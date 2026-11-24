# -*- coding: utf-8 -*-
"""
app.py
Ứng dụng Web Demo Độc lập: Hệ thống Dự đoán và Phân tích Khách hàng Rời bỏ Dịch vụ (Customer Churn Analytics).
Hỗ trợ đồng thời:
1. Giao diện Streamlit tương tác khi chạy cục bộ (`streamlit run app.py`).
2. Chuẩn WSGI (`app`, `application`) dùng Python thuần + `portable_engine` khi deploy trên Vercel Serverless.
"""

import json
import mimetypes
import os
import sys
from urllib.parse import parse_qs, unquote

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
if BASE_DIR not in sys.path:
    sys.path.insert(0, BASE_DIR)

from src.portable_engine import (
    FEATURE_VI_MAP,
    MODEL_NAMES,
    PRESET_PROFILES,
    batch_predict_csv,
    build_recommendations,
    compute_dynamic_roi_strategies,
    compute_threshold_metrics,
    find_optimal_threshold,
    generate_random_sample_csv,
    load_portable_model,
    predict_and_shap,
)

MODEL_PATH = os.path.join(BASE_DIR, "models", "best_churn_pipeline.joblib")
PORTABLE_MODEL_PATH = os.path.join(BASE_DIR, "models", "portable_lgbm_model.json")
RESULTS_PATH = os.path.join(BASE_DIR, "models", "experiment_results.json")
COMP_CSV_PATH = os.path.join(BASE_DIR, "models", "model_comparison.csv")
SAMPLE_CSV_PATH = os.path.join(BASE_DIR, "data", "sample_test_customers.csv")
RAW_CSV_PATH = os.path.join(BASE_DIR, "data", "WA_Fn-UseC_-Telco-Customer-Churn.csv")
FIG_DIR = os.path.join(BASE_DIR, "reports", "figures")

_WSGI_CACHE = {}


def _get_wsgi_data():
    """Tải và cache dữ liệu tĩnh + portable model 5 mô hình cho WSGI server."""
    if not _WSGI_CACHE:
        _WSGI_CACHE["model"] = load_portable_model(PORTABLE_MODEL_PATH)
        with open(RESULTS_PATH, "r", encoding="utf-8") as f:
            _WSGI_CACHE["results"] = json.load(f)
        import csv
        with open(COMP_CSV_PATH, "r", encoding="utf-8-sig") as f:
            _WSGI_CACHE["comp_rows"] = list(csv.DictReader(f))
        with open(SAMPLE_CSV_PATH, "r", encoding="utf-8-sig") as f:
            _WSGI_CACHE["sample_csv"] = f.read()
    return _WSGI_CACHE


def _compute_live_eval_payload(
    y_train: list,
    y_test: list,
    oof_probs_all: dict,
    test_probs_all: dict,
    model_name: str,
    threshold: float,
    c_fn: float,
    c_fp: float,
) -> dict:
    """Tính toàn bộ số liệu thực nghiệm động cho Tab 3 và 4 KPI tổng quan."""
    y_prob_test = test_probs_all[model_name]
    y_prob_oof = oof_probs_all[model_name]
    opt_thr, _ = find_optimal_threshold(y_train, y_prob_oof, c_fn=c_fn, c_fp=c_fp)
    _, cost_curve = find_optimal_threshold(y_test, y_prob_test, c_fn=c_fn, c_fp=c_fp)
    live_m = compute_threshold_metrics(y_test, y_prob_test, threshold, c_fn=c_fn, c_fp=c_fp)
    roi_rows = compute_dynamic_roi_strategies(y_test, y_prob_test, model_name, threshold, c_fn=c_fn, c_fp=c_fp)

    comp_live = []
    for m_name in MODEL_NAMES:
        m_thr_opt, _ = find_optimal_threshold(y_train, oof_probs_all[m_name], c_fn=c_fn, c_fp=c_fp)
        m_at_t = compute_threshold_metrics(y_test, test_probs_all[m_name], threshold, c_fn=c_fn, c_fp=c_fp)
        m_at_opt = compute_threshold_metrics(y_test, test_probs_all[m_name], m_thr_opt, c_fn=c_fn, c_fp=c_fp)
        comp_live.append({
            "Mô hình": m_name,
            "Accuracy (T)": f"{m_at_t['accuracy']*100:.2f}%",
            "Precision (T)": f"{m_at_t['precision']*100:.2f}%",
            "Recall (T)": f"{m_at_t['recall']*100:.2f}%",
            "F1 (T)": f"{m_at_t['f1']:.4f}",
            "TP / FN": f"{m_at_t['tp']} / {m_at_t['fn']}",
            "FP": m_at_t["fp"],
            "Chi phí tại T ($)": m_at_t["total_cost_usd"],
            "T* tối ưu": f"{m_thr_opt:.2f}",
            "Chi phí tại T* ($)": m_at_opt["total_cost_usd"],
        })

    return {
        "model_name": model_name,
        "optimal_threshold": opt_thr,
        "bayes_threshold": round(min(0.99, c_fp / c_fn) if c_fn > 0 else 0.5, 2),
        "live_metrics": live_m,
        "cost_curve": cost_curve,
        "roi_rows": roi_rows,
        "comp_live": comp_live,
    }


def _build_wsgi_html(data: dict) -> str:
    """Sinh trang HTML/CSS/JS độc lập cho Vercel WSGI với chế độ Sáng/Tối và 3 tính năng động."""
    results = data["results"]
    comp_rows = data["comp_rows"]
    model_meta = {
        "model_name": data["model"]["model_name"],
        "model_names": data["model"].get("model_names", MODEL_NAMES),
        "optimal_threshold": data["model"]["optimal_threshold"],
    }
    bootstrap_json = json.dumps(
        {
            "model": model_meta,
            "results": results,
            "comp_rows": comp_rows,
            "presets": PRESET_PROFILES,
        },
        ensure_ascii=False,
    )
    model_options_html = "".join(
        f'<option value="{m}">{m}</option>' for m in model_meta["model_names"]
    )
    return f"""<!DOCTYPE html>
<html lang="vi" data-theme="light">
<head>
  <meta charset="UTF-8" />
  <meta name="viewport" content="width=device-width, initial-scale=1.0" />
  <title>Customer Churn Analytics — Dự đoán Khách hàng Rời bỏ Dịch vụ</title>
  <style>
    :root {{
      --bg: #F8FAFC;
      --surface: #FFFFFF;
      --surface-alt: #F1F5F9;
      --border: #CBD5E1;
      --border-soft: #E2E8F0;
      --text: #0F172A;
      --muted: #475569;
      --accent: #1E40AF;
      --accent-text: #1E40AF;
      --accent-soft: #EFF6FF;
      --hero-bg: #0F2544;
      --hero-text: #FFFFFF;
      --hero-sub: #CBD5E1;
      --danger: #DC2626;
      --danger-bg: #FEF2F2;
      --danger-border: #FECACA;
      --danger-text: #7F1D1D;
      --warn: #D97706;
      --warn-bg: #FFFBEB;
      --warn-border: #FDE68A;
      --warn-text: #78350F;
      --ok: #059669;
      --ok-bg: #F0FDF4;
      --ok-border: #BBF7D0;
      --ok-text: #064E3B;
    }}
    [data-theme="dark"] {{
      --bg: #0B1120;
      --surface: #111827;
      --surface-alt: #1E293B;
      --border: #475569;
      --border-soft: #334155;
      --text: #F8FAFC;
      --muted: #CBD5E1;
      --accent: #2563EB;
      --accent-text: #60A5FA;
      --accent-soft: #1E293B;
      --hero-bg: #0F172A;
      --hero-text: #F8FAFC;
      --hero-sub: #CBD5E1;
      --danger: #EF4444;
      --danger-bg: #450A0A;
      --danger-border: #991B1B;
      --danger-text: #FECACA;
      --warn: #F59E0B;
      --warn-bg: #451A03;
      --warn-border: #92400E;
      --warn-text: #FDE68A;
      --ok: #10B981;
      --ok-bg: #022C22;
      --ok-border: #065F46;
      --ok-text: #A7F3D0;
    }}
    * {{ box-sizing: border-box; }}
    body {{
      margin: 0;
      background: var(--bg);
      color: var(--text);
      font-family: -apple-system, BlinkMacSystemFont, "Segoe UI", Roboto, Inter, sans-serif;
      font-size: 14px;
      line-height: 1.5;
    }}
    .container {{
      max-width: 1360px;
      margin: 0 auto;
      padding: 20px 24px 48px;
    }}
    .header {{
      background: var(--hero-bg);
      color: var(--hero-text);
      border-bottom: 3px solid #D97706;
      border-radius: 8px;
      padding: 20px 24px;
      display: flex;
      justify-content: space-between;
      align-items: center;
      gap: 16px;
      flex-wrap: wrap;
      margin-bottom: 18px;
    }}
    .header h1 {{ margin: 0; font-size: 21px; font-weight: 700; }}
    .header p {{ margin: 4px 0 0; color: var(--hero-sub); font-size: 13px; }}
    .theme-toggle {{
      background: rgba(255,255,255,0.14);
      color: #FFFFFF;
      border: 1px solid rgba(255,255,255,0.32);
      border-radius: 6px;
      padding: 8px 14px;
      font-size: 13px;
      font-weight: 600;
      cursor: pointer;
    }}
    .theme-toggle:hover {{ background: rgba(255,255,255,0.24); }}
    .kpi-grid {{
      display: grid;
      grid-template-columns: repeat(4, minmax(0, 1fr));
      gap: 14px;
      margin-bottom: 18px;
    }}
    .kpi-card {{
      background: var(--surface);
      border: 1px solid var(--border-soft);
      border-top: 3px solid var(--accent);
      border-radius: 8px;
      padding: 14px 16px;
      height: 120px;
      display: flex;
      flex-direction: column;
      justify-content: space-between;
    }}
    .kpi-card.amber {{ border-top-color: #D97706; }}
    .kpi-card.teal {{ border-top-color: #059669; }}
    .kpi-card.indigo {{ border-top-color: #6366F1; }}
    .kpi-label {{
      font-size: 11.5px;
      font-weight: 600;
      text-transform: uppercase;
      letter-spacing: 0.4px;
      color: var(--muted);
      white-space: nowrap;
      overflow: hidden;
      text-overflow: ellipsis;
    }}
    .kpi-val {{
      font-size: 19.5px;
      font-weight: 700;
      margin: 4px 0;
      white-space: nowrap;
      overflow: hidden;
      text-overflow: ellipsis;
    }}
    .kpi-sub {{ font-size: 12px; color: var(--muted); }}
    .control-bar {{
      background: var(--surface);
      border: 1px solid var(--border-soft);
      border-radius: 8px;
      padding: 12px 16px;
      display: flex;
      justify-content: space-between;
      align-items: center;
      flex-wrap: wrap;
      gap: 14px;
      margin-bottom: 16px;
    }}
    .ctrl-input {{
      padding: 6px 10px;
      border-radius: 6px;
      border: 1px solid var(--border);
      background: var(--surface-alt);
      color: var(--text);
      font-size: 13px;
    }}
    .tab-nav {{
      display: flex;
      gap: 8px;
      background: var(--surface);
      border: 1px solid var(--border-soft);
      border-radius: 10px;
      padding: 6px;
      margin-bottom: 16px;
      flex-wrap: wrap;
    }}
    .tab-btn {{
      border: 1px solid var(--border-soft);
      background: var(--surface-alt);
      color: var(--muted);
      padding: 9px 16px;
      border-radius: 7px;
      font-weight: 600;
      font-size: 13.5px;
      cursor: pointer;
      transition: all 0.18s ease;
    }}
    .tab-btn:hover {{
      background: var(--accent-soft);
      color: var(--accent-text);
      border-color: var(--accent-text);
    }}
    .tab-btn.active, .tab-btn.active:hover {{
      background: var(--accent);
      color: #FFFFFF;
      border-color: var(--accent);
    }}
    .tab-pane {{ display: none; }}
    .tab-pane.active {{ display: block; }}
    .btn-row {{ display: flex; gap: 10px; flex-wrap: wrap; margin-bottom: 16px; align-items: center; }}
    .btn {{
      background: var(--surface);
      color: var(--text);
      border: 1px solid var(--border);
      border-radius: 6px;
      padding: 8px 14px;
      font-size: 13px;
      font-weight: 600;
      cursor: pointer;
      height: 38px;
      transition: all 0.16s ease;
    }}
    .btn:hover {{ border-color: var(--accent-text); color: var(--accent-text); }}
    .btn-primary {{ background: var(--accent); color: #FFFFFF; border-color: var(--accent); }}
    .btn-primary:hover {{ color: #FFFFFF; opacity: 0.92; }}
    .btn-sm {{ padding: 5px 11px; font-size: 12px; height: 32px; }}
    .drop-zone {{
      border: 2px dashed var(--border);
      background: var(--surface);
      border-radius: 8px;
      padding: 24px 16px;
      text-align: center;
      cursor: pointer;
      transition: all 0.18s ease;
      margin-bottom: 12px;
    }}
    .drop-zone:hover, .drop-zone.dragover {{ border-color: var(--accent-text); background: var(--accent-soft); }}
    .sample-toolbar {{
      display: flex;
      align-items: center;
      gap: 10px;
      flex-wrap: wrap;
      background: var(--surface-alt);
      border: 1px solid var(--border-soft);
      border-radius: 6px;
      padding: 8px 12px;
      margin-bottom: 14px;
      font-size: 12.5px;
    }}
    .center-loading {{
      position: fixed; inset: 0; background: rgba(15, 23, 42, 0.42);
      backdrop-filter: blur(2px); display: none; align-items: center; justify-content: center; z-index: 99999;
    }}
    .center-loading.show {{ display: flex; }}
    .loading-card {{
      background: var(--surface); color: var(--text); border: 1px solid var(--border);
      border-radius: 10px; padding: 20px 28px; text-align: center; min-width: 280px;
    }}
    .spinner {{
      width: 28px; height: 28px; border: 3px solid var(--border); border-top-color: var(--accent-text);
      border-radius: 50%; margin: 0 auto 10px; animation: spin 0.75s linear infinite;
    }}
    @keyframes spin {{ to {{ transform: rotate(360deg); }} }}
    .form-grid {{ display: grid; grid-template-columns: repeat(3, minmax(0, 1fr)); gap: 14px; margin-bottom: 16px; }}
    .panel {{ background: var(--surface); border: 1px solid var(--border-soft); border-radius: 8px; padding: 16px; }}
    .panel h3 {{ margin: 0 0 12px; font-size: 14.5px; font-weight: 700; border-bottom: 1px solid var(--border-soft); padding-bottom: 8px; }}
    .field {{ margin-bottom: 10px; }}
    .field label {{ display: block; font-size: 12px; font-weight: 600; color: var(--muted); margin-bottom: 4px; }}
    .field select, .field input {{
      width: 100%; padding: 7px 10px; border-radius: 6px; border: 1px solid var(--border);
      background: var(--surface-alt); color: var(--text); font-size: 13px;
    }}
    .chips-grid {{ display: grid; grid-template-columns: repeat(6, minmax(0, 1fr)); gap: 10px; margin-bottom: 16px; }}
    .models-strip {{ display: grid; grid-template-columns: repeat(5, minmax(0, 1fr)); gap: 10px; margin-bottom: 16px; }}
    .chip {{
      background: var(--surface); border: 1px solid var(--border-soft); border-radius: 6px;
      padding: 10px 8px; text-align: center;
    }}
    .chip.selected {{ border: 2px solid var(--accent-text); background: var(--accent-soft); }}
    .chip-name {{ font-size: 11px; color: var(--muted); font-weight: 600; }}
    .chip-val {{ font-size: 14.5px; font-weight: 700; margin-top: 3px; }}
    .result-grid {{ display: grid; grid-template-columns: 1fr 1.35fr; gap: 16px; }}
    .live-chart-grid {{ display: grid; grid-template-columns: 1.35fr 1fr; gap: 16px; margin-bottom: 16px; }}
    .gauge-track {{
      position: relative; height: 14px; border-radius: 999px;
      background: linear-gradient(90deg, #10B981 0%, #F59E0B 45%, #EF4444 100%); margin: 12px 0 8px;
    }}
    .gauge-thr {{ position: absolute; top: -4px; bottom: -4px; width: 3px; background: var(--text); }}
    .gauge-mark {{
      position: absolute; top: -6px; width: 12px; height: 26px; transform: translateX(-50%);
      background: var(--surface); border: 3px solid var(--accent-text); border-radius: 4px;
    }}
    .status-box {{
      border-radius: 6px; padding: 12px 14px; border: 1px solid transparent; border-left-width: 4px; margin: 12px 0;
    }}
    .status-high {{ background: var(--danger-bg); border-color: var(--danger-border); border-left-color: var(--danger); color: var(--danger-text); }}
    .status-mid {{ background: var(--warn-bg); border-color: var(--warn-border); border-left-color: var(--warn); color: var(--warn-text); }}
    .status-low {{ background: var(--ok-bg); border-color: var(--ok-border); border-left-color: var(--ok); color: var(--ok-text); }}
    .rec-item {{
      background: var(--surface-alt); border-left: 3px solid var(--accent-text);
      border-radius: 6px; padding: 9px 12px; margin-bottom: 8px; font-size: 13px;
    }}
    .cm-grid {{ display: grid; grid-template-columns: 1fr 1fr; gap: 10px; margin-top: 8px; }}
    .cm-cell {{ border-radius: 8px; padding: 14px; text-align: center; border: 1px solid var(--border-soft); }}
    table {{ width: 100%; border-collapse: collapse; font-size: 12.5px; }}
    th, td {{ text-align: left; padding: 8px 10px; border-bottom: 1px solid var(--border-soft); }}
    th {{ background: var(--surface-alt); font-weight: 700; color: var(--muted); position: sticky; top: 0; }}
    .table-wrap {{ max-height: 420px; overflow: auto; border: 1px solid var(--border-soft); border-radius: 8px; background: var(--surface); }}
    .prog-bar {{ width: 90px; height: 8px; background: var(--surface-alt); border-radius: 999px; display: inline-block; overflow: hidden; vertical-align: middle; margin-right: 6px; }}
    .prog-fill {{ height: 100%; background: var(--accent-text); }}
    .fig-grid {{ display: grid; grid-template-columns: repeat(2, minmax(0, 1fr)); gap: 14px; margin-top: 14px; }}
    .fig-card {{ background: #FFFFFF; border: 1px solid var(--border-soft); border-radius: 8px; padding: 10px; color: #0F172A; }}
    .fig-card img {{ width: 100%; height: auto; display: block; border-radius: 4px; }}
    .fig-cap {{ font-size: 12px; color: #475569; margin-top: 6px; text-align: center; font-weight: 600; }}
    @media (max-width: 980px) {{
      .kpi-grid, .form-grid, .result-grid, .fig-grid, .live-chart-grid, .models-strip {{ grid-template-columns: 1fr; }}
      .chips-grid {{ grid-template-columns: repeat(2, minmax(0, 1fr)); }}
    }}
  </style>
</head>
<body>
  <div class="container">
    <header class="header">
      <div>
        <h1>Hệ thống Dự đoán và Phân tích Khách hàng Rời bỏ Dịch vụ (Customer Churn Analytics)</h1>
        <p>Mô hình phân loại nhạy chi phí (5 Thuật toán Học máy) và giải thích cục bộ TreeSHAP / Log-Odds</p>
      </div>
      <button class="theme-toggle" id="btn-theme" onclick="toggleTheme()">Giao diện: Tối</button>
    </header>

    <section class="kpi-grid" id="kpi-grid"></section>

    <section class="control-bar">
      <div style="display:flex; align-items:center; gap:10px; flex-wrap:wrap;">
        <strong>Mô hình:</strong>
        <select id="model-sel" class="ctrl-input" onchange="onControlChange(true)">{model_options_html}</select>
        <strong>Ngưỡng T:</strong>
        <input type="range" id="thr-slider" min="0.05" max="0.85" step="0.01" value="0.21" style="width:160px;" />
        <span id="thr-val" style="font-weight:700; min-width:38px;">0.21</span>
        <button class="btn btn-primary btn-sm" id="btn-apply-opt" onclick="applyOptimalThr()">Áp dụng T* = 0.21</button>
        <button class="btn btn-sm" onclick="setThr(0.50)">T = 0.50</button>
      </div>
      <div style="display:flex; align-items:center; gap:10px; flex-wrap:wrap; font-size:12.5px;">
        <label>C_FN (USD): <input type="number" id="c-fn" class="ctrl-input" value="500" min="50" max="5000" step="50" style="width:82px;" onchange="onControlChange(true)" /></label>
        <label>C_FP (USD): <input type="number" id="c-fp" class="ctrl-input" value="100" min="10" max="1000" step="10" style="width:78px;" onchange="onControlChange(true)" /></label>
        <span id="cost-ratio-lbl" style="color:var(--muted);">Tỷ lệ 5.0:1</span>
      </div>
    </section>

    <nav class="tab-nav">
      <button class="tab-btn active" onclick="switchTab('tab1', this)">1. Dự đoán khách hàng cá nhân</button>
      <button class="tab-btn" onclick="switchTab('tab2', this)">2. Chấm điểm danh sách (CSV)</button>
      <button class="tab-btn" onclick="switchTab('tab3', this)">3. Kết quả thực nghiệm mô hình</button>
    </nav>

    <!-- TAB 1 -->
    <div id="tab1" class="tab-pane active">
      <div class="btn-row">
        <button class="btn" onclick="applyPreset('high')">Hồ sơ mẫu 1: Nguy cơ cao (~87%)</button>
        <button class="btn" onclick="applyPreset('medium')">Hồ sơ mẫu 2: Nguy cơ trung bình (~30%)</button>
        <button class="btn" onclick="applyPreset('low')">Hồ sơ mẫu 3: Khách hàng ổn định (~2%)</button>
      </div>

      <div class="form-grid">
        <div class="panel">
          <h3>Nhân khẩu học &amp; Điện thoại</h3>
          <div class="field"><label>Giới tính (gender)</label><select id="f-gender"><option>Male</option><option>Female</option></select></div>
          <div class="field"><label>Khách hàng cao tuổi (SeniorCitizen)</label><select id="f-SeniorCitizen"><option value="0">Không</option><option value="1">Có</option></select></div>
          <div class="field"><label>Có bạn đời (Partner)</label><select id="f-Partner"><option>No</option><option>Yes</option></select></div>
          <div class="field"><label>Người phụ thuộc (Dependents)</label><select id="f-Dependents"><option>No</option><option>Yes</option></select></div>
          <div class="field"><label>Dịch vụ điện thoại (PhoneService)</label><select id="f-PhoneService"><option>Yes</option><option>No</option></select></div>
          <div class="field"><label>Nhiều đường dây (MultipleLines)</label><select id="f-MultipleLines"><option>No</option><option>Yes</option><option>No phone service</option></select></div>
        </div>
        <div class="panel">
          <h3>Dịch vụ Internet &amp; Gia tăng</h3>
          <div class="field"><label>Hạ tầng Internet (InternetService)</label><select id="f-InternetService"><option>Fiber optic</option><option>DSL</option><option>No</option></select></div>
          <div class="field"><label>Bảo mật trực tuyến (OnlineSecurity)</label><select id="f-OnlineSecurity"><option>No</option><option>Yes</option><option>No internet service</option></select></div>
          <div class="field"><label>Sao lưu trực tuyến (OnlineBackup)</label><select id="f-OnlineBackup"><option>No</option><option>Yes</option><option>No internet service</option></select></div>
          <div class="field"><label>Bảo vệ thiết bị (DeviceProtection)</label><select id="f-DeviceProtection"><option>No</option><option>Yes</option><option>No internet service</option></select></div>
          <div class="field"><label>Hỗ trợ kỹ thuật (TechSupport)</label><select id="f-TechSupport"><option>No</option><option>Yes</option><option>No internet service</option></select></div>
          <div class="field"><label>Truyền hình (StreamingTV)</label><select id="f-StreamingTV"><option>Yes</option><option>No</option><option>No internet service</option></select></div>
          <div class="field"><label>Phim trực tuyến (StreamingMovies)</label><select id="f-StreamingMovies"><option>Yes</option><option>No</option><option>No internet service</option></select></div>
        </div>
        <div class="panel">
          <h3>Hợp đồng &amp; Cước phí</h3>
          <div class="field"><label>Loại hợp đồng (Contract)</label><select id="f-Contract"><option>Month-to-month</option><option>One year</option><option>Two year</option></select></div>
          <div class="field"><label>Số tháng gắn bó (tenure): <span id="tenure-lbl">1</span></label><input type="range" id="f-tenure" min="0" max="72" value="1" /></div>
          <div class="field"><label>Cước hàng tháng (MonthlyCharges - USD)</label><input type="number" id="f-MonthlyCharges" min="18" max="130" step="0.5" value="95.1" /></div>
          <div class="field"><label>Tổng cước tích lũy (TotalCharges - USD)</label><input type="number" id="f-TotalCharges" min="0" max="10000" step="10" value="95.1" /></div>
          <div class="field"><label>Hóa đơn điện tử (PaperlessBilling)</label><select id="f-PaperlessBilling"><option>Yes</option><option>No</option></select></div>
          <div class="field"><label>Phương thức thanh toán (PaymentMethod)</label><select id="f-PaymentMethod"><option>Electronic check</option><option>Mailed check</option><option>Bank transfer (automatic)</option><option>Credit card (automatic)</option></select></div>
        </div>
      </div>

      <div class="chips-grid" id="domain-chips"></div>
      <h4 style="margin:0 0 8px;">Đối chiếu dự báo của 5 mô hình trên hồ sơ này</h4>
      <div class="models-strip" id="all-models-strip"></div>

      <div class="result-grid">
        <div class="panel">
          <div style="display:flex; justify-content:space-between; align-items:baseline;">
            <span id="pred-model-title" style="font-weight:600; color:var(--muted);">Xác suất rời bỏ (LightGBM)</span>
            <span id="prob-big" style="font-size:28px; font-weight:800;">--%</span>
          </div>
          <div class="gauge-track">
            <div class="gauge-thr" id="gauge-thr" style="left:21%;"></div>
            <div class="gauge-mark" id="gauge-mark" style="left:50%;"></div>
          </div>
          <div style="display:flex; justify-content:space-between; font-size:11.5px; color:var(--muted);">
            <span>0% (An toàn)</span>
            <span id="gauge-thr-txt">Ngưỡng T = 21%</span>
            <span>100% (Rời bỏ)</span>
          </div>
          <div id="status-box" class="status-box"></div>
          <h4 style="margin:12px 0 8px;">Khuyến nghị hành động CSKH</h4>
          <div id="rec-list"></div>
        </div>
        <div class="panel">
          <h3 id="shap-title">Giải thích cục bộ — Top 8 yếu tố ảnh hưởng</h3>
          <div id="shap-chart"></div>
        </div>
      </div>
    </div>

    <!-- TAB 2 -->
    <div id="tab2" class="tab-pane">
      <div id="csv-dropzone" class="drop-zone" onclick="document.getElementById('csv-upload').click()">
        <div style="font-weight:700; font-size:14.5px; margin-bottom:4px;">Thêm hoặc kéo thả tệp CSV danh sách khách hàng vào đây (Drag &amp; Drop)</div>
        <div style="font-size:12.5px; color:var(--muted);">Nhấp chuột để chọn tệp hoặc kéo thả file .csv (chỉ chạy chấm điểm khi bấm nút bên dưới)</div>
        <input type="file" id="csv-upload" accept=".csv" style="display:none;" onchange="handleCsvUpload(this)" />
      </div>

      <div class="sample-toolbar">
        <span style="font-weight:600; color:var(--muted);">Dữ liệu mẫu ngẫu nhiên:</span>
        <select id="sample-n" onchange="onSampleOptionChange()" class="ctrl-input">
          <option value="100" selected>100 khách hàng</option>
          <option value="200">200 khách hàng</option>
          <option value="500">500 khách hàng</option>
          <option value="1000">1000 khách hàng</option>
        </select>
        <button class="btn btn-sm" onclick="randomizeSample()">Đổi mẫu ngẫu nhiên (Random)</button>
        <a id="btn-dl-sample" class="btn btn-sm" href="/api/sample-csv?n=100&amp;seed=42" download="sample_customers_100.csv" style="text-decoration:none; display:inline-flex; align-items:center;">Tải tệp CSV mẫu (100 KH)</a>
      </div>

      <div id="batch-staged-status" class="status-box status-low" style="margin:0 0 12px 0;"></div>

      <div class="btn-row">
        <button class="btn btn-primary" onclick="runBatchScoring()">Chấm điểm</button>
        <button class="btn" onclick="resetBatchSelection()">Xóa kết quả / Chọn lại</button>
        <button id="btn-dl-result" class="btn" onclick="downloadBatchResult()" style="display:none;">Tải xuống kết quả dự đoán (CSV)</button>
      </div>

      <div id="batch-results-area" style="display:none;">
        <div class="kpi-grid" id="batch-kpis"></div>
        <div class="btn-row">
          <label>Lọc phân tầng:
            <select id="filter-tier" onchange="renderBatchTable()" class="ctrl-input">
              <option value="ALL">Tất cả phân tầng</option>
              <option value="high">Nguy cơ cao (&gt;= 50%)</option>
              <option value="mid">Cảnh báo (T – 50%)</option>
              <option value="low">Ổn định</option>
            </select>
          </label>
          <label>Lọc hợp đồng:
            <select id="filter-contract" onchange="renderBatchTable()" class="ctrl-input">
              <option value="ALL">Tất cả hợp đồng</option>
              <option value="Month-to-month">Month-to-month</option>
              <option value="One year">One year</option>
              <option value="Two year">Two year</option>
            </select>
          </label>
        </div>
        <div class="table-wrap">
          <table>
            <thead>
              <tr><th>Mã KH</th><th>Hợp đồng</th><th>Internet</th><th>Số tháng</th><th>Cước/tháng</th><th>Xác suất Churn</th><th>Phân tầng rủi ro</th></tr>
            </thead>
            <tbody id="batch-tbody"></tbody>
          </table>
        </div>
      </div>
    </div>

    <!-- TAB 3 -->
    <div id="tab3" class="tab-pane">
      <div class="tab-nav" style="margin-bottom:12px;">
        <button class="tab-btn active" onclick="switchSub('sub0', this)">Thực nghiệm động (Live)</button>
        <button class="tab-btn" onclick="switchSub('sub1', this)">Đối sánh cố định &amp; ROC/PR</button>
        <button class="tab-btn" onclick="switchSub('sub2', this)">Ablation Study</button>
        <button class="tab-btn" onclick="switchSub('sub3', this)">Khám phá dữ liệu (EDA)</button>
        <button class="tab-btn" onclick="switchSub('sub4', this)">Giải thích SHAP &amp; Luật CART</button>
      </div>
      <div id="sub0" class="sub-pane">
        <div class="kpi-grid" id="live-tab3-kpis"></div>
        <div class="live-chart-grid">
          <div class="panel"><h3>Đường cong Chi phí theo Ngưỡng T (Live Cost Curve)</h3><div id="live-cost-svg"></div></div>
          <div class="panel"><h3>Ma trận nhầm lẫn động trên Tập Test (N = 1.409)</h3><div id="live-cm-box"></div></div>
        </div>
        <h4 style="margin:0 0 8px;">Bảng đối sánh động 5 mô hình theo Ngưỡng T và Ma trận chi phí hiện tại</h4>
        <div id="live-comp-table"></div>
        <h4 style="margin:14px 0 8px;">Đối chứng chiến lược vận hành động</h4>
        <div id="live-roi-table"></div>
      </div>
      <div id="sub1" class="sub-pane" style="display:none;"></div>
      <div id="sub2" class="sub-pane" style="display:none;"></div>
      <div id="sub3" class="sub-pane" style="display:none;"></div>
      <div id="sub4" class="sub-pane" style="display:none;"></div>
    </div>
  </div>

  <div id="center-loading" class="center-loading">
    <div class="loading-card"><div class="spinner"></div><div id="loading-msg" style="font-weight:600; font-size:13.5px;">Đang xử lý...</div></div>
  </div>

  <script>
    const BOOT = {bootstrap_json};
    let currentBatch = [];
    let liveEvalCache = null;
    const FIELD_IDS = [
      "gender","SeniorCitizen","Partner","Dependents","PhoneService","MultipleLines",
      "InternetService","OnlineSecurity","OnlineBackup","DeviceProtection","TechSupport",
      "StreamingTV","StreamingMovies","Contract","tenure","MonthlyCharges","TotalCharges",
      "PaperlessBilling","PaymentMethod"
    ];

    function toggleTheme() {{
      const root = document.documentElement;
      const next = root.getAttribute("data-theme") === "dark" ? "light" : "dark";
      root.setAttribute("data-theme", next);
      localStorage.setItem("churn_theme", next);
      document.getElementById("btn-theme").textContent = next === "dark" ? "Giao diện: Sáng" : "Giao diện: Tối";
    }}
    (function initTheme() {{
      const saved = localStorage.getItem("churn_theme") || "light";
      document.documentElement.setAttribute("data-theme", saved);
      document.getElementById("btn-theme").textContent = saved === "dark" ? "Giao diện: Sáng" : "Giao diện: Tối";
    }})();

    function switchTab(id, btn) {{
      document.querySelectorAll(".tab-pane").forEach(el => el.classList.remove("active"));
      document.getElementById(id).classList.add("active");
      btn.parentElement.querySelectorAll(".tab-btn").forEach(b => b.classList.remove("active"));
      btn.classList.add("active");
    }}
    function switchSub(id, btn) {{
      document.querySelectorAll(".sub-pane").forEach(el => el.style.display = "none");
      document.getElementById(id).style.display = "block";
      btn.parentElement.querySelectorAll(".tab-btn").forEach(b => b.classList.remove("active"));
      btn.classList.add("active");
    }}
    function showLoading(msg) {{
      document.getElementById("loading-msg").textContent = msg || "Đang xử lý...";
      document.getElementById("center-loading").classList.add("show");
    }}
    function hideLoading() {{ document.getElementById("center-loading").classList.remove("show"); }}

    function getCtrl() {{
      return {{
        model_name: document.getElementById("model-sel").value,
        threshold: parseFloat(document.getElementById("thr-slider").value),
        c_fn: parseFloat(document.getElementById("c-fn").value) || 500,
        c_fp: parseFloat(document.getElementById("c-fp").value) || 100,
      }};
    }}

    async function refreshLiveEval() {{
      const c = getCtrl();
      document.getElementById("cost-ratio-lbl").textContent = `Tỷ lệ ${{(c.c_fn / c.c_fp).toFixed(1)}}:1`;
      const res = await fetch("/api/live-eval", {{
        method: "POST",
        headers: {{ "Content-Type": "application/json" }},
        body: JSON.stringify(c)
      }});
      liveEvalCache = await res.json();
      document.getElementById("btn-apply-opt").textContent = `Áp dụng T* = ${{liveEvalCache.optimal_threshold.toFixed(2)}}`;
      renderTopKpis();
      renderLiveTab3();
    }}

    function renderTopKpis() {{
      if (!liveEvalCache) return;
      const r = BOOT.results;
      const c = getCtrl();
      const cvAuc = r.cv_roc_auc[c.model_name] || 0;
      const testAuc = (r.test_metrics_at_050[c.model_name] || {{}}).roc_auc || 0;
      const lm = liveEvalCache.live_metrics;
      document.getElementById("kpi-grid").innerHTML = `
        <div class="kpi-card"><div class="kpi-label">Quy mô dữ liệu Telco</div><div class="kpi-val">${{r.dataset_audit.n_rows.toLocaleString()}} thuê bao</div><div class="kpi-sub">${{r.dataset_audit.n_raw_cols}} cột gốc -> ${{r.dataset_audit.n_features_engineered}} đặc trưng • Churn ${{(r.dataset_audit.churn_rate*100).toFixed(2)}}%</div></div>
        <div class="kpi-card indigo"><div class="kpi-label">Mô hình đang chọn</div><div class="kpi-val">${{c.model_name}}</div><div class="kpi-sub">5-Fold CV AUC: ${{cvAuc.toFixed(4)}} • Test AUC: ${{testAuc.toFixed(4)}}</div></div>
        <div class="kpi-card amber"><div class="kpi-label">Ngưỡng T / T* tối ưu</div><div class="kpi-val">T = ${{c.threshold.toFixed(2)}} (T*=${{liveEvalCache.optimal_threshold.toFixed(2)}})</div><div class="kpi-sub">Recall Test: ${{(lm.recall*100).toFixed(2)}}% (${{lm.tp}}/374 Churn)</div></div>
        <div class="kpi-card teal"><div class="kpi-label">Hiệu quả kinh tế (ROI)</div><div class="kpi-val">Tiết kiệm $${{lm.savings_usd.toLocaleString()}}</div><div class="kpi-sub">Giảm ${{lm.savings_pct.toFixed(1)}}% so với không can thiệp</div></div>
      `;
    }}

    function renderLiveTab3() {{
      if (!liveEvalCache) return;
      const c = getCtrl();
      const lm = liveEvalCache.live_metrics;
      document.getElementById("live-tab3-kpis").innerHTML = `
        <div class="kpi-card"><div class="kpi-label">Accuracy / F1-Score</div><div class="kpi-val">${{(lm.accuracy*100).toFixed(2)}}% / ${{lm.f1.toFixed(4)}}</div><div class="kpi-sub">Tập Test N = 1.409</div></div>
        <div class="kpi-card indigo"><div class="kpi-label">Recall / Precision</div><div class="kpi-val">${{(lm.recall*100).toFixed(2)}}% / ${{(lm.precision*100).toFixed(2)}}%</div><div class="kpi-sub">Giữ được ${{lm.tp}}/374 Churn • Bỏ sót ${{lm.fn}}</div></div>
        <div class="kpi-card amber"><div class="kpi-label">Tổng chi phí tại T = ${{c.threshold.toFixed(2)}}</div><div class="kpi-val">$${{lm.total_cost_usd.toLocaleString()}}</div><div class="kpi-sub">TB $${{lm.avg_cost_usd.toFixed(2)}} / khách hàng</div></div>
        <div class="kpi-card teal"><div class="kpi-label">Tiết kiệm vs Không can thiệp</div><div class="kpi-val">$${{lm.savings_usd.toLocaleString()}}</div><div class="kpi-sub">Giảm ${{lm.savings_pct.toFixed(1)}}% (Không can thiệp: $${{lm.do_nothing_cost_usd.toLocaleString()}})</div></div>
      `;

      // SVG Live Cost Curve
      const pts = liveEvalCache.cost_curve;
      const W = 560, H = 260, P = 48;
      const maxC = Math.max(...pts.map(p => p.total_cost_usd)) * 1.05;
      const minC = Math.min(...pts.map(p => p.total_cost_usd)) * 0.92;
      const toX = t => P + ((t - 0.05) / 0.80) * (W - P - 20);
      const toY = v => H - 34 - ((v - minC) / Math.max(1, maxC - minC)) * (H - 54);
      const poly = pts.map(p => `${{toX(p.threshold).toFixed(1)}},${{toY(p.total_cost_usd).toFixed(1)}}`).join(" ");
      const xOpt = toX(liveEvalCache.optimal_threshold);
      const xCur = toX(c.threshold);
      document.getElementById("live-cost-svg").innerHTML = `
        <svg viewBox="0 0 ${{W}} ${{H}}" style="width:100%; height:auto;">
          <polyline fill="none" stroke="var(--accent-text)" stroke-width="2.5" points="${{poly}}" />
          <line x1="${{xOpt}}" y1="16" x2="${{xOpt}}" y2="${{H-34}}" stroke="#10B981" stroke-width="2" stroke-dasharray="5,4" />
          <line x1="${{xCur}}" y1="16" x2="${{xCur}}" y2="${{H-34}}" stroke="#EF4444" stroke-width="2" />
          <text x="${{xOpt}}" y="14" text-anchor="middle" font-size="11" font-weight="700" fill="#10B981">T*=${{liveEvalCache.optimal_threshold.toFixed(2)}}</text>
          <text x="${{xCur}}" y="28" text-anchor="middle" font-size="11" font-weight="700" fill="#EF4444">T=${{c.threshold.toFixed(2)}}</text>
          <text x="${{P}}" y="${{H-10}}" font-size="11" fill="currentColor">0.05</text>
          <text x="${{W/2}}" y="${{H-10}}" text-anchor="middle" font-size="11" fill="currentColor">Ngưỡng quyết định T</text>
          <text x="${{W-20}}" y="${{H-10}}" text-anchor="end" font-size="11" fill="currentColor">0.85</text>
        </svg>
      `;

      document.getElementById("live-cm-box").innerHTML = `
        <div class="cm-grid">
          <div class="cm-cell" style="background:var(--ok-bg); color:var(--ok-text);"><div style="font-size:11px; font-weight:700;">ĐÚNG Ở LẠI (TN)</div><div style="font-size:24px; font-weight:800;">${{lm.tn}}</div><div style="font-size:11px;">Chi phí: $0</div></div>
          <div class="cm-cell" style="background:var(--warn-bg); color:var(--warn-text);"><div style="font-size:11px; font-weight:700;">CẢNH BÁO NHẦM (FP)</div><div style="font-size:24px; font-weight:800;">${{lm.fp}}</div><div style="font-size:11px;">$${{(lm.fp*c.c_fp).toLocaleString()}}</div></div>
          <div class="cm-cell" style="background:var(--danger-bg); color:var(--danger-text);"><div style="font-size:11px; font-weight:700;">BỎ SÓT CHURN (FN)</div><div style="font-size:24px; font-weight:800;">${{lm.fn}}</div><div style="font-size:11px;">$${{(lm.fn*c.c_fn).toLocaleString()}}</div></div>
          <div class="cm-cell" style="background:var(--accent-soft); color:var(--accent-text);"><div style="font-size:11px; font-weight:700;">BẮT TRÚNG CHURN (TP)</div><div style="font-size:24px; font-weight:800;">${{lm.tp}}</div><div style="font-size:11px;">$${{(lm.tp*c.c_fp).toLocaleString()}}</div></div>
        </div>
      `;

      const compCols = Object.keys(liveEvalCache.comp_live[0] || {{}});
      document.getElementById("live-comp-table").innerHTML = makeTable(compCols, liveEvalCache.comp_live.map(r => compCols.map(k => r[k])));
      const roiCols = Object.keys(liveEvalCache.roi_rows[0] || {{}});
      document.getElementById("live-roi-table").innerHTML = makeTable(roiCols, liveEvalCache.roi_rows.map(r => roiCols.map(k => r[k])));
    }}

    async function onControlChange(withLoader = false) {{
      if (withLoader) showLoading("Đang cập nhật mô hình & ma trận chi phí...");
      await refreshLiveEval();
      await runPredict();
      if (hasScoredBatch) await runBatchScoring(false);
      if (withLoader) setTimeout(hideLoading, 180);
    }}

    function applyOptimalThr() {{
      if (!liveEvalCache) return;
      setThr(liveEvalCache.optimal_threshold);
    }}

    async function setThr(val) {{
      document.getElementById("thr-slider").value = val.toFixed(2);
      document.getElementById("thr-val").textContent = val.toFixed(2);
      await onControlChange(true);
    }}
    document.getElementById("thr-slider").addEventListener("input", (e) => {{
      document.getElementById("thr-val").textContent = parseFloat(e.target.value).toFixed(2);
      runPredict();
    }});
    document.getElementById("thr-slider").addEventListener("change", () => onControlChange(true));

    function applyPreset(key) {{
      const p = BOOT.presets[key];
      FIELD_IDS.forEach(k => {{
        const el = document.getElementById("f-" + k);
        if (el) el.value = p[k];
      }});
      document.getElementById("tenure-lbl").textContent = p.tenure;
      runPredict();
    }}

    FIELD_IDS.forEach(k => {{
      const el = document.getElementById("f-" + k);
      if (el) {{
        el.addEventListener("change", runPredict);
        if (k === "tenure") {{
          el.addEventListener("input", () => {{
            document.getElementById("tenure-lbl").textContent = el.value;
            runPredict();
          }});
        }}
      }}
    }});

    async function runPredict() {{
      const c = getCtrl();
      const payload = {{ threshold: c.threshold, model_name: c.model_name }};
      FIELD_IDS.forEach(k => {{ payload[k] = document.getElementById("f-" + k).value; }});
      const res = await fetch("/api/predict", {{
        method: "POST",
        headers: {{ "Content-Type": "application/json" }},
        body: JSON.stringify(payload)
      }});
      const out = await res.json();
      const pct = out.probability * 100;
      const thr = c.threshold;
      const thrPct = thr * 100;

      document.getElementById("pred-model-title").textContent = `Xác suất rời bỏ (${{c.model_name}})`;
      document.getElementById("prob-big").textContent = pct.toFixed(2) + "%";
      document.getElementById("gauge-thr").style.left = thrPct.toFixed(1) + "%";
      document.getElementById("gauge-mark").style.left = Math.min(Math.max(pct, 1), 99).toFixed(1) + "%";
      document.getElementById("gauge-thr-txt").textContent = "Ngưỡng T = " + thrPct.toFixed(0) + "%";

      const sb = document.getElementById("status-box");
      if (out.probability >= 0.50) {{
        sb.className = "status-box status-high";
        sb.innerHTML = `<strong>CẢNH BÁO NGUY CƠ RỜI BỎ RẤT CAO</strong><br>Xác suất <b>${{pct.toFixed(2)}}%</b> vượt ngưỡng T = ${{thr.toFixed(2)}} và vượt mốc 50%.`;
      }} else if (out.probability >= thr) {{
        sb.className = "status-box status-mid";
        sb.innerHTML = `<strong>CẢNH BÁO RỦI RO TRUNG GIAN (BẮT NHỜ NGƯỠNG T)</strong><br>Xác suất <b>${{pct.toFixed(2)}}%</b> vượt ngưỡng T = ${{thr.toFixed(2)}}.`;
      }} else {{
        sb.className = "status-box status-low";
        sb.innerHTML = `<strong>NHÓM KHÁCH HÀNG ỔN ĐỊNH</strong><br>Xác suất <b>${{pct.toFixed(2)}}%</b> thấp hơn ngưỡng T = ${{thr.toFixed(2)}}.`;
      }}

      const df = out.domain_features;
      const chips = [
        ["Chi tiêu TB/tháng", "$" + df.AvgMonthlySpend.toFixed(2)],
        ["Chênh lệch cước", (df.ChargeDiff >= 0 ? "+" : "") + df.ChargeDiff.toFixed(2) + " USD"],
        ["Dịch vụ cộng thêm", df.TotalAddonServices + " / 6"],
        ["Gói Bảo mật/HTKT", df.HasProtectionBundle === 1 ? "Có" : "Không"],
        ["Thanh toán tự động", df.IsAutoPayment === 1 ? "Tự động" : "Thủ công"],
        ["Nhóm vòng đời", df.TenureGroup],
      ];
      document.getElementById("domain-chips").innerHTML = chips.map(ch =>
        `<div class="chip"><div class="chip-name">${{ch[0]}}</div><div class="chip-val">${{ch[1]}}</div></div>`
      ).join("");

      const allProbs = out.all_model_probs || {{}};
      document.getElementById("all-models-strip").innerHTML = Object.entries(allProbs).map(([m, p]) => {{
        const selCls = m === c.model_name ? "chip selected" : "chip";
        const tag = p >= thr ? "Cảnh báo" : "Ổn định";
        return `<div class="${{selCls}}"><div class="chip-name">${{m}}</div><div class="chip-val">${{(p*100).toFixed(2)}}%</div><div class="chip-name">${{tag}}</div></div>`;
      }}).join("");

      document.getElementById("rec-list").innerHTML = out.recommendations.map(r => `<div class="rec-item">${{r}}</div>`).join("");

      document.getElementById("shap-title").textContent = c.model_name === "Logistic Regression"
        ? "Đóng góp Log-Odds (Logistic Regression) — Top 8 yếu tố"
        : `Giải thích TreeSHAP (${{c.model_name}}) — Top 8 yếu tố`;

      const top = out.top_shap.slice().reverse();
      const maxAbs = Math.max(0.05, ...top.map(d => Math.abs(d.value))) * 1.35;
      const W = 580, H = 310, L = 230, R = 50, rowH = 32;
      const plotW = W - L - R;
      const zeroX = L + plotW / 2;
      const scale = (plotW / 2) / maxAbs;
      let barsSvg = `<line x1="${{zeroX}}" y1="12" x2="${{zeroX}}" y2="${{12 + top.length * rowH}}" stroke="currentColor" stroke-opacity="0.45" stroke-width="1.2"/>`;
      top.forEach((d, idx) => {{
        const y = 16 + idx * rowH;
        const w = Math.abs(d.value) * scale;
        const x = d.value >= 0 ? zeroX : zeroX - w;
        const fill = d.value >= 0 ? "#DC2626" : "#059669";
        const tx = d.value >= 0 ? zeroX + w + 6 : zeroX - w - 6;
        const anchor = d.value >= 0 ? "start" : "end";
        const sign = d.value >= 0 ? "+" : "";
        barsSvg += `
          <text x="${{L - 10}}" y="${{y + 15}}" text-anchor="end" font-size="11.5" fill="currentColor">${{d.label}}</text>
          <rect x="${{x}}" y="${{y + 3}}" width="${{Math.max(w, 2)}}" height="18" rx="3" fill="${{fill}}" />
          <text x="${{tx}}" y="${{y + 15}}" text-anchor="${{anchor}}" font-size="11" font-weight="700" fill="currentColor">${{sign}}${{d.value.toFixed(3)}}</text>
        `;
      }});
      document.getElementById("shap-chart").innerHTML = `<svg viewBox="0 0 ${{W}} ${{H}}" style="width:100%; height:auto;">${{barsSvg}}</svg>
        <div style="font-size:11.5px; color:var(--muted); text-align:center;">Đỏ: tăng nguy cơ rời bỏ | Xanh: giảm nguy cơ rời bỏ</div>`;
    }}

    let lastCsvText = null, uploadedFileName = "", sampleSeed = 42, hasScoredBatch = false;
    function updateStagedBanner() {{
      const n = document.getElementById("sample-n").value;
      const c = getCtrl();
      const dl = document.getElementById("btn-dl-sample");
      dl.href = `/api/sample-csv?n=${{n}}&seed=${{sampleSeed}}`;
      dl.download = `sample_customers_${{n}}.csv`;
      dl.textContent = `Tải tệp CSV mẫu (${{n}} KH)`;
      const box = document.getElementById("batch-staged-status");
      if (lastCsvText !== null) {{
        const lineCount = Math.max(0, lastCsvText.trim().split(/\\r?\\n/).length - 1);
        box.innerHTML = `<strong>Đã nạp tệp:</strong> ${{uploadedFileName}} (${{lineCount}} dòng) • Mô hình: <b>${{c.model_name}}</b>. Bấm nút <b>"Chấm điểm"</b> để chạy.`;
      }} else {{
        box.innerHTML = `<strong>Dữ liệu đang chọn:</strong> Mẫu ngẫu nhiên <b>${{n}} khách hàng</b> (Seed #${{sampleSeed}}) • Mô hình: <b>${{c.model_name}}</b>. Bấm nút <b>"Chấm điểm"</b> để chạy.`;
      }}
    }}
    function clearBatchResults() {{
      hasScoredBatch = false; currentBatch = [];
      document.getElementById("batch-results-area").style.display = "none";
      document.getElementById("btn-dl-result").style.display = "none";
      updateStagedBanner();
    }}
    function onSampleOptionChange() {{ lastCsvText = null; uploadedFileName = ""; document.getElementById("csv-upload").value = ""; clearBatchResults(); }}
    function randomizeSample() {{ sampleSeed = Math.floor(Math.random() * 90000) + 1000; lastCsvText = null; uploadedFileName = ""; document.getElementById("csv-upload").value = ""; clearBatchResults(); }}
    function resetBatchSelection() {{ lastCsvText = null; uploadedFileName = ""; document.getElementById("csv-upload").value = ""; clearBatchResults(); }}

    async function runBatchScoring(withSpinner = true) {{
      const c = getCtrl();
      const sampleN = parseInt(document.getElementById("sample-n").value, 10);
      if (withSpinner) showLoading(`Đang chạy mô hình ${{c.model_name}} chấm điểm...`);
      try {{
        const res = await fetch("/api/batch", {{
          method: "POST",
          headers: {{ "Content-Type": "application/json" }},
          body: JSON.stringify({{ threshold: c.threshold, model_name: c.model_name, c_fp: c.c_fp, csv_text: lastCsvText, sample_n: sampleN, sample_seed: sampleSeed }})
        }});
        const data = await res.json();
        currentBatch = data.rows || [];
        hasScoredBatch = true;
        document.getElementById("batch-results-area").style.display = "block";
        document.getElementById("btn-dl-result").style.display = "inline-block";
        const totalCount = data.total || 1;
        const highPct = ((data.n_high_50 / totalCount) * 100).toFixed(1);
        document.getElementById("batch-kpis").innerHTML = `
          <div class="kpi-card"><div class="kpi-label">Tổng số khách hàng</div><div class="kpi-val">${{data.total}}</div><div class="kpi-sub">Mô hình: ${{c.model_name}}</div></div>
          <div class="kpi-card amber"><div class="kpi-label">Cần can thiệp (&gt;= ${{c.threshold.toFixed(2)}})</div><div class="kpi-val">${{data.n_intervene}}</div><div class="kpi-sub">Chiếm ${{data.intervene_pct}}% danh sách</div></div>
          <div class="kpi-card indigo"><div class="kpi-label">Nguy cơ cao (&gt;= 0.50)</div><div class="kpi-val">${{data.n_high_50}}</div><div class="kpi-sub">Chiếm ${{highPct}}% danh sách</div></div>
          <div class="kpi-card teal"><div class="kpi-label">Ngân sách dự kiến</div><div class="kpi-val">$${{data.budget_usd.toLocaleString()}}</div><div class="kpi-sub">Định mức ${{c.c_fp}} USD / KH can thiệp</div></div>
        `;
        renderBatchTable();
      }} finally {{
        if (withSpinner) setTimeout(hideLoading, 180);
      }}
    }}

    function renderBatchTable() {{
      const ft = document.getElementById("filter-tier").value;
      const fc = document.getElementById("filter-contract").value;
      const filtered = currentBatch.filter(r => (ft === "ALL" || r.TierCode === ft) && (fc === "ALL" || r.Contract === fc));
      document.getElementById("batch-tbody").innerHTML = filtered.map(r => `
        <tr><td>${{r.customerID}}</td><td>${{r.Contract}}</td><td>${{r.InternetService}}</td><td>${{r.tenure}}</td><td>$${{r.MonthlyCharges.toFixed(2)}}</td>
        <td><span class="prog-bar"><span class="prog-fill" style="width:${{r.Churn_Probability_Pct}}%;"></span></span>${{r.Churn_Probability_Pct.toFixed(2)}}%</td><td>${{r.Tier}}</td></tr>
      `).join("");
    }}

    function readCsvFile(file) {{
      if (!file) return;
      const reader = new FileReader();
      reader.onload = () => {{ lastCsvText = reader.result; uploadedFileName = file.name; clearBatchResults(); }};
      reader.readAsText(file, "utf-8");
    }}
    function handleCsvUpload(input) {{ readCsvFile(input.files[0]); }}
    (function initDropZone() {{
      const dz = document.getElementById("csv-dropzone");
      if (!dz) return;
      dz.addEventListener("dragover", (e) => {{ e.preventDefault(); dz.classList.add("dragover"); }});
      dz.addEventListener("dragleave", () => dz.classList.remove("dragover"));
      dz.addEventListener("drop", (e) => {{
        e.preventDefault(); dz.classList.remove("dragover");
        if (e.dataTransfer && e.dataTransfer.files && e.dataTransfer.files[0]) readCsvFile(e.dataTransfer.files[0]);
      }});
    }})();

    function downloadBatchResult() {{
      if (!currentBatch.length) return;
      const hdr = ["customerID","Contract","InternetService","tenure","MonthlyCharges","Churn_Probability_Pct","Tier"];
      const lines = [hdr.join(",")].concat(currentBatch.map(r => hdr.map(k => `"${{r[k]}}"`).join(",")));
      const blob = new Blob(["\\uFEFF" + lines.join("\\n")], {{ type: "text/csv;charset=utf-8;" }});
      const a = document.createElement("a");
      a.href = URL.createObjectURL(blob); a.download = "churn_batch_predictions.csv"; a.click();
    }}

    function makeTable(headers, rows) {{
      return `<div class="table-wrap" style="max-height:none; margin-bottom:14px;"><table><thead><tr>${{headers.map(h => `<th>${{h}}</th>`).join("")}}</tr></thead><tbody>${{
        rows.map(r => `<tr>${{r.map(c => `<td>${{c}}</td>`).join("")}}</tr>`).join("")
      }}</tbody></table></div>`;
    }}

    function renderStaticExperiments() {{
      const r = BOOT.results;
      const compCols = Object.keys(BOOT.comp_rows[0] || {{}});
      document.getElementById("sub1").innerHTML =
        `<h4>Bảng đối sánh cố định 5 mô hình trên tập Test độc lập (N = 1.409)</h4>` +
        makeTable(compCols, BOOT.comp_rows.map(row => compCols.map(c => row[c]))) +
        `<div class="fig-grid">
          <div class="fig-card"><img src="/figures/fig04_model_comparison.png" /><div class="fig-cap">Hình 4: Đối sánh ROC-AUC, PR-AUC và F1-Score</div></div>
          <div class="fig-card"><img src="/figures/fig05_roc_pr_curves.png" /><div class="fig-cap">Hình 5: Đường cong ROC-AUC và PR-AUC</div></div>
        </div>
        <div class="fig-card" style="margin-top:14px;"><img src="/figures/fig06_confusion_matrices.png" /><div class="fig-cap">Hình 6: Ma trận nhầm lẫn LightGBM tại T = 0.50 và T* = 0.21</div></div>`;

      const abRows = Object.entries(r.ablation_study).map(([k, m]) => [k, m.threshold, m.accuracy, m.precision, m.recall, m.f1, m.roc_auc, m.brier, "$" + m.total_cost_usd.toLocaleString()]);
      document.getElementById("sub2").innerHTML =
        `<h4>Nghiên cứu cắt bỏ (Ablation Study) xử lý mất cân bằng trên LightGBM</h4>` +
        makeTable(["Phương án","Ngưỡng T","Accuracy","Precision","Recall","F1","ROC-AUC","Brier","Tổng chi phí"], abRows) +
        `<div class="fig-card"><img src="/figures/fig07_cost_threshold_roi.png" /><div class="fig-cap">Hình 7: Tối ưu ngưỡng chi phí T* = 0.21 và đối chứng ROI</div></div>`;

      document.getElementById("sub3").innerHTML = `
        <div class="fig-grid">
          <div class="fig-card"><img src="/figures/fig01_churn_distribution.png" /><div class="fig-cap">Hình 1: Phân phối biến mục tiêu Churn</div></div>
          <div class="fig-card"><img src="/figures/fig02_eda_drivers.png" /><div class="fig-cap">Hình 2: Các yếu tố cốt lõi dẫn tới Churn</div></div>
        </div>
        <div class="fig-card" style="margin-top:14px;"><img src="/figures/fig03_correlation_cramersv.png" /><div class="fig-cap">Hình 3: Ma trận tương quan và Cramér's V</div></div>`;

      const rulesRows = r.extracted_cart_rules.map(x => [x.conditions, x.n_samples, x.support_pct + "%", x.churn_prob_pct + "%", x.predicted_class]);
      document.getElementById("sub4").innerHTML = `
        <div class="fig-card"><img src="/figures/fig08_decision_tree_rules.png" /><div class="fig-cap">Hình 8: Cây quyết định CART rút trích luật IF-THEN</div></div>
        <div class="fig-grid">
          <div class="fig-card"><img src="/figures/fig09_shap_summary.png" /><div class="fig-cap">Hình 9: SHAP Summary Plot (Top 12 đặc trưng)</div></div>
          <div class="fig-card"><img src="/figures/fig10_shap_waterfall_cases.png" /><div class="fig-cap">Hình 10: SHAP Waterfall cho 2 hồ sơ điển hình</div></div>
        </div>
        <h4 style="margin-top:16px;">8 Luật phân loại IF-THEN rút trích từ cây quyết định CART</h4>` +
        makeTable(["Điều kiện luật (IF)","Số mẫu","Độ hỗ trợ","Xác suất Churn","Kết luận (THEN)"], rulesRows);
    }}

    refreshLiveEval();
    applyPreset("high");
    updateStagedBanner();
    renderStaticExperiments();
  </script>
</body>
</html>
"""


def _wsgi_response(start_response, status: str, content_type: str, body: bytes):
    headers = [
        ("Content-Type", content_type),
        ("Content-Length", str(len(body))),
        ("Cache-Control", "no-store"),
    ]
    start_response(status, headers)
    return [body]


def app(environ, start_response):
    """WSGI Entrypoint chuẩn cho Vercel (`app`, `application`)."""
    method = environ.get("REQUEST_METHOD", "GET").upper()
    path = unquote(environ.get("PATH_INFO", "/"))
    data = _get_wsgi_data()

    if method == "GET" and path in ("/", "/index.html"):
        html = _build_wsgi_html(data)
        return _wsgi_response(start_response, "200 OK", "text/html; charset=utf-8", html.encode("utf-8"))

    if method == "GET" and path == "/api/health":
        body = json.dumps({"status": "ok", "models": data["model"].get("model_names", MODEL_NAMES)}).encode("utf-8")
        return _wsgi_response(start_response, "200 OK", "application/json; charset=utf-8", body)

    if method == "GET" and path == "/api/sample-csv":
        qs = parse_qs(environ.get("QUERY_STRING", ""))
        n_rows = int(qs.get("n", ["100"])[0])
        seed_raw = qs.get("seed", [None])[0]
        seed_val = int(seed_raw) if seed_raw is not None else None
        if os.path.isfile(RAW_CSV_PATH):
            csv_body = generate_random_sample_csv(RAW_CSV_PATH, n_rows=n_rows, seed=seed_val)
        else:
            csv_body = data["sample_csv"]
        return _wsgi_response(start_response, "200 OK", "text/csv; charset=utf-8", csv_body.encode("utf-8-sig"))

    if method == "GET" and path.startswith("/figures/"):
        fname = os.path.basename(path)
        fpath = os.path.join(FIG_DIR, fname)
        if os.path.isfile(fpath):
            ctype, _ = mimetypes.guess_type(fpath)
            with open(fpath, "rb") as f:
                return _wsgi_response(start_response, "200 OK", ctype or "image/png", f.read())
        return _wsgi_response(start_response, "404 Not Found", "text/plain; charset=utf-8", b"Not Found")

    if method == "POST" and path == "/api/predict":
        try:
            length = int(environ.get("CONTENT_LENGTH") or 0)
            payload = json.loads(environ["wsgi.input"].read(length).decode("utf-8"))
            threshold = float(payload.pop("threshold", data["model"]["optimal_threshold"]))
            model_name = payload.pop("model_name", "LightGBM")
            res = predict_and_shap(payload, data["model"], top_k=8, model_name=model_name)
            res["recommendations"] = build_recommendations(payload, res["probability"], threshold)
            body = json.dumps(res, ensure_ascii=False).encode("utf-8")
            return _wsgi_response(start_response, "200 OK", "application/json; charset=utf-8", body)
        except Exception as exc:
            body = json.dumps({"error": str(exc)}, ensure_ascii=False).encode("utf-8")
            return _wsgi_response(start_response, "400 Bad Request", "application/json; charset=utf-8", body)

    if method == "POST" and path == "/api/batch":
        try:
            length = int(environ.get("CONTENT_LENGTH") or 0)
            payload = json.loads(environ["wsgi.input"].read(length).decode("utf-8") or "{}")
            threshold = float(payload.get("threshold", data["model"]["optimal_threshold"]))
            model_name = payload.get("model_name", "LightGBM")
            c_fp = float(payload.get("c_fp", 100.0))
            csv_text = payload.get("csv_text")
            if not csv_text:
                sample_n = int(payload.get("sample_n", 100))
                sample_seed = payload.get("sample_seed", 42)
                if os.path.isfile(RAW_CSV_PATH):
                    csv_text = generate_random_sample_csv(RAW_CSV_PATH, n_rows=sample_n, seed=sample_seed)
                else:
                    csv_text = data["sample_csv"]
            res = batch_predict_csv(csv_text, threshold, data["model"], model_name=model_name, c_fp=c_fp)
            body = json.dumps(res, ensure_ascii=False).encode("utf-8")
            return _wsgi_response(start_response, "200 OK", "application/json; charset=utf-8", body)
        except Exception as exc:
            body = json.dumps({"error": str(exc)}, ensure_ascii=False).encode("utf-8")
            return _wsgi_response(start_response, "400 Bad Request", "application/json; charset=utf-8", body)

    if method == "POST" and path == "/api/live-eval":
        try:
            length = int(environ.get("CONTENT_LENGTH") or 0)
            payload = json.loads(environ["wsgi.input"].read(length).decode("utf-8") or "{}")
            model_name = payload.get("model_name", "LightGBM")
            threshold = float(payload.get("threshold", data["model"]["optimal_threshold"]))
            c_fn = float(payload.get("c_fn", 500.0))
            c_fp = float(payload.get("c_fp", 100.0))
            m = data["model"]
            res = _compute_live_eval_payload(
                m["y_train"], m["y_test"], m["oof_probs_all"], m["test_probs_all"],
                model_name, threshold, c_fn, c_fp,
            )
            body = json.dumps(res, ensure_ascii=False).encode("utf-8")
            return _wsgi_response(start_response, "200 OK", "application/json; charset=utf-8", body)
        except Exception as exc:
            body = json.dumps({"error": str(exc)}, ensure_ascii=False).encode("utf-8")
            return _wsgi_response(start_response, "400 Bad Request", "application/json; charset=utf-8", body)

    return _wsgi_response(start_response, "404 Not Found", "text/plain; charset=utf-8", b"Not Found")


application = app


# ==============================================================================
# STREAMLIT LOCAL RUNTIME
# ==============================================================================
def _is_running_in_streamlit() -> bool:
    try:
        import streamlit.runtime as st_rt
        return bool(st_rt.exists())
    except Exception:
        return False


def _streamlit_css(is_dark: bool) -> str:
    if is_dark:
        t = {
            "bg": "#0B1120",
            "sidebar": "#111827",
            "surface": "#111827",
            "surface_alt": "#1E293B",
            "border": "#475569",
            "text": "#F8FAFC",
            "muted": "#CBD5E1",
            "accent_bg": "#2563EB",
            "accent_hover": "#3B82F6",
            "accent_text": "#60A5FA",
            "accent_soft": "#1E293B",
            "hero_bg": "#0F172A",
            "hero_sub": "#CBD5E1",
            "danger": "#EF4444",
            "danger_bg": "#450A0A",
            "danger_border": "#991B1B",
            "danger_text": "#FECACA",
            "warn": "#F59E0B",
            "warn_bg": "#451A03",
            "warn_border": "#92400E",
            "warn_text": "#FDE68A",
            "ok": "#10B981",
            "ok_bg": "#022C22",
            "ok_border": "#065F46",
            "ok_text": "#A7F3D0",
        }
    else:
        t = {
            "bg": "#F8FAFC",
            "sidebar": "#FFFFFF",
            "surface": "#FFFFFF",
            "surface_alt": "#F1F5F9",
            "border": "#CBD5E1",
            "text": "#0F172A",
            "muted": "#475569",
            "accent_bg": "#1E40AF",
            "accent_hover": "#1D4ED8",
            "accent_text": "#1E40AF",
            "accent_soft": "#EFF6FF",
            "hero_bg": "#0F2544",
            "hero_sub": "#CBD5E1",
            "danger": "#DC2626",
            "danger_bg": "#FEF2F2",
            "danger_border": "#FECACA",
            "danger_text": "#7F1D1D",
            "warn": "#D97706",
            "warn_bg": "#FFFBEB",
            "warn_border": "#FDE68A",
            "warn_text": "#78350F",
            "ok": "#059669",
            "ok_bg": "#F0FDF4",
            "ok_border": "#BBF7D0",
            "ok_text": "#064E3B",
        }

    return f"""
    <style>
    .stApp {{
        background-color: {t['bg']};
        color: {t['text']};
    }}
    .stApp .stMarkdown h1, .stApp .stMarkdown h2, .stApp .stMarkdown h3,
    .stApp .stMarkdown h4, .stApp .stMarkdown h5, .stApp .stMarkdown p,
    .stApp .stMarkdown li, .stApp label[data-testid="stWidgetLabel"] p,
    .stApp div[data-testid="stCaptionContainer"] p,
    .stApp div[data-testid="stRadio"] label p,
    .stApp div[data-testid="stRadio"] label div,
    .stApp div[data-testid="stTickBarMin"],
    .stApp div[data-testid="stTickBarMax"],
    .stApp div[data-testid="stThumbValue"],
    .stApp div[data-testid="stExpander"] summary p {{
        color: {t['text']} !important;
    }}
    .stApp div[data-testid="stAlert"] {{
        background-color: {t['surface_alt']} !important;
        border: 1px solid {t['border']} !important;
        color: {t['text']} !important;
    }}
    .stApp div[data-testid="stAlert"] * {{
        color: {t['text']} !important;
    }}
    section[data-testid="stSidebar"] {{
        background-color: {t['sidebar']};
        border-right: 1px solid {t['border']};
        color: {t['text']};
    }}
    header[data-testid="stHeader"] {{
        background-color: transparent;
    }}
    .block-container {{
        padding-top: 2.3rem;
        padding-bottom: 2.5rem;
        max-width: 1380px;
    }}

    /* Form widgets dark/light contrast */
    .stApp div[data-baseweb="select"] > div,
    .stApp div[data-baseweb="input"] > div,
    .stApp .stNumberInput input,
    .stApp .stTextInput input {{
        background-color: {t['surface']} !important;
        color: {t['text']} !important;
        border-color: {t['border']} !important;
    }}
    .stApp div[data-baseweb="select"] span,
    .stApp div[data-baseweb="select"] div,
    .stApp div[data-baseweb="select"] svg {{
        color: {t['text']} !important;
        fill: {t['text']} !important;
    }}
    div[data-testid="stNumberInput"] button {{
        background-color: {t['surface_alt']} !important;
        color: {t['text']} !important;
        border-color: {t['border']} !important;
    }}
    div[data-testid="stVerticalBlockBorderWrapper"] {{
        background-color: {t['surface']};
        border-color: {t['border']} !important;
        border-radius: 8px;
    }}

    /* Explicit Primary & Secondary Button Contrast (Light + Dark) */
    .stApp button[kind="primary"],
    .stApp button[data-testid="stBaseButton-primary"],
    .stApp .stButton > button[kind="primary"],
    .stApp .stDownloadButton > button[kind="primary"] {{
        background-color: {t['accent_bg']} !important;
        border: 1px solid {t['accent_bg']} !important;
        color: #FFFFFF !important;
        border-radius: 6px !important;
        font-weight: 600 !important;
        min-height: 40px !important;
        height: 40px !important;
        white-space: nowrap !important;
        transition: all 0.16s ease;
    }}
    .stApp button[kind="primary"] *,
    .stApp button[data-testid="stBaseButton-primary"] *,
    .stApp .stButton > button[kind="primary"] *,
    .stApp .stDownloadButton > button[kind="primary"] * {{
        color: #FFFFFF !important;
    }}
    .stApp button[kind="primary"]:hover,
    .stApp button[data-testid="stBaseButton-primary"]:hover {{
        background-color: {t['accent_hover']} !important;
        border-color: {t['accent_hover']} !important;
    }}

    .stApp button[kind="secondary"],
    .stApp button[data-testid="stBaseButton-secondary"],
    .stApp .stButton > button[kind="secondary"],
    .stApp .stDownloadButton > button {{
        background-color: {t['surface']} !important;
        border: 1px solid {t['border']} !important;
        color: {t['text']} !important;
        border-radius: 6px !important;
        font-weight: 600 !important;
        min-height: 40px !important;
        height: 40px !important;
        white-space: nowrap !important;
        transition: all 0.16s ease;
    }}
    .stApp button[kind="secondary"] *,
    .stApp button[data-testid="stBaseButton-secondary"] *,
    .stApp .stButton > button[kind="secondary"] *,
    .stApp .stDownloadButton > button * {{
        color: {t['text']} !important;
    }}
    .stApp button[kind="secondary"]:hover,
    .stApp button[data-testid="stBaseButton-secondary"]:hover,
    .stApp .stDownloadButton > button:hover {{
        background-color: {t['accent_soft']} !important;
        border-color: {t['accent_text']} !important;
        color: {t['accent_text']} !important;
    }}
    .stApp button[kind="secondary"]:hover *,
    .stApp button[data-testid="stBaseButton-secondary"]:hover *,
    .stApp .stDownloadButton > button:hover * {{
        color: {t['accent_text']} !important;
    }}

    /* Drag & Drop File Uploader */
    .stApp section[data-testid="stFileUploaderDropzone"],
    .stApp div[data-testid="stFileUploader"] section {{
        background-color: {t['surface']} !important;
        color: {t['text']} !important;
        border: 2px dashed {t['border']} !important;
        border-radius: 8px !important;
        padding: 18px 16px !important;
        transition: all 0.18s ease;
    }}
    .stApp section[data-testid="stFileUploaderDropzone"]:hover,
    .stApp div[data-testid="stFileUploader"] section:hover {{
        border-color: {t['accent_text']} !important;
        background-color: {t['accent_soft']} !important;
    }}
    .stApp section[data-testid="stFileUploaderDropzone"] *,
    .stApp div[data-testid="stFileUploader"] section * {{
        color: {t['text']} !important;
    }}
    .stApp section[data-testid="stFileUploaderDropzone"] button {{
        background-color: {t['surface_alt']} !important;
        color: {t['text']} !important;
        border: 1px solid {t['border']} !important;
    }}

    /* Center-Screen Loading */
    div[data-testid="stStatusWidget"]:has(button) {{
        position: fixed !important;
        inset: 0 !important;
        width: 100vw !important;
        height: 100vh !important;
        background: rgba(15, 23, 42, 0.36) !important;
        backdrop-filter: blur(2px) !important;
        z-index: 999998 !important;
        display: flex !important;
        align-items: center !important;
        justify-content: center !important;
    }}
    div[data-testid="stStatusWidget"]:has(button) > div {{
        background: {t['surface']} !important;
        color: {t['text']} !important;
        border: 1px solid {t['border']} !important;
        border-radius: 10px !important;
        padding: 16px 24px !important;
        box-shadow: 0 12px 32px rgba(0, 0, 0, 0.25) !important;
    }}
    .center-loading-overlay {{
        position: fixed; inset: 0; width: 100vw; height: 100vh;
        background: rgba(15, 23, 42, 0.40); backdrop-filter: blur(2px);
        z-index: 999999; display: flex; align-items: center; justify-content: center;
    }}
    .center-loading-card {{
        background: {t['surface']}; color: {t['text']}; border: 1px solid {t['border']};
        border-radius: 10px; padding: 20px 28px; text-align: center;
        box-shadow: 0 12px 32px rgba(0, 0, 0, 0.28); min-width: 280px;
    }}
    .center-spinner {{
        width: 28px; height: 28px; border: 3px solid {t['border']}; border-top-color: {t['accent_text']};
        border-radius: 50%; margin: 0 auto 10px auto; animation: stspin 0.75s linear infinite;
    }}
    @keyframes stspin {{ to {{ transform: rotate(360deg); }} }}

    /* High-Specificity Tabs Styling (Streamlit 1.64 stTab + React Aria + BaseWeb fallback) */
    .stApp .stTabs [data-baseweb="tab-highlight"],
    .stApp .stTabs [data-baseweb="tab-border"],
    .stApp .stTabs [role="tablist"]::after,
    .stApp .stTabs [data-testid="stTab"]::after,
    .stApp .stTabs .react-aria-SelectionIndicator {{
        display: none !important;
        height: 0 !important;
        border: none !important;
        background: transparent !important;
    }}
    .stApp .stTabs [role="tablist"],
    .stApp .stTabs [data-baseweb="tab-list"] {{
        gap: 8px !important;
        background-color: {t['surface']} !important;
        padding: 6px 8px !important;
        border-radius: 10px !important;
        border: 1px solid {t['border']} !important;
        border-bottom: 1px solid {t['border']} !important;
        margin-bottom: 14px !important;
    }}
    .stApp .stTabs [data-testid="stTab"],
    .stApp .stTabs [role="tab"],
    .stApp .stTabs [data-baseweb="tab"] {{
        height: 40px !important;
        padding: 8px 15px !important;
        border-radius: 7px !important;
        font-weight: 600 !important;
        font-size: 13.5px !important;
        color: {t['muted']} !important;
        background-color: {t['surface_alt']} !important;
        border: 1px solid {t['border']} !important;
        margin: 0 !important;
        cursor: pointer !important;
        transition: all 0.18s ease !important;
    }}
    .stApp .stTabs [data-testid="stTab"] *,
    .stApp .stTabs [data-testid="stTab"] div[data-testid="stMarkdownContainer"] p,
    .stApp .stTabs [role="tab"] *,
    .stApp .stTabs [role="tab"] p,
    .stApp .stTabs [data-baseweb="tab"] *,
    .stApp .stTabs [data-baseweb="tab"] p,
    .stApp .stTabs [data-baseweb="tab"] span {{
        color: {t['muted']} !important;
    }}
    .stApp .stTabs [data-testid="stTab"]:hover,
    .stApp .stTabs [role="tab"]:hover,
    .stApp .stTabs [data-baseweb="tab"]:hover {{
        background-color: {t['accent_soft']} !important;
        color: {t['accent_text']} !important;
        border-color: {t['accent_text']} !important;
    }}
    .stApp .stTabs [data-testid="stTab"]:hover *,
    .stApp .stTabs [data-testid="stTab"]:hover div[data-testid="stMarkdownContainer"] p,
    .stApp .stTabs [data-testid="stTab"]:hover p,
    .stApp .stTabs [role="tab"]:hover *,
    .stApp .stTabs [role="tab"]:hover p,
    .stApp .stTabs [data-baseweb="tab"]:hover *,
    .stApp .stTabs [data-baseweb="tab"]:hover p {{
        color: {t['accent_text']} !important;
    }}
    .stApp .stTabs [data-testid="stTab"][aria-selected="true"],
    .stApp .stTabs [data-testid="stTab"][aria-selected="true"]:hover,
    .stApp .stTabs [role="tab"][aria-selected="true"],
    .stApp .stTabs [role="tab"][aria-selected="true"]:hover,
    .stApp .stTabs [data-baseweb="tab"][aria-selected="true"],
    .stApp .stTabs [data-baseweb="tab"][aria-selected="true"]:hover {{
        background-color: {t['accent_bg']} !important;
        color: #FFFFFF !important;
        border-color: {t['accent_bg']} !important;
    }}
    .stApp .stTabs [data-testid="stTab"][aria-selected="true"] *,
    .stApp .stTabs [data-testid="stTab"][aria-selected="true"]:hover *,
    .stApp .stTabs [data-testid="stTab"][aria-selected="true"] div[data-testid="stMarkdownContainer"] p,
    .stApp .stTabs [data-testid="stTab"][aria-selected="true"]:hover div[data-testid="stMarkdownContainer"] p,
    .stApp .stTabs [role="tab"][aria-selected="true"] *,
    .stApp .stTabs [role="tab"][aria-selected="true"]:hover *,
    .stApp .stTabs [role="tab"][aria-selected="true"] p,
    .stApp .stTabs [data-baseweb="tab"][aria-selected="true"] *,
    .stApp .stTabs [data-baseweb="tab"][aria-selected="true"]:hover * {{
        color: #FFFFFF !important;
    }}
    .tbl-wrap {{
        width: 100%;
        overflow-x: auto;
        border: 1px solid {t['border']};
        border-radius: 8px;
        background: {t['surface']};
        margin-bottom: 14px;
    }}
    .tbl-clean {{
        width: 100%;
        border-collapse: collapse;
        font-size: 12px;
        color: {t['text']};
    }}
    .tbl-clean th {{
        background: {t['surface_alt']};
        color: {t['muted']};
        font-weight: 700;
        text-align: left;
        padding: 8px 8px;
        border-bottom: 1px solid {t['border']};
        white-space: nowrap;
    }}
    .tbl-clean td {{
        padding: 8px 8px;
        border-bottom: 1px solid {t['border']};
        color: {t['text']};
    }}
    .tbl-clean tr:last-child td {{ border-bottom: none; }}
    .tbl-clean tr.hl-row td {{
        background: {t['accent_soft']};
        font-weight: 700;
    }}

    .stApp .stMarkdown .hero-banner,
    .stApp .hero-banner {{
        background: {t['hero_bg']};
        padding: 20px 26px;
        border-radius: 8px;
        border-bottom: 3px solid #D97706;
        margin-bottom: 16px;
    }}
    .stApp .stMarkdown .hero-banner h1,
    .stApp .hero-banner h1 {{
        color: #FFFFFF !important;
        margin: 0;
        font-size: 21px;
        font-weight: 700;
        line-height: 1.3;
    }}
    .stApp .stMarkdown .hero-banner p,
    .stApp .hero-banner p {{
        color: {t['hero_sub']} !important;
        margin: 5px 0 0 0;
        font-size: 13px;
    }}
    .kpi-card {{
        background: {t['surface']};
        border: 1px solid {t['border']};
        border-top: 3px solid {t['accent_bg']};
        border-radius: 8px;
        padding: 14px 16px;
        height: 120px;
        display: flex;
        flex-direction: column;
        justify-content: space-between;
    }}
    .kpi-card.amber {{ border-top-color: #D97706; }}
    .kpi-card.teal {{ border-top-color: #059669; }}
    .kpi-card.indigo {{ border-top-color: #6366F1; }}
    .kpi-label {{
        font-size: 11.5px;
        font-weight: 600;
        color: {t['muted']} !important;
        text-transform: uppercase;
        letter-spacing: 0.4px;
        margin-bottom: 4px;
        white-space: nowrap;
        overflow: hidden;
        text-overflow: ellipsis;
    }}
    .kpi-value {{
        font-size: clamp(16px, 1.4vw, 19.5px);
        font-weight: 700;
        color: {t['text']} !important;
        line-height: 1.2;
        white-space: nowrap;
        overflow: hidden;
        text-overflow: ellipsis;
    }}
    .kpi-sub {{
        font-size: 12px;
        color: {t['muted']} !important;
        margin-top: 4px;
        line-height: 1.3;
    }}
    .risk-box {{
        padding: 14px 16px;
        border-radius: 8px;
        border: 1px solid transparent;
        border-left-width: 4px;
        margin-bottom: 12px;
    }}
    .risk-high {{
        background-color: {t['danger_bg']};
        border-color: {t['danger_border']};
        border-left-color: {t['danger']};
        color: {t['danger_text']} !important;
    }}
    .risk-mid {{
        background-color: {t['warn_bg']};
        border-color: {t['warn_border']};
        border-left-color: {t['warn']};
        color: {t['warn_text']} !important;
    }}
    .risk-low {{
        background-color: {t['ok_bg']};
        border-color: {t['ok_border']};
        border-left-color: {t['ok']};
        color: {t['ok_text']} !important;
    }}
    .gauge-wrap {{
        background: {t['surface']};
        border: 1px solid {t['border']};
        border-radius: 8px;
        padding: 14px 16px;
        margin-bottom: 12px;
    }}
    .gauge-track {{
        position: relative;
        height: 14px;
        border-radius: 999px;
        background: linear-gradient(90deg, #10B981 0%, #F59E0B 45%, #EF4444 100%);
        margin: 12px 0 8px 0;
    }}
    .gauge-threshold {{
        position: absolute; top: -5px; bottom: -5px; width: 3px;
        background: {t['text']}; border-radius: 2px;
    }}
    .gauge-marker {{
        position: absolute; top: -6px; width: 12px; height: 26px; transform: translateX(-50%);
        background: {t['surface']}; border: 3px solid {t['accent_text']}; border-radius: 4px;
    }}
    .feat-chip {{
        background: {t['surface']};
        border: 1px solid {t['border']};
        border-radius: 6px;
        padding: 10px 6px;
        text-align: center;
        height: 76px;
        display: flex;
        flex-direction: column;
        justify-content: center;
    }}
    .feat-chip.selected {{
        border: 2px solid {t['accent_text']} !important;
        background: {t['accent_soft']} !important;
    }}
    .feat-chip-name {{
        font-size: 10.5px;
        color: {t['muted']} !important;
        font-weight: 600;
        line-height: 1.25;
        word-break: normal;
        overflow-wrap: break-word;
    }}
    .feat-chip-val {{
        font-size: 14.5px;
        font-weight: 700;
        color: {t['text']} !important;
        margin-top: 3px;
    }}
    .rec-item {{
        background: {t['surface_alt']};
        border: 1px solid {t['border']};
        border-left: 3px solid {t['accent_text']};
        border-radius: 6px;
        padding: 10px 14px;
        margin-bottom: 8px;
        font-size: 13.5px;
        color: {t['text']} !important;
    }}
    .sidebar-box {{
        background: {t['surface_alt']};
        border: 1px solid {t['border']};
        border-radius: 8px;
        padding: 12px 14px;
        margin-bottom: 12px;
        font-size: 13px;
        color: {t['text']} !important;
    }}
    /* Sidebar collapse / expand buttons */
    [data-testid="stSidebarCollapseButton"] button,
    [data-testid="stSidebarCollapsedControl"] button,
    [data-testid="collapsedControl"] button,
    [data-testid="stExpandSidebarButton"],
    [data-testid="stSidebarHeader"] button {{
        background: transparent !important;
        border: 1px solid {t['border']} !important;
        border-radius: 6px !important;
        color: {t['muted']} !important;
        box-shadow: none !important;
        width: 32px !important;
        height: 32px !important;
        min-height: 32px !important;
        padding: 0 !important;
        transition: background .15s, color .15s, border-color .15s;
    }}
    [data-testid="stSidebarCollapseButton"] button:hover,
    [data-testid="stSidebarCollapsedControl"] button:hover,
    [data-testid="collapsedControl"] button:hover,
    [data-testid="stExpandSidebarButton"]:hover,
    [data-testid="stSidebarHeader"] button:hover {{
        background: {t['accent_soft']} !important;
        border-color: {t['accent_text']} !important;
        color: {t['accent_text']} !important;
    }}
    [data-testid="stSidebarCollapseButton"] svg,
    [data-testid="stSidebarCollapsedControl"] svg,
    [data-testid="collapsedControl"] svg,
    [data-testid="stExpandSidebarButton"] svg,
    [data-testid="stSidebarHeader"] button svg,
    [data-testid="stSidebarHeader"] button span {{
        color: inherit !important;
        fill: currentColor !important;
    }}
    header[data-testid="stHeader"] {{
        background: {t['bg']} !important;
    }}
    header[data-testid="stHeader"] button {{
        color: {t['muted']} !important;
    }}
    </style>
    """


def run_streamlit_app():
    import io
    import time
    import joblib
    import numpy as np
    import pandas as pd
    import matplotlib
    matplotlib.use("Agg")
    import matplotlib.pyplot as plt
    import shap
    import streamlit as st
    from src.data_loader import (
        BINARY_NUMERIC_FEATURES,
        CATEGORICAL_FEATURES,
        NUMERIC_FEATURES,
        add_domain_features,
    )

    st.set_page_config(
        page_title="Customer Churn Analytics — Dự đoán Khách hàng Rời bỏ Dịch vụ",
        layout="wide",
        initial_sidebar_state="expanded",
    )

    if "ui_theme" not in st.session_state:
        st.session_state["ui_theme"] = "Sáng (Light)"

    default = PRESET_PROFILES["high"]
    preset_keymap = {
        "gender": "gender",
        "senior": "SeniorCitizen",
        "partner": "Partner",
        "dependents": "Dependents",
        "phone_service": "PhoneService",
        "multiple_lines": "MultipleLines",
        "internet": "InternetService",
        "sec": "OnlineSecurity",
        "backup": "OnlineBackup",
        "prot": "DeviceProtection",
        "tech": "TechSupport",
        "tv": "StreamingTV",
        "movies": "StreamingMovies",
        "contract": "Contract",
        "tenure": "tenure",
        "monthly": "MonthlyCharges",
        "total": "TotalCharges",
        "paperless": "PaperlessBilling",
        "payment": "PaymentMethod",
    }
    for ui_k, raw_k in preset_keymap.items():
        if f"prof_{ui_k}" not in st.session_state:
            st.session_state[f"prof_{ui_k}"] = default[raw_k]

    def apply_preset(preset_key: str):
        p = PRESET_PROFILES[preset_key]
        for ui_k, raw_k in preset_keymap.items():
            st.session_state[f"prof_{ui_k}"] = p[raw_k]

    def set_threshold(val: float):
        st.session_state["threshold_slider"] = float(val)

    def show_center_loading(message: str, duration_sec: float = 0.35):
        ph = st.empty()
        ph.markdown(
            f"""
            <div class="center-loading-overlay">
                <div class="center-loading-card">
                    <div class="center-spinner"></div>
                    <div style="font-weight:600; font-size:13.5px;">{message}</div>
                </div>
            </div>
            """,
            unsafe_allow_html=True,
        )
        time.sleep(duration_sec)
        ph.empty()

    @st.cache_data
    def load_random_sample_csv(n_rows: int, seed: int) -> str:
        if os.path.exists(RAW_CSV_PATH):
            return generate_random_sample_csv(RAW_CSV_PATH, n_rows=n_rows, seed=seed)
        if os.path.exists(SAMPLE_CSV_PATH):
            with open(SAMPLE_CSV_PATH, "r", encoding="utf-8-sig") as f:
                return f.read()
        return ""

    @st.cache_resource
    def load_artifacts():
        bundle = joblib.load(MODEL_PATH)
        with open(RESULTS_PATH, "r", encoding="utf-8") as f:
            results = json.load(f)
        comp_df = pd.read_csv(COMP_CSV_PATH)
        explainers = {}
        all_pipes = bundle.get("all_pipelines", {bundle["best_model_name"]: bundle["best_model_pipeline"]})
        for mname, pipe in all_pipes.items():
            if mname != "Logistic Regression":
                explainers[mname] = shap.TreeExplainer(pipe.named_steps["clf"])
        return bundle, results, comp_df, explainers

    def plot_local_shap(shap_vals_row, feature_names, is_dark: bool, title_txt: str, xlabel_txt: str):
        bg = "#111827" if is_dark else "#FFFFFF"
        fg = "#F8FAFC" if is_dark else "#0F2544"
        sub = "#CBD5E1" if is_dark else "#475569"
        grid_c = "#334155" if is_dark else "#CBD5E1"
        pos_c = "#EF4444" if is_dark else "#DC2626"
        neg_c = "#10B981" if is_dark else "#059669"

        top_k = np.argsort(np.abs(shap_vals_row))[::-1][:8][::-1]
        labels = [FEATURE_VI_MAP.get(feature_names[j], feature_names[j]) for j in top_k]
        vals = [float(shap_vals_row[j]) for j in top_k]
        colors = [pos_c if v > 0 else neg_c for v in vals]

        fig, ax = plt.subplots(figsize=(7.4, 4.4), dpi=160)
        fig.patch.set_facecolor(bg)
        ax.set_facecolor(bg)

        bars = ax.barh(labels, vals, color=colors, height=0.58, edgecolor="none")
        ax.axvline(0, color=fg, lw=1.2)
        span = max(abs(min(vals)), abs(max(vals)), 0.05) + 0.10
        ax.set_xlim(-span * 1.32, span * 1.32)

        for b, v in zip(bars, vals):
            ha = "left" if v >= 0 else "right"
            off = span * 0.035 if v >= 0 else -span * 0.035
            ax.text(
                v + off,
                b.get_y() + b.get_height() / 2,
                f"{v:+.3f}",
                va="center",
                ha=ha,
                fontsize=8.5,
                fontweight="bold",
                color=fg,
            )

        ax.set_title(title_txt, fontsize=10.5, fontweight="bold", color=fg, pad=10)
        ax.set_xlabel(xlabel_txt, fontsize=8.8, color=sub)
        ax.tick_params(axis="y", labelsize=8.8, colors=fg)
        ax.tick_params(axis="x", labelsize=8.3, colors=sub)
        ax.xaxis.grid(True, linestyle="--", alpha=0.55, color=grid_c)
        ax.set_axisbelow(True)
        ax.spines["top"].set_visible(False)
        ax.spines["right"].set_visible(False)
        ax.spines["left"].set_color(grid_c)
        ax.spines["bottom"].set_color(grid_c)
        plt.tight_layout()
        return fig

    def plot_batch_distribution(probs: np.ndarray, threshold: float, is_dark: bool):
        bg = "#111827" if is_dark else "#FFFFFF"
        fg = "#F8FAFC" if is_dark else "#0F2544"
        sub = "#CBD5E1" if is_dark else "#475569"
        grid_c = "#334155" if is_dark else "#CBD5E1"

        fig, (ax1, ax2) = plt.subplots(1, 2, figsize=(11.0, 3.6), dpi=150, gridspec_kw={"width_ratios": [1.45, 1.0]})
        fig.patch.set_facecolor(bg)

        ax1.set_facecolor(bg)
        bins = np.linspace(0, 100, 21)
        _, edges, patches = ax1.hist(probs * 100, bins=bins, edgecolor=bg, linewidth=1.0)
        for left, p in zip(edges[:-1], patches):
            if left >= 50:
                p.set_facecolor("#DC2626")
            elif left >= threshold * 100:
                p.set_facecolor("#F59E0B")
            else:
                p.set_facecolor("#10B981")
        ax1.axvline(threshold * 100, color=fg, linestyle="--", linewidth=1.8, label=f"Ngưỡng T = {threshold:.2f}")
        ax1.set_title("Phân phối Xác suất Rời bỏ (%)", fontsize=10, fontweight="bold", color=fg)
        ax1.set_xlabel("Xác suất rời bỏ (%)", fontsize=8.8, color=sub)
        ax1.set_ylabel("Số lượng khách hàng", fontsize=8.8, color=sub)
        leg = ax1.legend(frameon=True, facecolor=bg, edgecolor=grid_c, fontsize=8.5)
        for text in leg.get_texts():
            text.set_color(fg)
        ax1.tick_params(colors=sub)
        ax1.yaxis.grid(True, linestyle="--", alpha=0.45, color=grid_c)
        ax1.set_axisbelow(True)
        ax1.spines["top"].set_visible(False)
        ax1.spines["right"].set_visible(False)
        ax1.spines["left"].set_color(grid_c)
        ax1.spines["bottom"].set_color(grid_c)

        ax2.set_facecolor(bg)
        n_high = int((probs >= 0.50).sum())
        n_mid = int(((probs >= threshold) & (probs < 0.50)).sum())
        n_low = int((probs < threshold).sum())
        tier_labels = [
            f"Ổn định (<{threshold*100:.0f}%)",
            f"Cảnh báo ({threshold*100:.0f}%–50%)",
            "Nguy cơ cao (>=50%)",
        ]
        tier_counts = [n_low, n_mid, n_high]
        tier_colors = ["#10B981", "#F59E0B", "#DC2626"]
        bars = ax2.barh(tier_labels, tier_counts, color=tier_colors, height=0.52)
        max_c = max(tier_counts) if max(tier_counts) > 0 else 1
        ax2.set_xlim(0, max_c * 1.28)
        total_n = len(probs) if len(probs) > 0 else 1
        for b, cnt in zip(bars, tier_counts):
            ax2.text(
                cnt + max_c * 0.03,
                b.get_y() + b.get_height() / 2,
                f"{cnt} ({cnt/total_n*100:.1f}%)",
                va="center",
                fontsize=8.8,
                fontweight="bold",
                color=fg,
            )
        ax2.set_title("Phân tầng Hành động CSKH", fontsize=10, fontweight="bold", color=fg)
        ax2.set_xlabel("Số khách hàng", fontsize=8.8, color=sub)
        ax2.tick_params(colors=sub)
        ax2.xaxis.grid(True, linestyle="--", alpha=0.45, color=grid_c)
        ax2.set_axisbelow(True)
        ax2.spines["top"].set_visible(False)
        ax2.spines["right"].set_visible(False)
        ax2.spines["left"].set_color(grid_c)
        ax2.spines["bottom"].set_color(grid_c)

        plt.tight_layout()
        return fig

    def plot_live_cost_curve(cost_curve: list, opt_thr: float, cur_thr: float, model_name: str, is_dark: bool):
        bg = "#111827" if is_dark else "#FFFFFF"
        fg = "#F8FAFC" if is_dark else "#0F2544"
        sub = "#CBD5E1" if is_dark else "#475569"
        grid_c = "#334155" if is_dark else "#CBD5E1"
        line_c = "#60A5FA" if is_dark else "#1E40AF"

        thrs = [p["threshold"] for p in cost_curve]
        costs = [p["total_cost_usd"] / 1000.0 for p in cost_curve]
        opt_idx = int(np.argmin(costs))
        cur_idx = min(range(len(thrs)), key=lambda i: abs(thrs[i] - cur_thr))

        fig, ax = plt.subplots(figsize=(6.8, 3.8), dpi=150)
        fig.patch.set_facecolor(bg)
        ax.set_facecolor(bg)
        ax.plot(thrs, costs, color=line_c, lw=2.4, label=f"Tổng chi phí ({model_name})")
        ax.axvline(opt_thr, color="#10B981", ls="--", lw=1.8, label=f"T* tối ưu = {opt_thr:.2f} (${costs[opt_idx]:.1f}k)")
        ax.axvline(cur_thr, color="#EF4444", ls="-", lw=1.8, label=f"T đang chọn = {cur_thr:.2f} (${costs[cur_idx]:.1f}k)")
        ax.scatter([opt_thr, cur_thr], [costs[opt_idx], costs[cur_idx]], color=["#10B981", "#EF4444"], zorder=5, s=48)
        ax.set_title("Đường cong Chi phí theo Ngưỡng T (Tập Test N = 1.409)", fontsize=10, fontweight="bold", color=fg)
        ax.set_xlabel("Ngưỡng phân loại (T)", fontsize=8.8, color=sub)
        ax.set_ylabel("Tổng chi phí (Nghìn USD)", fontsize=8.8, color=sub)
        leg = ax.legend(frameon=True, facecolor=bg, edgecolor=grid_c, fontsize=8.2)
        for text in leg.get_texts():
            text.set_color(fg)
        ax.tick_params(colors=sub)
        ax.grid(True, linestyle="--", alpha=0.45, color=grid_c)
        ax.spines["top"].set_visible(False)
        ax.spines["right"].set_visible(False)
        ax.spines["left"].set_color(grid_c)
        ax.spines["bottom"].set_color(grid_c)
        plt.tight_layout()
        return fig

    def plot_live_confusion_matrix(live_m: dict, model_name: str, cur_thr: float, c_fn: float, c_fp: float, is_dark: bool):
        bg = "#111827" if is_dark else "#FFFFFF"
        fg = "#F8FAFC" if is_dark else "#0F2544"
        sub = "#CBD5E1" if is_dark else "#475569"

        mat = np.array([[live_m["tn"], live_m["fp"]], [live_m["fn"], live_m["tp"]]])
        labels = [
            [f"TN = {live_m['tn']}\n(Chi phí $0)", f"FP = {live_m['fp']}\n(${live_m['fp']*c_fp:,.0f})"],
            [f"FN = {live_m['fn']}\n(${live_m['fn']*c_fn:,.0f})", f"TP = {live_m['tp']}\n(${live_m['tp']*c_fp:,.0f})"],
        ]
        fig, ax = plt.subplots(figsize=(5.0, 3.8), dpi=150)
        fig.patch.set_facecolor(bg)
        ax.set_facecolor(bg)
        cmap = "Blues"
        ax.imshow(mat, cmap=cmap, aspect="auto", vmin=0, vmax=max(1, int(mat.max())))
        ax.set_xticks([0, 1])
        ax.set_xticklabels(["Dự đoán: Ở lại (0)", "Dự đoán: Rời bỏ (1)"], color=fg, fontsize=8.8)
        ax.set_yticks([0, 1])
        ax.set_yticklabels(["Thực tế: Ở lại (0)", "Thực tế: Rời bỏ (1)"], color=fg, fontsize=8.8)
        thresh_val = mat.max() / 2.0
        for i in range(2):
            for j in range(2):
                txt_color = "#FFFFFF" if mat[i, j] > thresh_val else "#0F172A"
                ax.text(j, i, labels[i][j], ha="center", va="center", color=txt_color, fontsize=9.5, fontweight="bold")
        ax.set_title(f"Ma trận nhầm lẫn ({model_name} @ T={cur_thr:.2f})", fontsize=9.8, fontweight="bold", color=fg)
        for spine in ax.spines.values():
            spine.set_edgecolor(sub)
        plt.tight_layout()
        return fig

    if not os.path.exists(MODEL_PATH):
        st.error("Chưa tìm thấy mô hình đã huấn luyện. Vui lòng chạy `python src/train_pipeline.py` trước.")
        return

    bundle, results, comp_df, explainers = load_artifacts()
    all_pipes = bundle.get("all_pipelines", {bundle["best_model_name"]: bundle["best_model_pipeline"]})
    available_models = [m for m in MODEL_NAMES if m in all_pipes]
    tree_feat_names = bundle["tree_feature_names"]
    train_feat_means = np.array(bundle.get("train_feature_means", [0.0] * len(tree_feat_names)), dtype=float)
    y_train_arr = list(bundle["y_train"])
    y_test_arr = list(bundle["y_test"])
    oof_probs_all = {k: list(v) for k, v in bundle["oof_probs_all"].items()}
    test_probs_all = {k: list(v) for k, v in bundle["test_probs_all"].items()}
    feature_cols = NUMERIC_FEATURES + BINARY_NUMERIC_FEATURES + CATEGORICAL_FEATURES

    if "selected_model" not in st.session_state:
        st.session_state["selected_model"] = available_models[0]
    if "c_fn_input" not in st.session_state:
        st.session_state["c_fn_input"] = 500
    if "c_fp_input" not in st.session_state:
        st.session_state["c_fp_input"] = 100
    if "active_main_tab" not in st.session_state:
        st.session_state["active_main_tab"] = "tab1"
    if "active_sub_tab" not in st.session_state:
        st.session_state["active_sub_tab"] = "sub0"
    if "threshold_slider" not in st.session_state:
        st.session_state["threshold_slider"] = float(bundle["optimal_threshold"])

    with st.sidebar:
        st.markdown("### Giao diện hiển thị")
        ui_theme = st.radio(
            "Chế độ màu",
            ["Sáng (Light)", "Tối (Dark)"],
            horizontal=True,
            key="ui_theme",
            label_visibility="collapsed",
        )
        is_dark = ui_theme == "Tối (Dark)"

        st.divider()
        st.markdown("### Chọn mô hình & Chi phí")
        selected_model = st.selectbox("Thuật toán phân loại", available_models, key="selected_model")
        c_col1, c_col2 = st.columns(2)
        with c_col1:
            c_fn = float(st.number_input("C_FN (USD)", min_value=50, max_value=5000, step=50, key="c_fn_input"))
        with c_col2:
            c_fp = float(st.number_input("C_FP (USD)", min_value=10, max_value=1000, step=10, key="c_fp_input"))

        opt_thr, _ = find_optimal_threshold(y_train_arr, oof_probs_all[selected_model], c_fn=c_fn, c_fp=c_fp)
        _, cost_curve = find_optimal_threshold(y_test_arr, test_probs_all[selected_model], c_fn=c_fn, c_fp=c_fp)
        bayes_thr = min(0.99, c_fp / c_fn) if c_fn > 0 else 0.50
        ratio = c_fn / c_fp if c_fp > 0 else 1.0

        sel_cv_auc = results["cv_roc_auc"].get(selected_model, 0.0)
        sel_m05 = results["test_metrics_at_050"].get(selected_model, {})

        st.markdown(
            f"""
            <div class="sidebar-box">
                <strong>Mô hình:</strong> {selected_model}<br>
                <strong>CV / Test AUC:</strong> {sel_cv_auc:.4f} / {sel_m05.get('roc_auc', 0):.4f}<br>
                <strong>Test PR-AUC:</strong> {sel_m05.get('pr_auc', 0):.4f}<br>
                <hr style="margin:8px 0; border-color:rgba(148,163,184,0.3);">
                <strong>Ma trận chi phí động:</strong><br>
                • Bỏ sót Churn (C_FN): <b>{c_fn:,.0f} USD</b><br>
                • Tặng ưu đãi (C_FP/TP): <b>{c_fp:,.0f} USD</b><br>
                • Tỷ lệ đánh đổi: <b>{ratio:.1f} : 1</b><br>
                • T_Bayes: <b>{bayes_thr:.2f}</b> | T* OOF: <b>{opt_thr:.2f}</b>
            </div>
            """,
            unsafe_allow_html=True,
        )

        st.markdown("### Ngưỡng quyết định (T)")
        threshold = st.slider(
            "Ngưỡng phân loại rủi ro (T)",
            min_value=0.05,
            max_value=0.85,
            step=0.01,
            key="threshold_slider",
        )
        sb_c1, sb_c2 = st.columns(2)
        with sb_c1:
            st.button(
                f"T* = {opt_thr:.2f}",
                width="stretch",
                type="primary",
                on_click=set_threshold,
                args=(opt_thr,),
            )
        with sb_c2:
            st.button(
                "T = 0.50",
                width="stretch",
                on_click=set_threshold,
                args=(0.50,),
            )

    st.markdown(_streamlit_css(is_dark), unsafe_allow_html=True)

    # Trigger center loading on threshold / model / cost changes
    cur_ctrl_sig = f"{selected_model}|{threshold:.2f}|{c_fn:.0f}|{c_fp:.0f}"
    if "last_ctrl_sig" not in st.session_state:
        st.session_state["last_ctrl_sig"] = cur_ctrl_sig
    elif st.session_state["last_ctrl_sig"] != cur_ctrl_sig:
        st.session_state["last_ctrl_sig"] = cur_ctrl_sig
        show_center_loading(f"Đang cập nhật {selected_model} (T = {threshold:.2f}, C_FN = {c_fn:.0f}, C_FP = {c_fp:.0f})...", 0.30)

    # ==================== HERO BANNER ====================
    st.markdown(
        """
        <div class="hero-banner">
            <h1>Hệ thống Dự đoán và Phân tích Khách hàng Rời bỏ Dịch vụ (Customer Churn Analytics)</h1>
            <p>Đối chiếu 5 mô hình phân loại nhạy chi phí và giải thích cục bộ TreeSHAP / Log-Odds</p>
        </div>
        """,
        unsafe_allow_html=True,
    )

    # ==================== 4 KPI CARDS (LIVE) ====================
    audit = results["dataset_audit"]
    live_test_m = compute_threshold_metrics(y_test_arr, test_probs_all[selected_model], threshold, c_fn=c_fn, c_fp=c_fp)

    k1, k2, k3, k4 = st.columns(4)
    with k1:
        st.markdown(
            f"""
            <div class="kpi-card">
                <div class="kpi-label">Quy mô Dữ liệu Telco</div>
                <div class="kpi-value">{audit['n_rows']:,} thuê bao</div>
                <div class="kpi-sub">{audit['n_raw_cols']} cột gốc -> {audit['n_features_engineered']} đặc trưng • Churn {audit['churn_rate']*100:.2f}%</div>
            </div>
            """,
            unsafe_allow_html=True,
        )
    with k2:
        st.markdown(
            f"""
            <div class="kpi-card indigo">
                <div class="kpi-label">Mô hình Đang chọn</div>
                <div class="kpi-value">{selected_model}</div>
                <div class="kpi-sub">CV AUC: {sel_cv_auc:.4f} • Test AUC: {sel_m05.get('roc_auc', 0):.4f}</div>
            </div>
            """,
            unsafe_allow_html=True,
        )
    with k3:
        st.markdown(
            f"""
            <div class="kpi-card amber">
                <div class="kpi-label">Ngưỡng T / T* Tối ưu</div>
                <div class="kpi-value">T = {threshold:.2f} (T*={opt_thr:.2f})</div>
                <div class="kpi-sub">Recall Test: {live_test_m['recall']*100:.2f}% (giữ {live_test_m['tp']}/374 Churn)</div>
            </div>
            """,
            unsafe_allow_html=True,
        )
    with k4:
        st.markdown(
            f"""
            <div class="kpi-card teal">
                <div class="kpi-label">Hiệu quả Kinh tế (ROI)</div>
                <div class="kpi-value">Tiết kiệm ${live_test_m['savings_usd']:,.0f}</div>
                <div class="kpi-sub">Giảm {live_test_m['savings_pct']:.1f}% chi phí so với không can thiệp</div>
            </div>
            """,
            unsafe_allow_html=True,
        )

    st.markdown("<div style='height: 10px;'></div>", unsafe_allow_html=True)

    # ==================== MAIN TABS ====================
    def _set_main_tab(k):
        st.session_state["active_main_tab"] = k
    _main_labels = ["1. Dự đoán khách hàng cá nhân", "2. Chấm điểm danh sách (CSV)", "3. Kết quả thực nghiệm mô hình"]
    _main_cols = st.columns(3)
    for _i, _lbl in enumerate(_main_labels):
        _k = "tab%d" % (_i + 1)
        _main_cols[_i].button(_lbl, key="nav_" + _k, use_container_width=True,
            type="primary" if st.session_state["active_main_tab"] == _k else "secondary",
            on_click=_set_main_tab, args=(_k,))
    active_tab = st.session_state["active_main_tab"]

    # ==================== TAB 1 ====================
    if active_tab == "tab1":
        st.markdown("#### Nạp nhanh hồ sơ khách hàng mẫu")
        p1, p2, p3 = st.columns(3)
        with p1:
            st.button("Hồ sơ 1: Nguy cơ cao (~87%)", width="stretch", on_click=apply_preset, args=("high",))
        with p2:
            st.button("Hồ sơ 2: Trung bình (~30%)", width="stretch", on_click=apply_preset, args=("medium",))
        with p3:
            st.button("Hồ sơ 3: Ổn định (~2%)", width="stretch", on_click=apply_preset, args=("low",))

        st.markdown("#### Thông tin hồ sơ khách hàng (Tự động chấm điểm tức thì)")
        c1, c2, c3 = st.columns(3)

        with c1:
            with st.container(border=True):
                st.markdown("**Nhân khẩu học & Điện thoại**")
                gender = st.selectbox("Giới tính (gender)", ["Male", "Female"], key="prof_gender")
                senior = st.selectbox(
                    "Khách hàng cao tuổi (SeniorCitizen)",
                    [0, 1],
                    format_func=lambda x: "Có" if x == 1 else "Không",
                    key="prof_senior",
                )
                partner = st.selectbox("Có bạn đời (Partner)", ["No", "Yes"], key="prof_partner")
                dependents = st.selectbox("Người phụ thuộc (Dependents)", ["No", "Yes"], key="prof_dependents")
                phone_service = st.selectbox("Dịch vụ điện thoại (PhoneService)", ["Yes", "No"], key="prof_phone_service")
                multiple_lines = st.selectbox(
                    "Nhiều đường dây (MultipleLines)",
                    ["No", "Yes", "No phone service"],
                    key="prof_multiple_lines",
                )

        with c2:
            with st.container(border=True):
                st.markdown("**Dịch vụ Internet & Gia tăng**")
                internet = st.selectbox("Hạ tầng Internet (InternetService)", ["Fiber optic", "DSL", "No"], key="prof_internet")
                sec = st.selectbox("Bảo mật trực tuyến (OnlineSecurity)", ["No", "Yes", "No internet service"], key="prof_sec")
                backup = st.selectbox("Sao lưu trực tuyến (OnlineBackup)", ["No", "Yes", "No internet service"], key="prof_backup")
                prot = st.selectbox("Bảo vệ thiết bị (DeviceProtection)", ["No", "Yes", "No internet service"], key="prof_prot")
                tech = st.selectbox("Hỗ trợ kỹ thuật (TechSupport)", ["No", "Yes", "No internet service"], key="prof_tech")
                tv = st.selectbox("Truyền hình (StreamingTV)", ["No", "Yes", "No internet service"], key="prof_tv")
                movies = st.selectbox("Phim trực tuyến (StreamingMovies)", ["No", "Yes", "No internet service"], key="prof_movies")

        with c3:
            with st.container(border=True):
                st.markdown("**Hợp đồng & Cước phí**")
                contract = st.selectbox("Loại hợp đồng (Contract)", ["Month-to-month", "One year", "Two year"], key="prof_contract")
                tenure = st.slider("Số tháng gắn bó (tenure)", 0, 72, key="prof_tenure")
                monthly = st.number_input("Cước hàng tháng (MonthlyCharges - USD)", min_value=18.0, max_value=130.0, step=1.0, key="prof_monthly")
                total = st.number_input("Tổng cước tích lũy (TotalCharges - USD)", min_value=0.0, max_value=10000.0, step=10.0, key="prof_total")
                st.caption(f"Gợi ý theo kỳ hạn: ~{monthly * max(tenure, 1):,.2f} USD")
                paperless = st.selectbox("Hóa đơn điện tử (PaperlessBilling)", ["Yes", "No"], key="prof_paperless")
                payment = st.selectbox(
                    "Phương thức thanh toán (PaymentMethod)",
                    ["Electronic check", "Mailed check", "Bank transfer (automatic)", "Credit card (automatic)"],
                    key="prof_payment",
                )

        row_raw = {
            "gender": gender, "SeniorCitizen": senior, "Partner": partner, "Dependents": dependents,
            "tenure": tenure, "PhoneService": phone_service, "MultipleLines": multiple_lines,
            "InternetService": internet, "OnlineSecurity": sec, "OnlineBackup": backup,
            "DeviceProtection": prot, "TechSupport": tech, "StreamingTV": tv, "StreamingMovies": movies,
            "Contract": contract, "PaperlessBilling": paperless, "PaymentMethod": payment,
            "MonthlyCharges": monthly, "TotalCharges": total,
        }
        df_one = add_domain_features(pd.DataFrame([row_raw]))
        X_one = df_one[feature_cols]
        active_pipe = all_pipes[selected_model]
        prob = float(active_pipe.predict_proba(X_one)[0, 1])
        is_risk = prob >= threshold
        prob_pct = prob * 100
        thr_pct = threshold * 100

        st.markdown("##### 6 Đặc trưng miền tự động trích xuất")
        fe_row = df_one.iloc[0]
        f_cols = st.columns(6)
        chips = [
            ("Chi tiêu TB/tháng", f"${fe_row['AvgMonthlySpend']:.2f}"),
            ("Chênh lệch cước", f"{fe_row['ChargeDiff']:+.2f} USD"),
            ("Dịch vụ cộng thêm", f"{int(fe_row['TotalAddonServices'])} / 6"),
            ("Bảo mật / HTKT", "Có" if int(fe_row["HasProtectionBundle"]) == 1 else "Không"),
            ("Thanh toán tự động", "Tự động" if int(fe_row["IsAutoPayment"]) == 1 else "Thủ công"),
            ("Nhóm vòng đời", str(fe_row["TenureGroup"])),
        ]
        for col_obj, (cname, cval) in zip(f_cols, chips):
            col_obj.markdown(
                f'<div class="feat-chip"><div class="feat-chip-name">{cname}</div>'
                f'<div class="feat-chip-val">{cval}</div></div>',
                unsafe_allow_html=True,
            )

        st.markdown("##### Đối chiếu xác suất rời bỏ của cả 5 mô hình trên hồ sơ này")
        m_cols = st.columns(len(available_models))
        for col_obj, mname in zip(m_cols, available_models):
            mp = float(all_pipes[mname].predict_proba(X_one)[0, 1])
            sel_cls = "feat-chip selected" if mname == selected_model else "feat-chip"
            tag = "Cảnh báo" if mp >= threshold else "Ổn định"
            col_obj.markdown(
                f'<div class="{sel_cls}" style="height:88px;"><div class="feat-chip-name">{mname}</div>'
                f'<div class="feat-chip-val">{mp*100:.2f}%</div>'
                f'<div class="feat-chip-name">{tag} (T={threshold:.2f})</div></div>',
                unsafe_allow_html=True,
            )

        st.markdown("<div style='height: 12px;'></div>", unsafe_allow_html=True)

        r1, r2 = st.columns([1.0, 1.32])
        with r1:
            if prob >= 0.50:
                risk_css = "risk-high"
                risk_title = "CẢNH BÁO NGUY CƠ RỜI BỎ RẤT CAO"
                risk_desc = f"Xác suất <b>{prob_pct:.2f}%</b> vượt ngưỡng T = {threshold:.2f} và vượt mốc 50%. Cần can thiệp khẩn cấp."
            elif is_risk:
                risk_css = "risk-mid"
                risk_title = "CẢNH BÁO RỦI RO TRUNG GIAN (BẮT NHỜ NGƯỠNG T)"
                risk_desc = (
                    f"Xác suất <b>{prob_pct:.2f}%</b> vượt ngưỡng nhạy chi phí T = {threshold:.2f} "
                    f"(bị bỏ sót nếu dùng ngưỡng mặc định 0.50)."
                )
            else:
                risk_css = "risk-low"
                risk_title = "NHÓM KHÁCH HÀNG ỔN ĐỊNH"
                risk_desc = f"Xác suất <b>{prob_pct:.2f}%</b> thấp hơn ngưỡng cảnh báo T = {threshold:.2f}."

            st.markdown(
                f"""
                <div class="gauge-wrap">
                    <div style="display:flex; justify-content:space-between; align-items:baseline;">
                        <span style="font-size:13px; font-weight:600;">Xác suất Rời bỏ ({selected_model})</span>
                        <span style="font-size:28px; font-weight:800;">{prob_pct:.2f}%</span>
                    </div>
                    <div class="gauge-track">
                        <div class="gauge-threshold" style="left:{thr_pct:.1f}%;"></div>
                        <div class="gauge-marker" style="left:{min(max(prob_pct, 1.0), 99.0):.1f}%;"></div>
                    </div>
                    <div style="display:flex; justify-content:space-between; font-size:11.5px; opacity:0.85;">
                        <span>0% (An toàn)</span>
                        <span style="font-weight:700;">Ngưỡng T = {thr_pct:.0f}%</span>
                        <span>100% (Rời bỏ)</span>
                    </div>
                </div>
                <div class="risk-box {risk_css}">
                    <strong>{risk_title}</strong><br>{risk_desc}
                </div>
                """,
                unsafe_allow_html=True,
            )

            st.markdown("##### Khuyến nghị hành động CSKH")
            for rec in build_recommendations(row_raw, prob, threshold):
                st.markdown(f'<div class="rec-item">{rec}</div>', unsafe_allow_html=True)

        with r2:
            with st.container(border=True):
                X_one_trans = active_pipe.named_steps["prep"].transform(X_one)
                if selected_model == "Logistic Regression":
                    coef = active_pipe.named_steps["clf"].coef_[0]
                    sv_row = coef * (X_one_trans[0] - train_feat_means)
                    t_txt = "Đóng góp Log-Odds (Logistic Regression) — Top 8 yếu tố"
                    x_txt = "Đóng góp Log-Odds (Đỏ: tăng rủi ro | Xanh: giảm rủi ro)"
                else:
                    sv_raw = explainers[selected_model].shap_values(X_one_trans)
                    if isinstance(sv_raw, list):
                        sv_row = np.array(sv_raw[1])[0]
                    elif sv_raw.ndim == 3:
                        sv_row = sv_raw[0, :, 1]
                    else:
                        sv_row = np.array(sv_raw)[0]
                    t_txt = f"Giải thích TreeSHAP ({selected_model}) — Top 8 yếu tố"
                    x_txt = "Đóng góp SHAP (Đỏ: tăng rủi ro | Xanh: giảm rủi ro)"
                fig_local = plot_local_shap(sv_row, tree_feat_names, is_dark, t_txt, x_txt)
                st.pyplot(fig_local, width="stretch")
                plt.close(fig_local)

    # ==================== TAB 2 ====================
    elif active_tab == "tab2":
        st.markdown("#### Chấm điểm rủi ro hàng loạt & Phân tầng giữ chân khách hàng")

        for k, v in [
            ("sample_n", 100),
            ("sample_seed", 42),
            ("uploader_key", 0),
            ("batch_scored_raw", None),
            ("batch_scored_probs", None),
            ("batch_scored_sig", ""),
            ("batch_scored_source", ""),
        ]:
            if k not in st.session_state:
                st.session_state[k] = v

        uploaded = st.file_uploader(
            "Thêm hoặc kéo thả tệp CSV danh sách khách hàng vào đây (Drag & Drop)",
            type=["csv"],
            key=f"csv_uploader_{st.session_state['uploader_key']}",
            help="Hỗ trợ kéo thả file .csv. Sau khi chọn tệp, bấm nút 'Chấm điểm' bên dưới để chạy mô hình.",
        )

        sc1, sc2, sc3 = st.columns([1.2, 1.4, 1.4], vertical_alignment="bottom")
        with sc1:
            sample_n = st.selectbox("Kích thước mẫu ngẫu nhiên (KH)", [100, 200, 500, 1000], key="sample_n")
        sample_csv_str = load_random_sample_csv(sample_n, st.session_state["sample_seed"])
        with sc2:
            if st.button("Đổi mẫu ngẫu nhiên (Random)", width="stretch", key="btn_random_sample"):
                st.session_state["sample_seed"] = int(st.session_state["sample_seed"]) + 1
                st.session_state["batch_scored_raw"] = None
                st.session_state["batch_scored_probs"] = None
                st.session_state["batch_scored_sig"] = ""
                st.rerun()
        with sc3:
            st.download_button(
                f"Tải tệp CSV mẫu ({sample_n} KH)",
                data=sample_csv_str.encode("utf-8-sig"),
                file_name=f"sample_customers_{sample_n}_random.csv",
                mime="text/csv",
                width="stretch",
                key="btn_dl_sample_csv",
            )

        staged_df = None
        staged_sig = ""
        staged_source = ""
        if uploaded is not None:
            try:
                file_bytes = uploaded.getvalue()
                staged_df = pd.read_csv(io.BytesIO(file_bytes))
                staged_sig = f"upload:{uploaded.name}:{len(file_bytes)}:{selected_model}"
                staged_source = f"Tệp tải lên: **{uploaded.name}** ({len(staged_df):,} KH) • Mô hình: **{selected_model}**"
            except Exception as exc:
                st.error(f"Không thể đọc tệp CSV: {exc}")
        elif sample_csv_str:
            staged_df = pd.read_csv(io.StringIO(sample_csv_str))
            staged_sig = f"sample:{sample_n}:{st.session_state['sample_seed']}:{selected_model}"
            staged_source = (
                f"Mẫu ngẫu nhiên: **{sample_n} KH** (Seed #{st.session_state['sample_seed']}) • "
                f"Mô hình: **{selected_model}**"
            )

        if staged_sig and st.session_state["batch_scored_sig"] and st.session_state["batch_scored_sig"] != staged_sig:
            st.session_state["batch_scored_raw"] = None
            st.session_state["batch_scored_probs"] = None
            st.session_state["batch_scored_sig"] = ""

        if staged_df is not None:
            if st.session_state["batch_scored_probs"] is None:
                st.info(f"Đã sẵn sàng — {staged_source}. Bấm nút **\"Chấm điểm\"** để chạy mô hình.")
            else:
                st.success(f"Đã hoàn tất — {st.session_state['batch_scored_source']} | Ngưỡng: **T = {threshold:.2f}**")

        act1, act2, _ = st.columns([1.0, 1.2, 1.8], vertical_alignment="center")
        with act1:
            run_batch_btn = st.button("Chấm điểm", type="primary", width="stretch", key="btn_run_batch", disabled=(staged_df is None))
        with act2:
            reset_batch_btn = st.button("Xóa kết quả / Chọn lại", type="secondary", width="stretch", key="btn_reset_batch")

        if reset_batch_btn:
            st.session_state["batch_scored_raw"] = None
            st.session_state["batch_scored_probs"] = None
            st.session_state["batch_scored_sig"] = ""
            st.session_state["batch_scored_source"] = ""
            st.session_state["uploader_key"] += 1
            st.rerun()

        if run_batch_btn and staged_df is not None:
            try:
                show_center_loading(f"Đang chạy {selected_model} chấm điểm {len(staged_df):,} khách hàng...", 0.35)
                proc_batch = add_domain_features(staged_df)
                probs = active_pipe.predict_proba(proc_batch[feature_cols])[:, 1]
                st.session_state["batch_scored_raw"] = staged_df
                st.session_state["batch_scored_probs"] = probs
                st.session_state["batch_scored_sig"] = staged_sig
                st.session_state["batch_scored_source"] = staged_source
                st.rerun()
            except Exception as exc:
                st.error(f"Tệp CSV không đúng định dạng Telco Churn: {exc}")

        if st.session_state["batch_scored_probs"] is not None and st.session_state["batch_scored_raw"] is not None:
            raw_scored = st.session_state["batch_scored_raw"]
            probs = st.session_state["batch_scored_probs"]
            ids = raw_scored["customerID"] if "customerID" in raw_scored.columns else pd.Series(range(1, len(raw_scored) + 1))
            proc_scored = add_domain_features(raw_scored)

            tiers = np.where(
                probs >= 0.50,
                "Nguy cơ cao (>=50%)",
                np.where(probs >= threshold, f"Cảnh báo ({threshold*100:.0f}%–50%)", "Ổn định"),
            )

            out_df = pd.DataFrame({
                "customerID": ids,
                "Contract": raw_scored["Contract"],
                "InternetService": raw_scored["InternetService"],
                "tenure": proc_scored["tenure"],
                "MonthlyCharges": np.round(proc_scored["MonthlyCharges"], 2),
                "Xac_suat_Churn_Pct": np.round(probs * 100, 2),
                "Phan_tang_rui_ro": tiers,
            }).sort_values("Xac_suat_Churn_Pct", ascending=False)

            n_high = int((probs >= threshold).sum())
            n_very_high = int((probs >= 0.50).sum())
            est_campaign_cost = n_high * c_fp
            total_cnt = len(out_df) if len(out_df) > 0 else 1

            m1, m2, m3, m4 = st.columns(4)
            with m1:
                st.markdown(
                    f'<div class="kpi-card"><div class="kpi-label">Tổng số khách hàng</div>'
                    f'<div class="kpi-value">{len(out_df):,}</div>'
                    f'<div class="kpi-sub">Mô hình: {selected_model}</div></div>',
                    unsafe_allow_html=True,
                )
            with m2:
                st.markdown(
                    f'<div class="kpi-card amber"><div class="kpi-label">KH cần can thiệp (&gt;= {threshold:.2f})</div>'
                    f'<div class="kpi-value">{n_high:,}</div>'
                    f'<div class="kpi-sub">Chiếm {n_high/total_cnt*100:.1f}% tổng số KH</div></div>',
                    unsafe_allow_html=True,
                )
            with m3:
                st.markdown(
                    f'<div class="kpi-card indigo"><div class="kpi-label">KH nguy cơ cao (&gt;= 0.50)</div>'
                    f'<div class="kpi-value">{n_very_high:,}</div>'
                    f'<div class="kpi-sub">Chiếm {n_very_high/total_cnt*100:.1f}% tổng số KH</div></div>',
                    unsafe_allow_html=True,
                )
            with m4:
                st.markdown(
                    f'<div class="kpi-card teal"><div class="kpi-label">Ngân sách dự kiến</div>'
                    f'<div class="kpi-value">${est_campaign_cost:,.0f}</div>'
                    f'<div class="kpi-sub">Định mức {c_fp:,.0f} USD / KH can thiệp</div></div>',
                    unsafe_allow_html=True,
                )
            st.markdown("<div style='height: 10px;'></div>", unsafe_allow_html=True)

            with st.container(border=True):
                fig_batch = plot_batch_distribution(probs, threshold, is_dark)
                st.pyplot(fig_batch, width="stretch")
                plt.close(fig_batch)

            f_col1, f_col2 = st.columns(2)
            with f_col1:
                sel_tiers = st.multiselect(
                    "Lọc theo Phân tầng rủi ro",
                    options=list(out_df["Phan_tang_rui_ro"].unique()),
                    default=list(out_df["Phan_tang_rui_ro"].unique()),
                )
            with f_col2:
                sel_contracts = st.multiselect(
                    "Lọc theo Loại hợp đồng",
                    options=list(out_df["Contract"].unique()),
                    default=list(out_df["Contract"].unique()),
                )

            filtered_df = out_df[
                out_df["Phan_tang_rui_ro"].isin(sel_tiers)
                & out_df["Contract"].isin(sel_contracts)
            ]

            st.dataframe(
                filtered_df,
                width="stretch",
                height=380,
                hide_index=True,
                column_config={
                    "Xac_suat_Churn_Pct": st.column_config.ProgressColumn(
                        "Xác suất Churn (%)", format="%.2f%%", min_value=0, max_value=100
                    ),
                    "MonthlyCharges": st.column_config.NumberColumn("Cước tháng (USD)", format="$%.2f"),
                },
            )
            st.download_button(
                "Tải xuống kết quả dự đoán (CSV)",
                data=filtered_df.to_csv(index=False).encode("utf-8-sig"),
                file_name="churn_batch_predictions.csv",
                mime="text/csv",
            )

    def render_html_table(df_in, highlight_col: str = None, highlight_val: str = None):
        import html as _html
        ths = "".join(f"<th>{_html.escape(str(c))}</th>" for c in df_in.columns)
        trs = []
        for _, row in df_in.iterrows():
            cls = ' class="hl-row"' if (highlight_col and str(row[highlight_col]) == str(highlight_val)) else ""
            tds = "".join(f"<td>{_html.escape(str(row[c]))}</td>" for c in df_in.columns)
            trs.append(f"<tr{cls}>{tds}</tr>")
        st.markdown(
            f'<div class="tbl-wrap"><table class="tbl-clean"><thead><tr>{ths}</tr></thead>'
            f'<tbody>{"".join(trs)}</tbody></table></div>',
            unsafe_allow_html=True,
        )

    # ==================== TAB 3 ====================
    if active_tab == "tab3":
        st.markdown("#### Kết quả thực nghiệm & Tri thức khai phá dữ liệu")
        def _set_sub_tab(k):
            st.session_state["active_sub_tab"] = k
        _sub_labels = ["Thực nghiệm động (Live)", "Đối sánh 5 mô hình", "Ablation Study", "Khám phá (EDA)", "SHAP & Luật CART"]
        _sub_cols = st.columns([1.25, 1.05, 0.9, 0.9, 1.0], gap="small")
        for _i, _lbl in enumerate(_sub_labels):
            _k = "sub%d" % _i
            _sub_cols[_i].button(_lbl, key="nav_" + _k, use_container_width=True,
                type="primary" if st.session_state["active_sub_tab"] == _k else "secondary",
                on_click=_set_sub_tab, args=(_k,))
        active_sub = st.session_state["active_sub_tab"]

        if active_sub == "sub0":
            st.markdown(
                f"##### Đánh giá thời gian thực trên Tập Test độc lập (N = 1.409) — "
                f"Mô hình: **{selected_model}** | Ngưỡng: **T = {threshold:.2f}** | "
                f"Chi phí: **C_FN = {c_fn:,.0f} USD, C_FP = {c_fp:,.0f} USD**"
            )
            lk1, lk2, lk3, lk4 = st.columns(4)
            with lk1:
                st.markdown(
                    f'<div class="kpi-card"><div class="kpi-label">Accuracy / F1-Score</div>'
                    f'<div class="kpi-value">{live_test_m["accuracy"]*100:.2f}% / {live_test_m["f1"]:.4f}</div>'
                    f'<div class="kpi-sub">Ngưỡng phân loại T = {threshold:.2f}</div></div>',
                    unsafe_allow_html=True,
                )
            with lk2:
                st.markdown(
                    f'<div class="kpi-card indigo"><div class="kpi-label">Recall / Precision</div>'
                    f'<div class="kpi-value">{live_test_m["recall"]*100:.2f}% / {live_test_m["precision"]*100:.2f}%</div>'
                    f'<div class="kpi-sub">Giữ được {live_test_m["tp"]}/374 Churn • Bỏ sót {live_test_m["fn"]}</div></div>',
                    unsafe_allow_html=True,
                )
            with lk3:
                st.markdown(
                    f'<div class="kpi-card amber"><div class="kpi-label">Tổng chi phí tại T</div>'
                    f'<div class="kpi-value">${live_test_m["total_cost_usd"]:,.0f}</div>'
                    f'<div class="kpi-sub">Trung bình ${live_test_m["avg_cost_usd"]:.2f} / KH</div></div>',
                    unsafe_allow_html=True,
                )
            with lk4:
                st.markdown(
                    f'<div class="kpi-card teal"><div class="kpi-label">Tiết kiệm chi phí (ROI)</div>'
                    f'<div class="kpi-value">${live_test_m["savings_usd"]:,.0f}</div>'
                    f'<div class="kpi-sub">Giảm {live_test_m["savings_pct"]:.1f}% chi phí thiệt hại</div></div>',
                    unsafe_allow_html=True,
                )

            st.markdown("<div style='height: 12px;'></div>", unsafe_allow_html=True)
            lc1, lc2 = st.columns([1.35, 1.0])
            with lc1:
                with st.container(border=True):
                    fig_cost = plot_live_cost_curve(cost_curve, opt_thr, threshold, selected_model, is_dark)
                    st.pyplot(fig_cost, width="stretch")
                    plt.close(fig_cost)
            with lc2:
                with st.container(border=True):
                    fig_cm = plot_live_confusion_matrix(live_test_m, selected_model, threshold, c_fn, c_fp, is_dark)
                    st.pyplot(fig_cm, width="stretch")
                    plt.close(fig_cm)

            st.markdown(f"##### Bảng đối sánh động 5 mô hình tại T = {threshold:.2f} (C_FN = {c_fn:,.0f} USD, C_FP = {c_fp:,.0f} USD)")
            live_payload = _compute_live_eval_payload(
                y_train_arr, y_test_arr, oof_probs_all, test_probs_all,
                selected_model, threshold, c_fn, c_fp,
            )
            render_html_table(pd.DataFrame(live_payload["comp_live"]), highlight_col="Mô hình", highlight_val=selected_model)

            st.markdown("##### Đối chứng 3 chiến lược vận hành động")
            render_html_table(pd.DataFrame(live_payload["roi_rows"]))

        elif active_sub == "sub1":
            st.markdown("##### Bảng đối sánh cố định 5 mô hình phân loại trên tập Test độc lập (N = 1.409)")
            render_html_table(comp_df, highlight_col="Mô hình", highlight_val=selected_model)

            g1, g2 = st.columns(2)
            with g1:
                p_fig4 = os.path.join(FIG_DIR, "fig04_model_comparison.png")
                if os.path.exists(p_fig4):
                    st.image(p_fig4, caption="Hình 4: Đối sánh ROC-AUC, PR-AUC và F1-Score của 5 mô hình", width="stretch")
            with g2:
                p_fig5 = os.path.join(FIG_DIR, "fig05_roc_pr_curves.png")
                if os.path.exists(p_fig5):
                    st.image(p_fig5, caption="Hình 5: Đường cong ROC-AUC và Precision-Recall (PR-AUC)", width="stretch")

            p_fig6 = os.path.join(FIG_DIR, "fig06_confusion_matrices.png")
            if os.path.exists(p_fig6):
                st.image(p_fig6, caption="Hình 6: Ma trận nhầm lẫn của LightGBM tại T = 0.50 và T* = 0.21", width="stretch")

        elif active_sub == "sub2":
            st.markdown("##### Nghiên cứu cắt bỏ (Ablation Study) xử lý mất cân bằng trên LightGBM")
            ab_rows = []
            for strat_name, m in results["ablation_study"].items():
                ab_rows.append({
                    "Phương án": strat_name,
                    "Ngưỡng T": m["threshold"],
                    "Accuracy": m["accuracy"],
                    "Precision": m["precision"],
                    "Recall": m["recall"],
                    "F1-Score": m["f1"],
                    "ROC-AUC": m["roc_auc"],
                    "Brier Score": m["brier"],
                    "Tổng chi phí (USD)": m["total_cost_usd"],
                })
            render_html_table(pd.DataFrame(ab_rows))

            p_fig7 = os.path.join(FIG_DIR, "fig07_cost_threshold_roi.png")
            if os.path.exists(p_fig7):
                st.image(p_fig7, caption="Hình 7: Tối ưu ngưỡng chi phí T* = 0.21 và đối chứng 4 chiến lược vận hành", width="stretch")

        elif active_sub == "sub3":
            e1, e2 = st.columns(2)
            with e1:
                p_fig1 = os.path.join(FIG_DIR, "fig01_churn_distribution.png")
                if os.path.exists(p_fig1):
                    st.image(p_fig1, caption="Hình 1: Phân phối biến mục tiêu Churn (73.46% ở lại vs 26.54% rời bỏ)", width="stretch")
            with e2:
                p_fig2 = os.path.join(FIG_DIR, "fig02_eda_drivers.png")
                if os.path.exists(p_fig2):
                    st.image(p_fig2, caption="Hình 2: Các yếu tố cốt lõi dẫn đến hành vi rời bỏ dịch vụ", width="stretch")

            p_fig3 = os.path.join(FIG_DIR, "fig03_correlation_cramersv.png")
            if os.path.exists(p_fig3):
                st.image(p_fig3, caption="Hình 3: Ma trận tương quan và mức độ liên hệ Cramér's V", width="stretch")

        elif active_sub == "sub4":
            p_fig8 = os.path.join(FIG_DIR, "fig08_decision_tree_rules.png")
            if os.path.exists(p_fig8):
                st.image(p_fig8, caption="Hình 8: Cây quyết định CART cắt tỉa độ sâu 3 phục vụ rút trích luật IF-THEN", width="stretch")

            s1, s2 = st.columns(2)
            with s1:
                p_fig9 = os.path.join(FIG_DIR, "fig09_shap_summary.png")
                if os.path.exists(p_fig9):
                    st.image(p_fig9, caption="Hình 9: Giải thích toàn cục SHAP Summary Plot (Top 12 đặc trưng)", width="stretch")
            with s2:
                p_fig10 = os.path.join(FIG_DIR, "fig10_shap_waterfall_cases.png")
                if os.path.exists(p_fig10):
                    st.image(p_fig10, caption="Hình 10: Phân tích SHAP Waterfall cho 2 hồ sơ điển hình", width="stretch")

            st.markdown("##### 8 Luật phân loại IF-THEN rút trích từ cây quyết định CART")
            rules_df = pd.DataFrame(results["extracted_cart_rules"]).rename(columns={
                "conditions": "Điều kiện luật (IF)",
                "n_samples": "Số mẫu",
                "support_pct": "Độ hỗ trợ (%)",
                "churn_prob_pct": "Xác suất Churn (%)",
                "predicted_class": "Kết luận (THEN)",
            })
            render_html_table(rules_df)

            st.markdown("##### Phân tích lỗi dự báo (Error Analysis: T = 0.50 vs T* = 0.21)")
            err_rows = []
            for label, key in [("Ngưỡng mặc định T = 0.50", "at_default_050"), ("Ngưỡng tối ưu T* = 0.21", "at_optimal_tstar")]:
                e = results["error_analysis"][key]
                err_rows.append({
                    "Cấu hình": label,
                    "Tổng mẫu Test": e["total_test"],
                    "Tổng số lỗi": e["total_errors"],
                    "Tỷ lệ lỗi (%)": e["error_rate_pct"],
                    "Bỏ sót Churn (FN)": e["fn_count"],
                    "Cảnh báo nhầm (FP)": e["fp_count"],
                    "Lỗi vùng ranh giới": e["boundary_errors"],
                    "FN lâu năm âm thầm": e["fn_loyal_silent"],
                    "FN ngắn hạn biến động": e["fn_short_volatile"],
                })
            render_html_table(pd.DataFrame(err_rows))


if __name__ == "__main__" and _is_running_in_streamlit():
    run_streamlit_app()
