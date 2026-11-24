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
from urllib.parse import unquote

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
if BASE_DIR not in sys.path:
    sys.path.insert(0, BASE_DIR)

from src.portable_engine import (
    FEATURE_VI_MAP,
    PRESET_PROFILES,
    batch_predict_csv,
    build_recommendations,
    load_portable_model,
    predict_and_shap,
)

MODEL_PATH = os.path.join(BASE_DIR, "models", "best_churn_pipeline.joblib")
PORTABLE_MODEL_PATH = os.path.join(BASE_DIR, "models", "portable_lgbm_model.json")
RESULTS_PATH = os.path.join(BASE_DIR, "models", "experiment_results.json")
COMP_CSV_PATH = os.path.join(BASE_DIR, "models", "model_comparison.csv")
SAMPLE_CSV_PATH = os.path.join(BASE_DIR, "data", "sample_test_customers.csv")
FIG_DIR = os.path.join(BASE_DIR, "reports", "figures")

_WSGI_CACHE = {}


def _get_wsgi_data():
    """Tải và cache dữ liệu tĩnh + portable model cho WSGI server."""
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


def _build_wsgi_html(data: dict) -> str:
    """Sinh trang HTML/CSS/JS độc lập cho Vercel WSGI với chế độ Sáng/Tối và giao diện Basic."""
    results = data["results"]
    comp_rows = data["comp_rows"]
    model_meta = {
        "model_name": data["model"]["model_name"],
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
      --muted-2: #64748B;
      --accent: #1E40AF;
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
      --grid-line: #E2E8F0;
    }}
    [data-theme="dark"] {{
      --bg: #0B1120;
      --surface: #111827;
      --surface-alt: #1E293B;
      --border: #334155;
      --border-soft: #1F2937;
      --text: #F1F5F9;
      --muted: #94A3B8;
      --muted-2: #94A3B8;
      --accent: #3B82F6;
      --accent-soft: #172554;
      --hero-bg: #0F172A;
      --hero-text: #F8FAFC;
      --hero-sub: #94A3B8;
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
      --grid-line: #334155;
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
    .header h1 {{
      margin: 0;
      font-size: 21px;
      font-weight: 700;
    }}
    .header p {{
      margin: 4px 0 0;
      color: var(--hero-sub);
      font-size: 13px;
    }}
    .theme-toggle {{
      background: rgba(255,255,255,0.12);
      color: #FFFFFF;
      border: 1px solid rgba(255,255,255,0.28);
      border-radius: 6px;
      padding: 8px 14px;
      font-size: 13px;
      font-weight: 600;
      cursor: pointer;
    }}
    .theme-toggle:hover {{ background: rgba(255,255,255,0.22); }}
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
    }}
    .kpi-label {{
      font-size: 11.5px;
      font-weight: 600;
      text-transform: uppercase;
      letter-spacing: 0.4px;
      color: var(--muted);
    }}
    .kpi-val {{
      font-size: 21px;
      font-weight: 700;
      margin: 4px 0;
    }}
    .kpi-sub {{
      font-size: 12px;
      color: var(--muted);
    }}
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
    .tab-nav {{
      display: flex;
      gap: 8px;
      background: var(--surface);
      border: 1px solid var(--border-soft);
      border-radius: 8px;
      padding: 6px;
      margin-bottom: 16px;
      flex-wrap: wrap;
    }}
    .tab-btn {{
      border: 1px solid transparent;
      background: transparent;
      color: var(--muted);
      padding: 9px 16px;
      border-radius: 6px;
      font-weight: 600;
      font-size: 13.5px;
      cursor: pointer;
    }}
    .tab-btn.active {{
      background: var(--accent-soft);
      color: var(--accent);
      border-color: var(--border);
    }}
    .tab-pane {{ display: none; }}
    .tab-pane.active {{ display: block; }}
    .btn-row {{
      display: flex;
      gap: 10px;
      flex-wrap: wrap;
      margin-bottom: 16px;
    }}
    .btn {{
      background: var(--surface);
      color: var(--text);
      border: 1px solid var(--border);
      border-radius: 6px;
      padding: 8px 14px;
      font-size: 13px;
      font-weight: 600;
      cursor: pointer;
    }}
    .btn:hover {{ border-color: var(--accent); }}
    .btn-primary {{
      background: var(--accent);
      color: #FFFFFF;
      border-color: var(--accent);
    }}
    .form-grid {{
      display: grid;
      grid-template-columns: repeat(3, minmax(0, 1fr));
      gap: 14px;
      margin-bottom: 16px;
    }}
    .panel {{
      background: var(--surface);
      border: 1px solid var(--border-soft);
      border-radius: 8px;
      padding: 16px;
    }}
    .panel h3 {{
      margin: 0 0 12px;
      font-size: 14.5px;
      font-weight: 700;
      border-bottom: 1px solid var(--border-soft);
      padding-bottom: 8px;
    }}
    .field {{ margin-bottom: 10px; }}
    .field label {{
      display: block;
      font-size: 12px;
      font-weight: 600;
      color: var(--muted);
      margin-bottom: 4px;
    }}
    .field select, .field input {{
      width: 100%;
      padding: 7px 10px;
      border-radius: 6px;
      border: 1px solid var(--border);
      background: var(--surface-alt);
      color: var(--text);
      font-size: 13px;
    }}
    .chips-grid {{
      display: grid;
      grid-template-columns: repeat(6, minmax(0, 1fr));
      gap: 10px;
      margin-bottom: 16px;
    }}
    .chip {{
      background: var(--surface);
      border: 1px solid var(--border-soft);
      border-radius: 6px;
      padding: 10px;
      text-align: center;
    }}
    .chip-name {{ font-size: 11px; color: var(--muted); font-weight: 600; }}
    .chip-val {{ font-size: 14.5px; font-weight: 700; margin-top: 3px; }}
    .result-grid {{
      display: grid;
      grid-template-columns: 1fr 1.35fr;
      gap: 16px;
    }}
    .gauge-track {{
      position: relative;
      height: 14px;
      border-radius: 999px;
      background: linear-gradient(90deg, #10B981 0%, #F59E0B 45%, #EF4444 100%);
      margin: 12px 0 8px;
    }}
    .gauge-thr {{
      position: absolute; top: -4px; bottom: -4px; width: 3px; background: var(--text);
    }}
    .gauge-mark {{
      position: absolute; top: -6px; width: 12px; height: 26px;
      transform: translateX(-50%); background: var(--surface);
      border: 3px solid var(--accent); border-radius: 4px;
    }}
    .status-box {{
      border-radius: 6px;
      padding: 12px 14px;
      border: 1px solid transparent;
      border-left-width: 4px;
      margin: 12px 0;
    }}
    .status-high {{ background: var(--danger-bg); border-color: var(--danger-border); border-left-color: var(--danger); color: var(--danger-text); }}
    .status-mid {{ background: var(--warn-bg); border-color: var(--warn-border); border-left-color: var(--warn); color: var(--warn-text); }}
    .status-low {{ background: var(--ok-bg); border-color: var(--ok-border); border-left-color: var(--ok); color: var(--ok-text); }}
    .rec-item {{
      background: var(--surface-alt);
      border-left: 3px solid var(--accent);
      border-radius: 6px;
      padding: 9px 12px;
      margin-bottom: 8px;
      font-size: 13px;
    }}
    table {{
      width: 100%;
      border-collapse: collapse;
      font-size: 12.5px;
    }}
    th, td {{
      text-align: left;
      padding: 8px 10px;
      border-bottom: 1px solid var(--border-soft);
    }}
    th {{
      background: var(--surface-alt);
      font-weight: 700;
      color: var(--muted);
      position: sticky;
      top: 0;
    }}
    .table-wrap {{
      max-height: 420px;
      overflow: auto;
      border: 1px solid var(--border-soft);
      border-radius: 8px;
      background: var(--surface);
    }}
    .prog-bar {{
      width: 90px; height: 8px; background: var(--surface-alt); border-radius: 999px;
      display: inline-block; overflow: hidden; vertical-align: middle; margin-right: 6px;
    }}
    .prog-fill {{ height: 100%; background: var(--accent); }}
    .fig-grid {{
      display: grid;
      grid-template-columns: repeat(2, minmax(0, 1fr));
      gap: 14px;
      margin-top: 14px;
    }}
    .fig-card {{
      background: #FFFFFF;
      border: 1px solid var(--border-soft);
      border-radius: 8px;
      padding: 10px;
      color: #0F172A;
    }}
    .fig-card img {{ width: 100%; height: auto; display: block; border-radius: 4px; }}
    .fig-cap {{ font-size: 12px; color: #475569; margin-top: 6px; text-align: center; font-weight: 600; }}
    @media (max-width: 980px) {{
      .kpi-grid, .form-grid, .result-grid, .fig-grid {{ grid-template-columns: 1fr; }}
      .chips-grid {{ grid-template-columns: repeat(2, minmax(0, 1fr)); }}
    }}
  </style>
</head>
<body>
  <div class="container">
    <header class="header">
      <div>
        <h1>Hệ thống Dự đoán và Phân tích Khách hàng Rời bỏ Dịch vụ (Customer Churn Analytics)</h1>
        <p>Mô hình phân loại nhạy chi phí (Cost-Sensitive LightGBM) và giải thích cục bộ TreeSHAP</p>
      </div>
      <button class="theme-toggle" id="btn-theme" onclick="toggleTheme()">Giao diện: Tối</button>
    </header>

    <section class="kpi-grid" id="kpi-grid"></section>

    <section class="control-bar">
      <div style="display:flex; align-items:center; gap:12px; flex-wrap:wrap;">
        <strong>Ngưỡng quyết định (T):</strong>
        <input type="range" id="thr-slider" min="0.10" max="0.80" step="0.01" value="0.21" style="width:200px;" />
        <span id="thr-val" style="font-weight:700; min-width:44px;">0.21</span>
        <button class="btn btn-primary" onclick="setThr(0.21)">Đặt lại T* = 0.21</button>
        <button class="btn" onclick="setThr(0.50)">Ngưỡng mặc định T = 0.50</button>
      </div>
      <div style="font-size:12.5px; color:var(--muted);">
        Ma trận chi phí: C_FN = 500 USD | C_FP = 100 USD | Tỷ lệ 5:1
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

      <div class="result-grid">
        <div class="panel">
          <div style="display:flex; justify-content:space-between; align-items:baseline;">
            <span style="font-weight:600; color:var(--muted);">Xác suất rời bỏ dịch vụ (LightGBM)</span>
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
          <h3>Giải thích SHAP cục bộ — Top 8 yếu tố ảnh hưởng</h3>
          <div id="shap-chart"></div>
        </div>
      </div>
    </div>

    <!-- TAB 2 -->
    <div id="tab2" class="tab-pane">
      <div class="btn-row">
        <a class="btn" href="/api/sample-csv" download="sample_test_customers.csv">Tải tệp CSV mẫu (100 KH tập Test)</a>
        <label class="btn" style="display:inline-block;">
          Tải lên tệp CSV khác
          <input type="file" id="csv-upload" accept=".csv" style="display:none;" onchange="handleCsvUpload(this)" />
        </label>
        <button class="btn btn-primary" onclick="downloadBatchResult()">Tải xuống kết quả dự đoán (CSV)</button>
      </div>
      <div class="kpi-grid" id="batch-kpis"></div>
      <div class="btn-row" style="align-items:center;">
        <label>Lọc phân tầng:
          <select id="filter-tier" onchange="renderBatchTable()" style="padding:6px 10px; border-radius:6px; border:1px solid var(--border); background:var(--surface); color:var(--text);">
            <option value="ALL">Tất cả phân tầng</option>
            <option value="high">Nguy cơ cao (&gt;= 50%)</option>
            <option value="mid">Cảnh báo (T – 50%)</option>
            <option value="low">Ổn định</option>
          </select>
        </label>
        <label>Lọc hợp đồng:
          <select id="filter-contract" onchange="renderBatchTable()" style="padding:6px 10px; border-radius:6px; border:1px solid var(--border); background:var(--surface); color:var(--text);">
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
            <tr>
              <th>Mã KH</th><th>Hợp đồng</th><th>Internet</th><th>Số tháng</th><th>Cước/tháng</th><th>Xác suất Churn</th><th>Phân tầng rủi ro</th>
            </tr>
          </thead>
          <tbody id="batch-tbody"></tbody>
        </table>
      </div>
    </div>

    <!-- TAB 3 -->
    <div id="tab3" class="tab-pane">
      <div class="tab-nav" style="margin-bottom:12px;">
        <button class="tab-btn active" onclick="switchSub('sub1', this)">Đối sánh 5 mô hình</button>
        <button class="tab-btn" onclick="switchSub('sub2', this)">Tối ưu chi phí &amp; ROI</button>
        <button class="tab-btn" onclick="switchSub('sub3', this)">Khám phá dữ liệu (EDA)</button>
        <button class="tab-btn" onclick="switchSub('sub4', this)">Giải thích SHAP &amp; Luật CART</button>
      </div>
      <div id="sub1" class="sub-pane"></div>
      <div id="sub2" class="sub-pane" style="display:none;"></div>
      <div id="sub3" class="sub-pane" style="display:none;"></div>
      <div id="sub4" class="sub-pane" style="display:none;"></div>
    </div>
  </div>

  <script>
    const BOOT = {bootstrap_json};
    let currentBatch = [];
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

    function renderKpis() {{
      const r = BOOT.results;
      const best = BOOT.model.model_name;
      const m05 = r.test_metrics_at_050[best];
      const mt = r.test_metrics_at_tstar[best];
      const roi = r.roi_benchmark[r.roi_benchmark.length - 1];
      document.getElementById("kpi-grid").innerHTML = `
        <div class="kpi-card"><div class="kpi-label">Quy mô dữ liệu Telco</div><div class="kpi-val">${{r.dataset_audit.n_rows.toLocaleString()}} thuê bao</div><div class="kpi-sub">${{r.dataset_audit.n_raw_cols}} cột gốc → ${{r.dataset_audit.n_features_engineered}} đặc trưng • Churn ${{(r.dataset_audit.churn_rate*100).toFixed(2)}}%</div></div>
        <div class="kpi-card"><div class="kpi-label">Mô hình tối ưu</div><div class="kpi-val">${{best}}</div><div class="kpi-sub">5-Fold CV AUC: ${{r.cv_roc_auc[best].toFixed(4)}} • PR-AUC: ${{m05.pr_auc.toFixed(4)}}</div></div>
        <div class="kpi-card"><div class="kpi-label">Ngưỡng nhạy chi phí</div><div class="kpi-val">T* = ${{BOOT.model.optimal_threshold.toFixed(2)}}</div><div class="kpi-sub">Recall tăng ${{(m05.recall*100).toFixed(1)}}% → ${{(mt.recall*100).toFixed(2)}}%</div></div>
        <div class="kpi-card"><div class="kpi-label">Hiệu quả kinh tế (ROI)</div><div class="kpi-val">Tiết kiệm $${{roi.savings_vs_nothing_usd.toLocaleString()}}</div><div class="kpi-sub">Giảm ${{roi.savings_pct.toFixed(1)}}% chi phí so với không can thiệp</div></div>
      `;
    }}

    function setThr(val) {{
      document.getElementById("thr-slider").value = val.toFixed(2);
      document.getElementById("thr-val").textContent = val.toFixed(2);
      runPredict();
      loadBatch(null);
    }}
    document.getElementById("thr-slider").addEventListener("input", (e) => {{
      const v = parseFloat(e.target.value);
      document.getElementById("thr-val").textContent = v.toFixed(2);
      runPredict();
    }});
    document.getElementById("thr-slider").addEventListener("change", () => loadBatch(null));

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
      const payload = {{ threshold: parseFloat(document.getElementById("thr-slider").value) }};
      FIELD_IDS.forEach(k => {{
        const el = document.getElementById("f-" + k);
        payload[k] = el.value;
      }});
      const res = await fetch("/api/predict", {{
        method: "POST",
        headers: {{ "Content-Type": "application/json" }},
        body: JSON.stringify(payload)
      }});
      const out = await res.json();
      const pct = out.probability * 100;
      const thr = payload.threshold;
      const thrPct = thr * 100;

      document.getElementById("prob-big").textContent = pct.toFixed(2) + "%";
      document.getElementById("gauge-thr").style.left = thrPct.toFixed(1) + "%";
      document.getElementById("gauge-mark").style.left = Math.min(Math.max(pct, 1), 99).toFixed(1) + "%";
      document.getElementById("gauge-thr-txt").textContent = "Ngưỡng T = " + thrPct.toFixed(0) + "%";

      const sb = document.getElementById("status-box");
      if (out.probability >= 0.50) {{
        sb.className = "status-box status-high";
        sb.innerHTML = `<strong>CẢNH BÁO NGUY CƠ RỜI BỎ RẤT CAO</strong><br>Xác suất <b>${{pct.toFixed(2)}}%</b> vượt ngưỡng T = ${{thr.toFixed(2)}} và vượt mốc 50%. Cần can thiệp ngay.`;
      }} else if (out.probability >= thr) {{
        sb.className = "status-box status-mid";
        sb.innerHTML = `<strong>CẢNH BÁO RỦI RO TRUNG GIAN (BẮT NHỜ NGƯỠNG T*)</strong><br>Xác suất <b>${{pct.toFixed(2)}}%</b> vượt ngưỡng nhạy chi phí T = ${{thr.toFixed(2)}}.`;
      }} else {{
        sb.className = "status-box status-low";
        sb.innerHTML = `<strong>NHÓM KHÁCH HÀNG ỔN ĐỊNH</strong><br>Xác suất <b>${{pct.toFixed(2)}}%</b> thấp hơn ngưỡng cảnh báo T = ${{thr.toFixed(2)}}.`;
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
      document.getElementById("domain-chips").innerHTML = chips.map(c =>
        `<div class="chip"><div class="chip-name">${{c[0]}}</div><div class="chip-val">${{c[1]}}</div></div>`
      ).join("");

      document.getElementById("rec-list").innerHTML = out.recommendations.map(r => `<div class="rec-item">${{r}}</div>`).join("");

      // Render crisp horizontal SVG SHAP bar chart
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

    let lastCsvText = null;
    async function loadBatch(csvText) {{
      if (csvText !== null) lastCsvText = csvText;
      const thr = parseFloat(document.getElementById("thr-slider").value);
      const res = await fetch("/api/batch", {{
        method: "POST",
        headers: {{ "Content-Type": "application/json" }},
        body: JSON.stringify({{ threshold: thr, csv_text: lastCsvText }})
      }});
      const data = await res.json();
      currentBatch = data.rows || [];
      document.getElementById("batch-kpis").innerHTML = `
        <div class="kpi-card"><div class="kpi-label">Tổng số khách hàng</div><div class="kpi-val">${{data.total}}</div></div>
        <div class="kpi-card"><div class="kpi-label">Cần can thiệp (&gt;= ${{thr.toFixed(2)}})</div><div class="kpi-val">${{data.n_intervene}}</div><div class="kpi-sub">${{data.intervene_pct}}% danh sách</div></div>
        <div class="kpi-card"><div class="kpi-label">Nguy cơ cao (&gt;= 0.50)</div><div class="kpi-val">${{data.n_high_50}}</div></div>
        <div class="kpi-card"><div class="kpi-label">Ngân sách ưu đãi dự kiến</div><div class="kpi-val">$${{data.budget_usd.toLocaleString()}}</div><div class="kpi-sub">100 USD / KH can thiệp</div></div>
      `;
      renderBatchTable();
    }}

    function renderBatchTable() {{
      const ft = document.getElementById("filter-tier").value;
      const fc = document.getElementById("filter-contract").value;
      const filtered = currentBatch.filter(r => (ft === "ALL" || r.TierCode === ft) && (fc === "ALL" || r.Contract === fc));
      document.getElementById("batch-tbody").innerHTML = filtered.map(r => `
        <tr>
          <td>${{r.customerID}}</td>
          <td>${{r.Contract}}</td>
          <td>${{r.InternetService}}</td>
          <td>${{r.tenure}}</td>
          <td>$${{r.MonthlyCharges.toFixed(2)}}</td>
          <td><span class="prog-bar"><span class="prog-fill" style="width:${{r.Churn_Probability_Pct}}%;"></span></span>${{r.Churn_Probability_Pct.toFixed(2)}}%</td>
          <td>${{r.Tier}}</td>
        </tr>
      `).join("");
    }}

    function handleCsvUpload(input) {{
      const file = input.files[0];
      if (!file) return;
      const reader = new FileReader();
      reader.onload = () => loadBatch(reader.result);
      reader.readAsText(file, "utf-8");
    }}

    function downloadBatchResult() {{
      if (!currentBatch.length) return;
      const hdr = ["customerID","Contract","InternetService","tenure","MonthlyCharges","Churn_Probability_Pct","Tier"];
      const lines = [hdr.join(",")].concat(currentBatch.map(r => hdr.map(k => `"${{r[k]}}"`).join(",")));
      const blob = new Blob(["\\uFEFF" + lines.join("\\n")], {{ type: "text/csv;charset=utf-8;" }});
      const a = document.createElement("a");
      a.href = URL.createObjectURL(blob);
      a.download = "churn_batch_predictions.csv";
      a.click();
    }}

    function makeTable(headers, rows) {{
      return `<div class="table-wrap" style="max-height:none; margin-bottom:14px;"><table><thead><tr>${{headers.map(h => `<th>${{h}}</th>`).join("")}}</tr></thead><tbody>${{
        rows.map(r => `<tr>${{r.map(c => `<td>${{c}}</td>`).join("")}}</tr>`).join("")
      }}</tbody></table></div>`;
    }}

    function renderExperiments() {{
      const r = BOOT.results;
      const compCols = Object.keys(BOOT.comp_rows[0] || {{}});
      document.getElementById("sub1").innerHTML =
        `<h4>Bảng đối sánh 5 mô hình phân loại trên tập Test độc lập (N = 1.409)</h4>` +
        makeTable(compCols, BOOT.comp_rows.map(row => compCols.map(c => row[c]))) +
        `<div class="fig-grid">
          <div class="fig-card"><img src="/figures/fig04_model_comparison.png" /><div class="fig-cap">Hình 4: Đối sánh ROC-AUC, PR-AUC và F1-Score</div></div>
          <div class="fig-card"><img src="/figures/fig05_roc_pr_curves.png" /><div class="fig-cap">Hình 5: Đường cong ROC-AUC và PR-AUC</div></div>
        </div>
        <div class="fig-card" style="margin-top:14px;"><img src="/figures/fig06_confusion_matrices.png" /><div class="fig-cap">Hình 6: Ma trận nhầm lẫn LightGBM tại T = 0.50 và T* = 0.21</div></div>`;

      const roiRows = r.roi_benchmark.map(x => [x.strategy, x.targeted_customers, x.churn_caught, x.churn_missed, "$" + x.total_cost_usd.toLocaleString(), "$" + x.savings_vs_nothing_usd.toLocaleString(), x.savings_pct + "%"]);
      const abRows = Object.entries(r.ablation_study).map(([k, m]) => [k, m.threshold, m.accuracy, m.precision, m.recall, m.f1, m.roc_auc, m.brier, "$" + m.total_cost_usd.toLocaleString()]);
      document.getElementById("sub2").innerHTML =
        `<h4>Đối chứng chi phí nghiệp vụ theo 4 chiến lược (C_FN = 500 USD, C_FP = 100 USD)</h4>` +
        makeTable(["Chiến lược","Số KH can thiệp","Churn giữ được","Churn bỏ sót","Tổng chi phí","Tiết kiệm","Tỷ lệ tiết kiệm"], roiRows) +
        `<h4>Nghiên cứu cắt bỏ (Ablation Study) trên LightGBM</h4>` +
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

    renderKpis();
    applyPreset("high");
    loadBatch(null);
    renderExperiments();
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
        body = json.dumps({"status": "ok", "model": data["model"]["model_name"]}).encode("utf-8")
        return _wsgi_response(start_response, "200 OK", "application/json; charset=utf-8", body)

    if method == "GET" and path == "/api/sample-csv":
        return _wsgi_response(start_response, "200 OK", "text/csv; charset=utf-8", data["sample_csv"].encode("utf-8-sig"))

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
            res = predict_and_shap(payload, data["model"], top_k=8)
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
            csv_text = payload.get("csv_text") or data["sample_csv"]
            res = batch_predict_csv(csv_text, threshold, data["model"])
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
            "border": "#334155",
            "text": "#F1F5F9",
            "muted": "#94A3B8",
            "accent": "#3B82F6",
            "accent_soft": "#172554",
            "hero_bg": "#0F172A",
            "hero_sub": "#94A3B8",
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
            "border": "#E2E8F0",
            "text": "#0F172A",
            "muted": "#475569",
            "accent": "#1E40AF",
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
    .stApp h1, .stApp h2, .stApp h3, .stApp h4, .stApp h5, .stApp p, .stApp label,
    .stApp span, .stApp div {{
        color: inherit;
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
    div[data-baseweb="select"] > div,
    div[data-baseweb="input"] > div,
    .stNumberInput input,
    .stTextInput input {{
        background-color: {t['surface']} !important;
        color: {t['text']} !important;
        border-color: {t['border']} !important;
    }}
    div[data-testid="stVerticalBlockBorderWrapper"] {{
        background-color: {t['surface']};
        border-color: {t['border']} !important;
        border-radius: 8px;
    }}
    div[data-testid="stMetric"] {{
        background-color: {t['surface']};
        border: 1px solid {t['border']};
        border-radius: 8px;
        padding: 12px 14px;
    }}
    div[data-testid="stMetric"] label, div[data-testid="stMetricValue"] {{
        color: {t['text']} !important;
    }}
    .stButton > button {{
        background-color: {t['surface']};
        color: {t['text']};
        border: 1px solid {t['border']};
        border-radius: 6px;
        font-weight: 600;
    }}
    .stButton > button[kind="primary"] {{
        background-color: {t['accent']};
        color: #FFFFFF;
        border-color: {t['accent']};
    }}
    .hero-banner {{
        background: {t['hero_bg']};
        padding: 20px 26px;
        border-radius: 8px;
        border-bottom: 3px solid #D97706;
        margin-bottom: 16px;
    }}
    .hero-banner h1 {{
        color: #FFFFFF !important;
        margin: 0;
        font-size: 21px;
        font-weight: 700;
        line-height: 1.3;
    }}
    .hero-banner p {{
        color: {t['hero_sub']} !important;
        margin: 5px 0 0 0;
        font-size: 13px;
    }}
    .kpi-card {{
        background: {t['surface']};
        border: 1px solid {t['border']};
        border-top: 3px solid {t['accent']};
        border-radius: 8px;
        padding: 14px 16px;
        height: 100%;
    }}
    .kpi-card.amber {{ border-top-color: #D97706; }}
    .kpi-card.teal {{ border-top-color: #059669; }}
    .kpi-card.indigo {{ border-top-color: #4338CA; }}
    .kpi-label {{
        font-size: 11.5px;
        font-weight: 600;
        color: {t['muted']} !important;
        text-transform: uppercase;
        letter-spacing: 0.4px;
        margin-bottom: 4px;
    }}
    .kpi-value {{
        font-size: 21px;
        font-weight: 700;
        color: {t['text']} !important;
        line-height: 1.2;
    }}
    .kpi-sub {{
        font-size: 12px;
        color: {t['muted']} !important;
        margin-top: 5px;
    }}
    .stTabs [data-baseweb="tab-list"] {{
        gap: 6px;
        background-color: {t['surface']};
        padding: 6px;
        border-radius: 8px;
        border: 1px solid {t['border']};
        margin-bottom: 10px;
    }}
    .stTabs [data-baseweb="tab"] {{
        height: 40px;
        padding: 0 16px;
        border-radius: 6px;
        font-weight: 600;
        font-size: 13.5px;
        color: {t['muted']};
    }}
    .stTabs [aria-selected="true"] {{
        background-color: {t['accent_soft']} !important;
        color: {t['accent']} !important;
        border: 1px solid {t['border']};
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
        position: absolute;
        top: -5px;
        bottom: -5px;
        width: 3px;
        background: {t['text']};
        border-radius: 2px;
    }}
    .gauge-marker {{
        position: absolute;
        top: -6px;
        width: 12px;
        height: 26px;
        transform: translateX(-50%);
        background: {t['surface']};
        border: 3px solid {t['accent']};
        border-radius: 4px;
    }}
    .feat-chip {{
        background: {t['surface']};
        border: 1px solid {t['border']};
        border-radius: 6px;
        padding: 10px 8px;
        text-align: center;
        min-height: 64px;
        display: flex;
        flex-direction: column;
        justify-content: center;
    }}
    .feat-chip-name {{
        font-size: 11px;
        color: {t['muted']} !important;
        font-weight: 600;
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
        border-left: 3px solid {t['accent']};
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
    </style>
    """


def run_streamlit_app():
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

    @st.cache_resource
    def load_artifacts():
        bundle = joblib.load(MODEL_PATH)
        with open(RESULTS_PATH, "r", encoding="utf-8") as f:
            results = json.load(f)
        comp_df = pd.read_csv(COMP_CSV_PATH)
        explainer = shap.TreeExplainer(bundle["best_tree_pipeline"].named_steps["clf"])
        return bundle, results, comp_df, explainer

    def plot_local_shap(shap_vals_row, feature_names, is_dark: bool):
        bg = "#111827" if is_dark else "#FFFFFF"
        fg = "#F1F5F9" if is_dark else "#0F2544"
        sub = "#94A3B8" if is_dark else "#475569"
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

        ax.set_title(
            "Giải thích SHAP cục bộ — Top 8 yếu tố ảnh hưởng",
            fontsize=10.5,
            fontweight="bold",
            color=fg,
            pad=10,
        )
        ax.set_xlabel("Đóng góp SHAP (Đỏ: tăng rủi ro | Xanh: giảm rủi ro)", fontsize=8.8, color=sub)
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
        fg = "#F1F5F9" if is_dark else "#0F2544"
        sub = "#94A3B8" if is_dark else "#475569"
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

    if not os.path.exists(MODEL_PATH):
        st.error("Chưa tìm thấy mô hình đã huấn luyện. Vui lòng chạy `python src/train_pipeline.py` trước.")
        return

    bundle, results, comp_df, explainer = load_artifacts()
    best_pipe = bundle["best_model_pipeline"]
    tree_pipe = bundle["best_tree_pipeline"]
    best_name = bundle["best_model_name"]
    opt_thr = float(bundle["optimal_threshold"])
    tree_feat_names = bundle["tree_feature_names"]
    feature_cols = NUMERIC_FEATURES + BINARY_NUMERIC_FEATURES + CATEGORICAL_FEATURES

    if "threshold_slider" not in st.session_state:
        st.session_state["threshold_slider"] = opt_thr

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

    st.markdown(_streamlit_css(is_dark), unsafe_allow_html=True)

    # ==================== HERO BANNER ====================
    st.markdown(
        """
        <div class="hero-banner">
            <h1>Hệ thống Dự đoán và Phân tích Khách hàng Rời bỏ Dịch vụ (Customer Churn Analytics)</h1>
            <p>Mô hình phân loại nhạy chi phí (Cost-Sensitive LightGBM) và giải thích cục bộ TreeSHAP</p>
        </div>
        """,
        unsafe_allow_html=True,
    )

    # ==================== 4 KPI CARDS ====================
    audit = results["dataset_audit"]
    best_at_050 = results["test_metrics_at_050"][best_name]
    best_at_tstar = results["test_metrics_at_tstar"][best_name]
    roi_best = results["roi_benchmark"][-1]

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
                <div class="kpi-label">Mô hình Tối ưu</div>
                <div class="kpi-value">{best_name}</div>
                <div class="kpi-sub">5-Fold CV AUC: {results['cv_roc_auc'][best_name]:.4f} • PR-AUC: {best_at_050['pr_auc']:.4f}</div>
            </div>
            """,
            unsafe_allow_html=True,
        )
    with k3:
        st.markdown(
            f"""
            <div class="kpi-card amber">
                <div class="kpi-label">Ngưỡng Nhạy Chi phí</div>
                <div class="kpi-value">T* = {opt_thr:.2f}</div>
                <div class="kpi-sub">Recall tăng {best_at_050['recall']*100:.1f}% -> {best_at_tstar['recall']*100:.2f}% (giữ {best_at_tstar['tp']}/374 Churn)</div>
            </div>
            """,
            unsafe_allow_html=True,
        )
    with k4:
        st.markdown(
            f"""
            <div class="kpi-card teal">
                <div class="kpi-label">Hiệu quả Kinh tế (ROI)</div>
                <div class="kpi-value">Tiết kiệm ${roi_best['savings_vs_nothing_usd']:,.0f}</div>
                <div class="kpi-sub">Giảm {roi_best['savings_pct']:.1f}% chi phí so với không can thiệp</div>
            </div>
            """,
            unsafe_allow_html=True,
        )

    st.markdown("<div style='height: 10px;'></div>", unsafe_allow_html=True)

    # ==================== SIDEBAR CONTROLS ====================
    with st.sidebar:
        st.divider()
        st.markdown("### Thông số mô hình")
        st.markdown(
            f"""
            <div class="sidebar-box">
                <strong>Mô hình:</strong> {best_name}<br>
                <strong>Test ROC-AUC:</strong> {best_at_050['roc_auc']:.4f}<br>
                <strong>Test PR-AUC:</strong> {best_at_050['pr_auc']:.4f}<br>
                <strong>Brier Score:</strong> {best_at_050['brier']:.4f}<br>
                <hr style="margin:8px 0; border-color:rgba(148,163,184,0.3);">
                <strong>Ma trận chi phí nghiệp vụ:</strong><br>
                • Bỏ sót Churn (C_FN): <b>500 USD</b><br>
                • Tặng ưu đãi (C_FP/TP): <b>100 USD</b><br>
                • Tỷ lệ đánh đổi: <b>5 : 1</b>
            </div>
            """,
            unsafe_allow_html=True,
        )

        st.markdown("### Ngưỡng quyết định (T)")
        threshold = st.slider(
            "Ngưỡng phân loại rủi ro (T)",
            min_value=0.10,
            max_value=0.80,
            step=0.01,
            key="threshold_slider",
            help="Mặc định dùng ngưỡng tối ưu nhạy chi phí T* = 0.21.",
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
        st.caption(
            f"Lý thuyết Bayes: T_Bayes = {results['bayes_theoretical_threshold']:.2f} "
            f"(thực nghiệm OOF: T* = {opt_thr:.2f})."
        )

    # ==================== MAIN TABS ====================
    tab1, tab2, tab3 = st.tabs([
        "1. Dự đoán khách hàng cá nhân",
        "2. Chấm điểm danh sách (CSV)",
        "3. Kết quả thực nghiệm mô hình",
    ])

    # ==================== TAB 1 ====================
    with tab1:
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
                monthly = st.number_input(
                    "Cước hàng tháng (MonthlyCharges - USD)",
                    min_value=18.0,
                    max_value=130.0,
                    step=1.0,
                    key="prof_monthly",
                )
                total = st.number_input(
                    "Tổng cước tích lũy (TotalCharges - USD)",
                    min_value=0.0,
                    max_value=10000.0,
                    step=10.0,
                    key="prof_total",
                )
                st.caption(f"Gợi ý theo kỳ hạn: ~{monthly * max(tenure, 1):,.2f} USD")
                paperless = st.selectbox("Hóa đơn điện tử (PaperlessBilling)", ["Yes", "No"], key="prof_paperless")
                payment = st.selectbox(
                    "Phương thức thanh toán (PaymentMethod)",
                    [
                        "Electronic check",
                        "Mailed check",
                        "Bank transfer (automatic)",
                        "Credit card (automatic)",
                    ],
                    key="prof_payment",
                )

        row_raw = {
            "gender": gender,
            "SeniorCitizen": senior,
            "Partner": partner,
            "Dependents": dependents,
            "tenure": tenure,
            "PhoneService": phone_service,
            "MultipleLines": multiple_lines,
            "InternetService": internet,
            "OnlineSecurity": sec,
            "OnlineBackup": backup,
            "DeviceProtection": prot,
            "TechSupport": tech,
            "StreamingTV": tv,
            "StreamingMovies": movies,
            "Contract": contract,
            "PaperlessBilling": paperless,
            "PaymentMethod": payment,
            "MonthlyCharges": monthly,
            "TotalCharges": total,
        }
        df_one = add_domain_features(pd.DataFrame([row_raw]))
        X_one = df_one[feature_cols]
        prob = float(best_pipe.predict_proba(X_one)[0, 1])
        is_risk = prob >= threshold
        prob_pct = prob * 100
        thr_pct = threshold * 100

        st.markdown("##### 6 Đặc trưng miền tự động trích xuất")
        fe_row = df_one.iloc[0]
        f1, f2, f3, f4, f5, f6 = st.columns(6)
        chips = [
            (f1, "Chi tiêu TB/tháng", f"${fe_row['AvgMonthlySpend']:.2f}"),
            (f2, "Chênh lệch cước", f"{fe_row['ChargeDiff']:+.2f} USD"),
            (f3, "Dịch vụ cộng thêm", f"{int(fe_row['TotalAddonServices'])} / 6"),
            (f4, "Bảo mật / HTKT", "Có" if int(fe_row["HasProtectionBundle"]) == 1 else "Không"),
            (f5, "Thanh toán tự động", "Tự động" if int(fe_row["IsAutoPayment"]) == 1 else "Thủ công"),
            (f6, "Nhóm vòng đời", str(fe_row["TenureGroup"])),
        ]
        for col_obj, cname, cval in chips:
            col_obj.markdown(
                f'<div class="feat-chip"><div class="feat-chip-name">{cname}</div>'
                f'<div class="feat-chip-val">{cval}</div></div>',
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
                risk_title = "CẢNH BÁO RỦI RO TRUNG GIAN (BẮT NHỜ NGƯỠNG T*)"
                risk_desc = (
                    f"Xác suất <b>{prob_pct:.2f}%</b> vượt ngưỡng nhạy chi phí T = {threshold:.2f} "
                    f"(bị bỏ sót nếu dùng ngưỡng mặc định 0.50). Nên chủ động giữ chân."
                )
            else:
                risk_css = "risk-low"
                risk_title = "NHÓM KHÁCH HÀNG ỔN ĐỊNH"
                risk_desc = f"Xác suất <b>{prob_pct:.2f}%</b> thấp hơn ngưỡng cảnh báo T = {threshold:.2f}."

            st.markdown(
                f"""
                <div class="gauge-wrap">
                    <div style="display:flex; justify-content:space-between; align-items:baseline;">
                        <span style="font-size:13px; font-weight:600;">Xác suất Rời bỏ Dịch vụ (LightGBM)</span>
                        <span style="font-size:28px; font-weight:800;">{prob_pct:.2f}%</span>
                    </div>
                    <div class="gauge-track">
                        <div class="gauge-threshold" style="left:{thr_pct:.1f}%;"></div>
                        <div class="gauge-marker" style="left:{min(max(prob_pct, 1.0), 99.0):.1f}%;"></div>
                    </div>
                    <div style="display:flex; justify-content:space-between; font-size:11.5px; opacity:0.8;">
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
                X_one_trans = tree_pipe.named_steps["prep"].transform(X_one)
                sv_raw = explainer.shap_values(X_one_trans)
                if isinstance(sv_raw, list):
                    sv_row = np.array(sv_raw[1])[0]
                elif sv_raw.ndim == 3:
                    sv_row = sv_raw[0, :, 1]
                else:
                    sv_row = np.array(sv_raw)[0]
                fig_local = plot_local_shap(sv_row, tree_feat_names, is_dark)
                st.pyplot(fig_local, width="stretch")
                plt.close(fig_local)

    # ==================== TAB 2 ====================
    with tab2:
        st.markdown("#### Chấm điểm rủi ro hàng loạt & Phân tầng giữ chân khách hàng")
        col_dl, col_up = st.columns([1.0, 2.0])
        with col_dl:
            if os.path.exists(SAMPLE_CSV_PATH):
                with open(SAMPLE_CSV_PATH, "rb") as f_sample:
                    st.download_button(
                        "Tải tệp CSV mẫu (100 KH tập Test)",
                        data=f_sample.read(),
                        file_name="sample_test_customers.csv",
                        mime="text/csv",
                        width="stretch",
                    )
        with col_up:
            uploaded = st.file_uploader(
                "Tải lên tệp CSV khách hàng (hoặc xem kết quả chạy sẵn trên 100 KH tập Test bên dưới)",
                type=["csv"],
            )

        raw_batch = None
        source_label = ""
        if uploaded is not None:
            try:
                raw_batch = pd.read_csv(uploaded)
                source_label = f"Tệp tải lên: **{uploaded.name}** ({len(raw_batch):,} khách hàng)"
            except Exception as exc:
                st.error(f"Không thể đọc tệp CSV: {exc}")
        elif os.path.exists(SAMPLE_CSV_PATH):
            raw_batch = pd.read_csv(SAMPLE_CSV_PATH)
            source_label = f"Dữ liệu nạp sẵn: **sample_test_customers.csv** ({len(raw_batch):,} khách hàng tập Test)"

        if raw_batch is not None:
            try:
                st.info(f"Đang chấm điểm nguồn: {source_label} | Ngưỡng phân loại: **T = {threshold:.2f}**")
                ids = raw_batch["customerID"] if "customerID" in raw_batch.columns else pd.Series(range(1, len(raw_batch) + 1))
                proc_batch = add_domain_features(raw_batch)
                probs = best_pipe.predict_proba(proc_batch[feature_cols])[:, 1]

                tiers = np.where(
                    probs >= 0.50,
                    "Nguy cơ cao (>=50%)",
                    np.where(probs >= threshold, f"Cảnh báo ({threshold*100:.0f}%–50%)", "Ổn định"),
                )

                out_df = pd.DataFrame({
                    "customerID": ids,
                    "Contract": raw_batch["Contract"],
                    "InternetService": raw_batch["InternetService"],
                    "tenure": proc_batch["tenure"],
                    "MonthlyCharges": np.round(proc_batch["MonthlyCharges"], 2),
                    "Xac_suat_Churn_Pct": np.round(probs * 100, 2),
                    "Phan_tang_rui_ro": tiers,
                }).sort_values("Xac_suat_Churn_Pct", ascending=False)

                n_high = int((probs >= threshold).sum())
                n_very_high = int((probs >= 0.50).sum())
                est_campaign_cost = n_high * 100

                m1, m2, m3, m4 = st.columns(4)
                m1.metric("Tổng số khách hàng", f"{len(out_df):,}")
                m2.metric(f"KH cần can thiệp (>= {threshold:.2f})", f"{n_high:,}", f"{n_high/len(out_df)*100:.1f}%")
                m3.metric("KH nguy cơ cao (>= 0.50)", f"{n_very_high:,}")
                m4.metric("Ngân sách dự kiến", f"${est_campaign_cost:,}", "100 USD / KH can thiệp")

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
                            "Xác suất Churn (%)",
                            format="%.2f%%",
                            min_value=0,
                            max_value=100,
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
            except Exception as exc:
                st.error(f"Tệp CSV không đúng định dạng yêu cầu của Telco Churn: {exc}")

    # ==================== TAB 3 ====================
    with tab3:
        st.markdown("#### Kết quả thực nghiệm & Tri thức khai phá dữ liệu")
        sub1, sub2, sub3, sub4 = st.tabs([
            "Đối sánh 5 mô hình",
            "Tối ưu chi phí & ROI",
            "Khám phá dữ liệu (EDA)",
            "Giải thích SHAP & Luật CART",
        ])

        with sub1:
            st.markdown("##### Bảng đối sánh 5 mô hình phân loại trên tập Test độc lập (N = 1.409)")
            st.dataframe(comp_df, width="stretch", hide_index=True)

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

        with sub2:
            st.markdown("##### Đối chứng chi phí nghiệp vụ theo 4 chiến lược (C_FN = 500 USD, C_FP = 100 USD)")
            roi_df = pd.DataFrame(results["roi_benchmark"]).rename(columns={
                "strategy": "Chiến lược",
                "targeted_customers": "Số KH can thiệp",
                "churn_caught": "Churn giữ được (TP)",
                "churn_missed": "Churn bỏ sót (FN)",
                "total_cost_usd": "Tổng chi phí (USD)",
                "savings_vs_nothing_usd": "Tiết kiệm vs không can thiệp (USD)",
                "savings_pct": "Tỷ lệ tiết kiệm (%)",
            })
            st.dataframe(roi_df, width="stretch", hide_index=True)

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
            st.dataframe(pd.DataFrame(ab_rows), width="stretch", hide_index=True)

            p_fig7 = os.path.join(FIG_DIR, "fig07_cost_threshold_roi.png")
            if os.path.exists(p_fig7):
                st.image(p_fig7, caption="Hình 7: Tối ưu ngưỡng chi phí T* = 0.21 và đối chứng 4 chiến lược vận hành", width="stretch")

        with sub3:
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

        with sub4:
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
            st.dataframe(rules_df, width="stretch", hide_index=True)

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
            st.dataframe(pd.DataFrame(err_rows), width="stretch", hide_index=True)


if __name__ == "__main__" and _is_running_in_streamlit():
    run_streamlit_app()
