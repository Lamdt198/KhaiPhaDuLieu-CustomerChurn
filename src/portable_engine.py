# -*- coding: utf-8 -*-
"""
src/portable_engine.py
Động cơ suy luận LightGBM và giải thích Exact TreeSHAP bằng Python thuần (chuẩn thư viện).
Không phụ thuộc numpy, pandas, scikit-learn, lightgbm hay shap — tương thích 100% Vercel Serverless.
"""

import csv
import io
import json
import math
import random
from itertools import combinations

ADDON_SERVICE_COLS = [
    "OnlineSecurity",
    "OnlineBackup",
    "DeviceProtection",
    "TechSupport",
    "StreamingTV",
    "StreamingMovies",
]

NUMERIC_FEATURES = [
    "tenure",
    "MonthlyCharges",
    "TotalCharges",
    "AvgMonthlySpend",
    "ChargeDiff",
    "TotalAddonServices",
]

BINARY_NUMERIC_FEATURES = [
    "SeniorCitizen",
    "HasProtectionBundle",
    "IsAutoPayment",
]

CATEGORICAL_FEATURES = [
    "gender",
    "Partner",
    "Dependents",
    "PhoneService",
    "MultipleLines",
    "InternetService",
    "OnlineSecurity",
    "OnlineBackup",
    "DeviceProtection",
    "TechSupport",
    "StreamingTV",
    "StreamingMovies",
    "Contract",
    "PaperlessBilling",
    "PaymentMethod",
    "TenureGroup",
]

FEATURE_VI_MAP = {
    "tenure": "Số tháng gắn bó (tenure)",
    "MonthlyCharges": "Cước hàng tháng (USD)",
    "TotalCharges": "Tổng cước tích lũy (USD)",
    "AvgMonthlySpend": "Chi tiêu TB/tháng (Domain)",
    "ChargeDiff": "Biến động tăng cước (Domain)",
    "TotalAddonServices": "Số dịch vụ cộng thêm (Domain)",
    "SeniorCitizen": "Khách hàng cao tuổi",
    "HasProtectionBundle": "Có gói Bảo mật/HTKT (Domain)",
    "IsAutoPayment": "Thanh toán tự động (Domain)",
    "gender_Female": "Giới tính: Nữ",
    "gender_Male": "Giới tính: Nam",
    "Partner_No": "Không có bạn đời",
    "Partner_Yes": "Có bạn đời",
    "Dependents_No": "Không có người phụ thuộc",
    "Dependents_Yes": "Có người phụ thuộc",
    "PhoneService_No": "Không dùng điện thoại",
    "PhoneService_Yes": "Có dùng điện thoại",
    "MultipleLines_No": "Đơn đường dây điện thoại",
    "MultipleLines_No phone service": "Không có dịch vụ ĐT",
    "MultipleLines_Yes": "Nhiều đường dây điện thoại",
    "InternetService_DSL": "Internet: Cáp đồng DSL",
    "InternetService_Fiber optic": "Internet: Cáp quang Fiber optic",
    "InternetService_No": "Không dùng Internet",
    "OnlineSecurity_No": "Chưa có Bảo mật trực tuyến",
    "OnlineSecurity_No internet service": "Bảo mật: Không có Internet",
    "OnlineSecurity_Yes": "Đã đăng ký Bảo mật trực tuyến",
    "OnlineBackup_No": "Chưa có Sao lưu trực tuyến",
    "OnlineBackup_No internet service": "Sao lưu: Không có Internet",
    "OnlineBackup_Yes": "Đã đăng ký Sao lưu trực tuyến",
    "DeviceProtection_No": "Chưa có Bảo vệ thiết bị",
    "DeviceProtection_No internet service": "BV thiết bị: Không Internet",
    "DeviceProtection_Yes": "Đã đăng ký Bảo vệ thiết bị",
    "TechSupport_No": "Chưa có Hỗ trợ kỹ thuật",
    "TechSupport_No internet service": "HTKT: Không có Internet",
    "TechSupport_Yes": "Đã đăng ký Hỗ trợ kỹ thuật",
    "StreamingTV_No": "Không xem Streaming TV",
    "StreamingTV_No internet service": "TV: Không có Internet",
    "StreamingTV_Yes": "Có dùng Streaming TV",
    "StreamingMovies_No": "Không xem Phim trực tuyến",
    "StreamingMovies_No internet service": "Phim: Không có Internet",
    "StreamingMovies_Yes": "Có xem Phim trực tuyến",
    "Contract_Month-to-month": "Hợp đồng: Theo tháng",
    "Contract_One year": "Hợp đồng: 1 năm",
    "Contract_Two year": "Hợp đồng: 2 năm",
    "PaperlessBilling_No": "Hóa đơn giấy truyền thống",
    "PaperlessBilling_Yes": "Hóa đơn điện tử",
    "PaymentMethod_Bank transfer (automatic)": "Thanh toán: Chuyển khoản tự động",
    "PaymentMethod_Credit card (automatic)": "Thanh toán: Thẻ tín dụng tự động",
    "PaymentMethod_Electronic check": "Thanh toán: Séc điện tử",
    "PaymentMethod_Mailed check": "Thanh toán: Séc qua bưu điện",
    "TenureGroup_0-6m": "Vòng đời: 0–6 tháng",
    "TenureGroup_7-12m": "Vòng đời: 7–12 tháng",
    "TenureGroup_13-24m": "Vòng đời: 13–24 tháng",
    "TenureGroup_25-48m": "Vòng đời: 25–48 tháng",
    "TenureGroup_>48m": "Vòng đời: > 48 tháng",
}

PRESET_PROFILES = {
    "high": {
        "gender": "Male",
        "SeniorCitizen": 0,
        "Partner": "No",
        "Dependents": "No",
        "PhoneService": "Yes",
        "MultipleLines": "No",
        "InternetService": "Fiber optic",
        "OnlineSecurity": "No",
        "OnlineBackup": "No",
        "DeviceProtection": "No",
        "TechSupport": "No",
        "StreamingTV": "Yes",
        "StreamingMovies": "Yes",
        "Contract": "Month-to-month",
        "tenure": 1,
        "MonthlyCharges": 95.1,
        "TotalCharges": 95.1,
        "PaperlessBilling": "Yes",
        "PaymentMethod": "Electronic check",
    },
    "medium": {
        "gender": "Female",
        "SeniorCitizen": 0,
        "Partner": "No",
        "Dependents": "No",
        "PhoneService": "Yes",
        "MultipleLines": "No",
        "InternetService": "DSL",
        "OnlineSecurity": "No",
        "OnlineBackup": "Yes",
        "DeviceProtection": "No",
        "TechSupport": "No",
        "StreamingTV": "No",
        "StreamingMovies": "No",
        "Contract": "Month-to-month",
        "tenure": 10,
        "MonthlyCharges": 60.0,
        "TotalCharges": 600.0,
        "PaperlessBilling": "Yes",
        "PaymentMethod": "Mailed check",
    },
    "low": {
        "gender": "Female",
        "SeniorCitizen": 0,
        "Partner": "Yes",
        "Dependents": "Yes",
        "PhoneService": "Yes",
        "MultipleLines": "Yes",
        "InternetService": "DSL",
        "OnlineSecurity": "Yes",
        "OnlineBackup": "Yes",
        "DeviceProtection": "Yes",
        "TechSupport": "Yes",
        "StreamingTV": "Yes",
        "StreamingMovies": "Yes",
        "Contract": "Two year",
        "tenure": 68,
        "MonthlyCharges": 84.5,
        "TotalCharges": 5746.0,
        "PaperlessBilling": "No",
        "PaymentMethod": "Credit card (automatic)",
    },
}


def load_portable_model(model_path: str) -> dict:
    with open(model_path, "r", encoding="utf-8") as f:
        return json.load(f)


def _to_float(val, default: float = 0.0) -> float:
    if val is None or val == "":
        return default
    try:
        v = float(val)
        return default if math.isnan(v) else v
    except (ValueError, TypeError):
        return default


def _to_int(val, default: int = 0) -> int:
    return int(round(_to_float(val, float(default))))


def engineer_domain_features(row_dict: dict) -> dict:
    """Làm sạch và tạo 6 đặc trưng miền cho một hồ sơ khách hàng."""
    out = dict(row_dict)
    tenure = _to_int(out.get("tenure"), 0)
    monthly = _to_float(out.get("MonthlyCharges"), 0.0)
    total = _to_float(out.get("TotalCharges"), 0.0)
    senior = _to_int(out.get("SeniorCitizen"), 0)

    avg_spend = (total / tenure) if tenure > 0 else monthly
    charge_diff = monthly - avg_spend
    total_addons = sum(1 for col in ADDON_SERVICE_COLS if str(out.get(col, "")).strip() == "Yes")
    has_prot = 1 if (
        str(out.get("OnlineSecurity", "")).strip() == "Yes"
        or str(out.get("TechSupport", "")).strip() == "Yes"
    ) else 0
    is_auto = 1 if "automatic" in str(out.get("PaymentMethod", "")).lower() else 0

    if tenure <= 6:
        t_group = "0-6m"
    elif tenure <= 12:
        t_group = "7-12m"
    elif tenure <= 24:
        t_group = "13-24m"
    elif tenure <= 48:
        t_group = "25-48m"
    else:
        t_group = ">48m"

    out["tenure"] = tenure
    out["MonthlyCharges"] = monthly
    out["TotalCharges"] = total
    out["SeniorCitizen"] = senior
    out["AvgMonthlySpend"] = avg_spend
    out["ChargeDiff"] = charge_diff
    out["TotalAddonServices"] = total_addons
    out["HasProtectionBundle"] = has_prot
    out["IsAutoPayment"] = is_auto
    out["TenureGroup"] = t_group
    return out


def transform_features(proc_dict: dict, model_json: dict) -> list:
    """Biến đổi từ hồ sơ đã tạo đặc trưng miền sang vector 55 chiều khớp ColumnTransformer."""
    vec = []
    # 1. Numeric: SimpleImputer(median) + StandardScaler
    num_mean = model_json["num_mean"]
    num_scale = model_json["num_scale"]
    for i, col in enumerate(NUMERIC_FEATURES):
        v = _to_float(proc_dict.get(col), model_json["num_medians"][i])
        vec.append((v - num_mean[i]) / num_scale[i])

    # 2. Binary numeric: SimpleImputer(most_frequent)
    for i, col in enumerate(BINARY_NUMERIC_FEATURES):
        v = _to_float(proc_dict.get(col), model_json["bin_modes"][i])
        vec.append(v)

    # 3. Categorical: SimpleImputer(most_frequent) + OneHotEncoder
    cat_modes = model_json["cat_modes"]
    cat_cats = model_json["cat_categories"]
    for i, col in enumerate(CATEGORICAL_FEATURES):
        raw_val = proc_dict.get(col)
        s_val = str(raw_val).strip() if raw_val not in (None, "") else str(cat_modes[i])
        for cat in cat_cats[i]:
            vec.append(1.0 if s_val == cat else 0.0)

    return vec


def _eval_tree(node: dict, x: list) -> float:
    while "leaf_value" not in node:
        if x[node["split_feature"]] <= node["threshold"]:
            node = node["left_child"]
        else:
            node = node["right_child"]
    return node["leaf_value"]


def _eval_tree_cond(node: dict, x: list, active_set: set) -> float:
    if "leaf_value" in node:
        return node["leaf_value"]
    f = node["split_feature"]
    if f in active_set:
        nxt = node["left_child"] if x[f] <= node["threshold"] else node["right_child"]
        return _eval_tree_cond(nxt, x, active_set)
    c_left = node["left_child"].get("internal_count", node["left_child"].get("leaf_count", 1))
    c_right = node["right_child"].get("internal_count", node["right_child"].get("leaf_count", 1))
    tot = c_left + c_right
    return (
        (c_left / tot) * _eval_tree_cond(node["left_child"], x, active_set)
        + (c_right / tot) * _eval_tree_cond(node["right_child"], x, active_set)
    )


def _collect_tree_features(node: dict, acc: set):
    if "split_feature" in node:
        acc.add(node["split_feature"])
        _collect_tree_features(node["left_child"], acc)
        _collect_tree_features(node["right_child"], acc)


def predict_proba_row(row_dict: dict, model_json: dict) -> float:
    """Tính xác suất Churn nhanh cho một hồ sơ (dùng cho chấm điểm hàng loạt)."""
    proc = engineer_domain_features(row_dict)
    x = transform_features(proc, model_json)
    raw = sum(_eval_tree(tree, x) for tree in model_json["trees"])
    return 1.0 / (1.0 + math.exp(-raw))


def predict_and_shap(row_dict: dict, model_json: dict, top_k: int = 8) -> dict:
    """Dự đoán xác suất và tính Exact TreeSHAP cho một hồ sơ."""
    proc = engineer_domain_features(row_dict)
    x = transform_features(proc, model_json)
    feat_names = model_json["tree_feature_names"]
    shap_vals = [0.0] * len(feat_names)
    raw_score = 0.0

    for root in model_json["trees"]:
        feats_set = set()
        _collect_tree_features(root, feats_set)
        feats = sorted(feats_set)
        m = len(feats)
        if m == 0:
            raw_score += root.get("leaf_value", 0.0)
            continue

        vals = {}
        for k in range(m + 1):
            for sub in combinations(feats, k):
                vals[frozenset(sub)] = _eval_tree_cond(root, x, set(sub))

        raw_score += vals[frozenset(feats)]
        fact_m = math.factorial(m)
        for f in feats:
            others = [o for o in feats if o != f]
            for k in range(len(others) + 1):
                w = (math.factorial(k) * math.factorial(m - k - 1)) / fact_m
                for sub in combinations(others, k):
                    s0 = frozenset(sub)
                    s1 = frozenset(sub + (f,))
                    shap_vals[f] += w * (vals[s1] - vals[s0])

    prob = 1.0 / (1.0 + math.exp(-raw_score))
    ranked_idx = sorted(range(len(shap_vals)), key=lambda j: abs(shap_vals[j]), reverse=True)[:top_k]
    top_shap = [
        {
            "feature": feat_names[j],
            "label": FEATURE_VI_MAP.get(feat_names[j], feat_names[j]),
            "value": round(shap_vals[j], 4),
        }
        for j in ranked_idx
    ]
    return {
        "probability": prob,
        "shap_values": shap_vals,
        "top_shap": top_shap,
        "domain_features": {
            "AvgMonthlySpend": round(proc["AvgMonthlySpend"], 2),
            "ChargeDiff": round(proc["ChargeDiff"], 2),
            "TotalAddonServices": proc["TotalAddonServices"],
            "HasProtectionBundle": proc["HasProtectionBundle"],
            "IsAutoPayment": proc["IsAutoPayment"],
            "TenureGroup": proc["TenureGroup"],
        },
    }


def build_recommendations(row_dict: dict, prob: float, thr: float) -> list:
    """Sinh danh sách khuyến nghị hành động CSKH thuần văn bản (không dùng emoji)."""
    recs = []
    if prob < thr:
        recs.append(
            "Khách hàng thuộc nhóm ổn định; tiếp tục duy trì chăm sóc định kỳ và khảo sát mức độ hài lòng (CSAT)."
        )
        return recs

    if str(row_dict.get("Contract", "")).strip() == "Month-to-month":
        recs.append(
            "Đề xuất gói ưu đãi chuyển đổi từ hợp đồng theo tháng sang hợp đồng 1 năm "
            "(giảm 10–15% cước 3 tháng đầu hoặc tặng 1 tháng cước)."
        )
    if str(row_dict.get("InternetService", "")).strip() == "Fiber optic" and (
        str(row_dict.get("OnlineSecurity", "")).strip() == "No"
        or str(row_dict.get("TechSupport", "")).strip() == "No"
    ):
        recs.append(
            "Khách dùng cáp quang nhưng chưa kích hoạt bảo mật hoặc hỗ trợ kỹ thuật: "
            "tặng dùng thử miễn phí OnlineSecurity và TechSupport trong 3 tháng."
        )
    if str(row_dict.get("PaymentMethod", "")).strip() == "Electronic check":
        recs.append(
            "Khuyến khích chuyển đổi từ Séc điện tử (Electronic check) sang thanh toán tự động "
            "(tặng ưu đãi nhỏ vào hóa đơn kỳ đầu tiên)."
        )
    if _to_int(row_dict.get("tenure"), 0) <= 12:
        recs.append(
            "Khách hàng đang ở giai đoạn đầu vòng đời (<= 12 tháng): "
            "bộ phận CSKH chủ động gọi kiểm tra chất lượng đường truyền và giải đáp cước."
        )
    if not recs:
        recs.append(
            "Liên hệ khảo sát nguyên nhân chưa hài lòng và áp dụng gói ưu đãi giữ chân phù hợp giá trị cước."
        )
    return recs


def batch_predict_csv(csv_text: str, threshold: float, model_json: dict) -> dict:
    """Chấm điểm toàn bộ tệp CSV khách hàng bằng Python thuần."""
    reader = csv.DictReader(io.StringIO(csv_text.strip()))
    rows_out = []
    n_high_50 = 0
    n_warn = 0
    n_low = 0

    for idx, row in enumerate(reader, start=1):
        prob = predict_proba_row(row, model_json)
        prob_pct = round(prob * 100, 2)
        if prob >= 0.50:
            tier = "Nguy cơ cao (>= 50%)"
            tier_code = "high"
            n_high_50 += 1
        elif prob >= threshold:
            tier = f"Cảnh báo ({threshold * 100:.0f}%–50%)"
            tier_code = "mid"
            n_warn += 1
        else:
            tier = "Ổn định"
            tier_code = "low"
            n_low += 1

        rows_out.append({
            "customerID": row.get("customerID") or str(idx),
            "Contract": row.get("Contract", ""),
            "InternetService": row.get("InternetService", ""),
            "tenure": _to_int(row.get("tenure"), 0),
            "MonthlyCharges": round(_to_float(row.get("MonthlyCharges"), 0.0), 2),
            "Churn_Probability_Pct": prob_pct,
            "Tier": tier,
            "TierCode": tier_code,
        })

    rows_out.sort(key=lambda r: r["Churn_Probability_Pct"], reverse=True)
    total = len(rows_out)
    n_intervene = n_high_50 + n_warn
    return {
        "total": total,
        "n_intervene": n_intervene,
        "n_high_50": n_high_50,
        "n_warn": n_warn,
        "n_low": n_low,
        "intervene_pct": round((n_intervene / total * 100) if total else 0.0, 1),
        "budget_usd": n_intervene * 100,
        "rows": rows_out,
    }


def generate_random_sample_csv(raw_csv_path: str, n_rows: int = 100, seed: int = None) -> str:
    """Trích xuất ngẫu nhiên n_rows khách hàng từ tập dữ liệu gốc (bỏ cột nhãn Churn)."""
    with open(raw_csv_path, "r", encoding="utf-8-sig") as f:
        reader = csv.DictReader(f)
        fieldnames = [c for c in (reader.fieldnames or []) if c != "Churn"]
        all_rows = list(reader)

    n_take = max(1, min(int(n_rows), len(all_rows)))
    rng = random.Random(seed)
    sampled = rng.sample(all_rows, n_take)

    buf = io.StringIO()
    writer = csv.DictWriter(buf, fieldnames=fieldnames, extrasaction="ignore", lineterminator="\n")
    writer.writeheader()
    writer.writerows(sampled)
    return buf.getvalue()

