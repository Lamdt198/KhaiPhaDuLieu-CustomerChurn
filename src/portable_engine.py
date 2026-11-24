# -*- coding: utf-8 -*-
"""
src/portable_engine.py
Động cơ suy luận 5 mô hình (LightGBM, XGBoost, Random Forest, Decision Tree, Logistic Regression),
giải thích TreeSHAP / Log-Odds và tính toán ma trận chi phí động bằng Python thuần (chuẩn thư viện).
"""

import csv
import io
import json
import math
import random

MODEL_NAMES = [
    "LightGBM",
    "XGBoost",
    "Random Forest",
    "Decision Tree (CART)",
    "Logistic Regression",
]

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
    num_mean = model_json["num_mean"]
    num_scale = model_json["num_scale"]
    for i, col in enumerate(NUMERIC_FEATURES):
        v = _to_float(proc_dict.get(col), model_json["num_medians"][i])
        vec.append((v - num_mean[i]) / num_scale[i])

    for i, col in enumerate(BINARY_NUMERIC_FEATURES):
        v = _to_float(proc_dict.get(col), model_json["bin_modes"][i])
        vec.append(v)

    cat_modes = model_json["cat_modes"]
    cat_cats = model_json["cat_categories"]
    for i, col in enumerate(CATEGORICAL_FEATURES):
        raw_val = proc_dict.get(col)
        s_val = str(raw_val).strip() if raw_val not in (None, "") else str(cat_modes[i])
        for cat in cat_cats[i]:
            vec.append(1.0 if s_val == cat else 0.0)

    return vec


def _eval_compact_tree(node: dict, x: list, strict_lt: bool = False) -> float:
    while "l" not in node:
        go_left = (x[node["f"]] < node["t"]) if strict_lt else (x[node["f"]] <= node["t"])
        node = node["L"] if go_left else node["R"]
    return node["l"]


def _extend_path(path, unique_depth, zero_frac, one_frac, feat_idx):
    """Lundberg et al. (2018) EXTEND step for exact Path-Dependent TreeSHAP."""
    new_path = [dict(p) for p in path[:unique_depth]]
    new_path.append({"d": feat_idx, "z": zero_frac, "o": one_frac, "w": 1.0 if unique_depth == 0 else 0.0})
    for i in range(unique_depth - 1, -1, -1):
        new_path[i + 1]["w"] += one_frac * new_path[i]["w"] * (i + 1) / (unique_depth + 1)
        new_path[i]["w"] = zero_frac * new_path[i]["w"] * (unique_depth - i) / (unique_depth + 1)
    return new_path


def _unwind_path(path, unique_depth, idx):
    """Lundberg et al. (2018) UNWIND step for exact Path-Dependent TreeSHAP."""
    new_path = [dict(p) for p in path[: unique_depth + 1]]
    n = float(new_path[unique_depth]["w"])
    one_frac = new_path[idx]["o"]
    zero_frac = new_path[idx]["z"]
    for i in range(unique_depth - 1, -1, -1):
        if one_frac != 0.0:
            t = new_path[i]["w"]
            new_path[i]["w"] = n * (unique_depth + 1) / ((i + 1) * one_frac)
            n = t - new_path[i]["w"] * zero_frac * (unique_depth - i) / (unique_depth + 1)
        else:
            new_path[i]["w"] = (new_path[i]["w"] * (unique_depth + 1)) / (zero_frac * (unique_depth - i))
    for i in range(idx, unique_depth):
        new_path[i]["d"] = new_path[i + 1]["d"]
        new_path[i]["z"] = new_path[i + 1]["z"]
        new_path[i]["o"] = new_path[i + 1]["o"]
    return new_path


def _unwound_sum(path, unique_depth, idx):
    """Lundberg et al. (2018) UNWOUNDSUM helper."""
    one_frac = path[idx]["o"]
    zero_frac = path[idx]["z"]
    next_one = path[unique_depth]["w"]
    total = 0.0
    if one_frac != 0.0:
        for i in range(unique_depth - 1, -1, -1):
            w_val = next_one * (unique_depth + 1) / ((i + 1) * one_frac)
            total += w_val
            next_one = path[i]["w"] - w_val * zero_frac * (unique_depth - i) / (unique_depth + 1)
    else:
        for i in range(unique_depth - 1, -1, -1):
            total += (path[i]["w"] * (unique_depth + 1)) / (zero_frac * (unique_depth - i))
    return total


def _tree_shap_recursive(node: dict, x: list, shap_acc: list, strict_lt: bool = False, scale: float = 1.0):
    """Tính Exact Path-Dependent TreeSHAP trong O(L * D^2) cho cây quyết định bất kỳ."""
    def recurse(u, m_depth, p, parent_z, parent_o, parent_feat):
        m_ext = _extend_path(p, m_depth, parent_z, parent_o, parent_feat)
        if "l" in u:
            leaf_val = u["l"] * scale
            for i in range(1, m_depth + 1):
                w = _unwound_sum(m_ext, m_depth, i)
                d = m_ext[i]["d"]
                shap_acc[d] += w * (m_ext[i]["o"] - m_ext[i]["z"]) * leaf_val
            return

        feat = u["f"]
        go_left = (x[feat] < u["t"]) if strict_lt else (x[feat] <= u["t"])
        hot_child = u["L"] if go_left else u["R"]
        cold_child = u["R"] if go_left else u["L"]
        c_hot = float(hot_child["c"])
        c_cold = float(cold_child["c"])
        c_tot = float(u["c"]) if u["c"] else (c_hot + c_cold)

        inc_z = 1.0
        inc_o = 1.0
        new_depth = m_depth + 1
        k_idx = -1
        for i in range(1, new_depth):
            if m_ext[i]["d"] == feat:
                k_idx = i
                break
        if k_idx != -1:
            inc_z = m_ext[k_idx]["z"]
            inc_o = m_ext[k_idx]["o"]
            m_ext = _unwind_path(m_ext, new_depth - 1, k_idx)
            new_depth -= 1

        recurse(hot_child, new_depth, m_ext, inc_z * (c_hot / c_tot), inc_o, feat)
        recurse(cold_child, new_depth, m_ext, inc_z * (c_cold / c_tot), 0.0, feat)

    recurse(node, 0, [{}], 1.0, 1.0, -1)


def _predict_vec(x: list, model_json: dict, model_name: str = "LightGBM") -> float:
    if model_name == "LightGBM":
        raw = sum(_eval_compact_tree(t, x, strict_lt=False) for t in model_json["trees"])
        return 1.0 / (1.0 + math.exp(-raw))
    if model_name == "XGBoost":
        raw = model_json["xgb_base_margin"] + sum(_eval_compact_tree(t, x, strict_lt=True) for t in model_json["xgb_trees"])
        return 1.0 / (1.0 + math.exp(-raw))
    if model_name == "Random Forest":
        trees = model_json["rf_trees"]
        return sum(_eval_compact_tree(t, x, strict_lt=False) for t in trees) / len(trees)
    if model_name == "Decision Tree (CART)":
        return _eval_compact_tree(model_json["dt_tree"], x, strict_lt=False)
    if model_name == "Logistic Regression":
        coef = model_json["lr_coef"]
        z = model_json["lr_intercept"] + sum(w * xi for w, xi in zip(coef, x))
        return 1.0 / (1.0 + math.exp(-z))
    raw = sum(_eval_compact_tree(t, x, strict_lt=False) for t in model_json["trees"])
    return 1.0 / (1.0 + math.exp(-raw))


def predict_proba_row(row_dict: dict, model_json: dict, model_name: str = "LightGBM") -> float:
    """Tính xác suất Churn cho một hồ sơ bằng mô hình được chọn."""
    proc = engineer_domain_features(row_dict)
    x = transform_features(proc, model_json)
    return _predict_vec(x, model_json, model_name)


def predict_all_models(row_dict: dict, model_json: dict) -> dict:
    """Tính xác suất Churn của cả 5 mô hình trên cùng một hồ sơ."""
    proc = engineer_domain_features(row_dict)
    x = transform_features(proc, model_json)
    return {name: round(_predict_vec(x, model_json, name), 6) for name in MODEL_NAMES}


def predict_and_shap(row_dict: dict, model_json: dict, top_k: int = 8, model_name: str = "LightGBM") -> dict:
    """Dự đoán xác suất và tính giá trị đóng góp đặc trưng (TreeSHAP / Log-Odds) cho một hồ sơ."""
    proc = engineer_domain_features(row_dict)
    x = transform_features(proc, model_json)
    feat_names = model_json["tree_feature_names"]
    shap_vals = [0.0] * len(feat_names)

    if model_name == "LightGBM":
        for root in model_json["trees"]:
            _tree_shap_recursive(root, x, shap_vals, strict_lt=False, scale=1.0)
    elif model_name == "XGBoost":
        for root in model_json["xgb_trees"]:
            _tree_shap_recursive(root, x, shap_vals, strict_lt=True, scale=1.0)
    elif model_name == "Random Forest":
        trees = model_json["rf_trees"]
        sc = 1.0 / len(trees)
        for root in trees:
            _tree_shap_recursive(root, x, shap_vals, strict_lt=False, scale=sc)
    elif model_name == "Decision Tree (CART)":
        _tree_shap_recursive(model_json["dt_tree"], x, shap_vals, strict_lt=False, scale=1.0)
    elif model_name == "Logistic Regression":
        coef = model_json["lr_coef"]
        means = model_json.get("train_feature_means", [0.0] * len(x))
        shap_vals = [w * (xi - mu) for w, xi, mu in zip(coef, x, means)]

    prob = _predict_vec(x, model_json, model_name)
    all_probs = {name: round(_predict_vec(x, model_json, name), 6) for name in MODEL_NAMES}
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
        "model_name": model_name,
        "probability": prob,
        "all_model_probs": all_probs,
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


def batch_predict_csv(
    csv_text: str,
    threshold: float,
    model_json: dict,
    model_name: str = "LightGBM",
    c_fp: float = 100.0,
) -> dict:
    """Chấm điểm toàn bộ tệp CSV khách hàng bằng mô hình được chọn."""
    reader = csv.DictReader(io.StringIO(csv_text.strip()))
    rows_out = []
    n_high_50 = 0
    n_warn = 0
    n_low = 0

    for idx, row in enumerate(reader, start=1):
        prob = predict_proba_row(row, model_json, model_name=model_name)
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
        "model_name": model_name,
        "total": total,
        "n_intervene": n_intervene,
        "n_high_50": n_high_50,
        "n_warn": n_warn,
        "n_low": n_low,
        "intervene_pct": round((n_intervene / total * 100) if total else 0.0, 1),
        "budget_usd": round(n_intervene * float(c_fp), 2),
        "rows": rows_out,
    }


def compute_threshold_metrics(
    y_true: list,
    y_prob: list,
    threshold: float,
    c_fn: float = 500.0,
    c_fp: float = 100.0,
) -> dict:
    """Tính toàn bộ chỉ số phân loại và chi phí thực nghiệm tại ngưỡng threshold."""
    tn = fp = fn = tp = 0
    for yt, yp in zip(y_true, y_prob):
        pred = 1 if yp >= threshold else 0
        if yt == 1 and pred == 1:
            tp += 1
        elif yt == 0 and pred == 1:
            fp += 1
        elif yt == 1 and pred == 0:
            fn += 1
        else:
            tn += 1
    total = tn + fp + fn + tp
    acc = (tp + tn) / total if total else 0.0
    prec = tp / (tp + fp) if (tp + fp) else 0.0
    rec = tp / (tp + fn) if (tp + fn) else 0.0
    f1 = (2 * prec * rec / (prec + rec)) if (prec + rec) else 0.0
    total_cost = (fp + tp) * float(c_fp) + fn * float(c_fn)
    do_nothing_cost = (tp + fn) * float(c_fn)
    savings = do_nothing_cost - total_cost
    savings_pct = (savings / do_nothing_cost * 100.0) if do_nothing_cost > 0 else 0.0
    return {
        "threshold": round(float(threshold), 2),
        "accuracy": round(acc, 4),
        "precision": round(prec, 4),
        "recall": round(rec, 4),
        "f1": round(f1, 4),
        "tn": tn,
        "fp": fp,
        "fn": fn,
        "tp": tp,
        "targeted": tp + fp,
        "total_cost_usd": round(total_cost, 2),
        "avg_cost_usd": round(total_cost / total if total else 0.0, 2),
        "do_nothing_cost_usd": round(do_nothing_cost, 2),
        "savings_usd": round(savings, 2),
        "savings_pct": round(savings_pct, 1),
    }


def find_optimal_threshold(
    y_true: list,
    y_prob: list,
    c_fn: float = 500.0,
    c_fp: float = 100.0,
) -> tuple:
    """Tìm ngưỡng T* tối ưu chi phí thấp nhất trong [0.05, 0.85] và trả về đường cong chi phí."""
    curve = []
    best_thr = 0.20
    best_cost = float("inf")
    for i in range(5, 86):
        thr = round(i / 100.0, 2)
        m = compute_threshold_metrics(y_true, y_prob, thr, c_fn=c_fn, c_fp=c_fp)
        curve.append({"threshold": thr, "total_cost_usd": m["total_cost_usd"], "recall": m["recall"]})
        if m["total_cost_usd"] < best_cost:
            best_cost = m["total_cost_usd"]
            best_thr = thr
    return best_thr, curve


def compute_dynamic_roi_strategies(
    y_true: list,
    y_prob: list,
    model_name: str,
    threshold: float,
    c_fn: float = 500.0,
    c_fp: float = 100.0,
) -> list:
    """Đối chứng chiến lược Do Nothing vs Blanket 100% vs Mô hình đang chọn tại ngưỡng T."""
    total_n = len(y_true)
    total_churn = sum(1 for y in y_true if y == 1)
    cost_nothing = total_churn * float(c_fn)
    cost_blanket = total_n * float(c_fp)
    m = compute_threshold_metrics(y_true, y_prob, threshold, c_fn=c_fn, c_fp=c_fp)

    def _row(label, targeted, caught, missed, cost):
        sav = cost_nothing - cost
        pct = round((sav / cost_nothing * 100.0) if cost_nothing > 0 else 0.0, 1)
        return {
            "Chiến lược": label,
            "Số KH can thiệp": targeted,
            "Churn giữ được (TP)": caught,
            "Churn bỏ sót (FN)": missed,
            "Tổng chi phí (USD)": round(cost, 2),
            "Tiết kiệm (USD)": round(sav, 2),
            "Tỷ lệ tiết kiệm (%)": pct,
        }

    return [
        _row("Không can thiệp (Do Nothing)", 0, 0, total_churn, cost_nothing),
        _row("Ưu đãi đại trà 100% (Blanket)", total_n, total_churn, 0, cost_blanket),
        _row(f"{model_name} (T = {threshold:.2f})", m["targeted"], m["tp"], m["fn"], m["total_cost_usd"]),
    ]


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
