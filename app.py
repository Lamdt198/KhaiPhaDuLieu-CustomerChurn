# -*- coding: utf-8 -*-
"""
app.py
Ung dung Web minh hoa ket qua Khai pha du lieu - Du doan khach hang roi bo dich vu (Customer Churn).
Nhom 02 - Lop Cao hoc CNTT 20.2 - Hoc phan Khai pha du lieu (GVHD: TS. Bui Ngoc Thang).
Chay 100% local, khong su dung API key.
"""

import json
import os
import sys
import joblib
import numpy as np
import pandas as pd
import matplotlib
matplotlib.use("Agg")
import matplotlib.pyplot as plt
import shap
import streamlit as st

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, BASE_DIR)

from src.data_loader import (
    BINARY_NUMERIC_FEATURES,
    CATEGORICAL_FEATURES,
    NUMERIC_FEATURES,
    add_domain_features,
)

MODEL_PATH = os.path.join(BASE_DIR, "models", "best_churn_pipeline.joblib")
RESULTS_PATH = os.path.join(BASE_DIR, "models", "experiment_results.json")
COMP_CSV_PATH = os.path.join(BASE_DIR, "models", "model_comparison.csv")
SAMPLE_CSV_PATH = os.path.join(BASE_DIR, "data", "sample_test_customers.csv")
FIG_DIR = os.path.join(BASE_DIR, "reports", "figures")

st.set_page_config(
    page_title="Dự đoán Khách hàng Rời bỏ Dịch vụ - Nhóm 02",
    layout="wide",
    initial_sidebar_state="expanded",
)

st.markdown(
    """
    <style>
    .main-header {
        background-color: #0F2544;
        padding: 18px 24px;
        border-radius: 8px;
        border-bottom: 4px solid #D97706;
        margin-bottom: 18px;
    }
    .main-header h2 {
        color: #FFFFFF;
        margin: 0;
        font-size: 22px;
    }
    .main-header p {
        color: #CBD5E1;
        margin: 4px 0 0 0;
        font-size: 13.5px;
    }
    .risk-high {
        background-color: #FEF2F2;
        border-left: 5px solid #B91C1C;
        padding: 14px 18px;
        border-radius: 6px;
        color: #7F1D1D;
    }
    .risk-low {
        background-color: #F0FDF4;
        border-left: 5px solid #0F766E;
        padding: 14px 18px;
        border-radius: 6px;
        color: #134E4A;
    }
    </style>
    """,
    unsafe_allow_html=True,
)


@st.cache_resource
def load_artifacts():
    bundle = joblib.load(MODEL_PATH)
    with open(RESULTS_PATH, "r", encoding="utf-8") as f:
        results = json.load(f)
    comp_df = pd.read_csv(COMP_CSV_PATH)
    explainer = shap.TreeExplainer(bundle["best_tree_pipeline"].named_steps["clf"])
    return bundle, results, comp_df, explainer


def build_recommendations(row_dict: dict, prob: float, thr: float) -> list:
    recs = []
    if prob < thr:
        recs.append("Khách hàng thuộc nhóm ổn định; duy trì chăm sóc định kỳ và đo lường mức độ hài lòng.")
        return recs

    if row_dict["Contract"] == "Month-to-month":
        recs.append(
            "Đề xuất gói ưu đãi chuyển đổi từ hợp đồng theo tháng sang hợp đồng 1 năm "
            "(ví dụ: giảm 10-15% cước 3 tháng đầu hoặc tặng 1 tháng cước)."
        )
    if row_dict["InternetService"] == "Fiber optic" and (
        row_dict["OnlineSecurity"] == "No" or row_dict["TechSupport"] == "No"
    ):
        recs.append(
            "Khách dùng cáp quang nhưng chưa kích hoạt bảo mật/hỗ trợ kỹ thuật: "
            "tặng dùng thử miễn phí gói OnlineSecurity và TechSupport trong 3 tháng."
        )
    if row_dict["PaymentMethod"] == "Electronic check":
        recs.append(
            "Khuyến khích chuyển đổi phương thức thanh toán từ Electronic check sang thanh toán tự động "
            "(tặng ưu đãi nhỏ vào hóa đơn kỳ đầu tiên)."
        )
    if row_dict["tenure"] <= 12:
        recs.append(
            "Khách hàng đang ở giai đoạn đầu vòng đời (<= 12 tháng): "
            "bộ phận CSKH nên chủ động liên hệ kiểm tra chất lượng đường truyền và giải đáp cước phí."
        )
    if not recs:
        recs.append(
            "Liên hệ khảo sát nguyên nhân không hài lòng và áp dụng gói ưu đãi giữ chân phù hợp giá trị cước."
        )
    return recs


def plot_local_shap(shap_vals_row, feature_names):
    top_k = np.argsort(np.abs(shap_vals_row))[::-1][:8][::-1]
    labels = [feature_names[j] for j in top_k]
    vals = [float(shap_vals_row[j]) for j in top_k]
    colors = ["#B91C1C" if v > 0 else "#0F766E" for v in vals]

    fig, ax = plt.subplots(figsize=(7.2, 4.0), dpi=150)
    fig.patch.set_facecolor("white")
    bars = ax.barh(labels, vals, color=colors, height=0.58)
    ax.axvline(0, color="#0F2544", lw=1.0)
    span = max(abs(min(vals)), abs(max(vals))) + 0.12
    ax.set_xlim(-span * 1.28, span * 1.28)
    for b, v in zip(bars, vals):
        ha = "left" if v >= 0 else "right"
        off = span * 0.04 if v >= 0 else -span * 0.04
        ax.text(v + off, b.get_y() + b.get_height() / 2, f"{v:+.3f}",
                va="center", ha=ha, fontsize=8.5, fontweight="bold", color="#0F172A")
    ax.set_title("Đóng góp SHAP của 8 đặc trưng ảnh hưởng lớn nhất tới hồ sơ này",
                 fontsize=10, fontweight="bold", color="#0F2544")
    ax.set_xlabel("Giá trị SHAP (Đỏ: tăng nguy cơ rời bỏ | Xanh: giảm nguy cơ)", fontsize=8.8)
    ax.xaxis.grid(True, linestyle="--", alpha=0.6)
    ax.set_axisbelow(True)
    ax.spines["top"].set_visible(False)
    ax.spines["right"].set_visible(False)
    plt.tight_layout()
    return fig


def main():
    if not os.path.exists(MODEL_PATH):
        st.error("Chưa tìm thấy mô hình đã huấn luyện. Vui lòng chạy `python src/train_pipeline.py` trước.")
        return

    bundle, results, comp_df, explainer = load_artifacts()
    best_pipe = bundle["best_model_pipeline"]
    tree_pipe = bundle["best_tree_pipeline"]
    best_name = bundle["best_model_name"]
    opt_thr = bundle["optimal_threshold"]
    tree_feat_names = bundle["tree_feature_names"]
    feature_cols = NUMERIC_FEATURES + BINARY_NUMERIC_FEATURES + CATEGORICAL_FEATURES

    st.markdown(
        """
        <div class="main-header">
            <h2>Hệ thống Minh họa Phân loại Nguy cơ Rời bỏ Dịch vụ Khách hàng (Customer Churn)</h2>
            <p>Học phần Khai phá dữ liệu — Lớp Cao học CNTT 20.2 — Nhóm 02 — GVHD: TS. Bùi Ngọc Thăng</p>
        </div>
        """,
        unsafe_allow_html=True,
    )

    with st.sidebar:
        st.subheader("Thông tin mô hình")
        st.write(f"**Mô hình tối ưu:** {best_name}")
        st.write(f"**CV ROC-AUC (5-Fold):** {results['cv_roc_auc'][best_name]:.4f}")
        st.write(f"**Test ROC-AUC:** {results['test_metrics_at_050'][best_name]['roc_auc']:.4f}")
        st.write(f"**Test PR-AUC:** {results['test_metrics_at_050'][best_name]['pr_auc']:.4f}")
        st.divider()
        st.subheader("Cấu hình ngưỡng cảnh báo")
        threshold = st.slider(
            "Ngưỡng xác suất phân loại (T)",
            min_value=0.10,
            max_value=0.80,
            value=float(opt_thr),
            step=0.01,
            help="Mặc định sử dụng ngưỡng tối ưu nhạy chi phí T* tìm được trên tập kiểm thử chéo OOF.",
        )
        st.caption(
            f"Ngưỡng tối ưu chi phí thực nghiệm: T* = {opt_thr:.2f} "
            f"(Lý thuyết Bayes: {results['bayes_theoretical_threshold']:.2f}, tỷ lệ chi phí C_FN:C_FP = 5:1)."
        )

    tab1, tab2, tab3 = st.tabs([
        "1. Dự đoán khách hàng cá nhân",
        "2. Chấm điểm hàng loạt (CSV)",
        "3. Kết quả thực nghiệm & Tri thức khai phá",
    ])

    # ==================== TAB 1: DU DOAN CA NHAN + SHAP ====================
    with tab1:
        st.subheader("Nhập thông tin hồ sơ khách hàng")
        c1, c2, c3 = st.columns(3)

        with c1:
            st.markdown("**Nhân khẩu học & Điện thoại**")
            gender = st.selectbox("Giới tính (gender)", ["Male", "Female"])
            senior = st.selectbox("Khách hàng cao tuổi (SeniorCitizen)", [0, 1], format_func=lambda x: "Có" if x == 1 else "Không")
            partner = st.selectbox("Có bạn đời (Partner)", ["No", "Yes"])
            dependents = st.selectbox("Có người phụ thuộc (Dependents)", ["No", "Yes"])
            phone_service = st.selectbox("Dịch vụ điện thoại (PhoneService)", ["Yes", "No"])
            multiple_lines = st.selectbox("Nhiều đường dây (MultipleLines)", ["No", "Yes", "No phone service"])

        with c2:
            st.markdown("**Dịch vụ Internet & Gia tăng**")
            internet = st.selectbox("Hạ tầng Internet (InternetService)", ["Fiber optic", "DSL", "No"])
            sec = st.selectbox("Bảo mật trực tuyến (OnlineSecurity)", ["No", "Yes", "No internet service"])
            backup = st.selectbox("Sao lưu trực tuyến (OnlineBackup)", ["No", "Yes", "No internet service"])
            prot = st.selectbox("Bảo vệ thiết bị (DeviceProtection)", ["No", "Yes", "No internet service"])
            tech = st.selectbox("Hỗ trợ kỹ thuật (TechSupport)", ["No", "Yes", "No internet service"])
            tv = st.selectbox("Truyền hình (StreamingTV)", ["No", "Yes", "No internet service"])
            movies = st.selectbox("Phim trực tuyến (StreamingMovies)", ["No", "Yes", "No internet service"])

        with c3:
            st.markdown("**Hợp đồng & Cước phí**")
            contract = st.selectbox("Loại hợp đồng (Contract)", ["Month-to-month", "One year", "Two year"])
            tenure = st.slider("Số tháng gắn bó (tenure)", 0, 72, 4)
            monthly = st.number_input("Cước hàng tháng (MonthlyCharges - USD)", min_value=18.0, max_value=130.0, value=85.0, step=1.0)
            total_default = round(monthly * max(tenure, 1), 2)
            total = st.number_input("Tổng cước tích lũy (TotalCharges - USD)", min_value=0.0, max_value=10000.0, value=total_default, step=10.0)
            paperless = st.selectbox("Hóa đơn điện tử (PaperlessBilling)", ["Yes", "No"])
            payment = st.selectbox(
                "Phương thức thanh toán (PaymentMethod)",
                [
                    "Electronic check",
                    "Mailed check",
                    "Bank transfer (automatic)",
                    "Credit card (automatic)",
                ],
            )

        if st.button("Phân tích nguy cơ rời bỏ", type="primary"):
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

            st.divider()
            r1, r2 = st.columns([1.0, 1.35])
            with r1:
                st.metric("Xác suất rời bỏ dịch vụ (Churn Probability)", f"{prob*100:.2f}%")
                if is_risk:
                    st.markdown(
                        f'<div class="risk-high"><strong>CẢNH BÁO NGUY CƠ RỜI BỎ CAO</strong><br>'
                        f'Xác suất ({prob*100:.1f}%) vượt ngưỡng quyết định T = {threshold:.2f}. '
                        f'Cần đưa vào danh sách ưu tiên chăm sóc giữ chân.</div>',
                        unsafe_allow_html=True,
                    )
                else:
                    st.markdown(
                        f'<div class="risk-low"><strong>NHÓM KHÁCH HÀNG ỔN ĐỊNH</strong><br>'
                        f'Xác suất ({prob*100:.1f}%) thấp hơn ngưỡng cảnh báo T = {threshold:.2f}.</div>',
                        unsafe_allow_html=True,
                    )

                st.markdown("#### Khuyến nghị hành động CSKH")
                for rec in build_recommendations(row_raw, prob, threshold):
                    st.write(f"- {rec}")

            with r2:
                X_one_trans = tree_pipe.named_steps["prep"].transform(X_one)
                sv_raw = explainer.shap_values(X_one_trans)
                if isinstance(sv_raw, list):
                    sv_row = np.array(sv_raw[1])[0]
                elif sv_raw.ndim == 3:
                    sv_row = sv_raw[0, :, 1]
                else:
                    sv_row = np.array(sv_raw)[0]
                fig_local = plot_local_shap(sv_row, tree_feat_names)
                st.pyplot(fig_local)
                plt.close(fig_local)

    # ==================== TAB 2: CHAM DIEM HANG LOAT CSV ====================
    with tab2:
        st.subheader("Dự đoán rủi ro hàng loạt từ tệp CSV")
        col_dl, col_up = st.columns([1, 2])
        with col_dl:
            if os.path.exists(SAMPLE_CSV_PATH):
                with open(SAMPLE_CSV_PATH, "rb") as f_sample:
                    st.download_button(
                        "Tải tệp CSV mẫu (100 KH tập Test)",
                        data=f_sample.read(),
                        file_name="sample_test_customers.csv",
                        mime="text/csv",
                    )
        with col_up:
            uploaded = st.file_uploader("Tải lên tệp CSV danh sách khách hàng", type=["csv"])

        if uploaded is not None:
            try:
                raw_batch = pd.read_csv(uploaded)
                ids = raw_batch["customerID"] if "customerID" in raw_batch.columns else pd.Series(range(1, len(raw_batch) + 1))
                proc_batch = add_domain_features(raw_batch)
                probs = best_pipe.predict_proba(proc_batch[feature_cols])[:, 1]

                out_df = pd.DataFrame({
                    "customerID": ids,
                    "Contract": raw_batch["Contract"],
                    "tenure": proc_batch["tenure"],
                    "MonthlyCharges": proc_batch["MonthlyCharges"],
                    "Churn_Probability_Pct": np.round(probs * 100, 2),
                    "Phan_loai": np.where(probs >= threshold, "Nguy cơ cao", "Ổn định"),
                }).sort_values("Churn_Probability_Pct", ascending=False)

                n_high = int((probs >= threshold).sum())
                m1, m2, m3 = st.columns(3)
                m1.metric("Tổng số khách hàng", f"{len(out_df):,}")
                m2.metric(f"Số KH vượt ngưỡng T={threshold:.2f}", f"{n_high:,}")
                m3.metric("Tỷ lệ cảnh báo", f"{n_high/len(out_df)*100:.1f}%")

                st.dataframe(out_df, use_container_width=True, height=380)
                st.download_button(
                    "Tải xuống kết quả dự đoán (CSV)",
                    data=out_df.to_csv(index=False).encode("utf-8-sig"),
                    file_name="churn_batch_predictions.csv",
                    mime="text/csv",
                )
            except Exception as exc:
                st.error(f"Tệp CSV không đúng định dạng yêu cầu của Telco Churn: {exc}")

    # ==================== TAB 3: TONG QUAN KET QUA THUC NGHIEM ====================
    with tab3:
        st.subheader("Bảng đối sánh 5 mô hình phân loại trên tập Test độc lập (N = 1.409)")
        st.dataframe(comp_df, use_container_width=True, hide_index=True)

        st.subheader("Đối chứng chi phí nghiệp vụ theo 4 chiến lược (C_FN = $500, C_FP = $100)")
        roi_df = pd.DataFrame(results["roi_benchmark"]).rename(columns={
            "strategy": "Chiến lược",
            "targeted_customers": "Số KH can thiệp",
            "churn_caught": "Churn giữ được (TP)",
            "churn_missed": "Churn bỏ sót (FN)",
            "total_cost_usd": "Tổng chi phí (USD)",
            "savings_vs_nothing_usd": "Tiết kiệm so với không can thiệp (USD)",
            "savings_pct": "Tỷ lệ tiết kiệm (%)",
        })
        st.dataframe(roi_df, use_container_width=True, hide_index=True)

        st.subheader("Biểu đồ kết quả thực nghiệm và Giải thích SHAP")
        figs_to_show = [
            ("fig05_roc_pr_curves.png", "Đường cong ROC-AUC và Precision-Recall (PR-AUC) của 5 mô hình"),
            ("fig07_cost_threshold_roi.png", "Tối ưu ngưỡng chi phí T* và đối chứng 4 chiến lược vận hành"),
            ("fig09_shap_summary.png", "Giải thích toàn cục SHAP Summary Plot (Top 12 đặc trưng)"),
            ("fig08_decision_tree_rules.png", "Cây quyết định CART cắt tỉa độ sâu 3 phục vụ rút trích luật IF-THEN"),
        ]
        for fname, cap in figs_to_show:
            fpath = os.path.join(FIG_DIR, fname)
            if os.path.exists(fpath):
                st.image(fpath, caption=cap, use_container_width=True)


if __name__ == "__main__":
    main()
