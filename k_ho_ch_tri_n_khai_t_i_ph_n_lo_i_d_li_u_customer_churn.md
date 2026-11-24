# TỔNG HỢP TRIỂN KHAI BÀI TẬP LỚN — HỌC PHẦN KHAI PHÁ DỮ LIỆU

* **Đơn vị đào tạo:** Trường Đại học Kinh doanh và Công nghệ Hà Nội (HUBT) — Viện Sau Đại học
* **Lớp:** Công nghệ thông tin 20.2
* **Nhóm thực hiện:** Nhóm 02 (Đỗ Ngọc Điền, Dương Tùng Lâm, Trần Thị Thu Hương, Phạm Văn Chiến, Vàng A Hứ)
* **Giảng viên hướng dẫn:** TS. Bùi Ngọc Thăng
* **Chủ đề:** Chủ đề 2 — Phân loại dữ liệu (Data Classification) trong chăm sóc khách hàng
* **Tên đề tài:** *“Ứng dụng các mô hình phân loại dữ liệu trong dự đoán nguy cơ rời bỏ dịch vụ của khách hàng”*
* **Dữ liệu thực nghiệm:** IBM Telco Customer Churn (7.043 bản ghi, 21 cột gốc -> 25 đặc trưng sau làm sạch và tạo 6 đặc trưng miền)
* **Môi trường:** 100% Python cục bộ, không sử dụng API Key.

## 1. Cấu trúc thư mục đã triển khai

```text
Bai tap lon/
├── assets/
│   ├── BaoCao_CHPT-N02_Original_Cover.docx            # Phôi bìa chính thức Viện Sau Đại học HUBT
│   └── logo_hubt_transparent.png                      # Logo HUBT dùng cho trang bìa Slide
├── data/
│   ├── WA_Fn-UseC_-Telco-Customer-Churn.csv           # Bộ dữ liệu thật IBM Telco (7.043 dòng)
│   └── sample_test_customers.csv                      # 100 hồ sơ tập Test làm tệp mẫu upload CSV
├── src/
│   ├── data_loader.py                                 # Đọc, làm sạch và tạo 6 đặc trưng miền
│   ├── eda_stats.py                                   # Kiểm định Chi-Square, Cramér's V, Mann-Whitney U và fig01-fig03
│   └── train_pipeline.py                              # Anti-Leakage 5-Fold CV, 5 mô hình, tối ưu T*, CART, SHAP, Error Analysis, ROI
├── models/
│   ├── best_churn_pipeline.joblib                     # Pipeline LightGBM tối ưu
│   ├── eda_stats.json                                 # Kết quả kiểm định thống kê EDA
│   ├── experiment_results.json                        # Toàn bộ số liệu thực nghiệm đồng nhất
│   └── model_comparison.csv                           # Bảng đối sánh 5 mô hình
├── reports/
│   ├── figures/                                       # 10 biểu đồ độ phân giải 300 DPI (fig01 đến fig10)
│   ├── Bao_Cao_BTL_Khai_Pha_Du_Lieu_Nhom02.docx       # Báo cáo Word hoàn chỉnh chuẩn HUBT
│   └── Slide_Thuyet_Trinh_Khai_Pha_Du_Lieu_Nhom02.pptx# 17 Slides thuyết trình chuẩn HUBT kèm Speaker Notes
├── app.py                                             # Ứng dụng minh họa Streamlit 3 phân hệ
├── build_academic_report.py                           # Sinh tự động báo cáo Word từ số liệu thật
├── build_presentation.py                              # Sinh tự động Slide PPTX từ số liệu thật
├── run_demo.bat                                       # Chạy nhanh ứng dụng Streamlit trên Windows
└── requirements.txt                                   # Danh sách thư viện Python mã nguồn mở
```

## 2. Tóm tắt kết quả thực nghiệm chính (tập Test độc lập N = 1.409)

* **5 mô hình đối sánh (Stratified 5-Fold CV):**
  * LightGBM: CV ROC-AUC = `0,8494` | Test ROC-AUC = `0,8462` | PR-AUC = `0,6657` | Brier = `0,1348`
  * XGBoost: CV ROC-AUC = `0,8493` | Test ROC-AUC = `0,8470` | PR-AUC = `0,6645` | Brier = `0,1347`
  * Random Forest: CV ROC-AUC = `0,8477` | Test ROC-AUC = `0,8444` | PR-AUC = `0,6602` | Brier = `0,1355`
  * Logistic Regression: CV ROC-AUC = `0,8471` | Test ROC-AUC = `0,8452` | PR-AUC = `0,6518` | Brier = `0,1364`
  * Decision Tree (CART): CV ROC-AUC = `0,8295` | Test ROC-AUC = `0,8249` | PR-AUC = `0,6224` | Brier = `0,1410`
* **Tối ưu ngưỡng chi phí (C_FN = $500, C_FP = $100):** Ngưỡng tối ưu trên OOF Train là `T* = 0,21` (khớp sát lý thuyết Bayes `0,20`), nâng Recall của LightGBM từ `51,60%` lên `84,76%`, giữ nguyên Brier Score `0,1348` và giảm tổng chi phí trên tập Test từ `$187.000` (không can thiệp) xuống `$94.300` (tiết kiệm `49,6%`; rẻ hơn luật lọc thủ công Excel `$123.100` thêm `23,4%`).
* **Tri thức rút trích:** 8 luật IF-THEN tường minh từ cây CART (Luật R1: Hợp đồng theo tháng + Fiber optic + tenure <= 13,5 tháng có tỷ lệ rời bỏ `71,0%`), Top 5 đặc trưng SHAP toàn cục (`Contract_Month-to-month`, `tenure`, `OnlineSecurity_No`, `TechSupport_No`, `InternetService_Fiber optic`) và phân tích lỗi chi tiết trên tập Test.

## 3. Hướng dẫn chạy lại toàn bộ dự án

```powershell
python src/eda_stats.py
python src/train_pipeline.py
python build_academic_report.py
python build_presentation.py
python -m streamlit run app.py
```
