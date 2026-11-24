# Bài tập lớn Khai phá dữ liệu — Dự đoán Khách hàng Rời bỏ Dịch vụ (Customer Churn)

* **Trường:** Đại học Kinh doanh và Công nghệ Hà Nội (HUBT) — Viện Sau Đại học
* **Lớp:** Cao học Công nghệ thông tin 20.2
* **Học phần:** Khai phá dữ liệu (Data Mining)
* **Giảng viên hướng dẫn:** TS. Bùi Ngọc Thăng
* **Nhóm thực hiện:** Nhóm 02 (Đỗ Ngọc Điền, Dương Tùng Lâm, Trần Thị Thu Hương, Phạm Văn Chiến, Vàng A Hứ)

## 1. Giới thiệu đề tài

Đề tài thực hiện **Chủ đề 2 (Phân loại dữ liệu trong chăm sóc khách hàng)** trên bộ dữ liệu thực **IBM Telco Customer Churn** (7.043 thuê bao, 21 thuộc tính gốc, mở rộng thành 25 đặc trưng sau bước làm sạch và tạo 6 đặc trưng miền). Toàn bộ mã nguồn chạy 100% cục bộ bằng Python mã nguồn mở, không sử dụng API Key.

### Kết quả thực nghiệm chính (tập Test độc lập N = 1.409)

| Mô hình | CV ROC-AUC | Accuracy (0.5) | Recall (0.5) | F1 (0.5) | Test ROC-AUC | PR-AUC | Brier Score | Chi phí (T*=0,21) |
| :--- | :---: | :---: | :---: | :---: | :---: | :---: | :---: | :---: |
| Logistic Regression | 0,8471 | 79,91% | 51,34% | 0,5757 | 0,8452 | 0,6518 | 0,1364 | $94.000 |
| Decision Tree (CART) | 0,8295 | 79,56% | 52,41% | 0,5765 | 0,8249 | 0,6224 | 0,1410 | $102.600 |
| Random Forest | 0,8477 | 80,41% | 50,80% | 0,5793 | 0,8444 | 0,6602 | 0,1355 | $94.900 |
| XGBoost | 0,8493 | 80,84% | 53,21% | 0,5958 | 0,8470 | 0,6645 | 0,1347 | $95.200 |
| **LightGBM** | **0,8494** | **80,70%** | **51,60%** | **0,5866** | **0,8462** | **0,6657** | **0,1348** | **$94.300** |

* Ngưỡng quyết định nhạy chi phí tìm được trên OOF Train là **T\* = 0,21** (khớp sát ngưỡng lý thuyết Bayes 0,20 với tỷ lệ chi phí $C_{FN}:C_{FP} = 5:1$), nâng Recall của LightGBM lên **84,76%** và giảm tổng chi phí trên tập Test từ **$187.000** (không can thiệp) xuống **$94.300** (tiết kiệm **49,6%**).

## 2. Cấu trúc thư mục

```text
├── assets/                   # Phôi bìa gốc HUBT và logo trường
├── data/                     # Bộ dữ liệu IBM Telco Churn và tệp CSV mẫu 100 KH tập Test
├── models/                   # Pipeline LightGBM tối ưu và tệp JSON/CSV kết quả thực nghiệm
├── reports/
│   ├── figures/              # 10 biểu đồ thực nghiệm 300 DPI (fig01 - fig10)
│   ├── Bao_Cao_BTL_Khai_Pha_Du_Lieu_Nhom02.docx
│   └── Slide_Thuyet_Trinh_Khai_Pha_Du_Lieu_Nhom02.pptx
├── src/
│   ├── data_loader.py        # Làm sạch dữ liệu và tạo 6 đặc trưng miền
│   ├── eda_stats.py          # Kiểm định Chi-Square, Cramér's V, Mann-Whitney U và xuất biểu đồ EDA
│   └── train_pipeline.py     # Pipeline chống rò rỉ 5-Fold CV, 5 mô hình, tối ưu T*, CART, SHAP, Error Analysis
├── app.py                    # Ứng dụng minh họa Streamlit (3 phân hệ)
├── build_academic_report.py  # Sinh báo cáo Word từ số liệu thực nghiệm
├── build_presentation.py     # Sinh Slide thuyết trình 17 trang từ số liệu thực nghiệm
├── run_demo.bat              # Khởi chạy nhanh ứng dụng Streamlit trên Windows
└── requirements.txt          # Danh sách thư viện Python
```

## 3. Hướng dẫn cài đặt và chạy lại

```bash
python -m pip install -r requirements.txt
python src/eda_stats.py
python src/train_pipeline.py
python build_academic_report.py
python build_presentation.py
python -m streamlit run app.py
```
