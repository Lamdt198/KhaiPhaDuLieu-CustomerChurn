# -*- coding: utf-8 -*-
"""
src/generate_flowchart.py
Sinh so do khoi (Flowchart) quy trinh thuc nghiem KDD & Anti-Leakage Pipeline
do phan giai 300 DPI dong bo he mau hoc thuat HUBT cho Muc 1.2 cua Bao cao.
"""

import os
import matplotlib
matplotlib.use("Agg")
import matplotlib.pyplot as plt
from matplotlib.patches import FancyBboxPatch, FancyArrowPatch

BASE_DIR = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
FIG_DIR = os.path.join(BASE_DIR, "reports", "figures")
OUT_PATH = os.path.join(FIG_DIR, "fig00_kdd_pipeline_flowchart.png")


def draw_box(ax, x, y, w, h, step_badge, title, lines, header_color, border_color, bg_color="#FFFFFF"):
    """Ve hop bo goc 2 tang: thanh tieu de mau dam + than hop chua chi tiet."""
    # Bong do nhe
    shadow = FancyBboxPatch(
        (x + 0.06, y - 0.06), w, h,
        boxstyle="round,pad=0.02,rounding_size=0.14",
        linewidth=0,
        facecolor="#CBD5E1",
        alpha=0.45,
        zorder=1,
    )
    ax.add_patch(shadow)

    # Khung chinh
    body = FancyBboxPatch(
        (x, y), w, h,
        boxstyle="round,pad=0.02,rounding_size=0.14",
        linewidth=1.5,
        edgecolor=border_color,
        facecolor=bg_color,
        zorder=2,
    )
    ax.add_patch(body)

    # Thanh header
    hdr_h = 0.56
    hdr = FancyBboxPatch(
        (x, y + h - hdr_h), w, hdr_h,
        boxstyle="round,pad=0.02,rounding_size=0.12",
        linewidth=0,
        facecolor=header_color,
        zorder=3,
    )
    ax.add_patch(hdr)

    # Che goc duoi cua header de vuong phang
    from matplotlib.patches import Rectangle
    rect_fix = Rectangle(
        (x - 0.01, y + h - hdr_h - 0.02), w + 0.02, 0.16,
        linewidth=0,
        facecolor=header_color,
        zorder=3,
    )
    ax.add_patch(rect_fix)

    # Vien lai khung ngoai de sac net
    border_overlay = FancyBboxPatch(
        (x, y), w, h,
        boxstyle="round,pad=0.02,rounding_size=0.14",
        linewidth=1.5,
        edgecolor=border_color,
        facecolor="none",
        zorder=5,
    )
    ax.add_patch(border_overlay)

    # Badge buoc + tieu de
    ax.text(
        x + w / 2,
        y + h - hdr_h / 2 + 0.01,
        f"{step_badge}  {title}",
        ha="center",
        va="center",
        fontsize=9.6,
        fontweight="bold",
        color="#FFFFFF",
        zorder=6,
    )

    # Noi dung chi tiet trong hop
    content_top = y + h - hdr_h - 0.18
    line_step = 0.30
    for idx, line in enumerate(lines):
        ax.text(
            x + w / 2,
            content_top - idx * line_step,
            line,
            ha="center",
            va="top",
            fontsize=8.6,
            color="#1E293B",
            linespacing=1.22,
            zorder=6,
        )


def draw_arrow(ax, x1, y1, x2, y2, color="#1E3A8A", rad=0.0, label=None, lx=None, ly=None):
    arr = FancyArrowPatch(
        (x1, y1),
        (x2, y2),
        connectionstyle=f"arc3,rad={rad}",
        arrowstyle="-|>,head_length=7,head_width=4.5",
        linewidth=1.9,
        color=color,
        zorder=4,
    )
    ax.add_patch(arr)
    if label:
        ax.text(
            lx if lx is not None else (x1 + x2) / 2,
            ly if ly is not None else (y1 + y2) / 2 + 0.12,
            label,
            ha="center",
            va="center",
            fontsize=8.1,
            fontweight="bold",
            color=color,
            bbox=dict(boxstyle="round,pad=0.18", facecolor="#EFF6FF", edgecolor="#BFDBFE", lw=0.8),
            zorder=7,
        )


def generate_flowchart():
    os.makedirs(FIG_DIR, exist_ok=True)
    plt.rcParams["font.family"] = "DejaVu Sans"

    fig, ax = plt.subplots(figsize=(13.2, 6.9), dpi=300)
    ax.set_xlim(0, 13.2)
    ax.set_ylim(0, 6.9)
    ax.axis("off")
    fig.patch.set_facecolor("#FFFFFF")

    # Tieu de bieu do
    ax.text(
        6.6, 6.58,
        "SƠ ĐỒ TỔNG THỂ QUY TRÌNH THỰC NGHIỆM KHAI PHÁ DỮ LIỆU (CHUẨN KDD & ANTI-LEAKAGE)",
        ha="center", va="center",
        fontsize=11.8, fontweight="bold", color="#0F2544"
    )
    ax.text(
        6.6, 6.26,
        "Luồng biến đổi dữ liệu khép kín từ 7.043 thuê bao gốc đến hệ thống dự đoán nhạy chi phí và giải thích TreeSHAP",
        ha="center", va="center",
        fontsize=9.1, style="italic", color="#475569"
    )

    # Khung bao Anti-Leakage Zone cho hang duoi
    leak_zone = FancyBboxPatch(
        (0.25, 0.25), 12.70, 2.72,
        boxstyle="round,pad=0.02,rounding_size=0.18",
        linewidth=1.3,
        linestyle="--",
        edgecolor="#2563EB",
        facecolor="#F8FAFC",
        zorder=0,
    )
    ax.add_patch(leak_zone)
    ax.text(
        6.6, 2.82,
        "VÙNG HUẤN LUYỆN – TỐI ƯU – ĐÁNH GIÁ CHỐNG RÒ RỈ DỮ LIỆU (fit() chỉ thực hiện trên lát cắt Train)",
        ha="center", va="center",
        fontsize=8.3, fontweight="bold", color="#1E40AF",
        bbox=dict(boxstyle="round,pad=0.20", facecolor="#EFF6FF", edgecolor="#93C5FD", lw=0.9),
        zorder=7,
    )

    # ==================== TANG 1 (3 KHOI: TRAI -> PHAI) ====================
    y_top = 3.65
    h_top = 2.15
    w_top = 3.82

    # Buoc 1
    draw_box(
        ax, 0.35, y_top, w_top, h_top,
        "[1]", "THU THẬP & LÀM SẠCH",
        [
            "Bộ dữ liệu IBM Telco Customer Churn",
            "• Quy mô: 7.043 thuê bao × 21 cột gốc",
            "• Loại bỏ mã định danh customerID",
            "• Chuyển kiểu TotalCharges sang số thực,",
            "  điền 0,0 cho 11 mẫu mới (tenure = 0)",
        ],
        header_color="#0F2544",
        border_color="#1E3A8A",
    )

    # Buoc 2
    draw_box(
        ax, 4.69, y_top, w_top, h_top,
        "[2]", "TẠO ĐẶC TRƯNG MIỀN",
        [
            "Thiết kế 6 biến tri thức nghiệp vụ mới:",
            "• AvgMonthlySpend, ChargeDiff",
            "• TotalAddonServices (0–6 dịch vụ)",
            "• HasProtectionBundle, IsAutoPayment",
            "• TenureGroup (5 giai đoạn vòng đời)",
        ],
        header_color="#1E3A8A",
        border_color="#2563EB",
    )

    # Buoc 3
    draw_box(
        ax, 9.03, y_top, w_top, h_top,
        "[3]", "CHIA TẬP PHÂN TẦNG 80/20",
        [
            "Stratified Split (random_state = 42)",
            "• Giữ nguyên tỷ lệ Churn = 26,54%",
            "• Tập Train: 5.634 mẫu (80,0%)",
            "• Tập Test độc lập: 1.409 mẫu (20,0%)",
            "  (khóa hoàn toàn tới bước kiểm định)",
        ],
        header_color="#0F766E",
        border_color="#0D9488",
    )

    # Mui ten tang 1
    draw_arrow(ax, 4.19, y_top + h_top / 2, 4.67, y_top + h_top / 2, color="#1E3A8A", label="19 biến sạch", ly=y_top + h_top / 2 + 0.26)
    draw_arrow(ax, 8.53, y_top + h_top / 2, 9.01, y_top + h_top / 2, color="#1E3A8A", label="25 đặc trưng", ly=y_top + h_top / 2 + 0.26)

    # ==================== TANG 2 (4 KHOI: PHAI -> TRAI THEO LUONG KHEP KIN HOAC TRAI -> PHAI) ====================
    # De nguoi doc de theo doi tu trai sang phai o ca 2 tang, ta dan duong chuyen tang tu [3] xuong [4] ben trai
    y_bot = 0.45
    h_bot = 2.15
    w_bot = 2.88
    xs_bot = [0.42, 3.60, 6.78, 9.90]

    # Duong dan tu day khoi [3] (x=10.94, y=3.65) vong muot sang dinh khoi [4] (x=1.86, y=2.60)
    ax.plot(
        [10.45, 10.45, 1.86, 1.86],
        [3.65, 3.24, 3.24, 2.78],
        color="#0F766E", linewidth=1.9, zorder=4
    )
    draw_arrow(ax, 1.86, 2.86, 1.86, 2.62, color="#0F766E")
    ax.text(
        6.15, 3.24,
        "Đưa Tập Train (5.634 mẫu) vào Pipeline chống rò rỉ",
        ha="center", va="center",
        fontsize=8.1, fontweight="bold", color="#0F766E",
        bbox=dict(boxstyle="round,pad=0.18", facecolor="#F0FDFA", edgecolor="#99F6E4", lw=0.9),
        zorder=7,
    )

    # Duong rieng cho Tap Test (1.409 mau) di thang xuong khoi [7] de danh gia doc lap!
    ax.plot(
        [12.15, 12.15],
        [3.65, 2.78],
        color="#B91C1C", linewidth=1.6, linestyle="--", zorder=4
    )
    draw_arrow(ax, 12.15, 2.86, 12.15, 2.62, color="#B91C1C")
    ax.text(
        12.15, 3.24,
        "Test độc lập\n(1.409 mẫu)",
        ha="center", va="center",
        fontsize=7.7, fontweight="bold", color="#B91C1C",
        bbox=dict(boxstyle="round,pad=0.16", facecolor="#FEF2F2", edgecolor="#FECACA", lw=0.9),
        zorder=7,
    )

    # Buoc 4
    draw_box(
        ax, xs_bot[0], y_bot, w_bot, h_bot,
        "[4]", "COLUMN TRANSFORMER",
        [
            "Biến đổi 25 biến → 55 chiều:",
            "• 6 biến số: Median + Z-score",
            "• 3 biến nhị phân: giữ nguyên 0/1",
            "• 16 biến phân loại: One-Hot",
            "  (handle_unknown='ignore')",
        ],
        header_color="#1E40AF",
        border_color="#3B82F6",
    )

    # Buoc 5
    draw_box(
        ax, xs_bot[1], y_bot, w_bot, h_bot,
        "[5]", "5-FOLD GRIDSEARCH CV",
        [
            "Đối sánh 5 mô hình phân loại:",
            "• Logistic Regression, CART",
            "• Random Forest, XGBoost",
            "• LightGBM (CV AUC = 0,8494)",
            "Chọn mô hình tối ưu trên Train",
        ],
        header_color="#1E3A8A",
        border_color="#2563EB",
    )

    # Buoc 6
    draw_box(
        ax, xs_bot[2], y_bot, w_bot, h_bot,
        "[6]", "TỐI ƯU NGƯỠNG T*",
        [
            "Dự đoán Out-Of-Fold (OOF):",
            "• Ma trận chi phí bất đối xứng",
            "  C_FN = 500 USD | C_FP = 100 USD",
            "• Tối thiểu hóa tổng chi phí",
            "→ Ngưỡng tối ưu T* = 0,21",
        ],
        header_color="#B45309",
        border_color="#D97706",
    )

    # Buoc 7
    draw_box(
        ax, xs_bot[3], y_bot, w_bot, h_bot,
        "[7]", "XAI & ỨNG DỤNG WEB",
        [
            "Kiểm định Test & Triển khai:",
            "• Recall 84,76%, giảm 49,6% chi phí",
            "• 8 luật IF-THEN (CART depth 3)",
            "• Giải thích TreeSHAP toàn cục/cá nhân",
            "• Ứng dụng Streamlit (3 phân hệ)",
        ],
        header_color="#0F2544",
        border_color="#0F766E",
        bg_color="#F0FDFA",
    )

    # Mui ten tang 2
    draw_arrow(ax, xs_bot[0] + w_bot + 0.02, y_bot + h_bot / 2, xs_bot[1] - 0.02, y_bot + h_bot / 2, color="#1E3A8A")
    draw_arrow(ax, xs_bot[1] + w_bot + 0.02, y_bot + h_bot / 2, xs_bot[2] - 0.02, y_bot + h_bot / 2, color="#B45309")
    draw_arrow(ax, xs_bot[2] + w_bot + 0.02, y_bot + h_bot / 2, xs_bot[3] - 0.02, y_bot + h_bot / 2, color="#0F766E")

    plt.tight_layout(pad=0.25)
    fig.savefig(OUT_PATH, dpi=300, bbox_inches="tight", facecolor="#FFFFFF")
    plt.close(fig)
    print(f"Flowchart saved successfully: {OUT_PATH}")


if __name__ == "__main__":
    generate_flowchart()
