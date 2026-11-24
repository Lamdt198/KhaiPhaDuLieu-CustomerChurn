# -*- coding: utf-8 -*-
"""
build_presentation.py
Sinh bộ Slide thuyết trình 21 trang áp dụng chuẩn TemplateThemeEngine (Theme2) từ hệ thống HUBT K20.2:
- Canvas 20.00 x 11.25 in nền sáng thống nhất (#FFFFFF)
- Bảng màu học thuật Slate/Ice (#202A31, #707F90, #9EADBD, #D0DDED, #F2F5F8) + màu ngữ nghĩa hài hòa
- Thanh điều hướng Top Pill Navigation Bar 5 mục, góc nhấn hình học, thẻ bo góc, bảng chuẩn fixed-width
- 100% nội dung và số liệu thực nghiệm môn Khai phá dữ liệu (Nhóm 02 - GVHD: TS. Bùi Ngọc Thăng)
- Đầy đủ 21/21 Speaker Notes, không dùng emoji.
"""

import json
import os
import sys
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.enum.shapes import MSO_SHAPE

if hasattr(sys.stdout, "reconfigure"):
    try:
        sys.stdout.reconfigure(encoding="utf-8")
    except Exception:
        pass

# ==============================================================================
# 1. DESIGN TOKENS — ĐỒNG BỘ 100% VỚI THEME2 (TemplateThemeEngine)
# ==============================================================================
BASE_DIR = os.path.dirname(os.path.abspath(__file__))
ASSET_DIR = os.path.join(BASE_DIR, "assets")
MODEL_DIR = os.path.join(BASE_DIR, "models")
REPORT_DIR = os.path.join(BASE_DIR, "reports")
FIG_DIR = os.path.join(REPORT_DIR, "figures")
LOGO_PATH = os.path.join(ASSET_DIR, "logo_hubt_transparent.png")
OUTPUT_PATH = os.path.join(REPORT_DIR, "Slide_Thuyet_Trinh_Khai_Pha_Du_Lieu_Nhom02.pptx")

SLIDE_W = Inches(20.00)
SLIDE_H = Inches(11.25)

C_DARK_BG   = RGBColor(0x20, 0x2A, 0x31)  # #202A31 - Chữ chính / Active Pill / Header bảng
C_LIGHT_BG  = RGBColor(0xFF, 0xFF, 0xFF)  # #FFFFFF - Nền sáng thống nhất
C_STEEL     = RGBColor(0x70, 0x7F, 0x90)  # #707F90 - Badge / Viền nhấn / Chữ phụ đậm
C_SLATE     = RGBColor(0x9E, 0xAD, 0xBD)  # #9EADBD - Accent phụ
C_ICE       = RGBColor(0xD0, 0xDD, 0xED)  # #D0DDED - Top bar / Viền thẻ / Section pill
C_SURFACE   = RGBColor(0xF2, 0xF5, 0xF8)  # #F2F5F8 - Bề mặt thẻ trên nền sáng
C_MUTED_TXT = RGBColor(0x52, 0x60, 0x6D)  # #52606D - Chữ mô tả phụ

# Màu ngữ nghĩa hài hòa với tone Slate
C_DANGER    = RGBColor(0xC9, 0x4A, 0x4A)  # Rủi ro cao / Churn
C_WARN      = RGBColor(0xD9, 0x9B, 0x26)  # Cần chú ý / Ngưỡng
C_SAFE      = RGBColor(0x2B, 0x8A, 0x78)  # An toàn / Tối ưu / Highlight
C_SAFE_BG   = RGBColor(0xE1, 0xF3, 0xEE)  # Nền xanh nhạt dòng highlight bảng
C_HIGHLIGHT = RGBColor(0x3A, 0x7C, 0xA5)  # Điểm nhấn học thuật

FONT_MAIN   = "Arial"   # Tương đương kích thước 1:1 với Arimo, cài sẵn trên mọi Windows

NAV_ITEMS = [
    "Mở đầu & Ch.1",
    "Ch.2 EDA & Tiền xử lý",
    "Ch.3 Mô hình & T*",
    "Ch.4 Tri thức & Lỗi",
    "Ch.5 Demo & Kết luận",
]

TOTAL_SLIDES = 21


# ==============================================================================
# 2. TEMPLATE THEME ENGINE (CHUẨN THEME2)
# ==============================================================================
class TemplateThemeEngine:
    def __init__(self):
        self.prs = Presentation()
        self.prs.slide_width = SLIDE_W
        self.prs.slide_height = SLIDE_H
        self.blank_layout = self.prs.slide_layouts[6]
        self.slide_num = 0

    def new_slide(self, nav_idx=None, section_tag="", title="", subtitle="", notes=""):
        self.slide_num += 1
        slide = self.prs.slides.add_slide(self.blank_layout)

        bg = slide.background
        fill = bg.fill
        fill.solid()
        fill.fore_color.rgb = C_LIGHT_BG

        self._draw_corner_accents(slide)

        if nav_idx is not None:
            self._draw_top_nav(slide, nav_idx)

        if title:
            self._draw_slide_header(slide, section_tag, title, subtitle)

        if nav_idx is not None:
            self._draw_footer(slide)

        if notes:
            slide.notes_slide.notes_text_frame.text = notes.strip()

        return slide

    def _draw_corner_accents(self, slide):
        sh1 = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(0), Inches(0), Inches(20.0), Inches(0.10))
        sh1.fill.solid()
        sh1.fill.fore_color.rgb = C_ICE
        sh1.line.fill.background()

        sh2 = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(0), Inches(11.15), Inches(5.5), Inches(0.10))
        sh2.fill.solid()
        sh2.fill.fore_color.rgb = C_STEEL
        sh2.line.fill.background()

    def _draw_top_nav(self, slide, active_idx):
        if os.path.exists(LOGO_PATH):
            slide.shapes.add_picture(LOGO_PATH, Inches(0.65), Inches(0.36), height=Inches(0.44))

        tb_l = slide.shapes.add_textbox(Inches(1.48), Inches(0.38), Inches(4.15), Inches(0.45))
        tf_l = tb_l.text_frame
        tf_l.word_wrap = True
        p = tf_l.paragraphs[0]
        r = p.add_run()
        r.text = "KHAI PHÁ DỮ LIỆU • CUSTOMER CHURN"
        r.font.name = FONT_MAIN
        r.font.size = Pt(12)
        r.font.bold = True
        r.font.color.rgb = C_DARK_BG

        pill_w = 2.04
        pill_h = 0.44
        gap = 0.12
        start_x = 5.72
        y_nav = 0.36
        for idx, label in enumerate(NAV_ITEMS):
            x = start_x + idx * (pill_w + gap)
            is_active = (idx == active_idx)
            pill = slide.shapes.add_shape(
                MSO_SHAPE.ROUNDED_RECTANGLE,
                Inches(x), Inches(y_nav), Inches(pill_w), Inches(pill_h)
            )
            pill.adjustments[0] = 0.5
            if is_active:
                pill.fill.solid()
                pill.fill.fore_color.rgb = C_DARK_BG
                pill.line.fill.background()
                txt_color = C_LIGHT_BG
            else:
                pill.fill.solid()
                pill.fill.fore_color.rgb = C_SURFACE
                pill.line.color.rgb = C_ICE
                pill.line.width = Pt(1.0)
                txt_color = C_STEEL

            tf = pill.text_frame
            tf.margin_left = Inches(0.05)
            tf.margin_right = Inches(0.05)
            tf.margin_top = Inches(0.02)
            tf.margin_bottom = Inches(0.02)
            tf.vertical_anchor = MSO_ANCHOR.MIDDLE
            p = tf.paragraphs[0]
            p.alignment = PP_ALIGN.CENTER
            r = p.add_run()
            r.text = label
            r.font.name = FONT_MAIN
            r.font.size = Pt(10.5)
            r.font.bold = is_active
            r.font.color.rgb = txt_color

        tb_r = slide.shapes.add_textbox(Inches(16.65), Inches(0.38), Inches(2.70), Inches(0.45))
        p = tb_r.text_frame.paragraphs[0]
        p.alignment = PP_ALIGN.RIGHT
        r = p.add_run()
        r.text = "HUBT • K20.2 • NHÓM 02"
        r.font.name = FONT_MAIN
        r.font.size = Pt(12)
        r.font.bold = True
        r.font.color.rgb = C_STEEL

        line = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(0.65), Inches(0.96), Inches(18.70), Inches(0.02))
        line.fill.solid()
        line.fill.fore_color.rgb = C_ICE
        line.line.fill.background()

    def _draw_slide_header(self, slide, section_tag, title, subtitle):
        if section_tag:
            badge = slide.shapes.add_shape(
                MSO_SHAPE.ROUNDED_RECTANGLE, Inches(0.65), Inches(1.14), Inches(3.25), Inches(0.36)
            )
            badge.adjustments[0] = 0.5
            badge.fill.solid()
            badge.fill.fore_color.rgb = C_ICE
            badge.line.fill.background()
            tf = badge.text_frame
            tf.vertical_anchor = MSO_ANCHOR.MIDDLE
            p = tf.paragraphs[0]
            p.alignment = PP_ALIGN.CENTER
            r = p.add_run()
            r.text = section_tag.upper()
            r.font.name = FONT_MAIN
            r.font.size = Pt(10.5)
            r.font.bold = True
            r.font.color.rgb = C_DARK_BG

        tb = slide.shapes.add_textbox(Inches(0.65), Inches(1.56), Inches(18.70), Inches(0.62))
        tf = tb.text_frame
        tf.word_wrap = True
        p = tf.paragraphs[0]
        r = p.add_run()
        r.text = title
        r.font.name = FONT_MAIN
        r.font.size = Pt(26.5)
        r.font.bold = True
        r.font.color.rgb = C_DARK_BG

        if subtitle:
            tb_s = slide.shapes.add_textbox(Inches(0.65), Inches(2.16), Inches(18.70), Inches(0.42))
            tf_s = tb_s.text_frame
            tf_s.word_wrap = True
            p = tf_s.paragraphs[0]
            r = p.add_run()
            r.text = subtitle
            r.font.name = FONT_MAIN
            r.font.size = Pt(13.8)
            r.font.color.rgb = C_MUTED_TXT

    def _draw_footer(self, slide):
        tb = slide.shapes.add_textbox(Inches(0.65), Inches(10.68), Inches(14.0), Inches(0.35))
        p = tb.text_frame.paragraphs[0]
        r = p.add_run()
        r.text = "Bài tập lớn Học phần Khai phá dữ liệu • Lớp Cao học CNTT K20.2 – GVHD: TS. Bùi Ngọc Thăng"
        r.font.name = FONT_MAIN
        r.font.size = Pt(10.5)
        r.font.color.rgb = C_STEEL

        tb_n = slide.shapes.add_textbox(Inches(17.50), Inches(10.68), Inches(1.85), Inches(0.35))
        p = tb_n.text_frame.paragraphs[0]
        p.alignment = PP_ALIGN.RIGHT
        r = p.add_run()
        r.text = f"{self.slide_num:02d} / {TOTAL_SLIDES:02d}"
        r.font.name = FONT_MAIN
        r.font.size = Pt(11.5)
        r.font.bold = True
        r.font.color.rgb = C_DARK_BG

    def add_card(self, slide, x, y, w, h, fill_color=None, border_color=None, radius=0.06):
        card = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(x), Inches(y), Inches(w), Inches(h))
        card.adjustments[0] = radius
        card.fill.solid()
        card.fill.fore_color.rgb = fill_color if fill_color else C_SURFACE
        card.line.color.rgb = border_color if border_color else C_ICE
        card.line.width = Pt(1.25 if border_color else 1.0)
        return card

    def add_numbered_badge(self, slide, x, y, text, fill_color=C_DARK_BG, w=0.85, h=0.60, font_size=18):
        sh = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(x), Inches(y), Inches(w), Inches(h))
        sh.adjustments[0] = 0.25
        sh.fill.solid()
        sh.fill.fore_color.rgb = fill_color
        sh.line.fill.background()
        tf = sh.text_frame
        tf.vertical_anchor = MSO_ANCHOR.MIDDLE
        p = tf.paragraphs[0]
        p.alignment = PP_ALIGN.CENTER
        r = p.add_run()
        r.text = text
        r.font.name = FONT_MAIN
        r.font.size = Pt(font_size)
        r.font.bold = True
        r.font.color.rgb = C_LIGHT_BG
        return sh

    def add_text(self, slide, x, y, w, h, text, size=14.5, bold=False, color=C_DARK_BG, align=PP_ALIGN.LEFT):
        tb = slide.shapes.add_textbox(Inches(x), Inches(y), Inches(w), Inches(h))
        tf = tb.text_frame
        tf.word_wrap = True
        tf.margin_left = Inches(0.04)
        tf.margin_right = Inches(0.04)
        tf.margin_top = Inches(0.02)
        tf.margin_bottom = Inches(0.02)
        p = tf.paragraphs[0]
        p.alignment = align
        r = p.add_run()
        r.text = text
        r.font.name = FONT_MAIN
        r.font.size = Pt(size)
        r.font.bold = bold
        r.font.color.rgb = color
        return tb

    def add_bullets(self, slide, x, y, w, h, items, font_size=13.8, space_after=8, bullet_char="▪ "):
        tb = slide.shapes.add_textbox(Inches(x), Inches(y), Inches(w), Inches(h))
        tf = tb.text_frame
        tf.word_wrap = True
        tf.margin_left = Inches(0.05)
        tf.margin_right = Inches(0.05)
        tf.margin_top = Inches(0.04)
        tf.margin_bottom = Inches(0.04)
        for idx, item in enumerate(items):
            p = tf.paragraphs[0] if idx == 0 else tf.add_paragraph()
            p.space_after = Pt(space_after)
            p.line_spacing = 1.15
            if isinstance(item, tuple):
                lead, body = item
                r1 = p.add_run()
                r1.text = bullet_char + lead + " "
                r1.font.name = FONT_MAIN
                r1.font.size = Pt(font_size)
                r1.font.bold = True
                r1.font.color.rgb = C_DARK_BG

                r2 = p.add_run()
                r2.text = body
                r2.font.name = FONT_MAIN
                r2.font.size = Pt(font_size)
                r2.font.color.rgb = C_MUTED_TXT
            else:
                r = p.add_run()
                r.text = bullet_char + item
                r.font.name = FONT_MAIN
                r.font.size = Pt(font_size)
                r.font.color.rgb = C_DARK_BG
        return tb

    def add_kpi_card(self, slide, x, y, w, h, num_str, title_str, sub_str="", accent=C_HIGHLIGHT):
        self.add_card(slide, x, y, w, h, fill_color=C_SURFACE, border_color=accent, radius=0.08)
        bar = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(x), Inches(y), Inches(w), Inches(0.08))
        bar.fill.solid()
        bar.fill.fore_color.rgb = accent
        bar.line.fill.background()
        self.add_text(slide, x + 0.15, y + 0.22, w - 0.30, 0.72, num_str, size=34, bold=True, color=accent, align=PP_ALIGN.CENTER)
        self.add_text(slide, x + 0.15, y + 0.98, w - 0.30, 0.38, title_str, size=14.5, bold=True, color=C_DARK_BG, align=PP_ALIGN.CENTER)
        if sub_str:
            self.add_text(slide, x + 0.15, y + 1.38, w - 0.30, 0.36, sub_str, size=12, color=C_MUTED_TXT, align=PP_ALIGN.CENTER)

    def add_framed_image(self, slide, img_path, x, y, w, h, caption=""):
        self.add_card(slide, x, y, w, h, fill_color=C_LIGHT_BG, border_color=C_ICE, radius=0.04)
        if os.path.exists(img_path):
            pad = 0.14
            cap_h = 0.38 if caption else 0.0
            avail_w = w - 2 * pad
            avail_h = h - 2 * pad - cap_h
            pic = slide.shapes.add_picture(img_path, Inches(x + pad), Inches(y + pad))
            ratio = min(Inches(avail_w) / pic.width, Inches(avail_h) / pic.height)
            pic.width = int(pic.width * ratio)
            pic.height = int(pic.height * ratio)
            pic.left = int(Inches(x + pad) + (Inches(avail_w) - pic.width) / 2)
            pic.top = int(Inches(y + pad) + (Inches(avail_h) - pic.height) / 2)
        if caption:
            self.add_text(slide, x + 0.15, y + h - 0.40, w - 0.30, 0.34, caption,
                          size=11.5, bold=True, color=C_STEEL, align=PP_ALIGN.CENTER)

    def add_styled_table(self, slide, x, y, w, h, rows_data, col_widths=None, font_size=12.0,
                         highlight_row=None, left_align_cols=(0,)):
        n_rows = len(rows_data)
        n_cols = len(rows_data[0])
        tbl_shape = slide.shapes.add_table(n_rows, n_cols, Inches(x), Inches(y), Inches(w), Inches(h))
        tbl = tbl_shape.table
        if col_widths:
            scale = w / sum(col_widths)
            for c_idx, cw in enumerate(col_widths):
                tbl.columns[c_idx].width = Inches(cw * scale)

        for r_idx, row in enumerate(rows_data):
            for c_idx, val in enumerate(row):
                cell = tbl.cell(r_idx, c_idx)
                cell.vertical_anchor = MSO_ANCHOR.MIDDLE
                cell.margin_left = Inches(0.08)
                cell.margin_right = Inches(0.08)
                cell.margin_top = Inches(0.04)
                cell.margin_bottom = Inches(0.04)
                cell.fill.solid()

                if r_idx == 0:
                    cell.fill.fore_color.rgb = C_DARK_BG
                    txt_c = C_LIGHT_BG
                    is_b = True
                elif highlight_row is not None and r_idx == highlight_row:
                    cell.fill.fore_color.rgb = C_SAFE_BG
                    txt_c = C_DARK_BG
                    is_b = True
                else:
                    cell.fill.fore_color.rgb = C_LIGHT_BG if r_idx % 2 == 1 else C_SURFACE
                    txt_c = C_DARK_BG
                    is_b = (c_idx == 0)

                tf = cell.text_frame
                tf.word_wrap = True
                p = tf.paragraphs[0]
                if r_idx == 0:
                    p.alignment = PP_ALIGN.CENTER
                else:
                    p.alignment = PP_ALIGN.LEFT if c_idx in left_align_cols else PP_ALIGN.CENTER
                r = p.add_run()
                r.text = str(val)
                r.font.name = FONT_MAIN
                r.font.size = Pt(font_size)
                r.font.bold = is_b
                r.font.color.rgb = txt_c
        return tbl_shape

    def save(self, path):
        self.prs.save(path)


# ==============================================================================
# 3. XÂY DỰNG ĐẦY ĐỦ 21 SLIDES — MÔN KHAI PHÁ DỮ LIỆU (NHÓM 02)
# ==============================================================================
def build_all_slides():
    with open(os.path.join(MODEL_DIR, "eda_stats.json"), "r", encoding="utf-8") as f:
        eda = json.load(f)
    with open(os.path.join(MODEL_DIR, "experiment_results.json"), "r", encoding="utf-8") as f:
        res = json.load(f)

    t_star = res["optimal_threshold"]
    roi = res["roi_benchmark"]
    errt = res["error_analysis"]["at_optimal_tstar"]
    ab_unw = res["ablation_study"]["Unweighted (T=0.50)"]
    ab_bal = res["ablation_study"]["Class-Weighted Balanced (T=0.50)"]
    ab_opt = res["ablation_study"][f"Cost-Sensitive Optimal (T*={t_star:.2f})"]
    m_order = ["Logistic Regression", "Decision Tree (CART)", "Random Forest", "XGBoost", "LightGBM"]

    eng = TemplateThemeEngine()

    # --------------------------------------------------------------------------
    # SLIDE 01: COVER (Trang bìa chính thức chuẩn Theme2)
    # --------------------------------------------------------------------------
    s = eng.new_slide(
        nav_idx=None,
        notes="Kính chào TS. Bùi Ngọc Thăng và các anh chị học viên lớp Cao học CNTT K20.2. Thay mặt Nhóm 02, tôi xin trình bày Bài tập lớn học phần Khai phá dữ liệu với đề tài: Ứng dụng các mô hình phân loại dữ liệu trong dự đoán nguy cơ rời bỏ dịch vụ của khách hàng."
    )
    if os.path.exists(LOGO_PATH):
        pic_logo = s.shapes.add_picture(LOGO_PATH, Inches(9.35), Inches(0.60), height=Inches(1.15))
        pic_logo.left = int((SLIDE_W - pic_logo.width) / 2)

    eng.add_text(s, 1.5, 1.90, 17.0, 0.45, "TRƯỜNG ĐẠI HỌC KINH DOANH VÀ CÔNG NGHỆ HÀ NỘI", size=18, bold=True, color=C_DARK_BG, align=PP_ALIGN.CENTER)
    eng.add_text(s, 1.5, 2.38, 17.0, 0.40, "VIỆN ĐÀO TẠO SAU ĐẠI HỌC — LỚP CAO HỌC CNTT K20.2", size=14.5, color=C_STEEL, align=PP_ALIGN.CENTER)

    eng.add_card(s, 6.2, 3.08, 7.6, 0.50, fill_color=C_ICE, border_color=C_STEEL, radius=0.5)
    eng.add_text(s, 6.3, 3.16, 7.4, 0.34, "BÁO CÁO BÀI TẬP LỚN — CHỦ ĐỀ 2: PHÂN LOẠI DỮ LIỆU", size=13.5, bold=True, color=C_DARK_BG, align=PP_ALIGN.CENTER)

    eng.add_text(
        s, 0.80, 3.82, 18.40, 2.10,
        "ỨNG DỤNG CÁC MÔ HÌNH PHÂN LOẠI DỮ LIỆU\nTRONG DỰ ĐOÁN NGUY CƠ RỜI BỎ DỊCH VỤ CỦA KHÁCH HÀNG",
        size=37, bold=True, color=C_DARK_BG, align=PP_ALIGN.CENTER
    )
    eng.add_text(
        s, 1.00, 6.08, 18.00, 0.85,
        "Tuân thủ quy trình KDD chống rò rỉ dữ liệu, đối sánh 5 thuật toán phân loại, tối ưu ngưỡng nhạy chi phí\nvà rút trích tri thức (Luật CART & SHAP) trên bộ dữ liệu IBM Telco Customer Churn (7.043 thuê bao)",
        size=17.5, color=C_MUTED_TXT, align=PP_ALIGN.CENTER
    )

    eng.add_card(s, 0.80, 7.28, 18.40, 3.15, fill_color=C_SURFACE, border_color=C_ICE, radius=0.05)
    eng.add_bullets(s, 1.25, 7.56, 17.50, 2.65, [
        ("Học phần & Giảng viên:", "Khai phá dữ liệu (Data Mining) — Giảng viên hướng dẫn: TS. Bùi Ngọc Thăng"),
        ("Nhóm thực hiện (Nhóm 02):", "Đỗ Ngọc Điền • Dương Tùng Lâm • Trần Thị Thu Hương • Phạm Văn Chiến • Vàng A Hứ"),
        ("Sản phẩm bàn giao:", "Báo cáo Word 5 chương (.docx) • Bộ Slide 21 trang • Ứng dụng Web Streamlit (chạy 100% cục bộ) • Mã nguồn GitHub"),
    ], font_size=16.0, space_after=14)

    # --------------------------------------------------------------------------
    # SLIDE 02: ĐỘI NGŨ & PHÂN CÔNG NHIỆM VỤ (5 Vertical Cards chuẩn Theme2 Slide 02)
    # --------------------------------------------------------------------------
    s = eng.new_slide(
        nav_idx=0, section_tag="Mở đầu • Phân công nhiệm vụ",
        title="Đội Ngũ Thực Hiện Bài Tập Lớn (Nhóm 02 – Lớp Cao học CNTT K20.2)",
        subtitle="Phân công công việc rõ ràng theo 5 mảng cốt lõi của quy trình Khai phá dữ liệu — Tỷ trọng đóng góp đồng đều 20%/thành viên",
        notes="Bài tập lớn được phân công rõ ràng cho 5 thành viên Nhóm 02 theo đúng mạch quy trình KDD, đảm bảo tính đồng nhất giữa mã nguồn thực nghiệm, báo cáo Word, ứng dụng demo và slide bảo vệ."
    )
    members = [
        ("Đỗ Ngọc Điền", "KHUNG KDD & TỔNG HỢP", "Xây dựng khung đề tài theo quy trình KDD/CRISP-DM, tổng hợp cơ sở lý thuyết Chương 1 và Kết luận, rà soát thống nhất toàn bộ báo cáo."),
        ("Dương Tùng Lâm", "EDA & ĐẶC TRƯNG MIỀN", "Phân tích khám phá bộ dữ liệu 7.043 thuê bao, chạy kiểm định Chi-Square & Mann-Whitney U, thiết kế 6 đặc trưng miền và viết Chương 2."),
        ("Trần Thị Thu Hương", "MÔ HÌNH & 5-FOLD CV", "Xây dựng Anti-Leakage Pipeline, huấn luyện và tinh chỉnh 5 mô hình phân loại qua Stratified 5-Fold GridSearchCV, viết Chương 3."),
        ("Phạm Văn Chiến", "CHI PHÍ, LUẬT & XAI", "Tối ưu ngưỡng nhạy chi phí T* trên OOF Train, rút trích 8 luật CART, giải thích SHAP toàn cục/cục bộ và phân tích lỗi FN/FP (Chương 4)."),
        ("Vàng A Hứ", "ROI, DEMO & SLIDE", "Thực nghiệm đối chứng chi phí 4 kịch bản vận hành, xây dựng ứng dụng minh họa Streamlit chạy cục bộ, viết Chương 5 và biên tập Slide."),
    ]
    card_w = 3.56
    gap = 0.22
    start_x = 0.65
    for idx, (m_name, m_role, m_desc) in enumerate(members):
        cx = start_x + idx * (card_w + gap)
        eng.add_card(s, cx, 2.85, card_w, 7.35, fill_color=C_SURFACE, border_color=C_STEEL, radius=0.06)
        eng.add_numbered_badge(s, cx + 0.22, 3.12, f"0{idx+1}", fill_color=C_STEEL, w=0.72, h=0.55)
        eng.add_card(s, cx + 2.25, 3.16, 1.08, 0.44, fill_color=C_SAFE, border_color=C_SAFE, radius=0.5)
        eng.add_text(s, cx + 2.25, 3.22, 1.08, 0.34, "20%", size=12, bold=True, color=C_LIGHT_BG, align=PP_ALIGN.CENTER)

        eng.add_text(s, cx + 0.22, 3.92, card_w - 0.44, 0.50, m_name, size=18, bold=True, color=C_DARK_BG)
        eng.add_text(s, cx + 0.22, 4.44, card_w - 0.44, 0.38, m_role, size=11.5, bold=True, color=C_HIGHLIGHT)

        d = s.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(cx + 0.22), Inches(4.92), Inches(card_w - 0.44), Inches(0.02))
        d.fill.solid()
        d.fill.fore_color.rgb = C_STEEL
        d.line.fill.background()

        eng.add_text(s, cx + 0.22, 5.18, card_w - 0.44, 4.65, m_desc, size=13.5, color=C_MUTED_TXT)

    # --------------------------------------------------------------------------
    # SLIDE 03: MENU / KHUNG BÁO CÁO 5 CHƯƠNG (5 Stacked Rows chuẩn Theme2 Slide 03)
    # --------------------------------------------------------------------------
    s = eng.new_slide(
        nav_idx=0, section_tag="Mở đầu • Lộ trình trình bày",
        title="Khung Báo Cáo 5 Chương & Lộ Trình Thuyết Trình Đề Tài",
        subtitle="Bám sát đầy đủ quy trình Khai phá tri thức từ dữ liệu (KDD) cho bài toán phân loại rời bỏ khách hàng",
        notes="Báo cáo gồm 5 chương bám sát quy trình KDD: Chương 1 tổng quan và cơ sở toán học, Chương 2 khám phá dữ liệu và tiền xử lý chống rò rỉ, Chương 3 đối sánh 5 mô hình và tối ưu ngưỡng chi phí, Chương 4 rút trích luật CART, SHAP, phân tích lỗi và ROI, Chương 5 minh họa ứng dụng và đề xuất hành động CSKH."
    )
    chapters = [
        ("CHƯƠNG 1", "Tổng quan bài toán & Cơ sở lý thuyết Khai phá dữ liệu", "Bối cảnh Customer Churn, 3 điểm nghẽn CSKH truyền thống, quy trình KDD/CRISP-DM, cơ sở toán học 5 thuật toán phân loại, học nhạy chi phí & SHAP."),
        ("CHƯƠNG 2", "Phân tích dữ liệu (EDA), Kiểm định thống kê & Tiền xử lý", "Đặc tả 7.043 thuê bao IBM Telco, kiểm định Chi-Square & Mann-Whitney U, tạo 6 đặc trưng miền và Anti-Leakage Pipeline 4 bước."),
        ("CHƯƠNG 3", "Thực nghiệm 5 mô hình phân loại & Tối ưu ngưỡng chi phí", "Đối sánh Logistic Regression, CART, Random Forest, XGBoost, LightGBM qua 5-Fold CV; Ablation Study lệch lớp & tối ưu ngưỡng T* = 0,21."),
        ("CHƯƠNG 4", "Rút trích tri thức (CART & SHAP), Phân tích lỗi & Hiệu quả ROI", "Hệ 8 luật IF-THEN từ cây quyết định, SHAP toàn cục/cục bộ, giải phẫu lỗi FN/FP trên tập Test và đối chứng chi phí 4 chiến lược chăm sóc."),
        ("CHƯƠNG 5", "Triển khai ứng dụng minh họa Streamlit & Đề xuất giải pháp CSKH", "Nguyên mẫu ứng dụng web 3 phân hệ chạy 100% cục bộ (không API key), 4 nhóm hành động giữ chân khách hàng và hướng phát triển."),
    ]
    for i, (ch, ctitle, cdesc) in enumerate(chapters):
        cy = 2.75 + i * 1.50
        eng.add_card(s, 0.65, cy, 18.70, 1.32, fill_color=C_SURFACE if i % 2 == 0 else C_LIGHT_BG, border_color=C_DARK_BG, radius=0.08)
        eng.add_numbered_badge(s, 0.88, cy + 0.30, ch, fill_color=C_DARK_BG if i % 2 == 0 else C_STEEL, w=2.05, h=0.72, font_size=16)
        eng.add_text(s, 3.15, cy + 0.16, 15.8, 0.42, ctitle, size=17.5, bold=True, color=C_DARK_BG)
        eng.add_text(s, 3.15, cy + 0.62, 15.8, 0.60, cdesc, size=13.5, color=C_MUTED_TXT)

    # --------------------------------------------------------------------------
    # SLIDE 04: CHƯƠNG 1 — ĐẶT VẤN ĐỀ & BA ĐIỂM NGHẼN CSKH TRUYỀN THỐNG
    # --------------------------------------------------------------------------
    s = eng.new_slide(
        nav_idx=0, section_tag="Chương 1 • Mục 1.1",
        title="1.1 Đặt Vấn Đề: Bài Toán Dự Đoán Khách Hàng Rời Bỏ Dịch Vụ",
        subtitle="Chi phí thu hút thuê bao mới cao gấp 5–7 lần chi phí giữ chân thuê bao hiện hữu trong ngành viễn thông",
        notes="Trên bộ dữ liệu IBM Telco, hơn một phần tư thuê bao (26,54%) chấm dứt dịch vụ. Phương pháp chăm sóc truyền thống vướng ba điểm nghẽn: phản ứng bị động, lọc thủ công bỏ sót nhiều và khuyến mãi dàn trải gây lãng phí ngân sách."
    )
    eng.add_kpi_card(s, 0.65, 2.75, 4.45, 1.90, "7.043", "Thuê bao thực tế", "IBM Telco Customer Churn", C_DARK_BG)
    eng.add_kpi_card(s, 5.40, 2.75, 4.45, 1.90, "26,54%", "Tỷ lệ rời bỏ (Churn)", "1.869 Churn / 5.174 Ở lại", C_DANGER)
    eng.add_kpi_card(s, 10.15, 2.75, 4.45, 1.90, "5 : 1", "Bất đối xứng chi phí", "Bỏ sót $500 vs Ưu đãi $100", C_WARN)
    eng.add_kpi_card(s, 14.90, 2.75, 4.45, 1.90, "5 Mô hình", "Đối sánh có đối chứng", "2 Baselines + 3 Ensembles", C_SAFE)

    bottlenecks = [
        ("01.", "Phản Ứng Bị Động", "Doanh nghiệp thường chỉ biết khi khách hàng đã yêu cầu hủy hợp đồng; khi đó nỗ lực giữ chân tốn kém và tỷ lệ thành công rất thấp.", C_DANGER),
        ("02.", "Luật Lọc Thủ Công Bỏ Sót Nhiều", "Các tiêu chí lọc đơn giản trên bảng tính (chỉ xét hợp đồng tháng và cước cao) bỏ sót phần lớn khách có nguy cơ rời bỏ do thiếu dịch vụ hỗ trợ.", C_WARN),
        ("03.", "Khuyến Mãi Dàn Trải Lãng Phí", "Gửi ưu đãi đồng loạt cho 100% tập khách hàng làm đội chi phí vào nhóm vốn đã ổn định (hợp đồng dài hạn, thời gian gắn bó cao).", C_STEEL),
    ]
    for i, (num, btitle, bdesc, col) in enumerate(bottlenecks):
        bx = 0.65 + i * 6.32
        eng.add_card(s, bx, 4.95, 6.06, 3.20, fill_color=C_SURFACE, border_color=col)
        eng.add_numbered_badge(s, bx + 0.22, 5.18, num, fill_color=col, w=0.78, h=0.55)
        eng.add_text(s, bx + 1.15, 5.26, 4.70, 0.45, btitle, size=16.5, bold=True, color=C_DARK_BG)
        eng.add_text(s, bx + 0.22, 5.95, 5.62, 2.00, bdesc, size=13.5, color=C_MUTED_TXT)

    eng.add_card(s, 0.65, 8.45, 18.70, 1.85, fill_color=C_LIGHT_BG, border_color=C_HIGHLIGHT)
    eng.add_text(s, 0.95, 8.65, 18.10, 0.38, "MỤC TIÊU CỐT LÕI CỦA ĐỀ TÀI (TRỌNG TÂM KHAI PHÁ DỮ LIỆU)", size=14.5, bold=True, color=C_HIGHLIGHT)
    eng.add_bullets(s, 0.95, 9.08, 18.10, 1.10, [
        ("Quy trình KDD chặt chẽ & Tối ưu chi phí:", "Xây dựng pipeline chống rò rỉ dữ liệu, đối sánh 5 thuật toán phân loại, tìm ngưỡng quyết định nhạy chi phí T* và rút trích tri thức giải thích được (CART + SHAP) chạy 100% cục bộ."),
    ], font_size=13.5, space_after=4)

    # --------------------------------------------------------------------------
    # SLIDE 05: CHƯƠNG 1 — QUY TRÌNH KDD & KHUNG GIẢI PHÁP 3 TRỤ CỘT
    # --------------------------------------------------------------------------
    s = eng.new_slide(
        nav_idx=0, section_tag="Chương 1 • Mục 1.2",
        title="1.2 Quy Trình Khai Phá Dữ Liệu (KDD / CRISP-DM) & Khung 3 Trụ Cột",
        subtitle="Ánh xạ bài toán thực tế sang phân loại nhị phân có giám sát và tổ chức thực nghiệm khép kín không rò rỉ",
        notes="Quy trình thực hiện tuân thủ 5 giai đoạn KDD và 3 trụ cột: xử lý dữ liệu kèm kiểm định thống kê, mô hình hóa phân loại nhạy chi phí và rút trích tri thức giải thích được."
    )
    kdd_steps = [
        ("01", "Dữ liệu & Làm sạch", "7.043 mẫu, 21 cột; xử lý 11 dòng thiếu TotalCharges.", C_DARK_BG),
        ("02", "EDA & 6 Đặc trưng", "Chi-Square, Mann-Whitney U; mở rộng 25 biến (55 chiều).", C_STEEL),
        ("03", "5 Mô hình & 5-Fold", "Stratified 80/20; GridSearchCV trên ROC-AUC.", C_HIGHLIGHT),
        ("04", "Chi phí & Ngưỡng T*", "Ma trận phạt 5:1; tối ưu T* = 0,21 trên OOF Train.", C_WARN),
        ("05", "Tri thức & Demo", "8 luật CART, SHAP Global/Local, Error Analysis & App.", C_SAFE),
    ]
    for i, (num_s, st_t, st_d, col_s) in enumerate(kdd_steps):
        sx = 0.65 + i * 3.78
        eng.add_card(s, sx, 2.75, 3.55, 2.50, fill_color=C_SURFACE, border_color=col_s)
        eng.add_numbered_badge(s, sx + 0.18, 2.95, num_s, fill_color=col_s, w=0.72, h=0.52, font_size=16)
        eng.add_text(s, sx + 1.02, 3.04, 2.38, 0.45, st_t, size=14.2, bold=True, color=C_DARK_BG)
        eng.add_text(s, sx + 0.18, 3.68, 3.20, 1.40, st_d, size=12.8, color=C_MUTED_TXT)

    pillars = [
        ("TRỤ CỘT 1", "Dữ Liệu & Kiểm Định Thống Kê", [
            "Kiểm định phi tham số Chi-Square, Cramér's V và Mann-Whitney U trước khi mô hình hóa.",
            "Tạo 6 đặc trưng miền nắm bắt vòng đời (TenureGroup) và mức chi tiêu (AvgMonthlySpend).",
            "Cô lập hoàn toàn 1.409 hồ sơ tập Test trước mọi phép tính thống kê.",
        ], C_STEEL),
        ("TRỤ CỘT 2", "Phân Loại Nhạy Chi Phí", [
            "Đối sánh 2 Baselines (Logistic Regression, CART) với 3 Ensembles (RF, XGBoost, LightGBM).",
            "So sánh thực nghiệm Unweighted vs Balanced Weight vs Threshold Moving.",
            "Bảo toàn hiệu chuẩn xác suất (Brier Score) trong khi giảm 49,6% tổng chi phí.",
        ], C_HIGHLIGHT),
        ("TRỤ CỘT 3", "Rút Trích Tri Thức & Hành Động", [
            "Chuyển hóa mô hình thành 8 luật IF-THEN trực quan từ cây CART cắt tỉa độ sâu 3.",
            "Minh bạch hóa LightGBM bằng TreeSHAP toàn cục và cục bộ cho từng thuê bao.",
            "Phân tích nguyên nhân lỗi FN/FP và đóng gói nguyên mẫu Web Streamlit cục bộ.",
        ], C_SAFE),
    ]
    for i, (ptag, ptitle, pitems, pcol) in enumerate(pillars):
        px = 0.65 + i * 6.32
        eng.add_card(s, px, 5.55, 6.06, 4.70, fill_color=C_LIGHT_BG, border_color=pcol)
        eng.add_numbered_badge(s, px + 0.22, 5.78, ptag, fill_color=pcol, w=1.75, h=0.48, font_size=13)
        eng.add_text(s, px + 0.22, 6.40, 5.62, 0.45, ptitle, size=16.5, bold=True, color=C_DARK_BG)
        eng.add_bullets(s, px + 0.22, 6.95, 5.62, 3.10, pitems, font_size=13.0, space_after=8)

    # --------------------------------------------------------------------------
    # SLIDE 06: CHƯƠNG 1 — CƠ SỞ TOÁN HỌC PHÂN LOẠI, CHI PHÍ & SHAP
    # --------------------------------------------------------------------------
    s = eng.new_slide(
        nav_idx=0, section_tag="Chương 1 • Mục 1.3 – 1.4",
        title="1.3 Cơ Sở Toán Học: Thuật Toán Phân Loại, Ma Trận Chi Phí & SHAP",
        subtitle="Nền tảng giải thuật của 5 mô hình thử nghiệm, ngưỡng quyết định tối ưu Bayes và giá trị Shapley cộng tính",
        notes="Nhóm đối sánh từ mô hình tuyến tính Logistic Regression và cây đơn CART tới các tổ hợp cây Random Forest, XGBoost, LightGBM. Với tỷ lệ chi phí bỏ sót so với chăm sóc là 5:1, ngưỡng tối ưu lý thuyết Bayes bằng 0,20."
    )
    eng.add_card(s, 0.65, 2.75, 9.15, 4.45, fill_color=C_SURFACE, border_color=C_STEEL)
    eng.add_numbered_badge(s, 0.90, 2.98, "01.", fill_color=C_DARK_BG, w=0.78, h=0.52)
    eng.add_text(s, 1.82, 3.06, 7.70, 0.45, "5 THUẬT TOÁN PHÂN LOẠI ĐỐI CHỨNG", size=16.5, bold=True, color=C_DARK_BG)
    eng.add_bullets(s, 0.90, 3.68, 8.65, 3.35, [
        ("Logistic Regression:", "Hàm Sigmoid P(Y=1|x) = 1 / (1 + e^-(w^T x + b)), tối ưu Binary Cross-Entropy có phạt L2; giải thích qua Odds Ratio = exp(w_j)."),
        ("Decision Tree (CART):", "Phân tách nhánh cực đại hóa độ giảm vấy bẩn Gini(D) = 1 - (p_0^2 + p_1^2); rút trích trực tiếp luật IF-THEN."),
        ("Random Forest:", "Bagging 250 cây trên mẫu bootstrap và tập con đặc trưng ngẫu nhiên nhằm giảm mạnh phương sai dự báo."),
        ("XGBoost & LightGBM:", "Gradient Boosting theo khai triển Taylor bậc 2 + phạt L1/L2; LightGBM tăng trưởng nút lá trên histogram tốc độ cao."),
    ], font_size=13.0, space_after=8)

    eng.add_card(s, 10.20, 2.75, 9.15, 4.45, fill_color=C_SURFACE, border_color=C_HIGHLIGHT)
    eng.add_numbered_badge(s, 10.45, 2.98, "02.", fill_color=C_HIGHLIGHT, w=0.78, h=0.52)
    eng.add_text(s, 11.37, 3.06, 7.70, 0.45, "NHẠY CHI PHÍ BAYES & GIÁ TRỊ SHAP", size=16.5, bold=True, color=C_DARK_BG)
    eng.add_bullets(s, 10.45, 3.68, 8.65, 3.35, [
        ("Điều kiện can thiệp tối ưu Bayes:", "p·C_TP + (1-p)·C_FP <= p·C_FN  =>  T*_Bayes = C_FP / (C_FN - C_TP + C_FP) = 100 / 500 = 0,20."),
        ("Đánh giá trên lớp mất cân bằng:", "Sử dụng đồng bộ Recall, F1, ROC-AUC, PR-AUC (nhạy với lớp thiểu số) và Brier Score (đo độ hiệu chuẩn xác suất)."),
        ("Tính cộng tính của SHAP:", "f(x) = phi_0 + Sum phi_j; TreeSHAP tính chính xác đóng góp biên của từng đặc trưng trong thời gian đa thức."),
    ], font_size=13.0, space_after=10)

    eng.add_text(s, 0.65, 7.42, 18.70, 0.36, "BẢNG 3.1: MA TRẬN CHI PHÍ BẤT ĐỐI XỨNG TRONG BÀI TOÁN GIỮ CHÂN KHÁCH HÀNG (TỶ LỆ 5 : 1)", size=13.5, bold=True, color=C_DARK_BG)
    eng.add_styled_table(
        s, 0.65, 7.85, 18.70, 2.40,
        [
            ["Trạng thái thực tế \\ Dự đoán", "Dự đoán Ở lại (0 — Không can thiệp)", "Dự đoán Rời bỏ (1 — Gửi ưu đãi CSKH)"],
            ["Thực tế Ở lại (0)", "TN: Chi phí = $0 (Quyết định đúng, không lãng phí)", "FP: Chi phí = $100 (Chi phí gửi ưu đãi dư)"],
            ["Thực tế Rời bỏ (1)", "FN: Chi phí = $500 (Tổn thất giá trị vòng đời CLV)", "TP: Chi phí = $100 (Chi phí giữ chân thành công)"],
        ],
        col_widths=[4.8, 6.95, 6.95], font_size=13.0, left_align_cols=(0, 1, 2)
    )

    # --------------------------------------------------------------------------
    # SLIDE 07: CHƯƠNG 2 — ĐẶC TẢ DỮ LIỆU & PHÂN BỐ CHURN
    # --------------------------------------------------------------------------
    s = eng.new_slide(
        nav_idx=1, section_tag="Chương 2 • Mục 2.1 – 2.2",
        title="2.1 Đặc Tả Bộ Dữ Liệu IBM Telco Churn (N = 7.043) & Phân Bố Nhãn",
        subtitle="21 thuộc tính gốc chia 4 nhóm nghiệp vụ, mất cân bằng lớp 2,77 : 1 và chênh lệch rời bỏ rõ theo hợp đồng",
        notes="Bộ dữ liệu gốc có 7.043 dòng, không có bản ghi trùng lặp, chỉ có 11 ô khuyết ở cột TotalCharges đều thuộc thuê bao mới tenure bằng 0. Biểu đồ cho thấy hợp đồng theo tháng có tỷ lệ rời bỏ 42,7%, cao gấp hơn 15 lần hợp đồng 2 năm."
    )
    eng.add_styled_table(
        s, 0.65, 2.75, 9.10, 4.35,
        [
            ["Nhóm thuộc tính", "Số biến", "Các thuộc tính tiêu biểu"],
            ["Định danh", "1", "customerID (loại bỏ trước huấn luyện)"],
            ["Nhân khẩu học", "4", "gender, SeniorCitizen, Partner, Dependents"],
            ["Dịch vụ đăng ký", "9", "PhoneService, InternetService, OnlineSecurity, TechSupport, Streaming..."],
            ["Hợp đồng & Cước", "6", "tenure, Contract, PaperlessBilling, PaymentMethod, Monthly/TotalCharges"],
            ["Biến mục tiêu", "1", "Churn (Yes = 1.869 [26,54%]; No = 5.174 [73,46%])"],
        ],
        col_widths=[2.3, 1.1, 5.7], font_size=12.0, left_align_cols=(0, 2)
    )
    eng.add_card(s, 0.65, 7.35, 9.10, 2.95, fill_color=C_SURFACE, border_color=C_STEEL)
    eng.add_text(s, 0.90, 7.55, 8.60, 0.38, "KIỂM TOÁN CHẤT LƯỢNG DỮ LIỆU GỐC", size=14.5, bold=True, color=C_DARK_BG)
    eng.add_bullets(s, 0.90, 8.00, 8.60, 2.15, [
        ("Dữ liệu sạch, không trùng lặp:", "0 bản ghi trùng; 11 ô khoảng trắng ở TotalCharges đều có tenure = 0 -> điền 0,0 theo đúng bản chất nghiệp vụ."),
        ("Phân hóa cực mạnh theo hợp đồng:", "Month-to-month có 42,7% Churn (n = 3.875) so với 11,3% (1 năm) và chỉ 2,8% (2 năm)."),
    ], font_size=13.0, space_after=8)

    eng.add_framed_image(s, os.path.join(FIG_DIR, "fig01_churn_distribution.png"), 10.05, 2.75, 9.30, 7.55,
                         caption="Hình 2.1: Phân bố nhãn Churn (26,54%) và tỷ lệ rời bỏ theo loại hợp đồng")

    # --------------------------------------------------------------------------
    # SLIDE 08: CHƯƠNG 2 — CÁC YẾU TỐ HÀNH VI CHÍNH (EDA 2x2)
    # --------------------------------------------------------------------------
    s = eng.new_slide(
        nav_idx=1, section_tag="Chương 2 • Mục 2.2",
        title="2.2 Khám Phá Các Yếu Tố Hành Vi Chi Phối Nguy Cơ Rời Bỏ",
        subtitle="Đối chiếu phân phối biến liên tục và tỷ lệ rời bỏ theo hạ tầng Internet và phương thức thanh toán",
        notes="Phân tích trực quan cho thấy khách rời bỏ tập trung ở nhóm mới dùng dịch vụ (trung vị 10 tháng so với 38 tháng), có mức cước tháng cao (trung vị 79,65 USD), dùng cáp quang Fiber optic (41,9% Churn) và thanh toán qua Electronic check (45,3% Churn)."
    )
    eng.add_framed_image(s, os.path.join(FIG_DIR, "fig02_eda_drivers.png"), 0.65, 2.75, 11.75, 7.55,
                         caption="Hình 2.2: Phân phối tenure, MonthlyCharges và tỷ lệ Churn theo Internet & PaymentMethod")

    insights_eda = [
        ("01.", "Vòng Đời Khách Hàng (tenure)", "Trung vị thời gian gắn bó của nhóm Churn chỉ là 10 tháng so với 38 tháng ở nhóm Ở lại (Mann-Whitney p = 2,42e-208).", C_DANGER),
        ("02.", "Áp Lực Cước Phí Hàng Tháng", "Nhóm rời bỏ có trung vị cước 79,65 USD/tháng, cao hơn 23,6% so với nhóm ở lại 64,43 USD/tháng (p = 3,31e-54).", C_WARN),
        ("03.", "Hạ Tầng Cáp Quang Fiber Optic", "Tỷ lệ rời bỏ ở thuê bao Fiber optic đạt 41,9%, cao gấp đôi nhóm DSL (19,0%) và gấp gần 6 lần nhóm không dùng Internet (7,4%).", C_HIGHLIGHT),
        ("04.", "Thanh Toán Electronic Check", "Khách thanh toán thủ công qua Electronic check có tỷ lệ Churn 45,3%, cao gấp ~3 lần nhóm thanh toán tự động (15,2%–16,7%).", C_STEEL),
    ]
    for i, (num, itit, idesc, icol) in enumerate(insights_eda):
        iy = 2.75 + i * 1.94
        eng.add_card(s, 12.70, iy, 6.65, 1.72, fill_color=C_SURFACE, border_color=icol)
        eng.add_numbered_badge(s, 12.90, iy + 0.18, num, fill_color=icol, w=0.68, h=0.48, font_size=14)
        eng.add_text(s, 13.70, iy + 0.24, 5.45, 0.38, itit, size=14.5, bold=True, color=C_DARK_BG)
        eng.add_text(s, 12.90, iy + 0.78, 6.25, 0.85, idesc, size=12.2, color=C_MUTED_TXT)

    # --------------------------------------------------------------------------
    # SLIDE 09: CHƯƠNG 2 — KIỂM ĐỊNH THỐNG KÊ & 6 ĐẶC TRƯNG MIỀN
    # --------------------------------------------------------------------------
    s = eng.new_slide(
        nav_idx=1, section_tag="Chương 2 • Mục 2.2 – 2.3",
        title="2.3 Kiểm Định Thống Kê (Chi-Square, Mann-Whitney) & 6 Đặc Trưng Miền",
        subtitle="Kiểm chứng định lượng độ mạnh liên kết Cramér's V, phát hiện đa cộng tuyến và mở rộng không gian đặc trưng",
        notes="Kiểm định Chi-Square xác nhận Contract và đặc trưng tạo mới TenureGroup đứng đầu về Cramér's V. Ma trận tương quan chỉ rõ đa cộng tuyến mạnh r = 0,83 giữa tenure và TotalCharges. Riêng ChargeDiff không có ý nghĩa đơn biến (p = 0,693)."
    )
    eng.add_framed_image(s, os.path.join(FIG_DIR, "fig03_correlation_cramersv.png"), 0.65, 2.75, 10.10, 7.55,
                         caption="Hình 2.3: Ma trận tương quan Pearson biến số và xếp hạng Top 10 Cramér's V")

    eng.add_text(s, 11.05, 2.72, 8.30, 0.38, "BẢNG 2.4: 6 ĐẶC TRƯNG TRI THỨC MIỀN TẠO MỚI", size=14, bold=True, color=C_DARK_BG)
    eng.add_styled_table(
        s, 11.05, 3.16, 8.30, 4.55,
        [
            ["Đặc trưng mới", "Công thức / Cách xác định", "p-value"],
            ["AvgMonthlySpend", "TotalCharges / max(tenure, 1)", "2,09e-54"],
            ["ChargeDiff", "MonthlyCharges - AvgMonthlySpend", "0,693 (k.y.n)"],
            ["TotalAddonServices", "Tổng dịch vụ gia tăng = Yes (0–6)", "4,74e-07"],
            ["HasProtectionBundle", "OnlineSecurity = Yes OR TechSupport = Yes", "2,49e-117"],
            ["IsAutoPayment", "PaymentMethod chứa 'automatic'", "1,02e-68"],
            ["TenureGroup", "5 khoảng: 0-6, 7-12, 13-24, 25-48, >48m", "2,00e-199"],
        ],
        col_widths=[2.4, 4.3, 1.6], font_size=11.5, highlight_row=6, left_align_cols=(0, 1)
    )

    eng.add_card(s, 11.05, 7.95, 8.30, 2.35, fill_color=C_SURFACE, border_color=C_SAFE)
    eng.add_text(s, 11.28, 8.12, 7.85, 0.36, "NHẬN XÉT TỪ KIỂM ĐỊNH THỐNG KÊ", size=13.5, bold=True, color=C_SAFE)
    eng.add_bullets(s, 11.28, 8.52, 7.85, 1.68, [
        ("Hiệu quả Feature Engineering:", "TenureGroup đạt Cramér's V = 0,3629 (hạng 2 toàn tập); AvgMonthlySpend lọt Top 10 SHAP."),
        ("Phát hiện đa cộng tuyến:", "r(tenure, TotalCharges) = 0,83 -> áp dụng phạt L2 cho Logistic Regression và ưu tiên mô hình cây."),
    ], font_size=12.2, space_after=6)

    # --------------------------------------------------------------------------
    # SLIDE 10: CHƯƠNG 2 — QUY TRÌNH CHỐNG RÒ RỈ DỮ LIỆU 4 BƯỚC (Chuẩn Theme2)
    # --------------------------------------------------------------------------
    s = eng.new_slide(
        nav_idx=1, section_tag="Chương 2 • Mục 2.4",
        title="2.4 Quy Trình Tiền Xử Lý 4 Bước Chống Rò Rỉ Dữ Liệu (Anti-Leakage Pipeline)",
        subtitle="Đảm bảo mọi tham số thống kê chỉ được học trên tập Train và tập kiểm thử độc lập N = 1.409 hoàn toàn bị cô lập",
        notes="Đây là nguyên tắc kỹ thuật then chốt trong khai phá dữ liệu: tách phân tầng 80/20 ngay từ đầu, làm sạch 11 giá trị khuyết TotalCharges, đóng gói toàn bộ ColumnTransformer vào sklearn Pipeline và chỉ fit trong từng fold của Stratified 5-Fold CV."
    )
    eng.add_kpi_card(s, 0.65, 2.75, 4.45, 1.80, "5.634 (80%)", "Tập Huấn luyện (Train)", "Tỷ lệ Churn: 26,54% (Stratified)", C_DARK_BG)
    eng.add_kpi_card(s, 5.40, 2.75, 4.45, 1.80, "1.409 (20%)", "Tập Kiểm thử (Test)", "374 Churn / 1.035 Ở lại", C_DANGER)
    eng.add_kpi_card(s, 10.15, 2.75, 4.45, 1.80, "25 -> 55", "Chiều đặc trưng sau OHE", "6 số + 3 nhị phân + 16 phân loại", C_HIGHLIGHT)
    eng.add_kpi_card(s, 14.90, 2.75, 4.45, 1.80, "5-Fold CV", "Stratified GridSearchCV", "Tối ưu ROC-AUC & OOF T*", C_SAFE)

    pipe_steps = [
        ("BƯỚC 1", "Phân Tách Phân Tầng Độc Lập (Stratified Split 80/20)", "Tách ngay tập Train (5.634) và Test (1.409) với random_state = 42 trước khi tính bất kỳ giá trị trung vị, trung bình hay bộ mã hóa nào.", C_DARK_BG),
        ("BƯỚC 2", "Làm Sạch & Xử Lý Giá Trị Khuyết Thiếu", "Loại bỏ mã định danh customerID; chuyển đổi cột TotalCharges sang float64 và điền 0,0 cho 11 thuê bao mới có tenure = 0.", C_STEEL),
        ("BƯỚC 3", "Biến Đổi Qua ColumnTransformer (55 Chiều)", "6 biến số: Median Imputer + StandardScaler | 3 biến nhị phân: giữ nguyên 0/1 | 16 biến phân loại: OneHotEncoder(handle_unknown='ignore').", C_HIGHLIGHT),
        ("BƯỚC 4", "Huấn Luyện Nội Khối & Chọn Ngưỡng OOF", "Gói ColumnTransformer + Estimator vào sklearn Pipeline; fit() chỉ chạy trong 4/5 fold huấn luyện; ngưỡng T* = 0,21 chọn hoàn toàn trên OOF Train.", C_SAFE),
    ]
    for i, (stag, stit, sdesc, scol) in enumerate(pipe_steps):
        py = 4.85 + i * 1.38
        eng.add_card(s, 0.65, py, 18.70, 1.20, fill_color=C_SURFACE if i % 2 == 0 else C_LIGHT_BG, border_color=scol, radius=0.08)
        eng.add_numbered_badge(s, 0.88, py + 0.25, stag, fill_color=scol, w=1.75, h=0.68, font_size=14.5)
        eng.add_text(s, 2.88, py + 0.14, 16.15, 0.38, stit, size=15.5, bold=True, color=C_DARK_BG)
        eng.add_text(s, 2.88, py + 0.56, 16.15, 0.54, sdesc, size=13.0, color=C_MUTED_TXT)

    # --------------------------------------------------------------------------
    # SLIDE 11: CHƯƠNG 3 — BẢNG ĐỐI SÁNH 5 MÔ HÌNH TRÊN TẬP TEST (N = 1.409)
    # --------------------------------------------------------------------------
    s = eng.new_slide(
        nav_idx=2, section_tag="Chương 3 • Mục 3.2",
        title="3.1 Bảng Kết Quả Thực Nghiệm Đối Sánh 5 Mô Hình Trên Tập Test (N = 1.409)",
        subtitle="Đối chiếu toàn diện Accuracy, Recall, F1, ROC-AUC, PR-AUC, Brier Score và Tổng chi phí nghiệp vụ tại ngưỡng T* = 0,21",
        notes="Trên kiểm thử chéo 5-Fold, LightGBM đạt CV ROC-AUC cao nhất (0,8494) và PR-AUC cao nhất trên Test (0,6657). XGBoost đạt Test ROC-AUC cao nhất (0,8470). Đáng chú ý, sau khi có 6 đặc trưng miền, Logistic Regression cũng đạt 0,8452 và vượt xa cây CART đơn lẻ (0,8249)."
    )
    comp_rows = [["Mô hình", "CV ROC-AUC", "Acc (0.5)", "Rec (0.5)", "F1 (0.5)", "Test AUC", "PR-AUC", "Brier Score", f"Chi phí (T*={t_star})"]]
    for nm in m_order:
        m05 = res["test_metrics_at_050"][nm]
        mt = res["test_metrics_at_tstar"][nm]
        comp_rows.append([
            nm,
            f"{res['cv_roc_auc'][nm]:.4f}",
            f"{m05['accuracy']*100:.2f}%",
            f"{m05['recall']*100:.2f}%",
            f"{m05['f1']:.4f}",
            f"{m05['roc_auc']:.4f}",
            f"{m05['pr_auc']:.4f}",
            f"{m05['brier']:.4f}",
            f"${mt['total_cost_usd']:,.0f}",
        ])
    eng.add_styled_table(
        s, 0.65, 2.75, 18.70, 3.30, comp_rows,
        col_widths=[3.2, 1.85, 1.75, 1.75, 1.65, 1.80, 1.70, 1.80, 2.20],
        font_size=12.5, highlight_row=5, left_align_cols=(0,)
    )
    eng.add_framed_image(s, os.path.join(FIG_DIR, "fig04_model_comparison.png"), 0.65, 6.28, 12.10, 4.05,
                         caption="Hình 3.1: Trực quan hóa so sánh 5 chỉ số đánh giá chính của 5 mô hình")

    eng.add_card(s, 13.05, 6.28, 6.30, 4.05, fill_color=C_SURFACE, border_color=C_SAFE)
    eng.add_text(s, 13.30, 6.48, 5.80, 0.40, "NHẬN XÉT HỌC THUẬT QUAN TRỌNG", size=14.5, bold=True, color=C_SAFE)
    eng.add_bullets(s, 13.30, 6.95, 5.80, 3.25, [
        ("Gradient Boosting dẫn đầu:", "LightGBM đạt CV AUC cao nhất (0,8494) & PR-AUC cao nhất (0,6657); XGBoost sát nút (Test AUC 0,8470)."),
        ("Vai trò của Feature Engineering:", "Logistic Regression đạt 0,8452 nhờ 6 đặc trưng miền, bỏ xa cây CART đơn (0,8249)."),
        ("Nghịch lý ngưỡng 0,50:", "Dù Accuracy ~80%, Recall chỉ đạt 51%–53% (bỏ sót gần một nửa khách rời bỏ)."),
    ], font_size=12.5, space_after=8)

    # --------------------------------------------------------------------------
    # SLIDE 12: CHƯƠNG 3 — ĐƯỜNG CONG ROC, PR-AUC & MA TRẬN NHẦM LẪN
    # --------------------------------------------------------------------------
    s = eng.new_slide(
        nav_idx=2, section_tag="Chương 3 • Mục 3.2",
        title="3.2 Phân Tích Đường Cong ROC-AUC, PR-AUC & Ma Trận Nhầm Lẫn",
        subtitle="Đánh giá khả năng phân tách hai lớp trên toàn dải ngưỡng và tác động trực tiếp khi dịch ngưỡng sang T* = 0,21",
        notes="Hình trên cho thấy các mô hình tổ hợp cây ổn định hơn hẳn cây đơn trên cả đường cong ROC và PR. Hình dưới giải phẫu ma trận nhầm lẫn: khi chuyển LightGBM từ T = 0,50 sang T* = 0,21, số ca bỏ sót khách rời bỏ FN giảm từ 181 xuống chỉ còn 57 ca."
    )
    eng.add_framed_image(s, os.path.join(FIG_DIR, "fig05_roc_pr_curves.png"), 0.65, 2.72, 18.70, 3.85,
                         caption="Hình 3.2: Đường cong ROC-AUC (A) và Precision-Recall PR-AUC (B) của 5 mô hình trên tập Test")
    eng.add_framed_image(s, os.path.join(FIG_DIR, "fig06_confusion_matrices.png"), 0.65, 6.72, 18.70, 3.62,
                         caption="Hình 3.3: Ma trận nhầm lẫn các mô hình tiêu biểu tại T = 0,50 và LightGBM tại ngưỡng tối ưu T* = 0,21 (Recall 84,8%)")

    # --------------------------------------------------------------------------
    # SLIDE 13: CHƯƠNG 3 — ĐỐI CHỨNG XỬ LÝ LỆCH LỚP & TỐI ƯU NGƯỠNG CHI PHÍ T*
    # --------------------------------------------------------------------------
    s = eng.new_slide(
        nav_idx=2, section_tag="Chương 3 • Mục 3.3",
        title="3.3 Đối Chứng Xử Lý Lệch Lớp (Ablation) & Tối Ưu Ngưỡng Chi Phí T*",
        subtitle="So sánh Mặc định (0,50) vs Trọng số Balanced vs Dịch chuyển ngưỡng trên xác suất chuẩn (T* = 0,21)",
        notes="Kết quả Ablation Study chứng minh một điểm rất hay về mặt học thuật: gán trọng số balanced làm méo độ hiệu chuẩn xác suất (Brier Score tăng từ 0,1348 lên 0,1630). Trong khi đó, dịch chuyển ngưỡng trên mô hình gốc sang T* = 0,21 (khớp sát Bayes 0,20) vừa giữ nguyên Brier Score tốt nhất, vừa nâng Recall lên 84,76% và đạt chi phí thấp nhất $94.300."
    )
    eng.add_framed_image(s, os.path.join(FIG_DIR, "fig07_cost_threshold_roi.png"), 0.65, 2.72, 18.70, 4.35,
                         caption="Hình 3.4: Đường cong chi phí OOF Train vs Test xác định T* = 0,21 (A) và Đối chứng 4 chiến lược (B)")

    ab_rows = [
        ["Phương pháp trên LightGBM", "Ngưỡng", "Accuracy", "Precision", "Recall", "F1-Score", "Brier Score", "Bỏ sót (FN)", "Tổng chi phí"],
        ["Mặc định (Unweighted)", "0,50", f"{ab_unw['accuracy']*100:.2f}%", f"{ab_unw['precision']*100:.2f}%", f"{ab_unw['recall']*100:.2f}%", f"{ab_unw['f1']:.4f}", f"{ab_unw['brier']:.4f}", str(ab_unw["fn"]), f"${ab_unw['total_cost_usd']:,.0f}"],
        ["Trọng số lớp (Balanced)", "0,50", f"{ab_bal['accuracy']*100:.2f}%", f"{ab_bal['precision']*100:.2f}%", f"{ab_bal['recall']*100:.2f}%", f"{ab_bal['f1']:.4f}", f"{ab_bal['brier']:.4f}", str(ab_bal["fn"]), f"${ab_bal['total_cost_usd']:,.0f}"],
        [f"Tối ưu ngưỡng chi phí (T*)", f"{t_star:.2f}", f"{ab_opt['accuracy']*100:.2f}%", f"{ab_opt['precision']*100:.2f}%", f"{ab_opt['recall']*100:.2f}%", f"{ab_opt['f1']:.4f}", f"{ab_opt['brier']:.4f}", str(ab_opt["fn"]), f"${ab_opt['total_cost_usd']:,.0f}"],
    ]
    eng.add_styled_table(
        s, 0.65, 7.25, 18.70, 2.15, ab_rows,
        col_widths=[3.8, 1.3, 1.6, 1.6, 1.6, 1.6, 1.8, 1.7, 2.1],
        font_size=12.5, highlight_row=3, left_align_cols=(0,)
    )
    eng.add_card(s, 0.65, 9.56, 18.70, 0.78, fill_color=C_SAFE_BG, border_color=C_SAFE, radius=0.10)
    eng.add_text(s, 0.90, 9.73, 18.20, 0.42,
                 "Phát hiện cốt lõi: Ngưỡng T* = 0,21 (khớp sát lý thuyết Bayes 0,20) giữ nguyên Brier Score 0,1348 (không bị méo như balanced 0,1630) và giảm chi phí thêm $24.600 so với ngưỡng 0,50.",
                 size=13.0, bold=True, color=C_DARK_BG)

    # --------------------------------------------------------------------------
    # SLIDE 14: CHƯƠNG 3 — PHÂN TÍCH ODDS RATIO (LOGISTIC REGRESSION) & LUẬT CART
    # --------------------------------------------------------------------------
    s = eng.new_slide(
        nav_idx=3, section_tag="Chương 4 • Mục 4.1",
        title="4.1 Rút Trích Hệ Luật Phân Loại IF-THEN Từ Cây Quyết Định CART",
        subtitle="Chuyển hóa cây quyết định cắt tỉa độ sâu 3 trên thuộc tính gốc thành 8 luật nghiệp vụ dễ đọc (Support & Xác suất Churn)",
        notes="Từ cây quyết định CART cắt tỉa độ sâu 3, nhóm trích xuất được các luật tường minh. Đáng chú ý nhất là Luật R1: khách hợp đồng theo tháng, dùng cáp quang Fiber optic và gắn bó dưới 13,5 tháng chiếm 13,5% dữ liệu nhưng có tỷ lệ rời bỏ tới 71,0%; trái ngược với Luật R8 chỉ có 2,4%."
    )
    eng.add_framed_image(s, os.path.join(FIG_DIR, "fig08_decision_tree_rules.png"), 0.65, 2.72, 18.70, 4.30,
                         caption="Hình 4.1: Sơ đồ cây quyết định CART cắt tỉa độ sâu 3 trên tập huấn luyện (N = 5.634)")

    rule_rows = [
        ["Mã luật", "Điều kiện phân tách tường minh (IF ...)", "Số mẫu (Support)", "Tỷ lệ Churn", "Phân khúc rủi ro"],
        ["R1", "Hợp đồng = Theo tháng AND Internet = Fiber optic AND tenure <= 13,5 tháng", "762 (13,5%)", "71,0%", "Rủi ro rất cao"],
        ["R2", "Hợp đồng = Theo tháng AND Internet != Fiber optic AND tenure <= 3,5 tháng", "468 (8,3%)", "43,2%", "Rủi ro cao"],
        ["R3", "Hợp đồng = Theo tháng AND Internet = Fiber optic AND tenure > 13,5 tháng", "945 (16,8%)", "42,2%", "Rủi ro cao"],
        ["R5", "Hợp đồng = Theo tháng AND Internet != Fiber optic AND tenure > 3,5 tháng", "927 (16,5%)", "19,9%", "Trung bình"],
        ["R8", "Hợp đồng 1–2 năm AND MonthlyCharges <= $93,67 AND có OnlineSecurity", "1.505 (26,7%)", "2,4%", "Trung thành cốt lõi"],
    ]
    eng.add_styled_table(
        s, 0.65, 7.20, 18.70, 3.12, rule_rows,
        col_widths=[1.2, 9.6, 2.6, 2.0, 3.3],
        font_size=12.0, highlight_row=1, left_align_cols=(1,)
    )

    # --------------------------------------------------------------------------
    # SLIDE 15: CHƯƠNG 4 — GIẢI THÍCH TOÀN CỤC SHAP (GLOBAL XAI)
    # --------------------------------------------------------------------------
    s = eng.new_slide(
        nav_idx=3, section_tag="Chương 4 • Mục 4.2.1",
        title="4.2 Giải Thích Mô Hình Toàn Cục (Global XAI) Với SHAP Summary Plot",
        subtitle="Định lượng mức đóng góp và chiều hướng tác động của Top 12 đặc trưng trên mô hình LightGBM tối ưu",
        notes="Biểu đồ SHAP Beeswarm cho thấy sự nhất quán hoàn toàn giữa kiểm định thống kê, hệ luật CART và mô hình LightGBM. Đứng đầu là Contract_Month-to-month (0,663), tiếp theo là tenure (0,357), thiếu OnlineSecurity (0,238), thiếu TechSupport (0,201) và Fiber optic (0,200)."
    )
    eng.add_framed_image(s, os.path.join(FIG_DIR, "fig09_shap_summary.png"), 0.65, 2.75, 11.45, 7.55,
                         caption="Hình 4.2: SHAP Summary Beeswarm Plot — Top 12 đặc trưng trên tập Test (LightGBM)")

    eng.add_card(s, 12.40, 2.75, 6.95, 7.55, fill_color=C_SURFACE, border_color=C_HIGHLIGHT)
    eng.add_numbered_badge(s, 12.65, 3.00, "XAI", fill_color=C_HIGHLIGHT, w=0.95, h=0.52, font_size=15)
    eng.add_text(s, 13.75, 3.08, 5.40, 0.45, "Ý NGHĨA NGHIỆP VỤ TOP ĐẶC TRƯNG", size=15.5, bold=True, color=C_DARK_BG)
    eng.add_bullets(s, 12.65, 3.75, 6.45, 6.30, [
        ("1. Contract_Month-to-month (0,663):", "Yếu tố chi phối số 1; hợp đồng ngắn hạn không có ràng buộc làm tăng mạnh log-odds rời bỏ."),
        ("2. tenure (0,357):", "Quan hệ nghịch rõ rệt: thuê bao mới (chấm xanh) có SHAP dương cao; thuê bao lâu năm (chấm đỏ) rất ổn định."),
        ("3–4. OnlineSecurity_No & TechSupport_No:", "Việc thiếu hai dịch vụ bảo vệ/hỗ trợ đứng hạng 3 (0,238) và hạng 4 (0,201)."),
        ("5–6. Fiber optic & Electronic check:", "Cáp quang (0,200) và thanh toán thủ công (0,156) làm gia tăng rõ rệt xác suất hủy hợp đồng."),
        ("Kiểm chứng Odds Ratio (LR):", "Logistic Regression cũng xác nhận Fiber optic (OR = 2,09) & Month-to-month (OR = 1,95) là hai yếu tố rủi ro lớn nhất."),
        ("Đặc trưng miền AvgMonthlySpend:", "Đứng hạng 10 toàn cục (0,097), khẳng định giá trị thực tiễn của bước Feature Engineering."),
    ], font_size=12.8, space_after=10)

    # --------------------------------------------------------------------------
    # SLIDE 16: CHƯƠNG 4 — GIẢI THÍCH CỤC BỘ SHAP (LOCAL XAI)
    # --------------------------------------------------------------------------
    s = eng.new_slide(
        nav_idx=3, section_tag="Chương 4 • Mục 4.2.2",
        title="4.3 Giải Thích Quyết Định Cho Từng Cá Thể (Local SHAP)",
        subtitle="Đối chiếu phân rã đóng góp đặc trưng trên hai hồ sơ đối lập thuộc tập kiểm thử độc lập",
        notes="Local SHAP giúp nhân viên CSKH hiểu rõ lý do cụ thể của từng trường hợp cảnh báo thay vì chỉ thấy một con số xác suất. Hồ sơ 5178-LMXOP bị dự đoán 90,7% do mới dùng 1 tháng, hợp đồng tháng và thiếu hỗ trợ kỹ thuật; hồ sơ 7606-BPHHN chỉ 0,6% nhờ gắn bó 72 tháng và hợp đồng 2 năm."
    )
    eng.add_framed_image(s, os.path.join(FIG_DIR, "fig10_shap_waterfall_cases.png"), 0.65, 2.72, 18.70, 5.20,
                         caption="Hình 4.3: Đối chiếu đóng góp SHAP cục bộ giữa ca rủi ro cao (A) và ca trung thành (B)")

    eng.add_card(s, 0.65, 8.12, 9.15, 2.18, fill_color=C_SURFACE, border_color=C_DANGER)
    eng.add_numbered_badge(s, 0.88, 8.32, "90,7%", fill_color=C_DANGER, w=1.25, h=0.52, font_size=15)
    eng.add_text(s, 2.28, 8.38, 7.30, 0.42, "CA RỦI RO CAO: MÃ KH 5178-LMXOP (THỰC TẾ: RỜI BỎ)", size=14.5, bold=True, color=C_DANGER)
    eng.add_bullets(s, 0.88, 8.95, 8.70, 1.25, [
        ("Hồ sơ:", "tenure = 1 tháng (+0,777), Month-to-month (+0,770), TechSupport_No (+0,287), Fiber optic, cước $95,10 -> Cần ưu tiên gọi tư vấn chuyển đổi gói ngay."),
    ], font_size=12.6, space_after=4)

    eng.add_card(s, 10.20, 8.12, 9.15, 2.18, fill_color=C_SURFACE, border_color=C_SAFE)
    eng.add_numbered_badge(s, 10.43, 8.32, "0,6%", fill_color=C_SAFE, w=1.25, h=0.52, font_size=15)
    eng.add_text(s, 11.83, 8.38, 7.30, 0.42, "CA TRUNG THÀNH: MÃ KH 7606-BPHHN (THỰC TẾ: Ở LẠI)", size=14.5, bold=True, color=C_SAFE)
    eng.add_bullets(s, 10.43, 8.95, 8.70, 1.25, [
        ("Hồ sơ:", "tenure = 72 tháng (-0,967), không dùng hợp đồng tháng (-0,696), hợp đồng 2 năm (-0,303), thanh toán tự động -> Nhóm ổn định, tránh lãng phí ưu đãi."),
    ], font_size=12.6, space_after=4)

    # --------------------------------------------------------------------------
    # SLIDE 17: CHƯƠNG 4 — PHÂN TÍCH LỖI CHUYÊN SÂU & ĐỐI CHỨNG ROI 4 CHIẾN LƯỢC
    # --------------------------------------------------------------------------
    s = eng.new_slide(
        nav_idx=3, section_tag="Chương 4 • Mục 4.3 – 4.4",
        title="4.4 Phân Tích Lỗi Chuyên Sâu (Error Analysis) & Đối Chứng Hiệu Quả ROI",
        subtitle="Giải phẫu nguyên nhân gốc rễ các ca FN/FP và định lượng chi phí thực tế giữa 4 kịch bản vận hành trên tập Test",
        notes="Phân tích lỗi cho thấy 34,2% lỗi nằm ở vùng biên xác suất và 41/57 ca bỏ sót là khách dài hạn rời bỏ đột xuất do yếu tố ngoại sinh chưa có trong dữ liệu tĩnh. Về hiệu quả kinh tế, LightGBM tại T* = 0,21 tiết kiệm $92.700 (-49,6%) so với không can thiệp và rẻ hơn 23,4% so với luật lọc thủ công Excel."
    )
    err_rows = [
        ["Hạng mục giải phẫu lỗi (N = 1.409)", "Ngưỡng T = 0,50", f"Ngưỡng T* = {t_star:.2f}"],
        ["Số ca bỏ sót khách rời bỏ (FN)", "181 ca", f"{errt['fn_count']} ca (giảm 68,5%)"],
        ["  - Khách dài hạn / Hợp đồng 1–2 năm (FN)", "92 ca (50,8%)", f"{errt['fn_loyal_silent']} ca (71,9% số FN)"],
        ["  - Khách ngắn hạn hợp đồng tháng (FN)", "89 ca (49,2%)", f"{errt['fn_short_volatile']} ca (28,1% số FN)"],
        ["Số ca cảnh báo dư (FP)", "91 ca", f"{errt['fp_count']} ca"],
        ["  - Mang hồ sơ rủi ro cao (Tháng / Fiber)", "91 ca (100%)", f"{errt['fp_high_risk_profile']} ca (99,1% số FP)"],
        ["Lỗi sát vùng biên xác suất (|p - T| <= 0,10)", "93 ca (34,2%)", f"{errt['boundary_errors']} ca ({errt['boundary_share_pct']:.1f}%)"],
    ]
    eng.add_styled_table(
        s, 0.65, 2.72, 11.10, 3.85, err_rows,
        col_widths=[5.8, 2.4, 2.9], font_size=11.8, highlight_row=1, left_align_cols=(0,)
    )

    eng.add_card(s, 12.05, 2.72, 7.30, 3.85, fill_color=C_SURFACE, border_color=C_DANGER)
    eng.add_text(s, 12.30, 2.92, 6.80, 0.38, "NGUYÊN NHÂN GỐC RỄ & GIẢI PHÁP", size=14.5, bold=True, color=C_DANGER)
    eng.add_bullets(s, 12.30, 3.38, 6.80, 3.05, [
        ("Giới hạn dữ liệu tĩnh:", "71,9% ca FN còn lại là khách dài hạn rời bỏ đột xuất (sự cố mạng, chuyển chỗ ở, đối thủ) chưa có log khiếu nại."),
        ("Tính hợp lý của nhóm FP:", "99,1% ca FP thuộc nhóm hồ sơ rủi ro cao (hợp đồng tháng / Fiber) nên chăm sóc phòng ngừa vẫn có lợi."),
        ("Vùng biên 34,2%:", "Dải 0,21–0,40 nên chăm sóc qua kênh tự động chi phí thấp (SMS/Email)."),
    ], font_size=12.4, space_after=8)

    eng.add_text(s, 0.65, 6.76, 18.70, 0.36, "BẢNG 4.4: ĐỐI CHỨNG TỔNG CHI PHÍ 4 CHIẾN LƯỢC VẬN HÀNH TRÊN TẬP TEST (N = 1.409 | 374 CA CHURN)", size=13.5, bold=True, color=C_DARK_BG)
    roi_rows = [["Chiến lược chăm sóc", "KH can thiệp", "Churn giữ được (TP)", "Churn bỏ sót (FN)", "Tổng chi phí", "Tiết kiệm vs Không can thiệp"]]
    for r in roi:
        roi_rows.append([
            r["strategy"], f"{r['targeted_customers']:,}", str(r["churn_caught"]),
            str(r["churn_missed"]), f"${r['total_cost_usd']:,.0f}",
            f"${r['savings_vs_nothing_usd']:,.0f} ({r['savings_pct']:.1f}%)"
        ])
    eng.add_styled_table(
        s, 0.65, 7.18, 18.70, 3.12, roi_rows,
        col_widths=[4.6, 2.2, 2.6, 2.4, 2.6, 4.3],
        font_size=12.5, highlight_row=4, left_align_cols=(0,)
    )

    # --------------------------------------------------------------------------
    # SLIDE 18: CHƯƠNG 5 — ỨNG DỤNG MINH HỌA STREAMLIT & CHIẾN LƯỢC CSKH
    # --------------------------------------------------------------------------
    s = eng.new_slide(
        nav_idx=4, section_tag="Chương 5 • Mục 5.1 – 5.2",
        title="5.1 Ứng Dụng Minh Họa Streamlit & Chiến Lược CSKH Dựa Trên Tri Thức",
        subtitle="Đóng gói pipeline LightGBM + SHAP thành công cụ tra cứu cục bộ và bản đồ 4 nhóm hành động giữ chân khách hàng",
        notes="Ứng dụng Streamlit chạy 100% cục bộ không cần API key với 3 phân hệ: dự đoán cá nhân kèm biểu đồ SHAP trực tiếp, chấm điểm hàng loạt file CSV và bảng tổng quan kết quả. Bên phải là 4 nhóm khuyến nghị CSKH bám sát luật R1, SHAP và vùng lỗi biên."
    )
    eng.add_card(s, 0.65, 2.75, 7.55, 7.55, fill_color=C_SURFACE, border_color=C_HIGHLIGHT)
    eng.add_numbered_badge(s, 0.90, 3.00, "UI", fill_color=C_HIGHLIGHT, w=0.80, h=0.52, font_size=15)
    eng.add_text(s, 1.85, 3.08, 6.10, 0.45, "NGUYÊN MẪU WEB STREAMLIT (100% OFFLINE)", size=15.5, bold=True, color=C_DARK_BG)
    eng.add_bullets(s, 0.90, 3.78, 7.05, 6.25, [
        ("Tab 1 — Dự đoán khách hàng cá nhân:", "Nhập hồ sơ thuê bao, tự động tính 6 đặc trưng miền, đối chiếu ngưỡng T* = 0,21, vẽ biểu đồ thanh SHAP cục bộ và đề xuất kịch bản CSKH cụ thể."),
        ("Tab 2 — Chấm điểm hàng loạt (CSV):", "Tải lên danh sách khách hàng (kèm tệp mẫu 100 KH tập Test), xếp hạng rủi ro giảm dần và xuất tệp CSV danh sách ưu tiên liên hệ."),
        ("Tab 3 — Tổng quan thực nghiệm:", "Tra cứu trực tiếp bảng đối sánh 5 mô hình, bảng hiệu quả chi phí 4 chiến lược và các biểu đồ phân tích chính."),
        ("Vận hành độc lập:", "Khởi chạy nhanh bằng tệp run_demo.bat hoặc lệnh python -m streamlit run app.py; không phụ thuộc máy chủ ngoài hay API key."),
    ], font_size=13.2, space_after=14)

    eng.add_text(s, 8.50, 2.72, 10.85, 0.38, "BẢNG 5.1: BẢN ĐỒ HÀNH ĐỘNG CSKH THEO PHÂN KHÚC RỦI RO KHAI PHÁ", size=14.0, bold=True, color=C_DARK_BG)
    eng.add_styled_table(
        s, 8.50, 3.16, 10.85, 7.14,
        [
            ["Phân khúc rủi ro phát hiện", "Cơ sở từ kết quả khai phá", "Hành động chăm sóc đề xuất"],
            ["Hợp đồng tháng + Fiber optic + tenure <= 13 tháng", "Luật R1: chiếm 13,5% mẫu nhưng tỷ lệ Churn lên tới 71,0%.", "Ưu tiên gọi CSKH chủ động ở tháng 2 & 6; đề xuất gói chuyển đổi hợp đồng 1 năm giảm 10–15% cước 3 tháng đầu."],
            ["Thuê bao Cáp quang chưa dùng Security / TechSupport", "OnlineSecurity_No và TechSupport_No xếp hạng 3 & 4 trên SHAP.", "Tặng gói dùng thử miễn phí 3 tháng OnlineSecurity và TechSupport nhằm tăng trải nghiệm và rào cản chuyển đổi."],
            ["Thanh toán thủ công qua Electronic check", "Tỷ lệ Churn 45,3% (gấp ~3 lần thanh toán tự động); SHAP hạng 6.", "Khuyến khích liên kết thanh toán tự động qua ngân hàng / thẻ tín dụng kèm ưu đãi trừ thẳng hóa đơn kỳ đầu."],
            ["Nhóm xác suất vùng biên (0,21 – 0,40)", "Chiếm 34,2% vùng lỗi biên quanh ngưỡng T*.", "Triển khai kênh chăm sóc tự động chi phí thấp (SMS/Email khảo sát, ưu đãi băng thông) thay vì gọi điện trực tiếp."],
        ],
        col_widths=[2.9, 2.8, 5.15], font_size=12.0, left_align_cols=(0, 1, 2)
    )

    # --------------------------------------------------------------------------
    # SLIDE 19: KẾT LUẬN & HƯỚNG PHÁT TRIỂN (3-Column chuẩn Theme2 Slide 31)
    # --------------------------------------------------------------------------
    s = eng.new_slide(
        nav_idx=4, section_tag="Kết luận • Tổng kết đề tài",
        title="Tổng Kết Đóng Góp, Hạn Chế & Hướng Phát Triển Tiếp Theo",
        subtitle="Đánh giá mức độ hoàn thành mục tiêu học phần Khai phá dữ liệu và lộ trình mở rộng nghiên cứu",
        notes="Nhóm đã hoàn thành đầy đủ chu trình KDD trên dữ liệu thực IBM Telco Churn, chứng minh giá trị của học nhạy chi phí và giải thích mô hình. Hướng mở rộng tiếp theo tập trung vào dữ liệu chuỗi thời gian, Survival Analysis và kết hợp phân cụm RFM."
    )
    cols_kl = [
        ("01.", "KẾT QUẢ ĐẠT ĐƯỢC", C_SAFE, [
            ("Quy trình KDD chặt chẽ:", "Kiểm định thống kê đầy đủ, 6 đặc trưng miền và pipeline chống rò rỉ dữ liệu tuyệt đối."),
            ("Hiệu năng & Chi phí tối ưu:", "LightGBM đạt CV AUC 0,8494, PR-AUC 0,6657; ngưỡng T* = 0,21 nâng Recall lên 84,8% và giảm 49,6% chi phí."),
            ("Tri thức minh bạch:", "8 luật CART tường minh, SHAP toàn cục/cục bộ, giải phẫu lỗi FN/FP và nguyên mẫu Streamlit chạy cục bộ."),
        ]),
        ("02.", "HẠN CHẾ TỒN TẠI", C_WARN, [
            ("Dữ liệu tĩnh (Snapshot):", "Bộ dữ liệu chỉ phản ánh trạng thái tại một thời điểm, chưa ghi nhận biến động sử dụng theo từng tháng."),
            ("Thiếu nhật ký tương tác CSKH:", "Chưa có số lần gọi tổng đài khiếu nại hay sự cố đường truyền — nguyên nhân chính của 41 ca FN dài hạn."),
            ("Chi phí giữ chân giả định:", "Ma trận chi phí 5:1 theo chuẩn tham khảo ngành viễn thông; cần hiệu chỉnh theo CLV từng phân khúc thực tế."),
        ]),
        ("03.", "HƯỚNG PHÁT TRIỂN", C_HIGHLIGHT, [
            ("Tích hợp chuỗi thời gian:", "Bổ sung đặc trưng xu hướng lưu lượng và tần suất khiếu nại 30–90 ngày gần nhất để bắt sớm ca rời bỏ đột xuất."),
            ("Phân tích sống còn (Survival):", "Áp dụng Cox / Random Survival Forest để dự báo đồng thời xác suất và thời điểm dự kiến rời bỏ."),
            ("Kết hợp Phân cụm (Clustering):", "Gom cụm khách hàng theo RFM trước khi phân loại nhằm phân bổ ngân sách giữ chân theo đúng giá trị CLV."),
        ]),
    ]
    for i, (num, ctitle, col, citems) in enumerate(cols_kl):
        cx = 0.65 + i * 6.32
        eng.add_card(s, cx, 2.80, 6.06, 7.45, fill_color=C_SURFACE, border_color=col)
        eng.add_numbered_badge(s, cx + 0.25, 3.05, num, fill_color=col, w=0.85, h=0.58)
        eng.add_text(s, cx + 1.25, 3.14, 4.55, 0.45, ctitle, size=17, bold=True, color=C_DARK_BG)
        eng.add_bullets(s, cx + 0.25, 3.92, 5.56, 6.00, citems, font_size=13.0, space_after=12)

    # --------------------------------------------------------------------------
    # SLIDE 20: PHỤ LỤC — SIÊU THAM SỐ TỐI ƯU & CÂU HỎI PHẢN BIỆN Q&A
    # --------------------------------------------------------------------------
    s = eng.new_slide(
        nav_idx=4, section_tag="Phụ lục • Tra cứu & Q&A",
        title="[Phụ Lục] Siêu Tham Số Tối Ưu (5-Fold GridSearchCV) & Câu Hỏi Thảo Luận",
        subtitle="Tài liệu tra cứu kỹ thuật và phương án giải trình cho 4 câu hỏi trọng tâm trong buổi bảo vệ",
        notes="Slide phụ lục tổng hợp bộ siêu tham số tốt nhất của 5 mô hình sau Stratified 5-Fold GridSearchCV cùng 4 nội dung giải trình kỹ thuật thường gặp khi bảo vệ bài tập lớn môn Khai phá dữ liệu."
    )
    hp_rows = [["Mô hình", "Bộ siêu tham số tối ưu sau Stratified 5-Fold GridSearchCV", "CV ROC-AUC", "Thời gian"]]
    for nm in m_order:
        hp_rows.append([
            nm,
            ", ".join(f"{k} = {v}" for k, v in res["best_params"][nm].items()),
            f"{res['cv_roc_auc'][nm]:.4f}",
            f"{res['train_times_sec'][nm]:.2f}s",
        ])
    eng.add_styled_table(
        s, 0.65, 2.72, 18.70, 2.95, hp_rows,
        col_widths=[3.6, 10.6, 2.4, 2.1],
        font_size=12.2, highlight_row=5, left_align_cols=(0, 1)
    )

    qa_list = [
        ("Q1.", "Vì sao dùng Threshold Moving thay vì SMOTE hoặc class_weight='balanced'?",
         "Kết quả Ablation Study chứng minh class_weight='balanced' làm méo xác suất dự báo (Brier tăng từ 0,1348 lên 0,1630). Dịch ngưỡng T* = 0,21 trên xác suất gốc giữ nguyên hiệu chuẩn và cho tổng chi phí thấp nhất ($94.300)."),
        ("Q2.", "Vì sao ngưỡng tối ưu T* = 0,21 lại được chọn trên OOF Train chứ không phải Test?",
         "Chọn ngưỡng trên tập Test sẽ gây rò rỉ dữ liệu và đánh giá lạc quan sai lệch. T* = 0,21 tìm trên OOF 5-Fold khớp rất sát lý thuyết Bayes (0,20) và kiểm chứng độc lập trên Test."),
        ("Q3.", "Logistic Regression có Test AUC 0,8452 gần bằng LightGBM (0,8462), tại sao chọn LightGBM?",
         "LightGBM đạt CV ROC-AUC (0,8494) và PR-AUC (0,6657) cao nhất, xử lý tốt tương tác phi tuyến, không nhạy đa cộng tuyến (r = 0,83 giữa tenure và TotalCharges) và hỗ trợ TreeSHAP chính xác."),
        ("Q4.", "Tại sao vẫn còn 57 ca bỏ sót (FN) tại ngưỡng tối ưu T* = 0,21?",
         "Giải phẫu lỗi chỉ ra 41/57 ca (71,9%) là khách hàng dài hạn đột ngột rời bỏ do tác nhân ngoài dữ liệu tĩnh (chuyển nơi ở, sự cố mạng). Cần bổ sung log khiếu nại theo thời gian để cải thiện nhóm này."),
    ]
    q_pos = [(0.65, 5.95), (10.20, 5.95), (0.65, 8.20), (10.20, 8.20)]
    for (qx, qy), (qcode, qtxt, atxt) in zip(q_pos, qa_list):
        eng.add_card(s, qx, qy, 9.15, 2.08, fill_color=C_SURFACE, border_color=C_STEEL)
        eng.add_numbered_badge(s, qx + 0.18, qy + 0.18, qcode, fill_color=C_DARK_BG, w=0.72, h=0.45, font_size=13.5)
        eng.add_text(s, qx + 1.02, qy + 0.16, 7.95, 0.56, qtxt, size=12.8, bold=True, color=C_DARK_BG)
        eng.add_text(s, qx + 0.18, qy + 0.80, 8.80, 1.18, "→ " + atxt, size=11.8, color=C_MUTED_TXT)

    # --------------------------------------------------------------------------
    # SLIDE 21: THANK YOU & Q&A (Closing Hero chuẩn Theme2 Slide 34)
    # --------------------------------------------------------------------------
    s = eng.new_slide(
        nav_idx=None,
        notes="Xin trân trọng cảm ơn TS. Bùi Ngọc Thăng và các anh chị đã theo dõi phần trình bày của Nhóm 02. Nhóm rất mong nhận được ý kiến đóng góp và sẵn sàng trình diễn trực tiếp ứng dụng minh họa Streamlit."
    )
    b1 = s.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(13.6), Inches(0), Inches(6.4), Inches(0.95))
    b1.fill.solid()
    b1.fill.fore_color.rgb = C_DARK_BG
    b1.line.fill.background()

    b2 = s.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(0), Inches(10.30), Inches(6.4), Inches(0.95))
    b2.fill.solid()
    b2.fill.fore_color.rgb = C_DARK_BG
    b2.line.fill.background()

    if os.path.exists(LOGO_PATH):
        pic_end = s.shapes.add_picture(LOGO_PATH, Inches(9.35), Inches(1.40), height=Inches(1.30))
        pic_end.left = int((SLIDE_W - pic_end.width) / 2)

    eng.add_text(s, 2.0, 3.25, 16.0, 1.50, "TRÂN TRỌNG CẢM ƠN", size=58, bold=True, color=C_DARK_BG, align=PP_ALIGN.CENTER)
    eng.add_text(s, 2.5, 4.95, 15.0, 0.80, "KÍNH MỜI TS. BÙI NGỌC THĂNG & ANH/CHỊ ĐẶT CÂU HỎI THẢO LUẬN (Q&A)", size=21, bold=True, color=C_STEEL, align=PP_ALIGN.CENTER)

    div = s.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(7.5), Inches(6.10), Inches(5.0), Inches(0.06))
    div.fill.solid()
    div.fill.fore_color.rgb = C_STEEL
    div.line.fill.background()

    eng.add_text(s, 2.5, 6.58, 15.0, 0.50, "Đề tài: Ứng Dụng Các Mô Hình Phân Loại Dữ Liệu Trong Dự Đoán Nguy Cơ Rời Bỏ Dịch Vụ Của Khách Hàng", size=16, bold=True, color=C_DARK_BG, align=PP_ALIGN.CENTER)
    eng.add_text(s, 2.5, 7.18, 15.0, 0.45, "Nhóm 02 (Đỗ Ngọc Điền • Dương Tùng Lâm • Trần Thị Thu Hương • Phạm Văn Chiến • Vàng A Hứ)", size=14.5, color=C_MUTED_TXT, align=PP_ALIGN.CENTER)
    eng.add_text(s, 2.5, 7.72, 15.0, 0.45, "Lớp Cao học CNTT K20.2 • Trường Đại học Kinh doanh và Công nghệ Hà Nội (HUBT)", size=14.0, color=C_STEEL, align=PP_ALIGN.CENTER)
    eng.add_text(s, 2.5, 8.48, 15.0, 0.45, "Mã nguồn & Báo cáo: https://github.com/Lamdt198/KhaiPhaDuLieu-CustomerChurn  |  Demo: app.py", size=13.5, bold=True, color=C_HIGHLIGHT, align=PP_ALIGN.CENTER)

    os.makedirs(REPORT_DIR, exist_ok=True)
    final_out = OUTPUT_PATH
    try:
        eng.save(final_out)
    except PermissionError:
        final_out = os.path.join(REPORT_DIR, "Slide_Thuyet_Trinh_Khai_Pha_Du_Lieu_Nhom02_Theme2.pptx")
        eng.save(final_out)
        print(f"NOTE: {OUTPUT_PATH} is currently open in PowerPoint -> saved to {final_out}")
    print(f"SUCCESS: Generated {eng.slide_num} slides -> {final_out}")


if __name__ == "__main__":
    build_all_slides()
