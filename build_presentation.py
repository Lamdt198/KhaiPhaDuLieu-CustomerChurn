# -*- coding: utf-8 -*-
"""
build_presentation.py
Sinh Slide thuyet trinh PowerPoint (.pptx) 17 trang chuan Master Design System HUBT K20.2.
Hoc phan: Khai pha du lieu (Data Mining) - Nhom 02 - Lop Cao hoc CNTT 20.2.
Giang vien huong dan: TS. Bui Ngoc Thang.
Doc truc tiep so lieu tu models/eda_stats.json va models/experiment_results.json.
"""

import json
import os
import sys
from PIL import Image
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_SHAPE
from pptx.oxml.ns import qn
from pptx.oxml import parse_xml

if sys.stdout and hasattr(sys.stdout, "reconfigure"):
    try:
        sys.stdout.reconfigure(encoding="utf-8")
    except Exception:
        pass

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
ASSET_DIR = os.path.join(BASE_DIR, "assets")
MODEL_DIR = os.path.join(BASE_DIR, "models")
REPORT_DIR = os.path.join(BASE_DIR, "reports")
FIG_DIR = os.path.join(REPORT_DIR, "figures")
OUTPUT_PPTX = os.path.join(REPORT_DIR, "Slide_Thuyet_Trinh_Khai_Pha_Du_Lieu_Nhom02.pptx")
LOGO_PATH = os.path.join(ASSET_DIR, "logo_hubt_transparent.png")

# Design System HUBT K20.2
FONT_MAIN = "Arial"
FONT_CODE = "Consolas"

C_NAVY_DARK = RGBColor(15, 37, 68)      # #0F2544
C_NAVY_MED = RGBColor(26, 62, 114)      # #1A3E72
C_BLUE_ACCENT = RGBColor(37, 99, 235)   # #2563EB
C_GOLD = RGBColor(217, 119, 6)          # #D97706
C_TEAL = RGBColor(15, 118, 110)         # #0F766E
C_RED = RGBColor(185, 28, 28)           # #B91C1C
C_BG_SLIDE = RGBColor(244, 246, 250)    # #F4F6FA
C_BG_CARD = RGBColor(255, 255, 255)     # #FFFFFF
C_ROW_ALT = RGBColor(241, 245, 249)     # #F1F5F9
C_BORDER = RGBColor(203, 213, 225)      # #CBD5E1
C_TEXT_DARK = RGBColor(15, 23, 42)      # #0F172A
C_TEXT_MUTED = RGBColor(71, 85, 105)    # #475569
C_WHITE = RGBColor(255, 255, 255)

SLIDE_W = Inches(13.333)
SLIDE_H = Inches(7.500)


def set_cell_border(cell, border_color="CBD5E1", sz="6"):
    tcPr = cell._tc.get_or_add_tcPr()
    for edge in ("lnL", "lnR", "lnT", "lnB"):
        existing = tcPr.find(qn(f"a:{edge}"))
        if existing is not None:
            tcPr.remove(existing)
        tcPr.append(
            parse_xml(
                f'<a:{edge} xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" w="{int(sz)*1270}" cmpd="sng">'
                f'<a:solidFill><a:srgbClr val="{border_color}"/></a:solidFill>'
                f"</a:{edge}>"
            )
        )


def add_notes(slide, text):
    if text:
        slide.notes_slide.notes_text_frame.text = text.strip()


def apply_frame(slide, num, badge_txt, title_txt, notes=""):
    bg = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, 0, 0, SLIDE_W, SLIDE_H)
    bg.fill.solid()
    bg.fill.fore_color.rgb = C_BG_SLIDE
    bg.line.fill.background()

    hdr = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, 0, 0, SLIDE_W, Inches(0.84))
    hdr.fill.solid()
    hdr.fill.fore_color.rgb = C_NAVY_DARK
    hdr.line.fill.background()

    stripe = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, 0, Inches(0.84), SLIDE_W, Inches(0.04))
    stripe.fill.solid()
    stripe.fill.fore_color.rgb = C_GOLD
    stripe.line.fill.background()

    badge = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(0.45), Inches(0.18), Inches(2.25), Inches(0.46))
    badge.fill.solid()
    badge.fill.fore_color.rgb = C_NAVY_MED
    badge.line.color.rgb = C_GOLD
    badge.line.width = Pt(1)
    p_b = badge.text_frame.paragraphs[0]
    p_b.text = f"SLIDE {num:02d} | {badge_txt.upper()}"
    p_b.font.name = FONT_MAIN
    p_b.font.size = Pt(8.5)
    p_b.font.bold = True
    p_b.font.color.rgb = C_GOLD
    p_b.alignment = PP_ALIGN.CENTER

    tb = slide.shapes.add_textbox(Inches(2.85), Inches(0.14), Inches(10.0), Inches(0.56))
    tf = tb.text_frame
    tf.word_wrap = True
    p = tf.paragraphs[0]
    p.text = title_txt
    p.font.name = FONT_MAIN
    p.font.size = Pt(16)
    p.font.bold = True
    p.font.color.rgb = C_WHITE

    # Footer nhe
    ft = slide.shapes.add_textbox(Inches(0.45), Inches(7.18), Inches(12.4), Inches(0.24))
    pf = ft.text_frame.paragraphs[0]
    pf.text = "Học phần Khai phá dữ liệu  |  Nhóm 02 — Lớp Cao học CNTT 20.2  |  GVHD: TS. Bùi Ngọc Thăng"
    pf.font.name = FONT_MAIN
    pf.font.size = Pt(8.5)
    pf.font.color.rgb = C_TEXT_MUTED

    add_notes(slide, notes)


def add_card(slide, left, top, width, height, title=None, accent=C_NAVY_MED, bg=C_BG_CARD):
    card = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(left), Inches(top), Inches(width), Inches(height))
    card.fill.solid()
    card.fill.fore_color.rgb = bg
    card.line.color.rgb = C_BORDER
    card.line.width = Pt(1)

    if title:
        bar = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(left), Inches(top), Inches(width), Inches(0.36))
        bar.fill.solid()
        bar.fill.fore_color.rgb = accent
        bar.line.fill.background()
        p = bar.text_frame.paragraphs[0]
        p.text = f"  {title}"
        p.font.name = FONT_MAIN
        p.font.size = Pt(10.5)
        p.font.bold = True
        p.font.color.rgb = C_WHITE
        p.alignment = PP_ALIGN.LEFT
    return card


def add_bullets(slide, left, top, width, height, items, font_size=11.0, space_after=5):
    tb = slide.shapes.add_textbox(Inches(left), Inches(top), Inches(width), Inches(height))
    tf = tb.text_frame
    tf.word_wrap = True
    for i, item in enumerate(items):
        p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
        p.space_after = Pt(space_after)
        p.line_spacing = 1.18
        if isinstance(item, tuple):
            lead, body = item
            r1 = p.add_run()
            r1.text = f"- {lead}: "
            r1.font.name = FONT_MAIN
            r1.font.size = Pt(font_size)
            r1.font.bold = True
            r1.font.color.rgb = C_NAVY_DARK
            r2 = p.add_run()
            r2.text = body
            r2.font.name = FONT_MAIN
            r2.font.size = Pt(font_size)
            r2.font.color.rgb = C_TEXT_DARK
        else:
            r = p.add_run()
            r.text = f"- {item}"
            r.font.name = FONT_MAIN
            r.font.size = Pt(font_size)
            r.font.color.rgb = C_TEXT_DARK


def add_stat_box(slide, left, top, width, height, val_str, title_str, sub_str="", color=C_BLUE_ACCENT):
    add_card(slide, left, top, width, height)
    st = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(left), Inches(top), Inches(width), Inches(0.06))
    st.fill.solid()
    st.fill.fore_color.rgb = color
    st.line.fill.background()

    tb = slide.shapes.add_textbox(Inches(left + 0.08), Inches(top + 0.12), Inches(width - 0.16), Inches(height - 0.2))
    tf = tb.text_frame
    tf.word_wrap = True
    p1 = tf.paragraphs[0]
    p1.text = val_str
    p1.font.name = FONT_MAIN
    p1.font.size = Pt(20)
    p1.font.bold = True
    p1.font.color.rgb = color
    p1.alignment = PP_ALIGN.CENTER

    p2 = tf.add_paragraph()
    p2.text = title_str
    p2.font.name = FONT_MAIN
    p2.font.size = Pt(10)
    p2.font.bold = True
    p2.font.color.rgb = C_TEXT_DARK
    p2.alignment = PP_ALIGN.CENTER
    p2.space_before = Pt(3)

    if sub_str:
        p3 = tf.add_paragraph()
        p3.text = sub_str
        p3.font.name = FONT_MAIN
        p3.font.size = Pt(8.5)
        p3.font.color.rgb = C_TEXT_MUTED
        p3.alignment = PP_ALIGN.CENTER
        p3.space_before = Pt(2)


def add_table_styled(slide, left, top, width, height, headers, rows, col_ratios=None, font_size=9.5, highlight_row=None):
    n_rows = len(rows) + 1
    n_cols = len(headers)
    shape = slide.shapes.add_table(n_rows, n_cols, Inches(left), Inches(top), Inches(width), Inches(height))
    tbl = shape.table
    if col_ratios:
        tot = sum(col_ratios)
        for i, r in enumerate(col_ratios):
            tbl.columns[i].width = Inches(width * r / tot)

    for c_idx, h in enumerate(headers):
        cell = tbl.cell(0, c_idx)
        cell.fill.solid()
        cell.fill.fore_color.rgb = C_NAVY_MED
        cell.vertical_anchor = MSO_ANCHOR.MIDDLE
        set_cell_border(cell)
        p = cell.text_frame.paragraphs[0]
        p.text = h
        p.font.name = FONT_MAIN
        p.font.size = Pt(font_size)
        p.font.bold = True
        p.font.color.rgb = C_WHITE
        p.alignment = PP_ALIGN.CENTER

    for r_idx, row in enumerate(rows):
        is_hl = (highlight_row is not None and r_idx == highlight_row)
        bg_c = RGBColor(219, 234, 254) if is_hl else (C_ROW_ALT if r_idx % 2 == 1 else C_WHITE)
        for c_idx, val in enumerate(row):
            cell = tbl.cell(r_idx + 1, c_idx)
            cell.fill.solid()
            cell.fill.fore_color.rgb = bg_c
            cell.vertical_anchor = MSO_ANCHOR.MIDDLE
            set_cell_border(cell)
            p = cell.text_frame.paragraphs[0]
            p.text = str(val)
            p.font.name = FONT_MAIN
            p.font.size = Pt(font_size - 0.3)
            p.font.bold = is_hl or (c_idx == 0)
            p.font.color.rgb = C_NAVY_DARK if is_hl else C_TEXT_DARK
            p.alignment = PP_ALIGN.LEFT if c_idx == 0 else PP_ALIGN.CENTER


def add_picture_fit(slide, img_path, left, top, max_w, max_h):
    if not os.path.exists(img_path):
        return
    with Image.open(img_path) as im:
        w_px, h_px = im.size
    ratio = w_px / h_px
    box_ratio = max_w / max_h
    if ratio >= box_ratio:
        w = max_w
        h = max_w / ratio
    else:
        h = max_h
        w = max_h * ratio
    x = left + (max_w - w) / 2
    y = top + (max_h - h) / 2
    slide.shapes.add_picture(img_path, Inches(x), Inches(y), Inches(w), Inches(h))


def build_deck():
    with open(os.path.join(MODEL_DIR, "eda_stats.json"), "r", encoding="utf-8") as f:
        eda = json.load(f)
    with open(os.path.join(MODEL_DIR, "experiment_results.json"), "r", encoding="utf-8") as f:
        res = json.load(f)

    prs = Presentation()
    prs.slide_width = SLIDE_W
    prs.slide_height = SLIDE_H
    blank = prs.slide_layouts[6]

    best_name = res["best_model_name"]
    t_star = res["optimal_threshold"]
    roi = res["roi_benchmark"]
    errt = res["error_analysis"]["at_optimal_tstar"]
    ab_unw = res["ablation_study"]["Unweighted (T=0.50)"]
    ab_bal = res["ablation_study"]["Class-Weighted Balanced (T=0.50)"]
    ab_opt = res["ablation_study"][f"Cost-Sensitive Optimal (T*={t_star:.2f})"]

    # ==================== SLIDE 01: TRANG BIA ====================
    s1 = prs.slides.add_slide(blank)
    bg1 = s1.shapes.add_shape(MSO_SHAPE.RECTANGLE, 0, 0, SLIDE_W, SLIDE_H)
    bg1.fill.solid()
    bg1.fill.fore_color.rgb = C_NAVY_DARK
    bg1.line.fill.background()

    for y_bar in (0, 7.38):
        b = s1.shapes.add_shape(MSO_SHAPE.RECTANGLE, 0, Inches(y_bar), SLIDE_W, Inches(0.12))
        b.fill.solid()
        b.fill.fore_color.rgb = C_GOLD
        b.line.fill.background()

    if os.path.exists(LOGO_PATH):
        s1.shapes.add_picture(LOGO_PATH, Inches(0.85), Inches(0.26), Inches(1.45), Inches(0.95))

    tb_u = s1.shapes.add_textbox(Inches(2.45), Inches(0.24), Inches(10.0), Inches(1.0))
    tf_u = tb_u.text_frame
    tf_u.word_wrap = True
    p1 = tf_u.paragraphs[0]
    p1.text = "TRƯỜNG ĐẠI HỌC KINH DOANH VÀ CÔNG NGHỆ HÀ NỘI"
    p1.font.name = FONT_MAIN
    p1.font.size = Pt(15)
    p1.font.bold = True
    p1.font.color.rgb = RGBColor(254, 243, 199)
    p2 = tf_u.add_paragraph()
    p2.text = "VIỆN SAU ĐẠI HỌC  —  LỚP CAO HỌC CNTT 20.2  —  HỌC PHẦN: KHAI PHÁ DỮ LIỆU"
    p2.font.name = FONT_MAIN
    p2.font.size = Pt(11)
    p2.font.bold = True
    p2.font.color.rgb = RGBColor(191, 219, 254)
    p2.space_before = Pt(4)

    add_card(s1, 0.85, 1.48, 11.63, 3.65, bg=RGBColor(21, 46, 82))
    tb_t = s1.shapes.add_textbox(Inches(1.1), Inches(1.75), Inches(11.13), Inches(3.1))
    tf_t = tb_t.text_frame
    tf_t.word_wrap = True
    pt0 = tf_t.paragraphs[0]
    pt0.text = "BÁO CÁO BÀI TẬP LỚN — CHỦ ĐỀ 2: PHÂN LOẠI DỮ LIỆU TRONG CHĂM SÓC KHÁCH HÀNG"
    pt0.font.name = FONT_MAIN
    pt0.font.size = Pt(12.5)
    pt0.font.bold = True
    pt0.font.color.rgb = C_GOLD
    pt0.alignment = PP_ALIGN.CENTER

    pt1 = tf_t.add_paragraph()
    pt1.text = "“ỨNG DỤNG CÁC MÔ HÌNH PHÂN LOẠI DỮ LIỆU\nTRONG DỰ ĐOÁN NGUY CƠ RỜI BỎ DỊCH VỤ CỦA KHÁCH HÀNG”"
    pt1.font.name = FONT_MAIN
    pt1.font.size = Pt(22)
    pt1.font.bold = True
    pt1.font.color.rgb = C_WHITE
    pt1.alignment = PP_ALIGN.CENTER
    pt1.space_before = Pt(14)

    pt2 = tf_t.add_paragraph()
    pt2.text = "Thực nghiệm trên bộ dữ liệu IBM Telco Customer Churn (N = 7.043 thuê bao)"
    pt2.font.name = FONT_MAIN
    pt2.font.size = Pt(13)
    pt2.font.color.rgb = RGBColor(203, 213, 225)
    pt2.alignment = PP_ALIGN.CENTER
    pt2.space_before = Pt(14)

    tb_m = s1.shapes.add_textbox(Inches(0.85), Inches(5.40), Inches(11.63), Inches(1.75))
    tf_m = tb_m.text_frame
    tf_m.word_wrap = True
    pm1 = tf_m.paragraphs[0]
    pm1.text = "NHÓM THỰC HIỆN (NHÓM 02): Đỗ Ngọc Điền  •  Dương Tùng Lâm  •  Trần Thị Thu Hương  •  Phạm Văn Chiến  •  Vàng A Hứ"
    pm1.font.name = FONT_MAIN
    pm1.font.size = Pt(11.5)
    pm1.font.bold = True
    pm1.font.color.rgb = C_WHITE
    pm1.alignment = PP_ALIGN.CENTER

    pm2 = tf_m.add_paragraph()
    pm2.text = "GIẢNG VIÊN HƯỚNG DẪN: TS. Bùi Ngọc Thăng"
    pm2.font.name = FONT_MAIN
    pm2.font.size = Pt(13)
    pm2.font.bold = True
    pm2.font.color.rgb = C_GOLD
    pm2.alignment = PP_ALIGN.CENTER
    pm2.space_before = Pt(10)

    pm3 = tf_m.add_paragraph()
    pm3.text = "Hà Nội, 2026"
    pm3.font.name = FONT_MAIN
    pm3.font.size = Pt(11)
    pm3.font.color.rgb = RGBColor(148, 163, 184)
    pm3.alignment = PP_ALIGN.CENTER
    pm3.space_before = Pt(8)

    add_notes(s1, "Kính chào TS. Bùi Ngọc Thăng và các anh chị học viên lớp CNTT 20.2. Thay mặt Nhóm 02, tôi xin trình bày bài tập lớn học phần Khai phá dữ liệu với đề tài: Ứng dụng các mô hình phân loại dữ liệu trong dự đoán nguy cơ rời bỏ dịch vụ của khách hàng.")

    # ==================== SLIDE 02: NOI DUNG TRINH BAY ====================
    s2 = prs.slides.add_slide(blank)
    apply_frame(
        s2, 2, "Tổng quan", "Nội dung trình bày báo cáo (5 phần chính)",
        "Bài trình bày của nhóm gồm 5 phần bám sát quy trình Khai phá tri thức KDD: từ đặt vấn đề, phân tích và tiền xử lý dữ liệu, đối sánh 5 mô hình và tối ưu ngưỡng chi phí, rút trích tri thức luật CART và SHAP, cho đến phân tích lỗi, hiệu quả chi phí và demo ứng dụng."
    )
    sections = [
        ("01", "Đặt vấn đề & Quy trình KDD", "Bối cảnh bài toán Customer Churn, mục tiêu nghiên cứu và khung quy trình KDD / CRISP-DM áp dụng.", C_NAVY_MED),
        ("02", "Khám phá & Tiền xử lý dữ liệu", "Đặc tả 7.043 hồ sơ Telco, kiểm định Chi-Square & Mann-Whitney U, tạo 6 đặc trưng miền và pipeline chống rò rỉ.", C_BLUE_ACCENT),
        ("03", "Đối sánh 5 mô hình & Ngưỡng chi phí", "Thực nghiệm 5-Fold CV với Logistic Regression, CART, Random Forest, XGBoost, LightGBM và tối ưu ngưỡng T* = 0,21.", C_TEAL),
        ("04", "Rút trích tri thức & Phân tích lỗi", "Hệ luật IF-THEN từ cây quyết định, giải thích SHAP toàn cục/cục bộ và giải phẫu các ca dự đoán sai (FN/FP).", C_GOLD),
        ("05", "Hiệu quả kinh tế, Demo & Kết luận", "Đối chứng chi phí 4 chiến lược chăm sóc, minh họa ứng dụng Streamlit, đề xuất hành động CSKH và hướng mở rộng.", C_RED),
    ]
    for i, (num_s, tit_s, desc_s, col_s) in enumerate(sections):
        row = i // 3
        col = i % 3
        left = 0.55 + col * 4.15 if row == 0 else 2.60 + (i - 3) * 4.20
        top = 1.25 + row * 2.85
        add_card(s2, left, top, 3.85, 2.45, title=f"PHẦN {num_s}: {tit_s}", accent=col_s)
        add_bullets(s2, left + 0.15, top + 0.55, 3.55, 1.75, [desc_s], font_size=11.5)

    # ==================== SLIDE 03: DAT VAN DE & MUC TIEU ====================
    s3 = prs.slides.add_slide(blank)
    apply_frame(
        s3, 3, "Bối cảnh", "Đặt vấn đề và mục tiêu bài toán dự đoán Customer Churn",
        "Chi phí tìm khách mới cao gấp 5 đến 7 lần giữ khách cũ. Trong tập dữ liệu Telco, hơn một phần tư thuê bao (26,54%) rời bỏ dịch vụ. Mục tiêu của nhóm là xây dựng quy trình khai phá dữ liệu chặt chẽ, vừa dự đoán chính xác vừa giải thích được nguyên nhân để tối ưu chi phí giữ chân."
    )
    add_stat_box(s3, 0.55, 1.15, 2.9, 1.45, "7.043", "Thuê bao viễn thông", "IBM Telco Churn Dataset", C_NAVY_MED)
    add_stat_box(s3, 3.65, 1.15, 2.9, 1.45, "26,54%", "Tỷ lệ rời bỏ (Churn)", "1.869 Churn / 5.174 Ở lại", C_RED)
    add_stat_box(s3, 6.75, 1.15, 2.9, 1.45, "5 : 1", "Tỷ lệ chi phí C_FN : C_FP", "Bỏ sót $500 vs Ưu đãi $100", C_GOLD)
    add_stat_box(s3, 9.85, 1.15, 2.9, 1.45, "5 Mô hình", "Đối sánh có đối chứng", "2 Baselines + 3 Ensembles", C_TEAL)

    add_card(s3, 0.55, 2.85, 6.0, 4.05, title="Ba hạn chế của chăm sóc khách hàng truyền thống", accent=C_RED)
    add_bullets(s3, 0.72, 3.35, 5.65, 3.4, [
        ("Phản ứng bị động", "Chỉ biết khi khách hàng đã yêu cầu chấm dứt hợp đồng, khi đó tỷ lệ giữ chân rất thấp."),
        ("Luật lọc thủ công cứng nhắc", "Lọc đơn giản trên Excel (ví dụ cước cao + hợp đồng tháng) bỏ sót nhiều thuê bao rời bỏ (FN lớn)."),
        ("Chiến dịch ưu đãi dàn trải", "Tặng khuyến mãi đại trà cho toàn bộ khách hàng gây lãng phí ngân sách vào nhóm vốn đã trung thành."),
    ], font_size=11.5, space_after=10)

    add_card(s3, 6.80, 2.85, 5.95, 4.05, title="Mục tiêu bài tập lớn (Trọng tâm Khai phá dữ liệu)", accent=C_NAVY_MED)
    add_bullets(s3, 6.97, 3.35, 5.60, 3.4, [
        ("Quy trình KDD chống rò rỉ", "Tách tập kiểm thử độc lập trước mọi bước chuẩn hóa, kiểm thử chéo Stratified 5-Fold."),
        ("Tối ưu hóa nhạy chi phí", "Tìm ngưỡng quyết định T* tối thiểu hóa tổng chi phí nghiệp vụ thay vì dùng mặc định 0,50."),
        ("Rút trích tri thức giải thích được", "Kết hợp luật IF-THEN tường minh từ cây CART và giá trị Shapley (SHAP) toàn cục/cục bộ."),
        ("Chạy 100% cục bộ, tái lập được", "Mã nguồn Python mã nguồn mở, không phụ thuộc API bên ngoài, kèm ứng dụng Streamlit."),
    ], font_size=11.2, space_after=8)

    # ==================== SLIDE 04: QUY TRINH KDD ====================
    s4 = prs.slides.add_slide(blank)
    apply_frame(
        s4, 4, "Quy trình KDD", "Chu trình Khai phá dữ liệu (KDD / CRISP-DM) áp dụng trong đề tài",
        "Nhóm tuân thủ 5 giai đoạn chuẩn của quy trình KDD: chuẩn hóa dữ liệu gốc, kiểm định thống kê kết hợp tạo 6 đặc trưng miền, đóng gói tiền xử lý trong pipeline 5-Fold CV, đánh giá nhạy chi phí và rút trích tri thức luật kết hợp SHAP."
    )
    steps_kdd = [
        ("1. Dữ liệu & Làm sạch", "7.043 mẫu, 21 cột.\nLoại customerID.\nXử lý 11 dòng thiếu TotalCharges (tenure=0).", C_NAVY_MED),
        ("2. EDA & Đặc trưng", "Kiểm định Chi2 & Mann-Whitney U.\nTạo 6 đặc trưng miền (25 biến -> 55 chiều OHE).", C_BLUE_ACCENT),
        ("3. Pipeline & 5 Mô hình", "Stratified Split 80/20.\nColumnTransformer trong 5-Fold GridSearchCV.", C_TEAL),
        ("4. Đánh giá & T*", "So sánh AUC, PR-AUC, Brier.\nTìm ngưỡng T* = 0,21 trên OOF Train.", C_GOLD),
        ("5. Tri thức & Demo", "8 luật CART IF-THEN.\nSHAP Global/Local.\nError Analysis & App.", C_RED),
    ]
    for idx, (st_t, st_d, st_c) in enumerate(steps_kdd):
        x = 0.45 + idx * 2.52
        add_card(s4, x, 1.35, 2.34, 3.05, title=st_t, accent=st_c)
        add_bullets(s4, x + 0.10, 1.85, 2.14, 2.40, st_d.split("\n"), font_size=10.0, space_after=6)

    add_card(s4, 0.45, 4.70, 12.40, 2.25, title="Điểm nhấn kỹ thuật đảm bảo tính khách quan thực nghiệm", accent=C_NAVY_DARK)
    add_bullets(s4, 0.65, 5.20, 12.0, 1.65, [
        ("Tách tập Test độc lập ngay từ đầu", "Tập Test (1.409 mẫu) hoàn toàn không tham gia vào Imputer, StandardScaler, OneHotEncoder hay chọn siêu tham số."),
        ("Chọn mô hình và chọn ngưỡng T* trên OOF Train", "Tránh hoàn toàn thiên lệch lạc quan (optimistic bias) khi báo cáo hiệu năng và chi phí trên tập Test."),
        ("Đồng bộ dữ liệu xuyên suốt", "Toàn bộ số liệu trong Báo cáo Word, Slide thuyết trình và Web Demo được đọc trực tiếp từ kết quả chạy thực nghiệm (experiment_results.json)."),
    ], font_size=10.8, space_after=6)

    # ==================== SLIDE 05: EDA PHAN BO NHAN & YEU TO CHINH ====================
    s5 = prs.slides.add_slide(blank)
    apply_frame(
        s5, 5, "Khám phá dữ liệu", "Phân bố biến mục tiêu và các phân khúc có tỷ lệ rời bỏ cao",
        "Biểu đồ bên trái cho thấy tỷ lệ mất cân bằng lớp là 26,54% Churn. Hợp đồng theo tháng có tỷ lệ rời bỏ 42,7%, cao gấp 15 lần hợp đồng 2 năm. Hình bên phải chỉ rõ khách rời bỏ tập trung ở nhóm tenure ngắn, cước tháng cao, dùng cáp quang Fiber optic và thanh toán Electronic check."
    )
    add_card(s5, 0.45, 1.08, 6.1, 4.55, title="Phân bố Churn và tỷ lệ rời bỏ theo Hợp đồng", accent=C_NAVY_MED)
    add_picture_fit(s5, os.path.join(FIG_DIR, "fig01_churn_distribution.png"), 0.55, 1.52, 5.9, 4.0)

    add_card(s5, 6.75, 1.08, 6.1, 4.55, title="Phân phối theo tenure, cước, Internet và thanh toán", accent=C_NAVY_MED)
    add_picture_fit(s5, os.path.join(FIG_DIR, "fig02_eda_drivers.png"), 6.85, 1.52, 5.9, 4.0)

    add_card(s5, 0.45, 5.80, 12.4, 1.25)
    add_bullets(s5, 0.65, 5.92, 12.0, 1.05, [
        ("Hợp đồng & Vòng đời", "Hợp đồng theo tháng có 42,7% rời bỏ (so với 2,8% ở hợp đồng 2 năm); trung vị tenure nhóm rời bỏ chỉ là 10 tháng so với 38 tháng ở nhóm ở lại."),
        ("Dịch vụ & Thanh toán", "Cáp quang Fiber optic (41,9% Churn) và thanh toán thủ công Electronic check (45,3% Churn) là hai dấu hiệu cảnh báo rủi ro rõ nhất."),
    ], font_size=10.5, space_after=4)

    # ==================== SLIDE 06: KIEM DINH THONG KE & FEATURE ENGINEERING ====================
    s6 = prs.slides.add_slide(blank)
    apply_frame(
        s6, 6, "Đặc trưng", "Kiểm định thống kê và kỹ thuật tạo 6 đặc trưng tri thức miền",
        "Kiểm định Chi-Square và Cramér's V xác nhận Contract, TenureGroup, OnlineSecurity, TechSupport và InternetService liên kết rất mạnh với Churn (p < 10^-150). Ma trận tương quan chỉ ra đa cộng tuyến mạnh giữa tenure và TotalCharges (r = 0,83). Bên phải là 6 đặc trưng miền do nhóm thiết kế thêm."
    )
    add_card(s6, 0.45, 1.08, 7.25, 5.90, title="Tương quan Pearson (biến số) và Top 10 Cramér's V (biến phân loại)", accent=C_NAVY_MED)
    add_picture_fit(s6, os.path.join(FIG_DIR, "fig03_correlation_cramersv.png"), 0.55, 1.52, 7.05, 5.35)

    add_card(s6, 7.90, 1.08, 4.95, 5.90, title="6 đặc trưng tri thức miền được tạo mới", accent=C_TEAL)
    add_bullets(s6, 8.05, 1.58, 4.65, 5.25, [
        ("AvgMonthlySpend", "TotalCharges / max(tenure, 1) — đo chi tiêu bình quân thực tế (p = 2,1e-54; lọt Top 10 SHAP)."),
        ("ChargeDiff", "MonthlyCharges - AvgMonthlySpend — kiểm tra biến động tăng cước (đơn biến p = 0,69)."),
        ("TotalAddonServices", "Tổng số dịch vụ gia tăng đang đăng ký (0–6 dịch vụ, p = 4,7e-7)."),
        ("HasProtectionBundle", "Chỉ báo sử dụng gói OnlineSecurity hoặc TechSupport."),
        ("IsAutoPayment", "Phân tách thanh toán tự động (ngân hàng/thẻ) với thanh toán thủ công."),
        ("TenureGroup", "Phân 5 giai đoạn vòng đời (0-6, 7-12, 13-24, 25-48, >48m; Cramér's V = 0,3629 — hạng 2 toàn bộ)."),
    ], font_size=10.2, space_after=7)

    # ==================== SLIDE 07: ANTI-LEAKAGE PIPELINE & CV ====================
    s7 = prs.slides.add_slide(blank)
    apply_frame(
        s7, 7, "Tiền xử lý", "Kiến trúc Pipeline chống rò rỉ dữ liệu và Kiểm thử chéo 5-Fold",
        "Để đảm bảo kết quả đánh giá trung thực, nhóm tách tập Train 5.634 mẫu và Test 1.409 mẫu theo phân tầng nhãn ngay từ đầu. Mọi phép điền khuyết, chuẩn hóa và mã hóa One-Hot nằm trong ColumnTransformer, chỉ fit trên tập huấn luyện của từng fold."
    )
    add_stat_box(s7, 0.55, 1.15, 3.85, 1.35, "5.634 mẫu (80%)", "Tập Huấn luyện (Train)", "Tỷ lệ Churn: 26,54% (Stratified)", C_NAVY_MED)
    add_stat_box(s7, 4.75, 1.15, 3.85, 1.35, "1.409 mẫu (20%)", "Tập Kiểm thử độc lập (Test)", "Tỷ lệ Churn: 26,54% (374 ca Churn)", C_RED)
    add_stat_box(s7, 8.95, 1.15, 3.85, 1.35, "55 chiều", "Không gian đặc trưng sau OHE", "6 số + 3 nhị phân + 16 phân loại", C_TEAL)

    add_card(s7, 0.55, 2.75, 6.0, 4.20, title="Cấu trúc ColumnTransformer bên trong Pipeline", accent=C_BLUE_ACCENT)
    add_bullets(s7, 0.72, 3.25, 5.65, 3.55, [
        ("Nhánh biến số liên tục (6 biến)", "SimpleImputer(median) -> StandardScaler (đưa về trung bình 0, phương sai 1 để Logistic Regression hội tụ ổn định)."),
        ("Nhánh biến nhị phân (3 biến)", "SeniorCitizen, HasProtectionBundle, IsAutoPayment -> giữ nguyên giá trị 0/1."),
        ("Nhánh biến phân loại (16 biến)", "SimpleImputer(most_frequent) -> OneHotEncoder(handle_unknown='ignore') chống lỗi khi gặp giá trị mới."),
    ], font_size=11.0, space_after=10)

    add_card(s7, 6.80, 2.75, 6.0, 4.20, title="Chiến lược huấn luyện & chống quá khớp", accent=C_NAVY_MED)
    add_bullets(s7, 6.97, 3.25, 5.65, 3.55, [
        ("Stratified 5-Fold Cross-Validation", "Mỗi fold giữ đúng tỷ lệ 26,54% lớp dương; đo độ ổn định tổng quát hóa trước khi chạm tập Test."),
        ("GridSearchCV theo tiêu chí ROC-AUC", "Tối ưu siêu tham số theo năng lực xếp hạng xác suất thay vì Accuracy dễ bị lệch lớp."),
        ("Tối ưu ngưỡng T* trên OOF Predictions", "Tìm ngưỡng nhạy chi phí hoàn toàn từ dự đoán ngoài lát cắt trên tập Train, sau đó cố định áp dụng cho Test."),
    ], font_size=11.0, space_after=10)

    # ==================== SLIDE 08: CO SO 5 MO HINH ====================
    s8 = prs.slides.add_slide(blank)
    apply_frame(
        s8, 8, "Mô hình hóa", "Không gian 5 mô hình phân loại đối chứng (Baselines vs Ensembles)",
        "Nhóm đối sánh 5 thuật toán thuộc hai nhóm: 2 mô hình cơ sở theo đúng yêu cầu Chủ đề 2 là Logistic Regression và Decision Tree CART, cùng 3 mô hình học kết hợp cây là Random Forest, XGBoost và LightGBM."
    )
    add_card(s8, 0.55, 1.15, 5.95, 5.80, title="Nhóm mô hình cơ sở (Baselines)", accent=C_BLUE_ACCENT)
    add_bullets(s8, 0.72, 1.68, 5.60, 5.10, [
        ("1. Logistic Regression (Chính quy hóa L2)", "Hàm Sigmoid ánh xạ tuyến tính sang xác suất; tối ưu Binary Cross-Entropy. Cho phép đọc trực tiếp Odds Ratio = exp(w_j). Nhanh, xác suất hiệu chuẩn tốt."),
        ("2. Decision Tree (CART)", "Phân tách nhị phân giảm độ vấy bẩn Gini. Ưu điểm lớn nhất là chuyển trực tiếp nhánh cây thành luật IF-THEN tường minh; kiểm soát quá khớp bằng max_depth và min_samples_leaf."),
        ("Vai trò trong nghiên cứu", "Làm mốc đối chứng (benchmark) cả về độ chính xác lẫn khả năng diễn giải trực quan trước khi dùng mô hình phức tạp hơn."),
    ], font_size=11.2, space_after=12)

    add_card(s8, 6.80, 1.15, 6.0, 5.80, title="Nhóm mô hình học kết hợp cây (Tree Ensembles)", accent=C_TEAL)
    add_bullets(s8, 6.97, 1.68, 5.65, 5.10, [
        ("3. Random Forest (Bagging)", "Huấn luyện song song 250 cây trên các mẫu bootstrap và tập con đặc trưng ngẫu nhiên, giúp giảm mạnh phương sai so với cây đơn."),
        ("4. XGBoost (Gradient Boosting)", "Xây dựng cây tuần tự theo xấp xỉ Taylor bậc hai, bổ sung hàm phạt cấu trúc cây L1/L2 và lấy mẫu con (subsample = 0,8)."),
        ("5. LightGBM (Histogram GBDT)", "Rời rạc hóa đặc trưng theo histogram và tăng trưởng lá tối ưu (Leaf-wise với max_depth = 3, num_leaves = 15), tốc độ cao và độ hiệu chuẩn tốt."),
    ], font_size=11.2, space_after=12)

    # ==================== SLIDE 09: BANG DOI SANH 5 MO HINH ====================
    s9 = prs.slides.add_slide(blank)
    apply_frame(
        s9, 9, "Thực nghiệm", "Kết quả đối sánh 5 mô hình phân loại trên tập Test độc lập (N = 1.409)",
        "Trên kiểm thử chéo 5-Fold, LightGBM đạt CV ROC-AUC cao nhất là 0,8494 và PR-AUC cao nhất trên tập Test là 0,6657. XGBoost đứng sát ngay sau với Test ROC-AUC 0,8470. Đáng chú ý, sau khi có 6 đặc trưng miền, Logistic Regression cũng đạt 0,8452."
    )
    m_order = ["Logistic Regression", "Decision Tree (CART)", "Random Forest", "XGBoost", "LightGBM"]
    t_rows = []
    for nm in m_order:
        m05 = res["test_metrics_at_050"][nm]
        mt = res["test_metrics_at_tstar"][nm]
        t_rows.append([
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
    add_table_styled(
        s9, 0.45, 1.08, 12.43, 2.30,
        ["Mô hình", "CV AUC", "Acc (0.5)", "Rec (0.5)", "F1 (0.5)", "Test AUC", "PR-AUC", "Brier", f"Chi phí (T*={t_star})"],
        t_rows,
        col_ratios=[2.2, 0.9, 0.95, 0.95, 0.85, 0.95, 0.9, 0.85, 1.15],
        font_size=9.8,
        highlight_row=4,
    )
    add_card(s9, 0.45, 3.55, 12.43, 3.50, title="Trực quan hóa so sánh 5 chỉ số đánh giá chính", accent=C_NAVY_MED)
    add_picture_fit(s9, os.path.join(FIG_DIR, "fig04_model_comparison.png"), 0.60, 3.96, 12.13, 3.00)

    # ==================== SLIDE 10: ROC-AUC, PR-AUC & CONFUSION MATRIX ====================
    s10 = prs.slides.add_slide(blank)
    apply_frame(
        s10, 10, "Đánh giá", "Đường cong ROC-AUC, PR-AUC và Ma trận nhầm lẫn trên tập Test",
        "Hình trên cho thấy các mô hình tổ hợp cây (LightGBM, XGBoost, Random Forest) tách biệt rõ so với cây đơn lẻ trên cả đường cong ROC và PR. Hình dưới chỉ ra nghịch lý tại ngưỡng 0,50: LightGBM chỉ đạt Recall 51,6% (bỏ sót 181 ca), nhưng khi hạ ngưỡng về T* = 0,21, Recall tăng lên 84,8% (chỉ còn bỏ sót 57 ca)."
    )
    add_card(s10, 0.45, 1.05, 12.43, 3.05, title="Đường cong ROC và Precision-Recall của 5 mô hình", accent=C_NAVY_MED)
    add_picture_fit(s10, os.path.join(FIG_DIR, "fig05_roc_pr_curves.png"), 0.60, 1.44, 12.13, 2.58)

    add_card(s10, 0.45, 4.22, 12.43, 2.85, title="Ma trận nhầm lẫn so sánh các mô hình tại T = 0,50 và LightGBM tại T* = 0,21", accent=C_BLUE_ACCENT)
    add_picture_fit(s10, os.path.join(FIG_DIR, "fig06_confusion_matrices.png"), 0.55, 4.62, 12.23, 2.38)

    # ==================== SLIDE 11: TOI UU NGUONG CHI PHI & ABLATION ====================
    s11 = prs.slides.add_slide(blank)
    apply_frame(
        s11, 11, "Nhạy chi phí", "Đối chứng xử lý lệch lớp, Tối ưu ngưỡng T* và Hiệu quả kinh tế",
        "Đường cong chi phí trên OOF Train đạt đáy tại T* = 0,21, khớp rất sát ngưỡng lý thuyết Bayes 0,20. Bảng đối chứng cho thấy tối ưu ngưỡng trên xác suất chuẩn vừa giữ nguyên Brier Score tốt (0,1348 thay vì bị méo lên 0,1630 như balanced), vừa đạt tổng chi phí thấp nhất là $94.300 — tiết kiệm 49,6% so với không can thiệp."
    )
    add_card(s11, 0.45, 1.05, 12.43, 3.55, title="Đường cong chi phí theo ngưỡng T (A) và Đối chứng 4 chiến lược trên tập Test (B)", accent=C_TEAL)
    add_picture_fit(s11, os.path.join(FIG_DIR, "fig07_cost_threshold_roi.png"), 0.60, 1.46, 12.13, 3.06)

    ab_rows = [
        ["Mặc định Unweighted (T=0,50)", f"{ab_unw['accuracy']*100:.1f}%", f"{ab_unw['precision']*100:.1f}%", f"{ab_unw['recall']*100:.1f}%", f"{ab_unw['brier']:.4f}", str(ab_unw["fn"]), f"${ab_unw['total_cost_usd']:,.0f}"],
        ["Trọng số Balanced (T=0,50)", f"{ab_bal['accuracy']*100:.1f}%", f"{ab_bal['precision']*100:.1f}%", f"{ab_bal['recall']*100:.1f}%", f"{ab_bal['brier']:.4f}", str(ab_bal["fn"]), f"${ab_bal['total_cost_usd']:,.0f}"],
        [f"Tối ưu ngưỡng T* = {t_star:.2f}", f"{ab_opt['accuracy']*100:.1f}%", f"{ab_opt['precision']*100:.1f}%", f"{ab_opt['recall']*100:.1f}%", f"{ab_opt['brier']:.4f}", str(ab_opt["fn"]), f"${ab_opt['total_cost_usd']:,.0f}"],
    ]
    add_table_styled(
        s11, 0.45, 4.78, 12.43, 1.65,
        ["Phương pháp trên LightGBM", "Accuracy", "Precision", "Recall", "Brier Score", "Bỏ sót (FN)", "Tổng chi phí"],
        ab_rows,
        col_ratios=[2.7, 1.0, 1.0, 1.0, 1.1, 1.0, 1.2],
        font_size=10.0,
        highlight_row=2,
    )
    tb_note = s11.shapes.add_textbox(Inches(0.50), Inches(6.50), Inches(12.3), Inches(0.55))
    pn = tb_note.text_frame.paragraphs[0]
    pn.text = f"Hiệu quả kinh tế (N=1.409): LightGBM (T*=0,21) giảm chi phí từ $187.000 xuống $94.300 (-49,6%) và rẻ hơn luật lọc Excel ($123.100) thêm $28.800 (-23,4%)."
    pn.font.name = FONT_MAIN
    pn.font.size = Pt(10.5)
    pn.font.bold = True
    pn.font.color.rgb = C_TEAL

    # ==================== SLIDE 12: RUT TRICH LUAT CART ====================
    s12 = prs.slides.add_slide(blank)
    apply_frame(
        s12, 12, "Luật quyết định", "Rút trích hệ luật IF-THEN tường minh từ cây quyết định CART",
        "Từ cây quyết định cắt tỉa độ sâu 3 trên dữ liệu gốc, nhóm trích xuất được các luật kinh doanh trực quan. Đáng chú ý nhất là Luật R1: khách dùng hợp đồng theo tháng, cáp quang Fiber optic và thời gian gắn bó dưới 13,5 tháng chiếm 13,5% tập khách nhưng có tỷ lệ rời bỏ lên tới 71,0%."
    )
    add_card(s12, 0.45, 1.05, 12.43, 3.65, title="Sơ đồ cây quyết định CART độ sâu 3 (tập huấn luyện N = 5.634)", accent=C_NAVY_MED)
    add_picture_fit(s12, os.path.join(FIG_DIR, "fig08_decision_tree_rules.png"), 0.55, 1.44, 12.23, 3.18)

    r_top = [
        ["R1 (Rất cao)", "Hợp đồng = Theo tháng AND Internet = Fiber optic AND tenure <= 13,5 tháng", "762 (13,5%)", "71,0%"],
        ["R2 (Cao)", "Hợp đồng = Theo tháng AND Internet != Fiber optic AND tenure <= 3,5 tháng", "468 (8,3%)", "43,2%"],
        ["R3 (Cao)", "Hợp đồng = Theo tháng AND Internet = Fiber optic AND tenure > 13,5 tháng", "945 (16,8%)", "42,2%"],
        ["R8 (An toàn)", "Hợp đồng 1-2 năm AND MonthlyCharges <= $93,67 AND có dịch vụ bảo mật", "1.505 (26,7%)", "2,4%"],
    ]
    add_table_styled(
        s12, 0.45, 4.82, 12.43, 2.20,
        ["Phân khúc luật", "Điều kiện phân tách (IF ...)", "Số mẫu (Support)", "Tỷ lệ Churn"],
        r_top,
        col_ratios=[1.4, 5.2, 1.4, 1.1],
        font_size=9.8,
        highlight_row=0,
    )

    # ==================== SLIDE 13: SHAP TOAN CUC ====================
    s13 = prs.slides.add_slide(blank)
    apply_frame(
        s13, 13, "SHAP Toàn cục", "Giải thích mô hình LightGBM toàn cục với SHAP Summary Plot",
        "Biểu đồ SHAP Beeswarm cho thấy thứ tự quan trọng và chiều hướng tác động của 12 đặc trưng hàng đầu. Hợp đồng theo tháng đẩy xác suất rời bỏ mạnh nhất (mean |SHAP| = 0,663), trong khi thời gian gắn bó dài kéo mạnh xác suất về phía an toàn. Việc thiếu OnlineSecurity và TechSupport đứng thứ 3 và thứ 4."
    )
    add_card(s13, 0.45, 1.08, 8.05, 5.95, title="Biểu đồ SHAP Summary Plot — Top 12 đặc trưng (LightGBM)", accent=C_NAVY_MED)
    add_picture_fit(s13, os.path.join(FIG_DIR, "fig09_shap_summary.png"), 0.55, 1.50, 7.85, 5.42)

    add_card(s13, 8.70, 1.08, 4.18, 5.95, title="Nhận định từ tri thức SHAP", accent=C_GOLD)
    add_bullets(s13, 8.85, 1.58, 3.88, 5.25, [
        ("1. Contract_Month-to-month (0,663)", "Yếu tố chi phối số 1; không có ràng buộc hợp đồng làm tăng mạnh log-odds rời bỏ."),
        ("2. tenure (0,357)", "Quan hệ nghịch rõ rệt: khách mới (xanh) rủi ro cao, khách lâu năm (đỏ) rất ổn định."),
        ("3–4. Thiếu Security & TechSupport", "Khách chưa dùng hai dịch vụ này có xác suất rời bỏ cao hơn rõ rệt."),
        ("5–6. Fiber optic & Electronic check", "Cần rà soát chất lượng/giá gói cáp quang và trải nghiệm thanh toán thủ công."),
        ("Đóng góp đặc trưng miền", "AvgMonthlySpend (hạng 10) xác nhận hiệu quả của bước Feature Engineering."),
    ], font_size=10.3, space_after=8)

    # ==================== SLIDE 14: SHAP CUC BO ====================
    s14 = prs.slides.add_slide(blank)
    apply_frame(
        s14, 14, "SHAP Cục bộ", "Giải thích quyết định cho từng cá nhân bằng SHAP Cục bộ",
        "Ở cấp độ từng thuê bao, SHAP giúp nhân viên CSKH hiểu chính xác vì sao một khách hàng bị cảnh báo. Khách 5178-LMXOP có xác suất rời bỏ 90,7% do mới dùng 1 tháng, hợp đồng tháng và thiếu hỗ trợ kỹ thuật; ngược lại khách 7606-BPHHN có xác suất chỉ 0,6% nhờ gắn bó 72 tháng và hợp đồng 2 năm."
    )
    add_card(s14, 0.45, 1.08, 12.43, 4.55, title="Đối chiếu đóng góp SHAP cục bộ trên 2 hồ sơ điển hình thuộc tập Test", accent=C_NAVY_MED)
    add_picture_fit(s14, os.path.join(FIG_DIR, "fig10_shap_waterfall_cases.png"), 0.60, 1.50, 12.13, 4.02)

    add_card(s14, 0.45, 5.78, 6.08, 1.25, bg=RGBColor(254, 242, 242))
    add_bullets(s14, 0.60, 5.88, 5.78, 1.05, [
        ("Ca rủi ro cao (5178-LMXOP — 90,7%)", "tenure = 1 tháng (+0,777), Month-to-month (+0,770), TechSupport_No (+0,287), cước $95,10 -> Ưu tiên gọi CSKH ngay."),
    ], font_size=10.0)

    add_card(s14, 6.75, 5.78, 6.13, 1.25, bg=RGBColor(240, 253, 244))
    add_bullets(s14, 6.90, 5.88, 5.83, 1.05, [
        ("Ca trung thành (7606-BPHHN — 0,6%)", "tenure = 72 tháng (-0,967), không dùng hợp đồng tháng (-0,696), hợp đồng 2 năm (-0,303) -> Không cần gửi ưu đãi tốn kém."),
    ], font_size=10.0)

    # ==================== SLIDE 15: PHAN TICH LOI ====================
    s15 = prs.slides.add_slide(blank)
    apply_frame(
        s15, 15, "Phân tích lỗi", "Phân tích chuyên sâu các ca dự đoán sai (Error Analysis)",
        "Giải phẫu 398 ca sai tại ngưỡng T* = 0,21 cho thấy 34,2% lỗi nằm ngay dải biên sát ngưỡng. Đáng chú ý, trong 57 ca bỏ sót còn lại, tới 71,9% là khách lâu năm hoặc hợp đồng dài đột ngột rời bỏ do yếu tố ngoại sinh chưa có trong dữ liệu tĩnh; còn 99,1% ca FP thực chất đều là nhóm hồ sơ rủi ro cao."
    )
    err_rows = [
        ["Số ca bỏ sót khách rời bỏ (FN)", "181 ca", f"{errt['fn_count']} ca (giảm 68,5%)"],
        ["  - Khách dài hạn / Hợp đồng 1-2 năm (FN)", "92 ca (50,8%)", f"{errt['fn_loyal_silent']} ca (71,9% số FN)"],
        ["  - Khách ngắn hạn hợp đồng tháng (FN)", "89 ca (49,2%)", f"{errt['fn_short_volatile']} ca (28,1% số FN)"],
        ["Số ca cảnh báo dư (FP)", "91 ca", f"{errt['fp_count']} ca"],
        ["  - Thuộc hồ sơ rủi ro cao (Tháng / Fiber)", "91 ca (100%)", f"{errt['fp_high_risk_profile']} ca (99,1% số FP)"],
        ["Lỗi sát vùng biên (|p - T| <= 0,10)", "93 ca (34,2%)", f"{errt['boundary_errors']} ca ({errt['boundary_share_pct']:.1f}% tổng lỗi)"],
    ]
    add_table_styled(
        s15, 0.45, 1.15, 7.25, 3.65,
        ["Nhóm lỗi trên tập Test (N = 1.409)", "Ngưỡng T = 0,50", f"Ngưỡng T* = {t_star:.2f}"],
        err_rows,
        col_ratios=[3.6, 1.6, 2.0],
        font_size=10.0,
    )

    add_card(s15, 7.95, 1.15, 4.93, 5.80, title="Nguyên nhân gốc rễ & Hàm ý cải thiện", accent=C_RED)
    add_bullets(s15, 8.10, 1.65, 4.63, 5.15, [
        ("Giới hạn dữ liệu tĩnh (Snapshot)", "41/57 ca bỏ sót là khách dài hạn rời bỏ đột xuất — thường do sự cố mạng, khiếu nại hoặc chuyển nơi ở chưa được ghi nhận trong tập dữ liệu."),
        ("Bản chất hợp lý của nhóm FP", "99,1% ca cảnh báo dư đều mang đúng đặc trưng rủi ro cao (Month-to-month / Fiber optic), nên chăm sóc sớm vẫn có giá trị phòng ngừa."),
        ("Vùng biên xác suất (34,2%)", "Với dải xác suất 0,21–0,40, nên áp dụng kênh giữ chân chi phí thấp (SMS/Email) thay vì gọi điện trực tiếp."),
    ], font_size=10.5, space_after=10)

    add_card(s15, 0.45, 5.05, 7.25, 1.90, title="Đề xuất bổ sung dữ liệu cho giai đoạn sau", accent=C_NAVY_MED)
    add_bullets(s15, 0.62, 5.52, 6.90, 1.30, [
        ("Nhật ký tương tác CSKH", "Số lần gọi tổng đài khiếu nại, thời gian chờ xử lý sự cố trong 90 ngày gần nhất."),
        ("Chuỗi thời gian hành vi", "Xu hướng giảm lưu lượng sử dụng Internet/thoại và số lần chậm thanh toán hóa đơn."),
    ], font_size=10.5, space_after=6)

    # ==================== SLIDE 16: DEMO STREAMLIT & DE XUAT CSKH ====================
    s16 = prs.slides.add_slide(blank)
    apply_frame(
        s16, 16, "Ứng dụng & CRM", "Nguyên mẫu ứng dụng Streamlit và Đề xuất giải pháp CSKH",
        "Nhóm đã đóng gói mô hình LightGBM và giải thích SHAP vào ứng dụng Streamlit chạy 100% cục bộ gồm 3 phân hệ. Bên phải là 4 nhóm giải pháp chăm sóc khách hàng gắn trực tiếp với từng quy luật dữ liệu phát hiện được."
    )
    add_card(s16, 0.45, 1.12, 5.75, 5.82, title="Ứng dụng minh họa Streamlit (app.py — 100% Offline)", accent=C_BLUE_ACCENT)
    add_bullets(s16, 0.62, 1.65, 5.40, 5.10, [
        ("Tab 1 — Dự đoán khách hàng cá nhân", "Nhập hồ sơ thuê bao, tính tự động 6 đặc trưng miền, hiển thị xác suất so với ngưỡng T* = 0,21, vẽ biểu đồ SHAP cục bộ và đề xuất kịch bản CSKH."),
        ("Tab 2 — Chấm điểm hàng loạt (CSV)", "Tải lên tệp CSV khách hàng (có sẵn mẫu 100 KH tập Test), tự động phân nhóm rủi ro và xuất tệp kết quả CSV cho bộ phận CSKH."),
        ("Tab 3 — Tổng quan kết quả thực nghiệm", "Xem bảng so sánh 5 mô hình, đối chứng chi phí 4 chiến lược và các biểu đồ phân tích."),
        ("Cách chạy nhanh", "Nhấp đúp tệp run_demo.bat hoặc chạy lệnh: python -m streamlit run app.py (không cần API key)."),
    ], font_size=10.8, space_after=10)

    add_card(s16, 6.45, 1.12, 6.43, 5.82, title="Chiến lược giữ chân khách hàng dựa trên tri thức khai phá", accent=C_TEAL)
    add_bullets(s16, 6.62, 1.65, 6.08, 5.10, [
        ("1. Chuyển đổi hợp đồng sớm (Nhóm Luật R1 — 71% Churn)", "Với khách hợp đồng tháng + Cáp quang mới dùng <= 13 tháng: tặng gói chuyển đổi hợp đồng 1 năm (giảm 10–15% cước 3 tháng đầu)."),
        ("2. Đóng gói dịch vụ bảo vệ cho thuê bao Cáp quang", "Tặng dùng thử 3 tháng miễn phí OnlineSecurity & TechSupport để tăng trải nghiệm và chi phí chuyển đổi."),
        ("3. Dịch chuyển khỏi Electronic check (45,3% Churn)", "Ưu đãi nhỏ trên hóa đơn kỳ đầu cho khách chuyển sang thanh toán tự động qua ngân hàng hoặc thẻ tín dụng."),
        ("4. Phân tầng chi phí chăm sóc theo xác suất", "Xác suất > 0,40: gọi điện tư vấn trực tiếp; xác suất 0,21–0,40: gửi khảo sát và ưu đãi tự động qua SMS/Email."),
    ], font_size=10.6, space_after=10)

    # ==================== SLIDE 17: KET LUAN & PHU LUC ====================
    s17 = prs.slides.add_slide(blank)
    apply_frame(
        s17, 17, "Kết luận", "Tổng kết đóng góp, Hướng phát triển và Phụ lục siêu tham số",
        "Tóm lại, nhóm đã hoàn thiện đầy đủ chu trình Khai phá dữ liệu phân loại trên dữ liệu thực, chứng minh hiệu quả của tối ưu ngưỡng nhạy chi phí và rút trích tri thức giải thích được. Cảm ơn TS. Bùi Ngọc Thăng và các anh chị đã lắng nghe; nhóm xin phép nhận câu hỏi và demo trực tiếp."
    )
    add_card(s17, 0.45, 1.10, 6.10, 3.65, title="Tổng kết kết quả đạt được", accent=C_NAVY_MED)
    add_bullets(s17, 0.62, 1.58, 5.75, 3.05, [
        ("Pipeline KDD chuẩn mực", "Chống rò rỉ dữ liệu tuyệt đối, kiểm định thống kê đầy đủ và 6 đặc trưng miền."),
        ("Hiệu năng & Chi phí tối ưu", "LightGBM đạt CV ROC-AUC 0,8494, PR-AUC 0,6657; ngưỡng T* = 0,21 nâng Recall lên 84,8% và giảm 49,6% tổng chi phí."),
        ("Minh bạch tri thức", "8 luật CART tường minh + SHAP toàn cục/cục bộ + giải phẫu nguyên nhân lỗi chi tiết."),
    ], font_size=10.6, space_after=8)

    add_card(s17, 6.75, 1.10, 6.13, 3.65, title="Hướng phát triển tiếp theo", accent=C_GOLD)
    add_bullets(s17, 6.92, 1.58, 5.78, 3.05, [
        ("Dữ liệu chuỗi thời gian", "Tích hợp lịch sử khiếu nại CSKH và biến động lưu lượng sử dụng theo tháng để giảm nhóm FN dài hạn."),
        ("Phân tích sống còn (Survival Analysis)", "Dự báo đồng thời xác suất và thời điểm dự kiến rời bỏ để chọn đúng thời điểm can thiệp."),
        ("Kết hợp phân cụm (Clustering)", "Phân khúc khách hàng theo RFM trước khi phân loại để cá nhân hóa ngân sách giữ chân theo CLV."),
    ], font_size=10.6, space_after=8)

    hp_rows = [
        [nm, ", ".join(f"{k}={v}" for k, v in res["best_params"][nm].items()), f"{res['cv_roc_auc'][nm]:.4f}"]
        for nm in m_order
    ]
    add_table_styled(
        s17, 0.45, 4.92, 12.43, 2.15,
        ["[Phụ lục] Mô hình", "Bộ siêu tham số tối ưu sau Stratified 5-Fold GridSearchCV", "CV ROC-AUC"],
        hp_rows,
        col_ratios=[2.3, 6.5, 1.2],
        font_size=9.3,
    )

    os.makedirs(REPORT_DIR, exist_ok=True)
    prs.save(OUTPUT_PPTX)
    print(f"Presentation generated successfully: {OUTPUT_PPTX} ({len(prs.slides)} slides)")


if __name__ == "__main__":
    build_deck()
