# -*- coding: utf-8 -*-
"""
build_presentation.py
Sinh Slide thuyet trinh PowerPoint (.pptx) 19 trang theo dung Theme & Style cua 'Thiet ke mau.pptx'
(Canvas 20.0 x 11.25 in, nen toi #021229, nhan vang #FFDE59, trang #FFFFFF, do hoa vector/nen tu mau)
ket hop cau truc bang phan cong, menu 5 chuong va bo cuc bang bieu tham khao tu HeTriThucTinHocTHPT.
Hoc phan: Khai pha du lieu (Data Mining) - Nhom 02 - Lop Cao hoc CNTT 20.2 - GVHD: TS. Bui Ngoc Thang.
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
THEME_DIR = os.path.join(ASSET_DIR, "template_theme")
MODEL_DIR = os.path.join(BASE_DIR, "models")
REPORT_DIR = os.path.join(BASE_DIR, "reports")
FIG_DIR = os.path.join(REPORT_DIR, "figures")
OUTPUT_PPTX = os.path.join(REPORT_DIR, "Slide_Thuyet_Trinh_Khai_Pha_Du_Lieu_Nhom02.pptx")

LOGO_HUBT = os.path.join(ASSET_DIR, "logo_hubt_transparent.png")
IMG_COVER_BG = os.path.join(THEME_DIR, "image1.jpeg")
IMG_HERO_ART = os.path.join(THEME_DIR, "image2.png")
IMG_EMBLEM = os.path.join(THEME_DIR, "image4.png")
IMG_CIRCUIT = os.path.join(THEME_DIR, "image7.png")

# ==================== THEME CONSTANTS (TU THIET KE MAU.PPTX) ====================
SLIDE_W = Inches(20.000)
SLIDE_H = Inches(11.250)

FONT_DISPLAY = "Bahnschrift"   # Giong Anton/Montserrat dam, ho tro day du dau tieng Viet tren Windows
FONT_BODY = "Segoe UI"         # Hien dai, de doc tren nen toi, du dau tieng Viet
FONT_CODE = "Consolas"

C_BG_DARK = RGBColor(2, 18, 41)        # #021229 - Nen chinh cua Thiet ke mau.pptx
C_CARD_DARK = RGBColor(10, 31, 61)     # #0A1F3D - Nen card tren nen toi
C_CARD_ALT = RGBColor(15, 42, 80)      # #0F2A50 - Dong bang xen ke / card nhan
C_CARD_HL = RGBColor(26, 58, 108)      # #1A3A6C - Dong highlight trong bang
C_BORDER = RGBColor(38, 72, 118)       # #264876 - Vien card/bang tinh te
C_YELLOW = RGBColor(255, 222, 89)      # #FFDE59 - Mau nhan chinh cua Thiet ke mau.pptx
C_BLACK = RGBColor(0, 0, 2)            # #000002 - Chu tren nen vang
C_WHITE = RGBColor(255, 255, 255)      # #FFFFFF - Chu chinh tren nen toi
C_MUTED = RGBColor(186, 204, 226)      # #BACCE2 - Chu phu tren nen toi
C_TEAL = RGBColor(45, 212, 191)        # #2DD4BF - Nhan tich cuc / an toan
C_RED = RGBColor(248, 113, 113)        # #F87171 - Nhan canh bao / rui ro cao
C_BLUE_SOFT = RGBColor(96, 165, 250)   # #60A5FA - Nhan phu xanh duong


def set_slide_bg(slide, color_rgb):
    bg = slide.background
    fill = bg.fill
    fill.solid()
    fill.fore_color.rgb = color_rgb


def set_cell_border(cell, border_color="264876", sz="8"):
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


def add_yellow_pill(slide, left, top, width, height, text, font_size=15):
    pill = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(left), Inches(top), Inches(width), Inches(height))
    pill.fill.solid()
    pill.fill.fore_color.rgb = C_YELLOW
    pill.line.fill.background()
    tf = pill.text_frame
    tf.word_wrap = False
    p = tf.paragraphs[0]
    p.text = text
    p.font.name = FONT_DISPLAY
    p.font.size = Pt(font_size)
    p.font.bold = True
    p.font.color.rgb = C_BLACK
    p.alignment = PP_ALIGN.CENTER
    return pill


def apply_template_frame(slide, page_num, pill_label, title_yellow, title_white="", notes="", show_circuit=True):
    """Ap dung khung chuan cua Thiet ke mau.pptx (20 x 11.25 in, #021229, #FFDE59)."""
    set_slide_bg(slide, C_BG_DARK)

    # Hoa tiet mach vang goc duoi trai dac trung cua mau (dat lop duoi cung)
    if show_circuit and os.path.exists(IMG_CIRCUIT):
        slide.shapes.add_picture(IMG_CIRCUIT, Inches(-1.20), Inches(8.15), Inches(4.80), Inches(2.82))

    # Top-left brand (thay cho SALFORD & CO.)
    if os.path.exists(IMG_EMBLEM):
        slide.shapes.add_picture(IMG_EMBLEM, Inches(0.85), Inches(0.52), Inches(0.58), Inches(0.56))
    tb_brand = slide.shapes.add_textbox(Inches(1.58), Inches(0.58), Inches(6.5), Inches(0.45))
    pb = tb_brand.text_frame.paragraphs[0]
    pb.text = "NHÓM 02  |  CAO HỌC CNTT 20.2 — HUBT"
    pb.font.name = FONT_DISPLAY
    pb.font.size = Pt(15)
    pb.font.bold = True
    pb.font.color.rgb = C_WHITE

    # Top-right Page XX (giu dung vi tri cua mau)
    tb_pg = slide.shapes.add_textbox(Inches(16.90), Inches(0.55), Inches(2.25), Inches(0.45))
    pp = tb_pg.text_frame.paragraphs[0]
    pp.text = f"Page {page_num:02d}"
    pp.font.name = FONT_DISPLAY
    pp.font.size = Pt(17)
    pp.font.bold = True
    pp.font.color.rgb = C_YELLOW
    pp.alignment = PP_ALIGN.RIGHT

    # Tieu de 2 mau dac trung cua mau (#FFDE59 + #FFFFFF) + duong gach chan vang
    tb_ttl = slide.shapes.add_textbox(Inches(0.85), Inches(1.22), Inches(18.30), Inches(0.82))
    tf_t = tb_ttl.text_frame
    tf_t.word_wrap = True
    pt = tf_t.paragraphs[0]
    r1 = pt.add_run()
    r1.text = title_yellow.upper()
    r1.font.name = FONT_DISPLAY
    r1.font.size = Pt(31)
    r1.font.bold = True
    r1.font.color.rgb = C_YELLOW
    if title_white:
        r2 = pt.add_run()
        r2.text = f"  {title_white.upper()}"
        r2.font.name = FONT_DISPLAY
        r2.font.size = Pt(31)
        r2.font.bold = True
        r2.font.color.rgb = C_WHITE

    line = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(0.85), Inches(2.06), Inches(18.30), Inches(0.035))
    line.fill.solid()
    line.fill.fore_color.rgb = C_YELLOW
    line.line.fill.background()

    # Bottom-left Yellow Pill (thay cho nut "Next Page" cua mau)
    add_yellow_pill(slide, 0.85, 10.22, 3.30, 0.60, pill_label.upper(), font_size=14)

    # Bottom-right credit (thay cho Presentation By Benjamin Shah)
    tb_ft = slide.shapes.add_textbox(Inches(10.50), Inches(10.35), Inches(8.65), Inches(0.40))
    pf = tb_ft.text_frame.paragraphs[0]
    pf.text = "Học phần Khai phá dữ liệu  •  GVHD: TS. Bùi Ngọc Thăng"
    pf.font.name = FONT_BODY
    pf.font.size = Pt(13.5)
    pf.font.color.rgb = C_MUTED
    pf.alignment = PP_ALIGN.RIGHT

    add_notes(slide, notes)


def add_dark_card(slide, left, top, width, height, title=None, accent=C_YELLOW, bg=C_CARD_DARK):
    card = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(left), Inches(top), Inches(width), Inches(height))
    card.fill.solid()
    card.fill.fore_color.rgb = bg
    card.line.color.rgb = C_BORDER
    card.line.width = Pt(1.25)

    if title:
        # Thanh nhan tren cung card
        bar = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(left), Inches(top), Inches(width), Inches(0.08))
        bar.fill.solid()
        bar.fill.fore_color.rgb = accent
        bar.line.fill.background()

        tb = slide.shapes.add_textbox(Inches(left + 0.25), Inches(top + 0.20), Inches(width - 0.50), Inches(0.48))
        tf = tb.text_frame
        tf.word_wrap = True
        p = tf.paragraphs[0]
        p.text = title
        p.font.name = FONT_DISPLAY
        p.font.size = Pt(17)
        p.font.bold = True
        p.font.color.rgb = accent
    return card


def add_chart_card(slide, left, top, width, height, img_path, title=None):
    """Khung sang bao boc bieu do 300 DPI de hien thi sac net tren nen #021229."""
    outer = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(left), Inches(top), Inches(width), Inches(height))
    outer.fill.solid()
    outer.fill.fore_color.rgb = C_CARD_DARK
    outer.line.color.rgb = C_BORDER
    outer.line.width = Pt(1.25)

    inner_top = top + 0.14
    inner_h = height - 0.28
    if title:
        tb = slide.shapes.add_textbox(Inches(left + 0.22), Inches(top + 0.12), Inches(width - 0.44), Inches(0.42))
        p = tb.text_frame.paragraphs[0]
        p.text = title
        p.font.name = FONT_DISPLAY
        p.font.size = Pt(15.5)
        p.font.bold = True
        p.font.color.rgb = C_YELLOW
        inner_top = top + 0.56
        inner_h = height - 0.72

    canvas = slide.shapes.add_shape(
        MSO_SHAPE.ROUNDED_RECTANGLE,
        Inches(left + 0.16), Inches(inner_top), Inches(width - 0.32), Inches(inner_h)
    )
    canvas.fill.solid()
    canvas.fill.fore_color.rgb = C_WHITE
    canvas.line.fill.background()

    add_picture_fit(slide, img_path, left + 0.24, inner_top + 0.08, width - 0.48, inner_h - 0.16)


def add_bullets_dark(slide, left, top, width, height, items, font_size=15.0, space_after=9):
    tb = slide.shapes.add_textbox(Inches(left), Inches(top), Inches(width), Inches(height))
    tf = tb.text_frame
    tf.word_wrap = True
    for i, item in enumerate(items):
        p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
        p.space_after = Pt(space_after)
        p.line_spacing = 1.22
        if isinstance(item, tuple):
            lead, body = item
            r0 = p.add_run()
            r0.text = "▪  "
            r0.font.name = FONT_BODY
            r0.font.size = Pt(font_size)
            r0.font.bold = True
            r0.font.color.rgb = C_YELLOW

            r1 = p.add_run()
            r1.text = f"{lead}: "
            r1.font.name = FONT_BODY
            r1.font.size = Pt(font_size)
            r1.font.bold = True
            r1.font.color.rgb = C_WHITE

            r2 = p.add_run()
            r2.text = body
            r2.font.name = FONT_BODY
            r2.font.size = Pt(font_size)
            r2.font.color.rgb = C_MUTED
        else:
            r0 = p.add_run()
            r0.text = "▪  "
            r0.font.name = FONT_BODY
            r0.font.size = Pt(font_size)
            r0.font.bold = True
            r0.font.color.rgb = C_YELLOW

            r1 = p.add_run()
            r1.text = item
            r1.font.name = FONT_BODY
            r1.font.size = Pt(font_size)
            r1.font.color.rgb = C_WHITE


def add_kpi_tile(slide, left, top, width, height, num_str, title_str, sub_str="", num_color=C_YELLOW):
    add_dark_card(slide, left, top, width, height)
    bar = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(left), Inches(top), Inches(width), Inches(0.08))
    bar.fill.solid()
    bar.fill.fore_color.rgb = num_color
    bar.line.fill.background()

    tb = slide.shapes.add_textbox(Inches(left + 0.15), Inches(top + 0.22), Inches(width - 0.30), Inches(height - 0.30))
    tf = tb.text_frame
    tf.word_wrap = True
    p1 = tf.paragraphs[0]
    p1.text = num_str
    p1.font.name = FONT_DISPLAY
    p1.font.size = Pt(38)
    p1.font.bold = True
    p1.font.color.rgb = num_color
    p1.alignment = PP_ALIGN.CENTER

    p2 = tf.add_paragraph()
    p2.text = title_str
    p2.font.name = FONT_DISPLAY
    p2.font.size = Pt(15)
    p2.font.bold = True
    p2.font.color.rgb = C_WHITE
    p2.alignment = PP_ALIGN.CENTER
    p2.space_before = Pt(4)

    if sub_str:
        p3 = tf.add_paragraph()
        p3.text = sub_str
        p3.font.name = FONT_BODY
        p3.font.size = Pt(12.5)
        p3.font.color.rgb = C_MUTED
        p3.alignment = PP_ALIGN.CENTER
        p3.space_before = Pt(3)


def add_dark_table(slide, left, top, width, height, headers, rows, col_ratios=None, font_size=13.5, highlight_row=None, left_align_cols=(0,)):
    n_rows = len(rows) + 1
    n_cols = len(headers)
    shape = slide.shapes.add_table(n_rows, n_cols, Inches(left), Inches(top), Inches(width), Inches(height))
    tbl = shape.table
    if col_ratios:
        tot = sum(col_ratios)
        for i, r in enumerate(col_ratios):
            tbl.columns[i].width = Inches(width * r / tot)

    # Header row: Yellow #FFDE59 + Black Bold text (chuan phong cach Thiet ke mau.pptx)
    for c_idx, h in enumerate(headers):
        cell = tbl.cell(0, c_idx)
        cell.fill.solid()
        cell.fill.fore_color.rgb = C_YELLOW
        cell.vertical_anchor = MSO_ANCHOR.MIDDLE
        set_cell_border(cell, border_color="D4AF37", sz="8")
        p = cell.text_frame.paragraphs[0]
        p.text = h
        p.font.name = FONT_DISPLAY
        p.font.size = Pt(font_size)
        p.font.bold = True
        p.font.color.rgb = C_BLACK
        p.alignment = PP_ALIGN.CENTER

    for r_idx, row in enumerate(rows):
        is_hl = (highlight_row is not None and r_idx == highlight_row)
        bg = C_CARD_HL if is_hl else (C_CARD_ALT if r_idx % 2 == 1 else C_CARD_DARK)
        for c_idx, val in enumerate(row):
            cell = tbl.cell(r_idx + 1, c_idx)
            cell.fill.solid()
            cell.fill.fore_color.rgb = bg
            cell.vertical_anchor = MSO_ANCHOR.MIDDLE
            set_cell_border(cell, border_color="264876", sz="6")
            p = cell.text_frame.paragraphs[0]
            p.text = str(val)
            p.font.name = FONT_BODY
            p.font.size = Pt(font_size - 0.5)
            p.font.bold = is_hl or (c_idx == 0)
            p.font.color.rgb = C_YELLOW if is_hl else (C_WHITE if c_idx in left_align_cols else C_MUTED)
            if is_hl and c_idx != 0:
                p.font.color.rgb = C_YELLOW
            elif c_idx == 0:
                p.font.color.rgb = C_YELLOW if is_hl else C_WHITE
            p.alignment = PP_ALIGN.LEFT if c_idx in left_align_cols else PP_ALIGN.CENTER


def build_deck():
    with open(os.path.join(MODEL_DIR, "eda_stats.json"), "r", encoding="utf-8") as f:
        eda = json.load(f)
    with open(os.path.join(MODEL_DIR, "experiment_results.json"), "r", encoding="utf-8") as f:
        res = json.load(f)

    prs = Presentation()
    prs.slide_width = SLIDE_W
    prs.slide_height = SLIDE_H
    blank = prs.slide_layouts[6]

    t_star = res["optimal_threshold"]
    roi = res["roi_benchmark"]
    errt = res["error_analysis"]["at_optimal_tstar"]
    ab_unw = res["ablation_study"]["Unweighted (T=0.50)"]
    ab_bal = res["ablation_study"]["Class-Weighted Balanced (T=0.50)"]
    ab_opt = res["ablation_study"][f"Cost-Sensitive Optimal (T*={t_star:.2f})"]
    m_order = ["Logistic Regression", "Decision Tree (CART)", "Random Forest", "XGBoost", "LightGBM"]

    # ==================== SLIDE 01: TRANG BIA (THEO SLIDE 01 THIET KE MAU) ====================
    s1 = prs.slides.add_slide(blank)
    set_slide_bg(s1, C_BG_DARK)
    if os.path.exists(IMG_COVER_BG):
        s1.shapes.add_picture(IMG_COVER_BG, Inches(0), Inches(0), SLIDE_W, SLIDE_H)
    if os.path.exists(IMG_HERO_ART):
        add_picture_fit(s1, IMG_HERO_ART, 13.35, 2.55, 5.85, 5.75)

    # Top-left Logo HUBT + Ten truong
    if os.path.exists(LOGO_HUBT):
        s1.shapes.add_picture(LOGO_HUBT, Inches(1.05), Inches(0.65), Inches(1.35), Inches(0.90))
    tb_uni = s1.shapes.add_textbox(Inches(2.55), Inches(0.68), Inches(11.5), Inches(0.95))
    tf_u = tb_uni.text_frame
    tf_u.word_wrap = True
    pu1 = tf_u.paragraphs[0]
    pu1.text = "TRƯỜNG ĐẠI HỌC KINH DOANH VÀ CÔNG NGHỆ HÀ NỘI"
    pu1.font.name = FONT_DISPLAY
    pu1.font.size = Pt(18)
    pu1.font.bold = True
    pu1.font.color.rgb = C_YELLOW
    pu2 = tf_u.add_paragraph()
    pu2.text = "VIỆN SAU ĐẠI HỌC  •  LỚP CAO HỌC CNTT 20.2  •  HỌC PHẦN: KHAI PHÁ DỮ LIỆU"
    pu2.font.name = FONT_BODY
    pu2.font.size = Pt(14)
    pu2.font.bold = True
    pu2.font.color.rgb = C_WHITE
    pu2.space_before = Pt(3)

    # Main Bold Display Title (giong vi tri PROGRAMMING cua Slide 01 mau)
    tb_hero = s1.shapes.add_textbox(Inches(1.05), Inches(2.25), Inches(12.0), Inches(3.25))
    tf_h = tb_hero.text_frame
    tf_h.word_wrap = True
    ph0 = tf_h.paragraphs[0]
    ph0.text = "CHỦ ĐỀ 2: PHÂN LOẠI DỮ LIỆU TRONG CHĂM SÓC KHÁCH HÀNG"
    ph0.font.name = FONT_DISPLAY
    ph0.font.size = Pt(17)
    ph0.font.bold = True
    ph0.font.color.rgb = C_YELLOW

    ph1 = tf_h.add_paragraph()
    ph1.text = "DỰ ĐOÁN NGUY CƠ RỜI BỎ\nDỊCH VỤ KHÁCH HÀNG"
    ph1.font.name = FONT_DISPLAY
    ph1.font.size = Pt(54)
    ph1.font.bold = True
    ph1.font.color.rgb = C_WHITE
    ph1.space_before = Pt(10)

    # Mo ta ngan (vi tri TextBox 9 cua Slide 01 mau)
    tb_sub = s1.shapes.add_textbox(Inches(1.05), Inches(5.65), Inches(11.8), Inches(1.45))
    tf_s = tb_sub.text_frame
    tf_s.word_wrap = True
    ps1 = tf_s.paragraphs[0]
    ps1.text = (
        "Ứng dụng các mô hình phân loại dữ liệu (Logistic Regression, CART, Random Forest, XGBoost, LightGBM), "
        "tối ưu ngưỡng nhạy chi phí và giải thích SHAP trên bộ dữ liệu IBM Telco Customer Churn (7.043 thuê bao)."
    )
    ps1.font.name = FONT_BODY
    ps1.font.size = Pt(17)
    ps1.font.color.rgb = C_MUTED
    ps1.line_spacing = 1.28

    # Nut vang dac trung cua Slide 01 mau
    add_yellow_pill(s1, 1.05, 7.50, 4.65, 0.78, "NHÓM 02  —  LỚP CNTT 20.2", font_size=18)

    # Thanh vien & GVHD (vi tri duoi cua Slide 01 mau)
    tb_mem = s1.shapes.add_textbox(Inches(1.05), Inches(8.62), Inches(11.8), Inches(1.30))
    tf_m = tb_mem.text_frame
    tf_m.word_wrap = True
    pm1 = tf_m.paragraphs[0]
    pm1.text = "Thành viên: Đỗ Ngọc Điền  •  Dương Tùng Lâm  •  Trần Thị Thu Hương  •  Phạm Văn Chiến  •  Vàng A Hứ"
    pm1.font.name = FONT_BODY
    pm1.font.size = Pt(14.5)
    pm1.font.bold = True
    pm1.font.color.rgb = C_WHITE
    pm2 = tf_m.add_paragraph()
    pm2.text = "github.com/Lamdt198/KhaiPhaDuLieu-CustomerChurn"
    pm2.font.name = FONT_CODE
    pm2.font.size = Pt(13.5)
    pm2.font.color.rgb = C_YELLOW
    pm2.space_before = Pt(8)

    tb_gv = s1.shapes.add_textbox(Inches(13.10), Inches(9.75), Inches(6.05), Inches(0.60))
    pgv = tb_gv.text_frame.paragraphs[0]
    pgv.text = "Giảng viên hướng dẫn: TS. Bùi Ngọc Thăng"
    pgv.font.name = FONT_DISPLAY
    pgv.font.size = Pt(17.5)
    pgv.font.bold = True
    pgv.font.color.rgb = C_YELLOW
    pgv.alignment = PP_ALIGN.RIGHT

    add_notes(s1, "Kính chào TS. Bùi Ngọc Thăng và các anh chị học viên lớp Cao học CNTT 20.2. Thay mặt Nhóm 02, tôi xin trình bày Bài tập lớn học phần Khai phá dữ liệu: Ứng dụng các mô hình phân loại dữ liệu trong dự đoán nguy cơ rời bỏ dịch vụ của khách hàng.")

    # ==================== SLIDE 02: BANG PHAN CONG NHIEM VU (THAM KHAO SLIDE 02 HE TRI THUC) ====================
    s2 = prs.slides.add_slide(blank)
    apply_template_frame(
        s2, 2, "Mở đầu", "BẢNG PHÂN CÔNG NHIỆM VỤ", "NHÓM 02 — LỚP CNTT 20.2",
        "Bài tập lớn được phân công rõ ràng cho 5 thành viên Nhóm 02 theo từng công đoạn của quy trình Khai phá dữ liệu, đảm bảo sự thống nhất giữa mã nguồn thực nghiệm, báo cáo Word và slide bảo vệ."
    )
    add_dark_table(
        s2, 0.85, 2.45, 18.30, 5.90,
        ["STT", "Họ và tên", "Vai trò", "Nhiệm vụ chuyên môn đảm nhiệm", "Đóng góp"],
        [
            ["01", "Đỗ Ngọc Điền", "Nhóm trưởng", "Xây dựng khung đề tài theo quy trình KDD, viết Chương 1 & Kết luận, tổng hợp và rà soát toàn bộ báo cáo.", "20%"],
            ["02", "Dương Tùng Lâm", "Thành viên", "Khám phá dữ liệu (EDA), kiểm định Chi-Square & Mann-Whitney U, tạo 6 đặc trưng miền và viết Chương 2.", "20%"],
            ["03", "Trần Thị Thu Hương", "Thành viên", "Thiết kế pipeline chống rò rỉ, huấn luyện & tối ưu 5 mô hình qua Stratified 5-Fold CV, viết Chương 3.", "20%"],
            ["04", "Phạm Văn Chiến", "Thành viên", "Tối ưu ngưỡng chi phí T*, rút trích luật CART, giải thích SHAP toàn cục/cục bộ và phân tích lỗi (Chương 4).", "20%"],
            ["05", "Vàng A Hứ", "Thành viên", "Đối chứng chi phí 4 chiến lược vận hành, xây dựng ứng dụng minh họa Streamlit, viết Chương 5 & Slide.", "20%"],
        ],
        col_ratios=[0.7, 2.3, 1.5, 6.2, 1.1],
        font_size=15.0,
        left_align_cols=(1, 3),
    )
    add_dark_card(s2, 0.85, 8.65, 18.30, 1.25, bg=C_CARD_ALT)
    add_bullets_dark(s2, 1.15, 8.88, 17.70, 0.85, [
        ("Cam kết thực nghiệm", "100% mã nguồn Python mã nguồn mở chạy cục bộ, không dùng API Key; toàn bộ số liệu trong Báo cáo, Slide và Demo đều đọc trực tiếp từ kết quả chạy thật trên bộ dữ liệu IBM Telco Churn (N = 7.043)."),
    ], font_size=14.5)

    # ==================== SLIDE 03: MENU / KHUNG BAO CAO 5 CHUONG (THAM KHAO SLIDE 03 HE TRI THUC + SLIDE 06/07 MAU) ====================
    s3 = prs.slides.add_slide(blank)
    apply_template_frame(
        s3, 3, "Tổng quan", "KHUNG BÁO CÁO 5 CHƯƠNG", "& LỘ TRÌNH THUYẾT TRÌNH",
        "Báo cáo được tổ chức thành 5 chương chính cùng phần Kết luận, bám sát quy trình chuẩn KDD từ đặt vấn đề, phân tích và làm sạch dữ liệu, huấn luyện đối sánh 5 mô hình, rút trích tri thức luật CART & SHAP cho tới phân tích lỗi, đánh giá hiệu quả chi phí và minh họa ứng dụng."
    )
    menu_items = [
        ("01", "CHƯƠNG 1: TỔNG QUAN & LÝ THUYẾT", "Bối cảnh Customer Churn, quy trình KDD/CRISP-DM, cơ sở toán học 5 thuật toán phân loại, học nhạy chi phí và SHAP.", C_YELLOW),
        ("02", "CHƯƠNG 2: EDA & TIỀN XỬ LÝ", "Đặc tả 7.043 thuê bao Telco, kiểm định Chi-Square & Mann-Whitney U, tạo 6 đặc trưng miền và pipeline chống rò rỉ.", C_BLUE_SOFT),
        ("03", "CHƯƠNG 3: THỰC NGHIỆM 5 MÔ HÌNH", "Đối sánh Logistic Regression, CART, Random Forest, XGBoost, LightGBM qua 5-Fold CV và tối ưu ngưỡng T* = 0,21.", C_TEAL),
        ("04", "CHƯƠNG 4: TRI THỨC, LỖI & ROI", "Rút trích 8 luật IF-THEN từ CART, giải thích SHAP toàn cục/cục bộ, phân tích lỗi FN/FP và đối chứng 4 kịch bản chi phí.", C_YELLOW),
        ("05", "CHƯƠNG 5: ỨNG DỤNG & GIẢI PHÁP", "Nguyên mẫu web minh họa Streamlit 3 phân hệ (chạy cục bộ) và đề xuất 4 nhóm hành động CSKH theo rủi ro khai phá.", C_BLUE_SOFT),
        ("06", "KẾT LUẬN & PHỤ LỤC", "Tổng kết đóng góp khoa học, hạn chế dữ liệu tĩnh snapshot, hướng mở rộng Survival Analysis và bảng siêu tham số.", C_TEAL),
    ]
    for i, (num_s, tit_s, desc_s, acc_c) in enumerate(menu_items):
        row = i // 2
        col = i % 2
        x = 0.85 + col * 9.35
        y = 2.40 + row * 2.52
        add_dark_card(s3, x, y, 8.95, 2.25, bg=C_CARD_DARK)
        # So thu tu lon kieu Slide 06/07 Thiet ke mau.pptx
        tb_n = s3.shapes.add_textbox(Inches(x + 0.28), Inches(y + 0.35), Inches(1.35), Inches(1.45))
        pn = tb_n.text_frame.paragraphs[0]
        pn.text = num_s
        pn.font.name = FONT_DISPLAY
        pn.font.size = Pt(46)
        pn.font.bold = True
        pn.font.color.rgb = acc_c

        # Duong doc ngan cach
        vline = s3.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(x + 1.68), Inches(y + 0.32), Inches(0.04), Inches(1.60))
        vline.fill.solid()
        vline.fill.fore_color.rgb = acc_c
        vline.line.fill.background()

        tb_c = s3.shapes.add_textbox(Inches(x + 1.95), Inches(y + 0.25), Inches(6.75), Inches(1.80))
        tf_c = tb_c.text_frame
        tf_c.word_wrap = True
        pc1 = tf_c.paragraphs[0]
        pc1.text = tit_s
        pc1.font.name = FONT_DISPLAY
        pc1.font.size = Pt(17)
        pc1.font.bold = True
        pc1.font.color.rgb = C_WHITE
        pc2 = tf_c.add_paragraph()
        pc2.text = desc_s
        pc2.font.name = FONT_BODY
        pc2.font.size = Pt(13.8)
        pc2.font.color.rgb = C_MUTED
        pc2.space_before = Pt(6)
        pc2.line_spacing = 1.2

    # ==================== SLIDE 04: CHUONG 1 - DAT VAN DE & MUC TIEU ====================
    s4 = prs.slides.add_slide(blank)
    apply_template_frame(
        s4, 4, "Chương 1", "ĐẶT VẤN ĐỀ & MỤC TIÊU", "BÀI TOÁN CUSTOMER CHURN",
        "Chi phí thu hút khách mới cao gấp 5 đến 7 lần giữ khách cũ. Trên bộ dữ liệu thực IBM Telco, 26,54% thuê bao rời bỏ dịch vụ. Mục tiêu của nhóm là xây dựng pipeline khai phá dữ liệu khép kín, tối ưu chi phí giữ chân và giải thích rõ nguyên nhân."
    )
    add_kpi_tile(s4, 0.85, 2.38, 4.35, 2.05, "7.043", "Thuê bao viễn thông", "IBM Telco Customer Churn", C_YELLOW)
    add_kpi_tile(s4, 5.50, 2.38, 4.35, 2.05, "26,54%", "Tỷ lệ rời bỏ (Churn)", "1.869 Churn / 5.174 Ở lại", C_RED)
    add_kpi_tile(s4, 10.15, 2.38, 4.35, 2.05, "5 : 1", "Tỷ lệ chi phí C_FN : C_FP", "Bỏ sót $500 vs Ưu đãi $100", C_TEAL)
    add_kpi_tile(s4, 14.80, 2.38, 4.35, 2.05, "5 Mô hình", "Đối sánh thực nghiệm", "2 Baselines + 3 Ensembles", C_BLUE_SOFT)

    add_dark_card(s4, 0.85, 4.78, 8.95, 5.10, title="BA ĐIỂM NGHẼN CỦA CHĂM SÓC KHÁCH HÀNG TRUYỀN THỐNG", accent=C_RED)
    add_bullets_dark(s4, 1.12, 5.55, 8.40, 4.10, [
        ("Phản ứng bị động", "Doanh nghiệp chỉ phát hiện khi thuê bao đã yêu cầu hủy hợp đồng, lúc này khả năng giữ chân rất thấp."),
        ("Luật lọc thủ công cứng nhắc", "Lọc theo quy tắc đơn giản trên Excel (cước cao + hợp đồng tháng) bỏ sót số lượng lớn khách rời bỏ (FN cao)."),
        ("Chiến dịch ưu đãi dàn trải", "Tặng khuyến mãi đại trà cho 100% tập khách hàng gây lãng phí ngân sách vào nhóm vốn đã trung thành."),
    ], font_size=15.2, space_after=16)

    add_dark_card(s4, 10.20, 4.78, 8.95, 5.10, title="MỤC TIÊU NGHIÊN CỨU (TRỌNG TÂM KHAI PHÁ DỮ LIỆU)", accent=C_YELLOW)
    add_bullets_dark(s4, 10.47, 5.55, 8.40, 4.10, [
        ("Tuân thủ chu trình KDD chống rò rỉ", "Tách tập kiểm thử độc lập trước mọi bước biến đổi, kiểm định thống kê và kiểm thử chéo 5-Fold."),
        ("Tối ưu quyết định nhạy chi phí", "Tìm ngưỡng xác suất T* trên OOF Train nhằm giảm tối đa tổng chi phí nghiệp vụ thay vì mặc định 0,50."),
        ("Rút trích tri thức giải thích được", "Kết hợp luật IF-THEN tường minh từ cây CART và giá trị Shapley (SHAP) toàn cục/cục bộ."),
        ("Định lượng hiệu quả vận hành", "Phân tích lỗi chuyên sâu FN/FP và đối chứng trực tiếp chi phí thực tế giữa 4 kịch bản CSKH."),
    ], font_size=15.0, space_after=12)

    # ==================== SLIDE 05: CHUONG 1 - QUY TRINH KDD ====================
    s5 = prs.slides.add_slide(blank)
    apply_template_frame(
        s5, 5, "Chương 1", "QUY TRÌNH KHAI PHÁ DỮ LIỆU", "KDD & CRISP-DM ÁP DỤNG",
        "Đề tài triển khai trọn vẹn 5 giai đoạn của quy trình KDD: chuẩn hóa dữ liệu gốc, kiểm định thống kê kết hợp tạo 6 đặc trưng miền, đóng gói tiền xử lý chống rò rỉ trong 5-Fold CV, tối ưu ngưỡng nhạy chi phí và rút trích tri thức giải thích được."
    )
    kdd_stages = [
        ("01", "Dữ liệu & Làm sạch", "7.043 mẫu, 21 cột.\nLoại customerID.\nXử lý 11 dòng thiếu TotalCharges (tenure = 0).", C_YELLOW),
        ("02", "EDA & Đặc trưng", "Kiểm định Chi2 & Mann-Whitney U.\nTạo 6 đặc trưng miền (25 biến -> 55 chiều OHE).", C_BLUE_SOFT),
        ("03", "5 Mô hình & CV", "Stratified Split 80/20.\nColumnTransformer bên trong 5-Fold GridSearchCV.", C_TEAL),
        ("04", "Đánh giá & T*", "Đo ROC-AUC, PR-AUC, Brier Score.\nTìm ngưỡng T* = 0,21 trên OOF Train.", C_YELLOW),
        ("05", "Tri thức & Demo", "8 luật CART IF-THEN.\nSHAP Global/Local.\nError Analysis & App.", C_RED),
    ]
    for i, (num_s, st_t, st_d, col_s) in enumerate(kdd_stages):
        x = 0.85 + i * 3.72
        add_dark_card(s5, x, 2.55, 3.42, 4.35, bg=C_CARD_DARK)
        bar = s5.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(x), Inches(2.55), Inches(3.42), Inches(0.10))
        bar.fill.solid()
        bar.fill.fore_color.rgb = col_s
        bar.line.fill.background()

        tb_n = s5.shapes.add_textbox(Inches(x + 0.22), Inches(2.80), Inches(2.98), Inches(0.90))
        pn = tb_n.text_frame.paragraphs[0]
        pn.text = num_s
        pn.font.name = FONT_DISPLAY
        pn.font.size = Pt(40)
        pn.font.bold = True
        pn.font.color.rgb = col_s

        tb_h = s5.shapes.add_textbox(Inches(x + 0.22), Inches(3.70), Inches(2.98), Inches(0.55))
        ph = tb_h.text_frame.paragraphs[0]
        ph.text = st_t
        ph.font.name = FONT_DISPLAY
        ph.font.size = Pt(17)
        ph.font.bold = True
        ph.font.color.rgb = C_WHITE

        add_bullets_dark(s5, x + 0.22, 4.35, 2.98, 2.40, st_d.split("\n"), font_size=13.5, space_after=8)

    add_dark_card(s5, 0.85, 7.25, 18.30, 2.65, title="NGUYÊN TẮC BẢO ĐẢM TÍNH KHÁCH QUAN KHOA HỌC", accent=C_YELLOW, bg=C_CARD_ALT)
    add_bullets_dark(s5, 1.15, 7.98, 17.70, 1.75, [
        ("Cô lập tuyệt đối tập kiểm thử (Hold-out Test N = 1.409)", "Tập Test không tham gia bất kỳ bước Impute, StandardScaler, OneHotEncoder, chọn siêu tham số hay chọn ngưỡng T* nào."),
        ("Đồng nhất số liệu toàn hệ thống", "Toàn bộ bảng số liệu và 10 biểu đồ 300 DPI trong Báo cáo Word, Slide và ứng dụng Streamlit đều xuất tự động từ một nguồn kết quả duy nhất."),
    ], font_size=14.8, space_after=10)

    # ==================== SLIDE 06: CHUONG 1 - CO SO TOAN HOC ====================
    s6 = prs.slides.add_slide(blank)
    apply_template_frame(
        s6, 6, "Chương 1", "CƠ SỞ TOÁN HỌC", "PHÂN LOẠI, NHẠY CHI PHÍ & SHAP",
        "Đề tài đối sánh hai nhóm thuật toán: mô hình cơ sở (Logistic Regression tối ưu Log-Loss, CART giảm độ vấy bẩn Gini) và tổ hợp cây (Random Forest Bagging, XGBoost và LightGBM Gradient Boosting). Đồng thời áp dụng ngưỡng quyết định Bayes nhạy chi phí và phân rã đóng góp cộng tính SHAP."
    )
    add_dark_card(s6, 0.85, 2.38, 8.95, 4.55, title="NHÓM CƠ SỞ (BASELINES) & TỔ HỢP CÂY (ENSEMBLES)", accent=C_BLUE_SOFT)
    add_bullets_dark(s6, 1.12, 3.12, 8.40, 3.65, [
        ("Logistic Regression", "P(Y=1|x) = 1 / (1 + exp(-(w^T x + b))), tối ưu Binary Cross-Entropy có phạt L2; diễn giải trực tiếp qua Odds Ratio = exp(w_j)."),
        ("Decision Tree (CART)", "Tối đa hóa độ giảm vấy bẩn Gini(D) = 1 - (p_0^2 + p_1^2); chuyển trực tiếp nhánh cây thành luật IF-THEN."),
        ("Random Forest (Bagging)", "Tổng hợp 250 cây trên mẫu bootstrap & tập con đặc trưng ngẫu nhiên, giảm phương sai rõ rệt."),
        ("XGBoost & LightGBM (GBDT)", "Xây dựng cây tuần tự theo khai triển Taylor bậc hai + chính quy hóa L1/L2; LightGBM tăng trưởng theo nút lá trên histogram."),
    ], font_size=14.2, space_after=10)

    add_dark_card(s6, 10.20, 2.38, 8.95, 4.55, title="HỌC NHẠY CHI PHÍ & GIÁ TRỊ SHAPLEY (SHAP)", accent=C_YELLOW)
    add_bullets_dark(s6, 10.47, 3.12, 8.40, 3.65, [
        ("Ma trận chi phí bất đối xứng", "C_TN = $0 | C_FP = $100 (ưu đãi dư) | C_TP = $100 (chi phí giữ chân) | C_FN = $500 (tổn thất mất khách)."),
        ("Ngưỡng tối ưu lý thuyết Bayes", "Can thiệp khi p*C_TP + (1-p)*C_FP <= p*C_FN  =>  T*_Bayes = C_FP / (C_FN - C_TP + C_FP) = 100 / 500 = 0,20."),
        ("Hiệu chuẩn xác suất (Brier Score)", "Brier = (1/N) * Sum(p_i - y_i)^2 — đánh giá xác suất dự báo có phản ánh đúng tần suất thực tế hay không."),
        ("Giải thích cộng tính SHAP", "f(x) = phi_0 + Sum_{j=1}^d phi_j; TreeSHAP tính chính xác đóng góp biên của từng thuộc tính lên log-odds rời bỏ."),
    ], font_size=14.2, space_after=10)

    add_dark_table(
        s6, 0.85, 7.25, 18.30, 2.60,
        ["Trạng thái thực tế \\ Dự đoán", "Dự đoán Ở lại (0 — Không can thiệp)", "Dự đoán Rời bỏ (1 — Gửi ưu đãi CSKH)"],
        [
            ["Thực tế Ở lại (0)", "TN: Chi phí = $0 (Quyết định tối ưu)", "FP: Chi phí = $100 (Lãng phí ngân sách ưu đãi)"],
            ["Thực tế Rời bỏ (1)", "FN: Chi phí = $500 (Tổn thất giá trị vòng đời CLV)", "TP: Chi phí = $100 (Chi phí giữ chân thành công)"],
        ],
        col_ratios=[2.8, 3.8, 3.8],
        font_size=14.5,
        left_align_cols=(0, 1, 2),
    )

    # ==================== SLIDE 07: CHUONG 2 - EDA PHAN BO NHAN & YEU TO CHINH ====================
    s7 = prs.slides.add_slide(blank)
    apply_template_frame(
        s7, 7, "Chương 2", "PHÂN TÍCH KHÁM PHÁ DỮ LIỆU (EDA)", "PHÂN BỐ CHURN & CÁC YẾU TỐ CHÍNH",
        "Hợp đồng theo tháng có tỷ lệ rời bỏ lên tới 42,7%, cao gấp 15 lần hợp đồng 2 năm (2,8%). Khách rời bỏ tập trung dày ở giai đoạn đầu (trung vị tenure 10 tháng so với 38 tháng), mức cước tháng cao (79,65 USD so với 64,43 USD), dùng cáp quang Fiber optic (41,9%) và thanh toán Electronic check (45,3%)."
    )
    add_chart_card(s7, 0.85, 2.35, 8.95, 5.75, os.path.join(FIG_DIR, "fig01_churn_distribution.png"), "A. Phân bố nhãn Churn và Tỷ lệ rời bỏ theo Hợp đồng")
    add_chart_card(s7, 10.20, 2.35, 8.95, 5.75, os.path.join(FIG_DIR, "fig02_eda_drivers.png"), "B. Phân phối theo tenure, cước, hạ tầng Internet & thanh toán")

    add_dark_card(s7, 0.85, 8.35, 18.30, 1.58, bg=C_CARD_ALT)
    add_bullets_dark(s7, 1.15, 8.52, 17.70, 1.25, [
        ("Hợp đồng & Thời gian gắn bó", "Hợp đồng theo tháng chiếm 42,7% Churn (n = 3.875) so với 11,3% ở hợp đồng 1 năm và 2,8% ở 2 năm; trung vị tenure nhóm Churn chỉ là 10 tháng."),
        ("Hạ tầng & Thanh toán", "Thuê bao cáp quang Fiber optic (41,9% Churn) và thanh toán thủ công Electronic check (45,3% Churn) là hai nhóm rủi ro rõ nét nhất."),
    ], font_size=14.0, space_after=6)

    # ==================== SLIDE 08: CHUONG 2 - KIEM DINH THONG KE & 6 DAC TRUNG MIEN ====================
    s8 = prs.slides.add_slide(blank)
    apply_template_frame(
        s8, 8, "Chương 2", "KIỂM ĐỊNH THỐNG KÊ", "& TẠO 6 ĐẶC TRƯNG TRI THỨC MIỀN",
        "Kiểm định Chi-Square xác nhận Contract (Cramér's V = 0,410), TenureGroup (0,363), OnlineSecurity (0,347) và TechSupport (0,343) liên kết mạnh với Churn. Ma trận tương quan chỉ ra đa cộng tuyến mạnh giữa tenure và TotalCharges (r = 0,83). Bên phải là bảng 6 đặc trưng miền tạo mới."
    )
    add_chart_card(s8, 0.85, 2.35, 9.65, 7.55, os.path.join(FIG_DIR, "fig03_correlation_cramersv.png"), "Ma trận tương quan Pearson (biến số) & Top 10 Cramér's V")

    add_dark_table(
        s8, 10.80, 2.35, 8.35, 5.15,
        ["Đặc trưng miền mới", "Công thức / Định nghĩa", "Kiểm định (p)"],
        [
            ["AvgMonthlySpend", "TotalCharges / max(tenure, 1)", "2,09e-54"],
            ["ChargeDiff", "MonthlyCharges - AvgMonthlySpend", "0,693 (k.y.n)"],
            ["TotalAddonServices", "Tổng dịch vụ gia tăng = Yes (0-6)", "4,74e-07"],
            ["HasProtectionBundle", "OnlineSecurity = Yes OR TechSupport = Yes", "2,49e-117"],
            ["IsAutoPayment", "PaymentMethod chứa 'automatic'", "1,02e-68"],
            ["TenureGroup", "5 nhóm: 0-6, 7-12, 13-24, 25-48, >48m", "2,00e-199"],
        ],
        col_ratios=[2.3, 3.6, 1.5],
        font_size=12.8,
        left_align_cols=(0, 1),
    )

    add_dark_card(s8, 10.80, 7.75, 8.35, 2.15, title="Ý NGHĨA TỪ KIỂM ĐỊNH THỐNG KÊ", accent=C_TEAL)
    add_bullets_dark(s8, 11.02, 8.42, 7.90, 1.35, [
        ("TenureGroup & AvgMonthlySpend", "TenureGroup đạt Cramér's V hạng 2 toàn tập (0,363); AvgMonthlySpend lọt Top 10 SHAP."),
        ("Đa cộng tuyến (r = 0,83)", "tenure tương quan mạnh với TotalCharges -> cần chính quy hóa L2 cho Logistic Regression."),
    ], font_size=13.0, space_after=5)

    # ==================== SLIDE 09: CHUONG 2 - ANTI-LEAKAGE PIPELINE (THAM KHAO SLIDE 15 HE TRI THUC) ====================
    s9 = prs.slides.add_slide(blank)
    apply_template_frame(
        s9, 9, "Chương 2", "QUY TRÌNH TIỀN XỬ LÝ 4 BƯỚC", "CHỐNG RÒ RỈ DỮ LIỆU (ANTI-LEAKAGE)",
        "Quy trình tiền xử lý gồm 4 bước tuần tự nhằm triệt tiêu rò rỉ dữ liệu: tách phân tầng 80/20 ngay từ đầu, làm sạch 11 ô khuyết TotalCharges có tenure bằng 0, đóng gói ColumnTransformer vào Pipeline và chỉ fit trong từng lát cắt của kiểm thử chéo 5-Fold."
    )
    pipe_steps = [
        ("BƯỚC 1", "Phân tách phân tầng độc lập (Stratified Split 80/20)", "Tách ngay 5.634 mẫu Train (26,54% Churn) và 1.409 mẫu Test (26,54% Churn, 374 ca Churn) với random_state = 42 trước mọi bước chuẩn hóa.", C_YELLOW),
        ("BƯỚC 2", "Làm sạch & Khắc phục giá trị khuyết thiếu", "Loại bỏ customerID (không mang thông tin dự báo); chuyển TotalCharges sang float64 và điền 0,0 cho 11 thuê bao mới có tenure = 0.", C_BLUE_SOFT),
        ("BƯỚC 3", "Đóng gói ColumnTransformer (25 biến -> 55 chiều)", "Nhánh số (6 biến): Median Imputer + StandardScaler | Nhánh nhị phân (3 biến): giữ 0/1 | Nhánh phân loại (16 biến): OneHotEncoder(handle_unknown='ignore').", C_TEAL),
        ("BƯỚC 4", "Huấn luyện Stratified 5-Fold CV & Tối ưu OOF", "Gói ColumnTransformer + Estimator trong sklearn Pipeline; fit() chỉ chạy trên 4/5 fold huấn luyện, chọn siêu tham số theo ROC-AUC và chọn T* trên dự đoán OOF.", C_RED),
    ]
    for i, (b_lbl, b_tit, b_desc, b_col) in enumerate(pipe_steps):
        y = 2.40 + i * 1.88
        add_dark_card(s9, 0.85, y, 18.30, 1.62, bg=C_CARD_DARK)
        pill = s9.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(1.12), Inches(y + 0.36), Inches(2.25), Inches(0.88))
        pill.fill.solid()
        pill.fill.fore_color.rgb = b_col
        pill.line.fill.background()
        pp = pill.text_frame.paragraphs[0]
        pp.text = b_lbl
        pp.font.name = FONT_DISPLAY
        pp.font.size = Pt(20)
        pp.font.bold = True
        pp.font.color.rgb = C_BLACK if b_col in (C_YELLOW, C_TEAL) else C_WHITE
        pp.alignment = PP_ALIGN.CENTER

        tb = s9.shapes.add_textbox(Inches(3.68), Inches(y + 0.22), Inches(15.20), Inches(1.22))
        tf = tb.text_frame
        tf.word_wrap = True
        p1 = tf.paragraphs[0]
        p1.text = b_tit
        p1.font.name = FONT_DISPLAY
        p1.font.size = Pt(18.5)
        p1.font.bold = True
        p1.font.color.rgb = C_WHITE
        p2 = tf.add_paragraph()
        p2.text = b_desc
        p2.font.name = FONT_BODY
        p2.font.size = Pt(14.2)
        p2.font.color.rgb = C_MUTED
        p2.space_before = Pt(4)

    # ==================== SLIDE 10: CHUONG 3 - BANG DOI SANH 5 MO HINH ====================
    s10 = prs.slides.add_slide(blank)
    apply_template_frame(
        s10, 10, "Chương 3", "KẾT QUẢ ĐỐI SÁNH 5 MÔ HÌNH", "TRÊN TẬP TEST ĐỘC LẬP (N = 1.409)",
        "Bảng trên tổng hợp đầy đủ 8 chỉ số đánh giá của 5 mô hình trên tập Test độc lập. LightGBM dẫn đầu về CV ROC-AUC (0,8494) và PR-AUC (0,6657), XGBoost sát nút với Test ROC-AUC 0,8470. Logistic Regression sau khi có 6 đặc trưng miền cũng đạt 0,8452 và vượt xa cây CART đơn lẻ (0,8249)."
    )
    comp_rows = []
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
    add_dark_table(
        s10, 0.85, 2.35, 18.30, 3.10,
        ["Mô hình", "CV AUC", "Acc (0.5)", "Rec (0.5)", "F1 (0.5)", "Test AUC", "PR-AUC", "Brier", f"Chi phí (T*={t_star})"],
        comp_rows,
        col_ratios=[2.5, 1.0, 1.05, 1.05, 0.95, 1.05, 1.0, 0.95, 1.35],
        font_size=13.5,
        highlight_row=4,
    )
    add_chart_card(s10, 0.85, 5.68, 18.30, 4.28, os.path.join(FIG_DIR, "fig04_model_comparison.png"), "Trực quan hóa so sánh 5 chỉ số đánh giá chính trên tập Test (N = 1.409)")

    # ==================== SLIDE 11: CHUONG 3 - ROC-AUC, PR-AUC & CONFUSION MATRIX ====================
    s11 = prs.slides.add_slide(blank)
    apply_template_frame(
        s11, 11, "Chương 3", "ĐƯỜNG CONG ROC, PR-AUC", "& MA TRẬN NHẦM LẪN (CONFUSION MATRIX)",
        "Trên cả đường cong ROC và Precision-Recall, nhóm tổ hợp cây (LightGBM, XGBoost, Random Forest) tách biệt rõ so với cây CART đơn lẻ. Ma trận nhầm lẫn chỉ ra tác động trực tiếp của ngưỡng T* = 0,21: số ca bỏ sót khách rời bỏ FN giảm mạnh từ 181 xuống chỉ còn 57 ca (Recall tăng từ 51,6% lên 84,8%)."
    )
    add_chart_card(s11, 0.85, 2.32, 18.30, 3.85, os.path.join(FIG_DIR, "fig05_roc_pr_curves.png"), "A. Đường cong ROC-AUC và Precision-Recall (PR-AUC) của 5 mô hình")
    add_chart_card(s11, 0.85, 6.35, 18.30, 3.60, os.path.join(FIG_DIR, "fig06_confusion_matrices.png"), "B. Ma trận nhầm lẫn trên tập Test (N = 1.409): So sánh T = 0,50 và ngưỡng tối ưu T* = 0,21")

    # ==================== SLIDE 12: CHUONG 3 - ABLATION & TOI UU NGUONG CHI PHI T* ====================
    s12 = prs.slides.add_slide(blank)
    apply_template_frame(
        s12, 12, "Chương 3", "ĐỐI CHỨNG XỬ LÝ LỆCH LỚP", "& TỐI ƯU NGƯỠNG NHẠY CHI PHÍ T*",
        "Đường cong chi phí trên OOF Train đạt đáy tại T* = 0,21, khớp rất sát ngưỡng lý thuyết Bayes 0,20. Bảng đối chứng chỉ ra rằng tối ưu ngưỡng trên xác suất chuẩn vượt trội hơn gán trọng số balanced: giữ nguyên Brier Score tốt nhất (0,1348 thay vì bị méo lên 0,1630), nâng Recall lên 84,8% và giảm tổng chi phí còn $94.300."
    )
    add_chart_card(s12, 0.85, 2.32, 18.30, 4.45, os.path.join(FIG_DIR, "fig07_cost_threshold_roi.png"), "Đường cong chi phí theo ngưỡng T (OOF Train vs Test) và Đối chứng 4 chiến lược vận hành")

    ab_rows = [
        ["Mặc định Unweighted (T = 0,50)", f"{ab_unw['accuracy']*100:.2f}%", f"{ab_unw['precision']*100:.2f}%", f"{ab_unw['recall']*100:.2f}%", f"{ab_unw['f1']:.4f}", f"{ab_unw['brier']:.4f}", str(ab_unw["fn"]), f"${ab_unw['total_cost_usd']:,.0f}"],
        ["Gán trọng số Balanced (T = 0,50)", f"{ab_bal['accuracy']*100:.2f}%", f"{ab_bal['precision']*100:.2f}%", f"{ab_bal['recall']*100:.2f}%", f"{ab_bal['f1']:.4f}", f"{ab_bal['brier']:.4f}", str(ab_bal["fn"]), f"${ab_bal['total_cost_usd']:,.0f}"],
        [f"Tối ưu ngưỡng chi phí (T* = {t_star:.2f})", f"{ab_opt['accuracy']*100:.2f}%", f"{ab_opt['precision']*100:.2f}%", f"{ab_opt['recall']*100:.2f}%", f"{ab_opt['f1']:.4f}", f"{ab_opt['brier']:.4f}", str(ab_opt["fn"]), f"${ab_opt['total_cost_usd']:,.0f}"],
    ]
    add_dark_table(
        s12, 0.85, 7.00, 18.30, 2.25,
        ["Phương pháp trên LightGBM (Ablation Study)", "Accuracy", "Precision", "Recall", "F1-Score", "Brier Score", "Bỏ sót (FN)", "Tổng chi phí"],
        ab_rows,
        col_ratios=[3.2, 1.05, 1.05, 1.05, 1.0, 1.15, 1.1, 1.35],
        font_size=13.5,
        highlight_row=2,
    )
    tb_k = s12.shapes.add_textbox(Inches(0.85), Inches(9.40), Inches(18.30), Inches(0.55))
    pk = tb_k.text_frame.paragraphs[0]
    pk.text = "Kết luận thực nghiệm: Tối ưu ngưỡng T* = 0,21 vừa bảo toàn độ hiệu chuẩn xác suất (Brier = 0,1348), vừa tiết kiệm 49,6% tổn thất ($92.700) trên 1.409 thuê bao tập Test."
    pk.font.name = FONT_DISPLAY
    pk.font.size = Pt(15)
    pk.font.bold = True
    pk.font.color.rgb = C_TEAL

    # ==================== SLIDE 13: CHUONG 4 - RUT TRICH LUAT CART ====================
    s13 = prs.slides.add_slide(blank)
    apply_template_frame(
        s13, 13, "Chương 4", "RÚT TRÍCH HỆ LUẬT IF-THEN", "TỪ CÂY QUYẾT ĐỊNH CART CẮT TỈA",
        "Từ cây quyết định CART cắt tỉa độ sâu 3 trên thuộc tính gốc, nhóm rút trích được 8 luật phân loại tường minh. Nổi bật nhất là Luật R1 (Hợp đồng tháng + Cáp quang Fiber + tenure <= 13,5 tháng) chiếm 13,5% mẫu nhưng có xác suất rời bỏ tới 71,0%, đối lập hoàn toàn với nhóm Luật R8 chỉ có 2,4%."
    )
    add_chart_card(s13, 0.85, 2.32, 18.30, 4.40, os.path.join(FIG_DIR, "fig08_decision_tree_rules.png"), "Sơ đồ cây quyết định CART độ sâu 3 (tập huấn luyện N = 5.634)")

    rule_rows = [
        ["R1 (Rất cao)", "Hợp đồng = Theo tháng AND Internet = Fiber optic AND tenure <= 13,5 tháng", "762 (13,5%)", "71,0%"],
        ["R2 (Cao)", "Hợp đồng = Theo tháng AND Internet != Fiber optic AND tenure <= 3,5 tháng", "468 (8,3%)", "43,2%"],
        ["R3 (Cao)", "Hợp đồng = Theo tháng AND Internet = Fiber optic AND tenure > 13,5 tháng", "945 (16,8%)", "42,2%"],
        ["R8 (An toàn)", "Hợp đồng 1-2 năm AND MonthlyCharges <= $93,67 AND có dịch vụ OnlineSecurity", "1.505 (26,7%)", "2,4%"],
    ]
    add_dark_table(
        s13, 0.85, 6.95, 18.30, 2.95,
        ["Mã luật & Mức độ", "Điều kiện phân tách tường minh (IF ...)", "Số mẫu (Support)", "Tỷ lệ Churn"],
        rule_rows,
        col_ratios=[1.8, 6.6, 1.7, 1.3],
        font_size=13.8,
        highlight_row=0,
        left_align_cols=(0, 1),
    )

    # ==================== SLIDE 14: CHUONG 4 - SHAP TOAN CUC ====================
    s14 = prs.slides.add_slide(blank)
    apply_template_frame(
        s14, 14, "Chương 4", "GIẢI THÍCH TOÀN CỤC (GLOBAL XAI)", "BIỂU ĐỒ SHAP SUMMARY PLOT — LIGHTGBM",
        "Biểu đồ SHAP Beeswarm minh họa cả độ lớn lẫn chiều hướng tác động của 12 đặc trưng quan trọng nhất trên mô hình LightGBM. Hợp đồng theo tháng là yếu tố chi phối mạnh nhất (mean |SHAP| = 0,663), theo sau là tenure (0,357), thiếu OnlineSecurity (0,238), thiếu TechSupport (0,201) và Fiber optic (0,200)."
    )
    add_chart_card(s14, 0.85, 2.35, 11.45, 7.55, os.path.join(FIG_DIR, "fig09_shap_summary.png"), "SHAP Summary Beeswarm Plot — Top 12 đặc trưng (N = 1.409 tập Test)")

    add_dark_card(s14, 12.60, 2.35, 6.55, 7.55, title="TOP YẾU TỐ CHI PHỐI HÀNH VI RỜI BỎ", accent=C_YELLOW)
    add_bullets_dark(s14, 12.88, 3.12, 6.00, 6.50, [
        ("1. Contract_Month-to-month (0,663)", "Không có ràng buộc hợp đồng đẩy mạnh nhất xác suất rời bỏ; hợp đồng 2 năm kéo mạnh về phía an toàn."),
        ("2. tenure (0,357)", "Quan hệ nghịch rõ rệt: khách mới (chấm xanh) có SHAP dương cao; khách lâu năm (chấm đỏ) rất ổn định."),
        ("3–4. OnlineSecurity_No & TechSupport_No", "Thiếu hai dịch vụ bảo vệ và hỗ trợ kỹ thuật xếp hạng 3 (0,238) và hạng 4 (0,201)."),
        ("5–6. Fiber optic & Electronic check", "Hạ tầng cáp quang (0,200) và thanh toán thủ công Electronic check (0,156) làm tăng rõ nguy cơ hủy gói."),
        ("Hiệu quả đặc trưng miền", "AvgMonthlySpend (0,097) lọt Top 10 yếu tố quan trọng nhất, kiểm chứng giá trị bước Feature Engineering."),
    ], font_size=14.2, space_after=14)

    # ==================== SLIDE 15: CHUONG 4 - SHAP CUC BO ====================
    s15 = prs.slides.add_slide(blank)
    apply_template_frame(
        s15, 15, "Chương 4", "GIẢI THÍCH CỤC BỘ (LOCAL XAI)", "PHÂN RÃ SHAP CHO TỪNG THUÊ BAO ĐIỂN HÌNH",
        "Ở cấp độ từng cá nhân, SHAP giải thích rõ nguyên nhân khiến mô hình đưa ra cảnh báo. Khách 5178-LMXOP có xác suất rời bỏ 90,7% do mới dùng 1 tháng, hợp đồng tháng và thiếu TechSupport; ngược lại khách 7606-BPHHN có xác suất chỉ 0,6% nhờ thời gian gắn bó 72 tháng và hợp đồng 2 năm."
    )
    add_chart_card(s15, 0.85, 2.32, 18.30, 5.45, os.path.join(FIG_DIR, "fig10_shap_waterfall_cases.png"), "So sánh đóng góp SHAP cục bộ giữa 1 ca rủi ro cao (A) và 1 ca trung thành (B) thuộc tập Test")

    add_dark_card(s15, 0.85, 8.02, 8.95, 1.92, title="CA RỦI RO CAO: MÃ KH 5178-LMXOP (90,7% — THỰC TẾ: RỜI BỎ)", accent=C_RED)
    add_bullets_dark(s15, 1.10, 8.65, 8.45, 1.15, [
        "tenure = 1 tháng (+0,777), Contract_Month-to-month (+0,770), TechSupport_No (+0,287), cước $95,10 -> Đưa ngay vào danh sách CSKH ưu tiên chuyển đổi gói.",
    ], font_size=13.8)

    add_dark_card(s15, 10.20, 8.02, 8.95, 1.92, title="CA TRUNG THÀNH: MÃ KH 7606-BPHHN (0,6% — THỰC TẾ: Ở LẠI)", accent=C_TEAL)
    add_bullets_dark(s15, 10.45, 8.65, 8.45, 1.15, [
        "tenure = 72 tháng (-0,967), không dùng hợp đồng tháng (-0,696), hợp đồng 2 năm (-0,303) -> Nhóm ổn định, tránh gửi ưu đãi tốn kém không cần thiết.",
    ], font_size=13.8)

    # ==================== SLIDE 16: CHUONG 4 - PHAN TICH LOI & DOI CHUNG ROI ====================
    s16 = prs.slides.add_slide(blank)
    apply_template_frame(
        s16, 16, "Chương 4", "PHÂN TÍCH LỖI CHUYÊN SÂU", "& ĐỐI CHỨNG HIỆU QUẢ KINH TẾ (ROI)",
        "Giải phẫu các ca dự đoán sai tại T* = 0,21 cho thấy 34,2% lỗi nằm ở vùng biên xác suất, và 71,9% số ca bỏ sót FN rơi vào nhóm khách dài hạn rời bỏ đột xuất do yếu tố ngoại sinh chưa có trong dữ liệu tĩnh. Bảng dưới đối chứng tổng chi phí 4 chiến lược trên tập Test, khẳng định mô hình tiết kiệm $92.700 (-49,6%)."
    )
    err_rows = [
        ["Số ca bỏ sót khách rời bỏ (FN)", "181 ca", f"{errt['fn_count']} ca (-68,5%)"],
        ["  - Khách dài hạn / Hợp đồng 1-2 năm", "92 ca (50,8%)", f"{errt['fn_loyal_silent']} ca (71,9% FN)"],
        ["  - Khách ngắn hạn hợp đồng tháng", "89 ca (49,2%)", f"{errt['fn_short_volatile']} ca (28,1% FN)"],
        ["Số ca cảnh báo dư (FP)", "91 ca", f"{errt['fp_count']} ca"],
        ["  - Thuộc hồ sơ rủi ro (Tháng / Fiber)", "91 ca (100%)", f"{errt['fp_high_risk_profile']} ca (99,1% FP)"],
        ["Lỗi sát vùng biên (|p - T| <= 0,10)", "93 ca (34,2%)", f"{errt['boundary_errors']} ca ({errt['boundary_share_pct']:.1f}%)"],
    ]
    add_dark_table(
        s16, 0.85, 2.35, 10.65, 3.85,
        ["Giải phẫu lỗi trên tập Test (N = 1.409)", "T = 0,50", f"T* = {t_star:.2f}"],
        err_rows,
        col_ratios=[4.6, 1.9, 2.3],
        font_size=13.2,
        left_align_cols=(0,),
    )

    add_dark_card(s16, 11.80, 2.35, 7.35, 3.85, title="NGUYÊN NHÂN GỐC RỄ & HÀM Ý", accent=C_RED)
    add_bullets_dark(s16, 12.05, 3.05, 6.85, 2.95, [
        ("41/57 ca FN là khách dài hạn", "Rời bỏ đột xuất do yếu tố ngoại sinh (chuyển nơi ở, khiếu nại, đối thủ) không có trong dữ liệu snapshot."),
        ("99,1% ca FP mang hồ sơ rủi ro cao", "Thuộc nhóm hợp đồng tháng/cáp quang nên chăm sóc sớm vẫn mang tính phòng ngừa hợp lý."),
        ("Vùng biên 34,2%", "Dải xác suất 0,21–0,40 phù hợp kênh chăm sóc tự động chi phí thấp (SMS/Email)."),
    ], font_size=13.5, space_after=8)

    roi_rows = [
        [r["strategy"], f"{r['targeted_customers']:,}", str(r["churn_caught"]), str(r["churn_missed"]), f"${r['total_cost_usd']:,.0f}", f"${r['savings_vs_nothing_usd']:,.0f} ({r['savings_pct']:.1f}%)"]
        for r in roi
    ]
    add_dark_table(
        s16, 0.85, 6.50, 18.30, 3.40,
        ["Chiến lược chăm sóc (N = 1.409 tập Test)", "KH can thiệp", "Giữ được (TP)", "Bỏ sót (FN)", "Tổng chi phí", "Tiết kiệm vs Không can thiệp"],
        roi_rows,
        col_ratios=[3.2, 1.3, 1.4, 1.3, 1.6, 2.4],
        font_size=13.5,
        highlight_row=3,
        left_align_cols=(0,),
    )

    # ==================== SLIDE 17: CHUONG 5 - DEMO STREAMLIT & CHIEN LUOC CSKH ====================
    s17 = prs.slides.add_slide(blank)
    apply_template_frame(
        s17, 17, "Chương 5", "ỨNG DỤNG MINH HỌA STREAMLIT", "& ĐỀ XUẤT CHIẾN LƯỢC CSKH",
        "Toàn bộ pipeline LightGBM và bộ giải thích SHAP được đóng gói vào ứng dụng Streamlit chạy 100% cục bộ gồm 3 phân hệ. Bên phải là ma trận 4 nhóm giải pháp giữ chân khách hàng gắn trực tiếp với các quy luật đã khai phá."
    )
    add_dark_card(s17, 0.85, 2.38, 7.65, 7.52, title="NGUYÊN MẪU WEB STREAMLIT (100% OFFLINE)", accent=C_BLUE_SOFT)
    add_bullets_dark(s17, 1.12, 3.15, 7.10, 6.40, [
        ("Tab 1 — Dự đoán khách hàng cá nhân", "Nhập hồ sơ thuê bao, tính tự động 6 đặc trưng miền, đối chiếu ngưỡng T* = 0,21, vẽ biểu đồ SHAP cục bộ và đề xuất kịch bản CSKH."),
        ("Tab 2 — Chấm điểm hàng loạt (CSV)", "Tải lên tệp CSV khách hàng (kèm tệp mẫu 100 KH tập Test), xếp hạng rủi ro giảm dần và xuất tệp CSV danh sách ưu tiên."),
        ("Tab 3 — Tổng quan kết quả thực nghiệm", "Hiển thị bảng đối sánh 5 mô hình, bảng ROI 4 chiến lược và các biểu đồ thực nghiệm chính."),
        ("Khởi chạy nhanh", "Nhấp đúp run_demo.bat hoặc chạy lệnh: python -m streamlit run app.py (không cần API key)."),
    ], font_size=14.6, space_after=16)

    add_dark_table(
        s17, 8.80, 2.38, 10.35, 7.52,
        ["Phân khúc rủi ro phát hiện", "Cơ sở khai phá", "Hành động chăm sóc đề xuất"],
        [
            ["Hợp đồng tháng + Fiber optic + tenure <= 13m", "Luật R1: 13,5% mẫu nhưng 71,0% Churn.", "Chủ động gọi CSKH ở tháng 2 & 6; gói chuyển đổi hợp đồng 1 năm giảm 10–15% cước 3 tháng đầu."],
            ["Cáp quang chưa dùng Security / TechSupport", "Hạng 3 & 4 trên SHAP toàn cục.", "Tặng dùng thử miễn phí 3 tháng OnlineSecurity & TechSupport để tăng trải nghiệm và giữ chân."],
            ["Thanh toán thủ công Electronic check", "Tỷ lệ Churn 45,3%; SHAP hạng 6.", "Ưu đãi trừ trực tiếp kỳ hóa đơn đầu khi liên kết thanh toán tự động qua ngân hàng / thẻ."],
            ["Vùng xác suất biên (0,21 – 0,40)", "34,2% vùng lỗi biên quanh ngưỡng T*.", "Gửi khảo sát hài lòng & ưu đãi băng thông qua kênh chi phí thấp (SMS/Email tự động)."],
        ],
        col_ratios=[2.7, 2.3, 4.2],
        font_size=13.2,
        left_align_cols=(0, 1, 2),
    )

    # ==================== SLIDE 18: KET LUAN & PHU LUC SIEU THAM SO ====================
    s18 = prs.slides.add_slide(blank)
    apply_template_frame(
        s18, 18, "Kết luận", "TỔNG KẾT, HƯỚNG PHÁT TRIỂN", "& PHỤ LỤC SIÊU THAM SỐ TỐI ƯU",
        "Đề tài đã hoàn thiện trọn vẹn quy trình Khai phá dữ liệu phân loại cho bài toán Customer Churn với kết quả thực nghiệm nhất quán, minh bạch tri thức và hiệu quả chi phí rõ rệt. Bảng phía dưới là bộ siêu tham số tối ưu của 5 mô hình sau Stratified 5-Fold GridSearchCV."
    )
    add_dark_card(s18, 0.85, 2.35, 8.95, 3.70, title="KẾT QUẢ CHÍNH ĐẠT ĐƯỢC", accent=C_TEAL)
    add_bullets_dark(s18, 1.12, 3.05, 8.40, 2.80, [
        ("Quy trình KDD chống rò rỉ", "Kiểm định thống kê đầy đủ, 6 đặc trưng miền và tách tập Test độc lập ngay từ đầu."),
        ("Hiệu năng & Hiệu quả kinh tế", "LightGBM đạt CV ROC-AUC 0,8494, PR-AUC 0,6657; ngưỡng T* = 0,21 nâng Recall lên 84,8% và giảm 49,6% chi phí."),
        ("Tri thức giải thích được", "8 luật CART tường minh + SHAP toàn cục/cục bộ + giải phẫu lỗi FN/FP + Web Demo cục bộ."),
    ], font_size=14.2, space_after=10)

    add_dark_card(s18, 10.20, 2.35, 8.95, 3.70, title="HẠN CHẾ & HƯỚNG PHÁT TRIỂN", accent=C_YELLOW)
    add_bullets_dark(s18, 10.47, 3.05, 8.40, 2.80, [
        ("Hạn chế dữ liệu tĩnh (Snapshot)", "Chưa có log khiếu nại tổng đài hay xu hướng lưu lượng theo tháng -> gây ra 41 ca FN dài hạn."),
        ("Phân tích sống còn (Survival Analysis)", "Mở rộng dự báo đồng thời xác suất và thời điểm dự kiến rời bỏ để chọn thời điểm can thiệp."),
        ("Kết hợp phân cụm (Clustering)", "Phân khúc khách hàng theo RFM trước khi phân loại nhằm cá nhân hóa ngân sách giữ chân theo CLV."),
    ], font_size=14.2, space_after=10)

    hp_rows = [
        [nm, ", ".join(f"{k} = {v}" for k, v in res["best_params"][nm].items()), f"{res['cv_roc_auc'][nm]:.4f}", f"{res['train_times_sec'][nm]:.2f}s"]
        for nm in m_order
    ]
    add_dark_table(
        s18, 0.85, 6.32, 18.30, 3.55,
        ["[Phụ lục] Mô hình", "Bộ siêu tham số tối ưu sau Stratified 5-Fold GridSearchCV", "CV ROC-AUC", "Thời gian"],
        hp_rows,
        col_ratios=[2.3, 6.2, 1.3, 1.1],
        font_size=13.5,
        highlight_row=4,
        left_align_cols=(0, 1),
    )

    # ==================== SLIDE 19: TRANG CAM ON (THEO SLIDE 15 THIET KE MAU + SLIDE 32 HE TRI THUC) ====================
    s19 = prs.slides.add_slide(blank)
    set_slide_bg(s19, C_BG_DARK)
    if os.path.exists(IMG_COVER_BG):
        s19.shapes.add_picture(IMG_COVER_BG, Inches(0), Inches(0), SLIDE_W, SLIDE_H)
    if os.path.exists(IMG_CIRCUIT):
        s19.shapes.add_picture(IMG_CIRCUIT, Inches(-1.05), Inches(7.70), Inches(5.60), Inches(3.28))

    if os.path.exists(LOGO_HUBT):
        s19.shapes.add_picture(LOGO_HUBT, Inches(9.20), Inches(1.15), Inches(1.60), Inches(1.06))

    tb_ty = s19.shapes.add_textbox(Inches(2.0), Inches(2.65), Inches(16.0), Inches(2.20))
    tf_ty = tb_ty.text_frame
    tf_ty.word_wrap = True
    pty1 = tf_ty.paragraphs[0]
    pty1.text = "TRÂN TRỌNG CẢM ƠN!"
    pty1.font.name = FONT_DISPLAY
    pty1.font.size = Pt(76)
    pty1.font.bold = True
    pty1.font.color.rgb = C_YELLOW
    pty1.alignment = PP_ALIGN.CENTER

    pty2 = tf_ty.add_paragraph()
    pty2.text = "Kính mời TS. Bùi Ngọc Thăng và các anh/chị học viên đặt câu hỏi thảo luận & xem Demo trực tiếp"
    pty2.font.name = FONT_BODY
    pty2.font.size = Pt(21)
    pty2.font.color.rgb = C_WHITE
    pty2.alignment = PP_ALIGN.CENTER
    pty2.space_before = Pt(18)

    line_end = s19.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(6.20), Inches(5.75), Inches(7.60), Inches(0.05))
    line_end.fill.solid()
    line_end.fill.fore_color.rgb = C_YELLOW
    line_end.line.fill.background()

    tb_info = s19.shapes.add_textbox(Inches(2.5), Inches(6.15), Inches(15.0), Inches(2.50))
    tf_in = tb_info.text_frame
    tf_in.word_wrap = True
    pi1 = tf_in.paragraphs[0]
    pi1.text = "NHÓM 02 — LỚP CAO HỌC CNTT 20.2 — TRƯỜNG ĐẠI HỌC KINH DOANH VÀ CÔNG NGHỆ HÀ NỘI"
    pi1.font.name = FONT_DISPLAY
    pi1.font.size = Pt(17)
    pi1.font.bold = True
    pi1.font.color.rgb = C_WHITE
    pi1.alignment = PP_ALIGN.CENTER

    pi2 = tf_in.add_paragraph()
    pi2.text = "Đỗ Ngọc Điền  •  Dương Tùng Lâm  •  Trần Thị Thu Hương  •  Phạm Văn Chiến  •  Vàng A Hứ"
    pi2.font.name = FONT_BODY
    pi2.font.size = Pt(15.5)
    pi2.font.color.rgb = C_MUTED
    pi2.alignment = PP_ALIGN.CENTER
    pi2.space_before = Pt(10)

    pi3 = tf_in.add_paragraph()
    pi3.text = "Mã nguồn & Báo cáo: https://github.com/Lamdt198/KhaiPhaDuLieu-CustomerChurn"
    pi3.font.name = FONT_CODE
    pi3.font.size = Pt(15)
    pi3.font.color.rgb = C_YELLOW
    pi3.alignment = PP_ALIGN.CENTER
    pi3.space_before = Pt(18)

    add_notes(s19, "Xin trân trọng cảm ơn TS. Bùi Ngọc Thăng và các anh chị đã theo dõi phần trình bày của Nhóm 02. Nhóm rất mong nhận được các ý kiến đóng góp và sẵn sàng trình diễn trực tiếp ứng dụng minh họa trên Streamlit.")

    os.makedirs(REPORT_DIR, exist_ok=True)
    prs.save(OUTPUT_PPTX)
    print(f"Presentation generated successfully: {OUTPUT_PPTX} ({len(prs.slides)} slides, 20.0x11.25 in)")


if __name__ == "__main__":
    build_deck()
