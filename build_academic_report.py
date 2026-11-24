# -*- coding: utf-8 -*-
"""
build_academic_report.py
Sinh Bao cao Bai tap lon hoc phan Khai pha du lieu (Data Mining) - Nhom 02 - Lop CNTT 20.2.
Giang vien huong dan: TS. Bui Ngoc Thang.
Ke thua phoi bia chinh thuc Vien Sau Dai hoc HUBT (BaoCao_CHPT-N02_Original_Cover.docx),
tich hop Word Equation chuan OMML, so do khoi KDD, anh chup giao dien Streamlit thuc te
va dinh dang so chuan tieng Viet dong nhat 100%.
"""

import json
import os
import re
import sys
from PIL import Image
from docx import Document
from docx.shared import Inches, Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT, WD_ALIGN_VERTICAL
from docx.oxml import parse_xml
from docx.oxml.ns import nsdecls, qn

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

COVER_PATH = os.path.join(ASSET_DIR, "BaoCao_CHPT-N02_Original_Cover.docx")
OUTPUT_DOCX = os.path.join(REPORT_DIR, "Bao_Cao_BTL_Khai_Pha_Du_Lieu_Nhom02.docx")

NAVY_RGB = RGBColor(30, 58, 138)
DARK_RGB = RGBColor(15, 23, 42)
BODY_RGB = RGBColor(26, 26, 26)
MUTED_RGB = RGBColor(71, 85, 105)


# ==================== DINH DANG SO CHUAN TIENG VIET ====================
def fmt_int(val) -> str:
    """Dinh dang so nguyen phan tach hang nghin bang dau cham (VD: 7043 -> 7.043)."""
    return f"{int(round(val)):,}".replace(",", ".")


def fmt_dec(val: float, digits: int = 4, signed: bool = False) -> str:
    """Dinh dang so thuc chuan tieng Viet: dau cham hang nghin, dau phay thap phan."""
    s = f"{val:+,.{digits}f}" if signed else f"{val:,.{digits}f}"
    return s.replace(",", "X").replace(".", ",").replace("X", ".")


def fmt_pct(val: float, digits: int = 2) -> str:
    """Dinh dang phan tram tu gia tri da nhan 100 hoac ty le."""
    return f"{fmt_dec(val, digits)}%"


def fmt_sci(val: float) -> str:
    """Dinh dang ky hieu khoa hoc voi dau phay thap phan (VD: 1,23e-150)."""
    s = f"{val:.2e}"
    mant, exp = s.split("e")
    return f"{mant.replace('.', ',')}e{exp}"


def fmt_usd(val: float, symbol_prefix: bool = True) -> str:
    """Dinh dang tien te USD voi dau cham phan cach hang nghin."""
    num_s = fmt_int(val)
    return f"${num_s}" if symbol_prefix else f"{num_s} USD"


# ==================== TIEN ICH XML & WORD EQUATION (OMML) ====================
def set_cell_background(cell, fill_hex):
    tcPr = cell._element.get_or_add_tcPr()
    tcPr.append(parse_xml(f'<w:shd {nsdecls("w")} w:fill="{fill_hex}"/>'))


def set_cell_margins(cell, top=80, bottom=80, left=100, right=100):
    tcPr = cell._element.get_or_add_tcPr()
    tcPr.append(
        parse_xml(
            f'<w:tcMar {nsdecls("w")}>'
            f'<w:top w:w="{top}" w:type="dxa"/>'
            f'<w:bottom w:w="{bottom}" w:type="dxa"/>'
            f'<w:left w:w="{left}" w:type="dxa"/>'
            f'<w:right w:w="{right}" w:type="dxa"/>'
            f"</w:tcMar>"
        )
    )


def reset_indents(p):
    p.paragraph_format.first_line_indent = Pt(0)
    p.paragraph_format.left_indent = Pt(0)
    p.paragraph_format.right_indent = Pt(0)


def mr(txt: str, nor: bool = False, sty: str = None) -> str:
    """Tao phan tu <m:r> trong Word OMML (font Cambria Math 12pt)."""
    esc = txt.replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;")
    m_rpr = ""
    if nor:
        m_rpr = "<m:rPr><m:nor/></m:rPr>"
    elif sty:
        m_rpr = f'<m:rPr><m:sty m:val="{sty}"/></m:rPr>'
    w_rpr = (
        '<w:rPr>'
        '<w:rFonts w:ascii="Cambria Math" w:hAnsi="Cambria Math" w:cs="Cambria Math"/>'
        '<w:sz w:val="24"/><w:szCs w:val="24"/>'
        '</w:rPr>'
    )
    return f'<m:r>{m_rpr}{w_rpr}<m:t xml:space="preserve">{esc}</m:t></m:r>'


def mfrac(num_xml: str, den_xml: str) -> str:
    return (
        f"<m:f>"
        f'<m:fPr><m:type m:val="bar"/></m:fPr>'
        f"<m:num>{num_xml}</m:num>"
        f"<m:den>{den_xml}</m:den>"
        f"</m:f>"
    )


def msub(base_xml: str, sub_xml: str) -> str:
    return f"<m:sSub><m:e>{base_xml}</m:e><m:sub>{sub_xml}</m:sub></m:sSub>"


def msup(base_xml: str, sup_xml: str) -> str:
    return f"<m:sSup><m:e>{base_xml}</m:e><m:sup>{sup_xml}</m:sup></m:sSup>"


def msubsup(base_xml: str, sub_xml: str, sup_xml: str) -> str:
    return (
        f"<m:sSubSup>"
        f"<m:e>{base_xml}</m:e>"
        f"<m:sub>{sub_xml}</m:sub>"
        f"<m:sup>{sup_xml}</m:sup>"
        f"</m:sSubSup>"
    )


def msum(sub_xml: str, sup_xml: str, body_xml: str) -> str:
    return (
        f"<m:nary>"
        f'<m:naryPr><m:chr m:val="∑"/><m:limLoc m:val="undOvr"/></m:naryPr>'
        f"<m:sub>{sub_xml}</m:sub>"
        f"<m:sup>{sup_xml}</m:sup>"
        f"<m:e>{body_xml}</m:e>"
        f"</m:nary>"
    )


def mdelim(inner_xml: str, beg: str = "(", end: str = ")") -> str:
    esc_b = beg.replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;")
    esc_e = end.replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;")
    return (
        f"<m:d>"
        f'<m:dPr><m:begChr m:val="{esc_b}"/><m:endChr m:val="{esc_e}"/></m:dPr>'
        f"<m:e>{inner_xml}</m:e>"
        f"</m:d>"
    )


def enable_update_fields_on_open(doc):
    """Bat co updateFields de Microsoft Word tu dong cap nhat so trang Muc luc khi mo tep."""
    settings = doc.settings.element
    settings.append(parse_xml(f'<w:updateFields {nsdecls("w")} w:val="true"/>'))


def add_toc_field(doc):
    p = doc.add_paragraph()
    reset_indents(p)
    p.paragraph_format.space_before = Pt(4)
    p.paragraph_format.space_after = Pt(6)
    r = p.add_run()
    r.font.name = "Times New Roman"
    r.font.size = Pt(12)
    r._r.append(parse_xml(f'<w:fldChar {nsdecls("w")} w:fldCharType="begin"/>'))
    r._r.append(parse_xml(f'<w:instrText {nsdecls("w")} xml:space="preserve"> TOC \\o "1-3" \\h \\z \\u </w:instrText>'))
    r._r.append(parse_xml(f'<w:fldChar {nsdecls("w")} w:fldCharType="separate"/>'))
    t = parse_xml(f'<w:t {nsdecls("w")}>Nhấn chuột phải chọn Update Field (hoặc Ctrl+A rồi F9) trong Microsoft Word để cập nhật số trang mục lục.</w:t>')
    r._r.append(t)
    r._r.append(parse_xml(f'<w:fldChar {nsdecls("w")} w:fldCharType="end"/>'))


def format_cover_paragraph(p, text, size_pt=14, bold=False, italic=False, color=DARK_RGB, align=None, clear_indent=False):
    p.text = text
    if align is not None:
        p.alignment = align
    if clear_indent:
        reset_indents(p)
    for r in p.runs:
        r.font.name = "Times New Roman"
        r.font.size = Pt(size_pt)
        r.font.bold = bold
        r.font.italic = italic
        r.font.color.rgb = color


def build_report():
    with open(os.path.join(MODEL_DIR, "eda_stats.json"), "r", encoding="utf-8") as f:
        eda = json.load(f)
    with open(os.path.join(MODEL_DIR, "experiment_results.json"), "r", encoding="utf-8") as f:
        res = json.load(f)

    if os.path.exists(COVER_PATH):
        doc = Document(COVER_PATH)
        # Luu y: them dau cach truoc \n de khi xuat PDF hoac trich xuat van ban khong bi dinh chu
        format_cover_paragraph(
            doc.paragraphs[10],
            "HỌC PHẦN: KHAI PHÁ DỮ LIỆU \n(DATA MINING)",
            size_pt=16, bold=True, color=RGBColor(30, 41, 59),
            align=WD_ALIGN_PARAGRAPH.CENTER, clear_indent=True,
        )
        format_cover_paragraph(
            doc.paragraphs[12],
            "ĐỀ TÀI: “ỨNG DỤNG CÁC MÔ HÌNH PHÂN LOẠI DỮ LIỆU \nTRONG DỰ ĐOÁN NGUY CƠ RỜI BỎ DỊCH VỤ CỦA KHÁCH HÀNG”",
            size_pt=14, bold=True, color=NAVY_RGB,
            align=WD_ALIGN_PARAGRAPH.CENTER, clear_indent=True,
        )
        format_cover_paragraph(doc.paragraphs[14], "Thành viên nhóm:\tĐỗ Ngọc Điền", size_pt=14)
        for idx, member in enumerate([
            "\tDương Tùng Lâm",
            "\tTrần Thị Thu Hương",
            "\tPhạm Văn Chiến",
            "\tVàng A Hứ",
        ], start=15):
            format_cover_paragraph(doc.paragraphs[idx], member, size_pt=14)
        format_cover_paragraph(doc.paragraphs[19], "Lớp: Công nghệ thông tin 20.2", size_pt=14)
        format_cover_paragraph(doc.paragraphs[20], "Giảng viên hướng dẫn: TS. Bùi Ngọc Thăng", size_pt=14)

        # Xoa doan trong thua sau sectPr cua phoi bia (doan 21)
        if len(doc.paragraphs) > 21:
            p21 = doc.paragraphs[21]
            p21._element.getparent().remove(p21._element)

        # Xoa footer cu con sot ("Ha Noi, Thang 09/2026") trong Section 0 de tranh trung lap
        s0 = doc.sections[0]
        for ftr in (s0.first_page_footer, s0.footer):
            for p_f in ftr.paragraphs:
                p_f.text = ""

        # Chuyen sectPr tu doan 20 xuong doan moi "Ha Noi, nam 2026" o cuoi khung trang bia
        p20 = doc.paragraphs[20]
        p20_pPr = p20._element.get_or_add_pPr()
        sectPr_el = p20_pPr.find(qn("w:sectPr"))
        if sectPr_el is not None:
            p20_pPr.remove(sectPr_el)

        p_date = doc.add_paragraph()
        format_cover_paragraph(
            p_date,
            "Hà Nội, năm 2026",
            size_pt=14, bold=True, italic=False, color=DARK_RGB,
            align=WD_ALIGN_PARAGRAPH.CENTER, clear_indent=True,
        )
        p_date.paragraph_format.space_before = Pt(22)
        p_date.paragraph_format.space_after = Pt(0)
        if sectPr_el is not None:
            p_date._element.get_or_add_pPr().append(sectPr_el)

        # Rut bot 2 dong trong duoi logo (doan 8 va 9, giu nguyen doan 5 chua logo) de trang bia vua khit 1 trang
        for empty_idx in (9, 8):
            p_emp = doc.paragraphs[empty_idx]
            if p_emp.text == "" and "<w:drawing" not in p_emp._element.xml:
                p_emp._element.getparent().remove(p_emp._element)

        s1 = doc.sections[1]
        s1.different_first_page_header_footer = False
        s1.page_width = Inches(8.27)
        s1.page_height = Inches(11.69)
        s1.top_margin = Inches(0.79)
        s1.bottom_margin = Inches(0.79)
        s1.left_margin = Inches(0.98)
        s1.right_margin = Inches(0.79)

        s1.footer.is_linked_to_previous = False
        for child in list(s1.footer._element):
            s1.footer._element.remove(child)
        fp = s1.footer.add_paragraph()
        reset_indents(fp)
        fp.alignment = WD_ALIGN_PARAGRAPH.CENTER
        fp.paragraph_format.space_before = Pt(4)
        fp.paragraph_format.space_after = Pt(0)
        r_fp = fp.add_run()
        r_fp.font.name = "Times New Roman"
        r_fp.font.size = Pt(10.5)
        r_fp.font.color.rgb = MUTED_RGB
        r_fp._r.append(parse_xml(f'<w:fldChar {nsdecls("w")} w:fldCharType="begin"/>'))
        r_fp._r.append(parse_xml(f'<w:instrText {nsdecls("w")} xml:space="preserve"> PAGE </w:instrText>'))
        r_fp._r.append(parse_xml(f'<w:fldChar {nsdecls("w")} w:fldCharType="separate"/>'))
        r_fp._r.append(parse_xml(f'<w:fldChar {nsdecls("w")} w:fldCharType="end"/>'))
    else:
        doc = Document()

    enable_update_fields_on_open(doc)

    # Cau hinh Headings chuan
    for s_name, f_size, f_color, sp_bef, sp_aft in [
        ("Heading 1", Pt(15), NAVY_RGB, Pt(16), Pt(8)),
        ("Heading 2", Pt(13.5), DARK_RGB, Pt(12), Pt(4)),
        ("Heading 3", Pt(13), RGBColor(30, 41, 59), Pt(8), Pt(3)),
    ]:
        if s_name in doc.styles:
            st = doc.styles[s_name]
            st.font.name = "Times New Roman"
            st.font.size = f_size
            st.font.bold = True
            st.font.color.rgb = f_color
            st.paragraph_format.first_line_indent = Pt(0)
            st.paragraph_format.left_indent = Pt(0)
            st.paragraph_format.space_before = sp_bef
            st.paragraph_format.space_after = sp_aft
            st.paragraph_format.keep_with_next = True

    def add_h1(text, page_break=False):
        if page_break:
            doc.add_page_break()
        p = doc.add_paragraph(style="Heading 1")
        p._element.get_or_add_pPr().append(parse_xml(f'<w:outlineLvl {nsdecls("w")} w:val="0"/>'))
        reset_indents(p)
        p.alignment = WD_ALIGN_PARAGRAPH.LEFT
        p.paragraph_format.space_before = Pt(14 if not page_break else 4)
        p.paragraph_format.space_after = Pt(8)
        p.paragraph_format.keep_with_next = True
        r = p.add_run(text)
        r.font.name = "Times New Roman"
        r.font.size = Pt(15)
        r.font.bold = True
        r.font.color.rgb = NAVY_RGB
        return p

    def add_h2(text):
        p = doc.add_paragraph(style="Heading 2")
        p._element.get_or_add_pPr().append(parse_xml(f'<w:outlineLvl {nsdecls("w")} w:val="1"/>'))
        reset_indents(p)
        p.alignment = WD_ALIGN_PARAGRAPH.LEFT
        p.paragraph_format.space_before = Pt(11)
        p.paragraph_format.space_after = Pt(4)
        p.paragraph_format.keep_with_next = True
        r = p.add_run(text)
        r.font.name = "Times New Roman"
        r.font.size = Pt(13.5)
        r.font.bold = True
        r.font.color.rgb = DARK_RGB
        return p

    def add_h3(text):
        p = doc.add_paragraph(style="Heading 3")
        p._element.get_or_add_pPr().append(parse_xml(f'<w:outlineLvl {nsdecls("w")} w:val="2"/>'))
        reset_indents(p)
        p.alignment = WD_ALIGN_PARAGRAPH.LEFT
        p.paragraph_format.space_before = Pt(8)
        p.paragraph_format.space_after = Pt(3)
        p.paragraph_format.keep_with_next = True
        r = p.add_run(text)
        r.font.name = "Times New Roman"
        r.font.size = Pt(13)
        r.font.bold = True
        r.font.italic = True
        r.font.color.rgb = RGBColor(30, 41, 59)
        return p

    def add_para(text, bold_prefix=None, italic=False, space_after=6):
        p = doc.add_paragraph()
        p.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
        p.paragraph_format.space_before = Pt(0)
        p.paragraph_format.space_after = Pt(space_after)
        p.paragraph_format.line_spacing = 1.25
        if bold_prefix:
            rp = p.add_run(bold_prefix)
            rp.font.name = "Times New Roman"
            rp.font.size = Pt(13)
            rp.font.bold = True
            rp.font.color.rgb = DARK_RGB
        r = p.add_run(text)
        r.font.name = "Times New Roman"
        r.font.size = Pt(13)
        r.font.italic = italic
        r.font.color.rgb = BODY_RGB
        return p

    def add_bullet(text, bold_prefix=None):
        p = doc.add_paragraph()
        p.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
        p.paragraph_format.first_line_indent = Pt(0)
        p.paragraph_format.left_indent = Inches(0.28)
        p.paragraph_format.right_indent = Pt(0)
        p.paragraph_format.space_before = Pt(0)
        p.paragraph_format.space_after = Pt(4)
        p.paragraph_format.line_spacing = 1.25
        rb = p.add_run("-  ")
        rb.font.name = "Times New Roman"
        rb.font.size = Pt(13)
        rb.font.bold = True
        rb.font.color.rgb = NAVY_RGB
        if bold_prefix:
            rp = p.add_run(bold_prefix)
            rp.font.name = "Times New Roman"
            rp.font.size = Pt(13)
            rp.font.bold = True
            rp.font.color.rgb = DARK_RGB
        r = p.add_run(text)
        r.font.name = "Times New Roman"
        r.font.size = Pt(13)
        r.font.color.rgb = BODY_RGB
        return p

    def add_omml_equation(omml_inner: str, eq_num: str = None):
        """Chen cong thuc toan hoc Word Equation chuan OMML can giua kem so hieu can phai."""
        w_eq_in = 5.78
        w_num_in = 0.72
        widths_dxa = [int(round(w_eq_in * 1440)), int(round(w_num_in * 1440))]

        tbl = doc.add_table(rows=1, cols=2)
        tbl.alignment = WD_TABLE_ALIGNMENT.CENTER
        tbl.autofit = False
        tblPr = tbl._element.xpath("w:tblPr")[0]
        tblPr.append(parse_xml(f'<w:tblLayout {nsdecls("w")} w:type="fixed"/>'))
        tblPr.append(parse_xml(f'<w:tblW {nsdecls("w")} w:w="{sum(widths_dxa)}" w:type="dxa"/>'))
        tblPr.append(parse_xml(f'<w:tblInd {nsdecls("w")} w:w="0" w:type="dxa"/>'))
        tblPr.append(
            parse_xml(
                f'<w:tblBorders {nsdecls("w")}>'
                f'<w:top w:val="none"/><w:bottom w:val="none"/>'
                f'<w:left w:val="none"/><w:right w:val="none"/>'
                f'<w:insideH w:val="none"/><w:insideV w:val="none"/>'
                f"</w:tblBorders>"
            )
        )
        tblGrid = tbl._element.xpath("w:tblGrid")[0]
        tblGrid.clear()
        for wd in widths_dxa:
            tblGrid.append(parse_xml(f'<w:gridCol {nsdecls("w")} w:w="{wd}"/>'))

        row = tbl.rows[0]
        row._element.get_or_add_trPr().append(parse_xml(f'<w:cantSplit {nsdecls("w")}/>'))

        # Cot 0: Cong thuc OMML can giua
        c0 = row.cells[0]
        c0.width = Inches(w_eq_in)
        c0._element.get_or_add_tcPr().append(parse_xml(f'<w:tcW {nsdecls("w")} w:w="{widths_dxa[0]}" w:type="dxa"/>'))
        set_cell_margins(c0, top=45, bottom=45, left=20, right=20)
        c0.vertical_alignment = WD_ALIGN_VERTICAL.CENTER
        p0 = c0.paragraphs[0]
        reset_indents(p0)
        p0.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p0.paragraph_format.space_before = Pt(2)
        p0.paragraph_format.space_after = Pt(2)
        omml_el = parse_xml(
            f'<m:oMathPara {nsdecls("w", "m")}>'
            f'<m:oMathParaPr><m:jc m:val="center"/></m:oMathParaPr>'
            f"<m:oMath>{omml_inner}</m:oMath>"
            f"</m:oMathPara>"
        )
        p0._element.append(omml_el)

        # Cot 1: So hieu cong thuc (1.1)-(1.6) can phai
        c1 = row.cells[1]
        c1.width = Inches(w_num_in)
        c1._element.get_or_add_tcPr().append(parse_xml(f'<w:tcW {nsdecls("w")} w:w="{widths_dxa[1]}" w:type="dxa"/>'))
        set_cell_margins(c1, top=45, bottom=45, left=10, right=0)
        c1.vertical_alignment = WD_ALIGN_VERTICAL.CENTER
        p1 = c1.paragraphs[0]
        reset_indents(p1)
        p1.alignment = WD_ALIGN_PARAGRAPH.RIGHT
        p1.paragraph_format.space_before = Pt(2)
        p1.paragraph_format.space_after = Pt(2)
        if eq_num:
            rn = p1.add_run(f"({eq_num})")
            rn.font.name = "Times New Roman"
            rn.font.size = Pt(12.5)
            rn.font.color.rgb = DARK_RGB

    def add_note_box(text, label="Nhận xét"):
        tbl = doc.add_table(rows=1, cols=1)
        tbl.alignment = WD_TABLE_ALIGNMENT.CENTER
        tbl.autofit = False
        tblPr = tbl._element.xpath("w:tblPr")[0]
        tblPr.append(parse_xml(f'<w:tblInd {nsdecls("w")} w:w="0" w:type="dxa"/>'))
        tbl.columns[0].width = Inches(6.5)
        c = tbl.cell(0, 0)
        set_cell_background(c, "F1F5F9")
        set_cell_margins(c, top=110, bottom=110, left=160, right=150)
        tcPr = c._element.get_or_add_tcPr()
        tcPr.append(
            parse_xml(
                f'<w:tcBorders {nsdecls("w")}>'
                f'<w:top w:val="single" w:sz="4" w:space="0" w:color="CBD5E1"/>'
                f'<w:bottom w:val="single" w:sz="4" w:space="0" w:color="CBD5E1"/>'
                f'<w:right w:val="single" w:sz="4" w:space="0" w:color="CBD5E1"/>'
                f'<w:left w:val="single" w:sz="20" w:space="0" w:color="1E3A8A"/>'
                f"</w:tcBorders>"
            )
        )
        p = c.paragraphs[0]
        p.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
        reset_indents(p)
        p.paragraph_format.space_before = Pt(2)
        p.paragraph_format.space_after = Pt(2)
        p.paragraph_format.line_spacing = 1.2
        r1 = p.add_run(f"{label}: ")
        r1.font.name = "Times New Roman"
        r1.font.size = Pt(11.5)
        r1.font.bold = True
        r1.font.color.rgb = NAVY_RGB
        r2 = p.add_run(text)
        r2.font.name = "Times New Roman"
        r2.font.size = Pt(11.5)
        r2.font.color.rgb = DARK_RGB
        sp = doc.add_paragraph()
        reset_indents(sp)
        sp.paragraph_format.space_after = Pt(4)

    def add_table(headers, rows, col_widths=None, caption=None):
        if caption:
            pc = doc.add_paragraph()
            pc.alignment = WD_ALIGN_PARAGRAPH.CENTER
            reset_indents(pc)
            pc.paragraph_format.space_before = Pt(6)
            pc.paragraph_format.space_after = Pt(4)
            pc.paragraph_format.keep_with_next = True
            rc = pc.add_run(caption)
            rc.font.name = "Times New Roman"
            rc.font.size = Pt(11.5)
            rc.font.bold = True
            rc.font.italic = True
            rc.font.color.rgb = NAVY_RGB

        total_w_in = 6.50
        num_cols = len(headers)
        if col_widths:
            scale = total_w_in / sum(col_widths)
            widths_in = [round(w * scale, 3) for w in col_widths]
        else:
            widths_in = [round(total_w_in / num_cols, 3)] * num_cols
        widths_in[-1] = round(widths_in[-1] + (total_w_in - sum(widths_in)), 3)
        widths_dxa = [int(round(w * 1440)) for w in widths_in]

        table = doc.add_table(rows=len(rows) + 1, cols=num_cols)
        table.alignment = WD_TABLE_ALIGNMENT.CENTER
        table.autofit = False
        tblPr = table._element.xpath("w:tblPr")[0]
        tblPr.append(parse_xml(f'<w:tblLayout {nsdecls("w")} w:type="fixed"/>'))
        tblPr.append(parse_xml(f'<w:tblW {nsdecls("w")} w:w="{sum(widths_dxa)}" w:type="dxa"/>'))
        tblPr.append(parse_xml(f'<w:tblInd {nsdecls("w")} w:w="0" w:type="dxa"/>'))
        tblPr.append(
            parse_xml(
                f'<w:tblBorders {nsdecls("w")}>'
                f'<w:top w:val="single" w:sz="8" w:space="0" w:color="1E3A8A"/>'
                f'<w:bottom w:val="single" w:sz="8" w:space="0" w:color="1E3A8A"/>'
                f'<w:left w:val="single" w:sz="8" w:space="0" w:color="1E3A8A"/>'
                f'<w:right w:val="single" w:sz="8" w:space="0" w:color="1E3A8A"/>'
                f'<w:insideH w:val="single" w:sz="4" w:space="0" w:color="94A3B8"/>'
                f'<w:insideV w:val="single" w:sz="4" w:space="0" w:color="94A3B8"/>'
                f"</w:tblBorders>"
            )
        )
        tblGrid = table._element.xpath("w:tblGrid")[0]
        tblGrid.clear()
        for w_dxa in widths_dxa:
            tblGrid.append(parse_xml(f'<w:gridCol {nsdecls("w")} w:w="{w_dxa}"/>'))

        center_kw = [
            "stt", "hạng", "mã", "acc", "rec", "accuracy", "precision", "recall", "f1", "auc", "brier",
            "chi phí", "tỷ lệ", "số", "p-value", "cramér", "chi-square", "odds",
            "hệ số", "support", "xác suất", "ngưỡng", "thời gian", "tiết kiệm",
            "trung vị", "trung bình", "đóng góp", "tp", "fp", "fn", "tn", "shap",
        ]
        col_center = [any(k in h.lower() for k in center_kw) for h in headers]
        hdr_pt = Pt(9.2 if num_cols >= 6 else 10.0)
        body_pt = Pt(8.8 if num_cols >= 6 else 9.8)

        hdr_row = table.rows[0]
        trPr = hdr_row._element.get_or_add_trPr()
        trPr.append(parse_xml(f'<w:tblHeader {nsdecls("w")}/>'))
        trPr.append(parse_xml(f'<w:cantSplit {nsdecls("w")}/>'))
        for idx, htxt in enumerate(headers):
            cell = hdr_row.cells[idx]
            cell.width = Inches(widths_in[idx])
            cell._element.get_or_add_tcPr().append(
                parse_xml(f'<w:tcW {nsdecls("w")} w:w="{widths_dxa[idx]}" w:type="dxa"/>')
            )
            set_cell_background(cell, "1E3A8A")
            set_cell_margins(cell, top=80, bottom=80, left=60, right=60)
            cell.vertical_alignment = WD_ALIGN_VERTICAL.CENTER
            p = cell.paragraphs[0]
            p.alignment = WD_ALIGN_PARAGRAPH.CENTER
            reset_indents(p)
            p.paragraph_format.space_before = Pt(0)
            p.paragraph_format.space_after = Pt(0)
            r = p.add_run(htxt)
            r.font.name = "Times New Roman"
            r.font.size = hdr_pt
            r.font.bold = True
            r.font.color.rgb = RGBColor(255, 255, 255)

        for r_idx, row_vals in enumerate(rows):
            trow = table.rows[r_idx + 1]
            trow._element.get_or_add_trPr().append(parse_xml(f'<w:cantSplit {nsdecls("w")}/>'))
            bg = "F8FAFC" if r_idx % 2 == 1 else "FFFFFF"
            for c_idx, val in enumerate(row_vals):
                cell = trow.cells[c_idx]
                cell.width = Inches(widths_in[c_idx])
                cell._element.get_or_add_tcPr().append(
                    parse_xml(f'<w:tcW {nsdecls("w")} w:w="{widths_dxa[c_idx]}" w:type="dxa"/>')
                )
                set_cell_background(cell, bg)
                set_cell_margins(cell, top=55, bottom=55, left=60, right=60)
                cell.vertical_alignment = WD_ALIGN_VERTICAL.CENTER
                p = cell.paragraphs[0]
                p.alignment = WD_ALIGN_PARAGRAPH.CENTER if col_center[c_idx] else WD_ALIGN_PARAGRAPH.LEFT
                reset_indents(p)
                p.paragraph_format.space_before = Pt(0)
                p.paragraph_format.space_after = Pt(0)
                p.paragraph_format.line_spacing = 1.08
                r = p.add_run(str(val))
                r.font.name = "Times New Roman"
                r.font.size = body_pt
                if c_idx == 0 and not col_center[0]:
                    r.font.bold = True
                    r.font.color.rgb = DARK_RGB
                else:
                    r.font.color.rgb = BODY_RGB

        sp = doc.add_paragraph()
        reset_indents(sp)
        sp.paragraph_format.space_after = Pt(6)

    def add_figure(filename, caption, width_in=6.25):
        fpath = os.path.join(FIG_DIR, filename)
        if not os.path.exists(fpath):
            raise FileNotFoundError(f"Missing figure file: {fpath}")
        with Image.open(fpath) as im:
            w_px, h_px = im.size
        aspect = h_px / w_px if w_px > 0 else 0.6
        chosen_w = min(width_in, 5.5 / aspect) if aspect > 0.85 else width_in

        p_img = doc.add_paragraph()
        reset_indents(p_img)
        p_img.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p_img.paragraph_format.space_before = Pt(8)
        p_img.paragraph_format.space_after = Pt(3)
        p_img.paragraph_format.keep_with_next = True
        p_img.add_run().add_picture(fpath, width=Inches(chosen_w))

        p_cap = doc.add_paragraph()
        reset_indents(p_cap)
        p_cap.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p_cap.paragraph_format.space_before = Pt(2)
        p_cap.paragraph_format.space_after = Pt(8)
        r = p_cap.add_run(caption)
        r.font.name = "Times New Roman"
        r.font.size = Pt(11)
        r.font.bold = True
        r.font.italic = True
        r.font.color.rgb = NAVY_RGB

    # Lay cac so lieu thuc nghiem vao bien de viet bao cao chinh xac 100%
    audit = res["dataset_audit"]
    best_name = res["best_model_name"]
    t_star = res["optimal_threshold"]
    t_star_vn = fmt_dec(t_star, 2)
    m05_best = res["test_metrics_at_050"][best_name]
    mt_best = res["test_metrics_at_tstar"][best_name]
    ab_unw = res["ablation_study"]["Unweighted (T=0.50)"]
    ab_bal = res["ablation_study"]["Class-Weighted Balanced (T=0.50)"]
    ab_opt = res["ablation_study"][f"Cost-Sensitive Optimal (T*={t_star:.2f})"]
    err05 = res["error_analysis"]["at_default_050"]
    errt = res["error_analysis"]["at_optimal_tstar"]
    roi = res["roi_benchmark"]
    case_hi = res["shap_cases"]["high_risk_churn"]
    case_lo = res["shap_cases"]["safe_non_churn"]

    # ==================== PHAN MO DAU VA CAC DANH MUC ====================
    add_h1("BẢNG PHÂN CÔNG NHIỆM VỤ NHÓM 02", page_break=False)
    add_para(
        "Bài tập lớn học phần Khai phá dữ liệu được thực hiện bởi 5 thành viên Nhóm 02 (Lớp Cao học CNTT 20.2) "
        "trên cơ sở phân chia đều các mảng chuyên môn theo chu trình KDD. Nội dung công việc phụ trách và tỷ lệ đóng góp "
        "của từng thành viên được tổng hợp trong Bảng 0.1."
    )
    add_table(
        ["STT", "Họ và tên", "Mảng phụ trách chính", "Nội dung công việc thực hiện", "Đóng góp"],
        [
            ["1", "Đỗ Ngọc Điền", "Khung KDD & Tổng hợp", "Xây dựng khung đề tài theo quy trình KDD, viết Chương 1 và Kết luận, tổng hợp và rà soát toàn bộ báo cáo.", "20%"],
            ["2", "Dương Tùng Lâm", "EDA & Đặc trưng miền", "Khám phá dữ liệu (EDA), kiểm định thống kê Chi-Square và Mann-Whitney U, tạo 6 đặc trưng miền và viết Chương 2.", "20%"],
            ["3", "Trần Thị Thu Hương", "Pipeline & Mô hình", "Xây dựng pipeline chống rò rỉ, huấn luyện và tinh chỉnh 5 mô hình phân loại qua 5-Fold CV, viết Chương 3.", "20%"],
            ["4", "Phạm Văn Chiến", "Tối ưu ngưỡng & SHAP", "Tối ưu ngưỡng chi phí, trích xuất luật CART, giải thích SHAP toàn cục/cục bộ và phân tích lỗi, viết Chương 4.", "20%"],
            ["5", "Vàng A Hứ", "Ứng dụng & Hiệu quả KT", "Đối chứng chi phí 4 kịch bản vận hành, xây dựng ứng dụng minh họa Streamlit, viết Chương 5 và biên tập Slide.", "20%"],
        ],
        col_widths=[0.48, 1.38, 1.32, 2.60, 0.72],
        caption="Bảng 0.1: Phân công nhiệm vụ các thành viên Nhóm 02",
    )

    add_h1("LỜI MỞ ĐẦU", page_break=True)
    add_para(
        "Trong ngành viễn thông và các dịch vụ thuê bao định kỳ, tỷ lệ khách hàng chấm dứt sử dụng dịch vụ (Customer Churn) "
        "ảnh hưởng trực tiếp đến doanh thu và chi phí tăng trưởng. Các nghiên cứu thực nghiệm về quản trị quan hệ khách hàng "
        "cho thấy chi phí thu hút một thuê bao mới thường cao gấp 5 đến 7 lần chi phí duy trì một khách hàng hiện hữu. "
        "Do đó, việc khai phá dữ liệu lịch sử để nhận diện sớm các thuê bao có nguy cơ rời bỏ, đồng thời chỉ rõ nguyên nhân "
        "dẫn đến hành vi này là một bài toán phân loại dữ liệu mang tính thực tiễn cao."
    )
    add_para(
        "Bám sát nội dung Chủ đề 2 của học phần Khai phá dữ liệu (Phân loại dữ liệu trong chăm sóc khách hàng), "
        "Nhóm 02 thực hiện đề tài “Ứng dụng các mô hình phân loại dữ liệu trong dự đoán nguy cơ rời bỏ dịch vụ của khách hàng” "
        "trên tập dữ liệu thực tế IBM Telco Customer Churn gồm 7.043 thuê bao. Mục tiêu của bài tập lớn không dừng lại ở việc "
        "huấn luyện mô hình dự đoán, mà tuân thủ đầy đủ quy trình Khai phá tri thức từ dữ liệu (KDD): kiểm định thống kê khám phá, "
        "xây dựng pipeline tiền xử lý chống rò rỉ dữ liệu, đối sánh 5 thuật toán phân loại thuộc các trường phái khác nhau "
        "(Logistic Regression, Decision Tree, Random Forest, XGBoost và LightGBM), tối ưu ngưỡng quyết định theo ma trận chi phí "
        "bất đối xứng, rút trích luật phân loại IF-THEN kết hợp giải thích SHAP và triển khai ứng dụng minh họa thực tế."
    )
    add_para(
        "Nhóm xin gửi lời cảm ơn chân thành đến TS. Bùi Ngọc Thăng đã hướng dẫn, định hướng chuyên môn và cung cấp các kiến thức "
        "nền tảng về Khai phá dữ liệu để nhóm hoàn thành bài tập lớn này. Mọi góp ý của thầy là cơ sở để nhóm hoàn thiện nghiên cứu tốt hơn."
    )

    add_h1("DANH MỤC TỪ VIẾT TẮT, BẢNG BIỂU VÀ HÌNH VẼ", page_break=True)
    add_h2("Danh mục ký hiệu và từ viết tắt")
    add_table(
        ["Ký hiệu / Viết tắt", "Thuật ngữ tiếng Anh", "Nghĩa tiếng Việt"],
        [
            ["AUC / ROC-AUC", "Area Under the ROC Curve", "Diện tích dưới đường cong đặc trưng hoạt động máy thu"],
            ["CART", "Classification and Regression Trees", "Cây phân loại và hồi quy"],
            ["CLV", "Customer Lifetime Value", "Giá trị vòng đời khách hàng"],
            ["CRISP-DM", "Cross-Industry Standard Process for Data Mining", "Quy trình chuẩn liên ngành cho khai phá dữ liệu"],
            ["CRM", "Customer Relationship Management", "Quản trị quan hệ khách hàng"],
            ["CV", "Cross-Validation", "Kiểm thử chéo"],
            ["EDA", "Exploratory Data Analysis", "Phân tích khám phá dữ liệu"],
            ["FN / FP", "False Negative / False Positive", "Âm tính giả (bỏ sót) / Dương tính giả (cảnh báo nhầm)"],
            ["KDD", "Knowledge Discovery in Databases", "Khai phá tri thức trong cơ sở dữ liệu"],
            ["LGBM / XGB", "LightGBM / eXtreme Gradient Boosting", "Các thuật toán tăng cường độ dốc cây quyết định"],
            ["OOF", "Out-Of-Fold", "Dự đoán ngoài lát cắt trong kiểm thử chéo"],
            ["PR-AUC", "Precision-Recall Area Under Curve", "Diện tích dưới đường cong Precision-Recall"],
            ["SHAP", "SHapley Additive exPlanations", "Phương pháp giải thích đóng góp đặc trưng theo giá trị Shapley"],
            ["TN / TP", "True Negative / True Positive", "Âm tính đúng / Dương tính đúng"],
        ],
        col_widths=[1.45, 2.55, 2.50],
    )

    add_h2("Danh mục bảng biểu")
    add_table(
        ["Số hiệu", "Tên bảng biểu"],
        [
            ["Bảng 0.1", "Phân công nhiệm vụ các thành viên Nhóm 02"],
            ["Bảng 2.1", "Mô tả 21 thuộc tính gốc của bộ dữ liệu IBM Telco Customer Churn"],
            ["Bảng 2.2", "Kết quả kiểm định Chi-Square và Cramér's V của Top 10 biến phân loại với Churn"],
            ["Bảng 2.3", "Kết quả kiểm định Mann-Whitney U cho các biến số liên tục"],
            ["Bảng 2.4", "Danh sách 6 đặc trưng tri thức miền được tạo mới"],
            ["Bảng 3.1", "Ma trận chi phí bất đối xứng trong bài toán giữ chân khách hàng"],
            ["Bảng 3.2", "So sánh hiệu năng 5 mô hình phân loại trên tập Test độc lập (N = 1.409)"],
            ["Bảng 3.3", "Thực nghiệm đối chứng 3 phương pháp xử lý mất cân bằng lớp trên LightGBM"],
            ["Bảng 3.4", "Top các hệ số hồi quy và Odds Ratio của mô hình Logistic Regression"],
            ["Bảng 4.1", "Hệ luật phân loại IF-THEN trích xuất từ cây quyết định CART cắt tỉa độ sâu 3"],
            ["Bảng 4.2", "Xếp hạng Top 10 đặc trưng có mức đóng góp SHAP trung bình lớn nhất"],
            ["Bảng 4.3", "Phân tích lỗi dự đoán của LightGBM tại ngưỡng mặc định 0,50 và ngưỡng tối ưu 0,21"],
            ["Bảng 4.4", "Đối chứng tổng chi phí và mức tiết kiệm của 4 chiến lược chăm sóc trên tập Test"],
            ["Bảng 5.1", "Khuyến nghị hành động chăm sóc khách hàng theo nhóm rủi ro khai phá"],
            ["Bảng A.1", "Bộ siêu tham số tối ưu của 5 mô hình sau Stratified 5-Fold GridSearchCV"],
        ],
        col_widths=[1.1, 5.4],
    )

    add_h2("Danh mục hình vẽ")
    add_table(
        ["Số hiệu", "Tên hình vẽ"],
        [
            ["Hình 1.1", "Sơ đồ tổng thể quy trình thực nghiệm khai phá dữ liệu theo chuẩn KDD và Anti-Leakage"],
            ["Hình 2.1", "Phân bố biến mục tiêu Churn và tỷ lệ rời bỏ theo loại hợp đồng"],
            ["Hình 2.2", "Phân bố tỷ lệ rời bỏ theo tenure, cước hàng tháng, hạ tầng Internet và phương thức thanh toán"],
            ["Hình 2.3", "Ma trận tương quan Pearson giữa các biến số và xếp hạng Cramér's V"],
            ["Hình 3.1", "So sánh đa chỉ số hiệu năng của 5 mô hình trên tập Test độc lập (N = 1.409)"],
            ["Hình 3.2", "Đường cong ROC-AUC và Precision-Recall (PR-AUC) của 5 mô hình"],
            ["Hình 3.3", "Ma trận nhầm lẫn của các mô hình tiêu biểu và tác động của ngưỡng tối ưu T* = 0,21"],
            ["Hình 3.4", "Đường cong tối ưu ngưỡng chi phí T* (A) và đối chứng tổng chi phí 4 chiến lược (B)"],
            ["Hình 4.1", "Sơ đồ cây quyết định CART cắt tỉa độ sâu 3 phục vụ rút trích luật IF-THEN"],
            ["Hình 4.2", "Biểu đồ SHAP Summary Plot toàn cục cho Top 12 đặc trưng (LightGBM)"],
            ["Hình 4.3", "Giải thích SHAP cục bộ cho hai hồ sơ điển hình trên tập Test"],
            ["Hình 5.1", "Giao diện tổng quan ứng dụng và biểu mẫu nhập thông số thuê bao tại Phân hệ 1 (Streamlit)"],
            ["Hình 5.2", "Kết quả dự đoán cá nhân, biểu đồ giải thích TreeSHAP cục bộ và khuyến nghị chăm sóc tại Phân hệ 1"],
            ["Hình 5.3", "Giao diện chấm điểm hàng loạt 100 thuê bao từ tệp CSV, phân tầng rủi ro và bảng xếp hạng tại Phân hệ 2"],
        ],
        col_widths=[1.1, 5.4],
    )

    add_h1("MỤC LỤC", page_break=True)
    add_toc_field(doc)

    # ==================== CHUONG 1 ====================
    add_h1("CHƯƠNG 1. TỔNG QUAN ĐỀ TÀI VÀ CƠ SỞ LÝ THUYẾT KHAI PHÁ DỮ LIỆU", page_break=True)

    add_h2("1.1. Đặt vấn đề và mục tiêu nghiên cứu")
    add_para(
        "Trong thị trường dịch vụ viễn thông có mức độ cạnh tranh cao, rào cản chuyển đổi nhà cung cấp ngày càng giảm "
        "khiến tỷ lệ khách hàng rời bỏ dịch vụ (Customer Churn) trở thành chỉ số sống còn đối với doanh nghiệp. "
        "Khác với tổn thất ngắn hạn của một giao dịch đơn lẻ, việc mất một thuê bao kéo theo sự sụt giảm toàn bộ dòng tiền "
        "định kỳ trong tương lai (Customer Lifetime Value – CLV), đồng thời buộc doanh nghiệp phải chi trả ngân sách tiếp thị "
        "lớn hơn nhiều lần để tìm kiếm thuê bao thay thế."
    )
    add_para(
        "Trên thực tế, hoạt động chăm sóc giữ chân khách hàng theo cách truyền thống thường vướng ba hạn chế cơ bản: "
        "(1) phản ứng bị động sau khi khách hàng đã yêu cầu hủy hợp đồng; (2) áp dụng các quy tắc lọc cứng thủ công "
        "trên bảng tính (ví dụ chỉ lọc theo cước phí cao) dẫn đến tỷ lệ bỏ sót lớn; và (3) triển khai chương trình khuyến mãi "
        "dàn trải cho toàn bộ tập khách hàng gây lãng phí ngân sách vào nhóm khách hàng vốn đã trung thành."
    )
    add_para(
        "Từ bài toán thực tế trên, nhóm xác định các mục tiêu cụ thể cho bài tập lớn môn Khai phá dữ liệu:"
    )
    add_bullet("Áp dụng đầy đủ quy trình Khai phá tri thức từ dữ liệu (KDD) trên bộ dữ liệu thực IBM Telco Customer Churn (7.043 thuê bao), đảm bảo không rò rỉ dữ liệu từ bước tiền xử lý đến đánh giá.")
    add_bullet("Thực hiện kiểm định thống kê khám phá và tạo các đặc trưng tri thức miền nhằm tăng cường khả năng phân tách hành vi rời bỏ.")
    add_bullet("Huấn luyện và đối sánh thực nghiệm 5 mô hình phân loại gồm Logistic Regression, Decision Tree (CART), Random Forest, XGBoost và LightGBM qua kiểm thử chéo phân tầng 5-Fold.")
    add_bullet("Giải quyết tính bất đối xứng chi phí bằng phương pháp tối ưu ngưỡng quyết định T* trên dữ liệu kiểm thử chéo (OOF) thay vì phụ thuộc ngưỡng mặc định 0,50.")
    add_bullet("Rút trích tri thức giải thích được thông qua hệ luật IF-THEN từ cây quyết định và giá trị Shapley (SHAP), kết hợp phân tích lỗi dự đoán, đối chứng chi phí và triển khai ứng dụng minh họa.")

    add_h2("1.2. Quy trình Khai phá tri thức từ dữ liệu (KDD và CRISP-DM)")
    add_para(
        "Theo Fayyad và cộng sự [1] cùng Han, Kamber và Pei [2], Khai phá dữ liệu (Data Mining) là bước cốt lõi nằm trong "
        "chu trình Khai phá tri thức trong cơ sở dữ liệu (KDD). Một quy trình chuẩn gồm năm giai đoạn liên kết chặt chẽ: "
        "Lựa chọn dữ liệu (Selection), Tiền xử lý và làm sạch (Preprocessing), Biến đổi và tạo đặc trưng (Transformation), "
        "Khai phá mô hình (Data Mining), và Đánh giá – Rút trích tri thức (Evaluation & Knowledge Representation). "
        "Trong công nghiệp, quy trình này được cụ thể hóa theo chuẩn CRISP-DM gồm 6 bước lặp từ Hiểu nghiệp vụ đến Triển khai."
    )
    add_para(
        "Trong đề tài này, bài toán dự đoán rời bỏ khách hàng được mô hình hóa thành bài toán phân loại nhị phân có giám sát. "
        "Cho tập dữ liệu lịch sử D = {(x_i, y_i)}, i = 1..N, trong đó x_i là vectơ đặc trưng mô tả thông tin nhân khẩu học, "
        "dịch vụ đăng ký và hợp đồng của thuê bao thứ i, còn y_i thuộc {0, 1} là nhãn mục tiêu (0: Ở lại, 1: Rời bỏ). "
        "Mục tiêu là học một hàm ước lượng xác suất hậu nghiệm P(Y = 1 | X = x) sao cho tối ưu cả năng lực xếp hạng "
        "rủi ro lẫn tổng chi phí ra quyết định nghiệp vụ. Kiến trúc tổng thể 7 bước của quy trình thực nghiệm trong đề tài "
        "được minh họa tại Hình 1.1."
    )
    add_figure(
        "fig00_kdd_pipeline_flowchart.png",
        "Hình 1.1: Sơ đồ tổng thể quy trình thực nghiệm khai phá dữ liệu theo chuẩn KDD và Anti-Leakage",
        width_in=6.45,
    )

    add_h2("1.3. Cơ sở toán học các thuật toán phân loại áp dụng")
    add_h3("1.3.1. Hồi quy Logistic (Logistic Regression)")
    add_para(
        "Hồi quy Logistic là mô hình phân loại tuyến tính nền tảng, ánh xạ tổ hợp tuyến tính z = w^T x + b sang xác suất "
        "thuộc lớp dương thông qua hàm Sigmoid:"
    )
    # Cong thuc (1.1): P(Y = 1 | x) = sigma(w^T x + b) = 1 / (1 + exp(-(w^T x + b)))
    wtx_b = msup(mr("w", sty="bi"), mr("T", nor=True)) + mr("x", sty="bi") + mr(" + ", nor=True) + mr("b")
    eq_1_1 = (
        mr("P", nor=True)
        + mdelim(mr("Y") + mr(" = 1 | ", nor=True) + mr("x", sty="bi"))
        + mr(" = ", nor=True)
        + mr("σ")
        + mdelim(wtx_b)
        + mr(" = ", nor=True)
        + mfrac(
            mr("1", nor=True),
            mr("1 + exp", nor=True) + mdelim(mr("−", nor=True) + mdelim(wtx_b)),
        )
    )
    add_omml_equation(eq_1_1, "1.1")

    add_para(
        "Bộ tham số (w, b) được ước lượng bằng phương pháp hợp lý cực đại, tương đương tối thiểu hóa hàm mất mát "
        "Binary Cross-Entropy (Log-Loss) có thành phần chính quy hóa L2 nhằm hạn chế quá khớp khi có đa cộng tuyến:"
    )
    # Cong thuc (1.2): L(w, b) = -(1/N) * Sum_{i=1}^N [ y_i log(p_i) + (1 - y_i) log(1 - p_i) ] + (1 / 2C) ||w||_2^2
    yi = msub(mr("y"), mr("i"))
    pi = msub(mr("p"), mr("i"))
    ce_term = (
        yi + mr(" log", nor=True) + mdelim(pi)
        + mr(" + ", nor=True)
        + mdelim(mr("1 − ", nor=True) + yi)
        + mr(" log", nor=True)
        + mdelim(mr("1 − ", nor=True) + pi)
    )
    eq_1_2 = (
        mr("L")
        + mdelim(mr("w", sty="bi") + mr(", ", nor=True) + mr("b"))
        + mr(" = −", nor=True)
        + mfrac(mr("1", nor=True), mr("N"))
        + msum(mr("i") + mr("=1", nor=True), mr("N"), mdelim(ce_term, "[", "]"))
        + mr(" + ", nor=True)
        + mfrac(mr("1", nor=True), mr("2", nor=True) + mr("C"))
        + msubsup(mdelim(mr("w", sty="bi"), "‖", "‖"), mr("2", nor=True), mr("2", nor=True))
    )
    add_omml_equation(eq_1_2, "1.2")

    add_para(
        "Ưu điểm lớn của Logistic Regression là tốc độ tính toán nhanh, xác suất dự báo có độ hiệu chuẩn tốt và cho phép "
        "diễn giải trực tiếp tác động của từng đặc trưng thông qua tỷ số chênh Odds Ratio = exp(w_j)."
    )

    add_h3("1.3.2. Cây quyết định (Decision Tree – CART)")
    add_para(
        "Thuật toán CART xây dựng cây nhị phân bằng cách phân hoạch đệ quy không gian đặc trưng. Tại mỗi nút t chứa tập mẫu D_t, "
        "thuật toán chọn cặp (thuộc tính j, ngưỡng s) làm giảm mạnh nhất độ vấy bẩn Gini (hoặc Entropy) sau khi tách:"
    )
    # Cong thuc (1.3): Gini(D_t) = 1 - (p_0^2 + p_1^2) = 2 p_1 (1 - p_1)
    dt = msub(mr("D"), mr("t"))
    p0_sq = msubsup(mr("p"), mr("0", nor=True), mr("2", nor=True))
    p1_sq = msubsup(mr("p"), mr("1", nor=True), mr("2", nor=True))
    p1_sub = msub(mr("p"), mr("1", nor=True))
    eq_1_3 = (
        mr("Gini", nor=True) + mdelim(dt)
        + mr(" = 1 − ", nor=True)
        + mdelim(p0_sq + mr(" + ", nor=True) + p1_sq)
        + mr(" = 2", nor=True) + p1_sub + mdelim(mr("1 − ", nor=True) + p1_sub)
    )
    add_omml_equation(eq_1_3, "1.3")

    # Cong thuc (1.4): Delta Gini(D_t, j, s) = Gini(D_t) - (|D_L|/|D_t|) Gini(D_L) - (|D_R|/|D_t|) Gini(D_R)
    dl = msub(mr("D"), mr("L"))
    dr = msub(mr("D"), mr("R"))
    eq_1_4 = (
        mr("ΔGini", nor=True) + mdelim(dt + mr(", ", nor=True) + mr("j") + mr(", ", nor=True) + mr("s"))
        + mr(" = Gini", nor=True) + mdelim(dt)
        + mr(" − ", nor=True)
        + mfrac(mdelim(dl, "|", "|"), mdelim(dt, "|", "|"))
        + mr("Gini", nor=True) + mdelim(dl)
        + mr(" − ", nor=True)
        + mfrac(mdelim(dr, "|", "|"), mdelim(dt, "|", "|"))
        + mr("Gini", nor=True) + mdelim(dr)
    )
    add_omml_equation(eq_1_4, "1.4")

    add_para(
        "Cây quyết định có thể tự động mô hình hóa tương tác phi tuyến và chuyển trực tiếp mỗi nhánh từ gốc tới lá thành một "
        "luật phân loại dạng IF-THEN tường minh. Tuy nhiên, cây đơn lẻ dễ bị quá khớp và có phương sai cao, do đó cần kiểm soát "
        "bằng các tham số tiền cắt tỉa (pre-pruning) như max_depth và min_samples_leaf."
    )

    add_h3("1.3.3. Rừng ngẫu nhiên (Random Forest)")
    add_para(
        "Random Forest [3] thuộc họ Học kết hợp theo cơ chế Bagging (Bootstrap Aggregating). Thuật toán huấn luyện song song "
        "M cây quyết định trên các tập mẫu lấy lại có hoàn lại (bootstrap samples), đồng thời tại mỗi nút chia chỉ xét một tập con "
        "ngẫu nhiên gồm m = sqrt(d) đặc trưng. Việc giảm tương quan giữa các cây thành phần giúp giảm đáng kể phương sai dự báo "
        "và hạn chế hiện tượng quá khớp so với một cây CART đơn lẻ. Xác suất cuối cùng là trung bình cộng xác suất từ M cây."
    )

    add_h3("1.3.4. Tăng cường độ dốc: XGBoost và LightGBM")
    add_para(
        "Khác với Bagging, các thuật toán Gradient Boosted Decision Trees (GBDT) xây dựng các cây tuần tự, trong đó cây thứ k "
        "học cách bù đắp phần sai số còn lại của tổng (k − 1) cây trước đó. Với XGBoost [4], hàm mục tiêu tại vòng lặp k "
        "được xấp xỉ khai triển Taylor bậc hai sử dụng cả đạo hàm bậc nhất g_i và đạo hàm bậc hai h_i của hàm mất mát, kết hợp "
        "hàm phạt cấu trúc cây Omega(f_k) để kiểm soát độ phức tạp:"
    )
    # Cong thuc (1.5): Obj^(k) ≈ Sum_{i=1}^N [ g_i f_k(x_i) + (1/2) h_i f_k^2(x_i) ] + Omega(f_k)
    xi = msub(mr("x", sty="bi"), mr("i"))
    fk = msub(mr("f"), mr("k"))
    fk_sq = msubsup(mr("f"), mr("k"), mr("2", nor=True))
    gbdt_term = (
        msub(mr("g"), mr("i")) + fk + mdelim(xi)
        + mr(" + ", nor=True)
        + mfrac(mr("1", nor=True), mr("2", nor=True))
        + msub(mr("h"), mr("i")) + fk_sq + mdelim(xi)
    )
    eq_1_5 = (
        msup(mr("Obj", nor=True), mdelim(mr("k")))
        + mr(" ≈ ", nor=True)
        + msum(mr("i") + mr("=1", nor=True), mr("N"), mdelim(gbdt_term, "[", "]"))
        + mr(" + Ω", nor=True) + mdelim(fk)
    )
    add_omml_equation(eq_1_5, "1.5")

    add_para(
        "LightGBM [5] cải tiến tốc độ và hiệu quả bộ nhớ của GBDT thông qua kỹ thuật gom nhóm giá trị đặc trưng dạng biểu đồ tần suất "
        "(Histogram-based splitting) và chiến lược phát triển cây theo nút lá có độ giảm mất mát lớn nhất (Leaf-wise growth) "
        "có ràng buộc giới hạn độ sâu max_depth và số lá num_leaves nhằm tránh quá khớp trên dữ liệu cỡ trung bình."
    )

    add_h2("1.4. Độ đo đánh giá, học nhạy chi phí và giải thích SHAP")
    add_para(
        "Trên dữ liệu mất cân bằng lớp (lớp Churn chỉ chiếm 26,54%), độ chính xác tổng thể (Accuracy) không phản ánh đầy đủ "
        "năng lực nhận diện khách hàng rời bỏ. Do đó, nghiên cứu sử dụng đồng thời hệ thống độ đo sau:"
    )
    add_bullet("Precision = TP / (TP + FP): tỷ lệ khách hàng thực sự rời bỏ trong số các trường hợp bị cảnh báo.")
    add_bullet("Recall = TP / (TP + FN): tỷ lệ khách hàng rời bỏ thực tế được mô hình phát hiện đúng (độ nhạy).")
    add_bullet("F1-Score = 2 × Precision × Recall / (Precision + Recall): trung bình điều hòa cân bằng giữa Precision và Recall.")
    add_bullet("ROC-AUC và PR-AUC: đánh giá năng lực phân tách tổng thể trên mọi ngưỡng xác suất; trong đó PR-AUC đặc biệt nhạy với hiệu năng của lớp thiểu số.")
    add_bullet("Brier Score = (1/N) × Σ(p_i − y_i)^2: đo mức độ hiệu chuẩn xác suất (càng nhỏ xác suất dự đoán càng sát tần suất thực tế).")
    add_para(
        "Về mặt ra quyết định nhạy chi phí (Cost-Sensitive Learning) [6], khi chi phí bỏ sót một khách hàng rời bỏ (C_FN) lớn hơn "
        "nhiều so với chi phí gửi ưu đãi giữ chân (C_FP), ngưỡng quyết định mặc định T = 0,50 không còn tối ưu. Theo lý thuyết "
        "quyết định Bayes, với ma trận chi phí (C_TN = 0, C_FP, C_TP, C_FN), hành động can thiệp là tối ưu khi chi phí kỳ vọng "
        "của việc can thiệp nhỏ hơn hoặc bằng chi phí kỳ vọng khi không can thiệp:"
    )
    # Cong thuc (1.6): p C_TP + (1 - p) C_FP <= p C_FN + (1 - p) C_TN  =>  T*_Bayes = C_FP / (C_FN - C_TP + C_FP)
    ctp = msub(mr("C"), mr("TP", nor=True))
    cfp = msub(mr("C"), mr("FP", nor=True))
    cfn = msub(mr("C"), mr("FN", nor=True))
    ctn = msub(mr("C"), mr("TN", nor=True))
    eq_1_6 = (
        mr("p") + mr("·", nor=True) + ctp
        + mr(" + ", nor=True) + mdelim(mr("1 − ", nor=True) + mr("p")) + mr("·", nor=True) + cfp
        + mr(" ≤ ", nor=True)
        + mr("p") + mr("·", nor=True) + cfn
        + mr(" + ", nor=True) + mdelim(mr("1 − ", nor=True) + mr("p")) + mr("·", nor=True) + ctn
        + mr("   ⇒   ", nor=True)
        + msubsup(mr("T"), mr("Bayes", nor=True), mr("*", nor=True))
        + mr(" = ", nor=True)
        + mfrac(cfp, cfn + mr(" − ", nor=True) + ctp + mr(" + ", nor=True) + cfp)
    )
    add_omml_equation(eq_1_6, "1.6")

    add_para(
        "Để giải thích đầu ra của các mô hình tổ hợp cây, đề tài sử dụng phương pháp SHAP (SHapley Additive exPlanations) [7] "
        "dựa trên lý thuyết trò chơi hợp tác. Giá trị Shapley φ_j phân bổ công bằng phần đóng góp biên của đặc trưng j vào "
        "độ lệch giữa dự đoán f(x) và giá trị kỳ vọng cơ sở E[f(X)], thỏa mãn tính chất cộng tính chính xác. "
        "Thuật toán TreeSHAP cho phép tính chính xác giá trị này trong thời gian đa thức trên các mô hình cây như Random Forest, XGBoost và LightGBM."
    )

    # ==================== CHUONG 2 ====================
    add_h1("CHƯƠNG 2. PHÂN TÍCH DỮ LIỆU VÀ QUY TRÌNH TIỀN XỬ LÝ", page_break=True)

    add_h2("2.1. Đặc tả bộ dữ liệu IBM Telco Customer Churn")
    add_para(
        f"Bộ dữ liệu thực nghiệm gồm {fmt_int(audit['n_rows'])} bản ghi thuê bao viễn thông và {audit['n_raw_cols']} cột thuộc tính "
        f"(1 cột mã định danh customerID, 19 biến đặc trưng đầu vào và 1 biến mục tiêu Churn). Kiểm tra chất lượng dữ liệu gốc "
        f"cho thấy không có bản ghi trùng lặp ({audit['duplicate_rows']} dòng trùng), nhưng cột TotalCharges bị lưu dưới dạng "
        f"chuỗi văn bản và chứa {audit['missing_total_charges']} ô khoảng trắng. Kiểm tra chi tiết xác nhận cả 11 thuê bao này "
        f"đều có số tháng sử dụng tenure = 0 (thuê bao mới hòa mạng chưa phát sinh kỳ cước tích lũy), do đó việc chuyển đổi kiểu "
        f"số thực và điền giá trị 0,0 là hoàn toàn phù hợp bản chất nghiệp vụ."
    )
    add_table(
        ["Nhóm thuộc tính", "Tên cột (Attributes)", "Kiểu dữ liệu", "Ý nghĩa nghiệp vụ"],
        [
            ["Định danh (1)", "customerID", "Chuỗi", "Mã thuê bao duy nhất (loại bỏ trước khi huấn luyện)."],
            ["Nhân khẩu học (4)", "gender, SeniorCitizen, Partner, Dependents", "Phân loại / Nhị phân", "Giới tính, khách cao tuổi (0/1), có bạn đời, có người phụ thuộc."],
            ["Dịch vụ viễn thông (9)", "PhoneService, MultipleLines, InternetService, OnlineSecurity, OnlineBackup, DeviceProtection, TechSupport, StreamingTV, StreamingMovies", "Phân loại (2–3 mức)", "Hạ tầng điện thoại, Internet (DSL / Fiber optic / No) và 6 dịch vụ gia tăng đi kèm."],
            ["Hợp đồng & Cước (6)", "tenure, Contract, PaperlessBilling, PaymentMethod, MonthlyCharges, TotalCharges", "Số / Phân loại", "Số tháng gắn bó (0–72), loại hợp đồng (tháng/1 năm/2 năm), hóa đơn điện tử, thanh toán, cước tháng và tổng cước."],
            ["Biến mục tiêu (1)", "Churn", "Nhị phân (Yes/No)", "Nhãn rời bỏ dịch vụ trong tháng gần nhất (1 = Yes, 0 = No)."],
        ],
        col_widths=[1.25, 2.25, 1.05, 1.95],
        caption="Bảng 2.1: Mô tả 21 thuộc tính gốc của bộ dữ liệu IBM Telco Customer Churn",
    )

    add_h2("2.2. Phân tích khám phá dữ liệu (EDA) và kiểm định thống kê")
    add_para(
        f"Trong tổng số {fmt_int(audit['n_rows'])} khách hàng, có {fmt_int(audit['non_churn_count'])} khách ở lại ({fmt_pct((1-audit['churn_rate'])*100, 2)}) "
        f"và {fmt_int(audit['churn_count'])} khách rời bỏ ({fmt_pct(audit['churn_rate']*100, 2)}), tạo tỷ lệ mất cân bằng xấp xỉ 2,77 : 1. "
        f"Hình 2.1 và Hình 2.2 minh họa rõ sự chênh lệch lớn về tỷ lệ rời bỏ giữa các phân khúc hợp đồng, thời gian gắn bó, "
        f"mức cước hàng tháng, hạ tầng Internet và phương thức thanh toán."
    )
    add_figure("fig01_churn_distribution.png", "Hình 2.1: Phân bố biến mục tiêu Churn và tỷ lệ rời bỏ theo loại hợp đồng")

    ct_map = {x["category"]: x for x in eda["group_breakdown"]["Contract"]}
    inet_map = {x["category"]: x for x in eda["group_breakdown"]["InternetService"]}
    pm_map = {x["category"]: x for x in eda["group_breakdown"]["PaymentMethod"]}
    add_para(
        f"Nhóm khách hàng ký hợp đồng theo tháng (Month-to-month, n = {fmt_int(ct_map['Month-to-month']['count'])}) có tỷ lệ rời bỏ lên tới "
        f"{fmt_pct(ct_map['Month-to-month']['churn_rate_pct'], 1)}, cao gấp gần 4 lần nhóm hợp đồng 1 năm "
        f"({fmt_pct(ct_map['One year']['churn_rate_pct'], 1)}) và gấp hơn 15 lần nhóm hợp đồng 2 năm "
        f"({fmt_pct(ct_map['Two year']['churn_rate_pct'], 1)}). Tương tự, thuê bao dùng cáp quang Fiber optic có tỷ lệ rời bỏ "
        f"{fmt_pct(inet_map['Fiber optic']['churn_rate_pct'], 1)} (so với {fmt_pct(inet_map['DSL']['churn_rate_pct'], 1)} ở nhóm DSL), "
        f"và nhóm thanh toán qua Electronic check có tỷ lệ rời bỏ {fmt_pct(pm_map['Electronic check']['churn_rate_pct'], 1)} "
        f"so với chỉ {fmt_pct(pm_map['Credit card (automatic)']['churn_rate_pct'], 1)} ở nhóm thanh toán tự động qua thẻ tín dụng."
    )
    add_figure("fig02_eda_drivers.png", "Hình 2.2: Phân bố tỷ lệ rời bỏ theo tenure, cước hàng tháng, hạ tầng Internet và phương thức thanh toán")

    add_para(
        "Để kiểm chứng khách quan các quan sát trực quan thay vì kết luận cảm tính, nhóm thực hiện kiểm định phi tham số "
        "Chi-Square (kèm hệ số đo độ mạnh liên kết Cramér's V) cho các biến phân loại và kiểm định Mann-Whitney U cho các biến số liên tục. "
        "Kết quả trình bày tại Bảng 2.2, Bảng 2.3 và Hình 2.3."
    )
    chi2_rows = [
        [str(i), r["feature"], fmt_dec(r["chi2"], 2), fmt_sci(r["p_value"]), fmt_dec(r["cramers_v"], 4)]
        for i, r in enumerate(eda["chi2_top12"][:10], start=1)
    ]
    add_table(
        ["STT", "Thuộc tính phân loại", "Chi-Square", "p-value", "Cramér's V"],
        chi2_rows,
        col_widths=[0.5, 2.3, 1.2, 1.3, 1.2],
        caption="Bảng 2.2: Kết quả kiểm định Chi-Square và Cramér's V của Top 10 biến phân loại với Churn",
    )

    mw_rows = [
        [
            r["feature"],
            fmt_dec(r["median_non_churn"], 2),
            fmt_dec(r["median_churn"], 2),
            fmt_dec(r["mean_non_churn"], 2),
            fmt_dec(r["mean_churn"], 2),
            fmt_sci(r["p_value"]),
        ]
        for r in eda["mann_whitney"]
    ]
    add_table(
        ["Biến số", "Trung vị (Ở lại)", "Trung vị (Rời bỏ)", "Trung bình (Ở lại)", "Trung bình (Rời bỏ)", "p-value"],
        mw_rows,
        col_widths=[1.45, 0.95, 1.0, 1.0, 1.0, 1.1],
        caption="Bảng 2.3: Kết quả kiểm định Mann-Whitney U cho các biến số liên tục",
    )
    add_figure("fig03_correlation_cramersv.png", "Hình 2.3: Ma trận tương quan Pearson giữa các biến số và xếp hạng Cramér's V")

    add_note_box(
        f"Kết quả kiểm định xác nhận 5 đặc trưng phân loại có liên kết mạnh nhất với Churn là Contract (V = 0,4101), "
        f"TenureGroup (V = 0,3629), OnlineSecurity (V = 0,3474), TechSupport (V = 0,3429) và InternetService (V = 0,3225) "
        f"với p-value < 10^-150. Về biến số, trung vị tenure của nhóm rời bỏ chỉ là 10 tháng so với 38 tháng ở nhóm ở lại, "
        f"trong khi cước tháng trung vị của nhóm rời bỏ cao hơn đáng kể (79,65 USD so với 64,43 USD). Riêng biến ChargeDiff "
        f"có p = 0,6935 (không khác biệt có ý nghĩa thống kê đơn biến). Ngoài ra, hệ số tương quan giữa tenure và TotalCharges "
        f"đạt r = {fmt_dec(eda['pearson_tenure_totalcharges'], 2)}, cho thấy hiện tượng đa cộng tuyến mạnh — là cơ sở cần thiết để dùng "
        f"chính quy hóa L2 cho Logistic Regression và ưu tiên mô hình cây tổ hợp."
    )

    add_h2("2.3. Kỹ thuật tạo đặc trưng miền và pipeline tiền xử lý chống rò rỉ")
    add_para(
        "Từ phân tích khám phá, nhóm bổ sung 6 đặc trưng tri thức miền (Bảng 2.4), nâng tổng không gian đầu vào từ 19 biến gốc "
        "lên đúng 25 đặc trưng (sau khi biến đổi qua ColumnTransformer thu được 55 chiều đặc trưng số)."
    )
    add_table(
        ["STT", "Tên đặc trưng mới", "Công thức / Cách xác định", "Ý nghĩa khai phá"],
        [
            ["1", "AvgMonthlySpend", "TotalCharges / max(tenure, 1)", "Đo mức chi tiêu bình quân thực tế theo vòng đời thuê bao."],
            ["2", "ChargeDiff", "MonthlyCharges − AvgMonthlySpend", "Kiểm tra biến động chênh lệch giữa cước hiện tại và cước bình quân tích lũy."],
            ["3", "TotalAddonServices", "Tổng số dịch vụ gia tăng = 'Yes' (0–6)", "Phản ánh mức độ gắn kết hệ sinh thái dịch vụ của thuê bao."],
            ["4", "HasProtectionBundle", "1 nếu dùng OnlineSecurity hoặc TechSupport", "Nhận diện nhóm khách có rào cản chuyển đổi nhờ dịch vụ bảo vệ/hỗ trợ."],
            ["5", "IsAutoPayment", "1 nếu PaymentMethod chứa 'automatic'", "Phân tách nhóm thanh toán tự động (ổn định) với thanh toán thủ công."],
            ["6", "TenureGroup", "Chia 5 khoảng: 0–6, 7–12, 13–24, 25–48, >48 tháng", "Nắm bắt quan hệ phi tuyến theo từng giai đoạn vòng đời khách hàng."],
        ],
        col_widths=[0.45, 1.50, 2.15, 2.40],
        caption="Bảng 2.4: Danh sách 6 đặc trưng tri thức miền được tạo mới",
    )
    add_para(
        f"Để ngăn chặn tuyệt đối rò rỉ dữ liệu (Data Leakage), tập dữ liệu được phân tách phân tầng (Stratified Split) theo tỷ lệ "
        f"80% huấn luyện ({fmt_int(res['split']['train_size'])} mẫu) và 20% kiểm thử độc lập ({fmt_int(res['split']['test_size'])} mẫu) "
        f"với random_state = 42 trước khi ước lượng bất kỳ tham số thống kê nào. Toàn bộ 25 đặc trưng đầu vào được chia thành ba nhánh "
        f"xử lý riêng biệt bên trong ColumnTransformer của scikit-learn Pipeline: (1) 6 biến số liên tục (tenure, MonthlyCharges, "
        f"TotalCharges, AvgMonthlySpend, ChargeDiff, TotalAddonServices) được điền khuyết theo trung vị (SimpleImputer) và chuẩn hóa Z-score "
        f"(StandardScaler), giữ 6 chiều; (2) 3 biến nhị phân 0/1 (SeniorCitizen, HasProtectionBundle, IsAutoPayment) được điền khuyết theo "
        f"tần suất lớn nhất và giữ nguyên thang đo nhị phân, tạo 3 chiều; và (3) 16 biến phân loại (15 biến phân loại gốc cùng TenureGroup) "
        f"được mã hóa One-Hot (OneHotEncoder với handle_unknown='ignore'), sinh ra 46 chiều. Tổng hợp ba nhánh (6 + 3 + 16 = 25 đặc trưng "
        f"đầu vào) thu được chính xác 6 + 3 + 46 = 55 chiều đặc trưng số, trong đó mọi phép fit() chỉ chạy trên lát cắt huấn luyện của từng vòng 5-Fold CV."
    )

    # ==================== CHUONG 3 ====================
    add_h1("CHƯƠNG 3. THỰC NGHIỆM MÔ HÌNH PHÂN LOẠI VÀ TỐI ƯU NGƯỠNG CHI PHÍ", page_break=True)

    add_h2("3.1. Thiết lập kịch bản thực nghiệm và ma trận chi phí")
    add_para(
        "Cả 5 mô hình đại diện cho các trường phái phân loại khác nhau gồm 2 mô hình cơ sở (Logistic Regression, Decision Tree CART) "
        "và 3 mô hình tổ hợp cây (Random Forest, XGBoost, LightGBM) đều được tinh chỉnh siêu tham số bằng GridSearchCV trên "
        "Stratified 5-Fold Cross-Validation với tiêu chí tối ưu là ROC-AUC. Toàn bộ thực nghiệm chạy cục bộ trên CPU với thời gian "
        f"tổng cộng {fmt_dec(res['total_runtime_sec'], 1)} giây, đảm bảo khả năng tái lập hoàn toàn."
    )
    add_para(
        "Nhằm gắn kết đánh giá kỹ thuật với hiệu quả nghiệp vụ giữ chân khách hàng, nhóm xây dựng ma trận chi phí bất đối xứng "
        "theo tỷ lệ tiêu chuẩn ngành viễn thông 5 : 1 (Bảng 3.1): chi phí bỏ sót một thuê bao rời bỏ (C_FN = 500 USD, phản ánh "
        "tổn thất doanh thu vòng đời) cao gấp 5 lần chi phí triển khai ưu đãi/chăm sóc giữ chân (C_FP = C_TP = 100 USD). "
        "Theo công thức (1.6), ngưỡng quyết định lý thuyết Bayes tương ứng là T*_Bayes = 100 / (500 − 100 + 100) = 0,20."
    )
    add_table(
        ["Trạng thái thực tế \\ Dự đoán", "Dự đoán Ở lại (0 – Không can thiệp)", "Dự đoán Rời bỏ (1 – Gửi ưu đãi CSKH)"],
        [
            ["Thực tế Ở lại (0)", "TN: Chi phí = 0 USD (Quyết định đúng)", "FP: Chi phí = 100 USD (Lãng phí ưu đãi)"],
            ["Thực tế Rời bỏ (1)", "FN: Chi phí = 500 USD (Tổn thất mất khách)", "TP: Chi phí = 100 USD (Chi phí giữ chân thành công)"],
        ],
        col_widths=[1.9, 2.3, 2.3],
        caption="Bảng 3.1: Ma trận chi phí bất đối xứng trong bài toán giữ chân khách hàng",
    )

    add_h2("3.2. Kết quả đối sánh 5 mô hình trên tập Test độc lập")
    add_para(
        "Bảng 3.2 tổng hợp đầy đủ các chỉ số hiệu năng của 5 mô hình trên tập kiểm thử độc lập (N = 1.409), bao gồm cả đánh giá "
        f"tại ngưỡng mặc định T = 0,50 lẫn tổng chi phí tại ngưỡng tối ưu T* = {t_star_vn} (được chọn hoàn toàn trên dự đoán "
        "OOF của tập Train). Hình 3.1, Hình 3.2 và Hình 3.3 trực quan hóa sự so sánh này."
    )
    m_order = ["Logistic Regression", "Decision Tree (CART)", "Random Forest", "XGBoost", "LightGBM"]
    comp_rows = []
    for name in m_order:
        m05 = res["test_metrics_at_050"][name]
        mt = res["test_metrics_at_tstar"][name]
        comp_rows.append([
            name,
            fmt_dec(res["cv_roc_auc"][name], 4),
            fmt_pct(m05["accuracy"] * 100, 2),
            fmt_pct(m05["recall"] * 100, 2),
            fmt_dec(m05["f1"], 4),
            fmt_dec(m05["roc_auc"], 4),
            fmt_dec(m05["pr_auc"], 4),
            fmt_dec(m05["brier"], 4),
            fmt_usd(mt["total_cost_usd"]),
        ])
    add_table(
        ["Mô hình", "CV AUC", "Acc (0,50)", "Rec (0,50)", "F1 (0,50)", "Test AUC", "PR-AUC", "Brier", f"Chi phí (T*={t_star_vn})"],
        comp_rows,
        col_widths=[1.30, 0.62, 0.65, 0.65, 0.60, 0.64, 0.60, 0.56, 0.88],
        caption="Bảng 3.2: So sánh hiệu năng 5 mô hình phân loại trên tập Test độc lập (N = 1.409)",
    )
    add_figure("fig04_model_comparison.png", "Hình 3.1: So sánh đa chỉ số hiệu năng của 5 mô hình trên tập Test độc lập (N = 1.409)")
    add_figure("fig05_roc_pr_curves.png", "Hình 3.2: Đường cong ROC-AUC và Precision-Recall (PR-AUC) của 5 mô hình")

    add_para(
        f"Kết quả đối sánh cho thấy hai mô hình Gradient Boosting (LightGBM và XGBoost) đạt hiệu năng vượt trội nhất: "
        f"LightGBM dẫn đầu về CV ROC-AUC ({fmt_dec(res['cv_roc_auc']['LightGBM'], 4)}) và Test PR-AUC ({fmt_dec(res['test_metrics_at_050']['LightGBM']['pr_auc'], 4)}), "
        f"trong khi XGBoost đạt Test ROC-AUC cao nhất ({fmt_dec(res['test_metrics_at_050']['XGBoost']['roc_auc'], 4)}) và F1 cao nhất tại ngưỡng 0,50 "
        f"({fmt_dec(res['test_metrics_at_050']['XGBoost']['f1'], 4)}). Đáng chú ý, Logistic Regression sau khi được bổ sung 6 đặc trưng miền "
        f"(như TenureGroup và HasProtectionBundle) cũng đạt Test ROC-AUC = {fmt_dec(res['test_metrics_at_050']['Logistic Regression']['roc_auc'], 4)}, "
        f"tiệm cận sát các mô hình tổ hợp và bỏ xa Decision Tree đơn lẻ (Test ROC-AUC = {fmt_dec(res['test_metrics_at_050']['Decision Tree (CART)']['roc_auc'], 4)}). "
        f"Dựa trên tiêu chí lựa chọn mô hình nghiêm ngặt theo điểm kiểm thử chéo 5-Fold (không nhìn tập Test), LightGBM được chọn làm "
        f"mô hình chính cho các bước tối ưu chi phí và giải thích SHAP."
    )

    add_h2("3.3. Đối chứng xử lý lệch lớp và tối ưu ngưỡng nhạy chi phí")
    add_para(
        "Tại ngưỡng mặc định T = 0,50, dù cả 5 mô hình đều đạt Accuracy xấp xỉ 80%, độ nhạy (Recall) chỉ dao động quanh mức 51%–53%, "
        "tức bỏ sót gần một nửa số khách hàng thực sự rời bỏ (181/374 ca với LightGBM). Để xử lý vấn đề này, nhóm tiến hành "
        "thực nghiệm đối chứng (Ablation Study) giữa ba phương án trên cùng cấu hình LightGBM (Bảng 3.3): (1) giữ nguyên mặc định T = 0,50; "
        "(2) huấn luyện lại với trọng số lớp cân bằng class_weight='balanced'; và (3) giữ nguyên mô hình xác suất gốc kết hợp dịch chuyển "
        f"ngưỡng quyết định sang T* = {t_star_vn} tìm được trên đường cong chi phí OOF Train (Hình 3.4A)."
    )
    add_table(
        ["Phương pháp xử lý lệch lớp (LightGBM)", "Ngưỡng", "Accuracy", "Precision", "Recall", "F1-Score", "Brier Score", "Số ca FN", "Tổng chi phí"],
        [
            ["Mặc định (Unweighted)", "0,50", fmt_pct(ab_unw["accuracy"] * 100, 2), fmt_pct(ab_unw["precision"] * 100, 2), fmt_pct(ab_unw["recall"] * 100, 2), fmt_dec(ab_unw["f1"], 4), fmt_dec(ab_unw["brier"], 4), str(ab_unw["fn"]), fmt_usd(ab_unw["total_cost_usd"])],
            ["Gán trọng số lớp (Balanced)", "0,50", fmt_pct(ab_bal["accuracy"] * 100, 2), fmt_pct(ab_bal["precision"] * 100, 2), fmt_pct(ab_bal["recall"] * 100, 2), fmt_dec(ab_bal["f1"], 4), fmt_dec(ab_bal["brier"], 4), str(ab_bal["fn"]), fmt_usd(ab_bal["total_cost_usd"])],
            ["Tối ưu ngưỡng chi phí (T*)", t_star_vn, fmt_pct(ab_opt["accuracy"] * 100, 2), fmt_pct(ab_opt["precision"] * 100, 2), fmt_pct(ab_opt["recall"] * 100, 2), fmt_dec(ab_opt["f1"], 4), fmt_dec(ab_opt["brier"], 4), str(ab_opt["fn"]), fmt_usd(ab_opt["total_cost_usd"])],
        ],
        col_widths=[1.75, 0.55, 0.65, 0.65, 0.60, 0.60, 0.65, 0.45, 0.60],
        caption="Bảng 3.3: Thực nghiệm đối chứng 3 phương pháp xử lý mất cân bằng lớp trên LightGBM",
    )
    add_figure("fig06_confusion_matrices.png", "Hình 3.3: Ma trận nhầm lẫn của các mô hình tiêu biểu và tác động của ngưỡng tối ưu T* = 0,21")
    add_figure("fig07_cost_threshold_roi.png", "Hình 3.4: Đường cong tối ưu ngưỡng chi phí T* (A) và đối chứng tổng chi phí 4 chiến lược (B)")

    add_note_box(
        f"Kết quả Bảng 3.3 rút ra một phát hiện thực nghiệm rõ ràng: việc dùng class_weight='balanced' tuy kéo Recall tăng từ "
        f"51,60% lên 79,95%, nhưng làm méo đáng kể độ hiệu chuẩn xác suất (Brier Score tăng xấu từ 0,1348 lên 0,1630). "
        f"Ngược lại, phương pháp tối ưu ngưỡng chi phí trên xác suất chuẩn (T* = {t_star_vn}, trùng khớp rất sát ngưỡng lý thuyết Bayes 0,20) "
        f"vừa bảo toàn nguyên vẹn Brier Score tốt nhất (0,1348), vừa nâng Recall lên {fmt_pct(ab_opt['recall']*100, 2)} (giảm mạnh số ca bỏ sót FN "
        f"từ {ab_unw['fn']} xuống còn {ab_opt['fn']} ca) và đưa tổng chi phí về mức thấp nhất là {fmt_usd(ab_opt['total_cost_usd'])}."
    )

    add_h2("3.4. Phân tích định lượng hệ số Odds Ratio từ Logistic Regression")
    add_para(
        "Bên cạnh LightGBM, mô hình Logistic Regression cung cấp góc nhìn tham chiếu trực quan thông qua hệ số hồi quy và "
        "tỷ số chênh Odds Ratio = exp(w_j) (Bảng 3.4). Giữ cố định các biến khác, việc sử dụng cáp quang Fiber optic làm tăng "
        "tỷ số chênh rời bỏ lên 2,09 lần và hợp đồng theo tháng làm tăng 1,95 lần; ngược lại, hợp đồng 2 năm làm giảm tỷ số chênh "
        "xuống còn 0,41 lần (giảm ~59%) và thời gian gắn bó dài (tenure) là nhân tố bảo vệ mạnh thứ hai."
    )
    lr_rows = []
    for item in res["logistic_odds_ratios"]["top_increase_churn"][:5]:
        lr_rows.append([item["feature"], fmt_dec(item["coef"], 4, signed=True), fmt_dec(item["odds_ratio"], 4), "Tăng nguy cơ rời bỏ"])
    for item in res["logistic_odds_ratios"]["top_decrease_churn"][:5]:
        lr_rows.append([item["feature"], fmt_dec(item["coef"], 4, signed=True), fmt_dec(item["odds_ratio"], 4), "Giảm nguy cơ rời bỏ"])
    add_table(
        ["Đặc trưng (sau mã hóa)", "Hệ số hồi quy (w)", "Odds Ratio exp(w)", "Chiều hướng tác động"],
        lr_rows,
        col_widths=[2.2, 1.2, 1.3, 1.8],
        caption="Bảng 3.4: Top các hệ số hồi quy và Odds Ratio của mô hình Logistic Regression",
    )

    # ==================== CHUONG 4 ====================
    add_h1("CHƯƠNG 4. RÚT TRÍCH TRI THỨC, PHÂN TÍCH LỖI VÀ ĐÁNH GIÁ HIỆU QUẢ CHI PHÍ", page_break=True)

    add_h2("4.1. Rút trích hệ luật phân loại IF-THEN từ cây quyết định CART")
    add_para(
        "Để chuyển hóa kết quả khai phá dữ liệu thành tri thức nghiệp vụ dễ đọc cho nhân viên chăm sóc khách hàng, nhóm huấn luyện "
        "một cây quyết định CART cắt tỉa ở độ sâu 3 trên các thuộc tính gốc chưa chuẩn hóa (để đọc trực tiếp ngưỡng tháng và cước USD). "
        "Sơ đồ cây trình bày ở Hình 4.1 và toàn bộ 8 luật lá trích xuất được tổng hợp trong Bảng 4.1, sắp xếp theo xác suất rời bỏ giảm dần."
    )
    add_figure("fig08_decision_tree_rules.png", "Hình 4.1: Sơ đồ cây quyết định CART cắt tỉa độ sâu 3 phục vụ rút trích luật IF-THEN")

    rule_rows = []
    for idx, r in enumerate(res["extracted_cart_rules"], start=1):
        cond = (
            r["conditions"]
            .replace("Contract_Month-to-month > 0.50", "Hợp đồng = Theo tháng")
            .replace("Contract_Month-to-month <= 0.50", "Hợp đồng != Theo tháng")
            .replace("Contract_Two year > 0.50", "Hợp đồng = 2 năm")
            .replace("Contract_Two year <= 0.50", "Hợp đồng = 1 năm")
            .replace("InternetService_Fiber optic > 0.50", "Internet = Fiber optic")
            .replace("InternetService_Fiber optic <= 0.50", "Internet != Fiber optic")
            .replace("OnlineSecurity_No > 0.50", "OnlineSecurity = No")
            .replace("OnlineSecurity_No <= 0.50", "OnlineSecurity != No")
        )
        cond = re.sub(r"(\d)\.(\d)", r"\1,\2", cond)
        rule_rows.append([
            f"R{idx}",
            cond,
            f"{fmt_int(r['n_samples'])} ({fmt_pct(r['support_pct'], 1)})",
            fmt_pct(r["churn_prob_pct"], 1),
        ])
    add_table(
        ["Mã luật", "Điều kiện phân tách (IF ...)", "Số mẫu (Support)", "Xác suất Churn"],
        rule_rows,
        col_widths=[0.65, 3.65, 1.25, 0.95],
        caption="Bảng 4.1: Hệ luật phân loại IF-THEN trích xuất từ cây quyết định CART cắt tỉa độ sâu 3",
    )
    add_para(
        "Hệ luật ở Bảng 4.1 chỉ ra ba phân khúc hành vi rõ rệt: (1) Nhóm rủi ro rất cao (Luật R1): thuê bao hợp đồng theo tháng, "
        "dùng cáp quang Fiber optic và thời gian gắn bó <= 13,50 tháng chiếm 13,5% tổng khách hàng nhưng có tỷ lệ rời bỏ lên tới 71,0%; "
        "(2) Nhóm rủi ro trung bình – cao (Luật R2 và R3): khách hợp đồng theo tháng mới dùng <= 3,50 tháng (không dùng Fiber) "
        "hoặc dùng Fiber lâu hơn 13,50 tháng có tỷ lệ rời bỏ 42%–43%; (3) Nhóm trung thành cốt lõi (Luật R8): khách ký hợp đồng "
        "1–2 năm, cước <= 93,67 USD và có dịch vụ bảo mật chiếm 26,7% tập dữ liệu nhưng chỉ có 2,4% rời bỏ."
    )

    add_h2("4.2. Giải thích mô hình LightGBM bằng giá trị Shapley (SHAP)")
    add_h3("4.2.1. Giải thích toàn cục (Global SHAP)")
    add_para(
        "Áp dụng TreeSHAP trên mô hình LightGBM tối ưu với 1.409 hồ sơ tập Test, Bảng 4.2 và Hình 4.2 thể hiện mức đóng góp trung bình "
        "và chiều hướng tác động của 12 đặc trưng quan trọng nhất. Biểu đồ Beeswarm xác nhận nhất quán với cả kiểm định thống kê lẫn "
        "luật CART: đứng đầu là Contract_Month-to-month (mean |SHAP| = 0,6633), tiếp theo là tenure (0,3568), việc không sử dụng "
        "OnlineSecurity (0,2378), không sử dụng TechSupport (0,2006), dùng cáp quang Fiber optic (0,1997) và thanh toán qua "
        "Electronic check (0,1558). Đáng chú ý, đặc trưng tạo mới AvgMonthlySpend cũng lọt vào Top 10 yếu tố có đóng góp lớn nhất."
    )
    shap_rows = [
        [str(i), s["feature"], fmt_dec(s["mean_abs_shap"], 4)]
        for i, s in enumerate(res["shap_top12"][:10], start=1)
    ]
    add_table(
        ["Hạng", "Đặc trưng (sau mã hóa)", "Giá trị mean |SHAP|"],
        shap_rows,
        col_widths=[0.7, 3.8, 2.0],
        caption="Bảng 4.2: Xếp hạng Top 10 đặc trưng có mức đóng góp SHAP trung bình lớn nhất",
    )
    add_figure("fig09_shap_summary.png", "Hình 4.2: Biểu đồ SHAP Summary Plot toàn cục cho Top 12 đặc trưng (LightGBM)")

    add_h3("4.2.2. Giải thích cục bộ cho từng cá nhân (Local SHAP)")
    add_para(
        f"Hình 4.3 đối chiếu cơ chế ra quyết định của LightGBM trên hai hồ sơ đối lập thuộc tập Test. Ở trường hợp rủi ro cao "
        f"(Mã KH {case_hi['customer_id']}, xác suất dự đoán {fmt_pct(case_hi['pred_prob_pct'], 1)}, thực tế đã rời bỏ), thuê bao mới hòa mạng "
        f"{case_hi['tenure']} tháng, ký hợp đồng theo tháng, dùng cáp quang với cước {fmt_dec(case_hi['monthly_charges'], 2)} USD và thanh toán "
        f"qua Electronic check; các giá trị SHAP dương mạnh nhất đến từ tenure (+0,777), Contract_Month-to-month (+0,770) và "
        f"TechSupport_No (+0,287). Ngược lại, ở hồ sơ trung thành (Mã KH {case_lo['customer_id']}, xác suất chỉ {fmt_pct(case_lo['pred_prob_pct'], 2)}), "
        f"thời gian gắn bó {case_lo['tenure']} tháng (−0,967), không dùng hợp đồng theo tháng (−0,696) và hợp đồng 2 năm (−0,303) "
        f"kéo xác suất rời bỏ xuống sát mức 0."
    )
    add_figure("fig10_shap_waterfall_cases.png", "Hình 4.3: Giải thích SHAP cục bộ cho hai hồ sơ điển hình trên tập Test")

    add_h2("4.3. Phân tích lỗi chuyên sâu (Error Analysis)")
    add_para(
        "Để hiểu rõ giới hạn của mô hình và nguyên nhân gốc rễ của các dự đoán sai, nhóm phân tách toàn bộ lỗi trên tập Test "
        "(N = 1.409) của LightGBM tại cả hai ngưỡng T = 0,50 và T* = 0,21 (Bảng 4.3)."
    )
    add_table(
        ["Hạng mục phân tích lỗi", "Tại ngưỡng mặc định T = 0,50", f"Tại ngưỡng tối ưu T* = {t_star_vn}"],
        [
            ["Tổng số ca dự đoán sai / Tổng tập Test", f"{fmt_int(err05['total_errors'])} / {fmt_int(err05['total_test'])} ({fmt_pct(err05['error_rate_pct'], 2)})", f"{fmt_int(errt['total_errors'])} / {fmt_int(errt['total_test'])} ({fmt_pct(errt['error_rate_pct'], 2)})"],
            ["Số ca bỏ sót khách rời bỏ (False Negatives – FN)", fmt_int(err05["fn_count"]), fmt_int(errt["fn_count"])],
            ["  - Trong đó: Khách dài hạn / Hợp đồng 1–2 năm (FN)", f"{err05['fn_loyal_silent']} ({fmt_pct(err05['fn_loyal_silent']/max(err05['fn_count'],1)*100, 1)})", f"{errt['fn_loyal_silent']} ({fmt_pct(errt['fn_loyal_silent']/max(errt['fn_count'],1)*100, 1)})"],
            ["  - Trong đó: Khách ngắn hạn hợp đồng theo tháng (FN)", f"{err05['fn_short_volatile']} ({fmt_pct(err05['fn_short_volatile']/max(err05['fn_count'],1)*100, 1)})", f"{errt['fn_short_volatile']} ({fmt_pct(errt['fn_short_volatile']/max(errt['fn_count'],1)*100, 1)})"],
            ["Số ca cảnh báo dư (False Positives – FP)", fmt_int(err05["fp_count"]), fmt_int(errt["fp_count"])],
            ["  - Trong đó: Mang hồ sơ rủi ro cao (Month-to-month / Fiber)", f"{err05['fp_high_risk_profile']} ({fmt_pct(err05['fp_high_risk_profile']/max(err05['fp_count'],1)*100, 1)})", f"{errt['fp_high_risk_profile']} ({fmt_pct(errt['fp_high_risk_profile']/max(errt['fp_count'],1)*100, 1)})"],
            ["Số lỗi nằm sát vùng biên xác suất (|p − T| <= 0,10)", f"{err05['boundary_errors']} ({fmt_pct(err05['boundary_share_pct'], 1)} tổng lỗi)", f"{errt['boundary_errors']} ({fmt_pct(errt['boundary_share_pct'], 1)} tổng lỗi)"],
        ],
        col_widths=[3.1, 1.7, 1.7],
        caption="Bảng 4.3: Phân tích lỗi dự đoán của LightGBM tại ngưỡng mặc định 0,50 và ngưỡng tối ưu 0,21",
    )
    add_para(
        f"Kết quả giải phẫu lỗi cho thấy ba đặc điểm quan trọng: Thứ nhất, {fmt_pct(errt['boundary_share_pct'], 1)} tổng lỗi nằm trong "
        f"dải vùng biên sát ngưỡng quyết định (|p − T*| <= 0,10), nơi dữ liệu tĩnh khó tách bạch dứt khoát hai lớp. "
        f"Thứ hai, trong {errt['fn_count']} ca bỏ sót còn lại tại T* = {t_star_vn}, có tới {errt['fn_loyal_silent']} ca "
        f"({fmt_pct(errt['fn_loyal_silent']/errt['fn_count']*100, 1)}) là khách hàng lâu năm hoặc đang giữ hợp đồng dài hạn đột ngột rời bỏ; "
        f"nguyên nhân chủ yếu xuất phát từ yếu tố ngoại sinh chưa có trong bộ dữ liệu (chuyển nơi ở, sự cố mạng cục bộ hoặc khuyến mãi "
        f"từ đối thủ). Thứ ba, {errt['fp_high_risk_profile']}/{errt['fp_count']} ca cảnh báo dư (99,1%) thực tế đều thuộc nhóm "
        f"hồ sơ rủi ro cao (hợp đồng theo tháng hoặc cáp quang) — nghĩa là việc chủ động chăm sóc họ mang tính phòng ngừa hợp lý "
        f"chứ không phải dự đoán ngẫu nhiên."
    )

    add_h2("4.4. Đối chứng hiệu quả kinh tế giữa 4 chiến lược vận hành")
    add_para(
        "Để định lượng giá trị thực tiễn của mô hình phân loại trên 1.409 thuê bao tập Test (gồm 374 ca Churn thực tế), "
        "nhóm so sánh trực tiếp tổng chi phí của 4 chiến lược vận hành (Bảng 4.4 và Hình 3.4B)."
    )
    roi_rows = [
        [
            re.sub(r"(\d)\.(\d)", r"\1,\2", r["strategy"]),
            fmt_int(r["targeted_customers"]),
            fmt_int(r["churn_caught"]),
            fmt_int(r["churn_missed"]),
            fmt_usd(r["total_cost_usd"]),
            f"{fmt_usd(r['savings_vs_nothing_usd'])} ({fmt_pct(r['savings_pct'], 1)})",
        ]
        for r in roi
    ]
    add_table(
        ["Chiến lược chăm sóc", "Số KH can thiệp", "Churn giữ được (TP)", "Churn bỏ sót (FN)", "Tổng chi phí", "Tiết kiệm vs Không can thiệp"],
        roi_rows,
        col_widths=[1.85, 0.85, 0.90, 0.85, 0.95, 1.10],
        caption="Bảng 4.4: Đối chứng tổng chi phí và mức tiết kiệm của 4 chiến lược chăm sóc trên tập Test",
    )
    add_para(
        f"So với phương án không can thiệp (tổn thất {fmt_usd(roi[0]['total_cost_usd'])}), chiến lược sử dụng LightGBM tại ngưỡng "
        f"tối ưu T* = {t_star_vn} can thiệp đúng vào {fmt_int(roi[3]['targeted_customers'])} khách hàng có xác suất rủi ro cao nhất, "
        f"giữ chân được {roi[3]['churn_caught']}/374 ca Churn (84,8%) và đưa tổng chi phí xuống chỉ còn {fmt_usd(roi[3]['total_cost_usd'])} "
        f"— tiết kiệm ròng {fmt_usd(roi[3]['savings_vs_nothing_usd'])} (tương đương giảm {fmt_pct(roi[3]['savings_pct'], 1)} tổng tổn thất). "
        f"Ngay cả khi so với luật lọc thủ công trên Excel (Month-to-month & MonthlyCharges > 70 USD, chi phí {fmt_usd(roi[2]['total_cost_usd'])}), "
        f"mô hình khai phá dữ liệu vẫn giúp doanh nghiệp tiết kiệm thêm {fmt_usd(roi[2]['total_cost_usd'] - roi[3]['total_cost_usd'])} "
        f"(giảm thêm 23,4% chi phí) trên cùng quy mô 1.409 thuê bao."
    )

    # ==================== CHUONG 5 ====================
    add_h1("CHƯƠNG 5. TRIỂN KHAI ỨNG DỤNG MINH HỌA VÀ ĐỀ XUẤT GIẢI PHÁP", page_break=True)

    add_h2("5.1. Nguyên mẫu ứng dụng minh họa trên Streamlit")
    add_para(
        "Nhằm hỗ trợ bộ phận chăm sóc khách hàng kiểm chứng trực quan kết quả của quy trình khai phá dữ liệu mà không cần "
        "chạy dòng lệnh, nhóm đóng gói toàn bộ pipeline tiền xử lý và mô hình LightGBM tối ưu (models/best_churn_pipeline.joblib) "
        "vào một ứng dụng web cục bộ viết bằng Streamlit (app.py, vận hành hoàn toàn offline không cần API key). Ứng dụng gồm 3 phân hệ chính:"
    )
    add_bullet("Phân hệ 1 – Dự đoán khách hàng cá nhân: cho phép nhập thông số một thuê bao (hoặc nạp nhanh 3 hồ sơ mẫu), tự động trích xuất 6 đặc trưng miền, đối chiếu xác suất rời bỏ với ngưỡng nhạy chi phí T* = 0,21, đồng thời vẽ biểu đồ thanh đóng góp TreeSHAP cục bộ và đề xuất kịch bản chăm sóc cá nhân hóa.")
    add_bullet("Phân hệ 2 – Chấm điểm hàng loạt qua tệp CSV: cho phép tải lên danh sách nhiều thuê bao (tích hợp sẵn bộ mẫu 100 khách hàng từ tập Test), tự động phân tầng hành động CSKH, ước tính ngân sách can thiệp và xuất tệp CSV danh sách ưu tiên liên hệ theo xác suất rủi ro giảm dần.")
    add_bullet("Phân hệ 3 – Tổng quan kết quả thực nghiệm: hiển thị trực quan bảng đối sánh 5 mô hình, thực nghiệm đối chứng chi phí 4 chiến lược và hệ thống biểu đồ phân tích chính.")

    add_para(
        "Hình 5.1 và Hình 5.2 minh họa giao diện vận hành thực tế của Phân hệ 1 trên một hồ sơ thuê bao rủi ro cao "
        "(hợp đồng theo tháng, cáp quang Fiber optic, mới gắn bó 1 tháng, chưa kích hoạt bảo mật và hỗ trợ kỹ thuật). "
        "Hệ thống tự động tính toán 6 đặc trưng miền tức thì, đưa ra cảnh báo xác suất rời bỏ 87,15% (vượt xa ngưỡng T* = 0,21), "
        "chỉ rõ các nhân tố đẩy rủi ro tăng cao trên biểu đồ TreeSHAP (đứng đầu là hợp đồng theo tháng +0,769 và số tháng gắn bó thấp +0,752) "
        "kèm theo 4 khuyến nghị chăm sóc cụ thể."
    )
    add_figure(
        "fig11_streamlit_form.png",
        "Hình 5.1: Giao diện tổng quan ứng dụng và biểu mẫu nhập thông số thuê bao tại Phân hệ 1 (Streamlit)",
        width_in=6.35,
    )
    add_figure(
        "fig12_streamlit_shap.png",
        "Hình 5.2: Kết quả dự đoán cá nhân, biểu đồ giải thích TreeSHAP cục bộ và khuyến nghị chăm sóc tại Phân hệ 1",
        width_in=6.35,
    )

    add_para(
        "Hình 5.3 trình bày kết quả chấm điểm hàng loạt tại Phân hệ 2 trên danh sách 100 thuê bao mẫu thuộc tập Test với ngưỡng T* = 0,21. "
        "Ứng dụng tự động nhận diện 44 thuê bao cần can thiệp (gồm 21 thuê bao nguy cơ cao >= 50% và 23 thuê bao thuộc vùng cảnh báo sớm 21%–50%), "
        "ước tính ngân sách chăm sóc dự kiến 4.400 USD và cho phép bộ phận CSKH lọc nhanh theo loại hợp đồng hoặc phân tầng rủi ro để tải danh sách gọi điện."
    )
    add_figure(
        "fig13_streamlit_batch.png",
        "Hình 5.3: Giao diện chấm điểm hàng loạt 100 thuê bao từ tệp CSV, phân tầng rủi ro và bảng xếp hạng tại Phân hệ 2",
        width_in=6.35,
    )

    add_h2("5.2. Đề xuất chiến lược chăm sóc khách hàng dựa trên tri thức khai phá")
    add_para(
        "Dựa trên tổng hợp kết quả từ hệ luật IF-THEN (Bảng 4.1), biểu đồ SHAP toàn cục (Hình 4.2) và phân tích lỗi (Bảng 4.3), "
        "nhóm đề xuất bốn nhóm hành động ưu tiên cho doanh nghiệp viễn thông (Bảng 5.1)."
    )
    add_table(
        ["Phân khúc rủi ro phát hiện", "Cơ sở từ kết quả khai phá", "Hành động chăm sóc đề xuất"],
        [
            ["Thuê bao hợp đồng theo tháng + Cáp quang + tenure <= 13 tháng", "Luật R1: chiếm 13,5% tập mẫu nhưng tỷ lệ Churn lên tới 71,0%.", "Ưu tiên gọi CSKH chủ động ở tháng thứ 2 và tháng thứ 6; đề xuất gói chuyển đổi hợp đồng 1 năm kèm ưu đãi giảm 10%–15% cước 3 tháng đầu."],
            ["Thuê bao Cáp quang chưa dùng OnlineSecurity / TechSupport", "OnlineSecurity_No và TechSupport_No xếp hạng 3 và 4 trên SHAP toàn cục.", "Tặng gói dùng thử miễn phí 3 tháng OnlineSecurity và TechSupport nhằm tăng trải nghiệm an toàn và nâng cao chi phí chuyển đổi."],
            ["Thuê bao thanh toán thủ công qua Electronic check", "Tỷ lệ Churn 45,3% (gấp ~3 lần nhóm thanh toán tự động); SHAP hạng 6.", "Khuyến khích liên kết thanh toán tự động (ngân hàng / thẻ tín dụng) thông qua ưu đãi trừ trực tiếp vào hóa đơn kỳ đầu."],
            ["Nhóm xác suất trung bình quanh ngưỡng T* (0,21 – 0,40)", "Chiếm 34,2% vùng lỗi biên nhưng chi phí chăm sóc thấp hơn nhiều tổn thất mất khách.", "Áp dụng kênh chăm sóc chi phí thấp (SMS/Email khảo sát hài lòng, ưu đãi băng thông) thay vì gọi điện trực tiếp để tối ưu ngân sách."],
        ],
        col_widths=[1.85, 2.05, 2.60],
        caption="Bảng 5.1: Khuyến nghị hành động chăm sóc khách hàng theo nhóm rủi ro khai phá",
    )

    # ==================== KET LUAN ====================
    add_h1("KẾT LUẬN VÀ HƯỚNG PHÁT TRIỂN", page_break=True)
    add_h2("1. Kết luận")
    add_para(
        "Bài tập lớn đã hoàn thành đầy đủ chu trình Khai phá dữ liệu cho bài toán phân loại nguy cơ rời bỏ dịch vụ của khách hàng "
        "(Chủ đề 2) trên bộ dữ liệu thực IBM Telco Customer Churn (7.043 bản ghi). Các kết quả chính đạt được gồm:"
    )
    add_bullet("Kiểm định thống kê xác định rõ các nhân tố chi phối hành vi rời bỏ (Contract, tenure, OnlineSecurity, TechSupport, InternetService, PaymentMethod) và phát hiện hiện tượng đa cộng tuyến giữa tenure và TotalCharges (r = 0,83).")
    add_bullet("Xây dựng pipeline tiền xử lý chống rò rỉ dữ liệu kết hợp 6 đặc trưng miền (tổng 25 đặc trưng đầu vào, 55 chiều sau mã hóa); đối sánh 5 mô hình phân loại trên Stratified 5-Fold CV cho thấy LightGBM (CV ROC-AUC = 0,8494, Test PR-AUC = 0,6657) và XGBoost (Test ROC-AUC = 0,8470) đạt hiệu năng cao nhất.")
    add_bullet("Chứng minh qua thực nghiệm rằng tối ưu ngưỡng quyết định nhạy chi phí trên mô hình xác suất chuẩn (T* = 0,21) vượt trội hơn cả ngưỡng mặc định lẫn kỹ thuật gán trọng số lớp balanced: giữ nguyên Brier Score (0,1348), nâng Recall lên 84,76% và giảm 49,6% tổng chi phí so với không can thiệp.")
    add_bullet("Chuyển hóa mô hình thành tri thức giải thích được thông qua 8 luật IF-THEN tường minh từ cây CART, phân tích SHAP toàn cục/cục bộ và ứng dụng web minh họa Streamlit gồm 3 phân hệ chạy hoàn toàn cục bộ.")

    add_h2("2. Hạn chế và hướng phát triển")
    add_para(
        "Do bộ dữ liệu IBM Telco Customer Churn chỉ lưu trạng thái tĩnh tại một thời điểm (snapshot), mô hình chưa phản ánh được "
        "chuỗi biến động hành vi theo từng tháng (như số lần gọi tổng đài khiếu nại, suy giảm lưu lượng sử dụng hoặc số lần chậm thanh toán). "
        "Đây cũng là nguyên nhân chính dẫn đến 41 ca bỏ sót ở nhóm khách hàng dài hạn rời bỏ đột xuất. Trong các nghiên cứu tiếp theo, "
        "đề tài có thể mở rộng theo hai hướng: (1) tích hợp dữ liệu chuỗi thời gian kết hợp mô hình phân tích sống còn (Survival Analysis) "
        "để dự báo không chỉ khả năng rời bỏ mà cả thời điểm dự kiến rời bỏ; và (2) kết hợp phân cụm khách hàng (K-Means / RFM) "
        "trước khi phân loại nhằm cá nhân hóa sâu hơn chính sách ưu đãi theo từng nhóm giá trị."
    )

    # ==================== TAI LIEU THAM KHAO ====================
    add_h1("TÀI LIỆU THAM KHẢO", page_break=True)
    refs = [
        "[1] U. Fayyad, G. Piatetsky-Shapiro, and P. Smyth, \"From data mining to knowledge discovery in databases,\" AI Magazine, vol. 17, no. 3, pp. 37-54, 1996.",
        "[2] J. Han, M. Kamber, and J. Pei, Data Mining: Concepts and Techniques, 3rd ed. Waltham, MA: Morgan Kaufmann, 2011.",
        "[3] L. Breiman, \"Random forests,\" Machine Learning, vol. 45, no. 1, pp. 5-32, 2001.",
        "[4] T. Chen and C. Guestrin, \"XGBoost: A scalable tree boosting system,\" in Proc. 22nd ACM SIGKDD Int. Conf. Knowledge Discovery and Data Mining (KDD), 2016, pp. 785-794.",
        "[5] G. Ke et al., \"LightGBM: A highly efficient gradient boosting decision tree,\" in Advances in Neural Information Processing Systems (NeurIPS), vol. 30, 2017, pp. 3146-3154.",
        "[6] C. Elkan, \"The foundations of cost-sensitive learning,\" in Proc. 17th Int. Joint Conf. Artificial Intelligence (IJCAI), 2001, pp. 973-978.",
        "[7] S. M. Lundberg and S.-I. Lee, \"A unified approach to interpreting model predictions,\" in Advances in Neural Information Processing Systems (NeurIPS), vol. 30, 2017, pp. 4765-4774.",
        "[8] W. Verbeke, K. Dejaeger, D. Martens, J. Hur, and B. Baesens, \"New insights into churn prediction in the telecommunication sector: A profit driven data mining approach,\" European Journal of Operational Research, vol. 218, no. 1, pp. 211-229, 2012.",
        "[9] T. Verbraken, W. Verbeke, and B. Baesens, \"A novel profit maximizing metric for measuring classification performance of customer churn prediction models,\" IEEE Trans. Knowledge and Data Engineering, vol. 25, no. 5, pp. 961-973, 2013.",
        "[10] F. Pedregosa et al., \"Scikit-learn: Machine learning in Python,\" Journal of Machine Learning Research, vol. 12, pp. 2825-2830, 2011.",
        "[11] IBM Sample Data Sets, \"Telco Customer Churn Dataset,\" GitHub Repository: IBM/telco-customer-churn-on-icp4d. [Online]. Available: https://github.com/IBM/telco-customer-churn-on-icp4d.",
    ]
    for ref in refs:
        p = doc.add_paragraph()
        p.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
        p.paragraph_format.space_before = Pt(0)
        p.paragraph_format.space_after = Pt(5)
        p.paragraph_format.line_spacing = 1.2
        p.paragraph_format.left_indent = Inches(0.32)
        p.paragraph_format.first_line_indent = Inches(-0.32)
        r = p.add_run(ref)
        r.font.name = "Times New Roman"
        r.font.size = Pt(12)
        r.font.color.rgb = BODY_RGB

    # ==================== PHU LUC ====================
    add_h1("PHỤ LỤC", page_break=True)
    add_h2("Phụ lục A: Siêu tham số tối ưu của 5 mô hình sau Stratified 5-Fold GridSearchCV")
    hp_rows = [
        [
            name,
            ", ".join(f"{k} = {v}" for k, v in res["best_params"][name].items()),
            fmt_dec(res["cv_roc_auc"][name], 4),
            f"{fmt_dec(res['train_times_sec'][name], 2)}s",
        ]
        for name in m_order
    ]
    add_table(
        ["Mô hình", "Bộ siêu tham số tối ưu (GridSearchCV)", "CV ROC-AUC", "Thời gian"],
        hp_rows,
        col_widths=[1.45, 3.25, 0.95, 0.85],
        caption="Bảng A.1: Bộ siêu tham số tối ưu của 5 mô hình sau Stratified 5-Fold GridSearchCV",
    )

    add_h2("Phụ lục B: Hướng dẫn cài đặt và chạy mã nguồn thực nghiệm")
    add_para("Toàn bộ mã nguồn chạy cục bộ bằng Python 3.10+ và không yêu cầu bất kỳ API key nào:")
    add_bullet("Bước 1 – Cài đặt thư viện: python -m pip install -r requirements.txt")
    add_bullet("Bước 2 – Chạy kiểm định thống kê và xuất 3 biểu đồ EDA: python src/eda_stats.py")
    add_bullet("Bước 3 – Huấn luyện 5 mô hình, tối ưu chi phí, SHAP và xuất 7 biểu đồ: python src/train_pipeline.py")
    add_bullet("Bước 4 – Sinh sơ đồ khối quy trình KDD: python src/generate_flowchart.py")
    add_bullet("Bước 5 – Khởi chạy ứng dụng minh họa Streamlit: chạy tệp run_demo.bat hoặc lệnh python -m streamlit run app.py")

    os.makedirs(REPORT_DIR, exist_ok=True)
    doc.save(OUTPUT_DOCX)
    print(f"Word report generated successfully: {OUTPUT_DOCX}")


if __name__ == "__main__":
    build_report()
