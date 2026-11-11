#!/usr/bin/env python3
"""의결 안건 엑셀의 「00.안건 개요」 내용으로 위원회 의결 안건 문서(docx)를 만든다.

서식은 ktnet-policy-template.js 의 KTNET 문서 양식(맑은 고딕 · 차콜 + 앰버)을 따른다.
"""
import re
from collections import Counter
from pathlib import Path

from docx import Document
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_BREAK, WD_TAB_ALIGNMENT
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Pt, RGBColor, Twips

import source_data as src

OUT_DIR = Path(__file__).resolve().parent.parent
FONT = "맑은 고딕"
COMMITTEE = "AI 거버넌스 위원회"

CHARCOAL, DARK, MID, WHITE = "2D2D2D", "404040", "6B6B6B", "FFFFFF"
AMBER, AMBER_DARK, AMBER_LIGHT = "E8A020", "C47E00", "FDF3DC"
RULE_LIGHT, TABLE_BORDER, STRIPE = "DDDDDD", "CCCCCC", "F7F7F7"
CONTROL_FILL = {"금지": "F2F2F2", "조건부": "FFF2CC", "제한": "E2F0D9", "미해당": "FBE5E5"}

PAGE_W, PAGE_H = 11906, 16838
MARGIN_X, MARGIN_TOP, MARGIN_BOTTOM = 1417, 1500, 1300
CONTENT_W = PAGE_W - 2 * MARGIN_X

# OOXML은 속성 요소의 순서를 정해 두고 있어, 순서가 틀리면 Word가 문서를 손상된 것으로 본다.
PPR_ORDER = ["pStyle", "keepNext", "keepLines", "pageBreakBefore", "framePr", "widowControl", "numPr",
             "suppressLineNumbers", "pBdr", "shd", "tabs", "suppressAutoHyphens", "kinsoku", "wordWrap",
             "overflowPunct", "topLinePunct", "autoSpaceDE", "autoSpaceDN", "bidi", "adjustRightInd",
             "snapToGrid", "spacing", "ind", "contextualSpacing", "mirrorIndents", "suppressOverlap", "jc",
             "textDirection", "textAlignment", "textboxTightWrap", "outlineLvl", "divId", "cnfStyle", "rPr",
             "sectPr", "pPrChange"]
RPR_ORDER = ["rStyle", "rFonts", "b", "bCs", "i", "iCs", "caps", "smallCaps", "strike", "dstrike", "outline",
             "shadow", "emboss", "imprint", "noProof", "snapToGrid", "vanish", "webHidden", "color", "spacing",
             "w", "kern", "position", "sz", "szCs", "highlight", "u", "effect", "bdr", "shd", "fitText",
             "vertAlign", "rtl", "cs", "em", "lang", "eastAsianLayout", "specVanish", "oMath"]
TCPR_ORDER = ["cnfStyle", "tcW", "gridSpan", "hMerge", "vMerge", "tcBorders", "shd", "noWrap", "tcMar",
              "textDirection", "tcFitText", "vAlign", "hideMark"]
TBLPR_ORDER = ["tblStyle", "tblpPr", "tblOverlap", "bidiVisual", "tblStyleRowBandSize",
               "tblStyleColBandSize", "tblW", "jc", "tblCellSpacing", "tblInd", "tblBorders", "shd",
               "tblLayout", "tblCellMar", "tblLook"]


def el(tag, **attrs):
    node = OxmlElement(f"w:{tag}")
    for key, value in attrs.items():
        node.set(qn(f"w:{key}"), str(value))
    return node


def set_child(parent, child, order):
    """같은 태그의 기존 요소를 대체하고, 스키마 순서에 맞는 위치에 넣는다."""
    tag = child.tag.split("}")[1]
    existing = parent.find(qn(f"w:{tag}"))
    if existing is not None:
        parent.remove(existing)
    later = {qn(f"w:{name}") for name in order[order.index(tag) + 1:]}
    for sibling in parent:
        if sibling.tag in later:
            sibling.addprevious(child)
            return child
    parent.append(child)
    return child


def border(tag, size, color, style="single", space=0):
    return el(tag, val=style, sz=size, space=space, color=color)


def style_run(run, size=10, bold=False, color=DARK, fill=None):
    run.font.name = FONT
    run.font.size = Pt(size)
    run.font.bold = bold
    run.font.color.rgb = RGBColor.from_string(color)
    rpr = run._element.get_or_add_rPr()
    rpr.find(qn("w:rFonts")).set(qn("w:eastAsia"), FONT)
    if fill:
        set_child(rpr, el("shd", val="clear", color="auto", fill=fill), RPR_ORDER)
    return run


def para(container, text="", *, size=10, bold=False, color=DARK, before=4, after=4, align=None,
         indent=None, hanging=None, line=1.25, keep_next=False):
    p = container.add_paragraph()
    fmt = p.paragraph_format
    fmt.space_before, fmt.space_after, fmt.line_spacing = Pt(before), Pt(after), line
    fmt.keep_with_next = keep_next
    if align:
        fmt.alignment = align
    if indent is not None:
        fmt.left_indent = Twips(indent)
    if hanging is not None:
        fmt.first_line_indent = Twips(-hanging)
    if text:
        style_run(p.add_run(text), size, bold, color)
    return p


def para_borders(p, **sides):
    pbdr = el("pBdr")
    for side in ("top", "left", "bottom", "right"):
        if side in sides:
            pbdr.append(sides[side](side))
    set_child(p._p.get_or_add_pPr(), pbdr, PPR_ORDER)


def rule(container, color, size, before=0, after=0):
    p = para(container, before=before, after=after, line=1.0, size=2)
    para_borders(p, bottom=lambda s: border(s, size, color, space=1))
    return p


def heading(doc, number, text):
    p = para(doc, before=16, after=6, indent=140, keep_next=True)
    style_run(p.add_run(f"{number}  "), 12, True, AMBER)
    style_run(p.add_run(text), 12, True, CHARCOAL)
    para_borders(p, left=lambda s: border(s, 20, AMBER, style="thick", space=8))
    return p


def subheading(doc, label, title):
    p = para(doc, before=8, after=2, keep_next=True)
    style_run(p.add_run(label), 10, True, AMBER_DARK)
    style_run(p.add_run(f"  {title}"), 10, True, CHARCOAL)
    return p


def bullet(doc, text, size=10):
    p = para(doc, before=1, after=1, indent=440, hanging=220)
    style_run(p.add_run("–  "), size, False, MID)
    style_run(p.add_run(text), size)
    return p


def style_cell(cell, text, width, *, fill=None, bold=False, color=DARK, size=9, center=False, lines=None):
    tcpr = cell._tc.get_or_add_tcPr()
    set_child(tcpr, el("tcW", w=width, type="dxa"), TCPR_ORDER)
    borders = el("tcBorders")
    for side in ("top", "left", "bottom", "right"):
        borders.append(border(side, 4, TABLE_BORDER))
    set_child(tcpr, borders, TCPR_ORDER)
    if fill:
        set_child(tcpr, el("shd", val="clear", color="auto", fill=fill), TCPR_ORDER)
    margins = el("tcMar")
    pad_y, pad_x = (45, 90) if size < 9 else (70, 110)
    for side, w in (("top", pad_y), ("left", pad_x), ("bottom", pad_y), ("right", pad_x)):
        margins.append(el(side, w=w, type="dxa"))
    set_child(tcpr, margins, TCPR_ORDER)
    set_child(tcpr, el("vAlign", val="center"), TCPR_ORDER)

    for i, line_text in enumerate(lines if lines is not None else [text]):
        p = cell.paragraphs[0] if i == 0 else cell.add_paragraph()
        p.paragraph_format.space_before = p.paragraph_format.space_after = Pt(0)
        p.paragraph_format.line_spacing = 1.2
        if center:
            p.paragraph_format.alignment = WD_ALIGN_PARAGRAPH.CENTER
        style_run(p.add_run(line_text or ""), size, bold, color)


def table(doc, col_widths, rows, *, header=False, size=9):
    """rows: 셀 목록의 목록. 셀은 문자열 또는 style_cell 옵션 dict(text, span 포함)."""
    tbl = doc.add_table(rows=len(rows), cols=len(col_widths))
    tbl.alignment = WD_TABLE_ALIGNMENT.CENTER
    tbl.autofit = False
    tblpr = tbl._tbl.tblPr
    set_child(tblpr, el("tblW", w=sum(col_widths), type="dxa"), TBLPR_ORDER)
    set_child(tblpr, el("tblLayout", type="fixed"), TBLPR_ORDER)
    for grid_col, width in zip(tbl._tbl.tblGrid.findall(qn("w:gridCol")), col_widths):
        grid_col.set(qn("w:w"), str(width))

    for r, row in enumerate(rows):
        trpr = tbl.rows[r]._tr.get_or_add_trPr()
        trpr.append(el("cantSplit"))
        if header and r == 0:
            trpr.append(el("tblHeader"))
        col = 0
        for spec in row:
            spec = dict(spec) if isinstance(spec, dict) else {"text": spec}
            span = spec.pop("span", 1)
            cell = tbl.cell(r, col)
            if span > 1:
                cell = cell.merge(tbl.cell(r, col + span - 1))
            spec.setdefault("size", size)
            if header and r == 0:
                spec = {"fill": CHARCOAL, "bold": True, "color": WHITE, "center": True, **spec}
            style_cell(cell, spec.pop("text", ""), sum(col_widths[col:col + span]), **spec)
            col += span
    return tbl


def label_cell(text):
    return {"text": text, "fill": CHARCOAL, "bold": True, "color": WHITE, "center": True}


def add_field(paragraph, instruction, size, color):
    field = el("fldSimple", instr=f" {instruction} ")
    paragraph._p.append(field)
    run = paragraph.add_run("1")
    field.append(run._element)
    style_run(run, size, False, color)


def header_footer(section, right_text):
    # 기본 Header/Footer 스타일의 가운데·오른쪽 탭이 남아 있으면 우측 정렬 위치가 어긋난다.
    header_p = section.header.paragraphs[0]
    header_p.style = "Normal"
    header_p.paragraph_format.space_after = Pt(5)
    header_p.paragraph_format.tab_stops.add_tab_stop(Twips(CONTENT_W), WD_TAB_ALIGNMENT.RIGHT)
    style_run(header_p.add_run("KTNET 한국무역정보통신"), 9, True, CHARCOAL)
    style_run(header_p.add_run(f"\t{right_text}"), 8.5, False, MID)
    para_borders(header_p, bottom=lambda s: border(s, 6, AMBER, space=2))

    footer_p = section.footer.paragraphs[0]
    footer_p.style = "Normal"
    footer_p.paragraph_format.space_before = Pt(5)
    footer_p.paragraph_format.tab_stops.add_tab_stop(Twips(CONTENT_W), WD_TAB_ALIGNMENT.RIGHT)
    style_run(footer_p.add_run("본 문서는 KTNET 내부 문서입니다. 허가 없는 외부 배포를 금합니다."), 7.5, False, MID)
    style_run(footer_p.add_run("\tPage "), 7.5, False, MID)
    add_field(footer_p, "PAGE", 7.5, AMBER)
    style_run(footer_p.add_run(" / "), 7.5, False, MID)
    add_field(footer_p, "NUMPAGES", 7.5, MID)
    para_borders(footer_p, top=lambda s: border(s, 4, RULE_LIGHT, space=2))


def setup_document():
    doc = Document()
    normal = doc.styles["Normal"]
    normal.font.name = FONT
    normal.font.size = Pt(10)
    rfonts = normal.element.get_or_add_rPr().find(qn("w:rFonts"))
    rfonts.set(qn("w:eastAsia"), FONT)
    section = doc.sections[0]
    section.page_width, section.page_height = Twips(PAGE_W), Twips(PAGE_H)
    section.left_margin = section.right_margin = Twips(MARGIN_X)
    section.top_margin, section.bottom_margin = Twips(MARGIN_TOP), Twips(MARGIN_BOTTOM)
    section.header_distance = section.footer_distance = Twips(700)
    return doc, section


def title_block(doc, title, subtitle, date, department):
    para(doc, department, size=9.5, color=MID, before=0, after=3)
    badge = para(doc, before=0, after=8)
    style_run(badge.add_run(" 의결 안건 "), 8, True, WHITE, fill=AMBER)
    para(doc, title, size=22, bold=True, color=CHARCOAL, before=0, after=2, line=1.1)
    para(doc, subtitle, size=11, color=MID, before=0, after=8)
    rule(doc, AMBER, 20, after=8)

    lw, vw = 1300, (CONTENT_W - 2600) // 2
    table(doc, [lw, vw, lw, CONTENT_W - 2 * lw - vw], [
        [label_cell("안건 번호"), "", label_cell("안건 구분"), "의결"],
        [label_cell("상 정 일"), date, label_cell("의결 기구"), COMMITTEE],
        [label_cell("소관 부서"), {"text": department, "span": 3}],
        [label_cell("안 건 명"), {"text": src.OVERVIEW["안건명"], "span": 3, "bold": True, "color": CHARCOAL}],
    ])


def request_section(doc, requests):
    heading(doc, "1.", "의결 요청 사항")
    para(doc, "「AI 도구 승인 목록」을 [별표]와 같이 제정하고자 다음 사항의 의결을 요청한다.")
    for item in requests:
        p = para(doc, before=1, after=1, indent=440, hanging=260)
        style_run(p.add_run(item[0] + "  "), 10, True, AMBER)
        style_run(p.add_run(item[1:].strip()), 10, True, CHARCOAL)
    control = Counter(t[7] for t in src.TOOLS)
    approval = Counter(t[6] for t in src.TOOLS)
    note = para(doc, before=6, after=2, indent=160)
    style_run(note.add_run("의결 대상  "), 9.5, True, AMBER_DARK)
    style_run(note.add_run(
        f"{len(src.TOOLS)}개 도구 — 승인 가능 {approval['가능']} · 불가 {approval['불가']} / 개인정보 입력 "
        + " · ".join(f"{label} {control[label]}" for label in CONTROL_FILL)
        + ". 도구별 값은 [별표] 참조."), 9.5)
    note.paragraph_format.right_indent = Twips(160)
    set_child(note._p.get_or_add_pPr(), el("shd", val="clear", color="auto", fill=AMBER_LIGHT), PPR_ORDER)
    para_borders(note,
                 top=lambda s: border(s, 4, AMBER, space=4), bottom=lambda s: border(s, 4, AMBER, space=4),
                 left=lambda s: border(s, 24, AMBER, style="thick", space=6),
                 right=lambda s: border(s, 4, AMBER, space=4))


def core_section(doc):
    heading(doc, "2.", "핵심 사항")
    for label, text in src.CORE_POINTS:
        key, title = [part.strip() for part in label.split("—", 1)]
        subheading(doc, key, title)
        for line in text.split("\n"):
            bullet(doc, re.sub(r"\s{2,}", "   ", line.strip()))


def enforcement_section(doc):
    heading(doc, "3.", "시행 및 사후 관리")
    lw = 1500
    label = lambda text: {"text": text, "fill": STRIPE, "bold": True, "color": CHARCOAL, "center": True}
    table(doc, [lw, CONTENT_W - lw], [
        [label("시행 전제"), src.OVERVIEW["시행 전제"]],
        [label("관리 주기"), src.OVERVIEW["관리 주기"]],
        [label("사후 관리"),
         "의결된 승인 목록은 「AI 도구 승인 목록 관리대장」(엑셀)으로 관리하며, "
         "의결 결과와 이후의 변경 사항을 대장에 기록한다."],
    ], size=9.5)


def attachment_section(doc):
    heading(doc, "4.", "붙임")
    items = [
        f"[별표] 의결 대상 도구 목록 ({len(src.TOOLS)}개 도구)  1부",
        f"AI 도구 승인 목록 관리대장 ({src.LEDGER_FILENAME})  1부 — 범례, 도구별 상세(원문), "
        "현황 집계, 조건 해소 과제, 근거 자료 수록",
    ]
    for i, item in enumerate(items, start=1):
        p = para(doc, before=1, after=1, indent=440, hanging=260)
        style_run(p.add_run(f"{i}.  "), 10, True, AMBER)
        style_run(p.add_run(item), 10)


def result_section(doc):
    heading(doc, "", "의결 결과")
    lw = 1700
    half = (CONTENT_W - 2 * lw) // 2
    label = lambda text: {"text": text, "fill": STRIPE, "bold": True, "color": CHARCOAL, "center": True}
    table(doc, [lw, half, lw, CONTENT_W - 2 * lw - half], [
        [label("의결 결과"), {"text": "□ 원안 의결      □ 수정 의결      □ 보류      □ 부결", "span": 3}],
        [label("의 결 일"), "20      .       .       .", label("위 원 장"),
         {"text": "(서명)", "color": MID, "center": True}],
        [label("수정 · 부대 의견"), {"text": "", "span": 3, "lines": ["", "", ""]}],
    ], size=9.5)


def appendix(doc):
    p = para(doc, before=0, after=0)
    p.add_run().add_break(WD_BREAK.PAGE)
    title = para(doc, before=0, after=2)
    style_run(title.add_run("[별표]  "), 13, True, AMBER)
    style_run(title.add_run("의결 대상 도구 목록"), 13, True, CHARCOAL)
    para(doc, f"상정일({src.OVERVIEW['상정일 / 소관'].split('/')[0].strip()}) 기준 의결 요청 값. "
              "선결과제(C1~C7)와 도구별 조건·전제 원문은 관리대장 01·03 시트 참조.",
         size=9, color=MID, before=0, after=6)

    widths = [600, 1050, 1600, 1250, 1600, 650, 650, 950]
    widths.append(CONTENT_W - sum(widths))
    rows = [["연번", "분류", "도구", "제공사", "플랜", "등급", "승인", "개인정보 입력", "재검토"]]
    for no, category, name, vendor, plan, grade, approval, control, _, cycle, _ in src.TOOLS:
        fill = CONTROL_FILL[control]
        cells = [(str(no), True, False), (category, False, False), (name, False, False), (vendor, False, False),
                 (plan, False, False), (grade, True, False), (approval, True, True), (control, True, True),
                 (cycle, True, False)]
        rows.append([{"text": text, "center": center, "bold": bold, "fill": fill}
                     for text, center, bold in cells])
    table(doc, widths, rows, header=True, size=8.5)

    sub = para(doc, before=0, after=4, keep_next=True)
    sub.paragraph_format.page_break_before = True
    style_run(sub.add_run("범례"), 10, True, CHARCOAL)
    legend_rows = [["열", "값", "뜻"]]
    for column, value, meaning, _ in src.LEGEND_VALUES:
        if column not in ("등급", "승인", "개인정보 입력", "재검토"):
            continue
        fill = CONTROL_FILL.get(value) if column == "개인정보 입력" else None
        legend_rows.append([
            {"text": column, "center": True},
            {"text": value, "center": True, "bold": True, "fill": fill},
            meaning,
        ])
    table(doc, [1500, 1300, CONTENT_W - 2800], legend_rows, header=True, size=8.5)


def main():
    src.validate()
    ov = src.OVERVIEW
    title, subtitle = [part.strip() for part in ov["안건명"].split("—", 1)]
    date, department = [part.strip() for part in ov["상정일 / 소관"].split("/", 1)]
    requests = [part.strip() for part in re.split(r"\s{2,}", ov["의결 요청"]) if part.strip()]

    doc, section = setup_document()
    header_footer(section, f"{COMMITTEE} 의결 안건")
    title_block(doc, title, subtitle, date, department)
    request_section(doc, requests)
    core_section(doc)
    enforcement_section(doc)
    attachment_section(doc)
    result_section(doc)
    appendix(doc)
    doc.core_properties.title = title
    doc.core_properties.subject = subtitle
    doc.core_properties.author = department

    out = OUT_DIR / src.AGENDA_FILENAME
    doc.save(out)
    print(f"생성: {out}")


if __name__ == "__main__":
    main()
