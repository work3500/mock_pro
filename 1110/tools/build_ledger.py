#!/usr/bin/env python3
"""의결 안건 엑셀을 「AI 도구 승인 목록 관리대장」으로 전환한다.

주의: 관리대장은 생성 이후 직접 수정하며 관리하는 문서다. 이 스크립트를 다시 실행하면
대장에 손으로 기록한 내용(의결일, 과제 상태, 변경 이력 등)이 초기 상태로 덮어써진다.
기존 파일이 있으면 --force 없이는 실행되지 않는다.
"""
import argparse
import datetime as dt
import sys
import zipfile
from pathlib import Path

from lxml import etree
from openpyxl import Workbook
from openpyxl.formatting.rule import CellIsRule, FormulaRule
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.worksheet.properties import PageSetupProperties

import source_data as src

OUT_DIR = Path(__file__).resolve().parent.parent
MAX_ROW = 200  # 드롭다운·조건부 서식·집계 수식이 미리 걸려 있는 마지막 행
FONT_NAME = "맑은 고딕"

GRAY, YELLOW, GREEN, RED = "F2F2F2", "FFF2CC", "E2F0D9", "FBE5E5"
HEADER_BLUE, SUBHEADER_BLUE = "4472C4", "D9E2F3"
MARK_FILL, MARK_FONT = "FFD966", "7F6000"
CONTROL_FILL = {"금지": GRAY, "조건부": YELLOW, "제한": GREEN, "미해당": RED}

THIN = Side(style="thin", color="BFBFBF")
BORDER = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)

S01, S03, S05 = "01.승인 목록", "03.도구별 상세", "05.조건 해소 과제"
SHEET_GUIDE = (
    "01 승인 목록(한눈 표) → 02 범례 → 03 도구별 상세(원문) → 04 현황 집계 "
    "→ 05 조건 해소 과제 → 06 근거 자료 → 07 변경 이력"
)

# openpyxl은 수식의 계산값을 저장하지 않는다. 제한된 보기처럼 재계산이 꺼진 환경에서도
# 값이 보이도록 (시트명, 셀) → 계산값을 모아 두었다가 저장 후 XML에 써넣는다.
CACHED = {}


def fill(color):
    return PatternFill(start_color=color, end_color=color, fill_type="solid")


def font(bold=False, color=None):
    return Font(name=FONT_NAME, size=9, bold=bold, color=color)


def put(ws, row, col, value, *, bold=False, h=None, v="top", wrap=True, color=None, fmt=None):
    cell = ws.cell(row=row, column=col, value=value)
    cell.font = font(bold)
    cell.alignment = Alignment(horizontal=h, vertical=v, wrap_text=wrap)
    cell.border = BORDER
    if color:
        cell.fill = fill(color)
    if fmt:
        cell.number_format = fmt
    return cell


def put_formula(ws, row, col, formula, cached, **style):
    cell = put(ws, row, col, formula, **style)
    CACHED[(ws.title, cell.coordinate)] = cached
    return cell


def header(ws, labels, row=1, height=None):
    for col, label in enumerate(labels, start=1):
        cell = ws.cell(row=row, column=col, value=label)
        cell.font = font(bold=True, color="FFFFFF")
        cell.fill = fill(HEADER_BLUE)
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
    if height:
        ws.row_dimensions[row].height = height
    ws.freeze_panes = "A2"


def subheader(ws, row, labels, center_cols=()):
    for col, label in enumerate(labels, start=1):
        put(ws, row, col, label, bold=True, color=SUBHEADER_BLUE,
            h="center" if col in center_cols else None,
            v="center" if col in center_cols else "top")


def widths(ws, values):
    for col, width in enumerate(values, start=1):
        ws.column_dimensions[get_column_letter(col)].width = width


def page_setup(ws):
    ws.page_setup.orientation = "landscape"
    ws.page_setup.paperSize = ws.PAPERSIZE_A4
    ws.sheet_properties.pageSetUpPr = PageSetupProperties(fitToPage=True)
    ws.page_setup.fitToWidth = 1
    ws.page_setup.fitToHeight = 0
    ws.print_title_rows = "1:1"
    ws.page_margins.left = ws.page_margins.right = 0.4
    ws.page_margins.top = ws.page_margins.bottom = 0.5


def dropdown(ws, cells, options, title):
    dv = DataValidation(type="list", formula1='"' + ",".join(options) + '"', allow_blank=True)
    # 범례에 없는 값이 생길 수 있으므로 입력을 막지 않고 경고만 띄운다.
    dv.errorStyle = "warning"
    dv.errorTitle = title
    dv.error = "목록에 없는 값입니다. 02 범례에 정의를 추가한 뒤 사용하세요."
    ws.add_data_validation(dv)
    dv.add(cells)


def date_only(ws, cells):
    dv = DataValidation(type="date", operator="greaterThanOrEqual", formula1="43831", allow_blank=True)
    dv.errorTitle = "날짜 형식"
    dv.error = "날짜로 입력하세요. 예: 2026-11-10"
    ws.add_data_validation(dv)
    dv.add(cells)


def row_border_rule(ws, cells, key_col):
    """내용이 입력된 새 행에도 테두리가 따라붙게 한다."""
    ws.conditional_formatting.add(cells, FormulaRule(formula=[f'${key_col}2<>""'], border=BORDER))


def build_overview(wb):
    ws = wb.active
    ws.title = "00.대장 개요"
    header(ws, ["구분", "내용"])
    widths(ws, [24, 104])
    ov = src.OVERVIEW
    rows = [
        ("대장명", "AI 도구 승인 목록 관리대장"),
        ("소관", ov["상정일 / 소관"].split("/", 1)[1].strip()),
        ("관리 항목", "① 도구별 사용 승인·불가  ② 도구별 등급  ③ 도구별 「개인정보 입력」 통제값"),
        ("의결 근거",
         "「AI 도구 승인 목록 제정(안)」 — 2026-11-10 위원회 상정 "
         f"(의결 문서: {src.AGENDA_FILENAME})\n"
         "의결 결과와 의결일은 01 시트의 「위원회 의결」·「의결일」 열에 기록한다."),
        ("시행 전제", ov["시행 전제"]),
        ("관리 주기", ov["관리 주기"]),
        ("갱신 방법",
         "변경 사항은 01~05 시트의 해당 행에 반영하고, 변경 전·후 값과 근거를 07 변경 이력에 기록한다.\n"
         "신규 도구는 위원회 의결 후 01·03 시트의 마지막 행 바로 아래에 추가한다."),
    ]
    row = 2
    for label, text in rows:
        put(ws, row, 1, label)
        put(ws, row, 2, text)
        row += 1
    row += 1
    for label, text in src.CORE_POINTS:
        put(ws, row, 1, label, bold=True, color=YELLOW)
        put(ws, row, 2, text, color=YELLOW)
        row += 1
    row += 1
    put(ws, row, 1, "시트 안내")
    put(ws, row, 2, SHEET_GUIDE)
    page_setup(ws)


def build_tools(wb):
    ws = wb.create_sheet(S01)
    labels = src.TOOL_HEADERS + ["의결일", "최근\n재검토일", "차기\n재검토일", "비고"]
    header(ws, labels, height=30)
    widths(ws, [4, 10, 20, 15, 19, 5, 5, 8, 4, 4, 4, 4, 4, 4, 4, 7, 26, 9, 11, 11, 11, 24])
    ws.freeze_panes = "D2"

    centered = {1, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 16, 18, 19, 20, 21}
    for i, tool in enumerate(src.TOOLS):
        r = i + 2
        no, category, name, vendor, plan, grade, approval, control, codes, cycle, memo = tool
        marks = ["●" if f"C{n}" in codes.split() else None for n in range(1, 8)]
        values = [no, category, name, vendor, plan, grade, approval, control, *marks, cycle, memo,
                  None, None, None, None, None]
        for col, value in enumerate(values, start=1):
            is_center = col in centered
            put(ws, r, col, value, bold=col in (7, 8), wrap=col in (17, 22),
                h="center" if is_center else None, v="center" if is_center else "top",
                fmt="yyyy-mm-dd" if col in (19, 20, 21) else None)
        put_formula(
            ws, r, 21,
            f'=IF(OR($P{r}="반기",$P{r}="연"),IF(MAX($S{r},$T{r})=0,"",'
            f'EDATE(MAX($S{r},$T{r}),IF($P{r}="반기",6,12))),"")',
            "", h="center", v="center", wrap=False, fmt="yyyy-mm-dd")

    last = len(src.TOOLS) + 1
    ws.auto_filter.ref = f"A1:V{last}"

    cf = ws.conditional_formatting
    cf.add(f"I2:O{MAX_ROW}", CellIsRule(operator="equal", formula=['"●"'], fill=fill(MARK_FILL),
                                        font=Font(bold=True, color=MARK_FONT)))
    cf.add(f"U2:U{MAX_ROW}", FormulaRule(formula=['AND(ISNUMBER($U2),$U2<TODAY())'],
                                         font=Font(bold=True, color="C00000")))
    for value, color in CONTROL_FILL.items():
        cf.add(f"A2:V{MAX_ROW}", FormulaRule(formula=[f'$H2="{value}"'], fill=fill(color)))
    row_border_rule(ws, f"A2:V{MAX_ROW}", "C")

    dropdown(ws, f"F2:F{MAX_ROW}", ["개인", "회사", "－"], "등급")
    dropdown(ws, f"G2:G{MAX_ROW}", ["가능", "불가"], "승인")
    dropdown(ws, f"H2:H{MAX_ROW}", list(CONTROL_FILL), "개인정보 입력")
    dropdown(ws, f"I2:O{MAX_ROW}", ["●"], "선결과제")
    dropdown(ws, f"P2:P{MAX_ROW}", ["반기", "연", "－"], "재검토")
    dropdown(ws, f"R2:R{MAX_ROW}", ["원안의결", "수정의결", "보류", "부결"], "위원회 의결")
    date_only(ws, f"S2:T{MAX_ROW}")
    page_setup(ws)


def build_legend(wb):
    ws = wb.create_sheet("02.범례")
    header(ws, ["열", "값", "뜻", "색"])
    widths(ws, [16, 18, 86, 16])
    ledger_columns = [
        ("위원회 의결", "원안의결 / 수정의결 / 보류 / 부결", "위원회 의결 결과. 의결 전에는 비워 둔다", None),
        ("의결일", "날짜", "위원회 의결일. YYYY-MM-DD 형식으로 입력", None),
        ("최근 재검토일", "날짜", "정기·수시 재검토를 마친 날. 재검토 결과는 07 변경 이력에 기록", None),
        ("차기 재검토일", "자동 계산",
         "의결일과 최근 재검토일 중 늦은 날 + 재검토 주기(반기 6개월 · 연 12개월). 기한이 지나면 붉은 글씨로 표시",
         "붉은 글씨"),
        ("상태 (05 시트)", "미완료 / 진행중 / 완료", "조건 해소 과제의 진행 상태. 완료 시 완료일과 증빙을 함께 기록", None),
    ]
    row = 2
    for values in src.LEGEND_VALUES + ledger_columns:
        for col, value in enumerate(values, start=1):
            put(ws, row, col, value)
        row += 1
    row += 1
    subheader(ws, row, ["코드", "과제", "내용", "소관"])
    row += 1
    for values in src.LEGEND_CODES:
        for col, value in enumerate(values, start=1):
            put(ws, row, col, value)
        row += 1
    row += 1
    for i, values in enumerate(src.LEGEND_COMMON):
        if i == 0:
            subheader(ws, row, values)
        else:
            for col, value in enumerate(values, start=1):
                put(ws, row, col, value)
        row += 1
    page_setup(ws)


def build_details(wb):
    ws = wb.create_sheet(S03)
    header(ws, ["연번", "도구", "플랜", "개인정보 입력", "조건·전제 (원문)", "개별 선결과제 (원문)"])
    widths(ws, [4, 20, 19, 10, 58, 76])
    lookup = f"'{S01}'!$A$2:$H${MAX_ROW}"
    for i, tool in enumerate(src.TOOLS):
        r = i + 2
        no, _, name, _, plan, _, _, control = tool[:8]
        condition, task = src.DETAILS[no]
        put(ws, r, 1, no, h="center", wrap=False)
        # 도구·플랜·통제값은 01 시트가 원본이다. 연번으로 끌어와 두 시트가 어긋나지 않게 한다.
        put_formula(ws, r, 2, f'=IFERROR(VLOOKUP($A{r},{lookup},3,FALSE),"")', name)
        put_formula(ws, r, 3, f'=IFERROR(VLOOKUP($A{r},{lookup},5,FALSE),"")', plan)
        put_formula(ws, r, 4, f'=IFERROR(VLOOKUP($A{r},{lookup},8,FALSE),"")', control, h="center", wrap=False)
        put(ws, r, 5, condition)
        put(ws, r, 6, task)
    ws.auto_filter.ref = f"A1:F{len(src.TOOLS) + 1}"
    for value, color in CONTROL_FILL.items():
        ws.conditional_formatting.add(
            f"D2:D{MAX_ROW}", CellIsRule(operator="equal", formula=[f'"{value}"'], fill=fill(color)))
    row_border_rule(ws, f"A2:F{MAX_ROW}", "A")
    page_setup(ws)


def build_summary(wb):
    ws = wb.create_sheet("04.현황 집계")
    header(ws, ["개인정보 입력 통제값", "건수", "내용"])
    widths(ws, [26, 8, 72])

    def rng(col):
        return f"'{S01}'!${col}$2:${col}${MAX_ROW}"

    def count_cell(row, formula, cached, bold=False):
        put_formula(ws, row, 2, formula, cached, bold=bold, h="center", v="center", wrap=False)

    any_mark = "+".join(f'({rng(c)}="●")' for c in "IJKLMNO")

    row = 2
    for label, count, text in src.SUMMARY_CONTROL:
        put(ws, row, 1, label, bold=True, color=CONTROL_FILL[label])
        count_cell(row, f'=COUNTIF({rng("H")},"{label}")', count)
        put(ws, row, 3, text)
        row += 1
    put(ws, row, 1, "합계", bold=True)
    count_cell(row, "=SUM(B2:B5)", src.SUMMARY_TOTAL, bold=True)
    put(ws, row, 3, None)

    row += 2
    subheader(ws, row, ["그 밖의 집계", "건수", "내용"], center_cols=(2,))
    (ok, ok_n, ok_text), (no, no_n, no_text), (pre, pre_n, pre_text), (free, free_n, free_text) = src.SUMMARY_OTHER
    row += 1
    put(ws, row, 1, ok)
    count_cell(row, f'=COUNTIF({rng("G")},"가능")', ok_n)
    put_formula(
        ws, row, 3,
        f'="개인계정 "&COUNTIFS({rng("G")},"가능",{rng("F")},"개인")'
        f'&" · 회사계약 "&COUNTIFS({rng("G")},"가능",{rng("F")},"회사")',
        ok_text)
    row += 1
    put(ws, row, 1, no)
    count_cell(row, f'=COUNTIF({rng("G")},"불가")', no_n)
    put(ws, row, 3, no_text)
    row += 1
    put(ws, row, 1, pre)
    count_cell(row, f"=SUMPRODUCT(--(({any_mark})>0))", pre_n)
    put(ws, row, 3, pre_text)
    row += 1
    put(ws, row, 1, free)
    count_cell(row, f'=SUMPRODUCT(({rng("H")}="조건부")*(({any_mark})=0))', free_n)
    put(ws, row, 3, free_text)

    row += 2
    subheader(ws, row, ["관리 현황", "건수", "내용"], center_cols=(2,))
    row += 1
    put(ws, row, 1, "재검토 기한 경과")
    count_cell(row, f'=COUNTIF({rng("U")},"<"&TODAY())', 0)
    put(ws, row, 3, "01 시트 「차기 재검토일」이 오늘보다 이전인 도구")
    row += 1
    status = f"'{S05}'!$G$2:$G${MAX_ROW}"
    open_tasks = sum(1 for t in src.TASKS if initial_task_status(t))
    put(ws, row, 1, "조건 해소 과제 — 미완료·진행중")
    count_cell(row, f'=COUNTIF({status},"미완료")+COUNTIF({status},"진행중")', open_tasks)
    put(ws, row, 3, "05 시트 「상태」 기준")

    row += 2
    subheader(ws, row, ["검토 근거 (제정 시점)", "건수", "내용"], center_cols=(2,))
    for label, count, text in src.SUMMARY_EVIDENCE:
        row += 1
        put(ws, row, 1, label)
        put(ws, row, 2, count, h="center", v="center", wrap=False)
        put(ws, row, 3, text)
    page_setup(ws)


def initial_task_status(task):
    """원본이 '남아 있음'으로 명시한 과제(KTNET 이행 3건, C1~C7)만 미완료로 시작한다."""
    return "미완료" if task[0] in ("KTNET 이행(공통)", "도구별") else None


def build_tasks(wb):
    ws = wb.create_sheet(S05)
    header(ws, ["구분", "No", "과제", "해소되는 조문", "대상", "소관", "상태", "완료일", "증빙 · 비고"])
    widths(ws, [16, 5, 56, 34, 40, 16, 9, 11, 30])
    for i, task in enumerate(src.TASKS):
        r = i + 2
        color = YELLOW if task[0] == "KTNET 이행(공통)" else None
        for col, value in enumerate(task, start=1):
            put(ws, r, col, value, color=color,
                h="center" if col == 2 else None, v="center" if col == 2 else "top", wrap=col != 2)
        put(ws, r, 7, initial_task_status(task), h="center", v="center", wrap=False)
        put(ws, r, 8, None, h="center", v="center", wrap=False, fmt="yyyy-mm-dd")
        put(ws, r, 9, None)
    ws.auto_filter.ref = f"A1:I{len(src.TASKS) + 1}"
    for value, color in (("미완료", RED), ("진행중", YELLOW), ("완료", GREEN)):
        ws.conditional_formatting.add(
            f"G2:G{MAX_ROW}", CellIsRule(operator="equal", formula=[f'"{value}"'], fill=fill(color)))
    row_border_rule(ws, f"A2:I{MAX_ROW}", "C")
    dropdown(ws, f"G2:G{MAX_ROW}", ["미완료", "진행중", "완료"], "상태")
    date_only(ws, f"H2:H{MAX_ROW}")
    page_setup(ws)


def build_evidence(wb):
    ws = wb.create_sheet("06.근거 자료")
    header(ws, ["No", "문서", "내용"])
    widths(ws, [5, 42, 96])
    docs = src.EVIDENCE_DOCS + [
        (len(src.EVIDENCE_DOCS) + 1, src.AGENDA_FILENAME,
         "위원회 의결 안건 문서 — 안건 개요, 의결 요청 사항, 의결 대상 도구 목록(별표)"),
    ]
    row = 2
    for values in docs:
        for col, value in enumerate(values, start=1):
            put(ws, row, col, value)
        row += 1
    put(ws, row, 1, None)
    put(ws, row, 2, src.EVIDENCE_LAW[0])
    put(ws, row, 3, src.EVIDENCE_LAW[1])
    page_setup(ws)


def build_history(wb):
    ws = wb.create_sheet("07.변경 이력")
    header(ws, ["No", "일자", "구분", "대상 (연번 · 도구)", "변경 전", "변경 후", "근거 (의결 · 결재)", "작성"])
    widths(ws, [5, 11, 12, 24, 30, 44, 36, 12])
    first = [
        1, dt.date(2026, 11, 10), "제정", f"전체 ({len(src.TOOLS)}개 도구)", "－",
        f"의결 안건 엑셀({src.SOURCE_FILENAME})을 관리대장으로 전환. "
        f"{len(src.TOOLS)}개 도구를 의결 요청 값 그대로 등재(위원회 의결 전)",
        "「AI 도구 승인 목록 제정(안)」 2026-11-10 위원회 상정", "AX추진실",
    ]
    for col, value in enumerate(first, start=1):
        centered = col in (1, 2, 3)
        put(ws, 2, col, value, h="center" if centered else None, wrap=not centered,
            fmt="yyyy-mm-dd" if col == 2 else None)
    row_border_rule(ws, f"A2:H{MAX_ROW}", "B")
    dropdown(ws, f"C2:C{MAX_ROW}",
             ["제정", "신규 등재", "등급 변경", "통제값 변경", "폐지", "재검토", "과제 해소", "기타"], "구분")
    date_only(ws, f"B2:B{MAX_ROW}")
    page_setup(ws)


NS = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"


def inject_cached_values(path, sheet_titles):
    """저장된 xlsx의 수식 셀에 계산값(<v>)을 써넣는다."""
    tmp = path.with_suffix(".tmp")
    remaining = dict(CACHED)
    with zipfile.ZipFile(path) as zin, zipfile.ZipFile(tmp, "w", zipfile.ZIP_DEFLATED) as zout:
        for item in zin.infolist():
            data = zin.read(item.filename)
            name = Path(item.filename)
            if name.parent.as_posix() == "xl/worksheets" and name.stem.startswith("sheet"):
                title = sheet_titles[int(name.stem[len("sheet"):]) - 1]
                root = etree.fromstring(data)
                for cell in root.iter(f"{{{NS}}}c"):
                    if cell.find(f"{{{NS}}}f") is None:
                        continue
                    value = remaining.pop((title, cell.get("r")))
                    v = cell.find(f"{{{NS}}}v")
                    if v is None:
                        v = etree.SubElement(cell, f"{{{NS}}}v")
                    if isinstance(value, str):
                        cell.set("t", "str")
                    else:
                        cell.attrib.pop("t", None)
                    v.text = str(value)
                data = etree.tostring(root, xml_declaration=True, encoding="UTF-8", standalone=True)
            zout.writestr(item, data)
    assert not remaining, f"계산값을 써넣지 못한 수식 셀: {sorted(remaining)}"
    tmp.replace(path)


def main():
    parser = argparse.ArgumentParser(description=__doc__.splitlines()[0])
    parser.add_argument("--force", action="store_true", help="기존 관리대장을 덮어쓴다")
    parser.add_argument("--out", type=Path, default=OUT_DIR / src.LEDGER_FILENAME)
    args = parser.parse_args()
    if args.out.exists() and not args.force:
        sys.exit(f"{args.out} 이(가) 이미 있습니다. 직접 기록한 내용이 사라지므로 덮어쓰려면 --force 를 붙이세요.")

    src.validate()
    wb = Workbook()
    for build in (build_overview, build_tools, build_legend, build_details,
                  build_summary, build_tasks, build_evidence, build_history):
        build(wb)
    wb.properties.title = "AI 도구 승인 목록 관리대장"
    wb.save(args.out)
    inject_cached_values(args.out, wb.sheetnames)
    print(f"생성: {args.out}")


if __name__ == "__main__":
    main()
