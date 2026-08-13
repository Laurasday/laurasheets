#!/usr/bin/env python3
"""2026 SEOULCon 라이프스타일 서울 부스(8/23) 운영 관리 시트 생성 스크립트.

내용 없이 양식(헤더/드롭다운/서식)만 채운 빈 템플릿 8개 탭.

실행: python3 build_seoulcon_booth.py
"""

from __future__ import annotations

from pathlib import Path

from openpyxl import Workbook
from openpyxl.styles import Alignment, Font
from openpyxl.worksheet.worksheet import Worksheet

import build_shop_excel as base

OUTPUT_PATH = Path(__file__).parent / "서울콘_라이프스타일서울_부스운영.xlsx"

EVENT_DATE_FORMULA = 'DATE(2026,8,23)'

# ---------------------------------------------------------------------------
# 1. 개요
# ---------------------------------------------------------------------------

개요_라벨 = [
    "행사명",
    "일시",
    "장소",
    "주최",
    "주관",
    "예상 참석 규모",
    "부스 비용",
    "드레스코드",
    "담당 PM (본사)",
    "서브 담당",
    "대행사 (부스 장치)",
    "행사까지 D-day",
]


def build_개요(wb: Workbook) -> None:
    ws = wb.active
    ws.title = "개요"
    last_col = 2

    base.write_title_block(
        ws, "개요", "2026 SEOULCon 부속 '라이프스타일 서울' 부스 운영 개요", "", last_col,
    )
    header_row = 4
    base.write_header(ws, header_row, ["항목", "내용"])

    row_count = len(개요_라벨)
    base.style_body_rows(ws, header_row, row_count, last_col)

    for i, label in enumerate(개요_라벨):
        row = header_row + 1 + i
        ws.cell(row=row, column=1, value=label)
        ws.cell(row=row, column=1).alignment = Alignment(horizontal="left", vertical="center")
        ws.cell(row=row, column=2).alignment = Alignment(horizontal="left", vertical="center")

    dday_row = header_row + 개요_라벨.index("행사까지 D-day") + 1
    ws.cell(row=dday_row, column=2, value=f"={EVENT_DATE_FORMULA}-TODAY()")

    base.set_widths(ws, [20, 60])
    ws.freeze_panes = f"A{header_row + 1}"


# ---------------------------------------------------------------------------
# 2. 마감일·요청사항 트래커
# ---------------------------------------------------------------------------

def build_마감일트래커(wb: Workbook) -> None:
    ws = wb.create_sheet("마감일·요청사항 트래커")
    headers = ["No", "요청사항", "요청기관", "담당자", "마감일", "상태", "비고"]
    last_col = len(headers)
    header_row = 4
    row_count = 30

    base.write_title_block(
        ws, "마감일·요청사항 트래커", "행사 준비 관련 모든 요청사항·마감일 통합 관리",
        "", last_col,
    )
    base.write_header(ws, header_row, headers)
    base.style_body_rows(ws, header_row, row_count, last_col)
    base.add_no_column(ws, header_row, row_count, anchor_col_letter="B")

    base.add_dropdown(ws, "F", header_row, row_count, ["요청전", "진행중", "완료"])
    base.set_column_formats(ws, header_row, row_count, "E", base.DATE_FORMAT)

    base.set_widths(ws, [6, 32, 16, 12, 12, 10, 24])
    ws.freeze_panes = f"A{header_row + 1}"


# ---------------------------------------------------------------------------
# 3. 디자인 필요 리스트
# ---------------------------------------------------------------------------

def build_디자인리스트(wb: Workbook) -> None:
    ws = wb.create_sheet("디자인 필요 리스트")
    headers = [
        "No", "디자인물 명칭", "사이즈", "들어갈 내용 및 자료",
        "첨부·참고링크", "인쇄여부", "칼선유무", "담당자", "마감일", "진행상태", "비고",
    ]
    last_col = len(headers)
    header_row = 4
    row_count = 20

    base.write_title_block(
        ws, "디자인 필요 리스트", "부스 디자인물 제작 관리",
        "", last_col,
    )
    base.write_header(ws, header_row, headers)
    base.style_body_rows(ws, header_row, row_count, last_col)
    base.add_no_column(ws, header_row, row_count, anchor_col_letter="B")

    base.add_dropdown(ws, "F", header_row, row_count, ["Y", "N", "해당없음"])
    base.add_dropdown(ws, "G", header_row, row_count, ["Y", "N", "해당없음"])
    base.add_dropdown(ws, "J", header_row, row_count, ["요청전", "시안중", "컨펌완료", "인쇄중", "제작완료"])
    base.set_column_formats(ws, header_row, row_count, "I", base.DATE_FORMAT)

    base.set_widths(ws, [6, 22, 12, 30, 24, 10, 10, 12, 12, 12, 20])
    ws.freeze_panes = f"A{header_row + 1}"


# ---------------------------------------------------------------------------
# 4. 럭키드로우 상품 리스트
# ---------------------------------------------------------------------------

def build_럭드상품리스트(wb: Workbook) -> None:
    ws = wb.create_sheet("럭키드로우 상품 리스트")
    headers = ["No", "등수", "상품명·혜택", "수량", "지점배분(강남/명동/명동2)", "상당액", "확정여부", "비고"]
    last_col = len(headers)
    header_row = 4
    row_count = 15

    base.write_title_block(
        ws, "럭키드로우 상품 리스트", "럭키드로우 총 200개 상품 구성",
        "", last_col,
    )
    base.write_header(ws, header_row, headers)
    base.style_body_rows(ws, header_row, row_count, last_col)
    base.add_no_column(ws, header_row, row_count, anchor_col_letter="B")

    base.add_dropdown(ws, "G", header_row, row_count, ["미정", "확정"])
    base.set_column_formats(ws, header_row, row_count, "F", base.MONEY_FORMAT)

    sum_row = header_row + 1 + row_count
    ws.cell(row=sum_row, column=1, value="합계")
    ws.cell(row=sum_row, column=4, value=f"=SUM(D{header_row + 1}:D{header_row + row_count})")
    for col in range(1, last_col + 1):
        cell = ws.cell(row=sum_row, column=col)
        cell.font = Font(name=base.FONT_NAME, bold=True, size=10)
        cell.border = base.THIN_BORDER
        cell.alignment = base.BODY_ALIGN

    base.set_widths(ws, [6, 10, 26, 8, 22, 14, 10, 20])
    ws.freeze_panes = f"A{header_row + 1}"


# ---------------------------------------------------------------------------
# 5. 럭키드로우 넘버링 관리
# ---------------------------------------------------------------------------

def build_럭드넘버링(wb: Workbook) -> None:
    ws = wb.create_sheet("럭키드로우 넘버링 관리")
    last_col = 6

    base.write_title_block(
        ws, "지점별 비율 계산", "넘버링 배정 현황 대비 지점별 목표 비율", "", last_col,
    )
    ratio_header_row = 4
    ratio_headers = ["지점", "목표 비율(%)", "목표 수량", "실제 배정 수", "차이", "비고"]
    for col, text in enumerate(ratio_headers, start=1):
        cell = ws.cell(row=ratio_header_row, column=col, value=text)
        cell.font = base.HEADER_FONT
        cell.fill = base.HEADER_FILL
        cell.alignment = base.HEADER_ALIGN
        cell.border = base.THIN_BORDER

    지점들 = ["강남", "명동", "명동2"]
    numbering_header_row = ratio_header_row + 1 + len(지점들) + 2
    for i, 지점 in enumerate(지점들):
        row = ratio_header_row + 1 + i
        ws.cell(row=row, column=1, value=지점)
        ws.cell(row=row, column=3, value=f"=B{row}%*200")
        ws.cell(
            row=row, column=4,
            value=f'=COUNTIF(넘버링관리!$E${numbering_header_row + 1}:$E${numbering_header_row + 200},A{row})',
        )
        ws.cell(row=row, column=5, value=f"=D{row}-C{row}")
        for col in range(1, last_col + 1):
            cell = ws.cell(row=row, column=col)
            cell.font = base.BODY_FONT
            cell.alignment = base.BODY_ALIGN
            cell.border = base.THIN_BORDER
    base.set_widths(ws, [14, 14, 12, 14, 10, 24])

    headers = ["No", "넘버링No", "등수", "상품명", "지점", "당첨자명", "연락처", "수령여부", "비고"]
    last_col2 = len(headers)
    row_count = 200

    title_row = numbering_header_row - 1
    ws.merge_cells(start_row=title_row, start_column=1, end_row=title_row, end_column=last_col2)
    c = ws.cell(row=title_row, column=1, value="넘버링 관리")
    c.font = Font(name=base.FONT_NAME, bold=True, size=14)
    c.alignment = Alignment(horizontal="left", vertical="center")

    for col, text in enumerate(headers, start=1):
        cell = ws.cell(row=numbering_header_row, column=col, value=text)
        cell.font = base.HEADER_FONT
        cell.fill = base.HEADER_FILL
        cell.alignment = base.HEADER_ALIGN
        cell.border = base.THIN_BORDER

    base.style_body_rows(ws, numbering_header_row, row_count, last_col2)
    for i in range(row_count):
        row = numbering_header_row + 1 + i
        ws.cell(row=row, column=2, value=i + 1)

    base.add_dropdown(ws, "H", numbering_header_row, row_count, ["미수령", "수령완료"])
    ws.title = "넘버링관리"
    ws.freeze_panes = f"A{numbering_header_row + 1}"


# ---------------------------------------------------------------------------
# 6. 바우처
# ---------------------------------------------------------------------------

def build_바우처(wb: Workbook) -> None:
    ws = wb.create_sheet("바우처")
    headers = ["No", "바우처종류", "대상시술", "유효기간", "수량", "상당금액", "대상", "확정여부", "비고"]
    last_col = len(headers)
    header_row = 4
    row_count = 10

    base.write_title_block(
        ws, "바우처", "방문 크리에이터·셀럽·인플루언서 대상 시술 바우처 관리",
        "", last_col,
    )
    base.write_header(ws, header_row, headers)
    base.style_body_rows(ws, header_row, row_count, last_col)
    base.add_no_column(ws, header_row, row_count, anchor_col_letter="B")

    base.add_dropdown(ws, "G", header_row, row_count, ["방문 크리에이터", "셀럽", "인플루언서", "기타"])
    base.add_dropdown(ws, "H", header_row, row_count, ["미정", "확정"])
    base.set_column_formats(ws, header_row, row_count, "F", base.MONEY_FORMAT)

    base.set_widths(ws, [6, 18, 16, 12, 8, 14, 16, 10, 20])
    ws.freeze_panes = f"A{header_row + 1}"


# ---------------------------------------------------------------------------
# 7. 담당자 연락처 정리
# ---------------------------------------------------------------------------

def build_연락처(wb: Workbook) -> None:
    ws = wb.create_sheet("담당자 연락처 정리")
    headers = ["No", "소속", "이름·직함", "역할", "연락처", "이메일", "비고"]
    last_col = len(headers)
    header_row = 4
    row_count = 20

    base.write_title_block(
        ws, "담당자 연락처 정리", "내부·대행사·주최측 담당자 연락처 관리",
        "", last_col,
    )
    base.write_header(ws, header_row, headers)
    base.style_body_rows(ws, header_row, row_count, last_col)
    base.add_no_column(ws, header_row, row_count, anchor_col_letter="B")

    base.set_widths(ws, [6, 16, 16, 16, 16, 24, 20])
    ws.freeze_panes = f"A{header_row + 1}"


# ---------------------------------------------------------------------------
# 8. 현장 인력 배치
# ---------------------------------------------------------------------------

def build_인력배치(wb: Workbook) -> None:
    ws = wb.create_sheet("현장 인력 배치")
    headers = ["No", "이름", "역할", "외국어가능여부", "연락처", "비고"]
    last_col = len(headers)
    header_row = 4
    row_count = 5

    base.write_title_block(
        ws, "현장 인력 배치", "8/23 현장 상주 인원 배정 (원장 포함 최대 4명)",
        "", last_col,
    )
    base.write_header(ws, header_row, headers)
    base.style_body_rows(ws, header_row, row_count, last_col)
    base.add_no_column(ws, header_row, row_count, anchor_col_letter="B")

    base.add_dropdown(ws, "C", header_row, row_count, ["원장상담보조", "럭키드로우운영", "부스운영", "SNS이벤트"])
    base.add_dropdown(ws, "D", header_row, row_count, ["가능", "불가능"])

    base.set_widths(ws, [6, 14, 16, 14, 16, 20])
    ws.freeze_panes = f"A{header_row + 1}"


# ---------------------------------------------------------------------------
# 실행
# ---------------------------------------------------------------------------

def main() -> None:
    wb = Workbook()
    build_개요(wb)
    build_디자인리스트(wb)
    build_럭드상품리스트(wb)
    build_럭드넘버링(wb)
    build_바우처(wb)
    build_연락처(wb)
    build_인력배치(wb)

    wb.save(OUTPUT_PATH)
    print(f"생성 완료: {OUTPUT_PATH}")


if __name__ == "__main__":
    main()
