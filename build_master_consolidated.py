#!/usr/bin/env python3
"""쇼핑몰 관리 통합 워크북 생성 스크립트.

원본 라이브 구글시트(거래처 관리 대장 / 지인 판매 관리 / 에이블리 파트너스 정산 관리 /
리뷰 협찬 관리 대장 — 실제 데이터 포함)를 그대로 재현하고, 여기에 지출 관리 /
거래처 관리(위치순) / 거래처 관리(발행일순) / 월별 순수익금 4개 탭을 추가한
통합 워크북을 만든다. 전부 같은 파일 안에 있으므로 IMPORTRANGE 없이
시트 간 수식으로 바로 연동된다.

실행: python3 build_master_consolidated.py
"""

from __future__ import annotations

from pathlib import Path

from openpyxl import Workbook
from openpyxl.styles import Font
from openpyxl.worksheet.datavalidation import DataValidation

import build_shop_excel as base

OUTPUT_PATH = Path(__file__).parent / "쇼핑몰_관리_통합.xlsx"


# ---------------------------------------------------------------------------
# 1. 거래처 관리 대장 (실제 데이터 — 원본 라이브 시트 그대로 재현)
# ---------------------------------------------------------------------------

거래처_HEADERS = [
    "No", "거래처명", "위치정보(주소)", "담당자/연락처", "메신저",
    "개인평가 (★1~5)", "평가 메모", "계좌번호(은행)",
    "세금계산서 발행 여부", "세금계산서 발행일", "샘플 가능 여부", "비고",
]

거래처_REAL_ROWS = [
    ["위디(withee)", "디오트 B1 I-18호", "010-6789-6764", "K : withee2026", "★4",
     "디자인은 예쁜데 가격대 비쌈", "3010377709921(정한얼/농협)", "O", "6일", "", ""],
    ["시우팀팀", "디오트 1층 G-14호", "010-3935-4187 02-2117-4332", "", "★1",
     "하나 사입해왔는데 품절됨", "20791035298507(하나/한지연)", "O", "10일", "", ""],
    ["키치키치(kichkich) 마젠타(에이블리용)", "남평화 3층 가 33호", "010-9108-3331", "", "★3",
     "옷 저렴&예쁨/ 다만 무서움", "02314285104019(기업/김지혜)", "O", "5일", "", ""],
    ["브리니", "디오트 1층 A-15", "010-9231-4549", "", "★5",
     "", "140015691122(신한/(주)브리니랩)", "O", "5일", "", "입금 시 상호명으로"],
    ["퍼플그레이", "남평화 3층 100, 103호", "010-9639-3221", "K : with_button", "★5",
     "", "02314354504013(기업/위드버튼 고은석)", "O", "5일", "", ""],
    ["앨라", "누죤 지하1층 514", "010-5098-5860", "", "★1",
     "신마 이용 /장끼 X", "", "", "", "", ""],
    ["mineD 마인디", "디오트 지하2층 D24", "010-6518-4108 02-2117-4108", "K : mined24 I : mined4108", "★2",
     "신마 이용 /장끼 X / 품절건 카톡 문의 > 친절", "02211131104016(기업/박윤정)", "O", "10일", "", "화이트 컬러 제외 깔교 가능(2주 이내)"],
]

거래처_ROW_COUNT = 9


def build_거래처관리대장(wb: Workbook) -> None:
    ws = wb.active
    ws.title = "거래처 관리 대장"
    last_col = len(거래처_HEADERS)
    header_row = 4

    base.write_title_block(
        ws, "거래처 관리 대장", "사입처 정보 · 평가 · 세금계산서 · 계좌 · 샘플 가능 여부",
        "신마주문 메모X (매장경험 無)", last_col,
    )

    base.write_header(ws, header_row, 거래처_HEADERS)
    base.style_body_rows(ws, header_row, 거래처_ROW_COUNT, last_col)
    base.add_no_column(ws, header_row, 거래처_ROW_COUNT)

    for i, row_values in enumerate(거래처_REAL_ROWS):
        row = header_row + 1 + i
        for col_offset, value in enumerate(row_values, start=2):
            if value:
                ws.cell(row=row, column=col_offset, value=value)

    base.add_dropdown(ws, "F", header_row, 거래처_ROW_COUNT, ["★1", "★2", "★3", "★4", "★5"])
    base.add_dropdown(ws, "I", header_row, 거래처_ROW_COUNT, ["필요", "불필요", "요청완료", "발행완료", "O"])
    base.add_dropdown(ws, "K", header_row, 거래처_ROW_COUNT, ["가능", "불가능", "조건부가능"])

    base.set_widths(ws, [6, 24, 20, 20, 18, 10, 26, 22, 14, 12, 12, 24])
    ws.freeze_panes = f"A{header_row + 1}"


# ---------------------------------------------------------------------------
# 2 & 3. 거래처 관리 위치순 / 발행일순 뷰 (같은 파일 내 SORT 수식, 자동 연동)
# ---------------------------------------------------------------------------

def build_거래처_view(wb: Workbook, sheet_name: str, sort_key_formula: str, description: str) -> None:
    ws = wb.create_sheet(sheet_name)
    last_col = len(거래처_HEADERS)
    header_row = 4

    base.write_title_block(ws, sheet_name, description,
                            "※ '거래처 관리 대장' 탭 데이터를 자동으로 정렬해서 보여줍니다. 원본 데이터를 수정하면 여기도 같이 바뀝니다.",
                            last_col)
    base.write_header(ws, header_row, 거래처_HEADERS)
    ws.cell(row=header_row + 1, column=1, value=sort_key_formula)

    base.set_widths(ws, [6, 24, 20, 20, 18, 10, 26, 22, 14, 12, 12, 24])
    ws.freeze_panes = f"A{header_row + 1}"


거래처_DATA_START = 5
거래처_DATA_END = 4 + 거래처_ROW_COUNT

WI_FORMULA = (
    f"=SORT('거래처 관리 대장'!A{거래처_DATA_START}:L{거래처_DATA_END}, 3, TRUE)"
)
DATE_FORMULA = (
    f"=SORT('거래처 관리 대장'!A{거래처_DATA_START}:L{거래처_DATA_END}, "
    f'IFERROR(VALUE(REGEXREPLACE(\'거래처 관리 대장\'!J{거래처_DATA_START}:J{거래처_DATA_END},"[^0-9]","")),9999), TRUE)'
)


# ---------------------------------------------------------------------------
# 4. 지출 관리 (build_shop_excel.build_지출관리 재사용)
# ---------------------------------------------------------------------------

# main()에서 base.build_지출관리(wb) 호출


# ---------------------------------------------------------------------------
# 5. 지인 판매 관리 (원본 라이브 시트의 확장 컬럼 구조 그대로 재현, 데이터 없음)
# ---------------------------------------------------------------------------

지인판매_HEADERS = [
    "No", "날짜", "지인명", "상품명", "도매가", "세금(20%)",
    "판매가 (도매가+세금)", "결제여부", "비고", "결제수단", "비고",
]


def build_지인판매관리(wb: Workbook) -> None:
    ws = wb.create_sheet("지인 판매 관리")
    last_col = len(지인판매_HEADERS)
    header_row = 4
    row_count = 8
    sum_row = header_row + 1 + row_count

    base.write_title_block(
        ws, "지인 판매 관리", "도매가 + 세금 포함 금액으로 판매하는 지인 거래 관리",
        "※ 판매가 = 도매가 + 세금(도매가의 20%, 안전 버퍼 반영) — 지인에게는 이 금액으로만 판매",
        last_col,
    )
    base.write_header(ws, header_row, 지인판매_HEADERS)
    base.style_body_rows(ws, header_row, row_count, last_col)
    base.add_no_column(ws, header_row, row_count)
    base.add_dropdown(ws, "H", header_row, row_count, ["입금완료", "미입금"])
    base.add_dropdown(ws, "J", header_row, row_count, ["현금", "계좌이체", "카드"])
    base.set_column_formats(ws, header_row, row_count, "B", base.DATE_FORMAT)

    start, end = header_row + 1, header_row + row_count
    for row in range(start, end + 1):
        ws[f"F{row}"] = f"=E{row}*0.2"
        ws[f"G{row}"] = f"=E{row}+F{row}"
        for col_letter in ("E", "F", "G"):
            ws[f"{col_letter}{row}"].number_format = base.MONEY_FORMAT

    base.add_sum_row(ws, sum_row, 3, ["E", "F", "G"], header_row, row_count)
    base.set_widths(ws, [6, 12, 14, 18, 12, 12, 16, 12, 16, 12, 16])
    ws.freeze_panes = f"A{header_row + 1}"


# ---------------------------------------------------------------------------
# 6. 에이블리 파트너스 정산 관리 (원본 제목/메모 그대로, 데이터 없음)
# ---------------------------------------------------------------------------

정산_HEADERS = [
    "No", "정산월", "판매금액(매출)", "에이블리 수수료 (10%)", "정산 예정금액 (90%)",
    "원천징수세 (필요시 입력)", "실수령액", "정산 여부", "비고",
]


def build_에이블리파트너스정산관리(wb: Workbook) -> None:
    ws = wb.create_sheet("에이블리 파트너스 정산 관리")
    last_col = len(정산_HEADERS)
    header_row = 5
    row_count = 8
    sum_row = header_row + 1 + row_count

    base.write_title_block(
        ws, "에이블리 파트너스 정산 관리", "판매금의 10%를 에이블리에서 수수료로 수취 · 세금 처리는 별도 확인 필요",
        "※ 세금(사업소득세 원천징수 등) 처리 방식은 아직 수익이 없어 확정되지 않음 — 실제 정산 발생 시 세무사 등 전문가 확인 권장",
        last_col,
    )
    base.write_header(ws, header_row, 정산_HEADERS)
    base.style_body_rows(ws, header_row, row_count, last_col)
    base.add_no_column(ws, header_row, row_count)
    base.add_dropdown(ws, "H", header_row, row_count, ["정산대기", "정산완료"])
    base.set_column_formats(ws, header_row, row_count, "B", base.MONTH_FORMAT)

    start, end = header_row + 1, header_row + row_count
    for row in range(start, end + 1):
        ws[f"D{row}"] = f"=C{row}*0.1"
        ws[f"E{row}"] = f"=C{row}-D{row}"
        ws[f"G{row}"] = f"=E{row}-F{row}"
        for col_letter in ("C", "D", "E", "F", "G"):
            ws[f"{col_letter}{row}"].number_format = base.MONEY_FORMAT

    base.add_sum_row(ws, sum_row, 2, ["C", "D", "E", "F", "G"], header_row, row_count)
    base.set_widths(ws, [6, 12, 16, 16, 16, 18, 14, 12, 20])
    ws.freeze_panes = f"A{header_row + 1}"


SETTLEMENT_HEADER_ROW = 5
SETTLEMENT_ROW_COUNT = 8


# ---------------------------------------------------------------------------
# 7. 리뷰 협찬 관리 대장 (원본 제목/드롭다운 문구 그대로, 데이터 없음)
# ---------------------------------------------------------------------------

리뷰_HEADERS = [
    "No", "구분 (지인/인플루언서)", "이름/계정", "연락처", "플랫폼 (인스타/블로그/기타)",
    "협찬 상품", "리뷰어 결제금액", "페이백 금액 (결제금액 전액)", "리뷰 링크", "리뷰 게시일",
    "리뷰 작성 여부", "페이백 여부", "삭제금지 종료일", "삭제금지 상태", "비고",
]


def build_리뷰협찬관리대장(wb: Workbook) -> None:
    ws = wb.create_sheet("리뷰 협찬 관리 대장")
    last_col = len(리뷰_HEADERS)
    header_row = 4
    row_count = 8
    sum_row = header_row + 1 + row_count

    base.write_title_block(
        ws, "리뷰 협찬 관리 대장", "지인/인플루언서 리뷰 협찬 · 페이백 · 최소 3개월 삭제금지 계약 관리",
        "※ 진행방식: 리뷰어가 상품 결제 → 리뷰 게시 확인 → 결제금액 전액 페이백 / 리뷰는 게시일 기준 최소 3개월(EDATE 3개월) 삭제 금지",
        last_col,
    )
    base.write_header(ws, header_row, 리뷰_HEADERS)
    base.style_body_rows(ws, header_row, row_count, last_col)
    base.add_no_column(ws, header_row, row_count)

    base.add_dropdown(ws, "B", header_row, row_count, ["지인", "인플루언서"])
    base.add_dropdown(ws, "E", header_row, row_count, ["인스타", "블로그", "기타"])
    base.add_dropdown(ws, "K", header_row, row_count, ["완료", "미완료"])
    base.add_dropdown(ws, "L", header_row, row_count, ["완료", "미완료"])
    base.set_column_formats(ws, header_row, row_count, "J", base.DATE_FORMAT)
    base.set_column_formats(ws, header_row, row_count, "M", base.DATE_FORMAT)

    start, end = header_row + 1, header_row + row_count
    for row in range(start, end + 1):
        ws[f"H{row}"] = f"=G{row}"
        ws[f"M{row}"] = f'=IF(J{row}="","",EDATE(J{row},3))'
        ws[f"N{row}"] = f'=IF(J{row}="","",IF(TODAY()<M{row},"삭제금지 기간","기간 만료(삭제가능)"))'
        for col_letter in ("G", "H"):
            ws[f"{col_letter}{row}"].number_format = base.MONEY_FORMAT

    base.add_sum_row(ws, sum_row, 3, ["G", "H"], header_row, row_count)
    base.set_widths(ws, [6, 14, 14, 14, 16, 16, 14, 18, 16, 12, 12, 12, 14, 20, 20])
    ws.freeze_panes = f"A{header_row + 1}"


# ---------------------------------------------------------------------------
# 8. 월별 순수익금 (같은 파일 안의 지출 관리 + 에이블리 파트너스 정산 관리 참조, IMPORTRANGE 불필요)
# ---------------------------------------------------------------------------

def build_월별순수익금(wb: Workbook) -> None:
    ws = wb.create_sheet("월별 순수익금")
    last_col = 4
    header_row = 4

    base.write_title_block(
        ws, "월별 순수익금", "지출 관리 + 에이블리 정산 실수령액을 월별로 집계",
        "※ 같은 파일 안의 '지출 관리'와 '에이블리 파트너스 정산 관리' 탭을 자동으로 참조합니다.",
        last_col,
    )
    base.write_header(ws, header_row, ["월", "지출 합계", "수익(정산금) 합계", "순수익금"])

    row_count = 12
    base.style_body_rows(ws, header_row, row_count, last_col)

    exp_start = 5
    exp_end = 4 + base.ROW_COUNTS["지출 관리"]
    settle_start = SETTLEMENT_HEADER_ROW + 1
    settle_end = SETTLEMENT_HEADER_ROW + SETTLEMENT_ROW_COUNT

    start, end = header_row + 1, header_row + row_count
    for i, row in enumerate(range(start, end + 1)):
        month_index = i + 1
        ws.cell(row=row, column=1, value=f"=DATE(2026,{month_index},1)")
        ws.cell(row=row, column=1).number_format = base.MONTH_FORMAT
        ws.cell(
            row=row, column=2,
            value=(
                f"=SUMPRODUCT((TEXT('지출 관리'!$B${exp_start}:$B${exp_end},\"yyyy-mm\")=TEXT($A{row},\"yyyy-mm\"))"
                f"*'지출 관리'!$H${exp_start}:$H${exp_end})"
            ),
        )
        ws.cell(
            row=row, column=3,
            value=(
                f"=SUMPRODUCT((TEXT('에이블리 파트너스 정산 관리'!$B${settle_start}:$B${settle_end},\"yyyy-mm\")"
                f"=TEXT($A{row},\"yyyy-mm\"))*'에이블리 파트너스 정산 관리'!$G${settle_start}:$G${settle_end})"
            ),
        )
        ws.cell(row=row, column=4, value=f"=C{row}-B{row}")
        for col_letter in ("B", "C", "D"):
            ws[f"{col_letter}{row}"].number_format = base.MONEY_FORMAT

    sum_row = end + 1
    base.add_sum_row(ws, sum_row, 1, ["B", "C", "D"], header_row, row_count)
    base.set_widths(ws, [12, 14, 16, 14])
    ws.freeze_panes = f"A{header_row + 1}"


# ---------------------------------------------------------------------------
# 실행
# ---------------------------------------------------------------------------

def main() -> None:
    base.ROW_COUNTS["지출 관리"] = 8

    wb = Workbook()
    build_거래처관리대장(wb)
    build_거래처_view(wb, "거래처 관리 (위치순)", WI_FORMULA, "위치정보(주소) 가나다순 자동 정렬 뷰")
    build_거래처_view(wb, "거래처 관리 (발행일순)", DATE_FORMULA, "세금계산서 발행일(N일) 오름차순 자동 정렬 뷰")
    base.build_지출관리(wb)
    build_지인판매관리(wb)
    build_에이블리파트너스정산관리(wb)
    build_리뷰협찬관리대장(wb)
    build_월별순수익금(wb)

    wb.save(OUTPUT_PATH)
    print(f"생성 완료: {OUTPUT_PATH}")
    print(wb.sheetnames)


if __name__ == "__main__":
    main()
