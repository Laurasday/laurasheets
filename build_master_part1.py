#!/usr/bin/env python3
"""통합 워크북 파트 1: 거래처 관리 대장 / 위치순 / 발행일순 / 지인 판매 관리."""

from __future__ import annotations

from pathlib import Path

from openpyxl import Workbook

import build_master_consolidated as m

OUTPUT_PATH = Path(__file__).parent / "쇼핑몰_관리_통합_1.xlsx"


def main() -> None:
    wb = Workbook()
    m.build_거래처관리대장(wb)
    m.build_거래처_view(wb, "거래처 관리 (위치순)", m.WI_FORMULA, "위치정보(주소) 가나다순 자동 정렬 뷰")
    m.build_거래처_view(wb, "거래처 관리 (발행일순)", m.DATE_FORMULA, "세금계산서 발행일(N일) 오름차순 자동 정렬 뷰")
    m.build_지인판매관리(wb)

    wb.save(OUTPUT_PATH)
    print(f"생성 완료: {OUTPUT_PATH}")
    print(wb.sheetnames)


if __name__ == "__main__":
    main()
