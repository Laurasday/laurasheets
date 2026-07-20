#!/usr/bin/env python3
"""통합 워크북 파트 2: 지출 관리 / 에이블리 파트너스 정산 관리 / 리뷰 협찬 관리 대장 / 월별 순수익금.

이 파일은 그 자체로 완결적입니다 (월별 순수익금이 참조하는 지출 관리 +
에이블리 파트너스 정산 관리가 같은 파일 안에 있음).
"""

from __future__ import annotations

from pathlib import Path

from openpyxl import Workbook

import build_master_consolidated as m
import build_shop_excel as base

OUTPUT_PATH = Path(__file__).parent / "쇼핑몰_관리_통합_2.xlsx"


def main() -> None:
    base.ROW_COUNTS["지출 관리"] = 6

    wb = Workbook()
    default_sheet = wb.active
    base.build_지출관리(wb)
    m.build_에이블리파트너스정산관리(wb)
    m.build_리뷰협찬관리대장(wb)
    m.build_월별순수익금(wb)
    wb.remove(default_sheet)

    wb.save(OUTPUT_PATH)
    print(f"생성 완료: {OUTPUT_PATH}")
    print(wb.sheetnames)


if __name__ == "__main__":
    main()
