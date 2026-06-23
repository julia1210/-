"""
프레젠스월드 재고현황 자동화 스크립트

사용법:
    python inventory_update.py \
        --inventory <창고별재고현황.csv> \
        --master <ESA009M_품목마스터.csv> \
        --template <프레젠스월드_재고현황.xlsx> \
        --date 2026-06-20

필요 파일:
    1. 창고별재고현황 CSV: 이카운트 → 재고1 → 출력물 → 창고별재고현황
       (기준일자: 지난주 금요일, 창고: 각부서 하늘색 창고만 선택)
    2. 품목마스터 CSV: ESA009M (입고단가 기준)
    3. 프레젠스월드_재고현황.xlsx: 기존 작업 파일

출력:
    - 프레젠스월드_재고현황_YYYYMMDD.xlsx (업데이트된 파일)
    - 전주 대비 변동 리포트 출력
"""

import argparse
import re
import sys
from datetime import datetime, timedelta
from pathlib import Path

import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment
from openpyxl.utils import get_column_letter


# ──────────────────────────────────────────────
# 설정: 이카운트 창고명 → 엑셀 부서명 매핑
# 이카운트 창고별재고현황에서 나오는 창고명을 아래에 맞게 수정하세요.
# ──────────────────────────────────────────────
WAREHOUSE_TO_DEPT = {
    # 이카운트 창고명: 엑셀 부서 컬럼 접두사
    "온라인창고": "온라인",
    "온라인 창고": "온라인",
    "영업창고": "영업",
    "영업 창고": "영업",
    "위탁매장": "영업_위탁매장",
    "위탁매장창고": "영업_위탁매장",
    "사업지원": "사업지원",
    "사업지원창고": "사업지원",
    "컨기": "컨기",
    "컨텐츠기획": "컨기",
}

# ──────────────────────────────────────────────
# 설정: 브랜드명 정규화 (이카운트 → 엑셀 시트명)
# ──────────────────────────────────────────────
BRAND_TO_SHEET = {
    # 굿스마일
    "굿스마일": "굿스마일",
    "굿스마일컴퍼니": "굿스마일",
    "goodsmile": "굿스마일",
    "good smile": "굿스마일",
    "gsc": "굿스마일",
    # 메가하우스
    "메가하우스": "메가하우스",
    "megahouse": "메가하우스",
    # 부시로드
    "부시로드": "부시로드",
    "bushiroad": "부시로드",
    # 블로키
    "블로키": "블로키 ",
    "blockey": "블로키 ",
    # 반다이남코
    "반다이남코": "반다이남코",
    "bandai namco": "반다이남코",
    "bandainamco": "반다이남코",
    # 핫토이
    "핫토이": "핫토이_온라인",
    "hot toys": "핫토이_온라인",
    "hottoys": "핫토이_온라인",
    "핫토이(hottoys)": "핫토이_온라인",
    # 하비재팬
    "하비재팬": "하비재팬",
    "hobby japan": "하비재팬",
    "hobbyjapan": "하비재팬",
    # CCS
    "ccs toys": "CCS ",
    "ccs": "CCS ",
    # P001
    "p001": "P001 온라인",
    # 코코파
    "코코파": "코코파",
    "코코파스튜디오": "코코파",
    "coco pa": "코코파",
    # 기타
    "기타": "기타브랜드",
    "기타브랜드": "기타브랜드",
    "damtoys": "기타브랜드",
    "aniplex": "기타브랜드",
}

# 각 시트에서 사용하는 부서 컬럼 (수량 컬럼명)
SHEET_DEPT_COLS = {
    "굿스마일": ["온라인 수량", "영업 수량", "영업_위탁매장 수량", "사업지원 수량"],
    "메가하우스": ["온라인 수량", "영업 수량", "사업지원 수량"],
    "부시로드": ["온라인 수량", "영업 수량", "영업_위탁매장 수량", "사업지원 수량"],
    "블로키 ": ["온라인 수량", "영업 수량"],
    "반다이남코": ["온라인 수량", "영업 수량"],
    "핫토이_온라인": ["온라인 수량"],
    "하비재팬": ["온라인 수량"],
    "CCS ": ["온라인 수량"],
    "P001 온라인": ["온라인 수량"],
    "코코파": ["온라인 수량"],
    "기타브랜드": ["온라인 수량", "영업 수량"],
    "미쿠감사제": ["컨기 수량"],
}

# 통계 시트에서 각 부서별 집계 행의 브랜드 목록
STATS_SECTIONS = {
    "온라인사업부": ["굿스마일", "메가하우스", "부시로드", "블로키", "반다이남코", "핫토이", "하비재팬", "CCS TOYS", "P001", "코코파스튜디오", "기타브랜드"],
    "영업사업부": ["굿스마일", "메가하우스", "부시로드", "블로키", "반다이남코", "하비재팬", "CCS TOYS", "코코파스튜디오", "기타브랜드"],
    "영업사업부_위탁매장": ["굿스마일", "메가하우스", "부시로드"],
    "사업지원 TF": ["굿스마일", "메가하우스", "부시로드"],
    "컨텐츠기획부": ["홍대미쿠팝업", "메가하우스"],
}


def parse_args():
    p = argparse.ArgumentParser(description="프레젠스월드 재고현황 자동 업데이트")
    p.add_argument("--inventory", required=True, help="이카운트 창고별재고현황 CSV 경로")
    p.add_argument("--master", required=True, help="이카운트 ESA009M 품목마스터 CSV 경로")
    p.add_argument("--template", required=True, help="프레젠스월드_재고현황.xlsx 경로")
    p.add_argument("--date", help="작업 기준일 (YYYY-MM-DD), 기본값: 오늘")
    p.add_argument("--output", help="출력 파일 경로 (기본: 원본 덮어쓰기)")
    return p.parse_args()


def load_price_master(csv_path: str) -> dict:
    """ESA009M 품목마스터에서 품목명 → 입고단가 딕셔너리 반환"""
    prices = {}
    with open(csv_path, encoding="utf-8-sig") as f:
        lines = f.readlines()

    for line in lines[2:]:  # 회사명 행, 헤더 행 건너뜀
        line = line.strip()
        if not line:
            continue
        parts = re.split(r'"\s*,\s*"', line.strip('"'))
        if len(parts) < 4:
            continue
        name = parts[2].strip().rstrip("\t")
        price_str = parts[3].strip().rstrip("\t").replace(",", "")
        try:
            prices[name] = int(price_str)
        except ValueError:
            prices[name] = None
    return prices


def load_inventory_csv(csv_path: str) -> tuple[list[str], list[dict]]:
    """
    창고별재고현황 CSV 파싱.
    반환: (warehouse_columns, rows)
      - warehouse_columns: 창고 컬럼명 목록
      - rows: [{'브랜드': ..., '품목명': ..., warehouses: {창고명: 수량, ...}}, ...]

    ※ 이카운트 창고별재고현황 CSV 예상 형식:
       브랜드, 품목코드, 품목명, 온라인창고, 영업창고, 위탁매장, 사업지원창고, 합계
    """
    rows = []
    warehouse_cols = []

    with open(csv_path, encoding="utf-8-sig") as f:
        lines = f.readlines()

    # 헤더 찾기 (브랜드 컬럼이 있는 행)
    header_idx = None
    header_cols = []
    for i, line in enumerate(lines):
        clean = line.strip().strip('"')
        parts = [p.strip().strip('"').rstrip("\t") for p in re.split(r'[,\t]', clean)]
        if "브랜드" in parts or "품목명" in parts:
            header_idx = i
            header_cols = parts
            break

    if header_idx is None:
        print("[ERROR] 창고별재고현황 CSV에서 헤더를 찾지 못했습니다.")
        sys.exit(1)

    # 창고 컬럼 추출 (브랜드, 품목코드, 품목명, 입고단가 이후 컬럼)
    fixed_cols = {"브랜드", "품목코드", "품목명", "입고단가", "출고단가", "도매가", "합계", ""}
    warehouse_cols = [c for c in header_cols if c and c not in fixed_cols]

    for line in lines[header_idx + 1:]:
        line = line.strip()
        if not line:
            continue
        parts = re.split(r'"\s*,\s*"', line.strip('"'))
        if len(parts) < 3:
            continue

        # 각 셀에서 탭 제거
        parts = [p.strip().rstrip("\t") for p in parts]

        row = {"브랜드": "", "품목명": ""}
        for i, col in enumerate(header_cols):
            if i >= len(parts):
                break
            if col == "브랜드":
                row["브랜드"] = parts[i]
            elif col == "품목명":
                row["품목명"] = parts[i]
            elif col in warehouse_cols:
                val_str = parts[i].replace(",", "")
                try:
                    row[col] = int(val_str)
                except ValueError:
                    row[col] = 0

        if row["품목명"]:
            rows.append(row)

    return warehouse_cols, rows


def normalize_brand(raw: str) -> str:
    """브랜드명 정규화 → 엑셀 시트명 반환"""
    key = raw.lower().strip()
    return BRAND_TO_SHEET.get(key, BRAND_TO_SHEET.get(raw.strip(), "기타브랜드"))


def map_warehouse_to_dept(wh: str) -> str:
    """창고명 → 엑셀 부서 접두사"""
    return WAREHOUSE_TO_DEPT.get(wh, WAREHOUSE_TO_DEPT.get(wh.strip(), wh))


def update_brand_sheet(ws, sheet_name: str, brand_rows: list, prices: dict, warehouse_cols: list):
    """브랜드 시트의 수량 및 입고단가 업데이트"""
    header = {ws.cell(1, c).value: c for c in range(1, ws.max_column + 1)}

    # 기존 품목명 → 행 번호 매핑
    existing = {}
    for r in range(2, ws.max_row + 1):
        name_val = ws.cell(r, header.get("품목명", 2)).value
        if name_val:
            existing[name_val] = r

    qty_cols = SHEET_DEPT_COLS.get(sheet_name, ["온라인 수량"])

    # 모든 수량 컬럼 0으로 초기화
    for col_name in qty_cols:
        col_idx = header.get(col_name)
        if col_idx:
            for r in range(2, ws.max_row + 1):
                ws.cell(r, col_idx).value = 0

    updated = 0
    new_rows = []

    for row in brand_rows:
        item_name = row["품목명"]
        price = prices.get(item_name)

        # 각 창고 수량을 엑셀 부서 컬럼에 매핑
        qty_map = {}
        for wh in warehouse_cols:
            dept = map_warehouse_to_dept(wh)
            col_name = f"{dept} 수량"
            if col_name in qty_cols and wh in row:
                qty_map[col_name] = qty_map.get(col_name, 0) + row.get(wh, 0)

        if item_name in existing:
            r = existing[item_name]
            # 입고단가 업데이트
            price_col = header.get("입고단가")
            if price_col and price is not None:
                ws.cell(r, price_col).value = price
            # 수량 업데이트
            for col_name, qty in qty_map.items():
                col_idx = header.get(col_name)
                if col_idx:
                    ws.cell(r, col_idx).value = qty
            updated += 1
        else:
            new_rows.append((item_name, price, qty_map))

    # 새 품목 추가
    for item_name, price, qty_map in new_rows:
        r = ws.max_row + 1
        ws.cell(r, header.get("브랜드", 1)).value = ws.cell(2, header.get("브랜드", 1)).value
        ws.cell(r, header.get("품목명", 2)).value = item_name
        price_col = header.get("입고단가")
        if price_col:
            ws.cell(r, price_col).value = price

        for col_name, qty in qty_map.items():
            col_idx = header.get(col_name)
            if col_idx:
                ws.cell(r, col_idx).value = qty

        # 재고금액 수식 추가
        for col_name in header:
            if col_name and "재고금액" in str(col_name):
                amount_col = header[col_name]
                qty_col_name = col_name.replace("재고금액", "수량")
                qty_col = header.get(qty_col_name)
                price_col2 = header.get("입고단가")
                if qty_col and price_col2:
                    ws.cell(r, amount_col).value = f"={get_column_letter(qty_col)}{r}*{get_column_letter(price_col2)}{r}"

    return updated, len(new_rows)


def add_stats_date_column(ws_stats, new_date: datetime, wb: openpyxl.Workbook):
    """통계 시트에 새 날짜 컬럼 2개(수량, 금액) 추가"""
    # 마지막 날짜 컬럼 위치 찾기
    last_date_col = 1
    for c in range(1, ws_stats.max_column + 1):
        val = ws_stats.cell(1, c).value
        if isinstance(val, datetime):
            last_date_col = c

    new_col = last_date_col + 2  # 날짜 + 수량/금액 쌍

    # 날짜 헤더
    date_cell = ws_stats.cell(1, new_col)
    date_cell.value = new_date

    qty_cell = ws_stats.cell(2, new_col)
    qty_cell.value = "수량"
    amt_cell = ws_stats.cell(2, new_col + 1)
    amt_cell.value = "금액"

    print(f"\n[통계 시트] {new_date.strftime('%Y-%m-%d')} 컬럼을 {get_column_letter(new_col)}열에 추가")

    # 각 섹션 브랜드 행에 수식 연결
    for r in range(3, ws_stats.max_row + 1):
        cell_val = ws_stats.cell(r, 1).value
        if not isinstance(cell_val, str):
            continue

        # 브랜드명으로 시트와 합계 행 찾기
        target_sheet, qty_col_letter, amt_col_letter = _find_sheet_summary(cell_val, wb, ws_stats, r)
        if target_sheet:
            sheet_name = target_sheet.title
            ws_stats.cell(r, new_col).value = f"='{sheet_name}'!{qty_col_letter}"
            ws_stats.cell(r, new_col + 1).value = f"='{sheet_name}'!{amt_col_letter}"


def _find_sheet_summary(brand_label: str, wb, ws_stats, stat_row: int):
    """통계 시트 브랜드 행에 연결할 시트와 합계 셀 위치 반환"""
    # 기존 수식 패턴에서 시트와 셀 참조 추출
    for c in range(2, ws_stats.max_column):
        val = ws_stats.cell(stat_row, c).value
        if isinstance(val, str) and val.startswith("="):
            match = re.match(r"='?([^'!]+)'?!([A-Z]+\d+)", val)
            if match:
                sheet_name = match.group(1)
                if sheet_name in wb.sheetnames:
                    return wb[sheet_name], match.group(2), None
    return None, None, None


def print_comparison(wb_old_path: str, wb_new: openpyxl.Workbook, new_date: datetime):
    """전주 대비 변동 분석 출력"""
    print("\n" + "=" * 60)
    print(f"전주 대비 재고 변동 리포트 ({new_date.strftime('%Y-%m-%d')} 기준)")
    print("=" * 60)

    ws_stats = wb_new["통계"]

    sections = ["온라인사업부", "영업사업부", "영업사업부_위탁매장", "사업지원 TF", "컨텐츠기획부"]

    for r in range(1, ws_stats.max_row + 1):
        val = ws_stats.cell(r, 1).value
        if not isinstance(val, str) or "합 계" not in val:
            continue

        # 마지막 두 날짜의 합계 비교
        cols_with_dates = []
        for c in range(2, ws_stats.max_column + 1):
            hdr = ws_stats.cell(r - (r % 2), c).value  # approximate
            if isinstance(hdr, datetime):
                cols_with_dates.append(c)

        if len(cols_with_dates) >= 2:
            prev_qty = ws_stats.cell(r, cols_with_dates[-2]).value or 0
            curr_qty = ws_stats.cell(r, cols_with_dates[-1]).value or 0
            if isinstance(prev_qty, (int, float)) and isinstance(curr_qty, (int, float)):
                diff = curr_qty - prev_qty
                pct = (diff / prev_qty * 100) if prev_qty else 0
                if abs(pct) >= 5:
                    print(f"  행 {r}: 수량 {prev_qty:,} → {curr_qty:,} ({diff:+,}, {pct:+.1f}%)")

    print("=" * 60)


def main():
    args = parse_args()

    work_date = datetime.strptime(args.date, "%Y-%m-%d") if args.date else datetime.today()

    print(f"[1/4] 품목마스터 로딩: {args.master}")
    prices = load_price_master(args.master)
    print(f"      → {len(prices):,}개 품목 로드")

    print(f"\n[2/4] 창고별재고현황 로딩: {args.inventory}")
    warehouse_cols, inv_rows = load_inventory_csv(args.inventory)
    print(f"      → 창고 컬럼: {warehouse_cols}")
    print(f"      → {len(inv_rows):,}개 품목 로드")

    # 브랜드별로 분류
    by_sheet: dict[str, list] = {}
    for row in inv_rows:
        sheet_name = normalize_brand(row["브랜드"])
        by_sheet.setdefault(sheet_name, []).append(row)
    unmapped = [b for b in set(r["브랜드"] for r in inv_rows) if normalize_brand(b) == "기타브랜드" and b not in ("기타", "기타브랜드")]
    if unmapped:
        print(f"      ※ 기타브랜드로 분류된 브랜드: {set(unmapped)}")

    print(f"\n[3/4] 엑셀 파일 업데이트: {args.template}")
    wb = openpyxl.load_workbook(args.template)

    total_updated = total_new = 0
    for sheet_name, rows in by_sheet.items():
        if sheet_name not in wb.sheetnames:
            print(f"      [SKIP] 시트 없음: {sheet_name}")
            continue
        ws = wb[sheet_name]
        updated, new = update_brand_sheet(ws, sheet_name, rows, prices, warehouse_cols)
        total_updated += updated
        total_new += new
        print(f"      [{sheet_name}] 업데이트 {updated}건, 신규 {new}건")

    # 통계 시트 날짜 컬럼 추가
    add_stats_date_column(wb["통계"], work_date, wb)

    output_path = args.output or args.template
    wb.save(output_path)
    print(f"\n[4/4] 저장 완료: {output_path}")
    print(f"      총 업데이트 {total_updated}건, 신규 추가 {total_new}건")

    print_comparison(args.template, wb, work_date)


if __name__ == "__main__":
    main()
