"""
이카운트 재고 → Excel 누적 저장 스크립트
─────────────────────────────────────────
매주 금요일 실행해 이카운트 전 창고 재고를 조회하고,
기존 Excel(재고현황.xlsx)에 새 날짜 데이터를 추가한다.

동작 요약:
  1) 이카운트에서 설정된 모든 창고 재고를 한 번에 조회
  2) 창고 그룹(온라인/영업/사업지원TF/컨기부)별로 품목 수량 집계
  3) 기존 Excel이 있으면:
     - 품목코드 기준으로 수량 업데이트
     - 통계 시트에 오늘 날짜 열 추가(또는 덮어쓰기)
     - 기존 입고단가 보존
  4) 기존 Excel이 없으면: 조회 결과로 새 파일 생성
  5) output/ 폴더에 날짜_재고현황.xlsx 저장

사용 예:
  python build_inventory_excel.py                              # 오늘 날짜 기준
  python build_inventory_excel.py --date 20260616             # 날짜 지정
  python build_inventory_excel.py --base 재고현황.xlsx         # 기존 파일 기준으로 업데이트
"""

import sys
import re
import json
import argparse
import traceback
from pathlib import Path
from datetime import datetime, date
from collections import defaultdict

try:
    sys.stdout.reconfigure(encoding="utf-8", errors="replace")
    sys.stderr.reconfigure(encoding="utf-8", errors="replace")
except Exception:
    pass

import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

from ecount_inventory import get_inventory_by_location, extract_rows
from ecount_warehouses import WH_GROUPS, BRAND_SHEET_MAP, ALL_WH_CODES

# ─────────────────────────────────────────
# 브랜드 추출
# ─────────────────────────────────────────

BRAND_ALIASES = {
    "굿스마일컴퍼니": "굿스마일",
    "GSC": "굿스마일",
    "GOOD SMILE": "굿스마일",
    "MEGAHOUSE": "메가하우스",
    "BUSHIROAD": "부시로드",
    "BANDAI": "반다이남코",
    "HOT TOYS": "핫토이",
    "HOTTOYS": "핫토이",
    "HOBBY JAPAN": "하비재팬",
    "HOBBYJAPAN": "하비재팬",
}

KNOWN_BRANDS = list(BRAND_SHEET_MAP.keys())


def extract_brand(prod_des: str) -> str:
    """PROD_DES에서 브랜드명을 추출한다. '[브랜드명]...' 형식 우선."""
    if not prod_des:
        return "기타브랜드"
    m = re.match(r"\[([^\]]+)\]", prod_des.strip())
    if m:
        raw = m.group(1).strip()
        # 정확히 알려진 브랜드면 그대로
        for brand in KNOWN_BRANDS:
            if raw == brand or raw.lower() == brand.lower():
                return brand
        # alias 매핑
        for alias, canonical in BRAND_ALIASES.items():
            if alias.lower() in raw.lower():
                return canonical
        return raw
    # 브랜드 접두어 없으면 기타
    return "기타브랜드"


# ─────────────────────────────────────────
# ECOUNT 조회 + 집계
# ─────────────────────────────────────────

def fetch_all_groups(base_date: str) -> dict:
    """
    모든 창고 그룹의 재고를 조회해 집계한다.
    반환: {
        prod_cd: {
            "prod_des": str,
            "brand": str,
            "groups": { "온라인": qty, "영업": qty, ... }
        }
    }
    """
    # 그룹별 창고 재고를 한 번씩 조회
    group_data: dict[str, dict[str, float]] = {}   # group → {prod_cd: qty}

    for group, wh_codes in WH_GROUPS.items():
        group_data[group] = defaultdict(float)
        for wh_cd in wh_codes:
            print(f"  조회 중: [{group}] 창고 {wh_cd}...", end=" ")
            try:
                raw = get_inventory_by_location(base_date, wh_cd=wh_cd)
                rows = extract_rows(raw)
                for row in rows:
                    if not isinstance(row, dict):
                        continue
                    prod_cd = str(row.get("PROD_CD") or "").strip()
                    if not prod_cd:
                        continue
                    try:
                        qty = float(str(row.get("BAL_QTY") or 0).replace(",", ""))
                    except (ValueError, TypeError):
                        qty = 0.0
                    group_data[group][prod_cd] += qty
                print(f"{len(rows)}건")
            except Exception as e:
                print(f"[실패] {e}")

    # prod_cd 기준으로 통합 (PROD_DES / 브랜드 포함)
    # 마지막으로 조회에 성공한 행에서 prod_des 가져오기 위해 별도 pass
    prod_meta: dict[str, dict] = {}   # prod_cd → {prod_des, brand}

    for group, wh_codes in WH_GROUPS.items():
        for wh_cd in wh_codes:
            try:
                raw = get_inventory_by_location(base_date, wh_cd=wh_cd)
                for row in extract_rows(raw):
                    if not isinstance(row, dict):
                        continue
                    prod_cd = str(row.get("PROD_CD") or "").strip()
                    if prod_cd and prod_cd not in prod_meta:
                        prod_des = str(row.get("PROD_DES") or "").strip()
                        prod_meta[prod_cd] = {
                            "prod_des": prod_des,
                            "brand": extract_brand(prod_des),
                        }
            except Exception:
                pass

    # 통합
    result = {}
    all_prod_cds = set()
    for gd in group_data.values():
        all_prod_cds.update(gd.keys())

    for prod_cd in all_prod_cds:
        meta = prod_meta.get(prod_cd, {"prod_des": prod_cd, "brand": "기타브랜드"})
        result[prod_cd] = {
            "prod_des": meta["prod_des"],
            "brand": meta["brand"],
            "groups": {g: group_data[g].get(prod_cd, 0.0) for g in WH_GROUPS},
        }

    return result


# ─────────────────────────────────────────
# Excel 로드 / 파싱
# ─────────────────────────────────────────

BRAND_SHEET_HEADER = ("브랜드", "품목코드", "품목명[규격]", "입고단가",
                      "온라인 수량", "온라인 재고금액",
                      "영업 수량", "영업 재고금액",
                      "사업지원 수량", "사업지원 재고금액")

DIRECT_SHEET_HEADER = ("재고위치", "바코드", "제품명", "재고수량", "입고단가", "재고금액")


def load_base_excel(path: Path) -> dict:
    """
    기존 Excel에서 {품목코드: {brand, name, price}} 매핑을 로드한다.
    통계 시트의 기존 날짜 열도 읽어온다.
    """
    wb = openpyxl.load_workbook(path, data_only=True)
    prod_map = {}   # prod_cd → {brand, name, price, sheet}

    for sh_name in wb.sheetnames:
        if sh_name == "통계":
            continue
        ws = wb[sh_name]
        rows = list(ws.iter_rows(values_only=True))
        if not rows:
            continue
        header = rows[0]
        # 브랜드별 시트
        if len(header) >= 6 and str(header[1] or "").strip() == "품목코드":
            for row in rows[1:]:
                if not row or row[1] is None:
                    continue
                prod_cd = str(row[1]).strip()
                if not prod_cd:
                    continue
                prod_map[prod_cd] = {
                    "brand": str(row[0] or "").strip(),
                    "name": str(row[2] or "").strip(),
                    "price": row[3],
                    "sheet": sh_name,
                }
        # 직영/메가직영 형식 (바코드 컬럼이 2번째)
        elif len(header) >= 4 and str(header[1] or "").strip() == "바코드":
            for row in rows[1:]:
                if not row or row[1] is None:
                    continue
                prod_cd = str(row[1]).strip()
                if not prod_cd:
                    continue
                prod_map[prod_cd] = {
                    "brand": sh_name,
                    "name": str(row[2] or "").strip(),
                    "price": row[4] if len(row) > 4 else None,
                    "sheet": sh_name,
                }

    wb.close()
    return prod_map


# ─────────────────────────────────────────
# Excel 생성
# ─────────────────────────────────────────

# 스타일 상수
HEADER_FILL = PatternFill("solid", fgColor="4472C4")
HEADER_FONT = Font(bold=True, color="FFFFFF", size=10)
BRAND_FILL  = PatternFill("solid", fgColor="D9E1F2")
THIN = Side(style="thin", color="AAAAAA")
THIN_BORDER = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)
CENTER = Alignment(horizontal="center", vertical="center")


def _write_brand_sheet(ws, brand: str, rows: list, group_cols: list):
    """
    브랜드별 시트를 작성한다.
    rows: [ {prod_cd, name, price, groups:{온라인:qty, ...}} ]
    group_cols: ['온라인', '영업', '사업지원TF', '컨기부'] 중 사용할 것들
    """
    ws.freeze_panes = "A2"

    # 헤더 작성
    headers = ["브랜드", "품목코드", "품목명[규격]", "입고단가"]
    for g in group_cols:
        headers += [f"{g} 수량", f"{g} 재고금액"]

    for col, h in enumerate(headers, 1):
        cell = ws.cell(row=1, column=col, value=h)
        cell.fill = HEADER_FILL
        cell.font = HEADER_FONT
        cell.alignment = CENTER
        cell.border = THIN_BORDER

    # 열 너비
    col_widths = [10, 18, 50, 12] + [12, 16] * len(group_cols)
    for col, w in enumerate(col_widths, 1):
        ws.column_dimensions[get_column_letter(col)].width = w

    # 데이터 행
    for r_idx, item in enumerate(rows, 2):
        price_col = 4
        ws.cell(row=r_idx, column=1, value=brand)
        ws.cell(row=r_idx, column=2, value=item["prod_cd"])
        ws.cell(row=r_idx, column=3, value=item["name"])
        ws.cell(row=r_idx, column=4, value=item["price"])

        col_offset = 5
        for g in group_cols:
            qty = item["groups"].get(g, 0.0)
            qty_cell = ws.cell(row=r_idx, column=col_offset, value=int(qty) if qty == int(qty) else qty)
            price_letter = get_column_letter(price_col)
            qty_letter = get_column_letter(col_offset)
            amt_cell = ws.cell(row=r_idx, column=col_offset + 1,
                               value=f"={qty_letter}{r_idx}*{price_letter}{r_idx}")
            col_offset += 2

    # 합계 행
    total_row = len(rows) + 2
    ws.cell(row=total_row, column=3, value="합계").font = Font(bold=True)
    col_offset = 5
    for g in group_cols:
        qty_letter = get_column_letter(col_offset)
        amt_letter = get_column_letter(col_offset + 1)
        ws.cell(row=total_row, column=col_offset,
                value=f"=SUM({qty_letter}2:{qty_letter}{total_row - 1})").font = Font(bold=True)
        ws.cell(row=total_row, column=col_offset + 1,
                value=f"=SUM({amt_letter}2:{amt_letter}{total_row - 1})").font = Font(bold=True)
        col_offset += 2


def _append_stats_column(ws_stats, ref_date: date, brand_totals: dict):
    """
    통계 시트에 새 날짜 열(수량/금액)을 추가한다.
    brand_totals: {브랜드명: {수량: int, 금액: float}}
    """
    # 1행: 날짜들, 2행: 수량/금액 헤더
    # 마지막 데이터 열 찾기
    max_col = ws_stats.max_column
    # 빈 열 체크 (이미 이 날짜가 있으면 덮어쓰기)
    date_col = None
    for col in range(2, max_col + 1, 2):
        cell_val = ws_stats.cell(row=1, column=col).value
        if isinstance(cell_val, (date, datetime)):
            cell_date = cell_val.date() if isinstance(cell_val, datetime) else cell_val
            if cell_date == ref_date:
                date_col = col
                break

    if date_col is None:
        # 새 열 추가 (다음 빈 짝수 위치)
        date_col = max_col + 1
        if date_col % 2 == 0:
            date_col += 1

    # 날짜 헤더
    ws_stats.cell(row=1, column=date_col, value=ref_date).number_format = "M/D"
    ws_stats.cell(row=1, column=date_col).font = Font(bold=True)
    ws_stats.cell(row=2, column=date_col, value="수량")
    ws_stats.cell(row=2, column=date_col + 1, value="금액")

    # 브랜드별 합계 채우기
    for row in ws_stats.iter_rows(min_row=3, max_row=ws_stats.max_row, values_only=False):
        brand_cell = row[0]
        brand = str(brand_cell.value or "").strip()
        if not brand:
            continue
        totals = brand_totals.get(brand, {})
        ws_stats.cell(row=brand_cell.row, column=date_col, value=totals.get("수량", 0))
        ws_stats.cell(row=brand_cell.row, column=date_col + 1, value=totals.get("금액", 0))


def build_excel(inventory: dict, base_date: str, base_excel: Path | None,
                out_path: Path, group_cols: list | None = None):
    """
    메인 Excel 생성 함수.
    inventory : fetch_all_groups() 결과
    base_date : 'YYYYMMDD'
    base_excel: 기존 Excel 경로 (없으면 None)
    out_path  : 저장 경로
    group_cols: 포함할 그룹 열 목록 (None이면 전체)
    """
    if group_cols is None:
        group_cols = list(WH_GROUPS.keys())

    ref_date = datetime.strptime(base_date, "%Y%m%d").date()

    # 기존 파일 로드 or 새 워크북
    if base_excel and base_excel.exists():
        wb = openpyxl.load_workbook(base_excel)
        prod_map = load_base_excel(base_excel)
        print(f"\n기존 Excel 로드: {base_excel} ({len(prod_map)}개 품목코드)")
    else:
        wb = openpyxl.Workbook()
        # 기본 시트 제거
        if "Sheet" in wb.sheetnames:
            del wb["Sheet"]
        prod_map = {}
        print("\n새 Excel 파일 생성")

    # 브랜드별 품목 그룹핑 (이카운트 데이터 기준)
    brand_items: dict[str, list] = defaultdict(list)
    for prod_cd, info in sorted(inventory.items()):
        brand = info["brand"]
        # 기존 Excel에 있으면 기존 정보 우선
        existing = prod_map.get(prod_cd, {})
        brand_items[brand].append({
            "prod_cd": prod_cd,
            "name": existing.get("name") or info["prod_des"],
            "price": existing.get("price"),
            "groups": info["groups"],
        })

    # 브랜드 시트 생성/업데이트
    for brand, items in sorted(brand_items.items()):
        sheet_name = BRAND_SHEET_MAP.get(brand, brand)[:31]  # Excel 시트명 31자 제한
        if sheet_name in wb.sheetnames:
            del wb[sheet_name]
        ws = wb.create_sheet(title=sheet_name)
        _write_brand_sheet(ws, brand, items, group_cols)
        print(f"  시트 [{sheet_name}]: {len(items)}개 품목")

    # 통계 시트 처리
    brand_totals: dict[str, dict] = {}
    for brand, items in brand_items.items():
        online_qty = sum(it["groups"].get("온라인", 0) for it in items)
        online_amt = sum(
            it["groups"].get("온라인", 0) * (it["price"] or 0)
            for it in items
        )
        brand_totals[brand] = {"수량": int(online_qty), "금액": int(online_amt)}

    if "통계" in wb.sheetnames:
        ws_stats = wb["통계"]
    else:
        ws_stats = wb.create_sheet(title="통계", index=0)
        # 초기 구조 설정
        ws_stats.cell(row=1, column=1, value="온라인사업부")
        ws_stats.cell(row=2, column=1, value=None)
        for r_idx, brand in enumerate(sorted(brand_totals.keys()), 3):
            ws_stats.cell(row=r_idx, column=1, value=brand)

    _append_stats_column(ws_stats, ref_date, brand_totals)

    # 저장
    out_path.parent.mkdir(exist_ok=True)
    wb.save(out_path)
    print(f"\n저장 완료: {out_path}")


# ─────────────────────────────────────────
# main
# ─────────────────────────────────────────

def main():
    parser = argparse.ArgumentParser(description="이카운트 재고 → Excel 누적 저장")
    parser.add_argument("--date", default=datetime.now().strftime("%Y%m%d"),
                        help="기준일자 YYYYMMDD (기본: 오늘)")
    parser.add_argument("--base", default=None,
                        help="기존 Excel 파일 경로 (생략 시 output/ 폴더에서 최신 파일 자동 탐색)")
    parser.add_argument("--groups", default=None,
                        help="포함할 그룹, 쉼표 구분 (예: 온라인,영업). 기본: 전체")
    args = parser.parse_args()

    group_cols = args.groups.split(",") if args.groups else None

    print("=" * 60)
    print(f" 이카운트 재고 Excel 누적 저장")
    print("=" * 60)
    print(f"  기준일자: {args.date}")
    print(f"  그룹:     {group_cols or '전체'}")

    # 기존 파일 탐색
    base_excel = None
    if args.base:
        base_excel = Path(args.base)
    else:
        out_dir = Path("output")
        if out_dir.exists():
            xlsx_files = sorted(out_dir.glob("*재고현황*.xlsx"), reverse=True)
            if xlsx_files:
                base_excel = xlsx_files[0]
                print(f"  기존 파일: {base_excel} (자동 탐색)")

    print(f"\n[1/2] ECOUNT 재고 조회 시작...")
    try:
        inventory = fetch_all_groups(args.date)
    except Exception as e:
        print(f"\n[실패] 재고 조회 오류: {e}")
        traceback.print_exc()
        return 1

    total_items = len(inventory)
    print(f"\n조회 완료: 총 {total_items}개 품목")

    stamp = datetime.now().strftime("%y%m%d")
    out_path = Path("output") / f"{stamp}_재고현황.xlsx"

    print(f"\n[2/2] Excel 생성 중...")
    try:
        build_excel(inventory, args.date, base_excel, out_path, group_cols)
    except Exception as e:
        print(f"\n[실패] Excel 생성 오류: {e}")
        traceback.print_exc()
        return 1

    print("\n" + "=" * 60)
    print(" ✅ 완료!")
    print("=" * 60)
    return 0


if __name__ == "__main__":
    try:
        sys.exit(main())
    except Exception:
        print("\n[예상치 못한 오류]")
        traceback.print_exc()
        sys.exit(1)
