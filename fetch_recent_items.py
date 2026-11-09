#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
QA 환경 100일 이내 재고 연동 상품 조회 자동화 도구
- 제휴사별 / 채널별 100일 이내 재고 있는 상태로 연동된 상품(lnkgItemId)을 조회합니다.
- 상품번호와 영업점/사이트번호를 분리하여 표준 양식으로 CSV 및 Excel(xlsx)을 저장합니다.
- 앞자리 0(Leading Zeros)이 엑셀에서 삭제되지 않도록 텍스트 타입으로 안전하게 처리합니다.
"""

import sys
import json
import urllib.request
import urllib.error
import argparse
import csv
from concurrent.futures import ThreadPoolExecutor, as_completed
from datetime import datetime
import subprocess

try:
    import openpyxl
    from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
    HAS_OPENPYXL = True
except ImportError:
    HAS_OPENPYXL = False

API_URL_TEMPLATE = "https://qa-cloudapi.ssgadm.com/alln/market-collect-oracle/alliance-mcp/recent-linkage-items?chnlId={chnl_id}"

STORE_SITE_MAP = {
    # 영업점 번호 (백화점, 점포택배, 판매자스타배송 등)
    "1002": "본점(1002)",
    "1003": "타임스퀘어점 리빙관(1003)",
    "1004": "강남점(1004)",
    "1006": "신세계 사우스시티점(1006)",
    "1007": "의정부점(1007)",
    "1008": "광주신세계(1008)",
    "1009": "센텀시티점(1009)",
    "1010": "마산점(1010)",
    "1011": "천안아산점(1011)",
    "1013": "타임스퀘어점 패션관(1013)",
    "1017": "김해점(1017)",
    "1018": "하남점(1018)",
    "1019": "대구신세계(1019)",
    "1020": "시코르강남역점(1020)",
    "1021": "팩토리스토어 강남점(1021)",
    "1022": "팩토리스토어 고양점(1022)",
    "1023": "온라인점 (S파트너스)(1023)",
    "1024": "대전신세계 (ART & SCIENCE)(1024)",
    "1026": "팩토리스토어 타임스퀘어점(1026)",
    "1027": "시코르닷컴(C3)(1027)",
    "1029": "신세계푸드마켓 도곡점(1029)",
    "1030": "신세계백화점 화장품 온라인점(배스앤바디웍스 온라인점)(1030)",
    "2551": "NE.O.004(2551)",
    # 사이트 번호
    "6001": "이마트몰(6001)",
    "6002": "트레이더스(6002)",
    "6004": "신세계몰(6004)",
    "7008": "아울렛(7008)",
}

CHANNEL_DATA = [
    {
        "partner": "G마켓",
        "mainChnlId": "0000015780",
        "channels": [
            {"name": "신세계몰 (온라인)", "chnlId": "0000015781"},
            {"name": "신세계몰 (직매입)", "chnlId": "0000016345"},
            {"name": "신세계백화점", "chnlId": "0000015782"},
            {"name": "프리미엄아울렛", "chnlId": "0000017105"},
            {"name": "점포택배", "chnlId": "0000018380"},
            {"name": "판매자스타배송", "chnlId": "0000021357"},
            {"name": "장보기 (이마트몰/새벽배송)", "chnlId": "0000018820"},
            {"name": "장보기 (트레이더스)", "chnlId": "0000023802"},
            {"name": "1day", "chnlId": "0000019787"},
            {"name": "1day 트레이더스", "chnlId": "0000020006"},
        ]
    },
    {
        "partner": "옥션",
        "mainChnlId": "0000015777",
        "channels": [
            {"name": "신세계몰 (온라인)", "chnlId": "0000015778"},
            {"name": "신세계몰 (직매입)", "chnlId": "0000016346"},
            {"name": "신세계백화점", "chnlId": "0000015779"},
            {"name": "프리미엄아울렛", "chnlId": "0000017106"},
            {"name": "점포택배", "chnlId": "0000018381"},
            {"name": "판매자스타배송", "chnlId": "0000021358"},
            {"name": "1day", "chnlId": "0000019788"},
            {"name": "1day 트레이더스", "chnlId": "0000020007"},
        ]
    },
    {
        "partner": "라이브쇼핑",
        "mainChnlId": "0000019622",
        "channels": [
            {"name": "신세계몰", "chnlId": "0000019624"},
            {"name": "신세계백화점", "chnlId": "0000019623"},
        ]
    },
    {
        "partner": "11번가",
        "mainChnlId": "0000015958",
        "channels": [
            {"name": "신세계몰", "chnlId": "0000015959"},
            {"name": "신세계백화점", "chnlId": "0000015960"},
            {"name": "프리미엄아울렛", "chnlId": "0000017104"},
            {"name": "장보기 (이마트몰/새벽배송)", "chnlId": "0000020717"},
        ]
    },
    {
        "partner": "네이버",
        "mainChnlId": "0000016082",
        "channels": [
            {"name": "신세계백화점", "chnlId": "0000016083"},
            {"name": "점포택배", "chnlId": "0000018560"},
            {"name": "장보기 (이마트몰)", "chnlId": "0000018200"},
            {"name": "장보기 (새벽배송)", "chnlId": "0000018201"},
            {"name": "장보기 (트레이더스)", "chnlId": "0000018202"},
        ]
    },
    {
        "partner": "W컨셉",
        "mainChnlId": "0000019089",
        "channels": [
            {"name": "신세계몰", "chnlId": "0000019091"},
            {"name": "신세계백화점", "chnlId": "0000019090"},
        ]
    },
    {
        "partner": "카카오",
        "mainChnlId": "0000020537",
        "channels": [
            {"name": "신세계백화점", "chnlId": "0000020540"},
        ]
    },
    {
        "partner": "CTK",
        "mainChnlId": "0000021953",
        "channels": [
            {"name": "신세계백화점", "chnlId": "0000021954"},
        ]
    }
]

def parse_linkage_item_id(lnkg_item_id):
    """lnkgItemId를 상품번호와 영업점/사이트번호로 분리"""
    if not lnkg_item_id or len(lnkg_item_id) <= 4:
        return lnkg_item_id, "-"
    item_no = lnkg_item_id[:-4]
    suffix_code = lnkg_item_id[-4:]
    store_info = STORE_SITE_MAP.get(suffix_code, f"{suffix_code}")
    return item_no, store_info

def get_channel_map():
    mapping = {}
    for group in CHANNEL_DATA:
        partner = group["partner"]
        main_id = group["mainChnlId"]
        for ch in group["channels"]:
            mapping[ch["chnlId"]] = {
                "partner": partner,
                "mainChnlId": main_id,
                "name": ch["name"],
                "chnlId": ch["chnlId"]
            }
    return mapping

def fetch_items(chnl_id, timeout=10):
    url = API_URL_TEMPLATE.format(chnl_id=chnl_id)
    req = urllib.request.Request(
        url,
        headers={"User-Agent": "Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) QA-Linkage-Tool"}
    )
    try:
        with urllib.request.urlopen(req, timeout=timeout) as resp:
            data = json.loads(resp.read().decode("utf-8"))
            if data.get("status") == 200:
                return {"chnlId": chnl_id, "success": True, "items": data.get("data", [])}
            return {"chnlId": chnl_id, "success": False, "error": data.get("message", "Unknown error"), "items": []}
    except Exception as e:
        return {"chnlId": chnl_id, "success": False, "error": str(e), "items": []}

def print_channel_list():
    print("\n" + "="*80)
    print(f"{'채널명(제휴사)':<12} | {'대표채널ID':<12} | {'서비스명(하위채널)':<25} | {'채널ID':<15}")
    print("="*80)
    for group in CHANNEL_DATA:
        partner = group["partner"]
        main_id = group["mainChnlId"]
        for idx, ch in enumerate(group["channels"]):
            p_display = partner if idx == 0 else ""
            m_display = main_id if idx == 0 else ""
            print(f"{p_display:<12} | {m_display:<12} | {ch['name']:<25} | {ch['chnlId']:<15}")
    print("="*80 + "\n")

def save_to_csv(rows, filename):
    """엑셀에서 앞자리 0이 날아가지 않도록 수식 텍스트(=\"0000...\")로 보존"""
    with open(filename, mode="w", encoding="utf-8-sig", newline="") as f:
        fieldnames = ["채널명", "서비스명", "채널ID", "상품번호", "영업점번호", "마지막연동일자"]
        writer = csv.writer(f)
        writer.writerow(fieldnames)
        for r in rows:
            # 채널ID와 상품번호에 엑셀 텍스트 보존 수식 적용
            chnl_id_cell = f'="{r["채널ID"]}"' if r["채널ID"].startswith("0") else r["채널ID"]
            item_no_cell = f'="{r["상품번호"]}"' if r["상품번호"].startswith("0") else r["상품번호"]
            writer.writerow([
                r["채널명"],
                r["서비스명"],
                chnl_id_cell,
                item_no_cell,
                r["영업점번호"],
                r["마지막연동일자"]
            ])
    print(f"[SUCCESS] CSV 파일 저장 완료: {filename} (총 {len(rows)}행, 0 보존 텍스트 처리)")

def save_to_excel(rows, filename):
    """openpyxl을 사용하여 순수 텍스트 서식(@)으로 .xlsx 저장 (가장 깔끔함)"""
    if not HAS_OPENPYXL:
        print("[WARN] openpyxl 라이브러리가 없어 Excel(.xlsx) 생성을 건너뜁니다.")
        return

    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "연동상품목록"

    headers = ["채널명", "서비스명", "채널ID", "상품번호", "영업점번호", "마지막연동일자"]
    ws.append(headers)

    # 헤더 스타일
    header_fill = PatternFill(start_color="F1F5F9", end_color="F1F5F9", fill_type="solid")
    header_font = Font(name="맑은 고딕", size=11, bold=True, color="1E293B")
    border_thin = Border(
        left=Side(style='thin', color='E2E8F0'),
        right=Side(style='thin', color='E2E8F0'),
        top=Side(style='thin', color='CBD5E1'),
        bottom=Side(style='thin', color='CBD5E1')
    )

    for col_idx in range(1, len(headers) + 1):
        cell = ws.cell(row=1, column=col_idx)
        cell.fill = header_fill
        cell.font = header_font
        cell.alignment = Alignment(horizontal="center", vertical="center")
        cell.border = border_thin

    # 데이터 입력 (모두 문자열로 강제 지정하여 앞의 0 유지)
    body_font = Font(name="맑은 고딕", size=10)
    for r_idx, r in enumerate(rows, start=2):
        row_vals = [
            r["채널명"],
            r["서비스명"],
            str(r["채널ID"]),
            str(r["상품번호"]),
            r["영업점번호"],
            r["마지막연동일자"]
        ]
        for c_idx, val in enumerate(row_vals, start=1):
            cell = ws.cell(row=r_idx, column=c_idx)
            cell.value = str(val)
            cell.number_format = '@' # Text format
            cell.font = body_font
            cell.border = border_thin
            if c_idx in [1, 2, 5]:
                cell.alignment = Alignment(horizontal="left", vertical="center")
            elif c_idx in [3, 4]:
                cell.alignment = Alignment(horizontal="center", vertical="center")
            else:
                cell.alignment = Alignment(horizontal="center", vertical="center")

    # 열 너비 자동 조정
    ws.column_dimensions['A'].width = 14
    ws.column_dimensions['B'].width = 28
    ws.column_dimensions['C'].width = 16
    ws.column_dimensions['D'].width = 22
    ws.column_dimensions['E'].width = 24
    ws.column_dimensions['F'].width = 22

    wb.save(filename)
    print(f"[SUCCESS] Excel(.xlsx) 저장 완료: {filename} (총 {len(rows)}행, 텍스트 서식 완벽 보존)")

def run_query(target_channels, save_csv=None, save_xlsx=None, copy_to_clipboard=False):
    print(f"\n[INFO] 총 {len(target_channels)}개 채널 조회 시작...")
    results = []
    
    with ThreadPoolExecutor(max_workers=8) as executor:
        future_to_ch = {
            executor.submit(fetch_items, ch["chnlId"]): ch
            for ch in target_channels
        }
        for future in as_completed(future_to_ch):
            ch_info = future_to_ch[future]
            res = future.result()
            results.append({
                "partner": ch_info.get("partner", ""),
                "mainChnlId": ch_info.get("mainChnlId", ""),
                "channelName": ch_info.get("name", ""),
                "chnlId": ch_info["chnlId"],
                "success": res["success"],
                "error": res.get("error", ""),
                "items": res.get("items", [])
            })

    # 정렬: 제휴사 순서 유지
    partner_order = {group["partner"]: i for i, group in enumerate(CHANNEL_DATA)}
    results.sort(key=lambda x: (partner_order.get(x["partner"], 999), x["chnlId"]))

    total_items_count = sum(len(r["items"]) for r in results)
    all_item_nos = []

    print("\n" + "="*115)
    print(f"{'채널명':<8} | {'서비스명':<22} | {'채널ID':<12} | {'건수':<5} | {'최근연동일자':<20} | {'샘플 (상품번호 / 영업점)'}")
    print("="*115)
    
    csv_rows = []
    for r in results:
        count = len(r["items"])
        if not r["success"]:
            sample = f"조회 실패: {r['error']}"
            latest_dts = "-"
        elif count == 0:
            sample = "데이터 없음"
            latest_dts = "-"
        else:
            latest_dts = r["items"][0].get("modDts", "-")
            samples = []
            for item in r["items"][:2]:
                it_no, st_info = parse_linkage_item_id(item.get("lnkgItemId", ""))
                samples.append(f"{it_no} ({st_info})")
            sample = ", ".join(samples) + (f" ... 외 {count-2}건" if count > 2 else "")
            
            for item in r["items"]:
                raw_id = item.get("lnkgItemId", "")
                item_no, store_info = parse_linkage_item_id(raw_id)
                all_item_nos.append(item_no)
                csv_rows.append({
                    "채널명": r["partner"],
                    "서비스명": r["channelName"],
                    "채널ID": r["chnlId"],
                    "상품번호": item_no,
                    "영업점번호": store_info,
                    "마지막연동일자": item.get("modDts", "")
                })

        print(f"{r['partner']:<8} | {r['channelName']:<22} | {r['chnlId']:<12} | {count:<5} | {latest_dts:<20} | {sample}")

    print("="*115)
    print(f"총 {len(results)}개 채널 중 상품 보유: {sum(1 for r in results if len(r['items']) > 0)}개, 총 상품 수: {total_items_count}건\n")

    if save_csv:
        if not csv_rows:
            print("[WARN] 조회된 상품이 없어 CSV를 생성하지 않았습니다.")
        else:
            save_to_csv(csv_rows, save_csv)

    if save_xlsx:
        if not csv_rows:
            print("[WARN] 조회된 상품이 없어 Excel을 생성하지 않았습니다.")
        else:
            save_to_excel(csv_rows, save_xlsx)

    if copy_to_clipboard:
        if all_item_nos:
            text = "\n".join(all_item_nos)
            try:
                proc = subprocess.Popen(["pbcopy"], stdin=subprocess.PIPE)
                proc.communicate(text.encode("utf-8"))
                print(f"[SUCCESS] {len(all_item_nos)}개 상품번호가 클립보드에 복사되었습니다! (앞자리 0 유지)")
            except Exception as e:
                print(f"[WARN] 클립보드 복사 실패: {e}")
        else:
            print("[WARN] 복사할 상품번호가 없습니다.")

    return results

def main():
    parser = argparse.ArgumentParser(description="QA 100일 이내 재고 연동 상품 조회 스크립트")
    parser.add_argument("--all", action="store_true", help="전체 제휴사 모든 하위 채널 일괄 조회")
    parser.add_argument("--partner", type=str, help="특정 채널명(제휴사) 조회 (예: G마켓, 옥션, 11번가, 네이버 등)")
    parser.add_argument("--channel", type=str, help="특정 하위 채널 ID 직접 조회 (예: 0000018820)")
    parser.add_argument("--list", action="store_true", help="지원하는 제휴사 및 채널 ID 목록 출력")
    parser.add_argument("--csv", type=str, help="조회 결과를 저장할 CSV 파일명 (예: recent_all_items.csv)")
    parser.add_argument("--excel", type=str, help="조회 결과를 저장할 Excel 파일명 (예: recent_all_items.xlsx)")
    parser.add_argument("--copy", action="store_true", help="조회된 모든 상품번호를 macOS 클립보드로 복사")
    
    args = parser.parse_args()
    channel_map = get_channel_map()

    if args.list:
        print_channel_list()
        return

    targets = []

    if args.channel:
        if args.channel in channel_map:
            targets.append(channel_map[args.channel])
        else:
            targets.append({
                "partner": "기타/사용자지정",
                "mainChnlId": "-",
                "name": "직접입력 채널",
                "chnlId": args.channel
            })
    elif args.partner:
        matched = False
        for group in CHANNEL_DATA:
            if args.partner.lower() in group["partner"].lower():
                matched = True
                for ch in group["channels"]:
                    targets.append({
                        "partner": group["partner"],
                        "mainChnlId": group["mainChnlId"],
                        "name": ch["name"],
                        "chnlId": ch["chnlId"]
                    })
        if not matched:
            print(f"[ERROR] 제휴사 '{args.partner}'를 찾을 수 없습니다.")
            print("사용 가능한 채널명(제휴사):", ", ".join([g["partner"] for g in CHANNEL_DATA]))
            sys.exit(1)
    elif args.all:
        for group in CHANNEL_DATA:
            for ch in group["channels"]:
                targets.append({
                    "partner": group["partner"],
                    "mainChnlId": group["mainChnlId"],
                    "name": ch["name"],
                    "chnlId": ch["chnlId"]
                })
    else:
        print("\n=== QA 연동 상품(100일 이내 재고 有) 조회 도구 ===")
        print("1. 전체 제휴사 모든 채널 일괄 조회")
        print("2. 특정 제휴사별 조회 (G마켓, 옥션, 11번가, 네이버, etc)")
        print("3. 특정 하위 채널 ID 직접 조회 (예: 0000018820)")
        print("4. 제휴사 및 채널 ID 목록 확인")
        print("0. 종료")
        
        try:
            choice = input("\n메뉴 번호를 선택하세요 [1-4]: ").strip()
        except (KeyboardInterrupt, EOFError):
            print("\n종료합니다.")
            return

        if choice == "1":
            for group in CHANNEL_DATA:
                for ch in group["channels"]:
                    targets.append({
                        "partner": group["partner"],
                        "mainChnlId": group["mainChnlId"],
                        "name": ch["name"],
                        "chnlId": ch["chnlId"]
                    })
        elif choice == "2":
            print("\n[채널명(제휴사) 선택]")
            for idx, g in enumerate(CHANNEL_DATA, 1):
                print(f"{idx}. {g['partner']}")
            p_idx = input("번호 또는 이름 입력: ").strip()
            if p_idx.isdigit() and 1 <= int(p_idx) <= len(CHANNEL_DATA):
                sel = CHANNEL_DATA[int(p_idx) - 1]
                for ch in sel["channels"]:
                    targets.append({
                        "partner": sel["partner"],
                        "mainChnlId": sel["mainChnlId"],
                        "name": ch["name"],
                        "chnlId": ch["chnlId"]
                    })
            else:
                for g in CHANNEL_DATA:
                    if p_idx.lower() in g["partner"].lower():
                        for ch in g["channels"]:
                            targets.append({
                                "partner": g["partner"],
                                "mainChnlId": g["mainChnlId"],
                                "name": ch["name"],
                                "chnlId": ch["chnlId"]
                            })
                        break
        elif choice == "3":
            inp_id = input("조회할 하위 채널 ID(CHNL_ID)를 입력하세요 (예: 0000018820): ").strip()
            if inp_id in channel_map:
                targets.append(channel_map[inp_id])
            else:
                targets.append({
                    "partner": "기타/사용자지정",
                    "mainChnlId": "-",
                    "name": "직접입력 채널",
                    "chnlId": inp_id
                })
        elif choice == "4":
            print_channel_list()
            return
        else:
            print("종료합니다.")
            return

    if not targets:
        print("[ERROR] 조회 대상 채널이 없습니다.")
        return

    csv_file = args.csv
    xlsx_file = args.excel
    if not csv_file and not xlsx_file and not (args.all or args.partner or args.channel):
        save_yn = input("결과를 파일로 저장하시겠습니까? (1: CSV, 2: Excel, 3: 둘 다, N: 저장안함): ").strip().lower()
        now_str = datetime.now().strftime("%Y%m%d_%H%M%S")
        if save_yn in ["1", "csv"]:
            csv_file = f"qa_linkage_items_{now_str}.csv"
        elif save_yn in ["2", "excel", "xlsx"]:
            xlsx_file = f"qa_linkage_items_{now_str}.xlsx"
        elif save_yn in ["3", "both"]:
            csv_file = f"qa_linkage_items_{now_str}.csv"
            xlsx_file = f"qa_linkage_items_{now_str}.xlsx"

    run_query(targets, save_csv=csv_file, save_xlsx=xlsx_file, copy_to_clipboard=args.copy)

if __name__ == "__main__":
    main()
