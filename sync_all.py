#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
선연동 데이터 통합 동기화 & GitHub Pages 자동 배포 스크립트 (sync_all.py)
1. Downloads 폴더 내 '선연동_*.csv' 파일들을 자동 스캔하여 최신 제휴상품번호 매핑 수집
2. fetch_recent_items.py, index.html, qa_item_finder.html의 BUILTIN_ALLNCO_MAP 내장 캐시 갱신
3. fetch_recent_items.py 실행 (QA BO 현황 크롤링 -> recent_all_items.csv & recent_all_items.xlsx 생성)
4. Git 변경사항 감지 시 자동 commit & push (GitHub Pages 즉시 배포)
"""

import os
import sys
import glob
import re
import json
import subprocess
from datetime import datetime

WORKSPACE_DIR = os.path.dirname(os.path.abspath(__file__))
RESULT_DIR = os.path.expanduser("~/Documents/제휴상품/AllnProductResult")
DOWNLOADS_DIR = os.path.expanduser("~/Downloads")
LOG_FILE = os.path.join(WORKSPACE_DIR, "sync.log")

CHANNEL_NAME_TO_ID = {
    "지마켓(백화점)": "0000015782",
    "옥션(백화점)": "0000015779",
    "지마켓(신세계몰)": "0000015781",
    "옥션(신세계몰)": "0000015778",
    "프리미엄아울렛(지마켓)": "0000017105",
    "프리미엄아울렛(옥션)": "0000017106",
    "11번가(신세계몰)": "0000015959",
    "11번가(신세계백화점)": "0000015960",
    "11번가(백화점)": "0000015960",
    "11번가(프리미엄아울렛)": "0000017104",
    "11번가(아울렛)": "0000017104",
    "W컨셉(신세계백화점)": "0000019090",
    "W컨셉(백화점)": "0000019090",
    "W컨셉(신세계몰)": "0000019091",
    "W컨셉(몰)": "0000019091",
    "라이브쇼핑(백화점)": "0000019623",
    "라이브쇼핑(몰)": "0000019624",
    "현대Hmall": "0000016083",
    "현대Hmall(전문관)": "0000018200",
    "카카오(선물하기)": "0000018380",
    "카카오(백화점)": "0000018820",
    "카카오선물하기(몰)": "0000020717",
    "롯데ON(백화점)": "0000020540",
    "티몬(신세계몰)": "0000021954",
    "당근(백화점)": "0000023802"
}

def log(msg):
    ts = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    line = f"[{ts}] {msg}"
    print(line)
    try:
        with open(LOG_FILE, "a", encoding="utf-8") as f:
            f.write(line + "\n")
    except Exception:
        pass

def scan_downloads_for_prelinkage():
    # 1순위: ~/Documents/제휴상품/AllnProductResult, 2순위: ~/Downloads
    csv_files = []
    if os.path.exists(RESULT_DIR):
        csv_files.extend(glob.glob(os.path.join(RESULT_DIR, "선연동_*.csv")))
    if os.path.exists(DOWNLOADS_DIR):
        csv_files.extend(glob.glob(os.path.join(DOWNLOADS_DIR, "선연동_*.csv")))
    
    # 중복 제거
    csv_files = sorted(list(set(csv_files)))
    log(f"선연동 CSV 검색 결과: 총 {len(csv_files)}개 발견 (경로: {RESULT_DIR})")
    
    mapping = {}
    for filepath in sorted(csv_files):
        fname = os.path.basename(filepath)
        chnl_id = None
        for name, cid in CHANNEL_NAME_TO_ID.items():
            if name in fname:
                chnl_id = cid
                break
        
        # 파일 내용 읽기 (utf-8 또는 cp949)
        content = None
        for enc in ['utf-8', 'cp949', 'euc-kr']:
            try:
                with open(filepath, 'r', encoding=enc) as f:
                    content = f.read()
                break
            except Exception:
                continue
        
        if not content:
            continue
            
        lines = content.splitlines()
        if not lines:
            continue
            
        # 1행 주석에서 채널 재확인
        if not chnl_id and len(lines) > 0:
            first_line = lines[0]
            for name, cid in CHANNEL_NAME_TO_ID.items():
                if name in first_line:
                    chnl_id = cid
                    break
                    
        # 헤더 찾기
        ssg_col, alln_col, start_row = -1, -1, 0
        for i in range(min(10, len(lines))):
            cols = [c.strip().strip('"').strip("'") for c in lines[i].split(',')]
            for idx, c in enumerate(cols):
                if any(k in c for k in ["상품ID", "SSG상품ID", "ITEM_ID"]):
                    ssg_col = idx
                if any(k in c for k in ["제휴상품ID", "제휴사상품ID", "ALLNCO_ITEM_ID"]):
                    alln_col = idx
            if ssg_col != -1 and alln_col != -1:
                start_row = i + 1
                break
                
        if ssg_col == -1 or alln_col == -1:
            continue
            
        for line in lines[start_row:]:
            if not line.strip():
                continue
            cols = [c.strip().replace('="', '').replace('"', '').replace("'", "") for c in line.split(',')]
            if len(cols) > max(ssg_col, alln_col):
                ssg_id = cols[ssg_col].strip()
                alln_id = cols[alln_col].strip()
                if ssg_id and alln_id and alln_id.lower() not in ["null", "-", "", "none"]:
                    if chnl_id:
                        mapping[f"{chnl_id}_{ssg_id}"] = alln_id
                    mapping[ssg_id] = alln_id
                    
    log(f"총 {len(mapping)}건의 선연동 제휴상품번호 매핑 수집 완료")
    return mapping

def update_codebase_mapping(new_mapping):
    if not new_mapping:
        log("추가된 매핑이 없습니다.")
        return False
        
    fetch_py_path = os.path.join(WORKSPACE_DIR, "fetch_recent_items.py")
    index_html_path = os.path.join(WORKSPACE_DIR, "index.html")
    qa_html_path = os.path.join(WORKSPACE_DIR, "qa_item_finder.html")
    
    # 1. fetch_recent_items.py 읽기 및 기존 매핑과 병합
    with open(fetch_py_path, 'r', encoding='utf-8') as f:
        fetch_content = f.read()
        
    match = re.search(r'BUILTIN_ALLNCO_MAP\s*=\s*(\{.*?\})', fetch_content, re.DOTALL)
    existing_map = {}
    if match:
        try:
            existing_map = json.loads(match.group(1))
        except Exception:
            pass
            
    existing_map.update(new_mapping)
    log(f"통합된 BUILTIN_ALLNCO_MAP 항목 수: {len(existing_map)}건")
    map_json_str = json.dumps(existing_map, ensure_ascii=False)
    
    # fetch_recent_items.py 갱신
    new_fetch_content = re.sub(
        r'BUILTIN_ALLNCO_MAP\s*=\s*\{.*?\}',
        f'BUILTIN_ALLNCO_MAP = {map_json_str}',
        fetch_content,
        flags=re.DOTALL
    )
    with open(fetch_py_path, 'w', encoding='utf-8') as f:
        f.write(new_fetch_content)
        
    # index.html 갱신
    with open(index_html_path, 'r', encoding='utf-8') as f:
        index_content = f.read()
    new_index_content = re.sub(
        r'const BUILTIN_ALLNCO_MAP\s*=\s*\{.*?\};',
        f'const BUILTIN_ALLNCO_MAP = {map_json_str};',
        index_content,
        flags=re.DOTALL
    )
    with open(index_html_path, 'w', encoding='utf-8') as f:
        f.write(new_index_content)
        
    # qa_item_finder.html 갱신 (100% 동일 동기화)
    with open(qa_html_path, 'w', encoding='utf-8') as f:
        f.write(new_index_content)
        
    log("fetch_recent_items.py, index.html, qa_item_finder.html 내장 캐시 갱신 완료")
    return True

def run_fetch_script():
    log("최신 데이터 크롤링 시작 (fetch_recent_items.py --all --csv recent_all_items.csv --excel recent_all_items.xlsx)...")
    res = subprocess.run(
        [
            sys.executable, 
            os.path.join(WORKSPACE_DIR, "fetch_recent_items.py"),
            "--all",
            "--csv", "recent_all_items.csv",
            "--excel", "recent_all_items.xlsx"
        ],
        cwd=WORKSPACE_DIR,
        capture_output=True,
        text=True
    )
    if res.returncode == 0:
        log("데이터 크롤링 및 엑셀/CSV 생성 성공!")
    else:
        log(f"데이터 크롤링 실패 (코드 {res.returncode}): {res.stderr}")
        return False
    return True

def git_commit_and_push():
    log("Git 변경사항 검사 중...")
    status_res = subprocess.run(["git", "status", "--porcelain"], cwd=WORKSPACE_DIR, capture_output=True, text=True)
    changes = status_res.stdout.strip()
    
    if changes:
        log(f"Git 변경 파일 감지:\n{changes}")
        # Git add (웹 대시보드 및 스크립트 코드)
        subprocess.run(["git", "add", "index.html", "qa_item_finder.html", "fetch_recent_items.py", 
                        "auto_sync_daemon.py", "README.md", "sync_all.py", ".gitignore"], 
                       cwd=WORKSPACE_DIR, check=True)
                       
        now_str = datetime.now().strftime("%Y-%m-%d %H:%M")
        commit_msg = f"auto: sync partner linkage & prelinkage data [{now_str}]"
        
        commit_res = subprocess.run(["git", "commit", "-m", commit_msg], cwd=WORKSPACE_DIR, capture_output=True, text=True)
        log(f"Git commit 완료: {commit_res.stdout.strip()}")
    else:
        log("새로운 파일 변경사항이 없습니다.")

    # Git push -u origin main 시도
    log("Git push -u origin main 실행 중...")
    env = os.environ.copy()
    env["GIT_TERMINAL_PROMPT"] = "0"  # 무한 대기 방지
    
    push_res = subprocess.run(["git", "push", "-u", "origin", "main"], cwd=WORKSPACE_DIR, capture_output=True, text=True, env=env)
    if push_res.returncode == 0:
        log("Git push 성공! GitHub Pages에 최신 상태가 자동 배포됩니다.")
        return True
    else:
        err_msg = push_res.stderr.strip()
        log(f"Git push 실패: {err_msg}")
        if "Authentication failed" in err_msg or "could not read Username" in err_msg or "terminal prompts disabled" in err_msg:
            log("⚠️ [안내] GitHub 로그인 인증(PAT) 정보가 로컬 키체인에 등록되지 않았습니다.")
            log("👉 터미널에서 1회 'git push -u origin main'을 실행하여 GitHub 계정(토큰)으로 로그인하시면 Mac 키체인에 저장되어 다음부터 완전 자동 푸시됩니다.")
        return False

def main():
    log("========== 제휴 연동 데이터 통합 동기화 시작 ==========")
    mapping = scan_downloads_for_prelinkage()
    update_codebase_mapping(mapping)
    success = run_fetch_script()
    if success:
        git_commit_and_push()
    log("========== 제휴 연동 데이터 통합 동기화 종료 ==========\n")

if __name__ == "__main__":
    main()
