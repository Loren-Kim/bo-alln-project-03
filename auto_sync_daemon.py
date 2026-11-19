#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
선연동 자동 동기화 데몬 (auto_sync_daemon.py)
- 평일 업무시간(09:00 ~ 19:00) 1시간마다 주기적 동기화 실행
- Downloads 폴더 내 '선연동_*.csv' 다운로드 즉시 감지하여 자동 동기화 실행 (5초 디바운스)
- 실행: python3 auto_sync_daemon.py
- 상태: python3 auto_sync_daemon.py --status
- 종료: python3 auto_sync_daemon.py --stop
"""

import os
import sys
import time
import glob
import signal
import subprocess
from datetime import datetime

WORKSPACE_DIR = os.path.dirname(os.path.abspath(__file__))
DOWNLOADS_DIR = os.path.expanduser("~/Downloads")
PID_FILE = os.path.join(WORKSPACE_DIR, "daemon.pid")
LOG_FILE = os.path.join(WORKSPACE_DIR, "sync.log")
SYNC_SCRIPT = os.path.join(WORKSPACE_DIR, "sync_all.py")

WORK_START_HOUR = 9
WORK_END_HOUR = 19
CHECK_INTERVAL_SEC = 5     # 폴더 감시 주기 (초)
BATCH_INTERVAL_SEC = 3600  # 정기 배치 주기 (1시간 = 3600초)

def log(msg):
    ts = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    line = f"[DAEMON {ts}] {msg}"
    print(line)
    try:
        with open(LOG_FILE, "a", encoding="utf-8") as f:
            f.write(line + "\n")
    except Exception:
        pass

def is_work_hours():
    now = datetime.now()
    # 월요일=0, 금요일=4, 토요일=5, 일요일=6
    if now.weekday() >= 5:
        return False
    return WORK_START_HOUR <= now.hour < WORK_END_HOUR

def run_sync(reason="배치"):
    log(f"⚡ 동기화 실행 트리거 (사유: {reason})")
    try:
        res = subprocess.run([sys.executable, SYNC_SCRIPT], cwd=WORKSPACE_DIR, capture_output=True, text=True)
        if res.returncode == 0:
            log(f"✅ 동기화 완료 ({reason})")
        else:
            log(f"❌ 동기화 실패 ({reason}): {res.stderr[:200]}")
    except Exception as e:
        log(f"❌ 동기화 오류 ({reason}): {e}")

def get_prelinkage_files_snapshot():
    """Downloads 폴더의 선연동 CSV 파일들의 최신 수정시각(mtime) 스냅샷 생성"""
    files = glob.glob(os.path.join(DOWNLOADS_DIR, "선연동_*.csv"))
    snapshot = {}
    for f in files:
        try:
            snapshot[f] = os.path.getmtime(f)
        except OSError:
            pass
    return snapshot

def status():
    if not os.path.exists(PID_FILE):
        print("[-] 자동 동기화 데몬이 실행 중이지 않습니다.")
        return False
    try:
        with open(PID_FILE, "r") as f:
            pid = int(f.read().strip())
        os.kill(pid, 0)
        print(f"[+] 자동 동기화 데몬이 정상 작동 중입니다 (PID: {pid}).")
        return True
    except (OSError, ValueError):
        print("[-] 저장된 PID 프로세스가 존재하지 않습니다. 정리합니다.")
        try:
            os.remove(PID_FILE)
        except OSError:
            pass
        return False

def stop():
    if not os.path.exists(PID_FILE):
        print("[-] 실행 중인 데몬 PID 파일이 없습니다.")
        return
    try:
        with open(PID_FILE, "r") as f:
            pid = int(f.read().strip())
        os.kill(pid, signal.SIGTERM)
        print(f"[+] 데몬 프로세스(PID: {pid})를 정상 종료했습니다.")
    except Exception as e:
        print(f"[-] 데몬 종료 실패: {e}")
    finally:
        if os.path.exists(PID_FILE):
            os.remove(PID_FILE)

def run_daemon():
    # 중복 실행 방지
    if os.path.exists(PID_FILE):
        try:
            with open(PID_FILE, "r") as f:
                old_pid = int(f.read().strip())
            os.kill(old_pid, 0)
            log(f"이미 실행 중인 데몬 프로세스가 있습니다 (PID: {old_pid}).")
            print(f"이미 데몬이 실행 중입니다 (PID: {old_pid}).")
            return
        except (OSError, ValueError):
            pass

    with open(PID_FILE, "w") as f:
        f.write(str(os.getpid()))

    log(f"🚀 자동 동기화 데몬 시작 (PID: {os.getpid()})")
    log(f"설정: 평일 {WORK_START_HOUR:02d}:00 ~ {WORK_END_HOUR:02d}:00 (1시간 배치) + Downloads 감시")

    last_batch_time = 0
    known_snapshot = get_prelinkage_files_snapshot()

    try:
        while True:
            now_ts = time.time()

            # 1. 신규/수정 선연동 CSV 파일 감시
            current_snapshot = get_prelinkage_files_snapshot()
            has_new_or_modified = False
            for fpath, mtime in current_snapshot.items():
                if fpath not in known_snapshot or mtime > known_snapshot[fpath] + 0.1:
                    has_new_or_modified = True
                    fname = os.path.basename(fpath)
                    log(f"📂 새로운 선연동 파일 다운로드 감지: {fname}")
                    break

            if has_new_or_modified:
                time.sleep(3)  # 다운로드 완료를 위한 3초 버퍼
                known_snapshot = get_prelinkage_files_snapshot()
                run_sync(reason="새 파일 다운로드 감지")
                last_batch_time = now_ts

            # 2. 평일 업무시간 1시간 주기 정기 배치
            if is_work_hours():
                if now_ts - last_batch_time >= BATCH_INTERVAL_SEC:
                    run_sync(reason="1시간 정기 배치")
                    last_batch_time = now_ts

            time.sleep(CHECK_INTERVAL_SEC)

    except KeyboardInterrupt:
        log("사용자에 의해 데몬이 중단되었습니다.")
    finally:
        if os.path.exists(PID_FILE):
            os.remove(PID_FILE)
        log("자동 동기화 데몬이 종료되었습니다.")

def main():
    if len(sys.argv) > 1:
        arg = sys.argv[1].lower()
        if arg in ["--status", "-s", "status"]:
            status()
            return
        elif arg in ["--stop", "-k", "stop", "kill"]:
            stop()
            return

    run_daemon()

if __name__ == "__main__":
    main()
