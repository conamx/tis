# -*- coding: utf-8 -*-
"""
테더링 IP 변경 프로그램 (Android + ADB)

PC 를 안드로이드 폰의 USB 테더링(또는 핫스팟)으로 인터넷에 연결한 상태에서
폰의 비행기모드 / 모바일데이터를 껐다 켜서 통신사로부터 새 공인 IP 를 받는다.

필요한 것
  - Windows/Mac/Linux + Python 3.8 이상 (표준 라이브러리만 사용)
  - adb (Android platform-tools) : 이 파일과 같은 폴더의 platform-tools 폴더 안에 두거나 PATH 에 등록
  - 폰: 개발자 옵션 > USB 디버깅 ON, PC 연결 시 "USB 디버깅 허용" 승인

실행
  python tethering_ip_changer.py          -> GUI 실행
  python tethering_ip_changer.py --change -> GUI 없이 IP 1회 변경 (명령줄)
"""

import argparse
import os
import shutil
import subprocess
import sys
import threading
import time
import urllib.request

IP_SERVICES = [
    "https://api.ipify.org",
    "https://ifconfig.me/ip",
    "https://icanhazip.com",
    "https://checkip.amazonaws.com",
]

MODE_AIRPLANE = "airplane"
MODE_DATA = "data"


# ---------------------------------------------------------------- 공통 기능

def find_adb():
    base = os.path.dirname(os.path.abspath(sys.argv[0]))
    names = ("adb.exe", "adb") if os.name == "nt" else ("adb",)
    for name in names:
        for path in (os.path.join(base, "platform-tools", name), os.path.join(base, name)):
            if os.path.isfile(path):
                return path
    return shutil.which("adb")


def run_adb(adb, args, timeout=15):
    """adb 명령 실행 -> (returncode, 출력문자열)"""
    flags = 0x08000000 if os.name == "nt" else 0  # CREATE_NO_WINDOW: 콘솔창 깜빡임 방지
    try:
        p = subprocess.run([adb] + args, capture_output=True, text=True,
                           timeout=timeout, creationflags=flags,
                           encoding="utf-8", errors="replace")
        return p.returncode, (p.stdout + p.stderr).strip()
    except subprocess.TimeoutExpired:
        return -1, "시간 초과"
    except Exception as e:  # noqa: BLE001
        return -1, str(e)


def get_public_ip(timeout=4):
    for url in IP_SERVICES:
        try:
            req = urllib.request.Request(url, headers={"User-Agent": "curl/8"})
            with urllib.request.urlopen(req, timeout=timeout) as r:
                ip = r.read().decode().strip()
                if ip and len(ip) <= 45:
                    return ip
        except Exception:  # noqa: BLE001
            continue
    return None


def get_device(adb):
    """연결된 기기 상태 문자열 반환 (없으면 None)"""
    code, out = run_adb(adb, ["devices"])
    if code != 0:
        return None
    for line in out.splitlines()[1:]:
        parts = line.split()
        if len(parts) >= 2:
            return f"{parts[0]} ({parts[1]})"
    return None


class IpChanger:
    def __init__(self, adb, log=print):
        self.adb = adb
        self.log = log
        self.stop_flag = False

    def shell(self, *cmd, timeout=15):
        return run_adb(self.adb, ["shell"] + list(cmd), timeout=timeout)

    # --- 폰 제어
    def airplane(self, on):
        state = "enable" if on else "disable"
        code, out = self.shell("cmd", "connectivity", "airplane-mode", state)
        if code != 0 or "Unknown" in out or "Exception" in out:
            # 안드로이드 9 이하 대체 방식
            self.shell("settings", "put", "global", "airplane_mode_on", "1" if on else "0")
            self.shell("am", "broadcast", "-a", "android.intent.action.AIRPLANE_MODE",
                       "--ez", "state", "true" if on else "false")

    def mobile_data(self, on):
        self.shell("svc", "data", "enable" if on else "disable")

    def restore_usb_tethering(self):
        self.log("  USB 테더링 다시 켜는 중...")
        run_adb(self.adb, ["wait-for-device"], timeout=20)
        code, out = self.shell("svc", "usb", "setFunctions", "rndis")
        if code != 0:
            self.log(f"  (자동 재연결 실패: {out}) 폰에서 USB 테더링을 직접 켜주세요.")

    def wait_internet(self, max_wait, usb_tether):
        """인터넷이 다시 될 때까지 대기 후 IP 반환"""
        start = time.time()
        restored = False
        while time.time() - start < max_wait and not self.stop_flag:
            ip = get_public_ip(timeout=3)
            if ip:
                return ip
            if usb_tether and not restored and time.time() - start > 12:
                self.restore_usb_tethering()
                restored = True
            time.sleep(2)
        return None

    # --- 메인 동작
    def change_once(self, mode, off_seconds, usb_tether):
        if mode == MODE_AIRPLANE:
            self.log(f"  비행기모드 ON ({off_seconds}초 대기)")
            self.airplane(True)
            time.sleep(off_seconds)
            self.log("  비행기모드 OFF")
            self.airplane(False)
        else:
            self.log(f"  모바일데이터 OFF ({off_seconds}초 대기)")
            self.mobile_data(False)
            time.sleep(off_seconds)
            self.log("  모바일데이터 ON")
            self.mobile_data(True)
        self.log("  인터넷 재연결 대기 중...")
        return self.wait_internet(60, usb_tether)

    def change(self, mode=MODE_AIRPLANE, off_seconds=3, retries=3, usb_tether=True):
        self.stop_flag = False
        if not get_device(self.adb):
            self.log("[오류] 연결된 폰이 없습니다. USB 디버깅 / 케이블을 확인하세요.")
            return None
        old_ip = get_public_ip()
        self.log(f"현재 IP: {old_ip or '확인 불가'}")
        for attempt in range(1, retries + 1):
            if self.stop_flag:
                break
            self.log(f"[시도 {attempt}/{retries}]")
            new_ip = self.change_once(mode, off_seconds, usb_tether)
            if not new_ip:
                self.log("  인터넷 연결 실패 (테더링 상태를 확인하세요)")
                continue
            if new_ip != old_ip:
                self.log(f"[성공] {old_ip} -> {new_ip}")
                return new_ip
            self.log(f"  IP 동일 ({new_ip}), 대기시간을 늘려 다시 시도")
            off_seconds += 2
        self.log("[실패] IP 가 바뀌지 않았습니다.")
        return None


# ---------------------------------------------------------------- GUI

def run_gui(adb):
    import tkinter as tk
    from tkinter import ttk, scrolledtext

    root = tk.Tk()
    root.title("테더링 IP 변경기")
    root.geometry("520x560")
    root.minsize(460, 480)

    lines = []

    def log(msg):
        stamp = time.strftime("%H:%M:%S")
        root.after(0, lambda: (box.insert("end", f"[{stamp}] {msg}\n"), box.see("end")))
        lines.append(msg)

    changer = IpChanger(adb, log)
    busy = {"v": False}
    auto_job = {"id": None}

    # 상단 정보
    top = ttk.Frame(root, padding=10)
    top.pack(fill="x")
    ip_var = tk.StringVar(value="-")
    dev_var = tk.StringVar(value="-")
    ttk.Label(top, text="현재 IP", width=10).grid(row=0, column=0, sticky="w")
    ttk.Label(top, textvariable=ip_var, font=("Consolas", 16, "bold")).grid(row=0, column=1, sticky="w")
    ttk.Label(top, text="연결 기기", width=10).grid(row=1, column=0, sticky="w")
    ttk.Label(top, textvariable=dev_var).grid(row=1, column=1, sticky="w")

    # 설정
    opt = ttk.LabelFrame(root, text="설정", padding=10)
    opt.pack(fill="x", padx=10)
    mode_var = tk.StringVar(value=MODE_AIRPLANE)
    ttk.Radiobutton(opt, text="비행기모드 껐다켜기 (권장)", variable=mode_var,
                    value=MODE_AIRPLANE).grid(row=0, column=0, columnspan=2, sticky="w")
    ttk.Radiobutton(opt, text="모바일데이터 껐다켜기", variable=mode_var,
                    value=MODE_DATA).grid(row=1, column=0, columnspan=2, sticky="w")

    off_var = tk.IntVar(value=3)
    retry_var = tk.IntVar(value=3)
    auto_var = tk.IntVar(value=0)
    usb_var = tk.BooleanVar(value=True)
    ttk.Label(opt, text="끈 상태 유지(초)").grid(row=2, column=0, sticky="w", pady=2)
    ttk.Spinbox(opt, from_=1, to=60, textvariable=off_var, width=6).grid(row=2, column=1, sticky="w")
    ttk.Label(opt, text="최대 재시도").grid(row=3, column=0, sticky="w", pady=2)
    ttk.Spinbox(opt, from_=1, to=10, textvariable=retry_var, width=6).grid(row=3, column=1, sticky="w")
    ttk.Label(opt, text="자동 반복(분, 0=끔)").grid(row=4, column=0, sticky="w", pady=2)
    ttk.Spinbox(opt, from_=0, to=1440, textvariable=auto_var, width=6).grid(row=4, column=1, sticky="w")
    ttk.Checkbutton(opt, text="USB 테더링 끊기면 자동 재연결",
                    variable=usb_var).grid(row=5, column=0, columnspan=2, sticky="w")

    # 버튼
    btns = ttk.Frame(root, padding=10)
    btns.pack(fill="x")

    def refresh():
        def work():
            dev = get_device(adb) or "연결 안 됨"
            ip = get_public_ip() or "확인 불가"
            root.after(0, lambda: (dev_var.set(dev), ip_var.set(ip)))
        threading.Thread(target=work, daemon=True).start()

    def schedule_auto():
        if auto_job["id"]:
            root.after_cancel(auto_job["id"])
            auto_job["id"] = None
        try:
            minutes = int(auto_var.get())
        except (tk.TclError, ValueError):
            minutes = 0
        if minutes > 0:
            auto_job["id"] = root.after(minutes * 60_000, lambda: do_change(auto=True))
            log(f"{minutes}분 후 자동 변경 예약")

    def do_change(auto=False):
        if busy["v"]:
            return
        busy["v"] = True
        change_btn.config(state="disabled")

        def work():
            try:
                ip = changer.change(mode_var.get(), int(off_var.get()),
                                    int(retry_var.get()), usb_var.get())
                if ip:
                    root.after(0, lambda: ip_var.set(ip))
            except Exception as e:  # noqa: BLE001
                log(f"[오류] {e}")
            finally:
                busy["v"] = False
                root.after(0, lambda: change_btn.config(state="normal"))
                root.after(0, refresh)
                root.after(0, schedule_auto)
        threading.Thread(target=work, daemon=True).start()

    def stop():
        changer.stop_flag = True
        if auto_job["id"]:
            root.after_cancel(auto_job["id"])
            auto_job["id"] = None
        log("중지됨 (자동 반복 해제)")

    change_btn = ttk.Button(btns, text="IP 변경", command=do_change)
    change_btn.pack(side="left", expand=True, fill="x", padx=2, ipady=6)
    ttk.Button(btns, text="IP / 기기 확인", command=refresh).pack(side="left", expand=True, fill="x", padx=2, ipady=6)
    ttk.Button(btns, text="중지", command=stop).pack(side="left", expand=True, fill="x", padx=2, ipady=6)

    box = scrolledtext.ScrolledText(root, height=12, font=("Consolas", 9))
    box.pack(fill="both", expand=True, padx=10, pady=(0, 10))

    root.bind("<F5>", lambda e: do_change())
    log(f"adb: {adb}")
    log("단축키 F5 = IP 변경")
    refresh()
    root.mainloop()


# ---------------------------------------------------------------- 진입점

def main():
    ap = argparse.ArgumentParser(description="테더링 IP 변경 프로그램")
    ap.add_argument("--change", action="store_true", help="GUI 없이 IP 1회 변경")
    ap.add_argument("--mode", choices=[MODE_AIRPLANE, MODE_DATA], default=MODE_AIRPLANE)
    ap.add_argument("--off", type=int, default=3, help="끈 상태 유지 시간(초)")
    ap.add_argument("--retries", type=int, default=3)
    ap.add_argument("--no-usb", action="store_true", help="USB 테더링 자동 재연결 안 함")
    args = ap.parse_args()

    adb = find_adb()
    if not adb:
        msg = ("adb 를 찾을 수 없습니다.\n"
               "Android platform-tools 를 내려받아 이 프로그램 폴더에\n"
               "'platform-tools' 폴더째로 넣어주세요.")
        if args.change:
            print(msg)
        else:
            import tkinter as tk
            from tkinter import messagebox
            tk.Tk().withdraw()
            messagebox.showerror("adb 없음", msg)
        sys.exit(1)

    if args.change:
        ok = IpChanger(adb).change(args.mode, args.off, args.retries, not args.no_usb)
        sys.exit(0 if ok else 1)
    run_gui(adb)


if __name__ == "__main__":
    main()
