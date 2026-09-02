# -*- coding: utf-8 -*-
"""
ORG_FULL_NAME · 协同录入系统  客户端连接启动器（控制台版）
================================================
局域网自动配网：用户输入财务电脑服务器显示的「配网数字」，
本程序通过 UDP 广播自动发现服务器地址并打开浏览器，无需手动配置任何网络参数。
（采用控制台交互，零额外依赖，便于随安装包直接发布。）
"""
import socket
import sys
import time
import webbrowser

DISCOVER_PORT = 39997
MAGIC = "COOP_DISCOVER:"
RESP = "COOP_SERVER:"


def discover(code, timeout=6.0):
    """向局域网广播配网数字，返回 (ip, port) 或 None。"""
    sock = socket.socket(socket.AF_INET, socket.SOCK_DGRAM)
    sock.setsockopt(socket.SOL_SOCKET, socket.SO_BROADCAST, 1)
    sock.settimeout(1.0)
    end = time.monotonic() + timeout
    while time.monotonic() < end:
        try:
            sock.sendto((MAGIC + code).encode("utf-8"), ("<broadcast>", DISCOVER_PORT))
        except Exception:
            pass
        try:
            data, _addr = sock.recvfrom(1024)
            msg = data.decode("utf-8", "ignore").strip()
            if msg.startswith(RESP):
                rest = msg[len(RESP):]
                ip, port = rest.split(":", 1)
                return ip, int(port)
        except socket.timeout:
            continue
        except Exception:
            continue
    return None


def main():
    print("=" * 52)
    print("   协同录入系统 · 连接服务器（局域网自动配网）")
    print("=" * 52)
    print("请输入财务电脑服务器窗口显示的「配网数字」，按回车连接。")
    print("（直接回车可重新输入；按 Ctrl+C 退出）")
    while True:
        try:
            code = input("\n配网数字> ").strip()
        except (EOFError, KeyboardInterrupt):
            print("\n已退出。")
            return
        if not code:
            continue
        print("正在局域网内搜索服务器 ...")
        res = discover(code)
        if res:
            ip, port = res
            url = "http://%s:%d" % (ip, port)
            print("已找到服务器：%s" % url)
            print("正在打开浏览器 ...")
            try:
                webbrowser.open(url)
            except Exception:
                pass
            print("若浏览器未自动打开，请手动访问：%s" % url)
            try:
                input("按回车退出 ...")
            except Exception:
                pass
            return
        print("未找到匹配的服务器。请确认：")
        print("  1. 财务电脑已启动「协同录入系统」服务器；")
        print("  2. 输入的配网数字与服务器窗口显示的一致；")
        print("  3. 防火墙已放行 UDP 39997 端口。")


if __name__ == "__main__":
    main()
