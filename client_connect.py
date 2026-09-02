# -*- coding: utf-8 -*-
"""
局域网自动配网 · 客户端连接工具
用法：在其他 3 位人员的电脑上运行本脚本（双击或 python client_connect.py），
输入财务部分发的 6 位配网数字，自动发现服务器并打开浏览器进入系统。
无需手动配置 IP、端口等任何网络参数。
"""
import socket
import sys
import webbrowser

DISCOVER_PORT = 39997


def discover(pair_code, timeout=5):
    sock = socket.socket(socket.AF_INET, socket.SOCK_DGRAM)
    sock.setsockopt(socket.SOL_SOCKET, socket.SO_BROADCAST, 1)
    sock.settimeout(timeout)
    msg = ("COOP_DISCOVER:%s" % pair_code).encode()
    # 多播几个常见广播地址，兼容不同网段
    for bcast in ("255.255.255.255",):
        try:
            sock.sendto(msg, (bcast, DISCOVER_PORT))
        except Exception:
            pass
    try:
        data, _ = sock.recvfrom(1024)
        text = data.decode()
        if text.startswith("COOP_SERVER:"):
            _, ip, port = text.strip().split(":")
            return ip, int(port)
    except socket.timeout:
        return None
    finally:
        sock.close()
    return None


def main():
    print("=" * 50)
    print("  ORG_FULL_NAME · 协同录入系统 客户端连接")
    print("=" * 50)
    pair = input("请输入 6 位配网数字（向财务部索取）: ").strip()
    if not pair:
        print("配网数字不能为空")
        sys.exit(1)
    print("正在局域网内搜索服务器...")
    result = discover(pair)
    if not result:
        # 广播可能被交换机拦截时，重试一次
        result = discover(pair, timeout=8)
    if not result:
        print("未发现服务器。请确认：")
        print("  1. 财务部电脑上的服务已启动；")
        print("  2. 本机与服务器在同一局域网；")
        print("  3. 配网数字输入正确。")
        sys.exit(2)
    ip, port = result
    url = "http://%s:%d" % (ip, port)
    print("已找到服务器：%s" % url)
    print("正在打开浏览器...")
    webbrowser.open(url)
    print("如浏览器未自动打开，请手动访问：%s" % url)


if __name__ == "__main__":
    main()
