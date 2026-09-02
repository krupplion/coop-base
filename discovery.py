# -*- coding: utf-8 -*-
"""
局域网自动配网 · 服务端发现响应器（coop-base 通用分身）
原理：客户端向局域网广播 "COOP_DISCOVER:<配网数字>"，
本响应器监听 UDP 端口（默认 39997，可用 COOP_DISCOVERY_PORT 覆盖），数字匹配即回复 "COOP_SERVER:<IP>:<端口>"，
客户端收到后自动得到服务器地址——无需手动配置任何网络参数。
"""
import os
import socket
import threading

DISCOVER_PORT = int(os.environ.get("COOP_DISCOVERY_PORT", "39997"))


def _local_ips():
    ips = []
    try:
        s = socket.socket(socket.AF_INET, socket.SOCK_DGRAM)
        s.connect(("8.8.8.8", 80))
        ips.append(s.getsockname()[0])
        s.close()
    except Exception:
        pass
    try:
        for info in socket.getaddrinfo(socket.gethostname(), None, socket.AF_INET):
            ip = info[4][0]
            if not ip.startswith("127.") and ip not in ips:
                ips.append(ip)
    except Exception:
        pass
    return ips or ["127.0.0.1"]


def start_responder(pair_code, server_port):
    """后台线程运行：响应配网广播"""
    def loop():
        sock = socket.socket(socket.AF_INET, socket.SOCK_DGRAM)
        sock.setsockopt(socket.SOL_SOCKET, socket.SO_REUSEADDR, 1)
        sock.bind(("", DISCOVER_PORT))
        print("[配网] 发现服务已启动，UDP 端口 %d，等待客户端配网..." % DISCOVER_PORT)
        while True:
            try:
                data, addr = sock.recvfrom(1024)
                msg = data.decode("utf-8", "ignore").strip()
                if msg == "COOP_DISCOVER:%s" % pair_code:
                    reply = "COOP_SERVER:%s:%d" % (_local_ips()[0], server_port)
                    sock.sendto(reply.encode(), addr)
                    print("[配网] 已响应客户端 %s" % addr[0])
            except Exception:
                continue

    t = threading.Thread(target=loop, daemon=True)
    t.start()
    return t
