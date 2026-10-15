# -*- coding: utf-8 -*-
"""纯 Python 生成 icon.ico（32x32 32bpp）：深蓝底 + 表格格点 + 蓝色渐变角标。"""
import struct

W = H = 32
BG = (15, 23, 42, 255)        # #0f172a
GRID = (96, 165, 250, 255)    # #60a5fa 蓝
WHITE = (226, 232, 240, 255)  # 浅灰白

px = [[BG for _ in range(W)] for _ in range(H)]

# 外圈圆角：四角挖透明
for y in range(H):
    for x in range(W):
        if (x < 3 and y < 3) or (x >= W - 3 and y < 3) or \
           (x < 3 and y >= H - 3) or (x >= W - 3 and y >= H - 3):
            if (x, y) in ((0, 0), (1, 0), (0, 1), (2, 0), (0, 2), (1, 1)) or \
               x + y <= 2 or (W - 1 - x) + y <= 2 or x + (H - 1 - y) <= 2 or \
               (W - 1 - x) + (H - 1 - y) <= 2:
                px[y][x] = (0, 0, 0, 0)

# 表格网格线（第 8、16、24 行/列）
for i in range(4, W - 4):
    for line in (10, 21):
        px[line][i] = GRID
        px[i][line] = GRID

# 三个数据格点 + 一个高亮格点（第一格填蓝）
for cy, cx, col in ((6, 6, GRID), (6, 15, WHITE), (15, 6, WHITE), (15, 15, GRID)):
    for y in range(cy - 2, cy + 3):
        for x in range(cx - 2, cx + 3):
            if 0 <= y < H and 0 <= x < W:
                px[y][x] = col

# 像素（bottom-up、BGRA）
data = bytearray()
for y in range(H - 1, -1, -1):
    for x in range(W):
        r, g, b, a = px[y][x]
        data += bytes((b, g, r, a))

# AND mask（全 0，靠 alpha）
mask = bytearray([0] * (((W + 31) // 32) * 4 * H))

bih = struct.pack("<IiiHHIIiiII", 40, W, H * 2, 1, 32, 0, len(data) + len(mask), 0, 0, 0, 0)
img = bih + bytes(data) + bytes(mask)

hdr = struct.pack("<HHH", 0, 1, 1)
entry = struct.pack("<BBBBHHII", W % 256, H % 256, 0, 0, 1, 32, len(img), 6 + 16)

import os
out = os.path.join(os.path.dirname(os.path.abspath(__file__)), "src-tauri", "icons", "icon.ico")
os.makedirs(os.path.dirname(out), exist_ok=True)
with open(out, "wb") as f:
    f.write(hdr + entry + img)
print("icon.ico", os.path.getsize(out), "bytes")
