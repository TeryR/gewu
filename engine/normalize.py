# -*- coding: utf-8 -*-
"""归一化层：修复非标 xlsx（WPS/老版 Excel 产物）中 openpyxl 无法解析的问题。
修复产物写入 .cache，原文件永不改动。"""
import hashlib
import os
import re
import zipfile

CACHE_DIR = os.path.join(os.path.dirname(os.path.dirname(os.path.abspath(__file__))), ".cache")


def _cache_key(path: str) -> str:
    st = os.stat(path)
    raw = f"{os.path.abspath(path)}|{st.st_size}|{st.st_mtime_ns}"
    return hashlib.sha1(raw.encode("utf-8")).hexdigest()[:16]


def sanitize_xlsx(path: str) -> tuple[str, bool]:
    """返回 (可安全加载的路径, 是否生成了修复副本)。"""
    with zipfile.ZipFile(path) as z:
        if "xl/styles.xml" not in z.namelist():
            return path, False
        styles = z.read("xl/styles.xml").decode("utf-8", "ignore")

    # 已知问题1：空 <fill/> 节点 → openpyxl TypeError: Fill() takes no arguments
    if not re.search(r"<fill\s*/>", styles):
        return path, False

    fixed = re.sub(r"<fill\s*/>", '<fill><patternFill patternType="none"/></fill>', styles)
    os.makedirs(CACHE_DIR, exist_ok=True)
    dst = os.path.join(CACHE_DIR, f"{_cache_key(path)}.xlsx")
    if not os.path.exists(dst):
        with zipfile.ZipFile(path) as zin, zipfile.ZipFile(dst, "w", zipfile.ZIP_DEFLATED) as zout:
            for item in zin.infolist():
                data = zin.read(item.filename)
                if item.filename == "xl/styles.xml":
                    data = fixed.encode("utf-8")
                zout.writestr(item, data)
    return dst, True
