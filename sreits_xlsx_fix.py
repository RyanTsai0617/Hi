#!/usr/bin/env python3
"""S-REITs xlsx OneDrive/Excel 相容性修復（inline string → shared string）

問題
----
S-REITs 週報的 xlsx 由 openpyxl 產出。openpyxl 預設把字串以「行內字串」
(inline string, ``t="inlineStr"``) 寫入，而**不建立共用字串表**
(sharedStrings)。這種檔在桌機 Excel 多半開得起來，但在 **OneDrive /
Excel 網頁版**上容易被判定為「無法開啟活頁簿 / 內容需修復」——這正是使用者
遇到的「0713 打不開」。

實測比對可證：能正常開的「可開啟修復版」用的就是 sharedStrings；壞掉的主檔
用的是 inlineStr。兩者差別就在字串儲存方式。

這支工具把 openpyxl 產出的 inline-string xlsx **原地轉成 shared-string
格式**（就是那個「能開」的格式），過程只改寫「字串如何存放」，不動任何資料、
超連結、合併儲存格、欄寬、樣式，因此 100% 保真。同時移除空的
``<numFmts count="0"/>``（OOXML schema 不允許，Excel 會挑剔）。

用法
----
    # 原地修復（覆蓋原檔）—— 接在產表流程「wb.save() 之後」呼叫
    python sreits_xlsx_fix.py 新加坡REITs_新聞_20260712_含摘要.xlsx *.xlsx

    # 另存一份
    python sreits_xlsx_fix.py in.xlsx --out in_fixed.xlsx

內嵌用法（建議直接接在 sreits_build_xlsx.py 的每個 wb.save(out_xlsx) 之後）：
    from sreits_xlsx_fix import fix_xlsx
    wb.save(out_xlsx)
    fix_xlsx(out_xlsx)          # 原地轉 shared string，OneDrive 就開得起來
"""
from __future__ import annotations

import argparse
import os
import re
import shutil
import sys
import tempfile
import zipfile

# 命名空間
_MAIN_NS = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
_REL_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
_SST_REL = _REL_NS + "/sharedStrings"
_CT_NS = "http://schemas.openxmlformats.org/package/2006/content-types"

# 抓一個 inline-string 儲存格：<c ... t="inlineStr" ...>(<is>...</is>)</c>
# <is> 內不會再出現 </c>，故非貪婪比對安全。
_CELL_RE = re.compile(
    r'<c\b([^>]*?)\bt="inlineStr"([^>]*?)>(.*?)</c>',
    re.DOTALL,
)
_IS_RE = re.compile(r"<is>(.*)</is>", re.DOTALL)


def _convert_sheet(xml: str, sst: list[str], index: dict[str, int]) -> tuple[str, bool]:
    """把單一 worksheet 的 inlineStr 儲存格改成 shared string 參照。

    sst / index 為跨工作表共用的字串表（去重）。回傳 (新 XML, 是否有變動)。
    """
    changed = False

    def repl(m: re.Match) -> str:
        nonlocal changed
        pre, post, body = m.group(1), m.group(2), m.group(3)
        is_m = _IS_RE.search(body)
        # <si> 與 <is> 使用相同的子元素模型，故 <is> 內容可直接當成 <si> 內容
        inner = is_m.group(1) if is_m else "<t></t>"
        if inner not in index:
            index[inner] = len(sst)
            sst.append(inner)
        idx = index[inner]
        attrs = (pre + post)  # 保留 r=、s= 等其餘屬性，去掉 t="inlineStr"
        attrs = re.sub(r"\s+", " ", attrs).strip()
        attrs = (" " + attrs) if attrs else ""
        changed = True
        return f'<c{attrs} t="s"><v>{idx}</v></c>'

    return _CELL_RE.sub(repl, xml), changed


def _build_sharedstrings(sst: list[str]) -> str:
    count = len(sst)  # 這裡每個 inline 儲存格都各自參照，uniqueCount==count（已去重）
    parts = [
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>',
        f'<sst xmlns="{_MAIN_NS}" count="{count}" uniqueCount="{count}">',
    ]
    parts.extend(f"<si>{s}</si>" for s in sst)
    parts.append("</sst>")
    return "".join(parts)


def _ensure_content_type(ct_xml: str) -> str:
    if "sharedStrings.xml" in ct_xml:
        return ct_xml
    override = (
        '<Override PartName="/xl/sharedStrings.xml" '
        'ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml"/>'
    )
    return ct_xml.replace("</Types>", override + "</Types>", 1)


def _ensure_workbook_rel(rels_xml: str) -> str:
    if "sharedStrings.xml" in rels_xml:
        return rels_xml
    used = set(re.findall(r'Id="(rId\d+)"', rels_xml))
    n = 1
    while f"rId{n}" in used:
        n += 1
    rel = (
        f'<Relationship Id="rId{n}" Type="{_SST_REL}" Target="sharedStrings.xml"/>'
    )
    return rels_xml.replace("</Relationships>", rel + "</Relationships>", 1)


def fix_xlsx(path: str, out: str | None = None) -> str:
    """把 inline-string xlsx 原地（或另存）轉成 shared-string 格式。

    回傳實際寫出的路徑。若檔案本來就沒有 inline string，仍會移除空 numFmts。
    """
    target = out or path
    with zipfile.ZipFile(path) as zin:
        items = zin.infolist()
        blobs = {it.filename: zin.read(it.filename) for it in items}

    sst: list[str] = []
    index: dict[str, int] = {}
    any_change = False

    sheet_names = [n for n in blobs if re.fullmatch(r"xl/worksheets/sheet\d+\.xml", n)]
    for name in sheet_names:
        xml = blobs[name].decode("utf-8")
        new_xml, changed = _convert_sheet(xml, sst, index)
        if changed:
            blobs[name] = new_xml.encode("utf-8")
            any_change = True

    if any_change:
        blobs["xl/sharedStrings.xml"] = _build_sharedstrings(sst).encode("utf-8")
        if "[Content_Types].xml" in blobs:
            blobs["[Content_Types].xml"] = _ensure_content_type(
                blobs["[Content_Types].xml"].decode("utf-8")
            ).encode("utf-8")
        if "xl/_rels/workbook.xml.rels" in blobs:
            blobs["xl/_rels/workbook.xml.rels"] = _ensure_workbook_rel(
                blobs["xl/_rels/workbook.xml.rels"].decode("utf-8")
            ).encode("utf-8")

    # 移除空的 <numFmts count="0"/>（schema 不合法）
    if "xl/styles.xml" in blobs:
        styles = blobs["xl/styles.xml"].decode("utf-8")
        new_styles = re.sub(r'<numFmts count="0"\s*/>', "", styles)
        new_styles = re.sub(r'<numFmts count="0"\s*>\s*</numFmts>', "", new_styles)
        if new_styles != styles:
            blobs["xl/styles.xml"] = new_styles.encode("utf-8")

    # 寫回（保留 sharedStrings 這個新 entry 的順序不影響 Excel）
    tmp_fd, tmp_path = tempfile.mkstemp(suffix=".xlsx", dir=os.path.dirname(target) or ".")
    os.close(tmp_fd)
    ordered = list(blobs.keys())
    if "xl/sharedStrings.xml" in ordered:
        # 放在 styles.xml 附近，貼近 Excel 慣例（非必要，純美觀）
        ordered.remove("xl/sharedStrings.xml")
        pos = ordered.index("xl/styles.xml") if "xl/styles.xml" in ordered else len(ordered)
        ordered.insert(pos, "xl/sharedStrings.xml")
    try:
        with zipfile.ZipFile(tmp_path, "w", zipfile.ZIP_DEFLATED) as zout:
            for name in ordered:
                zout.writestr(name, blobs[name])
        # fail-fast：回讀驗證
        with zipfile.ZipFile(tmp_path) as zt:
            bad = zt.testzip()
            if bad:
                raise zipfile.BadZipFile(f"寫出的檔毀損：{bad}")
        shutil.move(tmp_path, target)
    finally:
        if os.path.exists(tmp_path):
            os.remove(tmp_path)
    return target


def main(argv: list[str] | None = None) -> int:
    ap = argparse.ArgumentParser(
        description="S-REITs xlsx 修復：inline string → shared string，讓 OneDrive/Excel 網頁版開得起來"
    )
    ap.add_argument("files", nargs="+", help="要修復的 xlsx（可多個 / 萬用字元）")
    ap.add_argument("--out", default=None, help="輸出路徑（只在單一輸入檔時有效）；省略則原地覆蓋")
    args = ap.parse_args(argv)

    if args.out and len(args.files) != 1:
        ap.error("--out 只能搭配單一輸入檔使用")

    rc = 0
    for f in args.files:
        if not os.path.isfile(f):
            print(f"[skip] 找不到檔案：{f}", file=sys.stderr)
            rc = 1
            continue
        try:
            written = fix_xlsx(f, out=args.out)
            print(f"[ok] 已轉 shared-string → {written}")
        except Exception as e:  # noqa: BLE001
            print(f"[fail] {f}: {type(e).__name__}: {e}", file=sys.stderr)
            rc = 1
    return rc


if __name__ == "__main__":
    raise SystemExit(main())
