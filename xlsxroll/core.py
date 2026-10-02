"""xlsx の中の XML のうち、変える文字と数字だけを書き換える。

Excel を開いて保存し直すライブラリ（openpyxl など）を使うと、既定のフォント・列幅・印刷設定・
条件付き書式などが変わってしまうことがある。ここでは zip の中身をそのままコピーし、
共有文字列（sharedStrings.xml）とシートの XML の「変えたいところ」だけを置き換える。
"""

from __future__ import annotations

import calendar
import re
import zipfile
from dataclasses import dataclass, field
from datetime import date, timedelta
from html import unescape
from pathlib import Path
from xml.sax.saxutils import escape

EXCEL_EPOCH = date(1899, 12, 30)

# 日付として表示される組み込みの表示形式（日本語版 Excel の 55〜58 なども含む）
BUILTIN_DATE_FORMATS = set(range(14, 18)) | {22} | set(range(27, 37)) | set(range(50, 59))

SI_RE = re.compile(r"<si>(.*?)</si>|<si/>", re.S)
T_RE = re.compile(r"(<t(?:\s[^>]*)?>)(.*?)(</t>)", re.S)
RPH_RE = re.compile(r"<rPh\b.*?</rPh>", re.S)
CELL_RE = re.compile(r"<c\b([^>]*?)(?:/>|>(.*?)</c>)", re.S)
ATTR_RE = re.compile(r'(\w+)="([^"]*)"')


@dataclass
class Change:
    sheet: str
    cell: str
    before: str
    after: str


@dataclass
class Report:
    changes: list[Change] = field(default_factory=list)
    leftovers: list[Change] = field(default_factory=list)  # 直し忘れの候補（after は空）


def add_months(d: date, months: int) -> date:
    """月末の日付は月末のまま進める（1/31 → 2/28、4/30 → 5/31）。"""
    y, m = divmod(d.month - 1 + months, 12)
    y, m = d.year + y, m + 1
    last = calendar.monthrange(y, m)[1]
    was_month_end = d.day == calendar.monthrange(d.year, d.month)[1]
    return date(y, m, last if was_month_end else min(d.day, last))


def serial_to_date(v: float) -> date:
    return EXCEL_EPOCH + timedelta(days=int(v))


def date_to_serial(d: date) -> int:
    return (d - EXCEL_EPOCH).days


def replace_text(text: str, pairs: list[tuple[str, str]]) -> str:
    for old, new in pairs:
        text = text.replace(old, new)
    return text


def _replace_in_si(si_xml: str, pairs) -> str:
    """1つの共有文字列の <t> だけを置き換える。ふりがな（rPh）の中は触らない。"""
    out, pos = [], 0
    for m in RPH_RE.finditer(si_xml):
        out.append(T_RE.sub(lambda t: _sub_t(t, pairs), si_xml[pos:m.start()]))
        out.append(m.group(0))
        pos = m.end()
    out.append(T_RE.sub(lambda t: _sub_t(t, pairs), si_xml[pos:]))
    return "".join(out)


def _sub_t(m: re.Match, pairs) -> str:
    raw = m.group(2)
    new = replace_text(unescape(raw), pairs)
    if new == unescape(raw):
        return m.group(0)
    return m.group(1) + escape(new) + m.group(3)


def plain_text(si_xml: str) -> str:
    return "".join(unescape(t.group(2)) for t in T_RE.finditer(RPH_RE.sub("", si_xml)))


def _date_styles(styles_xml: str) -> set[int]:
    """日付として表示される書式（cellXfs の番号）を集める。"""
    custom = {}
    for m in re.finditer(r'<numFmt\b[^>]*numFmtId="(\d+)"[^>]*formatCode="([^"]*)"', styles_xml):
        code = re.sub(r'"[^"]*"|\[[^\]]*\]|\\.', "", unescape(m.group(2))).lower()
        custom[int(m.group(1))] = "y" in code or "d" in code or "e" in code or "g" in code
    xfs = re.search(r"<cellXfs\b.*?</cellXfs>", styles_xml, re.S)
    result = set()
    if xfs:
        for i, xf in enumerate(re.findall(r"<xf\b[^>]*?(?:/>|>)", xfs.group(0))):
            fid = re.search(r'numFmtId="(\d+)"', xf)
            n = int(fid.group(1)) if fid else 0
            if n in BUILTIN_DATE_FORMATS or custom.get(n):
                result.add(i)
    return result


def _sheet_names(z: zipfile.ZipFile) -> dict[str, str]:
    """xl/worksheets/sheetN.xml → シート名。"""
    wb = z.read("xl/workbook.xml").decode("utf-8")
    rels = z.read("xl/_rels/workbook.xml.rels").decode("utf-8")
    targets = {}
    for m in re.finditer(r"<Relationship\b[^>]*>", rels):
        a = dict(ATTR_RE.findall(m.group(0)))
        t = a.get("Target", "").lstrip("/")
        targets[a.get("Id")] = t if t.startswith("xl/") else "xl/" + t
    names = {}
    for m in re.finditer(r"<sheet\b[^>]*>", wb):
        a = dict(re.findall(r'([\w:]+)="([^"]*)"', m.group(0)))
        path = targets.get(a.get("r:id"))
        if path:
            names[path] = unescape(a.get("name", path))
    return names


def roll(src: Path, dst: Path | None, pairs: list[tuple[str, str]], shift_months: int = 0,
         check_words: list[str] | None = None) -> Report:
    """src を読んで dst に書き出す。dst が None なら何も書かず、変更点だけを返す（確認用）。"""
    report = Report()
    check_words = check_words or []
    with zipfile.ZipFile(src) as z:
        names = _sheet_names(z)
        styles = z.read("xl/styles.xml").decode("utf-8") if "xl/styles.xml" in z.namelist() else ""
        date_styles = _date_styles(styles)

        # 共有文字列
        shared_before, shared_after, new_shared = [], [], None
        if "xl/sharedStrings.xml" in z.namelist():
            xml = z.read("xl/sharedStrings.xml").decode("utf-8")

            def sub_si(m):
                body = m.group(1) or ""
                shared_before.append(plain_text(body))
                new_body = _replace_in_si(body, pairs)
                shared_after.append(plain_text(new_body))
                return m.group(0) if new_body == body else f"<si>{new_body}</si>"

            new_xml = SI_RE.sub(sub_si, xml)
            new_shared = new_xml if new_xml != xml else None

        new_sheets = {}
        for path, sheet in names.items():
            xml = z.read(path).decode("utf-8")

            def sub_cell(m):
                attrs_raw, inner = m.group(1), m.group(2) or ""
                attrs = dict(ATTR_RE.findall(attrs_raw))
                ref, kind = attrs.get("r", "?"), attrs.get("t", "n")
                if kind == "s":
                    v = re.search(r"<v>(\d+)</v>", inner)
                    if v:
                        i = int(v.group(1))
                        before, after = shared_before[i], shared_after[i]
                        if before != after:
                            report.changes.append(Change(sheet, ref, before, after))
                        elif any(w in after for w in check_words):
                            report.leftovers.append(Change(sheet, ref, after, ""))
                    return m.group(0)
                if kind == "inlineStr":
                    new_inner = T_RE.sub(lambda t: _sub_t(t, pairs), inner)
                    before, after = plain_text(inner), plain_text(new_inner)
                    if before != after:
                        report.changes.append(Change(sheet, ref, before, after))
                        return f"<c{attrs_raw}>{new_inner}</c>"
                    if any(w in after for w in check_words):
                        report.leftovers.append(Change(sheet, ref, after, ""))
                    return m.group(0)
                # 数字のセル：日付の書式で、式でないものだけ進める
                if shift_months and kind == "n" and "<f" not in inner and int(attrs.get("s", "0")) in date_styles:
                    v = re.search(r"<v>([\d.]+)</v>", inner)
                    if v and float(v.group(1)) >= 61:
                        old = serial_to_date(float(v.group(1)))
                        new = add_months(old, shift_months)
                        frac = v.group(1).split(".")[1] if "." in v.group(1) else ""
                        new_v = str(date_to_serial(new)) + (f".{frac}" if frac else "")
                        report.changes.append(Change(sheet, ref, old.isoformat(), new.isoformat()))
                        return m.group(0).replace(v.group(0), f"<v>{new_v}</v>", 1)
                return m.group(0)

            new_xml = CELL_RE.sub(sub_cell, xml)
            if new_xml != xml:
                new_sheets[path] = new_xml

        if dst is not None:
            replaced = dict(new_sheets)
            if new_shared is not None:
                replaced["xl/sharedStrings.xml"] = new_shared
            with zipfile.ZipFile(dst, "w") as out:
                for info in z.infolist():
                    data = replaced[info.filename].encode("utf-8") if info.filename in replaced else z.read(info)
                    out.writestr(info, data)
    return report
