"""使い方:
    python -m xlsxroll 前月.xlsx 今月.xlsx --replace "9月=10月" --shift-months 1 --check "9月"
    python -m xlsxroll 前月.xlsx --replace "9月=10月" --dry-run     # 変わるところを見るだけ
"""

from __future__ import annotations

import argparse
import sys
from pathlib import Path

from .core import roll


def parse_pair(s: str) -> tuple[str, str]:
    if "=" not in s:
        raise argparse.ArgumentTypeError(f"「前=後」の形で書いてください: {s}")
    old, new = s.split("=", 1)
    return old, new


def main(argv: list[str] | None = None) -> int:
    if hasattr(sys.stdout, "reconfigure"):
        sys.stdout.reconfigure(encoding="utf-8")
    p = argparse.ArgumentParser(prog="xlsxroll", description="Excel の月の表記と日付を、書式を崩さずに進める")
    p.add_argument("src", type=Path, help="元のファイル（前月分など）")
    p.add_argument("dst", type=Path, nargs="?", help="保存先（--dry-run なら不要）")
    p.add_argument("--replace", type=parse_pair, action="append", default=[], metavar="前=後",
                   help="文字の置き換え。何回でも指定できる（上から順に適用）")
    p.add_argument("--shift-months", type=int, default=0, help="日付のセルを何か月進めるか（月末は月末のまま）")
    p.add_argument("--check", action="append", default=[], metavar="文字",
                   help="保存後に残っていたら知らせる文字（直し忘れの確認）")
    p.add_argument("--dry-run", action="store_true", help="保存せず、変わるところだけを表示する")
    args = p.parse_args(argv)

    if not args.dry_run and args.dst is None:
        p.error("保存先を指定するか、--dry-run を付けてください")
    if args.dst and args.dst.resolve() == args.src.resolve():
        p.error("元のファイルと同じ名前には保存できません（元のファイルは残します）")

    report = roll(args.src, None if args.dry_run else args.dst, args.replace, args.shift_months, args.check)
    for c in report.changes:
        print(f"変更  {c.sheet}!{c.cell}: {c.before} → {c.after}")
    for c in report.leftovers:
        print(f"要確認 {c.sheet}!{c.cell}: 「{c.before}」に確認したい文字が残っています")
    print(f"{'（確認のみ）' if args.dry_run else '保存しました: ' + str(args.dst)}  変更 {len(report.changes)}件 / 要確認 {len(report.leftovers)}件")
    return 1 if report.leftovers else 0


if __name__ == "__main__":
    sys.exit(main())
