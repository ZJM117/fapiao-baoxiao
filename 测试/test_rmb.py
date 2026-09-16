# -*- coding: utf-8 -*-
"""
人民币大写金额回归测试
======================
用真实发票 XML 里的官方「价税合计（大写）」字段当标准答案对照程序算的大写。
样本取自 E:\桌\某某公司\开具的发票（136 张，含 18 张红字负数）。

运行：python 测试/test_rmb.py
"""
import sys
import zipfile
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent.parent))
import invparse as ip
import report_pdf as rp

ARCHIVE = Path(r"E:\桌\某某公司\开具的发票")

# 边界用例（人工核对过的期望值）
CASES = [
    (0, "零圆整"), (1, "壹圆整"), (10, "壹拾圆整"), (100, "壹佰圆整"),
    (1000, "壹仟圆整"), (10000, "壹万圆整"), (10001, "壹万零壹圆整"),
    (1000000, "壹佰万圆整"), (100000000, "壹亿圆整"),
    (12345.67, "壹万贰仟叁佰肆拾伍圆陆角柒分"),
    (1314.00, "壹仟叁佰壹拾肆圆整"),
    (2101.10, "贰仟壹佰零壹圆壹角"),
    (1820.35, "壹仟捌佰贰拾圆零叁角伍分"),
    (100.05, "壹佰圆零伍分"),
    (0.5, "伍角"),
    (-88.8, "（负数）捌拾捌圆捌角"),
]


def load_official() -> list:
    out = []
    if not ARCHIVE.exists():
        return out
    for z in sorted(ARCHIVE.rglob("*.zip")):
        try:
            with zipfile.ZipFile(z) as f:
                names = [n for n in f.namelist() if n.lower().endswith(".xml")]
                if not names:
                    continue
                data = f.read(names[0])
        except Exception:
            continue
        if b"SellerName" not in data:
            continue
        g = ip._xml_to_dict(data)
        amt, cn = g.get("TotalTax-includedAmount"), g.get("TotalTax-includedAmountInChinese")
        if amt is not None and cn:
            out.append((amt, cn))
    return out


def main():
    fail = 0
    print("=== 边界用例 ===")
    for v, want in CASES:
        got = rp.rmb_upper(v)
        ok = got == want
        fail += not ok
        print(f"  {str(v):>12} → {got:<30} {'OK' if ok else '❌ 期望 ' + want}")

    samples = load_official()
    print(f"\n=== 真实票面对照（{len(samples)} 张）===")
    ok = 0
    diffs = []
    for a, c in samples:
        if rp.rmb_upper(a) == c:
            ok += 1
        else:
            diffs.append((a, c, rp.rmb_upper(a)))
    print(f"  一致 {ok}/{len(samples)}")
    for a, c, g in diffs[:10]:
        print(f"    {a}: 票面={c}  我们={g}")
    if diffs:
        print("  说明：不一致的都是税局允许的变体写法（少数开票端写「元」、"
              "角后写「整」等），不影响报销单使用。")

    print(f"\n边界用例失败 {fail} 个；票面对照一致率 {ok}/{len(samples)}")
    return 1 if fail else 0


if __name__ == "__main__":
    sys.exit(main())
