# -*- coding: utf-8 -*-
"""探测：行程单 / 机票 / 火车票 / 打车单 这些非发票凭证能抽出什么。
只读，不改任何东西，不弹窗。
"""
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent.parent))

SAMPLES = [
    r"E:\桌\审计\个人\报销凭证\某某投标\滴滴出行电子发票及行程报销单\滴滴出行行程报销单.pdf",
    r"E:\桌\审计\个人\报销凭证\某某投标\【T3出行-15.86元-1个行程】高德打车电子发票.pdf",
    r"E:\桌\审计\个人\报销凭证\某某项目项目\兰州—济南（行程单）.pdf",
    r"E:\桌\审计\个人\报销凭证\某某项目项目\某某项目差旅费（新）\1-济南-兰州-嘉峪关机票880（李四）.pdf",
    r"E:\桌\审计\个人\报销凭证\某某项目项目\某某项目差旅费（新）\11-酒泉南至张掖西火车票59(张三).pdf",
    r"E:\桌\审计\个人\报销凭证\某某公司\张三-地铁.pdf",
    r"E:\桌\审计\个人\报销凭证\某某项目项目\保险.pdf",
    r"E:\桌\审计\个人\报销凭证\某某项目项目\某某项目差旅费（新）\14-兰州某某区易佰酒店住宿费264.00.pdf",
    r"E:\桌\审计\个人\报销凭证\国网山东泰安肥城市供电公司发票和收据\国网山东泰安肥城市供电公司 110kV 仪南变电站一键顺控改造工程\仪南.pdf",
]


def dump_text(path: Path, max_lines=46):
    print("=" * 78)
    print(f"【{path.name}】  {path.stat().st_size // 1024} KB")
    if not path.exists():
        print("   !! 文件不存在")
        return
    try:
        from pypdf import PdfReader
        r = PdfReader(str(path))
        print(f"   页数：{len(r.pages)}")
        for i, pg in enumerate(r.pages[:2]):
            t = pg.extract_text() or ""
            lines = [ln.strip() for ln in t.splitlines() if ln.strip()]
            print(f"   --- 第 {i+1} 页 文本行 {len(lines)} ---")
            for ln in lines[:max_lines]:
                print(f"     {ln}")
    except Exception as e:
        print(f"   !! 抽取失败 {type(e).__name__}: {e}")


def main():
    print("### 现有解析器对它们的判断 ###")
    try:
        import invparse
        for s in SAMPLES:
            p = Path(s)
            if not p.exists():
                continue
            try:
                rec = invparse.parse_file(p)
                print(f"  {p.name[:44]:46s} -> kind={rec.get('kind')!r} no={rec.get('no')!r} "
                      f"date={rec.get('date')!r} total={rec.get('total')!r} warn={rec.get('warn')}")
            except Exception as e:
                print(f"  {p.name[:44]:46s} -> 异常 {type(e).__name__}: {e}")
    except Exception as e:
        print(f"  !! 导入 invparse 失败：{e}")

    print()
    print("### PDF 文本层实际内容 ###")
    for s in SAMPLES:
        dump_text(Path(s))


if __name__ == "__main__":
    main()
