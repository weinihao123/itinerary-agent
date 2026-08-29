# -*- coding: utf-8 -*-
"""行程智能体 命令行入口

子命令：
  add      口语/JSON 录入（解析→草稿→确认层级→落库）
  edit     按记录ID修改字段
  delete   软删除（默认）或物理删除（--hard）
  list     列出活跃记录
  export   导出 Excel（明细 + 汇总）
  dashboard 生成看板 HTML
  parse    仅解析口语文本，不落库

示例：
  python run.py add --text "中国移动河南郑州AI训战项目，8月1日到8月3日执行，7月31日出发8月4日返程" --yes
  python run.py edit 1 项目类型 云
  python run.py delete 1
  python run.py export
  python run.py dashboard
"""
import argparse
import json
import sys

import agent_lib as L


def _print_rows(rows):
    cols = ["记录ID", "项目名称", "运营商", "执行省份", "执行地市", "项目类型",
            "项目执行开始", "项目执行结束", "出发日期", "返程日期",
            "项目执行天数", "项目总计花费时间", "讲师层级"]
    width = [5, 22, 6, 8, 8, 6, 12, 12, 12, 12, 9, 11, 8]
    header = "  ".join(c.ljust(w) for c, w in zip(cols, width))
    print(header)
    print("-" * len(header))
    for r in rows:
        print("  ".join(str(r.get(c, "")).ljust(w) for c, w in zip(cols, width)))


def cmd_add(args):
    if args.json:
        found = json.loads(args.json)
        missing = [k for k in ["项目名称", "运营商", "执行省份", "执行地市", "项目类型",
                                "项目执行开始", "项目执行结束", "出发日期", "返程日期"]
                   if not found.get(k)]
    else:
        res = L.parse_oral_text(args.text or "")
        found = res["found"]
        missing = res["missing"]

    if missing:
        print("解析草稿：")
        print(json.dumps(found, ensure_ascii=False, indent=2))
        print("缺失字段（需补充）：", ", ".join(missing))
        print("请用 --json 补全后重试，或补充口语信息。")
        return 1

    rows = L.load_rows()
    level = L.compute_level(rows, found["项目名称"], override=args.level)
    print("解析草稿：")
    print(json.dumps(found, ensure_ascii=False, indent=2))
    print(f"拟写入讲师层级 = {level}（同项目上一期 +1，首期默认1）")
    if not args.yes:
        print("确认无误请加 --yes 写入；如需改写首期层级加 --level N。")
        return 0
    rec = L.add_trip(found, level_override=args.level)
    print(f"已写入，记录ID = {rec['记录ID']}")
    return 0


def cmd_edit(args):
    if len(args.pairs) % 2 != 0:
        print("edit 需成对提供 字段 值")
        return 1
    updates = {args.pairs[i]: args.pairs[i + 1] for i in range(0, len(args.pairs), 2)}
    rec = L.edit_trip(args.id, updates)
    if rec is None:
        print(f"未找到记录ID = {args.id}")
        return 1
    print(f"已更新记录ID = {args.id}")
    return 0


def cmd_delete(args):
    if args.hard:
        L.hard_delete(args.id)
        print(f"已物理删除记录ID = {args.id}")
    else:
        L.soft_delete(args.id)
        print(f"已软删除记录ID = {args.id}（看板与导出将过滤）")
    return 0


def cmd_list(args):
    rows = L.active_rows()
    if not rows:
        print("暂无活跃记录。")
        return 0
    _print_rows(rows)
    return 0


def cmd_export(args):
    path = L.export_excel()
    print("已导出：", path)
    return 0


def cmd_dashboard(args):
    path = L.build_dashboard()
    print("已生成看板：", path)
    return 0


def cmd_parse(args):
    res = L.parse_oral_text(args.text or "")
    print(json.dumps(res, ensure_ascii=False, indent=2))
    return 0


def main():
    p = argparse.ArgumentParser(description="行程记录与分析智能体")
    sub = p.add_subparsers(dest="cmd")

    a = sub.add_parser("add")
    a.add_argument("--text", help="口语化描述")
    a.add_argument("--json", help="JSON 字段")
    a.add_argument("--yes", action="store_true", help="确认写入")
    a.add_argument("--level", type=int, default=None, help="首期层级改写")

    e = sub.add_parser("edit")
    e.add_argument("id")
    e.add_argument("pairs", nargs="*")

    d = sub.add_parser("delete")
    d.add_argument("id")
    d.add_argument("--hard", action="store_true")

    sub.add_parser("list")
    sub.add_parser("export")
    sub.add_parser("dashboard")

    pa = sub.add_parser("parse")
    pa.add_argument("--text", required=True)

    args = p.parse_args()
    if not args.cmd:
        p.print_help()
        return 1
    return getattr(sys.modules[__name__], "cmd_" + args.cmd)(args)


if __name__ == "__main__":
    sys.exit(main())
