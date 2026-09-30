#!/usr/bin/env python
"""Create a reproducible hotel-standard search record."""

from __future__ import annotations

import argparse
import json
import urllib.parse
import webbrowser
from datetime import datetime
from pathlib import Path


def main() -> int:
    parser = argparse.ArgumentParser(description="按汉庭优先、全季备用生成附近酒店价格查询记录")
    parser.add_argument("--location", required=True, help="出差点或现场地址")
    parser.add_argument("--check-in", required=True, help="入住日期 YYYY-MM-DD")
    parser.add_argument("--check-out", required=True, help="离店日期 YYYY-MM-DD")
    parser.add_argument("--output", required=True, help="查询记录 JSON 文件")
    parser.add_argument("--hotel", help="人工确认的酒店名称")
    parser.add_argument("--price", type=float, help="人工确认的每晚价格")
    parser.add_argument("--source-url", help="价格来源 URL")
    parser.add_argument("--open-browser", action="store_true", help="打开查询页面")
    args = parser.parse_args()

    query = f"{args.location} 汉庭 全季 {args.check_in} {args.check_out} 酒店价格"
    encoded = urllib.parse.quote(query)
    links = {
        "Bing": f"https://www.bing.com/search?q={encoded}",
        "百度": f"https://www.baidu.com/s?wd={encoded}",
        "高德地图": f"https://ditu.amap.com/search?query={urllib.parse.quote(args.location + ' 汉庭酒店')}",
        "百度地图": f"https://map.baidu.com/search/{urllib.parse.quote(args.location + ' 汉庭酒店')}",
    }
    if args.open_browser:
        for url in links.values():
            webbrowser.open(url)

    data = {
        "created_at": datetime.now().isoformat(timespec="seconds"),
        "location": args.location,
        "check_in": args.check_in,
        "check_out": args.check_out,
        "standard": "最近的汉庭；没有汉庭时使用最近的全季",
        "priority": ["汉庭", "全季"],
        "selected_hotel": args.hotel or "待人工确认",
        "price_per_night": args.price,
        "price_source": args.source_url or "待人工确认",
        "search_links": links,
        "status": "已记录人工报价" if args.price is not None and args.hotel else "待人工查询",
        "note": "价格必须结合入住日期、房型、税费和可报销标准核对；本脚本不从搜索摘要推断价格。",
    }
    output = Path(args.output).expanduser().resolve()
    output.parent.mkdir(parents=True, exist_ok=True)
    output.write_text(json.dumps(data, ensure_ascii=False, indent=2), encoding="utf-8")
    print(output)
    print(json.dumps(data, ensure_ascii=False, indent=2))
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
