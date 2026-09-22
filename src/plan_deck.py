"""从命令行调用真实 OpenAI API 生成 DeckSpec JSON。"""

from __future__ import annotations

import argparse
import sys
from pathlib import Path

from llm import LLMError
from planner import PlanningRequestError, PromptLoadError, plan_deck

DEFAULT_REQUEST = "生成一份6页的研究生开题汇报，简洁蓝白风"


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(description="用自然语言规划一个合法 DeckSpec。")
    parser.add_argument("request", nargs="?", default=DEFAULT_REQUEST, help="PPT 自然语言要求")
    parser.add_argument(
        "--output",
        type=Path,
        default=Path("output/planned_deck.json"),
        help="DeckSpec JSON 输出路径",
    )
    return parser.parse_args()


def main() -> int:
    args = parse_args()
    try:
        deck = plan_deck(args.request)
    except (LLMError, PlanningRequestError, PromptLoadError) as exc:
        print(f"规划失败：{exc}", file=sys.stderr)
        return 1

    output_path: Path = args.output.resolve()
    output_path.parent.mkdir(parents=True, exist_ok=True)
    output_path.write_text(deck.model_dump_json(indent=2), encoding="utf-8")
    print(f"已生成 {len(deck.slides)} 页 DeckSpec：{output_path}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
