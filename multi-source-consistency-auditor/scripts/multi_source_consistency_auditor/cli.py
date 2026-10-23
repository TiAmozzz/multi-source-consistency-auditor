"""Command-line interface for the consistency auditor."""

from __future__ import annotations

import argparse
import sys
from pathlib import Path

from .profile import ProfileError, load_profile
from .reports import SUPPORTED_FORMATS, write_reports
from .rules import AuditEngine


if hasattr(sys.stdout, "reconfigure"):
    try:
        sys.stdout.reconfigure(encoding="utf-8")
        sys.stderr.reconfigure(encoding="utf-8")
    except (AttributeError, OSError):
        pass


def _is_within(path: Path, parent: Path) -> bool:
    try:
        path.resolve().relative_to(parent.resolve())
        return True
    except ValueError:
        return False


def _validate_output_location(profile, output_dir: str | Path) -> None:
    output = Path(output_dir)
    for source in profile.sources:
        for source_path in profile.resolve_paths(source):
            if source_path.is_dir() and _is_within(output, source_path):
                raise ValueError(
                    f"输出目录不能位于输入来源目录内：{source_path}"
                )
            if source_path.is_file() and output.resolve() == source_path.resolve():
                raise ValueError("输出目录不能覆盖输入文件")
            if source_path.is_file() and source_path.parent.resolve() == output.resolve():
                raise ValueError(
                    f"输出目录包含输入文件，可能造成覆盖：{source_path}"
                )


def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(
        prog="multi-source-audit",
        description="核验多个表格和文档来源中的数据一致性，并输出带证据的报告。",
    )
    parser.add_argument("--profile", required=True, help="JSON 核验配置路径")
    parser.add_argument("--output", default="audit-output", help="报告输出目录")
    parser.add_argument(
        "--format",
        dest="formats",
        action="append",
        choices=sorted(SUPPORTED_FORMATS | {"all"}),
        help=(
            "输出格式：html、xlsx、docx、pdf、markdown、json 或 csv。"
            "可重复指定；不指定时生成 HTML，指定 all 时生成全部格式"
        ),
    )
    parser.add_argument(
        "--fail-on-error",
        action="store_true",
        help="发现 error 级问题时返回退出码 1",
    )
    parser.add_argument(
        "--quiet",
        action="store_true",
        help="只输出错误，不打印核验摘要",
    )
    return parser


def main(argv: list[str] | None = None) -> int:
    args = build_parser().parse_args(argv)
    try:
        profile = load_profile(args.profile)
        _validate_output_location(profile, args.output)
        result = AuditEngine(profile).run()
        paths = write_reports(result, args.output, args.formats)
    except (ProfileError, OSError, ValueError) as exc:
        print(f"核验失败：{exc}", file=sys.stderr)
        return 2

    if not args.quiet:
        summary = result.to_dict()["summary"]
        print(f"核验配置：{result.profile_name}")
        print(
            f"结果：{result.status} | 实体 {summary['entities']} | "
            f"通过 {summary['passed_entities']} | "
            f"错误 {summary['errors']} | 警告 {summary['warnings']}"
        )
        for format_name, path in paths.items():
            print(f"{format_name.upper()} 差异汇总：{Path(path)}")

    if args.fail_on_error and result.error_count:
        return 1
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
