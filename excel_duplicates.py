"""
Excel 重复行检测工具。

功能：
    - 从包含“名字、邮箱、链接、国家”等表头的大体量 Excel 文件中读取数据；
    - 按“名字 + 链接”判断重复（两列同时相等即视为重复）；
    - 将所有属于重复组合的行导出到新的 Excel 文件。

核心对外函数：
    find_duplicate_rows(input_path, output_path, name_col="名字", link_col="链接")
"""

from __future__ import annotations

from pathlib import Path
from typing import Iterable, Optional, Sequence, Union

import pandas as pd


PathLike = Union[str, Path]


def find_duplicate_rows(
    input_path: PathLike,
    output_path: PathLike,
    name_col: str = "名字",
    link_col: str = "链接",
    extra_columns: Optional[Sequence[str]] = None,
    write_empty: bool = True,
) -> pd.DataFrame:
    """
    按“名字 + 链接”查找 Excel 中的重复行，并导出到新的 Excel 文件。

    重复定义：
        同一行的 `name_col` 与 `link_col` 值组成的键相同，
        且该键在整表中出现次数 >= 2。所有属于这些键的行都视为“重复行”。

    参数
    ----------
    input_path : str | pathlib.Path
        输入 Excel 文件路径（通常为 .xlsx）。
    output_path : str | pathlib.Path
        输出 Excel 文件路径。重复行将被写入该文件。
    name_col : str, 默认 "名字"
        作为“名字”的列名。
    link_col : str, 默认 "链接"
        作为“链接”的列名。
    extra_columns : Sequence[str], 可选
        如果指定，则在读取 Excel 时尽量只读取
        [name_col, link_col] + extra_columns 这些列，以减少内存占用。
        比如可以传入 ["邮箱", "国家"]。
        注意：如果某个列名在表头中不存在，函数会自动忽略该列名。
    write_empty : bool, 默认 True
        若未发现任何重复行时：
            - 为 True：仍然生成一个空的 Excel 文件（只有表头或完全为空）。
            - 为 False：不写出任何文件，只返回空 DataFrame。

    返回
    ----------
    pandas.DataFrame
        包含所有重复行的 DataFrame。调用方可以直接使用该结果进行后续处理。

    异常
    ----------
    FileNotFoundError
        当 input_path 指向的文件不存在时抛出。
    ValueError
        当 Excel 中缺少 name_col 或 link_col 时抛出。
    """

    input_path = Path(input_path)
    output_path = Path(output_path)

    if not input_path.is_file():
        raise FileNotFoundError(f"输入文件不存在：{input_path}")

    # 仅读取必要列，以降低内存压力
    usecols: Optional[Iterable[str]] = None
    if extra_columns is not None:
        # 暂定的列集合（稍后会根据实际表头过滤无效列）
        usecols = set([name_col, link_col, *extra_columns])

    # 读取 Excel；对百万级行数，建议确保本机内存 >= 8GB
    df = pd.read_excel(input_path, engine="openpyxl", usecols=usecols)

    missing_cols = [col for col in (name_col, link_col) if col not in df.columns]
    if missing_cols:
        raise ValueError(
            f"Excel 中缺少必需列：{missing_cols}，"
            f"当前表头为：{list(df.columns)}"
        )

    # duplicated(keep=False) 会把所有重复项标记为 True
    dup_mask = df.duplicated(subset=[name_col, link_col], keep=False)
    dup_df = df[dup_mask].copy()

    if dup_df.empty:
        if write_empty:
            # 写出一个空文件，方便调用方知道流程已执行
            dup_df.to_excel(output_path, index=False)
        # 直接返回空 DataFrame
        return dup_df

    # 将重复行写出到新的 Excel 中
    dup_df.to_excel(output_path, index=False)

    return dup_df


if __name__ == "__main__":
    # 简单示例：可自行修改路径进行本地测试
    example_input = Path("input.xlsx")
    example_output = Path("duplicates.xlsx")

    if example_input.exists():
        result = find_duplicate_rows(
            example_input,
            example_output,
            name_col="名字",
            link_col="链接",
            extra_columns=["邮箱", "国家"],
        )
        print(f"已检测到重复行数：{len(result)}，结果已写入：{example_output}")
    else:
        print(
            f"示例输入文件 {example_input} 不存在。"
            " 请将你的 Excel 文件重命名为 input.xlsx 放在当前目录后再次运行。"
        )


