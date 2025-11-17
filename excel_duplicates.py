"""
Excel 重复行检测工具。

功能：
    - 从包含“名字、邮箱、链接、国家”等表头的大体量 Excel 文件中读取数据；
    - 按“名字 + 链接”判断重复（两列同时相等即视为重复）；
    - 将所有属于重复组合的行导出到新的 Excel 文件。

核心对外函数：
    find_duplicate_rows(input_path, output_path, name_col="名字", link_col="链接")
    detect_duplicate_scenarios(input_path, output_full_path, output_duplicates_path, ...)
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


def detect_duplicate_scenarios(
    input_path: PathLike,
    output_full_path: PathLike,
    output_duplicates_path: PathLike,
    country_col: str = "国家",
    company_col: str = "企业名称",
    website_col: str = "网址",
    note_col: str = "备注",
    scenario_separator: str = "；",
) -> tuple[pd.DataFrame, pd.DataFrame]:
    """
    检测Excel文件中的6种重复场景，并为每行添加备注列。

    功能：
        - 逐行与其他所有行比较，检测6种重复场景
        - 为每行添加备注列，标记所有匹配的场景（用分隔符连接）
        - 非重复行标记"无重复"
        - 输出两个Excel文件：
            1. 完整文件：包含所有行和所有备注
            2. 重复文件：只包含有重复的行（备注不是"无重复"）

    重复场景定义：
        场景1：国家相同 && 企业名称相同 && 网址相同 → "完全一致"
        场景2：国家相同 && 企业名称相同 && 网址不同 → "国家与名称一致，网址不一致"
        场景3：国家相同 && 企业名称不同 && 网址相同 → "国家与网址一致，名称不一致"
        场景4：国家不同 && 企业名称相同 && 网址相同 → "国家不一致，名称与网址一致"
        场景5：国家不同 && 企业名称相同 && 网址不同 → "国家与网址不一致，名称一致"
        场景6：国家不同 && 企业名称不同 && 网址相同 → "国家与名称不一致，网址一致"

    参数
    ----------
    input_path : str | pathlib.Path
        输入 Excel 文件路径（通常为 .xlsx）。
    output_full_path : str | pathlib.Path
        完整输出 Excel 文件路径（包含所有行和备注）。
    output_duplicates_path : str | pathlib.Path
        重复行输出 Excel 文件路径（只包含有重复的行）。
    country_col : str, 默认 "国家"
        作为"国家"的列名。
    company_col : str, 默认 "企业名称"
        作为"企业名称"的列名。
    website_col : str, 默认 "网址"
        作为"网址"的列名。
    note_col : str, 默认 "备注"
        备注列的列名。
    scenario_separator : str, 默认 "；"
        多个场景描述之间的分隔符。

    返回
    ----------
    tuple[pandas.DataFrame, pandas.DataFrame]
        第一个DataFrame是完整数据（包含所有行和备注），
        第二个DataFrame是重复数据（只包含有重复的行）。

    异常
    ----------
    FileNotFoundError
        当 input_path 指向的文件不存在时抛出。
    ValueError
        当 Excel 中缺少必需列时抛出。
    """
    input_path = Path(input_path)
    output_full_path = Path(output_full_path)
    output_duplicates_path = Path(output_duplicates_path)

    if not input_path.is_file():
        raise FileNotFoundError(f"输入文件不存在：{input_path}")

    # 读取 Excel 文件
    df = pd.read_excel(input_path, engine="openpyxl")

    # 验证必需列是否存在
    required_cols = [country_col, company_col, website_col]
    missing_cols = [col for col in required_cols if col not in df.columns]
    if missing_cols:
        raise ValueError(
            f"Excel 中缺少必需列：{missing_cols}，"
            f"当前表头为：{list(df.columns)}"
        )

    # 场景描述映射
    scenario_descriptions = {
        1: "完全一致",
        2: "国家与名称一致，网址不一致",
        3: "国家与网址一致，名称不一致",
        4: "国家不一致，名称与网址一致",
        5: "国家与网址不一致，名称一致",
        6: "国家与名称不一致，网址一致",
    }

    # 初始化备注列
    notes = []

    # 逐行比较
    for idx in range(len(df)):
        current_row = df.iloc[idx]
        current_country = current_row[country_col]
        current_company = current_row[company_col]
        current_website = current_row[website_col]

        # 收集当前行匹配的所有场景
        matched_scenarios = set()

        # 与其他所有行比较
        for other_idx in range(len(df)):
            if idx == other_idx:
                continue  # 跳过自己

            other_row = df.iloc[other_idx]
            other_country = other_row[country_col]
            other_company = other_row[company_col]
            other_website = other_row[website_col]

            # 判断是否相同（处理NaN值）
            country_same = pd.isna(current_country) and pd.isna(other_country) or (
                not pd.isna(current_country)
                and not pd.isna(other_country)
                and str(current_country).strip() == str(other_country).strip()
            )
            company_same = pd.isna(current_company) and pd.isna(other_company) or (
                not pd.isna(current_company)
                and not pd.isna(other_company)
                and str(current_company).strip() == str(other_company).strip()
            )
            website_same = pd.isna(current_website) and pd.isna(other_website) or (
                not pd.isna(current_website)
                and not pd.isna(other_website)
                and str(current_website).strip() == str(other_website).strip()
            )

            # 检测6种场景
            if country_same and company_same and website_same:
                matched_scenarios.add(1)
            elif country_same and company_same and not website_same:
                matched_scenarios.add(2)
            elif country_same and not company_same and website_same:
                matched_scenarios.add(3)
            elif not country_same and company_same and website_same:
                matched_scenarios.add(4)
            elif not country_same and company_same and not website_same:
                matched_scenarios.add(5)
            elif not country_same and not company_same and website_same:
                matched_scenarios.add(6)

        # 生成备注
        if matched_scenarios:
            # 按场景编号排序，确保输出顺序一致
            sorted_scenarios = sorted(matched_scenarios)
            note_text = scenario_separator.join(
                [scenario_descriptions[s] for s in sorted_scenarios]
            )
        else:
            note_text = "无重复"

        notes.append(note_text)

    # 添加备注列
    df[note_col] = notes

    # 创建完整数据副本
    full_df = df.copy()

    # 创建重复数据（备注不是"无重复"的行）
    duplicates_df = df[df[note_col] != "无重复"].copy()

    # 保存完整文件
    full_df.to_excel(output_full_path, index=False)

    # 保存重复文件
    duplicates_df.to_excel(output_duplicates_path, index=False)

    return full_df, duplicates_df


if __name__ == "__main__":
    # 示例1：使用 find_duplicate_rows 函数
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

    # 示例2：使用 detect_duplicate_scenarios 函数
    scenario_input = Path("scenario_input.xlsx")
    scenario_output_full = Path("scenario_output_full.xlsx")
    scenario_output_duplicates = Path("scenario_output_duplicates.xlsx")

    if scenario_input.exists():
        full_df, duplicates_df = detect_duplicate_scenarios(
            input_path=scenario_input,
            output_full_path=scenario_output_full,
            output_duplicates_path=scenario_output_duplicates,
            country_col="国家",
            company_col="企业名称",
            website_col="网址",
            note_col="备注",
        )
        print(
            f"场景检测完成！\n"
            f"  完整文件（{len(full_df)}行）已保存到：{scenario_output_full}\n"
            f"  重复文件（{len(duplicates_df)}行）已保存到：{scenario_output_duplicates}"
        )
    else:
        print(
            f"场景检测示例输入文件 {scenario_input} 不存在。"
            " 请将包含'国家'、'企业名称'、'网址'列的 Excel 文件重命名为 scenario_input.xlsx 放在当前目录后再次运行。"
        )


