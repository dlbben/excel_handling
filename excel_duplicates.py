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

    # 预处理：标准化数据（处理NaN和空格）- 使用向量化操作
    # 直接在df上操作，避免不必要的复制
    df["_country_norm"] = df[country_col].fillna("").astype(str).str.strip()
    df["_company_norm"] = df[company_col].fillna("").astype(str).str.strip()
    df["_website_norm"] = df[website_col].fillna("").astype(str).str.strip()

    # 使用numpy数组存储场景标记（位标记：场景1=1, 场景2=2, 场景3=4, 场景4=8, 场景5=16, 场景6=32）
    import numpy as np
    scenario_flags = np.zeros(len(df), dtype=np.int32)

    # 场景1：国家相同 && 企业名称相同 && 网址相同
    # 使用transform批量标记，避免循环
    key_cols = ["_country_norm", "_company_norm", "_website_norm"]
    grouped = df.groupby(key_cols, dropna=False, sort=False)
    # 使用transform将组大小映射到每行
    df["_size_1"] = grouped["_country_norm"].transform("size")
    scenario_flags |= ((df["_size_1"] > 1).astype(np.int32) * 1)

    # 场景2：国家相同 && 企业名称相同 && 网址不同
    key_cols = ["_country_norm", "_company_norm"]
    grouped = df.groupby(key_cols, dropna=False, sort=False)
    df["_size_2"] = grouped["_country_norm"].transform("size")
    df["_website_nunique_2"] = grouped["_website_norm"].transform("nunique")
    scenario_flags |= ((df["_size_2"] > 1) & (df["_website_nunique_2"] > 1)).astype(np.int32) * 2

    # 场景3：国家相同 && 企业名称不同 && 网址相同
    key_cols = ["_country_norm", "_website_norm"]
    grouped = df.groupby(key_cols, dropna=False, sort=False)
    df["_size_3"] = grouped["_country_norm"].transform("size")
    df["_company_nunique_3"] = grouped["_company_norm"].transform("nunique")
    scenario_flags |= ((df["_size_3"] > 1) & (df["_company_nunique_3"] > 1)).astype(np.int32) * 4

    # 场景4：国家不同 && 企业名称相同 && 网址相同
    key_cols = ["_company_norm", "_website_norm"]
    grouped = df.groupby(key_cols, dropna=False, sort=False)
    df["_size_4"] = grouped["_company_norm"].transform("size")
    df["_country_nunique_4"] = grouped["_country_norm"].transform("nunique")
    scenario_flags |= ((df["_size_4"] > 1) & (df["_country_nunique_4"] > 1)).astype(np.int32) * 8

    # 场景5：国家不同 && 企业名称相同 && 网址不同
    key_cols = ["_company_norm"]
    grouped = df.groupby(key_cols, dropna=False, sort=False)
    df["_size_5"] = grouped["_company_norm"].transform("size")
    df["_country_nunique_5"] = grouped["_country_norm"].transform("nunique")
    df["_website_nunique_5"] = grouped["_website_norm"].transform("nunique")
    scenario_flags |= ((df["_size_5"] > 1) & (df["_country_nunique_5"] > 1) & (df["_website_nunique_5"] > 1)).astype(np.int32) * 16

    # 场景6：国家不同 && 企业名称不同 && 网址相同
    key_cols = ["_website_norm"]
    grouped = df.groupby(key_cols, dropna=False, sort=False)
    df["_size_6"] = grouped["_website_norm"].transform("size")
    df["_country_nunique_6"] = grouped["_country_norm"].transform("nunique")
    df["_company_nunique_6"] = grouped["_company_norm"].transform("nunique")
    scenario_flags |= ((df["_size_6"] > 1) & (df["_country_nunique_6"] > 1) & (df["_company_nunique_6"] > 1)).astype(np.int32) * 32

    # 清理临时列
    temp_cols = [
        "_country_norm", "_company_norm", "_website_norm",
        "_size_1", "_size_2", "_website_nunique_2", "_size_3", "_company_nunique_3",
        "_size_4", "_country_nunique_4", "_size_5", "_country_nunique_5", "_website_nunique_5",
        "_size_6", "_country_nunique_6", "_company_nunique_6"
    ]
    df.drop(columns=temp_cols, errors="ignore", inplace=True)

    # 预编译场景描述映射和位标记
    scenario_bits = np.array([1, 2, 4, 8, 16, 32], dtype=np.int32)
    scenario_nums = np.array([1, 2, 3, 4, 5, 6], dtype=np.int32)
    
    # 使用向量化操作生成备注（优化字符串生成）
    # 预编译所有可能的备注组合（最多63种组合：2^6-1）
    note_cache = {}
    note_cache[0] = "无重复"
    
    def flags_to_note(flag):
        """将标志位转换为备注文本（带缓存，使用numpy加速）"""
        if flag in note_cache:
            return note_cache[flag]
        if flag == 0:
            return "无重复"
        # 使用numpy位运算提取被标记的场景编号（向量化）
        matched = scenario_nums[scenario_bits & flag != 0]
        if len(matched) > 0:
            # 排序并生成备注文本
            matched_sorted = np.sort(matched)
            note_text = scenario_separator.join([scenario_descriptions[int(s)] for s in matched_sorted])
            note_cache[flag] = note_text
            return note_text
        return "无重复"
    
    # 优化：先预编译所有唯一的标志值，然后批量生成备注
    # 这样可以最大化缓存命中率
    unique_flags = np.unique(scenario_flags)
    for flag in unique_flags:
        if flag not in note_cache:
            flags_to_note(int(flag))
    
    # 使用pandas的map方法，比apply稍快，且可以利用缓存
    notes = pd.Series(scenario_flags).map(note_cache).values

    # 添加备注列
    df[note_col] = notes

    # 创建完整数据副本（避免修改原始df）
    full_df = df.copy()

    # 创建重复数据（备注不是"无重复"的行）
    # 使用布尔索引，更高效
    duplicates_mask = df[note_col] != "无重复"
    duplicates_df = df[duplicates_mask].copy()

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


