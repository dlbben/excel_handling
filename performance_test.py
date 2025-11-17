"""
性能测试脚本：测试 detect_duplicate_scenarios 函数处理十万行数据的性能
"""

import time
import pandas as pd
import numpy as np
from pathlib import Path
from excel_duplicates import detect_duplicate_scenarios


def generate_test_data(num_rows: int = 100000, output_path: str = "test_data_100k.xlsx"):
    """
    生成测试数据
    
    参数
    ----------
    num_rows : int
        要生成的行数
    output_path : str
        输出Excel文件路径
    """
    print(f"正在生成 {num_rows:,} 行测试数据...")
    
    # 生成测试数据
    np.random.seed(42)  # 固定随机种子，确保可重复
    
    # 生成国家列表（约20个国家）
    countries = [f"国家{i}" for i in range(1, 21)]
    
    # 生成企业名称列表（约5000个企业）
    companies = [f"企业{i}" for i in range(1, 5001)]
    
    # 生成网址列表（约10000个网址）
    websites = [f"https://www.example{i}.com" for i in range(1, 10001)]
    
    # 随机生成数据，确保有重复场景
    data = {
        "国家": np.random.choice(countries, num_rows),
        "企业名称": np.random.choice(companies, num_rows),
        "网址": np.random.choice(websites, num_rows),
        "其他列1": [f"数据{i}" for i in range(num_rows)],
        "其他列2": np.random.randint(1, 1000, num_rows),
    }
    
    df = pd.DataFrame(data)
    
    # 保存为Excel文件
    df.to_excel(output_path, index=False)
    print(f"测试数据已保存到: {output_path}")
    print(f"数据形状: {df.shape}")
    print(f"国家唯一值数量: {df['国家'].nunique()}")
    print(f"企业名称唯一值数量: {df['企业名称'].nunique()}")
    print(f"网址唯一值数量: {df['网址'].nunique()}")
    
    return output_path


def run_performance_test(
    input_path: str,
    num_runs: int = 3,
    target_time: float = 60.0
):
    """
    运行性能测试
    
    参数
    ----------
    input_path : str
        输入Excel文件路径
    num_runs : int
        运行次数（取平均值）
    target_time : float
        目标时间（秒）
    """
    print(f"\n{'='*60}")
    print(f"开始性能测试")
    print(f"{'='*60}")
    print(f"输入文件: {input_path}")
    print(f"运行次数: {num_runs}")
    print(f"目标时间: {target_time} 秒")
    print(f"{'='*60}\n")
    
    times = []
    
    for run in range(1, num_runs + 1):
        print(f"第 {run}/{num_runs} 次运行...")
        
        # 输出文件路径
        output_full = f"test_output_full_run{run}.xlsx"
        output_duplicates = f"test_output_duplicates_run{run}.xlsx"
        
        # 记录开始时间
        start_time = time.time()
        
        try:
            # 运行函数
            full_df, duplicates_df = detect_duplicate_scenarios(
                input_path=input_path,
                output_full_path=output_full,
                output_duplicates_path=output_duplicates,
                country_col="国家",
                company_col="企业名称",
                website_col="网址",
                note_col="备注",
            )
            
            # 记录结束时间
            end_time = time.time()
            elapsed_time = end_time - start_time
            times.append(elapsed_time)
            
            print(f"  运行时间: {elapsed_time:.2f} 秒")
            print(f"  完整文件行数: {len(full_df):,}")
            print(f"  重复文件行数: {len(duplicates_df):,}")
            
            # 统计备注分布
            note_counts = full_df["备注"].value_counts()
            print(f"  备注类型数量: {len(note_counts)}")
            print(f"  '无重复'行数: {note_counts.get('无重复', 0):,}")
            
        except Exception as e:
            print(f"  错误: {e}")
            continue
    
    if times:
        avg_time = np.mean(times)
        min_time = np.min(times)
        max_time = np.max(times)
        
        print(f"\n{'='*60}")
        print(f"性能测试结果")
        print(f"{'='*60}")
        print(f"平均时间: {avg_time:.2f} 秒")
        print(f"最短时间: {min_time:.2f} 秒")
        print(f"最长时间: {max_time:.2f} 秒")
        print(f"目标时间: {target_time:.2f} 秒")
        print(f"{'='*60}")
        
        if avg_time <= target_time:
            print(f"✅ 性能达标！平均时间 ({avg_time:.2f}秒) 小于目标时间 ({target_time:.2f}秒)")
        else:
            print(f"❌ 性能未达标！平均时间 ({avg_time:.2f}秒) 大于目标时间 ({target_time:.2f}秒)")
            print(f"   需要优化 {(avg_time - target_time):.2f} 秒")
        
        # 计算速度
        input_df = pd.read_excel(input_path, engine="openpyxl")
        num_rows = len(input_df)
        rows_per_second = num_rows / avg_time
        print(f"\n处理速度: {rows_per_second:,.0f} 行/秒")
        print(f"数据规模: {num_rows:,} 行")
        
        return avg_time, min_time, max_time
    else:
        print("所有运行都失败了")
        return None, None, None


def main():
    """主函数"""
    print("=" * 60)
    print("Excel 重复场景检测 - 性能测试")
    print("=" * 60)
    
    # 测试数据文件路径
    test_data_path = "test_data_100k.xlsx"
    num_rows = 100000
    
    # 检查测试数据是否存在
    if not Path(test_data_path).exists():
        print(f"\n测试数据文件不存在，正在生成...")
        generate_test_data(num_rows=num_rows, output_path=test_data_path)
    else:
        # 检查现有文件的行数
        try:
            existing_df = pd.read_excel(test_data_path, engine="openpyxl", nrows=5)
            print(f"\n发现现有测试数据文件: {test_data_path}")
            print("使用现有测试数据文件（如需重新生成，请删除该文件后重新运行）")
        except:
            print(f"\n现有测试数据文件可能损坏，正在重新生成...")
            generate_test_data(num_rows=num_rows, output_path=test_data_path)
    
    # 运行性能测试
    target_time = 60.0  # 目标：1分钟内完成
    run_performance_test(
        input_path=test_data_path,
        num_runs=3,
        target_time=target_time
    )
    
    print("\n测试完成！")


if __name__ == "__main__":
    main()
