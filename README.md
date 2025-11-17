## Excel 重复行检测工具说明

本项目提供基于 Python 的 Excel 重复行检测工具，包含两种功能：

1. **基础重复检测**：按"名字 + 链接"识别重复行
2. **多场景重复检测**：检测6种不同的重复场景（国家、企业名称、网址的组合）

### 1. 环境准备

#### 1.1 检查 Python 版本

首先确认已安装 Python 3.8 或更高版本（建议 64 位，内存 ≥ 8GB 处理 50 万行以上数据更稳定）。

在命令行中运行：

```bash
python --version
```

或

```bash
python3 --version
```

如果未安装 Python，请前往 [Python 官网](https://www.python.org/downloads/) 下载并安装。

#### 1.2 创建虚拟环境（推荐）

为了避免与系统 Python 环境冲突，建议创建独立的虚拟环境：

**步骤 1：创建虚拟环境**

在项目根目录下打开命令行（PowerShell 或 CMD），运行：

```bash
python -m venv venv
```

这会在当前目录创建一个名为 `venv` 的虚拟环境文件夹。

**步骤 2：激活虚拟环境**

根据你的操作系统，使用对应的命令激活虚拟环境：

**Windows（PowerShell 或 CMD）：**

```bash
venv\Scripts\activate
```

**Windows（Git Bash）：**

```bash
source venv/Scripts/activate
```

**Linux/Mac：**

```bash
source venv/bin/activate
```

**如何判断虚拟环境已激活？**

激活成功后，命令行提示符前面会显示 `(venv)`，例如：

```
(venv) C:\git\excel_handling\excel_handling>
```

**退出虚拟环境：**

当不再需要虚拟环境时，可以运行以下命令退出：

```bash
deactivate
```

#### 1.3 安装项目依赖

在项目根目录（包含 `excel_duplicates.py` 和 `requirements.txt` 的目录）下，使用以下命令安装依赖：

**方式一：使用默认 PyPI 源（国外源，速度可能较慢）**

```bash
pip install -r requirements.txt
```

**方式二：使用国内镜像源（推荐，下载速度更快）**

使用清华大学镜像源：

```bash
pip install -r requirements.txt -i https://pypi.tuna.tsinghua.edu.cn/simple
```

其他常用国内镜像源：

- 阿里云：`https://mirrors.aliyun.com/pypi/simple/`
- 中科大：`https://pypi.mirrors.ustc.edu.cn/simple/`
- 豆瓣：`https://pypi.douban.com/simple/`

例如使用阿里云镜像源：

```bash
pip install -r requirements.txt -i https://mirrors.aliyun.com/pypi/simple/
```

#### 1.4 验证安装

安装完成后，可以运行以下命令验证依赖是否安装成功：

```bash
python -c "import pandas; import openpyxl; print('依赖安装成功！')"
```

如果看到"依赖安装成功！"的输出，说明环境已准备就绪。

### 2. 核心功能介绍

核心代码位于 `excel_duplicates.py`，对外提供两个主要函数：

#### 2.1 find_duplicate_rows - 基础重复检测

```python
from excel_duplicates import find_duplicate_rows
```

函数的作用：

- 读取一个 Excel 文件（例如包含列：`名字`、`邮箱`、`链接`、`国家`）。
- 找出所有"名字相等且链接相等"的重复行（出现 2 次及以上的组合）。
- 将这些重复行写入到新的 Excel 文件中。

#### 2.2 detect_duplicate_scenarios - 多场景重复检测

```python
from excel_duplicates import detect_duplicate_scenarios
```

函数的作用：

- 读取一个 Excel 文件（包含列：`国家`、`企业名称`、`网址`等）。
- 逐行与其他所有行比较，检测6种重复场景。
- 为每行添加"备注"列，标记所有匹配的场景。
- 输出两个 Excel 文件：
  - 完整文件：包含所有行和所有备注
  - 重复文件：只包含有重复的行（备注不是"无重复"）

**6种重复场景：**

- **场景1**：国家相同 && 企业名称相同 && 网址相同 → 备注："完全一致"
- **场景2**：国家相同 && 企业名称相同 && 网址不同 → 备注："国家与名称一致，网址不一致"
- **场景3**：国家相同 && 企业名称不同 && 网址相同 → 备注："国家与网址一致，名称不一致"
- **场景4**：国家不同 && 企业名称相同 && 网址相同 → 备注："国家不一致，名称与网址一致"
- **场景5**：国家不同 && 企业名称相同 && 网址不同 → 备注："国家与网址不一致，名称一致"
- **场景6**：国家不同 && 企业名称不同 && 网址相同 → 备注："国家与名称不一致，网址一致"

如果一行与多行比较时匹配多个场景，备注会用分隔符（默认"；"）连接所有场景描述。如果一行没有匹配任何场景，备注标记为"无重复"。

### 3. 在 Python 代码中调用示例

#### 3.1 基础重复检测示例

假设你的原始 Excel 名为 `data.xlsx`，位于当前目录，希望把所有重复行导出到 `duplicates.xlsx`：

```python
from excel_duplicates import find_duplicate_rows


def main():
    dup_df = find_duplicate_rows(
        input_path="data.xlsx",        # 原始 Excel 路径
        output_path="duplicates.xlsx", # 输出重复行 Excel 路径
        name_col="名字",               # "名字"列名，如你的表头不同可在这里修改
        link_col="链接",               # "链接"列名
        extra_columns=["邮箱", "国家"], # 额外需要保留的列，可按需增删
        write_empty=True,              # 若没有重复行时是否仍然生成空的 Excel
    )

    print(f"共找到重复行：{len(dup_df)} 条")


if __name__ == "__main__":
    main()
```

#### 3.2 多场景重复检测示例

假设你的 Excel 文件包含"国家"、"企业名称"、"网址"等列，需要检测6种重复场景：

```python
from excel_duplicates import detect_duplicate_scenarios


def main():
    full_df, duplicates_df = detect_duplicate_scenarios(
        input_path="companies.xlsx",              # 原始 Excel 路径
        output_full_path="companies_full.xlsx",   # 完整输出文件（包含所有行和备注）
        output_duplicates_path="companies_duplicates.xlsx",  # 重复行输出文件
        country_col="国家",                        # "国家"列名
        company_col="企业名称",                    # "企业名称"列名
        website_col="网址",                        # "网址"列名
        note_col="备注",                           # 备注列名
        scenario_separator="；",                   # 多个场景描述之间的分隔符
    )

    print(f"完整文件包含 {len(full_df)} 行")
    print(f"重复文件包含 {len(duplicates_df)} 行")


if __name__ == "__main__":
    main()
```

运行方式：

```bash
python your_script.py
```

### 4. 直接运行模块内示例（可选）

`excel_duplicates.py` 中已经包含两个示例，用于快速测试：

#### 4.1 基础重复检测示例

1. 将你的 Excel 文件重命名为 `input.xlsx` 并放到与 `excel_duplicates.py` 同一目录。
2. 在该目录下运行：

```bash
python excel_duplicates.py
```

如果文件存在，脚本会：

- 读取 `input.xlsx`；
- 按 `名字 + 链接` 查找重复行；
- 将结果写出到 `duplicates.xlsx`；
- 在控制台打印重复行数量。

#### 4.2 多场景重复检测示例

1. 将包含"国家"、"企业名称"、"网址"列的 Excel 文件重命名为 `scenario_input.xlsx` 并放到与 `excel_duplicates.py` 同一目录。
2. 在该目录下运行：

```bash
python excel_duplicates.py
```

如果文件存在，脚本会：

- 读取 `scenario_input.xlsx`；
- 检测6种重复场景并为每行添加备注；
- 将完整结果写出到 `scenario_output_full.xlsx`（包含所有行和备注）；
- 将重复行结果写出到 `scenario_output_duplicates.xlsx`（只包含有重复的行）；
- 在控制台打印两个文件的统计信息。

### 5. 参数说明

#### 5.1 find_duplicate_rows 函数参数

- **input_path**：输入 Excel 文件路径（字符串或 `Path` 对象）。
- **output_path**：输出重复行 Excel 的路径。
- **name_col**：作为"名字"的列名，默认 `"名字"`。
- **link_col**：作为"链接"的列名，默认 `"链接"`。
- **extra_columns**：可选，想要一起读取、保留的其他列名列表，例如 `["邮箱", "国家"]`。
- **write_empty**：可选，默认 `True`。若没有重复行：
  - 为 `True`：仍然创建一个（可能为空的）输出 Excel 文件；
  - 为 `False`：不写文件，只返回空的 DataFrame。

#### 5.2 detect_duplicate_scenarios 函数参数

- **input_path**：输入 Excel 文件路径（字符串或 `Path` 对象）。
- **output_full_path**：完整输出 Excel 文件路径（包含所有行和备注）。
- **output_duplicates_path**：重复行输出 Excel 文件路径（只包含有重复的行）。
- **country_col**：可选，作为"国家"的列名，默认 `"国家"`。
- **company_col**：可选，作为"企业名称"的列名，默认 `"企业名称"`。
- **website_col**：可选，作为"网址"的列名，默认 `"网址"`。
- **note_col**：可选，备注列的列名，默认 `"备注"`。
- **scenario_separator**：可选，多个场景描述之间的分隔符，默认 `"；"`。

### 6. 注意事项

- 如需适配你自己的 Excel 表头，只需要在调用函数时把相应的列名参数改成对应的列名即可。
- 对于大数据量（如超过10万行），`detect_duplicate_scenarios` 函数可能需要较长时间，因为它是逐行比较（O(n²)复杂度）。
- 函数会自动处理空值（NaN）和字符串前后空格的情况。
- 如果一行与多行比较时匹配多个场景，备注会用分隔符连接所有场景描述。


