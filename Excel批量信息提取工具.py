"""
Excel 批量信息提取工具
作者：xiaoshu
功能：按配置批量提取多个 Excel 文件的指定列，生成汇总文件及处理报告
"""

import pandas as pd
import yaml
import os
from glob import glob
from tqdm import tqdm

def normalize_column_name(name: str) -> str:
    """
    标准化列名：小写 + 移除空格和特殊符号
    """
    return name.lower().replace(" ", "").replace("_", "").replace("-", "")

def load_config(config_path: str) -> dict:
    """
    读取 YAML 配置文件
    """
    with open(config_path, "r", encoding="utf-8") as f:
        return yaml.safe_load(f)

def main():
    # 读取配置文件
    config = load_config("config.yaml")
    source_dir = config["source_dir"]
    output_path = config["output_path"]
    columns_config = config["columns"]
    header_row = config["header_row"] - 1  # 转换为 pandas 的 0-based 索引

    # 解析列名配置
    target_columns = []
    aliases = {}
    for col_entry in columns_config:
        for target, alias_list in col_entry.items():
            target_columns.append(target)
            aliases[target] = [target] + alias_list

    # 检查源文件夹是否存在
    if not os.path.exists(source_dir):
        raise FileNotFoundError(f"源文件夹不存在：{source_dir}")

    # 准备输出目录
    output_dir = os.path.dirname(output_path)
    if not os.path.exists(output_dir):
        os.makedirs(output_dir)

    # 获取所有 Excel 文件
    files = glob(os.path.join(source_dir, "*.xlsx"))
    if not files:
        print("警告：未找到任何 .xlsx 文件！")
        return

    # 初始化数据容器
    summary_data = pd.DataFrame()
    process_report = []

    # 处理每个文件（带进度条）
    for file_path in tqdm(files, desc="正在处理文件"):
        try:
            # 读取 Excel 文件
            df = pd.read_excel(file_path, header=header_row)
            if df.empty:
                raise ValueError("表格为空或表头行不存在")

            # 构建列名映射表
            column_mapping = {}
            raw_columns = df.columns.tolist()
            normalized_raw = {col: normalize_column_name(col) for col in raw_columns}

            # 为每个目标列查找匹配的原始列
            for target in target_columns:
                found = False
                for alias in aliases[target]:
                    normalized_alias = normalize_column_name(alias)
                    for raw_col in raw_columns:
                        if normalized_alias == normalize_column_name(raw_col):
                            column_mapping[target] = raw_col
                            found = True
                            break
                    if found:
                        break

            # 提取数据并处理缺失列
            extracted = pd.DataFrame()
            for target in target_columns:
                if target in column_mapping:
                    extracted[target] = df[column_mapping[target]]
                else:
                    extracted[target] = pd.NA  # 缺失列填充空值

            # 清除当前文件中的全空行
            extracted.dropna(how="all", inplace=True)

            # 合并到总数据
            summary_data = pd.concat([summary_data, extracted], ignore_index=True)

            # 记录处理状态
            process_report.append({
                "文件名": os.path.basename(file_path),
                "状态": "成功",
                "错误信息": ""
            })

        except Exception as e:
            process_report.append({
                "文件名": os.path.basename(file_path),
                "状态": "失败",
                "错误信息": str(e)
            })

    # 清除所有空行后保存汇总文件
    summary_data.dropna(how="all", inplace=True)
    summary_data.to_excel(output_path, index=False)
    print(f"汇总文件已保存至：{output_path}")

    # 生成处理报告
    report_df = pd.DataFrame(process_report)
    report_path = os.path.join(output_dir, "处理报告.xlsx")
    report_df.to_excel(report_path, index=False)
    print(f"处理报告已生成：{report_path}")

if __name__ == "__main__":
    main()