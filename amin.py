import os
import re
from docx import Document
import pandas as pd
from tqdm import tqdm
import yaml
import logging
from datetime import datetime

# 配置日志记录
logging.basicConfig(filename='extraction_log.txt', level=logging.INFO,
                    format='%(asctime)s - %(levelname)s - %(message)s')


def load_config(config_path="config.yaml"):
    """
    加载配置文件
    """
    try:
        with open(config_path, 'r', encoding='utf-8') as f:
            config = yaml.safe_load(f)
        logging.info("配置文件加载成功")
        return config
    except Exception as e:
        logging.error(f"配置文件加载失败: {str(e)}")
        raise


def extract_info_from_docx(docx_path, patterns):
    """
    从单个Word文档中提取信息
    """
    try:
        doc = Document(docx_path)
        extracted_data = {}

        # 提取所有文本内容
        full_text = '\n'.join([para.text for para in doc.paragraphs])

        # 使用正则表达式匹配信息
        for field, regex_list in patterns.items():
            value = None
            for pattern in regex_list:
                match = re.search(pattern, full_text, re.IGNORECASE)
                if match:
                    value = match.group(1).strip()
                    break
            extracted_data[field] = value

        # 添加文件名信息
        extracted_data['源文件名'] = os.path.basename(docx_path)
        return extracted_data

    except Exception as e:
        logging.error(f"文件处理失败: {docx_path} - {str(e)}")
        return None


def process_docx_files(config):
    """
    批量处理Word文档
    """
    # 获取配置参数
    source_folder = config['source_folder']
    output_file = config['output_file']
    patterns = config['extraction_patterns']
    header_row = config.get('header_row', 0)

    # 获取所有Word文档
    docx_files = [f for f in os.listdir(source_folder)
                  if f.lower().endswith('.docx')]

    if not docx_files:
        logging.warning("未找到任何Word文档")
        return

    # 处理文件
    results = []
    success_count = 0
    fail_count = 0

    print(f"开始处理 {len(docx_files)} 个文档...")
    for filename in tqdm(docx_files, desc="处理进度"):
        file_path = os.path.join(source_folder, filename)
        result = extract_info_from_docx(file_path, patterns)

        if result:
            results.append(result)
            success_count += 1
        else:
            fail_count += 1

    # 创建DataFrame并保存
    if results:
        df = pd.DataFrame(results)

        # 清除空行
        df.dropna(how='all', inplace=True)

        # 保存到Excel
        df.to_excel(output_file, index=False)
        print(f"\n处理完成! 成功: {success_count}, 失败: {fail_count}")
        print(f"结果已保存到: {output_file}")
        logging.info(f"成功处理 {success_count} 个文件，失败 {fail_count} 个")
    else:
        print("未提取到有效数据")
        logging.warning("未提取到有效数据")


if __name__ == "__main__":
    try:
        config = load_config()
        process_docx_files(config)
    except Exception as e:
        print(f"程序执行出错: {str(e)}")
        logging.critical(f"程序终止: {str(e)}")