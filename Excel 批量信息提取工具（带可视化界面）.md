# Excel 批量信息提取工具（带可视化界面）

下面我将为之前的Excel批量信息提取工具添加一个完整的可视化界面，使用PyQt5实现。这个界面将包含所有必要的功能，让用户无需手动编辑配置文件即可完成操作。

## 环境安装

首先需要安装额外的依赖库：
```bash
pip install pandas openpyxl tqdm pyyaml PyQt5
```

## 完整代码实现

```python
"""
Excel 批量信息提取工具 (带可视化界面)
作者：AI助手
功能：通过图形界面批量提取多个Excel文件的指定列数据
"""

import sys
import os
import pandas as pd
import yaml
from glob import glob
from tqdm import tqdm
from PyQt5.QtWidgets import (
    QApplication, QMainWindow, QWidget, QVBoxLayout, QHBoxLayout, QLabel,
    QLineEdit, QPushButton, QFileDialog, QTableWidget, QTableWidgetItem,
    QHeaderView, QMessageBox, QProgressBar, QComboBox, QSpinBox, QCheckBox
)
from PyQt5.QtCore import Qt, QThread, pyqtSignal

def normalize_column_name(name: str) -> str:
    """标准化列名：小写 + 移除空格和特殊符号"""
    return str(name).lower().replace(" ", "").replace("_", "").replace("-", "")

class ExcelExtractorThread(QThread):
    """后台处理线程"""
    progress = pyqtSignal(int, str)  # 进度值, 当前文件名
    finished = pyqtSignal(bool, str)  # 是否成功, 消息
    report_ready = pyqtSignal(pd.DataFrame)  # 处理报告

    def __init__(self, config, parent=None):
        super().__init__(parent)
        self.config = config
        self.report_df = None

    def run(self):
        try:
            source_dir = self.config["source_dir"]
            output_path = self.config["output_path"]
            columns_config = self.config["columns"]
            header_row = self.config["header_row"] - 1  # 转换为0-based索引
            remove_empty = self.config.get("remove_empty", True)

            # 检查源文件夹
            if not os.path.exists(source_dir):
                self.finished.emit(False, f"源文件夹不存在：{source_dir}")
                return

            # 准备输出目录
            output_dir = os.path.dirname(output_path)
            if output_dir and not os.path.exists(output_dir):
                os.makedirs(output_dir)

            # 获取所有Excel文件
            files = glob(os.path.join(source_dir, "*.xlsx"))
            if not files:
                self.finished.emit(False, "警告：未找到任何 .xlsx 文件！")
                return

            # 解析目标列配置
            target_columns = []
            aliases = {}
            for col_entry in columns_config:
                for target, alias_list in col_entry.items():
                    target_columns.append(target)
                    aliases[target] = [target] + alias_list

            # 初始化数据容器
            summary_data = pd.DataFrame()
            process_report = []

            # 处理每个文件
            total_files = len(files)
            for i, file_path in enumerate(files):
                file_name = os.path.basename(file_path)
                self.progress.emit(int((i+1)/total_files*100), file_name)
                
                try:
                    # 读取Excel文件
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
                    if remove_empty:
                        extracted.dropna(how="all", inplace=True)

                    # 合并到总数据
                    summary_data = pd.concat([summary_data, extracted], ignore_index=True)

                    # 记录处理状态
                    process_report.append({
                        "文件名": file_name,
                        "状态": "成功",
                        "错误信息": ""
                    })

                except Exception as e:
                    process_report.append({
                        "文件名": file_name,
                        "状态": "失败",
                        "错误信息": str(e)
                    })

            # 清除所有空行后保存汇总文件
            if remove_empty:
                summary_data.dropna(how="all", inplace=True)
                
            summary_data.to_excel(output_path, index=False)
            
            # 生成处理报告
            self.report_df = pd.DataFrame(process_report)
            report_path = os.path.join(os.path.dirname(output_path), "处理报告.xlsx")
            self.report_df.to_excel(report_path, index=False)
            
            self.report_ready.emit(self.report_df)
            self.finished.emit(True, f"处理完成！成功处理 {len([r for r in process_report if r['状态']=='成功'])}/{len(files)} 个文件\n汇总文件已保存至：{output_path}")

        except Exception as e:
            self.finished.emit(False, f"处理过程中发生错误：{str(e)}")


class ExcelExtractorUI(QMainWindow):
    """Excel批量信息提取工具主界面"""
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Excel批量信息提取工具")
        self.setGeometry(100, 100, 900, 700)
        
        # 创建主控件和布局
        main_widget = QWidget()
        main_layout = QVBoxLayout()
        
        # 1. 源文件夹选择
        source_layout = QHBoxLayout()
        source_layout.addWidget(QLabel("源文件夹:"))
        self.source_input = QLineEdit()
        self.source_input.setPlaceholderText("请选择包含Excel文件的文件夹")
        source_layout.addWidget(self.source_input)
        
        self.source_btn = QPushButton("浏览...")
        self.source_btn.clicked.connect(self.select_source_folder)
        source_layout.addWidget(self.source_btn)
        
        # 2. 输出文件选择
        output_layout = QHBoxLayout()
        output_layout.addWidget(QLabel("输出文件:"))
        self.output_input = QLineEdit()
        self.output_input.setPlaceholderText("请选择汇总结果保存位置")
        output_layout.addWidget(self.output_input)
        
        self.output_btn = QPushButton("浏览...")
        self.output_btn.clicked.connect(self.select_output_file)
        output_layout.addWidget(self.output_btn)
        
        # 3. 表头设置
        header_layout = QHBoxLayout()
        header_layout.addWidget(QLabel("表头所在行:"))
        
        self.header_spin = QSpinBox()
        self.header_spin.setMinimum(1)
        self.header_spin.setMaximum(20)
        self.header_spin.setValue(1)
        header_layout.addWidget(self.header_spin)
        
        header_layout.addStretch()
        
        self.remove_empty_check = QCheckBox("自动清除空行")
        self.remove_empty_check.setChecked(True)
        header_layout.addWidget(self.remove_empty_check)
        
        # 4. 列配置表格
        column_layout = QVBoxLayout()
        column_layout.addWidget(QLabel("列配置:"))
        
        self.column_table = QTableWidget()
        self.column_table.setColumnCount(2)
        self.column_table.setHorizontalHeaderLabels(["目标列名", "别名列表(逗号分隔)"])
        self.column_table.horizontalHeader().setSectionResizeMode(QHeaderView.Stretch)
        self.column_table.setRowCount(3)
        
        # 设置默认值
        for i, (col, aliases) in enumerate([
            ("学号", "student number,学生编号,学号"),
            ("姓名", "Name,学生姓名,姓名"),
            ("参加模块", "Participation in modules,参加模块,课程名称")
        ]):
            self.column_table.setItem(i, 0, QTableWidgetItem(col))
            self.column_table.setItem(i, 1, QTableWidgetItem(aliases))
        
        column_layout.addWidget(self.column_table)
        
        # 表格操作按钮
        table_btn_layout = QHBoxLayout()
        self.add_row_btn = QPushButton("添加行")
        self.add_row_btn.clicked.connect(self.add_table_row)
        table_btn_layout.addWidget(self.add_row_btn)
        
        self.del_row_btn = QPushButton("删除行")
        self.del_row_btn.clicked.connect(self.delete_table_row)
        table_btn_layout.addWidget(self.del_row_btn)
        
        table_btn_layout.addStretch()
        column_layout.addLayout(table_btn_layout)
        
        # 5. 进度条
        progress_layout = QVBoxLayout()
        progress_layout.addWidget(QLabel("进度:"))
        
        self.progress_bar = QProgressBar()
        self.progress_bar.setRange(0, 100)
        self.progress_bar.setValue(0)
        progress_layout.addWidget(self.progress_bar)
        
        self.status_label = QLabel("就绪")
        self.status_label.setAlignment(Qt.AlignCenter)
        progress_layout.addWidget(self.status_label)
        
        # 6. 操作按钮
        btn_layout = QHBoxLayout()
        self.start_btn = QPushButton("开始处理")
        self.start_btn.clicked.connect(self.start_processing)
        btn_layout.addWidget(self.start_btn)
        
        self.export_report_btn = QPushButton("导出报告")
        self.export_report_btn.clicked.connect(self.export_report)
        self.export_report_btn.setEnabled(False)
        btn_layout.addWidget(self.export_report_btn)
        
        self.exit_btn = QPushButton("退出")
        self.exit_btn.clicked.connect(self.close)
        btn_layout.addWidget(self.exit_btn)
        
        # 组装主布局
        main_layout.addLayout(source_layout)
        main_layout.addLayout(output_layout)
        main_layout.addLayout(header_layout)
        main_layout.addLayout(column_layout)
        main_layout.addLayout(progress_layout)
        main_layout.addLayout(btn_layout)
        
        main_widget.setLayout(main_layout)
        self.setCentralWidget(main_widget)
        
        # 后台处理线程
        self.worker_thread = None
        self.report_df = None

    def select_source_folder(self):
        """选择源文件夹"""
        folder = QFileDialog.getExistingDirectory(self, "选择Excel文件所在文件夹")
        if folder:
            self.source_input.setText(folder)

    def select_output_file(self):
        """选择输出文件"""
        file_path, _ = QFileDialog.getSaveFileName(
            self, "保存汇总文件", "", "Excel Files (*.xlsx)"
        )
        if file_path:
            if not file_path.endswith(".xlsx"):
                file_path += ".xlsx"
            self.output_input.setText(file_path)

    def add_table_row(self):
        """添加新的列配置行"""
        row_count = self.column_table.rowCount()
        self.column_table.insertRow(row_count)

    def delete_table_row(self):
        """删除选中的列配置行"""
        current_row = self.column_table.currentRow()
        if current_row >= 0:
            self.column_table.removeRow(current_row)

    def get_config(self):
        """从界面获取配置"""
        config = {
            "source_dir": self.source_input.text().strip(),
            "output_path": self.output_input.text().strip(),
            "header_row": self.header_spin.value(),
            "remove_empty": self.remove_empty_check.isChecked(),
            "columns": []
        }
        
        # 获取列配置
        for row in range(self.column_table.rowCount()):
            target_item = self.column_table.item(row, 0)
            aliases_item = self.column_table.item(row, 1)
            
            if target_item and target_item.text().strip():
                target = target_item.text().strip()
                aliases = [a.strip() for a in aliases_item.text().split(",")] if aliases_item else []
                config["columns"].append({target: aliases})
        
        return config

    def validate_config(self, config):
        """验证配置是否有效"""
        if not config["source_dir"]:
            QMessageBox.warning(self, "配置错误", "请选择源文件夹！")
            return False
            
        if not config["output_path"]:
            QMessageBox.warning(self, "配置错误", "请选择输出文件路径！")
            return False
            
        if not config["columns"]:
            QMessageBox.warning(self, "配置错误", "请至少配置一列提取信息！")
            return False
            
        return True

    def start_processing(self):
        """开始处理Excel文件"""
        config = self.get_config()
        if not self.validate_config(config):
            return
            
        # 禁用按钮，防止重复点击
        self.start_btn.setEnabled(False)
        self.progress_bar.setValue(0)
        self.status_label.setText("正在初始化...")
        
        # 创建并启动工作线程
        self.worker_thread = ExcelExtractorThread(config)
        self.worker_thread.progress.connect(self.update_progress)
        self.worker_thread.finished.connect(self.on_finished)
        self.worker_thread.report_ready.connect(self.on_report_ready)
        self.worker_thread.start()

    def update_progress(self, value, filename):
        """更新进度条状态"""
        self.progress_bar.setValue(value)
        self.status_label.setText(f"正在处理: {filename}")

    def on_report_ready(self, report_df):
        """保存处理报告"""
        self.report_df = report_df
        self.export_report_btn.setEnabled(True)

    def on_finished(self, success, message):
        """处理完成回调"""
        self.start_btn.setEnabled(True)
        
        if success:
            QMessageBox.information(self, "处理完成", message)
        else:
            QMessageBox.critical(self, "处理失败", message)

    def export_report(self):
        """导出处理报告"""
        if self.report_df is None:
            QMessageBox.warning(self, "导出失败", "没有可导出的报告数据")
            return
            
        # 获取默认保存路径（与输出文件同目录）
        default_dir = os.path.dirname(self.output_input.text())
        if not default_dir:
            default_dir = os.path.expanduser("~")
        
        report_path, _ = QFileDialog.getSaveFileName(
            self, "保存处理报告", os.path.join(default_dir, "处理报告.xlsx"), 
            "Excel Files (*.xlsx)"
        )
        
        if report_path:
            if not report_path.endswith(".xlsx"):
                report_path += ".xlsx"
                
            try:
                self.report_df.to_excel(report_path, index=False)
                QMessageBox.information(self, "导出成功", f"处理报告已保存至:\n{report_path}")
            except Exception as e:
                QMessageBox.critical(self, "导出失败", f"保存报告时出错:\n{str(e)}")


if __name__ == "__main__":
    app = QApplication(sys.argv)
    window = ExcelExtractorUI()
    window.show()
    sys.exit(app.exec_())
```

## 功能说明

这个可视化工具包含以下主要功能：

### 1. 界面布局
- **源文件夹选择**：浏览按钮选择Excel文件所在文件夹
- **输出文件设置**：指定汇总结果保存路径
- **表头设置**：
  - 表头所在行号（默认为1）
  - 自动清除空行选项（默认启用）
- **列配置表格**：
  - 目标列名：汇总文件中显示的列名
  - 别名列表：源文件中可能的列名（用逗号分隔）
  - 支持添加/删除配置行
- **进度显示**：
  - 进度条显示处理进度
  - 状态标签显示当前处理的文件名
- **操作按钮**：
  - 开始处理：启动提取过程
  - 导出报告：生成处理报告（处理完成后可用）
  - 退出：关闭程序

### 2. 处理流程
1. 用户设置源文件夹、输出文件路径
2. 配置需要提取的列及其别名
3. 设置表头行号（Excel中的行号）
4. 点击"开始处理"按钮启动后台处理
5. 处理过程中显示进度和当前处理的文件名
6. 处理完成后显示结果统计信息
7. 点击"导出报告"可生成详细处理报告

### 3. 后台处理
- 使用单独的线程处理Excel文件，避免界面卡顿
- 智能匹配列名（忽略大小写、空格等）
- 自动处理不同表格结构（列顺序不同、缺少列）
- 自动清除空行（可选）
- 生成汇总文件和处理报告

## 使用指南

1. **启动程序**：
   ```bash
   python excel_extractor_gui.py
   ```

2. **配置设置**：
   - 选择源文件夹（包含Excel文件）
   - 设置输出文件路径
   - 配置需要提取的列及其别名
   - 设置表头行号（默认为1）

3. **开始处理**：
   - 点击"开始处理"按钮
   - 观察进度条和状态信息

4. **查看结果**：
   - 处理完成后，汇总文件将保存到指定位置
   - 点击"导出报告"可生成详细处理报告

## 优势特点

1. **用户友好**：
   - 直观的图形界面，无需编辑配置文件
   - 实时进度反馈
   - 错误信息清晰展示

2. **健壮性**：
   - 后台线程处理，避免界面冻结
   - 完善的错误处理机制
   - 详细的处理报告

3. **灵活性**：
   - 动态添加/删除列配置
   - 支持多种别名匹配
   - 可配置表头位置

4. **高效处理**：
   - 智能列名匹配算法
   - 批量处理多个文件
   - 自动清除空行

这个可视化工具让Excel批量信息提取变得简单直观，特别适合需要处理多个结构相似但又不完全一致的Excel文件的场景。