import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment
from openpyxl.chart import PieChart, Reference
from openpyxl.chart.label import DataLabelList
import os
import sys
import time
from collections import Counter


def print_step(step_number, message):
    """打印步骤提示"""
    time.sleep(0.25)
    print(f"\n=== {message} ===\n")

def wait_for_enter():
    """等待用户按回车继续"""
    input("按回车键继续...")

def find_header_row(sheet, header_keyword="片号"):
    for row_idx, row in enumerate(sheet.iter_rows(values_only=True), 1):
        # 确保行至少有3列，然后检查第三列（索引2）
        if row[0] is not None:
            if header_keyword in str(row[0]):
                print(f"检测到表头在第 {row_idx} 行")
                return row_idx
    print("未找到表头，默认返回第 1 行")
    return 1

def is_valid_code(code):
    """检查是否为有效的编码（不以#开头）"""
    if code is None:
        return False

    code_str = str(code)
    # 检查是否以#开头（Excel错误值通常以#开头）
    if code_str.startswith('#'):
        return False

    return True

def load_graph_data(file_path):
    """加载Excel数据，只读取所需的列"""
    wb = openpyxl.load_workbook(file_path)
    sheet = wb.active

    # 查找表头行
    header_row = find_header_row(sheet)

    data = []
    for row in sheet.iter_rows(min_row=header_row + 1, values_only=True):
        # 检查产品编码是否为有效值（不以#开头）
        slice_code = str(row[0]) if row[0] is not None else None
        if not is_valid_code(slice_code):
            continue

        row_data = {
            "片号": slice_code,
            "这个缺陷": row[1],
            "哪个缺陷": row[2],
            "就是这个缺陷": row[3],

        }
        data.append(row_data)

    return data

def analyze_defect_data(graph_data):
    """
    分析缺陷数据，统计各工序缺陷类型的占比
    参数:
    graph_data: 图形数据列表
    返回:
    字典，键为工序名，值为缺陷统计Counter对象
    """
    # 定义要分析的工序列
    process_columns = ["这个缺陷", "哪个缺陷", "就是这个缺陷"]

    defect_stats = {}

    for column in process_columns:
        # 收集该列的所有非空值
        defects = []
        for item in graph_data:
            defect = item.get(column)
            if defect is not None and str(defect).strip() != "":
                defects.append(str(defect).strip())

        # 统计缺陷类型
        defect_counter = Counter(defects)
        defect_stats[column] = defect_counter

    return defect_stats


def create_pie_charts(workbook, defect_stats):
    """
    创建饼图并添加到工作簿
    参数:
    workbook: openpyxl工作簿对象
    defect_stats: 缺陷统计字典
    """
    # 为每个工序创建一个工作表并添加饼图
    for process_name, counter in defect_stats.items():
        # 创建工作表
        sheet_name = f"{process_name}缺陷分析"
        if sheet_name in workbook.sheetnames:
            # 如果已存在，则删除原有工作表
            del workbook[sheet_name]
        ws = workbook.create_sheet(title=sheet_name)

        # 添加表头
        ws['A1'] = "缺陷类型"
        ws['B1'] = "数量"
        ws['C1'] = "占比"

        # 设置表头样式
        for cell in ['A1', 'B1', 'C1']:
            ws[cell].font = Font(bold=True)
            ws[cell].fill = PatternFill(start_color="DDDDDD", end_color="DDDDDD", fill_type="solid")

        # 添加数据
        total_count = sum(counter.values())
        row = 2
        for defect_type, count in counter.most_common():
            ws[f'A{row}'] = defect_type
            ws[f'B{row}'] = count
            ws[f'C{row}'] = count / total_count
            row += 1

        # 设置百分比格式
        for r in range(2, row):
            ws[f'C{r}'].number_format = '0.00%'

        # 创建饼图
        chart = PieChart()
        chart.title = f"{process_name}缺陷分布"

        # 设置数据范围
        labels = Reference(ws, min_col=1, min_row=2, max_row=row - 1)
        data = Reference(ws, min_col=2, min_row=1, max_row=row - 1)

        chart.add_data(data, titles_from_data=True)
        chart.set_categories(labels)

        # 设置图表样式
        chart.dataLabels = DataLabelList()
        chart.dataLabels.showPercent = True
        chart.dataLabels.showLegendKey = False
        chart.dataLabels.showVal = False
        chart.dataLabels.showCatName = True

        # 将图表添加到工作表
        ws.add_chart(chart, "E2")

        # 调整列宽
        ws.column_dimensions['A'].width = 20
        ws.column_dimensions['B'].width = 10
        ws.column_dimensions['C'].width = 10


def save_results_to_excel(defect_stats, output_file="工序缺陷统计.xlsx"):
    """将结果保存到Excel文件"""
    # 创建新的工作簿
    workbook = openpyxl.Workbook()


    # 如果有缺陷数据，添加饼图
    if defect_stats:
        create_pie_charts(workbook, defect_stats)

    # 保存文件
    workbook.save(output_file)
    print(f"结果已保存到 {output_file}")


# 主程序
if __name__ == "__main__":
    print_step(1, "长晶工艺缺陷分布")
    print("请确保Excel文件与此程序在同一文件夹中")
    wait_for_enter()

    # 查找Excel文件
    print_step(2, "查找Excel文件")
    excel_files = []
    for file in os.listdir('.'):
        if file.endswith('.xlsx') or file.endswith('.xls'):
            excel_files.append(file)

    if not excel_files:
        print("未找到Excel文件(.xlsx或.xls)")
        print("请将Excel文件放入与此程序相同的文件夹中")
        input("按回车键退出...")
        sys.exit(1)

    # 如果找到多个Excel文件，让用户选择
    if len(excel_files) > 1:
        print_step(3, "发现多个Excel文件，请选择要分析的文件:")
        for i, file in enumerate(excel_files, 1):
            print(f"{i}. {file}")

        while True:
            try:
                choice = int(input("请输入文件编号: "))
                if 1 <= choice <= len(excel_files):
                    file_path = excel_files[choice - 1]
                    break
                else:
                    print("编号无效，请重新输入")
            except ValueError:
                print("请输入有效数字")
    else:
        file_path = excel_files[0]
        print(f"找到文件: {file_path}")
        wait_for_enter()


    print_step(4, "开始分析数据")
    print("正在读取和分析Excel文件...")

    try:

        # 分析缺陷数据
        print_step(6, "分析缺陷数据")
        graph_data = load_graph_data(file_path)
        defect_stats = analyze_defect_data(graph_data)

        print_step(7, "保存结果")
        # 保存到Excel
        output_file = "工序缺陷统计.xlsx"
        save_results_to_excel(defect_stats, output_file)

        print_step(8, "完成")
        print("所有操作已完成!")
        print("您可以在同一文件夹中找到 '工序缺陷统计.xlsx' 文件")
        print("文件中包含了各工序缺陷的饼图分析")

    except Exception as e:
        print(f"处理过程中发生错误: {str(e)}")
        import traceback

        traceback.print_exc()

    input("按回车键退出程序...")