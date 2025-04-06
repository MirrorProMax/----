import numpy as np
import os
from datetime import datetime, timedelta
from openpyxl import Workbook, load_workbook
import SubFunc


def clear_console():
    """清理控制台"""
    os.system("clear")
    print(f"{datetime.now()}\n")


def load_excel_file(filename):
    """加载 Excel 文件"""
    try:
        workbook = load_workbook(filename=filename)
        return workbook.active if workbook else None
    except Exception as e:
        print(f"Error: 无法加载文件 {filename}. 错误信息: {e}")
        return None


def random_split_total(total_amount):
    # 参数校验 🔧
    if total_amount is None:
        return None
    if not isinstance(total_amount, (float, int)):
        raise TypeError("总金额必须是数值类型")
    if total_amount <= 0:
        raise ValueError("总金额必须大于0")

    # 随机生成3-6个分组
    group_count = np.random.randint(3, 7)

    # 生成金额分配权重（确保总和精确）
    fractions = np.random.dirichlet(np.ones(group_count), size=1)
    split_amounts = (fractions * total_amount).flatten()

    result = []
    remaining = total_amount  # 剩余待分配金额

    for i in range(group_count):
        # 动态调整最后一组逻辑 🔧
        if i == group_count - 1:
            current_amount = remaining
        else:
            current_amount = split_amounts[i]

        # 生成单价（严格限制范围）
        unit_price = round(np.clip(np.random.uniform(26, 28), 26, 28), 2)

        # 计算数量（保留两位小数）
        quantity = round(current_amount / unit_price, 2)

        # 记录结果
        result.append(
            {
                "分组": i + 1,
                "单价": unit_price,
                "数量": quantity,
                "小计": current_amount,
                "剩余金额": remaining - current_amount,  # 新增监控字段
            }
        )

        # 更新剩余金额（最后一组自动归零）
        remaining -= current_amount

    # 强制校验总和 🔧
    assert np.isclose(
        sum([x["小计"] for x in result]), total_amount
    ), "金额总和校验失败"

    return result


def process_rows(inSheet, output_data):
    rows = list(inSheet.iter_rows(min_row=1, values_only=True))
    columns = list(inSheet.iter_cols(min_row=1, values_only=True))
    if not columns or not rows:
        print("Error: 表格数据为空或格式不正确。")
        return
    date_row = rows[0]
    headers = columns[0]

    for row_index, row in enumerate(rows[1:], start=1):
        for col_index, cell_value in enumerate(row[1:], start=1):
            output_data.append(random_split_total(cell_value))


def save_to_excel(output_data, filename_prefix="Output"):
    """保存数据到 Excel 文件"""
    if not output_data:
        print("Warning: 没有数据需要保存。")
        return

    try:
        workbook = Workbook()
        sheet = workbook.active
        if sheet:
            sheet.append(["抬头", "日期", "序号", "单价", "数量", "金额"])
            for row in output_data:
                sheet.append(row)
            filename = (
                f"{filename_prefix}_{datetime.now().strftime('%Y-%m-%d %H-%M-%S')}.xlsx"
            )
            workbook.save(filename)
            print(f"数据已保存到文件: {filename}")
    except Exception as e:
        print(f"Error: 无法保存文件。错误信息: {e}")


def main():
    clear_console()
    sheet = load_excel_file("育肥牛.xlsx")
    if sheet is None:
        print("Error: 无法加载工作表。")
        return

    output_data = []
    process_rows(sheet, output_data)
    save_to_excel(output_data)


if __name__ == "__main__":
    main()
