import numpy as np
import os
from datetime import datetime, timedelta
from openpyxl import Workbook, load_workbook
import SubFunc
import random


def refresh_console():
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


def random_split_total(total_amount: float):
    # 参数校验 🔧
    if total_amount is None:
        return None
    if not isinstance(total_amount, (float, int)):
        raise TypeError("总金额必须是数值类型")
    if total_amount <= 0:
        raise ValueError("总金额必须大于0")

    # 随机生成3-6个分组
    group_count: int = int(np.random.randint(3, 7))

    # 生成金额分配权重（确保总和精确）
    fractions_NDArray = np.random.dirichlet(np.ones(group_count), size=1)
    split_amounts_NDArray = (fractions_NDArray * total_amount).flatten()

    split_amounts = split_amounts_NDArray.tolist()

    result: list = []
    remaining: float = float(total_amount)  # 剩余待分配金额

    unit_price_lower_limit = 26
    unit_price_upper_limit = 28

    for i in range(group_count):
        # 动态调整最后一组逻辑
        if i == group_count - 1:
            current_amount: float = float(remaining)
        else:
            current_amount: float = split_amounts[i]

        # 生成单价（严格限制范围）
        current_unit_price = round(
            SubFunc.randomFloat(unit_price_lower_limit, unit_price_upper_limit), 2
        )

        # 计算数量（保留两位小数）
        quantity: float = float(round(current_amount / current_unit_price, 2))

        # 记录结果
        result.append(
            [
                i + 1,
                current_unit_price,
                quantity,
                current_amount,
                remaining - current_amount,  # 新增监控字段
            ]
        )

        # 更新剩余金额（最后一组自动归零）
        remaining -= float(current_amount)

    # 强制校验总和 🔧
    if abs(sum(split_amounts) - total_amount) > 0.01:
        raise ValueError("金额分配不准确")

    return result


def process_rows(inSheet, output_data: list):
    rows = list(inSheet.iter_rows(min_row=1, values_only=True))
    columns = list(inSheet.iter_cols(min_row=1, values_only=True))
    if not columns or not rows:
        print("Error: 表格数据为空或格式不正确。")
        return
    date_row = rows[0]
    headers = columns[0]

    for row_index, row in enumerate(rows[1:], start=1):
        for col_index, cell_value in enumerate(row[1:], start=1):
            if cell_value is None:
                continue
            split_result = random_split_total(cell_value)

            # 随机生成“日”的值（1到28之间，避免月份天数问题）
            original_date = date_row[col_index]
            if isinstance(original_date, datetime):
                random_day = random.randint(1, 28)
                modified_date = original_date.replace(day=random_day)
            else:
                modified_date = original_date  # 如果不是日期类型，保持原样

            # 展平结果并添加主体和日期信息
            for item in split_result:
                output_data.append(
                    [
                        row[0],  # 主体
                        modified_date,  # 日期
                        *item,  # 展平的分组数据
                    ]
                )


def save_to_excel(output_data, filename_prefix="Output"):
    """保存数据到 Excel 文件"""
    if not output_data:
        print("Warning: 没有数据需要保存。")
        return

    try:
        workbook = Workbook()
        sheet = workbook.active
        if sheet:
            sheet.append(["主体", "日期", "序号", "单价", "数量", "金额", "剩余金额"])
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
    refresh_console()
    sheet = load_excel_file("育肥牛.xlsx")
    if sheet is None:
        print("Error: 无法加载工作表。")
        return

    output_data = []
    process_rows(sheet, output_data)
    save_to_excel(output_data)


if __name__ == "__main__":
    main()
