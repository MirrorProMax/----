import random
import SubFunc
import os
from datetime import datetime
from openpyxl import Workbook, load_workbook

wb: Workbook = load_workbook(filename="育肥牛.xlsx")
sheet = wb.active if wb else None
if sheet is not None:
    按行拆分池 = sheet.iter_rows(min_row=2, values_only=True)

    # exit()

    取整位数: int = 2

    os.system("clear")
    print(f"{datetime.now()}")
    print()
    for 标题, 剩余拆分金额 in 按行拆分池:

        目标波动率 = 0.5
        单价下限 = 26
        单价上限 = 28
        目标条数下限 = 5
        目标条数上限 = 15
        目标条数上下限差 = 目标条数上限 - 目标条数下限

        最终目标条数 = int((random.random() * 目标条数上下限差) + 目标条数下限)
        # print(最终目标条数)

        每一份金额基准 = int(剩余拆分金额 / 最终目标条数)
        # print(每一份金额基准)

        上次剩余拆分金额 = 0
        index = 1
        while 剩余拆分金额 > 0:
            本次随机比率 = SubFunc.randomFloat(下限=-目标波动率, 上限=目标波动率)
            # print(本次随机比率)

            if 剩余拆分金额 <= 每一份金额基准:
                本次拆出金额 = round(剩余拆分金额, 取整位数)
            else:
                本次拆出金额 = round(每一份金额基准 * (1 + 本次随机比率), 取整位数)

            本次单价 = round(SubFunc.randomFloat(单价下限, 单价上限), 2)
            本次数量 = round(本次拆出金额 / 本次单价, 取整位数)

            上次剩余拆分金额 = 剩余拆分金额
            剩余拆分金额 = round(剩余拆分金额 - 本次拆出金额, 取整位数)

            print()
            print(f"{标题}:")
            print(f"序号:{index}")
            print(f"本次金额:{本次拆出金额}")
            print(f"本次单价:{本次单价}")
            print(f"本次数量:{本次数量}")
            # print(f"剩余拆分金额:{剩余拆分金额}")
            print()

            index += 1

else:
    print("Error: Failed to load the workbook or sheet is None.")
