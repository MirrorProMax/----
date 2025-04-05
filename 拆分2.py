import numpy as np

import numpy as np


def random_split_total(total_amount):
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


# 示例使用
total = 1000  # 替换为你的总金额
split_result = random_split_total(total)

# 打印结果
print(f"总金额: {total}元")
print("拆分结果:")
for item in split_result:
    print(
        f"分组 {item['分组']}: 单价 {item['单价']}元 × {item['数量']}件 = {item['小计']}元"
    )
print(f"总计: {sum([x['小计'] for x in split_result]):.2f}元")
