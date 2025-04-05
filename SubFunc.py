# Sub Function
from typing import Any
import random


# 更自然的input方法
def inputNature(
    内容: str = "请输入",
) -> str:
    print(内容 + ": ", end="")
    return input()


# 封装的输入数字
def inputInt(inPrompt) -> int:
    return int(inputNature(inPrompt))


# 通用的批量打印
def subFunc_AutoMultiPrint(*inVs: str, inTitle=None, inLambda=lambda v: v):
    counter = 0
    theSymbolBetweenTheKeyAndTheValue = ": "
    intervals = ", "

    for v in inVs:
        if counter != 0:
            print("  |  ", end="")
        counter += 1

        content1: str = (
            f"{inTitle}{theSymbolBetweenTheKeyAndTheValue}{inLambda(v)}"
            if inTitle != None
            else ""
        )

        content2: str = f"{v}"

        print(
            f"{content1}{intervals}{content2}",
            end="",
        )
    print()


# 封装的查看变量的类型
def printType(*inVs):
    subFunc_AutoMultiPrint(*inVs, inTitle="Type", inLambda=lambda v: type(v))


# 封装的查看变量的ID
def printID(*inVs):
    subFunc_AutoMultiPrint(*inVs, inTitle="ID", inLambda=lambda v: id(v))


# 更自然的range
def rangeNature(inStart: int, inStop: int, inStep=1) -> range:
    offset = inStep
    if inStart > inStop:
        inStep = -inStep
        offset = -offset
    return range(
        inStart,
        inStop + offset,
        inStep,
    )


# 封装的列表推导式
def listComprehensionNature(
    inStartNum, inEndNum, inStep=1, inLambda=(lambda x: True)
) -> list:
    return [i for i in rangeNature(inStartNum, inEndNum, inStep) if inLambda(i)]


# print列表
def printList(inList: list[Any]):
    for i in inList:
        print(i)


# 如果脚本被直接运行，__name__ 的值就是 "__main__"
# 该部分可以用作调试
if __name__ == "__main__":
    pass
else:
    print("SubFunction Was Imported")


def randomInt(下限, 上限) -> int:
    return int(random.random() * (上限 - 下限) + 下限)


def randomFloat(下限, 上限) -> float:
    return random.random() * (上限 - 下限) + 下限
