"""
处理python使用round()四舍五入到个位数无法精确的问题，如round(0.5)=0
"""


def round_up(number, power=0):
    """
    实现精确四舍五入，包含正、负小数多种场景
    :param number: 需要四舍五入的小数
    :param power: 四舍五入位数，支持0-∞
    :return: 返回四舍五入后的结果
    """
    digit = 10 ** power
    num2 = float(int(number * digit))
    # 处理正数，power不为0的情况
    if number >= 0 and power != 0:
        tag = number * digit - num2 + 1 / (digit * 10)
        if tag >= 0.5:
            return (num2 + 1) / digit
        else:
            return num2 / digit
    # 处理正数，power为0取整的情况
    elif number >= 0 and power == 0:
        tag = number * digit - int(number)
        if tag >= 0.5:
            return (num2 + 1) / digit
        else:
            return num2 / digit
    # 处理负数，power为0取整的情况
    elif power == 0 and number < 0:
        tag = number * digit - int(number)
        if tag <= -0.5:
            return (num2 - 1) / digit
        else:
            return num2 / digit
    # 处理负数，power不为0的情况
    else:
        tag = number * digit - num2 - 1 / (digit * 10)
        if tag <= -0.5:
            return (num2 - 1) / digit
        else:
            return num2 / digit


if __name__ == "__main__":
    print(int(round_up(0.5)))
    print(int(round_up(1.5)))
    print(int(round_up(2.5)))
    print(int(round_up(3.5)))
