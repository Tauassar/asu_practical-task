def check_iin(iin):
    """
    Function to check whether the input iin is correct or not
    :param iin:str: actual iin
    :return: True if iin is correct, false otherwise
    :rtype: bool
    """
    # retrieve last number to check it with control sum
    a12 = int(iin[-1])
    # iin numbers to be used in control sum calculations
    numbers = iin[:-1]
    numbers_sum = 0
    for i in range(len(numbers)):
        numbers_sum += int(numbers[i]) * (i + 1)
    control_num = numbers_sum % 11
    if control_num == 10:
        weights = [3, 4, 5, 6, 7, 8, 9, 10, 11, 1, 2]
        numbers_sum = 0
        for i in range(len(numbers)):
            numbers_sum += int(numbers[i]) * weights[i]
        control_num = numbers_sum % 11
        if 0 > control_num > 9:
            return False
    return True if control_num == a12 else False


def check_bin(bin):
    # According to 5 chapter bin checked the same way as iin
    # https://adilet.zan.kz/rus/docs/P030000565_
    check_iin(bin)
