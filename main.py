import os
import xlrd
import time


def get_file_name():
    filePath = './data'
    return os.listdir(filePath)


def get_data_by_name(item, index, ):
    file_name = get_file_name()
    target_file = file_name[-1 * index]
    table = xlrd.open_workbook("./data/" + target_file).sheets()[0]
    code_values = table.col_values(0, 1)
    if item == 'code':
        return code_values
    name_index = table.row_values(0).index(item)
    item_values = table.col_values(name_index, 1)
    # target_values = []
    # for i in range(0, len(code_values)):
    #     target_values.append([code_values[i], item_values[i]])
    return item_values


def handle_arr(arr1, arr2, is_plus='-'):
    target_arr = []
    for i in range(0, len(arr1)):
        if is_plus == "+":
            target_arr.append(arr1[i] + arr2[i])
        elif is_plus == "/":
            if arr2[i] == 0:
                target_arr.append(float("inf"))
            else:
                target_arr.append(arr1[i] / arr2[i])
        else:
            target_arr.append(arr1[i] - arr2[i])
    return target_arr


def compare_arr(arr1, arr2):
    code_arr = get_data_by_name('code', 1)
    result_code = []
    for i in range(0, len(code_arr)):
        if arr1[i] > arr2[i]:
            result_code.append(code_arr[i])
    return result_code


# Press the green button in the gutter to run the script.
if __name__ == '__main__':
    code = get_data_by_name('code', 1)
    e3 = get_data_by_name('e', 3)
    e2 = get_data_by_name('e', 2)
    e1 = get_data_by_name('e', 1)
    b1 = get_data_by_name('b', 1)
    b2 = get_data_by_name('b', 2)
    b3 = get_data_by_name('b', 3)

    g1 = handle_arr(e3, e2)
    g2 = handle_arr(e2, e1)
    g3 = handle_arr(
        handle_arr(handle_arr(b1, b2, '+'), b3, '+'), [3] * len(code), '/'
    )
    g4 = handle_arr(e3, b3)
    g5 = handle_arr(e2, b2)
    g6 = handle_arr(
        handle_arr(g1, g2, '/'),
        g2
    )

    g1g2 = compare_arr(g1, g2)
    g3e1 = compare_arr(g3, e1)
    g5g4 = compare_arr(g5, g4)
    g20 = compare_arr(g2, [0] * len(code))
    e03 = compare_arr([0] * len(code), e3)
    g016 = compare_arr([0.1] * len(code), g6)
    result_arr = g1g2 + g3e1 + g5g4 + g20 + e03 + g016
    dict = {}
    for key in result_arr:
        dict[key] = dict.get(key, 0) + 1
    # print(dict)
    result = []
    for key, value in dict.items():
        if value == 6:
            result.append(key)
    print(result)
    with open(time.strftime("%Y_%m_%d_%H_%M_%S", time.localtime()) + ".txt", 'w') as f:
        f.writelines(",".join(map(str, result)))