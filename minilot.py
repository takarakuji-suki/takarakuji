import xlrd

zen_data = xlrd.open_workbook('T-data.xls')

mini_data = zen_data.sheet_by_name('MINI')

def get_data_all(self):
    return [mini_data.row_values(row) for row in range(mini_data.nrows) ]

mini_list = get_data_all(mini_data)


def yosou_num_max(n_kome):
    kari_num1 = []
    num1 = []
    for i in range(len(mini_list) - 1):
        for j in mini_list[i]:
            if mini_list[len(mini_list) - 1][n_kome] == j:
                for k in mini_list[i + 1]:
                    kari_num1.append(k)

    [num1.append(int(f)) for f in kari_num1]
    num1_dic = {}
    for j in range(31):
        num1_dic[j + 1] = num1.count(j + 1)

    max_count = max(num1_dic.values())
    max_num = []
    [max_num.append(k) for k, v in num1_dic.items() if v == max_count]
    return max_num

ans_max = []
for i in range(5):
    ans_max.append(yosou_num_max(i))
print(str(sorted(ans_max)) + ",", end="")


print("")


def yosou_num_min(n_kome):
    kari_num1 = []
    num1 = []
    for i in range(len(mini_list) - 1):
        for j in mini_list[i]:
            if mini_list[len(mini_list) - 1][n_kome] == j:
                for k in mini_list[i + 1]:
                    kari_num1.append(k)

    [num1.append(int(f)) for f in kari_num1]
    num1_dic = {}
    for j in range(31):
        num1_dic[j + 1] = num1.count(j + 1)

    min_count = min(num1_dic.values())
    min_num = []
    [min_num.append(k) for k, v in num1_dic.items() if v == min_count]
    return min_num


ans_min = []
for i in range(5):
    ans_min.append(yosou_num_min(i))
print(str(sorted(ans_min)) + ",", end="")