import xlrd

zen_data = xlrd.open_workbook('T-data.xls')

l6_data = zen_data.sheet_by_name('L6')

def get_data_all(self):
    return [l6_data.row_values(row) for row in range(l6_data.nrows) ]

l6_list = get_data_all(l6_data)


def yosou_num_max(n_kome):
    kari_num1 = []
    num1 = []
    for i in range(len(l6_list) - 1):
        for j in l6_list[i]:
            if l6_list[len(l6_list) - 1][n_kome] == j:
                for k in l6_list[i + 1]:
                    kari_num1.append(k)

    [num1.append(int(h)) for h in kari_num1]
    num1_dic = {}
    for j in range(43):
        num1_dic[j + 1] = num1.count(j + 1)

    max_count = max(num1_dic.values())
    max_num = []
    [max_num.append(k) for k, v in num1_dic.items() if v == max_count]
    return max_num

ans_max = []
for i in range(6):
    ans_max.append(yosou_num_max(i))
print(str(sorted(ans_max)) + ",", end="")


print("")


def yosou_num_min(n_kome):
    kari_num1 = []
    num1 = []
    for i in range(len(l6_list) - 1):
        for j in l6_list[i]:
            if l6_list[len(l6_list) - 1][n_kome] == j:
                for k in l6_list[i + 1]:
                    kari_num1.append(k)

    [num1.append(int(h)) for h in kari_num1]
    num1_dic = {}
    for j in range(43):
        num1_dic[j + 1] = num1.count(j + 1)

    min_count = min(num1_dic.values())
    min_num = []
    [min_num.append(k) for k, v in num1_dic.items() if v == min_count]
    return min_num


ans_min = []
for i in range(6):
    ans_min.append(yosou_num_min(i))
print(str(sorted(ans_min)) + ",", end="")