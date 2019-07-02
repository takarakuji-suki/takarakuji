import xlrd

zen_data = xlrd.open_workbook('T-data.xls')

l7_data = zen_data.sheet_by_name('L7')

def get_deta_all(self):
    return [l7_data.row_values(row) for row in range(l7_data.nrows) ]

l7_list = get_deta_all(l7_data)


def yosou_num_max(n_kome):
    kari_num1 = []
    num1 = []
    for i in range(len(l7_list) - 1):
        for j in l7_list[i]:
            if l7_list[len(l7_list) - 1][n_kome] == j:
                for k in l7_list[i + 1]:
                    kari_num1.append(k)

    [num1.append(int(f)) for f in kari_num1]
    num1_dic = {}
    for j in range(37):
        num1_dic[j + 1] = num1.count(j + 1)

    max_count = max(num1_dic.values())
    max_num = []
    [max_num.append(k) for k, v in num1_dic.items() if v == max_count]
    return max_num

ans_max = []
for i in range(7):
    ans_max.append(yosou_num_max(i))
print(str(sorted(ans_max)) + ",", end="")


print("")


def yosou_num_min(n_kome):
    kari_num1 = []
    num1 = []
    for i in range(len(l7_list) - 1):
        for j in l7_list[i]:
            if l7_list[len(l7_list) - 1][n_kome] == j:
                for k in l7_list[i + 1]:
                    kari_num1.append(k)

    [num1.append(int(f)) for f in kari_num1]
    num1_dic = {}
    for j in range(37):
        num1_dic[j + 1] = num1.count(j + 1)

    min_count = min(num1_dic.values())
    min_num = []
    [min_num.append(k) for k, v in num1_dic.items() if v == min_count]
    return min_num


ans_min = []
for i in range(7):
    ans_min.append(yosou_num_min(i))
print(str(sorted(ans_min)) + ",", end="")