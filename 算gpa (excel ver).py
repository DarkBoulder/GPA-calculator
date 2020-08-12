import os
import xlrd


class Course:
    def __init__(self, name, type, credit, score):
        self.name = name
        self.type = type
        self.credit = credit
        self.score = score

    def display(self):
        print("{} {} {} {}".format(self.name, self.type, self.credit, self.score))


def get_grade(score):
    if score >= 90:
        return 4.0, 4.0, 4.0, 4.0
    elif score >= 85:
        return 3.0, 4.0, 4.0, 3.7
    elif score >= 82:
        return 3.0, 3.0, 3.0, 3.3
    elif score >= 80:
        return 3.0, 3.0, 3.0, 3.0
    elif score >= 78:
        return 2.0, 3.0, 3.0, 3.0
    elif score >= 75:
        return 2.0, 3.0, 3.0, 2.7
    elif score >= 72:
        return 2.0, 3.0, 2.0, 2.3
    elif score >= 70:
        return 2.0, 3.0, 2.0, 2.0
    elif score >= 68:
        return 1.0, 2.0, 2.0, 2.0
    elif score >= 64:
        return 1.0, 2.0, 2.0, 1.5
    elif score >= 60:
        return 1.0, 2.0, 2.0, 1.0
    else:
        return 0.0, 0.0, 1.0, 0.0


def get_info(li):
    mul_sum = cre_sum = gpa_sum_std = gpa_sum_std1 = gpa_sum_wes = gpa_sum_pku = 0
    for ele in li:
        cre_sum += ele.credit
        mul_sum += ele.credit * ele.score
        std, std1, wes, pku = get_grade(ele.score)
        gpa_sum_std += std * ele.credit
        gpa_sum_std1 += std1 * ele.credit
        gpa_sum_wes += wes * ele.credit
        gpa_sum_pku += pku * ele.credit
    return mul_sum, cre_sum, gpa_sum_std, gpa_sum_std1, gpa_sum_wes, gpa_sum_pku


def print_info(ctype, mul, cre, std, std1, wes, pku):
    print('{}:\navg {}\n标准算法 {}\n标准改进算法1 {}\nWES算法 {}\n北大算法 {}\n'.format(ctype,
                                                                                    mul/cre,
                                                                                    std/cre,
                                                                                    std1/cre,
                                                                                    wes/cre,
                                                                                    pku/cre))


if __name__ == '__main__':
    data_path = '成绩.xls'
    bixiu = []
    xuanxiu = []
    gongxuan = []
    wb = xlrd.open_workbook('成绩.xls')
    sheet1 = wb.sheet_by_index(0)
    row_num = sheet1.nrows
    col_num = sheet1.ncols
    for row in range(row_num):
        li = sheet1.row_values(row)
        if li[1] == '必修':
            bixiu.append(Course(li[0], li[1], float(li[2]), float(li[3])))
        elif li[1] == '选修':
            xuanxiu.append(Course(li[0], li[1], float(li[2]), float(li[3])))
        elif li[1] == '校公选':
            gongxuan.append(Course(li[0], li[1], float(li[2]), float(li[3])))
    bixiu_mul, bixiu_cre, bixiu_std, bixiu_std1, bixiu_wes, bixiu_pku = get_info(bixiu)
    xuanxiu_mul, xuanxiu_cre, xuanxiu_std, xuanxiu_std1, xuanxiu_wes, xuanxiu_pku = get_info(xuanxiu)
    gongxuan_mul, gongxuan_cre, gongxuan_std, gongxuan_std1, gongxuan_wes, gongxuan_pku = get_info(gongxuan)
    print_info("必修",
               bixiu_mul,
               bixiu_cre,
               bixiu_std,
               bixiu_std1,
               bixiu_wes,
               bixiu_pku)
    print_info("必修+选修",
               bixiu_mul+xuanxiu_mul,
               bixiu_cre+xuanxiu_cre,
               bixiu_std+xuanxiu_std,
               bixiu_std1+xuanxiu_std1,
               bixiu_wes+xuanxiu_wes,
               bixiu_pku+xuanxiu_pku)
    print_info("必修+选修+校公选",
               bixiu_mul+xuanxiu_mul+gongxuan_mul,
               bixiu_cre+xuanxiu_cre+gongxuan_cre,
               bixiu_std+xuanxiu_std+gongxuan_std,
               bixiu_std1+xuanxiu_std1+gongxuan_std1,
               bixiu_wes+xuanxiu_wes+gongxuan_wes,
               bixiu_pku+xuanxiu_pku+gongxuan_pku)
    os.system('pause')
