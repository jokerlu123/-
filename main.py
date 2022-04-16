# -*- coding:utf-8 -*-

"""
作者：阿路阿路
日期：2022年04月15日
"""
import numpy as np
import xlrd
import random
n_table = 13      # 参与投票的人数
n_抽奖 = 13        # 参与抽奖的人数

# 汇总每个人的得分       n_table:参与投票的人数
def 计算分数(n_table):
    score = np.zeros(n_table)  # 初始化得分数组
    excel = xlrd.open_workbook("test.xls")  # 打开excel文件
    for j in range(n_table):  # 在13个表里提取数据

        sheet = excel.sheet_by_index(j)  # 根据下标获取工作薄，这里获取第一个sheet
        all_j = 总和(sheet,n_抽奖)
        for i in range(n_抽奖):
            date = sheet.row_values(i - 1)  # 获取i行的信息
            score[i] = score[i] + (date[1] + date[2]) / all_j

    print("得分：", score)

    return score

# 防止有人总体打分过高或过低影响结果
def 总和(sheet,n_抽奖):
    lie2 = sheet.col_values(1, 0, n_抽奖)  # 表格中第二列的信息
    lie3 = sheet.col_values(2, 0, n_抽奖)  # 表格中第三列的信息
    l2 = np.array(lie2)                # 讲list转换成数组
    l3 = np.array(lie3)
    all2 = l2.sum()                    # 求第二列的和
    all3 = l3.sum()                    # 求第三列的和
    all = all2 + all3
    return all

# 一个很白痴的if else判断，输出
"""
13个人给出分数，总分就是13    print(score.sum())  输出：13.000000000000

接下来随机输出一个0到13的数，判断落到什么地方，即谁中奖。
"""
def 抽奖(score):
    s = random.uniform(0, 13)
    if s < score[0]:
        print("恭喜1号获奖！")
    elif score[0] < s < sum(score[:2]):
        print("恭喜2号获奖！")
    elif sum(score[:2]) < s <= sum(score[:3]):
        print("恭喜3号获奖！")
    elif sum(score[:3]) < s <= sum(score[:4]):
        print("恭喜4号获奖！")
    elif sum(score[:4]) < s <= sum(score[:5]):
        print("恭喜5号获奖！")
    elif sum(score[:5]) < s <= sum(score[:6]):
        print("恭喜6号获奖！")
    elif sum(score[:6]) < s <= sum(score[:7]):
        print("恭喜7号获奖！")
    elif sum(score[:7]) < s <= sum(score[:8]):
        print("恭喜8号获奖！")
    elif sum(score[:8]) < s <= sum(score[:9]):
        print("恭喜9号获奖！")
    elif sum(score[:9]) < s <= sum(score[:10]):
        print("恭喜10号获奖！")
    elif sum(score[:10]) < s <= sum(score[:11]):
        print("恭喜11号获奖！")
    elif sum(score[:11]) < s <= sum(score[:12]):
        print("恭喜12号获奖！")
    else:
        print("恭喜13号获奖！")

score= 计算分数(n_table)
抽奖(score)




