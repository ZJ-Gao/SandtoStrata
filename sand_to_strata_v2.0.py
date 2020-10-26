'''
    作者：J.
    日期：2019.4.14
    功能：计算砂地比
    版本：2.0
    2.0 新增功能：分层序计算砂地比，模块化计算程序
'''
import xlrd

def list_section(col_top_depth, top, bottom):
    list_sec = []
    for i, element in enumerate(col_top_depth):
        if element != '顶深' and element >= top and element <= bottom:
            list_sec.append(i)
    return list_sec

def sand_to_strata(col_petrology, top, bottom, col_top_depth, col_bottom_depth):


    list_sand = []

    for i, element in enumerate(col_petrology):
        if '砂岩' in element or '砾岩' in element and '泥质粉砂岩' not in element:
            list_sand.append(i)

    list_sec = list_section(col_top_depth, top, bottom)

    list_sand_in_sec = list(set(list_sand).intersection(list_sec))
    # print(list_sand_in_sec)

    list_sand_to_strata = []
    for i in list_sand_in_sec:
        thickness = col_bottom_depth[i] - col_top_depth[i]
        list_sand_to_strata.append(round(thickness, 2))

    strata_thickness = bottom - top
    sand_to_strata_value = round(sum(list_sand_to_strata) / strata_thickness, 4)
    # print(sum(list_sand_to_strata), strata_thickness)
    return sand_to_strata_value

def main():

    data = xlrd.open_workbook('高11.xlsx')
    table = data.sheets()[0]


    col_top_depth = table.col_values(0)
    col_bottom_depth = table.col_values(1)
    col_petrology = table.col_values(2)

    bottom = col_bottom_depth[len(col_bottom_depth) -1]


    # print(top, bottom)
    #
    # print(sand_to_strata(col_petrology, top, bottom, col_top_depth, col_bottom_depth))



    # 读取层序划分列
    col_strata = table.col_values(6)
    while '' in col_strata:
        col_strata.remove('')

    # print(col_strata)
    Es1s_1 = sand_to_strata(col_petrology, col_strata[1], col_strata[2], col_top_depth, col_bottom_depth)
    Es1s_2 = sand_to_strata(col_petrology, col_strata[2], col_strata[3], col_top_depth, col_bottom_depth)
    Es1s_3 = sand_to_strata(col_petrology, col_strata[3], col_strata[4], col_top_depth, col_bottom_depth)
    Es1xt = sand_to_strata(col_petrology, col_strata[4], col_strata[5], col_top_depth, col_bottom_depth)
    Es1xw = sand_to_strata(col_petrology, col_strata[5], col_strata[6], col_top_depth, col_bottom_depth)
    Es2s = sand_to_strata(col_petrology, col_strata[6], col_strata[7], col_top_depth, col_bottom_depth)
    Es2x = sand_to_strata(col_petrology, col_strata[7], col_strata[8], col_top_depth, col_bottom_depth)
    Es3s = sand_to_strata(col_petrology, col_strata[8], bottom, col_top_depth, col_bottom_depth)
    well_name = table.cell(0, 6).value
    list_final = [well_name, Es1s_1, Es1s_2, Es1s_3, Es1xt, Es1xw, Es2s, Es2x, Es3s]
    print(list_final)



if __name__ == '__main__':
    main()