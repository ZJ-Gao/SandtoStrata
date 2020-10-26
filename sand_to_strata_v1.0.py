'''
    作者：J.
    日期：2019.4.13
    功能：计算砂地比
    版本：1.0
'''
import xlrd
def main():

    data = xlrd.open_workbook('高11.xlsx')
    table = data.sheets()[0]
    col_top_depth = table.col_values(0)
    col_bottom_depth = table.col_values(1)
    col_petrology = table.col_values(2)
    list1 = []
    list_sand = []
    for i, element in enumerate(col_petrology):
        if '砂岩' in element and element != '泥质粉砂岩':
            list1.append(i)

    for i in list1:
        thickness = col_bottom_depth[i] - col_top_depth[i]
        list_sand.append(round(thickness, 2))

    bottom = col_bottom_depth[len(col_bottom_depth) - 1]
    top = col_top_depth[1]

    strata_thickness = bottom
    sand_to_strata = round(sum(list_sand) / strata_thickness, 4)
    print(sand_to_strata)

if __name__ == '__main__':
    main()