'''
    Author: ZIJIE GAO
    Date: 4/4/2019
    Function: Calculate the ratio of sandstone to strata
    Version：.0
    2.0 New function: Calculate the ratio of sandstone to stratum for each sequence
    3.0 New function:
    1) Add codes which can process many documents at a time
    2) Write the results into .txt document
    4.0 New function: Modify codes to get used to new sequence division
    5.0 Debug
'''
import xlrd
import os

def get_filename(path,filetype):
    '''
        Read file name
    '''
    name = []
    for root,dirs,files in os.walk(path):
        for i in files:
            if filetype in i:
                name.append(i.replace(filetype,''))
    return name

def list_section(col_top_depth, top, bottom):
    list_sec = []
    for i, element in enumerate(col_top_depth):
        if element != 'top depth' and element >= top and element <= bottom:
            list_sec.append(i)
    return list_sec

def sand_to_strata(col_petrology, top, bottom, col_top_depth, col_bottom_depth):


    list_sand = []

    for i, element in enumerate(col_petrology):
        if 'sandstone' in element and 'muddy siltstone' not in element or 'conglomerate' in element:
            list_sand.append(i)

    list_sec = list_section(col_top_depth, top, bottom)

    list_sand_in_sec = list(set(list_sand).intersection(list_sec))
    # print(list_sand_in_sec)


    list_sand_to_strata = []
    for i in list_sand_in_sec:
        thickness = col_bottom_depth[i] - col_top_depth[i]
        list_sand_to_strata.append(round(thickness, 2))

    sand_thickness = sum(list_sand_to_strata)
    print(list_sand_to_strata, sand_thickness)
    # print(sum(list_sand_to_strata), strata_thickness)
    return sand_thickness


def write_txt(list):
    '''
        write results into .txt document
    '''
    fileObject = open('sand_to_strata.txt', 'a', encoding='utf-8')
    fileObject.write('\n')
    for ip in list:
        fileObject.writelines(str(ip))
        fileObject.write(' ')
    fileObject.write('\n')

    fileObject.close()


def main():
    path = 'C:\\changename_three'
    filetype = '.xlsx'

    name = get_filename(path, filetype)
    for word in name:

        data = xlrd.open_workbook(word + '.xlsx')
        table = data.sheets()[0]

        col_top_depth = table.col_values(0)
        col_bottom_depth = table.col_values(1)
        col_petrology = table.col_values(2)

        bottom = col_top_depth[len(col_top_depth) -1]
        # print(top, bottom)
        #
        # print(sand_to_strata(col_petrology, top, bottom, col_top_depth, col_bottom_depth))
        # Read the column recording depth of each sequence
        col_strata = table.col_values(6)
        while '' in col_strata:
            col_strata.remove('')

        # print(col_strata)
        Es1s_1 = sand_to_strata(col_petrology, col_strata[1], col_strata[2], col_top_depth, col_bottom_depth)
        Es1s_2 = sand_to_strata(col_petrology, col_strata[2], col_strata[3], col_top_depth, col_bottom_depth)
        Es1s_3 = sand_to_strata(col_petrology, col_strata[3], col_strata[4], col_top_depth, col_bottom_depth)
        Es1_4 = sand_to_strata(col_petrology, col_strata[4], col_strata[5], col_top_depth, col_bottom_depth)
        Es1_5 = sand_to_strata(col_petrology, col_strata[5], col_strata[6], col_top_depth, col_bottom_depth)
        Es1xt = sand_to_strata(col_petrology, col_strata[6], col_strata[7], col_top_depth, col_bottom_depth)
        Es1xw = sand_to_strata(col_petrology, col_strata[7], col_strata[8], col_top_depth, col_bottom_depth)
        Es2s = sand_to_strata(col_petrology, col_strata[8], col_strata[9], col_top_depth, col_bottom_depth)
        Es2x = sand_to_strata(col_petrology, col_strata[9], col_strata[10], col_top_depth, col_bottom_depth)

        well_name = table.cell(0, 6).value
        list_final = [well_name, Es1s_1, Es1s_2, Es1s_3, Es1_4, Es1_5, Es1xt, Es1xw, Es2s, Es2x]

        print(list_final)
        write_txt(list_final)

if __name__ == '__main__':
    main()