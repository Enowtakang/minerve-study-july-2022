import pandas as pd
from scipy import stats


"""
Specify paths to data files
"""
path_1 = 'C:/0minerve/1-LMO1-H1-a25.xlsx'
path_2 = 'C:/0minerve/2-LMO1-H1-b25.xlsx'
path_3 = 'C:/0minerve/3-LMO1-H2-a25.xlsx'
path_4 = 'C:/0minerve/4-LMO1-H2-b25.xlsx'
path_5 = 'C:/0minerve/5-LMO1-H3-a25.xlsx'
path_6 = 'C:/0minerve/6-LMO1-H3-b25.xlsx'
path_7 = 'C:/0minerve/7-LMO2-H1-a25.xlsx'
path_8 = 'C:/0minerve/8-LMO2-H1-b25.xlsx'
path_9 = 'C:/0minerve/9-LMO2-H2-a25.xlsx'
path_10 = 'C:/0minerve/10-LMO2-H2-b25.xlsx'
path_11 = 'C:/0minerve/11-LMO2-H3-a25.xlsx'
path_12 = 'C:/0minerve/12-LMO2-H3-b25.xlsx'
path_13 = 'C:/0minerve/13-LKO1-H1-a32.xlsx'
path_14 = 'C:/0minerve/14-LKO1-H1-b33.xlsx'
path_15 = 'C:/0minerve/15-LKO1-H2-a32.xlsx'
path_16 = 'C:/0minerve/16-LKO1-H2-b33.xlsx'
path_17 = 'C:/0minerve/17-LKO1-H3-a32.xlsx'
path_18 = 'C:/0minerve/18-LKO1-H3-b33.xlsx'
path_19 = 'C:/0minerve/19-LGL1-H1-a25.xlsx'
path_20 = 'C:/0minerve/20-LGL1-H1-b25.xlsx'
path_21 = 'C:/0minerve/21-LGL1-H2-a25.xlsx'
path_22 = 'C:/0minerve/22-LGL1-H2-b25.xlsx'
path_23 = 'C:/0minerve/23-LGL1-H3-a25.xlsx'
path_24 = 'C:/0minerve/24-LGL1-H3-b25.xlsx'
path_25 = 'C:/0minerve/25-LGL2-H1-a25.xlsx'
path_26 = 'C:/0minerve/26-LGL2-H1-b25.xlsx'
path_27 = 'C:/0minerve/27-LGL2-H2-a25.xlsx'
path_28 = 'C:/0minerve/28-LGL2-H2-b25.xlsx'
path_29 = 'C:/0minerve/29-LGL2-H3-a25.xlsx'
path_30 = 'C:/0minerve/30-LGL2-H3-b25.xlsx'
path_31 = 'C:/0minerve/31-LYNE1-H1-a25.xlsx'
path_32 = 'C:/0minerve/32-LYNE1-H1-b25.xlsx'
path_33 = 'C:/0minerve/33-LYNE1-H2-a25.xlsx'
path_34 = 'C:/0minerve/34-LYNE1-H2-b25.xlsx'
path_35 = 'C:/0minerve/35-LYNE1-H3-a25.xlsx'
path_36 = 'C:/0minerve/36-LYNE1-H3-b25.xlsx'
path_37 = 'C:/0minerve/37-LYNE2-H1-a25.xlsx'
path_38 = 'C:/0minerve/38-LYNE2-H1-b25.xlsx'
path_39 = 'C:/0minerve/39-LYNE2-H2-a25.xlsx'
path_40 = 'C:/0minerve/40-LYNE2-H2-b25.xlsx'
path_41 = 'C:/0minerve/41-LYNE2-H3-a25.xlsx'
path_42 = 'C:/0minerve/42-LYNE2-H3-b25.xlsx'


"""
Create list of data files
"""
data_list = [path_1, path_2,
             path_3, path_4, path_5, path_6,
             path_7, path_8, path_9, path_10,
             path_11, path_12, path_13, path_14,
             path_15, path_16, path_17,
             path_18, path_19, path_20,
             path_21, path_22, path_23,
             path_24, path_25, path_26,
             path_27, path_28, path_29,
             path_30, path_31, path_32,
             path_33, path_34, path_35,
             path_36, path_37, path_38,
             path_39, path_40, path_41, path_42]


"""
Read data files and calculate:
            
            Mann-Whitney U test per file
            
"""
for path in data_list:
    names = ['Pre-Test', 'Post-Test']
    df = pd.read_excel(path,
                       header=None,
                       names=names,
                       engine='openpyxl')
    result = stats.ttest_rel(
        df['Pre-Test'], df['Post-Test'])
    # print(df)
    print(path, result[0], result[1], result[1]/2)

