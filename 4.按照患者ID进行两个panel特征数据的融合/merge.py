import pandas as pd

# 读取 Excel 文件
# df = pd.read_excel('E:/研一/Community-Detection/dataset/数据处理/分社区（细胞接触）/feature_cd(con)2 - 副本.xlsx')
df = pd.read_excel("E:\研一\Community-Detection\dataset\新分割\Panel2\Panel2_社区优化特征.xlsx")

# 指定需要合并的列
merge_columns = ['患者id']  # 以病人id为合并列

# 根据指定列合并相同数值的行，并计算其他列的平均值
merged_df = df.groupby(merge_columns, as_index=False).mean()

# 保存结果到新的 Excel 文件
# merged_df.to_excel('E:/研一/Community-Detection/dataset/数据处理/分社区（细胞接触）/feature2.xlsx', index=False)
merged_df.to_excel('E:\研一\Community-Detection\dataset\新分割\Panel2\Panel2_特征.xlsx', index=False)

# 这里要分别对panel1和panel2进行相同的操作，然后直接用excel的VLOOKUP函数进行合并