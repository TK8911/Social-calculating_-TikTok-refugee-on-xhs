import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
from pprint import pprint
plt.rcParams['font.sans-serif'] = ['SimHei']
plt.rcParams['axes.unicode_minus'] = False

file_path = "D:\信管专业\社会计算\社会计算小组作业\评论update.xlsx"
excel_file = pd.ExcelFile(file_path)
df = excel_file.parse('有标记的')
df =df[df['user_origin'] != 0]

pivot_table = df.pivot_table(index='user_origin', columns='sentiment', values='评论ID', aggfunc='count').fillna(0)
percentage_table = pivot_table.div(pivot_table.sum(axis=1), axis=0) * 100
percentage_table = percentage_table.loc[:, (percentage_table > 0.1).any()]
print(percentage_table.round(2))

bar_width = 0.2
index = np.arange(len(percentage_table.columns))
for i, country in enumerate(percentage_table.index):
    plt.bar(index + i * bar_width, percentage_table.loc[country], bar_width, label=country)

plt.title('不同国家用户的sentiment分布百分比情况')
plt.xlabel('sentiment')
plt.xticks(rotation=45)
plt.ylabel('百分比（%）')
plt.xticks(index + bar_width * (len(percentage_table.index) - 1) / 2, percentage_table.columns)
plt.legend()


for i, country in enumerate(percentage_table.index):
    for j, value in enumerate(percentage_table.loc[country]):
        plt.text(index[j] + i * bar_width, value, f'{value:.2f}%', ha='center', va='bottom', fontsize=8)

# plt.show()


pd.set_option('display.max_columns', None)
pd.set_option('display.max_rows', None)
pd.set_option('display.width', None)
df['arousal'] = pd.to_numeric(df['arousal'], errors='coerce')
grouped_data = df.groupby('user_origin')[['valence','arousal', 'dominance']].describe()
print('愉悦度和支配度的唤醒度统计信息：')
pprint(grouped_data)
