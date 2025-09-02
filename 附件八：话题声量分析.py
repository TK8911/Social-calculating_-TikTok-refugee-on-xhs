import pandas as pd
import matplotlib.pyplot as plt
from datetime import datetime
import matplotlib.dates as mdates
plt.rcParams['font.sans-serif'] = ['SimHei']
plt.rcParams['axes.unicode_minus'] = False

def convert_date(date_string):
    if isinstance(date_string, datetime):
        return date_string.strftime('%Y/%m/%d')
    try:
        if isinstance(date_string, str) and ' ' in date_string:
            dt = datetime.strptime(date_string, '%m/%d/%Y %H:%M')
        elif isinstance(date_string, str):
            dt = datetime.strptime(date_string, '%Y/%m/%d')
        return dt.strftime('%Y/%m/%d')
    except ValueError:
        return None

# 读取 Excel 文件
file_path = r"D:\信管专业\社会计算\社会计算小组作业\评论update.xlsx"
excel_file = pd.ExcelFile(file_path)
df = excel_file.parse('筛选后的sum')

# 对评论时间列应用日期转换函数
df['评论时间'] = df['评论时间'].apply(convert_date)

# 将评论时间转换为 datetime 类型
df['评论时间'] = pd.to_datetime(df['评论时间'])

# 按评论时间和笔记 topic 分组，统计每个组合下评论 ID 的数量（即声量）
topic_volume_over_time = df.groupby(['评论时间', '笔记topic'])['评论ID'].count().unstack(fill_value=0)

# 创建画布
fig, ax = plt.subplots(figsize=(15, 8))

# 绘制不同笔记 topic 声量随时间变化的折线图
for topic in topic_volume_over_time.columns:
    ax.plot(topic_volume_over_time.index, topic_volume_over_time[topic], label=topic)

# 设置图表标题和坐标轴标签
ax.set_title('不同笔记 topic 声量随时间变化趋势')
ax.set_xlabel('评论时间')
ax.set_ylabel('声量（频次）')

# 设置 x 轴刻度为每月一次
ax.xaxis.set_major_locator(mdates.MonthLocator())
ax.xaxis.set_major_formatter(mdates.DateFormatter('%Y/%m'))

# 旋转 x 轴刻度标签
plt.xticks(rotation=45)

# 添加图例
ax.legend()

# 自动调整布局
plt.tight_layout()
start_date = datetime(2025, 1, 1)
end_date = datetime(2025, 7, 31)
ax.set_xlim(start_date, end_date)
# 显示图表
plt.show()

#---
def convert_date(date_string):
    if isinstance(date_string, datetime):
        return date_string.strftime('%Y/%m/%d')
    try:
        if isinstance(date_string, str) and ' ' in date_string:
            dt = datetime.strptime(date_string, '%m/%d/%Y %H:%M')
        elif isinstance(date_string, str):
            dt = datetime.strptime(date_string, '%Y/%m/%d')
        return dt.strftime('%Y/%m/%d')
    except ValueError:
        return None


# 对评论时间列应用日期转换函数
df['评论时间'] = df['评论时间'].apply(convert_date)

# 将评论时间转换为 datetime 类型
df['评论时间'] = pd.to_datetime(df['评论时间'])

# 按评论时间和笔记 topic 分组，统计每个组合下评论 ID 的数量（即声量）
topic_volume_over_time = df.groupby(['评论时间', '笔记topic'])['评论ID'].count().unstack(fill_value=0)

# 按评论日期分组，对每个日期下所有话题的声量求和
daily_total_volume = topic_volume_over_time.sum(axis=1)

# 创建画布
plt.figure(figsize=(15, 8))

# 绘制每日话题总声量折线图
plt.plot(daily_total_volume.index, daily_total_volume)
for x, y in zip(daily_total_volume.index, daily_total_volume):
    plt.annotate(f'{y}', (x, y), textcoords='offset points', xytext=(0,5), ha='center', fontsize=8)
# 设置图表标题和坐标轴标签
plt.title('每日话题总声量折线图')
plt.xlabel('评论日期')
plt.xticks(rotation=45)
plt.ylabel('总声量（频次）')

# 显示图表
plt.show()