from datetime import date
import pandas as pd
import numpy as np
import matplotlib.pyplot as plt

file_path = "D:\信管专业\社会计算\社会计算小组作业\评论update.xlsx"
excel_file = pd.ExcelFile(file_path)
df = excel_file.parse('有标记的')

# 情感映射
def map_sentiment(row):
    if row['sentiment'] in ['快乐', '赞美', '感动']:
        valence = 5
    elif row['sentiment'] in ['悲伤', '厌恶', '恐惧', '愤怒']:
        valence = 0
    elif row['sentiment'] in ['惊讶', '疑惑', '对比']:
        valence = 2
    else:
        valence = 2.5
    return valence

def map_user_origin(origin):
    if origin == '中国用户':
        return '中国用户'
    elif origin == '外国用户':
        return '外国用户'
    else:
        return '未知'

df['Valence'] = df.apply(map_sentiment, axis=1)
df['user_origin'] = df['user_origin'].apply(map_user_origin)
df['评论时间'] = pd.to_datetime(df['评论时间'])
df['日期'] = df['评论时间'].dt.date
df['日期'] = pd.to_datetime(df['日期'], errors='coerce').dropna().dt.date


topic_daily = df.groupby(['笔记topic', '日期']).agg(
    声量=('评论ID', 'count'),
    情感均值=('Valence', 'mean')
).reset_index()

# 话题生命周期计算
def safe_life_cycle(x):
    valid_dates = [i for i in x if isinstance(i, date)]
    if valid_dates:
        return (max(valid_dates) - min(valid_dates)).days
    else:
        return 0

# 计算整体话题指标
topic_overall = df.groupby('笔记topic').agg(
    总声量=('评论ID', 'count'),
    平均情感=('Valence', 'mean'),
    参与人数=('user_origin', 'nunique'),
    话题生命周期=('日期', safe_life_cycle)
).reset_index()

# 设置阈值线，可根据自己需求在此处更改阈值
plt.figure(figsize=(12, 8))
high_volume_threshold = topic_overall['总声量'].quantile(0.70)
positive_threshold = 2.5
plt.axvline(x=positive_threshold, color='gray', linestyle='--', alpha=0.7)
plt.axhline(y=high_volume_threshold, color='gray', linestyle='--', alpha=0.7)

# 按声量和情感倾向的阈值对话题分类
conditions = [
    (topic_overall['总声量'] >= high_volume_threshold) & (topic_overall['平均情感'] >= positive_threshold),
    (topic_overall['总声量'] < high_volume_threshold) & (topic_overall['平均情感'] >= positive_threshold),
    (topic_overall['总声量'] >= high_volume_threshold) & (topic_overall['平均情感'] < positive_threshold),
    (topic_overall['总声量'] < high_volume_threshold) & (topic_overall['平均情感'] < positive_threshold)
]
choices = ['优质话题(推荐)', '高潜力话题(关注)', '敏感话题(减少讨论)', '小众话题(降低关注)']
topic_overall['分类'] = np.select(conditions, choices, default='未分类')

colors = {'优质话题(推荐)': 'green', '高潜力话题(关注)': 'blue',
          '敏感话题(减少讨论)': 'red', '小众话题(降低关注)': 'gray'}
for category, group in topic_overall.groupby('分类'):
    plt.scatter(
        group['平均情感'],
        group['总声量'],
        s=group['参与人数'] * 10,
        alpha=0.6,
        c=colors[category],
        label=category
    )

for i, row in topic_overall.iterrows():
    plt.annotate(row['笔记topic'], (row['平均情感'], row['总声量']),
                 xytext=(5, 5), textcoords='offset points')

plt.rcParams['figure.dpi'] = 100
plt.rcParams['font.sans-serif'] = ['SimHei']
plt.rcParams['axes.unicode_minus'] = False
plt.xlabel('平均情感倾向')
plt.ylabel('话题声量')
plt.title('话题声量 - 情感分布图')
plt.legend()
plt.grid(True, alpha=0.3)
plt.tight_layout()
plt.show()

recommendations = topic_overall.sort_values(by=['总声量', '平均情感'],ascending=[False, False])
print("话题推荐策略：")
print(recommendations[['笔记topic', '分类', '总声量', '平均情感']])

try:
    quadrants = {
        '优质话题(推荐)': recommendations[recommendations['分类'] == '优质话题(推荐)'].iloc[0],
        '高潜力话题(关注)': recommendations[recommendations['分类'] == '高潜力话题(关注)'].iloc[0],
        '敏感话题(减少讨论)': recommendations[recommendations['分类'] == '敏感话题(减少讨论)'].iloc[0],
        '小众话题(降低关注)': recommendations[recommendations['分类'] == '小众话题(降低关注)'].iloc[0]
    }
except IndexError:
    print("部分分类中没有话题，无法输出各象限代表话题")
else:
    print("\n各象限代表话题：")
    for quad, data in quadrants.items():
        print(f"{quad}: {data['笔记topic']} (声量={data['总声量']}, 情感={data['平均情感']:.2f})")

#------------------------------------------------------------------------------------------此线下方为对话题日声量和日情感的分析
def volume_sentiment_analyse(topic):
    sample_topic = topic
    topic_data = topic_daily[topic_daily['笔记topic'] == sample_topic]
    fig, ax1 = plt.subplots(figsize=(12, 6))

    # 声量图
    ax1.bar(topic_data['日期'], topic_data['声量'], color='skyblue', alpha=0.7)
    ax1.set_xlabel('日期')
    ax1.set_ylabel('日声量', color='skyblue')
    ax1.tick_params(axis='y', labelcolor='skyblue')
    #折线图
    ax2 = ax1.twinx()
    ax2.plot(topic_data['日期'], topic_data['情感均值'], color='coral', marker='o')
    ax2.set_ylabel('情感倾向', color='coral')
    ax2.tick_params(axis='y', labelcolor='coral')
    ax2.axhline(y=0, color='gray', linestyle='--')
    ax2.axhline(y=2.5, color='r', linestyle='--', label='情感阈值线')

    for x, y in zip(topic_data['日期'], topic_data['情感均值']):
        color = 'b' if y < 2.5 else 'coral'
        ax2.plot(x, y, marker='o', color=color)

    plt.title(f'"{sample_topic}" 话题演化趋势')
    plt.grid(True, alpha=0.3)
    plt.tight_layout()
    plt.show()

# volume_sentiment_analyse("Music")
# volume_sentiment_analyse("cattax")
# volume_sentiment_analyse("remain")
# volume_sentiment_analyse("friend")
# volume_sentiment_analyse("daily")
# volume_sentiment_analyse("物价对比")
# volume_sentiment_analyse("取名")
# volume_sentiment_analyse("learn")
# volume_sentiment_analyse("communicate")