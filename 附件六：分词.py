import pandas as pd
import jieba
from wordcloud import WordCloud
import matplotlib.pyplot as plt

plt.rcParams['font.sans-serif'] = ['SimHei']
plt.rcParams['axes.unicode_minus'] = False
file_path = "D:\信管专业\社会计算\社会计算小组作业\评论update.xlsx"
excel_file = pd.ExcelFile(file_path)
df = excel_file.parse('有标记的')

def word_cloud(Topic):
    music_df = df[df['笔记topic'] == Topic]
    text = ' '.join(music_df['评论内容'].astype(str))
    text = text.lower()
    words = jieba.lcut(text)
    stopwords = set(["啊啊啊","这个","哈哈哈哈","哈哈","哈哈哈","哈","是不是","就是",'的', '了', '在', '是', '我', '有', '和', '就', '不', '人', '都', '一', '个', '上', '也', '很', '到', '说', '要', '去', '你', '会', '着', '没有', '看', '好', '自己', '这', '那',
                     'really','also','hello','very','or','so','but','and', 'the', 'a', 'an', 'is', 'are', 'was', 'were', 'be', 'been', 'being', 'have', 'has', 'had', 'do', 'does', 'did', 'will', 'would', 'could', 'should', 'can', 'this', 'that', 'it', 'its', 'he', 'she', 'they', 'them', 'their', 'we', 'our', 'you', 'your', 'i', 'me', 'my', 'mine', 'to', 'of', 'in', 'on', 'at', 'for', 'with', 'about', 'as', 'into', 'like', 'through', 'after', 'over', 'between', 'out', 'against', 'during', 'without', 'before', 'under', 'around', 'among'])
    filtered_words = [word for word in words if word not in stopwords and len(word) > 1]

    from collections import Counter
    word_counts = Counter(filtered_words)
    total_words = len(word_counts)
    top_n = max(1, total_words // 10) #前百分比参数
    top_words = word_counts.most_common(top_n)
    word_freq = dict(top_words)

    wc = WordCloud(
        width=800,
        height=400,
        background_color='white',
        font_path='simhei.ttf',
        max_words=100
    ).generate_from_frequencies(word_freq)

    plt.figure(figsize=(10, 5))
    plt.imshow(wc, interpolation='bilinear')
    plt.axis('off')
    plt.title(f"\"{Topic}\"话题词云 (词频排名前10%)")
    plt.show()

print('数据基本信息：')
df.info()
print(df['笔记topic'].unique())
# word_cloud("Music")
# word_cloud("cattax")
word_cloud("remain")
# word_cloud("friend")
# word_cloud("daily")
# word_cloud("物价对比")
# word_cloud("取名")
# word_cloud("learn")
# word_cloud("communicate")