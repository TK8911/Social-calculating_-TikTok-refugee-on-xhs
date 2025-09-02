import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
import seaborn as sns
from scipy.stats import pearsonr
from collections import defaultdict
import os

plt.rcParams["font.family"] = ["SimHei", "WenQuanYi Micro Hei", "Heiti TC"]


data_dir = "results" 
file_names = [f"result{i}.xlsx" for i in range(1, 7)] 

all_data = []
for file in file_names:
    file_path = os.path.join(data_dir, file)
    if not os.path.exists(file_path):
        print(f"警告：文件 {file_path} 不存在，已跳过")
        continue
    df = pd.read_excel(file_path)
    all_data.append(df)

df_combined = pd.concat(all_data, ignore_index=True)


valid_topics = [
    "tiktok难民", "cattax", "Music", "取名", "物价对比", 
    "daily", "learn", "communicate", "music", 
    "friend", "remain"
]
valid_sentiments = ["快乐", "悲伤", "厌恶", "恐惧", "愤怒", "惊讶", 
                    "赞美", "感动", "疑惑", "对比", "中性","失望"]
valid_origins = ["中国用户", "外国用户"]

df_clean = df_combined[
    (df_combined["topic"].isin(valid_topics)) &
    (df_combined["sentiment"].isin(valid_sentiments)) &
    (df_combined["user_origin"].isin(valid_origins))
].copy()

print(f"清洗后有效数据量：{len(df_clean)} 条")

# 计算情感共鸣强度
def calculate_topic_resonance(topic_group):
    cn_data = topic_group[topic_group["user_origin"] == "中国用户"]
    foreign_data = topic_group[topic_group["user_origin"] == "外国用户"]
    if len(cn_data) < 5 or len(foreign_data) < 5:  
        return None
    
    # 计算情感标签重合率
    cn_sentiments = set(cn_data["sentiment"].value_counts().index)
    foreign_sentiments = set(foreign_data["sentiment"].value_counts().index)
    common_sentiments = cn_sentiments & foreign_sentiments
    overlap_rate = len(common_sentiments) / len(cn_sentiments | foreign_sentiments) 
    
    # 计算三个情感维度相关性
    foreign_sample = foreign_data.sample(n=len(cn_data), replace=True, random_state=42)
    valence_corr, _ = pearsonr(cn_data["valence"], foreign_sample["valence"])
    arousal_corr, _ = pearsonr(cn_data["arousal"], foreign_sample["arousal"])
    dominance_corr, _ = pearsonr(cn_data["dominance"], foreign_sample["dominance"])
    
    return {
        "topic": topic_group["topic"].iloc[0],
        "cn_sample_size": len(cn_data),
        "foreign_sample_size": len(foreign_data),
        "sentiment_overlap_rate": round(overlap_rate, 2),  # 情感标签重合率
        "valence_corr": round(valence_corr, 2),  # 愉悦度相关性
        "arousal_corr": round(arousal_corr, 2),  # 唤醒度相关性
        "dominance_corr": round(dominance_corr, 2)  # 支配度相关性
    }

# 按主题分组计算共鸣强度
resonance_list = []
for topic, group in df_clean.groupby("topic"):
    resonance = calculate_topic_resonance(group)
    if resonance:
        resonance_list.append(resonance)

resonance_df = pd.DataFrame(resonance_list)
print("\n跨文化情感共鸣强度分析结果：")
print(resonance_df.sort_values("sentiment_overlap_rate", ascending=False))

# 共鸣强度可视化
plt.figure(figsize=(14, 6))
sns.barplot(
    data=resonance_df.sort_values("sentiment_overlap_rate", ascending=False),
    x="topic",
    y="sentiment_overlap_rate",
    palette="pastel"
)
plt.title("各主题跨文化情感标签重合率", fontsize=14)
plt.xlabel("主题")
plt.ylabel("情感标签重合率")
plt.xticks(rotation=45, ha="right")
plt.ylim(0, 1) 
plt.tight_layout()
plt.show()
