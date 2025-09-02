import os
import re
import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
import seaborn as sns
from datetime import datetime
import matplotlib.dates as mdates
from collections import defaultdict

plt.rcParams["font.family"] = ["SimHei", "WenQuanYi Micro Hei", "Heiti TC"]
plt.rcParams["axes.unicode_minus"] = False  


def clean_column(series):
    def extract_number(text):
        if pd.isna(text):
            return 0
        
        if isinstance(text, (int, float)):
            return int(text) if not pd.isna(text) else 0
        
        text = str(text).strip()
        if text == '' or text.lower() in ['nan', 'null', 'none']:
            return 0
        
        if '万' in text:
            try:
                num = float(re.findall(r'[\d.]+', text)[0]) * 10000
                return int(num)
            except:
                return 0
        elif 'k' in text.lower():
            try:
                num = float(re.findall(r'[\d.]+', text)[0]) * 1000
                return int(num)
            except:
                return 0
        else:
            numbers = re.findall(r'\d+', text)
            return int(numbers[0]) if numbers else 0
    
    return series.apply(extract_number)

# 数据加载与预处理
def load_data(data_dir="xhs_data"):
    
    # 读取所有Excel文件
    all_files = [f for f in os.listdir(data_dir) if f.endswith(('.xlsx', '.xls'))]
    if not all_files:
        raise ValueError(f"在 '{data_dir}' 中未找到Excel文件")
    
    data = {
        "posts": None,  
        "comments": None  
    }
    
    for file in all_files:
        file_path = os.path.join(data_dir, file)
        try:
            df = pd.read_excel(file_path)
            if "帖子" in file or "笔记" in file:
                if data["posts"] is None:
                    data["posts"] = df
                else:
                    data["posts"] = pd.concat([data["posts"], df], ignore_index=True)
            elif "评论" in file or "result" in file:
                if data["comments"] is None:
                    data["comments"] = df
                else:
                    data["comments"] = pd.concat([data["comments"], df], ignore_index=True)
            print(f"成功加载 {file}，数据量: {len(df)} 行")
        except Exception as e:
            print(f"加载 {file} 失败: {str(e)}")

    def classify_user_type(ip_location):
        if not isinstance(ip_location, str):
            return "未知"
        
        ip_location = ip_location.strip()
        
        # 定义外国地区列表
        foreign_locations = [
            "美国", "澳大利亚", "比利时", "意大利", "加拿大", "英国", "法国", "德国", "新加坡", "日本", "韩国", "俄罗斯", "西班牙", "荷兰",
            "瑞典", "挪威", "丹麦", "芬兰"
        ]
        
        # 判断是否为外国用户
        for location in foreign_locations:
            if location in ip_location:
                return "外国用户"
        
        # 如果包含中国省市名称，判断为中国用户
        chinese_locations = [
            "北京", "上海", "天津", "重庆", "河北", "山西", "辽宁", "吉林", "黑龙江",
            "江苏", "浙江", "安徽", "福建", "江西", "山东", "河南", "湖北", "湖南",
            "广东", "海南", "四川", "贵州", "云南", "陕西", "甘肃", "青海", "台湾",
            "内蒙古", "广西", "西藏", "宁夏", "新疆", "香港", "澳门", "中国"
        ]
        
        for location in chinese_locations:
            if location in ip_location:
                return "中国用户"
        return "未知"
    
    # 数据预处理
    if data["posts"] is not None:
        posts_df = data["posts"]
        
        if "发布时间" in posts_df.columns:
            posts_df["发布时间"] = pd.to_datetime(posts_df["发布时间"], errors="coerce")
        
        numeric_columns = ["点赞数", "评论数", "收藏数"]
        for col in numeric_columns:
            if col in posts_df.columns:
                posts_df[col] = clean_column(posts_df[col])
                print(f"清洗 {col} 列完成")
        
        if "IP地址" in posts_df.columns:
                posts_df["用户类型"] = posts_df["IP地址"].apply(classify_user_type)
        elif "IP属地" in posts_df.columns:
                posts_df["用户类型"] = posts_df["IP属地"].apply(classify_user_type)
        else:
                posts_df["用户类型"] = "未知"
            
        data["posts"] = posts_df
        
        if data["comments"] is not None:
            comments_df = data["comments"]
        
        if "评论时间" in comments_df.columns:
            comments_df["评论时间"] = pd.to_datetime(comments_df["评论时间"], errors="coerce")
        
        if "点赞数" in comments_df.columns:
            comments_df["点赞数"] = clean_column(comments_df["点赞数"])
        
        if "IP地址" in comments_df.columns:
            comments_df["评论用户类型"] = comments_df["IP地址"].apply(
                lambda x: "外国用户" if isinstance(x, str) and re.search(r'[A-Za-z]', x) else "中国用户"
            )
        elif "IP属地" in comments_df.columns:
            comments_df["评论用户类型"] = comments_df["IP属地"].apply(
                lambda x: "外国用户" if isinstance(x, str) and re.search(r'[A-Za-z]', x) else "中国用户"
            )
        
        data["comments"] = comments_df
    
    print(f"\n数据加载完成 - 帖子数: {len(data['posts']) if data['posts'] is not None else 0}, "
          f"评论数: {len(data['comments']) if data['comments'] is not None else 0}")
    return data

# 双语内容互动优势分析
def analyze_bilingual_advantage(data):
    print("\n开始分析双语内容的互动优势")
    
    posts_df = data["posts"].copy()
    
    required_cols = ["笔记详情", "评论数", "点赞数", "用户类型"]
    if not all(col in posts_df.columns for col in required_cols):
        print(f"缺少 {required_cols}")
        return
    
    def detect_language_type(text):
        if not isinstance(text, str) or text.strip() == "":
            return "未知"
        
        # 统计中英文比例
        chinese_chars = len(re.findall(r'[\u4e00-\u9fff]', text))  # 中文字符
        english_chars = len(re.findall(r'[A-Za-z]', text))  # 英文字符
        
        if chinese_chars > 0 and english_chars > 0:
            return "双语混合"
        elif chinese_chars > 0:
            return "纯中文"
        elif english_chars > 0:
            return "纯英文"
        else:
            return "未知"
    
    # 添加语言类型列
    posts_df["语言类型"] = posts_df["笔记详情"].apply(detect_language_type)
    valid_df = posts_df[posts_df["语言类型"] != "未知"].copy()
    
    # 计算互动率，定义浏览量=点赞+收藏+评论
    if "收藏数" in valid_df.columns:
        valid_df["浏览量估算"] = valid_df["点赞数"] + valid_df["评论数"] + valid_df["收藏数"]
    else:
        valid_df["浏览量估算"] = valid_df["点赞数"] + valid_df["评论数"]

    valid_df = valid_df[valid_df["浏览量估算"] > 0]
    
    valid_df["评论互动率"] = valid_df["评论数"] / valid_df["浏览量估算"]
    valid_df["点赞互动率"] = valid_df["点赞数"] / valid_df["浏览量估算"]
    
    result = valid_df.groupby(["用户类型", "语言类型"]).agg({
        "评论互动率": ["mean", "count"],
        "点赞互动率": ["mean"]
    }).round(4)
    
    result.columns = ["平均评论互动率", "样本数", "平均点赞互动率"]
    result = result.reset_index()
    
    # 可视化
    plt.figure(figsize=(14, 8))
    
    # 评论互动率
    plt.subplot(2, 1, 1)
    sns.barplot(data=result, x="语言类型", y="平均评论互动率", hue="用户类型", palette="Set2")
    plt.title("不同语言类型内容的平均评论互动率", fontsize=14)
    plt.ylabel("评论互动率 (评论数/浏览量)")
    plt.xticks(rotation=0)
    for i, row in enumerate(result.itertuples()):
        plt.text(i % 3, row.平均评论互动率 + 0.01, f"n={row.样本数}", 
                 ha='center', va='bottom', fontsize=10)
    
    # 点赞互动率
    plt.subplot(2, 1, 2)
    sns.barplot(data=result, x="语言类型", y="平均点赞互动率", hue="用户类型", palette="Set2")
    plt.title("不同语言类型内容的平均点赞互动率", fontsize=14)
    plt.ylabel("点赞互动率 (点赞数/浏览量)")
    plt.xticks(rotation=0)
    
    plt.tight_layout()
    plt.savefig("双语内容互动优势分析.png", dpi=300, bbox_inches="tight")
    plt.show()
    
    return result


# 内容形式的中外偏好差异分析
def analyze_content_preference(data):
    print("\n开始分析内容形式的中外偏好差异")
    if data["posts"] is None:
        print("没有帖子数据，无法进行分析")
        return
    
    posts_df = data["posts"].copy()

    required_cols = ["发布时间", "笔记类型", "用户类型"]
    if not all(col in posts_df.columns for col in required_cols):
        print(f"缺少{required_cols}")
        return
    
    valid_df = posts_df.dropna(subset=["发布时间", "笔记类型", "用户类型"]).copy()
    
    def simplify_note_type(note_type):
        if not isinstance(note_type, str):
            return "其他"
        note_type = note_type.lower()
        if "视频" in note_type:
            return "视频"
        elif "图文" in note_type or "文字" in note_type or "图片" in note_type:
            return "图文"
        else:
            return "其他"
    
    valid_df["内容类型"] = valid_df["笔记类型"].apply(simplify_note_type)
    
    # 按月份聚合数据
    valid_df["发布月份"] = valid_df["发布时间"].dt.to_period("M")
    valid_df = valid_df.sort_values("发布月份")
    
    # 按用户类型、月份和内容类型分组统计
    monthly_data = valid_df.groupby(["用户类型", "发布月份", "内容类型"]).size().reset_index(name="数量")
    
    # 计算每月各类内容占比
    total_per_month = valid_df.groupby(["用户类型", "发布月份"]).size().reset_index(name="总数量")
    monthly_data = pd.merge(monthly_data, total_per_month, on=["用户类型", "发布月份"])
    monthly_data["占比"] = monthly_data["数量"] / monthly_data["总数量"]
    
    # 转换为时间戳以便绘图
    monthly_data["发布日期"] = monthly_data["发布月份"].dt.to_timestamp()
    
    # 可视化
    plt.figure(figsize=(14, 8))
    
    # 中国用户内容类型偏好
    chinese_data = monthly_data[monthly_data["用户类型"] == "中国用户"]
    plt.subplot(2, 1, 1)
    sns.lineplot(data=chinese_data, x="发布日期", y="占比", hue="内容类型", marker='o', palette="Set1")
    plt.title("中国用户内容类型偏好变化", fontsize=14)
    plt.ylabel("内容类型占比")
    plt.ylim(0, 1)
    plt.xticks(rotation=45)
    plt.grid(alpha=0.3)
    
    # 外国用户内容类型偏好
    foreign_data = monthly_data[monthly_data["用户类型"] == "外国用户"]
    plt.subplot(2, 1, 2)
    sns.lineplot(data=foreign_data, x="发布日期", y="占比", hue="内容类型", marker='s', palette="Set1")
    plt.title("外国用户内容类型偏好变化", fontsize=14)
    plt.ylabel("内容类型占比")
    plt.ylim(0, 1)
    plt.xticks(rotation=45)
    plt.grid(alpha=0.3)
    
    plt.tight_layout()
    plt.savefig("内容形式中外偏好差异分析.png", dpi=300, bbox_inches="tight")
    plt.show()
    
    return monthly_data

# 关键事件对情感的冲击分析
def analyze_emotion_impact(data):
    print("\n开始分析关键事件对情感的冲击")
    if data["posts"] is None:
        print("没有帖子数据，无法进行分析")
        return
    
    posts_df = data["posts"].copy()
    
    required_cols = ["发布时间", "笔记详情", "用户类型"]
    if not all(col in posts_df.columns for col in required_cols):
        print(f"缺少{required_cols}")
        return
    
    valid_df = posts_df.dropna(subset=["发布时间", "笔记详情"]).copy()
    valid_df = valid_df[valid_df["发布时间"].dt.year == 2025].copy()
    # 定义情感关键词
    emotion_keywords = {
        "焦虑恐惧": [
            "害怕", "担心", "焦虑", "恐惧", "担忧", "不安", "紧张", "慌", "慌张", "忧虑",
            "恐慌", "害怕", "怕", "慌乱", "忐忑", "不踏实", "紧张", "压力", "压抑",
            "绝望", "崩溃", "完了", "怎么办", "末日", "危险", "风险", "威胁",
            "panic", "anxious", "worried", "fear", "afraid", "scary", "nervous", 
            "stress", "concern", "anxiety", "terrified", "frightened", "upset",
            "overwhelmed", "desperate", "crisis", "danger", "threat", "risk"
        ],
        "积极乐观": [
            "庆幸", "开心", "期待", "高兴", "希望", "兴奋", "激动", "欣慰", "满意",
            "幸福", "快乐", "喜悦", "乐观", "积极", "正能量", "美好", "棒", "好",
            "赞", "支持", "鼓励", "加油", "相信", "信心", "未来", "机会", "成功",
            "good", "happy", "glad", "excited", "looking forward", "hope", "great",
            "amazing", "wonderful", "awesome", "positive", "optimistic", "confident",
            "support", "encourage", "believe", "future", "opportunity", "success"
        ],
        "惊讶感慨": [
            "惊讶", "没想到", "对比", "惊讶的是", "震惊", "意外", "想不到", "竟然",
            "居然", "原来", "真的", "确实", "果然", "哇", "天啊", "我的天", "太",
            "感慨", "感叹", "变化", "差别", "不同", "反差", "对比",
            "surprise", "surprised", "amazing", "wow", "incredible", "unbelievable",
            "unexpected", "shocking", "astonishing", "compare", "comparison", "difference",
            "change", "contrast", "actually", "really", "truly"
        ],
        "适应融入": [
            "适应", "习惯", "融入", "学习", "了解", "体验", "尝试", "探索", "发现",
            "新", "不同", "文化", "生活", "朋友", "社交", "交流", "沟通", "互动",
            "分享", "帮助", "指导", "建议", "推荐", "介绍", "欢迎", "接受",
            "adapt", "adjust", "integrate", "learn", "explore", "discover", "new",
            "culture", "life", "friend", "social", "communicate", "share", "help",
            "guide", "recommend", "welcome", "accept", "experience", "try"
        ]
    }
    
    # 情感识别函数
    def detect_emotion_enhanced(text):
        if not isinstance(text, str) or len(text.strip()) < 5:
            return "未识别"
        
        text_lower = text.lower()
        emotion_scores = {}
        
        for emotion, keywords in emotion_keywords.items():
            score = 0
            for keyword in keywords:
                if keyword in text_lower:
                    weight = len(keyword) / 3.0 if len(keyword) > 2 else 1.0
                    score += weight
            emotion_scores[emotion] = score
        
        if max(emotion_scores.values()) > 0:
            dominant_emotion = max(emotion_scores, key=emotion_scores.get)
            return dominant_emotion
        else:
            return "中性"
    
    # 识别情感
    valid_df["情感类型"] = valid_df["笔记详情"].apply(detect_emotion_enhanced)
    
    all_emotion_data = []
    
    # 添加帖子数据
    posts_emotion = valid_df[["发布时间", "情感类型"]].copy()
    posts_emotion["数据来源"] = "帖子"
    all_emotion_data.append(posts_emotion)

    # 评论数据情感分析
    if data["comments"] is not None:
        comments_df = data["comments"].copy()
        if "评论内容" in comments_df.columns and "评论时间" in comments_df.columns:
            comments_valid = comments_df.dropna(subset=["评论内容", "评论时间"]).copy()
            comments_valid = comments_valid[comments_valid["评论时间"].dt.year == 2025].copy()
            
            comments_valid["情感类型"] = comments_valid["评论内容"].apply(detect_emotion_enhanced)
            
            comments_emotion = comments_valid[["评论时间", "情感类型"]].copy()
            comments_emotion = comments_emotion.rename(columns={"评论时间": "发布时间"})
            comments_emotion["数据来源"] = "评论"
            all_emotion_data.append(comments_emotion)
    
    if len(all_emotion_data) > 1:
        combined_df = pd.concat(all_emotion_data, ignore_index=True)
    else:
        combined_df = all_emotion_data[0]

    # 按日期聚合数据
    combined_df["发布日期"] = combined_df["发布时间"].dt.date
    combined_df = combined_df.sort_values("发布日期")
    
    # 定义关键事件日期点
    key_events = {
        "2025-01-17": "禁令合宪",
        "2025-01-19": "Tiktok从应用商店主动下架",
        "2025-01-20": "特朗普延期75天封禁",
        "2025-04-04": "特朗普再次延期75天",
        "2025-06-19": "特朗普再次延期90天",
    }
    
    # 按用户类型、日期和情感类型分组统计
    daily_emotion = combined_df.groupby(["发布日期", "情感类型"]).size().reset_index(name="数量")
    
    # 计算每日总条数
    daily_total = combined_df.groupby("发布日期").size().reset_index(name="总数量")
    daily_emotion = pd.merge(daily_emotion, daily_total, on="发布日期")
    daily_emotion["情感占比"] = daily_emotion["数量"] / daily_emotion["总数量"]
    
    # 过滤出有效情感数据
    emotion_df = daily_emotion[~daily_emotion["情感类型"].isin(["未识别", "中性"])]
    
    # 过滤出有情感标签的数据
    emotion_df = daily_emotion[daily_emotion["情感类型"] != "未识别"]
    
    emotion_df_smooth = emotion_df.copy()
    for emotion in emotion_df["情感类型"].unique():
        emotion_data = emotion_df[emotion_df["情感类型"] == emotion].sort_values("发布日期")
        if len(emotion_data) >= 3:
            emotion_df_smooth.loc[emotion_df_smooth["情感类型"] == emotion, "情感占比平滑"] = \
                emotion_data["情感占比"].rolling(window=3, center=True, min_periods=1).mean()
        else:
            emotion_df_smooth.loc[emotion_df_smooth["情感类型"] == emotion, "情感占比平滑"] = \
                emotion_data["情感占比"]
    
    # 可视化
    plt.figure(figsize=(16, 10))
    colors = {'焦虑恐惧': '#FF6B6B', '积极乐观': '#4ECDC4', '惊讶感慨': '#45B7D1', '适应融入': '#96CEB4'}
    
    # 绘制情感变化趋势
    for emotion in emotion_df["情感类型"].unique():
        emotion_data = emotion_df_smooth[emotion_df_smooth["情感类型"] == emotion].sort_values("发布日期")
        if not emotion_data.empty:
            plt.plot(emotion_data["发布日期"], emotion_data["情感占比平滑"], 
                    'o-', label=f'{emotion} (n={emotion_data["数量"].sum()})', 
                    color=colors.get(emotion, '#999999'), 
                    linewidth=2.5, markersize=5, alpha=0.8)
    
    plt.title("TikTok事件全体用户情感变化趋势", fontsize=16, fontweight='bold', pad=20)
    plt.ylabel("情感占比", fontsize=12)
    plt.xlabel("日期", fontsize=12)
   
    max_ratio = emotion_df["情感占比"].max()
    plt.ylim(0, min(max_ratio * 1.2, 1.0))
    
    # 添加关键事件标记
    for date_str, event in key_events.items():
        event_date = datetime.strptime(date_str, "%Y-%m-%d").date()
        if (event_date >= emotion_df["发布日期"].min() and 
            event_date <= emotion_df["发布日期"].max()):
            plt.axvline(x=event_date, color='red', linestyle='--', alpha=0.7, linewidth=2)
            plt.text(event_date, max_ratio * 1.1, event, 
                    rotation=45, verticalalignment='bottom', fontsize=10,
                    bbox=dict(boxstyle="round,pad=0.3", facecolor="yellow", alpha=0.7))
    
    plt.grid(alpha=0.3, linestyle='-', linewidth=0.5)
    plt.legend(loc='upper left', frameon=True, fancybox=True, shadow=True)
    
    from matplotlib.dates import DateFormatter, WeekdayLocator
    ax = plt.gca()
    ax.xaxis.set_major_formatter(DateFormatter('%m-%d'))
    ax.xaxis.set_major_locator(WeekdayLocator(interval=1))
    plt.xticks(rotation=45)
    
    plt.tight_layout()
    plt.savefig("关键事件对全体用户情感冲击分析.png", dpi=300, bbox_inches="tight")
    plt.show()
    
    return emotion_df

# 互助内容的类型与价值分析
def analyze_help_content(data):
    print("\n开始分析互助内容的类型与价值")
    if data["comments"] is None:
        print("没有评论数据，无法进行分析")
        return
    
    comments_df = data["comments"].copy()
    
    required_cols = ["评论内容", "点赞数"]
    if not all(col in comments_df.columns for col in required_cols):
        print(f"缺少{required_cols}")
        return
    
    valid_df = comments_df.dropna(subset=["评论内容", "点赞数"]).copy()
    
    # 定义互助内容类型
    help_categories = {
        "实用知识": ["教程", "方法", "步骤", "怎么", "如何", "攻略", "技巧", 
                  "teach", "how to", "guide", "method", "step"],
        "情感支持": ["欢迎", "加油", "支持", "鼓励", "开心", "理解", 
                  "welcome", "support", "encourage", "happy", "glad"],
        "娱乐互动": ["哈哈", "搞笑", "可爱", "有趣", "笑死", 
                  "funny", "cute", "haha", "lol", "interesting"]
    }
    
    # 内容类型分类函数
    def categorize_help_content(text):
        if not isinstance(text, str):
            return "其他"
        
        text_lower = text.lower()
        categories = []
        
        for category, keywords in help_categories.items():
            for keyword in keywords:
                if keyword in text_lower:
                    categories.append(category)
                    break  
        
        return ", ".join(categories) if categories else "其他"
    
    # 分类互助内容
    valid_df["互助类型"] = valid_df["评论内容"].apply(categorize_help_content)
    
    help_df = valid_df[valid_df["互助类型"] != "其他"].copy()
    
    # 统计各类互助内容的数量和平均点赞数
    help_stats = help_df.groupby("互助类型").agg({
        "评论内容": "count",
        "点赞数": "mean"
    }).rename(columns={"评论内容": "数量", "点赞数": "平均点赞数"})
    
    total = help_stats["数量"].sum()
    help_stats["占比"] = help_stats["数量"] / total
    
    help_stats = help_stats.sort_values("数量", ascending=False)
    help_stats = help_stats.round(4)
    
    
    # 可视化
    plt.figure(figsize=(14, 8))
    
    # 互助内容类型占比
    plt.subplot(2, 1, 1)
    help_stats["占比"].plot(kind="bar", color=['#5DA5DA', '#FAA43A', '#60BD68'])
    plt.title("互助内容类型占比分布", fontsize=14)
    plt.ylabel("占比")
    plt.xticks(rotation=0)
    for i, v in enumerate(help_stats["占比"]):
        plt.text(i, v + 0.01, f"{v:.2%}\n(n={help_stats['数量'][i]})", 
                 ha='center', va='bottom', fontsize=10)
    
    # 各类互助内容的平均点赞数
    plt.subplot(2, 1, 2)
    help_stats["平均点赞数"].plot(kind="bar", color=['#5DA5DA', '#FAA43A', '#60BD68'])
    plt.title("各类互助内容的平均点赞数", fontsize=14)
    plt.ylabel("平均点赞数")
    plt.xticks(rotation=0)
    for i, v in enumerate(help_stats["平均点赞数"]):
        plt.text(i, v + 0.1, f"{v:.1f}", ha='center', va='bottom', fontsize=10)
    
    plt.tight_layout()
    plt.savefig("互助内容类型与价值分析.png", dpi=300, bbox_inches="tight")
    plt.show()
    
    return help_stats

# 主函数
def main():
    print("=== 小红书'TikTok难民'事件数据分析 ===")
    
    try:
        data = load_data()
    except Exception as e:
        print(f"数据加载失败: {str(e)}")
        return
    
    # 执行各个角度的分析
    analyze_bilingual_advantage(data)
    analyze_content_preference(data)
    analyze_emotion_impact(data)
    analyze_help_content(data)
    
    print("\n所有分析完成")

if __name__ == "__main__":
    main()