import requests
import re
import pandas as pd
import time
from tqdm import tqdm

# 硅基流动 API 配置
URL = "https://api.siliconflow.cn/v1/chat/completions"
TOKEN = "……"  # 请替换为您的实际API密钥
# 配置路径
INPUT_EXCEL = "D:/111PythonLearning/data/comments.xlsx"  # 输入文件路径
OUTPUT_EXCEL = "D:/111PythonLearning/data/deal/result6.xlsx"  # 输出文件路径
# 精心设计的Prompt模板
PROMPT = """
你是一个社交媒体数据分析专家，正在分析TikTok被禁期间外国网友涌入小红书平台的用户评论。请根据以下要求分析评论：
分析维度：
1. sentiment：[快乐, 悲伤, 厌恶, 恐惧, 愤怒, 惊讶, 赞美, 感动, 疑惑, 对比]
   若不属于以上任何一类，标记为"中性"
2. user_origin（根据语义判断）：只能是[中国用户, 外国用户, 未知]中的一个
   - 中国用户：使用本土化表达、熟悉小红书文化、以主人姿态发言
   - 外国用户：表达不熟悉、跨文化视角、游客心态
3. 情感维度评分（0-5整数）：
   - Valence：情感积极程度（0=非常负面，5=非常积极）
   - Arousal：情感强烈程度（0=平静，5=兴奋）
   - Dominance：控制感程度（0=无助，5=掌控）
当前分析对象：
笔记topic: {topic}
评论内容: {comment}

输出示例："sentiment":"感动","user_origin":"外国用户","valence":4,"arousal":3,"dominance":2
"""
RETRY_TIMES = 3  # 失败重试次数
SLEEP_SECONDS = 2  # 每次请求间隔

def call_api(topic, comment,timeout=2):
    prompt = PROMPT.format(topic=topic,comment=comment)
    payload = {
        "model": "deepseek-ai/DeepSeek-V3",
        "messages": [{"role": "user", "content": prompt}],
        "max_tokens": 200,
        "temperature": 0.3,
        "top_p": 0.8
    }
    HEADERS = {
        "Authorization": f"Bearer {TOKEN}",
        "Content-Type": "application/json"
    }
    try:
        response = requests.request("POST", URL, json=payload, headers=HEADERS,timeout=timeout)
        response.raise_for_status()
        response_text = response.json()['choices'][0]['message']['content']
        return response_text
    except:
        return "API调用失败"  # 失败时返回固定标识

def parse_response(generated_text):
    if generated_text == "API调用失败":
        return (0, 0, 0, 0, 0)
    # 优化正则表达式，适配JSON格式
    patterns = {
        # 匹配带引号的字符串值，允许值中包含空格和中文
        "sentiment": r'"sentiment"\s*:\s*"([^"]+)"',
        "user_origin": r'"user_origin"\s*:\s*"([^"]+)"',
        # 匹配数字值
        "valence": r'"valence"\s*:\s*(\d+)',
        "arousal": r'"arousal"\s*:\s*(\d+)',
        "dominance": r'"dominance"\s*:\s*(\d+)'
    }
    result = []
    for key in ["sentiment", "user_origin", "valence", "arousal", "dominance"]:
        match = re.search(patterns[key], generated_text)
        if match:
            # 对于数字类型进行转换
            if key in ["valence", "arousal", "dominance"]:
                result.append(int(match.group(1)))
            else:
                result.append(match.group(1))
        else:
            result.append(0)  # 找不到则记为0
    return tuple(result)

def process_excel(INPUT_EXCEL, OUTPUT_EXCEL, batch_size=20, max_retries=2):
    # 1. 读取输入数据
    try:
        df= pd.read_excel(INPUT_EXCEL)
        df_input = df.iloc[8000:9336]  #控制条数
        required_cols = ['笔记topic', '评论内容']
        if not set(required_cols).issubset(df_input.columns):
            print(f"错误：必须包含列：{required_cols}")
            return
        total_rows = len(df_input)
        print(f"共 {total_rows} 条数据，开始处理...")
    except Exception as e:
        print(f"读取输入失败：{e}")
        return

    # 2. 初始化结果存储（内存暂存+续传支持）
    result_data = []
    processed_count = 0
    try:
        df_existing = pd.read_excel(OUTPUT_EXCEL)
        result_data = df_existing.values.tolist()
        processed_count = len(df_existing)
        print(f"从第 {processed_count+1} 行继续处理")
    except:
        result_data.append(['笔记topic', '评论内容', 'sentiment', 'user_origin', 'valence', 'arousal', 'dominance'])

    # 3. 核心处理循环（带超时和实时反馈）
    try:
        for i in tqdm(range(processed_count, total_rows), desc="处理进度", initial=processed_count):
            row = df_input.iloc[i]
            topic = str(row['笔记topic']) if pd.notna(row['笔记topic']) else ""
            comment = str(row['评论内容']) if pd.notna(row['评论内容']) else ""
            retries = 0
            success = False

            while retries <= max_retries and not success:
                try:
                    # 调用API（带超时）
                    api_result = call_api(topic, comment, timeout=10)  # 10秒超时

                    # 解析结果
                    if "API错误" in api_result:
                        # API调用失败，直接记录
                        result_data.append([topic, comment, api_result, "", "", "", ""])
                    else:
                        sentiment, user_origin, valence, arousal, dominance = parse_response(api_result)
                        result_data.append([topic, comment, sentiment, user_origin, valence, arousal, dominance])

                    success = True
                    processed_count += 1

                    # 批量保存（减小文件操作频率）
                    if processed_count % batch_size == 0 or i == total_rows - 1:
                        pd.DataFrame(result_data).to_excel(OUTPUT_EXCEL, index=False, header=False)
                        tqdm.write(f"已保存至 {OUTPUT_EXCEL}（共 {processed_count} 行）")

                except requests.exceptions.Timeout:
                    # 超时错误（重点处理，避免卡住）
                    retries += 1
                    if retries <= max_retries:
                        tqdm.write(f"第{i+1}行超时，重试 {retries}/{max_retries}")
                        time.sleep(3 * retries)  # 重试间隔递增
                    else:
                        tqdm.write(f"第{i+1}行超时重试耗尽，标记为失败")
                        result_data.append([topic, comment, "超时失败", "", "", "", ""])
                        processed_count += 1
                        success = True
                except Exception as e:
                    # 其他错误直接记录
                    tqdm.write(f"第{i+1}行错误：{str(e)}")
                    result_data.append([topic, comment, f"处理错误：{e}", "", "", "", ""])
                    processed_count += 1
                    success = True

        # 最终保存
        pd.DataFrame(result_data).to_excel(OUTPUT_EXCEL, index=False, header=False)
        print(f"\n全部完成！处理 {processed_count} 行，结果：{OUTPUT_EXCEL}")

    except Exception as e:
        # 致命错误时立即保存
        pd.DataFrame(result_data).to_excel(OUTPUT_EXCEL, index=False, header=False)
        print(f"\n程序中断：{e}，已保存 {processed_count} 行结果")

if __name__ == "__main__":
    process_excel(
        INPUT_EXCEL,
        OUTPUT_EXCEL,
        batch_size=80,  # 更小的批量，更频繁保存
        max_retries=2
    )
