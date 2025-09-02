import requests
import re
import pandas as pd
import time
from tqdm import tqdm

# 硅基流动 API 配置
URL = "https://api.siliconflow.cn/v1/chat/completions"
TOKEN = "……"  # 替换为你的API密钥
INPUT_EXCEL = "D:/111PythonLearning/data/帖子.xlsx"  # 输入文件路径（需包含'combine_notes'列）
OUTPUT_EXCEL = "D:/111PythonLearning/data/deal/result_topic.xlsx"  # 输出文件路径

# 定义待匹配的主题词列表（与Prompt保持一致）
TOPIC_LIST = ["cattax", "communicate", "daily", "learn", "Music", "取名", "物价对比"]

# 优化后的Prompt（明确要求返回指定主题词或"不相关"）
PROMPT = """
你是社交媒体数据分析专家，需要分析帖子主题：
1. 从列表中选择与帖子内容最匹配的主题词：{topic_list}
2. 若均不匹配，返回"不相关"
3. 仅返回结果，不添加任何解释、标点或多余文字
帖子内容：{content}
"""
RETRY_TIMES = 3  # 失败重试次数
SLEEP_SECONDS = 2  # 请求间隔时间
BATCH_SIZE = 80  # 批量保存间隔

def call_api(content):
    """调用API获取主题匹配结果"""
    # 格式化Prompt，传入主题词列表和帖子内容
    formatted_prompt = PROMPT.format(
        topic_list=TOPIC_LIST,
        content=content
    )
    payload = {
        "model": "deepseek-ai/DeepSeek-V3",
        "messages": [{"role": "user", "content": formatted_prompt}],
        "max_tokens": 50,  # 只需返回一个词，减少冗余
        "temperature": 0.1,  # 降低随机性，确保结果稳定
        "top_p": 0.8
    }
    headers = {
        "Authorization": f"Bearer {TOKEN}",
        "Content-Type": "application/json"
    }
    try:
        response = requests.post(
            URL,
            json=payload,
            headers=headers,
            timeout=10  # 超时时间设为10秒
        )
        response.raise_for_status()  # 触发HTTP错误（如401、500）
        # 提取模型返回的文本（去除首尾空格）
        return response.json()['choices'][0]['message']['content'].strip()
    except Exception as e:
        print(f"API调用异常: {str(e)[:30]}")
        return None  # 失败时返回None


def parse_topic(generated_text):
    """解析模型返回结果，验证是否为有效主题词"""
    if not generated_text:
        return "API调用失败"
    # 去除可能的标点符号（如模型误加的引号、句号）
    cleaned_text = re.sub(r'[^\w\u4e00-\u9fa5]', '', generated_text)
    # 检查是否在指定主题词列表中（不区分大小写，适配"music"和"Music"）
    for topic in TOPIC_LIST:
        if cleaned_text.lower() == topic.lower():
            return topic  # 返回原始规范的主题词（如"Music"而非"music"）
    # 检查是否为"不相关"（同样忽略标点和大小写）
    if cleaned_text.lower() == "不相关".lower():
        return "不相关"
    # 其他情况视为匹配失败
    return "匹配失败"

def process_excel():
    """主函数：处理Excel文件，批量分析帖子主题"""
    # 1. 读取输入数据（仅保留必要的'combine_notes'列）
    try:
        df_input = pd.read_excel(INPUT_EXCEL)
        # 检查必要列是否存在
        if 'combine_notes' not in df_input.columns:
            print("错误：输入文件必须包含'combine_notes'列（帖子内容）")
            return
        total_rows = len(df_input)
        print(f"成功读取数据，共 {total_rows} 条帖子")
    except Exception as e:
        print(f"读取输入文件失败：{e}")
        return

    # 2. 初始化结果存储（支持续传）
    result_data = []
    processed_count = 0  # 已处理的行数

    # 尝试读取已有结果（续传）
    try:
        df_existing = pd.read_excel(OUTPUT_EXCEL)
        result_data = df_existing.values.tolist()
        processed_count = len(df_existing)
        print(f"检测到已有结果，将从第 {processed_count + 1} 行开始处理")
    except:
        # 首次运行：添加表头
        result_data.append(['combine_notes', 'matched_topic'])  # 帖子内容 | 匹配的主题

    # 3. 逐行处理帖子
    try:
        # 进度条：从已处理的行数开始
        pbar = tqdm(range(processed_count, total_rows), desc="处理进度", initial=processed_count)
        for i in pbar:
            row = df_input.iloc[i]
            # 获取帖子内容（处理空值）
            content = str(row['combine_notes']).strip() if pd.notna(row['combine_notes']) else ""

            # 调用API（带重试机制）
            api_result = None
            for retry in range(RETRY_TIMES):
                api_result = call_api(content)
                if api_result is not None:
                    break  # 成功获取结果，退出重试
                time.sleep(SLEEP_SECONDS * (retry + 1))  # 重试间隔递增

            # 解析结果
            matched_topic = parse_topic(api_result)

            # 保存到结果列表
            result_data.append([content, matched_topic])
            processed_count += 1

            # 批量保存（减少文件写入次数）
            if processed_count % BATCH_SIZE == 0 or i == total_rows - 1:
                pd.DataFrame(result_data).to_excel(OUTPUT_EXCEL, index=False, header=False)
                pbar.set_postfix({"已保存": f"{processed_count}行"})  # 进度条显示保存状态

            # 控制API调用频率
            time.sleep(SLEEP_SECONDS)

        # 最终保存
        pd.DataFrame(result_data).to_excel(OUTPUT_EXCEL, index=False, header=False)
        print(f"\n全部处理完成！结果已保存至：{OUTPUT_EXCEL}")
        print(f"输出格式：2列（combine_notes: 帖子内容, matched_topic: 匹配的主题）")

    except Exception as e:
        # 遇到致命错误时，立即保存已处理的结果
        pd.DataFrame(result_data).to_excel(OUTPUT_EXCEL, index=False, header=False)
        print(f"\n程序中断：{e}，已保存 {processed_count} 行结果")


if __name__ == "__main__":
    process_excel()
