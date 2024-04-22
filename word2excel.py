from docx import Document
import re
import pandas as pd

# 读取Word文档
doc = Document(r"E:\mhyy\项目资料\大模型学习资料内容\qt\question01.docx")

# 存储题目的列表
questions = []
answers = []
# 题目正则表达式，匹配以普通数字（至少一位）+顿号开头的字符串
timu_regex = re.compile(r"\d+[、]+")


for i, paragraph in enumerate(doc.paragraphs):
    line = paragraph.text.strip()

    if not line:
        continue

    if timu_regex.match(line):
        print(f"题目 {i + 1}: {line}")
        if answers:  # 当前问题有答案时，将问题和答案添加到questions列表
            questions.append((current_question, "\n".join(answers)))
        answers = []  # 清空答案列表，准备收集新问题的答案
        current_question = line
    else:
        print(f"答案: {line}")
        answers.append(line)

# 处理最后一个问题（没有后续题目跟随）
if answers:
    questions.append((current_question, "\n".join(answers)))

# 将结果转换为DataFrame并保存到Excel文件
df = pd.DataFrame(questions, columns=["题目", "答案"])
df.to_excel(r"E:\mhyy\项目资料\结果01.xlsx", index=False)
