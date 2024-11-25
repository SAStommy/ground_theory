from docx import Document
import gradio as gr
from openai import OpenAI
import subprocess
import os

client = OpenAI(
    api_key= os.getenv("OPENAI_API_KEY")
)    

def generategraph(code_string):
    if os.path.exists("grounded_theory_tree"):
        os.remove("grounded_theory_tree")
    if os.path.exists("grounded_theory_tree.png"):
        os.remove("grounded_theory_tree.png")
        
    # 将代码写入文件
    with open("functiongraph.py", "w", encoding="utf-8") as f:
        f.write(code_string)
    
    # 执行 Python 文件并捕获输出
    subprocess.run(["python", "functiongraph.py"], capture_output=True, text=True)
    
    #if not os.path.exists("grounded_theory_tree.png"):
    #    raise FileNotFoundError("grounded_theory_tree.png 文件未生成，请检查代码生成过程是否正确。")

# 读取 docx 文件的函数
def load_docx_data(files):
    content = []
    for file in files:
        doc = Document(file)
        # 获取每个文档的段落并添加到内容列表
        content.extend([para.text for para in doc.paragraphs if para.text])  # 只添加非空段落
        content.append("")  # 添加一个空字符串作为段落分隔符
    return content

# AI 进行扎根理论三层编码的函数
def grounded_theory_analysis(content):
    global open_coding, axial_coding, selective_coding
    dialog_history = [{"role": "system", "content": "你是紮根理論的研究專家, 你也能夠截取資料文本原文(句子不會進行任何修改, 句子會與資料文本完全相同)作為根據"}]
    
    open_prompt = f"请對前面由{len(content)}段组成的文本進行開放式編碼，并列出主要概念後面用()加上對應資料作為根據"
    for i, part in enumerate(content):
        if i == len(content) - 1:
            dialog_history.append({"role": "user", "content": open_prompt})
        else:
            dialog_history.append({"role": "user", "content": f"你需要閱讀内容，即將接受在{len(content)}中的第{i + 1}部分\n\n{part}"})
        
    # 调用 API
    response = client.chat.completions.create(
        model="gpt-4o-mini",
        messages=dialog_history
    )
    
    open_coding =  response.choices[0].message.content
    dialog_history.append({"role": "assistant", "content": open_coding})
    
    
    # 主轴编码
    axial_prompt = f"根據已完成開放式編碼结果，進行主軸編碼，建立類别之間的聯繫並後面用()加上對應資料開放式編碼作為根據：\n\n{open_coding}"
    dialog_history.append({"role": "user", "content": axial_prompt})
    axial_response = client.chat.completions.create(
        model="gpt-4o-mini",
        messages=dialog_history
    )
    axial_coding = axial_response.choices[0].message.content
    dialog_history.append({"role": "assistant", "content": axial_coding})
    
    # 选择性编码
    selective_prompt = f"根據已完成主軸編碼结果，進行選擇性編碼，提煉核心類别并形成理論框架後面用()加上對應資料主軸編碼作為根據：\n\n{axial_coding}"
    dialog_history.append({"role": "user", "content": selective_prompt})
    selective_response = client.chat.completions.create(
        model="gpt-4o-mini",
        messages=dialog_history
    )
    selective_coding = selective_response.choices[0].message.content
    dialog_history.append({"role": "assistant", "content": selective_coding})
    
    # 汇总结果
    result = f"**開放式編碼：**\n{open_coding}\n\n"
    result += f"**主軸編碼：**\n{axial_coding}\n\n"
    result += f"**選擇性編碼：**\n{selective_coding}"
    
    return result

def getgraphcode(result):
    code_response = client.chat.completions.create(
        model="gpt-4o",
        messages=[{"role": "user", "content": f"""根據以下資料:\n{result}\n生成用畫出紮根理論三層樹狀圖及圖片以grounded_theory_tree.png儲存的python程式碼, 其中程式碼一定要包括以下12行:
                   \ndot.attr("graph", fontname='SimHei', fontweight="bold")\n
                   dot.attr("node", fontname='SimHei', fontweight="bold")\n
                   dot.attr("edge", fontname='SimHei', fontweight="bold")\n
                   # 添加层名标记节点\n
                   dot.node("OpenCodingLabel", "開放式編碼", shape="plaintext", fontname='SimHei', fontweight="bold")\n
                   dot.node("AxialCodingLabel", "主軸式編碼", shape="plaintext", fontname='SimHei', fontweight="bold")\n
                   dot.node("SelectiveCodingLabel", "選擇性編碼", shape="plaintext", fontname='SimHei', fontweight="bold")\n
                   # 连接层名标记到每一层的所有节点\n
                   for node_id in open_coding_nodes:\n
                        dot.edge("OpenCodingLabel", node_id, style="dashed")\n
                   for node_id in axial_coding_nodes:\n
                        dot.edge("AxialCodingLabel", node_id, style="dashed")\n
                   for node_id in selective_coding_nodes:\n
                        dot.edge("SelectiveCodingLabel", node_id, style="dashed")\n
                   dot.render(filename='grounded_theory_tree', format='png')\n
                   只需要回傳python程式碼, 不需要任何描述"""}]
    )
    code_string = code_response.choices[0].message.content
    code_string = code_string.replace("```python", "").replace("```", "")
     # 删除 import 之前的所有内容
    import_index = code_string.find('import')
    if import_index != -1:
        code_string = code_string[import_index:]
    code_string = code_string.replace("import Digraph", "from graphviz import Digraph")
    return code_string

# Gradio 主函数
def main(files):
    if files is not None:
        content = load_docx_data(files)  # 注意这里是 files，而不是 file
        analysis_result = grounded_theory_analysis(content)
        generategraph(getgraphcode(analysis_result))
        return analysis_result, None, None
    else:
        return "请上传一个 .docx 文件。", None, None

# 创建 Gradio 接口
iface = gr.Interface(
    fn=main,  # 处理函数
    inputs=gr.File(label="上傳 .docx 文件", file_count="multiple"),  # 输入文件
    outputs=[gr.Markdown(), gr.Image(type="pil"), gr.File(label="下載圖片")],  # 输出文本，支持 Markdown 格式
    title="紮根理論分析器",
    description="上傳一個10萬字以內 .docx 文件的文件，使用 AI 實現紮根理論三層結構分析自動化，并顯示每層结果和結構圖。"
)

# 启动界面
iface.launch(share=True)
