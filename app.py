import streamlit as st
import docx
import openai
import pandas as pd

# 1. 网页基础配置（Streamlit 自动渲染前端）
st.set_page_config(page_title="地铁车站结构设计审查评价系统", layout="wide")
st.title("🚇 地铁车站结构设计审查评价系统 (AI Agent)")
st.caption("基于国家 5 大核心规范 —— 强条硬核管控版")

# 2. 侧边栏：配置大模型密钥与审查依据
with st.sidebar:
    st.header("⚙️ 审查配置中心")
    api_key = st.text_input("请输入大模型 API Key", type="password", help="请输入支持长文本的LLM API Key")
    base_url = st.text_input("API 端点 URL", value="https://api.deepseek.com/v1")
    model_name = st.text_input("大模型名称", value="deepseek-chat")
    
    st.divider()
    st.subheader("📚 核心校审规范依据")
    st.markdown("""
    1. **建质[2013]160号** 编制深度规定
    2. **发改基础[2015]49号** 规划建设管理通知
    3. **建标104-2008** 工程项目建设标准
    4. **GB50490-2009** 城市轨道交通技术规范
    5. **GB50157-2013** 地铁设计规范
    """)

# 3. 核心功能：Word 文档文本与表格数据联合提取器
def extract_word_content(file):
    doc = docx.Document(file)
    content_lines = []
    
    # 提取正文和标题
    for para in doc.paragraphs:
        if para.text.strip():
            content_lines.append(para.text.strip())
            
    # 提取表格数据（关键参数往往藏在表格里，必须提取）
    for table in doc.tables:
        for row in table.rows:
            row_data = [cell.text.strip() for cell in row.cells if cell.text.strip()]
            if row_data:
                content_lines.append(f"[表格数据] " + " | ".join(row_data))
                
    return "\n".join(content_lines)

# 4. 前端上传组件
uploaded_file = st.file_uploader("📂 请上传地铁车站设计文件、说明书或计算书 (.docx)", type=["docx"])

if uploaded_file:
    with st.spinner("正在深度解析 Word 结构数据..."):
        document_text = extract_word_content(uploaded_file)
    st.success(f"🎉 文档解析成功！成功提取结构文本数据共计 {len(document_text)} 个字符。")
    
    # 文本预览区
    with st.expander("🔍 查看解析出的文本片段预览"):
        st.text(document_text[:800] + "\n\n[... 后续内容已收入缓存 ...]")

    # 5. 一键触发强条审查
    if st.button("🚀 开始自动化规范强条审查", type="primary"):
        if not api_key:
            st.error("❌ 触发失败：请先在左侧边栏配置您的大模型 API Key。")
        else:
            with st.spinner("🕵️ AI 强条猎手正在全网络扫描违规风险，请稍候..."):
                try:
                    # 强条优先审查灵魂提示词
                    system_prompt = """你是一个极其严厉、铁面无私的“地铁车站结构设计强条审查官”。
你的核心使命是死死盯住文本中任何违反国家强制性条文（强条）的蛛丝马迹。

# 最高行为准则
1. 【强条一票否决】：强条审查具有最高优先级。在输出报告时，必须将“强条违规项”置顶，并用最醒目的红色/警示符号标出。
2. 【证据链闭环】：每指出一项强条违规，必须严格满足三个要素：[文档原文定位] + [违反的规范具体条文号及内容] + [明确的整改闭环要求]。
3. 【无证据不放行】：如果文档对某项强条涉及的参数（如抗浮系数、人防荷载、裂缝值）未作说明，直接判定为“缺失关键安全设计要素”，视同违规。

# 重点盯防领域
- 主体设计年限是否明确为100年，安全等级是否为一级。
- 是否明确抗浮设防水位，抗浮安全系数是否合规。
- 迎水面构件混凝土裂缝最大允许值是否严格控制在0.2mm以内（背水面0.3mm）。
- 是否包含人防等效静荷载组合验算和抗震设防烈度。

# 输出格式
开篇必须立刻呈现【🚨强条合规判定结果】，如果有违规，必须列出【🛑 强条违规/漏项清单】置顶。只有当无强条违规时，才允许在下方输出常规的设计深度与优化建议。"""

                    # 初始化大模型客户端
                    client = openai.OpenAI(api_key=api_key, base_url=base_url)
                    
                    # 调用大模型进行推理
                    response = client.chat.completions.create(
                        model=model_name,
                        messages=[
                            {"role": "system", "content": system_prompt},
                            {"role": "user", "content": f"请审查以下地铁车站设计文件内容，重点抓取强条：\n\n{document_text}"}
                        ],
                        temperature=0.0 # 强制大模型以 0 随机性严谨输出
                    )
                    
                    # 6. 前端渲染审查报告
                    st.header("📊 AI 自动化审查评价报告")
                    report_content = response.choices[0].message.content
                    
                    # 智能高亮处理：如果报告中包含违规字眼，给予大红框提示
                    if "违反" in report_content or "违规" in report_content or "致命" in report_content:
                        st.error("⚠️ 系统检测到当前设计文件存在违反强制性条文的风险，已自动触发风险阻断！")
                    else:
                        st.success("✅ 初步强条扫描未见明显违规，请结合人工复核。")
                        
                    st.markdown(report_content)
                    
                except Exception as e:
                    st.error(f"审查引擎运行中断，错误信息: {str(e)}")