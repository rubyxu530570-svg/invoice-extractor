import streamlit as st
import pandas as pd
import re
from datetime import datetime
import os
import tempfile
from io import BytesIO

# 尝试导入 PDF 和 OCR 相关库
try:
    import pdfplumber
except ImportError:
    st.error("缺少 pdfplumber，请在 requirements.txt 中添加")
    st.stop()

try:
    from paddleocr import PaddleOCR
    ocr_available = True
    ocr = PaddleOCR(use_angle_cls=True, lang="ch", show_log=False)
except Exception as e:
    ocr_available = False
    st.warning(f"PaddleOCR 未安装或加载失败，仅支持可复制文本的PDF: {e}")

# ========================
# 发票信息提取函数（已优化）
# ========================
def extract_invoice_info(text):
    """从文本中提取发票关键信息"""
    result = {
        "发票号码": "",
        "发票日期": "",
        "购买方名称": "",
        "项目名称": "",
        "价税合计": ""
    }

    # 1. 提取发票号码（固定18位数字）
    inv_num_match = re.search(r'发票号码[:：\s]*(\d{18})', text)
    if inv_num_match:
        result["发票号码"] = inv_num_match.group(1)

    # 2. 提取开票日期（支持多种格式）
    date_patterns = [
        r'开票日期[:：\s]*(\d{4}年\d{1,2}月\d{1,2}日)',
        r'开票日期[:：\s]*(\d{4}-\d{1,2}-\d{1,2})',
        r'(\d{4})年(\d{1,2})月(\d{1,2})日'
    ]
    
    for pattern in date_patterns:
        match = re.search(pattern, text)
        if match:
            if len(match.groups()) == 3:
                year, month, day = match.groups()
                result["发票日期"] = f"{year}-{int(month):02d}-{int(day):02d}"
            else:
                date_str = match.group(1).replace('年', '-').replace('月', '-').replace('日', '')
                try:
                    dt = datetime.strptime(date_str, '%Y-%m-%d')
                    result["发票日期"] = dt.strftime('%Y-%m-%d')
                except ValueError:
                    pass
            break

    # 3. 提取购买方名称（只匹配“名称:”后面的内容）
    buyer_match = re.search(r'名称[:：]\s*(.*?)(?:公司|集团|中心|店|厂)', text)
    if buyer_match:
        name = buyer_match.group(1).strip()
        clean_name = re.sub(r'[^\u4e00-\u9fa5a-zA-Z0-9]', '', name)
        result["购买方名称"] = clean_name

    # 4. 提取项目名称（优先匹配 * 开头的行）
    project_lines = []
    lines = [line.strip() for line in text.split('\n') if line.strip()]
    
    for line in lines:
        if re.match(r'^[\*\u4e00-\u9fa5]+ $ ', line) and not re.search(r'规格|型号|单位|数量|单价|金额|合计', line):
            project_lines.append(line)
        elif line.startswith('*'):
            project_lines.append(line)
    
    if project_lines:
        result["项目名称"] = "，".join(project_lines[:3])
    else:
        star_line = re.search(r'\*([^*]+)\*', text)
        if star_line:
            result["项目名称"] = star_line.group(1).strip()

    # 5. 提取价税合计（小写）
    total_match = re.search(r'(?:价税合计|合计)[（  $ ]小写[） $  ]?[:：\s]*[¥￥]?([\d,]+\.?\d*)', text)
    if total_match:
        amount_str = total_match.group(1).replace(',', '')
        try:
            float(amount_str)
            result["价税合计"] = amount_str
        except ValueError:
            pass

    return result

# ========================
# PDF 转文本（支持扫描件）
# ========================
def pdf_to_text(pdf_file):
    """将PDF转换为文本，优先尝试直接提取，失败则用OCR"""
    text = ""
    
    try:
        with pdfplumber.open(pdf_file) as pdf:
            for page in pdf.pages:
                page_text = page.extract_text()
                if page_text:
                    text += page_text + "\n"
    except Exception as e:
        st.warning(f"直接提取文本失败: {e}")
        text = ""

    if not text.strip() and ocr_available:
        st.info("检测到可能是扫描件，正在使用OCR识别...")
        try:
            from pdf2image import convert_from_bytes
            images = convert_from_bytes(pdf_file.read(), dpi=200)
            pdf_file.seek(0)
            for img in images:
                result = ocr.ocr(img, cls=True)
                if result and result[0]:
                    for line in result[0]:
                        text += line[1][0] + "\n"
        except Exception as e:
            st.error(f"OCR识别失败: {e}")
    
    return text

# ========================
# 主程序界面
# ========================
st.set_page_config(page_title="发票信息提取工具", layout="wide")
st.title("📊 发票信息自动提取工具")
st.markdown("上传多个增值税发票 PDF 文件，自动识别并生成 Excel 表格")

uploaded_files = st.file_uploader(
    "📁 请选择一个或多个发票 PDF 文件",
    type=["pdf"],
    accept_multiple_files=True
)

if uploaded_files:
    st.write(f"已上传 {len(uploaded_files)} 个文件")
    
    all_results = []
    
    for file in uploaded_files:
        with st.spinner(f"正在处理 {file.name}..."):
            try:
                text = pdf_to_text(file)
                if not text.strip():
                    st.warning(f"{file.name} 未提取到任何文字，请检查是否为有效发票。")
                    continue
                
                info = extract_invoice_info(text)
                info["文件名"] = file.name
                all_results.append(info)
                
            except Exception as e:
                st.error(f"处理 {file.name} 时出错: {e}")
    
    if all_results:
        df = pd.DataFrame(all_results)
        st.subheader("📋 提取结果预览")
        st.dataframe(df.fillna(""), use_container_width=True)
        
        output = BytesIO()
        with pd.ExcelWriter(output, engine='openpyxl') as writer:
            df.to_excel(writer, index=False, sheet_name='发票信息')
        excel_data = output.getvalue()
        
        st.download_button(
            label="📥 下载Excel文件",
            data=excel_data,
            file_name="发票信息汇总.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
    else:
        st.warning("未成功提取任何发票信息，请检查PDF内容或格式。")
else:
    st.info("请上传PDF文件开始处理。")
