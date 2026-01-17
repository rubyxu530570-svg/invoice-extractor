import streamlit as st
import pandas as pd
import re
from datetime import datetime
import os
import tempfile
from io import BytesIO

# 尝试导入 PDF 和 OCR 相关库（按需安装）
try:
    import pdfplumber
except ImportError:
    st.error("缺少 pdfplumber，请在 requirements.txt 中添加")
    st.stop()

try:
    from paddleocr import PaddleOCR
    ocr_available = True
    # 初始化 PaddleOCR（只初始化一次）
    ocr = PaddleOCR(use_angle_cls=True, lang="ch", show_log=False)
except Exception as e:
    ocr_available = False
    st.warning(f"PaddleOCR 未安装或加载失败，仅支持可复制文本的PDF: {e}")

# ========================
# 发票信息提取函数
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

    # 1. 提取发票号码（通常为8位或更多数字）
    inv_num_match = re.search(r'发票号码[:：\s]*(\d{8,20})', text)
    if inv_num_match:
        result["发票号码"] = inv_num_match.group(1)

    # 2. 提取购买方名称（在"购买方"之后）
    buyer_match = re.search(r'购买方[:：\s]*([^\n\r]{1,50}?(?:公司|集团|中心|店|厂))', text)
    if buyer_match:
        result["购买方名称"] = buyer_match.group(1).strip()

    # 3. 提取项目名称（匹配商品行，简化处理）
    # 假设项目在“货物或应税劳务名称”之后
    items = []
    item_lines = re.findall(r'(?:货物或应税劳务名称|项目名称)[：:\s]*([^\n\r]{2,20})', text)
    if not item_lines:
        # 备用：找包含中文且长度适中的行（启发式）
        lines = [line.strip() for line in text.split('\n') if 2 <= len(line) <= 20 and re.search(r'[\u4e00-\u9fa5]', line)]
        # 过滤掉明显不是项目的行（如金额、日期等）
        items = [line for line in lines if not re.search(r'\d{4}年|\d+.\d+|小写|合计|税额', line)]
    else:
        items = item_lines
    result["项目名称"] = "，".join(items[:5])  # 最多取5个，避免过长

    # 4. 提取价税合计（小写）
    total_match = re.search(r'价税合计.*?[（  $ ]小写[） $  ]?[:：\s]*[¥￥]?([\d,]+\.?\d*)', text)
    if total_match:
        amount_str = total_match.group(1).replace(',', '')
        try:
            float(amount_str)  # 验证是否为数字
            result["价税合计"] = amount_str
        except ValueError:
            pass

    # 5. 提取发票日期（多种格式）
    date_patterns = [
        r'发票日期[:：\s]*(\d{4})[-/年](\d{1,2})[-/月](\d{1,2})日?',
        r'开票日期[:：\s]*(\d{4})[-/年](\d{1,2})[-/月](\d{1,2})日?',
        r'(\d{4})年(\d{1,2})月(\d{1,2})日',
        r'日期[:：\s]*(\d{4}-\d{1,2}-\d{1,2})'
    ]
    
    date_found = False
    for pattern in date_patterns:
        match = re.search(pattern, text)
        if match:
            groups = match.groups()
            if len(groups) == 3:
                year, month, day = groups
                try:
                    dt = datetime(int(year), int(month), int(day))
                    result["发票日期"] = dt.strftime('%Y-%m-%d')
                    date_found = True
                    break
                except ValueError:
                    continue
            elif len(groups) == 1:
                # YYYY-MM-DD 格式
                try:
                    dt = datetime.strptime(groups[0], '%Y-%m-%d')
                    result["发票日期"] = dt.strftime('%Y-%m-%d')
                    date_found = True
                    break
                except ValueError:
                    continue

    return result

# ========================
# PDF 转文本（支持扫描件）
# ========================
def pdf_to_text(pdf_file):
    """将PDF转换为文本，优先尝试直接提取，失败则用OCR"""
    text = ""
    
    # 方法1：尝试直接提取文本
    try:
        with pdfplumber.open(pdf_file) as pdf:
            for page in pdf.pages:
                page_text = page.extract_text()
                if page_text:
                    text += page_text + "\n"
    except Exception as e:
        st.warning(f"直接提取文本失败: {e}")
        text = ""

    # 如果没提取到文本，且OCR可用，则用OCR
    if not text.strip() and ocr_available:
        st.info("检测到可能是扫描件，正在使用OCR识别...")
        try:
            from pdf2image import convert_from_bytes
            images = convert_from_bytes(pdf_file.read(), dpi=200)
            pdf_file.seek(0)  # 重置文件指针
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
                # 读取PDF内容
                text = pdf_to_text(file)
                if not text.strip():
                    st.warning(f"{file.name} 未提取到任何文字，请检查是否为有效发票。")
                    continue
                
                # 提取信息
                info = extract_invoice_info(text)
                info["文件名"] = file.name
                all_results.append(info)
                
            except Exception as e:
                st.error(f"处理 {file.name} 时出错: {e}")
    
    if all_results:
        df = pd.DataFrame(all_results)
        st.subheader("📋 提取结果预览")
        st.dataframe(df.fillna(""), use_container_width=True)
        
        # 生成Excel下载
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
