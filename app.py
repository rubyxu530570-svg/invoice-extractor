import streamlit as st
import pandas as pd
import re
from datetime import datetime
from io import BytesIO

# 尝试导入依赖库
try:
    import pdfplumber
except ImportError:
    st.error("缺少 pdfplumber，请确保 requirements.txt 中包含它。")
    st.stop()

try:
    from paddleocr import PaddleOCR
    ocr_available = True
    ocr = PaddleOCR(use_angle_cls=True, lang="ch", show_log=False)
except Exception:
    ocr_available = False
    st.warning("PaddleOCR 未加载，仅支持可复制文本的PDF。")

def extract_invoice_info(text):
    result = {
        "发票号码": "",
        "发票日期": "",
        "购买方名称": "",
        "项目名称": "",
        "价税合计": ""
    }

    # 1. 发票号码（18位数字）
    inv_match = re.search(r'发票号码[:：\s]*(\d{18})', text)
    if inv_match:
        result["发票号码"] = inv_match.group(1)

    # 2. 开票日期
    date_match = re.search(r'开票日期[:：\s]*(\d{4}年\d{1,2}月\d{1,2}日)', text)
    if date_match:
        d = date_match.group(1)
        d_clean = d.replace('年', '-').replace('月', '-').replace('日', '')
        try:
            dt = datetime.strptime(d_clean, '%Y-%m-%d')
            result["发票日期"] = dt.strftime('%Y-%m-%d')
        except:
            pass

    # 3. 购买方名称（匹配“名称:”后的内容）
    buyer_match = re.search(r'名称[:：]\s*([^\n\r]*?公司)', text)
    if buyer_match:
        name = buyer_match.group(1).strip()
        # 清理非中文/字母/数字字符
        clean_name = re.sub(r'[^\u4e00-\u9fa5a-zA-Z0-9]', '', name)
        result["购买方名称"] = clean_name

    # 4. 项目名称（找 * 开头的行）
    lines = [line.strip() for line in text.split('\n') if line.strip()]
    project_lines = []
    for line in lines:
        if line.startswith('*') and len(line) > 2:
            project_lines.append(line)
    if project_lines:
        result["项目名称"] = "，".join(project_lines[:2])  # 取前1～2个

    # 5. 价税合计（小写）——重点修复
    # 尝试多种格式
    amount = ""
    # 格式1: (小写) ¥361.00
    m1 = re.search(r'[  $ （]小写[ $  ）][\s:：]*[¥￥]?\s*(\d{1,3}(?:,\d{3})*(?:\.\d+)?)', text)
    # 格式2: 价税合计（小写）¥361.00
    m2 = re.search(r'(?:价税合计|合计).*?[  $ （]小写[ $  ）].*?[¥￥]?\s*(\d{1,3}(?:,\d{3})*(?:\.\d+)?)', text)
    # 格式3: 直接找 ¥ 后的金额（兜底）
    m3 = re.search(r'[¥￥]\s*(\d{1,3}(?:,\d{3})*(?:\.\d+)?)', text)

    for m in [m1, m2, m3]:
        if m:
            amount_str = m.group(1).replace(',', '')
            try:
                float(amount_str)
                amount = amount_str
                break
            except:
                continue

    result["价税合计"] = amount
    return result

def pdf_to_text(pdf_file):
    text = ""
    # 先尝试直接提取
    try:
        with pdfplumber.open(pdf_file) as pdf:
            for page in pdf.pages:
                t = page.extract_text()
                if t:
                    text += t + "\n"
    except:
        text = ""

    # 如果没内容且OCR可用，则用OCR
    if not text.strip() and ocr_available:
        try:
            from pdf2image import convert_from_bytes
            images = convert_from_bytes(pdf_file.read(), dpi=200)
            pdf_file.seek(0)
            for img in images:
                ocr_result = ocr.ocr(img, cls=True)
                if ocr_result and ocr_result[0]:
                    for line in ocr_result[0]:
                        text += line[1][0] + "\n"
        except Exception as e:
            st.error(f"OCR失败: {e}")
    return text

# ===== 网页界面 =====
st.set_page_config(page_title="发票信息提取工具", layout="wide")
st.title("📊 发票信息自动提取工具")
st.markdown("上传多个增值税发票 PDF 文件，自动识别并生成 Excel 表格")

uploaded_files = st.file_uploader(
    "📁 请选择一个或多个发票 PDF 文件",
    type=["pdf"],
    accept_multiple_files=True
)

if uploaded_files:
    all_results = []
    for file in uploaded_files:
        with st.spinner(f"处理中: {file.name}"):
            try:
                text = pdf_to_text(file)
                if not text.strip():
                    st.warning(f"{file.name} 未提取到文字")
                    continue
                info = extract_invoice_info(text)
                info["文件名"] = file.name
                all_results.append(info)
            except Exception as e:
                st.error(f"处理 {file.name} 出错: {e}")

    if all_results:
        df = pd.DataFrame(all_results)
        st.subheader("📋 提取结果")
        st.dataframe(df.fillna(""), use_container_width=True)

        # 生成Excel
        output = BytesIO()
        with pd.ExcelWriter(output, engine='openpyxl') as writer:
            df.to_excel(writer, index=False, sheet_name='发票信息')
        excel_data = output.getvalue()

        st.download_button(
            label="📥 下载Excel",
            data=excel_data,
            file_name="发票信息汇总.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
else:
    st.info("请上传PDF文件开始处理。")
