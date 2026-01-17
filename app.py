#!/usr/bin/env python3
# -*- coding: utf-8 -*-

"""
增值税发票识别应用
功能：上传PDF发票，自动识别并提取关键信息，生成Excel表格
"""

import re
import io
import tempfile
import os
from datetime import datetime
# from PIL import Image  # 暂时未使用，保留注释以便将来扩展

import streamlit as st
import pdfplumber
import pdf2image
import pandas as pd
from paddleocr import PaddleOCR

# 设置页面配置
st.set_page_config(
    page_title="增值税发票识别系统",
    page_icon="📊",
    layout="wide",
    initial_sidebar_state="collapsed"
)

# 初始化PaddleOCR（只在需要时初始化，节省资源）
@st.cache_resource
def init_ocr():
    """初始化PaddleOCR模型"""
    return PaddleOCR(use_angle_cls=True, lang='ch')

# 判断PDF是否为图片型
def is_image_based_pdf(pdf_path):
    """
    判断PDF是否为图片型（扫描件）
    
    参数:
        pdf_path: PDF文件路径
        
    返回:
        bool: True为图片型PDF，False为文本型PDF
    """
    try:
        with pdfplumber.open(pdf_path) as pdf:
            # 如果PDF页数为0，返回False
            if len(pdf.pages) == 0:
                return False
                
            # 尝试提取第一页的文本
            first_page = pdf.pages[0]
            text = first_page.extract_text()
            
            # 如果提取的文本为空或很少（小于10个字符），则认为是图片型PDF
            return not text or len(text.strip()) < 10
    except Exception as e:
        st.error(f"判断PDF类型时出错: {e}")
        return False

# 从图片型PDF中提取文字
def extract_text_from_image_pdf(pdf_path):
    """
    从图片型PDF中提取文字
    
    参数:
        pdf_path: PDF文件路径
        
    返回:
        str: 提取的文字
    """
    try:
        # 将PDF转换为图片
        images = pdf2image.convert_from_path(pdf_path)
        
        # 初始化OCR
        ocr = init_ocr()
        
        # 存储所有页面的文本
        all_text = []
        
        # 对每页图片进行OCR识别
        for img in images:
            # 转换为OCR需要的格式
            img_byte_arr = io.BytesIO()
            img.save(img_byte_arr, format='PNG')
            img_byte_arr = img_byte_arr.getvalue()
            
            # 进行OCR识别
            result = ocr.ocr(img_byte_arr, cls=True)
            
            # 提取文本
            page_text = []
            if result and len(result) > 0:
                for line in result[0]:
                    page_text.append(line[1][0])  # line[1][0]是识别出的文本
            
            all_text.extend(page_text)
        
        return "\n".join(all_text)
    except Exception as e:
        st.error(f"OCR识别时出错: {e}")
        return ""

# 从文本型PDF中提取文字
def extract_text_from_text_pdf(pdf_path):
    """
    从文本型PDF中提取文字
    
    参数:
        pdf_path: PDF文件路径
        
    返回:
        str: 提取的文字
    """
    try:
        text = []
        with pdfplumber.open(pdf_path) as pdf:
            for page in pdf.pages:
                page_text = page.extract_text()
                if page_text:
                    text.append(page_text)
        return "\n".join(text)
    except Exception as e:
        st.error(f"提取PDF文本时出错: {e}")
        return ""

# 从发票文本中提取关键信息
def extract_invoice_info(text):
    """
    从发票文本中提取关键信息
    
    参数:
        text: 发票文本
        
    返回:
        dict: 包含发票信息的字典
    """
    result = {
        "发票号码": "",
        "发票日期": "",
        "购买方名称": "",
        "商品项目": "",
        "价税合计": ""
    }
    
    # 提取发票号码（通常格式为数字）
    invoice_no_pattern = re.compile(r'发票号码[:：\s]*([A-Z0-9]+)', re.IGNORECASE)
    invoice_no_match = invoice_no_pattern.search(text)
    if invoice_no_match:
        result["发票号码"] = invoice_no_match.group(1)
    
    # 提取发票日期（格式：YYYY-MM-DD或YYYY年MM月DD日）
    date_pattern1 = re.compile(r'发票日期[:：\s]*(\d{4}[-/]?\d{1,2}[-/]?\d{1,2})')
    date_pattern2 = re.compile(r'(\d{4})年(\d{1,2})月(\d{1,2})日')
    date_pattern3 = re.compile(r'开票日期[:：\s]*(\d{4}[-/]?\d{1,2}[-/]?\d{1,2})')
    
    # 尝试多种日期格式
    date_match = date_pattern1.search(text) or date_pattern3.search(text)
    if date_match:
        # 标准化日期格式为YYYY-MM-DD
        date_str = date_match.group(1).replace('/', '-')
        try:
            # 尝试不同的日期格式
            for fmt in ['%Y-%m-%d', '%Y-%m-%d', '%Y%m%d']:
                try:
                    date_obj = datetime.strptime(date_str, fmt)
                    result["发票日期"] = date_obj.strftime('%Y-%m-%d')
                    break
                except ValueError:
                    continue
        except Exception:
            pass
    else:
        date_match = date_pattern2.search(text)
        if date_match:
            year, month, day = date_match.groups()
            try:
                date_obj = datetime(int(year), int(month), int(day))
                result["发票日期"] = date_obj.strftime('%Y-%m-%d')
            except ValueError:
                pass
    
    # 提取购买方名称
    buyer_pattern = re.compile(r'购买方[:：\s]*([^\n]+)')
    buyer_match = buyer_pattern.search(text)
    if buyer_match:
        result["购买方名称"] = buyer_match.group(1).strip()
    
    # 提取商品项目（合并为一行，逗号分隔）
    items = []
    
    # 尝试多种商品项目的正则表达式模式
    item_patterns = [
        re.compile(r'货物或应税劳务、服务名称[:：\s]*([^\n]+)'),
        re.compile(r'货物或应税劳务名称[:：\s]*([^\n]+)'),
        re.compile(r'项目名称[:：\s]*([^\n]+)'),
        re.compile(r'商品名称[:：\s]*([^\n]+)')
    ]
    
    # 首先尝试直接匹配商品项目
    for pattern in item_patterns:
        match = pattern.search(text)
        if match:
            item = match.group(1).strip()
            if item and item not in items:
                items.append(item)
    
    # 如果没有找到，尝试行扫描方法
    if not items:
        lines = text.split('\n')
        capture_items = False
        
        for line in lines:
            line = line.strip()
            if not line:
                continue
                
            # 检查是否是商品项目行的开始
            keywords = ['货物或应税劳务、服务名称', '货物或应税劳务名称', 
                        '项目名称', '商品名称']
            if any(keyword in line for keyword in keywords):
                capture_items = True
                # 提取冒号后的内容
                if ':' in line or '：' in line:
                    sep = ':' if ':' in line else '：'
                    if len(line.split(sep)) > 1:
                        item = line.split(sep, 1)[1].strip()
                        if item and not any(keyword in item for keyword in ['规格', '型号', '单位', '数量']):
                            items.append(item)
            elif capture_items:
                # 检查是否应该停止捕获
                stop_keywords = ['价税合计', '合计', '小写', '大写', '备注']
                if any(keyword in line for keyword in stop_keywords):
                    break
                
                # 检查是否是商品行（不是表头行）
                exclude_keywords = ['规格型号', '单位', '数量', '单价', '金额', 
                                   '税率', '税额', '序号', 'No']
                if not any(keyword in line for keyword in exclude_keywords):
                    # 过滤掉纯数字和特殊字符行
                    if line and not (line.isdigit() or all(c in '0123456789.¥￥,，' for c in line)):
                        items.append(line)
    
    # 如果找到了商品项目，合并为一行
    if items:
        # 去重并合并
        unique_items = []
        for item in items:
            if item not in unique_items:
                unique_items.append(item)
        result["商品项目"] = '，'.join(unique_items)
    else:
        result["商品项目"] = "未识别到商品项目"
    
    # 提取价税合计（小写金额）
    total_patterns = [
        re.compile(r'价税合计\(小写\)[:：\s]*[¥￥\s]*([\d.,]+)'),
        re.compile(r'价税合计[:：\s]*[¥￥\s]*([\d.,]+)'),
        re.compile(r'合计[:：\s]*[¥￥\s]*([\d.,]+)'),
        re.compile(r'Total[:：\s]*[¥￥\s]*([\d.,]+)')
    ]
    
    for pattern in total_patterns:
        total_match = pattern.search(text)
        if total_match:
            # 只保留数字和小数点
            total_amount = re.sub(r'[^\d.]', '', total_match.group(1))
            # 确保只有一个小数点
            if total_amount.count('.') > 1:
                parts = total_amount.split('.')
                total_amount = parts[0] + '.' + ''.join(parts[1:])
            result["价税合计"] = total_amount
            break
    
    # 如果还是没有找到，尝试查找所有包含金额格式的行
    if not result["价税合计"]:
        amount_pattern = re.compile(r'[¥￥\s]*(\d{1,3}(?:,\d{3})*(?:\.\d{2})?)')
        matches = amount_pattern.findall(text)
        if matches:
            # 取最大的金额作为价税合计（通常是最大的金额）
            amounts = []
            for match in matches:
                try:
                    # 移除千位分隔符
                    clean_amount = match.replace(',', '')
                    amounts.append(float(clean_amount))
                except ValueError:
                    continue
            
            if amounts:
                max_amount = max(amounts)
                result["价税合计"] = f"{max_amount:.2f}"
    
    return result

# 处理单个PDF文件
def process_pdf(pdf_file):
    """
    处理单个PDF文件，提取发票信息
    
    参数:
        pdf_file: UploadedFile对象
        
    返回:
        dict: 包含发票信息的字典
    """
    temp_path = None
    try:
        # 创建临时文件
        with tempfile.NamedTemporaryFile(delete=False, suffix='.pdf') as temp_file:
            temp_file.write(pdf_file.getvalue())
            temp_path = temp_file.name
        
        # 判断PDF类型并提取文本
        pdf_type = "图片型" if is_image_based_pdf(temp_path) else "文本型"
        st.info(f"文件 {pdf_file.name} 是{pdf_type}PDF，正在处理...")
        
        # 尝试两种方法提取文本，提高成功率
        text = ""
        try:
            # 首先尝试直接提取文本
            text = extract_text_from_text_pdf(temp_path)
        except ValueError as e:
            st.warning(f"直接提取文本失败，尝试OCR: {e}")
        
        # 如果文本提取失败或文本太少，尝试OCR
        if not text or len(text.strip()) < 50:
            try:
                st.info(f"使用OCR技术识别文件 {pdf_file.name}...")
                ocr_text = extract_text_from_image_pdf(temp_path)
                # 合并两种方法的结果
                if ocr_text:
                    text = text + "\n" + ocr_text if text else ocr_text
            except ValueError as e:
                st.error(f"OCR识别失败: {e}")
        
        if not text:
            raise ValueError("无法从PDF中提取文本")
        
        # 提取发票信息
        invoice_info = extract_invoice_info(text)
        
        # 添加文件名和处理状态
        invoice_info["文件名"] = pdf_file.name
        invoice_info["处理状态"] = "成功"
        
        # 验证关键信息是否提取成功
        missing_fields = []
        for field in ["发票号码", "发票日期", "购买方名称", "价税合计"]:
            if not invoice_info.get(field):
                missing_fields.append(field)
        
        if missing_fields:
            invoice_info["处理状态"] = f"部分信息缺失: {', '.join(missing_fields)}"
            st.warning(f"文件 {pdf_file.name} 部分信息无法识别: {', '.join(missing_fields)}")
        
        return invoice_info
        
    except Exception as e:
        error_msg = str(e)
        st.error(f"处理文件 {pdf_file.name} 时出错: {error_msg}")
        return {
            "文件名": pdf_file.name,
            "发票号码": "",
            "发票日期": "",
            "购买方名称": "",
            "商品项目": "",
            "价税合计": "",
            "处理状态": f"处理出错: {error_msg}"
        }
    finally:
        # 清理临时文件
        if temp_path and os.path.exists(temp_path):
            try:
                os.unlink(temp_path)
            except ValueError as e:
                st.warning(f"无法删除临时文件: {e}")

# 主应用界面
def main():
    """主应用函数"""
    # 页面标题
    st.title("📊 增值税发票识别系统")
    st.markdown("---")
    
    # 文件上传区
    st.subheader("上传发票")
    uploaded_files = st.file_uploader(
        "请选择增值税发票PDF文件（支持扫描件和文本PDF）",
        type=["pdf"],
        accept_multiple_files=True,
        help="支持多个PDF文件同时上传"
    )
    
    # 处理按钮
    if uploaded_files and st.button("开始处理", type="primary", use_container_width=True):
        # 创建进度条
        progress_bar = st.progress(0)
        status_text = st.empty()
        
        # 存储所有发票信息
        all_invoices = []
        
        # 处理每个上传的文件
        for i, pdf_file in enumerate(uploaded_files):
            # 更新进度
            progress = (i + 1) / len(uploaded_files)
            progress_bar.progress(progress)
            status_text.text(f"处理中: {pdf_file.name} ({i+1}/{len(uploaded_files)})")
            
            # 处理PDF文件
            invoice_info = process_pdf(pdf_file)
            all_invoices.append(invoice_info)
        
        # 完成处理
        progress_bar.progress(1.0)
        status_text.text("处理完成！")
        
        # 显示结果表格
        if all_invoices:
            st.markdown("---")
            st.subheader("识别结果")
            
            # 创建DataFrame
            df = pd.DataFrame(all_invoices)
            
            # 调整列顺序
            columns_order = ["文件名", "发票号码", "发票日期", "购买方名称", "商品项目", "价税合计", "处理状态"]
            # 确保所有列都存在
            for col in columns_order:
                if col not in df.columns:
                    df[col] = ""
            df = df[columns_order]
            
            # 显示表格
            st.dataframe(df, use_container_width=True)
            
            # 下载按钮
            st.markdown("---")
            st.subheader("导出数据")
            
            # 创建Excel文件
            output = io.BytesIO()
            with pd.ExcelWriter(output, engine='openpyxl') as writer:
                df.to_excel(writer, index=False, sheet_name='发票数据')
            
            # 提供下载链接
            st.download_button(
                label="下载Excel文件",
                data=output.getvalue(),
                file_name=f"发票数据_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True
            )
    
    # 说明信息
    st.markdown("---")
    st.subheader("使用说明")
    st.markdown("""
    1. 点击"浏览文件"按钮，选择一个或多个增值税发票PDF文件
    2. 点击"开始处理"按钮，系统将自动识别发票信息
    3. 处理完成后，系统会显示识别结果表格
    4. 点击"下载Excel文件"按钮，将识别结果导出为Excel文件
    
    **注意事项：**
    - 系统支持扫描型PDF（图片PDF）和文本型PDF
    - 对于扫描型PDF，系统会使用OCR技术进行文字识别
    - 识别准确率受发票质量影响，如有识别错误，请手动修正
    """)

# 运行应用
if __name__ == "__main__":
    main()
