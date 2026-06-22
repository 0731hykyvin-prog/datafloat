"""
电信话单 PDF 清洗模块
=====================
遍历文件夹下所有 PDF 文件，转换为 XLSX 并清洗：
1. PDF → XLSX（文件名以第二行业务号码命名）
2. 删除前三行 + 后两行
3. 删除"开始时间"为空的行
"""

import os
import re

import pandas as pd

try:
    import pdfplumber
except ImportError:
    pdfplumber = None


def _extract_phone(df, row_idx=1):
    """从 DataFrame 第 row_idx 行提取 11 位手机号。"""
    if df is None or df.empty or row_idx >= len(df):
        return None
    row_text = " ".join([str(v) for v in df.iloc[row_idx].values if pd.notna(v)])
    match = re.search(r"1[3-9]\d{9}", row_text)
    return match.group() if match else None


def _clean_table(df):
    """清洗表格：去前三行 + 后两行 + 删除开始时间为空的行。"""
    if df is None or len(df) <= 5:
        return df
    # 去掉前三行和后两行
    df = df.iloc[3:].iloc[:-2] if len(df) > 5 else df.iloc[3:]
    # 删除"开始时间"为空的行
    for col in df.columns:
        if "开始时间" in str(col) or "通话起始时间" in str(col):
            df = df[df[col].notna() & (df[col].astype(str).str.strip() != "")]
            break
    return df


def convert_pdf_to_xlsx(pdf_path, output_dir=None):
    """
    单个 PDF → XLSX 转换 + 清洗。
    返回: (output_path, phone) 或 (None, None)
    """
    if pdfplumber is None:
        raise ImportError("请安装 pdfplumber: pip install pdfplumber")

    out_dir = output_dir or os.path.dirname(pdf_path) or "."
    os.makedirs(out_dir, exist_ok=True)

    with pdfplumber.open(pdf_path) as pdf:
        if len(pdf.pages) == 0:
            return None, None

        # 提取所有页面的表格
        all_rows = []
        for page in pdf.pages:
            tables = page.extract_tables()
            for tbl in tables:
                for row in tbl:
                    all_rows.append(row)

        if not all_rows:
            return None, None

        # 转为 DataFrame
        df = pd.DataFrame(all_rows)
        # 第一行作为表头
        if len(df) > 0:
            df.columns = df.iloc[0]
            df = df.iloc[1:].reset_index(drop=True)

        # 提取号码 — 从第2行（索引1）提取11位号码
        phone = _extract_phone(df)
        if not phone:
            phone = _extract_phone(df, row_idx=0)

        # 清洗
        df = _clean_table(df)

        # 输出 — 同名文件加序号防止覆盖
        base = phone if phone else os.path.splitext(os.path.basename(pdf_path))[0]
        filename = f"{base}.xlsx"
        output_path = os.path.join(out_dir, filename)
        counter = 2
        while os.path.exists(output_path):
            filename = f"{base}_{counter}.xlsx"
            output_path = os.path.join(out_dir, filename)
            counter += 1

        df.to_excel(output_path, index=False)

        return output_path, phone


def batch_convert_folder(folder_path, output_dir=None):
    """
    批量转换文件夹下所有 PDF。
    返回: (results, logs)
        results: [(pdf_path, output_path, phone), ...]
        logs: 日志字符串列表
    """
    out_dir = output_dir or os.path.join(folder_path, "xlsx输出")
    os.makedirs(out_dir, exist_ok=True)

    results = []
    logs = []
    pdf_files = [f for f in os.listdir(folder_path) if f.lower().endswith(".pdf")]

    logs.append(f"扫描到 {len(pdf_files)} 个 PDF 文件")
    logs.append(f"输出目录: {out_dir}")
    for filename in pdf_files:
        pdf_path = os.path.join(folder_path, filename)
        try:
            output_path, phone = convert_pdf_to_xlsx(pdf_path, out_dir)
            if output_path:
                logs.append(f"✅ {filename} → {os.path.basename(output_path)} (号码: {phone})")
                results.append((pdf_path, output_path, phone))
            else:
                logs.append(f"❌ {filename}: 无法提取表格")
        except Exception as e:
            logs.append(f"❌ {filename}: {e}")

    logs.append(f"完成: {len(results)}/{len(pdf_files)} 成功, 输出目录: {out_dir}")
    return results, logs
