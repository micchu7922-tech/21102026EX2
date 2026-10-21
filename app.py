# ============================================================
# Streamlit 資料清洗與品質分析工具
# 功能：上傳 Excel → 自動判斷完整/不完整資料 → 多 Sheet 報表 + 圖表 → 下載
# ============================================================

import streamlit as st
import pandas as pd
import numpy as np
import re
from datetime import datetime
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils.dataframe import dataframe_to_rows
from openpyxl.utils import get_column_letter
from openpyxl.chart import PieChart, BarChart, Reference
from openpyxl.chart.label import DataLabelList
import io
import plotly.express as px
import plotly.graph_objects as go

# -------------------- 頁面設定 --------------------
st.set_page_config(
    page_title="資料品質清洗工具",
    page_icon="📊",
    layout="wide",
    initial_sidebar_state="expanded"
)

st.title("📊 組織資料品質清洗與分析工具")
st.markdown("上傳 Excel 檔案後，系統會自動判斷**完整資料**與**不完整資料**，並產生多工作表報表與圖表。")

# -------------------- 資料品質檢查函數 --------------------
def is_null_or_empty(value):
    if pd.isna(value):
        return True
    if isinstance(value, str):
        cleaned = value.strip().upper()
        null_keywords = {
            '', 'NULL', 'N/A', 'NA', 'NONE', 'NIL', 'UNKNOWN',
            '#NA', '#N/A', 'NAN', '-', '--', 'NONE_PROVIDED'
        }
        return cleaned in null_keywords
    return False


def is_valid_date(value):
    if is_null_or_empty(value):
        return False
    value_str = str(value).strip()
    if '00-00' in value_str or '0000' in value_str:
        return False
    if re.search(r'02-31|02/31|02\.31', value_str):
        return False
    formats = [
        '%Y-%m-%d', '%Y/%m/%d', '%d/%m/%Y', '%m/%d/%Y',
        '%Y.%m.%d', '%d-%m-%Y', '%Y-%m-%d %H:%M:%S'
    ]
    for fmt in formats:
        try:
            datetime.strptime(value_str[:10], fmt)
            return True
        except:
            continue
    return False


def is_valid_phone(value):
    if is_null_or_empty(value):
        return False
    value_str = str(value).strip()
    cleaned = re.sub(r'[\s\-\(\)\+]', '', value_str)
    if not cleaned.isdigit():
        return False
    if len(cleaned) < 7 or len(cleaned) > 15:
        return False
    return True


def is_valid_email(value):
    if is_null_or_empty(value):
        return False
    value_str = str(value).strip()
    if '@@' in value_str or '..' in value_str:
        return False
    pattern = r'^[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,}$'
    return bool(re.match(pattern, value_str))


def is_valid_ip(value):
    if is_null_or_empty(value):
        return False
    value_str = str(value).strip().lower()
    if value_str in ['localhost', '999.999.999.999']:
        return False
    parts = value_str.split('.')
    if len(parts) != 4:
        return False
    try:
        return all(0 <= int(p) <= 255 for p in parts)
    except:
        return False


def is_valid_age(value):
    if is_null_or_empty(value):
        return False
    try:
        age = float(value)
        return 1 <= age <= 120
    except:
        return False


def is_valid_salary(value):
    if is_null_or_empty(value):
        return False
    try:
        salary = float(str(value).replace(',', '').replace('$', '').strip())
        return salary > 0
    except:
        return False


def is_valid_credit_score(value):
    if is_null_or_empty(value):
        return False
    try:
        score = float(value)
        return 300 <= score <= 850
    except:
        return False


def check_row_quality(row):
    issues = []

    critical_cols = [
        'First_Name', 'Last_Name', 'Email_Address', 'Phone_Number',
        'Department', 'Job_Title', 'Employment_Status', 'Salary_Monthly',
        'User_ID', 'DOB', 'Age'
    ]
    for col in critical_cols:
        if col in row.index and is_null_or_empty(row[col]):
            issues.append(f"{col} 為空值/NULL/N/A")

    date_cols = ['DOB', 'Hire_Date', 'Contract_End_Date', 'Registration_Date', 'Last_Purchase_Date']
    for col in date_cols:
        if col in row.index and not is_null_or_empty(row[col]):
            if not is_valid_date(row[col]):
                issues.append(f"{col} 日期格式異常 ({row[col]})")

    phone_cols = ['Phone_Number', 'Emergency_Contact_Phone']
    for col in phone_cols:
        if col in row.index and not is_null_or_empty(row[col]):
            if not is_valid_phone(row[col]):
                issues.append(f"{col} 電話格式異常 ({row[col]})")

    if 'Email_Address' in row.index and not is_null_or_empty(row['Email_Address']):
        if not is_valid_email(row['Email_Address']):
            issues.append(f"Email 格式異常 ({row['Email_Address']})")

    if 'IP_Address' in row.index and not is_null_or_empty(row['IP_Address']):
        if not is_valid_ip(row['IP_Address']):
            issues.append(f"IP 異常 ({row['IP_Address']})")

    if 'Age' in row.index and not is_null_or_empty(row['Age']):
        if not is_valid_age(row['Age']):
            issues.append(f"年齡異常 ({row['Age']})")

    if 'Salary_Monthly' in row.index and not is_null_or_empty(row['Salary_Monthly']):
        if not is_valid_salary(row['Salary_Monthly']):
            issues.append(f"薪資異常 ({row['Salary_Monthly']})")

    if 'Credit_Score' in row.index and not is_null_or_empty(row['Credit_Score']):
        if not is_valid_credit_score(row['Credit_Score']):
            issues.append(f"信用評分異常 ({row['Credit_Score']})")

    if 'Employment_Status' in row.index and not is_null_or_empty(row['Employment_Status']):
        status = str(row['Employment_Status']).strip()
        standard_status = {'Active', 'Pending', 'Terminated', 'Inactive', 'Probation'}
        if status not in standard_status and status.lower() not in {s.lower() for s in standard_status}:
            issues.append(f"就業狀態非標準值 ({status})")
        elif status not in standard_status:
            issues.append(f"就業狀態大小寫不一致 ({status})")

    if 'Department' in row.index and not is_null_or_empty(row['Department']):
        dept = str(row['Department']).strip()
        if dept in ['Fin.', 'hr', 'SALES', 'IT ', ' Information Tech']:
            issues.append(f"部門名稱不規範 ({dept})")

    if 'Two_Factor_Auth' in row.index and not is_null_or_empty(row['Two_Factor_Auth']):
        val = str(row['Two_Factor_Auth']).strip().lower()
        valid_bool = {'enabled', 'disabled', 'true', 'false', 'yes', 'no', 'y', 'n', '1', '0'}
        if val not in valid_bool:
            issues.append(f"Two_Factor_Auth 值異常 ({row['Two_Factor_Auth']})")

    if 'Years_of_Experience' in row.index and not is_null_or_empty(row['Years_of_Experience']):
        try:
            exp = float(str(row['Years_of_Experience']).replace('yrs', '').strip())
            if exp < 0 or exp > 50:
                issues.append(f"年資異常 ({row['Years_of_Experience']})")
        except:
            issues.append(f"年資格式異常 ({row['Years_of_Experience']})")

    if 'Customer_Satisfaction_Score' in row.index and not is_null_or_empty(row['Customer_Satisfaction_Score']):
        try:
            score = float(row['Customer_Satisfaction_Score'])
            if not (1 <= score <= 5):
                issues.append(f"滿意度超出範圍 ({row['Customer_Satisfaction_Score']})")
        except:
            issues.append(f"滿意度格式異常 ({row['Customer_Satisfaction_Score']})")

    is_complete = len(issues) == 0
    return is_complete, issues


# -------------------- 產生 Excel 報表 --------------------
def create_excel_report(complete_df, incomplete_df, issue_sorted, null_stats, numeric_compare,
                        complete_dept, incomplete_dept, complete_status, incomplete_status,
                        complete_count, incomplete_count, total_rows, complete_pct, incomplete_pct):
    wb = Workbook()

    header_font = Font(bold=True, color="FFFFFF", name="Arial", size=11)
    header_fill_complete = PatternFill(start_color="2E7D32", end_color="2E7D32", fill_type="solid")
    header_fill_incomplete = PatternFill(start_color="C62828", end_color="C62828", fill_type="solid")
    header_fill_stats = PatternFill(start_color="1565C0", end_color="1565C0", fill_type="solid")
    header_fill_chart = PatternFill(start_color="6A1B9A", end_color="6A1B9A", fill_type="solid")
    header_alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
    thin_border = Border(
        left=Side(style='thin'), right=Side(style='thin'),
        top=Side(style='thin'), bottom=Side(style='thin')
    )
    cell_font = Font(name="Arial", size=10)
    title_font = Font(name="Arial", size=14, bold=True, color="1565C0")
    section_font = Font(name="Arial", size=12, bold=True, color="333333")
    number_font = Font(name="Arial", size=10)

    def write_dataframe_sheet(ws, dataframe, header_fill):
        for r_idx, row in enumerate(dataframe_to_rows(dataframe, index=False, header=True), 1):
            for c_idx, value in enumerate(row, 1):
                cell = ws.cell(row=r_idx, column=c_idx, value=value)
                cell.font = cell_font
                cell.border = thin_border
                cell.alignment = Alignment(vertical="center", wrap_text=True)
                if r_idx == 1:
                    cell.font = header_font
                    cell.fill = header_fill
                    cell.alignment = header_alignment
        for col_idx, column in enumerate(dataframe.columns, 1):
            max_length = max(
                len(str(column)),
                dataframe[column].astype(str).str.len().max() if len(dataframe) > 0 else 0
            )
            adjusted_width = min(max_length + 2, 45)
            ws.column_dimensions[get_column_letter(col_idx)].width = adjusted_width
        ws.freeze_panes = "A2"
        ws.row_dimensions[1].height = 30

    # Sheet 1
    ws1 = wb.active
    ws1.title = "Complete_Data"
    write_dataframe_sheet(ws1, complete_df, header_fill_complete)

    # Sheet 2
    ws2 = wb.create_sheet("Incomplete_Data")
    write_dataframe_sheet(ws2, incomplete_df, header_fill_incomplete)

    # Sheet 3: Statistics
    ws3 = wb.create_sheet("Statistics")
    ws3['A1'] = "資料品質統計報表"
    ws3['A1'].font = title_font
    ws3.merge_cells('A1:F1')
    ws3.row_dimensions[1].height = 30
    ws3['A2'] = f"產生時間：{datetime.now().strftime('%Y-%m-%d %H:%M:%S')}"
    ws3['A2'].font = Font(name="Arial", size=9, italic=True, color="666666")

    ws3['A4'] = "一、資料完整性總覽"
    ws3['A4'].font = section_font
    for col, h in enumerate(["項目", "數值", "備註"], 1):
        cell = ws3.cell(row=5, column=col, value=h)
        cell.font = header_font
        cell.fill = header_fill_stats
        cell.alignment = header_alignment
        cell.border = thin_border

    summary_data = [
        ["原始總筆數", total_rows, ""],
        ["完整資料筆數", complete_count, f"{complete_pct}%"],
        ["不完整資料筆數", incomplete_count, f"{incomplete_pct}%"],
        ["完整資料比例", f"{complete_pct}%", ""],
        ["不完整資料比例", f"{incomplete_pct}%", ""],
    ]
    for r_idx, row_data in enumerate(summary_data, 6):
        for c_idx, val in enumerate(row_data, 1):
            cell = ws3.cell(row=r_idx, column=c_idx, value=val)
            cell.font = number_font
            cell.border = thin_border

    start_row = 13
    ws3[f'A{start_row}'] = "二、不完整資料「問題類型」出現次數排行"
    ws3[f'A{start_row}'].font = section_font
    for col, h in enumerate(["問題類型", "出現次數", "佔不完整資料比例"], 1):
        cell = ws3.cell(row=start_row+1, column=col, value=h)
        cell.font = header_font
        cell.fill = header_fill_incomplete
        cell.alignment = header_alignment
        cell.border = thin_border

    if issue_sorted:
        for r_idx, (issue_type, cnt) in enumerate(issue_sorted, start_row+2):
            pct = round(cnt / incomplete_count * 100, 1) if incomplete_count > 0 else 0
            ws3.cell(row=r_idx, column=1, value=issue_type).border = thin_border
            ws3.cell(row=r_idx, column=2, value=cnt).border = thin_border
            ws3.cell(row=r_idx, column=3, value=f"{pct}%").border = thin_border
            for c in range(1, 4):
                ws3.cell(row=r_idx, column=c).font = number_font

    start_row3 = start_row + 2 + max(len(issue_sorted), 1) + 3
    ws3[f'A{start_row3}'] = "三、關鍵欄位空值 / 異常值統計（全體資料）"
    ws3[f'A{start_row3}'].font = section_font
    for col, h in enumerate(["欄位名稱", "空值/異常筆數", "佔總筆數比例(%)"], 1):
        cell = ws3.cell(row=start_row3+1, column=col, value=h)
        cell.font = header_font
        cell.fill = header_fill_stats
        cell.alignment = header_alignment
        cell.border = thin_border

    for r_idx, row_data in enumerate(null_stats, start_row3+2):
        for c_idx, val in enumerate(row_data, 1):
            cell = ws3.cell(row=r_idx, column=c_idx, value=val)
            cell.font = number_font
            cell.border = thin_border

    start_row4 = start_row3 + 2 + len(null_stats) + 3
    ws3[f'A{start_row4}'] = "四、數值欄位比較（完整資料 vs 不完整資料）"
    ws3[f'A{start_row4}'].font = section_font
    headers4 = [
        "欄位", "完整_有效筆數", "完整_平均", "完整_最小", "完整_最大",
        "不完整_有效筆數", "不完整_平均", "不完整_最小", "不完整_最大"
    ]
    for col, h in enumerate(headers4, 1):
        cell = ws3.cell(row=start_row4+1, column=col, value=h)
        cell.font = header_font
        cell.fill = header_fill_stats
        cell.alignment = header_alignment
        cell.border = thin_border

    for r_idx, item in enumerate(numeric_compare, start_row4+2):
        values = [
            item['欄位'],
            item['完整_有效筆數'], item['完整_平均'], item['完整_最小'], item['完整_最大'],
            item['不完整_有效筆數'], item['不完整_平均'], item['不完整_最小'], item['不完整_最大']
        ]
        for c_idx, val in enumerate(values, 1):
            cell = ws3.cell(row=r_idx, column=c_idx, value=val if val is not None else "-")
            cell.font = number_font
            cell.border = thin_border
            cell.alignment = Alignment(horizontal="center", vertical="center")

    for col in range(1, 10):
        ws3.column_dimensions[get_column_letter(col)].width = 18
    ws3.column_dimensions['A'].width = 45

    # Sheet 4: Charts
    ws4 = wb.create_sheet("Charts")
    ws4['A1'] = "資料品質圖表統計"
    ws4['A1'].font = Font(name="Arial", size=16, bold=True, color="6A1B9A")
    ws4.merge_cells('A1:L1')
    ws4.row_dimensions[1].height = 32
    ws4['A2'] = f"產生時間：{datetime.now().strftime('%Y-%m-%d %H:%M:%S')}"
    ws4['A2'].font = Font(name="Arial", size=9, italic=True, color="666666")

    # 圖1 圓餅
    ws4['A4'] = "【圖1】完整資料 vs 不完整資料比例"
    ws4['A4'].font = section_font
    ws4['A5'] = "類別"
    ws4['B5'] = "筆數"
    ws4['A5'].font = header_font
    ws4['A5'].fill = header_fill_chart
    ws4['B5'].font = header_font
    ws4['B5'].fill = header_fill_chart
    ws4['A5'].border = thin_border
    ws4['B5'].border = thin_border
    ws4['A6'] = "完整資料"
    ws4['B6'] = complete_count
    ws4['A7'] = "不完整資料"
    ws4['B7'] = incomplete_count
    for r in range(6, 8):
        for c in range(1, 3):
            ws4.cell(row=r, column=c).border = thin_border
            ws4.cell(row=r, column=c).font = number_font

    pie1 = PieChart()
    pie1.title = "完整 vs 不完整資料比例"
    labels1 = Reference(ws4, min_col=1, min_row=6, max_row=7)
    data1 = Reference(ws4, min_col=2, min_row=5, max_row=7)
    pie1.add_data(data1, titles_from_data=True)
    pie1.set_categories(labels1)
    pie1.dataLabels = DataLabelList()
    pie1.dataLabels.showPercent = True
    pie1.dataLabels.showVal = True
    pie1.width = 14
    pie1.height = 10
    ws4.add_chart(pie1, "D4")

    # 圖2 問題類型
    ws4['A20'] = "【圖2】不完整資料問題類型 Top 10"
    ws4['A20'].font = section_font
    ws4['A21'] = "問題類型"
    ws4['B21'] = "出現次數"
    ws4['A21'].font = header_font
    ws4['A21'].fill = header_fill_incomplete
    ws4['B21'].font = header_font
    ws4['B21'].fill = header_fill_incomplete
    ws4['A21'].border = thin_border
    ws4['B21'].border = thin_border

    top_issues = issue_sorted[:10] if issue_sorted else []
    for i, (issue_type, cnt) in enumerate(top_issues):
        ws4.cell(row=22+i, column=1, value=issue_type[:40]).border = thin_border
        ws4.cell(row=22+i, column=2, value=cnt).border = thin_border
        ws4.cell(row=22+i, column=1).font = number_font
        ws4.cell(row=22+i, column=2).font = number_font

    if top_issues:
        bar1 = BarChart()
        bar1.type = "col"
        bar1.title = "問題類型出現次數 Top 10"
        bar1.y_axis.title = "出現次數"
        data2 = Reference(ws4, min_col=2, min_row=21, max_row=21+len(top_issues))
        cats2 = Reference(ws4, min_col=1, min_row=22, max_row=21+len(top_issues))
        bar1.add_data(data2, titles_from_data=True)
        bar1.set_categories(cats2)
        bar1.width = 18
        bar1.height = 10
        ws4.add_chart(bar1, "D20")

    # 圖3 空值比例
    ws4['A40'] = "【圖3】關鍵欄位空值 / 異常比例"
    ws4['A40'].font = section_font
    ws4['A41'] = "欄位"
    ws4['B41'] = "空值比例(%)"
    ws4['A41'].font = header_font
    ws4['A41'].fill = header_fill_stats
    ws4['B41'].font = header_font
    ws4['B41'].fill = header_fill_stats
    ws4['A41'].border = thin_border
    ws4['B41'].border = thin_border

    null_for_chart = sorted(
        [(name, pct) for name, cnt, pct in null_stats if pct > 0],
        key=lambda x: x[1], reverse=True
    )[:12]
    for i, (name, pct) in enumerate(null_for_chart):
        ws4.cell(row=42+i, column=1, value=name).border = thin_border
        ws4.cell(row=42+i, column=2, value=pct).border = thin_border
        ws4.cell(row=42+i, column=1).font = number_font
        ws4.cell(row=42+i, column=2).font = number_font

    if null_for_chart:
        bar2 = BarChart()
        bar2.type = "bar"
        bar2.title = "關鍵欄位空值/異常比例 (%)"
        data3 = Reference(ws4, min_col=2, min_row=41, max_row=41+len(null_for_chart))
        cats3 = Reference(ws4, min_col=1, min_row=42, max_row=41+len(null_for_chart))
        bar2.add_data(data3, titles_from_data=True)
        bar2.set_categories(cats3)
        bar2.width = 16
        bar2.height = 12
        ws4.add_chart(bar2, "D40")

    # 圖4 部門
    ws4['A60'] = "【圖4】部門分布比較（完整 vs 不完整）"
    ws4['A60'].font = section_font
    ws4['A61'] = "部門"
    ws4['B61'] = "完整資料"
    ws4['C61'] = "不完整資料"
    for col in range(1, 4):
        cell = ws4.cell(row=61, column=col)
        cell.font = header_font
        cell.fill = header_fill_chart
        cell.border = thin_border
        cell.alignment = header_alignment

    all_depts = sorted(set(list(complete_dept.keys()) + list(incomplete_dept.keys())))
    for i, dept in enumerate(all_depts[:15]):
        ws4.cell(row=62+i, column=1, value=dept[:30]).border = thin_border
        ws4.cell(row=62+i, column=2, value=complete_dept.get(dept, 0)).border = thin_border
        ws4.cell(row=62+i, column=3, value=incomplete_dept.get(dept, 0)).border = thin_border
        for c in range(1, 4):
            ws4.cell(row=62+i, column=c).font = number_font

    if all_depts:
        bar3 = BarChart()
        bar3.type = "col"
        bar3.grouping = "clustered"
        bar3.title = "部門分布比較（完整 vs 不完整）"
        data4 = Reference(ws4, min_col=2, min_row=61, max_col=3, max_row=61+min(len(all_depts), 15))
        cats4 = Reference(ws4, min_col=1, min_row=62, max_row=61+min(len(all_depts), 15))
        bar3.add_data(data4, titles_from_data=True)
        bar3.set_categories(cats4)
        bar3.width = 18
        bar3.height = 10
        ws4.add_chart(bar3, "E60")

    # 圖5 就業狀態
    ws4['A85'] = "【圖5】就業狀態分布比較（完整 vs 不完整）"
    ws4['A85'].font = section_font
    ws4['A86'] = "就業狀態"
    ws4['B86'] = "完整資料"
    ws4['C86'] = "不完整資料"
    for col in range(1, 4):
        cell = ws4.cell(row=86, column=col)
        cell.font = header_font
        cell.fill = header_fill_chart
        cell.border = thin_border
        cell.alignment = header_alignment

    all_status = sorted(set(list(complete_status.keys()) + list(incomplete_status.keys())))
    for i, status in enumerate(all_status):
        ws4.cell(row=87+i, column=1, value=status).border = thin_border
        ws4.cell(row=87+i, column=2, value=complete_status.get(status, 0)).border = thin_border
        ws4.cell(row=87+i, column=3, value=incomplete_status.get(status, 0)).border = thin_border
        for c in range(1, 4):
            ws4.cell(row=87+i, column=c).font = number_font

    if all_status:
        bar4 = BarChart()
        bar4.type = "col"
        bar4.grouping = "clustered"
        bar4.title = "就業狀態分布比較"
        data5 = Reference(ws4, min_col=2, min_row=86, max_col=3, max_row=86+len(all_status))
        cats5 = Reference(ws4, min_col=1, min_row=87, max_row=86+len(all_status))
        bar4.add_data(data5, titles_from_data=True)
        bar4.set_categories(cats5)
        bar4.width = 16
        bar4.height = 10
        ws4.add_chart(bar4, "E85")

    ws4.column_dimensions['A'].width = 42
    ws4.column_dimensions['B'].width = 14
    ws4.column_dimensions['C'].width = 14

    # 輸出到記憶體
    output = io.BytesIO()
    wb.save(output)
    output.seek(0)
    return output


# -------------------- Streamlit 主流程 --------------------
uploaded_file = st.file_uploader(
    "📂 請上傳 Excel 檔案（.xlsx）",
    type=["xlsx"],
    help="支援 .xlsx 格式"
)

if uploaded_file is not None:
    try:
        df = pd.read_excel(uploaded_file)
        st.success(f"✅ 成功讀取檔案：**{uploaded_file.name}**（共 {len(df)} 筆、{len(df.columns)} 個欄位）")

        with st.spinner("🔍 正在進行資料品質檢查，請稍候..."):
            quality_results = []
            for idx, row in df.iterrows():
                is_complete, issues = check_row_quality(row)
                quality_results.append({
                    'is_complete': is_complete,
                    'issue_count': len(issues),
                    'issues': ' | '.join(issues) if issues else ''
                })

            quality_df = pd.DataFrame(quality_results)
            df = pd.concat([df.reset_index(drop=True), quality_df], axis=1)

            complete_df = df[df['is_complete'] == True].drop(columns=['is_complete', 'issue_count', 'issues'])
            incomplete_df = df[df['is_complete'] == False].copy()
            incomplete_df = incomplete_df.rename(columns={
                'issue_count': '問題數量',
                'issues': '問題說明'
            }).drop(columns=['is_complete'])

            total_rows = len(df)
            complete_count = len(complete_df)
            incomplete_count = len(incomplete_df)
            complete_pct = round(complete_count / total_rows * 100, 2) if total_rows > 0 else 0
            incomplete_pct = round(incomplete_count / total_rows * 100, 2) if total_rows > 0 else 0

            # 問題類型統計
            issue_counter = {}
            if incomplete_count > 0 and '問題說明' in incomplete_df.columns:
                for issues_str in incomplete_df['問題說明']:
                    if pd.isna(issues_str) or issues_str == '':
                        continue
                    for issue in str(issues_str).split(' | '):
                        issue_type = re.sub(r'\s*\(.*?\)\s*', '', issue).strip()
                        if issue_type:
                            issue_counter[issue_type] = issue_counter.get(issue_type, 0) + 1
            issue_sorted = sorted(issue_counter.items(), key=lambda x: x[1], reverse=True)

            def count_null_like(series):
                return series.apply(is_null_or_empty).sum()

            null_stats = []
            key_cols = [
                'User_ID', 'First_Name', 'Last_Name', 'Email_Address', 'Phone_Number',
                'DOB', 'Age', 'Department', 'Job_Title', 'Employment_Status',
                'Salary_Monthly', 'Credit_Score', 'IP_Address', 'MAC_Address'
            ]
            for col in key_cols:
                if col in df.columns:
                    null_cnt = count_null_like(df[col])
                    null_stats.append([col, null_cnt, round(null_cnt / total_rows * 100, 2)])

            def safe_numeric_stats(series):
                numeric = pd.to_numeric(series, errors='coerce')
                return {
                    'count': numeric.notna().sum(),
                    'mean': round(numeric.mean(), 2) if numeric.notna().any() else None,
                    'min': numeric.min() if numeric.notna().any() else None,
                    'max': numeric.max() if numeric.notna().any() else None,
                }

            numeric_compare = []
            for col in ['Age', 'Salary_Monthly', 'Credit_Score', 'Years_of_Experience', 'Customer_Satisfaction_Score']:
                if col in complete_df.columns:
                    c_stats = safe_numeric_stats(complete_df[col])
                    i_stats = safe_numeric_stats(incomplete_df[col]) if col in incomplete_df.columns else {}
                    numeric_compare.append({
                        '欄位': col,
                        '完整_有效筆數': c_stats['count'],
                        '完整_平均': c_stats['mean'],
                        '完整_最小': c_stats['min'],
                        '完整_最大': c_stats['max'],
                        '不完整_有效筆數': i_stats.get('count', 0),
                        '不完整_平均': i_stats.get('mean'),
                        '不完整_最小': i_stats.get('min'),
                        '不完整_最大': i_stats.get('max'),
                    })

            def get_dept_counts(dataframe):
                if 'Department' not in dataframe.columns:
                    return {}
                return dataframe['Department'].fillna('未知').astype(str).str.strip().value_counts().to_dict()

            complete_dept = get_dept_counts(complete_df)
            incomplete_dept = get_dept_counts(incomplete_df)

            def get_status_counts(dataframe):
                if 'Employment_Status' not in dataframe.columns:
                    return {}
                return dataframe['Employment_Status'].fillna('未知').astype(str).str.strip().value_counts().to_dict()

            complete_status = get_status_counts(complete_df)
            incomplete_status = get_status_counts(incomplete_df)

        # ========== 顯示結果 ==========
        st.markdown("---")
        st.subheader("📈 資料完整性總覽")

        col1, col2, col3, col4 = st.columns(4)
        col1.metric("總筆數", total_rows)
        col2.metric("完整資料", complete_count, f"{complete_pct}%")
        col3.metric("不完整資料", incomplete_count, f"{incomplete_pct}%")
        col4.metric("完整比例", f"{complete_pct}%")

        # 互動圖表
        st.markdown("### 📊 互動圖表")
        tab1, tab2, tab3, tab4 = st.tabs(["完整性比例", "問題類型 Top10", "部門分布", "就業狀態"])

        with tab1:
            fig_pie = go.Figure(data=[go.Pie(
                labels=['完整資料', '不完整資料'],
                values=[complete_count, incomplete_count],
                hole=0.4,
                marker_colors=['#2E7D32', '#C62828']
            )])
            fig_pie.update_layout(title="完整 vs 不完整資料比例")
            st.plotly_chart(fig_pie, use_container_width=True)

        with tab2:
            if issue_sorted:
                issue_df = pd.DataFrame(issue_sorted[:10], columns=['問題類型', '出現次數'])
                fig_bar = px.bar(issue_df, x='問題類型', y='出現次數', title="問題類型出現次數 Top 10",
                                color='出現次數', color_continuous_scale='Reds')
                fig_bar.update_layout(xaxis_tickangle=-45)
                st.plotly_chart(fig_bar, use_container_width=True)
            else:
                st.info("沒有偵測到問題類型")

        with tab3:
            all_depts = sorted(set(list(complete_dept.keys()) + list(incomplete_dept.keys())))
            dept_data = []
            for d in all_depts:
                dept_data.append({'部門': d, '類型': '完整', '人數': complete_dept.get(d, 0)})
                dept_data.append({'部門': d, '類型': '不完整', '人數': incomplete_dept.get(d, 0)})
            if dept_data:
                dept_df = pd.DataFrame(dept_data)
                fig_dept = px.bar(dept_df, x='部門', y='人數', color='類型', barmode='group',
                                 title="部門分布比較", color_discrete_map={'完整': '#2E7D32', '不完整': '#C62828'})
                fig_dept.update_layout(xaxis_tickangle=-45)
                st.plotly_chart(fig_dept, use_container_width=True)

        with tab4:
            all_status = sorted(set(list(complete_status.keys()) + list(incomplete_status.keys())))
            status_data = []
            for s in all_status:
                status_data.append({'就業狀態': s, '類型': '完整', '人數': complete_status.get(s, 0)})
                status_data.append({'就業狀態': s, '類型': '不完整', '人數': incomplete_status.get(s, 0)})
            if status_data:
                status_df = pd.DataFrame(status_data)
                fig_status = px.bar(status_df, x='就業狀態', y='人數', color='類型', barmode='group',
                                   title="就業狀態分布比較", color_discrete_map={'完整': '#2E7D32', '不完整': '#C62828'})
                st.plotly_chart(fig_status, use_container_width=True)

        # 資料預覽
        st.markdown("### 📋 資料預覽")
        preview_tab1, preview_tab2 = st.tabs(["完整資料", "不完整資料"])
        with preview_tab1:
            st.dataframe(complete_df, use_container_width=True, height=400)
        with preview_tab2:
            st.dataframe(incomplete_df, use_container_width=True, height=400)

        # 下載按鈕
        st.markdown("---")
        st.subheader("⬇️ 下載整理後的 Excel 報表")

        excel_buffer = create_excel_report(
            complete_df, incomplete_df, issue_sorted, null_stats, numeric_compare,
            complete_dept, incomplete_dept, complete_status, incomplete_status,
            complete_count, incomplete_count, total_rows, complete_pct, incomplete_pct
        )

        st.download_button(
            label="📥 下載完整報表（含 Complete / Incomplete / Statistics / Charts 四個工作表）",
            data=excel_buffer,
            file_name=f"整理後_組織資料_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            type="primary"
        )

        st.info("報表包含四個工作表：\n"
                "1. **Complete_Data** – 完整資料\n"
                "2. **Incomplete_Data** – 不完整資料（含問題說明）\n"
                "3. **Statistics** – 統計報表\n"
                "4. **Charts** – Excel 內建圖表")

    except Exception as e:
        st.error(f"❌ 處理檔案時發生錯誤：{str(e)}")
        st.exception(e)

else:
    st.info("👆 請先上傳 Excel 檔案開始分析。")
    st.markdown("""
    ### 使用說明
    1. 點擊上方上傳按鈕，選擇 `.xlsx` 檔案
    2. 系統會自動進行資料品質檢查（空值、格式、異常值、大小寫等）
    3. 查看互動圖表與資料預覽
    4. 下載包含四個工作表的完整 Excel 報表
    """)