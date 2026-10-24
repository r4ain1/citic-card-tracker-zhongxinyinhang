#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
报表与系统文件生成器
输出专业 Excel 账簿系统（多标签页、公式、格式化、条件高亮）以及离线交互式 HTML 仪表盘。
"""

import os
import json
from typing import List, Dict, Any
import pandas as pd
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter


def create_excel_system(
    output_path: str,
    df_all: pd.DataFrame,
    summary_stats: Dict[str, Any],
    campaign_results: List[Dict[str, Any]],
    start_date: str,
    end_date: str
):
    """
    生成包含 5 个工作表的专业 Excel 系统工作簿
    """
    wb = Workbook()
    # 移除默认 sheet
    wb.remove(wb.active)
    
    # 样式定义
    font_title = Font(name='微软雅黑', size=16, bold=True, color='1F2937')
    font_subtitle = Font(name='微软雅黑', size=11, color='4B5563')
    font_section = Font(name='微软雅黑', size=12, bold=True, color='111827')
    font_header = Font(name='微软雅黑', size=10, bold=True, color='FFFFFF')
    font_body = Font(name='微软雅黑', size=10, color='1F2937')
    font_bold = Font(name='微软雅黑', size=10, bold=True, color='1F2937')
    font_green = Font(name='微软雅黑', size=10, bold=True, color='047857')
    font_red = Font(name='微软雅黑', size=10, bold=True, color='B91C1C')
    
    fill_header_navy = PatternFill(start_color='1E3A8A', end_color='1E3A8A', fill_type='solid')
    fill_header_blue = PatternFill(start_color='2563EB', end_color='2563EB', fill_type='solid')
    fill_header_teal = PatternFill(start_color='0D9488', end_color='0D9488', fill_type='solid')
    fill_header_purple = PatternFill(start_color='7C3AED', end_color='7C3AED', fill_type='solid')
    fill_header_amber = PatternFill(start_color='D97706', end_color='D97706', fill_type='solid')
    fill_kpi = PatternFill(start_color='F3F4F6', end_color='F3F4F6', fill_type='solid')
    fill_zebra = PatternFill(start_color='F9FAFB', end_color='F9FAFB', fill_type='solid')
    fill_achieved = PatternFill(start_color='D1FAE5', end_color='D1FAE5', fill_type='solid')
    fill_pending = PatternFill(start_color='FEF3C7', end_color='FEF3C7', fill_type='solid')
    
    thin_border_side = Side(border_style='thin', color='E5E7EB')
    border_cell = Border(left=thin_border_side, right=thin_border_side, top=thin_border_side, bottom=thin_border_side)
    
    align_left = Alignment(horizontal='left', vertical='center')
    align_right = Alignment(horizontal='right', vertical='center')
    align_center = Alignment(horizontal='center', vertical='center')
    
    # ----------------------------------------------------
    # Sheet 1: 📊 汇总看板
    # ----------------------------------------------------
    ws1 = wb.create_sheet(title="📊 汇总看板")
    ws1.views.sheetView[0].showGridLines = True
    
    ws1['B2'] = "中信银行信用卡当月消费与汇总系统"
    ws1['B2'].font = font_title
    ws1['B3'] = f"统计区间: {start_date} 至 {end_date}  |  生成时间: {pd.Timestamp.now().strftime('%Y-%m-%d %H:%M')}"
    ws1['B3'].font = font_subtitle
    
    # KPI 指标卡片
    kpi_cards = [
        ("当月净支出", f"¥ {summary_stats['net_spend']:,.2f}", "总支出 - 总退款", font_blue := Font(name='微软雅黑', size=14, bold=True, color='1D4ED8')),
        ("总支出消费", f"¥ {summary_stats['total_expense']:,.2f}", f"共 {summary_stats['expense_count']} 笔消费", font_red),
        ("总退款返还", f"¥ {summary_stats['total_refund']:,.2f}", f"共 {summary_stats['refund_count']} 笔退款", font_green),
        ("总交易笔数", f"{summary_stats['total_count']} 笔", "合并去重后记录", font_bold)
    ]
    
    for idx, (title, val, sub, val_font) in enumerate(kpi_cards):
        col_idx = 2 + idx * 2  # B, D, F, H
        ws1.merge_cells(start_row=5, start_column=col_idx, end_row=5, end_column=col_idx+1)
        ws1.merge_cells(start_row=6, start_column=col_idx, end_row=6, end_column=col_idx+1)
        ws1.merge_cells(start_row=7, start_column=col_idx, end_row=7, end_column=col_idx+1)
        
        cell_t = ws1.cell(row=5, column=col_idx, value=title)
        cell_v = ws1.cell(row=6, column=col_idx, value=val)
        cell_s = ws1.cell(row=7, column=col_idx, value=sub)
        
        cell_t.font = Font(name='微软雅黑', size=9, color='6B7280')
        cell_v.font = val_font
        cell_s.font = Font(name='微软雅黑', size=8, color='9CA3AF')
        
        for r in range(5, 8):
            for c in range(col_idx, col_idx + 2):
                ws1.cell(row=r, column=c).fill = fill_kpi
                ws1.cell(row=r, column=c).border = border_cell
                ws1.cell(row=r, column=c).alignment = align_center

    # 渠道汇总表格
    ws1['B9'] = "1. 各支付渠道支出与净额分布"
    ws1['B9'].font = font_section
    
    ch_headers = ["支付渠道", "支出金额 (元)", "退款金额 (元)", "净支出金额 (元)", "占比 (%)", "交易笔数"]
    for c_i, h in enumerate(ch_headers, start=2):
        cell = ws1.cell(row=10, column=c_i, value=h)
        cell.font = font_header
        cell.fill = fill_header_blue
        cell.alignment = align_center
        cell.border = border_cell
        
    cur_r = 11
    for item in summary_stats['channel_summary']:
        ws1.cell(row=cur_r, column=2, value=item['channel']).alignment = align_left
        ws1.cell(row=cur_r, column=3, value=item['expense']).alignment = align_right
        ws1.cell(row=cur_r, column=4, value=item['refund']).alignment = align_right
        ws1.cell(row=cur_r, column=5, value=item['net']).alignment = align_right
        ws1.cell(row=cur_r, column=6, value=f"{item['percentage']}%").alignment = align_right
        ws1.cell(row=cur_r, column=7, value=item['count']).alignment = align_center
        
        # 格式化数字
        ws1.cell(row=cur_r, column=3).number_format = '¥#,##0.00'
        ws1.cell(row=cur_r, column=4).number_format = '¥#,##0.00'
        ws1.cell(row=cur_r, column=5).number_format = '¥#,##0.00'
        
        for c_i in range(2, 8):
            ws1.cell(row=cur_r, column=c_i).font = font_body
            ws1.cell(row=cur_r, column=c_i).border = border_cell
            if cur_r % 2 == 1:
                ws1.cell(row=cur_r, column=c_i).fill = fill_zebra
        cur_r += 1
        
    # 渠道合计行
    ws1.cell(row=cur_r, column=2, value="合计").alignment = align_left
    ws1.cell(row=cur_r, column=3, value=summary_stats['total_expense']).alignment = align_right
    ws1.cell(row=cur_r, column=4, value=summary_stats['total_refund']).alignment = align_right
    ws1.cell(row=cur_r, column=5, value=summary_stats['net_spend']).alignment = align_right
    ws1.cell(row=cur_r, column=6, value="100.0%").alignment = align_right
    ws1.cell(row=cur_r, column=7, value=summary_stats['total_count']).alignment = align_center
    for c_i in range(2, 8):
        ws1.cell(row=cur_r, column=c_i).font = font_bold
        ws1.cell(row=cur_r, column=c_i).border = border_cell
    ws1.cell(row=cur_r, column=3).number_format = '¥#,##0.00'
    ws1.cell(row=cur_r, column=4).number_format = '¥#,##0.00'
    ws1.cell(row=cur_r, column=5).number_format = '¥#,##0.00'
    
    # 卡片汇总表格
    cur_r += 3
    ws1.cell(row=cur_r, column=2, value="2. 卡片明细汇总").font = font_section
    cur_r += 1
    card_headers = ["卡号", "总支出 (元)", "总退款 (元)", "净支出 (元)", "笔数"]
    for c_i, h in enumerate(card_headers, start=2):
        cell = ws1.cell(row=cur_r, column=c_i, value=h)
        cell.font = font_header
        cell.fill = fill_header_teal
        cell.alignment = align_center
        cell.border = border_cell
    cur_r += 1
    for c_item in summary_stats['card_summary']:
        ws1.cell(row=cur_r, column=2, value=c_item['card_no']).alignment = align_left
        ws1.cell(row=cur_r, column=3, value=c_item['expense']).alignment = align_right
        ws1.cell(row=cur_r, column=4, value=c_item['refund']).alignment = align_right
        ws1.cell(row=cur_r, column=5, value=c_item['net']).alignment = align_right
        ws1.cell(row=cur_r, column=6, value=c_item['count']).alignment = align_center
        ws1.cell(row=cur_r, column=3).number_format = '¥#,##0.00'
        ws1.cell(row=cur_r, column=4).number_format = '¥#,##0.00'
        ws1.cell(row=cur_r, column=5).number_format = '¥#,##0.00'
        for c_i in range(2, 7):
            ws1.cell(row=cur_r, column=c_i).font = font_body
            ws1.cell(row=cur_r, column=c_i).border = border_cell
        cur_r += 1

    # ----------------------------------------------------
    # Sheet 2: 🎯 活动达标监控
    # ----------------------------------------------------
    ws2 = wb.create_sheet(title="🎯 活动达标监控")
    ws2.views.sheetView[0].showGridLines = True
    
    ws2['B2'] = "信用卡活动与特定渠道达标进度监控"
    ws2['B2'].font = font_title
    ws2['B3'] = f"评估基准日: {end_date} (当月剩余 {campaign_results[0]['remaining_days'] if campaign_results else 0} 天)"
    ws2['B3'].font = font_subtitle
    
    cmp_headers = ["活动名称", "限定渠道/条件", "目标金额 (元)", "当前累计净消费", "完成进度", "还差金额 (元)", "状态", "建议日均消费"]
    for c_i, h in enumerate(cmp_headers, start=2):
        cell = ws2.cell(row=5, column=c_i, value=h)
        cell.font = font_header
        cell.fill = fill_header_purple
        cell.alignment = align_center
        cell.border = border_cell
        
    r2 = 6
    for cmp in campaign_results:
        ws2.cell(row=r2, column=2, value=cmp['name']).alignment = align_left
        ws2.cell(row=r2, column=3, value="、".join(cmp['channels'])).alignment = align_left
        ws2.cell(row=r2, column=4, value=cmp['target_amount']).alignment = align_right
        ws2.cell(row=r2, column=5, value=cmp['current_spend']).alignment = align_right
        ws2.cell(row=r2, column=6, value=f"{cmp['progress_pct']}%").alignment = align_right
        ws2.cell(row=r2, column=7, value=cmp['gap_amount']).alignment = align_right
        ws2.cell(row=r2, column=8, value=cmp['status_text']).alignment = align_center
        ws2.cell(row=r2, column=9, value=cmp['suggested_daily']).alignment = align_right
        
        ws2.cell(row=r2, column=4).number_format = '¥#,##0.00'
        ws2.cell(row=r2, column=5).number_format = '¥#,##0.00'
        ws2.cell(row=r2, column=7).number_format = '¥#,##0.00'
        ws2.cell(row=r2, column=9).number_format = '¥#,##0.00'
        
        for c_i in range(2, 10):
            ws2.cell(row=r2, column=c_i).font = font_body
            ws2.cell(row=r2, column=c_i).border = border_cell
            if cmp['is_reached']:
                ws2.cell(row=r2, column=8).font = font_green
                ws2.cell(row=r2, column=8).fill = fill_achieved
            else:
                ws2.cell(row=r2, column=8).font = font_red
                ws2.cell(row=r2, column=8).fill = fill_pending
        r2 += 1

    # ----------------------------------------------------
    # Sheet 3: 📑 全部合并交易明细
    # ----------------------------------------------------
    ws3 = wb.create_sheet(title="📑 全部合并交易明细")
    ws3.views.sheetView[0].showGridLines = True
    
    det_headers = ["交易日期", "记账日期", "卡号尾号", "交易描述", "识别渠道", "类型", "金额 (元)", "币种", "账单来源"]
    for c_i, h in enumerate(det_headers, start=1):
        cell = ws3.cell(row=1, column=c_i, value=h)
        cell.font = font_header
        cell.fill = fill_header_navy
        cell.alignment = align_center
        cell.border = border_cell
        
    for r_i, (_, row) in enumerate(df_all.iterrows(), start=2):
        ws3.cell(row=r_i, column=1, value=row['txn_date']).alignment = align_center
        ws3.cell(row=r_i, column=2, value=row['post_date']).alignment = align_center
        ws3.cell(row=r_i, column=3, value=row['card_no']).alignment = align_center
        ws3.cell(row=r_i, column=4, value=row['description']).alignment = align_left
        ws3.cell(row=r_i, column=5, value=row['channel']).alignment = align_center
        ws3.cell(row=r_i, column=6, value=row['trans_type']).alignment = align_center
        ws3.cell(row=r_i, column=7, value=row['amount']).alignment = align_right
        ws3.cell(row=r_i, column=8, value=row['currency']).alignment = align_center
        ws3.cell(row=r_i, column=9, value=row['source']).alignment = align_center
        
        ws3.cell(row=r_i, column=7).number_format = '¥#,##0.00'
        
        for c_i in range(1, 10):
            ws3.cell(row=r_i, column=c_i).font = font_body
            ws3.cell(row=r_i, column=c_i).border = border_cell
            if row['trans_type'] in ['商户退款', '退款/还款', '信用卡还款']:
                ws3.cell(row=r_i, column=6).font = font_green
            else:
                ws3.cell(row=r_i, column=6).font = font_red
                
    # ----------------------------------------------------
    # Sheet 4: 🔴 仅支出明细
    # ----------------------------------------------------
    ws4 = wb.create_sheet(title="🔴 仅支出明细")
    ws4.views.sheetView[0].showGridLines = True
    df_exp = df_all[df_all['trans_type'] == '支出'].sort_values(by='amount', ascending=False)
    for c_i, h in enumerate(det_headers, start=1):
        cell = ws4.cell(row=1, column=c_i, value=h)
        cell.font = font_header
        cell.fill = fill_header_amber
        cell.alignment = align_center
        cell.border = border_cell
        
    for r_i, (_, row) in enumerate(df_exp.iterrows(), start=2):
        ws4.cell(row=r_i, column=1, value=row['txn_date']).alignment = align_center
        ws4.cell(row=r_i, column=2, value=row['post_date']).alignment = align_center
        ws4.cell(row=r_i, column=3, value=row['card_no']).alignment = align_center
        ws4.cell(row=r_i, column=4, value=row['description']).alignment = align_left
        ws4.cell(row=r_i, column=5, value=row['channel']).alignment = align_center
        ws4.cell(row=r_i, column=6, value=row['trans_type']).alignment = align_center
        ws4.cell(row=r_i, column=7, value=row['amount']).alignment = align_right
        ws4.cell(row=r_i, column=8, value=row['currency']).alignment = align_center
        ws4.cell(row=r_i, column=9, value=row['source']).alignment = align_center
        ws4.cell(row=r_i, column=7).number_format = '¥#,##0.00'
        for c_i in range(1, 10):
            ws4.cell(row=r_i, column=c_i).font = font_body
            ws4.cell(row=r_i, column=c_i).border = border_cell

    # ----------------------------------------------------
    # Sheet 5: 🟢 仅退款明细
    # ----------------------------------------------------
    ws5 = wb.create_sheet(title="🟢 仅退款明细")
    ws5.views.sheetView[0].showGridLines = True
    df_ref = df_all[df_all['trans_type'].isin(['商户退款', '退款/还款', '信用卡还款'])].sort_values(by='amount', ascending=False)
    for c_i, h in enumerate(det_headers, start=1):
        cell = ws5.cell(row=1, column=c_i, value=h)
        cell.font = font_header
        cell.fill = fill_header_teal
        cell.alignment = align_center
        cell.border = border_cell
        
    for r_i, (_, row) in enumerate(df_ref.iterrows(), start=2):
        ws5.cell(row=r_i, column=1, value=row['txn_date']).alignment = align_center
        ws5.cell(row=r_i, column=2, value=row['post_date']).alignment = align_center
        ws5.cell(row=r_i, column=3, value=row['card_no']).alignment = align_center
        ws5.cell(row=r_i, column=4, value=row['description']).alignment = align_left
        ws5.cell(row=r_i, column=5, value=row['channel']).alignment = align_center
        ws5.cell(row=r_i, column=6, value=row['trans_type']).alignment = align_center
        ws5.cell(row=r_i, column=7, value=row['amount']).alignment = align_right
        ws5.cell(row=r_i, column=8, value=row['currency']).alignment = align_center
        ws5.cell(row=r_i, column=9, value=row['source']).alignment = align_center
        ws5.cell(row=r_i, column=7).number_format = '¥#,##0.00'
        for c_i in range(1, 10):
            ws5.cell(row=r_i, column=c_i).font = font_body
            ws5.cell(row=r_i, column=c_i).border = border_cell

    # 自动调整所有 Sheet 列宽
    for ws in [ws1, ws2, ws3, ws4, ws5]:
        for col in ws.columns:
            col_letter = get_column_letter(col[0].column)
            max_len = 0
            for cell in col:
                val_str = str(cell.value or '')
                # 中文字符按双倍长度计算
                cell_len = sum(2 if ord(c) > 127 else 1 for c in val_str)
                max_len = max(max_len, cell_len)
            ws.column_dimensions[col_letter].width = max(max_len + 3, 11)
            
    wb.save(output_path)
    print(f"[OK] Excel 账簿系统已生成: {output_path}")


def create_html_dashboard(
    output_path: str,
    df_all: pd.DataFrame,
    summary_stats: Dict[str, Any],
    campaign_results: List[Dict[str, Any]],
    start_date: str,
    end_date: str
):
    """
    生成自包含、零外部依赖的现代交互式 HTML 仪表盘
    """
    records = df_all.to_dict(orient='records')
    data_json = json.dumps(records, ensure_ascii=False)
    summary_json = json.dumps(summary_stats, ensure_ascii=False)
    campaigns_json = json.dumps(campaign_results, ensure_ascii=False)
    
    html_content = f"""<!DOCTYPE html>
<html lang="zh-CN">
<head>
  <meta charset="UTF-8">
  <meta name="viewport" content="width=device-width, initial-scale=1.0">
  <title>中信信用卡当月消费与达标监控系统</title>
  <style>
    :root {{
      --primary: #2563EB;
      --primary-dark: #1D4ED8;
      --success: #10B981;
      --warning: #F59E0B;
      --danger: #EF4444;
      --bg: #F8FAFC;
      --surface: #FFFFFF;
      --border: #E2E8F0;
      --text: #0F172A;
      --text-muted: #64748B;
      --radius: 12px;
      --shadow: 0 4px 6px -1px rgba(0, 0, 0, 0.05), 0 2px 4px -2px rgba(0, 0, 0, 0.05);
    }}
    * {{ box-sizing: border-box; margin: 0; padding: 0; font-family: -apple-system, BlinkMacSystemFont, "Segoe UI", Roboto, "PingFang SC", "Hiragino Sans GB", "Microsoft YaHei", sans-serif; }}
    body {{ background-color: var(--bg); color: var(--text); padding: 24px; line-height: 1.5; }}
    .container {{ max-width: 1280px; margin: 0 auto; }}
    
    header {{ margin-bottom: 24px; }}
    .badge {{ display: inline-block; padding: 4px 10px; background: #DBEAFE; color: var(--primary); border-radius: 9999px; font-size: 12px; font-weight: 600; margin-bottom: 8px; }}
    h1 {{ font-size: 26px; font-weight: 700; color: var(--text); }}
    .subtitle {{ color: var(--text-muted); font-size: 14px; margin-top: 4px; }}
    
    /* KPI Cards */
    .kpi-grid {{ display: grid; grid-template-columns: repeat(auto-fit, minmax(240px, 1fr)); gap: 16px; margin-bottom: 24px; }}
    .kpi-card {{ background: var(--surface); border: 1px solid var(--border); border-radius: var(--radius); padding: 20px; box-shadow: var(--shadow); }}
    .kpi-title {{ font-size: 13px; color: var(--text-muted); font-weight: 500; text-transform: uppercase; letter-spacing: 0.5px; }}
    .kpi-value {{ font-size: 28px; font-weight: 700; margin: 8px 0 4px; }}
    .kpi-sub {{ font-size: 13px; color: var(--text-muted); }}
    .c-primary {{ color: var(--primary); }}
    .c-danger {{ color: var(--danger); }}
    .c-success {{ color: var(--success); }}
    
    /* Campaign Cards */
    .section-title {{ font-size: 18px; font-weight: 700; margin: 32px 0 16px; display: flex; align-items: center; justify-content: space-between; }}
    .campaign-grid {{ display: grid; grid-template-columns: repeat(auto-fit, minmax(290px, 1fr)); gap: 16px; margin-bottom: 24px; }}
    .campaign-card {{ background: var(--surface); border: 1px solid var(--border); border-radius: var(--radius); padding: 18px; box-shadow: var(--shadow); position: relative; overflow: hidden; }}
    .campaign-card.achieved {{ border-left: 5px solid var(--success); }}
    .campaign-card.pending {{ border-left: 5px solid var(--warning); }}
    .campaign-name {{ font-size: 15px; font-weight: 600; margin-bottom: 4px; }}
    .campaign-desc {{ font-size: 12px; color: var(--text-muted); margin-bottom: 12px; }}
    .progress-bar-bg {{ height: 10px; background: #E2E8F0; border-radius: 9999px; overflow: hidden; margin-bottom: 8px; }}
    .progress-bar-fill {{ height: 100%; border-radius: 9999px; transition: width 0.4s ease; }}
    .campaign-stats {{ display: flex; justify-content: space-between; font-size: 13px; }}
    .status-tag {{ display: inline-block; padding: 2px 8px; border-radius: 6px; font-size: 11px; font-weight: 600; }}
    .tag-achieved {{ background: #D1FAE5; color: #065F46; }}
    .tag-pending {{ background: #FEF3C7; color: #92400E; }}
    
    /* Channels & Layout */
    .two-cols {{ display: grid; grid-template-columns: 1fr 1fr; gap: 20px; margin-bottom: 24px; }}
    @media (max-width: 900px) {{ .two-cols {{ grid-template-columns: 1fr; }} }}
    .panel {{ background: var(--surface); border: 1px solid var(--border); border-radius: var(--radius); padding: 20px; box-shadow: var(--shadow); }}
    
    .channel-item {{ display: flex; align-items: center; justify-content: space-between; padding: 10px 0; border-bottom: 1px solid var(--border); font-size: 14px; }}
    .channel-item:last-child {{ border-bottom: none; }}
    .channel-bar-wrap {{ flex: 1; margin: 0 16px; background: #E2E8F0; height: 8px; border-radius: 4px; overflow: hidden; }}
    .channel-bar {{ height: 100%; background: var(--primary); border-radius: 4px; }}
    
    /* Table & Controls */
    .table-controls {{ display: flex; flex-wrap: wrap; gap: 12px; margin-bottom: 16px; background: var(--surface); padding: 16px; border-radius: var(--radius); border: 1px solid var(--border); }}
    .table-controls input, .table-controls select {{ padding: 8px 12px; border: 1px solid var(--border); border-radius: 8px; font-size: 14px; outline: none; }}
    .table-controls input:focus, .table-controls select:focus {{ border-color: var(--primary); }}
    .table-wrap {{ background: var(--surface); border: 1px solid var(--border); border-radius: var(--radius); overflow: auto; box-shadow: var(--shadow); max-height: 600px; }}
    table {{ width: 100%; border-collapse: collapse; font-size: 13px; text-align: left; }}
    th {{ background: #F1F5F9; padding: 12px 14px; font-weight: 600; color: #475569; position: sticky; top: 0; z-index: 1; }}
    td {{ padding: 10px 14px; border-bottom: 1px solid var(--border); }}
    tr:hover td {{ background: #F8FAFC; }}
    .type-exp {{ color: var(--danger); font-weight: 600; }}
    .type-ref {{ color: var(--success); font-weight: 600; }}
  </style>
</head>
<body>
  <div class="container">
    <header>
      <span class="badge">中信银行信用卡 · 自动化汇总系统</span>
      <h1>月度消费与达标监控仪表盘</h1>
      <div class="subtitle">统计周期: <strong>{start_date}</strong> 至 <strong>{end_date}</strong> (仅统计当月1日到截止日已出与未出账单)</div>
    </header>
    
    <!-- KPI Grid -->
    <div class="kpi-grid">
      <div class="kpi-card">
        <div class="kpi-title">当月净消费 (支出 - 退款)</div>
        <div class="kpi-value c-primary">¥{summary_stats['net_spend']:,.2f}</div>
        <div class="kpi-sub">共统计 {summary_stats['total_count']} 笔交易 (已自动合并去重)</div>
      </div>
      <div class="kpi-card">
        <div class="kpi-title">当月总支出</div>
        <div class="kpi-value c-danger">¥{summary_stats['total_expense']:,.2f}</div>
        <div class="kpi-sub">累计消费 {summary_stats['expense_count']} 笔</div>
      </div>
      <div class="kpi-card">
        <div class="kpi-title">当月总退款 / 返还</div>
        <div class="kpi-value c-success">¥{summary_stats['total_refund']:,.2f}</div>
        <div class="kpi-sub">累计退款/对冲 {summary_stats['refund_count']} 笔</div>
      </div>
      <div class="kpi-card">
        <div class="kpi-title">统计进度 / 剩余天数</div>
        <div class="kpi-value">{end_date.split('-')[-1]} 日</div>
        <div class="kpi-sub">当月剩余 {campaign_results[0]['remaining_days'] if campaign_results else 0} 天达标冲刺窗口</div>
      </div>
    </div>
    
    <!-- Campaigns Section -->
    <div class="section-title">
      <span>🎯 信用卡活动达标监控</span>
      <span style="font-size:13px; font-weight:normal; color:var(--text-muted);">自动按渠道过滤与达标差额测算</span>
    </div>
    <div class="campaign-grid">
"""
    for cmp in campaign_results:
        is_ach = cmp['is_reached']
        card_cls = "achieved" if is_ach else "pending"
        fill_color = "var(--success)" if is_ach else "var(--primary)"
        tag_cls = "tag-achieved" if is_ach else "tag-pending"
        
        html_content += f"""
      <div class="campaign-card {card_cls}">
        <div style="display:flex; justify-content:space-between; align-items:center; margin-bottom:6px;">
          <div class="campaign-name">{cmp['name']}</div>
          <span class="status-tag {tag_cls}">{cmp['status_text']}</span>
        </div>
        <div class="campaign-desc">{cmp['description']}</div>
        <div class="progress-bar-bg">
          <div class="progress-bar-fill" style="width: {min(cmp['progress_pct'], 100.0)}%; background: {fill_color};"></div>
        </div>
        <div class="campaign-stats">
          <span>已累计: <strong>¥{cmp['current_spend']:,.2f}</strong></span>
          <span>目标: <strong>¥{cmp['target_amount']:,.2f}</strong> ({cmp['progress_pct']}%)</span>
        </div>
        {"<div style='margin-top:8px; font-size:12px; color:var(--text-muted);'>建议后续日均消费: <strong style=\"color:var(--danger)\">¥" + f"{cmp['suggested_daily']:,.2f}</strong>/天 (剩余{cmp['remaining_days']}天)</div>" if not is_ach and cmp['remaining_days'] > 0 else ""}
      </div>
"""

    html_content += f"""
    </div>
    
    <!-- Channels & Cards Section -->
    <div class="two-cols">
      <div class="panel">
        <div style="font-size:16px; font-weight:700; margin-bottom:16px;">💳 支付渠道分布 (净支出)</div>
        <div id="channelList">
"""
    max_ch_net = max([c['net'] for c in summary_stats['channel_summary']]) if summary_stats['channel_summary'] else 1.0
    for ch in summary_stats['channel_summary']:
        bar_w = round(max(0, ch['net']) / max_ch_net * 100.0, 1) if max_ch_net > 0 else 0
        html_content += f"""
          <div class="channel-item">
            <span style="width:110px; font-weight:500;">{ch['channel']}</span>
            <div class="channel-bar-wrap">
              <div class="channel-bar" style="width:{bar_w}%;"></div>
            </div>
            <div style="text-align:right; width:130px;">
              <strong>¥{ch['net']:,.2f}</strong>
              <span style="color:var(--text-muted); font-size:12px; margin-left:4px;">({ch['percentage']}%)</span>
            </div>
          </div>
"""
    html_content += f"""
        </div>
      </div>
      
      <div class="panel">
        <div style="font-size:16px; font-weight:700; margin-bottom:16px;">🪪 卡片汇总明细</div>
        <div>
"""
    for card in summary_stats['card_summary']:
        html_content += f"""
          <div class="channel-item">
            <span style="font-weight:600; width:100px;">{card['card_no']}</span>
            <span>支出: ¥{card['expense']:,.2f} | 退款: ¥{card['refund']:,.2f}</span>
            <span style="font-weight:700; color:var(--primary);">净额: ¥{card['net']:,.2f} ({card['count']}笔)</span>
          </div>
"""
    html_content += f"""
        </div>
      </div>
    </div>
    
    <!-- Details Table -->
    <div class="section-title">
      <span>📑 整合交易明细 (共 <span id="recordCount">{len(records)}</span> 笔)</span>
    </div>
    <div class="table-controls">
      <input type="text" id="searchInput" placeholder="搜索商户名/描述/卡号..." oninput="filterTable()">
      <select id="channelFilter" onchange="filterTable()">
        <option value="">全部支付渠道</option>
"""
    for ch in summary_stats['channel_summary']:
        html_content += f'<option value="{ch["channel"]}">{ch["channel"]}</option>'
        
    html_content += f"""
      </select>
      <select id="typeFilter" onchange="filterTable()">
        <option value="">全部交易类型</option>
        <option value="支出">仅支出</option>
        <option value="商户退款">仅商户退款</option><option value="信用卡还款">仅信用卡还款</option>
      </select>
    </div>
    
    <div class="table-wrap">
      <table id="txnTable">
        <thead>
          <tr>
            <th>交易日</th>
            <th>记账日</th>
            <th>卡号</th>
            <th>交易描述</th>
            <th>识别渠道</th>
            <th>交易类型</th>
            <th style="text-align:right;">金额</th>
            <th>来源</th>
          </tr>
        </thead>
        <tbody id="tableBody">
        </tbody>
      </table>
    </div>
  </div>

  <script>
    const rawData = {data_json};
    
    function renderTable(data) {{
      const tbody = document.getElementById('tableBody');
      tbody.innerHTML = '';
      data.forEach(item => {{
        const tr = document.createElement('tr');
        const isRefund = (item.trans_type === '商户退款' || item.trans_type === '退款/还款' || item.trans_type === '信用卡还款');
        tr.innerHTML = `
          <td>${{item.txn_date}}</td>
          <td>${{item.post_date}}</td>
          <td>尾号${{item.card_no}}</td>
          <td>${{item.description}}</td>
          <td><span style="padding:2px 6px; background:#F1F5F9; border-radius:4px;">${{item.channel}}</span></td>
          <td class="${{isRefund ? 'type-ref' : 'type-exp'}}">${{item.trans_type}}</td>
          <td style="text-align:right; font-weight:600; font-family:monospace;">${{isRefund ? '-' : ''}}¥${{Number(item.amount).toLocaleString('zh-CN', {{minimumFractionDigits: 2}})}}</td>
          <td style="color:#64748B;">${{item.source}}</td>
        `;
        tbody.appendChild(tr);
      }});
      document.getElementById('recordCount').innerText = data.length;
    }}
    
    function filterTable() {{
      const query = document.getElementById('searchInput').value.trim().toLowerCase();
      const ch = document.getElementById('channelFilter').value;
      const tp = document.getElementById('typeFilter').value;
      
      const filtered = rawData.filter(item => {{
        if (query && !item.description.toLowerCase().includes(query) && !item.card_no.includes(query) && !item.channel.toLowerCase().includes(query)) {{
          return false;
        }}
        if (ch && item.channel !== ch) return false;
        if (tp && item.trans_type !== tp) return false;
        return true;
      }});
      
      renderTable(filtered);
    }}
    
    renderTable(rawData);
  </script>
</body>
</html>
"""
    with open(output_path, 'w', encoding='utf-8') as f:
        f.write(html_content)
    print(f"[OK] 离线交互式 HTML 看板已生成: {output_path}")
