#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
中信信用卡账单全自动处理主入口
一键执行：网银下载 -> 账单合并 -> 智能清洗去重 -> 区间汇总 (1号~N号) -> 活动达标计算 -> 双模报表输出
"""

import os
import sys
import json
import argparse
from datetime import datetime, date
import yaml

from merge_and_analyze import (
    read_citic_excel,
    normalize_citic_dataframe,
    merge_and_filter_bills,
    calculate_summary_stats
)
from campaign_checker import evaluate_campaigns, DEFAULT_CAMPAIGNS
from generate_system import create_excel_system, create_html_dashboard


def load_config(config_path: str) -> dict:
    if os.path.exists(config_path):
        try:
            with open(config_path, 'r', encoding='utf-8') as f:
                if config_path.endswith(('.yaml', '.yml')):
                    return yaml.safe_load(f) or {}
                elif config_path.endswith('.json'):
                    return json.load(f) or {}
        except Exception as e:
            print(f"[警告] 读取配置文件失败: {e}，将采用默认活动规则。")
    return {}


def main():
    parser = argparse.ArgumentParser(description="中信银行信用卡账单自动化汇总与活动达标系统")
    parser.add_argument("--month", default=None, help="目标月份 (格式: YYYY-MM，如 2026-10，默认当前月份)")
    parser.add_argument("--start-day", type=int, default=1, help="统计起始日 (默认: 1)")
    parser.add_argument("--end-day", type=int, default=None, help="统计截止日 (默认: 当天日期，如每个月20号导出则默认20)")
    
    # 文件来源控制
    parser.add_argument("--no-export", action="store_true", help="跳过浏览器导出，使用本地已有文件")
    parser.add_argument("--import-dir", default=None, help="指定本地已下载账单的目录")
    parser.add_argument("--settled", default=None, help="已出账单文件路径")
    parser.add_argument("--unsettled", default=None, help="未出账单文件路径")
    
    # 浏览器控制
    parser.add_argument("--chrome-port", type=int, default=None, help="复用已有 Chrome 调试端口 (如 9222)")
    parser.add_argument("--user-data-dir", default="./chrome_profile", help="持久化浏览器目录")
    parser.add_argument("--headless", action="store_true", help="无头模式")
    
    # 输出与配置
    parser.add_argument("--output-dir", default="./outputs", help="系统文件生成输出目录")
    parser.add_argument("--config", default="./config.yaml", help="活动规则与参数配置文件")
    
    args = parser.parse_args()
    
    now = datetime.now()
    # 确定目标年月与截止日
    if args.month:
        parts = args.month.split('-')
        target_year = int(parts[0])
        target_month = int(parts[1])
    else:
        target_year = now.year
        target_month = now.month
        
    month_str = f"{target_year:04d}-{target_month:02d}"
    start_day = args.start_day
    end_day = args.end_day if args.end_day is not None else now.day
    
    start_date = f"{month_str}-{start_day:02d}"
    end_date = f"{month_str}-{end_day:02d}"
    
    print("=" * 65)
    print(f"💳 中信银行信用卡账单汇总与活动达标系统")
    print(f"📅 统计区间: {start_date} 至 {end_date} (当月 1 号到 {end_day} 号)")
    print("=" * 65)
    
    os.makedirs(args.output_dir, exist_ok=True)
    settled_path = args.settled
    unsettled_path = args.unsettled
    
    # 1. 查找或导出文件
    if not settled_path or not unsettled_path:
        if args.import_dir and os.path.exists(args.import_dir):
            print(f"[*] 正在从指定目录扫描已出和未出账单: {args.import_dir}")
            files = os.listdir(args.import_dir)
            for f in sorted(files, reverse=True):
                f_lower = f.lower()
                full = os.path.join(args.import_dir, f)
                if not settled_path and any(k in f for k in ['已出', 'settled']) and f_lower.endswith(('.xls', '.xlsx')):
                    settled_path = full
                elif not unsettled_path and any(k in f for k in ['未出', 'unsettled']) and f_lower.endswith(('.xls', '.xlsx')):
                    unsettled_path = full
                    
        # 如果仍然没有且未禁用导出，则启动浏览器自动下载
        if (not settled_path or not unsettled_path) and not args.no_export:
            from export_bills import export_citic_bills
            download_dir = os.path.join(args.output_dir, "downloads")
            print("[*] 正在启动网银自动下载流程...")
            settled_path, unsettled_path = export_citic_bills(
                download_dir=download_dir,
                target_month_str=month_str,
                user_data_dir=args.user_data_dir,
                remote_port=args.chrome_port,
                headless=args.headless
            )
            
    if not settled_path and not unsettled_path:
        print("[错误] 未能找到已出账单或未出账单文件，请提供文件路径或开启自动导出。")
        sys.exit(1)
        
    print(f"\n[已加载源文件]")
    print(f" - 已出账单: {settled_path if settled_path else '未提供'}")
    print(f" - 未出账单: {unsettled_path if unsettled_path else '未提供'}")
    
    # 2. 读取与标准化
    settled_df = None
    if settled_path and os.path.exists(settled_path):
        raw_s = read_citic_excel(settled_path)
        settled_df = normalize_citic_dataframe(raw_s, source_type="已出账单")
        print(f"[OK] 解析已出账单完成，有效交易行数: {len(settled_df)}")
        
    unsettled_df = None
    if unsettled_path and os.path.exists(unsettled_path):
        raw_u = read_citic_excel(unsettled_path)
        unsettled_df = normalize_citic_dataframe(raw_u, source_type="未出账单")
        print(f"[OK] 解析未出账单完成，有效交易行数: {len(unsettled_df)}")
        
    # 3. 合并、去重与日期过滤
    df_combined = merge_and_filter_bills(
        settled_df, unsettled_df, start_date=start_date, end_date=end_date
    )
    print(f"[OK] 账单合并与去重完毕，{start_date} ~ {end_date} 区间内共有 {len(df_combined)} 笔交易。")
    
    # 4. 指标统计
    stats = calculate_summary_stats(df_combined)
    
    # 5. 读取配置并执行活动监控计算
    cfg = load_config(args.config)
    campaigns_cfg = cfg.get('campaigns', DEFAULT_CAMPAIGNS)
    cmp_results = evaluate_campaigns(
        df_combined, campaigns_cfg, year=target_year, month=target_month, current_end_day=end_day
    )
    
    # 6. 生成输出系统文件
    excel_filename = f"中信信用卡当月明细与汇总系统_{target_year}{target_month:02d}_{end_day:02d}日.xlsx"
    excel_out = os.path.join(args.output_dir, excel_filename)
    html_out = os.path.join(args.output_dir, "citic_tracker_dashboard.html")
    
    create_excel_system(excel_out, df_combined, stats, cmp_results, start_date, end_date)
    create_html_dashboard(html_out, df_combined, stats, cmp_results, start_date, end_date)
    
    # 7. 控制台输出核心摘要
    print("\n" + "=" * 65)
    print(f"📊 【当月收支汇总】 (统计至 {end_date})")
    print(f" · 当月总支出:   ¥ {stats['total_expense']:>10,.2f}  (共 {stats['expense_count']} 笔消费)")
    print(f" · 当月总退款:   ¥ {stats['total_refund']:>10,.2f}  (共 {stats['refund_count']} 笔退款)")
    print(f" · 当月净支出:   ¥ {stats['net_spend']:>10,.2f}  (实际已计入信用卡支出)")
    print("-" * 65)
    print(f"💳 【主要渠道消费分布】:")
    for ch in stats['channel_summary'][:5]:
        print(f" · {ch['channel']:<10} 净额: ¥ {ch['net']:>8,.2f} | 占比: {ch['percentage']:>5.1f}% | {ch['count']}笔")
    print("-" * 65)
    print(f"🎯 【活动达标监控情况】:")
    for cmp in cmp_results:
        status_mark = "✅ 已达标" if cmp['is_reached'] else f"⏳ 差 ¥{cmp['gap_amount']:,.2f}"
        ch_tag = "全渠道" if cmp['channels'] == ["全渠道"] else ",".join(cmp['channels'])
        daily_advice = f"(建议日均 ¥{cmp['suggested_daily']:,.2f})" if not cmp['is_reached'] and cmp['remaining_days'] > 0 else ""
        print(f" · {cmp['name']:<20} 进度: {cmp['progress_pct']:>5.1f}% [{cmp['current_spend']:,.0f}/{cmp['target_amount']:,.0f}元] | {status_mark} {daily_advice}")
    print("=" * 65)
    print(f"\n🎉 系统生成完毕，可在以下路径查看完整账簿与网页看板:")
    print(f" 1. Excel 系统: {os.path.abspath(excel_out)}")
    print(f" 2. 网页 仪表盘: {os.path.abspath(html_out)}")


if __name__ == '__main__':
    main()
