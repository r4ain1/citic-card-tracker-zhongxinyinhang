#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
中信银行信用卡账单合并、清洗与多维汇总引擎
支持自动识别已出账单与未出账单、智能去重、渠道分类、正负退款对冲与日期区间过滤。
"""

import os
import re
import sys
from datetime import datetime, date
from typing import List, Dict, Any, Tuple, Optional
import pandas as pd


# 智能支付渠道正则匹配规则库
CHANNEL_RULES = [
    (r'(财付通|微信支付|微信|WX_|\bWX\b|TENPAY)', '微信支付'),
    (r'(支付宝|ALIPAY|蚂蚁金服|口碑)', '支付宝'),
    (r'(云闪付|银联在线|银联无卡|UPAY|银联二维码|银联闪付)', '云闪付'),
    (r'(美团|三快在线|美团支付|汉海信息)', '美团'),
    (r'(京东|京东白条|网银在线|广州晶东|宿迁联盛|京东支付)', '京东'),
    (r'(拼多多|寻梦信息|拼多多支付)', '拼多多'),
    (r'(动卡空间|中信信用卡|中信商城|动卡支付|积分商城)', '动卡空间'),
    (r'(滴滴|出行|高德|曹操|花小猪)', '出行交通'),
    (r'(瑞幸|星巴克|麦当劳|肯德基|喜茶|奈雪|海底捞)', '餐饮美食'),
    (r'(抖音|快手|字节跳动|B站|哔哩哔哩)', '内容电商/直播'),
    (r'(中石化|中石油|加油站|中国石化|中国石油)', '汽车/加油'),
    (r'(POS|刷卡|柜面|线下|超市|商场|便利店|百货)', '线下POS/刷卡'),
    (r'(PAYPAL|AMAZON|VISA|MASTERCARD|JCB|境外)', '境外/海淘'),
]


def identify_channel(desc: str) -> str:
    """根据交易描述智能识别支付渠道或商户类别"""
    if not desc or not isinstance(desc, str):
        return '其它'
    desc_upper = desc.upper()
    for pattern, channel in CHANNEL_RULES:
        if re.search(pattern, desc_upper):
            return channel
    return '其它/常规刷卡'


def parse_amount(val: Any) -> float:
    """解析金额字段，处理包含币种符号、千分位、CR标识、负号的各种情况"""
    if pd.isna(val) or val is None:
        return 0.0
    if isinstance(val, (int, float)):
        return float(val)
    s = str(val).strip()
    if not s:
        return 0.0
    
    # 移除常见币种标记
    s = re.sub(r'[￥¥$€\s]|CNY|RMB|USD', '', s, flags=re.I)
    # 移除千分位逗号
    s = s.replace(',', '')
    
    is_negative = False
    # 括号表示负数: (100.00)
    if s.startswith('(') and s.endswith(')'):
        is_negative = True
        s = s[1:-1].strip()
    elif s.endswith('CR') or s.endswith('cr'):  # 贷记记账标记
        is_negative = True
        s = s[:-2].strip()
    elif '-' in s:
        is_negative = True
        s = s.replace('-', '').strip()
    
    try:
        amt = float(s)
        return -amt if is_negative else amt
    except ValueError:
        # 尝试提取第一个有效浮点数
        m = re.search(r'[-+]?\d*\.?\d+', s)
        if m:
            try:
                amt = float(m.group(0))
                return -amt if is_negative else amt
            except ValueError:
                return 0.0
        return 0.0


def parse_date_str(val: Any) -> Optional[str]:
    """统一将日期格式化为 YYYY-MM-DD 字符串"""
    if pd.isna(val) or val is None:
        return None
    s = str(val).strip()
    if not s or s.lower() in ('nan', 'none', 'nat'):
        return None
    
    # 匹配常见格式: 20261015, 2026-10-15, 2026/10/15, 2026.10.15, 10/15/2026
    s_clean = re.sub(r'[^\d]', '', s)
    if len(s_clean) == 8:
        # YYYYMMDD
        y = s_clean[:4]
        m = s_clean[4:6]
        d = s_clean[6:8]
        if 1 <= int(m) <= 12 and 1 <= int(d) <= 31:
            return f"{y}-{m}-{d}"
    
    # 尝试多种标准日期格式
    for fmt in ('%Y-%m-%d', '%Y/%m/%d', '%Y.%m.%d', '%Y-%m-%d %H:%M:%S', '%m/%d/%Y'):
        try:
            dt = datetime.strptime(s, fmt)
            return dt.strftime('%Y-%m-%d')
        except ValueError:
            pass
    
    # 尝试 pandas to_datetime
    try:
        dt = pd.to_datetime(s)
        if pd.notna(dt):
            return dt.strftime('%Y-%m-%d')
    except Exception:
        pass
    
    return None


def read_citic_excel(file_path: str) -> pd.DataFrame:
    """
    读取中信网银导出的 XLS / XLSX 文件。
    由于银行导出的 .xls 常常是 HTML Table 或标准 BIFF8/openpyxl，这里做自适应多引擎解析。
    """
    if not os.path.exists(file_path):
        raise FileNotFoundError(f"文件未找到: {file_path}")
    
    df = None
    errors = []
    
    # 尝试 1: 用 pandas read_html (很多网银下载的 .xls 实际上是带有 HTML <table> 的文本)
    try:
        tables = pd.read_html(file_path, encoding='utf-8')
        if tables and len(tables) > 0:
            for t in tables:
                # 寻找包含交易描述、交易日或金额的表格
                t_str = t.to_string()
                if any(k in t_str for k in ['交易', '金额', '日期', '描述', '卡号']):
                    df = t
                    break
            if df is None and len(tables) > 0:
                df = tables[0]
    except Exception as e:
        errors.append(f"read_html(utf-8): {e}")
        try:
            tables = pd.read_html(file_path, encoding='gbk')
            if tables and len(tables) > 0:
                df = tables[0]
        except Exception as e2:
            errors.append(f"read_html(gbk): {e2}")

    # 尝试 2: 用 openpyxl / xlrd 引擎
    if df is None:
        for engine in ['xlrd', 'openpyxl', None]:
            try:
                df = pd.read_excel(file_path, engine=engine)
                break
            except Exception as e:
                errors.append(f"read_excel({engine}): {e}")

    if df is None or df.empty:
        raise ValueError(f"无法读取表格 {file_path}，已尝试所有解析引擎。错误列表: {errors}")

    # 寻找真实表头行：扫描前 10 行，找到包含核心列名的行
    header_row_idx = None
    for i in range(min(15, len(df))):
        row_values = [str(x) for x in df.iloc[i].values if pd.notna(x)]
        row_str = " ".join(row_values)
        if any(k in row_str for k in ['交易日', '交易日期', '交易描述', '记账日', '记账金额', '交易金额']):
            header_row_idx = i
            break

    if header_row_idx is not None:
        df.columns = [str(c).strip() for c in df.iloc[header_row_idx].values]
        df = df.iloc[header_row_idx + 1:].reset_index(drop=True)

    # 清洗空行和多余说明行
    df = df.dropna(how='all')
    return df


def normalize_citic_dataframe(df: pd.DataFrame, source_type: str) -> pd.DataFrame:
    """
    将中信账单字段标准化为统一的规范结构:
    [txn_date, post_date, card_no, description, amount, currency, channel, trans_type, source]
    """
    col_map = {}
    for col in df.columns:
        c_str = str(col).strip()
        if any(k in c_str for k in ['交易日', '交易日期', 'Txn Date']):
            col_map['txn_date'] = col
        elif any(k in c_str for k in ['记账日', '入账日', '结算日', 'Post Date']):
            col_map['post_date'] = col
        elif any(k in c_str for k in ['卡号', '卡号末四位', '卡号后四位', '卡末四位', 'Card']):
            col_map['card_no'] = col
        elif any(k in c_str for k in ['交易描述', '交易说明', '商户名称', '交易摘要', 'Description']):
            col_map['description'] = col
        elif any(k in c_str for k in ['记账币种/金额', '记账金额', '结算金额', '入账金额', 'Post Amt']):
            col_map['amount'] = col
        elif 'amount' not in col_map and any(k in c_str for k in ['交易币种/金额', '交易金额', '金额', 'Txn Amt']):
            col_map['amount'] = col
        elif any(k in c_str for k in ['币种', '货币', 'Currency']):
            col_map['currency'] = col

    # 如果列名未直接匹配，使用位置推测
    if 'description' not in col_map and len(df.columns) >= 3:
        # 交易描述通常是长文本列
        for col in df.columns:
            sample_vals = df[col].dropna().astype(str).tolist()[:5]
            if any(len(v) > 4 and not v.replace('.', '').isdigit() for v in sample_vals):
                col_map['description'] = col
                break

    standard_rows = []
    for _, row in df.iterrows():
        desc = str(row[col_map['description']]).strip() if 'description' in col_map else ''
        if not desc or desc in ['交易描述', '总计', '合计', '小计', '温馨提示']:
            continue
        
        # 金额解析
        raw_amt = row[col_map['amount']] if 'amount' in col_map else 0.0
        amt = parse_amount(raw_amt)
        
        # 日期解析
        raw_txn_date = row[col_map['txn_date']] if 'txn_date' in col_map else None
        raw_post_date = row[col_map['post_date']] if 'post_date' in col_map else None
        
        txn_date = parse_date_str(raw_txn_date)
        post_date = parse_date_str(raw_post_date) or txn_date
        
        if not txn_date and post_date:
            txn_date = post_date
        if not txn_date:
            # 若无日期，可能是说明文字行，过滤
            continue
            
        # 卡号末四位提取
        raw_card = str(row[col_map['card_no']]).strip() if 'card_no' in col_map else ''
        card_m = re.search(r'(\d{4})', raw_card)
        card_no = card_m.group(1) if card_m else (raw_card[-4:] if len(raw_card) >= 4 else '默认卡')
        
        # 币种
        raw_curr = str(row[col_map['currency']]).strip() if 'currency' in col_map else 'CNY'
        currency = 'CNY'
        if any(c in str(raw_amt).upper() or c in desc.upper() for c in ['USD', 'EUR', 'JPY', 'HKD']):
            for c in ['USD', 'EUR', 'JPY', 'HKD']:
                if c in str(raw_amt).upper() or c in desc.upper():
                    currency = c
                    break
        elif raw_curr and raw_curr.upper() in ['USD', 'EUR', 'JPY', 'HKD', 'GBP']:
            currency = raw_curr.upper()
            
        # 判断交易类型 (支出 vs 商户退款 vs 信用卡还款)
        # 区分用户主动还款与商户消费退款
        is_repayment = any(k in desc for k in ['还款', '转账存入', '自动还款', '存现', '网银互联还款'])
        is_refund_kw = any(k in desc for k in ['退款', '退货', '冲正', '撤销', '退回', '返还', '贷记'])
        
        if is_repayment:
            trans_type = '信用卡还款'
            abs_amt = abs(amt)
        elif amt < 0 or is_refund_kw:
            trans_type = '商户退款'
            abs_amt = abs(amt)
        else:
            trans_type = '支出'
            abs_amt = abs(amt)
            
        channel = identify_channel(desc)
        
        standard_rows.append({
            'txn_date': txn_date,
            'post_date': post_date,
            'card_no': card_no,
            'description': desc,
            'channel': channel,
            'trans_type': trans_type,
            'amount': abs_amt,          # 正数绝对值，便于统计
            'signed_amount': -abs_amt if trans_type in ('商户退款', '信用卡还款') else abs_amt, # 带符号金额
            'currency': currency,
            'source': source_type       # '已出账单' 或 '未出账单'
        })
        
    return pd.DataFrame(standard_rows)


def merge_and_filter_bills(
    settled_df: Optional[pd.DataFrame],
    unsettled_df: Optional[pd.DataFrame],
    start_date: str,
    end_date: str
) -> pd.DataFrame:
    """
    合并已出账单与未出账单、去重、并按指定日期区间 [start_date, end_date] 进行严格过滤
    """
    dfs = []
    if settled_df is not None and not settled_df.empty:
        dfs.append(settled_df)
    if unsettled_df is not None and not unsettled_df.empty:
        dfs.append(unsettled_df)
        
    if not dfs:
        return pd.DataFrame()
        
    combined = pd.concat(dfs, ignore_index=True)
    
    # 去重：若同一笔交易（相同交易日、相同卡号、相同金额、相同描述）同时出现在已出和未出中，保留已出账单
    # 优先排已出账单在前面，然后 drop_duplicates keep='first'
    combined['source_priority'] = combined['source'].apply(lambda x: 0 if '已出' in x else 1)
    combined = combined.sort_values(by=['txn_date', 'source_priority']).reset_index(drop=True)
    combined = combined.drop_duplicates(
        subset=['txn_date', 'card_no', 'amount', 'description', 'trans_type'],
        keep='first'
    ).drop(columns=['source_priority'])
    
    # 按交易日期过滤区间
    filtered = combined[
        (combined['txn_date'] >= start_date) & 
        (combined['txn_date'] <= end_date)
    ].sort_values(by='txn_date', ascending=True).reset_index(drop=True)
    
    return filtered


def calculate_summary_stats(df: pd.DataFrame) -> Dict[str, Any]:
    """
    计算全局统计核心指标：
    - 总支出
    - 总退款
    - 净支出 (总支出 - 总退款)
    - 笔数
    - 渠道汇总
    - 卡号汇总
    - 每日消费走势
    """
    if df.empty:
        return {
            'total_expense': 0.0,
            'total_refund': 0.0,
            'net_spend': 0.0,
            'expense_count': 0,
            'refund_count': 0,
            'total_count': 0,
            'channel_summary': [],
            'card_summary': [],
            'daily_summary': []
        }
        
    expenses = df[df['trans_type'] == '支出']
    refunds = df[df['trans_type'] == '商户退款']
    repayments = df[df['trans_type'] == '信用卡还款']
    
    total_expense = float(round(expenses['amount'].sum(), 2))
    total_refund = float(round(refunds['amount'].sum(), 2))
    total_repayment = float(round(repayments['amount'].sum(), 2))
    net_spend = float(round(total_expense - total_refund, 2))
    
    # 渠道汇总
    channel_groups = []
    for channel, group in df.groupby('channel'):
        exp = group[group['trans_type'] == '支出']['amount'].sum()
        ref = group[group['trans_type'] == '商户退款']['amount'].sum()
        net = exp - ref
        cnt = len(group)
        pct = (net / net_spend * 100.0) if net_spend > 0 else 0.0
        channel_groups.append({
            'channel': channel,
            'expense': round(exp, 2),
            'refund': round(ref, 2),
            'net': round(net, 2),
            'count': cnt,
            'percentage': round(pct, 1)
        })
    channel_groups.sort(key=lambda x: x['net'], reverse=True)
    
    # 卡片汇总
    card_groups = []
    for card, group in df.groupby('card_no'):
        exp = group[group['trans_type'] == '支出']['amount'].sum()
        ref = group[group['trans_type'] == '商户退款']['amount'].sum()
        net = exp - ref
        card_groups.append({
            'card_no': f"尾号{card}",
            'expense': round(exp, 2),
            'refund': round(ref, 2),
            'net': round(net, 2),
            'count': len(group)
        })
    card_groups.sort(key=lambda x: x['net'], reverse=True)
    
    # 每日明细走势
    daily_groups = []
    for dt, group in df.groupby('txn_date'):
        exp = group[group['trans_type'] == '支出']['amount'].sum()
        ref = group[group['trans_type'] == '商户退款']['amount'].sum()
        daily_groups.append({
            'date': dt,
            'expense': round(exp, 2),
            'refund': round(ref, 2),
            'net': round(exp - ref, 2),
            'count': len(group)
        })
    daily_groups.sort(key=lambda x: x['date'])
    
    return {
        'total_expense': total_expense,
        'total_refund': total_refund,
        'total_repayment': total_repayment,
        'net_spend': net_spend,
        'expense_count': len(expenses),
        'refund_count': len(refunds),
        'repayment_count': len(repayments),
        'total_count': len(df),
        'channel_summary': channel_groups,
        'card_summary': card_groups,
        'daily_summary': daily_groups
    }
