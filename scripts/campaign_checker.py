#!/usr/bin/env python3
# -*- coding: utf-8 -*-
import calendar
from datetime import datetime, date
from typing import List, Dict, Any, Optional
import pandas as pd

DEFAULT_CAMPAIGNS = [
    {
        "id": "c1",
        "name": "月度全渠道净支出达标 (6000元)",
        "target_amount": 6000.0,
        "channels": [],
        "keywords": [],
        "cards": [],
        "min_single_amount": 0.0,
        "description": "当月全卡全渠道累计净支出（总消费 - 商户退款）达到 6,000 元"
    },
    {
        "id": "c2",
        "name": "美团相关消费达标 (2000元)",
        "target_amount": 2000.0,
        "channels": ["美团"],
        "keywords": ["美团", "大众点评", "三快", "汉海"],
        "cards": [],
        "min_single_amount": 0.0,
        "description": "美团相关消费（含美团支付、大众点评及微信美团小程序）累计净支出达到 2,000 元"
    },
    {
        "id": "c3",
        "name": "抖音相关消费监控",
        "target_amount": 300.0,
        "channels": ["内容电商/直播"],
        "keywords": ["抖音", "字节"],
        "cards": [],
        "min_single_amount": 0.0,
        "description": "抖音支付及字节系平台累计净消费支出监控"
    }
]

def evaluate_campaigns(
    df: pd.DataFrame,
    campaigns_config: Optional[List[Dict[str, Any]]],
    year: int,
    month: int,
    current_end_day: int
) -> List[Dict[str, Any]]:
    rules = campaigns_config if (campaigns_config and len(campaigns_config) > 0) else DEFAULT_CAMPAIGNS
    
    _, days_in_month = calendar.monthrange(year, month)
    remaining_days = max(0, days_in_month - current_end_day)
    
    results = []
    for rule in rules:
        name = rule.get('name', '未命名活动')
        target = float(rule.get('target_amount', 0.0))
        channels = rule.get('channels', [])
        keywords = rule.get('keywords', [])
        cards = rule.get('cards', [])
        min_amt = float(rule.get('min_single_amount', 0.0))
        desc = rule.get('description', '')
        
        sub_df = df.copy() if not df.empty else pd.DataFrame()
        if not sub_df.empty:
            if channels and keywords:
                kw_regex = '|'.join(keywords)
                mask_ch = sub_df['channel'].isin(channels)
                mask_kw = sub_df['description'].str.contains(kw_regex, case=False, na=False)
                sub_df = sub_df[mask_ch | mask_kw]
            elif channels:
                sub_df = sub_df[sub_df['channel'].isin(channels)]
            elif keywords:
                kw_regex = '|'.join(keywords)
                sub_df = sub_df[sub_df['description'].str.contains(kw_regex, case=False, na=False)]
                
            if cards:
                sub_df = sub_df[sub_df['card_no'].isin([str(c) for c in cards])]
            if min_amt > 0:
                sub_df = sub_df[
                    (sub_df['trans_type'] == '商户退款') | 
                    ((sub_df['trans_type'] == '支出') & (sub_df['amount'] >= min_amt))
                ]
                
            exp = sub_df[sub_df['trans_type'] == '支出']['amount'].sum()
            ref = sub_df[sub_df['trans_type'] == '商户退款']['amount'].sum()
            current_spend = round(exp - ref, 2)
            tx_count = len(sub_df)
        else:
            current_spend = 0.0
            tx_count = 0
            
        is_reached = bool(current_spend >= target)
        progress_pct = round((current_spend / target * 100.0), 1) if target > 0 else 100.0
        gap = max(0.0, round(target - current_spend, 2))
        
        if is_reached:
            status_text = "已达标 🎉"
            status_code = "ACHIEVED"
            suggested_daily = 0.0
        else:
            status_text = f"还差 ¥{gap:,.2f}"
            status_code = "IN_PROGRESS"
            suggested_daily = round(gap / remaining_days, 2) if remaining_days > 0 else gap
            
        cond_display = []
        if keywords:
            cond_display.append(f"含关键词[{'/'.join(keywords)}]")
        elif channels:
            cond_display.append('/'.join(channels))
        else:
            cond_display.append("全渠道")
            
        results.append({
            "name": name,
            "description": desc,
            "target_amount": target,
            "current_spend": float(current_spend),
            "progress_pct": float(min(progress_pct, 999.0)),
            "gap_amount": float(gap),
            "is_reached": bool(is_reached),
            "status_text": status_text,
            "status_code": status_code,
            "tx_count": int(tx_count),
            "channels": cond_display,
            "remaining_days": int(remaining_days),
            "suggested_daily": float(suggested_daily)
        })
        
    return results
