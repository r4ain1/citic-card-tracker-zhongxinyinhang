#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
中信银行信用卡账单自动导出驱动器 (Selenium / CDP)
自动登录检测、账单页面导航、月份选择、点击导出【已出账单】与【未出账单】并监控下载。
"""

import os
import sys
import time
import glob
import argparse
from datetime import datetime
from typing import Tuple, Optional

from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.options import Options as ChromeOptions
from selenium.webdriver.chrome.service import Service as ChromeService
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC


BILL_URL = "https://e.creditcard.ecitic.com/citiccard/ebank-ocp/ebankpc/bill.html"


def get_chrome_driver(
    download_dir: str,
    user_data_dir: Optional[str] = None,
    remote_port: Optional[int] = None,
    headless: bool = False
) -> webdriver.Chrome:
    """配置并启动 Chrome Driver，支持直连已有调试端口或拉起独立受控浏览器"""
    options = ChromeOptions()
    
    # 如果指定了已有的远程调试端口（如 9222）
    if remote_port:
        options.add_experimental_option("debuggerAddress", f"127.0.0.1:{remote_port}")
        driver = webdriver.Chrome(options=options)
        return driver
        
    # 自动下载配置
    prefs = {
        "download.default_directory": os.path.abspath(download_dir),
        "download.prompt_for_download": False,
        "download.directory_upgrade": True,
        "safebrowsing.enabled": True,
        "profile.default_content_settings.popups": 0
    }
    options.add_experimental_option("prefs", prefs)
    
    # 禁用自动化检测特征
    options.add_argument("--disable-blink-features=AutomationControlled")
    options.add_experimental_option("excludeSwitches", ["enable-automation"])
    options.add_experimental_option("useAutomationExtension", False)
    
    if user_data_dir:
        options.add_argument(f"--user-data-dir={os.path.abspath(user_data_dir)}")
    
    if headless:
        options.add_argument("--headless=new")
        options.add_argument("--disable-gpu")
        options.add_argument("--window-size=1440,900")
    else:
        options.add_argument("--start-maximized")
        
    driver = webdriver.Chrome(options=options)
    return driver


def wait_for_download(download_dir: str, timeout: int = 40) -> Optional[str]:
    """等待下载目录下出现新的 xls/xlsx/csv 文件并确保已完成（无 .crdownload 后缀）"""
    start_time = time.time()
    initial_files = set(os.listdir(download_dir))
    
    while time.time() - start_time < timeout:
        current_files = set(os.listdir(download_dir))
        new_files = current_files - initial_files
        
        # 检查是否有下载中的临时文件
        has_crdownload = any(f.endswith('.crdownload') or f.endswith('.tmp') for f in current_files)
        if has_crdownload:
            time.sleep(1)
            continue
            
        valid_new = [
            f for f in new_files 
            if f.endswith(('.xls', '.xlsx', '.csv')) and not f.startswith('~')
        ]
        if valid_new:
            # 获取最新文件
            full_path = os.path.join(download_dir, valid_new[0])
            # 确认文件大小大于0且写入完毕
            time.sleep(0.5)
            if os.path.getsize(full_path) > 0:
                return full_path
                
        time.sleep(1)
    return None


def export_citic_bills(
    download_dir: str,
    target_month_str: Optional[str] = None, # 格式如 "2026-10" 或 "10" 或 "10月"
    user_data_dir: Optional[str] = None,
    remote_port: Optional[int] = None,
    headless: bool = False
) -> Tuple[str, str]:
    """
    主导出流程：
    1. 打开账单页，智能等待用户扫码/登录
    2. 切换并下载【已出账单】
    3. 切换并下载【未出账单】
    4. 重命名并返回两个文件的完整路径
    """
    os.makedirs(download_dir, exist_ok=True)
    driver = get_chrome_driver(download_dir, user_data_dir, remote_port, headless)
    
    try:
        print(f"[1/4] 正在导航至中信网银账单页面: {BILL_URL}")
        driver.get(BILL_URL)
        
        # 登录检测：循环检测是否已进入包含账单菜单的页面
        print("[2/4] 正在检查网银登录状态...")
        max_login_wait = 180  # 最多等待 3 分钟让用户扫码
        logged_in = False
        for _ in range(max_login_wait // 2):
            cur_url = driver.current_url
            page_src = driver.page_source
            
            # 检查是否有已出账单标签或卡号列表
            if "bill.html" in cur_url and ("已出账单" in page_src or "sub_menu" in page_src):
                logged_in = True
                print("[OK] 已成功识别到登录会话与账单页面！")
                break
            
            print("[提示] 检测到尚未登录，请在打开的浏览器窗口中使用【动卡空间 App】扫码或密码登录...")
            time.sleep(2)
            
        if not logged_in:
            raise TimeoutError("等待登录超时（超过180秒），请重试并在浏览器中及时完成登录。")
            
        time.sleep(3) # 等待页面异步接口加载
        
        # ----------------------------------------------------
        # 步骤 1: 导出【已出账单】
        # ----------------------------------------------------
        print("[3/4] 正在准备导出【已出账单】...")
        # 确保处于已出账单 Tab
        tabs = driver.find_elements(By.CSS_SELECTOR, ".sub_menu span")
        if tabs and len(tabs) >= 1:
            tabs[0].click()
            time.sleep(1.5)
            
        # 如果指定了目标月份，点击月份导航
        if target_month_str:
            # 提取两位数月份，例如 "10"
            m_clean = target_month_str.replace("月", "").split("-")[-1].zfill(2)
            month_label = f"{m_clean}月"
            print(f"[*] 正在切换至指定月份: {month_label}")
            month_items = driver.find_elements(By.CSS_SELECTOR, ".bar-scroll li")
            found_month = False
            for item in month_items:
                t_text = item.text.strip()
                t_name = item.get_attribute("name") or ""
                if month_label in t_text or (target_month_str in t_name):
                    item.click()
                    found_month = True
                    print(f"[OK] 已选中月份: {t_text} ({t_name})")
                    time.sleep(2)
                    break
            if not found_month:
                print(f"[警告] 未在月份条中找到完全匹配项 {month_label}，使用默认当前选中月份。")

        # 等待 loading 遮罩消失
        try:
            WebDriverWait(driver, 15).until(
                EC.invisibility_of_element_located((By.ID, "loading"))
            )
        except Exception:
            pass
            
        # 点击已出账单下载按钮 #downSettleBill
        settle_btn = WebDriverWait(driver, 15).until(
            EC.element_to_be_clickable((By.ID, "downSettleBill"))
        )
        driver.execute_script("arguments[0].scrollIntoView(true);", settle_btn)
        time.sleep(0.5)
        settle_btn.click()
        print("[*] 已触发【已出账单明细】下载，正在等待文件落盘...")
        
        settled_file = wait_for_download(download_dir, timeout=30)
        if not settled_file:
            # 尝试通过 JS 触发导出兜底
            print("[*] 尝试通过页面内置 JS 接口重试触发下载...")
            driver.execute_script("if(typeof h === 'function') { h('settleBill'); }")
            settled_file = wait_for_download(download_dir, timeout=20)
            
        if not settled_file:
            raise RuntimeError("已出账单下载超时，未能检测到新文件。")
        print(f"[OK] 已出账单已成功下载: {settled_file}")
        
        # ----------------------------------------------------
        # 步骤 2: 导出【未出账单】
        # ----------------------------------------------------
        print("[4/4] 正在切换并准备导出【未出账单】...")
        tabs = driver.find_elements(By.CSS_SELECTOR, ".sub_menu span")
        if tabs and len(tabs) >= 2:
            tabs[1].click()
            time.sleep(2)
        else:
            driver.execute_script("document.querySelectorAll('.sub_menu span')[1].click();")
            time.sleep(2)
            
        # 等待 loading
        try:
            WebDriverWait(driver, 15).until(
                EC.invisibility_of_element_located((By.ID, "loading"))
            )
        except Exception:
            pass
            
        # 点击未出账单下载按钮 #downUnsettleBill
        unsettle_btn = WebDriverWait(driver, 15).until(
            EC.element_to_be_clickable((By.ID, "downUnsettleBill"))
        )
        driver.execute_script("arguments[0].scrollIntoView(true);", unsettle_btn)
        time.sleep(0.5)
        unsettle_btn.click()
        print("[*] 已触发【未出账单明细】下载，正在等待文件落盘...")
        
        unsettled_file = wait_for_download(download_dir, timeout=30)
        if not unsettled_file:
            # 尝试通过页面 JS 兜底
            driver.execute_script("if(typeof d === 'function') { d('unSettleBill'); } else if(typeof h === 'function') { h('unSettleBill'); }")
            unsettled_file = wait_for_download(download_dir, timeout=20)
            
        if not unsettled_file:
            raise RuntimeError("未出账单下载超时，未能检测到新文件。")
        print(f"[OK] 未出账单已成功下载: {unsettled_file}")
        
        # 规范化重命名
        stamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        final_settled = os.path.join(download_dir, f"citic_settled_{stamp}.xls")
        final_unsettled = os.path.join(download_dir, f"citic_unsettled_{stamp}.xls")
        
        os.replace(settled_file, final_settled)
        os.replace(unsettled_file, final_unsettled)
        
        print(f"[完成] 两个账单文件已安全落盘并规范命名:\n - {final_settled}\n - {final_unsettled}")
        return final_settled, final_unsettled
        
    finally:
        try:
            driver.quit()
        except Exception:
            pass


if __name__ == '__main__':
    parser = argparse.ArgumentParser(description="中信信用卡账单自动化下载工具")
    parser.add_argument("--download-dir", default="./downloads", help="账单文件保存目录")
    parser.add_argument("--month", default=None, help="目标月份 (如 2026-10 或 10)")
    parser.add_argument("--user-data-dir", default="./chrome_profile", help="Chrome 用户数据目录（保留登录状态）")
    parser.add_argument("--remote-port", type=int, default=None, help="连接已有 Chrome 调试端口 (如 9222)")
    parser.add_argument("--headless", action="store_true", help="无头模式运行 (需已登录)")
    args = parser.parse_args()
    
    export_citic_bills(
        download_dir=args.download_dir,
        target_month_str=args.month,
        user_data_dir=args.user_data_dir,
        remote_port=args.remote_port,
        headless=args.headless
    )
