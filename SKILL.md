---
name: citic-card-tracker
description: 自动导出中信银行信用卡已出与未出账单 Excel、智能合并去重、汇总当月支出与退款明细，并按渠道监控信用卡活动消费达标进度的系统工具。适用于中信信用卡消费对账、活动刷卡达标监控及云电脑跨设备自动化。
---

# 中信信用卡账单助手 (CITIC Card Tracker)

针对中信银行信用卡网上银行（`https://e.creditcard.ecitic.com/citiccard/ebank-ocp/ebankpc/bill.html`）的端到端自动化工具，解决以下核心痛点：
1. **全自动导出**: 自动导航网银页面，智能处理登录态，点击下载当月【已出账单】与【未出账单】Excel。
2. **账单合并与清洗**: 自动处理两份账单的表头差异、重叠交易去重、日期统一、卡号识别。
3. **收支对冲与渠道分类**: 自动区分“支出”与“退款/还款”，智能识别微信支付、支付宝、云闪付、美团、京东、拼多多、线下POS等支付渠道。
4. **月度区间汇总 (1号~N号)**: 例如在每月 20 号导出，精准汇总 1 号至 20 号所有净支出（总支出 - 总退款）。
5. **活动消费达标监控系统**: 针对参加的返现、刷卡金、多倍积分等银行活动，实时测算全渠道及特定渠道消费完成度、差额及剩余天数建议日均消费。
6. **双模成果交付**: 自动生成专业格式的 Excel 账簿系统（含公式、条件高亮、5个分类工作表）与离线自包含的现代 HTML 交互看板。

---

## 快速使用指令

所有脚本均位于本 skill 的 `scripts/` 目录下。

### 1. 全自动一键运行 (自动打开浏览器并导出)
```bash
# 统计当月 1 号至今天（例如 20 号）的数据
python scripts/run.py --month 2026-10 --end-day 20
```
> 若尚未登录，终端会提示用户使用【动卡空间 App】扫码，扫码成功后脚本自动完成后续所有操作。

### 2. 挂载至已有 Chrome 浏览器 (推荐云电脑使用)
若云电脑已打开并登录了中信网银：
```bash
# 先启动带调试端口的 Chrome: chrome.exe --remote-debugging-port=9222
python scripts/run.py --chrome-port 9222 --end-day 20
```

### 3. 离线文件直接导入模式 (无需浏览器驱动)
若已手动下载好已出账单与未出账单 Excel：
```bash
python scripts/run.py --no-export --import-dir ./downloads --end-day 20
```
或者明确指定文件：
```bash
python scripts/run.py --no-export --settled ./已出账单明细.xls --unsettled ./未出账单明细.xls --end-day 20
```

---

## 输出产物说明

运行后将在指定输出目录（默认 `./outputs/`）生成两套交付文件：

1. **`中信信用卡当月明细与汇总系统_YYYYMM_DD日.xlsx`**:
   - `📊 汇总看板`: 当月总支出、总退款、净支出、笔数指标卡片，各渠道支出净额与占比汇总表，各卡片明细汇总。
   - `🎯 活动达标监控`: 实时展示各活动达标线、已累计消费、完成百分比、距达标差额、达标状态高亮、剩余天数建议日均消费。
   - `📑 全部合并交易明细`: 统一格式整合后的明细流水，带支付渠道标签与交易类型标注。
   - `🔴 仅支出明细`: 按消费金额降序排列的纯支出明细。
   - `🟢 仅退款明细`: 退款、对冲与还款明细，便于与商户退款对账。
2. **`citic_tracker_dashboard.html`**:
   - 自包含离线网页仪表盘，任何电脑双击即开。
   - 包含高颜值 KPI 卡片、活动达标动态进度条、渠道占比图、支持实时模糊搜索与按渠道/类型筛选的明细表格。

---

## 架构与核心模块

- **`scripts/export_bills.py`**: 基于 Selenium 的网银页面自动化导出器，支持 Chrome/Edge、调试端口挂载、下载流监听。
- **`scripts/merge_and_analyze.py`**: 核心数据引擎，多引擎（HTML table/BIFF8/openpyxl）表格解析、渠道正则分类、日期过滤与去重。
- **`scripts/campaign_checker.py`**: 活动监控引擎，根据规则测算达标情况与差额。
- **`scripts/generate_system.py`**: 生成 Excel 账簿系统与 HTML 离线看板。
- **`scripts/run.py`**: 统一调度主入口 CLI。

---

## 详细参考文档

- **[账单格式与渠道识别规范](references/citic_bill_structure.md)**: 网银字段映射、HTML 伪 Excel 兼容、退款判定与渠道识别正则库。
- **[活动达标规则配置指南](references/campaign_config_guide.md)**: 如何自定义添加活动、门槛金额、渠道限定与单笔金额限制。
- **[云电脑与跨设备使用指南](references/cloud_pc_setup.md)**: 解决在远程工作站、无影/腾讯云桌面等机器上的运行与扫码登录方案。
