# 中信银行信用卡账单自动化与活动消费达标监控系统 (CITIC Bank Card Tracker)

专为中信银行信用卡网银打造的账单自动化工具。一键自动导出已出与未出账单 Excel、智能合并去重、汇总当月净支出（支出与退款对冲），并实时监控各渠道活动消费达标进度。

## 1. 核心特性
- **全自动网银下载**: 自动接管 Chrome 打开中信网银，智能等待扫码登录，一键导出【已出账单】与【未出账单】。
- **智能数据清洗**: 兼容 HTML 伪 Excel 与标准 BIFF8 二进制表，自动清洗多余行，识别正负退款对冲。
- **精准渠道归类**: 正则识别微信支付、支付宝、云闪付、美团、京东、拼多多、线下POS等渠道。
- **活动达标测算**: 支持配置总消费与特定渠道达标要求，自动计算完成度、还差金额及剩余天数建议日均刷卡额。
- **双模交付**: 输出包含 5 个分页的专业 Excel 账簿系统与自包含离线交互 HTML 仪表盘。

## 2. 快速开始

```bash
# 安装基础依赖
pip install pandas openpyxl selenium pyyaml

# 运行统计当月 1 号至今天（例如 20 号）的数据
python scripts/run.py --month 2026-10 --end-day 20
```

## 3. 云电脑运行提示

- **直接运行**: 终端执行后，在弹出的 Chrome 窗口中用【动卡空间 App】扫码即可，后续全程自动。
- **复用已有浏览器**: 先启动 `chrome.exe --remote-debugging-port=9222` 并登录，再执行 `python scripts/run.py --chrome-port 9222`。
- **离线导入**: 若已手动下载账单，执行 `python scripts/run.py --no-export --import-dir ./downloads`。

详细文档请查阅 `references/` 目录。
