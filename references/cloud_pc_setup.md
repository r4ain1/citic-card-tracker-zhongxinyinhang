# 云电脑与跨设备使用指南 (Cloud PC & Multi-Device Setup)

在云电脑（如无影云电脑、腾讯云桌面、华为云工作空间、远程 Windows Server 或虚拟机）上运行本工具时，由于银行网银具备严格的风控与登录鉴权（短信验证码或动卡空间扫码），推荐以下三种运行方案：

---

## 方案 1: 图形界面直接运行（最推荐，最省心）

1. 在云电脑中打开终端（PowerShell 或 CMD）。
2. 运行主命令：
   ```bash
   python scripts/run.py --month 2026-10 --end-day 20
   ```
3. 脚本会自动启动云电脑上的 Chrome 浏览器并导航至中信网银账单页。
4. 屏幕上出现中信网银登录界面后，拿出手机打开【动卡空间 App】扫码完成登录。
5. 登录成功后，脚本会自动检测到账单表格加载，随即自动点击导出【已出账单】与【未出账单】，合并清洗并生成 Excel 与 HTML 看板！

> **登录态保留**: 脚本默认使用 `./chrome_profile` 保存浏览器缓存，只要在当前会话未过期前再次运行，即可直接免扫码导出。

---

## 方案 2: 复用已有 Chrome 调试端口（适合防封控 / 已登录环境）

如果你在云电脑上已经登录了中信网银，不想重新拉起新浏览器：

1. 先用调试端口模式启动云电脑的 Chrome：
   ```bash
   chrome.exe --remote-debugging-port=9222
   ```
2. 在该 Chrome 中正常访问 `https://e.creditcard.ecitic.com/citiccard/ebank-ocp/ebankpc/bill.html` 并登录。
3. 运行脚本，挂载到已有端口：
   ```bash
   python scripts/run.py --chrome-port 9222 --end-day 20
   ```
4. 脚本将直接在已有浏览器中接管页面，点击下载并汇总。

---

## 方案 3: 手动下载文件直接导入模式（无需任何自动化浏览器依赖）

如果在极端受限的云环境（无 GUI 界面或无 Chrome 驱动权限）：

1. 手动在网银页面分别点击【导出Excel】（已出账单）和【导出Excel】（未出账单）。
2. 将这两个文件放在任意目录（如 `./downloads`），或传给脚本参数：
   ```bash
   python scripts/run.py --no-export --import-dir ./downloads --end-day 20
   ```
   或者直接指定文件：
   ```bash
   python scripts/run.py --no-export --settled ./已出账单明细.xls --unsettled ./未出账单明细.xls --end-day 20
   ```
3. 系统会自动读取、合并、去重、统计渠道分布并输出达标报告。

---

## 依赖项安装 (零复杂配置)

```bash
pip install pandas openpyxl selenium pyyaml
```
如需在 Linux 云主机上运行图形 Chrome，需安装 `google-chrome-stable` 与 Xvfb 虚拟桌面，或者采用方案 3 离线解析。
