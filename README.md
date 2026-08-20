# Finance-Accountant

面向中小微企业的 **opencode 技能包（Skill Pack）**：把"财务报表生成 + 税务申报"做成一套可复用的 AI 工作流，强调**数据真实性核对链**、**国家标准报表格式**与**纳税人身份自适应**。

## 特性

  **真实性核对链**：生成银行流水明细、发票费用、资产负债表、利润表、现金流量表、个税申报、季度申报、工商年报后，用 `verify_chain.py` 逐条验证，数字必须来自原始文件（XLSX/PDF），禁止编造。
  **纳税人身份自适应**：自动判定小规模/一般纳税人（开票税率、申报表样式、销售额等证据），按身份切换增值税计税方式（简易 3% / 一般 13%9%6%）、申报周期（季报/月报）与申报表样式；证据不足自动标注"待确认"。
  **发票票种智能识别**：`invoice_detect.py` 批量识别专票/普票/数电票/红字发票/铁路电子客票/支付回单，自动提取票种、购销方、价税合计、税额、发票号码，未确定票种强制人工确认（不按默认票种入账）。
  **国家标准表格格式**：资产负债表/利润表/现金流量表/增值税（两种身份）/企业所得税（A类）/个税扣缴/工商年报，栏次名称一字不差；没有数据的科目也列示填 0。
  **申报日历**：个税月报、增值税季报/月报、企税季报、工商年报到期的期限与渠道。
  **可扩展**：新增月份/季度/年度，按文档步骤复制期间配置即可，脚本自动核算。

## 结构

```
finance-accountant/
├── README.md                     # 项目说明
├── LICENSE                       # MIT 许可证
├── .gitignore
├── agents/
│   └── finance-accountant.md     # 财务与税务申报助理（agent）
├── skills/
│   ├── financial-check/          # 核对链验证 skill
│   │   ├── SKILL.md
│   │   └── scripts/
│   │       ├── verify_chain.py   # 银行流水核对链 + LibreOffice 公式重算
│   │       └── invoice_detect.py # 发票票种智能识别（专/普/数电/红字/回单）
│   └── tax-report-format/        # 税务申报格式与日历 skill
│       └── SKILL.md
├── docs/
│   ├── verification-chain.md     # 核对链规则
│   ├── taxpayer-type.md          # 纳税人身份判定（小规模 vs 一般纳税人）
│   ├── invoice-recognition.md    # 发票票种识别（专票/普票/数电票/红字）
│   ├── tax-calendar.md           # 申报日历
│   ├── formats.md                # 报表/申报表格式规范
│   └── pitfalls.md               # 已知陷阱
├── templates/
│   └── handover-template.md      # 交接文档模板（换对话/换人可复用）
└── examples/
    └── verify_chain.ps1          # 核对链调用示例
```

## 安装（opencode）

1. 克隆本仓库（或把 `skills/` 复制到你的 opencode 配置目录）。
2. 在 `opencode.jsonc` 声明 skill 路径：

   ```jsonc
   {
     "$schema": "https://opencode.ai/config.json",
     "skills": { "paths": ["<本仓库>/skills"] }
   }
   ```

3. 安装依赖：`pip install openpyxl`（如校验含公式的 xlsx 台账，另需 LibreOffice）。
4. 在对话中触发（例如"生成 8 月份财务报表"或"核对 XX 台账"），agent 会自动加载对应 skill 并执行核对。

## 使用示例

```powershell
$env:PYTHONIOENCODING = "utf-8"
python skills/financial-check/scripts/verify_chain.py `
  --bank "<财务数据根目录>/公账文件/7月份.xlsx" `
  --debit 1000.00 --credit 5000.00 --balance 4000.00 `
  --capital 4999.00 --verify 1.00 `
  --taxbook "<财务数据根目录>/个税申报/…年度台账.xlsx"
# 输出全部 PASS 才可交付
```

数字说明：示例中的数值仅为演示格式，**实际使用时按当期原始文件核对填入，禁止照抄**。

## 许可

[MIT](LICENSE)

> 免责声明：本项目提供工具与方法论，不构成税务/会计法律意见。实际申报请以税务机关官方要求为准，并与有资质的专业人员确认。
