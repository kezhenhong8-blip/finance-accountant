# 核对链验证示例（PowerShell）
# 使用前替换尖括号内的占位符

param(
  [string]$Bank    = "<财务数据根目录>/公账文件/7月份.xlsx",
  [string]$Taxbook = "<财务数据根目录>/个税申报/<公司>_个税扣缴申报_<年度>年度台账.xlsx",  # 可选
  [string]$Python  = "python",   # 或 D:/path/to/python.exe
  [string]$Soffice = ""         # 或 C:/Program Files/LibreOffice/program/soffice.exe
)

$env:PYTHONIOENCODING = "utf-8"

# 按当期实际核对填入期望值（禁止编造，以下为演示数值）
$debit    = 1000.00     # 借方合计 = 期间费用合计
$credit   = 5000.00     # 贷方合计 = 现金流入
$balance  = 4000.00     # 期末余额
$capital  = 4999.00     # 实收资本
$verify   = 1.00        # 银行验证款

$skillDir = Join-Path $PSScriptRoot "..\skills\financial-check\scripts"

$args = @(
  "--bank", $Bank,
  "--debit", "$debit", "--credit", "$credit", "--balance", "$balance",
  "--capital", "$capital", "--verify", "$verify"
)
if ($Taxbook) { $args += @("--taxbook", $Taxbook) }
if ($Soffice) { $args += @("--soffice", $Soffice) }

& $Python (Join-Path $skillDir "verify_chain.py") @args
exit $LASTEXITCODE