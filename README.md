# Company Travel Skills

这是一个可持续迭代的差旅 Skill 套件，包含：

- `business-travel-workflow`：差旅案件主流程和状态机
- `travel-application`：出差申请
- `travel-daily-report`：出差每日汇报
- `travel-closure`：出差收尾
- `company-expense-reimbursement`：发票解析和报销单

仓库只保存通用 Skill 逻辑和空白模板。公司税号、银行账号、地址、员工姓名、职务、真实发票和差旅案件必须保存在本地或私有仓库。

## 安装

先查看仓库可安装的 Skill：

```powershell
npx skills add xxrust/company-travel-skills --list
```

安装全部 Skill：

```powershell
npx skills add xxrust/company-travel-skills --skill '*'
```

只安装主流程：

```powershell
npx skills add xxrust/company-travel-skills --skill business-travel-workflow
```

`npx skills add` 是跨 Agent 的分发方式。若要将选定 Skill 明确映射到本机 `C:\Users\<user>\.codex\skills`，使用：

```powershell
powershell -ExecutionPolicy Bypass -File .\scripts\install-junctions.ps1
```

开发时使用 Junction，发布时使用 GitHub Release 或 `npx skills add` 安装固定版本。不要把本机整个 `.codex\skills` 目录提交到仓库。

## PowerShell / GitHub 网络配置

如果浏览器通过 Clash 等 VPN 客户端访问 GitHub，而 Git 或 `npx` 不能访问，可以使用仓库中的脚本把代理配置到当前 PowerShell、Git 和 npm：

```powershell
powershell -ExecutionPolicy Bypass -File .\scripts\configure-github-network.ps1 -Persist
```

脚本默认使用 `http://127.0.0.1:7890`。如果 VPN 使用其他端口，传入实际地址：

```powershell
powershell -ExecutionPolicy Bypass -File .\scripts\configure-github-network.ps1 `
  -Proxy http://127.0.0.1:7890 -Persist
```

清除脚本写入的用户级代理配置：

```powershell
powershell -ExecutionPolicy Bypass -File .\scripts\configure-github-network.ps1 -Clear
```

脚本只配置代理，不保存 GitHub 令牌或账号密码。

## 本地私有配置

创建以下文件：

```text
C:\Users\<user>\.codex\private\company-profile.yaml
C:\Users\<user>\.codex\private\user-profile.yaml
```

可参考：

- `profiles/company-profile.example.yaml`
- `profiles/user-profile.example.yaml`

报销脚本会自动查找这些路径，也可以使用 `--profile`、`--user-profile` 或环境变量覆盖。

## 本地开发

运行全部 Skill 校验：

```powershell
python .\scripts\validate_skills.py
```

生成空白报销模板：

```powershell
python .\skills\company-expense-reimbursement\scripts\build_template.py --output .\.tmp\company-expense-template.xlsx
```

使用本地公司和个人配置生成报销单：

```powershell
python .\skills\company-expense-reimbursement\scripts\new_workbook.py `
  --output .\cases\TRIP-2026-001\05-报销\差旅费报销单.xlsx
```

## 依赖

处理 PDF 时需要单独安装或启用 `$mineru-pdf-to-md`。这个仓库只负责差旅业务流程，不复制外部 MinerU Skill 或模型文件。
