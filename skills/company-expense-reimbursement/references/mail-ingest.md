# 邮箱发票采集

报销 Skill 可以在用户明确要求“扫描邮箱”“从邮箱找发票”或启用自动采集时，使用只读邮箱接口发现并下载发票附件。邮箱地址、密码、OAuth 客户端文件和 token 必须保存在本地，不能写入这个公开仓库。

## 连接器选择

按以下顺序选择连接器：

1. **MCP 邮箱工具**：如果当前会话提供只读的 `mail.list_messages`、`mail.get_message` 和 `mail.download_attachment`，优先使用它们。不要调用删除、移动、标记已读或发信接口。
2. **Gmail API**：本仓库提供 `scripts/mail_ingest.py` 的 Gmail OAuth 只读实现。它需要本地 OAuth 客户端 JSON，第一次运行会打开浏览器授权，之后使用本地 token。
3. **IMAP**：网易企业邮箱使用 IMAP；Gmail 也可以使用 IMAP + 应用专用密码作为备用方案。服务器地址必须由邮箱管理员或邮箱设置确认，不能凭猜测写死。

当前仓库没有内置账号和凭据，也不会自动发现用户密码。没有 MCP 或本地配置时，继续支持手动上传 PDF/图片。

## 本地配置

复制 [mail-config.example.yaml](../../../profiles/mail-config.example.yaml) 到：

```text
C:\Users\<user>\.codex\private\mail-config.yaml
```

Gmail 示例：

```yaml
provider: gmail
gmail:
  credentials_json: C:\Users\<user>\.codex\private\google-oauth-client.json
  token_json: C:\Users\<user>\.codex\private\gmail-token.json
  account: name@example.com
  query: has:attachment
```

网易企业邮箱示例：

```yaml
provider: imap
imap:
  host: imap.qiye.163.com
  port: 993
  ssl: true
  username: name@company.example
  password_env: COMPANY_MAIL_PASSWORD
  folder: INBOX
```

优先使用 `password_env` 或 `password_file`。不要把明文密码、应用专用密码、OAuth token 或客户端密钥提交到 GitHub。

如果不想设置环境变量或密码文件，可以在交互式 PowerShell 中省略密码配置；采集器会在连接时隐藏式提示输入密码。通过自动任务、重定向或管道运行时不能使用隐藏提示，此时应使用临时环境变量或本地密码文件。

## 本地采集命令

Gmail API：

```powershell
python skills\company-expense-reimbursement\scripts\mail_ingest.py `
  --config "$env:USERPROFILE\.codex\private\mail-config.yaml" `
  --since 2026-09-01 `
  --output ".\cases\TRIP-YYYY-001\03-发票凭证\邮箱附件"
```

网易企业邮箱或 Gmail IMAP：

```powershell
$env:COMPANY_MAIL_PASSWORD = "在当前 PowerShell 会话临时设置的应用专用密码"
python skills\company-expense-reimbursement\scripts\mail_ingest.py `
  --provider imap `
  --config "$env:USERPROFILE\.codex\private\mail-config.yaml" `
  --since 2026-09-01 `
  --output ".\cases\TRIP-YYYY-001\03-发票凭证\邮箱附件"
```

第一次接入建议先执行 `--dry-run`，确认候选邮件和附件，再正式下载：

```powershell
python skills\company-expense-reimbursement\scripts\mail_ingest.py `
  --config "$env:USERPROFILE\.codex\private\mail-config.yaml" `
  --since 2026-09-01 --dry-run
```

如果配置文件还是示例值，脚本会在交互式终端中询问完整邮箱地址和 IMAP 服务器地址。也可以显式传入并保存非敏感设置：

```powershell
python skills\company-expense-reimbursement\scripts\mail_ingest.py `
  --provider imap `
  --config "$env:USERPROFILE\.codex\private\mail-config.yaml" `
  --email "你的企业邮箱地址" `
  --imap-host "企业邮箱后台显示的 IMAP 服务器" `
  --save-mail-settings --setup --dry-run
```

`--save-mail-settings` 只保存邮箱地址和 IMAP 服务器，不保存密码。邮箱域名和 IMAP 主机是两个独立设置，不能从一个值可靠推导另一个值。

采集器默认只下载 PDF、PNG、JPG、JPEG 和 OFD，并用邮件主题、发件人、正文和附件名判断是否像发票或交通凭证。需要把所有支持格式的附件交给人工筛选时，使用 `--include-all-attachments`。

## 去重与来源追踪

每个输出目录会生成：

- `.mail-ingest-state.json`：按 `provider + message_id + attachment_id` 去重，并按 SHA-256 防止同一文件重复保存。
- `mail-index.json`：记录邮件 ID、附件 ID、发件人、主题、接收时间、附件名、哈希、保存路径和候选判断理由。

采集器是只读的：不标记已读、不移动、不删除、不发送邮件。邮件源索引必须和附件一起保留，之后才能把附件关联到差旅案件和具体行程段。

## 交给 MinerU

下载完成后，对每一个 PDF 或页面图片运行 `$mineru-pdf-to-md`。不能仅凭邮件主题或附件名把文件直接当作可报销发票；仍需检查发票字段、公司信息、住宿入住/离店时间和路线闭环。识别失败或字段冲突时，保留原邮件来源并标记人工复核。

## MCP 适配器接口

如果以后接入 MCP，使用以下最小只读契约即可替换本地脚本：

```text
mail.list_messages(folder, since, until, query)
mail.get_message(message_id)
mail.download_attachment(message_id, attachment_id, output)
```

MCP 返回的消息至少应包含 `message_id`、`sender`、`subject`、`received_at` 和附件的 `attachment_id`、`filename`、`mime_type`。下载后仍要写入同样的 `mail-index.json`，这样 MCP 和本地 IMAP/Gmail API 的结果可以共用后续 MinerU 与报销流程。
