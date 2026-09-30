# 邮箱发票采集接口

报销 Skill 可以接入只读邮箱适配器，自动发现并下载发票附件。邮箱适配器不是当前仓库的固定凭据或服务端实现；凭据、OAuth 缓存和邮箱地址必须保存在本地。

## 适配器优先级

1. Gmail：使用 Gmail API 和 OAuth，只申请只读邮件与附件权限。
2. 网易企业邮箱：使用企业邮箱官方接口；没有可用接口时使用只读 IMAP。
3. 不把邮箱密码、OAuth token、客户端密钥或邮件原文提交到 GitHub。

## 最小处理链

```text
读取指定文件夹
→ 按邮件 ID 去重
→ 检查主题、发件人、正文和附件名
→ 下载 PDF/图片/电子行程单
→ 保存邮件 ID、附件名、接收时间和来源
→ 交给 $mineru-pdf-to-md
→ 关联差旅案件和行程段
```

默认行为是只读、只下载，不删除、不移动、不发送邮件。识别为发票的规则仍需经过发票字段校验；不能只凭主题或附件名判定为可报销凭证。

## 待接入接口

- `mail.list_messages(folder, since)`
- `mail.get_message(message_id)`
- `mail.download_attachment(message_id, attachment_id, output)`
- `mail.mark_source(message_id, attachment_id)`

当本机配置 Gmail 或网易邮箱 MCP/适配器后，`company-expense-reimbursement` 可以调用这些接口；没有连接器时继续支持手动上传文件。
