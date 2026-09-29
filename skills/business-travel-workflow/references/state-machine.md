# 差旅案件状态机

## States

| State | Meaning | Required gate to leave |
|---|---|---|
| `draft` | 案件已建立，信息仍在收集 | 申请字段完整并明确提交，或用户要求直接进入出差中 |
| `applied` | 出差申请已提交 | 有审批结果或用户明确记录未批准/已批准 |
| `approved` | 申请已获批准 | 出差实际开始或用户要求记录执行前准备 |
| `in_progress` | 出差进行中 | 每日记录已补齐，实际行程和成果可收尾 |
| `ready_to_close` | 已具备收尾材料 | 路线、成果、回程和问题已核对 |
| `ready_for_reimbursement` | 收尾通过，材料可进入报销 | 报销票据已关联且必需校验项通过 |
| `closed` | 案件归档完成 | 无未关闭的必需问题，输出文件已登记 |

## Transition rules

- `draft → applied`: 只在申请材料实际提交或用户明确说“已提交”时使用。
- `applied → approved`: 需要审批记录、审批截图、批准消息或用户明确确认。不能从申请表内容推断批准。
- `approved → in_progress`: 需要实际出差开始的日期或用户明确要求开始记录过程。
- `in_progress → ready_to_close`: 需要实际路线、工作成果、返程情况和每日记录的覆盖检查。
- `ready_to_close → ready_for_reimbursement`: 需要收尾检查通过；路线未闭环或关键事实缺失时保持阻塞。
- `ready_for_reimbursement → closed`: 需要报销技能返回已核对或已由用户接受的剩余问题，并登记最终文件。

## Parallel work

发票收集可以在 `in_progress` 阶段提前进行，但不改变案件主状态。把票据放入 `documents` 和 `expenses`，等收尾阶段再检查它们与实际行程的关联。

## Blockers

以下问题应阻止对应状态推进：缺少实际返程证据、住宿票缺入住或离店时间、购买方信息不一致、金额无法确认、审批状态未知、每日汇报与来源冲突。用户明确接受某个业务风险时，保留问题并记录接受者、时间和说明。
