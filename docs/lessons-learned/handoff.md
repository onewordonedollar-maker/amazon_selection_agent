# Codex / Cursor 交接规则

## 开始工作前

任何工具接手项目时先执行：

```powershell
cd D:\amazon_selection_agent
git pull
git status
git log --oneline -5
```

然后阅读：

- `README.md`
- `docs/lessons-learned/project-guardrails.md`
- `docs/lessons-learned/pitfalls.md`
- 与当前任务直接相关的代码和测试

## 推荐任务提示

交给新的编码工具时，可使用：

```text
先阅读 README.md 和 docs/lessons-learned/。
不要修改已锁定的采集等待、滚动、翻页和筛选语义。
每次只处理一个明确问题，修改后运行相关测试并验证真实界面。
如需改变核心采集或筛选规则，必须先停止并说明影响。
不要删除或覆盖用户未要求修改的代码。
```

## 工作粒度

- 一个任务对应一个清晰目标。
- 一个提交尽量只包含一种行为变化。
- UI 修改与采集逻辑修改分开提交。
- 重构与修复分开提交。
- 大规模修改前先创建稳定提交。

## 交接信息

每次切换工具前，至少留下：

- 当前分支与提交号
- 已完成内容
- 未完成内容
- 已运行的测试
- 已知风险
- 是否修改了核心规则

## 发生冲突时

- 不使用 `git reset --hard`。
- 不覆盖不理解的本地改动。
- 先执行 `git status` 和 `git diff`。
- 如果两个工具修改了同一文件，逐块合并并重新测试。

