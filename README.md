# Codex 学习与课程交付 Skills

两份可单独安装的 Codex 技能，分别用于学习和课程交付。主要使用中文。

| 技能 | 适用场景 | 可以直接这样说 |
| --- | --- | --- |
| [hhy-learning-coach](skills/hhy-learning-coach/SKILL.md) | 期末复习、代码和张量推导、英文论文段落、检查自己是否能独立解释或做变式 | 用 `$hhy-learning-coach` 从上次的决策树例题继续，按我的课件讲，最后说重点和易错点。 |
| [hhy-coursework-delivery](skills/hhy-coursework-delivery/SKILL.md) | 对照老师要求完成、复核报告与配套代码，核实数据和版式 | 用 `$hhy-coursework-delivery` 对照要求复查这份报告，保留原稿，核对数据与交付格式。 |

## 从一个小任务开始

| 想要的结果 | 先给 Codex 什么 | 这次可以检查什么 |
| --- | --- | --- |
| 自己会做题或读懂代码 | 一道课堂例题、课件范围、你已经做出的步骤 | 能否解释为什么这样做，或者按你的要求完成一道变式 |
| 作业材料齐全、结论可核对 | 老师要求、当前报告、能重现结果的数据和代码 | 是否逐条覆盖要求，正文与原始结果是否一致，配套文件能否使用 |

课程交付辅助脚本可以逐项使用。结构审计用于查看 DOCX 内容，页面比较用于定位变化，最终版式在 Word 或 PDF 中查看。两个技能的固定版和更新记录见 [Releases](https://github.com/han139242/codex-learning-skills/releases)。

## 安装

克隆仓库后，在根目录运行以下 PowerShell 命令。一次装一个技能。已有同名目录时，先比较内容，再自行备份并更新旧副本。重开 Codex 对话后输入技能名。

```powershell
git clone https://github.com/han139242/codex-learning-skills.git
Set-Location codex-learning-skills
$skill = 'hhy-learning-coach' # 可改为 hhy-coursework-delivery
$skillsDir = Join-Path ([Environment]::GetFolderPath('UserProfile')) '.agents/skills'
$target = Join-Path $skillsDir $skill
if (Test-Path -LiteralPath $target) { throw '已有同名技能，请先比较并备份旧版' }
New-Item -Path $skillsDir -ItemType Directory -Force | Out-Null
Copy-Item -LiteralPath (Join-Path 'skills' $skill) -Destination $target -Recurse
```

学习辅导与交付流程均可直接读取 Markdown。课程交付的辅助脚本按需使用：DOCX 结构审计需 Python，PNG 页面比较另需 Pillow，原生 Word 导出需 Windows、已安装的 Microsoft Word 和 PowerShell 7。具体参数见 [Word 终稿流程](skills/hhy-coursework-delivery/references/word-finalization.md)。完成报告准备后，由用户按课程要求办理正式提交。

代码和文档采用 [MIT](LICENSE)。借鉴来源及各自许可随 [学习辅导](skills/hhy-learning-coach/references/sources.md) 和 [课程交付](skills/hhy-coursework-delivery/references/methods.md) 分发。验证、更新和反馈见 [维护说明](MAINTAINERS.md)。

需要 Codex 与 DSH 独立讨论和复查时，可另装 [Codex + DSH 协作](https://github.com/han139242/codex-dsh-collaboration)。
