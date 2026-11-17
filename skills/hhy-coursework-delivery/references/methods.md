# 方法来源与取舍

核对日期：2026-10-16。下面是能力和设计参考，不是效果排名；未做跨产品同题对比，不宣称某个公开技能最好。正文与三个辅助工具为本次独立编写，不整包拼接第三方技能。

- [OpenAI 技能说明](https://learn.chatgpt.com/docs/build-skills)：采用简短入口、按需加载参考文件、明确触发范围。保留已有技能名称和入口，不新增功能重叠的技能。
- [OpenAI 技能评测方法](https://developers.openai.com/blog/eval-skills)：采用真实任务、可观察结果与异常场景的回归检查。格式验证、工具运行和实际代理行为评测分别记录，不把一个检查通过等同于整个流程有效。
- [Anthropic DOCX 参考](https://github.com/anthropics/skills/blob/dbd4588f9e1033efb41dad4bef2f7947c8993d44/skills/docx/SKILL.md)：参考渲染后查看页面，以及注意Word文本分散在多个run中的问题。该目录的LICENSE.txt为专有服务条款，未复制其代码、提示词或资产；沿用本机已有文档能力并独立实现只读检查工具。
- [Anthropic 文档共创参考](https://github.com/anthropics/skills/blob/dbd4588f9e1033efb41dad4bef2f7947c8993d44/skills/doc-coauthoring/SKILL.md)：只提取“脱离聊天，站在读者角度检查是否能理解”的概念。未复制其分阶段确认、频繁追问和默认另开代理的流程；独立评阅仍遵守当前授权。
- [Microsoft 目录样式说明](https://support.microsoft.com/en-au/word/format-or-customize-a-table-of-contents-in-word)：采用标题层级、目录样式、制表位和页码对齐，不用手打点线代替目录。
- [OfficeCLI 官方查看命令](https://github.com/iOfficeAI/OfficeCLI/wiki/command-view)：用结构、问题提示和原生截图辅助定位，参数以本机help为准；不把HTML分页或结构校验视为Word视觉验收。

两个Anthropic页面固定到仓库提交 dbd4588f9e1033efb41dad4bef2f7947c8993d44（本次通过git ls-remote核对），便于追溯。官方在线产品文档未固定版本，重新研究相关功能时应再查当前页面。

## 用户习惯如何落地

以本次用户明确要求和既有选择为准：正式论文风格；模板优先；质量优先于页数；结果要有解释；实验截图真实清晰；代码实际运行；修改保留原件；旧稿分开、正式入口唯一；结束时清理获准的中间产物。宋体等只作为无其他模板或风格时中文课程报告的起点，不硬套其他文体。

沟通保持简短直接，能从文件查到的不追问；例行步骤不逐条预告。有实质发现、改向或受限时才说明。已确认的事不重复问，技能也不改变当前发送、删除、覆盖等授权边界。

作业的学号、邮箱、参数、页数和通过项数留在任务记录，不写入通用流程。

## 分项参考与许可

以下链接供追溯，无需在日常使用时联网加载。参考审阅与验证思路，以本地课程场景重新编写，未移植第三方工具或固定问答轮次。

- [Superpowers verification-before-completion](https://github.com/obra/superpowers/blob/bb92a77741419a4ab5f06e711a283343f1ada0c3/skills/verification-before-completion/SKILL.md)：完成结论要有适用于当前产物的证据。

- [Huashu Report](https://github.com/alchaincyf/huashu-report/blob/eda161cf693f2f2016b09abc1385b8e3df5fb9ad/SKILL.md)：借鉴先核对数据口径再组织论证和图表的流程，按老师模板与课程范围重写。
- [Tramstop](https://github.com/alchaincyf/tramstop-skill/blob/01a2c7d2c4a5a32c3576d2d739117912972c9bc3/SKILL.md)：借鉴真实素材与结构服务保留原则，不引入作者口吻、检测率目标或人物经历。
- [Khazix neat-freak](https://github.com/KKKKhazix/khazix-skills/blob/322346ded8129436b3f64707789a73e732ae24d9/neat-freak/SKILL.md)：借鉴变更影响到相关文档的同步思路，仅用于本次交付入口和引用。

以上三项上游均为MIT许可，许可原文见[上游许可](upstream-licenses.txt)。

表格与答辩参考：
- [Huashu Excel](https://github.com/alchaincyf/huashu-excel/blob/5b8831257eeac8ce84d3c63c3c969dadaab15981/SKILL.md)：借鉴原表体检、追溯与核数，不采用强制外部基准或默认中位数。
- [Huashu Speech Coach](https://github.com/alchaincyf/huashu-skills/blob/fdb4086348cebda0ec9cf81cdcb546cc18ff3020/huashu-speech-coach/SKILL.md)：借鉴听众视角、讲稿与幻灯片分工，不照搬固定字号、结束页禁令或观众比例。
这两项上游MIT许可原文同存于[上游许可](upstream-licenses.txt)。
