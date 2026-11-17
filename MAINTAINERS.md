# 维护

- 在 `skills/<技能名>/` 修改共享技能。入口写触发范围和必要步骤，使用细节放到参考文件，源项目链接和归属许可随修改一起保留。
- 用匿名的题目、示例报告和合成数据复测。问题反馈只附能公开的示例、版本、复现步骤与实际结果。
- 提交前在本地检查入口、引用和辅助脚本。熟悉 GitHub Actions 的维护者可将 [工作流模板](ci/github-actions.yml) 放到 `.github/workflows/check.yml` 并提交，届时推送和 PR 会自动运行。

```powershell
python -m pip install -r requirements-check.txt
python -X utf8 -B -m unittest discover -s tests -v
git diff --check
```

涉及 Word 原生导出时，在安装了 Word 的 Windows 上另跑 [原生导出回归](skills/hhy-coursework-delivery/references/regression.md#辅助工具)。涉及教学行为或交付判断时，用受影响的完整任务手动走一次。新版发布时记录实际执行的验证、变更和依赖，更新安装副本时整目录替换并保留旧副本。
