# NOT_FOR_UPLOAD

以下内容按“**仅本地保留，不上传到公开仓库**”处理。

## 1. 本地验证与回归产物
- `e2e_*/`
- `scan_reports/`
- `tmp_md_audit/`
- `tmp_md_symbol_audit/`
- `tmp_toc_probe*/`
- `exec-clone-check-*/`
- `scene-upgrade-check-*/`
- `tmp_prompt_regression_exec_*/`
- 根目录下所有 `tmp_*.docx` / `_tmp_*.docx`
- 根目录下所有 `*_排版附件/`

## 2. 测试临时目录
- `tests/_tmp_*`
- `tests/_clone_*/`
- `tests/_scene_prompt_regression_*/`
- `tests/lf_*/`

## 3. 不公开的样本文档与规范素材
- `tests/00/`
- `tests/clone/`
- `tests/former-templates/`
- `tests/test-1/`
- `tests/test-equation/`
- `tests/test-markdown/`
- `tests/目录/`

## 4. 处理原则
- 这些文件可继续用于本地测试，但**不要提交到公开远程仓库**。
- 如需保留测试能力，请改为：
  - 代码动态生成的最小样本；
  - 你本人原创、可再分发的 synthetic 文档；
  - 不含真实姓名、学校规范正文、论文原文的脱敏样本。
- 对第三方规范原文、论文样本、对比稿、报告、日志，一律优先删除或替换。

详见：`docs/OPEN_SOURCE_MIT_RELEASE_CHECKLIST.md`

## 5. 发布前建议执行
- 运行 `python scripts/check_public_release.py` 做一次快速扫描。
- 若要用于发布前卡口或 CI，使用 `python scripts/check_public_release.py --strict`。
