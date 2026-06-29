# 普通中文报告交付规则

用于股票分析报告、课程报告、案例分析、管理类报告等非 USTC 学位论文的中文 Word/PDF 交付。

## 默认资产

- 默认报告模板：`assets/report_template.docx`
- 该模板只作为基础样式来源，不保证自动修复图表、三线表、标题边框或 PDF 字体。

## 推荐流程

1. 先生成 UTF-8 Markdown/JSON 内容源。
2. 生成临时 DOCX，不直接覆盖最终文件。
3. 对 DOCX 做结构后处理：字体、标题边框、图文块、三线表、段落缩进、表格长文本拆分。
4. 对报告类 DOCX 运行主题清理，移除 `styles.xml` 和 `theme1.xml` 中的蓝色标题残留、`accent1` 主题色、标题边框和下划线。推荐使用：

```bash
python3 /Users/xiaofei/.codex/skills/chinese-word-pro/scripts/clean_report_theme.py \
  --input-docx temp.docx \
  --output-docx temp_clean.docx
```

4. 检查 `word/document.xml`、`word/styles.xml` 和 `word/theme/theme1.xml`，确认无乱码、无蓝色标题残留、无异常段落边框。
5. 用 `scripts/export_pdf_fixed.py` 导出 PDF，并渲染 PNG 页面做视觉 QA。
6. 重点抽检标题页、第一张表所在页、主要图表页、最后一个风险表页。
7. QA 通过后再复制为最终 DOCX/PDF。

## 标题规则

- 中文报告标题和章节标题默认黑色，不使用蓝色主题色。
- 必须清理模板残留的 `w:pBdr` 段落边框，尤其是 `Title` 样式的蓝色底边框。
- 必须清理标题 run/style 中的 `w:u` 下划线、`themeColor="accent1"`、`4F81BD` 等蓝色主题残留。
- 必须同时清理 `word/theme/theme1.xml` 中的蓝色 `accent1`/`4F81BD` 残留。主题文件路径也匹配 `word/*.xml`，脚本分支中应先处理 `word/theme/*.xml`，再处理普通 `word/*.xml`。
- 标题字体：中文黑体，英文 Times New Roman。

## 正文字体规则

- 中文正文：宋体。
- 英文与数字：Times New Roman。
- 在 `python-docx` 中必须同时设置 `run.font.name` 和 `w:eastAsia`。
- 不要只依赖模板继承字体。

## 图片规则

- 图像必须是 inline，不使用浮动 anchor。
- 图像文件优先 PNG。
- 图片内部不要写“图 1”或标题，图题必须用 Word 文本生成。
- 图题在图片下方，单独段落，中文报告中可按用户要求居中。
- 普通段落插图容易被固定行距裁切；若曾出现裁切、遮挡、图片不显示，应使用无边框 1 列 2 行表格作为图文块：
  - 第一行：居中图片。
  - 第二行：居中图题。
  - 表格无边框，单元格边距为 0。
  - 图片宽度必须小于页面可用宽度。

## 表格规则

- 表题放在表格上方。
- 表格只承载结构化信息，例如指标、金额、占比、同比变化、风险类型、跟踪重点。
- 不要默认设置长篇分析列，例如 `分析判断`、`原因说明`、`投资含义`、`综合评价`。如果单元格需要一句以上的解释，应删除该列，把分析写在表格下方的正文段落中。
- 三线表必须使用单元格级边框，不能只依赖 `Table Grid` 或表格整体样式。
- 默认三线表：
  - 顶线：1.5 pt。
  - 表头下线：0.5 pt。
  - 底线：1.5 pt。
  - 无竖线，无普通内部横线。
- 表格段落必须清零首行缩进、左右缩进，设置合适行距，避免单元格文字被挤压或裁切。
- 表格生成后要做可读性检查：列数过多、单元格出现多行长句、表格横向拥挤时，应优先缩减列，而不是缩小字号硬塞。

## OOXML 清理规则

- 普通中文报告不能只依赖肉眼检查标题是否变黑，必须检查底层 XML。
- 蓝色残留检查至少覆盖 `word/document.xml`、`word/styles.xml` 和 `word/theme/theme1.xml`。
- 必须清理或替换 `4F81BD`、`w:themeColor="accent1"`、`w:themeFill="accent1"`、`<w:u .../>` 和标题相关 `<w:pBdr>...</w:pBdr>`。
- 对报告类文档，优先使用 `scripts/clean_report_theme.py` 作为通用清理器，再做人工或渲染复检。
- 如果仅剩 `word/settings.xml` 中的 `w:accent1="accent1"` 映射，一般不是可见样式问题；但 `styles.xml` 或 `theme1.xml` 中仍有蓝色值时，不能交付。
- XML 清理脚本要先匹配 `word/theme/*.xml`，再匹配普通 `word/*.xml`。不要让 `theme1.xml` 被普通分支提前处理，否则主题色可能残留。

## PDF 导出规则

- PDF 导出不能只看是否生成成功，必须检查字体和渲染页。
- macOS + LibreOffice headless 下，应使用 `scripts/export_pdf_fixed.py`：

```bash
python3 /Users/xiaofei/.codex/skills/chinese-word-pro/scripts/export_pdf_fixed.py \
  --input-docx final.docx \
  --output-pdf final.pdf \
  --render-dir rendered_pages
```

- 该脚本会生成 PDF 安全 DOCX，将 `宋体/黑体` 映射到 macOS/LibreOffice 可识别的 `宋体-简/黑体-简`，并写入本地 fontconfig。
- PDF 字体列表中不应出现异常替代字体，例如 `PingFang`、`HiraginoSans`、`HiraMaruPro`、`ArialUnicodeMS`。

## 最终 QA 清单

- 标题下无蓝色横线或下划线。
- `word/document.xml`、`word/styles.xml` 和 `word/theme/theme1.xml` 中无 `4F81BD`、标题相关 `themeColor="accent1"`、`w:u` 下划线、异常 `w:pBdr`。
- 正文中文为宋体，英文为 Times New Roman。
- 图能实际显示，且未被文字覆盖或裁切。
- 图题在图下方，表题在表上方。
- 表格为严格三线表。
- 表格无长篇分析列；分析判断写在表格下方正文中。
- PDF 字体没有异常替代。
- 渲染 PNG 中无乱码、无方框字、无明显错位。

## 重点页视觉抽检

对于图表较多的报告，至少抽检以下页面：

- 标题页或第一页：确认标题黑色、无蓝色横线、无下划线。
- 第一张表所在页：确认表题在上、三线表干净、无长文本挤压。
- 第一张图所在页：确认图片显示完整、图题在下方、没有被正文遮挡。
- 关键截图或 K 线图页：确认图像清晰、未裁切、未跨页错位。
- 最后一张风险表或结论页：确认表格底线完整、页面无大块异常空白。
