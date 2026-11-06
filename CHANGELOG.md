# Changelog

## 1.7.0

First public release. 首个公开版本。

- Package renamed from `@local/doc-viewer` to `dsh-doc-viewer`; MIT license.
  包名从 `@local/doc-viewer` 改为 `dsh-doc-viewer`，MIT 许可。装过旧版的要先卸掉旧包。
- SheetJS upgraded to 0.20.3, fixing the known prototype-pollution and ReDoS issues in 0.18.5. It now ships inside the package (`vendor/xlsx`), because SheetJS is no longer on npm.
  SheetJS 升到 0.20.3，修掉 0.18.5 的原型污染和 ReDoS 漏洞。新版不在 npm 上，所以直接随插件发布（`vendor/xlsx`）。
- The OCR helper is built from source when packing: one universal binary for Apple silicon and Intel, macOS 13 or later.
  OCR 小程序改为打包时从源码编译，Apple 芯片和 Intel 都能用，macOS 13 起。
- Declares `engines.dsh >= 0.2.2`.
- THIRD_PARTY_LICENSES.md carries the full license text of everything bundled or vendored.
  附上所有打包进来和随附代码的许可证全文。
- Text and Markdown files can be edited and saved in place, keeping their encoding and line endings; unsaved edits survive a restart.
  文本和 Markdown 可以直接编辑并存回原文件，保留编码和换行，没保存的修改重启后还在。
- Cells the agent changed stay outlined until you click into the table or type (a minute at most), instead of fading after six seconds, usually before the agent had finished replying.
  Agent 改过的格子一直亮到你点进表格或开始输入（最多 1 分钟）；以前 6 秒就淡出，常常 Agent 还没回答完就没了。
- When LibreOffice can't recalculate after an agent edit (busy or missing), the open table works the formulas out itself instead of showing them blank, and fills in LibreOffice's answers when they come.
  Agent 改完如果 LibreOffice 没算出公式结果（机器太忙或没有 LibreOffice），表格自己先算出来，不再显示成空白。
- `doc_viewer_edit` reports each cell before and after, and the formula results that moved because of it, so the agent needn't read the file back another way.
  `doc_viewer_edit` 直接报告改前改后和连带变化的公式结果，Agent 不用再自己跑命令核对。

## 1.6.1

- Opening a workbook no longer waits for LibreOffice: formulas without saved results are shown by the grid's own engine at once and filled in from LibreOffice in the background (a cold start used to take about 20 s).
  打开表格不再等 LibreOffice 冷启动，后台补公式结果。

## 1.6.0

- One motion system (140–200 ms, one easing curve), skeleton loading, tab and toast exit animations, reduced-motion respected.
- Consistent wording: 保存 / 另存为… / 导出 PDF…; an empty state that explains how to start; a one-time hint about opening beside the conversation.

Earlier versions were not published.
