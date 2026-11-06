# dsh-doc-viewer

[English](README.md) · 简体中文

[DeepSeek Harness](https://github.com/deepseek-ai/deepseek-harness) 插件：你和 Agent 一起直接改你自己的表格。Agent 改了哪些格子，会在你开着的表格里亮出来，所有内容都留在本机。

![在对话里说一句，Agent 改好开着的表格，改过的格子用蓝框标出](docs/screenshot.png)

## 功能

- **改的就是原文件。** 只写回改动的单元格。文件里有保留不了的内容（图表、数据透视表、宏）时，绝不覆盖。
- **你和 Agent 改同一份。** Agent 的改动直接进开着的表格，和你没保存的修改放在一起。
- **改了哪里看得见。** 改过的格子一直亮到你点进表格，每处改动都能跳过去看，也能撤销。
- **不添乱。** 不上云，也不在你的文件夹里多放副本。保存时自动留版本，可以撤销。

还能打开 PDF（阅读、合并、拆分、OCR）、Word（预览）、文本和 Markdown（编辑）。

## 安装

需要 DeepSeek Harness 0.2.2 或更新版本。

```bash
DSH="/Applications/DeepSeek Harness.app/Contents/Resources/runtime/cli/bin/dsh"
"$DSH" plugin --profile desktop add https://github.com/Wildchiken/dsh-doc-viewer/releases/download/v1.7.0/dsh-doc-viewer-1.7.0.tgz
```

然后在 `~/.dsh/profiles/desktop/package.json` 的 `dsh.profile.bundles` 里加上 `"dsh-doc-viewer"`，重启 Harness。

更新时先运行 `plugin --profile desktop remove dsh-doc-viewer`，再 `add` [最新 Release](https://github.com/Wildchiken/dsh-doc-viewer/releases/latest) 里的文件。

## 限制

- 只在 macOS 上测过。
- Word 和 PDF 不能编辑，超过 2 MB 的文本文件只读。
- 2 万行的表可能要几十秒才能打开。
- OCR 需要 macOS 13 或更新版本。

更多内容见[详细说明](docs/guide.zh-CN.md)。

## 许可

MIT，第三方许可见 [THIRD_PARTY_LICENSES.md](THIRD_PARTY_LICENSES.md)。
