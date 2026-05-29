# MD 转 Word 📝→📄

> 基于 [Cherry Studio](https://github.com/CherryHQ/cherry-studio) 开源代码的 Markdown 转 Word 文档工具
>
> **项目地址：** https://github.com/hu568/md-to-word

将 Markdown 文件（`.md`）一键转换为 Word 文档（`.docx`），支持标题、表格、代码块、列表等**常见 Markdown 语法**，保留整洁的排版风格。

## ✨ 功能特性

| 语法 | 转换效果 |
|------|----------|
| 标题 `# ~ ######` | Word 内置标题样式（Heading 1~6） |
| **粗体** / *斜体* / ~~删除线~~ | Word 字体样式（Bold / Italic / Strikethrough） |
| `` `行内代码` `` | Consolas 等宽字体 + 红色高亮 |
| ` ``` ` 代码块 | 灰色背景 + 四周边框 |
| `- ` 无序列表 / `1.` 有序列表 | Word 项目符号 / 编号列表 |
| `| 表格 |` | 三线表风格（表头加粗） |
| `> 引用` | 左侧灰线 + 斜体灰色文字 |
| `[超链接](url)` | 蓝色下划线样式 |
| `![图片](path)` | 内嵌图片（本地路径） |
| `---` 分割线 | 居中虚线 |
| `- [x]` 任务列表 | 支持识别（转为无序列表） |

## 🚀 快速开始

### 环境要求

- Python 3.8+
- [python-docx](https://python-docx.readthedocs.io/) — 生成 Word 文档
- [markdown-it-py](https://github.com/executablebooks/markdown-it-py) — Markdown 解析

### 安装依赖

```bash
pip install python-docx markdown-it-py
```

### 使用方法

**文件转换：**
```bash
python md_to_word.py input.md                     # 在同目录生成 input.docx
python md_to_word.py input.md output.docx          # 指定输出路径
```

**直接传字符串：**
```bash
python md_to_word.py --inline "# Hello\n\n**粗体**内容" hello.docx
```

## 🏗️ 实现原理

本工具参考了 Cherry Studio 的 `src/main/services/ExportService.ts` 实现：

```
Markdown 文本
    │
    ▼
markdown-it-py 解析为 Token 流
    │
    ▼
遍历 Token，映射为 python-docx 元素
    │
    ▼
生成 .docx 文件
```

- **解析层** — 使用 `markdown-it-py`（与 Cherry Studio 使用的 `markdown-it` JS 版同一标准）
- **构建层** — 使用 `python-docx` 构建 Word 文档元素
- **表格风格** — 三线表设计（表头粗体、顶部/底部边框）

## 📁 项目结构

```
md-to-word/
├── LICENSE              # AGPL-3.0 协议
├── README.md            # 本文件
├── SKILL.md             # CherryStudio Skill 描述
└── scripts/
    └── md_to_word.py    # 核心转换脚本 (~400行)
```

## 📜 开源协议

本项目基于 **GNU Affero General Public License v3.0 (AGPL-3.0)** 发布。

本作品参考了 [Cherry Studio](https://github.com/CherryHQ/cherry-studio)（AGPL-3.0）的 `ExportService.ts` 实现思路，并从 TypeScript (Electron) 移植到了 Python 3。

根据 AGPL-3.0 第5节要求：
- ✅ 修改时间：2026年5月
- ✅ 修改内容：语言移植（TS → Python），依赖库替换，移除 Electron 依赖，改为命令行界面
- ✅ 本作品同样以 AGPL-3.0 协议发布

## 🙏 致谢

- [Cherry Studio](https://github.com/CherryHQ/cherry-studio) — 优秀的开源 AI 客户端，提供导出功能的参考实现
- [markdown-it-py](https://github.com/executablebooks/markdown-it-py) — Python 版 Markdown 解析器
- [python-docx](https://python-docx.readthedocs.io/) — Python Word 文档生成库
