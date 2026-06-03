# MD 转 Word 📝→📄

> 基于 [Cherry Studio](https://github.com/CherryHQ/cherry-studio) 开源代码的 Markdown 转 Word 文档工具
>
> **项目地址：** https://github.com/hu568/md-to-word

将 Markdown 文件（`.md`）一键转换为 Word 文档（`.docx`），支持标题、表格、代码块、列表、图片嵌入、超链接等**常见 Markdown 语法**，保留整洁的排版风格。

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
| `[超链接](url)` | **真正的 Word 可点击超链接（Ctrl+单击跳转）** |
| `![图片](path)` | 内嵌图片（居中显示） |
| `![alt](path)<br>*图注*` | 图片居中 + 图注换行居中显示（小号斜体） |
| `---` 分割线 | 居中虚线 |
| `- [x]` 任务列表 | 支持识别（转为无序列表） |
| `$...$` 行内公式 | **Word 原生 OMML 公式（可直接编辑）** |
| `$$...$$` 块级公式 | 居中显示，Word 原生可编辑公式 |
| `\sum` `\int` `\frac` 等 | 完整 LaTeX 数学语法支持 |

## 🚀 快速开始

### 环境要求

- Python 3.8+
- [python-docx](https://python-docx.readthedocs.io/) — 生成 Word 文档
- [markdown-it-py](https://github.com/executablebooks/markdown-it-py) — Markdown 解析
- [latex2mathml](https://pypi.org/project/latex2mathml/) — LaTeX → MathML 转换
- [mathml2omml](https://pypi.org/project/mathml2omml/) — MathML → OMML（Word 原生公式格式）

### 安装依赖

```bash
pip install python-docx markdown-it-py latex2mathml mathml2omml
```

### 使用方法

**文件转换：**
```bash
python scripts/md_to_word.py input.md                     # 在同目录生成 input.docx
python scripts/md_to_word.py input.md output.docx          # 指定输出路径
```

**直接传字符串：**
```bash
python scripts/md_to_word.py --inline "# Hello\n\n**粗体**内容" hello.docx
```

## 🖼️ 图片与图注

支持 `![alt](path.jpg)<br>*▲ 图注文字*` 格式，自动将图片居中显示，图注换行至下一段并设为小号斜体居中。

转换效果：
- **图片段落** → 居中显示
- **图注段落** → 居中显示，小号斜体字

### 图片路径查找顺序

```
1. Markdown 中指定的原始路径
2. MD 源文件同目录下的相对路径
3. MD 源文件同目录下的 images/ 子目录
```

## 📐 LaTeX 数学公式

支持将 Markdown 中的 LaTeX 公式转换为 **Word 原生 OMML 公式**，可直接在 Word 中编辑和渲染，无需截图或插件。

### 语法支持

| 格式 | 说明 | 示例 |
|------|------|------|
| `$...$` | 行内公式 | `$E=mc^2$` |
| `$$...$$` | 块级公式（居中显示） | `$$\int_{-\infty}^{\infty} e^{-x^2}\,dx = \sqrt{\pi}$$` |

### 支持的 LaTeX 语法

基本运算、上下标、分式、根式、求和/积分/极限、矩阵、希腊字母、三角函数、矢量符号等标准 LaTeX 数学环境均可转换。

### 转换原理

借鉴 [md2word](https://md2word.com) 的转换管线思路：

```
LaTeX 公式文本 ($E=mc^2$)
    │
    ▼  latex2mathml
MathML (数学标记语言)
    │
    ▼  mathml2omml
OMML (Office Math Markup Language)
    │
    ▼  注入 python-docx
Word 文档中的可编辑公式
```

- **行内公式** `$...$` → 直接嵌入段落的 `m:oMath` XML 元素
- **块级公式** `$$...$$` → 创建居中 `m:oMathPara` 容器
- 纯本地转换，无需调用外部 API，完全离线可用
- 代码块和行内代码中的 `$` 符号会被自动保护，不会被误提取

## 🔗 超链接

表格内和段落中的链接都会渲染为 **Word 真正的可点击超链接**（蓝色 + 下划线），支持 Ctrl+单击直接跳转，而非单纯的蓝色文本样式。

## 🏗️ 实现原理

本工具参考了 Cherry Studio 的 `src/main/services/ExportService.ts` 实现，并借鉴了 [md2word](https://md2word.com) 的 LaTeX 转换管线（LaTeX → MathML → OMML）：

```
Markdown 文本（含 $...$ / $$...$$）
    │
    ├─ ① 预提取 LaTeX 公式，保护代码块
    │  （LaTeX → MathML → OMML 转换）
    │
    ▼
markdown-it-py 解析为 Token 流
    │
    ▼
遍历 Token，映射为 python-docx 元素
  ├─ 识别公式占位符 → 注入 m:oMath OMML 元素
  ├─ 普通文本 → 段落 / 标题 / 列表 / 表格
  └─ 图片 → 嵌入图片
    │
    ▼
生成 .docx 文件
    │
    ▼
图片后处理（可选）：扫描占位符，嵌入真实图片
```

- **解析层** — 使用 `markdown-it-py`（与 Cherry Studio 使用的 `markdown-it` JS 版同一标准）
- **构建层** — 使用 `python-docx` 构建 Word 文档元素
- **LaTeX 公式** — `latex2mathml` + `mathml2omml` 双库管线，输出 Word 原生 OMML 格式
- **表格风格** — 三线表设计（表头粗体、顶部/底部边框）
- **超链接** — 通过 `<w:hyperlink>` XML 元素 + 文档关系实现可点击超链接
- **图片策略** — 两阶段处理：主转换直接嵌入 + 后处理脚本替换占位符

## 📁 项目结构

```
md-to-word/
├── LICENSE                  # AGPL-3.0 协议
├── README.md                # 本文件
├── SKILL.md                 # CherryStudio Skill 描述
└── scripts/
    ├── md_to_word.py        # 核心转换脚本（主逻辑）
    └── embed_images.py      # 图片后处理模块（替换占位符）
```

## 📜 开源协议

本项目基于 **GNU Affero General Public License v3.0 (AGPL-3.0)** 发布。

本作品参考了 [Cherry Studio](https://github.com/CherryHQ/cherry-studio)（AGPL-3.0）的 `ExportService.ts` 实现思路，并从 TypeScript (Electron) 移植到了 Python 3。

根据 AGPL-3.0 第5节要求：
- ✅ 修改时间：2026年6月
- ✅ 修改内容：语言移植（TS → Python），依赖库替换，移除 Electron 依赖，改为命令行界面，增加图片居中与图注排版支持，增加可点击超链接支持，增加 LaTeX 数学公式转换支持
- ✅ 本作品同样以 AGPL-3.0 协议发布

## 🙏 致谢

- [Cherry Studio](https://github.com/CherryHQ/cherry-studio) — 优秀的开源 AI 客户端，提供导出功能的参考实现
- [markdown-it-py](https://github.com/executablebooks/markdown-it-py) — Python 版 Markdown 解析器
- [python-docx](https://python-docx.readthedocs.io/) — Python Word 文档生成库
- [latex2mathml](https://pypi.org/project/latex2mathml/) — LaTeX 转 MathML 库
- [mathml2omml](https://pypi.org/project/mathml2omml/) — MathML 转 OMML 库
- [md2word](https://md2word.com) — 在线 Markdown 转 Word 工具，提供 LaTeX → OMML 转换思路参考
