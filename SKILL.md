---
name: MD转Word
description: |
  将 Markdown 文件（.md）转换为 Word 文档（.docx），支持本地图片自动嵌入。
  支持标题、段落、粗体/斜体、代码块、列表（有序/无序）、表格、超链接、图片（自动嵌入）、引用块、水平线、任务列表等常见 Markdown 语法。
  当用户提到"md转word"、"markdown转docx"、"导出为word"、"md文件转换"、"markdown to word"、"保存为word"等时，自动触发此技能。
  当用户需要把笔记、文档、聊天记录等 Markdown 内容转换为 Word 格式时使用。
  **触发关键词**：md转word、md转docx、markdown转word、导出word、保存为word、markdown to word、convert md to docx、生成word文档
---

# MD 转 Word

参考 [Cherry Studio 开源实现](https://github.com/CherryHQ/cherry-studio) 思路，使用 `markdown-it-py` 解析 Markdown 为 token，再用 `python-docx` 生成 Word 文档。

## 依赖
- `python-docx` (~1.2.0) — 已安装
- `markdown-it-py` (~4.0.0) — 已安装

## 使用方法

```bash
python "scripts/md_to_word.py" <input.md> [output.docx]
```

直接传入 Markdown 字符串：
```bash
python "scripts/md_to_word.py" --inline "<markdown字符串>" output.docx
```

## 支持的 Markdown 语法

| 语法 | 说明 |
|------|------|
| `# ~ ######` | 标题 h1~h6 |
| `**text**` / `__text__` | 粗体 |
| `*text*` / `_text_` | 斜体 |
| `` `code` `` | 行内代码（Consolas 字体） |
| ` ```lang ... ``` ` | 代码块（灰色背景 + 边框） |
| `- item` / `* item` | 无序列表 |
| `1. item` | 有序列表 |
| `- [ ]` / `- [x]` | 任务列表 |
| `> quote` | 引用块（左侧灰线） |
| `---` / `***` | 水平分割线 |
| `| A | B |` | 表格（三线表风格） |
| `[text](url)` | 超链接 |
| `![alt](local_path)` | 图片（本地路径自动嵌入，支持图注） |
| `~~text~~` | 删除线 |

## 图片处理

支持以下三种图片格式：

1. **标准 Markdown 图片**：`![alt](images/file.jpg)` — 自动嵌入并保持原尺寸
2. **带图注的图片**：`![alt](path.jpg)<br>*▲ 图注文字*` — 自动嵌入图片 + 添加居中图注
3. **分行图注**：`![alt](path.jpg)` 下一行 `*图注文字*` — 自动识别绑定

### 工作原理

本技能采用**两阶段图片处理**策略：

1. **阶段一（主转换）**：`md_to_word.py` 在解析 Markdown 时尝试直接插入图片，优先查找绝对路径和相对路径
2. **阶段二（后处理）**：自动调用 `embed_images.py` 扫描生成的 DOCX 文件，将未能插入的 `[图片: ...]` 占位符替换为真实嵌入图片，并附上图注

### 图片查找顺序

```
1. Markdown 中指定的原始路径
2. MD 源文件同目录下的相对路径
3. MD 源文件同目录下的 images/ 子目录
4. 指定的 --img-dir 参数目录
```

## 项目文件

| 文件 | 说明 |
|------|------|
| `scripts/md_to_word.py` | 主转换脚本 — 解析 Markdown 生成 Word |
| `scripts/embed_images.py` | 图片后处理脚本 — 将占位符替换为真实图片 |
| `SKILL.md` | 本技能定义文件 |
| `LICENSE` | AGPL-3.0 许可证 |

## 工作流程

1. 读取用户指定的 Markdown 文件（或直接使用对话中的 Markdown 内容）
2. 用 `markdown-it-py` 解析为 token 流
3. 遍历 token，用 `python-docx` 构建对应的 Word 元素
4. 保存为 `.docx` 文件
5. **自动检测**是否包含图片引用，如有则执行图片后处理嵌入

## 参考实现

本技能参考了 Cherry Studio `src/main/services/ExportService.ts` 的实现思路：
- 使用 `markdown-it`（JS版）/ `markdown-it-py`（Python版）解析 Markdown
- 遍历 token 流构建文档元素
- 表格采用三线表风格（表头粗体、顶部底部有边框）

## 已知限制

- 表格中的嵌套复杂格式（如表格内代码块）支持有限
- 图片需要本地文件路径才能嵌入（不支持网络图片下载）
- 数学公式（LaTeX）暂不支持，建议在 Cherry Studio 中使用其内置导出功能
- 图片尺寸固定为5英寸宽度，后续可配置

## 开源协议

本技能基于 **GNU Affero General Public License v3.0 (AGPL-3.0)** 发布。

本项目参考了 [Cherry Studio](https://github.com/CherryHQ/cherry-studio)（AGPL-3.0）的 `ExportService.ts` 实现思路。
根据 AGPL-3.0 第5节的要求，特此声明：

- ✅ 本作品基于 Cherry Studio 的思路进行修改和移植
- ✅ 修改时间：2026年6月
- ✅ 修改内容：从 TypeScript (Electron) 移植到 Python 3，替换了所有依赖库
- ✅ 本作品同样以 AGPL-3.0 协议发布

详细信息请参阅 [LICENSE](./LICENSE) 文件。
