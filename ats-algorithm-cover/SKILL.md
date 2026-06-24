---
name: ats-algorithm-cover
description: 生成自主式交通系统（ATS）算法模型封面卡片。仅当用户要求创建、更新、渲染、批量生成或修复 ATS 算法封面/卡片 HTML，或透明高分辨率 PNG 图片时使用；所需封面数据可由用户直接提供，也可从文档或算法目录中推断。
---
# ATS 算法模型封面生成

## 目标

为 ATS 算法模型生成统一封面卡片：

- 输出 `cover.html` 和 `cover.png`
- HTML 从本 skill 内部模板 `assets/cover.html` 复制生成
- PNG 使用 Python + Playwright 通过 `uv run` 渲染，等待在线 Mermaid 加载完成
- PNG 默认透明背景、圆角外保留 alpha、DPI 元数据默认 300

## 必需数据

生成封面前必须具备以下字段。若用户没有直接提供，也没有提供可推断这些字段的文档或目录，先追问缺失项，不要编造。

1. 算法类型：基础算法或专用算法；也可提供可判断类型的算法编号，如 `4-4-J-1` 或 `4-4-Z-5`
2. 功能描述：一句不超过两行的模型主要功能描述，并指定或允许选择一个加粗关键词
3. 流程节点：2-6 个流程节点，或一段可转换为 Mermaid `flowchart LR` 的流程描述
4. 模型服务场景：用于页脚的服务场景文本
5. 主要编程语言：例如 `Python`、`C++`
6. 输出位置：目标 `cover.html` 和 `cover.png` 路径，或一个可写的算法目录

可选数据：

- 算法编号和算法名称，用于 HTML `<title>`
- 渲染参数：视口宽高、缩放倍率、DPI、是否保留不透明背景

## 数据推断

如果用户提供算法包目录、说明文档、Markdown frontmatter、表格或算法清单，先从这些材料中推断必需字段，再只追问无法推断的字段。

常见推断规则：

- 编号含 `-J-`：基础算法，`data-type="basic"`
- 编号含 `-Z-`：专用算法，`data-type="dedicated"`
- Markdown frontmatter 中的 `模型功能描述` 可压缩为功能句
- Markdown frontmatter 中的 `模型服务场景` 可直接用于页脚
- PlantUML 或条目式流程可压缩为 2-6 个 Mermaid 节点
- 语言可从源码、依赖文件、用户说明或算法包结构推断

推断后，如果字段会影响封面语义且不确定，应向用户确认。

## HTML 生成

从本 skill 的内部模板复制到目标位置。路径应相对本 `SKILL.md` 所在目录解析；使用当前平台的文件 API 或原生命令执行复制。

```text
<skill_dir>/assets/cover.html -> <target_dir>/cover.html
```

模板中只替换 `EDIT` 标记对应的 5 个位置：

1. `<section class="algorithm-card" data-type="...">`
2. `.function-text` 内的一句话功能描述和 `<strong>关键词</strong>`
3. `<pre class="mermaid">...</pre>` 内的 Mermaid 流程
4. `.service-scene` 内的模型服务场景
5. `.language` 内的主要编程语言 tag；多语言用多个 `<span class="language-tag">...</span>`，例如 `Python` 和 `C++`

不要改动模板的布局 CSS、卡片比例、主题变量、脚本结构，除非用户明确要求。模板内保留在线 Mermaid：

```html
<script src="https://cdn.jsdelivr.net/npm/mermaid@11/dist/mermaid.min.js"></script>
```

不要改成本地 Mermaid 文件，除非用户明确要求离线渲染。

## 渲染

使用本 skill 自带脚本，脚本路径相对本 `SKILL.md` 所在目录解析：

```bash
uv run <skill_dir>/scripts/render_cover.py <target_dir>/cover.html <target_dir>/cover.png
```

脚本参数：

- `--viewport-width` / `--viewport-height`: 逻辑视口，默认 `600x300`
- `--scale`: 设备缩放，默认 `3`
- `--dpi`: PNG DPI 元数据，默认 `300`
- `--opaque-background`: 保留页面背景；默认不加该参数，输出透明背景
- `--browser-channel`: 默认 `chrome`；如果目标环境没有 Chrome，可改用 `msedge` 或空字符串使用 Playwright 默认 Chromium

渲染脚本必须等待：

- `window.mermaid` 存在
- `.mermaid svg` 已生成并具有有效尺寸

## 验证

生成后至少检查：

1. `cover.png` 存在且分辨率不低于参考图 2x，推荐 3x
2. DPI 元数据为 300 或用户指定值
3. 四角 alpha 接近 0，圆角外不是实色背景
4. 流程图显示为 Mermaid SVG 节点和箭头，不是源码文本
5. 功能句不超过两行，页脚文字不溢出

可用跨平台 Python 检查尺寸、DPI 和角点 alpha：

```bash
uv run --with pillow python -c "from PIL import Image; im=Image.open('<target_dir>/cover.png'); print({'size': im.size, 'dpi': im.info.get('dpi'), 'top_left_alpha': im.convert('RGBA').getpixel((0, 0))[3]})"
```

## 批量生成

当用户要求批量生成封面时：

1. 先确认每张封面的必需数据来源：用户提供的数据表、多个算法目录、多个说明文档，或其他结构化材料
2. 对每个算法收集或推断 6 项必需数据
3. 生成每个 `cover.html`
4. 逐个运行渲染脚本生成 `cover.png`
5. 抽查基础算法和专用算法各至少一张，确认颜色、透明背景和 Mermaid 渲染
6. 汇报生成数量、失败项、缺失数据项和代表性尺寸
