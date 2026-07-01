# ATS Skills

本仓库包含了自主式交通系统（ATS）计算技术课题四算法模型相关的 Skills。

## 包含的 Skills

- `ats-algorithm-packaging`：将 Python 算法模型打包为标准化 Docker 格式，用于云计算平台部署，并生成ATS 算法模型“原始实验运行数据及测试报告”Markdown 和 Word 文档。（**仅专题4内部使用，打包+文档，其他专题可参考**）
- `ats-algorithm-doc-generation`：基于完整源代码、输入输出数据和真实运行证据，生成 ATS 算法模型“原始实验运行数据及测试报告”Markdown 和 Word 文档。（**供专题1、2、3参考，仅文档生成，不进行打包**）
- `ats-algorithm-cover`：生成 ATS 算法模型封面卡片，用于云计算平台展示。

## 仓库结构

所有可安装的 Skill 均放在 `skills/` 目录下：

```text
skills/
├── ats-algorithm-packaging/
├── ats-algorithm-doc-generation/
└── ats-algorithm-cover/
```

## 使用方式

推荐使用 `npx skills` 安装：

```bash
# 查看仓库中可安装的 Skills
npx skills add jexxl/ats-skills --list

# 安装全部 Skills
npx skills add jexxl/ats-skills --skill '*'
```

也可以只安装其中一个 Skill：

```bash
npx skills add jexxl/ats-skills --skill ats-algorithm-packaging
npx skills add jexxl/ats-skills --skill ats-algorithm-doc-generation
npx skills add jexxl/ats-skills --skill ats-algorithm-cover
```

执行安装时，根据 `npx skills` 的提示选择目标 Agent 和安装位置。如需手动安装，也可以克隆仓库后，将 `skills/` 下对应目录复制或链接到 AI Agent 的 `skills` 目录中。
