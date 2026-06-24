# ATS Skills

本仓库包含了自主式交通系统计算技术课题四算法模型相关的 Skills。

## 包含的 Skills

- `ats-algorithm-packaging`：将 Python 算法模型打包为标准化 Docker 格式，用于云计算平台部署，并生成ATS 算法模型“原始实验运行数据及测试报告”Markdown 和 Word 文档。（**仅专题4内部使用，打包+文档，其他专题可参考**）
- `ats-algorithm-doc-generation`：基于完整源代码、输入输出数据和真实运行证据，生成 ATS 算法模型“原始实验运行数据及测试报告”Markdown 和 Word 文档。（**仅文档生成，不进行打包**）
- `ats-algorithm-cover`：生成自主式交通系统（ATS）算法模型封面卡片，用于云计算平台展示。

## 使用方式

克隆仓库：

```bash
git clone https://github.com/jexxl/ats-skills.git
```

然后将 skill 目录复制或链接到对应 AI Agent 的 `skills` 目录中。
