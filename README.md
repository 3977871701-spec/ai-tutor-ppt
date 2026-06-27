# 学院AI辅导员系统 - PPT生成工具

基于 Node.js 和 pptxgenjs 库的 PPT 自动生成工具，用于批量生成AI辅导员系统的项目介绍演示文稿。

## 项目简介

本项目使用 JavaScript 脚本通过 pptxgenjs 库自动生成学院AI辅导员系统的 PowerPoint 演示文稿。包含多个生成脚本，可生成不同版本和用途的PPT，避免手动排版的重复劳动。

## 功能特性

- **脚本化生成** - 代码定义PPT内容和样式，一键生成标准化演示文稿
- **多版本支持** - 提供基础版、客户版、完整版等多种生成脚本
- **统一配色** - 预定义学院蓝+科技蓝+橙色强调的配色方案
- **内容大纲** - 附带 PPT_CONTENT.md 内容大纲文档，方便内容策划
- **批量生成** - 可快速修改内容后重新生成，适合迭代优化

## 技术栈

| 组件 | 技术 |
|------|------|
| 运行环境 | Node.js |
| PPT生成 | pptxgenjs ^4.0.1 |

## PPT内容大纲

1. 项目背景与目标（痛点/趋势/解决方案）
2. 系统架构（RAG + LLM + 微信公众号）
3. 知识库内容（7大类别覆盖学生日常所需）
4. 核心功能演示（智能问答/多场景覆盖）
5. 技术方案详解
6. 部署与运维方案
7. 成本分析与预期收益

## 项目结构

```
ai-tutor-ppt/
├── generate.js                # 基础版PPT生成脚本
├── generate_client.js         # 客户版PPT生成脚本
├── generate_full.js           # 完整版PPT生成脚本
├── PPT_CONTENT.md             # PPT内容大纲
├── package.json               # Node.js依赖配置
├── package-lock.json
├── node_modules/
└── *.pptx                     # 已生成的PPT文件
```

## 使用方法

### 1. 安装依赖

```bash
cd /Users/xylei/work/projects/ai-tutor-ppt
npm install
```

### 2. 生成PPT

```bash
# 生成基础版
node generate.js

# 生成客户版
node generate_client.js

# 生成完整版
node generate_full.js
```

生成的 `.pptx` 文件将保存在当前目录下。

### 3. 自定义内容

编辑 `PPT_CONTENT.md` 规划内容大纲，然后修改对应脚本中的文本和样式参数，重新运行生成。

## License

MIT
