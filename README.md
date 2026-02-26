# AI PPT Studio

AI 驱动的高质量 PPT 生成工具，支持结构化内容生成、预览编辑与一键导出 PPTX，适合课程汇报、答辩展示、项目路演等场景。

## 功能特性

- AI 生成结构化页面内容（JSON）
- 卡片式预览与 Markdown 视图切换
- 支持要点扩充/精简、拖拽排序
- 一键导出 PPTX（`pptxgenjs`）
- 多模型回退 + 超时兜底（含本地专家模板降级）
- 风格、听众、语气可配置

## 技术栈

- Frontend: HTML / CSS / Vanilla JavaScript
- Backend: Node.js + Express
- AI SDK: OpenAI SDK（兼容 ChatAnywhere 等 OpenAI 风格接口）
- PPT 导出: PptxGenJS

## 快速开始

### 1) 克隆项目

```bash
git clone https://github.com/dingchenyu957-ui/ai-ppt.git
cd ai-ppt
```

### 2) 安装依赖

```bash
npm install
```

### 3) 配置环境变量

在项目根目录创建 `.env` 文件：

```env
OPENAI_API_KEY=your_api_key_here

# 可选：自定义模型回退顺序（逗号分隔）
OPENAI_MODELS=gpt-5-mini,gpt-4.1-mini,gpt-4o-mini

# 可选：单模型超时（毫秒）
OPENAI_MODEL_TIMEOUT_MS=45000
```

### 4) 启动服务

```bash
node server.js
```

浏览器打开：

```text
http://localhost:3000
```

## 使用说明

1. 输入 PPT 主题  
2. 选择风格、听众、语气、页数  
3. 点击“生成高质量内容”  
4. 在右侧卡片区微调（扩充/精简、拖拽排序）  
5. 点击“下载 PPT”导出文件

## 常见问题

### 1) 生成慢或超时

- 检查网络与 API 服务稳定性
- 尝试减少页数（如 8-10 页）
- 调大 `OPENAI_MODEL_TIMEOUT_MS`
- 系统会在云端超时时自动切换本地专家模板

### 2) 导出失败

- 确认先成功生成内容后再导出
- 查看页面底部错误提示
- 检查后端日志输出

## 项目结构

```text
.
├── public/
│   └── index.html
├── server.js
├── package.json
└── README.md
```

## 许可证

本项目基于 [MIT License](./LICENSE) 开源。

