# WpsAgent

WpsAgent 是一个功能强大的 WPS Office 加载项，集成了 AI 功能来辅助用户处理文档。

## 功能特性

### AI 对话助手
- 支持云端 API（DeepSeek、OpenAI 等）
- 支持本地大模型（Ollama、vLLM、LocalAI 等）
- 支持 Markdown 格式渲染
- 对话历史管理

### 文档处理功能
| 功能 | 描述 |
|------|------|
| 校对 | 错别字检测，红底删除线+批注 |
| 润色 | 文档润色，支持全文分段处理 |
| 排版 | 智能排版，自定义字体/行距/边距 |
| 续写 | 根据上文自动续写 |
| 翻译 | 中英互译 |
| 总结 | 提取文档摘要 |
| 导航 | 嵌入豆包网页 |
| 写作 | 生成机关公文（通知/报告/请示/批复/函） |

## 快速开始

### 安装依赖

```bash
npm install
```

### 配置环境变量

编辑 `.env` 文件，设置你的模型配置：

#### 方式一：使用云端模型

```env
# 模型类型
MODEL_TYPE=cloud

# 云端模型配置
API_URL=https://api.deepseek.com/chat/completions
API_KEY=your-api-key-here
API_MODEL=deepseek-chat
```

支持的云端模型：
- **DeepSeek**: `deepseek-chat`, `deepseek-reasoner`, `deepseek-coder`
- **OpenAI**: `gpt-3.5-turbo`, `gpt-4`, `gpt-4-turbo`
- **Claude**: `claude-3-sonnet`, `claude-3-opus`

#### 方式二：使用本地大模型

```env
# 模型类型
MODEL_TYPE=local

# 本地模型配置
LOCAL_API_URL=http://localhost:11434/api/chat
LOCAL_MODEL_NAME=llama2
LOCAL_API_NEED_AUTH=false
LOCAL_API_KEY=
```

支持的本地模型框架：

| 框架 | 默认端口 | 配置示例 |
|------|----------|----------|
| **Ollama** | 11434 | `http://localhost:11434/api/chat` |
| **vLLM** | 8000 | `http://localhost:8000/v1/chat/completions` |
| **LocalAI** | 8080 | `http://localhost:8080/v1/chat/completions` |
| **text-generation-webui** | 5000 | `http://localhost:5000/v1/chat/completions` |

**Ollama 配置示例：**
```bash
# 安装 Ollama
curl -fsSL https://ollama.com/install.sh | sh

# 下载模型
ollama pull llama2
ollama pull qwen

# 启动服务
ollama serve
```

**vLLM 配置示例：**
```bash
# 安装 vLLM
pip install vllm

# 启动服务
python -m vllm.entrypoints.openai.api_server --model meta-llama/Llama-2-7b-chat-hf --port 8000
```

#### 排版配置

```env
# 页边距（厘米）
MARGIN_TOP=2.54
MARGIN_BOTTOM=2.54
MARGIN_LEFT=3.18
MARGIN_RIGHT=3.18

# 行距（磅）
LINE_SPACING=28

# 字体配置
TITLE_FONT=黑体
TITLE_FONT_SIZE=16
H1_FONT=黑体
H1_FONT_SIZE=16
H2_FONT=黑体
H2_FONT_SIZE=14
H3_FONT=楷体
H3_FONT_SIZE=12
BRACKET_FONT=仿宋
BRACKET_FONT_SIZE=12
```

### 启动服务

```bash
npm start
```

服务将运行在 `http://localhost:3889`

### 在 WPS 中加载插件

1. 打开 WPS Office
2. 进入 `文件` > `选项` > `加载项`
3. 点击 `添加`，选择 `dist/manifest.xml`
4. 重启 WPS Office

## 项目结构

```
WpsAgent-Deploy-choose-0213/
├── .env                    # 环境变量配置
├── .gitignore             # Git 忽略文件
├── package.json           # 项目配置
├── server.js             # Express 静态文件服务器
└── dist/                 # 构建输出目录
    ├── manifest.xml      # WPS 插件清单
    ├── ribbon.xml        # 功能区按钮配置
    ├── main.js          # 入口文件
    ├── index.html       # 主页面
    ├── js/             # JavaScript 模块
    │   ├── util.js     # 工具函数
    │   ├── ribbon.js   # Ribbon 按钮事件处理
    │   ├── config.js   # 配置管理
    │   ├── logger.js   # 日志系统
    │   ├── api.js      # API 调用
    │   ├── document.js # 文档操作
    │   ├── tabManager.js # Tab 管理
    │   ├── dialog.js   # 设置对话框逻辑
    │   └── taskpane.js # 任务面板核心逻辑
    ├── ui/            # UI 文件
    │   ├── taskpane.html
    │   └── dialog.html
    └── images/       # 图标资源
```

## 技术栈

| 技术 | 用途 |
|------|------|
| **Vite** | 构建工具 |
| **Express** | 静态文件服务器 |
| **WPS JSAPI** | WPS Office 接口调用 |
| **wpsjs** | WPS 插件开发框架 |
| **dotenv** | 环境变量管理 |

## 配置说明

### 模型配置

在设置面板中可以选择使用云端模型或本地大模型：

#### 云端模型配置
- **模型类型**: 选择"云端模型"
- **API URL**: AI API 端点地址（支持 OpenAI 兼容接口）
- **API Key**: 你的 API 密钥
- **模型名称**: 例如 `deepseek-chat`, `gpt-3.5-turbo`, `claude-3-sonnet`

#### 本地大模型配置
- **模型类型**: 选择"本地模型"
- **本地 API 地址**: 本地模型服务地址
  - Ollama: `http://localhost:11434/api/chat`
  - vLLM: `http://localhost:8000/v1/chat/completions`
  - LocalAI: `http://localhost:8080/v1/chat/completions`
- **模型名称**: 例如 `llama2`, `qwen`, `chatglm`, `mistral`
- **需要认证**: 如果本地服务需要 API Key，勾选此项
- **本地 API Key**: 本地模型的认证密钥（如果需要）

#### 快速切换模型

在设置面板中可以随时切换云端模型和本地模型，无需重启服务。配置会自动保存到 WPS PluginStorage 中。

### 排版配置

自定义文档排版标准：
- 页边距（上/下/左/右）
- 行距设置
- 标题字体和字号（1-3级）
- 括号内文字字体

## 开发

### 开发模式

```bash
npm run dev
```

### 构建

```bash
npm run build
```

### 代码检查

```bash
npm run lint
```

### 代码格式化

```bash
npm run format
```

## 安全性

- API Key 通过环境变量配置，不硬编码在代码中
- 敏感信息存储在 WPS PluginStorage 中
- 支持配置重置和恢复默认值

## 常见问题

### 1. 插件无法加载

确保：
- WPS Office 版本支持 JSAPI
- manifest.xml 路径正确
- 服务器正在运行

### 2. API 调用失败

检查：
- API Key 是否正确
- 网络连接是否正常
- API URL 是否有效

### 3. 功能按钮无响应

检查：
- 浏览器控制台是否有错误
- 服务器是否正常响应
- 配置是否正确加载

## 更新日志

### v1.2.0
- 支持本地大模型（Ollama、vLLM、LocalAI 等）
- 支持云端模型（DeepSeek、OpenAI、Claude 等）
- 添加模型类型切换功能
- 优化设置 UI，支持云端/本地模型配置
- 添加本地模型认证支持

### v1.1.0
- 添加环境变量支持
- 模块化代码结构
- 添加日志系统
- 优化错误处理
- 添加 API 重试机制

### v1.0.0
- 初始版本
- 基础 AI 功能
- 文档处理功能

## 许可证

MIT

## 贡献

欢迎提交 Issue 和 Pull Request！
