# Office-Word-MCP-Server 完整使用指南

> **版本**: v2.0
> **更新日期**: 2025年1月
> **适用平台**: Windows 10+、Linux (Ubuntu 18.04+ / CentOS 7+)
> **Python 版本**: 3.11+

---

## 目录

- [第1章：项目概述](#第1章项目概述)
  - [1.1 什么是 MCP](#11-什么是-mcp)
  - [1.2 项目架构](#12-项目架构)
  - [1.3 功能特性](#13-功能特性)
  - [1.4 适用场景](#14-适用场景)
- [第2章：MCP Server 安装](#第2章mcp-server-安装)
  - [2.1 Windows 环境安装](#21-windows-环境安装)
  - [2.2 Linux 环境安装](#22-linux-环境安装)
  - [2.3 传输方式选择](#23-传输方式选择)
- [第3章：MCP Client 配置](#第3章mcp-client-配置)
  - [3.1 Claude Code 配置](#31-claude-code-配置)
  - [3.2 VSCode + Cline 配置](#32-vscode--cline-配置)
  - [3.3 Spring AI Alibaba 配置](#33-spring-ai-alibaba-配置)
- [第4章：使用示例](#第4章使用示例)
  - [4.1 通过 Claude Code 使用](#41-通过-claude-code-使用)
  - [4.2 通过 Cline 使用](#42-通过-cline-使用)
  - [4.3 常用操作示例](#43-常用操作示例)
- [第5章：传输方式详解](#第5章传输方式详解)
  - [5.1 STDIO 传输](#51-stdio-传输)
  - [5.2 Streamable HTTP 传输](#52-streamable-http-传输)
  - [5.3 SSE 传输](#53-sse-传输)
- [第6章：故障排除](#第6章故障排除)
  - [6.1 安装问题](#61-安装问题)
  - [6.2 配置问题](#62-配置问题)
  - [6.3 连接问题](#63-连接问题)
- [第7章：企业级部署架构](#第7章企业级部署架构)
  - [7.1 B/S 部署架构](#71-bs-部署架构)
  - [7.2 C/S 部署架构](#72-cs-部署架构)
  - [7.3 架构选型指南](#73-架构选型指南)

---

## 第1章：项目概述

### 1.1 什么是 MCP

**MCP (Model Context Protocol)** 是一种标准化协议，用于 AI 应用与外部工具/服务的通信。

#### MCP 架构关系

```
┌─────────────────────────────────────────────────┐
│                                                 │
│  MCP Client/Host（客户端/宿主）                   │
│  ├─ Claude Code（CLI 工具）                      │
│  ├─ Claude Desktop（桌面应用）                   │
│  └─ VSCode + Cline（IDE 扩展）                   │
│                                                 │
│  作用：发起请求，调用工具，展示结果                 │
│                                                 │
└──────────────────┬──────────────────────────────┘
                   │
                   │ MCP Protocol（协议通信）
                   │ Transport: STDIO / HTTP / SSE
                   │
┌──────────────────▼──────────────────────────────┐
│                                                 │
│  MCP Server（服务端）                            │
│  └─ Office-Word-MCP-Server ← 本项目              │
│                                                 │
│  作用：提供具体功能（Word文档操作）                 │
│                                                 │
└──────────────────┬──────────────────────────────┘
                   │
                   ▼
          Word 文档文件（.docx）
```

### 1.2 项目架构

**Office-Word-MCP-Server** 是一个 **MCP Server**（服务端），提供 Microsoft Word 文档操作能力。

#### 核心组件

```
Office-Word-MCP-Server/
├── word_document_server/      # 核心服务模块
│   ├── main.py                # 服务入口
│   ├── tools/                 # 工具实现
│   └── utils/                 # 工具函数
├── word_mcp_server.py         # 启动脚本
├── setup_mcp.py               # 安装配置脚本
├── requirements.txt           # Python 依赖
├── pyproject.toml             # 项目配置
└── .env.example               # 环境变量模板
```

#### 文档处理原理 ⚠️ 重要概念

##### 核心要点：文档在哪里生成和处理？

**答案：所有文档操作都在 MCP Server 所在的服务器执行，文档文件保存在 MCP Server 的本地文件系统。**

##### 技术实现原理

本项目使用 **python-docx** 库操作 Word 文档，该库的工作方式决定了文档必须在 MCP Server 端处理：

```python
# word_document_server/tools/document_tools.py
async def create_document(filename: str, title: str = None, author: str = None):
    # 1. 在 MCP Server 的本地文件系统创建文档
    doc = Document()

    # 2. 设置文档属性
    if title:
        doc.core_properties.title = title

    # 3. 保存到 MCP Server 的本地磁盘
    doc.save(filename)  # ← 文件保存在 Server 端

    return f"Document {filename} created successfully"
```

##### 跨服务器部署场景说明

**场景：MCP Server 部署在服务器A，Spring AI（MCP Client）部署在服务器B**

```
服务器B（192.168.1.100）          服务器A（192.168.1.101）
┌──────────────────────┐          ┌──────────────────────┐
│  Spring AI Alibaba   │          │  MCP Server          │
│  (MCP Client)        │          │                      │
│                      │          │  ┌────────────────┐  │
│  1. 发送请求         │          │  │  Python 进程   │  │
│     create_document()│──HTTP──▶│  │  python-docx   │  │
│                      │          │  └────────┬───────┘  │
│                      │          │           │          │
│  3. 接收结果         │          │  2. 创建文档          │
│     "创建成功"        │◀─HTTP───│     ↓                │
│                      │          │  [report.docx]       │
│                      │          │  保存在 A 的磁盘     │
└──────────────────────┘          └──────────────────────┘
```

##### 关键说明

| 问题 | 答案 |
|------|------|
| **文档在哪里处理？** | MCP Server 所在的服务器（服务器A） |
| **文档保存在哪里？** | MCP Server 的本地文件系统（服务器A） |
| **MCP Client 能直接访问文档吗？** | ❌ 不能，Client 只接收文本结果（如"创建成功"） |
| **filename 参数是相对于谁？** | 相对于 MCP Server 的工作目录 |
| **python-docx 安装在哪里？** | MCP Server 所在的服务器 |

##### 为什么这样设计？

这是 **MCP 协议的标准设计模式**：

- **Server（服务器A）** = 能力提供者
  - 拥有文档处理能力（python-docx）
  - 访问本地文件系统
  - 执行实际操作

- **Client（服务器B）** = 能力消费者
  - 通过 MCP 协议调用工具
  - 只接收文本结果（JSON/String）
  - **不直接访问文档文件**

##### 实际影响和解决方案

如果需要在服务器B使用生成的文档，需要：

| 方案 | 实现方式 | 适用场景 |
|------|----------|----------|
| **文件共享** | 在服务器A设置 NFS/SMB | 内网环境 |
| **对象存储** | 文档上传至 MinIO/阿里云OSS | 生产环境 |
| **HTTP 下载** | 扩展 MCP Server 添加下载接口 | 需要细粒度控制 |

详细的部署架构和文件分发方案请参考 [第7章：企业级部署架构](#第7章企业级部署架构)。

---

### 1.3 功能特性

#### ✅ 文档管理
- 创建、读取、复制、合并 Word 文档
- 提取文本和分析文档结构
- 查看文档属性和统计信息
- 转换为 PDF 格式

#### ✅ 内容创建
- 添加标题、段落、表格、图片
- 插入分页符、列表（有序/无序）
- 添加脚注和尾注
- 相对位置插入内容

#### ✅ 格式化功能
- 文本格式化（粗体、斜体、下划线、颜色）
- 表格格式化（边框、底纹、对齐）
- 单元格合并、列宽调整
- 自定义样式创建

#### ✅ 高级功能
- 文档密码保护
- 数字签名
- 注释提取和管理
- 搜索和替换

### 1.4 适用场景

| 场景 | 说明 | 示例 |
|------|------|------|
| **自动化报告生成** | 批量创建结构化文档 | 周报、月报、销售报表 |
| **文档模板处理** | 基于模板填充数据 | 合同、发票、证书 |
| **AI 辅助写作** | 通过 AI 助手操作文档 | 使用 Claude 编写技术文档 |
| **批量文档处理** | 格式化、转换、合并文档 | PDF 转换、文档标准化 |

---

## 第2章：MCP Server 安装

### 2.1 Windows 环境安装

#### 系统要求

| 项目 | 要求 |
|------|------|
| 操作系统 | Windows 10 或更高版本 |
| Python | 3.11 或更高版本 |
| 内存 | 最少 512MB 可用内存 |
| 存储 | 最少 200MB 可用空间 |

#### 前置准备

##### 1. 安装 Python

1. 访问 [python.org](https://www.python.org/downloads/) 下载 Python 3.11+
2. 安装时**务必勾选** "Add Python to PATH"
3. 验证安装：

```cmd
python --version
# 输出示例: Python 3.11.5

pip --version
# 输出示例: pip 23.2.1
```

##### 2. 安装 Git（可选）

```cmd
# 下载安装 Git for Windows
# https://git-scm.com/download/win

git --version
# 输出示例: git version 2.42.0
```

#### 方法A：自动安装（推荐）

##### 步骤1：克隆项目

```cmd
# 使用 Git 克隆
git clone https://github.com/GongRzhe/Office-Word-MCP-Server.git
cd Office-Word-MCP-Server

# 或者下载 ZIP 后解压
# https://github.com/GongRzhe/Office-Word-MCP-Server/archive/refs/heads/main.zip
```

##### 步骤2：运行安装脚本

```cmd
python setup_mcp.py
```

##### 步骤3：按提示选择配置

```
Word Document MCP Server Setup (Multi-Transport)
===============================================

Transport Configuration:
1. STDIO (default, local execution)
2. Streamable HTTP (modern, recommended for web deployment)
3. SSE (Server-Sent Events, for compatibility)

Select transport type (1-3, default: 1): 1  ← 输入选择
```

**选择建议**：
- 选择 `1` (STDIO)：用于 Claude Code、Claude Desktop（推荐）
- 选择 `2` (HTTP)：用于网络访问或远程部署
- 选择 `3` (SSE)：用于特殊兼容性需求

##### 步骤4：选择安装方式

选择传输方式后，脚本会检测 `word-document-server` 是否已安装，并提示：

```
word-document-server is not installed.

Options:
1. Install from PyPI (recommended)
2. Set up local development environment

Enter your choice (1-2):
```

**选项说明**：

| 选项 | 说明 | 适用场景 |
|------|------|----------|
| **1. Install from PyPI** | 从 PyPI 安装已发布的包 | ✅ 普通用户<br>✅ 快速安装<br>✅ 生产环境 |
| **2. Local development** | 创建虚拟环境，本地开发 | ✅ 开发者<br>✅ 需要修改代码<br>✅ 调试和测试 |

**推荐选择**：

- **普通用户**：选择 `1`（从 PyPI 安装）
  - 优点：安装快速、版本稳定、自动管理依赖
  - 缺点：不能修改源代码

- **开发者**：选择 `2`（本地开发环境）
  - 优点：可以修改代码、方便调试、隔离依赖
  - 缺点：需要手动管理虚拟环境

**示例：选择 PyPI 安装（推荐）**

```
Enter your choice (1-2): 1

Installing word-document-server from PyPI...
Successfully installed word-mcp-server!

Now generating MCP config...
MCP configuration has been written to: D:\...\mcp-config.json
```

**示例：选择本地开发环境**

```
Enter your choice (1-2): 2

Creating new virtual environment...
Virtual environment created successfully!

Installing requirements...
Requirements installed successfully!

MCP configuration has been written to: D:\...\mcp-config.json
```

##### 安装完成提示

无论选择哪种方式，安装成功后都会显示：

```
Setup complete! You can now use the Word Document MCP server with compatible clients like Claude Desktop.

Transport Summary:
  - Transport: stdio
  - Configuration file: D:\...\mcp-config.json
```

#### 方法B：手动安装

##### 步骤1：创建虚拟环境

```cmd
# 进入项目目录
cd Office-Word-MCP-Server

# 创建虚拟环境
python -m venv .venv

# 激活虚拟环境
.venv\Scripts\activate

# 提示符会变为：
# (.venv) D:\...\Office-Word-MCP-Server>
```

##### 步骤2：安装依赖

```cmd
# 升级 pip
python -m pip install --upgrade pip

# 安装项目依赖
pip install -r requirements.txt

# 验证安装
pip list
```

预期看到的关键包：

```
Package              Version
-------------------- -------
fastmcp              2.8.1
python-docx          1.1.2
msoffcrypto-tool     5.4.2
docx2pdf             0.1.8
```

##### 步骤3：测试运行

```cmd
# 测试服务器启动
python word_mcp_server.py

# 如果看到类似输出，表示成功：
# Starting Word Document Server...
```

按 `Ctrl+C` 停止服务器。

#### 环境变量配置

创建 `.env` 文件（可选）：

```cmd
# 复制模板
copy .env.example .env

# 编辑 .env 文件
notepad .env
```

`.env` 文件内容：

```env
# 传输方式配置
MCP_TRANSPORT=stdio

# HTTP/SSE 配置（当不使用 stdio 时）
MCP_HOST=127.0.0.1
MCP_PORT=8000

# Streamable HTTP 路径
MCP_PATH=/mcp

# SSE 路径
MCP_SSE_PATH=/sse
```

---

### 2.2 Linux 环境安装

#### 系统要求

| 发行版 | 版本要求 |
|--------|----------|
| Ubuntu/Debian | 18.04+ / Debian 10+ |
| CentOS/RHEL | 7+ |
| Python | 3.11+ |

#### Ubuntu/Debian 安装

##### 步骤1：更新系统并安装依赖

```bash
# 更新包索引
sudo apt update
sudo apt upgrade -y

# 安装 Python 3.11+
sudo apt install -y python3.11 python3.11-venv python3-pip git

# 验证安装
python3.11 --version
# 输出: Python 3.11.x
```

##### 步骤2：克隆项目

```bash
# 克隆仓库
git clone https://github.com/GongRzhe/Office-Word-MCP-Server.git
cd Office-Word-MCP-Server
```

##### 步骤3：运行安装脚本（自动安装）

```bash
python3.11 setup_mcp.py
```

##### 步骤4：选择配置

**第一步：选择传输方式**

```
Transport Configuration:
1. STDIO (default, local execution)
2. Streamable HTTP (modern, recommended for web deployment)
3. SSE (Server-Sent Events, for compatibility)

Select transport type (1-3, default: 1): 1
```

推荐选择 `1` (STDIO)。

**第二步：选择安装方式**

```
word-document-server is not installed.

Options:
1. Install from PyPI (recommended)
2. Set up local development environment

Enter your choice (1-2): 1
```

- 普通用户选择 `1`（从 PyPI 安装，推荐）
- 开发者选择 `2`（本地开发环境）

详细说明见 [Windows 安装部分的步骤4](#步骤4选择安装方式)。

##### 步骤5：手动安装（可选）

```bash
# 创建虚拟环境
python3.11 -m venv .venv

# 激活虚拟环境
source .venv/bin/activate

# 升级 pip
pip install --upgrade pip

# 安装依赖
pip install -r requirements.txt

# 测试运行
python word_mcp_server.py
```

#### CentOS/RHEL 安装

##### 步骤1：安装开发工具和 Python

```bash
# 安装开发工具
sudo yum groupinstall -y "Development Tools"

# 安装 Python 3.11
sudo yum install -y python311 python311-devel python311-pip git

# 验证
python3.11 --version
```

##### 步骤2：克隆并安装

```bash
git clone https://github.com/GongRzhe/Office-Word-MCP-Server.git
cd Office-Word-MCP-Server

# 自动安装
python3.11 setup_mcp.py

# 或手动安装
python3.11 -m venv .venv
source .venv/bin/activate
pip install -r requirements.txt
```

#### Linux 环境变量

```bash
# 编辑 .env 文件
cp .env.example .env
nano .env

# 或使用 vim
vim .env
```

#### 设置为系统服务（可选）

创建 systemd 服务：

```bash
sudo nano /etc/systemd/system/word-mcp-server.service
```

内容：

```ini
[Unit]
Description=Office Word MCP Server
After=network.target

[Service]
Type=simple
User=yourusername
WorkingDirectory=/path/to/Office-Word-MCP-Server
Environment="PATH=/path/to/Office-Word-MCP-Server/.venv/bin"
ExecStart=/path/to/Office-Word-MCP-Server/.venv/bin/python word_mcp_server.py
Restart=always

[Install]
WantedBy=multi-user.target
```

启动服务：

```bash
sudo systemctl daemon-reload
sudo systemctl enable word-mcp-server
sudo systemctl start word-mcp-server
sudo systemctl status word-mcp-server
```

---

### 2.3 传输方式选择

#### 选择流程图

```
开始
  │
  ├─ 用于本地 AI 工具（Claude Code/Desktop）？
  │   ├─ 是 → 选择 STDIO（推荐）
  │   └─ 否 ↓
  │
  ├─ 需要网络访问或远程部署？
  │   ├─ 是 → 选择 Streamable HTTP
  │   └─ 否 ↓
  │
  └─ 有特殊兼容性需求？
      ├─ 是 → 选择 SSE
      └─ 否 → 选择 STDIO（默认）
```

#### 传输方式对比

| 传输方式 | 适用场景 | 优点 | 缺点 |
|----------|----------|------|------|
| **STDIO** | 本地 AI 工具 | ✅ 配置简单<br>✅ 性能最佳<br>✅ 延迟最低 | ❌ 仅限本地 |
| **Streamable HTTP** | 网络/远程访问 | ✅ 支持网络<br>✅ 现代协议<br>✅ 易于集成 | ❌ 需要配置网络参数 |
| **SSE** | 兼容性需求 | ✅ 兼容性好<br>✅ 支持网络 | ❌ 技术较旧<br>❌ 性能不如 HTTP |

#### 配置示例

**STDIO 配置**（.env）：

```env
MCP_TRANSPORT=stdio
```

**HTTP 配置**（.env）：

```env
MCP_TRANSPORT=streamable-http
MCP_HOST=127.0.0.1
MCP_PORT=8000
MCP_PATH=/mcp
```

**SSE 配置**（.env）：

```env
MCP_TRANSPORT=sse
MCP_HOST=127.0.0.1
MCP_PORT=8000
MCP_SSE_PATH=/sse
```

---

## 第3章：MCP Client 配置

### 3.1 Claude Code 配置

#### 什么是 Claude Code

**Claude Code** 是 Anthropic 官方推出的 CLI 工具，可以：
- 在终端或 VSCode 终端中运行
- 通过 MCP 协议连接外部工具
- 执行代码编辑、文件操作等任务

#### 安装 Claude Code

##### 使用 npm 安装

```bash
# 全局安装
npm install -g @anthropic-ai/claude-code

# 验证安装
claude-code --version
```

##### 使用 npx（无需安装）

```bash
# 直接运行
npx @anthropic-ai/claude-code
```

#### 配置方式

Claude Code 支持两种配置方式：

| 配置方式 | 位置 | 作用范围 |
|----------|------|----------|
| **项目级配置** | `.claude/mcp.json` | 仅当前项目 |
| **全局配置** | `~/.config/claude/mcp.json` | 所有项目 |

**推荐**：使用**项目级配置**，便于项目移植和共享。

#### 项目级配置（推荐）

##### Windows 配置示例

**步骤1**：在项目根目录创建 `.claude` 目录

```cmd
cd D:\02_Dev\Workspace\GitHub\Office-Word-MCP-Server
mkdir .claude
```

**步骤2**：创建 `mcp.json` 文件

```cmd
notepad .claude\mcp.json
```

**步骤3**：添加以下内容

```json
{
  "mcpServers": {
    "word-document-server": {
      "command": "python",
      "args": [
        "D:\\02_Dev\\Workspace\\GitHub\\Office-Word-MCP-Server\\word_mcp_server.py"
      ],
      "env": {
        "PYTHONPATH": "D:\\02_Dev\\Workspace\\GitHub\\Office-Word-MCP-Server",
        "MCP_TRANSPORT": "stdio"
      }
    }
  }
}
```

**⚠️ 重要提示**：
- 路径使用**双反斜杠** `\\` 或单正斜杠 `/`
- 将路径替换为你的实际项目路径

##### 使用虚拟环境的配置

如果使用虚拟环境，推荐指定虚拟环境中的 Python：

```json
{
  "mcpServers": {
    "word-document-server": {
      "command": "D:\\02_Dev\\Workspace\\GitHub\\Office-Word-MCP-Server\\.venv\\Scripts\\python.exe",
      "args": [
        "D:\\02_Dev\\Workspace\\GitHub\\Office-Word-MCP-Server\\word_mcp_server.py"
      ],
      "env": {
        "PYTHONPATH": "D:\\02_Dev\\Workspace\\GitHub\\Office-Word-MCP-Server",
        "MCP_TRANSPORT": "stdio"
      }
    }
  }
}
```

##### Linux 配置示例

**步骤1**：创建配置目录和文件

```bash
cd /home/user/Office-Word-MCP-Server
mkdir -p .claude
nano .claude/mcp.json
```

**步骤2**：添加配置内容

```json
{
  "mcpServers": {
    "word-document-server": {
      "command": "/home/user/Office-Word-MCP-Server/.venv/bin/python",
      "args": [
        "/home/user/Office-Word-MCP-Server/word_mcp_server.py"
      ],
      "env": {
        "PYTHONPATH": "/home/user/Office-Word-MCP-Server",
        "MCP_TRANSPORT": "stdio"
      }
    }
  }
}
```

#### 全局配置

##### Windows 全局配置

配置文件位置：`%USERPROFILE%\.config\claude\mcp.json`

```cmd
# 创建目录
mkdir %USERPROFILE%\.config\claude

# 编辑配置
notepad %USERPROFILE%\.config\claude\mcp.json
```

配置内容（同项目级配置）。

##### Linux/macOS 全局配置

配置文件位置：`~/.config/claude/mcp.json`

```bash
# 创建目录
mkdir -p ~/.config/claude

# 编辑配置
nano ~/.config/claude/mcp.json
```

#### 验证配置

##### 启动 Claude Code

```bash
# 进入项目目录
cd Office-Word-MCP-Server

# 启动 Claude Code
claude-code
```

##### 测试连接

在 Claude Code 中输入：

```
请列出可用的 MCP 工具
```

如果配置成功，会看到类似输出：

```
可用的 MCP Server:
- word-document-server

可用的工具：
- create_document
- add_paragraph
- add_heading
- add_table
...
```

##### 测试创建文档

```
请创建一个名为 test.docx 的 Word 文档，添加标题"测试文档"和一段文字"这是测试内容"
```

成功执行后，会在当前目录生成 `test.docx` 文件。

#### HTTP 传输配置（可选）

如果使用 HTTP 传输，修改 `env` 部分：

```json
{
  "mcpServers": {
    "word-document-server": {
      "command": "python",
      "args": ["D:\\...\\word_mcp_server.py"],
      "env": {
        "PYTHONPATH": "D:\\...\\Office-Word-MCP-Server",
        "MCP_TRANSPORT": "streamable-http",
        "MCP_HOST": "127.0.0.1",
        "MCP_PORT": "8000",
        "MCP_PATH": "/mcp"
      }
    }
  }
}
```

---

### 3.2 VSCode + Cline 配置

#### 什么是 Cline

**Cline** 是一个 VSCode 扩展，提供：
- AI 驱动的代码助手
- 自动化文件操作
- 原生支持 MCP 协议
- 集成 UI 界面

**官网**：[https://docs.cline.bot](https://docs.cline.bot)

#### 安装 Cline 扩展

##### 方法1：从 VSCode Marketplace 安装

1. 打开 VSCode
2. 点击左侧扩展图标（或按 `Ctrl+Shift+X`）
3. 搜索 "Cline"
4. 点击 "Install" 安装

##### 方法2：通过命令行安装

```bash
code --install-extension saoudrizwan.cline
```

##### 验证安装

安装后，VSCode 左侧会出现 Cline 图标。

#### 配置 MCP Server

##### 步骤1：打开 Cline MCP 配置

1. 点击 VSCode 左侧的 **Cline 图标**
2. 在 Cline 面板顶部，点击 **"MCP Servers"** 图标
3. 选择 **"Configure"** 标签
4. 点击底部的 **"Advanced MCP Settings"** 链接

这会打开 `cline_mcp_settings.json` 配置文件。

##### 步骤2：添加 Word MCP Server 配置

**Windows 配置示例**：

```json
{
  "mcpServers": {
    "word-document-server": {
      "command": "python",
      "args": [
        "D:\\02_Dev\\Workspace\\GitHub\\Office-Word-MCP-Server\\word_mcp_server.py"
      ],
      "env": {
        "PYTHONPATH": "D:\\02_Dev\\Workspace\\GitHub\\Office-Word-MCP-Server",
        "MCP_TRANSPORT": "stdio"
      },
      "alwaysAllow": [],
      "disabled": false
    }
  }
}
```

**Linux 配置示例**：

```json
{
  "mcpServers": {
    "word-document-server": {
      "command": "/home/user/Office-Word-MCP-Server/.venv/bin/python",
      "args": [
        "/home/user/Office-Word-MCP-Server/word_mcp_server.py"
      ],
      "env": {
        "PYTHONPATH": "/home/user/Office-Word-MCP-Server",
        "MCP_TRANSPORT": "stdio"
      },
      "alwaysAllow": [],
      "disabled": false
    }
  }
}
```

##### 配置字段说明

| 字段 | 说明 | 示例值 |
|------|------|--------|
| `command` | Python 解释器路径 | `python` 或绝对路径 |
| `args` | 服务器脚本路径（数组） | `["path/to/word_mcp_server.py"]` |
| `env` | 环境变量 | `{"MCP_TRANSPORT": "stdio"}` |
| `alwaysAllow` | 自动允许的工具（数组） | `[]` 或 `["create_document"]` |
| `disabled` | 是否禁用此服务器 | `false` |

##### 步骤3：重启 MCP Server

保存配置后：

1. 返回 Cline 的 **"MCP Servers"** 面板
2. 找到 `word-document-server`
3. 点击 **"Restart"** 按钮（循环箭头图标）

##### 步骤4：验证连接

在 **"Installed"** 标签中，应该看到：

```
✅ word-document-server
   Status: Connected
   Tools: 30+ available
```

#### 通过 UI 添加 Server（简化方式）

Cline 还支持通过 UI 添加 MCP Server：

1. 在 **"MCP Servers"** 面板，点击 **"+"** 按钮
2. 选择 **"Add STDIO Server"**
3. 填写信息：
   - **Name**: `word-document-server`
   - **Command**: `python`（或完整路径）
   - **Args**: `D:\...\word_mcp_server.py`
   - **Env**: 点击 "Add Variable"
     - Key: `MCP_TRANSPORT`
     - Value: `stdio`
4. 点击 **"Save"**

#### 验证和测试

##### 测试 1：查看可用工具

1. 打开 Cline 聊天界面
2. 输入：

```
列出所有可用的 Word 文档操作工具
```

##### 测试 2：创建文档

```
请创建一个 Word 文档 demo.docx，包含：
- 标题：测试文档
- 段落1：这是第一段内容
- 段落2：这是第二段内容
```

成功后，项目目录会生成 `demo.docx`。

#### HTTP/SSE 传输配置

如果使用 HTTP 或 SSE 传输，修改配置：

**HTTP 示例**：

```json
{
  "mcpServers": {
    "word-document-server": {
      "url": "http://127.0.0.1:8000/mcp",
      "headers": {
        "Content-Type": "application/json"
      },
      "alwaysAllow": [],
      "disabled": false
    }
  }
}
```

**SSE 示例**：

```json
{
  "mcpServers": {
    "word-document-server": {
      "url": "http://127.0.0.1:8000/sse",
      "headers": {},
      "alwaysAllow": [],
      "disabled": false
    }
  }
}
```

#### Cline 配置文件位置

| 系统 | 配置文件路径 |
|------|--------------|
| **Windows** | `%APPDATA%\Code\User\globalStorage\saoudrizwan.cline\settings\cline_mcp_settings.json` |
| **Linux** | `~/.config/Code/User/globalStorage/saoudrizwan.cline/settings/cline_mcp_settings.json` |
| **macOS** | `~/Library/Application Support/Code/User/globalStorage/saoudrizwan.cline/settings/cline_mcp_settings.json` |

---

### 3.3 Spring AI Alibaba 配置

#### 什么是 Spring AI Alibaba

**Spring AI Alibaba** 是阿里巴巴开源的 AI 应用开发框架，提供：
- 统一的 AI 服务调用接口
- 原生支持 MCP (Model Context Protocol) 协议
- 与 Spring Boot 无缝集成
- 支持多种 AI 模型和工具调用

**官网**: [https://sca.aliyun.com/ai/](https://sca.aliyun.com/ai/)

**GitHub**: [https://github.com/alibaba/spring-ai-alibaba](https://github.com/alibaba/spring-ai-alibaba)

#### 环境要求

| 项目 | 要求 |
|------|------|
| Java | JDK 17 或更高版本 |
| Spring Boot | 3.2.0 或更高版本 |
| Maven/Gradle | Maven 3.6+ / Gradle 7.0+ |
| Python | 3.11+ (用于运行 MCP Server) |

#### 项目依赖配置

##### Maven 配置

在 `pom.xml` 中添加依赖：

```xml
<dependencies>
    <!-- Spring AI Alibaba Starter -->
    <dependency>
        <groupId>com.alibaba.cloud.ai</groupId>
        <artifactId>spring-ai-alibaba-starter</artifactId>
        <version>1.0.0-M2</version>
    </dependency>

    <!-- Spring Boot Starter Web -->
    <dependency>
        <groupId>org.springframework.boot</groupId>
        <artifactId>spring-boot-starter-web</artifactId>
    </dependency>

    <!-- Lombok (可选，简化代码) -->
    <dependency>
        <groupId>org.projectlombok</groupId>
        <artifactId>lombok</artifactId>
        <optional>true</optional>
    </dependency>
</dependencies>
```

##### Gradle 配置

在 `build.gradle` 中添加依赖：

```gradle
dependencies {
    // Spring AI Alibaba Starter
    implementation 'com.alibaba.cloud.ai:spring-ai-alibaba-starter:1.0.0-M2'

    // Spring Boot Starter Web
    implementation 'org.springframework.boot:spring-boot-starter-web'

    // Lombok (可选)
    compileOnly 'org.projectlombok:lombok'
    annotationProcessor 'org.projectlombok:lombok'
}
```

#### MCP Server 连接配置

##### 配置方式一：使用 STDIO 传输（推荐）

在 `application.yml` 中配置：

```yaml
spring:
  ai:
    alibaba:
      mcp:
        servers:
          # Word 文档 MCP Server 配置
          word-document-server:
            # 使用 STDIO 传输
            transport: stdio
            # Python 解释器路径
            command: python
            # MCP Server 启动脚本路径（需要修改为实际路径）
            args:
              - D:/02_Dev/Workspace/GitHub/Office-Word-MCP-Server/word_mcp_server.py
            # 环境变量
            env:
              PYTHONPATH: D:/02_Dev/Workspace/GitHub/Office-Word-MCP-Server
              MCP_TRANSPORT: stdio
            # 是否启用该 Server
            enabled: true
```

**Linux 配置示例**：

```yaml
spring:
  ai:
    alibaba:
      mcp:
        servers:
          word-document-server:
            transport: stdio
            command: /home/user/Office-Word-MCP-Server/.venv/bin/python
            args:
              - /home/user/Office-Word-MCP-Server/word_mcp_server.py
            env:
              PYTHONPATH: /home/user/Office-Word-MCP-Server
              MCP_TRANSPORT: stdio
            enabled: true
```

##### 配置方式二：使用 HTTP 传输

**步骤1**：启动 MCP Server（HTTP 模式）

```bash
# 修改 .env 文件
MCP_TRANSPORT=streamable-http
MCP_HOST=127.0.0.1
MCP_PORT=8000
MCP_PATH=/mcp

# 启动 Server
python word_mcp_server.py
```

**步骤2**：配置 `application.yml`

```yaml
spring:
  ai:
    alibaba:
      mcp:
        servers:
          word-document-server:
            # 使用 HTTP 传输
            transport: http
            # MCP Server HTTP 地址
            url: http://127.0.0.1:8000/mcp
            # 超时设置（毫秒）
            timeout: 30000
            enabled: true
```

##### 配置方式三：使用 SSE 传输

**步骤1**：启动 MCP Server（SSE 模式）

```bash
# 修改 .env 文件
MCP_TRANSPORT=sse
MCP_HOST=127.0.0.1
MCP_PORT=8000
MCP_SSE_PATH=/sse

# 启动 Server
python word_mcp_server.py
```

**步骤2**：配置 `application.yml`

```yaml
spring:
  ai:
    alibaba:
      mcp:
        servers:
          word-document-server:
            # 使用 SSE 传输
            transport: sse
            # MCP Server SSE 地址
            url: http://127.0.0.1:8000/sse
            timeout: 30000
            enabled: true
```

#### 配置字段说明

| 字段 | 说明 | 必填 | 示例值 |
|------|------|------|--------|
| `transport` | 传输方式 | 是 | `stdio` / `http` / `sse` |
| `command` | Python 解释器路径 | STDIO 模式必填 | `python` 或完整路径 |
| `args` | 启动脚本参数 | STDIO 模式必填 | `["/path/to/word_mcp_server.py"]` |
| `env` | 环境变量 | 可选 | `PYTHONPATH`, `MCP_TRANSPORT` 等 |
| `url` | HTTP/SSE 地址 | HTTP/SSE 模式必填 | `http://127.0.0.1:8000/mcp` |
| `timeout` | 超时时间（毫秒） | 可选 | `30000` |
| `enabled` | 是否启用 | 可选 | `true` / `false` |

#### Java 代码示例

##### 示例1：注入 MCP 工具并使用

创建服务类 `WordDocumentService.java`：

```java
package com.example.demo.service;

import com.alibaba.cloud.ai.mcp.McpToolClient;
import org.springframework.stereotype.Service;
import java.util.Map;

@Service
public class WordDocumentService {

    private final McpToolClient mcpToolClient;

    public WordDocumentService(McpToolClient mcpToolClient) {
        this.mcpToolClient = mcpToolClient;
    }

    /**
     * 创建 Word 文档
     */
    public String createDocument(String filename, String title, String author) {
        Map<String, Object> params = Map.of(
            "filename", filename,
            "title", title,
            "author", author
        );

        return mcpToolClient.executeTool(
            "word-document-server",  // MCP Server 名称
            "create_document",        // 工具名称
            params                    // 参数
        );
    }

    /**
     * 添加段落
     */
    public String addParagraph(String filename, String text, String fontName, int fontSize) {
        Map<String, Object> params = Map.of(
            "filename", filename,
            "text", text,
            "font_name", fontName,
            "font_size", fontSize
        );

        return mcpToolClient.executeTool(
            "word-document-server",
            "add_paragraph",
            params
        );
    }

    /**
     * 添加标题
     */
    public String addHeading(String filename, String text, int level) {
        Map<String, Object> params = Map.of(
            "filename", filename,
            "text", text,
            "level", level
        );

        return mcpToolClient.executeTool(
            "word-document-server",
            "add_heading",
            params
        );
    }

    /**
     * 添加表格
     */
    public String addTable(String filename, int rows, int cols, String[][] data) {
        Map<String, Object> params = Map.of(
            "filename", filename,
            "rows", rows,
            "cols", cols,
            "data", data
        );

        return mcpToolClient.executeTool(
            "word-document-server",
            "add_table",
            params
        );
    }
}
```

##### 示例2：创建 REST 控制器

创建控制器 `WordDocumentController.java`：

```java
package com.example.demo.controller;

import com.example.demo.service.WordDocumentService;
import org.springframework.web.bind.annotation.*;

@RestController
@RequestMapping("/api/word")
public class WordDocumentController {

    private final WordDocumentService wordDocumentService;

    public WordDocumentController(WordDocumentService wordDocumentService) {
        this.wordDocumentService = wordDocumentService;
    }

    /**
     * 创建文档
     * POST /api/word/create
     */
    @PostMapping("/create")
    public String createDocument(
            @RequestParam String filename,
            @RequestParam String title,
            @RequestParam(required = false) String author) {
        return wordDocumentService.createDocument(
            filename,
            title,
            author != null ? author : "System"
        );
    }

    /**
     * 添加段落
     * POST /api/word/paragraph
     */
    @PostMapping("/paragraph")
    public String addParagraph(
            @RequestParam String filename,
            @RequestParam String text,
            @RequestParam(defaultValue = "宋体") String fontName,
            @RequestParam(defaultValue = "12") int fontSize) {
        return wordDocumentService.addParagraph(filename, text, fontName, fontSize);
    }

    /**
     * 添加标题
     * POST /api/word/heading
     */
    @PostMapping("/heading")
    public String addHeading(
            @RequestParam String filename,
            @RequestParam String text,
            @RequestParam(defaultValue = "1") int level) {
        return wordDocumentService.addHeading(filename, text, level);
    }

    /**
     * 生成报告（综合示例）
     * POST /api/word/report
     */
    @PostMapping("/report")
    public String generateReport(@RequestParam String filename) {
        // 创建文档
        wordDocumentService.createDocument(filename, "月度报告", "Spring AI");

        // 添加主标题
        wordDocumentService.addHeading(filename, "2025年1月工作报告", 1);

        // 添加副标题
        wordDocumentService.addHeading(filename, "一、工作概述", 2);

        // 添加段落
        wordDocumentService.addParagraph(
            filename,
            "本月主要完成了系统架构升级和性能优化工作。",
            "宋体",
            12
        );

        // 添加表格
        String[][] tableData = {
            {"任务", "状态", "完成度"},
            {"架构升级", "已完成", "100%"},
            {"性能优化", "进行中", "80%"}
        };
        wordDocumentService.addTable(filename, 3, 3, tableData);

        return "报告生成成功：" + filename;
    }
}
```

##### 示例3：与 AI 模型结合使用

```java
package com.example.demo.service;

import com.alibaba.cloud.ai.mcp.McpToolClient;
import org.springframework.ai.chat.ChatClient;
import org.springframework.ai.chat.ChatResponse;
import org.springframework.ai.chat.prompt.Prompt;
import org.springframework.stereotype.Service;

@Service
public class AIDocumentService {

    private final ChatClient chatClient;
    private final McpToolClient mcpToolClient;

    public AIDocumentService(ChatClient chatClient, McpToolClient mcpToolClient) {
        this.chatClient = chatClient;
        this.mcpToolClient = mcpToolClient;
    }

    /**
     * AI 驱动的文档生成
     */
    public String generateAIDocument(String topic, String filename) {
        // 1. 使用 AI 生成文档内容
        String prompt = String.format(
            "请为主题'%s'生成一份结构化的文档大纲，包含标题、章节和内容。",
            topic
        );

        ChatResponse response = chatClient.call(new Prompt(prompt));
        String content = response.getResult().getOutput().getContent();

        // 2. 使用 MCP 工具创建文档
        mcpToolClient.executeTool(
            "word-document-server",
            "create_document",
            Map.of("filename", filename, "title", topic)
        );

        // 3. 添加 AI 生成的内容
        mcpToolClient.executeTool(
            "word-document-server",
            "add_paragraph",
            Map.of("filename", filename, "text", content)
        );

        return "AI 文档生成完成：" + filename;
    }
}
```

#### 验证和测试

##### 测试1：启动应用

```bash
# 启动 Spring Boot 应用
mvn spring-boot:run

# 或使用 Gradle
gradle bootRun
```

查看日志，确认 MCP Server 连接成功：

```
INFO  c.a.c.a.m.McpClientManager - Connecting to MCP Server: word-document-server
INFO  c.a.c.a.m.McpClientManager - MCP Server connected successfully
INFO  c.a.c.a.m.McpClientManager - Available tools: create_document, add_paragraph, add_heading, ...
```

##### 测试2：使用 curl 测试 API

```bash
# 创建文档
curl -X POST "http://localhost:8080/api/word/create" \
  -d "filename=test.docx" \
  -d "title=测试文档"

# 添加标题
curl -X POST "http://localhost:8080/api/word/heading" \
  -d "filename=test.docx" \
  -d "text=第一章 引言" \
  -d "level=1"

# 添加段落
curl -X POST "http://localhost:8080/api/word/paragraph" \
  -d "filename=test.docx" \
  -d "text=这是一段测试内容。"

# 生成完整报告
curl -X POST "http://localhost:8080/api/word/report" \
  -d "filename=monthly_report.docx"
```

##### 测试3：查看生成的文档

```bash
# 检查文件是否生成
ls -la *.docx

# Windows
dir *.docx
```

打开生成的 `.docx` 文件验证内容。

#### 高级配置

##### 配置多个 MCP Server

```yaml
spring:
  ai:
    alibaba:
      mcp:
        servers:
          # Word 文档 Server
          word-document-server:
            transport: stdio
            command: python
            args:
              - D:/path/to/word_mcp_server.py
            enabled: true

          # 其他 MCP Server（示例）
          excel-server:
            transport: http
            url: http://127.0.0.1:8001/mcp
            enabled: true

          pdf-server:
            transport: sse
            url: http://127.0.0.1:8002/sse
            enabled: true
```

##### 配置工具权限控制

```yaml
spring:
  ai:
    alibaba:
      mcp:
        servers:
          word-document-server:
            transport: stdio
            command: python
            args:
              - D:/path/to/word_mcp_server.py
            # 允许的工具列表（可选）
            allowed-tools:
              - create_document
              - add_paragraph
              - add_heading
              - add_table
            # 禁止的工具列表（可选）
            denied-tools:
              - delete_document
            enabled: true
```

##### 配置连接池和超时

```yaml
spring:
  ai:
    alibaba:
      mcp:
        # 全局配置
        connection:
          timeout: 30000        # 连接超时（毫秒）
          read-timeout: 60000   # 读取超时（毫秒）
          max-retries: 3        # 最大重试次数
        servers:
          word-document-server:
            transport: http
            url: http://127.0.0.1:8000/mcp
            # Server 级别超时（覆盖全局配置）
            timeout: 45000
            enabled: true
```

#### 故障排除

##### 问题1：MCP Server 连接失败

**症状**：

```
ERROR c.a.c.a.m.McpClientManager - Failed to connect to MCP Server: word-document-server
```

**解决方案**：

1. 检查 Python 路径是否正确
2. 检查 MCP Server 脚本路径
3. 确认依赖已安装（`pip list | grep fastmcp`）
4. 手动测试启动：`python word_mcp_server.py`

##### 问题2：工具调用失败

**症状**：

```
Tool 'create_document' not found in server 'word-document-server'
```

**解决方案**：

1. 检查工具名称拼写
2. 确认 MCP Server 已正确启动
3. 查看 MCP Server 日志
4. 重启 Spring Boot 应用

##### 问题3：路径包含空格或中文

**症状**：

```
FileNotFoundError: No such file or directory
```

**解决方案**：

使用完整路径并正确转义：

```yaml
spring:
  ai:
    alibaba:
      mcp:
        servers:
          word-document-server:
            command: "D:/Program Files/Python311/python.exe"  # 使用引号
            args:
              - "D:/我的项目/Office-Word-MCP-Server/word_mcp_server.py"
```

或使用短路径（Windows）：

```yaml
command: "D:/PROGRA~1/Python311/python.exe"
```

#### 完整项目示例

完整的项目结构：

```
spring-ai-word-demo/
├── src/
│   └── main/
│       ├── java/
│       │   └── com/example/demo/
│       │       ├── DemoApplication.java
│       │       ├── controller/
│       │       │   └── WordDocumentController.java
│       │       └── service/
│       │           ├── WordDocumentService.java
│       │           └── AIDocumentService.java
│       └── resources/
│           └── application.yml
├── pom.xml
└── README.md
```

**application.yml** 完整配置：

```yaml
server:
  port: 8080

spring:
  application:
    name: spring-ai-word-demo

  ai:
    alibaba:
      # AI 模型配置（如使用通义千问）
      api-key: ${DASHSCOPE_API_KEY}

      # MCP 配置
      mcp:
        servers:
          word-document-server:
            transport: stdio
            command: python
            args:
              - D:/02_Dev/Workspace/GitHub/Office-Word-MCP-Server/word_mcp_server.py
            env:
              PYTHONPATH: D:/02_Dev/Workspace/GitHub/Office-Word-MCP-Server
              MCP_TRANSPORT: stdio
            enabled: true

# 日志配置
logging:
  level:
    com.alibaba.cloud.ai: DEBUG
    com.example.demo: DEBUG
```

#### 参考资源

- **Spring AI Alibaba 官方文档**: [https://sca.aliyun.com/ai/](https://sca.aliyun.com/ai/)
- **MCP 协议规范**: [https://modelcontextprotocol.io/](https://modelcontextprotocol.io/)
- **Office-Word-MCP-Server GitHub**: [https://github.com/GongRzhe/Office-Word-MCP-Server](https://github.com/GongRzhe/Office-Word-MCP-Server)
- **Spring AI 官方文档**: [https://docs.spring.io/spring-ai/reference/](https://docs.spring.io/spring-ai/reference/)

---

## 第4章：使用示例

### 4.1 通过 Claude Code 使用

#### 示例1：创建简单文档

**用户输入**：

```
创建一个 report.docx 文档，包含：
- 标题：月度工作报告
- 副标题：2025年1月
- 正文：本月主要完成了以下工作...
```

**Claude Code 响应**：

```
我将为您创建 Word 文档。

工具调用：
1. create_document(filename="report.docx")
2. add_heading(filename="report.docx", text="月度工作报告", level=1)
3. add_heading(filename="report.docx", text="2025年1月", level=2)
4. add_paragraph(filename="report.docx", text="本月主要完成了以下工作...")

✅ 文档已创建：report.docx
```

#### 示例2：添加表格

**用户输入**：

```
在 report.docx 中添加一个 3x4 的表格，包含销售数据：
第一行：产品名称 | 销售量 | 销售额 | 增长率
第二行：产品A | 1200 | ¥120,000 | 15%
第三行：产品B | 800 | ¥80,000 | 8%
```

**Claude Code 响应**：

```
工具调用：
add_table(
  filename="report.docx",
  rows=3,
  cols=4,
  data=[
    ["产品名称", "销售量", "销售额", "增长率"],
    ["产品A", "1200", "¥120,000", "15%"],
    ["产品B", "800", "¥80,000", "8%"]
  ]
)

✅ 表格已添加
```

#### 示例3：格式化文本

**用户输入**：

```
将 report.docx 中第二段的 "重要提示" 四个字设为红色粗体
```

**Claude Code 响应**：

```
工具调用：
1. find_text_in_document(filename="report.docx", text_to_find="重要提示")
   → 找到位置：段落索引 1, 起始位置 10

2. format_text(
     filename="report.docx",
     paragraph_index=1,
     start_pos=10,
     end_pos=14,
     bold=True,
     color="FF0000"
   )

✅ 文本格式已应用
```

---

### 4.2 通过 Cline 使用

#### 示例1：创建技术文档

**在 Cline 聊天中输入**：

```
帮我创建一个技术文档 tech_spec.docx，包含：

1. 标题："系统架构设计"
2. 章节1："概述"，内容："本文档描述系统的整体架构..."
3. 章节2："技术栈"，包含一个表格列出前端、后端、数据库技术
4. 章节3："部署方案"
```

**Cline 执行过程**：

```
🔧 使用工具：create_document
   参数：filename="tech_spec.docx"

🔧 使用工具：add_heading
   参数：text="系统架构设计", level=1

🔧 使用工具：add_heading
   参数：text="概述", level=2

🔧 使用工具：add_paragraph
   参数：text="本文档描述系统的整体架构..."

🔧 使用工具：add_heading
   参数：text="技术栈", level=2

🔧 使用工具：add_table
   参数：rows=4, cols=2, data=[...]

✅ 文档创建完成：tech_spec.docx
```

#### 示例2：批量处理文档

**在 Cline 中输入**：

```
为项目中的所有 .docx 文档添加页脚，内容为"公司机密 - 2025"
```

**Cline 执行**：

```
🔍 扫描目录，找到 5 个 .docx 文件

正在处理：
1. report1.docx ✅
2. report2.docx ✅
3. proposal.docx ✅
4. summary.docx ✅
5. notes.docx ✅

✅ 所有文档已更新
```

---

### 4.3 常用操作示例

#### 创建文档

```python
# 工具：create_document
create_document(
    filename="example.docx",
    title="示例文档",
    author="张三"
)
```

#### 添加内容

```python
# 添加标题
add_heading(
    filename="example.docx",
    text="第一章 引言",
    level=1,
    bold=True
)

# 添加段落
add_paragraph(
    filename="example.docx",
    text="这是一段示例文字。",
    font_name="宋体",
    font_size=12
)

# 添加表格
add_table(
    filename="example.docx",
    rows=3,
    cols=3,
    data=[
        ["列1", "列2", "列3"],
        ["数据1", "数据2", "数据3"],
        ["数据4", "数据5", "数据6"]
    ]
)
```

#### 格式化

```python
# 格式化文本
format_text(
    filename="example.docx",
    paragraph_index=2,
    start_pos=0,
    end_pos=5,
    bold=True,
    color="0000FF",
    font_size=14
)

# 搜索替换
search_and_replace(
    filename="example.docx",
    find_text="旧文本",
    replace_text="新文本"
)
```

#### 高级操作

```python
# 插入图片
add_picture(
    filename="example.docx",
    image_path="C:/images/chart.png",
    width=6.0  # 英寸
)

# 合并文档
merge_documents(
    output_filename="merged.docx",
    input_filenames=["doc1.docx", "doc2.docx", "doc3.docx"]
)

# 转换为 PDF
convert_to_pdf(
    filename="example.docx",
    output_filename="example.pdf"
)
```

---

## 第5章：传输方式详解

### 5.1 STDIO 传输

#### 概述

**STDIO (Standard Input/Output)** 是最简单的传输方式，通过标准输入输出流通信。

#### 适用场景

- ✅ 本地 AI 工具（Claude Code、Claude Desktop）
- ✅ 本地开发和测试
- ✅ 无网络需求的应用

#### 配置示例

**环境变量**（.env）：

```env
MCP_TRANSPORT=stdio
```

**MCP Client 配置**（mcp.json）：

```json
{
  "mcpServers": {
    "word-document-server": {
      "command": "python",
      "args": ["path/to/word_mcp_server.py"],
      "env": {
        "MCP_TRANSPORT": "stdio"
      }
    }
  }
}
```

#### 启动方式

```bash
# 直接运行
python word_mcp_server.py

# 或通过 MCP Client 自动启动
```

#### 优缺点

| 优点 | 缺点 |
|------|------|
| ✅ 配置简单 | ❌ 仅限本地 |
| ✅ 性能最佳 | ❌ 不支持网络访问 |
| ✅ 延迟最低 | ❌ 不支持多客户端 |
| ✅ 无需端口 | |

---

### 5.2 Streamable HTTP 传输

#### 概述

**Streamable HTTP** 使用现代 HTTP 流式传输协议，支持网络访问。

#### 适用场景

- ✅ 远程访问 MCP Server
- ✅ Web 应用集成
- ✅ 云端部署（Render、Heroku 等）
- ✅ 多客户端共享

#### 配置示例

**环境变量**（.env）：

```env
MCP_TRANSPORT=streamable-http
MCP_HOST=0.0.0.0
MCP_PORT=8000
MCP_PATH=/mcp
```

**启动服务器**：

```bash
python word_mcp_server.py
```

输出：

```
Starting Word Document Server...
Transport: streamable-http
Server listening at: http://0.0.0.0:8000/mcp
```

#### MCP Client 配置

**STDIO 方式（推荐）**：

```json
{
  "mcpServers": {
    "word-document-server": {
      "command": "python",
      "args": ["path/to/word_mcp_server.py"],
      "env": {
        "MCP_TRANSPORT": "streamable-http",
        "MCP_HOST": "127.0.0.1",
        "MCP_PORT": "8000",
        "MCP_PATH": "/mcp"
      }
    }
  }
}
```

**HTTP 直连方式**：

```json
{
  "mcpServers": {
    "word-document-server": {
      "url": "http://127.0.0.1:8000/mcp"
    }
  }
}
```

#### 测试访问

```bash
# 使用 curl 测试
curl -X POST http://127.0.0.1:8000/mcp \
  -H "Content-Type: application/json" \
  -d '{"method": "tools/list"}'
```

#### 优缺点

| 优点 | 缺点 |
|------|------|
| ✅ 支持网络访问 | ❌ 需要配置端口 |
| ✅ 现代协议 | ❌ 配置稍复杂 |
| ✅ 多客户端支持 | ❌ 需要防火墙配置 |
| ✅ 易于集成 | |

---

### 5.3 SSE 传输

#### 概述

**SSE (Server-Sent Events)** 是基于 HTTP 的单向通信协议，适用于特定兼容性需求。

#### 适用场景

- ✅ 需要 SSE 兼容性的应用
- ✅ 旧系统集成
- ✅ 云平台部署（如 Render）

#### 配置示例

**环境变量**（.env）：

```env
MCP_TRANSPORT=sse
MCP_HOST=0.0.0.0
MCP_PORT=8000
MCP_SSE_PATH=/sse
```

**启动服务器**：

```bash
python word_mcp_server.py
```

输出：

```
Starting Word Document Server...
Transport: sse
Server listening at: http://0.0.0.0:8000/sse
```

#### MCP Client 配置

**Cline SSE 配置**：

```json
{
  "mcpServers": {
    "word-document-server": {
      "url": "http://127.0.0.1:8000/sse"
    }
  }
}
```

#### 健康检查端点

SSE 模式自动提供健康检查：

```bash
curl http://127.0.0.1:8000/health
```

响应：

```json
{
  "status": "ok",
  "transport": "sse"
}
```

#### 优缺点

| 优点 | 缺点 |
|------|------|
| ✅ 兼容性好 | ❌ 技术较旧 |
| ✅ 支持网络 | ❌ 性能不如 HTTP |
| ✅ 单向通信简单 | ❌ 双向通信受限 |

---

### 传输方式选择决策树

```
需要使用 Office-Word-MCP-Server
        │
        ├─ 本地使用（Claude Code/Desktop）？
        │   └─ 是 → 选择 STDIO ✅
        │
        ├─ 需要远程访问？
        │   ├─ 现代应用 → 选择 Streamable HTTP ✅
        │   └─ 兼容性优先 → 选择 SSE
        │
        └─ 云平台部署（Render 等）？
            └─ 是 → 选择 SSE ✅
```

---

## 第6章：故障排除

### 6.1 安装问题

#### 问题1：Python 版本不兼容

**症状**：

```
Error: Python 3.11 or higher is required
```

**解决方案**：

```bash
# 检查当前版本
python --version

# Windows: 下载安装最新 Python
# https://www.python.org/downloads/

# Linux Ubuntu:
sudo apt install python3.11

# 验证
python3.11 --version
```

#### 问题2：依赖安装失败

**症状**：

```
ERROR: Failed building wheel for python-docx
```

**解决方案**：

```bash
# 方法1：升级 pip
pip install --upgrade pip setuptools wheel

# 方法2：使用国内镜像（中国用户）
pip install -r requirements.txt -i https://pypi.tuna.tsinghua.edu.cn/simple

# 方法3：手动安装核心依赖
pip install python-docx
pip install fastmcp
```

#### 问题3：权限不足

**症状**（Windows）：

```
PermissionError: [WinError 5] Access is denied
```

**解决方案**：

```cmd
# 以管理员身份运行 CMD
# 右键 "命令提示符" → "以管理员身份运行"

# 或使用用户目录安装
pip install --user -r requirements.txt
```

**症状**（Linux）：

```
Permission denied: '/usr/local/lib/python3.11'
```

**解决方案**：

```bash
# 使用虚拟环境（推荐）
python3.11 -m venv .venv
source .venv/bin/activate
pip install -r requirements.txt

# 或使用 --user 标志
pip install --user -r requirements.txt
```

#### 问题4：虚拟环境创建失败

**症状**：

```
Error: No module named 'venv'
```

**解决方案**：

```bash
# Ubuntu/Debian
sudo apt install python3.11-venv

# CentOS/RHEL
sudo yum install python311-devel
```

---

### 6.2 配置问题

#### 问题1：MCP Server 未识别

**症状**（Claude Code/Cline）：

```
No MCP servers found
```

**解决方案**：

**检查清单**：

1. ✅ 配置文件路径正确
   ```bash
   # 检查文件是否存在
   # Windows
   dir .claude\mcp.json

   # Linux
   ls -la .claude/mcp.json
   ```

2. ✅ JSON 格式正确
   ```bash
   # 验证 JSON 格式
   python -m json.tool .claude/mcp.json
   ```

3. ✅ Python 路径正确
   ```bash
   # 测试 Python 命令
   python --version

   # 测试脚本路径
   python path/to/word_mcp_server.py
   ```

4. ✅ 重启客户端
   - Claude Code: 退出后重新启动
   - Cline: 重启 VSCode

#### 问题2：路径包含空格或特殊字符

**症状**：

```
FileNotFoundError: [Errno 2] No such file or directory
```

**解决方案**：

**错误示例**：

```json
{
  "command": "D:\Program Files\Python\python.exe"
}
```

**正确示例**：

```json
{
  "command": "D:\\Program Files\\Python\\python.exe"
}
```

或使用短路径：

```cmd
# 获取短路径（Windows）
dir /x "C:\Program Files"
# 输出: PROGRA~1

# 使用短路径
"command": "D:\\PROGRA~1\\Python\\python.exe"
```

#### 问题3：环境变量未生效

**症状**：

```
MCP_TRANSPORT environment variable not set
```

**解决方案**：

**检查 .env 文件**：

```bash
# 确认 .env 文件存在
ls -la .env

# 查看内容
cat .env  # Linux
type .env  # Windows
```

**确认配置正确**：

```env
# .env 文件内容
MCP_TRANSPORT=stdio
```

**在 mcp.json 中显式设置**：

```json
{
  "mcpServers": {
    "word-document-server": {
      "command": "python",
      "args": ["word_mcp_server.py"],
      "env": {
        "MCP_TRANSPORT": "stdio"  ← 确保设置
      }
    }
  }
}
```

#### 问题4：Cline 配置文件未找到

**症状**：

```
Cannot find cline_mcp_settings.json
```

**解决方案**：

**手动创建配置文件**：

```bash
# Windows
mkdir %APPDATA%\Code\User\globalStorage\saoudrizwan.cline\settings
notepad %APPDATA%\Code\User\globalStorage\saoudrizwan.cline\settings\cline_mcp_settings.json

# Linux
mkdir -p ~/.config/Code/User/globalStorage/saoudrizwan.cline/settings
nano ~/.config/Code/User/globalStorage/saoudrizwan.cline/settings/cline_mcp_settings.json
```

添加基础配置：

```json
{
  "mcpServers": {}
}
```

---

### 6.3 连接问题

#### 问题1：Server 启动失败

**症状**：

```
Error: Server failed to start
```

**解决方案**：

**步骤1：检查 Python 依赖**

```bash
pip list | grep -E "fastmcp|python-docx"

# 输出应包含：
# fastmcp         2.8.1
# python-docx     1.1.2
```

**步骤2：手动测试启动**

```bash
python word_mcp_server.py
```

查看错误输出。

**步骤3：检查端口占用（HTTP/SSE 模式）**

```bash
# Windows
netstat -ano | findstr :8000

# Linux
netstat -tuln | grep :8000
```

如果端口被占用，修改端口：

```env
MCP_PORT=8001
```

#### 问题2：STDIO 通信失败

**症状**：

```
Timeout waiting for server response
```

**解决方案**：

**启用调试日志**：

```json
{
  "mcpServers": {
    "word-document-server": {
      "command": "python",
      "args": ["word_mcp_server.py"],
      "env": {
        "MCP_TRANSPORT": "stdio",
        "MCP_DEBUG": "1"  ← 添加调试
      }
    }
  }
}
```

**检查脚本可执行性**：

```bash
# Linux: 确保脚本有执行权限
chmod +x word_mcp_server.py
```

#### 问题3：HTTP 连接超时

**症状**：

```
Connection timeout: http://127.0.0.1:8000/mcp
```

**解决方案**：

**检查服务器运行状态**：

```bash
# 测试端点
curl http://127.0.0.1:8000/mcp

# 或使用浏览器访问
# http://127.0.0.1:8000/mcp
```

**检查防火墙**（Windows）：

```powershell
# 允许端口 8000
netsh advfirewall firewall add rule name="MCP Server" dir=in action=allow protocol=TCP localport=8000
```

**检查防火墙**（Linux）：

```bash
# Ubuntu (ufw)
sudo ufw allow 8000/tcp

# CentOS (firewalld)
sudo firewall-cmd --add-port=8000/tcp --permanent
sudo firewall-cmd --reload
```

#### 问题4：工具调用失败

**症状**：

```
Error: Tool 'create_document' not found
```

**解决方案**：

**验证 Server 连接**：

在客户端测试：

```
列出所有可用工具
```

如果没有工具列表，重启 Server：

```bash
# 重启 MCP Server
# Cline: 点击 "Restart" 按钮
# Claude Code: 重新启动应用
```

**检查工具权限（Cline）**：

在 `cline_mcp_settings.json` 中：

```json
{
  "mcpServers": {
    "word-document-server": {
      "alwaysAllow": ["create_document", "add_paragraph"],
      "disabled": false  ← 确保未禁用
    }
  }
}
```

#### 问题5：文档操作权限错误

**症状**：

```
PermissionError: Cannot write to 'document.docx'
```

**解决方案**：

**检查文件权限**：

```bash
# Windows: 右键文件 → 属性 → 安全
# 确保当前用户有"完全控制"权限

# Linux
ls -la document.docx
chmod 644 document.docx  # 设置可读写
```

**检查文件是否被占用**：

- 关闭 Word 应用程序
- 检查进程：

```bash
# Windows
tasklist | findstr WINWORD.EXE

# Linux
lsof document.docx
```

**使用绝对路径**：

```json
{
  "env": {
    "WORKING_DIR": "D:\\Documents\\output"
  }
}
```

---

### 常见错误速查表

| 错误信息 | 可能原因 | 解决方案 |
|----------|----------|----------|
| `Python 3.11+ required` | Python 版本过低 | 升级 Python 至 3.11+ |
| `Module 'fastmcp' not found` | 依赖未安装 | `pip install fastmcp` |
| `Permission denied` | 权限不足 | 使用虚拟环境或 `--user` 标志 |
| `JSON decode error` | 配置文件格式错误 | 验证 JSON 格式 |
| `Port 8000 already in use` | 端口被占用 | 修改端口或关闭占用进程 |
| `Server timeout` | Server 未启动 | 检查 Server 运行状态 |
| `Tool not found` | Server 未连接 | 重启 Server 和 Client |
| `File in use` | 文件被占用 | 关闭 Word 或其他程序 |

---

## 第7章：企业级部署架构

本章介绍如何在企业环境中部署 Office-Word-MCP-Server，实现多用户共享访问。

### 7.1 B/S 部署架构

#### 7.1.1 架构概述

**B/S（Browser/Server）架构** 通过 Web 浏览器访问服务，所有服务组件集中部署在服务器端。

##### 核心特点

| 特点 | 说明 |
|------|------|
| **服务端集中部署** | Spring AI（MCP Client）和 MCP Server 都部署在服务器 |
| **文档在服务器生成** | Word 文档在 MCP Server 所在服务器生成和存储 |
| **浏览器下载获取** | 用户通过浏览器下载生成的文档 |
| **零客户端安装** | 用户只需浏览器，无需安装任何软件 |

##### 设计目标

- 让内网所有用户通过浏览器使用 Word 文档生成服务
- 无需在用户电脑安装任何软件
- 统一管理、权限控制、审计日志
- 集中维护，降低运维成本

##### 适用场景

| 场景 | 说明 | 示例 |
|------|------|------|
| **企业内部文档自动化** | 批量生成结构化文档 | 报表、合同、通知 |
| **SaaS 服务集成** | 为 Web 应用添加文档能力 | 在线办公系统 |
| **多租户服务** | 统一入口，按需调用 | 集团多部门共享 |
| **轻量级使用** | 偶尔使用，不需要安装客户端 | 临时文档生成需求 |

##### 核心组件

| 层级 | 组件 | 部署位置 | 职责 |
|------|------|----------|------|
| **表现层** | 用户浏览器 | 用户电脑 | 界面交互、结果展示、文档下载 |
| **应用层** | Spring AI（MCP Client） | 应用服务器 | 业务逻辑、认证授权、MCP 调用 |
| **服务层** | MCP Server | MCP 服务器 | Word 文档操作、文件生成存储 |

---

#### 7.1.2 架构设计

##### 整体架构图

```
                   192.168.1.0/24 内网用户
                ┌─────┬─────┬─────┬─────┐
                │ PC1 │ PC2 │ PC3 │ ... │
                └──┬──┴──┬──┴──┬──┴─────┘
                   │     │     │
                   └─────┼─────┘
                         │ HTTP (80/443)
                         ▼
             ┌───────────────────────┐
             │   192.168.1.100       │
             │   ┌───────────────┐   │
             │   │  Spring AI    │   │
             │   │  Application  │   │
             │   │  (MCP Client) │   │
             │   └───────┬───────┘   │
             └───────────┼───────────┘
                         │ MCP/HTTP (8000)
                         ▼
             ┌───────────────────────┐
             │   192.168.1.101       │
             │   ┌───────────────┐   │
             │   │  MCP Server   │   │
             │   │  (Word 操作)  │   │
             │   └───────┬───────┘   │
             │           ▼           │
             │   [文档存储目录]       │
             └───────────────────────┘
```

##### 数据流转

```
用户请求                Spring AI              MCP Server
   │                       │                       │
   │  "生成月度报告"        │                       │
   ├──────────────────────▶│                       │
   │                       │  调用 create_document │
   │                       ├──────────────────────▶│
   │                       │                       │ 生成文档
   │                       │      返回结果         │
   │                       │◀──────────────────────┤
   │                       │                       │
   │   返回下载链接/文件    │                       │
   │◀──────────────────────┤                       │
```

##### 关键说明

| 要点 | 说明 |
|------|------|
| **文档生成位置** | 文档在 MCP Server（192.168.1.101）生成，不在用户电脑 |
| **传输协议** | 必须使用 HTTP 或 SSE（跨网络通信） |
| **中间层作用** | Spring AI 作为 MCP Client，统一管理所有用户请求 |

---

#### 7.1.3 部署配置

##### 服务层配置（192.168.1.101 - MCP Server）

**步骤1：配置环境变量**

创建或编辑 `.env` 文件：

```env
# 传输方式配置
MCP_TRANSPORT=streamable-http
MCP_HOST=0.0.0.0
MCP_PORT=8000
MCP_PATH=/mcp

# 文档存储路径（可选）
DOCUMENT_ROOT=/data/documents
```

**步骤2：启动服务**

```bash
# 激活虚拟环境（如果使用）
source .venv/bin/activate  # Linux
# 或
.venv\Scripts\activate  # Windows

# 启动服务
python word_mcp_server.py
```

**步骤3：配置防火墙**

```bash
# Linux (Ubuntu)
sudo ufw allow 8000/tcp

# Linux (CentOS)
sudo firewall-cmd --add-port=8000/tcp --permanent
sudo firewall-cmd --reload

# Windows PowerShell（管理员）
netsh advfirewall firewall add rule name="MCP Server" dir=in action=allow protocol=TCP localport=8000
```

**步骤4：验证服务**

```bash
# 从应用服务器测试连接
curl http://192.168.1.101:8000/mcp
```

##### 应用层配置（192.168.1.100 - Spring AI）

**步骤1：添加依赖**

`pom.xml`：

```xml
<dependencies>
    <dependency>
        <groupId>com.alibaba.cloud.ai</groupId>
        <artifactId>spring-ai-alibaba-starter</artifactId>
        <version>1.0.0-M2</version>
    </dependency>
    <dependency>
        <groupId>org.springframework.boot</groupId>
        <artifactId>spring-boot-starter-web</artifactId>
    </dependency>
</dependencies>
```

**步骤2：配置 MCP 连接**

`application.yml`：

```yaml
server:
  port: 8080

spring:
  ai:
    alibaba:
      mcp:
        servers:
          word-document-server:
            transport: http
            url: http://192.168.1.101:8000/mcp
            timeout: 30000
            enabled: true
```

**步骤3：创建 REST 接口**

```java
@RestController
@RequestMapping("/api/document")
public class DocumentController {

    private final McpToolClient mcpToolClient;

    public DocumentController(McpToolClient mcpToolClient) {
        this.mcpToolClient = mcpToolClient;
    }

    @PostMapping("/create")
    public ResponseEntity<String> createDocument(
            @RequestParam String filename,
            @RequestParam String title) {

        String result = mcpToolClient.executeTool(
            "word-document-server",
            "create_document",
            Map.of("filename", filename, "title", title)
        );

        return ResponseEntity.ok(result);
    }

    @PostMapping("/add-content")
    public ResponseEntity<String> addContent(
            @RequestParam String filename,
            @RequestParam String heading,
            @RequestParam String content) {

        // 添加标题
        mcpToolClient.executeTool(
            "word-document-server",
            "add_heading",
            Map.of("filename", filename, "text", heading, "level", 1)
        );

        // 添加段落
        String result = mcpToolClient.executeTool(
            "word-document-server",
            "add_paragraph",
            Map.of("filename", filename, "text", content)
        );

        return ResponseEntity.ok(result);
    }
}
```

**步骤4：启动应用**

```bash
mvn spring-boot:run
```

##### 用户访问

内网用户通过浏览器访问：

```
http://192.168.1.100:8080/api/document/create?filename=report.docx&title=月度报告
```

或通过前端页面调用 API。

---

#### 7.1.4 文档存储方案

由于文档在 MCP Server 生成，需要解决用户获取文档的问题。

##### 方案对比

| 方案 | 实现方式 | 优点 | 缺点 | 适用场景 |
|------|----------|------|------|----------|
| **共享存储** | SMB/NFS 挂载 | 简单直接、无需开发 | 需要网络共享配置 | 中小规模 |
| **文件 API** | MCP Server 提供下载接口 | 灵活可控 | 需要额外开发 | 需要细粒度控制 |
| **对象存储** | MinIO/阿里云 OSS | 高可用、易扩展 | 架构复杂度增加 | 大规模生产环境 |

##### 方案A：共享存储（推荐中小规模）

**MCP Server 端配置**：

```bash
# Linux: 创建共享目录
sudo mkdir -p /data/documents
sudo chmod 777 /data/documents

# 安装 Samba
sudo apt install samba -y

# 配置共享
sudo nano /etc/samba/smb.conf
```

添加配置：

```ini
[documents]
   path = /data/documents
   browseable = yes
   read only = no
   guest ok = no
   valid users = @documents
```

**Spring AI 端访问**：

```yaml
document:
  storage:
    type: shared
    # Windows 格式
    path: \\192.168.1.101\documents
    # 或 Linux 挂载点
    # path: /mnt/mcp-documents
```

##### 方案B：文件传输 API

在 MCP Server 端添加文件下载接口（需要自行扩展）：

```python
# 在 MCP Server 中添加文件下载工具
@mcp.tool()
def get_document_content(filename: str) -> bytes:
    """获取文档内容用于传输"""
    file_path = os.path.join(DOCUMENT_ROOT, filename)
    with open(file_path, 'rb') as f:
        return base64.b64encode(f.read()).decode()
```

Spring AI 端接收并返回给用户：

```java
@GetMapping("/download/{filename}")
public ResponseEntity<byte[]> downloadDocument(@PathVariable String filename) {
    // 调用 MCP 获取文件内容
    String base64Content = mcpToolClient.executeTool(
        "word-document-server",
        "get_document_content",
        Map.of("filename", filename)
    );

    byte[] content = Base64.getDecoder().decode(base64Content);

    return ResponseEntity.ok()
        .header("Content-Disposition", "attachment; filename=" + filename)
        .header("Content-Type", "application/vnd.openxmlformats-officedocument.wordprocessingml.document")
        .body(content);
}
```

##### 方案C：对象存储

```yaml
# Spring AI 配置
document:
  storage:
    type: oss
    endpoint: http://192.168.1.102:9000  # MinIO
    bucket: documents
    access-key: minioadmin
    secret-key: minioadmin
```

MCP Server 生成文档后上传到对象存储，Spring AI 返回预签名 URL 给用户下载。

---

#### 7.1.5 生产环境建议

##### 高可用架构

```
                    Load Balancer (Nginx/HAProxy)
                              │
               ┌──────────────┼──────────────┐
               ▼              ▼              ▼
          Spring AI      Spring AI      Spring AI
           Node 1         Node 2         Node 3
               │              │              │
               └──────────────┼──────────────┘
                              │
                    Load Balancer
                              │
               ┌──────────────┼──────────────┐
               ▼              ▼              ▼
          MCP Server     MCP Server     MCP Server
           Node 1         Node 2         Node 3
               │              │              │
               └──────────────┼──────────────┘
                              │
                      共享存储 / 对象存储
```

##### 安全加固清单

| 类别 | 措施 | 优先级 |
|------|------|--------|
| **传输安全** | HTTPS/TLS 加密 | 高 |
| **认证授权** | JWT/OAuth2 + RBAC | 高 |
| **网络隔离** | 内网部署、防火墙规则 | 高 |
| **审计日志** | 记录所有文档操作 | 中 |
| **文件安全** | 访问权限控制、病毒扫描 | 中 |
| **限流保护** | API 限流、防止滥用 | 中 |

##### 监控告警

```yaml
# Prometheus 监控指标示例
management:
  endpoints:
    web:
      exposure:
        include: health,metrics,prometheus
  metrics:
    export:
      prometheus:
        enabled: true
```

关键监控指标：

- MCP Server 连接状态
- 文档生成成功率
- API 响应时间
- 存储空间使用率

##### Nginx 反向代理配置

```nginx
upstream spring-ai {
    server 192.168.1.100:8080;
    # 多节点时添加更多 server
}

server {
    listen 80;
    server_name doc.company.com;

    location / {
        proxy_pass http://spring-ai;
        proxy_set_header Host $host;
        proxy_set_header X-Real-IP $remote_addr;
    }
}
```

---

#### 7.1.6 完整部署示例

##### 部署清单

| 服务器 | IP | 角色 | 端口 |
|--------|-----|------|------|
| app-server | 192.168.1.100 | Spring AI 应用 | 8080 |
| mcp-server | 192.168.1.101 | MCP Server | 8000 |
| storage | 192.168.1.102 | MinIO 对象存储（可选） | 9000 |

##### Docker Compose 一键部署

```yaml
# docker-compose.yml
version: '3.8'

services:
  mcp-server:
    build: ./Office-Word-MCP-Server
    ports:
      - "8000:8000"
    environment:
      - MCP_TRANSPORT=streamable-http
      - MCP_HOST=0.0.0.0
      - MCP_PORT=8000
    volumes:
      - ./documents:/data/documents

  spring-ai-app:
    build: ./spring-ai-word-demo
    ports:
      - "8080:8080"
    environment:
      - MCP_SERVER_URL=http://mcp-server:8000/mcp
    depends_on:
      - mcp-server

  # 可选：对象存储
  minio:
    image: minio/minio
    ports:
      - "9000:9000"
    environment:
      - MINIO_ROOT_USER=minioadmin
      - MINIO_ROOT_PASSWORD=minioadmin
    command: server /data
    volumes:
      - ./minio-data:/data
```

启动命令：

```bash
docker-compose up -d
```

##### 验证部署

```bash
# 1. 检查 MCP Server
curl http://192.168.1.101:8000/mcp

# 2. 检查 Spring AI 应用
curl http://192.168.1.100:8080/actuator/health

# 3. 测试创建文档
curl -X POST "http://192.168.1.100:8080/api/document/create" \
  -d "filename=test.docx" \
  -d "title=测试文档"

# 4. 检查文档是否生成
ls -la /data/documents/  # 在 MCP Server 上执行
```

---

### 7.2 C/S 部署架构

#### 7.2.1 架构概述

**C/S（Client/Server）架构** 通过桌面客户端访问服务，提供原生应用体验。根据 MCP Server 的部署位置，分为两种模式。

##### 两种模式对比

| 模式 | MCP Server 位置 | 文档生成位置 | 协议 | 特点 |
|------|-----------------|--------------|------|------|
| **远程服务器模式** | 远程服务器 | 服务器 | HTTP/SSE | 轻量客户端，集中管理 |
| **本地完整模式** | 用户电脑 | **用户本地** | STDIO | 离线可用，无需下载 |

##### 核心特点

| 特点 | 说明 |
|------|------|
| **桌面客户端** | 使用 Electron 构建原生桌面应用 |
| **MCP Client 分布式** | 每个用户电脑运行独立的 MCP Client |
| **直连 MCP Server** | 无需中间应用层，客户端直接连接 |
| **原生体验** | 系统集成、文件管理、离线能力 |

##### 技术选型

| 组件 | 技术 | 说明 |
|------|------|------|
| 桌面客户端 | Electron | 跨平台桌面应用框架 |
| MCP Client SDK | @modelcontextprotocol/sdk | 官方 TypeScript SDK |
| UI 框架 | React/Vue | 根据团队技术栈选择 |

---

#### 7.2.2 远程服务器模式

##### 架构设计

MCP Server 部署在远程服务器，客户端通过网络连接。

```
         192.168.1.0/24 内网用户
      ┌──────┬──────┬──────┬──────┐
      │ PC1  │ PC2  │ PC3  │ ...  │
      │      │      │      │      │
      │┌────┐│┌────┐│┌────┐│      │
      ││Elec││││Elec││││Elec│││      │
      ││tron││││tron││││tron│││      │
      ││    ││││    ││││    │││      │
      ││MCP ││││MCP ││││MCP │││      │
      ││Clie││││Clie││││Clie│││      │
      ││nt  ││││nt  ││││nt  │││      │
      │└──┬─┘│└──┬─┘│└──┬─┘│      │
      └───┼──┴───┼──┴───┼──┴──────┘
          │      │      │
          └──────┼──────┘
                 │ HTTP/SSE (8000)
                 ▼
      ┌─────────────────────┐
      │   192.168.1.101     │
      │   ┌─────────────┐   │
      │   │  MCP Server │   │
      │   └──────┬──────┘   │
      │          ▼          │
      │   [文档存储目录]     │
      └─────────────────────┘
```

##### 核心特点

| 特点 | 说明 |
|------|------|
| **文档位置** | 服务器端（需要下载到本地） |
| **传输协议** | HTTP 或 SSE（跨网络） |
| **与 B/S 区别** | 用 Electron 替代浏览器，无中间应用层 |
| **客户端职责** | UI 界面 + MCP Client |

##### 适用场景

- 需要原生桌面体验，但希望集中管理文档
- 客户端保持轻量，不安装 Python 环境
- 多用户共享同一 MCP Server

##### 部署配置

**服务端配置**：与 B/S 架构相同（参考 7.1.3）

**客户端配置要点**：

```javascript
// Electron 主进程中配置 MCP Client
// 使用 @modelcontextprotocol/sdk

const serverUrl = "http://192.168.1.101:8000/mcp";

// 通过 HTTP 传输连接远程 MCP Server
// 具体实现参考官方 SDK 文档
```

---

#### 7.2.3 本地完整模式

##### 架构设计

MCP Server 随客户端一起安装在用户电脑，实现完全本地化运行。

```
用户电脑
┌─────────────────────────────────┐
│  ┌─────────────────────────┐    │
│  │      Electron App       │    │
│  │  ┌─────────────────┐    │    │
│  │  │   MCP Client    │    │    │
│  │  │ (@modelcontext  │    │    │
│  │  │  protocol/sdk)  │    │    │
│  │  └────────┬────────┘    │    │
│  └───────────┼─────────────┘    │
│              │ STDIO            │
│              ▼                  │
│  ┌─────────────────────────┐    │
│  │      MCP Server         │    │
│  │  (随应用安装/打包)       │    │
│  └───────────┬─────────────┘    │
│              ▼                  │
│  ┌─────────────────────────┐    │
│  │   本地文档目录           │    │
│  │   C:\Users\xxx\Documents │    │
│  └─────────────────────────┘    │
└─────────────────────────────────┘

文档直接在用户电脑生成，无需下载 ✅
```

##### 核心特点

| 特点 | 说明 |
|------|------|
| **文档位置** | 用户本地电脑（无需下载） |
| **传输协议** | STDIO（本地进程通信） |
| **离线使用** | 完全支持，无需网络 |
| **独立运行** | 每个用户独立运行完整服务 |

##### 与 B/S 架构的关键区别

| 对比项 | B/S 架构 | C/S 本地完整模式 |
|--------|----------|------------------|
| 文档生成位置 | 服务器 | **用户本地** |
| 获取文档方式 | 浏览器下载 | **直接在本地打开** |
| 网络依赖 | 必须联网 | **可离线使用** |
| 文件管理 | 服务器统一管理 | **本地文件系统** |

##### 适用场景

- 需要离线使用文档生成功能
- 文档涉及敏感信息，不希望上传服务器
- 个人工具或小团队使用
- 需要深度系统集成（文件管理器、右键菜单等）

##### 部署配置

**打包方案**：

| 方案 | 说明 | 优点 | 缺点 |
|------|------|------|------|
| **Python 环境打包** | 使用 PyInstaller 将 MCP Server 打包为可执行文件 | 用户无需安装 Python | 应用体积增大 |
| **内嵌 Python** | 随 Electron 分发精简版 Python | 可控性强 | 配置复杂 |
| **依赖用户 Python** | 要求用户预装 Python | 应用体积小 | 增加用户安装成本 |

**推荐方案**：使用 PyInstaller 打包，提供开箱即用体验

**客户端配置要点**：

```javascript
// Electron 主进程中配置 MCP Client
// MCP Server 可执行文件路径（打包在应用内）

const serverPath = path.join(app.getAppPath(), 'mcp-server', 'word_mcp_server.exe');

// 通过 STDIO 传输连接本地 MCP Server
// 具体实现参考官方 SDK 文档
```

##### 应用分发

| 平台 | 分发格式 | 说明 |
|------|----------|------|
| Windows | .exe / .msi | 支持自动更新 |
| macOS | .dmg | 需要代码签名 |
| Linux | .AppImage / .deb | 跨发行版兼容 |

---

#### 7.2.4 文档存储对比

| 对比项 | B/S 架构 | C/S 远程模式 | C/S 本地模式 |
|--------|----------|--------------|--------------|
| 存储位置 | 服务器 | 服务器 | 用户本地 |
| 获取方式 | 浏览器下载 | 客户端下载 | 直接访问 |
| 共享协作 | 天然支持 | 天然支持 | 需要额外方案 |
| 隐私安全 | 文档在服务器 | 文档在服务器 | 文档不离开本地 |
| 备份恢复 | 服务端统一备份 | 服务端统一备份 | 用户自行备份 |

---

### 7.3 架构选型指南

#### 决策流程

```
需要部署文档生成服务
        │
        ├─ 用户需要离线使用？
        │   └─ 是 → C/S 本地完整模式
        │
        ├─ 文档涉及敏感信息，不能上服务器？
        │   └─ 是 → C/S 本地完整模式
        │
        ├─ 希望用户零安装？
        │   └─ 是 → B/S 架构
        │
        ├─ 需要原生桌面体验？
        │   ├─ 是，但要集中管理 → C/S 远程服务器模式
        │   └─ 是，要本地生成 → C/S 本地完整模式
        │
        └─ 需要统一管理和维护？
            └─ 是 → B/S 架构
```

#### 三种架构全面对比

| 对比项 | B/S 架构 | C/S 远程模式 | C/S 本地模式 |
|--------|----------|--------------|--------------|
| **客户端** | 浏览器 | Electron | Electron |
| **MCP Client 位置** | 服务器 | 用户电脑 | 用户电脑 |
| **MCP Server 位置** | 服务器 | 服务器 | 用户电脑 |
| **文档生成位置** | 服务器 | 服务器 | 用户本地 |
| **传输协议** | HTTP/SSE | HTTP/SSE | STDIO |
| **用户安装** | 无需 | 需要 | 需要 |
| **离线使用** | 否 | 否 | **是** |
| **集中管理** | 是 | 是 | 否 |
| **维护成本** | 低 | 中 | 高 |
| **开发成本** | 低 | 中 | 高 |
| **用户体验** | Web | 原生 | 原生 |

#### 推荐场景

| 场景 | 推荐架构 | 原因 |
|------|----------|------|
| 企业内部快速部署 | B/S | 零安装、集中管理 |
| 重度文档用户 | C/S 远程模式 | 原生体验、集中存储 |
| 敏感文档处理 | C/S 本地模式 | 数据不离开本地 |
| 离线办公环境 | C/S 本地模式 | 无需网络 |
| 个人效率工具 | C/S 本地模式 | 简单直接 |
| SaaS 产品集成 | B/S | 易于集成、统一入口 |

---

## 附录：快速参考

### 常用命令

#### Windows

```cmd
# 自动安装
python setup_mcp.py
# 步骤1: 选择传输方式 → 输入 1 (STDIO)
# 步骤2: 选择安装方式 → 输入 1 (从 PyPI 安装)

# 如果选择了本地开发环境，激活虚拟环境
.venv\Scripts\activate

# 启动 Server (通常由 MCP Client 自动启动，无需手动运行)
python word_mcp_server.py

# 测试 HTTP (仅当使用 HTTP 传输时)
curl http://127.0.0.1:8000/mcp
```

#### Linux

```bash
# 自动安装
python3.11 setup_mcp.py
# 步骤1: 选择传输方式 → 输入 1 (STDIO)
# 步骤2: 选择安装方式 → 输入 1 (从 PyPI 安装)

# 如果选择了本地开发环境，激活虚拟环境
source .venv/bin/activate

# 启动 Server (通常由 MCP Client 自动启动，无需手动运行)
python word_mcp_server.py

# 测试 HTTP (仅当使用 HTTP 传输时)
curl http://127.0.0.1:8000/mcp
```

### 配置模板

#### .claude/mcp.json (Windows)

```json
{
  "mcpServers": {
    "word-document-server": {
      "command": "D:\\path\\to\\.venv\\Scripts\\python.exe",
      "args": ["D:\\path\\to\\word_mcp_server.py"],
      "env": {
        "PYTHONPATH": "D:\\path\\to\\Office-Word-MCP-Server",
        "MCP_TRANSPORT": "stdio"
      }
    }
  }
}
```

#### .claude/mcp.json (Linux)

```json
{
  "mcpServers": {
    "word-document-server": {
      "command": "/path/to/.venv/bin/python",
      "args": ["/path/to/word_mcp_server.py"],
      "env": {
        "PYTHONPATH": "/path/to/Office-Word-MCP-Server",
        "MCP_TRANSPORT": "stdio"
      }
    }
  }
}
```

#### cline_mcp_settings.json

```json
{
  "mcpServers": {
    "word-document-server": {
      "command": "python",
      "args": ["path/to/word_mcp_server.py"],
      "env": {
        "MCP_TRANSPORT": "stdio"
      },
      "alwaysAllow": [],
      "disabled": false
    }
  }
}
```

---

## 总结

本指南涵盖了 **Office-Word-MCP-Server** 在 Windows 和 Linux 环境下的完整安装、配置和使用流程，包括：

### ✅ 已涵盖内容

1. **MCP 架构理解**：Server、Client、Host 关系
2. **多平台安装**：Windows 和 Linux 详细步骤
3. **Client 配置**：Claude Code 和 VSCode+Cline
4. **传输方式**：STDIO、HTTP、SSE 详解
5. **使用示例**：实际操作演示
6. **故障排除**：常见问题和解决方案

### 🎯 快速上手路径

1. **第一次使用**：
   - Windows: 运行 `python setup_mcp.py`
   - Linux: 运行 `python3.11 setup_mcp.py`

   **两步选择**：
   - 步骤1: 选择传输方式 → 输入 `1` (STDIO，推荐)
   - 步骤2: 选择安装方式 → 输入 `1` (从 PyPI 安装，推荐)

2. **配置 Client**：
   - Claude Code: 创建 `.claude/mcp.json`
   - Cline: 通过 UI 添加 MCP Server

3. **测试验证**：
   - 创建测试文档
   - 验证工具可用性

### 📚 进一步学习

- **官方文档**: [GitHub Repository](https://github.com/GongRzhe/Office-Word-MCP-Server)
- **MCP 协议**: [Model Context Protocol](https://modelcontextprotocol.io/)
- **Claude Code**: [Anthropic Documentation](https://docs.anthropic.com)
- **Cline**: [Cline Documentation](https://docs.cline.bot)

---

**文档版本**: v2.0
**最后更新**: 2025年1月
**维护者**: AI Assistant
**反馈**: [GitHub Issues](https://github.com/GongRzhe/Office-Word-MCP-Server/issues)

---

**祝使用愉快！** 🚀
