# m365-serve：零 npm 依赖的 M365 Copilot 本地代理

从 [chrischall/opencode-copilot-plugin](https://github.com/chrischall/opencode-copilot-plugin) 提取必要的代理功能，补上仅使用 Node.js 标准库的登录、OAuth 令牌缓存/刷新和 WebSocket 传输。

**需要 Node.js 22 或更高版本，以及本机已安装的 Microsoft Edge。推荐 Node.js 24 LTS。无需 npm install，无需构建。**

“零第三方依赖”指没有 npm 运行依赖、开发依赖、SDK、Playwright、MSAL、OpenCode 或第三方中转服务。浏览器和微软在线服务仍然必需。这不是离线模型；需要有 M365 Copilot 访问权限的组织账号。代码基于上游 MIT 源码，不代表没有使用开源代码。

## Windows 快速开始

解压到自己的目录，例如 `D:\tools\m365-serve`，打开 PowerShell：

```powershell
cd D:\tools\m365-serve
node --version
node src/cli.js login
node src/cli.js serve --port 4141
```

也可依次双击 `login.cmd` 和 `serve.cmd`。保持 serve 窗口运行。

1. `login` 自动找到 Microsoft Edge，打开独立浏览器配置目录。
2. 在浏览器完成组织账号登录、SSO 和 MFA；工程不收集密码或 MFA 种子。
3. 终端显示“登录成功”后启动 `serve`。输出 OpenAI Base URL 和 API key。
4. 客户端填 Base URL `http://127.0.0.1:4141/v1`、终端打印的 API key，以及模型名 `m365-copilot`。

默认端口 4141，只监听本机。Ctrl+C 停止服务。每次启动生成新 API key；固定 key 可在启动前设置：

```powershell
$env:M365_PROXY_KEY = '请替换为你自己的随机长密钥'
node src/cli.js serve
```

验证存活不需要密钥：

```powershell
Invoke-RestMethod http://127.0.0.1:4141/health
```

验证真实模型调用（替换 API key）：

```powershell
$headers = @{ Authorization = 'Bearer 替换为终端打印的API-key' }
$body = @{ model = 'm365-copilot'; messages = @(@{role='user'; content='Hello'}) } | ConvertTo-Json -Depth 10
Invoke-RestMethod -Method Post -Uri http://127.0.0.1:4141/v1/chat/completions -Headers $headers -ContentType 'application/json; charset=utf-8' -Body ([Text.Encoding]::UTF8.GetBytes($body))
```

登录固定使用 Microsoft Edge，不会回退到 Chrome，也不再读取 `CHROMIUM_PATH`。Edge 没有自动找到时：

```powershell
$env:EDGE_PATH = 'C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe'
node src/cli.js login
```

## 保留与删除

保留：

- `login`：交互式 OAuth 授权码 + PKCE，校验随机 state，拦截微软 nativeclient 回调。
- `serve`：`GET /health`、`GET /v1/models`、`POST /v1/chat/completions`。
- 普通/流式回复、模型 tone 映射、会话复用、同一会话串行请求。
- Markdown fence 工具调用转换和默认工具裁剪；工具由调用方执行，代理不执行本地命令。
- 图片 data URL 上传；远程图片 URL 不会被下载。上传失败沿用上游的降级行为。
- 本地标题模型、Bearer 鉴权、Host 校验、拒绝浏览器 Origin、请求体限制。

删除：OpenCode 插件与配置写入、setup、doctor、账号密码自动填写、TOTP、自动创建/发布 Copilot Studio Agent、npm 依赖、打包与编译链。

未自动创建 Agent 时工具调用的可靠性可能低于上游默认配置。已有 Agent 可通过 `M365_AGENT_ID` 指定；普通聊天不需要它。`--no-lean` 可关闭工具裁剪，`--multi-tool` 可允许一次返回多个工具调用。

采样参数 temperature/top_p/max_tokens 沿用上游行为：接受但忽略；这是协议桥接，不是完整 OpenAI API。具体模型可用性由微软租户决定，模型列表继承上游，不代表本工程已逐个验证。上游使用未公开的微软协议，服务变化可能要求调整代码。

## 缓存与刷新

默认目录：`%USERPROFILE%\.config\m365-serve`（其他系统是 `~/.config/m365-serve`）。

- `tokens.json`：access token、refresh token、过期时间。
- `browser-profile`：专用 Edge 登录配置，不修改默认浏览器配置。

令牌将在请求时提前两分钟刷新，刷新令牌轮换后原子写入缓存。条件访问、撤销授权或刷新失败时需重新运行 `login`。不读取上游 MSAL 缓存。

可用 `M365_CONFIG_DIR` 更改目录。此目录含凭据，不要提交或分享；Windows 使用当前用户目录权限，POSIX 创建目录/文件时使用 0700/0600。停止服务并删除此目录可清除本地缓存；服务端撤销需在微软账号中操作。

仍可选择通过 `M365_TOKEN_FILE` 提供纯文本令牌或含 `access_token` 的 JSON，或者设置 `M365_ACCESS_TOKEN`。这些覆盖登录缓存，不自动刷新；令牌文件每回合重新读取。

## 测试与验证范围

```powershell
node --test
```

已在 Node.js 24 / Linux 验证 8 项测试：HTTP 鉴权和请求校验；HTTP → 自研 WebSocket → SignalR 全链路；Metrics 与 chat 同包；TCP/WS 分片；流式输出；会话复用；工具返回转换；Disengaged；PKCE/state 校验；缓存和并发刷新、刷新令牌轮换。

还用真实 Chromium + 本地构造的授权回调验证了 CDP 回调拦截、授权码传递及缓存调用。此测试仅验证共用的 CDP 协议，不能替代真实 Edge 验证；测试使用临时包装器，正式登录仅查找 Edge。该浏览器测试在没有 `/usr/bin/chromium` 的系统上跳过。

**尚未在 Windows 上运行，也未使用真实微软账号验证租户登录或模型调用。** 登录逻辑遵循上游的客户端 ID、scope、redirect URI 和授权流程，但组织权限、管理员策略、浏览器调试限制可能影响结果。若报错，保留错误文字即可排查，不要发送 tokens.json、授权回调 URL 或密码。首次实际运行建议先用 `m365-copilot` 做普通聊天验证。

WebSocket 当前直接连接微软，不实现 HTTP CONNECT/SOCKS 代理；只允许强制代理出网的环境可能需要额外适配。不要通过禁用 TLS 证书验证来处理证书问题。

## 文件说明

- `src/cli.js`：login / serve 命令。
- `src/login.js`：浏览器发现、CDP 授权回调拦截。
- `src/oauth.js`：PKCE、微软 token 请求、缓存/刷新。
- `src/auth.js`：手动令牌输入，认证错误。
- `src/server.js`：OpenAI HTTP 接口。
- `src/session.js`、`src/websocket.js`：SignalR 和标准库 WebSocket 客户端。
- `src/translate.js`、`src/fenced.js`：消息、流式结果及工具调用转换。
- `src/config.js`、`src/models.js`、`src/image.js`、`src/validation.js`：必要辅助逻辑。

## 来源与许可

上游提交：`5a9c46061f0eb7bd6228a9ef277290d2def27a61`。

复用/改写 server、session、translate、fenced、image、models，以及 config 中的工具选择逻辑；使用 Node 自带类型擦除生成可直接运行的 JavaScript。已替换 WebSocket、验证、认证和日志模块，增加超时及断线处理。原版权和 MIT 许可保留在 LICENSE。自研 WebSocket 面向该文本协议，不作为通用 WebSocket 库交付。
