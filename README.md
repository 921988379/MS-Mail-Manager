# 一点微软工具箱（MS Mail Manager）

一个轻量级、自托管的 Microsoft / Outlook / Hotmail 邮箱管理工具箱，用来管理邮箱账号、Refresh Token、Access Token、验证码邮件、账号状态检测和外部查询 API。

- 官网/线上入口：`https://token.seoyh.net/`
- 仓库：`https://github.com/921988379/MS-Mail-Manager`
- 当前版本：`1.0.5`
- 版权支持：由 [一点优化](https://www.seoyh.net/) 提供

> 安全提醒：本项目会处理邮箱密码、Refresh Token、Access Token、验证码邮件等敏感数据。请只部署在自己可信服务器上，并妥善保存 `RTWEB_DATA_KEY`、数据库和环境变量文件。

## 功能特色

- 邮箱账号托管：单个/批量导入、分类管理、批量导出、删除账号
- Refresh Token 管理：检测状态、获取 Access Token、保存轮换后的 Refresh Token
- 邮件验证码读取：优先 Microsoft Graph，失败后回退 XOAUTH2 IMAP，再使用邮箱密码 IMAP 兜底
- 账号综合检测：统一检测 token、Graph、密码 IMAP、验证码摘要和最佳来源
- 批量操作进度条：实时显示总进度、当前处理邮箱、成功/失败数量
- 状态实时同步：批量执行时当前页面邮箱状态同步更新
- 外部 API：API Key、scope 权限、限流、调用日志、latest-code/account-status 查询
- 敏感字段加密：Refresh Token、邮箱密码、辅助密码支持加密存储
- 自动更新：后台版本页检查 GitHub 最新提交，可选启用安全更新命令

## 快速开始

### 1. 克隆项目

```bash
git clone https://github.com/921988379/MS-Mail-Manager.git
cd MS-Mail-Manager
```

### 2. 安装依赖

```bash
python3 -m venv .venv
source .venv/bin/activate
pip install -r requirements.txt
```

### 3. 配置环境变量

复制示例配置：

```bash
cp .env.example .env
```

请至少修改：

```text
RTWEB_DB=/www/server/rtweb/app.db
RTWEB_LOGIN_USERNAME=你的登录用户名
RTWEB_LOGIN_PASSWORD=第一层登录密码
RTWEB_ADMIN_PASSWORD=第二层管理密码
RTWEB_SESSION_SECRET=长随机字符串
RTWEB_DATA_KEY=长随机字符串，务必保存
RTWEB_API_ALLOW_QUERY_KEY=0
```

`RTWEB_DATA_KEY` 用于解密数据库里的敏感字段。**丢失后已加密的 Refresh Token / 密码无法恢复。**

### 4. 启动

本地开发：

```bash
set -a
source .env
set +a
python3 app.py
```

默认监听：

```text
http://127.0.0.1:18080
```

生产环境建议使用 systemd + Nginx HTTPS 反代。

## systemd 部署示例

`/etc/systemd/system/rtweb.service`：

```ini
[Unit]
Description=MS Mail Manager
After=network-online.target
Wants=network-online.target

[Service]
Type=simple
WorkingDirectory=/www/wwwroot/refresh-token-mail-web
EnvironmentFile=/www/server/rtweb/rtweb.env
ExecStart=/usr/bin/python3 /www/wwwroot/refresh-token-mail-web/app.py
Restart=always
RestartSec=3
User=rtweb
Group=rtweb
NoNewPrivileges=true
PrivateTmp=true
ProtectSystem=full
ReadWritePaths=/www/server/rtweb

[Install]
WantedBy=multi-user.target
```

启动：

```bash
systemctl daemon-reload
systemctl enable --now rtweb
systemctl status rtweb --no-pager -l
```

## Nginx 建议

建议启用 HTTPS，并阻止敏感文件访问：

```nginx
location ~* /(\.git|\.env|.*\.db|.*\.sqlite|.*\.sqlite3) {
    deny all;
}

location / {
    proxy_pass http://127.0.0.1:18080;
    proxy_set_header Host $host;
    proxy_set_header X-Real-IP $remote_addr;
    proxy_set_header X-Forwarded-For $proxy_add_x_forwarded_for;
    proxy_set_header X-Forwarded-Proto $scheme;
    proxy_read_timeout 300s;
    proxy_send_timeout 300s;
}
```


## 使用说明（完整流程）

### 1. 导入邮箱

入口：后台 `/mailboxes` → “批量导入”。一行一个邮箱：

```text
邮箱----密码----应用ID(Client ID)----Refresh Token----辅助邮箱----辅助密码----分类/项目
```

字段说明：

| 字段 | 是否必填 | 说明 |
| --- | --- | --- |
| 邮箱 | 必填 | Outlook / Hotmail 邮箱地址 |
| 密码 | 可选 | 用于密码 IMAP 兜底读取验证码；不会用于生成 Refresh Token |
| Client ID | 令牌账号必填 | Microsoft OAuth 应用 ID |
| Refresh Token | 令牌账号必填 | 用于刷新 Access Token，保存后加密入库 |
| 辅助邮箱/密码 | 可选 | 仅作为账号资料保存 |
| 分类/项目 | 可选 | 用于分组、API 按项目取码 |

兼容分隔符：`----`、`|`、逗号、Tab。旧格式 `邮箱----Client ID----Refresh Token` 也支持。

### 2. 令牌类型说明

| 类型 | 用途 | 适合场景 |
| --- | --- | --- |
| Graph 令牌 | 通过 Microsoft Graph `Mail.Read` 读取邮件 | 最推荐，稳定、速度快、可读收件箱和垃圾箱 |
| OAuth IMAP 令牌 | Access Token 通过 IMAP XOAUTH2 登录 | Graph 不可读时的兼容方案 |
| 密码 IMAP | 保存邮箱密码后用 IMAP 直接读取 | 令牌失效时兜底判断账号是否还能收信 |
| 无令牌账号 | 只保存邮箱/密码/分类 | 不能刷新 token；如有密码可尝试 IMAP 取码 |

系统读取验证码顺序：

```text
Graph → OAuth IMAP → 密码 IMAP
```

综合检测会告诉你哪个通道可用。

### 3. 刷新令牌 / 获取 Access Token

| 功能 | 入口 | 说明 |
| --- | --- | --- |
| 单个刷新 | `/tokens` | 输入 Client ID + Refresh Token，换取 Access Token |
| 批量获取令牌 | 邮箱管理 → 批量操作 → 获取令牌 | 后台执行，显示进度条和当前处理邮箱 |
| 批量刷新令牌 | 邮箱管理/令牌管理 | 异步执行，避免大批量请求 502 |
| 状态检测 | 检测状态 / 批量检测 | 验证 Refresh Token 是否可换 Access Token |

如 Microsoft 返回新的 Refresh Token，系统会自动轮换保存。Token 结果默认隐藏，只提供“显示/复制”按钮，降低误泄露风险。

### 4. 获取邮件和验证码

入口：`/mails` 或外部 API `/api/v1/latest-code`。

流程：

1. 选择邮箱，或 API 指定 `email/category`
2. 用 Refresh Token 换 Access Token
3. 优先通过 Microsoft Graph 读取收件箱/垃圾邮箱
4. Graph 失败后尝试 OAuth IMAP
5. 如令牌失败且保存了密码，再尝试密码 IMAP
6. 从最新邮件主题/摘要中识别验证码

取不到验证码不一定代表账号挂了，可能是最近没有验证码邮件、规则不匹配、邮件延迟或 Microsoft 风控。

### 5. 项目/分类管理

入口：`/project-manage`、`/categories`。

| 能力 | 用途 |
| --- | --- |
| 分类 | 给邮箱分组，批量导入时最后一列可直接写分类 |
| 项目规则 | 按项目配置邮箱组、关键词过滤和返回数量 |
| API 按项目查询 | `/api/v1/latest-code?category=项目名` |
| 批量改分类 | 勾选邮箱后可批量设置分类 |

### 6. 外部 API 使用

推荐统一使用请求头传 API Key：

```text
X-API-Key: ***
```

不要在 URL 明文传 key。

常用接口：

```bash
curl -H "X-API-Key: ***" "https://token.seoyh.net/api/v1/health"
curl -H "X-API-Key: ***" "https://token.seoyh.net/api/v1/projects"
curl -H "X-API-Key: ***" "https://token.seoyh.net/api/v1/accounts?category=项目名"
curl -H "X-API-Key: ***" "https://token.seoyh.net/api/v1/latest-code?email=xxx@outlook.com"
curl -H "X-API-Key: ***" "https://token.seoyh.net/api/v1/latest-code?category=项目名"
curl -H "X-API-Key: ***" "https://token.seoyh.net/api/v1/account-status?email=xxx@outlook.com"
```

| 接口 | scope | 说明 |
| --- | --- | --- |
| `/health` | `health` | 健康检查 |
| `/projects` | `projects` | 返回项目/分类列表 |
| `/accounts` | `accounts` | 返回账号元数据和状态，不返回密码/token |
| `/latest-code` | `latest_code` | 读取最新验证码 |
| `/account-status` | `accounts + latest_code` | 返回综合状态摘要 |

业务失败会返回 `ok:false` 和错误信息；服务本身正常时尽量不把账号失败伪装成 HTTP 502。

### 7. 免导入接口说明

当前版本默认不提供“外部直接传邮箱密码/Refresh Token 的免导入取码接口”。

原因：这类接口会让敏感凭证经过 URL、日志或第三方系统，安全风险更高。

推荐做法：先在后台导入并加密保存账号，再通过 API Key + `email/category` 调用。这样外部系统不需要持有邮箱密码或 Refresh Token。

### 8. 常见问题

| 问题 | 原因 | 处理建议 |
| --- | --- | --- |
| `invalid_grant` | Refresh Token 被撤销、过期或账号安全状态变化 | 重新 OAuth 授权获取新 Refresh Token |
| `AADSTS70000` | scope 未授权或授权过期 | 系统会兼容旧授权重试；仍失败时重新授权当前 scope |
| Graph 失败但密码 IMAP 可读 | Graph 权限/令牌异常，账号本身可能还能收信 | 继续用密码 IMAP 兜底，同时重新生成令牌 |
| OAuth IMAP 提示 authenticated but not connected | 账号未初始化 Exchange mailbox 或邮箱服务不可用 | 登录 Outlook 网页版初始化，或改用 Graph/密码 IMAP |
| API Key 无效 | Key 错误、被禁用或 scope 不足 | 到 API 密钥页检查启用状态、scope 和限流 |
| 批量任务很慢 | Microsoft 网络、邮箱数量、IMAP 超时都会影响速度 | 观察进度条；系统已限制并发并异步执行，避免 502 |

## 使用说明（简版入口）

### 登录

后台为双层验证：

1. 用户账号 + 用户密码
2. 应用管理密码

对应环境变量：

```text
RTWEB_LOGIN_USERNAME
RTWEB_LOGIN_PASSWORD
RTWEB_ADMIN_PASSWORD
```

### 邮箱导入

入口：`/mailboxes`

批量导入格式：

```text
邮箱----密码----Client ID----Refresh Token----辅助邮箱----辅助密码----分类
```

兼容旧格式：

```text
邮箱----Client ID----Refresh Token
```

支持分隔符：

```text
----
|
逗号
Tab
```

### 令牌管理

入口：`/tokens`

支持：

- 使用 Client ID + Refresh Token 获取 Access Token
- 检测 Refresh Token 是否可用
- 保存账号并更新轮换后的 Refresh Token
- 批量获取、批量刷新、批量检测
- 进度条显示当前处理邮箱

默认 scope：

```text
offline_access https://graph.microsoft.com/Mail.Read
```

### 获取验证码

入口：`/mails`

读取顺序：

```text
1. Refresh Token -> Access Token
2. Microsoft Graph 读取收件箱/垃圾箱
3. XOAUTH2 IMAP 读取收件箱/垃圾箱
4. 如果令牌失败且保存了密码，则用邮箱密码 IMAP 兜底
```

### 账号综合检测

入口：邮箱管理里的“综合检测”按钮，或批量操作里的“综合检测”。

判断参考：

| 现象 | 判断 | 建议 |
| --- | --- | --- |
| `token_failed + 密码 IMAP 可读` | 令牌挂了，账号未必挂 | 可继续用 IMAP 取码，重新 OAuth 授权刷新令牌 |
| `token_failed + 密码 IMAP 失败` | 账号/密码/风控可能异常 | 登录 Outlook 网页版验证或重新导入 |
| `Graph 可读` | 令牌和邮件权限可用 | 可正常用 API 取码 |

### 状态说明

| 状态 | 显示 | 含义 |
| --- | --- | --- |
| `token_ok` | 令牌可用 | Refresh Token 能换 Access Token |
| `token_failed` | 令牌失效 | Refresh Token 不可用 |
| `graph_ok` | Graph 可读 | Graph Mail.Read 能读取邮件 |
| `graph_failed` | Graph 失败 | Graph 读取失败 |
| `xoauth2_imap_ok` | OAuth IMAP 可读 | Access Token 可用于 IMAP XOAUTH2 |
| `xoauth2_imap_failed` | OAuth IMAP 失败 | XOAUTH2 IMAP 不可用 |
| `imap_password_ok` | 密码 IMAP 可读 | 邮箱密码可通过 IMAP 读取邮件 |
| `imap_password_failed` | 密码 IMAP 失败 | 邮箱密码 IMAP 不可用 |
| `all_failed` | 全部失败 | 所有通道都不可用 |

## 外部 API

推荐只使用请求头传 API Key：

```text
X-API-Key: ***
```

线上建议关闭 URL 参数传 key：

```text
RTWEB_API_ALLOW_QUERY_KEY=0
```

### 示例

```bash
curl -H "X-API-Key: ***" "https://token.seoyh.net/api/v1/health"

curl -H "X-API-Key: ***" "https://token.seoyh.net/api/v1/projects"

curl -H "X-API-Key: ***" "https://token.seoyh.net/api/v1/accounts?category=项目名"

curl -H "X-API-Key: ***" "https://token.seoyh.net/api/v1/latest-code?email=xxx@outlook.com"

curl -H "X-API-Key: ***" "https://token.seoyh.net/api/v1/latest-code?category=项目名"

curl -H "X-API-Key: ***" "https://token.seoyh.net/api/v1/account-status?email=xxx@outlook.com"

curl -H "X-API-Key: ***" "https://token.seoyh.net/api/v1/account-status?category=项目名"
```

### API scopes

| Scope | 用途 |
| --- | --- |
| `health` | 健康检查 |
| `latest_code` | 查询验证码 |
| `projects` | 查询项目/分类 |
| `accounts` | 查询账号元数据和状态 |

`/api/v1/account-status` 需要同时具备：

```text
accounts + latest_code
```

API 不返回邮箱密码、Refresh Token、Access Token 明文。

## 自动更新功能

后台入口：`/version`

支持：

- 查看当前应用版本
- 查看当前 Git 分支/commit
- 查看 GitHub 仓库地址
- 检查远程最新提交
- 可选执行更新命令

### 默认安全策略

自动更新默认关闭：

```text
RTWEB_AUTO_UPDATE_ENABLED=0
```

这样后台只能“检查更新”，不能直接覆盖线上代码。

### 启用自动更新

确认你已经做好数据库/env 备份，并确认更新命令安全后，再设置：

```text
RTWEB_AUTO_UPDATE_ENABLED=1
RTWEB_UPDATE_REPO=https://github.com/921988379/MS-Mail-Manager.git
RTWEB_UPDATE_BRANCH=main
RTWEB_UPDATE_COMMAND=./scripts/update.sh
```

更新脚本：

```bash
./scripts/update.sh
```

脚本会执行：

1. `git fetch`
2. `git pull --ff-only`
3. `python3 -m py_compile app.py`
4. 使用临时数据库运行 unittest
5. 尝试重启 `rtweb` 服务

也可以手动更新：

```bash
cd /www/wwwroot/refresh-token-mail-web
git fetch origin main
git pull --ff-only origin main
python3 -m py_compile app.py
TMPDB=$(mktemp /tmp/ms-mail-manager-test-db.XXXXXX)
rm -f "$TMPDB"
RTWEB_DB="$TMPDB" python3 -m unittest discover -s tests -p 'test_*.py'
rm -f "$TMPDB"
systemctl restart rtweb
```

## 安全设计

- 后台登录保护
- 双层密码验证
- API Key 保护外部 API
- API Key 只存 SHA-256 digest
- CSRF 防护
- 登录失败限流
- 安全响应头
- 请求日志脱敏
- Refresh Token、邮箱密码、辅助密码加密存储
- 令牌结果默认隐藏，需要点击显示或复制
- 批量任务并发上限，降低 502 和资源耗尽风险
- `.env`、数据库、缓存、密钥文件默认不提交

加密字段：

```text
saved_accounts.refresh_token
saved_accounts.password
saved_accounts.aux_password
```

加密前缀：

```text
enc:v1:
```

关键环境变量：

```text
RTWEB_DATA_KEY
```

## 测试

语法检查：

```bash
python3 -m py_compile app.py
```

使用临时数据库运行测试：

```bash
TMPDB=$(mktemp /tmp/ms-mail-manager-test-db.XXXXXX)
rm -f "$TMPDB"
RTWEB_DB="$TMPDB" python3 -m unittest discover -s tests -p 'test_*.py'
rm -f "$TMPDB"
```

## 发布前检查清单

```bash
python3 -m py_compile app.py
TMPDB=$(mktemp /tmp/ms-mail-manager-test-db.XXXXXX)
rm -f "$TMPDB"
RTWEB_DB="$TMPDB" python3 -m unittest discover -s tests -p 'test_*.py'
rm -f "$TMPDB"
git status --short
```

确认不要提交：

- `.env`
- `app.db`
- `*.db`
- `*.sqlite`
- 生产日志
- 任何真实 token / 密码 / API Key

## 版权说明

© 一点微软工具箱。技术与优化支持：[一点优化](https://www.seoyh.net/)。
