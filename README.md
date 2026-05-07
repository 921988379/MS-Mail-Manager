# Refresh Token 邮件验证码 Web 工具

一个轻量级自托管 Python Web 工具，用来保存 Microsoft 账号 Refresh Token，并读取最近的验证码邮件。

适合场景：

- 管理多个 Microsoft / Outlook 账号的 Refresh Token
- 通过 Microsoft Graph `Mail.Read` 读取收件箱和垃圾邮箱中的验证码邮件
- 在 Graph 不可用时回退到 IMAP XOAUTH2
- 在一个简单网页里查看已保存账号、批量导入账号、刷新验证码邮件

## 功能

- 基于 Python 标准库实现的单文件 HTTP 服务
- 使用 SQLite 保存账号数据
- 支持 Microsoft OAuth Refresh Token 换取 Access Token
- 优先使用 Microsoft Graph `Mail.Read` 检索验证码邮件
- 支持 IMAP XOAUTH2 兜底
- 支持保存账号、选择账号、批量导入
- 支持登录会话 Cookie
- 内置测试，覆盖 OAuth scope、Graph/IMAP 邮件读取、页面渲染和安全辅助函数

## 安全提醒

这个项目会处理非常敏感的数据，例如：

- Refresh Token
- Access Token
- 邮箱账号
- 可选的邮箱密码
- 验证码邮件内容

请不要提交或公开以下内容：

- `.env`
- `app.db`
- `*.db`
- `*.sqlite`
- Refresh Token
- Access Token
- 邮箱密码
- 含敏感信息的生产日志

建议只部署在自己的私有服务器上，默认绑定 `127.0.0.1`，如果要暴露到公网，请放在带认证的反向代理后面，并设置强密码。

## 配置

支持的环境变量：

- `RTWEB_DB`：SQLite 数据库路径，默认 `app.db`
- `RTWEB_ADMIN_PASSWORD` 或 `RTWEB_PASSWORD`：后台登录密码，默认 `change-me`
- `RTWEB_SESSION_SECRET`：会话签名密钥，默认每次进程启动随机生成
- `RTWEB_HOST`：监听地址，默认 `127.0.0.1`
- `RTWEB_PORT`：监听端口，默认 `8020`
- `RTWEB_COOKIE_SECURE`：Cookie 是否只允许 HTTPS；本地 HTTP 开发时可设为 `0`

示例：

```bash
export RTWEB_PASSWORD='change-this'
export RTWEB_SESSION_SECRET='use-a-long-random-secret'
python3 app.py
```

本地开发如果没有 HTTPS：

```bash
export RTWEB_COOKIE_SECURE=0
python3 app.py
```

## 批量导入格式

支持类似下面的格式，每行一个账号：

```text
邮箱----密码----Client ID----Refresh Token
```

也兼容不带密码的格式：

```text
邮箱----Client ID----Refresh Token
```

页面展示时会对密码和 Token 做打码处理。

## 测试

```bash
pytest tests -q
```

## 许可

这是一个自托管工具项目。公开使用前请自行补充 License，并自行承担账号与 Token 的安全管理责任。
