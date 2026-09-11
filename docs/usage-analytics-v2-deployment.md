# V2 使用统计：部署与验收

入口为 `/admin/usage`（用户使用）、`/admin/api`（API 用量），两个页面独立于 Office，只读访问。

## 当前配置

按 2026-09-11 确认的要求，后端保持单实例，存储改为 PostgreSQL，浏览器登录复用 `.env` 中现有 Entra 应用。

- 统计数据库：`office_translation_analytics`。
- 应用数据库账号：`office_translation_analytics_app`，没有超级用户、建库或建角色权限。
- schema：`usage_analytics`；业务表为 `users`、`api_usage_daily`，技术元数据表为 `meta`。
- 连接串保存在本机 `.env` 的 `ANALYTICS_DATABASE_URL`，不提交到 Git，也不写入日志。
- 管理员租户：`d7125684-0e28-40b5-aba2-ea9580f2a201`；邮箱：`jason@majun.fun`；仅允许查询该租户。

新环境运行 `npm ci`、`npm run build` 后启动后端。PostgreSQL 需先有数据库与可用账号；应用在连接成功后执行 schema 版本检查和事务建表，不使用超级用户创建生产数据库。

## 本次服务器更新步骤

1. 同步代码和 `package.json`、`package-lock.json`，包括新增的 `server/analytics*.js`、`server/admin.js` 和 `src/admin/`。
2. 将本机 `.env` 中的 `ANALYTICS_DATABASE_URL`、`ANALYTICS_ADMINS` 安全复制到服务器环境配置，保留服务器已有的其他配置。`.env` 被 Git 忽略，不会随代码同步；不要将它提交到仓库。复用已有应用，无需另设 `ADMIN_CLIENT_ID`。
3. 在服务器项目目录执行 `npm ci`、`npm run build`，然后按现有方式重启单实例 Node 服务。服务器必须能够连接该 PostgreSQL 地址和端口。
4. 打开 `https://www.majun.fun:30260/admin/usage`，以已配置管理员登录；再检查 `/admin/api`。新统计从采集开始累计，不补算旧版本历史请求。

本地已通过 435 项测试、类型检查、生产构建，验证了分页、平台与客户端筛选、API 汇总、真实数据库连接和未登录访问拦截。构建有管理页面脚本 267 KiB 的体积提示，不影响构建成功。页面交互测试使用隔离的合成数据，没有写入真实统计库。

## 浏览器登录与授权

`ADMIN_CLIENT_ID` 留空时使用现有 `CLIENT_ID`。在该 Entra 应用的身份验证配置中添加 **SPA** 重定向地址 `${APP_BASE_URL}/admin/usage`，保持已有 Office SSO 配置与其他回调地址。管理页面与 API 保持同源及 HTTPS。

浏览器采用 MSAL 授权码 + PKCE。复用同一应用时 scope 为 `${CLIENT_ID}/access_as_user`，使用 GUID 标识自身 API；配置独立浏览器应用时使用 `${SSO_RESOURCE}/access_as_user`。后端沿用现有签名、发行者、受众、有效期和 scope 校验；Graph 令牌与 ID Token 不能代替本 API 访问令牌。多租户登录需应用注册支持；也可将 `ADMIN_LOGIN_TENANT` 设为明确租户 GUID。

`ANALYTICS_ADMINS` 为服务端 JSON 白名单，示例结构见 `.env.example`：

- 优先支持真实 `tenant_id + user_oid`，不需为了授权额外获取 Graph 资料。
- 也支持 `tenant_id + email`：后端使用已验证令牌获取 Graph `/me`，核验返回 ID 与令牌对象 ID、租户一致后，再按邮箱授权。沿用已有 OBO、`User.Read` 权限及同意，资料最多缓存 5 分钟。不能根据客户端字段或未经核验的邮箱声明放行。
- 各管理员只能查询配置中的租户；无通配符和默认管理员，登录成功不等于有报表权限。
- 白名单修改后重启。网页不提供配置或统计删除接口。

用户已确认添加 `https://www.majun.fun:30260/admin/usage` 的 SPA 回调地址。当前应用读取注册元数据返回 403，未独立读取确认；真实管理员浏览器登录待用户同步服务器代码后验收。本次只完成本地代码和数据库配置，不更新线上 Node 服务。

登录参考：[微软 MSAL Browser 初始化](https://learn.microsoft.com/en-us/entra/msal/javascript/browser/initialization)。

## 存储与可靠性

`usage_analytics.meta` 保存 schema 版本、采集起点和已知最近缺口时间，不含个人信息。当前 schema 版本为 1，迁移在事务中完成，发现其他版本会拒绝统计访问，不覆盖已有数据。

PostgreSQL 使用独立异步连接、有界串行任务队列；最多 500 个任务、4 个管理查询。连接超时 2 秒，服务端语句超时 3 秒，客户端查询超时 4 秒，报表调用等待上限 5 秒；超时排队任务仍占用位置直到实际处理完成。数据库故障不会使翻译失败，也不自动重放结果不确定的统计写入。故障后暂停连接重试 5 秒，后续任务可重新连接；健康日志只输出固定提示。

连接持有 PostgreSQL 会话级 advisory lock，避免第二个服务进程或维护命令并发写入同一统计库。进程停止或连接结束后数据库自动释放锁，不需要本地锁文件。维护前仍必须停止全部后端／开发服务器，排空业务请求，不能在服务断线期间绕过这一要求。

每次请求结束后一次性提交汇总，失败也计数，日期归属自身服务端开始时刻的北京时间自然日。业务翻译日汇总与首次／最近使用时间在同一事务中更新。报表读取使用只读、可重复读事务，避免多个图表读取到不同提交状态。

启动及每天清理 12 个月前的日汇总，普通清理保留用户首次时间。请求未结束或进程崩溃时可能未计入；已知缺口在后续数据库可写时保存。采集前、已清理或已丢失的历史不显示为确认的零用量。

应用不新增备份、导出或恢复任务。PostgreSQL 服务器的自动备份、归档和快照范围仍需维护人员核实，不因应用没有备份任务就宣称没有恢复副本。回退版本不得覆盖更高版本 schema，也不为回退复制统计库；旧插件没有平台字段仍可请求，归入 unknown。

## 删除统计用户

由获授权维护人员确认真实账号和影响范围，停止所有后端进程并等候在途请求结束，再运行：

```text
npm run analytics:delete-user -- <tenant_id> <user_oid> --confirm
```

命令取得相同数据库会话锁后，在同一事务中清理两张业务表。锁被占用、参数错误或数据库不可用时不报告成功。不提供网页撤销；重启后的独立新使用继续采集并建立新的首用历史。邮箱查询权限不等于维护权限。

## 验证

运行 `npm test -- --runInBand`、`npm run typecheck`、`npm run build`。统计 SQL 测试通过 PGlite 的 PostgreSQL 引擎在隔离内存库中执行，测试脚本启用 Node VM 模块；不向真实统计库注入测试账号或模拟用量。

真实环境还需确认管理员普通浏览器登录、Word / Outlook 网页与桌面平台映射、实际分批调用，以及重启后统计连续性。数据库创建和空报表查询可单独验证，但不替代 Entra 与真实 Office 验收。
