# 统计后台安全审查

## 修复进展（2026-09-12）

第 1–3 项中危问题已在本地 `feat/api-usage-top10` 分支修复，尚未部署；下文保留原审查证据与当时版本说明。

- 第 1 项：授权只匹配 `tenant_id + user_oid`，删除登录时的邮箱匹配和缓存。旧邮箱条目不再授予权限，设置 API 拒绝保存邮箱授权，页面改为填写对象 ID。部署前须确保至少一位对象 ID 权限管理员可用，具体步骤见 [管理员迁移说明](docs/admin-settings.md)。
- 第 2 项：未认证 IP 配额与验证后的 `tid + oid` 配额分离。同出口匿名/无效令牌配额耗尽后，合法管理员仍可访问；不同账号和租户相互独立，同账号更换令牌不能重置配额。保持不信任客户端代理头；入口层验签资源保护和实际代理配置仍属部署要求。
- 第 3 项：权限变更与审计差异在同一事务内持久化，记录已验证操作者、数据库时间和版本。审计写入失败则回滚权限和版本。触发器拒绝修改、删除和清空审计，撤销 PUBLIC 权限；数据库所有者权限仍需由部署侧控制。

验证：针对性测试 8 个套件、63 项通过；随后全量 `npm test -- --runInBand` 的 31 个套件、467 项测试全部通过；`npm run typecheck` 通过。验证覆盖邮箱复用、旧条目迁移、匿名/无效令牌限流、不同账号/租户配额、增删改审计、操作者伪造、版本冲突、审计失败回滚和追加记录保护。未执行生产数据库迁移或线上角色验收。低危浏览器加固和未证实可利用的依赖公告不在此次修复范围。

审查日期：2026-09-12。目标：https://www.majun.fun:30260/admin/usage 。

## 结论与范围

未发现已证实的匿名读取统计数据、直接 SQL 注入、前端字段伪造管理员或跨租户查询绕过。发现 3 项应处理的安全风险，其中邮箱授权在账号回收/重建等条件下可能导致高权限错误授予；另有浏览器防护和依赖更新事项。此结论不等于完整渗透测试通过。

本地当前为 main，提交 `84dc069`。线上返回的 HTML 包含管理员设置和 TOP10，而 main 没有这些功能；因此补充审查了本地 `feat/api-usage-top10` 的 `2b6077d`。线上功能与该分支相符，但没有服务器部署提交、后端文件或镜像摘要，不能认定两者完全一致。下文功能分支引用均指 `2b6077d` 的文件和行号，而非当前 main 文件。

线上仅发送少量 GET/HEAD 请求及无效令牌、异源请求头检查，没有提交权限修改、查询真实个人报表、连接生产统计数据库或进行压力测试。独立复现使用本地临时 HTTP 服务、合成身份和模拟 Graph 返回值。未修改业务代码、配置或分支。

## 1. 中危：邮箱作为持续授权标识，权限可能随邮箱转移

**位置：** main `server/admin.js:34–46`；功能分支 `server/admin.js:43–60`；`server/auth.js` 的 profile 返回逻辑还会在 mail 为空时回退到 userPrincipalName。

服务端从 Graph 获取资料并核对 oid/tid，这能阻止客户端随意声称自己是管理员，但最终邮箱条目仍仅按 `tenant_id + mail` 匹配，没有绑定最初获批的用户对象 ID。Graph 验证证明“当前账号拥有这些目录资料”，不能证明“当前账号就是最初被授权的人”。

**触发条件：** 使用邮箱类型管理员条目，并且同一租户中的邮箱/UPN 被重新分配、账号删除重建后复用地址，或具有相应目录修改能力的人把目标地址赋给另一个账号。没有证据表明普通用户可任意修改这个生产租户的 mail，也不能跨不同 tid 仅靠相同邮箱获得权限。

**影响：** 新 oid 可能继承旧邮箱的统计查询权；功能分支中若邮箱条目含 `can_manage_admins:true`，还会继承修改全部管理员的能力。邮箱缓存最长 5 分钟也可能短暂延续旧映射。

**本地验证：** 对功能分支真实 registerAdmin 代码注入两个不同的合成已验证 oid，令模拟 Graph 分别返回相同 tid/mail 且 id 各自匹配。两者访问 session 均返回 200 和 `canManageAdmins:true`。这验证了应用侧授权转移机制，不代表已在生产 Entra 完成账号接管。

**建议：** 将正式管理员统一绑定 `tenant_id + user_oid`；邮箱只用于查找和展示。添加时先解析并核验目标 oid，再保存不可变 ID，后续不以邮箱重新匹配授权。迁移现有邮箱条目时核对真实对象身份。

参考：[Microsoft 授权标识建议](https://learn.microsoft.com/en-us/entra/identity-platform/claims-validation)、[Graph user 属性定义](https://learn.microsoft.com/en-us/graph/api/resources/user?view=graph-rest-1.0)。使用不可变 tid/oid 的建议与这里的目录属性授权分析一致。

## 2. 中危（部署条件相关）：匿名流量可耗尽共享 IP 的管理员配额

**位置：** `server/app.js:41–49`；`server/start.js:36–38` 的反向代理部署分支。两个审查版本的限流逻辑相同。

后台 120 次/分钟的限流运行在认证之前，使用默认 IP 键，失败的匿名请求也计数；应用没有配置 trust proxy。共享出口下，普通访问者可以耗尽与管理员相同的配额。若生产使用未正确识别真实客户端地址的反向代理，则所有请求还可能合并成代理 IP 的一个配额。

**本地验证：** 对临时服务发送 120 次无 Authorization 请求，均为 401；随后同一 IP 的合法模拟管理员请求返回 429。`app.get('trust proxy')` 为 false。没有在线上重复这一耗尽测试。

**影响与边界：** 在同出口/代理归并条件下，可短暂阻断正常后台访问。如果生产是直接 HTTPS 且攻击者与管理员 IP 不同，则不能据此断言攻击者能耗尽管理员配额。当前未确认生产代理拓扑。

**建议：** 保留认证前的粗粒度防滥用限流，同时为认证后的后台访问增加按已验证 tid/oid 的独立配额。代理信任应按实际可信代理地址/链配置，不能简单设为 true 或盲信客户端传入的 X-Forwarded-For。验证不同真实客户端不会共享全部后台配额。

参考：[express-rate-limit 代理排查说明](https://express-rate-limit.mintlify.app/guides/troubleshooting-proxy-issues)。

## 3. 中危：管理员权限变更没有应用级审计轨迹

**位置：** 功能分支 `server/admin.js:69–82`、`server/analytics-store.js:112–115`。main 尚无这个写接口。

PUT `/api/admin/settings` 做了角色检查、输入校验、自身权限保护和 revision 并发控制，但保存时只覆盖 entries 并递增版本。没有把操作者 tid/oid、变更时间、变更前后内容或权限差异持久记录；传到存储层的参数也只有 revision 和 entries。

**影响：** 高权限账号被盗或误操作后，仅凭应用数据不能可靠回答谁给谁加过权限、何时撤销、是否短暂添加过另一个管理员。revision 防止覆盖冲突，不提供追溯能力。该缺口本身不绕过角色鉴权。外部代理/数据库审计是否存在未核验，普通访问日志通常也无法替代业务级变更差异。

**建议：** 权限写入与追加审计记录置于同一事务，记录已验证操作者、时间、版本和最小必要的变更差异；审计记录应限制修改权限。不要记录 Bearer token 或客户端密钥。

## 防御加固与需核验事项

### 4. 低危：后台缺少 CSP 和嵌入限制

位置：`server/app.js:23–25`、功能分支同处。线上 `/admin/usage` 的 GET 与 `/admin/settings` 的 HEAD 均未返回 Content-Security-Policy 或 X-Frame-Options，HTML 也无相关策略。

建议对独立后台页面及 `/admin.html` 静态别名设置适配 MSAL 的 CSP，包含 `frame-ancestors 'none'`，并增加 `X-Content-Type-Options: nosniff`。不要把禁止嵌入策略直接套到需要被 Office 嵌入的插件页面。当前表格和设置界面用 textContent/value 渲染，未发现可达 DOM XSS；没有以缺少 CSP 为由断言已存在 XSS。是否可完成已登录点击劫持也未验证，受 MSAL 登录和浏览器存储隔离影响。

### 5. 依赖有安全公告，但未确认后台存在可利用调用链

对当前 main 执行 `npm audit --omit=dev --registry=https://registry.npmjs.org --json`：0 critical、0 high、5 moderate **受影响包条目**，涉及 qs、uuid 及其父依赖，不是 5 个已证实的应用漏洞。默认 npmmirror 审计接口不支持，因此改用官方 registry，没有改 npm 配置或 lockfile。

- `qs@6.15.3`：数组限制公告依赖 comma 解析选项；当前 Express 默认未开启 comma。isBuffer 公告依赖把不可信对象交给 qs.stringify；后台筛选路径未见该调用。参见 [数组限制公告](https://github.com/advisories/GHSA-x5fp-wj9c-mxmx)、[isBuffer 公告](https://github.com/advisories/GHSA-4mjr-xmp4-gh2g)。
- `uuid`：公告影响特定 v3/v5/v6 缓冲区调用；当前 MSAL Node 检查到的是 uuid.v4 调用，未见后台可控数据流入受影响接口。参见 [uuid 公告](https://github.com/advisories/GHSA-w5hq-g745-h8pq)。

定位：`package-lock.json:7823`（qs）、`:9248`（uuid）、`:4203`（Express）。应规划兼容升级并复测登录和查询；不建议仅按 audit 建议直接强制跨大版本升级。生产安装版本未读取，不能用本机结果代替生产依赖清单。

### 其他边界

- 日汇总清理不会删除 users 的姓名、邮箱和首用时间：`server/analytics-store.js:99–103`。这是现有保留策略，不能把“日汇总保留 12 个月”理解为全部个人资料在 12 个月后删除。需单独确定人员资料生命周期；已有离线 deleteUser 方法。
- 功能分支明确将权限管理员设计为可管理所有管理员，且导入时默认把旧 `tenants:['*']` 账号变成权限管理员。这是文档明确的角色语义，未作为额外越权漏洞重复计数；部署时仍应核对名单。
- 不要直接用 main 的授权逻辑替换当前线上新增的数据库权限机制。main 从环境配置读取授权，回退时可能重新启用环境中旧账号；需先核对两套权限名单。
- 数据库 TLS、网络暴露、账号实际授权、备份、Entra MFA/条件访问、生产服务器版本和真实管理员角色配置均未读取，不对其安全性作结论。

## 已验证的保护

线上观察（约北京时间 14:25）：

| 请求 | 结果 |
| --- | --- |
| GET /admin/usage | 200，仅页面结构，不含统计数据 |
| GET /api/admin/overview，无令牌 | 401 |
| GET /api/admin/settings，无令牌 | 401；仅证明 API 认证入口拦截，不能单独证明后端存在该路由 |
| GET /api/admin/session，Bearer invalid | 401 |
| GET /api/admin/overview，异源 Origin | 403 |

代码核验：JWT 限 RS256，核验签名、issuer、audience、exp、tid 和 access_as_user；管理员再单独鉴权。租户权限来自服务端，SQL 通过参数绑定、租户与用户组合关联，并对 LIKE 通配符转义；TOP10 也复用租户过滤。表格、姓名、邮箱和设置字段不通过 innerHTML 渲染。统计记录不含正文、译文或访问令牌。API 返回 no-store。

后台使用显式 Bearer 头，异站表单不能自动带上 token，Origin 检查提供附加保护。不能仅因未配置传统 CSRF token 就判定该接口存在 CSRF。Origin 也不能替代 Bearer 验证。

本地验证命令：

```text
npm test -- --runInBand server/admin.test.js server/auth.test.js server/auth.multitenant.test.js server/analytics.test.js server/app.test.js src/admin/admin.test.ts
```

结果：6 个套件、38 项测试通过；另完成上述邮箱映射和共享 IP 配额的两项隔离复现。功能分支进行了代码检查及邮箱授权局部执行，没有运行其完整测试套件。未以真实普通用户、租户管理员和权限管理员登录生产环境，因此真实角色矩阵与保存流程仍需后续验收。
