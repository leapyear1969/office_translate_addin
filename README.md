# Office 翻译插件（Outlook / Word）

Outlook 邮件翻译与 Word 文档翻译共用本项目的认证、授权、语言检测和翻译后台。

## 分宿主构建与清单

| 命令 | 生成的清单 |
|---|---|
| `npm run build` | Outlook 和 Word 两份清单及完整前端 |
| `npm run build:outlook` | Outlook 清单及完整前端 |
| `npm run build:word` | Word 清单及完整前端 |
| `npm run manifest:generate:outlook` | 仅生成根目录 `manifest.outlook.xml` 和兼容文件 `manifest.xml` |
| `npm run manifest:generate:word` | 仅生成根目录 `manifest.word.xml` |
| `npm run manifest:generate` | 生成两端清单及 Outlook 兼容文件 |
| `npm run validate:outlook` / `npm run validate:word` | 校验对应根目录清单 |
| `npm run validate` | 校验两端清单 |

构建会将本次选中的清单复制到 `dist/manifests/`。`dist` 每次构建都会清理；同时部署两端时使用 `npm run build`。单宿主构建只筛选发布清单，仍打包完整前端，保证既有页面地址可用。根目录另一宿主的清单不会被单宿主生成命令覆盖，可能保留旧环境地址；部署请使用本次 `dist/manifests/` 中的清单。不要手改生成文件，应修改 `scripts/manifests/outlook.js` 或 `scripts/manifests/word.js` 后重新生成。`prod/manifest.xml` 和备份清单不参与构建，也不会自动更新。

Outlook 继续使用 `/taskpane.html`、`/commands.html` 和原加载项 ID；根目录 `manifest.xml` 始终是 Outlook 的兼容副本。Word 使用 `/word/taskpane.html` 和独立加载项 ID，可与 Outlook 同时安装。`npm run dev-server` 同时提供两端页面与 API。

## Word 文档翻译

- 使用 `manifest.word.xml` 安装，功能区保留“翻译选项”；选中文字后右键提供“翻译成中文”和“翻译设置”。右键翻译固定使用简体中文，并打开面板显示进度或登录错误。
- 面板提供“翻译选中文字”“翻译当前段落”“翻译正文全文”，使用已保存的目标语言。当前段落是光标所在段落；跨段落选区取第一个段落。正文全文不包括页眉、页脚、脚注等独立内容。
- 保留账户、重新登录、浏览器授权、始终翻译/翻译前询问/从不自动翻译、目标语言和排除语言管理。偏好通过 Word 文档 settings 保存，随当前文档持久化，与 Outlook 的 roamingSettings 独立。
- 打开面板、重新登录或保存设置后检查文档语言；关闭面板不启动后台自动翻译。手动翻译忽略自动翻译模式和排除语言。手动操作后暂缓自动检测，重新保存设置可恢复。自动翻译写回前如果面板已关闭则取消。
- 译文直接替换所选范围，可使用 Word 撤销。首次替换前将整份正文的 OOXML 原文备份写入当前文档 settings，备份保存失败则不替换；后续翻译和恢复都不会覆盖或删除这份备份。保存文档后，关闭并重新打开仍可通过“恢复原文”恢复，不需要登录或调用翻译服务。备份包含原文并随文件共享，可能增加文件体积。
- “恢复原文”需在面板确认，会替换整份正文，撤回备份之后的所有正文翻译及编辑，即使之前只翻译了选区或段落。页眉、页脚等独立内容不在备份与恢复范围内。这是一份固定原文备份，并非逐次翻译历史。恢复后当前面板会暂停自动翻译；重新打开面板或保存翻译设置后，按自动翻译设置重新判断。
- 范围在登录和请求前捕获并跟踪，写回前重新读取并比较选区文字和 HTML；若文字或 HTML 中体现的格式已修改则取消，避免覆盖编辑，也避免将 OOXML 包元数据变化误判为原文修改。翻译使用 HTML，恢复使用备份的 OOXML；不能保证复杂 Word 格式、域、内容控件或修订的完整保真，正式使用前用代表性文档验收。
- 两端使用相同 `APP_BASE_URL`、`CLIENT_ID`、`SSO_RESOURCE` 和 `/api/me`、`/api/consent/start`、`/api/detect`、`/api/translate`。接口内容上限 1,000,000 字符，不支持 DOCX 上传。后台保持同源要求。
- Word 清单要求 WordApi 1.3 和 SharedRuntime 1.1，让右键命令与面板共享运行状态，避免并行覆盖和重复命令失效。目标 Word 客户端还需支持 Office SSO；需在真实 Microsoft 365 Word 中验收客户端预授权、授权回调、右键菜单、撤销、表格和复杂格式。

清单结构参考：[Microsoft 的 Word 清单示例](https://github.com/OfficeDev/generator-office/blob/master/src/app/templates/hosts/word/manifest.xml)、[WebApplicationInfo](https://learn.microsoft.com/en-us/javascript/api/manifest/webapplicationinfo)。

阅读邮件时点击“翻译”菜单中的“翻译邮件”，将整封正文翻译为设置的目标语言并直接显示在正文位置。服务器原始邮件不被修改；重新打开邮件即可查看原文。另一个菜单项“翻译选项”打开 taskpane.html。

## 本地运行

需要 Node.js 20.16 或更高版本。

```powershell
npm install
npm run build
npm start
```

开发时使用 `npm run dev-server`，同时提供 Webpack 页面和后端 API。不要同时启动两个服务，它们共用 3000 端口。服务地址为 https://localhost:3000，使用 Office 开发证书。浏览器只可预览设置布局，登录、设置保存和翻译需在对应的 Outlook 或 Word 客户端中测试。

项目已创建 `.env` 并写入用户提供的翻译 Key。`.env` 和用户原有的 `.local` 均被 Git 忽略。新环境请从 `.env.example` 复制后填写，不要覆盖已有密钥。

## Entra 应用配置（世纪互联）

### SSO 错误 13004

`13004` 表示加载项清单中的资源地址无效。检查 `APP_BASE_URL` 与 `SSO_RESOURCE` 的域名和端口是否一致，同时确保 `SSO_RESOURCE` 等于 Entra 应用注册的 Application ID URI。例如，生产页面 `https://www.majun.fun:30260/word/taskpane.html` 对应 `api://www.majun.fun:30260/59a13c0d-6f19-4c66-b8fd-2fa80d0186bc`；不能搭配 `https://localhost:3000` 的页面。本地开发则需要在 Entra 中配置匹配的本地 URI，不能只修改清单中的 Resource。

生成清单时会提前拒绝域名或端口不匹配的配置。修正 `.env` 后运行 `npm run build:word`，部署更新的前端并在 Word 中重新旁加载 `dist/manifests/manifest.word.xml`；若仍使用旧配置，移除旧测试加载项并重新加载。反复登录或执行 Graph 浏览器授权不能修复清单资源错误。

参考：[微软 13004 排查说明](https://learn.microsoft.com/zh-cn/office/dev/add-ins/develop/troubleshoot-sso-in-office-add-ins#13004)。

- Tenant ID：`d7125684-0e28-40b5-aba2-ea9580f2a201`
- Client ID：`59a13c0d-6f19-4c66-b8fd-2fa80d0186bc`
- Web Redirect URI：`https://localhost:3000`
- Application ID URI：`api://localhost:3000/59a13c0d-6f19-4c66-b8fd-2fa80d0186bc`
- 公开委托权限 `access_as_user`，按微软 Office SSO 文档预授权 Office 客户端。
- Microsoft Graph 委托权限 `User.Read`，完成所需的管理员同意。
- 在 `.env` 的 `CLIENT_SECRET` 填入该新应用客户端密钥的**值**，不是密钥 ID，然后重启服务。

Redirect URI 和 Application ID URI 不同。当前实现使用 Office SSO → 后端 JWT 签名/发行者/受众/租户/权限验证 → MSAL OBO → 中国区 Graph `/me`。首次缺少 Graph consent 时通过外部浏览器及服务端授权码 + PKCE 流程让用户授权，用户重新打开插件后验证 SSO 和 Graph 访问。前端每次手动翻译会先获取并验证登录用户信息。SSO 不可用时会明确报错，不冒用邮箱地址作为认证结果。

`AUTHORITY` 默认 `https://login.partner.microsoftonline.cn`，`GRAPH_BASE` 默认 `https://microsoftgraph.chinacloudapi.cn`，与参考项目一致。

参考：[Office SSO 配置](https://learn.microsoft.com/en-us/office/dev/add-ins/develop/use-sso-to-get-office-signed-in-user-token)。

## 中国版跨租户登录与授权

应用注册应支持中国云中的任意组织目录。服务器按登录令牌的 `tid` 在配置的中国云身份地址发现租户元数据，验证签名、发行者、受众、租户及 `access_as_user` 权限后，使用同一租户执行 Graph OBO。`TENANT_ID` 保留为应用注册所属租户的信息，不再是登录租户白名单，也无需改为 `common`。当前支持同一云内所有能取得本应用有效令牌的组织租户。

客户用户需要有效的 consent；开发者租户的管理员同意不覆盖客户租户。项目使用 Microsoft Graph `User.Read` 委托权限读取用户信息，该权限本身不强制管理员同意：客户租户策略允许时，普通用户可以为自己同意；管理员也可代表整个租户同意。部署 Outlook 清单不能作为 Graph 授权已完成的证明。Office 客户端对 `access_as_user` 的预授权仍须在应用注册中配置。

首次缺少 Graph consent 时，邮件提示栏显示“登录并授权”操作，点击打开“翻译选项”；配置页中的“登录并授权”按钮使用 openBrowserWindow 打开默认浏览器中的微软中国版授权页面；不支持该接口时显示可点击或复制的链接，网页版打开当前浏览器的新标签页，无法强制切换系统默认浏览器。用户应使用当前 Outlook 账户并同意 `User.Read`。完成页提示用户关闭窗口并重新打开插件；插件重新打开后验证登录并按现有设置检查邮件。静默登录不会自动弹窗，取消后可重新点击。旧版 Outlook 功能命令没有通用的直接打开侧栏接口，因此需要用户点击提示栏操作；不支持操作提示的客户端会显示手动打开“翻译选项”的说明。

部署前，在中国版 Entra 应用注册 → 身份验证 → Web 中添加 `${APP_BASE_URL}/auth/consent/callback`（将占位符替换为实际值）。当前生产清单对应 `https://www.majun.fun:30260/auth/consent/callback`，本地开发是 `https://localhost:3000/auth/consent/callback`。这是 **Web** 回调，不是 SPA；不要启用隐式授权。保留现有 Application ID URI、Office 预授权和服务器 CLIENT_SECRET。

授权会话保存在服务端内存中，10 分钟过期且只能使用一次，使用 state、nonce、PKCE 并检查授权用户的 tid/oid 与已验证的 Outlook SSO 身份相同。浏览器不向插件传递令牌或完成消息；插件不会将打开浏览器视为授权成功。多实例部署需要会话粘滞或共享会话存储；服务重启后用户需重新发起授权。反向代理应将 `/auth/consent/*` 转发至 Node 服务，避免记录授权回调查询参数。仅使用普通用户同意；租户策略禁止用户授权时，仍需管理员处理，插件不会绕过租户策略。

如选择统一授权，将下面地址中的 `CUSTOMER_TENANT_ID` 替换为客户租户 ID，让客户管理员登录并核对权限后同意：

```text
https://login.partner.microsoftonline.cn/CUSTOMER_TENANT_ID/adminconsent?client_id=59a13c0d-6f19-4c66-b8fd-2fa80d0186bc
```

管理员可在客户租户的“企业应用程序”中按上述应用 ID 查找应用，并在“权限”中确认租户范围的授权。管理员统一授权是可选方式，不是 `User.Read` 的强制要求。上述 `/auth/consent/callback` 用于插件发起的用户授权码流程，不用于独立管理员授权链接。

升级时部署新的 `server` 代码及重新构建的 `dist`，然后重启 Node 服务。仅重新部署清单不会更新后端。此次修复无需更换应用 ID、客户租户 ID 或客户端密钥。

错误码 `13013` 是 Office SSO 请求限流，不是缺少 consent 的直接证据。遇到此错误先停止重复点击，稍后重新打开加载项；仍失败时检查最初的 Office 错误和 Entra 登录日志。用户授权入口针对已取得 SSO 令牌但 Graph 返回 `consent_required` 的情况；它不能修复 Office 无法签发 SSO 令牌的问题。

参考：[中国版租户管理员授权](https://docs.azure.cn/zh-cn/entra/identity/enterprise-apps/grant-admin-consent?pivots=portal)、[Office SSO 错误码](https://learn.microsoft.com/en-us/office/dev/add-ins/develop/troubleshoot-sso-in-office-add-ins)。

## 翻译配置与行为

| .env 字段 | 用途 |
|---|---|
| TRANSLATOR_ENDPOINT | 默认 https://api.translator.azure.cn/ |
| TRANSLATOR_KEY | 翻译资源 Key，仅后端读取 |
| TRANSLATOR_REGION | 资源需要区域认证时，填写门户中准确的区域代码 |
| APP_BASE_URL | 插件 HTTPS 地址，用于清单生成与 API 来源检查 |
| SSO_RESOURCE | Entra 应用中实际的 Application ID URI |

翻译时保留 HTML 元素及属性，只提交可见文本节点；表格、链接、CID 图片和样式不会作为普通文字翻译。script、style、head、noscript、translate=no 和 notranslate 区域跳过。不同 HTML 文本节点分别翻译，因此跨行内标签的句子连贯性可能不如单段纯文本。附件、图片内文字和邮件主题不翻译。

每批最多 45,000 字符、100 段，每段最多 4,500 UTF-16 单元；大邮件分批完成后才应用译文。输入/输出上限为 1,000,000 字符。译文通过文本节点赋值，尖括号等内容会正确转义。失败或邮件切换时不应用未完成的译文。不把正文、令牌或密钥写入日志或 localStorage。

已使用 chinanorth3 区域通过真实 Azure 样例验证：语言检测识别为英语，英文成功译为简体中文，表格样式、链接和 CID 图片引用保留。此前的 401001 认证错误已随区域配置补齐解决。真实 Outlook SSO 和正文显示仍需在客户端验收。

参考：[Azure China Translator API](https://docs.azure.cn/en-us/ai-services/translator/text-translation/reference/v3/translate)、[服务限制](https://docs.azure.cn/en-us/ai-services/translator/service-limits)。

## 翻译选项

默认“翻译前询问我”，目标为简体中文。设置通过 Outlook roamingSettings 保存，含目标语言、翻译模式和不提示翻译的语言列表。

“始终翻译”和“翻译前询问我”只在选项面板打开时检查当前邮件，并在支持固定面板的客户端通过 ItemChanged 检查新邮件。需要跨邮件使用时请固定面板。面板关闭后不会在后台自动翻译。语言检测置信度低于 0.7、目标语言相同、命中排除列表、选择“从不自动翻译”时不会自动翻译。手动点击“翻译邮件”不受自动偏好限制。

`DisplayedBody.setAsync` 仍使用 Office.js Preview，支持范围以客户端实际能力为准；保留本项目已经测试可用的 beta CDN。

### OWA 正文显示未响应

2026-09-05 在世纪互联 OWA (`partner.outlook.cn`) 实测：SSO、账户查询、正文读取及翻译请求均成功，但 `display.body.setAsync` 未替换正文，也未返回回调。仅检查该函数是否存在，不能证明宿主已实现该预览接口。此结果不代表所有 OWA 环境均不支持。

翻译状态现在分别显示登录、读取正文、翻译、显示译文。正文读取和显示调用在 15 秒无回调时明确报错并解除等待，命令也会完成；显示超时不表示翻译服务失败。无法取消已经提交给 Office 的显示调用，因此超时后请重新打开邮件再重试，避免旧调用延迟完成。原始邮件不会被修改。目前该 OWA 环境的原位显示仍待解决，可先使用已验证可用的 Outlook 客户端。

接口说明：[DisplayedBody.setAsync（预览）](https://learn.microsoft.com/en-us/javascript/api/outlook/office.displayedbody?view=outlook-js-preview)。

后续最小复现发现，绕过插件的 15 秒超时后，Office.js 约 76 秒返回 `RequestTimeout`（5018），而非永久不回调。完整环境、对照实验、SDK 边界返回码和一手资料见 [OWA 原位显示诊断](docs/owa-display-diagnosis.md)。

## 安装新清单

重新旁加载项目根目录 `manifest.xml`（版本 1.1.0.0）。沿用原插件 ID，因此升级已有安装即可；若仍显示旧按钮，移除旧测试插件后加载新清单。新清单包含 v1.1 WebApplicationInfo 和两个菜单项。

## 验证

```powershell
npm test -- --runInBand
npm run typecheck
npm run build
npm run validate
```

自动测试使用伪造 Office 环境、签名测试令牌和模拟翻译服务，不代表实际租户 SSO 已通过。真实验收见 `TEST_CHECKLIST.md`。

## 移到服务器

1. 在服务器部署本项目，安全配置 `.env`，安装依赖并构建。
2. 修改 `APP_BASE_URL` 和 Entra `SSO_RESOURCE`，同步更新应用的 Application ID URI、Office 预授权及需要的重定向地址。
3. 执行 `npm run build` 重新生成清单，重新安装更新后的 `manifest.xml`。
4. 推荐设置 `NODE_ENV=production`，由反向代理终止 HTTPS 并转发到 `127.0.0.1:PORT`。也可配置 `SSL_CERT_PATH` 和 `SSL_KEY_PATH` 直接提供 HTTPS。
5. 重启 Node 服务。不要把 `.env`、源目录或密钥当静态资源发布；本服务器仅发布 `dist`。

后端限流当前为每 IP 每分钟 60 次。生产环境多个实例或反向代理需要按实际网络配置共享限流存储与代理信任，当前本地默认不信任转发头。
