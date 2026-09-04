# Outlook 邮件翻译插件

阅读邮件时点击“翻译”菜单中的“翻译邮件”，将整封正文翻译为设置的目标语言并直接显示在正文位置。服务器原始邮件不被修改；重新打开邮件即可查看原文。另一个菜单项“翻译选项”打开 taskpane.html。

## 本地运行

需要 Node.js 20.16 或更高版本。

```powershell
npm install
npm run build
npm start
```

开发时使用 `npm run dev-server`，同时提供 Webpack 页面和后端 API。不要同时启动两个服务，它们共用 3000 端口。服务地址为 https://localhost:3000，使用 Office 开发证书。浏览器只可预览设置布局，登录、设置保存和翻译需在 Outlook 中测试。

项目已创建 `.env` 并写入用户提供的翻译 Key。`.env` 和用户原有的 `.local` 均被 Git 忽略。新环境请从 `.env.example` 复制后填写，不要覆盖已有密钥。

## Entra 应用配置（世纪互联）

- Tenant ID：`d7125684-0e28-40b5-aba2-ea9580f2a201`
- Client ID：`59a13c0d-6f19-4c66-b8fd-2fa80d0186bc`
- Web Redirect URI：`https://localhost:3000`
- Application ID URI：`api://localhost:3000/59a13c0d-6f19-4c66-b8fd-2fa80d0186bc`
- 公开委托权限 `access_as_user`，按微软 Office SSO 文档预授权 Office 客户端。
- Microsoft Graph 委托权限 `User.Read`，完成所需的管理员同意。
- 在 `.env` 的 `CLIENT_SECRET` 填入该新应用客户端密钥的**值**，不是密钥 ID，然后重启服务。

Redirect URI 和 Application ID URI 不同。当前实现使用 Office SSO → 后端 JWT 签名/发行者/受众/租户/权限验证 → MSAL OBO → 中国区 Graph `/me`，不使用浏览器授权码回调，也没有跳过认证的开发接口。前端每次手动翻译会先获取并验证登录用户信息。SSO 不可用时会明确报错，不冒用邮箱地址作为认证结果。

`AUTHORITY` 默认 `https://login.partner.microsoftonline.cn`，`GRAPH_BASE` 默认 `https://microsoftgraph.chinacloudapi.cn`，与参考项目一致。

参考：[Office SSO 配置](https://learn.microsoft.com/en-us/office/dev/add-ins/develop/use-sso-to-get-office-signed-in-user-token)。

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
