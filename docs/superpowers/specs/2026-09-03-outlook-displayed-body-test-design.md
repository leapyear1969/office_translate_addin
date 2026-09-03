# Outlook DisplayedBody Preview API 验证 Demo 设计

## 目标

创建一个最小可运行的 Outlook Office.js Add-in，用于在 Microsoft 365 世纪互联邮箱环境中验证：用户阅读一封已收到的邮件时，Global Office.js Beta 库是否暴露 `Office.context.mailbox.item.display.body.setAsync()`，以及该方法能否临时替换邮件正文原显示区域。

本 Demo 只验证客户端显示层行为。它不修改服务器邮件，不调用 Microsoft Graph、EWS、VSTO、COM、翻译服务或生成式 AI，也不使用任务窗格。

## 已确认的技术选择

- 使用 XML Office Add-in manifest。
- 只声明 `MessageReadCommandSurface`，不支持 Compose。
- 提供两个 Ribbon 命令：“测试原文翻译”和“显示原文”。
- 使用 TypeScript、HTML、webpack、Node.js 和 npm，不引入 React。
- Commands 页面默认直接加载 Global Office.js Beta：
  `https://appsforoffice.microsoft.com/lib/beta/hosted/office.js`。
- TypeScript 使用 `@types/office-js-preview`。
- 不在 manifest 中声明 Mailbox Preview requirement set，因为微软明确说明客户端无法可靠报告该集合，且不应把它写入 manifest。
- Manifest 使用支持 Add-in Commands 的稳定 Mailbox requirement set；Preview API 是否可用完全通过运行时能力检测判断。

## CDN 与世纪互联边界

微软公开记录的正式版地址如下：

- Global 正式版：`https://appsforoffice.microsoft.com/lib/1/hosted/office.js`
- Global Preview：`https://appsforoffice.microsoft.com/lib/beta/hosted/office.js`
- 世纪互联正式版：`https://appsforoffice.cdn.partner.office365.cn/appsforoffice/lib/1/hosted/office.js`

截至设计日期，微软官方文档没有列出世纪互联专用 Preview CDN。由于 `DisplayedBody.setAsync()` 属于 Mailbox Preview，本 Demo 按用户确认默认使用 Global Preview 地址。测试结果因此同时受两个条件影响：世纪互联 Outlook 客户端是否实现该 Preview API，以及客户端网络是否可以加载 Global Preview CDN。

README 将给出手工切换 `<script src>` 的位置，但首版不引入构建变量或维护两份 Commands 页面。

## 项目结构

```text
office_tanslate_addin/
├─ manifest.xml
├─ package.json
├─ package-lock.json
├─ tsconfig.json
├─ webpack.config.js
├─ README.md
├─ TEST_CHECKLIST.md
└─ src/
   ├─ commands/
   │  ├─ commands.html
   │  └─ commands.ts
   └─ assets/
      ├─ icon-16.png
      ├─ icon-32.png
      └─ icon-80.png
```

图标只为满足 Ribbon manifest 的资源要求，保持简单，不承担产品视觉设计。

## Manifest 设计

Manifest 将包含：

- 唯一且有效的 Add-in ID。
- `ReadWriteItem` 权限，因为微软对 `DisplayedBody.setAsync()` 标注的最低权限是 read/write item。
- 仅对 `ItemIs`、`Message`、`Read` 场景激活。
- `VersionOverridesV1_0` 中的 `MessageReadCommandSurface`。
- 一个 `FunctionFile`，指向 `https://localhost:3000/commands.html`。
- 两个 `ExecuteFunction` 控件，分别绑定 `testTranslation` 和 `restoreOriginal`。
- 本地 HTTPS Commands URL、图标 URL与所需资源字符串。

不添加任务窗格入口，不声明 Compose 命令，也不声明 Preview requirement set。

## Commands Runtime 设计

`commands.html` 是无可见 UI 的 Functions/Commands 页面，负责加载 Global Beta Office.js 和 webpack 输出的命令脚本。

`commands.ts` 维护模块级状态：

```typescript
let originalHtml: string | null = null;
```

Outlook 不支持 function commands 的 shared runtime，且命令调用 `event.completed()` 后运行时会关闭。为使第二个 Ribbon 命令能够在新的运行时恢复正文，项目同时按当前邮件 `itemId` 将原始 HTML 缓存在同源 `localStorage`。该缓存仅保存在客户端，不修改邮件、不写服务器，也不发送到外部服务；模块内存也记录来源 `itemId`，只有当前邮件匹配时才使用。缓存超过 1 小时后不再用于恢复；每次命令启动时惰性删除该加载项前缀下的过期记录，并在成功恢复后立即删除当前记录。由于运行时关闭后没有定时进程，若用户不再运行任何命令，磁盘上的旧记录会保留到下一次命令运行或站点数据被清理。

它公开两个供 Outlook manifest 调用的函数：

- `testTranslation(event)`
- `restoreOriginal(event)`

函数通过 `Office.actions.associate` 与 manifest 中的函数名关联。代码不会仅依赖把函数挂到 `window` 上。

## 数据流

### 测试替换

1. 用户在 Message Read 中点击“测试原文翻译”。
2. 输出 `dumpEnvironmentInfo()`。
3. 调用 `item.body.getAsync(Office.CoercionType.Html, callback)`。
4. 检查读取结果；成功后将 HTML 保存到 `originalHtml`。
5. 逐层检查 `item`、`item.display`、`item.display.body` 和 `setAsync`。
6. 若支持，调用 `display.body.setAsync(testHtml, { coercionType: Html }, callback)`。
7. 输出成功或完整失败信息。
8. 在所有路径中最终调用一次 `event.completed()`。

严格按用户要求先读取正文，再进行 DisplayedBody 能力检测，以验证正文读取和保存链路。

### 恢复原文

1. 用户点击“显示原文”。
2. 输出环境信息并检查 `originalHtml`。
3. 若为空，显示“尚未保存原始正文，请先点击‘测试原文翻译’。”。
4. 再执行完整的 DisplayedBody 能力检测。
5. 调用 `display.body.setAsync(originalHtml, { coercionType: Html }, callback)`。
6. 输出成功或完整失败信息。
7. 在所有路径中最终调用一次 `event.completed()`。

原文保存在当前 Commands Runtime 的内存中，并按 `itemId` 在同源 `localStorage` 中保留跨命令运行时副本。内存与本地缓存都校验 `itemId`，不同邮件不会互相恢复；本地缓存超过 1 小时后失效，在下一次任一命令启动时惰性清理，并在成功恢复后删除。Outlook 自身会在离开当前邮件后恢复服务器中的原始正文，因为 Preview API 不会持久化修改。

## 完成回调约束

异步逻辑封装成 Promise，顶层命令处理器使用统一包装函数，并通过 `finally` 调用 `event.completed()`。这样能避免：

- 失败分支遗漏完成通知；
- 回调和异常路径重复调用；
- Outlook Ribbon 命令持续显示执行中。

## 能力检测与诊断

能力检测不会依赖 `Office.context.requirements.isSetSupported("Mailbox", ...)` 判断 Preview API，而是直接检查对象图。

不支持时将：

- 在 console 输出 `Office.context.mailbox.item`、`item.display`、`item.display?.body` 和 `typeof item.display?.body?.setAsync`。
- 向用户显示“当前 Outlook / Office.js 环境未提供 DisplayedBody.setAsync。”。

`dumpEnvironmentInfo()` 使用安全访问输出：

- `Office.context.host`
- `Office.context.platform`
- `mailbox.diagnostics.hostName`
- `mailbox.diagnostics.hostVersion`
- `mailbox.diagnostics.OWAView`
- `item.itemType`
- `item.display`
- `item.display?.body`

用户提示优先使用 `Office.context.mailbox.item.notificationMessages` 在当前邮件中显示可见通知；如果该 API 本身不可用或失败，则至少保留 console 日志。通知使用固定 key，以便后续状态覆盖旧状态。

## 错误处理

所有 Office.js `AsyncResult` 都检查 `Office.AsyncResultStatus.Succeeded` 和失败状态。失败日志统一包含：

- `error.code`
- `error.name`
- `error.message`

同步异常、Office 尚未初始化、当前 item 不存在、正文读取失败、Preview API 缺失、正文替换失败和恢复前无缓存均有独立清晰提示。日志不会包含除当前正文对象调试信息之外的外部传输；邮件正文不会发送到任何服务。

## 构建与本地 HTTPS

webpack 将 TypeScript 编译为单一 Commands bundle，并复制 Commands HTML 和图标到 `dist/`。开发服务器监听 `https://localhost:3000`。

使用 `office-addin-dev-certs` 生成并信任本地开发证书；webpack dev server 使用相同证书。npm 脚本至少包含：

- 安装后可直接执行的 build；
- `npm run dev-server` 启动 HTTPS 服务；
- manifest 验证命令。

不自动修改或安装 Outlook 配置，sideload 由用户依据 README 手工执行。

## 文档与人工验证

README 覆盖依赖安装、HTTPS 证书、开发服务器、URL 配置、Global/世纪互联 CDN 差异、XML manifest sideload、经典 Outlook、新 Outlook 和 OWA 的测试注意事项、完整测试步骤与结果记录模板。

`TEST_CHECKLIST.md` 提供逐项复选框，要求记录：客户端类型与版本、实际 Office.js URL、对象能力检测输出、调用结果、邮件正文是否原位替换、恢复结果和错误详情。

## 验证策略

本地自动验证包括：

- `npm install`
- TypeScript/webpack 生产构建
- XML manifest schema 验证
- 检查构建产物中 Commands HTML、脚本和图标是否齐全

真正的 API 支持结论只能通过用户在目标 Microsoft 365 世纪互联账户和具体 Outlook 客户端中人工运行得到。项目不会把“Global Beta 脚本成功加载”误报为“客户端支持 DisplayedBody”。

## 官方依据

- Office.DisplayedBody：<https://learn.microsoft.com/en-us/javascript/api/outlook/office.displayedbody?view=outlook-js-preview>
- Outlook Preview requirement set：<https://learn.microsoft.com/en-us/javascript/api/requirement-sets/outlook/outlook-requirement-set-preview?view=common-js-preview>
- Office.js 正式版与 Preview CDN：<https://learn.microsoft.com/en-us/office/dev/add-ins/develop/understand-the-javascript-api-for-office>
- 世纪互联 Office.js CDN 指引：<https://learn.microsoft.com/en-us/office/dev/add-ins/publish/government-cloud-guidance>
