# Outlook DisplayedBody.setAsync 验证 Demo

这是一个最小 Outlook Office.js Add-in，用来验证 Microsoft 365 世纪互联邮箱在 **Message Read / 阅读邮件** 场景中是否提供 Preview API：

```text
Office.context.mailbox.item.display.body.setAsync()
```

它只有两个 Ribbon 按钮：

- **测试原文翻译**：读取当前邮件 HTML，并把正文原显示区域临时替换为固定测试内容。
- **显示原文**：用此前缓存的 HTML 恢复当前邮件的显示正文。

这个项目不会修改服务器邮件，不使用 Microsoft Graph、EWS、VSTO、COM、任务窗格或任何翻译/API 服务。`DisplayedBody.setAsync` 设置的内容只在当前显示中生效；切换邮件或关闭当前邮件后，Outlook 仍显示服务器中的原始正文。

## 重要结论与限制

`DisplayedBody.setAsync` 截至 2026-09-03 仍是 **Mailbox Preview** API。微软官方文档要求使用 Office.js Preview 库，并明确说明 Preview requirement set 不应写入 manifest，因为客户端不会可靠报告该集合。

本项目使用：

```text
https://appsforoffice.microsoft.com/lib/beta/hosted/office.js
```

微软公开记录的三个相关地址是：

| 用途 | URL |
|---|---|
| Global 正式版 | `https://appsforoffice.microsoft.com/lib/1/hosted/office.js` |
| Global Preview/Beta | `https://appsforoffice.microsoft.com/lib/beta/hosted/office.js` |
| 世纪互联正式版 | `https://appsforoffice.cdn.partner.office365.cn/appsforoffice/lib/1/hosted/office.js` |

微软没有公开记录世纪互联专用的 Preview/Beta 地址。因此，本次测试依赖以下两个条件：

1. 世纪互联 Outlook 客户端实际实现 `DisplayedBody.setAsync`。
2. 客户端所在网络能够加载 Global Beta CDN。

如果 Office.js 加载失败，不能据此判定 API 不受支持；请先用浏览器或开发者工具确认 Global Beta URL 能成功访问。若 Office.js 已初始化，但运行时日志显示 `typeof item.display?.body?.setAsync === "undefined"`，才能判定当前客户端环境没有暴露该 API。

## 环境要求

- Node.js 18 或更高版本（已在 Node.js 20.16.0 验证）。
- npm。
- 可登录目标世纪互联租户的 Outlook 客户端或 Outlook on the web。
- 本机可以信任开发 HTTPS 证书。
- 网络策略允许访问 `https://localhost:3000` 和 Global Office.js Beta CDN。

## 安装依赖

在项目根目录运行：

```powershell
npm install
```

安装并信任 Microsoft Office Add-in 本地开发证书：

```powershell
npx office-addin-dev-certs install
```

Windows 可能显示证书信任或管理员确认窗口，请按提示接受。证书只用于本机 `localhost` 开发。

## 开发运行

启动本地 HTTPS Web Server：

```powershell
npm run dev-server
```

服务地址：

```text
https://localhost:3000/commands.html
```

先在浏览器中打开这个地址，确认没有证书警告。页面本身是空白的，这是预期行为；它是无 UI 的 Commands FunctionFile，不是任务窗格。

其他命令：

```powershell
npm test -- --runInBand
npm run typecheck
npm run build
npm run validate
```

`npm run validate` 固定使用微软 `office-addin-manifest@1.13.6`。测试时发现 `office-addin-manifest@2.1.6` 在 Node.js 20.16.0 下存在 CommonJS/ESM 传递依赖冲突；这不影响 Office.js 或 add-in 运行。

## Manifest URL 对应关系

[manifest.xml](./manifest.xml) 中的本地地址均为 `https://localhost:3000`：

- 基础 manifest 的 `SourceLocation`：`https://localhost:3000/commands.html`
- `FunctionFile resid="Commands.Url"`：引用资源 `Commands.Url`
- `Commands.Url`：`https://localhost:3000/commands.html`
- 图标：`https://localhost:3000/assets/icon-16.png`、`icon-32.png`、`icon-80.png`

如果端口或主机名改变，必须同时修改 `manifest.xml` 和 `webpack.config.js`，然后重新安装 manifest。

## 切换 Office.js URL

Office.js 地址位于：

```text
src/commands/commands.html
```

本测试默认使用 Global Beta。若只想对比正式版，可以把 `<script src>` 临时改成世纪互联正式地址后重启 dev server。正式版很可能没有 `DisplayedBody`，这正是对照结果，不代表加载项基础能力异常。

不要猜测或使用未经微软记录的世纪互联 `/beta/` 地址。

## 安装 XML Manifest

本项目使用 add-in-only XML manifest。微软当前通用手工安装流程是打开 **Add-Ins for Outlook / Outlook 加载项** 对话框，进入 **My add-ins / 我的加载项**，在 **Custom Addins / 自定义加载项** 中选择 **Add from File / 从文件添加**，然后选择本项目的 `manifest.xml`。

微软的快捷入口是：

```text
https://aka.ms/olksideload
```

如果该 Global 快捷入口不能正确进入世纪互联租户，请从你的世纪互联 Outlook/OWA 内部打开加载项管理界面，或让租户管理员集中部署这份 XML。世纪互联等主权云可能禁用用户自助安装，且公共 Microsoft 365 商店不可用；这时必须由管理员通过租户支持的集中部署入口上传 manifest，并确认本地 Web 服务和 Global Beta CDN 已被网络策略允许。

### 经典 Outlook for Windows

1. 保持 `npm run dev-server` 运行。
2. 在 Outlook 中选择 **文件 → 信息 → 管理加载项**；该入口会在浏览器打开当前邮箱对应的加载项对话框。
3. 进入 **我的加载项 → 自定义加载项 → 从文件添加**。
4. 选择 `manifest.xml` 并接受安装提示。
5. 回到 Outlook，打开一封已收到的邮件。

微软说明：经典 Outlook 手工 sideload 后可能因缓存最多延迟 24 小时显示。通常重启 Outlook、重新打开邮件或清理 Office 加载项缓存会更快刷新。

### New Outlook for Windows

1. 保持 `npm run dev-server` 运行。
2. 使用 `https://aka.ms/olksideload`，或从新 Outlook 的 **Apps/应用** 区域进入加载项管理。
3. 在 **我的加载项 → 自定义加载项 → 从文件添加** 中上传 `manifest.xml`。
4. 打开一封收到的邮件，在邮件操作栏、Ribbon 或 **Apps/应用** 溢出菜单中查找两个命令。

### Outlook on the web

1. 登录目标世纪互联 OWA，而不是 Global Microsoft 365 账户。
2. 从 OWA 的 **Apps/应用** 或加载项管理入口打开 **我的加载项**；若租户允许，也可以尝试 `https://aka.ms/olksideload`。
3. 选择 **自定义加载项 → 从文件添加** 并上传 `manifest.xml`。
4. 打开一封收到的邮件，在 Ribbon、邮件操作栏或 **Apps/应用** 菜单中查找命令。

安装一次后，同一邮箱的受支持 Outlook 客户端通常会同步显示该加载项，但世纪互联租户策略和客户端缓存可能导致差异。应分别记录每个客户端的实际结果。

微软参考：

- [Sideload Outlook add-ins for testing](https://learn.microsoft.com/en-us/office/dev/add-ins/outlook/sideload-outlook-add-ins-for-testing)
- [Guidance for deploying Office Add-ins on sovereign clouds](https://learn.microsoft.com/en-us/office/dev/add-ins/publish/government-cloud-guidance)
- [Office.DisplayedBody API](https://learn.microsoft.com/en-us/javascript/api/outlook/office.displayedbody?view=outlook-js-preview)
- [Outlook Preview requirement set](https://learn.microsoft.com/en-us/javascript/api/requirement-sets/outlook/outlook-requirement-set-preview?view=common-js-preview)

## 测试步骤

1. 运行 `npm run dev-server`，保持本地 HTTPS 服务运行。
2. 安装 `manifest.xml`。
3. 打开或重启 Outlook。
4. 打开一封已经收到的 HTML 邮件，确保处于阅读模式且只选择一封邮件。
5. 点击 **测试原文翻译**。
6. 检查正文原显示区域是否直接变成“翻译测试成功”。
7. 点击 **显示原文**。
8. 检查原邮件正文是否恢复。
9. 使用 Outlook/浏览器开发者工具保存 console 输出。

详细步骤见 [TEST_CHECKLIST.md](./TEST_CHECKLIST.md)。

## 如何判断结果

### 成功

以下三项必须同时满足：

```text
item.display.body.setAsync 存在
setAsync 回调返回 Succeeded
邮件正文原显示区域被替换
```

恢复按钮也应把同一封邮件恢复为测试前读取的 HTML。

### API 不支持

Outlook 中显示：

```text
当前 Outlook / Office.js 环境未提供 DisplayedBody.setAsync。
```

console 中的 `typeof item.display?.body?.setAsync` 为 `undefined`。

### 调用失败

console 会输出：

```text
error.code
error.name
error.message
```

请记录这些字段，不要只记录“按钮没反应”。

## 原文缓存机制

代码仍使用用户要求的模块变量：

```typescript
let originalHtml: string | null = null;
```

但 Outlook 不支持 function commands 的 shared runtime。每次命令在调用 `event.completed()` 后，浏览器运行时可以立即关闭，第二个按钮不能可靠读取前一次调用的模块内存。因此项目还把原 HTML 按 `itemId` 写入同源 `localStorage`，仅作为跨命令运行时的客户端缓存：

- 不写回邮件；
- 不发送到服务器；
- 不使用 Graph/EWS；
- 内存与本地缓存都会核对 `itemId`，不允许把一封邮件的缓存恢复到另一封邮件；
- 缓存最多保留 1 小时，成功恢复后立即删除。

这是为了让“显示原文”在真实 Outlook 命令生命周期中可运行，而不是改变核心测试目标。如果这仍不符合后续生产插件的数据策略，应在正式产品中重新设计，但本 Demo 不会扩展到该范围。

## 调试提示

每次命令都会输出：

```text
Office.context.host
Office.context.platform
Office.context.mailbox.diagnostics.hostName
Office.context.mailbox.diagnostics.hostVersion
Office.context.mailbox.diagnostics.OWAView
Office.context.mailbox.item.itemType
Office.context.mailbox.item.display
Office.context.mailbox.item.display?.body
typeof item.display?.body?.setAsync
```

若 Ribbon 按钮完全不出现，优先检查 manifest 安装、Message Read 场景、客户端缓存和租户策略。若按钮出现但没有日志，检查 `commands.html`、`commands.js` 以及 Global Beta CDN 的网络加载状态。
