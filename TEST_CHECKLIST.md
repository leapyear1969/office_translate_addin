# Outlook DisplayedBody.setAsync 测试清单

请对每一种目标客户端分别复制并填写一份本清单。

## A. 测试环境记录

- [ ] Outlook 类型：经典 Outlook / New Outlook / OWA
- [ ] Outlook/OWA 版本：`____________________________`
- [ ] Windows/macOS 和版本：`____________________________`
- [ ] 邮箱属于 Microsoft 365 世纪互联租户
- [ ] 测试账户：`____________________________`
- [ ] 测试时间：`____________________________`
- [ ] Office.js URL：`https://appsforoffice.microsoft.com/lib/beta/hosted/office.js`
- [ ] 网络/代理/VPN 说明：`____________________________`

## B. 本地项目准备

- [ ] 在项目根目录运行 `npm install`
- [ ] 运行 `npx office-addin-dev-certs install`
- [ ] 已接受并信任 localhost 开发证书
- [ ] 运行 `npm test -- --runInBand`，所有测试通过
- [ ] 运行 `npm run typecheck`，无 TypeScript 错误
- [ ] 运行 `npm run build`，webpack 构建成功
- [ ] 运行 `npm run validate`，最后显示 `The manifest is valid.`
- [ ] 运行 `npm run dev-server` 并保持终端窗口开启
- [ ] 浏览器可打开 `https://localhost:3000/commands.html`
- [ ] 页面没有 HTTPS 证书错误（空白页面是预期行为）
- [ ] 浏览器/客户端网络可加载 Global Beta Office.js URL

## C. Manifest 安装

- [ ] 打开目标世纪互联租户对应的 Outlook 加载项管理界面
- [ ] 进入“我的加载项 / My add-ins”
- [ ] 选择“自定义加载项 → 从文件添加 / Add from File”
- [ ] 选择项目根目录的 `manifest.xml`
- [ ] 接受安装提示
- [ ] 若用户自助安装被禁用，已请管理员集中部署 manifest
- [ ] 安装后已重启或刷新 Outlook
- [ ] 打开一封已收到的邮件后能找到“测试原文翻译”按钮
- [ ] 能找到“显示原文”按钮

按钮位置可能在 Ribbon、邮件操作栏或 Apps/应用溢出菜单中。

## D. 替换显示正文

- [ ] 打开一封已经收到的 HTML 邮件
- [ ] 确认处于 Message Read / 阅读模式
- [ ] 确认没有同时选择多封邮件
- [ ] 记下或截图原始正文
- [ ] 打开该命令运行时对应的开发者工具/console
- [ ] 点击“测试原文翻译”
- [ ] console 显示邮件正文读取成功
- [ ] `Office.context.mailbox.item`：`____________________________`
- [ ] `item.display`：`____________________________`
- [ ] `item.display?.body`：`____________________________`
- [ ] `typeof item.display?.body?.setAsync`：`____________________________`
- [ ] 正文原显示区域出现“翻译测试成功”
- [ ] 没有打开右侧任务窗格
- [ ] console 显示 `DisplayedBody.setAsync succeeded`

## E. 恢复原文

- [ ] 不切换邮件，点击“显示原文”
- [ ] 原始 HTML 邮件正文恢复
- [ ] console 显示原文恢复成功
- [ ] 恢复内容属于当前同一封邮件
- [ ] 邮件未被标记为已编辑或产生服务器端正文变化

## F. 临时显示行为确认

- [ ] 再次点击“测试原文翻译”后切换到另一封邮件
- [ ] 切回原邮件时 Outlook 显示服务器中的原始正文
- [ ] 在另一个 Outlook 客户端查看同一邮件，正文未被修改

## G. 失败信息记录

如果任何步骤失败，请完整填写：

```text
Outlook 类型：
Outlook 版本：
操作系统：
世纪互联租户：是 / 否
Office.js URL：
Global Beta CDN 是否成功加载：

Office.context.host：
Office.context.platform：
diagnostics.hostName：
diagnostics.hostVersion：
diagnostics.OWAView：
item.itemType：
item.display：
item.display.body：
typeof item.display.body.setAsync：

失败阶段：读取正文 / API 检测 / 替换 / 恢复 / manifest 安装 / CDN 加载
错误代码：
错误名称：
错误信息：
完整 console 日志：
```

## H. 最终判定

只选择一项：

- [ ] **成功**：`setAsync` 存在，回调成功，正文原区域被替换，且可以恢复。
- [ ] **API 不支持**：Office.js 已成功初始化，但 `setAsync` 为 `undefined`。
- [ ] **CDN/初始化失败**：Global Beta Office.js 没有成功加载，尚不能判断 API 支持。
- [ ] **API 存在但调用失败**：已记录错误代码、名称和信息。
- [ ] **Manifest/客户端问题**：按钮没有出现或 Commands Runtime 没有启动。

补充说明：

```text



```
