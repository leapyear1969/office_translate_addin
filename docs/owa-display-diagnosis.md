# 世纪互联 OWA 原位显示接口诊断

日期：2026-09-05。保留原位替换方案，不改为侧栏展示译文。

## 当前结论

在当前账号的 `partner.outlook.cn` 阅读窗格中，普通正文读取成功，`display.body.setAsync` 的固定短文本和短 HTML 请求均未更新正文。延长 HTML 测试的等待后，Office.js 在 76,046 毫秒返回 `RequestTimeout`，错误码 `5018`；底层通信返回码 `-6`、结果 `null`。这纠正了此前“永远不回调”的描述：插件原有 15 秒超时早于 SDK 自身的超时。

请求已通过 Office.js 参数检查，并进入发往 OWA 的 `ExecuteMethod` 调用（DispatchId 206）。这将故障范围缩小到 OWA 对该预览调用的处理或其通信返回链路；尚不能区分功能未部署、功能开关未启用与宿主实现缺陷。

目前没有找到可以在插件侧恢复原位显示的受支持修改。不能仅凭函数存在或一次超时断言所有世纪互联租户永久不支持该接口。

## 实测环境与对照

- 宿主：`OutlookWebApp`；平台：`OfficeOnline`。
- OWA 版本：`20260807011.15`；视图：`ThreeColumns`。
- 插件地址：`https://localhost:3000/taskpane.html`，Webpack 开发服务器。
- Office.js：Global Beta CDN，宿主脚本 `outlook-web-16.01.js`。
- `body.getAsync(Html)`：成功，正文长度 53,298 字符。日志未保存正文。
- 固定纯文本 `OWA 显示接口测试 / Hello`：18 字符，发出 DispatchId 206、coercionType 0，无正文更新，10 秒无回调。
- 固定 HTML `<p>OWA 显示接口测试 / Hello</p>`：25 字符，发出 DispatchId 206、coercionType 3，无正文更新，10 秒无回调。
- 延长上述 HTML 测试：76,046 毫秒后返回 `status: failed`、`error.name: RequestTimeout`、`error.code: 5018`、`error.message: 发生了内部错误。`。未使用项目 `officeCall` 的超时封装。
- 测试直接调用显示接口，不经过 SSO、Graph 或翻译服务，因此这些步骤及邮件大小不是该最小复现的触发条件。

## 截图中的控制台消息

1. `Permissions policy violation: unload is not allowed in this document`：微软宿主脚本在生命周期统计中注册 `unload`，用于页面退出时调用 `OSFAppTelemetry.onAppClosed`。这条警告不能证明显示调用被阻止。
2. `An iframe which has both allow-scripts and allow-same-origin ... can escape its sandboxing`：浏览器对 iframe 权限组合的警告，不是该脚本执行已被阻止的记录。
3. `upgrade-insecure-requests is ignored when delivered in a report-only policy`：此指令在 report-only 策略中不执行 URL 升级；它自身不是请求被 CSP 拦截的证据。

没有为消除这些消息修改浏览器权限、CSP、sandbox 或 Office.js。

## SDK 与服务支持证据

读取微软公开的 `outlook-web-16.01.debug.js` 可见：`display.body.setAsync` 经 `setBody(206)` 执行 Beta 检查、读写权限检查及参数检查，然后调用 `standardInvokeHostMethod`。在 OWA 路径中，该请求使用 `ExecuteMethod` 发送到宿主。诊断临时包装此边界，原样转发请求和回调，只记录方法编号、类型及长度，未改变参数、权限或返回结果。

微软预览要求说明：OWA 可能需要 Targeted release，预览 requirement set 尚未完整实现，客户端不能准确报告其支持状态。世纪互联服务说明将开启 Targeted release 列为不可用。因此不能把全球版启用定向发布的操作当成当前租户的确定解法。这些文档支持“宿主预览能力差异”的判断，但没有明确给出该 API 在此版本的支持承诺。

## 复现代码

在已初始化、具有 ReadWriteItem 权限的 Outlook 阅读加载项内执行：

```js
const started = Date.now();
Office.context.mailbox.item.display.body.setAsync(
  '<p>OWA 显示接口测试 / Hello</p>',
  { coercionType: Office.CoercionType.Html },
  result => console.log({
    elapsedMs: Date.now() - started,
    status: result.status,
    errorCode: result.error?.code,
  })
);
```

如需继续推进原位替换，应向微软/世纪互联确认 OWA 版本 `20260807011.15` 对 DispatchId 206 的实现与功能发布情况，并提交上述最小复现。不要更改内部功能开关、调用未公开接口或关闭浏览器安全策略来强行启用。

## 一手来源

- [Outlook 预览要求集](https://learn.microsoft.com/en-us/javascript/api/requirement-sets/outlook/outlook-requirement-set-preview?view=word-js-preview)
- [DisplayedBody.setAsync](https://learn.microsoft.com/en-us/javascript/api/outlook/office.displayedbody?view=outlook-js-preview)
- [世纪互联服务说明：Service updates](https://learn.microsoft.com/en-us/office365/servicedescriptions/office-365-platform-service-description/microsoft-365-operated-by-21vianet)
- [微软公开的 Beta 调试脚本](https://appsforoffice.microsoft.com/lib/beta/hosted/outlook-web-16.01.debug.js)
- [Chrome 的 unload 弃用说明](https://developer.chrome.com/docs/web-platform/deprecating-unload)
- [W3C：upgrade-insecure-requests 的 report-only 行为](https://www.w3.org/TR/upgrade-insecure-requests/#delivery)
