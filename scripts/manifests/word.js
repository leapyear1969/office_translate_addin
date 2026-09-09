module.exports = function wordManifest(config) {
  const xml = value => String(value).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/"/g, '&quot;');
  const origin = xml(config.origin);
  return `<?xml version="1.0" encoding="UTF-8"?>
<OfficeApp xmlns="http://schemas.microsoft.com/office/appforoffice/1.1" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance" xmlns:bt="http://schemas.microsoft.com/office/officeappbasictypes/1.0" xsi:type="TaskPaneApp">
  <Id>2c39aa04-7c4b-4e96-90f6-8d6d9d495727</Id><Version>1.0.0.0</Version><ProviderName>文档翻译</ProviderName><DefaultLocale>zh-CN</DefaultLocale>
  <DisplayName DefaultValue="Word 文档翻译"/><Description DefaultValue="使用共享翻译服务翻译选中文字、当前段落或文档正文。"/>
  <IconUrl DefaultValue="${origin}/assets/icon-32.png"/><HighResolutionIconUrl DefaultValue="${origin}/assets/icon-80.png"/><SupportUrl DefaultValue="${origin}/word/taskpane.html"/>
  <AppDomains><AppDomain>${origin}</AppDomain></AppDomains><Hosts><Host Name="Document"/></Hosts>
  <Requirements><Sets DefaultMinVersion="1.1"><Set Name="WordApi" MinVersion="1.3"/><Set Name="SharedRuntime" MinVersion="1.1"/></Sets></Requirements>
  <DefaultSettings><SourceLocation DefaultValue="${origin}/word/taskpane.html"/></DefaultSettings>
  <Permissions>ReadWriteDocument</Permissions>
  <VersionOverrides xmlns="http://schemas.microsoft.com/office/taskpaneappversionoverrides" xsi:type="VersionOverridesV1_0">
    <Requirements><bt:Sets DefaultMinVersion="1.1"><bt:Set Name="SharedRuntime" MinVersion="1.1"/></bt:Sets></Requirements>
    <Hosts><Host xsi:type="Document"><Runtimes><Runtime resid="Taskpane.Url" lifetime="long"/></Runtimes><DesktopFormFactor>
      <FunctionFile resid="Taskpane.Url"/>
      <ExtensionPoint xsi:type="PrimaryCommandSurface"><OfficeTab id="TabHome"><Group id="WordTranslation.Group">
        <Label resid="Group.Label"/>
        <Icon><bt:Image size="16" resid="Icon.16"/><bt:Image size="32" resid="Icon.32"/><bt:Image size="80" resid="Icon.80"/></Icon>
        <Control xsi:type="Button" id="WordTranslation.Open">
          <Label resid="Open.Label"/><Supertip><Title resid="Open.Label"/><Description resid="Open.Description"/></Supertip>
          <Icon><bt:Image size="16" resid="Icon.16"/><bt:Image size="32" resid="Icon.32"/><bt:Image size="80" resid="Icon.80"/></Icon>
          <Action xsi:type="ShowTaskpane"><TaskpaneId>WordTranslation.Pane</TaskpaneId><SourceLocation resid="Taskpane.Url"/></Action>
        </Control>
      </Group></OfficeTab></ExtensionPoint>
      <ExtensionPoint xsi:type="ContextMenu"><OfficeMenu id="ContextMenuText">
        <Control xsi:type="Button" id="WordTranslation.Settings">
          <Label resid="Settings.Label"/><Supertip><Title resid="Settings.Label"/><Description resid="Open.Description"/></Supertip>
          <Icon><bt:Image size="16" resid="Icon.16"/><bt:Image size="32" resid="Icon.32"/><bt:Image size="80" resid="Icon.80"/></Icon>
          <Action xsi:type="ShowTaskpane"><TaskpaneId>WordTranslation.Pane</TaskpaneId><SourceLocation resid="Taskpane.Url"/></Action>
        </Control>
      </OfficeMenu></ExtensionPoint>
    </DesktopFormFactor></Host></Hosts>
    <Resources>
      <bt:Images><bt:Image id="Icon.16" DefaultValue="${origin}/assets/icon-16.png"/><bt:Image id="Icon.32" DefaultValue="${origin}/assets/icon-32.png"/><bt:Image id="Icon.80" DefaultValue="${origin}/assets/icon-80.png"/></bt:Images>
      <bt:Urls><bt:Url id="Taskpane.Url" DefaultValue="${origin}/word/taskpane.html"/></bt:Urls>
      <bt:ShortStrings><bt:String id="Group.Label" DefaultValue="文档翻译"/><bt:String id="Open.Label" DefaultValue="翻译选项"/><bt:String id="Settings.Label" DefaultValue="翻译设置"/></bt:ShortStrings>
      <bt:LongStrings><bt:String id="Open.Description" DefaultValue="打开翻译选项，设置目标语言、自动翻译和排除语言。"/></bt:LongStrings>
    </Resources>
    <WebApplicationInfo><Id>${xml(config.clientId)}</Id><Resource>${xml(config.resource)}</Resource><Scopes><Scope>openid</Scope><Scope>profile</Scope><Scope>User.Read</Scope></Scopes></WebApplicationInfo>
  </VersionOverrides>
</OfficeApp>
`;
};
