const fs = require('fs');
const path = require('path');
const { readConfig } = require('../server/config');
const config = readConfig();
const xml = value => String(value).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/"/g, '&quot;');
const origin = xml(config.origin);
const icons = `<Icon><bt:Image size="16" resid="Icon.16"/><bt:Image size="32" resid="Icon.32"/><bt:Image size="64" resid="Icon.64"/><bt:Image size="80" resid="Icon.80"/></Icon>`;
function contents(v11) {
  return `<Requirements><bt:Sets DefaultMinVersion="1.5"><bt:Set Name="Mailbox"/></bt:Sets></Requirements>
    <Hosts><Host xsi:type="MailHost"><DesktopFormFactor>
      <FunctionFile resid="Commands.Url"/>
      <ExtensionPoint xsi:type="MessageReadCommandSurface"><OfficeTab id="TabDefault"><Group id="Translation.Group">
        <Label resid="Group.Label"/>
        <Control xsi:type="Menu" id="Translation.Menu">
          <Label resid="Menu.Label"/><Supertip><Title resid="Menu.Label"/><Description resid="Menu.Description"/></Supertip>${icons}
          <Items>
            <Item id="Translation.Message"><Label resid="Translate.Label"/>
              <Supertip><Title resid="Translate.Label"/><Description resid="Translate.Description"/></Supertip>${icons}
              <Action xsi:type="ExecuteFunction"><FunctionName>translateMessage</FunctionName></Action>
            </Item>
            <Item id="Translation.Options"><Label resid="Options.Label"/>
              <Supertip><Title resid="Options.Label"/><Description resid="Options.Description"/></Supertip>${icons}
              <Action xsi:type="ShowTaskpane"><SourceLocation resid="Taskpane.Url"/>${v11 ? '<SupportsPinning>true</SupportsPinning>' : ''}</Action>
            </Item>
          </Items>
        </Control>
      </Group></OfficeTab></ExtensionPoint>
    </DesktopFormFactor></Host></Hosts>
    <Resources>
      <bt:Images><bt:Image id="Icon.16" DefaultValue="${origin}/assets/icon-16.png"/><bt:Image id="Icon.32" DefaultValue="${origin}/assets/icon-32.png"/><bt:Image id="Icon.64" DefaultValue="${origin}/assets/icon-64.png"/><bt:Image id="Icon.80" DefaultValue="${origin}/assets/icon-80.png"/></bt:Images>
      <bt:Urls><bt:Url id="Commands.Url" DefaultValue="${origin}/commands.html"/><bt:Url id="Taskpane.Url" DefaultValue="${origin}/taskpane.html"/></bt:Urls>
      <bt:ShortStrings><bt:String id="Group.Label" DefaultValue="邮件翻译"/><bt:String id="Menu.Label" DefaultValue="翻译"/><bt:String id="Translate.Label" DefaultValue="翻译邮件"/><bt:String id="Options.Label" DefaultValue="翻译选项"/></bt:ShortStrings>
      <bt:LongStrings><bt:String id="Menu.Description" DefaultValue="翻译邮件并管理翻译选项。"/><bt:String id="Translate.Description" DefaultValue="翻译整封邮件并在原文位置显示译文。"/><bt:String id="Options.Description" DefaultValue="设置目标语言、自动翻译偏好和不翻译的语言。"/></bt:LongStrings>
    </Resources>`;
}
const manifest = `<?xml version="1.0" encoding="UTF-8"?>
<OfficeApp xmlns="http://schemas.microsoft.com/office/appforoffice/1.1" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance" xmlns:bt="http://schemas.microsoft.com/office/officeappbasictypes/1.0" xsi:type="MailApp">
  <Id>8cf39d33-31e8-4ed8-8ac0-f3e51d602286</Id><Version>1.1.0.0</Version><ProviderName>邮件翻译</ProviderName><DefaultLocale>zh-CN</DefaultLocale>
  <DisplayName DefaultValue="邮件翻译"/><Description DefaultValue="通过 Azure 翻译整封邮件，保留正文排版，支持 Office SSO 和翻译选项。"/>
  <IconUrl DefaultValue="${origin}/assets/icon-32.png"/><HighResolutionIconUrl DefaultValue="${origin}/assets/icon-80.png"/><SupportUrl DefaultValue="${origin}/taskpane.html"/>
  <AppDomains><AppDomain>${origin}</AppDomain></AppDomains><Hosts><Host Name="Mailbox"/></Hosts>
  <Requirements><Sets><Set Name="Mailbox" MinVersion="1.5"/></Sets></Requirements>
  <FormSettings><Form xsi:type="ItemRead"><DesktopSettings><SourceLocation DefaultValue="${origin}/taskpane.html"/><RequestedHeight>450</RequestedHeight></DesktopSettings></Form></FormSettings>
  <Permissions>ReadWriteItem</Permissions><Rule xsi:type="RuleCollection" Mode="Or"><Rule xsi:type="ItemIs" ItemType="Message" FormType="Read"/></Rule><DisableEntityHighlighting>false</DisableEntityHighlighting>
  <VersionOverrides xmlns="http://schemas.microsoft.com/office/mailappversionoverrides" xsi:type="VersionOverridesV1_0">
    ${contents(false)}
    <VersionOverrides xmlns="http://schemas.microsoft.com/office/mailappversionoverrides/1.1" xsi:type="VersionOverridesV1_1">
      ${contents(true)}
      <WebApplicationInfo><Id>${xml(config.clientId)}</Id><Resource>${xml(config.resource)}</Resource><Scopes><Scope>openid</Scope><Scope>profile</Scope><Scope>User.Read</Scope></Scopes></WebApplicationInfo>
    </VersionOverrides>
  </VersionOverrides>
</OfficeApp>
`;
fs.writeFileSync(path.resolve(__dirname, '../manifest.xml'), manifest);
console.log('Generated manifest.xml using APP_BASE_URL, CLIENT_ID and SSO_RESOURCE.');
