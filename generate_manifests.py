import uuid

# EDIT THIS: where you will host taskpane.html + assets/ (e.g. GitHub Pages root)
BASE_URL = "https://YOUR-USERNAME.github.io/ai-email-addin"

providers = {
    "openai":     {"name": "OpenAI (GPT) Email Assistant",      "domain": "api.openai.com",     "desc": "Summarize emails and draft replies using OpenAI GPT models."},
    "claude":     {"name": "Claude Email Assistant",             "domain": "api.anthropic.com",  "desc": "Summarize emails and draft replies using Anthropic Claude."},
    "gemini":     {"name": "Gemini Email Assistant",             "domain": "generativelanguage.googleapis.com", "desc": "Summarize emails and draft replies using Google Gemini."},
    "qwen":       {"name": "Qwen Email Assistant",               "domain": "dashscope-intl.aliyuncs.com", "desc": "Summarize emails and draft replies using Alibaba Qwen (DashScope)."},
    "deepseek":   {"name": "DeepSeek Email Assistant",           "domain": "api.deepseek.com",   "desc": "Summarize emails and draft replies using DeepSeek."},
    "grok":       {"name": "Grok Email Assistant",               "domain": "api.x.ai",            "desc": "Summarize emails and draft replies using xAI Grok."},
    "mistral":    {"name": "Mistral Email Assistant",            "domain": "api.mistral.ai",      "desc": "Summarize emails and draft replies using Mistral AI."},
    "groq":       {"name": "Groq Email Assistant",               "domain": "api.groq.com",        "desc": "Summarize emails and draft replies using Groq's fast open-model API."},
    "together":   {"name": "Together AI Email Assistant",        "domain": "api.together.xyz",    "desc": "Summarize emails and draft replies using Together AI open models."},
    "perplexity": {"name": "Perplexity Email Assistant",         "domain": "api.perplexity.ai",   "desc": "Summarize emails and draft replies using Perplexity."},
}

TEMPLATE = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<OfficeApp xmlns="http://schemas.microsoft.com/office/appforoffice/1.1"
  xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance"
  xmlns:bt="http://schemas.microsoft.com/office/officeappbasictypes/1.0"
  xmlns:mailappor="http://schemas.microsoft.com/office/mailappversionoverrides/1.0" xsi:type="MailApp">
  <Id>{guid}</Id>
  <Version>1.0.0.0</Version>
  <ProviderName>{key_title}</ProviderName>
  <DefaultLocale>en-US</DefaultLocale>
  <DisplayName DefaultValue="{name}"/>
  <Description DefaultValue="{desc}"/>
  <IconUrl DefaultValue="{base}/assets/icon-{key}-32.png"/>
  <HighResolutionIconUrl DefaultValue="{base}/assets/icon-{key}-80.png"/>
  <SupportUrl DefaultValue="{base}"/>
  <AppDomains>
    <AppDomain>https://{domain}</AppDomain>
  </AppDomains>
  <Hosts><Host Name="Mailbox"/></Hosts>
  <Requirements><Sets><Set Name="Mailbox" MinVersion="1.1"/></Sets></Requirements>
  <FormSettings><Form xsi:type="ItemRead"><DesktopSettings>
    <SourceLocation DefaultValue="{base}/taskpane.html?provider={key}"/>
    <RequestedHeight>480</RequestedHeight>
  </DesktopSettings></Form></FormSettings>
  <Permissions>ReadWriteItem</Permissions>
  <Rule xsi:type="RuleCollection" Mode="Or">
    <Rule xsi:type="ItemIs" ItemType="Message" FormType="Read"/>
    <Rule xsi:type="ItemIs" ItemType="Message" FormType="Edit"/>
  </Rule>
  <DisableEntityHighlighting>false</DisableEntityHighlighting>
  <VersionOverrides xmlns="http://schemas.microsoft.com/office/mailappversionoverrides" xsi:type="VersionOverridesV1_0">
    <Requirements><bt:Sets DefaultMinVersion="1.3"><bt:Set Name="Mailbox"/></bt:Sets></Requirements>
    <Hosts><Host xsi:type="MailHost"><DesktopFormFactor>
      <ExtensionPoint xsi:type="MessageReadCommandSurface"><OfficeTab id="TabDefault"><Group id="msgReadGroup_{key}">
        <Label resid="GroupLabel"/><Control xsi:type="Button" id="msgReadOpenPaneButton_{key}">
          <Label resid="TaskpaneButton.Label"/><Supertip><Title resid="TaskpaneButton.Label"/><Description resid="TaskpaneButton.Tooltip"/></Supertip>
          <Icon><bt:Image size="16" resid="Icon.16x16"/><bt:Image size="32" resid="Icon.32x32"/><bt:Image size="80" resid="Icon.80x80"/></Icon>
          <Action xsi:type="ShowTaskpane"><SourceLocation resid="Taskpane.Url"/></Action>
        </Control>
      </Group></OfficeTab></ExtensionPoint>
      <ExtensionPoint xsi:type="MessageComposeCommandSurface"><OfficeTab id="TabDefault"><Group id="msgComposeGroup_{key}">
        <Label resid="GroupLabel"/><Control xsi:type="Button" id="msgComposeOpenPaneButton_{key}">
          <Label resid="TaskpaneButton.Label"/><Supertip><Title resid="TaskpaneButton.Label"/><Description resid="TaskpaneButton.Tooltip"/></Supertip>
          <Icon><bt:Image size="16" resid="Icon.16x16"/><bt:Image size="32" resid="Icon.32x32"/><bt:Image size="80" resid="Icon.80x80"/></Icon>
          <Action xsi:type="ShowTaskpane"><SourceLocation resid="Taskpane.Url"/></Action>
        </Control>
      </Group></OfficeTab></ExtensionPoint>
    </DesktopFormFactor></Host></Hosts>
    <Resources>
      <bt:Images>
        <bt:Image id="Icon.16x16" DefaultValue="{base}/assets/icon-{key}-16.png"/>
        <bt:Image id="Icon.32x32" DefaultValue="{base}/assets/icon-{key}-32.png"/>
        <bt:Image id="Icon.80x80" DefaultValue="{base}/assets/icon-{key}-80.png"/>
      </bt:Images>
      <bt:Urls><bt:Url id="Taskpane.Url" DefaultValue="{base}/taskpane.html?provider={key}"/></bt:Urls>
      <bt:ShortStrings>
        <bt:String id="GroupLabel" DefaultValue="{short_label}"/>
        <bt:String id="TaskpaneButton.Label" DefaultValue="{short_label}"/>
      </bt:ShortStrings>
      <bt:LongStrings><bt:String id="TaskpaneButton.Tooltip" DefaultValue="Summarize or draft replies with {short_label}"/></bt:LongStrings>
    </Resources>
  </VersionOverrides>
</OfficeApp>
"""

for key, info in providers.items():
    xml = TEMPLATE.format(
        guid=str(uuid.uuid4()),
        key=key,
        key_title=key.capitalize(),
        name=info["name"],
        desc=info["desc"],
        domain=info["domain"],
        base=BASE_URL,
        short_label=info["name"].replace(" Email Assistant", ""),
    )
    with open(f"manifests/manifest-{key}.xml", "w") as f:
        f.write(xml)

print("Generated", len(providers), "manifests.")
