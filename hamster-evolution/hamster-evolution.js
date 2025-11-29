<?xml version="1.0" encoding="UTF-8"?>
<OfficeApp
  xmlns="http://schemas.microsoft.com/office/appforoffice/1.1"
  xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance"
  xmlns:bt="http://schemas.microsoft.com/office/officeappbasictypes/1.0"
  xmlns:mailappor="http://schemas.microsoft.com/office/mailappversionoverrides/1.0"
  xsi:type="MailApp">

  <!-- New GUID for this add-in -->
  <Id>1c3434d6-d45b-40fa-afbc-43a9b2f70a7c</Id>
  <Version>1.0.0.0</Version>
  <ProviderName>Billy Taylor</ProviderName>
  <DefaultLocale>en-US</DefaultLocale>

  <DisplayName DefaultValue="Hamster Evolution" />
  <Description DefaultValue="Hamster Evolution codes outgoing emails and logs them." />

  <!-- Icon in the add-ins dialog -->
  <IconUrl DefaultValue="https://billyht247.github.io/Outlook_Addin/evolution-80.png" />

  <SupportUrl DefaultValue="https://support.microsoft.com" />

  <Hosts>
    <Host Name="Mailbox" />
  </Hosts>

  <!-- Base requirements: we only target clients with commands support -->
  <Requirements>
    <Sets>
      <Set Name="Mailbox" MinVersion="1.3" />
    </Sets>
  </Requirements>

  <!-- Legacy FormSettings for old clients; modern clients will use VersionOverrides -->
  <FormSettings>
    <Form xsi:type="ItemRead">
      <DesktopSettings>
        <SourceLocation DefaultValue="https://billyht247.github.io/Outlook_Addin/hamster-evolution.html" />
        <RequestedHeight>300</RequestedHeight>
      </DesktopSettings>
    </Form>
  </FormSettings>

  <!-- Read/write item so we can change subject and BCC -->
  <Permissions>ReadWriteItem</Permissions>

  <!-- Legacy activation rule for old clients (message read). -->
  <Rule xsi:type="RuleCollection" Mode="Or">
    <Rule xsi:type="ItemIs" ItemType="Message" FormType="Read" />
  </Rule>

  <DisableEntityHighlighting>false</DisableEntityHighlighting>

  <!-- Commands + compose button for modern Outlook -->
  <VersionOverrides
    xmlns="http://schemas.microsoft.com/office/mailappversionoverrides"
    xsi:type="VersionOverridesV1_0">

    <VersionOverrides
      xmlns="http://schemas.microsoft.com/office/mailappversionoverrides/1.1"
      xsi:type="VersionOverridesV1_1">

      <Requirements>
        <bt:Sets DefaultMinVersion="1.3">
          <bt:Set Name="Mailbox" />
        </bt:Sets>
      </Requirements>

      <Hosts>
        <Host xsi:type="MailHost">
          <DesktopFormFactor>

            <!-- FunctionFile is required; we reuse the same page as our task pane -->
            <FunctionFile resid="EmailCode.TaskPane.Url" />

            <!-- Compose ribbon button -->
            <ExtensionPoint xsi:type="MessageComposeCommandSurface">
              <OfficeTab id="TabDefault">
                <Group id="EmailCode.Group">
                  <Label resid="EmailCode.Group.Label" />
                  <Control xsi:type="Button" id="EmailCode.Button">
                    <Label resid="EmailCode.Button.Label" />
                    <Supertip>
                      <Title resid="EmailCode.Button.Label" />
                      <Description resid="EmailCode.Button.Tooltip" />
                    </Supertip>
                    <Icon>
                      <bt:Image size="16" resid="Icon.16x16" />
                      <bt:Image size="32" resid="Icon.32x32" />
                      <bt:Image size="80" resid="Icon.80x80" />
                    </Icon>
                    <Action xsi:type="ShowTaskpane">
                      <SourceLocation resid="EmailCode.TaskPane.Url" />
                    </Action>
                  </Control>
                </Group>
              </OfficeTab>
            </ExtensionPoint>

          </DesktopFormFactor>
        </Host>
      </Hosts>

      <Resources>
        <bt:Images>
          <bt:Image id="Icon.16x16"
                    DefaultValue="https://billyht247.github.io/Outlook_Addin/evolution-16.png" />
          <bt:Image id="Icon.32x32"
                    DefaultValue="https://billyht247.github.io/Outlook_Addin/evolution-32.png" />
          <bt:Image id="Icon.80x80"
                    DefaultValue="https://billyht247.github.io/Outlook_Addin/evolution-80.png" />
        </bt:Images>

        <bt:Urls>
          <!-- Task pane + function file -->
          <bt:Url id="EmailCode.TaskPane.Url"
                  DefaultValue="https://billyht247.github.io/Outlook_Addin/hamster-evolution.html" />
        </bt:Urls>

        <bt:ShortStrings>
          <bt:String id="EmailCode.Group.Label" DefaultValue="Hamster Evolution" />
          <bt:String id="EmailCode.Button.Label" DefaultValue="Hamster Evolution" />
        </bt:ShortStrings>

        <bt:LongStrings>
          <bt:String id="EmailCode.Button.Tooltip"
                     DefaultValue="Open Hamster Evolution to set WHEN / TYPE / TIME and add the logging BCC." />
        </bt:LongStrings>
      </Resources>

    </VersionOverrides>
  </VersionOverrides>
</OfficeApp>