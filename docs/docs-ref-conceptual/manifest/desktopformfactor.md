# <a name="desktopformfactor-element"></a>DesktopFormFactor-Element

Gibt die Einstellungen für ein Add-In für den  Desktopformfaktor an. Der Desktopformfaktor umfasst Office für Windows, Office für Mac und Office Online. Es enthält alle Add-In-Informationen für den Desktopformfaktor mit Ausnahme des **Resources**-Knotens.

Jede DesktopFormFactor-Definition enthält das **FunctionFile**-Element und mindestens ein **ExtensionPoint**-Element. Weitere Informationen finden Sie unter [FunctionFile-Element](functionfile.md) und [ExtensionPoint-Element](extensionpoint.md).

> [!IMPORTANT]
> Das SupportsSharedFolders-Element ist nur in der Outlook-add-ins Preview Anforderung festgelegt mit Exchange Online verfügbar.
> Add-ins, die dieses Elements in der Office Store oder zentralisierte Bereitstellung sind nicht zulässig.

## <a name="child-elements"></a>Untergeordnete Elemente

| Element                               | Erforderlich | Beschreibung  |
|:--------------------------------------|:--------:|:-------------|
| [ExtensionPoint](extensionpoint.md)   | Ja      | Definiert, wo ein Add-In Funktionen verfügbar macht. |
| [FunctionFile](functionfile.md)       | Ja      | Eine URL zu einer Datei, die JavaScript-Funktionen enthält.|
| [GetStarted](getstarted.md)           | Nein       | Definiert die Beschriftung, die angezeigt wird, wenn Sie das Add-In in Word-, Excel- oder PowerPoint-Hosts installieren. |
| SupportsSharedFolders                 | Nein       | Definiert, ob das Outlook-add-in in Situationen Stellvertretung verfügbar ist und ist standardmäßig auf *false* festgelegt. Preview-Anforderungssatz.|

## <a name="desktopformfactor-example"></a>DesktopFormFactor-Beispiel

```xml
...
<Hosts>
  <Host xsi:type="Presentation">
    <DesktopFormFactor>
      <FunctionFile resid="residDesktopFuncUrl" />
      <GetStarted>
        <!-- GetStarted callout -->
      </GetStarted>
      <ExtensionPoint xsi:type="PrimaryCommandSurface">
        <!-- information on this extension point -->
      </ExtensionPoint>
      <!-- possibly more ExtensionPoint elements -->
    </DesktopFormFactor>
  </Host>
</Hosts>
...
```
