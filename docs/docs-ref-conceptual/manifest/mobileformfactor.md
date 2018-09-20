# <a name="mobileformfactor-element"></a>MobileFormFactor-Element

Gibt die Einstellungen für ein Add-In für den mobilen Formfaktor an. Es enthält alle Add-In-Informationen für den mobilen Formfaktor mit Ausnahme des **Resources**-Knotens.

Jede **MobileFormFactor**-Definition enthält das **FunctionFile**-Element und mindestens ein **ExtensionPoint**-Elemente. Weitere Informationen finden Sie unter [FunctionFile-Element](functionfile.md) und [ExtensionPoint-Element](extensionpoint.md)

Das **MobileFormFactor**-Element ist im VersionOverrides Schema 1.1 definiert. Das enthaltende [VersionOverrides](versionoverrides.md)-Element muss einen `xsi:type`-Attributwert von `VersionOverridesV1_1` aufweisen.

## <a name="child-elements"></a>Untergeordnete Elemente

| Element                               | Erforderlich | Beschreibung  |
|:--------------------------------------|:--------:|:-------------|
| [ExtensionPoint](extensionpoint.md) | Ja      | Definiert, wo ein Add-In Funktionen verfügbar macht. |
| [FunctionFile](functionfile.md)     | Ja      | Eine URL zu einer Datei, die JavaScript-Funktionen enthält.|

## <a name="mobileformfactor-example"></a>MobileFormFactor-Beispiel

```xml
...
<Hosts>
  <Host xsi:type="MailHost">
    ...
    <MobileFormFactor>
      <FunctionFile resid="residUILessFunctionFileUrl" />
      <ExtensionPoint xsi:type="MobileMessageReadCommandSurface">
        <!-- information on this extension point -->
      </ExtensionPoint> 
      <!-- possibly more ExtensionPoint elements -->
    </MobileFormFactor>
  </Host>
</Hosts>
...
```
