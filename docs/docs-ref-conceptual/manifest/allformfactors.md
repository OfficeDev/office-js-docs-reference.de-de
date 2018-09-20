# <a name="allformfactors-element"></a>AllFormFactors-Element

Gibt die Einstellungen für ein Add-In für alle Formfaktoren an. Derzeit ist das einzige Feature **AllFormFactors** mit benutzerdefinierten Funktionen. **AllFormFactors** ist ein erforderliches Element aus, wenn Sie benutzerdefinierte Funktionen verwenden.

## <a name="child-elements"></a>Untergeordnete Elemente

|  Element |  Erforderlich  |  Beschreibung  |
|:-----|:-----|:-----|
|  [ExtensionPoint](extensionpoint.md) |  Ja |  Definiert, wo ein Add-In Funktionen verfügbar macht. |

## <a name="allformfactors-example"></a>AllFormFactors-Beispiel

```xml
<Hosts>
    <Host xsi:type="Workbook">
        <AllFormFactors>
            <ExtensionPoint xsi:type="CustomFunctions">
                    <!-- Information on this extension point -->
            </ExtensionPoint>
        </AllFormFactors>
    </Host>
</Hosts>
```
