# <a name="customtab-element"></a>CustomTab-Element

Im Menüband geben Sie Registerkarte und Gruppe für die Add-In-Befehle an. Dies kann entweder in der Standardregisterkarte sein (  **Home**,  **Message** oder  **Meeting**), oder in einer durch das Add-In definierten benutzerdefinierten Registerkarte.

In benutzerdefinierten Registerkarten kann das Add-In bis zu 10 Gruppen erstellen. Jede Gruppe ist auf 6 Steuerelemente begrenzt, unabhängig davon, in welcher Registerkarte es angezeigt wird. Add-Ins sind auf eine benutzerdefinierte Registerkarte begrenzt.

Das Attribut  **id** muss im Manifest eindeutig sein.

## <a name="child-elements"></a>Untergeordnete Elemente

|  Element |  Erforderlich  |  Beschreibung  |
|:-----|:-----|:-----|
|  [Group](group.md)      | Ja |  Definiert eine Gruppe von Befehlen.  |
|  [Label](#label-tab)      | Ja |  Die Bezeichnung für CustomTab oder eine Gruppe  |
|  [Control](control.md)    | Ja |  Eine Auflistung von einem oder mehreren Control-Objekten.  |

### <a name="group"></a>Group

Erforderlich. Siehe [Group-Element](group.md).

### <a name="label-tab"></a>Label (Tab)

Erforderlich. Die Beschriftung der benutzerdefinierten Registerkarte. Das Attribut  **resid** muss auf den Wert des Attributs **id** eines **String**-Elements im  **ShortStrings**-Element im  [Resources](resources.md) -Element festgelegt werden.


## <a name="customtab-example"></a>CustomTab-Beispiel

```xml
<ExtensionPoint xsi:type="MessageReadCommandSurface">
  <CustomTab id="TabCustom1">
    <Group id="msgreadCustomTab.grp1">
    </Group>
    <Label resid="customTabLabel1"/>
  </CustomTab>
</ExtensionPoint>
```