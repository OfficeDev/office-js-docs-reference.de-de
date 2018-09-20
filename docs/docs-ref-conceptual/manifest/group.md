# <a name="group-element"></a>Group-Element

Definiert eine Gruppe von Benutzeroberflächensteuerelemente auf einer Registerkarte.  In benutzerdefinierten Registerkarten kann das Add-In bis zu 10 Gruppen erstellen. Jede Gruppe ist auf 6 Steuerelemente begrenzt, unabhängig davon, in welcher Registerkarte es angezeigt wird. Add-Ins sind auf eine benutzerdefinierte Registerkarte begrenzt.

## <a name="attributes"></a>Attribute

|  Attribut  |  Erforderlich  |  Beschreibung  |
|:-----|:-----|:-----|
|  [id](#id-attribute)  |  Ja  | Eine eindeutige ID für die Gruppe.|

### <a name="id-attribute"></a>id-Attribut

Erforderlich. Eindeutiger Bezeichner für die Gruppe. Es ist eine Zeichenfolge mit maximal 125 Zeichen. Dies muss im Manifest eindeutig sein, andernfalls wird die Gruppe nicht gerendert.

## <a name="child-elements"></a>Untergeordnete Elemente
|  Element |  Erforderlich  |  Beschreibung  |
|:-----|:-----|:-----|
|  [Label](#label)      | Ja |  Die Bezeichnung für CustomTab oder eine Gruppe  |
|  [Control](#control)    | Ja |  Auflistung von einem oder mehreren Control-Objekten.  |

### <a name="label"></a>Label 

Erforderlich. Die Beschriftung der Gruppe. Das Attribut  **resid** muss auf den Wert des Attributs **id** eines **String**-Elements im  **ShortStrings**-Element im  [Resources](resources.md) -Element festgelegt werden.

### <a name="control"></a>Control
Eine Gruppe erfordert mindestens ein Steuerelement.

```xml
<Group id="msgreadCustomTab.grp1">
    <Label resid="residCustomTabGroupLabel"/>
    <Control xsi:type="Button" id="Button2">
    <!-- information on the control -->
    </Control>
    <!-- other controls, as needed -->
</Group>
```