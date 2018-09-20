# <a name="host-element"></a>Host-Element

Gibt einen einzelnen Office-Anwendungstyp an, in dem das Add-In aktiviert werden sollte.

> [!IMPORTANT] 
> Die Syntax der **Host** -Element hängt davon ab, ob das Element innerhalb des [grundlegenden Manifest](#basic-manifest) oder in den [VersionOverrides](#versionoverrides-node) -Knoten definiert ist. Die Funktionalität ist jedoch dieselbe.  

## <a name="basic-manifest"></a>Grundlegendes Manifest

Wenn dies im grundlegenden Manifest (unter [OfficeApp](officeapp.md)) definiert ist, wird der Hosttyp vom `Name`-Attribut bestimmt.   

### <a name="attributes"></a>Attribute

| Attribut     | Typ   | Erforderlich | Beschreibung                                      |
|:--------------|:-------|:---------|:-------------------------------------------------|
| [Name](#name) | string | erforderlich | Der Name des Typs der Office-Hostanwendung. |

### <a name="name"></a>Name
Gibt den Hosttyp an, auf den von diesem Add-In abgezielt wird. Bei dem Wert muss es sich um Folgendes handeln:

- `Document` (Word)
- `Database` (Access)
- `Mailbox` (Outlook)
- `Notebook` (OneNote)
- `Presentation` (PowerPoint)
- `Project` (Project)
- `Workbook` (Excel)

### <a name="example"></a>Beispiel
```xml
<Hosts>
    <Host Name="Mailbox">
    </Host>
</Hosts>
```

## <a name="versionoverrides-node"></a>VersionOverrides-Knoten
Wenn dies im [VersionOverrides](versionoverrides.md) definiert ist, wird der Hosttyp vom `xsi:type`-Attribut bestimmt. 

### <a name="attributes"></a>Attribute

|  Attribut  |  Erforderlich  |  Beschreibung  |
|:-----|:-----|:-----|
|  [xsi:type](#xsitype)  |  Ja  | Beschreibt den Office-Host, für den diese Einstellungen gelten.|

### <a name="child-elements"></a>Untergeordnete Elemente

|  Element |  Erforderlich  |  Beschreibung  |
|:-----|:-----|:-----|
|  [DesktopFormFactor](desktopformfactor.md)    |  Ja   |  Definiert die Einstellungen für den Desktopformfaktor an. |
|  [MobileFormFactor](mobileformfactor.md)    |  Nein   |  Definiert die Einstellungen für den mobilen Formfaktor. **Hinweis:** Dieses Element wird nur in Outlook für iOS unterstützt. |
|  [AllFormFactors](allformfactors.md)    |  Nein   |  Definiert die Einstellungen für alle Formfaktoren. Nur von benutzerdefinierte Funktionen in Excel verwendet. |

### <a name="xsitype"></a>xsi:type

Steuert, auf welchen Office-Host (Word, Excel, PowerPoint, Outlook, OneNote) die enthaltenen Einstellungen angewendet werden. Bei dem Wert muss es sich um Folgendes handeln:

- `Document` (Word)
- `MailHost` (Outlook)    
- `Notebook` (OneNote)
- `Presentation` (PowerPoint)
- `Workbook` (Excel)

## <a name="host-example"></a>Host-Beispiel 
```xml
<Hosts>
    <Host xsi:type="MailHost">
        <!-- Host Settings -->
    </Host>
</Hosts>
```
