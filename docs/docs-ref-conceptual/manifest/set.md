# <a name="set-element"></a>Set-Element

Gibt einen Anforderungssatz aus der JavaScript-API für Office an, die Ihr Office-Add-In zum Aktivieren benötigt.

**Add-In-Typ:** Inhalt, Aufgabenbereich, E-Mail

## <a name="syntax"></a>Syntax

```XML
<Set Name="string" MinVersion="n .n">
```

## <a name="contained-in"></a>Enthalten in

[Sätzen](sets.md)

## <a name="attributes"></a>Attribute

|**Attribut**|**Typ**|**Erforderlich**|**Beschreibung**|
|:-----|:-----|:-----|:-----|
|Name|string|erforderlich|Der Name eines [Anforderungssatzes](https://docs.microsoft.com/office/dev/add-ins/develop/office-versions-and-requirement-sets).|
|MinVersion|string|Optional|Gibt die Mindestversion des API-Satzes an, die für das Add-In erforderlich ist. Setzt den Wert von **DefaultMinVersion** außer Kraft, wenn dies im übergeordneten [Sets](sets.md)-Element angegeben ist.|

## <a name="remarks"></a>Hinweise

Weitere Informationen zu anforderungssätze finden Sie unter [Office-Versionen und Anforderung festgelegt](https://docs.microsoft.com/office/dev/add-ins/develop/office-versions-and-requirement-sets).

Weitere Informationen zum **MinVersion**-Attribut des **Set**-Elements und zum **DefaultMinVersion**-Attribut des **Sets**-Elements finden Sie unter [Festlegen des Requirements-Element im Manifest](https://docs.microsoft.com/office/dev/add-ins/develop/specify-office-hosts-and-api-requirements#set-the-requirements-element-in-the-manifest).

> [!IMPORTANT] 
> Für Mail-add-ins, ist nur eine `"Mailbox"` Anforderung verfügbar festgelegt. Diese Anforderungssatz enthält die gesamte Teilmenge der API-Unterstützung in Mail-add-ins für Outlook, und geben Sie die `"Mailbox"` Anforderung in Ihrer Mail-add-in des Manifests festgelegt (es ist nicht optional, wie die Groß-/Kleinschreibung für Inhalts- und Aufgabenbereich Bereich-add-ins ist). Darüber hinaus können Sie Unterstützung für bestimmte Methoden in Mail-add-ins nicht deklarieren.
