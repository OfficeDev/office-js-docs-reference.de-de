# <a name="permissions-element"></a>Permissions-Element

Gibt die Ebene des API-Zugriffs für Ihr Office-Add-In an. Sie müssen Berechtigungen auf der Grundlage der „geringsten Rechte“ anfordern.

**Add-In-Typ:** Inhalt, Aufgabenbereich, E-Mail

## <a name="syntax"></a>Syntax

Für Inhalts- und Aufgabenbereich-Add-Ins:

```XML
 <Permissions> [Restricted | ReadDocument | ReadAllDocument | WriteDocument | ReadWriteDocument]</Permissions>
```

Für Mail-add-ins

```XML
 <Permissions>[Restricted | ReadItem | ReadWriteItem | ReadWriteMailbox]</Permissions>
```

## <a name="contained-in"></a>Enthalten in

[OfficeApp](officeapp.md)

## <a name="remarks"></a>Hinweise

Weitere Informationen finden Sie unter [Anfordern von Berechtigungen für API-Verwendung in Inhalts- und Aufgabenbereich-Add-Ins](https://docs.microsoft.com/office/dev/add-ins/develop/requesting-permissions-for-api-use-in-content-and-task-pane-add-ins) und [Grundlegendes zu Outlook-Add-In-Berechtigungen](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions).
