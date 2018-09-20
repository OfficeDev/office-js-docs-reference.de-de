# <a name="defaultsettings-element"></a>DefaultSettings-Element

Gibt den Standardspeicherort für die Quell- und anderen Standardeinstellungen für den Inhalt oder task Pane-add-in.

**Add-In-Typ:** Inhalt, Aufgabenbereich

## <a name="syntax"></a>Syntax

```XML
<DefaultSettings>
  ...
</DefaultSettings>
```

## <a name="contained-in"></a>Enthalten in

[OfficeApp](officeapp.md)

## <a name="can-contain"></a>Kann enthalten

|**Element**|**Inhalt**|**E-Mail**|**TaskPane**|
|:-----|:-----|:-----|:-----|
|[SourceLocation](sourcelocation.md)|x||x|
|[RequestedWidth](requestedwidth.md)|x|||
|[RequestedHeight](requestedheight.md)|x|||

## <a name="remarks"></a>Bemerkungen

Der Quellspeicherort und andere Einstellungen im **DefaultSettings**-Element gelten nur für Inhalts- und Aufgabenbereich-Add-Ins. Bei E-Mail-Add-Ins geben Sie die Standardspeicherorte für Quelldateien und andere Standardeinstellungen im [FormSettings](formsettings.md)-Element an.

