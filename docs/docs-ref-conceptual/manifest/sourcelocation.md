# <a name="sourcelocation-element"></a>SourceLocation-Element

Gibt die Quelldateispeicherorte für Ihr Office-Add-In als URL an, die zwischen 1 und 2018 Zeichen lang ist. Der Quellspeicherort muss eine HTTPS-Adresse, kein Dateipfad sein.

**Add-In-Typ:** Inhalt, Aufgabenbereich, E-Mail

## <a name="syntax"></a>Syntax

```XML
<SourceLocation DefaultValue="string" />
```

## <a name="contained-in"></a>Enthalten in

- [DefaultSettings](defaultsettings.md) (Inhalts- und Aufgabenbereich-Add-Ins)
- [FormSettings](formsettings.md) (E-Mail-Add-Ins)
- [ExtensionPoint](extensionpoint.md) (E-Mail-Kontext-Add-Ins)

## <a name="can-contain"></a>Kann enthalten

[Override](override.md)

## <a name="attributes"></a>Attribute

|**Attribut**|**Typ**|**Erforderlich**|**Beschreibung**|
|:-----|:-----|:-----|:-----|
|DefaultValue|URL|erforderlich|Gibt den Standardwert für diese Einstellung für das im [DefaultLocale](defaultlocale.md)-Element angegebene Gebietsschema an.|
