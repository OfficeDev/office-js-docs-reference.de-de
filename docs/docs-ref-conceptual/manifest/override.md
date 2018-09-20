# <a name="override-element"></a>Override-Element

Stellt eine Methode zum Angeben des Werts einer Einstellung für ein zusätzliches Gebietsschema bereit.

**Add-In-Typ:** Inhalt, Aufgabenbereich, E-Mail

## <a name="syntax"></a>Syntax

```XML
<Override Locale="string" Value="string" />
```

## <a name="contained-in"></a>Enthalten in

|**Element**|
|:-----|
|[CitationText](citationtext.md)|
|[Beschreibung](description.md)|
|[DictionaryName](dictionaryname.md)|
|[DictionaryHomePage](dictionaryhomepage.md)|
|[Anzeigename](displayname.md)|
|[HighResolutionIconUrl](highresolutioniconurl.md)|
|[IconUrl](iconurl.md)|
|[QueryUri](queryuri.md)|
|[SourceLocation](sourcelocation.md)|
|[SupportUrl](supporturl.md)|

## <a name="attributes"></a>Attribute

|**Attribut**|**Typ**|**Erforderlich**|**Beschreibung**|
|:-----|:-----|:-----|:-----|
|Locale|string|erforderlich|Gibt den Kulturnamen des Gebietsschemas für diese Außerkraftsetzung im BCP 47-Sprachtagformat an, z. B.  `"en-US"`.|
|Wert|string|erforderlich|Gibt den Wert der für das angegebene Gebietsschema ausgedrückten Einstellung an.|

## <a name="see-also"></a>Siehe auch

- [Lokalisierung für Office-Add-Ins](https://docs.microsoft.com/office/dev/add-ins/develop/localization)
    
