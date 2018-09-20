# <a name="iconurl-element"></a>IconUrl-Element

Gibt die URL des Bilds an, das zur Darstellung Ihres Office-Add-Ins in der Einfüge-UX und im Office Store verwendet wird.

**Add-In-Typ:** Inhalt, Aufgabenbereich, E-Mail

## <a name="syntax"></a>Syntax

```XML
<IconUrl DefaultValue="string" />
```

## <a name="can-contain"></a>Kann enthalten

[Override](override.md)

## <a name="attributes"></a>Attribute

|**Attribut**|**Typ**|**Erforderlich**|**Beschreibung**|
|:-----|:-----|:-----|:-----|
|DefaultValue|string|erforderlich|Gibt den Standardwert für diese Einstellung an, der für das im [DefaultLocale](defaultlocale.md)-Element angegebene Gebietsschema ausgedrückt wird.|

## <a name="remarks"></a>Hinweise

Bei E-Mail-Ad-Ins wird das Symbol in der Benutzeroberfläche **Datei**  >  **Add-Ins verwalten** (Outlook) oder **Einstellungen**  >  **Add-Ins verwalten** (Outlook Web App) angezeigt. Bei Inhalts- oder Aufgabenbereich-Add-Ins wird das Symbol in der Benutzeroberfläche **Einfügen**  >  **Add-Ins** angezeigt. Bei allen Add-In-Typen wird das Symbol auf der Office Store-Website verwendet, wenn Sie Ihr Add-In im Office Store veröffentlichen.

Das Bild muss eines der folgenden Dateiformate: GIF, JPG, PNG, EXIF, BMP oder TIFF. Bei Inhalts- und Aufgabenbereich-apps muss das angegebene Bild 32 x 32 Pixel sein. Für Mail-apps muss das Bild 64 x 64 Pixel. Sie sollten auch ein Symbol für die Verwendung mit Office-hostanwendungen unter hohe DPI Bildschirme im [HighResolutionIconUrl](highresolutioniconurl.md) -Element angeben. Weitere Informationen finden Sie im Abschnitt _Erstellen einer konsistenten visuellen Identität für Ihre app_ in den [effektiven erstellen-Codebeispielen in Elemente verwenden und innerhalb von Office](https://docs.microsoft.com/office/dev/store/create-effective-office-store-listings#create-a-consistent-visual-identity).
