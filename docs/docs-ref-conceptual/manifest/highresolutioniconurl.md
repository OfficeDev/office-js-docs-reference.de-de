# <a name="highresolutioniconurl-element"></a>HighResolutionIconUrl-Element

Gibt die URL des Bilds an, das zur Darstellung Ihres Office-Add-Ins in der Einfüge-UX und im Office Store auf Bildschirmen mit hohem DPI-Wert verwendet wird.

**Add-In-Typ:** Inhalt, Aufgabenbereich, E-Mail

## <a name="syntax"></a>Syntax

```XML
<HighResolutionIconUrl DefaultValue="string" />
```

## <a name="can-contain"></a>Kann enthalten

[Override](override.md)

## <a name="attributes"></a>Attribute

|**Attribut**|**Typ**|**Erforderlich**|**Beschreibung**|
|:-----|:-----|:-----|:-----|
|DefaultValue|Zeichenfolge (URL)|erforderlich|Gibt den Standardwert für diese Einstellung an, der für das im [DefaultLocale](defaultlocale.md)-Element angegebene Gebietsschema ausgedrückt wird.|

## <a name="remarks"></a>Hinweise

Bei E-Mail-Add-Ins wird das Symbol auf der Benutzeroberfläche **Datei**  >  **Add-Ins verwalten** angezeigt. Bei Inhalts- oder Aufgabenbereich-Add-Ins wird das Symbol in der Benutzeroberfläche **Einfügen**  >  **Add-Ins** angezeigt.

Das Bild muss in einem der folgenden Dateiformate mit einer empfohlenen Auflösung von 64 x 64 Pixel vorliegen: GIF, JPG, PNG, EXIF, BMP oder TIFF. Weitere Informationen finden Sie im Abschnitt _Erstellen einer konsistenten visuellen Identität für Ihre app_ in den [effektiven erstellen-Codebeispielen in Elemente verwenden und innerhalb von Office](https://docs.microsoft.com/office/dev/store/create-effective-office-store-listings).
