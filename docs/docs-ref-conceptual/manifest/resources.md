# <a name="resources-element"></a>Resources-Element

Enthält Symbole, Zeichenfolgen und URLs für den [VersionOverrides](versionoverrides.md)-Knoten. Ein Manifestelement gibt eine Ressource mithilfe der **ID** der Ressource an. Dadurch bleibt die Größe des Manifests verwaltbar, insbesondere wenn Ressourcen Versionen für unterschiedliche Gebietsschemas aufweisen. Eine **ID** muss innerhalb des Manifests eindeutig sein und kann maximal 32 Zeichen umfassen.

Jede Ressource kann ein oder mehrere untergeordnete **Override**-Elemente umfassen, um eine unterschiedliche Ressource für ein bestimmtes Gebietsschema zu definieren.

## <a name="child-elements"></a>Untergeordnete Elemente

|  Element |  Typ  |  Beschreibung  |
|:-----|:-----|:-----|
|  [Bilder](#images)            |  image   |  Stellt die HTTPS-URL des Bilds für ein Symbol bereit. |
|  **URLs**                |  url     |  Stellt einen HTTPS-URL-Speicherort bereit. Die maximale Länge einer URL beträgt 2048 Zeichen. |
|  **ShortStrings** |  string  |  Der Text für **Label**- und **Title**-Elemente. Jede **Zeichenfolge** enthält maximal 125 Zeichen. |
|  **LongStrings**  |  string  | Der Text für die **Description**-Attribute. Jede **Zeichenfolge** enthält maximal 250 Zeichen.|

> [!NOTE]
> Sie müssen für alle URLs in den Elementen **Bild** und **Url** Secure Sockets Layer (SSL) verwenden.

### <a name="images"></a>Bilder
Jedes Symbol muss drei  **Images**-Elemente aufweisen, jeweils eines für jede der drei obligatorischen Größen:

- 16x16
- 32x32
- 80x80

Die folgenden zusätzlichen Größen werden ebenfalls unterstützt, sind jedoch nicht erforderlich:

- 20x20
- 24x24
- 40x40
- 48x48
- 64x64

> [!IMPORTANT] 
> Outlook ist die Möglichkeit zum Cache Bildressourcen zur Optimierung der Systemleistung erforderlich. Aus diesem Grund darf der Server, der eine Bildressource hostet, keine CACHE-CONTROL-Richtlinien zum Antwortheader hinzufügen. Dies würde dazu führen, dass Outlook automatisch ein allgemeines oder Standardbild ersetzt.    

## <a name="resources-examples"></a>Beispiele für Ressourcen 

```XML
<Resources>
      <bt:Images>
        <bt:Image id="icon1_16x16" DefaultValue="https://www.contoso.com/icon_default.png">
          <bt:Override Locale="ja-jp" Value="https://www.contoso.com/ja-jp16-icon_default.png" />
        </bt:Image>
        <bt:Image id="icon1_32x32" DefaultValue="https://www.contoso.com/icon_default.png">
          <bt:Override Locale="ja-jp" Value="https://www.contoso.com/ja-jp32-icon_default.png" />
        </bt:Image>
        <bt:Image id="icon1_80x80" DefaultValue="https://www.contoso.com/icon_default.png">
          <bt:Override Locale="ja-jp" Value="https://www.contoso.com/ja-jp80-icon_default.png" />
        </bt:Image>
      </bt:Images>
      <bt:Urls>
        <bt:Url id="residDesktopFuncUrl" DefaultValue="https://www.contoso.com/Pages/Home.aspx">
          <bt:Override Locale="ja-jp" Value="https://www.contoso.com/Pages/Home.aspx" />
        </bt:Url>
      </bt:Urls>
      <bt:ShortStrings>
        <bt:String id="residLabel" DefaultValue="GetData">
          <bt:Override Locale="ja-jp" Value="JA-JP-GetData" />
        </bt:String>
      </bt:ShortStrings>
      <bt:LongStrings>
        <bt:String id="residToolTip" DefaultValue="Get data for your document.">
          <bt:Override Locale="ja-jp" Value="JA-JP - Get data for your document." />
        </bt:String>
      </bt:LongStrings>
    </Resources>
```

```xml
<Resources>
  <bt:Images>
    <!-- Blue icon -->
    <bt:Image id="blue-icon-16" DefaultValue="YOUR_WEB_SERVER/blue-16.png"/>
    <bt:Image id="blue-icon-32" DefaultValue="YOUR_WEB_SERVER//blue-32.png"/>
    <bt:Image id="blue-icon-80" DefaultValue="YOUR_WEB_SERVER/blue-80.png"/>
  </bt:Images>
  <bt:Urls>
    <bt:Url id="functionFile" DefaultValue="YOUR_WEB_SERVER/FunctionFile/Functions.html"/>
    <!-- other URLs -->
  </bt:Urls>
  <bt:ShortStrings>
    <bt:String id="groupLabel" DefaultValue="Add-in Demo">
      <bt:Override Locale="ar-sa" Value="<Localized text>" />
    </bt:String>
    <!-- Other short strings -->
  </bt:ShortStrings>
  <bt:LongStrings>
    <bt:String id="funcReadSuperTipDescription" DefaultValue="Gets the subject of the message or appointment.">
      <bt:Override Locale="ar-sa" Value="<Localized text>." />
    </bt:String>
    <!-- Other long strings -->
  </bt:LongStrings>
</Resources>
```
