# <a name="resources-element"></a><span data-ttu-id="22b06-101">Resources-Element</span><span class="sxs-lookup"><span data-stu-id="22b06-101">Resources element</span></span>

<span data-ttu-id="22b06-p101">Enthält Symbole, Zeichenfolgen und URLs für den [VersionOverrides](versionoverrides.md)-Knoten. Ein Manifestelement gibt eine Ressource mithilfe der **ID** der Ressource an. Dadurch bleibt die Größe des Manifests verwaltbar, insbesondere wenn Ressourcen Versionen für unterschiedliche Gebietsschemas aufweisen. Eine **ID** muss innerhalb des Manifests eindeutig sein und kann maximal 32 Zeichen umfassen.</span><span class="sxs-lookup"><span data-stu-id="22b06-p101">Contains icons, strings, and URLs for the [VersionOverrides](versionoverrides.md) node. A manifest element specifies a resource by using the **id** of the resource. This helps to keep the size of the manifest manageable, especially when resources have versions for different locales. An **id** must be unique within the manifest and can have a maximum of 32 characters.</span></span>

<span data-ttu-id="22b06-106">Jede Ressource kann ein oder mehrere untergeordnete **Override**-Elemente umfassen, um eine unterschiedliche Ressource für ein bestimmtes Gebietsschema zu definieren.</span><span class="sxs-lookup"><span data-stu-id="22b06-106">Each resource can have one or more **Override** child elements to define a different resource for a specific locale.</span></span>

## <a name="child-elements"></a><span data-ttu-id="22b06-107">Untergeordnete Elemente</span><span class="sxs-lookup"><span data-stu-id="22b06-107">Child elements</span></span>

|  <span data-ttu-id="22b06-108">Element</span><span class="sxs-lookup"><span data-stu-id="22b06-108">Element</span></span> |  <span data-ttu-id="22b06-109">Typ</span><span class="sxs-lookup"><span data-stu-id="22b06-109">Type</span></span>  |  <span data-ttu-id="22b06-110">Beschreibung</span><span class="sxs-lookup"><span data-stu-id="22b06-110">Description</span></span>  |
|:-----|:-----|:-----|
|  [<span data-ttu-id="22b06-111">Bilder</span><span class="sxs-lookup"><span data-stu-id="22b06-111">Images</span></span>](#images)            |  <span data-ttu-id="22b06-112">image</span><span class="sxs-lookup"><span data-stu-id="22b06-112">image</span></span>   |  <span data-ttu-id="22b06-113">Stellt die HTTPS-URL des Bilds für ein Symbol bereit.</span><span class="sxs-lookup"><span data-stu-id="22b06-113">Provides the HTTPS URL to an image for an icon.</span></span> |
|  <span data-ttu-id="22b06-114">**URLs**</span><span class="sxs-lookup"><span data-stu-id="22b06-114">**Urls**</span></span>                |  <span data-ttu-id="22b06-115">url</span><span class="sxs-lookup"><span data-stu-id="22b06-115">url</span></span>     |  <span data-ttu-id="22b06-p102">Stellt einen HTTPS-URL-Speicherort bereit. Die maximale Länge einer URL beträgt 2048 Zeichen.</span><span class="sxs-lookup"><span data-stu-id="22b06-p102">Provides an HTTPS URL location. A URL can have a maximum of 2048 characters.</span></span> |
|  <span data-ttu-id="22b06-118">**ShortStrings**</span><span class="sxs-lookup"><span data-stu-id="22b06-118">**ShortStrings**</span></span> |  <span data-ttu-id="22b06-119">string</span><span class="sxs-lookup"><span data-stu-id="22b06-119">string</span></span>  |  <span data-ttu-id="22b06-p103">Der Text für **Label**- und **Title**-Elemente. Jede **Zeichenfolge** enthält maximal 125 Zeichen. </span><span class="sxs-lookup"><span data-stu-id="22b06-p103">The text for **Label** and **Title** elements. Each **String** contains a maximum of 125 characters.</span></span>|
|  <span data-ttu-id="22b06-122">**LongStrings**</span><span class="sxs-lookup"><span data-stu-id="22b06-122">**LongStrings**</span></span>  |  <span data-ttu-id="22b06-123">string</span><span class="sxs-lookup"><span data-stu-id="22b06-123">string</span></span>  | <span data-ttu-id="22b06-p104">Der Text für die **Description**-Attribute. Jede **Zeichenfolge** enthält maximal 250 Zeichen.</span><span class="sxs-lookup"><span data-stu-id="22b06-p104">The text for **Description** attributes. Each **String** contains a maximum of 250 characters.</span></span>|

> [!NOTE]
> <span data-ttu-id="22b06-126">Sie müssen für alle URLs in den Elementen **Bild** und **Url** Secure Sockets Layer (SSL) verwenden.</span><span class="sxs-lookup"><span data-stu-id="22b06-126">You must use Secure Sockets Layer (SSL) for all URLs in the  **Image** and **Url** elements.</span></span>

### <a name="images"></a><span data-ttu-id="22b06-127">Bilder</span><span class="sxs-lookup"><span data-stu-id="22b06-127">Images</span></span>
<span data-ttu-id="22b06-128">Jedes Symbol muss drei  **Images**-Elemente aufweisen, jeweils eines für jede der drei obligatorischen Größen:</span><span class="sxs-lookup"><span data-stu-id="22b06-128">Each icon must have three  **Images** elements, one for each of the three mandatory sizes:</span></span>

- <span data-ttu-id="22b06-129">16x16</span><span class="sxs-lookup"><span data-stu-id="22b06-129">16x16</span></span>
- <span data-ttu-id="22b06-130">32x32</span><span class="sxs-lookup"><span data-stu-id="22b06-130">32x32</span></span>
- <span data-ttu-id="22b06-131">80x80</span><span class="sxs-lookup"><span data-stu-id="22b06-131">80x80</span></span>

<span data-ttu-id="22b06-132">Die folgenden zusätzlichen Größen werden ebenfalls unterstützt, sind jedoch nicht erforderlich:</span><span class="sxs-lookup"><span data-stu-id="22b06-132">The following additional sizes are also supported, but not required:</span></span>

- <span data-ttu-id="22b06-133">20x20</span><span class="sxs-lookup"><span data-stu-id="22b06-133">20x20</span></span>
- <span data-ttu-id="22b06-134">24x24</span><span class="sxs-lookup"><span data-stu-id="22b06-134">24x24</span></span>
- <span data-ttu-id="22b06-135">40x40</span><span class="sxs-lookup"><span data-stu-id="22b06-135">40x40</span></span>
- <span data-ttu-id="22b06-136">48x48</span><span class="sxs-lookup"><span data-stu-id="22b06-136">48x48</span></span>
- <span data-ttu-id="22b06-137">64x64</span><span class="sxs-lookup"><span data-stu-id="22b06-137">64x64</span></span>

> [!IMPORTANT] 
> <span data-ttu-id="22b06-138">Outlook ist die Möglichkeit zum Cache Bildressourcen zur Optimierung der Systemleistung erforderlich.</span><span class="sxs-lookup"><span data-stu-id="22b06-138">Outlook requires the ability to cache image resources for performance purposes.</span></span> <span data-ttu-id="22b06-139">Aus diesem Grund darf der Server, der eine Bildressource hostet, keine CACHE-CONTROL-Richtlinien zum Antwortheader hinzufügen.</span><span class="sxs-lookup"><span data-stu-id="22b06-139">For this reason, the server hosting an image resource must not add any CACHE-CONTROL directives to the response header.</span></span> <span data-ttu-id="22b06-140">Dies würde dazu führen, dass Outlook automatisch ein allgemeines oder Standardbild ersetzt.</span><span class="sxs-lookup"><span data-stu-id="22b06-140">This will result in Outlook automatically substituting a generic or default image.</span></span>    

## <a name="resources-examples"></a><span data-ttu-id="22b06-141">Beispiele für Ressourcen</span><span class="sxs-lookup"><span data-stu-id="22b06-141">Resources examples</span></span> 

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
