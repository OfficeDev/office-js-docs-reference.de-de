# <a name="highresolutioniconurl-element"></a><span data-ttu-id="35c01-101">HighResolutionIconUrl-Element</span><span class="sxs-lookup"><span data-stu-id="35c01-101">HighResolutionIconUrl element</span></span>

<span data-ttu-id="35c01-102">Gibt die URL des Bilds an, das zur Darstellung Ihres Office-Add-Ins in der Einfüge-UX und im Office Store auf Bildschirmen mit hohem DPI-Wert verwendet wird.</span><span class="sxs-lookup"><span data-stu-id="35c01-102">Specifies the URL of the image that is used to represent your Office Add-in in the insertion UX and Office Store on high DPI screens.</span></span>

<span data-ttu-id="35c01-103">**Add-In-Typ:** Inhalt, Aufgabenbereich, E-Mail</span><span class="sxs-lookup"><span data-stu-id="35c01-103">**Add-in type:** Content, Task pane, Mail</span></span>

## <a name="syntax"></a><span data-ttu-id="35c01-104">Syntax</span><span class="sxs-lookup"><span data-stu-id="35c01-104">Syntax</span></span>

```XML
<HighResolutionIconUrl DefaultValue="string" />
```

## <a name="can-contain"></a><span data-ttu-id="35c01-105">Kann enthalten</span><span class="sxs-lookup"><span data-stu-id="35c01-105">Can contain</span></span>

[<span data-ttu-id="35c01-106">Override</span><span class="sxs-lookup"><span data-stu-id="35c01-106">Override</span></span>](override.md)

## <a name="attributes"></a><span data-ttu-id="35c01-107">Attribute</span><span class="sxs-lookup"><span data-stu-id="35c01-107">Attributes</span></span>

|<span data-ttu-id="35c01-108">**Attribut**</span><span class="sxs-lookup"><span data-stu-id="35c01-108">**Attribute**</span></span>|<span data-ttu-id="35c01-109">**Typ**</span><span class="sxs-lookup"><span data-stu-id="35c01-109">**Type**</span></span>|<span data-ttu-id="35c01-110">**Erforderlich**</span><span class="sxs-lookup"><span data-stu-id="35c01-110">**Required**</span></span>|<span data-ttu-id="35c01-111">**Beschreibung**</span><span class="sxs-lookup"><span data-stu-id="35c01-111">**Description**</span></span>|
|:-----|:-----|:-----|:-----|
|<span data-ttu-id="35c01-112">DefaultValue</span><span class="sxs-lookup"><span data-stu-id="35c01-112">DefaultValue</span></span>|<span data-ttu-id="35c01-113">Zeichenfolge (URL)</span><span class="sxs-lookup"><span data-stu-id="35c01-113">string (URL)</span></span>|<span data-ttu-id="35c01-114">erforderlich</span><span class="sxs-lookup"><span data-stu-id="35c01-114">required</span></span>|<span data-ttu-id="35c01-115">Gibt den Standardwert für diese Einstellung an, der für das im [DefaultLocale](defaultlocale.md)-Element angegebene Gebietsschema ausgedrückt wird.</span><span class="sxs-lookup"><span data-stu-id="35c01-115">Specifies the default value for this setting, expressed for the locale specified in the [DefaultLocale](defaultlocale.md) element.</span></span>|

## <a name="remarks"></a><span data-ttu-id="35c01-116">Hinweise</span><span class="sxs-lookup"><span data-stu-id="35c01-116">Remarks</span></span>

<span data-ttu-id="35c01-p101">Bei E-Mail-Add-Ins wird das Symbol auf der Benutzeroberfläche **Datei**  >  **Add-Ins verwalten** angezeigt. Bei Inhalts- oder Aufgabenbereich-Add-Ins wird das Symbol in der Benutzeroberfläche **Einfügen**  >  **Add-Ins** angezeigt.</span><span class="sxs-lookup"><span data-stu-id="35c01-p101">For a mail add-in, the icon is displayed in the  **File** > **Manage add-ins** UI . For a content or task pane add-in, the icon is displayed in the **Insert** > **Add-ins** UI.</span></span>

<span data-ttu-id="35c01-119">Das Bild muss in einem der folgenden Dateiformate mit einer empfohlenen Auflösung von 64 x 64 Pixel vorliegen: GIF, JPG, PNG, EXIF, BMP oder TIFF.</span><span class="sxs-lookup"><span data-stu-id="35c01-119">The image must be in one of the following file formats at a recommended resolution of 64 x 64 pixels: GIF, JPG, PNG, EXIF, BMP or TIFF.</span></span> <span data-ttu-id="35c01-120">Weitere Informationen finden Sie im Abschnitt _Erstellen einer konsistenten visuellen Identität für Ihre app_ in den [effektiven erstellen-Codebeispielen in Elemente verwenden und innerhalb von Office](https://docs.microsoft.com/office/dev/store/create-effective-office-store-listings).</span><span class="sxs-lookup"><span data-stu-id="35c01-120">For more information, see the section  _Create a consistent visual identity for your app_ in [Create effective listings in AppSource and within Office](https://docs.microsoft.com/office/dev/store/create-effective-office-store-listings).</span></span>
