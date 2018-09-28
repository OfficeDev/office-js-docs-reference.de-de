# <a name="icon-element"></a><span data-ttu-id="1df35-101">Icon-Element</span><span class="sxs-lookup"><span data-stu-id="1df35-101">Icon element</span></span>

<span data-ttu-id="1df35-102">Definiert **Image**-Elemente für [Button](control.md#button-control)- oder [Menu](control.md#menu-dropdown-button-controls)-Steuerelemente.</span><span class="sxs-lookup"><span data-stu-id="1df35-102">Defines **Image** elements for [Button](control.md#button-control) or [Menu](control.md#menu-dropdown-button-controls) controls.</span></span>

## <a name="attributes"></a><span data-ttu-id="1df35-103">Attribute</span><span class="sxs-lookup"><span data-stu-id="1df35-103">Attributes</span></span>

|  <span data-ttu-id="1df35-104">Attribut</span><span class="sxs-lookup"><span data-stu-id="1df35-104">Attribute</span></span>  |  <span data-ttu-id="1df35-105">Erforderlich</span><span class="sxs-lookup"><span data-stu-id="1df35-105">Required</span></span>  |  <span data-ttu-id="1df35-106">Beschreibung</span><span class="sxs-lookup"><span data-stu-id="1df35-106">Description</span></span>  |
|:-----|:-----|:-----|
|  <span data-ttu-id="1df35-107">**xsi:type**</span><span class="sxs-lookup"><span data-stu-id="1df35-107">**xsi:type**</span></span>  |  <span data-ttu-id="1df35-108">Nein</span><span class="sxs-lookup"><span data-stu-id="1df35-108">No</span></span>  | <span data-ttu-id="1df35-p101">Der Typ des Symbols, das definiert wird. Dies gilt nur für Symbole in mobilen Formularfaktoren. Bei **Icon**-Elementen, die in einem [MobileFormFactor](mobileformfactor.md)-Element enthalten sind, muss dieses Attribut auf `bt:MobileIconList` festgelegt sein.</span><span class="sxs-lookup"><span data-stu-id="1df35-p101">The type of icon being defined. This is only applicable to icons in mobile form factors. **Icon** elements contained within a [MobileFormFactor](mobileformfactor.md) element must have this attribute set to `bt:MobileIconList`.</span></span> |

## <a name="child-elements"></a><span data-ttu-id="1df35-112">Untergeordnete Elemente</span><span class="sxs-lookup"><span data-stu-id="1df35-112">Child elements</span></span>

|  <span data-ttu-id="1df35-113">Element</span><span class="sxs-lookup"><span data-stu-id="1df35-113">Element</span></span> |  <span data-ttu-id="1df35-114">Erforderlich</span><span class="sxs-lookup"><span data-stu-id="1df35-114">Required</span></span>  |  <span data-ttu-id="1df35-115">Beschreibung</span><span class="sxs-lookup"><span data-stu-id="1df35-115">Description</span></span>  |
|:-----|:-----|:-----|
|  [<span data-ttu-id="1df35-116">Image</span><span class="sxs-lookup"><span data-stu-id="1df35-116">Image</span></span>](#image)        | <span data-ttu-id="1df35-117">Ja</span><span class="sxs-lookup"><span data-stu-id="1df35-117">Yes</span></span> |   <span data-ttu-id="1df35-118">resid-Element eines zu verwendenden Bilds</span><span class="sxs-lookup"><span data-stu-id="1df35-118">resid of an image to use</span></span>         |

### <a name="image"></a><span data-ttu-id="1df35-119">Bild</span><span class="sxs-lookup"><span data-stu-id="1df35-119">Image</span></span>

<span data-ttu-id="1df35-p102">Ein Bild für die Schaltfläche. Das Attribut  **resid** muss auf einen Wert des **id**-Attributs eines  **Image**-Elements im  **Images**-Element im  [Resources](resources.md) -Element festgelegt werden. Das Attribut **size** gibt die Größe des Bilds in Pixel an. Drei Bildgrößen sind erforderlich (16, 32 und 80 Pixel), es werden aber fünf weitere Größen unterstützt (20, 24, 40, 48 und 64 Pixel).|</span><span class="sxs-lookup"><span data-stu-id="1df35-p102">An image for the button. The  **resid** attribute must be set to the value of the **id** attribute of an **Image** element in the **Images** element in the [Resources](resources.md) element. The **size** attribute indicates the size in pixels of the image. Three image sizes are required (16, 32, and 80 pixels) while five other sizes are supported (20, 24, 40, 48, and 64 pixels).|</span></span>

```xml
<Icon>
  <bt:Image size="16" resid="blue-icon-16" />
  <bt:Image size="32" resid="blue-icon-32" />
  <bt:Image size="80" resid="blue-icon-80" />
</Icon>
```

## <a name="additional-requirements-for-mobile-form-factors"></a><span data-ttu-id="1df35-124">Weitere Anforderungen für mobile Formularfaktoren</span><span class="sxs-lookup"><span data-stu-id="1df35-124">Additional requirements for mobile form factors</span></span>

<span data-ttu-id="1df35-p103">Wenn das übergeordnete **Icon**-Element ist ein Nachfolger eines [MobileFormFactor](mobileformfactor.md)-Elements ist, sind die mindestens erforderlichen Größen geringfügig anders. Das Manifest muss mindestens 25, 32 und 48 Pixelgrößen bereitstellen. Jede bereitgestellte Größe muss mindestens dreimal angezeigt werden, wobei das `scale`-Attribut auf `1`, `2` oder `3` festgelegt ist.</span><span class="sxs-lookup"><span data-stu-id="1df35-p103">When the parent **Icon** element is a descendant of a [MobileFormFactor](mobileformfactor.md) element, the minimum required sizes are slightly different. The manifest must minimally provide 25, 32, and 48 pixel sizes. Each size provided must appear three times, with a `scale` attribute set to `1`, `2`, or `3`.</span></span>

```xml
<Icon xsi:type="bt:MobileIconList">
  <bt:Image resid="blue-icon-16-1" size="25" scale="1" />
  <bt:Image resid="blue-icon-16-2" size="25" scale="2" />
  <bt:Image resid="blue-icon-16-3" size="25" scale="3" />
  <bt:Image resid="blue-icon-32-1" size="32" scale="1" />
  <bt:Image resid="blue-icon-32-2" size="32" scale="2" />
  <bt:Image resid="blue-icon-32-3" size="32" scale="3" />
  <bt:Image resid="blue-icon-80-1" size="48" scale="1" />
  <bt:Image resid="blue-icon-80-2" size="48" scale="2" />
  <bt:Image resid="blue-icon-80-3" size="48" scale="3" />
</Icon>
```