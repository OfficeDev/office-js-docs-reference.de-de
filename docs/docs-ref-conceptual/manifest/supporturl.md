# <a name="supporturl-element"></a><span data-ttu-id="2932a-101">SupportUrl-Element</span><span class="sxs-lookup"><span data-stu-id="2932a-101">SupportUrl element</span></span>

<span data-ttu-id="2932a-102">Gibt die URL einer Seite an, die Supportinformationen für Ihr Add-In enthält.</span><span class="sxs-lookup"><span data-stu-id="2932a-102">Specifies the URL of a page that provides support information for your add-in.</span></span>

## <a name="syntax"></a><span data-ttu-id="2932a-103">Syntax</span><span class="sxs-lookup"><span data-stu-id="2932a-103">Syntax</span></span>

```XML
<OfficeApp>
...
  <IconUrl DefaultValue="https://contoso.com/assets/icon-32.png" />
  <HighResolutionIconUrl DefaultValue="https://contoso.com/assets/hi-res-icon.png"/>
  
  
  <SupportUrl DefaultValue="https://contoso.com/support " />
  
  
  <AppDomains>
  ...
  </AppDomains>
...
</OfficeApp>
```

## <a name="contained-in"></a><span data-ttu-id="2932a-104">Enthalten in</span><span class="sxs-lookup"><span data-stu-id="2932a-104">Contained in</span></span>

[<span data-ttu-id="2932a-105">OfficeApp</span><span class="sxs-lookup"><span data-stu-id="2932a-105">OfficeApp</span></span>](officeapp.md)

## <a name="can-contain"></a><span data-ttu-id="2932a-106">Kann enthalten</span><span class="sxs-lookup"><span data-stu-id="2932a-106">Can contain</span></span>

|  <span data-ttu-id="2932a-107">Element</span><span class="sxs-lookup"><span data-stu-id="2932a-107">Element</span></span> | <span data-ttu-id="2932a-108">Erforderlich</span><span class="sxs-lookup"><span data-stu-id="2932a-108">Required</span></span> | <span data-ttu-id="2932a-109">Beschreibung</span><span class="sxs-lookup"><span data-stu-id="2932a-109">Description</span></span>  |
|:-----|:-----|:-----|
|  [<span data-ttu-id="2932a-110">Override</span><span class="sxs-lookup"><span data-stu-id="2932a-110">Override</span></span>](override.md)   | <span data-ttu-id="2932a-111">Nein</span><span class="sxs-lookup"><span data-stu-id="2932a-111">No</span></span> | <span data-ttu-id="2932a-112">Gibt die Einstellung für zusätzliche Gebietsschema-URLS an.</span><span class="sxs-lookup"><span data-stu-id="2932a-112">Specifies the setting for additional locale urls</span></span> |

## <a name="attributes"></a><span data-ttu-id="2932a-113">Attribute</span><span class="sxs-lookup"><span data-stu-id="2932a-113">Attributes</span></span>

|<span data-ttu-id="2932a-114">**Attribut**</span><span class="sxs-lookup"><span data-stu-id="2932a-114">**Attribute**</span></span>|<span data-ttu-id="2932a-115">**Typ**</span><span class="sxs-lookup"><span data-stu-id="2932a-115">**Type**</span></span>|<span data-ttu-id="2932a-116">**Erforderlich**</span><span class="sxs-lookup"><span data-stu-id="2932a-116">**Required**</span></span>|<span data-ttu-id="2932a-117">**Beschreibung**</span><span class="sxs-lookup"><span data-stu-id="2932a-117">**Description**</span></span>|
|:-----|:-----|:-----|:-----|
|<span data-ttu-id="2932a-118">DefaultValue</span><span class="sxs-lookup"><span data-stu-id="2932a-118">DefaultValue</span></span>|<span data-ttu-id="2932a-119">URL</span><span class="sxs-lookup"><span data-stu-id="2932a-119">URL</span></span>|<span data-ttu-id="2932a-120">erforderlich</span><span class="sxs-lookup"><span data-stu-id="2932a-120">required</span></span>|<span data-ttu-id="2932a-121">Gibt den Standardwert für diese Einstellung an, der für das im [DefaultLocale](defaultlocale.md)-Element angegebene Gebietsschema ausgedrückt wird.</span><span class="sxs-lookup"><span data-stu-id="2932a-121">Specifies the default value for this setting, expressed for the locale specified in the [DefaultLocale](defaultlocale.md) element.</span></span>|
