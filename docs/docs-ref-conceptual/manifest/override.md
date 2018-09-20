# <a name="override-element"></a><span data-ttu-id="30c59-101">Override-Element</span><span class="sxs-lookup"><span data-stu-id="30c59-101">Override element</span></span>

<span data-ttu-id="30c59-102">Stellt eine Methode zum Angeben des Werts einer Einstellung für ein zusätzliches Gebietsschema bereit.</span><span class="sxs-lookup"><span data-stu-id="30c59-102">Provides a way to specify the value of a setting for an additional locale.</span></span>

<span data-ttu-id="30c59-103">**Add-In-Typ:** Inhalt, Aufgabenbereich, E-Mail</span><span class="sxs-lookup"><span data-stu-id="30c59-103">**Add-in type:** Content, Task pane, Mail</span></span>

## <a name="syntax"></a><span data-ttu-id="30c59-104">Syntax</span><span class="sxs-lookup"><span data-stu-id="30c59-104">Syntax</span></span>

```XML
<Override Locale="string" Value="string" />
```

## <a name="contained-in"></a><span data-ttu-id="30c59-105">Enthalten in</span><span class="sxs-lookup"><span data-stu-id="30c59-105">Contained in</span></span>

|<span data-ttu-id="30c59-106">**Element**</span><span class="sxs-lookup"><span data-stu-id="30c59-106">**Element**</span></span>|
|:-----|
|[<span data-ttu-id="30c59-107">CitationText</span><span class="sxs-lookup"><span data-stu-id="30c59-107">CitationText</span></span>](citationtext.md)|
|[<span data-ttu-id="30c59-108">Beschreibung</span><span class="sxs-lookup"><span data-stu-id="30c59-108">Description</span></span>](description.md)|
|[<span data-ttu-id="30c59-109">DictionaryName</span><span class="sxs-lookup"><span data-stu-id="30c59-109">DictionaryName</span></span>](dictionaryname.md)|
|[<span data-ttu-id="30c59-110">DictionaryHomePage</span><span class="sxs-lookup"><span data-stu-id="30c59-110">DictionaryHomePage</span></span>](dictionaryhomepage.md)|
|[<span data-ttu-id="30c59-111">Anzeigename</span><span class="sxs-lookup"><span data-stu-id="30c59-111">DisplayName</span></span>](displayname.md)|
|[<span data-ttu-id="30c59-112">HighResolutionIconUrl</span><span class="sxs-lookup"><span data-stu-id="30c59-112">HighResolutionIconUrl</span></span>](highresolutioniconurl.md)|
|[<span data-ttu-id="30c59-113">IconUrl</span><span class="sxs-lookup"><span data-stu-id="30c59-113">IconUrl</span></span>](iconurl.md)|
|[<span data-ttu-id="30c59-114">QueryUri</span><span class="sxs-lookup"><span data-stu-id="30c59-114">QueryUri</span></span>](queryuri.md)|
|[<span data-ttu-id="30c59-115">SourceLocation</span><span class="sxs-lookup"><span data-stu-id="30c59-115">SourceLocation</span></span>](sourcelocation.md)|
|[<span data-ttu-id="30c59-116">SupportUrl</span><span class="sxs-lookup"><span data-stu-id="30c59-116">SupportUrl</span></span>](supporturl.md)|

## <a name="attributes"></a><span data-ttu-id="30c59-117">Attribute</span><span class="sxs-lookup"><span data-stu-id="30c59-117">Attributes</span></span>

|<span data-ttu-id="30c59-118">**Attribut**</span><span class="sxs-lookup"><span data-stu-id="30c59-118">**Attribute**</span></span>|<span data-ttu-id="30c59-119">**Typ**</span><span class="sxs-lookup"><span data-stu-id="30c59-119">**Type**</span></span>|<span data-ttu-id="30c59-120">**Erforderlich**</span><span class="sxs-lookup"><span data-stu-id="30c59-120">**Required**</span></span>|<span data-ttu-id="30c59-121">**Beschreibung**</span><span class="sxs-lookup"><span data-stu-id="30c59-121">**Description**</span></span>|
|:-----|:-----|:-----|:-----|
|<span data-ttu-id="30c59-122">Locale</span><span class="sxs-lookup"><span data-stu-id="30c59-122">Locale</span></span>|<span data-ttu-id="30c59-123">string</span><span class="sxs-lookup"><span data-stu-id="30c59-123">string</span></span>|<span data-ttu-id="30c59-124">erforderlich</span><span class="sxs-lookup"><span data-stu-id="30c59-124">required</span></span>|<span data-ttu-id="30c59-125">Gibt den Kulturnamen des Gebietsschemas für diese Außerkraftsetzung im BCP 47-Sprachtagformat an, z. B.  `"en-US"`.</span><span class="sxs-lookup"><span data-stu-id="30c59-125">Specifies the culture name of the locale for this override in the BCP 47 language tag format, such as  `"en-US"`.</span></span>|
|<span data-ttu-id="30c59-126">Wert</span><span class="sxs-lookup"><span data-stu-id="30c59-126">Value</span></span>|<span data-ttu-id="30c59-127">string</span><span class="sxs-lookup"><span data-stu-id="30c59-127">string</span></span>|<span data-ttu-id="30c59-128">erforderlich</span><span class="sxs-lookup"><span data-stu-id="30c59-128">required</span></span>|<span data-ttu-id="30c59-129">Gibt den Wert der für das angegebene Gebietsschema ausgedrückten Einstellung an.</span><span class="sxs-lookup"><span data-stu-id="30c59-129">Specifies value of the setting expressed for the specified locale.</span></span>|

## <a name="see-also"></a><span data-ttu-id="30c59-130">Siehe auch</span><span class="sxs-lookup"><span data-stu-id="30c59-130">See also</span></span>

- [<span data-ttu-id="30c59-131">Lokalisierung für Office-Add-Ins</span><span class="sxs-lookup"><span data-stu-id="30c59-131">Localization for Office Add-ins</span></span>](https://docs.microsoft.com/office/dev/add-ins/develop/localization)
    
