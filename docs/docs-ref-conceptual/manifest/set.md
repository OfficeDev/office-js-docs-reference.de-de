# <a name="set-element"></a><span data-ttu-id="6320d-101">Set-Element</span><span class="sxs-lookup"><span data-stu-id="6320d-101">Set element</span></span>

<span data-ttu-id="6320d-102">Gibt einen Anforderungssatz aus der JavaScript-API für Office an, die Ihr Office-Add-In zum Aktivieren benötigt.</span><span class="sxs-lookup"><span data-stu-id="6320d-102">Specifies a requirement set from the JavaScript API for Office that your Office Add-in requires to activate.</span></span>

<span data-ttu-id="6320d-103">**Add-In-Typ:** Inhalt, Aufgabenbereich, E-Mail</span><span class="sxs-lookup"><span data-stu-id="6320d-103">**Add-in type:** Content, Task pane, Mail</span></span>

## <a name="syntax"></a><span data-ttu-id="6320d-104">Syntax</span><span class="sxs-lookup"><span data-stu-id="6320d-104">Syntax</span></span>

```XML
<Set Name="string" MinVersion="n .n">
```

## <a name="contained-in"></a><span data-ttu-id="6320d-105">Enthalten in</span><span class="sxs-lookup"><span data-stu-id="6320d-105">Contained in</span></span>

[<span data-ttu-id="6320d-106">Sätzen</span><span class="sxs-lookup"><span data-stu-id="6320d-106">Sets</span></span>](sets.md)

## <a name="attributes"></a><span data-ttu-id="6320d-107">Attribute</span><span class="sxs-lookup"><span data-stu-id="6320d-107">Attributes</span></span>

|<span data-ttu-id="6320d-108">**Attribut**</span><span class="sxs-lookup"><span data-stu-id="6320d-108">**Attribute**</span></span>|<span data-ttu-id="6320d-109">**Typ**</span><span class="sxs-lookup"><span data-stu-id="6320d-109">**Type**</span></span>|<span data-ttu-id="6320d-110">**Erforderlich**</span><span class="sxs-lookup"><span data-stu-id="6320d-110">**Required**</span></span>|<span data-ttu-id="6320d-111">**Beschreibung**</span><span class="sxs-lookup"><span data-stu-id="6320d-111">**Description**</span></span>|
|:-----|:-----|:-----|:-----|
|<span data-ttu-id="6320d-112">Name</span><span class="sxs-lookup"><span data-stu-id="6320d-112">Name</span></span>|<span data-ttu-id="6320d-113">string</span><span class="sxs-lookup"><span data-stu-id="6320d-113">string</span></span>|<span data-ttu-id="6320d-114">erforderlich</span><span class="sxs-lookup"><span data-stu-id="6320d-114">required</span></span>|<span data-ttu-id="6320d-115">Der Name eines [Anforderungssatzes](https://docs.microsoft.com/office/dev/add-ins/develop/office-versions-and-requirement-sets).</span><span class="sxs-lookup"><span data-stu-id="6320d-115">The name of a [requirement set](https://docs.microsoft.com/office/dev/add-ins/develop/office-versions-and-requirement-sets).</span></span>|
|<span data-ttu-id="6320d-116">MinVersion</span><span class="sxs-lookup"><span data-stu-id="6320d-116">MinVersion</span></span>|<span data-ttu-id="6320d-117">string</span><span class="sxs-lookup"><span data-stu-id="6320d-117">string</span></span>|<span data-ttu-id="6320d-118">Optional</span><span class="sxs-lookup"><span data-stu-id="6320d-118">optional</span></span>|<span data-ttu-id="6320d-p101">Gibt die Mindestversion des API-Satzes an, die für das Add-In erforderlich ist. Setzt den Wert von **DefaultMinVersion** außer Kraft, wenn dies im übergeordneten [Sets](sets.md)-Element angegeben ist.</span><span class="sxs-lookup"><span data-stu-id="6320d-p101">Specifies the minimum version of the API set required by your add-in. Overrides the value of  **DefaultMinVersion**, if it is specified in the parent [Sets](sets.md) element.</span></span>|

## <a name="remarks"></a><span data-ttu-id="6320d-121">Hinweise</span><span class="sxs-lookup"><span data-stu-id="6320d-121">Remarks</span></span>

<span data-ttu-id="6320d-122">Weitere Informationen zu anforderungssätze finden Sie unter [Office-Versionen und Anforderung festgelegt](https://docs.microsoft.com/office/dev/add-ins/develop/office-versions-and-requirement-sets).</span><span class="sxs-lookup"><span data-stu-id="6320d-122">For more information about requirement sets, see [Office versions and requirement sets](https://docs.microsoft.com/office/dev/add-ins/develop/office-versions-and-requirement-sets).</span></span>

<span data-ttu-id="6320d-123">Weitere Informationen zum **MinVersion**-Attribut des **Set**-Elements und zum **DefaultMinVersion**-Attribut des **Sets**-Elements finden Sie unter [Festlegen des Requirements-Element im Manifest](https://docs.microsoft.com/office/dev/add-ins/develop/specify-office-hosts-and-api-requirements#set-the-requirements-element-in-the-manifest).</span><span class="sxs-lookup"><span data-stu-id="6320d-123">For more information about the  **MinVersion** attribute of the **Set** element and the **DefaultMinVersion** attribute of the **Sets** element, see [Set the Requirements element in the manifest](https://docs.microsoft.com/office/dev/add-ins/develop/specify-office-hosts-and-api-requirements#set-the-requirements-element-in-the-manifest).</span></span>

> [!IMPORTANT] 
> <span data-ttu-id="6320d-124">Für Mail-add-ins, ist nur eine `"Mailbox"` Anforderung verfügbar festgelegt.</span><span class="sxs-lookup"><span data-stu-id="6320d-124">For mail add-ins, there is only one  `"Mailbox"` requirement set available.</span></span> <span data-ttu-id="6320d-125">Diese Anforderungssatz enthält die gesamte Teilmenge der API-Unterstützung in Mail-add-ins für Outlook, und geben Sie die `"Mailbox"` Anforderung in Ihrer Mail-add-in des Manifests festgelegt (es ist nicht optional, wie die Groß-/Kleinschreibung für Inhalts- und Aufgabenbereich Bereich-add-ins ist).</span><span class="sxs-lookup"><span data-stu-id="6320d-125">This requirement set contains the entire subset of API supported in mail add-ins for Outlook, and you must specify the `"Mailbox"` requirement set in your mail add-in's manifest (it's not optional as is the case for content and task pane add-ins).</span></span> <span data-ttu-id="6320d-126">Darüber hinaus können Sie Unterstützung für bestimmte Methoden in Mail-add-ins nicht deklarieren.</span><span class="sxs-lookup"><span data-stu-id="6320d-126">Also, you can't declare support for specific methods in mail add-ins.</span></span>
