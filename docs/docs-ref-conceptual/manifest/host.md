# <a name="host-element"></a><span data-ttu-id="16792-101">Host-Element</span><span class="sxs-lookup"><span data-stu-id="16792-101">Host element</span></span>

<span data-ttu-id="16792-102">Gibt einen einzelnen Office-Anwendungstyp an, in dem das Add-In aktiviert werden sollte.</span><span class="sxs-lookup"><span data-stu-id="16792-102">Specifies an individual Office application type where the add-in should activate.</span></span>

> [!IMPORTANT] 
> <span data-ttu-id="16792-103">Die Syntax der **Host** -Element hängt davon ab, ob das Element innerhalb des [grundlegenden Manifest](#basic-manifest) oder in den [VersionOverrides](#versionoverrides-node) -Knoten definiert ist.</span><span class="sxs-lookup"><span data-stu-id="16792-103">The **Host** element syntax varies depending on whether the element is defined within the [basic manifest](#basic-manifest) or within the [VersionOverrides](#versionoverrides-node) node.</span></span> <span data-ttu-id="16792-104">Die Funktionalität ist jedoch dieselbe.</span><span class="sxs-lookup"><span data-stu-id="16792-104">However, the functionality is the same.</span></span>  

## <a name="basic-manifest"></a><span data-ttu-id="16792-105">Grundlegendes Manifest</span><span class="sxs-lookup"><span data-stu-id="16792-105">Basic manifest</span></span>

<span data-ttu-id="16792-106">Wenn dies im grundlegenden Manifest (unter [OfficeApp](officeapp.md)) definiert ist, wird der Hosttyp vom `Name`-Attribut bestimmt.</span><span class="sxs-lookup"><span data-stu-id="16792-106">When defined in the basic manifest (under [OfficeApp](officeapp.md)), the host type is determined by the `Name` attribute.</span></span>   

### <a name="attributes"></a><span data-ttu-id="16792-107">Attribute</span><span class="sxs-lookup"><span data-stu-id="16792-107">Attributes</span></span>

| <span data-ttu-id="16792-108">Attribut</span><span class="sxs-lookup"><span data-stu-id="16792-108">Attribute</span></span>     | <span data-ttu-id="16792-109">Typ</span><span class="sxs-lookup"><span data-stu-id="16792-109">Type</span></span>   | <span data-ttu-id="16792-110">Erforderlich</span><span class="sxs-lookup"><span data-stu-id="16792-110">Required</span></span> | <span data-ttu-id="16792-111">Beschreibung</span><span class="sxs-lookup"><span data-stu-id="16792-111">Description</span></span>                                      |
|:--------------|:-------|:---------|:-------------------------------------------------|
| [<span data-ttu-id="16792-112">Name</span><span class="sxs-lookup"><span data-stu-id="16792-112">Name</span></span>](#name) | <span data-ttu-id="16792-113">string</span><span class="sxs-lookup"><span data-stu-id="16792-113">string</span></span> | <span data-ttu-id="16792-114">erforderlich</span><span class="sxs-lookup"><span data-stu-id="16792-114">required</span></span> | <span data-ttu-id="16792-115">Der Name des Typs der Office-Hostanwendung.</span><span class="sxs-lookup"><span data-stu-id="16792-115">The name of the type of Office host application.</span></span> |

### <a name="name"></a><span data-ttu-id="16792-116">Name</span><span class="sxs-lookup"><span data-stu-id="16792-116">Name</span></span>
<span data-ttu-id="16792-p102">Gibt den Hosttyp an, auf den von diesem Add-In abgezielt wird. Bei dem Wert muss es sich um Folgendes handeln:</span><span class="sxs-lookup"><span data-stu-id="16792-p102">Specifies the Host type targeted by this add-in. The value must be one of the following:</span></span>

- <span data-ttu-id="16792-119">`Document` (Word)</span><span class="sxs-lookup"><span data-stu-id="16792-119">`Document` (Word)</span></span>
- <span data-ttu-id="16792-120">`Database` (Access)</span><span class="sxs-lookup"><span data-stu-id="16792-120">`Database` (Access)</span></span>
- <span data-ttu-id="16792-121">`Mailbox` (Outlook)</span><span class="sxs-lookup"><span data-stu-id="16792-121">`Mailbox` (Outlook)</span></span>
- <span data-ttu-id="16792-122">`Notebook` (OneNote)</span><span class="sxs-lookup"><span data-stu-id="16792-122">`Notebook` (OneNote)</span></span>
- <span data-ttu-id="16792-123">`Presentation` (PowerPoint)</span><span class="sxs-lookup"><span data-stu-id="16792-123">`Presentation` (PowerPoint)</span></span>
- <span data-ttu-id="16792-124">`Project` (Project)</span><span class="sxs-lookup"><span data-stu-id="16792-124">`Project` (Project)</span></span>
- <span data-ttu-id="16792-125">`Workbook` (Excel)</span><span class="sxs-lookup"><span data-stu-id="16792-125">`Workbook` (Excel)</span></span>

### <a name="example"></a><span data-ttu-id="16792-126">Beispiel</span><span class="sxs-lookup"><span data-stu-id="16792-126">Example</span></span>
```xml
<Hosts>
    <Host Name="Mailbox">
    </Host>
</Hosts>
```

## <a name="versionoverrides-node"></a><span data-ttu-id="16792-127">VersionOverrides-Knoten</span><span class="sxs-lookup"><span data-stu-id="16792-127">VersionOverrides node</span></span>
<span data-ttu-id="16792-128">Wenn dies im [VersionOverrides](versionoverrides.md) definiert ist, wird der Hosttyp vom `xsi:type`-Attribut bestimmt.</span><span class="sxs-lookup"><span data-stu-id="16792-128">When defined in [VersionOverrides](versionoverrides.md), the host type is determined by the `xsi:type` attribute.</span></span> 

### <a name="attributes"></a><span data-ttu-id="16792-129">Attribute</span><span class="sxs-lookup"><span data-stu-id="16792-129">Attributes</span></span>

|  <span data-ttu-id="16792-130">Attribut</span><span class="sxs-lookup"><span data-stu-id="16792-130">Attribute</span></span>  |  <span data-ttu-id="16792-131">Erforderlich</span><span class="sxs-lookup"><span data-stu-id="16792-131">Required</span></span>  |  <span data-ttu-id="16792-132">Beschreibung</span><span class="sxs-lookup"><span data-stu-id="16792-132">Description</span></span>  |
|:-----|:-----|:-----|
|  [<span data-ttu-id="16792-133">xsi:type</span><span class="sxs-lookup"><span data-stu-id="16792-133">xsi:type</span></span>](#xsitype)  |  <span data-ttu-id="16792-134">Ja</span><span class="sxs-lookup"><span data-stu-id="16792-134">Yes</span></span>  | <span data-ttu-id="16792-135">Beschreibt den Office-Host, für den diese Einstellungen gelten.</span><span class="sxs-lookup"><span data-stu-id="16792-135">Describes the Office host where these settings apply.</span></span>|

### <a name="child-elements"></a><span data-ttu-id="16792-136">Untergeordnete Elemente</span><span class="sxs-lookup"><span data-stu-id="16792-136">Child elements</span></span>

|  <span data-ttu-id="16792-137">Element</span><span class="sxs-lookup"><span data-stu-id="16792-137">Element</span></span> |  <span data-ttu-id="16792-138">Erforderlich</span><span class="sxs-lookup"><span data-stu-id="16792-138">Required</span></span>  |  <span data-ttu-id="16792-139">Beschreibung</span><span class="sxs-lookup"><span data-stu-id="16792-139">Description</span></span>  |
|:-----|:-----|:-----|
|  [<span data-ttu-id="16792-140">DesktopFormFactor</span><span class="sxs-lookup"><span data-stu-id="16792-140">DesktopFormFactor</span></span>](desktopformfactor.md)    |  <span data-ttu-id="16792-141">Ja</span><span class="sxs-lookup"><span data-stu-id="16792-141">Yes</span></span>   |  <span data-ttu-id="16792-142">Definiert die Einstellungen für den Desktopformfaktor an.</span><span class="sxs-lookup"><span data-stu-id="16792-142">Defines the settings for the desktop form factor.</span></span> |
|  [<span data-ttu-id="16792-143">MobileFormFactor</span><span class="sxs-lookup"><span data-stu-id="16792-143">MobileFormFactor</span></span>](mobileformfactor.md)    |  <span data-ttu-id="16792-144">Nein</span><span class="sxs-lookup"><span data-stu-id="16792-144">No</span></span>   |  <span data-ttu-id="16792-p103">Definiert die Einstellungen für den mobilen Formfaktor. **Hinweis:** Dieses Element wird nur in Outlook für iOS unterstützt.</span><span class="sxs-lookup"><span data-stu-id="16792-p103">Defines the settings for the mobile form factor. **Note:** this element is only supported in Outlook for iOS.</span></span> |
|  [<span data-ttu-id="16792-147">AllFormFactors</span><span class="sxs-lookup"><span data-stu-id="16792-147">AllFormFactors</span></span>](allformfactors.md)    |  <span data-ttu-id="16792-148">Nein</span><span class="sxs-lookup"><span data-stu-id="16792-148">No</span></span>   |  <span data-ttu-id="16792-149">Definiert die Einstellungen für alle Formfaktoren.</span><span class="sxs-lookup"><span data-stu-id="16792-149">Defines the settings for all form factors.</span></span> <span data-ttu-id="16792-150">Nur von benutzerdefinierte Funktionen in Excel verwendet.</span><span class="sxs-lookup"><span data-stu-id="16792-150">Only used by custom functions in Excel.</span></span> |

### <a name="xsitype"></a><span data-ttu-id="16792-151">xsi:type</span><span class="sxs-lookup"><span data-stu-id="16792-151">xsi:type</span></span>

<span data-ttu-id="16792-152">Steuert, auf welchen Office-Host (Word, Excel, PowerPoint, Outlook, OneNote) die enthaltenen Einstellungen angewendet werden.</span><span class="sxs-lookup"><span data-stu-id="16792-152">Controls which Office host (Word, Excel, PowerPoint, Outlook, OneNote) where the contained settings apply.</span></span> <span data-ttu-id="16792-153">Bei dem Wert muss es sich um Folgendes handeln:</span><span class="sxs-lookup"><span data-stu-id="16792-153">The value must be one of the following:</span></span>

- <span data-ttu-id="16792-154">`Document` (Word)</span><span class="sxs-lookup"><span data-stu-id="16792-154">`Document` (Word)</span></span>
- <span data-ttu-id="16792-155">`MailHost` (Outlook)</span><span class="sxs-lookup"><span data-stu-id="16792-155">`MailHost` (Outlook)</span></span>    
- <span data-ttu-id="16792-156">`Notebook` (OneNote)</span><span class="sxs-lookup"><span data-stu-id="16792-156">`Notebook` (OneNote)</span></span>
- <span data-ttu-id="16792-157">`Presentation` (PowerPoint)</span><span class="sxs-lookup"><span data-stu-id="16792-157">`Presentation` (PowerPoint)</span></span>
- <span data-ttu-id="16792-158">`Workbook` (Excel)</span><span class="sxs-lookup"><span data-stu-id="16792-158">`Workbook` (Excel)</span></span>

## <a name="host-example"></a><span data-ttu-id="16792-159">Host-Beispiel</span><span class="sxs-lookup"><span data-stu-id="16792-159">Host example</span></span> 
```xml
<Hosts>
    <Host xsi:type="MailHost">
        <!-- Host Settings -->
    </Host>
</Hosts>
```
