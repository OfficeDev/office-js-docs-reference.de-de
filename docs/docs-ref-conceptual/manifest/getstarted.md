# <a name="getstarted-element"></a><span data-ttu-id="0817e-101">GetStarted-Element</span><span class="sxs-lookup"><span data-stu-id="0817e-101">GetStarted element</span></span>

<span data-ttu-id="0817e-p101">Stellt Informationen bereit, die von der Beschriftung verwendet werden, die angezeigt wird, wenn das Add-In auf Word-, Excel-, PowerPoint- und OneNote-Hosts installiert wird. Das **GetStarted**-Element ist ein untergeordnetes Element von [DesktopFormFactor](desktopformfactor.md).</span><span class="sxs-lookup"><span data-stu-id="0817e-p101">Provides information used by the callout that appears when the add-in is installed in Word, Excel, PowerPoint, and OneNote hosts. The **GetStarted** element is a child element of [DesktopFormFactor](desktopformfactor.md).</span></span>

## <a name="child-elements"></a><span data-ttu-id="0817e-104">Untergeordnete Elemente</span><span class="sxs-lookup"><span data-stu-id="0817e-104">Child elements</span></span>

| <span data-ttu-id="0817e-105">Element</span><span class="sxs-lookup"><span data-stu-id="0817e-105">Element</span></span>                       | <span data-ttu-id="0817e-106">Erforderlich</span><span class="sxs-lookup"><span data-stu-id="0817e-106">Required</span></span> | <span data-ttu-id="0817e-107">Beschreibung</span><span class="sxs-lookup"><span data-stu-id="0817e-107">Description</span></span>                                        |
|:------------------------------|:--------:|:---------------------------------------------------|
| [<span data-ttu-id="0817e-108">Titel</span><span class="sxs-lookup"><span data-stu-id="0817e-108">Title</span></span>](#title)               | <span data-ttu-id="0817e-109">Ja</span><span class="sxs-lookup"><span data-stu-id="0817e-109">Yes</span></span>      | <span data-ttu-id="0817e-110">Definiert, wo ein Add-In Funktionen verfügbar macht.</span><span class="sxs-lookup"><span data-stu-id="0817e-110">Defines where an add-in exposes functionality.</span></span>     |
| [<span data-ttu-id="0817e-111">Beschreibung</span><span class="sxs-lookup"><span data-stu-id="0817e-111">Description</span></span>](#description)   | <span data-ttu-id="0817e-112">Ja</span><span class="sxs-lookup"><span data-stu-id="0817e-112">Yes</span></span>      | <span data-ttu-id="0817e-113">Eine URL zu einer Datei, die JavaScript-Funktionen enthält.</span><span class="sxs-lookup"><span data-stu-id="0817e-113">A URL to a file that contains JavaScript functions.</span></span>|
| [<span data-ttu-id="0817e-114">LearnMoreUrl</span><span class="sxs-lookup"><span data-stu-id="0817e-114">LearnMoreUrl</span></span>](#learnmoreurl) | <span data-ttu-id="0817e-115">Nein</span><span class="sxs-lookup"><span data-stu-id="0817e-115">No</span></span>       | <span data-ttu-id="0817e-116">Eine URL zu einer Seite, auf der das Add-In ausführlich erläutert wird.</span><span class="sxs-lookup"><span data-stu-id="0817e-116">A URL to a page that explains the add-in in detail.</span></span>   |

### <a name="title"></a><span data-ttu-id="0817e-117">Titel</span><span class="sxs-lookup"><span data-stu-id="0817e-117">Title</span></span> 

<span data-ttu-id="0817e-p102">Erforderlich. Der Titel, der für den oberen Rand der Beschriftung verwendet wird. Das**resid**-Attribut verweist auf eine gültige ID im **ShortStrings**-Element im Abschnitt [Ressourcen](resources.md).</span><span class="sxs-lookup"><span data-stu-id="0817e-p102">Required. The title used for the top of the callout. The **resid** attribute references a valid ID in the **ShortStrings** element in the [Resources](resources.md) section.</span></span>

### <a name="description"></a><span data-ttu-id="0817e-121">Beschreibung</span><span class="sxs-lookup"><span data-stu-id="0817e-121">Description</span></span>

<span data-ttu-id="0817e-p103">Erforderlich. Die Beschreibung/der Textinhalt der Beschriftung. Das**resid**-Attribut verweist auf eine gültige ID im **LongStrings**-Element im Abschnitt [Ressourcen](resources.md).</span><span class="sxs-lookup"><span data-stu-id="0817e-p103">Required. The description / body content for the callout. The **resid** attribute references a valid ID in the **LongStrings** element in the [Resources](resources.md) section.</span></span>

### <a name="learnmoreurl"></a><span data-ttu-id="0817e-125">LearnMoreUrl</span><span class="sxs-lookup"><span data-stu-id="0817e-125">LearnMoreUrl</span></span>

<span data-ttu-id="0817e-p104">Erforderlich. Die URL zu einer Seite, auf der der Benutzer mehr über das Add-In erfahren kann. Das**resid**-Attribut verweist auf eine gültige ID im **Urls**-Element im Abschnitt [Ressourcen](resources.md).</span><span class="sxs-lookup"><span data-stu-id="0817e-p104">Required. The URL to a page where the user can learn more about your add-in. The **resid** attribute references a valid ID in the **Urls** element in the [Resources](resources.md) section.</span></span>

> [!NOTE]
> <span data-ttu-id="0817e-129">**LearnMoreUrl** wird derzeit nicht in Word, Excel oder PowerPoint-Clients gerendert.</span><span class="sxs-lookup"><span data-stu-id="0817e-129">**LearnMoreUrl** does not currently render in Word, Excel, or PowerPoint clients.</span></span> <span data-ttu-id="0817e-130">Es wird empfohlen, dass Sie diese URL für alle Clients hinzufügen, damit die URL dargestellt wird, sobald sie verfügbar ist.</span><span class="sxs-lookup"><span data-stu-id="0817e-130">We recommend that you add this URL for all clients so that the URL will render when it becomes available.</span></span> 

## <a name="see-also"></a><span data-ttu-id="0817e-131">Siehe auch</span><span class="sxs-lookup"><span data-stu-id="0817e-131">See also</span></span>

<span data-ttu-id="0817e-132">In den folgenden Codebeispielen wird das **GetStarted**-Element verwendet:</span><span class="sxs-lookup"><span data-stu-id="0817e-132">The following code samples use the **GetStarted** element:</span></span>

* [<span data-ttu-id="0817e-133">Excel-Web-Add-In zum Bearbeiten der Tabellen- und Diagrammformatierung</span><span class="sxs-lookup"><span data-stu-id="0817e-133">Excel Web Add-in for Manipulating Table and Chart Formatting</span></span>](https://github.com/OfficeDev/Excel-Add-in-JavaScript-SalesTracker)
* [<span data-ttu-id="0817e-134">Word Add-In JavaScript SpecKit</span><span class="sxs-lookup"><span data-stu-id="0817e-134">Word Add-in JavaScript SpecKit</span></span>](https://github.com/OfficeDev/Word-Add-in-JS-SpecKit)
* [<span data-ttu-id="0817e-135">Einfügen von Excel-Diagrammen mithilfe von Microsoft Graph in eine PowerPoint-add-in</span><span class="sxs-lookup"><span data-stu-id="0817e-135">Insert Excel charts using Microsoft Graph in a PowerPoint add-in</span></span>](https://github.com/OfficeDev/PowerPoint-Add-in-Microsoft-Graph-ASPNET-InsertChart)
