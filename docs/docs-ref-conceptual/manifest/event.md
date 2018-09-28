# <a name="event-element"></a><span data-ttu-id="77ed1-101">Element „Event“</span><span class="sxs-lookup"><span data-stu-id="77ed1-101">Event element</span></span>

<span data-ttu-id="77ed1-102">Dieses Element definiert einen Ereignishandler in einem Add-In.</span><span class="sxs-lookup"><span data-stu-id="77ed1-102">Defines an event handler in an add-in.</span></span>

> [!NOTE] 
> <span data-ttu-id="77ed1-103">Die `Event` Element ist derzeit nur von Outlook im Web in Office 365 unterstützt.</span><span class="sxs-lookup"><span data-stu-id="77ed1-103">The `Event` element is currently only supported by Outlook on the web in Office 365.</span></span>

## <a name="attributes"></a><span data-ttu-id="77ed1-104">Attribute</span><span class="sxs-lookup"><span data-stu-id="77ed1-104">Attributes</span></span>

|  <span data-ttu-id="77ed1-105">Attribut</span><span class="sxs-lookup"><span data-stu-id="77ed1-105">Attribute</span></span>  |  <span data-ttu-id="77ed1-106">Erforderlich</span><span class="sxs-lookup"><span data-stu-id="77ed1-106">Required</span></span>  |  <span data-ttu-id="77ed1-107">Beschreibung</span><span class="sxs-lookup"><span data-stu-id="77ed1-107">Description</span></span>  |
|:-----|:-----|:-----|
|  [<span data-ttu-id="77ed1-108">Typ</span><span class="sxs-lookup"><span data-stu-id="77ed1-108">Type</span></span>](#type-attribute)  |  <span data-ttu-id="77ed1-109">Ja</span><span class="sxs-lookup"><span data-stu-id="77ed1-109">Yes</span></span>  | <span data-ttu-id="77ed1-110">Gibt das zu behandelnde Ereignis an.</span><span class="sxs-lookup"><span data-stu-id="77ed1-110">Specifies the event to handle.</span></span> |
|  [<span data-ttu-id="77ed1-111">FunctionExecution</span><span class="sxs-lookup"><span data-stu-id="77ed1-111">FunctionExecution</span></span>](#functionexecution-attribute)  |  <span data-ttu-id="77ed1-112">Ja</span><span class="sxs-lookup"><span data-stu-id="77ed1-112">Yes</span></span>  | <span data-ttu-id="77ed1-p101">Gibt den Ausführungsstil für den Ereignishandler an: asynchron oder synchron. Zurzeit werden nur synchrone Ereignishandler unterstützt.</span><span class="sxs-lookup"><span data-stu-id="77ed1-p101">Specifies the execution style for the event handler, asynchronous or synchronous. Currently only synchronous event handlers are supported.</span></span> |
|  [<span data-ttu-id="77ed1-115">FunctionName</span><span class="sxs-lookup"><span data-stu-id="77ed1-115">FunctionName</span></span>](#functionname-attribute)  |  <span data-ttu-id="77ed1-116">Ja</span><span class="sxs-lookup"><span data-stu-id="77ed1-116">Yes</span></span>  | <span data-ttu-id="77ed1-117">Gibt den Funktionsnamen des Ereignishandlers an.</span><span class="sxs-lookup"><span data-stu-id="77ed1-117">Specifies the function name for the event handler.</span></span> |

### <a name="type-attribute"></a><span data-ttu-id="77ed1-118">Attribut „Type“</span><span class="sxs-lookup"><span data-stu-id="77ed1-118">Type attribute</span></span>

<span data-ttu-id="77ed1-p102">Dieses Attribut ist erforderlich. Es gibt an, welches Ereignis den Ereignishandler aufrufen wird. Die möglichen Werte für dieses Attribut sind in der nachfolgenden Tabelle aufgeführt.</span><span class="sxs-lookup"><span data-stu-id="77ed1-p102">Required. Specifies which event will invoke the event handler. The possible values for this attribute are specified in the following table.</span></span>

|  <span data-ttu-id="77ed1-122">Ereignistyp</span><span class="sxs-lookup"><span data-stu-id="77ed1-122">Event type</span></span>  |  <span data-ttu-id="77ed1-123">Beschreibung</span><span class="sxs-lookup"><span data-stu-id="77ed1-123">Description</span></span>  |
|:-----|:-----|
|  `ItemSend`  |  <span data-ttu-id="77ed1-124">Der Ereignishandler wird aufgerufen, wenn der Benutzer eine Nachricht oder eine Besprechungseinladung sendet.</span><span class="sxs-lookup"><span data-stu-id="77ed1-124">The event handler will be invoked when the user sends a message or meeting invitation.</span></span>  |

### <a name="functionexecution-attribute"></a><span data-ttu-id="77ed1-125">Attribut „FunctionExecution“</span><span class="sxs-lookup"><span data-stu-id="77ed1-125">FunctionExecution attribute</span></span>

<span data-ttu-id="77ed1-p103">Dieses Attribut ist erforderlich. Es MUSS auf `synchronous` gesetzt werden.</span><span class="sxs-lookup"><span data-stu-id="77ed1-p103">Required. MUST be set to `synchronous`.</span></span>

### <a name="functionname-attribute"></a><span data-ttu-id="77ed1-128">Attribut „FunctionName“</span><span class="sxs-lookup"><span data-stu-id="77ed1-128">FunctionName attribute</span></span>

<span data-ttu-id="77ed1-p104">Dieses Attribut ist erforderlich. Es gibt den Funktionsnamen des Ereignishandlers an. Dieser Wert muss einem Funktionsnamen in der [Funktionsdatei](functionfile.md) des Add-Ins entsprechen.</span><span class="sxs-lookup"><span data-stu-id="77ed1-p104">Required. Specifies the function name of the event handler. This value must match a function name in the add-in's [function file](functionfile.md).</span></span>

```xml
<Event Type="ItemSend" FunctionExecution="synchronous" FunctionName="itemSendHandler" /> 
```