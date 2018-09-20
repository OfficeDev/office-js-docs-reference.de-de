# <a name="outlook-javascript-api-requirement-sets"></a><span data-ttu-id="5dffa-101">Outlook-JavaScript-API-anforderungssätzen</span><span class="sxs-lookup"><span data-stu-id="5dffa-101">Outlook JavaScript API requirement sets</span></span>

<span data-ttu-id="5dffa-102">Outlook-add-ins deklarieren, welche API-Versionen, die hierfür erforderlich sind, mithilfe von Element " [Requirements](/javascript/office/manifest/requirements) " in ihre [manifest](https://docs.microsoft.com/office/dev/add-ins/develop/add-in-manifests).</span><span class="sxs-lookup"><span data-stu-id="5dffa-102">Outlook add-ins declare what API versions they require by using the [Requirements](/javascript/office/manifest/requirements) element in their [manifest](https://docs.microsoft.com/office/dev/add-ins/develop/add-in-manifests).</span></span> <span data-ttu-id="5dffa-103">Outlook-Add-Ins umfassen immer ein [Set](/javascript/office/manifest/set)-Element mit einem `Name`-Attribut, das auf `Mailbox` festgelegt ist, und einem `MinVersion`-Attribut, das auf den mindestens erforderlichen API-Anforderungssatz festgelegt ist, der das Add-In-Szenario unterstützt.</span><span class="sxs-lookup"><span data-stu-id="5dffa-103">Outlook add-ins always include a [Set](/javascript/office/manifest/set) element with a `Name` attribute set to `Mailbox` and a `MinVersion` attribute set to the minimum API requirement set that supports the add-in's scenarios.</span></span>

<span data-ttu-id="5dffa-104">Der folgende Manifestcodeausschnitt gibt z. B. als Mindestanforderungssatz 1.1 an:</span><span class="sxs-lookup"><span data-stu-id="5dffa-104">For example, the following manifest snippet indicates a minimum requirement set of 1.1:</span></span>

```xml
<Requirements>
  <Sets>
    <Set Name="MailBox" MinVersion="1.1" />
  </Sets>
</Requirements>
```

<span data-ttu-id="5dffa-105">Alle Outlook-APIs gehören zum `Mailbox` [-Anforderungssatz](https://docs.microsoft.com/office/dev/add-ins/develop/specify-office-hosts-and-api-requirements).</span><span class="sxs-lookup"><span data-stu-id="5dffa-105">All Outlook APIs belong to the `Mailbox` [requirement set](https://docs.microsoft.com/office/dev/add-ins/develop/specify-office-hosts-and-api-requirements).</span></span> <span data-ttu-id="5dffa-106">Der `Mailbox`-Anforderungssatz verfügt über Versionen, und jeder neuen Satz von APIs, den wir veröffentlichen, gehört zu einer höheren Version des Satzes.</span><span class="sxs-lookup"><span data-stu-id="5dffa-106">The `Mailbox` requirement set has versions, and each new set of APIs that we release belongs to a higher version of the set.</span></span> <span data-ttu-id="5dffa-107">Nicht alle Outlook-Clients unterstützen den neuesten Satz von APIs, aber wenn ein Outlook-Client Unterstützung für einen Anforderungssatz deklariert, unterstützt er alle APIs in diesem Anforderungssatz.</span><span class="sxs-lookup"><span data-stu-id="5dffa-107">Not all Outlook clients support the newest set of APIs, but if an Outlook client declares support for a requirement set, it supports all of the APIs in that requirement set.</span></span>

<span data-ttu-id="5dffa-p103">Das Festlegen einer Mindestanforderungssatzversion im Manifest steuert, in welchem Outlook-Client das Add-In angezeigt wird. Wenn ein Client den Mindestanforderungssatz nicht unterstützt, wird das Add-In nicht geladen. Wenn z. B. Anforderungssatzversion 1.3 angegeben wird, bedeutet dies, dass das Add-In nicht in einem Outlook-Client angezeigt wird, der nicht mindestens 1.3 unterstützt.</span><span class="sxs-lookup"><span data-stu-id="5dffa-p103">Setting a minimum requirement set version in the manifest controls which Outlook client the add-in will appear in. If a client does not support the minimum requirement set, it does not load the add-in. For example, if requirement set version 1.3 is specified, this means the add-in will not show up in any Outlook client that doesn't support at least 1.3.</span></span>

## <a name="using-apis-from-later-requirement-sets"></a><span data-ttu-id="5dffa-111">Verwenden von APIs aus neueren Anforderungssätzen</span><span class="sxs-lookup"><span data-stu-id="5dffa-111">Using APIs from later requirement sets</span></span>

<span data-ttu-id="5dffa-p104">Das Festlegen eines Anforderungssatzes beschränkt nicht die verfügbaren APIs, die das Add-In verwenden kann. Wenn das Add-In beispielsweise Anforderungssatz 1.1 angibt, aber in einem Outlook-Client ausgeführt wird, der 1.3 unterstützt, kann das Add-In APIs aus Anforderungssatz 1.3 verwenden.</span><span class="sxs-lookup"><span data-stu-id="5dffa-p104">Setting a requirement set does not limit the available APIs that the add-in can use. For example, if the add-in specifies requirement set 1.1, but it is running in an Outlook client which support 1.3, the add-in can use APIs from requirement set 1.3\.</span></span>

<span data-ttu-id="5dffa-114">Wenn Entwickler neuere APIs verwenden möchten, können sie einfach mithilfe der Standard-JavaScript-Methode ihr Vorhandensein überprüfen.</span><span class="sxs-lookup"><span data-stu-id="5dffa-114">To use newer APIs, developers can just check for their existence by using standard JavaScript technique</span></span>

```js
if (item.somePropertyOrFunction !== undefined) {
  item.somePropertyOrFunction ...
}
```

<span data-ttu-id="5dffa-115">Für APIs, die sich in der im Manifest angegebenen Version des Anforderungssatzes befinden, sind solche Überprüfungen nicht erforderlich.</span><span class="sxs-lookup"><span data-stu-id="5dffa-115">No such checks are necessary for any APIs which are present in the requirement set version specified in in the manifest.</span></span>

## <a name="choosing-a-minimum-requirement-set"></a><span data-ttu-id="5dffa-116">Auswählen einen Mindestanforderungssatzes</span><span class="sxs-lookup"><span data-stu-id="5dffa-116">Choosing a minimum requirement set</span></span>

<span data-ttu-id="5dffa-117">Entwickler sollten den ältesten Anforderungssatz verwenden, der die kritischen APIs für ihr Szenario enthält, ohne die das Add-In nicht funktioniert.</span><span class="sxs-lookup"><span data-stu-id="5dffa-117">Developers should use the earliest requirement set that contains the critical set of APIs for their scenario, without which the add-in won't work.</span></span>

## <a name="clients"></a><span data-ttu-id="5dffa-118">Clients</span><span class="sxs-lookup"><span data-stu-id="5dffa-118">Clients</span></span>

<span data-ttu-id="5dffa-119">Die folgenden Clients unterstützen Outlook-Add-Ins.</span><span class="sxs-lookup"><span data-stu-id="5dffa-119">The following clients support Outlook add-ins.</span></span>

| <span data-ttu-id="5dffa-120">Client</span><span class="sxs-lookup"><span data-stu-id="5dffa-120">Client</span></span> | <span data-ttu-id="5dffa-121">Unterstützte API-Anforderungssätze</span><span class="sxs-lookup"><span data-stu-id="5dffa-121">Supported API requirement sets</span></span> |
| --- | --- |
| <span data-ttu-id="5dffa-122">Outlook 2016 (Klick-und-Los) für Windows</span><span class="sxs-lookup"><span data-stu-id="5dffa-122">Outlook 2016 (Click-to-Run) for Windows</span></span> | <span data-ttu-id="5dffa-123">1.1, 1.2, 1.3, 1.4, 1,5, 1,6</span><span class="sxs-lookup"><span data-stu-id="5dffa-123">1.1, 1.2, 1.3, 1.4, 1.5, 1.6</span></span> |
| <span data-ttu-id="5dffa-124">Outlook 2016 (MSI) für Windows</span><span class="sxs-lookup"><span data-stu-id="5dffa-124">Outlook 2016 (MSI) for Windows</span></span> | <span data-ttu-id="5dffa-125">1.1, 1.2, 1.3, 1.4</span><span class="sxs-lookup"><span data-stu-id="5dffa-125">1.1, 1.2, 1.3, 1.4</span></span> |
| <span data-ttu-id="5dffa-126">Outlook 2016 für Mac</span><span class="sxs-lookup"><span data-stu-id="5dffa-126">Outlook 2016 for Mac</span></span> | <span data-ttu-id="5dffa-127">1.1, 1.2, 1.3, 1.4, 1,5, 1,6</span><span class="sxs-lookup"><span data-stu-id="5dffa-127">1.1, 1.2, 1.3, 1.4, 1.5, 1.6</span></span> |
| <span data-ttu-id="5dffa-128">Outlook 2013 für Windows</span><span class="sxs-lookup"><span data-stu-id="5dffa-128">Outlook 2013 for Windows</span></span> | <span data-ttu-id="5dffa-129">1.1, 1.2, 1.3, 1.4</span><span class="sxs-lookup"><span data-stu-id="5dffa-129">1.1, 1.2, 1.3, 1.4</span></span> |
| <span data-ttu-id="5dffa-130">Outlook für iPhone</span><span class="sxs-lookup"><span data-stu-id="5dffa-130">Outlook for iPhone</span></span> | <span data-ttu-id="5dffa-131">1.1, 1.2, 1.3, 1.4, 1.5</span><span class="sxs-lookup"><span data-stu-id="5dffa-131">1.1, 1.2, 1.3, 1.4, 1.5</span></span> |
| <span data-ttu-id="5dffa-132">Outlook für Android</span><span class="sxs-lookup"><span data-stu-id="5dffa-132">Outlook for Android</span></span> | <span data-ttu-id="5dffa-133">1.1, 1.2, 1.3, 1.4, 1.5</span><span class="sxs-lookup"><span data-stu-id="5dffa-133">1.1, 1.2, 1.3, 1.4, 1.5</span></span> |
| <span data-ttu-id="5dffa-134">Outlook im Web (Office 365 und Outlook.com)</span><span class="sxs-lookup"><span data-stu-id="5dffa-134">Outlook on the web (Office 365 and Outlook.com)</span></span> | <span data-ttu-id="5dffa-135">1.1, 1.2, 1.3, 1.4, 1,5, 1,6</span><span class="sxs-lookup"><span data-stu-id="5dffa-135">1.1, 1.2, 1.3, 1.4, 1.5, 1.6</span></span> |
| <span data-ttu-id="5dffa-136">Outlook Web App (Exchange 2013 lokal)</span><span class="sxs-lookup"><span data-stu-id="5dffa-136">Outlook Web App (Exchange 2013 On-Premise)</span></span> | <span data-ttu-id="5dffa-137">1.1</span><span class="sxs-lookup"><span data-stu-id="5dffa-137">1.1</span></span> |
| <span data-ttu-id="5dffa-138">Outlook Web App (Exchange 2016 lokal)</span><span class="sxs-lookup"><span data-stu-id="5dffa-138">Outlook Web App (Exchange 2016 On-Premise)</span></span> | <span data-ttu-id="5dffa-p105">1.1, 1.2. 1.3</span><span class="sxs-lookup"><span data-stu-id="5dffa-p105">1.1, 1.2. 1.3</span></span> |

> [!NOTE] 
> <span data-ttu-id="5dffa-141">Unterstützung für 1.3 in Outlook 2013 wurde als Teil der [Dezember 8 2015, aktualisieren für Outlook 2013 (KB3114349)](https://support.microsoft.com/kb/3114349)hinzugefügt.</span><span class="sxs-lookup"><span data-stu-id="5dffa-141">Support for 1.3 in Outlook 2013 was added as part of the [December 8, 2015, update for Outlook 2013 (KB3114349)](https://support.microsoft.com/kb/3114349).</span></span> <span data-ttu-id="5dffa-142">Unterstützung für 1.4 in Outlook 2013 wurde als Teil der [September 13, 2016, aktualisieren für Outlook 2013 (KB3118280)](https://support.microsoft.com/help/3118280)hinzugefügt.</span><span class="sxs-lookup"><span data-stu-id="5dffa-142">Support for 1.4 in Outlook 2013 was added as part of the [September 13, 2016, update for Outlook 2013 (KB3118280)](https://support.microsoft.com/help/3118280).</span></span>
