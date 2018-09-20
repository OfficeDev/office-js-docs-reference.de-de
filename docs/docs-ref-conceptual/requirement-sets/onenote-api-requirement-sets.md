# <a name="onenote-javascript-api-requirement-sets"></a><span data-ttu-id="7b64c-101">JavaScript-API-Anforderungssätze für OneNote</span><span class="sxs-lookup"><span data-stu-id="7b64c-101">OneNote JavaScript API requirement sets</span></span>

<span data-ttu-id="7b64c-102">Anforderungssätze sind benannte Gruppen von API-Mitgliedern.</span><span class="sxs-lookup"><span data-stu-id="7b64c-102">Requirement sets are named groups of API members.</span></span> <span data-ttu-id="7b64c-103">Office-Add-ins anforderungssätzen im Manifest angegebenen verwenden oder eine Überprüfung zur Laufzeit verwenden, um zu bestimmen, ob ein Office-Host APIs unterstützt, die ein Add-in muss.</span><span class="sxs-lookup"><span data-stu-id="7b64c-103">Office Add-ins use requirement sets specified in the manifest or use a runtime check to determine whether an Office host supports APIs that an add-in needs.</span></span> <span data-ttu-id="7b64c-104">Weitere Informationen finden Sie unter [Office-Versionen und Anforderung festgelegt](https://docs.microsoft.com/office/dev/add-ins/develop/office-versions-and-requirement-sets).</span><span class="sxs-lookup"><span data-stu-id="7b64c-104">For more information, see [Office versions and requirement sets](https://docs.microsoft.com/office/dev/add-ins/develop/office-versions-and-requirement-sets).</span></span>

<span data-ttu-id="7b64c-105">In der folgenden Tabelle sind die OneNote-Anforderungssätze, die Office-Hostanwendungen, die diese Anforderungssätze unterstützen, und die Buildversionen oder das Verfügbarkeitsdatum aufgeführt.</span><span class="sxs-lookup"><span data-stu-id="7b64c-105">The following table lists the OneNote requirement sets, the Office host applications that support those requirement sets, and the build versions or availability date.</span></span>

|  <span data-ttu-id="7b64c-106">Anforderungssatz</span><span class="sxs-lookup"><span data-stu-id="7b64c-106">Requirement set</span></span>  |  <span data-ttu-id="7b64c-107">Office Online</span><span class="sxs-lookup"><span data-stu-id="7b64c-107">Office Online</span></span> | 
|:-----|:-----|
| <span data-ttu-id="7b64c-108">OneNoteApi 1.1</span><span class="sxs-lookup"><span data-stu-id="7b64c-108">OneNoteApi 1.1</span></span>  | <span data-ttu-id="7b64c-109">September 2016</span><span class="sxs-lookup"><span data-stu-id="7b64c-109">September 2016</span></span> |  

## <a name="office-common-api-requirement-sets"></a><span data-ttu-id="7b64c-110">Allgemeine Office-API-Anforderungssätze</span><span class="sxs-lookup"><span data-stu-id="7b64c-110">Office common API requirement sets</span></span>

<span data-ttu-id="7b64c-111">Informationen zu allgemeinen API-Anforderungssätzen finden Sie unter [Allgemeine Office-API-Anforderungssätze](office-add-in-requirement-sets.md).</span><span class="sxs-lookup"><span data-stu-id="7b64c-111">For information about common API requirement sets, see [Office common API requirement sets](office-add-in-requirement-sets.md).</span></span>

## <a name="onenote-javascript-api-11"></a><span data-ttu-id="7b64c-112">OneNote-JavaScript-API 1.1</span><span class="sxs-lookup"><span data-stu-id="7b64c-112">OneNote JavaScript API 1.1</span></span> 

<span data-ttu-id="7b64c-113">Die OneNote-JavaScript-API 1.1 ist die erste Version der API.</span><span class="sxs-lookup"><span data-stu-id="7b64c-113">OneNote JavaScript API 1.1 is the first version of the API.</span></span> <span data-ttu-id="7b64c-114">Weitere Informationen zur API finden Sie unter der [OneNote-JavaScript-API programming (Übersicht)](https://docs.microsoft.com/office/dev/add-ins/onenote/onenote-add-ins-programming-overview).</span><span class="sxs-lookup"><span data-stu-id="7b64c-114">For details about the API, see the [OneNote JavaScript API programming overview](https://docs.microsoft.com/office/dev/add-ins/onenote/onenote-add-ins-programming-overview).</span></span>

## <a name="runtime-requirement-support-check"></a><span data-ttu-id="7b64c-115">Überprüfung der Unterstützung einer Anforderung zur Laufzeit</span><span class="sxs-lookup"><span data-stu-id="7b64c-115">Runtime requirement support check</span></span>

<span data-ttu-id="7b64c-116">Während der Laufzeit können Add-Ins überprüfen, ob ein bestimmter Host einen API-Anforderungssatz unterstützt, indem die folgende Überprüfung durchgeführt wird:</span><span class="sxs-lookup"><span data-stu-id="7b64c-116">During the runtime, add-ins can check if a particular host supports an API requirement set by doing the following-check:</span></span> 

```js
if (Office.context.requirements.isSetSupported('OneNoteApi', 1.1) === true) {
  /// perform actions
}
else {
  /// provide alternate flow/logic
}
```

## <a name="manifest-based-requirement-support-check"></a><span data-ttu-id="7b64c-117">Überprüfung der Unterstützung einer manifestbasierten Anforderung</span><span class="sxs-lookup"><span data-stu-id="7b64c-117">Manifest-based requirement support check</span></span>

<span data-ttu-id="7b64c-p103">Verwenden Sie das Requirements-Element im Add-In-Manifest, um wichtige Anforderungssätze oder API-Elemente anzugeben, die Ihr Add-In verwenden muss. Wenn der Office-Host oder die Plattform die im Requirements-Element angegebenen Anforderungssätze oder API-Elemente nicht unterstützt, wird das Add-In nicht auf diesem Host oder dieser Plattform ausgeführt und nicht unter „Meine Add-Ins“ angezeigt.</span><span class="sxs-lookup"><span data-stu-id="7b64c-p103">Use the Requirements element in the add-in manifest to specify critical requirement sets or API members that your add-in must use. If the Office host or platform doesn't support the requirement sets or API members specified in the Requirements element, the add-in won't run in that host or platform, and won't display in My Add-ins.</span></span>

<span data-ttu-id="7b64c-120">Das folgende Codebeispiel zeigt ein Add-In, das in allen Office-Hostanwendungen geladen wird, die den OneNoteApi-Anforderungssatz, Version 1.1, unterstützen.</span><span class="sxs-lookup"><span data-stu-id="7b64c-120">The following code example shows an add-in that loads in all Office host applications that support the OneNoteApi requirement set, version 1.1.</span></span>

```xml
<Requirements>
   <Sets DefaultMinVersion="1.1">
      <Set Name="OneNoteApi" MinVersion="1.1"/>
   </Sets>
</Requirements>
```

## <a name="see-also"></a><span data-ttu-id="7b64c-121">Siehe auch</span><span class="sxs-lookup"><span data-stu-id="7b64c-121">See also</span></span>

- [<span data-ttu-id="7b64c-122">Office-Versionen und Anforderungssätze</span><span class="sxs-lookup"><span data-stu-id="7b64c-122">Office versions and requirement sets</span></span>](https://docs.microsoft.com/office/dev/add-ins/develop/office-versions-and-requirement-sets)
- <span data-ttu-id="7b64c-123">
  [Angeben von Office-Hosts und API-Anforderungen](https://docs.microsoft.com/office/dev/add-ins/develop/specify-office-hosts-and-api-requirements)</span><span class="sxs-lookup"><span data-stu-id="7b64c-123">[Specify Office hosts and API requirements](https://docs.microsoft.com/office/dev/add-ins/develop/specify-office-hosts-and-api-requirements)</span></span>
- [<span data-ttu-id="7b64c-124">XML-Manifest für Office-Add-Ins</span><span class="sxs-lookup"><span data-stu-id="7b64c-124">Office Add-ins XML manifest</span></span>](https://docs.microsoft.com/office/dev/add-ins/develop/add-in-manifests)
