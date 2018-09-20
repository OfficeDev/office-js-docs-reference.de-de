
# <a name="diagnostics"></a><span data-ttu-id="7064a-101">diagnostics</span><span class="sxs-lookup"><span data-stu-id="7064a-101">diagnostics</span></span>

### <span data-ttu-id="7064a-p101">[Office](Office.md)[.context](Office.context.md)[.mailbox](Office.context.mailbox.md). diagnostics</span><span class="sxs-lookup"><span data-stu-id="7064a-p101">[Office](Office.md)[.context](Office.context.md)[.mailbox](Office.context.mailbox.md). diagnostics</span></span>

<span data-ttu-id="7064a-104">Stellt einem Outlook-Add-In Diagnoseinformationen bereit.</span><span class="sxs-lookup"><span data-stu-id="7064a-104">Provides diagnostic information to an Outlook add-in.</span></span>

##### <a name="requirements"></a><span data-ttu-id="7064a-105">Anforderungen</span><span class="sxs-lookup"><span data-stu-id="7064a-105">Requirements</span></span>

|<span data-ttu-id="7064a-106">Anforderung</span><span class="sxs-lookup"><span data-stu-id="7064a-106">Requirement</span></span>| <span data-ttu-id="7064a-107">Wert</span><span class="sxs-lookup"><span data-stu-id="7064a-107">Value</span></span>|
|---|---|
|[<span data-ttu-id="7064a-108">Mindestversion des Postfachanforderungssatzes</span><span class="sxs-lookup"><span data-stu-id="7064a-108">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="7064a-109">1.0</span><span class="sxs-lookup"><span data-stu-id="7064a-109">1.0</span></span>|
|[<span data-ttu-id="7064a-110">Mindestberechtigungsstufe</span><span class="sxs-lookup"><span data-stu-id="7064a-110">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="7064a-111">ReadItem</span><span class="sxs-lookup"><span data-stu-id="7064a-111">ReadItem</span></span>|
|[<span data-ttu-id="7064a-112">Zutreffender Outlook-Modus</span><span class="sxs-lookup"><span data-stu-id="7064a-112">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="7064a-113">Verfassen oder Lesen</span><span class="sxs-lookup"><span data-stu-id="7064a-113">Compose or read</span></span>|

### <a name="members"></a><span data-ttu-id="7064a-114">Elemente</span><span class="sxs-lookup"><span data-stu-id="7064a-114">Members</span></span>

####  <a name="hostname-string"></a><span data-ttu-id="7064a-115">hostName :String</span><span class="sxs-lookup"><span data-stu-id="7064a-115">hostName :String</span></span>

<span data-ttu-id="7064a-116">Ruft eine Zeichenfolge ab, die den Namen der Hostanwendung darstellt.</span><span class="sxs-lookup"><span data-stu-id="7064a-116">Gets a string that represents the name of the host application.</span></span>

<span data-ttu-id="7064a-117">Eine Zeichenfolge, die eine der folgenden Werte sein kann: `Outlook`, `OutlookIOS`, oder `OutlookWebApp`.</span><span class="sxs-lookup"><span data-stu-id="7064a-117">A string that can be one of the following values: `Outlook`, `OutlookIOS`, or `OutlookWebApp`.</span></span>

##### <a name="type"></a><span data-ttu-id="7064a-118">Typ:</span><span class="sxs-lookup"><span data-stu-id="7064a-118">Type:</span></span>

*   <span data-ttu-id="7064a-119">String</span><span class="sxs-lookup"><span data-stu-id="7064a-119">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="7064a-120">Anforderungen</span><span class="sxs-lookup"><span data-stu-id="7064a-120">Requirements</span></span>

|<span data-ttu-id="7064a-121">Anforderung</span><span class="sxs-lookup"><span data-stu-id="7064a-121">Requirement</span></span>| <span data-ttu-id="7064a-122">Wert</span><span class="sxs-lookup"><span data-stu-id="7064a-122">Value</span></span>|
|---|---|
|[<span data-ttu-id="7064a-123">Mindestversion des Postfachanforderungssatzes</span><span class="sxs-lookup"><span data-stu-id="7064a-123">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="7064a-124">1.0</span><span class="sxs-lookup"><span data-stu-id="7064a-124">1.0</span></span>|
|[<span data-ttu-id="7064a-125">Mindestberechtigungsstufe</span><span class="sxs-lookup"><span data-stu-id="7064a-125">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="7064a-126">ReadItem</span><span class="sxs-lookup"><span data-stu-id="7064a-126">ReadItem</span></span>|
|[<span data-ttu-id="7064a-127">Zutreffender Outlook-Modus</span><span class="sxs-lookup"><span data-stu-id="7064a-127">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="7064a-128">Verfassen oder Lesen</span><span class="sxs-lookup"><span data-stu-id="7064a-128">Compose or read</span></span>|

####  <a name="hostversion-string"></a><span data-ttu-id="7064a-129">hostVersion :String</span><span class="sxs-lookup"><span data-stu-id="7064a-129">hostVersion :String</span></span>

<span data-ttu-id="7064a-130">Ruft eine Zeichenfolge ab, die entweder die Version der Hostanwendung oder des Exchange-Servers darstellt.</span><span class="sxs-lookup"><span data-stu-id="7064a-130">Gets a string that represents the version of either the host application or the Exchange Server.</span></span>

<span data-ttu-id="7064a-p102">Wenn das E-Mail-Add-In im Outlook-Desktopclient oder Outlook für iOS ausgeführt wird, gibt die `hostVersion`-Eigenschaft die Version der Hostanwendung zurück. In Outlook Web App gibt die Eigenschaft die Version von Exchange Server zurück. Ein Beispiel ist die Zeichenfolge `15.0.468.0`.</span><span class="sxs-lookup"><span data-stu-id="7064a-p102">If the mail add-in is running on the Outlook desktop client or Outlook for iOS, the `hostVersion` property returns the version of the host application, Outlook. In Outlook Web App, the property returns the version of the Exchange Server. An example is the string `15.0.468.0`.</span></span>

##### <a name="type"></a><span data-ttu-id="7064a-134">Typ:</span><span class="sxs-lookup"><span data-stu-id="7064a-134">Type:</span></span>

*   <span data-ttu-id="7064a-135">String</span><span class="sxs-lookup"><span data-stu-id="7064a-135">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="7064a-136">Anforderungen</span><span class="sxs-lookup"><span data-stu-id="7064a-136">Requirements</span></span>

|<span data-ttu-id="7064a-137">Anforderung</span><span class="sxs-lookup"><span data-stu-id="7064a-137">Requirement</span></span>| <span data-ttu-id="7064a-138">Wert</span><span class="sxs-lookup"><span data-stu-id="7064a-138">Value</span></span>|
|---|---|
|[<span data-ttu-id="7064a-139">Mindestversion des Postfachanforderungssatzes</span><span class="sxs-lookup"><span data-stu-id="7064a-139">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="7064a-140">1.0</span><span class="sxs-lookup"><span data-stu-id="7064a-140">1.0</span></span>|
|[<span data-ttu-id="7064a-141">Mindestberechtigungsstufe</span><span class="sxs-lookup"><span data-stu-id="7064a-141">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="7064a-142">ReadItem</span><span class="sxs-lookup"><span data-stu-id="7064a-142">ReadItem</span></span>|
|[<span data-ttu-id="7064a-143">Zutreffender Outlook-Modus</span><span class="sxs-lookup"><span data-stu-id="7064a-143">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="7064a-144">Verfassen oder Lesen</span><span class="sxs-lookup"><span data-stu-id="7064a-144">Compose or read</span></span>|

####  <a name="owaview-string"></a><span data-ttu-id="7064a-145">OWAView :String</span><span class="sxs-lookup"><span data-stu-id="7064a-145">OWAView :String</span></span>

<span data-ttu-id="7064a-146">Ruft eine Zeichenfolge ab, die die aktuelle Ansicht von Outlook Web App darstellt.</span><span class="sxs-lookup"><span data-stu-id="7064a-146">Gets a string that represents the current view of Outlook Web App.</span></span>

<span data-ttu-id="7064a-147">Die zurückgegebene Zeichenfolge kann einen der folgenden Werte haben: `OneColumn`, `TwoColumns` oder `ThreeColumns`.</span><span class="sxs-lookup"><span data-stu-id="7064a-147">The returned string can be one of the following values: `OneColumn`, `TwoColumns`, or `ThreeColumns`.</span></span>

<span data-ttu-id="7064a-148">Handelt es sich bei der Hostanwendung nicht um Outlook Web App, wird durch den Zugriff auf diese Eigenschaft `undefined` zurückgegeben.</span><span class="sxs-lookup"><span data-stu-id="7064a-148">If the host application is not Outlook Web App, then accessing this property results in `undefined`.</span></span>

<span data-ttu-id="7064a-149">Outlook Web App verfügt über drei Ansichten, die der Bildschirm- und Fensterbreite sowie mit der Anzahl der Spalten übereinstimmt, die angezeigt werden können:</span><span class="sxs-lookup"><span data-stu-id="7064a-149">Outlook Web App has three views that correspond to the width of the screen and the window, and the number of columns that can be displayed:</span></span>

*   <span data-ttu-id="7064a-p103">`OneColumn` wird bei schmalem Bildschirm angezeigt. Outlook Web App verwendet dieses einspaltige Layout auf dem gesamten Bildschirm eines Smartphones.</span><span class="sxs-lookup"><span data-stu-id="7064a-p103">`OneColumn`, which is displayed when the screen is narrow. Outlook Web App uses this single-column layout on the entire screen of a smartphone.</span></span>
*   <span data-ttu-id="7064a-p104">`TwoColumns` wird bei etwas breiterem Bildschirm angezeigt. Outlook Web App verwendet diese Ansicht auf den meisten Tablet-PCs.</span><span class="sxs-lookup"><span data-stu-id="7064a-p104">`TwoColumns`, which is displayed when the screen is wider. Outlook Web App uses this view on most tablets.</span></span>
*   <span data-ttu-id="7064a-p105">`ThreeColumns` wird bei breitem Bildschirm angezeigt. Outlook Web App verwendet diese Ansicht z. B. in einem Fenster mit voller Bildschirmgröße auf einem Desktopcomputer.</span><span class="sxs-lookup"><span data-stu-id="7064a-p105">`ThreeColumns`, which is displayed when the screen is wide. For example, Outlook Web App uses this view in a full screen window on a desktop computer.</span></span>

##### <a name="type"></a><span data-ttu-id="7064a-156">Typ:</span><span class="sxs-lookup"><span data-stu-id="7064a-156">Type:</span></span>

*   <span data-ttu-id="7064a-157">String</span><span class="sxs-lookup"><span data-stu-id="7064a-157">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="7064a-158">Anforderungen</span><span class="sxs-lookup"><span data-stu-id="7064a-158">Requirements</span></span>

|<span data-ttu-id="7064a-159">Anforderung</span><span class="sxs-lookup"><span data-stu-id="7064a-159">Requirement</span></span>| <span data-ttu-id="7064a-160">Wert</span><span class="sxs-lookup"><span data-stu-id="7064a-160">Value</span></span>|
|---|---|
|[<span data-ttu-id="7064a-161">Mindestversion des Postfachanforderungssatzes</span><span class="sxs-lookup"><span data-stu-id="7064a-161">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="7064a-162">1.0</span><span class="sxs-lookup"><span data-stu-id="7064a-162">1.0</span></span>|
|[<span data-ttu-id="7064a-163">Mindestberechtigungsstufe</span><span class="sxs-lookup"><span data-stu-id="7064a-163">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="7064a-164">ReadItem</span><span class="sxs-lookup"><span data-stu-id="7064a-164">ReadItem</span></span>|
|[<span data-ttu-id="7064a-165">Zutreffender Outlook-Modus</span><span class="sxs-lookup"><span data-stu-id="7064a-165">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="7064a-166">Verfassen oder Lesen</span><span class="sxs-lookup"><span data-stu-id="7064a-166">Compose or read</span></span>|