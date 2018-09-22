
# <a name="diagnostics"></a><span data-ttu-id="31a48-101">diagnostics</span><span class="sxs-lookup"><span data-stu-id="31a48-101">diagnostics</span></span>

### <span data-ttu-id="31a48-p101">[Office](Office.md)[.context](Office.context.md)[.mailbox](Office.context.mailbox.md). diagnostics</span><span class="sxs-lookup"><span data-stu-id="31a48-p101">[Office](Office.md)[.context](Office.context.md)[.mailbox](Office.context.mailbox.md). diagnostics</span></span>

<span data-ttu-id="31a48-104">Stellt einem Outlook-Add-In Diagnoseinformationen bereit.</span><span class="sxs-lookup"><span data-stu-id="31a48-104">Provides diagnostic information to an Outlook add-in.</span></span>

##### <a name="requirements"></a><span data-ttu-id="31a48-105">Anforderungen</span><span class="sxs-lookup"><span data-stu-id="31a48-105">Requirements</span></span>

|<span data-ttu-id="31a48-106">Anforderung</span><span class="sxs-lookup"><span data-stu-id="31a48-106">Requirement</span></span>| <span data-ttu-id="31a48-107">Wert</span><span class="sxs-lookup"><span data-stu-id="31a48-107">Value</span></span>|
|---|---|
|[<span data-ttu-id="31a48-108">Mindestversion des Postfachanforderungssatzes</span><span class="sxs-lookup"><span data-stu-id="31a48-108">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="31a48-109">1.0</span><span class="sxs-lookup"><span data-stu-id="31a48-109">1.0</span></span>|
|[<span data-ttu-id="31a48-110">Mindestberechtigungsstufe</span><span class="sxs-lookup"><span data-stu-id="31a48-110">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="31a48-111">ReadItem</span><span class="sxs-lookup"><span data-stu-id="31a48-111">ReadItem</span></span>|
|[<span data-ttu-id="31a48-112">Zutreffender Outlook-Modus</span><span class="sxs-lookup"><span data-stu-id="31a48-112">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="31a48-113">Verfassen oder Lesen</span><span class="sxs-lookup"><span data-stu-id="31a48-113">Compose or read</span></span>|

##### <a name="members-and-methods"></a><span data-ttu-id="31a48-114">Elemente und Methoden</span><span class="sxs-lookup"><span data-stu-id="31a48-114">Members and methods</span></span>

| <span data-ttu-id="31a48-115">Element</span><span class="sxs-lookup"><span data-stu-id="31a48-115">Member</span></span> | <span data-ttu-id="31a48-116">Typ</span><span class="sxs-lookup"><span data-stu-id="31a48-116">Type</span></span> |
|--------|------|
| [<span data-ttu-id="31a48-117">hostName</span><span class="sxs-lookup"><span data-stu-id="31a48-117">hostName</span></span>](#hostname-string) | <span data-ttu-id="31a48-118">Element</span><span class="sxs-lookup"><span data-stu-id="31a48-118">Member</span></span> |
| [<span data-ttu-id="31a48-119">hostVersion</span><span class="sxs-lookup"><span data-stu-id="31a48-119">hostVersion</span></span>](#hostversion-string) | <span data-ttu-id="31a48-120">Element</span><span class="sxs-lookup"><span data-stu-id="31a48-120">Member</span></span> |
| [<span data-ttu-id="31a48-121">OWAView</span><span class="sxs-lookup"><span data-stu-id="31a48-121">OWAView</span></span>](#owaview-string) | <span data-ttu-id="31a48-122">Element</span><span class="sxs-lookup"><span data-stu-id="31a48-122">Member</span></span> |

### <a name="members"></a><span data-ttu-id="31a48-123">Elemente</span><span class="sxs-lookup"><span data-stu-id="31a48-123">Members</span></span>

####  <a name="hostname-string"></a><span data-ttu-id="31a48-124">hostName :String</span><span class="sxs-lookup"><span data-stu-id="31a48-124">hostName :String</span></span>

<span data-ttu-id="31a48-125">Ruft eine Zeichenfolge ab, die den Namen der Hostanwendung darstellt.</span><span class="sxs-lookup"><span data-stu-id="31a48-125">Gets a string that represents the name of the host application.</span></span>

<span data-ttu-id="31a48-126">Eine Zeichenfolge mit einem der folgenden Werte: `Outlook`, `Mac Outlook`, `OutlookIOS` oder `OutlookWebApp`.</span><span class="sxs-lookup"><span data-stu-id="31a48-126">A string that can be one of the following values: `Outlook`, `Mac Outlook`, `OutlookIOS`, or `OutlookWebApp`.</span></span>

##### <a name="type"></a><span data-ttu-id="31a48-127">Typ:</span><span class="sxs-lookup"><span data-stu-id="31a48-127">Type:</span></span>

*   <span data-ttu-id="31a48-128">String</span><span class="sxs-lookup"><span data-stu-id="31a48-128">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="31a48-129">Anforderungen</span><span class="sxs-lookup"><span data-stu-id="31a48-129">Requirements</span></span>

|<span data-ttu-id="31a48-130">Anforderung</span><span class="sxs-lookup"><span data-stu-id="31a48-130">Requirement</span></span>| <span data-ttu-id="31a48-131">Wert</span><span class="sxs-lookup"><span data-stu-id="31a48-131">Value</span></span>|
|---|---|
|[<span data-ttu-id="31a48-132">Mindestversion des Postfachanforderungssatzes</span><span class="sxs-lookup"><span data-stu-id="31a48-132">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="31a48-133">1.0</span><span class="sxs-lookup"><span data-stu-id="31a48-133">1.0</span></span>|
|[<span data-ttu-id="31a48-134">Mindestberechtigungsstufe</span><span class="sxs-lookup"><span data-stu-id="31a48-134">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="31a48-135">ReadItem</span><span class="sxs-lookup"><span data-stu-id="31a48-135">ReadItem</span></span>|
|[<span data-ttu-id="31a48-136">Zutreffender Outlook-Modus</span><span class="sxs-lookup"><span data-stu-id="31a48-136">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="31a48-137">Verfassen oder Lesen</span><span class="sxs-lookup"><span data-stu-id="31a48-137">Compose or read</span></span>|

####  <a name="hostversion-string"></a><span data-ttu-id="31a48-138">hostVersion :String</span><span class="sxs-lookup"><span data-stu-id="31a48-138">hostVersion :String</span></span>

<span data-ttu-id="31a48-139">Ruft eine Zeichenfolge ab, die entweder die Version der Hostanwendung oder des Exchange-Servers darstellt.</span><span class="sxs-lookup"><span data-stu-id="31a48-139">Gets a string that represents the version of either the host application or the Exchange Server.</span></span>

<span data-ttu-id="31a48-p102">Wenn das E-Mail-Add-In im Outlook-Desktopclient oder Outlook für iOS ausgeführt wird, gibt die `hostVersion`-Eigenschaft die Version der Hostanwendung zurück. In Outlook Web App gibt die Eigenschaft die Version von Exchange Server zurück. Ein Beispiel ist die Zeichenfolge `15.0.468.0`.</span><span class="sxs-lookup"><span data-stu-id="31a48-p102">If the mail add-in is running on the Outlook desktop client or Outlook for iOS, the `hostVersion` property returns the version of the host application, Outlook. In Outlook Web App, the property returns the version of the Exchange Server. An example is the string `15.0.468.0`.</span></span>

##### <a name="type"></a><span data-ttu-id="31a48-143">Typ:</span><span class="sxs-lookup"><span data-stu-id="31a48-143">Type:</span></span>

*   <span data-ttu-id="31a48-144">String</span><span class="sxs-lookup"><span data-stu-id="31a48-144">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="31a48-145">Anforderungen</span><span class="sxs-lookup"><span data-stu-id="31a48-145">Requirements</span></span>

|<span data-ttu-id="31a48-146">Anforderung</span><span class="sxs-lookup"><span data-stu-id="31a48-146">Requirement</span></span>| <span data-ttu-id="31a48-147">Wert</span><span class="sxs-lookup"><span data-stu-id="31a48-147">Value</span></span>|
|---|---|
|[<span data-ttu-id="31a48-148">Mindestversion des Postfachanforderungssatzes</span><span class="sxs-lookup"><span data-stu-id="31a48-148">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="31a48-149">1.0</span><span class="sxs-lookup"><span data-stu-id="31a48-149">1.0</span></span>|
|[<span data-ttu-id="31a48-150">Mindestberechtigungsstufe</span><span class="sxs-lookup"><span data-stu-id="31a48-150">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="31a48-151">ReadItem</span><span class="sxs-lookup"><span data-stu-id="31a48-151">ReadItem</span></span>|
|[<span data-ttu-id="31a48-152">Zutreffender Outlook-Modus</span><span class="sxs-lookup"><span data-stu-id="31a48-152">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="31a48-153">Verfassen oder Lesen</span><span class="sxs-lookup"><span data-stu-id="31a48-153">Compose or read</span></span>|

####  <a name="owaview-string"></a><span data-ttu-id="31a48-154">OWAView :String</span><span class="sxs-lookup"><span data-stu-id="31a48-154">OWAView :String</span></span>

<span data-ttu-id="31a48-155">Ruft eine Zeichenfolge ab, die die aktuelle Ansicht von Outlook Web App darstellt.</span><span class="sxs-lookup"><span data-stu-id="31a48-155">Gets a string that represents the current view of Outlook Web App.</span></span>

<span data-ttu-id="31a48-156">Die zurückgegebene Zeichenfolge kann einen der folgenden Werte haben: `OneColumn`, `TwoColumns` oder `ThreeColumns`.</span><span class="sxs-lookup"><span data-stu-id="31a48-156">The returned string can be one of the following values: `OneColumn`, `TwoColumns`, or `ThreeColumns`.</span></span>

<span data-ttu-id="31a48-157">Handelt es sich bei der Hostanwendung nicht um Outlook Web App, wird durch den Zugriff auf diese Eigenschaft `undefined` zurückgegeben.</span><span class="sxs-lookup"><span data-stu-id="31a48-157">If the host application is not Outlook Web App, then accessing this property results in `undefined`.</span></span>

<span data-ttu-id="31a48-158">Outlook Web App verfügt über drei Ansichten, die der Bildschirm- und Fensterbreite sowie mit der Anzahl der Spalten übereinstimmt, die angezeigt werden können:</span><span class="sxs-lookup"><span data-stu-id="31a48-158">Outlook Web App has three views that correspond to the width of the screen and the window, and the number of columns that can be displayed:</span></span>

*   <span data-ttu-id="31a48-p103">`OneColumn` wird bei schmalem Bildschirm angezeigt. Outlook Web App verwendet dieses einspaltige Layout auf dem gesamten Bildschirm eines Smartphones.</span><span class="sxs-lookup"><span data-stu-id="31a48-p103">`OneColumn`, which is displayed when the screen is narrow. Outlook Web App uses this single-column layout on the entire screen of a smartphone.</span></span>
*   <span data-ttu-id="31a48-p104">`TwoColumns` wird bei etwas breiterem Bildschirm angezeigt. Outlook Web App verwendet diese Ansicht auf den meisten Tablet-PCs.</span><span class="sxs-lookup"><span data-stu-id="31a48-p104">`TwoColumns`, which is displayed when the screen is wider. Outlook Web App uses this view on most tablets.</span></span>
*   <span data-ttu-id="31a48-p105">`ThreeColumns` wird bei breitem Bildschirm angezeigt. Outlook Web App verwendet diese Ansicht z. B. in einem Fenster mit voller Bildschirmgröße auf einem Desktopcomputer.</span><span class="sxs-lookup"><span data-stu-id="31a48-p105">`ThreeColumns`, which is displayed when the screen is wide. For example, Outlook Web App uses this view in a full screen window on a desktop computer.</span></span>

##### <a name="type"></a><span data-ttu-id="31a48-165">Typ:</span><span class="sxs-lookup"><span data-stu-id="31a48-165">Type:</span></span>

*   <span data-ttu-id="31a48-166">String</span><span class="sxs-lookup"><span data-stu-id="31a48-166">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="31a48-167">Anforderungen</span><span class="sxs-lookup"><span data-stu-id="31a48-167">Requirements</span></span>

|<span data-ttu-id="31a48-168">Anforderung</span><span class="sxs-lookup"><span data-stu-id="31a48-168">Requirement</span></span>| <span data-ttu-id="31a48-169">Wert</span><span class="sxs-lookup"><span data-stu-id="31a48-169">Value</span></span>|
|---|---|
|[<span data-ttu-id="31a48-170">Mindestversion des Postfachanforderungssatzes</span><span class="sxs-lookup"><span data-stu-id="31a48-170">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="31a48-171">1.0</span><span class="sxs-lookup"><span data-stu-id="31a48-171">1.0</span></span>|
|[<span data-ttu-id="31a48-172">Mindestberechtigungsstufe</span><span class="sxs-lookup"><span data-stu-id="31a48-172">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="31a48-173">ReadItem</span><span class="sxs-lookup"><span data-stu-id="31a48-173">ReadItem</span></span>|
|[<span data-ttu-id="31a48-174">Zutreffender Outlook-Modus</span><span class="sxs-lookup"><span data-stu-id="31a48-174">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="31a48-175">Verfassen oder Lesen</span><span class="sxs-lookup"><span data-stu-id="31a48-175">Compose or read</span></span>|