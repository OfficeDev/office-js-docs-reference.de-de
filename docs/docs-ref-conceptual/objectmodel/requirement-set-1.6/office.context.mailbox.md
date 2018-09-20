
# <a name="mailbox"></a><span data-ttu-id="3aa16-101">mailbox</span><span class="sxs-lookup"><span data-stu-id="3aa16-101">mailbox</span></span>

### <span data-ttu-id="3aa16-p101">[Office](Office.md)[.context](Office.context.md). mailbox</span><span class="sxs-lookup"><span data-stu-id="3aa16-p101">[Office](Office.md)[.context](Office.context.md). mailbox</span></span>

<span data-ttu-id="3aa16-104">Ermöglicht den Zugriff auf das Objektmodell von Outlook-add-in für Microsoft Outlook und Microsoft Outlook im Web.</span><span class="sxs-lookup"><span data-stu-id="3aa16-104">Provides access to the Outlook add-in object model for Microsoft Outlook and Microsoft Outlook on the web.</span></span>

##### <a name="requirements"></a><span data-ttu-id="3aa16-105">Anforderungen</span><span class="sxs-lookup"><span data-stu-id="3aa16-105">Requirements</span></span>

|<span data-ttu-id="3aa16-106">Anforderung</span><span class="sxs-lookup"><span data-stu-id="3aa16-106">Requirement</span></span>| <span data-ttu-id="3aa16-107">Wert</span><span class="sxs-lookup"><span data-stu-id="3aa16-107">Value</span></span>|
|---|---|
|[<span data-ttu-id="3aa16-108">Mindestversion des Postfachanforderungssatzes</span><span class="sxs-lookup"><span data-stu-id="3aa16-108">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="3aa16-109">1.0</span><span class="sxs-lookup"><span data-stu-id="3aa16-109">1.0</span></span>|
|[<span data-ttu-id="3aa16-110">Mindestberechtigungsstufe</span><span class="sxs-lookup"><span data-stu-id="3aa16-110">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="3aa16-111">Eingeschränkt</span><span class="sxs-lookup"><span data-stu-id="3aa16-111">Restricted</span></span>|
|[<span data-ttu-id="3aa16-112">Zutreffender Outlook-Modus</span><span class="sxs-lookup"><span data-stu-id="3aa16-112">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="3aa16-113">Verfassen oder Lesen</span><span class="sxs-lookup"><span data-stu-id="3aa16-113">Compose or read</span></span>|

##### <a name="members-and-methods"></a><span data-ttu-id="3aa16-114">Elemente und Methoden</span><span class="sxs-lookup"><span data-stu-id="3aa16-114">Members and methods</span></span>

| <span data-ttu-id="3aa16-115">Element</span><span class="sxs-lookup"><span data-stu-id="3aa16-115">Member</span></span> | <span data-ttu-id="3aa16-116">Typ</span><span class="sxs-lookup"><span data-stu-id="3aa16-116">Type</span></span> |
|--------|------|
| [<span data-ttu-id="3aa16-117">ewsUrl</span><span class="sxs-lookup"><span data-stu-id="3aa16-117">ewsUrl</span></span>](#ewsurl-string) | <span data-ttu-id="3aa16-118">Element</span><span class="sxs-lookup"><span data-stu-id="3aa16-118">Member</span></span> |
| [<span data-ttu-id="3aa16-119">restUrl</span><span class="sxs-lookup"><span data-stu-id="3aa16-119">restUrl</span></span>](#resturl-string) | <span data-ttu-id="3aa16-120">Element</span><span class="sxs-lookup"><span data-stu-id="3aa16-120">Member</span></span> |
| [<span data-ttu-id="3aa16-121">addHandlerAsync</span><span class="sxs-lookup"><span data-stu-id="3aa16-121">addHandlerAsync</span></span>](#addhandlerasynceventtype-handler-options-callback) | <span data-ttu-id="3aa16-122">Methode</span><span class="sxs-lookup"><span data-stu-id="3aa16-122">Method</span></span> |
| [<span data-ttu-id="3aa16-123">convertToEwsId</span><span class="sxs-lookup"><span data-stu-id="3aa16-123">convertToEwsId</span></span>](#converttoewsiditemid-restversion--string) | <span data-ttu-id="3aa16-124">Methode</span><span class="sxs-lookup"><span data-stu-id="3aa16-124">Method</span></span> |
| [<span data-ttu-id="3aa16-125">convertToLocalClientTime</span><span class="sxs-lookup"><span data-stu-id="3aa16-125">convertToLocalClientTime</span></span>](#converttolocalclienttimetimevalue--localclienttimejavascriptapioutlook16officelocalclienttime) | <span data-ttu-id="3aa16-126">Methode</span><span class="sxs-lookup"><span data-stu-id="3aa16-126">Method</span></span> |
| [<span data-ttu-id="3aa16-127">convertToRestId</span><span class="sxs-lookup"><span data-stu-id="3aa16-127">convertToRestId</span></span>](#converttorestiditemid-restversion--string) | <span data-ttu-id="3aa16-128">Methode</span><span class="sxs-lookup"><span data-stu-id="3aa16-128">Method</span></span> |
| [<span data-ttu-id="3aa16-129">convertToUtcClientTime</span><span class="sxs-lookup"><span data-stu-id="3aa16-129">convertToUtcClientTime</span></span>](#converttoutcclienttimeinput--date) | <span data-ttu-id="3aa16-130">Methode</span><span class="sxs-lookup"><span data-stu-id="3aa16-130">Method</span></span> |
| [<span data-ttu-id="3aa16-131">displayAppointmentForm</span><span class="sxs-lookup"><span data-stu-id="3aa16-131">displayAppointmentForm</span></span>](#displayappointmentformitemid) | <span data-ttu-id="3aa16-132">Methode</span><span class="sxs-lookup"><span data-stu-id="3aa16-132">Method</span></span> |
| [<span data-ttu-id="3aa16-133">displayMessageForm</span><span class="sxs-lookup"><span data-stu-id="3aa16-133">displayMessageForm</span></span>](#displaymessageformitemid) | <span data-ttu-id="3aa16-134">Methode</span><span class="sxs-lookup"><span data-stu-id="3aa16-134">Method</span></span> |
| [<span data-ttu-id="3aa16-135">displayNewAppointmentForm</span><span class="sxs-lookup"><span data-stu-id="3aa16-135">displayNewAppointmentForm</span></span>](#displaynewappointmentformparameters) | <span data-ttu-id="3aa16-136">Methode</span><span class="sxs-lookup"><span data-stu-id="3aa16-136">Method</span></span> |
| [<span data-ttu-id="3aa16-137">displayNewMessageForm</span><span class="sxs-lookup"><span data-stu-id="3aa16-137">displayNewMessageForm</span></span>](#displaynewmessageformparameters) | <span data-ttu-id="3aa16-138">Methode</span><span class="sxs-lookup"><span data-stu-id="3aa16-138">Method</span></span> |
| [<span data-ttu-id="3aa16-139">getCallbackTokenAsync</span><span class="sxs-lookup"><span data-stu-id="3aa16-139">getCallbackTokenAsync</span></span>](#getcallbacktokenasyncoptions-callback) | <span data-ttu-id="3aa16-140">Methode</span><span class="sxs-lookup"><span data-stu-id="3aa16-140">Method</span></span> |
| [<span data-ttu-id="3aa16-141">getCallbackTokenAsync</span><span class="sxs-lookup"><span data-stu-id="3aa16-141">getCallbackTokenAsync</span></span>](#getcallbacktokenasynccallback-usercontext) | <span data-ttu-id="3aa16-142">Methode</span><span class="sxs-lookup"><span data-stu-id="3aa16-142">Method</span></span> |
| [<span data-ttu-id="3aa16-143">getUserIdentityTokenAsync</span><span class="sxs-lookup"><span data-stu-id="3aa16-143">getUserIdentityTokenAsync</span></span>](#getuseridentitytokenasynccallback-usercontext) | <span data-ttu-id="3aa16-144">Methode</span><span class="sxs-lookup"><span data-stu-id="3aa16-144">Method</span></span> |
| [<span data-ttu-id="3aa16-145">makeEwsRequestAsync</span><span class="sxs-lookup"><span data-stu-id="3aa16-145">makeEwsRequestAsync</span></span>](#makeewsrequestasyncdata-callback-usercontext) | <span data-ttu-id="3aa16-146">Methode</span><span class="sxs-lookup"><span data-stu-id="3aa16-146">Method</span></span> |

### <a name="namespaces"></a><span data-ttu-id="3aa16-147">Namespaces</span><span class="sxs-lookup"><span data-stu-id="3aa16-147">Namespaces</span></span>

<span data-ttu-id="3aa16-148">[diagnostics](Office.context.mailbox.diagnostics.md): Stellt einem Outlook-Add-In Diagnoseinformationen bereit.</span><span class="sxs-lookup"><span data-stu-id="3aa16-148">[diagnostics](Office.context.mailbox.diagnostics.md): Provides diagnostic information to an Outlook add-in.</span></span>

<span data-ttu-id="3aa16-149">[item](Office.context.mailbox.item.md): Stellt Methoden und Eigenschaften für den Zugriff auf eine Nachricht oder einen Termins in einem Outlook-Add-In bereit.</span><span class="sxs-lookup"><span data-stu-id="3aa16-149">[item](Office.context.mailbox.item.md): Provides methods and properties for accessing a message or appointment in an Outlook add-in.</span></span>

<span data-ttu-id="3aa16-150">[userProfile](Office.context.mailbox.userProfile.md): Stellt Informationen zum Benutzer in einem Outlook-Add-In bereit.</span><span class="sxs-lookup"><span data-stu-id="3aa16-150">[userProfile](Office.context.mailbox.userProfile.md): Provides information about the user in an Outlook add-in.</span></span>

### <a name="members"></a><span data-ttu-id="3aa16-151">Elemente</span><span class="sxs-lookup"><span data-stu-id="3aa16-151">Members</span></span>

#### <a name="ewsurl-string"></a><span data-ttu-id="3aa16-152">ewsUrl :String</span><span class="sxs-lookup"><span data-stu-id="3aa16-152">ewsUrl :String</span></span>

<span data-ttu-id="3aa16-p102">Ruft die URL des EWS-Endpunkts (Exchange Web Services) für dieses E-Mail-Konto ab. Nur Lesemodus.</span><span class="sxs-lookup"><span data-stu-id="3aa16-p102">Gets the URL of the Exchange Web Services (EWS) endpoint for this email account. Read mode only.</span></span>

> [!NOTE]
> <span data-ttu-id="3aa16-155">Dieser Member wird in Outlook für iOS oder Outlook für Android nicht unterstützt.</span><span class="sxs-lookup"><span data-stu-id="3aa16-155">This member is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="3aa16-p103">Der `ewsUrl`-Wert kann durch einen Remotedienst zum Ausführen von EWS-Aufrufen an das Postfach des Benutzers verwendet werden. Sie können z. B. einen Remotedienst erstellen, um [Anlagen](https://docs.microsoft.com/outlook/add-ins/get-attachments-of-an-outlook-item) aus dem ausgewählten Element abzurufen.</span><span class="sxs-lookup"><span data-stu-id="3aa16-p103">The `ewsUrl` value can be used by a remote service to make EWS calls to the user's mailbox. For example, you can create a remote service to [get attachments from the selected item](https://docs.microsoft.com/outlook/add-ins/get-attachments-of-an-outlook-item).</span></span>

<span data-ttu-id="3aa16-158">Ihre App muss über die im Manifest angegebene **ReadItem**-Berechtigung verfügen, um das `ewsUrl`-Element im Lesemodus aufzurufen.</span><span class="sxs-lookup"><span data-stu-id="3aa16-158">Your app must have the **ReadItem** permission specified in its manifest to call the `ewsUrl` member in read mode.</span></span>

<span data-ttu-id="3aa16-p104">Im Verfassenmodus müssen Sie die [`saveAsync`](Office.context.mailbox.item.md#saveasyncoptions-callback)-Methode aufrufen, bevor Sie das `ewsUrl`-Element verwenden können. Die App muss über **ReadWriteItem**-Berechtigungen verfügen, um die `saveAsync`-Methode aufzurufen.</span><span class="sxs-lookup"><span data-stu-id="3aa16-p104">In compose mode you must call the [`saveAsync`](Office.context.mailbox.item.md#saveasyncoptions-callback) method before you can use the `ewsUrl` member. Your app must have **ReadWriteItem** permissions to call the `saveAsync` method.</span></span>

##### <a name="type"></a><span data-ttu-id="3aa16-161">Typ:</span><span class="sxs-lookup"><span data-stu-id="3aa16-161">Type:</span></span>

*   <span data-ttu-id="3aa16-162">String</span><span class="sxs-lookup"><span data-stu-id="3aa16-162">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="3aa16-163">Anforderungen</span><span class="sxs-lookup"><span data-stu-id="3aa16-163">Requirements</span></span>

|<span data-ttu-id="3aa16-164">Anforderung</span><span class="sxs-lookup"><span data-stu-id="3aa16-164">Requirement</span></span>| <span data-ttu-id="3aa16-165">Wert</span><span class="sxs-lookup"><span data-stu-id="3aa16-165">Value</span></span>|
|---|---|
|[<span data-ttu-id="3aa16-166">Mindestversion des Postfachanforderungssatzes</span><span class="sxs-lookup"><span data-stu-id="3aa16-166">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="3aa16-167">1.0</span><span class="sxs-lookup"><span data-stu-id="3aa16-167">1.0</span></span>|
|[<span data-ttu-id="3aa16-168">Mindestberechtigungsstufe</span><span class="sxs-lookup"><span data-stu-id="3aa16-168">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="3aa16-169">ReadItem</span><span class="sxs-lookup"><span data-stu-id="3aa16-169">ReadItem</span></span>|
|[<span data-ttu-id="3aa16-170">Zutreffender Outlook-Modus</span><span class="sxs-lookup"><span data-stu-id="3aa16-170">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="3aa16-171">Verfassen oder Lesen</span><span class="sxs-lookup"><span data-stu-id="3aa16-171">Compose or read</span></span>|

#### <a name="resturl-string"></a><span data-ttu-id="3aa16-172">restUrl :String</span><span class="sxs-lookup"><span data-stu-id="3aa16-172">restUrl :String</span></span>

<span data-ttu-id="3aa16-173">Ruft die URL des REST-Endpunkts für das betreffende E-Mail-Konto ab.</span><span class="sxs-lookup"><span data-stu-id="3aa16-173">Gets the URL of the REST endpoint for this email account.</span></span>

<span data-ttu-id="3aa16-174">Der Wert `restUrl` kann für [REST-API](https://docs.microsoft.com/outlook/rest/)-Aufrufe an das Postfach des Benutzers verwendet werden.</span><span class="sxs-lookup"><span data-stu-id="3aa16-174">The `restUrl` value can be used to make [REST API](https://docs.microsoft.com/outlook/rest/) calls to the user's mailbox.</span></span>

<span data-ttu-id="3aa16-175">Im Manifest der App muss die Berechtigung **ReadItem** angegeben sein, damit das Mitglied `restUrl` im Lesemodus abgerufen werden kann.</span><span class="sxs-lookup"><span data-stu-id="3aa16-175">Your app must have the **ReadItem** permission specified in its manifest to call the `restUrl` member in read mode.</span></span>

<span data-ttu-id="3aa16-p105">Im Verfassenmodus müssen Sie die [`saveAsync`](Office.context.mailbox.item.md#saveasyncoptions-callback)-Methode aufrufen, bevor Sie das `restUrl`-Element verwenden können. Die App muss über **ReadWriteItem**-Berechtigungen verfügen, um die `saveAsync`-Methode aufzurufen.</span><span class="sxs-lookup"><span data-stu-id="3aa16-p105">In compose mode you must call the [`saveAsync`](Office.context.mailbox.item.md#saveasyncoptions-callback) method before you can use the `restUrl` member. Your app must have **ReadWriteItem** permissions to call the `saveAsync` method.</span></span>

##### <a name="type"></a><span data-ttu-id="3aa16-178">Typ:</span><span class="sxs-lookup"><span data-stu-id="3aa16-178">Type:</span></span>

*   <span data-ttu-id="3aa16-179">String</span><span class="sxs-lookup"><span data-stu-id="3aa16-179">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="3aa16-180">Anforderungen</span><span class="sxs-lookup"><span data-stu-id="3aa16-180">Requirements</span></span>

|<span data-ttu-id="3aa16-181">Anforderung</span><span class="sxs-lookup"><span data-stu-id="3aa16-181">Requirement</span></span>| <span data-ttu-id="3aa16-182">Wert</span><span class="sxs-lookup"><span data-stu-id="3aa16-182">Value</span></span>|
|---|---|
|[<span data-ttu-id="3aa16-183">Mindestversion des Postfachanforderungssatzes</span><span class="sxs-lookup"><span data-stu-id="3aa16-183">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="3aa16-184">1,5</span><span class="sxs-lookup"><span data-stu-id="3aa16-184">1.5</span></span> |
|[<span data-ttu-id="3aa16-185">Mindestberechtigungsstufe</span><span class="sxs-lookup"><span data-stu-id="3aa16-185">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="3aa16-186">ReadItem</span><span class="sxs-lookup"><span data-stu-id="3aa16-186">ReadItem</span></span>|
|[<span data-ttu-id="3aa16-187">Zutreffender Outlook-Modus</span><span class="sxs-lookup"><span data-stu-id="3aa16-187">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="3aa16-188">Verfassen oder Lesen</span><span class="sxs-lookup"><span data-stu-id="3aa16-188">Compose or read</span></span>|

### <a name="methods"></a><span data-ttu-id="3aa16-189">Methoden</span><span class="sxs-lookup"><span data-stu-id="3aa16-189">Methods</span></span>

####  <a name="addhandlerasynceventtype-handler-options-callback"></a><span data-ttu-id="3aa16-190">addHandlerAsync(eventType, handler, [options], [callback])</span><span class="sxs-lookup"><span data-stu-id="3aa16-190">addHandlerAsync(eventType, handler, [options], [callback])</span></span>

<span data-ttu-id="3aa16-191">Fügt einen Ereignishandler für ein unterstütztes Ereignis hinzu.</span><span class="sxs-lookup"><span data-stu-id="3aa16-191">Adds an event handler for a supported event.</span></span>

<span data-ttu-id="3aa16-p106">Derzeit wird als Ereignistyp ausschließlich `Office.EventType.ItemChanged` unterstützt. Er wird aufgerufen, wenn der Benutzer ein neues Element auswählt. Dieses Ereignis wird von Add-Ins verwendet, die einen anheftbaren Aufgabenbereich implementieren. Es erlaubt es dem Add-In, die Benutzeroberfläche des Aufgabenbereichs basierend auf dem aktuell ausgewählten Element zu aktualisieren.</span><span class="sxs-lookup"><span data-stu-id="3aa16-p106">Currently the only supported event type is `Office.EventType.ItemChanged`, which is invoked when the user selects a new item. This event is used by add-ins that implement a pinnable taskpane, and allows the add-in to refresh the taskpane UI based on the currently selected item.</span></span>

##### <a name="parameters"></a><span data-ttu-id="3aa16-194">Parameter:</span><span class="sxs-lookup"><span data-stu-id="3aa16-194">Parameters:</span></span>

| <span data-ttu-id="3aa16-195">Name</span><span class="sxs-lookup"><span data-stu-id="3aa16-195">Name</span></span> | <span data-ttu-id="3aa16-196">Typ</span><span class="sxs-lookup"><span data-stu-id="3aa16-196">Type</span></span> | <span data-ttu-id="3aa16-197">Attribute</span><span class="sxs-lookup"><span data-stu-id="3aa16-197">Attributes</span></span> | <span data-ttu-id="3aa16-198">Beschreibung</span><span class="sxs-lookup"><span data-stu-id="3aa16-198">Description</span></span> |
|---|---|---|---|
| `eventType` | [<span data-ttu-id="3aa16-199">Office.EventType</span><span class="sxs-lookup"><span data-stu-id="3aa16-199">Office.EventType</span></span>](office.md#eventtype-string) || <span data-ttu-id="3aa16-200">Das Ereignis, das den Handler aufrufen soll</span><span class="sxs-lookup"><span data-stu-id="3aa16-200">The event that should invoke the handler.</span></span> |
| `handler` | <span data-ttu-id="3aa16-201">Function</span><span class="sxs-lookup"><span data-stu-id="3aa16-201">Function</span></span> || <span data-ttu-id="3aa16-p107">Die Funktion, die das Ereignis behandeln soll. Die Funktion muss einen einzigen Parameter akzeptieren (ein Objektliteral). Die Eigenschaft `type` dieses Parameters entspricht dem Parameter `eventType`, der an `addHandlerAsync` übergeben wird.</span><span class="sxs-lookup"><span data-stu-id="3aa16-p107">The function to handle the event. The function must accept a single parameter, which is an object literal. The `type` property on the parameter will match the `eventType` parameter passed to `addHandlerAsync`.</span></span> |
| `options` | <span data-ttu-id="3aa16-205">Objekt</span><span class="sxs-lookup"><span data-stu-id="3aa16-205">Object</span></span> | <span data-ttu-id="3aa16-206">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="3aa16-206">&lt;optional&gt;</span></span> | <span data-ttu-id="3aa16-207">Ein Objektliteral, das mindestens eine der folgenden Eigenschaften enthält.</span><span class="sxs-lookup"><span data-stu-id="3aa16-207">An object literal that contains one or more of the following properties.</span></span> |
| `options.asyncContext` | <span data-ttu-id="3aa16-208">Object</span><span class="sxs-lookup"><span data-stu-id="3aa16-208">Object</span></span> | <span data-ttu-id="3aa16-209">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="3aa16-209">&lt;optional&gt;</span></span> | <span data-ttu-id="3aa16-210">Entwickler können ein Objekt bereitstellen, auf das sie in der Callbackmethode zugreifen möchten.</span><span class="sxs-lookup"><span data-stu-id="3aa16-210">Developers can provide any object they wish to access in the callback method.</span></span> |
| `callback` | <span data-ttu-id="3aa16-211">Funktion</span><span class="sxs-lookup"><span data-stu-id="3aa16-211">function</span></span>| <span data-ttu-id="3aa16-212">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="3aa16-212">&lt;optional&gt;</span></span>|<span data-ttu-id="3aa16-213">Wenn die Methode abgeschlossen wird, wird die Funktion , die im `callback`-Parameter übergeben wird, mit einem einzelnen Parameter (`asyncResult`) aufgerufen, der ein [`AsyncResult`](/javascript/api/office/office.asyncresult)-Objekt darstellt.</span><span class="sxs-lookup"><span data-stu-id="3aa16-213">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="3aa16-214">Anforderungen</span><span class="sxs-lookup"><span data-stu-id="3aa16-214">Requirements</span></span>

|<span data-ttu-id="3aa16-215">Anforderung</span><span class="sxs-lookup"><span data-stu-id="3aa16-215">Requirement</span></span>| <span data-ttu-id="3aa16-216">Wert</span><span class="sxs-lookup"><span data-stu-id="3aa16-216">Value</span></span>|
|---|---|
|[<span data-ttu-id="3aa16-217">Mindestversion des Postfachanforderungssatzes</span><span class="sxs-lookup"><span data-stu-id="3aa16-217">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="3aa16-218">1,5</span><span class="sxs-lookup"><span data-stu-id="3aa16-218">1.5</span></span> |
|[<span data-ttu-id="3aa16-219">Mindestberechtigungsstufe</span><span class="sxs-lookup"><span data-stu-id="3aa16-219">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="3aa16-220">ReadItem</span><span class="sxs-lookup"><span data-stu-id="3aa16-220">ReadItem</span></span> |
|[<span data-ttu-id="3aa16-221">Zutreffender Outlook-Modus</span><span class="sxs-lookup"><span data-stu-id="3aa16-221">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="3aa16-222">Verfassen oder Lesen</span><span class="sxs-lookup"><span data-stu-id="3aa16-222">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="3aa16-223">Beispiel</span><span class="sxs-lookup"><span data-stu-id="3aa16-223">Example</span></span>

```
Office.initialize = function (reason) {
  $(document).ready(function () {
    Office.context.mailbox.addHandlerAsync(Office.EventType.ItemChanged, loadNewItem, function (result) {
      if (result.status === Office.AsyncResultStatus.Failed) {
        // Handle error
      }
    });
  });
};

function loadNewItem(eventArgs) {
  // Load the properties of the newly selected item
  loadProps(Office.context.mailbox.item);
};
```

####  <a name="converttoewsiditemid-restversion--string"></a><span data-ttu-id="3aa16-224">convertToEwsId(itemId, restVersion) → {String}</span><span class="sxs-lookup"><span data-stu-id="3aa16-224">convertToEwsId(itemId, restVersion) → {String}</span></span>

<span data-ttu-id="3aa16-225">Wandelt eine für REST formatierte Element-ID ins EWS-Format um.</span><span class="sxs-lookup"><span data-stu-id="3aa16-225">Converts an item ID formatted for REST into EWS format.</span></span>

> [!NOTE]
> <span data-ttu-id="3aa16-226">Diese Methode wird in Outlook für iOS oder Outlook für Android nicht unterstützt.</span><span class="sxs-lookup"><span data-stu-id="3aa16-226">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="3aa16-p108">Über eine REST-API abgerufene Element-IDs (z. B. die [Outlook Mail-API](https://docs.microsoft.com/previous-versions/office/office-365-api/api/version-2.0/mail-rest-operations) oder [Microsoft Graph](http://graph.microsoft.io/)) verwenden ein anderes Format, als das von Exchange-Webdiensten (EWS) verwendete. Die `convertToEwsId`-Methode wandelt eine für REST formatierte ID in das entsprechende EWS-Format um.</span><span class="sxs-lookup"><span data-stu-id="3aa16-p108">Item IDs retrieved via a REST API (such as the [Outlook Mail API](https://docs.microsoft.com/previous-versions/office/office-365-api/api/version-2.0/mail-rest-operations) or the [Microsoft Graph](http://graph.microsoft.io/)) use a different format than the format used by Exchange Web Services (EWS). The `convertToEwsId` method converts a REST-formatted ID into the proper format for EWS.</span></span>

##### <a name="parameters"></a><span data-ttu-id="3aa16-229">Parameter:</span><span class="sxs-lookup"><span data-stu-id="3aa16-229">Parameters:</span></span>

|<span data-ttu-id="3aa16-230">Name</span><span class="sxs-lookup"><span data-stu-id="3aa16-230">Name</span></span>| <span data-ttu-id="3aa16-231">Typ</span><span class="sxs-lookup"><span data-stu-id="3aa16-231">Type</span></span>| <span data-ttu-id="3aa16-232">Beschreibung</span><span class="sxs-lookup"><span data-stu-id="3aa16-232">Description</span></span>|
|---|---|---|
|`itemId`| <span data-ttu-id="3aa16-233">Zeichenfolge</span><span class="sxs-lookup"><span data-stu-id="3aa16-233">String</span></span>|<span data-ttu-id="3aa16-234">Eine für Outlook-REST-APIs formatierte Element-ID</span><span class="sxs-lookup"><span data-stu-id="3aa16-234">An item ID formatted for the Outlook REST APIs</span></span>|
|`restVersion`| [<span data-ttu-id="3aa16-235">Office.MailboxEnums.RestVersion</span><span class="sxs-lookup"><span data-stu-id="3aa16-235">Office.MailboxEnums.RestVersion</span></span>](/javascript/api/outlook_1_6/office.mailboxenums.restversion)|<span data-ttu-id="3aa16-236">Ein Wert, der die Version der Outlook-REST-API angibt, die zum Abrufen der Element-ID verwendet wurde.</span><span class="sxs-lookup"><span data-stu-id="3aa16-236">A value indicating the version of the Outlook REST API used to retrieve the item ID.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="3aa16-237">Anforderungen</span><span class="sxs-lookup"><span data-stu-id="3aa16-237">Requirements</span></span>

|<span data-ttu-id="3aa16-238">Anforderung</span><span class="sxs-lookup"><span data-stu-id="3aa16-238">Requirement</span></span>| <span data-ttu-id="3aa16-239">Wert</span><span class="sxs-lookup"><span data-stu-id="3aa16-239">Value</span></span>|
|---|---|
|[<span data-ttu-id="3aa16-240">Mindestversion des Postfachanforderungssatzes</span><span class="sxs-lookup"><span data-stu-id="3aa16-240">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="3aa16-241">1.3</span><span class="sxs-lookup"><span data-stu-id="3aa16-241">1.3</span></span>|
|[<span data-ttu-id="3aa16-242">Mindestberechtigungsstufe</span><span class="sxs-lookup"><span data-stu-id="3aa16-242">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="3aa16-243">Eingeschränkt</span><span class="sxs-lookup"><span data-stu-id="3aa16-243">Restricted</span></span>|
|[<span data-ttu-id="3aa16-244">Zutreffender Outlook-Modus</span><span class="sxs-lookup"><span data-stu-id="3aa16-244">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="3aa16-245">Verfassen oder Lesen</span><span class="sxs-lookup"><span data-stu-id="3aa16-245">Compose or read</span></span>|

##### <a name="returns"></a><span data-ttu-id="3aa16-246">Gibt zurück:</span><span class="sxs-lookup"><span data-stu-id="3aa16-246">Returns:</span></span>

<span data-ttu-id="3aa16-247">Typ: Zeichenfolge</span><span class="sxs-lookup"><span data-stu-id="3aa16-247">Type: String</span></span>

##### <a name="example"></a><span data-ttu-id="3aa16-248">Beispiel</span><span class="sxs-lookup"><span data-stu-id="3aa16-248">Example</span></span>

```
// Get an item's ID from a REST API
var restId = 'AAMkAGVlOTZjNTM3LW...';

// Treat restId as coming from the v2.0 version of the
// Outlook Mail API
var ewsId = Office.context.mailbox.convertToEwsId(restId, Office.MailboxEnums.RestVersion.v2_0);
```

####  <a name="converttolocalclienttimetimevalue--localclienttimejavascriptapioutlook16officelocalclienttime"></a><span data-ttu-id="3aa16-249">convertToLocalClientTime(timeValue) → {[LocalClientTime](/javascript/api/outlook_1_6/office.LocalClientTime)}</span><span class="sxs-lookup"><span data-stu-id="3aa16-249">convertToLocalClientTime(timeValue) → {[LocalClientTime](/javascript/api/outlook_1_6/office.LocalClientTime)}</span></span>

<span data-ttu-id="3aa16-250">Ruft ein Wörterbuch mit Uhrzeitinformationen basierend auf der Zeiteinstellung des lokalen Clients ab.</span><span class="sxs-lookup"><span data-stu-id="3aa16-250">Gets a dictionary containing time information in local client time.</span></span>

<span data-ttu-id="3aa16-p109">Die in einer Mail-App für Outlook oder Outlook Web App verwendeten Daten und Uhrzeiten stammen u. U. aus verschiedenen Zeitzonen. Outlook verwendet die Zeitzone des Client-Computers; Outlook Web App verwendet die im Exchange Admin Center (EAC) festgelegte Zeitzone. Sie sollten Datums- und Uhrzeitwerte bearbeiten, damit die auf der Benutzeroberfläche angezeigten Werte immer den von Benutzer erwarteten Zeitzonen entsprechen.</span><span class="sxs-lookup"><span data-stu-id="3aa16-p109">The dates and times used by a mail app for Outlook or Outlook Web App can use different time zones. Outlook uses the client computer time zone; Outlook Web App uses the time zone set on the Exchange Admin Center (EAC). You should handle date and time values so that the values you display on the user interface are always consistent with the time zone that the user expects.</span></span>

<span data-ttu-id="3aa16-p110">Wird die Mail-App in Outlook ausgeführt, wird mit der `convertToLocalClientTime`-Methode ein Wörterbuchobjekt zurückgegeben, dessen Werte auf die Zeitzone des Clientcomputers festgelegt sind. Wird die Mail-App in Outlook ausgeführt, wird mit der `convertToLocalClientTime`-Methode ein Wörterbuchobjekt zurückgegeben, dessen Werte auf die Zeitzone des Clientcomputers festgelegt sind.</span><span class="sxs-lookup"><span data-stu-id="3aa16-p110">If the mail app is running in Outlook, the `convertToLocalClientTime` method will return a dictionary object with the values set to the client computer time zone. If the mail app is running in Outlook Web App, the `convertToLocalClientTime` method will return a dictionary object with the values set to the time zone specified in the EAC.</span></span>

##### <a name="parameters"></a><span data-ttu-id="3aa16-256">Parameter:</span><span class="sxs-lookup"><span data-stu-id="3aa16-256">Parameters:</span></span>

|<span data-ttu-id="3aa16-257">Name</span><span class="sxs-lookup"><span data-stu-id="3aa16-257">Name</span></span>| <span data-ttu-id="3aa16-258">Typ</span><span class="sxs-lookup"><span data-stu-id="3aa16-258">Type</span></span>| <span data-ttu-id="3aa16-259">Beschreibung</span><span class="sxs-lookup"><span data-stu-id="3aa16-259">Description</span></span>|
|---|---|---|
|`timeValue`| <span data-ttu-id="3aa16-260">Datum</span><span class="sxs-lookup"><span data-stu-id="3aa16-260">Date</span></span>|<span data-ttu-id="3aa16-261">Ein Date-Objekt</span><span class="sxs-lookup"><span data-stu-id="3aa16-261">A Date object</span></span>|

##### <a name="requirements"></a><span data-ttu-id="3aa16-262">Anforderungen</span><span class="sxs-lookup"><span data-stu-id="3aa16-262">Requirements</span></span>

|<span data-ttu-id="3aa16-263">Anforderung</span><span class="sxs-lookup"><span data-stu-id="3aa16-263">Requirement</span></span>| <span data-ttu-id="3aa16-264">Wert</span><span class="sxs-lookup"><span data-stu-id="3aa16-264">Value</span></span>|
|---|---|
|[<span data-ttu-id="3aa16-265">Mindestversion des Postfachanforderungssatzes</span><span class="sxs-lookup"><span data-stu-id="3aa16-265">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="3aa16-266">1.0</span><span class="sxs-lookup"><span data-stu-id="3aa16-266">1.0</span></span>|
|[<span data-ttu-id="3aa16-267">Mindestberechtigungsstufe</span><span class="sxs-lookup"><span data-stu-id="3aa16-267">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="3aa16-268">ReadItem</span><span class="sxs-lookup"><span data-stu-id="3aa16-268">ReadItem</span></span>|
|[<span data-ttu-id="3aa16-269">Zutreffender Outlook-Modus</span><span class="sxs-lookup"><span data-stu-id="3aa16-269">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="3aa16-270">Verfassen oder Lesen</span><span class="sxs-lookup"><span data-stu-id="3aa16-270">Compose or read</span></span>|

##### <a name="returns"></a><span data-ttu-id="3aa16-271">Gibt zurück:</span><span class="sxs-lookup"><span data-stu-id="3aa16-271">Returns:</span></span>

<span data-ttu-id="3aa16-272">Typ: [LocalClientTime](/javascript/api/outlook_1_6/office.LocalClientTime)</span><span class="sxs-lookup"><span data-stu-id="3aa16-272">Type: [LocalClientTime](/javascript/api/outlook_1_6/office.LocalClientTime)</span></span>

####  <a name="converttorestiditemid-restversion--string"></a><span data-ttu-id="3aa16-273">convertToRestId(itemId, restVersion) → {String}</span><span class="sxs-lookup"><span data-stu-id="3aa16-273">convertToRestId(itemId, restVersion) → {String}</span></span>

<span data-ttu-id="3aa16-274">Wandelt eine für EWS formatierte Element-ID ins REST-Format um.</span><span class="sxs-lookup"><span data-stu-id="3aa16-274">Converts an item ID formatted for EWS into REST format.</span></span>

> [!NOTE]
> <span data-ttu-id="3aa16-275">Diese Methode wird in Outlook für iOS oder Outlook für Android nicht unterstützt.</span><span class="sxs-lookup"><span data-stu-id="3aa16-275">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="3aa16-p111">Über EWS oder die `itemId` abgerufene Element-IDs verwenden ein anderes Format als die über REST-APIs (z. B. die [Outlook Mail-API](https://docs.microsoft.com/previous-versions/office/office-365-api/api/version-2.0/mail-rest-operations) oder [Microsoft Graph](http://graph.microsoft.io/)) abgerufenen Element-IDs. Die `convertToRestId`-Methode wandelt eine für EWS formatierte ID in das entsprechende REST-Format um.</span><span class="sxs-lookup"><span data-stu-id="3aa16-p111">Item IDs retrieved via EWS or via the `itemId` property use a different format than the format used by REST APIs (such as the [Outlook Mail API](https://docs.microsoft.com/previous-versions/office/office-365-api/api/version-2.0/mail-rest-operations) or the [Microsoft Graph](http://graph.microsoft.io/)). The `convertToRestId` method converts an EWS-formatted ID into the proper format for REST.</span></span>

##### <a name="parameters"></a><span data-ttu-id="3aa16-278">Parameter:</span><span class="sxs-lookup"><span data-stu-id="3aa16-278">Parameters:</span></span>

|<span data-ttu-id="3aa16-279">Name</span><span class="sxs-lookup"><span data-stu-id="3aa16-279">Name</span></span>| <span data-ttu-id="3aa16-280">Typ</span><span class="sxs-lookup"><span data-stu-id="3aa16-280">Type</span></span>| <span data-ttu-id="3aa16-281">Beschreibung</span><span class="sxs-lookup"><span data-stu-id="3aa16-281">Description</span></span>|
|---|---|---|
|`itemId`| <span data-ttu-id="3aa16-282">Zeichenfolge</span><span class="sxs-lookup"><span data-stu-id="3aa16-282">String</span></span>|<span data-ttu-id="3aa16-283">Eine für Exchange-Webdienste (EWS) formatierte Element-ID</span><span class="sxs-lookup"><span data-stu-id="3aa16-283">An item ID formatted for Exchange Web Services (EWS)</span></span>|
|`restVersion`| [<span data-ttu-id="3aa16-284">Office.MailboxEnums.RestVersion</span><span class="sxs-lookup"><span data-stu-id="3aa16-284">Office.MailboxEnums.RestVersion</span></span>](/javascript/api/outlook_1_6/office.mailboxenums.restversion)|<span data-ttu-id="3aa16-285">Ein Wert, der die Version der Outlook-REST-API angibt, mit der die konvertierte ID verwendet wird.</span><span class="sxs-lookup"><span data-stu-id="3aa16-285">A value indicating the version of the Outlook REST API that the converted ID will be used with.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="3aa16-286">Anforderungen</span><span class="sxs-lookup"><span data-stu-id="3aa16-286">Requirements</span></span>

|<span data-ttu-id="3aa16-287">Anforderung</span><span class="sxs-lookup"><span data-stu-id="3aa16-287">Requirement</span></span>| <span data-ttu-id="3aa16-288">Wert</span><span class="sxs-lookup"><span data-stu-id="3aa16-288">Value</span></span>|
|---|---|
|[<span data-ttu-id="3aa16-289">Mindestversion des Postfachanforderungssatzes</span><span class="sxs-lookup"><span data-stu-id="3aa16-289">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="3aa16-290">1.3</span><span class="sxs-lookup"><span data-stu-id="3aa16-290">1.3</span></span>|
|[<span data-ttu-id="3aa16-291">Mindestberechtigungsstufe</span><span class="sxs-lookup"><span data-stu-id="3aa16-291">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="3aa16-292">Eingeschränkt</span><span class="sxs-lookup"><span data-stu-id="3aa16-292">Restricted</span></span>|
|[<span data-ttu-id="3aa16-293">Zutreffender Outlook-Modus</span><span class="sxs-lookup"><span data-stu-id="3aa16-293">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="3aa16-294">Verfassen oder Lesen</span><span class="sxs-lookup"><span data-stu-id="3aa16-294">Compose or read</span></span>|

##### <a name="returns"></a><span data-ttu-id="3aa16-295">Gibt zurück:</span><span class="sxs-lookup"><span data-stu-id="3aa16-295">Returns:</span></span>

<span data-ttu-id="3aa16-296">Typ: Zeichenfolge</span><span class="sxs-lookup"><span data-stu-id="3aa16-296">Type: String</span></span>

##### <a name="example"></a><span data-ttu-id="3aa16-297">Beispiel</span><span class="sxs-lookup"><span data-stu-id="3aa16-297">Example</span></span>

```
// Get the currently selected item's ID
var ewsId = Office.context.mailbox.item.itemId;

// Convert to a REST ID for the v2.0 version of the
// Outlook Mail API
var restId = Office.context.mailbox.convertToRestId(ewsId, Office.MailboxEnums.RestVersion.v2_0);
```

####  <a name="converttoutcclienttimeinput--date"></a><span data-ttu-id="3aa16-298">convertToUtcClientTime(input) → {Date}</span><span class="sxs-lookup"><span data-stu-id="3aa16-298">convertToUtcClientTime(input) → {Date}</span></span>

<span data-ttu-id="3aa16-299">Ruft ein Date-Objekt aus einem Wörterbuch mit Uhrzeitinformationen ab.</span><span class="sxs-lookup"><span data-stu-id="3aa16-299">Gets a Date object from a dictionary containing time information.</span></span>

<span data-ttu-id="3aa16-300">Mit der `convertToUtcClientTime`-Methode wird ein Wörterbuch mit lokalem Datum und lokaler Uhrzeit in ein Date-Objekt mit den richtigen Werten für das lokale Datum und die lokale Uhrzeit konvertiert.</span><span class="sxs-lookup"><span data-stu-id="3aa16-300">The `convertToUtcClientTime` method converts a dictionary containing a local date and time to a Date object with the correct values for the local date and time.</span></span>

##### <a name="parameters"></a><span data-ttu-id="3aa16-301">Parameter:</span><span class="sxs-lookup"><span data-stu-id="3aa16-301">Parameters:</span></span>

|<span data-ttu-id="3aa16-302">Name</span><span class="sxs-lookup"><span data-stu-id="3aa16-302">Name</span></span>| <span data-ttu-id="3aa16-303">Typ</span><span class="sxs-lookup"><span data-stu-id="3aa16-303">Type</span></span>| <span data-ttu-id="3aa16-304">Beschreibung</span><span class="sxs-lookup"><span data-stu-id="3aa16-304">Description</span></span>|
|---|---|---|
|`input`| [<span data-ttu-id="3aa16-305">LocalClientTime</span><span class="sxs-lookup"><span data-stu-id="3aa16-305">LocalClientTime</span></span>](/javascript/api/outlook_1_6/office.LocalClientTime)|<span data-ttu-id="3aa16-306">Der zu konvertierende Wert für die lokale Uhrzeit.</span><span class="sxs-lookup"><span data-stu-id="3aa16-306">The local time value to convert.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="3aa16-307">Anforderungen</span><span class="sxs-lookup"><span data-stu-id="3aa16-307">Requirements</span></span>

|<span data-ttu-id="3aa16-308">Anforderung</span><span class="sxs-lookup"><span data-stu-id="3aa16-308">Requirement</span></span>| <span data-ttu-id="3aa16-309">Wert</span><span class="sxs-lookup"><span data-stu-id="3aa16-309">Value</span></span>|
|---|---|
|[<span data-ttu-id="3aa16-310">Mindestversion des Postfachanforderungssatzes</span><span class="sxs-lookup"><span data-stu-id="3aa16-310">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="3aa16-311">1.0</span><span class="sxs-lookup"><span data-stu-id="3aa16-311">1.0</span></span>|
|[<span data-ttu-id="3aa16-312">Mindestberechtigungsstufe</span><span class="sxs-lookup"><span data-stu-id="3aa16-312">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="3aa16-313">ReadItem</span><span class="sxs-lookup"><span data-stu-id="3aa16-313">ReadItem</span></span>|
|[<span data-ttu-id="3aa16-314">Zutreffender Outlook-Modus</span><span class="sxs-lookup"><span data-stu-id="3aa16-314">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="3aa16-315">Verfassen oder Lesen</span><span class="sxs-lookup"><span data-stu-id="3aa16-315">Compose or read</span></span>|

##### <a name="returns"></a><span data-ttu-id="3aa16-316">Gibt zurück:</span><span class="sxs-lookup"><span data-stu-id="3aa16-316">Returns:</span></span>

<span data-ttu-id="3aa16-317">Ein Date-Objekt der Uhrzeit in UTC.</span><span class="sxs-lookup"><span data-stu-id="3aa16-317">A Date object with the time expressed in UTC.</span></span>

<dl class="param-type"><span data-ttu-id="3aa16-318">

<dt>
Typ</dt>


</span><span class="sxs-lookup"><span data-stu-id="3aa16-318">

<dt>Type</dt>

</span></span><dd><span data-ttu-id="3aa16-319">Datum</span><span class="sxs-lookup"><span data-stu-id="3aa16-319">Date</span></span></dd>

</dl>

####  <a name="displayappointmentformitemid"></a><span data-ttu-id="3aa16-320">displayAppointmentForm(itemId)</span><span class="sxs-lookup"><span data-stu-id="3aa16-320">displayAppointmentForm(itemId)</span></span>

<span data-ttu-id="3aa16-321">Zeigt einen bestehenden Kalendertermin an.</span><span class="sxs-lookup"><span data-stu-id="3aa16-321">Displays an existing calendar appointment.</span></span>

> [!NOTE]
> <span data-ttu-id="3aa16-322">Diese Methode wird in Outlook für iOS oder Outlook für Android nicht unterstützt.</span><span class="sxs-lookup"><span data-stu-id="3aa16-322">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="3aa16-323">Mit der `displayAppointmentForm`-Methode wird ein vorhandener Kalendertermin auf dem Desktop in einem neuen Fenster oder auf Mobilgeräten in einem Dialogfeld geöffnet.</span><span class="sxs-lookup"><span data-stu-id="3aa16-323">The `displayAppointmentForm` method opens an existing calendar appointment in a new window on the desktop or in a dialog box on mobile devices.</span></span>

<span data-ttu-id="3aa16-p112">In Outlook for Mac können Sie diese Methode verwenden, um einen einzelnen Termin anzuzeigen, der nicht Teil einer Terminserie ist, oder den Mastertermin einer Terminserie, jedoch keine Instanz der Terminserie. Dies liegt daran, dass Sie in Outlook for Mac nicht auf die Eigenschaften (einschließlich der Element-ID) von Instanzen einer Terminserie zugreifen können.</span><span class="sxs-lookup"><span data-stu-id="3aa16-p112">In Outlook for Mac, you can use this method to display a single appointment that is not part of a recurring series, or the master appointment of a recurring series, but you cannot display an instance of the series. This is because in Outlook for Mac, you cannot access the properties (including the item ID) of instances of a recurring series.</span></span>

<span data-ttu-id="3aa16-326">In Outlook Web App öffnet diese Methode das angegebene Formular nur, wenn der Textkörper des Formulars nicht größer ist als 32 KB.</span><span class="sxs-lookup"><span data-stu-id="3aa16-326">In Outlook Web App, this method opens the specified form only if the body of the form is less than or equal to 32KB number of characters.</span></span>

<span data-ttu-id="3aa16-327">Wenn der angegebene Elementbezeichner keinen vorhandenen Termin identifiziert, wird auf dem Clientcomputer oder Gerät ein leerer Bereich geöffnet, und es wird keine Fehlermeldung zurückgegeben.</span><span class="sxs-lookup"><span data-stu-id="3aa16-327">If the specified item identifier does not identify an existing appointment, a blank pane opens on the client computer or device, and no error message will be returned.</span></span>

##### <a name="parameters"></a><span data-ttu-id="3aa16-328">Parameter:</span><span class="sxs-lookup"><span data-stu-id="3aa16-328">Parameters:</span></span>

|<span data-ttu-id="3aa16-329">Name</span><span class="sxs-lookup"><span data-stu-id="3aa16-329">Name</span></span>| <span data-ttu-id="3aa16-330">Typ</span><span class="sxs-lookup"><span data-stu-id="3aa16-330">Type</span></span>| <span data-ttu-id="3aa16-331">Beschreibung</span><span class="sxs-lookup"><span data-stu-id="3aa16-331">Description</span></span>|
|---|---|---|
|`itemId`| <span data-ttu-id="3aa16-332">Zeichenfolge</span><span class="sxs-lookup"><span data-stu-id="3aa16-332">String</span></span>|<span data-ttu-id="3aa16-333">Der EWS-Bezeichner (Exchange Web Services, Exchange-Webdienste) für einen vorhandenen Kalendertermin.</span><span class="sxs-lookup"><span data-stu-id="3aa16-333">The Exchange Web Services (EWS) identifier for an existing calendar appointment.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="3aa16-334">Anforderungen</span><span class="sxs-lookup"><span data-stu-id="3aa16-334">Requirements</span></span>

|<span data-ttu-id="3aa16-335">Anforderung</span><span class="sxs-lookup"><span data-stu-id="3aa16-335">Requirement</span></span>| <span data-ttu-id="3aa16-336">Wert</span><span class="sxs-lookup"><span data-stu-id="3aa16-336">Value</span></span>|
|---|---|
|[<span data-ttu-id="3aa16-337">Mindestversion des Postfachanforderungssatzes</span><span class="sxs-lookup"><span data-stu-id="3aa16-337">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="3aa16-338">1.0</span><span class="sxs-lookup"><span data-stu-id="3aa16-338">1.0</span></span>|
|[<span data-ttu-id="3aa16-339">Mindestberechtigungsstufe</span><span class="sxs-lookup"><span data-stu-id="3aa16-339">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="3aa16-340">ReadItem</span><span class="sxs-lookup"><span data-stu-id="3aa16-340">ReadItem</span></span>|
|[<span data-ttu-id="3aa16-341">Zutreffender Outlook-Modus</span><span class="sxs-lookup"><span data-stu-id="3aa16-341">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="3aa16-342">Verfassen oder Lesen</span><span class="sxs-lookup"><span data-stu-id="3aa16-342">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="3aa16-343">Beispiel</span><span class="sxs-lookup"><span data-stu-id="3aa16-343">Example</span></span>

```
Office.context.mailbox.displayAppointmentForm(appointmentId);
```

####  <a name="displaymessageformitemid"></a><span data-ttu-id="3aa16-344">displayMessageForm(itemId)</span><span class="sxs-lookup"><span data-stu-id="3aa16-344">displayMessageForm(itemId)</span></span>

<span data-ttu-id="3aa16-345">Zeigt eine vorhandene Nachricht an.</span><span class="sxs-lookup"><span data-stu-id="3aa16-345">Displays an existing message.</span></span>

> [!NOTE]
> <span data-ttu-id="3aa16-346">Diese Methode wird in Outlook für iOS oder Outlook für Android nicht unterstützt.</span><span class="sxs-lookup"><span data-stu-id="3aa16-346">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="3aa16-347">Die `displayMessageForm`-Methode öffnet eine vorhandene Nachricht in einem neuen Fenster auf dem Desktop bzw. in einem Dialogfeld auf Mobilgeräten.</span><span class="sxs-lookup"><span data-stu-id="3aa16-347">The `displayMessageForm` method opens an existing message in a new window on the desktop or in a dialog box on mobile devices.</span></span>

<span data-ttu-id="3aa16-348">In Outlook Web App wird mit dieser Methode das angegebene Formular nur dann geöffnet, wenn der Textkörper des Formular Zeichen im Umfang vom maximal 32 KB umfasst.</span><span class="sxs-lookup"><span data-stu-id="3aa16-348">In Outlook Web App, this method opens the specified form only if the body of the form is less than or equal to 32 KB number of characters.</span></span>

<span data-ttu-id="3aa16-349">Wenn der angegebene Elementbezeichner keine vorhandenen Nachrichten erkennt, wird auf dem Client-Computer keine Nachricht angezeigt, und es werden keine Fehlermeldungen zurückgegeben.</span><span class="sxs-lookup"><span data-stu-id="3aa16-349">If the specified item identifier does not identify an existing message, no message will be displayed on the client computer, and no error message will be returned.</span></span>

<span data-ttu-id="3aa16-p113">Verwenden Sie `displayMessageForm` nicht mit einem `itemId`-Objekt, das einen Termin darstellt. Verwenden Sie die `displayAppointmentForm`-Methode, um einen vorhandenen Termin anzuzeigen, und `displayNewAppointmentForm`, um ein Formular zum Erstellen eines neuen Termins anzuzeigen.</span><span class="sxs-lookup"><span data-stu-id="3aa16-p113">Do not use the `displayMessageForm` with an `itemId` that represents an appointment. Use the `displayAppointmentForm` method to display an existing appointment, and `displayNewAppointmentForm` to display a form to create a new appointment.</span></span>

##### <a name="parameters"></a><span data-ttu-id="3aa16-352">Parameter:</span><span class="sxs-lookup"><span data-stu-id="3aa16-352">Parameters:</span></span>

|<span data-ttu-id="3aa16-353">Name</span><span class="sxs-lookup"><span data-stu-id="3aa16-353">Name</span></span>| <span data-ttu-id="3aa16-354">Typ</span><span class="sxs-lookup"><span data-stu-id="3aa16-354">Type</span></span>| <span data-ttu-id="3aa16-355">Beschreibung</span><span class="sxs-lookup"><span data-stu-id="3aa16-355">Description</span></span>|
|---|---|---|
|`itemId`| <span data-ttu-id="3aa16-356">Zeichenfolge</span><span class="sxs-lookup"><span data-stu-id="3aa16-356">String</span></span>|<span data-ttu-id="3aa16-357">Der Exchange-Webdienste (EWS) für eine vorhandene Nachricht.</span><span class="sxs-lookup"><span data-stu-id="3aa16-357">The Exchange Web Services (EWS) identifier for an existing message.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="3aa16-358">Anforderungen</span><span class="sxs-lookup"><span data-stu-id="3aa16-358">Requirements</span></span>

|<span data-ttu-id="3aa16-359">Anforderung</span><span class="sxs-lookup"><span data-stu-id="3aa16-359">Requirement</span></span>| <span data-ttu-id="3aa16-360">Wert</span><span class="sxs-lookup"><span data-stu-id="3aa16-360">Value</span></span>|
|---|---|
|[<span data-ttu-id="3aa16-361">Mindestversion des Postfachanforderungssatzes</span><span class="sxs-lookup"><span data-stu-id="3aa16-361">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="3aa16-362">1.0</span><span class="sxs-lookup"><span data-stu-id="3aa16-362">1.0</span></span>|
|[<span data-ttu-id="3aa16-363">Mindestberechtigungsstufe</span><span class="sxs-lookup"><span data-stu-id="3aa16-363">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="3aa16-364">ReadItem</span><span class="sxs-lookup"><span data-stu-id="3aa16-364">ReadItem</span></span>|
|[<span data-ttu-id="3aa16-365">Zutreffender Outlook-Modus</span><span class="sxs-lookup"><span data-stu-id="3aa16-365">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="3aa16-366">Verfassen oder Lesen</span><span class="sxs-lookup"><span data-stu-id="3aa16-366">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="3aa16-367">Beispiel</span><span class="sxs-lookup"><span data-stu-id="3aa16-367">Example</span></span>

```
Office.context.mailbox.displayMessageForm(messageId);
```

#### <a name="displaynewappointmentformparameters"></a><span data-ttu-id="3aa16-368">displayNewAppointmentForm(parameters)</span><span class="sxs-lookup"><span data-stu-id="3aa16-368">displayNewAppointmentForm(parameters)</span></span>

<span data-ttu-id="3aa16-369">Zeigt ein Formular zum Erstellen eines neuen Kalendertermins an.</span><span class="sxs-lookup"><span data-stu-id="3aa16-369">Displays a form for creating a new calendar appointment.</span></span>

> [!NOTE]
> <span data-ttu-id="3aa16-370">Diese Methode wird in Outlook für iOS oder Outlook für Android nicht unterstützt.</span><span class="sxs-lookup"><span data-stu-id="3aa16-370">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="3aa16-p114">Mit der `displayNewAppointmentForm`-Methode wird ein Formular geöffnet, mit dem der Benutzer einen neuen Termin oder eine Besprechung erstellen kann. Wenn Parameter angegeben wurden, werden die Felder im Terminformular automatisch mit dem Inhalt der Parameter ausgefüllt.</span><span class="sxs-lookup"><span data-stu-id="3aa16-p114">The `displayNewAppointmentForm` method opens a form that enables the user to create a new appointment or meeting. If parameters are specified, the appointment form fields are automatically populated with the contents of the parameters.</span></span>

<span data-ttu-id="3aa16-p115">In Outlook Web App und OWA for Devices zeigt diese Methode immer ein Formular mit einem Teilnehmerfeld an. Wenn Sie keine Teilnehmer als Eingabeargumente angeben, zeigt die Methode ein Formular mit einer Schaltfläche **Speichern** an. Wenn Sie Teilnehmer angegeben haben, enthält das Formular die Teilnehmer und eine Schaltfläche **Senden**.</span><span class="sxs-lookup"><span data-stu-id="3aa16-p115">In Outlook Web App and OWA for Devices, this method always displays a form with an attendees field. If you do not specify any attendees as input arguments, the method displays a form with a **Save** button. If you have specified attendees, the form would include the attendees and a **Send** button.</span></span>

<span data-ttu-id="3aa16-p116">Wenn Teilnehmer oder Ressourcen im `requiredAttendees`-, `optionalAttendees`- oder `resources`-Parameter festgelegt werden, wird mit dieser Methode im Outlook Rich Client und Outlook RT ein Besprechungsformular mit der Schaltfläche **Senden** angezeigt. Wenn keine Teilnehmer angegeben werden, wird mit dieser Methode ein Terminformular mit der Schaltfläche **Speichern & schließen** angezeigt.</span><span class="sxs-lookup"><span data-stu-id="3aa16-p116">In the Outlook rich client and Outlook RT, if you specify any attendees or resources in the `requiredAttendees`, `optionalAttendees`, or `resources` parameter, this method displays a meeting form with a **Send** button. If you don't specify any recipients, this method displays an appointment form with a **Save & Close** button.</span></span>

<span data-ttu-id="3aa16-378">Wenn einer der Parameter die angegebenen Größenbeschränkungen überschreitet oder wenn ein unbekannter Parametername angegeben wird, wird eine Ausnahme ausgelöst.</span><span class="sxs-lookup"><span data-stu-id="3aa16-378">If any of the parameters exceed the specified size limits, or if an unknown parameter name is specified, an exception is thrown.</span></span>

##### <a name="parameters"></a><span data-ttu-id="3aa16-379">Parameter:</span><span class="sxs-lookup"><span data-stu-id="3aa16-379">Parameters:</span></span>

> [!NOTE]
> <span data-ttu-id="3aa16-380">Alle Parameter sind optional.</span><span class="sxs-lookup"><span data-stu-id="3aa16-380">All parameters are optional.</span></span>

|<span data-ttu-id="3aa16-381">Name</span><span class="sxs-lookup"><span data-stu-id="3aa16-381">Name</span></span>| <span data-ttu-id="3aa16-382">Typ</span><span class="sxs-lookup"><span data-stu-id="3aa16-382">Type</span></span>| <span data-ttu-id="3aa16-383">Beschreibung</span><span class="sxs-lookup"><span data-stu-id="3aa16-383">Description</span></span>|
|---|---|---|
| `parameters` | <span data-ttu-id="3aa16-384">Object</span><span class="sxs-lookup"><span data-stu-id="3aa16-384">Object</span></span> | <span data-ttu-id="3aa16-385">Ein Wörterbuch mit Parametern, die den neuen Termin beschreiben</span><span class="sxs-lookup"><span data-stu-id="3aa16-385">A dictionary of parameters describing the new appointment.</span></span> |
| `parameters.requiredAttendees` | <span data-ttu-id="3aa16-386">Array.&lt;String&gt; &#124; Array.&lt;[EmailAddressDetails](/javascript/api/outlook_1_6/office.emailaddressdetails)&gt;</span><span class="sxs-lookup"><span data-stu-id="3aa16-386">Array.&lt;String&gt; &#124; Array.&lt;[EmailAddressDetails](/javascript/api/outlook_1_6/office.emailaddressdetails)&gt;</span></span> | <span data-ttu-id="3aa16-p117">Ein Array aus Zeichenfolgen mit den E-Mail-Adressen oder ein Array mit einem `EmailAddressDetails`-Objekt für jeden der erforderlichen Teilnehmer für den Termin. Das Array darf maximal 100 Einträge enthalten.</span><span class="sxs-lookup"><span data-stu-id="3aa16-p117">An array of strings containing the email addresses or an array containing an `EmailAddressDetails` object for each of the required attendees for the appointment. The array is limited to a maximum of 100 entries.</span></span> |
| `parameters.optionalAttendees` | <span data-ttu-id="3aa16-389">Array.&lt;String&gt; &#124; Array.&lt;[EmailAddressDetails](/javascript/api/outlook_1_6/office.emailaddressdetails)&gt;</span><span class="sxs-lookup"><span data-stu-id="3aa16-389">Array.&lt;String&gt; &#124; Array.&lt;[EmailAddressDetails](/javascript/api/outlook_1_6/office.emailaddressdetails)&gt;</span></span> | <span data-ttu-id="3aa16-p118">Ein Array aus Zeichenfolgen mit den E-Mail-Adressen oder ein Array mit einem Objekt des Typs `EmailAddressDetails` für jeden der optionalen Teilnehmer des Termins. Das Array darf maximal 100 Einträge enthalten.</span><span class="sxs-lookup"><span data-stu-id="3aa16-p118">An array of strings containing the email addresses or an array containing an `EmailAddressDetails` object for each of the optional attendees for the appointment. The array is limited to a maximum of 100 entries.</span></span> |
| `parameters.start` | <span data-ttu-id="3aa16-392">Datum</span><span class="sxs-lookup"><span data-stu-id="3aa16-392">Date</span></span> | <span data-ttu-id="3aa16-393">Ein Objekt des Typs `Date`. Gibt das Startdatum und den Beginn des Termins an.</span><span class="sxs-lookup"><span data-stu-id="3aa16-393">A `Date` object specifying the start date and time of the appointment.</span></span> |
| `parameters.end` | <span data-ttu-id="3aa16-394">Date</span><span class="sxs-lookup"><span data-stu-id="3aa16-394">Date</span></span> | <span data-ttu-id="3aa16-395">Ein Objekt des Typs `Date`. Gibt das Enddatum und das Ende des Termins an.</span><span class="sxs-lookup"><span data-stu-id="3aa16-395">A `Date` object specifying the end date and time of the appointment.</span></span> |
| `parameters.location` | <span data-ttu-id="3aa16-396">String</span><span class="sxs-lookup"><span data-stu-id="3aa16-396">String</span></span> | <span data-ttu-id="3aa16-p119">Eine Zeichenfolge mit dem Ort für den Termin. Die Zeichenfolge ist auf maximal 255 Zeichen beschränkt.</span><span class="sxs-lookup"><span data-stu-id="3aa16-p119">A string containing the location of the appointment. The string is limited to a maximum of 255 characters.</span></span> |
| `parameters.resources` | <span data-ttu-id="3aa16-399">Array.&lt;String&gt;</span><span class="sxs-lookup"><span data-stu-id="3aa16-399">Array.&lt;String&gt;</span></span> | <span data-ttu-id="3aa16-p120">Ein Array aus Zeichenfolgen, die die für den Termin erforderlichen Ressourcen enthalten. Das Array darf maximal 100 Einträge enthalten.</span><span class="sxs-lookup"><span data-stu-id="3aa16-p120">An array of strings containing the resources required for the appointment. The array is limited to a maximum of 100 entries.</span></span> |
| `parameters.subject` | <span data-ttu-id="3aa16-402">Zeichenfolge</span><span class="sxs-lookup"><span data-stu-id="3aa16-402">String</span></span> | <span data-ttu-id="3aa16-p121">Eine Zeichenfolge mit dem Betreff für den Termin. Die Zeichenfolge ist auf maximal 255 Zeichen beschränkt.</span><span class="sxs-lookup"><span data-stu-id="3aa16-p121">A string containing the subject of the appointment. The string is limited to a maximum of 255 characters.</span></span> |
| `parameters.body` | <span data-ttu-id="3aa16-405">Zeichenfolge</span><span class="sxs-lookup"><span data-stu-id="3aa16-405">String</span></span> | <span data-ttu-id="3aa16-p122">Der Text des Termins. Der Textinhalt darf maximal 32 KB umfassen.</span><span class="sxs-lookup"><span data-stu-id="3aa16-p122">The body of the appointment. The body content is limited to a maximum size of 32 KB.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="3aa16-408">Anforderungen</span><span class="sxs-lookup"><span data-stu-id="3aa16-408">Requirements</span></span>

|<span data-ttu-id="3aa16-409">Anforderung</span><span class="sxs-lookup"><span data-stu-id="3aa16-409">Requirement</span></span>| <span data-ttu-id="3aa16-410">Wert</span><span class="sxs-lookup"><span data-stu-id="3aa16-410">Value</span></span>|
|---|---|
|[<span data-ttu-id="3aa16-411">Mindestversion des Postfachanforderungssatzes</span><span class="sxs-lookup"><span data-stu-id="3aa16-411">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="3aa16-412">1.0</span><span class="sxs-lookup"><span data-stu-id="3aa16-412">1.0</span></span>|
|[<span data-ttu-id="3aa16-413">Mindestberechtigungsstufe</span><span class="sxs-lookup"><span data-stu-id="3aa16-413">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="3aa16-414">ReadItem</span><span class="sxs-lookup"><span data-stu-id="3aa16-414">ReadItem</span></span>|
|[<span data-ttu-id="3aa16-415">Zutreffender Outlook-Modus</span><span class="sxs-lookup"><span data-stu-id="3aa16-415">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="3aa16-416">Lesen</span><span class="sxs-lookup"><span data-stu-id="3aa16-416">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="3aa16-417">Beispiel</span><span class="sxs-lookup"><span data-stu-id="3aa16-417">Example</span></span>

```
var start = new Date();
var end = new Date();
end.setHours(start.getHours() + 1);

Office.context.mailbox.displayNewAppointmentForm(
  {
    requiredAttendees: ['bob@contoso.com'],
    optionalAttendees: ['sam@contoso.com'],
    start: start,
    end: end,
    location: 'Home',
    resources: ['projector@contoso.com'],
    subject: 'meeting',
    body: 'Hello World!'
  });
```

#### <a name="displaynewmessageformparameters"></a><span data-ttu-id="3aa16-418">displayNewMessageForm(parameters)</span><span class="sxs-lookup"><span data-stu-id="3aa16-418">displayNewMessageForm(parameters)</span></span>

<span data-ttu-id="3aa16-419">Zeigt ein Formular zum Erstellen einer neuen Nachricht an.</span><span class="sxs-lookup"><span data-stu-id="3aa16-419">Displays a form for creating a new message.</span></span>

<span data-ttu-id="3aa16-420">Mit der `displayNewMessageForm`-Methode wird ein Formular geöffnet, mit dem der Benutzer eine neue Nachricht erstellen kann.</span><span class="sxs-lookup"><span data-stu-id="3aa16-420">The `displayNewMessageForm` method opens a form that enables the user to create a new message.</span></span> <span data-ttu-id="3aa16-421">Wenn Parameter angegeben wurden, werden die Felder im Nachrichtenformular automatisch mit dem Inhalt der Parameter ausgefüllt.</span><span class="sxs-lookup"><span data-stu-id="3aa16-421">If parameters are specified, the message form fields are automatically populated with the contents of the parameters.</span></span>

<span data-ttu-id="3aa16-422">Wenn einer der Parameter die angegebenen Größenbeschränkungen überschreitet oder wenn ein unbekannter Parametername angegeben wird, wird eine Ausnahme ausgelöst.</span><span class="sxs-lookup"><span data-stu-id="3aa16-422">If any of the parameters exceed the specified size limits, or if an unknown parameter name is specified, an exception is thrown.</span></span>

##### <a name="parameters"></a><span data-ttu-id="3aa16-423">Parameter:</span><span class="sxs-lookup"><span data-stu-id="3aa16-423">Parameters:</span></span>

> [!NOTE]
> <span data-ttu-id="3aa16-424">Alle Parameter sind optional.</span><span class="sxs-lookup"><span data-stu-id="3aa16-424">All parameters are optional.</span></span>

|<span data-ttu-id="3aa16-425">Name</span><span class="sxs-lookup"><span data-stu-id="3aa16-425">Name</span></span>| <span data-ttu-id="3aa16-426">Typ</span><span class="sxs-lookup"><span data-stu-id="3aa16-426">Type</span></span>| <span data-ttu-id="3aa16-427">Beschreibung</span><span class="sxs-lookup"><span data-stu-id="3aa16-427">Description</span></span>|
|---|---|---|
| `parameters` | <span data-ttu-id="3aa16-428">Object</span><span class="sxs-lookup"><span data-stu-id="3aa16-428">Object</span></span> | <span data-ttu-id="3aa16-429">Ein Wörterbuch mit Parametern, die die neue Nachricht beschreiben.</span><span class="sxs-lookup"><span data-stu-id="3aa16-429">A dictionary of parameters describing the new message.</span></span> |
| `parameters.toRecipients` | <span data-ttu-id="3aa16-430">Array.&lt;String&gt; &#124; Array.&lt;[EmailAddressDetails](/javascript/api/outlook_1_6/office.emailaddressdetails)&gt;</span><span class="sxs-lookup"><span data-stu-id="3aa16-430">Array.&lt;String&gt; &#124; Array.&lt;[EmailAddressDetails](/javascript/api/outlook_1_6/office.emailaddressdetails)&gt;</span></span> | <span data-ttu-id="3aa16-431">Ein Array aus Zeichenfolgen mit den E-Mail-Adressen oder ein Array mit einem `EmailAddressDetails`-Objekt für jeden der Empfänger im Feld „An“.</span><span class="sxs-lookup"><span data-stu-id="3aa16-431">An array of strings containing the email addresses or an array containing an `EmailAddressDetails` object for each of the recipients on the To line.</span></span> <span data-ttu-id="3aa16-432">Das Array darf maximal 100 Einträge enthalten.</span><span class="sxs-lookup"><span data-stu-id="3aa16-432">The array is limited to a maximum of 100 entries.</span></span> |
| `parameters.ccRecipients` | <span data-ttu-id="3aa16-433">Array.&lt;String&gt; &#124; Array.&lt;[EmailAddressDetails](/javascript/api/outlook_1_6/office.emailaddressdetails)&gt;</span><span class="sxs-lookup"><span data-stu-id="3aa16-433">Array.&lt;String&gt; &#124; Array.&lt;[EmailAddressDetails](/javascript/api/outlook_1_6/office.emailaddressdetails)&gt;</span></span> | <span data-ttu-id="3aa16-434">Ein Array aus Zeichenfolgen mit den E-Mail-Adressen oder ein Array mit einem `EmailAddressDetails`-Objekt für jeden der Empfänger im Feld „Cc“.</span><span class="sxs-lookup"><span data-stu-id="3aa16-434">An array of strings containing the email addresses or an array containing an `EmailAddressDetails` object for each of the recipients on the Cc line.</span></span> <span data-ttu-id="3aa16-435">Das Array darf maximal 100 Einträge enthalten.</span><span class="sxs-lookup"><span data-stu-id="3aa16-435">The array is limited to a maximum of 100 entries.</span></span> |
| `parameters.bccRecipients` | <span data-ttu-id="3aa16-436">Array.&lt;String&gt; &#124; Array.&lt;[EmailAddressDetails](/javascript/api/outlook_1_6/office.emailaddressdetails)&gt;</span><span class="sxs-lookup"><span data-stu-id="3aa16-436">Array.&lt;String&gt; &#124; Array.&lt;[EmailAddressDetails](/javascript/api/outlook_1_6/office.emailaddressdetails)&gt;</span></span> | <span data-ttu-id="3aa16-437">Ein Array aus Zeichenfolgen mit den E-Mail-Adressen oder ein Array mit einem `EmailAddressDetails`-Objekt für jeden der Empfänger im Feld „Bcc“.</span><span class="sxs-lookup"><span data-stu-id="3aa16-437">An array of strings containing the email addresses or an array containing an `EmailAddressDetails` object for each of the recipients on the Bcc line.</span></span> <span data-ttu-id="3aa16-438">Das Array darf maximal 100 Einträge enthalten.</span><span class="sxs-lookup"><span data-stu-id="3aa16-438">The array is limited to a maximum of 100 entries.</span></span> |
| `parameters.subject` | <span data-ttu-id="3aa16-439">Zeichenfolge</span><span class="sxs-lookup"><span data-stu-id="3aa16-439">String</span></span> | <span data-ttu-id="3aa16-440">Eine Zeichenfolge mit dem Betreff für die Nachricht.</span><span class="sxs-lookup"><span data-stu-id="3aa16-440">A string containing the subject of the message.</span></span> <span data-ttu-id="3aa16-441">Die Zeichenfolge ist auf maximal 255 Zeichen beschränkt.</span><span class="sxs-lookup"><span data-stu-id="3aa16-441">The string is limited to a maximum of 255 characters.</span></span> |
| `parameters.htmlBody` | <span data-ttu-id="3aa16-442">Zeichenfolge</span><span class="sxs-lookup"><span data-stu-id="3aa16-442">String</span></span> | <span data-ttu-id="3aa16-443">Der HTML-Text der Nachricht.</span><span class="sxs-lookup"><span data-stu-id="3aa16-443">The HTML body of the message.</span></span> <span data-ttu-id="3aa16-444">Der Textinhalt darf maximal 32 KB umfassen.</span><span class="sxs-lookup"><span data-stu-id="3aa16-444">The body content is limited to a maximum size of 32 KB.</span></span> |
| `parameters.attachments` | <span data-ttu-id="3aa16-445">Array.&lt;Object&gt;</span><span class="sxs-lookup"><span data-stu-id="3aa16-445">Array.&lt;Object&gt;</span></span> | <span data-ttu-id="3aa16-446">Ein Array mit JSON-Objekten, die Datei- oder Elementanlagen sind.</span><span class="sxs-lookup"><span data-stu-id="3aa16-446">An array of JSON objects that are either file or item attachments.</span></span> |
| `parameters.attachments.type` | <span data-ttu-id="3aa16-447">Zeichenfolge</span><span class="sxs-lookup"><span data-stu-id="3aa16-447">String</span></span> | <span data-ttu-id="3aa16-p129">Gibt den Typ der Anlage an. Muss `file` für eine Dateianlage oder `item` für eine Elementanlage sein.</span><span class="sxs-lookup"><span data-stu-id="3aa16-p129">Indicates the type of attachment. Must be `file` for a file attachment or `item` for an item attachment.</span></span> |
| `parameters.attachments.name` | <span data-ttu-id="3aa16-450">Zeichenfolge</span><span class="sxs-lookup"><span data-stu-id="3aa16-450">String</span></span> | <span data-ttu-id="3aa16-451">Eine Zeichenfolge, die den Namen der Anlage mit bis zu 255 Zeichen enthält.</span><span class="sxs-lookup"><span data-stu-id="3aa16-451">A string that contains the name of the attachment, up to 255 characters in length.</span></span>|
| `parameters.attachments.url` | <span data-ttu-id="3aa16-452">Zeichenfolge</span><span class="sxs-lookup"><span data-stu-id="3aa16-452">String</span></span> | <span data-ttu-id="3aa16-p130">Wird nur verwendet, wenn `type` auf `file` gesetzt ist. Der URI des Speicherorts für die Datei.</span><span class="sxs-lookup"><span data-stu-id="3aa16-p130">Only used if `type` is set to `file`. The URI of the location for the file.</span></span> |
| `parameters.attachments.isInline` | <span data-ttu-id="3aa16-455">Boolean</span><span class="sxs-lookup"><span data-stu-id="3aa16-455">Boolean</span></span> | <span data-ttu-id="3aa16-p131">Wird nur verwendet, wenn `type` auf `file` gesetzt ist. Wenn `true`: Zeigt an, dass die Anlage inline im Nachrichtentext angezeigt wird und nicht in der Anlagenlisten angezeigt werden soll.</span><span class="sxs-lookup"><span data-stu-id="3aa16-p131">Only used if `type` is set to `file`. If `true`, indicates that the attachment will be shown inline in the message body, and should not be displayed in the attachment list.</span></span> |
| `parameters.attachments.itemId` | <span data-ttu-id="3aa16-458">Zeichenfolge</span><span class="sxs-lookup"><span data-stu-id="3aa16-458">String</span></span> | <span data-ttu-id="3aa16-459">Nur verwendet, wenn `type` auf `item` festgelegt ist.</span><span class="sxs-lookup"><span data-stu-id="3aa16-459">Only used if `type` is set to `item`.</span></span> <span data-ttu-id="3aa16-460">Der EWS-Element-Id der vorhandenen e-Mail, den Sie die neue Nachricht zuordnen möchten.</span><span class="sxs-lookup"><span data-stu-id="3aa16-460">The EWS item id of the existing e-mail you want to attach to the new message.</span></span> <span data-ttu-id="3aa16-461">Dies ist eine Zeichenfolge bis zu 100 Zeichen.</span><span class="sxs-lookup"><span data-stu-id="3aa16-461">This is a string up to 100 characters.</span></span> |


##### <a name="requirements"></a><span data-ttu-id="3aa16-462">Anforderungen</span><span class="sxs-lookup"><span data-stu-id="3aa16-462">Requirements</span></span>

|<span data-ttu-id="3aa16-463">Anforderung</span><span class="sxs-lookup"><span data-stu-id="3aa16-463">Requirement</span></span>| <span data-ttu-id="3aa16-464">Wert</span><span class="sxs-lookup"><span data-stu-id="3aa16-464">Value</span></span>|
|---|---|
|[<span data-ttu-id="3aa16-465">Mindestversion des Postfachanforderungssatzes</span><span class="sxs-lookup"><span data-stu-id="3aa16-465">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="3aa16-466">1.6</span><span class="sxs-lookup"><span data-stu-id="3aa16-466">1.6</span></span> |
|[<span data-ttu-id="3aa16-467">Mindestberechtigungsstufe</span><span class="sxs-lookup"><span data-stu-id="3aa16-467">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="3aa16-468">ReadItem</span><span class="sxs-lookup"><span data-stu-id="3aa16-468">ReadItem</span></span>|
|[<span data-ttu-id="3aa16-469">Zutreffender Outlook-Modus</span><span class="sxs-lookup"><span data-stu-id="3aa16-469">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="3aa16-470">Lesen</span><span class="sxs-lookup"><span data-stu-id="3aa16-470">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="3aa16-471">Beispiel</span><span class="sxs-lookup"><span data-stu-id="3aa16-471">Example</span></span>

```
Office.context.mailbox.displayNewMessageForm(
  {
    toRecipients: Office.context.mailbox.item.to, // Copy the To line from current item
    ccRecipients: ['sam@contoso.com'],
    subject: 'Outlook add-ins are cool!',
    htmlBody: 'Hello <b>World</b>!<br/><img src="cid:image.png"></i>',
    attachments: [
      {
        type: 'file',
        name: 'image.png',
        url: 'http://contoso.com/image.png',
        isInline: true
      }
    ]
  });
```

#### <a name="getcallbacktokenasyncoptions-callback"></a><span data-ttu-id="3aa16-472">getCallbackTokenAsync([options], callback)</span><span class="sxs-lookup"><span data-stu-id="3aa16-472">getCallbackTokenAsync([options], callback)</span></span>

<span data-ttu-id="3aa16-473">Ruft eine Zeichenfolge mit einem Token ab, das zum Aufrufen von REST-APIs oder Exchange-Webdiensten verwendet wird.</span><span class="sxs-lookup"><span data-stu-id="3aa16-473">Gets a string that contains a token used to call REST APIs or Exchange Web Services.</span></span>

<span data-ttu-id="3aa16-p133">Die `getCallbackTokenAsync`-Methode führt einen asynchronen Aufruf zum Abruf eines nicht transparenten Tokens vom Exchange-Server aus, der das Postfach des Benutzers hostet. Die Gültigkeitsdauer des Rückruftokens beträgt 5 Minuten.</span><span class="sxs-lookup"><span data-stu-id="3aa16-p133">The `getCallbackTokenAsync` method makes an asynchronous call to get an opaque token from the Exchange Server that hosts the user's mailbox. The lifetime of the callback token is 5 minutes.</span></span>

> [!NOTE]
> <span data-ttu-id="3aa16-476">Es wird empfohlen, dass-add-ins anstelle von Exchange-Webdienste nach Möglichkeit die REST-APIs verwenden.</span><span class="sxs-lookup"><span data-stu-id="3aa16-476">It is recommended that add-ins use the REST APIs instead of Exchange Web Services whenever possible.</span></span> 

<span data-ttu-id="3aa16-477">**REST-Tokens**</span><span class="sxs-lookup"><span data-stu-id="3aa16-477">**REST Tokens**</span></span>

<span data-ttu-id="3aa16-p134">Wenn ein REST-Token angefordert wird (`options.isRest = true`), kann das resultierende Token nicht zur Authentifizierung von Aufrufen der Exchange-Webdienste verwendet werden. Der Bereich des Tokens ist auf schreibgeschützten Zugriff auf das aktuelle Element und dessen Anlagen beschränkt, es sei denn im Manifest des Add-Ins ist die Berechtigung [`ReadWriteMailbox`](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions#readwritemailbox-permission) angegeben. Ist die Berechtigung `ReadWriteMailbox` angegeben, gewährt das resultierende Token Lese-/Schreibzugriff auf E-Mails, Kalender und Kontakte und erlaubt das Senden von E-Mails.</span><span class="sxs-lookup"><span data-stu-id="3aa16-p134">When a REST token is requested (`options.isRest = true`), the resulting token will not work to authenticate Exchange Web Services calls. The token will be limited in scope to read-only access to the current item and its attachments, unless the add-in has specified the [`ReadWriteMailbox`](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions#readwritemailbox-permission) permission in its manifest. If the `ReadWriteMailbox` permission is specified, the resulting token will grant read/write access to mail, calendar, and contacts, including the ability to send mail.</span></span>

<span data-ttu-id="3aa16-481">Das Add-In sollte die Eigenschaft `restUrl` verwenden, um die korrekte URL für REST-API-Aufrufe zu ermitteln.</span><span class="sxs-lookup"><span data-stu-id="3aa16-481">The add-in should use the `restUrl` property to determine the correct URL to use when making REST API calls.</span></span>

<span data-ttu-id="3aa16-482">**EWS-Tokens**</span><span class="sxs-lookup"><span data-stu-id="3aa16-482">**EWS Tokens**</span></span>

<span data-ttu-id="3aa16-p135">Wenn ein EWS-Token angefordert wird (`options.isRest = false`), kann das resultierende Token nicht zur Authentifizierung von REST-API-Aufrufen verwendet werden. Der Bereich des Tokens ist auf den Zugriff auf das aktuelle Element beschränkt.</span><span class="sxs-lookup"><span data-stu-id="3aa16-p135">When an EWS token is requested (`options.isRest = false`), the resulting token will not work to authenticate REST API calls. The token will be limited in scope to accessing the current item.</span></span>

<span data-ttu-id="3aa16-485">Das Add-In sollte die Eigenschaft `ewsUrl` verwenden, um die korrekte URL für EWS-Aufrufe zu ermitteln.</span><span class="sxs-lookup"><span data-stu-id="3aa16-485">The add-in should use the `ewsUrl` property to determine the correct URL to use when making EWS calls.</span></span>

##### <a name="parameters"></a><span data-ttu-id="3aa16-486">Parameter:</span><span class="sxs-lookup"><span data-stu-id="3aa16-486">Parameters:</span></span>

|<span data-ttu-id="3aa16-487">Name</span><span class="sxs-lookup"><span data-stu-id="3aa16-487">Name</span></span>| <span data-ttu-id="3aa16-488">Typ</span><span class="sxs-lookup"><span data-stu-id="3aa16-488">Type</span></span>| <span data-ttu-id="3aa16-489">Attribute</span><span class="sxs-lookup"><span data-stu-id="3aa16-489">Attributes</span></span>| <span data-ttu-id="3aa16-490">Beschreibung</span><span class="sxs-lookup"><span data-stu-id="3aa16-490">Description</span></span>|
|---|---|---|---|
| `options` | <span data-ttu-id="3aa16-491">Objekt</span><span class="sxs-lookup"><span data-stu-id="3aa16-491">Object</span></span> | <span data-ttu-id="3aa16-492">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="3aa16-492">&lt;optional&gt;</span></span> | <span data-ttu-id="3aa16-493">Ein Objektliteral, das mindestens eine der folgenden Eigenschaften enthält.</span><span class="sxs-lookup"><span data-stu-id="3aa16-493">An object literal that contains one or more of the following properties.</span></span> |
| `options.isRest` | <span data-ttu-id="3aa16-494">Boolean</span><span class="sxs-lookup"><span data-stu-id="3aa16-494">Boolean</span></span> |  <span data-ttu-id="3aa16-495">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="3aa16-495">&lt;optional&gt;</span></span> | <span data-ttu-id="3aa16-p136">Legt fest, ob das bereitgestellte Token für die Outlook-REST-APIs oder die Exchange-Webdienste verwendet wird. Der Standardwert ist `false`.</span><span class="sxs-lookup"><span data-stu-id="3aa16-p136">Determines if the token provided will be used for the Outlook REST APIs or Exchange Web Services. Default value is `false`.</span></span> |
| `options.asyncContext` | <span data-ttu-id="3aa16-498">Objekt</span><span class="sxs-lookup"><span data-stu-id="3aa16-498">Object</span></span> |  <span data-ttu-id="3aa16-499">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="3aa16-499">&lt;optional&gt;</span></span> | <span data-ttu-id="3aa16-500">Jegliche Statusdaten, die an die asynchrone Methode übergeben werden</span><span class="sxs-lookup"><span data-stu-id="3aa16-500">Any state data that is passed to the asynchronous method.</span></span> |
|`callback`| <span data-ttu-id="3aa16-501">function</span><span class="sxs-lookup"><span data-stu-id="3aa16-501">function</span></span>||<span data-ttu-id="3aa16-p137">Wenn die Methode abgeschlossen wird, wird die Funktion , die im `callback`-Parameter übergeben wird, mit einem einzelnen Parameter (`asyncResult`) aufgerufen, der ein [`AsyncResult`](/javascript/api/office/office.asyncresult)-Objekt darstellt. Das Token wird als Zeichenfolge in der `asyncResult.value`-Eigenschaft bereitgestellt.</span><span class="sxs-lookup"><span data-stu-id="3aa16-p137">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object. The token is provided as a string in the `asyncResult.value` property.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="3aa16-504">Anforderungen</span><span class="sxs-lookup"><span data-stu-id="3aa16-504">Requirements</span></span>

|<span data-ttu-id="3aa16-505">Anforderung</span><span class="sxs-lookup"><span data-stu-id="3aa16-505">Requirement</span></span>| <span data-ttu-id="3aa16-506">Wert</span><span class="sxs-lookup"><span data-stu-id="3aa16-506">Value</span></span>|
|---|---|
|[<span data-ttu-id="3aa16-507">Mindestversion des Postfachanforderungssatzes</span><span class="sxs-lookup"><span data-stu-id="3aa16-507">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="3aa16-508">1,5</span><span class="sxs-lookup"><span data-stu-id="3aa16-508">1.5</span></span> |
|[<span data-ttu-id="3aa16-509">Mindestberechtigungsstufe</span><span class="sxs-lookup"><span data-stu-id="3aa16-509">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="3aa16-510">ReadItem</span><span class="sxs-lookup"><span data-stu-id="3aa16-510">ReadItem</span></span>|
|[<span data-ttu-id="3aa16-511">Zutreffender Outlook-Modus</span><span class="sxs-lookup"><span data-stu-id="3aa16-511">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="3aa16-512">Verfassenmodus und Lesemodus</span><span class="sxs-lookup"><span data-stu-id="3aa16-512">Compose and read</span></span>|

##### <a name="example"></a><span data-ttu-id="3aa16-513">Beispiel</span><span class="sxs-lookup"><span data-stu-id="3aa16-513">Example</span></span>

```js
function getCallbackToken() {
  var options = {
    isRest: true,
    asyncContext: { message: 'Hello World!' }
  };

  Office.context.mailbox.getCallbackTokenAsync(options, cb);
}

function cb(asyncResult) {
  var token = asyncResult.value;
}
```

#### <a name="getcallbacktokenasynccallback-usercontext"></a><span data-ttu-id="3aa16-514">getCallbackTokenAsync(callback, [userContext])</span><span class="sxs-lookup"><span data-stu-id="3aa16-514">getCallbackTokenAsync(callback, [userContext])</span></span>

<span data-ttu-id="3aa16-515">Ruft eine Zeichenfolge ab, die einen Token enthält, der verwendet wird, um eine Anlage oder ein Element von einem Exchange Server abzurufen.</span><span class="sxs-lookup"><span data-stu-id="3aa16-515">Gets a string that contains a token used to get an attachment or item from an Exchange Server.</span></span>

<span data-ttu-id="3aa16-p138">Die `getCallbackTokenAsync`-Methode führt einen asynchronen Aufruf zum Abruf eines nicht transparenten Tokens vom Exchange-Server aus, der das Postfach des Benutzers hostet. Die Gültigkeitsdauer des Rückruftokens beträgt 5 Minuten.</span><span class="sxs-lookup"><span data-stu-id="3aa16-p138">The `getCallbackTokenAsync` method makes an asynchronous call to get an opaque token from the Exchange Server that hosts the user's mailbox. The lifetime of the callback token is 5 minutes.</span></span>

<span data-ttu-id="3aa16-p139">Sie können das Token und einen Anlagen- oder einen Elementbezeichner an ein Drittanbietersystem weitergeben. Das Drittanbietersystem verwendet das Token als Trägerautorisierungstoken, um den EWS-Vorgang (Exchange-Webdienste) [GetAttachment](https://docs.microsoft.com/exchange/client-developer/web-service-reference/getattachment-operation) oder [GetItem](https://docs.microsoft.com/exchange/client-developer/web-service-reference/getitem-operation) aufzurufen und eine Anlage oder ein Element zurückzugeben. Sie können z. B. einen Remotedienst erstellen, um [Anlagen aus dem ausgewählten Element abzurufen](https://docs.microsoft.com/outlook/add-ins/get-attachments-of-an-outlook-item).</span><span class="sxs-lookup"><span data-stu-id="3aa16-p139">You can pass the token and an attachment identifier or item identifier to a third-party system. The third-party system uses the token as a bearer authorization token to call the Exchange Web Services (EWS) [GetAttachment](https://docs.microsoft.com/exchange/client-developer/web-service-reference/getattachment-operation) or [GetItem](https://docs.microsoft.com/exchange/client-developer/web-service-reference/getitem-operation) operation to return an attachment or item. For example, you can create a remote service to [get attachments from the selected item](https://docs.microsoft.com/outlook/add-ins/get-attachments-of-an-outlook-item).</span></span>

<span data-ttu-id="3aa16-521">Ihre App muss über die im Manifest angegebene **ReadItem**-Berechtigung verfügen, um die `getCallbackTokenAsync`-Methode im Lesemodus aufzurufen.</span><span class="sxs-lookup"><span data-stu-id="3aa16-521">Your app must have the **ReadItem** permission specified in its manifest to call the `getCallbackTokenAsync` method in read mode.</span></span>

<span data-ttu-id="3aa16-p140">Im Verfassenmodus müssen Sie die [`saveAsync`](Office.context.mailbox.item.md#saveasyncoptions-callback)-Methode aufrufen, um einen Elementbezeichner an die `getCallbackTokenAsync`-Methode zu übergeben. Die App muss über **ReadWriteItem**-Berechtigungen verfügen, um die `saveAsync`-Methode aufzurufen.</span><span class="sxs-lookup"><span data-stu-id="3aa16-p140">In compose mode you must call the [`saveAsync`](Office.context.mailbox.item.md#saveasyncoptions-callback) method to get an item identifier to pass to the `getCallbackTokenAsync` method. Your app must have **ReadWriteItem** permissions to call the `saveAsync` method.</span></span>

##### <a name="parameters"></a><span data-ttu-id="3aa16-524">Parameter:</span><span class="sxs-lookup"><span data-stu-id="3aa16-524">Parameters:</span></span>

|<span data-ttu-id="3aa16-525">Name</span><span class="sxs-lookup"><span data-stu-id="3aa16-525">Name</span></span>| <span data-ttu-id="3aa16-526">Typ</span><span class="sxs-lookup"><span data-stu-id="3aa16-526">Type</span></span>| <span data-ttu-id="3aa16-527">Attribute</span><span class="sxs-lookup"><span data-stu-id="3aa16-527">Attributes</span></span>| <span data-ttu-id="3aa16-528">Beschreibung</span><span class="sxs-lookup"><span data-stu-id="3aa16-528">Description</span></span>|
|---|---|---|---|
|`callback`| <span data-ttu-id="3aa16-529">Funktion</span><span class="sxs-lookup"><span data-stu-id="3aa16-529">function</span></span>||<span data-ttu-id="3aa16-p141">Wenn die Methode abgeschlossen wird, wird die Funktion , die im `callback`-Parameter übergeben wird, mit einem einzelnen Parameter (`asyncResult`) aufgerufen, der ein [`AsyncResult`](/javascript/api/office/office.asyncresult)-Objekt darstellt. Das Token wird als Zeichenfolge in der `asyncResult.value`-Eigenschaft bereitgestellt.</span><span class="sxs-lookup"><span data-stu-id="3aa16-p141">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object. The token is provided as a string in the `asyncResult.value` property.</span></span>|
|`userContext`| <span data-ttu-id="3aa16-532">Objekt</span><span class="sxs-lookup"><span data-stu-id="3aa16-532">Object</span></span>| <span data-ttu-id="3aa16-533">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="3aa16-533">&lt;optional&gt;</span></span>|<span data-ttu-id="3aa16-534">Jegliche Zustandsdaten, die an die asynchrone Methode übergeben werden.</span><span class="sxs-lookup"><span data-stu-id="3aa16-534">Any state data that is passed to the asynchronous method.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="3aa16-535">Anforderungen</span><span class="sxs-lookup"><span data-stu-id="3aa16-535">Requirements</span></span>

|<span data-ttu-id="3aa16-536">Anforderung</span><span class="sxs-lookup"><span data-stu-id="3aa16-536">Requirement</span></span>| <span data-ttu-id="3aa16-537">Wert</span><span class="sxs-lookup"><span data-stu-id="3aa16-537">Value</span></span>|
|---|---|
|[<span data-ttu-id="3aa16-538">Mindestversion des Postfachanforderungssatzes</span><span class="sxs-lookup"><span data-stu-id="3aa16-538">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="3aa16-539">1.3</span><span class="sxs-lookup"><span data-stu-id="3aa16-539">1.3</span></span>|
|[<span data-ttu-id="3aa16-540">Mindestberechtigungsstufe</span><span class="sxs-lookup"><span data-stu-id="3aa16-540">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="3aa16-541">ReadItem</span><span class="sxs-lookup"><span data-stu-id="3aa16-541">ReadItem</span></span>|
|[<span data-ttu-id="3aa16-542">Zutreffender Outlook-Modus</span><span class="sxs-lookup"><span data-stu-id="3aa16-542">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="3aa16-543">Verfassenmodus und Lesemodus</span><span class="sxs-lookup"><span data-stu-id="3aa16-543">Compose and read</span></span>|

##### <a name="example"></a><span data-ttu-id="3aa16-544">Beispiel</span><span class="sxs-lookup"><span data-stu-id="3aa16-544">Example</span></span>

```js
function getCallbackToken() {
  Office.context.mailbox.getCallbackTokenAsync(cb);
}

function cb(asyncResult) {
  var token = asyncResult.value;
}
```

####  <a name="getuseridentitytokenasynccallback-usercontext"></a><span data-ttu-id="3aa16-545">getUserIdentityTokenAsync(callback, [userContext])</span><span class="sxs-lookup"><span data-stu-id="3aa16-545">getUserIdentityTokenAsync(callback, [userContext])</span></span>

<span data-ttu-id="3aa16-546">Ruft ein Token ab, das den Benutzer und das Office-Add-In identifiziert.</span><span class="sxs-lookup"><span data-stu-id="3aa16-546">Gets a token identifying the user and the Office Add-in.</span></span>

<span data-ttu-id="3aa16-547">Die `getUserIdentityTokenAsync`-Methode gibt ein Token zurück, das Sie dazu verwenden können, um das [Add-In und den Benutzer mit einem Drittanbietersystem zu identifizieren und zu authentifizieren](https://docs.microsoft.com/outlook/add-ins/authentication).</span><span class="sxs-lookup"><span data-stu-id="3aa16-547">The `getUserIdentityTokenAsync` method returns a token that you can use to identify and [authenticate the add-in and user with a third-party system](https://docs.microsoft.com/outlook/add-ins/authentication).</span></span>

##### <a name="parameters"></a><span data-ttu-id="3aa16-548">Parameter:</span><span class="sxs-lookup"><span data-stu-id="3aa16-548">Parameters:</span></span>

|<span data-ttu-id="3aa16-549">Name</span><span class="sxs-lookup"><span data-stu-id="3aa16-549">Name</span></span>| <span data-ttu-id="3aa16-550">Typ</span><span class="sxs-lookup"><span data-stu-id="3aa16-550">Type</span></span>| <span data-ttu-id="3aa16-551">Attribute</span><span class="sxs-lookup"><span data-stu-id="3aa16-551">Attributes</span></span>| <span data-ttu-id="3aa16-552">Beschreibung</span><span class="sxs-lookup"><span data-stu-id="3aa16-552">Description</span></span>|
|---|---|---|---|
|`callback`| <span data-ttu-id="3aa16-553">Funktion</span><span class="sxs-lookup"><span data-stu-id="3aa16-553">function</span></span>||<span data-ttu-id="3aa16-554">Wenn die Methode abgeschlossen wird, wird die Funktion , die im `callback`-Parameter übergeben wird, mit einem einzelnen Parameter (`asyncResult`) aufgerufen, der ein [`AsyncResult`](/javascript/api/office/office.asyncresult)-Objekt darstellt.</span><span class="sxs-lookup"><span data-stu-id="3aa16-554">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span><br/><br/><span data-ttu-id="3aa16-555">Das Token wird als Zeichenfolge in der `asyncResult.value`-Eigenschaft bereitgestellt.</span><span class="sxs-lookup"><span data-stu-id="3aa16-555">The token is provided as a string in the `asyncResult.value` property.</span></span>|
|`userContext`| <span data-ttu-id="3aa16-556">Objekt</span><span class="sxs-lookup"><span data-stu-id="3aa16-556">Object</span></span>| <span data-ttu-id="3aa16-557">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="3aa16-557">&lt;optional&gt;</span></span>|<span data-ttu-id="3aa16-558">Jegliche Zustandsdaten, die an die asynchrone Methode übergeben werden.</span><span class="sxs-lookup"><span data-stu-id="3aa16-558">Any state data that is passed to the asynchronous method.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="3aa16-559">Anforderungen</span><span class="sxs-lookup"><span data-stu-id="3aa16-559">Requirements</span></span>

|<span data-ttu-id="3aa16-560">Anforderung</span><span class="sxs-lookup"><span data-stu-id="3aa16-560">Requirement</span></span>| <span data-ttu-id="3aa16-561">Wert</span><span class="sxs-lookup"><span data-stu-id="3aa16-561">Value</span></span>|
|---|---|
|[<span data-ttu-id="3aa16-562">Mindestversion des Postfachanforderungssatzes</span><span class="sxs-lookup"><span data-stu-id="3aa16-562">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="3aa16-563">1.0</span><span class="sxs-lookup"><span data-stu-id="3aa16-563">1.0</span></span>|
|[<span data-ttu-id="3aa16-564">Mindestberechtigungsstufe</span><span class="sxs-lookup"><span data-stu-id="3aa16-564">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="3aa16-565">ReadItem</span><span class="sxs-lookup"><span data-stu-id="3aa16-565">ReadItem</span></span>|
|[<span data-ttu-id="3aa16-566">Zutreffender Outlook-Modus</span><span class="sxs-lookup"><span data-stu-id="3aa16-566">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="3aa16-567">Verfassen oder Lesen</span><span class="sxs-lookup"><span data-stu-id="3aa16-567">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="3aa16-568">Beispiel</span><span class="sxs-lookup"><span data-stu-id="3aa16-568">Example</span></span>

```js
function getIdentityToken() {
  Office.context.mailbox.getUserIdentityTokenAsync(cb);
}

function cb(asyncResult) {
  var token = asyncResult.value;
}
```

####  <a name="makeewsrequestasyncdata-callback-usercontext"></a><span data-ttu-id="3aa16-569">makeEwsRequestAsync(data, callback, [userContext])</span><span class="sxs-lookup"><span data-stu-id="3aa16-569">makeEwsRequestAsync(data, callback, [userContext])</span></span>

<span data-ttu-id="3aa16-570">Richtet eine asynchrone Anforderung an einen EWS-Dienst (Exchange Web Services) auf dem Exchange-Server, der das Postfach des Benutzers hostet.</span><span class="sxs-lookup"><span data-stu-id="3aa16-570">Makes an asynchronous request to an Exchange Web Services (EWS) service on the Exchange server that hosts the user’s mailbox.</span></span>

> [!NOTE]
> <span data-ttu-id="3aa16-571">Diese Methode wird in den folgenden Szenarien nicht unterstützt.</span><span class="sxs-lookup"><span data-stu-id="3aa16-571">This method is not supported in the following scenarios.</span></span>
> - <span data-ttu-id="3aa16-572">In Outlook für iOS oder Outlook für Android</span><span class="sxs-lookup"><span data-stu-id="3aa16-572">In Outlook for iOS or Outlook for Android</span></span>
> - <span data-ttu-id="3aa16-573">Wenn das Add-In geladen wird, in einem Postfach Google Mail</span><span class="sxs-lookup"><span data-stu-id="3aa16-573">When the add-in is loaded in a Gmail mailbox</span></span>
> 
> <span data-ttu-id="3aa16-574">In diesen Fällen sollte-add-ins [mithilfe von REST-APIs](https://docs.microsoft.com/outlook/add-ins/use-rest-api) , stattdessen Zugriff auf das Postfach des Benutzers.</span><span class="sxs-lookup"><span data-stu-id="3aa16-574">In these cases, add-ins should [use REST APIs](https://docs.microsoft.com/outlook/add-ins/use-rest-api) to access the user's mailbox instead.</span></span>

<span data-ttu-id="3aa16-575">Die `makeEwsRequestAsync`-Methode sendet eine EWS-Anforderung für das Add-In zu Exchange.</span><span class="sxs-lookup"><span data-stu-id="3aa16-575">The `makeEwsRequestAsync` method sends an EWS request on behalf of the add-in to Exchange.</span></span> <span data-ttu-id="3aa16-576">Eine Liste der unterstützten EWS-Vorgänge finden Sie unter [Aufrufen von Webdiensten aus einem Outlook-add-in](https://docs.microsoft.com/outlook/add-ins/web-services#ews-operations-that-add-ins-support) .</span><span class="sxs-lookup"><span data-stu-id="3aa16-576">See [Call web services from an Outlook add-in](https://docs.microsoft.com/outlook/add-ins/web-services#ews-operations-that-add-ins-support) for a list of the supported EWS operations.</span></span>

<span data-ttu-id="3aa16-577">Sie können keine Elemente, die Ordnern zugeordnet sind, mit der `makeEwsRequestAsync`-Methode anfordern.</span><span class="sxs-lookup"><span data-stu-id="3aa16-577">You cannot request Folder Associated Items with the `makeEwsRequestAsync` method.</span></span>

<span data-ttu-id="3aa16-578">Die XML-Anfrage muss UTF-8-Codierung angeben.</span><span class="sxs-lookup"><span data-stu-id="3aa16-578">The XML request must specify UTF-8 encoding.</span></span>

```
<?xml version="1.0" encoding="utf-8"?>
```

<span data-ttu-id="3aa16-p143">Das Add-In muss über die **ReadWriteMailbox**-Berechtigung verfügen, um die `makeEwsRequestAsync`-Methode zu verwenden. Informationen zur **ReadWriteMailbox**-Berechtigung und zu den EWS-Vorgängen, die Sie mit der `makeEwsRequestAsync`-Methode aufrufen können, finden Sie unter [Angeben von Berechtigungen für den E-Mail-Add-In-Zugriff auf das Benutzerpostfach](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions).</span><span class="sxs-lookup"><span data-stu-id="3aa16-p143">Your add-in must have the **ReadWriteMailbox** permission to use the `makeEwsRequestAsync` method. For information about using the **ReadWriteMailbox** permission and the EWS operations that you can call with the `makeEwsRequestAsync` method, see [Specify permissions for mail add-in access to the user's mailbox](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions).</span></span>

> [!NOTE]
> <span data-ttu-id="3aa16-581">Der Serveradministrator muss festgelegt `OAuthAuthentication` auf "true" auf dem Client Access Server EWS-Verzeichnis so aktivieren Sie die `makeEwsRequestAsync` -Methode, um EWS stellen anfordert.</span><span class="sxs-lookup"><span data-stu-id="3aa16-581">The server administrator must set `OAuthAuthentication` to true on the Client Access Server EWS directory to enable the `makeEwsRequestAsync` method to make EWS requests.</span></span>

##### <a name="version-differences"></a><span data-ttu-id="3aa16-582">Versionsunterschiede</span><span class="sxs-lookup"><span data-stu-id="3aa16-582">Version differences</span></span>

<span data-ttu-id="3aa16-583">Wenn Sie die `makeEwsRequestAsync`-Methode in Mail-Apps verwenden, die in älteren Outlook-Versionen als Version 15.0.4535.1004 ausgeführt werden, sollten Sie den Codierungswert auf `ISO-8859-1` festlegen.</span><span class="sxs-lookup"><span data-stu-id="3aa16-583">When you use the `makeEwsRequestAsync` method in mail apps running in Outlook versions earlier than version 15.0.4535.1004, you should set the encoding value to `ISO-8859-1`.</span></span>

```
<?xml version="1.0" encoding="iso-8859-1"?>
```

<span data-ttu-id="3aa16-p144">Sie müssen den Codierungswert nicht festlegen, wenn Ihre Mail-App in Outlook im Web ausgeführt wird. Sie können mithilfe der mailbox.diagnostics.hostName-Eigenschaft ermitteln, ob Ihre Mail-App in Outlook oder Outlook im Web ausgeführt wird. Sie können mithilfe der mailbox.diagnostics.hostVersion-Eigenschaft ermitteln, welche Version von Outlook ausgeführt wird.</span><span class="sxs-lookup"><span data-stu-id="3aa16-p144">You do not need to set the encoding value when your mail app is running in Outlook on the web. You can determine whether your mail app is running in Outlook or Outlook on the web by using the mailbox.diagnostics.hostName property. You can determine what version of Outlook is running by using the mailbox.diagnostics.hostVersion property.</span></span>

##### <a name="parameters"></a><span data-ttu-id="3aa16-587">Parameter:</span><span class="sxs-lookup"><span data-stu-id="3aa16-587">Parameters:</span></span>

|<span data-ttu-id="3aa16-588">Name</span><span class="sxs-lookup"><span data-stu-id="3aa16-588">Name</span></span>| <span data-ttu-id="3aa16-589">Typ</span><span class="sxs-lookup"><span data-stu-id="3aa16-589">Type</span></span>| <span data-ttu-id="3aa16-590">Attribute</span><span class="sxs-lookup"><span data-stu-id="3aa16-590">Attributes</span></span>| <span data-ttu-id="3aa16-591">Beschreibung</span><span class="sxs-lookup"><span data-stu-id="3aa16-591">Description</span></span>|
|---|---|---|---|
|`data`| <span data-ttu-id="3aa16-592">Zeichenfolge</span><span class="sxs-lookup"><span data-stu-id="3aa16-592">String</span></span>||<span data-ttu-id="3aa16-593">Die EWS-Anforderung.</span><span class="sxs-lookup"><span data-stu-id="3aa16-593">The EWS request.</span></span>|
|`callback`| <span data-ttu-id="3aa16-594">Funktion</span><span class="sxs-lookup"><span data-stu-id="3aa16-594">function</span></span>||<span data-ttu-id="3aa16-595">Wenn die Methode abgeschlossen wird, wird die Funktion , die im `callback`-Parameter übergeben wird, mit einem einzelnen Parameter (`asyncResult`) aufgerufen, der ein [`AsyncResult`](/javascript/api/office/office.asyncresult)-Objekt darstellt.</span><span class="sxs-lookup"><span data-stu-id="3aa16-595">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span><br/><br/><span data-ttu-id="3aa16-596">Das XML-Ergebnis des EWS-Aufrufs wird als Zeichenfolge in der `asyncResult.value`-Eigenschaft bereitgestellt.</span><span class="sxs-lookup"><span data-stu-id="3aa16-596">The XML result of the EWS call is provided as a string in the `asyncResult.value` property.</span></span> <span data-ttu-id="3aa16-597">Wenn das Ergebnis 1 MB Größe überschreitet, wird stattdessen eine Fehlermeldung zurückgegeben.</span><span class="sxs-lookup"><span data-stu-id="3aa16-597">If the result exceeds 1 MB in size, an error message is returned instead.</span></span>|
|`userContext`| <span data-ttu-id="3aa16-598">Objekt</span><span class="sxs-lookup"><span data-stu-id="3aa16-598">Object</span></span>| <span data-ttu-id="3aa16-599">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="3aa16-599">&lt;optional&gt;</span></span>|<span data-ttu-id="3aa16-600">Jegliche Zustandsdaten, die an die asynchrone Methode übergeben werden.</span><span class="sxs-lookup"><span data-stu-id="3aa16-600">Any state data that is passed to the asynchronous method.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="3aa16-601">Anforderungen</span><span class="sxs-lookup"><span data-stu-id="3aa16-601">Requirements</span></span>

|<span data-ttu-id="3aa16-602">Anforderung</span><span class="sxs-lookup"><span data-stu-id="3aa16-602">Requirement</span></span>| <span data-ttu-id="3aa16-603">Wert</span><span class="sxs-lookup"><span data-stu-id="3aa16-603">Value</span></span>|
|---|---|
|[<span data-ttu-id="3aa16-604">Mindestversion des Postfachanforderungssatzes</span><span class="sxs-lookup"><span data-stu-id="3aa16-604">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="3aa16-605">1.0</span><span class="sxs-lookup"><span data-stu-id="3aa16-605">1.0</span></span>|
|[<span data-ttu-id="3aa16-606">Mindestberechtigungsstufe</span><span class="sxs-lookup"><span data-stu-id="3aa16-606">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="3aa16-607">ReadWriteMailbox</span><span class="sxs-lookup"><span data-stu-id="3aa16-607">ReadWriteMailbox</span></span>|
|[<span data-ttu-id="3aa16-608">Zutreffender Outlook-Modus</span><span class="sxs-lookup"><span data-stu-id="3aa16-608">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="3aa16-609">Verfassen oder Lesen</span><span class="sxs-lookup"><span data-stu-id="3aa16-609">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="3aa16-610">Beispiel</span><span class="sxs-lookup"><span data-stu-id="3aa16-610">Example</span></span>

<span data-ttu-id="3aa16-611">Das folgende Beispiel ruft `makeEwsRequestAsync` zum Verwenden des `GetItem`-Vorgangs auf, um den Betreff eines Elements abzurufen.</span><span class="sxs-lookup"><span data-stu-id="3aa16-611">The following example calls `makeEwsRequestAsync` to use the `GetItem` operation to get the subject of an item.</span></span>

```js
function getSubjectRequest(id) {
   // Return a GetItem operation request for the subject of the specified item.
   var request =
    '<?xml version="1.0" encoding="utf-8"?>' +
    '<soap:Envelope xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance"' +
    '               xmlns:xsd="http://www.w3.org/2001/XMLSchema"' +
    '               xmlns:soap="http://schemas.xmlsoap.org/soap/envelope/"' +
    '               xmlns:t="http://schemas.microsoft.com/exchange/services/2006/types">' +
    '  <soap:Header>' +
    '    <RequestServerVersion Version="Exchange2013" xmlns="http://schemas.microsoft.com/exchange/services/2006/types" soap:mustUnderstand="0" />' +
    '  </soap:Header>' +
    '  <soap:Body>' +
    '    <GetItem xmlns="http://schemas.microsoft.com/exchange/services/2006/messages">' +
    '      <ItemShape>' +
    '        <t:BaseShape>IdOnly</t:BaseShape>' +
    '        <t:AdditionalProperties>' +
    '            <t:FieldURI FieldURI="item:Subject"/>' +
    '        </t:AdditionalProperties>' +
    '      </ItemShape>' +
    '      <ItemIds><t:ItemId Id="' + id + '"/></ItemIds>' +
    '    </GetItem>' +
    '  </soap:Body>' +
    '</soap:Envelope>';

   return request;
}

function sendRequest() {
   // Create a local variable that contains the mailbox.
   Office.context.mailbox.makeEwsRequestAsync(
    getSubjectRequest(mailbox.item.itemId), callback);
}

function callback(asyncResult)  {
   var result = asyncResult.value;
   var context = asyncResult.asyncContext;

   // Process the returned response here.
}
```