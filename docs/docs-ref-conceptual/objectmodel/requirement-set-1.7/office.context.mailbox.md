
# <a name="mailbox"></a><span data-ttu-id="99b1d-101">mailbox</span><span class="sxs-lookup"><span data-stu-id="99b1d-101">mailbox</span></span>

### <span data-ttu-id="99b1d-p101">[Office](Office.md)[.context](Office.context.md). mailbox</span><span class="sxs-lookup"><span data-stu-id="99b1d-p101">[Office](Office.md)[.context](Office.context.md). mailbox</span></span>

<span data-ttu-id="99b1d-104">Ermöglicht den Zugriff auf das Objektmodell von Outlook-add-in für Microsoft Outlook und Microsoft Outlook im Web.</span><span class="sxs-lookup"><span data-stu-id="99b1d-104">Provides access to the Outlook add-in object model for Microsoft Outlook and Microsoft Outlook on the web.</span></span>

##### <a name="requirements"></a><span data-ttu-id="99b1d-105">Anforderungen</span><span class="sxs-lookup"><span data-stu-id="99b1d-105">Requirements</span></span>

|<span data-ttu-id="99b1d-106">Anforderung</span><span class="sxs-lookup"><span data-stu-id="99b1d-106">Requirement</span></span>| <span data-ttu-id="99b1d-107">Wert</span><span class="sxs-lookup"><span data-stu-id="99b1d-107">Value</span></span>|
|---|---|
|[<span data-ttu-id="99b1d-108">Mindestversion des Postfachanforderungssatzes</span><span class="sxs-lookup"><span data-stu-id="99b1d-108">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="99b1d-109">1.0</span><span class="sxs-lookup"><span data-stu-id="99b1d-109">1.0</span></span>|
|[<span data-ttu-id="99b1d-110">Mindestberechtigungsstufe</span><span class="sxs-lookup"><span data-stu-id="99b1d-110">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="99b1d-111">Eingeschränkt</span><span class="sxs-lookup"><span data-stu-id="99b1d-111">Restricted</span></span>|
|[<span data-ttu-id="99b1d-112">Zutreffender Outlook-Modus</span><span class="sxs-lookup"><span data-stu-id="99b1d-112">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="99b1d-113">Verfassen oder Lesen</span><span class="sxs-lookup"><span data-stu-id="99b1d-113">Compose or read</span></span>|

##### <a name="members-and-methods"></a><span data-ttu-id="99b1d-114">Elemente und Methoden</span><span class="sxs-lookup"><span data-stu-id="99b1d-114">Members and methods</span></span>

| <span data-ttu-id="99b1d-115">Element</span><span class="sxs-lookup"><span data-stu-id="99b1d-115">Member</span></span> | <span data-ttu-id="99b1d-116">Typ</span><span class="sxs-lookup"><span data-stu-id="99b1d-116">Type</span></span> |
|--------|------|
| [<span data-ttu-id="99b1d-117">ewsUrl</span><span class="sxs-lookup"><span data-stu-id="99b1d-117">ewsUrl</span></span>](#ewsurl-string) | <span data-ttu-id="99b1d-118">Element</span><span class="sxs-lookup"><span data-stu-id="99b1d-118">Member</span></span> |
| [<span data-ttu-id="99b1d-119">restUrl</span><span class="sxs-lookup"><span data-stu-id="99b1d-119">restUrl</span></span>](#resturl-string) | <span data-ttu-id="99b1d-120">Element</span><span class="sxs-lookup"><span data-stu-id="99b1d-120">Member</span></span> |
| [<span data-ttu-id="99b1d-121">addHandlerAsync</span><span class="sxs-lookup"><span data-stu-id="99b1d-121">addHandlerAsync</span></span>](#addhandlerasynceventtype-handler-options-callback) | <span data-ttu-id="99b1d-122">Methode</span><span class="sxs-lookup"><span data-stu-id="99b1d-122">Method</span></span> |
| [<span data-ttu-id="99b1d-123">convertToEwsId</span><span class="sxs-lookup"><span data-stu-id="99b1d-123">convertToEwsId</span></span>](#converttoewsiditemid-restversion--string) | <span data-ttu-id="99b1d-124">Methode</span><span class="sxs-lookup"><span data-stu-id="99b1d-124">Method</span></span> |
| [<span data-ttu-id="99b1d-125">convertToLocalClientTime</span><span class="sxs-lookup"><span data-stu-id="99b1d-125">convertToLocalClientTime</span></span>](#converttolocalclienttimetimevalue--localclienttimejavascriptapioutlook17officelocalclienttime) | <span data-ttu-id="99b1d-126">Methode</span><span class="sxs-lookup"><span data-stu-id="99b1d-126">Method</span></span> |
| [<span data-ttu-id="99b1d-127">convertToRestId</span><span class="sxs-lookup"><span data-stu-id="99b1d-127">convertToRestId</span></span>](#converttorestiditemid-restversion--string) | <span data-ttu-id="99b1d-128">Methode</span><span class="sxs-lookup"><span data-stu-id="99b1d-128">Method</span></span> |
| [<span data-ttu-id="99b1d-129">convertToUtcClientTime</span><span class="sxs-lookup"><span data-stu-id="99b1d-129">convertToUtcClientTime</span></span>](#converttoutcclienttimeinput--date) | <span data-ttu-id="99b1d-130">Methode</span><span class="sxs-lookup"><span data-stu-id="99b1d-130">Method</span></span> |
| [<span data-ttu-id="99b1d-131">displayAppointmentForm</span><span class="sxs-lookup"><span data-stu-id="99b1d-131">displayAppointmentForm</span></span>](#displayappointmentformitemid) | <span data-ttu-id="99b1d-132">Methode</span><span class="sxs-lookup"><span data-stu-id="99b1d-132">Method</span></span> |
| [<span data-ttu-id="99b1d-133">displayMessageForm</span><span class="sxs-lookup"><span data-stu-id="99b1d-133">displayMessageForm</span></span>](#displaymessageformitemid) | <span data-ttu-id="99b1d-134">Methode</span><span class="sxs-lookup"><span data-stu-id="99b1d-134">Method</span></span> |
| [<span data-ttu-id="99b1d-135">displayNewAppointmentForm</span><span class="sxs-lookup"><span data-stu-id="99b1d-135">displayNewAppointmentForm</span></span>](#displaynewappointmentformparameters) | <span data-ttu-id="99b1d-136">Methode</span><span class="sxs-lookup"><span data-stu-id="99b1d-136">Method</span></span> |
| [<span data-ttu-id="99b1d-137">displayNewMessageForm</span><span class="sxs-lookup"><span data-stu-id="99b1d-137">displayNewMessageForm</span></span>](#displaynewmessageformparameters) | <span data-ttu-id="99b1d-138">Methode</span><span class="sxs-lookup"><span data-stu-id="99b1d-138">Method</span></span> |
| [<span data-ttu-id="99b1d-139">getCallbackTokenAsync</span><span class="sxs-lookup"><span data-stu-id="99b1d-139">getCallbackTokenAsync</span></span>](#getcallbacktokenasyncoptions-callback) | <span data-ttu-id="99b1d-140">Methode</span><span class="sxs-lookup"><span data-stu-id="99b1d-140">Method</span></span> |
| [<span data-ttu-id="99b1d-141">getCallbackTokenAsync</span><span class="sxs-lookup"><span data-stu-id="99b1d-141">getCallbackTokenAsync</span></span>](#getcallbacktokenasynccallback-usercontext) | <span data-ttu-id="99b1d-142">Methode</span><span class="sxs-lookup"><span data-stu-id="99b1d-142">Method</span></span> |
| [<span data-ttu-id="99b1d-143">getUserIdentityTokenAsync</span><span class="sxs-lookup"><span data-stu-id="99b1d-143">getUserIdentityTokenAsync</span></span>](#getuseridentitytokenasynccallback-usercontext) | <span data-ttu-id="99b1d-144">Methode</span><span class="sxs-lookup"><span data-stu-id="99b1d-144">Method</span></span> |
| [<span data-ttu-id="99b1d-145">makeEwsRequestAsync</span><span class="sxs-lookup"><span data-stu-id="99b1d-145">makeEwsRequestAsync</span></span>](#makeewsrequestasyncdata-callback-usercontext) | <span data-ttu-id="99b1d-146">Methode</span><span class="sxs-lookup"><span data-stu-id="99b1d-146">Method</span></span> |

### <a name="namespaces"></a><span data-ttu-id="99b1d-147">Namespaces</span><span class="sxs-lookup"><span data-stu-id="99b1d-147">Namespaces</span></span>

<span data-ttu-id="99b1d-148">[diagnostics](Office.context.mailbox.diagnostics.md): Stellt einem Outlook-Add-In Diagnoseinformationen bereit.</span><span class="sxs-lookup"><span data-stu-id="99b1d-148">[diagnostics](Office.context.mailbox.diagnostics.md): Provides diagnostic information to an Outlook add-in.</span></span>

<span data-ttu-id="99b1d-149">[item](Office.context.mailbox.item.md): Stellt Methoden und Eigenschaften für den Zugriff auf eine Nachricht oder einen Termins in einem Outlook-Add-In bereit.</span><span class="sxs-lookup"><span data-stu-id="99b1d-149">[item](Office.context.mailbox.item.md): Provides methods and properties for accessing a message or appointment in an Outlook add-in.</span></span>

<span data-ttu-id="99b1d-150">[userProfile](Office.context.mailbox.userProfile.md): Stellt Informationen zum Benutzer in einem Outlook-Add-In bereit.</span><span class="sxs-lookup"><span data-stu-id="99b1d-150">[userProfile](Office.context.mailbox.userProfile.md): Provides information about the user in an Outlook add-in.</span></span>

### <a name="members"></a><span data-ttu-id="99b1d-151">Elemente</span><span class="sxs-lookup"><span data-stu-id="99b1d-151">Members</span></span>

#### <a name="ewsurl-string"></a><span data-ttu-id="99b1d-152">ewsUrl :String</span><span class="sxs-lookup"><span data-stu-id="99b1d-152">ewsUrl :String</span></span>

<span data-ttu-id="99b1d-p102">Ruft die URL des EWS-Endpunkts (Exchange Web Services) für dieses E-Mail-Konto ab. Nur Lesemodus.</span><span class="sxs-lookup"><span data-stu-id="99b1d-p102">Gets the URL of the Exchange Web Services (EWS) endpoint for this email account. Read mode only.</span></span>

> [!NOTE]
> <span data-ttu-id="99b1d-155">Dieser Member wird in Outlook für iOS oder Outlook für Android nicht unterstützt.</span><span class="sxs-lookup"><span data-stu-id="99b1d-155">This member is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="99b1d-p103">Der `ewsUrl`-Wert kann durch einen Remotedienst zum Ausführen von EWS-Aufrufen an das Postfach des Benutzers verwendet werden. Sie können z. B. einen Remotedienst erstellen, um [Anlagen](https://docs.microsoft.com/outlook/add-ins/get-attachments-of-an-outlook-item) aus dem ausgewählten Element abzurufen.</span><span class="sxs-lookup"><span data-stu-id="99b1d-p103">The `ewsUrl` value can be used by a remote service to make EWS calls to the user's mailbox. For example, you can create a remote service to [get attachments from the selected item](https://docs.microsoft.com/outlook/add-ins/get-attachments-of-an-outlook-item).</span></span>

<span data-ttu-id="99b1d-158">Ihre App muss über die im Manifest angegebene **ReadItem**-Berechtigung verfügen, um das `ewsUrl`-Element im Lesemodus aufzurufen.</span><span class="sxs-lookup"><span data-stu-id="99b1d-158">Your app must have the **ReadItem** permission specified in its manifest to call the `ewsUrl` member in read mode.</span></span>

<span data-ttu-id="99b1d-p104">Im Verfassenmodus müssen Sie die [`saveAsync`](Office.context.mailbox.item.md#saveasyncoptions-callback)-Methode aufrufen, bevor Sie das `ewsUrl`-Element verwenden können. Die App muss über **ReadWriteItem**-Berechtigungen verfügen, um die `saveAsync`-Methode aufzurufen.</span><span class="sxs-lookup"><span data-stu-id="99b1d-p104">In compose mode you must call the [`saveAsync`](Office.context.mailbox.item.md#saveasyncoptions-callback) method before you can use the `ewsUrl` member. Your app must have **ReadWriteItem** permissions to call the `saveAsync` method.</span></span>

##### <a name="type"></a><span data-ttu-id="99b1d-161">Typ:</span><span class="sxs-lookup"><span data-stu-id="99b1d-161">Type:</span></span>

*   <span data-ttu-id="99b1d-162">String</span><span class="sxs-lookup"><span data-stu-id="99b1d-162">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="99b1d-163">Anforderungen</span><span class="sxs-lookup"><span data-stu-id="99b1d-163">Requirements</span></span>

|<span data-ttu-id="99b1d-164">Anforderung</span><span class="sxs-lookup"><span data-stu-id="99b1d-164">Requirement</span></span>| <span data-ttu-id="99b1d-165">Wert</span><span class="sxs-lookup"><span data-stu-id="99b1d-165">Value</span></span>|
|---|---|
|[<span data-ttu-id="99b1d-166">Mindestversion des Postfachanforderungssatzes</span><span class="sxs-lookup"><span data-stu-id="99b1d-166">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="99b1d-167">1.0</span><span class="sxs-lookup"><span data-stu-id="99b1d-167">1.0</span></span>|
|[<span data-ttu-id="99b1d-168">Mindestberechtigungsstufe</span><span class="sxs-lookup"><span data-stu-id="99b1d-168">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="99b1d-169">ReadItem</span><span class="sxs-lookup"><span data-stu-id="99b1d-169">ReadItem</span></span>|
|[<span data-ttu-id="99b1d-170">Zutreffender Outlook-Modus</span><span class="sxs-lookup"><span data-stu-id="99b1d-170">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="99b1d-171">Verfassen oder Lesen</span><span class="sxs-lookup"><span data-stu-id="99b1d-171">Compose or read</span></span>|

#### <a name="resturl-string"></a><span data-ttu-id="99b1d-172">restUrl :String</span><span class="sxs-lookup"><span data-stu-id="99b1d-172">restUrl :String</span></span>

<span data-ttu-id="99b1d-173">Ruft die URL des REST-Endpunkts für das betreffende E-Mail-Konto ab.</span><span class="sxs-lookup"><span data-stu-id="99b1d-173">Gets the URL of the REST endpoint for this email account.</span></span>

<span data-ttu-id="99b1d-174">Der Wert `restUrl` kann für [REST-API](https://docs.microsoft.com/outlook/rest/)-Aufrufe an das Postfach des Benutzers verwendet werden.</span><span class="sxs-lookup"><span data-stu-id="99b1d-174">The `restUrl` value can be used to make [REST API](https://docs.microsoft.com/outlook/rest/) calls to the user's mailbox.</span></span>

<span data-ttu-id="99b1d-175">Im Manifest der App muss die Berechtigung **ReadItem** angegeben sein, damit das Mitglied `restUrl` im Lesemodus abgerufen werden kann.</span><span class="sxs-lookup"><span data-stu-id="99b1d-175">Your app must have the **ReadItem** permission specified in its manifest to call the `restUrl` member in read mode.</span></span>

<span data-ttu-id="99b1d-p105">Im Verfassenmodus müssen Sie die [`saveAsync`](Office.context.mailbox.item.md#saveasyncoptions-callback)-Methode aufrufen, bevor Sie das `restUrl`-Element verwenden können. Die App muss über **ReadWriteItem**-Berechtigungen verfügen, um die `saveAsync`-Methode aufzurufen.</span><span class="sxs-lookup"><span data-stu-id="99b1d-p105">In compose mode you must call the [`saveAsync`](Office.context.mailbox.item.md#saveasyncoptions-callback) method before you can use the `restUrl` member. Your app must have **ReadWriteItem** permissions to call the `saveAsync` method.</span></span>

##### <a name="type"></a><span data-ttu-id="99b1d-178">Typ:</span><span class="sxs-lookup"><span data-stu-id="99b1d-178">Type:</span></span>

*   <span data-ttu-id="99b1d-179">String</span><span class="sxs-lookup"><span data-stu-id="99b1d-179">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="99b1d-180">Anforderungen</span><span class="sxs-lookup"><span data-stu-id="99b1d-180">Requirements</span></span>

|<span data-ttu-id="99b1d-181">Anforderung</span><span class="sxs-lookup"><span data-stu-id="99b1d-181">Requirement</span></span>| <span data-ttu-id="99b1d-182">Wert</span><span class="sxs-lookup"><span data-stu-id="99b1d-182">Value</span></span>|
|---|---|
|[<span data-ttu-id="99b1d-183">Mindestversion des Postfachanforderungssatzes</span><span class="sxs-lookup"><span data-stu-id="99b1d-183">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="99b1d-184">1,5</span><span class="sxs-lookup"><span data-stu-id="99b1d-184">1.5</span></span> |
|[<span data-ttu-id="99b1d-185">Mindestberechtigungsstufe</span><span class="sxs-lookup"><span data-stu-id="99b1d-185">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="99b1d-186">ReadItem</span><span class="sxs-lookup"><span data-stu-id="99b1d-186">ReadItem</span></span>|
|[<span data-ttu-id="99b1d-187">Zutreffender Outlook-Modus</span><span class="sxs-lookup"><span data-stu-id="99b1d-187">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="99b1d-188">Verfassen oder Lesen</span><span class="sxs-lookup"><span data-stu-id="99b1d-188">Compose or read</span></span>|

### <a name="methods"></a><span data-ttu-id="99b1d-189">Methoden</span><span class="sxs-lookup"><span data-stu-id="99b1d-189">Methods</span></span>

####  <a name="addhandlerasynceventtype-handler-options-callback"></a><span data-ttu-id="99b1d-190">addHandlerAsync(eventType, handler, [options], [callback])</span><span class="sxs-lookup"><span data-stu-id="99b1d-190">addHandlerAsync(eventType, handler, [options], [callback])</span></span>

<span data-ttu-id="99b1d-191">Fügt einen Ereignishandler für ein unterstütztes Ereignis hinzu.</span><span class="sxs-lookup"><span data-stu-id="99b1d-191">Adds an event handler for a supported event.</span></span>

<span data-ttu-id="99b1d-192">Derzeit sind die unterstützten Ereignistypen `Office.EventType.ItemChanged` und `Office.EventType.OfficeThemeChanged`.</span><span class="sxs-lookup"><span data-stu-id="99b1d-192">Currently, the supported event types are `Office.EventType.ItemChanged` and `Office.EventType.OfficeThemeChanged`.</span></span>

##### <a name="parameters"></a><span data-ttu-id="99b1d-193">Parameter:</span><span class="sxs-lookup"><span data-stu-id="99b1d-193">Parameters:</span></span>

| <span data-ttu-id="99b1d-194">Name</span><span class="sxs-lookup"><span data-stu-id="99b1d-194">Name</span></span> | <span data-ttu-id="99b1d-195">Typ</span><span class="sxs-lookup"><span data-stu-id="99b1d-195">Type</span></span> | <span data-ttu-id="99b1d-196">Attribute</span><span class="sxs-lookup"><span data-stu-id="99b1d-196">Attributes</span></span> | <span data-ttu-id="99b1d-197">Beschreibung</span><span class="sxs-lookup"><span data-stu-id="99b1d-197">Description</span></span> |
|---|---|---|---|
| `eventType` | [<span data-ttu-id="99b1d-198">Office.EventType</span><span class="sxs-lookup"><span data-stu-id="99b1d-198">Office.EventType</span></span>](office.md#eventtype-string) || <span data-ttu-id="99b1d-199">Das Ereignis, das den Handler aufrufen soll</span><span class="sxs-lookup"><span data-stu-id="99b1d-199">The event that should invoke the handler.</span></span> |
| `handler` | <span data-ttu-id="99b1d-200">Function</span><span class="sxs-lookup"><span data-stu-id="99b1d-200">Function</span></span> || <span data-ttu-id="99b1d-p106">Die Funktion, die das Ereignis behandeln soll. Die Funktion muss einen einzigen Parameter akzeptieren (ein Objektliteral). Die Eigenschaft `type` dieses Parameters entspricht dem Parameter `eventType`, der an `addHandlerAsync` übergeben wird.</span><span class="sxs-lookup"><span data-stu-id="99b1d-p106">The function to handle the event. The function must accept a single parameter, which is an object literal. The `type` property on the parameter will match the `eventType` parameter passed to `addHandlerAsync`.</span></span> |
| `options` | <span data-ttu-id="99b1d-204">Objekt</span><span class="sxs-lookup"><span data-stu-id="99b1d-204">Object</span></span> | <span data-ttu-id="99b1d-205">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="99b1d-205">&lt;optional&gt;</span></span> | <span data-ttu-id="99b1d-206">Ein Objektliteral, das mindestens eine der folgenden Eigenschaften enthält.</span><span class="sxs-lookup"><span data-stu-id="99b1d-206">An object literal that contains one or more of the following properties.</span></span> |
| `options.asyncContext` | <span data-ttu-id="99b1d-207">Object</span><span class="sxs-lookup"><span data-stu-id="99b1d-207">Object</span></span> | <span data-ttu-id="99b1d-208">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="99b1d-208">&lt;optional&gt;</span></span> | <span data-ttu-id="99b1d-209">Entwickler können ein Objekt bereitstellen, auf das sie in der Callbackmethode zugreifen möchten.</span><span class="sxs-lookup"><span data-stu-id="99b1d-209">Developers can provide any object they wish to access in the callback method.</span></span> |
| `callback` | <span data-ttu-id="99b1d-210">Funktion</span><span class="sxs-lookup"><span data-stu-id="99b1d-210">function</span></span>| <span data-ttu-id="99b1d-211">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="99b1d-211">&lt;optional&gt;</span></span>|<span data-ttu-id="99b1d-212">Wenn die Methode abgeschlossen wird, wird die Funktion , die im `callback`-Parameter übergeben wird, mit einem einzelnen Parameter (`asyncResult`) aufgerufen, der ein [`AsyncResult`](/javascript/api/office/office.asyncresult)-Objekt darstellt.</span><span class="sxs-lookup"><span data-stu-id="99b1d-212">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="99b1d-213">Anforderungen</span><span class="sxs-lookup"><span data-stu-id="99b1d-213">Requirements</span></span>

|<span data-ttu-id="99b1d-214">Anforderung</span><span class="sxs-lookup"><span data-stu-id="99b1d-214">Requirement</span></span>| <span data-ttu-id="99b1d-215">Wert</span><span class="sxs-lookup"><span data-stu-id="99b1d-215">Value</span></span>|
|---|---|
|[<span data-ttu-id="99b1d-216">Mindestversion des Postfachanforderungssatzes</span><span class="sxs-lookup"><span data-stu-id="99b1d-216">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="99b1d-217">1,5</span><span class="sxs-lookup"><span data-stu-id="99b1d-217">1.5</span></span> |
|[<span data-ttu-id="99b1d-218">Mindestberechtigungsstufe</span><span class="sxs-lookup"><span data-stu-id="99b1d-218">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="99b1d-219">ReadItem</span><span class="sxs-lookup"><span data-stu-id="99b1d-219">ReadItem</span></span> |
|[<span data-ttu-id="99b1d-220">Zutreffender Outlook-Modus</span><span class="sxs-lookup"><span data-stu-id="99b1d-220">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="99b1d-221">Verfassen oder Lesen</span><span class="sxs-lookup"><span data-stu-id="99b1d-221">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="99b1d-222">Beispiel</span><span class="sxs-lookup"><span data-stu-id="99b1d-222">Example</span></span>

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

####  <a name="converttoewsiditemid-restversion--string"></a><span data-ttu-id="99b1d-223">convertToEwsId(itemId, restVersion) → {String}</span><span class="sxs-lookup"><span data-stu-id="99b1d-223">convertToEwsId(itemId, restVersion) → {String}</span></span>

<span data-ttu-id="99b1d-224">Wandelt eine für REST formatierte Element-ID ins EWS-Format um.</span><span class="sxs-lookup"><span data-stu-id="99b1d-224">Converts an item ID formatted for REST into EWS format.</span></span>

> [!NOTE]
> <span data-ttu-id="99b1d-225">Diese Methode wird in Outlook für iOS oder Outlook für Android nicht unterstützt.</span><span class="sxs-lookup"><span data-stu-id="99b1d-225">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="99b1d-p107">Über eine REST-API abgerufene Element-IDs (z. B. die [Outlook Mail-API](https://docs.microsoft.com/previous-versions/office/office-365-api/api/version-2.0/mail-rest-operations) oder [Microsoft Graph](http://graph.microsoft.io/)) verwenden ein anderes Format, als das von Exchange-Webdiensten (EWS) verwendete. Die `convertToEwsId`-Methode wandelt eine für REST formatierte ID in das entsprechende EWS-Format um.</span><span class="sxs-lookup"><span data-stu-id="99b1d-p107">Item IDs retrieved via a REST API (such as the [Outlook Mail API](https://docs.microsoft.com/previous-versions/office/office-365-api/api/version-2.0/mail-rest-operations) or the [Microsoft Graph](http://graph.microsoft.io/)) use a different format than the format used by Exchange Web Services (EWS). The `convertToEwsId` method converts a REST-formatted ID into the proper format for EWS.</span></span>

##### <a name="parameters"></a><span data-ttu-id="99b1d-228">Parameter:</span><span class="sxs-lookup"><span data-stu-id="99b1d-228">Parameters:</span></span>

|<span data-ttu-id="99b1d-229">Name</span><span class="sxs-lookup"><span data-stu-id="99b1d-229">Name</span></span>| <span data-ttu-id="99b1d-230">Typ</span><span class="sxs-lookup"><span data-stu-id="99b1d-230">Type</span></span>| <span data-ttu-id="99b1d-231">Beschreibung</span><span class="sxs-lookup"><span data-stu-id="99b1d-231">Description</span></span>|
|---|---|---|
|`itemId`| <span data-ttu-id="99b1d-232">Zeichenfolge</span><span class="sxs-lookup"><span data-stu-id="99b1d-232">String</span></span>|<span data-ttu-id="99b1d-233">Eine für Outlook-REST-APIs formatierte Element-ID</span><span class="sxs-lookup"><span data-stu-id="99b1d-233">An item ID formatted for the Outlook REST APIs</span></span>|
|`restVersion`| [<span data-ttu-id="99b1d-234">Office.MailboxEnums.RestVersion</span><span class="sxs-lookup"><span data-stu-id="99b1d-234">Office.MailboxEnums.RestVersion</span></span>](/javascript/api/outlook_1_7/office.mailboxenums.restversion)|<span data-ttu-id="99b1d-235">Ein Wert, der die Version der Outlook-REST-API angibt, die zum Abrufen der Element-ID verwendet wurde.</span><span class="sxs-lookup"><span data-stu-id="99b1d-235">A value indicating the version of the Outlook REST API used to retrieve the item ID.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="99b1d-236">Anforderungen</span><span class="sxs-lookup"><span data-stu-id="99b1d-236">Requirements</span></span>

|<span data-ttu-id="99b1d-237">Anforderung</span><span class="sxs-lookup"><span data-stu-id="99b1d-237">Requirement</span></span>| <span data-ttu-id="99b1d-238">Wert</span><span class="sxs-lookup"><span data-stu-id="99b1d-238">Value</span></span>|
|---|---|
|[<span data-ttu-id="99b1d-239">Mindestversion des Postfachanforderungssatzes</span><span class="sxs-lookup"><span data-stu-id="99b1d-239">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="99b1d-240">1.3</span><span class="sxs-lookup"><span data-stu-id="99b1d-240">1.3</span></span>|
|[<span data-ttu-id="99b1d-241">Mindestberechtigungsstufe</span><span class="sxs-lookup"><span data-stu-id="99b1d-241">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="99b1d-242">Eingeschränkt</span><span class="sxs-lookup"><span data-stu-id="99b1d-242">Restricted</span></span>|
|[<span data-ttu-id="99b1d-243">Zutreffender Outlook-Modus</span><span class="sxs-lookup"><span data-stu-id="99b1d-243">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="99b1d-244">Verfassen oder Lesen</span><span class="sxs-lookup"><span data-stu-id="99b1d-244">Compose or read</span></span>|

##### <a name="returns"></a><span data-ttu-id="99b1d-245">Gibt zurück:</span><span class="sxs-lookup"><span data-stu-id="99b1d-245">Returns:</span></span>

<span data-ttu-id="99b1d-246">Typ: Zeichenfolge</span><span class="sxs-lookup"><span data-stu-id="99b1d-246">Type: String</span></span>

##### <a name="example"></a><span data-ttu-id="99b1d-247">Beispiel</span><span class="sxs-lookup"><span data-stu-id="99b1d-247">Example</span></span>

```
// Get an item's ID from a REST API
var restId = 'AAMkAGVlOTZjNTM3LW...';

// Treat restId as coming from the v2.0 version of the
// Outlook Mail API
var ewsId = Office.context.mailbox.convertToEwsId(restId, Office.MailboxEnums.RestVersion.v2_0);
```

####  <a name="converttolocalclienttimetimevalue--localclienttimejavascriptapioutlook17officelocalclienttime"></a><span data-ttu-id="99b1d-248">convertToLocalClientTime(timeValue) → {[LocalClientTime](/javascript/api/outlook_1_7/office.LocalClientTime)}</span><span class="sxs-lookup"><span data-stu-id="99b1d-248">convertToLocalClientTime(timeValue) → {[LocalClientTime](/javascript/api/outlook_1_7/office.LocalClientTime)}</span></span>

<span data-ttu-id="99b1d-249">Ruft ein Wörterbuch mit Uhrzeitinformationen basierend auf der Zeiteinstellung des lokalen Clients ab.</span><span class="sxs-lookup"><span data-stu-id="99b1d-249">Gets a dictionary containing time information in local client time.</span></span>

<span data-ttu-id="99b1d-p108">Die in einer Mail-App für Outlook oder Outlook Web App verwendeten Daten und Uhrzeiten stammen u. U. aus verschiedenen Zeitzonen. Outlook verwendet die Zeitzone des Client-Computers; Outlook Web App verwendet die im Exchange Admin Center (EAC) festgelegte Zeitzone. Sie sollten Datums- und Uhrzeitwerte bearbeiten, damit die auf der Benutzeroberfläche angezeigten Werte immer den von Benutzer erwarteten Zeitzonen entsprechen.</span><span class="sxs-lookup"><span data-stu-id="99b1d-p108">The dates and times used by a mail app for Outlook or Outlook Web App can use different time zones. Outlook uses the client computer time zone; Outlook Web App uses the time zone set on the Exchange Admin Center (EAC). You should handle date and time values so that the values you display on the user interface are always consistent with the time zone that the user expects.</span></span>

<span data-ttu-id="99b1d-p109">Wird die Mail-App in Outlook ausgeführt, wird mit der `convertToLocalClientTime`-Methode ein Wörterbuchobjekt zurückgegeben, dessen Werte auf die Zeitzone des Clientcomputers festgelegt sind. Wird die Mail-App in Outlook ausgeführt, wird mit der `convertToLocalClientTime`-Methode ein Wörterbuchobjekt zurückgegeben, dessen Werte auf die Zeitzone des Clientcomputers festgelegt sind.</span><span class="sxs-lookup"><span data-stu-id="99b1d-p109">If the mail app is running in Outlook, the `convertToLocalClientTime` method will return a dictionary object with the values set to the client computer time zone. If the mail app is running in Outlook Web App, the `convertToLocalClientTime` method will return a dictionary object with the values set to the time zone specified in the EAC.</span></span>

##### <a name="parameters"></a><span data-ttu-id="99b1d-255">Parameter:</span><span class="sxs-lookup"><span data-stu-id="99b1d-255">Parameters:</span></span>

|<span data-ttu-id="99b1d-256">Name</span><span class="sxs-lookup"><span data-stu-id="99b1d-256">Name</span></span>| <span data-ttu-id="99b1d-257">Typ</span><span class="sxs-lookup"><span data-stu-id="99b1d-257">Type</span></span>| <span data-ttu-id="99b1d-258">Beschreibung</span><span class="sxs-lookup"><span data-stu-id="99b1d-258">Description</span></span>|
|---|---|---|
|`timeValue`| <span data-ttu-id="99b1d-259">Datum</span><span class="sxs-lookup"><span data-stu-id="99b1d-259">Date</span></span>|<span data-ttu-id="99b1d-260">Ein Date-Objekt</span><span class="sxs-lookup"><span data-stu-id="99b1d-260">A Date object</span></span>|

##### <a name="requirements"></a><span data-ttu-id="99b1d-261">Anforderungen</span><span class="sxs-lookup"><span data-stu-id="99b1d-261">Requirements</span></span>

|<span data-ttu-id="99b1d-262">Anforderung</span><span class="sxs-lookup"><span data-stu-id="99b1d-262">Requirement</span></span>| <span data-ttu-id="99b1d-263">Wert</span><span class="sxs-lookup"><span data-stu-id="99b1d-263">Value</span></span>|
|---|---|
|[<span data-ttu-id="99b1d-264">Mindestversion des Postfachanforderungssatzes</span><span class="sxs-lookup"><span data-stu-id="99b1d-264">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="99b1d-265">1.0</span><span class="sxs-lookup"><span data-stu-id="99b1d-265">1.0</span></span>|
|[<span data-ttu-id="99b1d-266">Mindestberechtigungsstufe</span><span class="sxs-lookup"><span data-stu-id="99b1d-266">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="99b1d-267">ReadItem</span><span class="sxs-lookup"><span data-stu-id="99b1d-267">ReadItem</span></span>|
|[<span data-ttu-id="99b1d-268">Zutreffender Outlook-Modus</span><span class="sxs-lookup"><span data-stu-id="99b1d-268">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="99b1d-269">Verfassen oder Lesen</span><span class="sxs-lookup"><span data-stu-id="99b1d-269">Compose or read</span></span>|

##### <a name="returns"></a><span data-ttu-id="99b1d-270">Gibt zurück:</span><span class="sxs-lookup"><span data-stu-id="99b1d-270">Returns:</span></span>

<span data-ttu-id="99b1d-271">Typ: [LocalClientTime](/javascript/api/outlook_1_7/office.LocalClientTime)</span><span class="sxs-lookup"><span data-stu-id="99b1d-271">Type: [LocalClientTime](/javascript/api/outlook_1_7/office.LocalClientTime)</span></span>

####  <a name="converttorestiditemid-restversion--string"></a><span data-ttu-id="99b1d-272">convertToRestId(itemId, restVersion) → {String}</span><span class="sxs-lookup"><span data-stu-id="99b1d-272">convertToRestId(itemId, restVersion) → {String}</span></span>

<span data-ttu-id="99b1d-273">Wandelt eine für EWS formatierte Element-ID ins REST-Format um.</span><span class="sxs-lookup"><span data-stu-id="99b1d-273">Converts an item ID formatted for EWS into REST format.</span></span>

> [!NOTE]
> <span data-ttu-id="99b1d-274">Diese Methode wird in Outlook für iOS oder Outlook für Android nicht unterstützt.</span><span class="sxs-lookup"><span data-stu-id="99b1d-274">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="99b1d-p110">Über EWS oder die `itemId` abgerufene Element-IDs verwenden ein anderes Format als die über REST-APIs (z. B. die [Outlook Mail-API](https://docs.microsoft.com/previous-versions/office/office-365-api/api/version-2.0/mail-rest-operations) oder [Microsoft Graph](http://graph.microsoft.io/)) abgerufenen Element-IDs. Die `convertToRestId`-Methode wandelt eine für EWS formatierte ID in das entsprechende REST-Format um.</span><span class="sxs-lookup"><span data-stu-id="99b1d-p110">Item IDs retrieved via EWS or via the `itemId` property use a different format than the format used by REST APIs (such as the [Outlook Mail API](https://docs.microsoft.com/previous-versions/office/office-365-api/api/version-2.0/mail-rest-operations) or the [Microsoft Graph](http://graph.microsoft.io/)). The `convertToRestId` method converts an EWS-formatted ID into the proper format for REST.</span></span>

##### <a name="parameters"></a><span data-ttu-id="99b1d-277">Parameter:</span><span class="sxs-lookup"><span data-stu-id="99b1d-277">Parameters:</span></span>

|<span data-ttu-id="99b1d-278">Name</span><span class="sxs-lookup"><span data-stu-id="99b1d-278">Name</span></span>| <span data-ttu-id="99b1d-279">Typ</span><span class="sxs-lookup"><span data-stu-id="99b1d-279">Type</span></span>| <span data-ttu-id="99b1d-280">Beschreibung</span><span class="sxs-lookup"><span data-stu-id="99b1d-280">Description</span></span>|
|---|---|---|
|`itemId`| <span data-ttu-id="99b1d-281">Zeichenfolge</span><span class="sxs-lookup"><span data-stu-id="99b1d-281">String</span></span>|<span data-ttu-id="99b1d-282">Eine für Exchange-Webdienste (EWS) formatierte Element-ID</span><span class="sxs-lookup"><span data-stu-id="99b1d-282">An item ID formatted for Exchange Web Services (EWS)</span></span>|
|`restVersion`| [<span data-ttu-id="99b1d-283">Office.MailboxEnums.RestVersion</span><span class="sxs-lookup"><span data-stu-id="99b1d-283">Office.MailboxEnums.RestVersion</span></span>](/javascript/api/outlook_1_7/office.mailboxenums.restversion)|<span data-ttu-id="99b1d-284">Ein Wert, der die Version der Outlook-REST-API angibt, mit der die konvertierte ID verwendet wird.</span><span class="sxs-lookup"><span data-stu-id="99b1d-284">A value indicating the version of the Outlook REST API that the converted ID will be used with.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="99b1d-285">Anforderungen</span><span class="sxs-lookup"><span data-stu-id="99b1d-285">Requirements</span></span>

|<span data-ttu-id="99b1d-286">Anforderung</span><span class="sxs-lookup"><span data-stu-id="99b1d-286">Requirement</span></span>| <span data-ttu-id="99b1d-287">Wert</span><span class="sxs-lookup"><span data-stu-id="99b1d-287">Value</span></span>|
|---|---|
|[<span data-ttu-id="99b1d-288">Mindestversion des Postfachanforderungssatzes</span><span class="sxs-lookup"><span data-stu-id="99b1d-288">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="99b1d-289">1.3</span><span class="sxs-lookup"><span data-stu-id="99b1d-289">1.3</span></span>|
|[<span data-ttu-id="99b1d-290">Mindestberechtigungsstufe</span><span class="sxs-lookup"><span data-stu-id="99b1d-290">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="99b1d-291">Eingeschränkt</span><span class="sxs-lookup"><span data-stu-id="99b1d-291">Restricted</span></span>|
|[<span data-ttu-id="99b1d-292">Zutreffender Outlook-Modus</span><span class="sxs-lookup"><span data-stu-id="99b1d-292">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="99b1d-293">Verfassen oder Lesen</span><span class="sxs-lookup"><span data-stu-id="99b1d-293">Compose or read</span></span>|

##### <a name="returns"></a><span data-ttu-id="99b1d-294">Gibt zurück:</span><span class="sxs-lookup"><span data-stu-id="99b1d-294">Returns:</span></span>

<span data-ttu-id="99b1d-295">Typ: Zeichenfolge</span><span class="sxs-lookup"><span data-stu-id="99b1d-295">Type: String</span></span>

##### <a name="example"></a><span data-ttu-id="99b1d-296">Beispiel</span><span class="sxs-lookup"><span data-stu-id="99b1d-296">Example</span></span>

```
// Get the currently selected item's ID
var ewsId = Office.context.mailbox.item.itemId;

// Convert to a REST ID for the v2.0 version of the
// Outlook Mail API
var restId = Office.context.mailbox.convertToRestId(ewsId, Office.MailboxEnums.RestVersion.v2_0);
```

####  <a name="converttoutcclienttimeinput--date"></a><span data-ttu-id="99b1d-297">convertToUtcClientTime(input) → {Date}</span><span class="sxs-lookup"><span data-stu-id="99b1d-297">convertToUtcClientTime(input) → {Date}</span></span>

<span data-ttu-id="99b1d-298">Ruft ein Date-Objekt aus einem Wörterbuch mit Uhrzeitinformationen ab.</span><span class="sxs-lookup"><span data-stu-id="99b1d-298">Gets a Date object from a dictionary containing time information.</span></span>

<span data-ttu-id="99b1d-299">Mit der `convertToUtcClientTime`-Methode wird ein Wörterbuch mit lokalem Datum und lokaler Uhrzeit in ein Date-Objekt mit den richtigen Werten für das lokale Datum und die lokale Uhrzeit konvertiert.</span><span class="sxs-lookup"><span data-stu-id="99b1d-299">The `convertToUtcClientTime` method converts a dictionary containing a local date and time to a Date object with the correct values for the local date and time.</span></span>

##### <a name="parameters"></a><span data-ttu-id="99b1d-300">Parameter:</span><span class="sxs-lookup"><span data-stu-id="99b1d-300">Parameters:</span></span>

|<span data-ttu-id="99b1d-301">Name</span><span class="sxs-lookup"><span data-stu-id="99b1d-301">Name</span></span>| <span data-ttu-id="99b1d-302">Typ</span><span class="sxs-lookup"><span data-stu-id="99b1d-302">Type</span></span>| <span data-ttu-id="99b1d-303">Beschreibung</span><span class="sxs-lookup"><span data-stu-id="99b1d-303">Description</span></span>|
|---|---|---|
|`input`| [<span data-ttu-id="99b1d-304">LocalClientTime</span><span class="sxs-lookup"><span data-stu-id="99b1d-304">LocalClientTime</span></span>](/javascript/api/outlook_1_7/office.LocalClientTime)|<span data-ttu-id="99b1d-305">Der zu konvertierende Wert für die lokale Uhrzeit.</span><span class="sxs-lookup"><span data-stu-id="99b1d-305">The local time value to convert.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="99b1d-306">Anforderungen</span><span class="sxs-lookup"><span data-stu-id="99b1d-306">Requirements</span></span>

|<span data-ttu-id="99b1d-307">Anforderung</span><span class="sxs-lookup"><span data-stu-id="99b1d-307">Requirement</span></span>| <span data-ttu-id="99b1d-308">Wert</span><span class="sxs-lookup"><span data-stu-id="99b1d-308">Value</span></span>|
|---|---|
|[<span data-ttu-id="99b1d-309">Mindestversion des Postfachanforderungssatzes</span><span class="sxs-lookup"><span data-stu-id="99b1d-309">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="99b1d-310">1.0</span><span class="sxs-lookup"><span data-stu-id="99b1d-310">1.0</span></span>|
|[<span data-ttu-id="99b1d-311">Mindestberechtigungsstufe</span><span class="sxs-lookup"><span data-stu-id="99b1d-311">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="99b1d-312">ReadItem</span><span class="sxs-lookup"><span data-stu-id="99b1d-312">ReadItem</span></span>|
|[<span data-ttu-id="99b1d-313">Zutreffender Outlook-Modus</span><span class="sxs-lookup"><span data-stu-id="99b1d-313">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="99b1d-314">Verfassen oder Lesen</span><span class="sxs-lookup"><span data-stu-id="99b1d-314">Compose or read</span></span>|

##### <a name="returns"></a><span data-ttu-id="99b1d-315">Gibt zurück:</span><span class="sxs-lookup"><span data-stu-id="99b1d-315">Returns:</span></span>

<span data-ttu-id="99b1d-316">Ein Date-Objekt der Uhrzeit in UTC.</span><span class="sxs-lookup"><span data-stu-id="99b1d-316">A Date object with the time expressed in UTC.</span></span>

<dl class="param-type"><span data-ttu-id="99b1d-317">

<dt>
Typ</dt>


</span><span class="sxs-lookup"><span data-stu-id="99b1d-317">

<dt>Type</dt>

</span></span><dd><span data-ttu-id="99b1d-318">Datum</span><span class="sxs-lookup"><span data-stu-id="99b1d-318">Date</span></span></dd>

</dl>

####  <a name="displayappointmentformitemid"></a><span data-ttu-id="99b1d-319">displayAppointmentForm(itemId)</span><span class="sxs-lookup"><span data-stu-id="99b1d-319">displayAppointmentForm(itemId)</span></span>

<span data-ttu-id="99b1d-320">Zeigt einen bestehenden Kalendertermin an.</span><span class="sxs-lookup"><span data-stu-id="99b1d-320">Displays an existing calendar appointment.</span></span>

> [!NOTE]
> <span data-ttu-id="99b1d-321">Diese Methode wird in Outlook für iOS oder Outlook für Android nicht unterstützt.</span><span class="sxs-lookup"><span data-stu-id="99b1d-321">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="99b1d-322">Mit der `displayAppointmentForm`-Methode wird ein vorhandener Kalendertermin auf dem Desktop in einem neuen Fenster oder auf Mobilgeräten in einem Dialogfeld geöffnet.</span><span class="sxs-lookup"><span data-stu-id="99b1d-322">The `displayAppointmentForm` method opens an existing calendar appointment in a new window on the desktop or in a dialog box on mobile devices.</span></span>

<span data-ttu-id="99b1d-p111">In Outlook for Mac können Sie diese Methode verwenden, um einen einzelnen Termin anzuzeigen, der nicht Teil einer Terminserie ist, oder den Mastertermin einer Terminserie, jedoch keine Instanz der Terminserie. Dies liegt daran, dass Sie in Outlook for Mac nicht auf die Eigenschaften (einschließlich der Element-ID) von Instanzen einer Terminserie zugreifen können.</span><span class="sxs-lookup"><span data-stu-id="99b1d-p111">In Outlook for Mac, you can use this method to display a single appointment that is not part of a recurring series, or the master appointment of a recurring series, but you cannot display an instance of the series. This is because in Outlook for Mac, you cannot access the properties (including the item ID) of instances of a recurring series.</span></span>

<span data-ttu-id="99b1d-325">In Outlook Web App öffnet diese Methode das angegebene Formular nur, wenn der Textkörper des Formulars nicht größer ist als 32 KB.</span><span class="sxs-lookup"><span data-stu-id="99b1d-325">In Outlook Web App, this method opens the specified form only if the body of the form is less than or equal to 32KB number of characters.</span></span>

<span data-ttu-id="99b1d-326">Wenn der angegebene Elementbezeichner keinen vorhandenen Termin identifiziert, wird auf dem Clientcomputer oder Gerät ein leerer Bereich geöffnet, und es wird keine Fehlermeldung zurückgegeben.</span><span class="sxs-lookup"><span data-stu-id="99b1d-326">If the specified item identifier does not identify an existing appointment, a blank pane opens on the client computer or device, and no error message will be returned.</span></span>

##### <a name="parameters"></a><span data-ttu-id="99b1d-327">Parameter:</span><span class="sxs-lookup"><span data-stu-id="99b1d-327">Parameters:</span></span>

|<span data-ttu-id="99b1d-328">Name</span><span class="sxs-lookup"><span data-stu-id="99b1d-328">Name</span></span>| <span data-ttu-id="99b1d-329">Typ</span><span class="sxs-lookup"><span data-stu-id="99b1d-329">Type</span></span>| <span data-ttu-id="99b1d-330">Beschreibung</span><span class="sxs-lookup"><span data-stu-id="99b1d-330">Description</span></span>|
|---|---|---|
|`itemId`| <span data-ttu-id="99b1d-331">Zeichenfolge</span><span class="sxs-lookup"><span data-stu-id="99b1d-331">String</span></span>|<span data-ttu-id="99b1d-332">Der EWS-Bezeichner (Exchange Web Services, Exchange-Webdienste) für einen vorhandenen Kalendertermin.</span><span class="sxs-lookup"><span data-stu-id="99b1d-332">The Exchange Web Services (EWS) identifier for an existing calendar appointment.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="99b1d-333">Anforderungen</span><span class="sxs-lookup"><span data-stu-id="99b1d-333">Requirements</span></span>

|<span data-ttu-id="99b1d-334">Anforderung</span><span class="sxs-lookup"><span data-stu-id="99b1d-334">Requirement</span></span>| <span data-ttu-id="99b1d-335">Wert</span><span class="sxs-lookup"><span data-stu-id="99b1d-335">Value</span></span>|
|---|---|
|[<span data-ttu-id="99b1d-336">Mindestversion des Postfachanforderungssatzes</span><span class="sxs-lookup"><span data-stu-id="99b1d-336">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="99b1d-337">1.0</span><span class="sxs-lookup"><span data-stu-id="99b1d-337">1.0</span></span>|
|[<span data-ttu-id="99b1d-338">Mindestberechtigungsstufe</span><span class="sxs-lookup"><span data-stu-id="99b1d-338">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="99b1d-339">ReadItem</span><span class="sxs-lookup"><span data-stu-id="99b1d-339">ReadItem</span></span>|
|[<span data-ttu-id="99b1d-340">Zutreffender Outlook-Modus</span><span class="sxs-lookup"><span data-stu-id="99b1d-340">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="99b1d-341">Verfassen oder Lesen</span><span class="sxs-lookup"><span data-stu-id="99b1d-341">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="99b1d-342">Beispiel</span><span class="sxs-lookup"><span data-stu-id="99b1d-342">Example</span></span>

```
Office.context.mailbox.displayAppointmentForm(appointmentId);
```

####  <a name="displaymessageformitemid"></a><span data-ttu-id="99b1d-343">displayMessageForm(itemId)</span><span class="sxs-lookup"><span data-stu-id="99b1d-343">displayMessageForm(itemId)</span></span>

<span data-ttu-id="99b1d-344">Zeigt eine vorhandene Nachricht an.</span><span class="sxs-lookup"><span data-stu-id="99b1d-344">Displays an existing message.</span></span>

> [!NOTE]
> <span data-ttu-id="99b1d-345">Diese Methode wird in Outlook für iOS oder Outlook für Android nicht unterstützt.</span><span class="sxs-lookup"><span data-stu-id="99b1d-345">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="99b1d-346">Die `displayMessageForm`-Methode öffnet eine vorhandene Nachricht in einem neuen Fenster auf dem Desktop bzw. in einem Dialogfeld auf Mobilgeräten.</span><span class="sxs-lookup"><span data-stu-id="99b1d-346">The `displayMessageForm` method opens an existing message in a new window on the desktop or in a dialog box on mobile devices.</span></span>

<span data-ttu-id="99b1d-347">In Outlook Web App wird mit dieser Methode das angegebene Formular nur dann geöffnet, wenn der Textkörper des Formular Zeichen im Umfang vom maximal 32 KB umfasst.</span><span class="sxs-lookup"><span data-stu-id="99b1d-347">In Outlook Web App, this method opens the specified form only if the body of the form is less than or equal to 32 KB number of characters.</span></span>

<span data-ttu-id="99b1d-348">Wenn der angegebene Elementbezeichner keine vorhandenen Nachrichten erkennt, wird auf dem Client-Computer keine Nachricht angezeigt, und es werden keine Fehlermeldungen zurückgegeben.</span><span class="sxs-lookup"><span data-stu-id="99b1d-348">If the specified item identifier does not identify an existing message, no message will be displayed on the client computer, and no error message will be returned.</span></span>

<span data-ttu-id="99b1d-p112">Verwenden Sie `displayMessageForm` nicht mit einem `itemId`-Objekt, das einen Termin darstellt. Verwenden Sie die `displayAppointmentForm`-Methode, um einen vorhandenen Termin anzuzeigen, und `displayNewAppointmentForm`, um ein Formular zum Erstellen eines neuen Termins anzuzeigen.</span><span class="sxs-lookup"><span data-stu-id="99b1d-p112">Do not use the `displayMessageForm` with an `itemId` that represents an appointment. Use the `displayAppointmentForm` method to display an existing appointment, and `displayNewAppointmentForm` to display a form to create a new appointment.</span></span>

##### <a name="parameters"></a><span data-ttu-id="99b1d-351">Parameter:</span><span class="sxs-lookup"><span data-stu-id="99b1d-351">Parameters:</span></span>

|<span data-ttu-id="99b1d-352">Name</span><span class="sxs-lookup"><span data-stu-id="99b1d-352">Name</span></span>| <span data-ttu-id="99b1d-353">Typ</span><span class="sxs-lookup"><span data-stu-id="99b1d-353">Type</span></span>| <span data-ttu-id="99b1d-354">Beschreibung</span><span class="sxs-lookup"><span data-stu-id="99b1d-354">Description</span></span>|
|---|---|---|
|`itemId`| <span data-ttu-id="99b1d-355">Zeichenfolge</span><span class="sxs-lookup"><span data-stu-id="99b1d-355">String</span></span>|<span data-ttu-id="99b1d-356">Der Exchange-Webdienste (EWS) für eine vorhandene Nachricht.</span><span class="sxs-lookup"><span data-stu-id="99b1d-356">The Exchange Web Services (EWS) identifier for an existing message.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="99b1d-357">Anforderungen</span><span class="sxs-lookup"><span data-stu-id="99b1d-357">Requirements</span></span>

|<span data-ttu-id="99b1d-358">Anforderung</span><span class="sxs-lookup"><span data-stu-id="99b1d-358">Requirement</span></span>| <span data-ttu-id="99b1d-359">Wert</span><span class="sxs-lookup"><span data-stu-id="99b1d-359">Value</span></span>|
|---|---|
|[<span data-ttu-id="99b1d-360">Mindestversion des Postfachanforderungssatzes</span><span class="sxs-lookup"><span data-stu-id="99b1d-360">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="99b1d-361">1.0</span><span class="sxs-lookup"><span data-stu-id="99b1d-361">1.0</span></span>|
|[<span data-ttu-id="99b1d-362">Mindestberechtigungsstufe</span><span class="sxs-lookup"><span data-stu-id="99b1d-362">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="99b1d-363">ReadItem</span><span class="sxs-lookup"><span data-stu-id="99b1d-363">ReadItem</span></span>|
|[<span data-ttu-id="99b1d-364">Zutreffender Outlook-Modus</span><span class="sxs-lookup"><span data-stu-id="99b1d-364">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="99b1d-365">Verfassen oder Lesen</span><span class="sxs-lookup"><span data-stu-id="99b1d-365">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="99b1d-366">Beispiel</span><span class="sxs-lookup"><span data-stu-id="99b1d-366">Example</span></span>

```
Office.context.mailbox.displayMessageForm(messageId);
```

#### <a name="displaynewappointmentformparameters"></a><span data-ttu-id="99b1d-367">displayNewAppointmentForm(parameters)</span><span class="sxs-lookup"><span data-stu-id="99b1d-367">displayNewAppointmentForm(parameters)</span></span>

<span data-ttu-id="99b1d-368">Zeigt ein Formular zum Erstellen eines neuen Kalendertermins an.</span><span class="sxs-lookup"><span data-stu-id="99b1d-368">Displays a form for creating a new calendar appointment.</span></span>

> [!NOTE]
> <span data-ttu-id="99b1d-369">Diese Methode wird in Outlook für iOS oder Outlook für Android nicht unterstützt.</span><span class="sxs-lookup"><span data-stu-id="99b1d-369">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="99b1d-p113">Mit der `displayNewAppointmentForm`-Methode wird ein Formular geöffnet, mit dem der Benutzer einen neuen Termin oder eine Besprechung erstellen kann. Wenn Parameter angegeben wurden, werden die Felder im Terminformular automatisch mit dem Inhalt der Parameter ausgefüllt.</span><span class="sxs-lookup"><span data-stu-id="99b1d-p113">The `displayNewAppointmentForm` method opens a form that enables the user to create a new appointment or meeting. If parameters are specified, the appointment form fields are automatically populated with the contents of the parameters.</span></span>

<span data-ttu-id="99b1d-p114">In Outlook Web App und OWA for Devices zeigt diese Methode immer ein Formular mit einem Teilnehmerfeld an. Wenn Sie keine Teilnehmer als Eingabeargumente angeben, zeigt die Methode ein Formular mit einer Schaltfläche **Speichern** an. Wenn Sie Teilnehmer angegeben haben, enthält das Formular die Teilnehmer und eine Schaltfläche **Senden**.</span><span class="sxs-lookup"><span data-stu-id="99b1d-p114">In Outlook Web App and OWA for Devices, this method always displays a form with an attendees field. If you do not specify any attendees as input arguments, the method displays a form with a **Save** button. If you have specified attendees, the form would include the attendees and a **Send** button.</span></span>

<span data-ttu-id="99b1d-p115">Wenn Teilnehmer oder Ressourcen im `requiredAttendees`-, `optionalAttendees`- oder `resources`-Parameter festgelegt werden, wird mit dieser Methode im Outlook Rich Client und Outlook RT ein Besprechungsformular mit der Schaltfläche **Senden** angezeigt. Wenn keine Teilnehmer angegeben werden, wird mit dieser Methode ein Terminformular mit der Schaltfläche **Speichern & schließen** angezeigt.</span><span class="sxs-lookup"><span data-stu-id="99b1d-p115">In the Outlook rich client and Outlook RT, if you specify any attendees or resources in the `requiredAttendees`, `optionalAttendees`, or `resources` parameter, this method displays a meeting form with a **Send** button. If you don't specify any recipients, this method displays an appointment form with a **Save & Close** button.</span></span>

<span data-ttu-id="99b1d-377">Wenn einer der Parameter die angegebenen Größenbeschränkungen überschreitet oder wenn ein unbekannter Parametername angegeben wird, wird eine Ausnahme ausgelöst.</span><span class="sxs-lookup"><span data-stu-id="99b1d-377">If any of the parameters exceed the specified size limits, or if an unknown parameter name is specified, an exception is thrown.</span></span>

##### <a name="parameters"></a><span data-ttu-id="99b1d-378">Parameter:</span><span class="sxs-lookup"><span data-stu-id="99b1d-378">Parameters:</span></span>

> [!NOTE]
> <span data-ttu-id="99b1d-379">Alle Parameter sind optional.</span><span class="sxs-lookup"><span data-stu-id="99b1d-379">All parameters are optional.</span></span>

|<span data-ttu-id="99b1d-380">Name</span><span class="sxs-lookup"><span data-stu-id="99b1d-380">Name</span></span>| <span data-ttu-id="99b1d-381">Typ</span><span class="sxs-lookup"><span data-stu-id="99b1d-381">Type</span></span>| <span data-ttu-id="99b1d-382">Beschreibung</span><span class="sxs-lookup"><span data-stu-id="99b1d-382">Description</span></span>|
|---|---|---|
| `parameters` | <span data-ttu-id="99b1d-383">Object</span><span class="sxs-lookup"><span data-stu-id="99b1d-383">Object</span></span> | <span data-ttu-id="99b1d-384">Ein Wörterbuch mit Parametern, die den neuen Termin beschreiben</span><span class="sxs-lookup"><span data-stu-id="99b1d-384">A dictionary of parameters describing the new appointment.</span></span> |
| `parameters.requiredAttendees` | <span data-ttu-id="99b1d-385">Array.&lt;String&gt; &#124; Array.&lt;[EmailAddressDetails](/javascript/api/outlook_1_7/office.emailaddressdetails)&gt;</span><span class="sxs-lookup"><span data-stu-id="99b1d-385">Array.&lt;String&gt; &#124; Array.&lt;[EmailAddressDetails](/javascript/api/outlook_1_7/office.emailaddressdetails)&gt;</span></span> | <span data-ttu-id="99b1d-p116">Ein Array aus Zeichenfolgen mit den E-Mail-Adressen oder ein Array mit einem `EmailAddressDetails`-Objekt für jeden der erforderlichen Teilnehmer für den Termin. Das Array darf maximal 100 Einträge enthalten.</span><span class="sxs-lookup"><span data-stu-id="99b1d-p116">An array of strings containing the email addresses or an array containing an `EmailAddressDetails` object for each of the required attendees for the appointment. The array is limited to a maximum of 100 entries.</span></span> |
| `parameters.optionalAttendees` | <span data-ttu-id="99b1d-388">Array.&lt;String&gt; &#124; Array.&lt;[EmailAddressDetails](/javascript/api/outlook_1_7/office.emailaddressdetails)&gt;</span><span class="sxs-lookup"><span data-stu-id="99b1d-388">Array.&lt;String&gt; &#124; Array.&lt;[EmailAddressDetails](/javascript/api/outlook_1_7/office.emailaddressdetails)&gt;</span></span> | <span data-ttu-id="99b1d-p117">Ein Array aus Zeichenfolgen mit den E-Mail-Adressen oder ein Array mit einem Objekt des Typs `EmailAddressDetails` für jeden der optionalen Teilnehmer des Termins. Das Array darf maximal 100 Einträge enthalten.</span><span class="sxs-lookup"><span data-stu-id="99b1d-p117">An array of strings containing the email addresses or an array containing an `EmailAddressDetails` object for each of the optional attendees for the appointment. The array is limited to a maximum of 100 entries.</span></span> |
| `parameters.start` | <span data-ttu-id="99b1d-391">Datum</span><span class="sxs-lookup"><span data-stu-id="99b1d-391">Date</span></span> | <span data-ttu-id="99b1d-392">Ein Objekt des Typs `Date`. Gibt das Startdatum und den Beginn des Termins an.</span><span class="sxs-lookup"><span data-stu-id="99b1d-392">A `Date` object specifying the start date and time of the appointment.</span></span> |
| `parameters.end` | <span data-ttu-id="99b1d-393">Date</span><span class="sxs-lookup"><span data-stu-id="99b1d-393">Date</span></span> | <span data-ttu-id="99b1d-394">Ein Objekt des Typs `Date`. Gibt das Enddatum und das Ende des Termins an.</span><span class="sxs-lookup"><span data-stu-id="99b1d-394">A `Date` object specifying the end date and time of the appointment.</span></span> |
| `parameters.location` | <span data-ttu-id="99b1d-395">String</span><span class="sxs-lookup"><span data-stu-id="99b1d-395">String</span></span> | <span data-ttu-id="99b1d-p118">Eine Zeichenfolge mit dem Ort für den Termin. Die Zeichenfolge ist auf maximal 255 Zeichen beschränkt.</span><span class="sxs-lookup"><span data-stu-id="99b1d-p118">A string containing the location of the appointment. The string is limited to a maximum of 255 characters.</span></span> |
| `parameters.resources` | <span data-ttu-id="99b1d-398">Array.&lt;String&gt;</span><span class="sxs-lookup"><span data-stu-id="99b1d-398">Array.&lt;String&gt;</span></span> | <span data-ttu-id="99b1d-p119">Ein Array aus Zeichenfolgen, die die für den Termin erforderlichen Ressourcen enthalten. Das Array darf maximal 100 Einträge enthalten.</span><span class="sxs-lookup"><span data-stu-id="99b1d-p119">An array of strings containing the resources required for the appointment. The array is limited to a maximum of 100 entries.</span></span> |
| `parameters.subject` | <span data-ttu-id="99b1d-401">Zeichenfolge</span><span class="sxs-lookup"><span data-stu-id="99b1d-401">String</span></span> | <span data-ttu-id="99b1d-p120">Eine Zeichenfolge mit dem Betreff für den Termin. Die Zeichenfolge ist auf maximal 255 Zeichen beschränkt.</span><span class="sxs-lookup"><span data-stu-id="99b1d-p120">A string containing the subject of the appointment. The string is limited to a maximum of 255 characters.</span></span> |
| `parameters.body` | <span data-ttu-id="99b1d-404">Zeichenfolge</span><span class="sxs-lookup"><span data-stu-id="99b1d-404">String</span></span> | <span data-ttu-id="99b1d-p121">Der Text des Termins. Der Textinhalt darf maximal 32 KB umfassen.</span><span class="sxs-lookup"><span data-stu-id="99b1d-p121">The body of the appointment. The body content is limited to a maximum size of 32 KB.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="99b1d-407">Anforderungen</span><span class="sxs-lookup"><span data-stu-id="99b1d-407">Requirements</span></span>

|<span data-ttu-id="99b1d-408">Anforderung</span><span class="sxs-lookup"><span data-stu-id="99b1d-408">Requirement</span></span>| <span data-ttu-id="99b1d-409">Wert</span><span class="sxs-lookup"><span data-stu-id="99b1d-409">Value</span></span>|
|---|---|
|[<span data-ttu-id="99b1d-410">Mindestversion des Postfachanforderungssatzes</span><span class="sxs-lookup"><span data-stu-id="99b1d-410">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="99b1d-411">1.0</span><span class="sxs-lookup"><span data-stu-id="99b1d-411">1.0</span></span>|
|[<span data-ttu-id="99b1d-412">Mindestberechtigungsstufe</span><span class="sxs-lookup"><span data-stu-id="99b1d-412">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="99b1d-413">ReadItem</span><span class="sxs-lookup"><span data-stu-id="99b1d-413">ReadItem</span></span>|
|[<span data-ttu-id="99b1d-414">Zutreffender Outlook-Modus</span><span class="sxs-lookup"><span data-stu-id="99b1d-414">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="99b1d-415">Lesen</span><span class="sxs-lookup"><span data-stu-id="99b1d-415">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="99b1d-416">Beispiel</span><span class="sxs-lookup"><span data-stu-id="99b1d-416">Example</span></span>

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

#### <a name="displaynewmessageformparameters"></a><span data-ttu-id="99b1d-417">displayNewMessageForm(parameters)</span><span class="sxs-lookup"><span data-stu-id="99b1d-417">displayNewMessageForm(parameters)</span></span>

<span data-ttu-id="99b1d-418">Zeigt ein Formular zum Erstellen einer neuen Nachricht an.</span><span class="sxs-lookup"><span data-stu-id="99b1d-418">Displays a form for creating a new message.</span></span>

<span data-ttu-id="99b1d-419">Mit der `displayNewMessageForm`-Methode wird ein Formular geöffnet, mit dem der Benutzer eine neue Nachricht erstellen kann.</span><span class="sxs-lookup"><span data-stu-id="99b1d-419">The `displayNewMessageForm` method opens a form that enables the user to create a new message.</span></span> <span data-ttu-id="99b1d-420">Wenn Parameter angegeben wurden, werden die Felder im Nachrichtenformular automatisch mit dem Inhalt der Parameter ausgefüllt.</span><span class="sxs-lookup"><span data-stu-id="99b1d-420">If parameters are specified, the message form fields are automatically populated with the contents of the parameters.</span></span>

<span data-ttu-id="99b1d-421">Wenn einer der Parameter die angegebenen Größenbeschränkungen überschreitet oder wenn ein unbekannter Parametername angegeben wird, wird eine Ausnahme ausgelöst.</span><span class="sxs-lookup"><span data-stu-id="99b1d-421">If any of the parameters exceed the specified size limits, or if an unknown parameter name is specified, an exception is thrown.</span></span>

##### <a name="parameters"></a><span data-ttu-id="99b1d-422">Parameter:</span><span class="sxs-lookup"><span data-stu-id="99b1d-422">Parameters:</span></span>

> [!NOTE]
> <span data-ttu-id="99b1d-423">Alle Parameter sind optional.</span><span class="sxs-lookup"><span data-stu-id="99b1d-423">All parameters are optional.</span></span>

|<span data-ttu-id="99b1d-424">Name</span><span class="sxs-lookup"><span data-stu-id="99b1d-424">Name</span></span>| <span data-ttu-id="99b1d-425">Typ</span><span class="sxs-lookup"><span data-stu-id="99b1d-425">Type</span></span>| <span data-ttu-id="99b1d-426">Beschreibung</span><span class="sxs-lookup"><span data-stu-id="99b1d-426">Description</span></span>|
|---|---|---|
| `parameters` | <span data-ttu-id="99b1d-427">Object</span><span class="sxs-lookup"><span data-stu-id="99b1d-427">Object</span></span> | <span data-ttu-id="99b1d-428">Ein Wörterbuch mit Parametern, die die neue Nachricht beschreiben.</span><span class="sxs-lookup"><span data-stu-id="99b1d-428">A dictionary of parameters describing the new message.</span></span> |
| `parameters.toRecipients` | <span data-ttu-id="99b1d-429">Array.&lt;String&gt; &#124; Array.&lt;[EmailAddressDetails](/javascript/api/outlook_1_7/office.emailaddressdetails)&gt;</span><span class="sxs-lookup"><span data-stu-id="99b1d-429">Array.&lt;String&gt; &#124; Array.&lt;[EmailAddressDetails](/javascript/api/outlook_1_7/office.emailaddressdetails)&gt;</span></span> | <span data-ttu-id="99b1d-430">Ein Array aus Zeichenfolgen mit den E-Mail-Adressen oder ein Array mit einem `EmailAddressDetails`-Objekt für jeden der Empfänger im Feld „An“.</span><span class="sxs-lookup"><span data-stu-id="99b1d-430">An array of strings containing the email addresses or an array containing an `EmailAddressDetails` object for each of the recipients on the To line.</span></span> <span data-ttu-id="99b1d-431">Das Array darf maximal 100 Einträge enthalten.</span><span class="sxs-lookup"><span data-stu-id="99b1d-431">The array is limited to a maximum of 100 entries.</span></span> |
| `parameters.ccRecipients` | <span data-ttu-id="99b1d-432">Array.&lt;String&gt; &#124; Array.&lt;[EmailAddressDetails](/javascript/api/outlook_1_7/office.emailaddressdetails)&gt;</span><span class="sxs-lookup"><span data-stu-id="99b1d-432">Array.&lt;String&gt; &#124; Array.&lt;[EmailAddressDetails](/javascript/api/outlook_1_7/office.emailaddressdetails)&gt;</span></span> | <span data-ttu-id="99b1d-433">Ein Array aus Zeichenfolgen mit den E-Mail-Adressen oder ein Array mit einem `EmailAddressDetails`-Objekt für jeden der Empfänger im Feld „Cc“.</span><span class="sxs-lookup"><span data-stu-id="99b1d-433">An array of strings containing the email addresses or an array containing an `EmailAddressDetails` object for each of the recipients on the Cc line.</span></span> <span data-ttu-id="99b1d-434">Das Array darf maximal 100 Einträge enthalten.</span><span class="sxs-lookup"><span data-stu-id="99b1d-434">The array is limited to a maximum of 100 entries.</span></span> |
| `parameters.bccRecipients` | <span data-ttu-id="99b1d-435">Array.&lt;String&gt; &#124; Array.&lt;[EmailAddressDetails](/javascript/api/outlook_1_7/office.emailaddressdetails)&gt;</span><span class="sxs-lookup"><span data-stu-id="99b1d-435">Array.&lt;String&gt; &#124; Array.&lt;[EmailAddressDetails](/javascript/api/outlook_1_7/office.emailaddressdetails)&gt;</span></span> | <span data-ttu-id="99b1d-436">Ein Array aus Zeichenfolgen mit den E-Mail-Adressen oder ein Array mit einem `EmailAddressDetails`-Objekt für jeden der Empfänger im Feld „Bcc“.</span><span class="sxs-lookup"><span data-stu-id="99b1d-436">An array of strings containing the email addresses or an array containing an `EmailAddressDetails` object for each of the recipients on the Bcc line.</span></span> <span data-ttu-id="99b1d-437">Das Array darf maximal 100 Einträge enthalten.</span><span class="sxs-lookup"><span data-stu-id="99b1d-437">The array is limited to a maximum of 100 entries.</span></span> |
| `parameters.subject` | <span data-ttu-id="99b1d-438">Zeichenfolge</span><span class="sxs-lookup"><span data-stu-id="99b1d-438">String</span></span> | <span data-ttu-id="99b1d-439">Eine Zeichenfolge mit dem Betreff für die Nachricht.</span><span class="sxs-lookup"><span data-stu-id="99b1d-439">A string containing the subject of the message.</span></span> <span data-ttu-id="99b1d-440">Die Zeichenfolge ist auf maximal 255 Zeichen beschränkt.</span><span class="sxs-lookup"><span data-stu-id="99b1d-440">The string is limited to a maximum of 255 characters.</span></span> |
| `parameters.htmlBody` | <span data-ttu-id="99b1d-441">Zeichenfolge</span><span class="sxs-lookup"><span data-stu-id="99b1d-441">String</span></span> | <span data-ttu-id="99b1d-442">Der HTML-Text der Nachricht.</span><span class="sxs-lookup"><span data-stu-id="99b1d-442">The HTML body of the message.</span></span> <span data-ttu-id="99b1d-443">Der Textinhalt darf maximal 32 KB umfassen.</span><span class="sxs-lookup"><span data-stu-id="99b1d-443">The body content is limited to a maximum size of 32 KB.</span></span> |
| `parameters.attachments` | <span data-ttu-id="99b1d-444">Array.&lt;Object&gt;</span><span class="sxs-lookup"><span data-stu-id="99b1d-444">Array.&lt;Object&gt;</span></span> | <span data-ttu-id="99b1d-445">Ein Array mit JSON-Objekten, die Datei- oder Elementanlagen sind.</span><span class="sxs-lookup"><span data-stu-id="99b1d-445">An array of JSON objects that are either file or item attachments.</span></span> |
| `parameters.attachments.type` | <span data-ttu-id="99b1d-446">Zeichenfolge</span><span class="sxs-lookup"><span data-stu-id="99b1d-446">String</span></span> | <span data-ttu-id="99b1d-p128">Gibt den Typ der Anlage an. Muss `file` für eine Dateianlage oder `item` für eine Elementanlage sein.</span><span class="sxs-lookup"><span data-stu-id="99b1d-p128">Indicates the type of attachment. Must be `file` for a file attachment or `item` for an item attachment.</span></span> |
| `parameters.attachments.name` | <span data-ttu-id="99b1d-449">Zeichenfolge</span><span class="sxs-lookup"><span data-stu-id="99b1d-449">String</span></span> | <span data-ttu-id="99b1d-450">Eine Zeichenfolge, die den Namen der Anlage mit bis zu 255 Zeichen enthält.</span><span class="sxs-lookup"><span data-stu-id="99b1d-450">A string that contains the name of the attachment, up to 255 characters in length.</span></span>|
| `parameters.attachments.url` | <span data-ttu-id="99b1d-451">Zeichenfolge</span><span class="sxs-lookup"><span data-stu-id="99b1d-451">String</span></span> | <span data-ttu-id="99b1d-p129">Wird nur verwendet, wenn `type` auf `file` gesetzt ist. Der URI des Speicherorts für die Datei.</span><span class="sxs-lookup"><span data-stu-id="99b1d-p129">Only used if `type` is set to `file`. The URI of the location for the file.</span></span> |
| `parameters.attachments.isInline` | <span data-ttu-id="99b1d-454">Boolean</span><span class="sxs-lookup"><span data-stu-id="99b1d-454">Boolean</span></span> | <span data-ttu-id="99b1d-p130">Wird nur verwendet, wenn `type` auf `file` gesetzt ist. Wenn `true`: Zeigt an, dass die Anlage inline im Nachrichtentext angezeigt wird und nicht in der Anlagenlisten angezeigt werden soll.</span><span class="sxs-lookup"><span data-stu-id="99b1d-p130">Only used if `type` is set to `file`. If `true`, indicates that the attachment will be shown inline in the message body, and should not be displayed in the attachment list.</span></span> |
| `parameters.attachments.itemId` | <span data-ttu-id="99b1d-457">Zeichenfolge</span><span class="sxs-lookup"><span data-stu-id="99b1d-457">String</span></span> | <span data-ttu-id="99b1d-458">Nur verwendet, wenn `type` auf `item` festgelegt ist.</span><span class="sxs-lookup"><span data-stu-id="99b1d-458">Only used if `type` is set to `item`.</span></span> <span data-ttu-id="99b1d-459">Der EWS-Element-Id der vorhandenen e-Mail, den Sie die neue Nachricht zuordnen möchten.</span><span class="sxs-lookup"><span data-stu-id="99b1d-459">The EWS item id of the existing e-mail you want to attach to the new message.</span></span> <span data-ttu-id="99b1d-460">Dies ist eine Zeichenfolge bis zu 100 Zeichen.</span><span class="sxs-lookup"><span data-stu-id="99b1d-460">This is a string up to 100 characters.</span></span> |


##### <a name="requirements"></a><span data-ttu-id="99b1d-461">Anforderungen</span><span class="sxs-lookup"><span data-stu-id="99b1d-461">Requirements</span></span>

|<span data-ttu-id="99b1d-462">Anforderung</span><span class="sxs-lookup"><span data-stu-id="99b1d-462">Requirement</span></span>| <span data-ttu-id="99b1d-463">Wert</span><span class="sxs-lookup"><span data-stu-id="99b1d-463">Value</span></span>|
|---|---|
|[<span data-ttu-id="99b1d-464">Mindestversion des Postfachanforderungssatzes</span><span class="sxs-lookup"><span data-stu-id="99b1d-464">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="99b1d-465">1.6</span><span class="sxs-lookup"><span data-stu-id="99b1d-465">1.6</span></span> |
|[<span data-ttu-id="99b1d-466">Mindestberechtigungsstufe</span><span class="sxs-lookup"><span data-stu-id="99b1d-466">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="99b1d-467">ReadItem</span><span class="sxs-lookup"><span data-stu-id="99b1d-467">ReadItem</span></span>|
|[<span data-ttu-id="99b1d-468">Zutreffender Outlook-Modus</span><span class="sxs-lookup"><span data-stu-id="99b1d-468">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="99b1d-469">Lesen</span><span class="sxs-lookup"><span data-stu-id="99b1d-469">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="99b1d-470">Beispiel</span><span class="sxs-lookup"><span data-stu-id="99b1d-470">Example</span></span>

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

#### <a name="getcallbacktokenasyncoptions-callback"></a><span data-ttu-id="99b1d-471">getCallbackTokenAsync([options], callback)</span><span class="sxs-lookup"><span data-stu-id="99b1d-471">getCallbackTokenAsync([options], callback)</span></span>

<span data-ttu-id="99b1d-472">Ruft eine Zeichenfolge mit einem Token ab, das zum Aufrufen von REST-APIs oder Exchange-Webdiensten verwendet wird.</span><span class="sxs-lookup"><span data-stu-id="99b1d-472">Gets a string that contains a token used to call REST APIs or Exchange Web Services.</span></span>

<span data-ttu-id="99b1d-p132">Die `getCallbackTokenAsync`-Methode führt einen asynchronen Aufruf zum Abruf eines nicht transparenten Tokens vom Exchange-Server aus, der das Postfach des Benutzers hostet. Die Gültigkeitsdauer des Rückruftokens beträgt 5 Minuten.</span><span class="sxs-lookup"><span data-stu-id="99b1d-p132">The `getCallbackTokenAsync` method makes an asynchronous call to get an opaque token from the Exchange Server that hosts the user's mailbox. The lifetime of the callback token is 5 minutes.</span></span>

> [!NOTE]
> <span data-ttu-id="99b1d-475">Es wird empfohlen, dass-add-ins anstelle von Exchange-Webdienste nach Möglichkeit die REST-APIs verwenden.</span><span class="sxs-lookup"><span data-stu-id="99b1d-475">It is recommended that add-ins use the REST APIs instead of Exchange Web Services whenever possible.</span></span> 

<span data-ttu-id="99b1d-476">**REST-Tokens**</span><span class="sxs-lookup"><span data-stu-id="99b1d-476">**REST Tokens**</span></span>

<span data-ttu-id="99b1d-p133">Wenn ein REST-Token angefordert wird (`options.isRest = true`), kann das resultierende Token nicht zur Authentifizierung von Aufrufen der Exchange-Webdienste verwendet werden. Der Bereich des Tokens ist auf schreibgeschützten Zugriff auf das aktuelle Element und dessen Anlagen beschränkt, es sei denn im Manifest des Add-Ins ist die Berechtigung [`ReadWriteMailbox`](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions#readwritemailbox-permission) angegeben. Ist die Berechtigung `ReadWriteMailbox` angegeben, gewährt das resultierende Token Lese-/Schreibzugriff auf E-Mails, Kalender und Kontakte und erlaubt das Senden von E-Mails.</span><span class="sxs-lookup"><span data-stu-id="99b1d-p133">When a REST token is requested (`options.isRest = true`), the resulting token will not work to authenticate Exchange Web Services calls. The token will be limited in scope to read-only access to the current item and its attachments, unless the add-in has specified the [`ReadWriteMailbox`](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions#readwritemailbox-permission) permission in its manifest. If the `ReadWriteMailbox` permission is specified, the resulting token will grant read/write access to mail, calendar, and contacts, including the ability to send mail.</span></span>

<span data-ttu-id="99b1d-480">Das Add-In sollte die Eigenschaft `restUrl` verwenden, um die korrekte URL für REST-API-Aufrufe zu ermitteln.</span><span class="sxs-lookup"><span data-stu-id="99b1d-480">The add-in should use the `restUrl` property to determine the correct URL to use when making REST API calls.</span></span>

<span data-ttu-id="99b1d-481">**EWS-Tokens**</span><span class="sxs-lookup"><span data-stu-id="99b1d-481">**EWS Tokens**</span></span>

<span data-ttu-id="99b1d-p134">Wenn ein EWS-Token angefordert wird (`options.isRest = false`), kann das resultierende Token nicht zur Authentifizierung von REST-API-Aufrufen verwendet werden. Der Bereich des Tokens ist auf den Zugriff auf das aktuelle Element beschränkt.</span><span class="sxs-lookup"><span data-stu-id="99b1d-p134">When an EWS token is requested (`options.isRest = false`), the resulting token will not work to authenticate REST API calls. The token will be limited in scope to accessing the current item.</span></span>

<span data-ttu-id="99b1d-484">Das Add-In sollte die Eigenschaft `ewsUrl` verwenden, um die korrekte URL für EWS-Aufrufe zu ermitteln.</span><span class="sxs-lookup"><span data-stu-id="99b1d-484">The add-in should use the `ewsUrl` property to determine the correct URL to use when making EWS calls.</span></span>

##### <a name="parameters"></a><span data-ttu-id="99b1d-485">Parameter:</span><span class="sxs-lookup"><span data-stu-id="99b1d-485">Parameters:</span></span>

|<span data-ttu-id="99b1d-486">Name</span><span class="sxs-lookup"><span data-stu-id="99b1d-486">Name</span></span>| <span data-ttu-id="99b1d-487">Typ</span><span class="sxs-lookup"><span data-stu-id="99b1d-487">Type</span></span>| <span data-ttu-id="99b1d-488">Attribute</span><span class="sxs-lookup"><span data-stu-id="99b1d-488">Attributes</span></span>| <span data-ttu-id="99b1d-489">Beschreibung</span><span class="sxs-lookup"><span data-stu-id="99b1d-489">Description</span></span>|
|---|---|---|---|
| `options` | <span data-ttu-id="99b1d-490">Objekt</span><span class="sxs-lookup"><span data-stu-id="99b1d-490">Object</span></span> | <span data-ttu-id="99b1d-491">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="99b1d-491">&lt;optional&gt;</span></span> | <span data-ttu-id="99b1d-492">Ein Objektliteral, das mindestens eine der folgenden Eigenschaften enthält.</span><span class="sxs-lookup"><span data-stu-id="99b1d-492">An object literal that contains one or more of the following properties.</span></span> |
| `options.isRest` | <span data-ttu-id="99b1d-493">Boolean</span><span class="sxs-lookup"><span data-stu-id="99b1d-493">Boolean</span></span> |  <span data-ttu-id="99b1d-494">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="99b1d-494">&lt;optional&gt;</span></span> | <span data-ttu-id="99b1d-p135">Legt fest, ob das bereitgestellte Token für die Outlook-REST-APIs oder die Exchange-Webdienste verwendet wird. Der Standardwert ist `false`.</span><span class="sxs-lookup"><span data-stu-id="99b1d-p135">Determines if the token provided will be used for the Outlook REST APIs or Exchange Web Services. Default value is `false`.</span></span> |
| `options.asyncContext` | <span data-ttu-id="99b1d-497">Objekt</span><span class="sxs-lookup"><span data-stu-id="99b1d-497">Object</span></span> |  <span data-ttu-id="99b1d-498">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="99b1d-498">&lt;optional&gt;</span></span> | <span data-ttu-id="99b1d-499">Jegliche Statusdaten, die an die asynchrone Methode übergeben werden</span><span class="sxs-lookup"><span data-stu-id="99b1d-499">Any state data that is passed to the asynchronous method.</span></span> |
|`callback`| <span data-ttu-id="99b1d-500">function</span><span class="sxs-lookup"><span data-stu-id="99b1d-500">function</span></span>||<span data-ttu-id="99b1d-p136">Wenn die Methode abgeschlossen wird, wird die Funktion , die im `callback`-Parameter übergeben wird, mit einem einzelnen Parameter (`asyncResult`) aufgerufen, der ein [`AsyncResult`](/javascript/api/office/office.asyncresult)-Objekt darstellt. Das Token wird als Zeichenfolge in der `asyncResult.value`-Eigenschaft bereitgestellt.</span><span class="sxs-lookup"><span data-stu-id="99b1d-p136">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object. The token is provided as a string in the `asyncResult.value` property.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="99b1d-503">Anforderungen</span><span class="sxs-lookup"><span data-stu-id="99b1d-503">Requirements</span></span>

|<span data-ttu-id="99b1d-504">Anforderung</span><span class="sxs-lookup"><span data-stu-id="99b1d-504">Requirement</span></span>| <span data-ttu-id="99b1d-505">Wert</span><span class="sxs-lookup"><span data-stu-id="99b1d-505">Value</span></span>|
|---|---|
|[<span data-ttu-id="99b1d-506">Mindestversion des Postfachanforderungssatzes</span><span class="sxs-lookup"><span data-stu-id="99b1d-506">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="99b1d-507">1,5</span><span class="sxs-lookup"><span data-stu-id="99b1d-507">1.5</span></span> |
|[<span data-ttu-id="99b1d-508">Mindestberechtigungsstufe</span><span class="sxs-lookup"><span data-stu-id="99b1d-508">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="99b1d-509">ReadItem</span><span class="sxs-lookup"><span data-stu-id="99b1d-509">ReadItem</span></span>|
|[<span data-ttu-id="99b1d-510">Zutreffender Outlook-Modus</span><span class="sxs-lookup"><span data-stu-id="99b1d-510">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="99b1d-511">Verfassenmodus und Lesemodus</span><span class="sxs-lookup"><span data-stu-id="99b1d-511">Compose and read</span></span>|

##### <a name="example"></a><span data-ttu-id="99b1d-512">Beispiel</span><span class="sxs-lookup"><span data-stu-id="99b1d-512">Example</span></span>

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

#### <a name="getcallbacktokenasynccallback-usercontext"></a><span data-ttu-id="99b1d-513">getCallbackTokenAsync(callback, [userContext])</span><span class="sxs-lookup"><span data-stu-id="99b1d-513">getCallbackTokenAsync(callback, [userContext])</span></span>

<span data-ttu-id="99b1d-514">Ruft eine Zeichenfolge ab, die einen Token enthält, der verwendet wird, um eine Anlage oder ein Element von einem Exchange Server abzurufen.</span><span class="sxs-lookup"><span data-stu-id="99b1d-514">Gets a string that contains a token used to get an attachment or item from an Exchange Server.</span></span>

<span data-ttu-id="99b1d-p137">Die `getCallbackTokenAsync`-Methode führt einen asynchronen Aufruf zum Abruf eines nicht transparenten Tokens vom Exchange-Server aus, der das Postfach des Benutzers hostet. Die Gültigkeitsdauer des Rückruftokens beträgt 5 Minuten.</span><span class="sxs-lookup"><span data-stu-id="99b1d-p137">The `getCallbackTokenAsync` method makes an asynchronous call to get an opaque token from the Exchange Server that hosts the user's mailbox. The lifetime of the callback token is 5 minutes.</span></span>

<span data-ttu-id="99b1d-p138">Sie können das Token und einen Anlagen- oder einen Elementbezeichner an ein Drittanbietersystem weitergeben. Das Drittanbietersystem verwendet das Token als Trägerautorisierungstoken, um den EWS-Vorgang (Exchange-Webdienste) [GetAttachment](https://docs.microsoft.com/exchange/client-developer/web-service-reference/getattachment-operation) oder [GetItem](https://docs.microsoft.com/exchange/client-developer/web-service-reference/getitem-operation) aufzurufen und eine Anlage oder ein Element zurückzugeben. Sie können z. B. einen Remotedienst erstellen, um [Anlagen aus dem ausgewählten Element abzurufen](https://docs.microsoft.com/outlook/add-ins/get-attachments-of-an-outlook-item).</span><span class="sxs-lookup"><span data-stu-id="99b1d-p138">You can pass the token and an attachment identifier or item identifier to a third-party system. The third-party system uses the token as a bearer authorization token to call the Exchange Web Services (EWS) [GetAttachment](https://docs.microsoft.com/exchange/client-developer/web-service-reference/getattachment-operation) or [GetItem](https://docs.microsoft.com/exchange/client-developer/web-service-reference/getitem-operation) operation to return an attachment or item. For example, you can create a remote service to [get attachments from the selected item](https://docs.microsoft.com/outlook/add-ins/get-attachments-of-an-outlook-item).</span></span>

<span data-ttu-id="99b1d-520">Ihre App muss über die im Manifest angegebene **ReadItem**-Berechtigung verfügen, um die `getCallbackTokenAsync`-Methode im Lesemodus aufzurufen.</span><span class="sxs-lookup"><span data-stu-id="99b1d-520">Your app must have the **ReadItem** permission specified in its manifest to call the `getCallbackTokenAsync` method in read mode.</span></span>

<span data-ttu-id="99b1d-p139">Im Verfassenmodus müssen Sie die [`saveAsync`](Office.context.mailbox.item.md#saveasyncoptions-callback)-Methode aufrufen, um einen Elementbezeichner an die `getCallbackTokenAsync`-Methode zu übergeben. Die App muss über **ReadWriteItem**-Berechtigungen verfügen, um die `saveAsync`-Methode aufzurufen.</span><span class="sxs-lookup"><span data-stu-id="99b1d-p139">In compose mode you must call the [`saveAsync`](Office.context.mailbox.item.md#saveasyncoptions-callback) method to get an item identifier to pass to the `getCallbackTokenAsync` method. Your app must have **ReadWriteItem** permissions to call the `saveAsync` method.</span></span>

##### <a name="parameters"></a><span data-ttu-id="99b1d-523">Parameter:</span><span class="sxs-lookup"><span data-stu-id="99b1d-523">Parameters:</span></span>

|<span data-ttu-id="99b1d-524">Name</span><span class="sxs-lookup"><span data-stu-id="99b1d-524">Name</span></span>| <span data-ttu-id="99b1d-525">Typ</span><span class="sxs-lookup"><span data-stu-id="99b1d-525">Type</span></span>| <span data-ttu-id="99b1d-526">Attribute</span><span class="sxs-lookup"><span data-stu-id="99b1d-526">Attributes</span></span>| <span data-ttu-id="99b1d-527">Beschreibung</span><span class="sxs-lookup"><span data-stu-id="99b1d-527">Description</span></span>|
|---|---|---|---|
|`callback`| <span data-ttu-id="99b1d-528">Funktion</span><span class="sxs-lookup"><span data-stu-id="99b1d-528">function</span></span>||<span data-ttu-id="99b1d-p140">Wenn die Methode abgeschlossen wird, wird die Funktion , die im `callback`-Parameter übergeben wird, mit einem einzelnen Parameter (`asyncResult`) aufgerufen, der ein [`AsyncResult`](/javascript/api/office/office.asyncresult)-Objekt darstellt. Das Token wird als Zeichenfolge in der `asyncResult.value`-Eigenschaft bereitgestellt.</span><span class="sxs-lookup"><span data-stu-id="99b1d-p140">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object. The token is provided as a string in the `asyncResult.value` property.</span></span>|
|`userContext`| <span data-ttu-id="99b1d-531">Objekt</span><span class="sxs-lookup"><span data-stu-id="99b1d-531">Object</span></span>| <span data-ttu-id="99b1d-532">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="99b1d-532">&lt;optional&gt;</span></span>|<span data-ttu-id="99b1d-533">Jegliche Zustandsdaten, die an die asynchrone Methode übergeben werden.</span><span class="sxs-lookup"><span data-stu-id="99b1d-533">Any state data that is passed to the asynchronous method.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="99b1d-534">Anforderungen</span><span class="sxs-lookup"><span data-stu-id="99b1d-534">Requirements</span></span>

|<span data-ttu-id="99b1d-535">Anforderung</span><span class="sxs-lookup"><span data-stu-id="99b1d-535">Requirement</span></span>| <span data-ttu-id="99b1d-536">Wert</span><span class="sxs-lookup"><span data-stu-id="99b1d-536">Value</span></span>|
|---|---|
|[<span data-ttu-id="99b1d-537">Mindestversion des Postfachanforderungssatzes</span><span class="sxs-lookup"><span data-stu-id="99b1d-537">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="99b1d-538">1.3</span><span class="sxs-lookup"><span data-stu-id="99b1d-538">1.3</span></span>|
|[<span data-ttu-id="99b1d-539">Mindestberechtigungsstufe</span><span class="sxs-lookup"><span data-stu-id="99b1d-539">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="99b1d-540">ReadItem</span><span class="sxs-lookup"><span data-stu-id="99b1d-540">ReadItem</span></span>|
|[<span data-ttu-id="99b1d-541">Zutreffender Outlook-Modus</span><span class="sxs-lookup"><span data-stu-id="99b1d-541">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="99b1d-542">Verfassenmodus und Lesemodus</span><span class="sxs-lookup"><span data-stu-id="99b1d-542">Compose and read</span></span>|

##### <a name="example"></a><span data-ttu-id="99b1d-543">Beispiel</span><span class="sxs-lookup"><span data-stu-id="99b1d-543">Example</span></span>

```js
function getCallbackToken() {
  Office.context.mailbox.getCallbackTokenAsync(cb);
}

function cb(asyncResult) {
  var token = asyncResult.value;
}
```

####  <a name="getuseridentitytokenasynccallback-usercontext"></a><span data-ttu-id="99b1d-544">getUserIdentityTokenAsync(callback, [userContext])</span><span class="sxs-lookup"><span data-stu-id="99b1d-544">getUserIdentityTokenAsync(callback, [userContext])</span></span>

<span data-ttu-id="99b1d-545">Ruft ein Token ab, das den Benutzer und das Office-Add-In identifiziert.</span><span class="sxs-lookup"><span data-stu-id="99b1d-545">Gets a token identifying the user and the Office Add-in.</span></span>

<span data-ttu-id="99b1d-546">Die `getUserIdentityTokenAsync`-Methode gibt ein Token zurück, das Sie dazu verwenden können, um das [Add-In und den Benutzer mit einem Drittanbietersystem zu identifizieren und zu authentifizieren](https://docs.microsoft.com/outlook/add-ins/authentication).</span><span class="sxs-lookup"><span data-stu-id="99b1d-546">The `getUserIdentityTokenAsync` method returns a token that you can use to identify and [authenticate the add-in and user with a third-party system](https://docs.microsoft.com/outlook/add-ins/authentication).</span></span>

##### <a name="parameters"></a><span data-ttu-id="99b1d-547">Parameter:</span><span class="sxs-lookup"><span data-stu-id="99b1d-547">Parameters:</span></span>

|<span data-ttu-id="99b1d-548">Name</span><span class="sxs-lookup"><span data-stu-id="99b1d-548">Name</span></span>| <span data-ttu-id="99b1d-549">Typ</span><span class="sxs-lookup"><span data-stu-id="99b1d-549">Type</span></span>| <span data-ttu-id="99b1d-550">Attribute</span><span class="sxs-lookup"><span data-stu-id="99b1d-550">Attributes</span></span>| <span data-ttu-id="99b1d-551">Beschreibung</span><span class="sxs-lookup"><span data-stu-id="99b1d-551">Description</span></span>|
|---|---|---|---|
|`callback`| <span data-ttu-id="99b1d-552">Funktion</span><span class="sxs-lookup"><span data-stu-id="99b1d-552">function</span></span>||<span data-ttu-id="99b1d-553">Wenn die Methode abgeschlossen wird, wird die Funktion , die im `callback`-Parameter übergeben wird, mit einem einzelnen Parameter (`asyncResult`) aufgerufen, der ein [`AsyncResult`](/javascript/api/office/office.asyncresult)-Objekt darstellt.</span><span class="sxs-lookup"><span data-stu-id="99b1d-553">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span><br/><br/><span data-ttu-id="99b1d-554">Das Token wird als Zeichenfolge in der `asyncResult.value`-Eigenschaft bereitgestellt.</span><span class="sxs-lookup"><span data-stu-id="99b1d-554">The token is provided as a string in the `asyncResult.value` property.</span></span>|
|`userContext`| <span data-ttu-id="99b1d-555">Objekt</span><span class="sxs-lookup"><span data-stu-id="99b1d-555">Object</span></span>| <span data-ttu-id="99b1d-556">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="99b1d-556">&lt;optional&gt;</span></span>|<span data-ttu-id="99b1d-557">Jegliche Zustandsdaten, die an die asynchrone Methode übergeben werden.</span><span class="sxs-lookup"><span data-stu-id="99b1d-557">Any state data that is passed to the asynchronous method.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="99b1d-558">Anforderungen</span><span class="sxs-lookup"><span data-stu-id="99b1d-558">Requirements</span></span>

|<span data-ttu-id="99b1d-559">Anforderung</span><span class="sxs-lookup"><span data-stu-id="99b1d-559">Requirement</span></span>| <span data-ttu-id="99b1d-560">Wert</span><span class="sxs-lookup"><span data-stu-id="99b1d-560">Value</span></span>|
|---|---|
|[<span data-ttu-id="99b1d-561">Mindestversion des Postfachanforderungssatzes</span><span class="sxs-lookup"><span data-stu-id="99b1d-561">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="99b1d-562">1.0</span><span class="sxs-lookup"><span data-stu-id="99b1d-562">1.0</span></span>|
|[<span data-ttu-id="99b1d-563">Mindestberechtigungsstufe</span><span class="sxs-lookup"><span data-stu-id="99b1d-563">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="99b1d-564">ReadItem</span><span class="sxs-lookup"><span data-stu-id="99b1d-564">ReadItem</span></span>|
|[<span data-ttu-id="99b1d-565">Zutreffender Outlook-Modus</span><span class="sxs-lookup"><span data-stu-id="99b1d-565">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="99b1d-566">Verfassen oder Lesen</span><span class="sxs-lookup"><span data-stu-id="99b1d-566">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="99b1d-567">Beispiel</span><span class="sxs-lookup"><span data-stu-id="99b1d-567">Example</span></span>

```js
function getIdentityToken() {
  Office.context.mailbox.getUserIdentityTokenAsync(cb);
}

function cb(asyncResult) {
  var token = asyncResult.value;
}
```

####  <a name="makeewsrequestasyncdata-callback-usercontext"></a><span data-ttu-id="99b1d-568">makeEwsRequestAsync(data, callback, [userContext])</span><span class="sxs-lookup"><span data-stu-id="99b1d-568">makeEwsRequestAsync(data, callback, [userContext])</span></span>

<span data-ttu-id="99b1d-569">Richtet eine asynchrone Anforderung an einen EWS-Dienst (Exchange Web Services) auf dem Exchange-Server, der das Postfach des Benutzers hostet.</span><span class="sxs-lookup"><span data-stu-id="99b1d-569">Makes an asynchronous request to an Exchange Web Services (EWS) service on the Exchange server that hosts the user’s mailbox.</span></span>

> [!NOTE]
> <span data-ttu-id="99b1d-570">Diese Methode wird in den folgenden Szenarien nicht unterstützt.</span><span class="sxs-lookup"><span data-stu-id="99b1d-570">This method is not supported in the following scenarios.</span></span>
> - <span data-ttu-id="99b1d-571">In Outlook für iOS oder Outlook für Android</span><span class="sxs-lookup"><span data-stu-id="99b1d-571">In Outlook for iOS or Outlook for Android</span></span>
> - <span data-ttu-id="99b1d-572">Wenn das Add-In geladen wird, in einem Postfach Google Mail</span><span class="sxs-lookup"><span data-stu-id="99b1d-572">When the add-in is loaded in a Gmail mailbox</span></span>
> 
> <span data-ttu-id="99b1d-573">In diesen Fällen sollte-add-ins [mithilfe von REST-APIs](https://docs.microsoft.com/outlook/add-ins/use-rest-api) , stattdessen Zugriff auf das Postfach des Benutzers.</span><span class="sxs-lookup"><span data-stu-id="99b1d-573">In these cases, add-ins should [use REST APIs](https://docs.microsoft.com/outlook/add-ins/use-rest-api) to access the user's mailbox instead.</span></span>

<span data-ttu-id="99b1d-574">Die `makeEwsRequestAsync`-Methode sendet eine EWS-Anforderung für das Add-In zu Exchange.</span><span class="sxs-lookup"><span data-stu-id="99b1d-574">The `makeEwsRequestAsync` method sends an EWS request on behalf of the add-in to Exchange.</span></span> <span data-ttu-id="99b1d-575">Eine Liste der unterstützten EWS-Vorgänge finden Sie unter [Aufrufen von Webdiensten aus einem Outlook-add-in](https://docs.microsoft.com/outlook/add-ins/web-services#ews-operations-that-add-ins-support) .</span><span class="sxs-lookup"><span data-stu-id="99b1d-575">See [Call web services from an Outlook add-in](https://docs.microsoft.com/outlook/add-ins/web-services#ews-operations-that-add-ins-support) for a list of the supported EWS operations.</span></span>

<span data-ttu-id="99b1d-576">Sie können keine Elemente, die Ordnern zugeordnet sind, mit der `makeEwsRequestAsync`-Methode anfordern.</span><span class="sxs-lookup"><span data-stu-id="99b1d-576">You cannot request Folder Associated Items with the `makeEwsRequestAsync` method.</span></span>

<span data-ttu-id="99b1d-577">Die XML-Anfrage muss UTF-8-Codierung angeben.</span><span class="sxs-lookup"><span data-stu-id="99b1d-577">The XML request must specify UTF-8 encoding.</span></span>

```
<?xml version="1.0" encoding="utf-8"?>
```

<span data-ttu-id="99b1d-p142">Das Add-In muss über die **ReadWriteMailbox**-Berechtigung verfügen, um die `makeEwsRequestAsync`-Methode zu verwenden. Informationen zur **ReadWriteMailbox**-Berechtigung und zu den EWS-Vorgängen, die Sie mit der `makeEwsRequestAsync`-Methode aufrufen können, finden Sie unter [Angeben von Berechtigungen für den E-Mail-Add-In-Zugriff auf das Benutzerpostfach](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions).</span><span class="sxs-lookup"><span data-stu-id="99b1d-p142">Your add-in must have the **ReadWriteMailbox** permission to use the `makeEwsRequestAsync` method. For information about using the **ReadWriteMailbox** permission and the EWS operations that you can call with the `makeEwsRequestAsync` method, see [Specify permissions for mail add-in access to the user's mailbox](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions).</span></span>

> [!NOTE]
> <span data-ttu-id="99b1d-580">Der Serveradministrator muss festgelegt `OAuthAuthentication` auf "true" auf dem Client Access Server EWS-Verzeichnis so aktivieren Sie die `makeEwsRequestAsync` -Methode, um EWS stellen anfordert.</span><span class="sxs-lookup"><span data-stu-id="99b1d-580">The server administrator must set `OAuthAuthentication` to true on the Client Access Server EWS directory to enable the `makeEwsRequestAsync` method to make EWS requests.</span></span>

##### <a name="version-differences"></a><span data-ttu-id="99b1d-581">Versionsunterschiede</span><span class="sxs-lookup"><span data-stu-id="99b1d-581">Version differences</span></span>

<span data-ttu-id="99b1d-582">Wenn Sie die `makeEwsRequestAsync`-Methode in Mail-Apps verwenden, die in älteren Outlook-Versionen als Version 15.0.4535.1004 ausgeführt werden, sollten Sie den Codierungswert auf `ISO-8859-1` festlegen.</span><span class="sxs-lookup"><span data-stu-id="99b1d-582">When you use the `makeEwsRequestAsync` method in mail apps running in Outlook versions earlier than version 15.0.4535.1004, you should set the encoding value to `ISO-8859-1`.</span></span>

```
<?xml version="1.0" encoding="iso-8859-1"?>
```

<span data-ttu-id="99b1d-p143">Sie müssen den Codierungswert nicht festlegen, wenn Ihre Mail-App in Outlook im Web ausgeführt wird. Sie können mithilfe der mailbox.diagnostics.hostName-Eigenschaft ermitteln, ob Ihre Mail-App in Outlook oder Outlook im Web ausgeführt wird. Sie können mithilfe der mailbox.diagnostics.hostVersion-Eigenschaft ermitteln, welche Version von Outlook ausgeführt wird.</span><span class="sxs-lookup"><span data-stu-id="99b1d-p143">You do not need to set the encoding value when your mail app is running in Outlook on the web. You can determine whether your mail app is running in Outlook or Outlook on the web by using the mailbox.diagnostics.hostName property. You can determine what version of Outlook is running by using the mailbox.diagnostics.hostVersion property.</span></span>

##### <a name="parameters"></a><span data-ttu-id="99b1d-586">Parameter:</span><span class="sxs-lookup"><span data-stu-id="99b1d-586">Parameters:</span></span>

|<span data-ttu-id="99b1d-587">Name</span><span class="sxs-lookup"><span data-stu-id="99b1d-587">Name</span></span>| <span data-ttu-id="99b1d-588">Typ</span><span class="sxs-lookup"><span data-stu-id="99b1d-588">Type</span></span>| <span data-ttu-id="99b1d-589">Attribute</span><span class="sxs-lookup"><span data-stu-id="99b1d-589">Attributes</span></span>| <span data-ttu-id="99b1d-590">Beschreibung</span><span class="sxs-lookup"><span data-stu-id="99b1d-590">Description</span></span>|
|---|---|---|---|
|`data`| <span data-ttu-id="99b1d-591">Zeichenfolge</span><span class="sxs-lookup"><span data-stu-id="99b1d-591">String</span></span>||<span data-ttu-id="99b1d-592">Die EWS-Anforderung.</span><span class="sxs-lookup"><span data-stu-id="99b1d-592">The EWS request.</span></span>|
|`callback`| <span data-ttu-id="99b1d-593">Funktion</span><span class="sxs-lookup"><span data-stu-id="99b1d-593">function</span></span>||<span data-ttu-id="99b1d-594">Wenn die Methode abgeschlossen wird, wird die Funktion , die im `callback`-Parameter übergeben wird, mit einem einzelnen Parameter (`asyncResult`) aufgerufen, der ein [`AsyncResult`](/javascript/api/office/office.asyncresult)-Objekt darstellt.</span><span class="sxs-lookup"><span data-stu-id="99b1d-594">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span><br/><br/><span data-ttu-id="99b1d-595">Das XML-Ergebnis des EWS-Aufrufs wird als Zeichenfolge in der `asyncResult.value`-Eigenschaft bereitgestellt.</span><span class="sxs-lookup"><span data-stu-id="99b1d-595">The XML result of the EWS call is provided as a string in the `asyncResult.value` property.</span></span> <span data-ttu-id="99b1d-596">Wenn das Ergebnis 1 MB Größe überschreitet, wird stattdessen eine Fehlermeldung zurückgegeben.</span><span class="sxs-lookup"><span data-stu-id="99b1d-596">If the result exceeds 1 MB in size, an error message is returned instead.</span></span>|
|`userContext`| <span data-ttu-id="99b1d-597">Objekt</span><span class="sxs-lookup"><span data-stu-id="99b1d-597">Object</span></span>| <span data-ttu-id="99b1d-598">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="99b1d-598">&lt;optional&gt;</span></span>|<span data-ttu-id="99b1d-599">Jegliche Zustandsdaten, die an die asynchrone Methode übergeben werden.</span><span class="sxs-lookup"><span data-stu-id="99b1d-599">Any state data that is passed to the asynchronous method.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="99b1d-600">Anforderungen</span><span class="sxs-lookup"><span data-stu-id="99b1d-600">Requirements</span></span>

|<span data-ttu-id="99b1d-601">Anforderung</span><span class="sxs-lookup"><span data-stu-id="99b1d-601">Requirement</span></span>| <span data-ttu-id="99b1d-602">Wert</span><span class="sxs-lookup"><span data-stu-id="99b1d-602">Value</span></span>|
|---|---|
|[<span data-ttu-id="99b1d-603">Mindestversion des Postfachanforderungssatzes</span><span class="sxs-lookup"><span data-stu-id="99b1d-603">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="99b1d-604">1.0</span><span class="sxs-lookup"><span data-stu-id="99b1d-604">1.0</span></span>|
|[<span data-ttu-id="99b1d-605">Mindestberechtigungsstufe</span><span class="sxs-lookup"><span data-stu-id="99b1d-605">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="99b1d-606">ReadWriteMailbox</span><span class="sxs-lookup"><span data-stu-id="99b1d-606">ReadWriteMailbox</span></span>|
|[<span data-ttu-id="99b1d-607">Zutreffender Outlook-Modus</span><span class="sxs-lookup"><span data-stu-id="99b1d-607">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="99b1d-608">Verfassen oder Lesen</span><span class="sxs-lookup"><span data-stu-id="99b1d-608">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="99b1d-609">Beispiel</span><span class="sxs-lookup"><span data-stu-id="99b1d-609">Example</span></span>

<span data-ttu-id="99b1d-610">Das folgende Beispiel ruft `makeEwsRequestAsync` zum Verwenden des `GetItem`-Vorgangs auf, um den Betreff eines Elements abzurufen.</span><span class="sxs-lookup"><span data-stu-id="99b1d-610">The following example calls `makeEwsRequestAsync` to use the `GetItem` operation to get the subject of an item.</span></span>

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