
# <a name="mailbox"></a><span data-ttu-id="9e68b-101">mailbox</span><span class="sxs-lookup"><span data-stu-id="9e68b-101">mailbox</span></span>

### <span data-ttu-id="9e68b-p101">[Office](Office.md)[.context](Office.context.md). mailbox</span><span class="sxs-lookup"><span data-stu-id="9e68b-p101">[Office](Office.md)[.context](Office.context.md). mailbox</span></span>

<span data-ttu-id="9e68b-104">Ermöglicht den Zugriff auf das Objektmodell von Outlook-add-in für Microsoft Outlook und Microsoft Outlook im Web.</span><span class="sxs-lookup"><span data-stu-id="9e68b-104">Provides access to the Outlook add-in object model for Microsoft Outlook and Microsoft Outlook on the web.</span></span>

##### <a name="requirements"></a><span data-ttu-id="9e68b-105">Anforderungen</span><span class="sxs-lookup"><span data-stu-id="9e68b-105">Requirements</span></span>

|<span data-ttu-id="9e68b-106">Anforderung</span><span class="sxs-lookup"><span data-stu-id="9e68b-106">Requirement</span></span>| <span data-ttu-id="9e68b-107">Wert</span><span class="sxs-lookup"><span data-stu-id="9e68b-107">Value</span></span>|
|---|---|
|[<span data-ttu-id="9e68b-108">Mindestversion des Postfachanforderungssatzes</span><span class="sxs-lookup"><span data-stu-id="9e68b-108">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="9e68b-109">1.0</span><span class="sxs-lookup"><span data-stu-id="9e68b-109">1.0</span></span>|
|[<span data-ttu-id="9e68b-110">Mindestberechtigungsstufe</span><span class="sxs-lookup"><span data-stu-id="9e68b-110">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="9e68b-111">Eingeschränkt</span><span class="sxs-lookup"><span data-stu-id="9e68b-111">Restricted</span></span>|
|[<span data-ttu-id="9e68b-112">Zutreffender Outlook-Modus</span><span class="sxs-lookup"><span data-stu-id="9e68b-112">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="9e68b-113">Verfassen oder Lesen</span><span class="sxs-lookup"><span data-stu-id="9e68b-113">Compose or read</span></span>|

##### <a name="members-and-methods"></a><span data-ttu-id="9e68b-114">Elemente und Methoden</span><span class="sxs-lookup"><span data-stu-id="9e68b-114">Members and methods</span></span>

| <span data-ttu-id="9e68b-115">Element</span><span class="sxs-lookup"><span data-stu-id="9e68b-115">Member</span></span> | <span data-ttu-id="9e68b-116">Typ</span><span class="sxs-lookup"><span data-stu-id="9e68b-116">Type</span></span> |
|--------|------|
| [<span data-ttu-id="9e68b-117">ewsUrl</span><span class="sxs-lookup"><span data-stu-id="9e68b-117">ewsUrl</span></span>](#ewsurl-string) | <span data-ttu-id="9e68b-118">Element</span><span class="sxs-lookup"><span data-stu-id="9e68b-118">Member</span></span> |
| [<span data-ttu-id="9e68b-119">restUrl</span><span class="sxs-lookup"><span data-stu-id="9e68b-119">restUrl</span></span>](#resturl-string) | <span data-ttu-id="9e68b-120">Element</span><span class="sxs-lookup"><span data-stu-id="9e68b-120">Member</span></span> |
| [<span data-ttu-id="9e68b-121">addHandlerAsync</span><span class="sxs-lookup"><span data-stu-id="9e68b-121">addHandlerAsync</span></span>](#addhandlerasynceventtype-handler-options-callback) | <span data-ttu-id="9e68b-122">Methode</span><span class="sxs-lookup"><span data-stu-id="9e68b-122">Method</span></span> |
| [<span data-ttu-id="9e68b-123">convertToEwsId</span><span class="sxs-lookup"><span data-stu-id="9e68b-123">convertToEwsId</span></span>](#converttoewsiditemid-restversion--string) | <span data-ttu-id="9e68b-124">Methode</span><span class="sxs-lookup"><span data-stu-id="9e68b-124">Method</span></span> |
| [<span data-ttu-id="9e68b-125">convertToLocalClientTime</span><span class="sxs-lookup"><span data-stu-id="9e68b-125">convertToLocalClientTime</span></span>](#converttolocalclienttimetimevalue--localclienttimejavascriptapioutlook15officelocalclienttime) | <span data-ttu-id="9e68b-126">Methode</span><span class="sxs-lookup"><span data-stu-id="9e68b-126">Method</span></span> |
| [<span data-ttu-id="9e68b-127">convertToRestId</span><span class="sxs-lookup"><span data-stu-id="9e68b-127">convertToRestId</span></span>](#converttorestiditemid-restversion--string) | <span data-ttu-id="9e68b-128">Methode</span><span class="sxs-lookup"><span data-stu-id="9e68b-128">Method</span></span> |
| [<span data-ttu-id="9e68b-129">convertToUtcClientTime</span><span class="sxs-lookup"><span data-stu-id="9e68b-129">convertToUtcClientTime</span></span>](#converttoutcclienttimeinput--date) | <span data-ttu-id="9e68b-130">Methode</span><span class="sxs-lookup"><span data-stu-id="9e68b-130">Method</span></span> |
| [<span data-ttu-id="9e68b-131">displayAppointmentForm</span><span class="sxs-lookup"><span data-stu-id="9e68b-131">displayAppointmentForm</span></span>](#displayappointmentformitemid) | <span data-ttu-id="9e68b-132">Methode</span><span class="sxs-lookup"><span data-stu-id="9e68b-132">Method</span></span> |
| [<span data-ttu-id="9e68b-133">displayMessageForm</span><span class="sxs-lookup"><span data-stu-id="9e68b-133">displayMessageForm</span></span>](#displaymessageformitemid) | <span data-ttu-id="9e68b-134">Methode</span><span class="sxs-lookup"><span data-stu-id="9e68b-134">Method</span></span> |
| [<span data-ttu-id="9e68b-135">displayNewAppointmentForm</span><span class="sxs-lookup"><span data-stu-id="9e68b-135">displayNewAppointmentForm</span></span>](#displaynewappointmentformparameters) | <span data-ttu-id="9e68b-136">Methode</span><span class="sxs-lookup"><span data-stu-id="9e68b-136">Method</span></span> |
| [<span data-ttu-id="9e68b-137">getCallbackTokenAsync</span><span class="sxs-lookup"><span data-stu-id="9e68b-137">getCallbackTokenAsync</span></span>](#getcallbacktokenasyncoptions-callback) | <span data-ttu-id="9e68b-138">Methode</span><span class="sxs-lookup"><span data-stu-id="9e68b-138">Method</span></span> |
| [<span data-ttu-id="9e68b-139">getCallbackTokenAsync</span><span class="sxs-lookup"><span data-stu-id="9e68b-139">getCallbackTokenAsync</span></span>](#getcallbacktokenasynccallback-usercontext) | <span data-ttu-id="9e68b-140">Methode</span><span class="sxs-lookup"><span data-stu-id="9e68b-140">Method</span></span> |
| [<span data-ttu-id="9e68b-141">getUserIdentityTokenAsync</span><span class="sxs-lookup"><span data-stu-id="9e68b-141">getUserIdentityTokenAsync</span></span>](#getuseridentitytokenasynccallback-usercontext) | <span data-ttu-id="9e68b-142">Methode</span><span class="sxs-lookup"><span data-stu-id="9e68b-142">Method</span></span> |
| [<span data-ttu-id="9e68b-143">makeEwsRequestAsync</span><span class="sxs-lookup"><span data-stu-id="9e68b-143">makeEwsRequestAsync</span></span>](#makeewsrequestasyncdata-callback-usercontext) | <span data-ttu-id="9e68b-144">Methode</span><span class="sxs-lookup"><span data-stu-id="9e68b-144">Method</span></span> |

### <a name="namespaces"></a><span data-ttu-id="9e68b-145">Namespaces</span><span class="sxs-lookup"><span data-stu-id="9e68b-145">Namespaces</span></span>

<span data-ttu-id="9e68b-146">[diagnostics](Office.context.mailbox.diagnostics.md): Stellt einem Outlook-Add-In Diagnoseinformationen bereit.</span><span class="sxs-lookup"><span data-stu-id="9e68b-146">[diagnostics](Office.context.mailbox.diagnostics.md): Provides diagnostic information to an Outlook add-in.</span></span>

<span data-ttu-id="9e68b-147">[item](Office.context.mailbox.item.md): Stellt Methoden und Eigenschaften für den Zugriff auf eine Nachricht oder einen Termins in einem Outlook-Add-In bereit.</span><span class="sxs-lookup"><span data-stu-id="9e68b-147">[item](Office.context.mailbox.item.md): Provides methods and properties for accessing a message or appointment in an Outlook add-in.</span></span>

<span data-ttu-id="9e68b-148">[userProfile](Office.context.mailbox.userProfile.md): Stellt Informationen zum Benutzer in einem Outlook-Add-In bereit.</span><span class="sxs-lookup"><span data-stu-id="9e68b-148">[userProfile](Office.context.mailbox.userProfile.md): Provides information about the user in an Outlook add-in.</span></span>

### <a name="members"></a><span data-ttu-id="9e68b-149">Elemente</span><span class="sxs-lookup"><span data-stu-id="9e68b-149">Members</span></span>

#### <a name="ewsurl-string"></a><span data-ttu-id="9e68b-150">ewsUrl :String</span><span class="sxs-lookup"><span data-stu-id="9e68b-150">ewsUrl :String</span></span>

<span data-ttu-id="9e68b-p102">Ruft die URL des EWS-Endpunkts (Exchange Web Services) für dieses E-Mail-Konto ab. Nur Lesemodus.</span><span class="sxs-lookup"><span data-stu-id="9e68b-p102">Gets the URL of the Exchange Web Services (EWS) endpoint for this email account. Read mode only.</span></span>

> [!NOTE]
> <span data-ttu-id="9e68b-153">Dieser Member wird in Outlook für iOS oder Outlook für Android nicht unterstützt.</span><span class="sxs-lookup"><span data-stu-id="9e68b-153">This member is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="9e68b-p103">Der `ewsUrl`-Wert kann durch einen Remotedienst zum Ausführen von EWS-Aufrufen an das Postfach des Benutzers verwendet werden. Sie können z. B. einen Remotedienst erstellen, um [Anlagen](https://docs.microsoft.com/outlook/add-ins/get-attachments-of-an-outlook-item) aus dem ausgewählten Element abzurufen.</span><span class="sxs-lookup"><span data-stu-id="9e68b-p103">The `ewsUrl` value can be used by a remote service to make EWS calls to the user's mailbox. For example, you can create a remote service to [get attachments from the selected item](https://docs.microsoft.com/outlook/add-ins/get-attachments-of-an-outlook-item).</span></span>

<span data-ttu-id="9e68b-156">Ihre App muss über die im Manifest angegebene **ReadItem**-Berechtigung verfügen, um das `ewsUrl`-Element im Lesemodus aufzurufen.</span><span class="sxs-lookup"><span data-stu-id="9e68b-156">Your app must have the **ReadItem** permission specified in its manifest to call the `ewsUrl` member in read mode.</span></span>

<span data-ttu-id="9e68b-p104">Im Verfassenmodus müssen Sie die [`saveAsync`](Office.context.mailbox.item.md#saveasyncoptions-callback)-Methode aufrufen, bevor Sie das `ewsUrl`-Element verwenden können. Die App muss über **ReadWriteItem**-Berechtigungen verfügen, um die `saveAsync`-Methode aufzurufen.</span><span class="sxs-lookup"><span data-stu-id="9e68b-p104">In compose mode you must call the [`saveAsync`](Office.context.mailbox.item.md#saveasyncoptions-callback) method before you can use the `ewsUrl` member. Your app must have **ReadWriteItem** permissions to call the `saveAsync` method.</span></span>

##### <a name="type"></a><span data-ttu-id="9e68b-159">Typ:</span><span class="sxs-lookup"><span data-stu-id="9e68b-159">Type:</span></span>

*   <span data-ttu-id="9e68b-160">String</span><span class="sxs-lookup"><span data-stu-id="9e68b-160">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="9e68b-161">Anforderungen</span><span class="sxs-lookup"><span data-stu-id="9e68b-161">Requirements</span></span>

|<span data-ttu-id="9e68b-162">Anforderung</span><span class="sxs-lookup"><span data-stu-id="9e68b-162">Requirement</span></span>| <span data-ttu-id="9e68b-163">Wert</span><span class="sxs-lookup"><span data-stu-id="9e68b-163">Value</span></span>|
|---|---|
|[<span data-ttu-id="9e68b-164">Mindestversion des Postfachanforderungssatzes</span><span class="sxs-lookup"><span data-stu-id="9e68b-164">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="9e68b-165">1.0</span><span class="sxs-lookup"><span data-stu-id="9e68b-165">1.0</span></span>|
|[<span data-ttu-id="9e68b-166">Mindestberechtigungsstufe</span><span class="sxs-lookup"><span data-stu-id="9e68b-166">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="9e68b-167">ReadItem</span><span class="sxs-lookup"><span data-stu-id="9e68b-167">ReadItem</span></span>|
|[<span data-ttu-id="9e68b-168">Zutreffender Outlook-Modus</span><span class="sxs-lookup"><span data-stu-id="9e68b-168">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="9e68b-169">Verfassen oder Lesen</span><span class="sxs-lookup"><span data-stu-id="9e68b-169">Compose or read</span></span>|

#### <a name="resturl-string"></a><span data-ttu-id="9e68b-170">restUrl :String</span><span class="sxs-lookup"><span data-stu-id="9e68b-170">restUrl :String</span></span>

<span data-ttu-id="9e68b-171">Ruft die URL des REST-Endpunkts für das betreffende E-Mail-Konto ab.</span><span class="sxs-lookup"><span data-stu-id="9e68b-171">Gets the URL of the REST endpoint for this email account.</span></span>

<span data-ttu-id="9e68b-172">Der Wert `restUrl` kann für [REST-API](https://docs.microsoft.com/outlook/rest/)-Aufrufe an das Postfach des Benutzers verwendet werden.</span><span class="sxs-lookup"><span data-stu-id="9e68b-172">The `restUrl` value can be used to make [REST API](https://docs.microsoft.com/outlook/rest/) calls to the user's mailbox.</span></span>

<span data-ttu-id="9e68b-173">Im Manifest der App muss die Berechtigung **ReadItem** angegeben sein, damit das Mitglied `restUrl` im Lesemodus abgerufen werden kann.</span><span class="sxs-lookup"><span data-stu-id="9e68b-173">Your app must have the **ReadItem** permission specified in its manifest to call the `restUrl` member in read mode.</span></span>

<span data-ttu-id="9e68b-p105">Im Verfassenmodus müssen Sie die [`saveAsync`](Office.context.mailbox.item.md#saveasyncoptions-callback)-Methode aufrufen, bevor Sie das `restUrl`-Element verwenden können. Die App muss über **ReadWriteItem**-Berechtigungen verfügen, um die `saveAsync`-Methode aufzurufen.</span><span class="sxs-lookup"><span data-stu-id="9e68b-p105">In compose mode you must call the [`saveAsync`](Office.context.mailbox.item.md#saveasyncoptions-callback) method before you can use the `restUrl` member. Your app must have **ReadWriteItem** permissions to call the `saveAsync` method.</span></span>

> [!NOTE]
> <span data-ttu-id="9e68b-176">Outlook-Clients für lokale Installationen von Exchange 2016 mit einer benutzerdefinierten REST-URL konfiguriert verbunden werden einen ungültigen Wert für zurückgeben `restUrl`.</span><span class="sxs-lookup"><span data-stu-id="9e68b-176">Outlook clients connected to on-premises installations of Exchange 2016 with a custom REST URL configured will return an invalid value for `restUrl`.</span></span>

##### <a name="type"></a><span data-ttu-id="9e68b-177">Typ:</span><span class="sxs-lookup"><span data-stu-id="9e68b-177">Type:</span></span>

*   <span data-ttu-id="9e68b-178">String</span><span class="sxs-lookup"><span data-stu-id="9e68b-178">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="9e68b-179">Anforderungen</span><span class="sxs-lookup"><span data-stu-id="9e68b-179">Requirements</span></span>

|<span data-ttu-id="9e68b-180">Anforderung</span><span class="sxs-lookup"><span data-stu-id="9e68b-180">Requirement</span></span>| <span data-ttu-id="9e68b-181">Wert</span><span class="sxs-lookup"><span data-stu-id="9e68b-181">Value</span></span>|
|---|---|
|[<span data-ttu-id="9e68b-182">Mindestversion des Postfachanforderungssatzes</span><span class="sxs-lookup"><span data-stu-id="9e68b-182">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="9e68b-183">1,5</span><span class="sxs-lookup"><span data-stu-id="9e68b-183">1.5</span></span> |
|[<span data-ttu-id="9e68b-184">Mindestberechtigungsstufe</span><span class="sxs-lookup"><span data-stu-id="9e68b-184">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="9e68b-185">ReadItem</span><span class="sxs-lookup"><span data-stu-id="9e68b-185">ReadItem</span></span>|
|[<span data-ttu-id="9e68b-186">Zutreffender Outlook-Modus</span><span class="sxs-lookup"><span data-stu-id="9e68b-186">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="9e68b-187">Verfassen oder Lesen</span><span class="sxs-lookup"><span data-stu-id="9e68b-187">Compose or read</span></span>|

### <a name="methods"></a><span data-ttu-id="9e68b-188">Methoden</span><span class="sxs-lookup"><span data-stu-id="9e68b-188">Methods</span></span>

####  <a name="addhandlerasynceventtype-handler-options-callback"></a><span data-ttu-id="9e68b-189">addHandlerAsync(eventType, handler, [options], [callback])</span><span class="sxs-lookup"><span data-stu-id="9e68b-189">addHandlerAsync(eventType, handler, [options], [callback])</span></span>

<span data-ttu-id="9e68b-190">Fügt einen Ereignishandler für ein unterstütztes Ereignis hinzu.</span><span class="sxs-lookup"><span data-stu-id="9e68b-190">Adds an event handler for a supported event.</span></span>

<span data-ttu-id="9e68b-p106">Derzeit wird als Ereignistyp ausschließlich `Office.EventType.ItemChanged` unterstützt. Er wird aufgerufen, wenn der Benutzer ein neues Element auswählt. Dieses Ereignis wird von Add-Ins verwendet, die einen anheftbaren Aufgabenbereich implementieren. Es erlaubt es dem Add-In, die Benutzeroberfläche des Aufgabenbereichs basierend auf dem aktuell ausgewählten Element zu aktualisieren.</span><span class="sxs-lookup"><span data-stu-id="9e68b-p106">Currently the only supported event type is `Office.EventType.ItemChanged`, which is invoked when the user selects a new item. This event is used by add-ins that implement a pinnable taskpane, and allows the add-in to refresh the taskpane UI based on the currently selected item.</span></span>

##### <a name="parameters"></a><span data-ttu-id="9e68b-193">Parameter:</span><span class="sxs-lookup"><span data-stu-id="9e68b-193">Parameters:</span></span>

| <span data-ttu-id="9e68b-194">Name</span><span class="sxs-lookup"><span data-stu-id="9e68b-194">Name</span></span> | <span data-ttu-id="9e68b-195">Typ</span><span class="sxs-lookup"><span data-stu-id="9e68b-195">Type</span></span> | <span data-ttu-id="9e68b-196">Attribute</span><span class="sxs-lookup"><span data-stu-id="9e68b-196">Attributes</span></span> | <span data-ttu-id="9e68b-197">Beschreibung</span><span class="sxs-lookup"><span data-stu-id="9e68b-197">Description</span></span> |
|---|---|---|---|
| `eventType` | [<span data-ttu-id="9e68b-198">Office.EventType</span><span class="sxs-lookup"><span data-stu-id="9e68b-198">Office.EventType</span></span>](office.md#eventtype-string) || <span data-ttu-id="9e68b-199">Das Ereignis, das den Handler aufrufen soll</span><span class="sxs-lookup"><span data-stu-id="9e68b-199">The event that should invoke the handler.</span></span> |
| `handler` | <span data-ttu-id="9e68b-200">Function</span><span class="sxs-lookup"><span data-stu-id="9e68b-200">Function</span></span> || <span data-ttu-id="9e68b-p107">Die Funktion, die das Ereignis behandeln soll. Die Funktion muss einen einzigen Parameter akzeptieren (ein Objektliteral). Die Eigenschaft `type` dieses Parameters entspricht dem Parameter `eventType`, der an `addHandlerAsync` übergeben wird.</span><span class="sxs-lookup"><span data-stu-id="9e68b-p107">The function to handle the event. The function must accept a single parameter, which is an object literal. The `type` property on the parameter will match the `eventType` parameter passed to `addHandlerAsync`.</span></span> |
| `options` | <span data-ttu-id="9e68b-204">Objekt</span><span class="sxs-lookup"><span data-stu-id="9e68b-204">Object</span></span> | <span data-ttu-id="9e68b-205">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="9e68b-205">&lt;optional&gt;</span></span> | <span data-ttu-id="9e68b-206">Ein Objektliteral, das mindestens eine der folgenden Eigenschaften enthält.</span><span class="sxs-lookup"><span data-stu-id="9e68b-206">An object literal that contains one or more of the following properties.</span></span> |
| `options.asyncContext` | <span data-ttu-id="9e68b-207">Object</span><span class="sxs-lookup"><span data-stu-id="9e68b-207">Object</span></span> | <span data-ttu-id="9e68b-208">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="9e68b-208">&lt;optional&gt;</span></span> | <span data-ttu-id="9e68b-209">Entwickler können ein Objekt bereitstellen, auf das sie in der Callbackmethode zugreifen möchten.</span><span class="sxs-lookup"><span data-stu-id="9e68b-209">Developers can provide any object they wish to access in the callback method.</span></span> |
| `callback` | <span data-ttu-id="9e68b-210">Funktion</span><span class="sxs-lookup"><span data-stu-id="9e68b-210">function</span></span>| <span data-ttu-id="9e68b-211">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="9e68b-211">&lt;optional&gt;</span></span>|<span data-ttu-id="9e68b-212">Wenn die Methode abgeschlossen wird, wird die Funktion , die im `callback`-Parameter übergeben wird, mit einem einzelnen Parameter (`asyncResult`) aufgerufen, der ein [`AsyncResult`](/javascript/api/office/office.asyncresult)-Objekt darstellt.</span><span class="sxs-lookup"><span data-stu-id="9e68b-212">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="9e68b-213">Anforderungen</span><span class="sxs-lookup"><span data-stu-id="9e68b-213">Requirements</span></span>

|<span data-ttu-id="9e68b-214">Anforderung</span><span class="sxs-lookup"><span data-stu-id="9e68b-214">Requirement</span></span>| <span data-ttu-id="9e68b-215">Wert</span><span class="sxs-lookup"><span data-stu-id="9e68b-215">Value</span></span>|
|---|---|
|[<span data-ttu-id="9e68b-216">Mindestversion des Postfachanforderungssatzes</span><span class="sxs-lookup"><span data-stu-id="9e68b-216">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="9e68b-217">1,5</span><span class="sxs-lookup"><span data-stu-id="9e68b-217">1.5</span></span> |
|[<span data-ttu-id="9e68b-218">Mindestberechtigungsstufe</span><span class="sxs-lookup"><span data-stu-id="9e68b-218">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="9e68b-219">ReadItem</span><span class="sxs-lookup"><span data-stu-id="9e68b-219">ReadItem</span></span> |
|[<span data-ttu-id="9e68b-220">Zutreffender Outlook-Modus</span><span class="sxs-lookup"><span data-stu-id="9e68b-220">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="9e68b-221">Verfassen oder Lesen</span><span class="sxs-lookup"><span data-stu-id="9e68b-221">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="9e68b-222">Beispiel</span><span class="sxs-lookup"><span data-stu-id="9e68b-222">Example</span></span>

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

####  <a name="converttoewsiditemid-restversion--string"></a><span data-ttu-id="9e68b-223">convertToEwsId(itemId, restVersion) → {String}</span><span class="sxs-lookup"><span data-stu-id="9e68b-223">convertToEwsId(itemId, restVersion) → {String}</span></span>

<span data-ttu-id="9e68b-224">Wandelt eine für REST formatierte Element-ID ins EWS-Format um.</span><span class="sxs-lookup"><span data-stu-id="9e68b-224">Converts an item ID formatted for REST into EWS format.</span></span>

> [!NOTE]
> <span data-ttu-id="9e68b-225">Diese Methode wird in Outlook für iOS oder Outlook für Android nicht unterstützt.</span><span class="sxs-lookup"><span data-stu-id="9e68b-225">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="9e68b-p108">Über eine REST-API abgerufene Element-IDs (z. B. die [Outlook Mail-API](https://docs.microsoft.com/previous-versions/office/office-365-api/api/version-2.0/mail-rest-operations) oder [Microsoft Graph](http://graph.microsoft.io/)) verwenden ein anderes Format, als das von Exchange-Webdiensten (EWS) verwendete. Die `convertToEwsId`-Methode wandelt eine für REST formatierte ID in das entsprechende EWS-Format um.</span><span class="sxs-lookup"><span data-stu-id="9e68b-p108">Item IDs retrieved via a REST API (such as the [Outlook Mail API](https://docs.microsoft.com/previous-versions/office/office-365-api/api/version-2.0/mail-rest-operations) or the [Microsoft Graph](http://graph.microsoft.io/)) use a different format than the format used by Exchange Web Services (EWS). The `convertToEwsId` method converts a REST-formatted ID into the proper format for EWS.</span></span>

##### <a name="parameters"></a><span data-ttu-id="9e68b-228">Parameter:</span><span class="sxs-lookup"><span data-stu-id="9e68b-228">Parameters:</span></span>

|<span data-ttu-id="9e68b-229">Name</span><span class="sxs-lookup"><span data-stu-id="9e68b-229">Name</span></span>| <span data-ttu-id="9e68b-230">Typ</span><span class="sxs-lookup"><span data-stu-id="9e68b-230">Type</span></span>| <span data-ttu-id="9e68b-231">Beschreibung</span><span class="sxs-lookup"><span data-stu-id="9e68b-231">Description</span></span>|
|---|---|---|
|`itemId`| <span data-ttu-id="9e68b-232">Zeichenfolge</span><span class="sxs-lookup"><span data-stu-id="9e68b-232">String</span></span>|<span data-ttu-id="9e68b-233">Eine für Outlook-REST-APIs formatierte Element-ID</span><span class="sxs-lookup"><span data-stu-id="9e68b-233">An item ID formatted for the Outlook REST APIs</span></span>|
|`restVersion`| [<span data-ttu-id="9e68b-234">Office.MailboxEnums.RestVersion</span><span class="sxs-lookup"><span data-stu-id="9e68b-234">Office.MailboxEnums.RestVersion</span></span>](/javascript/api/outlook_1_5/office.mailboxenums.restversion)|<span data-ttu-id="9e68b-235">Ein Wert, der die Version der Outlook-REST-API angibt, die zum Abrufen der Element-ID verwendet wurde.</span><span class="sxs-lookup"><span data-stu-id="9e68b-235">A value indicating the version of the Outlook REST API used to retrieve the item ID.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="9e68b-236">Anforderungen</span><span class="sxs-lookup"><span data-stu-id="9e68b-236">Requirements</span></span>

|<span data-ttu-id="9e68b-237">Anforderung</span><span class="sxs-lookup"><span data-stu-id="9e68b-237">Requirement</span></span>| <span data-ttu-id="9e68b-238">Wert</span><span class="sxs-lookup"><span data-stu-id="9e68b-238">Value</span></span>|
|---|---|
|[<span data-ttu-id="9e68b-239">Mindestversion des Postfachanforderungssatzes</span><span class="sxs-lookup"><span data-stu-id="9e68b-239">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="9e68b-240">1.3</span><span class="sxs-lookup"><span data-stu-id="9e68b-240">1.3</span></span>|
|[<span data-ttu-id="9e68b-241">Mindestberechtigungsstufe</span><span class="sxs-lookup"><span data-stu-id="9e68b-241">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="9e68b-242">Eingeschränkt</span><span class="sxs-lookup"><span data-stu-id="9e68b-242">Restricted</span></span>|
|[<span data-ttu-id="9e68b-243">Zutreffender Outlook-Modus</span><span class="sxs-lookup"><span data-stu-id="9e68b-243">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="9e68b-244">Verfassen oder Lesen</span><span class="sxs-lookup"><span data-stu-id="9e68b-244">Compose or read</span></span>|

##### <a name="returns"></a><span data-ttu-id="9e68b-245">Gibt zurück:</span><span class="sxs-lookup"><span data-stu-id="9e68b-245">Returns:</span></span>

<span data-ttu-id="9e68b-246">Typ: Zeichenfolge</span><span class="sxs-lookup"><span data-stu-id="9e68b-246">Type: String</span></span>

##### <a name="example"></a><span data-ttu-id="9e68b-247">Beispiel</span><span class="sxs-lookup"><span data-stu-id="9e68b-247">Example</span></span>

```
// Get an item's ID from a REST API
var restId = 'AAMkAGVlOTZjNTM3LW...';

// Treat restId as coming from the v2.0 version of the
// Outlook Mail API
var ewsId = Office.context.mailbox.convertToEwsId(restId, Office.MailboxEnums.RestVersion.v2_0);
```

####  <a name="converttolocalclienttimetimevalue--localclienttimejavascriptapioutlook15officelocalclienttime"></a><span data-ttu-id="9e68b-248">convertToLocalClientTime(timeValue) → {[LocalClientTime](/javascript/api/outlook_1_5/office.LocalClientTime)}</span><span class="sxs-lookup"><span data-stu-id="9e68b-248">convertToLocalClientTime(timeValue) → {[LocalClientTime](/javascript/api/outlook_1_5/office.LocalClientTime)}</span></span>

<span data-ttu-id="9e68b-249">Ruft ein Wörterbuch mit Uhrzeitinformationen basierend auf der Zeiteinstellung des lokalen Clients ab.</span><span class="sxs-lookup"><span data-stu-id="9e68b-249">Gets a dictionary containing time information in local client time.</span></span>

<span data-ttu-id="9e68b-p109">Die in einer Mail-App für Outlook oder Outlook Web App verwendeten Daten und Uhrzeiten stammen u. U. aus verschiedenen Zeitzonen. Outlook verwendet die Zeitzone des Client-Computers; Outlook Web App verwendet die im Exchange Admin Center (EAC) festgelegte Zeitzone. Sie sollten Datums- und Uhrzeitwerte bearbeiten, damit die auf der Benutzeroberfläche angezeigten Werte immer den von Benutzer erwarteten Zeitzonen entsprechen.</span><span class="sxs-lookup"><span data-stu-id="9e68b-p109">The dates and times used by a mail app for Outlook or Outlook Web App can use different time zones. Outlook uses the client computer time zone; Outlook Web App uses the time zone set on the Exchange Admin Center (EAC). You should handle date and time values so that the values you display on the user interface are always consistent with the time zone that the user expects.</span></span>

<span data-ttu-id="9e68b-p110">Wird die Mail-App in Outlook ausgeführt, wird mit der `convertToLocalClientTime`-Methode ein Wörterbuchobjekt zurückgegeben, dessen Werte auf die Zeitzone des Clientcomputers festgelegt sind. Wird die Mail-App in Outlook ausgeführt, wird mit der `convertToLocalClientTime`-Methode ein Wörterbuchobjekt zurückgegeben, dessen Werte auf die Zeitzone des Clientcomputers festgelegt sind.</span><span class="sxs-lookup"><span data-stu-id="9e68b-p110">If the mail app is running in Outlook, the `convertToLocalClientTime` method will return a dictionary object with the values set to the client computer time zone. If the mail app is running in Outlook Web App, the `convertToLocalClientTime` method will return a dictionary object with the values set to the time zone specified in the EAC.</span></span>

##### <a name="parameters"></a><span data-ttu-id="9e68b-255">Parameter:</span><span class="sxs-lookup"><span data-stu-id="9e68b-255">Parameters:</span></span>

|<span data-ttu-id="9e68b-256">Name</span><span class="sxs-lookup"><span data-stu-id="9e68b-256">Name</span></span>| <span data-ttu-id="9e68b-257">Typ</span><span class="sxs-lookup"><span data-stu-id="9e68b-257">Type</span></span>| <span data-ttu-id="9e68b-258">Beschreibung</span><span class="sxs-lookup"><span data-stu-id="9e68b-258">Description</span></span>|
|---|---|---|
|`timeValue`| <span data-ttu-id="9e68b-259">Datum</span><span class="sxs-lookup"><span data-stu-id="9e68b-259">Date</span></span>|<span data-ttu-id="9e68b-260">Ein Date-Objekt</span><span class="sxs-lookup"><span data-stu-id="9e68b-260">A Date object</span></span>|

##### <a name="requirements"></a><span data-ttu-id="9e68b-261">Anforderungen</span><span class="sxs-lookup"><span data-stu-id="9e68b-261">Requirements</span></span>

|<span data-ttu-id="9e68b-262">Anforderung</span><span class="sxs-lookup"><span data-stu-id="9e68b-262">Requirement</span></span>| <span data-ttu-id="9e68b-263">Wert</span><span class="sxs-lookup"><span data-stu-id="9e68b-263">Value</span></span>|
|---|---|
|[<span data-ttu-id="9e68b-264">Mindestversion des Postfachanforderungssatzes</span><span class="sxs-lookup"><span data-stu-id="9e68b-264">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="9e68b-265">1.0</span><span class="sxs-lookup"><span data-stu-id="9e68b-265">1.0</span></span>|
|[<span data-ttu-id="9e68b-266">Mindestberechtigungsstufe</span><span class="sxs-lookup"><span data-stu-id="9e68b-266">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="9e68b-267">ReadItem</span><span class="sxs-lookup"><span data-stu-id="9e68b-267">ReadItem</span></span>|
|[<span data-ttu-id="9e68b-268">Zutreffender Outlook-Modus</span><span class="sxs-lookup"><span data-stu-id="9e68b-268">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="9e68b-269">Verfassen oder Lesen</span><span class="sxs-lookup"><span data-stu-id="9e68b-269">Compose or read</span></span>|

##### <a name="returns"></a><span data-ttu-id="9e68b-270">Gibt zurück:</span><span class="sxs-lookup"><span data-stu-id="9e68b-270">Returns:</span></span>

<span data-ttu-id="9e68b-271">Typ: [LocalClientTime](/javascript/api/outlook_1_5/office.LocalClientTime)</span><span class="sxs-lookup"><span data-stu-id="9e68b-271">Type: [LocalClientTime](/javascript/api/outlook_1_5/office.LocalClientTime)</span></span>

####  <a name="converttorestiditemid-restversion--string"></a><span data-ttu-id="9e68b-272">convertToRestId(itemId, restVersion) → {String}</span><span class="sxs-lookup"><span data-stu-id="9e68b-272">convertToRestId(itemId, restVersion) → {String}</span></span>

<span data-ttu-id="9e68b-273">Wandelt eine für EWS formatierte Element-ID ins REST-Format um.</span><span class="sxs-lookup"><span data-stu-id="9e68b-273">Converts an item ID formatted for EWS into REST format.</span></span>

> [!NOTE]
> <span data-ttu-id="9e68b-274">Diese Methode wird in Outlook für iOS oder Outlook für Android nicht unterstützt.</span><span class="sxs-lookup"><span data-stu-id="9e68b-274">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="9e68b-p111">Über EWS oder die `itemId` abgerufene Element-IDs verwenden ein anderes Format als die über REST-APIs (z. B. die [Outlook Mail-API](https://docs.microsoft.com/previous-versions/office/office-365-api/api/version-2.0/mail-rest-operations) oder [Microsoft Graph](http://graph.microsoft.io/)) abgerufenen Element-IDs. Die `convertToRestId`-Methode wandelt eine für EWS formatierte ID in das entsprechende REST-Format um.</span><span class="sxs-lookup"><span data-stu-id="9e68b-p111">Item IDs retrieved via EWS or via the `itemId` property use a different format than the format used by REST APIs (such as the [Outlook Mail API](https://docs.microsoft.com/previous-versions/office/office-365-api/api/version-2.0/mail-rest-operations) or the [Microsoft Graph](http://graph.microsoft.io/)). The `convertToRestId` method converts an EWS-formatted ID into the proper format for REST.</span></span>

##### <a name="parameters"></a><span data-ttu-id="9e68b-277">Parameter:</span><span class="sxs-lookup"><span data-stu-id="9e68b-277">Parameters:</span></span>

|<span data-ttu-id="9e68b-278">Name</span><span class="sxs-lookup"><span data-stu-id="9e68b-278">Name</span></span>| <span data-ttu-id="9e68b-279">Typ</span><span class="sxs-lookup"><span data-stu-id="9e68b-279">Type</span></span>| <span data-ttu-id="9e68b-280">Beschreibung</span><span class="sxs-lookup"><span data-stu-id="9e68b-280">Description</span></span>|
|---|---|---|
|`itemId`| <span data-ttu-id="9e68b-281">Zeichenfolge</span><span class="sxs-lookup"><span data-stu-id="9e68b-281">String</span></span>|<span data-ttu-id="9e68b-282">Eine für Exchange-Webdienste (EWS) formatierte Element-ID</span><span class="sxs-lookup"><span data-stu-id="9e68b-282">An item ID formatted for Exchange Web Services (EWS)</span></span>|
|`restVersion`| [<span data-ttu-id="9e68b-283">Office.MailboxEnums.RestVersion</span><span class="sxs-lookup"><span data-stu-id="9e68b-283">Office.MailboxEnums.RestVersion</span></span>](/javascript/api/outlook_1_5/office.mailboxenums.restversion)|<span data-ttu-id="9e68b-284">Ein Wert, der die Version der Outlook-REST-API angibt, mit der die konvertierte ID verwendet wird.</span><span class="sxs-lookup"><span data-stu-id="9e68b-284">A value indicating the version of the Outlook REST API that the converted ID will be used with.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="9e68b-285">Anforderungen</span><span class="sxs-lookup"><span data-stu-id="9e68b-285">Requirements</span></span>

|<span data-ttu-id="9e68b-286">Anforderung</span><span class="sxs-lookup"><span data-stu-id="9e68b-286">Requirement</span></span>| <span data-ttu-id="9e68b-287">Wert</span><span class="sxs-lookup"><span data-stu-id="9e68b-287">Value</span></span>|
|---|---|
|[<span data-ttu-id="9e68b-288">Mindestversion des Postfachanforderungssatzes</span><span class="sxs-lookup"><span data-stu-id="9e68b-288">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="9e68b-289">1.3</span><span class="sxs-lookup"><span data-stu-id="9e68b-289">1.3</span></span>|
|[<span data-ttu-id="9e68b-290">Mindestberechtigungsstufe</span><span class="sxs-lookup"><span data-stu-id="9e68b-290">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="9e68b-291">Eingeschränkt</span><span class="sxs-lookup"><span data-stu-id="9e68b-291">Restricted</span></span>|
|[<span data-ttu-id="9e68b-292">Zutreffender Outlook-Modus</span><span class="sxs-lookup"><span data-stu-id="9e68b-292">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="9e68b-293">Verfassen oder Lesen</span><span class="sxs-lookup"><span data-stu-id="9e68b-293">Compose or read</span></span>|

##### <a name="returns"></a><span data-ttu-id="9e68b-294">Gibt zurück:</span><span class="sxs-lookup"><span data-stu-id="9e68b-294">Returns:</span></span>

<span data-ttu-id="9e68b-295">Typ: Zeichenfolge</span><span class="sxs-lookup"><span data-stu-id="9e68b-295">Type: String</span></span>

##### <a name="example"></a><span data-ttu-id="9e68b-296">Beispiel</span><span class="sxs-lookup"><span data-stu-id="9e68b-296">Example</span></span>

```
// Get the currently selected item's ID
var ewsId = Office.context.mailbox.item.itemId;

// Convert to a REST ID for the v2.0 version of the
// Outlook Mail API
var restId = Office.context.mailbox.convertToRestId(ewsId, Office.MailboxEnums.RestVersion.v2_0);
```

####  <a name="converttoutcclienttimeinput--date"></a><span data-ttu-id="9e68b-297">convertToUtcClientTime(input) → {Date}</span><span class="sxs-lookup"><span data-stu-id="9e68b-297">convertToUtcClientTime(input) → {Date}</span></span>

<span data-ttu-id="9e68b-298">Ruft ein Date-Objekt aus einem Wörterbuch mit Uhrzeitinformationen ab.</span><span class="sxs-lookup"><span data-stu-id="9e68b-298">Gets a Date object from a dictionary containing time information.</span></span>

<span data-ttu-id="9e68b-299">Mit der `convertToUtcClientTime`-Methode wird ein Wörterbuch mit lokalem Datum und lokaler Uhrzeit in ein Date-Objekt mit den richtigen Werten für das lokale Datum und die lokale Uhrzeit konvertiert.</span><span class="sxs-lookup"><span data-stu-id="9e68b-299">The `convertToUtcClientTime` method converts a dictionary containing a local date and time to a Date object with the correct values for the local date and time.</span></span>

##### <a name="parameters"></a><span data-ttu-id="9e68b-300">Parameter:</span><span class="sxs-lookup"><span data-stu-id="9e68b-300">Parameters:</span></span>

|<span data-ttu-id="9e68b-301">Name</span><span class="sxs-lookup"><span data-stu-id="9e68b-301">Name</span></span>| <span data-ttu-id="9e68b-302">Typ</span><span class="sxs-lookup"><span data-stu-id="9e68b-302">Type</span></span>| <span data-ttu-id="9e68b-303">Beschreibung</span><span class="sxs-lookup"><span data-stu-id="9e68b-303">Description</span></span>|
|---|---|---|
|`input`| [<span data-ttu-id="9e68b-304">LocalClientTime</span><span class="sxs-lookup"><span data-stu-id="9e68b-304">LocalClientTime</span></span>](/javascript/api/outlook_1_5/office.LocalClientTime)|<span data-ttu-id="9e68b-305">Der zu konvertierende Wert für die lokale Uhrzeit.</span><span class="sxs-lookup"><span data-stu-id="9e68b-305">The local time value to convert.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="9e68b-306">Anforderungen</span><span class="sxs-lookup"><span data-stu-id="9e68b-306">Requirements</span></span>

|<span data-ttu-id="9e68b-307">Anforderung</span><span class="sxs-lookup"><span data-stu-id="9e68b-307">Requirement</span></span>| <span data-ttu-id="9e68b-308">Wert</span><span class="sxs-lookup"><span data-stu-id="9e68b-308">Value</span></span>|
|---|---|
|[<span data-ttu-id="9e68b-309">Mindestversion des Postfachanforderungssatzes</span><span class="sxs-lookup"><span data-stu-id="9e68b-309">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="9e68b-310">1.0</span><span class="sxs-lookup"><span data-stu-id="9e68b-310">1.0</span></span>|
|[<span data-ttu-id="9e68b-311">Mindestberechtigungsstufe</span><span class="sxs-lookup"><span data-stu-id="9e68b-311">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="9e68b-312">ReadItem</span><span class="sxs-lookup"><span data-stu-id="9e68b-312">ReadItem</span></span>|
|[<span data-ttu-id="9e68b-313">Zutreffender Outlook-Modus</span><span class="sxs-lookup"><span data-stu-id="9e68b-313">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="9e68b-314">Verfassen oder Lesen</span><span class="sxs-lookup"><span data-stu-id="9e68b-314">Compose or read</span></span>|

##### <a name="returns"></a><span data-ttu-id="9e68b-315">Gibt zurück:</span><span class="sxs-lookup"><span data-stu-id="9e68b-315">Returns:</span></span>

<span data-ttu-id="9e68b-316">Ein Date-Objekt der Uhrzeit in UTC.</span><span class="sxs-lookup"><span data-stu-id="9e68b-316">A Date object with the time expressed in UTC.</span></span>

<dl class="param-type"><span data-ttu-id="9e68b-317">

<dt>
Typ</dt>


</span><span class="sxs-lookup"><span data-stu-id="9e68b-317">

<dt>Type</dt>

</span></span><dd><span data-ttu-id="9e68b-318">Datum</span><span class="sxs-lookup"><span data-stu-id="9e68b-318">Date</span></span></dd>

</dl>

####  <a name="displayappointmentformitemid"></a><span data-ttu-id="9e68b-319">displayAppointmentForm(itemId)</span><span class="sxs-lookup"><span data-stu-id="9e68b-319">displayAppointmentForm(itemId)</span></span>

<span data-ttu-id="9e68b-320">Zeigt einen bestehenden Kalendertermin an.</span><span class="sxs-lookup"><span data-stu-id="9e68b-320">Displays an existing calendar appointment.</span></span>

> [!NOTE]
> <span data-ttu-id="9e68b-321">Diese Methode wird in Outlook für iOS oder Outlook für Android nicht unterstützt.</span><span class="sxs-lookup"><span data-stu-id="9e68b-321">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="9e68b-322">Mit der `displayAppointmentForm`-Methode wird ein vorhandener Kalendertermin auf dem Desktop in einem neuen Fenster oder auf Mobilgeräten in einem Dialogfeld geöffnet.</span><span class="sxs-lookup"><span data-stu-id="9e68b-322">The `displayAppointmentForm` method opens an existing calendar appointment in a new window on the desktop or in a dialog box on mobile devices.</span></span>

<span data-ttu-id="9e68b-p112">In Outlook for Mac können Sie diese Methode verwenden, um einen einzelnen Termin anzuzeigen, der nicht Teil einer Terminserie ist, oder den Mastertermin einer Terminserie, jedoch keine Instanz der Terminserie. Dies liegt daran, dass Sie in Outlook for Mac nicht auf die Eigenschaften (einschließlich der Element-ID) von Instanzen einer Terminserie zugreifen können.</span><span class="sxs-lookup"><span data-stu-id="9e68b-p112">In Outlook for Mac, you can use this method to display a single appointment that is not part of a recurring series, or the master appointment of a recurring series, but you cannot display an instance of the series. This is because in Outlook for Mac, you cannot access the properties (including the item ID) of instances of a recurring series.</span></span>

<span data-ttu-id="9e68b-325">In Outlook Web App öffnet diese Methode das angegebene Formular nur, wenn der Textkörper des Formulars nicht größer ist als 32 KB.</span><span class="sxs-lookup"><span data-stu-id="9e68b-325">In Outlook Web App, this method opens the specified form only if the body of the form is less than or equal to 32KB number of characters.</span></span>

<span data-ttu-id="9e68b-326">Wenn der angegebene Elementbezeichner keinen vorhandenen Termin identifiziert, wird auf dem Clientcomputer oder Gerät ein leerer Bereich geöffnet, und es wird keine Fehlermeldung zurückgegeben.</span><span class="sxs-lookup"><span data-stu-id="9e68b-326">If the specified item identifier does not identify an existing appointment, a blank pane opens on the client computer or device, and no error message will be returned.</span></span>

##### <a name="parameters"></a><span data-ttu-id="9e68b-327">Parameter:</span><span class="sxs-lookup"><span data-stu-id="9e68b-327">Parameters:</span></span>

|<span data-ttu-id="9e68b-328">Name</span><span class="sxs-lookup"><span data-stu-id="9e68b-328">Name</span></span>| <span data-ttu-id="9e68b-329">Typ</span><span class="sxs-lookup"><span data-stu-id="9e68b-329">Type</span></span>| <span data-ttu-id="9e68b-330">Beschreibung</span><span class="sxs-lookup"><span data-stu-id="9e68b-330">Description</span></span>|
|---|---|---|
|`itemId`| <span data-ttu-id="9e68b-331">Zeichenfolge</span><span class="sxs-lookup"><span data-stu-id="9e68b-331">String</span></span>|<span data-ttu-id="9e68b-332">Der EWS-Bezeichner (Exchange Web Services, Exchange-Webdienste) für einen vorhandenen Kalendertermin.</span><span class="sxs-lookup"><span data-stu-id="9e68b-332">The Exchange Web Services (EWS) identifier for an existing calendar appointment.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="9e68b-333">Anforderungen</span><span class="sxs-lookup"><span data-stu-id="9e68b-333">Requirements</span></span>

|<span data-ttu-id="9e68b-334">Anforderung</span><span class="sxs-lookup"><span data-stu-id="9e68b-334">Requirement</span></span>| <span data-ttu-id="9e68b-335">Wert</span><span class="sxs-lookup"><span data-stu-id="9e68b-335">Value</span></span>|
|---|---|
|[<span data-ttu-id="9e68b-336">Mindestversion des Postfachanforderungssatzes</span><span class="sxs-lookup"><span data-stu-id="9e68b-336">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="9e68b-337">1.0</span><span class="sxs-lookup"><span data-stu-id="9e68b-337">1.0</span></span>|
|[<span data-ttu-id="9e68b-338">Mindestberechtigungsstufe</span><span class="sxs-lookup"><span data-stu-id="9e68b-338">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="9e68b-339">ReadItem</span><span class="sxs-lookup"><span data-stu-id="9e68b-339">ReadItem</span></span>|
|[<span data-ttu-id="9e68b-340">Zutreffender Outlook-Modus</span><span class="sxs-lookup"><span data-stu-id="9e68b-340">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="9e68b-341">Verfassen oder Lesen</span><span class="sxs-lookup"><span data-stu-id="9e68b-341">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="9e68b-342">Beispiel</span><span class="sxs-lookup"><span data-stu-id="9e68b-342">Example</span></span>

```
Office.context.mailbox.displayAppointmentForm(appointmentId);
```

####  <a name="displaymessageformitemid"></a><span data-ttu-id="9e68b-343">displayMessageForm(itemId)</span><span class="sxs-lookup"><span data-stu-id="9e68b-343">displayMessageForm(itemId)</span></span>

<span data-ttu-id="9e68b-344">Zeigt eine vorhandene Nachricht an.</span><span class="sxs-lookup"><span data-stu-id="9e68b-344">Displays an existing message.</span></span>

> [!NOTE]
> <span data-ttu-id="9e68b-345">Diese Methode wird in Outlook für iOS oder Outlook für Android nicht unterstützt.</span><span class="sxs-lookup"><span data-stu-id="9e68b-345">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="9e68b-346">Die `displayMessageForm`-Methode öffnet eine vorhandene Nachricht in einem neuen Fenster auf dem Desktop bzw. in einem Dialogfeld auf Mobilgeräten.</span><span class="sxs-lookup"><span data-stu-id="9e68b-346">The `displayMessageForm` method opens an existing message in a new window on the desktop or in a dialog box on mobile devices.</span></span>

<span data-ttu-id="9e68b-347">In Outlook Web App wird mit dieser Methode das angegebene Formular nur dann geöffnet, wenn der Textkörper des Formular Zeichen im Umfang vom maximal 32 KB umfasst.</span><span class="sxs-lookup"><span data-stu-id="9e68b-347">In Outlook Web App, this method opens the specified form only if the body of the form is less than or equal to 32 KB number of characters.</span></span>

<span data-ttu-id="9e68b-348">Wenn der angegebene Elementbezeichner keine vorhandenen Nachrichten erkennt, wird auf dem Client-Computer keine Nachricht angezeigt, und es werden keine Fehlermeldungen zurückgegeben.</span><span class="sxs-lookup"><span data-stu-id="9e68b-348">If the specified item identifier does not identify an existing message, no message will be displayed on the client computer, and no error message will be returned.</span></span>

<span data-ttu-id="9e68b-p113">Verwenden Sie `displayMessageForm` nicht mit einem `itemId`-Objekt, das einen Termin darstellt. Verwenden Sie die `displayAppointmentForm`-Methode, um einen vorhandenen Termin anzuzeigen, und `displayNewAppointmentForm`, um ein Formular zum Erstellen eines neuen Termins anzuzeigen.</span><span class="sxs-lookup"><span data-stu-id="9e68b-p113">Do not use the `displayMessageForm` with an `itemId` that represents an appointment. Use the `displayAppointmentForm` method to display an existing appointment, and `displayNewAppointmentForm` to display a form to create a new appointment.</span></span>

##### <a name="parameters"></a><span data-ttu-id="9e68b-351">Parameter:</span><span class="sxs-lookup"><span data-stu-id="9e68b-351">Parameters:</span></span>

|<span data-ttu-id="9e68b-352">Name</span><span class="sxs-lookup"><span data-stu-id="9e68b-352">Name</span></span>| <span data-ttu-id="9e68b-353">Typ</span><span class="sxs-lookup"><span data-stu-id="9e68b-353">Type</span></span>| <span data-ttu-id="9e68b-354">Beschreibung</span><span class="sxs-lookup"><span data-stu-id="9e68b-354">Description</span></span>|
|---|---|---|
|`itemId`| <span data-ttu-id="9e68b-355">Zeichenfolge</span><span class="sxs-lookup"><span data-stu-id="9e68b-355">String</span></span>|<span data-ttu-id="9e68b-356">Der Exchange-Webdienste (EWS) für eine vorhandene Nachricht.</span><span class="sxs-lookup"><span data-stu-id="9e68b-356">The Exchange Web Services (EWS) identifier for an existing message.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="9e68b-357">Anforderungen</span><span class="sxs-lookup"><span data-stu-id="9e68b-357">Requirements</span></span>

|<span data-ttu-id="9e68b-358">Anforderung</span><span class="sxs-lookup"><span data-stu-id="9e68b-358">Requirement</span></span>| <span data-ttu-id="9e68b-359">Wert</span><span class="sxs-lookup"><span data-stu-id="9e68b-359">Value</span></span>|
|---|---|
|[<span data-ttu-id="9e68b-360">Mindestversion des Postfachanforderungssatzes</span><span class="sxs-lookup"><span data-stu-id="9e68b-360">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="9e68b-361">1.0</span><span class="sxs-lookup"><span data-stu-id="9e68b-361">1.0</span></span>|
|[<span data-ttu-id="9e68b-362">Mindestberechtigungsstufe</span><span class="sxs-lookup"><span data-stu-id="9e68b-362">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="9e68b-363">ReadItem</span><span class="sxs-lookup"><span data-stu-id="9e68b-363">ReadItem</span></span>|
|[<span data-ttu-id="9e68b-364">Zutreffender Outlook-Modus</span><span class="sxs-lookup"><span data-stu-id="9e68b-364">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="9e68b-365">Verfassen oder Lesen</span><span class="sxs-lookup"><span data-stu-id="9e68b-365">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="9e68b-366">Beispiel</span><span class="sxs-lookup"><span data-stu-id="9e68b-366">Example</span></span>

```
Office.context.mailbox.displayMessageForm(messageId);
```

#### <a name="displaynewappointmentformparameters"></a><span data-ttu-id="9e68b-367">displayNewAppointmentForm(parameters)</span><span class="sxs-lookup"><span data-stu-id="9e68b-367">displayNewAppointmentForm(parameters)</span></span>

<span data-ttu-id="9e68b-368">Zeigt ein Formular zum Erstellen eines neuen Kalendertermins an.</span><span class="sxs-lookup"><span data-stu-id="9e68b-368">Displays a form for creating a new calendar appointment.</span></span>

> [!NOTE]
> <span data-ttu-id="9e68b-369">Diese Methode wird in Outlook für iOS oder Outlook für Android nicht unterstützt.</span><span class="sxs-lookup"><span data-stu-id="9e68b-369">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="9e68b-p114">Mit der `displayNewAppointmentForm`-Methode wird ein Formular geöffnet, mit dem der Benutzer einen neuen Termin oder eine Besprechung erstellen kann. Wenn Parameter angegeben wurden, werden die Felder im Terminformular automatisch mit dem Inhalt der Parameter ausgefüllt.</span><span class="sxs-lookup"><span data-stu-id="9e68b-p114">The `displayNewAppointmentForm` method opens a form that enables the user to create a new appointment or meeting. If parameters are specified, the appointment form fields are automatically populated with the contents of the parameters.</span></span>

<span data-ttu-id="9e68b-p115">In Outlook Web App und OWA for Devices zeigt diese Methode immer ein Formular mit einem Teilnehmerfeld an. Wenn Sie keine Teilnehmer als Eingabeargumente angeben, zeigt die Methode ein Formular mit einer Schaltfläche **Speichern** an. Wenn Sie Teilnehmer angegeben haben, enthält das Formular die Teilnehmer und eine Schaltfläche **Senden**.</span><span class="sxs-lookup"><span data-stu-id="9e68b-p115">In Outlook Web App and OWA for Devices, this method always displays a form with an attendees field. If you do not specify any attendees as input arguments, the method displays a form with a **Save** button. If you have specified attendees, the form would include the attendees and a **Send** button.</span></span>

<span data-ttu-id="9e68b-p116">Wenn Teilnehmer oder Ressourcen im `requiredAttendees`-, `optionalAttendees`- oder `resources`-Parameter festgelegt werden, wird mit dieser Methode im Outlook Rich Client und Outlook RT ein Besprechungsformular mit der Schaltfläche **Senden** angezeigt. Wenn keine Teilnehmer angegeben werden, wird mit dieser Methode ein Terminformular mit der Schaltfläche **Speichern & schließen** angezeigt.</span><span class="sxs-lookup"><span data-stu-id="9e68b-p116">In the Outlook rich client and Outlook RT, if you specify any attendees or resources in the `requiredAttendees`, `optionalAttendees`, or `resources` parameter, this method displays a meeting form with a **Send** button. If you don't specify any recipients, this method displays an appointment form with a **Save & Close** button.</span></span>

<span data-ttu-id="9e68b-377">Wenn einer der Parameter die angegebenen Größenbeschränkungen überschreitet oder wenn ein unbekannter Parametername angegeben wird, wird eine Ausnahme ausgelöst.</span><span class="sxs-lookup"><span data-stu-id="9e68b-377">If any of the parameters exceed the specified size limits, or if an unknown parameter name is specified, an exception is thrown.</span></span>

##### <a name="parameters"></a><span data-ttu-id="9e68b-378">Parameter:</span><span class="sxs-lookup"><span data-stu-id="9e68b-378">Parameters:</span></span>

|<span data-ttu-id="9e68b-379">Name</span><span class="sxs-lookup"><span data-stu-id="9e68b-379">Name</span></span>| <span data-ttu-id="9e68b-380">Typ</span><span class="sxs-lookup"><span data-stu-id="9e68b-380">Type</span></span>| <span data-ttu-id="9e68b-381">Beschreibung</span><span class="sxs-lookup"><span data-stu-id="9e68b-381">Description</span></span>|
|---|---|---|
| `parameters` | <span data-ttu-id="9e68b-382">Object</span><span class="sxs-lookup"><span data-stu-id="9e68b-382">Object</span></span> | <span data-ttu-id="9e68b-383">Ein Wörterbuch mit Parametern, die den neuen Termin beschreiben</span><span class="sxs-lookup"><span data-stu-id="9e68b-383">A dictionary of parameters describing the new appointment.</span></span> |
| `parameters.requiredAttendees` | <span data-ttu-id="9e68b-384">Array.&lt;String&gt; &#124; Array.&lt;[EmailAddressDetails](/javascript/api/outlook_1_5/office.emailaddressdetails)&gt;</span><span class="sxs-lookup"><span data-stu-id="9e68b-384">Array.&lt;String&gt; &#124; Array.&lt;[EmailAddressDetails](/javascript/api/outlook_1_5/office.emailaddressdetails)&gt;</span></span> | <span data-ttu-id="9e68b-p117">Ein Array aus Zeichenfolgen mit den E-Mail-Adressen oder ein Array mit einem `EmailAddressDetails`-Objekt für jeden der erforderlichen Teilnehmer für den Termin. Das Array darf maximal 100 Einträge enthalten.</span><span class="sxs-lookup"><span data-stu-id="9e68b-p117">An array of strings containing the email addresses or an array containing an `EmailAddressDetails` object for each of the required attendees for the appointment. The array is limited to a maximum of 100 entries.</span></span> |
| `parameters.optionalAttendees` | <span data-ttu-id="9e68b-387">Array.&lt;String&gt; &#124; Array.&lt;[EmailAddressDetails](/javascript/api/outlook_1_5/office.emailaddressdetails)&gt;</span><span class="sxs-lookup"><span data-stu-id="9e68b-387">Array.&lt;String&gt; &#124; Array.&lt;[EmailAddressDetails](/javascript/api/outlook_1_5/office.emailaddressdetails)&gt;</span></span> | <span data-ttu-id="9e68b-p118">Ein Array aus Zeichenfolgen mit den E-Mail-Adressen oder ein Array mit einem Objekt des Typs `EmailAddressDetails` für jeden der optionalen Teilnehmer des Termins. Das Array darf maximal 100 Einträge enthalten.</span><span class="sxs-lookup"><span data-stu-id="9e68b-p118">An array of strings containing the email addresses or an array containing an `EmailAddressDetails` object for each of the optional attendees for the appointment. The array is limited to a maximum of 100 entries.</span></span> |
| `parameters.start` | <span data-ttu-id="9e68b-390">Datum</span><span class="sxs-lookup"><span data-stu-id="9e68b-390">Date</span></span> | <span data-ttu-id="9e68b-391">Ein Objekt des Typs `Date`. Gibt das Startdatum und den Beginn des Termins an.</span><span class="sxs-lookup"><span data-stu-id="9e68b-391">A `Date` object specifying the start date and time of the appointment.</span></span> |
| `parameters.end` | <span data-ttu-id="9e68b-392">Date</span><span class="sxs-lookup"><span data-stu-id="9e68b-392">Date</span></span> | <span data-ttu-id="9e68b-393">Ein Objekt des Typs `Date`. Gibt das Enddatum und das Ende des Termins an.</span><span class="sxs-lookup"><span data-stu-id="9e68b-393">A `Date` object specifying the end date and time of the appointment.</span></span> |
| `parameters.location` | <span data-ttu-id="9e68b-394">String</span><span class="sxs-lookup"><span data-stu-id="9e68b-394">String</span></span> | <span data-ttu-id="9e68b-p119">Eine Zeichenfolge mit dem Ort für den Termin. Die Zeichenfolge ist auf maximal 255 Zeichen beschränkt.</span><span class="sxs-lookup"><span data-stu-id="9e68b-p119">A string containing the location of the appointment. The string is limited to a maximum of 255 characters.</span></span> |
| `parameters.resources` | <span data-ttu-id="9e68b-397">Array.&lt;String&gt;</span><span class="sxs-lookup"><span data-stu-id="9e68b-397">Array.&lt;String&gt;</span></span> | <span data-ttu-id="9e68b-p120">Ein Array aus Zeichenfolgen, die die für den Termin erforderlichen Ressourcen enthalten. Das Array darf maximal 100 Einträge enthalten.</span><span class="sxs-lookup"><span data-stu-id="9e68b-p120">An array of strings containing the resources required for the appointment. The array is limited to a maximum of 100 entries.</span></span> |
| `parameters.subject` | <span data-ttu-id="9e68b-400">Zeichenfolge</span><span class="sxs-lookup"><span data-stu-id="9e68b-400">String</span></span> | <span data-ttu-id="9e68b-p121">Eine Zeichenfolge mit dem Betreff für den Termin. Die Zeichenfolge ist auf maximal 255 Zeichen beschränkt.</span><span class="sxs-lookup"><span data-stu-id="9e68b-p121">A string containing the subject of the appointment. The string is limited to a maximum of 255 characters.</span></span> |
| `parameters.body` | <span data-ttu-id="9e68b-403">Zeichenfolge</span><span class="sxs-lookup"><span data-stu-id="9e68b-403">String</span></span> | <span data-ttu-id="9e68b-p122">Der Text des Termins. Der Textinhalt darf maximal 32 KB umfassen.</span><span class="sxs-lookup"><span data-stu-id="9e68b-p122">The body of the appointment. The body content is limited to a maximum size of 32 KB.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="9e68b-406">Anforderungen</span><span class="sxs-lookup"><span data-stu-id="9e68b-406">Requirements</span></span>

|<span data-ttu-id="9e68b-407">Anforderung</span><span class="sxs-lookup"><span data-stu-id="9e68b-407">Requirement</span></span>| <span data-ttu-id="9e68b-408">Wert</span><span class="sxs-lookup"><span data-stu-id="9e68b-408">Value</span></span>|
|---|---|
|[<span data-ttu-id="9e68b-409">Mindestversion des Postfachanforderungssatzes</span><span class="sxs-lookup"><span data-stu-id="9e68b-409">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="9e68b-410">1.0</span><span class="sxs-lookup"><span data-stu-id="9e68b-410">1.0</span></span>|
|[<span data-ttu-id="9e68b-411">Mindestberechtigungsstufe</span><span class="sxs-lookup"><span data-stu-id="9e68b-411">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="9e68b-412">ReadItem</span><span class="sxs-lookup"><span data-stu-id="9e68b-412">ReadItem</span></span>|
|[<span data-ttu-id="9e68b-413">Zutreffender Outlook-Modus</span><span class="sxs-lookup"><span data-stu-id="9e68b-413">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="9e68b-414">Lesen</span><span class="sxs-lookup"><span data-stu-id="9e68b-414">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="9e68b-415">Beispiel</span><span class="sxs-lookup"><span data-stu-id="9e68b-415">Example</span></span>

```js
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

#### <a name="getcallbacktokenasyncoptions-callback"></a><span data-ttu-id="9e68b-416">getCallbackTokenAsync([options], callback)</span><span class="sxs-lookup"><span data-stu-id="9e68b-416">getCallbackTokenAsync([options], callback)</span></span>

<span data-ttu-id="9e68b-417">Ruft eine Zeichenfolge mit einem Token ab, das zum Aufrufen von REST-APIs oder Exchange-Webdiensten verwendet wird.</span><span class="sxs-lookup"><span data-stu-id="9e68b-417">Gets a string that contains a token used to call REST APIs or Exchange Web Services.</span></span>

<span data-ttu-id="9e68b-p123">Die `getCallbackTokenAsync`-Methode führt einen asynchronen Aufruf zum Abruf eines nicht transparenten Tokens vom Exchange-Server aus, der das Postfach des Benutzers hostet. Die Gültigkeitsdauer des Rückruftokens beträgt 5 Minuten.</span><span class="sxs-lookup"><span data-stu-id="9e68b-p123">The `getCallbackTokenAsync` method makes an asynchronous call to get an opaque token from the Exchange Server that hosts the user's mailbox. The lifetime of the callback token is 5 minutes.</span></span>

> [!NOTE]
> <span data-ttu-id="9e68b-420">Es wird empfohlen, dass-add-ins anstelle von Exchange-Webdienste nach Möglichkeit die REST-APIs verwenden.</span><span class="sxs-lookup"><span data-stu-id="9e68b-420">It is recommended that add-ins use the REST APIs instead of Exchange Web Services whenever possible.</span></span> 

<span data-ttu-id="9e68b-421">**REST-Tokens**</span><span class="sxs-lookup"><span data-stu-id="9e68b-421">**REST Tokens**</span></span>

<span data-ttu-id="9e68b-p124">Wenn ein REST-Token angefordert wird (`options.isRest = true`), kann das resultierende Token nicht zur Authentifizierung von Aufrufen der Exchange-Webdienste verwendet werden. Der Bereich des Tokens ist auf schreibgeschützten Zugriff auf das aktuelle Element und dessen Anlagen beschränkt, es sei denn im Manifest des Add-Ins ist die Berechtigung [`ReadWriteMailbox`](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions#readwritemailbox-permission) angegeben. Ist die Berechtigung `ReadWriteMailbox` angegeben, gewährt das resultierende Token Lese-/Schreibzugriff auf E-Mails, Kalender und Kontakte und erlaubt das Senden von E-Mails.</span><span class="sxs-lookup"><span data-stu-id="9e68b-p124">When a REST token is requested (`options.isRest = true`), the resulting token will not work to authenticate Exchange Web Services calls. The token will be limited in scope to read-only access to the current item and its attachments, unless the add-in has specified the [`ReadWriteMailbox`](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions#readwritemailbox-permission) permission in its manifest. If the `ReadWriteMailbox` permission is specified, the resulting token will grant read/write access to mail, calendar, and contacts, including the ability to send mail.</span></span>

<span data-ttu-id="9e68b-425">Das Add-In sollte die Eigenschaft `restUrl` verwenden, um die korrekte URL für REST-API-Aufrufe zu ermitteln.</span><span class="sxs-lookup"><span data-stu-id="9e68b-425">The add-in should use the `restUrl` property to determine the correct URL to use when making REST API calls.</span></span>

<span data-ttu-id="9e68b-426">**EWS-Tokens**</span><span class="sxs-lookup"><span data-stu-id="9e68b-426">**EWS Tokens**</span></span>

<span data-ttu-id="9e68b-p125">Wenn ein EWS-Token angefordert wird (`options.isRest = false`), kann das resultierende Token nicht zur Authentifizierung von REST-API-Aufrufen verwendet werden. Der Bereich des Tokens ist auf den Zugriff auf das aktuelle Element beschränkt.</span><span class="sxs-lookup"><span data-stu-id="9e68b-p125">When an EWS token is requested (`options.isRest = false`), the resulting token will not work to authenticate REST API calls. The token will be limited in scope to accessing the current item.</span></span>

<span data-ttu-id="9e68b-429">Das Add-In sollte die Eigenschaft `ewsUrl` verwenden, um die korrekte URL für EWS-Aufrufe zu ermitteln.</span><span class="sxs-lookup"><span data-stu-id="9e68b-429">The add-in should use the `ewsUrl` property to determine the correct URL to use when making EWS calls.</span></span>

##### <a name="parameters"></a><span data-ttu-id="9e68b-430">Parameter:</span><span class="sxs-lookup"><span data-stu-id="9e68b-430">Parameters:</span></span>

|<span data-ttu-id="9e68b-431">Name</span><span class="sxs-lookup"><span data-stu-id="9e68b-431">Name</span></span>| <span data-ttu-id="9e68b-432">Typ</span><span class="sxs-lookup"><span data-stu-id="9e68b-432">Type</span></span>| <span data-ttu-id="9e68b-433">Attribute</span><span class="sxs-lookup"><span data-stu-id="9e68b-433">Attributes</span></span>| <span data-ttu-id="9e68b-434">Beschreibung</span><span class="sxs-lookup"><span data-stu-id="9e68b-434">Description</span></span>|
|---|---|---|---|
| `options` | <span data-ttu-id="9e68b-435">Objekt</span><span class="sxs-lookup"><span data-stu-id="9e68b-435">Object</span></span> | <span data-ttu-id="9e68b-436">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="9e68b-436">&lt;optional&gt;</span></span> | <span data-ttu-id="9e68b-437">Ein Objektliteral, das mindestens eine der folgenden Eigenschaften enthält.</span><span class="sxs-lookup"><span data-stu-id="9e68b-437">An object literal that contains one or more of the following properties.</span></span> |
| `options.isRest` | <span data-ttu-id="9e68b-438">Boolean</span><span class="sxs-lookup"><span data-stu-id="9e68b-438">Boolean</span></span> |  <span data-ttu-id="9e68b-439">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="9e68b-439">&lt;optional&gt;</span></span> | <span data-ttu-id="9e68b-p126">Legt fest, ob das bereitgestellte Token für die Outlook-REST-APIs oder die Exchange-Webdienste verwendet wird. Der Standardwert ist `false`.</span><span class="sxs-lookup"><span data-stu-id="9e68b-p126">Determines if the token provided will be used for the Outlook REST APIs or Exchange Web Services. Default value is `false`.</span></span> |
| `options.asyncContext` | <span data-ttu-id="9e68b-442">Objekt</span><span class="sxs-lookup"><span data-stu-id="9e68b-442">Object</span></span> |  <span data-ttu-id="9e68b-443">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="9e68b-443">&lt;optional&gt;</span></span> | <span data-ttu-id="9e68b-444">Jegliche Statusdaten, die an die asynchrone Methode übergeben werden</span><span class="sxs-lookup"><span data-stu-id="9e68b-444">Any state data that is passed to the asynchronous method.</span></span> |
|`callback`| <span data-ttu-id="9e68b-445">function</span><span class="sxs-lookup"><span data-stu-id="9e68b-445">function</span></span>||<span data-ttu-id="9e68b-p127">Wenn die Methode abgeschlossen wird, wird die Funktion , die im `callback`-Parameter übergeben wird, mit einem einzelnen Parameter (`asyncResult`) aufgerufen, der ein [`AsyncResult`](/javascript/api/office/office.asyncresult)-Objekt darstellt. Das Token wird als Zeichenfolge in der `asyncResult.value`-Eigenschaft bereitgestellt.</span><span class="sxs-lookup"><span data-stu-id="9e68b-p127">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object. The token is provided as a string in the `asyncResult.value` property.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="9e68b-448">Anforderungen</span><span class="sxs-lookup"><span data-stu-id="9e68b-448">Requirements</span></span>

|<span data-ttu-id="9e68b-449">Anforderung</span><span class="sxs-lookup"><span data-stu-id="9e68b-449">Requirement</span></span>| <span data-ttu-id="9e68b-450">Wert</span><span class="sxs-lookup"><span data-stu-id="9e68b-450">Value</span></span>|
|---|---|
|[<span data-ttu-id="9e68b-451">Mindestversion des Postfachanforderungssatzes</span><span class="sxs-lookup"><span data-stu-id="9e68b-451">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="9e68b-452">1,5</span><span class="sxs-lookup"><span data-stu-id="9e68b-452">1.5</span></span> |
|[<span data-ttu-id="9e68b-453">Mindestberechtigungsstufe</span><span class="sxs-lookup"><span data-stu-id="9e68b-453">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="9e68b-454">ReadItem</span><span class="sxs-lookup"><span data-stu-id="9e68b-454">ReadItem</span></span>|
|[<span data-ttu-id="9e68b-455">Zutreffender Outlook-Modus</span><span class="sxs-lookup"><span data-stu-id="9e68b-455">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="9e68b-456">Verfassenmodus und Lesemodus</span><span class="sxs-lookup"><span data-stu-id="9e68b-456">Compose and read</span></span>|

##### <a name="example"></a><span data-ttu-id="9e68b-457">Beispiel</span><span class="sxs-lookup"><span data-stu-id="9e68b-457">Example</span></span>

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

#### <a name="getcallbacktokenasynccallback-usercontext"></a><span data-ttu-id="9e68b-458">getCallbackTokenAsync(callback, [userContext])</span><span class="sxs-lookup"><span data-stu-id="9e68b-458">getCallbackTokenAsync(callback, [userContext])</span></span>

<span data-ttu-id="9e68b-459">Ruft eine Zeichenfolge ab, die einen Token enthält, der verwendet wird, um eine Anlage oder ein Element von einem Exchange Server abzurufen.</span><span class="sxs-lookup"><span data-stu-id="9e68b-459">Gets a string that contains a token used to get an attachment or item from an Exchange Server.</span></span>

<span data-ttu-id="9e68b-p128">Die `getCallbackTokenAsync`-Methode führt einen asynchronen Aufruf zum Abruf eines nicht transparenten Tokens vom Exchange-Server aus, der das Postfach des Benutzers hostet. Die Gültigkeitsdauer des Rückruftokens beträgt 5 Minuten.</span><span class="sxs-lookup"><span data-stu-id="9e68b-p128">The `getCallbackTokenAsync` method makes an asynchronous call to get an opaque token from the Exchange Server that hosts the user's mailbox. The lifetime of the callback token is 5 minutes.</span></span>

<span data-ttu-id="9e68b-p129">Sie können das Token und einen Anlagen- oder einen Elementbezeichner an ein Drittanbietersystem weitergeben. Das Drittanbietersystem verwendet das Token als Trägerautorisierungstoken, um den EWS-Vorgang (Exchange-Webdienste) [GetAttachment](https://docs.microsoft.com/exchange/client-developer/web-service-reference/getattachment-operation) oder [GetItem](https://docs.microsoft.com/exchange/client-developer/web-service-reference/getitem-operation) aufzurufen und eine Anlage oder ein Element zurückzugeben. Sie können z. B. einen Remotedienst erstellen, um [Anlagen aus dem ausgewählten Element abzurufen](https://docs.microsoft.com/outlook/add-ins/get-attachments-of-an-outlook-item).</span><span class="sxs-lookup"><span data-stu-id="9e68b-p129">You can pass the token and an attachment identifier or item identifier to a third-party system. The third-party system uses the token as a bearer authorization token to call the Exchange Web Services (EWS) [GetAttachment](https://docs.microsoft.com/exchange/client-developer/web-service-reference/getattachment-operation) or [GetItem](https://docs.microsoft.com/exchange/client-developer/web-service-reference/getitem-operation) operation to return an attachment or item. For example, you can create a remote service to [get attachments from the selected item](https://docs.microsoft.com/outlook/add-ins/get-attachments-of-an-outlook-item).</span></span>

<span data-ttu-id="9e68b-465">Ihre App muss über die im Manifest angegebene **ReadItem**-Berechtigung verfügen, um die `getCallbackTokenAsync`-Methode im Lesemodus aufzurufen.</span><span class="sxs-lookup"><span data-stu-id="9e68b-465">Your app must have the **ReadItem** permission specified in its manifest to call the `getCallbackTokenAsync` method in read mode.</span></span>

<span data-ttu-id="9e68b-p130">Im Verfassenmodus müssen Sie die [`saveAsync`](Office.context.mailbox.item.md#saveasyncoptions-callback)-Methode aufrufen, um einen Elementbezeichner an die `getCallbackTokenAsync`-Methode zu übergeben. Die App muss über **ReadWriteItem**-Berechtigungen verfügen, um die `saveAsync`-Methode aufzurufen.</span><span class="sxs-lookup"><span data-stu-id="9e68b-p130">In compose mode you must call the [`saveAsync`](Office.context.mailbox.item.md#saveasyncoptions-callback) method to get an item identifier to pass to the `getCallbackTokenAsync` method. Your app must have **ReadWriteItem** permissions to call the `saveAsync` method.</span></span>

##### <a name="parameters"></a><span data-ttu-id="9e68b-468">Parameter:</span><span class="sxs-lookup"><span data-stu-id="9e68b-468">Parameters:</span></span>

|<span data-ttu-id="9e68b-469">Name</span><span class="sxs-lookup"><span data-stu-id="9e68b-469">Name</span></span>| <span data-ttu-id="9e68b-470">Typ</span><span class="sxs-lookup"><span data-stu-id="9e68b-470">Type</span></span>| <span data-ttu-id="9e68b-471">Attribute</span><span class="sxs-lookup"><span data-stu-id="9e68b-471">Attributes</span></span>| <span data-ttu-id="9e68b-472">Beschreibung</span><span class="sxs-lookup"><span data-stu-id="9e68b-472">Description</span></span>|
|---|---|---|---|
|`callback`| <span data-ttu-id="9e68b-473">Funktion</span><span class="sxs-lookup"><span data-stu-id="9e68b-473">function</span></span>||<span data-ttu-id="9e68b-p131">Wenn die Methode abgeschlossen wird, wird die Funktion , die im `callback`-Parameter übergeben wird, mit einem einzelnen Parameter (`asyncResult`) aufgerufen, der ein [`AsyncResult`](/javascript/api/office/office.asyncresult)-Objekt darstellt. Das Token wird als Zeichenfolge in der `asyncResult.value`-Eigenschaft bereitgestellt.</span><span class="sxs-lookup"><span data-stu-id="9e68b-p131">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object. The token is provided as a string in the `asyncResult.value` property.</span></span>|
|`userContext`| <span data-ttu-id="9e68b-476">Objekt</span><span class="sxs-lookup"><span data-stu-id="9e68b-476">Object</span></span>| <span data-ttu-id="9e68b-477">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="9e68b-477">&lt;optional&gt;</span></span>|<span data-ttu-id="9e68b-478">Jegliche Zustandsdaten, die an die asynchrone Methode übergeben werden.</span><span class="sxs-lookup"><span data-stu-id="9e68b-478">Any state data that is passed to the asynchronous method.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="9e68b-479">Anforderungen</span><span class="sxs-lookup"><span data-stu-id="9e68b-479">Requirements</span></span>

|<span data-ttu-id="9e68b-480">Anforderung</span><span class="sxs-lookup"><span data-stu-id="9e68b-480">Requirement</span></span>| <span data-ttu-id="9e68b-481">Wert</span><span class="sxs-lookup"><span data-stu-id="9e68b-481">Value</span></span>|
|---|---|
|[<span data-ttu-id="9e68b-482">Mindestversion des Postfachanforderungssatzes</span><span class="sxs-lookup"><span data-stu-id="9e68b-482">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="9e68b-483">1.3</span><span class="sxs-lookup"><span data-stu-id="9e68b-483">1.3</span></span>|
|[<span data-ttu-id="9e68b-484">Mindestberechtigungsstufe</span><span class="sxs-lookup"><span data-stu-id="9e68b-484">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="9e68b-485">ReadItem</span><span class="sxs-lookup"><span data-stu-id="9e68b-485">ReadItem</span></span>|
|[<span data-ttu-id="9e68b-486">Zutreffender Outlook-Modus</span><span class="sxs-lookup"><span data-stu-id="9e68b-486">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="9e68b-487">Verfassenmodus und Lesemodus</span><span class="sxs-lookup"><span data-stu-id="9e68b-487">Compose and read</span></span>|

##### <a name="example"></a><span data-ttu-id="9e68b-488">Beispiel</span><span class="sxs-lookup"><span data-stu-id="9e68b-488">Example</span></span>

```js
function getCallbackToken() {
  Office.context.mailbox.getCallbackTokenAsync(cb);
}

function cb(asyncResult) {
  var token = asyncResult.value;
}
```

####  <a name="getuseridentitytokenasynccallback-usercontext"></a><span data-ttu-id="9e68b-489">getUserIdentityTokenAsync(callback, [userContext])</span><span class="sxs-lookup"><span data-stu-id="9e68b-489">getUserIdentityTokenAsync(callback, [userContext])</span></span>

<span data-ttu-id="9e68b-490">Ruft ein Token ab, das den Benutzer und das Office-Add-In identifiziert.</span><span class="sxs-lookup"><span data-stu-id="9e68b-490">Gets a token identifying the user and the Office Add-in.</span></span>

<span data-ttu-id="9e68b-491">Die `getUserIdentityTokenAsync`-Methode gibt ein Token zurück, das Sie dazu verwenden können, um das [Add-In und den Benutzer mit einem Drittanbietersystem zu identifizieren und zu authentifizieren](https://docs.microsoft.com/outlook/add-ins/authentication).</span><span class="sxs-lookup"><span data-stu-id="9e68b-491">The `getUserIdentityTokenAsync` method returns a token that you can use to identify and [authenticate the add-in and user with a third-party system](https://docs.microsoft.com/outlook/add-ins/authentication).</span></span>

##### <a name="parameters"></a><span data-ttu-id="9e68b-492">Parameter:</span><span class="sxs-lookup"><span data-stu-id="9e68b-492">Parameters:</span></span>

|<span data-ttu-id="9e68b-493">Name</span><span class="sxs-lookup"><span data-stu-id="9e68b-493">Name</span></span>| <span data-ttu-id="9e68b-494">Typ</span><span class="sxs-lookup"><span data-stu-id="9e68b-494">Type</span></span>| <span data-ttu-id="9e68b-495">Attribute</span><span class="sxs-lookup"><span data-stu-id="9e68b-495">Attributes</span></span>| <span data-ttu-id="9e68b-496">Beschreibung</span><span class="sxs-lookup"><span data-stu-id="9e68b-496">Description</span></span>|
|---|---|---|---|
|`callback`| <span data-ttu-id="9e68b-497">Funktion</span><span class="sxs-lookup"><span data-stu-id="9e68b-497">function</span></span>||<span data-ttu-id="9e68b-498">Wenn die Methode abgeschlossen wird, wird die Funktion , die im `callback`-Parameter übergeben wird, mit einem einzelnen Parameter (`asyncResult`) aufgerufen, der ein [`AsyncResult`](/javascript/api/office/office.asyncresult)-Objekt darstellt.</span><span class="sxs-lookup"><span data-stu-id="9e68b-498">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span><br/><br/><span data-ttu-id="9e68b-499">Das Token wird als Zeichenfolge in der `asyncResult.value`-Eigenschaft bereitgestellt.</span><span class="sxs-lookup"><span data-stu-id="9e68b-499">The token is provided as a string in the `asyncResult.value` property.</span></span>|
|`userContext`| <span data-ttu-id="9e68b-500">Objekt</span><span class="sxs-lookup"><span data-stu-id="9e68b-500">Object</span></span>| <span data-ttu-id="9e68b-501">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="9e68b-501">&lt;optional&gt;</span></span>|<span data-ttu-id="9e68b-502">Jegliche Zustandsdaten, die an die asynchrone Methode übergeben werden.</span><span class="sxs-lookup"><span data-stu-id="9e68b-502">Any state data that is passed to the asynchronous method.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="9e68b-503">Anforderungen</span><span class="sxs-lookup"><span data-stu-id="9e68b-503">Requirements</span></span>

|<span data-ttu-id="9e68b-504">Anforderung</span><span class="sxs-lookup"><span data-stu-id="9e68b-504">Requirement</span></span>| <span data-ttu-id="9e68b-505">Wert</span><span class="sxs-lookup"><span data-stu-id="9e68b-505">Value</span></span>|
|---|---|
|[<span data-ttu-id="9e68b-506">Mindestversion des Postfachanforderungssatzes</span><span class="sxs-lookup"><span data-stu-id="9e68b-506">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="9e68b-507">1.0</span><span class="sxs-lookup"><span data-stu-id="9e68b-507">1.0</span></span>|
|[<span data-ttu-id="9e68b-508">Mindestberechtigungsstufe</span><span class="sxs-lookup"><span data-stu-id="9e68b-508">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="9e68b-509">ReadItem</span><span class="sxs-lookup"><span data-stu-id="9e68b-509">ReadItem</span></span>|
|[<span data-ttu-id="9e68b-510">Zutreffender Outlook-Modus</span><span class="sxs-lookup"><span data-stu-id="9e68b-510">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="9e68b-511">Verfassen oder Lesen</span><span class="sxs-lookup"><span data-stu-id="9e68b-511">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="9e68b-512">Beispiel</span><span class="sxs-lookup"><span data-stu-id="9e68b-512">Example</span></span>

```js
function getIdentityToken() {
  Office.context.mailbox.getUserIdentityTokenAsync(cb);
}

function cb(asyncResult) {
  var token = asyncResult.value;
}
```

####  <a name="makeewsrequestasyncdata-callback-usercontext"></a><span data-ttu-id="9e68b-513">makeEwsRequestAsync(data, callback, [userContext])</span><span class="sxs-lookup"><span data-stu-id="9e68b-513">makeEwsRequestAsync(data, callback, [userContext])</span></span>

<span data-ttu-id="9e68b-514">Richtet eine asynchrone Anforderung an einen EWS-Dienst (Exchange Web Services) auf dem Exchange-Server, der das Postfach des Benutzers hostet.</span><span class="sxs-lookup"><span data-stu-id="9e68b-514">Makes an asynchronous request to an Exchange Web Services (EWS) service on the Exchange server that hosts the user’s mailbox.</span></span>

> [!NOTE]
> <span data-ttu-id="9e68b-515">Diese Methode wird in den folgenden Szenarien nicht unterstützt.</span><span class="sxs-lookup"><span data-stu-id="9e68b-515">This method is not supported in the following scenarios.</span></span>
> - <span data-ttu-id="9e68b-516">In Outlook für iOS oder Outlook für Android</span><span class="sxs-lookup"><span data-stu-id="9e68b-516">In Outlook for iOS or Outlook for Android</span></span>
> - <span data-ttu-id="9e68b-517">Wenn das Add-In geladen wird, in einem Postfach Google Mail</span><span class="sxs-lookup"><span data-stu-id="9e68b-517">When the add-in is loaded in a Gmail mailbox</span></span>
> 
> <span data-ttu-id="9e68b-518">In diesen Fällen sollte-add-ins [mithilfe von REST-APIs](https://docs.microsoft.com/outlook/add-ins/use-rest-api) , stattdessen Zugriff auf das Postfach des Benutzers.</span><span class="sxs-lookup"><span data-stu-id="9e68b-518">In these cases, add-ins should [use REST APIs](https://docs.microsoft.com/outlook/add-ins/use-rest-api) to access the user's mailbox instead.</span></span>

<span data-ttu-id="9e68b-519">Die `makeEwsRequestAsync`-Methode sendet eine EWS-Anforderung für das Add-In zu Exchange.</span><span class="sxs-lookup"><span data-stu-id="9e68b-519">The `makeEwsRequestAsync` method sends an EWS request on behalf of the add-in to Exchange.</span></span> <span data-ttu-id="9e68b-520">Eine Liste der unterstützten EWS-Vorgänge finden Sie unter [Aufrufen von Webdiensten aus einem Outlook-add-in](https://docs.microsoft.com/outlook/add-ins/web-services#ews-operations-that-add-ins-support) .</span><span class="sxs-lookup"><span data-stu-id="9e68b-520">See [Call web services from an Outlook add-in](https://docs.microsoft.com/outlook/add-ins/web-services#ews-operations-that-add-ins-support) for a list of the supported EWS operations.</span></span>

<span data-ttu-id="9e68b-521">Sie können keine Elemente, die Ordnern zugeordnet sind, mit der `makeEwsRequestAsync`-Methode anfordern.</span><span class="sxs-lookup"><span data-stu-id="9e68b-521">You cannot request Folder Associated Items with the `makeEwsRequestAsync` method.</span></span>

<span data-ttu-id="9e68b-522">Die XML-Anfrage muss UTF-8-Codierung angeben.</span><span class="sxs-lookup"><span data-stu-id="9e68b-522">The XML request must specify UTF-8 encoding.</span></span>

```xml
<?xml version="1.0" encoding="utf-8"?>
```

<span data-ttu-id="9e68b-p133">Das Add-In muss über die **ReadWriteMailbox**-Berechtigung verfügen, um die `makeEwsRequestAsync`-Methode zu verwenden. Informationen zur **ReadWriteMailbox**-Berechtigung und zu den EWS-Vorgängen, die Sie mit der `makeEwsRequestAsync`-Methode aufrufen können, finden Sie unter [Angeben von Berechtigungen für den E-Mail-Add-In-Zugriff auf das Benutzerpostfach](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions).</span><span class="sxs-lookup"><span data-stu-id="9e68b-p133">Your add-in must have the **ReadWriteMailbox** permission to use the `makeEwsRequestAsync` method. For information about using the **ReadWriteMailbox** permission and the EWS operations that you can call with the `makeEwsRequestAsync` method, see [Specify permissions for mail add-in access to the user's mailbox](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions).</span></span>

> [!NOTE]
> <span data-ttu-id="9e68b-525">Der Serveradministrator muss festgelegt `OAuthAuthentication` auf "true" auf dem Client Access Server EWS-Verzeichnis so aktivieren Sie die `makeEwsRequestAsync` -Methode, um EWS stellen anfordert.</span><span class="sxs-lookup"><span data-stu-id="9e68b-525">The server administrator must set `OAuthAuthentication` to true on the Client Access Server EWS directory to enable the `makeEwsRequestAsync` method to make EWS requests.</span></span>

##### <a name="version-differences"></a><span data-ttu-id="9e68b-526">Versionsunterschiede</span><span class="sxs-lookup"><span data-stu-id="9e68b-526">Version differences</span></span>

<span data-ttu-id="9e68b-527">Wenn Sie die `makeEwsRequestAsync`-Methode in Mail-Apps verwenden, die in älteren Outlook-Versionen als Version 15.0.4535.1004 ausgeführt werden, sollten Sie den Codierungswert auf `ISO-8859-1` festlegen.</span><span class="sxs-lookup"><span data-stu-id="9e68b-527">When you use the `makeEwsRequestAsync` method in mail apps running in Outlook versions earlier than version 15.0.4535.1004, you should set the encoding value to `ISO-8859-1`.</span></span>

```xml
<?xml version="1.0" encoding="iso-8859-1"?>
```

<span data-ttu-id="9e68b-p134">Sie müssen den Codierungswert nicht festlegen, wenn Ihre Mail-App in Outlook im Web ausgeführt wird. Sie können mithilfe der mailbox.diagnostics.hostName-Eigenschaft ermitteln, ob Ihre Mail-App in Outlook oder Outlook im Web ausgeführt wird. Sie können mithilfe der mailbox.diagnostics.hostVersion-Eigenschaft ermitteln, welche Version von Outlook ausgeführt wird.</span><span class="sxs-lookup"><span data-stu-id="9e68b-p134">You do not need to set the encoding value when your mail app is running in Outlook on the web. You can determine whether your mail app is running in Outlook or Outlook on the web by using the mailbox.diagnostics.hostName property. You can determine what version of Outlook is running by using the mailbox.diagnostics.hostVersion property.</span></span>

##### <a name="parameters"></a><span data-ttu-id="9e68b-531">Parameter:</span><span class="sxs-lookup"><span data-stu-id="9e68b-531">Parameters:</span></span>

|<span data-ttu-id="9e68b-532">Name</span><span class="sxs-lookup"><span data-stu-id="9e68b-532">Name</span></span>| <span data-ttu-id="9e68b-533">Typ</span><span class="sxs-lookup"><span data-stu-id="9e68b-533">Type</span></span>| <span data-ttu-id="9e68b-534">Attribute</span><span class="sxs-lookup"><span data-stu-id="9e68b-534">Attributes</span></span>| <span data-ttu-id="9e68b-535">Beschreibung</span><span class="sxs-lookup"><span data-stu-id="9e68b-535">Description</span></span>|
|---|---|---|---|
|`data`| <span data-ttu-id="9e68b-536">Zeichenfolge</span><span class="sxs-lookup"><span data-stu-id="9e68b-536">String</span></span>||<span data-ttu-id="9e68b-537">Die EWS-Anforderung.</span><span class="sxs-lookup"><span data-stu-id="9e68b-537">The EWS request.</span></span>|
|`callback`| <span data-ttu-id="9e68b-538">Funktion</span><span class="sxs-lookup"><span data-stu-id="9e68b-538">function</span></span>||<span data-ttu-id="9e68b-539">Wenn die Methode abgeschlossen wird, wird die Funktion , die im `callback`-Parameter übergeben wird, mit einem einzelnen Parameter (`asyncResult`) aufgerufen, der ein [`AsyncResult`](/javascript/api/office/office.asyncresult)-Objekt darstellt.</span><span class="sxs-lookup"><span data-stu-id="9e68b-539">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span><br/><br/><span data-ttu-id="9e68b-540">Das XML-Ergebnis des EWS-Aufrufs wird als Zeichenfolge in der `asyncResult.value`-Eigenschaft bereitgestellt.</span><span class="sxs-lookup"><span data-stu-id="9e68b-540">The XML result of the EWS call is provided as a string in the `asyncResult.value` property.</span></span> <span data-ttu-id="9e68b-541">Wenn das Ergebnis 1 MB Größe überschreitet, wird stattdessen eine Fehlermeldung zurückgegeben.</span><span class="sxs-lookup"><span data-stu-id="9e68b-541">If the result exceeds 1 MB in size, an error message is returned instead.</span></span>|
|`userContext`| <span data-ttu-id="9e68b-542">Objekt</span><span class="sxs-lookup"><span data-stu-id="9e68b-542">Object</span></span>| <span data-ttu-id="9e68b-543">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="9e68b-543">&lt;optional&gt;</span></span>|<span data-ttu-id="9e68b-544">Jegliche Zustandsdaten, die an die asynchrone Methode übergeben werden.</span><span class="sxs-lookup"><span data-stu-id="9e68b-544">Any state data that is passed to the asynchronous method.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="9e68b-545">Anforderungen</span><span class="sxs-lookup"><span data-stu-id="9e68b-545">Requirements</span></span>

|<span data-ttu-id="9e68b-546">Anforderung</span><span class="sxs-lookup"><span data-stu-id="9e68b-546">Requirement</span></span>| <span data-ttu-id="9e68b-547">Wert</span><span class="sxs-lookup"><span data-stu-id="9e68b-547">Value</span></span>|
|---|---|
|[<span data-ttu-id="9e68b-548">Mindestversion des Postfachanforderungssatzes</span><span class="sxs-lookup"><span data-stu-id="9e68b-548">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="9e68b-549">1.0</span><span class="sxs-lookup"><span data-stu-id="9e68b-549">1.0</span></span>|
|[<span data-ttu-id="9e68b-550">Mindestberechtigungsstufe</span><span class="sxs-lookup"><span data-stu-id="9e68b-550">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="9e68b-551">ReadWriteMailbox</span><span class="sxs-lookup"><span data-stu-id="9e68b-551">ReadWriteMailbox</span></span>|
|[<span data-ttu-id="9e68b-552">Zutreffender Outlook-Modus</span><span class="sxs-lookup"><span data-stu-id="9e68b-552">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="9e68b-553">Verfassen oder Lesen</span><span class="sxs-lookup"><span data-stu-id="9e68b-553">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="9e68b-554">Beispiel</span><span class="sxs-lookup"><span data-stu-id="9e68b-554">Example</span></span>

<span data-ttu-id="9e68b-555">Das folgende Beispiel ruft `makeEwsRequestAsync` zum Verwenden des `GetItem`-Vorgangs auf, um den Betreff eines Elements abzurufen.</span><span class="sxs-lookup"><span data-stu-id="9e68b-555">The following example calls `makeEwsRequestAsync` to use the `GetItem` operation to get the subject of an item.</span></span>

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