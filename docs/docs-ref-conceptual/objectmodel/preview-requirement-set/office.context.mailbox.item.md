
# <a name="item"></a><span data-ttu-id="22e1f-101">item</span><span class="sxs-lookup"><span data-stu-id="22e1f-101">item</span></span>

### <a name="officeofficemdcontextofficecontextmdmailboxofficecontextmailboxmditem"></a><span data-ttu-id="22e1f-102">[Office](office.md)[.context](office.context.md)[.mailbox](office.context.mailbox.md).item</span><span class="sxs-lookup"><span data-stu-id="22e1f-102">[Office](office.md)[.context](office.context.md)[.mailbox](office.context.mailbox.md).item</span></span>

<span data-ttu-id="22e1f-p101">Der `item`-Namespace wird für den Zugriff auf die aktuell ausgewählte Nachricht, Besprechungsanfrage oder den aktuell ausgewählten Termin verwendet. Sie können den Typ von `item` mithilfe der [itemType](#itemtype-officemailboxenumsitemtypejavascriptapioutlookofficemailboxenumsitemtype)-Eigenschaft bestimmen.</span><span class="sxs-lookup"><span data-stu-id="22e1f-p101">The `item` namespace is used to access the currently selected message, meeting request, or appointment. You can determine the type of the `item` by using the [itemType](#itemtype-officemailboxenumsitemtypejavascriptapioutlookofficemailboxenumsitemtype) property.</span></span>

##### <a name="requirements"></a><span data-ttu-id="22e1f-105">Anforderungen</span><span class="sxs-lookup"><span data-stu-id="22e1f-105">Requirements</span></span>

|<span data-ttu-id="22e1f-106">Anforderung</span><span class="sxs-lookup"><span data-stu-id="22e1f-106">Requirement</span></span>|<span data-ttu-id="22e1f-107">Wert</span><span class="sxs-lookup"><span data-stu-id="22e1f-107">Value</span></span>|
|---|---|
|[<span data-ttu-id="22e1f-108">Mindestversion des Postfachanforderungssatzes</span><span class="sxs-lookup"><span data-stu-id="22e1f-108">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="22e1f-109">1.0</span><span class="sxs-lookup"><span data-stu-id="22e1f-109">1.0</span></span>|
|[<span data-ttu-id="22e1f-110">Mindestberechtigungsstufe</span><span class="sxs-lookup"><span data-stu-id="22e1f-110">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="22e1f-111">Eingeschränkt</span><span class="sxs-lookup"><span data-stu-id="22e1f-111">Restricted</span></span>|
|[<span data-ttu-id="22e1f-112">Zutreffender Outlook-Modus</span><span class="sxs-lookup"><span data-stu-id="22e1f-112">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="22e1f-113">Verfassen oder Lesen</span><span class="sxs-lookup"><span data-stu-id="22e1f-113">Compose or read</span></span>|

##### <a name="members-and-methods"></a><span data-ttu-id="22e1f-114">Elemente und Methoden</span><span class="sxs-lookup"><span data-stu-id="22e1f-114">Members and methods</span></span>

| <span data-ttu-id="22e1f-115">Element</span><span class="sxs-lookup"><span data-stu-id="22e1f-115">Member</span></span> | <span data-ttu-id="22e1f-116">Typ</span><span class="sxs-lookup"><span data-stu-id="22e1f-116">Type</span></span> |
|--------|------|
| [<span data-ttu-id="22e1f-117">attachments</span><span class="sxs-lookup"><span data-stu-id="22e1f-117">attachments</span></span>](#attachments-arrayattachmentdetailsjavascriptapioutlookofficeattachmentdetails) | <span data-ttu-id="22e1f-118">Element</span><span class="sxs-lookup"><span data-stu-id="22e1f-118">Member</span></span> |
| [<span data-ttu-id="22e1f-119">bcc</span><span class="sxs-lookup"><span data-stu-id="22e1f-119">bcc</span></span>](#bcc-recipientsjavascriptapioutlookofficerecipients) | <span data-ttu-id="22e1f-120">Element</span><span class="sxs-lookup"><span data-stu-id="22e1f-120">Member</span></span> |
| [<span data-ttu-id="22e1f-121">body</span><span class="sxs-lookup"><span data-stu-id="22e1f-121">body</span></span>](#body-bodyjavascriptapioutlookofficebody) | <span data-ttu-id="22e1f-122">Element</span><span class="sxs-lookup"><span data-stu-id="22e1f-122">Member</span></span> |
| [<span data-ttu-id="22e1f-123">cc</span><span class="sxs-lookup"><span data-stu-id="22e1f-123">cc</span></span>](#cc-arrayemailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsrecipientsjavascriptapioutlookofficerecipients) | <span data-ttu-id="22e1f-124">Element</span><span class="sxs-lookup"><span data-stu-id="22e1f-124">Member</span></span> |
| [<span data-ttu-id="22e1f-125">conversationId</span><span class="sxs-lookup"><span data-stu-id="22e1f-125">conversationId</span></span>](#nullable-conversationid-string) | <span data-ttu-id="22e1f-126">Element</span><span class="sxs-lookup"><span data-stu-id="22e1f-126">Member</span></span> |
| [<span data-ttu-id="22e1f-127">dateTimeCreated</span><span class="sxs-lookup"><span data-stu-id="22e1f-127">dateTimeCreated</span></span>](#datetimecreated-date) | <span data-ttu-id="22e1f-128">Element</span><span class="sxs-lookup"><span data-stu-id="22e1f-128">Member</span></span> |
| [<span data-ttu-id="22e1f-129">dateTimeModified</span><span class="sxs-lookup"><span data-stu-id="22e1f-129">dateTimeModified</span></span>](#datetimemodified-date) | <span data-ttu-id="22e1f-130">Element</span><span class="sxs-lookup"><span data-stu-id="22e1f-130">Member</span></span> |
| [<span data-ttu-id="22e1f-131">end</span><span class="sxs-lookup"><span data-stu-id="22e1f-131">end</span></span>](#end-datetimejavascriptapioutlookofficetime) | <span data-ttu-id="22e1f-132">Element</span><span class="sxs-lookup"><span data-stu-id="22e1f-132">Member</span></span> |
| [<span data-ttu-id="22e1f-133">from</span><span class="sxs-lookup"><span data-stu-id="22e1f-133">from</span></span>](#from-emailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsfromjavascriptapioutlookofficefrom) | <span data-ttu-id="22e1f-134">Element</span><span class="sxs-lookup"><span data-stu-id="22e1f-134">Member</span></span> |
| [<span data-ttu-id="22e1f-135">internetMessageId</span><span class="sxs-lookup"><span data-stu-id="22e1f-135">internetMessageId</span></span>](#internetmessageid-string) | <span data-ttu-id="22e1f-136">Element</span><span class="sxs-lookup"><span data-stu-id="22e1f-136">Member</span></span> |
| [<span data-ttu-id="22e1f-137">itemClass</span><span class="sxs-lookup"><span data-stu-id="22e1f-137">itemClass</span></span>](#itemclass-string) | <span data-ttu-id="22e1f-138">Element</span><span class="sxs-lookup"><span data-stu-id="22e1f-138">Member</span></span> |
| [<span data-ttu-id="22e1f-139">itemId</span><span class="sxs-lookup"><span data-stu-id="22e1f-139">itemId</span></span>](#nullable-itemid-string) | <span data-ttu-id="22e1f-140">Element</span><span class="sxs-lookup"><span data-stu-id="22e1f-140">Member</span></span> |
| [<span data-ttu-id="22e1f-141">itemType</span><span class="sxs-lookup"><span data-stu-id="22e1f-141">itemType</span></span>](#itemtype-officemailboxenumsitemtypejavascriptapioutlookofficemailboxenumsitemtype) | <span data-ttu-id="22e1f-142">Element</span><span class="sxs-lookup"><span data-stu-id="22e1f-142">Member</span></span> |
| [<span data-ttu-id="22e1f-143">location</span><span class="sxs-lookup"><span data-stu-id="22e1f-143">location</span></span>](#location-stringlocationjavascriptapioutlookofficelocation) | <span data-ttu-id="22e1f-144">Element</span><span class="sxs-lookup"><span data-stu-id="22e1f-144">Member</span></span> |
| [<span data-ttu-id="22e1f-145">normalizedSubject</span><span class="sxs-lookup"><span data-stu-id="22e1f-145">normalizedSubject</span></span>](#normalizedsubject-string) | <span data-ttu-id="22e1f-146">Element</span><span class="sxs-lookup"><span data-stu-id="22e1f-146">Member</span></span> |
| [<span data-ttu-id="22e1f-147">notificationMessages</span><span class="sxs-lookup"><span data-stu-id="22e1f-147">notificationMessages</span></span>](#notificationmessages-notificationmessagesjavascriptapioutlookofficenotificationmessages) | <span data-ttu-id="22e1f-148">Element</span><span class="sxs-lookup"><span data-stu-id="22e1f-148">Member</span></span> |
| [<span data-ttu-id="22e1f-149">optionalAttendees</span><span class="sxs-lookup"><span data-stu-id="22e1f-149">optionalAttendees</span></span>](#optionalattendees-arrayemailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsrecipientsjavascriptapioutlookofficerecipients) | <span data-ttu-id="22e1f-150">Element</span><span class="sxs-lookup"><span data-stu-id="22e1f-150">Member</span></span> |
| [<span data-ttu-id="22e1f-151">organizer</span><span class="sxs-lookup"><span data-stu-id="22e1f-151">organizer</span></span>](#organizer-emailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsorganizerjavascriptapioutlookofficeorganizer) | <span data-ttu-id="22e1f-152">Element</span><span class="sxs-lookup"><span data-stu-id="22e1f-152">Member</span></span> |
| [<span data-ttu-id="22e1f-153">recurrence</span><span class="sxs-lookup"><span data-stu-id="22e1f-153">recurrence</span></span>](#nullable-recurrence-recurrencejavascriptapioutlookofficerecurrence) | <span data-ttu-id="22e1f-154">Element</span><span class="sxs-lookup"><span data-stu-id="22e1f-154">Member</span></span> |
| [<span data-ttu-id="22e1f-155">requiredAttendees</span><span class="sxs-lookup"><span data-stu-id="22e1f-155">requiredAttendees</span></span>](#requiredattendees-arrayemailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsrecipientsjavascriptapioutlookofficerecipients) | <span data-ttu-id="22e1f-156">Element</span><span class="sxs-lookup"><span data-stu-id="22e1f-156">Member</span></span> |
| [<span data-ttu-id="22e1f-157">sender</span><span class="sxs-lookup"><span data-stu-id="22e1f-157">sender</span></span>](#sender-emailaddressdetailsjavascriptapioutlookofficeemailaddressdetails) | <span data-ttu-id="22e1f-158">Element</span><span class="sxs-lookup"><span data-stu-id="22e1f-158">Member</span></span> |
| [<span data-ttu-id="22e1f-159">seriesId</span><span class="sxs-lookup"><span data-stu-id="22e1f-159">seriesId</span></span>](#nullable-seriesid-string) | <span data-ttu-id="22e1f-160">Element</span><span class="sxs-lookup"><span data-stu-id="22e1f-160">Member</span></span> |
| [<span data-ttu-id="22e1f-161">start</span><span class="sxs-lookup"><span data-stu-id="22e1f-161">start</span></span>](#start-datetimejavascriptapioutlookofficetime) | <span data-ttu-id="22e1f-162">Element</span><span class="sxs-lookup"><span data-stu-id="22e1f-162">Member</span></span> |
| [<span data-ttu-id="22e1f-163">subject</span><span class="sxs-lookup"><span data-stu-id="22e1f-163">subject</span></span>](#subject-stringsubjectjavascriptapioutlookofficesubject) | <span data-ttu-id="22e1f-164">Element</span><span class="sxs-lookup"><span data-stu-id="22e1f-164">Member</span></span> |
| [<span data-ttu-id="22e1f-165">to</span><span class="sxs-lookup"><span data-stu-id="22e1f-165">to</span></span>](#to-arrayemailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsrecipientsjavascriptapioutlookofficerecipients) | <span data-ttu-id="22e1f-166">Element</span><span class="sxs-lookup"><span data-stu-id="22e1f-166">Member</span></span> |
| [<span data-ttu-id="22e1f-167">addFileAttachmentAsync</span><span class="sxs-lookup"><span data-stu-id="22e1f-167">addFileAttachmentAsync</span></span>](#addfileattachmentasyncuri-attachmentname-options-callback) | <span data-ttu-id="22e1f-168">Methode</span><span class="sxs-lookup"><span data-stu-id="22e1f-168">Method</span></span> |
| [<span data-ttu-id="22e1f-169">addFileAttachmentFromBase64Async</span><span class="sxs-lookup"><span data-stu-id="22e1f-169">addFileAttachmentFromBase64Async</span></span>](#addfileattachmentfrombase64asyncbase64file-attachmentname-options-callback) | <span data-ttu-id="22e1f-170">Methode</span><span class="sxs-lookup"><span data-stu-id="22e1f-170">Method</span></span> |
| [<span data-ttu-id="22e1f-171">addHandlerAsync</span><span class="sxs-lookup"><span data-stu-id="22e1f-171">addHandlerAsync</span></span>](#addhandlerasynceventtype-handler-options-callback) | <span data-ttu-id="22e1f-172">Methode</span><span class="sxs-lookup"><span data-stu-id="22e1f-172">Method</span></span> |
| [<span data-ttu-id="22e1f-173">addItemAttachmentAsync</span><span class="sxs-lookup"><span data-stu-id="22e1f-173">addItemAttachmentAsync</span></span>](#additemattachmentasyncitemid-attachmentname-options-callback) | <span data-ttu-id="22e1f-174">Methode</span><span class="sxs-lookup"><span data-stu-id="22e1f-174">Method</span></span> |
| [<span data-ttu-id="22e1f-175">close</span><span class="sxs-lookup"><span data-stu-id="22e1f-175">close</span></span>](#close) | <span data-ttu-id="22e1f-176">Methode</span><span class="sxs-lookup"><span data-stu-id="22e1f-176">Method</span></span> |
| [<span data-ttu-id="22e1f-177">displayReplyAllForm</span><span class="sxs-lookup"><span data-stu-id="22e1f-177">displayReplyAllForm</span></span>](#displayreplyallformformdata) | <span data-ttu-id="22e1f-178">Methode</span><span class="sxs-lookup"><span data-stu-id="22e1f-178">Method</span></span> |
| [<span data-ttu-id="22e1f-179">displayReplyForm</span><span class="sxs-lookup"><span data-stu-id="22e1f-179">displayReplyForm</span></span>](#displayreplyformformdata) | <span data-ttu-id="22e1f-180">Methode</span><span class="sxs-lookup"><span data-stu-id="22e1f-180">Method</span></span> |
| [<span data-ttu-id="22e1f-181">getEntities</span><span class="sxs-lookup"><span data-stu-id="22e1f-181">getEntities</span></span>](#getentities--entitiesjavascriptapioutlookofficeentities) | <span data-ttu-id="22e1f-182">Methode</span><span class="sxs-lookup"><span data-stu-id="22e1f-182">Method</span></span> |
| [<span data-ttu-id="22e1f-183">getEntitiesByType</span><span class="sxs-lookup"><span data-stu-id="22e1f-183">getEntitiesByType</span></span>](#getentitiesbytypeentitytype--nullable-arraystringcontactjavascriptapioutlookofficecontactmeetingsuggestionjavascriptapioutlookofficemeetingsuggestionphonenumberjavascriptapioutlookofficephonenumbertasksuggestionjavascriptapioutlookofficetasksuggestion) | <span data-ttu-id="22e1f-184">Methode</span><span class="sxs-lookup"><span data-stu-id="22e1f-184">Method</span></span> |
| [<span data-ttu-id="22e1f-185">getFilteredEntitiesByName</span><span class="sxs-lookup"><span data-stu-id="22e1f-185">getFilteredEntitiesByName</span></span>](#getfilteredentitiesbynamename--nullable-arraystringcontactjavascriptapioutlookofficecontactmeetingsuggestionjavascriptapioutlookofficemeetingsuggestionphonenumberjavascriptapioutlookofficephonenumbertasksuggestionjavascriptapioutlookofficetasksuggestion) | <span data-ttu-id="22e1f-186">Methode</span><span class="sxs-lookup"><span data-stu-id="22e1f-186">Method</span></span> |
| [<span data-ttu-id="22e1f-187">getInitializationContextAsync</span><span class="sxs-lookup"><span data-stu-id="22e1f-187">getInitializationContextAsync</span></span>](#getinitializationcontextasyncoptions-callback) | <span data-ttu-id="22e1f-188">Methode</span><span class="sxs-lookup"><span data-stu-id="22e1f-188">Method</span></span> |
| [<span data-ttu-id="22e1f-189">getRegExMatches</span><span class="sxs-lookup"><span data-stu-id="22e1f-189">getRegExMatches</span></span>](#getregexmatches--object) | <span data-ttu-id="22e1f-190">Methode</span><span class="sxs-lookup"><span data-stu-id="22e1f-190">Method</span></span> |
| [<span data-ttu-id="22e1f-191">getRegExMatchesByName</span><span class="sxs-lookup"><span data-stu-id="22e1f-191">getRegExMatchesByName</span></span>](#getregexmatchesbynamename--nullable-array-string-) | <span data-ttu-id="22e1f-192">Methode</span><span class="sxs-lookup"><span data-stu-id="22e1f-192">Method</span></span> |
| [<span data-ttu-id="22e1f-193">getSelectedDataAsync</span><span class="sxs-lookup"><span data-stu-id="22e1f-193">getSelectedDataAsync</span></span>](#getselecteddataasynccoerciontype-options-callback--string) | <span data-ttu-id="22e1f-194">Methode</span><span class="sxs-lookup"><span data-stu-id="22e1f-194">Method</span></span> |
| [<span data-ttu-id="22e1f-195">getSelectedEntities</span><span class="sxs-lookup"><span data-stu-id="22e1f-195">getSelectedEntities</span></span>](#getselectedentities--entitiesjavascriptapioutlookofficeentities) | <span data-ttu-id="22e1f-196">Methode</span><span class="sxs-lookup"><span data-stu-id="22e1f-196">Method</span></span> |
| [<span data-ttu-id="22e1f-197">getSelectedRegExMatches</span><span class="sxs-lookup"><span data-stu-id="22e1f-197">getSelectedRegExMatches</span></span>](#getselectedregexmatches--object) | <span data-ttu-id="22e1f-198">Methode</span><span class="sxs-lookup"><span data-stu-id="22e1f-198">Method</span></span> |
| [<span data-ttu-id="22e1f-199">loadCustomPropertiesAsync</span><span class="sxs-lookup"><span data-stu-id="22e1f-199">loadCustomPropertiesAsync</span></span>](#loadcustompropertiesasynccallback-usercontext) | <span data-ttu-id="22e1f-200">Methode</span><span class="sxs-lookup"><span data-stu-id="22e1f-200">Method</span></span> |
| [<span data-ttu-id="22e1f-201">removeAttachmentAsync</span><span class="sxs-lookup"><span data-stu-id="22e1f-201">removeAttachmentAsync</span></span>](#removeattachmentasyncattachmentid-options-callback) | <span data-ttu-id="22e1f-202">Methode</span><span class="sxs-lookup"><span data-stu-id="22e1f-202">Method</span></span> |
| [<span data-ttu-id="22e1f-203">removeHandlerAsync</span><span class="sxs-lookup"><span data-stu-id="22e1f-203">removeHandlerAsync</span></span>](#removehandlerasynceventtype-handler-options-callback) | <span data-ttu-id="22e1f-204">Methode</span><span class="sxs-lookup"><span data-stu-id="22e1f-204">Method</span></span> |
| [<span data-ttu-id="22e1f-205">saveAsync</span><span class="sxs-lookup"><span data-stu-id="22e1f-205">saveAsync</span></span>](#saveasyncoptions-callback) | <span data-ttu-id="22e1f-206">Methode</span><span class="sxs-lookup"><span data-stu-id="22e1f-206">Method</span></span> |
| [<span data-ttu-id="22e1f-207">setSelectedDataAsync</span><span class="sxs-lookup"><span data-stu-id="22e1f-207">setSelectedDataAsync</span></span>](#setselecteddataasyncdata-options-callback) | <span data-ttu-id="22e1f-208">Methode</span><span class="sxs-lookup"><span data-stu-id="22e1f-208">Method</span></span> |

### <a name="example"></a><span data-ttu-id="22e1f-209">Beispiel</span><span class="sxs-lookup"><span data-stu-id="22e1f-209">Example</span></span>

<span data-ttu-id="22e1f-210">Im folgenden JavaScript-Codebeispiel wird der Zugriff auf die `subject`-Eigenschaft des aktuellen Elements in Outlook veranschaulicht.</span><span class="sxs-lookup"><span data-stu-id="22e1f-210">The following JavaScript code example shows how to access the `subject` property of the current item in Outlook.</span></span>

```
// The initialize function is required for all apps.
Office.initialize = function () {
    // Checks for the DOM to load using the jQuery ready function.
    $(document).ready(function () {
    // After the DOM is loaded, app-specific code can run.
    var item = Office.context.mailbox.item;
    var subject = item.subject;
    // Continue with processing the subject of the current item,
    // which can be a message or appointment.
    });
}
```

### <a name="members"></a><span data-ttu-id="22e1f-211">Elemente</span><span class="sxs-lookup"><span data-stu-id="22e1f-211">Members</span></span>

#### <a name="attachments-arrayattachmentdetailsjavascriptapioutlookofficeattachmentdetails"></a><span data-ttu-id="22e1f-212">attachments :Array.<[AttachmentDetails](/javascript/api/outlook/office.attachmentdetails)></span><span class="sxs-lookup"><span data-stu-id="22e1f-212">attachments :Array.<[AttachmentDetails](/javascript/api/outlook/office.attachmentdetails)></span></span>

<span data-ttu-id="22e1f-p102">Ruft ein Array mit Anlagen für das Element ab. Nur Lesemodus.</span><span class="sxs-lookup"><span data-stu-id="22e1f-p102">Gets an array of attachments for the item. Read mode only.</span></span>

> [!NOTE]
> <span data-ttu-id="22e1f-215">Bestimmte Dateitypen werden aufgrund von Sicherheitsproblemen von Outlook blockiert und daher nicht zurückgegeben werden.</span><span class="sxs-lookup"><span data-stu-id="22e1f-215">Certain types of files are blocked by Outlook due to potential security issues and are therefore not returned.</span></span> <span data-ttu-id="22e1f-216">Weitere Informationen finden Sie unter [Blockierte Anlagen in Outlook](https://support.office.com/article/Blocked-attachments-in-Outlook-434752E1-02D3-4E90-9124-8B81E49A8519).</span><span class="sxs-lookup"><span data-stu-id="22e1f-216">For more information, see [Blocked attachments in Outlook](https://support.office.com/article/Blocked-attachments-in-Outlook-434752E1-02D3-4E90-9124-8B81E49A8519).</span></span>

##### <a name="type"></a><span data-ttu-id="22e1f-217">Typ:</span><span class="sxs-lookup"><span data-stu-id="22e1f-217">Type:</span></span>

*   <span data-ttu-id="22e1f-218">Array.<[AttachmentDetails](/javascript/api/outlook/office.attachmentdetails)></span><span class="sxs-lookup"><span data-stu-id="22e1f-218">Array.<[AttachmentDetails](/javascript/api/outlook/office.attachmentdetails)></span></span>

##### <a name="requirements"></a><span data-ttu-id="22e1f-219">Anforderungen</span><span class="sxs-lookup"><span data-stu-id="22e1f-219">Requirements</span></span>

|<span data-ttu-id="22e1f-220">Anforderung</span><span class="sxs-lookup"><span data-stu-id="22e1f-220">Requirement</span></span>|<span data-ttu-id="22e1f-221">Wert</span><span class="sxs-lookup"><span data-stu-id="22e1f-221">Value</span></span>|
|---|---|
|[<span data-ttu-id="22e1f-222">Mindestversion des Postfachanforderungssatzes</span><span class="sxs-lookup"><span data-stu-id="22e1f-222">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="22e1f-223">1.0</span><span class="sxs-lookup"><span data-stu-id="22e1f-223">1.0</span></span>|
|[<span data-ttu-id="22e1f-224">Mindestberechtigungsstufe</span><span class="sxs-lookup"><span data-stu-id="22e1f-224">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="22e1f-225">ReadItem</span><span class="sxs-lookup"><span data-stu-id="22e1f-225">ReadItem</span></span>|
|[<span data-ttu-id="22e1f-226">Zutreffender Outlook-Modus</span><span class="sxs-lookup"><span data-stu-id="22e1f-226">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="22e1f-227">Lesen</span><span class="sxs-lookup"><span data-stu-id="22e1f-227">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="22e1f-228">Beispiel</span><span class="sxs-lookup"><span data-stu-id="22e1f-228">Example</span></span>

<span data-ttu-id="22e1f-229">Mit dem folgende Code wird eine HTML-Zeichenfolge mit Details aller Anlagen für das aktuelle Element erstellt.</span><span class="sxs-lookup"><span data-stu-id="22e1f-229">The following code builds an HTML string with details of all attachments on the current item.</span></span>

```
var _Item = Office.context.mailbox.item;
var outputString = "";

if (_Item.attachments.length > 0) {
  for (i = 0 ; i < _Item.attachments.length ; i++) {
    var _att = _Item.attachments[i];
    outputString += "<BR>" + i + ". Name: ";
    outputString += _att.name;
    outputString += "<BR>ID: " + _att.id;
    outputString += "<BR>contentType: " + _att.contentType;
    outputString += "<BR>size: " + _att.size;
    outputString += "<BR>attachmentType: " + _att.attachmentType;
    outputString += "<BR>isInline: " + _att.isInline;
  }
}

// Do something with outputString
```

####  <a name="bcc-recipientsjavascriptapioutlookofficerecipients"></a><span data-ttu-id="22e1f-230">bcc :[Recipients](/javascript/api/outlook/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="22e1f-230">bcc :[Recipients](/javascript/api/outlook/office.recipients)</span></span>

<span data-ttu-id="22e1f-231">Ruft ein Objekt, das Methoden zum Abrufen oder aktualisieren die Empfänger in der Zeile Bcc (blind Carbon Copy, Blindkopie) einer Nachricht bereitstellt.</span><span class="sxs-lookup"><span data-stu-id="22e1f-231">Gets an object that provides methods to get or update the recipients on the Bcc (blind carbon copy) line of a message.</span></span> <span data-ttu-id="22e1f-232">Nur Verfassenmodus.</span><span class="sxs-lookup"><span data-stu-id="22e1f-232">Compose mode only.</span></span>

##### <a name="type"></a><span data-ttu-id="22e1f-233">Typ:</span><span class="sxs-lookup"><span data-stu-id="22e1f-233">Type:</span></span>

*   [<span data-ttu-id="22e1f-234">Recipients</span><span class="sxs-lookup"><span data-stu-id="22e1f-234">Recipients</span></span>](/javascript/api/outlook/office.recipients)

##### <a name="requirements"></a><span data-ttu-id="22e1f-235">Anforderungen</span><span class="sxs-lookup"><span data-stu-id="22e1f-235">Requirements</span></span>

|<span data-ttu-id="22e1f-236">Anforderung</span><span class="sxs-lookup"><span data-stu-id="22e1f-236">Requirement</span></span>|<span data-ttu-id="22e1f-237">Wert</span><span class="sxs-lookup"><span data-stu-id="22e1f-237">Value</span></span>|
|---|---|
|[<span data-ttu-id="22e1f-238">Mindestversion des Postfachanforderungssatzes</span><span class="sxs-lookup"><span data-stu-id="22e1f-238">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="22e1f-239">1.1</span><span class="sxs-lookup"><span data-stu-id="22e1f-239">1.1</span></span>|
|[<span data-ttu-id="22e1f-240">Mindestberechtigungsstufe</span><span class="sxs-lookup"><span data-stu-id="22e1f-240">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="22e1f-241">ReadItem</span><span class="sxs-lookup"><span data-stu-id="22e1f-241">ReadItem</span></span>|
|[<span data-ttu-id="22e1f-242">Zutreffender Outlook-Modus</span><span class="sxs-lookup"><span data-stu-id="22e1f-242">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="22e1f-243">Verfassen</span><span class="sxs-lookup"><span data-stu-id="22e1f-243">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="22e1f-244">Beispiel</span><span class="sxs-lookup"><span data-stu-id="22e1f-244">Example</span></span>

```
Office.context.mailbox.item.bcc.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.bcc.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.bcc.getAsync(callback);

function callback(asyncResult) {
  var arrayOfBccRecipients = asyncResult.value;
}
```

####  <a name="body-bodyjavascriptapioutlookofficebody"></a><span data-ttu-id="22e1f-245">body :[Body](/javascript/api/outlook/office.body)</span><span class="sxs-lookup"><span data-stu-id="22e1f-245">body :[Body](/javascript/api/outlook/office.body)</span></span>

<span data-ttu-id="22e1f-246">Ruft ein Objekt ab, das Methoden zum Bearbeiten des Textkörpers eines Elements bereitstellt.</span><span class="sxs-lookup"><span data-stu-id="22e1f-246">Gets an object that provides methods for manipulating the body of an item.</span></span>

##### <a name="type"></a><span data-ttu-id="22e1f-247">Typ:</span><span class="sxs-lookup"><span data-stu-id="22e1f-247">Type:</span></span>

*   [<span data-ttu-id="22e1f-248">Body</span><span class="sxs-lookup"><span data-stu-id="22e1f-248">Body</span></span>](/javascript/api/outlook/office.body)

##### <a name="requirements"></a><span data-ttu-id="22e1f-249">Anforderungen</span><span class="sxs-lookup"><span data-stu-id="22e1f-249">Requirements</span></span>

|<span data-ttu-id="22e1f-250">Anforderung</span><span class="sxs-lookup"><span data-stu-id="22e1f-250">Requirement</span></span>|<span data-ttu-id="22e1f-251">Wert</span><span class="sxs-lookup"><span data-stu-id="22e1f-251">Value</span></span>|
|---|---|
|[<span data-ttu-id="22e1f-252">Mindestversion des Postfachanforderungssatzes</span><span class="sxs-lookup"><span data-stu-id="22e1f-252">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="22e1f-253">1.1</span><span class="sxs-lookup"><span data-stu-id="22e1f-253">1.1</span></span>|
|[<span data-ttu-id="22e1f-254">Mindestberechtigungsstufe</span><span class="sxs-lookup"><span data-stu-id="22e1f-254">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="22e1f-255">ReadItem</span><span class="sxs-lookup"><span data-stu-id="22e1f-255">ReadItem</span></span>|
|[<span data-ttu-id="22e1f-256">Zutreffender Outlook-Modus</span><span class="sxs-lookup"><span data-stu-id="22e1f-256">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="22e1f-257">Verfassen oder Lesen</span><span class="sxs-lookup"><span data-stu-id="22e1f-257">Compose or read</span></span>|

####  <a name="cc-arrayemailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsrecipientsjavascriptapioutlookofficerecipients"></a><span data-ttu-id="22e1f-258">cc: Array. <[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)>|[Empfänger](/javascript/api/outlook/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="22e1f-258">cc :Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook/office.recipients)</span></span>

<span data-ttu-id="22e1f-259">Ermöglicht den Zugriff auf die (Carbon Copy, Blindkopie) Cc-Empfänger einer Nachricht.</span><span class="sxs-lookup"><span data-stu-id="22e1f-259">Provides access to the Cc (carbon copy) recipients of a message.</span></span> <span data-ttu-id="22e1f-260">Der Typ des Objekts und die Zugriffsebene, hängt von den Modus des aktuellen Elements ab.</span><span class="sxs-lookup"><span data-stu-id="22e1f-260">The type of object and level of access depends on the mode of the current item.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="22e1f-261">Lesemodus</span><span class="sxs-lookup"><span data-stu-id="22e1f-261">Read mode</span></span>

<span data-ttu-id="22e1f-p106">Die `cc`-Eigenschaft gibt ein Array mit einem `EmailAddressDetails`-Objekt für jeden Empfänger in der **Cc**-Zeile der Nachricht zurück. Die Auflistung ist auf höchstens 100 Elemente beschränkt.</span><span class="sxs-lookup"><span data-stu-id="22e1f-p106">The `cc` property returns an array that contains an `EmailAddressDetails` object for each recipient listed on the **Cc** line of the message. The collection is limited to a maximum of 100 members.</span></span>

##### <a name="compose-mode"></a><span data-ttu-id="22e1f-264">Verfassenmodus</span><span class="sxs-lookup"><span data-stu-id="22e1f-264">Compose mode</span></span>

<span data-ttu-id="22e1f-265">Die `cc` -Eigenschaft gibt eine `Recipients` -Objekt, das Methoden zum Abrufen oder aktualisieren die Empfänger in der Zeile **Cc** der Nachricht bereitstellt.</span><span class="sxs-lookup"><span data-stu-id="22e1f-265">The `cc` property returns a `Recipients` object that provides methods to get or update the recipients on the **Cc** line of the message.</span></span>

##### <a name="type"></a><span data-ttu-id="22e1f-266">Typ:</span><span class="sxs-lookup"><span data-stu-id="22e1f-266">Type:</span></span>

*   <span data-ttu-id="22e1f-267">Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="22e1f-267">Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook/office.recipients)</span></span>

##### <a name="requirements"></a><span data-ttu-id="22e1f-268">Anforderungen</span><span class="sxs-lookup"><span data-stu-id="22e1f-268">Requirements</span></span>

|<span data-ttu-id="22e1f-269">Anforderung</span><span class="sxs-lookup"><span data-stu-id="22e1f-269">Requirement</span></span>|<span data-ttu-id="22e1f-270">Wert</span><span class="sxs-lookup"><span data-stu-id="22e1f-270">Value</span></span>|
|---|---|
|[<span data-ttu-id="22e1f-271">Mindestversion des Postfachanforderungssatzes</span><span class="sxs-lookup"><span data-stu-id="22e1f-271">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="22e1f-272">1.0</span><span class="sxs-lookup"><span data-stu-id="22e1f-272">1.0</span></span>|
|[<span data-ttu-id="22e1f-273">Mindestberechtigungsstufe</span><span class="sxs-lookup"><span data-stu-id="22e1f-273">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="22e1f-274">ReadItem</span><span class="sxs-lookup"><span data-stu-id="22e1f-274">ReadItem</span></span>|
|[<span data-ttu-id="22e1f-275">Zutreffender Outlook-Modus</span><span class="sxs-lookup"><span data-stu-id="22e1f-275">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="22e1f-276">Verfassen oder Lesen</span><span class="sxs-lookup"><span data-stu-id="22e1f-276">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="22e1f-277">Beispiel</span><span class="sxs-lookup"><span data-stu-id="22e1f-277">Example</span></span>

```
Office.context.mailbox.item.cc.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.cc.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.cc.getAsync(callback);

function callback(asyncResult) {
  var arrayOfCcRecipients = asyncResult.value;
}
```

####  <a name="nullable-conversationid-string"></a><span data-ttu-id="22e1f-278">(nullable) conversationId :String</span><span class="sxs-lookup"><span data-stu-id="22e1f-278">(nullable) conversationId :String</span></span>

<span data-ttu-id="22e1f-279">Ruft einen Bezeichner für die E-Mail-Unterhaltung ab, in der eine bestimmte Nachricht enthalten ist.</span><span class="sxs-lookup"><span data-stu-id="22e1f-279">Gets an identifier for the email conversation that contains a particular message.</span></span>

<span data-ttu-id="22e1f-p107">Sie können für diese Eigenschaft eine ganze Zahl abrufen, wenn Ihre Mail-App in Formularen zum Lesen oder Antworten in Formularen zum Verfassen aktiviert wird. Wenn der Benutzer den Betreff der Antwortnachricht ändert, ändert sich beim Versenden die Konversations-ID für die entsprechende Nachricht, und der Wert, den Sie vorher bezogen haben, trifft nicht länger zu.</span><span class="sxs-lookup"><span data-stu-id="22e1f-p107">You can get an integer for this property if your mail app is activated in read forms or responses in compose forms. If subsequently the user changes the subject of the reply message, upon sending the reply, the conversation ID for that message will change and that value you obtained earlier will no longer apply.</span></span>

<span data-ttu-id="22e1f-p108">Sie erhalten in einem Formular zum Verfassen für diese Eigenschaft für ein neues Element null. Wenn der Benutzer einen Betreff festlegt und das Element speichert, gibt die `conversationId`-Eigenschaft einen Wert zurück.</span><span class="sxs-lookup"><span data-stu-id="22e1f-p108">You get null for this property for a new item in a compose form. If the user sets a subject and saves the item, the `conversationId` property will return a value.</span></span>

##### <a name="type"></a><span data-ttu-id="22e1f-284">Typ:</span><span class="sxs-lookup"><span data-stu-id="22e1f-284">Type:</span></span>

*   <span data-ttu-id="22e1f-285">String</span><span class="sxs-lookup"><span data-stu-id="22e1f-285">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="22e1f-286">Anforderungen</span><span class="sxs-lookup"><span data-stu-id="22e1f-286">Requirements</span></span>

|<span data-ttu-id="22e1f-287">Anforderung</span><span class="sxs-lookup"><span data-stu-id="22e1f-287">Requirement</span></span>|<span data-ttu-id="22e1f-288">Wert</span><span class="sxs-lookup"><span data-stu-id="22e1f-288">Value</span></span>|
|---|---|
|[<span data-ttu-id="22e1f-289">Mindestversion des Postfachanforderungssatzes</span><span class="sxs-lookup"><span data-stu-id="22e1f-289">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="22e1f-290">1.0</span><span class="sxs-lookup"><span data-stu-id="22e1f-290">1.0</span></span>|
|[<span data-ttu-id="22e1f-291">Mindestberechtigungsstufe</span><span class="sxs-lookup"><span data-stu-id="22e1f-291">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="22e1f-292">ReadItem</span><span class="sxs-lookup"><span data-stu-id="22e1f-292">ReadItem</span></span>|
|[<span data-ttu-id="22e1f-293">Zutreffender Outlook-Modus</span><span class="sxs-lookup"><span data-stu-id="22e1f-293">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="22e1f-294">Verfassen oder Lesen</span><span class="sxs-lookup"><span data-stu-id="22e1f-294">Compose or read</span></span>|

#### <a name="datetimecreated-date"></a><span data-ttu-id="22e1f-295">dateTimeCreated :Date</span><span class="sxs-lookup"><span data-stu-id="22e1f-295">dateTimeCreated :Date</span></span>

<span data-ttu-id="22e1f-p109">Ruft das Datum und die Uhrzeit der Erstellung eines Elements ab. Nur Lesemodus.</span><span class="sxs-lookup"><span data-stu-id="22e1f-p109">Gets the date and time that an item was created. Read mode only.</span></span>

##### <a name="type"></a><span data-ttu-id="22e1f-298">Typ:</span><span class="sxs-lookup"><span data-stu-id="22e1f-298">Type:</span></span>

*   <span data-ttu-id="22e1f-299">Datum</span><span class="sxs-lookup"><span data-stu-id="22e1f-299">Date</span></span>

##### <a name="requirements"></a><span data-ttu-id="22e1f-300">Anforderungen</span><span class="sxs-lookup"><span data-stu-id="22e1f-300">Requirements</span></span>

|<span data-ttu-id="22e1f-301">Anforderung</span><span class="sxs-lookup"><span data-stu-id="22e1f-301">Requirement</span></span>|<span data-ttu-id="22e1f-302">Wert</span><span class="sxs-lookup"><span data-stu-id="22e1f-302">Value</span></span>|
|---|---|
|[<span data-ttu-id="22e1f-303">Mindestversion des Postfachanforderungssatzes</span><span class="sxs-lookup"><span data-stu-id="22e1f-303">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="22e1f-304">1.0</span><span class="sxs-lookup"><span data-stu-id="22e1f-304">1.0</span></span>|
|[<span data-ttu-id="22e1f-305">Mindestberechtigungsstufe</span><span class="sxs-lookup"><span data-stu-id="22e1f-305">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="22e1f-306">ReadItem</span><span class="sxs-lookup"><span data-stu-id="22e1f-306">ReadItem</span></span>|
|[<span data-ttu-id="22e1f-307">Zutreffender Outlook-Modus</span><span class="sxs-lookup"><span data-stu-id="22e1f-307">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="22e1f-308">Lesen</span><span class="sxs-lookup"><span data-stu-id="22e1f-308">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="22e1f-309">Beispiel</span><span class="sxs-lookup"><span data-stu-id="22e1f-309">Example</span></span>

```
var created = Office.context.mailbox.item.dateTimeCreated;
```

#### <a name="datetimemodified-date"></a><span data-ttu-id="22e1f-310">dateTimeModified :Date</span><span class="sxs-lookup"><span data-stu-id="22e1f-310">dateTimeModified :Date</span></span>

<span data-ttu-id="22e1f-p110">Ruft das Datum und die Uhrzeit der letzten Änderung eines Elements ab. Nur Lesemodus.</span><span class="sxs-lookup"><span data-stu-id="22e1f-p110">Gets the date and time that an item was last modified. Read mode only.</span></span>

> [!NOTE]
> <span data-ttu-id="22e1f-313">Dieser Member wird in Outlook für iOS oder Outlook für Android nicht unterstützt.</span><span class="sxs-lookup"><span data-stu-id="22e1f-313">This member is not supported in Outlook for iOS or Outlook for Android.</span></span>

##### <a name="type"></a><span data-ttu-id="22e1f-314">Typ:</span><span class="sxs-lookup"><span data-stu-id="22e1f-314">Type:</span></span>

*   <span data-ttu-id="22e1f-315">Datum</span><span class="sxs-lookup"><span data-stu-id="22e1f-315">Date</span></span>

##### <a name="requirements"></a><span data-ttu-id="22e1f-316">Anforderungen</span><span class="sxs-lookup"><span data-stu-id="22e1f-316">Requirements</span></span>

|<span data-ttu-id="22e1f-317">Anforderung</span><span class="sxs-lookup"><span data-stu-id="22e1f-317">Requirement</span></span>|<span data-ttu-id="22e1f-318">Wert</span><span class="sxs-lookup"><span data-stu-id="22e1f-318">Value</span></span>|
|---|---|
|[<span data-ttu-id="22e1f-319">Mindestversion des Postfachanforderungssatzes</span><span class="sxs-lookup"><span data-stu-id="22e1f-319">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="22e1f-320">1.0</span><span class="sxs-lookup"><span data-stu-id="22e1f-320">1.0</span></span>|
|[<span data-ttu-id="22e1f-321">Mindestberechtigungsstufe</span><span class="sxs-lookup"><span data-stu-id="22e1f-321">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="22e1f-322">ReadItem</span><span class="sxs-lookup"><span data-stu-id="22e1f-322">ReadItem</span></span>|
|[<span data-ttu-id="22e1f-323">Zutreffender Outlook-Modus</span><span class="sxs-lookup"><span data-stu-id="22e1f-323">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="22e1f-324">Lesen</span><span class="sxs-lookup"><span data-stu-id="22e1f-324">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="22e1f-325">Beispiel</span><span class="sxs-lookup"><span data-stu-id="22e1f-325">Example</span></span>

```
var modified = Office.context.mailbox.item.dateTimeModified;
```

####  <a name="end-datetimejavascriptapioutlookofficetime"></a><span data-ttu-id="22e1f-326">end :Date|[Time](/javascript/api/outlook/office.time)</span><span class="sxs-lookup"><span data-stu-id="22e1f-326">end :Date|[Time](/javascript/api/outlook/office.time)</span></span>

<span data-ttu-id="22e1f-327">Ruft Datum und Zeit für das Ende des Termins ab oder legt diese fest.</span><span class="sxs-lookup"><span data-stu-id="22e1f-327">Gets or sets the date and time that the appointment is to end.</span></span>

<span data-ttu-id="22e1f-p111">Die `end`-Eigenschaft wird als Datums- und Uhrzeitwert in UTC ausgedrückt. Sie können die [`convertToLocalClientTime`](office.context.mailbox.md#converttolocalclienttimetimevalue--localclienttimejavascriptapioutlookofficelocalclienttime)-Methode verwenden, um den Endeigenschaftswert in den lokalen Uhrzeit- und Datumswert des Clients umzuwandeln.</span><span class="sxs-lookup"><span data-stu-id="22e1f-p111">The `end` property is expressed as a Coordinated Universal Time (UTC) date and time value. You can use the [`convertToLocalClientTime`](office.context.mailbox.md#converttolocalclienttimetimevalue--localclienttimejavascriptapioutlookofficelocalclienttime) method to convert the end property value to the client’s local date and time.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="22e1f-330">Lesemodus</span><span class="sxs-lookup"><span data-stu-id="22e1f-330">Read mode</span></span>

<span data-ttu-id="22e1f-331">Die `end`-Eigenschaft gibt ein `Date`-Objekt zurück.</span><span class="sxs-lookup"><span data-stu-id="22e1f-331">The `end` property returns a `Date` object.</span></span>

##### <a name="compose-mode"></a><span data-ttu-id="22e1f-332">Verfassenmodus</span><span class="sxs-lookup"><span data-stu-id="22e1f-332">Compose mode</span></span>

<span data-ttu-id="22e1f-333">Die `end`-Eigenschaft gibt ein `Time`-Objekt zurück.</span><span class="sxs-lookup"><span data-stu-id="22e1f-333">The `end` property returns a `Time` object.</span></span>

<span data-ttu-id="22e1f-334">Wenn Sie die [`Time.setAsync`](/javascript/api/outlook/office.time#setasync-datetime--options--callback-)-Methode verwenden, um die Endzeit im Verfassenmodus festzulegen, sollten Sie die [`convertToUtcClientTime`](office.context.mailbox.md#converttoutcclienttimeinput--date)-Methode verwenden, um die Ortszeit auf dem Client für den Server in UTC umzuwandeln.</span><span class="sxs-lookup"><span data-stu-id="22e1f-334">When you use the [`Time.setAsync`](/javascript/api/outlook/office.time#setasync-datetime--options--callback-) method to set the end time, you should use the [`convertToUtcClientTime`](office.context.mailbox.md#converttoutcclienttimeinput--date) method to convert the local time on the client to UTC for the server.</span></span>

##### <a name="type"></a><span data-ttu-id="22e1f-335">Typ:</span><span class="sxs-lookup"><span data-stu-id="22e1f-335">Type:</span></span>

*   <span data-ttu-id="22e1f-336">Date | [Time](/javascript/api/outlook/office.time)</span><span class="sxs-lookup"><span data-stu-id="22e1f-336">Date | [Time](/javascript/api/outlook/office.time)</span></span>

##### <a name="requirements"></a><span data-ttu-id="22e1f-337">Anforderungen</span><span class="sxs-lookup"><span data-stu-id="22e1f-337">Requirements</span></span>

|<span data-ttu-id="22e1f-338">Anforderung</span><span class="sxs-lookup"><span data-stu-id="22e1f-338">Requirement</span></span>|<span data-ttu-id="22e1f-339">Wert</span><span class="sxs-lookup"><span data-stu-id="22e1f-339">Value</span></span>|
|---|---|
|[<span data-ttu-id="22e1f-340">Mindestversion des Postfachanforderungssatzes</span><span class="sxs-lookup"><span data-stu-id="22e1f-340">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="22e1f-341">1.0</span><span class="sxs-lookup"><span data-stu-id="22e1f-341">1.0</span></span>|
|[<span data-ttu-id="22e1f-342">Mindestberechtigungsstufe</span><span class="sxs-lookup"><span data-stu-id="22e1f-342">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="22e1f-343">ReadItem</span><span class="sxs-lookup"><span data-stu-id="22e1f-343">ReadItem</span></span>|
|[<span data-ttu-id="22e1f-344">Zutreffender Outlook-Modus</span><span class="sxs-lookup"><span data-stu-id="22e1f-344">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="22e1f-345">Verfassen oder Lesen</span><span class="sxs-lookup"><span data-stu-id="22e1f-345">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="22e1f-346">Beispiel</span><span class="sxs-lookup"><span data-stu-id="22e1f-346">Example</span></span>

<span data-ttu-id="22e1f-347">Im folgenden Beispiel wird die Endzeit eines Termins im Verfassenmodus mithilfe der [`setAsync`](/javascript/api/outlook/office.time#setasync-datetime--options--callback-)-Methode des `Time`-Objekts festgelegt.</span><span class="sxs-lookup"><span data-stu-id="22e1f-347">The following example sets the end time of an appointment in compose mode by using the [`setAsync`](/javascript/api/outlook/office.time#setasync-datetime--options--callback-) method of the `Time` object.</span></span>

```
var endTime = new Date("3/14/2015");
var options = {
  // Pass information that can be used
  // in the callback
     asyncContext: {verb:"Set"}
}
Office.context.mailbox.item.end.setAsync(endTime, options, function(result) {
  if (result.error) {
    console.debug(result.error);
  } else {
    // Access the asyncContext that was passed to the setAsync function
    console.debug("End Time " + result.asyncContext.verb);
  }
});
```

#### <a name="from-emailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsfromjavascriptapioutlookofficefrom"></a><span data-ttu-id="22e1f-348">von:[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)|[aus](/javascript/api/outlook/office.from)</span><span class="sxs-lookup"><span data-stu-id="22e1f-348">from :[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)|[From](/javascript/api/outlook/office.from)</span></span>

<span data-ttu-id="22e1f-349">Ruft die E-Mail-Adresse des Absenders einer Nachricht ab.</span><span class="sxs-lookup"><span data-stu-id="22e1f-349">Gets the email address of the sender of a message.</span></span>

<span data-ttu-id="22e1f-p112">Die `from`- und [`sender`](#sender-emailaddressdetailsjavascriptapioutlookofficeemailaddressdetails)-Eigenschaften stellen dieselbe Person dar, außer die Nachricht wurde von einem Delegaten gesendet. In diesem Fall stellt die `from`-Eigenschaft die Stellvertretung dar, und die sender-Eigenschaft stellt den Delegaten dar.</span><span class="sxs-lookup"><span data-stu-id="22e1f-p112">The `from` and [`sender`](#sender-emailaddressdetailsjavascriptapioutlookofficeemailaddressdetails) properties represent the same person unless the message is sent by a delegate. In that case, the `from` property represents the delegator, and the sender property represents the delegate.</span></span>

> [!NOTE]
> <span data-ttu-id="22e1f-352">Die `recipientType` -Eigenschaft des der `EmailAddressDetails` -Objekts der `from` -Eigenschaft ist `undefined`.</span><span class="sxs-lookup"><span data-stu-id="22e1f-352">The `recipientType` property of the `EmailAddressDetails` object in the `from` property is `undefined`.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="22e1f-353">Lesemodus</span><span class="sxs-lookup"><span data-stu-id="22e1f-353">Read mode</span></span>

<span data-ttu-id="22e1f-354">Die `from` -Eigenschaft gibt eine `EmailAddressDetails` Objekt.</span><span class="sxs-lookup"><span data-stu-id="22e1f-354">The `from` property returns an `EmailAddressDetails` object.</span></span>

```
var subject = Office.context.mailbox.item.from;
```

##### <a name="compose-mode"></a><span data-ttu-id="22e1f-355">Verfassenmodus</span><span class="sxs-lookup"><span data-stu-id="22e1f-355">Compose mode</span></span>

<span data-ttu-id="22e1f-356">Die `from` -Eigenschaft gibt eine `From` -Objekt, eine Methode zum Abrufen bietet, der vom Wert.</span><span class="sxs-lookup"><span data-stu-id="22e1f-356">The `from` property returns a `From` object that provides a method to get the from value.</span></span>

```
Office.context.mailbox.item.from.getAsync(callback);

function callback(asyncResult) {
  var from = asyncResult.value;
}
```

##### <a name="type"></a><span data-ttu-id="22e1f-357">Typ:</span><span class="sxs-lookup"><span data-stu-id="22e1f-357">Type:</span></span>

*   <span data-ttu-id="22e1f-358">[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails) | [aus](/javascript/api/outlook/office.from)</span><span class="sxs-lookup"><span data-stu-id="22e1f-358">[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails) | [From](/javascript/api/outlook/office.from)</span></span>

##### <a name="requirements"></a><span data-ttu-id="22e1f-359">Anforderungen</span><span class="sxs-lookup"><span data-stu-id="22e1f-359">Requirements</span></span>

|<span data-ttu-id="22e1f-360">Anforderung</span><span class="sxs-lookup"><span data-stu-id="22e1f-360">Requirement</span></span>|||
|---|---|---|
|[<span data-ttu-id="22e1f-361">Mindestversion des Postfachanforderungssatzes</span><span class="sxs-lookup"><span data-stu-id="22e1f-361">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="22e1f-362">1.0</span><span class="sxs-lookup"><span data-stu-id="22e1f-362">1.0</span></span>|<span data-ttu-id="22e1f-363">Vorschau</span><span class="sxs-lookup"><span data-stu-id="22e1f-363">Preview</span></span>|
|[<span data-ttu-id="22e1f-364">Mindestberechtigungsstufe</span><span class="sxs-lookup"><span data-stu-id="22e1f-364">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="22e1f-365">ReadItem</span><span class="sxs-lookup"><span data-stu-id="22e1f-365">ReadItem</span></span>|<span data-ttu-id="22e1f-366">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="22e1f-366">ReadWriteItem</span></span>|
|[<span data-ttu-id="22e1f-367">Zutreffender Outlook-Modus</span><span class="sxs-lookup"><span data-stu-id="22e1f-367">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="22e1f-368">Lesen</span><span class="sxs-lookup"><span data-stu-id="22e1f-368">Read</span></span>|<span data-ttu-id="22e1f-369">Verfassen</span><span class="sxs-lookup"><span data-stu-id="22e1f-369">Compose</span></span>|

#### <a name="internetmessageid-string"></a><span data-ttu-id="22e1f-370">internetMessageId :String</span><span class="sxs-lookup"><span data-stu-id="22e1f-370">internetMessageId :String</span></span>

<span data-ttu-id="22e1f-p113">Ruft die Internetnachricht-ID einer E-Mail-Nachricht ab. Nur Lesemodus.</span><span class="sxs-lookup"><span data-stu-id="22e1f-p113">Gets the Internet message identifier for an email message. Read mode only.</span></span>

##### <a name="type"></a><span data-ttu-id="22e1f-373">Typ:</span><span class="sxs-lookup"><span data-stu-id="22e1f-373">Type:</span></span>

*   <span data-ttu-id="22e1f-374">String</span><span class="sxs-lookup"><span data-stu-id="22e1f-374">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="22e1f-375">Anforderungen</span><span class="sxs-lookup"><span data-stu-id="22e1f-375">Requirements</span></span>

|<span data-ttu-id="22e1f-376">Anforderung</span><span class="sxs-lookup"><span data-stu-id="22e1f-376">Requirement</span></span>|<span data-ttu-id="22e1f-377">Wert</span><span class="sxs-lookup"><span data-stu-id="22e1f-377">Value</span></span>|
|---|---|
|[<span data-ttu-id="22e1f-378">Mindestversion des Postfachanforderungssatzes</span><span class="sxs-lookup"><span data-stu-id="22e1f-378">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="22e1f-379">1.0</span><span class="sxs-lookup"><span data-stu-id="22e1f-379">1.0</span></span>|
|[<span data-ttu-id="22e1f-380">Mindestberechtigungsstufe</span><span class="sxs-lookup"><span data-stu-id="22e1f-380">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="22e1f-381">ReadItem</span><span class="sxs-lookup"><span data-stu-id="22e1f-381">ReadItem</span></span>|
|[<span data-ttu-id="22e1f-382">Zutreffender Outlook-Modus</span><span class="sxs-lookup"><span data-stu-id="22e1f-382">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="22e1f-383">Lesen</span><span class="sxs-lookup"><span data-stu-id="22e1f-383">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="22e1f-384">Beispiel</span><span class="sxs-lookup"><span data-stu-id="22e1f-384">Example</span></span>

```
var internetMessageId = Office.context.mailbox.item.internetMessageId;
```

#### <a name="itemclass-string"></a><span data-ttu-id="22e1f-385">itemClass :String</span><span class="sxs-lookup"><span data-stu-id="22e1f-385">itemClass :String</span></span>

<span data-ttu-id="22e1f-p114">Ruft die Exchange-Webdienste-Elementklasse des ausgewählten Elements ab. Nur Lesemodus.</span><span class="sxs-lookup"><span data-stu-id="22e1f-p114">Gets the Exchange Web Services item class of the selected item. Read mode only.</span></span>

<span data-ttu-id="22e1f-p115">Die `itemClass`-Eigenschaft gibt die Nachrichtenklasse des ausgewählten Elements an. Folgende Standardnachrichtenklassen für das Nachrichten- oder Terminelement sind vorhanden:</span><span class="sxs-lookup"><span data-stu-id="22e1f-p115">The `itemClass` property specifies the message class of the selected item. The following are the default message classes for the message or appointment item.</span></span>

|<span data-ttu-id="22e1f-390">Typ</span><span class="sxs-lookup"><span data-stu-id="22e1f-390">Type</span></span>|<span data-ttu-id="22e1f-391">Beschreibung</span><span class="sxs-lookup"><span data-stu-id="22e1f-391">Description</span></span>|<span data-ttu-id="22e1f-392">Elementklasse</span><span class="sxs-lookup"><span data-stu-id="22e1f-392">item class</span></span>|
|---|---|---|
|<span data-ttu-id="22e1f-393">Terminelemente</span><span class="sxs-lookup"><span data-stu-id="22e1f-393">Appointment items</span></span>|<span data-ttu-id="22e1f-394">Dies sind Kalenderelemente der Elementklasse `IPM.Appointment` oder `IPM.Appointment.Occurence`.</span><span class="sxs-lookup"><span data-stu-id="22e1f-394">These are calendar items of the item class `IPM.Appointment` or `IPM.Appointment.Occurence`.</span></span>|`IPM.Appointment`<br />`IPM.Appointment.Occurence`|
|<span data-ttu-id="22e1f-395">Nachrichtenelemente</span><span class="sxs-lookup"><span data-stu-id="22e1f-395">Message items</span></span>|<span data-ttu-id="22e1f-396">Diese Elemente umfassen E-Mail-Nachrichten der Standardnachrichtenklasse `IPM.Note` sowie Besprechungsanfragen, -antworten und -absagen, die `IPM.Schedule.Meeting` als Basisnachrichtenklasse verwenden.</span><span class="sxs-lookup"><span data-stu-id="22e1f-396">These include email messages that have the default message class `IPM.Note`, and meeting requests, responses, and cancellations, that use `IPM.Schedule.Meeting` as the base message class.</span></span>|`IPM.Note`<br />`IPM.Schedule.Meeting.Request`<br />`IPM.Schedule.Meeting.Neg`<br />`IPM.Schedule.Meeting.Pos`<br />`IPM.Schedule.Meeting.Tent`<br />`IPM.Schedule.Meeting.Canceled`|

<span data-ttu-id="22e1f-397">Sie können benutzerdefinierte Nachrichtenklassen zum Erweitern einer Standardnachrichtenklasse erstellen, z. B. eine benutzerdefinierte Terminnachrichtenklasse wie `IPM.Appointment.Contoso`.</span><span class="sxs-lookup"><span data-stu-id="22e1f-397">You can create custom message classes that extends a default message class, for example, a custom appointment message class `IPM.Appointment.Contoso`.</span></span>

##### <a name="type"></a><span data-ttu-id="22e1f-398">Typ:</span><span class="sxs-lookup"><span data-stu-id="22e1f-398">Type:</span></span>

*   <span data-ttu-id="22e1f-399">String</span><span class="sxs-lookup"><span data-stu-id="22e1f-399">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="22e1f-400">Anforderungen</span><span class="sxs-lookup"><span data-stu-id="22e1f-400">Requirements</span></span>

|<span data-ttu-id="22e1f-401">Anforderung</span><span class="sxs-lookup"><span data-stu-id="22e1f-401">Requirement</span></span>|<span data-ttu-id="22e1f-402">Wert</span><span class="sxs-lookup"><span data-stu-id="22e1f-402">Value</span></span>|
|---|---|
|[<span data-ttu-id="22e1f-403">Mindestversion des Postfachanforderungssatzes</span><span class="sxs-lookup"><span data-stu-id="22e1f-403">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="22e1f-404">1.0</span><span class="sxs-lookup"><span data-stu-id="22e1f-404">1.0</span></span>|
|[<span data-ttu-id="22e1f-405">Mindestberechtigungsstufe</span><span class="sxs-lookup"><span data-stu-id="22e1f-405">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="22e1f-406">ReadItem</span><span class="sxs-lookup"><span data-stu-id="22e1f-406">ReadItem</span></span>|
|[<span data-ttu-id="22e1f-407">Zutreffender Outlook-Modus</span><span class="sxs-lookup"><span data-stu-id="22e1f-407">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="22e1f-408">Lesen</span><span class="sxs-lookup"><span data-stu-id="22e1f-408">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="22e1f-409">Beispiel</span><span class="sxs-lookup"><span data-stu-id="22e1f-409">Example</span></span>

```
var itemClass = Office.context.mailbox.item.itemClass;
```

#### <a name="nullable-itemid-string"></a><span data-ttu-id="22e1f-410">(nullable) itemId :String</span><span class="sxs-lookup"><span data-stu-id="22e1f-410">(nullable) itemId :String</span></span>

<span data-ttu-id="22e1f-p116">Ruft den Exchange-Webdienste-Elementbezeichner für das aktuelle Element ab. Nur Lesemodus.</span><span class="sxs-lookup"><span data-stu-id="22e1f-p116">Gets the Exchange Web Services item identifier for the current item. Read mode only.</span></span>

> [!NOTE]
> <span data-ttu-id="22e1f-413">Der Bezeichner, der von der `itemId`-Eigenschaft zurückgegeben wird, ist der gleiche Bezeichner wie der Exchange-Webdienste-Elementbezeichner.</span><span class="sxs-lookup"><span data-stu-id="22e1f-413">The identifier returned by the `itemId` property is the same as the Exchange Web Services item identifier.</span></span> <span data-ttu-id="22e1f-414">Die `itemId` -Eigenschaft ist nicht identisch mit der Outlook-Eintrags-ID oder die ID von der Outlook-REST-API verwendet.</span><span class="sxs-lookup"><span data-stu-id="22e1f-414">The `itemId` property is not identical to the Outlook Entry ID or the ID used by the Outlook REST API.</span></span> <span data-ttu-id="22e1f-415">REST-API-Aufrufe dieser Wert verwenden, sollten sie konvertiert werden, bevor [Office.context.mailbox.convertToRestId](office.context.mailbox.md#converttorestiditemid-restversion--string)verwenden.</span><span class="sxs-lookup"><span data-stu-id="22e1f-415">Before making REST API calls using this value, it should be converted using [Office.context.mailbox.convertToRestId](office.context.mailbox.md#converttorestiditemid-restversion--string).</span></span> <span data-ttu-id="22e1f-416">Weitere Informationen finden Sie unter [Verwenden der Outlook-REST-APIs aus einem Outlook-add-in](https://docs.microsoft.com/outlook/add-ins/use-rest-api#get-the-item-id).</span><span class="sxs-lookup"><span data-stu-id="22e1f-416">For more details, see [Use the Outlook REST APIs from an Outlook add-in](https://docs.microsoft.com/outlook/add-ins/use-rest-api#get-the-item-id).</span></span>

<span data-ttu-id="22e1f-p118">Die Eigenschaft `itemId` ist im Modus „Verfassen“ nicht verfügbar. Wenn ein Elementbezeichner erforderlich ist, kann die [`saveAsync`](#saveasyncoptions-callback)-Methode verwendet werden, um das Element zu speichern. Dadurch wird der Elementbezeichner im [`AsyncResult.value`](/javascript/api/office/office.asyncresult)-Parameter in der Rückruffunktion zurückgegeben.</span><span class="sxs-lookup"><span data-stu-id="22e1f-p118">The `itemId` property is not available in compose mode. If an item identifier is required, the [`saveAsync`](#saveasyncoptions-callback) method can be used to save the item to the store, which will return the item identifier in the [`AsyncResult.value`](/javascript/api/office/office.asyncresult) parameter in the callback function.</span></span>

##### <a name="type"></a><span data-ttu-id="22e1f-419">Typ:</span><span class="sxs-lookup"><span data-stu-id="22e1f-419">Type:</span></span>

*   <span data-ttu-id="22e1f-420">String</span><span class="sxs-lookup"><span data-stu-id="22e1f-420">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="22e1f-421">Anforderungen</span><span class="sxs-lookup"><span data-stu-id="22e1f-421">Requirements</span></span>

|<span data-ttu-id="22e1f-422">Anforderung</span><span class="sxs-lookup"><span data-stu-id="22e1f-422">Requirement</span></span>|<span data-ttu-id="22e1f-423">Wert</span><span class="sxs-lookup"><span data-stu-id="22e1f-423">Value</span></span>|
|---|---|
|[<span data-ttu-id="22e1f-424">Mindestversion des Postfachanforderungssatzes</span><span class="sxs-lookup"><span data-stu-id="22e1f-424">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="22e1f-425">1.0</span><span class="sxs-lookup"><span data-stu-id="22e1f-425">1.0</span></span>|
|[<span data-ttu-id="22e1f-426">Mindestberechtigungsstufe</span><span class="sxs-lookup"><span data-stu-id="22e1f-426">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="22e1f-427">ReadItem</span><span class="sxs-lookup"><span data-stu-id="22e1f-427">ReadItem</span></span>|
|[<span data-ttu-id="22e1f-428">Zutreffender Outlook-Modus</span><span class="sxs-lookup"><span data-stu-id="22e1f-428">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="22e1f-429">Lesen</span><span class="sxs-lookup"><span data-stu-id="22e1f-429">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="22e1f-430">Beispiel</span><span class="sxs-lookup"><span data-stu-id="22e1f-430">Example</span></span>

<span data-ttu-id="22e1f-p119">Mit dem folgende Code wird das Vorhandensein eines Elementbezeichners überprüft. Wenn die `itemId`-Eigenschaft `null` oder `undefined` zurückgibt, speichert sie das Element und ruft den Elementbezeichner aus dem asynchronen Ergebnis ab.</span><span class="sxs-lookup"><span data-stu-id="22e1f-p119">The following code checks for the presence of an item identifier. If the `itemId` property returns `null` or `undefined`, it saves the item to the store and gets the item identifier from the asynchronous result.</span></span>

```
var itemId = Office.context.mailbox.item.itemId;
if (itemId === null || itemId == undefined) {
  Office.context.mailbox.item.saveAsync(function(result){
    itemId = result.value;
  });
}
```

####  <a name="itemtype-officemailboxenumsitemtypejavascriptapioutlookofficemailboxenumsitemtype"></a><span data-ttu-id="22e1f-433">itemType :[Office.MailboxEnums.ItemType](/javascript/api/outlook/office.mailboxenums.itemtype)</span><span class="sxs-lookup"><span data-stu-id="22e1f-433">itemType :[Office.MailboxEnums.ItemType](/javascript/api/outlook/office.mailboxenums.itemtype)</span></span>

<span data-ttu-id="22e1f-434">Ruft den Typ des Elements ab, das eine Instanz darstellt.</span><span class="sxs-lookup"><span data-stu-id="22e1f-434">Gets the type of item that an instance represents.</span></span>

<span data-ttu-id="22e1f-435">Die `itemType`-Eigenschaft gibt einen der Werte der `ItemType`-Enumeration zurück, der angibt, ob es sich bei der `item`-Objektinstanz um eine Nachricht oder einen Termin handelt.</span><span class="sxs-lookup"><span data-stu-id="22e1f-435">The `itemType` property returns one of the `ItemType` enumeration values, indicating whether the `item` object instance is a message or an appointment.</span></span>

##### <a name="type"></a><span data-ttu-id="22e1f-436">Typ:</span><span class="sxs-lookup"><span data-stu-id="22e1f-436">Type:</span></span>

*   [<span data-ttu-id="22e1f-437">Office.MailboxEnums.ItemType</span><span class="sxs-lookup"><span data-stu-id="22e1f-437">Office.MailboxEnums.ItemType</span></span>](/javascript/api/outlook/office.mailboxenums.itemtype)

##### <a name="requirements"></a><span data-ttu-id="22e1f-438">Anforderungen</span><span class="sxs-lookup"><span data-stu-id="22e1f-438">Requirements</span></span>

|<span data-ttu-id="22e1f-439">Anforderung</span><span class="sxs-lookup"><span data-stu-id="22e1f-439">Requirement</span></span>|<span data-ttu-id="22e1f-440">Wert</span><span class="sxs-lookup"><span data-stu-id="22e1f-440">Value</span></span>|
|---|---|
|[<span data-ttu-id="22e1f-441">Mindestversion des Postfachanforderungssatzes</span><span class="sxs-lookup"><span data-stu-id="22e1f-441">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="22e1f-442">1.0</span><span class="sxs-lookup"><span data-stu-id="22e1f-442">1.0</span></span>|
|[<span data-ttu-id="22e1f-443">Mindestberechtigungsstufe</span><span class="sxs-lookup"><span data-stu-id="22e1f-443">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="22e1f-444">ReadItem</span><span class="sxs-lookup"><span data-stu-id="22e1f-444">ReadItem</span></span>|
|[<span data-ttu-id="22e1f-445">Zutreffender Outlook-Modus</span><span class="sxs-lookup"><span data-stu-id="22e1f-445">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="22e1f-446">Verfassen oder Lesen</span><span class="sxs-lookup"><span data-stu-id="22e1f-446">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="22e1f-447">Beispiel</span><span class="sxs-lookup"><span data-stu-id="22e1f-447">Example</span></span>

```
if (Office.context.mailbox.item.itemType == Office.MailboxEnums.ItemType.Message)
  // do something
else
  // do something else
```

####  <a name="location-stringlocationjavascriptapioutlookofficelocation"></a><span data-ttu-id="22e1f-448">location :String|[Location](/javascript/api/outlook/office.location)</span><span class="sxs-lookup"><span data-stu-id="22e1f-448">location :String|[Location](/javascript/api/outlook/office.location)</span></span>

<span data-ttu-id="22e1f-449">Ruft den Ort eines Termins ab bzw. legt ihn fest.</span><span class="sxs-lookup"><span data-stu-id="22e1f-449">Gets or sets the location of an appointment.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="22e1f-450">Lesemodus</span><span class="sxs-lookup"><span data-stu-id="22e1f-450">Read mode</span></span>

<span data-ttu-id="22e1f-451">Die `location`-Eigenschaft gibt eine Zeichenfolge zurück, die den Ort des Termins enthält.</span><span class="sxs-lookup"><span data-stu-id="22e1f-451">The `location` property returns a string that contains the location of the appointment.</span></span>

##### <a name="compose-mode"></a><span data-ttu-id="22e1f-452">Verfassenmodus</span><span class="sxs-lookup"><span data-stu-id="22e1f-452">Compose mode</span></span>

<span data-ttu-id="22e1f-453">Die `location`-Eigenschaft gibt ein `Location`-Objekt zurück, das Methoden zum Abrufen und Festlegen des Orts für einen Termin bereitstellt.</span><span class="sxs-lookup"><span data-stu-id="22e1f-453">The `location` property returns a `Location` object that provides methods that are used to get and set the location of the appointment.</span></span>

##### <a name="type"></a><span data-ttu-id="22e1f-454">Typ:</span><span class="sxs-lookup"><span data-stu-id="22e1f-454">Type:</span></span>

*   <span data-ttu-id="22e1f-455">String | [Location](/javascript/api/outlook/office.location)</span><span class="sxs-lookup"><span data-stu-id="22e1f-455">String | [Location](/javascript/api/outlook/office.location)</span></span>

##### <a name="requirements"></a><span data-ttu-id="22e1f-456">Anforderungen</span><span class="sxs-lookup"><span data-stu-id="22e1f-456">Requirements</span></span>

|<span data-ttu-id="22e1f-457">Anforderung</span><span class="sxs-lookup"><span data-stu-id="22e1f-457">Requirement</span></span>|<span data-ttu-id="22e1f-458">Wert</span><span class="sxs-lookup"><span data-stu-id="22e1f-458">Value</span></span>|
|---|---|
|[<span data-ttu-id="22e1f-459">Mindestversion des Postfachanforderungssatzes</span><span class="sxs-lookup"><span data-stu-id="22e1f-459">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="22e1f-460">1.0</span><span class="sxs-lookup"><span data-stu-id="22e1f-460">1.0</span></span>|
|[<span data-ttu-id="22e1f-461">Mindestberechtigungsstufe</span><span class="sxs-lookup"><span data-stu-id="22e1f-461">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="22e1f-462">ReadItem</span><span class="sxs-lookup"><span data-stu-id="22e1f-462">ReadItem</span></span>|
|[<span data-ttu-id="22e1f-463">Zutreffender Outlook-Modus</span><span class="sxs-lookup"><span data-stu-id="22e1f-463">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="22e1f-464">Verfassen oder Lesen</span><span class="sxs-lookup"><span data-stu-id="22e1f-464">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="22e1f-465">Beispiel</span><span class="sxs-lookup"><span data-stu-id="22e1f-465">Example</span></span>

```
var userContext = { value : 1 };
Office.context.mailbox.item.location.getAsync( { context: userContext}, callback);

function callback(asyncResult) {
  var context = asyncResult.context;
  var location = asyncResult.value;
}
```

#### <a name="normalizedsubject-string"></a><span data-ttu-id="22e1f-466">normalizedSubject :String</span><span class="sxs-lookup"><span data-stu-id="22e1f-466">normalizedSubject :String</span></span>

<span data-ttu-id="22e1f-p120">Ruft den Betreff für ein Element ab, wobei alle Präfixe entfernt werden (einschließlich `RE:` und `FWD:`). Nur Lesemodus.</span><span class="sxs-lookup"><span data-stu-id="22e1f-p120">Gets the subject of an item, with all prefixes removed (including `RE:` and `FWD:`). Read mode only.</span></span>

<span data-ttu-id="22e1f-p121">Die normalizedSubject-Eigenschaft ruft den Betreff des Elements mit allen Standardpräfixen (z. B. `RE:` und `FW:`) ab, die von E-Mail-Programmen hinzugefügt werden. Wenn der Betreff des Elements mit vollständigen Präfixen abgerufen werden soll, verwenden Sie die [`subject`](#subject-stringsubjectjavascriptapioutlookofficesubject)-Eigenschaft.</span><span class="sxs-lookup"><span data-stu-id="22e1f-p121">The normalizedSubject property gets the subject of the item, with any standard prefixes (such as `RE:` and `FW:`) that are added by email programs. To get the subject of the item with the prefixes intact, use the [`subject`](#subject-stringsubjectjavascriptapioutlookofficesubject) property.</span></span>

##### <a name="type"></a><span data-ttu-id="22e1f-471">Typ:</span><span class="sxs-lookup"><span data-stu-id="22e1f-471">Type:</span></span>

*   <span data-ttu-id="22e1f-472">String</span><span class="sxs-lookup"><span data-stu-id="22e1f-472">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="22e1f-473">Anforderungen</span><span class="sxs-lookup"><span data-stu-id="22e1f-473">Requirements</span></span>

|<span data-ttu-id="22e1f-474">Anforderung</span><span class="sxs-lookup"><span data-stu-id="22e1f-474">Requirement</span></span>|<span data-ttu-id="22e1f-475">Wert</span><span class="sxs-lookup"><span data-stu-id="22e1f-475">Value</span></span>|
|---|---|
|[<span data-ttu-id="22e1f-476">Mindestversion des Postfachanforderungssatzes</span><span class="sxs-lookup"><span data-stu-id="22e1f-476">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="22e1f-477">1.0</span><span class="sxs-lookup"><span data-stu-id="22e1f-477">1.0</span></span>|
|[<span data-ttu-id="22e1f-478">Mindestberechtigungsstufe</span><span class="sxs-lookup"><span data-stu-id="22e1f-478">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="22e1f-479">ReadItem</span><span class="sxs-lookup"><span data-stu-id="22e1f-479">ReadItem</span></span>|
|[<span data-ttu-id="22e1f-480">Zutreffender Outlook-Modus</span><span class="sxs-lookup"><span data-stu-id="22e1f-480">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="22e1f-481">Lesen</span><span class="sxs-lookup"><span data-stu-id="22e1f-481">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="22e1f-482">Beispiel</span><span class="sxs-lookup"><span data-stu-id="22e1f-482">Example</span></span>

```
var normalizedSubject = Office.context.mailbox.item.normalizedSubject;
```

####  <a name="notificationmessages-notificationmessagesjavascriptapioutlookofficenotificationmessages"></a><span data-ttu-id="22e1f-483">notificationMessages :[NotificationMessages](/javascript/api/outlook/office.notificationmessages)</span><span class="sxs-lookup"><span data-stu-id="22e1f-483">notificationMessages :[NotificationMessages](/javascript/api/outlook/office.notificationmessages)</span></span>

<span data-ttu-id="22e1f-484">Ruft die Benachrichtigungen für ein Element ab.</span><span class="sxs-lookup"><span data-stu-id="22e1f-484">Gets the notification messages for an item.</span></span>

##### <a name="type"></a><span data-ttu-id="22e1f-485">Typ:</span><span class="sxs-lookup"><span data-stu-id="22e1f-485">Type:</span></span>

*   [<span data-ttu-id="22e1f-486">NotificationMessages</span><span class="sxs-lookup"><span data-stu-id="22e1f-486">NotificationMessages</span></span>](/javascript/api/outlook/office.notificationmessages)

##### <a name="requirements"></a><span data-ttu-id="22e1f-487">Anforderungen</span><span class="sxs-lookup"><span data-stu-id="22e1f-487">Requirements</span></span>

|<span data-ttu-id="22e1f-488">Anforderung</span><span class="sxs-lookup"><span data-stu-id="22e1f-488">Requirement</span></span>|<span data-ttu-id="22e1f-489">Wert</span><span class="sxs-lookup"><span data-stu-id="22e1f-489">Value</span></span>|
|---|---|
|[<span data-ttu-id="22e1f-490">Mindestversion des Postfachanforderungssatzes</span><span class="sxs-lookup"><span data-stu-id="22e1f-490">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="22e1f-491">1.3</span><span class="sxs-lookup"><span data-stu-id="22e1f-491">1.3</span></span>|
|[<span data-ttu-id="22e1f-492">Mindestberechtigungsstufe</span><span class="sxs-lookup"><span data-stu-id="22e1f-492">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="22e1f-493">ReadItem</span><span class="sxs-lookup"><span data-stu-id="22e1f-493">ReadItem</span></span>|
|[<span data-ttu-id="22e1f-494">Zutreffender Outlook-Modus</span><span class="sxs-lookup"><span data-stu-id="22e1f-494">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="22e1f-495">Verfassen oder Lesen</span><span class="sxs-lookup"><span data-stu-id="22e1f-495">Compose or read</span></span>|

####  <a name="optionalattendees-arrayemailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsrecipientsjavascriptapioutlookofficerecipients"></a><span data-ttu-id="22e1f-496">optionalAttendees :Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="22e1f-496">optionalAttendees :Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook/office.recipients)</span></span>

<span data-ttu-id="22e1f-497">Ermöglicht den Zugriff auf die optionalen Teilnehmer eines Ereignisses.</span><span class="sxs-lookup"><span data-stu-id="22e1f-497">Provides access to the optional attendees of an event.</span></span> <span data-ttu-id="22e1f-498">Der Typ des Objekts und die Zugriffsebene, hängt von den Modus des aktuellen Elements ab.</span><span class="sxs-lookup"><span data-stu-id="22e1f-498">The type of object and level of access depends on the mode of the current item.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="22e1f-499">Lesemodus</span><span class="sxs-lookup"><span data-stu-id="22e1f-499">Read mode</span></span>

<span data-ttu-id="22e1f-500">Die `optionalAttendees`-Eigenschaft gibt ein Array mit einem `EmailAddressDetails`-Objekt für jeden optionalen Teilnehmer an der Besprechung zurück.</span><span class="sxs-lookup"><span data-stu-id="22e1f-500">The `optionalAttendees` property returns an array that contains an `EmailAddressDetails` object for each optional attendee to the meeting.</span></span>

##### <a name="compose-mode"></a><span data-ttu-id="22e1f-501">Verfassenmodus</span><span class="sxs-lookup"><span data-stu-id="22e1f-501">Compose mode</span></span>

<span data-ttu-id="22e1f-502">Die `optionalAttendees` -Eigenschaft gibt eine `Recipients` -Objekt, das Methoden zum Abrufen oder Aktualisieren der optionalen Teilnehmer für eine Besprechung bereitstellt.</span><span class="sxs-lookup"><span data-stu-id="22e1f-502">The `optionalAttendees` property returns a `Recipients` object that provides methods to get or update the optional attendees for a meeting.</span></span>

##### <a name="type"></a><span data-ttu-id="22e1f-503">Typ:</span><span class="sxs-lookup"><span data-stu-id="22e1f-503">Type:</span></span>

*   <span data-ttu-id="22e1f-504">Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="22e1f-504">Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook/office.recipients)</span></span>

##### <a name="requirements"></a><span data-ttu-id="22e1f-505">Anforderungen</span><span class="sxs-lookup"><span data-stu-id="22e1f-505">Requirements</span></span>

|<span data-ttu-id="22e1f-506">Anforderung</span><span class="sxs-lookup"><span data-stu-id="22e1f-506">Requirement</span></span>|<span data-ttu-id="22e1f-507">Wert</span><span class="sxs-lookup"><span data-stu-id="22e1f-507">Value</span></span>|
|---|---|
|[<span data-ttu-id="22e1f-508">Mindestversion des Postfachanforderungssatzes</span><span class="sxs-lookup"><span data-stu-id="22e1f-508">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="22e1f-509">1.0</span><span class="sxs-lookup"><span data-stu-id="22e1f-509">1.0</span></span>|
|[<span data-ttu-id="22e1f-510">Mindestberechtigungsstufe</span><span class="sxs-lookup"><span data-stu-id="22e1f-510">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="22e1f-511">ReadItem</span><span class="sxs-lookup"><span data-stu-id="22e1f-511">ReadItem</span></span>|
|[<span data-ttu-id="22e1f-512">Zutreffender Outlook-Modus</span><span class="sxs-lookup"><span data-stu-id="22e1f-512">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="22e1f-513">Verfassen oder Lesen</span><span class="sxs-lookup"><span data-stu-id="22e1f-513">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="22e1f-514">Beispiel</span><span class="sxs-lookup"><span data-stu-id="22e1f-514">Example</span></span>

```
Office.context.mailbox.item.optionalAttendees.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.optionalAttendees.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.optionalAttendees.getAsync(callback);

function callback(asyncResult) {
  var arrayOfOptionalAttendeesRecipients = asyncResult.value;
}
```

#### <a name="organizer-emailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsorganizerjavascriptapioutlookofficeorganizer"></a><span data-ttu-id="22e1f-515">Organizer:[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)|[Organizer](/javascript/api/outlook/office.organizer)</span><span class="sxs-lookup"><span data-stu-id="22e1f-515">organizer :[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)|[Organizer](/javascript/api/outlook/office.organizer)</span></span>

<span data-ttu-id="22e1f-516">Ruft die e-Mail-Adresse des Dialogfelds Organisieren für eine angegebene Besprechung ab.</span><span class="sxs-lookup"><span data-stu-id="22e1f-516">Gets the email address of the organizer for a specified meeting.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="22e1f-517">Lesemodus</span><span class="sxs-lookup"><span data-stu-id="22e1f-517">Read mode</span></span>

<span data-ttu-id="22e1f-518">Die `organizer` -Eigenschaft gibt ein [EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails) -Objekt, die den Besprechungsorganisator darstellt.</span><span class="sxs-lookup"><span data-stu-id="22e1f-518">The `organizer` property returns an [EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails) object that represents the meeting organizer.</span></span>

##### <a name="compose-mode"></a><span data-ttu-id="22e1f-519">Verfassenmodus</span><span class="sxs-lookup"><span data-stu-id="22e1f-519">Compose mode</span></span>

<span data-ttu-id="22e1f-520">Die `organizer` -Eigenschaft gibt ein [Organisator](/javascript/api/outlook/office.organizer) -Objekt, das eine Methode zum Abrufen des Werts der Organisator bereitstellt.</span><span class="sxs-lookup"><span data-stu-id="22e1f-520">The `organizer` property returns an [Organizer](/javascript/api/outlook/office.organizer) object that provides a method to get the organizer value.</span></span>

##### <a name="type"></a><span data-ttu-id="22e1f-521">Typ:</span><span class="sxs-lookup"><span data-stu-id="22e1f-521">Type:</span></span>

*   <span data-ttu-id="22e1f-522">[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails) | [Organizer](/javascript/api/outlook/office.organizer)</span><span class="sxs-lookup"><span data-stu-id="22e1f-522">[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails) | [Organizer](/javascript/api/outlook/office.organizer)</span></span>

##### <a name="requirements"></a><span data-ttu-id="22e1f-523">Anforderungen</span><span class="sxs-lookup"><span data-stu-id="22e1f-523">Requirements</span></span>

|<span data-ttu-id="22e1f-524">Anforderung</span><span class="sxs-lookup"><span data-stu-id="22e1f-524">Requirement</span></span>|||
|---|---|---|
|[<span data-ttu-id="22e1f-525">Mindestversion des Postfachanforderungssatzes</span><span class="sxs-lookup"><span data-stu-id="22e1f-525">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="22e1f-526">1.0</span><span class="sxs-lookup"><span data-stu-id="22e1f-526">1.0</span></span>|<span data-ttu-id="22e1f-527">Vorschau</span><span class="sxs-lookup"><span data-stu-id="22e1f-527">Preview</span></span>|
|[<span data-ttu-id="22e1f-528">Mindestberechtigungsstufe</span><span class="sxs-lookup"><span data-stu-id="22e1f-528">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="22e1f-529">ReadItem</span><span class="sxs-lookup"><span data-stu-id="22e1f-529">ReadItem</span></span>|<span data-ttu-id="22e1f-530">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="22e1f-530">ReadWriteItem</span></span>|
|[<span data-ttu-id="22e1f-531">Zutreffender Outlook-Modus</span><span class="sxs-lookup"><span data-stu-id="22e1f-531">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="22e1f-532">Lesen</span><span class="sxs-lookup"><span data-stu-id="22e1f-532">Read</span></span>|<span data-ttu-id="22e1f-533">Verfassen</span><span class="sxs-lookup"><span data-stu-id="22e1f-533">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="22e1f-534">Beispiel</span><span class="sxs-lookup"><span data-stu-id="22e1f-534">Example</span></span>

```
var organizerName = Office.context.mailbox.item.organizer.displayName;
var organizerAddress = Office.context.mailbox.item.organizer.emailAddress;
```

#### <a name="nullable-recurrence-recurrencejavascriptapioutlookofficerecurrence"></a><span data-ttu-id="22e1f-535">(Nullwerte zulassen) Serie:[Serie](/javascript/api/outlook/office.recurrence)</span><span class="sxs-lookup"><span data-stu-id="22e1f-535">(nullable) recurrence :[Recurrence](/javascript/api/outlook/office.recurrence)</span></span>

<span data-ttu-id="22e1f-536">Ruft ab oder legt das Serienmuster eines Termins.</span><span class="sxs-lookup"><span data-stu-id="22e1f-536">Gets or sets the recurrence pattern of an appointment.</span></span> <span data-ttu-id="22e1f-537">Ruft das Serienmuster eine Besprechungsanfrage an.</span><span class="sxs-lookup"><span data-stu-id="22e1f-537">Gets the recurrence pattern of a meeting request.</span></span> <span data-ttu-id="22e1f-538">Lesen Sie und verfassen Sie Modi für Terminelemente.</span><span class="sxs-lookup"><span data-stu-id="22e1f-538">Read and compose modes for appointment items.</span></span> <span data-ttu-id="22e1f-539">Zur Erfüllung der Anforderung Elemente im Lesemodus.</span><span class="sxs-lookup"><span data-stu-id="22e1f-539">Read mode for meeting request items.</span></span>

<span data-ttu-id="22e1f-540">Die `recurrence` -Eigenschaft gibt ein Objekt [Serie](/javascript/api/outlook/office.recurrence) für Termin- oder Besprechungsanfragen zurück, wenn ein Element eine Datenreihe oder eine Instanz einer Serie ist.</span><span class="sxs-lookup"><span data-stu-id="22e1f-540">The `recurrence` property returns a [recurrence](/javascript/api/outlook/office.recurrence) object for recurring appointments or meetings requests if an item is a series or an instance in a series.</span></span> <span data-ttu-id="22e1f-541">`null`wird für einzelne Termine und Besprechungsanfragen von einzelnen Terminen zurückgegeben.</span><span class="sxs-lookup"><span data-stu-id="22e1f-541">`null` is returned for single appointments and meeting requests of single appointments.</span></span> <span data-ttu-id="22e1f-542">`undefined`wird für Nachrichten zurückgegeben, die nicht Besprechungsanfragen sind.</span><span class="sxs-lookup"><span data-stu-id="22e1f-542">`undefined` is returned for messages that are not meeting requests.</span></span>

> <span data-ttu-id="22e1f-543">Hinweis: Besprechungsanfragen verfügen über eine `itemClass` Wert der IPM. Schedule.Meeting.Request.</span><span class="sxs-lookup"><span data-stu-id="22e1f-543">Note: Meeting requests have an `itemClass` value of IPM.Schedule.Meeting.Request.</span></span>

> <span data-ttu-id="22e1f-544">Hinweis: Wenn das Objekt Serie ist `null`, darauf hin, dass das Objekt einen einzelnen Termin oder eine Besprechungsanfrage an einen einzelnen Termin und nicht um einen Teil einer Reihe ist.</span><span class="sxs-lookup"><span data-stu-id="22e1f-544">Note: If the recurrence object is `null`, this indicates that the object is a single appointment or a meeting request of a single appointment and NOT a part of a series.</span></span>

##### <a name="type"></a><span data-ttu-id="22e1f-545">Typ:</span><span class="sxs-lookup"><span data-stu-id="22e1f-545">Type:</span></span>

* [<span data-ttu-id="22e1f-546">Serie</span><span class="sxs-lookup"><span data-stu-id="22e1f-546">Recurrence</span></span>](/javascript/api/outlook/office.recurrence)

|<span data-ttu-id="22e1f-547">Anforderung</span><span class="sxs-lookup"><span data-stu-id="22e1f-547">Requirement</span></span>|<span data-ttu-id="22e1f-548">Wert</span><span class="sxs-lookup"><span data-stu-id="22e1f-548">Value</span></span>|
|---|---|
|[<span data-ttu-id="22e1f-549">Mindestversion des Postfachanforderungssatzes</span><span class="sxs-lookup"><span data-stu-id="22e1f-549">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="22e1f-550">Vorschau</span><span class="sxs-lookup"><span data-stu-id="22e1f-550">Preview</span></span>|
|[<span data-ttu-id="22e1f-551">Mindestberechtigungsstufe</span><span class="sxs-lookup"><span data-stu-id="22e1f-551">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="22e1f-552">ReadItem</span><span class="sxs-lookup"><span data-stu-id="22e1f-552">ReadItem</span></span>|
|[<span data-ttu-id="22e1f-553">Zutreffender Outlook-Modus</span><span class="sxs-lookup"><span data-stu-id="22e1f-553">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="22e1f-554">Verfassen oder Lesen</span><span class="sxs-lookup"><span data-stu-id="22e1f-554">Compose or read</span></span>|

####  <a name="requiredattendees-arrayemailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsrecipientsjavascriptapioutlookofficerecipients"></a><span data-ttu-id="22e1f-555">requiredAttendees :Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="22e1f-555">requiredAttendees :Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook/office.recipients)</span></span>

<span data-ttu-id="22e1f-556">Ermöglicht den Zugriff auf die erforderlichen Teilnehmer eines Ereignisses.</span><span class="sxs-lookup"><span data-stu-id="22e1f-556">Provides access to the required attendees of an event.</span></span> <span data-ttu-id="22e1f-557">Der Typ des Objekts und die Zugriffsebene, hängt von den Modus des aktuellen Elements ab.</span><span class="sxs-lookup"><span data-stu-id="22e1f-557">The type of object and level of access depends on the mode of the current item.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="22e1f-558">Lesemodus</span><span class="sxs-lookup"><span data-stu-id="22e1f-558">Read mode</span></span>

<span data-ttu-id="22e1f-559">Die `requiredAttendees`-Eigenschaft gibt ein Array mit einem `EmailAddressDetails`-Objekt für jeden erforderlichen Teilnehmer an der Besprechung zurück.</span><span class="sxs-lookup"><span data-stu-id="22e1f-559">The `requiredAttendees` property returns an array that contains an `EmailAddressDetails` object for each required attendee to the meeting.</span></span>

##### <a name="compose-mode"></a><span data-ttu-id="22e1f-560">Verfassenmodus</span><span class="sxs-lookup"><span data-stu-id="22e1f-560">Compose mode</span></span>

<span data-ttu-id="22e1f-561">Die `requiredAttendees` -Eigenschaft gibt eine `Recipients` -Objekt, das Methoden zum Abrufen oder Aktualisieren der erforderlichen Teilnehmer für eine Besprechung bereitstellt.</span><span class="sxs-lookup"><span data-stu-id="22e1f-561">The `requiredAttendees` property returns a `Recipients` object that provides methods to get or update the required attendees for a meeting.</span></span>

##### <a name="type"></a><span data-ttu-id="22e1f-562">Typ:</span><span class="sxs-lookup"><span data-stu-id="22e1f-562">Type:</span></span>

*   <span data-ttu-id="22e1f-563">Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="22e1f-563">Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook/office.recipients)</span></span>

##### <a name="requirements"></a><span data-ttu-id="22e1f-564">Anforderungen</span><span class="sxs-lookup"><span data-stu-id="22e1f-564">Requirements</span></span>

|<span data-ttu-id="22e1f-565">Anforderung</span><span class="sxs-lookup"><span data-stu-id="22e1f-565">Requirement</span></span>|<span data-ttu-id="22e1f-566">Wert</span><span class="sxs-lookup"><span data-stu-id="22e1f-566">Value</span></span>|
|---|---|
|[<span data-ttu-id="22e1f-567">Mindestversion des Postfachanforderungssatzes</span><span class="sxs-lookup"><span data-stu-id="22e1f-567">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="22e1f-568">1.0</span><span class="sxs-lookup"><span data-stu-id="22e1f-568">1.0</span></span>|
|[<span data-ttu-id="22e1f-569">Mindestberechtigungsstufe</span><span class="sxs-lookup"><span data-stu-id="22e1f-569">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="22e1f-570">ReadItem</span><span class="sxs-lookup"><span data-stu-id="22e1f-570">ReadItem</span></span>|
|[<span data-ttu-id="22e1f-571">Zutreffender Outlook-Modus</span><span class="sxs-lookup"><span data-stu-id="22e1f-571">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="22e1f-572">Verfassen oder Lesen</span><span class="sxs-lookup"><span data-stu-id="22e1f-572">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="22e1f-573">Beispiel</span><span class="sxs-lookup"><span data-stu-id="22e1f-573">Example</span></span>

```
Office.context.mailbox.item.requiredAttendees.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.requiredAttendees.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.requiredAttendees.getAsync(callback);

function callback(asyncResult) {
  var arrayOfRequiredAttendeesRecipients = asyncResult.value;
}
```

#### <a name="sender-emailaddressdetailsjavascriptapioutlookofficeemailaddressdetails"></a><span data-ttu-id="22e1f-574">sender :[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)</span><span class="sxs-lookup"><span data-stu-id="22e1f-574">sender :[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)</span></span>

<span data-ttu-id="22e1f-p126">Ruft die E-Mail-Adresse des Absenders einer E-Mail-Nachricht ab. Nur Lesemodus.</span><span class="sxs-lookup"><span data-stu-id="22e1f-p126">Gets the email address of the sender of an email message. Read mode only.</span></span>

<span data-ttu-id="22e1f-p127">Die [`from`](#from-emailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsfromjavascriptapioutlookofficefrom)- und `sender`-Eigenschaften stellen dieselbe Person dar, außer die Nachricht wurde von einem Delegaten gesendet. In diesem Fall stellt die `from`-Eigenschaft die Stellvertretung dar, und die sender-Eigenschaft stellt den Delegaten dar.</span><span class="sxs-lookup"><span data-stu-id="22e1f-p127">The [`from`](#from-emailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsfromjavascriptapioutlookofficefrom) and `sender` properties represent the same person unless the message is sent by a delegate. In that case, the `from` property represents the delegator, and the sender property represents the delegate.</span></span>

> [!NOTE]
> <span data-ttu-id="22e1f-579">Die `recipientType` -Eigenschaft des der `EmailAddressDetails` -Objekts der `sender` -Eigenschaft ist `undefined`.</span><span class="sxs-lookup"><span data-stu-id="22e1f-579">The `recipientType` property of the `EmailAddressDetails` object in the `sender` property is `undefined`.</span></span>

##### <a name="type"></a><span data-ttu-id="22e1f-580">Typ:</span><span class="sxs-lookup"><span data-stu-id="22e1f-580">Type:</span></span>

*   [<span data-ttu-id="22e1f-581">EmailAddressDetails</span><span class="sxs-lookup"><span data-stu-id="22e1f-581">EmailAddressDetails</span></span>](/javascript/api/outlook/office.emailaddressdetails)

##### <a name="requirements"></a><span data-ttu-id="22e1f-582">Anforderungen</span><span class="sxs-lookup"><span data-stu-id="22e1f-582">Requirements</span></span>

|<span data-ttu-id="22e1f-583">Anforderung</span><span class="sxs-lookup"><span data-stu-id="22e1f-583">Requirement</span></span>|<span data-ttu-id="22e1f-584">Wert</span><span class="sxs-lookup"><span data-stu-id="22e1f-584">Value</span></span>|
|---|---|
|[<span data-ttu-id="22e1f-585">Mindestversion des Postfachanforderungssatzes</span><span class="sxs-lookup"><span data-stu-id="22e1f-585">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="22e1f-586">1.0</span><span class="sxs-lookup"><span data-stu-id="22e1f-586">1.0</span></span>|
|[<span data-ttu-id="22e1f-587">Mindestberechtigungsstufe</span><span class="sxs-lookup"><span data-stu-id="22e1f-587">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="22e1f-588">ReadItem</span><span class="sxs-lookup"><span data-stu-id="22e1f-588">ReadItem</span></span>|
|[<span data-ttu-id="22e1f-589">Zutreffender Outlook-Modus</span><span class="sxs-lookup"><span data-stu-id="22e1f-589">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="22e1f-590">Lesen</span><span class="sxs-lookup"><span data-stu-id="22e1f-590">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="22e1f-591">Beispiel</span><span class="sxs-lookup"><span data-stu-id="22e1f-591">Example</span></span>

```
var senderName = Office.context.mailbox.item.sender.displayName;
var senderAddress = Office.context.mailbox.item.sender.emailAddress;
```

#### <a name="nullable-seriesid-string"></a><span data-ttu-id="22e1f-592">(Nullwerte zulassen) SeriesId: Zeichenfolge</span><span class="sxs-lookup"><span data-stu-id="22e1f-592">(nullable) seriesId :String</span></span>

<span data-ttu-id="22e1f-593">Ruft die Id der Datenreihe, zu der eine Instanz gehört.</span><span class="sxs-lookup"><span data-stu-id="22e1f-593">Gets the id of the series that an instance belongs to.</span></span>

<span data-ttu-id="22e1f-594">In OWA und Outlook die `seriesId` gibt die Exchange-Webdienste (EWS)-ID des übergeordneten (Reihe)-Elements, das dieses Element gehört.</span><span class="sxs-lookup"><span data-stu-id="22e1f-594">In OWA and Outlook, the `seriesId` returns the Exchange Web Services (EWS) ID of the parent (series) item that this item belongs to.</span></span> <span data-ttu-id="22e1f-595">Jedoch in iOS und Android die `seriesId` gibt die REST-ID des übergeordneten Elements zurück.</span><span class="sxs-lookup"><span data-stu-id="22e1f-595">However, in iOS and Android, the `seriesId` returns the REST ID of the parent item.</span></span>

> [!NOTE]
> <span data-ttu-id="22e1f-596">Der Bezeichner, der von der `seriesId`-Eigenschaft zurückgegeben wird, ist der gleiche Bezeichner wie der Exchange-Webdienste-Elementbezeichner.</span><span class="sxs-lookup"><span data-stu-id="22e1f-596">The identifier returned by the `seriesId` property is the same as the Exchange Web Services item identifier.</span></span> <span data-ttu-id="22e1f-597">Die `seriesId` -Eigenschaft ist nicht identisch mit der Outlook-IDs von der Outlook-REST-API verwendet.</span><span class="sxs-lookup"><span data-stu-id="22e1f-597">The `seriesId` property is not identical to the Outlook IDs used by the Outlook REST API.</span></span> <span data-ttu-id="22e1f-598">REST-API-Aufrufe dieser Wert verwenden, sollten sie konvertiert werden, bevor [Office.context.mailbox.convertToRestId](office.context.mailbox.md#converttorestiditemid-restversion--string)verwenden.</span><span class="sxs-lookup"><span data-stu-id="22e1f-598">Before making REST API calls using this value, it should be converted using [Office.context.mailbox.convertToRestId](office.context.mailbox.md#converttorestiditemid-restversion--string).</span></span> <span data-ttu-id="22e1f-599">Weitere Informationen finden Sie unter [Verwenden der Outlook-REST-APIs aus einem Outlook-add-in](https://docs.microsoft.com/outlook/add-ins/use-rest-api).</span><span class="sxs-lookup"><span data-stu-id="22e1f-599">For more details, see [Use the Outlook REST APIs from an Outlook add-in](https://docs.microsoft.com/outlook/add-ins/use-rest-api).</span></span>

<span data-ttu-id="22e1f-600">Die `seriesId` -Eigenschaft gibt `null` für Elemente, deren nicht das übergeordnete Elemente wie Termine Serienelementen einzelne, oder Besprechungsanfragen und gibt `undefined` für alle anderen Elemente, die nicht Anforderungen erfüllt sind.</span><span class="sxs-lookup"><span data-stu-id="22e1f-600">The `seriesId` property returns `null` for items that do not have parent items such as single appointments, series items, or meeting requests and returns `undefined` for any other items that are not meeting requests.</span></span>

##### <a name="type"></a><span data-ttu-id="22e1f-601">Typ:</span><span class="sxs-lookup"><span data-stu-id="22e1f-601">Type:</span></span>

* <span data-ttu-id="22e1f-602">String</span><span class="sxs-lookup"><span data-stu-id="22e1f-602">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="22e1f-603">Anforderungen</span><span class="sxs-lookup"><span data-stu-id="22e1f-603">Requirements</span></span>

|<span data-ttu-id="22e1f-604">Anforderung</span><span class="sxs-lookup"><span data-stu-id="22e1f-604">Requirement</span></span>|<span data-ttu-id="22e1f-605">Wert</span><span class="sxs-lookup"><span data-stu-id="22e1f-605">Value</span></span>|
|---|---|
|[<span data-ttu-id="22e1f-606">Mindestversion des Postfachanforderungssatzes</span><span class="sxs-lookup"><span data-stu-id="22e1f-606">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="22e1f-607">Vorschau</span><span class="sxs-lookup"><span data-stu-id="22e1f-607">Preview</span></span>|
|[<span data-ttu-id="22e1f-608">Mindestberechtigungsstufe</span><span class="sxs-lookup"><span data-stu-id="22e1f-608">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="22e1f-609">ReadItem</span><span class="sxs-lookup"><span data-stu-id="22e1f-609">ReadItem</span></span>|
|[<span data-ttu-id="22e1f-610">Zutreffender Outlook-Modus</span><span class="sxs-lookup"><span data-stu-id="22e1f-610">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="22e1f-611">Verfassen oder Lesen</span><span class="sxs-lookup"><span data-stu-id="22e1f-611">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="22e1f-612">Beispiel</span><span class="sxs-lookup"><span data-stu-id="22e1f-612">Example</span></span>

```
var seriesId = Office.context.mailbox.item.seriesId; 
var isSeries = (seriesId == null);
```

####  <a name="start-datetimejavascriptapioutlookofficetime"></a><span data-ttu-id="22e1f-613">start :Date|[Time](/javascript/api/outlook/office.time)</span><span class="sxs-lookup"><span data-stu-id="22e1f-613">start :Date|[Time](/javascript/api/outlook/office.time)</span></span>

<span data-ttu-id="22e1f-614">Ruft Datum und Zeit für den Beginn des Termins ab oder legt Datum und Uhrzeit fest.</span><span class="sxs-lookup"><span data-stu-id="22e1f-614">Gets or sets the date and time that the appointment is to begin.</span></span>

<span data-ttu-id="22e1f-p130">Die `start`-Eigenschaft wird als Datums- und Uhrzeitwert in UTC ausgedrückt. Sie können die [`convertToLocalClientTime`](office.context.mailbox.md#converttolocalclienttimetimevalue--localclienttimejavascriptapioutlookofficelocalclienttime)-Methode verwenden, um den Wert in den lokalen Uhrzeit- und Datumswert des Clients umzuwandeln.</span><span class="sxs-lookup"><span data-stu-id="22e1f-p130">The `start` property is expressed as a Coordinated Universal Time (UTC) date and time value. You can use the [`convertToLocalClientTime`](office.context.mailbox.md#converttolocalclienttimetimevalue--localclienttimejavascriptapioutlookofficelocalclienttime) method to convert the value to the client’s local date and time.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="22e1f-617">Lesemodus</span><span class="sxs-lookup"><span data-stu-id="22e1f-617">Read mode</span></span>

<span data-ttu-id="22e1f-618">Die `start`-Eigenschaft gibt ein `Date`-Objekt zurück.</span><span class="sxs-lookup"><span data-stu-id="22e1f-618">The `start` property returns a `Date` object.</span></span>

##### <a name="compose-mode"></a><span data-ttu-id="22e1f-619">Verfassenmodus</span><span class="sxs-lookup"><span data-stu-id="22e1f-619">Compose mode</span></span>

<span data-ttu-id="22e1f-620">Die `start`-Eigenschaft gibt ein `Time`-Objekt zurück.</span><span class="sxs-lookup"><span data-stu-id="22e1f-620">The `start` property returns a `Time` object.</span></span>

<span data-ttu-id="22e1f-621">Wenn Sie die [`Time.setAsync`](/javascript/api/outlook/office.time#setasync-datetime--options--callback-)-Methode verwenden, um die Startzeit im Verfassenmodus festzulegen, sollten Sie die [`convertToUtcClientTime`](office.context.mailbox.md#converttoutcclienttimeinput--date)-Methode verwenden, um die Ortszeit auf dem Client für den Server in UTC umzuwandeln.</span><span class="sxs-lookup"><span data-stu-id="22e1f-621">When you use the [`Time.setAsync`](/javascript/api/outlook/office.time#setasync-datetime--options--callback-) method to set the start time, you should use the [`convertToUtcClientTime`](office.context.mailbox.md#converttoutcclienttimeinput--date) method to convert the local time on the client to UTC for the server.</span></span>

##### <a name="type"></a><span data-ttu-id="22e1f-622">Typ:</span><span class="sxs-lookup"><span data-stu-id="22e1f-622">Type:</span></span>

*   <span data-ttu-id="22e1f-623">Date | [Time](/javascript/api/outlook/office.time)</span><span class="sxs-lookup"><span data-stu-id="22e1f-623">Date | [Time](/javascript/api/outlook/office.time)</span></span>

##### <a name="requirements"></a><span data-ttu-id="22e1f-624">Anforderungen</span><span class="sxs-lookup"><span data-stu-id="22e1f-624">Requirements</span></span>

|<span data-ttu-id="22e1f-625">Anforderung</span><span class="sxs-lookup"><span data-stu-id="22e1f-625">Requirement</span></span>|<span data-ttu-id="22e1f-626">Wert</span><span class="sxs-lookup"><span data-stu-id="22e1f-626">Value</span></span>|
|---|---|
|[<span data-ttu-id="22e1f-627">Mindestversion des Postfachanforderungssatzes</span><span class="sxs-lookup"><span data-stu-id="22e1f-627">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="22e1f-628">1.0</span><span class="sxs-lookup"><span data-stu-id="22e1f-628">1.0</span></span>|
|[<span data-ttu-id="22e1f-629">Mindestberechtigungsstufe</span><span class="sxs-lookup"><span data-stu-id="22e1f-629">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="22e1f-630">ReadItem</span><span class="sxs-lookup"><span data-stu-id="22e1f-630">ReadItem</span></span>|
|[<span data-ttu-id="22e1f-631">Zutreffender Outlook-Modus</span><span class="sxs-lookup"><span data-stu-id="22e1f-631">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="22e1f-632">Verfassen oder Lesen</span><span class="sxs-lookup"><span data-stu-id="22e1f-632">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="22e1f-633">Beispiel</span><span class="sxs-lookup"><span data-stu-id="22e1f-633">Example</span></span>

<span data-ttu-id="22e1f-634">Im folgenden Beispiel wird die Startzeit eines Termins im Verfassenmodus mithilfe der [`setAsync`](/javascript/api/outlook/office.time#setasync-datetime--options--callback-)-Methode des `Time`-Objekts festgelegt.</span><span class="sxs-lookup"><span data-stu-id="22e1f-634">The following example sets the start time of an appointment in compose mode by using the [`setAsync`](/javascript/api/outlook/office.time#setasync-datetime--options--callback-) method of the `Time` object.</span></span>

```
var startTime = new Date("3/14/2015");
var options = {
  // Pass information that can be used
  // in the callback
     asyncContext: {verb:"Set"}
}
Office.context.mailbox.item.start.setAsync(startTime, options, function(result) {
  if (result.error) {
    console.debug(result.error);
  } else {
    // Access the asyncContext that was passed to the setAsync function
    console.debug("Start Time " + result.asyncContext.verb);
  }
});
```

####  <a name="subject-stringsubjectjavascriptapioutlookofficesubject"></a><span data-ttu-id="22e1f-635">subject :String|[Subject](/javascript/api/outlook/office.subject)</span><span class="sxs-lookup"><span data-stu-id="22e1f-635">subject :String|[Subject](/javascript/api/outlook/office.subject)</span></span>

<span data-ttu-id="22e1f-636">Ruft die Beschreibung ab, die im Betrefffeld eines Elements angezeigt wird, oder legt sie fest.</span><span class="sxs-lookup"><span data-stu-id="22e1f-636">Gets or sets the description that appears in the subject field of an item.</span></span>

<span data-ttu-id="22e1f-637">Die `subject`-Eigenschaft ruft den gesamten Betreff des Elements ab oder legt ihn fest – so, wie er vom E-Mail-Server gesendet wird.</span><span class="sxs-lookup"><span data-stu-id="22e1f-637">The `subject` property gets or sets the entire subject of the item, as sent by the email server.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="22e1f-638">Lesemodus</span><span class="sxs-lookup"><span data-stu-id="22e1f-638">Read mode</span></span>

<span data-ttu-id="22e1f-p131">Die `subject`-Eigenschaft gibt eine Zeichenfolge zurück. Verwenden Sie die [`normalizedSubject`](#normalizedsubject-string)-Eigenschaft, um den Betreff ohne führende Präfixe wie `RE:` und `FW:` abzurufen.</span><span class="sxs-lookup"><span data-stu-id="22e1f-p131">The `subject` property returns a string. Use the [`normalizedSubject`](#normalizedsubject-string) property to get the subject minus any leading prefixes such as `RE:` and `FW:`.</span></span>

```
var subject = Office.context.mailbox.item.subject;
```

##### <a name="compose-mode"></a><span data-ttu-id="22e1f-641">Verfassenmodus</span><span class="sxs-lookup"><span data-stu-id="22e1f-641">Compose mode</span></span>

<span data-ttu-id="22e1f-642">Die `subject`-Eigenschaft gibt ein `Subject`-Objekt zurück, das Methoden zum Abrufen und Festlegen des Betreffs bereitstellt.</span><span class="sxs-lookup"><span data-stu-id="22e1f-642">The `subject` property returns a `Subject` object that provides methods to get and set the subject.</span></span>

```
Office.context.mailbox.item.subject.getAsync(callback);

function callback(asyncResult) {
  var subject = asyncResult.value;
}
```

##### <a name="type"></a><span data-ttu-id="22e1f-643">Typ:</span><span class="sxs-lookup"><span data-stu-id="22e1f-643">Type:</span></span>

*   <span data-ttu-id="22e1f-644">String | [Subject](/javascript/api/outlook/office.subject)</span><span class="sxs-lookup"><span data-stu-id="22e1f-644">String | [Subject](/javascript/api/outlook/office.subject)</span></span>

##### <a name="requirements"></a><span data-ttu-id="22e1f-645">Anforderungen</span><span class="sxs-lookup"><span data-stu-id="22e1f-645">Requirements</span></span>

|<span data-ttu-id="22e1f-646">Anforderung</span><span class="sxs-lookup"><span data-stu-id="22e1f-646">Requirement</span></span>|<span data-ttu-id="22e1f-647">Wert</span><span class="sxs-lookup"><span data-stu-id="22e1f-647">Value</span></span>|
|---|---|
|[<span data-ttu-id="22e1f-648">Mindestversion des Postfachanforderungssatzes</span><span class="sxs-lookup"><span data-stu-id="22e1f-648">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="22e1f-649">1.0</span><span class="sxs-lookup"><span data-stu-id="22e1f-649">1.0</span></span>|
|[<span data-ttu-id="22e1f-650">Mindestberechtigungsstufe</span><span class="sxs-lookup"><span data-stu-id="22e1f-650">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="22e1f-651">ReadItem</span><span class="sxs-lookup"><span data-stu-id="22e1f-651">ReadItem</span></span>|
|[<span data-ttu-id="22e1f-652">Zutreffender Outlook-Modus</span><span class="sxs-lookup"><span data-stu-id="22e1f-652">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="22e1f-653">Verfassen oder Lesen</span><span class="sxs-lookup"><span data-stu-id="22e1f-653">Compose or read</span></span>|

####  <a name="to-arrayemailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsrecipientsjavascriptapioutlookofficerecipients"></a><span data-ttu-id="22e1f-654">an: Array. <[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)>|[Empfänger](/javascript/api/outlook/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="22e1f-654">to :Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook/office.recipients)</span></span>

<span data-ttu-id="22e1f-655">Ermöglicht den Zugriff auf die Empfänger in der Zeile **an** einer Nachricht.</span><span class="sxs-lookup"><span data-stu-id="22e1f-655">Provides access to the recipients on the **To** line of a message.</span></span> <span data-ttu-id="22e1f-656">Der Typ des Objekts und die Zugriffsebene, hängt von den Modus des aktuellen Elements ab.</span><span class="sxs-lookup"><span data-stu-id="22e1f-656">The type of object and level of access depends on the mode of the current item.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="22e1f-657">Lesemodus</span><span class="sxs-lookup"><span data-stu-id="22e1f-657">Read mode</span></span>

<span data-ttu-id="22e1f-p133">Die `to`-Eigenschaft gibt ein Array mit einem `EmailAddressDetails`-Objekt für jeden Empfänger in der**An**-Zeile der Nachricht zurück. Die Auflistung ist auf höchstens 100 Elemente beschränkt.</span><span class="sxs-lookup"><span data-stu-id="22e1f-p133">The `to` property returns an array that contains an `EmailAddressDetails` object for each recipient listed on the **To** line of the message. The collection is limited to a maximum of 100 members.</span></span>

##### <a name="compose-mode"></a><span data-ttu-id="22e1f-660">Verfassenmodus</span><span class="sxs-lookup"><span data-stu-id="22e1f-660">Compose mode</span></span>

<span data-ttu-id="22e1f-661">Die `to` -Eigenschaft gibt eine `Recipients` -Objekt, das Methoden zum Abrufen oder aktualisieren die Empfänger in der Zeile **an** der Nachricht bereitstellt.</span><span class="sxs-lookup"><span data-stu-id="22e1f-661">The `to` property returns a `Recipients` object that provides methods to get or update the recipients on the **To** line of the message.</span></span>

##### <a name="type"></a><span data-ttu-id="22e1f-662">Typ:</span><span class="sxs-lookup"><span data-stu-id="22e1f-662">Type:</span></span>

*   <span data-ttu-id="22e1f-663">Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="22e1f-663">Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook/office.recipients)</span></span>

##### <a name="requirements"></a><span data-ttu-id="22e1f-664">Anforderungen</span><span class="sxs-lookup"><span data-stu-id="22e1f-664">Requirements</span></span>

|<span data-ttu-id="22e1f-665">Anforderung</span><span class="sxs-lookup"><span data-stu-id="22e1f-665">Requirement</span></span>|<span data-ttu-id="22e1f-666">Wert</span><span class="sxs-lookup"><span data-stu-id="22e1f-666">Value</span></span>|
|---|---|
|[<span data-ttu-id="22e1f-667">Mindestversion des Postfachanforderungssatzes</span><span class="sxs-lookup"><span data-stu-id="22e1f-667">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="22e1f-668">1.0</span><span class="sxs-lookup"><span data-stu-id="22e1f-668">1.0</span></span>|
|[<span data-ttu-id="22e1f-669">Mindestberechtigungsstufe</span><span class="sxs-lookup"><span data-stu-id="22e1f-669">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="22e1f-670">ReadItem</span><span class="sxs-lookup"><span data-stu-id="22e1f-670">ReadItem</span></span>|
|[<span data-ttu-id="22e1f-671">Zutreffender Outlook-Modus</span><span class="sxs-lookup"><span data-stu-id="22e1f-671">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="22e1f-672">Verfassen oder Lesen</span><span class="sxs-lookup"><span data-stu-id="22e1f-672">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="22e1f-673">Beispiel</span><span class="sxs-lookup"><span data-stu-id="22e1f-673">Example</span></span>

```
Office.context.mailbox.item.to.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.to.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.to.getAsync(callback);

function callback(asyncResult) {
  var arrayOfToRecipients = asyncResult.value;
}
```

### <a name="methods"></a><span data-ttu-id="22e1f-674">Methoden</span><span class="sxs-lookup"><span data-stu-id="22e1f-674">Methods</span></span>

####  <a name="addfileattachmentasyncuri-attachmentname-options-callback"></a><span data-ttu-id="22e1f-675">addFileAttachmentAsync(uri, attachmentName, [options], [callback])</span><span class="sxs-lookup"><span data-stu-id="22e1f-675">addFileAttachmentAsync(uri, attachmentName, [options], [callback])</span></span>

<span data-ttu-id="22e1f-676">Fügt eine Datei zu einer Nachricht oder einem Termin als Anlage hinzu.</span><span class="sxs-lookup"><span data-stu-id="22e1f-676">Adds a file to a message or appointment as an attachment.</span></span>

<span data-ttu-id="22e1f-677">Die `addFileAttachmentAsync`-Methode lädt die Datei am angegebenen URI hoch und fügt sie an das Element im Verfassenformular an.</span><span class="sxs-lookup"><span data-stu-id="22e1f-677">The `addFileAttachmentAsync` method uploads the file at the specified URI and attaches it to the item in the compose form.</span></span>

<span data-ttu-id="22e1f-678">Anschließend können Sie den Bezeichner mit der [`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback)-Methode in der gleichen Sitzung zum Entfernen der Anlage verwenden.</span><span class="sxs-lookup"><span data-stu-id="22e1f-678">You can subsequently use the identifier with the [`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback) method to remove the attachment in the same session.</span></span>

##### <a name="parameters"></a><span data-ttu-id="22e1f-679">Parameter:</span><span class="sxs-lookup"><span data-stu-id="22e1f-679">Parameters:</span></span>
|<span data-ttu-id="22e1f-680">Name</span><span class="sxs-lookup"><span data-stu-id="22e1f-680">Name</span></span>|<span data-ttu-id="22e1f-681">Typ</span><span class="sxs-lookup"><span data-stu-id="22e1f-681">Type</span></span>|<span data-ttu-id="22e1f-682">Attribute</span><span class="sxs-lookup"><span data-stu-id="22e1f-682">Attributes</span></span>|<span data-ttu-id="22e1f-683">Beschreibung</span><span class="sxs-lookup"><span data-stu-id="22e1f-683">Description</span></span>|
|---|---|---|---|
|`uri`|<span data-ttu-id="22e1f-684">Zeichenfolge</span><span class="sxs-lookup"><span data-stu-id="22e1f-684">String</span></span>||<span data-ttu-id="22e1f-p134">Der URI, der den Speicherort der an die Nachricht oder den Termin anzuhängenden Datei angibt. Die maximale Länge ist 2048 Zeichen.</span><span class="sxs-lookup"><span data-stu-id="22e1f-p134">The URI that provides the location of the file to attach to the message or appointment. The maximum length is 2048 characters.</span></span>|
|`attachmentName`|<span data-ttu-id="22e1f-687">Zeichenfolge</span><span class="sxs-lookup"><span data-stu-id="22e1f-687">String</span></span>||<span data-ttu-id="22e1f-p135">Der Name der Anlage, der beim Hochladen der Anlage angezeigt wird. Die maximale Länge ist 255 Zeichen.</span><span class="sxs-lookup"><span data-stu-id="22e1f-p135">The name of the attachment that is shown while the attachment is uploading. The maximum length is 255 characters.</span></span>|
|`options`|<span data-ttu-id="22e1f-690">Object</span><span class="sxs-lookup"><span data-stu-id="22e1f-690">Object</span></span>|<span data-ttu-id="22e1f-691">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="22e1f-691">&lt;optional&gt;</span></span>|<span data-ttu-id="22e1f-692">Ein Objektliteral, das mindestens eine der folgenden Eigenschaften enthält.</span><span class="sxs-lookup"><span data-stu-id="22e1f-692">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`|<span data-ttu-id="22e1f-693">Object</span><span class="sxs-lookup"><span data-stu-id="22e1f-693">Object</span></span>|<span data-ttu-id="22e1f-694">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="22e1f-694">&lt;optional&gt;</span></span>|<span data-ttu-id="22e1f-695">Entwickler können jedes Objekt angeben, auf das sie in der Rückrufmethode zugreifen möchten.</span><span class="sxs-lookup"><span data-stu-id="22e1f-695">Developers can provide any object they wish to access in the callback method.</span></span>|
|`options.isInline`|<span data-ttu-id="22e1f-696">Boolean</span><span class="sxs-lookup"><span data-stu-id="22e1f-696">Boolean</span></span>|<span data-ttu-id="22e1f-697">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="22e1f-697">&lt;optional&gt;</span></span>|<span data-ttu-id="22e1f-698">Wenn `true`: Zeigt an, dass die Anlage inline im Nachrichtentext angezeigt wird und nicht in der Anlagenlisten angezeigt werden soll.</span><span class="sxs-lookup"><span data-stu-id="22e1f-698">If `true`, indicates that the attachment will be shown inline in the message body, and should not be displayed in the attachment list.</span></span>|
|`callback`|<span data-ttu-id="22e1f-699">function</span><span class="sxs-lookup"><span data-stu-id="22e1f-699">function</span></span>|<span data-ttu-id="22e1f-700">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="22e1f-700">&lt;optional&gt;</span></span>|<span data-ttu-id="22e1f-701">Wenn die Methode abgeschlossen wird, wird die Funktion , die im `callback`-Parameter übergeben wird, mit einem einzelnen Parameter (`asyncResult`) aufgerufen, der ein [`AsyncResult`](/javascript/api/office/office.asyncresult)-Objekt darstellt.</span><span class="sxs-lookup"><span data-stu-id="22e1f-701">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span> <br/><span data-ttu-id="22e1f-702">Bei Erfolg wird der Anlagenbezeichner in der `asyncResult.value`-Eigenschaft bereitgestellt.</span><span class="sxs-lookup"><span data-stu-id="22e1f-702">On success, the attachment identifier will be provided in the `asyncResult.value` property.</span></span><br/><span data-ttu-id="22e1f-703">Wenn beim Hochladen der Anlage ein Fehler auftritt, enthält das `asyncResult`-Objekt ein `Error`-Objekt mit einer Beschreibung des Fehlers.</span><span class="sxs-lookup"><span data-stu-id="22e1f-703">If uploading the attachment fails, the `asyncResult` object will contain an `Error` object that provides a description of the error.</span></span>|

##### <a name="errors"></a><span data-ttu-id="22e1f-704">Fehler</span><span class="sxs-lookup"><span data-stu-id="22e1f-704">Errors</span></span>

|<span data-ttu-id="22e1f-705">Fehlercode</span><span class="sxs-lookup"><span data-stu-id="22e1f-705">Error code</span></span>|<span data-ttu-id="22e1f-706">Beschreibung</span><span class="sxs-lookup"><span data-stu-id="22e1f-706">Description</span></span>|
|------------|-------------|
|`AttachmentSizeExceeded`|<span data-ttu-id="22e1f-707">Die Anlage ist zu groß.</span><span class="sxs-lookup"><span data-stu-id="22e1f-707">The attachment is larger than allowed.</span></span>|
|`FileTypeNotSupported`|<span data-ttu-id="22e1f-708">Die Anlage enthält eine Erweiterung, die nicht zulässig ist.</span><span class="sxs-lookup"><span data-stu-id="22e1f-708">The attachment has an extension that is not allowed.</span></span>|
|`NumberOfAttachmentsExceeded`|<span data-ttu-id="22e1f-709">Die Nachricht oder der Termin enthält zu viele Anlagen.</span><span class="sxs-lookup"><span data-stu-id="22e1f-709">The message or appointment has too many attachments.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="22e1f-710">Anforderungen</span><span class="sxs-lookup"><span data-stu-id="22e1f-710">Requirements</span></span>

|<span data-ttu-id="22e1f-711">Anforderung</span><span class="sxs-lookup"><span data-stu-id="22e1f-711">Requirement</span></span>|<span data-ttu-id="22e1f-712">Wert</span><span class="sxs-lookup"><span data-stu-id="22e1f-712">Value</span></span>|
|---|---|
|[<span data-ttu-id="22e1f-713">Mindestversion des Postfachanforderungssatzes</span><span class="sxs-lookup"><span data-stu-id="22e1f-713">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="22e1f-714">1.1</span><span class="sxs-lookup"><span data-stu-id="22e1f-714">1.1</span></span>|
|[<span data-ttu-id="22e1f-715">Mindestberechtigungsstufe</span><span class="sxs-lookup"><span data-stu-id="22e1f-715">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="22e1f-716">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="22e1f-716">ReadWriteItem</span></span>|
|[<span data-ttu-id="22e1f-717">Zutreffender Outlook-Modus</span><span class="sxs-lookup"><span data-stu-id="22e1f-717">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="22e1f-718">Verfassen</span><span class="sxs-lookup"><span data-stu-id="22e1f-718">Compose</span></span>|

##### <a name="examples"></a><span data-ttu-id="22e1f-719">Beispiele</span><span class="sxs-lookup"><span data-stu-id="22e1f-719">Examples</span></span>

```js
function callback(result) {
  if (result.error) {
    showMessage(result.error);
  } else {
    showMessage("Attachment added");
  }
}

function addAttachment() {
  // The values in asyncContext can be accessed in the callback
  var options = { 'asyncContext': { var1: 1, var2: 2 } };

  var attachmentURL = "https://contoso.com/rtm/icon.png";
  Office.context.mailbox.item.addFileAttachmentAsync(attachmentURL, attachmentURL, options, callback);
}
```

<span data-ttu-id="22e1f-720">Im folgenden Beispiel wird eine Bilddatei als Inlineanlage hinzugefügt, und die Anlage wird im Nachrichtentext referenziert.</span><span class="sxs-lookup"><span data-stu-id="22e1f-720">The following example adds an image file as an inline attachment and references the attachment in the message body.</span></span>

```js
Office.context.mailbox.item.addFileAttachmentAsync
(
  "http://i.imgur.com/WJXklif.png",
  "cute_bird.png",
  {
    isInline: true
  },
  function (asyncResult) {
    Office.context.mailbox.item.body.setAsync(
      "<p>Here's a cute bird!</p><img src='cid:cute_bird.png'>",
      {
        "coercionType": "html"
      },
      function (asyncResult) {
        
      }
    );
  }
);
```

#### <a name="addfileattachmentfrombase64asyncbase64file-attachmentname-options-callback"></a><span data-ttu-id="22e1f-721">addFileAttachmentFromBase64Async (base64File, AttachmentName, [Optionen], [Rückruf])</span><span class="sxs-lookup"><span data-stu-id="22e1f-721">addFileAttachmentFromBase64Async(base64File, attachmentName, [options], [callback])</span></span>

<span data-ttu-id="22e1f-722">Fügt eine Datei aus der base64-Codierung in einer Nachricht oder eines Termins als Anlage.</span><span class="sxs-lookup"><span data-stu-id="22e1f-722">Adds a file from the base64 encoding to a message or appointment as an attachment.</span></span>

<span data-ttu-id="22e1f-723">Die `addFileAttachmentFromBase64Async` -Methode die Datei aus der base64-Codierung und fügt diese an das Element in dem Formular zum Verfassen.</span><span class="sxs-lookup"><span data-stu-id="22e1f-723">The `addFileAttachmentFromBase64Async` method uploads the file from the base64 encoding and attaches it to the item in the compose form.</span></span> <span data-ttu-id="22e1f-724">Diese Methode gibt die Anlage-ID in der AsyncResult.value-Objekt zurück.</span><span class="sxs-lookup"><span data-stu-id="22e1f-724">This method returns the attachment identifier in the AsyncResult.value object.</span></span>

<span data-ttu-id="22e1f-725">Anschließend können Sie den Bezeichner mit der [`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback)-Methode in der gleichen Sitzung zum Entfernen der Anlage verwenden.</span><span class="sxs-lookup"><span data-stu-id="22e1f-725">You can subsequently use the identifier with the [`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback) method to remove the attachment in the same session.</span></span>

##### <a name="parameters"></a><span data-ttu-id="22e1f-726">Parameter:</span><span class="sxs-lookup"><span data-stu-id="22e1f-726">Parameters:</span></span>
|<span data-ttu-id="22e1f-727">Name</span><span class="sxs-lookup"><span data-stu-id="22e1f-727">Name</span></span>|<span data-ttu-id="22e1f-728">Typ</span><span class="sxs-lookup"><span data-stu-id="22e1f-728">Type</span></span>|<span data-ttu-id="22e1f-729">Attribute</span><span class="sxs-lookup"><span data-stu-id="22e1f-729">Attributes</span></span>|<span data-ttu-id="22e1f-730">Beschreibung</span><span class="sxs-lookup"><span data-stu-id="22e1f-730">Description</span></span>|
|---|---|---|---|
|`base64File`|<span data-ttu-id="22e1f-731">Zeichenfolge</span><span class="sxs-lookup"><span data-stu-id="22e1f-731">String</span></span>||<span data-ttu-id="22e1f-732">Die base64-codierte Inhalt eines Bilds oder eine Datei, eine e-Mail oder ein Ereignis hinzugefügt werden soll.</span><span class="sxs-lookup"><span data-stu-id="22e1f-732">The base64 encoded content of an image or file to be added to an email or event.</span></span>|
|`attachmentName`|<span data-ttu-id="22e1f-733">Zeichenfolge</span><span class="sxs-lookup"><span data-stu-id="22e1f-733">String</span></span>||<span data-ttu-id="22e1f-p137">Der Name der Anlage, der beim Hochladen der Anlage angezeigt wird. Die maximale Länge ist 255 Zeichen.</span><span class="sxs-lookup"><span data-stu-id="22e1f-p137">The name of the attachment that is shown while the attachment is uploading. The maximum length is 255 characters.</span></span>|
|`options`|<span data-ttu-id="22e1f-736">Object</span><span class="sxs-lookup"><span data-stu-id="22e1f-736">Object</span></span>|<span data-ttu-id="22e1f-737">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="22e1f-737">&lt;optional&gt;</span></span>|<span data-ttu-id="22e1f-738">Ein Objektliteral, das mindestens eine der folgenden Eigenschaften enthält.</span><span class="sxs-lookup"><span data-stu-id="22e1f-738">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`|<span data-ttu-id="22e1f-739">Object</span><span class="sxs-lookup"><span data-stu-id="22e1f-739">Object</span></span>|<span data-ttu-id="22e1f-740">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="22e1f-740">&lt;optional&gt;</span></span>|<span data-ttu-id="22e1f-741">Entwickler können jedes Objekt angeben, auf das sie in der Rückrufmethode zugreifen möchten.</span><span class="sxs-lookup"><span data-stu-id="22e1f-741">Developers can provide any object they wish to access in the callback method.</span></span>|
|`options.isInline`|<span data-ttu-id="22e1f-742">Boolean</span><span class="sxs-lookup"><span data-stu-id="22e1f-742">Boolean</span></span>|<span data-ttu-id="22e1f-743">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="22e1f-743">&lt;optional&gt;</span></span>|<span data-ttu-id="22e1f-744">Wenn `true`: Zeigt an, dass die Anlage inline im Nachrichtentext angezeigt wird und nicht in der Anlagenlisten angezeigt werden soll.</span><span class="sxs-lookup"><span data-stu-id="22e1f-744">If `true`, indicates that the attachment will be shown inline in the message body, and should not be displayed in the attachment list.</span></span>|
|`callback`|<span data-ttu-id="22e1f-745">function</span><span class="sxs-lookup"><span data-stu-id="22e1f-745">function</span></span>|<span data-ttu-id="22e1f-746">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="22e1f-746">&lt;optional&gt;</span></span>|<span data-ttu-id="22e1f-747">Wenn die Methode abgeschlossen wird, wird die Funktion , die im `callback`-Parameter übergeben wird, mit einem einzelnen Parameter (`asyncResult`) aufgerufen, der ein [`AsyncResult`](/javascript/api/office/office.asyncresult)-Objekt darstellt.</span><span class="sxs-lookup"><span data-stu-id="22e1f-747">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span> <br/><span data-ttu-id="22e1f-748">Bei Erfolg wird der Anlagenbezeichner in der `asyncResult.value`-Eigenschaft bereitgestellt.</span><span class="sxs-lookup"><span data-stu-id="22e1f-748">On success, the attachment identifier will be provided in the `asyncResult.value` property.</span></span><br/><span data-ttu-id="22e1f-749">Wenn beim Hochladen der Anlage ein Fehler auftritt, enthält das `asyncResult`-Objekt ein `Error`-Objekt mit einer Beschreibung des Fehlers.</span><span class="sxs-lookup"><span data-stu-id="22e1f-749">If uploading the attachment fails, the `asyncResult` object will contain an `Error` object that provides a description of the error.</span></span>|

##### <a name="errors"></a><span data-ttu-id="22e1f-750">Fehler</span><span class="sxs-lookup"><span data-stu-id="22e1f-750">Errors</span></span>

|<span data-ttu-id="22e1f-751">Fehlercode</span><span class="sxs-lookup"><span data-stu-id="22e1f-751">Error code</span></span>|<span data-ttu-id="22e1f-752">Beschreibung</span><span class="sxs-lookup"><span data-stu-id="22e1f-752">Description</span></span>|
|------------|-------------|
|`AttachmentSizeExceeded`|<span data-ttu-id="22e1f-753">Die Anlage ist zu groß.</span><span class="sxs-lookup"><span data-stu-id="22e1f-753">The attachment is larger than allowed.</span></span>|
|`FileTypeNotSupported`|<span data-ttu-id="22e1f-754">Die Anlage enthält eine Erweiterung, die nicht zulässig ist.</span><span class="sxs-lookup"><span data-stu-id="22e1f-754">The attachment has an extension that is not allowed.</span></span>|
|`NumberOfAttachmentsExceeded`|<span data-ttu-id="22e1f-755">Die Nachricht oder der Termin enthält zu viele Anlagen.</span><span class="sxs-lookup"><span data-stu-id="22e1f-755">The message or appointment has too many attachments.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="22e1f-756">Anforderungen</span><span class="sxs-lookup"><span data-stu-id="22e1f-756">Requirements</span></span>

|<span data-ttu-id="22e1f-757">Anforderung</span><span class="sxs-lookup"><span data-stu-id="22e1f-757">Requirement</span></span>|<span data-ttu-id="22e1f-758">Wert</span><span class="sxs-lookup"><span data-stu-id="22e1f-758">Value</span></span>|
|---|---|
|[<span data-ttu-id="22e1f-759">Mindestversion des Postfachanforderungssatzes</span><span class="sxs-lookup"><span data-stu-id="22e1f-759">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="22e1f-760">Vorschau</span><span class="sxs-lookup"><span data-stu-id="22e1f-760">Preview</span></span>|
|[<span data-ttu-id="22e1f-761">Mindestberechtigungsstufe</span><span class="sxs-lookup"><span data-stu-id="22e1f-761">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="22e1f-762">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="22e1f-762">ReadWriteItem</span></span>|
|[<span data-ttu-id="22e1f-763">Zutreffender Outlook-Modus</span><span class="sxs-lookup"><span data-stu-id="22e1f-763">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="22e1f-764">Verfassen</span><span class="sxs-lookup"><span data-stu-id="22e1f-764">Compose</span></span>|

##### <a name="examples"></a><span data-ttu-id="22e1f-765">Beispiele</span><span class="sxs-lookup"><span data-stu-id="22e1f-765">Examples</span></span>

```js
Office.context.mailbox.item.addFileAttachmentFromBase64Async(
  base64String,
  "cute_bird.png",
  {
    isInline: true
  },
  function (asyncResult) {
    Office.context.mailbox.item.body.setAsync(
      "<p>Here's a cute bird!</p><img src='cid:cute_bird.png'>",
      {
        "coercionType": "html"
      },
      function (asyncResult) {
      }
    );
  }
);
```

####  <a name="addhandlerasynceventtype-handler-options-callback"></a><span data-ttu-id="22e1f-766">addHandlerAsync(eventType, handler, [options], [callback])</span><span class="sxs-lookup"><span data-stu-id="22e1f-766">addHandlerAsync(eventType, handler, [options], [callback])</span></span>

<span data-ttu-id="22e1f-767">Fügt einen Ereignishandler für ein unterstütztes Ereignis hinzu.</span><span class="sxs-lookup"><span data-stu-id="22e1f-767">Adds an event handler for a supported event.</span></span>

<span data-ttu-id="22e1f-768">Derzeit sind die unterstützten Ereignistypen `Office.EventType.AppointmentTimeChanged`, `Office.EventType.RecipientsChanged`, und`Office.EventType.RecurrencePatternChanged`</span><span class="sxs-lookup"><span data-stu-id="22e1f-768">Currently the supported event types are `Office.EventType.AppointmentTimeChanged`, `Office.EventType.RecipientsChanged`, and `Office.EventType.RecurrencePatternChanged`</span></span>

##### <a name="parameters"></a><span data-ttu-id="22e1f-769">Parameter:</span><span class="sxs-lookup"><span data-stu-id="22e1f-769">Parameters:</span></span>

| <span data-ttu-id="22e1f-770">Name</span><span class="sxs-lookup"><span data-stu-id="22e1f-770">Name</span></span> | <span data-ttu-id="22e1f-771">Typ</span><span class="sxs-lookup"><span data-stu-id="22e1f-771">Type</span></span> | <span data-ttu-id="22e1f-772">Attribute</span><span class="sxs-lookup"><span data-stu-id="22e1f-772">Attributes</span></span> | <span data-ttu-id="22e1f-773">Beschreibung</span><span class="sxs-lookup"><span data-stu-id="22e1f-773">Description</span></span> |
|---|---|---|---|
| `eventType` | [<span data-ttu-id="22e1f-774">Office.EventType</span><span class="sxs-lookup"><span data-stu-id="22e1f-774">Office.EventType</span></span>](office.md#eventtype-string) || <span data-ttu-id="22e1f-775">Das Ereignis, das den Handler aufrufen soll</span><span class="sxs-lookup"><span data-stu-id="22e1f-775">The event that should invoke the handler.</span></span> |
| `handler` | <span data-ttu-id="22e1f-776">Function</span><span class="sxs-lookup"><span data-stu-id="22e1f-776">Function</span></span> || <span data-ttu-id="22e1f-p138">Die Funktion, die das Ereignis behandeln soll. Die Funktion muss einen einzigen Parameter akzeptieren (ein Objektliteral). Die Eigenschaft `type` dieses Parameters entspricht dem Parameter `eventType`, der an `addHandlerAsync` übergeben wird.</span><span class="sxs-lookup"><span data-stu-id="22e1f-p138">The function to handle the event. The function must accept a single parameter, which is an object literal. The `type` property on the parameter will match the `eventType` parameter passed to `addHandlerAsync`.</span></span> |
| `options` | <span data-ttu-id="22e1f-780">Objekt</span><span class="sxs-lookup"><span data-stu-id="22e1f-780">Object</span></span> | <span data-ttu-id="22e1f-781">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="22e1f-781">&lt;optional&gt;</span></span> | <span data-ttu-id="22e1f-782">Ein Objektliteral, das mindestens eine der folgenden Eigenschaften enthält.</span><span class="sxs-lookup"><span data-stu-id="22e1f-782">An object literal that contains one or more of the following properties.</span></span> |
| `options.asyncContext` | <span data-ttu-id="22e1f-783">Object</span><span class="sxs-lookup"><span data-stu-id="22e1f-783">Object</span></span> | <span data-ttu-id="22e1f-784">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="22e1f-784">&lt;optional&gt;</span></span> | <span data-ttu-id="22e1f-785">Entwickler können ein Objekt bereitstellen, auf das sie in der Callbackmethode zugreifen möchten.</span><span class="sxs-lookup"><span data-stu-id="22e1f-785">Developers can provide any object they wish to access in the callback method.</span></span> |
| `callback` | <span data-ttu-id="22e1f-786">Funktion</span><span class="sxs-lookup"><span data-stu-id="22e1f-786">function</span></span>| <span data-ttu-id="22e1f-787">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="22e1f-787">&lt;optional&gt;</span></span>|<span data-ttu-id="22e1f-788">Wenn die Methode abgeschlossen wird, wird die Funktion , die im `callback`-Parameter übergeben wird, mit einem einzelnen Parameter (`asyncResult`) aufgerufen, der ein [`AsyncResult`](/javascript/api/office/office.asyncresult)-Objekt darstellt.</span><span class="sxs-lookup"><span data-stu-id="22e1f-788">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="22e1f-789">Anforderungen</span><span class="sxs-lookup"><span data-stu-id="22e1f-789">Requirements</span></span>

|<span data-ttu-id="22e1f-790">Anforderung</span><span class="sxs-lookup"><span data-stu-id="22e1f-790">Requirement</span></span>| <span data-ttu-id="22e1f-791">Wert</span><span class="sxs-lookup"><span data-stu-id="22e1f-791">Value</span></span>|
|---|---|
|[<span data-ttu-id="22e1f-792">Mindestversion des Postfachanforderungssatzes</span><span class="sxs-lookup"><span data-stu-id="22e1f-792">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="22e1f-793">Vorschau</span><span class="sxs-lookup"><span data-stu-id="22e1f-793">Preview</span></span> |
|[<span data-ttu-id="22e1f-794">Mindestberechtigungsstufe</span><span class="sxs-lookup"><span data-stu-id="22e1f-794">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="22e1f-795">ReadItem</span><span class="sxs-lookup"><span data-stu-id="22e1f-795">ReadItem</span></span> |
|[<span data-ttu-id="22e1f-796">Zutreffender Outlook-Modus</span><span class="sxs-lookup"><span data-stu-id="22e1f-796">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="22e1f-797">Verfassen oder Lesen</span><span class="sxs-lookup"><span data-stu-id="22e1f-797">Compose or read</span></span> |

##### <a name="example"></a><span data-ttu-id="22e1f-798">Beispiel</span><span class="sxs-lookup"><span data-stu-id="22e1f-798">Example</span></span>

```
Office.initialize = function (reason) {
  $(document).ready(function () {
    Office.context.mailbox.item.addHandlerAsync(Office.EventType.RecurrencePatternChanged, loadNewItem, function (result) {
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

####  <a name="additemattachmentasyncitemid-attachmentname-options-callback"></a><span data-ttu-id="22e1f-799">addItemAttachmentAsync(itemId, attachmentName, [options], [callback])</span><span class="sxs-lookup"><span data-stu-id="22e1f-799">addItemAttachmentAsync(itemId, attachmentName, [options], [callback])</span></span>

<span data-ttu-id="22e1f-800">Fügt der Nachricht oder dem Termin ein Exchange-Objekt, wie z. B. eine Nachricht, als Anhang hinzu.</span><span class="sxs-lookup"><span data-stu-id="22e1f-800">Adds an Exchange item, such as a message, as an attachment to the message or appointment.</span></span>

<span data-ttu-id="22e1f-p139">Die `addItemAttachmentAsync`-Methode fügt das Element mit dem angegebenen Exchange-Bezeichner an das Element im Formular zum Verfassen an. Wenn Sie eine Rückrufmethode angeben, wird die Methode mit einem `asyncResult`-Parameter aufgerufen, der entweder den Anlagenbezeichner oder einen Code enthält, der etwaige Fehler angibt, die beim Anfügen des Objekts aufgetreten sind. Sie können ggf. den `options`-Parameter verwenden, um Statusinformationen an die Rückrufmethode zu übergeben.</span><span class="sxs-lookup"><span data-stu-id="22e1f-p139">The `addItemAttachmentAsync` method attaches the item with the specified Exchange identifier to the item in the compose form. If you specify a callback method, the method is called with one parameter, `asyncResult`, which contains either the attachment identifier or a code that indicates any error that occurred while attaching the item. You can use the `options` parameter to pass state information to the callback method, if needed.</span></span>

<span data-ttu-id="22e1f-804">Anschließend können Sie den Bezeichner mit der [`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback)-Methode in der gleichen Sitzung zum Entfernen der Anlage verwenden.</span><span class="sxs-lookup"><span data-stu-id="22e1f-804">You can subsequently use the identifier with the [`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback) method to remove the attachment in the same session.</span></span>

<span data-ttu-id="22e1f-805">Wenn Ihre Office-Add-Ins in Outlook Web App ausgeführt wird die `addItemAttachmentAsync` -Methode kann Anfügen von Elementen in der Elemente, die Sie bearbeiten; Allerdings Dies wird nicht unterstützt und wird nicht empfohlen.</span><span class="sxs-lookup"><span data-stu-id="22e1f-805">If your Office Add-in is running in Outlook Web App, the `addItemAttachmentAsync` method can attach items to items other than the item that you are editing; however, this is not supported and is not recommended.</span></span>

##### <a name="parameters"></a><span data-ttu-id="22e1f-806">Parameter:</span><span class="sxs-lookup"><span data-stu-id="22e1f-806">Parameters:</span></span>

|<span data-ttu-id="22e1f-807">Name</span><span class="sxs-lookup"><span data-stu-id="22e1f-807">Name</span></span>|<span data-ttu-id="22e1f-808">Typ</span><span class="sxs-lookup"><span data-stu-id="22e1f-808">Type</span></span>|<span data-ttu-id="22e1f-809">Attribute</span><span class="sxs-lookup"><span data-stu-id="22e1f-809">Attributes</span></span>|<span data-ttu-id="22e1f-810">Beschreibung</span><span class="sxs-lookup"><span data-stu-id="22e1f-810">Description</span></span>|
|---|---|---|---|
|`itemId`|<span data-ttu-id="22e1f-811">Zeichenfolge</span><span class="sxs-lookup"><span data-stu-id="22e1f-811">String</span></span>||<span data-ttu-id="22e1f-p140">Der Exchange-Bezeichner des Objekts, das angehängt werden soll. Die maximale Länge beträgt 100 Zeichen.</span><span class="sxs-lookup"><span data-stu-id="22e1f-p140">The Exchange identifier of the item to attach. The maximum length is 100 characters.</span></span>|
|`attachmentName`|<span data-ttu-id="22e1f-814">Zeichenfolge</span><span class="sxs-lookup"><span data-stu-id="22e1f-814">String</span></span>||<span data-ttu-id="22e1f-p141">Der Betreff des Elements, das angehängt werden soll. Die maximale Länge ist 255 Zeichen.</span><span class="sxs-lookup"><span data-stu-id="22e1f-p141">The sujbect of the item to be attached. The maximum length is 255 characters.</span></span>|
|`options`|<span data-ttu-id="22e1f-817">Object</span><span class="sxs-lookup"><span data-stu-id="22e1f-817">Object</span></span>|<span data-ttu-id="22e1f-818">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="22e1f-818">&lt;optional&gt;</span></span>|<span data-ttu-id="22e1f-819">Ein Objektliteral, das mindestens eine der folgenden Eigenschaften enthält.</span><span class="sxs-lookup"><span data-stu-id="22e1f-819">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`|<span data-ttu-id="22e1f-820">Object</span><span class="sxs-lookup"><span data-stu-id="22e1f-820">Object</span></span>|<span data-ttu-id="22e1f-821">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="22e1f-821">&lt;optional&gt;</span></span>|<span data-ttu-id="22e1f-822">Entwickler können ein Objekt bereitstellen, auf das sie in der Callbackmethode zugreifen möchten.</span><span class="sxs-lookup"><span data-stu-id="22e1f-822">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`|<span data-ttu-id="22e1f-823">Funktion</span><span class="sxs-lookup"><span data-stu-id="22e1f-823">function</span></span>|<span data-ttu-id="22e1f-824">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="22e1f-824">&lt;optional&gt;</span></span>|<span data-ttu-id="22e1f-825">Wenn die Methode abgeschlossen wird, wird die Funktion , die im `callback`-Parameter übergeben wird, mit einem einzelnen Parameter (`asyncResult`) aufgerufen, der ein [`AsyncResult`](/javascript/api/office/office.asyncresult)-Objekt darstellt.</span><span class="sxs-lookup"><span data-stu-id="22e1f-825">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span> <br/><span data-ttu-id="22e1f-826">Bei Erfolg wird der Anlagenbezeichner in der `asyncResult.value`-Eigenschaft bereitgestellt.</span><span class="sxs-lookup"><span data-stu-id="22e1f-826">On success, the attachment identifier will be provided in the `asyncResult.value` property.</span></span><br/><span data-ttu-id="22e1f-827">Wenn beim Hinzufügen der Anlage ein Fehler auftritt, enthält das `asyncResult`-Objekt ein `Error`-Objekt mit einer Beschreibung des Fehlers.</span><span class="sxs-lookup"><span data-stu-id="22e1f-827">If adding the attachment fails, the `asyncResult` object will contain an `Error` object that provides a description of the error.</span></span>|

##### <a name="errors"></a><span data-ttu-id="22e1f-828">Fehler</span><span class="sxs-lookup"><span data-stu-id="22e1f-828">Errors</span></span>

|<span data-ttu-id="22e1f-829">Fehlercode</span><span class="sxs-lookup"><span data-stu-id="22e1f-829">Error code</span></span>|<span data-ttu-id="22e1f-830">Beschreibung</span><span class="sxs-lookup"><span data-stu-id="22e1f-830">Description</span></span>|
|------------|-------------|
|`NumberOfAttachmentsExceeded`|<span data-ttu-id="22e1f-831">Die Nachricht oder der Termin enthält zu viele Anlagen.</span><span class="sxs-lookup"><span data-stu-id="22e1f-831">The message or appointment has too many attachments.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="22e1f-832">Anforderungen</span><span class="sxs-lookup"><span data-stu-id="22e1f-832">Requirements</span></span>

|<span data-ttu-id="22e1f-833">Anforderung</span><span class="sxs-lookup"><span data-stu-id="22e1f-833">Requirement</span></span>|<span data-ttu-id="22e1f-834">Wert</span><span class="sxs-lookup"><span data-stu-id="22e1f-834">Value</span></span>|
|---|---|
|[<span data-ttu-id="22e1f-835">Mindestversion des Postfachanforderungssatzes</span><span class="sxs-lookup"><span data-stu-id="22e1f-835">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="22e1f-836">1.1</span><span class="sxs-lookup"><span data-stu-id="22e1f-836">1.1</span></span>|
|[<span data-ttu-id="22e1f-837">Mindestberechtigungsstufe</span><span class="sxs-lookup"><span data-stu-id="22e1f-837">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="22e1f-838">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="22e1f-838">ReadWriteItem</span></span>|
|[<span data-ttu-id="22e1f-839">Zutreffender Outlook-Modus</span><span class="sxs-lookup"><span data-stu-id="22e1f-839">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="22e1f-840">Verfassen</span><span class="sxs-lookup"><span data-stu-id="22e1f-840">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="22e1f-841">Beispiel</span><span class="sxs-lookup"><span data-stu-id="22e1f-841">Example</span></span>

<span data-ttu-id="22e1f-842">Im folgenden Beispiel wird ein vorhandenes Outlook-Element als Anlage mit dem Namen `My Attachment` hinzugefügt.</span><span class="sxs-lookup"><span data-stu-id="22e1f-842">The following example adds an existing Outlook item as an attachment with the name `My Attachment`.</span></span>

```
function callback(result) {
  if (result.error) {
    showMessage(result.error);
  } else {
    showMessage("Attachment added");
  }
}

function addAttachment() {
  // EWS ID of item to attach
  // (Shortened for readability)
  var itemId = "AAMkADI1...AAA=";

  // The values in asyncContext can be accessed in the callback
  var options = { 'asyncContext': { var1: 1, var2: 2 } };

  Office.context.mailbox.item.addItemAttachmentAsync(itemId, "My Attachment", options, callback);
}
```

####  <a name="close"></a><span data-ttu-id="22e1f-843">close()</span><span class="sxs-lookup"><span data-stu-id="22e1f-843">close()</span></span>

<span data-ttu-id="22e1f-844">Schließt das aktuelle Element, das gerade verfasst wird.</span><span class="sxs-lookup"><span data-stu-id="22e1f-844">Closes the current item that is being composed.</span></span>

<span data-ttu-id="22e1f-p142">Das Verhalten der `close`-Methode hängt vom aktuellen Status des verfassten Elements ab. Wenn das Element über nicht gespeicherte Änderungen verfügt, fordert der Client den Benutzer auf, die Schließenaktion zu speichern, zu verwerfen oder abzubrechen.</span><span class="sxs-lookup"><span data-stu-id="22e1f-p142">The behavior of the `close` method depends on the current state of the item being composed. If the item has unsaved changes, the client prompts the user to save, discard, or cancel the close action.</span></span>

> [!NOTE]
> <span data-ttu-id="22e1f-847">In Outlook im Web, wenn das Element einen Termin handelt, und es wurde bereits zuvor gespeichert wurde mit `saveAsync`, der Benutzer wird aufgefordert, speichern, löschen oder Abbrechen, auch wenn keine Änderungen aufgetreten sind, da das Element zuletzt gespeichert.</span><span class="sxs-lookup"><span data-stu-id="22e1f-847">In Outlook on the web, if the item is an appointment and it has previously been saved using `saveAsync`, the user is prompted to save, discard, or cancel even if no changes have occurred since the item was last saved.</span></span>

<span data-ttu-id="22e1f-848">Wenn die Nachricht im Outlook-Desktopclient eine Inlineantwort ist, hat die `close`-Methode keine Auswirkung.</span><span class="sxs-lookup"><span data-stu-id="22e1f-848">In the Outlook desktop client, if the message is an inline reply, the `close` method has no effect.</span></span>

##### <a name="requirements"></a><span data-ttu-id="22e1f-849">Anforderungen</span><span class="sxs-lookup"><span data-stu-id="22e1f-849">Requirements</span></span>

|<span data-ttu-id="22e1f-850">Anforderung</span><span class="sxs-lookup"><span data-stu-id="22e1f-850">Requirement</span></span>|<span data-ttu-id="22e1f-851">Wert</span><span class="sxs-lookup"><span data-stu-id="22e1f-851">Value</span></span>|
|---|---|
|[<span data-ttu-id="22e1f-852">Mindestversion des Postfachanforderungssatzes</span><span class="sxs-lookup"><span data-stu-id="22e1f-852">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="22e1f-853">1.3</span><span class="sxs-lookup"><span data-stu-id="22e1f-853">1.3</span></span>|
|[<span data-ttu-id="22e1f-854">Mindestberechtigungsstufe</span><span class="sxs-lookup"><span data-stu-id="22e1f-854">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="22e1f-855">Eingeschränkt</span><span class="sxs-lookup"><span data-stu-id="22e1f-855">Restricted</span></span>|
|[<span data-ttu-id="22e1f-856">Zutreffender Outlook-Modus</span><span class="sxs-lookup"><span data-stu-id="22e1f-856">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="22e1f-857">Verfassen</span><span class="sxs-lookup"><span data-stu-id="22e1f-857">Compose</span></span>|

#### <a name="displayreplyallformformdata"></a><span data-ttu-id="22e1f-858">displayReplyAllForm(formData)</span><span class="sxs-lookup"><span data-stu-id="22e1f-858">displayReplyAllForm(formData)</span></span>

<span data-ttu-id="22e1f-859">Zeigt ein Antwortformular an, das den Absender und alle Empfänger der ausgewählten Nachricht oder des Organisators und alle Teilnehmer des ausgewählten Termins enthält.</span><span class="sxs-lookup"><span data-stu-id="22e1f-859">Displays a reply form that includes the sender and all recipients of the selected message or the organizer and all attendees of the selected appointment.</span></span>

> [!NOTE]
> <span data-ttu-id="22e1f-860">Diese Methode wird in Outlook für iOS oder Outlook für Android nicht unterstützt.</span><span class="sxs-lookup"><span data-stu-id="22e1f-860">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="22e1f-861">In Outlook Web App wird das Antwortformular als Popupformular in der Dreispaltenansicht und als Popupformular in der Zwei- oder Einspaltenansicht angezeigt.</span><span class="sxs-lookup"><span data-stu-id="22e1f-861">In Outlook Web App, the reply form is displayed as a pop-out form in the 3-column view and a pop-up form in the 2- or 1-column view.</span></span>

<span data-ttu-id="22e1f-862">Wenn einer der Zeichenfolgenparameter seinen Grenzwert überschreitet, löst `displayReplyAllForm` eine Ausnahme aus.</span><span class="sxs-lookup"><span data-stu-id="22e1f-862">If any of the string parameters exceed their limits, `displayReplyAllForm` throws an exception.</span></span>

<span data-ttu-id="22e1f-p143">Wenn Anlagen im `formData.attachments`-Parameter angegeben werden, versuchen Outlook und Outlook Web App alle Anlagen herunterzuladen und sie an das Antwortformular anzufügen. Wenn Anlagen nicht hinzugefügt werden können, wird in der Formularbenutzeroberfläche ein Fehler ausgegeben. Wenn dies nicht möglich ist, wird keine Fehlermeldung ausgegeben.</span><span class="sxs-lookup"><span data-stu-id="22e1f-p143">When attachments are specified in the `formData.attachments` parameter, Outlook and Outlook Web App attempt to download all attachments and attach them to the reply form. If any attachments fail to be added, an error is shown in the form UI. If this isn't possible, then no error message is thrown.</span></span>

##### <a name="parameters"></a><span data-ttu-id="22e1f-866">Parameter:</span><span class="sxs-lookup"><span data-stu-id="22e1f-866">Parameters:</span></span>

|<span data-ttu-id="22e1f-867">Name</span><span class="sxs-lookup"><span data-stu-id="22e1f-867">Name</span></span>|<span data-ttu-id="22e1f-868">Typ</span><span class="sxs-lookup"><span data-stu-id="22e1f-868">Type</span></span>|<span data-ttu-id="22e1f-869">Attribute</span><span class="sxs-lookup"><span data-stu-id="22e1f-869">Attributes</span></span>|<span data-ttu-id="22e1f-870">Beschreibung</span><span class="sxs-lookup"><span data-stu-id="22e1f-870">Description</span></span>|
|---|---|---|---|
|`formData`|<span data-ttu-id="22e1f-871">Zeichenfolge &#124; Objekt</span><span class="sxs-lookup"><span data-stu-id="22e1f-871">String &#124; Object</span></span>||<span data-ttu-id="22e1f-p144">Eine Zeichenfolge, die Text- und HTML-Code enthält, die den Hauptteil des Antwortformulars darstellen. Die Zeichenfolge ist auf 32 KB beschränkt.</span><span class="sxs-lookup"><span data-stu-id="22e1f-p144">A string that contains text and HTML and that represents the body of the reply form. The string is limited to 32 KB.</span></span><br/><span data-ttu-id="22e1f-874">**ODER**</span><span class="sxs-lookup"><span data-stu-id="22e1f-874">**OR**</span></span><br/><span data-ttu-id="22e1f-p145">Ein Objekt, das Text- oder Anlagendaten und eine Rückruffunktion enthält. Das Objekt ist wie folgt definiert:</span><span class="sxs-lookup"><span data-stu-id="22e1f-p145">An object that contains body or attachment data and a callback function. The object is defined as follows.</span></span>|
|`formData.htmlBody`|<span data-ttu-id="22e1f-877">Zeichenfolge</span><span class="sxs-lookup"><span data-stu-id="22e1f-877">String</span></span>|<span data-ttu-id="22e1f-878">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="22e1f-878">&lt;optional&gt;</span></span>|<span data-ttu-id="22e1f-p146">Eine Zeichenfolge, die Text- und HTML-Code enthält, die den Hauptteil des Antwortformulars darstellen. Die Zeichenfolge ist auf 32 KB beschränkt.</span><span class="sxs-lookup"><span data-stu-id="22e1f-p146">A string that contains text and HTML and that represents the body of the reply form. The string is limited to 32 KB.</span></span>
|`formData.attachments`|<span data-ttu-id="22e1f-881">Array.&lt;Object&gt;</span><span class="sxs-lookup"><span data-stu-id="22e1f-881">Array.&lt;Object&gt;</span></span>|<span data-ttu-id="22e1f-882">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="22e1f-882">&lt;optional&gt;</span></span>|<span data-ttu-id="22e1f-883">Ein Array mit JSON-Objekten, die Datei- oder Elementanlagen sind.</span><span class="sxs-lookup"><span data-stu-id="22e1f-883">An array of JSON objects that are either file or item attachments.</span></span>|
|`formData.attachments.type`|<span data-ttu-id="22e1f-884">Zeichenfolge</span><span class="sxs-lookup"><span data-stu-id="22e1f-884">String</span></span>||<span data-ttu-id="22e1f-p147">Gibt den Typ der Anlage an. Muss `file` für eine Dateianlage oder `item` für eine Elementanlage sein.</span><span class="sxs-lookup"><span data-stu-id="22e1f-p147">Indicates the type of attachment. Must be `file` for a file attachment or `item` for an item attachment.</span></span>|
|`formData.attachments.name`|<span data-ttu-id="22e1f-887">Zeichenfolge</span><span class="sxs-lookup"><span data-stu-id="22e1f-887">String</span></span>||<span data-ttu-id="22e1f-888">Eine Zeichenfolge, die den Namen der Anlage mit bis zu 255 Zeichen enthält.</span><span class="sxs-lookup"><span data-stu-id="22e1f-888">A string that contains the name of the attachment, up to 255 characters in length.</span></span>|
|`formData.attachments.url`|<span data-ttu-id="22e1f-889">Zeichenfolge</span><span class="sxs-lookup"><span data-stu-id="22e1f-889">String</span></span>||<span data-ttu-id="22e1f-p148">Wird nur verwendet, wenn `type` auf `file` gesetzt ist. Der URI des Speicherorts für die Datei.</span><span class="sxs-lookup"><span data-stu-id="22e1f-p148">Only used if `type` is set to `file`. The URI of the location for the file.</span></span>|
|`formData.attachments.isInline`|<span data-ttu-id="22e1f-892">Boolean</span><span class="sxs-lookup"><span data-stu-id="22e1f-892">Boolean</span></span>||<span data-ttu-id="22e1f-p149">Wird nur verwendet, wenn `type` auf `file` gesetzt ist. Wenn `true`: Zeigt an, dass die Anlage inline im Nachrichtentext angezeigt wird und nicht in der Anlagenlisten angezeigt werden soll.</span><span class="sxs-lookup"><span data-stu-id="22e1f-p149">Only used if `type` is set to `file`. If `true`, indicates that the attachment will be shown inline in the message body, and should not be displayed in the attachment list.</span></span>|
|`formData.attachments.itemId`|<span data-ttu-id="22e1f-895">Zeichenfolge</span><span class="sxs-lookup"><span data-stu-id="22e1f-895">String</span></span>||<span data-ttu-id="22e1f-p150">Wird nur verwendet, wenn `type` auf `item` gesetzt ist. Die EWS-Element-ID der Anlage. Dies ist eine Zeichenfolge bis zu 100 Zeichen.</span><span class="sxs-lookup"><span data-stu-id="22e1f-p150">Only used if `type` is set to `item`. The EWS item id of the attachment. This is a string up to 100 characters.</span></span>|
|`callback`|<span data-ttu-id="22e1f-899">Funktion</span><span class="sxs-lookup"><span data-stu-id="22e1f-899">function</span></span>|<span data-ttu-id="22e1f-900">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="22e1f-900">&lt;optional&gt;</span></span>|<span data-ttu-id="22e1f-901">Bei Abschluss der Methode wird die im Parameter `callback` übergebene Funktion mit einem einzigen Parameter aufgerufen: `asyncResult`. Dieser Parameter ist ein Objekt des Typs [AsyncResult](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="22e1f-901">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [AsyncResult](/javascript/api/office/office.asyncresult) object.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="22e1f-902">Anforderungen</span><span class="sxs-lookup"><span data-stu-id="22e1f-902">Requirements</span></span>

|<span data-ttu-id="22e1f-903">Anforderung</span><span class="sxs-lookup"><span data-stu-id="22e1f-903">Requirement</span></span>|<span data-ttu-id="22e1f-904">Wert</span><span class="sxs-lookup"><span data-stu-id="22e1f-904">Value</span></span>|
|---|---|
|[<span data-ttu-id="22e1f-905">Mindestversion des Postfachanforderungssatzes</span><span class="sxs-lookup"><span data-stu-id="22e1f-905">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="22e1f-906">1.0</span><span class="sxs-lookup"><span data-stu-id="22e1f-906">1.0</span></span>|
|[<span data-ttu-id="22e1f-907">Mindestberechtigungsstufe</span><span class="sxs-lookup"><span data-stu-id="22e1f-907">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="22e1f-908">ReadItem</span><span class="sxs-lookup"><span data-stu-id="22e1f-908">ReadItem</span></span>|
|[<span data-ttu-id="22e1f-909">Zutreffender Outlook-Modus</span><span class="sxs-lookup"><span data-stu-id="22e1f-909">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="22e1f-910">Lesen</span><span class="sxs-lookup"><span data-stu-id="22e1f-910">Read</span></span>|

##### <a name="examples"></a><span data-ttu-id="22e1f-911">Beispiele</span><span class="sxs-lookup"><span data-stu-id="22e1f-911">Examples</span></span>

<span data-ttu-id="22e1f-912">Im folgenden Code wird eine Zeichenfolge an die `displayReplyAllForm`-Funktion übergeben.</span><span class="sxs-lookup"><span data-stu-id="22e1f-912">The following code passes a string to the `displayReplyAllForm` function.</span></span>

```
Office.context.mailbox.item.displayReplyAllForm('hello there');
Office.context.mailbox.item.displayReplyAllForm('<b>hello there</b>');
```

<span data-ttu-id="22e1f-913">Antworten mit einem leeren Textkörper.</span><span class="sxs-lookup"><span data-stu-id="22e1f-913">Reply with an empty body.</span></span>

```
Office.context.mailbox.item.displayReplyAllForm({});
```

<span data-ttu-id="22e1f-914">Antworten nur einem Textkörper.</span><span class="sxs-lookup"><span data-stu-id="22e1f-914">Reply with just a body.</span></span>

```
Office.context.mailbox.item.displayReplyAllForm(
{
  'htmlBody' : 'hi'
});
```

<span data-ttu-id="22e1f-915">Antworten mit einem Textkörper und einer Dateianlage.</span><span class="sxs-lookup"><span data-stu-id="22e1f-915">Reply with a body and a file attachment.</span></span>

```
Office.context.mailbox.item.displayReplyAllForm(
{
  'htmlBody' : 'hi',
  'attachments' :
  [
    {
      'type' : Office.MailboxEnums.AttachmentType.File,
      'name' : 'squirrel.png',
      'url' : 'http://i.imgur.com/sRgTlGR.jpg'
    }
  ]
});
```

<span data-ttu-id="22e1f-916">Antworten mit einem Textkörper und einer Elementanlage.</span><span class="sxs-lookup"><span data-stu-id="22e1f-916">Reply with a body and an item attachment.</span></span>

```
Office.context.mailbox.item.displayReplyAllForm(
{
  'htmlBody' : 'hi',
  'attachments' :
  [
    {
      'type' : 'item',
      'name' : 'rand',
      'itemId' : Office.context.mailbox.item.itemId
    }
  ]
});
```

<span data-ttu-id="22e1f-917">Antworten Sie mit einem Textkörper, einer Dateianlage, einer Elementanlage und einem Callback.</span><span class="sxs-lookup"><span data-stu-id="22e1f-917">Reply with a body, file attachment, item attachment, and a callback.</span></span>

```
Office.context.mailbox.item.displayReplyAllForm(
{
  'htmlBody' : 'hi',
  'attachments' :
  [
    {
      'type' : Office.MailboxEnums.AttachmentType.File,
      'name' : 'squirrel.png',
      'url' : 'http://i.imgur.com/sRgTlGR.jpg'
    },
    {
      'type' : 'item',
      'name' : 'rand',
      'itemId' : Office.context.mailbox.item.itemId
    }
  ],
  'callback' : function(asyncResult)
  {
    console.log(asyncResult.value);
  }
});
```

#### <a name="displayreplyformformdata"></a><span data-ttu-id="22e1f-918">displayReplyForm(formData)</span><span class="sxs-lookup"><span data-stu-id="22e1f-918">displayReplyForm(formData)</span></span>

<span data-ttu-id="22e1f-919">Zeigt ein Antwortformular an, das nur den Absender der ausgewählten Nachricht oder den Organisator des ausgewählten Termins enthält.</span><span class="sxs-lookup"><span data-stu-id="22e1f-919">Displays a reply form that includes only the sender of the selected message or the organizer of the selected appointment.</span></span>

> [!NOTE]
> <span data-ttu-id="22e1f-920">Diese Methode wird in Outlook für iOS oder Outlook für Android nicht unterstützt.</span><span class="sxs-lookup"><span data-stu-id="22e1f-920">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="22e1f-921">In Outlook Web App wird das Antwortformular als Popupformular in der Dreispaltenansicht und als Popupformular in der Zwei- oder Einspaltenansicht angezeigt.</span><span class="sxs-lookup"><span data-stu-id="22e1f-921">In Outlook Web App, the reply form is displayed as a pop-out form in the 3-column view and a pop-up form in the 2- or 1-column view.</span></span>

<span data-ttu-id="22e1f-922">Wenn einer der Zeichenfolgenparameter seinen Grenzwert überschreitet, löst `displayReplyForm` eine Ausnahme aus.</span><span class="sxs-lookup"><span data-stu-id="22e1f-922">If any of the string parameters exceed their limits, `displayReplyForm` throws an exception.</span></span>

<span data-ttu-id="22e1f-p151">Wenn Anlagen im `formData.attachments`-Parameter angegeben werden, versuchen Outlook und Outlook Web App alle Anlagen herunterzuladen und sie an das Antwortformular anzufügen. Wenn Anlagen nicht hinzugefügt werden können, wird in der Formularbenutzeroberfläche ein Fehler ausgegeben. Wenn dies nicht möglich ist, wird keine Fehlermeldung ausgegeben.</span><span class="sxs-lookup"><span data-stu-id="22e1f-p151">When attachments are specified in the `formData.attachments` parameter, Outlook and Outlook Web App attempt to download all attachments and attach them to the reply form. If any attachments fail to be added, an error is shown in the form UI. If this isn't possible, then no error message is thrown.</span></span>

##### <a name="parameters"></a><span data-ttu-id="22e1f-926">Parameter:</span><span class="sxs-lookup"><span data-stu-id="22e1f-926">Parameters:</span></span>

|<span data-ttu-id="22e1f-927">Name</span><span class="sxs-lookup"><span data-stu-id="22e1f-927">Name</span></span>|<span data-ttu-id="22e1f-928">Typ</span><span class="sxs-lookup"><span data-stu-id="22e1f-928">Type</span></span>|<span data-ttu-id="22e1f-929">Attribute</span><span class="sxs-lookup"><span data-stu-id="22e1f-929">Attributes</span></span>|<span data-ttu-id="22e1f-930">Beschreibung</span><span class="sxs-lookup"><span data-stu-id="22e1f-930">Description</span></span>|
|---|---|---|---|
|`formData`|<span data-ttu-id="22e1f-931">Zeichenfolge &#124; Objekt</span><span class="sxs-lookup"><span data-stu-id="22e1f-931">String &#124; Object</span></span>||<span data-ttu-id="22e1f-p152">Eine Zeichenfolge, die Text- und HTML-Code enthält, die den Hauptteil des Antwortformulars darstellen. Die Zeichenfolge ist auf 32 KB beschränkt.</span><span class="sxs-lookup"><span data-stu-id="22e1f-p152">A string that contains text and HTML and that represents the body of the reply form. The string is limited to 32 KB.</span></span><br/><span data-ttu-id="22e1f-934">**ODER**</span><span class="sxs-lookup"><span data-stu-id="22e1f-934">**OR**</span></span><br/><span data-ttu-id="22e1f-p153">Ein Objekt, das Text- oder Anlagendaten und eine Rückruffunktion enthält. Das Objekt ist wie folgt definiert:</span><span class="sxs-lookup"><span data-stu-id="22e1f-p153">An object that contains body or attachment data and a callback function. The object is defined as follows.</span></span>|
|`formData.htmlBody`|<span data-ttu-id="22e1f-937">Zeichenfolge</span><span class="sxs-lookup"><span data-stu-id="22e1f-937">String</span></span>|<span data-ttu-id="22e1f-938">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="22e1f-938">&lt;optional&gt;</span></span>|<span data-ttu-id="22e1f-p154">Eine Zeichenfolge, die Text- und HTML-Code enthält, die den Hauptteil des Antwortformulars darstellen. Die Zeichenfolge ist auf 32 KB beschränkt.</span><span class="sxs-lookup"><span data-stu-id="22e1f-p154">A string that contains text and HTML and that represents the body of the reply form. The string is limited to 32 KB.</span></span>
|`formData.attachments`|<span data-ttu-id="22e1f-941">Array.&lt;Object&gt;</span><span class="sxs-lookup"><span data-stu-id="22e1f-941">Array.&lt;Object&gt;</span></span>|<span data-ttu-id="22e1f-942">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="22e1f-942">&lt;optional&gt;</span></span>|<span data-ttu-id="22e1f-943">Ein Array mit JSON-Objekten, die Datei- oder Elementanlagen sind.</span><span class="sxs-lookup"><span data-stu-id="22e1f-943">An array of JSON objects that are either file or item attachments.</span></span>|
|`formData.attachments.type`|<span data-ttu-id="22e1f-944">Zeichenfolge</span><span class="sxs-lookup"><span data-stu-id="22e1f-944">String</span></span>||<span data-ttu-id="22e1f-p155">Gibt den Typ der Anlage an. Muss `file` für eine Dateianlage oder `item` für eine Elementanlage sein.</span><span class="sxs-lookup"><span data-stu-id="22e1f-p155">Indicates the type of attachment. Must be `file` for a file attachment or `item` for an item attachment.</span></span>|
|`formData.attachments.name`|<span data-ttu-id="22e1f-947">Zeichenfolge</span><span class="sxs-lookup"><span data-stu-id="22e1f-947">String</span></span>||<span data-ttu-id="22e1f-948">Eine Zeichenfolge, die den Namen der Anlage mit bis zu 255 Zeichen enthält.</span><span class="sxs-lookup"><span data-stu-id="22e1f-948">A string that contains the name of the attachment, up to 255 characters in length.</span></span>|
|`formData.attachments.url`|<span data-ttu-id="22e1f-949">Zeichenfolge</span><span class="sxs-lookup"><span data-stu-id="22e1f-949">String</span></span>||<span data-ttu-id="22e1f-p156">Wird nur verwendet, wenn `type` auf `file` gesetzt ist. Der URI des Speicherorts für die Datei.</span><span class="sxs-lookup"><span data-stu-id="22e1f-p156">Only used if `type` is set to `file`. The URI of the location for the file.</span></span>|
|`formData.attachments.isInline`|<span data-ttu-id="22e1f-952">Boolean</span><span class="sxs-lookup"><span data-stu-id="22e1f-952">Boolean</span></span>||<span data-ttu-id="22e1f-p157">Wird nur verwendet, wenn `type` auf `file` gesetzt ist. Wenn `true`: Zeigt an, dass die Anlage inline im Nachrichtentext angezeigt wird und nicht in der Anlagenlisten angezeigt werden soll.</span><span class="sxs-lookup"><span data-stu-id="22e1f-p157">Only used if `type` is set to `file`. If `true`, indicates that the attachment will be shown inline in the message body, and should not be displayed in the attachment list.</span></span>|
|`formData.attachments.itemId`|<span data-ttu-id="22e1f-955">Zeichenfolge</span><span class="sxs-lookup"><span data-stu-id="22e1f-955">String</span></span>||<span data-ttu-id="22e1f-p158">Wird nur verwendet, wenn `type` auf `item` gesetzt ist. Die EWS-Element-ID der Anlage. Dies ist eine Zeichenfolge bis zu 100 Zeichen.</span><span class="sxs-lookup"><span data-stu-id="22e1f-p158">Only used if `type` is set to `item`. The EWS item id of the attachment. This is a string up to 100 characters.</span></span>|
|`callback`|<span data-ttu-id="22e1f-959">Funktion</span><span class="sxs-lookup"><span data-stu-id="22e1f-959">function</span></span>|<span data-ttu-id="22e1f-960">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="22e1f-960">&lt;optional&gt;</span></span>|<span data-ttu-id="22e1f-961">Bei Abschluss der Methode wird die im Parameter `callback` übergebene Funktion mit einem einzigen Parameter aufgerufen: `asyncResult`. Dieser Parameter ist ein Objekt des Typs [AsyncResult](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="22e1f-961">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [AsyncResult](/javascript/api/office/office.asyncresult) object.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="22e1f-962">Anforderungen</span><span class="sxs-lookup"><span data-stu-id="22e1f-962">Requirements</span></span>

|<span data-ttu-id="22e1f-963">Anforderung</span><span class="sxs-lookup"><span data-stu-id="22e1f-963">Requirement</span></span>|<span data-ttu-id="22e1f-964">Wert</span><span class="sxs-lookup"><span data-stu-id="22e1f-964">Value</span></span>|
|---|---|
|[<span data-ttu-id="22e1f-965">Mindestversion des Postfachanforderungssatzes</span><span class="sxs-lookup"><span data-stu-id="22e1f-965">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="22e1f-966">1.0</span><span class="sxs-lookup"><span data-stu-id="22e1f-966">1.0</span></span>|
|[<span data-ttu-id="22e1f-967">Mindestberechtigungsstufe</span><span class="sxs-lookup"><span data-stu-id="22e1f-967">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="22e1f-968">ReadItem</span><span class="sxs-lookup"><span data-stu-id="22e1f-968">ReadItem</span></span>|
|[<span data-ttu-id="22e1f-969">Zutreffender Outlook-Modus</span><span class="sxs-lookup"><span data-stu-id="22e1f-969">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="22e1f-970">Lesen</span><span class="sxs-lookup"><span data-stu-id="22e1f-970">Read</span></span>|

##### <a name="examples"></a><span data-ttu-id="22e1f-971">Beispiele</span><span class="sxs-lookup"><span data-stu-id="22e1f-971">Examples</span></span>

<span data-ttu-id="22e1f-972">Im folgenden Code wird eine Zeichenfolge an die `displayReplyForm`-Funktion übergeben.</span><span class="sxs-lookup"><span data-stu-id="22e1f-972">The following code passes a string to the `displayReplyForm` function.</span></span>

```
Office.context.mailbox.item.displayReplyForm('hello there');
Office.context.mailbox.item.displayReplyForm('<b>hello there</b>');
```

<span data-ttu-id="22e1f-973">Antworten mit einem leeren Textkörper.</span><span class="sxs-lookup"><span data-stu-id="22e1f-973">Reply with an empty body.</span></span>

```
Office.context.mailbox.item.displayReplyForm({});
```

<span data-ttu-id="22e1f-974">Antworten nur einem Textkörper.</span><span class="sxs-lookup"><span data-stu-id="22e1f-974">Reply with just a body.</span></span>

```
Office.context.mailbox.item.displayReplyForm(
{
  'htmlBody' : 'hi'
});
```

<span data-ttu-id="22e1f-975">Antworten mit einem Textkörper und einer Dateianlage.</span><span class="sxs-lookup"><span data-stu-id="22e1f-975">Reply with a body and a file attachment.</span></span>

```
Office.context.mailbox.item.displayReplyForm(
{
  'htmlBody' : 'hi',
  'attachments' :
  [
    {
      'type' : Office.MailboxEnums.AttachmentType.File,
      'name' : 'squirrel.png',
      'url' : 'http://i.imgur.com/sRgTlGR.jpg'
    }
  ]
});
```

<span data-ttu-id="22e1f-976">Antworten mit einem Textkörper und einer Elementanlage.</span><span class="sxs-lookup"><span data-stu-id="22e1f-976">Reply with a body and an item attachment.</span></span>

```
Office.context.mailbox.item.displayReplyForm(
{
  'htmlBody' : 'hi',
  'attachments' :
  [
    {
      'type' : 'item',
      'name' : 'rand',
      'itemId' : Office.context.mailbox.item.itemId
    }
  ]
});
```

<span data-ttu-id="22e1f-977">Antworten Sie mit einem Textkörper, einer Dateianlage, einer Elementanlage und einem Callback.</span><span class="sxs-lookup"><span data-stu-id="22e1f-977">Reply with a body, file attachment, item attachment, and a callback.</span></span>

```
Office.context.mailbox.item.displayReplyForm(
{
  'htmlBody' : 'hi',
  'attachments' :
  [
    {
      'type' : Office.MailboxEnums.AttachmentType.File,
      'name' : 'squirrel.png',
      'url' : 'http://i.imgur.com/sRgTlGR.jpg'
    },
    {
      'type' : 'item',
      'name' : 'rand',
      'itemId' : Office.context.mailbox.item.itemId
    }
  ],
  'callback' : function(asyncResult)
  {
    console.log(asyncResult.value);
  }
});
```

#### <a name="getentities--entitiesjavascriptapioutlookofficeentities"></a><span data-ttu-id="22e1f-978">getEntities() → {[Entities](/javascript/api/outlook/office.entities)}</span><span class="sxs-lookup"><span data-stu-id="22e1f-978">getEntities() → {[Entities](/javascript/api/outlook/office.entities)}</span></span>

<span data-ttu-id="22e1f-979">Ruft das ausgewählte Element Textkörper gefundenen Entitäten ab.</span><span class="sxs-lookup"><span data-stu-id="22e1f-979">Gets the entities found in the selected item's body.</span></span>

> [!NOTE]
> <span data-ttu-id="22e1f-980">Diese Methode wird in Outlook für iOS oder Outlook für Android nicht unterstützt.</span><span class="sxs-lookup"><span data-stu-id="22e1f-980">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

##### <a name="requirements"></a><span data-ttu-id="22e1f-981">Anforderungen</span><span class="sxs-lookup"><span data-stu-id="22e1f-981">Requirements</span></span>

|<span data-ttu-id="22e1f-982">Anforderung</span><span class="sxs-lookup"><span data-stu-id="22e1f-982">Requirement</span></span>|<span data-ttu-id="22e1f-983">Wert</span><span class="sxs-lookup"><span data-stu-id="22e1f-983">Value</span></span>|
|---|---|
|[<span data-ttu-id="22e1f-984">Mindestversion des Postfachanforderungssatzes</span><span class="sxs-lookup"><span data-stu-id="22e1f-984">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="22e1f-985">1.0</span><span class="sxs-lookup"><span data-stu-id="22e1f-985">1.0</span></span>|
|[<span data-ttu-id="22e1f-986">Mindestberechtigungsstufe</span><span class="sxs-lookup"><span data-stu-id="22e1f-986">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="22e1f-987">ReadItem</span><span class="sxs-lookup"><span data-stu-id="22e1f-987">ReadItem</span></span>|
|[<span data-ttu-id="22e1f-988">Zutreffender Outlook-Modus</span><span class="sxs-lookup"><span data-stu-id="22e1f-988">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="22e1f-989">Lesen</span><span class="sxs-lookup"><span data-stu-id="22e1f-989">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="22e1f-990">Gibt zurück:</span><span class="sxs-lookup"><span data-stu-id="22e1f-990">Returns:</span></span>

<span data-ttu-id="22e1f-991">Typ: [Entitäten](/javascript/api/outlook/office.entities)</span><span class="sxs-lookup"><span data-stu-id="22e1f-991">Type: [Entities](/javascript/api/outlook/office.entities)</span></span>

##### <a name="example"></a><span data-ttu-id="22e1f-992">Beispiel</span><span class="sxs-lookup"><span data-stu-id="22e1f-992">Example</span></span>

<span data-ttu-id="22e1f-993">Das folgende Beispiel greift auf die Kontakte Entitäten im Textkörper des aktuellen Elements.</span><span class="sxs-lookup"><span data-stu-id="22e1f-993">The following example accesses the contacts entities in the current item's body.</span></span>

```
var contacts = Office.context.mailbox.item.getEntities().contacts;
```

#### <a name="getentitiesbytypeentitytype--nullable-arraystringcontactjavascriptapioutlookofficecontactmeetingsuggestionjavascriptapioutlookofficemeetingsuggestionphonenumberjavascriptapioutlookofficephonenumbertasksuggestionjavascriptapioutlookofficetasksuggestion"></a><span data-ttu-id="22e1f-994">getEntitiesByType(entityType) → (nullable) {Array.<(String|[Contact](/javascript/api/outlook/office.contact)|[MeetingSuggestion](/javascript/api/outlook/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook/office.tasksuggestion))>}</span><span class="sxs-lookup"><span data-stu-id="22e1f-994">getEntitiesByType(entityType) → (nullable) {Array.<(String|[Contact](/javascript/api/outlook/office.contact)|[MeetingSuggestion](/javascript/api/outlook/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook/office.tasksuggestion))>}</span></span>

<span data-ttu-id="22e1f-995">Ruft ein Array aller Entitäten des angegebenen Entitätstyps im Hauptteil des ausgewählten Elements gefunden.</span><span class="sxs-lookup"><span data-stu-id="22e1f-995">Gets an array of all the entities of the specified entity type found in the selected item's body.</span></span>

> [!NOTE]
> <span data-ttu-id="22e1f-996">Diese Methode wird in Outlook für iOS oder Outlook für Android nicht unterstützt.</span><span class="sxs-lookup"><span data-stu-id="22e1f-996">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

##### <a name="parameters"></a><span data-ttu-id="22e1f-997">Parameter:</span><span class="sxs-lookup"><span data-stu-id="22e1f-997">Parameters:</span></span>

|<span data-ttu-id="22e1f-998">Name</span><span class="sxs-lookup"><span data-stu-id="22e1f-998">Name</span></span>|<span data-ttu-id="22e1f-999">Typ</span><span class="sxs-lookup"><span data-stu-id="22e1f-999">Type</span></span>|<span data-ttu-id="22e1f-1000">Beschreibung</span><span class="sxs-lookup"><span data-stu-id="22e1f-1000">Description</span></span>|
|---|---|---|
|`entityType`|[<span data-ttu-id="22e1f-1001">Office.MailboxEnums.EntityType</span><span class="sxs-lookup"><span data-stu-id="22e1f-1001">Office.MailboxEnums.EntityType</span></span>](/javascript/api/outlook/office.mailboxenums.entitytype)|<span data-ttu-id="22e1f-1002">Einer der Werte der EntityType-Enumeration.</span><span class="sxs-lookup"><span data-stu-id="22e1f-1002">One of the EntityType enumeration values.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="22e1f-1003">Anforderungen</span><span class="sxs-lookup"><span data-stu-id="22e1f-1003">Requirements</span></span>

|<span data-ttu-id="22e1f-1004">Anforderung</span><span class="sxs-lookup"><span data-stu-id="22e1f-1004">Requirement</span></span>|<span data-ttu-id="22e1f-1005">Wert</span><span class="sxs-lookup"><span data-stu-id="22e1f-1005">Value</span></span>|
|---|---|
|[<span data-ttu-id="22e1f-1006">Mindestversion des Postfachanforderungssatzes</span><span class="sxs-lookup"><span data-stu-id="22e1f-1006">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="22e1f-1007">1.0</span><span class="sxs-lookup"><span data-stu-id="22e1f-1007">1.0</span></span>|
|[<span data-ttu-id="22e1f-1008">Mindestberechtigungsstufe</span><span class="sxs-lookup"><span data-stu-id="22e1f-1008">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="22e1f-1009">Eingeschränkt</span><span class="sxs-lookup"><span data-stu-id="22e1f-1009">Restricted</span></span>|
|[<span data-ttu-id="22e1f-1010">Zutreffender Outlook-Modus</span><span class="sxs-lookup"><span data-stu-id="22e1f-1010">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="22e1f-1011">Lesen</span><span class="sxs-lookup"><span data-stu-id="22e1f-1011">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="22e1f-1012">Gibt zurück:</span><span class="sxs-lookup"><span data-stu-id="22e1f-1012">Returns:</span></span>

<span data-ttu-id="22e1f-1013">Wenn der in `entityType` übergebene Wert kein gültiges Element der `EntityType`-Enumeration ist, gibt die Methode null zurück.</span><span class="sxs-lookup"><span data-stu-id="22e1f-1013">If the value passed in `entityType` is not a valid member of the `EntityType` enumeration, the method returns null.</span></span> <span data-ttu-id="22e1f-1014">Wenn keine Entitäten des angegebenen Typs im Textkörper des Elements vorhanden sind, gibt die Methode ein leeres Array zurück.</span><span class="sxs-lookup"><span data-stu-id="22e1f-1014">If no entities of the specified type are present in the item's body, the method returns an empty array.</span></span> <span data-ttu-id="22e1f-1015">Andernfalls hängt der Typ der Objekte im zurückgegebenen Array vom Typ der Entität ab, die im `entityType`-Parameter angefordert wurde.</span><span class="sxs-lookup"><span data-stu-id="22e1f-1015">Otherwise, the type of the objects in the returned array depends on the type of entity requested in the `entityType` parameter.</span></span>

<span data-ttu-id="22e1f-1016">Während Sie die minimale Berechtigungsstufe für diese Methode **Restricted** ist, erfordern einige Entitätstypen **ReadItem** für den Zugriff, wie in der folgenden Tabelle angegeben wird.</span><span class="sxs-lookup"><span data-stu-id="22e1f-1016">While the minimum permission level to use this method is **Restricted**, some entity types require **ReadItem** to access, as specified in the following table.</span></span>

|<span data-ttu-id="22e1f-1017">Wert von `entityType`</span><span class="sxs-lookup"><span data-stu-id="22e1f-1017">Value of `entityType`</span></span>|<span data-ttu-id="22e1f-1018">Typ der Objekte im zurückgegebenen Array</span><span class="sxs-lookup"><span data-stu-id="22e1f-1018">Type of objects in returned array</span></span>|<span data-ttu-id="22e1f-1019">Erforderliche Berechtigungsstufe</span><span class="sxs-lookup"><span data-stu-id="22e1f-1019">Required Permission Level</span></span>|
|---|---|---|
|`Address`|<span data-ttu-id="22e1f-1020">Zeichenfolge</span><span class="sxs-lookup"><span data-stu-id="22e1f-1020">String</span></span>|<span data-ttu-id="22e1f-1021">**Restricted**</span><span class="sxs-lookup"><span data-stu-id="22e1f-1021">**Restricted**</span></span>|
|`Contact`|<span data-ttu-id="22e1f-1022">Contact</span><span class="sxs-lookup"><span data-stu-id="22e1f-1022">Contact</span></span>|<span data-ttu-id="22e1f-1023">**ReadItem**</span><span class="sxs-lookup"><span data-stu-id="22e1f-1023">**ReadItem**</span></span>|
|`EmailAddress`|<span data-ttu-id="22e1f-1024">Zeichenfolge</span><span class="sxs-lookup"><span data-stu-id="22e1f-1024">String</span></span>|<span data-ttu-id="22e1f-1025">**ReadItem**</span><span class="sxs-lookup"><span data-stu-id="22e1f-1025">**ReadItem**</span></span>|
|`MeetingSuggestion`|<span data-ttu-id="22e1f-1026">MeetingSuggestion</span><span class="sxs-lookup"><span data-stu-id="22e1f-1026">MeetingSuggestion</span></span>|<span data-ttu-id="22e1f-1027">**ReadItem**</span><span class="sxs-lookup"><span data-stu-id="22e1f-1027">**ReadItem**</span></span>|
|`PhoneNumber`|<span data-ttu-id="22e1f-1028">PhoneNumber</span><span class="sxs-lookup"><span data-stu-id="22e1f-1028">PhoneNumber</span></span>|<span data-ttu-id="22e1f-1029">**Restricted**</span><span class="sxs-lookup"><span data-stu-id="22e1f-1029">**Restricted**</span></span>|
|`TaskSuggestion`|<span data-ttu-id="22e1f-1030">TaskSuggestion</span><span class="sxs-lookup"><span data-stu-id="22e1f-1030">TaskSuggestion</span></span>|<span data-ttu-id="22e1f-1031">**ReadItem**</span><span class="sxs-lookup"><span data-stu-id="22e1f-1031">**ReadItem**</span></span>|
|`URL`|<span data-ttu-id="22e1f-1032">Zeichenfolge</span><span class="sxs-lookup"><span data-stu-id="22e1f-1032">String</span></span>|<span data-ttu-id="22e1f-1033">**Restricted**</span><span class="sxs-lookup"><span data-stu-id="22e1f-1033">**Restricted**</span></span>|

<span data-ttu-id="22e1f-1034">Typ: Array.<(String|[Contact](/javascript/api/outlook/office.contact)|[MeetingSuggestion](/javascript/api/outlook/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook/office.tasksuggestion))></span><span class="sxs-lookup"><span data-stu-id="22e1f-1034">Type: Array.<(String|[Contact](/javascript/api/outlook/office.contact)|[MeetingSuggestion](/javascript/api/outlook/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook/office.tasksuggestion))></span></span>

##### <a name="example"></a><span data-ttu-id="22e1f-1035">Beispiel</span><span class="sxs-lookup"><span data-stu-id="22e1f-1035">Example</span></span>

<span data-ttu-id="22e1f-1036">Im folgenden Beispiel wird veranschaulicht, wie auf ein Array von Zeichenfolgen, die Postadressen im Textkörper des aktuellen Elements darstellen.</span><span class="sxs-lookup"><span data-stu-id="22e1f-1036">The following example shows how to access an array of strings that represent postal addresses in the current item's body.</span></span>

```
// The initialize function is required for all apps.
Office.initialize = function () {
  // Checks for the DOM to load using the jQuery ready function.
  $(document).ready(function () {
    // After the DOM is loaded, app-specific code can run.
    var item = Office.context.mailbox.item;
    // Get an array of strings that represent postal addresses in the current item's body.
    var addresses = item.getEntitiesByType(Office.MailboxEnums.EntityType.Address);
    // Continue processing the array of addresses.
  });
}
```

#### <a name="getfilteredentitiesbynamename--nullable-arraystringcontactjavascriptapioutlookofficecontactmeetingsuggestionjavascriptapioutlookofficemeetingsuggestionphonenumberjavascriptapioutlookofficephonenumbertasksuggestionjavascriptapioutlookofficetasksuggestion"></a><span data-ttu-id="22e1f-1037">getFilteredEntitiesByName(name) → (nullable) {Array.<(String|[Contact](/javascript/api/outlook/office.contact)|[MeetingSuggestion](/javascript/api/outlook/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook/office.tasksuggestion))>}</span><span class="sxs-lookup"><span data-stu-id="22e1f-1037">getFilteredEntitiesByName(name) → (nullable) {Array.<(String|[Contact](/javascript/api/outlook/office.contact)|[MeetingSuggestion](/javascript/api/outlook/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook/office.tasksuggestion))>}</span></span>

<span data-ttu-id="22e1f-1038">Gibt bekannte Entitäten im ausgewählten Element zurück, die den in der XML-Manifestdatei definierten benannten Filter übergeben.</span><span class="sxs-lookup"><span data-stu-id="22e1f-1038">Returns well-known entities in the selected item that pass the named filter defined in the manifest XML file.</span></span>

> [!NOTE]
> <span data-ttu-id="22e1f-1039">Diese Methode wird in Outlook für iOS oder Outlook für Android nicht unterstützt.</span><span class="sxs-lookup"><span data-stu-id="22e1f-1039">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="22e1f-1040">Die `getFilteredEntitiesByName`-Methode gibt die Entitäten zurück, die dem im [ItemHasKnownEntity](/javascript/office/manifest/rule#itemhasknownentity-rule)-Regelelement der XML-Manifestdatei definierten regulären Ausdruck mit dem angegebenen `FilterName`-Elementwert entsprechen.</span><span class="sxs-lookup"><span data-stu-id="22e1f-1040">The `getFilteredEntitiesByName` method returns the entities that match the regular expression defined in the [ItemHasKnownEntity](/javascript/office/manifest/rule#itemhasknownentity-rule) rule element in the manifest XML file with the specified `FilterName` element value.</span></span>

##### <a name="parameters"></a><span data-ttu-id="22e1f-1041">Parameter:</span><span class="sxs-lookup"><span data-stu-id="22e1f-1041">Parameters:</span></span>

|<span data-ttu-id="22e1f-1042">Name</span><span class="sxs-lookup"><span data-stu-id="22e1f-1042">Name</span></span>|<span data-ttu-id="22e1f-1043">Typ</span><span class="sxs-lookup"><span data-stu-id="22e1f-1043">Type</span></span>|<span data-ttu-id="22e1f-1044">Beschreibung</span><span class="sxs-lookup"><span data-stu-id="22e1f-1044">Description</span></span>|
|---|---|---|
|`name`|<span data-ttu-id="22e1f-1045">Zeichenfolge</span><span class="sxs-lookup"><span data-stu-id="22e1f-1045">String</span></span>|<span data-ttu-id="22e1f-1046">Der Name des `ItemHasKnownEntity`-Regelelements, das den entsprechenden Filter definiert.</span><span class="sxs-lookup"><span data-stu-id="22e1f-1046">The name of the `ItemHasKnownEntity` rule element that defines the filter to match.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="22e1f-1047">Anforderungen</span><span class="sxs-lookup"><span data-stu-id="22e1f-1047">Requirements</span></span>

|<span data-ttu-id="22e1f-1048">Anforderung</span><span class="sxs-lookup"><span data-stu-id="22e1f-1048">Requirement</span></span>|<span data-ttu-id="22e1f-1049">Wert</span><span class="sxs-lookup"><span data-stu-id="22e1f-1049">Value</span></span>|
|---|---|
|[<span data-ttu-id="22e1f-1050">Mindestversion des Postfachanforderungssatzes</span><span class="sxs-lookup"><span data-stu-id="22e1f-1050">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="22e1f-1051">1.0</span><span class="sxs-lookup"><span data-stu-id="22e1f-1051">1.0</span></span>|
|[<span data-ttu-id="22e1f-1052">Mindestberechtigungsstufe</span><span class="sxs-lookup"><span data-stu-id="22e1f-1052">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="22e1f-1053">ReadItem</span><span class="sxs-lookup"><span data-stu-id="22e1f-1053">ReadItem</span></span>|
|[<span data-ttu-id="22e1f-1054">Zutreffender Outlook-Modus</span><span class="sxs-lookup"><span data-stu-id="22e1f-1054">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="22e1f-1055">Lesen</span><span class="sxs-lookup"><span data-stu-id="22e1f-1055">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="22e1f-1056">Gibt zurück:</span><span class="sxs-lookup"><span data-stu-id="22e1f-1056">Returns:</span></span>

<span data-ttu-id="22e1f-p160">Befindet sich kein `ItemHasKnownEntity`-Element im Manifest mit einem `FilterName`-Elementwert, der dem `name`-Parameter entspricht, gibt die Methode `null` zurück. Wenn der `name`-Parameter einem `ItemHasKnownEntity`-Element im Manifest entspricht, es aber keine entsprechenden Entitäten im aktuellen Element gibt, gibt die Methode ein leeres Array zurück.</span><span class="sxs-lookup"><span data-stu-id="22e1f-p160">If there is no `ItemHasKnownEntity` element in the manifest with a `FilterName` element value that matches the `name` parameter, the method returns `null`. If the `name` parameter does match an `ItemHasKnownEntity` element in the manifest, but there are no entities in the current item that match, the method return an empty array.</span></span>

<span data-ttu-id="22e1f-1059">Typ: Array.<(String|[Contact](/javascript/api/outlook/office.contact)|[MeetingSuggestion](/javascript/api/outlook/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook/office.tasksuggestion))></span><span class="sxs-lookup"><span data-stu-id="22e1f-1059">Type: Array.<(String|[Contact](/javascript/api/outlook/office.contact)|[MeetingSuggestion](/javascript/api/outlook/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook/office.tasksuggestion))></span></span>

#### <a name="getinitializationcontextasyncoptions-callback"></a><span data-ttu-id="22e1f-1060">getInitializationContextAsync([options], [callback])</span><span class="sxs-lookup"><span data-stu-id="22e1f-1060">getInitializationContextAsync([options], [callback])</span></span>

<span data-ttu-id="22e1f-1061">Ruft Initialisierungsdaten ab, die übergeben werden, wenn das Add-In [durch eine Nachricht mit Aktionen aktiviert](https://docs.microsoft.com/outlook/actionable-messages/invoke-add-in-from-actionable-message) wird.</span><span class="sxs-lookup"><span data-stu-id="22e1f-1061">Gets initialization data passed when the add-in is [activated by an actionable message](https://docs.microsoft.com/outlook/actionable-messages/invoke-add-in-from-actionable-message).</span></span>

> [!NOTE]
> <span data-ttu-id="22e1f-1062">Diese Methode wird nur vom Outlook 2016 für Windows (größer als 16.0.8413.1000 Klick-und-Los-Versionen) und Outlook im Web für Office 365 unterstützt.</span><span class="sxs-lookup"><span data-stu-id="22e1f-1062">This method is only supported by Outlook 2016 for Windows (Click-to-Run versions greater than 16.0.8413.1000) and Outlook on the web for Office 365.</span></span>

##### <a name="parameters"></a><span data-ttu-id="22e1f-1063">Parameter:</span><span class="sxs-lookup"><span data-stu-id="22e1f-1063">Parameters:</span></span>
|<span data-ttu-id="22e1f-1064">Name</span><span class="sxs-lookup"><span data-stu-id="22e1f-1064">Name</span></span>|<span data-ttu-id="22e1f-1065">Typ</span><span class="sxs-lookup"><span data-stu-id="22e1f-1065">Type</span></span>|<span data-ttu-id="22e1f-1066">Attribute</span><span class="sxs-lookup"><span data-stu-id="22e1f-1066">Attributes</span></span>|<span data-ttu-id="22e1f-1067">Beschreibung</span><span class="sxs-lookup"><span data-stu-id="22e1f-1067">Description</span></span>|
|---|---|---|---|
|`options`|<span data-ttu-id="22e1f-1068">Objekt</span><span class="sxs-lookup"><span data-stu-id="22e1f-1068">Object</span></span>|<span data-ttu-id="22e1f-1069">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="22e1f-1069">&lt;optional&gt;</span></span>|<span data-ttu-id="22e1f-1070">Ein Objektliteral, das mindestens eine der folgenden Eigenschaften enthält.</span><span class="sxs-lookup"><span data-stu-id="22e1f-1070">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`|<span data-ttu-id="22e1f-1071">Object</span><span class="sxs-lookup"><span data-stu-id="22e1f-1071">Object</span></span>|<span data-ttu-id="22e1f-1072">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="22e1f-1072">&lt;optional&gt;</span></span>|<span data-ttu-id="22e1f-1073">Entwickler können ein Objekt bereitstellen, auf das sie in der Callbackmethode zugreifen möchten.</span><span class="sxs-lookup"><span data-stu-id="22e1f-1073">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`|<span data-ttu-id="22e1f-1074">Funktion</span><span class="sxs-lookup"><span data-stu-id="22e1f-1074">function</span></span>|<span data-ttu-id="22e1f-1075">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="22e1f-1075">&lt;optional&gt;</span></span>|<span data-ttu-id="22e1f-1076">Wenn die Methode abgeschlossen wird, wird die Funktion , die im `callback`-Parameter übergeben wird, mit einem einzelnen Parameter (`asyncResult`) aufgerufen, der ein [`AsyncResult`](/javascript/api/office/office.asyncresult)-Objekt darstellt.</span><span class="sxs-lookup"><span data-stu-id="22e1f-1076">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span> <br/><span data-ttu-id="22e1f-1077">Initialisierungsdaten erfolgt auf Erfolg, in der `asyncResult.value` -Eigenschaft, wie eine Zeichenfolge.</span><span class="sxs-lookup"><span data-stu-id="22e1f-1077">On success, the initialization data is provided in the `asyncResult.value` property as a string.</span></span><br/><span data-ttu-id="22e1f-1078">Wenn es keinen Initialisierungskontext gibt, enthält das `asyncResult`-Objekt ein `Error`-Objekt, dessen `code`-Eigenschaft auf `9020` und dessen `name`-Eigenschaft auf `GenericResponseError` festgelegt ist.</span><span class="sxs-lookup"><span data-stu-id="22e1f-1078">If there is no initialization context, the `asyncResult` object will contain an `Error` object with its `code` property set to `9020` and its `name` property set to `GenericResponseError`.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="22e1f-1079">Anforderungen</span><span class="sxs-lookup"><span data-stu-id="22e1f-1079">Requirements</span></span>

|<span data-ttu-id="22e1f-1080">Anforderung</span><span class="sxs-lookup"><span data-stu-id="22e1f-1080">Requirement</span></span>|<span data-ttu-id="22e1f-1081">Wert</span><span class="sxs-lookup"><span data-stu-id="22e1f-1081">Value</span></span>|
|---|---|
|[<span data-ttu-id="22e1f-1082">Mindestversion des Postfachanforderungssatzes</span><span class="sxs-lookup"><span data-stu-id="22e1f-1082">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="22e1f-1083">Vorschau</span><span class="sxs-lookup"><span data-stu-id="22e1f-1083">Preview</span></span>|
|[<span data-ttu-id="22e1f-1084">Mindestberechtigungsstufe</span><span class="sxs-lookup"><span data-stu-id="22e1f-1084">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="22e1f-1085">ReadItem</span><span class="sxs-lookup"><span data-stu-id="22e1f-1085">ReadItem</span></span>|
|[<span data-ttu-id="22e1f-1086">Zutreffender Outlook-Modus</span><span class="sxs-lookup"><span data-stu-id="22e1f-1086">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="22e1f-1087">Lesen</span><span class="sxs-lookup"><span data-stu-id="22e1f-1087">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="22e1f-1088">Beispiel</span><span class="sxs-lookup"><span data-stu-id="22e1f-1088">Example</span></span>

```
// Get the initialization context (if present)
Office.context.mailbox.item.getInitializationContextAsync(
  function(asyncResult) {
    if (asyncResult.status == Office.AsyncResultStatus.Succeeded) {
      if (asyncResult.value != null && asyncResult.value.length > 0) {
        // The value is a string, parse to an object
        var context = JSON.parse(asyncResult.value);
        // Do something with context
      } else {
        // Empty context, treat as no context
      }
    } else {
      if (asyncResult.error.code == 9020) {
        // GenericResponseError returned when there is
        // no context
        // Treat as no context
      } else {
        // Handle the error
      }
    }
  }
);
```

#### <a name="getregexmatches--object"></a><span data-ttu-id="22e1f-1089">getRegExMatches() → {Object}</span><span class="sxs-lookup"><span data-stu-id="22e1f-1089">getRegExMatches() → {Object}</span></span>

<span data-ttu-id="22e1f-1090">Gibt Zeichenfolgenwerte im ausgewählten Element zurück, die den in der XML-Manifestdatei definierten regulären Ausdrücken entsprechen.</span><span class="sxs-lookup"><span data-stu-id="22e1f-1090">Returns string values in the selected item that match the regular expressions defined in the manifest XML file.</span></span>

> [!NOTE]
> <span data-ttu-id="22e1f-1091">Diese Methode wird in Outlook für iOS oder Outlook für Android nicht unterstützt.</span><span class="sxs-lookup"><span data-stu-id="22e1f-1091">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="22e1f-p161">Die `getRegExMatches`-Methode gibt die Zeichenfolgen zurück, die dem im `ItemHasRegularExpressionMatch`- oder `ItemHasKnownEntity`-Regelelement der XML-Manifestdatei definierten regulären Ausdruck entsprechen. Bei einer `ItemHasRegularExpressionMatch`-Regel muss eine entsprechende Zeichenfolge in der Eigenschaft des Elements vorhanden sein, das von dieser Regel angegeben wird. Der einfache `PropertyName`-Typ definiert die unterstützten Eigenschaften.</span><span class="sxs-lookup"><span data-stu-id="22e1f-p161">The `getRegExMatches` method returns the strings that match the regular expression defined in each `ItemHasRegularExpressionMatch` or `ItemHasKnownEntity` rule element in the manifest XML file. For an `ItemHasRegularExpressionMatch` rule, a matching string has to occur in the property of the item that is specified by that rule. The `PropertyName` simple type defines the supported properties.</span></span>

<span data-ttu-id="22e1f-1095">Nehmen Sie z. B. an, dass ein Add-In-Manifest das folgende `Rule`-Element aufweist:</span><span class="sxs-lookup"><span data-stu-id="22e1f-1095">For example, consider an add-in manifest has the following `Rule` element:</span></span>

```
<Rule xsi:type="RuleCollection" Mode="And">
  <Rule xsi:type="ItemIs" FormType="Read" ItemType="Message" />
  <Rule xsi:type="RuleCollection" Mode="Or">
    <Rule xsi:type="ItemHasRegularExpressionMatch" RegExName="fruits" RegExValue="apple|banana|coconut" PropertyName="BodyAsPlaintext" IgnoreCase="true" />
    <Rule xsi:type="ItemHasRegularExpressionMatch" RegExName="veggies" RegExValue="tomato|onion|spinach|broccoli" PropertyName="BodyAsPlaintext" IgnoreCase="true" />
  </Rule>
</Rule>
```

<span data-ttu-id="22e1f-1096">Das von `getRegExMatches` zurückgegebene Objekt hätte zwei Eigenschaften: `fruits` und `veggies`.</span><span class="sxs-lookup"><span data-stu-id="22e1f-1096">The object returned from `getRegExMatches` would have two properties: `fruits` and `veggies`.</span></span>

```
{
  'fruits': ['apple','banana','Banana','coconut'],
  'veggies': ['tomato','onion','spinach','broccoli']
}
```

<span data-ttu-id="22e1f-p162">Wenn Sie eine `ItemHasRegularExpressionMatch`-Regel für die Textkörpereigenschaft eines Elements festlegen, sollte der reguläre Ausdruck den Textkörper weiter filtern und nicht versuchen, den gesamten Textkörper des Elements zurückzugeben. Wenn der gesamte Textkörper eines Elements mit einem regulären Ausdruck wie `.*` abgerufen wird, werden nicht immer die gewünschten Ergebnisse erzielt. Verwenden Sie stattdessen die [`Body.getAsync`](/javascript/api/outlook/office.body#getasync-coerciontype--options--callback-)-Methode, um den gesamten Textkörper abzurufen.</span><span class="sxs-lookup"><span data-stu-id="22e1f-p162">If you specify an `ItemHasRegularExpressionMatch` rule on the body property of an item, the regular expression should further filter the body and should not attempt to return the entire body of the item. Using a regular expression such as `.*` to obtain the entire body of an item does not always return the expected results. Instead, use the [`Body.getAsync`](/javascript/api/outlook/office.body#getasync-coerciontype--options--callback-) method to retrieve the entire body.</span></span>

##### <a name="requirements"></a><span data-ttu-id="22e1f-1100">Anforderungen</span><span class="sxs-lookup"><span data-stu-id="22e1f-1100">Requirements</span></span>

|<span data-ttu-id="22e1f-1101">Anforderung</span><span class="sxs-lookup"><span data-stu-id="22e1f-1101">Requirement</span></span>|<span data-ttu-id="22e1f-1102">Wert</span><span class="sxs-lookup"><span data-stu-id="22e1f-1102">Value</span></span>|
|---|---|
|[<span data-ttu-id="22e1f-1103">Mindestversion des Postfachanforderungssatzes</span><span class="sxs-lookup"><span data-stu-id="22e1f-1103">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="22e1f-1104">1.0</span><span class="sxs-lookup"><span data-stu-id="22e1f-1104">1.0</span></span>|
|[<span data-ttu-id="22e1f-1105">Mindestberechtigungsstufe</span><span class="sxs-lookup"><span data-stu-id="22e1f-1105">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="22e1f-1106">ReadItem</span><span class="sxs-lookup"><span data-stu-id="22e1f-1106">ReadItem</span></span>|
|[<span data-ttu-id="22e1f-1107">Zutreffender Outlook-Modus</span><span class="sxs-lookup"><span data-stu-id="22e1f-1107">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="22e1f-1108">Lesen</span><span class="sxs-lookup"><span data-stu-id="22e1f-1108">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="22e1f-1109">Gibt zurück:</span><span class="sxs-lookup"><span data-stu-id="22e1f-1109">Returns:</span></span>

<span data-ttu-id="22e1f-p163">Ein Objekt mit Arrays aus Zeichenfolgen, die den in der XML-Manifestdatei definierten regulären Ausdrücken entsprechen. Der Name der einzelnen Arrays ist gleich dem entsprechenden Wert des `RegExName`-Attributs der entsprechenden `ItemHasRegularExpressionMatch`-Regel oder des `FilterName`-Attributs der entsprechenden `ItemHasKnownEntity`-Regel.</span><span class="sxs-lookup"><span data-stu-id="22e1f-p163">An object that contains arrays of strings that match the regular expressions defined in the manifest XML file. The name of each array is equal to the corresponding value of the `RegExName` attribute of the matching `ItemHasRegularExpressionMatch` rule or the `FilterName` attribute of the matching `ItemHasKnownEntity` rule.</span></span>

<dl class="param-type"><span data-ttu-id="22e1f-1112">

<dt>
Typ</dt>


</span><span class="sxs-lookup"><span data-stu-id="22e1f-1112">

<dt>Type</dt>

</span></span><dd><span data-ttu-id="22e1f-1113">Object</span><span class="sxs-lookup"><span data-stu-id="22e1f-1113">Object</span></span></dd>

</dl>

##### <a name="example"></a><span data-ttu-id="22e1f-1114">Beispiel</span><span class="sxs-lookup"><span data-stu-id="22e1f-1114">Example</span></span>

<span data-ttu-id="22e1f-1115">Das folgende Beispiel zeigt, wie Sie auf das Array von Übereinstimmungen für die regulären Ausdrucksregelelemente `fruits` und `veggies` zugreifen, die im Manifest angegeben sind.</span><span class="sxs-lookup"><span data-stu-id="22e1f-1115">The following example shows how to access the array of matches for the regular expression rule elements `fruits` and `veggies`, which are specified in the manifest.</span></span>

```
var allMatches = Office.context.mailbox.item.getRegExMatches();
var fruits = allMatches.fruits;
var veges = allMatches.veggies;
```

#### <a name="getregexmatchesbynamename--nullable-array-string-"></a><span data-ttu-id="22e1f-1116">getRegExMatchesByName(name) → (Nullwerte zulassen) {Array. < Zeichenfolge >}</span><span class="sxs-lookup"><span data-stu-id="22e1f-1116">getRegExMatchesByName(name) → (nullable) {Array.< String >}</span></span>

<span data-ttu-id="22e1f-1117">Gibt Zeichenfolgenwerte im ausgewählten Element zurück, die dem in der XML-Manifestdatei definierten benannten regulären Ausdruck entsprechen.</span><span class="sxs-lookup"><span data-stu-id="22e1f-1117">Returns string values in the selected item that match the named regular expression defined in the manifest XML file.</span></span>

> [!NOTE]
> <span data-ttu-id="22e1f-1118">Diese Methode wird in Outlook für iOS oder Outlook für Android nicht unterstützt.</span><span class="sxs-lookup"><span data-stu-id="22e1f-1118">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="22e1f-1119">Die `getRegExMatchesByName`-Methode gibt die Zeichenfolgen zurück, die dem im `ItemHasRegularExpressionMatch`-Regelelement der XML-Manifestdatei definierten regulären Ausdruck mit dem angegebenen `RegExName`-Elementwert entsprechen.</span><span class="sxs-lookup"><span data-stu-id="22e1f-1119">The `getRegExMatchesByName` method returns the strings that match the regular expression defined in the `ItemHasRegularExpressionMatch` rule element in the manifest XML file with the specified `RegExName` element value.</span></span>

<span data-ttu-id="22e1f-p164">Wenn Sie eine `ItemHasRegularExpressionMatch`-Regel für die Textkörpereigenschaft eines Elements festlegen, sollte der reguläre Ausdruck den Textkörper weiter filtern und nicht versuchen, den gesamten Textkörper des Elements zurückzugeben. Wenn der gesamte Textkörper eines Elements mit einem regulären Ausdruck wie `.*` abgerufen wird, werden nicht immer die gewünschten Ergebnisse erzielt.</span><span class="sxs-lookup"><span data-stu-id="22e1f-p164">If you specify an `ItemHasRegularExpressionMatch` rule on the body property of an item, the regular expression should further filter the body and should not attempt to return the entire body of the item. Using a regular expression such as `.*` to obtain the entire body of an item does not always return the expected results.</span></span>

##### <a name="parameters"></a><span data-ttu-id="22e1f-1122">Parameter:</span><span class="sxs-lookup"><span data-stu-id="22e1f-1122">Parameters:</span></span>

|<span data-ttu-id="22e1f-1123">Name</span><span class="sxs-lookup"><span data-stu-id="22e1f-1123">Name</span></span>|<span data-ttu-id="22e1f-1124">Typ</span><span class="sxs-lookup"><span data-stu-id="22e1f-1124">Type</span></span>|<span data-ttu-id="22e1f-1125">Beschreibung</span><span class="sxs-lookup"><span data-stu-id="22e1f-1125">Description</span></span>|
|---|---|---|
|`name`|<span data-ttu-id="22e1f-1126">Zeichenfolge</span><span class="sxs-lookup"><span data-stu-id="22e1f-1126">String</span></span>|<span data-ttu-id="22e1f-1127">Der Name des `ItemHasRegularExpressionMatch`-Regelelements, das den entsprechenden Filter definiert.</span><span class="sxs-lookup"><span data-stu-id="22e1f-1127">The name of the `ItemHasRegularExpressionMatch` rule element that defines the filter to match.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="22e1f-1128">Anforderungen</span><span class="sxs-lookup"><span data-stu-id="22e1f-1128">Requirements</span></span>

|<span data-ttu-id="22e1f-1129">Anforderung</span><span class="sxs-lookup"><span data-stu-id="22e1f-1129">Requirement</span></span>|<span data-ttu-id="22e1f-1130">Wert</span><span class="sxs-lookup"><span data-stu-id="22e1f-1130">Value</span></span>|
|---|---|
|[<span data-ttu-id="22e1f-1131">Mindestversion des Postfachanforderungssatzes</span><span class="sxs-lookup"><span data-stu-id="22e1f-1131">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="22e1f-1132">1.0</span><span class="sxs-lookup"><span data-stu-id="22e1f-1132">1.0</span></span>|
|[<span data-ttu-id="22e1f-1133">Mindestberechtigungsstufe</span><span class="sxs-lookup"><span data-stu-id="22e1f-1133">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="22e1f-1134">ReadItem</span><span class="sxs-lookup"><span data-stu-id="22e1f-1134">ReadItem</span></span>|
|[<span data-ttu-id="22e1f-1135">Zutreffender Outlook-Modus</span><span class="sxs-lookup"><span data-stu-id="22e1f-1135">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="22e1f-1136">Lesen</span><span class="sxs-lookup"><span data-stu-id="22e1f-1136">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="22e1f-1137">Gibt zurück:</span><span class="sxs-lookup"><span data-stu-id="22e1f-1137">Returns:</span></span>

<span data-ttu-id="22e1f-1138">Ein Array mit den Zeichenfolgen, die dem in der XML-Manifestdatei definierten regulären Ausdruck entsprechen.</span><span class="sxs-lookup"><span data-stu-id="22e1f-1138">An array that contains the strings that match the regular expression defined in the manifest XML file.</span></span>

<dl class="param-type"><span data-ttu-id="22e1f-1139">

<dt>Typ</dt>

</span><span class="sxs-lookup"><span data-stu-id="22e1f-1139">

<dt>Type</dt>

</span></span><dd><span data-ttu-id="22e1f-1140">Array. < Zeichenfolge ></span><span class="sxs-lookup"><span data-stu-id="22e1f-1140">Array.< String ></span></span></dd>

</dl>

##### <a name="example"></a><span data-ttu-id="22e1f-1141">Beispiel</span><span class="sxs-lookup"><span data-stu-id="22e1f-1141">Example</span></span>

```
var fruits = Office.context.mailbox.item.getRegExMatchesByName("fruits");
var veggies = Office.context.mailbox.item.getRegExMatchesByName("veggies");
```

####  <a name="getselecteddataasynccoerciontype-options-callback--string"></a><span data-ttu-id="22e1f-1142">getSelectedDataAsync(coercionType, [options], callback) → {String}</span><span class="sxs-lookup"><span data-stu-id="22e1f-1142">getSelectedDataAsync(coercionType, [options], callback) → {String}</span></span>

<span data-ttu-id="22e1f-1143">Gibt asynchron ausgewählte Daten aus dem Betreff oder Textkörper einer Nachricht zurück.</span><span class="sxs-lookup"><span data-stu-id="22e1f-1143">Asynchronously returns selected data from the subject or body of a message.</span></span>

<span data-ttu-id="22e1f-p165">Wenn keine Auswahl vorhanden ist, aber der Cursor sich im Nachrichtentext oder Betreff befindet, gibt die Methode für die ausgewählten Daten NULL zurück. Wenn ein anderes Feld als der Textkörper oder Betreff ausgewählt ist, gibt die Methode den `InvalidSelection`-Fehler zurück.</span><span class="sxs-lookup"><span data-stu-id="22e1f-p165">If there is no selection but the cursor is in the body or subject, the method returns null for the selected data. If a field other than the body or subject is selected, the method returns the `InvalidSelection` error.</span></span>

##### <a name="parameters"></a><span data-ttu-id="22e1f-1146">Parameter:</span><span class="sxs-lookup"><span data-stu-id="22e1f-1146">Parameters:</span></span>

|<span data-ttu-id="22e1f-1147">Name</span><span class="sxs-lookup"><span data-stu-id="22e1f-1147">Name</span></span>|<span data-ttu-id="22e1f-1148">Typ</span><span class="sxs-lookup"><span data-stu-id="22e1f-1148">Type</span></span>|<span data-ttu-id="22e1f-1149">Attribute</span><span class="sxs-lookup"><span data-stu-id="22e1f-1149">Attributes</span></span>|<span data-ttu-id="22e1f-1150">Beschreibung</span><span class="sxs-lookup"><span data-stu-id="22e1f-1150">Description</span></span>|
|---|---|---|---|
|`coercionType`|[<span data-ttu-id="22e1f-1151">Office.CoercionType</span><span class="sxs-lookup"><span data-stu-id="22e1f-1151">Office.CoercionType</span></span>](office.md#coerciontype-string)||<span data-ttu-id="22e1f-p166">Fordert ein Format für die Daten an. Wenn es sich um Texthandelt, gibt die Methode den unformatierten Text als Zeichenfolge zurück und entfernt eventuell vorhandene HTML-Tags. Wenn es sich um HTML handelt, gibt die Methode den ausgewählten Text zurück, entweder als unformatierten Text oder als HTML.</span><span class="sxs-lookup"><span data-stu-id="22e1f-p166">Requests a format for the data. If Text, the method returns the plain text as a string , removing any HTML tags present. If HTML, the method returns the selected text, whether it is plaintext or HTML.</span></span>|
|`options`|<span data-ttu-id="22e1f-1155">Object</span><span class="sxs-lookup"><span data-stu-id="22e1f-1155">Object</span></span>|<span data-ttu-id="22e1f-1156">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="22e1f-1156">&lt;optional&gt;</span></span>|<span data-ttu-id="22e1f-1157">Ein Objektliteral, das mindestens eine der folgenden Eigenschaften enthält.</span><span class="sxs-lookup"><span data-stu-id="22e1f-1157">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`|<span data-ttu-id="22e1f-1158">Object</span><span class="sxs-lookup"><span data-stu-id="22e1f-1158">Object</span></span>|<span data-ttu-id="22e1f-1159">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="22e1f-1159">&lt;optional&gt;</span></span>|<span data-ttu-id="22e1f-1160">Entwickler können ein Objekt bereitstellen, auf das sie in der Callbackmethode zugreifen möchten.</span><span class="sxs-lookup"><span data-stu-id="22e1f-1160">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`|<span data-ttu-id="22e1f-1161">Funktion</span><span class="sxs-lookup"><span data-stu-id="22e1f-1161">function</span></span>||<span data-ttu-id="22e1f-1162">Wenn die Methode abgeschlossen wird, wird die Funktion , die im `callback`-Parameter übergeben wird, mit einem einzelnen Parameter (`asyncResult`) aufgerufen, der ein [`AsyncResult`](/javascript/api/office/office.asyncresult)-Objekt darstellt.</span><span class="sxs-lookup"><span data-stu-id="22e1f-1162">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span><br/><br/><span data-ttu-id="22e1f-1163">Rufen Sie für den Zugriff auf die ausgewählten Daten aus der Rückrufmethode `asyncResult.value.data` auf.</span><span class="sxs-lookup"><span data-stu-id="22e1f-1163">To access the selected data from the callback method, call `asyncResult.value.data`.</span></span> <span data-ttu-id="22e1f-1164">Rufen Sie die Source-Eigenschaft für den Zugriff, die die Auswahl stammen, `asyncResult.value.sourceProperty`, die entweder sein `body` oder `subject`.</span><span class="sxs-lookup"><span data-stu-id="22e1f-1164">To access the source property that the selection comes from, call `asyncResult.value.sourceProperty`, which will be either `body` or `subject`.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="22e1f-1165">Anforderungen</span><span class="sxs-lookup"><span data-stu-id="22e1f-1165">Requirements</span></span>

|<span data-ttu-id="22e1f-1166">Anforderung</span><span class="sxs-lookup"><span data-stu-id="22e1f-1166">Requirement</span></span>|<span data-ttu-id="22e1f-1167">Wert</span><span class="sxs-lookup"><span data-stu-id="22e1f-1167">Value</span></span>|
|---|---|
|[<span data-ttu-id="22e1f-1168">Mindestversion des Postfachanforderungssatzes</span><span class="sxs-lookup"><span data-stu-id="22e1f-1168">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="22e1f-1169">1.2</span><span class="sxs-lookup"><span data-stu-id="22e1f-1169">1.2</span></span>|
|[<span data-ttu-id="22e1f-1170">Mindestberechtigungsstufe</span><span class="sxs-lookup"><span data-stu-id="22e1f-1170">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="22e1f-1171">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="22e1f-1171">ReadWriteItem</span></span>|
|[<span data-ttu-id="22e1f-1172">Zutreffender Outlook-Modus</span><span class="sxs-lookup"><span data-stu-id="22e1f-1172">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="22e1f-1173">Verfassen</span><span class="sxs-lookup"><span data-stu-id="22e1f-1173">Compose</span></span>|

##### <a name="returns"></a><span data-ttu-id="22e1f-1174">Gibt zurück:</span><span class="sxs-lookup"><span data-stu-id="22e1f-1174">Returns:</span></span>

<span data-ttu-id="22e1f-1175">Die ausgewählten Daten als Zeichenfolge mit dem durch `coercionType` bestimmten Format.</span><span class="sxs-lookup"><span data-stu-id="22e1f-1175">The selected data as a string with format determined by `coercionType`.</span></span>

<dl class="param-type"><span data-ttu-id="22e1f-1176">

<dt>
Typ</dt>


</span><span class="sxs-lookup"><span data-stu-id="22e1f-1176">

<dt>Type</dt>

</span></span><dd><span data-ttu-id="22e1f-1177">Zeichenfolge</span><span class="sxs-lookup"><span data-stu-id="22e1f-1177">String</span></span></dd>

</dl>

##### <a name="example"></a><span data-ttu-id="22e1f-1178">Beispiel</span><span class="sxs-lookup"><span data-stu-id="22e1f-1178">Example</span></span>

```
// getting selected data
Office.initialize = function () {
    Office.context.mailbox.item.getSelectedDataAsync(Office.CoercionType.Text, {}, getCallback);
}

function getCallback(asyncResult) {
    var text = asyncResult.value.data;
    var prop = asyncResult.value.sourceProperty;

    Office.context.mailbox.item.setSelectedDataAsync('Setting ' + prop + ': ' + text, {}, setCallback);
}

function setCallback(asyncResult) {
    // check for errors
}
```

#### <a name="getselectedentities--entitiesjavascriptapioutlookofficeentities"></a><span data-ttu-id="22e1f-1179">getSelectedEntities() → {[Entitäten](/javascript/api/outlook/office.entities)}</span><span class="sxs-lookup"><span data-stu-id="22e1f-1179">getSelectedEntities() → {[Entities](/javascript/api/outlook/office.entities)}</span></span>

<span data-ttu-id="22e1f-p168">Ruft die Entitäten ab, die in einer hervorgehobenen Übereinstimmung gefunden werden, die ein Benutzer ausgewählt hat. Hervorgehobene Übereinstimmungen gelten für [Kontext-Add-Ins](https://docs.microsoft.com/outlook/add-ins/contextual-outlook-add-ins).</span><span class="sxs-lookup"><span data-stu-id="22e1f-p168">Gets the entities found in a highlighted match a user has selected. Highlighted matches apply to [contextual add-ins](https://docs.microsoft.com/outlook/add-ins/contextual-outlook-add-ins).</span></span>

> [!NOTE]
> <span data-ttu-id="22e1f-1182">Diese Methode wird in Outlook für iOS oder Outlook für Android nicht unterstützt.</span><span class="sxs-lookup"><span data-stu-id="22e1f-1182">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

##### <a name="requirements"></a><span data-ttu-id="22e1f-1183">Anforderungen</span><span class="sxs-lookup"><span data-stu-id="22e1f-1183">Requirements</span></span>

|<span data-ttu-id="22e1f-1184">Anforderung</span><span class="sxs-lookup"><span data-stu-id="22e1f-1184">Requirement</span></span>|<span data-ttu-id="22e1f-1185">Wert</span><span class="sxs-lookup"><span data-stu-id="22e1f-1185">Value</span></span>|
|---|---|
|[<span data-ttu-id="22e1f-1186">Mindestversion des Postfachanforderungssatzes</span><span class="sxs-lookup"><span data-stu-id="22e1f-1186">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="22e1f-1187">1.6</span><span class="sxs-lookup"><span data-stu-id="22e1f-1187">1.6</span></span>|
|[<span data-ttu-id="22e1f-1188">Mindestberechtigungsstufe</span><span class="sxs-lookup"><span data-stu-id="22e1f-1188">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="22e1f-1189">ReadItem</span><span class="sxs-lookup"><span data-stu-id="22e1f-1189">ReadItem</span></span>|
|[<span data-ttu-id="22e1f-1190">Zutreffender Outlook-Modus</span><span class="sxs-lookup"><span data-stu-id="22e1f-1190">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="22e1f-1191">Lesen</span><span class="sxs-lookup"><span data-stu-id="22e1f-1191">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="22e1f-1192">Gibt zurück:</span><span class="sxs-lookup"><span data-stu-id="22e1f-1192">Returns:</span></span>

<span data-ttu-id="22e1f-1193">Typ: [Entitäten](/javascript/api/outlook/office.entities)</span><span class="sxs-lookup"><span data-stu-id="22e1f-1193">Type: [Entities](/javascript/api/outlook/office.entities)</span></span>

##### <a name="example"></a><span data-ttu-id="22e1f-1194">Beispiel</span><span class="sxs-lookup"><span data-stu-id="22e1f-1194">Example</span></span>

<span data-ttu-id="22e1f-1195">Im folgenden Beispiel wird auf die Adressentitäten in der hervorgehobenen Übereinstimmung zugegriffen, die der Benutzer ausgewählt hat.</span><span class="sxs-lookup"><span data-stu-id="22e1f-1195">The following example accesses the addresses entities in the highlighted match selected by the user.</span></span>

```
var contacts = Office.context.mailbox.item.getSelectedEntities().addresses;
```

#### <a name="getselectedregexmatches--object"></a><span data-ttu-id="22e1f-1196">getSelectedRegExMatches() → {Object}</span><span class="sxs-lookup"><span data-stu-id="22e1f-1196">getSelectedRegExMatches() → {Object}</span></span>

<span data-ttu-id="22e1f-p169">Gibt Zeichenfolgenwerte in einer hervorgehobenen Übereinstimmung zurück, die den in der XML-Manifestdatei definierten regulären Ausdrücken entsprechen. Hervorgehobene Übereinstimmungen gelten für [Kontext-Add-Ins](https://docs.microsoft.com/outlook/add-ins/contextual-outlook-add-ins).</span><span class="sxs-lookup"><span data-stu-id="22e1f-p169">Returns string values in a highlighted match that match the regular expressions defined in the manifest XML file. Highlighted matches apply to [contextual add-ins](https://docs.microsoft.com/outlook/add-ins/contextual-outlook-add-ins).</span></span>

> [!NOTE]
> <span data-ttu-id="22e1f-1199">Diese Methode wird in Outlook für iOS oder Outlook für Android nicht unterstützt.</span><span class="sxs-lookup"><span data-stu-id="22e1f-1199">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="22e1f-p170">Die `getSelectedRegExMatches`-Methode gibt die Zeichenfolgen zurück, die dem im `ItemHasRegularExpressionMatch`- oder `ItemHasKnownEntity`-Regelelement der XML-Manifestdatei definierten regulären Ausdruck entsprechen. Bei einer `ItemHasRegularExpressionMatch`-Regel muss eine entsprechende Zeichenfolge in der Eigenschaft des Elements vorhanden sein, das von dieser Regel angegeben wird. Der einfache `PropertyName`-Typ definiert die unterstützten Eigenschaften.</span><span class="sxs-lookup"><span data-stu-id="22e1f-p170">The `getSelectedRegExMatches` method returns the strings that match the regular expression defined in each `ItemHasRegularExpressionMatch` or `ItemHasKnownEntity` rule element in the manifest XML file. For an `ItemHasRegularExpressionMatch` rule, a matching string has to occur in the property of the item that is specified by that rule. The `PropertyName` simple type defines the supported properties.</span></span>

<span data-ttu-id="22e1f-1203">Nehmen Sie z. B. an, dass ein Add-In-Manifest das folgende `Rule`-Element aufweist:</span><span class="sxs-lookup"><span data-stu-id="22e1f-1203">For example, consider an add-in manifest has the following `Rule` element:</span></span>

```
<Rule xsi:type="RuleCollection" Mode="And">
  <Rule xsi:type="ItemIs" FormType="Read" ItemType="Message" />
  <Rule xsi:type="RuleCollection" Mode="Or">
    <Rule xsi:type="ItemHasRegularExpressionMatch" RegExName="fruits" RegExValue="apple|banana|coconut" PropertyName="BodyAsPlaintext" IgnoreCase="true" />
    <Rule xsi:type="ItemHasRegularExpressionMatch" RegExName="veggies" RegExValue="tomato|onion|spinach|broccoli" PropertyName="BodyAsPlaintext" IgnoreCase="true" />
  </Rule>
</Rule>
```

<span data-ttu-id="22e1f-1204">Das von `getRegExMatches` zurückgegebene Objekt hätte zwei Eigenschaften: `fruits` und `veggies`.</span><span class="sxs-lookup"><span data-stu-id="22e1f-1204">The object returned from `getRegExMatches` would have two properties: `fruits` and `veggies`.</span></span>

```
{
  'fruits': ['apple','banana','Banana','coconut'],
  'veggies': ['tomato','onion','spinach','broccoli']
}
```

<span data-ttu-id="22e1f-p171">Wenn Sie eine `ItemHasRegularExpressionMatch`-Regel für die Textkörpereigenschaft eines Elements festlegen, sollte der reguläre Ausdruck den Textkörper weiter filtern und nicht versuchen, den gesamten Textkörper des Elements zurückzugeben. Wenn der gesamte Textkörper eines Elements mit einem regulären Ausdruck wie `.*` abgerufen wird, werden nicht immer die gewünschten Ergebnisse erzielt. Verwenden Sie stattdessen die [`Body.getAsync`](/javascript/api/outlook/office.body#getasync-coerciontype--options--callback-)-Methode, um den gesamten Textkörper abzurufen.</span><span class="sxs-lookup"><span data-stu-id="22e1f-p171">If you specify an `ItemHasRegularExpressionMatch` rule on the body property of an item, the regular expression should further filter the body and should not attempt to return the entire body of the item. Using a regular expression such as `.*` to obtain the entire body of an item does not always return the expected results. Instead, use the [`Body.getAsync`](/javascript/api/outlook/office.body#getasync-coerciontype--options--callback-) method to retrieve the entire body.</span></span>

##### <a name="requirements"></a><span data-ttu-id="22e1f-1208">Anforderungen</span><span class="sxs-lookup"><span data-stu-id="22e1f-1208">Requirements</span></span>

|<span data-ttu-id="22e1f-1209">Anforderung</span><span class="sxs-lookup"><span data-stu-id="22e1f-1209">Requirement</span></span>|<span data-ttu-id="22e1f-1210">Wert</span><span class="sxs-lookup"><span data-stu-id="22e1f-1210">Value</span></span>|
|---|---|
|[<span data-ttu-id="22e1f-1211">Mindestversion des Postfachanforderungssatzes</span><span class="sxs-lookup"><span data-stu-id="22e1f-1211">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="22e1f-1212">1.6</span><span class="sxs-lookup"><span data-stu-id="22e1f-1212">1.6</span></span>|
|[<span data-ttu-id="22e1f-1213">Mindestberechtigungsstufe</span><span class="sxs-lookup"><span data-stu-id="22e1f-1213">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="22e1f-1214">ReadItem</span><span class="sxs-lookup"><span data-stu-id="22e1f-1214">ReadItem</span></span>|
|[<span data-ttu-id="22e1f-1215">Zutreffender Outlook-Modus</span><span class="sxs-lookup"><span data-stu-id="22e1f-1215">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="22e1f-1216">Lesen</span><span class="sxs-lookup"><span data-stu-id="22e1f-1216">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="22e1f-1217">Gibt zurück:</span><span class="sxs-lookup"><span data-stu-id="22e1f-1217">Returns:</span></span>

<span data-ttu-id="22e1f-p172">Ein Objekt mit Arrays aus Zeichenfolgen, die den in der XML-Manifestdatei definierten regulären Ausdrücken entsprechen. Der Name der einzelnen Arrays ist gleich dem entsprechenden Wert des `RegExName`-Attributs der entsprechenden `ItemHasRegularExpressionMatch`-Regel oder des `FilterName`-Attributs der entsprechenden `ItemHasKnownEntity`-Regel.</span><span class="sxs-lookup"><span data-stu-id="22e1f-p172">An object that contains arrays of strings that match the regular expressions defined in the manifest XML file. The name of each array is equal to the corresponding value of the `RegExName` attribute of the matching `ItemHasRegularExpressionMatch` rule or the `FilterName` attribute of the matching `ItemHasKnownEntity` rule.</span></span>

##### <a name="example"></a><span data-ttu-id="22e1f-1220">Beispiel</span><span class="sxs-lookup"><span data-stu-id="22e1f-1220">Example</span></span>

<span data-ttu-id="22e1f-1221">Das folgende Beispiel zeigt, wie Sie auf das Array von Übereinstimmungen für die regulären Ausdrucksregelelemente `fruits` und `veggies` zugreifen, die im Manifest angegeben sind.</span><span class="sxs-lookup"><span data-stu-id="22e1f-1221">The following example shows how to access the array of matches for the regular expression rule elements `fruits` and `veggies`, which are specified in the manifest.</span></span>

```
var selectedMatches = Office.context.mailbox.item.getSelectedRegExMatches();
var fruits = selectedMatches.fruits;
var veggies = selectedMatches.veggies;
```

####  <a name="loadcustompropertiesasynccallback-usercontext"></a><span data-ttu-id="22e1f-1222">loadCustomPropertiesAsync(callback, [userContext])</span><span class="sxs-lookup"><span data-stu-id="22e1f-1222">loadCustomPropertiesAsync(callback, [userContext])</span></span>

<span data-ttu-id="22e1f-1223">Lädt asynchron benutzerdefinierte Eigenschaften für dieses Add-In für das ausgewählte Element.</span><span class="sxs-lookup"><span data-stu-id="22e1f-1223">Asynchronously loads custom properties for this add-in on the selected item.</span></span>

<span data-ttu-id="22e1f-p173">Benutzerdefinierte Eigenschaften werden als Schlüssel-/Wert-Paare pro App und pro Element gespeichert. Diese Methode gibt ein `CustomProperties`-Objekt im Rückruf zurück, das Methoden für den Zugriff auf die benutzerdefinierten Eigenschaften für das aktuelle Element und das aktuelle Add-In bietet. Benutzerdefinierte Eigenschaften sind für das Element nicht verschlüsselt und sollten darum nicht als sicherer Speicher verwendet werden.</span><span class="sxs-lookup"><span data-stu-id="22e1f-p173">Custom properties are stored as key/value pairs on a per-app, per-item basis. This method returns a `CustomProperties` object in the callback, which provides methods to access the custom properties specific to the current item and the current add-in. Custom properties are not encrypted on the item, so this should not be used as secure storage.</span></span>

##### <a name="parameters"></a><span data-ttu-id="22e1f-1227">Parameter:</span><span class="sxs-lookup"><span data-stu-id="22e1f-1227">Parameters:</span></span>

|<span data-ttu-id="22e1f-1228">Name</span><span class="sxs-lookup"><span data-stu-id="22e1f-1228">Name</span></span>|<span data-ttu-id="22e1f-1229">Typ</span><span class="sxs-lookup"><span data-stu-id="22e1f-1229">Type</span></span>|<span data-ttu-id="22e1f-1230">Attribute</span><span class="sxs-lookup"><span data-stu-id="22e1f-1230">Attributes</span></span>|<span data-ttu-id="22e1f-1231">Beschreibung</span><span class="sxs-lookup"><span data-stu-id="22e1f-1231">Description</span></span>|
|---|---|---|---|
|`callback`|<span data-ttu-id="22e1f-1232">Funktion</span><span class="sxs-lookup"><span data-stu-id="22e1f-1232">function</span></span>||<span data-ttu-id="22e1f-1233">Wenn die Methode abgeschlossen wird, wird die Funktion , die im `callback`-Parameter übergeben wird, mit einem einzelnen Parameter (`asyncResult`) aufgerufen, der ein [`AsyncResult`](/javascript/api/office/office.asyncresult)-Objekt darstellt.</span><span class="sxs-lookup"><span data-stu-id="22e1f-1233">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span><br/><br/><span data-ttu-id="22e1f-1234">Die benutzerdefinierten Eigenschaften werden als [`CustomProperties`](/javascript/api/outlook/office.customproperties)-Objekt in der `asyncResult.value`-Eigenschaft bereitgestellt.</span><span class="sxs-lookup"><span data-stu-id="22e1f-1234">The custom properties are provided as a [`CustomProperties`](/javascript/api/outlook/office.customproperties) object in the `asyncResult.value` property.</span></span> <span data-ttu-id="22e1f-1235">Dieses Objekt kann zum Abrufen, festlegen und Entfernen benutzerdefinierter Eigenschaften aus dem Element und speichern Sie Änderungen an den benutzerdefinierten Eigenschaftensatz zurück an den Server verwendet werden.</span><span class="sxs-lookup"><span data-stu-id="22e1f-1235">This object can be used to get, set, and remove custom properties from the item and save changes to the custom property set back to the server.</span></span>|
|`userContext`|<span data-ttu-id="22e1f-1236">Objekt</span><span class="sxs-lookup"><span data-stu-id="22e1f-1236">Object</span></span>|<span data-ttu-id="22e1f-1237">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="22e1f-1237">&lt;optional&gt;</span></span>|<span data-ttu-id="22e1f-1238">Entwickler können ein beliebiges Objekt bereitstellen, den sie in der Rückruffunktion zugreifen möchten.</span><span class="sxs-lookup"><span data-stu-id="22e1f-1238">Developers can provide any object they wish to access in the callback function.</span></span> <span data-ttu-id="22e1f-1239">Dieses Objekt zugegriffen werden kann, indem die `asyncResult.asyncContext` -Eigenschaft in der Callback-Funktion.</span><span class="sxs-lookup"><span data-stu-id="22e1f-1239">This object can be accessed by the `asyncResult.asyncContext` property in the callback function.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="22e1f-1240">Anforderungen</span><span class="sxs-lookup"><span data-stu-id="22e1f-1240">Requirements</span></span>

|<span data-ttu-id="22e1f-1241">Anforderung</span><span class="sxs-lookup"><span data-stu-id="22e1f-1241">Requirement</span></span>|<span data-ttu-id="22e1f-1242">Wert</span><span class="sxs-lookup"><span data-stu-id="22e1f-1242">Value</span></span>|
|---|---|
|[<span data-ttu-id="22e1f-1243">Mindestversion des Postfachanforderungssatzes</span><span class="sxs-lookup"><span data-stu-id="22e1f-1243">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="22e1f-1244">1.0</span><span class="sxs-lookup"><span data-stu-id="22e1f-1244">1.0</span></span>|
|[<span data-ttu-id="22e1f-1245">Mindestberechtigungsstufe</span><span class="sxs-lookup"><span data-stu-id="22e1f-1245">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="22e1f-1246">ReadItem</span><span class="sxs-lookup"><span data-stu-id="22e1f-1246">ReadItem</span></span>|
|[<span data-ttu-id="22e1f-1247">Zutreffender Outlook-Modus</span><span class="sxs-lookup"><span data-stu-id="22e1f-1247">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="22e1f-1248">Verfassen oder Lesen</span><span class="sxs-lookup"><span data-stu-id="22e1f-1248">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="22e1f-1249">Beispiel</span><span class="sxs-lookup"><span data-stu-id="22e1f-1249">Example</span></span>

<span data-ttu-id="22e1f-p176">Das folgende Codebeispiel veranschaulicht die Verwendung der `loadCustomPropertiesAsync`-Methode zum asynchronen Laden der benutzerdefinierten Eigenschaften, die für das aktuelle Element spezifisch sind. Des Weiteren veranschaulicht das Beispiel die Verwendung der `CustomProperties.saveAsync`-Methode zum Speichern der Eigenschaften auf dem Server. Nach dem Laden der benutzerdefinierten Eigenschaften wird in dem Codebeispiel die `CustomProperties.get`-Methode dazu verwendet, die benutzerdefinierte `myProp`-Eigenschaft zu lesen. Die `CustomProperties.set`-Methode wird dazu verwendet, die benutzerdefinierte `otherProp`-Eigenschaft zu schreiben. Schließlich wird die `saveAsync`-Methode aufgerufen, um die benutzerdefinierten Eigenschaften zu speichern.</span><span class="sxs-lookup"><span data-stu-id="22e1f-p176">The following code example shows how to use the `loadCustomPropertiesAsync` method to asynchronously load custom properties that are specific to the current item. The example also shows how to use the `CustomProperties.saveAsync` method to save these properties back to the server. After loading the custom properties, the code sample uses the `CustomProperties.get` method to read the custom property `myProp`, the `CustomProperties.set` method to write the custom property `otherProp`, and then finally calls the `saveAsync` method to save the custom properties.</span></span>

```
// The initialize function is required for all add-ins.
Office.initialize = function () {
  // Checks for the DOM to load using the jQuery ready function.
  $(document).ready(function () {
  // After the DOM is loaded, add-in-specific code can run.
  var item = Office.context.mailbox.item;
  item.loadCustomPropertiesAsync(customPropsCallback);
  });
}

function customPropsCallback(asyncResult) {
  var customProps = asyncResult.value;
  var myProp = customProps.get("myProp");

  customProps.set("otherProp", "value");
  customProps.saveAsync(saveCallback);
}

function saveCallback(asyncResult) {
}
```

####  <a name="removeattachmentasyncattachmentid-options-callback"></a><span data-ttu-id="22e1f-1253">removeAttachmentAsync(attachmentId, [options], [callback])</span><span class="sxs-lookup"><span data-stu-id="22e1f-1253">removeAttachmentAsync(attachmentId, [options], [callback])</span></span>

<span data-ttu-id="22e1f-1254">Entfernt eine Anlage aus einer Nachricht oder einem Termin.</span><span class="sxs-lookup"><span data-stu-id="22e1f-1254">Removes an attachment from a message or appointment.</span></span>

<span data-ttu-id="22e1f-p177">Die `removeAttachmentAsync`-Methode entfernt die Anlage mit dem angegebenen Bezeichner aus dem Element. Als bewährte Vorgehensweise sollten Sie den Anlagenbezeichner nur dann zum Entfernen einer Anlage verwenden, wenn die gleiche Mail-App die Anlage in der gleichen Sitzung hinzugefügt hat. In Outlook Web App und OWA für Geräte ist der Anlagenbezeichner nur innerhalb der gleichen Sitzung gültig. Eine Sitzung ist abgeschlossen, wenn der Benutzer die App schließt, oder wenn der Benutzer in einem eingebetteten Formular mit dem Verfassen beginnt und den Vorgang dann in einem separaten Fenster fortsetzt.</span><span class="sxs-lookup"><span data-stu-id="22e1f-p177">The `removeAttachmentAsync` method removes the attachment with the specified identifier from the item. As a best practice, you should use the attachment identifier to remove an attachment only if the same mail app has added that attachment in the same session. In Outlook Web App and OWA for Devices, the attachment identifier is valid only within the same session. A session is over when the user closes the app, or if the user starts composing in an inline form and subsequently pops out the inline form to continue in a separate window.</span></span>

##### <a name="parameters"></a><span data-ttu-id="22e1f-1259">Parameter:</span><span class="sxs-lookup"><span data-stu-id="22e1f-1259">Parameters:</span></span>

|<span data-ttu-id="22e1f-1260">Name</span><span class="sxs-lookup"><span data-stu-id="22e1f-1260">Name</span></span>|<span data-ttu-id="22e1f-1261">Typ</span><span class="sxs-lookup"><span data-stu-id="22e1f-1261">Type</span></span>|<span data-ttu-id="22e1f-1262">Attribute</span><span class="sxs-lookup"><span data-stu-id="22e1f-1262">Attributes</span></span>|<span data-ttu-id="22e1f-1263">Beschreibung</span><span class="sxs-lookup"><span data-stu-id="22e1f-1263">Description</span></span>|
|---|---|---|---|
|`attachmentId`|<span data-ttu-id="22e1f-1264">Zeichenfolge</span><span class="sxs-lookup"><span data-stu-id="22e1f-1264">String</span></span>||<span data-ttu-id="22e1f-p178">Der Bezeichner des Anhangs, der entfernt werden soll. Die maximale Länge der Zeichenfolge ist 100 Zeichen.</span><span class="sxs-lookup"><span data-stu-id="22e1f-p178">The identifier of the attachment to remove. The maximum length of the string is 100 characters.</span></span>|
|`options`|<span data-ttu-id="22e1f-1267">Object</span><span class="sxs-lookup"><span data-stu-id="22e1f-1267">Object</span></span>|<span data-ttu-id="22e1f-1268">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="22e1f-1268">&lt;optional&gt;</span></span>|<span data-ttu-id="22e1f-1269">Ein Objektliteral, das mindestens eine der folgenden Eigenschaften enthält.</span><span class="sxs-lookup"><span data-stu-id="22e1f-1269">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`|<span data-ttu-id="22e1f-1270">Object</span><span class="sxs-lookup"><span data-stu-id="22e1f-1270">Object</span></span>|<span data-ttu-id="22e1f-1271">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="22e1f-1271">&lt;optional&gt;</span></span>|<span data-ttu-id="22e1f-1272">Entwickler können ein Objekt bereitstellen, auf das sie in der Callbackmethode zugreifen möchten.</span><span class="sxs-lookup"><span data-stu-id="22e1f-1272">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`|<span data-ttu-id="22e1f-1273">Funktion</span><span class="sxs-lookup"><span data-stu-id="22e1f-1273">function</span></span>|<span data-ttu-id="22e1f-1274">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="22e1f-1274">&lt;optional&gt;</span></span>|<span data-ttu-id="22e1f-1275">Wenn die Methode abgeschlossen wird, wird die Funktion , die im `callback`-Parameter übergeben wird, mit einem einzelnen Parameter (`asyncResult`) aufgerufen, der ein [`AsyncResult`](/javascript/api/office/office.asyncresult)-Objekt darstellt.</span><span class="sxs-lookup"><span data-stu-id="22e1f-1275">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span> <br/><span data-ttu-id="22e1f-1276">Wenn beim Entfernen der Anlage ein Fehler auftritt, enthält die Eigenschaft `asyncResult.error` einen Fehlercode mit dem Grund für den Fehler.</span><span class="sxs-lookup"><span data-stu-id="22e1f-1276">If removing the attachment fails, the `asyncResult.error` property will contain an error code with the reason for the failure.</span></span>|

##### <a name="errors"></a><span data-ttu-id="22e1f-1277">Fehler</span><span class="sxs-lookup"><span data-stu-id="22e1f-1277">Errors</span></span>

|<span data-ttu-id="22e1f-1278">Fehlercode</span><span class="sxs-lookup"><span data-stu-id="22e1f-1278">Error code</span></span>|<span data-ttu-id="22e1f-1279">Beschreibung</span><span class="sxs-lookup"><span data-stu-id="22e1f-1279">Description</span></span>|
|------------|-------------|
|`InvalidAttachmentId`|<span data-ttu-id="22e1f-1280">Der Bezeichner für die Anlage ist nicht vorhanden.</span><span class="sxs-lookup"><span data-stu-id="22e1f-1280">The attachment identifier does not exist.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="22e1f-1281">Anforderungen</span><span class="sxs-lookup"><span data-stu-id="22e1f-1281">Requirements</span></span>

|<span data-ttu-id="22e1f-1282">Anforderung</span><span class="sxs-lookup"><span data-stu-id="22e1f-1282">Requirement</span></span>|<span data-ttu-id="22e1f-1283">Wert</span><span class="sxs-lookup"><span data-stu-id="22e1f-1283">Value</span></span>|
|---|---|
|[<span data-ttu-id="22e1f-1284">Mindestversion des Postfachanforderungssatzes</span><span class="sxs-lookup"><span data-stu-id="22e1f-1284">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="22e1f-1285">1.1</span><span class="sxs-lookup"><span data-stu-id="22e1f-1285">1.1</span></span>|
|[<span data-ttu-id="22e1f-1286">Mindestberechtigungsstufe</span><span class="sxs-lookup"><span data-stu-id="22e1f-1286">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="22e1f-1287">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="22e1f-1287">ReadWriteItem</span></span>|
|[<span data-ttu-id="22e1f-1288">Zutreffender Outlook-Modus</span><span class="sxs-lookup"><span data-stu-id="22e1f-1288">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="22e1f-1289">Verfassen</span><span class="sxs-lookup"><span data-stu-id="22e1f-1289">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="22e1f-1290">Beispiel</span><span class="sxs-lookup"><span data-stu-id="22e1f-1290">Example</span></span>

<span data-ttu-id="22e1f-1291">Mit dem folgende Code wird eine Anlage mit dem Bezeichner "0" entfernt.</span><span class="sxs-lookup"><span data-stu-id="22e1f-1291">The following code removes an attachment with an identifier of '0'.</span></span>

```
Office.context.mailbox.item.removeAttachmentAsync(
  '0',
  { asyncContext : null },
  function (asyncResult)
  {
    console.log(asyncResult.status);
  }
);
```

####  <a name="removehandlerasynceventtype-handler-options-callback"></a><span data-ttu-id="22e1f-1292">RemoveHandlerAsync (EventType, Handler, [Optionen], [Rückruf])</span><span class="sxs-lookup"><span data-stu-id="22e1f-1292">removeHandlerAsync(eventType, handler, [options], [callback])</span></span>

<span data-ttu-id="22e1f-1293">Entfernt einen Ereignishandler für ein Ereignis unterstützt.</span><span class="sxs-lookup"><span data-stu-id="22e1f-1293">Removes an event handler for a supported event.</span></span>

<span data-ttu-id="22e1f-1294">Derzeit sind die unterstützten Ereignistypen `Office.EventType.AppointmentTimeChanged`, `Office.EventType.RecipientsChanged`, und`Office.EventType.RecurrencePatternChanged`</span><span class="sxs-lookup"><span data-stu-id="22e1f-1294">Currently the supported event types are `Office.EventType.AppointmentTimeChanged`, `Office.EventType.RecipientsChanged`, and `Office.EventType.RecurrencePatternChanged`</span></span>

##### <a name="parameters"></a><span data-ttu-id="22e1f-1295">Parameter:</span><span class="sxs-lookup"><span data-stu-id="22e1f-1295">Parameters:</span></span>

| <span data-ttu-id="22e1f-1296">Name</span><span class="sxs-lookup"><span data-stu-id="22e1f-1296">Name</span></span> | <span data-ttu-id="22e1f-1297">Typ</span><span class="sxs-lookup"><span data-stu-id="22e1f-1297">Type</span></span> | <span data-ttu-id="22e1f-1298">Attribute</span><span class="sxs-lookup"><span data-stu-id="22e1f-1298">Attributes</span></span> | <span data-ttu-id="22e1f-1299">Beschreibung</span><span class="sxs-lookup"><span data-stu-id="22e1f-1299">Description</span></span> |
|---|---|---|---|
| `eventType` | [<span data-ttu-id="22e1f-1300">Office.EventType</span><span class="sxs-lookup"><span data-stu-id="22e1f-1300">Office.EventType</span></span>](office.md#eventtype-string) || <span data-ttu-id="22e1f-1301">Das Ereignis, das den Handler aufrufen soll</span><span class="sxs-lookup"><span data-stu-id="22e1f-1301">The event that should invoke the handler.</span></span> |
| `handler` | <span data-ttu-id="22e1f-1302">Function</span><span class="sxs-lookup"><span data-stu-id="22e1f-1302">Function</span></span> || <span data-ttu-id="22e1f-p179">Die Funktion, die das Ereignis behandeln soll. Die Funktion muss einen einzigen Parameter akzeptieren (ein Objektliteral). Die Eigenschaft `type` dieses Parameters entspricht dem Parameter `eventType`, der an `removeHandlerAsync` übergeben wird.</span><span class="sxs-lookup"><span data-stu-id="22e1f-p179">The function to handle the event. The function must accept a single parameter, which is an object literal. The `type` property on the parameter will match the `eventType` parameter passed to `removeHandlerAsync`.</span></span> |
| `options` | <span data-ttu-id="22e1f-1306">Objekt</span><span class="sxs-lookup"><span data-stu-id="22e1f-1306">Object</span></span> | <span data-ttu-id="22e1f-1307">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="22e1f-1307">&lt;optional&gt;</span></span> | <span data-ttu-id="22e1f-1308">Ein Objektliteral, das mindestens eine der folgenden Eigenschaften enthält.</span><span class="sxs-lookup"><span data-stu-id="22e1f-1308">An object literal that contains one or more of the following properties.</span></span> |
| `options.asyncContext` | <span data-ttu-id="22e1f-1309">Object</span><span class="sxs-lookup"><span data-stu-id="22e1f-1309">Object</span></span> | <span data-ttu-id="22e1f-1310">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="22e1f-1310">&lt;optional&gt;</span></span> | <span data-ttu-id="22e1f-1311">Entwickler können ein Objekt bereitstellen, auf das sie in der Callbackmethode zugreifen möchten.</span><span class="sxs-lookup"><span data-stu-id="22e1f-1311">Developers can provide any object they wish to access in the callback method.</span></span> |
| `callback` | <span data-ttu-id="22e1f-1312">Funktion</span><span class="sxs-lookup"><span data-stu-id="22e1f-1312">function</span></span>| <span data-ttu-id="22e1f-1313">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="22e1f-1313">&lt;optional&gt;</span></span>|<span data-ttu-id="22e1f-1314">Wenn die Methode abgeschlossen wird, wird die Funktion , die im `callback`-Parameter übergeben wird, mit einem einzelnen Parameter (`asyncResult`) aufgerufen, der ein [`AsyncResult`](/javascript/api/office/office.asyncresult)-Objekt darstellt.</span><span class="sxs-lookup"><span data-stu-id="22e1f-1314">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="22e1f-1315">Anforderungen</span><span class="sxs-lookup"><span data-stu-id="22e1f-1315">Requirements</span></span>

|<span data-ttu-id="22e1f-1316">Anforderung</span><span class="sxs-lookup"><span data-stu-id="22e1f-1316">Requirement</span></span>| <span data-ttu-id="22e1f-1317">Wert</span><span class="sxs-lookup"><span data-stu-id="22e1f-1317">Value</span></span>|
|---|---|
|[<span data-ttu-id="22e1f-1318">Mindestversion des Postfachanforderungssatzes</span><span class="sxs-lookup"><span data-stu-id="22e1f-1318">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="22e1f-1319">Vorschau</span><span class="sxs-lookup"><span data-stu-id="22e1f-1319">Preview</span></span> |
|[<span data-ttu-id="22e1f-1320">Mindestberechtigungsstufe</span><span class="sxs-lookup"><span data-stu-id="22e1f-1320">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="22e1f-1321">ReadItem</span><span class="sxs-lookup"><span data-stu-id="22e1f-1321">ReadItem</span></span> |
|[<span data-ttu-id="22e1f-1322">Zutreffender Outlook-Modus</span><span class="sxs-lookup"><span data-stu-id="22e1f-1322">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="22e1f-1323">Verfassen oder Lesen</span><span class="sxs-lookup"><span data-stu-id="22e1f-1323">Compose or read</span></span> |

##### <a name="example"></a><span data-ttu-id="22e1f-1324">Beispiel</span><span class="sxs-lookup"><span data-stu-id="22e1f-1324">Example</span></span>

```
Office.initialize = function (reason) {
  $(document).ready(function () {
    Office.context.mailbox.item.removeHandlerAsync(Office.EventType.RecurrencePatternChanged, loadNewItem, function (result) {
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

####  <a name="saveasyncoptions-callback"></a><span data-ttu-id="22e1f-1325">saveAsync([options], callback)</span><span class="sxs-lookup"><span data-stu-id="22e1f-1325">saveAsync([options], callback)</span></span>

<span data-ttu-id="22e1f-1326">Speicher asynchron ein Element. .</span><span class="sxs-lookup"><span data-stu-id="22e1f-1326">Asynchronously saves an item.</span></span>

<span data-ttu-id="22e1f-p180">Beim Aufrufen speichert diese Methode die aktuelle Nachricht als Entwurf und  gibt die Element-ID über die Callbackmethode zurück. In Outlook Web App oder Outlook im Onlinemodus wird das Element auf dem Server gespeichert. In Outlook im Cache-Modus wird das Element im lokalen Cache gespeichert.</span><span class="sxs-lookup"><span data-stu-id="22e1f-p180">When invoked, this method saves the current message as a draft and returns the item id via the callback method. In Outlook Web App or Outlook in online mode, the item is saved to the server. In Outlook in cached mode, the item is saved to the local cache.</span></span>

> [!NOTE]
> <span data-ttu-id="22e1f-1330">Wenn Ihr Add-In ruft `saveAsync` für ein Element im Entwurfsmodus, um das Abrufen einer `itemId` um mit EWS oder REST-API verwenden, beachten Sie, dass wenn Outlook im Cache-Modus ist, es kann einige Zeit dauern, bevor das Element tatsächlich auf dem Server synchronisiert wird.</span><span class="sxs-lookup"><span data-stu-id="22e1f-1330">If your add-in calls `saveAsync` on an item in compose mode in order to get an `itemId` to use with EWS or the REST API, be aware that when Outlook is in cached mode, it may take some time before the item is actually synced to the server.</span></span> <span data-ttu-id="22e1f-1331">Bis das Element mit synchronisiert ist, die `itemId` wird ein Fehler zurückgegeben.</span><span class="sxs-lookup"><span data-stu-id="22e1f-1331">Until the item is synced, using the `itemId` will return an error.</span></span>

<span data-ttu-id="22e1f-p182">Da Termine keinen Entwurfsstatus haben, wird das Element bei Aufruf von `saveAsync` für einen Termin im Verfassenmodus als normaler Termin im Kalender des Benutzers gespeichert. Für neue Termine, die noch nicht gespeichert wurden, wird keine Einladung gesendet. Beim Speichern eines vorhandenen Termins wird eine Aktualisierung an die hinzugefügten oder entfernten Teilnehmer gesendet.</span><span class="sxs-lookup"><span data-stu-id="22e1f-p182">Since appointments have no draft state, if `saveAsync` is called on an appointment in compose mode, the item will be saved as a normal appointment on the user's calendar. For new appointments that have not been saved before, no invitation will be sent. Saving an existing appointment will send an update to added or removed attendees.</span></span>

> [!NOTE]
> <span data-ttu-id="22e1f-1335">Die folgenden Clients haben unterschiedlichem Verhalten für `saveAsync` Termine im Verfassenmodus:</span><span class="sxs-lookup"><span data-stu-id="22e1f-1335">The following clients have different behavior for `saveAsync` on appointments in compose mode:</span></span>
>
> - <span data-ttu-id="22e1f-1336">Outlook für Mac unterstützt keine `saveAsync` auf einer Besprechung im Verfassenmodus.</span><span class="sxs-lookup"><span data-stu-id="22e1f-1336">Mac Outlook does not support `saveAsync` on a meeting in compose mode.</span></span> <span data-ttu-id="22e1f-1337">Aufrufen von `saveAsync` auf einer Besprechung in Outlook für Mac wird ein Fehler zurückgegeben.</span><span class="sxs-lookup"><span data-stu-id="22e1f-1337">Calling `saveAsync` on a meeting in Mac Outlook will return an error.</span></span>
> - <span data-ttu-id="22e1f-1338">Outlook im Web immer sendet eine Einladung oder zu aktualisieren, wenn `saveAsync` für einen Termin aufgerufen wird, im Verfassenmodus.</span><span class="sxs-lookup"><span data-stu-id="22e1f-1338">Outlook on the web always sends an invitation or update when `saveAsync` is called on an appointment in compose mode.</span></span>

##### <a name="parameters"></a><span data-ttu-id="22e1f-1339">Parameter:</span><span class="sxs-lookup"><span data-stu-id="22e1f-1339">Parameters:</span></span>

|<span data-ttu-id="22e1f-1340">Name</span><span class="sxs-lookup"><span data-stu-id="22e1f-1340">Name</span></span>|<span data-ttu-id="22e1f-1341">Typ</span><span class="sxs-lookup"><span data-stu-id="22e1f-1341">Type</span></span>|<span data-ttu-id="22e1f-1342">Attribute</span><span class="sxs-lookup"><span data-stu-id="22e1f-1342">Attributes</span></span>|<span data-ttu-id="22e1f-1343">Beschreibung</span><span class="sxs-lookup"><span data-stu-id="22e1f-1343">Description</span></span>|
|---|---|---|---|
|`options`|<span data-ttu-id="22e1f-1344">Objekt</span><span class="sxs-lookup"><span data-stu-id="22e1f-1344">Object</span></span>|<span data-ttu-id="22e1f-1345">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="22e1f-1345">&lt;optional&gt;</span></span>|<span data-ttu-id="22e1f-1346">Ein Objektliteral, das mindestens eine der folgenden Eigenschaften enthält.</span><span class="sxs-lookup"><span data-stu-id="22e1f-1346">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`|<span data-ttu-id="22e1f-1347">Object</span><span class="sxs-lookup"><span data-stu-id="22e1f-1347">Object</span></span>|<span data-ttu-id="22e1f-1348">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="22e1f-1348">&lt;optional&gt;</span></span>|<span data-ttu-id="22e1f-1349">Entwickler können ein Objekt bereitstellen, auf das sie in der Callbackmethode zugreifen möchten.</span><span class="sxs-lookup"><span data-stu-id="22e1f-1349">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`|<span data-ttu-id="22e1f-1350">Funktion</span><span class="sxs-lookup"><span data-stu-id="22e1f-1350">function</span></span>||<span data-ttu-id="22e1f-1351">Wenn die Methode abgeschlossen wird, wird die Funktion , die im `callback`-Parameter übergeben wird, mit einem einzelnen Parameter (`asyncResult`) aufgerufen, der ein [`AsyncResult`](/javascript/api/office/office.asyncresult)-Objekt darstellt.</span><span class="sxs-lookup"><span data-stu-id="22e1f-1351">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span><br/><br/><span data-ttu-id="22e1f-1352">Auf Erfolg, erfolgt die Element-ID der `asyncResult.value` Eigenschaft.</span><span class="sxs-lookup"><span data-stu-id="22e1f-1352">On success, the item identifier is provided in the `asyncResult.value` property.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="22e1f-1353">Anforderungen</span><span class="sxs-lookup"><span data-stu-id="22e1f-1353">Requirements</span></span>

|<span data-ttu-id="22e1f-1354">Anforderung</span><span class="sxs-lookup"><span data-stu-id="22e1f-1354">Requirement</span></span>|<span data-ttu-id="22e1f-1355">Wert</span><span class="sxs-lookup"><span data-stu-id="22e1f-1355">Value</span></span>|
|---|---|
|[<span data-ttu-id="22e1f-1356">Mindestversion des Postfachanforderungssatzes</span><span class="sxs-lookup"><span data-stu-id="22e1f-1356">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="22e1f-1357">1.3</span><span class="sxs-lookup"><span data-stu-id="22e1f-1357">1.3</span></span>|
|[<span data-ttu-id="22e1f-1358">Mindestberechtigungsstufe</span><span class="sxs-lookup"><span data-stu-id="22e1f-1358">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="22e1f-1359">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="22e1f-1359">ReadWriteItem</span></span>|
|[<span data-ttu-id="22e1f-1360">Zutreffender Outlook-Modus</span><span class="sxs-lookup"><span data-stu-id="22e1f-1360">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="22e1f-1361">Verfassen</span><span class="sxs-lookup"><span data-stu-id="22e1f-1361">Compose</span></span>|

##### <a name="examples"></a><span data-ttu-id="22e1f-1362">Beispiele</span><span class="sxs-lookup"><span data-stu-id="22e1f-1362">Examples</span></span>

```
Office.context.mailbox.item.saveAsync(
  function callback(result) {
    // Process the result
  });
```

<span data-ttu-id="22e1f-p184">Es folgt ein Beispiel für den `result`-Parameter, der an die Callbackfunktion übergeben wird. Die `value`-Eigenschaft enthält die Element-ID des Elements.</span><span class="sxs-lookup"><span data-stu-id="22e1f-p184">The following is an example of the `result` parameter passed to the callback function. The `value` property contains the item ID of the item.</span></span>

```
{
  "value":"AAMkADI5...AAA=",
  "status":"succeeded"
}
```

####  <a name="setselecteddataasyncdata-options-callback"></a><span data-ttu-id="22e1f-1365">setSelectedDataAsync(data, [options], callback)</span><span class="sxs-lookup"><span data-stu-id="22e1f-1365">setSelectedDataAsync(data, [options], callback)</span></span>

<span data-ttu-id="22e1f-1366">Fügt asynchron Daten in den Textkörper oder Betreff einer Nachricht ein.</span><span class="sxs-lookup"><span data-stu-id="22e1f-1366">Asynchronously inserts data into the body or subject of a message.</span></span>

<span data-ttu-id="22e1f-p185">Die `setSelectedDataAsync`-Methode fügt die angegebene Zeichenfolge an der Cursorposition im Betreff oder im Textkörper ein, oder, falls im Editor Text ausgewählt ist, ersetzt den markierten Text. Wenn sich der Cursor nicht im Textkörper oder im Betreffsfeld befindet, wird ein Fehler zurückgegeben. Nach dem Einfügen wird der Cursor am Ende der eingefügten Inhalte platziert.</span><span class="sxs-lookup"><span data-stu-id="22e1f-p185">The `setSelectedDataAsync` method inserts the specified string at the cursor location in the subject or body of the item, or, if text is selected in the editor, it replaces the selected text. If the cursor is not in the body or subject field, an error is returned. After insertion, the cursor is placed at the end of the inserted content.</span></span>

##### <a name="parameters"></a><span data-ttu-id="22e1f-1370">Parameter:</span><span class="sxs-lookup"><span data-stu-id="22e1f-1370">Parameters:</span></span>

|<span data-ttu-id="22e1f-1371">Name</span><span class="sxs-lookup"><span data-stu-id="22e1f-1371">Name</span></span>|<span data-ttu-id="22e1f-1372">Typ</span><span class="sxs-lookup"><span data-stu-id="22e1f-1372">Type</span></span>|<span data-ttu-id="22e1f-1373">Attribute</span><span class="sxs-lookup"><span data-stu-id="22e1f-1373">Attributes</span></span>|<span data-ttu-id="22e1f-1374">Beschreibung</span><span class="sxs-lookup"><span data-stu-id="22e1f-1374">Description</span></span>|
|---|---|---|---|
|`data`|<span data-ttu-id="22e1f-1375">Zeichenfolge</span><span class="sxs-lookup"><span data-stu-id="22e1f-1375">String</span></span>||<span data-ttu-id="22e1f-p186">Die einzufügenden Daten. Daten dürfen 1.000.000 Zeichen nicht überschreiten. Werden mehr als 1.000.000 Zeichen übergeben, wird eine `ArgumentOutOfRange`-Ausnahme ausgelöst.</span><span class="sxs-lookup"><span data-stu-id="22e1f-p186">The data to be inserted. Data is not to exceed 1,000,000 characters. If more than 1,000,000 characters are passed in, an `ArgumentOutOfRange` exception is thrown.</span></span>|
|`options`|<span data-ttu-id="22e1f-1379">Object</span><span class="sxs-lookup"><span data-stu-id="22e1f-1379">Object</span></span>|<span data-ttu-id="22e1f-1380">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="22e1f-1380">&lt;optional&gt;</span></span>|<span data-ttu-id="22e1f-1381">Ein Objektliteral, das mindestens eine der folgenden Eigenschaften enthält.</span><span class="sxs-lookup"><span data-stu-id="22e1f-1381">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`|<span data-ttu-id="22e1f-1382">Object</span><span class="sxs-lookup"><span data-stu-id="22e1f-1382">Object</span></span>|<span data-ttu-id="22e1f-1383">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="22e1f-1383">&lt;optional&gt;</span></span>|<span data-ttu-id="22e1f-1384">Entwickler können ein Objekt bereitstellen, auf das sie in der Rückrufmethode zugreifen möchten.</span><span class="sxs-lookup"><span data-stu-id="22e1f-1384">Developers can provide any object they wish to access in the callback method.</span></span>|
|`options.coercionType`|[<span data-ttu-id="22e1f-1385">Office.CoercionType</span><span class="sxs-lookup"><span data-stu-id="22e1f-1385">Office.CoercionType</span></span>](office.md#coerciontype-string)|<span data-ttu-id="22e1f-1386">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="22e1f-1386">&lt;optional&gt;</span></span>|<span data-ttu-id="22e1f-p187">Wenn `text`, wird das aktuelle Format in Outlook Web App und Outlook angewendet. Wenn das Feld ein HTML-Editor ist, werden nur die Textdaten eingefügt, selbst wenn es sich bei den Daten um HTML-Daten handelt.</span><span class="sxs-lookup"><span data-stu-id="22e1f-p187">If `text`, the current style is applied in Outlook Web App and Outlook. If the field is an HTML editor, only the text data is inserted, even if the data is HTML.</span></span><br/><br/><span data-ttu-id="22e1f-p188">Wenn `html` gesetzt ist und das Feld HTML unterstützt (der Betreff jedoch nicht), wird die aktuelle Formatvorlage in Outlook Web App angewendet und die Standardformatvorlage in Outlook. Ist das Feld ein Textfeld, wird ein Fehler des Typs `InvalidDataFormat` zurückgegeben.</span><span class="sxs-lookup"><span data-stu-id="22e1f-p188">If `html` and the field supports HTML (the subject doesn't), the current style is applied in Outlook Web App and the default style is applied in Outlook. If the field is a text field, an `InvalidDataFormat` error is returned.</span></span><br/><br/><span data-ttu-id="22e1f-1391">Wenn `coercionType` nicht festgelegt wird, hängt das Ergebnis vom Feld ab: Wenn das Feld HTML ist, wird HTML verwendet. Wenn das Feld Text ist, wird Nur-Text verwendet.</span><span class="sxs-lookup"><span data-stu-id="22e1f-1391">If `coercionType` is not set, the result depends on the field: if the field is HTML then HTML is used; if the field is text, then plain text is used.</span></span>|
|`callback`|<span data-ttu-id="22e1f-1392">function</span><span class="sxs-lookup"><span data-stu-id="22e1f-1392">function</span></span>||<span data-ttu-id="22e1f-1393">Wenn die Methode abgeschlossen wird, wird die Funktion , die im `callback`-Parameter übergeben wird, mit einem einzelnen Parameter (`asyncResult`) aufgerufen, der ein [`AsyncResult`](/javascript/api/office/office.asyncresult)-Objekt darstellt.</span><span class="sxs-lookup"><span data-stu-id="22e1f-1393">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="22e1f-1394">Anforderungen</span><span class="sxs-lookup"><span data-stu-id="22e1f-1394">Requirements</span></span>

|<span data-ttu-id="22e1f-1395">Anforderung</span><span class="sxs-lookup"><span data-stu-id="22e1f-1395">Requirement</span></span>|<span data-ttu-id="22e1f-1396">Wert</span><span class="sxs-lookup"><span data-stu-id="22e1f-1396">Value</span></span>|
|---|---|
|[<span data-ttu-id="22e1f-1397">Mindestversion des Postfachanforderungssatzes</span><span class="sxs-lookup"><span data-stu-id="22e1f-1397">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="22e1f-1398">1.2</span><span class="sxs-lookup"><span data-stu-id="22e1f-1398">1.2</span></span>|
|[<span data-ttu-id="22e1f-1399">Mindestberechtigungsstufe</span><span class="sxs-lookup"><span data-stu-id="22e1f-1399">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="22e1f-1400">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="22e1f-1400">ReadWriteItem</span></span>|
|[<span data-ttu-id="22e1f-1401">Zutreffender Outlook-Modus</span><span class="sxs-lookup"><span data-stu-id="22e1f-1401">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="22e1f-1402">Verfassen</span><span class="sxs-lookup"><span data-stu-id="22e1f-1402">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="22e1f-1403">Beispiel</span><span class="sxs-lookup"><span data-stu-id="22e1f-1403">Example</span></span>

```
Office.context.mailbox.item.setSelectedDataAsync("Hello World!");
Office.context.mailbox.item.setSelectedDataAsync("<b>Hello World!</b>", { coercionType : "html" });
```