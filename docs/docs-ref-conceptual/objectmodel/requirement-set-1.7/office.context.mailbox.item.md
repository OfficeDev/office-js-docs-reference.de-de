
# <a name="item"></a><span data-ttu-id="d9554-101">item</span><span class="sxs-lookup"><span data-stu-id="d9554-101">item</span></span>

### <a name="officeofficemdcontextofficecontextmdmailboxofficecontextmailboxmditem"></a><span data-ttu-id="d9554-102">[Office](office.md)[.context](office.context.md)[.mailbox](office.context.mailbox.md).item</span><span class="sxs-lookup"><span data-stu-id="d9554-102">[Office](office.md)[.context](office.context.md)[.mailbox](office.context.mailbox.md).item</span></span>

<span data-ttu-id="d9554-p101">Der `item`-Namespace wird für den Zugriff auf die aktuell ausgewählte Nachricht, Besprechungsanfrage oder den aktuell ausgewählten Termin verwendet. Sie können den Typ von `item` mithilfe der [itemType](#itemtype-officemailboxenumsitemtypejavascriptapioutlook17officemailboxenumsitemtype)-Eigenschaft bestimmen.</span><span class="sxs-lookup"><span data-stu-id="d9554-p101">The `item` namespace is used to access the currently selected message, meeting request, or appointment. You can determine the type of the `item` by using the [itemType](#itemtype-officemailboxenumsitemtypejavascriptapioutlook17officemailboxenumsitemtype) property.</span></span>

##### <a name="requirements"></a><span data-ttu-id="d9554-105">Anforderungen</span><span class="sxs-lookup"><span data-stu-id="d9554-105">Requirements</span></span>

|<span data-ttu-id="d9554-106">Anforderung</span><span class="sxs-lookup"><span data-stu-id="d9554-106">Requirement</span></span>|<span data-ttu-id="d9554-107">Wert</span><span class="sxs-lookup"><span data-stu-id="d9554-107">Value</span></span>|
|---|---|
|[<span data-ttu-id="d9554-108">Mindestversion des Postfachanforderungssatzes</span><span class="sxs-lookup"><span data-stu-id="d9554-108">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="d9554-109">1.0</span><span class="sxs-lookup"><span data-stu-id="d9554-109">1.0</span></span>|
|[<span data-ttu-id="d9554-110">Mindestberechtigungsstufe</span><span class="sxs-lookup"><span data-stu-id="d9554-110">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="d9554-111">Eingeschränkt</span><span class="sxs-lookup"><span data-stu-id="d9554-111">Restricted</span></span>|
|[<span data-ttu-id="d9554-112">Zutreffender Outlook-Modus</span><span class="sxs-lookup"><span data-stu-id="d9554-112">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="d9554-113">Verfassen oder Lesen</span><span class="sxs-lookup"><span data-stu-id="d9554-113">Compose or read</span></span>|

##### <a name="members-and-methods"></a><span data-ttu-id="d9554-114">Elemente und Methoden</span><span class="sxs-lookup"><span data-stu-id="d9554-114">Members and methods</span></span>

| <span data-ttu-id="d9554-115">Element</span><span class="sxs-lookup"><span data-stu-id="d9554-115">Member</span></span> | <span data-ttu-id="d9554-116">Typ</span><span class="sxs-lookup"><span data-stu-id="d9554-116">Type</span></span> |
|--------|------|
| [<span data-ttu-id="d9554-117">attachments</span><span class="sxs-lookup"><span data-stu-id="d9554-117">attachments</span></span>](#attachments-arrayattachmentdetailsjavascriptapioutlook17officeattachmentdetails) | <span data-ttu-id="d9554-118">Element</span><span class="sxs-lookup"><span data-stu-id="d9554-118">Member</span></span> |
| [<span data-ttu-id="d9554-119">bcc</span><span class="sxs-lookup"><span data-stu-id="d9554-119">bcc</span></span>](#bcc-recipientsjavascriptapioutlook17officerecipients) | <span data-ttu-id="d9554-120">Element</span><span class="sxs-lookup"><span data-stu-id="d9554-120">Member</span></span> |
| [<span data-ttu-id="d9554-121">body</span><span class="sxs-lookup"><span data-stu-id="d9554-121">body</span></span>](#body-bodyjavascriptapioutlook17officebody) | <span data-ttu-id="d9554-122">Element</span><span class="sxs-lookup"><span data-stu-id="d9554-122">Member</span></span> |
| [<span data-ttu-id="d9554-123">cc</span><span class="sxs-lookup"><span data-stu-id="d9554-123">cc</span></span>](#cc-arrayemailaddressdetailsjavascriptapioutlook17officeemailaddressdetailsrecipientsjavascriptapioutlook17officerecipients) | <span data-ttu-id="d9554-124">Element</span><span class="sxs-lookup"><span data-stu-id="d9554-124">Member</span></span> |
| [<span data-ttu-id="d9554-125">conversationId</span><span class="sxs-lookup"><span data-stu-id="d9554-125">conversationId</span></span>](#nullable-conversationid-string) | <span data-ttu-id="d9554-126">Element</span><span class="sxs-lookup"><span data-stu-id="d9554-126">Member</span></span> |
| [<span data-ttu-id="d9554-127">dateTimeCreated</span><span class="sxs-lookup"><span data-stu-id="d9554-127">dateTimeCreated</span></span>](#datetimecreated-date) | <span data-ttu-id="d9554-128">Element</span><span class="sxs-lookup"><span data-stu-id="d9554-128">Member</span></span> |
| [<span data-ttu-id="d9554-129">dateTimeModified</span><span class="sxs-lookup"><span data-stu-id="d9554-129">dateTimeModified</span></span>](#datetimemodified-date) | <span data-ttu-id="d9554-130">Element</span><span class="sxs-lookup"><span data-stu-id="d9554-130">Member</span></span> |
| [<span data-ttu-id="d9554-131">end</span><span class="sxs-lookup"><span data-stu-id="d9554-131">end</span></span>](#end-datetimejavascriptapioutlook17officetime) | <span data-ttu-id="d9554-132">Element</span><span class="sxs-lookup"><span data-stu-id="d9554-132">Member</span></span> |
| [<span data-ttu-id="d9554-133">from</span><span class="sxs-lookup"><span data-stu-id="d9554-133">from</span></span>](#from-emailaddressdetailsjavascriptapioutlook17officeemailaddressdetailsfromjavascriptapioutlook17officefrom) | <span data-ttu-id="d9554-134">Element</span><span class="sxs-lookup"><span data-stu-id="d9554-134">Member</span></span> |
| [<span data-ttu-id="d9554-135">internetMessageId</span><span class="sxs-lookup"><span data-stu-id="d9554-135">internetMessageId</span></span>](#internetmessageid-string) | <span data-ttu-id="d9554-136">Element</span><span class="sxs-lookup"><span data-stu-id="d9554-136">Member</span></span> |
| [<span data-ttu-id="d9554-137">itemClass</span><span class="sxs-lookup"><span data-stu-id="d9554-137">itemClass</span></span>](#itemclass-string) | <span data-ttu-id="d9554-138">Element</span><span class="sxs-lookup"><span data-stu-id="d9554-138">Member</span></span> |
| [<span data-ttu-id="d9554-139">itemId</span><span class="sxs-lookup"><span data-stu-id="d9554-139">itemId</span></span>](#nullable-itemid-string) | <span data-ttu-id="d9554-140">Element</span><span class="sxs-lookup"><span data-stu-id="d9554-140">Member</span></span> |
| [<span data-ttu-id="d9554-141">itemType</span><span class="sxs-lookup"><span data-stu-id="d9554-141">itemType</span></span>](#itemtype-officemailboxenumsitemtypejavascriptapioutlook17officemailboxenumsitemtype) | <span data-ttu-id="d9554-142">Element</span><span class="sxs-lookup"><span data-stu-id="d9554-142">Member</span></span> |
| [<span data-ttu-id="d9554-143">location</span><span class="sxs-lookup"><span data-stu-id="d9554-143">location</span></span>](#location-stringlocationjavascriptapioutlook17officelocation) | <span data-ttu-id="d9554-144">Element</span><span class="sxs-lookup"><span data-stu-id="d9554-144">Member</span></span> |
| [<span data-ttu-id="d9554-145">normalizedSubject</span><span class="sxs-lookup"><span data-stu-id="d9554-145">normalizedSubject</span></span>](#normalizedsubject-string) | <span data-ttu-id="d9554-146">Element</span><span class="sxs-lookup"><span data-stu-id="d9554-146">Member</span></span> |
| [<span data-ttu-id="d9554-147">notificationMessages</span><span class="sxs-lookup"><span data-stu-id="d9554-147">notificationMessages</span></span>](#notificationmessages-notificationmessagesjavascriptapioutlook17officenotificationmessages) | <span data-ttu-id="d9554-148">Element</span><span class="sxs-lookup"><span data-stu-id="d9554-148">Member</span></span> |
| [<span data-ttu-id="d9554-149">optionalAttendees</span><span class="sxs-lookup"><span data-stu-id="d9554-149">optionalAttendees</span></span>](#optionalattendees-arrayemailaddressdetailsjavascriptapioutlook17officeemailaddressdetailsrecipientsjavascriptapioutlook17officerecipients) | <span data-ttu-id="d9554-150">Element</span><span class="sxs-lookup"><span data-stu-id="d9554-150">Member</span></span> |
| [<span data-ttu-id="d9554-151">organizer</span><span class="sxs-lookup"><span data-stu-id="d9554-151">organizer</span></span>](#organizer-emailaddressdetailsjavascriptapioutlook17officeemailaddressdetailsorganizerjavascriptapioutlook17officeorganizer) | <span data-ttu-id="d9554-152">Element</span><span class="sxs-lookup"><span data-stu-id="d9554-152">Member</span></span> |
| [<span data-ttu-id="d9554-153">recurrence</span><span class="sxs-lookup"><span data-stu-id="d9554-153">recurrence</span></span>](#nullable-recurrence-recurrencejavascriptapioutlook17officerecurrence) | <span data-ttu-id="d9554-154">Element</span><span class="sxs-lookup"><span data-stu-id="d9554-154">Member</span></span> |
| [<span data-ttu-id="d9554-155">requiredAttendees</span><span class="sxs-lookup"><span data-stu-id="d9554-155">requiredAttendees</span></span>](#requiredattendees-arrayemailaddressdetailsjavascriptapioutlook17officeemailaddressdetailsrecipientsjavascriptapioutlook17officerecipients) | <span data-ttu-id="d9554-156">Element</span><span class="sxs-lookup"><span data-stu-id="d9554-156">Member</span></span> |
| [<span data-ttu-id="d9554-157">sender</span><span class="sxs-lookup"><span data-stu-id="d9554-157">sender</span></span>](#sender-emailaddressdetailsjavascriptapioutlook17officeemailaddressdetails) | <span data-ttu-id="d9554-158">Element</span><span class="sxs-lookup"><span data-stu-id="d9554-158">Member</span></span> |
| [<span data-ttu-id="d9554-159">seriesId</span><span class="sxs-lookup"><span data-stu-id="d9554-159">seriesId</span></span>](#nullable-seriesid-string) | <span data-ttu-id="d9554-160">Element</span><span class="sxs-lookup"><span data-stu-id="d9554-160">Member</span></span> |
| [<span data-ttu-id="d9554-161">start</span><span class="sxs-lookup"><span data-stu-id="d9554-161">start</span></span>](#start-datetimejavascriptapioutlook17officetime) | <span data-ttu-id="d9554-162">Element</span><span class="sxs-lookup"><span data-stu-id="d9554-162">Member</span></span> |
| [<span data-ttu-id="d9554-163">subject</span><span class="sxs-lookup"><span data-stu-id="d9554-163">subject</span></span>](#subject-stringsubjectjavascriptapioutlook17officesubject) | <span data-ttu-id="d9554-164">Element</span><span class="sxs-lookup"><span data-stu-id="d9554-164">Member</span></span> |
| [<span data-ttu-id="d9554-165">to</span><span class="sxs-lookup"><span data-stu-id="d9554-165">to</span></span>](#to-arrayemailaddressdetailsjavascriptapioutlook17officeemailaddressdetailsrecipientsjavascriptapioutlook17officerecipients) | <span data-ttu-id="d9554-166">Element</span><span class="sxs-lookup"><span data-stu-id="d9554-166">Member</span></span> |
| [<span data-ttu-id="d9554-167">addFileAttachmentAsync</span><span class="sxs-lookup"><span data-stu-id="d9554-167">addFileAttachmentAsync</span></span>](#addfileattachmentasyncuri-attachmentname-options-callback) | <span data-ttu-id="d9554-168">Methode</span><span class="sxs-lookup"><span data-stu-id="d9554-168">Method</span></span> |
| [<span data-ttu-id="d9554-169">addHandlerAsync</span><span class="sxs-lookup"><span data-stu-id="d9554-169">addHandlerAsync</span></span>](#addhandlerasynceventtype-handler-options-callback) | <span data-ttu-id="d9554-170">Methode</span><span class="sxs-lookup"><span data-stu-id="d9554-170">Method</span></span> |
| [<span data-ttu-id="d9554-171">addItemAttachmentAsync</span><span class="sxs-lookup"><span data-stu-id="d9554-171">addItemAttachmentAsync</span></span>](#additemattachmentasyncitemid-attachmentname-options-callback) | <span data-ttu-id="d9554-172">Methode</span><span class="sxs-lookup"><span data-stu-id="d9554-172">Method</span></span> |
| [<span data-ttu-id="d9554-173">close</span><span class="sxs-lookup"><span data-stu-id="d9554-173">close</span></span>](#close) | <span data-ttu-id="d9554-174">Methode</span><span class="sxs-lookup"><span data-stu-id="d9554-174">Method</span></span> |
| [<span data-ttu-id="d9554-175">displayReplyAllForm</span><span class="sxs-lookup"><span data-stu-id="d9554-175">displayReplyAllForm</span></span>](#displayreplyallformformdata) | <span data-ttu-id="d9554-176">Methode</span><span class="sxs-lookup"><span data-stu-id="d9554-176">Method</span></span> |
| [<span data-ttu-id="d9554-177">displayReplyForm</span><span class="sxs-lookup"><span data-stu-id="d9554-177">displayReplyForm</span></span>](#displayreplyformformdata) | <span data-ttu-id="d9554-178">Methode</span><span class="sxs-lookup"><span data-stu-id="d9554-178">Method</span></span> |
| [<span data-ttu-id="d9554-179">getEntities</span><span class="sxs-lookup"><span data-stu-id="d9554-179">getEntities</span></span>](#getentities--entitiesjavascriptapioutlook17officeentities) | <span data-ttu-id="d9554-180">Methode</span><span class="sxs-lookup"><span data-stu-id="d9554-180">Method</span></span> |
| [<span data-ttu-id="d9554-181">getEntitiesByType</span><span class="sxs-lookup"><span data-stu-id="d9554-181">getEntitiesByType</span></span>](#getentitiesbytypeentitytype--nullable-arraystringcontactjavascriptapioutlook17officecontactmeetingsuggestionjavascriptapioutlook17officemeetingsuggestionphonenumberjavascriptapioutlook17officephonenumbertasksuggestionjavascriptapioutlook17officetasksuggestion) | <span data-ttu-id="d9554-182">Methode</span><span class="sxs-lookup"><span data-stu-id="d9554-182">Method</span></span> |
| [<span data-ttu-id="d9554-183">getFilteredEntitiesByName</span><span class="sxs-lookup"><span data-stu-id="d9554-183">getFilteredEntitiesByName</span></span>](#getfilteredentitiesbynamename--nullable-arraystringcontactjavascriptapioutlook17officecontactmeetingsuggestionjavascriptapioutlook17officemeetingsuggestionphonenumberjavascriptapioutlook17officephonenumbertasksuggestionjavascriptapioutlook17officetasksuggestion) | <span data-ttu-id="d9554-184">Methode</span><span class="sxs-lookup"><span data-stu-id="d9554-184">Method</span></span> |
| [<span data-ttu-id="d9554-185">getRegExMatches</span><span class="sxs-lookup"><span data-stu-id="d9554-185">getRegExMatches</span></span>](#getregexmatches--object) | <span data-ttu-id="d9554-186">Methode</span><span class="sxs-lookup"><span data-stu-id="d9554-186">Method</span></span> |
| [<span data-ttu-id="d9554-187">getRegExMatchesByName</span><span class="sxs-lookup"><span data-stu-id="d9554-187">getRegExMatchesByName</span></span>](#getregexmatchesbynamename--nullable-array-string-) | <span data-ttu-id="d9554-188">Methode</span><span class="sxs-lookup"><span data-stu-id="d9554-188">Method</span></span> |
| [<span data-ttu-id="d9554-189">getSelectedDataAsync</span><span class="sxs-lookup"><span data-stu-id="d9554-189">getSelectedDataAsync</span></span>](#getselecteddataasynccoerciontype-options-callback--string) | <span data-ttu-id="d9554-190">Methode</span><span class="sxs-lookup"><span data-stu-id="d9554-190">Method</span></span> |
| [<span data-ttu-id="d9554-191">getSelectedEntities</span><span class="sxs-lookup"><span data-stu-id="d9554-191">getSelectedEntities</span></span>](#getselectedentities--entitiesjavascriptapioutlook17officeentities) | <span data-ttu-id="d9554-192">Methode</span><span class="sxs-lookup"><span data-stu-id="d9554-192">Method</span></span> |
| [<span data-ttu-id="d9554-193">getSelectedRegExMatches</span><span class="sxs-lookup"><span data-stu-id="d9554-193">getSelectedRegExMatches</span></span>](#getselectedregexmatches--object) | <span data-ttu-id="d9554-194">Methode</span><span class="sxs-lookup"><span data-stu-id="d9554-194">Method</span></span> |
| [<span data-ttu-id="d9554-195">loadCustomPropertiesAsync</span><span class="sxs-lookup"><span data-stu-id="d9554-195">loadCustomPropertiesAsync</span></span>](#loadcustompropertiesasynccallback-usercontext) | <span data-ttu-id="d9554-196">Methode</span><span class="sxs-lookup"><span data-stu-id="d9554-196">Method</span></span> |
| [<span data-ttu-id="d9554-197">removeAttachmentAsync</span><span class="sxs-lookup"><span data-stu-id="d9554-197">removeAttachmentAsync</span></span>](#removeattachmentasyncattachmentid-options-callback) | <span data-ttu-id="d9554-198">Methode</span><span class="sxs-lookup"><span data-stu-id="d9554-198">Method</span></span> |
| [<span data-ttu-id="d9554-199">removeHandlerAsync</span><span class="sxs-lookup"><span data-stu-id="d9554-199">removeHandlerAsync</span></span>](#removehandlerasynceventtype-handler-options-callback) | <span data-ttu-id="d9554-200">Methode</span><span class="sxs-lookup"><span data-stu-id="d9554-200">Method</span></span> |
| [<span data-ttu-id="d9554-201">saveAsync</span><span class="sxs-lookup"><span data-stu-id="d9554-201">saveAsync</span></span>](#saveasyncoptions-callback) | <span data-ttu-id="d9554-202">Methode</span><span class="sxs-lookup"><span data-stu-id="d9554-202">Method</span></span> |
| [<span data-ttu-id="d9554-203">setSelectedDataAsync</span><span class="sxs-lookup"><span data-stu-id="d9554-203">setSelectedDataAsync</span></span>](#setselecteddataasyncdata-options-callback) | <span data-ttu-id="d9554-204">Methode</span><span class="sxs-lookup"><span data-stu-id="d9554-204">Method</span></span> |

### <a name="example"></a><span data-ttu-id="d9554-205">Beispiel</span><span class="sxs-lookup"><span data-stu-id="d9554-205">Example</span></span>

<span data-ttu-id="d9554-206">Im folgenden JavaScript-Codebeispiel wird der Zugriff auf die `subject`-Eigenschaft des aktuellen Elements in Outlook veranschaulicht.</span><span class="sxs-lookup"><span data-stu-id="d9554-206">The following JavaScript code example shows how to access the `subject` property of the current item in Outlook.</span></span>

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

### <a name="members"></a><span data-ttu-id="d9554-207">Elemente</span><span class="sxs-lookup"><span data-stu-id="d9554-207">Members</span></span>

#### <a name="attachments-arrayattachmentdetailsjavascriptapioutlook17officeattachmentdetails"></a><span data-ttu-id="d9554-208">attachments :Array.<[AttachmentDetails](/javascript/api/outlook_1_7/office.attachmentdetails)></span><span class="sxs-lookup"><span data-stu-id="d9554-208">attachments :Array.<[AttachmentDetails](/javascript/api/outlook_1_7/office.attachmentdetails)></span></span>

<span data-ttu-id="d9554-p102">Ruft ein Array mit Anlagen für das Element ab. Nur Lesemodus.</span><span class="sxs-lookup"><span data-stu-id="d9554-p102">Gets an array of attachments for the item. Read mode only.</span></span>

> [!NOTE]
> <span data-ttu-id="d9554-211">Bestimmte Dateitypen werden aufgrund von Sicherheitsproblemen von Outlook blockiert und daher nicht zurückgegeben werden.</span><span class="sxs-lookup"><span data-stu-id="d9554-211">Certain types of files are blocked by Outlook due to potential security issues and are therefore not returned.</span></span> <span data-ttu-id="d9554-212">Weitere Informationen finden Sie unter [Blockierte Anlagen in Outlook](https://support.office.com/article/Blocked-attachments-in-Outlook-434752E1-02D3-4E90-9124-8B81E49A8519).</span><span class="sxs-lookup"><span data-stu-id="d9554-212">For more information, see [Blocked attachments in Outlook](https://support.office.com/article/Blocked-attachments-in-Outlook-434752E1-02D3-4E90-9124-8B81E49A8519).</span></span>

##### <a name="type"></a><span data-ttu-id="d9554-213">Typ:</span><span class="sxs-lookup"><span data-stu-id="d9554-213">Type:</span></span>

*   <span data-ttu-id="d9554-214">Array.<[AttachmentDetails](/javascript/api/outlook_1_7/office.attachmentdetails)></span><span class="sxs-lookup"><span data-stu-id="d9554-214">Array.<[AttachmentDetails](/javascript/api/outlook_1_7/office.attachmentdetails)></span></span>

##### <a name="requirements"></a><span data-ttu-id="d9554-215">Anforderungen</span><span class="sxs-lookup"><span data-stu-id="d9554-215">Requirements</span></span>

|<span data-ttu-id="d9554-216">Anforderung</span><span class="sxs-lookup"><span data-stu-id="d9554-216">Requirement</span></span>|<span data-ttu-id="d9554-217">Wert</span><span class="sxs-lookup"><span data-stu-id="d9554-217">Value</span></span>|
|---|---|
|[<span data-ttu-id="d9554-218">Mindestversion des Postfachanforderungssatzes</span><span class="sxs-lookup"><span data-stu-id="d9554-218">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="d9554-219">1.0</span><span class="sxs-lookup"><span data-stu-id="d9554-219">1.0</span></span>|
|[<span data-ttu-id="d9554-220">Mindestberechtigungsstufe</span><span class="sxs-lookup"><span data-stu-id="d9554-220">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="d9554-221">ReadItem</span><span class="sxs-lookup"><span data-stu-id="d9554-221">ReadItem</span></span>|
|[<span data-ttu-id="d9554-222">Zutreffender Outlook-Modus</span><span class="sxs-lookup"><span data-stu-id="d9554-222">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="d9554-223">Lesen</span><span class="sxs-lookup"><span data-stu-id="d9554-223">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="d9554-224">Beispiel</span><span class="sxs-lookup"><span data-stu-id="d9554-224">Example</span></span>

<span data-ttu-id="d9554-225">Mit dem folgende Code wird eine HTML-Zeichenfolge mit Details aller Anlagen für das aktuelle Element erstellt.</span><span class="sxs-lookup"><span data-stu-id="d9554-225">The following code builds an HTML string with details of all attachments on the current item.</span></span>

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

####  <a name="bcc-recipientsjavascriptapioutlook17officerecipients"></a><span data-ttu-id="d9554-226">bcc :[Recipients](/javascript/api/outlook_1_7/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="d9554-226">bcc :[Recipients](/javascript/api/outlook_1_7/office.recipients)</span></span>

<span data-ttu-id="d9554-227">Ruft ein Objekt, das Methoden zum Abrufen oder aktualisieren die Empfänger in der Zeile Bcc (blind Carbon Copy, Blindkopie) einer Nachricht bereitstellt.</span><span class="sxs-lookup"><span data-stu-id="d9554-227">Gets an object that provides methods to get or update the recipients on the Bcc (blind carbon copy) line of a message.</span></span> <span data-ttu-id="d9554-228">Nur Verfassenmodus.</span><span class="sxs-lookup"><span data-stu-id="d9554-228">Compose mode only.</span></span>

##### <a name="type"></a><span data-ttu-id="d9554-229">Typ:</span><span class="sxs-lookup"><span data-stu-id="d9554-229">Type:</span></span>

*   [<span data-ttu-id="d9554-230">Recipients</span><span class="sxs-lookup"><span data-stu-id="d9554-230">Recipients</span></span>](/javascript/api/outlook_1_7/office.recipients)

##### <a name="requirements"></a><span data-ttu-id="d9554-231">Anforderungen</span><span class="sxs-lookup"><span data-stu-id="d9554-231">Requirements</span></span>

|<span data-ttu-id="d9554-232">Anforderung</span><span class="sxs-lookup"><span data-stu-id="d9554-232">Requirement</span></span>|<span data-ttu-id="d9554-233">Wert</span><span class="sxs-lookup"><span data-stu-id="d9554-233">Value</span></span>|
|---|---|
|[<span data-ttu-id="d9554-234">Mindestversion des Postfachanforderungssatzes</span><span class="sxs-lookup"><span data-stu-id="d9554-234">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="d9554-235">1.1</span><span class="sxs-lookup"><span data-stu-id="d9554-235">1.1</span></span>|
|[<span data-ttu-id="d9554-236">Mindestberechtigungsstufe</span><span class="sxs-lookup"><span data-stu-id="d9554-236">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="d9554-237">ReadItem</span><span class="sxs-lookup"><span data-stu-id="d9554-237">ReadItem</span></span>|
|[<span data-ttu-id="d9554-238">Zutreffender Outlook-Modus</span><span class="sxs-lookup"><span data-stu-id="d9554-238">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="d9554-239">Verfassen</span><span class="sxs-lookup"><span data-stu-id="d9554-239">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="d9554-240">Beispiel</span><span class="sxs-lookup"><span data-stu-id="d9554-240">Example</span></span>

```
Office.context.mailbox.item.bcc.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.bcc.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.bcc.getAsync(callback);

function callback(asyncResult) {
  var arrayOfBccRecipients = asyncResult.value;
}
```

####  <a name="body-bodyjavascriptapioutlook17officebody"></a><span data-ttu-id="d9554-241">body :[Body](/javascript/api/outlook_1_7/office.body)</span><span class="sxs-lookup"><span data-stu-id="d9554-241">body :[Body](/javascript/api/outlook_1_7/office.body)</span></span>

<span data-ttu-id="d9554-242">Ruft ein Objekt ab, das Methoden zum Bearbeiten des Textkörpers eines Elements bereitstellt.</span><span class="sxs-lookup"><span data-stu-id="d9554-242">Gets an object that provides methods for manipulating the body of an item.</span></span>

##### <a name="type"></a><span data-ttu-id="d9554-243">Typ:</span><span class="sxs-lookup"><span data-stu-id="d9554-243">Type:</span></span>

*   [<span data-ttu-id="d9554-244">Body</span><span class="sxs-lookup"><span data-stu-id="d9554-244">Body</span></span>](/javascript/api/outlook_1_7/office.body)

##### <a name="requirements"></a><span data-ttu-id="d9554-245">Anforderungen</span><span class="sxs-lookup"><span data-stu-id="d9554-245">Requirements</span></span>

|<span data-ttu-id="d9554-246">Anforderung</span><span class="sxs-lookup"><span data-stu-id="d9554-246">Requirement</span></span>|<span data-ttu-id="d9554-247">Wert</span><span class="sxs-lookup"><span data-stu-id="d9554-247">Value</span></span>|
|---|---|
|[<span data-ttu-id="d9554-248">Mindestversion des Postfachanforderungssatzes</span><span class="sxs-lookup"><span data-stu-id="d9554-248">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="d9554-249">1.1</span><span class="sxs-lookup"><span data-stu-id="d9554-249">1.1</span></span>|
|[<span data-ttu-id="d9554-250">Mindestberechtigungsstufe</span><span class="sxs-lookup"><span data-stu-id="d9554-250">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="d9554-251">ReadItem</span><span class="sxs-lookup"><span data-stu-id="d9554-251">ReadItem</span></span>|
|[<span data-ttu-id="d9554-252">Zutreffender Outlook-Modus</span><span class="sxs-lookup"><span data-stu-id="d9554-252">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="d9554-253">Verfassen oder Lesen</span><span class="sxs-lookup"><span data-stu-id="d9554-253">Compose or read</span></span>|

####  <a name="cc-arrayemailaddressdetailsjavascriptapioutlook17officeemailaddressdetailsrecipientsjavascriptapioutlook17officerecipients"></a><span data-ttu-id="d9554-254">cc: Array. <[EmailAddressDetails](/javascript/api/outlook_1_7/office.emailaddressdetails)>|[Empfänger](/javascript/api/outlook_1_7/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="d9554-254">cc :Array.<[EmailAddressDetails](/javascript/api/outlook_1_7/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook_1_7/office.recipients)</span></span>

<span data-ttu-id="d9554-255">Ermöglicht den Zugriff auf die (Carbon Copy, Blindkopie) Cc-Empfänger einer Nachricht.</span><span class="sxs-lookup"><span data-stu-id="d9554-255">Provides access to the Cc (carbon copy) recipients of a message.</span></span> <span data-ttu-id="d9554-256">Der Typ des Objekts und die Zugriffsebene, hängt von den Modus des aktuellen Elements ab.</span><span class="sxs-lookup"><span data-stu-id="d9554-256">The type of object and level of access depends on the mode of the current item.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="d9554-257">Lesemodus</span><span class="sxs-lookup"><span data-stu-id="d9554-257">Read mode</span></span>

<span data-ttu-id="d9554-p106">Die `cc`-Eigenschaft gibt ein Array mit einem `EmailAddressDetails`-Objekt für jeden Empfänger in der **Cc**-Zeile der Nachricht zurück. Die Auflistung ist auf höchstens 100 Elemente beschränkt.</span><span class="sxs-lookup"><span data-stu-id="d9554-p106">The `cc` property returns an array that contains an `EmailAddressDetails` object for each recipient listed on the **Cc** line of the message. The collection is limited to a maximum of 100 members.</span></span>

##### <a name="compose-mode"></a><span data-ttu-id="d9554-260">Verfassenmodus</span><span class="sxs-lookup"><span data-stu-id="d9554-260">Compose mode</span></span>

<span data-ttu-id="d9554-261">Die `cc` -Eigenschaft gibt eine `Recipients` -Objekt, das Methoden zum Abrufen oder aktualisieren die Empfänger in der Zeile **Cc** der Nachricht bereitstellt.</span><span class="sxs-lookup"><span data-stu-id="d9554-261">The `cc` property returns a `Recipients` object that provides methods to get or update the recipients on the **Cc** line of the message.</span></span>

##### <a name="type"></a><span data-ttu-id="d9554-262">Typ:</span><span class="sxs-lookup"><span data-stu-id="d9554-262">Type:</span></span>

*   <span data-ttu-id="d9554-263">Array.<[EmailAddressDetails](/javascript/api/outlook_1_7/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook_1_7/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="d9554-263">Array.<[EmailAddressDetails](/javascript/api/outlook_1_7/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook_1_7/office.recipients)</span></span>

##### <a name="requirements"></a><span data-ttu-id="d9554-264">Anforderungen</span><span class="sxs-lookup"><span data-stu-id="d9554-264">Requirements</span></span>

|<span data-ttu-id="d9554-265">Anforderung</span><span class="sxs-lookup"><span data-stu-id="d9554-265">Requirement</span></span>|<span data-ttu-id="d9554-266">Wert</span><span class="sxs-lookup"><span data-stu-id="d9554-266">Value</span></span>|
|---|---|
|[<span data-ttu-id="d9554-267">Mindestversion des Postfachanforderungssatzes</span><span class="sxs-lookup"><span data-stu-id="d9554-267">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="d9554-268">1.0</span><span class="sxs-lookup"><span data-stu-id="d9554-268">1.0</span></span>|
|[<span data-ttu-id="d9554-269">Mindestberechtigungsstufe</span><span class="sxs-lookup"><span data-stu-id="d9554-269">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="d9554-270">ReadItem</span><span class="sxs-lookup"><span data-stu-id="d9554-270">ReadItem</span></span>|
|[<span data-ttu-id="d9554-271">Zutreffender Outlook-Modus</span><span class="sxs-lookup"><span data-stu-id="d9554-271">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="d9554-272">Verfassen oder Lesen</span><span class="sxs-lookup"><span data-stu-id="d9554-272">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="d9554-273">Beispiel</span><span class="sxs-lookup"><span data-stu-id="d9554-273">Example</span></span>

```
Office.context.mailbox.item.cc.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.cc.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.cc.getAsync(callback);

function callback(asyncResult) {
  var arrayOfCcRecipients = asyncResult.value;
}
```

####  <a name="nullable-conversationid-string"></a><span data-ttu-id="d9554-274">(nullable) conversationId :String</span><span class="sxs-lookup"><span data-stu-id="d9554-274">(nullable) conversationId :String</span></span>

<span data-ttu-id="d9554-275">Ruft einen Bezeichner für die E-Mail-Unterhaltung ab, in der eine bestimmte Nachricht enthalten ist.</span><span class="sxs-lookup"><span data-stu-id="d9554-275">Gets an identifier for the email conversation that contains a particular message.</span></span>

<span data-ttu-id="d9554-p107">Sie können für diese Eigenschaft eine ganze Zahl abrufen, wenn Ihre Mail-App in Formularen zum Lesen oder Antworten in Formularen zum Verfassen aktiviert wird. Wenn der Benutzer den Betreff der Antwortnachricht ändert, ändert sich beim Versenden die Konversations-ID für die entsprechende Nachricht, und der Wert, den Sie vorher bezogen haben, trifft nicht länger zu.</span><span class="sxs-lookup"><span data-stu-id="d9554-p107">You can get an integer for this property if your mail app is activated in read forms or responses in compose forms. If subsequently the user changes the subject of the reply message, upon sending the reply, the conversation ID for that message will change and that value you obtained earlier will no longer apply.</span></span>

<span data-ttu-id="d9554-p108">Sie erhalten in einem Formular zum Verfassen für diese Eigenschaft für ein neues Element null. Wenn der Benutzer einen Betreff festlegt und das Element speichert, gibt die `conversationId`-Eigenschaft einen Wert zurück.</span><span class="sxs-lookup"><span data-stu-id="d9554-p108">You get null for this property for a new item in a compose form. If the user sets a subject and saves the item, the `conversationId` property will return a value.</span></span>

##### <a name="type"></a><span data-ttu-id="d9554-280">Typ:</span><span class="sxs-lookup"><span data-stu-id="d9554-280">Type:</span></span>

*   <span data-ttu-id="d9554-281">String</span><span class="sxs-lookup"><span data-stu-id="d9554-281">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="d9554-282">Anforderungen</span><span class="sxs-lookup"><span data-stu-id="d9554-282">Requirements</span></span>

|<span data-ttu-id="d9554-283">Anforderung</span><span class="sxs-lookup"><span data-stu-id="d9554-283">Requirement</span></span>|<span data-ttu-id="d9554-284">Wert</span><span class="sxs-lookup"><span data-stu-id="d9554-284">Value</span></span>|
|---|---|
|[<span data-ttu-id="d9554-285">Mindestversion des Postfachanforderungssatzes</span><span class="sxs-lookup"><span data-stu-id="d9554-285">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="d9554-286">1.0</span><span class="sxs-lookup"><span data-stu-id="d9554-286">1.0</span></span>|
|[<span data-ttu-id="d9554-287">Mindestberechtigungsstufe</span><span class="sxs-lookup"><span data-stu-id="d9554-287">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="d9554-288">ReadItem</span><span class="sxs-lookup"><span data-stu-id="d9554-288">ReadItem</span></span>|
|[<span data-ttu-id="d9554-289">Zutreffender Outlook-Modus</span><span class="sxs-lookup"><span data-stu-id="d9554-289">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="d9554-290">Verfassen oder Lesen</span><span class="sxs-lookup"><span data-stu-id="d9554-290">Compose or read</span></span>|

#### <a name="datetimecreated-date"></a><span data-ttu-id="d9554-291">dateTimeCreated :Date</span><span class="sxs-lookup"><span data-stu-id="d9554-291">dateTimeCreated :Date</span></span>

<span data-ttu-id="d9554-p109">Ruft das Datum und die Uhrzeit der Erstellung eines Elements ab. Nur Lesemodus.</span><span class="sxs-lookup"><span data-stu-id="d9554-p109">Gets the date and time that an item was created. Read mode only.</span></span>

##### <a name="type"></a><span data-ttu-id="d9554-294">Typ:</span><span class="sxs-lookup"><span data-stu-id="d9554-294">Type:</span></span>

*   <span data-ttu-id="d9554-295">Datum</span><span class="sxs-lookup"><span data-stu-id="d9554-295">Date</span></span>

##### <a name="requirements"></a><span data-ttu-id="d9554-296">Anforderungen</span><span class="sxs-lookup"><span data-stu-id="d9554-296">Requirements</span></span>

|<span data-ttu-id="d9554-297">Anforderung</span><span class="sxs-lookup"><span data-stu-id="d9554-297">Requirement</span></span>|<span data-ttu-id="d9554-298">Wert</span><span class="sxs-lookup"><span data-stu-id="d9554-298">Value</span></span>|
|---|---|
|[<span data-ttu-id="d9554-299">Mindestversion des Postfachanforderungssatzes</span><span class="sxs-lookup"><span data-stu-id="d9554-299">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="d9554-300">1.0</span><span class="sxs-lookup"><span data-stu-id="d9554-300">1.0</span></span>|
|[<span data-ttu-id="d9554-301">Mindestberechtigungsstufe</span><span class="sxs-lookup"><span data-stu-id="d9554-301">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="d9554-302">ReadItem</span><span class="sxs-lookup"><span data-stu-id="d9554-302">ReadItem</span></span>|
|[<span data-ttu-id="d9554-303">Zutreffender Outlook-Modus</span><span class="sxs-lookup"><span data-stu-id="d9554-303">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="d9554-304">Lesen</span><span class="sxs-lookup"><span data-stu-id="d9554-304">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="d9554-305">Beispiel</span><span class="sxs-lookup"><span data-stu-id="d9554-305">Example</span></span>

```
var created = Office.context.mailbox.item.dateTimeCreated;
```

#### <a name="datetimemodified-date"></a><span data-ttu-id="d9554-306">dateTimeModified :Date</span><span class="sxs-lookup"><span data-stu-id="d9554-306">dateTimeModified :Date</span></span>

<span data-ttu-id="d9554-p110">Ruft das Datum und die Uhrzeit der letzten Änderung eines Elements ab. Nur Lesemodus.</span><span class="sxs-lookup"><span data-stu-id="d9554-p110">Gets the date and time that an item was last modified. Read mode only.</span></span>

> [!NOTE]
> <span data-ttu-id="d9554-309">Dieser Member wird in Outlook für iOS oder Outlook für Android nicht unterstützt.</span><span class="sxs-lookup"><span data-stu-id="d9554-309">This member is not supported in Outlook for iOS or Outlook for Android.</span></span>

##### <a name="type"></a><span data-ttu-id="d9554-310">Typ:</span><span class="sxs-lookup"><span data-stu-id="d9554-310">Type:</span></span>

*   <span data-ttu-id="d9554-311">Datum</span><span class="sxs-lookup"><span data-stu-id="d9554-311">Date</span></span>

##### <a name="requirements"></a><span data-ttu-id="d9554-312">Anforderungen</span><span class="sxs-lookup"><span data-stu-id="d9554-312">Requirements</span></span>

|<span data-ttu-id="d9554-313">Anforderung</span><span class="sxs-lookup"><span data-stu-id="d9554-313">Requirement</span></span>|<span data-ttu-id="d9554-314">Wert</span><span class="sxs-lookup"><span data-stu-id="d9554-314">Value</span></span>|
|---|---|
|[<span data-ttu-id="d9554-315">Mindestversion des Postfachanforderungssatzes</span><span class="sxs-lookup"><span data-stu-id="d9554-315">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="d9554-316">1.0</span><span class="sxs-lookup"><span data-stu-id="d9554-316">1.0</span></span>|
|[<span data-ttu-id="d9554-317">Mindestberechtigungsstufe</span><span class="sxs-lookup"><span data-stu-id="d9554-317">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="d9554-318">ReadItem</span><span class="sxs-lookup"><span data-stu-id="d9554-318">ReadItem</span></span>|
|[<span data-ttu-id="d9554-319">Zutreffender Outlook-Modus</span><span class="sxs-lookup"><span data-stu-id="d9554-319">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="d9554-320">Lesen</span><span class="sxs-lookup"><span data-stu-id="d9554-320">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="d9554-321">Beispiel</span><span class="sxs-lookup"><span data-stu-id="d9554-321">Example</span></span>

```
var modified = Office.context.mailbox.item.dateTimeModified;
```

####  <a name="end-datetimejavascriptapioutlook17officetime"></a><span data-ttu-id="d9554-322">end :Date|[Time](/javascript/api/outlook_1_7/office.time)</span><span class="sxs-lookup"><span data-stu-id="d9554-322">end :Date|[Time](/javascript/api/outlook_1_7/office.time)</span></span>

<span data-ttu-id="d9554-323">Ruft Datum und Zeit für das Ende des Termins ab oder legt diese fest.</span><span class="sxs-lookup"><span data-stu-id="d9554-323">Gets or sets the date and time that the appointment is to end.</span></span>

<span data-ttu-id="d9554-p111">Die `end`-Eigenschaft wird als Datums- und Uhrzeitwert in UTC ausgedrückt. Sie können die [`convertToLocalClientTime`](office.context.mailbox.md#converttolocalclienttimetimevalue--localclienttimejavascriptapioutlook17officelocalclienttime)-Methode verwenden, um den Endeigenschaftswert in den lokalen Uhrzeit- und Datumswert des Clients umzuwandeln.</span><span class="sxs-lookup"><span data-stu-id="d9554-p111">The `end` property is expressed as a Coordinated Universal Time (UTC) date and time value. You can use the [`convertToLocalClientTime`](office.context.mailbox.md#converttolocalclienttimetimevalue--localclienttimejavascriptapioutlook17officelocalclienttime) method to convert the end property value to the client’s local date and time.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="d9554-326">Lesemodus</span><span class="sxs-lookup"><span data-stu-id="d9554-326">Read mode</span></span>

<span data-ttu-id="d9554-327">Die `end`-Eigenschaft gibt ein `Date`-Objekt zurück.</span><span class="sxs-lookup"><span data-stu-id="d9554-327">The `end` property returns a `Date` object.</span></span>

##### <a name="compose-mode"></a><span data-ttu-id="d9554-328">Verfassenmodus</span><span class="sxs-lookup"><span data-stu-id="d9554-328">Compose mode</span></span>

<span data-ttu-id="d9554-329">Die `end`-Eigenschaft gibt ein `Time`-Objekt zurück.</span><span class="sxs-lookup"><span data-stu-id="d9554-329">The `end` property returns a `Time` object.</span></span>

<span data-ttu-id="d9554-330">Wenn Sie die [`Time.setAsync`](/javascript/api/outlook_1_7/office.time#setasync-datetime--options--callback-)-Methode verwenden, um die Endzeit im Verfassenmodus festzulegen, sollten Sie die [`convertToUtcClientTime`](office.context.mailbox.md#converttoutcclienttimeinput--date)-Methode verwenden, um die Ortszeit auf dem Client für den Server in UTC umzuwandeln.</span><span class="sxs-lookup"><span data-stu-id="d9554-330">When you use the [`Time.setAsync`](/javascript/api/outlook_1_7/office.time#setasync-datetime--options--callback-) method to set the end time, you should use the [`convertToUtcClientTime`](office.context.mailbox.md#converttoutcclienttimeinput--date) method to convert the local time on the client to UTC for the server.</span></span>

##### <a name="type"></a><span data-ttu-id="d9554-331">Typ:</span><span class="sxs-lookup"><span data-stu-id="d9554-331">Type:</span></span>

*   <span data-ttu-id="d9554-332">Date | [Time](/javascript/api/outlook_1_7/office.time)</span><span class="sxs-lookup"><span data-stu-id="d9554-332">Date | [Time](/javascript/api/outlook_1_7/office.time)</span></span>

##### <a name="requirements"></a><span data-ttu-id="d9554-333">Anforderungen</span><span class="sxs-lookup"><span data-stu-id="d9554-333">Requirements</span></span>

|<span data-ttu-id="d9554-334">Anforderung</span><span class="sxs-lookup"><span data-stu-id="d9554-334">Requirement</span></span>|<span data-ttu-id="d9554-335">Wert</span><span class="sxs-lookup"><span data-stu-id="d9554-335">Value</span></span>|
|---|---|
|[<span data-ttu-id="d9554-336">Mindestversion des Postfachanforderungssatzes</span><span class="sxs-lookup"><span data-stu-id="d9554-336">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="d9554-337">1.0</span><span class="sxs-lookup"><span data-stu-id="d9554-337">1.0</span></span>|
|[<span data-ttu-id="d9554-338">Mindestberechtigungsstufe</span><span class="sxs-lookup"><span data-stu-id="d9554-338">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="d9554-339">ReadItem</span><span class="sxs-lookup"><span data-stu-id="d9554-339">ReadItem</span></span>|
|[<span data-ttu-id="d9554-340">Zutreffender Outlook-Modus</span><span class="sxs-lookup"><span data-stu-id="d9554-340">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="d9554-341">Verfassen oder Lesen</span><span class="sxs-lookup"><span data-stu-id="d9554-341">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="d9554-342">Beispiel</span><span class="sxs-lookup"><span data-stu-id="d9554-342">Example</span></span>

<span data-ttu-id="d9554-343">Im folgenden Beispiel wird die Endzeit eines Termins im Verfassenmodus mithilfe der [`setAsync`](/javascript/api/outlook_1_7/office.time#setasync-datetime--options--callback-)-Methode des `Time`-Objekts festgelegt.</span><span class="sxs-lookup"><span data-stu-id="d9554-343">The following example sets the end time of an appointment in compose mode by using the [`setAsync`](/javascript/api/outlook_1_7/office.time#setasync-datetime--options--callback-) method of the `Time` object.</span></span>

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

#### <a name="from-emailaddressdetailsjavascriptapioutlook17officeemailaddressdetailsfromjavascriptapioutlook17officefrom"></a><span data-ttu-id="d9554-344">von:[EmailAddressDetails](/javascript/api/outlook_1_7/office.emailaddressdetails)|[aus](/javascript/api/outlook_1_7/office.from)</span><span class="sxs-lookup"><span data-stu-id="d9554-344">from :[EmailAddressDetails](/javascript/api/outlook_1_7/office.emailaddressdetails)|[From](/javascript/api/outlook_1_7/office.from)</span></span>

<span data-ttu-id="d9554-345">Ruft die E-Mail-Adresse des Absenders einer Nachricht ab.</span><span class="sxs-lookup"><span data-stu-id="d9554-345">Gets the email address of the sender of a message.</span></span>

<span data-ttu-id="d9554-p112">Die `from`- und [`sender`](#sender-emailaddressdetailsjavascriptapioutlook17officeemailaddressdetails)-Eigenschaften stellen dieselbe Person dar, außer die Nachricht wurde von einem Delegaten gesendet. In diesem Fall stellt die `from`-Eigenschaft die Stellvertretung dar, und die sender-Eigenschaft stellt den Delegaten dar.</span><span class="sxs-lookup"><span data-stu-id="d9554-p112">The `from` and [`sender`](#sender-emailaddressdetailsjavascriptapioutlook17officeemailaddressdetails) properties represent the same person unless the message is sent by a delegate. In that case, the `from` property represents the delegator, and the sender property represents the delegate.</span></span>

> [!NOTE]
> <span data-ttu-id="d9554-348">Die `recipientType` -Eigenschaft des der `EmailAddressDetails` -Objekts der `from` -Eigenschaft ist `undefined`.</span><span class="sxs-lookup"><span data-stu-id="d9554-348">The `recipientType` property of the `EmailAddressDetails` object in the `from` property is `undefined`.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="d9554-349">Lesemodus</span><span class="sxs-lookup"><span data-stu-id="d9554-349">Read mode</span></span>

<span data-ttu-id="d9554-350">Die `from` -Eigenschaft gibt eine `EmailAddressDetails` Objekt.</span><span class="sxs-lookup"><span data-stu-id="d9554-350">The `from` property returns an `EmailAddressDetails` object.</span></span>

```
var subject = Office.context.mailbox.item.from;
```

##### <a name="compose-mode"></a><span data-ttu-id="d9554-351">Verfassenmodus</span><span class="sxs-lookup"><span data-stu-id="d9554-351">Compose mode</span></span>

<span data-ttu-id="d9554-352">Die `from` -Eigenschaft gibt eine `From` -Objekt, eine Methode zum Abrufen bietet, der vom Wert.</span><span class="sxs-lookup"><span data-stu-id="d9554-352">The `from` property returns a `From` object that provides a method to get the from value.</span></span>

```
Office.context.mailbox.item.from.getAsync(callback);

function callback(asyncResult) {
  var from = asyncResult.value;
}
```

##### <a name="type"></a><span data-ttu-id="d9554-353">Typ:</span><span class="sxs-lookup"><span data-stu-id="d9554-353">Type:</span></span>

*   <span data-ttu-id="d9554-354">[EmailAddressDetails](/javascript/api/outlook_1_7/office.emailaddressdetails) | [aus](/javascript/api/outlook_1_7/office.from)</span><span class="sxs-lookup"><span data-stu-id="d9554-354">[EmailAddressDetails](/javascript/api/outlook_1_7/office.emailaddressdetails) | [From](/javascript/api/outlook_1_7/office.from)</span></span>

##### <a name="requirements"></a><span data-ttu-id="d9554-355">Anforderungen</span><span class="sxs-lookup"><span data-stu-id="d9554-355">Requirements</span></span>

|<span data-ttu-id="d9554-356">Anforderung</span><span class="sxs-lookup"><span data-stu-id="d9554-356">Requirement</span></span>|||
|---|---|---|
|[<span data-ttu-id="d9554-357">Mindestversion des Postfachanforderungssatzes</span><span class="sxs-lookup"><span data-stu-id="d9554-357">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="d9554-358">1.0</span><span class="sxs-lookup"><span data-stu-id="d9554-358">1.0</span></span>|<span data-ttu-id="d9554-359">1.7</span><span class="sxs-lookup"><span data-stu-id="d9554-359">1.7</span></span>|
|[<span data-ttu-id="d9554-360">Mindestberechtigungsstufe</span><span class="sxs-lookup"><span data-stu-id="d9554-360">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="d9554-361">ReadItem</span><span class="sxs-lookup"><span data-stu-id="d9554-361">ReadItem</span></span>|<span data-ttu-id="d9554-362">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="d9554-362">ReadWriteItem</span></span>|
|[<span data-ttu-id="d9554-363">Zutreffender Outlook-Modus</span><span class="sxs-lookup"><span data-stu-id="d9554-363">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="d9554-364">Lesen</span><span class="sxs-lookup"><span data-stu-id="d9554-364">Read</span></span>|<span data-ttu-id="d9554-365">Verfassen</span><span class="sxs-lookup"><span data-stu-id="d9554-365">Compose</span></span>|

#### <a name="internetmessageid-string"></a><span data-ttu-id="d9554-366">internetMessageId :String</span><span class="sxs-lookup"><span data-stu-id="d9554-366">internetMessageId :String</span></span>

<span data-ttu-id="d9554-p113">Ruft die Internetnachricht-ID einer E-Mail-Nachricht ab. Nur Lesemodus.</span><span class="sxs-lookup"><span data-stu-id="d9554-p113">Gets the Internet message identifier for an email message. Read mode only.</span></span>

##### <a name="type"></a><span data-ttu-id="d9554-369">Typ:</span><span class="sxs-lookup"><span data-stu-id="d9554-369">Type:</span></span>

*   <span data-ttu-id="d9554-370">String</span><span class="sxs-lookup"><span data-stu-id="d9554-370">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="d9554-371">Anforderungen</span><span class="sxs-lookup"><span data-stu-id="d9554-371">Requirements</span></span>

|<span data-ttu-id="d9554-372">Anforderung</span><span class="sxs-lookup"><span data-stu-id="d9554-372">Requirement</span></span>|<span data-ttu-id="d9554-373">Wert</span><span class="sxs-lookup"><span data-stu-id="d9554-373">Value</span></span>|
|---|---|
|[<span data-ttu-id="d9554-374">Mindestversion des Postfachanforderungssatzes</span><span class="sxs-lookup"><span data-stu-id="d9554-374">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="d9554-375">1.0</span><span class="sxs-lookup"><span data-stu-id="d9554-375">1.0</span></span>|
|[<span data-ttu-id="d9554-376">Mindestberechtigungsstufe</span><span class="sxs-lookup"><span data-stu-id="d9554-376">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="d9554-377">ReadItem</span><span class="sxs-lookup"><span data-stu-id="d9554-377">ReadItem</span></span>|
|[<span data-ttu-id="d9554-378">Zutreffender Outlook-Modus</span><span class="sxs-lookup"><span data-stu-id="d9554-378">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="d9554-379">Lesen</span><span class="sxs-lookup"><span data-stu-id="d9554-379">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="d9554-380">Beispiel</span><span class="sxs-lookup"><span data-stu-id="d9554-380">Example</span></span>

```
var internetMessageId = Office.context.mailbox.item.internetMessageId;
```

#### <a name="itemclass-string"></a><span data-ttu-id="d9554-381">itemClass :String</span><span class="sxs-lookup"><span data-stu-id="d9554-381">itemClass :String</span></span>

<span data-ttu-id="d9554-p114">Ruft die Exchange-Webdienste-Elementklasse des ausgewählten Elements ab. Nur Lesemodus.</span><span class="sxs-lookup"><span data-stu-id="d9554-p114">Gets the Exchange Web Services item class of the selected item. Read mode only.</span></span>

<span data-ttu-id="d9554-p115">Die `itemClass`-Eigenschaft gibt die Nachrichtenklasse des ausgewählten Elements an. Folgende Standardnachrichtenklassen für das Nachrichten- oder Terminelement sind vorhanden:</span><span class="sxs-lookup"><span data-stu-id="d9554-p115">The `itemClass` property specifies the message class of the selected item. The following are the default message classes for the message or appointment item.</span></span>

|<span data-ttu-id="d9554-386">Typ</span><span class="sxs-lookup"><span data-stu-id="d9554-386">Type</span></span>|<span data-ttu-id="d9554-387">Beschreibung</span><span class="sxs-lookup"><span data-stu-id="d9554-387">Description</span></span>|<span data-ttu-id="d9554-388">Elementklasse</span><span class="sxs-lookup"><span data-stu-id="d9554-388">item class</span></span>|
|---|---|---|
|<span data-ttu-id="d9554-389">Terminelemente</span><span class="sxs-lookup"><span data-stu-id="d9554-389">Appointment items</span></span>|<span data-ttu-id="d9554-390">Dies sind Kalenderelemente der Elementklasse `IPM.Appointment` oder `IPM.Appointment.Occurence`.</span><span class="sxs-lookup"><span data-stu-id="d9554-390">These are calendar items of the item class `IPM.Appointment` or `IPM.Appointment.Occurence`.</span></span>|`IPM.Appointment`<br />`IPM.Appointment.Occurence`|
|<span data-ttu-id="d9554-391">Nachrichtenelemente</span><span class="sxs-lookup"><span data-stu-id="d9554-391">Message items</span></span>|<span data-ttu-id="d9554-392">Diese Elemente umfassen E-Mail-Nachrichten der Standardnachrichtenklasse `IPM.Note` sowie Besprechungsanfragen, -antworten und -absagen, die `IPM.Schedule.Meeting` als Basisnachrichtenklasse verwenden.</span><span class="sxs-lookup"><span data-stu-id="d9554-392">These include email messages that have the default message class `IPM.Note`, and meeting requests, responses, and cancellations, that use `IPM.Schedule.Meeting` as the base message class.</span></span>|`IPM.Note`<br />`IPM.Schedule.Meeting.Request`<br />`IPM.Schedule.Meeting.Neg`<br />`IPM.Schedule.Meeting.Pos`<br />`IPM.Schedule.Meeting.Tent`<br />`IPM.Schedule.Meeting.Canceled`|

<span data-ttu-id="d9554-393">Sie können benutzerdefinierte Nachrichtenklassen zum Erweitern einer Standardnachrichtenklasse erstellen, z. B. eine benutzerdefinierte Terminnachrichtenklasse wie `IPM.Appointment.Contoso`.</span><span class="sxs-lookup"><span data-stu-id="d9554-393">You can create custom message classes that extends a default message class, for example, a custom appointment message class `IPM.Appointment.Contoso`.</span></span>

##### <a name="type"></a><span data-ttu-id="d9554-394">Typ:</span><span class="sxs-lookup"><span data-stu-id="d9554-394">Type:</span></span>

*   <span data-ttu-id="d9554-395">String</span><span class="sxs-lookup"><span data-stu-id="d9554-395">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="d9554-396">Anforderungen</span><span class="sxs-lookup"><span data-stu-id="d9554-396">Requirements</span></span>

|<span data-ttu-id="d9554-397">Anforderung</span><span class="sxs-lookup"><span data-stu-id="d9554-397">Requirement</span></span>|<span data-ttu-id="d9554-398">Wert</span><span class="sxs-lookup"><span data-stu-id="d9554-398">Value</span></span>|
|---|---|
|[<span data-ttu-id="d9554-399">Mindestversion des Postfachanforderungssatzes</span><span class="sxs-lookup"><span data-stu-id="d9554-399">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="d9554-400">1.0</span><span class="sxs-lookup"><span data-stu-id="d9554-400">1.0</span></span>|
|[<span data-ttu-id="d9554-401">Mindestberechtigungsstufe</span><span class="sxs-lookup"><span data-stu-id="d9554-401">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="d9554-402">ReadItem</span><span class="sxs-lookup"><span data-stu-id="d9554-402">ReadItem</span></span>|
|[<span data-ttu-id="d9554-403">Zutreffender Outlook-Modus</span><span class="sxs-lookup"><span data-stu-id="d9554-403">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="d9554-404">Lesen</span><span class="sxs-lookup"><span data-stu-id="d9554-404">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="d9554-405">Beispiel</span><span class="sxs-lookup"><span data-stu-id="d9554-405">Example</span></span>

```
var itemClass = Office.context.mailbox.item.itemClass;
```

#### <a name="nullable-itemid-string"></a><span data-ttu-id="d9554-406">(nullable) itemId :String</span><span class="sxs-lookup"><span data-stu-id="d9554-406">(nullable) itemId :String</span></span>

<span data-ttu-id="d9554-p116">Ruft den Exchange-Webdienste-Elementbezeichner für das aktuelle Element ab. Nur Lesemodus.</span><span class="sxs-lookup"><span data-stu-id="d9554-p116">Gets the Exchange Web Services item identifier for the current item. Read mode only.</span></span>

> [!NOTE]
> <span data-ttu-id="d9554-409">Der Bezeichner, der von der `itemId`-Eigenschaft zurückgegeben wird, ist der gleiche Bezeichner wie der Exchange-Webdienste-Elementbezeichner.</span><span class="sxs-lookup"><span data-stu-id="d9554-409">The identifier returned by the `itemId` property is the same as the Exchange Web Services item identifier.</span></span> <span data-ttu-id="d9554-410">Die `itemId` -Eigenschaft ist nicht identisch mit der Outlook-Eintrags-ID oder die ID von der Outlook-REST-API verwendet.</span><span class="sxs-lookup"><span data-stu-id="d9554-410">The `itemId` property is not identical to the Outlook Entry ID or the ID used by the Outlook REST API.</span></span> <span data-ttu-id="d9554-411">REST-API-Aufrufe dieser Wert verwenden, sollten sie konvertiert werden, bevor [Office.context.mailbox.convertToRestId](office.context.mailbox.md#converttorestiditemid-restversion--string)verwenden.</span><span class="sxs-lookup"><span data-stu-id="d9554-411">Before making REST API calls using this value, it should be converted using [Office.context.mailbox.convertToRestId](office.context.mailbox.md#converttorestiditemid-restversion--string).</span></span> <span data-ttu-id="d9554-412">Weitere Informationen finden Sie unter [Verwenden der Outlook-REST-APIs aus einem Outlook-add-in](https://docs.microsoft.com/outlook/add-ins/use-rest-api#get-the-item-id).</span><span class="sxs-lookup"><span data-stu-id="d9554-412">For more details, see [Use the Outlook REST APIs from an Outlook add-in](https://docs.microsoft.com/outlook/add-ins/use-rest-api#get-the-item-id).</span></span>

<span data-ttu-id="d9554-p118">Die Eigenschaft `itemId` ist im Modus „Verfassen“ nicht verfügbar. Wenn ein Elementbezeichner erforderlich ist, kann die [`saveAsync`](#saveasyncoptions-callback)-Methode verwendet werden, um das Element zu speichern. Dadurch wird der Elementbezeichner im [`AsyncResult.value`](/javascript/api/office/office.asyncresult)-Parameter in der Rückruffunktion zurückgegeben.</span><span class="sxs-lookup"><span data-stu-id="d9554-p118">The `itemId` property is not available in compose mode. If an item identifier is required, the [`saveAsync`](#saveasyncoptions-callback) method can be used to save the item to the store, which will return the item identifier in the [`AsyncResult.value`](/javascript/api/office/office.asyncresult) parameter in the callback function.</span></span>

##### <a name="type"></a><span data-ttu-id="d9554-415">Typ:</span><span class="sxs-lookup"><span data-stu-id="d9554-415">Type:</span></span>

*   <span data-ttu-id="d9554-416">String</span><span class="sxs-lookup"><span data-stu-id="d9554-416">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="d9554-417">Anforderungen</span><span class="sxs-lookup"><span data-stu-id="d9554-417">Requirements</span></span>

|<span data-ttu-id="d9554-418">Anforderung</span><span class="sxs-lookup"><span data-stu-id="d9554-418">Requirement</span></span>|<span data-ttu-id="d9554-419">Wert</span><span class="sxs-lookup"><span data-stu-id="d9554-419">Value</span></span>|
|---|---|
|[<span data-ttu-id="d9554-420">Mindestversion des Postfachanforderungssatzes</span><span class="sxs-lookup"><span data-stu-id="d9554-420">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="d9554-421">1.0</span><span class="sxs-lookup"><span data-stu-id="d9554-421">1.0</span></span>|
|[<span data-ttu-id="d9554-422">Mindestberechtigungsstufe</span><span class="sxs-lookup"><span data-stu-id="d9554-422">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="d9554-423">ReadItem</span><span class="sxs-lookup"><span data-stu-id="d9554-423">ReadItem</span></span>|
|[<span data-ttu-id="d9554-424">Zutreffender Outlook-Modus</span><span class="sxs-lookup"><span data-stu-id="d9554-424">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="d9554-425">Lesen</span><span class="sxs-lookup"><span data-stu-id="d9554-425">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="d9554-426">Beispiel</span><span class="sxs-lookup"><span data-stu-id="d9554-426">Example</span></span>

<span data-ttu-id="d9554-p119">Mit dem folgende Code wird das Vorhandensein eines Elementbezeichners überprüft. Wenn die `itemId`-Eigenschaft `null` oder `undefined` zurückgibt, speichert sie das Element und ruft den Elementbezeichner aus dem asynchronen Ergebnis ab.</span><span class="sxs-lookup"><span data-stu-id="d9554-p119">The following code checks for the presence of an item identifier. If the `itemId` property returns `null` or `undefined`, it saves the item to the store and gets the item identifier from the asynchronous result.</span></span>

```
var itemId = Office.context.mailbox.item.itemId;
if (itemId === null || itemId == undefined) {
  Office.context.mailbox.item.saveAsync(function(result){
    itemId = result.value;
  });
}
```

####  <a name="itemtype-officemailboxenumsitemtypejavascriptapioutlook17officemailboxenumsitemtype"></a><span data-ttu-id="d9554-429">itemType :[Office.MailboxEnums.ItemType](/javascript/api/outlook_1_7/office.mailboxenums.itemtype)</span><span class="sxs-lookup"><span data-stu-id="d9554-429">itemType :[Office.MailboxEnums.ItemType](/javascript/api/outlook_1_7/office.mailboxenums.itemtype)</span></span>

<span data-ttu-id="d9554-430">Ruft den Typ des Elements ab, das eine Instanz darstellt.</span><span class="sxs-lookup"><span data-stu-id="d9554-430">Gets the type of item that an instance represents.</span></span>

<span data-ttu-id="d9554-431">Die `itemType`-Eigenschaft gibt einen der Werte der `ItemType`-Enumeration zurück, der angibt, ob es sich bei der `item`-Objektinstanz um eine Nachricht oder einen Termin handelt.</span><span class="sxs-lookup"><span data-stu-id="d9554-431">The `itemType` property returns one of the `ItemType` enumeration values, indicating whether the `item` object instance is a message or an appointment.</span></span>

##### <a name="type"></a><span data-ttu-id="d9554-432">Typ:</span><span class="sxs-lookup"><span data-stu-id="d9554-432">Type:</span></span>

*   [<span data-ttu-id="d9554-433">Office.MailboxEnums.ItemType</span><span class="sxs-lookup"><span data-stu-id="d9554-433">Office.MailboxEnums.ItemType</span></span>](/javascript/api/outlook_1_7/office.mailboxenums.itemtype)

##### <a name="requirements"></a><span data-ttu-id="d9554-434">Anforderungen</span><span class="sxs-lookup"><span data-stu-id="d9554-434">Requirements</span></span>

|<span data-ttu-id="d9554-435">Anforderung</span><span class="sxs-lookup"><span data-stu-id="d9554-435">Requirement</span></span>|<span data-ttu-id="d9554-436">Wert</span><span class="sxs-lookup"><span data-stu-id="d9554-436">Value</span></span>|
|---|---|
|[<span data-ttu-id="d9554-437">Mindestversion des Postfachanforderungssatzes</span><span class="sxs-lookup"><span data-stu-id="d9554-437">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="d9554-438">1.0</span><span class="sxs-lookup"><span data-stu-id="d9554-438">1.0</span></span>|
|[<span data-ttu-id="d9554-439">Mindestberechtigungsstufe</span><span class="sxs-lookup"><span data-stu-id="d9554-439">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="d9554-440">ReadItem</span><span class="sxs-lookup"><span data-stu-id="d9554-440">ReadItem</span></span>|
|[<span data-ttu-id="d9554-441">Zutreffender Outlook-Modus</span><span class="sxs-lookup"><span data-stu-id="d9554-441">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="d9554-442">Verfassen oder Lesen</span><span class="sxs-lookup"><span data-stu-id="d9554-442">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="d9554-443">Beispiel</span><span class="sxs-lookup"><span data-stu-id="d9554-443">Example</span></span>

```
if (Office.context.mailbox.item.itemType == Office.MailboxEnums.ItemType.Message)
  // do something
else
  // do something else
```

####  <a name="location-stringlocationjavascriptapioutlook17officelocation"></a><span data-ttu-id="d9554-444">location :String|[Location](/javascript/api/outlook_1_7/office.location)</span><span class="sxs-lookup"><span data-stu-id="d9554-444">location :String|[Location](/javascript/api/outlook_1_7/office.location)</span></span>

<span data-ttu-id="d9554-445">Ruft den Ort eines Termins ab bzw. legt ihn fest.</span><span class="sxs-lookup"><span data-stu-id="d9554-445">Gets or sets the location of an appointment.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="d9554-446">Lesemodus</span><span class="sxs-lookup"><span data-stu-id="d9554-446">Read mode</span></span>

<span data-ttu-id="d9554-447">Die `location`-Eigenschaft gibt eine Zeichenfolge zurück, die den Ort des Termins enthält.</span><span class="sxs-lookup"><span data-stu-id="d9554-447">The `location` property returns a string that contains the location of the appointment.</span></span>

##### <a name="compose-mode"></a><span data-ttu-id="d9554-448">Verfassenmodus</span><span class="sxs-lookup"><span data-stu-id="d9554-448">Compose mode</span></span>

<span data-ttu-id="d9554-449">Die `location`-Eigenschaft gibt ein `Location`-Objekt zurück, das Methoden zum Abrufen und Festlegen des Orts für einen Termin bereitstellt.</span><span class="sxs-lookup"><span data-stu-id="d9554-449">The `location` property returns a `Location` object that provides methods that are used to get and set the location of the appointment.</span></span>

##### <a name="type"></a><span data-ttu-id="d9554-450">Typ:</span><span class="sxs-lookup"><span data-stu-id="d9554-450">Type:</span></span>

*   <span data-ttu-id="d9554-451">String | [Location](/javascript/api/outlook_1_7/office.location)</span><span class="sxs-lookup"><span data-stu-id="d9554-451">String | [Location](/javascript/api/outlook_1_7/office.location)</span></span>

##### <a name="requirements"></a><span data-ttu-id="d9554-452">Anforderungen</span><span class="sxs-lookup"><span data-stu-id="d9554-452">Requirements</span></span>

|<span data-ttu-id="d9554-453">Anforderung</span><span class="sxs-lookup"><span data-stu-id="d9554-453">Requirement</span></span>|<span data-ttu-id="d9554-454">Wert</span><span class="sxs-lookup"><span data-stu-id="d9554-454">Value</span></span>|
|---|---|
|[<span data-ttu-id="d9554-455">Mindestversion des Postfachanforderungssatzes</span><span class="sxs-lookup"><span data-stu-id="d9554-455">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="d9554-456">1.0</span><span class="sxs-lookup"><span data-stu-id="d9554-456">1.0</span></span>|
|[<span data-ttu-id="d9554-457">Mindestberechtigungsstufe</span><span class="sxs-lookup"><span data-stu-id="d9554-457">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="d9554-458">ReadItem</span><span class="sxs-lookup"><span data-stu-id="d9554-458">ReadItem</span></span>|
|[<span data-ttu-id="d9554-459">Zutreffender Outlook-Modus</span><span class="sxs-lookup"><span data-stu-id="d9554-459">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="d9554-460">Verfassen oder Lesen</span><span class="sxs-lookup"><span data-stu-id="d9554-460">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="d9554-461">Beispiel</span><span class="sxs-lookup"><span data-stu-id="d9554-461">Example</span></span>

```
var userContext = { value : 1 };
Office.context.mailbox.item.location.getAsync( { context: userContext}, callback);

function callback(asyncResult) {
  var context = asyncResult.context;
  var location = asyncResult.value;
}
```

#### <a name="normalizedsubject-string"></a><span data-ttu-id="d9554-462">normalizedSubject :String</span><span class="sxs-lookup"><span data-stu-id="d9554-462">normalizedSubject :String</span></span>

<span data-ttu-id="d9554-p120">Ruft den Betreff für ein Element ab, wobei alle Präfixe entfernt werden (einschließlich `RE:` und `FWD:`). Nur Lesemodus.</span><span class="sxs-lookup"><span data-stu-id="d9554-p120">Gets the subject of an item, with all prefixes removed (including `RE:` and `FWD:`). Read mode only.</span></span>

<span data-ttu-id="d9554-p121">Die normalizedSubject-Eigenschaft ruft den Betreff des Elements mit allen Standardpräfixen (z. B. `RE:` und `FW:`) ab, die von E-Mail-Programmen hinzugefügt werden. Wenn der Betreff des Elements mit vollständigen Präfixen abgerufen werden soll, verwenden Sie die [`subject`](#subject-stringsubjectjavascriptapioutlook17officesubject)-Eigenschaft.</span><span class="sxs-lookup"><span data-stu-id="d9554-p121">The normalizedSubject property gets the subject of the item, with any standard prefixes (such as `RE:` and `FW:`) that are added by email programs. To get the subject of the item with the prefixes intact, use the [`subject`](#subject-stringsubjectjavascriptapioutlook17officesubject) property.</span></span>

##### <a name="type"></a><span data-ttu-id="d9554-467">Typ:</span><span class="sxs-lookup"><span data-stu-id="d9554-467">Type:</span></span>

*   <span data-ttu-id="d9554-468">String</span><span class="sxs-lookup"><span data-stu-id="d9554-468">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="d9554-469">Anforderungen</span><span class="sxs-lookup"><span data-stu-id="d9554-469">Requirements</span></span>

|<span data-ttu-id="d9554-470">Anforderung</span><span class="sxs-lookup"><span data-stu-id="d9554-470">Requirement</span></span>|<span data-ttu-id="d9554-471">Wert</span><span class="sxs-lookup"><span data-stu-id="d9554-471">Value</span></span>|
|---|---|
|[<span data-ttu-id="d9554-472">Mindestversion des Postfachanforderungssatzes</span><span class="sxs-lookup"><span data-stu-id="d9554-472">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="d9554-473">1.0</span><span class="sxs-lookup"><span data-stu-id="d9554-473">1.0</span></span>|
|[<span data-ttu-id="d9554-474">Mindestberechtigungsstufe</span><span class="sxs-lookup"><span data-stu-id="d9554-474">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="d9554-475">ReadItem</span><span class="sxs-lookup"><span data-stu-id="d9554-475">ReadItem</span></span>|
|[<span data-ttu-id="d9554-476">Zutreffender Outlook-Modus</span><span class="sxs-lookup"><span data-stu-id="d9554-476">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="d9554-477">Lesen</span><span class="sxs-lookup"><span data-stu-id="d9554-477">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="d9554-478">Beispiel</span><span class="sxs-lookup"><span data-stu-id="d9554-478">Example</span></span>

```
var normalizedSubject = Office.context.mailbox.item.normalizedSubject;
```

####  <a name="notificationmessages-notificationmessagesjavascriptapioutlook17officenotificationmessages"></a><span data-ttu-id="d9554-479">notificationMessages :[NotificationMessages](/javascript/api/outlook_1_7/office.notificationmessages)</span><span class="sxs-lookup"><span data-stu-id="d9554-479">notificationMessages :[NotificationMessages](/javascript/api/outlook_1_7/office.notificationmessages)</span></span>

<span data-ttu-id="d9554-480">Ruft die Benachrichtigungen für ein Element ab.</span><span class="sxs-lookup"><span data-stu-id="d9554-480">Gets the notification messages for an item.</span></span>

##### <a name="type"></a><span data-ttu-id="d9554-481">Typ:</span><span class="sxs-lookup"><span data-stu-id="d9554-481">Type:</span></span>

*   [<span data-ttu-id="d9554-482">NotificationMessages</span><span class="sxs-lookup"><span data-stu-id="d9554-482">NotificationMessages</span></span>](/javascript/api/outlook_1_7/office.notificationmessages)

##### <a name="requirements"></a><span data-ttu-id="d9554-483">Anforderungen</span><span class="sxs-lookup"><span data-stu-id="d9554-483">Requirements</span></span>

|<span data-ttu-id="d9554-484">Anforderung</span><span class="sxs-lookup"><span data-stu-id="d9554-484">Requirement</span></span>|<span data-ttu-id="d9554-485">Wert</span><span class="sxs-lookup"><span data-stu-id="d9554-485">Value</span></span>|
|---|---|
|[<span data-ttu-id="d9554-486">Mindestversion des Postfachanforderungssatzes</span><span class="sxs-lookup"><span data-stu-id="d9554-486">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="d9554-487">1.3</span><span class="sxs-lookup"><span data-stu-id="d9554-487">1.3</span></span>|
|[<span data-ttu-id="d9554-488">Mindestberechtigungsstufe</span><span class="sxs-lookup"><span data-stu-id="d9554-488">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="d9554-489">ReadItem</span><span class="sxs-lookup"><span data-stu-id="d9554-489">ReadItem</span></span>|
|[<span data-ttu-id="d9554-490">Zutreffender Outlook-Modus</span><span class="sxs-lookup"><span data-stu-id="d9554-490">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="d9554-491">Verfassen oder Lesen</span><span class="sxs-lookup"><span data-stu-id="d9554-491">Compose or read</span></span>|

####  <a name="optionalattendees-arrayemailaddressdetailsjavascriptapioutlook17officeemailaddressdetailsrecipientsjavascriptapioutlook17officerecipients"></a><span data-ttu-id="d9554-492">optionalAttendees :Array.<[EmailAddressDetails](/javascript/api/outlook_1_7/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook_1_7/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="d9554-492">optionalAttendees :Array.<[EmailAddressDetails](/javascript/api/outlook_1_7/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook_1_7/office.recipients)</span></span>

<span data-ttu-id="d9554-493">Ermöglicht den Zugriff auf die optionalen Teilnehmer eines Ereignisses.</span><span class="sxs-lookup"><span data-stu-id="d9554-493">Provides access to the optional attendees of an event.</span></span> <span data-ttu-id="d9554-494">Der Typ des Objekts und die Zugriffsebene, hängt von den Modus des aktuellen Elements ab.</span><span class="sxs-lookup"><span data-stu-id="d9554-494">The type of object and level of access depends on the mode of the current item.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="d9554-495">Lesemodus</span><span class="sxs-lookup"><span data-stu-id="d9554-495">Read mode</span></span>

<span data-ttu-id="d9554-496">Die `optionalAttendees`-Eigenschaft gibt ein Array mit einem `EmailAddressDetails`-Objekt für jeden optionalen Teilnehmer an der Besprechung zurück.</span><span class="sxs-lookup"><span data-stu-id="d9554-496">The `optionalAttendees` property returns an array that contains an `EmailAddressDetails` object for each optional attendee to the meeting.</span></span>

##### <a name="compose-mode"></a><span data-ttu-id="d9554-497">Verfassenmodus</span><span class="sxs-lookup"><span data-stu-id="d9554-497">Compose mode</span></span>

<span data-ttu-id="d9554-498">Die `optionalAttendees` -Eigenschaft gibt eine `Recipients` -Objekt, das Methoden zum Abrufen oder Aktualisieren der optionalen Teilnehmer für eine Besprechung bereitstellt.</span><span class="sxs-lookup"><span data-stu-id="d9554-498">The `optionalAttendees` property returns a `Recipients` object that provides methods to get or update the optional attendees for a meeting.</span></span>

##### <a name="type"></a><span data-ttu-id="d9554-499">Typ:</span><span class="sxs-lookup"><span data-stu-id="d9554-499">Type:</span></span>

*   <span data-ttu-id="d9554-500">Array.<[EmailAddressDetails](/javascript/api/outlook_1_7/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook_1_7/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="d9554-500">Array.<[EmailAddressDetails](/javascript/api/outlook_1_7/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook_1_7/office.recipients)</span></span>

##### <a name="requirements"></a><span data-ttu-id="d9554-501">Anforderungen</span><span class="sxs-lookup"><span data-stu-id="d9554-501">Requirements</span></span>

|<span data-ttu-id="d9554-502">Anforderung</span><span class="sxs-lookup"><span data-stu-id="d9554-502">Requirement</span></span>|<span data-ttu-id="d9554-503">Wert</span><span class="sxs-lookup"><span data-stu-id="d9554-503">Value</span></span>|
|---|---|
|[<span data-ttu-id="d9554-504">Mindestversion des Postfachanforderungssatzes</span><span class="sxs-lookup"><span data-stu-id="d9554-504">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="d9554-505">1.0</span><span class="sxs-lookup"><span data-stu-id="d9554-505">1.0</span></span>|
|[<span data-ttu-id="d9554-506">Mindestberechtigungsstufe</span><span class="sxs-lookup"><span data-stu-id="d9554-506">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="d9554-507">ReadItem</span><span class="sxs-lookup"><span data-stu-id="d9554-507">ReadItem</span></span>|
|[<span data-ttu-id="d9554-508">Zutreffender Outlook-Modus</span><span class="sxs-lookup"><span data-stu-id="d9554-508">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="d9554-509">Verfassen oder Lesen</span><span class="sxs-lookup"><span data-stu-id="d9554-509">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="d9554-510">Beispiel</span><span class="sxs-lookup"><span data-stu-id="d9554-510">Example</span></span>

```
Office.context.mailbox.item.optionalAttendees.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.optionalAttendees.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.optionalAttendees.getAsync(callback);

function callback(asyncResult) {
  var arrayOfOptionalAttendeesRecipients = asyncResult.value;
}
```

#### <a name="organizer-emailaddressdetailsjavascriptapioutlook17officeemailaddressdetailsorganizerjavascriptapioutlook17officeorganizer"></a><span data-ttu-id="d9554-511">Organizer:[EmailAddressDetails](/javascript/api/outlook_1_7/office.emailaddressdetails)|[Organizer](/javascript/api/outlook_1_7/office.organizer)</span><span class="sxs-lookup"><span data-stu-id="d9554-511">organizer :[EmailAddressDetails](/javascript/api/outlook_1_7/office.emailaddressdetails)|[Organizer](/javascript/api/outlook_1_7/office.organizer)</span></span>

<span data-ttu-id="d9554-512">Ruft die e-Mail-Adresse des Dialogfelds Organisieren für eine angegebene Besprechung ab.</span><span class="sxs-lookup"><span data-stu-id="d9554-512">Gets the email address of the organizer for a specified meeting.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="d9554-513">Lesemodus</span><span class="sxs-lookup"><span data-stu-id="d9554-513">Read mode</span></span>

<span data-ttu-id="d9554-514">Die `organizer` -Eigenschaft gibt ein [EmailAddressDetails](/javascript/api/outlook_1_7/office.emailaddressdetails) -Objekt, die den Besprechungsorganisator darstellt.</span><span class="sxs-lookup"><span data-stu-id="d9554-514">The `organizer` property returns an [EmailAddressDetails](/javascript/api/outlook_1_7/office.emailaddressdetails) object that represents the meeting organizer.</span></span>

##### <a name="compose-mode"></a><span data-ttu-id="d9554-515">Verfassenmodus</span><span class="sxs-lookup"><span data-stu-id="d9554-515">Compose mode</span></span>

<span data-ttu-id="d9554-516">Die `organizer` -Eigenschaft gibt ein [Organisator](/javascript/api/outlook_1_7/office.organizer) -Objekt, das eine Methode zum Abrufen des Werts der Organisator bereitstellt.</span><span class="sxs-lookup"><span data-stu-id="d9554-516">The `organizer` property returns an [Organizer](/javascript/api/outlook_1_7/office.organizer) object that provides a method to get the organizer value.</span></span>

##### <a name="type"></a><span data-ttu-id="d9554-517">Typ:</span><span class="sxs-lookup"><span data-stu-id="d9554-517">Type:</span></span>

*   <span data-ttu-id="d9554-518">[EmailAddressDetails](/javascript/api/outlook_1_7/office.emailaddressdetails) | [Organizer](/javascript/api/outlook_1_7/office.organizer)</span><span class="sxs-lookup"><span data-stu-id="d9554-518">[EmailAddressDetails](/javascript/api/outlook_1_7/office.emailaddressdetails) | [Organizer](/javascript/api/outlook_1_7/office.organizer)</span></span>

##### <a name="requirements"></a><span data-ttu-id="d9554-519">Anforderungen</span><span class="sxs-lookup"><span data-stu-id="d9554-519">Requirements</span></span>

|<span data-ttu-id="d9554-520">Anforderung</span><span class="sxs-lookup"><span data-stu-id="d9554-520">Requirement</span></span>|||
|---|---|---|
|[<span data-ttu-id="d9554-521">Mindestversion des Postfachanforderungssatzes</span><span class="sxs-lookup"><span data-stu-id="d9554-521">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="d9554-522">1.0</span><span class="sxs-lookup"><span data-stu-id="d9554-522">1.0</span></span>|<span data-ttu-id="d9554-523">1.7</span><span class="sxs-lookup"><span data-stu-id="d9554-523">1.7</span></span>|
|[<span data-ttu-id="d9554-524">Mindestberechtigungsstufe</span><span class="sxs-lookup"><span data-stu-id="d9554-524">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="d9554-525">ReadItem</span><span class="sxs-lookup"><span data-stu-id="d9554-525">ReadItem</span></span>|<span data-ttu-id="d9554-526">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="d9554-526">ReadWriteItem</span></span>|
|[<span data-ttu-id="d9554-527">Zutreffender Outlook-Modus</span><span class="sxs-lookup"><span data-stu-id="d9554-527">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="d9554-528">Lesen</span><span class="sxs-lookup"><span data-stu-id="d9554-528">Read</span></span>|<span data-ttu-id="d9554-529">Verfassen</span><span class="sxs-lookup"><span data-stu-id="d9554-529">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="d9554-530">Beispiel</span><span class="sxs-lookup"><span data-stu-id="d9554-530">Example</span></span>

```
var organizerName = Office.context.mailbox.item.organizer.displayName;
var organizerAddress = Office.context.mailbox.item.organizer.emailAddress;
```

#### <a name="nullable-recurrence-recurrencejavascriptapioutlook17officerecurrence"></a><span data-ttu-id="d9554-531">(Nullwerte zulassen) Serie:[Serie](/javascript/api/outlook_1_7/office.recurrence)</span><span class="sxs-lookup"><span data-stu-id="d9554-531">(nullable) recurrence :[Recurrence](/javascript/api/outlook_1_7/office.recurrence)</span></span>

<span data-ttu-id="d9554-532">Ruft ab oder legt das Serienmuster eines Termins.</span><span class="sxs-lookup"><span data-stu-id="d9554-532">Gets or sets the recurrence pattern of an appointment.</span></span> <span data-ttu-id="d9554-533">Ruft das Serienmuster eine Besprechungsanfrage an.</span><span class="sxs-lookup"><span data-stu-id="d9554-533">Gets the recurrence pattern of a meeting request.</span></span> <span data-ttu-id="d9554-534">Lesen Sie und verfassen Sie Modi für Terminelemente.</span><span class="sxs-lookup"><span data-stu-id="d9554-534">Read and compose modes for appointment items.</span></span> <span data-ttu-id="d9554-535">Zur Erfüllung der Anforderung Elemente im Lesemodus.</span><span class="sxs-lookup"><span data-stu-id="d9554-535">Read mode for meeting request items.</span></span>

<span data-ttu-id="d9554-536">Die `recurrence` -Eigenschaft gibt ein Objekt [Serie](/javascript/api/outlook_1_7/office.recurrence) für Termin- oder Besprechungsanfragen zurück, wenn ein Element eine Datenreihe oder eine Instanz einer Serie ist.</span><span class="sxs-lookup"><span data-stu-id="d9554-536">The `recurrence` property returns a [recurrence](/javascript/api/outlook_1_7/office.recurrence) object for recurring appointments or meetings requests if an item is a series or an instance in a series.</span></span> <span data-ttu-id="d9554-537">`null`wird für einzelne Termine und Besprechungsanfragen von einzelnen Terminen zurückgegeben.</span><span class="sxs-lookup"><span data-stu-id="d9554-537">`null` is returned for single appointments and meeting requests of single appointments.</span></span> <span data-ttu-id="d9554-538">`undefined`wird für Nachrichten zurückgegeben, die nicht Besprechungsanfragen sind.</span><span class="sxs-lookup"><span data-stu-id="d9554-538">`undefined` is returned for messages that are not meeting requests.</span></span>

> <span data-ttu-id="d9554-539">Hinweis: Besprechungsanfragen verfügen über eine `itemClass` Wert der IPM. Schedule.Meeting.Request.</span><span class="sxs-lookup"><span data-stu-id="d9554-539">Note: Meeting requests have an `itemClass` value of IPM.Schedule.Meeting.Request.</span></span>

> <span data-ttu-id="d9554-540">Hinweis: Wenn das Objekt Serie ist `null`, darauf hin, dass das Objekt einen einzelnen Termin oder eine Besprechungsanfrage an einen einzelnen Termin und nicht um einen Teil einer Reihe ist.</span><span class="sxs-lookup"><span data-stu-id="d9554-540">Note: If the recurrence object is `null`, this indicates that the object is a single appointment or a meeting request of a single appointment and NOT a part of a series.</span></span>

##### <a name="type"></a><span data-ttu-id="d9554-541">Typ:</span><span class="sxs-lookup"><span data-stu-id="d9554-541">Type:</span></span>

* [<span data-ttu-id="d9554-542">Serie</span><span class="sxs-lookup"><span data-stu-id="d9554-542">Recurrence</span></span>](/javascript/api/outlook_1_7/office.recurrence)

|<span data-ttu-id="d9554-543">Anforderung</span><span class="sxs-lookup"><span data-stu-id="d9554-543">Requirement</span></span>|<span data-ttu-id="d9554-544">Wert</span><span class="sxs-lookup"><span data-stu-id="d9554-544">Value</span></span>|
|---|---|
|[<span data-ttu-id="d9554-545">Mindestversion des Postfachanforderungssatzes</span><span class="sxs-lookup"><span data-stu-id="d9554-545">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="d9554-546">1.7</span><span class="sxs-lookup"><span data-stu-id="d9554-546">1.7</span></span>|
|[<span data-ttu-id="d9554-547">Mindestberechtigungsstufe</span><span class="sxs-lookup"><span data-stu-id="d9554-547">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="d9554-548">ReadItem</span><span class="sxs-lookup"><span data-stu-id="d9554-548">ReadItem</span></span>|
|[<span data-ttu-id="d9554-549">Zutreffender Outlook-Modus</span><span class="sxs-lookup"><span data-stu-id="d9554-549">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="d9554-550">Verfassen oder Lesen</span><span class="sxs-lookup"><span data-stu-id="d9554-550">Compose or read</span></span>|

####  <a name="requiredattendees-arrayemailaddressdetailsjavascriptapioutlook17officeemailaddressdetailsrecipientsjavascriptapioutlook17officerecipients"></a><span data-ttu-id="d9554-551">requiredAttendees :Array.<[EmailAddressDetails](/javascript/api/outlook_1_7/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook_1_7/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="d9554-551">requiredAttendees :Array.<[EmailAddressDetails](/javascript/api/outlook_1_7/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook_1_7/office.recipients)</span></span>

<span data-ttu-id="d9554-552">Ermöglicht den Zugriff auf die erforderlichen Teilnehmer eines Ereignisses.</span><span class="sxs-lookup"><span data-stu-id="d9554-552">Provides access to the required attendees of an event.</span></span> <span data-ttu-id="d9554-553">Der Typ des Objekts und die Zugriffsebene, hängt von den Modus des aktuellen Elements ab.</span><span class="sxs-lookup"><span data-stu-id="d9554-553">The type of object and level of access depends on the mode of the current item.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="d9554-554">Lesemodus</span><span class="sxs-lookup"><span data-stu-id="d9554-554">Read mode</span></span>

<span data-ttu-id="d9554-555">Die `requiredAttendees`-Eigenschaft gibt ein Array mit einem `EmailAddressDetails`-Objekt für jeden erforderlichen Teilnehmer an der Besprechung zurück.</span><span class="sxs-lookup"><span data-stu-id="d9554-555">The `requiredAttendees` property returns an array that contains an `EmailAddressDetails` object for each required attendee to the meeting.</span></span>

##### <a name="compose-mode"></a><span data-ttu-id="d9554-556">Verfassenmodus</span><span class="sxs-lookup"><span data-stu-id="d9554-556">Compose mode</span></span>

<span data-ttu-id="d9554-557">Die `requiredAttendees` -Eigenschaft gibt eine `Recipients` -Objekt, das Methoden zum Abrufen oder Aktualisieren der erforderlichen Teilnehmer für eine Besprechung bereitstellt.</span><span class="sxs-lookup"><span data-stu-id="d9554-557">The `requiredAttendees` property returns a `Recipients` object that provides methods to get or update the required attendees for a meeting.</span></span>

##### <a name="type"></a><span data-ttu-id="d9554-558">Typ:</span><span class="sxs-lookup"><span data-stu-id="d9554-558">Type:</span></span>

*   <span data-ttu-id="d9554-559">Array.<[EmailAddressDetails](/javascript/api/outlook_1_7/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook_1_7/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="d9554-559">Array.<[EmailAddressDetails](/javascript/api/outlook_1_7/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook_1_7/office.recipients)</span></span>

##### <a name="requirements"></a><span data-ttu-id="d9554-560">Anforderungen</span><span class="sxs-lookup"><span data-stu-id="d9554-560">Requirements</span></span>

|<span data-ttu-id="d9554-561">Anforderung</span><span class="sxs-lookup"><span data-stu-id="d9554-561">Requirement</span></span>|<span data-ttu-id="d9554-562">Wert</span><span class="sxs-lookup"><span data-stu-id="d9554-562">Value</span></span>|
|---|---|
|[<span data-ttu-id="d9554-563">Mindestversion des Postfachanforderungssatzes</span><span class="sxs-lookup"><span data-stu-id="d9554-563">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="d9554-564">1.0</span><span class="sxs-lookup"><span data-stu-id="d9554-564">1.0</span></span>|
|[<span data-ttu-id="d9554-565">Mindestberechtigungsstufe</span><span class="sxs-lookup"><span data-stu-id="d9554-565">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="d9554-566">ReadItem</span><span class="sxs-lookup"><span data-stu-id="d9554-566">ReadItem</span></span>|
|[<span data-ttu-id="d9554-567">Zutreffender Outlook-Modus</span><span class="sxs-lookup"><span data-stu-id="d9554-567">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="d9554-568">Verfassen oder Lesen</span><span class="sxs-lookup"><span data-stu-id="d9554-568">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="d9554-569">Beispiel</span><span class="sxs-lookup"><span data-stu-id="d9554-569">Example</span></span>

```
Office.context.mailbox.item.requiredAttendees.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.requiredAttendees.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.requiredAttendees.getAsync(callback);

function callback(asyncResult) {
  var arrayOfRequiredAttendeesRecipients = asyncResult.value;
}
```

#### <a name="sender-emailaddressdetailsjavascriptapioutlook17officeemailaddressdetails"></a><span data-ttu-id="d9554-570">sender :[EmailAddressDetails](/javascript/api/outlook_1_7/office.emailaddressdetails)</span><span class="sxs-lookup"><span data-stu-id="d9554-570">sender :[EmailAddressDetails](/javascript/api/outlook_1_7/office.emailaddressdetails)</span></span>

<span data-ttu-id="d9554-p126">Ruft die E-Mail-Adresse des Absenders einer E-Mail-Nachricht ab. Nur Lesemodus.</span><span class="sxs-lookup"><span data-stu-id="d9554-p126">Gets the email address of the sender of an email message. Read mode only.</span></span>

<span data-ttu-id="d9554-p127">Die [`from`](#from-emailaddressdetailsjavascriptapioutlook17officeemailaddressdetailsfromjavascriptapioutlook17officefrom)- und `sender`-Eigenschaften stellen dieselbe Person dar, außer die Nachricht wurde von einem Delegaten gesendet. In diesem Fall stellt die `from`-Eigenschaft die Stellvertretung dar, und die sender-Eigenschaft stellt den Delegaten dar.</span><span class="sxs-lookup"><span data-stu-id="d9554-p127">The [`from`](#from-emailaddressdetailsjavascriptapioutlook17officeemailaddressdetailsfromjavascriptapioutlook17officefrom) and `sender` properties represent the same person unless the message is sent by a delegate. In that case, the `from` property represents the delegator, and the sender property represents the delegate.</span></span>

> [!NOTE]
> <span data-ttu-id="d9554-575">Die `recipientType` -Eigenschaft des der `EmailAddressDetails` -Objekts der `sender` -Eigenschaft ist `undefined`.</span><span class="sxs-lookup"><span data-stu-id="d9554-575">The `recipientType` property of the `EmailAddressDetails` object in the `sender` property is `undefined`.</span></span>

##### <a name="type"></a><span data-ttu-id="d9554-576">Typ:</span><span class="sxs-lookup"><span data-stu-id="d9554-576">Type:</span></span>

*   [<span data-ttu-id="d9554-577">EmailAddressDetails</span><span class="sxs-lookup"><span data-stu-id="d9554-577">EmailAddressDetails</span></span>](/javascript/api/outlook_1_7/office.emailaddressdetails)

##### <a name="requirements"></a><span data-ttu-id="d9554-578">Anforderungen</span><span class="sxs-lookup"><span data-stu-id="d9554-578">Requirements</span></span>

|<span data-ttu-id="d9554-579">Anforderung</span><span class="sxs-lookup"><span data-stu-id="d9554-579">Requirement</span></span>|<span data-ttu-id="d9554-580">Wert</span><span class="sxs-lookup"><span data-stu-id="d9554-580">Value</span></span>|
|---|---|
|[<span data-ttu-id="d9554-581">Mindestversion des Postfachanforderungssatzes</span><span class="sxs-lookup"><span data-stu-id="d9554-581">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="d9554-582">1.0</span><span class="sxs-lookup"><span data-stu-id="d9554-582">1.0</span></span>|
|[<span data-ttu-id="d9554-583">Mindestberechtigungsstufe</span><span class="sxs-lookup"><span data-stu-id="d9554-583">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="d9554-584">ReadItem</span><span class="sxs-lookup"><span data-stu-id="d9554-584">ReadItem</span></span>|
|[<span data-ttu-id="d9554-585">Zutreffender Outlook-Modus</span><span class="sxs-lookup"><span data-stu-id="d9554-585">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="d9554-586">Lesen</span><span class="sxs-lookup"><span data-stu-id="d9554-586">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="d9554-587">Beispiel</span><span class="sxs-lookup"><span data-stu-id="d9554-587">Example</span></span>

```
var senderName = Office.context.mailbox.item.sender.displayName;
var senderAddress = Office.context.mailbox.item.sender.emailAddress;
```

#### <a name="nullable-seriesid-string"></a><span data-ttu-id="d9554-588">(Nullwerte zulassen) SeriesId: Zeichenfolge</span><span class="sxs-lookup"><span data-stu-id="d9554-588">(nullable) seriesId :String</span></span>

<span data-ttu-id="d9554-589">Ruft die Id der Datenreihe, zu der eine Instanz gehört.</span><span class="sxs-lookup"><span data-stu-id="d9554-589">Gets the id of the series that an instance belongs to.</span></span>

<span data-ttu-id="d9554-590">In OWA und Outlook die `seriesId` gibt die Exchange-Webdienste (EWS)-ID des übergeordneten (Reihe)-Elements, das dieses Element gehört.</span><span class="sxs-lookup"><span data-stu-id="d9554-590">In OWA and Outlook, the `seriesId` returns the Exchange Web Services (EWS) ID of the parent (series) item that this item belongs to.</span></span> <span data-ttu-id="d9554-591">Jedoch in iOS und Android die `seriesId` gibt die REST-ID des übergeordneten Elements zurück.</span><span class="sxs-lookup"><span data-stu-id="d9554-591">However, in iOS and Android, the `seriesId` returns the REST ID of the parent item.</span></span>

> [!NOTE]
> <span data-ttu-id="d9554-592">Der Bezeichner, der von der `seriesId`-Eigenschaft zurückgegeben wird, ist der gleiche Bezeichner wie der Exchange-Webdienste-Elementbezeichner.</span><span class="sxs-lookup"><span data-stu-id="d9554-592">The identifier returned by the `seriesId` property is the same as the Exchange Web Services item identifier.</span></span> <span data-ttu-id="d9554-593">Die `seriesId` -Eigenschaft ist nicht identisch mit der Outlook-IDs von der Outlook-REST-API verwendet.</span><span class="sxs-lookup"><span data-stu-id="d9554-593">The `seriesId` property is not identical to the Outlook IDs used by the Outlook REST API.</span></span> <span data-ttu-id="d9554-594">REST-API-Aufrufe dieser Wert verwenden, sollten sie konvertiert werden, bevor [Office.context.mailbox.convertToRestId](office.context.mailbox.md#converttorestiditemid-restversion--string)verwenden.</span><span class="sxs-lookup"><span data-stu-id="d9554-594">Before making REST API calls using this value, it should be converted using [Office.context.mailbox.convertToRestId](office.context.mailbox.md#converttorestiditemid-restversion--string).</span></span> <span data-ttu-id="d9554-595">Weitere Informationen finden Sie unter [Verwenden der Outlook-REST-APIs aus einem Outlook-add-in](https://docs.microsoft.com/outlook/add-ins/use-rest-api).</span><span class="sxs-lookup"><span data-stu-id="d9554-595">For more details, see [Use the Outlook REST APIs from an Outlook add-in](https://docs.microsoft.com/outlook/add-ins/use-rest-api).</span></span>

<span data-ttu-id="d9554-596">Die `seriesId` -Eigenschaft gibt `null` für Elemente, deren nicht das übergeordnete Elemente wie Termine Serienelementen einzelne, oder Besprechungsanfragen und gibt `undefined` für alle anderen Elemente, die nicht Anforderungen erfüllt sind.</span><span class="sxs-lookup"><span data-stu-id="d9554-596">The `seriesId` property returns `null` for items that do not have parent items such as single appointments, series items, or meeting requests and returns `undefined` for any other items that are not meeting requests.</span></span>

##### <a name="type"></a><span data-ttu-id="d9554-597">Typ:</span><span class="sxs-lookup"><span data-stu-id="d9554-597">Type:</span></span>

* <span data-ttu-id="d9554-598">String</span><span class="sxs-lookup"><span data-stu-id="d9554-598">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="d9554-599">Anforderungen</span><span class="sxs-lookup"><span data-stu-id="d9554-599">Requirements</span></span>

|<span data-ttu-id="d9554-600">Anforderung</span><span class="sxs-lookup"><span data-stu-id="d9554-600">Requirement</span></span>|<span data-ttu-id="d9554-601">Wert</span><span class="sxs-lookup"><span data-stu-id="d9554-601">Value</span></span>|
|---|---|
|[<span data-ttu-id="d9554-602">Mindestversion des Postfachanforderungssatzes</span><span class="sxs-lookup"><span data-stu-id="d9554-602">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="d9554-603">1.7</span><span class="sxs-lookup"><span data-stu-id="d9554-603">1.7</span></span>|
|[<span data-ttu-id="d9554-604">Mindestberechtigungsstufe</span><span class="sxs-lookup"><span data-stu-id="d9554-604">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="d9554-605">ReadItem</span><span class="sxs-lookup"><span data-stu-id="d9554-605">ReadItem</span></span>|
|[<span data-ttu-id="d9554-606">Zutreffender Outlook-Modus</span><span class="sxs-lookup"><span data-stu-id="d9554-606">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="d9554-607">Verfassen oder Lesen</span><span class="sxs-lookup"><span data-stu-id="d9554-607">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="d9554-608">Beispiel</span><span class="sxs-lookup"><span data-stu-id="d9554-608">Example</span></span>

```
var seriesId = Office.context.mailbox.item.seriesId; 
var isSeries = (seriesId == null);
```

####  <a name="start-datetimejavascriptapioutlook17officetime"></a><span data-ttu-id="d9554-609">start :Date|[Time](/javascript/api/outlook_1_7/office.time)</span><span class="sxs-lookup"><span data-stu-id="d9554-609">start :Date|[Time](/javascript/api/outlook_1_7/office.time)</span></span>

<span data-ttu-id="d9554-610">Ruft Datum und Zeit für den Beginn des Termins ab oder legt Datum und Uhrzeit fest.</span><span class="sxs-lookup"><span data-stu-id="d9554-610">Gets or sets the date and time that the appointment is to begin.</span></span>

<span data-ttu-id="d9554-p130">Die `start`-Eigenschaft wird als Datums- und Uhrzeitwert in UTC ausgedrückt. Sie können die [`convertToLocalClientTime`](office.context.mailbox.md#converttolocalclienttimetimevalue--localclienttimejavascriptapioutlook17officelocalclienttime)-Methode verwenden, um den Wert in den lokalen Uhrzeit- und Datumswert des Clients umzuwandeln.</span><span class="sxs-lookup"><span data-stu-id="d9554-p130">The `start` property is expressed as a Coordinated Universal Time (UTC) date and time value. You can use the [`convertToLocalClientTime`](office.context.mailbox.md#converttolocalclienttimetimevalue--localclienttimejavascriptapioutlook17officelocalclienttime) method to convert the value to the client’s local date and time.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="d9554-613">Lesemodus</span><span class="sxs-lookup"><span data-stu-id="d9554-613">Read mode</span></span>

<span data-ttu-id="d9554-614">Die `start`-Eigenschaft gibt ein `Date`-Objekt zurück.</span><span class="sxs-lookup"><span data-stu-id="d9554-614">The `start` property returns a `Date` object.</span></span>

##### <a name="compose-mode"></a><span data-ttu-id="d9554-615">Verfassenmodus</span><span class="sxs-lookup"><span data-stu-id="d9554-615">Compose mode</span></span>

<span data-ttu-id="d9554-616">Die `start`-Eigenschaft gibt ein `Time`-Objekt zurück.</span><span class="sxs-lookup"><span data-stu-id="d9554-616">The `start` property returns a `Time` object.</span></span>

<span data-ttu-id="d9554-617">Wenn Sie die [`Time.setAsync`](/javascript/api/outlook_1_7/office.time#setasync-datetime--options--callback-)-Methode verwenden, um die Startzeit im Verfassenmodus festzulegen, sollten Sie die [`convertToUtcClientTime`](office.context.mailbox.md#converttoutcclienttimeinput--date)-Methode verwenden, um die Ortszeit auf dem Client für den Server in UTC umzuwandeln.</span><span class="sxs-lookup"><span data-stu-id="d9554-617">When you use the [`Time.setAsync`](/javascript/api/outlook_1_7/office.time#setasync-datetime--options--callback-) method to set the start time, you should use the [`convertToUtcClientTime`](office.context.mailbox.md#converttoutcclienttimeinput--date) method to convert the local time on the client to UTC for the server.</span></span>

##### <a name="type"></a><span data-ttu-id="d9554-618">Typ:</span><span class="sxs-lookup"><span data-stu-id="d9554-618">Type:</span></span>

*   <span data-ttu-id="d9554-619">Date | [Time](/javascript/api/outlook_1_7/office.time)</span><span class="sxs-lookup"><span data-stu-id="d9554-619">Date | [Time](/javascript/api/outlook_1_7/office.time)</span></span>

##### <a name="requirements"></a><span data-ttu-id="d9554-620">Anforderungen</span><span class="sxs-lookup"><span data-stu-id="d9554-620">Requirements</span></span>

|<span data-ttu-id="d9554-621">Anforderung</span><span class="sxs-lookup"><span data-stu-id="d9554-621">Requirement</span></span>|<span data-ttu-id="d9554-622">Wert</span><span class="sxs-lookup"><span data-stu-id="d9554-622">Value</span></span>|
|---|---|
|[<span data-ttu-id="d9554-623">Mindestversion des Postfachanforderungssatzes</span><span class="sxs-lookup"><span data-stu-id="d9554-623">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="d9554-624">1.0</span><span class="sxs-lookup"><span data-stu-id="d9554-624">1.0</span></span>|
|[<span data-ttu-id="d9554-625">Mindestberechtigungsstufe</span><span class="sxs-lookup"><span data-stu-id="d9554-625">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="d9554-626">ReadItem</span><span class="sxs-lookup"><span data-stu-id="d9554-626">ReadItem</span></span>|
|[<span data-ttu-id="d9554-627">Zutreffender Outlook-Modus</span><span class="sxs-lookup"><span data-stu-id="d9554-627">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="d9554-628">Verfassen oder Lesen</span><span class="sxs-lookup"><span data-stu-id="d9554-628">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="d9554-629">Beispiel</span><span class="sxs-lookup"><span data-stu-id="d9554-629">Example</span></span>

<span data-ttu-id="d9554-630">Im folgenden Beispiel wird die Startzeit eines Termins im Verfassenmodus mithilfe der [`setAsync`](/javascript/api/outlook_1_7/office.time#setasync-datetime--options--callback-)-Methode des `Time`-Objekts festgelegt.</span><span class="sxs-lookup"><span data-stu-id="d9554-630">The following example sets the start time of an appointment in compose mode by using the [`setAsync`](/javascript/api/outlook_1_7/office.time#setasync-datetime--options--callback-) method of the `Time` object.</span></span>

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

####  <a name="subject-stringsubjectjavascriptapioutlook17officesubject"></a><span data-ttu-id="d9554-631">subject :String|[Subject](/javascript/api/outlook_1_7/office.subject)</span><span class="sxs-lookup"><span data-stu-id="d9554-631">subject :String|[Subject](/javascript/api/outlook_1_7/office.subject)</span></span>

<span data-ttu-id="d9554-632">Ruft die Beschreibung ab, die im Betrefffeld eines Elements angezeigt wird, oder legt sie fest.</span><span class="sxs-lookup"><span data-stu-id="d9554-632">Gets or sets the description that appears in the subject field of an item.</span></span>

<span data-ttu-id="d9554-633">Die `subject`-Eigenschaft ruft den gesamten Betreff des Elements ab oder legt ihn fest – so, wie er vom E-Mail-Server gesendet wird.</span><span class="sxs-lookup"><span data-stu-id="d9554-633">The `subject` property gets or sets the entire subject of the item, as sent by the email server.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="d9554-634">Lesemodus</span><span class="sxs-lookup"><span data-stu-id="d9554-634">Read mode</span></span>

<span data-ttu-id="d9554-p131">Die `subject`-Eigenschaft gibt eine Zeichenfolge zurück. Verwenden Sie die [`normalizedSubject`](#normalizedsubject-string)-Eigenschaft, um den Betreff ohne führende Präfixe wie `RE:` und `FW:` abzurufen.</span><span class="sxs-lookup"><span data-stu-id="d9554-p131">The `subject` property returns a string. Use the [`normalizedSubject`](#normalizedsubject-string) property to get the subject minus any leading prefixes such as `RE:` and `FW:`.</span></span>

```
var subject = Office.context.mailbox.item.subject;
```

##### <a name="compose-mode"></a><span data-ttu-id="d9554-637">Verfassenmodus</span><span class="sxs-lookup"><span data-stu-id="d9554-637">Compose mode</span></span>

<span data-ttu-id="d9554-638">Die `subject`-Eigenschaft gibt ein `Subject`-Objekt zurück, das Methoden zum Abrufen und Festlegen des Betreffs bereitstellt.</span><span class="sxs-lookup"><span data-stu-id="d9554-638">The `subject` property returns a `Subject` object that provides methods to get and set the subject.</span></span>

```
Office.context.mailbox.item.subject.getAsync(callback);

function callback(asyncResult) {
  var subject = asyncResult.value;
}
```

##### <a name="type"></a><span data-ttu-id="d9554-639">Typ:</span><span class="sxs-lookup"><span data-stu-id="d9554-639">Type:</span></span>

*   <span data-ttu-id="d9554-640">String | [Subject](/javascript/api/outlook_1_7/office.subject)</span><span class="sxs-lookup"><span data-stu-id="d9554-640">String | [Subject](/javascript/api/outlook_1_7/office.subject)</span></span>

##### <a name="requirements"></a><span data-ttu-id="d9554-641">Anforderungen</span><span class="sxs-lookup"><span data-stu-id="d9554-641">Requirements</span></span>

|<span data-ttu-id="d9554-642">Anforderung</span><span class="sxs-lookup"><span data-stu-id="d9554-642">Requirement</span></span>|<span data-ttu-id="d9554-643">Wert</span><span class="sxs-lookup"><span data-stu-id="d9554-643">Value</span></span>|
|---|---|
|[<span data-ttu-id="d9554-644">Mindestversion des Postfachanforderungssatzes</span><span class="sxs-lookup"><span data-stu-id="d9554-644">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="d9554-645">1.0</span><span class="sxs-lookup"><span data-stu-id="d9554-645">1.0</span></span>|
|[<span data-ttu-id="d9554-646">Mindestberechtigungsstufe</span><span class="sxs-lookup"><span data-stu-id="d9554-646">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="d9554-647">ReadItem</span><span class="sxs-lookup"><span data-stu-id="d9554-647">ReadItem</span></span>|
|[<span data-ttu-id="d9554-648">Zutreffender Outlook-Modus</span><span class="sxs-lookup"><span data-stu-id="d9554-648">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="d9554-649">Verfassen oder Lesen</span><span class="sxs-lookup"><span data-stu-id="d9554-649">Compose or read</span></span>|

####  <a name="to-arrayemailaddressdetailsjavascriptapioutlook17officeemailaddressdetailsrecipientsjavascriptapioutlook17officerecipients"></a><span data-ttu-id="d9554-650">an: Array. <[EmailAddressDetails](/javascript/api/outlook_1_7/office.emailaddressdetails)>|[Empfänger](/javascript/api/outlook_1_7/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="d9554-650">to :Array.<[EmailAddressDetails](/javascript/api/outlook_1_7/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook_1_7/office.recipients)</span></span>

<span data-ttu-id="d9554-651">Ermöglicht den Zugriff auf die Empfänger in der Zeile **an** einer Nachricht.</span><span class="sxs-lookup"><span data-stu-id="d9554-651">Provides access to the recipients on the **To** line of a message.</span></span> <span data-ttu-id="d9554-652">Der Typ des Objekts und die Zugriffsebene, hängt von den Modus des aktuellen Elements ab.</span><span class="sxs-lookup"><span data-stu-id="d9554-652">The type of object and level of access depends on the mode of the current item.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="d9554-653">Lesemodus</span><span class="sxs-lookup"><span data-stu-id="d9554-653">Read mode</span></span>

<span data-ttu-id="d9554-p133">Die `to`-Eigenschaft gibt ein Array mit einem `EmailAddressDetails`-Objekt für jeden Empfänger in der**An**-Zeile der Nachricht zurück. Die Auflistung ist auf höchstens 100 Elemente beschränkt.</span><span class="sxs-lookup"><span data-stu-id="d9554-p133">The `to` property returns an array that contains an `EmailAddressDetails` object for each recipient listed on the **To** line of the message. The collection is limited to a maximum of 100 members.</span></span>

##### <a name="compose-mode"></a><span data-ttu-id="d9554-656">Verfassenmodus</span><span class="sxs-lookup"><span data-stu-id="d9554-656">Compose mode</span></span>

<span data-ttu-id="d9554-657">Die `to` -Eigenschaft gibt eine `Recipients` -Objekt, das Methoden zum Abrufen oder aktualisieren die Empfänger in der Zeile **an** der Nachricht bereitstellt.</span><span class="sxs-lookup"><span data-stu-id="d9554-657">The `to` property returns a `Recipients` object that provides methods to get or update the recipients on the **To** line of the message.</span></span>

##### <a name="type"></a><span data-ttu-id="d9554-658">Typ:</span><span class="sxs-lookup"><span data-stu-id="d9554-658">Type:</span></span>

*   <span data-ttu-id="d9554-659">Array.<[EmailAddressDetails](/javascript/api/outlook_1_7/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook_1_7/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="d9554-659">Array.<[EmailAddressDetails](/javascript/api/outlook_1_7/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook_1_7/office.recipients)</span></span>

##### <a name="requirements"></a><span data-ttu-id="d9554-660">Anforderungen</span><span class="sxs-lookup"><span data-stu-id="d9554-660">Requirements</span></span>

|<span data-ttu-id="d9554-661">Anforderung</span><span class="sxs-lookup"><span data-stu-id="d9554-661">Requirement</span></span>|<span data-ttu-id="d9554-662">Wert</span><span class="sxs-lookup"><span data-stu-id="d9554-662">Value</span></span>|
|---|---|
|[<span data-ttu-id="d9554-663">Mindestversion des Postfachanforderungssatzes</span><span class="sxs-lookup"><span data-stu-id="d9554-663">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="d9554-664">1.0</span><span class="sxs-lookup"><span data-stu-id="d9554-664">1.0</span></span>|
|[<span data-ttu-id="d9554-665">Mindestberechtigungsstufe</span><span class="sxs-lookup"><span data-stu-id="d9554-665">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="d9554-666">ReadItem</span><span class="sxs-lookup"><span data-stu-id="d9554-666">ReadItem</span></span>|
|[<span data-ttu-id="d9554-667">Zutreffender Outlook-Modus</span><span class="sxs-lookup"><span data-stu-id="d9554-667">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="d9554-668">Verfassen oder Lesen</span><span class="sxs-lookup"><span data-stu-id="d9554-668">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="d9554-669">Beispiel</span><span class="sxs-lookup"><span data-stu-id="d9554-669">Example</span></span>

```
Office.context.mailbox.item.to.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.to.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.to.getAsync(callback);

function callback(asyncResult) {
  var arrayOfToRecipients = asyncResult.value;
}
```

### <a name="methods"></a><span data-ttu-id="d9554-670">Methoden</span><span class="sxs-lookup"><span data-stu-id="d9554-670">Methods</span></span>

####  <a name="addfileattachmentasyncuri-attachmentname-options-callback"></a><span data-ttu-id="d9554-671">addFileAttachmentAsync(uri, attachmentName, [options], [callback])</span><span class="sxs-lookup"><span data-stu-id="d9554-671">addFileAttachmentAsync(uri, attachmentName, [options], [callback])</span></span>

<span data-ttu-id="d9554-672">Fügt eine Datei zu einer Nachricht oder einem Termin als Anlage hinzu.</span><span class="sxs-lookup"><span data-stu-id="d9554-672">Adds a file to a message or appointment as an attachment.</span></span>

<span data-ttu-id="d9554-673">Die `addFileAttachmentAsync`-Methode lädt die Datei am angegebenen URI hoch und fügt sie an das Element im Verfassenformular an.</span><span class="sxs-lookup"><span data-stu-id="d9554-673">The `addFileAttachmentAsync` method uploads the file at the specified URI and attaches it to the item in the compose form.</span></span>

<span data-ttu-id="d9554-674">Anschließend können Sie den Bezeichner mit der [`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback)-Methode in der gleichen Sitzung zum Entfernen der Anlage verwenden.</span><span class="sxs-lookup"><span data-stu-id="d9554-674">You can subsequently use the identifier with the [`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback) method to remove the attachment in the same session.</span></span>

##### <a name="parameters"></a><span data-ttu-id="d9554-675">Parameter:</span><span class="sxs-lookup"><span data-stu-id="d9554-675">Parameters:</span></span>
|<span data-ttu-id="d9554-676">Name</span><span class="sxs-lookup"><span data-stu-id="d9554-676">Name</span></span>|<span data-ttu-id="d9554-677">Typ</span><span class="sxs-lookup"><span data-stu-id="d9554-677">Type</span></span>|<span data-ttu-id="d9554-678">Attribute</span><span class="sxs-lookup"><span data-stu-id="d9554-678">Attributes</span></span>|<span data-ttu-id="d9554-679">Beschreibung</span><span class="sxs-lookup"><span data-stu-id="d9554-679">Description</span></span>|
|---|---|---|---|
|`uri`|<span data-ttu-id="d9554-680">Zeichenfolge</span><span class="sxs-lookup"><span data-stu-id="d9554-680">String</span></span>||<span data-ttu-id="d9554-p134">Der URI, der den Speicherort der an die Nachricht oder den Termin anzuhängenden Datei angibt. Die maximale Länge ist 2048 Zeichen.</span><span class="sxs-lookup"><span data-stu-id="d9554-p134">The URI that provides the location of the file to attach to the message or appointment. The maximum length is 2048 characters.</span></span>|
|`attachmentName`|<span data-ttu-id="d9554-683">Zeichenfolge</span><span class="sxs-lookup"><span data-stu-id="d9554-683">String</span></span>||<span data-ttu-id="d9554-p135">Der Name der Anlage, der beim Hochladen der Anlage angezeigt wird. Die maximale Länge ist 255 Zeichen.</span><span class="sxs-lookup"><span data-stu-id="d9554-p135">The name of the attachment that is shown while the attachment is uploading. The maximum length is 255 characters.</span></span>|
|`options`|<span data-ttu-id="d9554-686">Object</span><span class="sxs-lookup"><span data-stu-id="d9554-686">Object</span></span>|<span data-ttu-id="d9554-687">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="d9554-687">&lt;optional&gt;</span></span>|<span data-ttu-id="d9554-688">Ein Objektliteral, das mindestens eine der folgenden Eigenschaften enthält.</span><span class="sxs-lookup"><span data-stu-id="d9554-688">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`|<span data-ttu-id="d9554-689">Object</span><span class="sxs-lookup"><span data-stu-id="d9554-689">Object</span></span>|<span data-ttu-id="d9554-690">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="d9554-690">&lt;optional&gt;</span></span>|<span data-ttu-id="d9554-691">Entwickler können jedes Objekt angeben, auf das sie in der Rückrufmethode zugreifen möchten.</span><span class="sxs-lookup"><span data-stu-id="d9554-691">Developers can provide any object they wish to access in the callback method.</span></span>|
|`options.isInline`|<span data-ttu-id="d9554-692">Boolean</span><span class="sxs-lookup"><span data-stu-id="d9554-692">Boolean</span></span>|<span data-ttu-id="d9554-693">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="d9554-693">&lt;optional&gt;</span></span>|<span data-ttu-id="d9554-694">Wenn `true`: Zeigt an, dass die Anlage inline im Nachrichtentext angezeigt wird und nicht in der Anlagenlisten angezeigt werden soll.</span><span class="sxs-lookup"><span data-stu-id="d9554-694">If `true`, indicates that the attachment will be shown inline in the message body, and should not be displayed in the attachment list.</span></span>|
|`callback`|<span data-ttu-id="d9554-695">function</span><span class="sxs-lookup"><span data-stu-id="d9554-695">function</span></span>|<span data-ttu-id="d9554-696">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="d9554-696">&lt;optional&gt;</span></span>|<span data-ttu-id="d9554-697">Wenn die Methode abgeschlossen wird, wird die Funktion , die im `callback`-Parameter übergeben wird, mit einem einzelnen Parameter (`asyncResult`) aufgerufen, der ein [`AsyncResult`](/javascript/api/office/office.asyncresult)-Objekt darstellt.</span><span class="sxs-lookup"><span data-stu-id="d9554-697">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span> <br/><span data-ttu-id="d9554-698">Bei Erfolg wird der Anlagenbezeichner in der `asyncResult.value`-Eigenschaft bereitgestellt.</span><span class="sxs-lookup"><span data-stu-id="d9554-698">On success, the attachment identifier will be provided in the `asyncResult.value` property.</span></span><br/><span data-ttu-id="d9554-699">Wenn beim Hochladen der Anlage ein Fehler auftritt, enthält das `asyncResult`-Objekt ein `Error`-Objekt mit einer Beschreibung des Fehlers.</span><span class="sxs-lookup"><span data-stu-id="d9554-699">If uploading the attachment fails, the `asyncResult` object will contain an `Error` object that provides a description of the error.</span></span>|

##### <a name="errors"></a><span data-ttu-id="d9554-700">Fehler</span><span class="sxs-lookup"><span data-stu-id="d9554-700">Errors</span></span>

|<span data-ttu-id="d9554-701">Fehlercode</span><span class="sxs-lookup"><span data-stu-id="d9554-701">Error code</span></span>|<span data-ttu-id="d9554-702">Beschreibung</span><span class="sxs-lookup"><span data-stu-id="d9554-702">Description</span></span>|
|------------|-------------|
|`AttachmentSizeExceeded`|<span data-ttu-id="d9554-703">Die Anlage ist zu groß.</span><span class="sxs-lookup"><span data-stu-id="d9554-703">The attachment is larger than allowed.</span></span>|
|`FileTypeNotSupported`|<span data-ttu-id="d9554-704">Die Anlage enthält eine Erweiterung, die nicht zulässig ist.</span><span class="sxs-lookup"><span data-stu-id="d9554-704">The attachment has an extension that is not allowed.</span></span>|
|`NumberOfAttachmentsExceeded`|<span data-ttu-id="d9554-705">Die Nachricht oder der Termin enthält zu viele Anlagen.</span><span class="sxs-lookup"><span data-stu-id="d9554-705">The message or appointment has too many attachments.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="d9554-706">Anforderungen</span><span class="sxs-lookup"><span data-stu-id="d9554-706">Requirements</span></span>

|<span data-ttu-id="d9554-707">Anforderung</span><span class="sxs-lookup"><span data-stu-id="d9554-707">Requirement</span></span>|<span data-ttu-id="d9554-708">Wert</span><span class="sxs-lookup"><span data-stu-id="d9554-708">Value</span></span>|
|---|---|
|[<span data-ttu-id="d9554-709">Mindestversion des Postfachanforderungssatzes</span><span class="sxs-lookup"><span data-stu-id="d9554-709">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="d9554-710">1.1</span><span class="sxs-lookup"><span data-stu-id="d9554-710">1.1</span></span>|
|[<span data-ttu-id="d9554-711">Mindestberechtigungsstufe</span><span class="sxs-lookup"><span data-stu-id="d9554-711">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="d9554-712">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="d9554-712">ReadWriteItem</span></span>|
|[<span data-ttu-id="d9554-713">Zutreffender Outlook-Modus</span><span class="sxs-lookup"><span data-stu-id="d9554-713">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="d9554-714">Verfassen</span><span class="sxs-lookup"><span data-stu-id="d9554-714">Compose</span></span>|

##### <a name="examples"></a><span data-ttu-id="d9554-715">Beispiele</span><span class="sxs-lookup"><span data-stu-id="d9554-715">Examples</span></span>

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

<span data-ttu-id="d9554-716">Im folgenden Beispiel wird eine Bilddatei als Inlineanlage hinzugefügt, und die Anlage wird im Nachrichtentext referenziert.</span><span class="sxs-lookup"><span data-stu-id="d9554-716">The following example adds an image file as an inline attachment and references the attachment in the message body.</span></span>

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

####  <a name="addhandlerasynceventtype-handler-options-callback"></a><span data-ttu-id="d9554-717">addHandlerAsync(eventType, handler, [options], [callback])</span><span class="sxs-lookup"><span data-stu-id="d9554-717">addHandlerAsync(eventType, handler, [options], [callback])</span></span>

<span data-ttu-id="d9554-718">Fügt einen Ereignishandler für ein unterstütztes Ereignis hinzu.</span><span class="sxs-lookup"><span data-stu-id="d9554-718">Adds an event handler for a supported event.</span></span>

<span data-ttu-id="d9554-719">Derzeit sind die unterstützten Ereignistypen `Office.EventType.AppointmentTimeChanged`, `Office.EventType.RecipientsChanged`, und`Office.EventType.RecurrenceChanged`</span><span class="sxs-lookup"><span data-stu-id="d9554-719">Currently the supported event types are `Office.EventType.AppointmentTimeChanged`, `Office.EventType.RecipientsChanged`, and `Office.EventType.RecurrenceChanged`</span></span>

##### <a name="parameters"></a><span data-ttu-id="d9554-720">Parameter:</span><span class="sxs-lookup"><span data-stu-id="d9554-720">Parameters:</span></span>

| <span data-ttu-id="d9554-721">Name</span><span class="sxs-lookup"><span data-stu-id="d9554-721">Name</span></span> | <span data-ttu-id="d9554-722">Typ</span><span class="sxs-lookup"><span data-stu-id="d9554-722">Type</span></span> | <span data-ttu-id="d9554-723">Attribute</span><span class="sxs-lookup"><span data-stu-id="d9554-723">Attributes</span></span> | <span data-ttu-id="d9554-724">Beschreibung</span><span class="sxs-lookup"><span data-stu-id="d9554-724">Description</span></span> |
|---|---|---|---|
| `eventType` | [<span data-ttu-id="d9554-725">Office.EventType</span><span class="sxs-lookup"><span data-stu-id="d9554-725">Office.EventType</span></span>](office.md#eventtype-string) || <span data-ttu-id="d9554-726">Das Ereignis, das den Handler aufrufen soll</span><span class="sxs-lookup"><span data-stu-id="d9554-726">The event that should invoke the handler.</span></span> |
| `handler` | <span data-ttu-id="d9554-727">Function</span><span class="sxs-lookup"><span data-stu-id="d9554-727">Function</span></span> || <span data-ttu-id="d9554-p136">Die Funktion, die das Ereignis behandeln soll. Die Funktion muss einen einzigen Parameter akzeptieren (ein Objektliteral). Die Eigenschaft `type` dieses Parameters entspricht dem Parameter `eventType`, der an `addHandlerAsync` übergeben wird.</span><span class="sxs-lookup"><span data-stu-id="d9554-p136">The function to handle the event. The function must accept a single parameter, which is an object literal. The `type` property on the parameter will match the `eventType` parameter passed to `addHandlerAsync`.</span></span> |
| `options` | <span data-ttu-id="d9554-731">Objekt</span><span class="sxs-lookup"><span data-stu-id="d9554-731">Object</span></span> | <span data-ttu-id="d9554-732">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="d9554-732">&lt;optional&gt;</span></span> | <span data-ttu-id="d9554-733">Ein Objektliteral, das mindestens eine der folgenden Eigenschaften enthält.</span><span class="sxs-lookup"><span data-stu-id="d9554-733">An object literal that contains one or more of the following properties.</span></span> |
| `options.asyncContext` | <span data-ttu-id="d9554-734">Object</span><span class="sxs-lookup"><span data-stu-id="d9554-734">Object</span></span> | <span data-ttu-id="d9554-735">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="d9554-735">&lt;optional&gt;</span></span> | <span data-ttu-id="d9554-736">Entwickler können ein Objekt bereitstellen, auf das sie in der Callbackmethode zugreifen möchten.</span><span class="sxs-lookup"><span data-stu-id="d9554-736">Developers can provide any object they wish to access in the callback method.</span></span> |
| `callback` | <span data-ttu-id="d9554-737">Funktion</span><span class="sxs-lookup"><span data-stu-id="d9554-737">function</span></span>| <span data-ttu-id="d9554-738">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="d9554-738">&lt;optional&gt;</span></span>|<span data-ttu-id="d9554-739">Wenn die Methode abgeschlossen wird, wird die Funktion , die im `callback`-Parameter übergeben wird, mit einem einzelnen Parameter (`asyncResult`) aufgerufen, der ein [`AsyncResult`](/javascript/api/office/office.asyncresult)-Objekt darstellt.</span><span class="sxs-lookup"><span data-stu-id="d9554-739">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="d9554-740">Anforderungen</span><span class="sxs-lookup"><span data-stu-id="d9554-740">Requirements</span></span>

|<span data-ttu-id="d9554-741">Anforderung</span><span class="sxs-lookup"><span data-stu-id="d9554-741">Requirement</span></span>| <span data-ttu-id="d9554-742">Wert</span><span class="sxs-lookup"><span data-stu-id="d9554-742">Value</span></span>|
|---|---|
|[<span data-ttu-id="d9554-743">Mindestversion des Postfachanforderungssatzes</span><span class="sxs-lookup"><span data-stu-id="d9554-743">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="d9554-744">1.7</span><span class="sxs-lookup"><span data-stu-id="d9554-744">1.7</span></span> |
|[<span data-ttu-id="d9554-745">Mindestberechtigungsstufe</span><span class="sxs-lookup"><span data-stu-id="d9554-745">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="d9554-746">ReadItem</span><span class="sxs-lookup"><span data-stu-id="d9554-746">ReadItem</span></span> |
|[<span data-ttu-id="d9554-747">Zutreffender Outlook-Modus</span><span class="sxs-lookup"><span data-stu-id="d9554-747">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="d9554-748">Verfassen oder Lesen</span><span class="sxs-lookup"><span data-stu-id="d9554-748">Compose or read</span></span> |

##### <a name="example"></a><span data-ttu-id="d9554-749">Beispiel</span><span class="sxs-lookup"><span data-stu-id="d9554-749">Example</span></span>

```
Office.initialize = function (reason) {
  $(document).ready(function () {
    Office.context.mailbox.item.addHandlerAsync(Office.EventType.RecurrenceChanged, loadNewItem, function (result) {
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

####  <a name="additemattachmentasyncitemid-attachmentname-options-callback"></a><span data-ttu-id="d9554-750">addItemAttachmentAsync(itemId, attachmentName, [options], [callback])</span><span class="sxs-lookup"><span data-stu-id="d9554-750">addItemAttachmentAsync(itemId, attachmentName, [options], [callback])</span></span>

<span data-ttu-id="d9554-751">Fügt der Nachricht oder dem Termin ein Exchange-Objekt, wie z. B. eine Nachricht, als Anhang hinzu.</span><span class="sxs-lookup"><span data-stu-id="d9554-751">Adds an Exchange item, such as a message, as an attachment to the message or appointment.</span></span>

<span data-ttu-id="d9554-p137">Die `addItemAttachmentAsync`-Methode fügt das Element mit dem angegebenen Exchange-Bezeichner an das Element im Formular zum Verfassen an. Wenn Sie eine Rückrufmethode angeben, wird die Methode mit einem `asyncResult`-Parameter aufgerufen, der entweder den Anlagenbezeichner oder einen Code enthält, der etwaige Fehler angibt, die beim Anfügen des Objekts aufgetreten sind. Sie können ggf. den `options`-Parameter verwenden, um Statusinformationen an die Rückrufmethode zu übergeben.</span><span class="sxs-lookup"><span data-stu-id="d9554-p137">The `addItemAttachmentAsync` method attaches the item with the specified Exchange identifier to the item in the compose form. If you specify a callback method, the method is called with one parameter, `asyncResult`, which contains either the attachment identifier or a code that indicates any error that occurred while attaching the item. You can use the `options` parameter to pass state information to the callback method, if needed.</span></span>

<span data-ttu-id="d9554-755">Anschließend können Sie den Bezeichner mit der [`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback)-Methode in der gleichen Sitzung zum Entfernen der Anlage verwenden.</span><span class="sxs-lookup"><span data-stu-id="d9554-755">You can subsequently use the identifier with the [`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback) method to remove the attachment in the same session.</span></span>

<span data-ttu-id="d9554-756">Wenn Ihre Office-Add-Ins in Outlook Web App ausgeführt wird die `addItemAttachmentAsync` -Methode kann Anfügen von Elementen in der Elemente, die Sie bearbeiten; Allerdings Dies wird nicht unterstützt und wird nicht empfohlen.</span><span class="sxs-lookup"><span data-stu-id="d9554-756">If your Office Add-in is running in Outlook Web App, the `addItemAttachmentAsync` method can attach items to items other than the item that you are editing; however, this is not supported and is not recommended.</span></span>

##### <a name="parameters"></a><span data-ttu-id="d9554-757">Parameter:</span><span class="sxs-lookup"><span data-stu-id="d9554-757">Parameters:</span></span>

|<span data-ttu-id="d9554-758">Name</span><span class="sxs-lookup"><span data-stu-id="d9554-758">Name</span></span>|<span data-ttu-id="d9554-759">Typ</span><span class="sxs-lookup"><span data-stu-id="d9554-759">Type</span></span>|<span data-ttu-id="d9554-760">Attribute</span><span class="sxs-lookup"><span data-stu-id="d9554-760">Attributes</span></span>|<span data-ttu-id="d9554-761">Beschreibung</span><span class="sxs-lookup"><span data-stu-id="d9554-761">Description</span></span>|
|---|---|---|---|
|`itemId`|<span data-ttu-id="d9554-762">Zeichenfolge</span><span class="sxs-lookup"><span data-stu-id="d9554-762">String</span></span>||<span data-ttu-id="d9554-p138">Der Exchange-Bezeichner des Objekts, das angehängt werden soll. Die maximale Länge beträgt 100 Zeichen.</span><span class="sxs-lookup"><span data-stu-id="d9554-p138">The Exchange identifier of the item to attach. The maximum length is 100 characters.</span></span>|
|`attachmentName`|<span data-ttu-id="d9554-765">Zeichenfolge</span><span class="sxs-lookup"><span data-stu-id="d9554-765">String</span></span>||<span data-ttu-id="d9554-p139">Der Betreff des Elements, das angehängt werden soll. Die maximale Länge ist 255 Zeichen.</span><span class="sxs-lookup"><span data-stu-id="d9554-p139">The sujbect of the item to be attached. The maximum length is 255 characters.</span></span>|
|`options`|<span data-ttu-id="d9554-768">Object</span><span class="sxs-lookup"><span data-stu-id="d9554-768">Object</span></span>|<span data-ttu-id="d9554-769">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="d9554-769">&lt;optional&gt;</span></span>|<span data-ttu-id="d9554-770">Ein Objektliteral, das mindestens eine der folgenden Eigenschaften enthält.</span><span class="sxs-lookup"><span data-stu-id="d9554-770">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`|<span data-ttu-id="d9554-771">Object</span><span class="sxs-lookup"><span data-stu-id="d9554-771">Object</span></span>|<span data-ttu-id="d9554-772">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="d9554-772">&lt;optional&gt;</span></span>|<span data-ttu-id="d9554-773">Entwickler können ein Objekt bereitstellen, auf das sie in der Callbackmethode zugreifen möchten.</span><span class="sxs-lookup"><span data-stu-id="d9554-773">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`|<span data-ttu-id="d9554-774">Funktion</span><span class="sxs-lookup"><span data-stu-id="d9554-774">function</span></span>|<span data-ttu-id="d9554-775">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="d9554-775">&lt;optional&gt;</span></span>|<span data-ttu-id="d9554-776">Wenn die Methode abgeschlossen wird, wird die Funktion , die im `callback`-Parameter übergeben wird, mit einem einzelnen Parameter (`asyncResult`) aufgerufen, der ein [`AsyncResult`](/javascript/api/office/office.asyncresult)-Objekt darstellt.</span><span class="sxs-lookup"><span data-stu-id="d9554-776">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span> <br/><span data-ttu-id="d9554-777">Bei Erfolg wird der Anlagenbezeichner in der `asyncResult.value`-Eigenschaft bereitgestellt.</span><span class="sxs-lookup"><span data-stu-id="d9554-777">On success, the attachment identifier will be provided in the `asyncResult.value` property.</span></span><br/><span data-ttu-id="d9554-778">Wenn beim Hinzufügen der Anlage ein Fehler auftritt, enthält das `asyncResult`-Objekt ein `Error`-Objekt mit einer Beschreibung des Fehlers.</span><span class="sxs-lookup"><span data-stu-id="d9554-778">If adding the attachment fails, the `asyncResult` object will contain an `Error` object that provides a description of the error.</span></span>|

##### <a name="errors"></a><span data-ttu-id="d9554-779">Fehler</span><span class="sxs-lookup"><span data-stu-id="d9554-779">Errors</span></span>

|<span data-ttu-id="d9554-780">Fehlercode</span><span class="sxs-lookup"><span data-stu-id="d9554-780">Error code</span></span>|<span data-ttu-id="d9554-781">Beschreibung</span><span class="sxs-lookup"><span data-stu-id="d9554-781">Description</span></span>|
|------------|-------------|
|`NumberOfAttachmentsExceeded`|<span data-ttu-id="d9554-782">Die Nachricht oder der Termin enthält zu viele Anlagen.</span><span class="sxs-lookup"><span data-stu-id="d9554-782">The message or appointment has too many attachments.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="d9554-783">Anforderungen</span><span class="sxs-lookup"><span data-stu-id="d9554-783">Requirements</span></span>

|<span data-ttu-id="d9554-784">Anforderung</span><span class="sxs-lookup"><span data-stu-id="d9554-784">Requirement</span></span>|<span data-ttu-id="d9554-785">Wert</span><span class="sxs-lookup"><span data-stu-id="d9554-785">Value</span></span>|
|---|---|
|[<span data-ttu-id="d9554-786">Mindestversion des Postfachanforderungssatzes</span><span class="sxs-lookup"><span data-stu-id="d9554-786">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="d9554-787">1.1</span><span class="sxs-lookup"><span data-stu-id="d9554-787">1.1</span></span>|
|[<span data-ttu-id="d9554-788">Mindestberechtigungsstufe</span><span class="sxs-lookup"><span data-stu-id="d9554-788">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="d9554-789">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="d9554-789">ReadWriteItem</span></span>|
|[<span data-ttu-id="d9554-790">Zutreffender Outlook-Modus</span><span class="sxs-lookup"><span data-stu-id="d9554-790">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="d9554-791">Verfassen</span><span class="sxs-lookup"><span data-stu-id="d9554-791">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="d9554-792">Beispiel</span><span class="sxs-lookup"><span data-stu-id="d9554-792">Example</span></span>

<span data-ttu-id="d9554-793">Im folgenden Beispiel wird ein vorhandenes Outlook-Element als Anlage mit dem Namen `My Attachment` hinzugefügt.</span><span class="sxs-lookup"><span data-stu-id="d9554-793">The following example adds an existing Outlook item as an attachment with the name `My Attachment`.</span></span>

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

####  <a name="close"></a><span data-ttu-id="d9554-794">close()</span><span class="sxs-lookup"><span data-stu-id="d9554-794">close()</span></span>

<span data-ttu-id="d9554-795">Schließt das aktuelle Element, das gerade verfasst wird.</span><span class="sxs-lookup"><span data-stu-id="d9554-795">Closes the current item that is being composed.</span></span>

<span data-ttu-id="d9554-p140">Das Verhalten der `close`-Methode hängt vom aktuellen Status des verfassten Elements ab. Wenn das Element über nicht gespeicherte Änderungen verfügt, fordert der Client den Benutzer auf, die Schließenaktion zu speichern, zu verwerfen oder abzubrechen.</span><span class="sxs-lookup"><span data-stu-id="d9554-p140">The behavior of the `close` method depends on the current state of the item being composed. If the item has unsaved changes, the client prompts the user to save, discard, or cancel the close action.</span></span>

> [!NOTE]
> <span data-ttu-id="d9554-798">In Outlook im Web, wenn das Element einen Termin handelt, und es wurde bereits zuvor gespeichert wurde mit `saveAsync`, der Benutzer wird aufgefordert, speichern, löschen oder Abbrechen, auch wenn keine Änderungen aufgetreten sind, da das Element zuletzt gespeichert.</span><span class="sxs-lookup"><span data-stu-id="d9554-798">In Outlook on the web, if the item is an appointment and it has previously been saved using `saveAsync`, the user is prompted to save, discard, or cancel even if no changes have occurred since the item was last saved.</span></span>

<span data-ttu-id="d9554-799">Wenn die Nachricht im Outlook-Desktopclient eine Inlineantwort ist, hat die `close`-Methode keine Auswirkung.</span><span class="sxs-lookup"><span data-stu-id="d9554-799">In the Outlook desktop client, if the message is an inline reply, the `close` method has no effect.</span></span>

##### <a name="requirements"></a><span data-ttu-id="d9554-800">Anforderungen</span><span class="sxs-lookup"><span data-stu-id="d9554-800">Requirements</span></span>

|<span data-ttu-id="d9554-801">Anforderung</span><span class="sxs-lookup"><span data-stu-id="d9554-801">Requirement</span></span>|<span data-ttu-id="d9554-802">Wert</span><span class="sxs-lookup"><span data-stu-id="d9554-802">Value</span></span>|
|---|---|
|[<span data-ttu-id="d9554-803">Mindestversion des Postfachanforderungssatzes</span><span class="sxs-lookup"><span data-stu-id="d9554-803">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="d9554-804">1.3</span><span class="sxs-lookup"><span data-stu-id="d9554-804">1.3</span></span>|
|[<span data-ttu-id="d9554-805">Mindestberechtigungsstufe</span><span class="sxs-lookup"><span data-stu-id="d9554-805">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="d9554-806">Eingeschränkt</span><span class="sxs-lookup"><span data-stu-id="d9554-806">Restricted</span></span>|
|[<span data-ttu-id="d9554-807">Zutreffender Outlook-Modus</span><span class="sxs-lookup"><span data-stu-id="d9554-807">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="d9554-808">Verfassen</span><span class="sxs-lookup"><span data-stu-id="d9554-808">Compose</span></span>|

#### <a name="displayreplyallformformdata"></a><span data-ttu-id="d9554-809">displayReplyAllForm(formData)</span><span class="sxs-lookup"><span data-stu-id="d9554-809">displayReplyAllForm(formData)</span></span>

<span data-ttu-id="d9554-810">Zeigt ein Antwortformular an, das den Absender und alle Empfänger der ausgewählten Nachricht oder des Organisators und alle Teilnehmer des ausgewählten Termins enthält.</span><span class="sxs-lookup"><span data-stu-id="d9554-810">Displays a reply form that includes the sender and all recipients of the selected message or the organizer and all attendees of the selected appointment.</span></span>

> [!NOTE]
> <span data-ttu-id="d9554-811">Diese Methode wird in Outlook für iOS oder Outlook für Android nicht unterstützt.</span><span class="sxs-lookup"><span data-stu-id="d9554-811">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="d9554-812">In Outlook Web App wird das Antwortformular als Popupformular in der Dreispaltenansicht und als Popupformular in der Zwei- oder Einspaltenansicht angezeigt.</span><span class="sxs-lookup"><span data-stu-id="d9554-812">In Outlook Web App, the reply form is displayed as a pop-out form in the 3-column view and a pop-up form in the 2- or 1-column view.</span></span>

<span data-ttu-id="d9554-813">Wenn einer der Zeichenfolgenparameter seinen Grenzwert überschreitet, löst `displayReplyAllForm` eine Ausnahme aus.</span><span class="sxs-lookup"><span data-stu-id="d9554-813">If any of the string parameters exceed their limits, `displayReplyAllForm` throws an exception.</span></span>

<span data-ttu-id="d9554-p141">Wenn Anlagen im `formData.attachments`-Parameter angegeben werden, versuchen Outlook und Outlook Web App alle Anlagen herunterzuladen und sie an das Antwortformular anzufügen. Wenn Anlagen nicht hinzugefügt werden können, wird in der Formularbenutzeroberfläche ein Fehler ausgegeben. Wenn dies nicht möglich ist, wird keine Fehlermeldung ausgegeben.</span><span class="sxs-lookup"><span data-stu-id="d9554-p141">When attachments are specified in the `formData.attachments` parameter, Outlook and Outlook Web App attempt to download all attachments and attach them to the reply form. If any attachments fail to be added, an error is shown in the form UI. If this isn't possible, then no error message is thrown.</span></span>

##### <a name="parameters"></a><span data-ttu-id="d9554-817">Parameter:</span><span class="sxs-lookup"><span data-stu-id="d9554-817">Parameters:</span></span>

|<span data-ttu-id="d9554-818">Name</span><span class="sxs-lookup"><span data-stu-id="d9554-818">Name</span></span>|<span data-ttu-id="d9554-819">Typ</span><span class="sxs-lookup"><span data-stu-id="d9554-819">Type</span></span>|<span data-ttu-id="d9554-820">Attribute</span><span class="sxs-lookup"><span data-stu-id="d9554-820">Attributes</span></span>|<span data-ttu-id="d9554-821">Beschreibung</span><span class="sxs-lookup"><span data-stu-id="d9554-821">Description</span></span>|
|---|---|---|---|
|`formData`|<span data-ttu-id="d9554-822">Zeichenfolge &#124; Objekt</span><span class="sxs-lookup"><span data-stu-id="d9554-822">String &#124; Object</span></span>||<span data-ttu-id="d9554-p142">Eine Zeichenfolge, die Text- und HTML-Code enthält, die den Hauptteil des Antwortformulars darstellen. Die Zeichenfolge ist auf 32 KB beschränkt.</span><span class="sxs-lookup"><span data-stu-id="d9554-p142">A string that contains text and HTML and that represents the body of the reply form. The string is limited to 32 KB.</span></span><br/><span data-ttu-id="d9554-825">**ODER**</span><span class="sxs-lookup"><span data-stu-id="d9554-825">**OR**</span></span><br/><span data-ttu-id="d9554-p143">Ein Objekt, das Text- oder Anlagendaten und eine Rückruffunktion enthält. Das Objekt ist wie folgt definiert:</span><span class="sxs-lookup"><span data-stu-id="d9554-p143">An object that contains body or attachment data and a callback function. The object is defined as follows.</span></span>|
|`formData.htmlBody`|<span data-ttu-id="d9554-828">Zeichenfolge</span><span class="sxs-lookup"><span data-stu-id="d9554-828">String</span></span>|<span data-ttu-id="d9554-829">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="d9554-829">&lt;optional&gt;</span></span>|<span data-ttu-id="d9554-p144">Eine Zeichenfolge, die Text- und HTML-Code enthält, die den Hauptteil des Antwortformulars darstellen. Die Zeichenfolge ist auf 32 KB beschränkt.</span><span class="sxs-lookup"><span data-stu-id="d9554-p144">A string that contains text and HTML and that represents the body of the reply form. The string is limited to 32 KB.</span></span>
|`formData.attachments`|<span data-ttu-id="d9554-832">Array.&lt;Object&gt;</span><span class="sxs-lookup"><span data-stu-id="d9554-832">Array.&lt;Object&gt;</span></span>|<span data-ttu-id="d9554-833">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="d9554-833">&lt;optional&gt;</span></span>|<span data-ttu-id="d9554-834">Ein Array mit JSON-Objekten, die Datei- oder Elementanlagen sind.</span><span class="sxs-lookup"><span data-stu-id="d9554-834">An array of JSON objects that are either file or item attachments.</span></span>|
|`formData.attachments.type`|<span data-ttu-id="d9554-835">Zeichenfolge</span><span class="sxs-lookup"><span data-stu-id="d9554-835">String</span></span>||<span data-ttu-id="d9554-p145">Gibt den Typ der Anlage an. Muss `file` für eine Dateianlage oder `item` für eine Elementanlage sein.</span><span class="sxs-lookup"><span data-stu-id="d9554-p145">Indicates the type of attachment. Must be `file` for a file attachment or `item` for an item attachment.</span></span>|
|`formData.attachments.name`|<span data-ttu-id="d9554-838">Zeichenfolge</span><span class="sxs-lookup"><span data-stu-id="d9554-838">String</span></span>||<span data-ttu-id="d9554-839">Eine Zeichenfolge, die den Namen der Anlage mit bis zu 255 Zeichen enthält.</span><span class="sxs-lookup"><span data-stu-id="d9554-839">A string that contains the name of the attachment, up to 255 characters in length.</span></span>|
|`formData.attachments.url`|<span data-ttu-id="d9554-840">Zeichenfolge</span><span class="sxs-lookup"><span data-stu-id="d9554-840">String</span></span>||<span data-ttu-id="d9554-p146">Wird nur verwendet, wenn `type` auf `file` gesetzt ist. Der URI des Speicherorts für die Datei.</span><span class="sxs-lookup"><span data-stu-id="d9554-p146">Only used if `type` is set to `file`. The URI of the location for the file.</span></span>|
|`formData.attachments.isInline`|<span data-ttu-id="d9554-843">Boolean</span><span class="sxs-lookup"><span data-stu-id="d9554-843">Boolean</span></span>||<span data-ttu-id="d9554-p147">Wird nur verwendet, wenn `type` auf `file` gesetzt ist. Wenn `true`: Zeigt an, dass die Anlage inline im Nachrichtentext angezeigt wird und nicht in der Anlagenlisten angezeigt werden soll.</span><span class="sxs-lookup"><span data-stu-id="d9554-p147">Only used if `type` is set to `file`. If `true`, indicates that the attachment will be shown inline in the message body, and should not be displayed in the attachment list.</span></span>|
|`formData.attachments.itemId`|<span data-ttu-id="d9554-846">Zeichenfolge</span><span class="sxs-lookup"><span data-stu-id="d9554-846">String</span></span>||<span data-ttu-id="d9554-p148">Wird nur verwendet, wenn `type` auf `item` gesetzt ist. Die EWS-Element-ID der Anlage. Dies ist eine Zeichenfolge bis zu 100 Zeichen.</span><span class="sxs-lookup"><span data-stu-id="d9554-p148">Only used if `type` is set to `item`. The EWS item id of the attachment. This is a string up to 100 characters.</span></span>|
|`callback`|<span data-ttu-id="d9554-850">Funktion</span><span class="sxs-lookup"><span data-stu-id="d9554-850">function</span></span>|<span data-ttu-id="d9554-851">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="d9554-851">&lt;optional&gt;</span></span>|<span data-ttu-id="d9554-852">Bei Abschluss der Methode wird die im Parameter `callback` übergebene Funktion mit einem einzigen Parameter aufgerufen: `asyncResult`. Dieser Parameter ist ein Objekt des Typs [AsyncResult](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="d9554-852">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [AsyncResult](/javascript/api/office/office.asyncresult) object.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="d9554-853">Anforderungen</span><span class="sxs-lookup"><span data-stu-id="d9554-853">Requirements</span></span>

|<span data-ttu-id="d9554-854">Anforderung</span><span class="sxs-lookup"><span data-stu-id="d9554-854">Requirement</span></span>|<span data-ttu-id="d9554-855">Wert</span><span class="sxs-lookup"><span data-stu-id="d9554-855">Value</span></span>|
|---|---|
|[<span data-ttu-id="d9554-856">Mindestversion des Postfachanforderungssatzes</span><span class="sxs-lookup"><span data-stu-id="d9554-856">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="d9554-857">1.0</span><span class="sxs-lookup"><span data-stu-id="d9554-857">1.0</span></span>|
|[<span data-ttu-id="d9554-858">Mindestberechtigungsstufe</span><span class="sxs-lookup"><span data-stu-id="d9554-858">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="d9554-859">ReadItem</span><span class="sxs-lookup"><span data-stu-id="d9554-859">ReadItem</span></span>|
|[<span data-ttu-id="d9554-860">Zutreffender Outlook-Modus</span><span class="sxs-lookup"><span data-stu-id="d9554-860">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="d9554-861">Lesen</span><span class="sxs-lookup"><span data-stu-id="d9554-861">Read</span></span>|

##### <a name="examples"></a><span data-ttu-id="d9554-862">Beispiele</span><span class="sxs-lookup"><span data-stu-id="d9554-862">Examples</span></span>

<span data-ttu-id="d9554-863">Im folgenden Code wird eine Zeichenfolge an die `displayReplyAllForm`-Funktion übergeben.</span><span class="sxs-lookup"><span data-stu-id="d9554-863">The following code passes a string to the `displayReplyAllForm` function.</span></span>

```
Office.context.mailbox.item.displayReplyAllForm('hello there');
Office.context.mailbox.item.displayReplyAllForm('<b>hello there</b>');
```

<span data-ttu-id="d9554-864">Antworten mit einem leeren Textkörper.</span><span class="sxs-lookup"><span data-stu-id="d9554-864">Reply with an empty body.</span></span>

```
Office.context.mailbox.item.displayReplyAllForm({});
```

<span data-ttu-id="d9554-865">Antworten nur einem Textkörper.</span><span class="sxs-lookup"><span data-stu-id="d9554-865">Reply with just a body.</span></span>

```
Office.context.mailbox.item.displayReplyAllForm(
{
  'htmlBody' : 'hi'
});
```

<span data-ttu-id="d9554-866">Antworten mit einem Textkörper und einer Dateianlage.</span><span class="sxs-lookup"><span data-stu-id="d9554-866">Reply with a body and a file attachment.</span></span>

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

<span data-ttu-id="d9554-867">Antworten mit einem Textkörper und einer Elementanlage.</span><span class="sxs-lookup"><span data-stu-id="d9554-867">Reply with a body and an item attachment.</span></span>

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

<span data-ttu-id="d9554-868">Antworten Sie mit einem Textkörper, einer Dateianlage, einer Elementanlage und einem Callback.</span><span class="sxs-lookup"><span data-stu-id="d9554-868">Reply with a body, file attachment, item attachment, and a callback.</span></span>

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

#### <a name="displayreplyformformdata"></a><span data-ttu-id="d9554-869">displayReplyForm(formData)</span><span class="sxs-lookup"><span data-stu-id="d9554-869">displayReplyForm(formData)</span></span>

<span data-ttu-id="d9554-870">Zeigt ein Antwortformular an, das nur den Absender der ausgewählten Nachricht oder den Organisator des ausgewählten Termins enthält.</span><span class="sxs-lookup"><span data-stu-id="d9554-870">Displays a reply form that includes only the sender of the selected message or the organizer of the selected appointment.</span></span>

> [!NOTE]
> <span data-ttu-id="d9554-871">Diese Methode wird in Outlook für iOS oder Outlook für Android nicht unterstützt.</span><span class="sxs-lookup"><span data-stu-id="d9554-871">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="d9554-872">In Outlook Web App wird das Antwortformular als Popupformular in der Dreispaltenansicht und als Popupformular in der Zwei- oder Einspaltenansicht angezeigt.</span><span class="sxs-lookup"><span data-stu-id="d9554-872">In Outlook Web App, the reply form is displayed as a pop-out form in the 3-column view and a pop-up form in the 2- or 1-column view.</span></span>

<span data-ttu-id="d9554-873">Wenn einer der Zeichenfolgenparameter seinen Grenzwert überschreitet, löst `displayReplyForm` eine Ausnahme aus.</span><span class="sxs-lookup"><span data-stu-id="d9554-873">If any of the string parameters exceed their limits, `displayReplyForm` throws an exception.</span></span>

<span data-ttu-id="d9554-p149">Wenn Anlagen im `formData.attachments`-Parameter angegeben werden, versuchen Outlook und Outlook Web App alle Anlagen herunterzuladen und sie an das Antwortformular anzufügen. Wenn Anlagen nicht hinzugefügt werden können, wird in der Formularbenutzeroberfläche ein Fehler ausgegeben. Wenn dies nicht möglich ist, wird keine Fehlermeldung ausgegeben.</span><span class="sxs-lookup"><span data-stu-id="d9554-p149">When attachments are specified in the `formData.attachments` parameter, Outlook and Outlook Web App attempt to download all attachments and attach them to the reply form. If any attachments fail to be added, an error is shown in the form UI. If this isn't possible, then no error message is thrown.</span></span>

##### <a name="parameters"></a><span data-ttu-id="d9554-877">Parameter:</span><span class="sxs-lookup"><span data-stu-id="d9554-877">Parameters:</span></span>

|<span data-ttu-id="d9554-878">Name</span><span class="sxs-lookup"><span data-stu-id="d9554-878">Name</span></span>|<span data-ttu-id="d9554-879">Typ</span><span class="sxs-lookup"><span data-stu-id="d9554-879">Type</span></span>|<span data-ttu-id="d9554-880">Attribute</span><span class="sxs-lookup"><span data-stu-id="d9554-880">Attributes</span></span>|<span data-ttu-id="d9554-881">Beschreibung</span><span class="sxs-lookup"><span data-stu-id="d9554-881">Description</span></span>|
|---|---|---|---|
|`formData`|<span data-ttu-id="d9554-882">Zeichenfolge &#124; Objekt</span><span class="sxs-lookup"><span data-stu-id="d9554-882">String &#124; Object</span></span>||<span data-ttu-id="d9554-p150">Eine Zeichenfolge, die Text- und HTML-Code enthält, die den Hauptteil des Antwortformulars darstellen. Die Zeichenfolge ist auf 32 KB beschränkt.</span><span class="sxs-lookup"><span data-stu-id="d9554-p150">A string that contains text and HTML and that represents the body of the reply form. The string is limited to 32 KB.</span></span><br/><span data-ttu-id="d9554-885">**ODER**</span><span class="sxs-lookup"><span data-stu-id="d9554-885">**OR**</span></span><br/><span data-ttu-id="d9554-p151">Ein Objekt, das Text- oder Anlagendaten und eine Rückruffunktion enthält. Das Objekt ist wie folgt definiert:</span><span class="sxs-lookup"><span data-stu-id="d9554-p151">An object that contains body or attachment data and a callback function. The object is defined as follows.</span></span>|
|`formData.htmlBody`|<span data-ttu-id="d9554-888">Zeichenfolge</span><span class="sxs-lookup"><span data-stu-id="d9554-888">String</span></span>|<span data-ttu-id="d9554-889">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="d9554-889">&lt;optional&gt;</span></span>|<span data-ttu-id="d9554-p152">Eine Zeichenfolge, die Text- und HTML-Code enthält, die den Hauptteil des Antwortformulars darstellen. Die Zeichenfolge ist auf 32 KB beschränkt.</span><span class="sxs-lookup"><span data-stu-id="d9554-p152">A string that contains text and HTML and that represents the body of the reply form. The string is limited to 32 KB.</span></span>
|`formData.attachments`|<span data-ttu-id="d9554-892">Array.&lt;Object&gt;</span><span class="sxs-lookup"><span data-stu-id="d9554-892">Array.&lt;Object&gt;</span></span>|<span data-ttu-id="d9554-893">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="d9554-893">&lt;optional&gt;</span></span>|<span data-ttu-id="d9554-894">Ein Array mit JSON-Objekten, die Datei- oder Elementanlagen sind.</span><span class="sxs-lookup"><span data-stu-id="d9554-894">An array of JSON objects that are either file or item attachments.</span></span>|
|`formData.attachments.type`|<span data-ttu-id="d9554-895">Zeichenfolge</span><span class="sxs-lookup"><span data-stu-id="d9554-895">String</span></span>||<span data-ttu-id="d9554-p153">Gibt den Typ der Anlage an. Muss `file` für eine Dateianlage oder `item` für eine Elementanlage sein.</span><span class="sxs-lookup"><span data-stu-id="d9554-p153">Indicates the type of attachment. Must be `file` for a file attachment or `item` for an item attachment.</span></span>|
|`formData.attachments.name`|<span data-ttu-id="d9554-898">Zeichenfolge</span><span class="sxs-lookup"><span data-stu-id="d9554-898">String</span></span>||<span data-ttu-id="d9554-899">Eine Zeichenfolge, die den Namen der Anlage mit bis zu 255 Zeichen enthält.</span><span class="sxs-lookup"><span data-stu-id="d9554-899">A string that contains the name of the attachment, up to 255 characters in length.</span></span>|
|`formData.attachments.url`|<span data-ttu-id="d9554-900">Zeichenfolge</span><span class="sxs-lookup"><span data-stu-id="d9554-900">String</span></span>||<span data-ttu-id="d9554-p154">Wird nur verwendet, wenn `type` auf `file` gesetzt ist. Der URI des Speicherorts für die Datei.</span><span class="sxs-lookup"><span data-stu-id="d9554-p154">Only used if `type` is set to `file`. The URI of the location for the file.</span></span>|
|`formData.attachments.isInline`|<span data-ttu-id="d9554-903">Boolean</span><span class="sxs-lookup"><span data-stu-id="d9554-903">Boolean</span></span>||<span data-ttu-id="d9554-p155">Wird nur verwendet, wenn `type` auf `file` gesetzt ist. Wenn `true`: Zeigt an, dass die Anlage inline im Nachrichtentext angezeigt wird und nicht in der Anlagenlisten angezeigt werden soll.</span><span class="sxs-lookup"><span data-stu-id="d9554-p155">Only used if `type` is set to `file`. If `true`, indicates that the attachment will be shown inline in the message body, and should not be displayed in the attachment list.</span></span>|
|`formData.attachments.itemId`|<span data-ttu-id="d9554-906">Zeichenfolge</span><span class="sxs-lookup"><span data-stu-id="d9554-906">String</span></span>||<span data-ttu-id="d9554-p156">Wird nur verwendet, wenn `type` auf `item` gesetzt ist. Die EWS-Element-ID der Anlage. Dies ist eine Zeichenfolge bis zu 100 Zeichen.</span><span class="sxs-lookup"><span data-stu-id="d9554-p156">Only used if `type` is set to `item`. The EWS item id of the attachment. This is a string up to 100 characters.</span></span>|
|`callback`|<span data-ttu-id="d9554-910">Funktion</span><span class="sxs-lookup"><span data-stu-id="d9554-910">function</span></span>|<span data-ttu-id="d9554-911">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="d9554-911">&lt;optional&gt;</span></span>|<span data-ttu-id="d9554-912">Bei Abschluss der Methode wird die im Parameter `callback` übergebene Funktion mit einem einzigen Parameter aufgerufen: `asyncResult`. Dieser Parameter ist ein Objekt des Typs [AsyncResult](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="d9554-912">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [AsyncResult](/javascript/api/office/office.asyncresult) object.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="d9554-913">Anforderungen</span><span class="sxs-lookup"><span data-stu-id="d9554-913">Requirements</span></span>

|<span data-ttu-id="d9554-914">Anforderung</span><span class="sxs-lookup"><span data-stu-id="d9554-914">Requirement</span></span>|<span data-ttu-id="d9554-915">Wert</span><span class="sxs-lookup"><span data-stu-id="d9554-915">Value</span></span>|
|---|---|
|[<span data-ttu-id="d9554-916">Mindestversion des Postfachanforderungssatzes</span><span class="sxs-lookup"><span data-stu-id="d9554-916">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="d9554-917">1.0</span><span class="sxs-lookup"><span data-stu-id="d9554-917">1.0</span></span>|
|[<span data-ttu-id="d9554-918">Mindestberechtigungsstufe</span><span class="sxs-lookup"><span data-stu-id="d9554-918">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="d9554-919">ReadItem</span><span class="sxs-lookup"><span data-stu-id="d9554-919">ReadItem</span></span>|
|[<span data-ttu-id="d9554-920">Zutreffender Outlook-Modus</span><span class="sxs-lookup"><span data-stu-id="d9554-920">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="d9554-921">Lesen</span><span class="sxs-lookup"><span data-stu-id="d9554-921">Read</span></span>|

##### <a name="examples"></a><span data-ttu-id="d9554-922">Beispiele</span><span class="sxs-lookup"><span data-stu-id="d9554-922">Examples</span></span>

<span data-ttu-id="d9554-923">Im folgenden Code wird eine Zeichenfolge an die `displayReplyForm`-Funktion übergeben.</span><span class="sxs-lookup"><span data-stu-id="d9554-923">The following code passes a string to the `displayReplyForm` function.</span></span>

```
Office.context.mailbox.item.displayReplyForm('hello there');
Office.context.mailbox.item.displayReplyForm('<b>hello there</b>');
```

<span data-ttu-id="d9554-924">Antworten mit einem leeren Textkörper.</span><span class="sxs-lookup"><span data-stu-id="d9554-924">Reply with an empty body.</span></span>

```
Office.context.mailbox.item.displayReplyForm({});
```

<span data-ttu-id="d9554-925">Antworten nur einem Textkörper.</span><span class="sxs-lookup"><span data-stu-id="d9554-925">Reply with just a body.</span></span>

```
Office.context.mailbox.item.displayReplyForm(
{
  'htmlBody' : 'hi'
});
```

<span data-ttu-id="d9554-926">Antworten mit einem Textkörper und einer Dateianlage.</span><span class="sxs-lookup"><span data-stu-id="d9554-926">Reply with a body and a file attachment.</span></span>

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

<span data-ttu-id="d9554-927">Antworten mit einem Textkörper und einer Elementanlage.</span><span class="sxs-lookup"><span data-stu-id="d9554-927">Reply with a body and an item attachment.</span></span>

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

<span data-ttu-id="d9554-928">Antworten Sie mit einem Textkörper, einer Dateianlage, einer Elementanlage und einem Callback.</span><span class="sxs-lookup"><span data-stu-id="d9554-928">Reply with a body, file attachment, item attachment, and a callback.</span></span>

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

#### <a name="getentities--entitiesjavascriptapioutlook17officeentities"></a><span data-ttu-id="d9554-929">getEntities() → {[Entities](/javascript/api/outlook_1_7/office.entities)}</span><span class="sxs-lookup"><span data-stu-id="d9554-929">getEntities() → {[Entities](/javascript/api/outlook_1_7/office.entities)}</span></span>

<span data-ttu-id="d9554-930">Ruft das ausgewählte Element Textkörper gefundenen Entitäten ab.</span><span class="sxs-lookup"><span data-stu-id="d9554-930">Gets the entities found in the selected item's body.</span></span>

> [!NOTE]
> <span data-ttu-id="d9554-931">Diese Methode wird in Outlook für iOS oder Outlook für Android nicht unterstützt.</span><span class="sxs-lookup"><span data-stu-id="d9554-931">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

##### <a name="requirements"></a><span data-ttu-id="d9554-932">Anforderungen</span><span class="sxs-lookup"><span data-stu-id="d9554-932">Requirements</span></span>

|<span data-ttu-id="d9554-933">Anforderung</span><span class="sxs-lookup"><span data-stu-id="d9554-933">Requirement</span></span>|<span data-ttu-id="d9554-934">Wert</span><span class="sxs-lookup"><span data-stu-id="d9554-934">Value</span></span>|
|---|---|
|[<span data-ttu-id="d9554-935">Mindestversion des Postfachanforderungssatzes</span><span class="sxs-lookup"><span data-stu-id="d9554-935">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="d9554-936">1.0</span><span class="sxs-lookup"><span data-stu-id="d9554-936">1.0</span></span>|
|[<span data-ttu-id="d9554-937">Mindestberechtigungsstufe</span><span class="sxs-lookup"><span data-stu-id="d9554-937">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="d9554-938">ReadItem</span><span class="sxs-lookup"><span data-stu-id="d9554-938">ReadItem</span></span>|
|[<span data-ttu-id="d9554-939">Zutreffender Outlook-Modus</span><span class="sxs-lookup"><span data-stu-id="d9554-939">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="d9554-940">Lesen</span><span class="sxs-lookup"><span data-stu-id="d9554-940">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="d9554-941">Gibt zurück:</span><span class="sxs-lookup"><span data-stu-id="d9554-941">Returns:</span></span>

<span data-ttu-id="d9554-942">Typ: [Entitäten](/javascript/api/outlook_1_7/office.entities)</span><span class="sxs-lookup"><span data-stu-id="d9554-942">Type: [Entities](/javascript/api/outlook_1_7/office.entities)</span></span>

##### <a name="example"></a><span data-ttu-id="d9554-943">Beispiel</span><span class="sxs-lookup"><span data-stu-id="d9554-943">Example</span></span>

<span data-ttu-id="d9554-944">Das folgende Beispiel greift auf die Kontakte Entitäten im Textkörper des aktuellen Elements.</span><span class="sxs-lookup"><span data-stu-id="d9554-944">The following example accesses the contacts entities in the current item's body.</span></span>

```
var contacts = Office.context.mailbox.item.getEntities().contacts;
```

#### <a name="getentitiesbytypeentitytype--nullable-arraystringcontactjavascriptapioutlook17officecontactmeetingsuggestionjavascriptapioutlook17officemeetingsuggestionphonenumberjavascriptapioutlook17officephonenumbertasksuggestionjavascriptapioutlook17officetasksuggestion"></a><span data-ttu-id="d9554-945">getEntitiesByType(entityType) → (nullable) {Array.<(String|[Contact](/javascript/api/outlook_1_7/office.contact)|[MeetingSuggestion](/javascript/api/outlook_1_7/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook_1_7/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook_1_7/office.tasksuggestion))>}</span><span class="sxs-lookup"><span data-stu-id="d9554-945">getEntitiesByType(entityType) → (nullable) {Array.<(String|[Contact](/javascript/api/outlook_1_7/office.contact)|[MeetingSuggestion](/javascript/api/outlook_1_7/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook_1_7/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook_1_7/office.tasksuggestion))>}</span></span>

<span data-ttu-id="d9554-946">Ruft ein Array aller Entitäten des angegebenen Entitätstyps im Hauptteil des ausgewählten Elements gefunden.</span><span class="sxs-lookup"><span data-stu-id="d9554-946">Gets an array of all the entities of the specified entity type found in the selected item's body.</span></span>

> [!NOTE]
> <span data-ttu-id="d9554-947">Diese Methode wird in Outlook für iOS oder Outlook für Android nicht unterstützt.</span><span class="sxs-lookup"><span data-stu-id="d9554-947">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

##### <a name="parameters"></a><span data-ttu-id="d9554-948">Parameter:</span><span class="sxs-lookup"><span data-stu-id="d9554-948">Parameters:</span></span>

|<span data-ttu-id="d9554-949">Name</span><span class="sxs-lookup"><span data-stu-id="d9554-949">Name</span></span>|<span data-ttu-id="d9554-950">Typ</span><span class="sxs-lookup"><span data-stu-id="d9554-950">Type</span></span>|<span data-ttu-id="d9554-951">Beschreibung</span><span class="sxs-lookup"><span data-stu-id="d9554-951">Description</span></span>|
|---|---|---|
|`entityType`|[<span data-ttu-id="d9554-952">Office.MailboxEnums.EntityType</span><span class="sxs-lookup"><span data-stu-id="d9554-952">Office.MailboxEnums.EntityType</span></span>](/javascript/api/outlook_1_7/office.mailboxenums.entitytype)|<span data-ttu-id="d9554-953">Einer der Werte der EntityType-Enumeration.</span><span class="sxs-lookup"><span data-stu-id="d9554-953">One of the EntityType enumeration values.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="d9554-954">Anforderungen</span><span class="sxs-lookup"><span data-stu-id="d9554-954">Requirements</span></span>

|<span data-ttu-id="d9554-955">Anforderung</span><span class="sxs-lookup"><span data-stu-id="d9554-955">Requirement</span></span>|<span data-ttu-id="d9554-956">Wert</span><span class="sxs-lookup"><span data-stu-id="d9554-956">Value</span></span>|
|---|---|
|[<span data-ttu-id="d9554-957">Mindestversion des Postfachanforderungssatzes</span><span class="sxs-lookup"><span data-stu-id="d9554-957">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="d9554-958">1.0</span><span class="sxs-lookup"><span data-stu-id="d9554-958">1.0</span></span>|
|[<span data-ttu-id="d9554-959">Mindestberechtigungsstufe</span><span class="sxs-lookup"><span data-stu-id="d9554-959">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="d9554-960">Eingeschränkt</span><span class="sxs-lookup"><span data-stu-id="d9554-960">Restricted</span></span>|
|[<span data-ttu-id="d9554-961">Zutreffender Outlook-Modus</span><span class="sxs-lookup"><span data-stu-id="d9554-961">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="d9554-962">Lesen</span><span class="sxs-lookup"><span data-stu-id="d9554-962">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="d9554-963">Gibt zurück:</span><span class="sxs-lookup"><span data-stu-id="d9554-963">Returns:</span></span>

<span data-ttu-id="d9554-964">Wenn der in `entityType` übergebene Wert kein gültiges Element der `EntityType`-Enumeration ist, gibt die Methode null zurück.</span><span class="sxs-lookup"><span data-stu-id="d9554-964">If the value passed in `entityType` is not a valid member of the `EntityType` enumeration, the method returns null.</span></span> <span data-ttu-id="d9554-965">Wenn keine Entitäten des angegebenen Typs im Textkörper des Elements vorhanden sind, gibt die Methode ein leeres Array zurück.</span><span class="sxs-lookup"><span data-stu-id="d9554-965">If no entities of the specified type are present in the item's body, the method returns an empty array.</span></span> <span data-ttu-id="d9554-966">Andernfalls hängt der Typ der Objekte im zurückgegebenen Array vom Typ der Entität ab, die im `entityType`-Parameter angefordert wurde.</span><span class="sxs-lookup"><span data-stu-id="d9554-966">Otherwise, the type of the objects in the returned array depends on the type of entity requested in the `entityType` parameter.</span></span>

<span data-ttu-id="d9554-967">Während Sie die minimale Berechtigungsstufe für diese Methode **Restricted** ist, erfordern einige Entitätstypen **ReadItem** für den Zugriff, wie in der folgenden Tabelle angegeben wird.</span><span class="sxs-lookup"><span data-stu-id="d9554-967">While the minimum permission level to use this method is **Restricted**, some entity types require **ReadItem** to access, as specified in the following table.</span></span>

|<span data-ttu-id="d9554-968">Wert von `entityType`</span><span class="sxs-lookup"><span data-stu-id="d9554-968">Value of `entityType`</span></span>|<span data-ttu-id="d9554-969">Typ der Objekte im zurückgegebenen Array</span><span class="sxs-lookup"><span data-stu-id="d9554-969">Type of objects in returned array</span></span>|<span data-ttu-id="d9554-970">Erforderliche Berechtigungsstufe</span><span class="sxs-lookup"><span data-stu-id="d9554-970">Required Permission Level</span></span>|
|---|---|---|
|`Address`|<span data-ttu-id="d9554-971">Zeichenfolge</span><span class="sxs-lookup"><span data-stu-id="d9554-971">String</span></span>|<span data-ttu-id="d9554-972">**Restricted**</span><span class="sxs-lookup"><span data-stu-id="d9554-972">**Restricted**</span></span>|
|`Contact`|<span data-ttu-id="d9554-973">Contact</span><span class="sxs-lookup"><span data-stu-id="d9554-973">Contact</span></span>|<span data-ttu-id="d9554-974">**ReadItem**</span><span class="sxs-lookup"><span data-stu-id="d9554-974">**ReadItem**</span></span>|
|`EmailAddress`|<span data-ttu-id="d9554-975">Zeichenfolge</span><span class="sxs-lookup"><span data-stu-id="d9554-975">String</span></span>|<span data-ttu-id="d9554-976">**ReadItem**</span><span class="sxs-lookup"><span data-stu-id="d9554-976">**ReadItem**</span></span>|
|`MeetingSuggestion`|<span data-ttu-id="d9554-977">MeetingSuggestion</span><span class="sxs-lookup"><span data-stu-id="d9554-977">MeetingSuggestion</span></span>|<span data-ttu-id="d9554-978">**ReadItem**</span><span class="sxs-lookup"><span data-stu-id="d9554-978">**ReadItem**</span></span>|
|`PhoneNumber`|<span data-ttu-id="d9554-979">PhoneNumber</span><span class="sxs-lookup"><span data-stu-id="d9554-979">PhoneNumber</span></span>|<span data-ttu-id="d9554-980">**Restricted**</span><span class="sxs-lookup"><span data-stu-id="d9554-980">**Restricted**</span></span>|
|`TaskSuggestion`|<span data-ttu-id="d9554-981">TaskSuggestion</span><span class="sxs-lookup"><span data-stu-id="d9554-981">TaskSuggestion</span></span>|<span data-ttu-id="d9554-982">**ReadItem**</span><span class="sxs-lookup"><span data-stu-id="d9554-982">**ReadItem**</span></span>|
|`URL`|<span data-ttu-id="d9554-983">Zeichenfolge</span><span class="sxs-lookup"><span data-stu-id="d9554-983">String</span></span>|<span data-ttu-id="d9554-984">**Restricted**</span><span class="sxs-lookup"><span data-stu-id="d9554-984">**Restricted**</span></span>|

<span data-ttu-id="d9554-985">Typ: Array.<(String|[Contact](/javascript/api/outlook_1_7/office.contact)|[MeetingSuggestion](/javascript/api/outlook_1_7/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook_1_7/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook_1_7/office.tasksuggestion))></span><span class="sxs-lookup"><span data-stu-id="d9554-985">Type: Array.<(String|[Contact](/javascript/api/outlook_1_7/office.contact)|[MeetingSuggestion](/javascript/api/outlook_1_7/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook_1_7/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook_1_7/office.tasksuggestion))></span></span>

##### <a name="example"></a><span data-ttu-id="d9554-986">Beispiel</span><span class="sxs-lookup"><span data-stu-id="d9554-986">Example</span></span>

<span data-ttu-id="d9554-987">Im folgenden Beispiel wird veranschaulicht, wie auf ein Array von Zeichenfolgen, die Postadressen im Textkörper des aktuellen Elements darstellen.</span><span class="sxs-lookup"><span data-stu-id="d9554-987">The following example shows how to access an array of strings that represent postal addresses in the current item's body.</span></span>

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

#### <a name="getfilteredentitiesbynamename--nullable-arraystringcontactjavascriptapioutlook17officecontactmeetingsuggestionjavascriptapioutlook17officemeetingsuggestionphonenumberjavascriptapioutlook17officephonenumbertasksuggestionjavascriptapioutlook17officetasksuggestion"></a><span data-ttu-id="d9554-988">getFilteredEntitiesByName(name) → (nullable) {Array.<(String|[Contact](/javascript/api/outlook_1_7/office.contact)|[MeetingSuggestion](/javascript/api/outlook_1_7/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook_1_7/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook_1_7/office.tasksuggestion))>}</span><span class="sxs-lookup"><span data-stu-id="d9554-988">getFilteredEntitiesByName(name) → (nullable) {Array.<(String|[Contact](/javascript/api/outlook_1_7/office.contact)|[MeetingSuggestion](/javascript/api/outlook_1_7/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook_1_7/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook_1_7/office.tasksuggestion))>}</span></span>

<span data-ttu-id="d9554-989">Gibt bekannte Entitäten im ausgewählten Element zurück, die den in der XML-Manifestdatei definierten benannten Filter übergeben.</span><span class="sxs-lookup"><span data-stu-id="d9554-989">Returns well-known entities in the selected item that pass the named filter defined in the manifest XML file.</span></span>

> [!NOTE]
> <span data-ttu-id="d9554-990">Diese Methode wird in Outlook für iOS oder Outlook für Android nicht unterstützt.</span><span class="sxs-lookup"><span data-stu-id="d9554-990">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="d9554-991">Die `getFilteredEntitiesByName`-Methode gibt die Entitäten zurück, die dem im [ItemHasKnownEntity](/javascript/office/manifest/rule#itemhasknownentity-rule)-Regelelement der XML-Manifestdatei definierten regulären Ausdruck mit dem angegebenen `FilterName`-Elementwert entsprechen.</span><span class="sxs-lookup"><span data-stu-id="d9554-991">The `getFilteredEntitiesByName` method returns the entities that match the regular expression defined in the [ItemHasKnownEntity](/javascript/office/manifest/rule#itemhasknownentity-rule) rule element in the manifest XML file with the specified `FilterName` element value.</span></span>

##### <a name="parameters"></a><span data-ttu-id="d9554-992">Parameter:</span><span class="sxs-lookup"><span data-stu-id="d9554-992">Parameters:</span></span>

|<span data-ttu-id="d9554-993">Name</span><span class="sxs-lookup"><span data-stu-id="d9554-993">Name</span></span>|<span data-ttu-id="d9554-994">Typ</span><span class="sxs-lookup"><span data-stu-id="d9554-994">Type</span></span>|<span data-ttu-id="d9554-995">Beschreibung</span><span class="sxs-lookup"><span data-stu-id="d9554-995">Description</span></span>|
|---|---|---|
|`name`|<span data-ttu-id="d9554-996">Zeichenfolge</span><span class="sxs-lookup"><span data-stu-id="d9554-996">String</span></span>|<span data-ttu-id="d9554-997">Der Name des `ItemHasKnownEntity`-Regelelements, das den entsprechenden Filter definiert.</span><span class="sxs-lookup"><span data-stu-id="d9554-997">The name of the `ItemHasKnownEntity` rule element that defines the filter to match.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="d9554-998">Anforderungen</span><span class="sxs-lookup"><span data-stu-id="d9554-998">Requirements</span></span>

|<span data-ttu-id="d9554-999">Anforderung</span><span class="sxs-lookup"><span data-stu-id="d9554-999">Requirement</span></span>|<span data-ttu-id="d9554-1000">Wert</span><span class="sxs-lookup"><span data-stu-id="d9554-1000">Value</span></span>|
|---|---|
|[<span data-ttu-id="d9554-1001">Mindestversion des Postfachanforderungssatzes</span><span class="sxs-lookup"><span data-stu-id="d9554-1001">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="d9554-1002">1.0</span><span class="sxs-lookup"><span data-stu-id="d9554-1002">1.0</span></span>|
|[<span data-ttu-id="d9554-1003">Mindestberechtigungsstufe</span><span class="sxs-lookup"><span data-stu-id="d9554-1003">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="d9554-1004">ReadItem</span><span class="sxs-lookup"><span data-stu-id="d9554-1004">ReadItem</span></span>|
|[<span data-ttu-id="d9554-1005">Zutreffender Outlook-Modus</span><span class="sxs-lookup"><span data-stu-id="d9554-1005">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="d9554-1006">Lesen</span><span class="sxs-lookup"><span data-stu-id="d9554-1006">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="d9554-1007">Gibt zurück:</span><span class="sxs-lookup"><span data-stu-id="d9554-1007">Returns:</span></span>

<span data-ttu-id="d9554-p158">Befindet sich kein `ItemHasKnownEntity`-Element im Manifest mit einem `FilterName`-Elementwert, der dem `name`-Parameter entspricht, gibt die Methode `null` zurück. Wenn der `name`-Parameter einem `ItemHasKnownEntity`-Element im Manifest entspricht, es aber keine entsprechenden Entitäten im aktuellen Element gibt, gibt die Methode ein leeres Array zurück.</span><span class="sxs-lookup"><span data-stu-id="d9554-p158">If there is no `ItemHasKnownEntity` element in the manifest with a `FilterName` element value that matches the `name` parameter, the method returns `null`. If the `name` parameter does match an `ItemHasKnownEntity` element in the manifest, but there are no entities in the current item that match, the method return an empty array.</span></span>

<span data-ttu-id="d9554-1010">Typ: Array.<(String|[Contact](/javascript/api/outlook_1_7/office.contact)|[MeetingSuggestion](/javascript/api/outlook_1_7/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook_1_7/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook_1_7/office.tasksuggestion))></span><span class="sxs-lookup"><span data-stu-id="d9554-1010">Type: Array.<(String|[Contact](/javascript/api/outlook_1_7/office.contact)|[MeetingSuggestion](/javascript/api/outlook_1_7/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook_1_7/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook_1_7/office.tasksuggestion))></span></span>

#### <a name="getregexmatches--object"></a><span data-ttu-id="d9554-1011">getRegExMatches() → {Object}</span><span class="sxs-lookup"><span data-stu-id="d9554-1011">getRegExMatches() → {Object}</span></span>

<span data-ttu-id="d9554-1012">Gibt Zeichenfolgenwerte im ausgewählten Element zurück, die den in der XML-Manifestdatei definierten regulären Ausdrücken entsprechen.</span><span class="sxs-lookup"><span data-stu-id="d9554-1012">Returns string values in the selected item that match the regular expressions defined in the manifest XML file.</span></span>

> [!NOTE]
> <span data-ttu-id="d9554-1013">Diese Methode wird in Outlook für iOS oder Outlook für Android nicht unterstützt.</span><span class="sxs-lookup"><span data-stu-id="d9554-1013">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="d9554-p159">Die `getRegExMatches`-Methode gibt die Zeichenfolgen zurück, die dem im `ItemHasRegularExpressionMatch`- oder `ItemHasKnownEntity`-Regelelement der XML-Manifestdatei definierten regulären Ausdruck entsprechen. Bei einer `ItemHasRegularExpressionMatch`-Regel muss eine entsprechende Zeichenfolge in der Eigenschaft des Elements vorhanden sein, das von dieser Regel angegeben wird. Der einfache `PropertyName`-Typ definiert die unterstützten Eigenschaften.</span><span class="sxs-lookup"><span data-stu-id="d9554-p159">The `getRegExMatches` method returns the strings that match the regular expression defined in each `ItemHasRegularExpressionMatch` or `ItemHasKnownEntity` rule element in the manifest XML file. For an `ItemHasRegularExpressionMatch` rule, a matching string has to occur in the property of the item that is specified by that rule. The `PropertyName` simple type defines the supported properties.</span></span>

<span data-ttu-id="d9554-1017">Nehmen Sie z. B. an, dass ein Add-In-Manifest das folgende `Rule`-Element aufweist:</span><span class="sxs-lookup"><span data-stu-id="d9554-1017">For example, consider an add-in manifest has the following `Rule` element:</span></span>

```
<Rule xsi:type="RuleCollection" Mode="And">
  <Rule xsi:type="ItemIs" FormType="Read" ItemType="Message" />
  <Rule xsi:type="RuleCollection" Mode="Or">
    <Rule xsi:type="ItemHasRegularExpressionMatch" RegExName="fruits" RegExValue="apple|banana|coconut" PropertyName="BodyAsPlaintext" IgnoreCase="true" />
    <Rule xsi:type="ItemHasRegularExpressionMatch" RegExName="veggies" RegExValue="tomato|onion|spinach|broccoli" PropertyName="BodyAsPlaintext" IgnoreCase="true" />
  </Rule>
</Rule>
```

<span data-ttu-id="d9554-1018">Das von `getRegExMatches` zurückgegebene Objekt hätte zwei Eigenschaften: `fruits` und `veggies`.</span><span class="sxs-lookup"><span data-stu-id="d9554-1018">The object returned from `getRegExMatches` would have two properties: `fruits` and `veggies`.</span></span>

```
{
  'fruits': ['apple','banana','Banana','coconut'],
  'veggies': ['tomato','onion','spinach','broccoli']
}
```

<span data-ttu-id="d9554-p160">Wenn Sie eine `ItemHasRegularExpressionMatch`-Regel für die Textkörpereigenschaft eines Elements festlegen, sollte der reguläre Ausdruck den Textkörper weiter filtern und nicht versuchen, den gesamten Textkörper des Elements zurückzugeben. Wenn der gesamte Textkörper eines Elements mit einem regulären Ausdruck wie `.*` abgerufen wird, werden nicht immer die gewünschten Ergebnisse erzielt. Verwenden Sie stattdessen die [`Body.getAsync`](/javascript/api/outlook_1_7/office.body#getasync-coerciontype--options--callback-)-Methode, um den gesamten Textkörper abzurufen.</span><span class="sxs-lookup"><span data-stu-id="d9554-p160">If you specify an `ItemHasRegularExpressionMatch` rule on the body property of an item, the regular expression should further filter the body and should not attempt to return the entire body of the item. Using a regular expression such as `.*` to obtain the entire body of an item does not always return the expected results. Instead, use the [`Body.getAsync`](/javascript/api/outlook_1_7/office.body#getasync-coerciontype--options--callback-) method to retrieve the entire body.</span></span>

##### <a name="requirements"></a><span data-ttu-id="d9554-1022">Anforderungen</span><span class="sxs-lookup"><span data-stu-id="d9554-1022">Requirements</span></span>

|<span data-ttu-id="d9554-1023">Anforderung</span><span class="sxs-lookup"><span data-stu-id="d9554-1023">Requirement</span></span>|<span data-ttu-id="d9554-1024">Wert</span><span class="sxs-lookup"><span data-stu-id="d9554-1024">Value</span></span>|
|---|---|
|[<span data-ttu-id="d9554-1025">Mindestversion des Postfachanforderungssatzes</span><span class="sxs-lookup"><span data-stu-id="d9554-1025">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="d9554-1026">1.0</span><span class="sxs-lookup"><span data-stu-id="d9554-1026">1.0</span></span>|
|[<span data-ttu-id="d9554-1027">Mindestberechtigungsstufe</span><span class="sxs-lookup"><span data-stu-id="d9554-1027">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="d9554-1028">ReadItem</span><span class="sxs-lookup"><span data-stu-id="d9554-1028">ReadItem</span></span>|
|[<span data-ttu-id="d9554-1029">Zutreffender Outlook-Modus</span><span class="sxs-lookup"><span data-stu-id="d9554-1029">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="d9554-1030">Lesen</span><span class="sxs-lookup"><span data-stu-id="d9554-1030">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="d9554-1031">Gibt zurück:</span><span class="sxs-lookup"><span data-stu-id="d9554-1031">Returns:</span></span>

<span data-ttu-id="d9554-p161">Ein Objekt mit Arrays aus Zeichenfolgen, die den in der XML-Manifestdatei definierten regulären Ausdrücken entsprechen. Der Name der einzelnen Arrays ist gleich dem entsprechenden Wert des `RegExName`-Attributs der entsprechenden `ItemHasRegularExpressionMatch`-Regel oder des `FilterName`-Attributs der entsprechenden `ItemHasKnownEntity`-Regel.</span><span class="sxs-lookup"><span data-stu-id="d9554-p161">An object that contains arrays of strings that match the regular expressions defined in the manifest XML file. The name of each array is equal to the corresponding value of the `RegExName` attribute of the matching `ItemHasRegularExpressionMatch` rule or the `FilterName` attribute of the matching `ItemHasKnownEntity` rule.</span></span>

<dl class="param-type"><span data-ttu-id="d9554-1034">

<dt>
Typ</dt>


</span><span class="sxs-lookup"><span data-stu-id="d9554-1034">

<dt>Type</dt>

</span></span><dd><span data-ttu-id="d9554-1035">Object</span><span class="sxs-lookup"><span data-stu-id="d9554-1035">Object</span></span></dd>

</dl>

##### <a name="example"></a><span data-ttu-id="d9554-1036">Beispiel</span><span class="sxs-lookup"><span data-stu-id="d9554-1036">Example</span></span>

<span data-ttu-id="d9554-1037">Das folgende Beispiel zeigt, wie Sie auf das Array von Übereinstimmungen für die regulären Ausdrucksregelelemente `fruits` und `veggies` zugreifen, die im Manifest angegeben sind.</span><span class="sxs-lookup"><span data-stu-id="d9554-1037">The following example shows how to access the array of matches for the regular expression rule elements `fruits` and `veggies`, which are specified in the manifest.</span></span>

```
var allMatches = Office.context.mailbox.item.getRegExMatches();
var fruits = allMatches.fruits;
var veges = allMatches.veggies;
```

#### <a name="getregexmatchesbynamename--nullable-array-string-"></a><span data-ttu-id="d9554-1038">getRegExMatchesByName(name) → (Nullwerte zulassen) {Array. < Zeichenfolge >}</span><span class="sxs-lookup"><span data-stu-id="d9554-1038">getRegExMatchesByName(name) → (nullable) {Array.< String >}</span></span>

<span data-ttu-id="d9554-1039">Gibt Zeichenfolgenwerte im ausgewählten Element zurück, die dem in der XML-Manifestdatei definierten benannten regulären Ausdruck entsprechen.</span><span class="sxs-lookup"><span data-stu-id="d9554-1039">Returns string values in the selected item that match the named regular expression defined in the manifest XML file.</span></span>

> [!NOTE]
> <span data-ttu-id="d9554-1040">Diese Methode wird in Outlook für iOS oder Outlook für Android nicht unterstützt.</span><span class="sxs-lookup"><span data-stu-id="d9554-1040">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="d9554-1041">Die `getRegExMatchesByName`-Methode gibt die Zeichenfolgen zurück, die dem im `ItemHasRegularExpressionMatch`-Regelelement der XML-Manifestdatei definierten regulären Ausdruck mit dem angegebenen `RegExName`-Elementwert entsprechen.</span><span class="sxs-lookup"><span data-stu-id="d9554-1041">The `getRegExMatchesByName` method returns the strings that match the regular expression defined in the `ItemHasRegularExpressionMatch` rule element in the manifest XML file with the specified `RegExName` element value.</span></span>

<span data-ttu-id="d9554-p162">Wenn Sie eine `ItemHasRegularExpressionMatch`-Regel für die Textkörpereigenschaft eines Elements festlegen, sollte der reguläre Ausdruck den Textkörper weiter filtern und nicht versuchen, den gesamten Textkörper des Elements zurückzugeben. Wenn der gesamte Textkörper eines Elements mit einem regulären Ausdruck wie `.*` abgerufen wird, werden nicht immer die gewünschten Ergebnisse erzielt.</span><span class="sxs-lookup"><span data-stu-id="d9554-p162">If you specify an `ItemHasRegularExpressionMatch` rule on the body property of an item, the regular expression should further filter the body and should not attempt to return the entire body of the item. Using a regular expression such as `.*` to obtain the entire body of an item does not always return the expected results.</span></span>

##### <a name="parameters"></a><span data-ttu-id="d9554-1044">Parameter:</span><span class="sxs-lookup"><span data-stu-id="d9554-1044">Parameters:</span></span>

|<span data-ttu-id="d9554-1045">Name</span><span class="sxs-lookup"><span data-stu-id="d9554-1045">Name</span></span>|<span data-ttu-id="d9554-1046">Typ</span><span class="sxs-lookup"><span data-stu-id="d9554-1046">Type</span></span>|<span data-ttu-id="d9554-1047">Beschreibung</span><span class="sxs-lookup"><span data-stu-id="d9554-1047">Description</span></span>|
|---|---|---|
|`name`|<span data-ttu-id="d9554-1048">Zeichenfolge</span><span class="sxs-lookup"><span data-stu-id="d9554-1048">String</span></span>|<span data-ttu-id="d9554-1049">Der Name des `ItemHasRegularExpressionMatch`-Regelelements, das den entsprechenden Filter definiert.</span><span class="sxs-lookup"><span data-stu-id="d9554-1049">The name of the `ItemHasRegularExpressionMatch` rule element that defines the filter to match.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="d9554-1050">Anforderungen</span><span class="sxs-lookup"><span data-stu-id="d9554-1050">Requirements</span></span>

|<span data-ttu-id="d9554-1051">Anforderung</span><span class="sxs-lookup"><span data-stu-id="d9554-1051">Requirement</span></span>|<span data-ttu-id="d9554-1052">Wert</span><span class="sxs-lookup"><span data-stu-id="d9554-1052">Value</span></span>|
|---|---|
|[<span data-ttu-id="d9554-1053">Mindestversion des Postfachanforderungssatzes</span><span class="sxs-lookup"><span data-stu-id="d9554-1053">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="d9554-1054">1.0</span><span class="sxs-lookup"><span data-stu-id="d9554-1054">1.0</span></span>|
|[<span data-ttu-id="d9554-1055">Mindestberechtigungsstufe</span><span class="sxs-lookup"><span data-stu-id="d9554-1055">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="d9554-1056">ReadItem</span><span class="sxs-lookup"><span data-stu-id="d9554-1056">ReadItem</span></span>|
|[<span data-ttu-id="d9554-1057">Zutreffender Outlook-Modus</span><span class="sxs-lookup"><span data-stu-id="d9554-1057">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="d9554-1058">Lesen</span><span class="sxs-lookup"><span data-stu-id="d9554-1058">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="d9554-1059">Gibt zurück:</span><span class="sxs-lookup"><span data-stu-id="d9554-1059">Returns:</span></span>

<span data-ttu-id="d9554-1060">Ein Array mit den Zeichenfolgen, die dem in der XML-Manifestdatei definierten regulären Ausdruck entsprechen.</span><span class="sxs-lookup"><span data-stu-id="d9554-1060">An array that contains the strings that match the regular expression defined in the manifest XML file.</span></span>

<dl class="param-type"><span data-ttu-id="d9554-1061">

<dt>Typ</dt>

</span><span class="sxs-lookup"><span data-stu-id="d9554-1061">

<dt>Type</dt>

</span></span><dd><span data-ttu-id="d9554-1062">Array. < Zeichenfolge ></span><span class="sxs-lookup"><span data-stu-id="d9554-1062">Array.< String ></span></span></dd>

</dl>

##### <a name="example"></a><span data-ttu-id="d9554-1063">Beispiel</span><span class="sxs-lookup"><span data-stu-id="d9554-1063">Example</span></span>

```
var fruits = Office.context.mailbox.item.getRegExMatchesByName("fruits");
var veggies = Office.context.mailbox.item.getRegExMatchesByName("veggies");
```

####  <a name="getselecteddataasynccoerciontype-options-callback--string"></a><span data-ttu-id="d9554-1064">getSelectedDataAsync(coercionType, [options], callback) → {String}</span><span class="sxs-lookup"><span data-stu-id="d9554-1064">getSelectedDataAsync(coercionType, [options], callback) → {String}</span></span>

<span data-ttu-id="d9554-1065">Gibt asynchron ausgewählte Daten aus dem Betreff oder Textkörper einer Nachricht zurück.</span><span class="sxs-lookup"><span data-stu-id="d9554-1065">Asynchronously returns selected data from the subject or body of a message.</span></span>

<span data-ttu-id="d9554-p163">Wenn keine Auswahl vorhanden ist, aber der Cursor sich im Nachrichtentext oder Betreff befindet, gibt die Methode für die ausgewählten Daten NULL zurück. Wenn ein anderes Feld als der Textkörper oder Betreff ausgewählt ist, gibt die Methode den `InvalidSelection`-Fehler zurück.</span><span class="sxs-lookup"><span data-stu-id="d9554-p163">If there is no selection but the cursor is in the body or subject, the method returns null for the selected data. If a field other than the body or subject is selected, the method returns the `InvalidSelection` error.</span></span>

##### <a name="parameters"></a><span data-ttu-id="d9554-1068">Parameter:</span><span class="sxs-lookup"><span data-stu-id="d9554-1068">Parameters:</span></span>

|<span data-ttu-id="d9554-1069">Name</span><span class="sxs-lookup"><span data-stu-id="d9554-1069">Name</span></span>|<span data-ttu-id="d9554-1070">Typ</span><span class="sxs-lookup"><span data-stu-id="d9554-1070">Type</span></span>|<span data-ttu-id="d9554-1071">Attribute</span><span class="sxs-lookup"><span data-stu-id="d9554-1071">Attributes</span></span>|<span data-ttu-id="d9554-1072">Beschreibung</span><span class="sxs-lookup"><span data-stu-id="d9554-1072">Description</span></span>|
|---|---|---|---|
|`coercionType`|[<span data-ttu-id="d9554-1073">Office.CoercionType</span><span class="sxs-lookup"><span data-stu-id="d9554-1073">Office.CoercionType</span></span>](office.md#coerciontype-string)||<span data-ttu-id="d9554-p164">Fordert ein Format für die Daten an. Wenn es sich um Texthandelt, gibt die Methode den unformatierten Text als Zeichenfolge zurück und entfernt eventuell vorhandene HTML-Tags. Wenn es sich um HTML handelt, gibt die Methode den ausgewählten Text zurück, entweder als unformatierten Text oder als HTML.</span><span class="sxs-lookup"><span data-stu-id="d9554-p164">Requests a format for the data. If Text, the method returns the plain text as a string , removing any HTML tags present. If HTML, the method returns the selected text, whether it is plaintext or HTML.</span></span>|
|`options`|<span data-ttu-id="d9554-1077">Object</span><span class="sxs-lookup"><span data-stu-id="d9554-1077">Object</span></span>|<span data-ttu-id="d9554-1078">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="d9554-1078">&lt;optional&gt;</span></span>|<span data-ttu-id="d9554-1079">Ein Objektliteral, das mindestens eine der folgenden Eigenschaften enthält.</span><span class="sxs-lookup"><span data-stu-id="d9554-1079">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`|<span data-ttu-id="d9554-1080">Object</span><span class="sxs-lookup"><span data-stu-id="d9554-1080">Object</span></span>|<span data-ttu-id="d9554-1081">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="d9554-1081">&lt;optional&gt;</span></span>|<span data-ttu-id="d9554-1082">Entwickler können ein Objekt bereitstellen, auf das sie in der Callbackmethode zugreifen möchten.</span><span class="sxs-lookup"><span data-stu-id="d9554-1082">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`|<span data-ttu-id="d9554-1083">Funktion</span><span class="sxs-lookup"><span data-stu-id="d9554-1083">function</span></span>||<span data-ttu-id="d9554-1084">Wenn die Methode abgeschlossen wird, wird die Funktion , die im `callback`-Parameter übergeben wird, mit einem einzelnen Parameter (`asyncResult`) aufgerufen, der ein [`AsyncResult`](/javascript/api/office/office.asyncresult)-Objekt darstellt.</span><span class="sxs-lookup"><span data-stu-id="d9554-1084">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span><br/><br/><span data-ttu-id="d9554-1085">Rufen Sie für den Zugriff auf die ausgewählten Daten aus der Rückrufmethode `asyncResult.value.data` auf.</span><span class="sxs-lookup"><span data-stu-id="d9554-1085">To access the selected data from the callback method, call `asyncResult.value.data`.</span></span> <span data-ttu-id="d9554-1086">Rufen Sie die Source-Eigenschaft für den Zugriff, die die Auswahl stammen, `asyncResult.value.sourceProperty`, die entweder sein `body` oder `subject`.</span><span class="sxs-lookup"><span data-stu-id="d9554-1086">To access the source property that the selection comes from, call `asyncResult.value.sourceProperty`, which will be either `body` or `subject`.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="d9554-1087">Anforderungen</span><span class="sxs-lookup"><span data-stu-id="d9554-1087">Requirements</span></span>

|<span data-ttu-id="d9554-1088">Anforderung</span><span class="sxs-lookup"><span data-stu-id="d9554-1088">Requirement</span></span>|<span data-ttu-id="d9554-1089">Wert</span><span class="sxs-lookup"><span data-stu-id="d9554-1089">Value</span></span>|
|---|---|
|[<span data-ttu-id="d9554-1090">Mindestversion des Postfachanforderungssatzes</span><span class="sxs-lookup"><span data-stu-id="d9554-1090">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="d9554-1091">1.2</span><span class="sxs-lookup"><span data-stu-id="d9554-1091">1.2</span></span>|
|[<span data-ttu-id="d9554-1092">Mindestberechtigungsstufe</span><span class="sxs-lookup"><span data-stu-id="d9554-1092">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="d9554-1093">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="d9554-1093">ReadWriteItem</span></span>|
|[<span data-ttu-id="d9554-1094">Zutreffender Outlook-Modus</span><span class="sxs-lookup"><span data-stu-id="d9554-1094">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="d9554-1095">Verfassen</span><span class="sxs-lookup"><span data-stu-id="d9554-1095">Compose</span></span>|

##### <a name="returns"></a><span data-ttu-id="d9554-1096">Gibt zurück:</span><span class="sxs-lookup"><span data-stu-id="d9554-1096">Returns:</span></span>

<span data-ttu-id="d9554-1097">Die ausgewählten Daten als Zeichenfolge mit dem durch `coercionType` bestimmten Format.</span><span class="sxs-lookup"><span data-stu-id="d9554-1097">The selected data as a string with format determined by `coercionType`.</span></span>

<dl class="param-type"><span data-ttu-id="d9554-1098">

<dt>
Typ</dt>


</span><span class="sxs-lookup"><span data-stu-id="d9554-1098">

<dt>Type</dt>

</span></span><dd><span data-ttu-id="d9554-1099">Zeichenfolge</span><span class="sxs-lookup"><span data-stu-id="d9554-1099">String</span></span></dd>

</dl>

##### <a name="example"></a><span data-ttu-id="d9554-1100">Beispiel</span><span class="sxs-lookup"><span data-stu-id="d9554-1100">Example</span></span>

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

#### <a name="getselectedentities--entitiesjavascriptapioutlook17officeentities"></a><span data-ttu-id="d9554-1101">getSelectedEntities() → {[Entitäten](/javascript/api/outlook_1_7/office.entities)}</span><span class="sxs-lookup"><span data-stu-id="d9554-1101">getSelectedEntities() → {[Entities](/javascript/api/outlook_1_7/office.entities)}</span></span>

<span data-ttu-id="d9554-p166">Ruft die Entitäten ab, die in einer hervorgehobenen Übereinstimmung gefunden werden, die ein Benutzer ausgewählt hat. Hervorgehobene Übereinstimmungen gelten für [Kontext-Add-Ins](https://docs.microsoft.com/outlook/add-ins/contextual-outlook-add-ins).</span><span class="sxs-lookup"><span data-stu-id="d9554-p166">Gets the entities found in a highlighted match a user has selected. Highlighted matches apply to [contextual add-ins](https://docs.microsoft.com/outlook/add-ins/contextual-outlook-add-ins).</span></span>

> [!NOTE]
> <span data-ttu-id="d9554-1104">Diese Methode wird in Outlook für iOS oder Outlook für Android nicht unterstützt.</span><span class="sxs-lookup"><span data-stu-id="d9554-1104">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

##### <a name="requirements"></a><span data-ttu-id="d9554-1105">Anforderungen</span><span class="sxs-lookup"><span data-stu-id="d9554-1105">Requirements</span></span>

|<span data-ttu-id="d9554-1106">Anforderung</span><span class="sxs-lookup"><span data-stu-id="d9554-1106">Requirement</span></span>|<span data-ttu-id="d9554-1107">Wert</span><span class="sxs-lookup"><span data-stu-id="d9554-1107">Value</span></span>|
|---|---|
|[<span data-ttu-id="d9554-1108">Mindestversion des Postfachanforderungssatzes</span><span class="sxs-lookup"><span data-stu-id="d9554-1108">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="d9554-1109">1.6</span><span class="sxs-lookup"><span data-stu-id="d9554-1109">1.6</span></span>|
|[<span data-ttu-id="d9554-1110">Mindestberechtigungsstufe</span><span class="sxs-lookup"><span data-stu-id="d9554-1110">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="d9554-1111">ReadItem</span><span class="sxs-lookup"><span data-stu-id="d9554-1111">ReadItem</span></span>|
|[<span data-ttu-id="d9554-1112">Zutreffender Outlook-Modus</span><span class="sxs-lookup"><span data-stu-id="d9554-1112">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="d9554-1113">Lesen</span><span class="sxs-lookup"><span data-stu-id="d9554-1113">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="d9554-1114">Gibt zurück:</span><span class="sxs-lookup"><span data-stu-id="d9554-1114">Returns:</span></span>

<span data-ttu-id="d9554-1115">Typ: [Entitäten](/javascript/api/outlook_1_7/office.entities)</span><span class="sxs-lookup"><span data-stu-id="d9554-1115">Type: [Entities](/javascript/api/outlook_1_7/office.entities)</span></span>

##### <a name="example"></a><span data-ttu-id="d9554-1116">Beispiel</span><span class="sxs-lookup"><span data-stu-id="d9554-1116">Example</span></span>

<span data-ttu-id="d9554-1117">Im folgenden Beispiel wird auf die Adressentitäten in der hervorgehobenen Übereinstimmung zugegriffen, die der Benutzer ausgewählt hat.</span><span class="sxs-lookup"><span data-stu-id="d9554-1117">The following example accesses the addresses entities in the highlighted match selected by the user.</span></span>

```
var contacts = Office.context.mailbox.item.getSelectedEntities().addresses;
```

#### <a name="getselectedregexmatches--object"></a><span data-ttu-id="d9554-1118">getSelectedRegExMatches() → {Object}</span><span class="sxs-lookup"><span data-stu-id="d9554-1118">getSelectedRegExMatches() → {Object}</span></span>

<span data-ttu-id="d9554-p167">Gibt Zeichenfolgenwerte in einer hervorgehobenen Übereinstimmung zurück, die den in der XML-Manifestdatei definierten regulären Ausdrücken entsprechen. Hervorgehobene Übereinstimmungen gelten für [Kontext-Add-Ins](https://docs.microsoft.com/outlook/add-ins/contextual-outlook-add-ins).</span><span class="sxs-lookup"><span data-stu-id="d9554-p167">Returns string values in a highlighted match that match the regular expressions defined in the manifest XML file. Highlighted matches apply to [contextual add-ins](https://docs.microsoft.com/outlook/add-ins/contextual-outlook-add-ins).</span></span>

> [!NOTE]
> <span data-ttu-id="d9554-1121">Diese Methode wird in Outlook für iOS oder Outlook für Android nicht unterstützt.</span><span class="sxs-lookup"><span data-stu-id="d9554-1121">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="d9554-p168">Die `getSelectedRegExMatches`-Methode gibt die Zeichenfolgen zurück, die dem im `ItemHasRegularExpressionMatch`- oder `ItemHasKnownEntity`-Regelelement der XML-Manifestdatei definierten regulären Ausdruck entsprechen. Bei einer `ItemHasRegularExpressionMatch`-Regel muss eine entsprechende Zeichenfolge in der Eigenschaft des Elements vorhanden sein, das von dieser Regel angegeben wird. Der einfache `PropertyName`-Typ definiert die unterstützten Eigenschaften.</span><span class="sxs-lookup"><span data-stu-id="d9554-p168">The `getSelectedRegExMatches` method returns the strings that match the regular expression defined in each `ItemHasRegularExpressionMatch` or `ItemHasKnownEntity` rule element in the manifest XML file. For an `ItemHasRegularExpressionMatch` rule, a matching string has to occur in the property of the item that is specified by that rule. The `PropertyName` simple type defines the supported properties.</span></span>

<span data-ttu-id="d9554-1125">Nehmen Sie z. B. an, dass ein Add-In-Manifest das folgende `Rule`-Element aufweist:</span><span class="sxs-lookup"><span data-stu-id="d9554-1125">For example, consider an add-in manifest has the following `Rule` element:</span></span>

```
<Rule xsi:type="RuleCollection" Mode="And">
  <Rule xsi:type="ItemIs" FormType="Read" ItemType="Message" />
  <Rule xsi:type="RuleCollection" Mode="Or">
    <Rule xsi:type="ItemHasRegularExpressionMatch" RegExName="fruits" RegExValue="apple|banana|coconut" PropertyName="BodyAsPlaintext" IgnoreCase="true" />
    <Rule xsi:type="ItemHasRegularExpressionMatch" RegExName="veggies" RegExValue="tomato|onion|spinach|broccoli" PropertyName="BodyAsPlaintext" IgnoreCase="true" />
  </Rule>
</Rule>
```

<span data-ttu-id="d9554-1126">Das von `getRegExMatches` zurückgegebene Objekt hätte zwei Eigenschaften: `fruits` und `veggies`.</span><span class="sxs-lookup"><span data-stu-id="d9554-1126">The object returned from `getRegExMatches` would have two properties: `fruits` and `veggies`.</span></span>

```
{
  'fruits': ['apple','banana','Banana','coconut'],
  'veggies': ['tomato','onion','spinach','broccoli']
}
```

<span data-ttu-id="d9554-p169">Wenn Sie eine `ItemHasRegularExpressionMatch`-Regel für die Textkörpereigenschaft eines Elements festlegen, sollte der reguläre Ausdruck den Textkörper weiter filtern und nicht versuchen, den gesamten Textkörper des Elements zurückzugeben. Wenn der gesamte Textkörper eines Elements mit einem regulären Ausdruck wie `.*` abgerufen wird, werden nicht immer die gewünschten Ergebnisse erzielt. Verwenden Sie stattdessen die [`Body.getAsync`](/javascript/api/outlook_1_7/office.body#getasync-coerciontype--options--callback-)-Methode, um den gesamten Textkörper abzurufen.</span><span class="sxs-lookup"><span data-stu-id="d9554-p169">If you specify an `ItemHasRegularExpressionMatch` rule on the body property of an item, the regular expression should further filter the body and should not attempt to return the entire body of the item. Using a regular expression such as `.*` to obtain the entire body of an item does not always return the expected results. Instead, use the [`Body.getAsync`](/javascript/api/outlook_1_7/office.body#getasync-coerciontype--options--callback-) method to retrieve the entire body.</span></span>

##### <a name="requirements"></a><span data-ttu-id="d9554-1130">Anforderungen</span><span class="sxs-lookup"><span data-stu-id="d9554-1130">Requirements</span></span>

|<span data-ttu-id="d9554-1131">Anforderung</span><span class="sxs-lookup"><span data-stu-id="d9554-1131">Requirement</span></span>|<span data-ttu-id="d9554-1132">Wert</span><span class="sxs-lookup"><span data-stu-id="d9554-1132">Value</span></span>|
|---|---|
|[<span data-ttu-id="d9554-1133">Mindestversion des Postfachanforderungssatzes</span><span class="sxs-lookup"><span data-stu-id="d9554-1133">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="d9554-1134">1.6</span><span class="sxs-lookup"><span data-stu-id="d9554-1134">1.6</span></span>|
|[<span data-ttu-id="d9554-1135">Mindestberechtigungsstufe</span><span class="sxs-lookup"><span data-stu-id="d9554-1135">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="d9554-1136">ReadItem</span><span class="sxs-lookup"><span data-stu-id="d9554-1136">ReadItem</span></span>|
|[<span data-ttu-id="d9554-1137">Zutreffender Outlook-Modus</span><span class="sxs-lookup"><span data-stu-id="d9554-1137">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="d9554-1138">Lesen</span><span class="sxs-lookup"><span data-stu-id="d9554-1138">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="d9554-1139">Gibt zurück:</span><span class="sxs-lookup"><span data-stu-id="d9554-1139">Returns:</span></span>

<span data-ttu-id="d9554-p170">Ein Objekt mit Arrays aus Zeichenfolgen, die den in der XML-Manifestdatei definierten regulären Ausdrücken entsprechen. Der Name der einzelnen Arrays ist gleich dem entsprechenden Wert des `RegExName`-Attributs der entsprechenden `ItemHasRegularExpressionMatch`-Regel oder des `FilterName`-Attributs der entsprechenden `ItemHasKnownEntity`-Regel.</span><span class="sxs-lookup"><span data-stu-id="d9554-p170">An object that contains arrays of strings that match the regular expressions defined in the manifest XML file. The name of each array is equal to the corresponding value of the `RegExName` attribute of the matching `ItemHasRegularExpressionMatch` rule or the `FilterName` attribute of the matching `ItemHasKnownEntity` rule.</span></span>

##### <a name="example"></a><span data-ttu-id="d9554-1142">Beispiel</span><span class="sxs-lookup"><span data-stu-id="d9554-1142">Example</span></span>

<span data-ttu-id="d9554-1143">Das folgende Beispiel zeigt, wie Sie auf das Array von Übereinstimmungen für die regulären Ausdrucksregelelemente `fruits` und `veggies` zugreifen, die im Manifest angegeben sind.</span><span class="sxs-lookup"><span data-stu-id="d9554-1143">The following example shows how to access the array of matches for the regular expression rule elements `fruits` and `veggies`, which are specified in the manifest.</span></span>

```
var selectedMatches = Office.context.mailbox.item.getSelectedRegExMatches();
var fruits = selectedMatches.fruits;
var veggies = selectedMatches.veggies;
```

####  <a name="loadcustompropertiesasynccallback-usercontext"></a><span data-ttu-id="d9554-1144">loadCustomPropertiesAsync(callback, [userContext])</span><span class="sxs-lookup"><span data-stu-id="d9554-1144">loadCustomPropertiesAsync(callback, [userContext])</span></span>

<span data-ttu-id="d9554-1145">Lädt asynchron benutzerdefinierte Eigenschaften für dieses Add-In für das ausgewählte Element.</span><span class="sxs-lookup"><span data-stu-id="d9554-1145">Asynchronously loads custom properties for this add-in on the selected item.</span></span>

<span data-ttu-id="d9554-p171">Benutzerdefinierte Eigenschaften werden als Schlüssel-/Wert-Paare pro App und pro Element gespeichert. Diese Methode gibt ein `CustomProperties`-Objekt im Rückruf zurück, das Methoden für den Zugriff auf die benutzerdefinierten Eigenschaften für das aktuelle Element und das aktuelle Add-In bietet. Benutzerdefinierte Eigenschaften sind für das Element nicht verschlüsselt und sollten darum nicht als sicherer Speicher verwendet werden.</span><span class="sxs-lookup"><span data-stu-id="d9554-p171">Custom properties are stored as key/value pairs on a per-app, per-item basis. This method returns a `CustomProperties` object in the callback, which provides methods to access the custom properties specific to the current item and the current add-in. Custom properties are not encrypted on the item, so this should not be used as secure storage.</span></span>

##### <a name="parameters"></a><span data-ttu-id="d9554-1149">Parameter:</span><span class="sxs-lookup"><span data-stu-id="d9554-1149">Parameters:</span></span>

|<span data-ttu-id="d9554-1150">Name</span><span class="sxs-lookup"><span data-stu-id="d9554-1150">Name</span></span>|<span data-ttu-id="d9554-1151">Typ</span><span class="sxs-lookup"><span data-stu-id="d9554-1151">Type</span></span>|<span data-ttu-id="d9554-1152">Attribute</span><span class="sxs-lookup"><span data-stu-id="d9554-1152">Attributes</span></span>|<span data-ttu-id="d9554-1153">Beschreibung</span><span class="sxs-lookup"><span data-stu-id="d9554-1153">Description</span></span>|
|---|---|---|---|
|`callback`|<span data-ttu-id="d9554-1154">Funktion</span><span class="sxs-lookup"><span data-stu-id="d9554-1154">function</span></span>||<span data-ttu-id="d9554-1155">Wenn die Methode abgeschlossen wird, wird die Funktion , die im `callback`-Parameter übergeben wird, mit einem einzelnen Parameter (`asyncResult`) aufgerufen, der ein [`AsyncResult`](/javascript/api/office/office.asyncresult)-Objekt darstellt.</span><span class="sxs-lookup"><span data-stu-id="d9554-1155">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span><br/><br/><span data-ttu-id="d9554-1156">Die benutzerdefinierten Eigenschaften werden als [`CustomProperties`](/javascript/api/outlook_1_7/office.customproperties)-Objekt in der `asyncResult.value`-Eigenschaft bereitgestellt.</span><span class="sxs-lookup"><span data-stu-id="d9554-1156">The custom properties are provided as a [`CustomProperties`](/javascript/api/outlook_1_7/office.customproperties) object in the `asyncResult.value` property.</span></span> <span data-ttu-id="d9554-1157">Dieses Objekt kann zum Abrufen, festlegen und Entfernen benutzerdefinierter Eigenschaften aus dem Element und speichern Sie Änderungen an den benutzerdefinierten Eigenschaftensatz zurück an den Server verwendet werden.</span><span class="sxs-lookup"><span data-stu-id="d9554-1157">This object can be used to get, set, and remove custom properties from the item and save changes to the custom property set back to the server.</span></span>|
|`userContext`|<span data-ttu-id="d9554-1158">Objekt</span><span class="sxs-lookup"><span data-stu-id="d9554-1158">Object</span></span>|<span data-ttu-id="d9554-1159">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="d9554-1159">&lt;optional&gt;</span></span>|<span data-ttu-id="d9554-1160">Entwickler können ein beliebiges Objekt bereitstellen, den sie in der Rückruffunktion zugreifen möchten.</span><span class="sxs-lookup"><span data-stu-id="d9554-1160">Developers can provide any object they wish to access in the callback function.</span></span> <span data-ttu-id="d9554-1161">Dieses Objekt zugegriffen werden kann, indem die `asyncResult.asyncContext` -Eigenschaft in der Callback-Funktion.</span><span class="sxs-lookup"><span data-stu-id="d9554-1161">This object can be accessed by the `asyncResult.asyncContext` property in the callback function.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="d9554-1162">Anforderungen</span><span class="sxs-lookup"><span data-stu-id="d9554-1162">Requirements</span></span>

|<span data-ttu-id="d9554-1163">Anforderung</span><span class="sxs-lookup"><span data-stu-id="d9554-1163">Requirement</span></span>|<span data-ttu-id="d9554-1164">Wert</span><span class="sxs-lookup"><span data-stu-id="d9554-1164">Value</span></span>|
|---|---|
|[<span data-ttu-id="d9554-1165">Mindestversion des Postfachanforderungssatzes</span><span class="sxs-lookup"><span data-stu-id="d9554-1165">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="d9554-1166">1.0</span><span class="sxs-lookup"><span data-stu-id="d9554-1166">1.0</span></span>|
|[<span data-ttu-id="d9554-1167">Mindestberechtigungsstufe</span><span class="sxs-lookup"><span data-stu-id="d9554-1167">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="d9554-1168">ReadItem</span><span class="sxs-lookup"><span data-stu-id="d9554-1168">ReadItem</span></span>|
|[<span data-ttu-id="d9554-1169">Zutreffender Outlook-Modus</span><span class="sxs-lookup"><span data-stu-id="d9554-1169">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="d9554-1170">Verfassen oder Lesen</span><span class="sxs-lookup"><span data-stu-id="d9554-1170">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="d9554-1171">Beispiel</span><span class="sxs-lookup"><span data-stu-id="d9554-1171">Example</span></span>

<span data-ttu-id="d9554-p174">Das folgende Codebeispiel veranschaulicht die Verwendung der `loadCustomPropertiesAsync`-Methode zum asynchronen Laden der benutzerdefinierten Eigenschaften, die für das aktuelle Element spezifisch sind. Des Weiteren veranschaulicht das Beispiel die Verwendung der `CustomProperties.saveAsync`-Methode zum Speichern der Eigenschaften auf dem Server. Nach dem Laden der benutzerdefinierten Eigenschaften wird in dem Codebeispiel die `CustomProperties.get`-Methode dazu verwendet, die benutzerdefinierte `myProp`-Eigenschaft zu lesen. Die `CustomProperties.set`-Methode wird dazu verwendet, die benutzerdefinierte `otherProp`-Eigenschaft zu schreiben. Schließlich wird die `saveAsync`-Methode aufgerufen, um die benutzerdefinierten Eigenschaften zu speichern.</span><span class="sxs-lookup"><span data-stu-id="d9554-p174">The following code example shows how to use the `loadCustomPropertiesAsync` method to asynchronously load custom properties that are specific to the current item. The example also shows how to use the `CustomProperties.saveAsync` method to save these properties back to the server. After loading the custom properties, the code sample uses the `CustomProperties.get` method to read the custom property `myProp`, the `CustomProperties.set` method to write the custom property `otherProp`, and then finally calls the `saveAsync` method to save the custom properties.</span></span>

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

####  <a name="removeattachmentasyncattachmentid-options-callback"></a><span data-ttu-id="d9554-1175">removeAttachmentAsync(attachmentId, [options], [callback])</span><span class="sxs-lookup"><span data-stu-id="d9554-1175">removeAttachmentAsync(attachmentId, [options], [callback])</span></span>

<span data-ttu-id="d9554-1176">Entfernt eine Anlage aus einer Nachricht oder einem Termin.</span><span class="sxs-lookup"><span data-stu-id="d9554-1176">Removes an attachment from a message or appointment.</span></span>

<span data-ttu-id="d9554-p175">Die `removeAttachmentAsync`-Methode entfernt die Anlage mit dem angegebenen Bezeichner aus dem Element. Als bewährte Vorgehensweise sollten Sie den Anlagenbezeichner nur dann zum Entfernen einer Anlage verwenden, wenn die gleiche Mail-App die Anlage in der gleichen Sitzung hinzugefügt hat. In Outlook Web App und OWA für Geräte ist der Anlagenbezeichner nur innerhalb der gleichen Sitzung gültig. Eine Sitzung ist abgeschlossen, wenn der Benutzer die App schließt, oder wenn der Benutzer in einem eingebetteten Formular mit dem Verfassen beginnt und den Vorgang dann in einem separaten Fenster fortsetzt.</span><span class="sxs-lookup"><span data-stu-id="d9554-p175">The `removeAttachmentAsync` method removes the attachment with the specified identifier from the item. As a best practice, you should use the attachment identifier to remove an attachment only if the same mail app has added that attachment in the same session. In Outlook Web App and OWA for Devices, the attachment identifier is valid only within the same session. A session is over when the user closes the app, or if the user starts composing in an inline form and subsequently pops out the inline form to continue in a separate window.</span></span>

##### <a name="parameters"></a><span data-ttu-id="d9554-1181">Parameter:</span><span class="sxs-lookup"><span data-stu-id="d9554-1181">Parameters:</span></span>

|<span data-ttu-id="d9554-1182">Name</span><span class="sxs-lookup"><span data-stu-id="d9554-1182">Name</span></span>|<span data-ttu-id="d9554-1183">Typ</span><span class="sxs-lookup"><span data-stu-id="d9554-1183">Type</span></span>|<span data-ttu-id="d9554-1184">Attribute</span><span class="sxs-lookup"><span data-stu-id="d9554-1184">Attributes</span></span>|<span data-ttu-id="d9554-1185">Beschreibung</span><span class="sxs-lookup"><span data-stu-id="d9554-1185">Description</span></span>|
|---|---|---|---|
|`attachmentId`|<span data-ttu-id="d9554-1186">Zeichenfolge</span><span class="sxs-lookup"><span data-stu-id="d9554-1186">String</span></span>||<span data-ttu-id="d9554-p176">Der Bezeichner des Anhangs, der entfernt werden soll. Die maximale Länge der Zeichenfolge ist 100 Zeichen.</span><span class="sxs-lookup"><span data-stu-id="d9554-p176">The identifier of the attachment to remove. The maximum length of the string is 100 characters.</span></span>|
|`options`|<span data-ttu-id="d9554-1189">Object</span><span class="sxs-lookup"><span data-stu-id="d9554-1189">Object</span></span>|<span data-ttu-id="d9554-1190">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="d9554-1190">&lt;optional&gt;</span></span>|<span data-ttu-id="d9554-1191">Ein Objektliteral, das mindestens eine der folgenden Eigenschaften enthält.</span><span class="sxs-lookup"><span data-stu-id="d9554-1191">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`|<span data-ttu-id="d9554-1192">Object</span><span class="sxs-lookup"><span data-stu-id="d9554-1192">Object</span></span>|<span data-ttu-id="d9554-1193">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="d9554-1193">&lt;optional&gt;</span></span>|<span data-ttu-id="d9554-1194">Entwickler können ein Objekt bereitstellen, auf das sie in der Callbackmethode zugreifen möchten.</span><span class="sxs-lookup"><span data-stu-id="d9554-1194">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`|<span data-ttu-id="d9554-1195">Funktion</span><span class="sxs-lookup"><span data-stu-id="d9554-1195">function</span></span>|<span data-ttu-id="d9554-1196">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="d9554-1196">&lt;optional&gt;</span></span>|<span data-ttu-id="d9554-1197">Wenn die Methode abgeschlossen wird, wird die Funktion , die im `callback`-Parameter übergeben wird, mit einem einzelnen Parameter (`asyncResult`) aufgerufen, der ein [`AsyncResult`](/javascript/api/office/office.asyncresult)-Objekt darstellt.</span><span class="sxs-lookup"><span data-stu-id="d9554-1197">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span> <br/><span data-ttu-id="d9554-1198">Wenn beim Entfernen der Anlage ein Fehler auftritt, enthält die Eigenschaft `asyncResult.error` einen Fehlercode mit dem Grund für den Fehler.</span><span class="sxs-lookup"><span data-stu-id="d9554-1198">If removing the attachment fails, the `asyncResult.error` property will contain an error code with the reason for the failure.</span></span>|

##### <a name="errors"></a><span data-ttu-id="d9554-1199">Fehler</span><span class="sxs-lookup"><span data-stu-id="d9554-1199">Errors</span></span>

|<span data-ttu-id="d9554-1200">Fehlercode</span><span class="sxs-lookup"><span data-stu-id="d9554-1200">Error code</span></span>|<span data-ttu-id="d9554-1201">Beschreibung</span><span class="sxs-lookup"><span data-stu-id="d9554-1201">Description</span></span>|
|------------|-------------|
|`InvalidAttachmentId`|<span data-ttu-id="d9554-1202">Der Bezeichner für die Anlage ist nicht vorhanden.</span><span class="sxs-lookup"><span data-stu-id="d9554-1202">The attachment identifier does not exist.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="d9554-1203">Anforderungen</span><span class="sxs-lookup"><span data-stu-id="d9554-1203">Requirements</span></span>

|<span data-ttu-id="d9554-1204">Anforderung</span><span class="sxs-lookup"><span data-stu-id="d9554-1204">Requirement</span></span>|<span data-ttu-id="d9554-1205">Wert</span><span class="sxs-lookup"><span data-stu-id="d9554-1205">Value</span></span>|
|---|---|
|[<span data-ttu-id="d9554-1206">Mindestversion des Postfachanforderungssatzes</span><span class="sxs-lookup"><span data-stu-id="d9554-1206">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="d9554-1207">1.1</span><span class="sxs-lookup"><span data-stu-id="d9554-1207">1.1</span></span>|
|[<span data-ttu-id="d9554-1208">Mindestberechtigungsstufe</span><span class="sxs-lookup"><span data-stu-id="d9554-1208">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="d9554-1209">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="d9554-1209">ReadWriteItem</span></span>|
|[<span data-ttu-id="d9554-1210">Zutreffender Outlook-Modus</span><span class="sxs-lookup"><span data-stu-id="d9554-1210">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="d9554-1211">Verfassen</span><span class="sxs-lookup"><span data-stu-id="d9554-1211">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="d9554-1212">Beispiel</span><span class="sxs-lookup"><span data-stu-id="d9554-1212">Example</span></span>

<span data-ttu-id="d9554-1213">Mit dem folgende Code wird eine Anlage mit dem Bezeichner "0" entfernt.</span><span class="sxs-lookup"><span data-stu-id="d9554-1213">The following code removes an attachment with an identifier of '0'.</span></span>

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

####  <a name="removehandlerasynceventtype-handler-options-callback"></a><span data-ttu-id="d9554-1214">RemoveHandlerAsync (EventType, Handler, [Optionen], [Rückruf])</span><span class="sxs-lookup"><span data-stu-id="d9554-1214">removeHandlerAsync(eventType, handler, [options], [callback])</span></span>

<span data-ttu-id="d9554-1215">Entfernt einen Ereignishandler für ein Ereignis unterstützt.</span><span class="sxs-lookup"><span data-stu-id="d9554-1215">Removes an event handler for a supported event.</span></span>

<span data-ttu-id="d9554-1216">Derzeit sind die unterstützten Ereignistypen `Office.EventType.AppointmentTimeChanged`, `Office.EventType.RecipientsChanged`, und`Office.EventType.RecurrenceChanged`</span><span class="sxs-lookup"><span data-stu-id="d9554-1216">Currently the supported event types are `Office.EventType.AppointmentTimeChanged`, `Office.EventType.RecipientsChanged`, and `Office.EventType.RecurrenceChanged`</span></span>

##### <a name="parameters"></a><span data-ttu-id="d9554-1217">Parameter:</span><span class="sxs-lookup"><span data-stu-id="d9554-1217">Parameters:</span></span>

| <span data-ttu-id="d9554-1218">Name</span><span class="sxs-lookup"><span data-stu-id="d9554-1218">Name</span></span> | <span data-ttu-id="d9554-1219">Typ</span><span class="sxs-lookup"><span data-stu-id="d9554-1219">Type</span></span> | <span data-ttu-id="d9554-1220">Attribute</span><span class="sxs-lookup"><span data-stu-id="d9554-1220">Attributes</span></span> | <span data-ttu-id="d9554-1221">Beschreibung</span><span class="sxs-lookup"><span data-stu-id="d9554-1221">Description</span></span> |
|---|---|---|---|
| `eventType` | [<span data-ttu-id="d9554-1222">Office.EventType</span><span class="sxs-lookup"><span data-stu-id="d9554-1222">Office.EventType</span></span>](office.md#eventtype-string) || <span data-ttu-id="d9554-1223">Das Ereignis, das den Handler aufrufen soll</span><span class="sxs-lookup"><span data-stu-id="d9554-1223">The event that should invoke the handler.</span></span> |
| `handler` | <span data-ttu-id="d9554-1224">Function</span><span class="sxs-lookup"><span data-stu-id="d9554-1224">Function</span></span> || <span data-ttu-id="d9554-p177">Die Funktion, die das Ereignis behandeln soll. Die Funktion muss einen einzigen Parameter akzeptieren (ein Objektliteral). Die Eigenschaft `type` dieses Parameters entspricht dem Parameter `eventType`, der an `removeHandlerAsync` übergeben wird.</span><span class="sxs-lookup"><span data-stu-id="d9554-p177">The function to handle the event. The function must accept a single parameter, which is an object literal. The `type` property on the parameter will match the `eventType` parameter passed to `removeHandlerAsync`.</span></span> |
| `options` | <span data-ttu-id="d9554-1228">Objekt</span><span class="sxs-lookup"><span data-stu-id="d9554-1228">Object</span></span> | <span data-ttu-id="d9554-1229">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="d9554-1229">&lt;optional&gt;</span></span> | <span data-ttu-id="d9554-1230">Ein Objektliteral, das mindestens eine der folgenden Eigenschaften enthält.</span><span class="sxs-lookup"><span data-stu-id="d9554-1230">An object literal that contains one or more of the following properties.</span></span> |
| `options.asyncContext` | <span data-ttu-id="d9554-1231">Object</span><span class="sxs-lookup"><span data-stu-id="d9554-1231">Object</span></span> | <span data-ttu-id="d9554-1232">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="d9554-1232">&lt;optional&gt;</span></span> | <span data-ttu-id="d9554-1233">Entwickler können ein Objekt bereitstellen, auf das sie in der Callbackmethode zugreifen möchten.</span><span class="sxs-lookup"><span data-stu-id="d9554-1233">Developers can provide any object they wish to access in the callback method.</span></span> |
| `callback` | <span data-ttu-id="d9554-1234">Funktion</span><span class="sxs-lookup"><span data-stu-id="d9554-1234">function</span></span>| <span data-ttu-id="d9554-1235">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="d9554-1235">&lt;optional&gt;</span></span>|<span data-ttu-id="d9554-1236">Wenn die Methode abgeschlossen wird, wird die Funktion , die im `callback`-Parameter übergeben wird, mit einem einzelnen Parameter (`asyncResult`) aufgerufen, der ein [`AsyncResult`](/javascript/api/office/office.asyncresult)-Objekt darstellt.</span><span class="sxs-lookup"><span data-stu-id="d9554-1236">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="d9554-1237">Anforderungen</span><span class="sxs-lookup"><span data-stu-id="d9554-1237">Requirements</span></span>

|<span data-ttu-id="d9554-1238">Anforderung</span><span class="sxs-lookup"><span data-stu-id="d9554-1238">Requirement</span></span>| <span data-ttu-id="d9554-1239">Wert</span><span class="sxs-lookup"><span data-stu-id="d9554-1239">Value</span></span>|
|---|---|
|[<span data-ttu-id="d9554-1240">Mindestversion des Postfachanforderungssatzes</span><span class="sxs-lookup"><span data-stu-id="d9554-1240">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="d9554-1241">1.7</span><span class="sxs-lookup"><span data-stu-id="d9554-1241">1.7</span></span> |
|[<span data-ttu-id="d9554-1242">Mindestberechtigungsstufe</span><span class="sxs-lookup"><span data-stu-id="d9554-1242">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="d9554-1243">ReadItem</span><span class="sxs-lookup"><span data-stu-id="d9554-1243">ReadItem</span></span> |
|[<span data-ttu-id="d9554-1244">Zutreffender Outlook-Modus</span><span class="sxs-lookup"><span data-stu-id="d9554-1244">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="d9554-1245">Verfassen oder Lesen</span><span class="sxs-lookup"><span data-stu-id="d9554-1245">Compose or read</span></span> |

##### <a name="example"></a><span data-ttu-id="d9554-1246">Beispiel</span><span class="sxs-lookup"><span data-stu-id="d9554-1246">Example</span></span>

```
Office.initialize = function (reason) {
  $(document).ready(function () {
    Office.context.mailbox.item.removeHandlerAsync(Office.EventType.RecurrenceChanged, loadNewItem, function (result) {
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

####  <a name="saveasyncoptions-callback"></a><span data-ttu-id="d9554-1247">saveAsync([options], callback)</span><span class="sxs-lookup"><span data-stu-id="d9554-1247">saveAsync([options], callback)</span></span>

<span data-ttu-id="d9554-1248">Speicher asynchron ein Element. .</span><span class="sxs-lookup"><span data-stu-id="d9554-1248">Asynchronously saves an item.</span></span>

<span data-ttu-id="d9554-p178">Beim Aufrufen speichert diese Methode die aktuelle Nachricht als Entwurf und  gibt die Element-ID über die Callbackmethode zurück. In Outlook Web App oder Outlook im Onlinemodus wird das Element auf dem Server gespeichert. In Outlook im Cache-Modus wird das Element im lokalen Cache gespeichert.</span><span class="sxs-lookup"><span data-stu-id="d9554-p178">When invoked, this method saves the current message as a draft and returns the item id via the callback method. In Outlook Web App or Outlook in online mode, the item is saved to the server. In Outlook in cached mode, the item is saved to the local cache.</span></span>

> [!NOTE]
> <span data-ttu-id="d9554-1252">Wenn Ihr Add-In ruft `saveAsync` für ein Element im Entwurfsmodus, um das Abrufen einer `itemId` um mit EWS oder REST-API verwenden, beachten Sie, dass wenn Outlook im Cache-Modus ist, es kann einige Zeit dauern, bevor das Element tatsächlich auf dem Server synchronisiert wird.</span><span class="sxs-lookup"><span data-stu-id="d9554-1252">If your add-in calls `saveAsync` on an item in compose mode in order to get an `itemId` to use with EWS or the REST API, be aware that when Outlook is in cached mode, it may take some time before the item is actually synced to the server.</span></span> <span data-ttu-id="d9554-1253">Bis das Element mit synchronisiert ist, die `itemId` wird ein Fehler zurückgegeben.</span><span class="sxs-lookup"><span data-stu-id="d9554-1253">Until the item is synced, using the `itemId` will return an error.</span></span>

<span data-ttu-id="d9554-p180">Da Termine keinen Entwurfsstatus haben, wird das Element bei Aufruf von `saveAsync` für einen Termin im Verfassenmodus als normaler Termin im Kalender des Benutzers gespeichert. Für neue Termine, die noch nicht gespeichert wurden, wird keine Einladung gesendet. Beim Speichern eines vorhandenen Termins wird eine Aktualisierung an die hinzugefügten oder entfernten Teilnehmer gesendet.</span><span class="sxs-lookup"><span data-stu-id="d9554-p180">Since appointments have no draft state, if `saveAsync` is called on an appointment in compose mode, the item will be saved as a normal appointment on the user's calendar. For new appointments that have not been saved before, no invitation will be sent. Saving an existing appointment will send an update to added or removed attendees.</span></span>

> [!NOTE]
> <span data-ttu-id="d9554-1257">Die folgenden Clients haben unterschiedlichem Verhalten für `saveAsync` Termine im Verfassenmodus:</span><span class="sxs-lookup"><span data-stu-id="d9554-1257">The following clients have different behavior for `saveAsync` on appointments in compose mode:</span></span>
>
> - <span data-ttu-id="d9554-1258">Outlook für Mac unterstützt keine `saveAsync` auf einer Besprechung im Verfassenmodus.</span><span class="sxs-lookup"><span data-stu-id="d9554-1258">Mac Outlook does not support `saveAsync` on a meeting in compose mode.</span></span> <span data-ttu-id="d9554-1259">Aufrufen von `saveAsync` auf einer Besprechung in Outlook für Mac wird ein Fehler zurückgegeben.</span><span class="sxs-lookup"><span data-stu-id="d9554-1259">Calling `saveAsync` on a meeting in Mac Outlook will return an error.</span></span>
> - <span data-ttu-id="d9554-1260">Outlook im Web immer sendet eine Einladung oder zu aktualisieren, wenn `saveAsync` für einen Termin aufgerufen wird, im Verfassenmodus.</span><span class="sxs-lookup"><span data-stu-id="d9554-1260">Outlook on the web always sends an invitation or update when `saveAsync` is called on an appointment in compose mode.</span></span>

##### <a name="parameters"></a><span data-ttu-id="d9554-1261">Parameter:</span><span class="sxs-lookup"><span data-stu-id="d9554-1261">Parameters:</span></span>

|<span data-ttu-id="d9554-1262">Name</span><span class="sxs-lookup"><span data-stu-id="d9554-1262">Name</span></span>|<span data-ttu-id="d9554-1263">Typ</span><span class="sxs-lookup"><span data-stu-id="d9554-1263">Type</span></span>|<span data-ttu-id="d9554-1264">Attribute</span><span class="sxs-lookup"><span data-stu-id="d9554-1264">Attributes</span></span>|<span data-ttu-id="d9554-1265">Beschreibung</span><span class="sxs-lookup"><span data-stu-id="d9554-1265">Description</span></span>|
|---|---|---|---|
|`options`|<span data-ttu-id="d9554-1266">Objekt</span><span class="sxs-lookup"><span data-stu-id="d9554-1266">Object</span></span>|<span data-ttu-id="d9554-1267">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="d9554-1267">&lt;optional&gt;</span></span>|<span data-ttu-id="d9554-1268">Ein Objektliteral, das mindestens eine der folgenden Eigenschaften enthält.</span><span class="sxs-lookup"><span data-stu-id="d9554-1268">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`|<span data-ttu-id="d9554-1269">Object</span><span class="sxs-lookup"><span data-stu-id="d9554-1269">Object</span></span>|<span data-ttu-id="d9554-1270">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="d9554-1270">&lt;optional&gt;</span></span>|<span data-ttu-id="d9554-1271">Entwickler können ein Objekt bereitstellen, auf das sie in der Callbackmethode zugreifen möchten.</span><span class="sxs-lookup"><span data-stu-id="d9554-1271">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`|<span data-ttu-id="d9554-1272">Funktion</span><span class="sxs-lookup"><span data-stu-id="d9554-1272">function</span></span>||<span data-ttu-id="d9554-1273">Wenn die Methode abgeschlossen wird, wird die Funktion , die im `callback`-Parameter übergeben wird, mit einem einzelnen Parameter (`asyncResult`) aufgerufen, der ein [`AsyncResult`](/javascript/api/office/office.asyncresult)-Objekt darstellt.</span><span class="sxs-lookup"><span data-stu-id="d9554-1273">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span><br/><br/><span data-ttu-id="d9554-1274">Auf Erfolg, erfolgt die Element-ID der `asyncResult.value` Eigenschaft.</span><span class="sxs-lookup"><span data-stu-id="d9554-1274">On success, the item identifier is provided in the `asyncResult.value` property.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="d9554-1275">Anforderungen</span><span class="sxs-lookup"><span data-stu-id="d9554-1275">Requirements</span></span>

|<span data-ttu-id="d9554-1276">Anforderung</span><span class="sxs-lookup"><span data-stu-id="d9554-1276">Requirement</span></span>|<span data-ttu-id="d9554-1277">Wert</span><span class="sxs-lookup"><span data-stu-id="d9554-1277">Value</span></span>|
|---|---|
|[<span data-ttu-id="d9554-1278">Mindestversion des Postfachanforderungssatzes</span><span class="sxs-lookup"><span data-stu-id="d9554-1278">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="d9554-1279">1.3</span><span class="sxs-lookup"><span data-stu-id="d9554-1279">1.3</span></span>|
|[<span data-ttu-id="d9554-1280">Mindestberechtigungsstufe</span><span class="sxs-lookup"><span data-stu-id="d9554-1280">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="d9554-1281">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="d9554-1281">ReadWriteItem</span></span>|
|[<span data-ttu-id="d9554-1282">Zutreffender Outlook-Modus</span><span class="sxs-lookup"><span data-stu-id="d9554-1282">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="d9554-1283">Verfassen</span><span class="sxs-lookup"><span data-stu-id="d9554-1283">Compose</span></span>|

##### <a name="examples"></a><span data-ttu-id="d9554-1284">Beispiele</span><span class="sxs-lookup"><span data-stu-id="d9554-1284">Examples</span></span>

```
Office.context.mailbox.item.saveAsync(
  function callback(result) {
    // Process the result
  });
```

<span data-ttu-id="d9554-p182">Es folgt ein Beispiel für den `result`-Parameter, der an die Callbackfunktion übergeben wird. Die `value`-Eigenschaft enthält die Element-ID des Elements.</span><span class="sxs-lookup"><span data-stu-id="d9554-p182">The following is an example of the `result` parameter passed to the callback function. The `value` property contains the item ID of the item.</span></span>

```
{
  "value":"AAMkADI5...AAA=",
  "status":"succeeded"
}
```

####  <a name="setselecteddataasyncdata-options-callback"></a><span data-ttu-id="d9554-1287">setSelectedDataAsync(data, [options], callback)</span><span class="sxs-lookup"><span data-stu-id="d9554-1287">setSelectedDataAsync(data, [options], callback)</span></span>

<span data-ttu-id="d9554-1288">Fügt asynchron Daten in den Textkörper oder Betreff einer Nachricht ein.</span><span class="sxs-lookup"><span data-stu-id="d9554-1288">Asynchronously inserts data into the body or subject of a message.</span></span>

<span data-ttu-id="d9554-p183">Die `setSelectedDataAsync`-Methode fügt die angegebene Zeichenfolge an der Cursorposition im Betreff oder im Textkörper ein, oder, falls im Editor Text ausgewählt ist, ersetzt den markierten Text. Wenn sich der Cursor nicht im Textkörper oder im Betreffsfeld befindet, wird ein Fehler zurückgegeben. Nach dem Einfügen wird der Cursor am Ende der eingefügten Inhalte platziert.</span><span class="sxs-lookup"><span data-stu-id="d9554-p183">The `setSelectedDataAsync` method inserts the specified string at the cursor location in the subject or body of the item, or, if text is selected in the editor, it replaces the selected text. If the cursor is not in the body or subject field, an error is returned. After insertion, the cursor is placed at the end of the inserted content.</span></span>

##### <a name="parameters"></a><span data-ttu-id="d9554-1292">Parameter:</span><span class="sxs-lookup"><span data-stu-id="d9554-1292">Parameters:</span></span>

|<span data-ttu-id="d9554-1293">Name</span><span class="sxs-lookup"><span data-stu-id="d9554-1293">Name</span></span>|<span data-ttu-id="d9554-1294">Typ</span><span class="sxs-lookup"><span data-stu-id="d9554-1294">Type</span></span>|<span data-ttu-id="d9554-1295">Attribute</span><span class="sxs-lookup"><span data-stu-id="d9554-1295">Attributes</span></span>|<span data-ttu-id="d9554-1296">Beschreibung</span><span class="sxs-lookup"><span data-stu-id="d9554-1296">Description</span></span>|
|---|---|---|---|
|`data`|<span data-ttu-id="d9554-1297">Zeichenfolge</span><span class="sxs-lookup"><span data-stu-id="d9554-1297">String</span></span>||<span data-ttu-id="d9554-p184">Die einzufügenden Daten. Daten dürfen 1.000.000 Zeichen nicht überschreiten. Werden mehr als 1.000.000 Zeichen übergeben, wird eine `ArgumentOutOfRange`-Ausnahme ausgelöst.</span><span class="sxs-lookup"><span data-stu-id="d9554-p184">The data to be inserted. Data is not to exceed 1,000,000 characters. If more than 1,000,000 characters are passed in, an `ArgumentOutOfRange` exception is thrown.</span></span>|
|`options`|<span data-ttu-id="d9554-1301">Object</span><span class="sxs-lookup"><span data-stu-id="d9554-1301">Object</span></span>|<span data-ttu-id="d9554-1302">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="d9554-1302">&lt;optional&gt;</span></span>|<span data-ttu-id="d9554-1303">Ein Objektliteral, das mindestens eine der folgenden Eigenschaften enthält.</span><span class="sxs-lookup"><span data-stu-id="d9554-1303">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`|<span data-ttu-id="d9554-1304">Object</span><span class="sxs-lookup"><span data-stu-id="d9554-1304">Object</span></span>|<span data-ttu-id="d9554-1305">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="d9554-1305">&lt;optional&gt;</span></span>|<span data-ttu-id="d9554-1306">Entwickler können ein Objekt bereitstellen, auf das sie in der Rückrufmethode zugreifen möchten.</span><span class="sxs-lookup"><span data-stu-id="d9554-1306">Developers can provide any object they wish to access in the callback method.</span></span>|
|`options.coercionType`|[<span data-ttu-id="d9554-1307">Office.CoercionType</span><span class="sxs-lookup"><span data-stu-id="d9554-1307">Office.CoercionType</span></span>](office.md#coerciontype-string)|<span data-ttu-id="d9554-1308">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="d9554-1308">&lt;optional&gt;</span></span>|<span data-ttu-id="d9554-p185">Wenn `text`, wird das aktuelle Format in Outlook Web App und Outlook angewendet. Wenn das Feld ein HTML-Editor ist, werden nur die Textdaten eingefügt, selbst wenn es sich bei den Daten um HTML-Daten handelt.</span><span class="sxs-lookup"><span data-stu-id="d9554-p185">If `text`, the current style is applied in Outlook Web App and Outlook. If the field is an HTML editor, only the text data is inserted, even if the data is HTML.</span></span><br/><br/><span data-ttu-id="d9554-p186">Wenn `html` gesetzt ist und das Feld HTML unterstützt (der Betreff jedoch nicht), wird die aktuelle Formatvorlage in Outlook Web App angewendet und die Standardformatvorlage in Outlook. Ist das Feld ein Textfeld, wird ein Fehler des Typs `InvalidDataFormat` zurückgegeben.</span><span class="sxs-lookup"><span data-stu-id="d9554-p186">If `html` and the field supports HTML (the subject doesn't), the current style is applied in Outlook Web App and the default style is applied in Outlook. If the field is a text field, an `InvalidDataFormat` error is returned.</span></span><br/><br/><span data-ttu-id="d9554-1313">Wenn `coercionType` nicht festgelegt wird, hängt das Ergebnis vom Feld ab: Wenn das Feld HTML ist, wird HTML verwendet. Wenn das Feld Text ist, wird Nur-Text verwendet.</span><span class="sxs-lookup"><span data-stu-id="d9554-1313">If `coercionType` is not set, the result depends on the field: if the field is HTML then HTML is used; if the field is text, then plain text is used.</span></span>|
|`callback`|<span data-ttu-id="d9554-1314">function</span><span class="sxs-lookup"><span data-stu-id="d9554-1314">function</span></span>||<span data-ttu-id="d9554-1315">Wenn die Methode abgeschlossen wird, wird die Funktion , die im `callback`-Parameter übergeben wird, mit einem einzelnen Parameter (`asyncResult`) aufgerufen, der ein [`AsyncResult`](/javascript/api/office/office.asyncresult)-Objekt darstellt.</span><span class="sxs-lookup"><span data-stu-id="d9554-1315">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="d9554-1316">Anforderungen</span><span class="sxs-lookup"><span data-stu-id="d9554-1316">Requirements</span></span>

|<span data-ttu-id="d9554-1317">Anforderung</span><span class="sxs-lookup"><span data-stu-id="d9554-1317">Requirement</span></span>|<span data-ttu-id="d9554-1318">Wert</span><span class="sxs-lookup"><span data-stu-id="d9554-1318">Value</span></span>|
|---|---|
|[<span data-ttu-id="d9554-1319">Mindestversion des Postfachanforderungssatzes</span><span class="sxs-lookup"><span data-stu-id="d9554-1319">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="d9554-1320">1.2</span><span class="sxs-lookup"><span data-stu-id="d9554-1320">1.2</span></span>|
|[<span data-ttu-id="d9554-1321">Mindestberechtigungsstufe</span><span class="sxs-lookup"><span data-stu-id="d9554-1321">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="d9554-1322">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="d9554-1322">ReadWriteItem</span></span>|
|[<span data-ttu-id="d9554-1323">Zutreffender Outlook-Modus</span><span class="sxs-lookup"><span data-stu-id="d9554-1323">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="d9554-1324">Verfassen</span><span class="sxs-lookup"><span data-stu-id="d9554-1324">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="d9554-1325">Beispiel</span><span class="sxs-lookup"><span data-stu-id="d9554-1325">Example</span></span>

```
Office.context.mailbox.item.setSelectedDataAsync("Hello World!");
Office.context.mailbox.item.setSelectedDataAsync("<b>Hello World!</b>", { coercionType : "html" });
```