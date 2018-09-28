
# <a name="item"></a><span data-ttu-id="a7c7c-101">item</span><span class="sxs-lookup"><span data-stu-id="a7c7c-101">item</span></span>

### <a name="officeofficemdcontextofficecontextmdmailboxofficecontextmailboxmditem"></a><span data-ttu-id="a7c7c-102">[Office](office.md)[.context](office.context.md)[.mailbox](office.context.mailbox.md).item</span><span class="sxs-lookup"><span data-stu-id="a7c7c-102">[Office](office.md)[.context](office.context.md)[.mailbox](office.context.mailbox.md).item</span></span>

<span data-ttu-id="a7c7c-p101">Der `item`-Namespace wird für den Zugriff auf die aktuell ausgewählte Nachricht, Besprechungsanfrage oder den aktuell ausgewählten Termin verwendet. Sie können den Typ von `item` mithilfe der [itemType](#itemtype-officemailboxenumsitemtypejavascriptapioutlookofficemailboxenumsitemtype)-Eigenschaft bestimmen.</span><span class="sxs-lookup"><span data-stu-id="a7c7c-p101">The `item` namespace is used to access the currently selected message, meeting request, or appointment. You can determine the type of the `item` by using the [itemType](#itemtype-officemailboxenumsitemtypejavascriptapioutlookofficemailboxenumsitemtype) property.</span></span>

##### <a name="requirements"></a><span data-ttu-id="a7c7c-105">Anforderungen</span><span class="sxs-lookup"><span data-stu-id="a7c7c-105">Requirements</span></span>

|<span data-ttu-id="a7c7c-106">Anforderung</span><span class="sxs-lookup"><span data-stu-id="a7c7c-106">Requirement</span></span>|<span data-ttu-id="a7c7c-107">Wert</span><span class="sxs-lookup"><span data-stu-id="a7c7c-107">Value</span></span>|
|---|---|
|[<span data-ttu-id="a7c7c-108">Mindestversion des Postfachanforderungssatzes</span><span class="sxs-lookup"><span data-stu-id="a7c7c-108">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="a7c7c-109">1.0</span><span class="sxs-lookup"><span data-stu-id="a7c7c-109">1.0</span></span>|
|[<span data-ttu-id="a7c7c-110">Mindestberechtigungsstufe</span><span class="sxs-lookup"><span data-stu-id="a7c7c-110">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="a7c7c-111">Eingeschränkt</span><span class="sxs-lookup"><span data-stu-id="a7c7c-111">Restricted</span></span>|
|[<span data-ttu-id="a7c7c-112">Zutreffender Outlook-Modus</span><span class="sxs-lookup"><span data-stu-id="a7c7c-112">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="a7c7c-113">Verfassen oder Lesen</span><span class="sxs-lookup"><span data-stu-id="a7c7c-113">Compose or read</span></span>|

##### <a name="members-and-methods"></a><span data-ttu-id="a7c7c-114">Elemente und Methoden</span><span class="sxs-lookup"><span data-stu-id="a7c7c-114">Members and methods</span></span>

| <span data-ttu-id="a7c7c-115">Element</span><span class="sxs-lookup"><span data-stu-id="a7c7c-115">Member</span></span> | <span data-ttu-id="a7c7c-116">Typ</span><span class="sxs-lookup"><span data-stu-id="a7c7c-116">Type</span></span> |
|--------|------|
| [<span data-ttu-id="a7c7c-117">attachments</span><span class="sxs-lookup"><span data-stu-id="a7c7c-117">attachments</span></span>](#attachments-arrayattachmentdetailsjavascriptapioutlookofficeattachmentdetails) | <span data-ttu-id="a7c7c-118">Element</span><span class="sxs-lookup"><span data-stu-id="a7c7c-118">Member</span></span> |
| [<span data-ttu-id="a7c7c-119">bcc</span><span class="sxs-lookup"><span data-stu-id="a7c7c-119">bcc</span></span>](#bcc-recipientsjavascriptapioutlookofficerecipients) | <span data-ttu-id="a7c7c-120">Element</span><span class="sxs-lookup"><span data-stu-id="a7c7c-120">Member</span></span> |
| [<span data-ttu-id="a7c7c-121">body</span><span class="sxs-lookup"><span data-stu-id="a7c7c-121">body</span></span>](#body-bodyjavascriptapioutlookofficebody) | <span data-ttu-id="a7c7c-122">Element</span><span class="sxs-lookup"><span data-stu-id="a7c7c-122">Member</span></span> |
| [<span data-ttu-id="a7c7c-123">cc</span><span class="sxs-lookup"><span data-stu-id="a7c7c-123">cc</span></span>](#cc-arrayemailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsrecipientsjavascriptapioutlookofficerecipients) | <span data-ttu-id="a7c7c-124">Element</span><span class="sxs-lookup"><span data-stu-id="a7c7c-124">Member</span></span> |
| [<span data-ttu-id="a7c7c-125">conversationId</span><span class="sxs-lookup"><span data-stu-id="a7c7c-125">conversationId</span></span>](#nullable-conversationid-string) | <span data-ttu-id="a7c7c-126">Element</span><span class="sxs-lookup"><span data-stu-id="a7c7c-126">Member</span></span> |
| [<span data-ttu-id="a7c7c-127">dateTimeCreated</span><span class="sxs-lookup"><span data-stu-id="a7c7c-127">dateTimeCreated</span></span>](#datetimecreated-date) | <span data-ttu-id="a7c7c-128">Element</span><span class="sxs-lookup"><span data-stu-id="a7c7c-128">Member</span></span> |
| [<span data-ttu-id="a7c7c-129">dateTimeModified</span><span class="sxs-lookup"><span data-stu-id="a7c7c-129">dateTimeModified</span></span>](#datetimemodified-date) | <span data-ttu-id="a7c7c-130">Element</span><span class="sxs-lookup"><span data-stu-id="a7c7c-130">Member</span></span> |
| [<span data-ttu-id="a7c7c-131">end</span><span class="sxs-lookup"><span data-stu-id="a7c7c-131">end</span></span>](#end-datetimejavascriptapioutlookofficetime) | <span data-ttu-id="a7c7c-132">Element</span><span class="sxs-lookup"><span data-stu-id="a7c7c-132">Member</span></span> |
| [<span data-ttu-id="a7c7c-133">from</span><span class="sxs-lookup"><span data-stu-id="a7c7c-133">from</span></span>](#from-emailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsfromjavascriptapioutlookofficefrom) | <span data-ttu-id="a7c7c-134">Element</span><span class="sxs-lookup"><span data-stu-id="a7c7c-134">Member</span></span> |
| [<span data-ttu-id="a7c7c-135">internetMessageId</span><span class="sxs-lookup"><span data-stu-id="a7c7c-135">internetMessageId</span></span>](#internetmessageid-string) | <span data-ttu-id="a7c7c-136">Element</span><span class="sxs-lookup"><span data-stu-id="a7c7c-136">Member</span></span> |
| [<span data-ttu-id="a7c7c-137">itemClass</span><span class="sxs-lookup"><span data-stu-id="a7c7c-137">itemClass</span></span>](#itemclass-string) | <span data-ttu-id="a7c7c-138">Element</span><span class="sxs-lookup"><span data-stu-id="a7c7c-138">Member</span></span> |
| [<span data-ttu-id="a7c7c-139">itemId</span><span class="sxs-lookup"><span data-stu-id="a7c7c-139">itemId</span></span>](#nullable-itemid-string) | <span data-ttu-id="a7c7c-140">Element</span><span class="sxs-lookup"><span data-stu-id="a7c7c-140">Member</span></span> |
| [<span data-ttu-id="a7c7c-141">itemType</span><span class="sxs-lookup"><span data-stu-id="a7c7c-141">itemType</span></span>](#itemtype-officemailboxenumsitemtypejavascriptapioutlookofficemailboxenumsitemtype) | <span data-ttu-id="a7c7c-142">Element</span><span class="sxs-lookup"><span data-stu-id="a7c7c-142">Member</span></span> |
| [<span data-ttu-id="a7c7c-143">location</span><span class="sxs-lookup"><span data-stu-id="a7c7c-143">location</span></span>](#location-stringlocationjavascriptapioutlookofficelocation) | <span data-ttu-id="a7c7c-144">Element</span><span class="sxs-lookup"><span data-stu-id="a7c7c-144">Member</span></span> |
| [<span data-ttu-id="a7c7c-145">normalizedSubject</span><span class="sxs-lookup"><span data-stu-id="a7c7c-145">normalizedSubject</span></span>](#normalizedsubject-string) | <span data-ttu-id="a7c7c-146">Element</span><span class="sxs-lookup"><span data-stu-id="a7c7c-146">Member</span></span> |
| [<span data-ttu-id="a7c7c-147">notificationMessages</span><span class="sxs-lookup"><span data-stu-id="a7c7c-147">notificationMessages</span></span>](#notificationmessages-notificationmessagesjavascriptapioutlookofficenotificationmessages) | <span data-ttu-id="a7c7c-148">Element</span><span class="sxs-lookup"><span data-stu-id="a7c7c-148">Member</span></span> |
| [<span data-ttu-id="a7c7c-149">optionalAttendees</span><span class="sxs-lookup"><span data-stu-id="a7c7c-149">optionalAttendees</span></span>](#optionalattendees-arrayemailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsrecipientsjavascriptapioutlookofficerecipients) | <span data-ttu-id="a7c7c-150">Element</span><span class="sxs-lookup"><span data-stu-id="a7c7c-150">Member</span></span> |
| [<span data-ttu-id="a7c7c-151">organizer</span><span class="sxs-lookup"><span data-stu-id="a7c7c-151">organizer</span></span>](#organizer-emailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsorganizerjavascriptapioutlookofficeorganizer) | <span data-ttu-id="a7c7c-152">Element</span><span class="sxs-lookup"><span data-stu-id="a7c7c-152">Member</span></span> |
| [<span data-ttu-id="a7c7c-153">recurrence</span><span class="sxs-lookup"><span data-stu-id="a7c7c-153">recurrence</span></span>](#nullable-recurrence-recurrencejavascriptapioutlookofficerecurrence) | <span data-ttu-id="a7c7c-154">Element</span><span class="sxs-lookup"><span data-stu-id="a7c7c-154">Member</span></span> |
| [<span data-ttu-id="a7c7c-155">requiredAttendees</span><span class="sxs-lookup"><span data-stu-id="a7c7c-155">requiredAttendees</span></span>](#requiredattendees-arrayemailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsrecipientsjavascriptapioutlookofficerecipients) | <span data-ttu-id="a7c7c-156">Element</span><span class="sxs-lookup"><span data-stu-id="a7c7c-156">Member</span></span> |
| [<span data-ttu-id="a7c7c-157">sender</span><span class="sxs-lookup"><span data-stu-id="a7c7c-157">sender</span></span>](#sender-emailaddressdetailsjavascriptapioutlookofficeemailaddressdetails) | <span data-ttu-id="a7c7c-158">Element</span><span class="sxs-lookup"><span data-stu-id="a7c7c-158">Member</span></span> |
| [<span data-ttu-id="a7c7c-159">seriesId</span><span class="sxs-lookup"><span data-stu-id="a7c7c-159">seriesId</span></span>](#nullable-seriesid-string) | <span data-ttu-id="a7c7c-160">Element</span><span class="sxs-lookup"><span data-stu-id="a7c7c-160">Member</span></span> |
| [<span data-ttu-id="a7c7c-161">start</span><span class="sxs-lookup"><span data-stu-id="a7c7c-161">start</span></span>](#start-datetimejavascriptapioutlookofficetime) | <span data-ttu-id="a7c7c-162">Element</span><span class="sxs-lookup"><span data-stu-id="a7c7c-162">Member</span></span> |
| [<span data-ttu-id="a7c7c-163">subject</span><span class="sxs-lookup"><span data-stu-id="a7c7c-163">subject</span></span>](#subject-stringsubjectjavascriptapioutlookofficesubject) | <span data-ttu-id="a7c7c-164">Element</span><span class="sxs-lookup"><span data-stu-id="a7c7c-164">Member</span></span> |
| [<span data-ttu-id="a7c7c-165">to</span><span class="sxs-lookup"><span data-stu-id="a7c7c-165">to</span></span>](#to-arrayemailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsrecipientsjavascriptapioutlookofficerecipients) | <span data-ttu-id="a7c7c-166">Element</span><span class="sxs-lookup"><span data-stu-id="a7c7c-166">Member</span></span> |
| [<span data-ttu-id="a7c7c-167">addFileAttachmentAsync</span><span class="sxs-lookup"><span data-stu-id="a7c7c-167">addFileAttachmentAsync</span></span>](#addfileattachmentasyncuri-attachmentname-options-callback) | <span data-ttu-id="a7c7c-168">Methode</span><span class="sxs-lookup"><span data-stu-id="a7c7c-168">Method</span></span> |
| [<span data-ttu-id="a7c7c-169">addFileAttachmentFromBase64Async</span><span class="sxs-lookup"><span data-stu-id="a7c7c-169">addFileAttachmentFromBase64Async</span></span>](#addfileattachmentfrombase64asyncbase64file-attachmentname-options-callback) | <span data-ttu-id="a7c7c-170">Methode</span><span class="sxs-lookup"><span data-stu-id="a7c7c-170">Method</span></span> |
| [<span data-ttu-id="a7c7c-171">addHandlerAsync</span><span class="sxs-lookup"><span data-stu-id="a7c7c-171">addHandlerAsync</span></span>](#addhandlerasynceventtype-handler-options-callback) | <span data-ttu-id="a7c7c-172">Methode</span><span class="sxs-lookup"><span data-stu-id="a7c7c-172">Method</span></span> |
| [<span data-ttu-id="a7c7c-173">addItemAttachmentAsync</span><span class="sxs-lookup"><span data-stu-id="a7c7c-173">addItemAttachmentAsync</span></span>](#additemattachmentasyncitemid-attachmentname-options-callback) | <span data-ttu-id="a7c7c-174">Methode</span><span class="sxs-lookup"><span data-stu-id="a7c7c-174">Method</span></span> |
| [<span data-ttu-id="a7c7c-175">close</span><span class="sxs-lookup"><span data-stu-id="a7c7c-175">close</span></span>](#close) | <span data-ttu-id="a7c7c-176">Methode</span><span class="sxs-lookup"><span data-stu-id="a7c7c-176">Method</span></span> |
| [<span data-ttu-id="a7c7c-177">displayReplyAllForm</span><span class="sxs-lookup"><span data-stu-id="a7c7c-177">displayReplyAllForm</span></span>](#displayreplyallformformdata) | <span data-ttu-id="a7c7c-178">Methode</span><span class="sxs-lookup"><span data-stu-id="a7c7c-178">Method</span></span> |
| [<span data-ttu-id="a7c7c-179">displayReplyForm</span><span class="sxs-lookup"><span data-stu-id="a7c7c-179">displayReplyForm</span></span>](#displayreplyformformdata) | <span data-ttu-id="a7c7c-180">Methode</span><span class="sxs-lookup"><span data-stu-id="a7c7c-180">Method</span></span> |
| [<span data-ttu-id="a7c7c-181">getEntities</span><span class="sxs-lookup"><span data-stu-id="a7c7c-181">getEntities</span></span>](#getentities--entitiesjavascriptapioutlookofficeentities) | <span data-ttu-id="a7c7c-182">Methode</span><span class="sxs-lookup"><span data-stu-id="a7c7c-182">Method</span></span> |
| [<span data-ttu-id="a7c7c-183">getEntitiesByType</span><span class="sxs-lookup"><span data-stu-id="a7c7c-183">getEntitiesByType</span></span>](#getentitiesbytypeentitytype--nullable-arraystringcontactjavascriptapioutlookofficecontactmeetingsuggestionjavascriptapioutlookofficemeetingsuggestionphonenumberjavascriptapioutlookofficephonenumbertasksuggestionjavascriptapioutlookofficetasksuggestion) | <span data-ttu-id="a7c7c-184">Methode</span><span class="sxs-lookup"><span data-stu-id="a7c7c-184">Method</span></span> |
| [<span data-ttu-id="a7c7c-185">getFilteredEntitiesByName</span><span class="sxs-lookup"><span data-stu-id="a7c7c-185">getFilteredEntitiesByName</span></span>](#getfilteredentitiesbynamename--nullable-arraystringcontactjavascriptapioutlookofficecontactmeetingsuggestionjavascriptapioutlookofficemeetingsuggestionphonenumberjavascriptapioutlookofficephonenumbertasksuggestionjavascriptapioutlookofficetasksuggestion) | <span data-ttu-id="a7c7c-186">Methode</span><span class="sxs-lookup"><span data-stu-id="a7c7c-186">Method</span></span> |
| [<span data-ttu-id="a7c7c-187">getInitializationContextAsync</span><span class="sxs-lookup"><span data-stu-id="a7c7c-187">getInitializationContextAsync</span></span>](#getinitializationcontextasyncoptions-callback) | <span data-ttu-id="a7c7c-188">Methode</span><span class="sxs-lookup"><span data-stu-id="a7c7c-188">Method</span></span> |
| [<span data-ttu-id="a7c7c-189">getRegExMatches</span><span class="sxs-lookup"><span data-stu-id="a7c7c-189">getRegExMatches</span></span>](#getregexmatches--object) | <span data-ttu-id="a7c7c-190">Methode</span><span class="sxs-lookup"><span data-stu-id="a7c7c-190">Method</span></span> |
| [<span data-ttu-id="a7c7c-191">getRegExMatchesByName</span><span class="sxs-lookup"><span data-stu-id="a7c7c-191">getRegExMatchesByName</span></span>](#getregexmatchesbynamename--nullable-array-string-) | <span data-ttu-id="a7c7c-192">Methode</span><span class="sxs-lookup"><span data-stu-id="a7c7c-192">Method</span></span> |
| [<span data-ttu-id="a7c7c-193">getSelectedDataAsync</span><span class="sxs-lookup"><span data-stu-id="a7c7c-193">getSelectedDataAsync</span></span>](#getselecteddataasynccoerciontype-options-callback--string) | <span data-ttu-id="a7c7c-194">Methode</span><span class="sxs-lookup"><span data-stu-id="a7c7c-194">Method</span></span> |
| [<span data-ttu-id="a7c7c-195">getSelectedEntities</span><span class="sxs-lookup"><span data-stu-id="a7c7c-195">getSelectedEntities</span></span>](#getselectedentities--entitiesjavascriptapioutlookofficeentities) | <span data-ttu-id="a7c7c-196">Methode</span><span class="sxs-lookup"><span data-stu-id="a7c7c-196">Method</span></span> |
| [<span data-ttu-id="a7c7c-197">getSelectedRegExMatches</span><span class="sxs-lookup"><span data-stu-id="a7c7c-197">getSelectedRegExMatches</span></span>](#getselectedregexmatches--object) | <span data-ttu-id="a7c7c-198">Methode</span><span class="sxs-lookup"><span data-stu-id="a7c7c-198">Method</span></span> |
| [<span data-ttu-id="a7c7c-199">getSharedPropertiesAsync</span><span class="sxs-lookup"><span data-stu-id="a7c7c-199">getSharedPropertiesAsync</span></span>](#getsharedpropertiesasyncoptions-callback) | <span data-ttu-id="a7c7c-200">Methode</span><span class="sxs-lookup"><span data-stu-id="a7c7c-200">Method</span></span> |
| [<span data-ttu-id="a7c7c-201">loadCustomPropertiesAsync</span><span class="sxs-lookup"><span data-stu-id="a7c7c-201">loadCustomPropertiesAsync</span></span>](#loadcustompropertiesasynccallback-usercontext) | <span data-ttu-id="a7c7c-202">Methode</span><span class="sxs-lookup"><span data-stu-id="a7c7c-202">Method</span></span> |
| [<span data-ttu-id="a7c7c-203">removeAttachmentAsync</span><span class="sxs-lookup"><span data-stu-id="a7c7c-203">removeAttachmentAsync</span></span>](#removeattachmentasyncattachmentid-options-callback) | <span data-ttu-id="a7c7c-204">Methode</span><span class="sxs-lookup"><span data-stu-id="a7c7c-204">Method</span></span> |
| [<span data-ttu-id="a7c7c-205">removeHandlerAsync</span><span class="sxs-lookup"><span data-stu-id="a7c7c-205">removeHandlerAsync</span></span>](#removehandlerasynceventtype-handler-options-callback) | <span data-ttu-id="a7c7c-206">Methode</span><span class="sxs-lookup"><span data-stu-id="a7c7c-206">Method</span></span> |
| [<span data-ttu-id="a7c7c-207">saveAsync</span><span class="sxs-lookup"><span data-stu-id="a7c7c-207">saveAsync</span></span>](#saveasyncoptions-callback) | <span data-ttu-id="a7c7c-208">Methode</span><span class="sxs-lookup"><span data-stu-id="a7c7c-208">Method</span></span> |
| [<span data-ttu-id="a7c7c-209">setSelectedDataAsync</span><span class="sxs-lookup"><span data-stu-id="a7c7c-209">setSelectedDataAsync</span></span>](#setselecteddataasyncdata-options-callback) | <span data-ttu-id="a7c7c-210">Methode</span><span class="sxs-lookup"><span data-stu-id="a7c7c-210">Method</span></span> |

### <a name="example"></a><span data-ttu-id="a7c7c-211">Beispiel</span><span class="sxs-lookup"><span data-stu-id="a7c7c-211">Example</span></span>

<span data-ttu-id="a7c7c-212">Im folgenden JavaScript-Codebeispiel wird der Zugriff auf die `subject`-Eigenschaft des aktuellen Elements in Outlook veranschaulicht.</span><span class="sxs-lookup"><span data-stu-id="a7c7c-212">The following JavaScript code example shows how to access the `subject` property of the current item in Outlook.</span></span>

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

### <a name="members"></a><span data-ttu-id="a7c7c-213">Elemente</span><span class="sxs-lookup"><span data-stu-id="a7c7c-213">Members</span></span>

#### <a name="attachments-arrayattachmentdetailsjavascriptapioutlookofficeattachmentdetails"></a><span data-ttu-id="a7c7c-214">attachments :Array.<[AttachmentDetails](/javascript/api/outlook/office.attachmentdetails)></span><span class="sxs-lookup"><span data-stu-id="a7c7c-214">attachments :Array.<[AttachmentDetails](/javascript/api/outlook/office.attachmentdetails)></span></span>

<span data-ttu-id="a7c7c-p102">Ruft ein Array mit Anlagen für das Element ab. Nur Lesemodus.</span><span class="sxs-lookup"><span data-stu-id="a7c7c-p102">Gets an array of attachments for the item. Read mode only.</span></span>

> [!NOTE]
> <span data-ttu-id="a7c7c-217">Bestimmte Dateitypen werden aufgrund von Sicherheitsproblemen von Outlook blockiert und daher nicht zurückgegeben werden.</span><span class="sxs-lookup"><span data-stu-id="a7c7c-217">Certain types of files are blocked by Outlook due to potential security issues and are therefore not returned.</span></span> <span data-ttu-id="a7c7c-218">Weitere Informationen finden Sie unter [Blockierte Anlagen in Outlook](https://support.office.com/article/Blocked-attachments-in-Outlook-434752E1-02D3-4E90-9124-8B81E49A8519).</span><span class="sxs-lookup"><span data-stu-id="a7c7c-218">For more information, see [Blocked attachments in Outlook](https://support.office.com/article/Blocked-attachments-in-Outlook-434752E1-02D3-4E90-9124-8B81E49A8519).</span></span>

##### <a name="type"></a><span data-ttu-id="a7c7c-219">Typ:</span><span class="sxs-lookup"><span data-stu-id="a7c7c-219">Type:</span></span>

*   <span data-ttu-id="a7c7c-220">Array.<[AttachmentDetails](/javascript/api/outlook/office.attachmentdetails)></span><span class="sxs-lookup"><span data-stu-id="a7c7c-220">Array.<[AttachmentDetails](/javascript/api/outlook/office.attachmentdetails)></span></span>

##### <a name="requirements"></a><span data-ttu-id="a7c7c-221">Anforderungen</span><span class="sxs-lookup"><span data-stu-id="a7c7c-221">Requirements</span></span>

|<span data-ttu-id="a7c7c-222">Anforderung</span><span class="sxs-lookup"><span data-stu-id="a7c7c-222">Requirement</span></span>|<span data-ttu-id="a7c7c-223">Wert</span><span class="sxs-lookup"><span data-stu-id="a7c7c-223">Value</span></span>|
|---|---|
|[<span data-ttu-id="a7c7c-224">Mindestversion des Postfachanforderungssatzes</span><span class="sxs-lookup"><span data-stu-id="a7c7c-224">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="a7c7c-225">1.0</span><span class="sxs-lookup"><span data-stu-id="a7c7c-225">1.0</span></span>|
|[<span data-ttu-id="a7c7c-226">Mindestberechtigungsstufe</span><span class="sxs-lookup"><span data-stu-id="a7c7c-226">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="a7c7c-227">ReadItem</span><span class="sxs-lookup"><span data-stu-id="a7c7c-227">ReadItem</span></span>|
|[<span data-ttu-id="a7c7c-228">Zutreffender Outlook-Modus</span><span class="sxs-lookup"><span data-stu-id="a7c7c-228">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="a7c7c-229">Lesen</span><span class="sxs-lookup"><span data-stu-id="a7c7c-229">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="a7c7c-230">Beispiel</span><span class="sxs-lookup"><span data-stu-id="a7c7c-230">Example</span></span>

<span data-ttu-id="a7c7c-231">Mit dem folgende Code wird eine HTML-Zeichenfolge mit Details aller Anlagen für das aktuelle Element erstellt.</span><span class="sxs-lookup"><span data-stu-id="a7c7c-231">The following code builds an HTML string with details of all attachments on the current item.</span></span>

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

####  <a name="bcc-recipientsjavascriptapioutlookofficerecipients"></a><span data-ttu-id="a7c7c-232">bcc :[Recipients](/javascript/api/outlook/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="a7c7c-232">bcc :[Recipients](/javascript/api/outlook/office.recipients)</span></span>

<span data-ttu-id="a7c7c-233">Ruft ein Objekt, das Methoden zum Abrufen oder aktualisieren die Empfänger in der Zeile Bcc (blind Carbon Copy, Blindkopie) einer Nachricht bereitstellt.</span><span class="sxs-lookup"><span data-stu-id="a7c7c-233">Gets an object that provides methods to get or update the recipients on the Bcc (blind carbon copy) line of a message.</span></span> <span data-ttu-id="a7c7c-234">Nur Verfassenmodus.</span><span class="sxs-lookup"><span data-stu-id="a7c7c-234">Compose mode only.</span></span>

##### <a name="type"></a><span data-ttu-id="a7c7c-235">Typ:</span><span class="sxs-lookup"><span data-stu-id="a7c7c-235">Type:</span></span>

*   [<span data-ttu-id="a7c7c-236">Recipients</span><span class="sxs-lookup"><span data-stu-id="a7c7c-236">Recipients</span></span>](/javascript/api/outlook/office.recipients)

##### <a name="requirements"></a><span data-ttu-id="a7c7c-237">Anforderungen</span><span class="sxs-lookup"><span data-stu-id="a7c7c-237">Requirements</span></span>

|<span data-ttu-id="a7c7c-238">Anforderung</span><span class="sxs-lookup"><span data-stu-id="a7c7c-238">Requirement</span></span>|<span data-ttu-id="a7c7c-239">Wert</span><span class="sxs-lookup"><span data-stu-id="a7c7c-239">Value</span></span>|
|---|---|
|[<span data-ttu-id="a7c7c-240">Mindestversion des Postfachanforderungssatzes</span><span class="sxs-lookup"><span data-stu-id="a7c7c-240">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="a7c7c-241">1.1</span><span class="sxs-lookup"><span data-stu-id="a7c7c-241">1.1</span></span>|
|[<span data-ttu-id="a7c7c-242">Mindestberechtigungsstufe</span><span class="sxs-lookup"><span data-stu-id="a7c7c-242">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="a7c7c-243">ReadItem</span><span class="sxs-lookup"><span data-stu-id="a7c7c-243">ReadItem</span></span>|
|[<span data-ttu-id="a7c7c-244">Zutreffender Outlook-Modus</span><span class="sxs-lookup"><span data-stu-id="a7c7c-244">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="a7c7c-245">Verfassen</span><span class="sxs-lookup"><span data-stu-id="a7c7c-245">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="a7c7c-246">Beispiel</span><span class="sxs-lookup"><span data-stu-id="a7c7c-246">Example</span></span>

```
Office.context.mailbox.item.bcc.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.bcc.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.bcc.getAsync(callback);

function callback(asyncResult) {
  var arrayOfBccRecipients = asyncResult.value;
}
```

####  <a name="body-bodyjavascriptapioutlookofficebody"></a><span data-ttu-id="a7c7c-247">body :[Body](/javascript/api/outlook/office.body)</span><span class="sxs-lookup"><span data-stu-id="a7c7c-247">body :[Body](/javascript/api/outlook/office.body)</span></span>

<span data-ttu-id="a7c7c-248">Ruft ein Objekt ab, das Methoden zum Bearbeiten des Textkörpers eines Elements bereitstellt.</span><span class="sxs-lookup"><span data-stu-id="a7c7c-248">Gets an object that provides methods for manipulating the body of an item.</span></span>

##### <a name="type"></a><span data-ttu-id="a7c7c-249">Typ:</span><span class="sxs-lookup"><span data-stu-id="a7c7c-249">Type:</span></span>

*   [<span data-ttu-id="a7c7c-250">Body</span><span class="sxs-lookup"><span data-stu-id="a7c7c-250">Body</span></span>](/javascript/api/outlook/office.body)

##### <a name="requirements"></a><span data-ttu-id="a7c7c-251">Anforderungen</span><span class="sxs-lookup"><span data-stu-id="a7c7c-251">Requirements</span></span>

|<span data-ttu-id="a7c7c-252">Anforderung</span><span class="sxs-lookup"><span data-stu-id="a7c7c-252">Requirement</span></span>|<span data-ttu-id="a7c7c-253">Wert</span><span class="sxs-lookup"><span data-stu-id="a7c7c-253">Value</span></span>|
|---|---|
|[<span data-ttu-id="a7c7c-254">Mindestversion des Postfachanforderungssatzes</span><span class="sxs-lookup"><span data-stu-id="a7c7c-254">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="a7c7c-255">1.1</span><span class="sxs-lookup"><span data-stu-id="a7c7c-255">1.1</span></span>|
|[<span data-ttu-id="a7c7c-256">Mindestberechtigungsstufe</span><span class="sxs-lookup"><span data-stu-id="a7c7c-256">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="a7c7c-257">ReadItem</span><span class="sxs-lookup"><span data-stu-id="a7c7c-257">ReadItem</span></span>|
|[<span data-ttu-id="a7c7c-258">Zutreffender Outlook-Modus</span><span class="sxs-lookup"><span data-stu-id="a7c7c-258">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="a7c7c-259">Verfassen oder Lesen</span><span class="sxs-lookup"><span data-stu-id="a7c7c-259">Compose or read</span></span>|

####  <a name="cc-arrayemailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsrecipientsjavascriptapioutlookofficerecipients"></a><span data-ttu-id="a7c7c-260">cc: Array. <[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)>|[Empfänger](/javascript/api/outlook/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="a7c7c-260">cc :Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook/office.recipients)</span></span>

<span data-ttu-id="a7c7c-261">Ermöglicht den Zugriff auf die (Carbon Copy, Blindkopie) Cc-Empfänger einer Nachricht.</span><span class="sxs-lookup"><span data-stu-id="a7c7c-261">Provides access to the Cc (carbon copy) recipients of a message.</span></span> <span data-ttu-id="a7c7c-262">Der Typ des Objekts und die Zugriffsebene, hängt von den Modus des aktuellen Elements ab.</span><span class="sxs-lookup"><span data-stu-id="a7c7c-262">The type of object and level of access depends on the mode of the current item.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="a7c7c-263">Lesemodus</span><span class="sxs-lookup"><span data-stu-id="a7c7c-263">Read mode</span></span>

<span data-ttu-id="a7c7c-p106">Die `cc`-Eigenschaft gibt ein Array mit einem `EmailAddressDetails`-Objekt für jeden Empfänger in der **Cc**-Zeile der Nachricht zurück. Die Auflistung ist auf höchstens 100 Elemente beschränkt.</span><span class="sxs-lookup"><span data-stu-id="a7c7c-p106">The `cc` property returns an array that contains an `EmailAddressDetails` object for each recipient listed on the **Cc** line of the message. The collection is limited to a maximum of 100 members.</span></span>

##### <a name="compose-mode"></a><span data-ttu-id="a7c7c-266">Verfassenmodus</span><span class="sxs-lookup"><span data-stu-id="a7c7c-266">Compose mode</span></span>

<span data-ttu-id="a7c7c-267">Die `cc` -Eigenschaft gibt eine `Recipients` -Objekt, das Methoden zum Abrufen oder aktualisieren die Empfänger in der Zeile **Cc** der Nachricht bereitstellt.</span><span class="sxs-lookup"><span data-stu-id="a7c7c-267">The `cc` property returns a `Recipients` object that provides methods to get or update the recipients on the **Cc** line of the message.</span></span>

##### <a name="type"></a><span data-ttu-id="a7c7c-268">Typ:</span><span class="sxs-lookup"><span data-stu-id="a7c7c-268">Type:</span></span>

*   <span data-ttu-id="a7c7c-269">Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="a7c7c-269">Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook/office.recipients)</span></span>

##### <a name="requirements"></a><span data-ttu-id="a7c7c-270">Anforderungen</span><span class="sxs-lookup"><span data-stu-id="a7c7c-270">Requirements</span></span>

|<span data-ttu-id="a7c7c-271">Anforderung</span><span class="sxs-lookup"><span data-stu-id="a7c7c-271">Requirement</span></span>|<span data-ttu-id="a7c7c-272">Wert</span><span class="sxs-lookup"><span data-stu-id="a7c7c-272">Value</span></span>|
|---|---|
|[<span data-ttu-id="a7c7c-273">Mindestversion des Postfachanforderungssatzes</span><span class="sxs-lookup"><span data-stu-id="a7c7c-273">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="a7c7c-274">1.0</span><span class="sxs-lookup"><span data-stu-id="a7c7c-274">1.0</span></span>|
|[<span data-ttu-id="a7c7c-275">Mindestberechtigungsstufe</span><span class="sxs-lookup"><span data-stu-id="a7c7c-275">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="a7c7c-276">ReadItem</span><span class="sxs-lookup"><span data-stu-id="a7c7c-276">ReadItem</span></span>|
|[<span data-ttu-id="a7c7c-277">Zutreffender Outlook-Modus</span><span class="sxs-lookup"><span data-stu-id="a7c7c-277">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="a7c7c-278">Verfassen oder Lesen</span><span class="sxs-lookup"><span data-stu-id="a7c7c-278">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="a7c7c-279">Beispiel</span><span class="sxs-lookup"><span data-stu-id="a7c7c-279">Example</span></span>

```
Office.context.mailbox.item.cc.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.cc.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.cc.getAsync(callback);

function callback(asyncResult) {
  var arrayOfCcRecipients = asyncResult.value;
}
```

####  <a name="nullable-conversationid-string"></a><span data-ttu-id="a7c7c-280">(nullable) conversationId :String</span><span class="sxs-lookup"><span data-stu-id="a7c7c-280">(nullable) conversationId :String</span></span>

<span data-ttu-id="a7c7c-281">Ruft einen Bezeichner für die E-Mail-Unterhaltung ab, in der eine bestimmte Nachricht enthalten ist.</span><span class="sxs-lookup"><span data-stu-id="a7c7c-281">Gets an identifier for the email conversation that contains a particular message.</span></span>

<span data-ttu-id="a7c7c-p107">Sie können für diese Eigenschaft eine ganze Zahl abrufen, wenn Ihre Mail-App in Formularen zum Lesen oder Antworten in Formularen zum Verfassen aktiviert wird. Wenn der Benutzer den Betreff der Antwortnachricht ändert, ändert sich beim Versenden die Konversations-ID für die entsprechende Nachricht, und der Wert, den Sie vorher bezogen haben, trifft nicht länger zu.</span><span class="sxs-lookup"><span data-stu-id="a7c7c-p107">You can get an integer for this property if your mail app is activated in read forms or responses in compose forms. If subsequently the user changes the subject of the reply message, upon sending the reply, the conversation ID for that message will change and that value you obtained earlier will no longer apply.</span></span>

<span data-ttu-id="a7c7c-p108">Sie erhalten in einem Formular zum Verfassen für diese Eigenschaft für ein neues Element null. Wenn der Benutzer einen Betreff festlegt und das Element speichert, gibt die `conversationId`-Eigenschaft einen Wert zurück.</span><span class="sxs-lookup"><span data-stu-id="a7c7c-p108">You get null for this property for a new item in a compose form. If the user sets a subject and saves the item, the `conversationId` property will return a value.</span></span>

##### <a name="type"></a><span data-ttu-id="a7c7c-286">Typ:</span><span class="sxs-lookup"><span data-stu-id="a7c7c-286">Type:</span></span>

*   <span data-ttu-id="a7c7c-287">String</span><span class="sxs-lookup"><span data-stu-id="a7c7c-287">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="a7c7c-288">Anforderungen</span><span class="sxs-lookup"><span data-stu-id="a7c7c-288">Requirements</span></span>

|<span data-ttu-id="a7c7c-289">Anforderung</span><span class="sxs-lookup"><span data-stu-id="a7c7c-289">Requirement</span></span>|<span data-ttu-id="a7c7c-290">Wert</span><span class="sxs-lookup"><span data-stu-id="a7c7c-290">Value</span></span>|
|---|---|
|[<span data-ttu-id="a7c7c-291">Mindestversion des Postfachanforderungssatzes</span><span class="sxs-lookup"><span data-stu-id="a7c7c-291">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="a7c7c-292">1.0</span><span class="sxs-lookup"><span data-stu-id="a7c7c-292">1.0</span></span>|
|[<span data-ttu-id="a7c7c-293">Mindestberechtigungsstufe</span><span class="sxs-lookup"><span data-stu-id="a7c7c-293">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="a7c7c-294">ReadItem</span><span class="sxs-lookup"><span data-stu-id="a7c7c-294">ReadItem</span></span>|
|[<span data-ttu-id="a7c7c-295">Zutreffender Outlook-Modus</span><span class="sxs-lookup"><span data-stu-id="a7c7c-295">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="a7c7c-296">Verfassen oder Lesen</span><span class="sxs-lookup"><span data-stu-id="a7c7c-296">Compose or read</span></span>|

#### <a name="datetimecreated-date"></a><span data-ttu-id="a7c7c-297">dateTimeCreated :Date</span><span class="sxs-lookup"><span data-stu-id="a7c7c-297">dateTimeCreated :Date</span></span>

<span data-ttu-id="a7c7c-p109">Ruft das Datum und die Uhrzeit der Erstellung eines Elements ab. Nur Lesemodus.</span><span class="sxs-lookup"><span data-stu-id="a7c7c-p109">Gets the date and time that an item was created. Read mode only.</span></span>

##### <a name="type"></a><span data-ttu-id="a7c7c-300">Typ:</span><span class="sxs-lookup"><span data-stu-id="a7c7c-300">Type:</span></span>

*   <span data-ttu-id="a7c7c-301">Datum</span><span class="sxs-lookup"><span data-stu-id="a7c7c-301">Date</span></span>

##### <a name="requirements"></a><span data-ttu-id="a7c7c-302">Anforderungen</span><span class="sxs-lookup"><span data-stu-id="a7c7c-302">Requirements</span></span>

|<span data-ttu-id="a7c7c-303">Anforderung</span><span class="sxs-lookup"><span data-stu-id="a7c7c-303">Requirement</span></span>|<span data-ttu-id="a7c7c-304">Wert</span><span class="sxs-lookup"><span data-stu-id="a7c7c-304">Value</span></span>|
|---|---|
|[<span data-ttu-id="a7c7c-305">Mindestversion des Postfachanforderungssatzes</span><span class="sxs-lookup"><span data-stu-id="a7c7c-305">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="a7c7c-306">1.0</span><span class="sxs-lookup"><span data-stu-id="a7c7c-306">1.0</span></span>|
|[<span data-ttu-id="a7c7c-307">Mindestberechtigungsstufe</span><span class="sxs-lookup"><span data-stu-id="a7c7c-307">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="a7c7c-308">ReadItem</span><span class="sxs-lookup"><span data-stu-id="a7c7c-308">ReadItem</span></span>|
|[<span data-ttu-id="a7c7c-309">Zutreffender Outlook-Modus</span><span class="sxs-lookup"><span data-stu-id="a7c7c-309">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="a7c7c-310">Lesen</span><span class="sxs-lookup"><span data-stu-id="a7c7c-310">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="a7c7c-311">Beispiel</span><span class="sxs-lookup"><span data-stu-id="a7c7c-311">Example</span></span>

```
var created = Office.context.mailbox.item.dateTimeCreated;
```

#### <a name="datetimemodified-date"></a><span data-ttu-id="a7c7c-312">dateTimeModified :Date</span><span class="sxs-lookup"><span data-stu-id="a7c7c-312">dateTimeModified :Date</span></span>

<span data-ttu-id="a7c7c-p110">Ruft das Datum und die Uhrzeit der letzten Änderung eines Elements ab. Nur Lesemodus.</span><span class="sxs-lookup"><span data-stu-id="a7c7c-p110">Gets the date and time that an item was last modified. Read mode only.</span></span>

> [!NOTE]
> <span data-ttu-id="a7c7c-315">Dieser Member wird in Outlook für iOS oder Outlook für Android nicht unterstützt.</span><span class="sxs-lookup"><span data-stu-id="a7c7c-315">This member is not supported in Outlook for iOS or Outlook for Android.</span></span>

##### <a name="type"></a><span data-ttu-id="a7c7c-316">Typ:</span><span class="sxs-lookup"><span data-stu-id="a7c7c-316">Type:</span></span>

*   <span data-ttu-id="a7c7c-317">Datum</span><span class="sxs-lookup"><span data-stu-id="a7c7c-317">Date</span></span>

##### <a name="requirements"></a><span data-ttu-id="a7c7c-318">Anforderungen</span><span class="sxs-lookup"><span data-stu-id="a7c7c-318">Requirements</span></span>

|<span data-ttu-id="a7c7c-319">Anforderung</span><span class="sxs-lookup"><span data-stu-id="a7c7c-319">Requirement</span></span>|<span data-ttu-id="a7c7c-320">Wert</span><span class="sxs-lookup"><span data-stu-id="a7c7c-320">Value</span></span>|
|---|---|
|[<span data-ttu-id="a7c7c-321">Mindestversion des Postfachanforderungssatzes</span><span class="sxs-lookup"><span data-stu-id="a7c7c-321">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="a7c7c-322">1.0</span><span class="sxs-lookup"><span data-stu-id="a7c7c-322">1.0</span></span>|
|[<span data-ttu-id="a7c7c-323">Mindestberechtigungsstufe</span><span class="sxs-lookup"><span data-stu-id="a7c7c-323">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="a7c7c-324">ReadItem</span><span class="sxs-lookup"><span data-stu-id="a7c7c-324">ReadItem</span></span>|
|[<span data-ttu-id="a7c7c-325">Zutreffender Outlook-Modus</span><span class="sxs-lookup"><span data-stu-id="a7c7c-325">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="a7c7c-326">Lesen</span><span class="sxs-lookup"><span data-stu-id="a7c7c-326">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="a7c7c-327">Beispiel</span><span class="sxs-lookup"><span data-stu-id="a7c7c-327">Example</span></span>

```
var modified = Office.context.mailbox.item.dateTimeModified;
```

####  <a name="end-datetimejavascriptapioutlookofficetime"></a><span data-ttu-id="a7c7c-328">end :Date|[Time](/javascript/api/outlook/office.time)</span><span class="sxs-lookup"><span data-stu-id="a7c7c-328">end :Date|[Time](/javascript/api/outlook/office.time)</span></span>

<span data-ttu-id="a7c7c-329">Ruft Datum und Zeit für das Ende des Termins ab oder legt diese fest.</span><span class="sxs-lookup"><span data-stu-id="a7c7c-329">Gets or sets the date and time that the appointment is to end.</span></span>

<span data-ttu-id="a7c7c-p111">Die `end`-Eigenschaft wird als Datums- und Uhrzeitwert in UTC ausgedrückt. Sie können die [`convertToLocalClientTime`](office.context.mailbox.md#converttolocalclienttimetimevalue--localclienttimejavascriptapioutlookofficelocalclienttime)-Methode verwenden, um den Endeigenschaftswert in den lokalen Uhrzeit- und Datumswert des Clients umzuwandeln.</span><span class="sxs-lookup"><span data-stu-id="a7c7c-p111">The `end` property is expressed as a Coordinated Universal Time (UTC) date and time value. You can use the [`convertToLocalClientTime`](office.context.mailbox.md#converttolocalclienttimetimevalue--localclienttimejavascriptapioutlookofficelocalclienttime) method to convert the end property value to the client’s local date and time.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="a7c7c-332">Lesemodus</span><span class="sxs-lookup"><span data-stu-id="a7c7c-332">Read mode</span></span>

<span data-ttu-id="a7c7c-333">Die `end`-Eigenschaft gibt ein `Date`-Objekt zurück.</span><span class="sxs-lookup"><span data-stu-id="a7c7c-333">The `end` property returns a `Date` object.</span></span>

##### <a name="compose-mode"></a><span data-ttu-id="a7c7c-334">Verfassenmodus</span><span class="sxs-lookup"><span data-stu-id="a7c7c-334">Compose mode</span></span>

<span data-ttu-id="a7c7c-335">Die `end`-Eigenschaft gibt ein `Time`-Objekt zurück.</span><span class="sxs-lookup"><span data-stu-id="a7c7c-335">The `end` property returns a `Time` object.</span></span>

<span data-ttu-id="a7c7c-336">Wenn Sie die [`Time.setAsync`](/javascript/api/outlook/office.time#setasync-datetime--options--callback-)-Methode verwenden, um die Endzeit im Verfassenmodus festzulegen, sollten Sie die [`convertToUtcClientTime`](office.context.mailbox.md#converttoutcclienttimeinput--date)-Methode verwenden, um die Ortszeit auf dem Client für den Server in UTC umzuwandeln.</span><span class="sxs-lookup"><span data-stu-id="a7c7c-336">When you use the [`Time.setAsync`](/javascript/api/outlook/office.time#setasync-datetime--options--callback-) method to set the end time, you should use the [`convertToUtcClientTime`](office.context.mailbox.md#converttoutcclienttimeinput--date) method to convert the local time on the client to UTC for the server.</span></span>

##### <a name="type"></a><span data-ttu-id="a7c7c-337">Typ:</span><span class="sxs-lookup"><span data-stu-id="a7c7c-337">Type:</span></span>

*   <span data-ttu-id="a7c7c-338">Date | [Time](/javascript/api/outlook/office.time)</span><span class="sxs-lookup"><span data-stu-id="a7c7c-338">Date | [Time](/javascript/api/outlook/office.time)</span></span>

##### <a name="requirements"></a><span data-ttu-id="a7c7c-339">Anforderungen</span><span class="sxs-lookup"><span data-stu-id="a7c7c-339">Requirements</span></span>

|<span data-ttu-id="a7c7c-340">Anforderung</span><span class="sxs-lookup"><span data-stu-id="a7c7c-340">Requirement</span></span>|<span data-ttu-id="a7c7c-341">Wert</span><span class="sxs-lookup"><span data-stu-id="a7c7c-341">Value</span></span>|
|---|---|
|[<span data-ttu-id="a7c7c-342">Mindestversion des Postfachanforderungssatzes</span><span class="sxs-lookup"><span data-stu-id="a7c7c-342">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="a7c7c-343">1.0</span><span class="sxs-lookup"><span data-stu-id="a7c7c-343">1.0</span></span>|
|[<span data-ttu-id="a7c7c-344">Mindestberechtigungsstufe</span><span class="sxs-lookup"><span data-stu-id="a7c7c-344">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="a7c7c-345">ReadItem</span><span class="sxs-lookup"><span data-stu-id="a7c7c-345">ReadItem</span></span>|
|[<span data-ttu-id="a7c7c-346">Zutreffender Outlook-Modus</span><span class="sxs-lookup"><span data-stu-id="a7c7c-346">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="a7c7c-347">Verfassen oder Lesen</span><span class="sxs-lookup"><span data-stu-id="a7c7c-347">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="a7c7c-348">Beispiel</span><span class="sxs-lookup"><span data-stu-id="a7c7c-348">Example</span></span>

<span data-ttu-id="a7c7c-349">Im folgenden Beispiel wird die Endzeit eines Termins im Verfassenmodus mithilfe der [`setAsync`](/javascript/api/outlook/office.time#setasync-datetime--options--callback-)-Methode des `Time`-Objekts festgelegt.</span><span class="sxs-lookup"><span data-stu-id="a7c7c-349">The following example sets the end time of an appointment in compose mode by using the [`setAsync`](/javascript/api/outlook/office.time#setasync-datetime--options--callback-) method of the `Time` object.</span></span>

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

#### <a name="from-emailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsfromjavascriptapioutlookofficefrom"></a><span data-ttu-id="a7c7c-350">von:[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)|[aus](/javascript/api/outlook/office.from)</span><span class="sxs-lookup"><span data-stu-id="a7c7c-350">from :[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)|[From](/javascript/api/outlook/office.from)</span></span>

<span data-ttu-id="a7c7c-351">Ruft die E-Mail-Adresse des Absenders einer Nachricht ab.</span><span class="sxs-lookup"><span data-stu-id="a7c7c-351">Gets the email address of the sender of a message.</span></span>

<span data-ttu-id="a7c7c-p112">Die `from`- und [`sender`](#sender-emailaddressdetailsjavascriptapioutlookofficeemailaddressdetails)-Eigenschaften stellen dieselbe Person dar, außer die Nachricht wurde von einem Delegaten gesendet. In diesem Fall stellt die `from`-Eigenschaft die Stellvertretung dar, und die sender-Eigenschaft stellt den Delegaten dar.</span><span class="sxs-lookup"><span data-stu-id="a7c7c-p112">The `from` and [`sender`](#sender-emailaddressdetailsjavascriptapioutlookofficeemailaddressdetails) properties represent the same person unless the message is sent by a delegate. In that case, the `from` property represents the delegator, and the sender property represents the delegate.</span></span>

> [!NOTE]
> <span data-ttu-id="a7c7c-354">Die `recipientType` -Eigenschaft des der `EmailAddressDetails` -Objekts der `from` -Eigenschaft ist `undefined`.</span><span class="sxs-lookup"><span data-stu-id="a7c7c-354">The `recipientType` property of the `EmailAddressDetails` object in the `from` property is `undefined`.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="a7c7c-355">Lesemodus</span><span class="sxs-lookup"><span data-stu-id="a7c7c-355">Read mode</span></span>

<span data-ttu-id="a7c7c-356">Die `from` -Eigenschaft gibt eine `EmailAddressDetails` Objekt.</span><span class="sxs-lookup"><span data-stu-id="a7c7c-356">The `from` property returns an `EmailAddressDetails` object.</span></span>

```
var subject = Office.context.mailbox.item.from;
```

##### <a name="compose-mode"></a><span data-ttu-id="a7c7c-357">Verfassenmodus</span><span class="sxs-lookup"><span data-stu-id="a7c7c-357">Compose mode</span></span>

<span data-ttu-id="a7c7c-358">Die `from` -Eigenschaft gibt eine `From` -Objekt, eine Methode zum Abrufen bietet, der vom Wert.</span><span class="sxs-lookup"><span data-stu-id="a7c7c-358">The `from` property returns a `From` object that provides a method to get the from value.</span></span>

```
Office.context.mailbox.item.from.getAsync(callback);

function callback(asyncResult) {
  var from = asyncResult.value;
}
```

##### <a name="type"></a><span data-ttu-id="a7c7c-359">Typ:</span><span class="sxs-lookup"><span data-stu-id="a7c7c-359">Type:</span></span>

*   <span data-ttu-id="a7c7c-360">[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails) | [aus](/javascript/api/outlook/office.from)</span><span class="sxs-lookup"><span data-stu-id="a7c7c-360">[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails) | [From](/javascript/api/outlook/office.from)</span></span>

##### <a name="requirements"></a><span data-ttu-id="a7c7c-361">Anforderungen</span><span class="sxs-lookup"><span data-stu-id="a7c7c-361">Requirements</span></span>

|<span data-ttu-id="a7c7c-362">Anforderung</span><span class="sxs-lookup"><span data-stu-id="a7c7c-362">Requirement</span></span>|||
|---|---|---|
|[<span data-ttu-id="a7c7c-363">Mindestversion des Postfachanforderungssatzes</span><span class="sxs-lookup"><span data-stu-id="a7c7c-363">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="a7c7c-364">1.0</span><span class="sxs-lookup"><span data-stu-id="a7c7c-364">1.0</span></span>|<span data-ttu-id="a7c7c-365">1.7</span><span class="sxs-lookup"><span data-stu-id="a7c7c-365">1.7</span></span>|
|[<span data-ttu-id="a7c7c-366">Mindestberechtigungsstufe</span><span class="sxs-lookup"><span data-stu-id="a7c7c-366">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="a7c7c-367">ReadItem</span><span class="sxs-lookup"><span data-stu-id="a7c7c-367">ReadItem</span></span>|<span data-ttu-id="a7c7c-368">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="a7c7c-368">ReadWriteItem</span></span>|
|[<span data-ttu-id="a7c7c-369">Zutreffender Outlook-Modus</span><span class="sxs-lookup"><span data-stu-id="a7c7c-369">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="a7c7c-370">Lesen</span><span class="sxs-lookup"><span data-stu-id="a7c7c-370">Read</span></span>|<span data-ttu-id="a7c7c-371">Verfassen</span><span class="sxs-lookup"><span data-stu-id="a7c7c-371">Compose</span></span>|

#### <a name="internetmessageid-string"></a><span data-ttu-id="a7c7c-372">internetMessageId :String</span><span class="sxs-lookup"><span data-stu-id="a7c7c-372">internetMessageId :String</span></span>

<span data-ttu-id="a7c7c-p113">Ruft die Internetnachricht-ID einer E-Mail-Nachricht ab. Nur Lesemodus.</span><span class="sxs-lookup"><span data-stu-id="a7c7c-p113">Gets the Internet message identifier for an email message. Read mode only.</span></span>

##### <a name="type"></a><span data-ttu-id="a7c7c-375">Typ:</span><span class="sxs-lookup"><span data-stu-id="a7c7c-375">Type:</span></span>

*   <span data-ttu-id="a7c7c-376">String</span><span class="sxs-lookup"><span data-stu-id="a7c7c-376">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="a7c7c-377">Anforderungen</span><span class="sxs-lookup"><span data-stu-id="a7c7c-377">Requirements</span></span>

|<span data-ttu-id="a7c7c-378">Anforderung</span><span class="sxs-lookup"><span data-stu-id="a7c7c-378">Requirement</span></span>|<span data-ttu-id="a7c7c-379">Wert</span><span class="sxs-lookup"><span data-stu-id="a7c7c-379">Value</span></span>|
|---|---|
|[<span data-ttu-id="a7c7c-380">Mindestversion des Postfachanforderungssatzes</span><span class="sxs-lookup"><span data-stu-id="a7c7c-380">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="a7c7c-381">1.0</span><span class="sxs-lookup"><span data-stu-id="a7c7c-381">1.0</span></span>|
|[<span data-ttu-id="a7c7c-382">Mindestberechtigungsstufe</span><span class="sxs-lookup"><span data-stu-id="a7c7c-382">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="a7c7c-383">ReadItem</span><span class="sxs-lookup"><span data-stu-id="a7c7c-383">ReadItem</span></span>|
|[<span data-ttu-id="a7c7c-384">Zutreffender Outlook-Modus</span><span class="sxs-lookup"><span data-stu-id="a7c7c-384">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="a7c7c-385">Lesen</span><span class="sxs-lookup"><span data-stu-id="a7c7c-385">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="a7c7c-386">Beispiel</span><span class="sxs-lookup"><span data-stu-id="a7c7c-386">Example</span></span>

```
var internetMessageId = Office.context.mailbox.item.internetMessageId;
```

#### <a name="itemclass-string"></a><span data-ttu-id="a7c7c-387">itemClass :String</span><span class="sxs-lookup"><span data-stu-id="a7c7c-387">itemClass :String</span></span>

<span data-ttu-id="a7c7c-p114">Ruft die Exchange-Webdienste-Elementklasse des ausgewählten Elements ab. Nur Lesemodus.</span><span class="sxs-lookup"><span data-stu-id="a7c7c-p114">Gets the Exchange Web Services item class of the selected item. Read mode only.</span></span>

<span data-ttu-id="a7c7c-p115">Die `itemClass`-Eigenschaft gibt die Nachrichtenklasse des ausgewählten Elements an. Folgende Standardnachrichtenklassen für das Nachrichten- oder Terminelement sind vorhanden:</span><span class="sxs-lookup"><span data-stu-id="a7c7c-p115">The `itemClass` property specifies the message class of the selected item. The following are the default message classes for the message or appointment item.</span></span>

|<span data-ttu-id="a7c7c-392">Typ</span><span class="sxs-lookup"><span data-stu-id="a7c7c-392">Type</span></span>|<span data-ttu-id="a7c7c-393">Beschreibung</span><span class="sxs-lookup"><span data-stu-id="a7c7c-393">Description</span></span>|<span data-ttu-id="a7c7c-394">Elementklasse</span><span class="sxs-lookup"><span data-stu-id="a7c7c-394">item class</span></span>|
|---|---|---|
|<span data-ttu-id="a7c7c-395">Terminelemente</span><span class="sxs-lookup"><span data-stu-id="a7c7c-395">Appointment items</span></span>|<span data-ttu-id="a7c7c-396">Dies sind Kalenderelemente der Elementklasse `IPM.Appointment` oder `IPM.Appointment.Occurence`.</span><span class="sxs-lookup"><span data-stu-id="a7c7c-396">These are calendar items of the item class `IPM.Appointment` or `IPM.Appointment.Occurence`.</span></span>|`IPM.Appointment`<br />`IPM.Appointment.Occurence`|
|<span data-ttu-id="a7c7c-397">Nachrichtenelemente</span><span class="sxs-lookup"><span data-stu-id="a7c7c-397">Message items</span></span>|<span data-ttu-id="a7c7c-398">Diese Elemente umfassen E-Mail-Nachrichten der Standardnachrichtenklasse `IPM.Note` sowie Besprechungsanfragen, -antworten und -absagen, die `IPM.Schedule.Meeting` als Basisnachrichtenklasse verwenden.</span><span class="sxs-lookup"><span data-stu-id="a7c7c-398">These include email messages that have the default message class `IPM.Note`, and meeting requests, responses, and cancellations, that use `IPM.Schedule.Meeting` as the base message class.</span></span>|`IPM.Note`<br />`IPM.Schedule.Meeting.Request`<br />`IPM.Schedule.Meeting.Neg`<br />`IPM.Schedule.Meeting.Pos`<br />`IPM.Schedule.Meeting.Tent`<br />`IPM.Schedule.Meeting.Canceled`|

<span data-ttu-id="a7c7c-399">Sie können benutzerdefinierte Nachrichtenklassen zum Erweitern einer Standardnachrichtenklasse erstellen, z. B. eine benutzerdefinierte Terminnachrichtenklasse wie `IPM.Appointment.Contoso`.</span><span class="sxs-lookup"><span data-stu-id="a7c7c-399">You can create custom message classes that extends a default message class, for example, a custom appointment message class `IPM.Appointment.Contoso`.</span></span>

##### <a name="type"></a><span data-ttu-id="a7c7c-400">Typ:</span><span class="sxs-lookup"><span data-stu-id="a7c7c-400">Type:</span></span>

*   <span data-ttu-id="a7c7c-401">String</span><span class="sxs-lookup"><span data-stu-id="a7c7c-401">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="a7c7c-402">Anforderungen</span><span class="sxs-lookup"><span data-stu-id="a7c7c-402">Requirements</span></span>

|<span data-ttu-id="a7c7c-403">Anforderung</span><span class="sxs-lookup"><span data-stu-id="a7c7c-403">Requirement</span></span>|<span data-ttu-id="a7c7c-404">Wert</span><span class="sxs-lookup"><span data-stu-id="a7c7c-404">Value</span></span>|
|---|---|
|[<span data-ttu-id="a7c7c-405">Mindestversion des Postfachanforderungssatzes</span><span class="sxs-lookup"><span data-stu-id="a7c7c-405">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="a7c7c-406">1.0</span><span class="sxs-lookup"><span data-stu-id="a7c7c-406">1.0</span></span>|
|[<span data-ttu-id="a7c7c-407">Mindestberechtigungsstufe</span><span class="sxs-lookup"><span data-stu-id="a7c7c-407">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="a7c7c-408">ReadItem</span><span class="sxs-lookup"><span data-stu-id="a7c7c-408">ReadItem</span></span>|
|[<span data-ttu-id="a7c7c-409">Zutreffender Outlook-Modus</span><span class="sxs-lookup"><span data-stu-id="a7c7c-409">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="a7c7c-410">Lesen</span><span class="sxs-lookup"><span data-stu-id="a7c7c-410">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="a7c7c-411">Beispiel</span><span class="sxs-lookup"><span data-stu-id="a7c7c-411">Example</span></span>

```
var itemClass = Office.context.mailbox.item.itemClass;
```

#### <a name="nullable-itemid-string"></a><span data-ttu-id="a7c7c-412">(nullable) itemId :String</span><span class="sxs-lookup"><span data-stu-id="a7c7c-412">(nullable) itemId :String</span></span>

<span data-ttu-id="a7c7c-p116">Ruft den Exchange-Webdienste-Elementbezeichner für das aktuelle Element ab. Nur Lesemodus.</span><span class="sxs-lookup"><span data-stu-id="a7c7c-p116">Gets the Exchange Web Services item identifier for the current item. Read mode only.</span></span>

> [!NOTE]
> <span data-ttu-id="a7c7c-415">Der Bezeichner, der von der `itemId`-Eigenschaft zurückgegeben wird, ist der gleiche Bezeichner wie der Exchange-Webdienste-Elementbezeichner.</span><span class="sxs-lookup"><span data-stu-id="a7c7c-415">The identifier returned by the `itemId` property is the same as the Exchange Web Services item identifier.</span></span> <span data-ttu-id="a7c7c-416">Die `itemId` -Eigenschaft ist nicht identisch mit der Outlook-Eintrags-ID oder die ID von der Outlook-REST-API verwendet.</span><span class="sxs-lookup"><span data-stu-id="a7c7c-416">The `itemId` property is not identical to the Outlook Entry ID or the ID used by the Outlook REST API.</span></span> <span data-ttu-id="a7c7c-417">REST-API-Aufrufe dieser Wert verwenden, sollten sie konvertiert werden, bevor [Office.context.mailbox.convertToRestId](office.context.mailbox.md#converttorestiditemid-restversion--string)verwenden.</span><span class="sxs-lookup"><span data-stu-id="a7c7c-417">Before making REST API calls using this value, it should be converted using [Office.context.mailbox.convertToRestId](office.context.mailbox.md#converttorestiditemid-restversion--string).</span></span> <span data-ttu-id="a7c7c-418">Weitere Informationen finden Sie unter [Verwenden der Outlook-REST-APIs aus einem Outlook-add-in](https://docs.microsoft.com/outlook/add-ins/use-rest-api#get-the-item-id).</span><span class="sxs-lookup"><span data-stu-id="a7c7c-418">For more details, see [Use the Outlook REST APIs from an Outlook add-in](https://docs.microsoft.com/outlook/add-ins/use-rest-api#get-the-item-id).</span></span>

<span data-ttu-id="a7c7c-p118">Die Eigenschaft `itemId` ist im Modus „Verfassen“ nicht verfügbar. Wenn ein Elementbezeichner erforderlich ist, kann die [`saveAsync`](#saveasyncoptions-callback)-Methode verwendet werden, um das Element zu speichern. Dadurch wird der Elementbezeichner im [`AsyncResult.value`](/javascript/api/office/office.asyncresult)-Parameter in der Rückruffunktion zurückgegeben.</span><span class="sxs-lookup"><span data-stu-id="a7c7c-p118">The `itemId` property is not available in compose mode. If an item identifier is required, the [`saveAsync`](#saveasyncoptions-callback) method can be used to save the item to the store, which will return the item identifier in the [`AsyncResult.value`](/javascript/api/office/office.asyncresult) parameter in the callback function.</span></span>

##### <a name="type"></a><span data-ttu-id="a7c7c-421">Typ:</span><span class="sxs-lookup"><span data-stu-id="a7c7c-421">Type:</span></span>

*   <span data-ttu-id="a7c7c-422">String</span><span class="sxs-lookup"><span data-stu-id="a7c7c-422">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="a7c7c-423">Anforderungen</span><span class="sxs-lookup"><span data-stu-id="a7c7c-423">Requirements</span></span>

|<span data-ttu-id="a7c7c-424">Anforderung</span><span class="sxs-lookup"><span data-stu-id="a7c7c-424">Requirement</span></span>|<span data-ttu-id="a7c7c-425">Wert</span><span class="sxs-lookup"><span data-stu-id="a7c7c-425">Value</span></span>|
|---|---|
|[<span data-ttu-id="a7c7c-426">Mindestversion des Postfachanforderungssatzes</span><span class="sxs-lookup"><span data-stu-id="a7c7c-426">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="a7c7c-427">1.0</span><span class="sxs-lookup"><span data-stu-id="a7c7c-427">1.0</span></span>|
|[<span data-ttu-id="a7c7c-428">Mindestberechtigungsstufe</span><span class="sxs-lookup"><span data-stu-id="a7c7c-428">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="a7c7c-429">ReadItem</span><span class="sxs-lookup"><span data-stu-id="a7c7c-429">ReadItem</span></span>|
|[<span data-ttu-id="a7c7c-430">Zutreffender Outlook-Modus</span><span class="sxs-lookup"><span data-stu-id="a7c7c-430">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="a7c7c-431">Lesen</span><span class="sxs-lookup"><span data-stu-id="a7c7c-431">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="a7c7c-432">Beispiel</span><span class="sxs-lookup"><span data-stu-id="a7c7c-432">Example</span></span>

<span data-ttu-id="a7c7c-p119">Mit dem folgende Code wird das Vorhandensein eines Elementbezeichners überprüft. Wenn die `itemId`-Eigenschaft `null` oder `undefined` zurückgibt, speichert sie das Element und ruft den Elementbezeichner aus dem asynchronen Ergebnis ab.</span><span class="sxs-lookup"><span data-stu-id="a7c7c-p119">The following code checks for the presence of an item identifier. If the `itemId` property returns `null` or `undefined`, it saves the item to the store and gets the item identifier from the asynchronous result.</span></span>

```
var itemId = Office.context.mailbox.item.itemId;
if (itemId === null || itemId == undefined) {
  Office.context.mailbox.item.saveAsync(function(result){
    itemId = result.value;
  });
}
```

####  <a name="itemtype-officemailboxenumsitemtypejavascriptapioutlookofficemailboxenumsitemtype"></a><span data-ttu-id="a7c7c-435">itemType :[Office.MailboxEnums.ItemType](/javascript/api/outlook/office.mailboxenums.itemtype)</span><span class="sxs-lookup"><span data-stu-id="a7c7c-435">itemType :[Office.MailboxEnums.ItemType](/javascript/api/outlook/office.mailboxenums.itemtype)</span></span>

<span data-ttu-id="a7c7c-436">Ruft den Typ des Elements ab, das eine Instanz darstellt.</span><span class="sxs-lookup"><span data-stu-id="a7c7c-436">Gets the type of item that an instance represents.</span></span>

<span data-ttu-id="a7c7c-437">Die `itemType`-Eigenschaft gibt einen der Werte der `ItemType`-Enumeration zurück, der angibt, ob es sich bei der `item`-Objektinstanz um eine Nachricht oder einen Termin handelt.</span><span class="sxs-lookup"><span data-stu-id="a7c7c-437">The `itemType` property returns one of the `ItemType` enumeration values, indicating whether the `item` object instance is a message or an appointment.</span></span>

##### <a name="type"></a><span data-ttu-id="a7c7c-438">Typ:</span><span class="sxs-lookup"><span data-stu-id="a7c7c-438">Type:</span></span>

*   [<span data-ttu-id="a7c7c-439">Office.MailboxEnums.ItemType</span><span class="sxs-lookup"><span data-stu-id="a7c7c-439">Office.MailboxEnums.ItemType</span></span>](/javascript/api/outlook/office.mailboxenums.itemtype)

##### <a name="requirements"></a><span data-ttu-id="a7c7c-440">Anforderungen</span><span class="sxs-lookup"><span data-stu-id="a7c7c-440">Requirements</span></span>

|<span data-ttu-id="a7c7c-441">Anforderung</span><span class="sxs-lookup"><span data-stu-id="a7c7c-441">Requirement</span></span>|<span data-ttu-id="a7c7c-442">Wert</span><span class="sxs-lookup"><span data-stu-id="a7c7c-442">Value</span></span>|
|---|---|
|[<span data-ttu-id="a7c7c-443">Mindestversion des Postfachanforderungssatzes</span><span class="sxs-lookup"><span data-stu-id="a7c7c-443">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="a7c7c-444">1.0</span><span class="sxs-lookup"><span data-stu-id="a7c7c-444">1.0</span></span>|
|[<span data-ttu-id="a7c7c-445">Mindestberechtigungsstufe</span><span class="sxs-lookup"><span data-stu-id="a7c7c-445">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="a7c7c-446">ReadItem</span><span class="sxs-lookup"><span data-stu-id="a7c7c-446">ReadItem</span></span>|
|[<span data-ttu-id="a7c7c-447">Zutreffender Outlook-Modus</span><span class="sxs-lookup"><span data-stu-id="a7c7c-447">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="a7c7c-448">Verfassen oder Lesen</span><span class="sxs-lookup"><span data-stu-id="a7c7c-448">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="a7c7c-449">Beispiel</span><span class="sxs-lookup"><span data-stu-id="a7c7c-449">Example</span></span>

```
if (Office.context.mailbox.item.itemType == Office.MailboxEnums.ItemType.Message)
  // do something
else
  // do something else
```

####  <a name="location-stringlocationjavascriptapioutlookofficelocation"></a><span data-ttu-id="a7c7c-450">location :String|[Location](/javascript/api/outlook/office.location)</span><span class="sxs-lookup"><span data-stu-id="a7c7c-450">location :String|[Location](/javascript/api/outlook/office.location)</span></span>

<span data-ttu-id="a7c7c-451">Ruft den Ort eines Termins ab bzw. legt ihn fest.</span><span class="sxs-lookup"><span data-stu-id="a7c7c-451">Gets or sets the location of an appointment.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="a7c7c-452">Lesemodus</span><span class="sxs-lookup"><span data-stu-id="a7c7c-452">Read mode</span></span>

<span data-ttu-id="a7c7c-453">Die `location`-Eigenschaft gibt eine Zeichenfolge zurück, die den Ort des Termins enthält.</span><span class="sxs-lookup"><span data-stu-id="a7c7c-453">The `location` property returns a string that contains the location of the appointment.</span></span>

##### <a name="compose-mode"></a><span data-ttu-id="a7c7c-454">Verfassenmodus</span><span class="sxs-lookup"><span data-stu-id="a7c7c-454">Compose mode</span></span>

<span data-ttu-id="a7c7c-455">Die `location`-Eigenschaft gibt ein `Location`-Objekt zurück, das Methoden zum Abrufen und Festlegen des Orts für einen Termin bereitstellt.</span><span class="sxs-lookup"><span data-stu-id="a7c7c-455">The `location` property returns a `Location` object that provides methods that are used to get and set the location of the appointment.</span></span>

##### <a name="type"></a><span data-ttu-id="a7c7c-456">Typ:</span><span class="sxs-lookup"><span data-stu-id="a7c7c-456">Type:</span></span>

*   <span data-ttu-id="a7c7c-457">String | [Location](/javascript/api/outlook/office.location)</span><span class="sxs-lookup"><span data-stu-id="a7c7c-457">String | [Location](/javascript/api/outlook/office.location)</span></span>

##### <a name="requirements"></a><span data-ttu-id="a7c7c-458">Anforderungen</span><span class="sxs-lookup"><span data-stu-id="a7c7c-458">Requirements</span></span>

|<span data-ttu-id="a7c7c-459">Anforderung</span><span class="sxs-lookup"><span data-stu-id="a7c7c-459">Requirement</span></span>|<span data-ttu-id="a7c7c-460">Wert</span><span class="sxs-lookup"><span data-stu-id="a7c7c-460">Value</span></span>|
|---|---|
|[<span data-ttu-id="a7c7c-461">Mindestversion des Postfachanforderungssatzes</span><span class="sxs-lookup"><span data-stu-id="a7c7c-461">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="a7c7c-462">1.0</span><span class="sxs-lookup"><span data-stu-id="a7c7c-462">1.0</span></span>|
|[<span data-ttu-id="a7c7c-463">Mindestberechtigungsstufe</span><span class="sxs-lookup"><span data-stu-id="a7c7c-463">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="a7c7c-464">ReadItem</span><span class="sxs-lookup"><span data-stu-id="a7c7c-464">ReadItem</span></span>|
|[<span data-ttu-id="a7c7c-465">Zutreffender Outlook-Modus</span><span class="sxs-lookup"><span data-stu-id="a7c7c-465">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="a7c7c-466">Verfassen oder Lesen</span><span class="sxs-lookup"><span data-stu-id="a7c7c-466">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="a7c7c-467">Beispiel</span><span class="sxs-lookup"><span data-stu-id="a7c7c-467">Example</span></span>

```
var userContext = { value : 1 };
Office.context.mailbox.item.location.getAsync( { context: userContext}, callback);

function callback(asyncResult) {
  var context = asyncResult.context;
  var location = asyncResult.value;
}
```

#### <a name="normalizedsubject-string"></a><span data-ttu-id="a7c7c-468">normalizedSubject :String</span><span class="sxs-lookup"><span data-stu-id="a7c7c-468">normalizedSubject :String</span></span>

<span data-ttu-id="a7c7c-p120">Ruft den Betreff für ein Element ab, wobei alle Präfixe entfernt werden (einschließlich `RE:` und `FWD:`). Nur Lesemodus.</span><span class="sxs-lookup"><span data-stu-id="a7c7c-p120">Gets the subject of an item, with all prefixes removed (including `RE:` and `FWD:`). Read mode only.</span></span>

<span data-ttu-id="a7c7c-p121">Die normalizedSubject-Eigenschaft ruft den Betreff des Elements mit allen Standardpräfixen (z. B. `RE:` und `FW:`) ab, die von E-Mail-Programmen hinzugefügt werden. Wenn der Betreff des Elements mit vollständigen Präfixen abgerufen werden soll, verwenden Sie die [`subject`](#subject-stringsubjectjavascriptapioutlookofficesubject)-Eigenschaft.</span><span class="sxs-lookup"><span data-stu-id="a7c7c-p121">The normalizedSubject property gets the subject of the item, with any standard prefixes (such as `RE:` and `FW:`) that are added by email programs. To get the subject of the item with the prefixes intact, use the [`subject`](#subject-stringsubjectjavascriptapioutlookofficesubject) property.</span></span>

##### <a name="type"></a><span data-ttu-id="a7c7c-473">Typ:</span><span class="sxs-lookup"><span data-stu-id="a7c7c-473">Type:</span></span>

*   <span data-ttu-id="a7c7c-474">String</span><span class="sxs-lookup"><span data-stu-id="a7c7c-474">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="a7c7c-475">Anforderungen</span><span class="sxs-lookup"><span data-stu-id="a7c7c-475">Requirements</span></span>

|<span data-ttu-id="a7c7c-476">Anforderung</span><span class="sxs-lookup"><span data-stu-id="a7c7c-476">Requirement</span></span>|<span data-ttu-id="a7c7c-477">Wert</span><span class="sxs-lookup"><span data-stu-id="a7c7c-477">Value</span></span>|
|---|---|
|[<span data-ttu-id="a7c7c-478">Mindestversion des Postfachanforderungssatzes</span><span class="sxs-lookup"><span data-stu-id="a7c7c-478">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="a7c7c-479">1.0</span><span class="sxs-lookup"><span data-stu-id="a7c7c-479">1.0</span></span>|
|[<span data-ttu-id="a7c7c-480">Mindestberechtigungsstufe</span><span class="sxs-lookup"><span data-stu-id="a7c7c-480">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="a7c7c-481">ReadItem</span><span class="sxs-lookup"><span data-stu-id="a7c7c-481">ReadItem</span></span>|
|[<span data-ttu-id="a7c7c-482">Zutreffender Outlook-Modus</span><span class="sxs-lookup"><span data-stu-id="a7c7c-482">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="a7c7c-483">Lesen</span><span class="sxs-lookup"><span data-stu-id="a7c7c-483">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="a7c7c-484">Beispiel</span><span class="sxs-lookup"><span data-stu-id="a7c7c-484">Example</span></span>

```
var normalizedSubject = Office.context.mailbox.item.normalizedSubject;
```

####  <a name="notificationmessages-notificationmessagesjavascriptapioutlookofficenotificationmessages"></a><span data-ttu-id="a7c7c-485">notificationMessages :[NotificationMessages](/javascript/api/outlook/office.notificationmessages)</span><span class="sxs-lookup"><span data-stu-id="a7c7c-485">notificationMessages :[NotificationMessages](/javascript/api/outlook/office.notificationmessages)</span></span>

<span data-ttu-id="a7c7c-486">Ruft die Benachrichtigungen für ein Element ab.</span><span class="sxs-lookup"><span data-stu-id="a7c7c-486">Gets the notification messages for an item.</span></span>

##### <a name="type"></a><span data-ttu-id="a7c7c-487">Typ:</span><span class="sxs-lookup"><span data-stu-id="a7c7c-487">Type:</span></span>

*   [<span data-ttu-id="a7c7c-488">NotificationMessages</span><span class="sxs-lookup"><span data-stu-id="a7c7c-488">NotificationMessages</span></span>](/javascript/api/outlook/office.notificationmessages)

##### <a name="requirements"></a><span data-ttu-id="a7c7c-489">Anforderungen</span><span class="sxs-lookup"><span data-stu-id="a7c7c-489">Requirements</span></span>

|<span data-ttu-id="a7c7c-490">Anforderung</span><span class="sxs-lookup"><span data-stu-id="a7c7c-490">Requirement</span></span>|<span data-ttu-id="a7c7c-491">Wert</span><span class="sxs-lookup"><span data-stu-id="a7c7c-491">Value</span></span>|
|---|---|
|[<span data-ttu-id="a7c7c-492">Mindestversion des Postfachanforderungssatzes</span><span class="sxs-lookup"><span data-stu-id="a7c7c-492">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="a7c7c-493">1.3</span><span class="sxs-lookup"><span data-stu-id="a7c7c-493">1.3</span></span>|
|[<span data-ttu-id="a7c7c-494">Mindestberechtigungsstufe</span><span class="sxs-lookup"><span data-stu-id="a7c7c-494">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="a7c7c-495">ReadItem</span><span class="sxs-lookup"><span data-stu-id="a7c7c-495">ReadItem</span></span>|
|[<span data-ttu-id="a7c7c-496">Zutreffender Outlook-Modus</span><span class="sxs-lookup"><span data-stu-id="a7c7c-496">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="a7c7c-497">Verfassen oder Lesen</span><span class="sxs-lookup"><span data-stu-id="a7c7c-497">Compose or read</span></span>|

####  <a name="optionalattendees-arrayemailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsrecipientsjavascriptapioutlookofficerecipients"></a><span data-ttu-id="a7c7c-498">optionalAttendees :Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="a7c7c-498">optionalAttendees :Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook/office.recipients)</span></span>

<span data-ttu-id="a7c7c-499">Ermöglicht den Zugriff auf die optionalen Teilnehmer eines Ereignisses.</span><span class="sxs-lookup"><span data-stu-id="a7c7c-499">Provides access to the optional attendees of an event.</span></span> <span data-ttu-id="a7c7c-500">Der Typ des Objekts und die Zugriffsebene, hängt von den Modus des aktuellen Elements ab.</span><span class="sxs-lookup"><span data-stu-id="a7c7c-500">The type of object and level of access depends on the mode of the current item.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="a7c7c-501">Lesemodus</span><span class="sxs-lookup"><span data-stu-id="a7c7c-501">Read mode</span></span>

<span data-ttu-id="a7c7c-502">Die `optionalAttendees`-Eigenschaft gibt ein Array mit einem `EmailAddressDetails`-Objekt für jeden optionalen Teilnehmer an der Besprechung zurück.</span><span class="sxs-lookup"><span data-stu-id="a7c7c-502">The `optionalAttendees` property returns an array that contains an `EmailAddressDetails` object for each optional attendee to the meeting.</span></span>

##### <a name="compose-mode"></a><span data-ttu-id="a7c7c-503">Verfassenmodus</span><span class="sxs-lookup"><span data-stu-id="a7c7c-503">Compose mode</span></span>

<span data-ttu-id="a7c7c-504">Die `optionalAttendees` -Eigenschaft gibt eine `Recipients` -Objekt, das Methoden zum Abrufen oder Aktualisieren der optionalen Teilnehmer für eine Besprechung bereitstellt.</span><span class="sxs-lookup"><span data-stu-id="a7c7c-504">The `optionalAttendees` property returns a `Recipients` object that provides methods to get or update the optional attendees for a meeting.</span></span>

##### <a name="type"></a><span data-ttu-id="a7c7c-505">Typ:</span><span class="sxs-lookup"><span data-stu-id="a7c7c-505">Type:</span></span>

*   <span data-ttu-id="a7c7c-506">Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="a7c7c-506">Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook/office.recipients)</span></span>

##### <a name="requirements"></a><span data-ttu-id="a7c7c-507">Anforderungen</span><span class="sxs-lookup"><span data-stu-id="a7c7c-507">Requirements</span></span>

|<span data-ttu-id="a7c7c-508">Anforderung</span><span class="sxs-lookup"><span data-stu-id="a7c7c-508">Requirement</span></span>|<span data-ttu-id="a7c7c-509">Wert</span><span class="sxs-lookup"><span data-stu-id="a7c7c-509">Value</span></span>|
|---|---|
|[<span data-ttu-id="a7c7c-510">Mindestversion des Postfachanforderungssatzes</span><span class="sxs-lookup"><span data-stu-id="a7c7c-510">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="a7c7c-511">1.0</span><span class="sxs-lookup"><span data-stu-id="a7c7c-511">1.0</span></span>|
|[<span data-ttu-id="a7c7c-512">Mindestberechtigungsstufe</span><span class="sxs-lookup"><span data-stu-id="a7c7c-512">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="a7c7c-513">ReadItem</span><span class="sxs-lookup"><span data-stu-id="a7c7c-513">ReadItem</span></span>|
|[<span data-ttu-id="a7c7c-514">Zutreffender Outlook-Modus</span><span class="sxs-lookup"><span data-stu-id="a7c7c-514">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="a7c7c-515">Verfassen oder Lesen</span><span class="sxs-lookup"><span data-stu-id="a7c7c-515">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="a7c7c-516">Beispiel</span><span class="sxs-lookup"><span data-stu-id="a7c7c-516">Example</span></span>

```
Office.context.mailbox.item.optionalAttendees.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.optionalAttendees.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.optionalAttendees.getAsync(callback);

function callback(asyncResult) {
  var arrayOfOptionalAttendeesRecipients = asyncResult.value;
}
```

#### <a name="organizer-emailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsorganizerjavascriptapioutlookofficeorganizer"></a><span data-ttu-id="a7c7c-517">Organizer:[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)|[Organizer](/javascript/api/outlook/office.organizer)</span><span class="sxs-lookup"><span data-stu-id="a7c7c-517">organizer :[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)|[Organizer](/javascript/api/outlook/office.organizer)</span></span>

<span data-ttu-id="a7c7c-518">Ruft die e-Mail-Adresse des Dialogfelds Organisieren für eine angegebene Besprechung ab.</span><span class="sxs-lookup"><span data-stu-id="a7c7c-518">Gets the email address of the organizer for a specified meeting.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="a7c7c-519">Lesemodus</span><span class="sxs-lookup"><span data-stu-id="a7c7c-519">Read mode</span></span>

<span data-ttu-id="a7c7c-520">Die `organizer` -Eigenschaft gibt ein [EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails) -Objekt, die den Besprechungsorganisator darstellt.</span><span class="sxs-lookup"><span data-stu-id="a7c7c-520">The `organizer` property returns an [EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails) object that represents the meeting organizer.</span></span>

##### <a name="compose-mode"></a><span data-ttu-id="a7c7c-521">Verfassenmodus</span><span class="sxs-lookup"><span data-stu-id="a7c7c-521">Compose mode</span></span>

<span data-ttu-id="a7c7c-522">Die `organizer` -Eigenschaft gibt ein [Organisator](/javascript/api/outlook/office.organizer) -Objekt, das eine Methode zum Abrufen des Werts der Organisator bereitstellt.</span><span class="sxs-lookup"><span data-stu-id="a7c7c-522">The `organizer` property returns an [Organizer](/javascript/api/outlook/office.organizer) object that provides a method to get the organizer value.</span></span>

##### <a name="type"></a><span data-ttu-id="a7c7c-523">Typ:</span><span class="sxs-lookup"><span data-stu-id="a7c7c-523">Type:</span></span>

*   <span data-ttu-id="a7c7c-524">[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails) | [Organizer](/javascript/api/outlook/office.organizer)</span><span class="sxs-lookup"><span data-stu-id="a7c7c-524">[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails) | [Organizer](/javascript/api/outlook/office.organizer)</span></span>

##### <a name="requirements"></a><span data-ttu-id="a7c7c-525">Anforderungen</span><span class="sxs-lookup"><span data-stu-id="a7c7c-525">Requirements</span></span>

|<span data-ttu-id="a7c7c-526">Anforderung</span><span class="sxs-lookup"><span data-stu-id="a7c7c-526">Requirement</span></span>|||
|---|---|---|
|[<span data-ttu-id="a7c7c-527">Mindestversion des Postfachanforderungssatzes</span><span class="sxs-lookup"><span data-stu-id="a7c7c-527">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="a7c7c-528">1.0</span><span class="sxs-lookup"><span data-stu-id="a7c7c-528">1.0</span></span>|<span data-ttu-id="a7c7c-529">1.7</span><span class="sxs-lookup"><span data-stu-id="a7c7c-529">1.7</span></span>|
|[<span data-ttu-id="a7c7c-530">Mindestberechtigungsstufe</span><span class="sxs-lookup"><span data-stu-id="a7c7c-530">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="a7c7c-531">ReadItem</span><span class="sxs-lookup"><span data-stu-id="a7c7c-531">ReadItem</span></span>|<span data-ttu-id="a7c7c-532">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="a7c7c-532">ReadWriteItem</span></span>|
|[<span data-ttu-id="a7c7c-533">Zutreffender Outlook-Modus</span><span class="sxs-lookup"><span data-stu-id="a7c7c-533">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="a7c7c-534">Lesen</span><span class="sxs-lookup"><span data-stu-id="a7c7c-534">Read</span></span>|<span data-ttu-id="a7c7c-535">Verfassen</span><span class="sxs-lookup"><span data-stu-id="a7c7c-535">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="a7c7c-536">Beispiel</span><span class="sxs-lookup"><span data-stu-id="a7c7c-536">Example</span></span>

```
var organizerName = Office.context.mailbox.item.organizer.displayName;
var organizerAddress = Office.context.mailbox.item.organizer.emailAddress;
```

#### <a name="nullable-recurrence-recurrencejavascriptapioutlookofficerecurrence"></a><span data-ttu-id="a7c7c-537">(Nullwerte zulassen) Serie:[Serie](/javascript/api/outlook/office.recurrence)</span><span class="sxs-lookup"><span data-stu-id="a7c7c-537">(nullable) recurrence :[Recurrence](/javascript/api/outlook/office.recurrence)</span></span>

<span data-ttu-id="a7c7c-538">Ruft ab oder legt das Serienmuster eines Termins.</span><span class="sxs-lookup"><span data-stu-id="a7c7c-538">Gets or sets the recurrence pattern of an appointment.</span></span> <span data-ttu-id="a7c7c-539">Ruft das Serienmuster eine Besprechungsanfrage an.</span><span class="sxs-lookup"><span data-stu-id="a7c7c-539">Gets the recurrence pattern of a meeting request.</span></span> <span data-ttu-id="a7c7c-540">Lesen Sie und verfassen Sie Modi für Terminelemente.</span><span class="sxs-lookup"><span data-stu-id="a7c7c-540">Read and compose modes for appointment items.</span></span> <span data-ttu-id="a7c7c-541">Zur Erfüllung der Anforderung Elemente im Lesemodus.</span><span class="sxs-lookup"><span data-stu-id="a7c7c-541">Read mode for meeting request items.</span></span>

<span data-ttu-id="a7c7c-542">Die `recurrence` -Eigenschaft gibt ein Objekt [Serie](/javascript/api/outlook/office.recurrence) für Termin- oder Besprechungsanfragen zurück, wenn ein Element eine Datenreihe oder eine Instanz einer Serie ist.</span><span class="sxs-lookup"><span data-stu-id="a7c7c-542">The `recurrence` property returns a [recurrence](/javascript/api/outlook/office.recurrence) object for recurring appointments or meetings requests if an item is a series or an instance in a series.</span></span> <span data-ttu-id="a7c7c-543">`null`wird für einzelne Termine und Besprechungsanfragen von einzelnen Terminen zurückgegeben.</span><span class="sxs-lookup"><span data-stu-id="a7c7c-543">`null` is returned for single appointments and meeting requests of single appointments.</span></span> <span data-ttu-id="a7c7c-544">`undefined`wird für Nachrichten zurückgegeben, die nicht Besprechungsanfragen sind.</span><span class="sxs-lookup"><span data-stu-id="a7c7c-544">`undefined` is returned for messages that are not meeting requests.</span></span>

> <span data-ttu-id="a7c7c-545">Hinweis: Besprechungsanfragen verfügen über eine `itemClass` Wert der IPM. Schedule.Meeting.Request.</span><span class="sxs-lookup"><span data-stu-id="a7c7c-545">Note: Meeting requests have an `itemClass` value of IPM.Schedule.Meeting.Request.</span></span>

> <span data-ttu-id="a7c7c-546">Hinweis: Wenn das Objekt Serie ist `null`, darauf hin, dass das Objekt einen einzelnen Termin oder eine Besprechungsanfrage an einen einzelnen Termin und nicht um einen Teil einer Reihe ist.</span><span class="sxs-lookup"><span data-stu-id="a7c7c-546">Note: If the recurrence object is `null`, this indicates that the object is a single appointment or a meeting request of a single appointment and NOT a part of a series.</span></span>

##### <a name="type"></a><span data-ttu-id="a7c7c-547">Typ:</span><span class="sxs-lookup"><span data-stu-id="a7c7c-547">Type:</span></span>

* [<span data-ttu-id="a7c7c-548">Serie</span><span class="sxs-lookup"><span data-stu-id="a7c7c-548">Recurrence</span></span>](/javascript/api/outlook/office.recurrence)

|<span data-ttu-id="a7c7c-549">Anforderung</span><span class="sxs-lookup"><span data-stu-id="a7c7c-549">Requirement</span></span>|<span data-ttu-id="a7c7c-550">Wert</span><span class="sxs-lookup"><span data-stu-id="a7c7c-550">Value</span></span>|
|---|---|
|[<span data-ttu-id="a7c7c-551">Mindestversion des Postfachanforderungssatzes</span><span class="sxs-lookup"><span data-stu-id="a7c7c-551">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="a7c7c-552">1.7</span><span class="sxs-lookup"><span data-stu-id="a7c7c-552">1.7</span></span>|
|[<span data-ttu-id="a7c7c-553">Mindestberechtigungsstufe</span><span class="sxs-lookup"><span data-stu-id="a7c7c-553">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="a7c7c-554">ReadItem</span><span class="sxs-lookup"><span data-stu-id="a7c7c-554">ReadItem</span></span>|
|[<span data-ttu-id="a7c7c-555">Zutreffender Outlook-Modus</span><span class="sxs-lookup"><span data-stu-id="a7c7c-555">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="a7c7c-556">Verfassen oder Lesen</span><span class="sxs-lookup"><span data-stu-id="a7c7c-556">Compose or read</span></span>|

####  <a name="requiredattendees-arrayemailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsrecipientsjavascriptapioutlookofficerecipients"></a><span data-ttu-id="a7c7c-557">requiredAttendees :Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="a7c7c-557">requiredAttendees :Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook/office.recipients)</span></span>

<span data-ttu-id="a7c7c-558">Ermöglicht den Zugriff auf die erforderlichen Teilnehmer eines Ereignisses.</span><span class="sxs-lookup"><span data-stu-id="a7c7c-558">Provides access to the required attendees of an event.</span></span> <span data-ttu-id="a7c7c-559">Der Typ des Objekts und die Zugriffsebene, hängt von den Modus des aktuellen Elements ab.</span><span class="sxs-lookup"><span data-stu-id="a7c7c-559">The type of object and level of access depends on the mode of the current item.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="a7c7c-560">Lesemodus</span><span class="sxs-lookup"><span data-stu-id="a7c7c-560">Read mode</span></span>

<span data-ttu-id="a7c7c-561">Die `requiredAttendees`-Eigenschaft gibt ein Array mit einem `EmailAddressDetails`-Objekt für jeden erforderlichen Teilnehmer an der Besprechung zurück.</span><span class="sxs-lookup"><span data-stu-id="a7c7c-561">The `requiredAttendees` property returns an array that contains an `EmailAddressDetails` object for each required attendee to the meeting.</span></span>

##### <a name="compose-mode"></a><span data-ttu-id="a7c7c-562">Verfassenmodus</span><span class="sxs-lookup"><span data-stu-id="a7c7c-562">Compose mode</span></span>

<span data-ttu-id="a7c7c-563">Die `requiredAttendees` -Eigenschaft gibt eine `Recipients` -Objekt, das Methoden zum Abrufen oder Aktualisieren der erforderlichen Teilnehmer für eine Besprechung bereitstellt.</span><span class="sxs-lookup"><span data-stu-id="a7c7c-563">The `requiredAttendees` property returns a `Recipients` object that provides methods to get or update the required attendees for a meeting.</span></span>

##### <a name="type"></a><span data-ttu-id="a7c7c-564">Typ:</span><span class="sxs-lookup"><span data-stu-id="a7c7c-564">Type:</span></span>

*   <span data-ttu-id="a7c7c-565">Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="a7c7c-565">Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook/office.recipients)</span></span>

##### <a name="requirements"></a><span data-ttu-id="a7c7c-566">Anforderungen</span><span class="sxs-lookup"><span data-stu-id="a7c7c-566">Requirements</span></span>

|<span data-ttu-id="a7c7c-567">Anforderung</span><span class="sxs-lookup"><span data-stu-id="a7c7c-567">Requirement</span></span>|<span data-ttu-id="a7c7c-568">Wert</span><span class="sxs-lookup"><span data-stu-id="a7c7c-568">Value</span></span>|
|---|---|
|[<span data-ttu-id="a7c7c-569">Mindestversion des Postfachanforderungssatzes</span><span class="sxs-lookup"><span data-stu-id="a7c7c-569">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="a7c7c-570">1.0</span><span class="sxs-lookup"><span data-stu-id="a7c7c-570">1.0</span></span>|
|[<span data-ttu-id="a7c7c-571">Mindestberechtigungsstufe</span><span class="sxs-lookup"><span data-stu-id="a7c7c-571">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="a7c7c-572">ReadItem</span><span class="sxs-lookup"><span data-stu-id="a7c7c-572">ReadItem</span></span>|
|[<span data-ttu-id="a7c7c-573">Zutreffender Outlook-Modus</span><span class="sxs-lookup"><span data-stu-id="a7c7c-573">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="a7c7c-574">Verfassen oder Lesen</span><span class="sxs-lookup"><span data-stu-id="a7c7c-574">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="a7c7c-575">Beispiel</span><span class="sxs-lookup"><span data-stu-id="a7c7c-575">Example</span></span>

```
Office.context.mailbox.item.requiredAttendees.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.requiredAttendees.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.requiredAttendees.getAsync(callback);

function callback(asyncResult) {
  var arrayOfRequiredAttendeesRecipients = asyncResult.value;
}
```

#### <a name="sender-emailaddressdetailsjavascriptapioutlookofficeemailaddressdetails"></a><span data-ttu-id="a7c7c-576">sender :[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)</span><span class="sxs-lookup"><span data-stu-id="a7c7c-576">sender :[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)</span></span>

<span data-ttu-id="a7c7c-p126">Ruft die E-Mail-Adresse des Absenders einer E-Mail-Nachricht ab. Nur Lesemodus.</span><span class="sxs-lookup"><span data-stu-id="a7c7c-p126">Gets the email address of the sender of an email message. Read mode only.</span></span>

<span data-ttu-id="a7c7c-p127">Die [`from`](#from-emailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsfromjavascriptapioutlookofficefrom)- und `sender`-Eigenschaften stellen dieselbe Person dar, außer die Nachricht wurde von einem Delegaten gesendet. In diesem Fall stellt die `from`-Eigenschaft die Stellvertretung dar, und die sender-Eigenschaft stellt den Delegaten dar.</span><span class="sxs-lookup"><span data-stu-id="a7c7c-p127">The [`from`](#from-emailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsfromjavascriptapioutlookofficefrom) and `sender` properties represent the same person unless the message is sent by a delegate. In that case, the `from` property represents the delegator, and the sender property represents the delegate.</span></span>

> [!NOTE]
> <span data-ttu-id="a7c7c-581">Die `recipientType` -Eigenschaft des der `EmailAddressDetails` -Objekts der `sender` -Eigenschaft ist `undefined`.</span><span class="sxs-lookup"><span data-stu-id="a7c7c-581">The `recipientType` property of the `EmailAddressDetails` object in the `sender` property is `undefined`.</span></span>

##### <a name="type"></a><span data-ttu-id="a7c7c-582">Typ:</span><span class="sxs-lookup"><span data-stu-id="a7c7c-582">Type:</span></span>

*   [<span data-ttu-id="a7c7c-583">EmailAddressDetails</span><span class="sxs-lookup"><span data-stu-id="a7c7c-583">EmailAddressDetails</span></span>](/javascript/api/outlook/office.emailaddressdetails)

##### <a name="requirements"></a><span data-ttu-id="a7c7c-584">Anforderungen</span><span class="sxs-lookup"><span data-stu-id="a7c7c-584">Requirements</span></span>

|<span data-ttu-id="a7c7c-585">Anforderung</span><span class="sxs-lookup"><span data-stu-id="a7c7c-585">Requirement</span></span>|<span data-ttu-id="a7c7c-586">Wert</span><span class="sxs-lookup"><span data-stu-id="a7c7c-586">Value</span></span>|
|---|---|
|[<span data-ttu-id="a7c7c-587">Mindestversion des Postfachanforderungssatzes</span><span class="sxs-lookup"><span data-stu-id="a7c7c-587">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="a7c7c-588">1.0</span><span class="sxs-lookup"><span data-stu-id="a7c7c-588">1.0</span></span>|
|[<span data-ttu-id="a7c7c-589">Mindestberechtigungsstufe</span><span class="sxs-lookup"><span data-stu-id="a7c7c-589">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="a7c7c-590">ReadItem</span><span class="sxs-lookup"><span data-stu-id="a7c7c-590">ReadItem</span></span>|
|[<span data-ttu-id="a7c7c-591">Zutreffender Outlook-Modus</span><span class="sxs-lookup"><span data-stu-id="a7c7c-591">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="a7c7c-592">Lesen</span><span class="sxs-lookup"><span data-stu-id="a7c7c-592">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="a7c7c-593">Beispiel</span><span class="sxs-lookup"><span data-stu-id="a7c7c-593">Example</span></span>

```
var senderName = Office.context.mailbox.item.sender.displayName;
var senderAddress = Office.context.mailbox.item.sender.emailAddress;
```

#### <a name="nullable-seriesid-string"></a><span data-ttu-id="a7c7c-594">(Nullwerte zulassen) SeriesId: Zeichenfolge</span><span class="sxs-lookup"><span data-stu-id="a7c7c-594">(nullable) seriesId :String</span></span>

<span data-ttu-id="a7c7c-595">Ruft die Id der Datenreihe, zu der eine Instanz gehört.</span><span class="sxs-lookup"><span data-stu-id="a7c7c-595">Gets the id of the series that an instance belongs to.</span></span>

<span data-ttu-id="a7c7c-596">In OWA und Outlook die `seriesId` gibt die Exchange-Webdienste (EWS)-ID des übergeordneten (Reihe)-Elements, das dieses Element gehört.</span><span class="sxs-lookup"><span data-stu-id="a7c7c-596">In OWA and Outlook, the `seriesId` returns the Exchange Web Services (EWS) ID of the parent (series) item that this item belongs to.</span></span> <span data-ttu-id="a7c7c-597">Jedoch in iOS und Android die `seriesId` gibt die REST-ID des übergeordneten Elements zurück.</span><span class="sxs-lookup"><span data-stu-id="a7c7c-597">However, in iOS and Android, the `seriesId` returns the REST ID of the parent item.</span></span>

> [!NOTE]
> <span data-ttu-id="a7c7c-598">Der Bezeichner, der von der `seriesId`-Eigenschaft zurückgegeben wird, ist der gleiche Bezeichner wie der Exchange-Webdienste-Elementbezeichner.</span><span class="sxs-lookup"><span data-stu-id="a7c7c-598">The identifier returned by the `seriesId` property is the same as the Exchange Web Services item identifier.</span></span> <span data-ttu-id="a7c7c-599">Die `seriesId` -Eigenschaft ist nicht identisch mit der Outlook-IDs von der Outlook-REST-API verwendet.</span><span class="sxs-lookup"><span data-stu-id="a7c7c-599">The `seriesId` property is not identical to the Outlook IDs used by the Outlook REST API.</span></span> <span data-ttu-id="a7c7c-600">REST-API-Aufrufe dieser Wert verwenden, sollten sie konvertiert werden, bevor [Office.context.mailbox.convertToRestId](office.context.mailbox.md#converttorestiditemid-restversion--string)verwenden.</span><span class="sxs-lookup"><span data-stu-id="a7c7c-600">Before making REST API calls using this value, it should be converted using [Office.context.mailbox.convertToRestId](office.context.mailbox.md#converttorestiditemid-restversion--string).</span></span> <span data-ttu-id="a7c7c-601">Weitere Informationen finden Sie unter [Verwenden der Outlook-REST-APIs aus einem Outlook-add-in](https://docs.microsoft.com/outlook/add-ins/use-rest-api).</span><span class="sxs-lookup"><span data-stu-id="a7c7c-601">For more details, see [Use the Outlook REST APIs from an Outlook add-in](https://docs.microsoft.com/outlook/add-ins/use-rest-api).</span></span>

<span data-ttu-id="a7c7c-602">Die `seriesId` -Eigenschaft gibt `null` für Elemente, deren nicht das übergeordnete Elemente wie Termine Serienelementen einzelne, oder Besprechungsanfragen und gibt `undefined` für alle anderen Elemente, die nicht Anforderungen erfüllt sind.</span><span class="sxs-lookup"><span data-stu-id="a7c7c-602">The `seriesId` property returns `null` for items that do not have parent items such as single appointments, series items, or meeting requests and returns `undefined` for any other items that are not meeting requests.</span></span>

##### <a name="type"></a><span data-ttu-id="a7c7c-603">Typ:</span><span class="sxs-lookup"><span data-stu-id="a7c7c-603">Type:</span></span>

* <span data-ttu-id="a7c7c-604">String</span><span class="sxs-lookup"><span data-stu-id="a7c7c-604">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="a7c7c-605">Anforderungen</span><span class="sxs-lookup"><span data-stu-id="a7c7c-605">Requirements</span></span>

|<span data-ttu-id="a7c7c-606">Anforderung</span><span class="sxs-lookup"><span data-stu-id="a7c7c-606">Requirement</span></span>|<span data-ttu-id="a7c7c-607">Wert</span><span class="sxs-lookup"><span data-stu-id="a7c7c-607">Value</span></span>|
|---|---|
|[<span data-ttu-id="a7c7c-608">Mindestversion des Postfachanforderungssatzes</span><span class="sxs-lookup"><span data-stu-id="a7c7c-608">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="a7c7c-609">1.7</span><span class="sxs-lookup"><span data-stu-id="a7c7c-609">1.7</span></span>|
|[<span data-ttu-id="a7c7c-610">Mindestberechtigungsstufe</span><span class="sxs-lookup"><span data-stu-id="a7c7c-610">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="a7c7c-611">ReadItem</span><span class="sxs-lookup"><span data-stu-id="a7c7c-611">ReadItem</span></span>|
|[<span data-ttu-id="a7c7c-612">Zutreffender Outlook-Modus</span><span class="sxs-lookup"><span data-stu-id="a7c7c-612">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="a7c7c-613">Verfassen oder Lesen</span><span class="sxs-lookup"><span data-stu-id="a7c7c-613">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="a7c7c-614">Beispiel</span><span class="sxs-lookup"><span data-stu-id="a7c7c-614">Example</span></span>

```
var seriesId = Office.context.mailbox.item.seriesId;
var isSeries = (seriesId == null);
```

####  <a name="start-datetimejavascriptapioutlookofficetime"></a><span data-ttu-id="a7c7c-615">start :Date|[Time](/javascript/api/outlook/office.time)</span><span class="sxs-lookup"><span data-stu-id="a7c7c-615">start :Date|[Time](/javascript/api/outlook/office.time)</span></span>

<span data-ttu-id="a7c7c-616">Ruft Datum und Zeit für den Beginn des Termins ab oder legt Datum und Uhrzeit fest.</span><span class="sxs-lookup"><span data-stu-id="a7c7c-616">Gets or sets the date and time that the appointment is to begin.</span></span>

<span data-ttu-id="a7c7c-p130">Die `start`-Eigenschaft wird als Datums- und Uhrzeitwert in UTC ausgedrückt. Sie können die [`convertToLocalClientTime`](office.context.mailbox.md#converttolocalclienttimetimevalue--localclienttimejavascriptapioutlookofficelocalclienttime)-Methode verwenden, um den Wert in den lokalen Uhrzeit- und Datumswert des Clients umzuwandeln.</span><span class="sxs-lookup"><span data-stu-id="a7c7c-p130">The `start` property is expressed as a Coordinated Universal Time (UTC) date and time value. You can use the [`convertToLocalClientTime`](office.context.mailbox.md#converttolocalclienttimetimevalue--localclienttimejavascriptapioutlookofficelocalclienttime) method to convert the value to the client’s local date and time.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="a7c7c-619">Lesemodus</span><span class="sxs-lookup"><span data-stu-id="a7c7c-619">Read mode</span></span>

<span data-ttu-id="a7c7c-620">Die `start`-Eigenschaft gibt ein `Date`-Objekt zurück.</span><span class="sxs-lookup"><span data-stu-id="a7c7c-620">The `start` property returns a `Date` object.</span></span>

##### <a name="compose-mode"></a><span data-ttu-id="a7c7c-621">Verfassenmodus</span><span class="sxs-lookup"><span data-stu-id="a7c7c-621">Compose mode</span></span>

<span data-ttu-id="a7c7c-622">Die `start`-Eigenschaft gibt ein `Time`-Objekt zurück.</span><span class="sxs-lookup"><span data-stu-id="a7c7c-622">The `start` property returns a `Time` object.</span></span>

<span data-ttu-id="a7c7c-623">Wenn Sie die [`Time.setAsync`](/javascript/api/outlook/office.time#setasync-datetime--options--callback-)-Methode verwenden, um die Startzeit im Verfassenmodus festzulegen, sollten Sie die [`convertToUtcClientTime`](office.context.mailbox.md#converttoutcclienttimeinput--date)-Methode verwenden, um die Ortszeit auf dem Client für den Server in UTC umzuwandeln.</span><span class="sxs-lookup"><span data-stu-id="a7c7c-623">When you use the [`Time.setAsync`](/javascript/api/outlook/office.time#setasync-datetime--options--callback-) method to set the start time, you should use the [`convertToUtcClientTime`](office.context.mailbox.md#converttoutcclienttimeinput--date) method to convert the local time on the client to UTC for the server.</span></span>

##### <a name="type"></a><span data-ttu-id="a7c7c-624">Typ:</span><span class="sxs-lookup"><span data-stu-id="a7c7c-624">Type:</span></span>

*   <span data-ttu-id="a7c7c-625">Date | [Time](/javascript/api/outlook/office.time)</span><span class="sxs-lookup"><span data-stu-id="a7c7c-625">Date | [Time](/javascript/api/outlook/office.time)</span></span>

##### <a name="requirements"></a><span data-ttu-id="a7c7c-626">Anforderungen</span><span class="sxs-lookup"><span data-stu-id="a7c7c-626">Requirements</span></span>

|<span data-ttu-id="a7c7c-627">Anforderung</span><span class="sxs-lookup"><span data-stu-id="a7c7c-627">Requirement</span></span>|<span data-ttu-id="a7c7c-628">Wert</span><span class="sxs-lookup"><span data-stu-id="a7c7c-628">Value</span></span>|
|---|---|
|[<span data-ttu-id="a7c7c-629">Mindestversion des Postfachanforderungssatzes</span><span class="sxs-lookup"><span data-stu-id="a7c7c-629">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="a7c7c-630">1.0</span><span class="sxs-lookup"><span data-stu-id="a7c7c-630">1.0</span></span>|
|[<span data-ttu-id="a7c7c-631">Mindestberechtigungsstufe</span><span class="sxs-lookup"><span data-stu-id="a7c7c-631">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="a7c7c-632">ReadItem</span><span class="sxs-lookup"><span data-stu-id="a7c7c-632">ReadItem</span></span>|
|[<span data-ttu-id="a7c7c-633">Zutreffender Outlook-Modus</span><span class="sxs-lookup"><span data-stu-id="a7c7c-633">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="a7c7c-634">Verfassen oder Lesen</span><span class="sxs-lookup"><span data-stu-id="a7c7c-634">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="a7c7c-635">Beispiel</span><span class="sxs-lookup"><span data-stu-id="a7c7c-635">Example</span></span>

<span data-ttu-id="a7c7c-636">Im folgenden Beispiel wird die Startzeit eines Termins im Verfassenmodus mithilfe der [`setAsync`](/javascript/api/outlook/office.time#setasync-datetime--options--callback-)-Methode des `Time`-Objekts festgelegt.</span><span class="sxs-lookup"><span data-stu-id="a7c7c-636">The following example sets the start time of an appointment in compose mode by using the [`setAsync`](/javascript/api/outlook/office.time#setasync-datetime--options--callback-) method of the `Time` object.</span></span>

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

####  <a name="subject-stringsubjectjavascriptapioutlookofficesubject"></a><span data-ttu-id="a7c7c-637">subject :String|[Subject](/javascript/api/outlook/office.subject)</span><span class="sxs-lookup"><span data-stu-id="a7c7c-637">subject :String|[Subject](/javascript/api/outlook/office.subject)</span></span>

<span data-ttu-id="a7c7c-638">Ruft die Beschreibung ab, die im Betrefffeld eines Elements angezeigt wird, oder legt sie fest.</span><span class="sxs-lookup"><span data-stu-id="a7c7c-638">Gets or sets the description that appears in the subject field of an item.</span></span>

<span data-ttu-id="a7c7c-639">Die `subject`-Eigenschaft ruft den gesamten Betreff des Elements ab oder legt ihn fest – so, wie er vom E-Mail-Server gesendet wird.</span><span class="sxs-lookup"><span data-stu-id="a7c7c-639">The `subject` property gets or sets the entire subject of the item, as sent by the email server.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="a7c7c-640">Lesemodus</span><span class="sxs-lookup"><span data-stu-id="a7c7c-640">Read mode</span></span>

<span data-ttu-id="a7c7c-p131">Die `subject`-Eigenschaft gibt eine Zeichenfolge zurück. Verwenden Sie die [`normalizedSubject`](#normalizedsubject-string)-Eigenschaft, um den Betreff ohne führende Präfixe wie `RE:` und `FW:` abzurufen.</span><span class="sxs-lookup"><span data-stu-id="a7c7c-p131">The `subject` property returns a string. Use the [`normalizedSubject`](#normalizedsubject-string) property to get the subject minus any leading prefixes such as `RE:` and `FW:`.</span></span>

```
var subject = Office.context.mailbox.item.subject;
```

##### <a name="compose-mode"></a><span data-ttu-id="a7c7c-643">Verfassenmodus</span><span class="sxs-lookup"><span data-stu-id="a7c7c-643">Compose mode</span></span>

<span data-ttu-id="a7c7c-644">Die `subject`-Eigenschaft gibt ein `Subject`-Objekt zurück, das Methoden zum Abrufen und Festlegen des Betreffs bereitstellt.</span><span class="sxs-lookup"><span data-stu-id="a7c7c-644">The `subject` property returns a `Subject` object that provides methods to get and set the subject.</span></span>

```
Office.context.mailbox.item.subject.getAsync(callback);

function callback(asyncResult) {
  var subject = asyncResult.value;
}
```

##### <a name="type"></a><span data-ttu-id="a7c7c-645">Typ:</span><span class="sxs-lookup"><span data-stu-id="a7c7c-645">Type:</span></span>

*   <span data-ttu-id="a7c7c-646">String | [Subject](/javascript/api/outlook/office.subject)</span><span class="sxs-lookup"><span data-stu-id="a7c7c-646">String | [Subject](/javascript/api/outlook/office.subject)</span></span>

##### <a name="requirements"></a><span data-ttu-id="a7c7c-647">Anforderungen</span><span class="sxs-lookup"><span data-stu-id="a7c7c-647">Requirements</span></span>

|<span data-ttu-id="a7c7c-648">Anforderung</span><span class="sxs-lookup"><span data-stu-id="a7c7c-648">Requirement</span></span>|<span data-ttu-id="a7c7c-649">Wert</span><span class="sxs-lookup"><span data-stu-id="a7c7c-649">Value</span></span>|
|---|---|
|[<span data-ttu-id="a7c7c-650">Mindestversion des Postfachanforderungssatzes</span><span class="sxs-lookup"><span data-stu-id="a7c7c-650">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="a7c7c-651">1.0</span><span class="sxs-lookup"><span data-stu-id="a7c7c-651">1.0</span></span>|
|[<span data-ttu-id="a7c7c-652">Mindestberechtigungsstufe</span><span class="sxs-lookup"><span data-stu-id="a7c7c-652">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="a7c7c-653">ReadItem</span><span class="sxs-lookup"><span data-stu-id="a7c7c-653">ReadItem</span></span>|
|[<span data-ttu-id="a7c7c-654">Zutreffender Outlook-Modus</span><span class="sxs-lookup"><span data-stu-id="a7c7c-654">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="a7c7c-655">Verfassen oder Lesen</span><span class="sxs-lookup"><span data-stu-id="a7c7c-655">Compose or read</span></span>|

####  <a name="to-arrayemailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsrecipientsjavascriptapioutlookofficerecipients"></a><span data-ttu-id="a7c7c-656">an: Array. <[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)>|[Empfänger](/javascript/api/outlook/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="a7c7c-656">to :Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook/office.recipients)</span></span>

<span data-ttu-id="a7c7c-657">Ermöglicht den Zugriff auf die Empfänger in der Zeile **an** einer Nachricht.</span><span class="sxs-lookup"><span data-stu-id="a7c7c-657">Provides access to the recipients on the **To** line of a message.</span></span> <span data-ttu-id="a7c7c-658">Der Typ des Objekts und die Zugriffsebene, hängt von den Modus des aktuellen Elements ab.</span><span class="sxs-lookup"><span data-stu-id="a7c7c-658">The type of object and level of access depends on the mode of the current item.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="a7c7c-659">Lesemodus</span><span class="sxs-lookup"><span data-stu-id="a7c7c-659">Read mode</span></span>

<span data-ttu-id="a7c7c-p133">Die `to`-Eigenschaft gibt ein Array mit einem `EmailAddressDetails`-Objekt für jeden Empfänger in der**An**-Zeile der Nachricht zurück. Die Auflistung ist auf höchstens 100 Elemente beschränkt.</span><span class="sxs-lookup"><span data-stu-id="a7c7c-p133">The `to` property returns an array that contains an `EmailAddressDetails` object for each recipient listed on the **To** line of the message. The collection is limited to a maximum of 100 members.</span></span>

##### <a name="compose-mode"></a><span data-ttu-id="a7c7c-662">Verfassenmodus</span><span class="sxs-lookup"><span data-stu-id="a7c7c-662">Compose mode</span></span>

<span data-ttu-id="a7c7c-663">Die `to` -Eigenschaft gibt eine `Recipients` -Objekt, das Methoden zum Abrufen oder aktualisieren die Empfänger in der Zeile **an** der Nachricht bereitstellt.</span><span class="sxs-lookup"><span data-stu-id="a7c7c-663">The `to` property returns a `Recipients` object that provides methods to get or update the recipients on the **To** line of the message.</span></span>

##### <a name="type"></a><span data-ttu-id="a7c7c-664">Typ:</span><span class="sxs-lookup"><span data-stu-id="a7c7c-664">Type:</span></span>

*   <span data-ttu-id="a7c7c-665">Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="a7c7c-665">Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook/office.recipients)</span></span>

##### <a name="requirements"></a><span data-ttu-id="a7c7c-666">Anforderungen</span><span class="sxs-lookup"><span data-stu-id="a7c7c-666">Requirements</span></span>

|<span data-ttu-id="a7c7c-667">Anforderung</span><span class="sxs-lookup"><span data-stu-id="a7c7c-667">Requirement</span></span>|<span data-ttu-id="a7c7c-668">Wert</span><span class="sxs-lookup"><span data-stu-id="a7c7c-668">Value</span></span>|
|---|---|
|[<span data-ttu-id="a7c7c-669">Mindestversion des Postfachanforderungssatzes</span><span class="sxs-lookup"><span data-stu-id="a7c7c-669">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="a7c7c-670">1.0</span><span class="sxs-lookup"><span data-stu-id="a7c7c-670">1.0</span></span>|
|[<span data-ttu-id="a7c7c-671">Mindestberechtigungsstufe</span><span class="sxs-lookup"><span data-stu-id="a7c7c-671">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="a7c7c-672">ReadItem</span><span class="sxs-lookup"><span data-stu-id="a7c7c-672">ReadItem</span></span>|
|[<span data-ttu-id="a7c7c-673">Zutreffender Outlook-Modus</span><span class="sxs-lookup"><span data-stu-id="a7c7c-673">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="a7c7c-674">Verfassen oder Lesen</span><span class="sxs-lookup"><span data-stu-id="a7c7c-674">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="a7c7c-675">Beispiel</span><span class="sxs-lookup"><span data-stu-id="a7c7c-675">Example</span></span>

```
Office.context.mailbox.item.to.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.to.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.to.getAsync(callback);

function callback(asyncResult) {
  var arrayOfToRecipients = asyncResult.value;
}
```

### <a name="methods"></a><span data-ttu-id="a7c7c-676">Methoden</span><span class="sxs-lookup"><span data-stu-id="a7c7c-676">Methods</span></span>

####  <a name="addfileattachmentasyncuri-attachmentname-options-callback"></a><span data-ttu-id="a7c7c-677">addFileAttachmentAsync(uri, attachmentName, [options], [callback])</span><span class="sxs-lookup"><span data-stu-id="a7c7c-677">addFileAttachmentAsync(uri, attachmentName, [options], [callback])</span></span>

<span data-ttu-id="a7c7c-678">Fügt eine Datei zu einer Nachricht oder einem Termin als Anlage hinzu.</span><span class="sxs-lookup"><span data-stu-id="a7c7c-678">Adds a file to a message or appointment as an attachment.</span></span>

<span data-ttu-id="a7c7c-679">Die `addFileAttachmentAsync`-Methode lädt die Datei am angegebenen URI hoch und fügt sie an das Element im Verfassenformular an.</span><span class="sxs-lookup"><span data-stu-id="a7c7c-679">The `addFileAttachmentAsync` method uploads the file at the specified URI and attaches it to the item in the compose form.</span></span>

<span data-ttu-id="a7c7c-680">Anschließend können Sie den Bezeichner mit der [`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback)-Methode in der gleichen Sitzung zum Entfernen der Anlage verwenden.</span><span class="sxs-lookup"><span data-stu-id="a7c7c-680">You can subsequently use the identifier with the [`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback) method to remove the attachment in the same session.</span></span>

##### <a name="parameters"></a><span data-ttu-id="a7c7c-681">Parameter:</span><span class="sxs-lookup"><span data-stu-id="a7c7c-681">Parameters:</span></span>
|<span data-ttu-id="a7c7c-682">Name</span><span class="sxs-lookup"><span data-stu-id="a7c7c-682">Name</span></span>|<span data-ttu-id="a7c7c-683">Typ</span><span class="sxs-lookup"><span data-stu-id="a7c7c-683">Type</span></span>|<span data-ttu-id="a7c7c-684">Attribute</span><span class="sxs-lookup"><span data-stu-id="a7c7c-684">Attributes</span></span>|<span data-ttu-id="a7c7c-685">Beschreibung</span><span class="sxs-lookup"><span data-stu-id="a7c7c-685">Description</span></span>|
|---|---|---|---|
|`uri`|<span data-ttu-id="a7c7c-686">Zeichenfolge</span><span class="sxs-lookup"><span data-stu-id="a7c7c-686">String</span></span>||<span data-ttu-id="a7c7c-p134">Der URI, der den Speicherort der an die Nachricht oder den Termin anzuhängenden Datei angibt. Die maximale Länge ist 2048 Zeichen.</span><span class="sxs-lookup"><span data-stu-id="a7c7c-p134">The URI that provides the location of the file to attach to the message or appointment. The maximum length is 2048 characters.</span></span>|
|`attachmentName`|<span data-ttu-id="a7c7c-689">Zeichenfolge</span><span class="sxs-lookup"><span data-stu-id="a7c7c-689">String</span></span>||<span data-ttu-id="a7c7c-p135">Der Name der Anlage, der beim Hochladen der Anlage angezeigt wird. Die maximale Länge ist 255 Zeichen.</span><span class="sxs-lookup"><span data-stu-id="a7c7c-p135">The name of the attachment that is shown while the attachment is uploading. The maximum length is 255 characters.</span></span>|
|`options`|<span data-ttu-id="a7c7c-692">Object</span><span class="sxs-lookup"><span data-stu-id="a7c7c-692">Object</span></span>|<span data-ttu-id="a7c7c-693">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="a7c7c-693">&lt;optional&gt;</span></span>|<span data-ttu-id="a7c7c-694">Ein Objektliteral, das mindestens eine der folgenden Eigenschaften enthält.</span><span class="sxs-lookup"><span data-stu-id="a7c7c-694">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`|<span data-ttu-id="a7c7c-695">Object</span><span class="sxs-lookup"><span data-stu-id="a7c7c-695">Object</span></span>|<span data-ttu-id="a7c7c-696">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="a7c7c-696">&lt;optional&gt;</span></span>|<span data-ttu-id="a7c7c-697">Entwickler können jedes Objekt angeben, auf das sie in der Rückrufmethode zugreifen möchten.</span><span class="sxs-lookup"><span data-stu-id="a7c7c-697">Developers can provide any object they wish to access in the callback method.</span></span>|
|`options.isInline`|<span data-ttu-id="a7c7c-698">Boolean</span><span class="sxs-lookup"><span data-stu-id="a7c7c-698">Boolean</span></span>|<span data-ttu-id="a7c7c-699">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="a7c7c-699">&lt;optional&gt;</span></span>|<span data-ttu-id="a7c7c-700">Wenn `true`: Zeigt an, dass die Anlage inline im Nachrichtentext angezeigt wird und nicht in der Anlagenlisten angezeigt werden soll.</span><span class="sxs-lookup"><span data-stu-id="a7c7c-700">If `true`, indicates that the attachment will be shown inline in the message body, and should not be displayed in the attachment list.</span></span>|
|`callback`|<span data-ttu-id="a7c7c-701">function</span><span class="sxs-lookup"><span data-stu-id="a7c7c-701">function</span></span>|<span data-ttu-id="a7c7c-702">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="a7c7c-702">&lt;optional&gt;</span></span>|<span data-ttu-id="a7c7c-703">Wenn die Methode abgeschlossen wird, wird die Funktion , die im `callback`-Parameter übergeben wird, mit einem einzelnen Parameter (`asyncResult`) aufgerufen, der ein [`AsyncResult`](/javascript/api/office/office.asyncresult)-Objekt darstellt.</span><span class="sxs-lookup"><span data-stu-id="a7c7c-703">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span> <br/><span data-ttu-id="a7c7c-704">Bei Erfolg wird der Anlagenbezeichner in der `asyncResult.value`-Eigenschaft bereitgestellt.</span><span class="sxs-lookup"><span data-stu-id="a7c7c-704">On success, the attachment identifier will be provided in the `asyncResult.value` property.</span></span><br/><span data-ttu-id="a7c7c-705">Wenn beim Hochladen der Anlage ein Fehler auftritt, enthält das `asyncResult`-Objekt ein `Error`-Objekt mit einer Beschreibung des Fehlers.</span><span class="sxs-lookup"><span data-stu-id="a7c7c-705">If uploading the attachment fails, the `asyncResult` object will contain an `Error` object that provides a description of the error.</span></span>|

##### <a name="errors"></a><span data-ttu-id="a7c7c-706">Fehler</span><span class="sxs-lookup"><span data-stu-id="a7c7c-706">Errors</span></span>

|<span data-ttu-id="a7c7c-707">Fehlercode</span><span class="sxs-lookup"><span data-stu-id="a7c7c-707">Error code</span></span>|<span data-ttu-id="a7c7c-708">Beschreibung</span><span class="sxs-lookup"><span data-stu-id="a7c7c-708">Description</span></span>|
|------------|-------------|
|`AttachmentSizeExceeded`|<span data-ttu-id="a7c7c-709">Die Anlage ist zu groß.</span><span class="sxs-lookup"><span data-stu-id="a7c7c-709">The attachment is larger than allowed.</span></span>|
|`FileTypeNotSupported`|<span data-ttu-id="a7c7c-710">Die Anlage enthält eine Erweiterung, die nicht zulässig ist.</span><span class="sxs-lookup"><span data-stu-id="a7c7c-710">The attachment has an extension that is not allowed.</span></span>|
|`NumberOfAttachmentsExceeded`|<span data-ttu-id="a7c7c-711">Die Nachricht oder der Termin enthält zu viele Anlagen.</span><span class="sxs-lookup"><span data-stu-id="a7c7c-711">The message or appointment has too many attachments.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="a7c7c-712">Anforderungen</span><span class="sxs-lookup"><span data-stu-id="a7c7c-712">Requirements</span></span>

|<span data-ttu-id="a7c7c-713">Anforderung</span><span class="sxs-lookup"><span data-stu-id="a7c7c-713">Requirement</span></span>|<span data-ttu-id="a7c7c-714">Wert</span><span class="sxs-lookup"><span data-stu-id="a7c7c-714">Value</span></span>|
|---|---|
|[<span data-ttu-id="a7c7c-715">Mindestversion des Postfachanforderungssatzes</span><span class="sxs-lookup"><span data-stu-id="a7c7c-715">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="a7c7c-716">1.1</span><span class="sxs-lookup"><span data-stu-id="a7c7c-716">1.1</span></span>|
|[<span data-ttu-id="a7c7c-717">Mindestberechtigungsstufe</span><span class="sxs-lookup"><span data-stu-id="a7c7c-717">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="a7c7c-718">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="a7c7c-718">ReadWriteItem</span></span>|
|[<span data-ttu-id="a7c7c-719">Zutreffender Outlook-Modus</span><span class="sxs-lookup"><span data-stu-id="a7c7c-719">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="a7c7c-720">Verfassen</span><span class="sxs-lookup"><span data-stu-id="a7c7c-720">Compose</span></span>|

##### <a name="examples"></a><span data-ttu-id="a7c7c-721">Beispiele</span><span class="sxs-lookup"><span data-stu-id="a7c7c-721">Examples</span></span>

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

<span data-ttu-id="a7c7c-722">Im folgenden Beispiel wird eine Bilddatei als Inlineanlage hinzugefügt, und die Anlage wird im Nachrichtentext referenziert.</span><span class="sxs-lookup"><span data-stu-id="a7c7c-722">The following example adds an image file as an inline attachment and references the attachment in the message body.</span></span>

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

#### <a name="addfileattachmentfrombase64asyncbase64file-attachmentname-options-callback"></a><span data-ttu-id="a7c7c-723">addFileAttachmentFromBase64Async (base64File, AttachmentName, [Optionen], [Rückruf])</span><span class="sxs-lookup"><span data-stu-id="a7c7c-723">addFileAttachmentFromBase64Async(base64File, attachmentName, [options], [callback])</span></span>

<span data-ttu-id="a7c7c-724">Fügt eine Datei aus der base64-Codierung in einer Nachricht oder eines Termins als Anlage.</span><span class="sxs-lookup"><span data-stu-id="a7c7c-724">Adds a file from the base64 encoding to a message or appointment as an attachment.</span></span>

<span data-ttu-id="a7c7c-725">Die `addFileAttachmentFromBase64Async` -Methode die Datei aus der base64-Codierung und fügt diese an das Element in dem Formular zum Verfassen.</span><span class="sxs-lookup"><span data-stu-id="a7c7c-725">The `addFileAttachmentFromBase64Async` method uploads the file from the base64 encoding and attaches it to the item in the compose form.</span></span> <span data-ttu-id="a7c7c-726">Diese Methode gibt die Anlage-ID in der AsyncResult.value-Objekt zurück.</span><span class="sxs-lookup"><span data-stu-id="a7c7c-726">This method returns the attachment identifier in the AsyncResult.value object.</span></span>

<span data-ttu-id="a7c7c-727">Anschließend können Sie den Bezeichner mit der [`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback)-Methode in der gleichen Sitzung zum Entfernen der Anlage verwenden.</span><span class="sxs-lookup"><span data-stu-id="a7c7c-727">You can subsequently use the identifier with the [`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback) method to remove the attachment in the same session.</span></span>

##### <a name="parameters"></a><span data-ttu-id="a7c7c-728">Parameter:</span><span class="sxs-lookup"><span data-stu-id="a7c7c-728">Parameters:</span></span>
|<span data-ttu-id="a7c7c-729">Name</span><span class="sxs-lookup"><span data-stu-id="a7c7c-729">Name</span></span>|<span data-ttu-id="a7c7c-730">Typ</span><span class="sxs-lookup"><span data-stu-id="a7c7c-730">Type</span></span>|<span data-ttu-id="a7c7c-731">Attribute</span><span class="sxs-lookup"><span data-stu-id="a7c7c-731">Attributes</span></span>|<span data-ttu-id="a7c7c-732">Beschreibung</span><span class="sxs-lookup"><span data-stu-id="a7c7c-732">Description</span></span>|
|---|---|---|---|
|`base64File`|<span data-ttu-id="a7c7c-733">Zeichenfolge</span><span class="sxs-lookup"><span data-stu-id="a7c7c-733">String</span></span>||<span data-ttu-id="a7c7c-734">Die base64-codierte Inhalt eines Bilds oder eine Datei, eine e-Mail oder ein Ereignis hinzugefügt werden soll.</span><span class="sxs-lookup"><span data-stu-id="a7c7c-734">The base64 encoded content of an image or file to be added to an email or event.</span></span>|
|`attachmentName`|<span data-ttu-id="a7c7c-735">Zeichenfolge</span><span class="sxs-lookup"><span data-stu-id="a7c7c-735">String</span></span>||<span data-ttu-id="a7c7c-p137">Der Name der Anlage, der beim Hochladen der Anlage angezeigt wird. Die maximale Länge ist 255 Zeichen.</span><span class="sxs-lookup"><span data-stu-id="a7c7c-p137">The name of the attachment that is shown while the attachment is uploading. The maximum length is 255 characters.</span></span>|
|`options`|<span data-ttu-id="a7c7c-738">Object</span><span class="sxs-lookup"><span data-stu-id="a7c7c-738">Object</span></span>|<span data-ttu-id="a7c7c-739">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="a7c7c-739">&lt;optional&gt;</span></span>|<span data-ttu-id="a7c7c-740">Ein Objektliteral, das mindestens eine der folgenden Eigenschaften enthält.</span><span class="sxs-lookup"><span data-stu-id="a7c7c-740">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`|<span data-ttu-id="a7c7c-741">Object</span><span class="sxs-lookup"><span data-stu-id="a7c7c-741">Object</span></span>|<span data-ttu-id="a7c7c-742">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="a7c7c-742">&lt;optional&gt;</span></span>|<span data-ttu-id="a7c7c-743">Entwickler können jedes Objekt angeben, auf das sie in der Rückrufmethode zugreifen möchten.</span><span class="sxs-lookup"><span data-stu-id="a7c7c-743">Developers can provide any object they wish to access in the callback method.</span></span>|
|`options.isInline`|<span data-ttu-id="a7c7c-744">Boolean</span><span class="sxs-lookup"><span data-stu-id="a7c7c-744">Boolean</span></span>|<span data-ttu-id="a7c7c-745">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="a7c7c-745">&lt;optional&gt;</span></span>|<span data-ttu-id="a7c7c-746">Wenn `true`: Zeigt an, dass die Anlage inline im Nachrichtentext angezeigt wird und nicht in der Anlagenlisten angezeigt werden soll.</span><span class="sxs-lookup"><span data-stu-id="a7c7c-746">If `true`, indicates that the attachment will be shown inline in the message body, and should not be displayed in the attachment list.</span></span>|
|`callback`|<span data-ttu-id="a7c7c-747">function</span><span class="sxs-lookup"><span data-stu-id="a7c7c-747">function</span></span>|<span data-ttu-id="a7c7c-748">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="a7c7c-748">&lt;optional&gt;</span></span>|<span data-ttu-id="a7c7c-749">Wenn die Methode abgeschlossen wird, wird die Funktion , die im `callback`-Parameter übergeben wird, mit einem einzelnen Parameter (`asyncResult`) aufgerufen, der ein [`AsyncResult`](/javascript/api/office/office.asyncresult)-Objekt darstellt.</span><span class="sxs-lookup"><span data-stu-id="a7c7c-749">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span> <br/><span data-ttu-id="a7c7c-750">Bei Erfolg wird der Anlagenbezeichner in der `asyncResult.value`-Eigenschaft bereitgestellt.</span><span class="sxs-lookup"><span data-stu-id="a7c7c-750">On success, the attachment identifier will be provided in the `asyncResult.value` property.</span></span><br/><span data-ttu-id="a7c7c-751">Wenn beim Hochladen der Anlage ein Fehler auftritt, enthält das `asyncResult`-Objekt ein `Error`-Objekt mit einer Beschreibung des Fehlers.</span><span class="sxs-lookup"><span data-stu-id="a7c7c-751">If uploading the attachment fails, the `asyncResult` object will contain an `Error` object that provides a description of the error.</span></span>|

##### <a name="errors"></a><span data-ttu-id="a7c7c-752">Fehler</span><span class="sxs-lookup"><span data-stu-id="a7c7c-752">Errors</span></span>

|<span data-ttu-id="a7c7c-753">Fehlercode</span><span class="sxs-lookup"><span data-stu-id="a7c7c-753">Error code</span></span>|<span data-ttu-id="a7c7c-754">Beschreibung</span><span class="sxs-lookup"><span data-stu-id="a7c7c-754">Description</span></span>|
|------------|-------------|
|`AttachmentSizeExceeded`|<span data-ttu-id="a7c7c-755">Die Anlage ist zu groß.</span><span class="sxs-lookup"><span data-stu-id="a7c7c-755">The attachment is larger than allowed.</span></span>|
|`FileTypeNotSupported`|<span data-ttu-id="a7c7c-756">Die Anlage enthält eine Erweiterung, die nicht zulässig ist.</span><span class="sxs-lookup"><span data-stu-id="a7c7c-756">The attachment has an extension that is not allowed.</span></span>|
|`NumberOfAttachmentsExceeded`|<span data-ttu-id="a7c7c-757">Die Nachricht oder der Termin enthält zu viele Anlagen.</span><span class="sxs-lookup"><span data-stu-id="a7c7c-757">The message or appointment has too many attachments.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="a7c7c-758">Anforderungen</span><span class="sxs-lookup"><span data-stu-id="a7c7c-758">Requirements</span></span>

|<span data-ttu-id="a7c7c-759">Anforderung</span><span class="sxs-lookup"><span data-stu-id="a7c7c-759">Requirement</span></span>|<span data-ttu-id="a7c7c-760">Wert</span><span class="sxs-lookup"><span data-stu-id="a7c7c-760">Value</span></span>|
|---|---|
|[<span data-ttu-id="a7c7c-761">Mindestversion des Postfachanforderungssatzes</span><span class="sxs-lookup"><span data-stu-id="a7c7c-761">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="a7c7c-762">Vorschau</span><span class="sxs-lookup"><span data-stu-id="a7c7c-762">Preview</span></span>|
|[<span data-ttu-id="a7c7c-763">Mindestberechtigungsstufe</span><span class="sxs-lookup"><span data-stu-id="a7c7c-763">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="a7c7c-764">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="a7c7c-764">ReadWriteItem</span></span>|
|[<span data-ttu-id="a7c7c-765">Zutreffender Outlook-Modus</span><span class="sxs-lookup"><span data-stu-id="a7c7c-765">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="a7c7c-766">Verfassen</span><span class="sxs-lookup"><span data-stu-id="a7c7c-766">Compose</span></span>|

##### <a name="examples"></a><span data-ttu-id="a7c7c-767">Beispiele</span><span class="sxs-lookup"><span data-stu-id="a7c7c-767">Examples</span></span>

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

####  <a name="addhandlerasynceventtype-handler-options-callback"></a><span data-ttu-id="a7c7c-768">addHandlerAsync(eventType, handler, [options], [callback])</span><span class="sxs-lookup"><span data-stu-id="a7c7c-768">addHandlerAsync(eventType, handler, [options], [callback])</span></span>

<span data-ttu-id="a7c7c-769">Fügt einen Ereignishandler für ein unterstütztes Ereignis hinzu.</span><span class="sxs-lookup"><span data-stu-id="a7c7c-769">Adds an event handler for a supported event.</span></span>

<span data-ttu-id="a7c7c-770">Derzeit sind die unterstützten Ereignistypen `Office.EventType.AppointmentTimeChanged`, `Office.EventType.RecipientsChanged`, und`Office.EventType.RecurrenceChanged`</span><span class="sxs-lookup"><span data-stu-id="a7c7c-770">Currently the supported event types are `Office.EventType.AppointmentTimeChanged`, `Office.EventType.RecipientsChanged`, and `Office.EventType.RecurrenceChanged`</span></span>

##### <a name="parameters"></a><span data-ttu-id="a7c7c-771">Parameter:</span><span class="sxs-lookup"><span data-stu-id="a7c7c-771">Parameters:</span></span>

| <span data-ttu-id="a7c7c-772">Name</span><span class="sxs-lookup"><span data-stu-id="a7c7c-772">Name</span></span> | <span data-ttu-id="a7c7c-773">Typ</span><span class="sxs-lookup"><span data-stu-id="a7c7c-773">Type</span></span> | <span data-ttu-id="a7c7c-774">Attribute</span><span class="sxs-lookup"><span data-stu-id="a7c7c-774">Attributes</span></span> | <span data-ttu-id="a7c7c-775">Beschreibung</span><span class="sxs-lookup"><span data-stu-id="a7c7c-775">Description</span></span> |
|---|---|---|---|
| `eventType` | [<span data-ttu-id="a7c7c-776">Office.EventType</span><span class="sxs-lookup"><span data-stu-id="a7c7c-776">Office.EventType</span></span>](office.md#eventtype-string) || <span data-ttu-id="a7c7c-777">Das Ereignis, das den Handler aufrufen soll</span><span class="sxs-lookup"><span data-stu-id="a7c7c-777">The event that should invoke the handler.</span></span> |
| `handler` | <span data-ttu-id="a7c7c-778">Function</span><span class="sxs-lookup"><span data-stu-id="a7c7c-778">Function</span></span> || <span data-ttu-id="a7c7c-p138">Die Funktion, die das Ereignis behandeln soll. Die Funktion muss einen einzigen Parameter akzeptieren (ein Objektliteral). Die Eigenschaft `type` dieses Parameters entspricht dem Parameter `eventType`, der an `addHandlerAsync` übergeben wird.</span><span class="sxs-lookup"><span data-stu-id="a7c7c-p138">The function to handle the event. The function must accept a single parameter, which is an object literal. The `type` property on the parameter will match the `eventType` parameter passed to `addHandlerAsync`.</span></span> |
| `options` | <span data-ttu-id="a7c7c-782">Objekt</span><span class="sxs-lookup"><span data-stu-id="a7c7c-782">Object</span></span> | <span data-ttu-id="a7c7c-783">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="a7c7c-783">&lt;optional&gt;</span></span> | <span data-ttu-id="a7c7c-784">Ein Objektliteral, das mindestens eine der folgenden Eigenschaften enthält.</span><span class="sxs-lookup"><span data-stu-id="a7c7c-784">An object literal that contains one or more of the following properties.</span></span> |
| `options.asyncContext` | <span data-ttu-id="a7c7c-785">Object</span><span class="sxs-lookup"><span data-stu-id="a7c7c-785">Object</span></span> | <span data-ttu-id="a7c7c-786">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="a7c7c-786">&lt;optional&gt;</span></span> | <span data-ttu-id="a7c7c-787">Entwickler können ein Objekt bereitstellen, auf das sie in der Callbackmethode zugreifen möchten.</span><span class="sxs-lookup"><span data-stu-id="a7c7c-787">Developers can provide any object they wish to access in the callback method.</span></span> |
| `callback` | <span data-ttu-id="a7c7c-788">Funktion</span><span class="sxs-lookup"><span data-stu-id="a7c7c-788">function</span></span>| <span data-ttu-id="a7c7c-789">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="a7c7c-789">&lt;optional&gt;</span></span>|<span data-ttu-id="a7c7c-790">Wenn die Methode abgeschlossen wird, wird die Funktion , die im `callback`-Parameter übergeben wird, mit einem einzelnen Parameter (`asyncResult`) aufgerufen, der ein [`AsyncResult`](/javascript/api/office/office.asyncresult)-Objekt darstellt.</span><span class="sxs-lookup"><span data-stu-id="a7c7c-790">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="a7c7c-791">Anforderungen</span><span class="sxs-lookup"><span data-stu-id="a7c7c-791">Requirements</span></span>

|<span data-ttu-id="a7c7c-792">Anforderung</span><span class="sxs-lookup"><span data-stu-id="a7c7c-792">Requirement</span></span>| <span data-ttu-id="a7c7c-793">Wert</span><span class="sxs-lookup"><span data-stu-id="a7c7c-793">Value</span></span>|
|---|---|
|[<span data-ttu-id="a7c7c-794">Mindestversion des Postfachanforderungssatzes</span><span class="sxs-lookup"><span data-stu-id="a7c7c-794">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="a7c7c-795">1.7</span><span class="sxs-lookup"><span data-stu-id="a7c7c-795">1.7</span></span> |
|[<span data-ttu-id="a7c7c-796">Mindestberechtigungsstufe</span><span class="sxs-lookup"><span data-stu-id="a7c7c-796">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="a7c7c-797">ReadItem</span><span class="sxs-lookup"><span data-stu-id="a7c7c-797">ReadItem</span></span> |
|[<span data-ttu-id="a7c7c-798">Zutreffender Outlook-Modus</span><span class="sxs-lookup"><span data-stu-id="a7c7c-798">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="a7c7c-799">Verfassen oder Lesen</span><span class="sxs-lookup"><span data-stu-id="a7c7c-799">Compose or read</span></span> |

####  <a name="additemattachmentasyncitemid-attachmentname-options-callback"></a><span data-ttu-id="a7c7c-800">addItemAttachmentAsync(itemId, attachmentName, [options], [callback])</span><span class="sxs-lookup"><span data-stu-id="a7c7c-800">addItemAttachmentAsync(itemId, attachmentName, [options], [callback])</span></span>

<span data-ttu-id="a7c7c-801">Fügt der Nachricht oder dem Termin ein Exchange-Objekt, wie z. B. eine Nachricht, als Anhang hinzu.</span><span class="sxs-lookup"><span data-stu-id="a7c7c-801">Adds an Exchange item, such as a message, as an attachment to the message or appointment.</span></span>

<span data-ttu-id="a7c7c-p139">Die `addItemAttachmentAsync`-Methode fügt das Element mit dem angegebenen Exchange-Bezeichner an das Element im Formular zum Verfassen an. Wenn Sie eine Rückrufmethode angeben, wird die Methode mit einem `asyncResult`-Parameter aufgerufen, der entweder den Anlagenbezeichner oder einen Code enthält, der etwaige Fehler angibt, die beim Anfügen des Objekts aufgetreten sind. Sie können ggf. den `options`-Parameter verwenden, um Statusinformationen an die Rückrufmethode zu übergeben.</span><span class="sxs-lookup"><span data-stu-id="a7c7c-p139">The `addItemAttachmentAsync` method attaches the item with the specified Exchange identifier to the item in the compose form. If you specify a callback method, the method is called with one parameter, `asyncResult`, which contains either the attachment identifier or a code that indicates any error that occurred while attaching the item. You can use the `options` parameter to pass state information to the callback method, if needed.</span></span>

<span data-ttu-id="a7c7c-805">Anschließend können Sie den Bezeichner mit der [`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback)-Methode in der gleichen Sitzung zum Entfernen der Anlage verwenden.</span><span class="sxs-lookup"><span data-stu-id="a7c7c-805">You can subsequently use the identifier with the [`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback) method to remove the attachment in the same session.</span></span>

<span data-ttu-id="a7c7c-806">Wenn Ihre Office-Add-Ins in Outlook Web App ausgeführt wird die `addItemAttachmentAsync` -Methode kann Anfügen von Elementen in der Elemente, die Sie bearbeiten; Allerdings Dies wird nicht unterstützt und wird nicht empfohlen.</span><span class="sxs-lookup"><span data-stu-id="a7c7c-806">If your Office Add-in is running in Outlook Web App, the `addItemAttachmentAsync` method can attach items to items other than the item that you are editing; however, this is not supported and is not recommended.</span></span>

##### <a name="parameters"></a><span data-ttu-id="a7c7c-807">Parameter:</span><span class="sxs-lookup"><span data-stu-id="a7c7c-807">Parameters:</span></span>

|<span data-ttu-id="a7c7c-808">Name</span><span class="sxs-lookup"><span data-stu-id="a7c7c-808">Name</span></span>|<span data-ttu-id="a7c7c-809">Typ</span><span class="sxs-lookup"><span data-stu-id="a7c7c-809">Type</span></span>|<span data-ttu-id="a7c7c-810">Attribute</span><span class="sxs-lookup"><span data-stu-id="a7c7c-810">Attributes</span></span>|<span data-ttu-id="a7c7c-811">Beschreibung</span><span class="sxs-lookup"><span data-stu-id="a7c7c-811">Description</span></span>|
|---|---|---|---|
|`itemId`|<span data-ttu-id="a7c7c-812">Zeichenfolge</span><span class="sxs-lookup"><span data-stu-id="a7c7c-812">String</span></span>||<span data-ttu-id="a7c7c-p140">Der Exchange-Bezeichner des Objekts, das angehängt werden soll. Die maximale Länge beträgt 100 Zeichen.</span><span class="sxs-lookup"><span data-stu-id="a7c7c-p140">The Exchange identifier of the item to attach. The maximum length is 100 characters.</span></span>|
|`attachmentName`|<span data-ttu-id="a7c7c-815">Zeichenfolge</span><span class="sxs-lookup"><span data-stu-id="a7c7c-815">String</span></span>||<span data-ttu-id="a7c7c-p141">Der Betreff des Elements, das angehängt werden soll. Die maximale Länge ist 255 Zeichen.</span><span class="sxs-lookup"><span data-stu-id="a7c7c-p141">The sujbect of the item to be attached. The maximum length is 255 characters.</span></span>|
|`options`|<span data-ttu-id="a7c7c-818">Object</span><span class="sxs-lookup"><span data-stu-id="a7c7c-818">Object</span></span>|<span data-ttu-id="a7c7c-819">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="a7c7c-819">&lt;optional&gt;</span></span>|<span data-ttu-id="a7c7c-820">Ein Objektliteral, das mindestens eine der folgenden Eigenschaften enthält.</span><span class="sxs-lookup"><span data-stu-id="a7c7c-820">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`|<span data-ttu-id="a7c7c-821">Object</span><span class="sxs-lookup"><span data-stu-id="a7c7c-821">Object</span></span>|<span data-ttu-id="a7c7c-822">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="a7c7c-822">&lt;optional&gt;</span></span>|<span data-ttu-id="a7c7c-823">Entwickler können ein Objekt bereitstellen, auf das sie in der Callbackmethode zugreifen möchten.</span><span class="sxs-lookup"><span data-stu-id="a7c7c-823">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`|<span data-ttu-id="a7c7c-824">Funktion</span><span class="sxs-lookup"><span data-stu-id="a7c7c-824">function</span></span>|<span data-ttu-id="a7c7c-825">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="a7c7c-825">&lt;optional&gt;</span></span>|<span data-ttu-id="a7c7c-826">Wenn die Methode abgeschlossen wird, wird die Funktion , die im `callback`-Parameter übergeben wird, mit einem einzelnen Parameter (`asyncResult`) aufgerufen, der ein [`AsyncResult`](/javascript/api/office/office.asyncresult)-Objekt darstellt.</span><span class="sxs-lookup"><span data-stu-id="a7c7c-826">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span> <br/><span data-ttu-id="a7c7c-827">Bei Erfolg wird der Anlagenbezeichner in der `asyncResult.value`-Eigenschaft bereitgestellt.</span><span class="sxs-lookup"><span data-stu-id="a7c7c-827">On success, the attachment identifier will be provided in the `asyncResult.value` property.</span></span><br/><span data-ttu-id="a7c7c-828">Wenn beim Hinzufügen der Anlage ein Fehler auftritt, enthält das `asyncResult`-Objekt ein `Error`-Objekt mit einer Beschreibung des Fehlers.</span><span class="sxs-lookup"><span data-stu-id="a7c7c-828">If adding the attachment fails, the `asyncResult` object will contain an `Error` object that provides a description of the error.</span></span>|

##### <a name="errors"></a><span data-ttu-id="a7c7c-829">Fehler</span><span class="sxs-lookup"><span data-stu-id="a7c7c-829">Errors</span></span>

|<span data-ttu-id="a7c7c-830">Fehlercode</span><span class="sxs-lookup"><span data-stu-id="a7c7c-830">Error code</span></span>|<span data-ttu-id="a7c7c-831">Beschreibung</span><span class="sxs-lookup"><span data-stu-id="a7c7c-831">Description</span></span>|
|------------|-------------|
|`NumberOfAttachmentsExceeded`|<span data-ttu-id="a7c7c-832">Die Nachricht oder der Termin enthält zu viele Anlagen.</span><span class="sxs-lookup"><span data-stu-id="a7c7c-832">The message or appointment has too many attachments.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="a7c7c-833">Anforderungen</span><span class="sxs-lookup"><span data-stu-id="a7c7c-833">Requirements</span></span>

|<span data-ttu-id="a7c7c-834">Anforderung</span><span class="sxs-lookup"><span data-stu-id="a7c7c-834">Requirement</span></span>|<span data-ttu-id="a7c7c-835">Wert</span><span class="sxs-lookup"><span data-stu-id="a7c7c-835">Value</span></span>|
|---|---|
|[<span data-ttu-id="a7c7c-836">Mindestversion des Postfachanforderungssatzes</span><span class="sxs-lookup"><span data-stu-id="a7c7c-836">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="a7c7c-837">1.1</span><span class="sxs-lookup"><span data-stu-id="a7c7c-837">1.1</span></span>|
|[<span data-ttu-id="a7c7c-838">Mindestberechtigungsstufe</span><span class="sxs-lookup"><span data-stu-id="a7c7c-838">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="a7c7c-839">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="a7c7c-839">ReadWriteItem</span></span>|
|[<span data-ttu-id="a7c7c-840">Zutreffender Outlook-Modus</span><span class="sxs-lookup"><span data-stu-id="a7c7c-840">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="a7c7c-841">Verfassen</span><span class="sxs-lookup"><span data-stu-id="a7c7c-841">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="a7c7c-842">Beispiel</span><span class="sxs-lookup"><span data-stu-id="a7c7c-842">Example</span></span>

<span data-ttu-id="a7c7c-843">Im folgenden Beispiel wird ein vorhandenes Outlook-Element als Anlage mit dem Namen `My Attachment` hinzugefügt.</span><span class="sxs-lookup"><span data-stu-id="a7c7c-843">The following example adds an existing Outlook item as an attachment with the name `My Attachment`.</span></span>

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

####  <a name="close"></a><span data-ttu-id="a7c7c-844">close()</span><span class="sxs-lookup"><span data-stu-id="a7c7c-844">close()</span></span>

<span data-ttu-id="a7c7c-845">Schließt das aktuelle Element, das gerade verfasst wird.</span><span class="sxs-lookup"><span data-stu-id="a7c7c-845">Closes the current item that is being composed.</span></span>

<span data-ttu-id="a7c7c-p142">Das Verhalten der `close`-Methode hängt vom aktuellen Status des verfassten Elements ab. Wenn das Element über nicht gespeicherte Änderungen verfügt, fordert der Client den Benutzer auf, die Schließenaktion zu speichern, zu verwerfen oder abzubrechen.</span><span class="sxs-lookup"><span data-stu-id="a7c7c-p142">The behavior of the `close` method depends on the current state of the item being composed. If the item has unsaved changes, the client prompts the user to save, discard, or cancel the close action.</span></span>

> [!NOTE]
> <span data-ttu-id="a7c7c-848">In Outlook im Web, wenn das Element einen Termin handelt, und es wurde bereits zuvor gespeichert wurde mit `saveAsync`, der Benutzer wird aufgefordert, speichern, löschen oder Abbrechen, auch wenn keine Änderungen aufgetreten sind, da das Element zuletzt gespeichert.</span><span class="sxs-lookup"><span data-stu-id="a7c7c-848">In Outlook on the web, if the item is an appointment and it has previously been saved using `saveAsync`, the user is prompted to save, discard, or cancel even if no changes have occurred since the item was last saved.</span></span>

<span data-ttu-id="a7c7c-849">Wenn die Nachricht im Outlook-Desktopclient eine Inlineantwort ist, hat die `close`-Methode keine Auswirkung.</span><span class="sxs-lookup"><span data-stu-id="a7c7c-849">In the Outlook desktop client, if the message is an inline reply, the `close` method has no effect.</span></span>

##### <a name="requirements"></a><span data-ttu-id="a7c7c-850">Anforderungen</span><span class="sxs-lookup"><span data-stu-id="a7c7c-850">Requirements</span></span>

|<span data-ttu-id="a7c7c-851">Anforderung</span><span class="sxs-lookup"><span data-stu-id="a7c7c-851">Requirement</span></span>|<span data-ttu-id="a7c7c-852">Wert</span><span class="sxs-lookup"><span data-stu-id="a7c7c-852">Value</span></span>|
|---|---|
|[<span data-ttu-id="a7c7c-853">Mindestversion des Postfachanforderungssatzes</span><span class="sxs-lookup"><span data-stu-id="a7c7c-853">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="a7c7c-854">1.3</span><span class="sxs-lookup"><span data-stu-id="a7c7c-854">1.3</span></span>|
|[<span data-ttu-id="a7c7c-855">Mindestberechtigungsstufe</span><span class="sxs-lookup"><span data-stu-id="a7c7c-855">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="a7c7c-856">Eingeschränkt</span><span class="sxs-lookup"><span data-stu-id="a7c7c-856">Restricted</span></span>|
|[<span data-ttu-id="a7c7c-857">Zutreffender Outlook-Modus</span><span class="sxs-lookup"><span data-stu-id="a7c7c-857">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="a7c7c-858">Verfassen</span><span class="sxs-lookup"><span data-stu-id="a7c7c-858">Compose</span></span>|

#### <a name="displayreplyallformformdata"></a><span data-ttu-id="a7c7c-859">displayReplyAllForm(formData)</span><span class="sxs-lookup"><span data-stu-id="a7c7c-859">displayReplyAllForm(formData)</span></span>

<span data-ttu-id="a7c7c-860">Zeigt ein Antwortformular an, das den Absender und alle Empfänger der ausgewählten Nachricht oder des Organisators und alle Teilnehmer des ausgewählten Termins enthält.</span><span class="sxs-lookup"><span data-stu-id="a7c7c-860">Displays a reply form that includes the sender and all recipients of the selected message or the organizer and all attendees of the selected appointment.</span></span>

> [!NOTE]
> <span data-ttu-id="a7c7c-861">Diese Methode wird in Outlook für iOS oder Outlook für Android nicht unterstützt.</span><span class="sxs-lookup"><span data-stu-id="a7c7c-861">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="a7c7c-862">In Outlook Web App wird das Antwortformular als Popupformular in der Dreispaltenansicht und als Popupformular in der Zwei- oder Einspaltenansicht angezeigt.</span><span class="sxs-lookup"><span data-stu-id="a7c7c-862">In Outlook Web App, the reply form is displayed as a pop-out form in the 3-column view and a pop-up form in the 2- or 1-column view.</span></span>

<span data-ttu-id="a7c7c-863">Wenn einer der Zeichenfolgenparameter seinen Grenzwert überschreitet, löst `displayReplyAllForm` eine Ausnahme aus.</span><span class="sxs-lookup"><span data-stu-id="a7c7c-863">If any of the string parameters exceed their limits, `displayReplyAllForm` throws an exception.</span></span>

<span data-ttu-id="a7c7c-p143">Wenn Anlagen im `formData.attachments`-Parameter angegeben werden, versuchen Outlook und Outlook Web App alle Anlagen herunterzuladen und sie an das Antwortformular anzufügen. Wenn Anlagen nicht hinzugefügt werden können, wird in der Formularbenutzeroberfläche ein Fehler ausgegeben. Wenn dies nicht möglich ist, wird keine Fehlermeldung ausgegeben.</span><span class="sxs-lookup"><span data-stu-id="a7c7c-p143">When attachments are specified in the `formData.attachments` parameter, Outlook and Outlook Web App attempt to download all attachments and attach them to the reply form. If any attachments fail to be added, an error is shown in the form UI. If this isn't possible, then no error message is thrown.</span></span>

##### <a name="parameters"></a><span data-ttu-id="a7c7c-867">Parameter:</span><span class="sxs-lookup"><span data-stu-id="a7c7c-867">Parameters:</span></span>

|<span data-ttu-id="a7c7c-868">Name</span><span class="sxs-lookup"><span data-stu-id="a7c7c-868">Name</span></span>|<span data-ttu-id="a7c7c-869">Typ</span><span class="sxs-lookup"><span data-stu-id="a7c7c-869">Type</span></span>|<span data-ttu-id="a7c7c-870">Attribute</span><span class="sxs-lookup"><span data-stu-id="a7c7c-870">Attributes</span></span>|<span data-ttu-id="a7c7c-871">Beschreibung</span><span class="sxs-lookup"><span data-stu-id="a7c7c-871">Description</span></span>|
|---|---|---|---|
|`formData`|<span data-ttu-id="a7c7c-872">Zeichenfolge &#124; Objekt</span><span class="sxs-lookup"><span data-stu-id="a7c7c-872">String &#124; Object</span></span>||<span data-ttu-id="a7c7c-p144">Eine Zeichenfolge, die Text- und HTML-Code enthält, die den Hauptteil des Antwortformulars darstellen. Die Zeichenfolge ist auf 32 KB beschränkt.</span><span class="sxs-lookup"><span data-stu-id="a7c7c-p144">A string that contains text and HTML and that represents the body of the reply form. The string is limited to 32 KB.</span></span><br/><span data-ttu-id="a7c7c-875">**ODER**</span><span class="sxs-lookup"><span data-stu-id="a7c7c-875">**OR**</span></span><br/><span data-ttu-id="a7c7c-p145">Ein Objekt, das Text- oder Anlagendaten und eine Rückruffunktion enthält. Das Objekt ist wie folgt definiert:</span><span class="sxs-lookup"><span data-stu-id="a7c7c-p145">An object that contains body or attachment data and a callback function. The object is defined as follows.</span></span>|
|`formData.htmlBody`|<span data-ttu-id="a7c7c-878">Zeichenfolge</span><span class="sxs-lookup"><span data-stu-id="a7c7c-878">String</span></span>|<span data-ttu-id="a7c7c-879">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="a7c7c-879">&lt;optional&gt;</span></span>|<span data-ttu-id="a7c7c-p146">Eine Zeichenfolge, die Text- und HTML-Code enthält, die den Hauptteil des Antwortformulars darstellen. Die Zeichenfolge ist auf 32 KB beschränkt.</span><span class="sxs-lookup"><span data-stu-id="a7c7c-p146">A string that contains text and HTML and that represents the body of the reply form. The string is limited to 32 KB.</span></span>
|`formData.attachments`|<span data-ttu-id="a7c7c-882">Array.&lt;Object&gt;</span><span class="sxs-lookup"><span data-stu-id="a7c7c-882">Array.&lt;Object&gt;</span></span>|<span data-ttu-id="a7c7c-883">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="a7c7c-883">&lt;optional&gt;</span></span>|<span data-ttu-id="a7c7c-884">Ein Array mit JSON-Objekten, die Datei- oder Elementanlagen sind.</span><span class="sxs-lookup"><span data-stu-id="a7c7c-884">An array of JSON objects that are either file or item attachments.</span></span>|
|`formData.attachments.type`|<span data-ttu-id="a7c7c-885">Zeichenfolge</span><span class="sxs-lookup"><span data-stu-id="a7c7c-885">String</span></span>||<span data-ttu-id="a7c7c-p147">Gibt den Typ der Anlage an. Muss `file` für eine Dateianlage oder `item` für eine Elementanlage sein.</span><span class="sxs-lookup"><span data-stu-id="a7c7c-p147">Indicates the type of attachment. Must be `file` for a file attachment or `item` for an item attachment.</span></span>|
|`formData.attachments.name`|<span data-ttu-id="a7c7c-888">Zeichenfolge</span><span class="sxs-lookup"><span data-stu-id="a7c7c-888">String</span></span>||<span data-ttu-id="a7c7c-889">Eine Zeichenfolge, die den Namen der Anlage mit bis zu 255 Zeichen enthält.</span><span class="sxs-lookup"><span data-stu-id="a7c7c-889">A string that contains the name of the attachment, up to 255 characters in length.</span></span>|
|`formData.attachments.url`|<span data-ttu-id="a7c7c-890">Zeichenfolge</span><span class="sxs-lookup"><span data-stu-id="a7c7c-890">String</span></span>||<span data-ttu-id="a7c7c-p148">Wird nur verwendet, wenn `type` auf `file` gesetzt ist. Der URI des Speicherorts für die Datei.</span><span class="sxs-lookup"><span data-stu-id="a7c7c-p148">Only used if `type` is set to `file`. The URI of the location for the file.</span></span>|
|`formData.attachments.isInline`|<span data-ttu-id="a7c7c-893">Boolean</span><span class="sxs-lookup"><span data-stu-id="a7c7c-893">Boolean</span></span>||<span data-ttu-id="a7c7c-p149">Wird nur verwendet, wenn `type` auf `file` gesetzt ist. Wenn `true`: Zeigt an, dass die Anlage inline im Nachrichtentext angezeigt wird und nicht in der Anlagenlisten angezeigt werden soll.</span><span class="sxs-lookup"><span data-stu-id="a7c7c-p149">Only used if `type` is set to `file`. If `true`, indicates that the attachment will be shown inline in the message body, and should not be displayed in the attachment list.</span></span>|
|`formData.attachments.itemId`|<span data-ttu-id="a7c7c-896">Zeichenfolge</span><span class="sxs-lookup"><span data-stu-id="a7c7c-896">String</span></span>||<span data-ttu-id="a7c7c-p150">Wird nur verwendet, wenn `type` auf `item` gesetzt ist. Die EWS-Element-ID der Anlage. Dies ist eine Zeichenfolge bis zu 100 Zeichen.</span><span class="sxs-lookup"><span data-stu-id="a7c7c-p150">Only used if `type` is set to `item`. The EWS item id of the attachment. This is a string up to 100 characters.</span></span>|
|`callback`|<span data-ttu-id="a7c7c-900">Funktion</span><span class="sxs-lookup"><span data-stu-id="a7c7c-900">function</span></span>|<span data-ttu-id="a7c7c-901">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="a7c7c-901">&lt;optional&gt;</span></span>|<span data-ttu-id="a7c7c-902">Bei Abschluss der Methode wird die im Parameter `callback` übergebene Funktion mit einem einzigen Parameter aufgerufen: `asyncResult`. Dieser Parameter ist ein Objekt des Typs [AsyncResult](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="a7c7c-902">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [AsyncResult](/javascript/api/office/office.asyncresult) object.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="a7c7c-903">Anforderungen</span><span class="sxs-lookup"><span data-stu-id="a7c7c-903">Requirements</span></span>

|<span data-ttu-id="a7c7c-904">Anforderung</span><span class="sxs-lookup"><span data-stu-id="a7c7c-904">Requirement</span></span>|<span data-ttu-id="a7c7c-905">Wert</span><span class="sxs-lookup"><span data-stu-id="a7c7c-905">Value</span></span>|
|---|---|
|[<span data-ttu-id="a7c7c-906">Mindestversion des Postfachanforderungssatzes</span><span class="sxs-lookup"><span data-stu-id="a7c7c-906">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="a7c7c-907">1.0</span><span class="sxs-lookup"><span data-stu-id="a7c7c-907">1.0</span></span>|
|[<span data-ttu-id="a7c7c-908">Mindestberechtigungsstufe</span><span class="sxs-lookup"><span data-stu-id="a7c7c-908">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="a7c7c-909">ReadItem</span><span class="sxs-lookup"><span data-stu-id="a7c7c-909">ReadItem</span></span>|
|[<span data-ttu-id="a7c7c-910">Zutreffender Outlook-Modus</span><span class="sxs-lookup"><span data-stu-id="a7c7c-910">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="a7c7c-911">Lesen</span><span class="sxs-lookup"><span data-stu-id="a7c7c-911">Read</span></span>|

##### <a name="examples"></a><span data-ttu-id="a7c7c-912">Beispiele</span><span class="sxs-lookup"><span data-stu-id="a7c7c-912">Examples</span></span>

<span data-ttu-id="a7c7c-913">Im folgenden Code wird eine Zeichenfolge an die `displayReplyAllForm`-Funktion übergeben.</span><span class="sxs-lookup"><span data-stu-id="a7c7c-913">The following code passes a string to the `displayReplyAllForm` function.</span></span>

```
Office.context.mailbox.item.displayReplyAllForm('hello there');
Office.context.mailbox.item.displayReplyAllForm('<b>hello there</b>');
```

<span data-ttu-id="a7c7c-914">Antworten mit einem leeren Textkörper.</span><span class="sxs-lookup"><span data-stu-id="a7c7c-914">Reply with an empty body.</span></span>

```
Office.context.mailbox.item.displayReplyAllForm({});
```

<span data-ttu-id="a7c7c-915">Antworten nur einem Textkörper.</span><span class="sxs-lookup"><span data-stu-id="a7c7c-915">Reply with just a body.</span></span>

```
Office.context.mailbox.item.displayReplyAllForm(
{
  'htmlBody' : 'hi'
});
```

<span data-ttu-id="a7c7c-916">Antworten mit einem Textkörper und einer Dateianlage.</span><span class="sxs-lookup"><span data-stu-id="a7c7c-916">Reply with a body and a file attachment.</span></span>

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

<span data-ttu-id="a7c7c-917">Antworten mit einem Textkörper und einer Elementanlage.</span><span class="sxs-lookup"><span data-stu-id="a7c7c-917">Reply with a body and an item attachment.</span></span>

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

<span data-ttu-id="a7c7c-918">Antworten Sie mit einem Textkörper, einer Dateianlage, einer Elementanlage und einem Callback.</span><span class="sxs-lookup"><span data-stu-id="a7c7c-918">Reply with a body, file attachment, item attachment, and a callback.</span></span>

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

#### <a name="displayreplyformformdata"></a><span data-ttu-id="a7c7c-919">displayReplyForm(formData)</span><span class="sxs-lookup"><span data-stu-id="a7c7c-919">displayReplyForm(formData)</span></span>

<span data-ttu-id="a7c7c-920">Zeigt ein Antwortformular an, das nur den Absender der ausgewählten Nachricht oder den Organisator des ausgewählten Termins enthält.</span><span class="sxs-lookup"><span data-stu-id="a7c7c-920">Displays a reply form that includes only the sender of the selected message or the organizer of the selected appointment.</span></span>

> [!NOTE]
> <span data-ttu-id="a7c7c-921">Diese Methode wird in Outlook für iOS oder Outlook für Android nicht unterstützt.</span><span class="sxs-lookup"><span data-stu-id="a7c7c-921">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="a7c7c-922">In Outlook Web App wird das Antwortformular als Popupformular in der Dreispaltenansicht und als Popupformular in der Zwei- oder Einspaltenansicht angezeigt.</span><span class="sxs-lookup"><span data-stu-id="a7c7c-922">In Outlook Web App, the reply form is displayed as a pop-out form in the 3-column view and a pop-up form in the 2- or 1-column view.</span></span>

<span data-ttu-id="a7c7c-923">Wenn einer der Zeichenfolgenparameter seinen Grenzwert überschreitet, löst `displayReplyForm` eine Ausnahme aus.</span><span class="sxs-lookup"><span data-stu-id="a7c7c-923">If any of the string parameters exceed their limits, `displayReplyForm` throws an exception.</span></span>

<span data-ttu-id="a7c7c-p151">Wenn Anlagen im `formData.attachments`-Parameter angegeben werden, versuchen Outlook und Outlook Web App alle Anlagen herunterzuladen und sie an das Antwortformular anzufügen. Wenn Anlagen nicht hinzugefügt werden können, wird in der Formularbenutzeroberfläche ein Fehler ausgegeben. Wenn dies nicht möglich ist, wird keine Fehlermeldung ausgegeben.</span><span class="sxs-lookup"><span data-stu-id="a7c7c-p151">When attachments are specified in the `formData.attachments` parameter, Outlook and Outlook Web App attempt to download all attachments and attach them to the reply form. If any attachments fail to be added, an error is shown in the form UI. If this isn't possible, then no error message is thrown.</span></span>

##### <a name="parameters"></a><span data-ttu-id="a7c7c-927">Parameter:</span><span class="sxs-lookup"><span data-stu-id="a7c7c-927">Parameters:</span></span>

|<span data-ttu-id="a7c7c-928">Name</span><span class="sxs-lookup"><span data-stu-id="a7c7c-928">Name</span></span>|<span data-ttu-id="a7c7c-929">Typ</span><span class="sxs-lookup"><span data-stu-id="a7c7c-929">Type</span></span>|<span data-ttu-id="a7c7c-930">Attribute</span><span class="sxs-lookup"><span data-stu-id="a7c7c-930">Attributes</span></span>|<span data-ttu-id="a7c7c-931">Beschreibung</span><span class="sxs-lookup"><span data-stu-id="a7c7c-931">Description</span></span>|
|---|---|---|---|
|`formData`|<span data-ttu-id="a7c7c-932">Zeichenfolge &#124; Objekt</span><span class="sxs-lookup"><span data-stu-id="a7c7c-932">String &#124; Object</span></span>||<span data-ttu-id="a7c7c-p152">Eine Zeichenfolge, die Text- und HTML-Code enthält, die den Hauptteil des Antwortformulars darstellen. Die Zeichenfolge ist auf 32 KB beschränkt.</span><span class="sxs-lookup"><span data-stu-id="a7c7c-p152">A string that contains text and HTML and that represents the body of the reply form. The string is limited to 32 KB.</span></span><br/><span data-ttu-id="a7c7c-935">**ODER**</span><span class="sxs-lookup"><span data-stu-id="a7c7c-935">**OR**</span></span><br/><span data-ttu-id="a7c7c-p153">Ein Objekt, das Text- oder Anlagendaten und eine Rückruffunktion enthält. Das Objekt ist wie folgt definiert:</span><span class="sxs-lookup"><span data-stu-id="a7c7c-p153">An object that contains body or attachment data and a callback function. The object is defined as follows.</span></span>|
|`formData.htmlBody`|<span data-ttu-id="a7c7c-938">Zeichenfolge</span><span class="sxs-lookup"><span data-stu-id="a7c7c-938">String</span></span>|<span data-ttu-id="a7c7c-939">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="a7c7c-939">&lt;optional&gt;</span></span>|<span data-ttu-id="a7c7c-p154">Eine Zeichenfolge, die Text- und HTML-Code enthält, die den Hauptteil des Antwortformulars darstellen. Die Zeichenfolge ist auf 32 KB beschränkt.</span><span class="sxs-lookup"><span data-stu-id="a7c7c-p154">A string that contains text and HTML and that represents the body of the reply form. The string is limited to 32 KB.</span></span>
|`formData.attachments`|<span data-ttu-id="a7c7c-942">Array.&lt;Object&gt;</span><span class="sxs-lookup"><span data-stu-id="a7c7c-942">Array.&lt;Object&gt;</span></span>|<span data-ttu-id="a7c7c-943">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="a7c7c-943">&lt;optional&gt;</span></span>|<span data-ttu-id="a7c7c-944">Ein Array mit JSON-Objekten, die Datei- oder Elementanlagen sind.</span><span class="sxs-lookup"><span data-stu-id="a7c7c-944">An array of JSON objects that are either file or item attachments.</span></span>|
|`formData.attachments.type`|<span data-ttu-id="a7c7c-945">Zeichenfolge</span><span class="sxs-lookup"><span data-stu-id="a7c7c-945">String</span></span>||<span data-ttu-id="a7c7c-p155">Gibt den Typ der Anlage an. Muss `file` für eine Dateianlage oder `item` für eine Elementanlage sein.</span><span class="sxs-lookup"><span data-stu-id="a7c7c-p155">Indicates the type of attachment. Must be `file` for a file attachment or `item` for an item attachment.</span></span>|
|`formData.attachments.name`|<span data-ttu-id="a7c7c-948">Zeichenfolge</span><span class="sxs-lookup"><span data-stu-id="a7c7c-948">String</span></span>||<span data-ttu-id="a7c7c-949">Eine Zeichenfolge, die den Namen der Anlage mit bis zu 255 Zeichen enthält.</span><span class="sxs-lookup"><span data-stu-id="a7c7c-949">A string that contains the name of the attachment, up to 255 characters in length.</span></span>|
|`formData.attachments.url`|<span data-ttu-id="a7c7c-950">Zeichenfolge</span><span class="sxs-lookup"><span data-stu-id="a7c7c-950">String</span></span>||<span data-ttu-id="a7c7c-p156">Wird nur verwendet, wenn `type` auf `file` gesetzt ist. Der URI des Speicherorts für die Datei.</span><span class="sxs-lookup"><span data-stu-id="a7c7c-p156">Only used if `type` is set to `file`. The URI of the location for the file.</span></span>|
|`formData.attachments.isInline`|<span data-ttu-id="a7c7c-953">Boolean</span><span class="sxs-lookup"><span data-stu-id="a7c7c-953">Boolean</span></span>||<span data-ttu-id="a7c7c-p157">Wird nur verwendet, wenn `type` auf `file` gesetzt ist. Wenn `true`: Zeigt an, dass die Anlage inline im Nachrichtentext angezeigt wird und nicht in der Anlagenlisten angezeigt werden soll.</span><span class="sxs-lookup"><span data-stu-id="a7c7c-p157">Only used if `type` is set to `file`. If `true`, indicates that the attachment will be shown inline in the message body, and should not be displayed in the attachment list.</span></span>|
|`formData.attachments.itemId`|<span data-ttu-id="a7c7c-956">Zeichenfolge</span><span class="sxs-lookup"><span data-stu-id="a7c7c-956">String</span></span>||<span data-ttu-id="a7c7c-p158">Wird nur verwendet, wenn `type` auf `item` gesetzt ist. Die EWS-Element-ID der Anlage. Dies ist eine Zeichenfolge bis zu 100 Zeichen.</span><span class="sxs-lookup"><span data-stu-id="a7c7c-p158">Only used if `type` is set to `item`. The EWS item id of the attachment. This is a string up to 100 characters.</span></span>|
|`callback`|<span data-ttu-id="a7c7c-960">Funktion</span><span class="sxs-lookup"><span data-stu-id="a7c7c-960">function</span></span>|<span data-ttu-id="a7c7c-961">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="a7c7c-961">&lt;optional&gt;</span></span>|<span data-ttu-id="a7c7c-962">Bei Abschluss der Methode wird die im Parameter `callback` übergebene Funktion mit einem einzigen Parameter aufgerufen: `asyncResult`. Dieser Parameter ist ein Objekt des Typs [AsyncResult](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="a7c7c-962">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [AsyncResult](/javascript/api/office/office.asyncresult) object.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="a7c7c-963">Anforderungen</span><span class="sxs-lookup"><span data-stu-id="a7c7c-963">Requirements</span></span>

|<span data-ttu-id="a7c7c-964">Anforderung</span><span class="sxs-lookup"><span data-stu-id="a7c7c-964">Requirement</span></span>|<span data-ttu-id="a7c7c-965">Wert</span><span class="sxs-lookup"><span data-stu-id="a7c7c-965">Value</span></span>|
|---|---|
|[<span data-ttu-id="a7c7c-966">Mindestversion des Postfachanforderungssatzes</span><span class="sxs-lookup"><span data-stu-id="a7c7c-966">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="a7c7c-967">1.0</span><span class="sxs-lookup"><span data-stu-id="a7c7c-967">1.0</span></span>|
|[<span data-ttu-id="a7c7c-968">Mindestberechtigungsstufe</span><span class="sxs-lookup"><span data-stu-id="a7c7c-968">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="a7c7c-969">ReadItem</span><span class="sxs-lookup"><span data-stu-id="a7c7c-969">ReadItem</span></span>|
|[<span data-ttu-id="a7c7c-970">Zutreffender Outlook-Modus</span><span class="sxs-lookup"><span data-stu-id="a7c7c-970">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="a7c7c-971">Lesen</span><span class="sxs-lookup"><span data-stu-id="a7c7c-971">Read</span></span>|

##### <a name="examples"></a><span data-ttu-id="a7c7c-972">Beispiele</span><span class="sxs-lookup"><span data-stu-id="a7c7c-972">Examples</span></span>

<span data-ttu-id="a7c7c-973">Im folgenden Code wird eine Zeichenfolge an die `displayReplyForm`-Funktion übergeben.</span><span class="sxs-lookup"><span data-stu-id="a7c7c-973">The following code passes a string to the `displayReplyForm` function.</span></span>

```
Office.context.mailbox.item.displayReplyForm('hello there');
Office.context.mailbox.item.displayReplyForm('<b>hello there</b>');
```

<span data-ttu-id="a7c7c-974">Antworten mit einem leeren Textkörper.</span><span class="sxs-lookup"><span data-stu-id="a7c7c-974">Reply with an empty body.</span></span>

```
Office.context.mailbox.item.displayReplyForm({});
```

<span data-ttu-id="a7c7c-975">Antworten nur einem Textkörper.</span><span class="sxs-lookup"><span data-stu-id="a7c7c-975">Reply with just a body.</span></span>

```
Office.context.mailbox.item.displayReplyForm(
{
  'htmlBody' : 'hi'
});
```

<span data-ttu-id="a7c7c-976">Antworten mit einem Textkörper und einer Dateianlage.</span><span class="sxs-lookup"><span data-stu-id="a7c7c-976">Reply with a body and a file attachment.</span></span>

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

<span data-ttu-id="a7c7c-977">Antworten mit einem Textkörper und einer Elementanlage.</span><span class="sxs-lookup"><span data-stu-id="a7c7c-977">Reply with a body and an item attachment.</span></span>

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

<span data-ttu-id="a7c7c-978">Antworten Sie mit einem Textkörper, einer Dateianlage, einer Elementanlage und einem Callback.</span><span class="sxs-lookup"><span data-stu-id="a7c7c-978">Reply with a body, file attachment, item attachment, and a callback.</span></span>

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

#### <a name="getentities--entitiesjavascriptapioutlookofficeentities"></a><span data-ttu-id="a7c7c-979">getEntities() → {[Entities](/javascript/api/outlook/office.entities)}</span><span class="sxs-lookup"><span data-stu-id="a7c7c-979">getEntities() → {[Entities](/javascript/api/outlook/office.entities)}</span></span>

<span data-ttu-id="a7c7c-980">Ruft das ausgewählte Element Textkörper gefundenen Entitäten ab.</span><span class="sxs-lookup"><span data-stu-id="a7c7c-980">Gets the entities found in the selected item's body.</span></span>

> [!NOTE]
> <span data-ttu-id="a7c7c-981">Diese Methode wird in Outlook für iOS oder Outlook für Android nicht unterstützt.</span><span class="sxs-lookup"><span data-stu-id="a7c7c-981">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

##### <a name="requirements"></a><span data-ttu-id="a7c7c-982">Anforderungen</span><span class="sxs-lookup"><span data-stu-id="a7c7c-982">Requirements</span></span>

|<span data-ttu-id="a7c7c-983">Anforderung</span><span class="sxs-lookup"><span data-stu-id="a7c7c-983">Requirement</span></span>|<span data-ttu-id="a7c7c-984">Wert</span><span class="sxs-lookup"><span data-stu-id="a7c7c-984">Value</span></span>|
|---|---|
|[<span data-ttu-id="a7c7c-985">Mindestversion des Postfachanforderungssatzes</span><span class="sxs-lookup"><span data-stu-id="a7c7c-985">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="a7c7c-986">1.0</span><span class="sxs-lookup"><span data-stu-id="a7c7c-986">1.0</span></span>|
|[<span data-ttu-id="a7c7c-987">Mindestberechtigungsstufe</span><span class="sxs-lookup"><span data-stu-id="a7c7c-987">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="a7c7c-988">ReadItem</span><span class="sxs-lookup"><span data-stu-id="a7c7c-988">ReadItem</span></span>|
|[<span data-ttu-id="a7c7c-989">Zutreffender Outlook-Modus</span><span class="sxs-lookup"><span data-stu-id="a7c7c-989">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="a7c7c-990">Lesen</span><span class="sxs-lookup"><span data-stu-id="a7c7c-990">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="a7c7c-991">Gibt zurück:</span><span class="sxs-lookup"><span data-stu-id="a7c7c-991">Returns:</span></span>

<span data-ttu-id="a7c7c-992">Typ: [Entitäten](/javascript/api/outlook/office.entities)</span><span class="sxs-lookup"><span data-stu-id="a7c7c-992">Type: [Entities](/javascript/api/outlook/office.entities)</span></span>

##### <a name="example"></a><span data-ttu-id="a7c7c-993">Beispiel</span><span class="sxs-lookup"><span data-stu-id="a7c7c-993">Example</span></span>

<span data-ttu-id="a7c7c-994">Das folgende Beispiel greift auf die Kontakte Entitäten im Textkörper des aktuellen Elements.</span><span class="sxs-lookup"><span data-stu-id="a7c7c-994">The following example accesses the contacts entities in the current item's body.</span></span>

```
var contacts = Office.context.mailbox.item.getEntities().contacts;
```

#### <a name="getentitiesbytypeentitytype--nullable-arraystringcontactjavascriptapioutlookofficecontactmeetingsuggestionjavascriptapioutlookofficemeetingsuggestionphonenumberjavascriptapioutlookofficephonenumbertasksuggestionjavascriptapioutlookofficetasksuggestion"></a><span data-ttu-id="a7c7c-995">getEntitiesByType(entityType) → (nullable) {Array.<(String|[Contact](/javascript/api/outlook/office.contact)|[MeetingSuggestion](/javascript/api/outlook/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook/office.tasksuggestion))>}</span><span class="sxs-lookup"><span data-stu-id="a7c7c-995">getEntitiesByType(entityType) → (nullable) {Array.<(String|[Contact](/javascript/api/outlook/office.contact)|[MeetingSuggestion](/javascript/api/outlook/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook/office.tasksuggestion))>}</span></span>

<span data-ttu-id="a7c7c-996">Ruft ein Array aller Entitäten des angegebenen Entitätstyps im Hauptteil des ausgewählten Elements gefunden.</span><span class="sxs-lookup"><span data-stu-id="a7c7c-996">Gets an array of all the entities of the specified entity type found in the selected item's body.</span></span>

> [!NOTE]
> <span data-ttu-id="a7c7c-997">Diese Methode wird in Outlook für iOS oder Outlook für Android nicht unterstützt.</span><span class="sxs-lookup"><span data-stu-id="a7c7c-997">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

##### <a name="parameters"></a><span data-ttu-id="a7c7c-998">Parameter:</span><span class="sxs-lookup"><span data-stu-id="a7c7c-998">Parameters:</span></span>

|<span data-ttu-id="a7c7c-999">Name</span><span class="sxs-lookup"><span data-stu-id="a7c7c-999">Name</span></span>|<span data-ttu-id="a7c7c-1000">Typ</span><span class="sxs-lookup"><span data-stu-id="a7c7c-1000">Type</span></span>|<span data-ttu-id="a7c7c-1001">Beschreibung</span><span class="sxs-lookup"><span data-stu-id="a7c7c-1001">Description</span></span>|
|---|---|---|
|`entityType`|[<span data-ttu-id="a7c7c-1002">Office.MailboxEnums.EntityType</span><span class="sxs-lookup"><span data-stu-id="a7c7c-1002">Office.MailboxEnums.EntityType</span></span>](/javascript/api/outlook/office.mailboxenums.entitytype)|<span data-ttu-id="a7c7c-1003">Einer der Werte der EntityType-Enumeration.</span><span class="sxs-lookup"><span data-stu-id="a7c7c-1003">One of the EntityType enumeration values.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="a7c7c-1004">Anforderungen</span><span class="sxs-lookup"><span data-stu-id="a7c7c-1004">Requirements</span></span>

|<span data-ttu-id="a7c7c-1005">Anforderung</span><span class="sxs-lookup"><span data-stu-id="a7c7c-1005">Requirement</span></span>|<span data-ttu-id="a7c7c-1006">Wert</span><span class="sxs-lookup"><span data-stu-id="a7c7c-1006">Value</span></span>|
|---|---|
|[<span data-ttu-id="a7c7c-1007">Mindestversion des Postfachanforderungssatzes</span><span class="sxs-lookup"><span data-stu-id="a7c7c-1007">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="a7c7c-1008">1.0</span><span class="sxs-lookup"><span data-stu-id="a7c7c-1008">1.0</span></span>|
|[<span data-ttu-id="a7c7c-1009">Mindestberechtigungsstufe</span><span class="sxs-lookup"><span data-stu-id="a7c7c-1009">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="a7c7c-1010">Eingeschränkt</span><span class="sxs-lookup"><span data-stu-id="a7c7c-1010">Restricted</span></span>|
|[<span data-ttu-id="a7c7c-1011">Zutreffender Outlook-Modus</span><span class="sxs-lookup"><span data-stu-id="a7c7c-1011">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="a7c7c-1012">Lesen</span><span class="sxs-lookup"><span data-stu-id="a7c7c-1012">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="a7c7c-1013">Gibt zurück:</span><span class="sxs-lookup"><span data-stu-id="a7c7c-1013">Returns:</span></span>

<span data-ttu-id="a7c7c-1014">Wenn der in `entityType` übergebene Wert kein gültiges Element der `EntityType`-Enumeration ist, gibt die Methode null zurück.</span><span class="sxs-lookup"><span data-stu-id="a7c7c-1014">If the value passed in `entityType` is not a valid member of the `EntityType` enumeration, the method returns null.</span></span> <span data-ttu-id="a7c7c-1015">Wenn keine Entitäten des angegebenen Typs im Textkörper des Elements vorhanden sind, gibt die Methode ein leeres Array zurück.</span><span class="sxs-lookup"><span data-stu-id="a7c7c-1015">If no entities of the specified type are present in the item's body, the method returns an empty array.</span></span> <span data-ttu-id="a7c7c-1016">Andernfalls hängt der Typ der Objekte im zurückgegebenen Array vom Typ der Entität ab, die im `entityType`-Parameter angefordert wurde.</span><span class="sxs-lookup"><span data-stu-id="a7c7c-1016">Otherwise, the type of the objects in the returned array depends on the type of entity requested in the `entityType` parameter.</span></span>

<span data-ttu-id="a7c7c-1017">Während Sie die minimale Berechtigungsstufe für diese Methode **Restricted** ist, erfordern einige Entitätstypen **ReadItem** für den Zugriff, wie in der folgenden Tabelle angegeben wird.</span><span class="sxs-lookup"><span data-stu-id="a7c7c-1017">While the minimum permission level to use this method is **Restricted**, some entity types require **ReadItem** to access, as specified in the following table.</span></span>

|<span data-ttu-id="a7c7c-1018">Wert von `entityType`</span><span class="sxs-lookup"><span data-stu-id="a7c7c-1018">Value of `entityType`</span></span>|<span data-ttu-id="a7c7c-1019">Typ der Objekte im zurückgegebenen Array</span><span class="sxs-lookup"><span data-stu-id="a7c7c-1019">Type of objects in returned array</span></span>|<span data-ttu-id="a7c7c-1020">Erforderliche Berechtigungsstufe</span><span class="sxs-lookup"><span data-stu-id="a7c7c-1020">Required Permission Level</span></span>|
|---|---|---|
|`Address`|<span data-ttu-id="a7c7c-1021">Zeichenfolge</span><span class="sxs-lookup"><span data-stu-id="a7c7c-1021">String</span></span>|<span data-ttu-id="a7c7c-1022">**Restricted**</span><span class="sxs-lookup"><span data-stu-id="a7c7c-1022">**Restricted**</span></span>|
|`Contact`|<span data-ttu-id="a7c7c-1023">Contact</span><span class="sxs-lookup"><span data-stu-id="a7c7c-1023">Contact</span></span>|<span data-ttu-id="a7c7c-1024">**ReadItem**</span><span class="sxs-lookup"><span data-stu-id="a7c7c-1024">**ReadItem**</span></span>|
|`EmailAddress`|<span data-ttu-id="a7c7c-1025">Zeichenfolge</span><span class="sxs-lookup"><span data-stu-id="a7c7c-1025">String</span></span>|<span data-ttu-id="a7c7c-1026">**ReadItem**</span><span class="sxs-lookup"><span data-stu-id="a7c7c-1026">**ReadItem**</span></span>|
|`MeetingSuggestion`|<span data-ttu-id="a7c7c-1027">MeetingSuggestion</span><span class="sxs-lookup"><span data-stu-id="a7c7c-1027">MeetingSuggestion</span></span>|<span data-ttu-id="a7c7c-1028">**ReadItem**</span><span class="sxs-lookup"><span data-stu-id="a7c7c-1028">**ReadItem**</span></span>|
|`PhoneNumber`|<span data-ttu-id="a7c7c-1029">PhoneNumber</span><span class="sxs-lookup"><span data-stu-id="a7c7c-1029">PhoneNumber</span></span>|<span data-ttu-id="a7c7c-1030">**Restricted**</span><span class="sxs-lookup"><span data-stu-id="a7c7c-1030">**Restricted**</span></span>|
|`TaskSuggestion`|<span data-ttu-id="a7c7c-1031">TaskSuggestion</span><span class="sxs-lookup"><span data-stu-id="a7c7c-1031">TaskSuggestion</span></span>|<span data-ttu-id="a7c7c-1032">**ReadItem**</span><span class="sxs-lookup"><span data-stu-id="a7c7c-1032">**ReadItem**</span></span>|
|`URL`|<span data-ttu-id="a7c7c-1033">Zeichenfolge</span><span class="sxs-lookup"><span data-stu-id="a7c7c-1033">String</span></span>|<span data-ttu-id="a7c7c-1034">**Restricted**</span><span class="sxs-lookup"><span data-stu-id="a7c7c-1034">**Restricted**</span></span>|

<span data-ttu-id="a7c7c-1035">Typ: Array.<(String|[Contact](/javascript/api/outlook/office.contact)|[MeetingSuggestion](/javascript/api/outlook/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook/office.tasksuggestion))></span><span class="sxs-lookup"><span data-stu-id="a7c7c-1035">Type: Array.<(String|[Contact](/javascript/api/outlook/office.contact)|[MeetingSuggestion](/javascript/api/outlook/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook/office.tasksuggestion))></span></span>

##### <a name="example"></a><span data-ttu-id="a7c7c-1036">Beispiel</span><span class="sxs-lookup"><span data-stu-id="a7c7c-1036">Example</span></span>

<span data-ttu-id="a7c7c-1037">Im folgenden Beispiel wird veranschaulicht, wie auf ein Array von Zeichenfolgen, die Postadressen im Textkörper des aktuellen Elements darstellen.</span><span class="sxs-lookup"><span data-stu-id="a7c7c-1037">The following example shows how to access an array of strings that represent postal addresses in the current item's body.</span></span>

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

#### <a name="getfilteredentitiesbynamename--nullable-arraystringcontactjavascriptapioutlookofficecontactmeetingsuggestionjavascriptapioutlookofficemeetingsuggestionphonenumberjavascriptapioutlookofficephonenumbertasksuggestionjavascriptapioutlookofficetasksuggestion"></a><span data-ttu-id="a7c7c-1038">getFilteredEntitiesByName(name) → (nullable) {Array.<(String|[Contact](/javascript/api/outlook/office.contact)|[MeetingSuggestion](/javascript/api/outlook/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook/office.tasksuggestion))>}</span><span class="sxs-lookup"><span data-stu-id="a7c7c-1038">getFilteredEntitiesByName(name) → (nullable) {Array.<(String|[Contact](/javascript/api/outlook/office.contact)|[MeetingSuggestion](/javascript/api/outlook/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook/office.tasksuggestion))>}</span></span>

<span data-ttu-id="a7c7c-1039">Gibt bekannte Entitäten im ausgewählten Element zurück, die den in der XML-Manifestdatei definierten benannten Filter übergeben.</span><span class="sxs-lookup"><span data-stu-id="a7c7c-1039">Returns well-known entities in the selected item that pass the named filter defined in the manifest XML file.</span></span>

> [!NOTE]
> <span data-ttu-id="a7c7c-1040">Diese Methode wird in Outlook für iOS oder Outlook für Android nicht unterstützt.</span><span class="sxs-lookup"><span data-stu-id="a7c7c-1040">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="a7c7c-1041">Die `getFilteredEntitiesByName`-Methode gibt die Entitäten zurück, die dem im [ItemHasKnownEntity](/javascript/office/manifest/rule#itemhasknownentity-rule)-Regelelement der XML-Manifestdatei definierten regulären Ausdruck mit dem angegebenen `FilterName`-Elementwert entsprechen.</span><span class="sxs-lookup"><span data-stu-id="a7c7c-1041">The `getFilteredEntitiesByName` method returns the entities that match the regular expression defined in the [ItemHasKnownEntity](/javascript/office/manifest/rule#itemhasknownentity-rule) rule element in the manifest XML file with the specified `FilterName` element value.</span></span>

##### <a name="parameters"></a><span data-ttu-id="a7c7c-1042">Parameter:</span><span class="sxs-lookup"><span data-stu-id="a7c7c-1042">Parameters:</span></span>

|<span data-ttu-id="a7c7c-1043">Name</span><span class="sxs-lookup"><span data-stu-id="a7c7c-1043">Name</span></span>|<span data-ttu-id="a7c7c-1044">Typ</span><span class="sxs-lookup"><span data-stu-id="a7c7c-1044">Type</span></span>|<span data-ttu-id="a7c7c-1045">Beschreibung</span><span class="sxs-lookup"><span data-stu-id="a7c7c-1045">Description</span></span>|
|---|---|---|
|`name`|<span data-ttu-id="a7c7c-1046">Zeichenfolge</span><span class="sxs-lookup"><span data-stu-id="a7c7c-1046">String</span></span>|<span data-ttu-id="a7c7c-1047">Der Name des `ItemHasKnownEntity`-Regelelements, das den entsprechenden Filter definiert.</span><span class="sxs-lookup"><span data-stu-id="a7c7c-1047">The name of the `ItemHasKnownEntity` rule element that defines the filter to match.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="a7c7c-1048">Anforderungen</span><span class="sxs-lookup"><span data-stu-id="a7c7c-1048">Requirements</span></span>

|<span data-ttu-id="a7c7c-1049">Anforderung</span><span class="sxs-lookup"><span data-stu-id="a7c7c-1049">Requirement</span></span>|<span data-ttu-id="a7c7c-1050">Wert</span><span class="sxs-lookup"><span data-stu-id="a7c7c-1050">Value</span></span>|
|---|---|
|[<span data-ttu-id="a7c7c-1051">Mindestversion des Postfachanforderungssatzes</span><span class="sxs-lookup"><span data-stu-id="a7c7c-1051">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="a7c7c-1052">1.0</span><span class="sxs-lookup"><span data-stu-id="a7c7c-1052">1.0</span></span>|
|[<span data-ttu-id="a7c7c-1053">Mindestberechtigungsstufe</span><span class="sxs-lookup"><span data-stu-id="a7c7c-1053">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="a7c7c-1054">ReadItem</span><span class="sxs-lookup"><span data-stu-id="a7c7c-1054">ReadItem</span></span>|
|[<span data-ttu-id="a7c7c-1055">Zutreffender Outlook-Modus</span><span class="sxs-lookup"><span data-stu-id="a7c7c-1055">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="a7c7c-1056">Lesen</span><span class="sxs-lookup"><span data-stu-id="a7c7c-1056">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="a7c7c-1057">Gibt zurück:</span><span class="sxs-lookup"><span data-stu-id="a7c7c-1057">Returns:</span></span>

<span data-ttu-id="a7c7c-p160">Befindet sich kein `ItemHasKnownEntity`-Element im Manifest mit einem `FilterName`-Elementwert, der dem `name`-Parameter entspricht, gibt die Methode `null` zurück. Wenn der `name`-Parameter einem `ItemHasKnownEntity`-Element im Manifest entspricht, es aber keine entsprechenden Entitäten im aktuellen Element gibt, gibt die Methode ein leeres Array zurück.</span><span class="sxs-lookup"><span data-stu-id="a7c7c-p160">If there is no `ItemHasKnownEntity` element in the manifest with a `FilterName` element value that matches the `name` parameter, the method returns `null`. If the `name` parameter does match an `ItemHasKnownEntity` element in the manifest, but there are no entities in the current item that match, the method return an empty array.</span></span>

<span data-ttu-id="a7c7c-1060">Typ: Array.<(String|[Contact](/javascript/api/outlook/office.contact)|[MeetingSuggestion](/javascript/api/outlook/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook/office.tasksuggestion))></span><span class="sxs-lookup"><span data-stu-id="a7c7c-1060">Type: Array.<(String|[Contact](/javascript/api/outlook/office.contact)|[MeetingSuggestion](/javascript/api/outlook/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook/office.tasksuggestion))></span></span>

#### <a name="getinitializationcontextasyncoptions-callback"></a><span data-ttu-id="a7c7c-1061">getInitializationContextAsync([options], [callback])</span><span class="sxs-lookup"><span data-stu-id="a7c7c-1061">getInitializationContextAsync([options], [callback])</span></span>

<span data-ttu-id="a7c7c-1062">Ruft Initialisierungsdaten ab, die übergeben werden, wenn das Add-In [durch eine Nachricht mit Aktionen aktiviert](https://docs.microsoft.com/outlook/actionable-messages/invoke-add-in-from-actionable-message) wird.</span><span class="sxs-lookup"><span data-stu-id="a7c7c-1062">Gets initialization data passed when the add-in is [activated by an actionable message](https://docs.microsoft.com/outlook/actionable-messages/invoke-add-in-from-actionable-message).</span></span>

> [!NOTE]
> <span data-ttu-id="a7c7c-1063">Diese Methode wird nur vom Outlook 2016 oder höher für Windows (höher als 16.0.8413.1000 Klick-und-Los-Versionen) und Outlook im Web für Office 365 unterstützt.</span><span class="sxs-lookup"><span data-stu-id="a7c7c-1063">This method is only supported by Outlook 2016 or later for Windows (Click-to-Run versions later than 16.0.8413.1000) and Outlook on the web for Office 365.</span></span>

##### <a name="parameters"></a><span data-ttu-id="a7c7c-1064">Parameter:</span><span class="sxs-lookup"><span data-stu-id="a7c7c-1064">Parameters:</span></span>
|<span data-ttu-id="a7c7c-1065">Name</span><span class="sxs-lookup"><span data-stu-id="a7c7c-1065">Name</span></span>|<span data-ttu-id="a7c7c-1066">Typ</span><span class="sxs-lookup"><span data-stu-id="a7c7c-1066">Type</span></span>|<span data-ttu-id="a7c7c-1067">Attribute</span><span class="sxs-lookup"><span data-stu-id="a7c7c-1067">Attributes</span></span>|<span data-ttu-id="a7c7c-1068">Beschreibung</span><span class="sxs-lookup"><span data-stu-id="a7c7c-1068">Description</span></span>|
|---|---|---|---|
|`options`|<span data-ttu-id="a7c7c-1069">Objekt</span><span class="sxs-lookup"><span data-stu-id="a7c7c-1069">Object</span></span>|<span data-ttu-id="a7c7c-1070">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="a7c7c-1070">&lt;optional&gt;</span></span>|<span data-ttu-id="a7c7c-1071">Ein Objektliteral, das mindestens eine der folgenden Eigenschaften enthält.</span><span class="sxs-lookup"><span data-stu-id="a7c7c-1071">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`|<span data-ttu-id="a7c7c-1072">Object</span><span class="sxs-lookup"><span data-stu-id="a7c7c-1072">Object</span></span>|<span data-ttu-id="a7c7c-1073">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="a7c7c-1073">&lt;optional&gt;</span></span>|<span data-ttu-id="a7c7c-1074">Entwickler können ein Objekt bereitstellen, auf das sie in der Callbackmethode zugreifen möchten.</span><span class="sxs-lookup"><span data-stu-id="a7c7c-1074">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`|<span data-ttu-id="a7c7c-1075">Funktion</span><span class="sxs-lookup"><span data-stu-id="a7c7c-1075">function</span></span>|<span data-ttu-id="a7c7c-1076">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="a7c7c-1076">&lt;optional&gt;</span></span>|<span data-ttu-id="a7c7c-1077">Wenn die Methode abgeschlossen wird, wird die Funktion , die im `callback`-Parameter übergeben wird, mit einem einzelnen Parameter (`asyncResult`) aufgerufen, der ein [`AsyncResult`](/javascript/api/office/office.asyncresult)-Objekt darstellt.</span><span class="sxs-lookup"><span data-stu-id="a7c7c-1077">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span> <br/><span data-ttu-id="a7c7c-1078">Initialisierungsdaten erfolgt auf Erfolg, in der `asyncResult.value` -Eigenschaft, wie eine Zeichenfolge.</span><span class="sxs-lookup"><span data-stu-id="a7c7c-1078">On success, the initialization data is provided in the `asyncResult.value` property as a string.</span></span><br/><span data-ttu-id="a7c7c-1079">Wenn es keinen Initialisierungskontext gibt, enthält das `asyncResult`-Objekt ein `Error`-Objekt, dessen `code`-Eigenschaft auf `9020` und dessen `name`-Eigenschaft auf `GenericResponseError` festgelegt ist.</span><span class="sxs-lookup"><span data-stu-id="a7c7c-1079">If there is no initialization context, the `asyncResult` object will contain an `Error` object with its `code` property set to `9020` and its `name` property set to `GenericResponseError`.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="a7c7c-1080">Anforderungen</span><span class="sxs-lookup"><span data-stu-id="a7c7c-1080">Requirements</span></span>

|<span data-ttu-id="a7c7c-1081">Anforderung</span><span class="sxs-lookup"><span data-stu-id="a7c7c-1081">Requirement</span></span>|<span data-ttu-id="a7c7c-1082">Wert</span><span class="sxs-lookup"><span data-stu-id="a7c7c-1082">Value</span></span>|
|---|---|
|[<span data-ttu-id="a7c7c-1083">Mindestversion des Postfachanforderungssatzes</span><span class="sxs-lookup"><span data-stu-id="a7c7c-1083">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="a7c7c-1084">Vorschau</span><span class="sxs-lookup"><span data-stu-id="a7c7c-1084">Preview</span></span>|
|[<span data-ttu-id="a7c7c-1085">Mindestberechtigungsstufe</span><span class="sxs-lookup"><span data-stu-id="a7c7c-1085">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="a7c7c-1086">ReadItem</span><span class="sxs-lookup"><span data-stu-id="a7c7c-1086">ReadItem</span></span>|
|[<span data-ttu-id="a7c7c-1087">Zutreffender Outlook-Modus</span><span class="sxs-lookup"><span data-stu-id="a7c7c-1087">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="a7c7c-1088">Lesen</span><span class="sxs-lookup"><span data-stu-id="a7c7c-1088">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="a7c7c-1089">Beispiel</span><span class="sxs-lookup"><span data-stu-id="a7c7c-1089">Example</span></span>

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

#### <a name="getregexmatches--object"></a><span data-ttu-id="a7c7c-1090">getRegExMatches() → {Object}</span><span class="sxs-lookup"><span data-stu-id="a7c7c-1090">getRegExMatches() → {Object}</span></span>

<span data-ttu-id="a7c7c-1091">Gibt Zeichenfolgenwerte im ausgewählten Element zurück, die den in der XML-Manifestdatei definierten regulären Ausdrücken entsprechen.</span><span class="sxs-lookup"><span data-stu-id="a7c7c-1091">Returns string values in the selected item that match the regular expressions defined in the manifest XML file.</span></span>

> [!NOTE]
> <span data-ttu-id="a7c7c-1092">Diese Methode wird in Outlook für iOS oder Outlook für Android nicht unterstützt.</span><span class="sxs-lookup"><span data-stu-id="a7c7c-1092">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="a7c7c-p161">Die `getRegExMatches`-Methode gibt die Zeichenfolgen zurück, die dem im `ItemHasRegularExpressionMatch`- oder `ItemHasKnownEntity`-Regelelement der XML-Manifestdatei definierten regulären Ausdruck entsprechen. Bei einer `ItemHasRegularExpressionMatch`-Regel muss eine entsprechende Zeichenfolge in der Eigenschaft des Elements vorhanden sein, das von dieser Regel angegeben wird. Der einfache `PropertyName`-Typ definiert die unterstützten Eigenschaften.</span><span class="sxs-lookup"><span data-stu-id="a7c7c-p161">The `getRegExMatches` method returns the strings that match the regular expression defined in each `ItemHasRegularExpressionMatch` or `ItemHasKnownEntity` rule element in the manifest XML file. For an `ItemHasRegularExpressionMatch` rule, a matching string has to occur in the property of the item that is specified by that rule. The `PropertyName` simple type defines the supported properties.</span></span>

<span data-ttu-id="a7c7c-1096">Nehmen Sie z. B. an, dass ein Add-In-Manifest das folgende `Rule`-Element aufweist:</span><span class="sxs-lookup"><span data-stu-id="a7c7c-1096">For example, consider an add-in manifest has the following `Rule` element:</span></span>

```
<Rule xsi:type="RuleCollection" Mode="And">
  <Rule xsi:type="ItemIs" FormType="Read" ItemType="Message" />
  <Rule xsi:type="RuleCollection" Mode="Or">
    <Rule xsi:type="ItemHasRegularExpressionMatch" RegExName="fruits" RegExValue="apple|banana|coconut" PropertyName="BodyAsPlaintext" IgnoreCase="true" />
    <Rule xsi:type="ItemHasRegularExpressionMatch" RegExName="veggies" RegExValue="tomato|onion|spinach|broccoli" PropertyName="BodyAsPlaintext" IgnoreCase="true" />
  </Rule>
</Rule>
```

<span data-ttu-id="a7c7c-1097">Das von `getRegExMatches` zurückgegebene Objekt hätte zwei Eigenschaften: `fruits` und `veggies`.</span><span class="sxs-lookup"><span data-stu-id="a7c7c-1097">The object returned from `getRegExMatches` would have two properties: `fruits` and `veggies`.</span></span>

```
{
  'fruits': ['apple','banana','Banana','coconut'],
  'veggies': ['tomato','onion','spinach','broccoli']
}
```

<span data-ttu-id="a7c7c-p162">Wenn Sie eine `ItemHasRegularExpressionMatch`-Regel für die Textkörpereigenschaft eines Elements festlegen, sollte der reguläre Ausdruck den Textkörper weiter filtern und nicht versuchen, den gesamten Textkörper des Elements zurückzugeben. Wenn der gesamte Textkörper eines Elements mit einem regulären Ausdruck wie `.*` abgerufen wird, werden nicht immer die gewünschten Ergebnisse erzielt. Verwenden Sie stattdessen die [`Body.getAsync`](/javascript/api/outlook/office.body#getasync-coerciontype--options--callback-)-Methode, um den gesamten Textkörper abzurufen.</span><span class="sxs-lookup"><span data-stu-id="a7c7c-p162">If you specify an `ItemHasRegularExpressionMatch` rule on the body property of an item, the regular expression should further filter the body and should not attempt to return the entire body of the item. Using a regular expression such as `.*` to obtain the entire body of an item does not always return the expected results. Instead, use the [`Body.getAsync`](/javascript/api/outlook/office.body#getasync-coerciontype--options--callback-) method to retrieve the entire body.</span></span>

##### <a name="requirements"></a><span data-ttu-id="a7c7c-1101">Anforderungen</span><span class="sxs-lookup"><span data-stu-id="a7c7c-1101">Requirements</span></span>

|<span data-ttu-id="a7c7c-1102">Anforderung</span><span class="sxs-lookup"><span data-stu-id="a7c7c-1102">Requirement</span></span>|<span data-ttu-id="a7c7c-1103">Wert</span><span class="sxs-lookup"><span data-stu-id="a7c7c-1103">Value</span></span>|
|---|---|
|[<span data-ttu-id="a7c7c-1104">Mindestversion des Postfachanforderungssatzes</span><span class="sxs-lookup"><span data-stu-id="a7c7c-1104">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="a7c7c-1105">1.0</span><span class="sxs-lookup"><span data-stu-id="a7c7c-1105">1.0</span></span>|
|[<span data-ttu-id="a7c7c-1106">Mindestberechtigungsstufe</span><span class="sxs-lookup"><span data-stu-id="a7c7c-1106">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="a7c7c-1107">ReadItem</span><span class="sxs-lookup"><span data-stu-id="a7c7c-1107">ReadItem</span></span>|
|[<span data-ttu-id="a7c7c-1108">Zutreffender Outlook-Modus</span><span class="sxs-lookup"><span data-stu-id="a7c7c-1108">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="a7c7c-1109">Lesen</span><span class="sxs-lookup"><span data-stu-id="a7c7c-1109">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="a7c7c-1110">Gibt zurück:</span><span class="sxs-lookup"><span data-stu-id="a7c7c-1110">Returns:</span></span>

<span data-ttu-id="a7c7c-p163">Ein Objekt mit Arrays aus Zeichenfolgen, die den in der XML-Manifestdatei definierten regulären Ausdrücken entsprechen. Der Name der einzelnen Arrays ist gleich dem entsprechenden Wert des `RegExName`-Attributs der entsprechenden `ItemHasRegularExpressionMatch`-Regel oder des `FilterName`-Attributs der entsprechenden `ItemHasKnownEntity`-Regel.</span><span class="sxs-lookup"><span data-stu-id="a7c7c-p163">An object that contains arrays of strings that match the regular expressions defined in the manifest XML file. The name of each array is equal to the corresponding value of the `RegExName` attribute of the matching `ItemHasRegularExpressionMatch` rule or the `FilterName` attribute of the matching `ItemHasKnownEntity` rule.</span></span>

<dl class="param-type"><span data-ttu-id="a7c7c-1113">

<dt>
Typ</dt>


</span><span class="sxs-lookup"><span data-stu-id="a7c7c-1113">

<dt>Type</dt>

</span></span><dd><span data-ttu-id="a7c7c-1114">Object</span><span class="sxs-lookup"><span data-stu-id="a7c7c-1114">Object</span></span></dd>

</dl>

##### <a name="example"></a><span data-ttu-id="a7c7c-1115">Beispiel</span><span class="sxs-lookup"><span data-stu-id="a7c7c-1115">Example</span></span>

<span data-ttu-id="a7c7c-1116">Das folgende Beispiel zeigt, wie Sie auf das Array von Übereinstimmungen für die regulären Ausdrucksregelelemente `fruits` und `veggies` zugreifen, die im Manifest angegeben sind.</span><span class="sxs-lookup"><span data-stu-id="a7c7c-1116">The following example shows how to access the array of matches for the regular expression rule elements `fruits` and `veggies`, which are specified in the manifest.</span></span>

```
var allMatches = Office.context.mailbox.item.getRegExMatches();
var fruits = allMatches.fruits;
var veges = allMatches.veggies;
```

#### <a name="getregexmatchesbynamename--nullable-array-string-"></a><span data-ttu-id="a7c7c-1117">getRegExMatchesByName(name) → (Nullwerte zulassen) {Array. < Zeichenfolge >}</span><span class="sxs-lookup"><span data-stu-id="a7c7c-1117">getRegExMatchesByName(name) → (nullable) {Array.< String >}</span></span>

<span data-ttu-id="a7c7c-1118">Gibt Zeichenfolgenwerte im ausgewählten Element zurück, die dem in der XML-Manifestdatei definierten benannten regulären Ausdruck entsprechen.</span><span class="sxs-lookup"><span data-stu-id="a7c7c-1118">Returns string values in the selected item that match the named regular expression defined in the manifest XML file.</span></span>

> [!NOTE]
> <span data-ttu-id="a7c7c-1119">Diese Methode wird in Outlook für iOS oder Outlook für Android nicht unterstützt.</span><span class="sxs-lookup"><span data-stu-id="a7c7c-1119">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="a7c7c-1120">Die `getRegExMatchesByName`-Methode gibt die Zeichenfolgen zurück, die dem im `ItemHasRegularExpressionMatch`-Regelelement der XML-Manifestdatei definierten regulären Ausdruck mit dem angegebenen `RegExName`-Elementwert entsprechen.</span><span class="sxs-lookup"><span data-stu-id="a7c7c-1120">The `getRegExMatchesByName` method returns the strings that match the regular expression defined in the `ItemHasRegularExpressionMatch` rule element in the manifest XML file with the specified `RegExName` element value.</span></span>

<span data-ttu-id="a7c7c-p164">Wenn Sie eine `ItemHasRegularExpressionMatch`-Regel für die Textkörpereigenschaft eines Elements festlegen, sollte der reguläre Ausdruck den Textkörper weiter filtern und nicht versuchen, den gesamten Textkörper des Elements zurückzugeben. Wenn der gesamte Textkörper eines Elements mit einem regulären Ausdruck wie `.*` abgerufen wird, werden nicht immer die gewünschten Ergebnisse erzielt.</span><span class="sxs-lookup"><span data-stu-id="a7c7c-p164">If you specify an `ItemHasRegularExpressionMatch` rule on the body property of an item, the regular expression should further filter the body and should not attempt to return the entire body of the item. Using a regular expression such as `.*` to obtain the entire body of an item does not always return the expected results.</span></span>

##### <a name="parameters"></a><span data-ttu-id="a7c7c-1123">Parameter:</span><span class="sxs-lookup"><span data-stu-id="a7c7c-1123">Parameters:</span></span>

|<span data-ttu-id="a7c7c-1124">Name</span><span class="sxs-lookup"><span data-stu-id="a7c7c-1124">Name</span></span>|<span data-ttu-id="a7c7c-1125">Typ</span><span class="sxs-lookup"><span data-stu-id="a7c7c-1125">Type</span></span>|<span data-ttu-id="a7c7c-1126">Beschreibung</span><span class="sxs-lookup"><span data-stu-id="a7c7c-1126">Description</span></span>|
|---|---|---|
|`name`|<span data-ttu-id="a7c7c-1127">Zeichenfolge</span><span class="sxs-lookup"><span data-stu-id="a7c7c-1127">String</span></span>|<span data-ttu-id="a7c7c-1128">Der Name des `ItemHasRegularExpressionMatch`-Regelelements, das den entsprechenden Filter definiert.</span><span class="sxs-lookup"><span data-stu-id="a7c7c-1128">The name of the `ItemHasRegularExpressionMatch` rule element that defines the filter to match.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="a7c7c-1129">Anforderungen</span><span class="sxs-lookup"><span data-stu-id="a7c7c-1129">Requirements</span></span>

|<span data-ttu-id="a7c7c-1130">Anforderung</span><span class="sxs-lookup"><span data-stu-id="a7c7c-1130">Requirement</span></span>|<span data-ttu-id="a7c7c-1131">Wert</span><span class="sxs-lookup"><span data-stu-id="a7c7c-1131">Value</span></span>|
|---|---|
|[<span data-ttu-id="a7c7c-1132">Mindestversion des Postfachanforderungssatzes</span><span class="sxs-lookup"><span data-stu-id="a7c7c-1132">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="a7c7c-1133">1.0</span><span class="sxs-lookup"><span data-stu-id="a7c7c-1133">1.0</span></span>|
|[<span data-ttu-id="a7c7c-1134">Mindestberechtigungsstufe</span><span class="sxs-lookup"><span data-stu-id="a7c7c-1134">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="a7c7c-1135">ReadItem</span><span class="sxs-lookup"><span data-stu-id="a7c7c-1135">ReadItem</span></span>|
|[<span data-ttu-id="a7c7c-1136">Zutreffender Outlook-Modus</span><span class="sxs-lookup"><span data-stu-id="a7c7c-1136">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="a7c7c-1137">Lesen</span><span class="sxs-lookup"><span data-stu-id="a7c7c-1137">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="a7c7c-1138">Gibt zurück:</span><span class="sxs-lookup"><span data-stu-id="a7c7c-1138">Returns:</span></span>

<span data-ttu-id="a7c7c-1139">Ein Array mit den Zeichenfolgen, die dem in der XML-Manifestdatei definierten regulären Ausdruck entsprechen.</span><span class="sxs-lookup"><span data-stu-id="a7c7c-1139">An array that contains the strings that match the regular expression defined in the manifest XML file.</span></span>

<dl class="param-type"><span data-ttu-id="a7c7c-1140">

<dt>Typ</dt>

</span><span class="sxs-lookup"><span data-stu-id="a7c7c-1140">

<dt>Type</dt>

</span></span><dd><span data-ttu-id="a7c7c-1141">Array. < Zeichenfolge ></span><span class="sxs-lookup"><span data-stu-id="a7c7c-1141">Array.< String ></span></span></dd>

</dl>

##### <a name="example"></a><span data-ttu-id="a7c7c-1142">Beispiel</span><span class="sxs-lookup"><span data-stu-id="a7c7c-1142">Example</span></span>

```
var fruits = Office.context.mailbox.item.getRegExMatchesByName("fruits");
var veggies = Office.context.mailbox.item.getRegExMatchesByName("veggies");
```

####  <a name="getselecteddataasynccoerciontype-options-callback--string"></a><span data-ttu-id="a7c7c-1143">getSelectedDataAsync(coercionType, [options], callback) → {String}</span><span class="sxs-lookup"><span data-stu-id="a7c7c-1143">getSelectedDataAsync(coercionType, [options], callback) → {String}</span></span>

<span data-ttu-id="a7c7c-1144">Gibt asynchron ausgewählte Daten aus dem Betreff oder Textkörper einer Nachricht zurück.</span><span class="sxs-lookup"><span data-stu-id="a7c7c-1144">Asynchronously returns selected data from the subject or body of a message.</span></span>

<span data-ttu-id="a7c7c-p165">Wenn keine Auswahl vorhanden ist, aber der Cursor sich im Nachrichtentext oder Betreff befindet, gibt die Methode für die ausgewählten Daten NULL zurück. Wenn ein anderes Feld als der Textkörper oder Betreff ausgewählt ist, gibt die Methode den `InvalidSelection`-Fehler zurück.</span><span class="sxs-lookup"><span data-stu-id="a7c7c-p165">If there is no selection but the cursor is in the body or subject, the method returns null for the selected data. If a field other than the body or subject is selected, the method returns the `InvalidSelection` error.</span></span>

##### <a name="parameters"></a><span data-ttu-id="a7c7c-1147">Parameter:</span><span class="sxs-lookup"><span data-stu-id="a7c7c-1147">Parameters:</span></span>

|<span data-ttu-id="a7c7c-1148">Name</span><span class="sxs-lookup"><span data-stu-id="a7c7c-1148">Name</span></span>|<span data-ttu-id="a7c7c-1149">Typ</span><span class="sxs-lookup"><span data-stu-id="a7c7c-1149">Type</span></span>|<span data-ttu-id="a7c7c-1150">Attribute</span><span class="sxs-lookup"><span data-stu-id="a7c7c-1150">Attributes</span></span>|<span data-ttu-id="a7c7c-1151">Beschreibung</span><span class="sxs-lookup"><span data-stu-id="a7c7c-1151">Description</span></span>|
|---|---|---|---|
|`coercionType`|[<span data-ttu-id="a7c7c-1152">Office.CoercionType</span><span class="sxs-lookup"><span data-stu-id="a7c7c-1152">Office.CoercionType</span></span>](office.md#coerciontype-string)||<span data-ttu-id="a7c7c-p166">Fordert ein Format für die Daten an. Wenn es sich um Texthandelt, gibt die Methode den unformatierten Text als Zeichenfolge zurück und entfernt eventuell vorhandene HTML-Tags. Wenn es sich um HTML handelt, gibt die Methode den ausgewählten Text zurück, entweder als unformatierten Text oder als HTML.</span><span class="sxs-lookup"><span data-stu-id="a7c7c-p166">Requests a format for the data. If Text, the method returns the plain text as a string , removing any HTML tags present. If HTML, the method returns the selected text, whether it is plaintext or HTML.</span></span>|
|`options`|<span data-ttu-id="a7c7c-1156">Object</span><span class="sxs-lookup"><span data-stu-id="a7c7c-1156">Object</span></span>|<span data-ttu-id="a7c7c-1157">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="a7c7c-1157">&lt;optional&gt;</span></span>|<span data-ttu-id="a7c7c-1158">Ein Objektliteral, das mindestens eine der folgenden Eigenschaften enthält.</span><span class="sxs-lookup"><span data-stu-id="a7c7c-1158">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`|<span data-ttu-id="a7c7c-1159">Object</span><span class="sxs-lookup"><span data-stu-id="a7c7c-1159">Object</span></span>|<span data-ttu-id="a7c7c-1160">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="a7c7c-1160">&lt;optional&gt;</span></span>|<span data-ttu-id="a7c7c-1161">Entwickler können ein Objekt bereitstellen, auf das sie in der Callbackmethode zugreifen möchten.</span><span class="sxs-lookup"><span data-stu-id="a7c7c-1161">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`|<span data-ttu-id="a7c7c-1162">Funktion</span><span class="sxs-lookup"><span data-stu-id="a7c7c-1162">function</span></span>||<span data-ttu-id="a7c7c-1163">Wenn die Methode abgeschlossen wird, wird die Funktion , die im `callback`-Parameter übergeben wird, mit einem einzelnen Parameter (`asyncResult`) aufgerufen, der ein [`AsyncResult`](/javascript/api/office/office.asyncresult)-Objekt darstellt.</span><span class="sxs-lookup"><span data-stu-id="a7c7c-1163">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span><br/><br/><span data-ttu-id="a7c7c-1164">Rufen Sie für den Zugriff auf die ausgewählten Daten aus der Rückrufmethode `asyncResult.value.data` auf.</span><span class="sxs-lookup"><span data-stu-id="a7c7c-1164">To access the selected data from the callback method, call `asyncResult.value.data`.</span></span> <span data-ttu-id="a7c7c-1165">Rufen Sie die Source-Eigenschaft für den Zugriff, die die Auswahl stammen, `asyncResult.value.sourceProperty`, die entweder sein `body` oder `subject`.</span><span class="sxs-lookup"><span data-stu-id="a7c7c-1165">To access the source property that the selection comes from, call `asyncResult.value.sourceProperty`, which will be either `body` or `subject`.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="a7c7c-1166">Anforderungen</span><span class="sxs-lookup"><span data-stu-id="a7c7c-1166">Requirements</span></span>

|<span data-ttu-id="a7c7c-1167">Anforderung</span><span class="sxs-lookup"><span data-stu-id="a7c7c-1167">Requirement</span></span>|<span data-ttu-id="a7c7c-1168">Wert</span><span class="sxs-lookup"><span data-stu-id="a7c7c-1168">Value</span></span>|
|---|---|
|[<span data-ttu-id="a7c7c-1169">Mindestversion des Postfachanforderungssatzes</span><span class="sxs-lookup"><span data-stu-id="a7c7c-1169">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="a7c7c-1170">1.2</span><span class="sxs-lookup"><span data-stu-id="a7c7c-1170">1.2</span></span>|
|[<span data-ttu-id="a7c7c-1171">Mindestberechtigungsstufe</span><span class="sxs-lookup"><span data-stu-id="a7c7c-1171">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="a7c7c-1172">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="a7c7c-1172">ReadWriteItem</span></span>|
|[<span data-ttu-id="a7c7c-1173">Zutreffender Outlook-Modus</span><span class="sxs-lookup"><span data-stu-id="a7c7c-1173">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="a7c7c-1174">Verfassen</span><span class="sxs-lookup"><span data-stu-id="a7c7c-1174">Compose</span></span>|

##### <a name="returns"></a><span data-ttu-id="a7c7c-1175">Gibt zurück:</span><span class="sxs-lookup"><span data-stu-id="a7c7c-1175">Returns:</span></span>

<span data-ttu-id="a7c7c-1176">Die ausgewählten Daten als Zeichenfolge mit dem durch `coercionType` bestimmten Format.</span><span class="sxs-lookup"><span data-stu-id="a7c7c-1176">The selected data as a string with format determined by `coercionType`.</span></span>

<dl class="param-type"><span data-ttu-id="a7c7c-1177">

<dt>
Typ</dt>


</span><span class="sxs-lookup"><span data-stu-id="a7c7c-1177">

<dt>Type</dt>

</span></span><dd><span data-ttu-id="a7c7c-1178">Zeichenfolge</span><span class="sxs-lookup"><span data-stu-id="a7c7c-1178">String</span></span></dd>

</dl>

##### <a name="example"></a><span data-ttu-id="a7c7c-1179">Beispiel</span><span class="sxs-lookup"><span data-stu-id="a7c7c-1179">Example</span></span>

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

#### <a name="getselectedentities--entitiesjavascriptapioutlookofficeentities"></a><span data-ttu-id="a7c7c-1180">getSelectedEntities() → {[Entitäten](/javascript/api/outlook/office.entities)}</span><span class="sxs-lookup"><span data-stu-id="a7c7c-1180">getSelectedEntities() → {[Entities](/javascript/api/outlook/office.entities)}</span></span>

<span data-ttu-id="a7c7c-p168">Ruft die Entitäten ab, die in einer hervorgehobenen Übereinstimmung gefunden werden, die ein Benutzer ausgewählt hat. Hervorgehobene Übereinstimmungen gelten für [Kontext-Add-Ins](https://docs.microsoft.com/outlook/add-ins/contextual-outlook-add-ins).</span><span class="sxs-lookup"><span data-stu-id="a7c7c-p168">Gets the entities found in a highlighted match a user has selected. Highlighted matches apply to [contextual add-ins](https://docs.microsoft.com/outlook/add-ins/contextual-outlook-add-ins).</span></span>

> [!NOTE]
> <span data-ttu-id="a7c7c-1183">Diese Methode wird in Outlook für iOS oder Outlook für Android nicht unterstützt.</span><span class="sxs-lookup"><span data-stu-id="a7c7c-1183">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

##### <a name="requirements"></a><span data-ttu-id="a7c7c-1184">Anforderungen</span><span class="sxs-lookup"><span data-stu-id="a7c7c-1184">Requirements</span></span>

|<span data-ttu-id="a7c7c-1185">Anforderung</span><span class="sxs-lookup"><span data-stu-id="a7c7c-1185">Requirement</span></span>|<span data-ttu-id="a7c7c-1186">Wert</span><span class="sxs-lookup"><span data-stu-id="a7c7c-1186">Value</span></span>|
|---|---|
|[<span data-ttu-id="a7c7c-1187">Mindestversion des Postfachanforderungssatzes</span><span class="sxs-lookup"><span data-stu-id="a7c7c-1187">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="a7c7c-1188">1.6</span><span class="sxs-lookup"><span data-stu-id="a7c7c-1188">1.6</span></span>|
|[<span data-ttu-id="a7c7c-1189">Mindestberechtigungsstufe</span><span class="sxs-lookup"><span data-stu-id="a7c7c-1189">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="a7c7c-1190">ReadItem</span><span class="sxs-lookup"><span data-stu-id="a7c7c-1190">ReadItem</span></span>|
|[<span data-ttu-id="a7c7c-1191">Zutreffender Outlook-Modus</span><span class="sxs-lookup"><span data-stu-id="a7c7c-1191">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="a7c7c-1192">Lesen</span><span class="sxs-lookup"><span data-stu-id="a7c7c-1192">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="a7c7c-1193">Gibt zurück:</span><span class="sxs-lookup"><span data-stu-id="a7c7c-1193">Returns:</span></span>

<span data-ttu-id="a7c7c-1194">Typ: [Entitäten](/javascript/api/outlook/office.entities)</span><span class="sxs-lookup"><span data-stu-id="a7c7c-1194">Type: [Entities](/javascript/api/outlook/office.entities)</span></span>

##### <a name="example"></a><span data-ttu-id="a7c7c-1195">Beispiel</span><span class="sxs-lookup"><span data-stu-id="a7c7c-1195">Example</span></span>

<span data-ttu-id="a7c7c-1196">Im folgenden Beispiel wird auf die Adressentitäten in der hervorgehobenen Übereinstimmung zugegriffen, die der Benutzer ausgewählt hat.</span><span class="sxs-lookup"><span data-stu-id="a7c7c-1196">The following example accesses the addresses entities in the highlighted match selected by the user.</span></span>

```
var contacts = Office.context.mailbox.item.getSelectedEntities().addresses;
```

#### <a name="getselectedregexmatches--object"></a><span data-ttu-id="a7c7c-1197">getSelectedRegExMatches() → {Object}</span><span class="sxs-lookup"><span data-stu-id="a7c7c-1197">getSelectedRegExMatches() → {Object}</span></span>

<span data-ttu-id="a7c7c-p169">Gibt Zeichenfolgenwerte in einer hervorgehobenen Übereinstimmung zurück, die den in der XML-Manifestdatei definierten regulären Ausdrücken entsprechen. Hervorgehobene Übereinstimmungen gelten für [Kontext-Add-Ins](https://docs.microsoft.com/outlook/add-ins/contextual-outlook-add-ins).</span><span class="sxs-lookup"><span data-stu-id="a7c7c-p169">Returns string values in a highlighted match that match the regular expressions defined in the manifest XML file. Highlighted matches apply to [contextual add-ins](https://docs.microsoft.com/outlook/add-ins/contextual-outlook-add-ins).</span></span>

> [!NOTE]
> <span data-ttu-id="a7c7c-1200">Diese Methode wird in Outlook für iOS oder Outlook für Android nicht unterstützt.</span><span class="sxs-lookup"><span data-stu-id="a7c7c-1200">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="a7c7c-p170">Die `getSelectedRegExMatches`-Methode gibt die Zeichenfolgen zurück, die dem im `ItemHasRegularExpressionMatch`- oder `ItemHasKnownEntity`-Regelelement der XML-Manifestdatei definierten regulären Ausdruck entsprechen. Bei einer `ItemHasRegularExpressionMatch`-Regel muss eine entsprechende Zeichenfolge in der Eigenschaft des Elements vorhanden sein, das von dieser Regel angegeben wird. Der einfache `PropertyName`-Typ definiert die unterstützten Eigenschaften.</span><span class="sxs-lookup"><span data-stu-id="a7c7c-p170">The `getSelectedRegExMatches` method returns the strings that match the regular expression defined in each `ItemHasRegularExpressionMatch` or `ItemHasKnownEntity` rule element in the manifest XML file. For an `ItemHasRegularExpressionMatch` rule, a matching string has to occur in the property of the item that is specified by that rule. The `PropertyName` simple type defines the supported properties.</span></span>

<span data-ttu-id="a7c7c-1204">Nehmen Sie z. B. an, dass ein Add-In-Manifest das folgende `Rule`-Element aufweist:</span><span class="sxs-lookup"><span data-stu-id="a7c7c-1204">For example, consider an add-in manifest has the following `Rule` element:</span></span>

```
<Rule xsi:type="RuleCollection" Mode="And">
  <Rule xsi:type="ItemIs" FormType="Read" ItemType="Message" />
  <Rule xsi:type="RuleCollection" Mode="Or">
    <Rule xsi:type="ItemHasRegularExpressionMatch" RegExName="fruits" RegExValue="apple|banana|coconut" PropertyName="BodyAsPlaintext" IgnoreCase="true" />
    <Rule xsi:type="ItemHasRegularExpressionMatch" RegExName="veggies" RegExValue="tomato|onion|spinach|broccoli" PropertyName="BodyAsPlaintext" IgnoreCase="true" />
  </Rule>
</Rule>
```

<span data-ttu-id="a7c7c-1205">Das von `getRegExMatches` zurückgegebene Objekt hätte zwei Eigenschaften: `fruits` und `veggies`.</span><span class="sxs-lookup"><span data-stu-id="a7c7c-1205">The object returned from `getRegExMatches` would have two properties: `fruits` and `veggies`.</span></span>

```
{
  'fruits': ['apple','banana','Banana','coconut'],
  'veggies': ['tomato','onion','spinach','broccoli']
}
```

<span data-ttu-id="a7c7c-p171">Wenn Sie eine `ItemHasRegularExpressionMatch`-Regel für die Textkörpereigenschaft eines Elements festlegen, sollte der reguläre Ausdruck den Textkörper weiter filtern und nicht versuchen, den gesamten Textkörper des Elements zurückzugeben. Wenn der gesamte Textkörper eines Elements mit einem regulären Ausdruck wie `.*` abgerufen wird, werden nicht immer die gewünschten Ergebnisse erzielt. Verwenden Sie stattdessen die [`Body.getAsync`](/javascript/api/outlook/office.body#getasync-coerciontype--options--callback-)-Methode, um den gesamten Textkörper abzurufen.</span><span class="sxs-lookup"><span data-stu-id="a7c7c-p171">If you specify an `ItemHasRegularExpressionMatch` rule on the body property of an item, the regular expression should further filter the body and should not attempt to return the entire body of the item. Using a regular expression such as `.*` to obtain the entire body of an item does not always return the expected results. Instead, use the [`Body.getAsync`](/javascript/api/outlook/office.body#getasync-coerciontype--options--callback-) method to retrieve the entire body.</span></span>

##### <a name="requirements"></a><span data-ttu-id="a7c7c-1209">Anforderungen</span><span class="sxs-lookup"><span data-stu-id="a7c7c-1209">Requirements</span></span>

|<span data-ttu-id="a7c7c-1210">Anforderung</span><span class="sxs-lookup"><span data-stu-id="a7c7c-1210">Requirement</span></span>|<span data-ttu-id="a7c7c-1211">Wert</span><span class="sxs-lookup"><span data-stu-id="a7c7c-1211">Value</span></span>|
|---|---|
|[<span data-ttu-id="a7c7c-1212">Mindestversion des Postfachanforderungssatzes</span><span class="sxs-lookup"><span data-stu-id="a7c7c-1212">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="a7c7c-1213">1.6</span><span class="sxs-lookup"><span data-stu-id="a7c7c-1213">1.6</span></span>|
|[<span data-ttu-id="a7c7c-1214">Mindestberechtigungsstufe</span><span class="sxs-lookup"><span data-stu-id="a7c7c-1214">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="a7c7c-1215">ReadItem</span><span class="sxs-lookup"><span data-stu-id="a7c7c-1215">ReadItem</span></span>|
|[<span data-ttu-id="a7c7c-1216">Zutreffender Outlook-Modus</span><span class="sxs-lookup"><span data-stu-id="a7c7c-1216">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="a7c7c-1217">Lesen</span><span class="sxs-lookup"><span data-stu-id="a7c7c-1217">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="a7c7c-1218">Gibt zurück:</span><span class="sxs-lookup"><span data-stu-id="a7c7c-1218">Returns:</span></span>

<span data-ttu-id="a7c7c-p172">Ein Objekt mit Arrays aus Zeichenfolgen, die den in der XML-Manifestdatei definierten regulären Ausdrücken entsprechen. Der Name der einzelnen Arrays ist gleich dem entsprechenden Wert des `RegExName`-Attributs der entsprechenden `ItemHasRegularExpressionMatch`-Regel oder des `FilterName`-Attributs der entsprechenden `ItemHasKnownEntity`-Regel.</span><span class="sxs-lookup"><span data-stu-id="a7c7c-p172">An object that contains arrays of strings that match the regular expressions defined in the manifest XML file. The name of each array is equal to the corresponding value of the `RegExName` attribute of the matching `ItemHasRegularExpressionMatch` rule or the `FilterName` attribute of the matching `ItemHasKnownEntity` rule.</span></span>

##### <a name="example"></a><span data-ttu-id="a7c7c-1221">Beispiel</span><span class="sxs-lookup"><span data-stu-id="a7c7c-1221">Example</span></span>

<span data-ttu-id="a7c7c-1222">Das folgende Beispiel zeigt, wie Sie auf das Array von Übereinstimmungen für die regulären Ausdrucksregelelemente `fruits` und `veggies` zugreifen, die im Manifest angegeben sind.</span><span class="sxs-lookup"><span data-stu-id="a7c7c-1222">The following example shows how to access the array of matches for the regular expression rule elements `fruits` and `veggies`, which are specified in the manifest.</span></span>

```
var selectedMatches = Office.context.mailbox.item.getSelectedRegExMatches();
var fruits = selectedMatches.fruits;
var veggies = selectedMatches.veggies;
```

#### <a name="getsharedpropertiesasyncoptions-callback"></a><span data-ttu-id="a7c7c-1223">GetSharedPropertiesAsync ([Optionen] Rückruf)</span><span class="sxs-lookup"><span data-stu-id="a7c7c-1223">getSharedPropertiesAsync([options], callback)</span></span>

<span data-ttu-id="a7c7c-1224">Ruft die Eigenschaften des ausgewählten Termins oder einer Nachricht in einen freigegebenen Ordner, Kalender oder Postfach an.</span><span class="sxs-lookup"><span data-stu-id="a7c7c-1224">Gets the properties of the selected appointment or message in a shared folder, calendar, or mailbox.</span></span>

##### <a name="parameters"></a><span data-ttu-id="a7c7c-1225">Parameter:</span><span class="sxs-lookup"><span data-stu-id="a7c7c-1225">Parameters:</span></span>

|<span data-ttu-id="a7c7c-1226">Name</span><span class="sxs-lookup"><span data-stu-id="a7c7c-1226">Name</span></span>|<span data-ttu-id="a7c7c-1227">Typ</span><span class="sxs-lookup"><span data-stu-id="a7c7c-1227">Type</span></span>|<span data-ttu-id="a7c7c-1228">Attribute</span><span class="sxs-lookup"><span data-stu-id="a7c7c-1228">Attributes</span></span>|<span data-ttu-id="a7c7c-1229">Beschreibung</span><span class="sxs-lookup"><span data-stu-id="a7c7c-1229">Description</span></span>|
|---|---|---|---|
|`options`|<span data-ttu-id="a7c7c-1230">Objekt</span><span class="sxs-lookup"><span data-stu-id="a7c7c-1230">Object</span></span>|<span data-ttu-id="a7c7c-1231">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="a7c7c-1231">&lt;optional&gt;</span></span>|<span data-ttu-id="a7c7c-1232">Ein Objektliteral, das mindestens eine der folgenden Eigenschaften enthält.</span><span class="sxs-lookup"><span data-stu-id="a7c7c-1232">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`|<span data-ttu-id="a7c7c-1233">Object</span><span class="sxs-lookup"><span data-stu-id="a7c7c-1233">Object</span></span>|<span data-ttu-id="a7c7c-1234">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="a7c7c-1234">&lt;optional&gt;</span></span>|<span data-ttu-id="a7c7c-1235">Entwickler können ein Objekt bereitstellen, auf das sie in der Callbackmethode zugreifen möchten.</span><span class="sxs-lookup"><span data-stu-id="a7c7c-1235">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`|<span data-ttu-id="a7c7c-1236">Funktion</span><span class="sxs-lookup"><span data-stu-id="a7c7c-1236">function</span></span>||<span data-ttu-id="a7c7c-1237">Wenn die Methode abgeschlossen wird, wird die Funktion , die im `callback`-Parameter übergeben wird, mit einem einzelnen Parameter (`asyncResult`) aufgerufen, der ein [`AsyncResult`](/javascript/api/office/office.asyncresult)-Objekt darstellt.</span><span class="sxs-lookup"><span data-stu-id="a7c7c-1237">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span><br/><br/><span data-ttu-id="a7c7c-1238">Freigegebenen Eigenschaften dienen als ein [`SharedProperties`](/javascript/api/outlook/office.sharedproperties) -Objekts der `asyncResult.value` Eigenschaft.</span><span class="sxs-lookup"><span data-stu-id="a7c7c-1238">The shared properties are provided as a [`SharedProperties`](/javascript/api/outlook/office.sharedproperties) object in the `asyncResult.value` property.</span></span> <span data-ttu-id="a7c7c-1239">Dieses Objekt kann zum Abrufen von freigegebenen Elementeigenschaften verwendet werden.</span><span class="sxs-lookup"><span data-stu-id="a7c7c-1239">This object can be used to get the item's shared properties.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="a7c7c-1240">Anforderungen</span><span class="sxs-lookup"><span data-stu-id="a7c7c-1240">Requirements</span></span>

|<span data-ttu-id="a7c7c-1241">Anforderung</span><span class="sxs-lookup"><span data-stu-id="a7c7c-1241">Requirement</span></span>|<span data-ttu-id="a7c7c-1242">Wert</span><span class="sxs-lookup"><span data-stu-id="a7c7c-1242">Value</span></span>|
|---|---|
|[<span data-ttu-id="a7c7c-1243">Mindestversion des Postfachanforderungssatzes</span><span class="sxs-lookup"><span data-stu-id="a7c7c-1243">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="a7c7c-1244">Vorschau</span><span class="sxs-lookup"><span data-stu-id="a7c7c-1244">Preview</span></span>|
|[<span data-ttu-id="a7c7c-1245">Mindestberechtigungsstufe</span><span class="sxs-lookup"><span data-stu-id="a7c7c-1245">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="a7c7c-1246">ReadItem</span><span class="sxs-lookup"><span data-stu-id="a7c7c-1246">ReadItem</span></span>|
|[<span data-ttu-id="a7c7c-1247">Zutreffender Outlook-Modus</span><span class="sxs-lookup"><span data-stu-id="a7c7c-1247">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="a7c7c-1248">Verfassen oder Lesen</span><span class="sxs-lookup"><span data-stu-id="a7c7c-1248">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="a7c7c-1249">Beispiel</span><span class="sxs-lookup"><span data-stu-id="a7c7c-1249">Example</span></span>

```js
Office.context.mailbox.item.getSharedPropertiesAsync(callback);
function callback (asyncResult) {
  var context=asyncResult.context;
  var sharedProperties = asyncResult.value;
}
```

####  <a name="loadcustompropertiesasynccallback-usercontext"></a><span data-ttu-id="a7c7c-1250">loadCustomPropertiesAsync(callback, [userContext])</span><span class="sxs-lookup"><span data-stu-id="a7c7c-1250">loadCustomPropertiesAsync(callback, [userContext])</span></span>

<span data-ttu-id="a7c7c-1251">Lädt asynchron benutzerdefinierte Eigenschaften für dieses Add-In für das ausgewählte Element.</span><span class="sxs-lookup"><span data-stu-id="a7c7c-1251">Asynchronously loads custom properties for this add-in on the selected item.</span></span>

<span data-ttu-id="a7c7c-p174">Benutzerdefinierte Eigenschaften werden als Schlüssel-/Wert-Paare pro App und pro Element gespeichert. Diese Methode gibt ein `CustomProperties`-Objekt im Rückruf zurück, das Methoden für den Zugriff auf die benutzerdefinierten Eigenschaften für das aktuelle Element und das aktuelle Add-In bietet. Benutzerdefinierte Eigenschaften sind für das Element nicht verschlüsselt und sollten darum nicht als sicherer Speicher verwendet werden.</span><span class="sxs-lookup"><span data-stu-id="a7c7c-p174">Custom properties are stored as key/value pairs on a per-app, per-item basis. This method returns a `CustomProperties` object in the callback, which provides methods to access the custom properties specific to the current item and the current add-in. Custom properties are not encrypted on the item, so this should not be used as secure storage.</span></span>

##### <a name="parameters"></a><span data-ttu-id="a7c7c-1255">Parameter:</span><span class="sxs-lookup"><span data-stu-id="a7c7c-1255">Parameters:</span></span>

|<span data-ttu-id="a7c7c-1256">Name</span><span class="sxs-lookup"><span data-stu-id="a7c7c-1256">Name</span></span>|<span data-ttu-id="a7c7c-1257">Typ</span><span class="sxs-lookup"><span data-stu-id="a7c7c-1257">Type</span></span>|<span data-ttu-id="a7c7c-1258">Attribute</span><span class="sxs-lookup"><span data-stu-id="a7c7c-1258">Attributes</span></span>|<span data-ttu-id="a7c7c-1259">Beschreibung</span><span class="sxs-lookup"><span data-stu-id="a7c7c-1259">Description</span></span>|
|---|---|---|---|
|`callback`|<span data-ttu-id="a7c7c-1260">Funktion</span><span class="sxs-lookup"><span data-stu-id="a7c7c-1260">function</span></span>||<span data-ttu-id="a7c7c-1261">Wenn die Methode abgeschlossen wird, wird die Funktion , die im `callback`-Parameter übergeben wird, mit einem einzelnen Parameter (`asyncResult`) aufgerufen, der ein [`AsyncResult`](/javascript/api/office/office.asyncresult)-Objekt darstellt.</span><span class="sxs-lookup"><span data-stu-id="a7c7c-1261">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span><br/><br/><span data-ttu-id="a7c7c-1262">Die benutzerdefinierten Eigenschaften werden als [`CustomProperties`](/javascript/api/outlook/office.customproperties)-Objekt in der `asyncResult.value`-Eigenschaft bereitgestellt.</span><span class="sxs-lookup"><span data-stu-id="a7c7c-1262">The custom properties are provided as a [`CustomProperties`](/javascript/api/outlook/office.customproperties) object in the `asyncResult.value` property.</span></span> <span data-ttu-id="a7c7c-1263">Dieses Objekt kann zum Abrufen, festlegen und Entfernen benutzerdefinierter Eigenschaften aus dem Element und speichern Sie Änderungen an den benutzerdefinierten Eigenschaftensatz zurück an den Server verwendet werden.</span><span class="sxs-lookup"><span data-stu-id="a7c7c-1263">This object can be used to get, set, and remove custom properties from the item and save changes to the custom property set back to the server.</span></span>|
|`userContext`|<span data-ttu-id="a7c7c-1264">Objekt</span><span class="sxs-lookup"><span data-stu-id="a7c7c-1264">Object</span></span>|<span data-ttu-id="a7c7c-1265">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="a7c7c-1265">&lt;optional&gt;</span></span>|<span data-ttu-id="a7c7c-1266">Entwickler können ein beliebiges Objekt bereitstellen, den sie in der Rückruffunktion zugreifen möchten.</span><span class="sxs-lookup"><span data-stu-id="a7c7c-1266">Developers can provide any object they wish to access in the callback function.</span></span> <span data-ttu-id="a7c7c-1267">Dieses Objekt zugegriffen werden kann, indem die `asyncResult.asyncContext` -Eigenschaft in der Callback-Funktion.</span><span class="sxs-lookup"><span data-stu-id="a7c7c-1267">This object can be accessed by the `asyncResult.asyncContext` property in the callback function.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="a7c7c-1268">Anforderungen</span><span class="sxs-lookup"><span data-stu-id="a7c7c-1268">Requirements</span></span>

|<span data-ttu-id="a7c7c-1269">Anforderung</span><span class="sxs-lookup"><span data-stu-id="a7c7c-1269">Requirement</span></span>|<span data-ttu-id="a7c7c-1270">Wert</span><span class="sxs-lookup"><span data-stu-id="a7c7c-1270">Value</span></span>|
|---|---|
|[<span data-ttu-id="a7c7c-1271">Mindestversion des Postfachanforderungssatzes</span><span class="sxs-lookup"><span data-stu-id="a7c7c-1271">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="a7c7c-1272">1.0</span><span class="sxs-lookup"><span data-stu-id="a7c7c-1272">1.0</span></span>|
|[<span data-ttu-id="a7c7c-1273">Mindestberechtigungsstufe</span><span class="sxs-lookup"><span data-stu-id="a7c7c-1273">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="a7c7c-1274">ReadItem</span><span class="sxs-lookup"><span data-stu-id="a7c7c-1274">ReadItem</span></span>|
|[<span data-ttu-id="a7c7c-1275">Zutreffender Outlook-Modus</span><span class="sxs-lookup"><span data-stu-id="a7c7c-1275">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="a7c7c-1276">Verfassen oder Lesen</span><span class="sxs-lookup"><span data-stu-id="a7c7c-1276">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="a7c7c-1277">Beispiel</span><span class="sxs-lookup"><span data-stu-id="a7c7c-1277">Example</span></span>

<span data-ttu-id="a7c7c-p177">Das folgende Codebeispiel veranschaulicht die Verwendung der `loadCustomPropertiesAsync`-Methode zum asynchronen Laden der benutzerdefinierten Eigenschaften, die für das aktuelle Element spezifisch sind. Des Weiteren veranschaulicht das Beispiel die Verwendung der `CustomProperties.saveAsync`-Methode zum Speichern der Eigenschaften auf dem Server. Nach dem Laden der benutzerdefinierten Eigenschaften wird in dem Codebeispiel die `CustomProperties.get`-Methode dazu verwendet, die benutzerdefinierte `myProp`-Eigenschaft zu lesen. Die `CustomProperties.set`-Methode wird dazu verwendet, die benutzerdefinierte `otherProp`-Eigenschaft zu schreiben. Schließlich wird die `saveAsync`-Methode aufgerufen, um die benutzerdefinierten Eigenschaften zu speichern.</span><span class="sxs-lookup"><span data-stu-id="a7c7c-p177">The following code example shows how to use the `loadCustomPropertiesAsync` method to asynchronously load custom properties that are specific to the current item. The example also shows how to use the `CustomProperties.saveAsync` method to save these properties back to the server. After loading the custom properties, the code sample uses the `CustomProperties.get` method to read the custom property `myProp`, the `CustomProperties.set` method to write the custom property `otherProp`, and then finally calls the `saveAsync` method to save the custom properties.</span></span>

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

####  <a name="removeattachmentasyncattachmentid-options-callback"></a><span data-ttu-id="a7c7c-1281">removeAttachmentAsync(attachmentId, [options], [callback])</span><span class="sxs-lookup"><span data-stu-id="a7c7c-1281">removeAttachmentAsync(attachmentId, [options], [callback])</span></span>

<span data-ttu-id="a7c7c-1282">Entfernt eine Anlage aus einer Nachricht oder einem Termin.</span><span class="sxs-lookup"><span data-stu-id="a7c7c-1282">Removes an attachment from a message or appointment.</span></span>

<span data-ttu-id="a7c7c-p178">Die `removeAttachmentAsync`-Methode entfernt die Anlage mit dem angegebenen Bezeichner aus dem Element. Als bewährte Vorgehensweise sollten Sie den Anlagenbezeichner nur dann zum Entfernen einer Anlage verwenden, wenn die gleiche Mail-App die Anlage in der gleichen Sitzung hinzugefügt hat. In Outlook Web App und OWA für Geräte ist der Anlagenbezeichner nur innerhalb der gleichen Sitzung gültig. Eine Sitzung ist abgeschlossen, wenn der Benutzer die App schließt, oder wenn der Benutzer in einem eingebetteten Formular mit dem Verfassen beginnt und den Vorgang dann in einem separaten Fenster fortsetzt.</span><span class="sxs-lookup"><span data-stu-id="a7c7c-p178">The `removeAttachmentAsync` method removes the attachment with the specified identifier from the item. As a best practice, you should use the attachment identifier to remove an attachment only if the same mail app has added that attachment in the same session. In Outlook Web App and OWA for Devices, the attachment identifier is valid only within the same session. A session is over when the user closes the app, or if the user starts composing in an inline form and subsequently pops out the inline form to continue in a separate window.</span></span>

##### <a name="parameters"></a><span data-ttu-id="a7c7c-1287">Parameter:</span><span class="sxs-lookup"><span data-stu-id="a7c7c-1287">Parameters:</span></span>

|<span data-ttu-id="a7c7c-1288">Name</span><span class="sxs-lookup"><span data-stu-id="a7c7c-1288">Name</span></span>|<span data-ttu-id="a7c7c-1289">Typ</span><span class="sxs-lookup"><span data-stu-id="a7c7c-1289">Type</span></span>|<span data-ttu-id="a7c7c-1290">Attribute</span><span class="sxs-lookup"><span data-stu-id="a7c7c-1290">Attributes</span></span>|<span data-ttu-id="a7c7c-1291">Beschreibung</span><span class="sxs-lookup"><span data-stu-id="a7c7c-1291">Description</span></span>|
|---|---|---|---|
|`attachmentId`|<span data-ttu-id="a7c7c-1292">Zeichenfolge</span><span class="sxs-lookup"><span data-stu-id="a7c7c-1292">String</span></span>||<span data-ttu-id="a7c7c-p179">Der Bezeichner des Anhangs, der entfernt werden soll. Die maximale Länge der Zeichenfolge ist 100 Zeichen.</span><span class="sxs-lookup"><span data-stu-id="a7c7c-p179">The identifier of the attachment to remove. The maximum length of the string is 100 characters.</span></span>|
|`options`|<span data-ttu-id="a7c7c-1295">Object</span><span class="sxs-lookup"><span data-stu-id="a7c7c-1295">Object</span></span>|<span data-ttu-id="a7c7c-1296">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="a7c7c-1296">&lt;optional&gt;</span></span>|<span data-ttu-id="a7c7c-1297">Ein Objektliteral, das mindestens eine der folgenden Eigenschaften enthält.</span><span class="sxs-lookup"><span data-stu-id="a7c7c-1297">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`|<span data-ttu-id="a7c7c-1298">Object</span><span class="sxs-lookup"><span data-stu-id="a7c7c-1298">Object</span></span>|<span data-ttu-id="a7c7c-1299">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="a7c7c-1299">&lt;optional&gt;</span></span>|<span data-ttu-id="a7c7c-1300">Entwickler können ein Objekt bereitstellen, auf das sie in der Callbackmethode zugreifen möchten.</span><span class="sxs-lookup"><span data-stu-id="a7c7c-1300">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`|<span data-ttu-id="a7c7c-1301">Funktion</span><span class="sxs-lookup"><span data-stu-id="a7c7c-1301">function</span></span>|<span data-ttu-id="a7c7c-1302">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="a7c7c-1302">&lt;optional&gt;</span></span>|<span data-ttu-id="a7c7c-1303">Wenn die Methode abgeschlossen wird, wird die Funktion , die im `callback`-Parameter übergeben wird, mit einem einzelnen Parameter (`asyncResult`) aufgerufen, der ein [`AsyncResult`](/javascript/api/office/office.asyncresult)-Objekt darstellt.</span><span class="sxs-lookup"><span data-stu-id="a7c7c-1303">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span> <br/><span data-ttu-id="a7c7c-1304">Wenn beim Entfernen der Anlage ein Fehler auftritt, enthält die Eigenschaft `asyncResult.error` einen Fehlercode mit dem Grund für den Fehler.</span><span class="sxs-lookup"><span data-stu-id="a7c7c-1304">If removing the attachment fails, the `asyncResult.error` property will contain an error code with the reason for the failure.</span></span>|

##### <a name="errors"></a><span data-ttu-id="a7c7c-1305">Fehler</span><span class="sxs-lookup"><span data-stu-id="a7c7c-1305">Errors</span></span>

|<span data-ttu-id="a7c7c-1306">Fehlercode</span><span class="sxs-lookup"><span data-stu-id="a7c7c-1306">Error code</span></span>|<span data-ttu-id="a7c7c-1307">Beschreibung</span><span class="sxs-lookup"><span data-stu-id="a7c7c-1307">Description</span></span>|
|------------|-------------|
|`InvalidAttachmentId`|<span data-ttu-id="a7c7c-1308">Der Bezeichner für die Anlage ist nicht vorhanden.</span><span class="sxs-lookup"><span data-stu-id="a7c7c-1308">The attachment identifier does not exist.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="a7c7c-1309">Anforderungen</span><span class="sxs-lookup"><span data-stu-id="a7c7c-1309">Requirements</span></span>

|<span data-ttu-id="a7c7c-1310">Anforderung</span><span class="sxs-lookup"><span data-stu-id="a7c7c-1310">Requirement</span></span>|<span data-ttu-id="a7c7c-1311">Wert</span><span class="sxs-lookup"><span data-stu-id="a7c7c-1311">Value</span></span>|
|---|---|
|[<span data-ttu-id="a7c7c-1312">Mindestversion des Postfachanforderungssatzes</span><span class="sxs-lookup"><span data-stu-id="a7c7c-1312">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="a7c7c-1313">1.1</span><span class="sxs-lookup"><span data-stu-id="a7c7c-1313">1.1</span></span>|
|[<span data-ttu-id="a7c7c-1314">Mindestberechtigungsstufe</span><span class="sxs-lookup"><span data-stu-id="a7c7c-1314">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="a7c7c-1315">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="a7c7c-1315">ReadWriteItem</span></span>|
|[<span data-ttu-id="a7c7c-1316">Zutreffender Outlook-Modus</span><span class="sxs-lookup"><span data-stu-id="a7c7c-1316">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="a7c7c-1317">Verfassen</span><span class="sxs-lookup"><span data-stu-id="a7c7c-1317">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="a7c7c-1318">Beispiel</span><span class="sxs-lookup"><span data-stu-id="a7c7c-1318">Example</span></span>

<span data-ttu-id="a7c7c-1319">Mit dem folgende Code wird eine Anlage mit dem Bezeichner "0" entfernt.</span><span class="sxs-lookup"><span data-stu-id="a7c7c-1319">The following code removes an attachment with an identifier of '0'.</span></span>

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

####  <a name="removehandlerasynceventtype-handler-options-callback"></a><span data-ttu-id="a7c7c-1320">RemoveHandlerAsync (EventType, Handler, [Optionen], [Rückruf])</span><span class="sxs-lookup"><span data-stu-id="a7c7c-1320">removeHandlerAsync(eventType, handler, [options], [callback])</span></span>

<span data-ttu-id="a7c7c-1321">Entfernt einen Ereignishandler für ein Ereignis unterstützt.</span><span class="sxs-lookup"><span data-stu-id="a7c7c-1321">Removes an event handler for a supported event.</span></span>

<span data-ttu-id="a7c7c-1322">Derzeit sind die unterstützten Ereignistypen `Office.EventType.AppointmentTimeChanged`, `Office.EventType.RecipientsChanged`, und`Office.EventType.RecurrenceChanged`</span><span class="sxs-lookup"><span data-stu-id="a7c7c-1322">Currently the supported event types are `Office.EventType.AppointmentTimeChanged`, `Office.EventType.RecipientsChanged`, and `Office.EventType.RecurrenceChanged`</span></span>

##### <a name="parameters"></a><span data-ttu-id="a7c7c-1323">Parameter:</span><span class="sxs-lookup"><span data-stu-id="a7c7c-1323">Parameters:</span></span>

| <span data-ttu-id="a7c7c-1324">Name</span><span class="sxs-lookup"><span data-stu-id="a7c7c-1324">Name</span></span> | <span data-ttu-id="a7c7c-1325">Typ</span><span class="sxs-lookup"><span data-stu-id="a7c7c-1325">Type</span></span> | <span data-ttu-id="a7c7c-1326">Attribute</span><span class="sxs-lookup"><span data-stu-id="a7c7c-1326">Attributes</span></span> | <span data-ttu-id="a7c7c-1327">Beschreibung</span><span class="sxs-lookup"><span data-stu-id="a7c7c-1327">Description</span></span> |
|---|---|---|---|
| `eventType` | [<span data-ttu-id="a7c7c-1328">Office.EventType</span><span class="sxs-lookup"><span data-stu-id="a7c7c-1328">Office.EventType</span></span>](office.md#eventtype-string) || <span data-ttu-id="a7c7c-1329">Das Ereignis, das den Handler aufrufen soll</span><span class="sxs-lookup"><span data-stu-id="a7c7c-1329">The event that should invoke the handler.</span></span> |
| `handler` | <span data-ttu-id="a7c7c-1330">Function</span><span class="sxs-lookup"><span data-stu-id="a7c7c-1330">Function</span></span> || <span data-ttu-id="a7c7c-p180">Die Funktion, die das Ereignis behandeln soll. Die Funktion muss einen einzigen Parameter akzeptieren (ein Objektliteral). Die Eigenschaft `type` dieses Parameters entspricht dem Parameter `eventType`, der an `removeHandlerAsync` übergeben wird.</span><span class="sxs-lookup"><span data-stu-id="a7c7c-p180">The function to handle the event. The function must accept a single parameter, which is an object literal. The `type` property on the parameter will match the `eventType` parameter passed to `removeHandlerAsync`.</span></span> |
| `options` | <span data-ttu-id="a7c7c-1334">Objekt</span><span class="sxs-lookup"><span data-stu-id="a7c7c-1334">Object</span></span> | <span data-ttu-id="a7c7c-1335">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="a7c7c-1335">&lt;optional&gt;</span></span> | <span data-ttu-id="a7c7c-1336">Ein Objektliteral, das mindestens eine der folgenden Eigenschaften enthält.</span><span class="sxs-lookup"><span data-stu-id="a7c7c-1336">An object literal that contains one or more of the following properties.</span></span> |
| `options.asyncContext` | <span data-ttu-id="a7c7c-1337">Object</span><span class="sxs-lookup"><span data-stu-id="a7c7c-1337">Object</span></span> | <span data-ttu-id="a7c7c-1338">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="a7c7c-1338">&lt;optional&gt;</span></span> | <span data-ttu-id="a7c7c-1339">Entwickler können ein Objekt bereitstellen, auf das sie in der Callbackmethode zugreifen möchten.</span><span class="sxs-lookup"><span data-stu-id="a7c7c-1339">Developers can provide any object they wish to access in the callback method.</span></span> |
| `callback` | <span data-ttu-id="a7c7c-1340">Funktion</span><span class="sxs-lookup"><span data-stu-id="a7c7c-1340">function</span></span>| <span data-ttu-id="a7c7c-1341">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="a7c7c-1341">&lt;optional&gt;</span></span>|<span data-ttu-id="a7c7c-1342">Wenn die Methode abgeschlossen wird, wird die Funktion , die im `callback`-Parameter übergeben wird, mit einem einzelnen Parameter (`asyncResult`) aufgerufen, der ein [`AsyncResult`](/javascript/api/office/office.asyncresult)-Objekt darstellt.</span><span class="sxs-lookup"><span data-stu-id="a7c7c-1342">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="a7c7c-1343">Anforderungen</span><span class="sxs-lookup"><span data-stu-id="a7c7c-1343">Requirements</span></span>

|<span data-ttu-id="a7c7c-1344">Anforderung</span><span class="sxs-lookup"><span data-stu-id="a7c7c-1344">Requirement</span></span>| <span data-ttu-id="a7c7c-1345">Wert</span><span class="sxs-lookup"><span data-stu-id="a7c7c-1345">Value</span></span>|
|---|---|
|[<span data-ttu-id="a7c7c-1346">Mindestversion des Postfachanforderungssatzes</span><span class="sxs-lookup"><span data-stu-id="a7c7c-1346">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="a7c7c-1347">1.7</span><span class="sxs-lookup"><span data-stu-id="a7c7c-1347">1.7</span></span> |
|[<span data-ttu-id="a7c7c-1348">Mindestberechtigungsstufe</span><span class="sxs-lookup"><span data-stu-id="a7c7c-1348">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="a7c7c-1349">ReadItem</span><span class="sxs-lookup"><span data-stu-id="a7c7c-1349">ReadItem</span></span> |
|[<span data-ttu-id="a7c7c-1350">Zutreffender Outlook-Modus</span><span class="sxs-lookup"><span data-stu-id="a7c7c-1350">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="a7c7c-1351">Verfassen oder Lesen</span><span class="sxs-lookup"><span data-stu-id="a7c7c-1351">Compose or read</span></span> |

####  <a name="saveasyncoptions-callback"></a><span data-ttu-id="a7c7c-1352">saveAsync([options], callback)</span><span class="sxs-lookup"><span data-stu-id="a7c7c-1352">saveAsync([options], callback)</span></span>

<span data-ttu-id="a7c7c-1353">Speicher asynchron ein Element. .</span><span class="sxs-lookup"><span data-stu-id="a7c7c-1353">Asynchronously saves an item.</span></span>

<span data-ttu-id="a7c7c-p181">Beim Aufrufen speichert diese Methode die aktuelle Nachricht als Entwurf und  gibt die Element-ID über die Callbackmethode zurück. In Outlook Web App oder Outlook im Onlinemodus wird das Element auf dem Server gespeichert. In Outlook im Cache-Modus wird das Element im lokalen Cache gespeichert.</span><span class="sxs-lookup"><span data-stu-id="a7c7c-p181">When invoked, this method saves the current message as a draft and returns the item id via the callback method. In Outlook Web App or Outlook in online mode, the item is saved to the server. In Outlook in cached mode, the item is saved to the local cache.</span></span>

> [!NOTE]
> <span data-ttu-id="a7c7c-1357">Wenn Ihr Add-In ruft `saveAsync` für ein Element im Entwurfsmodus, um das Abrufen einer `itemId` um mit EWS oder REST-API verwenden, beachten Sie, dass wenn Outlook im Cache-Modus ist, es kann einige Zeit dauern, bevor das Element tatsächlich auf dem Server synchronisiert wird.</span><span class="sxs-lookup"><span data-stu-id="a7c7c-1357">If your add-in calls `saveAsync` on an item in compose mode in order to get an `itemId` to use with EWS or the REST API, be aware that when Outlook is in cached mode, it may take some time before the item is actually synced to the server.</span></span> <span data-ttu-id="a7c7c-1358">Bis das Element mit synchronisiert ist, die `itemId` wird ein Fehler zurückgegeben.</span><span class="sxs-lookup"><span data-stu-id="a7c7c-1358">Until the item is synced, using the `itemId` will return an error.</span></span>

<span data-ttu-id="a7c7c-p183">Da Termine keinen Entwurfsstatus haben, wird das Element bei Aufruf von `saveAsync` für einen Termin im Verfassenmodus als normaler Termin im Kalender des Benutzers gespeichert. Für neue Termine, die noch nicht gespeichert wurden, wird keine Einladung gesendet. Beim Speichern eines vorhandenen Termins wird eine Aktualisierung an die hinzugefügten oder entfernten Teilnehmer gesendet.</span><span class="sxs-lookup"><span data-stu-id="a7c7c-p183">Since appointments have no draft state, if `saveAsync` is called on an appointment in compose mode, the item will be saved as a normal appointment on the user's calendar. For new appointments that have not been saved before, no invitation will be sent. Saving an existing appointment will send an update to added or removed attendees.</span></span>

> [!NOTE]
> <span data-ttu-id="a7c7c-1362">Die folgenden Clients haben unterschiedlichem Verhalten für `saveAsync` Termine im Verfassenmodus:</span><span class="sxs-lookup"><span data-stu-id="a7c7c-1362">The following clients have different behavior for `saveAsync` on appointments in compose mode:</span></span>
>
> - <span data-ttu-id="a7c7c-1363">Outlook für Mac unterstützt keine `saveAsync` auf einer Besprechung im Verfassenmodus.</span><span class="sxs-lookup"><span data-stu-id="a7c7c-1363">Mac Outlook does not support `saveAsync` on a meeting in compose mode.</span></span> <span data-ttu-id="a7c7c-1364">Aufrufen von `saveAsync` auf einer Besprechung in Outlook für Mac wird ein Fehler zurückgegeben.</span><span class="sxs-lookup"><span data-stu-id="a7c7c-1364">Calling `saveAsync` on a meeting in Mac Outlook will return an error.</span></span>
> - <span data-ttu-id="a7c7c-1365">Outlook im Web immer sendet eine Einladung oder zu aktualisieren, wenn `saveAsync` für einen Termin aufgerufen wird, im Verfassenmodus.</span><span class="sxs-lookup"><span data-stu-id="a7c7c-1365">Outlook on the web always sends an invitation or update when `saveAsync` is called on an appointment in compose mode.</span></span>

##### <a name="parameters"></a><span data-ttu-id="a7c7c-1366">Parameter:</span><span class="sxs-lookup"><span data-stu-id="a7c7c-1366">Parameters:</span></span>

|<span data-ttu-id="a7c7c-1367">Name</span><span class="sxs-lookup"><span data-stu-id="a7c7c-1367">Name</span></span>|<span data-ttu-id="a7c7c-1368">Typ</span><span class="sxs-lookup"><span data-stu-id="a7c7c-1368">Type</span></span>|<span data-ttu-id="a7c7c-1369">Attribute</span><span class="sxs-lookup"><span data-stu-id="a7c7c-1369">Attributes</span></span>|<span data-ttu-id="a7c7c-1370">Beschreibung</span><span class="sxs-lookup"><span data-stu-id="a7c7c-1370">Description</span></span>|
|---|---|---|---|
|`options`|<span data-ttu-id="a7c7c-1371">Objekt</span><span class="sxs-lookup"><span data-stu-id="a7c7c-1371">Object</span></span>|<span data-ttu-id="a7c7c-1372">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="a7c7c-1372">&lt;optional&gt;</span></span>|<span data-ttu-id="a7c7c-1373">Ein Objektliteral, das mindestens eine der folgenden Eigenschaften enthält.</span><span class="sxs-lookup"><span data-stu-id="a7c7c-1373">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`|<span data-ttu-id="a7c7c-1374">Object</span><span class="sxs-lookup"><span data-stu-id="a7c7c-1374">Object</span></span>|<span data-ttu-id="a7c7c-1375">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="a7c7c-1375">&lt;optional&gt;</span></span>|<span data-ttu-id="a7c7c-1376">Entwickler können ein Objekt bereitstellen, auf das sie in der Callbackmethode zugreifen möchten.</span><span class="sxs-lookup"><span data-stu-id="a7c7c-1376">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`|<span data-ttu-id="a7c7c-1377">Funktion</span><span class="sxs-lookup"><span data-stu-id="a7c7c-1377">function</span></span>||<span data-ttu-id="a7c7c-1378">Wenn die Methode abgeschlossen wird, wird die Funktion , die im `callback`-Parameter übergeben wird, mit einem einzelnen Parameter (`asyncResult`) aufgerufen, der ein [`AsyncResult`](/javascript/api/office/office.asyncresult)-Objekt darstellt.</span><span class="sxs-lookup"><span data-stu-id="a7c7c-1378">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span><br/><br/><span data-ttu-id="a7c7c-1379">Auf Erfolg, erfolgt die Element-ID der `asyncResult.value` Eigenschaft.</span><span class="sxs-lookup"><span data-stu-id="a7c7c-1379">On success, the item identifier is provided in the `asyncResult.value` property.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="a7c7c-1380">Anforderungen</span><span class="sxs-lookup"><span data-stu-id="a7c7c-1380">Requirements</span></span>

|<span data-ttu-id="a7c7c-1381">Anforderung</span><span class="sxs-lookup"><span data-stu-id="a7c7c-1381">Requirement</span></span>|<span data-ttu-id="a7c7c-1382">Wert</span><span class="sxs-lookup"><span data-stu-id="a7c7c-1382">Value</span></span>|
|---|---|
|[<span data-ttu-id="a7c7c-1383">Mindestversion des Postfachanforderungssatzes</span><span class="sxs-lookup"><span data-stu-id="a7c7c-1383">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="a7c7c-1384">1.3</span><span class="sxs-lookup"><span data-stu-id="a7c7c-1384">1.3</span></span>|
|[<span data-ttu-id="a7c7c-1385">Mindestberechtigungsstufe</span><span class="sxs-lookup"><span data-stu-id="a7c7c-1385">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="a7c7c-1386">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="a7c7c-1386">ReadWriteItem</span></span>|
|[<span data-ttu-id="a7c7c-1387">Zutreffender Outlook-Modus</span><span class="sxs-lookup"><span data-stu-id="a7c7c-1387">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="a7c7c-1388">Verfassen</span><span class="sxs-lookup"><span data-stu-id="a7c7c-1388">Compose</span></span>|

##### <a name="examples"></a><span data-ttu-id="a7c7c-1389">Beispiele</span><span class="sxs-lookup"><span data-stu-id="a7c7c-1389">Examples</span></span>

```
Office.context.mailbox.item.saveAsync(
  function callback(result) {
    // Process the result
  });
```

<span data-ttu-id="a7c7c-p185">Es folgt ein Beispiel für den `result`-Parameter, der an die Callbackfunktion übergeben wird. Die `value`-Eigenschaft enthält die Element-ID des Elements.</span><span class="sxs-lookup"><span data-stu-id="a7c7c-p185">The following is an example of the `result` parameter passed to the callback function. The `value` property contains the item ID of the item.</span></span>

```
{
  "value":"AAMkADI5...AAA=",
  "status":"succeeded"
}
```

####  <a name="setselecteddataasyncdata-options-callback"></a><span data-ttu-id="a7c7c-1392">setSelectedDataAsync(data, [options], callback)</span><span class="sxs-lookup"><span data-stu-id="a7c7c-1392">setSelectedDataAsync(data, [options], callback)</span></span>

<span data-ttu-id="a7c7c-1393">Fügt asynchron Daten in den Textkörper oder Betreff einer Nachricht ein.</span><span class="sxs-lookup"><span data-stu-id="a7c7c-1393">Asynchronously inserts data into the body or subject of a message.</span></span>

<span data-ttu-id="a7c7c-p186">Die `setSelectedDataAsync`-Methode fügt die angegebene Zeichenfolge an der Cursorposition im Betreff oder im Textkörper ein, oder, falls im Editor Text ausgewählt ist, ersetzt den markierten Text. Wenn sich der Cursor nicht im Textkörper oder im Betreffsfeld befindet, wird ein Fehler zurückgegeben. Nach dem Einfügen wird der Cursor am Ende der eingefügten Inhalte platziert.</span><span class="sxs-lookup"><span data-stu-id="a7c7c-p186">The `setSelectedDataAsync` method inserts the specified string at the cursor location in the subject or body of the item, or, if text is selected in the editor, it replaces the selected text. If the cursor is not in the body or subject field, an error is returned. After insertion, the cursor is placed at the end of the inserted content.</span></span>

##### <a name="parameters"></a><span data-ttu-id="a7c7c-1397">Parameter:</span><span class="sxs-lookup"><span data-stu-id="a7c7c-1397">Parameters:</span></span>

|<span data-ttu-id="a7c7c-1398">Name</span><span class="sxs-lookup"><span data-stu-id="a7c7c-1398">Name</span></span>|<span data-ttu-id="a7c7c-1399">Typ</span><span class="sxs-lookup"><span data-stu-id="a7c7c-1399">Type</span></span>|<span data-ttu-id="a7c7c-1400">Attribute</span><span class="sxs-lookup"><span data-stu-id="a7c7c-1400">Attributes</span></span>|<span data-ttu-id="a7c7c-1401">Beschreibung</span><span class="sxs-lookup"><span data-stu-id="a7c7c-1401">Description</span></span>|
|---|---|---|---|
|`data`|<span data-ttu-id="a7c7c-1402">Zeichenfolge</span><span class="sxs-lookup"><span data-stu-id="a7c7c-1402">String</span></span>||<span data-ttu-id="a7c7c-p187">Die einzufügenden Daten. Daten dürfen 1.000.000 Zeichen nicht überschreiten. Werden mehr als 1.000.000 Zeichen übergeben, wird eine `ArgumentOutOfRange`-Ausnahme ausgelöst.</span><span class="sxs-lookup"><span data-stu-id="a7c7c-p187">The data to be inserted. Data is not to exceed 1,000,000 characters. If more than 1,000,000 characters are passed in, an `ArgumentOutOfRange` exception is thrown.</span></span>|
|`options`|<span data-ttu-id="a7c7c-1406">Object</span><span class="sxs-lookup"><span data-stu-id="a7c7c-1406">Object</span></span>|<span data-ttu-id="a7c7c-1407">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="a7c7c-1407">&lt;optional&gt;</span></span>|<span data-ttu-id="a7c7c-1408">Ein Objektliteral, das mindestens eine der folgenden Eigenschaften enthält.</span><span class="sxs-lookup"><span data-stu-id="a7c7c-1408">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`|<span data-ttu-id="a7c7c-1409">Object</span><span class="sxs-lookup"><span data-stu-id="a7c7c-1409">Object</span></span>|<span data-ttu-id="a7c7c-1410">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="a7c7c-1410">&lt;optional&gt;</span></span>|<span data-ttu-id="a7c7c-1411">Entwickler können ein Objekt bereitstellen, auf das sie in der Rückrufmethode zugreifen möchten.</span><span class="sxs-lookup"><span data-stu-id="a7c7c-1411">Developers can provide any object they wish to access in the callback method.</span></span>|
|`options.coercionType`|[<span data-ttu-id="a7c7c-1412">Office.CoercionType</span><span class="sxs-lookup"><span data-stu-id="a7c7c-1412">Office.CoercionType</span></span>](office.md#coerciontype-string)|<span data-ttu-id="a7c7c-1413">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="a7c7c-1413">&lt;optional&gt;</span></span>|<span data-ttu-id="a7c7c-p188">Wenn `text`, wird das aktuelle Format in Outlook Web App und Outlook angewendet. Wenn das Feld ein HTML-Editor ist, werden nur die Textdaten eingefügt, selbst wenn es sich bei den Daten um HTML-Daten handelt.</span><span class="sxs-lookup"><span data-stu-id="a7c7c-p188">If `text`, the current style is applied in Outlook Web App and Outlook. If the field is an HTML editor, only the text data is inserted, even if the data is HTML.</span></span><br/><br/><span data-ttu-id="a7c7c-p189">Wenn `html` gesetzt ist und das Feld HTML unterstützt (der Betreff jedoch nicht), wird die aktuelle Formatvorlage in Outlook Web App angewendet und die Standardformatvorlage in Outlook. Ist das Feld ein Textfeld, wird ein Fehler des Typs `InvalidDataFormat` zurückgegeben.</span><span class="sxs-lookup"><span data-stu-id="a7c7c-p189">If `html` and the field supports HTML (the subject doesn't), the current style is applied in Outlook Web App and the default style is applied in Outlook. If the field is a text field, an `InvalidDataFormat` error is returned.</span></span><br/><br/><span data-ttu-id="a7c7c-1418">Wenn `coercionType` nicht festgelegt wird, hängt das Ergebnis vom Feld ab: Wenn das Feld HTML ist, wird HTML verwendet. Wenn das Feld Text ist, wird Nur-Text verwendet.</span><span class="sxs-lookup"><span data-stu-id="a7c7c-1418">If `coercionType` is not set, the result depends on the field: if the field is HTML then HTML is used; if the field is text, then plain text is used.</span></span>|
|`callback`|<span data-ttu-id="a7c7c-1419">function</span><span class="sxs-lookup"><span data-stu-id="a7c7c-1419">function</span></span>||<span data-ttu-id="a7c7c-1420">Wenn die Methode abgeschlossen wird, wird die Funktion , die im `callback`-Parameter übergeben wird, mit einem einzelnen Parameter (`asyncResult`) aufgerufen, der ein [`AsyncResult`](/javascript/api/office/office.asyncresult)-Objekt darstellt.</span><span class="sxs-lookup"><span data-stu-id="a7c7c-1420">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="a7c7c-1421">Anforderungen</span><span class="sxs-lookup"><span data-stu-id="a7c7c-1421">Requirements</span></span>

|<span data-ttu-id="a7c7c-1422">Anforderung</span><span class="sxs-lookup"><span data-stu-id="a7c7c-1422">Requirement</span></span>|<span data-ttu-id="a7c7c-1423">Wert</span><span class="sxs-lookup"><span data-stu-id="a7c7c-1423">Value</span></span>|
|---|---|
|[<span data-ttu-id="a7c7c-1424">Mindestversion des Postfachanforderungssatzes</span><span class="sxs-lookup"><span data-stu-id="a7c7c-1424">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="a7c7c-1425">1.2</span><span class="sxs-lookup"><span data-stu-id="a7c7c-1425">1.2</span></span>|
|[<span data-ttu-id="a7c7c-1426">Mindestberechtigungsstufe</span><span class="sxs-lookup"><span data-stu-id="a7c7c-1426">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="a7c7c-1427">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="a7c7c-1427">ReadWriteItem</span></span>|
|[<span data-ttu-id="a7c7c-1428">Zutreffender Outlook-Modus</span><span class="sxs-lookup"><span data-stu-id="a7c7c-1428">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="a7c7c-1429">Verfassen</span><span class="sxs-lookup"><span data-stu-id="a7c7c-1429">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="a7c7c-1430">Beispiel</span><span class="sxs-lookup"><span data-stu-id="a7c7c-1430">Example</span></span>

```
Office.context.mailbox.item.setSelectedDataAsync("Hello World!");
Office.context.mailbox.item.setSelectedDataAsync("<b>Hello World!</b>", { coercionType : "html" });
```