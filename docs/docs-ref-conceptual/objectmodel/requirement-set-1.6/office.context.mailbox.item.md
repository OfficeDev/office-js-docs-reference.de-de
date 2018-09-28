
# <a name="item"></a><span data-ttu-id="bac6a-101">item</span><span class="sxs-lookup"><span data-stu-id="bac6a-101">item</span></span>

### <a name="officeofficemdcontextofficecontextmdmailboxofficecontextmailboxmditem"></a><span data-ttu-id="bac6a-102">[Office](office.md)[.context](office.context.md)[.mailbox](office.context.mailbox.md).item</span><span class="sxs-lookup"><span data-stu-id="bac6a-102">[Office](office.md)[.context](office.context.md)[.mailbox](office.context.mailbox.md).item</span></span>

<span data-ttu-id="bac6a-p101">Der `item`-Namespace wird für den Zugriff auf die aktuell ausgewählte Nachricht, Besprechungsanfrage oder den aktuell ausgewählten Termin verwendet. Sie können den Typ von `item` mithilfe der [itemType](#itemtype-officemailboxenumsitemtypejavascriptapioutlook16officemailboxenumsitemtype)-Eigenschaft bestimmen.</span><span class="sxs-lookup"><span data-stu-id="bac6a-p101">The `item` namespace is used to access the currently selected message, meeting request, or appointment. You can determine the type of the `item` by using the [itemType](#itemtype-officemailboxenumsitemtypejavascriptapioutlook16officemailboxenumsitemtype) property.</span></span>

##### <a name="requirements"></a><span data-ttu-id="bac6a-105">Anforderungen</span><span class="sxs-lookup"><span data-stu-id="bac6a-105">Requirements</span></span>

|<span data-ttu-id="bac6a-106">Anforderung</span><span class="sxs-lookup"><span data-stu-id="bac6a-106">Requirement</span></span>| <span data-ttu-id="bac6a-107">Wert</span><span class="sxs-lookup"><span data-stu-id="bac6a-107">Value</span></span>|
|---|---|
|[<span data-ttu-id="bac6a-108">Mindestversion des Postfachanforderungssatzes</span><span class="sxs-lookup"><span data-stu-id="bac6a-108">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="bac6a-109">1.0</span><span class="sxs-lookup"><span data-stu-id="bac6a-109">1.0</span></span>|
|[<span data-ttu-id="bac6a-110">Mindestberechtigungsstufe</span><span class="sxs-lookup"><span data-stu-id="bac6a-110">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="bac6a-111">Eingeschränkt</span><span class="sxs-lookup"><span data-stu-id="bac6a-111">Restricted</span></span>|
|[<span data-ttu-id="bac6a-112">Zutreffender Outlook-Modus</span><span class="sxs-lookup"><span data-stu-id="bac6a-112">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="bac6a-113">Verfassen oder Lesen</span><span class="sxs-lookup"><span data-stu-id="bac6a-113">Compose or read</span></span>|

##### <a name="members-and-methods"></a><span data-ttu-id="bac6a-114">Elemente und Methoden</span><span class="sxs-lookup"><span data-stu-id="bac6a-114">Members and methods</span></span>

| <span data-ttu-id="bac6a-115">Element</span><span class="sxs-lookup"><span data-stu-id="bac6a-115">Member</span></span> | <span data-ttu-id="bac6a-116">Typ</span><span class="sxs-lookup"><span data-stu-id="bac6a-116">Type</span></span> |
|--------|------|
| [<span data-ttu-id="bac6a-117">attachments</span><span class="sxs-lookup"><span data-stu-id="bac6a-117">attachments</span></span>](#attachments-arrayattachmentdetailsjavascriptapioutlook16officeattachmentdetails) | <span data-ttu-id="bac6a-118">Element</span><span class="sxs-lookup"><span data-stu-id="bac6a-118">Member</span></span> |
| [<span data-ttu-id="bac6a-119">bcc</span><span class="sxs-lookup"><span data-stu-id="bac6a-119">bcc</span></span>](#bcc-recipientsjavascriptapioutlook16officerecipients) | <span data-ttu-id="bac6a-120">Element</span><span class="sxs-lookup"><span data-stu-id="bac6a-120">Member</span></span> |
| [<span data-ttu-id="bac6a-121">body</span><span class="sxs-lookup"><span data-stu-id="bac6a-121">body</span></span>](#body-bodyjavascriptapioutlook16officebody) | <span data-ttu-id="bac6a-122">Element</span><span class="sxs-lookup"><span data-stu-id="bac6a-122">Member</span></span> |
| [<span data-ttu-id="bac6a-123">cc</span><span class="sxs-lookup"><span data-stu-id="bac6a-123">cc</span></span>](#cc-arrayemailaddressdetailsjavascriptapioutlook16officeemailaddressdetailsrecipientsjavascriptapioutlook16officerecipients) | <span data-ttu-id="bac6a-124">Element</span><span class="sxs-lookup"><span data-stu-id="bac6a-124">Member</span></span> |
| [<span data-ttu-id="bac6a-125">conversationId</span><span class="sxs-lookup"><span data-stu-id="bac6a-125">conversationId</span></span>](#nullable-conversationid-string) | <span data-ttu-id="bac6a-126">Element</span><span class="sxs-lookup"><span data-stu-id="bac6a-126">Member</span></span> |
| [<span data-ttu-id="bac6a-127">dateTimeCreated</span><span class="sxs-lookup"><span data-stu-id="bac6a-127">dateTimeCreated</span></span>](#datetimecreated-date) | <span data-ttu-id="bac6a-128">Element</span><span class="sxs-lookup"><span data-stu-id="bac6a-128">Member</span></span> |
| [<span data-ttu-id="bac6a-129">dateTimeModified</span><span class="sxs-lookup"><span data-stu-id="bac6a-129">dateTimeModified</span></span>](#datetimemodified-date) | <span data-ttu-id="bac6a-130">Element</span><span class="sxs-lookup"><span data-stu-id="bac6a-130">Member</span></span> |
| [<span data-ttu-id="bac6a-131">end</span><span class="sxs-lookup"><span data-stu-id="bac6a-131">end</span></span>](#end-datetimejavascriptapioutlook16officetime) | <span data-ttu-id="bac6a-132">Element</span><span class="sxs-lookup"><span data-stu-id="bac6a-132">Member</span></span> |
| [<span data-ttu-id="bac6a-133">from</span><span class="sxs-lookup"><span data-stu-id="bac6a-133">from</span></span>](#from-emailaddressdetailsjavascriptapioutlook16officeemailaddressdetails) | <span data-ttu-id="bac6a-134">Element</span><span class="sxs-lookup"><span data-stu-id="bac6a-134">Member</span></span> |
| [<span data-ttu-id="bac6a-135">internetMessageId</span><span class="sxs-lookup"><span data-stu-id="bac6a-135">internetMessageId</span></span>](#internetmessageid-string) | <span data-ttu-id="bac6a-136">Element</span><span class="sxs-lookup"><span data-stu-id="bac6a-136">Member</span></span> |
| [<span data-ttu-id="bac6a-137">itemClass</span><span class="sxs-lookup"><span data-stu-id="bac6a-137">itemClass</span></span>](#itemclass-string) | <span data-ttu-id="bac6a-138">Element</span><span class="sxs-lookup"><span data-stu-id="bac6a-138">Member</span></span> |
| [<span data-ttu-id="bac6a-139">itemId</span><span class="sxs-lookup"><span data-stu-id="bac6a-139">itemId</span></span>](#nullable-itemid-string) | <span data-ttu-id="bac6a-140">Element</span><span class="sxs-lookup"><span data-stu-id="bac6a-140">Member</span></span> |
| [<span data-ttu-id="bac6a-141">itemType</span><span class="sxs-lookup"><span data-stu-id="bac6a-141">itemType</span></span>](#itemtype-officemailboxenumsitemtypejavascriptapioutlook16officemailboxenumsitemtype) | <span data-ttu-id="bac6a-142">Element</span><span class="sxs-lookup"><span data-stu-id="bac6a-142">Member</span></span> |
| [<span data-ttu-id="bac6a-143">location</span><span class="sxs-lookup"><span data-stu-id="bac6a-143">location</span></span>](#location-stringlocationjavascriptapioutlook16officelocation) | <span data-ttu-id="bac6a-144">Element</span><span class="sxs-lookup"><span data-stu-id="bac6a-144">Member</span></span> |
| [<span data-ttu-id="bac6a-145">normalizedSubject</span><span class="sxs-lookup"><span data-stu-id="bac6a-145">normalizedSubject</span></span>](#normalizedsubject-string) | <span data-ttu-id="bac6a-146">Element</span><span class="sxs-lookup"><span data-stu-id="bac6a-146">Member</span></span> |
| [<span data-ttu-id="bac6a-147">notificationMessages</span><span class="sxs-lookup"><span data-stu-id="bac6a-147">notificationMessages</span></span>](#notificationmessages-notificationmessagesjavascriptapioutlook16officenotificationmessages) | <span data-ttu-id="bac6a-148">Element</span><span class="sxs-lookup"><span data-stu-id="bac6a-148">Member</span></span> |
| [<span data-ttu-id="bac6a-149">optionalAttendees</span><span class="sxs-lookup"><span data-stu-id="bac6a-149">optionalAttendees</span></span>](#optionalattendees-arrayemailaddressdetailsjavascriptapioutlook16officeemailaddressdetailsrecipientsjavascriptapioutlook16officerecipients) | <span data-ttu-id="bac6a-150">Element</span><span class="sxs-lookup"><span data-stu-id="bac6a-150">Member</span></span> |
| [<span data-ttu-id="bac6a-151">organizer</span><span class="sxs-lookup"><span data-stu-id="bac6a-151">organizer</span></span>](#organizer-emailaddressdetailsjavascriptapioutlook16officeemailaddressdetails) | <span data-ttu-id="bac6a-152">Element</span><span class="sxs-lookup"><span data-stu-id="bac6a-152">Member</span></span> |
| [<span data-ttu-id="bac6a-153">requiredAttendees</span><span class="sxs-lookup"><span data-stu-id="bac6a-153">requiredAttendees</span></span>](#requiredattendees-arrayemailaddressdetailsjavascriptapioutlook16officeemailaddressdetailsrecipientsjavascriptapioutlook16officerecipients) | <span data-ttu-id="bac6a-154">Element</span><span class="sxs-lookup"><span data-stu-id="bac6a-154">Member</span></span> |
| [<span data-ttu-id="bac6a-155">sender</span><span class="sxs-lookup"><span data-stu-id="bac6a-155">sender</span></span>](#sender-emailaddressdetailsjavascriptapioutlook16officeemailaddressdetails) | <span data-ttu-id="bac6a-156">Element</span><span class="sxs-lookup"><span data-stu-id="bac6a-156">Member</span></span> |
| [<span data-ttu-id="bac6a-157">start</span><span class="sxs-lookup"><span data-stu-id="bac6a-157">start</span></span>](#start-datetimejavascriptapioutlook16officetime) | <span data-ttu-id="bac6a-158">Element</span><span class="sxs-lookup"><span data-stu-id="bac6a-158">Member</span></span> |
| [<span data-ttu-id="bac6a-159">subject</span><span class="sxs-lookup"><span data-stu-id="bac6a-159">subject</span></span>](#subject-stringsubjectjavascriptapioutlook16officesubject) | <span data-ttu-id="bac6a-160">Element</span><span class="sxs-lookup"><span data-stu-id="bac6a-160">Member</span></span> |
| [<span data-ttu-id="bac6a-161">to</span><span class="sxs-lookup"><span data-stu-id="bac6a-161">to</span></span>](#to-arrayemailaddressdetailsjavascriptapioutlook16officeemailaddressdetailsrecipientsjavascriptapioutlook16officerecipients) | <span data-ttu-id="bac6a-162">Element</span><span class="sxs-lookup"><span data-stu-id="bac6a-162">Member</span></span> |
| [<span data-ttu-id="bac6a-163">addFileAttachmentAsync</span><span class="sxs-lookup"><span data-stu-id="bac6a-163">addFileAttachmentAsync</span></span>](#addfileattachmentasyncuri-attachmentname-options-callback) | <span data-ttu-id="bac6a-164">Methode</span><span class="sxs-lookup"><span data-stu-id="bac6a-164">Method</span></span> |
| [<span data-ttu-id="bac6a-165">addItemAttachmentAsync</span><span class="sxs-lookup"><span data-stu-id="bac6a-165">addItemAttachmentAsync</span></span>](#additemattachmentasyncitemid-attachmentname-options-callback) | <span data-ttu-id="bac6a-166">Methode</span><span class="sxs-lookup"><span data-stu-id="bac6a-166">Method</span></span> |
| [<span data-ttu-id="bac6a-167">close</span><span class="sxs-lookup"><span data-stu-id="bac6a-167">close</span></span>](#close) | <span data-ttu-id="bac6a-168">Methode</span><span class="sxs-lookup"><span data-stu-id="bac6a-168">Method</span></span> |
| [<span data-ttu-id="bac6a-169">displayReplyAllForm</span><span class="sxs-lookup"><span data-stu-id="bac6a-169">displayReplyAllForm</span></span>](#displayreplyallformformdata) | <span data-ttu-id="bac6a-170">Methode</span><span class="sxs-lookup"><span data-stu-id="bac6a-170">Method</span></span> |
| [<span data-ttu-id="bac6a-171">displayReplyForm</span><span class="sxs-lookup"><span data-stu-id="bac6a-171">displayReplyForm</span></span>](#displayreplyformformdata) | <span data-ttu-id="bac6a-172">Methode</span><span class="sxs-lookup"><span data-stu-id="bac6a-172">Method</span></span> |
| [<span data-ttu-id="bac6a-173">getEntities</span><span class="sxs-lookup"><span data-stu-id="bac6a-173">getEntities</span></span>](#getentities--entitiesjavascriptapioutlook16officeentities) | <span data-ttu-id="bac6a-174">Methode</span><span class="sxs-lookup"><span data-stu-id="bac6a-174">Method</span></span> |
| [<span data-ttu-id="bac6a-175">getEntitiesByType</span><span class="sxs-lookup"><span data-stu-id="bac6a-175">getEntitiesByType</span></span>](#getentitiesbytypeentitytype--nullable-arraystringcontactjavascriptapioutlook16officecontactmeetingsuggestionjavascriptapioutlook16officemeetingsuggestionphonenumberjavascriptapioutlook16officephonenumbertasksuggestionjavascriptapioutlook16officetasksuggestion) | <span data-ttu-id="bac6a-176">Methode</span><span class="sxs-lookup"><span data-stu-id="bac6a-176">Method</span></span> |
| [<span data-ttu-id="bac6a-177">getFilteredEntitiesByName</span><span class="sxs-lookup"><span data-stu-id="bac6a-177">getFilteredEntitiesByName</span></span>](#getfilteredentitiesbynamename--nullable-arraystringcontactjavascriptapioutlook16officecontactmeetingsuggestionjavascriptapioutlook16officemeetingsuggestionphonenumberjavascriptapioutlook16officephonenumbertasksuggestionjavascriptapioutlook16officetasksuggestion) | <span data-ttu-id="bac6a-178">Methode</span><span class="sxs-lookup"><span data-stu-id="bac6a-178">Method</span></span> |
| [<span data-ttu-id="bac6a-179">getRegExMatches</span><span class="sxs-lookup"><span data-stu-id="bac6a-179">getRegExMatches</span></span>](#getregexmatches--object) | <span data-ttu-id="bac6a-180">Methode</span><span class="sxs-lookup"><span data-stu-id="bac6a-180">Method</span></span> |
| [<span data-ttu-id="bac6a-181">getRegExMatchesByName</span><span class="sxs-lookup"><span data-stu-id="bac6a-181">getRegExMatchesByName</span></span>](#getregexmatchesbynamename--nullable-array-string-) | <span data-ttu-id="bac6a-182">Methode</span><span class="sxs-lookup"><span data-stu-id="bac6a-182">Method</span></span> |
| [<span data-ttu-id="bac6a-183">getSelectedDataAsync</span><span class="sxs-lookup"><span data-stu-id="bac6a-183">getSelectedDataAsync</span></span>](#getselecteddataasynccoerciontype-options-callback--string) | <span data-ttu-id="bac6a-184">Methode</span><span class="sxs-lookup"><span data-stu-id="bac6a-184">Method</span></span> |
| [<span data-ttu-id="bac6a-185">getSelectedEntities</span><span class="sxs-lookup"><span data-stu-id="bac6a-185">getSelectedEntities</span></span>](#getselectedentities--entitiesjavascriptapioutlook16officeentities) | <span data-ttu-id="bac6a-186">Methode</span><span class="sxs-lookup"><span data-stu-id="bac6a-186">Method</span></span> |
| [<span data-ttu-id="bac6a-187">getSelectedRegExMatches</span><span class="sxs-lookup"><span data-stu-id="bac6a-187">getSelectedRegExMatches</span></span>](#getselectedregexmatches--object) | <span data-ttu-id="bac6a-188">Methode</span><span class="sxs-lookup"><span data-stu-id="bac6a-188">Method</span></span> |
| [<span data-ttu-id="bac6a-189">loadCustomPropertiesAsync</span><span class="sxs-lookup"><span data-stu-id="bac6a-189">loadCustomPropertiesAsync</span></span>](#loadcustompropertiesasynccallback-usercontext) | <span data-ttu-id="bac6a-190">Methode</span><span class="sxs-lookup"><span data-stu-id="bac6a-190">Method</span></span> |
| [<span data-ttu-id="bac6a-191">removeAttachmentAsync</span><span class="sxs-lookup"><span data-stu-id="bac6a-191">removeAttachmentAsync</span></span>](#removeattachmentasyncattachmentid-options-callback) | <span data-ttu-id="bac6a-192">Methode</span><span class="sxs-lookup"><span data-stu-id="bac6a-192">Method</span></span> |
| [<span data-ttu-id="bac6a-193">saveAsync</span><span class="sxs-lookup"><span data-stu-id="bac6a-193">saveAsync</span></span>](#saveasyncoptions-callback) | <span data-ttu-id="bac6a-194">Methode</span><span class="sxs-lookup"><span data-stu-id="bac6a-194">Method</span></span> |
| [<span data-ttu-id="bac6a-195">setSelectedDataAsync</span><span class="sxs-lookup"><span data-stu-id="bac6a-195">setSelectedDataAsync</span></span>](#setselecteddataasyncdata-options-callback) | <span data-ttu-id="bac6a-196">Methode</span><span class="sxs-lookup"><span data-stu-id="bac6a-196">Method</span></span> |

### <a name="example"></a><span data-ttu-id="bac6a-197">Beispiel</span><span class="sxs-lookup"><span data-stu-id="bac6a-197">Example</span></span>

<span data-ttu-id="bac6a-198">Im folgenden JavaScript-Codebeispiel wird der Zugriff auf die `subject`-Eigenschaft des aktuellen Elements in Outlook veranschaulicht.</span><span class="sxs-lookup"><span data-stu-id="bac6a-198">The following JavaScript code example shows how to access the `subject` property of the current item in Outlook.</span></span>

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

### <a name="members"></a><span data-ttu-id="bac6a-199">Elemente</span><span class="sxs-lookup"><span data-stu-id="bac6a-199">Members</span></span>

#### <a name="attachments-arrayattachmentdetailsjavascriptapioutlook16officeattachmentdetails"></a><span data-ttu-id="bac6a-200">attachments :Array.<[AttachmentDetails](/javascript/api/outlook_1_6/office.attachmentdetails)></span><span class="sxs-lookup"><span data-stu-id="bac6a-200">attachments :Array.<[AttachmentDetails](/javascript/api/outlook_1_6/office.attachmentdetails)></span></span>

<span data-ttu-id="bac6a-p102">Ruft ein Array mit Anlagen für das Element ab. Nur Lesemodus.</span><span class="sxs-lookup"><span data-stu-id="bac6a-p102">Gets an array of attachments for the item. Read mode only.</span></span>

> [!NOTE]
> <span data-ttu-id="bac6a-203">Bestimmte Dateitypen werden aufgrund von Sicherheitsproblemen von Outlook blockiert und daher nicht zurückgegeben werden.</span><span class="sxs-lookup"><span data-stu-id="bac6a-203">Certain types of files are blocked by Outlook due to potential security issues and are therefore not returned.</span></span> <span data-ttu-id="bac6a-204">Weitere Informationen finden Sie unter [Blockierte Anlagen in Outlook](https://support.office.com/article/Blocked-attachments-in-Outlook-434752E1-02D3-4E90-9124-8B81E49A8519).</span><span class="sxs-lookup"><span data-stu-id="bac6a-204">For more information, see [Blocked attachments in Outlook](https://support.office.com/article/Blocked-attachments-in-Outlook-434752E1-02D3-4E90-9124-8B81E49A8519).</span></span>

##### <a name="type"></a><span data-ttu-id="bac6a-205">Typ:</span><span class="sxs-lookup"><span data-stu-id="bac6a-205">Type:</span></span>

*   <span data-ttu-id="bac6a-206">Array.<[AttachmentDetails](/javascript/api/outlook_1_6/office.attachmentdetails)></span><span class="sxs-lookup"><span data-stu-id="bac6a-206">Array.<[AttachmentDetails](/javascript/api/outlook_1_6/office.attachmentdetails)></span></span>

##### <a name="requirements"></a><span data-ttu-id="bac6a-207">Anforderungen</span><span class="sxs-lookup"><span data-stu-id="bac6a-207">Requirements</span></span>

|<span data-ttu-id="bac6a-208">Anforderung</span><span class="sxs-lookup"><span data-stu-id="bac6a-208">Requirement</span></span>| <span data-ttu-id="bac6a-209">Wert</span><span class="sxs-lookup"><span data-stu-id="bac6a-209">Value</span></span>|
|---|---|
|[<span data-ttu-id="bac6a-210">Mindestversion des Postfachanforderungssatzes</span><span class="sxs-lookup"><span data-stu-id="bac6a-210">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="bac6a-211">1.0</span><span class="sxs-lookup"><span data-stu-id="bac6a-211">1.0</span></span>|
|[<span data-ttu-id="bac6a-212">Mindestberechtigungsstufe</span><span class="sxs-lookup"><span data-stu-id="bac6a-212">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="bac6a-213">ReadItem</span><span class="sxs-lookup"><span data-stu-id="bac6a-213">ReadItem</span></span>|
|[<span data-ttu-id="bac6a-214">Zutreffender Outlook-Modus</span><span class="sxs-lookup"><span data-stu-id="bac6a-214">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="bac6a-215">Lesen</span><span class="sxs-lookup"><span data-stu-id="bac6a-215">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="bac6a-216">Beispiel</span><span class="sxs-lookup"><span data-stu-id="bac6a-216">Example</span></span>

<span data-ttu-id="bac6a-217">Mit dem folgende Code wird eine HTML-Zeichenfolge mit Details aller Anlagen für das aktuelle Element erstellt.</span><span class="sxs-lookup"><span data-stu-id="bac6a-217">The following code builds an HTML string with details of all attachments on the current item.</span></span>

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

####  <a name="bcc-recipientsjavascriptapioutlook16officerecipients"></a><span data-ttu-id="bac6a-218">bcc :[Recipients](/javascript/api/outlook_1_6/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="bac6a-218">bcc :[Recipients](/javascript/api/outlook_1_6/office.recipients)</span></span>

<span data-ttu-id="bac6a-219">Ruft ein Objekt, das Methoden zum Abrufen oder aktualisieren die Empfänger in der Zeile Bcc (blind Carbon Copy, Blindkopie) einer Nachricht bereitstellt.</span><span class="sxs-lookup"><span data-stu-id="bac6a-219">Gets an object that provides methods to get or update the recipients on the Bcc (blind carbon copy) line of a message.</span></span> <span data-ttu-id="bac6a-220">Nur Verfassenmodus.</span><span class="sxs-lookup"><span data-stu-id="bac6a-220">Compose mode only.</span></span>

##### <a name="type"></a><span data-ttu-id="bac6a-221">Typ:</span><span class="sxs-lookup"><span data-stu-id="bac6a-221">Type:</span></span>

*   [<span data-ttu-id="bac6a-222">Recipients</span><span class="sxs-lookup"><span data-stu-id="bac6a-222">Recipients</span></span>](/javascript/api/outlook_1_6/office.recipients)

##### <a name="requirements"></a><span data-ttu-id="bac6a-223">Anforderungen</span><span class="sxs-lookup"><span data-stu-id="bac6a-223">Requirements</span></span>

|<span data-ttu-id="bac6a-224">Anforderung</span><span class="sxs-lookup"><span data-stu-id="bac6a-224">Requirement</span></span>| <span data-ttu-id="bac6a-225">Wert</span><span class="sxs-lookup"><span data-stu-id="bac6a-225">Value</span></span>|
|---|---|
|[<span data-ttu-id="bac6a-226">Mindestversion des Postfachanforderungssatzes</span><span class="sxs-lookup"><span data-stu-id="bac6a-226">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="bac6a-227">1.1</span><span class="sxs-lookup"><span data-stu-id="bac6a-227">1.1</span></span>|
|[<span data-ttu-id="bac6a-228">Mindestberechtigungsstufe</span><span class="sxs-lookup"><span data-stu-id="bac6a-228">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="bac6a-229">ReadItem</span><span class="sxs-lookup"><span data-stu-id="bac6a-229">ReadItem</span></span>|
|[<span data-ttu-id="bac6a-230">Zutreffender Outlook-Modus</span><span class="sxs-lookup"><span data-stu-id="bac6a-230">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="bac6a-231">Verfassen</span><span class="sxs-lookup"><span data-stu-id="bac6a-231">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="bac6a-232">Beispiel</span><span class="sxs-lookup"><span data-stu-id="bac6a-232">Example</span></span>

```
Office.context.mailbox.item.bcc.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.bcc.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.bcc.getAsync(callback);

function callback(asyncResult) {
  var arrayOfBccRecipients = asyncResult.value;
}
```

####  <a name="body-bodyjavascriptapioutlook16officebody"></a><span data-ttu-id="bac6a-233">body :[Body](/javascript/api/outlook_1_6/office.body)</span><span class="sxs-lookup"><span data-stu-id="bac6a-233">body :[Body](/javascript/api/outlook_1_6/office.body)</span></span>

<span data-ttu-id="bac6a-234">Ruft ein Objekt ab, das Methoden zum Bearbeiten des Textkörpers eines Elements bereitstellt.</span><span class="sxs-lookup"><span data-stu-id="bac6a-234">Gets an object that provides methods for manipulating the body of an item.</span></span>

##### <a name="type"></a><span data-ttu-id="bac6a-235">Typ:</span><span class="sxs-lookup"><span data-stu-id="bac6a-235">Type:</span></span>

*   [<span data-ttu-id="bac6a-236">Body</span><span class="sxs-lookup"><span data-stu-id="bac6a-236">Body</span></span>](/javascript/api/outlook_1_6/office.body)

##### <a name="requirements"></a><span data-ttu-id="bac6a-237">Anforderungen</span><span class="sxs-lookup"><span data-stu-id="bac6a-237">Requirements</span></span>

|<span data-ttu-id="bac6a-238">Anforderung</span><span class="sxs-lookup"><span data-stu-id="bac6a-238">Requirement</span></span>| <span data-ttu-id="bac6a-239">Wert</span><span class="sxs-lookup"><span data-stu-id="bac6a-239">Value</span></span>|
|---|---|
|[<span data-ttu-id="bac6a-240">Mindestversion des Postfachanforderungssatzes</span><span class="sxs-lookup"><span data-stu-id="bac6a-240">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="bac6a-241">1.1</span><span class="sxs-lookup"><span data-stu-id="bac6a-241">1.1</span></span>|
|[<span data-ttu-id="bac6a-242">Mindestberechtigungsstufe</span><span class="sxs-lookup"><span data-stu-id="bac6a-242">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="bac6a-243">ReadItem</span><span class="sxs-lookup"><span data-stu-id="bac6a-243">ReadItem</span></span>|
|[<span data-ttu-id="bac6a-244">Zutreffender Outlook-Modus</span><span class="sxs-lookup"><span data-stu-id="bac6a-244">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="bac6a-245">Verfassen oder Lesen</span><span class="sxs-lookup"><span data-stu-id="bac6a-245">Compose or read</span></span>|

####  <a name="cc-arrayemailaddressdetailsjavascriptapioutlook16officeemailaddressdetailsrecipientsjavascriptapioutlook16officerecipients"></a><span data-ttu-id="bac6a-246">cc: Array. <[EmailAddressDetails](/javascript/api/outlook_1_6/office.emailaddressdetails)>|[Empfänger](/javascript/api/outlook_1_6/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="bac6a-246">cc :Array.<[EmailAddressDetails](/javascript/api/outlook_1_6/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook_1_6/office.recipients)</span></span>

<span data-ttu-id="bac6a-247">Ermöglicht den Zugriff auf die (Carbon Copy, Blindkopie) Cc-Empfänger einer Nachricht.</span><span class="sxs-lookup"><span data-stu-id="bac6a-247">Provides access to the Cc (carbon copy) recipients of a message.</span></span> <span data-ttu-id="bac6a-248">Der Typ des Objekts und die Zugriffsebene, hängt von den Modus des aktuellen Elements ab.</span><span class="sxs-lookup"><span data-stu-id="bac6a-248">The type of object and level of access depends on the mode of the current item.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="bac6a-249">Lesemodus</span><span class="sxs-lookup"><span data-stu-id="bac6a-249">Read mode</span></span>

<span data-ttu-id="bac6a-p106">Die `cc`-Eigenschaft gibt ein Array mit einem `EmailAddressDetails`-Objekt für jeden Empfänger in der **Cc**-Zeile der Nachricht zurück. Die Auflistung ist auf höchstens 100 Elemente beschränkt.</span><span class="sxs-lookup"><span data-stu-id="bac6a-p106">The `cc` property returns an array that contains an `EmailAddressDetails` object for each recipient listed on the **Cc** line of the message. The collection is limited to a maximum of 100 members.</span></span>

##### <a name="compose-mode"></a><span data-ttu-id="bac6a-252">Verfassenmodus</span><span class="sxs-lookup"><span data-stu-id="bac6a-252">Compose mode</span></span>

<span data-ttu-id="bac6a-253">Die `cc` -Eigenschaft gibt eine `Recipients` -Objekt, das Methoden zum Abrufen oder aktualisieren die Empfänger in der Zeile **Cc** der Nachricht bereitstellt.</span><span class="sxs-lookup"><span data-stu-id="bac6a-253">The `cc` property returns a `Recipients` object that provides methods to get or update the recipients on the **Cc** line of the message.</span></span>

##### <a name="type"></a><span data-ttu-id="bac6a-254">Typ:</span><span class="sxs-lookup"><span data-stu-id="bac6a-254">Type:</span></span>

*   <span data-ttu-id="bac6a-255">Array.<[EmailAddressDetails](/javascript/api/outlook_1_6/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook_1_6/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="bac6a-255">Array.<[EmailAddressDetails](/javascript/api/outlook_1_6/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook_1_6/office.recipients)</span></span>

##### <a name="requirements"></a><span data-ttu-id="bac6a-256">Anforderungen</span><span class="sxs-lookup"><span data-stu-id="bac6a-256">Requirements</span></span>

|<span data-ttu-id="bac6a-257">Anforderung</span><span class="sxs-lookup"><span data-stu-id="bac6a-257">Requirement</span></span>| <span data-ttu-id="bac6a-258">Wert</span><span class="sxs-lookup"><span data-stu-id="bac6a-258">Value</span></span>|
|---|---|
|[<span data-ttu-id="bac6a-259">Mindestversion des Postfachanforderungssatzes</span><span class="sxs-lookup"><span data-stu-id="bac6a-259">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="bac6a-260">1.0</span><span class="sxs-lookup"><span data-stu-id="bac6a-260">1.0</span></span>|
|[<span data-ttu-id="bac6a-261">Mindestberechtigungsstufe</span><span class="sxs-lookup"><span data-stu-id="bac6a-261">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="bac6a-262">ReadItem</span><span class="sxs-lookup"><span data-stu-id="bac6a-262">ReadItem</span></span>|
|[<span data-ttu-id="bac6a-263">Zutreffender Outlook-Modus</span><span class="sxs-lookup"><span data-stu-id="bac6a-263">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="bac6a-264">Verfassen oder Lesen</span><span class="sxs-lookup"><span data-stu-id="bac6a-264">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="bac6a-265">Beispiel</span><span class="sxs-lookup"><span data-stu-id="bac6a-265">Example</span></span>

```
Office.context.mailbox.item.cc.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.cc.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.cc.getAsync(callback);

function callback(asyncResult) {
  var arrayOfCcRecipients = asyncResult.value;
}
```

####  <a name="nullable-conversationid-string"></a><span data-ttu-id="bac6a-266">(nullable) conversationId :String</span><span class="sxs-lookup"><span data-stu-id="bac6a-266">(nullable) conversationId :String</span></span>

<span data-ttu-id="bac6a-267">Ruft einen Bezeichner für die E-Mail-Unterhaltung ab, in der eine bestimmte Nachricht enthalten ist.</span><span class="sxs-lookup"><span data-stu-id="bac6a-267">Gets an identifier for the email conversation that contains a particular message.</span></span>

<span data-ttu-id="bac6a-p107">Sie können für diese Eigenschaft eine ganze Zahl abrufen, wenn Ihre Mail-App in Formularen zum Lesen oder Antworten in Formularen zum Verfassen aktiviert wird. Wenn der Benutzer den Betreff der Antwortnachricht ändert, ändert sich beim Versenden die Konversations-ID für die entsprechende Nachricht, und der Wert, den Sie vorher bezogen haben, trifft nicht länger zu.</span><span class="sxs-lookup"><span data-stu-id="bac6a-p107">You can get an integer for this property if your mail app is activated in read forms or responses in compose forms. If subsequently the user changes the subject of the reply message, upon sending the reply, the conversation ID for that message will change and that value you obtained earlier will no longer apply.</span></span>

<span data-ttu-id="bac6a-p108">Sie erhalten in einem Formular zum Verfassen für diese Eigenschaft für ein neues Element null. Wenn der Benutzer einen Betreff festlegt und das Element speichert, gibt die `conversationId`-Eigenschaft einen Wert zurück.</span><span class="sxs-lookup"><span data-stu-id="bac6a-p108">You get null for this property for a new item in a compose form. If the user sets a subject and saves the item, the `conversationId` property will return a value.</span></span>

##### <a name="type"></a><span data-ttu-id="bac6a-272">Typ:</span><span class="sxs-lookup"><span data-stu-id="bac6a-272">Type:</span></span>

*   <span data-ttu-id="bac6a-273">String</span><span class="sxs-lookup"><span data-stu-id="bac6a-273">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="bac6a-274">Anforderungen</span><span class="sxs-lookup"><span data-stu-id="bac6a-274">Requirements</span></span>

|<span data-ttu-id="bac6a-275">Anforderung</span><span class="sxs-lookup"><span data-stu-id="bac6a-275">Requirement</span></span>| <span data-ttu-id="bac6a-276">Wert</span><span class="sxs-lookup"><span data-stu-id="bac6a-276">Value</span></span>|
|---|---|
|[<span data-ttu-id="bac6a-277">Mindestversion des Postfachanforderungssatzes</span><span class="sxs-lookup"><span data-stu-id="bac6a-277">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="bac6a-278">1.0</span><span class="sxs-lookup"><span data-stu-id="bac6a-278">1.0</span></span>|
|[<span data-ttu-id="bac6a-279">Mindestberechtigungsstufe</span><span class="sxs-lookup"><span data-stu-id="bac6a-279">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="bac6a-280">ReadItem</span><span class="sxs-lookup"><span data-stu-id="bac6a-280">ReadItem</span></span>|
|[<span data-ttu-id="bac6a-281">Zutreffender Outlook-Modus</span><span class="sxs-lookup"><span data-stu-id="bac6a-281">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="bac6a-282">Verfassen oder Lesen</span><span class="sxs-lookup"><span data-stu-id="bac6a-282">Compose or read</span></span>|

#### <a name="datetimecreated-date"></a><span data-ttu-id="bac6a-283">dateTimeCreated :Date</span><span class="sxs-lookup"><span data-stu-id="bac6a-283">dateTimeCreated :Date</span></span>

<span data-ttu-id="bac6a-p109">Ruft das Datum und die Uhrzeit der Erstellung eines Elements ab. Nur Lesemodus.</span><span class="sxs-lookup"><span data-stu-id="bac6a-p109">Gets the date and time that an item was created. Read mode only.</span></span>

##### <a name="type"></a><span data-ttu-id="bac6a-286">Typ:</span><span class="sxs-lookup"><span data-stu-id="bac6a-286">Type:</span></span>

*   <span data-ttu-id="bac6a-287">Datum</span><span class="sxs-lookup"><span data-stu-id="bac6a-287">Date</span></span>

##### <a name="requirements"></a><span data-ttu-id="bac6a-288">Anforderungen</span><span class="sxs-lookup"><span data-stu-id="bac6a-288">Requirements</span></span>

|<span data-ttu-id="bac6a-289">Anforderung</span><span class="sxs-lookup"><span data-stu-id="bac6a-289">Requirement</span></span>| <span data-ttu-id="bac6a-290">Wert</span><span class="sxs-lookup"><span data-stu-id="bac6a-290">Value</span></span>|
|---|---|
|[<span data-ttu-id="bac6a-291">Mindestversion des Postfachanforderungssatzes</span><span class="sxs-lookup"><span data-stu-id="bac6a-291">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="bac6a-292">1.0</span><span class="sxs-lookup"><span data-stu-id="bac6a-292">1.0</span></span>|
|[<span data-ttu-id="bac6a-293">Mindestberechtigungsstufe</span><span class="sxs-lookup"><span data-stu-id="bac6a-293">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="bac6a-294">ReadItem</span><span class="sxs-lookup"><span data-stu-id="bac6a-294">ReadItem</span></span>|
|[<span data-ttu-id="bac6a-295">Zutreffender Outlook-Modus</span><span class="sxs-lookup"><span data-stu-id="bac6a-295">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="bac6a-296">Lesen</span><span class="sxs-lookup"><span data-stu-id="bac6a-296">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="bac6a-297">Beispiel</span><span class="sxs-lookup"><span data-stu-id="bac6a-297">Example</span></span>

```
var created = Office.context.mailbox.item.dateTimeCreated;
```

#### <a name="datetimemodified-date"></a><span data-ttu-id="bac6a-298">dateTimeModified :Date</span><span class="sxs-lookup"><span data-stu-id="bac6a-298">dateTimeModified :Date</span></span>

<span data-ttu-id="bac6a-p110">Ruft das Datum und die Uhrzeit der letzten Änderung eines Elements ab. Nur Lesemodus.</span><span class="sxs-lookup"><span data-stu-id="bac6a-p110">Gets the date and time that an item was last modified. Read mode only.</span></span>

> [!NOTE]
> <span data-ttu-id="bac6a-301">Dieser Member wird in Outlook für iOS oder Outlook für Android nicht unterstützt.</span><span class="sxs-lookup"><span data-stu-id="bac6a-301">This member is not supported in Outlook for iOS or Outlook for Android.</span></span>

##### <a name="type"></a><span data-ttu-id="bac6a-302">Typ:</span><span class="sxs-lookup"><span data-stu-id="bac6a-302">Type:</span></span>

*   <span data-ttu-id="bac6a-303">Datum</span><span class="sxs-lookup"><span data-stu-id="bac6a-303">Date</span></span>

##### <a name="requirements"></a><span data-ttu-id="bac6a-304">Anforderungen</span><span class="sxs-lookup"><span data-stu-id="bac6a-304">Requirements</span></span>

|<span data-ttu-id="bac6a-305">Anforderung</span><span class="sxs-lookup"><span data-stu-id="bac6a-305">Requirement</span></span>| <span data-ttu-id="bac6a-306">Wert</span><span class="sxs-lookup"><span data-stu-id="bac6a-306">Value</span></span>|
|---|---|
|[<span data-ttu-id="bac6a-307">Mindestversion des Postfachanforderungssatzes</span><span class="sxs-lookup"><span data-stu-id="bac6a-307">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="bac6a-308">1.0</span><span class="sxs-lookup"><span data-stu-id="bac6a-308">1.0</span></span>|
|[<span data-ttu-id="bac6a-309">Mindestberechtigungsstufe</span><span class="sxs-lookup"><span data-stu-id="bac6a-309">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="bac6a-310">ReadItem</span><span class="sxs-lookup"><span data-stu-id="bac6a-310">ReadItem</span></span>|
|[<span data-ttu-id="bac6a-311">Zutreffender Outlook-Modus</span><span class="sxs-lookup"><span data-stu-id="bac6a-311">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="bac6a-312">Lesen</span><span class="sxs-lookup"><span data-stu-id="bac6a-312">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="bac6a-313">Beispiel</span><span class="sxs-lookup"><span data-stu-id="bac6a-313">Example</span></span>

```
var modified = Office.context.mailbox.item.dateTimeModified;
```

####  <a name="end-datetimejavascriptapioutlook16officetime"></a><span data-ttu-id="bac6a-314">end :Date|[Time](/javascript/api/outlook_1_6/office.time)</span><span class="sxs-lookup"><span data-stu-id="bac6a-314">end :Date|[Time](/javascript/api/outlook_1_6/office.time)</span></span>

<span data-ttu-id="bac6a-315">Ruft Datum und Zeit für das Ende des Termins ab oder legt diese fest.</span><span class="sxs-lookup"><span data-stu-id="bac6a-315">Gets or sets the date and time that the appointment is to end.</span></span>

<span data-ttu-id="bac6a-p111">Die `end`-Eigenschaft wird als Datums- und Uhrzeitwert in UTC ausgedrückt. Sie können die [`convertToLocalClientTime`](office.context.mailbox.md#converttolocalclienttimetimevalue--localclienttimejavascriptapioutlook16officelocalclienttime)-Methode verwenden, um den Endeigenschaftswert in den lokalen Uhrzeit- und Datumswert des Clients umzuwandeln.</span><span class="sxs-lookup"><span data-stu-id="bac6a-p111">The `end` property is expressed as a Coordinated Universal Time (UTC) date and time value. You can use the [`convertToLocalClientTime`](office.context.mailbox.md#converttolocalclienttimetimevalue--localclienttimejavascriptapioutlook16officelocalclienttime) method to convert the end property value to the client’s local date and time.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="bac6a-318">Lesemodus</span><span class="sxs-lookup"><span data-stu-id="bac6a-318">Read mode</span></span>

<span data-ttu-id="bac6a-319">Die `end`-Eigenschaft gibt ein `Date`-Objekt zurück.</span><span class="sxs-lookup"><span data-stu-id="bac6a-319">The `end` property returns a `Date` object.</span></span>

##### <a name="compose-mode"></a><span data-ttu-id="bac6a-320">Verfassenmodus</span><span class="sxs-lookup"><span data-stu-id="bac6a-320">Compose mode</span></span>

<span data-ttu-id="bac6a-321">Die `end`-Eigenschaft gibt ein `Time`-Objekt zurück.</span><span class="sxs-lookup"><span data-stu-id="bac6a-321">The `end` property returns a `Time` object.</span></span>

<span data-ttu-id="bac6a-322">Wenn Sie die [`Time.setAsync`](/javascript/api/outlook_1_6/office.time#setasync-datetime--options--callback-)-Methode verwenden, um die Endzeit im Verfassenmodus festzulegen, sollten Sie die [`convertToUtcClientTime`](office.context.mailbox.md#converttoutcclienttimeinput--date)-Methode verwenden, um die Ortszeit auf dem Client für den Server in UTC umzuwandeln.</span><span class="sxs-lookup"><span data-stu-id="bac6a-322">When you use the [`Time.setAsync`](/javascript/api/outlook_1_6/office.time#setasync-datetime--options--callback-) method to set the end time, you should use the [`convertToUtcClientTime`](office.context.mailbox.md#converttoutcclienttimeinput--date) method to convert the local time on the client to UTC for the server.</span></span>

##### <a name="type"></a><span data-ttu-id="bac6a-323">Typ:</span><span class="sxs-lookup"><span data-stu-id="bac6a-323">Type:</span></span>

*   <span data-ttu-id="bac6a-324">Date | [Time](/javascript/api/outlook_1_6/office.time)</span><span class="sxs-lookup"><span data-stu-id="bac6a-324">Date | [Time](/javascript/api/outlook_1_6/office.time)</span></span>

##### <a name="requirements"></a><span data-ttu-id="bac6a-325">Anforderungen</span><span class="sxs-lookup"><span data-stu-id="bac6a-325">Requirements</span></span>

|<span data-ttu-id="bac6a-326">Anforderung</span><span class="sxs-lookup"><span data-stu-id="bac6a-326">Requirement</span></span>| <span data-ttu-id="bac6a-327">Wert</span><span class="sxs-lookup"><span data-stu-id="bac6a-327">Value</span></span>|
|---|---|
|[<span data-ttu-id="bac6a-328">Mindestversion des Postfachanforderungssatzes</span><span class="sxs-lookup"><span data-stu-id="bac6a-328">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="bac6a-329">1.0</span><span class="sxs-lookup"><span data-stu-id="bac6a-329">1.0</span></span>|
|[<span data-ttu-id="bac6a-330">Mindestberechtigungsstufe</span><span class="sxs-lookup"><span data-stu-id="bac6a-330">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="bac6a-331">ReadItem</span><span class="sxs-lookup"><span data-stu-id="bac6a-331">ReadItem</span></span>|
|[<span data-ttu-id="bac6a-332">Zutreffender Outlook-Modus</span><span class="sxs-lookup"><span data-stu-id="bac6a-332">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="bac6a-333">Verfassen oder Lesen</span><span class="sxs-lookup"><span data-stu-id="bac6a-333">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="bac6a-334">Beispiel</span><span class="sxs-lookup"><span data-stu-id="bac6a-334">Example</span></span>

<span data-ttu-id="bac6a-335">Im folgenden Beispiel wird die Endzeit eines Termins im Verfassenmodus mithilfe der [`setAsync`](/javascript/api/outlook_1_6/office.time#setasync-datetime--options--callback-)-Methode des `Time`-Objekts festgelegt.</span><span class="sxs-lookup"><span data-stu-id="bac6a-335">The following example sets the end time of an appointment in compose mode by using the [`setAsync`](/javascript/api/outlook_1_6/office.time#setasync-datetime--options--callback-) method of the `Time` object.</span></span>

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

#### <a name="from-emailaddressdetailsjavascriptapioutlook16officeemailaddressdetails"></a><span data-ttu-id="bac6a-336">from :[EmailAddressDetails](/javascript/api/outlook_1_6/office.emailaddressdetails)</span><span class="sxs-lookup"><span data-stu-id="bac6a-336">from :[EmailAddressDetails](/javascript/api/outlook_1_6/office.emailaddressdetails)</span></span>

<span data-ttu-id="bac6a-p112">Ruft die E-Mail-Adresse des Absenders einer Nachricht ab. Nur Lesemodus.</span><span class="sxs-lookup"><span data-stu-id="bac6a-p112">Gets the email address of the sender of a message. Read mode only.</span></span>

<span data-ttu-id="bac6a-p113">Die `from`- und [`sender`](#sender-emailaddressdetailsjavascriptapioutlook16officeemailaddressdetails)-Eigenschaften stellen dieselbe Person dar, außer die Nachricht wurde von einem Delegaten gesendet. In diesem Fall stellt die `from`-Eigenschaft die Stellvertretung dar, und die sender-Eigenschaft stellt den Delegaten dar.</span><span class="sxs-lookup"><span data-stu-id="bac6a-p113">The `from` and [`sender`](#sender-emailaddressdetailsjavascriptapioutlook16officeemailaddressdetails) properties represent the same person unless the message is sent by a delegate. In that case, the `from` property represents the delegator, and the sender property represents the delegate.</span></span>

> [!NOTE]
> <span data-ttu-id="bac6a-341">Die `recipientType` -Eigenschaft des der `EmailAddressDetails` -Objekts der `from` -Eigenschaft ist `undefined`.</span><span class="sxs-lookup"><span data-stu-id="bac6a-341">The `recipientType` property of the `EmailAddressDetails` object in the `from` property is `undefined`.</span></span>

##### <a name="type"></a><span data-ttu-id="bac6a-342">Typ:</span><span class="sxs-lookup"><span data-stu-id="bac6a-342">Type:</span></span>

*   [<span data-ttu-id="bac6a-343">EmailAddressDetails</span><span class="sxs-lookup"><span data-stu-id="bac6a-343">EmailAddressDetails</span></span>](/javascript/api/outlook_1_6/office.emailaddressdetails)

##### <a name="requirements"></a><span data-ttu-id="bac6a-344">Anforderungen</span><span class="sxs-lookup"><span data-stu-id="bac6a-344">Requirements</span></span>

|<span data-ttu-id="bac6a-345">Anforderung</span><span class="sxs-lookup"><span data-stu-id="bac6a-345">Requirement</span></span>| <span data-ttu-id="bac6a-346">Wert</span><span class="sxs-lookup"><span data-stu-id="bac6a-346">Value</span></span>|
|---|---|
|[<span data-ttu-id="bac6a-347">Mindestversion des Postfachanforderungssatzes</span><span class="sxs-lookup"><span data-stu-id="bac6a-347">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="bac6a-348">1.0</span><span class="sxs-lookup"><span data-stu-id="bac6a-348">1.0</span></span>|
|[<span data-ttu-id="bac6a-349">Mindestberechtigungsstufe</span><span class="sxs-lookup"><span data-stu-id="bac6a-349">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="bac6a-350">ReadItem</span><span class="sxs-lookup"><span data-stu-id="bac6a-350">ReadItem</span></span>|
|[<span data-ttu-id="bac6a-351">Zutreffender Outlook-Modus</span><span class="sxs-lookup"><span data-stu-id="bac6a-351">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="bac6a-352">Lesen</span><span class="sxs-lookup"><span data-stu-id="bac6a-352">Read</span></span>|

#### <a name="internetmessageid-string"></a><span data-ttu-id="bac6a-353">internetMessageId :String</span><span class="sxs-lookup"><span data-stu-id="bac6a-353">internetMessageId :String</span></span>

<span data-ttu-id="bac6a-p114">Ruft die Internetnachricht-ID einer E-Mail-Nachricht ab. Nur Lesemodus.</span><span class="sxs-lookup"><span data-stu-id="bac6a-p114">Gets the Internet message identifier for an email message. Read mode only.</span></span>

##### <a name="type"></a><span data-ttu-id="bac6a-356">Typ:</span><span class="sxs-lookup"><span data-stu-id="bac6a-356">Type:</span></span>

*   <span data-ttu-id="bac6a-357">String</span><span class="sxs-lookup"><span data-stu-id="bac6a-357">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="bac6a-358">Anforderungen</span><span class="sxs-lookup"><span data-stu-id="bac6a-358">Requirements</span></span>

|<span data-ttu-id="bac6a-359">Anforderung</span><span class="sxs-lookup"><span data-stu-id="bac6a-359">Requirement</span></span>| <span data-ttu-id="bac6a-360">Wert</span><span class="sxs-lookup"><span data-stu-id="bac6a-360">Value</span></span>|
|---|---|
|[<span data-ttu-id="bac6a-361">Mindestversion des Postfachanforderungssatzes</span><span class="sxs-lookup"><span data-stu-id="bac6a-361">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="bac6a-362">1.0</span><span class="sxs-lookup"><span data-stu-id="bac6a-362">1.0</span></span>|
|[<span data-ttu-id="bac6a-363">Mindestberechtigungsstufe</span><span class="sxs-lookup"><span data-stu-id="bac6a-363">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="bac6a-364">ReadItem</span><span class="sxs-lookup"><span data-stu-id="bac6a-364">ReadItem</span></span>|
|[<span data-ttu-id="bac6a-365">Zutreffender Outlook-Modus</span><span class="sxs-lookup"><span data-stu-id="bac6a-365">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="bac6a-366">Lesen</span><span class="sxs-lookup"><span data-stu-id="bac6a-366">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="bac6a-367">Beispiel</span><span class="sxs-lookup"><span data-stu-id="bac6a-367">Example</span></span>

```
var internetMessageId = Office.context.mailbox.item.internetMessageId;
```

#### <a name="itemclass-string"></a><span data-ttu-id="bac6a-368">itemClass :String</span><span class="sxs-lookup"><span data-stu-id="bac6a-368">itemClass :String</span></span>

<span data-ttu-id="bac6a-p115">Ruft die Exchange-Webdienste-Elementklasse des ausgewählten Elements ab. Nur Lesemodus.</span><span class="sxs-lookup"><span data-stu-id="bac6a-p115">Gets the Exchange Web Services item class of the selected item. Read mode only.</span></span>

<span data-ttu-id="bac6a-p116">Die `itemClass`-Eigenschaft gibt die Nachrichtenklasse des ausgewählten Elements an. Folgende Standardnachrichtenklassen für das Nachrichten- oder Terminelement sind vorhanden:</span><span class="sxs-lookup"><span data-stu-id="bac6a-p116">The `itemClass` property specifies the message class of the selected item. The following are the default message classes for the message or appointment item.</span></span>

| <span data-ttu-id="bac6a-373">Typ</span><span class="sxs-lookup"><span data-stu-id="bac6a-373">Type</span></span> | <span data-ttu-id="bac6a-374">Beschreibung</span><span class="sxs-lookup"><span data-stu-id="bac6a-374">Description</span></span> | <span data-ttu-id="bac6a-375">Elementklasse</span><span class="sxs-lookup"><span data-stu-id="bac6a-375">item class</span></span> |
| --- | --- | --- |
| <span data-ttu-id="bac6a-376">Terminelemente</span><span class="sxs-lookup"><span data-stu-id="bac6a-376">Appointment items</span></span> | <span data-ttu-id="bac6a-377">Dies sind Kalenderelemente der Elementklasse `IPM.Appointment` oder `IPM.Appointment.Occurence`.</span><span class="sxs-lookup"><span data-stu-id="bac6a-377">These are calendar items of the item class `IPM.Appointment` or `IPM.Appointment.Occurence`.</span></span> | `IPM.Appointment`<br />`IPM.Appointment.Occurence` |
| <span data-ttu-id="bac6a-378">Nachrichtenelemente</span><span class="sxs-lookup"><span data-stu-id="bac6a-378">Message items</span></span> | <span data-ttu-id="bac6a-379">Diese Elemente umfassen E-Mail-Nachrichten der Standardnachrichtenklasse `IPM.Note` sowie Besprechungsanfragen, -antworten und -absagen, die `IPM.Schedule.Meeting` als Basisnachrichtenklasse verwenden.</span><span class="sxs-lookup"><span data-stu-id="bac6a-379">These include email messages that have the default message class `IPM.Note`, and meeting requests, responses, and cancellations, that use `IPM.Schedule.Meeting` as the base message class.</span></span> | `IPM.Note`<br />`IPM.Schedule.Meeting.Request`<br />`IPM.Schedule.Meeting.Neg`<br />`IPM.Schedule.Meeting.Pos`<br />`IPM.Schedule.Meeting.Tent`<br />`IPM.Schedule.Meeting.Canceled` |

<span data-ttu-id="bac6a-380">Sie können benutzerdefinierte Nachrichtenklassen zum Erweitern einer Standardnachrichtenklasse erstellen, z. B. eine benutzerdefinierte Terminnachrichtenklasse wie `IPM.Appointment.Contoso`.</span><span class="sxs-lookup"><span data-stu-id="bac6a-380">You can create custom message classes that extends a default message class, for example, a custom appointment message class `IPM.Appointment.Contoso`.</span></span>

##### <a name="type"></a><span data-ttu-id="bac6a-381">Typ:</span><span class="sxs-lookup"><span data-stu-id="bac6a-381">Type:</span></span>

*   <span data-ttu-id="bac6a-382">String</span><span class="sxs-lookup"><span data-stu-id="bac6a-382">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="bac6a-383">Anforderungen</span><span class="sxs-lookup"><span data-stu-id="bac6a-383">Requirements</span></span>

|<span data-ttu-id="bac6a-384">Anforderung</span><span class="sxs-lookup"><span data-stu-id="bac6a-384">Requirement</span></span>| <span data-ttu-id="bac6a-385">Wert</span><span class="sxs-lookup"><span data-stu-id="bac6a-385">Value</span></span>|
|---|---|
|[<span data-ttu-id="bac6a-386">Mindestversion des Postfachanforderungssatzes</span><span class="sxs-lookup"><span data-stu-id="bac6a-386">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="bac6a-387">1.0</span><span class="sxs-lookup"><span data-stu-id="bac6a-387">1.0</span></span>|
|[<span data-ttu-id="bac6a-388">Mindestberechtigungsstufe</span><span class="sxs-lookup"><span data-stu-id="bac6a-388">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="bac6a-389">ReadItem</span><span class="sxs-lookup"><span data-stu-id="bac6a-389">ReadItem</span></span>|
|[<span data-ttu-id="bac6a-390">Zutreffender Outlook-Modus</span><span class="sxs-lookup"><span data-stu-id="bac6a-390">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="bac6a-391">Lesen</span><span class="sxs-lookup"><span data-stu-id="bac6a-391">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="bac6a-392">Beispiel</span><span class="sxs-lookup"><span data-stu-id="bac6a-392">Example</span></span>

```
var itemClass = Office.context.mailbox.item.itemClass;
```

#### <a name="nullable-itemid-string"></a><span data-ttu-id="bac6a-393">(nullable) itemId :String</span><span class="sxs-lookup"><span data-stu-id="bac6a-393">(nullable) itemId :String</span></span>

<span data-ttu-id="bac6a-p117">Ruft den Exchange-Webdienste-Elementbezeichner für das aktuelle Element ab. Nur Lesemodus.</span><span class="sxs-lookup"><span data-stu-id="bac6a-p117">Gets the Exchange Web Services item identifier for the current item. Read mode only.</span></span>

> [!NOTE]
> <span data-ttu-id="bac6a-396">Der Bezeichner, der von der `itemId`-Eigenschaft zurückgegeben wird, ist der gleiche Bezeichner wie der Exchange-Webdienste-Elementbezeichner.</span><span class="sxs-lookup"><span data-stu-id="bac6a-396">The identifier returned by the `itemId` property is the same as the Exchange Web Services item identifier.</span></span> <span data-ttu-id="bac6a-397">Die `itemId` -Eigenschaft ist nicht identisch mit der Outlook-Eintrags-ID oder die ID von der Outlook-REST-API verwendet.</span><span class="sxs-lookup"><span data-stu-id="bac6a-397">The `itemId` property is not identical to the Outlook Entry ID or the ID used by the Outlook REST API.</span></span> <span data-ttu-id="bac6a-398">REST-API-Aufrufe dieser Wert verwenden, sollten sie konvertiert werden, bevor [Office.context.mailbox.convertToRestId](office.context.mailbox.md#converttorestiditemid-restversion--string)verwenden.</span><span class="sxs-lookup"><span data-stu-id="bac6a-398">Before making REST API calls using this value, it should be converted using [Office.context.mailbox.convertToRestId](office.context.mailbox.md#converttorestiditemid-restversion--string).</span></span> <span data-ttu-id="bac6a-399">Weitere Informationen finden Sie unter [Verwenden der Outlook-REST-APIs aus einem Outlook-add-in](https://docs.microsoft.com/outlook/add-ins/use-rest-api#get-the-item-id).</span><span class="sxs-lookup"><span data-stu-id="bac6a-399">For more details, see [Use the Outlook REST APIs from an Outlook add-in](https://docs.microsoft.com/outlook/add-ins/use-rest-api#get-the-item-id).</span></span>

<span data-ttu-id="bac6a-p119">Die Eigenschaft `itemId` ist im Modus „Verfassen“ nicht verfügbar. Wenn ein Elementbezeichner erforderlich ist, kann die [`saveAsync`](#saveasyncoptions-callback)-Methode verwendet werden, um das Element zu speichern. Dadurch wird der Elementbezeichner im [`AsyncResult.value`](/javascript/api/office/office.asyncresult)-Parameter in der Rückruffunktion zurückgegeben.</span><span class="sxs-lookup"><span data-stu-id="bac6a-p119">The `itemId` property is not available in compose mode. If an item identifier is required, the [`saveAsync`](#saveasyncoptions-callback) method can be used to save the item to the store, which will return the item identifier in the [`AsyncResult.value`](/javascript/api/office/office.asyncresult) parameter in the callback function.</span></span>

##### <a name="type"></a><span data-ttu-id="bac6a-402">Typ:</span><span class="sxs-lookup"><span data-stu-id="bac6a-402">Type:</span></span>

*   <span data-ttu-id="bac6a-403">String</span><span class="sxs-lookup"><span data-stu-id="bac6a-403">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="bac6a-404">Anforderungen</span><span class="sxs-lookup"><span data-stu-id="bac6a-404">Requirements</span></span>

|<span data-ttu-id="bac6a-405">Anforderung</span><span class="sxs-lookup"><span data-stu-id="bac6a-405">Requirement</span></span>| <span data-ttu-id="bac6a-406">Wert</span><span class="sxs-lookup"><span data-stu-id="bac6a-406">Value</span></span>|
|---|---|
|[<span data-ttu-id="bac6a-407">Mindestversion des Postfachanforderungssatzes</span><span class="sxs-lookup"><span data-stu-id="bac6a-407">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="bac6a-408">1.0</span><span class="sxs-lookup"><span data-stu-id="bac6a-408">1.0</span></span>|
|[<span data-ttu-id="bac6a-409">Mindestberechtigungsstufe</span><span class="sxs-lookup"><span data-stu-id="bac6a-409">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="bac6a-410">ReadItem</span><span class="sxs-lookup"><span data-stu-id="bac6a-410">ReadItem</span></span>|
|[<span data-ttu-id="bac6a-411">Zutreffender Outlook-Modus</span><span class="sxs-lookup"><span data-stu-id="bac6a-411">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="bac6a-412">Lesen</span><span class="sxs-lookup"><span data-stu-id="bac6a-412">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="bac6a-413">Beispiel</span><span class="sxs-lookup"><span data-stu-id="bac6a-413">Example</span></span>

<span data-ttu-id="bac6a-p120">Mit dem folgende Code wird das Vorhandensein eines Elementbezeichners überprüft. Wenn die `itemId`-Eigenschaft `null` oder `undefined` zurückgibt, speichert sie das Element und ruft den Elementbezeichner aus dem asynchronen Ergebnis ab.</span><span class="sxs-lookup"><span data-stu-id="bac6a-p120">The following code checks for the presence of an item identifier. If the `itemId` property returns `null` or `undefined`, it saves the item to the store and gets the item identifier from the asynchronous result.</span></span>

```
var itemId = Office.context.mailbox.item.itemId;
if (itemId === null || itemId == undefined) {
  Office.context.mailbox.item.saveAsync(function(result){
    itemId = result.value;
  });
}
```

####  <a name="itemtype-officemailboxenumsitemtypejavascriptapioutlook16officemailboxenumsitemtype"></a><span data-ttu-id="bac6a-416">itemType :[Office.MailboxEnums.ItemType](/javascript/api/outlook_1_6/office.mailboxenums.itemtype)</span><span class="sxs-lookup"><span data-stu-id="bac6a-416">itemType :[Office.MailboxEnums.ItemType](/javascript/api/outlook_1_6/office.mailboxenums.itemtype)</span></span>

<span data-ttu-id="bac6a-417">Ruft den Typ des Elements ab, das eine Instanz darstellt.</span><span class="sxs-lookup"><span data-stu-id="bac6a-417">Gets the type of item that an instance represents.</span></span>

<span data-ttu-id="bac6a-418">Die `itemType`-Eigenschaft gibt einen der Werte der `ItemType`-Enumeration zurück, der angibt, ob es sich bei der `item`-Objektinstanz um eine Nachricht oder einen Termin handelt.</span><span class="sxs-lookup"><span data-stu-id="bac6a-418">The `itemType` property returns one of the `ItemType` enumeration values, indicating whether the `item` object instance is a message or an appointment.</span></span>

##### <a name="type"></a><span data-ttu-id="bac6a-419">Typ:</span><span class="sxs-lookup"><span data-stu-id="bac6a-419">Type:</span></span>

*   [<span data-ttu-id="bac6a-420">Office.MailboxEnums.ItemType</span><span class="sxs-lookup"><span data-stu-id="bac6a-420">Office.MailboxEnums.ItemType</span></span>](/javascript/api/outlook_1_6/office.mailboxenums.itemtype)

##### <a name="requirements"></a><span data-ttu-id="bac6a-421">Anforderungen</span><span class="sxs-lookup"><span data-stu-id="bac6a-421">Requirements</span></span>

|<span data-ttu-id="bac6a-422">Anforderung</span><span class="sxs-lookup"><span data-stu-id="bac6a-422">Requirement</span></span>| <span data-ttu-id="bac6a-423">Wert</span><span class="sxs-lookup"><span data-stu-id="bac6a-423">Value</span></span>|
|---|---|
|[<span data-ttu-id="bac6a-424">Mindestversion des Postfachanforderungssatzes</span><span class="sxs-lookup"><span data-stu-id="bac6a-424">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="bac6a-425">1.0</span><span class="sxs-lookup"><span data-stu-id="bac6a-425">1.0</span></span>|
|[<span data-ttu-id="bac6a-426">Mindestberechtigungsstufe</span><span class="sxs-lookup"><span data-stu-id="bac6a-426">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="bac6a-427">ReadItem</span><span class="sxs-lookup"><span data-stu-id="bac6a-427">ReadItem</span></span>|
|[<span data-ttu-id="bac6a-428">Zutreffender Outlook-Modus</span><span class="sxs-lookup"><span data-stu-id="bac6a-428">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="bac6a-429">Verfassen oder Lesen</span><span class="sxs-lookup"><span data-stu-id="bac6a-429">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="bac6a-430">Beispiel</span><span class="sxs-lookup"><span data-stu-id="bac6a-430">Example</span></span>

```
if (Office.context.mailbox.item.itemType == Office.MailboxEnums.ItemType.Message)
  // do something
else
  // do something else
```

####  <a name="location-stringlocationjavascriptapioutlook16officelocation"></a><span data-ttu-id="bac6a-431">location :String|[Location](/javascript/api/outlook_1_6/office.location)</span><span class="sxs-lookup"><span data-stu-id="bac6a-431">location :String|[Location](/javascript/api/outlook_1_6/office.location)</span></span>

<span data-ttu-id="bac6a-432">Ruft den Ort eines Termins ab bzw. legt ihn fest.</span><span class="sxs-lookup"><span data-stu-id="bac6a-432">Gets or sets the location of an appointment.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="bac6a-433">Lesemodus</span><span class="sxs-lookup"><span data-stu-id="bac6a-433">Read mode</span></span>

<span data-ttu-id="bac6a-434">Die `location`-Eigenschaft gibt eine Zeichenfolge zurück, die den Ort des Termins enthält.</span><span class="sxs-lookup"><span data-stu-id="bac6a-434">The `location` property returns a string that contains the location of the appointment.</span></span>

##### <a name="compose-mode"></a><span data-ttu-id="bac6a-435">Verfassenmodus</span><span class="sxs-lookup"><span data-stu-id="bac6a-435">Compose mode</span></span>

<span data-ttu-id="bac6a-436">Die `location`-Eigenschaft gibt ein `Location`-Objekt zurück, das Methoden zum Abrufen und Festlegen des Orts für einen Termin bereitstellt.</span><span class="sxs-lookup"><span data-stu-id="bac6a-436">The `location` property returns a `Location` object that provides methods that are used to get and set the location of the appointment.</span></span>

##### <a name="type"></a><span data-ttu-id="bac6a-437">Typ:</span><span class="sxs-lookup"><span data-stu-id="bac6a-437">Type:</span></span>

*   <span data-ttu-id="bac6a-438">String | [Location](/javascript/api/outlook_1_6/office.location)</span><span class="sxs-lookup"><span data-stu-id="bac6a-438">String | [Location](/javascript/api/outlook_1_6/office.location)</span></span>

##### <a name="requirements"></a><span data-ttu-id="bac6a-439">Anforderungen</span><span class="sxs-lookup"><span data-stu-id="bac6a-439">Requirements</span></span>

|<span data-ttu-id="bac6a-440">Anforderung</span><span class="sxs-lookup"><span data-stu-id="bac6a-440">Requirement</span></span>| <span data-ttu-id="bac6a-441">Wert</span><span class="sxs-lookup"><span data-stu-id="bac6a-441">Value</span></span>|
|---|---|
|[<span data-ttu-id="bac6a-442">Mindestversion des Postfachanforderungssatzes</span><span class="sxs-lookup"><span data-stu-id="bac6a-442">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="bac6a-443">1.0</span><span class="sxs-lookup"><span data-stu-id="bac6a-443">1.0</span></span>|
|[<span data-ttu-id="bac6a-444">Mindestberechtigungsstufe</span><span class="sxs-lookup"><span data-stu-id="bac6a-444">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="bac6a-445">ReadItem</span><span class="sxs-lookup"><span data-stu-id="bac6a-445">ReadItem</span></span>|
|[<span data-ttu-id="bac6a-446">Zutreffender Outlook-Modus</span><span class="sxs-lookup"><span data-stu-id="bac6a-446">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="bac6a-447">Verfassen oder Lesen</span><span class="sxs-lookup"><span data-stu-id="bac6a-447">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="bac6a-448">Beispiel</span><span class="sxs-lookup"><span data-stu-id="bac6a-448">Example</span></span>

```
var userContext = { value : 1 };
Office.context.mailbox.item.location.getAsync( { context: userContext}, callback);

function callback(asyncResult) {
  var context = asyncResult.context;
  var location = asyncResult.value;
}
```

#### <a name="normalizedsubject-string"></a><span data-ttu-id="bac6a-449">normalizedSubject :String</span><span class="sxs-lookup"><span data-stu-id="bac6a-449">normalizedSubject :String</span></span>

<span data-ttu-id="bac6a-p121">Ruft den Betreff für ein Element ab, wobei alle Präfixe entfernt werden (einschließlich `RE:` und `FWD:`). Nur Lesemodus.</span><span class="sxs-lookup"><span data-stu-id="bac6a-p121">Gets the subject of an item, with all prefixes removed (including `RE:` and `FWD:`). Read mode only.</span></span>

<span data-ttu-id="bac6a-p122">Die normalizedSubject-Eigenschaft ruft den Betreff des Elements mit allen Standardpräfixen (z. B. `RE:` und `FW:`) ab, die von E-Mail-Programmen hinzugefügt werden. Wenn der Betreff des Elements mit vollständigen Präfixen abgerufen werden soll, verwenden Sie die [`subject`](#subject-stringsubjectjavascriptapioutlook16officesubject)-Eigenschaft.</span><span class="sxs-lookup"><span data-stu-id="bac6a-p122">The normalizedSubject property gets the subject of the item, with any standard prefixes (such as `RE:` and `FW:`) that are added by email programs. To get the subject of the item with the prefixes intact, use the [`subject`](#subject-stringsubjectjavascriptapioutlook16officesubject) property.</span></span>

##### <a name="type"></a><span data-ttu-id="bac6a-454">Typ:</span><span class="sxs-lookup"><span data-stu-id="bac6a-454">Type:</span></span>

*   <span data-ttu-id="bac6a-455">String</span><span class="sxs-lookup"><span data-stu-id="bac6a-455">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="bac6a-456">Anforderungen</span><span class="sxs-lookup"><span data-stu-id="bac6a-456">Requirements</span></span>

|<span data-ttu-id="bac6a-457">Anforderung</span><span class="sxs-lookup"><span data-stu-id="bac6a-457">Requirement</span></span>| <span data-ttu-id="bac6a-458">Wert</span><span class="sxs-lookup"><span data-stu-id="bac6a-458">Value</span></span>|
|---|---|
|[<span data-ttu-id="bac6a-459">Mindestversion des Postfachanforderungssatzes</span><span class="sxs-lookup"><span data-stu-id="bac6a-459">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="bac6a-460">1.0</span><span class="sxs-lookup"><span data-stu-id="bac6a-460">1.0</span></span>|
|[<span data-ttu-id="bac6a-461">Mindestberechtigungsstufe</span><span class="sxs-lookup"><span data-stu-id="bac6a-461">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="bac6a-462">ReadItem</span><span class="sxs-lookup"><span data-stu-id="bac6a-462">ReadItem</span></span>|
|[<span data-ttu-id="bac6a-463">Zutreffender Outlook-Modus</span><span class="sxs-lookup"><span data-stu-id="bac6a-463">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="bac6a-464">Lesen</span><span class="sxs-lookup"><span data-stu-id="bac6a-464">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="bac6a-465">Beispiel</span><span class="sxs-lookup"><span data-stu-id="bac6a-465">Example</span></span>

```
var normalizedSubject = Office.context.mailbox.item.normalizedSubject;
```

####  <a name="notificationmessages-notificationmessagesjavascriptapioutlook16officenotificationmessages"></a><span data-ttu-id="bac6a-466">notificationMessages :[NotificationMessages](/javascript/api/outlook_1_6/office.notificationmessages)</span><span class="sxs-lookup"><span data-stu-id="bac6a-466">notificationMessages :[NotificationMessages](/javascript/api/outlook_1_6/office.notificationmessages)</span></span>

<span data-ttu-id="bac6a-467">Ruft die Benachrichtigungen für ein Element ab.</span><span class="sxs-lookup"><span data-stu-id="bac6a-467">Gets the notification messages for an item.</span></span>

##### <a name="type"></a><span data-ttu-id="bac6a-468">Typ:</span><span class="sxs-lookup"><span data-stu-id="bac6a-468">Type:</span></span>

*   [<span data-ttu-id="bac6a-469">NotificationMessages</span><span class="sxs-lookup"><span data-stu-id="bac6a-469">NotificationMessages</span></span>](/javascript/api/outlook_1_6/office.notificationmessages)

##### <a name="requirements"></a><span data-ttu-id="bac6a-470">Anforderungen</span><span class="sxs-lookup"><span data-stu-id="bac6a-470">Requirements</span></span>

|<span data-ttu-id="bac6a-471">Anforderung</span><span class="sxs-lookup"><span data-stu-id="bac6a-471">Requirement</span></span>| <span data-ttu-id="bac6a-472">Wert</span><span class="sxs-lookup"><span data-stu-id="bac6a-472">Value</span></span>|
|---|---|
|[<span data-ttu-id="bac6a-473">Mindestversion des Postfachanforderungssatzes</span><span class="sxs-lookup"><span data-stu-id="bac6a-473">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="bac6a-474">1.3</span><span class="sxs-lookup"><span data-stu-id="bac6a-474">1.3</span></span>|
|[<span data-ttu-id="bac6a-475">Mindestberechtigungsstufe</span><span class="sxs-lookup"><span data-stu-id="bac6a-475">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="bac6a-476">ReadItem</span><span class="sxs-lookup"><span data-stu-id="bac6a-476">ReadItem</span></span>|
|[<span data-ttu-id="bac6a-477">Zutreffender Outlook-Modus</span><span class="sxs-lookup"><span data-stu-id="bac6a-477">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="bac6a-478">Verfassen oder Lesen</span><span class="sxs-lookup"><span data-stu-id="bac6a-478">Compose or read</span></span>|

####  <a name="optionalattendees-arrayemailaddressdetailsjavascriptapioutlook16officeemailaddressdetailsrecipientsjavascriptapioutlook16officerecipients"></a><span data-ttu-id="bac6a-479">optionalAttendees :Array.<[EmailAddressDetails](/javascript/api/outlook_1_6/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook_1_6/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="bac6a-479">optionalAttendees :Array.<[EmailAddressDetails](/javascript/api/outlook_1_6/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook_1_6/office.recipients)</span></span>

<span data-ttu-id="bac6a-480">Ermöglicht den Zugriff auf die optionalen Teilnehmer eines Ereignisses.</span><span class="sxs-lookup"><span data-stu-id="bac6a-480">Provides access to the optional attendees of an event.</span></span> <span data-ttu-id="bac6a-481">Der Typ des Objekts und die Zugriffsebene, hängt von den Modus des aktuellen Elements ab.</span><span class="sxs-lookup"><span data-stu-id="bac6a-481">The type of object and level of access depends on the mode of the current item.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="bac6a-482">Lesemodus</span><span class="sxs-lookup"><span data-stu-id="bac6a-482">Read mode</span></span>

<span data-ttu-id="bac6a-483">Die `optionalAttendees`-Eigenschaft gibt ein Array mit einem `EmailAddressDetails`-Objekt für jeden optionalen Teilnehmer an der Besprechung zurück.</span><span class="sxs-lookup"><span data-stu-id="bac6a-483">The `optionalAttendees` property returns an array that contains an `EmailAddressDetails` object for each optional attendee to the meeting.</span></span>

##### <a name="compose-mode"></a><span data-ttu-id="bac6a-484">Verfassenmodus</span><span class="sxs-lookup"><span data-stu-id="bac6a-484">Compose mode</span></span>

<span data-ttu-id="bac6a-485">Die `optionalAttendees` -Eigenschaft gibt eine `Recipients` -Objekt, das Methoden zum Abrufen oder Aktualisieren der optionalen Teilnehmer für eine Besprechung bereitstellt.</span><span class="sxs-lookup"><span data-stu-id="bac6a-485">The `optionalAttendees` property returns a `Recipients` object that provides methods to get or update the optional attendees for a meeting.</span></span>

##### <a name="type"></a><span data-ttu-id="bac6a-486">Typ:</span><span class="sxs-lookup"><span data-stu-id="bac6a-486">Type:</span></span>

*   <span data-ttu-id="bac6a-487">Array.<[EmailAddressDetails](/javascript/api/outlook_1_6/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="bac6a-487">Array.<[EmailAddressDetails](/javascript/api/outlook_1_6/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook/office.recipients)</span></span>

##### <a name="requirements"></a><span data-ttu-id="bac6a-488">Anforderungen</span><span class="sxs-lookup"><span data-stu-id="bac6a-488">Requirements</span></span>

|<span data-ttu-id="bac6a-489">Anforderung</span><span class="sxs-lookup"><span data-stu-id="bac6a-489">Requirement</span></span>| <span data-ttu-id="bac6a-490">Wert</span><span class="sxs-lookup"><span data-stu-id="bac6a-490">Value</span></span>|
|---|---|
|[<span data-ttu-id="bac6a-491">Mindestversion des Postfachanforderungssatzes</span><span class="sxs-lookup"><span data-stu-id="bac6a-491">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="bac6a-492">1.0</span><span class="sxs-lookup"><span data-stu-id="bac6a-492">1.0</span></span>|
|[<span data-ttu-id="bac6a-493">Mindestberechtigungsstufe</span><span class="sxs-lookup"><span data-stu-id="bac6a-493">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="bac6a-494">ReadItem</span><span class="sxs-lookup"><span data-stu-id="bac6a-494">ReadItem</span></span>|
|[<span data-ttu-id="bac6a-495">Zutreffender Outlook-Modus</span><span class="sxs-lookup"><span data-stu-id="bac6a-495">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="bac6a-496">Verfassen oder Lesen</span><span class="sxs-lookup"><span data-stu-id="bac6a-496">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="bac6a-497">Beispiel</span><span class="sxs-lookup"><span data-stu-id="bac6a-497">Example</span></span>

```
Office.context.mailbox.item.optionalAttendees.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.optionalAttendees.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.optionalAttendees.getAsync(callback);

function callback(asyncResult) {
  var arrayOfOptionalAttendeesRecipients = asyncResult.value;
}
```

#### <a name="organizer-emailaddressdetailsjavascriptapioutlook16officeemailaddressdetails"></a><span data-ttu-id="bac6a-498">organizer :[EmailAddressDetails](/javascript/api/outlook_1_6/office.emailaddressdetails)</span><span class="sxs-lookup"><span data-stu-id="bac6a-498">organizer :[EmailAddressDetails](/javascript/api/outlook_1_6/office.emailaddressdetails)</span></span>

<span data-ttu-id="bac6a-p124">Ruft für eine angegebene Besprechung die E-Mail-Adresse des Organisators der Besprechung ab. Nur Lesemodus.</span><span class="sxs-lookup"><span data-stu-id="bac6a-p124">Gets the email address of the meeting organizer for a specified meeting. Read mode only.</span></span>

##### <a name="type"></a><span data-ttu-id="bac6a-501">Typ:</span><span class="sxs-lookup"><span data-stu-id="bac6a-501">Type:</span></span>

*   [<span data-ttu-id="bac6a-502">EmailAddressDetails</span><span class="sxs-lookup"><span data-stu-id="bac6a-502">EmailAddressDetails</span></span>](/javascript/api/outlook_1_6/office.emailaddressdetails)

##### <a name="requirements"></a><span data-ttu-id="bac6a-503">Anforderungen</span><span class="sxs-lookup"><span data-stu-id="bac6a-503">Requirements</span></span>

|<span data-ttu-id="bac6a-504">Anforderung</span><span class="sxs-lookup"><span data-stu-id="bac6a-504">Requirement</span></span>| <span data-ttu-id="bac6a-505">Wert</span><span class="sxs-lookup"><span data-stu-id="bac6a-505">Value</span></span>|
|---|---|
|[<span data-ttu-id="bac6a-506">Mindestversion des Postfachanforderungssatzes</span><span class="sxs-lookup"><span data-stu-id="bac6a-506">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="bac6a-507">1.0</span><span class="sxs-lookup"><span data-stu-id="bac6a-507">1.0</span></span>|
|[<span data-ttu-id="bac6a-508">Mindestberechtigungsstufe</span><span class="sxs-lookup"><span data-stu-id="bac6a-508">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="bac6a-509">ReadItem</span><span class="sxs-lookup"><span data-stu-id="bac6a-509">ReadItem</span></span>|
|[<span data-ttu-id="bac6a-510">Zutreffender Outlook-Modus</span><span class="sxs-lookup"><span data-stu-id="bac6a-510">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="bac6a-511">Lesen</span><span class="sxs-lookup"><span data-stu-id="bac6a-511">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="bac6a-512">Beispiel</span><span class="sxs-lookup"><span data-stu-id="bac6a-512">Example</span></span>

```
var organizerName = Office.context.mailbox.item.organizer.displayName;
var organizerAddress = Office.context.mailbox.item.organizer.emailAddress;
```

####  <a name="requiredattendees-arrayemailaddressdetailsjavascriptapioutlook16officeemailaddressdetailsrecipientsjavascriptapioutlook16officerecipients"></a><span data-ttu-id="bac6a-513">requiredAttendees :Array.<[EmailAddressDetails](/javascript/api/outlook_1_6/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook_1_6/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="bac6a-513">requiredAttendees :Array.<[EmailAddressDetails](/javascript/api/outlook_1_6/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook_1_6/office.recipients)</span></span>

<span data-ttu-id="bac6a-514">Ermöglicht den Zugriff auf die erforderlichen Teilnehmer eines Ereignisses.</span><span class="sxs-lookup"><span data-stu-id="bac6a-514">Provides access to the required attendees of an event.</span></span> <span data-ttu-id="bac6a-515">Der Typ des Objekts und die Zugriffsebene, hängt von den Modus des aktuellen Elements ab.</span><span class="sxs-lookup"><span data-stu-id="bac6a-515">The type of object and level of access depends on the mode of the current item.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="bac6a-516">Lesemodus</span><span class="sxs-lookup"><span data-stu-id="bac6a-516">Read mode</span></span>

<span data-ttu-id="bac6a-517">Die `requiredAttendees`-Eigenschaft gibt ein Array mit einem `EmailAddressDetails`-Objekt für jeden erforderlichen Teilnehmer an der Besprechung zurück.</span><span class="sxs-lookup"><span data-stu-id="bac6a-517">The `requiredAttendees` property returns an array that contains an `EmailAddressDetails` object for each required attendee to the meeting.</span></span>

##### <a name="compose-mode"></a><span data-ttu-id="bac6a-518">Verfassenmodus</span><span class="sxs-lookup"><span data-stu-id="bac6a-518">Compose mode</span></span>

<span data-ttu-id="bac6a-519">Die `requiredAttendees` -Eigenschaft gibt eine `Recipients` -Objekt, das Methoden zum Abrufen oder Aktualisieren der erforderlichen Teilnehmer für eine Besprechung bereitstellt.</span><span class="sxs-lookup"><span data-stu-id="bac6a-519">The `requiredAttendees` property returns a `Recipients` object that provides methods to get or update the required attendees for a meeting.</span></span>

##### <a name="type"></a><span data-ttu-id="bac6a-520">Typ:</span><span class="sxs-lookup"><span data-stu-id="bac6a-520">Type:</span></span>

*   <span data-ttu-id="bac6a-521">Array.<[EmailAddressDetails](/javascript/api/outlook_1_6/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook_1_6/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="bac6a-521">Array.<[EmailAddressDetails](/javascript/api/outlook_1_6/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook_1_6/office.recipients)</span></span>

##### <a name="requirements"></a><span data-ttu-id="bac6a-522">Anforderungen</span><span class="sxs-lookup"><span data-stu-id="bac6a-522">Requirements</span></span>

|<span data-ttu-id="bac6a-523">Anforderung</span><span class="sxs-lookup"><span data-stu-id="bac6a-523">Requirement</span></span>| <span data-ttu-id="bac6a-524">Wert</span><span class="sxs-lookup"><span data-stu-id="bac6a-524">Value</span></span>|
|---|---|
|[<span data-ttu-id="bac6a-525">Mindestversion des Postfachanforderungssatzes</span><span class="sxs-lookup"><span data-stu-id="bac6a-525">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="bac6a-526">1.0</span><span class="sxs-lookup"><span data-stu-id="bac6a-526">1.0</span></span>|
|[<span data-ttu-id="bac6a-527">Mindestberechtigungsstufe</span><span class="sxs-lookup"><span data-stu-id="bac6a-527">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="bac6a-528">ReadItem</span><span class="sxs-lookup"><span data-stu-id="bac6a-528">ReadItem</span></span>|
|[<span data-ttu-id="bac6a-529">Zutreffender Outlook-Modus</span><span class="sxs-lookup"><span data-stu-id="bac6a-529">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="bac6a-530">Verfassen oder Lesen</span><span class="sxs-lookup"><span data-stu-id="bac6a-530">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="bac6a-531">Beispiel</span><span class="sxs-lookup"><span data-stu-id="bac6a-531">Example</span></span>

```
Office.context.mailbox.item.requiredAttendees.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.requiredAttendees.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.requiredAttendees.getAsync(callback);

function callback(asyncResult) {
  var arrayOfRequiredAttendeesRecipients = asyncResult.value;
}
```

#### <a name="sender-emailaddressdetailsjavascriptapioutlook16officeemailaddressdetails"></a><span data-ttu-id="bac6a-532">sender :[EmailAddressDetails](/javascript/api/outlook_1_6/office.emailaddressdetails)</span><span class="sxs-lookup"><span data-stu-id="bac6a-532">sender :[EmailAddressDetails](/javascript/api/outlook_1_6/office.emailaddressdetails)</span></span>

<span data-ttu-id="bac6a-p126">Ruft die E-Mail-Adresse des Absenders einer E-Mail-Nachricht ab. Nur Lesemodus.</span><span class="sxs-lookup"><span data-stu-id="bac6a-p126">Gets the email address of the sender of an email message. Read mode only.</span></span>

<span data-ttu-id="bac6a-p127">Die [`from`](#from-emailaddressdetailsjavascriptapioutlook16officeemailaddressdetails)- und `sender`-Eigenschaften stellen dieselbe Person dar, außer die Nachricht wurde von einem Delegaten gesendet. In diesem Fall stellt die `from`-Eigenschaft die Stellvertretung dar, und die sender-Eigenschaft stellt den Delegaten dar.</span><span class="sxs-lookup"><span data-stu-id="bac6a-p127">The [`from`](#from-emailaddressdetailsjavascriptapioutlook16officeemailaddressdetails) and `sender` properties represent the same person unless the message is sent by a delegate. In that case, the `from` property represents the delegator, and the sender property represents the delegate.</span></span>

> [!NOTE]
> <span data-ttu-id="bac6a-537">Die `recipientType` -Eigenschaft des der `EmailAddressDetails` -Objekts der `sender` -Eigenschaft ist `undefined`.</span><span class="sxs-lookup"><span data-stu-id="bac6a-537">The `recipientType` property of the `EmailAddressDetails` object in the `sender` property is `undefined`.</span></span>

##### <a name="type"></a><span data-ttu-id="bac6a-538">Typ:</span><span class="sxs-lookup"><span data-stu-id="bac6a-538">Type:</span></span>

*   [<span data-ttu-id="bac6a-539">EmailAddressDetails</span><span class="sxs-lookup"><span data-stu-id="bac6a-539">EmailAddressDetails</span></span>](/javascript/api/outlook_1_6/office.emailaddressdetails)

##### <a name="requirements"></a><span data-ttu-id="bac6a-540">Anforderungen</span><span class="sxs-lookup"><span data-stu-id="bac6a-540">Requirements</span></span>

|<span data-ttu-id="bac6a-541">Anforderung</span><span class="sxs-lookup"><span data-stu-id="bac6a-541">Requirement</span></span>| <span data-ttu-id="bac6a-542">Wert</span><span class="sxs-lookup"><span data-stu-id="bac6a-542">Value</span></span>|
|---|---|
|[<span data-ttu-id="bac6a-543">Mindestversion des Postfachanforderungssatzes</span><span class="sxs-lookup"><span data-stu-id="bac6a-543">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="bac6a-544">1.0</span><span class="sxs-lookup"><span data-stu-id="bac6a-544">1.0</span></span>|
|[<span data-ttu-id="bac6a-545">Mindestberechtigungsstufe</span><span class="sxs-lookup"><span data-stu-id="bac6a-545">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="bac6a-546">ReadItem</span><span class="sxs-lookup"><span data-stu-id="bac6a-546">ReadItem</span></span>|
|[<span data-ttu-id="bac6a-547">Zutreffender Outlook-Modus</span><span class="sxs-lookup"><span data-stu-id="bac6a-547">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="bac6a-548">Lesen</span><span class="sxs-lookup"><span data-stu-id="bac6a-548">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="bac6a-549">Beispiel</span><span class="sxs-lookup"><span data-stu-id="bac6a-549">Example</span></span>

```
var senderName = Office.context.mailbox.item.sender.displayName;
var senderAddress = Office.context.mailbox.item.sender.emailAddress;
```

####  <a name="start-datetimejavascriptapioutlook16officetime"></a><span data-ttu-id="bac6a-550">start :Date|[Time](/javascript/api/outlook_1_6/office.time)</span><span class="sxs-lookup"><span data-stu-id="bac6a-550">start :Date|[Time](/javascript/api/outlook_1_6/office.time)</span></span>

<span data-ttu-id="bac6a-551">Ruft Datum und Zeit für den Beginn des Termins ab oder legt Datum und Uhrzeit fest.</span><span class="sxs-lookup"><span data-stu-id="bac6a-551">Gets or sets the date and time that the appointment is to begin.</span></span>

<span data-ttu-id="bac6a-p128">Die `start`-Eigenschaft wird als Datums- und Uhrzeitwert in UTC ausgedrückt. Sie können die [`convertToLocalClientTime`](office.context.mailbox.md#converttolocalclienttimetimevalue--localclienttimejavascriptapioutlook16officelocalclienttime)-Methode verwenden, um den Wert in den lokalen Uhrzeit- und Datumswert des Clients umzuwandeln.</span><span class="sxs-lookup"><span data-stu-id="bac6a-p128">The `start` property is expressed as a Coordinated Universal Time (UTC) date and time value. You can use the [`convertToLocalClientTime`](office.context.mailbox.md#converttolocalclienttimetimevalue--localclienttimejavascriptapioutlook16officelocalclienttime) method to convert the value to the client’s local date and time.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="bac6a-554">Lesemodus</span><span class="sxs-lookup"><span data-stu-id="bac6a-554">Read mode</span></span>

<span data-ttu-id="bac6a-555">Die `start`-Eigenschaft gibt ein `Date`-Objekt zurück.</span><span class="sxs-lookup"><span data-stu-id="bac6a-555">The `start` property returns a `Date` object.</span></span>

##### <a name="compose-mode"></a><span data-ttu-id="bac6a-556">Verfassenmodus</span><span class="sxs-lookup"><span data-stu-id="bac6a-556">Compose mode</span></span>

<span data-ttu-id="bac6a-557">Die `start`-Eigenschaft gibt ein `Time`-Objekt zurück.</span><span class="sxs-lookup"><span data-stu-id="bac6a-557">The `start` property returns a `Time` object.</span></span>

<span data-ttu-id="bac6a-558">Wenn Sie die [`Time.setAsync`](/javascript/api/outlook_1_6/office.time#setasync-datetime--options--callback-)-Methode verwenden, um die Startzeit im Verfassenmodus festzulegen, sollten Sie die [`convertToUtcClientTime`](office.context.mailbox.md#converttoutcclienttimeinput--date)-Methode verwenden, um die Ortszeit auf dem Client für den Server in UTC umzuwandeln.</span><span class="sxs-lookup"><span data-stu-id="bac6a-558">When you use the [`Time.setAsync`](/javascript/api/outlook_1_6/office.time#setasync-datetime--options--callback-) method to set the start time, you should use the [`convertToUtcClientTime`](office.context.mailbox.md#converttoutcclienttimeinput--date) method to convert the local time on the client to UTC for the server.</span></span>

##### <a name="type"></a><span data-ttu-id="bac6a-559">Typ:</span><span class="sxs-lookup"><span data-stu-id="bac6a-559">Type:</span></span>

*   <span data-ttu-id="bac6a-560">Date | [Time](/javascript/api/outlook_1_6/office.time)</span><span class="sxs-lookup"><span data-stu-id="bac6a-560">Date | [Time](/javascript/api/outlook_1_6/office.time)</span></span>

##### <a name="requirements"></a><span data-ttu-id="bac6a-561">Anforderungen</span><span class="sxs-lookup"><span data-stu-id="bac6a-561">Requirements</span></span>

|<span data-ttu-id="bac6a-562">Anforderung</span><span class="sxs-lookup"><span data-stu-id="bac6a-562">Requirement</span></span>| <span data-ttu-id="bac6a-563">Wert</span><span class="sxs-lookup"><span data-stu-id="bac6a-563">Value</span></span>|
|---|---|
|[<span data-ttu-id="bac6a-564">Mindestversion des Postfachanforderungssatzes</span><span class="sxs-lookup"><span data-stu-id="bac6a-564">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="bac6a-565">1.0</span><span class="sxs-lookup"><span data-stu-id="bac6a-565">1.0</span></span>|
|[<span data-ttu-id="bac6a-566">Mindestberechtigungsstufe</span><span class="sxs-lookup"><span data-stu-id="bac6a-566">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="bac6a-567">ReadItem</span><span class="sxs-lookup"><span data-stu-id="bac6a-567">ReadItem</span></span>|
|[<span data-ttu-id="bac6a-568">Zutreffender Outlook-Modus</span><span class="sxs-lookup"><span data-stu-id="bac6a-568">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="bac6a-569">Verfassen oder Lesen</span><span class="sxs-lookup"><span data-stu-id="bac6a-569">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="bac6a-570">Beispiel</span><span class="sxs-lookup"><span data-stu-id="bac6a-570">Example</span></span>

<span data-ttu-id="bac6a-571">Im folgenden Beispiel wird die Startzeit eines Termins im Verfassenmodus mithilfe der [`setAsync`](/javascript/api/outlook_1_6/office.time#setasync-datetime--options--callback-)-Methode des `Time`-Objekts festgelegt.</span><span class="sxs-lookup"><span data-stu-id="bac6a-571">The following example sets the start time of an appointment in compose mode by using the [`setAsync`](/javascript/api/outlook_1_6/office.time#setasync-datetime--options--callback-) method of the `Time` object.</span></span>

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

####  <a name="subject-stringsubjectjavascriptapioutlook16officesubject"></a><span data-ttu-id="bac6a-572">subject :String|[Subject](/javascript/api/outlook_1_6/office.subject)</span><span class="sxs-lookup"><span data-stu-id="bac6a-572">subject :String|[Subject](/javascript/api/outlook_1_6/office.subject)</span></span>

<span data-ttu-id="bac6a-573">Ruft die Beschreibung ab, die im Betrefffeld eines Elements angezeigt wird, oder legt sie fest.</span><span class="sxs-lookup"><span data-stu-id="bac6a-573">Gets or sets the description that appears in the subject field of an item.</span></span>

<span data-ttu-id="bac6a-574">Die `subject`-Eigenschaft ruft den gesamten Betreff des Elements ab oder legt ihn fest – so, wie er vom E-Mail-Server gesendet wird.</span><span class="sxs-lookup"><span data-stu-id="bac6a-574">The `subject` property gets or sets the entire subject of the item, as sent by the email server.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="bac6a-575">Lesemodus</span><span class="sxs-lookup"><span data-stu-id="bac6a-575">Read mode</span></span>

<span data-ttu-id="bac6a-p129">Die `subject`-Eigenschaft gibt eine Zeichenfolge zurück. Verwenden Sie die [`normalizedSubject`](#normalizedsubject-string)-Eigenschaft, um den Betreff ohne führende Präfixe wie `RE:` und `FW:` abzurufen.</span><span class="sxs-lookup"><span data-stu-id="bac6a-p129">The `subject` property returns a string. Use the [`normalizedSubject`](#normalizedsubject-string) property to get the subject minus any leading prefixes such as `RE:` and `FW:`.</span></span>

```
var subject = Office.context.mailbox.item.subject;
```

##### <a name="compose-mode"></a><span data-ttu-id="bac6a-578">Verfassenmodus</span><span class="sxs-lookup"><span data-stu-id="bac6a-578">Compose mode</span></span>

<span data-ttu-id="bac6a-579">Die `subject`-Eigenschaft gibt ein `Subject`-Objekt zurück, das Methoden zum Abrufen und Festlegen des Betreffs bereitstellt.</span><span class="sxs-lookup"><span data-stu-id="bac6a-579">The `subject` property returns a `Subject` object that provides methods to get and set the subject.</span></span>

```
Office.context.mailbox.item.subject.getAsync(callback);

function callback(asyncResult) {
  var subject = asyncResult.value;
}
```

##### <a name="type"></a><span data-ttu-id="bac6a-580">Typ:</span><span class="sxs-lookup"><span data-stu-id="bac6a-580">Type:</span></span>

*   <span data-ttu-id="bac6a-581">String | [Subject](/javascript/api/outlook_1_6/office.subject)</span><span class="sxs-lookup"><span data-stu-id="bac6a-581">String | [Subject](/javascript/api/outlook_1_6/office.subject)</span></span>

##### <a name="requirements"></a><span data-ttu-id="bac6a-582">Anforderungen</span><span class="sxs-lookup"><span data-stu-id="bac6a-582">Requirements</span></span>

|<span data-ttu-id="bac6a-583">Anforderung</span><span class="sxs-lookup"><span data-stu-id="bac6a-583">Requirement</span></span>| <span data-ttu-id="bac6a-584">Wert</span><span class="sxs-lookup"><span data-stu-id="bac6a-584">Value</span></span>|
|---|---|
|[<span data-ttu-id="bac6a-585">Mindestversion des Postfachanforderungssatzes</span><span class="sxs-lookup"><span data-stu-id="bac6a-585">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="bac6a-586">1.0</span><span class="sxs-lookup"><span data-stu-id="bac6a-586">1.0</span></span>|
|[<span data-ttu-id="bac6a-587">Mindestberechtigungsstufe</span><span class="sxs-lookup"><span data-stu-id="bac6a-587">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="bac6a-588">ReadItem</span><span class="sxs-lookup"><span data-stu-id="bac6a-588">ReadItem</span></span>|
|[<span data-ttu-id="bac6a-589">Zutreffender Outlook-Modus</span><span class="sxs-lookup"><span data-stu-id="bac6a-589">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="bac6a-590">Verfassen oder Lesen</span><span class="sxs-lookup"><span data-stu-id="bac6a-590">Compose or read</span></span>|

####  <a name="to-arrayemailaddressdetailsjavascriptapioutlook16officeemailaddressdetailsrecipientsjavascriptapioutlook16officerecipients"></a><span data-ttu-id="bac6a-591">an: Array. <[EmailAddressDetails](/javascript/api/outlook_1_6/office.emailaddressdetails)>|[Empfänger](/javascript/api/outlook_1_6/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="bac6a-591">to :Array.<[EmailAddressDetails](/javascript/api/outlook_1_6/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook_1_6/office.recipients)</span></span>

<span data-ttu-id="bac6a-592">Ermöglicht den Zugriff auf die Empfänger in der Zeile **an** einer Nachricht.</span><span class="sxs-lookup"><span data-stu-id="bac6a-592">Provides access to the recipients on the **To** line of a message.</span></span> <span data-ttu-id="bac6a-593">Der Typ des Objekts und die Zugriffsebene, hängt von den Modus des aktuellen Elements ab.</span><span class="sxs-lookup"><span data-stu-id="bac6a-593">The type of object and level of access depends on the mode of the current item.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="bac6a-594">Lesemodus</span><span class="sxs-lookup"><span data-stu-id="bac6a-594">Read mode</span></span>

<span data-ttu-id="bac6a-p131">Die `to`-Eigenschaft gibt ein Array mit einem `EmailAddressDetails`-Objekt für jeden Empfänger in der**An**-Zeile der Nachricht zurück. Die Auflistung ist auf höchstens 100 Elemente beschränkt.</span><span class="sxs-lookup"><span data-stu-id="bac6a-p131">The `to` property returns an array that contains an `EmailAddressDetails` object for each recipient listed on the **To** line of the message. The collection is limited to a maximum of 100 members.</span></span>

##### <a name="compose-mode"></a><span data-ttu-id="bac6a-597">Verfassenmodus</span><span class="sxs-lookup"><span data-stu-id="bac6a-597">Compose mode</span></span>

<span data-ttu-id="bac6a-598">Die `to` -Eigenschaft gibt eine `Recipients` -Objekt, das Methoden zum Abrufen oder aktualisieren die Empfänger in der Zeile **an** der Nachricht bereitstellt.</span><span class="sxs-lookup"><span data-stu-id="bac6a-598">The `to` property returns a `Recipients` object that provides methods to get or update the recipients on the **To** line of the message.</span></span>

##### <a name="type"></a><span data-ttu-id="bac6a-599">Typ:</span><span class="sxs-lookup"><span data-stu-id="bac6a-599">Type:</span></span>

*   <span data-ttu-id="bac6a-600">Array.<[EmailAddressDetails](/javascript/api/outlook_1_6/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook_1_6/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="bac6a-600">Array.<[EmailAddressDetails](/javascript/api/outlook_1_6/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook_1_6/office.recipients)</span></span>

##### <a name="requirements"></a><span data-ttu-id="bac6a-601">Anforderungen</span><span class="sxs-lookup"><span data-stu-id="bac6a-601">Requirements</span></span>

|<span data-ttu-id="bac6a-602">Anforderung</span><span class="sxs-lookup"><span data-stu-id="bac6a-602">Requirement</span></span>| <span data-ttu-id="bac6a-603">Wert</span><span class="sxs-lookup"><span data-stu-id="bac6a-603">Value</span></span>|
|---|---|
|[<span data-ttu-id="bac6a-604">Mindestversion des Postfachanforderungssatzes</span><span class="sxs-lookup"><span data-stu-id="bac6a-604">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="bac6a-605">1.0</span><span class="sxs-lookup"><span data-stu-id="bac6a-605">1.0</span></span>|
|[<span data-ttu-id="bac6a-606">Mindestberechtigungsstufe</span><span class="sxs-lookup"><span data-stu-id="bac6a-606">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="bac6a-607">ReadItem</span><span class="sxs-lookup"><span data-stu-id="bac6a-607">ReadItem</span></span>|
|[<span data-ttu-id="bac6a-608">Zutreffender Outlook-Modus</span><span class="sxs-lookup"><span data-stu-id="bac6a-608">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="bac6a-609">Verfassen oder Lesen</span><span class="sxs-lookup"><span data-stu-id="bac6a-609">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="bac6a-610">Beispiel</span><span class="sxs-lookup"><span data-stu-id="bac6a-610">Example</span></span>

```
Office.context.mailbox.item.to.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.to.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.to.getAsync(callback);

function callback(asyncResult) {
  var arrayOfToRecipients = asyncResult.value;
}
```

### <a name="methods"></a><span data-ttu-id="bac6a-611">Methoden</span><span class="sxs-lookup"><span data-stu-id="bac6a-611">Methods</span></span>

####  <a name="addfileattachmentasyncuri-attachmentname-options-callback"></a><span data-ttu-id="bac6a-612">addFileAttachmentAsync(uri, attachmentName, [options], [callback])</span><span class="sxs-lookup"><span data-stu-id="bac6a-612">addFileAttachmentAsync(uri, attachmentName, [options], [callback])</span></span>

<span data-ttu-id="bac6a-613">Fügt eine Datei zu einer Nachricht oder einem Termin als Anlage hinzu.</span><span class="sxs-lookup"><span data-stu-id="bac6a-613">Adds a file to a message or appointment as an attachment.</span></span>

<span data-ttu-id="bac6a-614">Die `addFileAttachmentAsync`-Methode lädt die Datei am angegebenen URI hoch und fügt sie an das Element im Verfassenformular an.</span><span class="sxs-lookup"><span data-stu-id="bac6a-614">The `addFileAttachmentAsync` method uploads the file at the specified URI and attaches it to the item in the compose form.</span></span>

<span data-ttu-id="bac6a-615">Anschließend können Sie den Bezeichner mit der [`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback)-Methode in der gleichen Sitzung zum Entfernen der Anlage verwenden.</span><span class="sxs-lookup"><span data-stu-id="bac6a-615">You can subsequently use the identifier with the [`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback) method to remove the attachment in the same session.</span></span>

##### <a name="parameters"></a><span data-ttu-id="bac6a-616">Parameter:</span><span class="sxs-lookup"><span data-stu-id="bac6a-616">Parameters:</span></span>

|<span data-ttu-id="bac6a-617">Name</span><span class="sxs-lookup"><span data-stu-id="bac6a-617">Name</span></span>| <span data-ttu-id="bac6a-618">Typ</span><span class="sxs-lookup"><span data-stu-id="bac6a-618">Type</span></span>| <span data-ttu-id="bac6a-619">Attribute</span><span class="sxs-lookup"><span data-stu-id="bac6a-619">Attributes</span></span>| <span data-ttu-id="bac6a-620">Beschreibung</span><span class="sxs-lookup"><span data-stu-id="bac6a-620">Description</span></span>|
|---|---|---|---|
|`uri`| <span data-ttu-id="bac6a-621">Zeichenfolge</span><span class="sxs-lookup"><span data-stu-id="bac6a-621">String</span></span>||<span data-ttu-id="bac6a-p132">Der URI, der den Speicherort der an die Nachricht oder den Termin anzuhängenden Datei angibt. Die maximale Länge ist 2048 Zeichen.</span><span class="sxs-lookup"><span data-stu-id="bac6a-p132">The URI that provides the location of the file to attach to the message or appointment. The maximum length is 2048 characters.</span></span>|
|`attachmentName`| <span data-ttu-id="bac6a-624">Zeichenfolge</span><span class="sxs-lookup"><span data-stu-id="bac6a-624">String</span></span>||<span data-ttu-id="bac6a-p133">Der Name der Anlage, der beim Hochladen der Anlage angezeigt wird. Die maximale Länge ist 255 Zeichen.</span><span class="sxs-lookup"><span data-stu-id="bac6a-p133">The name of the attachment that is shown while the attachment is uploading. The maximum length is 255 characters.</span></span>|
|`options`| <span data-ttu-id="bac6a-627">Object</span><span class="sxs-lookup"><span data-stu-id="bac6a-627">Object</span></span>| <span data-ttu-id="bac6a-628">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="bac6a-628">&lt;optional&gt;</span></span>|<span data-ttu-id="bac6a-629">Ein Objektliteral, das mindestens eine der folgenden Eigenschaften enthält.</span><span class="sxs-lookup"><span data-stu-id="bac6a-629">An object literal that contains one or more of the following properties.</span></span>|
| `options.asyncContext` | <span data-ttu-id="bac6a-630">Object</span><span class="sxs-lookup"><span data-stu-id="bac6a-630">Object</span></span> | <span data-ttu-id="bac6a-631">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="bac6a-631">&lt;optional&gt;</span></span> | <span data-ttu-id="bac6a-632">Entwickler können jedes Objekt angeben, auf das sie in der Rückrufmethode zugreifen möchten.</span><span class="sxs-lookup"><span data-stu-id="bac6a-632">Developers can provide any object they wish to access in the callback method.</span></span> |
| `options.isInline` | <span data-ttu-id="bac6a-633">Boolean</span><span class="sxs-lookup"><span data-stu-id="bac6a-633">Boolean</span></span> | <span data-ttu-id="bac6a-634">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="bac6a-634">&lt;optional&gt;</span></span> | <span data-ttu-id="bac6a-635">Wenn `true`: Zeigt an, dass die Anlage inline im Nachrichtentext angezeigt wird und nicht in der Anlagenlisten angezeigt werden soll.</span><span class="sxs-lookup"><span data-stu-id="bac6a-635">If `true`, indicates that the attachment will be shown inline in the message body, and should not be displayed in the attachment list.</span></span> |
|`callback`| <span data-ttu-id="bac6a-636">function</span><span class="sxs-lookup"><span data-stu-id="bac6a-636">function</span></span>| <span data-ttu-id="bac6a-637">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="bac6a-637">&lt;optional&gt;</span></span>|<span data-ttu-id="bac6a-638">Wenn die Methode abgeschlossen wird, wird die Funktion , die im `callback`-Parameter übergeben wird, mit einem einzelnen Parameter (`asyncResult`) aufgerufen, der ein [`AsyncResult`](/javascript/api/office/office.asyncresult)-Objekt darstellt.</span><span class="sxs-lookup"><span data-stu-id="bac6a-638">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span> <br/><span data-ttu-id="bac6a-639">Bei Erfolg wird der Anlagenbezeichner in der `asyncResult.value`-Eigenschaft bereitgestellt.</span><span class="sxs-lookup"><span data-stu-id="bac6a-639">On success, the attachment identifier will be provided in the `asyncResult.value` property.</span></span><br/><span data-ttu-id="bac6a-640">Wenn beim Hochladen der Anlage ein Fehler auftritt, enthält das `asyncResult`-Objekt ein `Error`-Objekt mit einer Beschreibung des Fehlers.</span><span class="sxs-lookup"><span data-stu-id="bac6a-640">If uploading the attachment fails, the `asyncResult` object will contain an `Error` object that provides a description of the error.</span></span>|

##### <a name="errors"></a><span data-ttu-id="bac6a-641">Fehler</span><span class="sxs-lookup"><span data-stu-id="bac6a-641">Errors</span></span>

| <span data-ttu-id="bac6a-642">Fehlercode</span><span class="sxs-lookup"><span data-stu-id="bac6a-642">Error code</span></span> | <span data-ttu-id="bac6a-643">Beschreibung</span><span class="sxs-lookup"><span data-stu-id="bac6a-643">Description</span></span> |
|------------|-------------|
| `AttachmentSizeExceeded` | <span data-ttu-id="bac6a-644">Die Anlage ist zu groß.</span><span class="sxs-lookup"><span data-stu-id="bac6a-644">The attachment is larger than allowed.</span></span> |
| `FileTypeNotSupported` | <span data-ttu-id="bac6a-645">Die Anlage enthält eine Erweiterung, die nicht zulässig ist.</span><span class="sxs-lookup"><span data-stu-id="bac6a-645">The attachment has an extension that is not allowed.</span></span> |
| `NumberOfAttachmentsExceeded` | <span data-ttu-id="bac6a-646">Die Nachricht oder der Termin enthält zu viele Anlagen.</span><span class="sxs-lookup"><span data-stu-id="bac6a-646">The message or appointment has too many attachments.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="bac6a-647">Anforderungen</span><span class="sxs-lookup"><span data-stu-id="bac6a-647">Requirements</span></span>

|<span data-ttu-id="bac6a-648">Anforderung</span><span class="sxs-lookup"><span data-stu-id="bac6a-648">Requirement</span></span>| <span data-ttu-id="bac6a-649">Wert</span><span class="sxs-lookup"><span data-stu-id="bac6a-649">Value</span></span>|
|---|---|
|[<span data-ttu-id="bac6a-650">Mindestversion des Postfachanforderungssatzes</span><span class="sxs-lookup"><span data-stu-id="bac6a-650">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="bac6a-651">1.1</span><span class="sxs-lookup"><span data-stu-id="bac6a-651">1.1</span></span>|
|[<span data-ttu-id="bac6a-652">Mindestberechtigungsstufe</span><span class="sxs-lookup"><span data-stu-id="bac6a-652">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="bac6a-653">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="bac6a-653">ReadWriteItem</span></span>|
|[<span data-ttu-id="bac6a-654">Zutreffender Outlook-Modus</span><span class="sxs-lookup"><span data-stu-id="bac6a-654">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="bac6a-655">Verfassen</span><span class="sxs-lookup"><span data-stu-id="bac6a-655">Compose</span></span>|

##### <a name="examples"></a><span data-ttu-id="bac6a-656">Beispiele</span><span class="sxs-lookup"><span data-stu-id="bac6a-656">Examples</span></span>

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

<span data-ttu-id="bac6a-657">Im folgenden Beispiel wird eine Bilddatei als Inlineanlage hinzugefügt, und die Anlage wird im Nachrichtentext referenziert.</span><span class="sxs-lookup"><span data-stu-id="bac6a-657">The following example adds an image file as an inline attachment and references the attachment in the message body.</span></span>

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

####  <a name="additemattachmentasyncitemid-attachmentname-options-callback"></a><span data-ttu-id="bac6a-658">addItemAttachmentAsync(itemId, attachmentName, [options], [callback])</span><span class="sxs-lookup"><span data-stu-id="bac6a-658">addItemAttachmentAsync(itemId, attachmentName, [options], [callback])</span></span>

<span data-ttu-id="bac6a-659">Fügt der Nachricht oder dem Termin ein Exchange-Objekt, wie z. B. eine Nachricht, als Anhang hinzu.</span><span class="sxs-lookup"><span data-stu-id="bac6a-659">Adds an Exchange item, such as a message, as an attachment to the message or appointment.</span></span>

<span data-ttu-id="bac6a-p134">Die `addItemAttachmentAsync`-Methode fügt das Element mit dem angegebenen Exchange-Bezeichner an das Element im Formular zum Verfassen an. Wenn Sie eine Rückrufmethode angeben, wird die Methode mit einem `asyncResult`-Parameter aufgerufen, der entweder den Anlagenbezeichner oder einen Code enthält, der etwaige Fehler angibt, die beim Anfügen des Objekts aufgetreten sind. Sie können ggf. den `options`-Parameter verwenden, um Statusinformationen an die Rückrufmethode zu übergeben.</span><span class="sxs-lookup"><span data-stu-id="bac6a-p134">The `addItemAttachmentAsync` method attaches the item with the specified Exchange identifier to the item in the compose form. If you specify a callback method, the method is called with one parameter, `asyncResult`, which contains either the attachment identifier or a code that indicates any error that occurred while attaching the item. You can use the `options` parameter to pass state information to the callback method, if needed.</span></span>

<span data-ttu-id="bac6a-663">Anschließend können Sie den Bezeichner mit der [`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback)-Methode in der gleichen Sitzung zum Entfernen der Anlage verwenden.</span><span class="sxs-lookup"><span data-stu-id="bac6a-663">You can subsequently use the identifier with the [`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback) method to remove the attachment in the same session.</span></span>

<span data-ttu-id="bac6a-664">Wenn Ihre Office-Add-Ins in Outlook Web App ausgeführt wird die `addItemAttachmentAsync` -Methode kann Anfügen von Elementen in der Elemente, die Sie bearbeiten; Allerdings Dies wird nicht unterstützt und wird nicht empfohlen.</span><span class="sxs-lookup"><span data-stu-id="bac6a-664">If your Office Add-in is running in Outlook Web App, the `addItemAttachmentAsync` method can attach items to items other than the item that you are editing; however, this is not supported and is not recommended.</span></span>

##### <a name="parameters"></a><span data-ttu-id="bac6a-665">Parameter:</span><span class="sxs-lookup"><span data-stu-id="bac6a-665">Parameters:</span></span>

|<span data-ttu-id="bac6a-666">Name</span><span class="sxs-lookup"><span data-stu-id="bac6a-666">Name</span></span>| <span data-ttu-id="bac6a-667">Typ</span><span class="sxs-lookup"><span data-stu-id="bac6a-667">Type</span></span>| <span data-ttu-id="bac6a-668">Attribute</span><span class="sxs-lookup"><span data-stu-id="bac6a-668">Attributes</span></span>| <span data-ttu-id="bac6a-669">Beschreibung</span><span class="sxs-lookup"><span data-stu-id="bac6a-669">Description</span></span>|
|---|---|---|---|
|`itemId`| <span data-ttu-id="bac6a-670">Zeichenfolge</span><span class="sxs-lookup"><span data-stu-id="bac6a-670">String</span></span>||<span data-ttu-id="bac6a-p135">Der Exchange-Bezeichner des Objekts, das angehängt werden soll. Die maximale Länge beträgt 100 Zeichen.</span><span class="sxs-lookup"><span data-stu-id="bac6a-p135">The Exchange identifier of the item to attach. The maximum length is 100 characters.</span></span>|
|`attachmentName`| <span data-ttu-id="bac6a-673">Zeichenfolge</span><span class="sxs-lookup"><span data-stu-id="bac6a-673">String</span></span>||<span data-ttu-id="bac6a-p136">Der Betreff des Elements, das angehängt werden soll. Die maximale Länge ist 255 Zeichen.</span><span class="sxs-lookup"><span data-stu-id="bac6a-p136">The sujbect of the item to be attached. The maximum length is 255 characters.</span></span>|
|`options`| <span data-ttu-id="bac6a-676">Object</span><span class="sxs-lookup"><span data-stu-id="bac6a-676">Object</span></span>| <span data-ttu-id="bac6a-677">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="bac6a-677">&lt;optional&gt;</span></span>|<span data-ttu-id="bac6a-678">Ein Objektliteral, das mindestens eine der folgenden Eigenschaften enthält.</span><span class="sxs-lookup"><span data-stu-id="bac6a-678">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`| <span data-ttu-id="bac6a-679">Object</span><span class="sxs-lookup"><span data-stu-id="bac6a-679">Object</span></span>| <span data-ttu-id="bac6a-680">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="bac6a-680">&lt;optional&gt;</span></span>|<span data-ttu-id="bac6a-681">Entwickler können ein Objekt bereitstellen, auf das sie in der Callbackmethode zugreifen möchten.</span><span class="sxs-lookup"><span data-stu-id="bac6a-681">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`| <span data-ttu-id="bac6a-682">Funktion</span><span class="sxs-lookup"><span data-stu-id="bac6a-682">function</span></span>| <span data-ttu-id="bac6a-683">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="bac6a-683">&lt;optional&gt;</span></span>|<span data-ttu-id="bac6a-684">Wenn die Methode abgeschlossen wird, wird die Funktion , die im `callback`-Parameter übergeben wird, mit einem einzelnen Parameter (`asyncResult`) aufgerufen, der ein [`AsyncResult`](/javascript/api/office/office.asyncresult)-Objekt darstellt.</span><span class="sxs-lookup"><span data-stu-id="bac6a-684">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span> <br/><span data-ttu-id="bac6a-685">Bei Erfolg wird der Anlagenbezeichner in der `asyncResult.value`-Eigenschaft bereitgestellt.</span><span class="sxs-lookup"><span data-stu-id="bac6a-685">On success, the attachment identifier will be provided in the `asyncResult.value` property.</span></span><br/><span data-ttu-id="bac6a-686">Wenn beim Hinzufügen der Anlage ein Fehler auftritt, enthält das `asyncResult`-Objekt ein `Error`-Objekt mit einer Beschreibung des Fehlers.</span><span class="sxs-lookup"><span data-stu-id="bac6a-686">If adding the attachment fails, the `asyncResult` object will contain an `Error` object that provides a description of the error.</span></span>|

##### <a name="errors"></a><span data-ttu-id="bac6a-687">Fehler</span><span class="sxs-lookup"><span data-stu-id="bac6a-687">Errors</span></span>

| <span data-ttu-id="bac6a-688">Fehlercode</span><span class="sxs-lookup"><span data-stu-id="bac6a-688">Error code</span></span> | <span data-ttu-id="bac6a-689">Beschreibung</span><span class="sxs-lookup"><span data-stu-id="bac6a-689">Description</span></span> |
|------------|-------------|
| `NumberOfAttachmentsExceeded` | <span data-ttu-id="bac6a-690">Die Nachricht oder der Termin enthält zu viele Anlagen.</span><span class="sxs-lookup"><span data-stu-id="bac6a-690">The message or appointment has too many attachments.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="bac6a-691">Anforderungen</span><span class="sxs-lookup"><span data-stu-id="bac6a-691">Requirements</span></span>

|<span data-ttu-id="bac6a-692">Anforderung</span><span class="sxs-lookup"><span data-stu-id="bac6a-692">Requirement</span></span>| <span data-ttu-id="bac6a-693">Wert</span><span class="sxs-lookup"><span data-stu-id="bac6a-693">Value</span></span>|
|---|---|
|[<span data-ttu-id="bac6a-694">Mindestversion des Postfachanforderungssatzes</span><span class="sxs-lookup"><span data-stu-id="bac6a-694">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="bac6a-695">1.1</span><span class="sxs-lookup"><span data-stu-id="bac6a-695">1.1</span></span>|
|[<span data-ttu-id="bac6a-696">Mindestberechtigungsstufe</span><span class="sxs-lookup"><span data-stu-id="bac6a-696">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="bac6a-697">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="bac6a-697">ReadWriteItem</span></span>|
|[<span data-ttu-id="bac6a-698">Zutreffender Outlook-Modus</span><span class="sxs-lookup"><span data-stu-id="bac6a-698">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="bac6a-699">Verfassen</span><span class="sxs-lookup"><span data-stu-id="bac6a-699">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="bac6a-700">Beispiel</span><span class="sxs-lookup"><span data-stu-id="bac6a-700">Example</span></span>

<span data-ttu-id="bac6a-701">Im folgenden Beispiel wird ein vorhandenes Outlook-Element als Anlage mit dem Namen `My Attachment` hinzugefügt.</span><span class="sxs-lookup"><span data-stu-id="bac6a-701">The following example adds an existing Outlook item as an attachment with the name `My Attachment`.</span></span>

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

####  <a name="close"></a><span data-ttu-id="bac6a-702">close()</span><span class="sxs-lookup"><span data-stu-id="bac6a-702">close()</span></span>

<span data-ttu-id="bac6a-703">Schließt das aktuelle Element, das gerade verfasst wird.</span><span class="sxs-lookup"><span data-stu-id="bac6a-703">Closes the current item that is being composed.</span></span>

<span data-ttu-id="bac6a-p137">Das Verhalten der `close`-Methode hängt vom aktuellen Status des verfassten Elements ab. Wenn das Element über nicht gespeicherte Änderungen verfügt, fordert der Client den Benutzer auf, die Schließenaktion zu speichern, zu verwerfen oder abzubrechen.</span><span class="sxs-lookup"><span data-stu-id="bac6a-p137">The behavior of the `close` method depends on the current state of the item being composed. If the item has unsaved changes, the client prompts the user to save, discard, or cancel the close action.</span></span>

> [!NOTE]
> <span data-ttu-id="bac6a-706">In Outlook im Web, wenn das Element einen Termin handelt, und es wurde bereits zuvor gespeichert wurde mit `saveAsync`, der Benutzer wird aufgefordert, speichern, löschen oder Abbrechen, auch wenn keine Änderungen aufgetreten sind, da das Element zuletzt gespeichert.</span><span class="sxs-lookup"><span data-stu-id="bac6a-706">In Outlook on the web, if the item is an appointment and it has previously been saved using `saveAsync`, the user is prompted to save, discard, or cancel even if no changes have occurred since the item was last saved.</span></span>

<span data-ttu-id="bac6a-707">Wenn die Nachricht im Outlook-Desktopclient eine Inlineantwort ist, hat die `close`-Methode keine Auswirkung.</span><span class="sxs-lookup"><span data-stu-id="bac6a-707">In the Outlook desktop client, if the message is an inline reply, the `close` method has no effect.</span></span>

##### <a name="requirements"></a><span data-ttu-id="bac6a-708">Anforderungen</span><span class="sxs-lookup"><span data-stu-id="bac6a-708">Requirements</span></span>

|<span data-ttu-id="bac6a-709">Anforderung</span><span class="sxs-lookup"><span data-stu-id="bac6a-709">Requirement</span></span>| <span data-ttu-id="bac6a-710">Wert</span><span class="sxs-lookup"><span data-stu-id="bac6a-710">Value</span></span>|
|---|---|
|[<span data-ttu-id="bac6a-711">Mindestversion des Postfachanforderungssatzes</span><span class="sxs-lookup"><span data-stu-id="bac6a-711">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="bac6a-712">1.3</span><span class="sxs-lookup"><span data-stu-id="bac6a-712">1.3</span></span>|
|[<span data-ttu-id="bac6a-713">Mindestberechtigungsstufe</span><span class="sxs-lookup"><span data-stu-id="bac6a-713">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="bac6a-714">Eingeschränkt</span><span class="sxs-lookup"><span data-stu-id="bac6a-714">Restricted</span></span>|
|[<span data-ttu-id="bac6a-715">Zutreffender Outlook-Modus</span><span class="sxs-lookup"><span data-stu-id="bac6a-715">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="bac6a-716">Verfassen</span><span class="sxs-lookup"><span data-stu-id="bac6a-716">Compose</span></span>|

#### <a name="displayreplyallformformdata"></a><span data-ttu-id="bac6a-717">displayReplyAllForm(formData)</span><span class="sxs-lookup"><span data-stu-id="bac6a-717">displayReplyAllForm(formData)</span></span>

<span data-ttu-id="bac6a-718">Zeigt ein Antwortformular an, das den Absender und alle Empfänger der ausgewählten Nachricht oder des Organisators und alle Teilnehmer des ausgewählten Termins enthält.</span><span class="sxs-lookup"><span data-stu-id="bac6a-718">Displays a reply form that includes the sender and all recipients of the selected message or the organizer and all attendees of the selected appointment.</span></span>

> [!NOTE]
> <span data-ttu-id="bac6a-719">Diese Methode wird in Outlook für iOS oder Outlook für Android nicht unterstützt.</span><span class="sxs-lookup"><span data-stu-id="bac6a-719">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="bac6a-720">In Outlook Web App wird das Antwortformular als Popupformular in der Dreispaltenansicht und als Popupformular in der Zwei- oder Einspaltenansicht angezeigt.</span><span class="sxs-lookup"><span data-stu-id="bac6a-720">In Outlook Web App, the reply form is displayed as a pop-out form in the 3-column view and a pop-up form in the 2- or 1-column view.</span></span>

<span data-ttu-id="bac6a-721">Wenn einer der Zeichenfolgenparameter seinen Grenzwert überschreitet, löst `displayReplyAllForm` eine Ausnahme aus.</span><span class="sxs-lookup"><span data-stu-id="bac6a-721">If any of the string parameters exceed their limits, `displayReplyAllForm` throws an exception.</span></span>

<span data-ttu-id="bac6a-p138">Wenn Anlagen im `formData.attachments`-Parameter angegeben werden, versuchen Outlook und Outlook Web App alle Anlagen herunterzuladen und sie an das Antwortformular anzufügen. Wenn Anlagen nicht hinzugefügt werden können, wird in der Formularbenutzeroberfläche ein Fehler ausgegeben. Wenn dies nicht möglich ist, wird keine Fehlermeldung ausgegeben.</span><span class="sxs-lookup"><span data-stu-id="bac6a-p138">When attachments are specified in the `formData.attachments` parameter, Outlook and Outlook Web App attempt to download all attachments and attach them to the reply form. If any attachments fail to be added, an error is shown in the form UI. If this isn't possible, then no error message is thrown.</span></span>

##### <a name="parameters"></a><span data-ttu-id="bac6a-725">Parameter:</span><span class="sxs-lookup"><span data-stu-id="bac6a-725">Parameters:</span></span>

| <span data-ttu-id="bac6a-726">Name</span><span class="sxs-lookup"><span data-stu-id="bac6a-726">Name</span></span> | <span data-ttu-id="bac6a-727">Typ</span><span class="sxs-lookup"><span data-stu-id="bac6a-727">Type</span></span> | <span data-ttu-id="bac6a-728">Attribute</span><span class="sxs-lookup"><span data-stu-id="bac6a-728">Attributes</span></span> | <span data-ttu-id="bac6a-729">Beschreibung</span><span class="sxs-lookup"><span data-stu-id="bac6a-729">Description</span></span> |
|---|---|---|---|
|`formData`| <span data-ttu-id="bac6a-730">Zeichenfolge &#124; Objekt</span><span class="sxs-lookup"><span data-stu-id="bac6a-730">String &#124; Object</span></span>| |<span data-ttu-id="bac6a-p139">Eine Zeichenfolge, die Text- und HTML-Code enthält, die den Hauptteil des Antwortformulars darstellen. Die Zeichenfolge ist auf 32 KB beschränkt.</span><span class="sxs-lookup"><span data-stu-id="bac6a-p139">A string that contains text and HTML and that represents the body of the reply form. The string is limited to 32 KB.</span></span><br/><span data-ttu-id="bac6a-733">**ODER**</span><span class="sxs-lookup"><span data-stu-id="bac6a-733">**OR**</span></span><br/><span data-ttu-id="bac6a-p140">Ein Objekt, das Text- oder Anlagendaten und eine Rückruffunktion enthält. Das Objekt ist wie folgt definiert:</span><span class="sxs-lookup"><span data-stu-id="bac6a-p140">An object that contains body or attachment data and a callback function. The object is defined as follows.</span></span> |
| `formData.htmlBody` | <span data-ttu-id="bac6a-736">Zeichenfolge</span><span class="sxs-lookup"><span data-stu-id="bac6a-736">String</span></span> | <span data-ttu-id="bac6a-737">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="bac6a-737">&lt;optional&gt;</span></span> | <span data-ttu-id="bac6a-p141">Eine Zeichenfolge, die Text- und HTML-Code enthält, die den Hauptteil des Antwortformulars darstellen. Die Zeichenfolge ist auf 32 KB beschränkt.</span><span class="sxs-lookup"><span data-stu-id="bac6a-p141">A string that contains text and HTML and that represents the body of the reply form. The string is limited to 32 KB.</span></span>
| `formData.attachments` | <span data-ttu-id="bac6a-740">Array.&lt;Object&gt;</span><span class="sxs-lookup"><span data-stu-id="bac6a-740">Array.&lt;Object&gt;</span></span> | <span data-ttu-id="bac6a-741">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="bac6a-741">&lt;optional&gt;</span></span> | <span data-ttu-id="bac6a-742">Ein Array mit JSON-Objekten, die Datei- oder Elementanlagen sind.</span><span class="sxs-lookup"><span data-stu-id="bac6a-742">An array of JSON objects that are either file or item attachments.</span></span> |
| `formData.attachments.type` | <span data-ttu-id="bac6a-743">Zeichenfolge</span><span class="sxs-lookup"><span data-stu-id="bac6a-743">String</span></span> | | <span data-ttu-id="bac6a-p142">Gibt den Typ der Anlage an. Muss `file` für eine Dateianlage oder `item` für eine Elementanlage sein.</span><span class="sxs-lookup"><span data-stu-id="bac6a-p142">Indicates the type of attachment. Must be `file` for a file attachment or `item` for an item attachment.</span></span> |
| `formData.attachments.name` | <span data-ttu-id="bac6a-746">Zeichenfolge</span><span class="sxs-lookup"><span data-stu-id="bac6a-746">String</span></span> | | <span data-ttu-id="bac6a-747">Eine Zeichenfolge, die den Namen der Anlage mit bis zu 255 Zeichen enthält.</span><span class="sxs-lookup"><span data-stu-id="bac6a-747">A string that contains the name of the attachment, up to 255 characters in length.</span></span>|
| `formData.attachments.url` | <span data-ttu-id="bac6a-748">Zeichenfolge</span><span class="sxs-lookup"><span data-stu-id="bac6a-748">String</span></span> | | <span data-ttu-id="bac6a-p143">Wird nur verwendet, wenn `type` auf `file` gesetzt ist. Der URI des Speicherorts für die Datei.</span><span class="sxs-lookup"><span data-stu-id="bac6a-p143">Only used if `type` is set to `file`. The URI of the location for the file.</span></span> |
| `formData.attachments.isInline` | <span data-ttu-id="bac6a-751">Boolean</span><span class="sxs-lookup"><span data-stu-id="bac6a-751">Boolean</span></span> | | <span data-ttu-id="bac6a-p144">Wird nur verwendet, wenn `type` auf `file` gesetzt ist. Wenn `true`: Zeigt an, dass die Anlage inline im Nachrichtentext angezeigt wird und nicht in der Anlagenlisten angezeigt werden soll.</span><span class="sxs-lookup"><span data-stu-id="bac6a-p144">Only used if `type` is set to `file`. If `true`, indicates that the attachment will be shown inline in the message body, and should not be displayed in the attachment list.</span></span> |
| `formData.attachments.itemId` | <span data-ttu-id="bac6a-754">Zeichenfolge</span><span class="sxs-lookup"><span data-stu-id="bac6a-754">String</span></span> | | <span data-ttu-id="bac6a-p145">Wird nur verwendet, wenn `type` auf `item` gesetzt ist. Die EWS-Element-ID der Anlage. Dies ist eine Zeichenfolge bis zu 100 Zeichen.</span><span class="sxs-lookup"><span data-stu-id="bac6a-p145">Only used if `type` is set to `item`. The EWS item id of the attachment. This is a string up to 100 characters.</span></span> |
| `callback` | <span data-ttu-id="bac6a-758">Funktion</span><span class="sxs-lookup"><span data-stu-id="bac6a-758">function</span></span> | <span data-ttu-id="bac6a-759">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="bac6a-759">&lt;optional&gt;</span></span> | <span data-ttu-id="bac6a-760">Bei Abschluss der Methode wird die im Parameter `callback` übergebene Funktion mit einem einzigen Parameter aufgerufen: `asyncResult`. Dieser Parameter ist ein Objekt des Typs [AsyncResult](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="bac6a-760">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [AsyncResult](/javascript/api/office/office.asyncresult) object.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="bac6a-761">Anforderungen</span><span class="sxs-lookup"><span data-stu-id="bac6a-761">Requirements</span></span>

|<span data-ttu-id="bac6a-762">Anforderung</span><span class="sxs-lookup"><span data-stu-id="bac6a-762">Requirement</span></span>| <span data-ttu-id="bac6a-763">Wert</span><span class="sxs-lookup"><span data-stu-id="bac6a-763">Value</span></span>|
|---|---|
|[<span data-ttu-id="bac6a-764">Mindestversion des Postfachanforderungssatzes</span><span class="sxs-lookup"><span data-stu-id="bac6a-764">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="bac6a-765">1.0</span><span class="sxs-lookup"><span data-stu-id="bac6a-765">1.0</span></span>|
|[<span data-ttu-id="bac6a-766">Mindestberechtigungsstufe</span><span class="sxs-lookup"><span data-stu-id="bac6a-766">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="bac6a-767">ReadItem</span><span class="sxs-lookup"><span data-stu-id="bac6a-767">ReadItem</span></span>|
|[<span data-ttu-id="bac6a-768">Zutreffender Outlook-Modus</span><span class="sxs-lookup"><span data-stu-id="bac6a-768">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="bac6a-769">Lesen</span><span class="sxs-lookup"><span data-stu-id="bac6a-769">Read</span></span>|

##### <a name="examples"></a><span data-ttu-id="bac6a-770">Beispiele</span><span class="sxs-lookup"><span data-stu-id="bac6a-770">Examples</span></span>

<span data-ttu-id="bac6a-771">Im folgenden Code wird eine Zeichenfolge an die `displayReplyAllForm`-Funktion übergeben.</span><span class="sxs-lookup"><span data-stu-id="bac6a-771">The following code passes a string to the `displayReplyAllForm` function.</span></span>

```
Office.context.mailbox.item.displayReplyAllForm('hello there');
Office.context.mailbox.item.displayReplyAllForm('<b>hello there</b>');
```

<span data-ttu-id="bac6a-772">Antworten mit einem leeren Textkörper.</span><span class="sxs-lookup"><span data-stu-id="bac6a-772">Reply with an empty body.</span></span>

```
Office.context.mailbox.item.displayReplyAllForm({});
```

<span data-ttu-id="bac6a-773">Antworten nur einem Textkörper.</span><span class="sxs-lookup"><span data-stu-id="bac6a-773">Reply with just a body.</span></span>

```
Office.context.mailbox.item.displayReplyAllForm(
{
  'htmlBody' : 'hi'
});
```

<span data-ttu-id="bac6a-774">Antworten mit einem Textkörper und einer Dateianlage.</span><span class="sxs-lookup"><span data-stu-id="bac6a-774">Reply with a body and a file attachment.</span></span>

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

<span data-ttu-id="bac6a-775">Antworten mit einem Textkörper und einer Elementanlage.</span><span class="sxs-lookup"><span data-stu-id="bac6a-775">Reply with a body and an item attachment.</span></span>

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

<span data-ttu-id="bac6a-776">Antworten Sie mit einem Textkörper, einer Dateianlage, einer Elementanlage und einem Callback.</span><span class="sxs-lookup"><span data-stu-id="bac6a-776">Reply with a body, file attachment, item attachment, and a callback.</span></span>

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

#### <a name="displayreplyformformdata"></a><span data-ttu-id="bac6a-777">displayReplyForm(formData)</span><span class="sxs-lookup"><span data-stu-id="bac6a-777">displayReplyForm(formData)</span></span>

<span data-ttu-id="bac6a-778">Zeigt ein Antwortformular an, das nur den Absender der ausgewählten Nachricht oder den Organisator des ausgewählten Termins enthält.</span><span class="sxs-lookup"><span data-stu-id="bac6a-778">Displays a reply form that includes only the sender of the selected message or the organizer of the selected appointment.</span></span>

> [!NOTE]
> <span data-ttu-id="bac6a-779">Diese Methode wird in Outlook für iOS oder Outlook für Android nicht unterstützt.</span><span class="sxs-lookup"><span data-stu-id="bac6a-779">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="bac6a-780">In Outlook Web App wird das Antwortformular als Popupformular in der Dreispaltenansicht und als Popupformular in der Zwei- oder Einspaltenansicht angezeigt.</span><span class="sxs-lookup"><span data-stu-id="bac6a-780">In Outlook Web App, the reply form is displayed as a pop-out form in the 3-column view and a pop-up form in the 2- or 1-column view.</span></span>

<span data-ttu-id="bac6a-781">Wenn einer der Zeichenfolgenparameter seinen Grenzwert überschreitet, löst `displayReplyForm` eine Ausnahme aus.</span><span class="sxs-lookup"><span data-stu-id="bac6a-781">If any of the string parameters exceed their limits, `displayReplyForm` throws an exception.</span></span>

<span data-ttu-id="bac6a-p146">Wenn Anlagen im `formData.attachments`-Parameter angegeben werden, versuchen Outlook und Outlook Web App alle Anlagen herunterzuladen und sie an das Antwortformular anzufügen. Wenn Anlagen nicht hinzugefügt werden können, wird in der Formularbenutzeroberfläche ein Fehler ausgegeben. Wenn dies nicht möglich ist, wird keine Fehlermeldung ausgegeben.</span><span class="sxs-lookup"><span data-stu-id="bac6a-p146">When attachments are specified in the `formData.attachments` parameter, Outlook and Outlook Web App attempt to download all attachments and attach them to the reply form. If any attachments fail to be added, an error is shown in the form UI. If this isn't possible, then no error message is thrown.</span></span>

##### <a name="parameters"></a><span data-ttu-id="bac6a-785">Parameter:</span><span class="sxs-lookup"><span data-stu-id="bac6a-785">Parameters:</span></span>

| <span data-ttu-id="bac6a-786">Name</span><span class="sxs-lookup"><span data-stu-id="bac6a-786">Name</span></span> | <span data-ttu-id="bac6a-787">Typ</span><span class="sxs-lookup"><span data-stu-id="bac6a-787">Type</span></span> | <span data-ttu-id="bac6a-788">Attribute</span><span class="sxs-lookup"><span data-stu-id="bac6a-788">Attributes</span></span> | <span data-ttu-id="bac6a-789">Beschreibung</span><span class="sxs-lookup"><span data-stu-id="bac6a-789">Description</span></span> |
|---|---|---|---|
|`formData`| <span data-ttu-id="bac6a-790">Zeichenfolge &#124; Objekt</span><span class="sxs-lookup"><span data-stu-id="bac6a-790">String &#124; Object</span></span>| | <span data-ttu-id="bac6a-p147">Eine Zeichenfolge, die Text- und HTML-Code enthält, die den Hauptteil des Antwortformulars darstellen. Die Zeichenfolge ist auf 32 KB beschränkt.</span><span class="sxs-lookup"><span data-stu-id="bac6a-p147">A string that contains text and HTML and that represents the body of the reply form. The string is limited to 32 KB.</span></span><br/><span data-ttu-id="bac6a-793">**ODER**</span><span class="sxs-lookup"><span data-stu-id="bac6a-793">**OR**</span></span><br/><span data-ttu-id="bac6a-p148">Ein Objekt, das Text- oder Anlagendaten und eine Rückruffunktion enthält. Das Objekt ist wie folgt definiert:</span><span class="sxs-lookup"><span data-stu-id="bac6a-p148">An object that contains body or attachment data and a callback function. The object is defined as follows.</span></span> |
| `formData.htmlBody` | <span data-ttu-id="bac6a-796">Zeichenfolge</span><span class="sxs-lookup"><span data-stu-id="bac6a-796">String</span></span> | <span data-ttu-id="bac6a-797">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="bac6a-797">&lt;optional&gt;</span></span> | <span data-ttu-id="bac6a-p149">Eine Zeichenfolge, die Text- und HTML-Code enthält, die den Hauptteil des Antwortformulars darstellen. Die Zeichenfolge ist auf 32 KB beschränkt.</span><span class="sxs-lookup"><span data-stu-id="bac6a-p149">A string that contains text and HTML and that represents the body of the reply form. The string is limited to 32 KB.</span></span>
| `formData.attachments` | <span data-ttu-id="bac6a-800">Array.&lt;Object&gt;</span><span class="sxs-lookup"><span data-stu-id="bac6a-800">Array.&lt;Object&gt;</span></span> | <span data-ttu-id="bac6a-801">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="bac6a-801">&lt;optional&gt;</span></span> | <span data-ttu-id="bac6a-802">Ein Array mit JSON-Objekten, die Datei- oder Elementanlagen sind.</span><span class="sxs-lookup"><span data-stu-id="bac6a-802">An array of JSON objects that are either file or item attachments.</span></span> |
| `formData.attachments.type` | <span data-ttu-id="bac6a-803">Zeichenfolge</span><span class="sxs-lookup"><span data-stu-id="bac6a-803">String</span></span> | | <span data-ttu-id="bac6a-p150">Gibt den Typ der Anlage an. Muss `file` für eine Dateianlage oder `item` für eine Elementanlage sein.</span><span class="sxs-lookup"><span data-stu-id="bac6a-p150">Indicates the type of attachment. Must be `file` for a file attachment or `item` for an item attachment.</span></span> |
| `formData.attachments.name` | <span data-ttu-id="bac6a-806">Zeichenfolge</span><span class="sxs-lookup"><span data-stu-id="bac6a-806">String</span></span> | | <span data-ttu-id="bac6a-807">Eine Zeichenfolge, die den Namen der Anlage mit bis zu 255 Zeichen enthält.</span><span class="sxs-lookup"><span data-stu-id="bac6a-807">A string that contains the name of the attachment, up to 255 characters in length.</span></span>|
| `formData.attachments.url` | <span data-ttu-id="bac6a-808">Zeichenfolge</span><span class="sxs-lookup"><span data-stu-id="bac6a-808">String</span></span> | | <span data-ttu-id="bac6a-p151">Wird nur verwendet, wenn `type` auf `file` gesetzt ist. Der URI des Speicherorts für die Datei.</span><span class="sxs-lookup"><span data-stu-id="bac6a-p151">Only used if `type` is set to `file`. The URI of the location for the file.</span></span> |
| `formData.attachments.isInline` | <span data-ttu-id="bac6a-811">Boolean</span><span class="sxs-lookup"><span data-stu-id="bac6a-811">Boolean</span></span> | | <span data-ttu-id="bac6a-p152">Wird nur verwendet, wenn `type` auf `file` gesetzt ist. Wenn `true`: Zeigt an, dass die Anlage inline im Nachrichtentext angezeigt wird und nicht in der Anlagenlisten angezeigt werden soll.</span><span class="sxs-lookup"><span data-stu-id="bac6a-p152">Only used if `type` is set to `file`. If `true`, indicates that the attachment will be shown inline in the message body, and should not be displayed in the attachment list.</span></span> |
| `formData.attachments.itemId` | <span data-ttu-id="bac6a-814">Zeichenfolge</span><span class="sxs-lookup"><span data-stu-id="bac6a-814">String</span></span> | | <span data-ttu-id="bac6a-p153">Wird nur verwendet, wenn `type` auf `item` gesetzt ist. Die EWS-Element-ID der Anlage. Dies ist eine Zeichenfolge bis zu 100 Zeichen.</span><span class="sxs-lookup"><span data-stu-id="bac6a-p153">Only used if `type` is set to `item`. The EWS item id of the attachment. This is a string up to 100 characters.</span></span> |
| `callback` | <span data-ttu-id="bac6a-818">Funktion</span><span class="sxs-lookup"><span data-stu-id="bac6a-818">function</span></span> | <span data-ttu-id="bac6a-819">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="bac6a-819">&lt;optional&gt;</span></span> | <span data-ttu-id="bac6a-820">Bei Abschluss der Methode wird die im Parameter `callback` übergebene Funktion mit einem einzigen Parameter aufgerufen: `asyncResult`. Dieser Parameter ist ein Objekt des Typs [AsyncResult](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="bac6a-820">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [AsyncResult](/javascript/api/office/office.asyncresult) object.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="bac6a-821">Anforderungen</span><span class="sxs-lookup"><span data-stu-id="bac6a-821">Requirements</span></span>

|<span data-ttu-id="bac6a-822">Anforderung</span><span class="sxs-lookup"><span data-stu-id="bac6a-822">Requirement</span></span>| <span data-ttu-id="bac6a-823">Wert</span><span class="sxs-lookup"><span data-stu-id="bac6a-823">Value</span></span>|
|---|---|
|[<span data-ttu-id="bac6a-824">Mindestversion des Postfachanforderungssatzes</span><span class="sxs-lookup"><span data-stu-id="bac6a-824">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="bac6a-825">1.0</span><span class="sxs-lookup"><span data-stu-id="bac6a-825">1.0</span></span>|
|[<span data-ttu-id="bac6a-826">Mindestberechtigungsstufe</span><span class="sxs-lookup"><span data-stu-id="bac6a-826">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="bac6a-827">ReadItem</span><span class="sxs-lookup"><span data-stu-id="bac6a-827">ReadItem</span></span>|
|[<span data-ttu-id="bac6a-828">Zutreffender Outlook-Modus</span><span class="sxs-lookup"><span data-stu-id="bac6a-828">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="bac6a-829">Lesen</span><span class="sxs-lookup"><span data-stu-id="bac6a-829">Read</span></span>|

##### <a name="examples"></a><span data-ttu-id="bac6a-830">Beispiele</span><span class="sxs-lookup"><span data-stu-id="bac6a-830">Examples</span></span>

<span data-ttu-id="bac6a-831">Im folgenden Code wird eine Zeichenfolge an die `displayReplyForm`-Funktion übergeben.</span><span class="sxs-lookup"><span data-stu-id="bac6a-831">The following code passes a string to the `displayReplyForm` function.</span></span>

```
Office.context.mailbox.item.displayReplyForm('hello there');
Office.context.mailbox.item.displayReplyForm('<b>hello there</b>');
```

<span data-ttu-id="bac6a-832">Antworten mit einem leeren Textkörper.</span><span class="sxs-lookup"><span data-stu-id="bac6a-832">Reply with an empty body.</span></span>

```
Office.context.mailbox.item.displayReplyForm({});
```

<span data-ttu-id="bac6a-833">Antworten nur einem Textkörper.</span><span class="sxs-lookup"><span data-stu-id="bac6a-833">Reply with just a body.</span></span>

```
Office.context.mailbox.item.displayReplyForm(
{
  'htmlBody' : 'hi'
});
```

<span data-ttu-id="bac6a-834">Antworten mit einem Textkörper und einer Dateianlage.</span><span class="sxs-lookup"><span data-stu-id="bac6a-834">Reply with a body and a file attachment.</span></span>

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

<span data-ttu-id="bac6a-835">Antworten mit einem Textkörper und einer Elementanlage.</span><span class="sxs-lookup"><span data-stu-id="bac6a-835">Reply with a body and an item attachment.</span></span>

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

<span data-ttu-id="bac6a-836">Antworten Sie mit einem Textkörper, einer Dateianlage, einer Elementanlage und einem Callback.</span><span class="sxs-lookup"><span data-stu-id="bac6a-836">Reply with a body, file attachment, item attachment, and a callback.</span></span>

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

#### <a name="getentities--entitiesjavascriptapioutlook16officeentities"></a><span data-ttu-id="bac6a-837">getEntities() → {[Entities](/javascript/api/outlook_1_6/office.entities)}</span><span class="sxs-lookup"><span data-stu-id="bac6a-837">getEntities() → {[Entities](/javascript/api/outlook_1_6/office.entities)}</span></span>

<span data-ttu-id="bac6a-838">Ruft das ausgewählte Element Textkörper gefundenen Entitäten ab.</span><span class="sxs-lookup"><span data-stu-id="bac6a-838">Gets the entities found in the selected item's body.</span></span>

> [!NOTE]
> <span data-ttu-id="bac6a-839">Diese Methode wird in Outlook für iOS oder Outlook für Android nicht unterstützt.</span><span class="sxs-lookup"><span data-stu-id="bac6a-839">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

##### <a name="requirements"></a><span data-ttu-id="bac6a-840">Anforderungen</span><span class="sxs-lookup"><span data-stu-id="bac6a-840">Requirements</span></span>

|<span data-ttu-id="bac6a-841">Anforderung</span><span class="sxs-lookup"><span data-stu-id="bac6a-841">Requirement</span></span>| <span data-ttu-id="bac6a-842">Wert</span><span class="sxs-lookup"><span data-stu-id="bac6a-842">Value</span></span>|
|---|---|
|[<span data-ttu-id="bac6a-843">Mindestversion des Postfachanforderungssatzes</span><span class="sxs-lookup"><span data-stu-id="bac6a-843">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="bac6a-844">1.0</span><span class="sxs-lookup"><span data-stu-id="bac6a-844">1.0</span></span>|
|[<span data-ttu-id="bac6a-845">Mindestberechtigungsstufe</span><span class="sxs-lookup"><span data-stu-id="bac6a-845">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="bac6a-846">ReadItem</span><span class="sxs-lookup"><span data-stu-id="bac6a-846">ReadItem</span></span>|
|[<span data-ttu-id="bac6a-847">Zutreffender Outlook-Modus</span><span class="sxs-lookup"><span data-stu-id="bac6a-847">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="bac6a-848">Lesen</span><span class="sxs-lookup"><span data-stu-id="bac6a-848">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="bac6a-849">Gibt zurück:</span><span class="sxs-lookup"><span data-stu-id="bac6a-849">Returns:</span></span>

<span data-ttu-id="bac6a-850">Typ: [Entitäten](/javascript/api/outlook_1_6/office.entities)</span><span class="sxs-lookup"><span data-stu-id="bac6a-850">Type: [Entities](/javascript/api/outlook_1_6/office.entities)</span></span>

##### <a name="example"></a><span data-ttu-id="bac6a-851">Beispiel</span><span class="sxs-lookup"><span data-stu-id="bac6a-851">Example</span></span>

<span data-ttu-id="bac6a-852">Das folgende Beispiel greift auf die Kontakte Entitäten im Textkörper des aktuellen Elements.</span><span class="sxs-lookup"><span data-stu-id="bac6a-852">The following example accesses the contacts entities in the current item's body.</span></span>

```
var contacts = Office.context.mailbox.item.getEntities().contacts;
```

#### <a name="getentitiesbytypeentitytype--nullable-arraystringcontactjavascriptapioutlook16officecontactmeetingsuggestionjavascriptapioutlook16officemeetingsuggestionphonenumberjavascriptapioutlook16officephonenumbertasksuggestionjavascriptapioutlook16officetasksuggestion"></a><span data-ttu-id="bac6a-853">getEntitiesByType(entityType) → (nullable) {Array.<(String|[Contact](/javascript/api/outlook_1_6/office.contact)|[MeetingSuggestion](/javascript/api/outlook_1_6/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook_1_6/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook_1_6/office.tasksuggestion))>}</span><span class="sxs-lookup"><span data-stu-id="bac6a-853">getEntitiesByType(entityType) → (nullable) {Array.<(String|[Contact](/javascript/api/outlook_1_6/office.contact)|[MeetingSuggestion](/javascript/api/outlook_1_6/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook_1_6/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook_1_6/office.tasksuggestion))>}</span></span>

<span data-ttu-id="bac6a-854">Ruft ein Array aller Entitäten des angegebenen Entitätstyps im Hauptteil des ausgewählten Elements gefunden.</span><span class="sxs-lookup"><span data-stu-id="bac6a-854">Gets an array of all the entities of the specified entity type found in the selected item's body.</span></span>

> [!NOTE]
> <span data-ttu-id="bac6a-855">Diese Methode wird in Outlook für iOS oder Outlook für Android nicht unterstützt.</span><span class="sxs-lookup"><span data-stu-id="bac6a-855">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

##### <a name="parameters"></a><span data-ttu-id="bac6a-856">Parameter:</span><span class="sxs-lookup"><span data-stu-id="bac6a-856">Parameters:</span></span>

|<span data-ttu-id="bac6a-857">Name</span><span class="sxs-lookup"><span data-stu-id="bac6a-857">Name</span></span>| <span data-ttu-id="bac6a-858">Typ</span><span class="sxs-lookup"><span data-stu-id="bac6a-858">Type</span></span>| <span data-ttu-id="bac6a-859">Beschreibung</span><span class="sxs-lookup"><span data-stu-id="bac6a-859">Description</span></span>|
|---|---|---|
|`entityType`| [<span data-ttu-id="bac6a-860">Office.MailboxEnums.EntityType</span><span class="sxs-lookup"><span data-stu-id="bac6a-860">Office.MailboxEnums.EntityType</span></span>](/javascript/api/outlook_1_6/office.mailboxenums.entitytype)|<span data-ttu-id="bac6a-861">Einer der Werte der EntityType-Enumeration.</span><span class="sxs-lookup"><span data-stu-id="bac6a-861">One of the EntityType enumeration values.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="bac6a-862">Anforderungen</span><span class="sxs-lookup"><span data-stu-id="bac6a-862">Requirements</span></span>

|<span data-ttu-id="bac6a-863">Anforderung</span><span class="sxs-lookup"><span data-stu-id="bac6a-863">Requirement</span></span>| <span data-ttu-id="bac6a-864">Wert</span><span class="sxs-lookup"><span data-stu-id="bac6a-864">Value</span></span>|
|---|---|
|[<span data-ttu-id="bac6a-865">Mindestversion des Postfachanforderungssatzes</span><span class="sxs-lookup"><span data-stu-id="bac6a-865">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="bac6a-866">1.0</span><span class="sxs-lookup"><span data-stu-id="bac6a-866">1.0</span></span>|
|[<span data-ttu-id="bac6a-867">Mindestberechtigungsstufe</span><span class="sxs-lookup"><span data-stu-id="bac6a-867">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="bac6a-868">Eingeschränkt</span><span class="sxs-lookup"><span data-stu-id="bac6a-868">Restricted</span></span>|
|[<span data-ttu-id="bac6a-869">Zutreffender Outlook-Modus</span><span class="sxs-lookup"><span data-stu-id="bac6a-869">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="bac6a-870">Lesen</span><span class="sxs-lookup"><span data-stu-id="bac6a-870">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="bac6a-871">Gibt zurück:</span><span class="sxs-lookup"><span data-stu-id="bac6a-871">Returns:</span></span>

<span data-ttu-id="bac6a-872">Wenn der in `entityType` übergebene Wert kein gültiges Element der `EntityType`-Enumeration ist, gibt die Methode null zurück.</span><span class="sxs-lookup"><span data-stu-id="bac6a-872">If the value passed in `entityType` is not a valid member of the `EntityType` enumeration, the method returns null.</span></span> <span data-ttu-id="bac6a-873">Wenn keine Entitäten des angegebenen Typs im Textkörper des Elements vorhanden sind, gibt die Methode ein leeres Array zurück.</span><span class="sxs-lookup"><span data-stu-id="bac6a-873">If no entities of the specified type are present in the item's body, the method returns an empty array.</span></span> <span data-ttu-id="bac6a-874">Andernfalls hängt der Typ der Objekte im zurückgegebenen Array vom Typ der Entität ab, die im `entityType`-Parameter angefordert wurde.</span><span class="sxs-lookup"><span data-stu-id="bac6a-874">Otherwise, the type of the objects in the returned array depends on the type of entity requested in the `entityType` parameter.</span></span>

<span data-ttu-id="bac6a-875">Während Sie die minimale Berechtigungsstufe für diese Methode **Restricted** ist, erfordern einige Entitätstypen **ReadItem** für den Zugriff, wie in der folgenden Tabelle angegeben wird.</span><span class="sxs-lookup"><span data-stu-id="bac6a-875">While the minimum permission level to use this method is **Restricted**, some entity types require **ReadItem** to access, as specified in the following table.</span></span>

| <span data-ttu-id="bac6a-876">Wert von `entityType`</span><span class="sxs-lookup"><span data-stu-id="bac6a-876">Value of `entityType`</span></span> | <span data-ttu-id="bac6a-877">Typ der Objekte im zurückgegebenen Array</span><span class="sxs-lookup"><span data-stu-id="bac6a-877">Type of objects in returned array</span></span> | <span data-ttu-id="bac6a-878">Erforderliche Berechtigungsstufe</span><span class="sxs-lookup"><span data-stu-id="bac6a-878">Required Permission Level</span></span> |
| --- | --- | --- |
| `Address` | <span data-ttu-id="bac6a-879">Zeichenfolge</span><span class="sxs-lookup"><span data-stu-id="bac6a-879">String</span></span> | <span data-ttu-id="bac6a-880">**Restricted**</span><span class="sxs-lookup"><span data-stu-id="bac6a-880">**Restricted**</span></span> |
| `Contact` | <span data-ttu-id="bac6a-881">Contact</span><span class="sxs-lookup"><span data-stu-id="bac6a-881">Contact</span></span> | <span data-ttu-id="bac6a-882">**ReadItem**</span><span class="sxs-lookup"><span data-stu-id="bac6a-882">**ReadItem**</span></span> |
| `EmailAddress` | <span data-ttu-id="bac6a-883">Zeichenfolge</span><span class="sxs-lookup"><span data-stu-id="bac6a-883">String</span></span> | <span data-ttu-id="bac6a-884">**ReadItem**</span><span class="sxs-lookup"><span data-stu-id="bac6a-884">**ReadItem**</span></span> |
| `MeetingSuggestion` | <span data-ttu-id="bac6a-885">MeetingSuggestion</span><span class="sxs-lookup"><span data-stu-id="bac6a-885">MeetingSuggestion</span></span> | <span data-ttu-id="bac6a-886">**ReadItem**</span><span class="sxs-lookup"><span data-stu-id="bac6a-886">**ReadItem**</span></span> |
| `PhoneNumber` | <span data-ttu-id="bac6a-887">PhoneNumber</span><span class="sxs-lookup"><span data-stu-id="bac6a-887">PhoneNumber</span></span> | <span data-ttu-id="bac6a-888">**Restricted**</span><span class="sxs-lookup"><span data-stu-id="bac6a-888">**Restricted**</span></span> |
| `TaskSuggestion` | <span data-ttu-id="bac6a-889">TaskSuggestion</span><span class="sxs-lookup"><span data-stu-id="bac6a-889">TaskSuggestion</span></span> | <span data-ttu-id="bac6a-890">**ReadItem**</span><span class="sxs-lookup"><span data-stu-id="bac6a-890">**ReadItem**</span></span> |
| `URL` | <span data-ttu-id="bac6a-891">Zeichenfolge</span><span class="sxs-lookup"><span data-stu-id="bac6a-891">String</span></span> | <span data-ttu-id="bac6a-892">**Restricted**</span><span class="sxs-lookup"><span data-stu-id="bac6a-892">**Restricted**</span></span> |

<span data-ttu-id="bac6a-893">Typ: Array.<(String|[Contact](/javascript/api/outlook_1_6/office.contact)|[MeetingSuggestion](/javascript/api/outlook_1_6/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook_1_6/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook_1_6/office.tasksuggestion))></span><span class="sxs-lookup"><span data-stu-id="bac6a-893">Type: Array.<(String|[Contact](/javascript/api/outlook_1_6/office.contact)|[MeetingSuggestion](/javascript/api/outlook_1_6/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook_1_6/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook_1_6/office.tasksuggestion))></span></span>

##### <a name="example"></a><span data-ttu-id="bac6a-894">Beispiel</span><span class="sxs-lookup"><span data-stu-id="bac6a-894">Example</span></span>

<span data-ttu-id="bac6a-895">Im folgenden Beispiel wird veranschaulicht, wie auf ein Array von Zeichenfolgen, die Postadressen im Textkörper des aktuellen Elements darstellen.</span><span class="sxs-lookup"><span data-stu-id="bac6a-895">The following example shows how to access an array of strings that represent postal addresses in the current item's body.</span></span>

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

#### <a name="getfilteredentitiesbynamename--nullable-arraystringcontactjavascriptapioutlook16officecontactmeetingsuggestionjavascriptapioutlook16officemeetingsuggestionphonenumberjavascriptapioutlook16officephonenumbertasksuggestionjavascriptapioutlook16officetasksuggestion"></a><span data-ttu-id="bac6a-896">getFilteredEntitiesByName(name) → (nullable) {Array.<(String|[Contact](/javascript/api/outlook_1_6/office.contact)|[MeetingSuggestion](/javascript/api/outlook_1_6/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook_1_6/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook_1_6/office.tasksuggestion))>}</span><span class="sxs-lookup"><span data-stu-id="bac6a-896">getFilteredEntitiesByName(name) → (nullable) {Array.<(String|[Contact](/javascript/api/outlook_1_6/office.contact)|[MeetingSuggestion](/javascript/api/outlook_1_6/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook_1_6/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook_1_6/office.tasksuggestion))>}</span></span>

<span data-ttu-id="bac6a-897">Gibt bekannte Entitäten im ausgewählten Element zurück, die den in der XML-Manifestdatei definierten benannten Filter übergeben.</span><span class="sxs-lookup"><span data-stu-id="bac6a-897">Returns well-known entities in the selected item that pass the named filter defined in the manifest XML file.</span></span>

> [!NOTE]
> <span data-ttu-id="bac6a-898">Diese Methode wird in Outlook für iOS oder Outlook für Android nicht unterstützt.</span><span class="sxs-lookup"><span data-stu-id="bac6a-898">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="bac6a-899">Die `getFilteredEntitiesByName`-Methode gibt die Entitäten zurück, die dem im [ItemHasKnownEntity](/javascript/office/manifest/rule#itemhasknownentity-rule)-Regelelement der XML-Manifestdatei definierten regulären Ausdruck mit dem angegebenen `FilterName`-Elementwert entsprechen.</span><span class="sxs-lookup"><span data-stu-id="bac6a-899">The `getFilteredEntitiesByName` method returns the entities that match the regular expression defined in the [ItemHasKnownEntity](/javascript/office/manifest/rule#itemhasknownentity-rule) rule element in the manifest XML file with the specified `FilterName` element value.</span></span>

##### <a name="parameters"></a><span data-ttu-id="bac6a-900">Parameter:</span><span class="sxs-lookup"><span data-stu-id="bac6a-900">Parameters:</span></span>

|<span data-ttu-id="bac6a-901">Name</span><span class="sxs-lookup"><span data-stu-id="bac6a-901">Name</span></span>| <span data-ttu-id="bac6a-902">Typ</span><span class="sxs-lookup"><span data-stu-id="bac6a-902">Type</span></span>| <span data-ttu-id="bac6a-903">Beschreibung</span><span class="sxs-lookup"><span data-stu-id="bac6a-903">Description</span></span>|
|---|---|---|
|`name`| <span data-ttu-id="bac6a-904">Zeichenfolge</span><span class="sxs-lookup"><span data-stu-id="bac6a-904">String</span></span>|<span data-ttu-id="bac6a-905">Der Name des `ItemHasKnownEntity`-Regelelements, das den entsprechenden Filter definiert.</span><span class="sxs-lookup"><span data-stu-id="bac6a-905">The name of the `ItemHasKnownEntity` rule element that defines the filter to match.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="bac6a-906">Anforderungen</span><span class="sxs-lookup"><span data-stu-id="bac6a-906">Requirements</span></span>

|<span data-ttu-id="bac6a-907">Anforderung</span><span class="sxs-lookup"><span data-stu-id="bac6a-907">Requirement</span></span>| <span data-ttu-id="bac6a-908">Wert</span><span class="sxs-lookup"><span data-stu-id="bac6a-908">Value</span></span>|
|---|---|
|[<span data-ttu-id="bac6a-909">Mindestversion des Postfachanforderungssatzes</span><span class="sxs-lookup"><span data-stu-id="bac6a-909">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="bac6a-910">1.0</span><span class="sxs-lookup"><span data-stu-id="bac6a-910">1.0</span></span>|
|[<span data-ttu-id="bac6a-911">Mindestberechtigungsstufe</span><span class="sxs-lookup"><span data-stu-id="bac6a-911">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="bac6a-912">ReadItem</span><span class="sxs-lookup"><span data-stu-id="bac6a-912">ReadItem</span></span>|
|[<span data-ttu-id="bac6a-913">Zutreffender Outlook-Modus</span><span class="sxs-lookup"><span data-stu-id="bac6a-913">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="bac6a-914">Lesen</span><span class="sxs-lookup"><span data-stu-id="bac6a-914">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="bac6a-915">Gibt zurück:</span><span class="sxs-lookup"><span data-stu-id="bac6a-915">Returns:</span></span>

<span data-ttu-id="bac6a-p155">Befindet sich kein `ItemHasKnownEntity`-Element im Manifest mit einem `FilterName`-Elementwert, der dem `name`-Parameter entspricht, gibt die Methode `null` zurück. Wenn der `name`-Parameter einem `ItemHasKnownEntity`-Element im Manifest entspricht, es aber keine entsprechenden Entitäten im aktuellen Element gibt, gibt die Methode ein leeres Array zurück.</span><span class="sxs-lookup"><span data-stu-id="bac6a-p155">If there is no `ItemHasKnownEntity` element in the manifest with a `FilterName` element value that matches the `name` parameter, the method returns `null`. If the `name` parameter does match an `ItemHasKnownEntity` element in the manifest, but there are no entities in the current item that match, the method return an empty array.</span></span>

<span data-ttu-id="bac6a-918">Typ: Array.<(String|[Contact](/javascript/api/outlook_1_6/office.contact)|[MeetingSuggestion](/javascript/api/outlook_1_6/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook_1_6/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook_1_6/office.tasksuggestion))></span><span class="sxs-lookup"><span data-stu-id="bac6a-918">Type: Array.<(String|[Contact](/javascript/api/outlook_1_6/office.contact)|[MeetingSuggestion](/javascript/api/outlook_1_6/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook_1_6/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook_1_6/office.tasksuggestion))></span></span>

#### <a name="getregexmatches--object"></a><span data-ttu-id="bac6a-919">getRegExMatches() → {Object}</span><span class="sxs-lookup"><span data-stu-id="bac6a-919">getRegExMatches() → {Object}</span></span>

<span data-ttu-id="bac6a-920">Gibt Zeichenfolgenwerte im ausgewählten Element zurück, die den in der XML-Manifestdatei definierten regulären Ausdrücken entsprechen.</span><span class="sxs-lookup"><span data-stu-id="bac6a-920">Returns string values in the selected item that match the regular expressions defined in the manifest XML file.</span></span>

> [!NOTE]
> <span data-ttu-id="bac6a-921">Diese Methode wird in Outlook für iOS oder Outlook für Android nicht unterstützt.</span><span class="sxs-lookup"><span data-stu-id="bac6a-921">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="bac6a-p156">Die `getRegExMatches`-Methode gibt die Zeichenfolgen zurück, die dem im `ItemHasRegularExpressionMatch`- oder `ItemHasKnownEntity`-Regelelement der XML-Manifestdatei definierten regulären Ausdruck entsprechen. Bei einer `ItemHasRegularExpressionMatch`-Regel muss eine entsprechende Zeichenfolge in der Eigenschaft des Elements vorhanden sein, das von dieser Regel angegeben wird. Der einfache `PropertyName`-Typ definiert die unterstützten Eigenschaften.</span><span class="sxs-lookup"><span data-stu-id="bac6a-p156">The `getRegExMatches` method returns the strings that match the regular expression defined in each `ItemHasRegularExpressionMatch` or `ItemHasKnownEntity` rule element in the manifest XML file. For an `ItemHasRegularExpressionMatch` rule, a matching string has to occur in the property of the item that is specified by that rule. The `PropertyName` simple type defines the supported properties.</span></span>

<span data-ttu-id="bac6a-925">Nehmen Sie z. B. an, dass ein Add-In-Manifest das folgende `Rule`-Element aufweist:</span><span class="sxs-lookup"><span data-stu-id="bac6a-925">For example, consider an add-in manifest has the following `Rule` element:</span></span>

```
<Rule xsi:type="RuleCollection" Mode="And">
  <Rule xsi:type="ItemIs" FormType="Read" ItemType="Message" />
  <Rule xsi:type="RuleCollection" Mode="Or">
    <Rule xsi:type="ItemHasRegularExpressionMatch" RegExName="fruits" RegExValue="apple|banana|coconut" PropertyName="BodyAsPlaintext" IgnoreCase="true" />
    <Rule xsi:type="ItemHasRegularExpressionMatch" RegExName="veggies" RegExValue="tomato|onion|spinach|broccoli" PropertyName="BodyAsPlaintext" IgnoreCase="true" />
  </Rule>
</Rule>
```

<span data-ttu-id="bac6a-926">Das von `getRegExMatches` zurückgegebene Objekt hätte zwei Eigenschaften: `fruits` und `veggies`.</span><span class="sxs-lookup"><span data-stu-id="bac6a-926">The object returned from `getRegExMatches` would have two properties: `fruits` and `veggies`.</span></span>

```
{
  'fruits': ['apple','banana','Banana','coconut'],
  'veggies': ['tomato','onion','spinach','broccoli']
}
```

<span data-ttu-id="bac6a-p157">Wenn Sie eine `ItemHasRegularExpressionMatch`-Regel für die Textkörpereigenschaft eines Elements festlegen, sollte der reguläre Ausdruck den Textkörper weiter filtern und nicht versuchen, den gesamten Textkörper des Elements zurückzugeben. Wenn der gesamte Textkörper eines Elements mit einem regulären Ausdruck wie `.*` abgerufen wird, werden nicht immer die gewünschten Ergebnisse erzielt. Verwenden Sie stattdessen die [`Body.getAsync`](/javascript/api/outlook_1_6/office.body#getasync-coerciontype--options--callback-)-Methode, um den gesamten Textkörper abzurufen.</span><span class="sxs-lookup"><span data-stu-id="bac6a-p157">If you specify an `ItemHasRegularExpressionMatch` rule on the body property of an item, the regular expression should further filter the body and should not attempt to return the entire body of the item. Using a regular expression such as `.*` to obtain the entire body of an item does not always return the expected results. Instead, use the [`Body.getAsync`](/javascript/api/outlook_1_6/office.body#getasync-coerciontype--options--callback-) method to retrieve the entire body.</span></span>

##### <a name="requirements"></a><span data-ttu-id="bac6a-930">Anforderungen</span><span class="sxs-lookup"><span data-stu-id="bac6a-930">Requirements</span></span>

|<span data-ttu-id="bac6a-931">Anforderung</span><span class="sxs-lookup"><span data-stu-id="bac6a-931">Requirement</span></span>| <span data-ttu-id="bac6a-932">Wert</span><span class="sxs-lookup"><span data-stu-id="bac6a-932">Value</span></span>|
|---|---|
|[<span data-ttu-id="bac6a-933">Mindestversion des Postfachanforderungssatzes</span><span class="sxs-lookup"><span data-stu-id="bac6a-933">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="bac6a-934">1.0</span><span class="sxs-lookup"><span data-stu-id="bac6a-934">1.0</span></span>|
|[<span data-ttu-id="bac6a-935">Mindestberechtigungsstufe</span><span class="sxs-lookup"><span data-stu-id="bac6a-935">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="bac6a-936">ReadItem</span><span class="sxs-lookup"><span data-stu-id="bac6a-936">ReadItem</span></span>|
|[<span data-ttu-id="bac6a-937">Zutreffender Outlook-Modus</span><span class="sxs-lookup"><span data-stu-id="bac6a-937">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="bac6a-938">Lesen</span><span class="sxs-lookup"><span data-stu-id="bac6a-938">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="bac6a-939">Gibt zurück:</span><span class="sxs-lookup"><span data-stu-id="bac6a-939">Returns:</span></span>

<span data-ttu-id="bac6a-p158">Ein Objekt mit Arrays aus Zeichenfolgen, die den in der XML-Manifestdatei definierten regulären Ausdrücken entsprechen. Der Name der einzelnen Arrays ist gleich dem entsprechenden Wert des `RegExName`-Attributs der entsprechenden `ItemHasRegularExpressionMatch`-Regel oder des `FilterName`-Attributs der entsprechenden `ItemHasKnownEntity`-Regel.</span><span class="sxs-lookup"><span data-stu-id="bac6a-p158">An object that contains arrays of strings that match the regular expressions defined in the manifest XML file. The name of each array is equal to the corresponding value of the `RegExName` attribute of the matching `ItemHasRegularExpressionMatch` rule or the `FilterName` attribute of the matching `ItemHasKnownEntity` rule.</span></span>

<dl class="param-type"><span data-ttu-id="bac6a-942">

<dt>
Typ</dt>


</span><span class="sxs-lookup"><span data-stu-id="bac6a-942">

<dt>Type</dt>

</span></span><dd><span data-ttu-id="bac6a-943">Object</span><span class="sxs-lookup"><span data-stu-id="bac6a-943">Object</span></span></dd>

</dl>

##### <a name="example"></a><span data-ttu-id="bac6a-944">Beispiel</span><span class="sxs-lookup"><span data-stu-id="bac6a-944">Example</span></span>

<span data-ttu-id="bac6a-945">Das folgende Beispiel zeigt, wie Sie auf das Array von Übereinstimmungen für die regulären Ausdrucksregelelemente `fruits` und `veggies` zugreifen, die im Manifest angegeben sind.</span><span class="sxs-lookup"><span data-stu-id="bac6a-945">The following example shows how to access the array of matches for the regular expression rule elements `fruits` and `veggies`, which are specified in the manifest.</span></span>

```
var allMatches = Office.context.mailbox.item.getRegExMatches();
var fruits = allMatches.fruits;
var veges = allMatches.veggies;
```

#### <a name="getregexmatchesbynamename--nullable-array-string-"></a><span data-ttu-id="bac6a-946">getRegExMatchesByName(name) → (Nullwerte zulassen) {Array. < Zeichenfolge >}</span><span class="sxs-lookup"><span data-stu-id="bac6a-946">getRegExMatchesByName(name) → (nullable) {Array.< String >}</span></span>

<span data-ttu-id="bac6a-947">Gibt Zeichenfolgenwerte im ausgewählten Element zurück, die dem in der XML-Manifestdatei definierten benannten regulären Ausdruck entsprechen.</span><span class="sxs-lookup"><span data-stu-id="bac6a-947">Returns string values in the selected item that match the named regular expression defined in the manifest XML file.</span></span>

> [!NOTE]
> <span data-ttu-id="bac6a-948">Diese Methode wird in Outlook für iOS oder Outlook für Android nicht unterstützt.</span><span class="sxs-lookup"><span data-stu-id="bac6a-948">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="bac6a-949">Die `getRegExMatchesByName`-Methode gibt die Zeichenfolgen zurück, die dem im `ItemHasRegularExpressionMatch`-Regelelement der XML-Manifestdatei definierten regulären Ausdruck mit dem angegebenen `RegExName`-Elementwert entsprechen.</span><span class="sxs-lookup"><span data-stu-id="bac6a-949">The `getRegExMatchesByName` method returns the strings that match the regular expression defined in the `ItemHasRegularExpressionMatch` rule element in the manifest XML file with the specified `RegExName` element value.</span></span>

<span data-ttu-id="bac6a-p159">Wenn Sie eine `ItemHasRegularExpressionMatch`-Regel für die Textkörpereigenschaft eines Elements festlegen, sollte der reguläre Ausdruck den Textkörper weiter filtern und nicht versuchen, den gesamten Textkörper des Elements zurückzugeben. Wenn der gesamte Textkörper eines Elements mit einem regulären Ausdruck wie `.*` abgerufen wird, werden nicht immer die gewünschten Ergebnisse erzielt.</span><span class="sxs-lookup"><span data-stu-id="bac6a-p159">If you specify an `ItemHasRegularExpressionMatch` rule on the body property of an item, the regular expression should further filter the body and should not attempt to return the entire body of the item. Using a regular expression such as `.*` to obtain the entire body of an item does not always return the expected results.</span></span>

##### <a name="parameters"></a><span data-ttu-id="bac6a-952">Parameter:</span><span class="sxs-lookup"><span data-stu-id="bac6a-952">Parameters:</span></span>

|<span data-ttu-id="bac6a-953">Name</span><span class="sxs-lookup"><span data-stu-id="bac6a-953">Name</span></span>| <span data-ttu-id="bac6a-954">Typ</span><span class="sxs-lookup"><span data-stu-id="bac6a-954">Type</span></span>| <span data-ttu-id="bac6a-955">Beschreibung</span><span class="sxs-lookup"><span data-stu-id="bac6a-955">Description</span></span>|
|---|---|---|
|`name`| <span data-ttu-id="bac6a-956">Zeichenfolge</span><span class="sxs-lookup"><span data-stu-id="bac6a-956">String</span></span>|<span data-ttu-id="bac6a-957">Der Name des `ItemHasRegularExpressionMatch`-Regelelements, das den entsprechenden Filter definiert.</span><span class="sxs-lookup"><span data-stu-id="bac6a-957">The name of the `ItemHasRegularExpressionMatch` rule element that defines the filter to match.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="bac6a-958">Anforderungen</span><span class="sxs-lookup"><span data-stu-id="bac6a-958">Requirements</span></span>

|<span data-ttu-id="bac6a-959">Anforderung</span><span class="sxs-lookup"><span data-stu-id="bac6a-959">Requirement</span></span>| <span data-ttu-id="bac6a-960">Wert</span><span class="sxs-lookup"><span data-stu-id="bac6a-960">Value</span></span>|
|---|---|
|[<span data-ttu-id="bac6a-961">Mindestversion des Postfachanforderungssatzes</span><span class="sxs-lookup"><span data-stu-id="bac6a-961">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="bac6a-962">1.0</span><span class="sxs-lookup"><span data-stu-id="bac6a-962">1.0</span></span>|
|[<span data-ttu-id="bac6a-963">Mindestberechtigungsstufe</span><span class="sxs-lookup"><span data-stu-id="bac6a-963">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="bac6a-964">ReadItem</span><span class="sxs-lookup"><span data-stu-id="bac6a-964">ReadItem</span></span>|
|[<span data-ttu-id="bac6a-965">Zutreffender Outlook-Modus</span><span class="sxs-lookup"><span data-stu-id="bac6a-965">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="bac6a-966">Lesen</span><span class="sxs-lookup"><span data-stu-id="bac6a-966">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="bac6a-967">Gibt zurück:</span><span class="sxs-lookup"><span data-stu-id="bac6a-967">Returns:</span></span>

<span data-ttu-id="bac6a-968">Ein Array mit den Zeichenfolgen, die dem in der XML-Manifestdatei definierten regulären Ausdruck entsprechen.</span><span class="sxs-lookup"><span data-stu-id="bac6a-968">An array that contains the strings that match the regular expression defined in the manifest XML file.</span></span>

<dl class="param-type"><span data-ttu-id="bac6a-969">

<dt>Typ</dt>

</span><span class="sxs-lookup"><span data-stu-id="bac6a-969">

<dt>Type</dt>

</span></span><dd><span data-ttu-id="bac6a-970">Array. < Zeichenfolge ></span><span class="sxs-lookup"><span data-stu-id="bac6a-970">Array.< String ></span></span></dd>

</dl>

##### <a name="example"></a><span data-ttu-id="bac6a-971">Beispiel</span><span class="sxs-lookup"><span data-stu-id="bac6a-971">Example</span></span>

```
var fruits = Office.context.mailbox.item.getRegExMatchesByName("fruits");
var veggies = Office.context.mailbox.item.getRegExMatchesByName("veggies");
```

####  <a name="getselecteddataasynccoerciontype-options-callback--string"></a><span data-ttu-id="bac6a-972">getSelectedDataAsync(coercionType, [options], callback) → {String}</span><span class="sxs-lookup"><span data-stu-id="bac6a-972">getSelectedDataAsync(coercionType, [options], callback) → {String}</span></span>

<span data-ttu-id="bac6a-973">Gibt asynchron ausgewählte Daten aus dem Betreff oder Textkörper einer Nachricht zurück.</span><span class="sxs-lookup"><span data-stu-id="bac6a-973">Asynchronously returns selected data from the subject or body of a message.</span></span>

<span data-ttu-id="bac6a-p160">Wenn keine Auswahl vorhanden ist, aber der Cursor sich im Nachrichtentext oder Betreff befindet, gibt die Methode für die ausgewählten Daten NULL zurück. Wenn ein anderes Feld als der Textkörper oder Betreff ausgewählt ist, gibt die Methode den `InvalidSelection`-Fehler zurück.</span><span class="sxs-lookup"><span data-stu-id="bac6a-p160">If there is no selection but the cursor is in the body or subject, the method returns null for the selected data. If a field other than the body or subject is selected, the method returns the `InvalidSelection` error.</span></span>

##### <a name="parameters"></a><span data-ttu-id="bac6a-976">Parameter:</span><span class="sxs-lookup"><span data-stu-id="bac6a-976">Parameters:</span></span>

|<span data-ttu-id="bac6a-977">Name</span><span class="sxs-lookup"><span data-stu-id="bac6a-977">Name</span></span>| <span data-ttu-id="bac6a-978">Typ</span><span class="sxs-lookup"><span data-stu-id="bac6a-978">Type</span></span>| <span data-ttu-id="bac6a-979">Attribute</span><span class="sxs-lookup"><span data-stu-id="bac6a-979">Attributes</span></span>| <span data-ttu-id="bac6a-980">Beschreibung</span><span class="sxs-lookup"><span data-stu-id="bac6a-980">Description</span></span>|
|---|---|---|---|
|`coercionType`| [<span data-ttu-id="bac6a-981">Office.CoercionType</span><span class="sxs-lookup"><span data-stu-id="bac6a-981">Office.CoercionType</span></span>](office.md#coerciontype-string)||<span data-ttu-id="bac6a-p161">Fordert ein Format für die Daten an. Wenn es sich um Texthandelt, gibt die Methode den unformatierten Text als Zeichenfolge zurück und entfernt eventuell vorhandene HTML-Tags. Wenn es sich um HTML handelt, gibt die Methode den ausgewählten Text zurück, entweder als unformatierten Text oder als HTML.</span><span class="sxs-lookup"><span data-stu-id="bac6a-p161">Requests a format for the data. If Text, the method returns the plain text as a string , removing any HTML tags present. If HTML, the method returns the selected text, whether it is plaintext or HTML.</span></span>|
|`options`| <span data-ttu-id="bac6a-985">Object</span><span class="sxs-lookup"><span data-stu-id="bac6a-985">Object</span></span>| <span data-ttu-id="bac6a-986">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="bac6a-986">&lt;optional&gt;</span></span>|<span data-ttu-id="bac6a-987">Ein Objektliteral, das mindestens eine der folgenden Eigenschaften enthält.</span><span class="sxs-lookup"><span data-stu-id="bac6a-987">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`| <span data-ttu-id="bac6a-988">Object</span><span class="sxs-lookup"><span data-stu-id="bac6a-988">Object</span></span>| <span data-ttu-id="bac6a-989">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="bac6a-989">&lt;optional&gt;</span></span>|<span data-ttu-id="bac6a-990">Entwickler können ein Objekt bereitstellen, auf das sie in der Callbackmethode zugreifen möchten.</span><span class="sxs-lookup"><span data-stu-id="bac6a-990">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`| <span data-ttu-id="bac6a-991">Funktion</span><span class="sxs-lookup"><span data-stu-id="bac6a-991">function</span></span>||<span data-ttu-id="bac6a-992">Wenn die Methode abgeschlossen wird, wird die Funktion , die im `callback`-Parameter übergeben wird, mit einem einzelnen Parameter (`asyncResult`) aufgerufen, der ein [`AsyncResult`](/javascript/api/office/office.asyncresult)-Objekt darstellt.</span><span class="sxs-lookup"><span data-stu-id="bac6a-992">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span><br/><br/><span data-ttu-id="bac6a-993">Rufen Sie für den Zugriff auf die ausgewählten Daten aus der Rückrufmethode `asyncResult.value.data` auf.</span><span class="sxs-lookup"><span data-stu-id="bac6a-993">To access the selected data from the callback method, call `asyncResult.value.data`.</span></span> <span data-ttu-id="bac6a-994">Rufen Sie die Source-Eigenschaft für den Zugriff, die die Auswahl stammen, `asyncResult.value.sourceProperty`, die entweder sein `body` oder `subject`.</span><span class="sxs-lookup"><span data-stu-id="bac6a-994">To access the source property that the selection comes from, call `asyncResult.value.sourceProperty`, which will be either `body` or `subject`.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="bac6a-995">Anforderungen</span><span class="sxs-lookup"><span data-stu-id="bac6a-995">Requirements</span></span>

|<span data-ttu-id="bac6a-996">Anforderung</span><span class="sxs-lookup"><span data-stu-id="bac6a-996">Requirement</span></span>| <span data-ttu-id="bac6a-997">Wert</span><span class="sxs-lookup"><span data-stu-id="bac6a-997">Value</span></span>|
|---|---|
|[<span data-ttu-id="bac6a-998">Mindestversion des Postfachanforderungssatzes</span><span class="sxs-lookup"><span data-stu-id="bac6a-998">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="bac6a-999">1.2</span><span class="sxs-lookup"><span data-stu-id="bac6a-999">1.2</span></span>|
|[<span data-ttu-id="bac6a-1000">Mindestberechtigungsstufe</span><span class="sxs-lookup"><span data-stu-id="bac6a-1000">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="bac6a-1001">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="bac6a-1001">ReadWriteItem</span></span>|
|[<span data-ttu-id="bac6a-1002">Zutreffender Outlook-Modus</span><span class="sxs-lookup"><span data-stu-id="bac6a-1002">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="bac6a-1003">Verfassen</span><span class="sxs-lookup"><span data-stu-id="bac6a-1003">Compose</span></span>|

##### <a name="returns"></a><span data-ttu-id="bac6a-1004">Gibt zurück:</span><span class="sxs-lookup"><span data-stu-id="bac6a-1004">Returns:</span></span>

<span data-ttu-id="bac6a-1005">Die ausgewählten Daten als Zeichenfolge mit dem durch `coercionType` bestimmten Format.</span><span class="sxs-lookup"><span data-stu-id="bac6a-1005">The selected data as a string with format determined by `coercionType`.</span></span>

<dl class="param-type"><span data-ttu-id="bac6a-1006">

<dt>
Typ</dt>


</span><span class="sxs-lookup"><span data-stu-id="bac6a-1006">

<dt>Type</dt>

</span></span><dd><span data-ttu-id="bac6a-1007">Zeichenfolge</span><span class="sxs-lookup"><span data-stu-id="bac6a-1007">String</span></span></dd>

</dl>

##### <a name="example"></a><span data-ttu-id="bac6a-1008">Beispiel</span><span class="sxs-lookup"><span data-stu-id="bac6a-1008">Example</span></span>

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

#### <a name="getselectedentities--entitiesjavascriptapioutlook16officeentities"></a><span data-ttu-id="bac6a-1009">getSelectedEntities() → {[Entitäten](/javascript/api/outlook_1_6/office.entities)}</span><span class="sxs-lookup"><span data-stu-id="bac6a-1009">getSelectedEntities() → {[Entities](/javascript/api/outlook_1_6/office.entities)}</span></span>

<span data-ttu-id="bac6a-p163">Ruft die Entitäten ab, die in einer hervorgehobenen Übereinstimmung gefunden werden, die ein Benutzer ausgewählt hat. Hervorgehobene Übereinstimmungen gelten für [Kontext-Add-Ins](https://docs.microsoft.com/outlook/add-ins/contextual-outlook-add-ins).</span><span class="sxs-lookup"><span data-stu-id="bac6a-p163">Gets the entities found in a highlighted match a user has selected. Highlighted matches apply to [contextual add-ins](https://docs.microsoft.com/outlook/add-ins/contextual-outlook-add-ins).</span></span>

> [!NOTE]
> <span data-ttu-id="bac6a-1012">Diese Methode wird in Outlook für iOS oder Outlook für Android nicht unterstützt.</span><span class="sxs-lookup"><span data-stu-id="bac6a-1012">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

##### <a name="requirements"></a><span data-ttu-id="bac6a-1013">Anforderungen</span><span class="sxs-lookup"><span data-stu-id="bac6a-1013">Requirements</span></span>

|<span data-ttu-id="bac6a-1014">Anforderung</span><span class="sxs-lookup"><span data-stu-id="bac6a-1014">Requirement</span></span>| <span data-ttu-id="bac6a-1015">Wert</span><span class="sxs-lookup"><span data-stu-id="bac6a-1015">Value</span></span>|
|---|---|
|[<span data-ttu-id="bac6a-1016">Mindestversion des Postfachanforderungssatzes</span><span class="sxs-lookup"><span data-stu-id="bac6a-1016">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="bac6a-1017">1.6</span><span class="sxs-lookup"><span data-stu-id="bac6a-1017">1.6</span></span> |
|[<span data-ttu-id="bac6a-1018">Mindestberechtigungsstufe</span><span class="sxs-lookup"><span data-stu-id="bac6a-1018">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="bac6a-1019">ReadItem</span><span class="sxs-lookup"><span data-stu-id="bac6a-1019">ReadItem</span></span>|
|[<span data-ttu-id="bac6a-1020">Zutreffender Outlook-Modus</span><span class="sxs-lookup"><span data-stu-id="bac6a-1020">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="bac6a-1021">Lesen</span><span class="sxs-lookup"><span data-stu-id="bac6a-1021">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="bac6a-1022">Gibt zurück:</span><span class="sxs-lookup"><span data-stu-id="bac6a-1022">Returns:</span></span>

<span data-ttu-id="bac6a-1023">Typ: [Entitäten](/javascript/api/outlook_1_6/office.entities)</span><span class="sxs-lookup"><span data-stu-id="bac6a-1023">Type: [Entities](/javascript/api/outlook_1_6/office.entities)</span></span>

##### <a name="example"></a><span data-ttu-id="bac6a-1024">Beispiel</span><span class="sxs-lookup"><span data-stu-id="bac6a-1024">Example</span></span>

<span data-ttu-id="bac6a-1025">Im folgenden Beispiel wird auf die Adressentitäten in der hervorgehobenen Übereinstimmung zugegriffen, die der Benutzer ausgewählt hat.</span><span class="sxs-lookup"><span data-stu-id="bac6a-1025">The following example accesses the addresses entities in the highlighted match selected by the user.</span></span>

```
var contacts = Office.context.mailbox.item.getSelectedEntities().addresses;
```

#### <a name="getselectedregexmatches--object"></a><span data-ttu-id="bac6a-1026">getSelectedRegExMatches() → {Object}</span><span class="sxs-lookup"><span data-stu-id="bac6a-1026">getSelectedRegExMatches() → {Object}</span></span>

<span data-ttu-id="bac6a-p164">Gibt Zeichenfolgenwerte in einer hervorgehobenen Übereinstimmung zurück, die den in der XML-Manifestdatei definierten regulären Ausdrücken entsprechen. Hervorgehobene Übereinstimmungen gelten für [Kontext-Add-Ins](https://docs.microsoft.com/outlook/add-ins/contextual-outlook-add-ins).</span><span class="sxs-lookup"><span data-stu-id="bac6a-p164">Returns string values in a highlighted match that match the regular expressions defined in the manifest XML file. Highlighted matches apply to [contextual add-ins](https://docs.microsoft.com/outlook/add-ins/contextual-outlook-add-ins).</span></span>

> [!NOTE]
> <span data-ttu-id="bac6a-1029">Diese Methode wird in Outlook für iOS oder Outlook für Android nicht unterstützt.</span><span class="sxs-lookup"><span data-stu-id="bac6a-1029">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="bac6a-p165">Die `getSelectedRegExMatches`-Methode gibt die Zeichenfolgen zurück, die dem im `ItemHasRegularExpressionMatch`- oder `ItemHasKnownEntity`-Regelelement der XML-Manifestdatei definierten regulären Ausdruck entsprechen. Bei einer `ItemHasRegularExpressionMatch`-Regel muss eine entsprechende Zeichenfolge in der Eigenschaft des Elements vorhanden sein, das von dieser Regel angegeben wird. Der einfache `PropertyName`-Typ definiert die unterstützten Eigenschaften.</span><span class="sxs-lookup"><span data-stu-id="bac6a-p165">The `getSelectedRegExMatches` method returns the strings that match the regular expression defined in each `ItemHasRegularExpressionMatch` or `ItemHasKnownEntity` rule element in the manifest XML file. For an `ItemHasRegularExpressionMatch` rule, a matching string has to occur in the property of the item that is specified by that rule. The `PropertyName` simple type defines the supported properties.</span></span>

<span data-ttu-id="bac6a-1033">Nehmen Sie z. B. an, dass ein Add-In-Manifest das folgende `Rule`-Element aufweist:</span><span class="sxs-lookup"><span data-stu-id="bac6a-1033">For example, consider an add-in manifest has the following `Rule` element:</span></span>

```
<Rule xsi:type="RuleCollection" Mode="And">
  <Rule xsi:type="ItemIs" FormType="Read" ItemType="Message" />
  <Rule xsi:type="RuleCollection" Mode="Or">
    <Rule xsi:type="ItemHasRegularExpressionMatch" RegExName="fruits" RegExValue="apple|banana|coconut" PropertyName="BodyAsPlaintext" IgnoreCase="true" />
    <Rule xsi:type="ItemHasRegularExpressionMatch" RegExName="veggies" RegExValue="tomato|onion|spinach|broccoli" PropertyName="BodyAsPlaintext" IgnoreCase="true" />
  </Rule>
</Rule>
```

<span data-ttu-id="bac6a-1034">Das von `getRegExMatches` zurückgegebene Objekt hätte zwei Eigenschaften: `fruits` und `veggies`.</span><span class="sxs-lookup"><span data-stu-id="bac6a-1034">The object returned from `getRegExMatches` would have two properties: `fruits` and `veggies`.</span></span>

```
{
  'fruits': ['apple','banana','Banana','coconut'],
  'veggies': ['tomato','onion','spinach','broccoli']
}
```

<span data-ttu-id="bac6a-p166">Wenn Sie eine `ItemHasRegularExpressionMatch`-Regel für die Textkörpereigenschaft eines Elements festlegen, sollte der reguläre Ausdruck den Textkörper weiter filtern und nicht versuchen, den gesamten Textkörper des Elements zurückzugeben. Wenn der gesamte Textkörper eines Elements mit einem regulären Ausdruck wie `.*` abgerufen wird, werden nicht immer die gewünschten Ergebnisse erzielt. Verwenden Sie stattdessen die [`Body.getAsync`](/javascript/api/outlook_1_6/office.body#getasync-coerciontype--options--callback-)-Methode, um den gesamten Textkörper abzurufen.</span><span class="sxs-lookup"><span data-stu-id="bac6a-p166">If you specify an `ItemHasRegularExpressionMatch` rule on the body property of an item, the regular expression should further filter the body and should not attempt to return the entire body of the item. Using a regular expression such as `.*` to obtain the entire body of an item does not always return the expected results. Instead, use the [`Body.getAsync`](/javascript/api/outlook_1_6/office.body#getasync-coerciontype--options--callback-) method to retrieve the entire body.</span></span>

##### <a name="requirements"></a><span data-ttu-id="bac6a-1038">Anforderungen</span><span class="sxs-lookup"><span data-stu-id="bac6a-1038">Requirements</span></span>

|<span data-ttu-id="bac6a-1039">Anforderung</span><span class="sxs-lookup"><span data-stu-id="bac6a-1039">Requirement</span></span>| <span data-ttu-id="bac6a-1040">Wert</span><span class="sxs-lookup"><span data-stu-id="bac6a-1040">Value</span></span>|
|---|---|
|[<span data-ttu-id="bac6a-1041">Mindestversion des Postfachanforderungssatzes</span><span class="sxs-lookup"><span data-stu-id="bac6a-1041">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="bac6a-1042">1.6</span><span class="sxs-lookup"><span data-stu-id="bac6a-1042">1.6</span></span> |
|[<span data-ttu-id="bac6a-1043">Mindestberechtigungsstufe</span><span class="sxs-lookup"><span data-stu-id="bac6a-1043">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="bac6a-1044">ReadItem</span><span class="sxs-lookup"><span data-stu-id="bac6a-1044">ReadItem</span></span>|
|[<span data-ttu-id="bac6a-1045">Zutreffender Outlook-Modus</span><span class="sxs-lookup"><span data-stu-id="bac6a-1045">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="bac6a-1046">Lesen</span><span class="sxs-lookup"><span data-stu-id="bac6a-1046">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="bac6a-1047">Gibt zurück:</span><span class="sxs-lookup"><span data-stu-id="bac6a-1047">Returns:</span></span>

<span data-ttu-id="bac6a-p167">Ein Objekt mit Arrays aus Zeichenfolgen, die den in der XML-Manifestdatei definierten regulären Ausdrücken entsprechen. Der Name der einzelnen Arrays ist gleich dem entsprechenden Wert des `RegExName`-Attributs der entsprechenden `ItemHasRegularExpressionMatch`-Regel oder des `FilterName`-Attributs der entsprechenden `ItemHasKnownEntity`-Regel.</span><span class="sxs-lookup"><span data-stu-id="bac6a-p167">An object that contains arrays of strings that match the regular expressions defined in the manifest XML file. The name of each array is equal to the corresponding value of the `RegExName` attribute of the matching `ItemHasRegularExpressionMatch` rule or the `FilterName` attribute of the matching `ItemHasKnownEntity` rule.</span></span>

##### <a name="example"></a><span data-ttu-id="bac6a-1050">Beispiel</span><span class="sxs-lookup"><span data-stu-id="bac6a-1050">Example</span></span>

<span data-ttu-id="bac6a-1051">Das folgende Beispiel zeigt, wie Sie auf das Array von Übereinstimmungen für die regulären Ausdrucksregelelemente `fruits` und `veggies` zugreifen, die im Manifest angegeben sind.</span><span class="sxs-lookup"><span data-stu-id="bac6a-1051">The following example shows how to access the array of matches for the regular expression rule elements `fruits` and `veggies`, which are specified in the manifest.</span></span>

```
var selectedMatches = Office.context.mailbox.item.getSelectedRegExMatches();
var fruits = selectedMatches.fruits;
var veggies = selectedMatches.veggies;
```

####  <a name="loadcustompropertiesasynccallback-usercontext"></a><span data-ttu-id="bac6a-1052">loadCustomPropertiesAsync(callback, [userContext])</span><span class="sxs-lookup"><span data-stu-id="bac6a-1052">loadCustomPropertiesAsync(callback, [userContext])</span></span>

<span data-ttu-id="bac6a-1053">Lädt asynchron benutzerdefinierte Eigenschaften für dieses Add-In für das ausgewählte Element.</span><span class="sxs-lookup"><span data-stu-id="bac6a-1053">Asynchronously loads custom properties for this add-in on the selected item.</span></span>

<span data-ttu-id="bac6a-p168">Benutzerdefinierte Eigenschaften werden als Schlüssel-/Wert-Paare pro App und pro Element gespeichert. Diese Methode gibt ein `CustomProperties`-Objekt im Rückruf zurück, das Methoden für den Zugriff auf die benutzerdefinierten Eigenschaften für das aktuelle Element und das aktuelle Add-In bietet. Benutzerdefinierte Eigenschaften sind für das Element nicht verschlüsselt und sollten darum nicht als sicherer Speicher verwendet werden.</span><span class="sxs-lookup"><span data-stu-id="bac6a-p168">Custom properties are stored as key/value pairs on a per-app, per-item basis. This method returns a `CustomProperties` object in the callback, which provides methods to access the custom properties specific to the current item and the current add-in. Custom properties are not encrypted on the item, so this should not be used as secure storage.</span></span>

##### <a name="parameters"></a><span data-ttu-id="bac6a-1057">Parameter:</span><span class="sxs-lookup"><span data-stu-id="bac6a-1057">Parameters:</span></span>

|<span data-ttu-id="bac6a-1058">Name</span><span class="sxs-lookup"><span data-stu-id="bac6a-1058">Name</span></span>| <span data-ttu-id="bac6a-1059">Typ</span><span class="sxs-lookup"><span data-stu-id="bac6a-1059">Type</span></span>| <span data-ttu-id="bac6a-1060">Attribute</span><span class="sxs-lookup"><span data-stu-id="bac6a-1060">Attributes</span></span>| <span data-ttu-id="bac6a-1061">Beschreibung</span><span class="sxs-lookup"><span data-stu-id="bac6a-1061">Description</span></span>|
|---|---|---|---|
|`callback`| <span data-ttu-id="bac6a-1062">Funktion</span><span class="sxs-lookup"><span data-stu-id="bac6a-1062">function</span></span>||<span data-ttu-id="bac6a-1063">Wenn die Methode abgeschlossen wird, wird die Funktion , die im `callback`-Parameter übergeben wird, mit einem einzelnen Parameter (`asyncResult`) aufgerufen, der ein [`AsyncResult`](/javascript/api/office/office.asyncresult)-Objekt darstellt.</span><span class="sxs-lookup"><span data-stu-id="bac6a-1063">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span><br/><br/><span data-ttu-id="bac6a-1064">Die benutzerdefinierten Eigenschaften werden als [`CustomProperties`](/javascript/api/outlook_1_6/office.customproperties)-Objekt in der `asyncResult.value`-Eigenschaft bereitgestellt.</span><span class="sxs-lookup"><span data-stu-id="bac6a-1064">The custom properties are provided as a [`CustomProperties`](/javascript/api/outlook_1_6/office.customproperties) object in the `asyncResult.value` property.</span></span> <span data-ttu-id="bac6a-1065">Dieses Objekt kann zum Abrufen, festlegen und Entfernen benutzerdefinierter Eigenschaften aus dem Element und speichern Sie Änderungen an den benutzerdefinierten Eigenschaftensatz zurück an den Server verwendet werden.</span><span class="sxs-lookup"><span data-stu-id="bac6a-1065">This object can be used to get, set, and remove custom properties from the item and save changes to the custom property set back to the server.</span></span>|
|`userContext`| <span data-ttu-id="bac6a-1066">Objekt</span><span class="sxs-lookup"><span data-stu-id="bac6a-1066">Object</span></span>| <span data-ttu-id="bac6a-1067">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="bac6a-1067">&lt;optional&gt;</span></span>|<span data-ttu-id="bac6a-1068">Entwickler können ein beliebiges Objekt bereitstellen, den sie in der Rückruffunktion zugreifen möchten.</span><span class="sxs-lookup"><span data-stu-id="bac6a-1068">Developers can provide any object they wish to access in the callback function.</span></span> <span data-ttu-id="bac6a-1069">Dieses Objekt zugegriffen werden kann, indem die `asyncResult.asyncContext` -Eigenschaft in der Callback-Funktion.</span><span class="sxs-lookup"><span data-stu-id="bac6a-1069">This object can be accessed by the `asyncResult.asyncContext` property in the callback function.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="bac6a-1070">Anforderungen</span><span class="sxs-lookup"><span data-stu-id="bac6a-1070">Requirements</span></span>

|<span data-ttu-id="bac6a-1071">Anforderung</span><span class="sxs-lookup"><span data-stu-id="bac6a-1071">Requirement</span></span>| <span data-ttu-id="bac6a-1072">Wert</span><span class="sxs-lookup"><span data-stu-id="bac6a-1072">Value</span></span>|
|---|---|
|[<span data-ttu-id="bac6a-1073">Mindestversion des Postfachanforderungssatzes</span><span class="sxs-lookup"><span data-stu-id="bac6a-1073">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="bac6a-1074">1.0</span><span class="sxs-lookup"><span data-stu-id="bac6a-1074">1.0</span></span>|
|[<span data-ttu-id="bac6a-1075">Mindestberechtigungsstufe</span><span class="sxs-lookup"><span data-stu-id="bac6a-1075">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="bac6a-1076">ReadItem</span><span class="sxs-lookup"><span data-stu-id="bac6a-1076">ReadItem</span></span>|
|[<span data-ttu-id="bac6a-1077">Zutreffender Outlook-Modus</span><span class="sxs-lookup"><span data-stu-id="bac6a-1077">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="bac6a-1078">Verfassen oder Lesen</span><span class="sxs-lookup"><span data-stu-id="bac6a-1078">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="bac6a-1079">Beispiel</span><span class="sxs-lookup"><span data-stu-id="bac6a-1079">Example</span></span>

<span data-ttu-id="bac6a-p171">Das folgende Codebeispiel veranschaulicht die Verwendung der `loadCustomPropertiesAsync`-Methode zum asynchronen Laden der benutzerdefinierten Eigenschaften, die für das aktuelle Element spezifisch sind. Des Weiteren veranschaulicht das Beispiel die Verwendung der `CustomProperties.saveAsync`-Methode zum Speichern der Eigenschaften auf dem Server. Nach dem Laden der benutzerdefinierten Eigenschaften wird in dem Codebeispiel die `CustomProperties.get`-Methode dazu verwendet, die benutzerdefinierte `myProp`-Eigenschaft zu lesen. Die `CustomProperties.set`-Methode wird dazu verwendet, die benutzerdefinierte `otherProp`-Eigenschaft zu schreiben. Schließlich wird die `saveAsync`-Methode aufgerufen, um die benutzerdefinierten Eigenschaften zu speichern.</span><span class="sxs-lookup"><span data-stu-id="bac6a-p171">The following code example shows how to use the `loadCustomPropertiesAsync` method to asynchronously load custom properties that are specific to the current item. The example also shows how to use the `CustomProperties.saveAsync` method to save these properties back to the server. After loading the custom properties, the code sample uses the `CustomProperties.get` method to read the custom property `myProp`, the `CustomProperties.set` method to write the custom property `otherProp`, and then finally calls the `saveAsync` method to save the custom properties.</span></span>

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

####  <a name="removeattachmentasyncattachmentid-options-callback"></a><span data-ttu-id="bac6a-1083">removeAttachmentAsync(attachmentId, [options], [callback])</span><span class="sxs-lookup"><span data-stu-id="bac6a-1083">removeAttachmentAsync(attachmentId, [options], [callback])</span></span>

<span data-ttu-id="bac6a-1084">Entfernt eine Anlage aus einer Nachricht oder einem Termin.</span><span class="sxs-lookup"><span data-stu-id="bac6a-1084">Removes an attachment from a message or appointment.</span></span>

<span data-ttu-id="bac6a-p172">Die `removeAttachmentAsync`-Methode entfernt die Anlage mit dem angegebenen Bezeichner aus dem Element. Als bewährte Vorgehensweise sollten Sie den Anlagenbezeichner nur dann zum Entfernen einer Anlage verwenden, wenn die gleiche Mail-App die Anlage in der gleichen Sitzung hinzugefügt hat. In Outlook Web App und OWA für Geräte ist der Anlagenbezeichner nur innerhalb der gleichen Sitzung gültig. Eine Sitzung ist abgeschlossen, wenn der Benutzer die App schließt, oder wenn der Benutzer in einem eingebetteten Formular mit dem Verfassen beginnt und den Vorgang dann in einem separaten Fenster fortsetzt.</span><span class="sxs-lookup"><span data-stu-id="bac6a-p172">The `removeAttachmentAsync` method removes the attachment with the specified identifier from the item. As a best practice, you should use the attachment identifier to remove an attachment only if the same mail app has added that attachment in the same session. In Outlook Web App and OWA for Devices, the attachment identifier is valid only within the same session. A session is over when the user closes the app, or if the user starts composing in an inline form and subsequently pops out the inline form to continue in a separate window.</span></span>

##### <a name="parameters"></a><span data-ttu-id="bac6a-1089">Parameter:</span><span class="sxs-lookup"><span data-stu-id="bac6a-1089">Parameters:</span></span>

|<span data-ttu-id="bac6a-1090">Name</span><span class="sxs-lookup"><span data-stu-id="bac6a-1090">Name</span></span>| <span data-ttu-id="bac6a-1091">Typ</span><span class="sxs-lookup"><span data-stu-id="bac6a-1091">Type</span></span>| <span data-ttu-id="bac6a-1092">Attribute</span><span class="sxs-lookup"><span data-stu-id="bac6a-1092">Attributes</span></span>| <span data-ttu-id="bac6a-1093">Beschreibung</span><span class="sxs-lookup"><span data-stu-id="bac6a-1093">Description</span></span>|
|---|---|---|---|
|`attachmentId`| <span data-ttu-id="bac6a-1094">Zeichenfolge</span><span class="sxs-lookup"><span data-stu-id="bac6a-1094">String</span></span>||<span data-ttu-id="bac6a-p173">Der Bezeichner des Anhangs, der entfernt werden soll. Die maximale Länge der Zeichenfolge ist 100 Zeichen.</span><span class="sxs-lookup"><span data-stu-id="bac6a-p173">The identifier of the attachment to remove. The maximum length of the string is 100 characters.</span></span>|
|`options`| <span data-ttu-id="bac6a-1097">Object</span><span class="sxs-lookup"><span data-stu-id="bac6a-1097">Object</span></span>| <span data-ttu-id="bac6a-1098">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="bac6a-1098">&lt;optional&gt;</span></span>|<span data-ttu-id="bac6a-1099">Ein Objektliteral, das mindestens eine der folgenden Eigenschaften enthält.</span><span class="sxs-lookup"><span data-stu-id="bac6a-1099">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`| <span data-ttu-id="bac6a-1100">Object</span><span class="sxs-lookup"><span data-stu-id="bac6a-1100">Object</span></span>| <span data-ttu-id="bac6a-1101">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="bac6a-1101">&lt;optional&gt;</span></span>|<span data-ttu-id="bac6a-1102">Entwickler können ein Objekt bereitstellen, auf das sie in der Callbackmethode zugreifen möchten.</span><span class="sxs-lookup"><span data-stu-id="bac6a-1102">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`| <span data-ttu-id="bac6a-1103">Funktion</span><span class="sxs-lookup"><span data-stu-id="bac6a-1103">function</span></span>| <span data-ttu-id="bac6a-1104">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="bac6a-1104">&lt;optional&gt;</span></span>|<span data-ttu-id="bac6a-1105">Wenn die Methode abgeschlossen wird, wird die Funktion , die im `callback`-Parameter übergeben wird, mit einem einzelnen Parameter (`asyncResult`) aufgerufen, der ein [`AsyncResult`](/javascript/api/office/office.asyncresult)-Objekt darstellt.</span><span class="sxs-lookup"><span data-stu-id="bac6a-1105">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span> <br/><span data-ttu-id="bac6a-1106">Wenn beim Entfernen der Anlage ein Fehler auftritt, enthält die Eigenschaft `asyncResult.error` einen Fehlercode mit dem Grund für den Fehler.</span><span class="sxs-lookup"><span data-stu-id="bac6a-1106">If removing the attachment fails, the `asyncResult.error` property will contain an error code with the reason for the failure.</span></span>|

##### <a name="errors"></a><span data-ttu-id="bac6a-1107">Fehler</span><span class="sxs-lookup"><span data-stu-id="bac6a-1107">Errors</span></span>

| <span data-ttu-id="bac6a-1108">Fehlercode</span><span class="sxs-lookup"><span data-stu-id="bac6a-1108">Error code</span></span> | <span data-ttu-id="bac6a-1109">Beschreibung</span><span class="sxs-lookup"><span data-stu-id="bac6a-1109">Description</span></span> |
|------------|-------------|
| `InvalidAttachmentId` | <span data-ttu-id="bac6a-1110">Der Bezeichner für die Anlage ist nicht vorhanden.</span><span class="sxs-lookup"><span data-stu-id="bac6a-1110">The attachment identifier does not exist.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="bac6a-1111">Anforderungen</span><span class="sxs-lookup"><span data-stu-id="bac6a-1111">Requirements</span></span>

|<span data-ttu-id="bac6a-1112">Anforderung</span><span class="sxs-lookup"><span data-stu-id="bac6a-1112">Requirement</span></span>| <span data-ttu-id="bac6a-1113">Wert</span><span class="sxs-lookup"><span data-stu-id="bac6a-1113">Value</span></span>|
|---|---|
|[<span data-ttu-id="bac6a-1114">Mindestversion des Postfachanforderungssatzes</span><span class="sxs-lookup"><span data-stu-id="bac6a-1114">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="bac6a-1115">1.1</span><span class="sxs-lookup"><span data-stu-id="bac6a-1115">1.1</span></span>|
|[<span data-ttu-id="bac6a-1116">Mindestberechtigungsstufe</span><span class="sxs-lookup"><span data-stu-id="bac6a-1116">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="bac6a-1117">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="bac6a-1117">ReadWriteItem</span></span>|
|[<span data-ttu-id="bac6a-1118">Zutreffender Outlook-Modus</span><span class="sxs-lookup"><span data-stu-id="bac6a-1118">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="bac6a-1119">Verfassen</span><span class="sxs-lookup"><span data-stu-id="bac6a-1119">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="bac6a-1120">Beispiel</span><span class="sxs-lookup"><span data-stu-id="bac6a-1120">Example</span></span>

<span data-ttu-id="bac6a-1121">Mit dem folgende Code wird eine Anlage mit dem Bezeichner "0" entfernt.</span><span class="sxs-lookup"><span data-stu-id="bac6a-1121">The following code removes an attachment with an identifier of '0'.</span></span>

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

####  <a name="saveasyncoptions-callback"></a><span data-ttu-id="bac6a-1122">saveAsync([options], callback)</span><span class="sxs-lookup"><span data-stu-id="bac6a-1122">saveAsync([options], callback)</span></span>

<span data-ttu-id="bac6a-1123">Speicher asynchron ein Element. .</span><span class="sxs-lookup"><span data-stu-id="bac6a-1123">Asynchronously saves an item.</span></span>

<span data-ttu-id="bac6a-p174">Beim Aufrufen speichert diese Methode die aktuelle Nachricht als Entwurf und  gibt die Element-ID über die Callbackmethode zurück. In Outlook Web App oder Outlook im Onlinemodus wird das Element auf dem Server gespeichert. In Outlook im Cache-Modus wird das Element im lokalen Cache gespeichert.</span><span class="sxs-lookup"><span data-stu-id="bac6a-p174">When invoked, this method saves the current message as a draft and returns the item id via the callback method. In Outlook Web App or Outlook in online mode, the item is saved to the server. In Outlook in cached mode, the item is saved to the local cache.</span></span>

> [!NOTE]
> <span data-ttu-id="bac6a-1127">Wenn Ihr Add-In ruft `saveAsync` für ein Element im Entwurfsmodus, um das Abrufen einer `itemId` um mit EWS oder REST-API verwenden, beachten Sie, dass wenn Outlook im Cache-Modus ist, es kann einige Zeit dauern, bevor das Element tatsächlich auf dem Server synchronisiert wird.</span><span class="sxs-lookup"><span data-stu-id="bac6a-1127">If your add-in calls `saveAsync` on an item in compose mode in order to get an `itemId` to use with EWS or the REST API, be aware that when Outlook is in cached mode, it may take some time before the item is actually synced to the server.</span></span> <span data-ttu-id="bac6a-1128">Bis das Element mit synchronisiert ist, die `itemId` wird ein Fehler zurückgegeben.</span><span class="sxs-lookup"><span data-stu-id="bac6a-1128">Until the item is synced, using the `itemId` will return an error.</span></span>

<span data-ttu-id="bac6a-p176">Da Termine keinen Entwurfsstatus haben, wird das Element bei Aufruf von `saveAsync` für einen Termin im Verfassenmodus als normaler Termin im Kalender des Benutzers gespeichert. Für neue Termine, die noch nicht gespeichert wurden, wird keine Einladung gesendet. Beim Speichern eines vorhandenen Termins wird eine Aktualisierung an die hinzugefügten oder entfernten Teilnehmer gesendet.</span><span class="sxs-lookup"><span data-stu-id="bac6a-p176">Since appointments have no draft state, if `saveAsync` is called on an appointment in compose mode, the item will be saved as a normal appointment on the user's calendar. For new appointments that have not been saved before, no invitation will be sent. Saving an existing appointment will send an update to added or removed attendees.</span></span>

> [!NOTE]
> <span data-ttu-id="bac6a-1132">Die folgenden Clients haben unterschiedlichem Verhalten für `saveAsync` Termine im Verfassenmodus:</span><span class="sxs-lookup"><span data-stu-id="bac6a-1132">The following clients have different behavior for `saveAsync` on appointments in compose mode:</span></span>
>
> - <span data-ttu-id="bac6a-1133">Outlook für Mac unterstützt keine `saveAsync` auf einer Besprechung im Verfassenmodus.</span><span class="sxs-lookup"><span data-stu-id="bac6a-1133">Mac Outlook does not support `saveAsync` on a meeting in compose mode.</span></span> <span data-ttu-id="bac6a-1134">Aufrufen von `saveAsync` auf einer Besprechung in Outlook für Mac wird ein Fehler zurückgegeben.</span><span class="sxs-lookup"><span data-stu-id="bac6a-1134">Calling `saveAsync` on a meeting in Mac Outlook will return an error.</span></span>
> - <span data-ttu-id="bac6a-1135">Outlook im Web immer sendet eine Einladung oder zu aktualisieren, wenn `saveAsync` für einen Termin aufgerufen wird, im Verfassenmodus.</span><span class="sxs-lookup"><span data-stu-id="bac6a-1135">Outlook on the web always sends an invitation or update when `saveAsync` is called on an appointment in compose mode.</span></span>

##### <a name="parameters"></a><span data-ttu-id="bac6a-1136">Parameter:</span><span class="sxs-lookup"><span data-stu-id="bac6a-1136">Parameters:</span></span>

|<span data-ttu-id="bac6a-1137">Name</span><span class="sxs-lookup"><span data-stu-id="bac6a-1137">Name</span></span>| <span data-ttu-id="bac6a-1138">Typ</span><span class="sxs-lookup"><span data-stu-id="bac6a-1138">Type</span></span>| <span data-ttu-id="bac6a-1139">Attribute</span><span class="sxs-lookup"><span data-stu-id="bac6a-1139">Attributes</span></span>| <span data-ttu-id="bac6a-1140">Beschreibung</span><span class="sxs-lookup"><span data-stu-id="bac6a-1140">Description</span></span>|
|---|---|---|---|
|`options`| <span data-ttu-id="bac6a-1141">Objekt</span><span class="sxs-lookup"><span data-stu-id="bac6a-1141">Object</span></span>| <span data-ttu-id="bac6a-1142">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="bac6a-1142">&lt;optional&gt;</span></span>|<span data-ttu-id="bac6a-1143">Ein Objektliteral, das mindestens eine der folgenden Eigenschaften enthält.</span><span class="sxs-lookup"><span data-stu-id="bac6a-1143">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`| <span data-ttu-id="bac6a-1144">Object</span><span class="sxs-lookup"><span data-stu-id="bac6a-1144">Object</span></span>| <span data-ttu-id="bac6a-1145">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="bac6a-1145">&lt;optional&gt;</span></span>|<span data-ttu-id="bac6a-1146">Entwickler können ein Objekt bereitstellen, auf das sie in der Callbackmethode zugreifen möchten.</span><span class="sxs-lookup"><span data-stu-id="bac6a-1146">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`| <span data-ttu-id="bac6a-1147">Funktion</span><span class="sxs-lookup"><span data-stu-id="bac6a-1147">function</span></span>||<span data-ttu-id="bac6a-1148">Wenn die Methode abgeschlossen wird, wird die Funktion , die im `callback`-Parameter übergeben wird, mit einem einzelnen Parameter (`asyncResult`) aufgerufen, der ein [`AsyncResult`](/javascript/api/office/office.asyncresult)-Objekt darstellt.</span><span class="sxs-lookup"><span data-stu-id="bac6a-1148">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span><br/><br/><span data-ttu-id="bac6a-1149">Auf Erfolg, erfolgt die Element-ID der `asyncResult.value` Eigenschaft.</span><span class="sxs-lookup"><span data-stu-id="bac6a-1149">On success, the item identifier is provided in the `asyncResult.value` property.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="bac6a-1150">Anforderungen</span><span class="sxs-lookup"><span data-stu-id="bac6a-1150">Requirements</span></span>

|<span data-ttu-id="bac6a-1151">Anforderung</span><span class="sxs-lookup"><span data-stu-id="bac6a-1151">Requirement</span></span>| <span data-ttu-id="bac6a-1152">Wert</span><span class="sxs-lookup"><span data-stu-id="bac6a-1152">Value</span></span>|
|---|---|
|[<span data-ttu-id="bac6a-1153">Mindestversion des Postfachanforderungssatzes</span><span class="sxs-lookup"><span data-stu-id="bac6a-1153">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="bac6a-1154">1.3</span><span class="sxs-lookup"><span data-stu-id="bac6a-1154">1.3</span></span>|
|[<span data-ttu-id="bac6a-1155">Mindestberechtigungsstufe</span><span class="sxs-lookup"><span data-stu-id="bac6a-1155">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="bac6a-1156">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="bac6a-1156">ReadWriteItem</span></span>|
|[<span data-ttu-id="bac6a-1157">Zutreffender Outlook-Modus</span><span class="sxs-lookup"><span data-stu-id="bac6a-1157">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="bac6a-1158">Verfassen</span><span class="sxs-lookup"><span data-stu-id="bac6a-1158">Compose</span></span>|

##### <a name="examples"></a><span data-ttu-id="bac6a-1159">Beispiele</span><span class="sxs-lookup"><span data-stu-id="bac6a-1159">Examples</span></span>

```
Office.context.mailbox.item.saveAsync(
  function callback(result) {
    // Process the result
  });
```

<span data-ttu-id="bac6a-p178">Es folgt ein Beispiel für den `result`-Parameter, der an die Callbackfunktion übergeben wird. Die `value`-Eigenschaft enthält die Element-ID des Elements.</span><span class="sxs-lookup"><span data-stu-id="bac6a-p178">The following is an example of the `result` parameter passed to the callback function. The `value` property contains the item ID of the item.</span></span>

```
{
  "value":"AAMkADI5...AAA=",
  "status":"succeeded"
}
```

####  <a name="setselecteddataasyncdata-options-callback"></a><span data-ttu-id="bac6a-1162">setSelectedDataAsync(data, [options], callback)</span><span class="sxs-lookup"><span data-stu-id="bac6a-1162">setSelectedDataAsync(data, [options], callback)</span></span>

<span data-ttu-id="bac6a-1163">Fügt asynchron Daten in den Textkörper oder Betreff einer Nachricht ein.</span><span class="sxs-lookup"><span data-stu-id="bac6a-1163">Asynchronously inserts data into the body or subject of a message.</span></span>

<span data-ttu-id="bac6a-p179">Die `setSelectedDataAsync`-Methode fügt die angegebene Zeichenfolge an der Cursorposition im Betreff oder im Textkörper ein, oder, falls im Editor Text ausgewählt ist, ersetzt den markierten Text. Wenn sich der Cursor nicht im Textkörper oder im Betreffsfeld befindet, wird ein Fehler zurückgegeben. Nach dem Einfügen wird der Cursor am Ende der eingefügten Inhalte platziert.</span><span class="sxs-lookup"><span data-stu-id="bac6a-p179">The `setSelectedDataAsync` method inserts the specified string at the cursor location in the subject or body of the item, or, if text is selected in the editor, it replaces the selected text. If the cursor is not in the body or subject field, an error is returned. After insertion, the cursor is placed at the end of the inserted content.</span></span>

##### <a name="parameters"></a><span data-ttu-id="bac6a-1167">Parameter:</span><span class="sxs-lookup"><span data-stu-id="bac6a-1167">Parameters:</span></span>

|<span data-ttu-id="bac6a-1168">Name</span><span class="sxs-lookup"><span data-stu-id="bac6a-1168">Name</span></span>| <span data-ttu-id="bac6a-1169">Typ</span><span class="sxs-lookup"><span data-stu-id="bac6a-1169">Type</span></span>| <span data-ttu-id="bac6a-1170">Attribute</span><span class="sxs-lookup"><span data-stu-id="bac6a-1170">Attributes</span></span>| <span data-ttu-id="bac6a-1171">Beschreibung</span><span class="sxs-lookup"><span data-stu-id="bac6a-1171">Description</span></span>|
|---|---|---|---|
|`data`| <span data-ttu-id="bac6a-1172">Zeichenfolge</span><span class="sxs-lookup"><span data-stu-id="bac6a-1172">String</span></span>||<span data-ttu-id="bac6a-p180">Die einzufügenden Daten. Daten dürfen 1.000.000 Zeichen nicht überschreiten. Werden mehr als 1.000.000 Zeichen übergeben, wird eine `ArgumentOutOfRange`-Ausnahme ausgelöst.</span><span class="sxs-lookup"><span data-stu-id="bac6a-p180">The data to be inserted. Data is not to exceed 1,000,000 characters. If more than 1,000,000 characters are passed in, an `ArgumentOutOfRange` exception is thrown.</span></span>|
|`options`| <span data-ttu-id="bac6a-1176">Object</span><span class="sxs-lookup"><span data-stu-id="bac6a-1176">Object</span></span>| <span data-ttu-id="bac6a-1177">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="bac6a-1177">&lt;optional&gt;</span></span>|<span data-ttu-id="bac6a-1178">Ein Objektliteral, das mindestens eine der folgenden Eigenschaften enthält.</span><span class="sxs-lookup"><span data-stu-id="bac6a-1178">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`| <span data-ttu-id="bac6a-1179">Object</span><span class="sxs-lookup"><span data-stu-id="bac6a-1179">Object</span></span>| <span data-ttu-id="bac6a-1180">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="bac6a-1180">&lt;optional&gt;</span></span>|<span data-ttu-id="bac6a-1181">Entwickler können ein Objekt bereitstellen, auf das sie in der Rückrufmethode zugreifen möchten.</span><span class="sxs-lookup"><span data-stu-id="bac6a-1181">Developers can provide any object they wish to access in the callback method.</span></span>|
|`options.coercionType`| [<span data-ttu-id="bac6a-1182">Office.CoercionType</span><span class="sxs-lookup"><span data-stu-id="bac6a-1182">Office.CoercionType</span></span>](office.md#coerciontype-string)| <span data-ttu-id="bac6a-1183">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="bac6a-1183">&lt;optional&gt;</span></span>|<span data-ttu-id="bac6a-p181">Wenn `text`, wird das aktuelle Format in Outlook Web App und Outlook angewendet. Wenn das Feld ein HTML-Editor ist, werden nur die Textdaten eingefügt, selbst wenn es sich bei den Daten um HTML-Daten handelt.</span><span class="sxs-lookup"><span data-stu-id="bac6a-p181">If `text`, the current style is applied in Outlook Web App and Outlook. If the field is an HTML editor, only the text data is inserted, even if the data is HTML.</span></span><br/><br/><span data-ttu-id="bac6a-p182">Wenn `html` gesetzt ist und das Feld HTML unterstützt (der Betreff jedoch nicht), wird die aktuelle Formatvorlage in Outlook Web App angewendet und die Standardformatvorlage in Outlook. Ist das Feld ein Textfeld, wird ein Fehler des Typs `InvalidDataFormat` zurückgegeben.</span><span class="sxs-lookup"><span data-stu-id="bac6a-p182">If `html` and the field supports HTML (the subject doesn't), the current style is applied in Outlook Web App and the default style is applied in Outlook. If the field is a text field, an `InvalidDataFormat` error is returned.</span></span><br/><br/><span data-ttu-id="bac6a-1188">Wenn `coercionType` nicht festgelegt wird, hängt das Ergebnis vom Feld ab: Wenn das Feld HTML ist, wird HTML verwendet. Wenn das Feld Text ist, wird Nur-Text verwendet.</span><span class="sxs-lookup"><span data-stu-id="bac6a-1188">If `coercionType` is not set, the result depends on the field: if the field is HTML then HTML is used; if the field is text, then plain text is used.</span></span>|
|`callback`| <span data-ttu-id="bac6a-1189">function</span><span class="sxs-lookup"><span data-stu-id="bac6a-1189">function</span></span>||<span data-ttu-id="bac6a-1190">Wenn die Methode abgeschlossen wird, wird die Funktion , die im `callback`-Parameter übergeben wird, mit einem einzelnen Parameter (`asyncResult`) aufgerufen, der ein [`AsyncResult`](/javascript/api/office/office.asyncresult)-Objekt darstellt.</span><span class="sxs-lookup"><span data-stu-id="bac6a-1190">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="bac6a-1191">Anforderungen</span><span class="sxs-lookup"><span data-stu-id="bac6a-1191">Requirements</span></span>

|<span data-ttu-id="bac6a-1192">Anforderung</span><span class="sxs-lookup"><span data-stu-id="bac6a-1192">Requirement</span></span>| <span data-ttu-id="bac6a-1193">Wert</span><span class="sxs-lookup"><span data-stu-id="bac6a-1193">Value</span></span>|
|---|---|
|[<span data-ttu-id="bac6a-1194">Mindestversion des Postfachanforderungssatzes</span><span class="sxs-lookup"><span data-stu-id="bac6a-1194">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="bac6a-1195">1.2</span><span class="sxs-lookup"><span data-stu-id="bac6a-1195">1.2</span></span>|
|[<span data-ttu-id="bac6a-1196">Mindestberechtigungsstufe</span><span class="sxs-lookup"><span data-stu-id="bac6a-1196">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="bac6a-1197">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="bac6a-1197">ReadWriteItem</span></span>|
|[<span data-ttu-id="bac6a-1198">Zutreffender Outlook-Modus</span><span class="sxs-lookup"><span data-stu-id="bac6a-1198">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="bac6a-1199">Verfassen</span><span class="sxs-lookup"><span data-stu-id="bac6a-1199">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="bac6a-1200">Beispiel</span><span class="sxs-lookup"><span data-stu-id="bac6a-1200">Example</span></span>

```
Office.context.mailbox.item.setSelectedDataAsync("Hello World!");
Office.context.mailbox.item.setSelectedDataAsync("<b>Hello World!</b>", { coercionType : "html" });
```