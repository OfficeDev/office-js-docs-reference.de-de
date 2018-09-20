
# <a name="item"></a><span data-ttu-id="406e6-101">item</span><span class="sxs-lookup"><span data-stu-id="406e6-101">item</span></span>

### <a name="officeofficemdcontextofficecontextmdmailboxofficecontextmailboxmditem"></a><span data-ttu-id="406e6-102">[Office](office.md)[.context](office.context.md)[.mailbox](office.context.mailbox.md).item</span><span class="sxs-lookup"><span data-stu-id="406e6-102">[Office](office.md)[.context](office.context.md)[.mailbox](office.context.mailbox.md).item</span></span>

<span data-ttu-id="406e6-p101">Der `item`-Namespace wird für den Zugriff auf die aktuell ausgewählte Nachricht, Besprechungsanfrage oder den aktuell ausgewählten Termin verwendet. Sie können den Typ von `item` mithilfe der [itemType](#itemtype-officemailboxenumsitemtypejavascriptapioutlook15officemailboxenumsitemtype)-Eigenschaft bestimmen.</span><span class="sxs-lookup"><span data-stu-id="406e6-p101">The `item` namespace is used to access the currently selected message, meeting request, or appointment. You can determine the type of the `item` by using the [itemType](#itemtype-officemailboxenumsitemtypejavascriptapioutlook15officemailboxenumsitemtype) property.</span></span>

##### <a name="requirements"></a><span data-ttu-id="406e6-105">Anforderungen</span><span class="sxs-lookup"><span data-stu-id="406e6-105">Requirements</span></span>

|<span data-ttu-id="406e6-106">Anforderung</span><span class="sxs-lookup"><span data-stu-id="406e6-106">Requirement</span></span>| <span data-ttu-id="406e6-107">Wert</span><span class="sxs-lookup"><span data-stu-id="406e6-107">Value</span></span>|
|---|---|
|[<span data-ttu-id="406e6-108">Mindestversion des Postfachanforderungssatzes</span><span class="sxs-lookup"><span data-stu-id="406e6-108">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="406e6-109">1.0</span><span class="sxs-lookup"><span data-stu-id="406e6-109">1.0</span></span>|
|[<span data-ttu-id="406e6-110">Mindestberechtigungsstufe</span><span class="sxs-lookup"><span data-stu-id="406e6-110">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="406e6-111">Eingeschränkt</span><span class="sxs-lookup"><span data-stu-id="406e6-111">Restricted</span></span>|
|[<span data-ttu-id="406e6-112">Zutreffender Outlook-Modus</span><span class="sxs-lookup"><span data-stu-id="406e6-112">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="406e6-113">Verfassen oder Lesen</span><span class="sxs-lookup"><span data-stu-id="406e6-113">Compose or read</span></span>|

##### <a name="members-and-methods"></a><span data-ttu-id="406e6-114">Elemente und Methoden</span><span class="sxs-lookup"><span data-stu-id="406e6-114">Members and methods</span></span>

| <span data-ttu-id="406e6-115">Element</span><span class="sxs-lookup"><span data-stu-id="406e6-115">Member</span></span> | <span data-ttu-id="406e6-116">Typ</span><span class="sxs-lookup"><span data-stu-id="406e6-116">Type</span></span> |
|--------|------|
| [<span data-ttu-id="406e6-117">attachments</span><span class="sxs-lookup"><span data-stu-id="406e6-117">attachments</span></span>](#attachments-arrayattachmentdetailsjavascriptapioutlook15officeattachmentdetails) | <span data-ttu-id="406e6-118">Element</span><span class="sxs-lookup"><span data-stu-id="406e6-118">Member</span></span> |
| [<span data-ttu-id="406e6-119">bcc</span><span class="sxs-lookup"><span data-stu-id="406e6-119">bcc</span></span>](#bcc-recipientsjavascriptapioutlook15officerecipients) | <span data-ttu-id="406e6-120">Element</span><span class="sxs-lookup"><span data-stu-id="406e6-120">Member</span></span> |
| [<span data-ttu-id="406e6-121">body</span><span class="sxs-lookup"><span data-stu-id="406e6-121">body</span></span>](#body-bodyjavascriptapioutlook15officebody) | <span data-ttu-id="406e6-122">Element</span><span class="sxs-lookup"><span data-stu-id="406e6-122">Member</span></span> |
| [<span data-ttu-id="406e6-123">cc</span><span class="sxs-lookup"><span data-stu-id="406e6-123">cc</span></span>](#cc-arrayemailaddressdetailsjavascriptapioutlook15officeemailaddressdetailsrecipientsjavascriptapioutlook15officerecipients) | <span data-ttu-id="406e6-124">Element</span><span class="sxs-lookup"><span data-stu-id="406e6-124">Member</span></span> |
| [<span data-ttu-id="406e6-125">conversationId</span><span class="sxs-lookup"><span data-stu-id="406e6-125">conversationId</span></span>](#nullable-conversationid-string) | <span data-ttu-id="406e6-126">Element</span><span class="sxs-lookup"><span data-stu-id="406e6-126">Member</span></span> |
| [<span data-ttu-id="406e6-127">dateTimeCreated</span><span class="sxs-lookup"><span data-stu-id="406e6-127">dateTimeCreated</span></span>](#datetimecreated-date) | <span data-ttu-id="406e6-128">Element</span><span class="sxs-lookup"><span data-stu-id="406e6-128">Member</span></span> |
| [<span data-ttu-id="406e6-129">dateTimeModified</span><span class="sxs-lookup"><span data-stu-id="406e6-129">dateTimeModified</span></span>](#datetimemodified-date) | <span data-ttu-id="406e6-130">Element</span><span class="sxs-lookup"><span data-stu-id="406e6-130">Member</span></span> |
| [<span data-ttu-id="406e6-131">end</span><span class="sxs-lookup"><span data-stu-id="406e6-131">end</span></span>](#end-datetimejavascriptapioutlook15officetime) | <span data-ttu-id="406e6-132">Element</span><span class="sxs-lookup"><span data-stu-id="406e6-132">Member</span></span> |
| [<span data-ttu-id="406e6-133">from</span><span class="sxs-lookup"><span data-stu-id="406e6-133">from</span></span>](#from-emailaddressdetailsjavascriptapioutlook15officeemailaddressdetails) | <span data-ttu-id="406e6-134">Element</span><span class="sxs-lookup"><span data-stu-id="406e6-134">Member</span></span> |
| [<span data-ttu-id="406e6-135">internetMessageId</span><span class="sxs-lookup"><span data-stu-id="406e6-135">internetMessageId</span></span>](#internetmessageid-string) | <span data-ttu-id="406e6-136">Element</span><span class="sxs-lookup"><span data-stu-id="406e6-136">Member</span></span> |
| [<span data-ttu-id="406e6-137">itemClass</span><span class="sxs-lookup"><span data-stu-id="406e6-137">itemClass</span></span>](#itemclass-string) | <span data-ttu-id="406e6-138">Element</span><span class="sxs-lookup"><span data-stu-id="406e6-138">Member</span></span> |
| [<span data-ttu-id="406e6-139">itemId</span><span class="sxs-lookup"><span data-stu-id="406e6-139">itemId</span></span>](#nullable-itemid-string) | <span data-ttu-id="406e6-140">Element</span><span class="sxs-lookup"><span data-stu-id="406e6-140">Member</span></span> |
| [<span data-ttu-id="406e6-141">itemType</span><span class="sxs-lookup"><span data-stu-id="406e6-141">itemType</span></span>](#itemtype-officemailboxenumsitemtypejavascriptapioutlook15officemailboxenumsitemtype) | <span data-ttu-id="406e6-142">Element</span><span class="sxs-lookup"><span data-stu-id="406e6-142">Member</span></span> |
| [<span data-ttu-id="406e6-143">location</span><span class="sxs-lookup"><span data-stu-id="406e6-143">location</span></span>](#location-stringlocationjavascriptapioutlook15officelocation) | <span data-ttu-id="406e6-144">Element</span><span class="sxs-lookup"><span data-stu-id="406e6-144">Member</span></span> |
| [<span data-ttu-id="406e6-145">normalizedSubject</span><span class="sxs-lookup"><span data-stu-id="406e6-145">normalizedSubject</span></span>](#normalizedsubject-string) | <span data-ttu-id="406e6-146">Element</span><span class="sxs-lookup"><span data-stu-id="406e6-146">Member</span></span> |
| [<span data-ttu-id="406e6-147">notificationMessages</span><span class="sxs-lookup"><span data-stu-id="406e6-147">notificationMessages</span></span>](#notificationmessages-notificationmessagesjavascriptapioutlook15officenotificationmessages) | <span data-ttu-id="406e6-148">Element</span><span class="sxs-lookup"><span data-stu-id="406e6-148">Member</span></span> |
| [<span data-ttu-id="406e6-149">optionalAttendees</span><span class="sxs-lookup"><span data-stu-id="406e6-149">optionalAttendees</span></span>](#optionalattendees-arrayemailaddressdetailsjavascriptapioutlook15officeemailaddressdetailsrecipientsjavascriptapioutlook15officerecipients) | <span data-ttu-id="406e6-150">Element</span><span class="sxs-lookup"><span data-stu-id="406e6-150">Member</span></span> |
| [<span data-ttu-id="406e6-151">organizer</span><span class="sxs-lookup"><span data-stu-id="406e6-151">organizer</span></span>](#organizer-emailaddressdetailsjavascriptapioutlook15officeemailaddressdetails) | <span data-ttu-id="406e6-152">Element</span><span class="sxs-lookup"><span data-stu-id="406e6-152">Member</span></span> |
| [<span data-ttu-id="406e6-153">requiredAttendees</span><span class="sxs-lookup"><span data-stu-id="406e6-153">requiredAttendees</span></span>](#requiredattendees-arrayemailaddressdetailsjavascriptapioutlook15officeemailaddressdetailsrecipientsjavascriptapioutlook15officerecipients) | <span data-ttu-id="406e6-154">Element</span><span class="sxs-lookup"><span data-stu-id="406e6-154">Member</span></span> |
| [<span data-ttu-id="406e6-155">sender</span><span class="sxs-lookup"><span data-stu-id="406e6-155">sender</span></span>](#sender-emailaddressdetailsjavascriptapioutlook15officeemailaddressdetails) | <span data-ttu-id="406e6-156">Element</span><span class="sxs-lookup"><span data-stu-id="406e6-156">Member</span></span> |
| [<span data-ttu-id="406e6-157">start</span><span class="sxs-lookup"><span data-stu-id="406e6-157">start</span></span>](#start-datetimejavascriptapioutlook15officetime) | <span data-ttu-id="406e6-158">Element</span><span class="sxs-lookup"><span data-stu-id="406e6-158">Member</span></span> |
| [<span data-ttu-id="406e6-159">subject</span><span class="sxs-lookup"><span data-stu-id="406e6-159">subject</span></span>](#subject-stringsubjectjavascriptapioutlook15officesubject) | <span data-ttu-id="406e6-160">Element</span><span class="sxs-lookup"><span data-stu-id="406e6-160">Member</span></span> |
| [<span data-ttu-id="406e6-161">to</span><span class="sxs-lookup"><span data-stu-id="406e6-161">to</span></span>](#to-arrayemailaddressdetailsjavascriptapioutlook15officeemailaddressdetailsrecipientsjavascriptapioutlook15officerecipients) | <span data-ttu-id="406e6-162">Element</span><span class="sxs-lookup"><span data-stu-id="406e6-162">Member</span></span> |
| [<span data-ttu-id="406e6-163">addFileAttachmentAsync</span><span class="sxs-lookup"><span data-stu-id="406e6-163">addFileAttachmentAsync</span></span>](#addfileattachmentasyncuri-attachmentname-options-callback) | <span data-ttu-id="406e6-164">Methode</span><span class="sxs-lookup"><span data-stu-id="406e6-164">Method</span></span> |
| [<span data-ttu-id="406e6-165">addItemAttachmentAsync</span><span class="sxs-lookup"><span data-stu-id="406e6-165">addItemAttachmentAsync</span></span>](#additemattachmentasyncitemid-attachmentname-options-callback) | <span data-ttu-id="406e6-166">Methode</span><span class="sxs-lookup"><span data-stu-id="406e6-166">Method</span></span> |
| [<span data-ttu-id="406e6-167">close</span><span class="sxs-lookup"><span data-stu-id="406e6-167">close</span></span>](#close) | <span data-ttu-id="406e6-168">Methode</span><span class="sxs-lookup"><span data-stu-id="406e6-168">Method</span></span> |
| [<span data-ttu-id="406e6-169">displayReplyAllForm</span><span class="sxs-lookup"><span data-stu-id="406e6-169">displayReplyAllForm</span></span>](#displayreplyallformformdata) | <span data-ttu-id="406e6-170">Methode</span><span class="sxs-lookup"><span data-stu-id="406e6-170">Method</span></span> |
| [<span data-ttu-id="406e6-171">displayReplyForm</span><span class="sxs-lookup"><span data-stu-id="406e6-171">displayReplyForm</span></span>](#displayreplyformformdata) | <span data-ttu-id="406e6-172">Methode</span><span class="sxs-lookup"><span data-stu-id="406e6-172">Method</span></span> |
| [<span data-ttu-id="406e6-173">getEntities</span><span class="sxs-lookup"><span data-stu-id="406e6-173">getEntities</span></span>](#getentities--entitiesjavascriptapioutlook15officeentities) | <span data-ttu-id="406e6-174">Methode</span><span class="sxs-lookup"><span data-stu-id="406e6-174">Method</span></span> |
| [<span data-ttu-id="406e6-175">getEntitiesByType</span><span class="sxs-lookup"><span data-stu-id="406e6-175">getEntitiesByType</span></span>](#getentitiesbytypeentitytype--nullable-arraystringcontactjavascriptapioutlook15officecontactmeetingsuggestionjavascriptapioutlook15officemeetingsuggestionphonenumberjavascriptapioutlook15officephonenumbertasksuggestionjavascriptapioutlook15officetasksuggestion) | <span data-ttu-id="406e6-176">Methode</span><span class="sxs-lookup"><span data-stu-id="406e6-176">Method</span></span> |
| [<span data-ttu-id="406e6-177">getFilteredEntitiesByName</span><span class="sxs-lookup"><span data-stu-id="406e6-177">getFilteredEntitiesByName</span></span>](#getfilteredentitiesbynamename--nullable-arraystringcontactjavascriptapioutlook15officecontactmeetingsuggestionjavascriptapioutlook15officemeetingsuggestionphonenumberjavascriptapioutlook15officephonenumbertasksuggestionjavascriptapioutlook15officetasksuggestion) | <span data-ttu-id="406e6-178">Methode</span><span class="sxs-lookup"><span data-stu-id="406e6-178">Method</span></span> |
| [<span data-ttu-id="406e6-179">getRegExMatches</span><span class="sxs-lookup"><span data-stu-id="406e6-179">getRegExMatches</span></span>](#getregexmatches--object) | <span data-ttu-id="406e6-180">Methode</span><span class="sxs-lookup"><span data-stu-id="406e6-180">Method</span></span> |
| [<span data-ttu-id="406e6-181">getRegExMatchesByName</span><span class="sxs-lookup"><span data-stu-id="406e6-181">getRegExMatchesByName</span></span>](#getregexmatchesbynamename--nullable-array-string-) | <span data-ttu-id="406e6-182">Methode</span><span class="sxs-lookup"><span data-stu-id="406e6-182">Method</span></span> |
| [<span data-ttu-id="406e6-183">getSelectedDataAsync</span><span class="sxs-lookup"><span data-stu-id="406e6-183">getSelectedDataAsync</span></span>](#getselecteddataasynccoerciontype-options-callback--string) | <span data-ttu-id="406e6-184">Methode</span><span class="sxs-lookup"><span data-stu-id="406e6-184">Method</span></span> |
| [<span data-ttu-id="406e6-185">loadCustomPropertiesAsync</span><span class="sxs-lookup"><span data-stu-id="406e6-185">loadCustomPropertiesAsync</span></span>](#loadcustompropertiesasynccallback-usercontext) | <span data-ttu-id="406e6-186">Methode</span><span class="sxs-lookup"><span data-stu-id="406e6-186">Method</span></span> |
| [<span data-ttu-id="406e6-187">removeAttachmentAsync</span><span class="sxs-lookup"><span data-stu-id="406e6-187">removeAttachmentAsync</span></span>](#removeattachmentasyncattachmentid-options-callback) | <span data-ttu-id="406e6-188">Methode</span><span class="sxs-lookup"><span data-stu-id="406e6-188">Method</span></span> |
| [<span data-ttu-id="406e6-189">saveAsync</span><span class="sxs-lookup"><span data-stu-id="406e6-189">saveAsync</span></span>](#saveasyncoptions-callback) | <span data-ttu-id="406e6-190">Methode</span><span class="sxs-lookup"><span data-stu-id="406e6-190">Method</span></span> |
| [<span data-ttu-id="406e6-191">setSelectedDataAsync</span><span class="sxs-lookup"><span data-stu-id="406e6-191">setSelectedDataAsync</span></span>](#setselecteddataasyncdata-options-callback) | <span data-ttu-id="406e6-192">Methode</span><span class="sxs-lookup"><span data-stu-id="406e6-192">Method</span></span> |

### <a name="example"></a><span data-ttu-id="406e6-193">Beispiel</span><span class="sxs-lookup"><span data-stu-id="406e6-193">Example</span></span>

<span data-ttu-id="406e6-194">Im folgenden JavaScript-Codebeispiel wird der Zugriff auf die `subject`-Eigenschaft des aktuellen Elements in Outlook veranschaulicht.</span><span class="sxs-lookup"><span data-stu-id="406e6-194">The following JavaScript code example shows how to access the `subject` property of the current item in Outlook.</span></span>

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

### <a name="members"></a><span data-ttu-id="406e6-195">Elemente</span><span class="sxs-lookup"><span data-stu-id="406e6-195">Members</span></span>

#### <a name="attachments-arrayattachmentdetailsjavascriptapioutlook15officeattachmentdetails"></a><span data-ttu-id="406e6-196">attachments :Array.<[AttachmentDetails](/javascript/api/outlook_1_5/office.attachmentdetails)></span><span class="sxs-lookup"><span data-stu-id="406e6-196">attachments :Array.<[AttachmentDetails](/javascript/api/outlook_1_5/office.attachmentdetails)></span></span>

<span data-ttu-id="406e6-p102">Ruft ein Array mit Anlagen für das Element ab. Nur Lesemodus.</span><span class="sxs-lookup"><span data-stu-id="406e6-p102">Gets an array of attachments for the item. Read mode only.</span></span>

> [!NOTE]
> <span data-ttu-id="406e6-199">Bestimmte Dateitypen werden aufgrund von Sicherheitsproblemen von Outlook blockiert und daher nicht zurückgegeben werden.</span><span class="sxs-lookup"><span data-stu-id="406e6-199">Certain types of files are blocked by Outlook due to potential security issues and are therefore not returned.</span></span> <span data-ttu-id="406e6-200">Weitere Informationen finden Sie unter [Blockierte Anlagen in Outlook](https://support.office.com/article/Blocked-attachments-in-Outlook-434752E1-02D3-4E90-9124-8B81E49A8519).</span><span class="sxs-lookup"><span data-stu-id="406e6-200">For more information, see [Blocked attachments in Outlook](https://support.office.com/article/Blocked-attachments-in-Outlook-434752E1-02D3-4E90-9124-8B81E49A8519).</span></span>

##### <a name="type"></a><span data-ttu-id="406e6-201">Typ:</span><span class="sxs-lookup"><span data-stu-id="406e6-201">Type:</span></span>

*   <span data-ttu-id="406e6-202">Array.<[AttachmentDetails](/javascript/api/outlook_1_5/office.attachmentdetails)></span><span class="sxs-lookup"><span data-stu-id="406e6-202">Array.<[AttachmentDetails](/javascript/api/outlook_1_5/office.attachmentdetails)></span></span>

##### <a name="requirements"></a><span data-ttu-id="406e6-203">Anforderungen</span><span class="sxs-lookup"><span data-stu-id="406e6-203">Requirements</span></span>

|<span data-ttu-id="406e6-204">Anforderung</span><span class="sxs-lookup"><span data-stu-id="406e6-204">Requirement</span></span>| <span data-ttu-id="406e6-205">Wert</span><span class="sxs-lookup"><span data-stu-id="406e6-205">Value</span></span>|
|---|---|
|[<span data-ttu-id="406e6-206">Mindestversion des Postfachanforderungssatzes</span><span class="sxs-lookup"><span data-stu-id="406e6-206">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="406e6-207">1.0</span><span class="sxs-lookup"><span data-stu-id="406e6-207">1.0</span></span>|
|[<span data-ttu-id="406e6-208">Mindestberechtigungsstufe</span><span class="sxs-lookup"><span data-stu-id="406e6-208">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="406e6-209">ReadItem</span><span class="sxs-lookup"><span data-stu-id="406e6-209">ReadItem</span></span>|
|[<span data-ttu-id="406e6-210">Zutreffender Outlook-Modus</span><span class="sxs-lookup"><span data-stu-id="406e6-210">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="406e6-211">Lesen</span><span class="sxs-lookup"><span data-stu-id="406e6-211">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="406e6-212">Beispiel</span><span class="sxs-lookup"><span data-stu-id="406e6-212">Example</span></span>

<span data-ttu-id="406e6-213">Mit dem folgende Code wird eine HTML-Zeichenfolge mit Details aller Anlagen für das aktuelle Element erstellt.</span><span class="sxs-lookup"><span data-stu-id="406e6-213">The following code builds an HTML string with details of all attachments on the current item.</span></span>

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

####  <a name="bcc-recipientsjavascriptapioutlook15officerecipients"></a><span data-ttu-id="406e6-214">bcc :[Recipients](/javascript/api/outlook_1_5/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="406e6-214">bcc :[Recipients](/javascript/api/outlook_1_5/office.recipients)</span></span>

<span data-ttu-id="406e6-215">Ruft ein Objekt, das Methoden zum Abrufen oder aktualisieren die Empfänger in der Zeile Bcc (blind Carbon Copy, Blindkopie) einer Nachricht bereitstellt.</span><span class="sxs-lookup"><span data-stu-id="406e6-215">Gets an object that provides methods to get or update the recipients on the Bcc (blind carbon copy) line of a message.</span></span> <span data-ttu-id="406e6-216">Nur Verfassenmodus.</span><span class="sxs-lookup"><span data-stu-id="406e6-216">Compose mode only.</span></span>

##### <a name="type"></a><span data-ttu-id="406e6-217">Typ:</span><span class="sxs-lookup"><span data-stu-id="406e6-217">Type:</span></span>

*   [<span data-ttu-id="406e6-218">Recipients</span><span class="sxs-lookup"><span data-stu-id="406e6-218">Recipients</span></span>](/javascript/api/outlook_1_5/office.recipients)

##### <a name="requirements"></a><span data-ttu-id="406e6-219">Anforderungen</span><span class="sxs-lookup"><span data-stu-id="406e6-219">Requirements</span></span>

|<span data-ttu-id="406e6-220">Anforderung</span><span class="sxs-lookup"><span data-stu-id="406e6-220">Requirement</span></span>| <span data-ttu-id="406e6-221">Wert</span><span class="sxs-lookup"><span data-stu-id="406e6-221">Value</span></span>|
|---|---|
|[<span data-ttu-id="406e6-222">Mindestversion des Postfachanforderungssatzes</span><span class="sxs-lookup"><span data-stu-id="406e6-222">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="406e6-223">1.1</span><span class="sxs-lookup"><span data-stu-id="406e6-223">1.1</span></span>|
|[<span data-ttu-id="406e6-224">Mindestberechtigungsstufe</span><span class="sxs-lookup"><span data-stu-id="406e6-224">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="406e6-225">ReadItem</span><span class="sxs-lookup"><span data-stu-id="406e6-225">ReadItem</span></span>|
|[<span data-ttu-id="406e6-226">Zutreffender Outlook-Modus</span><span class="sxs-lookup"><span data-stu-id="406e6-226">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="406e6-227">Verfassen</span><span class="sxs-lookup"><span data-stu-id="406e6-227">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="406e6-228">Beispiel</span><span class="sxs-lookup"><span data-stu-id="406e6-228">Example</span></span>

```
Office.context.mailbox.item.bcc.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.bcc.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.bcc.getAsync(callback);

function callback(asyncResult) {
  var arrayOfBccRecipients = asyncResult.value;
}
```

####  <a name="body-bodyjavascriptapioutlook15officebody"></a><span data-ttu-id="406e6-229">body :[Body](/javascript/api/outlook_1_5/office.body)</span><span class="sxs-lookup"><span data-stu-id="406e6-229">body :[Body](/javascript/api/outlook_1_5/office.body)</span></span>

<span data-ttu-id="406e6-230">Ruft ein Objekt ab, das Methoden zum Bearbeiten des Textkörpers eines Elements bereitstellt.</span><span class="sxs-lookup"><span data-stu-id="406e6-230">Gets an object that provides methods for manipulating the body of an item.</span></span>

##### <a name="type"></a><span data-ttu-id="406e6-231">Typ:</span><span class="sxs-lookup"><span data-stu-id="406e6-231">Type:</span></span>

*   [<span data-ttu-id="406e6-232">Body</span><span class="sxs-lookup"><span data-stu-id="406e6-232">Body</span></span>](/javascript/api/outlook_1_5/office.body)

##### <a name="requirements"></a><span data-ttu-id="406e6-233">Anforderungen</span><span class="sxs-lookup"><span data-stu-id="406e6-233">Requirements</span></span>

|<span data-ttu-id="406e6-234">Anforderung</span><span class="sxs-lookup"><span data-stu-id="406e6-234">Requirement</span></span>| <span data-ttu-id="406e6-235">Wert</span><span class="sxs-lookup"><span data-stu-id="406e6-235">Value</span></span>|
|---|---|
|[<span data-ttu-id="406e6-236">Mindestversion des Postfachanforderungssatzes</span><span class="sxs-lookup"><span data-stu-id="406e6-236">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="406e6-237">1.1</span><span class="sxs-lookup"><span data-stu-id="406e6-237">1.1</span></span>|
|[<span data-ttu-id="406e6-238">Mindestberechtigungsstufe</span><span class="sxs-lookup"><span data-stu-id="406e6-238">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="406e6-239">ReadItem</span><span class="sxs-lookup"><span data-stu-id="406e6-239">ReadItem</span></span>|
|[<span data-ttu-id="406e6-240">Zutreffender Outlook-Modus</span><span class="sxs-lookup"><span data-stu-id="406e6-240">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="406e6-241">Verfassen oder Lesen</span><span class="sxs-lookup"><span data-stu-id="406e6-241">Compose or read</span></span>|

####  <a name="cc-arrayemailaddressdetailsjavascriptapioutlook15officeemailaddressdetailsrecipientsjavascriptapioutlook15officerecipients"></a><span data-ttu-id="406e6-242">cc: Array. <[EmailAddressDetails](/javascript/api/outlook_1_5/office.emailaddressdetails)>|[Empfänger](/javascript/api/outlook_1_5/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="406e6-242">cc :Array.<[EmailAddressDetails](/javascript/api/outlook_1_5/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook_1_5/office.recipients)</span></span>

<span data-ttu-id="406e6-243">Ermöglicht den Zugriff auf die (Carbon Copy, Blindkopie) Cc-Empfänger einer Nachricht.</span><span class="sxs-lookup"><span data-stu-id="406e6-243">Provides access to the Cc (carbon copy) recipients of a message.</span></span> <span data-ttu-id="406e6-244">Der Typ des Objekts und die Zugriffsebene, hängt von den Modus des aktuellen Elements ab.</span><span class="sxs-lookup"><span data-stu-id="406e6-244">The type of object and level of access depends on the mode of the current item.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="406e6-245">Lesemodus</span><span class="sxs-lookup"><span data-stu-id="406e6-245">Read mode</span></span>

<span data-ttu-id="406e6-p106">Die `cc`-Eigenschaft gibt ein Array mit einem `EmailAddressDetails`-Objekt für jeden Empfänger in der **Cc**-Zeile der Nachricht zurück. Die Auflistung ist auf höchstens 100 Elemente beschränkt.</span><span class="sxs-lookup"><span data-stu-id="406e6-p106">The `cc` property returns an array that contains an `EmailAddressDetails` object for each recipient listed on the **Cc** line of the message. The collection is limited to a maximum of 100 members.</span></span>

##### <a name="compose-mode"></a><span data-ttu-id="406e6-248">Verfassenmodus</span><span class="sxs-lookup"><span data-stu-id="406e6-248">Compose mode</span></span>

<span data-ttu-id="406e6-249">Die `cc` -Eigenschaft gibt eine `Recipients` -Objekt, das Methoden zum Abrufen oder aktualisieren die Empfänger in der Zeile **Cc** der Nachricht bereitstellt.</span><span class="sxs-lookup"><span data-stu-id="406e6-249">The `cc` property returns a `Recipients` object that provides methods to get or update the recipients on the **Cc** line of the message.</span></span>

##### <a name="type"></a><span data-ttu-id="406e6-250">Typ:</span><span class="sxs-lookup"><span data-stu-id="406e6-250">Type:</span></span>

*   <span data-ttu-id="406e6-251">Array.<[EmailAddressDetails](/javascript/api/outlook_1_5/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook_1_5/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="406e6-251">Array.<[EmailAddressDetails](/javascript/api/outlook_1_5/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook_1_5/office.recipients)</span></span>

##### <a name="requirements"></a><span data-ttu-id="406e6-252">Anforderungen</span><span class="sxs-lookup"><span data-stu-id="406e6-252">Requirements</span></span>

|<span data-ttu-id="406e6-253">Anforderung</span><span class="sxs-lookup"><span data-stu-id="406e6-253">Requirement</span></span>| <span data-ttu-id="406e6-254">Wert</span><span class="sxs-lookup"><span data-stu-id="406e6-254">Value</span></span>|
|---|---|
|[<span data-ttu-id="406e6-255">Mindestversion des Postfachanforderungssatzes</span><span class="sxs-lookup"><span data-stu-id="406e6-255">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="406e6-256">1.0</span><span class="sxs-lookup"><span data-stu-id="406e6-256">1.0</span></span>|
|[<span data-ttu-id="406e6-257">Mindestberechtigungsstufe</span><span class="sxs-lookup"><span data-stu-id="406e6-257">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="406e6-258">ReadItem</span><span class="sxs-lookup"><span data-stu-id="406e6-258">ReadItem</span></span>|
|[<span data-ttu-id="406e6-259">Zutreffender Outlook-Modus</span><span class="sxs-lookup"><span data-stu-id="406e6-259">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="406e6-260">Verfassen oder Lesen</span><span class="sxs-lookup"><span data-stu-id="406e6-260">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="406e6-261">Beispiel</span><span class="sxs-lookup"><span data-stu-id="406e6-261">Example</span></span>

```
Office.context.mailbox.item.cc.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.cc.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.cc.getAsync(callback);

function callback(asyncResult) {
  var arrayOfCcRecipients = asyncResult.value;
}
```

####  <a name="nullable-conversationid-string"></a><span data-ttu-id="406e6-262">(nullable) conversationId :String</span><span class="sxs-lookup"><span data-stu-id="406e6-262">(nullable) conversationId :String</span></span>

<span data-ttu-id="406e6-263">Ruft einen Bezeichner für die E-Mail-Unterhaltung ab, in der eine bestimmte Nachricht enthalten ist.</span><span class="sxs-lookup"><span data-stu-id="406e6-263">Gets an identifier for the email conversation that contains a particular message.</span></span>

<span data-ttu-id="406e6-p107">Sie können für diese Eigenschaft eine ganze Zahl abrufen, wenn Ihre Mail-App in Formularen zum Lesen oder Antworten in Formularen zum Verfassen aktiviert wird. Wenn der Benutzer den Betreff der Antwortnachricht ändert, ändert sich beim Versenden die Konversations-ID für die entsprechende Nachricht, und der Wert, den Sie vorher bezogen haben, trifft nicht länger zu.</span><span class="sxs-lookup"><span data-stu-id="406e6-p107">You can get an integer for this property if your mail app is activated in read forms or responses in compose forms. If subsequently the user changes the subject of the reply message, upon sending the reply, the conversation ID for that message will change and that value you obtained earlier will no longer apply.</span></span>

<span data-ttu-id="406e6-p108">Sie erhalten in einem Formular zum Verfassen für diese Eigenschaft für ein neues Element null. Wenn der Benutzer einen Betreff festlegt und das Element speichert, gibt die `conversationId`-Eigenschaft einen Wert zurück.</span><span class="sxs-lookup"><span data-stu-id="406e6-p108">You get null for this property for a new item in a compose form. If the user sets a subject and saves the item, the `conversationId` property will return a value.</span></span>

##### <a name="type"></a><span data-ttu-id="406e6-268">Typ:</span><span class="sxs-lookup"><span data-stu-id="406e6-268">Type:</span></span>

*   <span data-ttu-id="406e6-269">String</span><span class="sxs-lookup"><span data-stu-id="406e6-269">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="406e6-270">Anforderungen</span><span class="sxs-lookup"><span data-stu-id="406e6-270">Requirements</span></span>

|<span data-ttu-id="406e6-271">Anforderung</span><span class="sxs-lookup"><span data-stu-id="406e6-271">Requirement</span></span>| <span data-ttu-id="406e6-272">Wert</span><span class="sxs-lookup"><span data-stu-id="406e6-272">Value</span></span>|
|---|---|
|[<span data-ttu-id="406e6-273">Mindestversion des Postfachanforderungssatzes</span><span class="sxs-lookup"><span data-stu-id="406e6-273">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="406e6-274">1.0</span><span class="sxs-lookup"><span data-stu-id="406e6-274">1.0</span></span>|
|[<span data-ttu-id="406e6-275">Mindestberechtigungsstufe</span><span class="sxs-lookup"><span data-stu-id="406e6-275">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="406e6-276">ReadItem</span><span class="sxs-lookup"><span data-stu-id="406e6-276">ReadItem</span></span>|
|[<span data-ttu-id="406e6-277">Zutreffender Outlook-Modus</span><span class="sxs-lookup"><span data-stu-id="406e6-277">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="406e6-278">Verfassen oder Lesen</span><span class="sxs-lookup"><span data-stu-id="406e6-278">Compose or read</span></span>|

#### <a name="datetimecreated-date"></a><span data-ttu-id="406e6-279">dateTimeCreated :Date</span><span class="sxs-lookup"><span data-stu-id="406e6-279">dateTimeCreated :Date</span></span>

<span data-ttu-id="406e6-p109">Ruft das Datum und die Uhrzeit der Erstellung eines Elements ab. Nur Lesemodus.</span><span class="sxs-lookup"><span data-stu-id="406e6-p109">Gets the date and time that an item was created. Read mode only.</span></span>

##### <a name="type"></a><span data-ttu-id="406e6-282">Typ:</span><span class="sxs-lookup"><span data-stu-id="406e6-282">Type:</span></span>

*   <span data-ttu-id="406e6-283">Datum</span><span class="sxs-lookup"><span data-stu-id="406e6-283">Date</span></span>

##### <a name="requirements"></a><span data-ttu-id="406e6-284">Anforderungen</span><span class="sxs-lookup"><span data-stu-id="406e6-284">Requirements</span></span>

|<span data-ttu-id="406e6-285">Anforderung</span><span class="sxs-lookup"><span data-stu-id="406e6-285">Requirement</span></span>| <span data-ttu-id="406e6-286">Wert</span><span class="sxs-lookup"><span data-stu-id="406e6-286">Value</span></span>|
|---|---|
|[<span data-ttu-id="406e6-287">Mindestversion des Postfachanforderungssatzes</span><span class="sxs-lookup"><span data-stu-id="406e6-287">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="406e6-288">1.0</span><span class="sxs-lookup"><span data-stu-id="406e6-288">1.0</span></span>|
|[<span data-ttu-id="406e6-289">Mindestberechtigungsstufe</span><span class="sxs-lookup"><span data-stu-id="406e6-289">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="406e6-290">ReadItem</span><span class="sxs-lookup"><span data-stu-id="406e6-290">ReadItem</span></span>|
|[<span data-ttu-id="406e6-291">Zutreffender Outlook-Modus</span><span class="sxs-lookup"><span data-stu-id="406e6-291">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="406e6-292">Lesen</span><span class="sxs-lookup"><span data-stu-id="406e6-292">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="406e6-293">Beispiel</span><span class="sxs-lookup"><span data-stu-id="406e6-293">Example</span></span>

```
var created = Office.context.mailbox.item.dateTimeCreated;
```

#### <a name="datetimemodified-date"></a><span data-ttu-id="406e6-294">dateTimeModified :Date</span><span class="sxs-lookup"><span data-stu-id="406e6-294">dateTimeModified :Date</span></span>

<span data-ttu-id="406e6-p110">Ruft das Datum und die Uhrzeit der letzten Änderung eines Elements ab. Nur Lesemodus.</span><span class="sxs-lookup"><span data-stu-id="406e6-p110">Gets the date and time that an item was last modified. Read mode only.</span></span>

> [!NOTE]
> <span data-ttu-id="406e6-297">Dieser Member wird in Outlook für iOS oder Outlook für Android nicht unterstützt.</span><span class="sxs-lookup"><span data-stu-id="406e6-297">This member is not supported in Outlook for iOS or Outlook for Android.</span></span>

##### <a name="type"></a><span data-ttu-id="406e6-298">Typ:</span><span class="sxs-lookup"><span data-stu-id="406e6-298">Type:</span></span>

*   <span data-ttu-id="406e6-299">Datum</span><span class="sxs-lookup"><span data-stu-id="406e6-299">Date</span></span>

##### <a name="requirements"></a><span data-ttu-id="406e6-300">Anforderungen</span><span class="sxs-lookup"><span data-stu-id="406e6-300">Requirements</span></span>

|<span data-ttu-id="406e6-301">Anforderung</span><span class="sxs-lookup"><span data-stu-id="406e6-301">Requirement</span></span>| <span data-ttu-id="406e6-302">Wert</span><span class="sxs-lookup"><span data-stu-id="406e6-302">Value</span></span>|
|---|---|
|[<span data-ttu-id="406e6-303">Mindestversion des Postfachanforderungssatzes</span><span class="sxs-lookup"><span data-stu-id="406e6-303">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="406e6-304">1.0</span><span class="sxs-lookup"><span data-stu-id="406e6-304">1.0</span></span>|
|[<span data-ttu-id="406e6-305">Mindestberechtigungsstufe</span><span class="sxs-lookup"><span data-stu-id="406e6-305">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="406e6-306">ReadItem</span><span class="sxs-lookup"><span data-stu-id="406e6-306">ReadItem</span></span>|
|[<span data-ttu-id="406e6-307">Zutreffender Outlook-Modus</span><span class="sxs-lookup"><span data-stu-id="406e6-307">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="406e6-308">Lesen</span><span class="sxs-lookup"><span data-stu-id="406e6-308">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="406e6-309">Beispiel</span><span class="sxs-lookup"><span data-stu-id="406e6-309">Example</span></span>

```
var modified = Office.context.mailbox.item.dateTimeModified;
```

####  <a name="end-datetimejavascriptapioutlook15officetime"></a><span data-ttu-id="406e6-310">end :Date|[Time](/javascript/api/outlook_1_5/office.time)</span><span class="sxs-lookup"><span data-stu-id="406e6-310">end :Date|[Time](/javascript/api/outlook_1_5/office.time)</span></span>

<span data-ttu-id="406e6-311">Ruft Datum und Zeit für das Ende des Termins ab oder legt diese fest.</span><span class="sxs-lookup"><span data-stu-id="406e6-311">Gets or sets the date and time that the appointment is to end.</span></span>

<span data-ttu-id="406e6-p111">Die `end`-Eigenschaft wird als Datums- und Uhrzeitwert in UTC ausgedrückt. Sie können die [`convertToLocalClientTime`](office.context.mailbox.md#converttolocalclienttimetimevalue--localclienttimejavascriptapioutlook15officelocalclienttime)-Methode verwenden, um den Endeigenschaftswert in den lokalen Uhrzeit- und Datumswert des Clients umzuwandeln.</span><span class="sxs-lookup"><span data-stu-id="406e6-p111">The `end` property is expressed as a Coordinated Universal Time (UTC) date and time value. You can use the [`convertToLocalClientTime`](office.context.mailbox.md#converttolocalclienttimetimevalue--localclienttimejavascriptapioutlook15officelocalclienttime) method to convert the end property value to the client’s local date and time.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="406e6-314">Lesemodus</span><span class="sxs-lookup"><span data-stu-id="406e6-314">Read mode</span></span>

<span data-ttu-id="406e6-315">Die `end`-Eigenschaft gibt ein `Date`-Objekt zurück.</span><span class="sxs-lookup"><span data-stu-id="406e6-315">The `end` property returns a `Date` object.</span></span>

##### <a name="compose-mode"></a><span data-ttu-id="406e6-316">Verfassenmodus</span><span class="sxs-lookup"><span data-stu-id="406e6-316">Compose mode</span></span>

<span data-ttu-id="406e6-317">Die `end`-Eigenschaft gibt ein `Time`-Objekt zurück.</span><span class="sxs-lookup"><span data-stu-id="406e6-317">The `end` property returns a `Time` object.</span></span>

<span data-ttu-id="406e6-318">Wenn Sie die [`Time.setAsync`](/javascript/api/outlook_1_5/office.time#setasync-datetime--options--callback-)-Methode verwenden, um die Endzeit im Verfassenmodus festzulegen, sollten Sie die [`convertToUtcClientTime`](office.context.mailbox.md#converttoutcclienttimeinput--date)-Methode verwenden, um die Ortszeit auf dem Client für den Server in UTC umzuwandeln.</span><span class="sxs-lookup"><span data-stu-id="406e6-318">When you use the [`Time.setAsync`](/javascript/api/outlook_1_5/office.time#setasync-datetime--options--callback-) method to set the end time, you should use the [`convertToUtcClientTime`](office.context.mailbox.md#converttoutcclienttimeinput--date) method to convert the local time on the client to UTC for the server.</span></span>

##### <a name="type"></a><span data-ttu-id="406e6-319">Typ:</span><span class="sxs-lookup"><span data-stu-id="406e6-319">Type:</span></span>

*   <span data-ttu-id="406e6-320">Date | [Time](/javascript/api/outlook_1_5/office.time)</span><span class="sxs-lookup"><span data-stu-id="406e6-320">Date | [Time](/javascript/api/outlook_1_5/office.time)</span></span>

##### <a name="requirements"></a><span data-ttu-id="406e6-321">Anforderungen</span><span class="sxs-lookup"><span data-stu-id="406e6-321">Requirements</span></span>

|<span data-ttu-id="406e6-322">Anforderung</span><span class="sxs-lookup"><span data-stu-id="406e6-322">Requirement</span></span>| <span data-ttu-id="406e6-323">Wert</span><span class="sxs-lookup"><span data-stu-id="406e6-323">Value</span></span>|
|---|---|
|[<span data-ttu-id="406e6-324">Mindestversion des Postfachanforderungssatzes</span><span class="sxs-lookup"><span data-stu-id="406e6-324">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="406e6-325">1.0</span><span class="sxs-lookup"><span data-stu-id="406e6-325">1.0</span></span>|
|[<span data-ttu-id="406e6-326">Mindestberechtigungsstufe</span><span class="sxs-lookup"><span data-stu-id="406e6-326">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="406e6-327">ReadItem</span><span class="sxs-lookup"><span data-stu-id="406e6-327">ReadItem</span></span>|
|[<span data-ttu-id="406e6-328">Zutreffender Outlook-Modus</span><span class="sxs-lookup"><span data-stu-id="406e6-328">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="406e6-329">Verfassen oder Lesen</span><span class="sxs-lookup"><span data-stu-id="406e6-329">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="406e6-330">Beispiel</span><span class="sxs-lookup"><span data-stu-id="406e6-330">Example</span></span>

<span data-ttu-id="406e6-331">Im folgenden Beispiel wird die Endzeit eines Termins im Verfassenmodus mithilfe der [`setAsync`](/javascript/api/outlook_1_5/office.time#setasync-datetime--options--callback-)-Methode des `Time`-Objekts festgelegt.</span><span class="sxs-lookup"><span data-stu-id="406e6-331">The following example sets the end time of an appointment in compose mode by using the [`setAsync`](/javascript/api/outlook_1_5/office.time#setasync-datetime--options--callback-) method of the `Time` object.</span></span>

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

#### <a name="from-emailaddressdetailsjavascriptapioutlook15officeemailaddressdetails"></a><span data-ttu-id="406e6-332">from :[EmailAddressDetails](/javascript/api/outlook_1_5/office.emailaddressdetails)</span><span class="sxs-lookup"><span data-stu-id="406e6-332">from :[EmailAddressDetails](/javascript/api/outlook_1_5/office.emailaddressdetails)</span></span>

<span data-ttu-id="406e6-p112">Ruft die E-Mail-Adresse des Absenders einer Nachricht ab. Nur Lesemodus.</span><span class="sxs-lookup"><span data-stu-id="406e6-p112">Gets the email address of the sender of a message. Read mode only.</span></span>

<span data-ttu-id="406e6-p113">Die `from`- und [`sender`](#sender-emailaddressdetailsjavascriptapioutlook15officeemailaddressdetails)-Eigenschaften stellen dieselbe Person dar, außer die Nachricht wurde von einem Delegaten gesendet. In diesem Fall stellt die `from`-Eigenschaft die Stellvertretung dar, und die sender-Eigenschaft stellt den Delegaten dar.</span><span class="sxs-lookup"><span data-stu-id="406e6-p113">The `from` and [`sender`](#sender-emailaddressdetailsjavascriptapioutlook15officeemailaddressdetails) properties represent the same person unless the message is sent by a delegate. In that case, the `from` property represents the delegator, and the sender property represents the delegate.</span></span>

> [!NOTE]
> <span data-ttu-id="406e6-337">Die `recipientType` -Eigenschaft des der `EmailAddressDetails` -Objekts der `from` -Eigenschaft ist `undefined`.</span><span class="sxs-lookup"><span data-stu-id="406e6-337">The `recipientType` property of the `EmailAddressDetails` object in the `from` property is `undefined`.</span></span>

##### <a name="type"></a><span data-ttu-id="406e6-338">Typ:</span><span class="sxs-lookup"><span data-stu-id="406e6-338">Type:</span></span>

*   [<span data-ttu-id="406e6-339">EmailAddressDetails</span><span class="sxs-lookup"><span data-stu-id="406e6-339">EmailAddressDetails</span></span>](/javascript/api/outlook_1_5/office.emailaddressdetails)

##### <a name="requirements"></a><span data-ttu-id="406e6-340">Anforderungen</span><span class="sxs-lookup"><span data-stu-id="406e6-340">Requirements</span></span>

|<span data-ttu-id="406e6-341">Anforderung</span><span class="sxs-lookup"><span data-stu-id="406e6-341">Requirement</span></span>| <span data-ttu-id="406e6-342">Wert</span><span class="sxs-lookup"><span data-stu-id="406e6-342">Value</span></span>|
|---|---|
|[<span data-ttu-id="406e6-343">Mindestversion des Postfachanforderungssatzes</span><span class="sxs-lookup"><span data-stu-id="406e6-343">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="406e6-344">1.0</span><span class="sxs-lookup"><span data-stu-id="406e6-344">1.0</span></span>|
|[<span data-ttu-id="406e6-345">Mindestberechtigungsstufe</span><span class="sxs-lookup"><span data-stu-id="406e6-345">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="406e6-346">ReadItem</span><span class="sxs-lookup"><span data-stu-id="406e6-346">ReadItem</span></span>|
|[<span data-ttu-id="406e6-347">Zutreffender Outlook-Modus</span><span class="sxs-lookup"><span data-stu-id="406e6-347">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="406e6-348">Lesen</span><span class="sxs-lookup"><span data-stu-id="406e6-348">Read</span></span>|

#### <a name="internetmessageid-string"></a><span data-ttu-id="406e6-349">internetMessageId :String</span><span class="sxs-lookup"><span data-stu-id="406e6-349">internetMessageId :String</span></span>

<span data-ttu-id="406e6-p114">Ruft die Internetnachricht-ID einer E-Mail-Nachricht ab. Nur Lesemodus.</span><span class="sxs-lookup"><span data-stu-id="406e6-p114">Gets the Internet message identifier for an email message. Read mode only.</span></span>

##### <a name="type"></a><span data-ttu-id="406e6-352">Typ:</span><span class="sxs-lookup"><span data-stu-id="406e6-352">Type:</span></span>

*   <span data-ttu-id="406e6-353">String</span><span class="sxs-lookup"><span data-stu-id="406e6-353">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="406e6-354">Anforderungen</span><span class="sxs-lookup"><span data-stu-id="406e6-354">Requirements</span></span>

|<span data-ttu-id="406e6-355">Anforderung</span><span class="sxs-lookup"><span data-stu-id="406e6-355">Requirement</span></span>| <span data-ttu-id="406e6-356">Wert</span><span class="sxs-lookup"><span data-stu-id="406e6-356">Value</span></span>|
|---|---|
|[<span data-ttu-id="406e6-357">Mindestversion des Postfachanforderungssatzes</span><span class="sxs-lookup"><span data-stu-id="406e6-357">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="406e6-358">1.0</span><span class="sxs-lookup"><span data-stu-id="406e6-358">1.0</span></span>|
|[<span data-ttu-id="406e6-359">Mindestberechtigungsstufe</span><span class="sxs-lookup"><span data-stu-id="406e6-359">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="406e6-360">ReadItem</span><span class="sxs-lookup"><span data-stu-id="406e6-360">ReadItem</span></span>|
|[<span data-ttu-id="406e6-361">Zutreffender Outlook-Modus</span><span class="sxs-lookup"><span data-stu-id="406e6-361">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="406e6-362">Lesen</span><span class="sxs-lookup"><span data-stu-id="406e6-362">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="406e6-363">Beispiel</span><span class="sxs-lookup"><span data-stu-id="406e6-363">Example</span></span>

```
var internetMessageId = Office.context.mailbox.item.internetMessageId;
```

#### <a name="itemclass-string"></a><span data-ttu-id="406e6-364">itemClass :String</span><span class="sxs-lookup"><span data-stu-id="406e6-364">itemClass :String</span></span>

<span data-ttu-id="406e6-p115">Ruft die Exchange-Webdienste-Elementklasse des ausgewählten Elements ab. Nur Lesemodus.</span><span class="sxs-lookup"><span data-stu-id="406e6-p115">Gets the Exchange Web Services item class of the selected item. Read mode only.</span></span>

<span data-ttu-id="406e6-p116">Die `itemClass`-Eigenschaft gibt die Nachrichtenklasse des ausgewählten Elements an. Folgende Standardnachrichtenklassen für das Nachrichten- oder Terminelement sind vorhanden:</span><span class="sxs-lookup"><span data-stu-id="406e6-p116">The `itemClass` property specifies the message class of the selected item. The following are the default message classes for the message or appointment item.</span></span>

| <span data-ttu-id="406e6-369">Typ</span><span class="sxs-lookup"><span data-stu-id="406e6-369">Type</span></span> | <span data-ttu-id="406e6-370">Beschreibung</span><span class="sxs-lookup"><span data-stu-id="406e6-370">Description</span></span> | <span data-ttu-id="406e6-371">Elementklasse</span><span class="sxs-lookup"><span data-stu-id="406e6-371">item class</span></span> |
| --- | --- | --- |
| <span data-ttu-id="406e6-372">Terminelemente</span><span class="sxs-lookup"><span data-stu-id="406e6-372">Appointment items</span></span> | <span data-ttu-id="406e6-373">Dies sind Kalenderelemente der Elementklasse `IPM.Appointment` oder `IPM.Appointment.Occurence`.</span><span class="sxs-lookup"><span data-stu-id="406e6-373">These are calendar items of the item class `IPM.Appointment` or `IPM.Appointment.Occurence`.</span></span> | `IPM.Appointment`<br />`IPM.Appointment.Occurence` |
| <span data-ttu-id="406e6-374">Nachrichtenelemente</span><span class="sxs-lookup"><span data-stu-id="406e6-374">Message items</span></span> | <span data-ttu-id="406e6-375">Diese Elemente umfassen E-Mail-Nachrichten der Standardnachrichtenklasse `IPM.Note` sowie Besprechungsanfragen, -antworten und -absagen, die `IPM.Schedule.Meeting` als Basisnachrichtenklasse verwenden.</span><span class="sxs-lookup"><span data-stu-id="406e6-375">These include email messages that have the default message class `IPM.Note`, and meeting requests, responses, and cancellations, that use `IPM.Schedule.Meeting` as the base message class.</span></span> | `IPM.Note`<br />`IPM.Schedule.Meeting.Request`<br />`IPM.Schedule.Meeting.Neg`<br />`IPM.Schedule.Meeting.Pos`<br />`IPM.Schedule.Meeting.Tent`<br />`IPM.Schedule.Meeting.Canceled` |

<span data-ttu-id="406e6-376">Sie können benutzerdefinierte Nachrichtenklassen zum Erweitern einer Standardnachrichtenklasse erstellen, z. B. eine benutzerdefinierte Terminnachrichtenklasse wie `IPM.Appointment.Contoso`.</span><span class="sxs-lookup"><span data-stu-id="406e6-376">You can create custom message classes that extends a default message class, for example, a custom appointment message class `IPM.Appointment.Contoso`.</span></span>

##### <a name="type"></a><span data-ttu-id="406e6-377">Typ:</span><span class="sxs-lookup"><span data-stu-id="406e6-377">Type:</span></span>

*   <span data-ttu-id="406e6-378">String</span><span class="sxs-lookup"><span data-stu-id="406e6-378">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="406e6-379">Anforderungen</span><span class="sxs-lookup"><span data-stu-id="406e6-379">Requirements</span></span>

|<span data-ttu-id="406e6-380">Anforderung</span><span class="sxs-lookup"><span data-stu-id="406e6-380">Requirement</span></span>| <span data-ttu-id="406e6-381">Wert</span><span class="sxs-lookup"><span data-stu-id="406e6-381">Value</span></span>|
|---|---|
|[<span data-ttu-id="406e6-382">Mindestversion des Postfachanforderungssatzes</span><span class="sxs-lookup"><span data-stu-id="406e6-382">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="406e6-383">1.0</span><span class="sxs-lookup"><span data-stu-id="406e6-383">1.0</span></span>|
|[<span data-ttu-id="406e6-384">Mindestberechtigungsstufe</span><span class="sxs-lookup"><span data-stu-id="406e6-384">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="406e6-385">ReadItem</span><span class="sxs-lookup"><span data-stu-id="406e6-385">ReadItem</span></span>|
|[<span data-ttu-id="406e6-386">Zutreffender Outlook-Modus</span><span class="sxs-lookup"><span data-stu-id="406e6-386">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="406e6-387">Lesen</span><span class="sxs-lookup"><span data-stu-id="406e6-387">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="406e6-388">Beispiel</span><span class="sxs-lookup"><span data-stu-id="406e6-388">Example</span></span>

```
var itemClass = Office.context.mailbox.item.itemClass;
```

#### <a name="nullable-itemid-string"></a><span data-ttu-id="406e6-389">(nullable) itemId :String</span><span class="sxs-lookup"><span data-stu-id="406e6-389">(nullable) itemId :String</span></span>

<span data-ttu-id="406e6-p117">Ruft den Exchange-Webdienste-Elementbezeichner für das aktuelle Element ab. Nur Lesemodus.</span><span class="sxs-lookup"><span data-stu-id="406e6-p117">Gets the Exchange Web Services item identifier for the current item. Read mode only.</span></span>

> [!NOTE]
> <span data-ttu-id="406e6-392">Der Bezeichner, der von der `itemId`-Eigenschaft zurückgegeben wird, ist der gleiche Bezeichner wie der Exchange-Webdienste-Elementbezeichner.</span><span class="sxs-lookup"><span data-stu-id="406e6-392">The identifier returned by the `itemId` property is the same as the Exchange Web Services item identifier.</span></span> <span data-ttu-id="406e6-393">Die `itemId` -Eigenschaft ist nicht identisch mit der Outlook-Eintrags-ID oder die ID von der Outlook-REST-API verwendet.</span><span class="sxs-lookup"><span data-stu-id="406e6-393">The `itemId` property is not identical to the Outlook Entry ID or the ID used by the Outlook REST API.</span></span> <span data-ttu-id="406e6-394">REST-API-Aufrufe dieser Wert verwenden, sollten sie konvertiert werden, bevor [Office.context.mailbox.convertToRestId](office.context.mailbox.md#converttorestiditemid-restversion--string)verwenden.</span><span class="sxs-lookup"><span data-stu-id="406e6-394">Before making REST API calls using this value, it should be converted using [Office.context.mailbox.convertToRestId](office.context.mailbox.md#converttorestiditemid-restversion--string).</span></span> <span data-ttu-id="406e6-395">Weitere Informationen finden Sie unter [Verwenden der Outlook-REST-APIs aus einem Outlook-add-in](https://docs.microsoft.com/outlook/add-ins/use-rest-api#get-the-item-id).</span><span class="sxs-lookup"><span data-stu-id="406e6-395">For more details, see [Use the Outlook REST APIs from an Outlook add-in](https://docs.microsoft.com/outlook/add-ins/use-rest-api#get-the-item-id).</span></span>

<span data-ttu-id="406e6-p119">Die Eigenschaft `itemId` ist im Modus „Verfassen“ nicht verfügbar. Wenn ein Elementbezeichner erforderlich ist, kann die [`saveAsync`](#saveasyncoptions-callback)-Methode verwendet werden, um das Element zu speichern. Dadurch wird der Elementbezeichner im [`AsyncResult.value`](/javascript/api/office/office.asyncresult)-Parameter in der Rückruffunktion zurückgegeben.</span><span class="sxs-lookup"><span data-stu-id="406e6-p119">The `itemId` property is not available in compose mode. If an item identifier is required, the [`saveAsync`](#saveasyncoptions-callback) method can be used to save the item to the store, which will return the item identifier in the [`AsyncResult.value`](/javascript/api/office/office.asyncresult) parameter in the callback function.</span></span>

##### <a name="type"></a><span data-ttu-id="406e6-398">Typ:</span><span class="sxs-lookup"><span data-stu-id="406e6-398">Type:</span></span>

*   <span data-ttu-id="406e6-399">String</span><span class="sxs-lookup"><span data-stu-id="406e6-399">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="406e6-400">Anforderungen</span><span class="sxs-lookup"><span data-stu-id="406e6-400">Requirements</span></span>

|<span data-ttu-id="406e6-401">Anforderung</span><span class="sxs-lookup"><span data-stu-id="406e6-401">Requirement</span></span>| <span data-ttu-id="406e6-402">Wert</span><span class="sxs-lookup"><span data-stu-id="406e6-402">Value</span></span>|
|---|---|
|[<span data-ttu-id="406e6-403">Mindestversion des Postfachanforderungssatzes</span><span class="sxs-lookup"><span data-stu-id="406e6-403">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="406e6-404">1.0</span><span class="sxs-lookup"><span data-stu-id="406e6-404">1.0</span></span>|
|[<span data-ttu-id="406e6-405">Mindestberechtigungsstufe</span><span class="sxs-lookup"><span data-stu-id="406e6-405">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="406e6-406">ReadItem</span><span class="sxs-lookup"><span data-stu-id="406e6-406">ReadItem</span></span>|
|[<span data-ttu-id="406e6-407">Zutreffender Outlook-Modus</span><span class="sxs-lookup"><span data-stu-id="406e6-407">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="406e6-408">Lesen</span><span class="sxs-lookup"><span data-stu-id="406e6-408">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="406e6-409">Beispiel</span><span class="sxs-lookup"><span data-stu-id="406e6-409">Example</span></span>

<span data-ttu-id="406e6-p120">Mit dem folgende Code wird das Vorhandensein eines Elementbezeichners überprüft. Wenn die `itemId`-Eigenschaft `null` oder `undefined` zurückgibt, speichert sie das Element und ruft den Elementbezeichner aus dem asynchronen Ergebnis ab.</span><span class="sxs-lookup"><span data-stu-id="406e6-p120">The following code checks for the presence of an item identifier. If the `itemId` property returns `null` or `undefined`, it saves the item to the store and gets the item identifier from the asynchronous result.</span></span>

```
var itemId = Office.context.mailbox.item.itemId;
if (itemId === null || itemId == undefined) {
  Office.context.mailbox.item.saveAsync(function(result){
    itemId = result.value;
  });
}
```

####  <a name="itemtype-officemailboxenumsitemtypejavascriptapioutlook15officemailboxenumsitemtype"></a><span data-ttu-id="406e6-412">itemType :[Office.MailboxEnums.ItemType](/javascript/api/outlook_1_5/office.mailboxenums.itemtype)</span><span class="sxs-lookup"><span data-stu-id="406e6-412">itemType :[Office.MailboxEnums.ItemType](/javascript/api/outlook_1_5/office.mailboxenums.itemtype)</span></span>

<span data-ttu-id="406e6-413">Ruft den Typ des Elements ab, das eine Instanz darstellt.</span><span class="sxs-lookup"><span data-stu-id="406e6-413">Gets the type of item that an instance represents.</span></span>

<span data-ttu-id="406e6-414">Die `itemType`-Eigenschaft gibt einen der Werte der `ItemType`-Enumeration zurück, der angibt, ob es sich bei der `item`-Objektinstanz um eine Nachricht oder einen Termin handelt.</span><span class="sxs-lookup"><span data-stu-id="406e6-414">The `itemType` property returns one of the `ItemType` enumeration values, indicating whether the `item` object instance is a message or an appointment.</span></span>

##### <a name="type"></a><span data-ttu-id="406e6-415">Typ:</span><span class="sxs-lookup"><span data-stu-id="406e6-415">Type:</span></span>

*   [<span data-ttu-id="406e6-416">Office.MailboxEnums.ItemType</span><span class="sxs-lookup"><span data-stu-id="406e6-416">Office.MailboxEnums.ItemType</span></span>](/javascript/api/outlook_1_5/office.mailboxenums.itemtype)

##### <a name="requirements"></a><span data-ttu-id="406e6-417">Anforderungen</span><span class="sxs-lookup"><span data-stu-id="406e6-417">Requirements</span></span>

|<span data-ttu-id="406e6-418">Anforderung</span><span class="sxs-lookup"><span data-stu-id="406e6-418">Requirement</span></span>| <span data-ttu-id="406e6-419">Wert</span><span class="sxs-lookup"><span data-stu-id="406e6-419">Value</span></span>|
|---|---|
|[<span data-ttu-id="406e6-420">Mindestversion des Postfachanforderungssatzes</span><span class="sxs-lookup"><span data-stu-id="406e6-420">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="406e6-421">1.0</span><span class="sxs-lookup"><span data-stu-id="406e6-421">1.0</span></span>|
|[<span data-ttu-id="406e6-422">Mindestberechtigungsstufe</span><span class="sxs-lookup"><span data-stu-id="406e6-422">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="406e6-423">ReadItem</span><span class="sxs-lookup"><span data-stu-id="406e6-423">ReadItem</span></span>|
|[<span data-ttu-id="406e6-424">Zutreffender Outlook-Modus</span><span class="sxs-lookup"><span data-stu-id="406e6-424">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="406e6-425">Verfassen oder Lesen</span><span class="sxs-lookup"><span data-stu-id="406e6-425">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="406e6-426">Beispiel</span><span class="sxs-lookup"><span data-stu-id="406e6-426">Example</span></span>

```
if (Office.context.mailbox.item.itemType == Office.MailboxEnums.ItemType.Message)
  // do something
else
  // do something else
```

####  <a name="location-stringlocationjavascriptapioutlook15officelocation"></a><span data-ttu-id="406e6-427">location :String|[Location](/javascript/api/outlook_1_5/office.location)</span><span class="sxs-lookup"><span data-stu-id="406e6-427">location :String|[Location](/javascript/api/outlook_1_5/office.location)</span></span>

<span data-ttu-id="406e6-428">Ruft den Ort eines Termins ab bzw. legt ihn fest.</span><span class="sxs-lookup"><span data-stu-id="406e6-428">Gets or sets the location of an appointment.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="406e6-429">Lesemodus</span><span class="sxs-lookup"><span data-stu-id="406e6-429">Read mode</span></span>

<span data-ttu-id="406e6-430">Die `location`-Eigenschaft gibt eine Zeichenfolge zurück, die den Ort des Termins enthält.</span><span class="sxs-lookup"><span data-stu-id="406e6-430">The `location` property returns a string that contains the location of the appointment.</span></span>

##### <a name="compose-mode"></a><span data-ttu-id="406e6-431">Verfassenmodus</span><span class="sxs-lookup"><span data-stu-id="406e6-431">Compose mode</span></span>

<span data-ttu-id="406e6-432">Die `location`-Eigenschaft gibt ein `Location`-Objekt zurück, das Methoden zum Abrufen und Festlegen des Orts für einen Termin bereitstellt.</span><span class="sxs-lookup"><span data-stu-id="406e6-432">The `location` property returns a `Location` object that provides methods that are used to get and set the location of the appointment.</span></span>

##### <a name="type"></a><span data-ttu-id="406e6-433">Typ:</span><span class="sxs-lookup"><span data-stu-id="406e6-433">Type:</span></span>

*   <span data-ttu-id="406e6-434">String | [Location](/javascript/api/outlook_1_5/office.location)</span><span class="sxs-lookup"><span data-stu-id="406e6-434">String | [Location](/javascript/api/outlook_1_5/office.location)</span></span>

##### <a name="requirements"></a><span data-ttu-id="406e6-435">Anforderungen</span><span class="sxs-lookup"><span data-stu-id="406e6-435">Requirements</span></span>

|<span data-ttu-id="406e6-436">Anforderung</span><span class="sxs-lookup"><span data-stu-id="406e6-436">Requirement</span></span>| <span data-ttu-id="406e6-437">Wert</span><span class="sxs-lookup"><span data-stu-id="406e6-437">Value</span></span>|
|---|---|
|[<span data-ttu-id="406e6-438">Mindestversion des Postfachanforderungssatzes</span><span class="sxs-lookup"><span data-stu-id="406e6-438">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="406e6-439">1.0</span><span class="sxs-lookup"><span data-stu-id="406e6-439">1.0</span></span>|
|[<span data-ttu-id="406e6-440">Mindestberechtigungsstufe</span><span class="sxs-lookup"><span data-stu-id="406e6-440">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="406e6-441">ReadItem</span><span class="sxs-lookup"><span data-stu-id="406e6-441">ReadItem</span></span>|
|[<span data-ttu-id="406e6-442">Zutreffender Outlook-Modus</span><span class="sxs-lookup"><span data-stu-id="406e6-442">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="406e6-443">Verfassen oder Lesen</span><span class="sxs-lookup"><span data-stu-id="406e6-443">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="406e6-444">Beispiel</span><span class="sxs-lookup"><span data-stu-id="406e6-444">Example</span></span>

```
var userContext = { value : 1 };
Office.context.mailbox.item.location.getAsync( { context: userContext}, callback);

function callback(asyncResult) {
  var context = asyncResult.context;
  var location = asyncResult.value;
}
```

#### <a name="normalizedsubject-string"></a><span data-ttu-id="406e6-445">normalizedSubject :String</span><span class="sxs-lookup"><span data-stu-id="406e6-445">normalizedSubject :String</span></span>

<span data-ttu-id="406e6-p121">Ruft den Betreff für ein Element ab, wobei alle Präfixe entfernt werden (einschließlich `RE:` und `FWD:`). Nur Lesemodus.</span><span class="sxs-lookup"><span data-stu-id="406e6-p121">Gets the subject of an item, with all prefixes removed (including `RE:` and `FWD:`). Read mode only.</span></span>

<span data-ttu-id="406e6-p122">Die normalizedSubject-Eigenschaft ruft den Betreff des Elements mit allen Standardpräfixen (z. B. `RE:` und `FW:`) ab, die von E-Mail-Programmen hinzugefügt werden. Wenn der Betreff des Elements mit vollständigen Präfixen abgerufen werden soll, verwenden Sie die [`subject`](#subject-stringsubjectjavascriptapioutlook15officesubject)-Eigenschaft.</span><span class="sxs-lookup"><span data-stu-id="406e6-p122">The normalizedSubject property gets the subject of the item, with any standard prefixes (such as `RE:` and `FW:`) that are added by email programs. To get the subject of the item with the prefixes intact, use the [`subject`](#subject-stringsubjectjavascriptapioutlook15officesubject) property.</span></span>

##### <a name="type"></a><span data-ttu-id="406e6-450">Typ:</span><span class="sxs-lookup"><span data-stu-id="406e6-450">Type:</span></span>

*   <span data-ttu-id="406e6-451">String</span><span class="sxs-lookup"><span data-stu-id="406e6-451">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="406e6-452">Anforderungen</span><span class="sxs-lookup"><span data-stu-id="406e6-452">Requirements</span></span>

|<span data-ttu-id="406e6-453">Anforderung</span><span class="sxs-lookup"><span data-stu-id="406e6-453">Requirement</span></span>| <span data-ttu-id="406e6-454">Wert</span><span class="sxs-lookup"><span data-stu-id="406e6-454">Value</span></span>|
|---|---|
|[<span data-ttu-id="406e6-455">Mindestversion des Postfachanforderungssatzes</span><span class="sxs-lookup"><span data-stu-id="406e6-455">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="406e6-456">1.0</span><span class="sxs-lookup"><span data-stu-id="406e6-456">1.0</span></span>|
|[<span data-ttu-id="406e6-457">Mindestberechtigungsstufe</span><span class="sxs-lookup"><span data-stu-id="406e6-457">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="406e6-458">ReadItem</span><span class="sxs-lookup"><span data-stu-id="406e6-458">ReadItem</span></span>|
|[<span data-ttu-id="406e6-459">Zutreffender Outlook-Modus</span><span class="sxs-lookup"><span data-stu-id="406e6-459">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="406e6-460">Lesen</span><span class="sxs-lookup"><span data-stu-id="406e6-460">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="406e6-461">Beispiel</span><span class="sxs-lookup"><span data-stu-id="406e6-461">Example</span></span>

```
var normalizedSubject = Office.context.mailbox.item.normalizedSubject;
```

####  <a name="notificationmessages-notificationmessagesjavascriptapioutlook15officenotificationmessages"></a><span data-ttu-id="406e6-462">notificationMessages :[NotificationMessages](/javascript/api/outlook_1_5/office.notificationmessages)</span><span class="sxs-lookup"><span data-stu-id="406e6-462">notificationMessages :[NotificationMessages](/javascript/api/outlook_1_5/office.notificationmessages)</span></span>

<span data-ttu-id="406e6-463">Ruft die Benachrichtigungen für ein Element ab.</span><span class="sxs-lookup"><span data-stu-id="406e6-463">Gets the notification messages for an item.</span></span>

##### <a name="type"></a><span data-ttu-id="406e6-464">Typ:</span><span class="sxs-lookup"><span data-stu-id="406e6-464">Type:</span></span>

*   [<span data-ttu-id="406e6-465">NotificationMessages</span><span class="sxs-lookup"><span data-stu-id="406e6-465">NotificationMessages</span></span>](/javascript/api/outlook_1_5/office.notificationmessages)

##### <a name="requirements"></a><span data-ttu-id="406e6-466">Anforderungen</span><span class="sxs-lookup"><span data-stu-id="406e6-466">Requirements</span></span>

|<span data-ttu-id="406e6-467">Anforderung</span><span class="sxs-lookup"><span data-stu-id="406e6-467">Requirement</span></span>| <span data-ttu-id="406e6-468">Wert</span><span class="sxs-lookup"><span data-stu-id="406e6-468">Value</span></span>|
|---|---|
|[<span data-ttu-id="406e6-469">Mindestversion des Postfachanforderungssatzes</span><span class="sxs-lookup"><span data-stu-id="406e6-469">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="406e6-470">1.3</span><span class="sxs-lookup"><span data-stu-id="406e6-470">1.3</span></span>|
|[<span data-ttu-id="406e6-471">Mindestberechtigungsstufe</span><span class="sxs-lookup"><span data-stu-id="406e6-471">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="406e6-472">ReadItem</span><span class="sxs-lookup"><span data-stu-id="406e6-472">ReadItem</span></span>|
|[<span data-ttu-id="406e6-473">Zutreffender Outlook-Modus</span><span class="sxs-lookup"><span data-stu-id="406e6-473">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="406e6-474">Verfassen oder Lesen</span><span class="sxs-lookup"><span data-stu-id="406e6-474">Compose or read</span></span>|

####  <a name="optionalattendees-arrayemailaddressdetailsjavascriptapioutlook15officeemailaddressdetailsrecipientsjavascriptapioutlook15officerecipients"></a><span data-ttu-id="406e6-475">optionalAttendees :Array.<[EmailAddressDetails](/javascript/api/outlook_1_5/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook_1_5/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="406e6-475">optionalAttendees :Array.<[EmailAddressDetails](/javascript/api/outlook_1_5/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook_1_5/office.recipients)</span></span>

<span data-ttu-id="406e6-476">Ermöglicht den Zugriff auf die optionalen Teilnehmer eines Ereignisses.</span><span class="sxs-lookup"><span data-stu-id="406e6-476">Provides access to the optional attendees of an event.</span></span> <span data-ttu-id="406e6-477">Der Typ des Objekts und die Zugriffsebene, hängt von den Modus des aktuellen Elements ab.</span><span class="sxs-lookup"><span data-stu-id="406e6-477">The type of object and level of access depends on the mode of the current item.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="406e6-478">Lesemodus</span><span class="sxs-lookup"><span data-stu-id="406e6-478">Read mode</span></span>

<span data-ttu-id="406e6-479">Die `optionalAttendees`-Eigenschaft gibt ein Array mit einem `EmailAddressDetails`-Objekt für jeden optionalen Teilnehmer an der Besprechung zurück.</span><span class="sxs-lookup"><span data-stu-id="406e6-479">The `optionalAttendees` property returns an array that contains an `EmailAddressDetails` object for each optional attendee to the meeting.</span></span>

##### <a name="compose-mode"></a><span data-ttu-id="406e6-480">Verfassenmodus</span><span class="sxs-lookup"><span data-stu-id="406e6-480">Compose mode</span></span>

<span data-ttu-id="406e6-481">Die `optionalAttendees` -Eigenschaft gibt eine `Recipients` -Objekt, das Methoden zum Abrufen oder Aktualisieren der optionalen Teilnehmer für eine Besprechung bereitstellt.</span><span class="sxs-lookup"><span data-stu-id="406e6-481">The `optionalAttendees` property returns a `Recipients` object that provides methods to get or update the optional attendees for a meeting.</span></span>

##### <a name="type"></a><span data-ttu-id="406e6-482">Typ:</span><span class="sxs-lookup"><span data-stu-id="406e6-482">Type:</span></span>

*   <span data-ttu-id="406e6-483">Array.<[EmailAddressDetails](/javascript/api/outlook_1_5/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook_1_5/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="406e6-483">Array.<[EmailAddressDetails](/javascript/api/outlook_1_5/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook_1_5/office.recipients)</span></span>

##### <a name="requirements"></a><span data-ttu-id="406e6-484">Anforderungen</span><span class="sxs-lookup"><span data-stu-id="406e6-484">Requirements</span></span>

|<span data-ttu-id="406e6-485">Anforderung</span><span class="sxs-lookup"><span data-stu-id="406e6-485">Requirement</span></span>| <span data-ttu-id="406e6-486">Wert</span><span class="sxs-lookup"><span data-stu-id="406e6-486">Value</span></span>|
|---|---|
|[<span data-ttu-id="406e6-487">Mindestversion des Postfachanforderungssatzes</span><span class="sxs-lookup"><span data-stu-id="406e6-487">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="406e6-488">1.0</span><span class="sxs-lookup"><span data-stu-id="406e6-488">1.0</span></span>|
|[<span data-ttu-id="406e6-489">Mindestberechtigungsstufe</span><span class="sxs-lookup"><span data-stu-id="406e6-489">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="406e6-490">ReadItem</span><span class="sxs-lookup"><span data-stu-id="406e6-490">ReadItem</span></span>|
|[<span data-ttu-id="406e6-491">Zutreffender Outlook-Modus</span><span class="sxs-lookup"><span data-stu-id="406e6-491">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="406e6-492">Verfassen oder Lesen</span><span class="sxs-lookup"><span data-stu-id="406e6-492">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="406e6-493">Beispiel</span><span class="sxs-lookup"><span data-stu-id="406e6-493">Example</span></span>

```
Office.context.mailbox.item.optionalAttendees.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.optionalAttendees.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.optionalAttendees.getAsync(callback);

function callback(asyncResult) {
  var arrayOfOptionalAttendeesRecipients = asyncResult.value;
}
```

#### <a name="organizer-emailaddressdetailsjavascriptapioutlook15officeemailaddressdetails"></a><span data-ttu-id="406e6-494">organizer :[EmailAddressDetails](/javascript/api/outlook_1_5/office.emailaddressdetails)</span><span class="sxs-lookup"><span data-stu-id="406e6-494">organizer :[EmailAddressDetails](/javascript/api/outlook_1_5/office.emailaddressdetails)</span></span>

<span data-ttu-id="406e6-p124">Ruft für eine angegebene Besprechung die E-Mail-Adresse des Organisators der Besprechung ab. Nur Lesemodus.</span><span class="sxs-lookup"><span data-stu-id="406e6-p124">Gets the email address of the meeting organizer for a specified meeting. Read mode only.</span></span>

##### <a name="type"></a><span data-ttu-id="406e6-497">Typ:</span><span class="sxs-lookup"><span data-stu-id="406e6-497">Type:</span></span>

*   [<span data-ttu-id="406e6-498">EmailAddressDetails</span><span class="sxs-lookup"><span data-stu-id="406e6-498">EmailAddressDetails</span></span>](/javascript/api/outlook_1_5/office.emailaddressdetails)

##### <a name="requirements"></a><span data-ttu-id="406e6-499">Anforderungen</span><span class="sxs-lookup"><span data-stu-id="406e6-499">Requirements</span></span>

|<span data-ttu-id="406e6-500">Anforderung</span><span class="sxs-lookup"><span data-stu-id="406e6-500">Requirement</span></span>| <span data-ttu-id="406e6-501">Wert</span><span class="sxs-lookup"><span data-stu-id="406e6-501">Value</span></span>|
|---|---|
|[<span data-ttu-id="406e6-502">Mindestversion des Postfachanforderungssatzes</span><span class="sxs-lookup"><span data-stu-id="406e6-502">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="406e6-503">1.0</span><span class="sxs-lookup"><span data-stu-id="406e6-503">1.0</span></span>|
|[<span data-ttu-id="406e6-504">Mindestberechtigungsstufe</span><span class="sxs-lookup"><span data-stu-id="406e6-504">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="406e6-505">ReadItem</span><span class="sxs-lookup"><span data-stu-id="406e6-505">ReadItem</span></span>|
|[<span data-ttu-id="406e6-506">Zutreffender Outlook-Modus</span><span class="sxs-lookup"><span data-stu-id="406e6-506">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="406e6-507">Lesen</span><span class="sxs-lookup"><span data-stu-id="406e6-507">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="406e6-508">Beispiel</span><span class="sxs-lookup"><span data-stu-id="406e6-508">Example</span></span>

```
var organizerName = Office.context.mailbox.item.organizer.displayName;
var organizerAddress = Office.context.mailbox.item.organizer.emailAddress;
```

####  <a name="requiredattendees-arrayemailaddressdetailsjavascriptapioutlook15officeemailaddressdetailsrecipientsjavascriptapioutlook15officerecipients"></a><span data-ttu-id="406e6-509">requiredAttendees :Array.<[EmailAddressDetails](/javascript/api/outlook_1_5/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook_1_5/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="406e6-509">requiredAttendees :Array.<[EmailAddressDetails](/javascript/api/outlook_1_5/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook_1_5/office.recipients)</span></span>

<span data-ttu-id="406e6-510">Ermöglicht den Zugriff auf die erforderlichen Teilnehmer eines Ereignisses.</span><span class="sxs-lookup"><span data-stu-id="406e6-510">Provides access to the required attendees of an event.</span></span> <span data-ttu-id="406e6-511">Der Typ des Objekts und die Zugriffsebene, hängt von den Modus des aktuellen Elements ab.</span><span class="sxs-lookup"><span data-stu-id="406e6-511">The type of object and level of access depends on the mode of the current item.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="406e6-512">Lesemodus</span><span class="sxs-lookup"><span data-stu-id="406e6-512">Read mode</span></span>

<span data-ttu-id="406e6-513">Die `requiredAttendees`-Eigenschaft gibt ein Array mit einem `EmailAddressDetails`-Objekt für jeden erforderlichen Teilnehmer an der Besprechung zurück.</span><span class="sxs-lookup"><span data-stu-id="406e6-513">The `requiredAttendees` property returns an array that contains an `EmailAddressDetails` object for each required attendee to the meeting.</span></span>

##### <a name="compose-mode"></a><span data-ttu-id="406e6-514">Verfassenmodus</span><span class="sxs-lookup"><span data-stu-id="406e6-514">Compose mode</span></span>

<span data-ttu-id="406e6-515">Die `requiredAttendees` -Eigenschaft gibt eine `Recipients` -Objekt, das Methoden zum Abrufen oder Aktualisieren der erforderlichen Teilnehmer für eine Besprechung bereitstellt.</span><span class="sxs-lookup"><span data-stu-id="406e6-515">The `requiredAttendees` property returns a `Recipients` object that provides methods to get or update the required attendees for a meeting.</span></span>

##### <a name="type"></a><span data-ttu-id="406e6-516">Typ:</span><span class="sxs-lookup"><span data-stu-id="406e6-516">Type:</span></span>

*   <span data-ttu-id="406e6-517">Array.<[EmailAddressDetails](/javascript/api/outlook_1_5/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook_1_5/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="406e6-517">Array.<[EmailAddressDetails](/javascript/api/outlook_1_5/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook_1_5/office.recipients)</span></span>

##### <a name="requirements"></a><span data-ttu-id="406e6-518">Anforderungen</span><span class="sxs-lookup"><span data-stu-id="406e6-518">Requirements</span></span>

|<span data-ttu-id="406e6-519">Anforderung</span><span class="sxs-lookup"><span data-stu-id="406e6-519">Requirement</span></span>| <span data-ttu-id="406e6-520">Wert</span><span class="sxs-lookup"><span data-stu-id="406e6-520">Value</span></span>|
|---|---|
|[<span data-ttu-id="406e6-521">Mindestversion des Postfachanforderungssatzes</span><span class="sxs-lookup"><span data-stu-id="406e6-521">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="406e6-522">1.0</span><span class="sxs-lookup"><span data-stu-id="406e6-522">1.0</span></span>|
|[<span data-ttu-id="406e6-523">Mindestberechtigungsstufe</span><span class="sxs-lookup"><span data-stu-id="406e6-523">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="406e6-524">ReadItem</span><span class="sxs-lookup"><span data-stu-id="406e6-524">ReadItem</span></span>|
|[<span data-ttu-id="406e6-525">Zutreffender Outlook-Modus</span><span class="sxs-lookup"><span data-stu-id="406e6-525">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="406e6-526">Verfassen oder Lesen</span><span class="sxs-lookup"><span data-stu-id="406e6-526">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="406e6-527">Beispiel</span><span class="sxs-lookup"><span data-stu-id="406e6-527">Example</span></span>

```
Office.context.mailbox.item.requiredAttendees.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.requiredAttendees.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.requiredAttendees.getAsync(callback);

function callback(asyncResult) {
  var arrayOfRequiredAttendeesRecipients = asyncResult.value;
}
```

#### <a name="sender-emailaddressdetailsjavascriptapioutlook15officeemailaddressdetails"></a><span data-ttu-id="406e6-528">sender :[EmailAddressDetails](/javascript/api/outlook_1_5/office.emailaddressdetails)</span><span class="sxs-lookup"><span data-stu-id="406e6-528">sender :[EmailAddressDetails](/javascript/api/outlook_1_5/office.emailaddressdetails)</span></span>

<span data-ttu-id="406e6-p126">Ruft die E-Mail-Adresse des Absenders einer E-Mail-Nachricht ab. Nur Lesemodus.</span><span class="sxs-lookup"><span data-stu-id="406e6-p126">Gets the email address of the sender of an email message. Read mode only.</span></span>

<span data-ttu-id="406e6-p127">Die [`from`](#from-emailaddressdetailsjavascriptapioutlook15officeemailaddressdetails)- und `sender`-Eigenschaften stellen dieselbe Person dar, außer die Nachricht wurde von einem Delegaten gesendet. In diesem Fall stellt die `from`-Eigenschaft die Stellvertretung dar, und die sender-Eigenschaft stellt den Delegaten dar.</span><span class="sxs-lookup"><span data-stu-id="406e6-p127">The [`from`](#from-emailaddressdetailsjavascriptapioutlook15officeemailaddressdetails) and `sender` properties represent the same person unless the message is sent by a delegate. In that case, the `from` property represents the delegator, and the sender property represents the delegate.</span></span>

> [!NOTE]
> <span data-ttu-id="406e6-533">Die `recipientType` -Eigenschaft des der `EmailAddressDetails` -Objekts der `sender` -Eigenschaft ist `undefined`.</span><span class="sxs-lookup"><span data-stu-id="406e6-533">The `recipientType` property of the `EmailAddressDetails` object in the `sender` property is `undefined`.</span></span>

##### <a name="type"></a><span data-ttu-id="406e6-534">Typ:</span><span class="sxs-lookup"><span data-stu-id="406e6-534">Type:</span></span>

*   [<span data-ttu-id="406e6-535">EmailAddressDetails</span><span class="sxs-lookup"><span data-stu-id="406e6-535">EmailAddressDetails</span></span>](/javascript/api/outlook_1_5/office.emailaddressdetails)

##### <a name="requirements"></a><span data-ttu-id="406e6-536">Anforderungen</span><span class="sxs-lookup"><span data-stu-id="406e6-536">Requirements</span></span>

|<span data-ttu-id="406e6-537">Anforderung</span><span class="sxs-lookup"><span data-stu-id="406e6-537">Requirement</span></span>| <span data-ttu-id="406e6-538">Wert</span><span class="sxs-lookup"><span data-stu-id="406e6-538">Value</span></span>|
|---|---|
|[<span data-ttu-id="406e6-539">Mindestversion des Postfachanforderungssatzes</span><span class="sxs-lookup"><span data-stu-id="406e6-539">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="406e6-540">1.0</span><span class="sxs-lookup"><span data-stu-id="406e6-540">1.0</span></span>|
|[<span data-ttu-id="406e6-541">Mindestberechtigungsstufe</span><span class="sxs-lookup"><span data-stu-id="406e6-541">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="406e6-542">ReadItem</span><span class="sxs-lookup"><span data-stu-id="406e6-542">ReadItem</span></span>|
|[<span data-ttu-id="406e6-543">Zutreffender Outlook-Modus</span><span class="sxs-lookup"><span data-stu-id="406e6-543">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="406e6-544">Lesen</span><span class="sxs-lookup"><span data-stu-id="406e6-544">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="406e6-545">Beispiel</span><span class="sxs-lookup"><span data-stu-id="406e6-545">Example</span></span>

```
var senderName = Office.context.mailbox.item.sender.displayName;
var senderAddress = Office.context.mailbox.item.sender.emailAddress;
```

####  <a name="start-datetimejavascriptapioutlook15officetime"></a><span data-ttu-id="406e6-546">start :Date|[Time](/javascript/api/outlook_1_5/office.time)</span><span class="sxs-lookup"><span data-stu-id="406e6-546">start :Date|[Time](/javascript/api/outlook_1_5/office.time)</span></span>

<span data-ttu-id="406e6-547">Ruft Datum und Zeit für den Beginn des Termins ab oder legt Datum und Uhrzeit fest.</span><span class="sxs-lookup"><span data-stu-id="406e6-547">Gets or sets the date and time that the appointment is to begin.</span></span>

<span data-ttu-id="406e6-p128">Die `start`-Eigenschaft wird als Datums- und Uhrzeitwert in UTC ausgedrückt. Sie können die [`convertToLocalClientTime`](office.context.mailbox.md#converttolocalclienttimetimevalue--localclienttimejavascriptapioutlook15officelocalclienttime)-Methode verwenden, um den Wert in den lokalen Uhrzeit- und Datumswert des Clients umzuwandeln.</span><span class="sxs-lookup"><span data-stu-id="406e6-p128">The `start` property is expressed as a Coordinated Universal Time (UTC) date and time value. You can use the [`convertToLocalClientTime`](office.context.mailbox.md#converttolocalclienttimetimevalue--localclienttimejavascriptapioutlook15officelocalclienttime) method to convert the value to the client’s local date and time.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="406e6-550">Lesemodus</span><span class="sxs-lookup"><span data-stu-id="406e6-550">Read mode</span></span>

<span data-ttu-id="406e6-551">Die `start`-Eigenschaft gibt ein `Date`-Objekt zurück.</span><span class="sxs-lookup"><span data-stu-id="406e6-551">The `start` property returns a `Date` object.</span></span>

##### <a name="compose-mode"></a><span data-ttu-id="406e6-552">Verfassenmodus</span><span class="sxs-lookup"><span data-stu-id="406e6-552">Compose mode</span></span>

<span data-ttu-id="406e6-553">Die `start`-Eigenschaft gibt ein `Time`-Objekt zurück.</span><span class="sxs-lookup"><span data-stu-id="406e6-553">The `start` property returns a `Time` object.</span></span>

<span data-ttu-id="406e6-554">Wenn Sie die [`Time.setAsync`](/javascript/api/outlook_1_5/office.time#setasync-datetime--options--callback-)-Methode verwenden, um die Startzeit im Verfassenmodus festzulegen, sollten Sie die [`convertToUtcClientTime`](office.context.mailbox.md#converttoutcclienttimeinput--date)-Methode verwenden, um die Ortszeit auf dem Client für den Server in UTC umzuwandeln.</span><span class="sxs-lookup"><span data-stu-id="406e6-554">When you use the [`Time.setAsync`](/javascript/api/outlook_1_5/office.time#setasync-datetime--options--callback-) method to set the start time, you should use the [`convertToUtcClientTime`](office.context.mailbox.md#converttoutcclienttimeinput--date) method to convert the local time on the client to UTC for the server.</span></span>

##### <a name="type"></a><span data-ttu-id="406e6-555">Typ:</span><span class="sxs-lookup"><span data-stu-id="406e6-555">Type:</span></span>

*   <span data-ttu-id="406e6-556">Date | [Time](/javascript/api/outlook_1_5/office.time)</span><span class="sxs-lookup"><span data-stu-id="406e6-556">Date | [Time](/javascript/api/outlook_1_5/office.time)</span></span>

##### <a name="requirements"></a><span data-ttu-id="406e6-557">Anforderungen</span><span class="sxs-lookup"><span data-stu-id="406e6-557">Requirements</span></span>

|<span data-ttu-id="406e6-558">Anforderung</span><span class="sxs-lookup"><span data-stu-id="406e6-558">Requirement</span></span>| <span data-ttu-id="406e6-559">Wert</span><span class="sxs-lookup"><span data-stu-id="406e6-559">Value</span></span>|
|---|---|
|[<span data-ttu-id="406e6-560">Mindestversion des Postfachanforderungssatzes</span><span class="sxs-lookup"><span data-stu-id="406e6-560">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="406e6-561">1.0</span><span class="sxs-lookup"><span data-stu-id="406e6-561">1.0</span></span>|
|[<span data-ttu-id="406e6-562">Mindestberechtigungsstufe</span><span class="sxs-lookup"><span data-stu-id="406e6-562">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="406e6-563">ReadItem</span><span class="sxs-lookup"><span data-stu-id="406e6-563">ReadItem</span></span>|
|[<span data-ttu-id="406e6-564">Zutreffender Outlook-Modus</span><span class="sxs-lookup"><span data-stu-id="406e6-564">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="406e6-565">Verfassen oder Lesen</span><span class="sxs-lookup"><span data-stu-id="406e6-565">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="406e6-566">Beispiel</span><span class="sxs-lookup"><span data-stu-id="406e6-566">Example</span></span>

<span data-ttu-id="406e6-567">Im folgenden Beispiel wird die Startzeit eines Termins im Verfassenmodus mithilfe der [`setAsync`](/javascript/api/outlook_1_5/office.time#setasync-datetime--options--callback-)-Methode des `Time`-Objekts festgelegt.</span><span class="sxs-lookup"><span data-stu-id="406e6-567">The following example sets the start time of an appointment in compose mode by using the [`setAsync`](/javascript/api/outlook_1_5/office.time#setasync-datetime--options--callback-) method of the `Time` object.</span></span>

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

####  <a name="subject-stringsubjectjavascriptapioutlook15officesubject"></a><span data-ttu-id="406e6-568">subject :String|[Subject](/javascript/api/outlook_1_5/office.subject)</span><span class="sxs-lookup"><span data-stu-id="406e6-568">subject :String|[Subject](/javascript/api/outlook_1_5/office.subject)</span></span>

<span data-ttu-id="406e6-569">Ruft die Beschreibung ab, die im Betrefffeld eines Elements angezeigt wird, oder legt sie fest.</span><span class="sxs-lookup"><span data-stu-id="406e6-569">Gets or sets the description that appears in the subject field of an item.</span></span>

<span data-ttu-id="406e6-570">Die `subject`-Eigenschaft ruft den gesamten Betreff des Elements ab oder legt ihn fest – so, wie er vom E-Mail-Server gesendet wird.</span><span class="sxs-lookup"><span data-stu-id="406e6-570">The `subject` property gets or sets the entire subject of the item, as sent by the email server.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="406e6-571">Lesemodus</span><span class="sxs-lookup"><span data-stu-id="406e6-571">Read mode</span></span>

<span data-ttu-id="406e6-p129">Die `subject`-Eigenschaft gibt eine Zeichenfolge zurück. Verwenden Sie die [`normalizedSubject`](#normalizedsubject-string)-Eigenschaft, um den Betreff ohne führende Präfixe wie `RE:` und `FW:` abzurufen.</span><span class="sxs-lookup"><span data-stu-id="406e6-p129">The `subject` property returns a string. Use the [`normalizedSubject`](#normalizedsubject-string) property to get the subject minus any leading prefixes such as `RE:` and `FW:`.</span></span>

```
var subject = Office.context.mailbox.item.subject;
```

##### <a name="compose-mode"></a><span data-ttu-id="406e6-574">Verfassenmodus</span><span class="sxs-lookup"><span data-stu-id="406e6-574">Compose mode</span></span>

<span data-ttu-id="406e6-575">Die `subject`-Eigenschaft gibt ein `Subject`-Objekt zurück, das Methoden zum Abrufen und Festlegen des Betreffs bereitstellt.</span><span class="sxs-lookup"><span data-stu-id="406e6-575">The `subject` property returns a `Subject` object that provides methods to get and set the subject.</span></span>

```
Office.context.mailbox.item.subject.getAsync(callback);

function callback(asyncResult) {
  var subject = asyncResult.value;
}
```

##### <a name="type"></a><span data-ttu-id="406e6-576">Typ:</span><span class="sxs-lookup"><span data-stu-id="406e6-576">Type:</span></span>

*   <span data-ttu-id="406e6-577">String | [Subject](/javascript/api/outlook_1_5/office.subject)</span><span class="sxs-lookup"><span data-stu-id="406e6-577">String | [Subject](/javascript/api/outlook_1_5/office.subject)</span></span>

##### <a name="requirements"></a><span data-ttu-id="406e6-578">Anforderungen</span><span class="sxs-lookup"><span data-stu-id="406e6-578">Requirements</span></span>

|<span data-ttu-id="406e6-579">Anforderung</span><span class="sxs-lookup"><span data-stu-id="406e6-579">Requirement</span></span>| <span data-ttu-id="406e6-580">Wert</span><span class="sxs-lookup"><span data-stu-id="406e6-580">Value</span></span>|
|---|---|
|[<span data-ttu-id="406e6-581">Mindestversion des Postfachanforderungssatzes</span><span class="sxs-lookup"><span data-stu-id="406e6-581">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="406e6-582">1.0</span><span class="sxs-lookup"><span data-stu-id="406e6-582">1.0</span></span>|
|[<span data-ttu-id="406e6-583">Mindestberechtigungsstufe</span><span class="sxs-lookup"><span data-stu-id="406e6-583">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="406e6-584">ReadItem</span><span class="sxs-lookup"><span data-stu-id="406e6-584">ReadItem</span></span>|
|[<span data-ttu-id="406e6-585">Zutreffender Outlook-Modus</span><span class="sxs-lookup"><span data-stu-id="406e6-585">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="406e6-586">Verfassen oder Lesen</span><span class="sxs-lookup"><span data-stu-id="406e6-586">Compose or read</span></span>|

####  <a name="to-arrayemailaddressdetailsjavascriptapioutlook15officeemailaddressdetailsrecipientsjavascriptapioutlook15officerecipients"></a><span data-ttu-id="406e6-587">an: Array. <[EmailAddressDetails](/javascript/api/outlook_1_5/office.emailaddressdetails)>|[Empfänger](/javascript/api/outlook_1_5/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="406e6-587">to :Array.<[EmailAddressDetails](/javascript/api/outlook_1_5/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook_1_5/office.recipients)</span></span>

<span data-ttu-id="406e6-588">Ermöglicht den Zugriff auf die Empfänger in der Zeile **an** einer Nachricht.</span><span class="sxs-lookup"><span data-stu-id="406e6-588">Provides access to the recipients on the **To** line of a message.</span></span> <span data-ttu-id="406e6-589">Der Typ des Objekts und die Zugriffsebene, hängt von den Modus des aktuellen Elements ab.</span><span class="sxs-lookup"><span data-stu-id="406e6-589">The type of object and level of access depends on the mode of the current item.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="406e6-590">Lesemodus</span><span class="sxs-lookup"><span data-stu-id="406e6-590">Read mode</span></span>

<span data-ttu-id="406e6-p131">Die `to`-Eigenschaft gibt ein Array mit einem `EmailAddressDetails`-Objekt für jeden Empfänger in der**An**-Zeile der Nachricht zurück. Die Auflistung ist auf höchstens 100 Elemente beschränkt.</span><span class="sxs-lookup"><span data-stu-id="406e6-p131">The `to` property returns an array that contains an `EmailAddressDetails` object for each recipient listed on the **To** line of the message. The collection is limited to a maximum of 100 members.</span></span>

##### <a name="compose-mode"></a><span data-ttu-id="406e6-593">Verfassenmodus</span><span class="sxs-lookup"><span data-stu-id="406e6-593">Compose mode</span></span>

<span data-ttu-id="406e6-594">Die `to` -Eigenschaft gibt eine `Recipients` -Objekt, das Methoden zum Abrufen oder aktualisieren die Empfänger in der Zeile **an** der Nachricht bereitstellt.</span><span class="sxs-lookup"><span data-stu-id="406e6-594">The `to` property returns a `Recipients` object that provides methods to get or update the recipients on the **To** line of the message.</span></span>

##### <a name="type"></a><span data-ttu-id="406e6-595">Typ:</span><span class="sxs-lookup"><span data-stu-id="406e6-595">Type:</span></span>

*   <span data-ttu-id="406e6-596">Array.<[EmailAddressDetails](/javascript/api/outlook_1_5/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook_1_5/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="406e6-596">Array.<[EmailAddressDetails](/javascript/api/outlook_1_5/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook_1_5/office.recipients)</span></span>

##### <a name="requirements"></a><span data-ttu-id="406e6-597">Anforderungen</span><span class="sxs-lookup"><span data-stu-id="406e6-597">Requirements</span></span>

|<span data-ttu-id="406e6-598">Anforderung</span><span class="sxs-lookup"><span data-stu-id="406e6-598">Requirement</span></span>| <span data-ttu-id="406e6-599">Wert</span><span class="sxs-lookup"><span data-stu-id="406e6-599">Value</span></span>|
|---|---|
|[<span data-ttu-id="406e6-600">Mindestversion des Postfachanforderungssatzes</span><span class="sxs-lookup"><span data-stu-id="406e6-600">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="406e6-601">1.0</span><span class="sxs-lookup"><span data-stu-id="406e6-601">1.0</span></span>|
|[<span data-ttu-id="406e6-602">Mindestberechtigungsstufe</span><span class="sxs-lookup"><span data-stu-id="406e6-602">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="406e6-603">ReadItem</span><span class="sxs-lookup"><span data-stu-id="406e6-603">ReadItem</span></span>|
|[<span data-ttu-id="406e6-604">Zutreffender Outlook-Modus</span><span class="sxs-lookup"><span data-stu-id="406e6-604">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="406e6-605">Verfassen oder Lesen</span><span class="sxs-lookup"><span data-stu-id="406e6-605">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="406e6-606">Beispiel</span><span class="sxs-lookup"><span data-stu-id="406e6-606">Example</span></span>

```
Office.context.mailbox.item.to.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.to.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.to.getAsync(callback);

function callback(asyncResult) {
  var arrayOfToRecipients = asyncResult.value;
}
```

### <a name="methods"></a><span data-ttu-id="406e6-607">Methoden</span><span class="sxs-lookup"><span data-stu-id="406e6-607">Methods</span></span>

####  <a name="addfileattachmentasyncuri-attachmentname-options-callback"></a><span data-ttu-id="406e6-608">addFileAttachmentAsync(uri, attachmentName, [options], [callback])</span><span class="sxs-lookup"><span data-stu-id="406e6-608">addFileAttachmentAsync(uri, attachmentName, [options], [callback])</span></span>

<span data-ttu-id="406e6-609">Fügt eine Datei zu einer Nachricht oder einem Termin als Anlage hinzu.</span><span class="sxs-lookup"><span data-stu-id="406e6-609">Adds a file to a message or appointment as an attachment.</span></span>

<span data-ttu-id="406e6-610">Die `addFileAttachmentAsync`-Methode lädt die Datei am angegebenen URI hoch und fügt sie an das Element im Verfassenformular an.</span><span class="sxs-lookup"><span data-stu-id="406e6-610">The `addFileAttachmentAsync` method uploads the file at the specified URI and attaches it to the item in the compose form.</span></span>

<span data-ttu-id="406e6-611">Anschließend können Sie den Bezeichner mit der [`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback)-Methode in der gleichen Sitzung zum Entfernen der Anlage verwenden.</span><span class="sxs-lookup"><span data-stu-id="406e6-611">You can subsequently use the identifier with the [`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback) method to remove the attachment in the same session.</span></span>

##### <a name="parameters"></a><span data-ttu-id="406e6-612">Parameter:</span><span class="sxs-lookup"><span data-stu-id="406e6-612">Parameters:</span></span>

|<span data-ttu-id="406e6-613">Name</span><span class="sxs-lookup"><span data-stu-id="406e6-613">Name</span></span>| <span data-ttu-id="406e6-614">Typ</span><span class="sxs-lookup"><span data-stu-id="406e6-614">Type</span></span>| <span data-ttu-id="406e6-615">Attribute</span><span class="sxs-lookup"><span data-stu-id="406e6-615">Attributes</span></span>| <span data-ttu-id="406e6-616">Beschreibung</span><span class="sxs-lookup"><span data-stu-id="406e6-616">Description</span></span>|
|---|---|---|---|
|`uri`| <span data-ttu-id="406e6-617">Zeichenfolge</span><span class="sxs-lookup"><span data-stu-id="406e6-617">String</span></span>||<span data-ttu-id="406e6-p132">Der URI, der den Speicherort der an die Nachricht oder den Termin anzuhängenden Datei angibt. Die maximale Länge ist 2048 Zeichen.</span><span class="sxs-lookup"><span data-stu-id="406e6-p132">The URI that provides the location of the file to attach to the message or appointment. The maximum length is 2048 characters.</span></span>|
|`attachmentName`| <span data-ttu-id="406e6-620">Zeichenfolge</span><span class="sxs-lookup"><span data-stu-id="406e6-620">String</span></span>||<span data-ttu-id="406e6-p133">Der Name der Anlage, der beim Hochladen der Anlage angezeigt wird. Die maximale Länge ist 255 Zeichen.</span><span class="sxs-lookup"><span data-stu-id="406e6-p133">The name of the attachment that is shown while the attachment is uploading. The maximum length is 255 characters.</span></span>|
|`options`| <span data-ttu-id="406e6-623">Object</span><span class="sxs-lookup"><span data-stu-id="406e6-623">Object</span></span>| <span data-ttu-id="406e6-624">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="406e6-624">&lt;optional&gt;</span></span>|<span data-ttu-id="406e6-625">Ein Objektliteral, das mindestens eine der folgenden Eigenschaften enthält.</span><span class="sxs-lookup"><span data-stu-id="406e6-625">An object literal that contains one or more of the following properties.</span></span>|
| `options.asyncContext` | <span data-ttu-id="406e6-626">Object</span><span class="sxs-lookup"><span data-stu-id="406e6-626">Object</span></span> | <span data-ttu-id="406e6-627">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="406e6-627">&lt;optional&gt;</span></span> | <span data-ttu-id="406e6-628">Entwickler können jedes Objekt angeben, auf das sie in der Rückrufmethode zugreifen möchten.</span><span class="sxs-lookup"><span data-stu-id="406e6-628">Developers can provide any object they wish to access in the callback method.</span></span> |
| `options.isInline` | <span data-ttu-id="406e6-629">Boolean</span><span class="sxs-lookup"><span data-stu-id="406e6-629">Boolean</span></span> | <span data-ttu-id="406e6-630">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="406e6-630">&lt;optional&gt;</span></span> | <span data-ttu-id="406e6-631">Wenn `true`: Zeigt an, dass die Anlage inline im Nachrichtentext angezeigt wird und nicht in der Anlagenlisten angezeigt werden soll.</span><span class="sxs-lookup"><span data-stu-id="406e6-631">If `true`, indicates that the attachment will be shown inline in the message body, and should not be displayed in the attachment list.</span></span> |
|`callback`| <span data-ttu-id="406e6-632">function</span><span class="sxs-lookup"><span data-stu-id="406e6-632">function</span></span>| <span data-ttu-id="406e6-633">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="406e6-633">&lt;optional&gt;</span></span>|<span data-ttu-id="406e6-634">Wenn die Methode abgeschlossen wird, wird die Funktion , die im `callback`-Parameter übergeben wird, mit einem einzelnen Parameter (`asyncResult`) aufgerufen, der ein [`AsyncResult`](/javascript/api/office/office.asyncresult)-Objekt darstellt.</span><span class="sxs-lookup"><span data-stu-id="406e6-634">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span> <br/><span data-ttu-id="406e6-635">Bei Erfolg wird der Anlagenbezeichner in der `asyncResult.value`-Eigenschaft bereitgestellt.</span><span class="sxs-lookup"><span data-stu-id="406e6-635">On success, the attachment identifier will be provided in the `asyncResult.value` property.</span></span><br/><span data-ttu-id="406e6-636">Wenn beim Hochladen der Anlage ein Fehler auftritt, enthält das `asyncResult`-Objekt ein `Error`-Objekt mit einer Beschreibung des Fehlers.</span><span class="sxs-lookup"><span data-stu-id="406e6-636">If uploading the attachment fails, the `asyncResult` object will contain an `Error` object that provides a description of the error.</span></span>|

##### <a name="errors"></a><span data-ttu-id="406e6-637">Fehler</span><span class="sxs-lookup"><span data-stu-id="406e6-637">Errors</span></span>

| <span data-ttu-id="406e6-638">Fehlercode</span><span class="sxs-lookup"><span data-stu-id="406e6-638">Error code</span></span> | <span data-ttu-id="406e6-639">Beschreibung</span><span class="sxs-lookup"><span data-stu-id="406e6-639">Description</span></span> |
|------------|-------------|
| `AttachmentSizeExceeded` | <span data-ttu-id="406e6-640">Die Anlage ist zu groß.</span><span class="sxs-lookup"><span data-stu-id="406e6-640">The attachment is larger than allowed.</span></span> |
| `FileTypeNotSupported` | <span data-ttu-id="406e6-641">Die Anlage enthält eine Erweiterung, die nicht zulässig ist.</span><span class="sxs-lookup"><span data-stu-id="406e6-641">The attachment has an extension that is not allowed.</span></span> |
| `NumberOfAttachmentsExceeded` | <span data-ttu-id="406e6-642">Die Nachricht oder der Termin enthält zu viele Anlagen.</span><span class="sxs-lookup"><span data-stu-id="406e6-642">The message or appointment has too many attachments.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="406e6-643">Anforderungen</span><span class="sxs-lookup"><span data-stu-id="406e6-643">Requirements</span></span>

|<span data-ttu-id="406e6-644">Anforderung</span><span class="sxs-lookup"><span data-stu-id="406e6-644">Requirement</span></span>| <span data-ttu-id="406e6-645">Wert</span><span class="sxs-lookup"><span data-stu-id="406e6-645">Value</span></span>|
|---|---|
|[<span data-ttu-id="406e6-646">Mindestversion des Postfachanforderungssatzes</span><span class="sxs-lookup"><span data-stu-id="406e6-646">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="406e6-647">1.1</span><span class="sxs-lookup"><span data-stu-id="406e6-647">1.1</span></span>|
|[<span data-ttu-id="406e6-648">Mindestberechtigungsstufe</span><span class="sxs-lookup"><span data-stu-id="406e6-648">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="406e6-649">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="406e6-649">ReadWriteItem</span></span>|
|[<span data-ttu-id="406e6-650">Zutreffender Outlook-Modus</span><span class="sxs-lookup"><span data-stu-id="406e6-650">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="406e6-651">Verfassen</span><span class="sxs-lookup"><span data-stu-id="406e6-651">Compose</span></span>|

##### <a name="examples"></a><span data-ttu-id="406e6-652">Beispiele</span><span class="sxs-lookup"><span data-stu-id="406e6-652">Examples</span></span>

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

<span data-ttu-id="406e6-653">Im folgenden Beispiel wird eine Bilddatei als Inlineanlage hinzugefügt, und die Anlage wird im Nachrichtentext referenziert.</span><span class="sxs-lookup"><span data-stu-id="406e6-653">The following example adds an image file as an inline attachment and references the attachment in the message body.</span></span>

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

####  <a name="additemattachmentasyncitemid-attachmentname-options-callback"></a><span data-ttu-id="406e6-654">addItemAttachmentAsync(itemId, attachmentName, [options], [callback])</span><span class="sxs-lookup"><span data-stu-id="406e6-654">addItemAttachmentAsync(itemId, attachmentName, [options], [callback])</span></span>

<span data-ttu-id="406e6-655">Fügt der Nachricht oder dem Termin ein Exchange-Objekt, wie z. B. eine Nachricht, als Anhang hinzu.</span><span class="sxs-lookup"><span data-stu-id="406e6-655">Adds an Exchange item, such as a message, as an attachment to the message or appointment.</span></span>

<span data-ttu-id="406e6-p134">Die `addItemAttachmentAsync`-Methode fügt das Element mit dem angegebenen Exchange-Bezeichner an das Element im Formular zum Verfassen an. Wenn Sie eine Rückrufmethode angeben, wird die Methode mit einem `asyncResult`-Parameter aufgerufen, der entweder den Anlagenbezeichner oder einen Code enthält, der etwaige Fehler angibt, die beim Anfügen des Objekts aufgetreten sind. Sie können ggf. den `options`-Parameter verwenden, um Statusinformationen an die Rückrufmethode zu übergeben.</span><span class="sxs-lookup"><span data-stu-id="406e6-p134">The `addItemAttachmentAsync` method attaches the item with the specified Exchange identifier to the item in the compose form. If you specify a callback method, the method is called with one parameter, `asyncResult`, which contains either the attachment identifier or a code that indicates any error that occurred while attaching the item. You can use the `options` parameter to pass state information to the callback method, if needed.</span></span>

<span data-ttu-id="406e6-659">Anschließend können Sie den Bezeichner mit der [`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback)-Methode in der gleichen Sitzung zum Entfernen der Anlage verwenden.</span><span class="sxs-lookup"><span data-stu-id="406e6-659">You can subsequently use the identifier with the [`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback) method to remove the attachment in the same session.</span></span>

<span data-ttu-id="406e6-660">Wenn Ihre Office-Add-Ins in Outlook Web App ausgeführt wird die `addItemAttachmentAsync` -Methode kann Anfügen von Elementen in der Elemente, die Sie bearbeiten; Allerdings Dies wird nicht unterstützt und wird nicht empfohlen.</span><span class="sxs-lookup"><span data-stu-id="406e6-660">If your Office Add-in is running in Outlook Web App, the `addItemAttachmentAsync` method can attach items to items other than the item that you are editing; however, this is not supported and is not recommended.</span></span>

##### <a name="parameters"></a><span data-ttu-id="406e6-661">Parameter:</span><span class="sxs-lookup"><span data-stu-id="406e6-661">Parameters:</span></span>

|<span data-ttu-id="406e6-662">Name</span><span class="sxs-lookup"><span data-stu-id="406e6-662">Name</span></span>| <span data-ttu-id="406e6-663">Typ</span><span class="sxs-lookup"><span data-stu-id="406e6-663">Type</span></span>| <span data-ttu-id="406e6-664">Attribute</span><span class="sxs-lookup"><span data-stu-id="406e6-664">Attributes</span></span>| <span data-ttu-id="406e6-665">Beschreibung</span><span class="sxs-lookup"><span data-stu-id="406e6-665">Description</span></span>|
|---|---|---|---|
|`itemId`| <span data-ttu-id="406e6-666">Zeichenfolge</span><span class="sxs-lookup"><span data-stu-id="406e6-666">String</span></span>||<span data-ttu-id="406e6-p135">Der Exchange-Bezeichner des Objekts, das angehängt werden soll. Die maximale Länge beträgt 100 Zeichen.</span><span class="sxs-lookup"><span data-stu-id="406e6-p135">The Exchange identifier of the item to attach. The maximum length is 100 characters.</span></span>|
|`attachmentName`| <span data-ttu-id="406e6-669">Zeichenfolge</span><span class="sxs-lookup"><span data-stu-id="406e6-669">String</span></span>||<span data-ttu-id="406e6-p136">Der Betreff des Elements, das angehängt werden soll. Die maximale Länge ist 255 Zeichen.</span><span class="sxs-lookup"><span data-stu-id="406e6-p136">The sujbect of the item to be attached. The maximum length is 255 characters.</span></span>|
|`options`| <span data-ttu-id="406e6-672">Object</span><span class="sxs-lookup"><span data-stu-id="406e6-672">Object</span></span>| <span data-ttu-id="406e6-673">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="406e6-673">&lt;optional&gt;</span></span>|<span data-ttu-id="406e6-674">Ein Objektliteral, das mindestens eine der folgenden Eigenschaften enthält.</span><span class="sxs-lookup"><span data-stu-id="406e6-674">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`| <span data-ttu-id="406e6-675">Object</span><span class="sxs-lookup"><span data-stu-id="406e6-675">Object</span></span>| <span data-ttu-id="406e6-676">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="406e6-676">&lt;optional&gt;</span></span>|<span data-ttu-id="406e6-677">Entwickler können ein Objekt bereitstellen, auf das sie in der Callbackmethode zugreifen möchten.</span><span class="sxs-lookup"><span data-stu-id="406e6-677">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`| <span data-ttu-id="406e6-678">Funktion</span><span class="sxs-lookup"><span data-stu-id="406e6-678">function</span></span>| <span data-ttu-id="406e6-679">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="406e6-679">&lt;optional&gt;</span></span>|<span data-ttu-id="406e6-680">Wenn die Methode abgeschlossen wird, wird die Funktion , die im `callback`-Parameter übergeben wird, mit einem einzelnen Parameter (`asyncResult`) aufgerufen, der ein [`AsyncResult`](/javascript/api/office/office.asyncresult)-Objekt darstellt.</span><span class="sxs-lookup"><span data-stu-id="406e6-680">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span> <br/><span data-ttu-id="406e6-681">Bei Erfolg wird der Anlagenbezeichner in der `asyncResult.value`-Eigenschaft bereitgestellt.</span><span class="sxs-lookup"><span data-stu-id="406e6-681">On success, the attachment identifier will be provided in the `asyncResult.value` property.</span></span><br/><span data-ttu-id="406e6-682">Wenn beim Hinzufügen der Anlage ein Fehler auftritt, enthält das `asyncResult`-Objekt ein `Error`-Objekt mit einer Beschreibung des Fehlers.</span><span class="sxs-lookup"><span data-stu-id="406e6-682">If adding the attachment fails, the `asyncResult` object will contain an `Error` object that provides a description of the error.</span></span>|

##### <a name="errors"></a><span data-ttu-id="406e6-683">Fehler</span><span class="sxs-lookup"><span data-stu-id="406e6-683">Errors</span></span>

| <span data-ttu-id="406e6-684">Fehlercode</span><span class="sxs-lookup"><span data-stu-id="406e6-684">Error code</span></span> | <span data-ttu-id="406e6-685">Beschreibung</span><span class="sxs-lookup"><span data-stu-id="406e6-685">Description</span></span> |
|------------|-------------|
| `NumberOfAttachmentsExceeded` | <span data-ttu-id="406e6-686">Die Nachricht oder der Termin enthält zu viele Anlagen.</span><span class="sxs-lookup"><span data-stu-id="406e6-686">The message or appointment has too many attachments.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="406e6-687">Anforderungen</span><span class="sxs-lookup"><span data-stu-id="406e6-687">Requirements</span></span>

|<span data-ttu-id="406e6-688">Anforderung</span><span class="sxs-lookup"><span data-stu-id="406e6-688">Requirement</span></span>| <span data-ttu-id="406e6-689">Wert</span><span class="sxs-lookup"><span data-stu-id="406e6-689">Value</span></span>|
|---|---|
|[<span data-ttu-id="406e6-690">Mindestversion des Postfachanforderungssatzes</span><span class="sxs-lookup"><span data-stu-id="406e6-690">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="406e6-691">1.1</span><span class="sxs-lookup"><span data-stu-id="406e6-691">1.1</span></span>|
|[<span data-ttu-id="406e6-692">Mindestberechtigungsstufe</span><span class="sxs-lookup"><span data-stu-id="406e6-692">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="406e6-693">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="406e6-693">ReadWriteItem</span></span>|
|[<span data-ttu-id="406e6-694">Zutreffender Outlook-Modus</span><span class="sxs-lookup"><span data-stu-id="406e6-694">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="406e6-695">Verfassen</span><span class="sxs-lookup"><span data-stu-id="406e6-695">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="406e6-696">Beispiel</span><span class="sxs-lookup"><span data-stu-id="406e6-696">Example</span></span>

<span data-ttu-id="406e6-697">Im folgenden Beispiel wird ein vorhandenes Outlook-Element als Anlage mit dem Namen `My Attachment` hinzugefügt.</span><span class="sxs-lookup"><span data-stu-id="406e6-697">The following example adds an existing Outlook item as an attachment with the name `My Attachment`.</span></span>

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

####  <a name="close"></a><span data-ttu-id="406e6-698">close()</span><span class="sxs-lookup"><span data-stu-id="406e6-698">close()</span></span>

<span data-ttu-id="406e6-699">Schließt das aktuelle Element, das gerade verfasst wird.</span><span class="sxs-lookup"><span data-stu-id="406e6-699">Closes the current item that is being composed.</span></span>

<span data-ttu-id="406e6-p137">Das Verhalten der `close`-Methode hängt vom aktuellen Status des verfassten Elements ab. Wenn das Element über nicht gespeicherte Änderungen verfügt, fordert der Client den Benutzer auf, die Schließenaktion zu speichern, zu verwerfen oder abzubrechen.</span><span class="sxs-lookup"><span data-stu-id="406e6-p137">The behavior of the `close` method depends on the current state of the item being composed. If the item has unsaved changes, the client prompts the user to save, discard, or cancel the close action.</span></span>

> [!NOTE]
> <span data-ttu-id="406e6-702">In Outlook im Web, wenn das Element einen Termin handelt, und es wurde bereits zuvor gespeichert wurde mit `saveAsync`, der Benutzer wird aufgefordert, speichern, löschen oder Abbrechen, auch wenn keine Änderungen aufgetreten sind, da das Element zuletzt gespeichert.</span><span class="sxs-lookup"><span data-stu-id="406e6-702">In Outlook on the web, if the item is an appointment and it has previously been saved using `saveAsync`, the user is prompted to save, discard, or cancel even if no changes have occurred since the item was last saved.</span></span>

<span data-ttu-id="406e6-703">Wenn die Nachricht im Outlook-Desktopclient eine Inlineantwort ist, hat die `close`-Methode keine Auswirkung.</span><span class="sxs-lookup"><span data-stu-id="406e6-703">In the Outlook desktop client, if the message is an inline reply, the `close` method has no effect.</span></span>

##### <a name="requirements"></a><span data-ttu-id="406e6-704">Anforderungen</span><span class="sxs-lookup"><span data-stu-id="406e6-704">Requirements</span></span>

|<span data-ttu-id="406e6-705">Anforderung</span><span class="sxs-lookup"><span data-stu-id="406e6-705">Requirement</span></span>| <span data-ttu-id="406e6-706">Wert</span><span class="sxs-lookup"><span data-stu-id="406e6-706">Value</span></span>|
|---|---|
|[<span data-ttu-id="406e6-707">Mindestversion des Postfachanforderungssatzes</span><span class="sxs-lookup"><span data-stu-id="406e6-707">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="406e6-708">1.3</span><span class="sxs-lookup"><span data-stu-id="406e6-708">1.3</span></span>|
|[<span data-ttu-id="406e6-709">Mindestberechtigungsstufe</span><span class="sxs-lookup"><span data-stu-id="406e6-709">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="406e6-710">Eingeschränkt</span><span class="sxs-lookup"><span data-stu-id="406e6-710">Restricted</span></span>|
|[<span data-ttu-id="406e6-711">Zutreffender Outlook-Modus</span><span class="sxs-lookup"><span data-stu-id="406e6-711">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="406e6-712">Verfassen</span><span class="sxs-lookup"><span data-stu-id="406e6-712">Compose</span></span>|

#### <a name="displayreplyallformformdata"></a><span data-ttu-id="406e6-713">displayReplyAllForm(formData)</span><span class="sxs-lookup"><span data-stu-id="406e6-713">displayReplyAllForm(formData)</span></span>

<span data-ttu-id="406e6-714">Zeigt ein Antwortformular an, das den Absender und alle Empfänger der ausgewählten Nachricht oder des Organisators und alle Teilnehmer des ausgewählten Termins enthält.</span><span class="sxs-lookup"><span data-stu-id="406e6-714">Displays a reply form that includes the sender and all recipients of the selected message or the organizer and all attendees of the selected appointment.</span></span>

> [!NOTE]
> <span data-ttu-id="406e6-715">Diese Methode wird in Outlook für iOS oder Outlook für Android nicht unterstützt.</span><span class="sxs-lookup"><span data-stu-id="406e6-715">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="406e6-716">In Outlook Web App wird das Antwortformular als Popupformular in der Dreispaltenansicht und als Popupformular in der Zwei- oder Einspaltenansicht angezeigt.</span><span class="sxs-lookup"><span data-stu-id="406e6-716">In Outlook Web App, the reply form is displayed as a pop-out form in the 3-column view and a pop-up form in the 2- or 1-column view.</span></span>

<span data-ttu-id="406e6-717">Wenn einer der Zeichenfolgenparameter seinen Grenzwert überschreitet, löst `displayReplyAllForm` eine Ausnahme aus.</span><span class="sxs-lookup"><span data-stu-id="406e6-717">If any of the string parameters exceed their limits, `displayReplyAllForm` throws an exception.</span></span>

<span data-ttu-id="406e6-p138">Wenn Anlagen im `formData.attachments`-Parameter angegeben werden, versuchen Outlook und Outlook Web App alle Anlagen herunterzuladen und sie an das Antwortformular anzufügen. Wenn Anlagen nicht hinzugefügt werden können, wird in der Formularbenutzeroberfläche ein Fehler ausgegeben. Wenn dies nicht möglich ist, wird keine Fehlermeldung ausgegeben.</span><span class="sxs-lookup"><span data-stu-id="406e6-p138">When attachments are specified in the `formData.attachments` parameter, Outlook and Outlook Web App attempt to download all attachments and attach them to the reply form. If any attachments fail to be added, an error is shown in the form UI. If this isn't possible, then no error message is thrown.</span></span>

##### <a name="parameters"></a><span data-ttu-id="406e6-721">Parameter:</span><span class="sxs-lookup"><span data-stu-id="406e6-721">Parameters:</span></span>

| <span data-ttu-id="406e6-722">Name</span><span class="sxs-lookup"><span data-stu-id="406e6-722">Name</span></span> | <span data-ttu-id="406e6-723">Typ</span><span class="sxs-lookup"><span data-stu-id="406e6-723">Type</span></span> | <span data-ttu-id="406e6-724">Attribute</span><span class="sxs-lookup"><span data-stu-id="406e6-724">Attributes</span></span> | <span data-ttu-id="406e6-725">Beschreibung</span><span class="sxs-lookup"><span data-stu-id="406e6-725">Description</span></span> |
|---|---|---|---|
|`formData`| <span data-ttu-id="406e6-726">Zeichenfolge &#124; Objekt</span><span class="sxs-lookup"><span data-stu-id="406e6-726">String &#124; Object</span></span>| |<span data-ttu-id="406e6-p139">Eine Zeichenfolge, die Text- und HTML-Code enthält, die den Hauptteil des Antwortformulars darstellen. Die Zeichenfolge ist auf 32 KB beschränkt.</span><span class="sxs-lookup"><span data-stu-id="406e6-p139">A string that contains text and HTML and that represents the body of the reply form. The string is limited to 32 KB.</span></span><br/><span data-ttu-id="406e6-729">**ODER**</span><span class="sxs-lookup"><span data-stu-id="406e6-729">**OR**</span></span><br/><span data-ttu-id="406e6-p140">Ein Objekt, das Text- oder Anlagendaten und eine Rückruffunktion enthält. Das Objekt ist wie folgt definiert:</span><span class="sxs-lookup"><span data-stu-id="406e6-p140">An object that contains body or attachment data and a callback function. The object is defined as follows.</span></span> |
| `formData.htmlBody` | <span data-ttu-id="406e6-732">Zeichenfolge</span><span class="sxs-lookup"><span data-stu-id="406e6-732">String</span></span> | <span data-ttu-id="406e6-733">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="406e6-733">&lt;optional&gt;</span></span> | <span data-ttu-id="406e6-p141">Eine Zeichenfolge, die Text- und HTML-Code enthält, die den Hauptteil des Antwortformulars darstellen. Die Zeichenfolge ist auf 32 KB beschränkt.</span><span class="sxs-lookup"><span data-stu-id="406e6-p141">A string that contains text and HTML and that represents the body of the reply form. The string is limited to 32 KB.</span></span>
| `formData.attachments` | <span data-ttu-id="406e6-736">Array.&lt;Object&gt;</span><span class="sxs-lookup"><span data-stu-id="406e6-736">Array.&lt;Object&gt;</span></span> | <span data-ttu-id="406e6-737">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="406e6-737">&lt;optional&gt;</span></span> | <span data-ttu-id="406e6-738">Ein Array mit JSON-Objekten, die Datei- oder Elementanlagen sind.</span><span class="sxs-lookup"><span data-stu-id="406e6-738">An array of JSON objects that are either file or item attachments.</span></span> |
| `formData.attachments.type` | <span data-ttu-id="406e6-739">Zeichenfolge</span><span class="sxs-lookup"><span data-stu-id="406e6-739">String</span></span> | | <span data-ttu-id="406e6-p142">Gibt den Typ der Anlage an. Muss `file` für eine Dateianlage oder `item` für eine Elementanlage sein.</span><span class="sxs-lookup"><span data-stu-id="406e6-p142">Indicates the type of attachment. Must be `file` for a file attachment or `item` for an item attachment.</span></span> |
| `formData.attachments.name` | <span data-ttu-id="406e6-742">Zeichenfolge</span><span class="sxs-lookup"><span data-stu-id="406e6-742">String</span></span> | | <span data-ttu-id="406e6-743">Eine Zeichenfolge, die den Namen der Anlage mit bis zu 255 Zeichen enthält.</span><span class="sxs-lookup"><span data-stu-id="406e6-743">A string that contains the name of the attachment, up to 255 characters in length.</span></span>|
| `formData.attachments.url` | <span data-ttu-id="406e6-744">Zeichenfolge</span><span class="sxs-lookup"><span data-stu-id="406e6-744">String</span></span> | | <span data-ttu-id="406e6-p143">Wird nur verwendet, wenn `type` auf `file` gesetzt ist. Der URI des Speicherorts für die Datei.</span><span class="sxs-lookup"><span data-stu-id="406e6-p143">Only used if `type` is set to `file`. The URI of the location for the file.</span></span> |
| `formData.attachments.isInline` | <span data-ttu-id="406e6-747">Boolean</span><span class="sxs-lookup"><span data-stu-id="406e6-747">Boolean</span></span> | | <span data-ttu-id="406e6-p144">Wird nur verwendet, wenn `type` auf `file` gesetzt ist. Wenn `true`: Zeigt an, dass die Anlage inline im Nachrichtentext angezeigt wird und nicht in der Anlagenlisten angezeigt werden soll.</span><span class="sxs-lookup"><span data-stu-id="406e6-p144">Only used if `type` is set to `file`. If `true`, indicates that the attachment will be shown inline in the message body, and should not be displayed in the attachment list.</span></span> |
| `formData.attachments.itemId` | <span data-ttu-id="406e6-750">Zeichenfolge</span><span class="sxs-lookup"><span data-stu-id="406e6-750">String</span></span> | | <span data-ttu-id="406e6-p145">Wird nur verwendet, wenn `type` auf `item` gesetzt ist. Die EWS-Element-ID der Anlage. Dies ist eine Zeichenfolge bis zu 100 Zeichen.</span><span class="sxs-lookup"><span data-stu-id="406e6-p145">Only used if `type` is set to `item`. The EWS item id of the attachment. This is a string up to 100 characters.</span></span> |
| `callback` | <span data-ttu-id="406e6-754">Funktion</span><span class="sxs-lookup"><span data-stu-id="406e6-754">function</span></span> | <span data-ttu-id="406e6-755">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="406e6-755">&lt;optional&gt;</span></span> | <span data-ttu-id="406e6-756">Bei Abschluss der Methode wird die im Parameter `callback` übergebene Funktion mit einem einzigen Parameter aufgerufen: `asyncResult`. Dieser Parameter ist ein Objekt des Typs [AsyncResult](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="406e6-756">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [AsyncResult](/javascript/api/office/office.asyncresult) object.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="406e6-757">Anforderungen</span><span class="sxs-lookup"><span data-stu-id="406e6-757">Requirements</span></span>

|<span data-ttu-id="406e6-758">Anforderung</span><span class="sxs-lookup"><span data-stu-id="406e6-758">Requirement</span></span>| <span data-ttu-id="406e6-759">Wert</span><span class="sxs-lookup"><span data-stu-id="406e6-759">Value</span></span>|
|---|---|
|[<span data-ttu-id="406e6-760">Mindestversion des Postfachanforderungssatzes</span><span class="sxs-lookup"><span data-stu-id="406e6-760">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="406e6-761">1.0</span><span class="sxs-lookup"><span data-stu-id="406e6-761">1.0</span></span>|
|[<span data-ttu-id="406e6-762">Mindestberechtigungsstufe</span><span class="sxs-lookup"><span data-stu-id="406e6-762">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="406e6-763">ReadItem</span><span class="sxs-lookup"><span data-stu-id="406e6-763">ReadItem</span></span>|
|[<span data-ttu-id="406e6-764">Zutreffender Outlook-Modus</span><span class="sxs-lookup"><span data-stu-id="406e6-764">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="406e6-765">Lesen</span><span class="sxs-lookup"><span data-stu-id="406e6-765">Read</span></span>|

##### <a name="examples"></a><span data-ttu-id="406e6-766">Beispiele</span><span class="sxs-lookup"><span data-stu-id="406e6-766">Examples</span></span>

<span data-ttu-id="406e6-767">Im folgenden Code wird eine Zeichenfolge an die `displayReplyAllForm`-Funktion übergeben.</span><span class="sxs-lookup"><span data-stu-id="406e6-767">The following code passes a string to the `displayReplyAllForm` function.</span></span>

```
Office.context.mailbox.item.displayReplyAllForm('hello there');
Office.context.mailbox.item.displayReplyAllForm('<b>hello there</b>');
```

<span data-ttu-id="406e6-768">Antworten mit einem leeren Textkörper.</span><span class="sxs-lookup"><span data-stu-id="406e6-768">Reply with an empty body.</span></span>

```
Office.context.mailbox.item.displayReplyAllForm({});
```

<span data-ttu-id="406e6-769">Antworten nur einem Textkörper.</span><span class="sxs-lookup"><span data-stu-id="406e6-769">Reply with just a body.</span></span>

```
Office.context.mailbox.item.displayReplyAllForm(
{
  'htmlBody' : 'hi'
});
```

<span data-ttu-id="406e6-770">Antworten mit einem Textkörper und einer Dateianlage.</span><span class="sxs-lookup"><span data-stu-id="406e6-770">Reply with a body and a file attachment.</span></span>

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

<span data-ttu-id="406e6-771">Antworten mit einem Textkörper und einer Elementanlage.</span><span class="sxs-lookup"><span data-stu-id="406e6-771">Reply with a body and an item attachment.</span></span>

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

<span data-ttu-id="406e6-772">Antworten Sie mit einem Textkörper, einer Dateianlage, einer Elementanlage und einem Callback.</span><span class="sxs-lookup"><span data-stu-id="406e6-772">Reply with a body, file attachment, item attachment, and a callback.</span></span>

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

#### <a name="displayreplyformformdata"></a><span data-ttu-id="406e6-773">displayReplyForm(formData)</span><span class="sxs-lookup"><span data-stu-id="406e6-773">displayReplyForm(formData)</span></span>

<span data-ttu-id="406e6-774">Zeigt ein Antwortformular an, das nur den Absender der ausgewählten Nachricht oder den Organisator des ausgewählten Termins enthält.</span><span class="sxs-lookup"><span data-stu-id="406e6-774">Displays a reply form that includes only the sender of the selected message or the organizer of the selected appointment.</span></span>

> [!NOTE]
> <span data-ttu-id="406e6-775">Diese Methode wird in Outlook für iOS oder Outlook für Android nicht unterstützt.</span><span class="sxs-lookup"><span data-stu-id="406e6-775">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="406e6-776">In Outlook Web App wird das Antwortformular als Popupformular in der Dreispaltenansicht und als Popupformular in der Zwei- oder Einspaltenansicht angezeigt.</span><span class="sxs-lookup"><span data-stu-id="406e6-776">In Outlook Web App, the reply form is displayed as a pop-out form in the 3-column view and a pop-up form in the 2- or 1-column view.</span></span>

<span data-ttu-id="406e6-777">Wenn einer der Zeichenfolgenparameter seinen Grenzwert überschreitet, löst `displayReplyForm` eine Ausnahme aus.</span><span class="sxs-lookup"><span data-stu-id="406e6-777">If any of the string parameters exceed their limits, `displayReplyForm` throws an exception.</span></span>

<span data-ttu-id="406e6-p146">Wenn Anlagen im `formData.attachments`-Parameter angegeben werden, versuchen Outlook und Outlook Web App alle Anlagen herunterzuladen und sie an das Antwortformular anzufügen. Wenn Anlagen nicht hinzugefügt werden können, wird in der Formularbenutzeroberfläche ein Fehler ausgegeben. Wenn dies nicht möglich ist, wird keine Fehlermeldung ausgegeben.</span><span class="sxs-lookup"><span data-stu-id="406e6-p146">When attachments are specified in the `formData.attachments` parameter, Outlook and Outlook Web App attempt to download all attachments and attach them to the reply form. If any attachments fail to be added, an error is shown in the form UI. If this isn't possible, then no error message is thrown.</span></span>

##### <a name="parameters"></a><span data-ttu-id="406e6-781">Parameter:</span><span class="sxs-lookup"><span data-stu-id="406e6-781">Parameters:</span></span>

| <span data-ttu-id="406e6-782">Name</span><span class="sxs-lookup"><span data-stu-id="406e6-782">Name</span></span> | <span data-ttu-id="406e6-783">Typ</span><span class="sxs-lookup"><span data-stu-id="406e6-783">Type</span></span> | <span data-ttu-id="406e6-784">Attribute</span><span class="sxs-lookup"><span data-stu-id="406e6-784">Attributes</span></span> | <span data-ttu-id="406e6-785">Beschreibung</span><span class="sxs-lookup"><span data-stu-id="406e6-785">Description</span></span> |
|---|---|---|---|
|`formData`| <span data-ttu-id="406e6-786">Zeichenfolge &#124; Objekt</span><span class="sxs-lookup"><span data-stu-id="406e6-786">String &#124; Object</span></span>| | <span data-ttu-id="406e6-p147">Eine Zeichenfolge, die Text- und HTML-Code enthält, die den Hauptteil des Antwortformulars darstellen. Die Zeichenfolge ist auf 32 KB beschränkt.</span><span class="sxs-lookup"><span data-stu-id="406e6-p147">A string that contains text and HTML and that represents the body of the reply form. The string is limited to 32 KB.</span></span><br/><span data-ttu-id="406e6-789">**ODER**</span><span class="sxs-lookup"><span data-stu-id="406e6-789">**OR**</span></span><br/><span data-ttu-id="406e6-p148">Ein Objekt, das Text- oder Anlagendaten und eine Rückruffunktion enthält. Das Objekt ist wie folgt definiert:</span><span class="sxs-lookup"><span data-stu-id="406e6-p148">An object that contains body or attachment data and a callback function. The object is defined as follows.</span></span> |
| `formData.htmlBody` | <span data-ttu-id="406e6-792">Zeichenfolge</span><span class="sxs-lookup"><span data-stu-id="406e6-792">String</span></span> | <span data-ttu-id="406e6-793">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="406e6-793">&lt;optional&gt;</span></span> | <span data-ttu-id="406e6-p149">Eine Zeichenfolge, die Text- und HTML-Code enthält, die den Hauptteil des Antwortformulars darstellen. Die Zeichenfolge ist auf 32 KB beschränkt.</span><span class="sxs-lookup"><span data-stu-id="406e6-p149">A string that contains text and HTML and that represents the body of the reply form. The string is limited to 32 KB.</span></span>
| `formData.attachments` | <span data-ttu-id="406e6-796">Array.&lt;Object&gt;</span><span class="sxs-lookup"><span data-stu-id="406e6-796">Array.&lt;Object&gt;</span></span> | <span data-ttu-id="406e6-797">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="406e6-797">&lt;optional&gt;</span></span> | <span data-ttu-id="406e6-798">Ein Array mit JSON-Objekten, die Datei- oder Elementanlagen sind.</span><span class="sxs-lookup"><span data-stu-id="406e6-798">An array of JSON objects that are either file or item attachments.</span></span> |
| `formData.attachments.type` | <span data-ttu-id="406e6-799">Zeichenfolge</span><span class="sxs-lookup"><span data-stu-id="406e6-799">String</span></span> | | <span data-ttu-id="406e6-p150">Gibt den Typ der Anlage an. Muss `file` für eine Dateianlage oder `item` für eine Elementanlage sein.</span><span class="sxs-lookup"><span data-stu-id="406e6-p150">Indicates the type of attachment. Must be `file` for a file attachment or `item` for an item attachment.</span></span> |
| `formData.attachments.name` | <span data-ttu-id="406e6-802">Zeichenfolge</span><span class="sxs-lookup"><span data-stu-id="406e6-802">String</span></span> | | <span data-ttu-id="406e6-803">Eine Zeichenfolge, die den Namen der Anlage mit bis zu 255 Zeichen enthält.</span><span class="sxs-lookup"><span data-stu-id="406e6-803">A string that contains the name of the attachment, up to 255 characters in length.</span></span>|
| `formData.attachments.url` | <span data-ttu-id="406e6-804">Zeichenfolge</span><span class="sxs-lookup"><span data-stu-id="406e6-804">String</span></span> | | <span data-ttu-id="406e6-p151">Wird nur verwendet, wenn `type` auf `file` gesetzt ist. Der URI des Speicherorts für die Datei.</span><span class="sxs-lookup"><span data-stu-id="406e6-p151">Only used if `type` is set to `file`. The URI of the location for the file.</span></span> |
| `formData.attachments.isInline` | <span data-ttu-id="406e6-807">Boolean</span><span class="sxs-lookup"><span data-stu-id="406e6-807">Boolean</span></span> | | <span data-ttu-id="406e6-p152">Wird nur verwendet, wenn `type` auf `file` gesetzt ist. Wenn `true`: Zeigt an, dass die Anlage inline im Nachrichtentext angezeigt wird und nicht in der Anlagenlisten angezeigt werden soll.</span><span class="sxs-lookup"><span data-stu-id="406e6-p152">Only used if `type` is set to `file`. If `true`, indicates that the attachment will be shown inline in the message body, and should not be displayed in the attachment list.</span></span> |
| `formData.attachments.itemId` | <span data-ttu-id="406e6-810">Zeichenfolge</span><span class="sxs-lookup"><span data-stu-id="406e6-810">String</span></span> | | <span data-ttu-id="406e6-p153">Wird nur verwendet, wenn `type` auf `item` gesetzt ist. Die EWS-Element-ID der Anlage. Dies ist eine Zeichenfolge bis zu 100 Zeichen.</span><span class="sxs-lookup"><span data-stu-id="406e6-p153">Only used if `type` is set to `item`. The EWS item id of the attachment. This is a string up to 100 characters.</span></span> |
| `callback` | <span data-ttu-id="406e6-814">Funktion</span><span class="sxs-lookup"><span data-stu-id="406e6-814">function</span></span> | <span data-ttu-id="406e6-815">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="406e6-815">&lt;optional&gt;</span></span> | <span data-ttu-id="406e6-816">Bei Abschluss der Methode wird die im Parameter `callback` übergebene Funktion mit einem einzigen Parameter aufgerufen: `asyncResult`. Dieser Parameter ist ein Objekt des Typs [AsyncResult](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="406e6-816">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [AsyncResult](/javascript/api/office/office.asyncresult) object.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="406e6-817">Anforderungen</span><span class="sxs-lookup"><span data-stu-id="406e6-817">Requirements</span></span>

|<span data-ttu-id="406e6-818">Anforderung</span><span class="sxs-lookup"><span data-stu-id="406e6-818">Requirement</span></span>| <span data-ttu-id="406e6-819">Wert</span><span class="sxs-lookup"><span data-stu-id="406e6-819">Value</span></span>|
|---|---|
|[<span data-ttu-id="406e6-820">Mindestversion des Postfachanforderungssatzes</span><span class="sxs-lookup"><span data-stu-id="406e6-820">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="406e6-821">1.0</span><span class="sxs-lookup"><span data-stu-id="406e6-821">1.0</span></span>|
|[<span data-ttu-id="406e6-822">Mindestberechtigungsstufe</span><span class="sxs-lookup"><span data-stu-id="406e6-822">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="406e6-823">ReadItem</span><span class="sxs-lookup"><span data-stu-id="406e6-823">ReadItem</span></span>|
|[<span data-ttu-id="406e6-824">Zutreffender Outlook-Modus</span><span class="sxs-lookup"><span data-stu-id="406e6-824">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="406e6-825">Lesen</span><span class="sxs-lookup"><span data-stu-id="406e6-825">Read</span></span>|

##### <a name="examples"></a><span data-ttu-id="406e6-826">Beispiele</span><span class="sxs-lookup"><span data-stu-id="406e6-826">Examples</span></span>

<span data-ttu-id="406e6-827">Im folgenden Code wird eine Zeichenfolge an die `displayReplyForm`-Funktion übergeben.</span><span class="sxs-lookup"><span data-stu-id="406e6-827">The following code passes a string to the `displayReplyForm` function.</span></span>

```
Office.context.mailbox.item.displayReplyForm('hello there');
Office.context.mailbox.item.displayReplyForm('<b>hello there</b>');
```

<span data-ttu-id="406e6-828">Antworten mit einem leeren Textkörper.</span><span class="sxs-lookup"><span data-stu-id="406e6-828">Reply with an empty body.</span></span>

```
Office.context.mailbox.item.displayReplyForm({});
```

<span data-ttu-id="406e6-829">Antworten nur einem Textkörper.</span><span class="sxs-lookup"><span data-stu-id="406e6-829">Reply with just a body.</span></span>

```
Office.context.mailbox.item.displayReplyForm(
{
  'htmlBody' : 'hi'
});
```

<span data-ttu-id="406e6-830">Antworten mit einem Textkörper und einer Dateianlage.</span><span class="sxs-lookup"><span data-stu-id="406e6-830">Reply with a body and a file attachment.</span></span>

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

<span data-ttu-id="406e6-831">Antworten mit einem Textkörper und einer Elementanlage.</span><span class="sxs-lookup"><span data-stu-id="406e6-831">Reply with a body and an item attachment.</span></span>

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

<span data-ttu-id="406e6-832">Antworten Sie mit einem Textkörper, einer Dateianlage, einer Elementanlage und einem Callback.</span><span class="sxs-lookup"><span data-stu-id="406e6-832">Reply with a body, file attachment, item attachment, and a callback.</span></span>

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

#### <a name="getentities--entitiesjavascriptapioutlook15officeentities"></a><span data-ttu-id="406e6-833">getEntities() → {[Entities](/javascript/api/outlook_1_5/office.entities)}</span><span class="sxs-lookup"><span data-stu-id="406e6-833">getEntities() → {[Entities](/javascript/api/outlook_1_5/office.entities)}</span></span>

<span data-ttu-id="406e6-834">Ruft das ausgewählte Element Textkörper gefundenen Entitäten ab.</span><span class="sxs-lookup"><span data-stu-id="406e6-834">Gets the entities found in the selected item's body.</span></span>

> [!NOTE]
> <span data-ttu-id="406e6-835">Diese Methode wird in Outlook für iOS oder Outlook für Android nicht unterstützt.</span><span class="sxs-lookup"><span data-stu-id="406e6-835">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

##### <a name="requirements"></a><span data-ttu-id="406e6-836">Anforderungen</span><span class="sxs-lookup"><span data-stu-id="406e6-836">Requirements</span></span>

|<span data-ttu-id="406e6-837">Anforderung</span><span class="sxs-lookup"><span data-stu-id="406e6-837">Requirement</span></span>| <span data-ttu-id="406e6-838">Wert</span><span class="sxs-lookup"><span data-stu-id="406e6-838">Value</span></span>|
|---|---|
|[<span data-ttu-id="406e6-839">Mindestversion des Postfachanforderungssatzes</span><span class="sxs-lookup"><span data-stu-id="406e6-839">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="406e6-840">1.0</span><span class="sxs-lookup"><span data-stu-id="406e6-840">1.0</span></span>|
|[<span data-ttu-id="406e6-841">Mindestberechtigungsstufe</span><span class="sxs-lookup"><span data-stu-id="406e6-841">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="406e6-842">ReadItem</span><span class="sxs-lookup"><span data-stu-id="406e6-842">ReadItem</span></span>|
|[<span data-ttu-id="406e6-843">Zutreffender Outlook-Modus</span><span class="sxs-lookup"><span data-stu-id="406e6-843">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="406e6-844">Lesen</span><span class="sxs-lookup"><span data-stu-id="406e6-844">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="406e6-845">Gibt zurück:</span><span class="sxs-lookup"><span data-stu-id="406e6-845">Returns:</span></span>

<span data-ttu-id="406e6-846">Typ: [Entitäten](/javascript/api/outlook_1_5/office.entities)</span><span class="sxs-lookup"><span data-stu-id="406e6-846">Type: [Entities](/javascript/api/outlook_1_5/office.entities)</span></span>

##### <a name="example"></a><span data-ttu-id="406e6-847">Beispiel</span><span class="sxs-lookup"><span data-stu-id="406e6-847">Example</span></span>

<span data-ttu-id="406e6-848">Das folgende Beispiel greift auf die Kontakte Entitäten im Textkörper des aktuellen Elements.</span><span class="sxs-lookup"><span data-stu-id="406e6-848">The following example accesses the contacts entities in the current item's body.</span></span>

```
var contacts = Office.context.mailbox.item.getEntities().contacts;
```

#### <a name="getentitiesbytypeentitytype--nullable-arraystringcontactjavascriptapioutlook15officecontactmeetingsuggestionjavascriptapioutlook15officemeetingsuggestionphonenumberjavascriptapioutlook15officephonenumbertasksuggestionjavascriptapioutlook15officetasksuggestion"></a><span data-ttu-id="406e6-849">getEntitiesByType(entityType) → (nullable) {Array.<(String|[Contact](/javascript/api/outlook_1_5/office.contact)|[MeetingSuggestion](/javascript/api/outlook_1_5/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook_1_5/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook_1_5/office.tasksuggestion))>}</span><span class="sxs-lookup"><span data-stu-id="406e6-849">getEntitiesByType(entityType) → (nullable) {Array.<(String|[Contact](/javascript/api/outlook_1_5/office.contact)|[MeetingSuggestion](/javascript/api/outlook_1_5/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook_1_5/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook_1_5/office.tasksuggestion))>}</span></span>

<span data-ttu-id="406e6-850">Ruft ein Array aller Entitäten des angegebenen Entitätstyps im Hauptteil des ausgewählten Elements gefunden.</span><span class="sxs-lookup"><span data-stu-id="406e6-850">Gets an array of all the entities of the specified entity type found in the selected item's body.</span></span>

> [!NOTE]
> <span data-ttu-id="406e6-851">Diese Methode wird in Outlook für iOS oder Outlook für Android nicht unterstützt.</span><span class="sxs-lookup"><span data-stu-id="406e6-851">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

##### <a name="parameters"></a><span data-ttu-id="406e6-852">Parameter:</span><span class="sxs-lookup"><span data-stu-id="406e6-852">Parameters:</span></span>

|<span data-ttu-id="406e6-853">Name</span><span class="sxs-lookup"><span data-stu-id="406e6-853">Name</span></span>| <span data-ttu-id="406e6-854">Typ</span><span class="sxs-lookup"><span data-stu-id="406e6-854">Type</span></span>| <span data-ttu-id="406e6-855">Beschreibung</span><span class="sxs-lookup"><span data-stu-id="406e6-855">Description</span></span>|
|---|---|---|
|`entityType`| [<span data-ttu-id="406e6-856">Office.MailboxEnums.EntityType</span><span class="sxs-lookup"><span data-stu-id="406e6-856">Office.MailboxEnums.EntityType</span></span>](/javascript/api/outlook_1_5/office.mailboxenums.entitytype)|<span data-ttu-id="406e6-857">Einer der Werte der EntityType-Enumeration.</span><span class="sxs-lookup"><span data-stu-id="406e6-857">One of the EntityType enumeration values.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="406e6-858">Anforderungen</span><span class="sxs-lookup"><span data-stu-id="406e6-858">Requirements</span></span>

|<span data-ttu-id="406e6-859">Anforderung</span><span class="sxs-lookup"><span data-stu-id="406e6-859">Requirement</span></span>| <span data-ttu-id="406e6-860">Wert</span><span class="sxs-lookup"><span data-stu-id="406e6-860">Value</span></span>|
|---|---|
|[<span data-ttu-id="406e6-861">Mindestversion des Postfachanforderungssatzes</span><span class="sxs-lookup"><span data-stu-id="406e6-861">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="406e6-862">1.0</span><span class="sxs-lookup"><span data-stu-id="406e6-862">1.0</span></span>|
|[<span data-ttu-id="406e6-863">Mindestberechtigungsstufe</span><span class="sxs-lookup"><span data-stu-id="406e6-863">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="406e6-864">Eingeschränkt</span><span class="sxs-lookup"><span data-stu-id="406e6-864">Restricted</span></span>|
|[<span data-ttu-id="406e6-865">Zutreffender Outlook-Modus</span><span class="sxs-lookup"><span data-stu-id="406e6-865">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="406e6-866">Lesen</span><span class="sxs-lookup"><span data-stu-id="406e6-866">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="406e6-867">Gibt zurück:</span><span class="sxs-lookup"><span data-stu-id="406e6-867">Returns:</span></span>

<span data-ttu-id="406e6-868">Wenn der in `entityType` übergebene Wert kein gültiges Element der `EntityType`-Enumeration ist, gibt die Methode null zurück.</span><span class="sxs-lookup"><span data-stu-id="406e6-868">If the value passed in `entityType` is not a valid member of the `EntityType` enumeration, the method returns null.</span></span> <span data-ttu-id="406e6-869">Wenn keine Entitäten des angegebenen Typs im Textkörper des Elements vorhanden sind, gibt die Methode ein leeres Array zurück.</span><span class="sxs-lookup"><span data-stu-id="406e6-869">If no entities of the specified type are present in the item's body, the method returns an empty array.</span></span> <span data-ttu-id="406e6-870">Andernfalls hängt der Typ der Objekte im zurückgegebenen Array vom Typ der Entität ab, die im `entityType`-Parameter angefordert wurde.</span><span class="sxs-lookup"><span data-stu-id="406e6-870">Otherwise, the type of the objects in the returned array depends on the type of entity requested in the `entityType` parameter.</span></span>

<span data-ttu-id="406e6-871">Während Sie die minimale Berechtigungsstufe für diese Methode **Restricted** ist, erfordern einige Entitätstypen **ReadItem** für den Zugriff, wie in der folgenden Tabelle angegeben wird.</span><span class="sxs-lookup"><span data-stu-id="406e6-871">While the minimum permission level to use this method is **Restricted**, some entity types require **ReadItem** to access, as specified in the following table.</span></span>

| <span data-ttu-id="406e6-872">Wert von `entityType`</span><span class="sxs-lookup"><span data-stu-id="406e6-872">Value of `entityType`</span></span> | <span data-ttu-id="406e6-873">Typ der Objekte im zurückgegebenen Array</span><span class="sxs-lookup"><span data-stu-id="406e6-873">Type of objects in returned array</span></span> | <span data-ttu-id="406e6-874">Erforderliche Berechtigungsstufe</span><span class="sxs-lookup"><span data-stu-id="406e6-874">Required Permission Level</span></span> |
| --- | --- | --- |
| `Address` | <span data-ttu-id="406e6-875">Zeichenfolge</span><span class="sxs-lookup"><span data-stu-id="406e6-875">String</span></span> | <span data-ttu-id="406e6-876">**Restricted**</span><span class="sxs-lookup"><span data-stu-id="406e6-876">**Restricted**</span></span> |
| `Contact` | <span data-ttu-id="406e6-877">Contact</span><span class="sxs-lookup"><span data-stu-id="406e6-877">Contact</span></span> | <span data-ttu-id="406e6-878">**ReadItem**</span><span class="sxs-lookup"><span data-stu-id="406e6-878">**ReadItem**</span></span> |
| `EmailAddress` | <span data-ttu-id="406e6-879">Zeichenfolge</span><span class="sxs-lookup"><span data-stu-id="406e6-879">String</span></span> | <span data-ttu-id="406e6-880">**ReadItem**</span><span class="sxs-lookup"><span data-stu-id="406e6-880">**ReadItem**</span></span> |
| `MeetingSuggestion` | <span data-ttu-id="406e6-881">MeetingSuggestion</span><span class="sxs-lookup"><span data-stu-id="406e6-881">MeetingSuggestion</span></span> | <span data-ttu-id="406e6-882">**ReadItem**</span><span class="sxs-lookup"><span data-stu-id="406e6-882">**ReadItem**</span></span> |
| `PhoneNumber` | <span data-ttu-id="406e6-883">PhoneNumber</span><span class="sxs-lookup"><span data-stu-id="406e6-883">PhoneNumber</span></span> | <span data-ttu-id="406e6-884">**Restricted**</span><span class="sxs-lookup"><span data-stu-id="406e6-884">**Restricted**</span></span> |
| `TaskSuggestion` | <span data-ttu-id="406e6-885">TaskSuggestion</span><span class="sxs-lookup"><span data-stu-id="406e6-885">TaskSuggestion</span></span> | <span data-ttu-id="406e6-886">**ReadItem**</span><span class="sxs-lookup"><span data-stu-id="406e6-886">**ReadItem**</span></span> |
| `URL` | <span data-ttu-id="406e6-887">Zeichenfolge</span><span class="sxs-lookup"><span data-stu-id="406e6-887">String</span></span> | <span data-ttu-id="406e6-888">**Restricted**</span><span class="sxs-lookup"><span data-stu-id="406e6-888">**Restricted**</span></span> |

<span data-ttu-id="406e6-889">Typ: Array.<(String|[Contact](/javascript/api/outlook_1_5/office.contact)|[MeetingSuggestion](/javascript/api/outlook_1_5/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook_1_5/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook_1_5/office.tasksuggestion))></span><span class="sxs-lookup"><span data-stu-id="406e6-889">Type: Array.<(String|[Contact](/javascript/api/outlook_1_5/office.contact)|[MeetingSuggestion](/javascript/api/outlook_1_5/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook_1_5/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook_1_5/office.tasksuggestion))></span></span>

##### <a name="example"></a><span data-ttu-id="406e6-890">Beispiel</span><span class="sxs-lookup"><span data-stu-id="406e6-890">Example</span></span>

<span data-ttu-id="406e6-891">Im folgenden Beispiel wird veranschaulicht, wie auf ein Array von Zeichenfolgen, die Postadressen im Textkörper des aktuellen Elements darstellen.</span><span class="sxs-lookup"><span data-stu-id="406e6-891">The following example shows how to access an array of strings that represent postal addresses in the current item's body.</span></span>

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

#### <a name="getfilteredentitiesbynamename--nullable-arraystringcontactjavascriptapioutlook15officecontactmeetingsuggestionjavascriptapioutlook15officemeetingsuggestionphonenumberjavascriptapioutlook15officephonenumbertasksuggestionjavascriptapioutlook15officetasksuggestion"></a><span data-ttu-id="406e6-892">getFilteredEntitiesByName(name) → (nullable) {Array.<(String|[Contact](/javascript/api/outlook_1_5/office.contact)|[MeetingSuggestion](/javascript/api/outlook_1_5/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook_1_5/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook_1_5/office.tasksuggestion))>}</span><span class="sxs-lookup"><span data-stu-id="406e6-892">getFilteredEntitiesByName(name) → (nullable) {Array.<(String|[Contact](/javascript/api/outlook_1_5/office.contact)|[MeetingSuggestion](/javascript/api/outlook_1_5/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook_1_5/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook_1_5/office.tasksuggestion))>}</span></span>

<span data-ttu-id="406e6-893">Gibt bekannte Entitäten im ausgewählten Element zurück, die den in der XML-Manifestdatei definierten benannten Filter übergeben.</span><span class="sxs-lookup"><span data-stu-id="406e6-893">Returns well-known entities in the selected item that pass the named filter defined in the manifest XML file.</span></span>

> [!NOTE]
> <span data-ttu-id="406e6-894">Diese Methode wird in Outlook für iOS oder Outlook für Android nicht unterstützt.</span><span class="sxs-lookup"><span data-stu-id="406e6-894">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="406e6-895">Die `getFilteredEntitiesByName`-Methode gibt die Entitäten zurück, die dem im [ItemHasKnownEntity](/javascript/office/manifest/rule#itemhasknownentity-rule)-Regelelement der XML-Manifestdatei definierten regulären Ausdruck mit dem angegebenen `FilterName`-Elementwert entsprechen.</span><span class="sxs-lookup"><span data-stu-id="406e6-895">The `getFilteredEntitiesByName` method returns the entities that match the regular expression defined in the [ItemHasKnownEntity](/javascript/office/manifest/rule#itemhasknownentity-rule) rule element in the manifest XML file with the specified `FilterName` element value.</span></span>

##### <a name="parameters"></a><span data-ttu-id="406e6-896">Parameter:</span><span class="sxs-lookup"><span data-stu-id="406e6-896">Parameters:</span></span>

|<span data-ttu-id="406e6-897">Name</span><span class="sxs-lookup"><span data-stu-id="406e6-897">Name</span></span>| <span data-ttu-id="406e6-898">Typ</span><span class="sxs-lookup"><span data-stu-id="406e6-898">Type</span></span>| <span data-ttu-id="406e6-899">Beschreibung</span><span class="sxs-lookup"><span data-stu-id="406e6-899">Description</span></span>|
|---|---|---|
|`name`| <span data-ttu-id="406e6-900">Zeichenfolge</span><span class="sxs-lookup"><span data-stu-id="406e6-900">String</span></span>|<span data-ttu-id="406e6-901">Der Name des `ItemHasKnownEntity`-Regelelements, das den entsprechenden Filter definiert.</span><span class="sxs-lookup"><span data-stu-id="406e6-901">The name of the `ItemHasKnownEntity` rule element that defines the filter to match.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="406e6-902">Anforderungen</span><span class="sxs-lookup"><span data-stu-id="406e6-902">Requirements</span></span>

|<span data-ttu-id="406e6-903">Anforderung</span><span class="sxs-lookup"><span data-stu-id="406e6-903">Requirement</span></span>| <span data-ttu-id="406e6-904">Wert</span><span class="sxs-lookup"><span data-stu-id="406e6-904">Value</span></span>|
|---|---|
|[<span data-ttu-id="406e6-905">Mindestversion des Postfachanforderungssatzes</span><span class="sxs-lookup"><span data-stu-id="406e6-905">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="406e6-906">1.0</span><span class="sxs-lookup"><span data-stu-id="406e6-906">1.0</span></span>|
|[<span data-ttu-id="406e6-907">Mindestberechtigungsstufe</span><span class="sxs-lookup"><span data-stu-id="406e6-907">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="406e6-908">ReadItem</span><span class="sxs-lookup"><span data-stu-id="406e6-908">ReadItem</span></span>|
|[<span data-ttu-id="406e6-909">Zutreffender Outlook-Modus</span><span class="sxs-lookup"><span data-stu-id="406e6-909">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="406e6-910">Lesen</span><span class="sxs-lookup"><span data-stu-id="406e6-910">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="406e6-911">Gibt zurück:</span><span class="sxs-lookup"><span data-stu-id="406e6-911">Returns:</span></span>

<span data-ttu-id="406e6-p155">Befindet sich kein `ItemHasKnownEntity`-Element im Manifest mit einem `FilterName`-Elementwert, der dem `name`-Parameter entspricht, gibt die Methode `null` zurück. Wenn der `name`-Parameter einem `ItemHasKnownEntity`-Element im Manifest entspricht, es aber keine entsprechenden Entitäten im aktuellen Element gibt, gibt die Methode ein leeres Array zurück.</span><span class="sxs-lookup"><span data-stu-id="406e6-p155">If there is no `ItemHasKnownEntity` element in the manifest with a `FilterName` element value that matches the `name` parameter, the method returns `null`. If the `name` parameter does match an `ItemHasKnownEntity` element in the manifest, but there are no entities in the current item that match, the method return an empty array.</span></span>

<span data-ttu-id="406e6-914">Typ: Array.<(String|[Contact](/javascript/api/outlook_1_5/office.contact)|[MeetingSuggestion](/javascript/api/outlook_1_5/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook_1_5/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook_1_5/office.tasksuggestion))></span><span class="sxs-lookup"><span data-stu-id="406e6-914">Type: Array.<(String|[Contact](/javascript/api/outlook_1_5/office.contact)|[MeetingSuggestion](/javascript/api/outlook_1_5/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook_1_5/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook_1_5/office.tasksuggestion))></span></span>

#### <a name="getregexmatches--object"></a><span data-ttu-id="406e6-915">getRegExMatches() → {Object}</span><span class="sxs-lookup"><span data-stu-id="406e6-915">getRegExMatches() → {Object}</span></span>

<span data-ttu-id="406e6-916">Gibt Zeichenfolgenwerte im ausgewählten Element zurück, die den in der XML-Manifestdatei definierten regulären Ausdrücken entsprechen.</span><span class="sxs-lookup"><span data-stu-id="406e6-916">Returns string values in the selected item that match the regular expressions defined in the manifest XML file.</span></span>

> [!NOTE]
> <span data-ttu-id="406e6-917">Diese Methode wird in Outlook für iOS oder Outlook für Android nicht unterstützt.</span><span class="sxs-lookup"><span data-stu-id="406e6-917">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="406e6-p156">Die `getRegExMatches`-Methode gibt die Zeichenfolgen zurück, die dem im `ItemHasRegularExpressionMatch`- oder `ItemHasKnownEntity`-Regelelement der XML-Manifestdatei definierten regulären Ausdruck entsprechen. Bei einer `ItemHasRegularExpressionMatch`-Regel muss eine entsprechende Zeichenfolge in der Eigenschaft des Elements vorhanden sein, das von dieser Regel angegeben wird. Der einfache `PropertyName`-Typ definiert die unterstützten Eigenschaften.</span><span class="sxs-lookup"><span data-stu-id="406e6-p156">The `getRegExMatches` method returns the strings that match the regular expression defined in each `ItemHasRegularExpressionMatch` or `ItemHasKnownEntity` rule element in the manifest XML file. For an `ItemHasRegularExpressionMatch` rule, a matching string has to occur in the property of the item that is specified by that rule. The `PropertyName` simple type defines the supported properties.</span></span>

<span data-ttu-id="406e6-921">Nehmen Sie z. B. an, dass ein Add-In-Manifest das folgende `Rule`-Element aufweist:</span><span class="sxs-lookup"><span data-stu-id="406e6-921">For example, consider an add-in manifest has the following `Rule` element:</span></span>

```
<Rule xsi:type="RuleCollection" Mode="And">
  <Rule xsi:type="ItemIs" FormType="Read" ItemType="Message" />
  <Rule xsi:type="RuleCollection" Mode="Or">
    <Rule xsi:type="ItemHasRegularExpressionMatch" RegExName="fruits" RegExValue="apple|banana|coconut" PropertyName="BodyAsPlaintext" IgnoreCase="true" />
    <Rule xsi:type="ItemHasRegularExpressionMatch" RegExName="veggies" RegExValue="tomato|onion|spinach|broccoli" PropertyName="BodyAsPlaintext" IgnoreCase="true" />
  </Rule>
</Rule>
```

<span data-ttu-id="406e6-922">Das von `getRegExMatches` zurückgegebene Objekt hätte zwei Eigenschaften: `fruits` und `veggies`.</span><span class="sxs-lookup"><span data-stu-id="406e6-922">The object returned from `getRegExMatches` would have two properties: `fruits` and `veggies`.</span></span>

```
{
  'fruits': ['apple','banana','Banana','coconut'],
  'veggies': ['tomato','onion','spinach','broccoli']
}
```

<span data-ttu-id="406e6-p157">Wenn Sie eine `ItemHasRegularExpressionMatch`-Regel für die Textkörpereigenschaft eines Elements festlegen, sollte der reguläre Ausdruck den Textkörper weiter filtern und nicht versuchen, den gesamten Textkörper des Elements zurückzugeben. Wenn der gesamte Textkörper eines Elements mit einem regulären Ausdruck wie `.*` abgerufen wird, werden nicht immer die gewünschten Ergebnisse erzielt. Verwenden Sie stattdessen die [`Body.getAsync`](/javascript/api/outlook_1_5/office.body#getasync-coerciontype--options--callback-)-Methode, um den gesamten Textkörper abzurufen.</span><span class="sxs-lookup"><span data-stu-id="406e6-p157">If you specify an `ItemHasRegularExpressionMatch` rule on the body property of an item, the regular expression should further filter the body and should not attempt to return the entire body of the item. Using a regular expression such as `.*` to obtain the entire body of an item does not always return the expected results. Instead, use the [`Body.getAsync`](/javascript/api/outlook_1_5/office.body#getasync-coerciontype--options--callback-) method to retrieve the entire body.</span></span>

##### <a name="requirements"></a><span data-ttu-id="406e6-926">Anforderungen</span><span class="sxs-lookup"><span data-stu-id="406e6-926">Requirements</span></span>

|<span data-ttu-id="406e6-927">Anforderung</span><span class="sxs-lookup"><span data-stu-id="406e6-927">Requirement</span></span>| <span data-ttu-id="406e6-928">Wert</span><span class="sxs-lookup"><span data-stu-id="406e6-928">Value</span></span>|
|---|---|
|[<span data-ttu-id="406e6-929">Mindestversion des Postfachanforderungssatzes</span><span class="sxs-lookup"><span data-stu-id="406e6-929">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="406e6-930">1.0</span><span class="sxs-lookup"><span data-stu-id="406e6-930">1.0</span></span>|
|[<span data-ttu-id="406e6-931">Mindestberechtigungsstufe</span><span class="sxs-lookup"><span data-stu-id="406e6-931">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="406e6-932">ReadItem</span><span class="sxs-lookup"><span data-stu-id="406e6-932">ReadItem</span></span>|
|[<span data-ttu-id="406e6-933">Zutreffender Outlook-Modus</span><span class="sxs-lookup"><span data-stu-id="406e6-933">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="406e6-934">Lesen</span><span class="sxs-lookup"><span data-stu-id="406e6-934">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="406e6-935">Gibt zurück:</span><span class="sxs-lookup"><span data-stu-id="406e6-935">Returns:</span></span>

<span data-ttu-id="406e6-p158">Ein Objekt mit Arrays aus Zeichenfolgen, die den in der XML-Manifestdatei definierten regulären Ausdrücken entsprechen. Der Name der einzelnen Arrays ist gleich dem entsprechenden Wert des `RegExName`-Attributs der entsprechenden `ItemHasRegularExpressionMatch`-Regel oder des `FilterName`-Attributs der entsprechenden `ItemHasKnownEntity`-Regel.</span><span class="sxs-lookup"><span data-stu-id="406e6-p158">An object that contains arrays of strings that match the regular expressions defined in the manifest XML file. The name of each array is equal to the corresponding value of the `RegExName` attribute of the matching `ItemHasRegularExpressionMatch` rule or the `FilterName` attribute of the matching `ItemHasKnownEntity` rule.</span></span>

<dl class="param-type"><span data-ttu-id="406e6-938">

<dt>
Typ</dt>


</span><span class="sxs-lookup"><span data-stu-id="406e6-938">

<dt>Type</dt>

</span></span><dd><span data-ttu-id="406e6-939">Object</span><span class="sxs-lookup"><span data-stu-id="406e6-939">Object</span></span></dd>

</dl>

##### <a name="example"></a><span data-ttu-id="406e6-940">Beispiel</span><span class="sxs-lookup"><span data-stu-id="406e6-940">Example</span></span>

<span data-ttu-id="406e6-941">Das folgende Beispiel zeigt, wie Sie auf das Array von Übereinstimmungen für die regulären <rule>-Ausdruckselemente `fruits` und `veggies` zugreifen, die im Manifest angegeben sind.</rule></span><span class="sxs-lookup"><span data-stu-id="406e6-941">The following example shows how to access the array of matches for the regular expression <rule>elements `fruits` and `veggies`, which are specified in the manifest.</rule></span></span>

```
var allMatches = Office.context.mailbox.item.getRegExMatches();
var fruits = allMatches.fruits;
var veges = allMatches.veggies;
```

#### <a name="getregexmatchesbynamename--nullable-array-string-"></a><span data-ttu-id="406e6-942">getRegExMatchesByName(name) → (Nullwerte zulassen) {Array. < Zeichenfolge >}</span><span class="sxs-lookup"><span data-stu-id="406e6-942">getRegExMatchesByName(name) → (nullable) {Array.< String >}</span></span>

<span data-ttu-id="406e6-943">Gibt Zeichenfolgenwerte im ausgewählten Element zurück, die dem in der XML-Manifestdatei definierten benannten regulären Ausdruck entsprechen.</span><span class="sxs-lookup"><span data-stu-id="406e6-943">Returns string values in the selected item that match the named regular expression defined in the manifest XML file.</span></span>

> [!NOTE]
> <span data-ttu-id="406e6-944">Diese Methode wird in Outlook für iOS oder Outlook für Android nicht unterstützt.</span><span class="sxs-lookup"><span data-stu-id="406e6-944">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="406e6-945">Die `getRegExMatchesByName`-Methode gibt die Zeichenfolgen zurück, die dem im `ItemHasRegularExpressionMatch`-Regelelement der XML-Manifestdatei definierten regulären Ausdruck mit dem angegebenen `RegExName`-Elementwert entsprechen.</span><span class="sxs-lookup"><span data-stu-id="406e6-945">The `getRegExMatchesByName` method returns the strings that match the regular expression defined in the `ItemHasRegularExpressionMatch` rule element in the manifest XML file with the specified `RegExName` element value.</span></span>

<span data-ttu-id="406e6-p159">Wenn Sie eine `ItemHasRegularExpressionMatch`-Regel für die Textkörpereigenschaft eines Elements festlegen, sollte der reguläre Ausdruck den Textkörper weiter filtern und nicht versuchen, den gesamten Textkörper des Elements zurückzugeben. Wenn der gesamte Textkörper eines Elements mit einem regulären Ausdruck wie `.*` abgerufen wird, werden nicht immer die gewünschten Ergebnisse erzielt.</span><span class="sxs-lookup"><span data-stu-id="406e6-p159">If you specify an `ItemHasRegularExpressionMatch` rule on the body property of an item, the regular expression should further filter the body and should not attempt to return the entire body of the item. Using a regular expression such as `.*` to obtain the entire body of an item does not always return the expected results.</span></span>

##### <a name="parameters"></a><span data-ttu-id="406e6-948">Parameter:</span><span class="sxs-lookup"><span data-stu-id="406e6-948">Parameters:</span></span>

|<span data-ttu-id="406e6-949">Name</span><span class="sxs-lookup"><span data-stu-id="406e6-949">Name</span></span>| <span data-ttu-id="406e6-950">Typ</span><span class="sxs-lookup"><span data-stu-id="406e6-950">Type</span></span>| <span data-ttu-id="406e6-951">Beschreibung</span><span class="sxs-lookup"><span data-stu-id="406e6-951">Description</span></span>|
|---|---|---|
|`name`| <span data-ttu-id="406e6-952">Zeichenfolge</span><span class="sxs-lookup"><span data-stu-id="406e6-952">String</span></span>|<span data-ttu-id="406e6-953">Der Name des `ItemHasRegularExpressionMatch`-Regelelements, das den entsprechenden Filter definiert.</span><span class="sxs-lookup"><span data-stu-id="406e6-953">The name of the `ItemHasRegularExpressionMatch` rule element that defines the filter to match.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="406e6-954">Anforderungen</span><span class="sxs-lookup"><span data-stu-id="406e6-954">Requirements</span></span>

|<span data-ttu-id="406e6-955">Anforderung</span><span class="sxs-lookup"><span data-stu-id="406e6-955">Requirement</span></span>| <span data-ttu-id="406e6-956">Wert</span><span class="sxs-lookup"><span data-stu-id="406e6-956">Value</span></span>|
|---|---|
|[<span data-ttu-id="406e6-957">Mindestversion des Postfachanforderungssatzes</span><span class="sxs-lookup"><span data-stu-id="406e6-957">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="406e6-958">1.0</span><span class="sxs-lookup"><span data-stu-id="406e6-958">1.0</span></span>|
|[<span data-ttu-id="406e6-959">Mindestberechtigungsstufe</span><span class="sxs-lookup"><span data-stu-id="406e6-959">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="406e6-960">ReadItem</span><span class="sxs-lookup"><span data-stu-id="406e6-960">ReadItem</span></span>|
|[<span data-ttu-id="406e6-961">Zutreffender Outlook-Modus</span><span class="sxs-lookup"><span data-stu-id="406e6-961">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="406e6-962">Lesen</span><span class="sxs-lookup"><span data-stu-id="406e6-962">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="406e6-963">Gibt zurück:</span><span class="sxs-lookup"><span data-stu-id="406e6-963">Returns:</span></span>

<span data-ttu-id="406e6-964">Ein Array mit den Zeichenfolgen, die dem in der XML-Manifestdatei definierten regulären Ausdruck entsprechen.</span><span class="sxs-lookup"><span data-stu-id="406e6-964">An array that contains the strings that match the regular expression defined in the manifest XML file.</span></span>

<dl class="param-type"><span data-ttu-id="406e6-965">

<dt>Typ</dt>

</span><span class="sxs-lookup"><span data-stu-id="406e6-965">

<dt>Type</dt>

</span></span><dd><span data-ttu-id="406e6-966">Array. < Zeichenfolge ></span><span class="sxs-lookup"><span data-stu-id="406e6-966">Array.< String ></span></span></dd>

</dl>

##### <a name="example"></a><span data-ttu-id="406e6-967">Beispiel</span><span class="sxs-lookup"><span data-stu-id="406e6-967">Example</span></span>

```
var fruits = Office.context.mailbox.item.getRegExMatchesByName("fruits");
var veggies = Office.context.mailbox.item.getRegExMatchesByName("veggies");
```

####  <a name="getselecteddataasynccoerciontype-options-callback--string"></a><span data-ttu-id="406e6-968">getSelectedDataAsync(coercionType, [options], callback) → {String}</span><span class="sxs-lookup"><span data-stu-id="406e6-968">getSelectedDataAsync(coercionType, [options], callback) → {String}</span></span>

<span data-ttu-id="406e6-969">Gibt asynchron ausgewählte Daten aus dem Betreff oder Textkörper einer Nachricht zurück.</span><span class="sxs-lookup"><span data-stu-id="406e6-969">Asynchronously returns selected data from the subject or body of a message.</span></span>

<span data-ttu-id="406e6-p160">Wenn keine Auswahl vorhanden ist, aber der Cursor sich im Nachrichtentext oder Betreff befindet, gibt die Methode für die ausgewählten Daten NULL zurück. Wenn ein anderes Feld als der Textkörper oder Betreff ausgewählt ist, gibt die Methode den `InvalidSelection`-Fehler zurück.</span><span class="sxs-lookup"><span data-stu-id="406e6-p160">If there is no selection but the cursor is in the body or subject, the method returns null for the selected data. If a field other than the body or subject is selected, the method returns the `InvalidSelection` error.</span></span>

##### <a name="parameters"></a><span data-ttu-id="406e6-972">Parameter:</span><span class="sxs-lookup"><span data-stu-id="406e6-972">Parameters:</span></span>

|<span data-ttu-id="406e6-973">Name</span><span class="sxs-lookup"><span data-stu-id="406e6-973">Name</span></span>| <span data-ttu-id="406e6-974">Typ</span><span class="sxs-lookup"><span data-stu-id="406e6-974">Type</span></span>| <span data-ttu-id="406e6-975">Attribute</span><span class="sxs-lookup"><span data-stu-id="406e6-975">Attributes</span></span>| <span data-ttu-id="406e6-976">Beschreibung</span><span class="sxs-lookup"><span data-stu-id="406e6-976">Description</span></span>|
|---|---|---|---|
|`coercionType`| [<span data-ttu-id="406e6-977">Office.CoercionType</span><span class="sxs-lookup"><span data-stu-id="406e6-977">Office.CoercionType</span></span>](office.md#coerciontype-string)||<span data-ttu-id="406e6-p161">Fordert ein Format für die Daten an. Wenn es sich um Texthandelt, gibt die Methode den unformatierten Text als Zeichenfolge zurück und entfernt eventuell vorhandene HTML-Tags. Wenn es sich um HTML handelt, gibt die Methode den ausgewählten Text zurück, entweder als unformatierten Text oder als HTML.</span><span class="sxs-lookup"><span data-stu-id="406e6-p161">Requests a format for the data. If Text, the method returns the plain text as a string , removing any HTML tags present. If HTML, the method returns the selected text, whether it is plaintext or HTML.</span></span>|
|`options`| <span data-ttu-id="406e6-981">Object</span><span class="sxs-lookup"><span data-stu-id="406e6-981">Object</span></span>| <span data-ttu-id="406e6-982">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="406e6-982">&lt;optional&gt;</span></span>|<span data-ttu-id="406e6-983">Ein Objektliteral, das mindestens eine der folgenden Eigenschaften enthält.</span><span class="sxs-lookup"><span data-stu-id="406e6-983">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`| <span data-ttu-id="406e6-984">Object</span><span class="sxs-lookup"><span data-stu-id="406e6-984">Object</span></span>| <span data-ttu-id="406e6-985">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="406e6-985">&lt;optional&gt;</span></span>|<span data-ttu-id="406e6-986">Entwickler können ein Objekt bereitstellen, auf das sie in der Callbackmethode zugreifen möchten.</span><span class="sxs-lookup"><span data-stu-id="406e6-986">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`| <span data-ttu-id="406e6-987">Funktion</span><span class="sxs-lookup"><span data-stu-id="406e6-987">function</span></span>||<span data-ttu-id="406e6-988">Wenn die Methode abgeschlossen wird, wird die Funktion , die im `callback`-Parameter übergeben wird, mit einem einzelnen Parameter (`asyncResult`) aufgerufen, der ein [`AsyncResult`](/javascript/api/office/office.asyncresult)-Objekt darstellt.</span><span class="sxs-lookup"><span data-stu-id="406e6-988">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span><br/><br/><span data-ttu-id="406e6-989">Rufen Sie für den Zugriff auf die ausgewählten Daten aus der Rückrufmethode `asyncResult.value.data` auf.</span><span class="sxs-lookup"><span data-stu-id="406e6-989">To access the selected data from the callback method, call `asyncResult.value.data`.</span></span> <span data-ttu-id="406e6-990">Rufen Sie die Source-Eigenschaft für den Zugriff, die die Auswahl stammen, `asyncResult.value.sourceProperty`, die entweder sein `body` oder `subject`.</span><span class="sxs-lookup"><span data-stu-id="406e6-990">To access the source property that the selection comes from, call `asyncResult.value.sourceProperty`, which will be either `body` or `subject`.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="406e6-991">Anforderungen</span><span class="sxs-lookup"><span data-stu-id="406e6-991">Requirements</span></span>

|<span data-ttu-id="406e6-992">Anforderung</span><span class="sxs-lookup"><span data-stu-id="406e6-992">Requirement</span></span>| <span data-ttu-id="406e6-993">Wert</span><span class="sxs-lookup"><span data-stu-id="406e6-993">Value</span></span>|
|---|---|
|[<span data-ttu-id="406e6-994">Mindestversion des Postfachanforderungssatzes</span><span class="sxs-lookup"><span data-stu-id="406e6-994">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="406e6-995">1.2</span><span class="sxs-lookup"><span data-stu-id="406e6-995">1.2</span></span>|
|[<span data-ttu-id="406e6-996">Mindestberechtigungsstufe</span><span class="sxs-lookup"><span data-stu-id="406e6-996">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="406e6-997">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="406e6-997">ReadWriteItem</span></span>|
|[<span data-ttu-id="406e6-998">Zutreffender Outlook-Modus</span><span class="sxs-lookup"><span data-stu-id="406e6-998">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="406e6-999">Verfassen</span><span class="sxs-lookup"><span data-stu-id="406e6-999">Compose</span></span>|

##### <a name="returns"></a><span data-ttu-id="406e6-1000">Gibt zurück:</span><span class="sxs-lookup"><span data-stu-id="406e6-1000">Returns:</span></span>

<span data-ttu-id="406e6-1001">Die ausgewählten Daten als Zeichenfolge mit dem durch `coercionType` bestimmten Format.</span><span class="sxs-lookup"><span data-stu-id="406e6-1001">The selected data as a string with format determined by `coercionType`.</span></span>

<dl class="param-type"><span data-ttu-id="406e6-1002">

<dt>
Typ</dt>


</span><span class="sxs-lookup"><span data-stu-id="406e6-1002">

<dt>Type</dt>

</span></span><dd><span data-ttu-id="406e6-1003">Zeichenfolge</span><span class="sxs-lookup"><span data-stu-id="406e6-1003">String</span></span></dd>

</dl>

##### <a name="example"></a><span data-ttu-id="406e6-1004">Beispiel</span><span class="sxs-lookup"><span data-stu-id="406e6-1004">Example</span></span>

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

####  <a name="loadcustompropertiesasynccallback-usercontext"></a><span data-ttu-id="406e6-1005">loadCustomPropertiesAsync(callback, [userContext])</span><span class="sxs-lookup"><span data-stu-id="406e6-1005">loadCustomPropertiesAsync(callback, [userContext])</span></span>

<span data-ttu-id="406e6-1006">Lädt asynchron benutzerdefinierte Eigenschaften für dieses Add-In für das ausgewählte Element.</span><span class="sxs-lookup"><span data-stu-id="406e6-1006">Asynchronously loads custom properties for this add-in on the selected item.</span></span>

<span data-ttu-id="406e6-p163">Benutzerdefinierte Eigenschaften werden als Schlüssel-/Wert-Paare pro App und pro Element gespeichert. Diese Methode gibt ein `CustomProperties`-Objekt im Rückruf zurück, das Methoden für den Zugriff auf die benutzerdefinierten Eigenschaften für das aktuelle Element und das aktuelle Add-In bietet. Benutzerdefinierte Eigenschaften sind für das Element nicht verschlüsselt und sollten darum nicht als sicherer Speicher verwendet werden.</span><span class="sxs-lookup"><span data-stu-id="406e6-p163">Custom properties are stored as key/value pairs on a per-app, per-item basis. This method returns a `CustomProperties` object in the callback, which provides methods to access the custom properties specific to the current item and the current add-in. Custom properties are not encrypted on the item, so this should not be used as secure storage.</span></span>

##### <a name="parameters"></a><span data-ttu-id="406e6-1010">Parameter:</span><span class="sxs-lookup"><span data-stu-id="406e6-1010">Parameters:</span></span>

|<span data-ttu-id="406e6-1011">Name</span><span class="sxs-lookup"><span data-stu-id="406e6-1011">Name</span></span>| <span data-ttu-id="406e6-1012">Typ</span><span class="sxs-lookup"><span data-stu-id="406e6-1012">Type</span></span>| <span data-ttu-id="406e6-1013">Attribute</span><span class="sxs-lookup"><span data-stu-id="406e6-1013">Attributes</span></span>| <span data-ttu-id="406e6-1014">Beschreibung</span><span class="sxs-lookup"><span data-stu-id="406e6-1014">Description</span></span>|
|---|---|---|---|
|`callback`| <span data-ttu-id="406e6-1015">Funktion</span><span class="sxs-lookup"><span data-stu-id="406e6-1015">function</span></span>||<span data-ttu-id="406e6-1016">Wenn die Methode abgeschlossen wird, wird die Funktion , die im `callback`-Parameter übergeben wird, mit einem einzelnen Parameter (`asyncResult`) aufgerufen, der ein [`AsyncResult`](/javascript/api/office/office.asyncresult)-Objekt darstellt.</span><span class="sxs-lookup"><span data-stu-id="406e6-1016">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span><br/><br/><span data-ttu-id="406e6-1017">Die benutzerdefinierten Eigenschaften werden als [`CustomProperties`](/javascript/api/outlook_1_5/office.customproperties)-Objekt in der `asyncResult.value`-Eigenschaft bereitgestellt.</span><span class="sxs-lookup"><span data-stu-id="406e6-1017">The custom properties are provided as a [`CustomProperties`](/javascript/api/outlook_1_5/office.customproperties) object in the `asyncResult.value` property.</span></span> <span data-ttu-id="406e6-1018">Dieses Objekt kann zum Abrufen, festlegen und Entfernen benutzerdefinierter Eigenschaften aus dem Element und speichern Sie Änderungen an den benutzerdefinierten Eigenschaftensatz zurück an den Server verwendet werden.</span><span class="sxs-lookup"><span data-stu-id="406e6-1018">This object can be used to get, set, and remove custom properties from the item and save changes to the custom property set back to the server.</span></span>|
|`userContext`| <span data-ttu-id="406e6-1019">Objekt</span><span class="sxs-lookup"><span data-stu-id="406e6-1019">Object</span></span>| <span data-ttu-id="406e6-1020">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="406e6-1020">&lt;optional&gt;</span></span>|<span data-ttu-id="406e6-1021">Entwickler können ein beliebiges Objekt bereitstellen, den sie in der Rückruffunktion zugreifen möchten.</span><span class="sxs-lookup"><span data-stu-id="406e6-1021">Developers can provide any object they wish to access in the callback function.</span></span> <span data-ttu-id="406e6-1022">Dieses Objekt zugegriffen werden kann, indem die `asyncResult.asyncContext` -Eigenschaft in der Callback-Funktion.</span><span class="sxs-lookup"><span data-stu-id="406e6-1022">This object can be accessed by the `asyncResult.asyncContext` property in the callback function.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="406e6-1023">Anforderungen</span><span class="sxs-lookup"><span data-stu-id="406e6-1023">Requirements</span></span>

|<span data-ttu-id="406e6-1024">Anforderung</span><span class="sxs-lookup"><span data-stu-id="406e6-1024">Requirement</span></span>| <span data-ttu-id="406e6-1025">Wert</span><span class="sxs-lookup"><span data-stu-id="406e6-1025">Value</span></span>|
|---|---|
|[<span data-ttu-id="406e6-1026">Mindestversion des Postfachanforderungssatzes</span><span class="sxs-lookup"><span data-stu-id="406e6-1026">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="406e6-1027">1.0</span><span class="sxs-lookup"><span data-stu-id="406e6-1027">1.0</span></span>|
|[<span data-ttu-id="406e6-1028">Mindestberechtigungsstufe</span><span class="sxs-lookup"><span data-stu-id="406e6-1028">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="406e6-1029">ReadItem</span><span class="sxs-lookup"><span data-stu-id="406e6-1029">ReadItem</span></span>|
|[<span data-ttu-id="406e6-1030">Zutreffender Outlook-Modus</span><span class="sxs-lookup"><span data-stu-id="406e6-1030">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="406e6-1031">Verfassen oder Lesen</span><span class="sxs-lookup"><span data-stu-id="406e6-1031">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="406e6-1032">Beispiel</span><span class="sxs-lookup"><span data-stu-id="406e6-1032">Example</span></span>

<span data-ttu-id="406e6-p166">Das folgende Codebeispiel veranschaulicht die Verwendung der `loadCustomPropertiesAsync`-Methode zum asynchronen Laden der benutzerdefinierten Eigenschaften, die für das aktuelle Element spezifisch sind. Des Weiteren veranschaulicht das Beispiel die Verwendung der `CustomProperties.saveAsync`-Methode zum Speichern der Eigenschaften auf dem Server. Nach dem Laden der benutzerdefinierten Eigenschaften wird in dem Codebeispiel die `CustomProperties.get`-Methode dazu verwendet, die benutzerdefinierte `myProp`-Eigenschaft zu lesen. Die `CustomProperties.set`-Methode wird dazu verwendet, die benutzerdefinierte `otherProp`-Eigenschaft zu schreiben. Schließlich wird die `saveAsync`-Methode aufgerufen, um die benutzerdefinierten Eigenschaften zu speichern.</span><span class="sxs-lookup"><span data-stu-id="406e6-p166">The following code example shows how to use the `loadCustomPropertiesAsync` method to asynchronously load custom properties that are specific to the current item. The example also shows how to use the `CustomProperties.saveAsync` method to save these properties back to the server. After loading the custom properties, the code sample uses the `CustomProperties.get` method to read the custom property `myProp`, the `CustomProperties.set` method to write the custom property `otherProp`, and then finally calls the `saveAsync` method to save the custom properties.</span></span>

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

####  <a name="removeattachmentasyncattachmentid-options-callback"></a><span data-ttu-id="406e6-1036">removeAttachmentAsync(attachmentId, [options], [callback])</span><span class="sxs-lookup"><span data-stu-id="406e6-1036">removeAttachmentAsync(attachmentId, [options], [callback])</span></span>

<span data-ttu-id="406e6-1037">Entfernt eine Anlage aus einer Nachricht oder einem Termin.</span><span class="sxs-lookup"><span data-stu-id="406e6-1037">Removes an attachment from a message or appointment.</span></span>

<span data-ttu-id="406e6-p167">Die `removeAttachmentAsync`-Methode entfernt die Anlage mit dem angegebenen Bezeichner aus dem Element. Als bewährte Vorgehensweise sollten Sie den Anlagenbezeichner nur dann zum Entfernen einer Anlage verwenden, wenn die gleiche Mail-App die Anlage in der gleichen Sitzung hinzugefügt hat. In Outlook Web App und OWA für Geräte ist der Anlagenbezeichner nur innerhalb der gleichen Sitzung gültig. Eine Sitzung ist abgeschlossen, wenn der Benutzer die App schließt, oder wenn der Benutzer in einem eingebetteten Formular mit dem Verfassen beginnt und den Vorgang dann in einem separaten Fenster fortsetzt.</span><span class="sxs-lookup"><span data-stu-id="406e6-p167">The `removeAttachmentAsync` method removes the attachment with the specified identifier from the item. As a best practice, you should use the attachment identifier to remove an attachment only if the same mail app has added that attachment in the same session. In Outlook Web App and OWA for Devices, the attachment identifier is valid only within the same session. A session is over when the user closes the app, or if the user starts composing in an inline form and subsequently pops out the inline form to continue in a separate window.</span></span>

##### <a name="parameters"></a><span data-ttu-id="406e6-1042">Parameter:</span><span class="sxs-lookup"><span data-stu-id="406e6-1042">Parameters:</span></span>

|<span data-ttu-id="406e6-1043">Name</span><span class="sxs-lookup"><span data-stu-id="406e6-1043">Name</span></span>| <span data-ttu-id="406e6-1044">Typ</span><span class="sxs-lookup"><span data-stu-id="406e6-1044">Type</span></span>| <span data-ttu-id="406e6-1045">Attribute</span><span class="sxs-lookup"><span data-stu-id="406e6-1045">Attributes</span></span>| <span data-ttu-id="406e6-1046">Beschreibung</span><span class="sxs-lookup"><span data-stu-id="406e6-1046">Description</span></span>|
|---|---|---|---|
|`attachmentId`| <span data-ttu-id="406e6-1047">Zeichenfolge</span><span class="sxs-lookup"><span data-stu-id="406e6-1047">String</span></span>||<span data-ttu-id="406e6-p168">Der Bezeichner des Anhangs, der entfernt werden soll. Die maximale Länge der Zeichenfolge ist 100 Zeichen.</span><span class="sxs-lookup"><span data-stu-id="406e6-p168">The identifier of the attachment to remove. The maximum length of the string is 100 characters.</span></span>|
|`options`| <span data-ttu-id="406e6-1050">Object</span><span class="sxs-lookup"><span data-stu-id="406e6-1050">Object</span></span>| <span data-ttu-id="406e6-1051">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="406e6-1051">&lt;optional&gt;</span></span>|<span data-ttu-id="406e6-1052">Ein Objektliteral, das mindestens eine der folgenden Eigenschaften enthält.</span><span class="sxs-lookup"><span data-stu-id="406e6-1052">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`| <span data-ttu-id="406e6-1053">Object</span><span class="sxs-lookup"><span data-stu-id="406e6-1053">Object</span></span>| <span data-ttu-id="406e6-1054">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="406e6-1054">&lt;optional&gt;</span></span>|<span data-ttu-id="406e6-1055">Entwickler können ein Objekt bereitstellen, auf das sie in der Callbackmethode zugreifen möchten.</span><span class="sxs-lookup"><span data-stu-id="406e6-1055">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`| <span data-ttu-id="406e6-1056">Funktion</span><span class="sxs-lookup"><span data-stu-id="406e6-1056">function</span></span>| <span data-ttu-id="406e6-1057">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="406e6-1057">&lt;optional&gt;</span></span>|<span data-ttu-id="406e6-1058">Wenn die Methode abgeschlossen wird, wird die Funktion , die im `callback`-Parameter übergeben wird, mit einem einzelnen Parameter (`asyncResult`) aufgerufen, der ein [`AsyncResult`](/javascript/api/office/office.asyncresult)-Objekt darstellt.</span><span class="sxs-lookup"><span data-stu-id="406e6-1058">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span> <br/><span data-ttu-id="406e6-1059">Wenn beim Entfernen der Anlage ein Fehler auftritt, enthält die Eigenschaft `asyncResult.error` einen Fehlercode mit dem Grund für den Fehler.</span><span class="sxs-lookup"><span data-stu-id="406e6-1059">If removing the attachment fails, the `asyncResult.error` property will contain an error code with the reason for the failure.</span></span>|

##### <a name="errors"></a><span data-ttu-id="406e6-1060">Fehler</span><span class="sxs-lookup"><span data-stu-id="406e6-1060">Errors</span></span>

| <span data-ttu-id="406e6-1061">Fehlercode</span><span class="sxs-lookup"><span data-stu-id="406e6-1061">Error code</span></span> | <span data-ttu-id="406e6-1062">Beschreibung</span><span class="sxs-lookup"><span data-stu-id="406e6-1062">Description</span></span> |
|------------|-------------|
| `InvalidAttachmentId` | <span data-ttu-id="406e6-1063">Der Bezeichner für die Anlage ist nicht vorhanden.</span><span class="sxs-lookup"><span data-stu-id="406e6-1063">The attachment identifier does not exist.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="406e6-1064">Anforderungen</span><span class="sxs-lookup"><span data-stu-id="406e6-1064">Requirements</span></span>

|<span data-ttu-id="406e6-1065">Anforderung</span><span class="sxs-lookup"><span data-stu-id="406e6-1065">Requirement</span></span>| <span data-ttu-id="406e6-1066">Wert</span><span class="sxs-lookup"><span data-stu-id="406e6-1066">Value</span></span>|
|---|---|
|[<span data-ttu-id="406e6-1067">Mindestversion des Postfachanforderungssatzes</span><span class="sxs-lookup"><span data-stu-id="406e6-1067">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="406e6-1068">1.1</span><span class="sxs-lookup"><span data-stu-id="406e6-1068">1.1</span></span>|
|[<span data-ttu-id="406e6-1069">Mindestberechtigungsstufe</span><span class="sxs-lookup"><span data-stu-id="406e6-1069">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="406e6-1070">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="406e6-1070">ReadWriteItem</span></span>|
|[<span data-ttu-id="406e6-1071">Zutreffender Outlook-Modus</span><span class="sxs-lookup"><span data-stu-id="406e6-1071">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="406e6-1072">Verfassen</span><span class="sxs-lookup"><span data-stu-id="406e6-1072">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="406e6-1073">Beispiel</span><span class="sxs-lookup"><span data-stu-id="406e6-1073">Example</span></span>

<span data-ttu-id="406e6-1074">Mit dem folgende Code wird eine Anlage mit dem Bezeichner "0" entfernt.</span><span class="sxs-lookup"><span data-stu-id="406e6-1074">The following code removes an attachment with an identifier of '0'.</span></span>

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

####  <a name="saveasyncoptions-callback"></a><span data-ttu-id="406e6-1075">saveAsync([options], callback)</span><span class="sxs-lookup"><span data-stu-id="406e6-1075">saveAsync([options], callback)</span></span>

<span data-ttu-id="406e6-1076">Speicher asynchron ein Element. .</span><span class="sxs-lookup"><span data-stu-id="406e6-1076">Asynchronously saves an item.</span></span>

<span data-ttu-id="406e6-p169">Beim Aufrufen speichert diese Methode die aktuelle Nachricht als Entwurf und  gibt die Element-ID über die Callbackmethode zurück. In Outlook Web App oder Outlook im Onlinemodus wird das Element auf dem Server gespeichert. In Outlook im Cache-Modus wird das Element im lokalen Cache gespeichert.</span><span class="sxs-lookup"><span data-stu-id="406e6-p169">When invoked, this method saves the current message as a draft and returns the item id via the callback method. In Outlook Web App or Outlook in online mode, the item is saved to the server. In Outlook in cached mode, the item is saved to the local cache.</span></span>

> [!NOTE]
> <span data-ttu-id="406e6-1080">Wenn Ihr Add-In ruft `saveAsync` für ein Element im Entwurfsmodus, um das Abrufen einer `itemId` um mit EWS oder REST-API verwenden, beachten Sie, dass wenn Outlook im Cache-Modus ist, es kann einige Zeit dauern, bevor das Element tatsächlich auf dem Server synchronisiert wird.</span><span class="sxs-lookup"><span data-stu-id="406e6-1080">If your add-in calls `saveAsync` on an item in compose mode in order to get an `itemId` to use with EWS or the REST API, be aware that when Outlook is in cached mode, it may take some time before the item is actually synced to the server.</span></span> <span data-ttu-id="406e6-1081">Bis das Element mit synchronisiert ist, die `itemId` wird ein Fehler zurückgegeben.</span><span class="sxs-lookup"><span data-stu-id="406e6-1081">Until the item is synced, using the `itemId` will return an error.</span></span>

<span data-ttu-id="406e6-p171">Da Termine keinen Entwurfsstatus haben, wird das Element bei Aufruf von `saveAsync` für einen Termin im Verfassenmodus als normaler Termin im Kalender des Benutzers gespeichert. Für neue Termine, die noch nicht gespeichert wurden, wird keine Einladung gesendet. Beim Speichern eines vorhandenen Termins wird eine Aktualisierung an die hinzugefügten oder entfernten Teilnehmer gesendet.</span><span class="sxs-lookup"><span data-stu-id="406e6-p171">Since appointments have no draft state, if `saveAsync` is called on an appointment in compose mode, the item will be saved as a normal appointment on the user's calendar. For new appointments that have not been saved before, no invitation will be sent. Saving an existing appointment will send an update to added or removed attendees.</span></span>

> [!NOTE]
> <span data-ttu-id="406e6-1085">Die folgenden Clients haben unterschiedlichem Verhalten für `saveAsync` Termine im Verfassenmodus:</span><span class="sxs-lookup"><span data-stu-id="406e6-1085">The following clients have different behavior for `saveAsync` on appointments in compose mode:</span></span>
>
> - <span data-ttu-id="406e6-1086">Outlook für Mac unterstützt keine `saveAsync` auf einer Besprechung im Verfassenmodus.</span><span class="sxs-lookup"><span data-stu-id="406e6-1086">Mac Outlook does not support `saveAsync` on a meeting in compose mode.</span></span> <span data-ttu-id="406e6-1087">Aufrufen von `saveAsync` auf einer Besprechung in Outlook für Mac wird ein Fehler zurückgegeben.</span><span class="sxs-lookup"><span data-stu-id="406e6-1087">Calling `saveAsync` on a meeting in Mac Outlook will return an error.</span></span>
> - <span data-ttu-id="406e6-1088">Outlook im Web immer sendet eine Einladung oder zu aktualisieren, wenn `saveAsync` für einen Termin aufgerufen wird, im Verfassenmodus.</span><span class="sxs-lookup"><span data-stu-id="406e6-1088">Outlook on the web always sends an invitation or update when `saveAsync` is called on an appointment in compose mode.</span></span>

##### <a name="parameters"></a><span data-ttu-id="406e6-1089">Parameter:</span><span class="sxs-lookup"><span data-stu-id="406e6-1089">Parameters:</span></span>

|<span data-ttu-id="406e6-1090">Name</span><span class="sxs-lookup"><span data-stu-id="406e6-1090">Name</span></span>| <span data-ttu-id="406e6-1091">Typ</span><span class="sxs-lookup"><span data-stu-id="406e6-1091">Type</span></span>| <span data-ttu-id="406e6-1092">Attribute</span><span class="sxs-lookup"><span data-stu-id="406e6-1092">Attributes</span></span>| <span data-ttu-id="406e6-1093">Beschreibung</span><span class="sxs-lookup"><span data-stu-id="406e6-1093">Description</span></span>|
|---|---|---|---|
|`options`| <span data-ttu-id="406e6-1094">Objekt</span><span class="sxs-lookup"><span data-stu-id="406e6-1094">Object</span></span>| <span data-ttu-id="406e6-1095">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="406e6-1095">&lt;optional&gt;</span></span>|<span data-ttu-id="406e6-1096">Ein Objektliteral, das mindestens eine der folgenden Eigenschaften enthält.</span><span class="sxs-lookup"><span data-stu-id="406e6-1096">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`| <span data-ttu-id="406e6-1097">Object</span><span class="sxs-lookup"><span data-stu-id="406e6-1097">Object</span></span>| <span data-ttu-id="406e6-1098">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="406e6-1098">&lt;optional&gt;</span></span>|<span data-ttu-id="406e6-1099">Entwickler können ein Objekt bereitstellen, auf das sie in der Callbackmethode zugreifen möchten.</span><span class="sxs-lookup"><span data-stu-id="406e6-1099">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`| <span data-ttu-id="406e6-1100">Funktion</span><span class="sxs-lookup"><span data-stu-id="406e6-1100">function</span></span>||<span data-ttu-id="406e6-1101">Wenn die Methode abgeschlossen wird, wird die Funktion , die im `callback`-Parameter übergeben wird, mit einem einzelnen Parameter (`asyncResult`) aufgerufen, der ein [`AsyncResult`](/javascript/api/office/office.asyncresult)-Objekt darstellt.</span><span class="sxs-lookup"><span data-stu-id="406e6-1101">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span><br/><br/><span data-ttu-id="406e6-1102">Auf Erfolg, erfolgt die Element-ID der `asyncResult.value` Eigenschaft.</span><span class="sxs-lookup"><span data-stu-id="406e6-1102">On success, the item identifier is provided in the `asyncResult.value` property.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="406e6-1103">Anforderungen</span><span class="sxs-lookup"><span data-stu-id="406e6-1103">Requirements</span></span>

|<span data-ttu-id="406e6-1104">Anforderung</span><span class="sxs-lookup"><span data-stu-id="406e6-1104">Requirement</span></span>| <span data-ttu-id="406e6-1105">Wert</span><span class="sxs-lookup"><span data-stu-id="406e6-1105">Value</span></span>|
|---|---|
|[<span data-ttu-id="406e6-1106">Mindestversion des Postfachanforderungssatzes</span><span class="sxs-lookup"><span data-stu-id="406e6-1106">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="406e6-1107">1.3</span><span class="sxs-lookup"><span data-stu-id="406e6-1107">1.3</span></span>|
|[<span data-ttu-id="406e6-1108">Mindestberechtigungsstufe</span><span class="sxs-lookup"><span data-stu-id="406e6-1108">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="406e6-1109">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="406e6-1109">ReadWriteItem</span></span>|
|[<span data-ttu-id="406e6-1110">Zutreffender Outlook-Modus</span><span class="sxs-lookup"><span data-stu-id="406e6-1110">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="406e6-1111">Verfassen</span><span class="sxs-lookup"><span data-stu-id="406e6-1111">Compose</span></span>|

##### <a name="examples"></a><span data-ttu-id="406e6-1112">Beispiele</span><span class="sxs-lookup"><span data-stu-id="406e6-1112">Examples</span></span>

```
Office.context.mailbox.item.saveAsync(
  function callback(result) {
    // Process the result
  });
```

<span data-ttu-id="406e6-p173">Es folgt ein Beispiel für den `result`-Parameter, der an die Callbackfunktion übergeben wird. Die `value`-Eigenschaft enthält die Element-ID des Elements.</span><span class="sxs-lookup"><span data-stu-id="406e6-p173">The following is an example of the `result` parameter passed to the callback function. The `value` property contains the item ID of the item.</span></span>

```
{
  "value":"AAMkADI5...AAA=",
  "status":"succeeded"
}
```

####  <a name="setselecteddataasyncdata-options-callback"></a><span data-ttu-id="406e6-1115">setSelectedDataAsync(data, [options], callback)</span><span class="sxs-lookup"><span data-stu-id="406e6-1115">setSelectedDataAsync(data, [options], callback)</span></span>

<span data-ttu-id="406e6-1116">Fügt asynchron Daten in den Textkörper oder Betreff einer Nachricht ein.</span><span class="sxs-lookup"><span data-stu-id="406e6-1116">Asynchronously inserts data into the body or subject of a message.</span></span>

<span data-ttu-id="406e6-p174">Die `setSelectedDataAsync`-Methode fügt die angegebene Zeichenfolge an der Cursorposition im Betreff oder im Textkörper ein, oder, falls im Editor Text ausgewählt ist, ersetzt den markierten Text. Wenn sich der Cursor nicht im Textkörper oder im Betreffsfeld befindet, wird ein Fehler zurückgegeben. Nach dem Einfügen wird der Cursor am Ende der eingefügten Inhalte platziert.</span><span class="sxs-lookup"><span data-stu-id="406e6-p174">The `setSelectedDataAsync` method inserts the specified string at the cursor location in the subject or body of the item, or, if text is selected in the editor, it replaces the selected text. If the cursor is not in the body or subject field, an error is returned. After insertion, the cursor is placed at the end of the inserted content.</span></span>

##### <a name="parameters"></a><span data-ttu-id="406e6-1120">Parameter:</span><span class="sxs-lookup"><span data-stu-id="406e6-1120">Parameters:</span></span>

|<span data-ttu-id="406e6-1121">Name</span><span class="sxs-lookup"><span data-stu-id="406e6-1121">Name</span></span>| <span data-ttu-id="406e6-1122">Typ</span><span class="sxs-lookup"><span data-stu-id="406e6-1122">Type</span></span>| <span data-ttu-id="406e6-1123">Attribute</span><span class="sxs-lookup"><span data-stu-id="406e6-1123">Attributes</span></span>| <span data-ttu-id="406e6-1124">Beschreibung</span><span class="sxs-lookup"><span data-stu-id="406e6-1124">Description</span></span>|
|---|---|---|---|
|`data`| <span data-ttu-id="406e6-1125">Zeichenfolge</span><span class="sxs-lookup"><span data-stu-id="406e6-1125">String</span></span>||<span data-ttu-id="406e6-p175">Die einzufügenden Daten. Daten dürfen 1.000.000 Zeichen nicht überschreiten. Werden mehr als 1.000.000 Zeichen übergeben, wird eine `ArgumentOutOfRange`-Ausnahme ausgelöst.</span><span class="sxs-lookup"><span data-stu-id="406e6-p175">The data to be inserted. Data is not to exceed 1,000,000 characters. If more than 1,000,000 characters are passed in, an `ArgumentOutOfRange` exception is thrown.</span></span>|
|`options`| <span data-ttu-id="406e6-1129">Object</span><span class="sxs-lookup"><span data-stu-id="406e6-1129">Object</span></span>| <span data-ttu-id="406e6-1130">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="406e6-1130">&lt;optional&gt;</span></span>|<span data-ttu-id="406e6-1131">Ein Objektliteral, das mindestens eine der folgenden Eigenschaften enthält.</span><span class="sxs-lookup"><span data-stu-id="406e6-1131">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`| <span data-ttu-id="406e6-1132">Object</span><span class="sxs-lookup"><span data-stu-id="406e6-1132">Object</span></span>| <span data-ttu-id="406e6-1133">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="406e6-1133">&lt;optional&gt;</span></span>|<span data-ttu-id="406e6-1134">Entwickler können ein Objekt bereitstellen, auf das sie in der Rückrufmethode zugreifen möchten.</span><span class="sxs-lookup"><span data-stu-id="406e6-1134">Developers can provide any object they wish to access in the callback method.</span></span>|
|`options.coercionType`| [<span data-ttu-id="406e6-1135">Office.CoercionType</span><span class="sxs-lookup"><span data-stu-id="406e6-1135">Office.CoercionType</span></span>](office.md#coerciontype-string)| <span data-ttu-id="406e6-1136">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="406e6-1136">&lt;optional&gt;</span></span>|<span data-ttu-id="406e6-p176">Wenn `text`, wird das aktuelle Format in Outlook Web App und Outlook angewendet. Wenn das Feld ein HTML-Editor ist, werden nur die Textdaten eingefügt, selbst wenn es sich bei den Daten um HTML-Daten handelt.</span><span class="sxs-lookup"><span data-stu-id="406e6-p176">If `text`, the current style is applied in Outlook Web App and Outlook. If the field is an HTML editor, only the text data is inserted, even if the data is HTML.</span></span><br/><br/><span data-ttu-id="406e6-p177">Wenn `html` gesetzt ist und das Feld HTML unterstützt (der Betreff jedoch nicht), wird die aktuelle Formatvorlage in Outlook Web App angewendet und die Standardformatvorlage in Outlook. Ist das Feld ein Textfeld, wird ein Fehler des Typs `InvalidDataFormat` zurückgegeben.</span><span class="sxs-lookup"><span data-stu-id="406e6-p177">If `html` and the field supports HTML (the subject doesn't), the current style is applied in Outlook Web App and the default style is applied in Outlook. If the field is a text field, an `InvalidDataFormat` error is returned.</span></span><br/><br/><span data-ttu-id="406e6-1141">Wenn `coercionType` nicht festgelegt wird, hängt das Ergebnis vom Feld ab: Wenn das Feld HTML ist, wird HTML verwendet. Wenn das Feld Text ist, wird Nur-Text verwendet.</span><span class="sxs-lookup"><span data-stu-id="406e6-1141">If `coercionType` is not set, the result depends on the field: if the field is HTML then HTML is used; if the field is text, then plain text is used.</span></span>|
|`callback`| <span data-ttu-id="406e6-1142">function</span><span class="sxs-lookup"><span data-stu-id="406e6-1142">function</span></span>||<span data-ttu-id="406e6-1143">Wenn die Methode abgeschlossen wird, wird die Funktion , die im `callback`-Parameter übergeben wird, mit einem einzelnen Parameter (`asyncResult`) aufgerufen, der ein [`AsyncResult`](/javascript/api/office/office.asyncresult)-Objekt darstellt.</span><span class="sxs-lookup"><span data-stu-id="406e6-1143">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="406e6-1144">Anforderungen</span><span class="sxs-lookup"><span data-stu-id="406e6-1144">Requirements</span></span>

|<span data-ttu-id="406e6-1145">Anforderung</span><span class="sxs-lookup"><span data-stu-id="406e6-1145">Requirement</span></span>| <span data-ttu-id="406e6-1146">Wert</span><span class="sxs-lookup"><span data-stu-id="406e6-1146">Value</span></span>|
|---|---|
|[<span data-ttu-id="406e6-1147">Mindestversion des Postfachanforderungssatzes</span><span class="sxs-lookup"><span data-stu-id="406e6-1147">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="406e6-1148">1.2</span><span class="sxs-lookup"><span data-stu-id="406e6-1148">1.2</span></span>|
|[<span data-ttu-id="406e6-1149">Mindestberechtigungsstufe</span><span class="sxs-lookup"><span data-stu-id="406e6-1149">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="406e6-1150">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="406e6-1150">ReadWriteItem</span></span>|
|[<span data-ttu-id="406e6-1151">Zutreffender Outlook-Modus</span><span class="sxs-lookup"><span data-stu-id="406e6-1151">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="406e6-1152">Verfassen</span><span class="sxs-lookup"><span data-stu-id="406e6-1152">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="406e6-1153">Beispiel</span><span class="sxs-lookup"><span data-stu-id="406e6-1153">Example</span></span>

```
Office.context.mailbox.item.setSelectedDataAsync("Hello World!");
Office.context.mailbox.item.setSelectedDataAsync("<b>Hello World!</b>", { coercionType : "html" });
```