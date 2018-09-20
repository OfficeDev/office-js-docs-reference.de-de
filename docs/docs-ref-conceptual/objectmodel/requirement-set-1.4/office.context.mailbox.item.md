
# <a name="item"></a><span data-ttu-id="462e2-101">item</span><span class="sxs-lookup"><span data-stu-id="462e2-101">item</span></span>

### <a name="officeofficemdcontextofficecontextmdmailboxofficecontextmailboxmditem"></a><span data-ttu-id="462e2-102">[Office](office.md)[.context](office.context.md)[.mailbox](office.context.mailbox.md).item</span><span class="sxs-lookup"><span data-stu-id="462e2-102">[Office](office.md)[.context](office.context.md)[.mailbox](office.context.mailbox.md).item</span></span>

<span data-ttu-id="462e2-p101">Der `item`-Namespace wird für den Zugriff auf die aktuell ausgewählte Nachricht, Besprechungsanfrage oder den aktuell ausgewählten Termin verwendet. Sie können den Typ von `item` mithilfe der [itemType](#itemtype-officemailboxenumsitemtypejavascriptapioutlook14officemailboxenumsitemtype)-Eigenschaft bestimmen.</span><span class="sxs-lookup"><span data-stu-id="462e2-p101">The `item` namespace is used to access the currently selected message, meeting request, or appointment. You can determine the type of the `item` by using the [itemType](#itemtype-officemailboxenumsitemtypejavascriptapioutlook14officemailboxenumsitemtype) property.</span></span>

##### <a name="requirements"></a><span data-ttu-id="462e2-105">Anforderungen</span><span class="sxs-lookup"><span data-stu-id="462e2-105">Requirements</span></span>

|<span data-ttu-id="462e2-106">Anforderung</span><span class="sxs-lookup"><span data-stu-id="462e2-106">Requirement</span></span>| <span data-ttu-id="462e2-107">Wert</span><span class="sxs-lookup"><span data-stu-id="462e2-107">Value</span></span>|
|---|---|
|[<span data-ttu-id="462e2-108">Mindestversion des Postfachanforderungssatzes</span><span class="sxs-lookup"><span data-stu-id="462e2-108">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="462e2-109">1.0</span><span class="sxs-lookup"><span data-stu-id="462e2-109">1.0</span></span>|
|[<span data-ttu-id="462e2-110">Mindestberechtigungsstufe</span><span class="sxs-lookup"><span data-stu-id="462e2-110">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="462e2-111">Eingeschränkt</span><span class="sxs-lookup"><span data-stu-id="462e2-111">Restricted</span></span>|
|[<span data-ttu-id="462e2-112">Zutreffender Outlook-Modus</span><span class="sxs-lookup"><span data-stu-id="462e2-112">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="462e2-113">Verfassen oder Lesen</span><span class="sxs-lookup"><span data-stu-id="462e2-113">Compose or read</span></span>|

### <a name="example"></a><span data-ttu-id="462e2-114">Beispiel</span><span class="sxs-lookup"><span data-stu-id="462e2-114">Example</span></span>

<span data-ttu-id="462e2-115">Im folgenden JavaScript-Codebeispiel wird der Zugriff auf die `subject`-Eigenschaft des aktuellen Elements in Outlook veranschaulicht.</span><span class="sxs-lookup"><span data-stu-id="462e2-115">The following JavaScript code example shows how to access the `subject` property of the current item in Outlook.</span></span>

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

### <a name="members"></a><span data-ttu-id="462e2-116">Elemente</span><span class="sxs-lookup"><span data-stu-id="462e2-116">Members</span></span>

#### <a name="attachments-arrayattachmentdetailsjavascriptapioutlook14officeattachmentdetails"></a><span data-ttu-id="462e2-117">attachments :Array.<[AttachmentDetails](/javascript/api/outlook_1_4/office.attachmentdetails)></span><span class="sxs-lookup"><span data-stu-id="462e2-117">attachments :Array.<[AttachmentDetails](/javascript/api/outlook_1_4/office.attachmentdetails)></span></span>

<span data-ttu-id="462e2-p102">Ruft ein Array mit Anlagen für das Element ab. Nur Lesemodus.</span><span class="sxs-lookup"><span data-stu-id="462e2-p102">Gets an array of attachments for the item. Read mode only.</span></span>

> [!NOTE]
> <span data-ttu-id="462e2-120">Bestimmte Dateitypen werden aufgrund von Sicherheitsproblemen von Outlook blockiert und daher nicht zurückgegeben werden.</span><span class="sxs-lookup"><span data-stu-id="462e2-120">Certain types of files are blocked by Outlook due to potential security issues and are therefore not returned.</span></span> <span data-ttu-id="462e2-121">Weitere Informationen finden Sie unter [Blockierte Anlagen in Outlook](https://support.office.com/article/Blocked-attachments-in-Outlook-434752E1-02D3-4E90-9124-8B81E49A8519).</span><span class="sxs-lookup"><span data-stu-id="462e2-121">For more information, see [Blocked attachments in Outlook](https://support.office.com/article/Blocked-attachments-in-Outlook-434752E1-02D3-4E90-9124-8B81E49A8519).</span></span>

##### <a name="type"></a><span data-ttu-id="462e2-122">Typ:</span><span class="sxs-lookup"><span data-stu-id="462e2-122">Type:</span></span>

*   <span data-ttu-id="462e2-123">Array.<[AttachmentDetails](/javascript/api/outlook_1_4/office.attachmentdetails)></span><span class="sxs-lookup"><span data-stu-id="462e2-123">Array.<[AttachmentDetails](/javascript/api/outlook_1_4/office.attachmentdetails)></span></span>

##### <a name="requirements"></a><span data-ttu-id="462e2-124">Anforderungen</span><span class="sxs-lookup"><span data-stu-id="462e2-124">Requirements</span></span>

|<span data-ttu-id="462e2-125">Anforderung</span><span class="sxs-lookup"><span data-stu-id="462e2-125">Requirement</span></span>| <span data-ttu-id="462e2-126">Wert</span><span class="sxs-lookup"><span data-stu-id="462e2-126">Value</span></span>|
|---|---|
|[<span data-ttu-id="462e2-127">Mindestversion des Postfachanforderungssatzes</span><span class="sxs-lookup"><span data-stu-id="462e2-127">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="462e2-128">1.0</span><span class="sxs-lookup"><span data-stu-id="462e2-128">1.0</span></span>|
|[<span data-ttu-id="462e2-129">Mindestberechtigungsstufe</span><span class="sxs-lookup"><span data-stu-id="462e2-129">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="462e2-130">ReadItem</span><span class="sxs-lookup"><span data-stu-id="462e2-130">ReadItem</span></span>|
|[<span data-ttu-id="462e2-131">Zutreffender Outlook-Modus</span><span class="sxs-lookup"><span data-stu-id="462e2-131">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="462e2-132">Lesen</span><span class="sxs-lookup"><span data-stu-id="462e2-132">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="462e2-133">Beispiel</span><span class="sxs-lookup"><span data-stu-id="462e2-133">Example</span></span>

<span data-ttu-id="462e2-134">Mit dem folgende Code wird eine HTML-Zeichenfolge mit Details aller Anlagen für das aktuelle Element erstellt.</span><span class="sxs-lookup"><span data-stu-id="462e2-134">The following code builds an HTML string with details of all attachments on the current item.</span></span>

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

####  <a name="bcc-recipientsjavascriptapioutlook14officerecipients"></a><span data-ttu-id="462e2-135">bcc :[Recipients](/javascript/api/outlook_1_4/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="462e2-135">bcc :[Recipients](/javascript/api/outlook_1_4/office.recipients)</span></span>

<span data-ttu-id="462e2-136">Ruft ein Objekt, das Methoden zum Abrufen oder aktualisieren Sie die Zeile Bcc (blind Carbon Copy, Blindkopie) einer Nachricht bereitstellt.</span><span class="sxs-lookup"><span data-stu-id="462e2-136">Gets an object that provides methods to get or update the Bcc (blind carbon copy) line of a message.</span></span> <span data-ttu-id="462e2-137">Nur Verfassenmodus.</span><span class="sxs-lookup"><span data-stu-id="462e2-137">Compose mode only.</span></span>

##### <a name="type"></a><span data-ttu-id="462e2-138">Typ:</span><span class="sxs-lookup"><span data-stu-id="462e2-138">Type:</span></span>

*   [<span data-ttu-id="462e2-139">Recipients</span><span class="sxs-lookup"><span data-stu-id="462e2-139">Recipients</span></span>](/javascript/api/outlook_1_4/office.recipients)

##### <a name="requirements"></a><span data-ttu-id="462e2-140">Anforderungen</span><span class="sxs-lookup"><span data-stu-id="462e2-140">Requirements</span></span>

|<span data-ttu-id="462e2-141">Anforderung</span><span class="sxs-lookup"><span data-stu-id="462e2-141">Requirement</span></span>| <span data-ttu-id="462e2-142">Wert</span><span class="sxs-lookup"><span data-stu-id="462e2-142">Value</span></span>|
|---|---|
|[<span data-ttu-id="462e2-143">Mindestversion des Postfachanforderungssatzes</span><span class="sxs-lookup"><span data-stu-id="462e2-143">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="462e2-144">1.1</span><span class="sxs-lookup"><span data-stu-id="462e2-144">1.1</span></span>|
|[<span data-ttu-id="462e2-145">Mindestberechtigungsstufe</span><span class="sxs-lookup"><span data-stu-id="462e2-145">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="462e2-146">ReadItem</span><span class="sxs-lookup"><span data-stu-id="462e2-146">ReadItem</span></span>|
|[<span data-ttu-id="462e2-147">Zutreffender Outlook-Modus</span><span class="sxs-lookup"><span data-stu-id="462e2-147">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="462e2-148">Verfassen</span><span class="sxs-lookup"><span data-stu-id="462e2-148">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="462e2-149">Beispiel</span><span class="sxs-lookup"><span data-stu-id="462e2-149">Example</span></span>

```
Office.context.mailbox.item.bcc.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.bcc.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.bcc.getAsync(callback);

function callback(asyncResult) {
  var arrayOfBccRecipients = asyncResult.value;
}
```

####  <a name="body-bodyjavascriptapioutlook14officebody"></a><span data-ttu-id="462e2-150">body :[Body](/javascript/api/outlook_1_4/office.body)</span><span class="sxs-lookup"><span data-stu-id="462e2-150">body :[Body](/javascript/api/outlook_1_4/office.body)</span></span>

<span data-ttu-id="462e2-151">Ruft ein Objekt ab, das Methoden zum Bearbeiten des Textkörpers eines Elements bereitstellt.</span><span class="sxs-lookup"><span data-stu-id="462e2-151">Gets an object that provides methods for manipulating the body of an item.</span></span>

##### <a name="type"></a><span data-ttu-id="462e2-152">Typ:</span><span class="sxs-lookup"><span data-stu-id="462e2-152">Type:</span></span>

*   [<span data-ttu-id="462e2-153">Body</span><span class="sxs-lookup"><span data-stu-id="462e2-153">Body</span></span>](/javascript/api/outlook_1_4/office.body)

##### <a name="requirements"></a><span data-ttu-id="462e2-154">Anforderungen</span><span class="sxs-lookup"><span data-stu-id="462e2-154">Requirements</span></span>

|<span data-ttu-id="462e2-155">Anforderung</span><span class="sxs-lookup"><span data-stu-id="462e2-155">Requirement</span></span>| <span data-ttu-id="462e2-156">Wert</span><span class="sxs-lookup"><span data-stu-id="462e2-156">Value</span></span>|
|---|---|
|[<span data-ttu-id="462e2-157">Mindestversion des Postfachanforderungssatzes</span><span class="sxs-lookup"><span data-stu-id="462e2-157">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="462e2-158">1.1</span><span class="sxs-lookup"><span data-stu-id="462e2-158">1.1</span></span>|
|[<span data-ttu-id="462e2-159">Mindestberechtigungsstufe</span><span class="sxs-lookup"><span data-stu-id="462e2-159">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="462e2-160">ReadItem</span><span class="sxs-lookup"><span data-stu-id="462e2-160">ReadItem</span></span>|
|[<span data-ttu-id="462e2-161">Zutreffender Outlook-Modus</span><span class="sxs-lookup"><span data-stu-id="462e2-161">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="462e2-162">Verfassen oder Lesen</span><span class="sxs-lookup"><span data-stu-id="462e2-162">Compose or read</span></span>|

####  <a name="cc-arrayemailaddressdetailsjavascriptapioutlook14officeemailaddressdetailsrecipientsjavascriptapioutlook14officerecipients"></a><span data-ttu-id="462e2-163">cc: Array. <[EmailAddressDetails](/javascript/api/outlook_1_4/office.emailaddressdetails)>|[Empfänger](/javascript/api/outlook_1_4/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="462e2-163">cc :Array.<[EmailAddressDetails](/javascript/api/outlook_1_4/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook_1_4/office.recipients)</span></span>

<span data-ttu-id="462e2-164">Ermöglicht den Zugriff auf die (Carbon Copy, Blindkopie) Cc-Empfänger einer Nachricht.</span><span class="sxs-lookup"><span data-stu-id="462e2-164">Provides access to the Cc (carbon copy) recipients of a message.</span></span> <span data-ttu-id="462e2-165">Der Typ des Objekts und die Zugriffsebene, hängt von den Modus des aktuellen Elements ab.</span><span class="sxs-lookup"><span data-stu-id="462e2-165">The type of object and level of access depends on the mode of the current item.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="462e2-166">Lesemodus</span><span class="sxs-lookup"><span data-stu-id="462e2-166">Read mode</span></span>

<span data-ttu-id="462e2-p106">Die `cc`-Eigenschaft gibt ein Array mit einem `EmailAddressDetails`-Objekt für jeden Empfänger in der **Cc**-Zeile der Nachricht zurück. Die Auflistung ist auf höchstens 100 Elemente beschränkt.</span><span class="sxs-lookup"><span data-stu-id="462e2-p106">The `cc` property returns an array that contains an `EmailAddressDetails` object for each recipient listed on the **Cc** line of the message. The collection is limited to a maximum of 100 members.</span></span>

##### <a name="compose-mode"></a><span data-ttu-id="462e2-169">Verfassenmodus</span><span class="sxs-lookup"><span data-stu-id="462e2-169">Compose mode</span></span>

<span data-ttu-id="462e2-170">Die `cc` -Eigenschaft gibt eine `Recipients` -Objekt, das Methoden zum Abrufen oder aktualisieren die Empfänger in der Zeile **Cc** der Nachricht bereitstellt.</span><span class="sxs-lookup"><span data-stu-id="462e2-170">The `cc` property returns a `Recipients` object that provides methods to get or update the recipients on the **Cc** line of the message.</span></span>

##### <a name="type"></a><span data-ttu-id="462e2-171">Typ:</span><span class="sxs-lookup"><span data-stu-id="462e2-171">Type:</span></span>

*   <span data-ttu-id="462e2-172">Array.<[EmailAddressDetails](/javascript/api/outlook_1_4/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook_1_4/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="462e2-172">Array.<[EmailAddressDetails](/javascript/api/outlook_1_4/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook_1_4/office.recipients)</span></span>

##### <a name="requirements"></a><span data-ttu-id="462e2-173">Anforderungen</span><span class="sxs-lookup"><span data-stu-id="462e2-173">Requirements</span></span>

|<span data-ttu-id="462e2-174">Anforderung</span><span class="sxs-lookup"><span data-stu-id="462e2-174">Requirement</span></span>| <span data-ttu-id="462e2-175">Wert</span><span class="sxs-lookup"><span data-stu-id="462e2-175">Value</span></span>|
|---|---|
|[<span data-ttu-id="462e2-176">Mindestversion des Postfachanforderungssatzes</span><span class="sxs-lookup"><span data-stu-id="462e2-176">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="462e2-177">1.0</span><span class="sxs-lookup"><span data-stu-id="462e2-177">1.0</span></span>|
|[<span data-ttu-id="462e2-178">Mindestberechtigungsstufe</span><span class="sxs-lookup"><span data-stu-id="462e2-178">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="462e2-179">ReadItem</span><span class="sxs-lookup"><span data-stu-id="462e2-179">ReadItem</span></span>|
|[<span data-ttu-id="462e2-180">Zutreffender Outlook-Modus</span><span class="sxs-lookup"><span data-stu-id="462e2-180">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="462e2-181">Verfassen oder Lesen</span><span class="sxs-lookup"><span data-stu-id="462e2-181">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="462e2-182">Beispiel</span><span class="sxs-lookup"><span data-stu-id="462e2-182">Example</span></span>

```
Office.context.mailbox.item.cc.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.cc.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.cc.getAsync(callback);

function callback(asyncResult) {
  var arrayOfCcRecipients = asyncResult.value;
}
```

####  <a name="nullable-conversationid-string"></a><span data-ttu-id="462e2-183">(nullable) conversationId :String</span><span class="sxs-lookup"><span data-stu-id="462e2-183">(nullable) conversationId :String</span></span>

<span data-ttu-id="462e2-184">Ruft einen Bezeichner für die E-Mail-Unterhaltung ab, in der eine bestimmte Nachricht enthalten ist.</span><span class="sxs-lookup"><span data-stu-id="462e2-184">Gets an identifier for the email conversation that contains a particular message.</span></span>

<span data-ttu-id="462e2-p107">Sie können für diese Eigenschaft eine ganze Zahl abrufen, wenn Ihre Mail-App in Formularen zum Lesen oder Antworten in Formularen zum Verfassen aktiviert wird. Wenn der Benutzer den Betreff der Antwortnachricht ändert, ändert sich beim Versenden die Konversations-ID für die entsprechende Nachricht, und der Wert, den Sie vorher bezogen haben, trifft nicht länger zu.</span><span class="sxs-lookup"><span data-stu-id="462e2-p107">You can get an integer for this property if your mail app is activated in read forms or responses in compose forms. If subsequently the user changes the subject of the reply message, upon sending the reply, the conversation ID for that message will change and that value you obtained earlier will no longer apply.</span></span>

<span data-ttu-id="462e2-p108">Sie erhalten in einem Formular zum Verfassen für diese Eigenschaft für ein neues Element null. Wenn der Benutzer einen Betreff festlegt und das Element speichert, gibt die `conversationId`-Eigenschaft einen Wert zurück.</span><span class="sxs-lookup"><span data-stu-id="462e2-p108">You get null for this property for a new item in a compose form. If the user sets a subject and saves the item, the `conversationId` property will return a value.</span></span>

##### <a name="type"></a><span data-ttu-id="462e2-189">Typ:</span><span class="sxs-lookup"><span data-stu-id="462e2-189">Type:</span></span>

*   <span data-ttu-id="462e2-190">String</span><span class="sxs-lookup"><span data-stu-id="462e2-190">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="462e2-191">Anforderungen</span><span class="sxs-lookup"><span data-stu-id="462e2-191">Requirements</span></span>

|<span data-ttu-id="462e2-192">Anforderung</span><span class="sxs-lookup"><span data-stu-id="462e2-192">Requirement</span></span>| <span data-ttu-id="462e2-193">Wert</span><span class="sxs-lookup"><span data-stu-id="462e2-193">Value</span></span>|
|---|---|
|[<span data-ttu-id="462e2-194">Mindestversion des Postfachanforderungssatzes</span><span class="sxs-lookup"><span data-stu-id="462e2-194">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="462e2-195">1.0</span><span class="sxs-lookup"><span data-stu-id="462e2-195">1.0</span></span>|
|[<span data-ttu-id="462e2-196">Mindestberechtigungsstufe</span><span class="sxs-lookup"><span data-stu-id="462e2-196">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="462e2-197">ReadItem</span><span class="sxs-lookup"><span data-stu-id="462e2-197">ReadItem</span></span>|
|[<span data-ttu-id="462e2-198">Zutreffender Outlook-Modus</span><span class="sxs-lookup"><span data-stu-id="462e2-198">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="462e2-199">Verfassen oder Lesen</span><span class="sxs-lookup"><span data-stu-id="462e2-199">Compose or read</span></span>|

#### <a name="datetimecreated-date"></a><span data-ttu-id="462e2-200">dateTimeCreated :Date</span><span class="sxs-lookup"><span data-stu-id="462e2-200">dateTimeCreated :Date</span></span>

<span data-ttu-id="462e2-p109">Ruft das Datum und die Uhrzeit der Erstellung eines Elements ab. Nur Lesemodus.</span><span class="sxs-lookup"><span data-stu-id="462e2-p109">Gets the date and time that an item was created. Read mode only.</span></span>

##### <a name="type"></a><span data-ttu-id="462e2-203">Typ:</span><span class="sxs-lookup"><span data-stu-id="462e2-203">Type:</span></span>

*   <span data-ttu-id="462e2-204">Datum</span><span class="sxs-lookup"><span data-stu-id="462e2-204">Date</span></span>

##### <a name="requirements"></a><span data-ttu-id="462e2-205">Anforderungen</span><span class="sxs-lookup"><span data-stu-id="462e2-205">Requirements</span></span>

|<span data-ttu-id="462e2-206">Anforderung</span><span class="sxs-lookup"><span data-stu-id="462e2-206">Requirement</span></span>| <span data-ttu-id="462e2-207">Wert</span><span class="sxs-lookup"><span data-stu-id="462e2-207">Value</span></span>|
|---|---|
|[<span data-ttu-id="462e2-208">Mindestversion des Postfachanforderungssatzes</span><span class="sxs-lookup"><span data-stu-id="462e2-208">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="462e2-209">1.0</span><span class="sxs-lookup"><span data-stu-id="462e2-209">1.0</span></span>|
|[<span data-ttu-id="462e2-210">Mindestberechtigungsstufe</span><span class="sxs-lookup"><span data-stu-id="462e2-210">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="462e2-211">ReadItem</span><span class="sxs-lookup"><span data-stu-id="462e2-211">ReadItem</span></span>|
|[<span data-ttu-id="462e2-212">Zutreffender Outlook-Modus</span><span class="sxs-lookup"><span data-stu-id="462e2-212">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="462e2-213">Lesen</span><span class="sxs-lookup"><span data-stu-id="462e2-213">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="462e2-214">Beispiel</span><span class="sxs-lookup"><span data-stu-id="462e2-214">Example</span></span>

```
var created = Office.context.mailbox.item.dateTimeCreated;
```

#### <a name="datetimemodified-date"></a><span data-ttu-id="462e2-215">dateTimeModified :Date</span><span class="sxs-lookup"><span data-stu-id="462e2-215">dateTimeModified :Date</span></span>

<span data-ttu-id="462e2-p110">Ruft das Datum und die Uhrzeit der letzten Änderung eines Elements ab. Nur Lesemodus.</span><span class="sxs-lookup"><span data-stu-id="462e2-p110">Gets the date and time that an item was last modified. Read mode only.</span></span>

> [!NOTE]
> <span data-ttu-id="462e2-218">Dieser Member wird in Outlook für iOS oder Outlook für Android nicht unterstützt.</span><span class="sxs-lookup"><span data-stu-id="462e2-218">This member is not supported in Outlook for iOS or Outlook for Android.</span></span>

##### <a name="type"></a><span data-ttu-id="462e2-219">Typ:</span><span class="sxs-lookup"><span data-stu-id="462e2-219">Type:</span></span>

*   <span data-ttu-id="462e2-220">Datum</span><span class="sxs-lookup"><span data-stu-id="462e2-220">Date</span></span>

##### <a name="requirements"></a><span data-ttu-id="462e2-221">Anforderungen</span><span class="sxs-lookup"><span data-stu-id="462e2-221">Requirements</span></span>

|<span data-ttu-id="462e2-222">Anforderung</span><span class="sxs-lookup"><span data-stu-id="462e2-222">Requirement</span></span>| <span data-ttu-id="462e2-223">Wert</span><span class="sxs-lookup"><span data-stu-id="462e2-223">Value</span></span>|
|---|---|
|[<span data-ttu-id="462e2-224">Mindestversion des Postfachanforderungssatzes</span><span class="sxs-lookup"><span data-stu-id="462e2-224">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="462e2-225">1.0</span><span class="sxs-lookup"><span data-stu-id="462e2-225">1.0</span></span>|
|[<span data-ttu-id="462e2-226">Mindestberechtigungsstufe</span><span class="sxs-lookup"><span data-stu-id="462e2-226">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="462e2-227">ReadItem</span><span class="sxs-lookup"><span data-stu-id="462e2-227">ReadItem</span></span>|
|[<span data-ttu-id="462e2-228">Zutreffender Outlook-Modus</span><span class="sxs-lookup"><span data-stu-id="462e2-228">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="462e2-229">Lesen</span><span class="sxs-lookup"><span data-stu-id="462e2-229">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="462e2-230">Beispiel</span><span class="sxs-lookup"><span data-stu-id="462e2-230">Example</span></span>

```
var modified = Office.context.mailbox.item.dateTimeModified;
```

####  <a name="end-datetimejavascriptapioutlook14officetime"></a><span data-ttu-id="462e2-231">end :Date|[Time](/javascript/api/outlook_1_4/office.time)</span><span class="sxs-lookup"><span data-stu-id="462e2-231">end :Date|[Time](/javascript/api/outlook_1_4/office.time)</span></span>

<span data-ttu-id="462e2-232">Ruft Datum und Zeit für das Ende des Termins ab oder legt diese fest.</span><span class="sxs-lookup"><span data-stu-id="462e2-232">Gets or sets the date and time that the appointment is to end.</span></span>

<span data-ttu-id="462e2-p111">Die `end`-Eigenschaft wird als Datums- und Uhrzeitwert in UTC ausgedrückt. Sie können die [`convertToLocalClientTime`](office.context.mailbox.md#converttolocalclienttimetimevalue--localclienttimejavascriptapioutlook14officelocalclienttime)-Methode verwenden, um den Endeigenschaftswert in den lokalen Uhrzeit- und Datumswert des Clients umzuwandeln.</span><span class="sxs-lookup"><span data-stu-id="462e2-p111">The `end` property is expressed as a Coordinated Universal Time (UTC) date and time value. You can use the [`convertToLocalClientTime`](office.context.mailbox.md#converttolocalclienttimetimevalue--localclienttimejavascriptapioutlook14officelocalclienttime) method to convert the end property value to the client’s local date and time.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="462e2-235">Lesemodus</span><span class="sxs-lookup"><span data-stu-id="462e2-235">Read mode</span></span>

<span data-ttu-id="462e2-236">Die `end`-Eigenschaft gibt ein `Date`-Objekt zurück.</span><span class="sxs-lookup"><span data-stu-id="462e2-236">The `end` property returns a `Date` object.</span></span>

##### <a name="compose-mode"></a><span data-ttu-id="462e2-237">Verfassenmodus</span><span class="sxs-lookup"><span data-stu-id="462e2-237">Compose mode</span></span>

<span data-ttu-id="462e2-238">Die `end`-Eigenschaft gibt ein `Time`-Objekt zurück.</span><span class="sxs-lookup"><span data-stu-id="462e2-238">The `end` property returns a `Time` object.</span></span>

<span data-ttu-id="462e2-239">Wenn Sie die [`Time.setAsync`](/javascript/api/outlook_1_4/office.time#setasync-datetime--options--callback-)-Methode verwenden, um die Endzeit im Verfassenmodus festzulegen, sollten Sie die [`convertToUtcClientTime`](office.context.mailbox.md#converttoutcclienttimeinput--date)-Methode verwenden, um die Ortszeit auf dem Client für den Server in UTC umzuwandeln.</span><span class="sxs-lookup"><span data-stu-id="462e2-239">When you use the [`Time.setAsync`](/javascript/api/outlook_1_4/office.time#setasync-datetime--options--callback-) method to set the end time, you should use the [`convertToUtcClientTime`](office.context.mailbox.md#converttoutcclienttimeinput--date) method to convert the local time on the client to UTC for the server.</span></span>

##### <a name="type"></a><span data-ttu-id="462e2-240">Typ:</span><span class="sxs-lookup"><span data-stu-id="462e2-240">Type:</span></span>

*   <span data-ttu-id="462e2-241">Date | [Time](/javascript/api/outlook_1_4/office.time)</span><span class="sxs-lookup"><span data-stu-id="462e2-241">Date | [Time](/javascript/api/outlook_1_4/office.time)</span></span>

##### <a name="requirements"></a><span data-ttu-id="462e2-242">Anforderungen</span><span class="sxs-lookup"><span data-stu-id="462e2-242">Requirements</span></span>

|<span data-ttu-id="462e2-243">Anforderung</span><span class="sxs-lookup"><span data-stu-id="462e2-243">Requirement</span></span>| <span data-ttu-id="462e2-244">Wert</span><span class="sxs-lookup"><span data-stu-id="462e2-244">Value</span></span>|
|---|---|
|[<span data-ttu-id="462e2-245">Mindestversion des Postfachanforderungssatzes</span><span class="sxs-lookup"><span data-stu-id="462e2-245">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="462e2-246">1.0</span><span class="sxs-lookup"><span data-stu-id="462e2-246">1.0</span></span>|
|[<span data-ttu-id="462e2-247">Mindestberechtigungsstufe</span><span class="sxs-lookup"><span data-stu-id="462e2-247">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="462e2-248">ReadItem</span><span class="sxs-lookup"><span data-stu-id="462e2-248">ReadItem</span></span>|
|[<span data-ttu-id="462e2-249">Zutreffender Outlook-Modus</span><span class="sxs-lookup"><span data-stu-id="462e2-249">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="462e2-250">Verfassen oder Lesen</span><span class="sxs-lookup"><span data-stu-id="462e2-250">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="462e2-251">Beispiel</span><span class="sxs-lookup"><span data-stu-id="462e2-251">Example</span></span>

<span data-ttu-id="462e2-252">Im folgenden Beispiel wird die Endzeit eines Termins im Verfassenmodus mithilfe der [`setAsync`](/javascript/api/outlook_1_4/office.time#setasync-datetime--options--callback-)-Methode des `Time`-Objekts festgelegt.</span><span class="sxs-lookup"><span data-stu-id="462e2-252">The following example sets the end time of an appointment in compose mode by using the [`setAsync`](/javascript/api/outlook_1_4/office.time#setasync-datetime--options--callback-) method of the `Time` object.</span></span>

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

#### <a name="from-emailaddressdetailsjavascriptapioutlook14officeemailaddressdetails"></a><span data-ttu-id="462e2-253">from :[EmailAddressDetails](/javascript/api/outlook_1_4/office.emailaddressdetails)</span><span class="sxs-lookup"><span data-stu-id="462e2-253">from :[EmailAddressDetails](/javascript/api/outlook_1_4/office.emailaddressdetails)</span></span>

<span data-ttu-id="462e2-p112">Ruft die E-Mail-Adresse des Absenders einer Nachricht ab. Nur Lesemodus.</span><span class="sxs-lookup"><span data-stu-id="462e2-p112">Gets the email address of the sender of a message. Read mode only.</span></span>

<span data-ttu-id="462e2-p113">Die `from`- und [`sender`](#sender-emailaddressdetailsjavascriptapioutlook14officeemailaddressdetails)-Eigenschaften stellen dieselbe Person dar, außer die Nachricht wurde von einem Delegaten gesendet. In diesem Fall stellt die `from`-Eigenschaft die Stellvertretung dar, und die sender-Eigenschaft stellt den Delegaten dar.</span><span class="sxs-lookup"><span data-stu-id="462e2-p113">The `from` and [`sender`](#sender-emailaddressdetailsjavascriptapioutlook14officeemailaddressdetails) properties represent the same person unless the message is sent by a delegate. In that case, the `from` property represents the delegator, and the sender property represents the delegate.</span></span>

> [!NOTE]
> <span data-ttu-id="462e2-258">Die `recipientType` -Eigenschaft des der `EmailAddressDetails` -Objekts der `from` -Eigenschaft ist `undefined`.</span><span class="sxs-lookup"><span data-stu-id="462e2-258">The `recipientType` property of the `EmailAddressDetails` object in the `from` property is `undefined`.</span></span>

##### <a name="type"></a><span data-ttu-id="462e2-259">Typ:</span><span class="sxs-lookup"><span data-stu-id="462e2-259">Type:</span></span>

*   [<span data-ttu-id="462e2-260">EmailAddressDetails</span><span class="sxs-lookup"><span data-stu-id="462e2-260">EmailAddressDetails</span></span>](/javascript/api/outlook_1_4/office.emailaddressdetails)

##### <a name="requirements"></a><span data-ttu-id="462e2-261">Anforderungen</span><span class="sxs-lookup"><span data-stu-id="462e2-261">Requirements</span></span>

|<span data-ttu-id="462e2-262">Anforderung</span><span class="sxs-lookup"><span data-stu-id="462e2-262">Requirement</span></span>| <span data-ttu-id="462e2-263">Wert</span><span class="sxs-lookup"><span data-stu-id="462e2-263">Value</span></span>|
|---|---|
|[<span data-ttu-id="462e2-264">Mindestversion des Postfachanforderungssatzes</span><span class="sxs-lookup"><span data-stu-id="462e2-264">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="462e2-265">1.0</span><span class="sxs-lookup"><span data-stu-id="462e2-265">1.0</span></span>|
|[<span data-ttu-id="462e2-266">Mindestberechtigungsstufe</span><span class="sxs-lookup"><span data-stu-id="462e2-266">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="462e2-267">ReadItem</span><span class="sxs-lookup"><span data-stu-id="462e2-267">ReadItem</span></span>|
|[<span data-ttu-id="462e2-268">Zutreffender Outlook-Modus</span><span class="sxs-lookup"><span data-stu-id="462e2-268">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="462e2-269">Lesen</span><span class="sxs-lookup"><span data-stu-id="462e2-269">Read</span></span>|

#### <a name="internetmessageid-string"></a><span data-ttu-id="462e2-270">internetMessageId :String</span><span class="sxs-lookup"><span data-stu-id="462e2-270">internetMessageId :String</span></span>

<span data-ttu-id="462e2-p114">Ruft die Internetnachricht-ID einer E-Mail-Nachricht ab. Nur Lesemodus.</span><span class="sxs-lookup"><span data-stu-id="462e2-p114">Gets the Internet message identifier for an email message. Read mode only.</span></span>

##### <a name="type"></a><span data-ttu-id="462e2-273">Typ:</span><span class="sxs-lookup"><span data-stu-id="462e2-273">Type:</span></span>

*   <span data-ttu-id="462e2-274">String</span><span class="sxs-lookup"><span data-stu-id="462e2-274">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="462e2-275">Anforderungen</span><span class="sxs-lookup"><span data-stu-id="462e2-275">Requirements</span></span>

|<span data-ttu-id="462e2-276">Anforderung</span><span class="sxs-lookup"><span data-stu-id="462e2-276">Requirement</span></span>| <span data-ttu-id="462e2-277">Wert</span><span class="sxs-lookup"><span data-stu-id="462e2-277">Value</span></span>|
|---|---|
|[<span data-ttu-id="462e2-278">Mindestversion des Postfachanforderungssatzes</span><span class="sxs-lookup"><span data-stu-id="462e2-278">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="462e2-279">1.0</span><span class="sxs-lookup"><span data-stu-id="462e2-279">1.0</span></span>|
|[<span data-ttu-id="462e2-280">Mindestberechtigungsstufe</span><span class="sxs-lookup"><span data-stu-id="462e2-280">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="462e2-281">ReadItem</span><span class="sxs-lookup"><span data-stu-id="462e2-281">ReadItem</span></span>|
|[<span data-ttu-id="462e2-282">Zutreffender Outlook-Modus</span><span class="sxs-lookup"><span data-stu-id="462e2-282">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="462e2-283">Lesen</span><span class="sxs-lookup"><span data-stu-id="462e2-283">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="462e2-284">Beispiel</span><span class="sxs-lookup"><span data-stu-id="462e2-284">Example</span></span>

```
var internetMessageId = Office.context.mailbox.item.internetMessageId;
```

#### <a name="itemclass-string"></a><span data-ttu-id="462e2-285">itemClass :String</span><span class="sxs-lookup"><span data-stu-id="462e2-285">itemClass :String</span></span>

<span data-ttu-id="462e2-p115">Ruft die Exchange-Webdienste-Elementklasse des ausgewählten Elements ab. Nur Lesemodus.</span><span class="sxs-lookup"><span data-stu-id="462e2-p115">Gets the Exchange Web Services item class of the selected item. Read mode only.</span></span>

<span data-ttu-id="462e2-p116">Die `itemClass`-Eigenschaft gibt die Nachrichtenklasse des ausgewählten Elements an. Folgende Standardnachrichtenklassen für das Nachrichten- oder Terminelement sind vorhanden:</span><span class="sxs-lookup"><span data-stu-id="462e2-p116">The `itemClass` property specifies the message class of the selected item. The following are the default message classes for the message or appointment item.</span></span>

| <span data-ttu-id="462e2-290">Typ</span><span class="sxs-lookup"><span data-stu-id="462e2-290">Type</span></span> | <span data-ttu-id="462e2-291">Beschreibung</span><span class="sxs-lookup"><span data-stu-id="462e2-291">Description</span></span> | <span data-ttu-id="462e2-292">Elementklasse</span><span class="sxs-lookup"><span data-stu-id="462e2-292">item class</span></span> |
| --- | --- | --- |
| <span data-ttu-id="462e2-293">Terminelemente</span><span class="sxs-lookup"><span data-stu-id="462e2-293">Appointment items</span></span> | <span data-ttu-id="462e2-294">Dies sind Kalenderelemente der Elementklasse `IPM.Appointment` oder `IPM.Appointment.Occurence`.</span><span class="sxs-lookup"><span data-stu-id="462e2-294">These are calendar items of the item class `IPM.Appointment` or `IPM.Appointment.Occurence`.</span></span> | `IPM.Appointment`<br />`IPM.Appointment.Occurence` |
| <span data-ttu-id="462e2-295">Nachrichtenelemente</span><span class="sxs-lookup"><span data-stu-id="462e2-295">Message items</span></span> | <span data-ttu-id="462e2-296">Diese Elemente umfassen E-Mail-Nachrichten der Standardnachrichtenklasse `IPM.Note` sowie Besprechungsanfragen, -antworten und -absagen, die `IPM.Schedule.Meeting` als Basisnachrichtenklasse verwenden.</span><span class="sxs-lookup"><span data-stu-id="462e2-296">These include email messages that have the default message class `IPM.Note`, and meeting requests, responses, and cancellations, that use `IPM.Schedule.Meeting` as the base message class.</span></span> | `IPM.Note`<br />`IPM.Schedule.Meeting.Request`<br />`IPM.Schedule.Meeting.Neg`<br />`IPM.Schedule.Meeting.Pos`<br />`IPM.Schedule.Meeting.Tent`<br />`IPM.Schedule.Meeting.Canceled` |

<span data-ttu-id="462e2-297">Sie können benutzerdefinierte Nachrichtenklassen zum Erweitern einer Standardnachrichtenklasse erstellen, z. B. eine benutzerdefinierte Terminnachrichtenklasse wie `IPM.Appointment.Contoso`.</span><span class="sxs-lookup"><span data-stu-id="462e2-297">You can create custom message classes that extends a default message class, for example, a custom appointment message class `IPM.Appointment.Contoso`.</span></span>

##### <a name="type"></a><span data-ttu-id="462e2-298">Typ:</span><span class="sxs-lookup"><span data-stu-id="462e2-298">Type:</span></span>

*   <span data-ttu-id="462e2-299">String</span><span class="sxs-lookup"><span data-stu-id="462e2-299">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="462e2-300">Anforderungen</span><span class="sxs-lookup"><span data-stu-id="462e2-300">Requirements</span></span>

|<span data-ttu-id="462e2-301">Anforderung</span><span class="sxs-lookup"><span data-stu-id="462e2-301">Requirement</span></span>| <span data-ttu-id="462e2-302">Wert</span><span class="sxs-lookup"><span data-stu-id="462e2-302">Value</span></span>|
|---|---|
|[<span data-ttu-id="462e2-303">Mindestversion des Postfachanforderungssatzes</span><span class="sxs-lookup"><span data-stu-id="462e2-303">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="462e2-304">1.0</span><span class="sxs-lookup"><span data-stu-id="462e2-304">1.0</span></span>|
|[<span data-ttu-id="462e2-305">Mindestberechtigungsstufe</span><span class="sxs-lookup"><span data-stu-id="462e2-305">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="462e2-306">ReadItem</span><span class="sxs-lookup"><span data-stu-id="462e2-306">ReadItem</span></span>|
|[<span data-ttu-id="462e2-307">Zutreffender Outlook-Modus</span><span class="sxs-lookup"><span data-stu-id="462e2-307">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="462e2-308">Lesen</span><span class="sxs-lookup"><span data-stu-id="462e2-308">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="462e2-309">Beispiel</span><span class="sxs-lookup"><span data-stu-id="462e2-309">Example</span></span>

```
var itemClass = Office.context.mailbox.item.itemClass;
```

#### <a name="nullable-itemid-string"></a><span data-ttu-id="462e2-310">(nullable) itemId :String</span><span class="sxs-lookup"><span data-stu-id="462e2-310">(nullable) itemId :String</span></span>

<span data-ttu-id="462e2-p117">Ruft den Exchange-Webdienste-Elementbezeichner für das aktuelle Element ab. Nur Lesemodus.</span><span class="sxs-lookup"><span data-stu-id="462e2-p117">Gets the Exchange Web Services item identifier for the current item. Read mode only.</span></span>

> [!NOTE]
> <span data-ttu-id="462e2-313">Der Bezeichner, der von der `itemId`-Eigenschaft zurückgegeben wird, ist der gleiche Bezeichner wie der Exchange-Webdienste-Elementbezeichner.</span><span class="sxs-lookup"><span data-stu-id="462e2-313">The identifier returned by the `itemId` property is the same as the Exchange Web Services item identifier.</span></span> <span data-ttu-id="462e2-314">Die `itemId` -Eigenschaft ist nicht identisch mit der Outlook-Eintrags-ID oder die ID von der Outlook-REST-API verwendet.</span><span class="sxs-lookup"><span data-stu-id="462e2-314">The `itemId` property is not identical to the Outlook Entry ID or the ID used by the Outlook REST API.</span></span> <span data-ttu-id="462e2-315">REST-API-Aufrufe dieser Wert verwenden, sollten sie konvertiert werden, bevor [Office.context.mailbox.convertToRestId](office.context.mailbox.md#converttorestiditemid-restversion--string)verwenden.</span><span class="sxs-lookup"><span data-stu-id="462e2-315">Before making REST API calls using this value, it should be converted using [Office.context.mailbox.convertToRestId](office.context.mailbox.md#converttorestiditemid-restversion--string).</span></span> <span data-ttu-id="462e2-316">Weitere Informationen finden Sie unter [Verwenden der Outlook-REST-APIs aus einem Outlook-add-in](https://docs.microsoft.com/outlook/add-ins/use-rest-api#get-the-item-id).</span><span class="sxs-lookup"><span data-stu-id="462e2-316">For more details, see [Use the Outlook REST APIs from an Outlook add-in](https://docs.microsoft.com/outlook/add-ins/use-rest-api#get-the-item-id).</span></span>

<span data-ttu-id="462e2-p119">Die Eigenschaft `itemId` ist im Modus „Verfassen“ nicht verfügbar. Wenn ein Elementbezeichner erforderlich ist, kann die [`saveAsync`](#saveasyncoptions-callback)-Methode verwendet werden, um das Element zu speichern. Dadurch wird der Elementbezeichner im [`AsyncResult.value`](/javascript/api/office/office.asyncresult)-Parameter in der Rückruffunktion zurückgegeben.</span><span class="sxs-lookup"><span data-stu-id="462e2-p119">The `itemId` property is not available in compose mode. If an item identifier is required, the [`saveAsync`](#saveasyncoptions-callback) method can be used to save the item to the store, which will return the item identifier in the [`AsyncResult.value`](/javascript/api/office/office.asyncresult) parameter in the callback function.</span></span>

##### <a name="type"></a><span data-ttu-id="462e2-319">Typ:</span><span class="sxs-lookup"><span data-stu-id="462e2-319">Type:</span></span>

*   <span data-ttu-id="462e2-320">String</span><span class="sxs-lookup"><span data-stu-id="462e2-320">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="462e2-321">Anforderungen</span><span class="sxs-lookup"><span data-stu-id="462e2-321">Requirements</span></span>

|<span data-ttu-id="462e2-322">Anforderung</span><span class="sxs-lookup"><span data-stu-id="462e2-322">Requirement</span></span>| <span data-ttu-id="462e2-323">Wert</span><span class="sxs-lookup"><span data-stu-id="462e2-323">Value</span></span>|
|---|---|
|[<span data-ttu-id="462e2-324">Mindestversion des Postfachanforderungssatzes</span><span class="sxs-lookup"><span data-stu-id="462e2-324">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="462e2-325">1.0</span><span class="sxs-lookup"><span data-stu-id="462e2-325">1.0</span></span>|
|[<span data-ttu-id="462e2-326">Mindestberechtigungsstufe</span><span class="sxs-lookup"><span data-stu-id="462e2-326">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="462e2-327">ReadItem</span><span class="sxs-lookup"><span data-stu-id="462e2-327">ReadItem</span></span>|
|[<span data-ttu-id="462e2-328">Zutreffender Outlook-Modus</span><span class="sxs-lookup"><span data-stu-id="462e2-328">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="462e2-329">Lesen</span><span class="sxs-lookup"><span data-stu-id="462e2-329">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="462e2-330">Beispiel</span><span class="sxs-lookup"><span data-stu-id="462e2-330">Example</span></span>

<span data-ttu-id="462e2-p120">Mit dem folgende Code wird das Vorhandensein eines Elementbezeichners überprüft. Wenn die `itemId`-Eigenschaft `null` oder `undefined` zurückgibt, speichert sie das Element und ruft den Elementbezeichner aus dem asynchronen Ergebnis ab.</span><span class="sxs-lookup"><span data-stu-id="462e2-p120">The following code checks for the presence of an item identifier. If the `itemId` property returns `null` or `undefined`, it saves the item to the store and gets the item identifier from the asynchronous result.</span></span>

```
var itemId = Office.context.mailbox.item.itemId;
if (itemId === null || itemId == undefined) {
  Office.context.mailbox.item.saveAsync(function(result){
    itemId = result.value;
  });
}
```

####  <a name="itemtype-officemailboxenumsitemtypejavascriptapioutlook14officemailboxenumsitemtype"></a><span data-ttu-id="462e2-333">itemType :[Office.MailboxEnums.ItemType](/javascript/api/outlook_1_4/office.mailboxenums.itemtype)</span><span class="sxs-lookup"><span data-stu-id="462e2-333">itemType :[Office.MailboxEnums.ItemType](/javascript/api/outlook_1_4/office.mailboxenums.itemtype)</span></span>

<span data-ttu-id="462e2-334">Ruft den Typ des Elements ab, das eine Instanz darstellt.</span><span class="sxs-lookup"><span data-stu-id="462e2-334">Gets the type of item that an instance represents.</span></span>

<span data-ttu-id="462e2-335">Die `itemType`-Eigenschaft gibt einen der Werte der `ItemType`-Enumeration zurück, der angibt, ob es sich bei der `item`-Objektinstanz um eine Nachricht oder einen Termin handelt.</span><span class="sxs-lookup"><span data-stu-id="462e2-335">The `itemType` property returns one of the `ItemType` enumeration values, indicating whether the `item` object instance is a message or an appointment.</span></span>

##### <a name="type"></a><span data-ttu-id="462e2-336">Typ:</span><span class="sxs-lookup"><span data-stu-id="462e2-336">Type:</span></span>

*   [<span data-ttu-id="462e2-337">Office.MailboxEnums.ItemType</span><span class="sxs-lookup"><span data-stu-id="462e2-337">Office.MailboxEnums.ItemType</span></span>](/javascript/api/outlook_1_4/office.mailboxenums.itemtype)

##### <a name="requirements"></a><span data-ttu-id="462e2-338">Anforderungen</span><span class="sxs-lookup"><span data-stu-id="462e2-338">Requirements</span></span>

|<span data-ttu-id="462e2-339">Anforderung</span><span class="sxs-lookup"><span data-stu-id="462e2-339">Requirement</span></span>| <span data-ttu-id="462e2-340">Wert</span><span class="sxs-lookup"><span data-stu-id="462e2-340">Value</span></span>|
|---|---|
|[<span data-ttu-id="462e2-341">Mindestversion des Postfachanforderungssatzes</span><span class="sxs-lookup"><span data-stu-id="462e2-341">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="462e2-342">1.0</span><span class="sxs-lookup"><span data-stu-id="462e2-342">1.0</span></span>|
|[<span data-ttu-id="462e2-343">Mindestberechtigungsstufe</span><span class="sxs-lookup"><span data-stu-id="462e2-343">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="462e2-344">ReadItem</span><span class="sxs-lookup"><span data-stu-id="462e2-344">ReadItem</span></span>|
|[<span data-ttu-id="462e2-345">Zutreffender Outlook-Modus</span><span class="sxs-lookup"><span data-stu-id="462e2-345">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="462e2-346">Verfassen oder Lesen</span><span class="sxs-lookup"><span data-stu-id="462e2-346">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="462e2-347">Beispiel</span><span class="sxs-lookup"><span data-stu-id="462e2-347">Example</span></span>

```
if (Office.context.mailbox.item.itemType == Office.MailboxEnums.ItemType.Message)
  // do something
else
  // do something else
```

####  <a name="location-stringlocationjavascriptapioutlook14officelocation"></a><span data-ttu-id="462e2-348">location :String|[Location](/javascript/api/outlook_1_4/office.location)</span><span class="sxs-lookup"><span data-stu-id="462e2-348">location :String|[Location](/javascript/api/outlook_1_4/office.location)</span></span>

<span data-ttu-id="462e2-349">Ruft den Ort eines Termins ab bzw. legt ihn fest.</span><span class="sxs-lookup"><span data-stu-id="462e2-349">Gets or sets the location of an appointment.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="462e2-350">Lesemodus</span><span class="sxs-lookup"><span data-stu-id="462e2-350">Read mode</span></span>

<span data-ttu-id="462e2-351">Die `location`-Eigenschaft gibt eine Zeichenfolge zurück, die den Ort des Termins enthält.</span><span class="sxs-lookup"><span data-stu-id="462e2-351">The `location` property returns a string that contains the location of the appointment.</span></span>

##### <a name="compose-mode"></a><span data-ttu-id="462e2-352">Verfassenmodus</span><span class="sxs-lookup"><span data-stu-id="462e2-352">Compose mode</span></span>

<span data-ttu-id="462e2-353">Die `location`-Eigenschaft gibt ein `Location`-Objekt zurück, das Methoden zum Abrufen und Festlegen des Orts für einen Termin bereitstellt.</span><span class="sxs-lookup"><span data-stu-id="462e2-353">The `location` property returns a `Location` object that provides methods that are used to get and set the location of the appointment.</span></span>

##### <a name="type"></a><span data-ttu-id="462e2-354">Typ:</span><span class="sxs-lookup"><span data-stu-id="462e2-354">Type:</span></span>

*   <span data-ttu-id="462e2-355">String | [Location](/javascript/api/outlook_1_4/office.location)</span><span class="sxs-lookup"><span data-stu-id="462e2-355">String | [Location](/javascript/api/outlook_1_4/office.location)</span></span>

##### <a name="requirements"></a><span data-ttu-id="462e2-356">Anforderungen</span><span class="sxs-lookup"><span data-stu-id="462e2-356">Requirements</span></span>

|<span data-ttu-id="462e2-357">Anforderung</span><span class="sxs-lookup"><span data-stu-id="462e2-357">Requirement</span></span>| <span data-ttu-id="462e2-358">Wert</span><span class="sxs-lookup"><span data-stu-id="462e2-358">Value</span></span>|
|---|---|
|[<span data-ttu-id="462e2-359">Mindestversion des Postfachanforderungssatzes</span><span class="sxs-lookup"><span data-stu-id="462e2-359">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="462e2-360">1.0</span><span class="sxs-lookup"><span data-stu-id="462e2-360">1.0</span></span>|
|[<span data-ttu-id="462e2-361">Mindestberechtigungsstufe</span><span class="sxs-lookup"><span data-stu-id="462e2-361">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="462e2-362">ReadItem</span><span class="sxs-lookup"><span data-stu-id="462e2-362">ReadItem</span></span>|
|[<span data-ttu-id="462e2-363">Zutreffender Outlook-Modus</span><span class="sxs-lookup"><span data-stu-id="462e2-363">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="462e2-364">Verfassen oder Lesen</span><span class="sxs-lookup"><span data-stu-id="462e2-364">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="462e2-365">Beispiel</span><span class="sxs-lookup"><span data-stu-id="462e2-365">Example</span></span>

```
var userContext = { value : 1 };
Office.context.mailbox.item.location.getAsync( { context: userContext}, callback);

function callback(asyncResult) {
  var context = asyncResult.context;
  var location = asyncResult.value;
}
```

#### <a name="normalizedsubject-string"></a><span data-ttu-id="462e2-366">normalizedSubject :String</span><span class="sxs-lookup"><span data-stu-id="462e2-366">normalizedSubject :String</span></span>

<span data-ttu-id="462e2-p121">Ruft den Betreff für ein Element ab, wobei alle Präfixe entfernt werden (einschließlich `RE:` und `FWD:`). Nur Lesemodus.</span><span class="sxs-lookup"><span data-stu-id="462e2-p121">Gets the subject of an item, with all prefixes removed (including `RE:` and `FWD:`). Read mode only.</span></span>

<span data-ttu-id="462e2-p122">Die normalizedSubject-Eigenschaft ruft den Betreff des Elements mit allen Standardpräfixen (z. B. `RE:` und `FW:`) ab, die von E-Mail-Programmen hinzugefügt werden. Wenn der Betreff des Elements mit vollständigen Präfixen abgerufen werden soll, verwenden Sie die [`subject`](#subject-stringsubjectjavascriptapioutlook14officesubject)-Eigenschaft.</span><span class="sxs-lookup"><span data-stu-id="462e2-p122">The normalizedSubject property gets the subject of the item, with any standard prefixes (such as `RE:` and `FW:`) that are added by email programs. To get the subject of the item with the prefixes intact, use the [`subject`](#subject-stringsubjectjavascriptapioutlook14officesubject) property.</span></span>

##### <a name="type"></a><span data-ttu-id="462e2-371">Typ:</span><span class="sxs-lookup"><span data-stu-id="462e2-371">Type:</span></span>

*   <span data-ttu-id="462e2-372">String</span><span class="sxs-lookup"><span data-stu-id="462e2-372">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="462e2-373">Anforderungen</span><span class="sxs-lookup"><span data-stu-id="462e2-373">Requirements</span></span>

|<span data-ttu-id="462e2-374">Anforderung</span><span class="sxs-lookup"><span data-stu-id="462e2-374">Requirement</span></span>| <span data-ttu-id="462e2-375">Wert</span><span class="sxs-lookup"><span data-stu-id="462e2-375">Value</span></span>|
|---|---|
|[<span data-ttu-id="462e2-376">Mindestversion des Postfachanforderungssatzes</span><span class="sxs-lookup"><span data-stu-id="462e2-376">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="462e2-377">1.0</span><span class="sxs-lookup"><span data-stu-id="462e2-377">1.0</span></span>|
|[<span data-ttu-id="462e2-378">Mindestberechtigungsstufe</span><span class="sxs-lookup"><span data-stu-id="462e2-378">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="462e2-379">ReadItem</span><span class="sxs-lookup"><span data-stu-id="462e2-379">ReadItem</span></span>|
|[<span data-ttu-id="462e2-380">Zutreffender Outlook-Modus</span><span class="sxs-lookup"><span data-stu-id="462e2-380">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="462e2-381">Lesen</span><span class="sxs-lookup"><span data-stu-id="462e2-381">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="462e2-382">Beispiel</span><span class="sxs-lookup"><span data-stu-id="462e2-382">Example</span></span>

```
var normalizedSubject = Office.context.mailbox.item.normalizedSubject;
```

####  <a name="notificationmessages-notificationmessagesjavascriptapioutlook14officenotificationmessages"></a><span data-ttu-id="462e2-383">notificationMessages :[NotificationMessages](/javascript/api/outlook_1_4/office.notificationmessages)</span><span class="sxs-lookup"><span data-stu-id="462e2-383">notificationMessages :[NotificationMessages](/javascript/api/outlook_1_4/office.notificationmessages)</span></span>

<span data-ttu-id="462e2-384">Ruft die Benachrichtigungen für ein Element ab.</span><span class="sxs-lookup"><span data-stu-id="462e2-384">Gets the notification messages for an item.</span></span>

##### <a name="type"></a><span data-ttu-id="462e2-385">Typ:</span><span class="sxs-lookup"><span data-stu-id="462e2-385">Type:</span></span>

*   [<span data-ttu-id="462e2-386">NotificationMessages</span><span class="sxs-lookup"><span data-stu-id="462e2-386">NotificationMessages</span></span>](/javascript/api/outlook_1_4/office.notificationmessages)

##### <a name="requirements"></a><span data-ttu-id="462e2-387">Anforderungen</span><span class="sxs-lookup"><span data-stu-id="462e2-387">Requirements</span></span>

|<span data-ttu-id="462e2-388">Anforderung</span><span class="sxs-lookup"><span data-stu-id="462e2-388">Requirement</span></span>| <span data-ttu-id="462e2-389">Wert</span><span class="sxs-lookup"><span data-stu-id="462e2-389">Value</span></span>|
|---|---|
|[<span data-ttu-id="462e2-390">Mindestversion des Postfachanforderungssatzes</span><span class="sxs-lookup"><span data-stu-id="462e2-390">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="462e2-391">1.3</span><span class="sxs-lookup"><span data-stu-id="462e2-391">1.3</span></span>|
|[<span data-ttu-id="462e2-392">Mindestberechtigungsstufe</span><span class="sxs-lookup"><span data-stu-id="462e2-392">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="462e2-393">ReadItem</span><span class="sxs-lookup"><span data-stu-id="462e2-393">ReadItem</span></span>|
|[<span data-ttu-id="462e2-394">Zutreffender Outlook-Modus</span><span class="sxs-lookup"><span data-stu-id="462e2-394">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="462e2-395">Verfassen oder Lesen</span><span class="sxs-lookup"><span data-stu-id="462e2-395">Compose or read</span></span>|

####  <a name="optionalattendees-arrayemailaddressdetailsjavascriptapioutlook14officeemailaddressdetailsrecipientsjavascriptapioutlook14officerecipients"></a><span data-ttu-id="462e2-396">optionalAttendees :Array.<[EmailAddressDetails](/javascript/api/outlook_1_4/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook_1_4/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="462e2-396">optionalAttendees :Array.<[EmailAddressDetails](/javascript/api/outlook_1_4/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook_1_4/office.recipients)</span></span>

<span data-ttu-id="462e2-397">Ermöglicht den Zugriff auf die optionalen Teilnehmer eines Ereignisses.</span><span class="sxs-lookup"><span data-stu-id="462e2-397">Provides access to the optional attendees of an event.</span></span> <span data-ttu-id="462e2-398">Der Typ des Objekts und die Zugriffsebene, hängt von den Modus des aktuellen Elements ab.</span><span class="sxs-lookup"><span data-stu-id="462e2-398">The type of object and level of access depends on the mode of the current item.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="462e2-399">Lesemodus</span><span class="sxs-lookup"><span data-stu-id="462e2-399">Read mode</span></span>

<span data-ttu-id="462e2-400">Die `optionalAttendees`-Eigenschaft gibt ein Array mit einem `EmailAddressDetails`-Objekt für jeden optionalen Teilnehmer an der Besprechung zurück.</span><span class="sxs-lookup"><span data-stu-id="462e2-400">The `optionalAttendees` property returns an array that contains an `EmailAddressDetails` object for each optional attendee to the meeting.</span></span>

##### <a name="compose-mode"></a><span data-ttu-id="462e2-401">Verfassenmodus</span><span class="sxs-lookup"><span data-stu-id="462e2-401">Compose mode</span></span>

<span data-ttu-id="462e2-402">Die `optionalAttendees` -Eigenschaft gibt eine `Recipients` -Objekt, das Methoden zum Abrufen oder Aktualisieren der optionalen Teilnehmer für eine Besprechung bereitstellt.</span><span class="sxs-lookup"><span data-stu-id="462e2-402">The `optionalAttendees` property returns a `Recipients` object that provides methods to get or update the optional attendees for a meeting.</span></span>

##### <a name="type"></a><span data-ttu-id="462e2-403">Typ:</span><span class="sxs-lookup"><span data-stu-id="462e2-403">Type:</span></span>

*   <span data-ttu-id="462e2-404">Array.<[EmailAddressDetails](/javascript/api/outlook_1_4/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook_1_4/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="462e2-404">Array.<[EmailAddressDetails](/javascript/api/outlook_1_4/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook_1_4/office.recipients)</span></span>

##### <a name="requirements"></a><span data-ttu-id="462e2-405">Anforderungen</span><span class="sxs-lookup"><span data-stu-id="462e2-405">Requirements</span></span>

|<span data-ttu-id="462e2-406">Anforderung</span><span class="sxs-lookup"><span data-stu-id="462e2-406">Requirement</span></span>| <span data-ttu-id="462e2-407">Wert</span><span class="sxs-lookup"><span data-stu-id="462e2-407">Value</span></span>|
|---|---|
|[<span data-ttu-id="462e2-408">Mindestversion des Postfachanforderungssatzes</span><span class="sxs-lookup"><span data-stu-id="462e2-408">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="462e2-409">1.0</span><span class="sxs-lookup"><span data-stu-id="462e2-409">1.0</span></span>|
|[<span data-ttu-id="462e2-410">Mindestberechtigungsstufe</span><span class="sxs-lookup"><span data-stu-id="462e2-410">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="462e2-411">ReadItem</span><span class="sxs-lookup"><span data-stu-id="462e2-411">ReadItem</span></span>|
|[<span data-ttu-id="462e2-412">Zutreffender Outlook-Modus</span><span class="sxs-lookup"><span data-stu-id="462e2-412">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="462e2-413">Verfassen oder Lesen</span><span class="sxs-lookup"><span data-stu-id="462e2-413">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="462e2-414">Beispiel</span><span class="sxs-lookup"><span data-stu-id="462e2-414">Example</span></span>

```
Office.context.mailbox.item.optionalAttendees.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.optionalAttendees.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.optionalAttendees.getAsync(callback);

function callback(asyncResult) {
  var arrayOfOptionalAttendeesRecipients = asyncResult.value;
}
```

#### <a name="organizer-emailaddressdetailsjavascriptapioutlook14officeemailaddressdetails"></a><span data-ttu-id="462e2-415">organizer :[EmailAddressDetails](/javascript/api/outlook_1_4/office.emailaddressdetails)</span><span class="sxs-lookup"><span data-stu-id="462e2-415">organizer :[EmailAddressDetails](/javascript/api/outlook_1_4/office.emailaddressdetails)</span></span>

<span data-ttu-id="462e2-p124">Ruft für eine angegebene Besprechung die E-Mail-Adresse des Organisators der Besprechung ab. Nur Lesemodus.</span><span class="sxs-lookup"><span data-stu-id="462e2-p124">Gets the email address of the meeting organizer for a specified meeting. Read mode only.</span></span>

##### <a name="type"></a><span data-ttu-id="462e2-418">Typ:</span><span class="sxs-lookup"><span data-stu-id="462e2-418">Type:</span></span>

*   [<span data-ttu-id="462e2-419">EmailAddressDetails</span><span class="sxs-lookup"><span data-stu-id="462e2-419">EmailAddressDetails</span></span>](/javascript/api/outlook_1_4/office.emailaddressdetails)

##### <a name="requirements"></a><span data-ttu-id="462e2-420">Anforderungen</span><span class="sxs-lookup"><span data-stu-id="462e2-420">Requirements</span></span>

|<span data-ttu-id="462e2-421">Anforderung</span><span class="sxs-lookup"><span data-stu-id="462e2-421">Requirement</span></span>| <span data-ttu-id="462e2-422">Wert</span><span class="sxs-lookup"><span data-stu-id="462e2-422">Value</span></span>|
|---|---|
|[<span data-ttu-id="462e2-423">Mindestversion des Postfachanforderungssatzes</span><span class="sxs-lookup"><span data-stu-id="462e2-423">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="462e2-424">1.0</span><span class="sxs-lookup"><span data-stu-id="462e2-424">1.0</span></span>|
|[<span data-ttu-id="462e2-425">Mindestberechtigungsstufe</span><span class="sxs-lookup"><span data-stu-id="462e2-425">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="462e2-426">ReadItem</span><span class="sxs-lookup"><span data-stu-id="462e2-426">ReadItem</span></span>|
|[<span data-ttu-id="462e2-427">Zutreffender Outlook-Modus</span><span class="sxs-lookup"><span data-stu-id="462e2-427">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="462e2-428">Lesen</span><span class="sxs-lookup"><span data-stu-id="462e2-428">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="462e2-429">Beispiel</span><span class="sxs-lookup"><span data-stu-id="462e2-429">Example</span></span>

```
var organizerName = Office.context.mailbox.item.organizer.displayName;
var organizerAddress = Office.context.mailbox.item.organizer.emailAddress;
```

####  <a name="requiredattendees-arrayemailaddressdetailsjavascriptapioutlook14officeemailaddressdetailsrecipientsjavascriptapioutlook14officerecipients"></a><span data-ttu-id="462e2-430">requiredAttendees :Array.<[EmailAddressDetails](/javascript/api/outlook_1_4/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook_1_4/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="462e2-430">requiredAttendees :Array.<[EmailAddressDetails](/javascript/api/outlook_1_4/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook_1_4/office.recipients)</span></span>

<span data-ttu-id="462e2-431">Ermöglicht den Zugriff auf die erforderlichen Teilnehmer eines Ereignisses.</span><span class="sxs-lookup"><span data-stu-id="462e2-431">Provides access to the required attendees of an event.</span></span> <span data-ttu-id="462e2-432">Der Typ des Objekts und die Zugriffsebene, hängt von den Modus des aktuellen Elements ab.</span><span class="sxs-lookup"><span data-stu-id="462e2-432">The type of object and level of access depends on the mode of the current item.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="462e2-433">Lesemodus</span><span class="sxs-lookup"><span data-stu-id="462e2-433">Read mode</span></span>

<span data-ttu-id="462e2-434">Die `requiredAttendees`-Eigenschaft gibt ein Array mit einem `EmailAddressDetails`-Objekt für jeden erforderlichen Teilnehmer an der Besprechung zurück.</span><span class="sxs-lookup"><span data-stu-id="462e2-434">The `requiredAttendees` property returns an array that contains an `EmailAddressDetails` object for each required attendee to the meeting.</span></span>

##### <a name="compose-mode"></a><span data-ttu-id="462e2-435">Verfassenmodus</span><span class="sxs-lookup"><span data-stu-id="462e2-435">Compose mode</span></span>

<span data-ttu-id="462e2-436">Die `requiredAttendees` -Eigenschaft gibt eine `Recipients` -Objekt, das Methoden zum Abrufen oder Aktualisieren der erforderlichen Teilnehmer für eine Besprechung bereitstellt.</span><span class="sxs-lookup"><span data-stu-id="462e2-436">The `requiredAttendees` property returns a `Recipients` object that provides methods to get or update the required attendees for a meeting.</span></span>

##### <a name="type"></a><span data-ttu-id="462e2-437">Typ:</span><span class="sxs-lookup"><span data-stu-id="462e2-437">Type:</span></span>

*   <span data-ttu-id="462e2-438">Array.<[EmailAddressDetails](/javascript/api/outlook_1_4/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook_1_4/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="462e2-438">Array.<[EmailAddressDetails](/javascript/api/outlook_1_4/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook_1_4/office.recipients)</span></span>

##### <a name="requirements"></a><span data-ttu-id="462e2-439">Anforderungen</span><span class="sxs-lookup"><span data-stu-id="462e2-439">Requirements</span></span>

|<span data-ttu-id="462e2-440">Anforderung</span><span class="sxs-lookup"><span data-stu-id="462e2-440">Requirement</span></span>| <span data-ttu-id="462e2-441">Wert</span><span class="sxs-lookup"><span data-stu-id="462e2-441">Value</span></span>|
|---|---|
|[<span data-ttu-id="462e2-442">Mindestversion des Postfachanforderungssatzes</span><span class="sxs-lookup"><span data-stu-id="462e2-442">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="462e2-443">1.0</span><span class="sxs-lookup"><span data-stu-id="462e2-443">1.0</span></span>|
|[<span data-ttu-id="462e2-444">Mindestberechtigungsstufe</span><span class="sxs-lookup"><span data-stu-id="462e2-444">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="462e2-445">ReadItem</span><span class="sxs-lookup"><span data-stu-id="462e2-445">ReadItem</span></span>|
|[<span data-ttu-id="462e2-446">Zutreffender Outlook-Modus</span><span class="sxs-lookup"><span data-stu-id="462e2-446">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="462e2-447">Verfassen oder Lesen</span><span class="sxs-lookup"><span data-stu-id="462e2-447">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="462e2-448">Beispiel</span><span class="sxs-lookup"><span data-stu-id="462e2-448">Example</span></span>

```
Office.context.mailbox.item.requiredAttendees.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.requiredAttendees.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.requiredAttendees.getAsync(callback);

function callback(asyncResult) {
  var arrayOfRequiredAttendeesRecipients = asyncResult.value;
}
```

#### <a name="sender-emailaddressdetailsjavascriptapioutlook14officeemailaddressdetails"></a><span data-ttu-id="462e2-449">sender :[EmailAddressDetails](/javascript/api/outlook_1_4/office.emailaddressdetails)</span><span class="sxs-lookup"><span data-stu-id="462e2-449">sender :[EmailAddressDetails](/javascript/api/outlook_1_4/office.emailaddressdetails)</span></span>

<span data-ttu-id="462e2-p126">Ruft die E-Mail-Adresse des Absenders einer E-Mail-Nachricht ab. Nur Lesemodus.</span><span class="sxs-lookup"><span data-stu-id="462e2-p126">Gets the email address of the sender of an email message. Read mode only.</span></span>

<span data-ttu-id="462e2-p127">Die [`from`](#from-emailaddressdetailsjavascriptapioutlook14officeemailaddressdetails)- und `sender`-Eigenschaften stellen dieselbe Person dar, außer die Nachricht wurde von einem Delegaten gesendet. In diesem Fall stellt die `from`-Eigenschaft die Stellvertretung dar, und die sender-Eigenschaft stellt den Delegaten dar.</span><span class="sxs-lookup"><span data-stu-id="462e2-p127">The [`from`](#from-emailaddressdetailsjavascriptapioutlook14officeemailaddressdetails) and `sender` properties represent the same person unless the message is sent by a delegate. In that case, the `from` property represents the delegator, and the sender property represents the delegate.</span></span>

> [!NOTE]
> <span data-ttu-id="462e2-454">Die `recipientType` -Eigenschaft des der `EmailAddressDetails` -Objekts der `sender` -Eigenschaft ist `undefined`.</span><span class="sxs-lookup"><span data-stu-id="462e2-454">The `recipientType` property of the `EmailAddressDetails` object in the `sender` property is `undefined`.</span></span>

##### <a name="type"></a><span data-ttu-id="462e2-455">Typ:</span><span class="sxs-lookup"><span data-stu-id="462e2-455">Type:</span></span>

*   [<span data-ttu-id="462e2-456">EmailAddressDetails</span><span class="sxs-lookup"><span data-stu-id="462e2-456">EmailAddressDetails</span></span>](/javascript/api/outlook_1_4/office.emailaddressdetails)

##### <a name="requirements"></a><span data-ttu-id="462e2-457">Anforderungen</span><span class="sxs-lookup"><span data-stu-id="462e2-457">Requirements</span></span>

|<span data-ttu-id="462e2-458">Anforderung</span><span class="sxs-lookup"><span data-stu-id="462e2-458">Requirement</span></span>| <span data-ttu-id="462e2-459">Wert</span><span class="sxs-lookup"><span data-stu-id="462e2-459">Value</span></span>|
|---|---|
|[<span data-ttu-id="462e2-460">Mindestversion des Postfachanforderungssatzes</span><span class="sxs-lookup"><span data-stu-id="462e2-460">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="462e2-461">1.0</span><span class="sxs-lookup"><span data-stu-id="462e2-461">1.0</span></span>|
|[<span data-ttu-id="462e2-462">Mindestberechtigungsstufe</span><span class="sxs-lookup"><span data-stu-id="462e2-462">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="462e2-463">ReadItem</span><span class="sxs-lookup"><span data-stu-id="462e2-463">ReadItem</span></span>|
|[<span data-ttu-id="462e2-464">Zutreffender Outlook-Modus</span><span class="sxs-lookup"><span data-stu-id="462e2-464">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="462e2-465">Lesen</span><span class="sxs-lookup"><span data-stu-id="462e2-465">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="462e2-466">Beispiel</span><span class="sxs-lookup"><span data-stu-id="462e2-466">Example</span></span>

```
var senderName = Office.context.mailbox.item.sender.displayName;
var senderAddress = Office.context.mailbox.item.sender.emailAddress;
```

####  <a name="start-datetimejavascriptapioutlook14officetime"></a><span data-ttu-id="462e2-467">start :Date|[Time](/javascript/api/outlook_1_4/office.time)</span><span class="sxs-lookup"><span data-stu-id="462e2-467">start :Date|[Time](/javascript/api/outlook_1_4/office.time)</span></span>

<span data-ttu-id="462e2-468">Ruft Datum und Zeit für den Beginn des Termins ab oder legt Datum und Uhrzeit fest.</span><span class="sxs-lookup"><span data-stu-id="462e2-468">Gets or sets the date and time that the appointment is to begin.</span></span>

<span data-ttu-id="462e2-p128">Die `start`-Eigenschaft wird als Datums- und Uhrzeitwert in UTC ausgedrückt. Sie können die [`convertToLocalClientTime`](office.context.mailbox.md#converttolocalclienttimetimevalue--localclienttimejavascriptapioutlook14officelocalclienttime)-Methode verwenden, um den Wert in den lokalen Uhrzeit- und Datumswert des Clients umzuwandeln.</span><span class="sxs-lookup"><span data-stu-id="462e2-p128">The `start` property is expressed as a Coordinated Universal Time (UTC) date and time value. You can use the [`convertToLocalClientTime`](office.context.mailbox.md#converttolocalclienttimetimevalue--localclienttimejavascriptapioutlook14officelocalclienttime) method to convert the value to the client’s local date and time.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="462e2-471">Lesemodus</span><span class="sxs-lookup"><span data-stu-id="462e2-471">Read mode</span></span>

<span data-ttu-id="462e2-472">Die `start`-Eigenschaft gibt ein `Date`-Objekt zurück.</span><span class="sxs-lookup"><span data-stu-id="462e2-472">The `start` property returns a `Date` object.</span></span>

##### <a name="compose-mode"></a><span data-ttu-id="462e2-473">Verfassenmodus</span><span class="sxs-lookup"><span data-stu-id="462e2-473">Compose mode</span></span>

<span data-ttu-id="462e2-474">Die `start`-Eigenschaft gibt ein `Time`-Objekt zurück.</span><span class="sxs-lookup"><span data-stu-id="462e2-474">The `start` property returns a `Time` object.</span></span>

<span data-ttu-id="462e2-475">Wenn Sie die [`Time.setAsync`](/javascript/api/outlook_1_4/office.time#setasync-datetime--options--callback-)-Methode verwenden, um die Startzeit im Verfassenmodus festzulegen, sollten Sie die [`convertToUtcClientTime`](office.context.mailbox.md#converttoutcclienttimeinput--date)-Methode verwenden, um die Ortszeit auf dem Client für den Server in UTC umzuwandeln.</span><span class="sxs-lookup"><span data-stu-id="462e2-475">When you use the [`Time.setAsync`](/javascript/api/outlook_1_4/office.time#setasync-datetime--options--callback-) method to set the start time, you should use the [`convertToUtcClientTime`](office.context.mailbox.md#converttoutcclienttimeinput--date) method to convert the local time on the client to UTC for the server.</span></span>

##### <a name="type"></a><span data-ttu-id="462e2-476">Typ:</span><span class="sxs-lookup"><span data-stu-id="462e2-476">Type:</span></span>

*   <span data-ttu-id="462e2-477">Date | [Time](/javascript/api/outlook_1_4/office.time)</span><span class="sxs-lookup"><span data-stu-id="462e2-477">Date | [Time](/javascript/api/outlook_1_4/office.time)</span></span>

##### <a name="requirements"></a><span data-ttu-id="462e2-478">Anforderungen</span><span class="sxs-lookup"><span data-stu-id="462e2-478">Requirements</span></span>

|<span data-ttu-id="462e2-479">Anforderung</span><span class="sxs-lookup"><span data-stu-id="462e2-479">Requirement</span></span>| <span data-ttu-id="462e2-480">Wert</span><span class="sxs-lookup"><span data-stu-id="462e2-480">Value</span></span>|
|---|---|
|[<span data-ttu-id="462e2-481">Mindestversion des Postfachanforderungssatzes</span><span class="sxs-lookup"><span data-stu-id="462e2-481">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="462e2-482">1.0</span><span class="sxs-lookup"><span data-stu-id="462e2-482">1.0</span></span>|
|[<span data-ttu-id="462e2-483">Mindestberechtigungsstufe</span><span class="sxs-lookup"><span data-stu-id="462e2-483">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="462e2-484">ReadItem</span><span class="sxs-lookup"><span data-stu-id="462e2-484">ReadItem</span></span>|
|[<span data-ttu-id="462e2-485">Zutreffender Outlook-Modus</span><span class="sxs-lookup"><span data-stu-id="462e2-485">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="462e2-486">Verfassen oder Lesen</span><span class="sxs-lookup"><span data-stu-id="462e2-486">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="462e2-487">Beispiel</span><span class="sxs-lookup"><span data-stu-id="462e2-487">Example</span></span>

<span data-ttu-id="462e2-488">Im folgenden Beispiel wird die Startzeit eines Termins im Verfassenmodus mithilfe der [`setAsync`](/javascript/api/outlook_1_4/office.time#setasync-datetime--options--callback-)-Methode des `Time`-Objekts festgelegt.</span><span class="sxs-lookup"><span data-stu-id="462e2-488">The following example sets the start time of an appointment in compose mode by using the [`setAsync`](/javascript/api/outlook_1_4/office.time#setasync-datetime--options--callback-) method of the `Time` object.</span></span>

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

####  <a name="subject-stringsubjectjavascriptapioutlook14officesubject"></a><span data-ttu-id="462e2-489">subject :String|[Subject](/javascript/api/outlook_1_4/office.subject)</span><span class="sxs-lookup"><span data-stu-id="462e2-489">subject :String|[Subject](/javascript/api/outlook_1_4/office.subject)</span></span>

<span data-ttu-id="462e2-490">Ruft die Beschreibung ab, die im Betrefffeld eines Elements angezeigt wird, oder legt sie fest.</span><span class="sxs-lookup"><span data-stu-id="462e2-490">Gets or sets the description that appears in the subject field of an item.</span></span>

<span data-ttu-id="462e2-491">Die `subject`-Eigenschaft ruft den gesamten Betreff des Elements ab oder legt ihn fest – so, wie er vom E-Mail-Server gesendet wird.</span><span class="sxs-lookup"><span data-stu-id="462e2-491">The `subject` property gets or sets the entire subject of the item, as sent by the email server.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="462e2-492">Lesemodus</span><span class="sxs-lookup"><span data-stu-id="462e2-492">Read mode</span></span>

<span data-ttu-id="462e2-p129">Die `subject`-Eigenschaft gibt eine Zeichenfolge zurück. Verwenden Sie die [`normalizedSubject`](#normalizedsubject-string)-Eigenschaft, um den Betreff ohne führende Präfixe wie `RE:` und `FW:` abzurufen.</span><span class="sxs-lookup"><span data-stu-id="462e2-p129">The `subject` property returns a string. Use the [`normalizedSubject`](#normalizedsubject-string) property to get the subject minus any leading prefixes such as `RE:` and `FW:`.</span></span>

```
var subject = Office.context.mailbox.item.subject;
```

##### <a name="compose-mode"></a><span data-ttu-id="462e2-495">Verfassenmodus</span><span class="sxs-lookup"><span data-stu-id="462e2-495">Compose mode</span></span>

<span data-ttu-id="462e2-496">Die `subject`-Eigenschaft gibt ein `Subject`-Objekt zurück, das Methoden zum Abrufen und Festlegen des Betreffs bereitstellt.</span><span class="sxs-lookup"><span data-stu-id="462e2-496">The `subject` property returns a `Subject` object that provides methods to get and set the subject.</span></span>

```
Office.context.mailbox.item.subject.getAsync(callback);

function callback(asyncResult) {
  var subject = asyncResult.value;
}
```

##### <a name="type"></a><span data-ttu-id="462e2-497">Typ:</span><span class="sxs-lookup"><span data-stu-id="462e2-497">Type:</span></span>

*   <span data-ttu-id="462e2-498">String | [Subject](/javascript/api/outlook_1_4/office.subject)</span><span class="sxs-lookup"><span data-stu-id="462e2-498">String | [Subject](/javascript/api/outlook_1_4/office.subject)</span></span>

##### <a name="requirements"></a><span data-ttu-id="462e2-499">Anforderungen</span><span class="sxs-lookup"><span data-stu-id="462e2-499">Requirements</span></span>

|<span data-ttu-id="462e2-500">Anforderung</span><span class="sxs-lookup"><span data-stu-id="462e2-500">Requirement</span></span>| <span data-ttu-id="462e2-501">Wert</span><span class="sxs-lookup"><span data-stu-id="462e2-501">Value</span></span>|
|---|---|
|[<span data-ttu-id="462e2-502">Mindestversion des Postfachanforderungssatzes</span><span class="sxs-lookup"><span data-stu-id="462e2-502">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="462e2-503">1.0</span><span class="sxs-lookup"><span data-stu-id="462e2-503">1.0</span></span>|
|[<span data-ttu-id="462e2-504">Mindestberechtigungsstufe</span><span class="sxs-lookup"><span data-stu-id="462e2-504">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="462e2-505">ReadItem</span><span class="sxs-lookup"><span data-stu-id="462e2-505">ReadItem</span></span>|
|[<span data-ttu-id="462e2-506">Zutreffender Outlook-Modus</span><span class="sxs-lookup"><span data-stu-id="462e2-506">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="462e2-507">Verfassen oder Lesen</span><span class="sxs-lookup"><span data-stu-id="462e2-507">Compose or read</span></span>|

####  <a name="to-arrayemailaddressdetailsjavascriptapioutlook14officeemailaddressdetailsrecipientsjavascriptapioutlook14officerecipients"></a><span data-ttu-id="462e2-508">an: Array. <[EmailAddressDetails](/javascript/api/outlook_1_4/office.emailaddressdetails)>|[Empfänger](/javascript/api/outlook_1_4/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="462e2-508">to :Array.<[EmailAddressDetails](/javascript/api/outlook_1_4/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook_1_4/office.recipients)</span></span>

<span data-ttu-id="462e2-509">Ermöglicht den Zugriff auf die Empfänger in der Zeile **an** einer Nachricht.</span><span class="sxs-lookup"><span data-stu-id="462e2-509">Provides access to the recipients on the **To** line of a message.</span></span> <span data-ttu-id="462e2-510">Der Typ des Objekts und die Zugriffsebene, hängt von den Modus des aktuellen Elements ab.</span><span class="sxs-lookup"><span data-stu-id="462e2-510">The type of object and level of access depends on the mode of the current item.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="462e2-511">Lesemodus</span><span class="sxs-lookup"><span data-stu-id="462e2-511">Read mode</span></span>

<span data-ttu-id="462e2-p131">Die `to`-Eigenschaft gibt ein Array mit einem `EmailAddressDetails`-Objekt für jeden Empfänger in der**An**-Zeile der Nachricht zurück. Die Auflistung ist auf höchstens 100 Elemente beschränkt.</span><span class="sxs-lookup"><span data-stu-id="462e2-p131">The `to` property returns an array that contains an `EmailAddressDetails` object for each recipient listed on the **To** line of the message. The collection is limited to a maximum of 100 members.</span></span>

##### <a name="compose-mode"></a><span data-ttu-id="462e2-514">Verfassenmodus</span><span class="sxs-lookup"><span data-stu-id="462e2-514">Compose mode</span></span>

<span data-ttu-id="462e2-515">Die `to` -Eigenschaft gibt eine `Recipients` -Objekt, das Methoden zum Abrufen oder aktualisieren die Empfänger in der Zeile **an** der Nachricht bereitstellt.</span><span class="sxs-lookup"><span data-stu-id="462e2-515">The `to` property returns a `Recipients` object that provides methods to get or update the recipients on the **To** line of the message.</span></span>

##### <a name="type"></a><span data-ttu-id="462e2-516">Typ:</span><span class="sxs-lookup"><span data-stu-id="462e2-516">Type:</span></span>

*   <span data-ttu-id="462e2-517">Array.<[EmailAddressDetails](/javascript/api/outlook_1_4/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook_1_4/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="462e2-517">Array.<[EmailAddressDetails](/javascript/api/outlook_1_4/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook_1_4/office.recipients)</span></span>

##### <a name="requirements"></a><span data-ttu-id="462e2-518">Anforderungen</span><span class="sxs-lookup"><span data-stu-id="462e2-518">Requirements</span></span>

|<span data-ttu-id="462e2-519">Anforderung</span><span class="sxs-lookup"><span data-stu-id="462e2-519">Requirement</span></span>| <span data-ttu-id="462e2-520">Wert</span><span class="sxs-lookup"><span data-stu-id="462e2-520">Value</span></span>|
|---|---|
|[<span data-ttu-id="462e2-521">Mindestversion des Postfachanforderungssatzes</span><span class="sxs-lookup"><span data-stu-id="462e2-521">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="462e2-522">1.0</span><span class="sxs-lookup"><span data-stu-id="462e2-522">1.0</span></span>|
|[<span data-ttu-id="462e2-523">Mindestberechtigungsstufe</span><span class="sxs-lookup"><span data-stu-id="462e2-523">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="462e2-524">ReadItem</span><span class="sxs-lookup"><span data-stu-id="462e2-524">ReadItem</span></span>|
|[<span data-ttu-id="462e2-525">Zutreffender Outlook-Modus</span><span class="sxs-lookup"><span data-stu-id="462e2-525">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="462e2-526">Verfassen oder Lesen</span><span class="sxs-lookup"><span data-stu-id="462e2-526">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="462e2-527">Beispiel</span><span class="sxs-lookup"><span data-stu-id="462e2-527">Example</span></span>

```
Office.context.mailbox.item.to.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.to.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.to.getAsync(callback);

function callback(asyncResult) {
  var arrayOfToRecipients = asyncResult.value;
}
```

### <a name="methods"></a><span data-ttu-id="462e2-528">Methoden</span><span class="sxs-lookup"><span data-stu-id="462e2-528">Methods</span></span>

####  <a name="addfileattachmentasyncuri-attachmentname-options-callback"></a><span data-ttu-id="462e2-529">addFileAttachmentAsync(uri, attachmentName, [options], [callback])</span><span class="sxs-lookup"><span data-stu-id="462e2-529">addFileAttachmentAsync(uri, attachmentName, [options], [callback])</span></span>

<span data-ttu-id="462e2-530">Fügt eine Datei zu einer Nachricht oder einem Termin als Anlage hinzu.</span><span class="sxs-lookup"><span data-stu-id="462e2-530">Adds a file to a message or appointment as an attachment.</span></span>

<span data-ttu-id="462e2-531">Die `addFileAttachmentAsync`-Methode lädt die Datei am angegebenen URI hoch und fügt sie an das Element im Verfassenformular an.</span><span class="sxs-lookup"><span data-stu-id="462e2-531">The `addFileAttachmentAsync` method uploads the file at the specified URI and attaches it to the item in the compose form.</span></span>

<span data-ttu-id="462e2-532">Anschließend können Sie den Bezeichner mit der [`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback)-Methode in der gleichen Sitzung zum Entfernen der Anlage verwenden.</span><span class="sxs-lookup"><span data-stu-id="462e2-532">You can subsequently use the identifier with the [`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback) method to remove the attachment in the same session.</span></span>

##### <a name="parameters"></a><span data-ttu-id="462e2-533">Parameter:</span><span class="sxs-lookup"><span data-stu-id="462e2-533">Parameters:</span></span>

|<span data-ttu-id="462e2-534">Name</span><span class="sxs-lookup"><span data-stu-id="462e2-534">Name</span></span>| <span data-ttu-id="462e2-535">Typ</span><span class="sxs-lookup"><span data-stu-id="462e2-535">Type</span></span>| <span data-ttu-id="462e2-536">Attribute</span><span class="sxs-lookup"><span data-stu-id="462e2-536">Attributes</span></span>| <span data-ttu-id="462e2-537">Beschreibung</span><span class="sxs-lookup"><span data-stu-id="462e2-537">Description</span></span>|
|---|---|---|---|
|`uri`| <span data-ttu-id="462e2-538">Zeichenfolge</span><span class="sxs-lookup"><span data-stu-id="462e2-538">String</span></span>||<span data-ttu-id="462e2-p132">Der URI, der den Speicherort der an die Nachricht oder den Termin anzuhängenden Datei angibt. Die maximale Länge ist 2048 Zeichen.</span><span class="sxs-lookup"><span data-stu-id="462e2-p132">The URI that provides the location of the file to attach to the message or appointment. The maximum length is 2048 characters.</span></span>|
|`attachmentName`| <span data-ttu-id="462e2-541">Zeichenfolge</span><span class="sxs-lookup"><span data-stu-id="462e2-541">String</span></span>||<span data-ttu-id="462e2-p133">Der Name der Anlage, der beim Hochladen der Anlage angezeigt wird. Die maximale Länge ist 255 Zeichen.</span><span class="sxs-lookup"><span data-stu-id="462e2-p133">The name of the attachment that is shown while the attachment is uploading. The maximum length is 255 characters.</span></span>|
|`options`| <span data-ttu-id="462e2-544">Object</span><span class="sxs-lookup"><span data-stu-id="462e2-544">Object</span></span>| <span data-ttu-id="462e2-545">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="462e2-545">&lt;optional&gt;</span></span>|<span data-ttu-id="462e2-546">Ein Objektliteral, das mindestens eine der folgenden Eigenschaften enthält.</span><span class="sxs-lookup"><span data-stu-id="462e2-546">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`| <span data-ttu-id="462e2-547">Object</span><span class="sxs-lookup"><span data-stu-id="462e2-547">Object</span></span>| <span data-ttu-id="462e2-548">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="462e2-548">&lt;optional&gt;</span></span>|<span data-ttu-id="462e2-549">Entwickler können ein Objekt bereitstellen, auf das sie in der Callbackmethode zugreifen möchten.</span><span class="sxs-lookup"><span data-stu-id="462e2-549">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`| <span data-ttu-id="462e2-550">Funktion</span><span class="sxs-lookup"><span data-stu-id="462e2-550">function</span></span>| <span data-ttu-id="462e2-551">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="462e2-551">&lt;optional&gt;</span></span>|<span data-ttu-id="462e2-552">Wenn die Methode abgeschlossen wird, wird die Funktion , die im `callback`-Parameter übergeben wird, mit einem einzelnen Parameter (`asyncResult`) aufgerufen, der ein [`AsyncResult`](/javascript/api/office/office.asyncresult)-Objekt darstellt.</span><span class="sxs-lookup"><span data-stu-id="462e2-552">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span> <br/><span data-ttu-id="462e2-553">Bei Erfolg wird der Anlagenbezeichner in der `asyncResult.value`-Eigenschaft bereitgestellt.</span><span class="sxs-lookup"><span data-stu-id="462e2-553">On success, the attachment identifier will be provided in the `asyncResult.value` property.</span></span><br/><span data-ttu-id="462e2-554">Wenn beim Hochladen der Anlage ein Fehler auftritt, enthält das `asyncResult`-Objekt ein `Error`-Objekt mit einer Beschreibung des Fehlers.</span><span class="sxs-lookup"><span data-stu-id="462e2-554">If uploading the attachment fails, the `asyncResult` object will contain an `Error` object that provides a description of the error.</span></span>|

##### <a name="errors"></a><span data-ttu-id="462e2-555">Fehler</span><span class="sxs-lookup"><span data-stu-id="462e2-555">Errors</span></span>

| <span data-ttu-id="462e2-556">Fehlercode</span><span class="sxs-lookup"><span data-stu-id="462e2-556">Error code</span></span> | <span data-ttu-id="462e2-557">Beschreibung</span><span class="sxs-lookup"><span data-stu-id="462e2-557">Description</span></span> |
|------------|-------------|
| `AttachmentSizeExceeded` | <span data-ttu-id="462e2-558">Die Anlage ist zu groß.</span><span class="sxs-lookup"><span data-stu-id="462e2-558">The attachment is larger than allowed.</span></span> |
| `FileTypeNotSupported` | <span data-ttu-id="462e2-559">Die Anlage enthält eine Erweiterung, die nicht zulässig ist.</span><span class="sxs-lookup"><span data-stu-id="462e2-559">The attachment has an extension that is not allowed.</span></span> |
| `NumberOfAttachmentsExceeded` | <span data-ttu-id="462e2-560">Die Nachricht oder der Termin enthält zu viele Anlagen.</span><span class="sxs-lookup"><span data-stu-id="462e2-560">The message or appointment has too many attachments.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="462e2-561">Anforderungen</span><span class="sxs-lookup"><span data-stu-id="462e2-561">Requirements</span></span>

|<span data-ttu-id="462e2-562">Anforderung</span><span class="sxs-lookup"><span data-stu-id="462e2-562">Requirement</span></span>| <span data-ttu-id="462e2-563">Wert</span><span class="sxs-lookup"><span data-stu-id="462e2-563">Value</span></span>|
|---|---|
|[<span data-ttu-id="462e2-564">Mindestversion des Postfachanforderungssatzes</span><span class="sxs-lookup"><span data-stu-id="462e2-564">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="462e2-565">1.1</span><span class="sxs-lookup"><span data-stu-id="462e2-565">1.1</span></span>|
|[<span data-ttu-id="462e2-566">Mindestberechtigungsstufe</span><span class="sxs-lookup"><span data-stu-id="462e2-566">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="462e2-567">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="462e2-567">ReadWriteItem</span></span>|
|[<span data-ttu-id="462e2-568">Zutreffender Outlook-Modus</span><span class="sxs-lookup"><span data-stu-id="462e2-568">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="462e2-569">Verfassen</span><span class="sxs-lookup"><span data-stu-id="462e2-569">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="462e2-570">Beispiel</span><span class="sxs-lookup"><span data-stu-id="462e2-570">Example</span></span>

```
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

####  <a name="additemattachmentasyncitemid-attachmentname-options-callback"></a><span data-ttu-id="462e2-571">addItemAttachmentAsync(itemId, attachmentName, [options], [callback])</span><span class="sxs-lookup"><span data-stu-id="462e2-571">addItemAttachmentAsync(itemId, attachmentName, [options], [callback])</span></span>

<span data-ttu-id="462e2-572">Fügt der Nachricht oder dem Termin ein Exchange-Objekt, wie z. B. eine Nachricht, als Anhang hinzu.</span><span class="sxs-lookup"><span data-stu-id="462e2-572">Adds an Exchange item, such as a message, as an attachment to the message or appointment.</span></span>

<span data-ttu-id="462e2-p134">Die `addItemAttachmentAsync`-Methode fügt das Element mit dem angegebenen Exchange-Bezeichner an das Element im Formular zum Verfassen an. Wenn Sie eine Rückrufmethode angeben, wird die Methode mit einem `asyncResult`-Parameter aufgerufen, der entweder den Anlagenbezeichner oder einen Code enthält, der etwaige Fehler angibt, die beim Anfügen des Objekts aufgetreten sind. Sie können ggf. den `options`-Parameter verwenden, um Statusinformationen an die Rückrufmethode zu übergeben.</span><span class="sxs-lookup"><span data-stu-id="462e2-p134">The `addItemAttachmentAsync` method attaches the item with the specified Exchange identifier to the item in the compose form. If you specify a callback method, the method is called with one parameter, `asyncResult`, which contains either the attachment identifier or a code that indicates any error that occurred while attaching the item. You can use the `options` parameter to pass state information to the callback method, if needed.</span></span>

<span data-ttu-id="462e2-576">Anschließend können Sie den Bezeichner mit der [`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback)-Methode in der gleichen Sitzung zum Entfernen der Anlage verwenden.</span><span class="sxs-lookup"><span data-stu-id="462e2-576">You can subsequently use the identifier with the [`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback) method to remove the attachment in the same session.</span></span>

<span data-ttu-id="462e2-577">Wenn Ihre Office-Add-Ins in Outlook Web App ausgeführt wird die `addItemAttachmentAsync` -Methode kann Anfügen von Elementen in der Elemente, die Sie bearbeiten; Allerdings Dies wird nicht unterstützt und wird nicht empfohlen.</span><span class="sxs-lookup"><span data-stu-id="462e2-577">If your Office Add-in is running in Outlook Web App, the `addItemAttachmentAsync` method can attach items to items other than the item that you are editing; however, this is not supported and is not recommended.</span></span>

##### <a name="parameters"></a><span data-ttu-id="462e2-578">Parameter:</span><span class="sxs-lookup"><span data-stu-id="462e2-578">Parameters:</span></span>

|<span data-ttu-id="462e2-579">Name</span><span class="sxs-lookup"><span data-stu-id="462e2-579">Name</span></span>| <span data-ttu-id="462e2-580">Typ</span><span class="sxs-lookup"><span data-stu-id="462e2-580">Type</span></span>| <span data-ttu-id="462e2-581">Attribute</span><span class="sxs-lookup"><span data-stu-id="462e2-581">Attributes</span></span>| <span data-ttu-id="462e2-582">Beschreibung</span><span class="sxs-lookup"><span data-stu-id="462e2-582">Description</span></span>|
|---|---|---|---|
|`itemId`| <span data-ttu-id="462e2-583">Zeichenfolge</span><span class="sxs-lookup"><span data-stu-id="462e2-583">String</span></span>||<span data-ttu-id="462e2-p135">Der Exchange-Bezeichner des Objekts, das angehängt werden soll. Die maximale Länge beträgt 100 Zeichen.</span><span class="sxs-lookup"><span data-stu-id="462e2-p135">The Exchange identifier of the item to attach. The maximum length is 100 characters.</span></span>|
|`attachmentName`| <span data-ttu-id="462e2-586">Zeichenfolge</span><span class="sxs-lookup"><span data-stu-id="462e2-586">String</span></span>||<span data-ttu-id="462e2-p136">Der Betreff des Elements, das angehängt werden soll. Die maximale Länge ist 255 Zeichen.</span><span class="sxs-lookup"><span data-stu-id="462e2-p136">The sujbect of the item to be attached. The maximum length is 255 characters.</span></span>|
|`options`| <span data-ttu-id="462e2-589">Object</span><span class="sxs-lookup"><span data-stu-id="462e2-589">Object</span></span>| <span data-ttu-id="462e2-590">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="462e2-590">&lt;optional&gt;</span></span>|<span data-ttu-id="462e2-591">Ein Objektliteral, das mindestens eine der folgenden Eigenschaften enthält.</span><span class="sxs-lookup"><span data-stu-id="462e2-591">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`| <span data-ttu-id="462e2-592">Object</span><span class="sxs-lookup"><span data-stu-id="462e2-592">Object</span></span>| <span data-ttu-id="462e2-593">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="462e2-593">&lt;optional&gt;</span></span>|<span data-ttu-id="462e2-594">Entwickler können ein Objekt bereitstellen, auf das sie in der Callbackmethode zugreifen möchten.</span><span class="sxs-lookup"><span data-stu-id="462e2-594">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`| <span data-ttu-id="462e2-595">Funktion</span><span class="sxs-lookup"><span data-stu-id="462e2-595">function</span></span>| <span data-ttu-id="462e2-596">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="462e2-596">&lt;optional&gt;</span></span>|<span data-ttu-id="462e2-597">Wenn die Methode abgeschlossen wird, wird die Funktion , die im `callback`-Parameter übergeben wird, mit einem einzelnen Parameter (`asyncResult`) aufgerufen, der ein [`AsyncResult`](/javascript/api/office/office.asyncresult)-Objekt darstellt.</span><span class="sxs-lookup"><span data-stu-id="462e2-597">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span> <br/><span data-ttu-id="462e2-598">Bei Erfolg wird der Anlagenbezeichner in der `asyncResult.value`-Eigenschaft bereitgestellt.</span><span class="sxs-lookup"><span data-stu-id="462e2-598">On success, the attachment identifier will be provided in the `asyncResult.value` property.</span></span><br/><span data-ttu-id="462e2-599">Wenn beim Hinzufügen der Anlage ein Fehler auftritt, enthält das `asyncResult`-Objekt ein `Error`-Objekt mit einer Beschreibung des Fehlers.</span><span class="sxs-lookup"><span data-stu-id="462e2-599">If adding the attachment fails, the `asyncResult` object will contain an `Error` object that provides a description of the error.</span></span>|

##### <a name="errors"></a><span data-ttu-id="462e2-600">Fehler</span><span class="sxs-lookup"><span data-stu-id="462e2-600">Errors</span></span>

| <span data-ttu-id="462e2-601">Fehlercode</span><span class="sxs-lookup"><span data-stu-id="462e2-601">Error code</span></span> | <span data-ttu-id="462e2-602">Beschreibung</span><span class="sxs-lookup"><span data-stu-id="462e2-602">Description</span></span> |
|------------|-------------|
| `NumberOfAttachmentsExceeded` | <span data-ttu-id="462e2-603">Die Nachricht oder der Termin enthält zu viele Anlagen.</span><span class="sxs-lookup"><span data-stu-id="462e2-603">The message or appointment has too many attachments.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="462e2-604">Anforderungen</span><span class="sxs-lookup"><span data-stu-id="462e2-604">Requirements</span></span>

|<span data-ttu-id="462e2-605">Anforderung</span><span class="sxs-lookup"><span data-stu-id="462e2-605">Requirement</span></span>| <span data-ttu-id="462e2-606">Wert</span><span class="sxs-lookup"><span data-stu-id="462e2-606">Value</span></span>|
|---|---|
|[<span data-ttu-id="462e2-607">Mindestversion des Postfachanforderungssatzes</span><span class="sxs-lookup"><span data-stu-id="462e2-607">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="462e2-608">1.1</span><span class="sxs-lookup"><span data-stu-id="462e2-608">1.1</span></span>|
|[<span data-ttu-id="462e2-609">Mindestberechtigungsstufe</span><span class="sxs-lookup"><span data-stu-id="462e2-609">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="462e2-610">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="462e2-610">ReadWriteItem</span></span>|
|[<span data-ttu-id="462e2-611">Zutreffender Outlook-Modus</span><span class="sxs-lookup"><span data-stu-id="462e2-611">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="462e2-612">Verfassen</span><span class="sxs-lookup"><span data-stu-id="462e2-612">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="462e2-613">Beispiel</span><span class="sxs-lookup"><span data-stu-id="462e2-613">Example</span></span>

<span data-ttu-id="462e2-614">Im folgenden Beispiel wird ein vorhandenes Outlook-Element als Anlage mit dem Namen `My Attachment` hinzugefügt.</span><span class="sxs-lookup"><span data-stu-id="462e2-614">The following example adds an existing Outlook item as an attachment with the name `My Attachment`.</span></span>

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

####  <a name="close"></a><span data-ttu-id="462e2-615">close()</span><span class="sxs-lookup"><span data-stu-id="462e2-615">close()</span></span>

<span data-ttu-id="462e2-616">Schließt das aktuelle Element, das gerade verfasst wird.</span><span class="sxs-lookup"><span data-stu-id="462e2-616">Closes the current item that is being composed.</span></span>

<span data-ttu-id="462e2-p137">Das Verhalten der `close`-Methode hängt vom aktuellen Status des verfassten Elements ab. Wenn das Element über nicht gespeicherte Änderungen verfügt, fordert der Client den Benutzer auf, die Schließenaktion zu speichern, zu verwerfen oder abzubrechen.</span><span class="sxs-lookup"><span data-stu-id="462e2-p137">The behavior of the `close` method depends on the current state of the item being composed. If the item has unsaved changes, the client prompts the user to save, discard, or cancel the close action.</span></span>

> [!NOTE]
> <span data-ttu-id="462e2-619">In Outlook im Web, wenn das Element einen Termin handelt, und es wurde bereits zuvor gespeichert wurde mit `saveAsync`, der Benutzer wird aufgefordert, speichern, löschen oder Abbrechen, auch wenn keine Änderungen aufgetreten sind, da das Element zuletzt gespeichert.</span><span class="sxs-lookup"><span data-stu-id="462e2-619">In Outlook on the web, if the item is an appointment and it has previously been saved using `saveAsync`, the user is prompted to save, discard, or cancel even if no changes have occurred since the item was last saved.</span></span>

<span data-ttu-id="462e2-620">Wenn die Nachricht im Outlook-Desktopclient eine Inlineantwort ist, hat die `close`-Methode keine Auswirkung.</span><span class="sxs-lookup"><span data-stu-id="462e2-620">In the Outlook desktop client, if the message is an inline reply, the `close` method has no effect.</span></span>

##### <a name="requirements"></a><span data-ttu-id="462e2-621">Anforderungen</span><span class="sxs-lookup"><span data-stu-id="462e2-621">Requirements</span></span>

|<span data-ttu-id="462e2-622">Anforderung</span><span class="sxs-lookup"><span data-stu-id="462e2-622">Requirement</span></span>| <span data-ttu-id="462e2-623">Wert</span><span class="sxs-lookup"><span data-stu-id="462e2-623">Value</span></span>|
|---|---|
|[<span data-ttu-id="462e2-624">Mindestversion des Postfachanforderungssatzes</span><span class="sxs-lookup"><span data-stu-id="462e2-624">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="462e2-625">1.3</span><span class="sxs-lookup"><span data-stu-id="462e2-625">1.3</span></span>|
|[<span data-ttu-id="462e2-626">Mindestberechtigungsstufe</span><span class="sxs-lookup"><span data-stu-id="462e2-626">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="462e2-627">Eingeschränkt</span><span class="sxs-lookup"><span data-stu-id="462e2-627">Restricted</span></span>|
|[<span data-ttu-id="462e2-628">Zutreffender Outlook-Modus</span><span class="sxs-lookup"><span data-stu-id="462e2-628">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="462e2-629">Verfassen</span><span class="sxs-lookup"><span data-stu-id="462e2-629">Compose</span></span>|

#### <a name="displayreplyallformformdata"></a><span data-ttu-id="462e2-630">displayReplyAllForm(formData)</span><span class="sxs-lookup"><span data-stu-id="462e2-630">displayReplyAllForm(formData)</span></span>

<span data-ttu-id="462e2-631">Zeigt ein Antwortformular an, das den Absender und alle Empfänger der ausgewählten Nachricht oder des Organisators und alle Teilnehmer des ausgewählten Termins enthält.</span><span class="sxs-lookup"><span data-stu-id="462e2-631">Displays a reply form that includes the sender and all recipients of the selected message or the organizer and all attendees of the selected appointment.</span></span>

> [!NOTE]
> <span data-ttu-id="462e2-632">Diese Methode wird in Outlook für iOS oder Outlook für Android nicht unterstützt.</span><span class="sxs-lookup"><span data-stu-id="462e2-632">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="462e2-633">In Outlook Web App wird das Antwortformular als Popupformular in der Dreispaltenansicht und als Popupformular in der Zwei- oder Einspaltenansicht angezeigt.</span><span class="sxs-lookup"><span data-stu-id="462e2-633">In Outlook Web App, the reply form is displayed as a pop-out form in the 3-column view and a pop-up form in the 2- or 1-column view.</span></span>

<span data-ttu-id="462e2-634">Wenn einer der Zeichenfolgenparameter seinen Grenzwert überschreitet, löst `displayReplyAllForm` eine Ausnahme aus.</span><span class="sxs-lookup"><span data-stu-id="462e2-634">If any of the string parameters exceed their limits, `displayReplyAllForm` throws an exception.</span></span>

<span data-ttu-id="462e2-p138">Wenn Anlagen im `formData.attachments`-Parameter angegeben werden, versuchen Outlook und Outlook Web App alle Anlagen herunterzuladen und sie an das Antwortformular anzufügen. Wenn Anlagen nicht hinzugefügt werden können, wird in der Formularbenutzeroberfläche ein Fehler ausgegeben. Wenn dies nicht möglich ist, wird keine Fehlermeldung ausgegeben.</span><span class="sxs-lookup"><span data-stu-id="462e2-p138">When attachments are specified in the `formData.attachments` parameter, Outlook and Outlook Web App attempt to download all attachments and attach them to the reply form. If any attachments fail to be added, an error is shown in the form UI. If this isn't possible, then no error message is thrown.</span></span>

##### <a name="parameters"></a><span data-ttu-id="462e2-638">Parameter:</span><span class="sxs-lookup"><span data-stu-id="462e2-638">Parameters:</span></span>

|<span data-ttu-id="462e2-639">Name</span><span class="sxs-lookup"><span data-stu-id="462e2-639">Name</span></span>| <span data-ttu-id="462e2-640">Typ</span><span class="sxs-lookup"><span data-stu-id="462e2-640">Type</span></span>| <span data-ttu-id="462e2-641">Beschreibung</span><span class="sxs-lookup"><span data-stu-id="462e2-641">Description</span></span>|
|---|---|---|
|`formData`| <span data-ttu-id="462e2-642">Zeichenfolge &#124; Objekt</span><span class="sxs-lookup"><span data-stu-id="462e2-642">String &#124; Object</span></span>| |<span data-ttu-id="462e2-p139">Eine Zeichenfolge, die Text- und HTML-Code enthält, die den Hauptteil des Antwortformulars darstellen. Die Zeichenfolge ist auf 32 KB beschränkt.</span><span class="sxs-lookup"><span data-stu-id="462e2-p139">A string that contains text and HTML and that represents the body of the reply form. The string is limited to 32 KB.</span></span><br/><span data-ttu-id="462e2-645">**ODER**</span><span class="sxs-lookup"><span data-stu-id="462e2-645">**OR**</span></span><br/><span data-ttu-id="462e2-p140">Ein Objekt, das Text- oder Anlagendaten und eine Rückruffunktion enthält. Das Objekt ist wie folgt definiert:</span><span class="sxs-lookup"><span data-stu-id="462e2-p140">An object that contains body or attachment data and a callback function. The object is defined as follows.</span></span> |
| `formData.htmlBody` | <span data-ttu-id="462e2-648">Zeichenfolge</span><span class="sxs-lookup"><span data-stu-id="462e2-648">String</span></span> | <span data-ttu-id="462e2-649">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="462e2-649">&lt;optional&gt;</span></span> | <span data-ttu-id="462e2-p141">Eine Zeichenfolge, die Text- und HTML-Code enthält, die den Hauptteil des Antwortformulars darstellen. Die Zeichenfolge ist auf 32 KB beschränkt.</span><span class="sxs-lookup"><span data-stu-id="462e2-p141">A string that contains text and HTML and that represents the body of the reply form. The string is limited to 32 KB.</span></span>
| `formData.attachments` | <span data-ttu-id="462e2-652">Array.&lt;Object&gt;</span><span class="sxs-lookup"><span data-stu-id="462e2-652">Array.&lt;Object&gt;</span></span> | <span data-ttu-id="462e2-653">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="462e2-653">&lt;optional&gt;</span></span> | <span data-ttu-id="462e2-654">Ein Array mit JSON-Objekten, die Datei- oder Elementanlagen sind.</span><span class="sxs-lookup"><span data-stu-id="462e2-654">An array of JSON objects that are either file or item attachments.</span></span> |
| `formData.attachments.type` | <span data-ttu-id="462e2-655">Zeichenfolge</span><span class="sxs-lookup"><span data-stu-id="462e2-655">String</span></span> | | <span data-ttu-id="462e2-p142">Gibt den Typ der Anlage an. Muss `file` für eine Dateianlage oder `item` für eine Elementanlage sein.</span><span class="sxs-lookup"><span data-stu-id="462e2-p142">Indicates the type of attachment. Must be `file` for a file attachment or `item` for an item attachment.</span></span> |
| `formData.attachments.name` | <span data-ttu-id="462e2-658">Zeichenfolge</span><span class="sxs-lookup"><span data-stu-id="462e2-658">String</span></span> | | <span data-ttu-id="462e2-659">Eine Zeichenfolge, die den Namen der Anlage mit bis zu 255 Zeichen enthält.</span><span class="sxs-lookup"><span data-stu-id="462e2-659">A string that contains the name of the attachment, up to 255 characters in length.</span></span>|
| `formData.attachments.url` | <span data-ttu-id="462e2-660">Zeichenfolge</span><span class="sxs-lookup"><span data-stu-id="462e2-660">String</span></span> | | <span data-ttu-id="462e2-p143">Wird nur verwendet, wenn `type` auf `file` gesetzt ist. Der URI des Speicherorts für die Datei.</span><span class="sxs-lookup"><span data-stu-id="462e2-p143">Only used if `type` is set to `file`. The URI of the location for the file.</span></span> |
| `formData.attachments.itemId` | <span data-ttu-id="462e2-663">Zeichenfolge</span><span class="sxs-lookup"><span data-stu-id="462e2-663">String</span></span> | | <span data-ttu-id="462e2-p144">Wird nur verwendet, wenn `type` auf `item` gesetzt ist. Die EWS-Element-ID der Anlage. Dies ist eine Zeichenfolge bis zu 100 Zeichen.</span><span class="sxs-lookup"><span data-stu-id="462e2-p144">Only used if `type` is set to `item`. The EWS item id of the attachment. This is a string up to 100 characters.</span></span> |
| `callback` | <span data-ttu-id="462e2-667">Funktion</span><span class="sxs-lookup"><span data-stu-id="462e2-667">function</span></span> | <span data-ttu-id="462e2-668">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="462e2-668">&lt;optional&gt;</span></span> | <span data-ttu-id="462e2-669">Bei Abschluss der Methode wird die im Parameter `callback` übergebene Funktion mit einem einzigen Parameter aufgerufen: `asyncResult`. Dieser Parameter ist ein Objekt des Typs [AsyncResult](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="462e2-669">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [AsyncResult](/javascript/api/office/office.asyncresult) object.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="462e2-670">Anforderungen</span><span class="sxs-lookup"><span data-stu-id="462e2-670">Requirements</span></span>

|<span data-ttu-id="462e2-671">Anforderung</span><span class="sxs-lookup"><span data-stu-id="462e2-671">Requirement</span></span>| <span data-ttu-id="462e2-672">Wert</span><span class="sxs-lookup"><span data-stu-id="462e2-672">Value</span></span>|
|---|---|
|[<span data-ttu-id="462e2-673">Mindestversion des Postfachanforderungssatzes</span><span class="sxs-lookup"><span data-stu-id="462e2-673">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="462e2-674">1.0</span><span class="sxs-lookup"><span data-stu-id="462e2-674">1.0</span></span>|
|[<span data-ttu-id="462e2-675">Mindestberechtigungsstufe</span><span class="sxs-lookup"><span data-stu-id="462e2-675">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="462e2-676">ReadItem</span><span class="sxs-lookup"><span data-stu-id="462e2-676">ReadItem</span></span>|
|[<span data-ttu-id="462e2-677">Zutreffender Outlook-Modus</span><span class="sxs-lookup"><span data-stu-id="462e2-677">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="462e2-678">Lesen</span><span class="sxs-lookup"><span data-stu-id="462e2-678">Read</span></span>|

##### <a name="examples"></a><span data-ttu-id="462e2-679">Beispiele</span><span class="sxs-lookup"><span data-stu-id="462e2-679">Examples</span></span>

<span data-ttu-id="462e2-680">Im folgenden Code wird eine Zeichenfolge an die `displayReplyAllForm`-Funktion übergeben.</span><span class="sxs-lookup"><span data-stu-id="462e2-680">The following code passes a string to the `displayReplyAllForm` function.</span></span>

```
Office.context.mailbox.item.displayReplyAllForm('hello there');
Office.context.mailbox.item.displayReplyAllForm('<b>hello there</b>');
```

<span data-ttu-id="462e2-681">Antworten mit einem leeren Textkörper.</span><span class="sxs-lookup"><span data-stu-id="462e2-681">Reply with an empty body.</span></span>

```
Office.context.mailbox.item.displayReplyAllForm({});
```

<span data-ttu-id="462e2-682">Antworten nur einem Textkörper.</span><span class="sxs-lookup"><span data-stu-id="462e2-682">Reply with just a body.</span></span>

```
Office.context.mailbox.item.displayReplyAllForm(
{
  'htmlBody' : 'hi'
});
```

<span data-ttu-id="462e2-683">Antworten mit einem Textkörper und einer Dateianlage.</span><span class="sxs-lookup"><span data-stu-id="462e2-683">Reply with a body and a file attachment.</span></span>

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

<span data-ttu-id="462e2-684">Antworten mit einem Textkörper und einer Elementanlage.</span><span class="sxs-lookup"><span data-stu-id="462e2-684">Reply with a body and an item attachment.</span></span>

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

<span data-ttu-id="462e2-685">Antworten Sie mit einem Textkörper, einer Dateianlage, einer Elementanlage und einem Callback.</span><span class="sxs-lookup"><span data-stu-id="462e2-685">Reply with a body, file attachment, item attachment, and a callback.</span></span>

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

#### <a name="displayreplyformformdata"></a><span data-ttu-id="462e2-686">displayReplyForm(formData)</span><span class="sxs-lookup"><span data-stu-id="462e2-686">displayReplyForm(formData)</span></span>

<span data-ttu-id="462e2-687">Zeigt ein Antwortformular an, das nur den Absender der ausgewählten Nachricht oder den Organisator des ausgewählten Termins enthält.</span><span class="sxs-lookup"><span data-stu-id="462e2-687">Displays a reply form that includes only the sender of the selected message or the organizer of the selected appointment.</span></span>

> [!NOTE]
> <span data-ttu-id="462e2-688">Diese Methode wird in Outlook für iOS oder Outlook für Android nicht unterstützt.</span><span class="sxs-lookup"><span data-stu-id="462e2-688">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="462e2-689">In Outlook Web App wird das Antwortformular als Popupformular in der Dreispaltenansicht und als Popupformular in der Zwei- oder Einspaltenansicht angezeigt.</span><span class="sxs-lookup"><span data-stu-id="462e2-689">In Outlook Web App, the reply form is displayed as a pop-out form in the 3-column view and a pop-up form in the 2- or 1-column view.</span></span>

<span data-ttu-id="462e2-690">Wenn einer der Zeichenfolgenparameter seinen Grenzwert überschreitet, löst `displayReplyForm` eine Ausnahme aus.</span><span class="sxs-lookup"><span data-stu-id="462e2-690">If any of the string parameters exceed their limits, `displayReplyForm` throws an exception.</span></span>

<span data-ttu-id="462e2-p145">Wenn Anlagen im `formData.attachments`-Parameter angegeben werden, versuchen Outlook und Outlook Web App alle Anlagen herunterzuladen und sie an das Antwortformular anzufügen. Wenn Anlagen nicht hinzugefügt werden können, wird in der Formularbenutzeroberfläche ein Fehler ausgegeben. Wenn dies nicht möglich ist, wird keine Fehlermeldung ausgegeben.</span><span class="sxs-lookup"><span data-stu-id="462e2-p145">When attachments are specified in the `formData.attachments` parameter, Outlook and Outlook Web App attempt to download all attachments and attach them to the reply form. If any attachments fail to be added, an error is shown in the form UI. If this isn't possible, then no error message is thrown.</span></span>

##### <a name="parameters"></a><span data-ttu-id="462e2-694">Parameter:</span><span class="sxs-lookup"><span data-stu-id="462e2-694">Parameters:</span></span>

|<span data-ttu-id="462e2-695">Name</span><span class="sxs-lookup"><span data-stu-id="462e2-695">Name</span></span>| <span data-ttu-id="462e2-696">Typ</span><span class="sxs-lookup"><span data-stu-id="462e2-696">Type</span></span>| <span data-ttu-id="462e2-697">Beschreibung</span><span class="sxs-lookup"><span data-stu-id="462e2-697">Description</span></span>|
|---|---|---|
|`formData`| <span data-ttu-id="462e2-698">Zeichenfolge &#124; Objekt</span><span class="sxs-lookup"><span data-stu-id="462e2-698">String &#124; Object</span></span>| | <span data-ttu-id="462e2-p146">Eine Zeichenfolge, die Text- und HTML-Code enthält, die den Hauptteil des Antwortformulars darstellen. Die Zeichenfolge ist auf 32 KB beschränkt.</span><span class="sxs-lookup"><span data-stu-id="462e2-p146">A string that contains text and HTML and that represents the body of the reply form. The string is limited to 32 KB.</span></span><br/><span data-ttu-id="462e2-701">**ODER**</span><span class="sxs-lookup"><span data-stu-id="462e2-701">**OR**</span></span><br/><span data-ttu-id="462e2-p147">Ein Objekt, das Text- oder Anlagendaten und eine Rückruffunktion enthält. Das Objekt ist wie folgt definiert:</span><span class="sxs-lookup"><span data-stu-id="462e2-p147">An object that contains body or attachment data and a callback function. The object is defined as follows.</span></span> |
| `formData.htmlBody` | <span data-ttu-id="462e2-704">Zeichenfolge</span><span class="sxs-lookup"><span data-stu-id="462e2-704">String</span></span> | <span data-ttu-id="462e2-705">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="462e2-705">&lt;optional&gt;</span></span> | <span data-ttu-id="462e2-p148">Eine Zeichenfolge, die Text- und HTML-Code enthält, die den Hauptteil des Antwortformulars darstellen. Die Zeichenfolge ist auf 32 KB beschränkt.</span><span class="sxs-lookup"><span data-stu-id="462e2-p148">A string that contains text and HTML and that represents the body of the reply form. The string is limited to 32 KB.</span></span>
| `formData.attachments` | <span data-ttu-id="462e2-708">Array.&lt;Object&gt;</span><span class="sxs-lookup"><span data-stu-id="462e2-708">Array.&lt;Object&gt;</span></span> | <span data-ttu-id="462e2-709">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="462e2-709">&lt;optional&gt;</span></span> | <span data-ttu-id="462e2-710">Ein Array mit JSON-Objekten, die Datei- oder Elementanlagen sind.</span><span class="sxs-lookup"><span data-stu-id="462e2-710">An array of JSON objects that are either file or item attachments.</span></span> |
| `formData.attachments.type` | <span data-ttu-id="462e2-711">Zeichenfolge</span><span class="sxs-lookup"><span data-stu-id="462e2-711">String</span></span> | | <span data-ttu-id="462e2-p149">Gibt den Typ der Anlage an. Muss `file` für eine Dateianlage oder `item` für eine Elementanlage sein.</span><span class="sxs-lookup"><span data-stu-id="462e2-p149">Indicates the type of attachment. Must be `file` for a file attachment or `item` for an item attachment.</span></span> |
| `formData.attachments.name` | <span data-ttu-id="462e2-714">Zeichenfolge</span><span class="sxs-lookup"><span data-stu-id="462e2-714">String</span></span> | | <span data-ttu-id="462e2-715">Eine Zeichenfolge, die den Namen der Anlage mit bis zu 255 Zeichen enthält.</span><span class="sxs-lookup"><span data-stu-id="462e2-715">A string that contains the name of the attachment, up to 255 characters in length.</span></span>|
| `formData.attachments.url` | <span data-ttu-id="462e2-716">Zeichenfolge</span><span class="sxs-lookup"><span data-stu-id="462e2-716">String</span></span> | | <span data-ttu-id="462e2-p150">Wird nur verwendet, wenn `type` auf `file` gesetzt ist. Der URI des Speicherorts für die Datei.</span><span class="sxs-lookup"><span data-stu-id="462e2-p150">Only used if `type` is set to `file`. The URI of the location for the file.</span></span> |
| `formData.attachments.itemId` | <span data-ttu-id="462e2-719">Zeichenfolge</span><span class="sxs-lookup"><span data-stu-id="462e2-719">String</span></span> | | <span data-ttu-id="462e2-p151">Wird nur verwendet, wenn `type` auf `item` gesetzt ist. Die EWS-Element-ID der Anlage. Dies ist eine Zeichenfolge bis zu 100 Zeichen.</span><span class="sxs-lookup"><span data-stu-id="462e2-p151">Only used if `type` is set to `item`. The EWS item id of the attachment. This is a string up to 100 characters.</span></span> |
| `callback` | <span data-ttu-id="462e2-723">Funktion</span><span class="sxs-lookup"><span data-stu-id="462e2-723">function</span></span> | <span data-ttu-id="462e2-724">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="462e2-724">&lt;optional&gt;</span></span> | <span data-ttu-id="462e2-725">Bei Abschluss der Methode wird die im Parameter `callback` übergebene Funktion mit einem einzigen Parameter aufgerufen: `asyncResult`. Dieser Parameter ist ein Objekt des Typs [AsyncResult](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="462e2-725">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [AsyncResult](/javascript/api/office/office.asyncresult) object.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="462e2-726">Anforderungen</span><span class="sxs-lookup"><span data-stu-id="462e2-726">Requirements</span></span>

|<span data-ttu-id="462e2-727">Anforderung</span><span class="sxs-lookup"><span data-stu-id="462e2-727">Requirement</span></span>| <span data-ttu-id="462e2-728">Wert</span><span class="sxs-lookup"><span data-stu-id="462e2-728">Value</span></span>|
|---|---|
|[<span data-ttu-id="462e2-729">Mindestversion des Postfachanforderungssatzes</span><span class="sxs-lookup"><span data-stu-id="462e2-729">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="462e2-730">1.0</span><span class="sxs-lookup"><span data-stu-id="462e2-730">1.0</span></span>|
|[<span data-ttu-id="462e2-731">Mindestberechtigungsstufe</span><span class="sxs-lookup"><span data-stu-id="462e2-731">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="462e2-732">ReadItem</span><span class="sxs-lookup"><span data-stu-id="462e2-732">ReadItem</span></span>|
|[<span data-ttu-id="462e2-733">Zutreffender Outlook-Modus</span><span class="sxs-lookup"><span data-stu-id="462e2-733">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="462e2-734">Lesen</span><span class="sxs-lookup"><span data-stu-id="462e2-734">Read</span></span>|

##### <a name="examples"></a><span data-ttu-id="462e2-735">Beispiele</span><span class="sxs-lookup"><span data-stu-id="462e2-735">Examples</span></span>

<span data-ttu-id="462e2-736">Im folgenden Code wird eine Zeichenfolge an die `displayReplyForm`-Funktion übergeben.</span><span class="sxs-lookup"><span data-stu-id="462e2-736">The following code passes a string to the `displayReplyForm` function.</span></span>

```
Office.context.mailbox.item.displayReplyForm('hello there');
Office.context.mailbox.item.displayReplyForm('<b>hello there</b>');
```

<span data-ttu-id="462e2-737">Antworten mit einem leeren Textkörper.</span><span class="sxs-lookup"><span data-stu-id="462e2-737">Reply with an empty body.</span></span>

```
Office.context.mailbox.item.displayReplyForm({});
```

<span data-ttu-id="462e2-738">Antworten nur einem Textkörper.</span><span class="sxs-lookup"><span data-stu-id="462e2-738">Reply with just a body.</span></span>

```
Office.context.mailbox.item.displayReplyForm(
{
  'htmlBody' : 'hi'
});
```

<span data-ttu-id="462e2-739">Antworten mit einem Textkörper und einer Dateianlage.</span><span class="sxs-lookup"><span data-stu-id="462e2-739">Reply with a body and a file attachment.</span></span>

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

<span data-ttu-id="462e2-740">Antworten mit einem Textkörper und einer Elementanlage.</span><span class="sxs-lookup"><span data-stu-id="462e2-740">Reply with a body and an item attachment.</span></span>

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

<span data-ttu-id="462e2-741">Antworten Sie mit einem Textkörper, einer Dateianlage, einer Elementanlage und einem Callback.</span><span class="sxs-lookup"><span data-stu-id="462e2-741">Reply with a body, file attachment, item attachment, and a callback.</span></span>

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

#### <a name="getentities--entitiesjavascriptapioutlook14officeentities"></a><span data-ttu-id="462e2-742">getEntities() → {[Entities](/javascript/api/outlook_1_4/office.entities)}</span><span class="sxs-lookup"><span data-stu-id="462e2-742">getEntities() → {[Entities](/javascript/api/outlook_1_4/office.entities)}</span></span>

<span data-ttu-id="462e2-743">Ruft das ausgewählte Element Textkörper gefundenen Entitäten ab.</span><span class="sxs-lookup"><span data-stu-id="462e2-743">Gets the entities found in the selected item's body.</span></span>

> [!NOTE]
> <span data-ttu-id="462e2-744">Diese Methode wird in Outlook für iOS oder Outlook für Android nicht unterstützt.</span><span class="sxs-lookup"><span data-stu-id="462e2-744">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

##### <a name="requirements"></a><span data-ttu-id="462e2-745">Anforderungen</span><span class="sxs-lookup"><span data-stu-id="462e2-745">Requirements</span></span>

|<span data-ttu-id="462e2-746">Anforderung</span><span class="sxs-lookup"><span data-stu-id="462e2-746">Requirement</span></span>| <span data-ttu-id="462e2-747">Wert</span><span class="sxs-lookup"><span data-stu-id="462e2-747">Value</span></span>|
|---|---|
|[<span data-ttu-id="462e2-748">Mindestversion des Postfachanforderungssatzes</span><span class="sxs-lookup"><span data-stu-id="462e2-748">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="462e2-749">1.0</span><span class="sxs-lookup"><span data-stu-id="462e2-749">1.0</span></span>|
|[<span data-ttu-id="462e2-750">Mindestberechtigungsstufe</span><span class="sxs-lookup"><span data-stu-id="462e2-750">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="462e2-751">ReadItem</span><span class="sxs-lookup"><span data-stu-id="462e2-751">ReadItem</span></span>|
|[<span data-ttu-id="462e2-752">Zutreffender Outlook-Modus</span><span class="sxs-lookup"><span data-stu-id="462e2-752">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="462e2-753">Lesen</span><span class="sxs-lookup"><span data-stu-id="462e2-753">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="462e2-754">Gibt zurück:</span><span class="sxs-lookup"><span data-stu-id="462e2-754">Returns:</span></span>

<span data-ttu-id="462e2-755">Typ: [Entitäten](/javascript/api/outlook_1_4/office.entities)</span><span class="sxs-lookup"><span data-stu-id="462e2-755">Type: [Entities](/javascript/api/outlook_1_4/office.entities)</span></span>

##### <a name="example"></a><span data-ttu-id="462e2-756">Beispiel</span><span class="sxs-lookup"><span data-stu-id="462e2-756">Example</span></span>

<span data-ttu-id="462e2-757">Das folgende Beispiel greift auf die Kontakte Entitäten im Textkörper des aktuellen Elements.</span><span class="sxs-lookup"><span data-stu-id="462e2-757">The following example accesses the contacts entities in the current item's body.</span></span>

```
var contacts = Office.context.mailbox.item.getEntities().contacts;
```

#### <a name="getentitiesbytypeentitytype--nullable-arraystringcontactjavascriptapioutlook14officecontactmeetingsuggestionjavascriptapioutlook14officemeetingsuggestionphonenumberjavascriptapioutlook14officephonenumbertasksuggestionjavascriptapioutlook14officetasksuggestion"></a><span data-ttu-id="462e2-758">getEntitiesByType(entityType) → (nullable) {Array.<(String|[Contact](/javascript/api/outlook_1_4/office.contact)|[MeetingSuggestion](/javascript/api/outlook_1_4/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook_1_4/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook_1_4/office.tasksuggestion))>}</span><span class="sxs-lookup"><span data-stu-id="462e2-758">getEntitiesByType(entityType) → (nullable) {Array.<(String|[Contact](/javascript/api/outlook_1_4/office.contact)|[MeetingSuggestion](/javascript/api/outlook_1_4/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook_1_4/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook_1_4/office.tasksuggestion))>}</span></span>

<span data-ttu-id="462e2-759">Ruft ein Array aller Entitäten des angegebenen Entitätstyps im Hauptteil des ausgewählten Elements gefunden.</span><span class="sxs-lookup"><span data-stu-id="462e2-759">Gets an array of all the entities of the specified entity type found in the selected item's body.</span></span>

> [!NOTE]
> <span data-ttu-id="462e2-760">Diese Methode wird in Outlook für iOS oder Outlook für Android nicht unterstützt.</span><span class="sxs-lookup"><span data-stu-id="462e2-760">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

##### <a name="parameters"></a><span data-ttu-id="462e2-761">Parameter:</span><span class="sxs-lookup"><span data-stu-id="462e2-761">Parameters:</span></span>

|<span data-ttu-id="462e2-762">Name</span><span class="sxs-lookup"><span data-stu-id="462e2-762">Name</span></span>| <span data-ttu-id="462e2-763">Typ</span><span class="sxs-lookup"><span data-stu-id="462e2-763">Type</span></span>| <span data-ttu-id="462e2-764">Beschreibung</span><span class="sxs-lookup"><span data-stu-id="462e2-764">Description</span></span>|
|---|---|---|
|`entityType`| [<span data-ttu-id="462e2-765">Office.MailboxEnums.EntityType</span><span class="sxs-lookup"><span data-stu-id="462e2-765">Office.MailboxEnums.EntityType</span></span>](/javascript/api/outlook_1_4/office.mailboxenums.entitytype)|<span data-ttu-id="462e2-766">Einer der Werte der EntityType-Enumeration.</span><span class="sxs-lookup"><span data-stu-id="462e2-766">One of the EntityType enumeration values.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="462e2-767">Anforderungen</span><span class="sxs-lookup"><span data-stu-id="462e2-767">Requirements</span></span>

|<span data-ttu-id="462e2-768">Anforderung</span><span class="sxs-lookup"><span data-stu-id="462e2-768">Requirement</span></span>| <span data-ttu-id="462e2-769">Wert</span><span class="sxs-lookup"><span data-stu-id="462e2-769">Value</span></span>|
|---|---|
|[<span data-ttu-id="462e2-770">Mindestversion des Postfachanforderungssatzes</span><span class="sxs-lookup"><span data-stu-id="462e2-770">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="462e2-771">1.0</span><span class="sxs-lookup"><span data-stu-id="462e2-771">1.0</span></span>|
|[<span data-ttu-id="462e2-772">Mindestberechtigungsstufe</span><span class="sxs-lookup"><span data-stu-id="462e2-772">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="462e2-773">Eingeschränkt</span><span class="sxs-lookup"><span data-stu-id="462e2-773">Restricted</span></span>|
|[<span data-ttu-id="462e2-774">Zutreffender Outlook-Modus</span><span class="sxs-lookup"><span data-stu-id="462e2-774">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="462e2-775">Lesen</span><span class="sxs-lookup"><span data-stu-id="462e2-775">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="462e2-776">Gibt zurück:</span><span class="sxs-lookup"><span data-stu-id="462e2-776">Returns:</span></span>

<span data-ttu-id="462e2-777">Wenn der in `entityType` übergebene Wert kein gültiges Element der `EntityType`-Enumeration ist, gibt die Methode null zurück.</span><span class="sxs-lookup"><span data-stu-id="462e2-777">If the value passed in `entityType` is not a valid member of the `EntityType` enumeration, the method returns null.</span></span> <span data-ttu-id="462e2-778">Wenn keine Entitäten des angegebenen Typs im Textkörper des Elements vorhanden sind, gibt die Methode ein leeres Array zurück.</span><span class="sxs-lookup"><span data-stu-id="462e2-778">If no entities of the specified type are present in the item's body, the method returns an empty array.</span></span> <span data-ttu-id="462e2-779">Andernfalls hängt der Typ der Objekte im zurückgegebenen Array vom Typ der Entität ab, die im `entityType`-Parameter angefordert wurde.</span><span class="sxs-lookup"><span data-stu-id="462e2-779">Otherwise, the type of the objects in the returned array depends on the type of entity requested in the `entityType` parameter.</span></span>

<span data-ttu-id="462e2-780">Während Sie die minimale Berechtigungsstufe für diese Methode **Restricted** ist, erfordern einige Entitätstypen **ReadItem** für den Zugriff, wie in der folgenden Tabelle angegeben wird.</span><span class="sxs-lookup"><span data-stu-id="462e2-780">While the minimum permission level to use this method is **Restricted**, some entity types require **ReadItem** to access, as specified in the following table.</span></span>

| <span data-ttu-id="462e2-781">Wert von `entityType`</span><span class="sxs-lookup"><span data-stu-id="462e2-781">Value of `entityType`</span></span> | <span data-ttu-id="462e2-782">Typ der Objekte im zurückgegebenen Array</span><span class="sxs-lookup"><span data-stu-id="462e2-782">Type of objects in returned array</span></span> | <span data-ttu-id="462e2-783">Erforderliche Berechtigungsstufe</span><span class="sxs-lookup"><span data-stu-id="462e2-783">Required Permission Level</span></span> |
| --- | --- | --- |
| `Address` | <span data-ttu-id="462e2-784">Zeichenfolge</span><span class="sxs-lookup"><span data-stu-id="462e2-784">String</span></span> | <span data-ttu-id="462e2-785">**Restricted**</span><span class="sxs-lookup"><span data-stu-id="462e2-785">**Restricted**</span></span> |
| `Contact` | <span data-ttu-id="462e2-786">Contact</span><span class="sxs-lookup"><span data-stu-id="462e2-786">Contact</span></span> | <span data-ttu-id="462e2-787">**ReadItem**</span><span class="sxs-lookup"><span data-stu-id="462e2-787">**ReadItem**</span></span> |
| `EmailAddress` | <span data-ttu-id="462e2-788">Zeichenfolge</span><span class="sxs-lookup"><span data-stu-id="462e2-788">String</span></span> | <span data-ttu-id="462e2-789">**ReadItem**</span><span class="sxs-lookup"><span data-stu-id="462e2-789">**ReadItem**</span></span> |
| `MeetingSuggestion` | <span data-ttu-id="462e2-790">MeetingSuggestion</span><span class="sxs-lookup"><span data-stu-id="462e2-790">MeetingSuggestion</span></span> | <span data-ttu-id="462e2-791">**ReadItem**</span><span class="sxs-lookup"><span data-stu-id="462e2-791">**ReadItem**</span></span> |
| `PhoneNumber` | <span data-ttu-id="462e2-792">PhoneNumber</span><span class="sxs-lookup"><span data-stu-id="462e2-792">PhoneNumber</span></span> | <span data-ttu-id="462e2-793">**Restricted**</span><span class="sxs-lookup"><span data-stu-id="462e2-793">**Restricted**</span></span> |
| `TaskSuggestion` | <span data-ttu-id="462e2-794">TaskSuggestion</span><span class="sxs-lookup"><span data-stu-id="462e2-794">TaskSuggestion</span></span> | <span data-ttu-id="462e2-795">**ReadItem**</span><span class="sxs-lookup"><span data-stu-id="462e2-795">**ReadItem**</span></span> |
| `URL` | <span data-ttu-id="462e2-796">Zeichenfolge</span><span class="sxs-lookup"><span data-stu-id="462e2-796">String</span></span> | <span data-ttu-id="462e2-797">**Restricted**</span><span class="sxs-lookup"><span data-stu-id="462e2-797">**Restricted**</span></span> |

<span data-ttu-id="462e2-798">Typ: Array.<(String|[Contact](/javascript/api/outlook_1_4/office.contact)|[MeetingSuggestion](/javascript/api/outlook_1_4/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook_1_4/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook_1_4/office.tasksuggestion))></span><span class="sxs-lookup"><span data-stu-id="462e2-798">Type: Array.<(String|[Contact](/javascript/api/outlook_1_4/office.contact)|[MeetingSuggestion](/javascript/api/outlook_1_4/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook_1_4/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook_1_4/office.tasksuggestion))></span></span>

##### <a name="example"></a><span data-ttu-id="462e2-799">Beispiel</span><span class="sxs-lookup"><span data-stu-id="462e2-799">Example</span></span>

<span data-ttu-id="462e2-800">Im folgenden Beispiel wird veranschaulicht, wie auf ein Array von Zeichenfolgen, die Postadressen im Textkörper des aktuellen Elements darstellen.</span><span class="sxs-lookup"><span data-stu-id="462e2-800">The following example shows how to access an array of strings that represent postal addresses in the current item's body.</span></span>

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

#### <a name="getfilteredentitiesbynamename--nullable-arraystringcontactjavascriptapioutlook14officecontactmeetingsuggestionjavascriptapioutlook14officemeetingsuggestionphonenumberjavascriptapioutlook14officephonenumbertasksuggestionjavascriptapioutlook14officetasksuggestion"></a><span data-ttu-id="462e2-801">getFilteredEntitiesByName(name) → (nullable) {Array.<(String|[Contact](/javascript/api/outlook_1_4/office.contact)|[MeetingSuggestion](/javascript/api/outlook_1_4/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook_1_4/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook_1_4/office.tasksuggestion))>}</span><span class="sxs-lookup"><span data-stu-id="462e2-801">getFilteredEntitiesByName(name) → (nullable) {Array.<(String|[Contact](/javascript/api/outlook_1_4/office.contact)|[MeetingSuggestion](/javascript/api/outlook_1_4/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook_1_4/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook_1_4/office.tasksuggestion))>}</span></span>

<span data-ttu-id="462e2-802">Gibt bekannte Entitäten im ausgewählten Element zurück, die den in der XML-Manifestdatei definierten benannten Filter übergeben.</span><span class="sxs-lookup"><span data-stu-id="462e2-802">Returns well-known entities in the selected item that pass the named filter defined in the manifest XML file.</span></span>

> [!NOTE]
> <span data-ttu-id="462e2-803">Diese Methode wird in Outlook für iOS oder Outlook für Android nicht unterstützt.</span><span class="sxs-lookup"><span data-stu-id="462e2-803">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="462e2-804">Die `getFilteredEntitiesByName`-Methode gibt die Entitäten zurück, die dem im [ItemHasKnownEntity](/javascript/office/manifest/rule#itemhasknownentity-rule)-Regelelement der XML-Manifestdatei definierten regulären Ausdruck mit dem angegebenen `FilterName`-Elementwert entsprechen.</span><span class="sxs-lookup"><span data-stu-id="462e2-804">The `getFilteredEntitiesByName` method returns the entities that match the regular expression defined in the [ItemHasKnownEntity](/javascript/office/manifest/rule#itemhasknownentity-rule) rule element in the manifest XML file with the specified `FilterName` element value.</span></span>

##### <a name="parameters"></a><span data-ttu-id="462e2-805">Parameter:</span><span class="sxs-lookup"><span data-stu-id="462e2-805">Parameters:</span></span>

|<span data-ttu-id="462e2-806">Name</span><span class="sxs-lookup"><span data-stu-id="462e2-806">Name</span></span>| <span data-ttu-id="462e2-807">Typ</span><span class="sxs-lookup"><span data-stu-id="462e2-807">Type</span></span>| <span data-ttu-id="462e2-808">Beschreibung</span><span class="sxs-lookup"><span data-stu-id="462e2-808">Description</span></span>|
|---|---|---|
|`name`| <span data-ttu-id="462e2-809">Zeichenfolge</span><span class="sxs-lookup"><span data-stu-id="462e2-809">String</span></span>|<span data-ttu-id="462e2-810">Der Name des `ItemHasKnownEntity`-Regelelements, das den entsprechenden Filter definiert.</span><span class="sxs-lookup"><span data-stu-id="462e2-810">The name of the `ItemHasKnownEntity` rule element that defines the filter to match.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="462e2-811">Anforderungen</span><span class="sxs-lookup"><span data-stu-id="462e2-811">Requirements</span></span>

|<span data-ttu-id="462e2-812">Anforderung</span><span class="sxs-lookup"><span data-stu-id="462e2-812">Requirement</span></span>| <span data-ttu-id="462e2-813">Wert</span><span class="sxs-lookup"><span data-stu-id="462e2-813">Value</span></span>|
|---|---|
|[<span data-ttu-id="462e2-814">Mindestversion des Postfachanforderungssatzes</span><span class="sxs-lookup"><span data-stu-id="462e2-814">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="462e2-815">1.0</span><span class="sxs-lookup"><span data-stu-id="462e2-815">1.0</span></span>|
|[<span data-ttu-id="462e2-816">Mindestberechtigungsstufe</span><span class="sxs-lookup"><span data-stu-id="462e2-816">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="462e2-817">ReadItem</span><span class="sxs-lookup"><span data-stu-id="462e2-817">ReadItem</span></span>|
|[<span data-ttu-id="462e2-818">Zutreffender Outlook-Modus</span><span class="sxs-lookup"><span data-stu-id="462e2-818">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="462e2-819">Lesen</span><span class="sxs-lookup"><span data-stu-id="462e2-819">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="462e2-820">Gibt zurück:</span><span class="sxs-lookup"><span data-stu-id="462e2-820">Returns:</span></span>

<span data-ttu-id="462e2-p153">Befindet sich kein `ItemHasKnownEntity`-Element im Manifest mit einem `FilterName`-Elementwert, der dem `name`-Parameter entspricht, gibt die Methode `null` zurück. Wenn der `name`-Parameter einem `ItemHasKnownEntity`-Element im Manifest entspricht, es aber keine entsprechenden Entitäten im aktuellen Element gibt, gibt die Methode ein leeres Array zurück.</span><span class="sxs-lookup"><span data-stu-id="462e2-p153">If there is no `ItemHasKnownEntity` element in the manifest with a `FilterName` element value that matches the `name` parameter, the method returns `null`. If the `name` parameter does match an `ItemHasKnownEntity` element in the manifest, but there are no entities in the current item that match, the method return an empty array.</span></span>

<span data-ttu-id="462e2-823">Typ: Array.<(String|[Contact](/javascript/api/outlook_1_4/office.contact)|[MeetingSuggestion](/javascript/api/outlook_1_4/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook_1_4/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook_1_4/office.tasksuggestion))></span><span class="sxs-lookup"><span data-stu-id="462e2-823">Type: Array.<(String|[Contact](/javascript/api/outlook_1_4/office.contact)|[MeetingSuggestion](/javascript/api/outlook_1_4/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook_1_4/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook_1_4/office.tasksuggestion))></span></span>

#### <a name="getregexmatches--object"></a><span data-ttu-id="462e2-824">getRegExMatches() → {Object}</span><span class="sxs-lookup"><span data-stu-id="462e2-824">getRegExMatches() → {Object}</span></span>

<span data-ttu-id="462e2-825">Gibt Zeichenfolgenwerte im ausgewählten Element zurück, die den in der XML-Manifestdatei definierten regulären Ausdrücken entsprechen.</span><span class="sxs-lookup"><span data-stu-id="462e2-825">Returns string values in the selected item that match the regular expressions defined in the manifest XML file.</span></span>

> [!NOTE]
> <span data-ttu-id="462e2-826">Diese Methode wird in Outlook für iOS oder Outlook für Android nicht unterstützt.</span><span class="sxs-lookup"><span data-stu-id="462e2-826">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="462e2-p154">Die `getRegExMatches`-Methode gibt die Zeichenfolgen zurück, die dem im `ItemHasRegularExpressionMatch`- oder `ItemHasKnownEntity`-Regelelement der XML-Manifestdatei definierten regulären Ausdruck entsprechen. Bei einer `ItemHasRegularExpressionMatch`-Regel muss eine entsprechende Zeichenfolge in der Eigenschaft des Elements vorhanden sein, das von dieser Regel angegeben wird. Der einfache `PropertyName`-Typ definiert die unterstützten Eigenschaften.</span><span class="sxs-lookup"><span data-stu-id="462e2-p154">The `getRegExMatches` method returns the strings that match the regular expression defined in each `ItemHasRegularExpressionMatch` or `ItemHasKnownEntity` rule element in the manifest XML file. For an `ItemHasRegularExpressionMatch` rule, a matching string has to occur in the property of the item that is specified by that rule. The `PropertyName` simple type defines the supported properties.</span></span>

<span data-ttu-id="462e2-830">Nehmen Sie z. B. an, dass ein Add-In-Manifest das folgende `Rule`-Element aufweist:</span><span class="sxs-lookup"><span data-stu-id="462e2-830">For example, consider an add-in manifest has the following `Rule` element:</span></span>

```
<Rule xsi:type="RuleCollection" Mode="And">
  <Rule xsi:type="ItemIs" FormType="Read" ItemType="Message" />
  <Rule xsi:type="RuleCollection" Mode="Or">
    <Rule xsi:type="ItemHasRegularExpressionMatch" RegExName="fruits" RegExValue="apple|banana|coconut" PropertyName="BodyAsPlaintext" IgnoreCase="true" />
    <Rule xsi:type="ItemHasRegularExpressionMatch" RegExName="veggies" RegExValue="tomato|onion|spinach|broccoli" PropertyName="BodyAsPlaintext" IgnoreCase="true" />
  </Rule>
</Rule>
```

<span data-ttu-id="462e2-831">Das von `getRegExMatches` zurückgegebene Objekt hätte zwei Eigenschaften: `fruits` und `veggies`.</span><span class="sxs-lookup"><span data-stu-id="462e2-831">The object returned from `getRegExMatches` would have two properties: `fruits` and `veggies`.</span></span>

```
{
  'fruits': ['apple','banana','Banana','coconut'],
  'veggies': ['tomato','onion','spinach','broccoli']
}
```

<span data-ttu-id="462e2-p155">Wenn Sie eine `ItemHasRegularExpressionMatch`-Regel für die Textkörpereigenschaft eines Elements festlegen, sollte der reguläre Ausdruck den Textkörper weiter filtern und nicht versuchen, den gesamten Textkörper des Elements zurückzugeben. Wenn der gesamte Textkörper eines Elements mit einem regulären Ausdruck wie `.*` abgerufen wird, werden nicht immer die gewünschten Ergebnisse erzielt. Verwenden Sie stattdessen die [`Body.getAsync`](/javascript/api/outlook_1_4/office.body#getasync-coerciontype--options--callback-)-Methode, um den gesamten Textkörper abzurufen.</span><span class="sxs-lookup"><span data-stu-id="462e2-p155">If you specify an `ItemHasRegularExpressionMatch` rule on the body property of an item, the regular expression should further filter the body and should not attempt to return the entire body of the item. Using a regular expression such as `.*` to obtain the entire body of an item does not always return the expected results. Instead, use the [`Body.getAsync`](/javascript/api/outlook_1_4/office.body#getasync-coerciontype--options--callback-) method to retrieve the entire body.</span></span>

##### <a name="requirements"></a><span data-ttu-id="462e2-835">Anforderungen</span><span class="sxs-lookup"><span data-stu-id="462e2-835">Requirements</span></span>

|<span data-ttu-id="462e2-836">Anforderung</span><span class="sxs-lookup"><span data-stu-id="462e2-836">Requirement</span></span>| <span data-ttu-id="462e2-837">Wert</span><span class="sxs-lookup"><span data-stu-id="462e2-837">Value</span></span>|
|---|---|
|[<span data-ttu-id="462e2-838">Mindestversion des Postfachanforderungssatzes</span><span class="sxs-lookup"><span data-stu-id="462e2-838">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="462e2-839">1.0</span><span class="sxs-lookup"><span data-stu-id="462e2-839">1.0</span></span>|
|[<span data-ttu-id="462e2-840">Mindestberechtigungsstufe</span><span class="sxs-lookup"><span data-stu-id="462e2-840">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="462e2-841">ReadItem</span><span class="sxs-lookup"><span data-stu-id="462e2-841">ReadItem</span></span>|
|[<span data-ttu-id="462e2-842">Zutreffender Outlook-Modus</span><span class="sxs-lookup"><span data-stu-id="462e2-842">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="462e2-843">Lesen</span><span class="sxs-lookup"><span data-stu-id="462e2-843">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="462e2-844">Gibt zurück:</span><span class="sxs-lookup"><span data-stu-id="462e2-844">Returns:</span></span>

<span data-ttu-id="462e2-p156">Ein Objekt mit Arrays aus Zeichenfolgen, die den in der XML-Manifestdatei definierten regulären Ausdrücken entsprechen. Der Name der einzelnen Arrays ist gleich dem entsprechenden Wert des `RegExName`-Attributs der entsprechenden `ItemHasRegularExpressionMatch`-Regel oder des `FilterName`-Attributs der entsprechenden `ItemHasKnownEntity`-Regel.</span><span class="sxs-lookup"><span data-stu-id="462e2-p156">An object that contains arrays of strings that match the regular expressions defined in the manifest XML file. The name of each array is equal to the corresponding value of the `RegExName` attribute of the matching `ItemHasRegularExpressionMatch` rule or the `FilterName` attribute of the matching `ItemHasKnownEntity` rule.</span></span>

<dl class="param-type"><span data-ttu-id="462e2-847">

<dt>
Typ</dt>


</span><span class="sxs-lookup"><span data-stu-id="462e2-847">

<dt>Type</dt>

</span></span><dd><span data-ttu-id="462e2-848">Object</span><span class="sxs-lookup"><span data-stu-id="462e2-848">Object</span></span></dd>

</dl>

##### <a name="example"></a><span data-ttu-id="462e2-849">Beispiel</span><span class="sxs-lookup"><span data-stu-id="462e2-849">Example</span></span>

<span data-ttu-id="462e2-850">Das folgende Beispiel zeigt, wie Sie auf das Array von Übereinstimmungen für die regulären <rule>-Ausdruckselemente `fruits` und `veggies` zugreifen, die im Manifest angegeben sind.</rule></span><span class="sxs-lookup"><span data-stu-id="462e2-850">The following example shows how to access the array of matches for the regular expression <rule>elements `fruits` and `veggies`, which are specified in the manifest.</rule></span></span>

```
var allMatches = Office.context.mailbox.item.getRegExMatches();
var fruits = allMatches.fruits;
var veges = allMatches.veggies;
```

#### <a name="getregexmatchesbynamename--nullable-array-string-"></a><span data-ttu-id="462e2-851">getRegExMatchesByName(name) → (Nullwerte zulassen) {Array. < Zeichenfolge >}</span><span class="sxs-lookup"><span data-stu-id="462e2-851">getRegExMatchesByName(name) → (nullable) {Array.< String >}</span></span>

<span data-ttu-id="462e2-852">Gibt Zeichenfolgenwerte im ausgewählten Element zurück, die dem in der XML-Manifestdatei definierten benannten regulären Ausdruck entsprechen.</span><span class="sxs-lookup"><span data-stu-id="462e2-852">Returns string values in the selected item that match the named regular expression defined in the manifest XML file.</span></span>

> [!NOTE]
> <span data-ttu-id="462e2-853">Diese Methode wird in Outlook für iOS oder Outlook für Android nicht unterstützt.</span><span class="sxs-lookup"><span data-stu-id="462e2-853">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="462e2-854">Die `getRegExMatchesByName`-Methode gibt die Zeichenfolgen zurück, die dem im `ItemHasRegularExpressionMatch`-Regelelement der XML-Manifestdatei definierten regulären Ausdruck mit dem angegebenen `RegExName`-Elementwert entsprechen.</span><span class="sxs-lookup"><span data-stu-id="462e2-854">The `getRegExMatchesByName` method returns the strings that match the regular expression defined in the `ItemHasRegularExpressionMatch` rule element in the manifest XML file with the specified `RegExName` element value.</span></span>

<span data-ttu-id="462e2-p157">Wenn Sie eine `ItemHasRegularExpressionMatch`-Regel für die Textkörpereigenschaft eines Elements festlegen, sollte der reguläre Ausdruck den Textkörper weiter filtern und nicht versuchen, den gesamten Textkörper des Elements zurückzugeben. Wenn der gesamte Textkörper eines Elements mit einem regulären Ausdruck wie `.*` abgerufen wird, werden nicht immer die gewünschten Ergebnisse erzielt.</span><span class="sxs-lookup"><span data-stu-id="462e2-p157">If you specify an `ItemHasRegularExpressionMatch` rule on the body property of an item, the regular expression should further filter the body and should not attempt to return the entire body of the item. Using a regular expression such as `.*` to obtain the entire body of an item does not always return the expected results.</span></span>

##### <a name="parameters"></a><span data-ttu-id="462e2-857">Parameter:</span><span class="sxs-lookup"><span data-stu-id="462e2-857">Parameters:</span></span>

|<span data-ttu-id="462e2-858">Name</span><span class="sxs-lookup"><span data-stu-id="462e2-858">Name</span></span>| <span data-ttu-id="462e2-859">Typ</span><span class="sxs-lookup"><span data-stu-id="462e2-859">Type</span></span>| <span data-ttu-id="462e2-860">Beschreibung</span><span class="sxs-lookup"><span data-stu-id="462e2-860">Description</span></span>|
|---|---|---|
|`name`| <span data-ttu-id="462e2-861">Zeichenfolge</span><span class="sxs-lookup"><span data-stu-id="462e2-861">String</span></span>|<span data-ttu-id="462e2-862">Der Name des `ItemHasRegularExpressionMatch`-Regelelements, das den entsprechenden Filter definiert.</span><span class="sxs-lookup"><span data-stu-id="462e2-862">The name of the `ItemHasRegularExpressionMatch` rule element that defines the filter to match.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="462e2-863">Anforderungen</span><span class="sxs-lookup"><span data-stu-id="462e2-863">Requirements</span></span>

|<span data-ttu-id="462e2-864">Anforderung</span><span class="sxs-lookup"><span data-stu-id="462e2-864">Requirement</span></span>| <span data-ttu-id="462e2-865">Wert</span><span class="sxs-lookup"><span data-stu-id="462e2-865">Value</span></span>|
|---|---|
|[<span data-ttu-id="462e2-866">Mindestversion des Postfachanforderungssatzes</span><span class="sxs-lookup"><span data-stu-id="462e2-866">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="462e2-867">1.0</span><span class="sxs-lookup"><span data-stu-id="462e2-867">1.0</span></span>|
|[<span data-ttu-id="462e2-868">Mindestberechtigungsstufe</span><span class="sxs-lookup"><span data-stu-id="462e2-868">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="462e2-869">ReadItem</span><span class="sxs-lookup"><span data-stu-id="462e2-869">ReadItem</span></span>|
|[<span data-ttu-id="462e2-870">Zutreffender Outlook-Modus</span><span class="sxs-lookup"><span data-stu-id="462e2-870">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="462e2-871">Lesen</span><span class="sxs-lookup"><span data-stu-id="462e2-871">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="462e2-872">Gibt zurück:</span><span class="sxs-lookup"><span data-stu-id="462e2-872">Returns:</span></span>

<span data-ttu-id="462e2-873">Ein Array mit den Zeichenfolgen, die dem in der XML-Manifestdatei definierten regulären Ausdruck entsprechen.</span><span class="sxs-lookup"><span data-stu-id="462e2-873">An array that contains the strings that match the regular expression defined in the manifest XML file.</span></span>

<dl class="param-type"><span data-ttu-id="462e2-874">

<dt>Typ</dt>

</span><span class="sxs-lookup"><span data-stu-id="462e2-874">

<dt>Type</dt>

</span></span><dd><span data-ttu-id="462e2-875">Array. < Zeichenfolge ></span><span class="sxs-lookup"><span data-stu-id="462e2-875">Array.< String ></span></span></dd>

</dl>

##### <a name="example"></a><span data-ttu-id="462e2-876">Beispiel</span><span class="sxs-lookup"><span data-stu-id="462e2-876">Example</span></span>

```
var fruits = Office.context.mailbox.item.getRegExMatchesByName("fruits");
var veggies = Office.context.mailbox.item.getRegExMatchesByName("veggies");
```

####  <a name="getselecteddataasynccoerciontype-options-callback--string"></a><span data-ttu-id="462e2-877">getSelectedDataAsync(coercionType, [options], callback) → {String}</span><span class="sxs-lookup"><span data-stu-id="462e2-877">getSelectedDataAsync(coercionType, [options], callback) → {String}</span></span>

<span data-ttu-id="462e2-878">Gibt asynchron ausgewählte Daten aus dem Betreff oder Textkörper einer Nachricht zurück.</span><span class="sxs-lookup"><span data-stu-id="462e2-878">Asynchronously returns selected data from the subject or body of a message.</span></span>

<span data-ttu-id="462e2-p158">Wenn keine Auswahl vorhanden ist, aber der Cursor sich im Nachrichtentext oder Betreff befindet, gibt die Methode für die ausgewählten Daten NULL zurück. Wenn ein anderes Feld als der Textkörper oder Betreff ausgewählt ist, gibt die Methode den `InvalidSelection`-Fehler zurück.</span><span class="sxs-lookup"><span data-stu-id="462e2-p158">If there is no selection but the cursor is in the body or subject, the method returns null for the selected data. If a field other than the body or subject is selected, the method returns the `InvalidSelection` error.</span></span>

##### <a name="parameters"></a><span data-ttu-id="462e2-881">Parameter:</span><span class="sxs-lookup"><span data-stu-id="462e2-881">Parameters:</span></span>

|<span data-ttu-id="462e2-882">Name</span><span class="sxs-lookup"><span data-stu-id="462e2-882">Name</span></span>| <span data-ttu-id="462e2-883">Typ</span><span class="sxs-lookup"><span data-stu-id="462e2-883">Type</span></span>| <span data-ttu-id="462e2-884">Attribute</span><span class="sxs-lookup"><span data-stu-id="462e2-884">Attributes</span></span>| <span data-ttu-id="462e2-885">Beschreibung</span><span class="sxs-lookup"><span data-stu-id="462e2-885">Description</span></span>|
|---|---|---|---|
|`coercionType`| [<span data-ttu-id="462e2-886">Office.CoercionType</span><span class="sxs-lookup"><span data-stu-id="462e2-886">Office.CoercionType</span></span>](office.md#coerciontype-string)||<span data-ttu-id="462e2-p159">Fordert ein Format für die Daten an. Wenn es sich um Texthandelt, gibt die Methode den unformatierten Text als Zeichenfolge zurück und entfernt eventuell vorhandene HTML-Tags. Wenn es sich um HTML handelt, gibt die Methode den ausgewählten Text zurück, entweder als unformatierten Text oder als HTML.</span><span class="sxs-lookup"><span data-stu-id="462e2-p159">Requests a format for the data. If Text, the method returns the plain text as a string , removing any HTML tags present. If HTML, the method returns the selected text, whether it is plaintext or HTML.</span></span>|
|`options`| <span data-ttu-id="462e2-890">Object</span><span class="sxs-lookup"><span data-stu-id="462e2-890">Object</span></span>| <span data-ttu-id="462e2-891">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="462e2-891">&lt;optional&gt;</span></span>|<span data-ttu-id="462e2-892">Ein Objektliteral, das mindestens eine der folgenden Eigenschaften enthält.</span><span class="sxs-lookup"><span data-stu-id="462e2-892">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`| <span data-ttu-id="462e2-893">Object</span><span class="sxs-lookup"><span data-stu-id="462e2-893">Object</span></span>| <span data-ttu-id="462e2-894">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="462e2-894">&lt;optional&gt;</span></span>|<span data-ttu-id="462e2-895">Entwickler können ein Objekt bereitstellen, auf das sie in der Callbackmethode zugreifen möchten.</span><span class="sxs-lookup"><span data-stu-id="462e2-895">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`| <span data-ttu-id="462e2-896">Funktion</span><span class="sxs-lookup"><span data-stu-id="462e2-896">function</span></span>||<span data-ttu-id="462e2-897">Wenn die Methode abgeschlossen wird, wird die Funktion , die im `callback`-Parameter übergeben wird, mit einem einzelnen Parameter (`asyncResult`) aufgerufen, der ein [`AsyncResult`](/javascript/api/office/office.asyncresult)-Objekt darstellt.</span><span class="sxs-lookup"><span data-stu-id="462e2-897">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span><br/><br/><span data-ttu-id="462e2-898">Rufen Sie für den Zugriff auf die ausgewählten Daten aus der Rückrufmethode `asyncResult.value.data` auf.</span><span class="sxs-lookup"><span data-stu-id="462e2-898">To access the selected data from the callback method, call `asyncResult.value.data`.</span></span> <span data-ttu-id="462e2-899">Rufen Sie die Source-Eigenschaft für den Zugriff, die die Auswahl stammen, `asyncResult.value.sourceProperty`, die entweder sein `body` oder `subject`.</span><span class="sxs-lookup"><span data-stu-id="462e2-899">To access the source property that the selection comes from, call `asyncResult.value.sourceProperty`, which will be either `body` or `subject`.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="462e2-900">Anforderungen</span><span class="sxs-lookup"><span data-stu-id="462e2-900">Requirements</span></span>

|<span data-ttu-id="462e2-901">Anforderung</span><span class="sxs-lookup"><span data-stu-id="462e2-901">Requirement</span></span>| <span data-ttu-id="462e2-902">Wert</span><span class="sxs-lookup"><span data-stu-id="462e2-902">Value</span></span>|
|---|---|
|[<span data-ttu-id="462e2-903">Mindestversion des Postfachanforderungssatzes</span><span class="sxs-lookup"><span data-stu-id="462e2-903">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="462e2-904">1.2</span><span class="sxs-lookup"><span data-stu-id="462e2-904">1.2</span></span>|
|[<span data-ttu-id="462e2-905">Mindestberechtigungsstufe</span><span class="sxs-lookup"><span data-stu-id="462e2-905">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="462e2-906">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="462e2-906">ReadWriteItem</span></span>|
|[<span data-ttu-id="462e2-907">Zutreffender Outlook-Modus</span><span class="sxs-lookup"><span data-stu-id="462e2-907">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="462e2-908">Verfassen</span><span class="sxs-lookup"><span data-stu-id="462e2-908">Compose</span></span>|

##### <a name="returns"></a><span data-ttu-id="462e2-909">Gibt zurück:</span><span class="sxs-lookup"><span data-stu-id="462e2-909">Returns:</span></span>

<span data-ttu-id="462e2-910">Die ausgewählten Daten als Zeichenfolge mit dem durch `coercionType` bestimmten Format.</span><span class="sxs-lookup"><span data-stu-id="462e2-910">The selected data as a string with format determined by `coercionType`.</span></span>

<dl class="param-type"><span data-ttu-id="462e2-911">

<dt>
Typ</dt>


</span><span class="sxs-lookup"><span data-stu-id="462e2-911">

<dt>Type</dt>

</span></span><dd><span data-ttu-id="462e2-912">Zeichenfolge</span><span class="sxs-lookup"><span data-stu-id="462e2-912">String</span></span></dd>

</dl>

##### <a name="example"></a><span data-ttu-id="462e2-913">Beispiel</span><span class="sxs-lookup"><span data-stu-id="462e2-913">Example</span></span>

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

####  <a name="loadcustompropertiesasynccallback-usercontext"></a><span data-ttu-id="462e2-914">loadCustomPropertiesAsync(callback, [userContext])</span><span class="sxs-lookup"><span data-stu-id="462e2-914">loadCustomPropertiesAsync(callback, [userContext])</span></span>

<span data-ttu-id="462e2-915">Lädt asynchron benutzerdefinierte Eigenschaften für dieses Add-In für das ausgewählte Element.</span><span class="sxs-lookup"><span data-stu-id="462e2-915">Asynchronously loads custom properties for this add-in on the selected item.</span></span>

<span data-ttu-id="462e2-p161">Benutzerdefinierte Eigenschaften werden als Schlüssel-/Wert-Paare pro App und pro Element gespeichert. Diese Methode gibt ein `CustomProperties`-Objekt im Rückruf zurück, das Methoden für den Zugriff auf die benutzerdefinierten Eigenschaften für das aktuelle Element und das aktuelle Add-In bietet. Benutzerdefinierte Eigenschaften sind für das Element nicht verschlüsselt und sollten darum nicht als sicherer Speicher verwendet werden.</span><span class="sxs-lookup"><span data-stu-id="462e2-p161">Custom properties are stored as key/value pairs on a per-app, per-item basis. This method returns a `CustomProperties` object in the callback, which provides methods to access the custom properties specific to the current item and the current add-in. Custom properties are not encrypted on the item, so this should not be used as secure storage.</span></span>

##### <a name="parameters"></a><span data-ttu-id="462e2-919">Parameter:</span><span class="sxs-lookup"><span data-stu-id="462e2-919">Parameters:</span></span>

|<span data-ttu-id="462e2-920">Name</span><span class="sxs-lookup"><span data-stu-id="462e2-920">Name</span></span>| <span data-ttu-id="462e2-921">Typ</span><span class="sxs-lookup"><span data-stu-id="462e2-921">Type</span></span>| <span data-ttu-id="462e2-922">Attribute</span><span class="sxs-lookup"><span data-stu-id="462e2-922">Attributes</span></span>| <span data-ttu-id="462e2-923">Beschreibung</span><span class="sxs-lookup"><span data-stu-id="462e2-923">Description</span></span>|
|---|---|---|---|
|`callback`| <span data-ttu-id="462e2-924">Funktion</span><span class="sxs-lookup"><span data-stu-id="462e2-924">function</span></span>||<span data-ttu-id="462e2-925">Wenn die Methode abgeschlossen wird, wird die Funktion , die im `callback`-Parameter übergeben wird, mit einem einzelnen Parameter (`asyncResult`) aufgerufen, der ein [`AsyncResult`](/javascript/api/office/office.asyncresult)-Objekt darstellt.</span><span class="sxs-lookup"><span data-stu-id="462e2-925">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span><br/><br/><span data-ttu-id="462e2-926">Die benutzerdefinierten Eigenschaften werden als [`CustomProperties`](/javascript/api/outlook_1_4/office.customproperties)-Objekt in der `asyncResult.value`-Eigenschaft bereitgestellt.</span><span class="sxs-lookup"><span data-stu-id="462e2-926">The custom properties are provided as a [`CustomProperties`](/javascript/api/outlook_1_4/office.customproperties) object in the `asyncResult.value` property.</span></span> <span data-ttu-id="462e2-927">Dieses Objekt kann zum Abrufen, festlegen und Entfernen benutzerdefinierter Eigenschaften aus dem Element und speichern Sie Änderungen an den benutzerdefinierten Eigenschaftensatz zurück an den Server verwendet werden.</span><span class="sxs-lookup"><span data-stu-id="462e2-927">This object can be used to get, set, and remove custom properties from the item and save changes to the custom property set back to the server.</span></span>|
|`userContext`| <span data-ttu-id="462e2-928">Objekt</span><span class="sxs-lookup"><span data-stu-id="462e2-928">Object</span></span>| <span data-ttu-id="462e2-929">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="462e2-929">&lt;optional&gt;</span></span>|<span data-ttu-id="462e2-930">Entwickler können ein beliebiges Objekt bereitstellen, den sie in der Rückruffunktion zugreifen möchten.</span><span class="sxs-lookup"><span data-stu-id="462e2-930">Developers can provide any object they wish to access in the callback function.</span></span> <span data-ttu-id="462e2-931">Dieses Objekt zugegriffen werden kann, indem die `asyncResult.asyncContext` -Eigenschaft in der Callback-Funktion.</span><span class="sxs-lookup"><span data-stu-id="462e2-931">This object can be accessed by the `asyncResult.asyncContext` property in the callback function.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="462e2-932">Anforderungen</span><span class="sxs-lookup"><span data-stu-id="462e2-932">Requirements</span></span>

|<span data-ttu-id="462e2-933">Anforderung</span><span class="sxs-lookup"><span data-stu-id="462e2-933">Requirement</span></span>| <span data-ttu-id="462e2-934">Wert</span><span class="sxs-lookup"><span data-stu-id="462e2-934">Value</span></span>|
|---|---|
|[<span data-ttu-id="462e2-935">Mindestversion des Postfachanforderungssatzes</span><span class="sxs-lookup"><span data-stu-id="462e2-935">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="462e2-936">1.0</span><span class="sxs-lookup"><span data-stu-id="462e2-936">1.0</span></span>|
|[<span data-ttu-id="462e2-937">Mindestberechtigungsstufe</span><span class="sxs-lookup"><span data-stu-id="462e2-937">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="462e2-938">ReadItem</span><span class="sxs-lookup"><span data-stu-id="462e2-938">ReadItem</span></span>|
|[<span data-ttu-id="462e2-939">Zutreffender Outlook-Modus</span><span class="sxs-lookup"><span data-stu-id="462e2-939">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="462e2-940">Verfassen oder Lesen</span><span class="sxs-lookup"><span data-stu-id="462e2-940">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="462e2-941">Beispiel</span><span class="sxs-lookup"><span data-stu-id="462e2-941">Example</span></span>

<span data-ttu-id="462e2-p164">Das folgende Codebeispiel veranschaulicht die Verwendung der `loadCustomPropertiesAsync`-Methode zum asynchronen Laden der benutzerdefinierten Eigenschaften, die für das aktuelle Element spezifisch sind. Des Weiteren veranschaulicht das Beispiel die Verwendung der `CustomProperties.saveAsync`-Methode zum Speichern der Eigenschaften auf dem Server. Nach dem Laden der benutzerdefinierten Eigenschaften wird in dem Codebeispiel die `CustomProperties.get`-Methode dazu verwendet, die benutzerdefinierte `myProp`-Eigenschaft zu lesen. Die `CustomProperties.set`-Methode wird dazu verwendet, die benutzerdefinierte `otherProp`-Eigenschaft zu schreiben. Schließlich wird die `saveAsync`-Methode aufgerufen, um die benutzerdefinierten Eigenschaften zu speichern.</span><span class="sxs-lookup"><span data-stu-id="462e2-p164">The following code example shows how to use the `loadCustomPropertiesAsync` method to asynchronously load custom properties that are specific to the current item. The example also shows how to use the `CustomProperties.saveAsync` method to save these properties back to the server. After loading the custom properties, the code sample uses the `CustomProperties.get` method to read the custom property `myProp`, the `CustomProperties.set` method to write the custom property `otherProp`, and then finally calls the `saveAsync` method to save the custom properties.</span></span>

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

####  <a name="removeattachmentasyncattachmentid-options-callback"></a><span data-ttu-id="462e2-945">removeAttachmentAsync(attachmentId, [options], [callback])</span><span class="sxs-lookup"><span data-stu-id="462e2-945">removeAttachmentAsync(attachmentId, [options], [callback])</span></span>

<span data-ttu-id="462e2-946">Entfernt eine Anlage aus einer Nachricht oder einem Termin.</span><span class="sxs-lookup"><span data-stu-id="462e2-946">Removes an attachment from a message or appointment.</span></span>

<span data-ttu-id="462e2-p165">Die `removeAttachmentAsync`-Methode entfernt die Anlage mit dem angegebenen Bezeichner aus dem Element. Als bewährte Vorgehensweise sollten Sie den Anlagenbezeichner nur dann zum Entfernen einer Anlage verwenden, wenn die gleiche Mail-App die Anlage in der gleichen Sitzung hinzugefügt hat. In Outlook Web App und OWA für Geräte ist der Anlagenbezeichner nur innerhalb der gleichen Sitzung gültig. Eine Sitzung ist abgeschlossen, wenn der Benutzer die App schließt, oder wenn der Benutzer in einem eingebetteten Formular mit dem Verfassen beginnt und den Vorgang dann in einem separaten Fenster fortsetzt.</span><span class="sxs-lookup"><span data-stu-id="462e2-p165">The `removeAttachmentAsync` method removes the attachment with the specified identifier from the item. As a best practice, you should use the attachment identifier to remove an attachment only if the same mail app has added that attachment in the same session. In Outlook Web App and OWA for Devices, the attachment identifier is valid only within the same session. A session is over when the user closes the app, or if the user starts composing in an inline form and subsequently pops out the inline form to continue in a separate window.</span></span>

##### <a name="parameters"></a><span data-ttu-id="462e2-951">Parameter:</span><span class="sxs-lookup"><span data-stu-id="462e2-951">Parameters:</span></span>

|<span data-ttu-id="462e2-952">Name</span><span class="sxs-lookup"><span data-stu-id="462e2-952">Name</span></span>| <span data-ttu-id="462e2-953">Typ</span><span class="sxs-lookup"><span data-stu-id="462e2-953">Type</span></span>| <span data-ttu-id="462e2-954">Attribute</span><span class="sxs-lookup"><span data-stu-id="462e2-954">Attributes</span></span>| <span data-ttu-id="462e2-955">Beschreibung</span><span class="sxs-lookup"><span data-stu-id="462e2-955">Description</span></span>|
|---|---|---|---|
|`attachmentId`| <span data-ttu-id="462e2-956">Zeichenfolge</span><span class="sxs-lookup"><span data-stu-id="462e2-956">String</span></span>||<span data-ttu-id="462e2-p166">Der Bezeichner des Anhangs, der entfernt werden soll. Die maximale Länge der Zeichenfolge ist 100 Zeichen.</span><span class="sxs-lookup"><span data-stu-id="462e2-p166">The identifier of the attachment to remove. The maximum length of the string is 100 characters.</span></span>|
|`options`| <span data-ttu-id="462e2-959">Object</span><span class="sxs-lookup"><span data-stu-id="462e2-959">Object</span></span>| <span data-ttu-id="462e2-960">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="462e2-960">&lt;optional&gt;</span></span>|<span data-ttu-id="462e2-961">Ein Objektliteral, das mindestens eine der folgenden Eigenschaften enthält.</span><span class="sxs-lookup"><span data-stu-id="462e2-961">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`| <span data-ttu-id="462e2-962">Object</span><span class="sxs-lookup"><span data-stu-id="462e2-962">Object</span></span>| <span data-ttu-id="462e2-963">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="462e2-963">&lt;optional&gt;</span></span>|<span data-ttu-id="462e2-964">Entwickler können ein Objekt bereitstellen, auf das sie in der Callbackmethode zugreifen möchten.</span><span class="sxs-lookup"><span data-stu-id="462e2-964">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`| <span data-ttu-id="462e2-965">Funktion</span><span class="sxs-lookup"><span data-stu-id="462e2-965">function</span></span>| <span data-ttu-id="462e2-966">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="462e2-966">&lt;optional&gt;</span></span>|<span data-ttu-id="462e2-967">Wenn die Methode abgeschlossen wird, wird die Funktion , die im `callback`-Parameter übergeben wird, mit einem einzelnen Parameter (`asyncResult`) aufgerufen, der ein [`AsyncResult`](/javascript/api/office/office.asyncresult)-Objekt darstellt.</span><span class="sxs-lookup"><span data-stu-id="462e2-967">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span> <br/><span data-ttu-id="462e2-968">Wenn beim Entfernen der Anlage ein Fehler auftritt, enthält die Eigenschaft `asyncResult.error` einen Fehlercode mit dem Grund für den Fehler.</span><span class="sxs-lookup"><span data-stu-id="462e2-968">If removing the attachment fails, the `asyncResult.error` property will contain an error code with the reason for the failure.</span></span>|

##### <a name="errors"></a><span data-ttu-id="462e2-969">Fehler</span><span class="sxs-lookup"><span data-stu-id="462e2-969">Errors</span></span>

| <span data-ttu-id="462e2-970">Fehlercode</span><span class="sxs-lookup"><span data-stu-id="462e2-970">Error code</span></span> | <span data-ttu-id="462e2-971">Beschreibung</span><span class="sxs-lookup"><span data-stu-id="462e2-971">Description</span></span> |
|------------|-------------|
| `InvalidAttachmentId` | <span data-ttu-id="462e2-972">Der Bezeichner für die Anlage ist nicht vorhanden.</span><span class="sxs-lookup"><span data-stu-id="462e2-972">The attachment identifier does not exist.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="462e2-973">Anforderungen</span><span class="sxs-lookup"><span data-stu-id="462e2-973">Requirements</span></span>

|<span data-ttu-id="462e2-974">Anforderung</span><span class="sxs-lookup"><span data-stu-id="462e2-974">Requirement</span></span>| <span data-ttu-id="462e2-975">Wert</span><span class="sxs-lookup"><span data-stu-id="462e2-975">Value</span></span>|
|---|---|
|[<span data-ttu-id="462e2-976">Mindestversion des Postfachanforderungssatzes</span><span class="sxs-lookup"><span data-stu-id="462e2-976">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="462e2-977">1.1</span><span class="sxs-lookup"><span data-stu-id="462e2-977">1.1</span></span>|
|[<span data-ttu-id="462e2-978">Mindestberechtigungsstufe</span><span class="sxs-lookup"><span data-stu-id="462e2-978">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="462e2-979">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="462e2-979">ReadWriteItem</span></span>|
|[<span data-ttu-id="462e2-980">Zutreffender Outlook-Modus</span><span class="sxs-lookup"><span data-stu-id="462e2-980">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="462e2-981">Verfassen</span><span class="sxs-lookup"><span data-stu-id="462e2-981">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="462e2-982">Beispiel</span><span class="sxs-lookup"><span data-stu-id="462e2-982">Example</span></span>

<span data-ttu-id="462e2-983">Mit dem folgende Code wird eine Anlage mit dem Bezeichner "0" entfernt.</span><span class="sxs-lookup"><span data-stu-id="462e2-983">The following code removes an attachment with an identifier of '0'.</span></span>

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

####  <a name="saveasyncoptions-callback"></a><span data-ttu-id="462e2-984">saveAsync([options], callback)</span><span class="sxs-lookup"><span data-stu-id="462e2-984">saveAsync([options], callback)</span></span>

<span data-ttu-id="462e2-985">Speicher asynchron ein Element. .</span><span class="sxs-lookup"><span data-stu-id="462e2-985">Asynchronously saves an item.</span></span>

<span data-ttu-id="462e2-p167">Beim Aufrufen speichert diese Methode die aktuelle Nachricht als Entwurf und  gibt die Element-ID über die Callbackmethode zurück. In Outlook Web App oder Outlook im Onlinemodus wird das Element auf dem Server gespeichert. In Outlook im Cache-Modus wird das Element im lokalen Cache gespeichert.</span><span class="sxs-lookup"><span data-stu-id="462e2-p167">When invoked, this method saves the current message as a draft and returns the item id via the callback method. In Outlook Web App or Outlook in online mode, the item is saved to the server. In Outlook in cached mode, the item is saved to the local cache.</span></span>

> [!NOTE]
> <span data-ttu-id="462e2-989">Wenn Ihr Add-In ruft `saveAsync` für ein Element im Entwurfsmodus, um das Abrufen einer `itemId` um mit EWS oder REST-API verwenden, beachten Sie, dass wenn Outlook im Cache-Modus ist, es kann einige Zeit dauern, bevor das Element tatsächlich auf dem Server synchronisiert wird.</span><span class="sxs-lookup"><span data-stu-id="462e2-989">If your add-in calls `saveAsync` on an item in compose mode in order to get an `itemId` to use with EWS or the REST API, be aware that when Outlook is in cached mode, it may take some time before the item is actually synced to the server.</span></span> <span data-ttu-id="462e2-990">Bis das Element mit synchronisiert ist, die `itemId` wird ein Fehler zurückgegeben.</span><span class="sxs-lookup"><span data-stu-id="462e2-990">Until the item is synced, using the `itemId` will return an error.</span></span>

<span data-ttu-id="462e2-p169">Da Termine keinen Entwurfsstatus haben, wird das Element bei Aufruf von `saveAsync` für einen Termin im Verfassenmodus als normaler Termin im Kalender des Benutzers gespeichert. Für neue Termine, die noch nicht gespeichert wurden, wird keine Einladung gesendet. Beim Speichern eines vorhandenen Termins wird eine Aktualisierung an die hinzugefügten oder entfernten Teilnehmer gesendet.</span><span class="sxs-lookup"><span data-stu-id="462e2-p169">Since appointments have no draft state, if `saveAsync` is called on an appointment in compose mode, the item will be saved as a normal appointment on the user's calendar. For new appointments that have not been saved before, no invitation will be sent. Saving an existing appointment will send an update to added or removed attendees.</span></span>

> [!NOTE]
> <span data-ttu-id="462e2-994">Die folgenden Clients haben unterschiedlichem Verhalten für `saveAsync` Termine im Verfassenmodus:</span><span class="sxs-lookup"><span data-stu-id="462e2-994">The following clients have different behavior for `saveAsync` on appointments in compose mode:</span></span>
>
> - <span data-ttu-id="462e2-995">Outlook für Mac unterstützt keine `saveAsync` auf einer Besprechung im Verfassenmodus.</span><span class="sxs-lookup"><span data-stu-id="462e2-995">Mac Outlook does not support `saveAsync` on a meeting in compose mode.</span></span> <span data-ttu-id="462e2-996">Aufrufen von `saveAsync` auf einer Besprechung in Outlook für Mac wird ein Fehler zurückgegeben.</span><span class="sxs-lookup"><span data-stu-id="462e2-996">Calling `saveAsync` on a meeting in Mac Outlook will return an error.</span></span>
> - <span data-ttu-id="462e2-997">Outlook im Web immer sendet eine Einladung oder zu aktualisieren, wenn `saveAsync` für einen Termin aufgerufen wird, im Verfassenmodus.</span><span class="sxs-lookup"><span data-stu-id="462e2-997">Outlook on the web always sends an invitation or update when `saveAsync` is called on an appointment in compose mode.</span></span>

##### <a name="parameters"></a><span data-ttu-id="462e2-998">Parameter:</span><span class="sxs-lookup"><span data-stu-id="462e2-998">Parameters:</span></span>

|<span data-ttu-id="462e2-999">Name</span><span class="sxs-lookup"><span data-stu-id="462e2-999">Name</span></span>| <span data-ttu-id="462e2-1000">Typ</span><span class="sxs-lookup"><span data-stu-id="462e2-1000">Type</span></span>| <span data-ttu-id="462e2-1001">Attribute</span><span class="sxs-lookup"><span data-stu-id="462e2-1001">Attributes</span></span>| <span data-ttu-id="462e2-1002">Beschreibung</span><span class="sxs-lookup"><span data-stu-id="462e2-1002">Description</span></span>|
|---|---|---|---|
|`options`| <span data-ttu-id="462e2-1003">Objekt</span><span class="sxs-lookup"><span data-stu-id="462e2-1003">Object</span></span>| <span data-ttu-id="462e2-1004">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="462e2-1004">&lt;optional&gt;</span></span>|<span data-ttu-id="462e2-1005">Ein Objektliteral, das mindestens eine der folgenden Eigenschaften enthält.</span><span class="sxs-lookup"><span data-stu-id="462e2-1005">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`| <span data-ttu-id="462e2-1006">Object</span><span class="sxs-lookup"><span data-stu-id="462e2-1006">Object</span></span>| <span data-ttu-id="462e2-1007">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="462e2-1007">&lt;optional&gt;</span></span>|<span data-ttu-id="462e2-1008">Entwickler können ein Objekt bereitstellen, auf das sie in der Callbackmethode zugreifen möchten.</span><span class="sxs-lookup"><span data-stu-id="462e2-1008">Developers can provide any object they wish to access in the callback method.</span></span>||
|`callback`| <span data-ttu-id="462e2-1009">Funktion</span><span class="sxs-lookup"><span data-stu-id="462e2-1009">function</span></span>||<span data-ttu-id="462e2-1010">Wenn die Methode abgeschlossen wird, wird die Funktion , die im `callback`-Parameter übergeben wird, mit einem einzelnen Parameter (`asyncResult`) aufgerufen, der ein [`AsyncResult`](/javascript/api/office/office.asyncresult)-Objekt darstellt.</span><span class="sxs-lookup"><span data-stu-id="462e2-1010">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span><br/><br/><span data-ttu-id="462e2-1011">Auf Erfolg, erfolgt die Element-ID der `asyncResult.value` Eigenschaft.</span><span class="sxs-lookup"><span data-stu-id="462e2-1011">On success, the item identifier is provided in the `asyncResult.value` property.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="462e2-1012">Anforderungen</span><span class="sxs-lookup"><span data-stu-id="462e2-1012">Requirements</span></span>

|<span data-ttu-id="462e2-1013">Anforderung</span><span class="sxs-lookup"><span data-stu-id="462e2-1013">Requirement</span></span>| <span data-ttu-id="462e2-1014">Wert</span><span class="sxs-lookup"><span data-stu-id="462e2-1014">Value</span></span>|
|---|---|
|[<span data-ttu-id="462e2-1015">Mindestversion des Postfachanforderungssatzes</span><span class="sxs-lookup"><span data-stu-id="462e2-1015">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="462e2-1016">1.3</span><span class="sxs-lookup"><span data-stu-id="462e2-1016">1.3</span></span>|
|[<span data-ttu-id="462e2-1017">Mindestberechtigungsstufe</span><span class="sxs-lookup"><span data-stu-id="462e2-1017">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="462e2-1018">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="462e2-1018">ReadWriteItem</span></span>|
|[<span data-ttu-id="462e2-1019">Zutreffender Outlook-Modus</span><span class="sxs-lookup"><span data-stu-id="462e2-1019">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="462e2-1020">Verfassen</span><span class="sxs-lookup"><span data-stu-id="462e2-1020">Compose</span></span>|

##### <a name="examples"></a><span data-ttu-id="462e2-1021">Beispiele</span><span class="sxs-lookup"><span data-stu-id="462e2-1021">Examples</span></span>

```
Office.context.mailbox.item.saveAsync(
  function callback(result) {
    // Process the result
  });
```

<span data-ttu-id="462e2-p171">Es folgt ein Beispiel für den `result`-Parameter, der an die Callbackfunktion übergeben wird. Die `value`-Eigenschaft enthält die Element-ID des Elements.</span><span class="sxs-lookup"><span data-stu-id="462e2-p171">The following is an example of the `result` parameter passed to the callback function. The `value` property contains the item ID of the item.</span></span>

```
{
  "value":"AAMkADI5...AAA=",
  "status":"succeeded"
}
```

####  <a name="setselecteddataasyncdata-options-callback"></a><span data-ttu-id="462e2-1024">setSelectedDataAsync(data, [options], callback)</span><span class="sxs-lookup"><span data-stu-id="462e2-1024">setSelectedDataAsync(data, [options], callback)</span></span>

<span data-ttu-id="462e2-1025">Fügt asynchron Daten in den Textkörper oder Betreff einer Nachricht ein.</span><span class="sxs-lookup"><span data-stu-id="462e2-1025">Asynchronously inserts data into the body or subject of a message.</span></span>

<span data-ttu-id="462e2-p172">Die `setSelectedDataAsync`-Methode fügt die angegebene Zeichenfolge an der Cursorposition im Betreff oder im Textkörper ein, oder, falls im Editor Text ausgewählt ist, ersetzt den markierten Text. Wenn sich der Cursor nicht im Textkörper oder im Betreffsfeld befindet, wird ein Fehler zurückgegeben. Nach dem Einfügen wird der Cursor am Ende der eingefügten Inhalte platziert.</span><span class="sxs-lookup"><span data-stu-id="462e2-p172">The `setSelectedDataAsync` method inserts the specified string at the cursor location in the subject or body of the item, or, if text is selected in the editor, it replaces the selected text. If the cursor is not in the body or subject field, an error is returned. After insertion, the cursor is placed at the end of the inserted content.</span></span>

##### <a name="parameters"></a><span data-ttu-id="462e2-1029">Parameter:</span><span class="sxs-lookup"><span data-stu-id="462e2-1029">Parameters:</span></span>

|<span data-ttu-id="462e2-1030">Name</span><span class="sxs-lookup"><span data-stu-id="462e2-1030">Name</span></span>| <span data-ttu-id="462e2-1031">Typ</span><span class="sxs-lookup"><span data-stu-id="462e2-1031">Type</span></span>| <span data-ttu-id="462e2-1032">Attribute</span><span class="sxs-lookup"><span data-stu-id="462e2-1032">Attributes</span></span>| <span data-ttu-id="462e2-1033">Beschreibung</span><span class="sxs-lookup"><span data-stu-id="462e2-1033">Description</span></span>|
|---|---|---|---|
|`data`| <span data-ttu-id="462e2-1034">Zeichenfolge</span><span class="sxs-lookup"><span data-stu-id="462e2-1034">String</span></span>||<span data-ttu-id="462e2-p173">Die einzufügenden Daten. Daten dürfen 1.000.000 Zeichen nicht überschreiten. Werden mehr als 1.000.000 Zeichen übergeben, wird eine `ArgumentOutOfRange`-Ausnahme ausgelöst.</span><span class="sxs-lookup"><span data-stu-id="462e2-p173">The data to be inserted. Data is not to exceed 1,000,000 characters. If more than 1,000,000 characters are passed in, an `ArgumentOutOfRange` exception is thrown.</span></span>|
|`options`| <span data-ttu-id="462e2-1038">Object</span><span class="sxs-lookup"><span data-stu-id="462e2-1038">Object</span></span>| <span data-ttu-id="462e2-1039">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="462e2-1039">&lt;optional&gt;</span></span>|<span data-ttu-id="462e2-1040">Ein Objektliteral, das mindestens eine der folgenden Eigenschaften enthält.</span><span class="sxs-lookup"><span data-stu-id="462e2-1040">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`| <span data-ttu-id="462e2-1041">Object</span><span class="sxs-lookup"><span data-stu-id="462e2-1041">Object</span></span>| <span data-ttu-id="462e2-1042">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="462e2-1042">&lt;optional&gt;</span></span>|<span data-ttu-id="462e2-1043">Entwickler können ein Objekt bereitstellen, auf das sie in der Rückrufmethode zugreifen möchten.</span><span class="sxs-lookup"><span data-stu-id="462e2-1043">Developers can provide any object they wish to access in the callback method.</span></span>|
|`options.coercionType`| [<span data-ttu-id="462e2-1044">Office.CoercionType</span><span class="sxs-lookup"><span data-stu-id="462e2-1044">Office.CoercionType</span></span>](office.md#coerciontype-string)| <span data-ttu-id="462e2-1045">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="462e2-1045">&lt;optional&gt;</span></span>|<span data-ttu-id="462e2-p174">Wenn `text`, wird das aktuelle Format in Outlook Web App und Outlook angewendet. Wenn das Feld ein HTML-Editor ist, werden nur die Textdaten eingefügt, selbst wenn es sich bei den Daten um HTML-Daten handelt.</span><span class="sxs-lookup"><span data-stu-id="462e2-p174">If `text`, the current style is applied in Outlook Web App and Outlook. If the field is an HTML editor, only the text data is inserted, even if the data is HTML.</span></span><br/><br/><span data-ttu-id="462e2-p175">Wenn `html` gesetzt ist und das Feld HTML unterstützt (der Betreff jedoch nicht), wird die aktuelle Formatvorlage in Outlook Web App angewendet und die Standardformatvorlage in Outlook. Ist das Feld ein Textfeld, wird ein Fehler des Typs `InvalidDataFormat` zurückgegeben.</span><span class="sxs-lookup"><span data-stu-id="462e2-p175">If `html` and the field supports HTML (the subject doesn't), the current style is applied in Outlook Web App and the default style is applied in Outlook. If the field is a text field, an `InvalidDataFormat` error is returned.</span></span><br/><br/><span data-ttu-id="462e2-1050">Wenn `coercionType` nicht festgelegt wird, hängt das Ergebnis vom Feld ab: Wenn das Feld HTML ist, wird HTML verwendet. Wenn das Feld Text ist, wird Nur-Text verwendet.</span><span class="sxs-lookup"><span data-stu-id="462e2-1050">If `coercionType` is not set, the result depends on the field: if the field is HTML then HTML is used; if the field is text, then plain text is used.</span></span>|
|`callback`| <span data-ttu-id="462e2-1051">function</span><span class="sxs-lookup"><span data-stu-id="462e2-1051">function</span></span>||<span data-ttu-id="462e2-1052">Wenn die Methode abgeschlossen wird, wird die Funktion , die im `callback`-Parameter übergeben wird, mit einem einzelnen Parameter (`asyncResult`) aufgerufen, der ein [`AsyncResult`](/javascript/api/office/office.asyncresult)-Objekt darstellt.</span><span class="sxs-lookup"><span data-stu-id="462e2-1052">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="462e2-1053">Anforderungen</span><span class="sxs-lookup"><span data-stu-id="462e2-1053">Requirements</span></span>

|<span data-ttu-id="462e2-1054">Anforderung</span><span class="sxs-lookup"><span data-stu-id="462e2-1054">Requirement</span></span>| <span data-ttu-id="462e2-1055">Wert</span><span class="sxs-lookup"><span data-stu-id="462e2-1055">Value</span></span>|
|---|---|
|[<span data-ttu-id="462e2-1056">Mindestversion des Postfachanforderungssatzes</span><span class="sxs-lookup"><span data-stu-id="462e2-1056">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="462e2-1057">1.2</span><span class="sxs-lookup"><span data-stu-id="462e2-1057">1.2</span></span>|
|[<span data-ttu-id="462e2-1058">Mindestberechtigungsstufe</span><span class="sxs-lookup"><span data-stu-id="462e2-1058">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="462e2-1059">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="462e2-1059">ReadWriteItem</span></span>|
|[<span data-ttu-id="462e2-1060">Zutreffender Outlook-Modus</span><span class="sxs-lookup"><span data-stu-id="462e2-1060">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="462e2-1061">Verfassen</span><span class="sxs-lookup"><span data-stu-id="462e2-1061">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="462e2-1062">Beispiel</span><span class="sxs-lookup"><span data-stu-id="462e2-1062">Example</span></span>

```
Office.context.mailbox.item.setSelectedDataAsync("Hello World!");
Office.context.mailbox.item.setSelectedDataAsync("<b>Hello World!</b>", { coercionType : "html" });
```