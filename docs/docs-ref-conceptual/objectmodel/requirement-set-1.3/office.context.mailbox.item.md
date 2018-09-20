
# <a name="item"></a><span data-ttu-id="496ae-101">item</span><span class="sxs-lookup"><span data-stu-id="496ae-101">item</span></span>

### <a name="officeofficemdcontextofficecontextmdmailboxofficecontextmailboxmditem"></a><span data-ttu-id="496ae-102">[Office](office.md)[.context](office.context.md)[.mailbox](office.context.mailbox.md).item</span><span class="sxs-lookup"><span data-stu-id="496ae-102">[Office](office.md)[.context](office.context.md)[.mailbox](office.context.mailbox.md).item</span></span>

<span data-ttu-id="496ae-p101">Der `item`-Namespace wird für den Zugriff auf die aktuell ausgewählte Nachricht, Besprechungsanfrage oder den aktuell ausgewählten Termin verwendet. Sie können den Typ von `item` mithilfe der [itemType](#itemtype-officemailboxenumsitemtypejavascriptapioutlook13officemailboxenumsitemtype)-Eigenschaft bestimmen.</span><span class="sxs-lookup"><span data-stu-id="496ae-p101">The `item` namespace is used to access the currently selected message, meeting request, or appointment. You can determine the type of the `item` by using the [itemType](#itemtype-officemailboxenumsitemtypejavascriptapioutlook13officemailboxenumsitemtype) property.</span></span>

##### <a name="requirements"></a><span data-ttu-id="496ae-105">Anforderungen</span><span class="sxs-lookup"><span data-stu-id="496ae-105">Requirements</span></span>

|<span data-ttu-id="496ae-106">Anforderung</span><span class="sxs-lookup"><span data-stu-id="496ae-106">Requirement</span></span>| <span data-ttu-id="496ae-107">Wert</span><span class="sxs-lookup"><span data-stu-id="496ae-107">Value</span></span>|
|---|---|
|[<span data-ttu-id="496ae-108">Mindestversion des Postfachanforderungssatzes</span><span class="sxs-lookup"><span data-stu-id="496ae-108">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="496ae-109">1.0</span><span class="sxs-lookup"><span data-stu-id="496ae-109">1.0</span></span>|
|[<span data-ttu-id="496ae-110">Mindestberechtigungsstufe</span><span class="sxs-lookup"><span data-stu-id="496ae-110">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="496ae-111">Eingeschränkt</span><span class="sxs-lookup"><span data-stu-id="496ae-111">Restricted</span></span>|
|[<span data-ttu-id="496ae-112">Zutreffender Outlook-Modus</span><span class="sxs-lookup"><span data-stu-id="496ae-112">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="496ae-113">Verfassen oder Lesen</span><span class="sxs-lookup"><span data-stu-id="496ae-113">Compose or read</span></span>|

### <a name="example"></a><span data-ttu-id="496ae-114">Beispiel</span><span class="sxs-lookup"><span data-stu-id="496ae-114">Example</span></span>

<span data-ttu-id="496ae-115">Im folgenden JavaScript-Codebeispiel wird der Zugriff auf die `subject`-Eigenschaft des aktuellen Elements in Outlook veranschaulicht.</span><span class="sxs-lookup"><span data-stu-id="496ae-115">The following JavaScript code example shows how to access the `subject` property of the current item in Outlook.</span></span>

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

### <a name="members"></a><span data-ttu-id="496ae-116">Elemente</span><span class="sxs-lookup"><span data-stu-id="496ae-116">Members</span></span>

#### <a name="attachments-arrayattachmentdetailsjavascriptapioutlook13officeattachmentdetails"></a><span data-ttu-id="496ae-117">attachments :Array.<[AttachmentDetails](/javascript/api/outlook_1_3/office.attachmentdetails)></span><span class="sxs-lookup"><span data-stu-id="496ae-117">attachments :Array.<[AttachmentDetails](/javascript/api/outlook_1_3/office.attachmentdetails)></span></span>

<span data-ttu-id="496ae-p102">Ruft ein Array mit Anlagen für das Element ab. Nur Lesemodus.</span><span class="sxs-lookup"><span data-stu-id="496ae-p102">Gets an array of attachments for the item. Read mode only.</span></span>

> [!NOTE]
> <span data-ttu-id="496ae-120">Bestimmte Dateitypen werden aufgrund von Sicherheitsproblemen von Outlook blockiert und daher nicht zurückgegeben werden.</span><span class="sxs-lookup"><span data-stu-id="496ae-120">Certain types of files are blocked by Outlook due to potential security issues and are therefore not returned.</span></span> <span data-ttu-id="496ae-121">Weitere Informationen finden Sie unter [Blockierte Anlagen in Outlook](https://support.office.com/article/Blocked-attachments-in-Outlook-434752E1-02D3-4E90-9124-8B81E49A8519).</span><span class="sxs-lookup"><span data-stu-id="496ae-121">For more information, see [Blocked attachments in Outlook](https://support.office.com/article/Blocked-attachments-in-Outlook-434752E1-02D3-4E90-9124-8B81E49A8519).</span></span>

##### <a name="type"></a><span data-ttu-id="496ae-122">Typ:</span><span class="sxs-lookup"><span data-stu-id="496ae-122">Type:</span></span>

*   <span data-ttu-id="496ae-123">Array.<[AttachmentDetails](/javascript/api/outlook_1_3/office.attachmentdetails)></span><span class="sxs-lookup"><span data-stu-id="496ae-123">Array.<[AttachmentDetails](/javascript/api/outlook_1_3/office.attachmentdetails)></span></span>

##### <a name="requirements"></a><span data-ttu-id="496ae-124">Anforderungen</span><span class="sxs-lookup"><span data-stu-id="496ae-124">Requirements</span></span>

|<span data-ttu-id="496ae-125">Anforderung</span><span class="sxs-lookup"><span data-stu-id="496ae-125">Requirement</span></span>| <span data-ttu-id="496ae-126">Wert</span><span class="sxs-lookup"><span data-stu-id="496ae-126">Value</span></span>|
|---|---|
|[<span data-ttu-id="496ae-127">Mindestversion des Postfachanforderungssatzes</span><span class="sxs-lookup"><span data-stu-id="496ae-127">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="496ae-128">1.0</span><span class="sxs-lookup"><span data-stu-id="496ae-128">1.0</span></span>|
|[<span data-ttu-id="496ae-129">Mindestberechtigungsstufe</span><span class="sxs-lookup"><span data-stu-id="496ae-129">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="496ae-130">ReadItem</span><span class="sxs-lookup"><span data-stu-id="496ae-130">ReadItem</span></span>|
|[<span data-ttu-id="496ae-131">Zutreffender Outlook-Modus</span><span class="sxs-lookup"><span data-stu-id="496ae-131">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="496ae-132">Lesen</span><span class="sxs-lookup"><span data-stu-id="496ae-132">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="496ae-133">Beispiel</span><span class="sxs-lookup"><span data-stu-id="496ae-133">Example</span></span>

<span data-ttu-id="496ae-134">Mit dem folgende Code wird eine HTML-Zeichenfolge mit Details aller Anlagen für das aktuelle Element erstellt.</span><span class="sxs-lookup"><span data-stu-id="496ae-134">The following code builds an HTML string with details of all attachments on the current item.</span></span>

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

####  <a name="bcc-recipientsjavascriptapioutlook13officerecipients"></a><span data-ttu-id="496ae-135">bcc :[Recipients](/javascript/api/outlook_1_3/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="496ae-135">bcc :[Recipients](/javascript/api/outlook_1_3/office.recipients)</span></span>

<span data-ttu-id="496ae-136">Ruft ein Objekt, das Methoden zum Abrufen oder aktualisieren die Empfänger in der Zeile Bcc (blind Carbon Copy, Blindkopie) einer Nachricht bereitstellt.</span><span class="sxs-lookup"><span data-stu-id="496ae-136">Gets an object that provides methods to get or update the recipients on the Bcc (blind carbon copy) line of a message.</span></span> <span data-ttu-id="496ae-137">Nur Verfassenmodus.</span><span class="sxs-lookup"><span data-stu-id="496ae-137">Compose mode only.</span></span>

##### <a name="type"></a><span data-ttu-id="496ae-138">Typ:</span><span class="sxs-lookup"><span data-stu-id="496ae-138">Type:</span></span>

*   [<span data-ttu-id="496ae-139">Recipients</span><span class="sxs-lookup"><span data-stu-id="496ae-139">Recipients</span></span>](/javascript/api/outlook_1_3/office.recipients)

##### <a name="requirements"></a><span data-ttu-id="496ae-140">Anforderungen</span><span class="sxs-lookup"><span data-stu-id="496ae-140">Requirements</span></span>

|<span data-ttu-id="496ae-141">Anforderung</span><span class="sxs-lookup"><span data-stu-id="496ae-141">Requirement</span></span>| <span data-ttu-id="496ae-142">Wert</span><span class="sxs-lookup"><span data-stu-id="496ae-142">Value</span></span>|
|---|---|
|[<span data-ttu-id="496ae-143">Mindestversion des Postfachanforderungssatzes</span><span class="sxs-lookup"><span data-stu-id="496ae-143">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="496ae-144">1.1</span><span class="sxs-lookup"><span data-stu-id="496ae-144">1.1</span></span>|
|[<span data-ttu-id="496ae-145">Mindestberechtigungsstufe</span><span class="sxs-lookup"><span data-stu-id="496ae-145">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="496ae-146">ReadItem</span><span class="sxs-lookup"><span data-stu-id="496ae-146">ReadItem</span></span>|
|[<span data-ttu-id="496ae-147">Zutreffender Outlook-Modus</span><span class="sxs-lookup"><span data-stu-id="496ae-147">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="496ae-148">Verfassen</span><span class="sxs-lookup"><span data-stu-id="496ae-148">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="496ae-149">Beispiel</span><span class="sxs-lookup"><span data-stu-id="496ae-149">Example</span></span>

```
Office.context.mailbox.item.bcc.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.bcc.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.bcc.getAsync(callback);

function callback(asyncResult) {
  var arrayOfBccRecipients = asyncResult.value;
}
```

####  <a name="body-bodyjavascriptapioutlook13officebody"></a><span data-ttu-id="496ae-150">body :[Body](/javascript/api/outlook_1_3/office.body)</span><span class="sxs-lookup"><span data-stu-id="496ae-150">body :[Body](/javascript/api/outlook_1_3/office.body)</span></span>

<span data-ttu-id="496ae-151">Ruft ein Objekt ab, das Methoden zum Bearbeiten des Textkörpers eines Elements bereitstellt.</span><span class="sxs-lookup"><span data-stu-id="496ae-151">Gets an object that provides methods for manipulating the body of an item.</span></span>

##### <a name="type"></a><span data-ttu-id="496ae-152">Typ:</span><span class="sxs-lookup"><span data-stu-id="496ae-152">Type:</span></span>

*   [<span data-ttu-id="496ae-153">Body</span><span class="sxs-lookup"><span data-stu-id="496ae-153">Body</span></span>](/javascript/api/outlook_1_3/office.body)

##### <a name="requirements"></a><span data-ttu-id="496ae-154">Anforderungen</span><span class="sxs-lookup"><span data-stu-id="496ae-154">Requirements</span></span>

|<span data-ttu-id="496ae-155">Anforderung</span><span class="sxs-lookup"><span data-stu-id="496ae-155">Requirement</span></span>| <span data-ttu-id="496ae-156">Wert</span><span class="sxs-lookup"><span data-stu-id="496ae-156">Value</span></span>|
|---|---|
|[<span data-ttu-id="496ae-157">Mindestversion des Postfachanforderungssatzes</span><span class="sxs-lookup"><span data-stu-id="496ae-157">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="496ae-158">1.1</span><span class="sxs-lookup"><span data-stu-id="496ae-158">1.1</span></span>|
|[<span data-ttu-id="496ae-159">Mindestberechtigungsstufe</span><span class="sxs-lookup"><span data-stu-id="496ae-159">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="496ae-160">ReadItem</span><span class="sxs-lookup"><span data-stu-id="496ae-160">ReadItem</span></span>|
|[<span data-ttu-id="496ae-161">Zutreffender Outlook-Modus</span><span class="sxs-lookup"><span data-stu-id="496ae-161">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="496ae-162">Verfassen oder Lesen</span><span class="sxs-lookup"><span data-stu-id="496ae-162">Compose or read</span></span>|

####  <a name="cc-arrayemailaddressdetailsjavascriptapioutlook13officeemailaddressdetailsrecipientsjavascriptapioutlook13officerecipients"></a><span data-ttu-id="496ae-163">cc: Array. <[EmailAddressDetails](/javascript/api/outlook_1_3/office.emailaddressdetails)>|[Empfänger](/javascript/api/outlook_1_3/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="496ae-163">cc :Array.<[EmailAddressDetails](/javascript/api/outlook_1_3/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook_1_3/office.recipients)</span></span>

<span data-ttu-id="496ae-164">Ermöglicht den Zugriff auf die (Carbon Copy, Blindkopie) Cc-Empfänger einer Nachricht.</span><span class="sxs-lookup"><span data-stu-id="496ae-164">Provides access to the Cc (carbon copy) recipients of a message.</span></span> <span data-ttu-id="496ae-165">Der Typ des Objekts und die Zugriffsebene, hängt von den Modus des aktuellen Elements ab.</span><span class="sxs-lookup"><span data-stu-id="496ae-165">The type of object and level of access depends on the mode of the current item.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="496ae-166">Lesemodus</span><span class="sxs-lookup"><span data-stu-id="496ae-166">Read mode</span></span>

<span data-ttu-id="496ae-p106">Die `cc`-Eigenschaft gibt ein Array mit einem `EmailAddressDetails`-Objekt für jeden Empfänger in der **Cc**-Zeile der Nachricht zurück. Die Auflistung ist auf höchstens 100 Elemente beschränkt.</span><span class="sxs-lookup"><span data-stu-id="496ae-p106">The `cc` property returns an array that contains an `EmailAddressDetails` object for each recipient listed on the **Cc** line of the message. The collection is limited to a maximum of 100 members.</span></span>

##### <a name="compose-mode"></a><span data-ttu-id="496ae-169">Verfassenmodus</span><span class="sxs-lookup"><span data-stu-id="496ae-169">Compose mode</span></span>

<span data-ttu-id="496ae-170">Die `cc` -Eigenschaft gibt eine `Recipients` -Objekt, das Methoden zum Abrufen oder aktualisieren die Empfänger in der Zeile **Cc** der Nachricht bereitstellt.</span><span class="sxs-lookup"><span data-stu-id="496ae-170">The `cc` property returns a `Recipients` object that provides methods to get or update the recipients on the **Cc** line of the message.</span></span>

##### <a name="type"></a><span data-ttu-id="496ae-171">Typ:</span><span class="sxs-lookup"><span data-stu-id="496ae-171">Type:</span></span>

*   <span data-ttu-id="496ae-172">Array.<[EmailAddressDetails](/javascript/api/outlook_1_3/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook_1_3/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="496ae-172">Array.<[EmailAddressDetails](/javascript/api/outlook_1_3/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook_1_3/office.recipients)</span></span>

##### <a name="requirements"></a><span data-ttu-id="496ae-173">Anforderungen</span><span class="sxs-lookup"><span data-stu-id="496ae-173">Requirements</span></span>

|<span data-ttu-id="496ae-174">Anforderung</span><span class="sxs-lookup"><span data-stu-id="496ae-174">Requirement</span></span>| <span data-ttu-id="496ae-175">Wert</span><span class="sxs-lookup"><span data-stu-id="496ae-175">Value</span></span>|
|---|---|
|[<span data-ttu-id="496ae-176">Mindestversion des Postfachanforderungssatzes</span><span class="sxs-lookup"><span data-stu-id="496ae-176">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="496ae-177">1.0</span><span class="sxs-lookup"><span data-stu-id="496ae-177">1.0</span></span>|
|[<span data-ttu-id="496ae-178">Mindestberechtigungsstufe</span><span class="sxs-lookup"><span data-stu-id="496ae-178">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="496ae-179">ReadItem</span><span class="sxs-lookup"><span data-stu-id="496ae-179">ReadItem</span></span>|
|[<span data-ttu-id="496ae-180">Zutreffender Outlook-Modus</span><span class="sxs-lookup"><span data-stu-id="496ae-180">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="496ae-181">Verfassen oder Lesen</span><span class="sxs-lookup"><span data-stu-id="496ae-181">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="496ae-182">Beispiel</span><span class="sxs-lookup"><span data-stu-id="496ae-182">Example</span></span>

```
Office.context.mailbox.item.cc.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.cc.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.cc.getAsync(callback);

function callback(asyncResult) {
  var arrayOfCcRecipients = asyncResult.value;
}
```

####  <a name="nullable-conversationid-string"></a><span data-ttu-id="496ae-183">(nullable) conversationId :String</span><span class="sxs-lookup"><span data-stu-id="496ae-183">(nullable) conversationId :String</span></span>

<span data-ttu-id="496ae-184">Ruft einen Bezeichner für die E-Mail-Unterhaltung ab, in der eine bestimmte Nachricht enthalten ist.</span><span class="sxs-lookup"><span data-stu-id="496ae-184">Gets an identifier for the email conversation that contains a particular message.</span></span>

<span data-ttu-id="496ae-p107">Sie können für diese Eigenschaft eine ganze Zahl abrufen, wenn Ihre Mail-App in Formularen zum Lesen oder Antworten in Formularen zum Verfassen aktiviert wird. Wenn der Benutzer den Betreff der Antwortnachricht ändert, ändert sich beim Versenden die Konversations-ID für die entsprechende Nachricht, und der Wert, den Sie vorher bezogen haben, trifft nicht länger zu.</span><span class="sxs-lookup"><span data-stu-id="496ae-p107">You can get an integer for this property if your mail app is activated in read forms or responses in compose forms. If subsequently the user changes the subject of the reply message, upon sending the reply, the conversation ID for that message will change and that value you obtained earlier will no longer apply.</span></span>

<span data-ttu-id="496ae-p108">Sie erhalten in einem Formular zum Verfassen für diese Eigenschaft für ein neues Element null. Wenn der Benutzer einen Betreff festlegt und das Element speichert, gibt die `conversationId`-Eigenschaft einen Wert zurück.</span><span class="sxs-lookup"><span data-stu-id="496ae-p108">You get null for this property for a new item in a compose form. If the user sets a subject and saves the item, the `conversationId` property will return a value.</span></span>

##### <a name="type"></a><span data-ttu-id="496ae-189">Typ:</span><span class="sxs-lookup"><span data-stu-id="496ae-189">Type:</span></span>

*   <span data-ttu-id="496ae-190">String</span><span class="sxs-lookup"><span data-stu-id="496ae-190">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="496ae-191">Anforderungen</span><span class="sxs-lookup"><span data-stu-id="496ae-191">Requirements</span></span>

|<span data-ttu-id="496ae-192">Anforderung</span><span class="sxs-lookup"><span data-stu-id="496ae-192">Requirement</span></span>| <span data-ttu-id="496ae-193">Wert</span><span class="sxs-lookup"><span data-stu-id="496ae-193">Value</span></span>|
|---|---|
|[<span data-ttu-id="496ae-194">Mindestversion des Postfachanforderungssatzes</span><span class="sxs-lookup"><span data-stu-id="496ae-194">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="496ae-195">1.0</span><span class="sxs-lookup"><span data-stu-id="496ae-195">1.0</span></span>|
|[<span data-ttu-id="496ae-196">Mindestberechtigungsstufe</span><span class="sxs-lookup"><span data-stu-id="496ae-196">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="496ae-197">ReadItem</span><span class="sxs-lookup"><span data-stu-id="496ae-197">ReadItem</span></span>|
|[<span data-ttu-id="496ae-198">Zutreffender Outlook-Modus</span><span class="sxs-lookup"><span data-stu-id="496ae-198">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="496ae-199">Verfassen oder Lesen</span><span class="sxs-lookup"><span data-stu-id="496ae-199">Compose or read</span></span>|

#### <a name="datetimecreated-date"></a><span data-ttu-id="496ae-200">dateTimeCreated :Date</span><span class="sxs-lookup"><span data-stu-id="496ae-200">dateTimeCreated :Date</span></span>

<span data-ttu-id="496ae-p109">Ruft das Datum und die Uhrzeit der Erstellung eines Elements ab. Nur Lesemodus.</span><span class="sxs-lookup"><span data-stu-id="496ae-p109">Gets the date and time that an item was created. Read mode only.</span></span>

##### <a name="type"></a><span data-ttu-id="496ae-203">Typ:</span><span class="sxs-lookup"><span data-stu-id="496ae-203">Type:</span></span>

*   <span data-ttu-id="496ae-204">Datum</span><span class="sxs-lookup"><span data-stu-id="496ae-204">Date</span></span>

##### <a name="requirements"></a><span data-ttu-id="496ae-205">Anforderungen</span><span class="sxs-lookup"><span data-stu-id="496ae-205">Requirements</span></span>

|<span data-ttu-id="496ae-206">Anforderung</span><span class="sxs-lookup"><span data-stu-id="496ae-206">Requirement</span></span>| <span data-ttu-id="496ae-207">Wert</span><span class="sxs-lookup"><span data-stu-id="496ae-207">Value</span></span>|
|---|---|
|[<span data-ttu-id="496ae-208">Mindestversion des Postfachanforderungssatzes</span><span class="sxs-lookup"><span data-stu-id="496ae-208">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="496ae-209">1.0</span><span class="sxs-lookup"><span data-stu-id="496ae-209">1.0</span></span>|
|[<span data-ttu-id="496ae-210">Mindestberechtigungsstufe</span><span class="sxs-lookup"><span data-stu-id="496ae-210">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="496ae-211">ReadItem</span><span class="sxs-lookup"><span data-stu-id="496ae-211">ReadItem</span></span>|
|[<span data-ttu-id="496ae-212">Zutreffender Outlook-Modus</span><span class="sxs-lookup"><span data-stu-id="496ae-212">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="496ae-213">Lesen</span><span class="sxs-lookup"><span data-stu-id="496ae-213">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="496ae-214">Beispiel</span><span class="sxs-lookup"><span data-stu-id="496ae-214">Example</span></span>

```
var created = Office.context.mailbox.item.dateTimeCreated;
```

#### <a name="datetimemodified-date"></a><span data-ttu-id="496ae-215">dateTimeModified :Date</span><span class="sxs-lookup"><span data-stu-id="496ae-215">dateTimeModified :Date</span></span>

<span data-ttu-id="496ae-p110">Ruft das Datum und die Uhrzeit der letzten Änderung eines Elements ab. Nur Lesemodus.</span><span class="sxs-lookup"><span data-stu-id="496ae-p110">Gets the date and time that an item was last modified. Read mode only.</span></span>

> [!NOTE]
> <span data-ttu-id="496ae-218">Dieser Member wird in Outlook für iOS oder Outlook für Android nicht unterstützt.</span><span class="sxs-lookup"><span data-stu-id="496ae-218">This member is not supported in Outlook for iOS or Outlook for Android.</span></span>

##### <a name="type"></a><span data-ttu-id="496ae-219">Typ:</span><span class="sxs-lookup"><span data-stu-id="496ae-219">Type:</span></span>

*   <span data-ttu-id="496ae-220">Datum</span><span class="sxs-lookup"><span data-stu-id="496ae-220">Date</span></span>

##### <a name="requirements"></a><span data-ttu-id="496ae-221">Anforderungen</span><span class="sxs-lookup"><span data-stu-id="496ae-221">Requirements</span></span>

|<span data-ttu-id="496ae-222">Anforderung</span><span class="sxs-lookup"><span data-stu-id="496ae-222">Requirement</span></span>| <span data-ttu-id="496ae-223">Wert</span><span class="sxs-lookup"><span data-stu-id="496ae-223">Value</span></span>|
|---|---|
|[<span data-ttu-id="496ae-224">Mindestversion des Postfachanforderungssatzes</span><span class="sxs-lookup"><span data-stu-id="496ae-224">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="496ae-225">1.0</span><span class="sxs-lookup"><span data-stu-id="496ae-225">1.0</span></span>|
|[<span data-ttu-id="496ae-226">Mindestberechtigungsstufe</span><span class="sxs-lookup"><span data-stu-id="496ae-226">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="496ae-227">ReadItem</span><span class="sxs-lookup"><span data-stu-id="496ae-227">ReadItem</span></span>|
|[<span data-ttu-id="496ae-228">Zutreffender Outlook-Modus</span><span class="sxs-lookup"><span data-stu-id="496ae-228">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="496ae-229">Lesen</span><span class="sxs-lookup"><span data-stu-id="496ae-229">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="496ae-230">Beispiel</span><span class="sxs-lookup"><span data-stu-id="496ae-230">Example</span></span>

```
var modified = Office.context.mailbox.item.dateTimeModified;
```

####  <a name="end-datetimejavascriptapioutlook13officetime"></a><span data-ttu-id="496ae-231">end :Date|[Time](/javascript/api/outlook_1_3/office.time)</span><span class="sxs-lookup"><span data-stu-id="496ae-231">end :Date|[Time](/javascript/api/outlook_1_3/office.time)</span></span>

<span data-ttu-id="496ae-232">Ruft Datum und Zeit für das Ende des Termins ab oder legt diese fest.</span><span class="sxs-lookup"><span data-stu-id="496ae-232">Gets or sets the date and time that the appointment is to end.</span></span>

<span data-ttu-id="496ae-p111">Die `end`-Eigenschaft wird als Datums- und Uhrzeitwert in UTC ausgedrückt. Sie können die [`convertToLocalClientTime`](office.context.mailbox.md#converttolocalclienttimetimevalue--localclienttimejavascriptapioutlook13officelocalclienttime)-Methode verwenden, um den Endeigenschaftswert in den lokalen Uhrzeit- und Datumswert des Clients umzuwandeln.</span><span class="sxs-lookup"><span data-stu-id="496ae-p111">The `end` property is expressed as a Coordinated Universal Time (UTC) date and time value. You can use the [`convertToLocalClientTime`](office.context.mailbox.md#converttolocalclienttimetimevalue--localclienttimejavascriptapioutlook13officelocalclienttime) method to convert the end property value to the client’s local date and time.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="496ae-235">Lesemodus</span><span class="sxs-lookup"><span data-stu-id="496ae-235">Read mode</span></span>

<span data-ttu-id="496ae-236">Die `end`-Eigenschaft gibt ein `Date`-Objekt zurück.</span><span class="sxs-lookup"><span data-stu-id="496ae-236">The `end` property returns a `Date` object.</span></span>

##### <a name="compose-mode"></a><span data-ttu-id="496ae-237">Verfassenmodus</span><span class="sxs-lookup"><span data-stu-id="496ae-237">Compose mode</span></span>

<span data-ttu-id="496ae-238">Die `end`-Eigenschaft gibt ein `Time`-Objekt zurück.</span><span class="sxs-lookup"><span data-stu-id="496ae-238">The `end` property returns a `Time` object.</span></span>

<span data-ttu-id="496ae-239">Wenn Sie die [`Time.setAsync`](/javascript/api/outlook_1_3/office.time#setasync-datetime--options--callback-)-Methode verwenden, um die Endzeit im Verfassenmodus festzulegen, sollten Sie die [`convertToUtcClientTime`](office.context.mailbox.md#converttoutcclienttimeinput--date)-Methode verwenden, um die Ortszeit auf dem Client für den Server in UTC umzuwandeln.</span><span class="sxs-lookup"><span data-stu-id="496ae-239">When you use the [`Time.setAsync`](/javascript/api/outlook_1_3/office.time#setasync-datetime--options--callback-) method to set the end time, you should use the [`convertToUtcClientTime`](office.context.mailbox.md#converttoutcclienttimeinput--date) method to convert the local time on the client to UTC for the server.</span></span>

##### <a name="type"></a><span data-ttu-id="496ae-240">Typ:</span><span class="sxs-lookup"><span data-stu-id="496ae-240">Type:</span></span>

*   <span data-ttu-id="496ae-241">Date | [Time](/javascript/api/outlook_1_3/office.time)</span><span class="sxs-lookup"><span data-stu-id="496ae-241">Date | [Time](/javascript/api/outlook_1_3/office.time)</span></span>

##### <a name="requirements"></a><span data-ttu-id="496ae-242">Anforderungen</span><span class="sxs-lookup"><span data-stu-id="496ae-242">Requirements</span></span>

|<span data-ttu-id="496ae-243">Anforderung</span><span class="sxs-lookup"><span data-stu-id="496ae-243">Requirement</span></span>| <span data-ttu-id="496ae-244">Wert</span><span class="sxs-lookup"><span data-stu-id="496ae-244">Value</span></span>|
|---|---|
|[<span data-ttu-id="496ae-245">Mindestversion des Postfachanforderungssatzes</span><span class="sxs-lookup"><span data-stu-id="496ae-245">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="496ae-246">1.0</span><span class="sxs-lookup"><span data-stu-id="496ae-246">1.0</span></span>|
|[<span data-ttu-id="496ae-247">Mindestberechtigungsstufe</span><span class="sxs-lookup"><span data-stu-id="496ae-247">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="496ae-248">ReadItem</span><span class="sxs-lookup"><span data-stu-id="496ae-248">ReadItem</span></span>|
|[<span data-ttu-id="496ae-249">Zutreffender Outlook-Modus</span><span class="sxs-lookup"><span data-stu-id="496ae-249">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="496ae-250">Verfassen oder Lesen</span><span class="sxs-lookup"><span data-stu-id="496ae-250">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="496ae-251">Beispiel</span><span class="sxs-lookup"><span data-stu-id="496ae-251">Example</span></span>

<span data-ttu-id="496ae-252">Im folgenden Beispiel wird die Endzeit eines Termins im Verfassenmodus mithilfe der [`setAsync`](/javascript/api/outlook_1_3/office.time#setasync-datetime--options--callback-)-Methode des `Time`-Objekts festgelegt.</span><span class="sxs-lookup"><span data-stu-id="496ae-252">The following example sets the end time of an appointment in compose mode by using the [`setAsync`](/javascript/api/outlook_1_3/office.time#setasync-datetime--options--callback-) method of the `Time` object.</span></span>

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

#### <a name="from-emailaddressdetailsjavascriptapioutlook13officeemailaddressdetails"></a><span data-ttu-id="496ae-253">from :[EmailAddressDetails](/javascript/api/outlook_1_3/office.emailaddressdetails)</span><span class="sxs-lookup"><span data-stu-id="496ae-253">from :[EmailAddressDetails](/javascript/api/outlook_1_3/office.emailaddressdetails)</span></span>

<span data-ttu-id="496ae-p112">Ruft die E-Mail-Adresse des Absenders einer Nachricht ab. Nur Lesemodus.</span><span class="sxs-lookup"><span data-stu-id="496ae-p112">Gets the email address of the sender of a message. Read mode only.</span></span>

<span data-ttu-id="496ae-p113">Die `from`- und [`sender`](#sender-emailaddressdetailsjavascriptapioutlook13officeemailaddressdetails)-Eigenschaften stellen dieselbe Person dar, außer die Nachricht wurde von einem Delegaten gesendet. In diesem Fall stellt die `from`-Eigenschaft die Stellvertretung dar, und die sender-Eigenschaft stellt den Delegaten dar.</span><span class="sxs-lookup"><span data-stu-id="496ae-p113">The `from` and [`sender`](#sender-emailaddressdetailsjavascriptapioutlook13officeemailaddressdetails) properties represent the same person unless the message is sent by a delegate. In that case, the `from` property represents the delegator, and the sender property represents the delegate.</span></span>

> [!NOTE]
> <span data-ttu-id="496ae-258">Die `recipientType` -Eigenschaft des der `EmailAddressDetails` -Objekts der `from` -Eigenschaft ist `undefined`.</span><span class="sxs-lookup"><span data-stu-id="496ae-258">The `recipientType` property of the `EmailAddressDetails` object in the `from` property is `undefined`.</span></span>

##### <a name="type"></a><span data-ttu-id="496ae-259">Typ:</span><span class="sxs-lookup"><span data-stu-id="496ae-259">Type:</span></span>

*   [<span data-ttu-id="496ae-260">EmailAddressDetails</span><span class="sxs-lookup"><span data-stu-id="496ae-260">EmailAddressDetails</span></span>](/javascript/api/outlook_1_3/office.emailaddressdetails)

##### <a name="requirements"></a><span data-ttu-id="496ae-261">Anforderungen</span><span class="sxs-lookup"><span data-stu-id="496ae-261">Requirements</span></span>

|<span data-ttu-id="496ae-262">Anforderung</span><span class="sxs-lookup"><span data-stu-id="496ae-262">Requirement</span></span>| <span data-ttu-id="496ae-263">Wert</span><span class="sxs-lookup"><span data-stu-id="496ae-263">Value</span></span>|
|---|---|
|[<span data-ttu-id="496ae-264">Mindestversion des Postfachanforderungssatzes</span><span class="sxs-lookup"><span data-stu-id="496ae-264">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="496ae-265">1.0</span><span class="sxs-lookup"><span data-stu-id="496ae-265">1.0</span></span>|
|[<span data-ttu-id="496ae-266">Mindestberechtigungsstufe</span><span class="sxs-lookup"><span data-stu-id="496ae-266">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="496ae-267">ReadItem</span><span class="sxs-lookup"><span data-stu-id="496ae-267">ReadItem</span></span>|
|[<span data-ttu-id="496ae-268">Zutreffender Outlook-Modus</span><span class="sxs-lookup"><span data-stu-id="496ae-268">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="496ae-269">Lesen</span><span class="sxs-lookup"><span data-stu-id="496ae-269">Read</span></span>|

#### <a name="internetmessageid-string"></a><span data-ttu-id="496ae-270">internetMessageId :String</span><span class="sxs-lookup"><span data-stu-id="496ae-270">internetMessageId :String</span></span>

<span data-ttu-id="496ae-p114">Ruft die Internetnachricht-ID einer E-Mail-Nachricht ab. Nur Lesemodus.</span><span class="sxs-lookup"><span data-stu-id="496ae-p114">Gets the Internet message identifier for an email message. Read mode only.</span></span>

##### <a name="type"></a><span data-ttu-id="496ae-273">Typ:</span><span class="sxs-lookup"><span data-stu-id="496ae-273">Type:</span></span>

*   <span data-ttu-id="496ae-274">String</span><span class="sxs-lookup"><span data-stu-id="496ae-274">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="496ae-275">Anforderungen</span><span class="sxs-lookup"><span data-stu-id="496ae-275">Requirements</span></span>

|<span data-ttu-id="496ae-276">Anforderung</span><span class="sxs-lookup"><span data-stu-id="496ae-276">Requirement</span></span>| <span data-ttu-id="496ae-277">Wert</span><span class="sxs-lookup"><span data-stu-id="496ae-277">Value</span></span>|
|---|---|
|[<span data-ttu-id="496ae-278">Mindestversion des Postfachanforderungssatzes</span><span class="sxs-lookup"><span data-stu-id="496ae-278">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="496ae-279">1.0</span><span class="sxs-lookup"><span data-stu-id="496ae-279">1.0</span></span>|
|[<span data-ttu-id="496ae-280">Mindestberechtigungsstufe</span><span class="sxs-lookup"><span data-stu-id="496ae-280">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="496ae-281">ReadItem</span><span class="sxs-lookup"><span data-stu-id="496ae-281">ReadItem</span></span>|
|[<span data-ttu-id="496ae-282">Zutreffender Outlook-Modus</span><span class="sxs-lookup"><span data-stu-id="496ae-282">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="496ae-283">Lesen</span><span class="sxs-lookup"><span data-stu-id="496ae-283">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="496ae-284">Beispiel</span><span class="sxs-lookup"><span data-stu-id="496ae-284">Example</span></span>

```
var internetMessageId = Office.context.mailbox.item.internetMessageId;
```

#### <a name="itemclass-string"></a><span data-ttu-id="496ae-285">itemClass :String</span><span class="sxs-lookup"><span data-stu-id="496ae-285">itemClass :String</span></span>

<span data-ttu-id="496ae-p115">Ruft die Exchange-Webdienste-Elementklasse des ausgewählten Elements ab. Nur Lesemodus.</span><span class="sxs-lookup"><span data-stu-id="496ae-p115">Gets the Exchange Web Services item class of the selected item. Read mode only.</span></span>

<span data-ttu-id="496ae-p116">Die `itemClass`-Eigenschaft gibt die Nachrichtenklasse des ausgewählten Elements an. Folgende Standardnachrichtenklassen für das Nachrichten- oder Terminelement sind vorhanden:</span><span class="sxs-lookup"><span data-stu-id="496ae-p116">The `itemClass` property specifies the message class of the selected item. The following are the default message classes for the message or appointment item.</span></span>

| <span data-ttu-id="496ae-290">Typ</span><span class="sxs-lookup"><span data-stu-id="496ae-290">Type</span></span> | <span data-ttu-id="496ae-291">Beschreibung</span><span class="sxs-lookup"><span data-stu-id="496ae-291">Description</span></span> | <span data-ttu-id="496ae-292">Elementklasse</span><span class="sxs-lookup"><span data-stu-id="496ae-292">item class</span></span> |
| --- | --- | --- |
| <span data-ttu-id="496ae-293">Terminelemente</span><span class="sxs-lookup"><span data-stu-id="496ae-293">Appointment items</span></span> | <span data-ttu-id="496ae-294">Dies sind Kalenderelemente der Elementklasse `IPM.Appointment` oder `IPM.Appointment.Occurence`.</span><span class="sxs-lookup"><span data-stu-id="496ae-294">These are calendar items of the item class `IPM.Appointment` or `IPM.Appointment.Occurence`.</span></span> | `IPM.Appointment`<br />`IPM.Appointment.Occurence` |
| <span data-ttu-id="496ae-295">Nachrichtenelemente</span><span class="sxs-lookup"><span data-stu-id="496ae-295">Message items</span></span> | <span data-ttu-id="496ae-296">Diese Elemente umfassen E-Mail-Nachrichten der Standardnachrichtenklasse `IPM.Note` sowie Besprechungsanfragen, -antworten und -absagen, die `IPM.Schedule.Meeting` als Basisnachrichtenklasse verwenden.</span><span class="sxs-lookup"><span data-stu-id="496ae-296">These include email messages that have the default message class `IPM.Note`, and meeting requests, responses, and cancellations, that use `IPM.Schedule.Meeting` as the base message class.</span></span> | `IPM.Note`<br />`IPM.Schedule.Meeting.Request`<br />`IPM.Schedule.Meeting.Neg`<br />`IPM.Schedule.Meeting.Pos`<br />`IPM.Schedule.Meeting.Tent`<br />`IPM.Schedule.Meeting.Canceled` |

<span data-ttu-id="496ae-297">Sie können benutzerdefinierte Nachrichtenklassen zum Erweitern einer Standardnachrichtenklasse erstellen, z. B. eine benutzerdefinierte Terminnachrichtenklasse wie `IPM.Appointment.Contoso`.</span><span class="sxs-lookup"><span data-stu-id="496ae-297">You can create custom message classes that extends a default message class, for example, a custom appointment message class `IPM.Appointment.Contoso`.</span></span>

##### <a name="type"></a><span data-ttu-id="496ae-298">Typ:</span><span class="sxs-lookup"><span data-stu-id="496ae-298">Type:</span></span>

*   <span data-ttu-id="496ae-299">String</span><span class="sxs-lookup"><span data-stu-id="496ae-299">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="496ae-300">Anforderungen</span><span class="sxs-lookup"><span data-stu-id="496ae-300">Requirements</span></span>

|<span data-ttu-id="496ae-301">Anforderung</span><span class="sxs-lookup"><span data-stu-id="496ae-301">Requirement</span></span>| <span data-ttu-id="496ae-302">Wert</span><span class="sxs-lookup"><span data-stu-id="496ae-302">Value</span></span>|
|---|---|
|[<span data-ttu-id="496ae-303">Mindestversion des Postfachanforderungssatzes</span><span class="sxs-lookup"><span data-stu-id="496ae-303">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="496ae-304">1.0</span><span class="sxs-lookup"><span data-stu-id="496ae-304">1.0</span></span>|
|[<span data-ttu-id="496ae-305">Mindestberechtigungsstufe</span><span class="sxs-lookup"><span data-stu-id="496ae-305">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="496ae-306">ReadItem</span><span class="sxs-lookup"><span data-stu-id="496ae-306">ReadItem</span></span>|
|[<span data-ttu-id="496ae-307">Zutreffender Outlook-Modus</span><span class="sxs-lookup"><span data-stu-id="496ae-307">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="496ae-308">Lesen</span><span class="sxs-lookup"><span data-stu-id="496ae-308">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="496ae-309">Beispiel</span><span class="sxs-lookup"><span data-stu-id="496ae-309">Example</span></span>

```
var itemClass = Office.context.mailbox.item.itemClass;
```

#### <a name="nullable-itemid-string"></a><span data-ttu-id="496ae-310">(nullable) itemId :String</span><span class="sxs-lookup"><span data-stu-id="496ae-310">(nullable) itemId :String</span></span>

<span data-ttu-id="496ae-p117">Ruft den Exchange-Webdienste-Elementbezeichner für das aktuelle Element ab. Nur Lesemodus.</span><span class="sxs-lookup"><span data-stu-id="496ae-p117">Gets the Exchange Web Services item identifier for the current item. Read mode only.</span></span>

> [!NOTE]
> <span data-ttu-id="496ae-313">Der Bezeichner, der von der `itemId`-Eigenschaft zurückgegeben wird, ist der gleiche Bezeichner wie der Exchange-Webdienste-Elementbezeichner.</span><span class="sxs-lookup"><span data-stu-id="496ae-313">The identifier returned by the `itemId` property is the same as the Exchange Web Services item identifier.</span></span> <span data-ttu-id="496ae-314">Die `itemId` -Eigenschaft ist nicht identisch mit der Outlook-Eintrags-ID oder die ID von der Outlook-REST-API verwendet.</span><span class="sxs-lookup"><span data-stu-id="496ae-314">The `itemId` property is not identical to the Outlook Entry ID or the ID used by the Outlook REST API.</span></span> <span data-ttu-id="496ae-315">REST-API-Aufrufe dieser Wert verwenden, sollten sie konvertiert werden, bevor [Office.context.mailbox.convertToRestId](office.context.mailbox.md#converttorestiditemid-restversion--string)verwenden.</span><span class="sxs-lookup"><span data-stu-id="496ae-315">Before making REST API calls using this value, it should be converted using [Office.context.mailbox.convertToRestId](office.context.mailbox.md#converttorestiditemid-restversion--string).</span></span> <span data-ttu-id="496ae-316">Weitere Informationen finden Sie unter [Verwenden der Outlook-REST-APIs aus einem Outlook-add-in](https://docs.microsoft.com/outlook/add-ins/use-rest-api#get-the-item-id).</span><span class="sxs-lookup"><span data-stu-id="496ae-316">For more details, see [Use the Outlook REST APIs from an Outlook add-in](https://docs.microsoft.com/outlook/add-ins/use-rest-api#get-the-item-id).</span></span>

<span data-ttu-id="496ae-p119">Die Eigenschaft `itemId` ist im Modus „Verfassen“ nicht verfügbar. Wenn ein Elementbezeichner erforderlich ist, kann die [`saveAsync`](#saveasyncoptions-callback)-Methode verwendet werden, um das Element zu speichern. Dadurch wird der Elementbezeichner im [`AsyncResult.value`](/javascript/api/office/office.asyncresult)-Parameter in der Rückruffunktion zurückgegeben.</span><span class="sxs-lookup"><span data-stu-id="496ae-p119">The `itemId` property is not available in compose mode. If an item identifier is required, the [`saveAsync`](#saveasyncoptions-callback) method can be used to save the item to the store, which will return the item identifier in the [`AsyncResult.value`](/javascript/api/office/office.asyncresult) parameter in the callback function.</span></span>

##### <a name="type"></a><span data-ttu-id="496ae-319">Typ:</span><span class="sxs-lookup"><span data-stu-id="496ae-319">Type:</span></span>

*   <span data-ttu-id="496ae-320">String</span><span class="sxs-lookup"><span data-stu-id="496ae-320">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="496ae-321">Anforderungen</span><span class="sxs-lookup"><span data-stu-id="496ae-321">Requirements</span></span>

|<span data-ttu-id="496ae-322">Anforderung</span><span class="sxs-lookup"><span data-stu-id="496ae-322">Requirement</span></span>| <span data-ttu-id="496ae-323">Wert</span><span class="sxs-lookup"><span data-stu-id="496ae-323">Value</span></span>|
|---|---|
|[<span data-ttu-id="496ae-324">Mindestversion des Postfachanforderungssatzes</span><span class="sxs-lookup"><span data-stu-id="496ae-324">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="496ae-325">1.0</span><span class="sxs-lookup"><span data-stu-id="496ae-325">1.0</span></span>|
|[<span data-ttu-id="496ae-326">Mindestberechtigungsstufe</span><span class="sxs-lookup"><span data-stu-id="496ae-326">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="496ae-327">ReadItem</span><span class="sxs-lookup"><span data-stu-id="496ae-327">ReadItem</span></span>|
|[<span data-ttu-id="496ae-328">Zutreffender Outlook-Modus</span><span class="sxs-lookup"><span data-stu-id="496ae-328">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="496ae-329">Lesen</span><span class="sxs-lookup"><span data-stu-id="496ae-329">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="496ae-330">Beispiel</span><span class="sxs-lookup"><span data-stu-id="496ae-330">Example</span></span>

<span data-ttu-id="496ae-p120">Mit dem folgende Code wird das Vorhandensein eines Elementbezeichners überprüft. Wenn die `itemId`-Eigenschaft `null` oder `undefined` zurückgibt, speichert sie das Element und ruft den Elementbezeichner aus dem asynchronen Ergebnis ab.</span><span class="sxs-lookup"><span data-stu-id="496ae-p120">The following code checks for the presence of an item identifier. If the `itemId` property returns `null` or `undefined`, it saves the item to the store and gets the item identifier from the asynchronous result.</span></span>

```
var itemId = Office.context.mailbox.item.itemId;
if (itemId === null || itemId == undefined) {
  Office.context.mailbox.item.saveAsync(function(result){
    itemId = result.value;
  });
}
```

####  <a name="itemtype-officemailboxenumsitemtypejavascriptapioutlook13officemailboxenumsitemtype"></a><span data-ttu-id="496ae-333">itemType :[Office.MailboxEnums.ItemType](/javascript/api/outlook_1_3/office.mailboxenums.itemtype)</span><span class="sxs-lookup"><span data-stu-id="496ae-333">itemType :[Office.MailboxEnums.ItemType](/javascript/api/outlook_1_3/office.mailboxenums.itemtype)</span></span>

<span data-ttu-id="496ae-334">Ruft den Typ des Elements ab, das eine Instanz darstellt.</span><span class="sxs-lookup"><span data-stu-id="496ae-334">Gets the type of item that an instance represents.</span></span>

<span data-ttu-id="496ae-335">Die `itemType`-Eigenschaft gibt einen der Werte der `ItemType`-Enumeration zurück, der angibt, ob es sich bei der `item`-Objektinstanz um eine Nachricht oder einen Termin handelt.</span><span class="sxs-lookup"><span data-stu-id="496ae-335">The `itemType` property returns one of the `ItemType` enumeration values, indicating whether the `item` object instance is a message or an appointment.</span></span>

##### <a name="type"></a><span data-ttu-id="496ae-336">Typ:</span><span class="sxs-lookup"><span data-stu-id="496ae-336">Type:</span></span>

*   [<span data-ttu-id="496ae-337">Office.MailboxEnums.ItemType</span><span class="sxs-lookup"><span data-stu-id="496ae-337">Office.MailboxEnums.ItemType</span></span>](/javascript/api/outlook_1_3/office.mailboxenums.itemtype)

##### <a name="requirements"></a><span data-ttu-id="496ae-338">Anforderungen</span><span class="sxs-lookup"><span data-stu-id="496ae-338">Requirements</span></span>

|<span data-ttu-id="496ae-339">Anforderung</span><span class="sxs-lookup"><span data-stu-id="496ae-339">Requirement</span></span>| <span data-ttu-id="496ae-340">Wert</span><span class="sxs-lookup"><span data-stu-id="496ae-340">Value</span></span>|
|---|---|
|[<span data-ttu-id="496ae-341">Mindestversion des Postfachanforderungssatzes</span><span class="sxs-lookup"><span data-stu-id="496ae-341">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="496ae-342">1.0</span><span class="sxs-lookup"><span data-stu-id="496ae-342">1.0</span></span>|
|[<span data-ttu-id="496ae-343">Mindestberechtigungsstufe</span><span class="sxs-lookup"><span data-stu-id="496ae-343">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="496ae-344">ReadItem</span><span class="sxs-lookup"><span data-stu-id="496ae-344">ReadItem</span></span>|
|[<span data-ttu-id="496ae-345">Zutreffender Outlook-Modus</span><span class="sxs-lookup"><span data-stu-id="496ae-345">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="496ae-346">Verfassen oder Lesen</span><span class="sxs-lookup"><span data-stu-id="496ae-346">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="496ae-347">Beispiel</span><span class="sxs-lookup"><span data-stu-id="496ae-347">Example</span></span>

```
if (Office.context.mailbox.item.itemType == Office.MailboxEnums.ItemType.Message)
  // do something
else
  // do something else
```

####  <a name="location-stringlocationjavascriptapioutlook13officelocation"></a><span data-ttu-id="496ae-348">location :String|[Location](/javascript/api/outlook_1_3/office.location)</span><span class="sxs-lookup"><span data-stu-id="496ae-348">location :String|[Location](/javascript/api/outlook_1_3/office.location)</span></span>

<span data-ttu-id="496ae-349">Ruft den Ort eines Termins ab bzw. legt ihn fest.</span><span class="sxs-lookup"><span data-stu-id="496ae-349">Gets or sets the location of an appointment.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="496ae-350">Lesemodus</span><span class="sxs-lookup"><span data-stu-id="496ae-350">Read mode</span></span>

<span data-ttu-id="496ae-351">Die `location`-Eigenschaft gibt eine Zeichenfolge zurück, die den Ort des Termins enthält.</span><span class="sxs-lookup"><span data-stu-id="496ae-351">The `location` property returns a string that contains the location of the appointment.</span></span>

##### <a name="compose-mode"></a><span data-ttu-id="496ae-352">Verfassenmodus</span><span class="sxs-lookup"><span data-stu-id="496ae-352">Compose mode</span></span>

<span data-ttu-id="496ae-353">Die `location`-Eigenschaft gibt ein `Location`-Objekt zurück, das Methoden zum Abrufen und Festlegen des Orts für einen Termin bereitstellt.</span><span class="sxs-lookup"><span data-stu-id="496ae-353">The `location` property returns a `Location` object that provides methods that are used to get and set the location of the appointment.</span></span>

##### <a name="type"></a><span data-ttu-id="496ae-354">Typ:</span><span class="sxs-lookup"><span data-stu-id="496ae-354">Type:</span></span>

*   <span data-ttu-id="496ae-355">String | [Location](/javascript/api/outlook_1_3/office.location)</span><span class="sxs-lookup"><span data-stu-id="496ae-355">String | [Location](/javascript/api/outlook_1_3/office.location)</span></span>

##### <a name="requirements"></a><span data-ttu-id="496ae-356">Anforderungen</span><span class="sxs-lookup"><span data-stu-id="496ae-356">Requirements</span></span>

|<span data-ttu-id="496ae-357">Anforderung</span><span class="sxs-lookup"><span data-stu-id="496ae-357">Requirement</span></span>| <span data-ttu-id="496ae-358">Wert</span><span class="sxs-lookup"><span data-stu-id="496ae-358">Value</span></span>|
|---|---|
|[<span data-ttu-id="496ae-359">Mindestversion des Postfachanforderungssatzes</span><span class="sxs-lookup"><span data-stu-id="496ae-359">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="496ae-360">1.0</span><span class="sxs-lookup"><span data-stu-id="496ae-360">1.0</span></span>|
|[<span data-ttu-id="496ae-361">Mindestberechtigungsstufe</span><span class="sxs-lookup"><span data-stu-id="496ae-361">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="496ae-362">ReadItem</span><span class="sxs-lookup"><span data-stu-id="496ae-362">ReadItem</span></span>|
|[<span data-ttu-id="496ae-363">Zutreffender Outlook-Modus</span><span class="sxs-lookup"><span data-stu-id="496ae-363">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="496ae-364">Verfassen oder Lesen</span><span class="sxs-lookup"><span data-stu-id="496ae-364">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="496ae-365">Beispiel</span><span class="sxs-lookup"><span data-stu-id="496ae-365">Example</span></span>

```
var userContext = { value : 1 };
Office.context.mailbox.item.location.getAsync( { context: userContext}, callback);

function callback(asyncResult) {
  var context = asyncResult.context;
  var location = asyncResult.value;
}
```

#### <a name="normalizedsubject-string"></a><span data-ttu-id="496ae-366">normalizedSubject :String</span><span class="sxs-lookup"><span data-stu-id="496ae-366">normalizedSubject :String</span></span>

<span data-ttu-id="496ae-p121">Ruft den Betreff für ein Element ab, wobei alle Präfixe entfernt werden (einschließlich `RE:` und `FWD:`). Nur Lesemodus.</span><span class="sxs-lookup"><span data-stu-id="496ae-p121">Gets the subject of an item, with all prefixes removed (including `RE:` and `FWD:`). Read mode only.</span></span>

<span data-ttu-id="496ae-p122">Die normalizedSubject-Eigenschaft ruft den Betreff des Elements mit allen Standardpräfixen (z. B. `RE:` und `FW:`) ab, die von E-Mail-Programmen hinzugefügt werden. Wenn der Betreff des Elements mit vollständigen Präfixen abgerufen werden soll, verwenden Sie die [`subject`](#subject-stringsubjectjavascriptapioutlook13officesubject)-Eigenschaft.</span><span class="sxs-lookup"><span data-stu-id="496ae-p122">The normalizedSubject property gets the subject of the item, with any standard prefixes (such as `RE:` and `FW:`) that are added by email programs. To get the subject of the item with the prefixes intact, use the [`subject`](#subject-stringsubjectjavascriptapioutlook13officesubject) property.</span></span>

##### <a name="type"></a><span data-ttu-id="496ae-371">Typ:</span><span class="sxs-lookup"><span data-stu-id="496ae-371">Type:</span></span>

*   <span data-ttu-id="496ae-372">String</span><span class="sxs-lookup"><span data-stu-id="496ae-372">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="496ae-373">Anforderungen</span><span class="sxs-lookup"><span data-stu-id="496ae-373">Requirements</span></span>

|<span data-ttu-id="496ae-374">Anforderung</span><span class="sxs-lookup"><span data-stu-id="496ae-374">Requirement</span></span>| <span data-ttu-id="496ae-375">Wert</span><span class="sxs-lookup"><span data-stu-id="496ae-375">Value</span></span>|
|---|---|
|[<span data-ttu-id="496ae-376">Mindestversion des Postfachanforderungssatzes</span><span class="sxs-lookup"><span data-stu-id="496ae-376">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="496ae-377">1.0</span><span class="sxs-lookup"><span data-stu-id="496ae-377">1.0</span></span>|
|[<span data-ttu-id="496ae-378">Mindestberechtigungsstufe</span><span class="sxs-lookup"><span data-stu-id="496ae-378">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="496ae-379">ReadItem</span><span class="sxs-lookup"><span data-stu-id="496ae-379">ReadItem</span></span>|
|[<span data-ttu-id="496ae-380">Zutreffender Outlook-Modus</span><span class="sxs-lookup"><span data-stu-id="496ae-380">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="496ae-381">Lesen</span><span class="sxs-lookup"><span data-stu-id="496ae-381">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="496ae-382">Beispiel</span><span class="sxs-lookup"><span data-stu-id="496ae-382">Example</span></span>

```
var normalizedSubject = Office.context.mailbox.item.normalizedSubject;
```

####  <a name="notificationmessages-notificationmessagesjavascriptapioutlook13officenotificationmessages"></a><span data-ttu-id="496ae-383">notificationMessages :[NotificationMessages](/javascript/api/outlook_1_3/office.notificationmessages)</span><span class="sxs-lookup"><span data-stu-id="496ae-383">notificationMessages :[NotificationMessages](/javascript/api/outlook_1_3/office.notificationmessages)</span></span>

<span data-ttu-id="496ae-384">Ruft die Benachrichtigungen für ein Element ab.</span><span class="sxs-lookup"><span data-stu-id="496ae-384">Gets the notification messages for an item.</span></span>

##### <a name="type"></a><span data-ttu-id="496ae-385">Typ:</span><span class="sxs-lookup"><span data-stu-id="496ae-385">Type:</span></span>

*   [<span data-ttu-id="496ae-386">NotificationMessages</span><span class="sxs-lookup"><span data-stu-id="496ae-386">NotificationMessages</span></span>](/javascript/api/outlook_1_3/office.notificationmessages)

##### <a name="requirements"></a><span data-ttu-id="496ae-387">Anforderungen</span><span class="sxs-lookup"><span data-stu-id="496ae-387">Requirements</span></span>

|<span data-ttu-id="496ae-388">Anforderung</span><span class="sxs-lookup"><span data-stu-id="496ae-388">Requirement</span></span>| <span data-ttu-id="496ae-389">Wert</span><span class="sxs-lookup"><span data-stu-id="496ae-389">Value</span></span>|
|---|---|
|[<span data-ttu-id="496ae-390">Mindestversion des Postfachanforderungssatzes</span><span class="sxs-lookup"><span data-stu-id="496ae-390">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="496ae-391">1.3</span><span class="sxs-lookup"><span data-stu-id="496ae-391">1.3</span></span>|
|[<span data-ttu-id="496ae-392">Mindestberechtigungsstufe</span><span class="sxs-lookup"><span data-stu-id="496ae-392">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="496ae-393">ReadItem</span><span class="sxs-lookup"><span data-stu-id="496ae-393">ReadItem</span></span>|
|[<span data-ttu-id="496ae-394">Zutreffender Outlook-Modus</span><span class="sxs-lookup"><span data-stu-id="496ae-394">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="496ae-395">Verfassen oder Lesen</span><span class="sxs-lookup"><span data-stu-id="496ae-395">Compose or read</span></span>|

####  <a name="optionalattendees-arrayemailaddressdetailsjavascriptapioutlook13officeemailaddressdetailsrecipientsjavascriptapioutlook13officerecipients"></a><span data-ttu-id="496ae-396">optionalAttendees :Array.<[EmailAddressDetails](/javascript/api/outlook_1_3/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook_1_3/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="496ae-396">optionalAttendees :Array.<[EmailAddressDetails](/javascript/api/outlook_1_3/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook_1_3/office.recipients)</span></span>

<span data-ttu-id="496ae-397">Ermöglicht den Zugriff auf die optionalen Teilnehmer eines Ereignisses.</span><span class="sxs-lookup"><span data-stu-id="496ae-397">Provides access to the optional attendees of an event.</span></span> <span data-ttu-id="496ae-398">Der Typ des Objekts und die Zugriffsebene, hängt von den Modus des aktuellen Elements ab.</span><span class="sxs-lookup"><span data-stu-id="496ae-398">The type of object and level of access depends on the mode of the current item.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="496ae-399">Lesemodus</span><span class="sxs-lookup"><span data-stu-id="496ae-399">Read mode</span></span>

<span data-ttu-id="496ae-400">Die `optionalAttendees`-Eigenschaft gibt ein Array mit einem `EmailAddressDetails`-Objekt für jeden optionalen Teilnehmer an der Besprechung zurück.</span><span class="sxs-lookup"><span data-stu-id="496ae-400">The `optionalAttendees` property returns an array that contains an `EmailAddressDetails` object for each optional attendee to the meeting.</span></span>

##### <a name="compose-mode"></a><span data-ttu-id="496ae-401">Verfassenmodus</span><span class="sxs-lookup"><span data-stu-id="496ae-401">Compose mode</span></span>

<span data-ttu-id="496ae-402">Die `optionalAttendees` -Eigenschaft gibt eine `Recipients` -Objekt, das Methoden zum Abrufen oder Aktualisieren der optionalen Teilnehmer für eine Besprechung bereitstellt.</span><span class="sxs-lookup"><span data-stu-id="496ae-402">The `optionalAttendees` property returns a `Recipients` object that provides methods to get or update the optional attendees for a meeting.</span></span>

##### <a name="type"></a><span data-ttu-id="496ae-403">Typ:</span><span class="sxs-lookup"><span data-stu-id="496ae-403">Type:</span></span>

*   <span data-ttu-id="496ae-404">Array.<[EmailAddressDetails](/javascript/api/outlook_1_3/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook_1_3/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="496ae-404">Array.<[EmailAddressDetails](/javascript/api/outlook_1_3/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook_1_3/office.recipients)</span></span>

##### <a name="requirements"></a><span data-ttu-id="496ae-405">Anforderungen</span><span class="sxs-lookup"><span data-stu-id="496ae-405">Requirements</span></span>

|<span data-ttu-id="496ae-406">Anforderung</span><span class="sxs-lookup"><span data-stu-id="496ae-406">Requirement</span></span>| <span data-ttu-id="496ae-407">Wert</span><span class="sxs-lookup"><span data-stu-id="496ae-407">Value</span></span>|
|---|---|
|[<span data-ttu-id="496ae-408">Mindestversion des Postfachanforderungssatzes</span><span class="sxs-lookup"><span data-stu-id="496ae-408">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="496ae-409">1.0</span><span class="sxs-lookup"><span data-stu-id="496ae-409">1.0</span></span>|
|[<span data-ttu-id="496ae-410">Mindestberechtigungsstufe</span><span class="sxs-lookup"><span data-stu-id="496ae-410">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="496ae-411">ReadItem</span><span class="sxs-lookup"><span data-stu-id="496ae-411">ReadItem</span></span>|
|[<span data-ttu-id="496ae-412">Zutreffender Outlook-Modus</span><span class="sxs-lookup"><span data-stu-id="496ae-412">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="496ae-413">Verfassen oder Lesen</span><span class="sxs-lookup"><span data-stu-id="496ae-413">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="496ae-414">Beispiel</span><span class="sxs-lookup"><span data-stu-id="496ae-414">Example</span></span>

```
Office.context.mailbox.item.optionalAttendees.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.optionalAttendees.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.optionalAttendees.getAsync(callback);

function callback(asyncResult) {
  var arrayOfOptionalAttendeesRecipients = asyncResult.value;
}
```

#### <a name="organizer-emailaddressdetailsjavascriptapioutlook13officeemailaddressdetails"></a><span data-ttu-id="496ae-415">organizer :[EmailAddressDetails](/javascript/api/outlook_1_3/office.emailaddressdetails)</span><span class="sxs-lookup"><span data-stu-id="496ae-415">organizer :[EmailAddressDetails](/javascript/api/outlook_1_3/office.emailaddressdetails)</span></span>

<span data-ttu-id="496ae-p124">Ruft für eine angegebene Besprechung die E-Mail-Adresse des Organisators der Besprechung ab. Nur Lesemodus.</span><span class="sxs-lookup"><span data-stu-id="496ae-p124">Gets the email address of the meeting organizer for a specified meeting. Read mode only.</span></span>

##### <a name="type"></a><span data-ttu-id="496ae-418">Typ:</span><span class="sxs-lookup"><span data-stu-id="496ae-418">Type:</span></span>

*   [<span data-ttu-id="496ae-419">EmailAddressDetails</span><span class="sxs-lookup"><span data-stu-id="496ae-419">EmailAddressDetails</span></span>](/javascript/api/outlook_1_3/office.emailaddressdetails)

##### <a name="requirements"></a><span data-ttu-id="496ae-420">Anforderungen</span><span class="sxs-lookup"><span data-stu-id="496ae-420">Requirements</span></span>

|<span data-ttu-id="496ae-421">Anforderung</span><span class="sxs-lookup"><span data-stu-id="496ae-421">Requirement</span></span>| <span data-ttu-id="496ae-422">Wert</span><span class="sxs-lookup"><span data-stu-id="496ae-422">Value</span></span>|
|---|---|
|[<span data-ttu-id="496ae-423">Mindestversion des Postfachanforderungssatzes</span><span class="sxs-lookup"><span data-stu-id="496ae-423">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="496ae-424">1.0</span><span class="sxs-lookup"><span data-stu-id="496ae-424">1.0</span></span>|
|[<span data-ttu-id="496ae-425">Mindestberechtigungsstufe</span><span class="sxs-lookup"><span data-stu-id="496ae-425">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="496ae-426">ReadItem</span><span class="sxs-lookup"><span data-stu-id="496ae-426">ReadItem</span></span>|
|[<span data-ttu-id="496ae-427">Zutreffender Outlook-Modus</span><span class="sxs-lookup"><span data-stu-id="496ae-427">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="496ae-428">Lesen</span><span class="sxs-lookup"><span data-stu-id="496ae-428">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="496ae-429">Beispiel</span><span class="sxs-lookup"><span data-stu-id="496ae-429">Example</span></span>

```
var organizerName = Office.context.mailbox.item.organizer.displayName;
var organizerAddress = Office.context.mailbox.item.organizer.emailAddress;
```

####  <a name="requiredattendees-arrayemailaddressdetailsjavascriptapioutlook13officeemailaddressdetailsrecipientsjavascriptapioutlook13officerecipients"></a><span data-ttu-id="496ae-430">requiredAttendees :Array.<[EmailAddressDetails](/javascript/api/outlook_1_3/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook_1_3/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="496ae-430">requiredAttendees :Array.<[EmailAddressDetails](/javascript/api/outlook_1_3/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook_1_3/office.recipients)</span></span>

<span data-ttu-id="496ae-431">Ermöglicht den Zugriff auf die erforderlichen Teilnehmer eines Ereignisses.</span><span class="sxs-lookup"><span data-stu-id="496ae-431">Provides access to the required attendees of an event.</span></span> <span data-ttu-id="496ae-432">Der Typ des Objekts und die Zugriffsebene, hängt von den Modus des aktuellen Elements ab.</span><span class="sxs-lookup"><span data-stu-id="496ae-432">The type of object and level of access depends on the mode of the current item.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="496ae-433">Lesemodus</span><span class="sxs-lookup"><span data-stu-id="496ae-433">Read mode</span></span>

<span data-ttu-id="496ae-434">Die `requiredAttendees`-Eigenschaft gibt ein Array mit einem `EmailAddressDetails`-Objekt für jeden erforderlichen Teilnehmer an der Besprechung zurück.</span><span class="sxs-lookup"><span data-stu-id="496ae-434">The `requiredAttendees` property returns an array that contains an `EmailAddressDetails` object for each required attendee to the meeting.</span></span>

##### <a name="compose-mode"></a><span data-ttu-id="496ae-435">Verfassenmodus</span><span class="sxs-lookup"><span data-stu-id="496ae-435">Compose mode</span></span>

<span data-ttu-id="496ae-436">Die `requiredAttendees` -Eigenschaft gibt eine `Recipients` -Objekt, das Methoden zum Abrufen oder Aktualisieren der erforderlichen Teilnehmer für eine Besprechung bereitstellt.</span><span class="sxs-lookup"><span data-stu-id="496ae-436">The `requiredAttendees` property returns a `Recipients` object that provides methods to get or update the required attendees for a meeting.</span></span>

##### <a name="type"></a><span data-ttu-id="496ae-437">Typ:</span><span class="sxs-lookup"><span data-stu-id="496ae-437">Type:</span></span>

*   <span data-ttu-id="496ae-438">Array.<[EmailAddressDetails](/javascript/api/outlook_1_3/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook_1_3/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="496ae-438">Array.<[EmailAddressDetails](/javascript/api/outlook_1_3/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook_1_3/office.recipients)</span></span>

##### <a name="requirements"></a><span data-ttu-id="496ae-439">Anforderungen</span><span class="sxs-lookup"><span data-stu-id="496ae-439">Requirements</span></span>

|<span data-ttu-id="496ae-440">Anforderung</span><span class="sxs-lookup"><span data-stu-id="496ae-440">Requirement</span></span>| <span data-ttu-id="496ae-441">Wert</span><span class="sxs-lookup"><span data-stu-id="496ae-441">Value</span></span>|
|---|---|
|[<span data-ttu-id="496ae-442">Mindestversion des Postfachanforderungssatzes</span><span class="sxs-lookup"><span data-stu-id="496ae-442">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="496ae-443">1.0</span><span class="sxs-lookup"><span data-stu-id="496ae-443">1.0</span></span>|
|[<span data-ttu-id="496ae-444">Mindestberechtigungsstufe</span><span class="sxs-lookup"><span data-stu-id="496ae-444">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="496ae-445">ReadItem</span><span class="sxs-lookup"><span data-stu-id="496ae-445">ReadItem</span></span>|
|[<span data-ttu-id="496ae-446">Zutreffender Outlook-Modus</span><span class="sxs-lookup"><span data-stu-id="496ae-446">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="496ae-447">Verfassen oder Lesen</span><span class="sxs-lookup"><span data-stu-id="496ae-447">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="496ae-448">Beispiel</span><span class="sxs-lookup"><span data-stu-id="496ae-448">Example</span></span>

```
Office.context.mailbox.item.requiredAttendees.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.requiredAttendees.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.requiredAttendees.getAsync(callback);

function callback(asyncResult) {
  var arrayOfRequiredAttendeesRecipients = asyncResult.value;
}
```

#### <a name="sender-emailaddressdetailsjavascriptapioutlook13officeemailaddressdetails"></a><span data-ttu-id="496ae-449">sender :[EmailAddressDetails](/javascript/api/outlook_1_3/office.emailaddressdetails)</span><span class="sxs-lookup"><span data-stu-id="496ae-449">sender :[EmailAddressDetails](/javascript/api/outlook_1_3/office.emailaddressdetails)</span></span>

<span data-ttu-id="496ae-p126">Ruft die E-Mail-Adresse des Absenders einer E-Mail-Nachricht ab. Nur Lesemodus.</span><span class="sxs-lookup"><span data-stu-id="496ae-p126">Gets the email address of the sender of an email message. Read mode only.</span></span>

<span data-ttu-id="496ae-p127">Die [`from`](#from-emailaddressdetailsjavascriptapioutlook13officeemailaddressdetails)- und `sender`-Eigenschaften stellen dieselbe Person dar, außer die Nachricht wurde von einem Delegaten gesendet. In diesem Fall stellt die `from`-Eigenschaft die Stellvertretung dar, und die sender-Eigenschaft stellt den Delegaten dar.</span><span class="sxs-lookup"><span data-stu-id="496ae-p127">The [`from`](#from-emailaddressdetailsjavascriptapioutlook13officeemailaddressdetails) and `sender` properties represent the same person unless the message is sent by a delegate. In that case, the `from` property represents the delegator, and the sender property represents the delegate.</span></span>

> [!NOTE]
> <span data-ttu-id="496ae-454">Die `recipientType` -Eigenschaft des der `EmailAddressDetails` -Objekts der `sender` -Eigenschaft ist `undefined`.</span><span class="sxs-lookup"><span data-stu-id="496ae-454">The `recipientType` property of the `EmailAddressDetails` object in the `sender` property is `undefined`.</span></span>

##### <a name="type"></a><span data-ttu-id="496ae-455">Typ:</span><span class="sxs-lookup"><span data-stu-id="496ae-455">Type:</span></span>

*   [<span data-ttu-id="496ae-456">EmailAddressDetails</span><span class="sxs-lookup"><span data-stu-id="496ae-456">EmailAddressDetails</span></span>](/javascript/api/outlook_1_3/office.emailaddressdetails)

##### <a name="requirements"></a><span data-ttu-id="496ae-457">Anforderungen</span><span class="sxs-lookup"><span data-stu-id="496ae-457">Requirements</span></span>

|<span data-ttu-id="496ae-458">Anforderung</span><span class="sxs-lookup"><span data-stu-id="496ae-458">Requirement</span></span>| <span data-ttu-id="496ae-459">Wert</span><span class="sxs-lookup"><span data-stu-id="496ae-459">Value</span></span>|
|---|---|
|[<span data-ttu-id="496ae-460">Mindestversion des Postfachanforderungssatzes</span><span class="sxs-lookup"><span data-stu-id="496ae-460">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="496ae-461">1.0</span><span class="sxs-lookup"><span data-stu-id="496ae-461">1.0</span></span>|
|[<span data-ttu-id="496ae-462">Mindestberechtigungsstufe</span><span class="sxs-lookup"><span data-stu-id="496ae-462">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="496ae-463">ReadItem</span><span class="sxs-lookup"><span data-stu-id="496ae-463">ReadItem</span></span>|
|[<span data-ttu-id="496ae-464">Zutreffender Outlook-Modus</span><span class="sxs-lookup"><span data-stu-id="496ae-464">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="496ae-465">Lesen</span><span class="sxs-lookup"><span data-stu-id="496ae-465">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="496ae-466">Beispiel</span><span class="sxs-lookup"><span data-stu-id="496ae-466">Example</span></span>

```
var senderName = Office.context.mailbox.item.sender.displayName;
var senderAddress = Office.context.mailbox.item.sender.emailAddress;
```

####  <a name="start-datetimejavascriptapioutlook13officetime"></a><span data-ttu-id="496ae-467">start :Date|[Time](/javascript/api/outlook_1_3/office.time)</span><span class="sxs-lookup"><span data-stu-id="496ae-467">start :Date|[Time](/javascript/api/outlook_1_3/office.time)</span></span>

<span data-ttu-id="496ae-468">Ruft Datum und Zeit für den Beginn des Termins ab oder legt Datum und Uhrzeit fest.</span><span class="sxs-lookup"><span data-stu-id="496ae-468">Gets or sets the date and time that the appointment is to begin.</span></span>

<span data-ttu-id="496ae-p128">Die `start`-Eigenschaft wird als Datums- und Uhrzeitwert in UTC ausgedrückt. Sie können die [`convertToLocalClientTime`](office.context.mailbox.md#converttolocalclienttimetimevalue--localclienttimejavascriptapioutlook13officelocalclienttime)-Methode verwenden, um den Wert in den lokalen Uhrzeit- und Datumswert des Clients umzuwandeln.</span><span class="sxs-lookup"><span data-stu-id="496ae-p128">The `start` property is expressed as a Coordinated Universal Time (UTC) date and time value. You can use the [`convertToLocalClientTime`](office.context.mailbox.md#converttolocalclienttimetimevalue--localclienttimejavascriptapioutlook13officelocalclienttime) method to convert the value to the client’s local date and time.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="496ae-471">Lesemodus</span><span class="sxs-lookup"><span data-stu-id="496ae-471">Read mode</span></span>

<span data-ttu-id="496ae-472">Die `start`-Eigenschaft gibt ein `Date`-Objekt zurück.</span><span class="sxs-lookup"><span data-stu-id="496ae-472">The `start` property returns a `Date` object.</span></span>

##### <a name="compose-mode"></a><span data-ttu-id="496ae-473">Verfassenmodus</span><span class="sxs-lookup"><span data-stu-id="496ae-473">Compose mode</span></span>

<span data-ttu-id="496ae-474">Die `start`-Eigenschaft gibt ein `Time`-Objekt zurück.</span><span class="sxs-lookup"><span data-stu-id="496ae-474">The `start` property returns a `Time` object.</span></span>

<span data-ttu-id="496ae-475">Wenn Sie die [`Time.setAsync`](/javascript/api/outlook_1_3/office.time#setasync-datetime--options--callback-)-Methode verwenden, um die Startzeit im Verfassenmodus festzulegen, sollten Sie die [`convertToUtcClientTime`](office.context.mailbox.md#converttoutcclienttimeinput--date)-Methode verwenden, um die Ortszeit auf dem Client für den Server in UTC umzuwandeln.</span><span class="sxs-lookup"><span data-stu-id="496ae-475">When you use the [`Time.setAsync`](/javascript/api/outlook_1_3/office.time#setasync-datetime--options--callback-) method to set the start time, you should use the [`convertToUtcClientTime`](office.context.mailbox.md#converttoutcclienttimeinput--date) method to convert the local time on the client to UTC for the server.</span></span>

##### <a name="type"></a><span data-ttu-id="496ae-476">Typ:</span><span class="sxs-lookup"><span data-stu-id="496ae-476">Type:</span></span>

*   <span data-ttu-id="496ae-477">Date | [Time](/javascript/api/outlook_1_3/office.time)</span><span class="sxs-lookup"><span data-stu-id="496ae-477">Date | [Time](/javascript/api/outlook_1_3/office.time)</span></span>

##### <a name="requirements"></a><span data-ttu-id="496ae-478">Anforderungen</span><span class="sxs-lookup"><span data-stu-id="496ae-478">Requirements</span></span>

|<span data-ttu-id="496ae-479">Anforderung</span><span class="sxs-lookup"><span data-stu-id="496ae-479">Requirement</span></span>| <span data-ttu-id="496ae-480">Wert</span><span class="sxs-lookup"><span data-stu-id="496ae-480">Value</span></span>|
|---|---|
|[<span data-ttu-id="496ae-481">Mindestversion des Postfachanforderungssatzes</span><span class="sxs-lookup"><span data-stu-id="496ae-481">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="496ae-482">1.0</span><span class="sxs-lookup"><span data-stu-id="496ae-482">1.0</span></span>|
|[<span data-ttu-id="496ae-483">Mindestberechtigungsstufe</span><span class="sxs-lookup"><span data-stu-id="496ae-483">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="496ae-484">ReadItem</span><span class="sxs-lookup"><span data-stu-id="496ae-484">ReadItem</span></span>|
|[<span data-ttu-id="496ae-485">Zutreffender Outlook-Modus</span><span class="sxs-lookup"><span data-stu-id="496ae-485">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="496ae-486">Verfassen oder Lesen</span><span class="sxs-lookup"><span data-stu-id="496ae-486">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="496ae-487">Beispiel</span><span class="sxs-lookup"><span data-stu-id="496ae-487">Example</span></span>

<span data-ttu-id="496ae-488">Im folgenden Beispiel wird die Startzeit eines Termins im Verfassenmodus mithilfe der [`setAsync`](/javascript/api/outlook_1_3/office.time#setasync-datetime--options--callback-)-Methode des `Time`-Objekts festgelegt.</span><span class="sxs-lookup"><span data-stu-id="496ae-488">The following example sets the start time of an appointment in compose mode by using the [`setAsync`](/javascript/api/outlook_1_3/office.time#setasync-datetime--options--callback-) method of the `Time` object.</span></span>

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

####  <a name="subject-stringsubjectjavascriptapioutlook13officesubject"></a><span data-ttu-id="496ae-489">subject :String|[Subject](/javascript/api/outlook_1_3/office.subject)</span><span class="sxs-lookup"><span data-stu-id="496ae-489">subject :String|[Subject](/javascript/api/outlook_1_3/office.subject)</span></span>

<span data-ttu-id="496ae-490">Ruft die Beschreibung ab, die im Betrefffeld eines Elements angezeigt wird, oder legt sie fest.</span><span class="sxs-lookup"><span data-stu-id="496ae-490">Gets or sets the description that appears in the subject field of an item.</span></span>

<span data-ttu-id="496ae-491">Die `subject`-Eigenschaft ruft den gesamten Betreff des Elements ab oder legt ihn fest – so, wie er vom E-Mail-Server gesendet wird.</span><span class="sxs-lookup"><span data-stu-id="496ae-491">The `subject` property gets or sets the entire subject of the item, as sent by the email server.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="496ae-492">Lesemodus</span><span class="sxs-lookup"><span data-stu-id="496ae-492">Read mode</span></span>

<span data-ttu-id="496ae-p129">Die `subject`-Eigenschaft gibt eine Zeichenfolge zurück. Verwenden Sie die [`normalizedSubject`](#normalizedsubject-string)-Eigenschaft, um den Betreff ohne führende Präfixe wie `RE:` und `FW:` abzurufen.</span><span class="sxs-lookup"><span data-stu-id="496ae-p129">The `subject` property returns a string. Use the [`normalizedSubject`](#normalizedsubject-string) property to get the subject minus any leading prefixes such as `RE:` and `FW:`.</span></span>

```
var subject = Office.context.mailbox.item.subject;
```

##### <a name="compose-mode"></a><span data-ttu-id="496ae-495">Verfassenmodus</span><span class="sxs-lookup"><span data-stu-id="496ae-495">Compose mode</span></span>

<span data-ttu-id="496ae-496">Die `subject`-Eigenschaft gibt ein `Subject`-Objekt zurück, das Methoden zum Abrufen und Festlegen des Betreffs bereitstellt.</span><span class="sxs-lookup"><span data-stu-id="496ae-496">The `subject` property returns a `Subject` object that provides methods to get and set the subject.</span></span>

```
Office.context.mailbox.item.subject.getAsync(callback);

function callback(asyncResult) {
  var subject = asyncResult.value;
}
```

##### <a name="type"></a><span data-ttu-id="496ae-497">Typ:</span><span class="sxs-lookup"><span data-stu-id="496ae-497">Type:</span></span>

*   <span data-ttu-id="496ae-498">String | [Subject](/javascript/api/outlook_1_3/office.subject)</span><span class="sxs-lookup"><span data-stu-id="496ae-498">String | [Subject](/javascript/api/outlook_1_3/office.subject)</span></span>

##### <a name="requirements"></a><span data-ttu-id="496ae-499">Anforderungen</span><span class="sxs-lookup"><span data-stu-id="496ae-499">Requirements</span></span>

|<span data-ttu-id="496ae-500">Anforderung</span><span class="sxs-lookup"><span data-stu-id="496ae-500">Requirement</span></span>| <span data-ttu-id="496ae-501">Wert</span><span class="sxs-lookup"><span data-stu-id="496ae-501">Value</span></span>|
|---|---|
|[<span data-ttu-id="496ae-502">Mindestversion des Postfachanforderungssatzes</span><span class="sxs-lookup"><span data-stu-id="496ae-502">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="496ae-503">1.0</span><span class="sxs-lookup"><span data-stu-id="496ae-503">1.0</span></span>|
|[<span data-ttu-id="496ae-504">Mindestberechtigungsstufe</span><span class="sxs-lookup"><span data-stu-id="496ae-504">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="496ae-505">ReadItem</span><span class="sxs-lookup"><span data-stu-id="496ae-505">ReadItem</span></span>|
|[<span data-ttu-id="496ae-506">Zutreffender Outlook-Modus</span><span class="sxs-lookup"><span data-stu-id="496ae-506">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="496ae-507">Verfassen oder Lesen</span><span class="sxs-lookup"><span data-stu-id="496ae-507">Compose or read</span></span>|

####  <a name="to-arrayemailaddressdetailsjavascriptapioutlook13officeemailaddressdetailsrecipientsjavascriptapioutlook13officerecipients"></a><span data-ttu-id="496ae-508">an: Array. <[EmailAddressDetails](/javascript/api/outlook_1_3/office.emailaddressdetails)>|[Empfänger](/javascript/api/outlook_1_3/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="496ae-508">to :Array.<[EmailAddressDetails](/javascript/api/outlook_1_3/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook_1_3/office.recipients)</span></span>

<span data-ttu-id="496ae-509">Ermöglicht den Zugriff auf die Empfänger in der Zeile **an** einer Nachricht.</span><span class="sxs-lookup"><span data-stu-id="496ae-509">Provides access to the recipients on the **To** line of a message.</span></span> <span data-ttu-id="496ae-510">Der Typ des Objekts und die Zugriffsebene, hängt von den Modus des aktuellen Elements ab.</span><span class="sxs-lookup"><span data-stu-id="496ae-510">The type of object and level of access depends on the mode of the current item.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="496ae-511">Lesemodus</span><span class="sxs-lookup"><span data-stu-id="496ae-511">Read mode</span></span>

<span data-ttu-id="496ae-p131">Die `to`-Eigenschaft gibt ein Array mit einem `EmailAddressDetails`-Objekt für jeden Empfänger in der**An**-Zeile der Nachricht zurück. Die Auflistung ist auf höchstens 100 Elemente beschränkt.</span><span class="sxs-lookup"><span data-stu-id="496ae-p131">The `to` property returns an array that contains an `EmailAddressDetails` object for each recipient listed on the **To** line of the message. The collection is limited to a maximum of 100 members.</span></span>

##### <a name="compose-mode"></a><span data-ttu-id="496ae-514">Verfassenmodus</span><span class="sxs-lookup"><span data-stu-id="496ae-514">Compose mode</span></span>

<span data-ttu-id="496ae-515">Die `to` -Eigenschaft gibt eine `Recipients` -Objekt, das Methoden zum Abrufen oder aktualisieren die Empfänger in der Zeile **an** der Nachricht bereitstellt.</span><span class="sxs-lookup"><span data-stu-id="496ae-515">The `to` property returns a `Recipients` object that provides methods to get or update the recipients on the **To** line of the message.</span></span>

##### <a name="type"></a><span data-ttu-id="496ae-516">Typ:</span><span class="sxs-lookup"><span data-stu-id="496ae-516">Type:</span></span>

*   <span data-ttu-id="496ae-517">Array.<[EmailAddressDetails](/javascript/api/outlook_1_3/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook_1_3/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="496ae-517">Array.<[EmailAddressDetails](/javascript/api/outlook_1_3/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook_1_3/office.recipients)</span></span>

##### <a name="requirements"></a><span data-ttu-id="496ae-518">Anforderungen</span><span class="sxs-lookup"><span data-stu-id="496ae-518">Requirements</span></span>

|<span data-ttu-id="496ae-519">Anforderung</span><span class="sxs-lookup"><span data-stu-id="496ae-519">Requirement</span></span>| <span data-ttu-id="496ae-520">Wert</span><span class="sxs-lookup"><span data-stu-id="496ae-520">Value</span></span>|
|---|---|
|[<span data-ttu-id="496ae-521">Mindestversion des Postfachanforderungssatzes</span><span class="sxs-lookup"><span data-stu-id="496ae-521">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="496ae-522">1.0</span><span class="sxs-lookup"><span data-stu-id="496ae-522">1.0</span></span>|
|[<span data-ttu-id="496ae-523">Mindestberechtigungsstufe</span><span class="sxs-lookup"><span data-stu-id="496ae-523">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="496ae-524">ReadItem</span><span class="sxs-lookup"><span data-stu-id="496ae-524">ReadItem</span></span>|
|[<span data-ttu-id="496ae-525">Zutreffender Outlook-Modus</span><span class="sxs-lookup"><span data-stu-id="496ae-525">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="496ae-526">Verfassen oder Lesen</span><span class="sxs-lookup"><span data-stu-id="496ae-526">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="496ae-527">Beispiel</span><span class="sxs-lookup"><span data-stu-id="496ae-527">Example</span></span>

```
Office.context.mailbox.item.to.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.to.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.to.getAsync(callback);

function callback(asyncResult) {
  var arrayOfToRecipients = asyncResult.value;
}
```

### <a name="methods"></a><span data-ttu-id="496ae-528">Methoden</span><span class="sxs-lookup"><span data-stu-id="496ae-528">Methods</span></span>

####  <a name="addfileattachmentasyncuri-attachmentname-options-callback"></a><span data-ttu-id="496ae-529">addFileAttachmentAsync(uri, attachmentName, [options], [callback])</span><span class="sxs-lookup"><span data-stu-id="496ae-529">addFileAttachmentAsync(uri, attachmentName, [options], [callback])</span></span>

<span data-ttu-id="496ae-530">Fügt eine Datei zu einer Nachricht oder einem Termin als Anlage hinzu.</span><span class="sxs-lookup"><span data-stu-id="496ae-530">Adds a file to a message or appointment as an attachment.</span></span>

<span data-ttu-id="496ae-531">Die `addFileAttachmentAsync`-Methode lädt die Datei am angegebenen URI hoch und fügt sie an das Element im Verfassenformular an.</span><span class="sxs-lookup"><span data-stu-id="496ae-531">The `addFileAttachmentAsync` method uploads the file at the specified URI and attaches it to the item in the compose form.</span></span>

<span data-ttu-id="496ae-532">Anschließend können Sie den Bezeichner mit der [`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback)-Methode in der gleichen Sitzung zum Entfernen der Anlage verwenden.</span><span class="sxs-lookup"><span data-stu-id="496ae-532">You can subsequently use the identifier with the [`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback) method to remove the attachment in the same session.</span></span>

##### <a name="parameters"></a><span data-ttu-id="496ae-533">Parameter:</span><span class="sxs-lookup"><span data-stu-id="496ae-533">Parameters:</span></span>

|<span data-ttu-id="496ae-534">Name</span><span class="sxs-lookup"><span data-stu-id="496ae-534">Name</span></span>| <span data-ttu-id="496ae-535">Typ</span><span class="sxs-lookup"><span data-stu-id="496ae-535">Type</span></span>| <span data-ttu-id="496ae-536">Attribute</span><span class="sxs-lookup"><span data-stu-id="496ae-536">Attributes</span></span>| <span data-ttu-id="496ae-537">Beschreibung</span><span class="sxs-lookup"><span data-stu-id="496ae-537">Description</span></span>|
|---|---|---|---|
|`uri`| <span data-ttu-id="496ae-538">Zeichenfolge</span><span class="sxs-lookup"><span data-stu-id="496ae-538">String</span></span>||<span data-ttu-id="496ae-p132">Der URI, der den Speicherort der an die Nachricht oder den Termin anzuhängenden Datei angibt. Die maximale Länge ist 2048 Zeichen.</span><span class="sxs-lookup"><span data-stu-id="496ae-p132">The URI that provides the location of the file to attach to the message or appointment. The maximum length is 2048 characters.</span></span>|
|`attachmentName`| <span data-ttu-id="496ae-541">Zeichenfolge</span><span class="sxs-lookup"><span data-stu-id="496ae-541">String</span></span>||<span data-ttu-id="496ae-p133">Der Name der Anlage, der beim Hochladen der Anlage angezeigt wird. Die maximale Länge ist 255 Zeichen.</span><span class="sxs-lookup"><span data-stu-id="496ae-p133">The name of the attachment that is shown while the attachment is uploading. The maximum length is 255 characters.</span></span>|
|`options`| <span data-ttu-id="496ae-544">Object</span><span class="sxs-lookup"><span data-stu-id="496ae-544">Object</span></span>| <span data-ttu-id="496ae-545">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="496ae-545">&lt;optional&gt;</span></span>|<span data-ttu-id="496ae-546">Ein Objektliteral, das mindestens eine der folgenden Eigenschaften enthält.</span><span class="sxs-lookup"><span data-stu-id="496ae-546">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`| <span data-ttu-id="496ae-547">Object</span><span class="sxs-lookup"><span data-stu-id="496ae-547">Object</span></span>| <span data-ttu-id="496ae-548">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="496ae-548">&lt;optional&gt;</span></span>|<span data-ttu-id="496ae-549">Entwickler können ein Objekt bereitstellen, auf das sie in der Callbackmethode zugreifen möchten.</span><span class="sxs-lookup"><span data-stu-id="496ae-549">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`| <span data-ttu-id="496ae-550">Funktion</span><span class="sxs-lookup"><span data-stu-id="496ae-550">function</span></span>| <span data-ttu-id="496ae-551">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="496ae-551">&lt;optional&gt;</span></span>|<span data-ttu-id="496ae-552">Wenn die Methode abgeschlossen wird, wird die Funktion , die im `callback`-Parameter übergeben wird, mit einem einzelnen Parameter (`asyncResult`) aufgerufen, der ein [`AsyncResult`](/javascript/api/office/office.asyncresult)-Objekt darstellt.</span><span class="sxs-lookup"><span data-stu-id="496ae-552">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span> <br/><span data-ttu-id="496ae-553">Bei Erfolg wird der Anlagenbezeichner in der `asyncResult.value`-Eigenschaft bereitgestellt.</span><span class="sxs-lookup"><span data-stu-id="496ae-553">On success, the attachment identifier will be provided in the `asyncResult.value` property.</span></span><br/><span data-ttu-id="496ae-554">Wenn beim Hochladen der Anlage ein Fehler auftritt, enthält das `asyncResult`-Objekt ein `Error`-Objekt mit einer Beschreibung des Fehlers.</span><span class="sxs-lookup"><span data-stu-id="496ae-554">If uploading the attachment fails, the `asyncResult` object will contain an `Error` object that provides a description of the error.</span></span>|

##### <a name="errors"></a><span data-ttu-id="496ae-555">Fehler</span><span class="sxs-lookup"><span data-stu-id="496ae-555">Errors</span></span>

| <span data-ttu-id="496ae-556">Fehlercode</span><span class="sxs-lookup"><span data-stu-id="496ae-556">Error code</span></span> | <span data-ttu-id="496ae-557">Beschreibung</span><span class="sxs-lookup"><span data-stu-id="496ae-557">Description</span></span> |
|------------|-------------|
| `AttachmentSizeExceeded` | <span data-ttu-id="496ae-558">Die Anlage ist zu groß.</span><span class="sxs-lookup"><span data-stu-id="496ae-558">The attachment is larger than allowed.</span></span> |
| `FileTypeNotSupported` | <span data-ttu-id="496ae-559">Die Anlage enthält eine Erweiterung, die nicht zulässig ist.</span><span class="sxs-lookup"><span data-stu-id="496ae-559">The attachment has an extension that is not allowed.</span></span> |
| `NumberOfAttachmentsExceeded` | <span data-ttu-id="496ae-560">Die Nachricht oder der Termin enthält zu viele Anlagen.</span><span class="sxs-lookup"><span data-stu-id="496ae-560">The message or appointment has too many attachments.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="496ae-561">Anforderungen</span><span class="sxs-lookup"><span data-stu-id="496ae-561">Requirements</span></span>

|<span data-ttu-id="496ae-562">Anforderung</span><span class="sxs-lookup"><span data-stu-id="496ae-562">Requirement</span></span>| <span data-ttu-id="496ae-563">Wert</span><span class="sxs-lookup"><span data-stu-id="496ae-563">Value</span></span>|
|---|---|
|[<span data-ttu-id="496ae-564">Mindestversion des Postfachanforderungssatzes</span><span class="sxs-lookup"><span data-stu-id="496ae-564">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="496ae-565">1.1</span><span class="sxs-lookup"><span data-stu-id="496ae-565">1.1</span></span>|
|[<span data-ttu-id="496ae-566">Mindestberechtigungsstufe</span><span class="sxs-lookup"><span data-stu-id="496ae-566">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="496ae-567">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="496ae-567">ReadWriteItem</span></span>|
|[<span data-ttu-id="496ae-568">Zutreffender Outlook-Modus</span><span class="sxs-lookup"><span data-stu-id="496ae-568">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="496ae-569">Verfassen</span><span class="sxs-lookup"><span data-stu-id="496ae-569">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="496ae-570">Beispiel</span><span class="sxs-lookup"><span data-stu-id="496ae-570">Example</span></span>

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

####  <a name="additemattachmentasyncitemid-attachmentname-options-callback"></a><span data-ttu-id="496ae-571">addItemAttachmentAsync(itemId, attachmentName, [options], [callback])</span><span class="sxs-lookup"><span data-stu-id="496ae-571">addItemAttachmentAsync(itemId, attachmentName, [options], [callback])</span></span>

<span data-ttu-id="496ae-572">Fügt der Nachricht oder dem Termin ein Exchange-Objekt, wie z. B. eine Nachricht, als Anhang hinzu.</span><span class="sxs-lookup"><span data-stu-id="496ae-572">Adds an Exchange item, such as a message, as an attachment to the message or appointment.</span></span>

<span data-ttu-id="496ae-p134">Die `addItemAttachmentAsync`-Methode fügt das Element mit dem angegebenen Exchange-Bezeichner an das Element im Formular zum Verfassen an. Wenn Sie eine Rückrufmethode angeben, wird die Methode mit einem `asyncResult`-Parameter aufgerufen, der entweder den Anlagenbezeichner oder einen Code enthält, der etwaige Fehler angibt, die beim Anfügen des Objekts aufgetreten sind. Sie können ggf. den `options`-Parameter verwenden, um Statusinformationen an die Rückrufmethode zu übergeben.</span><span class="sxs-lookup"><span data-stu-id="496ae-p134">The `addItemAttachmentAsync` method attaches the item with the specified Exchange identifier to the item in the compose form. If you specify a callback method, the method is called with one parameter, `asyncResult`, which contains either the attachment identifier or a code that indicates any error that occurred while attaching the item. You can use the `options` parameter to pass state information to the callback method, if needed.</span></span>

<span data-ttu-id="496ae-576">Anschließend können Sie den Bezeichner mit der [`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback)-Methode in der gleichen Sitzung zum Entfernen der Anlage verwenden.</span><span class="sxs-lookup"><span data-stu-id="496ae-576">You can subsequently use the identifier with the [`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback) method to remove the attachment in the same session.</span></span>

<span data-ttu-id="496ae-577">Wenn Ihre Office-Add-Ins in Outlook Web App ausgeführt wird die `addItemAttachmentAsync` -Methode kann Anfügen von Elementen in der Elemente, die Sie bearbeiten; Allerdings Dies wird nicht unterstützt und wird nicht empfohlen.</span><span class="sxs-lookup"><span data-stu-id="496ae-577">If your Office Add-in is running in Outlook Web App, the `addItemAttachmentAsync` method can attach items to items other than the item that you are editing; however, this is not supported and is not recommended.</span></span>

##### <a name="parameters"></a><span data-ttu-id="496ae-578">Parameter:</span><span class="sxs-lookup"><span data-stu-id="496ae-578">Parameters:</span></span>

|<span data-ttu-id="496ae-579">Name</span><span class="sxs-lookup"><span data-stu-id="496ae-579">Name</span></span>| <span data-ttu-id="496ae-580">Typ</span><span class="sxs-lookup"><span data-stu-id="496ae-580">Type</span></span>| <span data-ttu-id="496ae-581">Attribute</span><span class="sxs-lookup"><span data-stu-id="496ae-581">Attributes</span></span>| <span data-ttu-id="496ae-582">Beschreibung</span><span class="sxs-lookup"><span data-stu-id="496ae-582">Description</span></span>|
|---|---|---|---|
|`itemId`| <span data-ttu-id="496ae-583">Zeichenfolge</span><span class="sxs-lookup"><span data-stu-id="496ae-583">String</span></span>||<span data-ttu-id="496ae-p135">Der Exchange-Bezeichner des Objekts, das angehängt werden soll. Die maximale Länge beträgt 100 Zeichen.</span><span class="sxs-lookup"><span data-stu-id="496ae-p135">The Exchange identifier of the item to attach. The maximum length is 100 characters.</span></span>|
|`attachmentName`| <span data-ttu-id="496ae-586">Zeichenfolge</span><span class="sxs-lookup"><span data-stu-id="496ae-586">String</span></span>||<span data-ttu-id="496ae-p136">Der Betreff des Elements, das angehängt werden soll. Die maximale Länge ist 255 Zeichen.</span><span class="sxs-lookup"><span data-stu-id="496ae-p136">The sujbect of the item to be attached. The maximum length is 255 characters.</span></span>|
|`options`| <span data-ttu-id="496ae-589">Object</span><span class="sxs-lookup"><span data-stu-id="496ae-589">Object</span></span>| <span data-ttu-id="496ae-590">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="496ae-590">&lt;optional&gt;</span></span>|<span data-ttu-id="496ae-591">Ein Objektliteral, das mindestens eine der folgenden Eigenschaften enthält.</span><span class="sxs-lookup"><span data-stu-id="496ae-591">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`| <span data-ttu-id="496ae-592">Object</span><span class="sxs-lookup"><span data-stu-id="496ae-592">Object</span></span>| <span data-ttu-id="496ae-593">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="496ae-593">&lt;optional&gt;</span></span>|<span data-ttu-id="496ae-594">Entwickler können ein Objekt bereitstellen, auf das sie in der Callbackmethode zugreifen möchten.</span><span class="sxs-lookup"><span data-stu-id="496ae-594">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`| <span data-ttu-id="496ae-595">Funktion</span><span class="sxs-lookup"><span data-stu-id="496ae-595">function</span></span>| <span data-ttu-id="496ae-596">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="496ae-596">&lt;optional&gt;</span></span>|<span data-ttu-id="496ae-597">Wenn die Methode abgeschlossen wird, wird die Funktion , die im `callback`-Parameter übergeben wird, mit einem einzelnen Parameter (`asyncResult`) aufgerufen, der ein [`AsyncResult`](/javascript/api/office/office.asyncresult)-Objekt darstellt.</span><span class="sxs-lookup"><span data-stu-id="496ae-597">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span> <br/><span data-ttu-id="496ae-598">Bei Erfolg wird der Anlagenbezeichner in der `asyncResult.value`-Eigenschaft bereitgestellt.</span><span class="sxs-lookup"><span data-stu-id="496ae-598">On success, the attachment identifier will be provided in the `asyncResult.value` property.</span></span><br/><span data-ttu-id="496ae-599">Wenn beim Hinzufügen der Anlage ein Fehler auftritt, enthält das `asyncResult`-Objekt ein `Error`-Objekt mit einer Beschreibung des Fehlers.</span><span class="sxs-lookup"><span data-stu-id="496ae-599">If adding the attachment fails, the `asyncResult` object will contain an `Error` object that provides a description of the error.</span></span>|

##### <a name="errors"></a><span data-ttu-id="496ae-600">Fehler</span><span class="sxs-lookup"><span data-stu-id="496ae-600">Errors</span></span>

| <span data-ttu-id="496ae-601">Fehlercode</span><span class="sxs-lookup"><span data-stu-id="496ae-601">Error code</span></span> | <span data-ttu-id="496ae-602">Beschreibung</span><span class="sxs-lookup"><span data-stu-id="496ae-602">Description</span></span> |
|------------|-------------|
| `NumberOfAttachmentsExceeded` | <span data-ttu-id="496ae-603">Die Nachricht oder der Termin enthält zu viele Anlagen.</span><span class="sxs-lookup"><span data-stu-id="496ae-603">The message or appointment has too many attachments.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="496ae-604">Anforderungen</span><span class="sxs-lookup"><span data-stu-id="496ae-604">Requirements</span></span>

|<span data-ttu-id="496ae-605">Anforderung</span><span class="sxs-lookup"><span data-stu-id="496ae-605">Requirement</span></span>| <span data-ttu-id="496ae-606">Wert</span><span class="sxs-lookup"><span data-stu-id="496ae-606">Value</span></span>|
|---|---|
|[<span data-ttu-id="496ae-607">Mindestversion des Postfachanforderungssatzes</span><span class="sxs-lookup"><span data-stu-id="496ae-607">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="496ae-608">1.1</span><span class="sxs-lookup"><span data-stu-id="496ae-608">1.1</span></span>|
|[<span data-ttu-id="496ae-609">Mindestberechtigungsstufe</span><span class="sxs-lookup"><span data-stu-id="496ae-609">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="496ae-610">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="496ae-610">ReadWriteItem</span></span>|
|[<span data-ttu-id="496ae-611">Zutreffender Outlook-Modus</span><span class="sxs-lookup"><span data-stu-id="496ae-611">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="496ae-612">Verfassen</span><span class="sxs-lookup"><span data-stu-id="496ae-612">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="496ae-613">Beispiel</span><span class="sxs-lookup"><span data-stu-id="496ae-613">Example</span></span>

<span data-ttu-id="496ae-614">Im folgenden Beispiel wird ein vorhandenes Outlook-Element als Anlage mit dem Namen `My Attachment` hinzugefügt.</span><span class="sxs-lookup"><span data-stu-id="496ae-614">The following example adds an existing Outlook item as an attachment with the name `My Attachment`.</span></span>

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

####  <a name="close"></a><span data-ttu-id="496ae-615">close()</span><span class="sxs-lookup"><span data-stu-id="496ae-615">close()</span></span>

<span data-ttu-id="496ae-616">Schließt das aktuelle Element, das gerade verfasst wird.</span><span class="sxs-lookup"><span data-stu-id="496ae-616">Closes the current item that is being composed.</span></span>

<span data-ttu-id="496ae-p137">Das Verhalten der `close`-Methode hängt vom aktuellen Status des verfassten Elements ab. Wenn das Element über nicht gespeicherte Änderungen verfügt, fordert der Client den Benutzer auf, die Schließenaktion zu speichern, zu verwerfen oder abzubrechen.</span><span class="sxs-lookup"><span data-stu-id="496ae-p137">The behavior of the `close` method depends on the current state of the item being composed. If the item has unsaved changes, the client prompts the user to save, discard, or cancel the close action.</span></span>

> [!NOTE]
> <span data-ttu-id="496ae-619">In Outlook im Web, wenn das Element einen Termin handelt, und es wurde bereits zuvor gespeichert wurde mit `saveAsync`, der Benutzer wird aufgefordert, speichern, löschen oder Abbrechen, auch wenn keine Änderungen aufgetreten sind, da das Element zuletzt gespeichert.</span><span class="sxs-lookup"><span data-stu-id="496ae-619">In Outlook on the web, if the item is an appointment and it has previously been saved using `saveAsync`, the user is prompted to save, discard, or cancel even if no changes have occurred since the item was last saved.</span></span>

<span data-ttu-id="496ae-620">Wenn die Nachricht im Outlook-Desktopclient eine Inlineantwort ist, hat die `close`-Methode keine Auswirkung.</span><span class="sxs-lookup"><span data-stu-id="496ae-620">In the Outlook desktop client, if the message is an inline reply, the `close` method has no effect.</span></span>

##### <a name="requirements"></a><span data-ttu-id="496ae-621">Anforderungen</span><span class="sxs-lookup"><span data-stu-id="496ae-621">Requirements</span></span>

|<span data-ttu-id="496ae-622">Anforderung</span><span class="sxs-lookup"><span data-stu-id="496ae-622">Requirement</span></span>| <span data-ttu-id="496ae-623">Wert</span><span class="sxs-lookup"><span data-stu-id="496ae-623">Value</span></span>|
|---|---|
|[<span data-ttu-id="496ae-624">Mindestversion des Postfachanforderungssatzes</span><span class="sxs-lookup"><span data-stu-id="496ae-624">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="496ae-625">1.3</span><span class="sxs-lookup"><span data-stu-id="496ae-625">1.3</span></span>|
|[<span data-ttu-id="496ae-626">Mindestberechtigungsstufe</span><span class="sxs-lookup"><span data-stu-id="496ae-626">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="496ae-627">Eingeschränkt</span><span class="sxs-lookup"><span data-stu-id="496ae-627">Restricted</span></span>|
|[<span data-ttu-id="496ae-628">Zutreffender Outlook-Modus</span><span class="sxs-lookup"><span data-stu-id="496ae-628">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="496ae-629">Verfassen</span><span class="sxs-lookup"><span data-stu-id="496ae-629">Compose</span></span>|

#### <a name="displayreplyallformformdata"></a><span data-ttu-id="496ae-630">displayReplyAllForm(formData)</span><span class="sxs-lookup"><span data-stu-id="496ae-630">displayReplyAllForm(formData)</span></span>

<span data-ttu-id="496ae-631">Zeigt ein Antwortformular an, das den Absender und alle Empfänger der ausgewählten Nachricht oder des Organisators und alle Teilnehmer des ausgewählten Termins enthält.</span><span class="sxs-lookup"><span data-stu-id="496ae-631">Displays a reply form that includes the sender and all recipients of the selected message or the organizer and all attendees of the selected appointment.</span></span>

> [!NOTE]
> <span data-ttu-id="496ae-632">Diese Methode wird in Outlook für iOS oder Outlook für Android nicht unterstützt.</span><span class="sxs-lookup"><span data-stu-id="496ae-632">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="496ae-633">In Outlook Web App wird das Antwortformular als Popupformular in der Dreispaltenansicht und als Popupformular in der Zwei- oder Einspaltenansicht angezeigt.</span><span class="sxs-lookup"><span data-stu-id="496ae-633">In Outlook Web App, the reply form is displayed as a pop-out form in the 3-column view and a pop-up form in the 2- or 1-column view.</span></span>

<span data-ttu-id="496ae-634">Wenn einer der Zeichenfolgenparameter seinen Grenzwert überschreitet, löst `displayReplyAllForm` eine Ausnahme aus.</span><span class="sxs-lookup"><span data-stu-id="496ae-634">If any of the string parameters exceed their limits, `displayReplyAllForm` throws an exception.</span></span>

<span data-ttu-id="496ae-p138">Wenn Anlagen im `formData.attachments`-Parameter angegeben werden, versuchen Outlook und Outlook Web App alle Anlagen herunterzuladen und sie an das Antwortformular anzufügen. Wenn Anlagen nicht hinzugefügt werden können, wird in der Formularbenutzeroberfläche ein Fehler ausgegeben. Wenn dies nicht möglich ist, wird keine Fehlermeldung ausgegeben.</span><span class="sxs-lookup"><span data-stu-id="496ae-p138">When attachments are specified in the `formData.attachments` parameter, Outlook and Outlook Web App attempt to download all attachments and attach them to the reply form. If any attachments fail to be added, an error is shown in the form UI. If this isn't possible, then no error message is thrown.</span></span>

##### <a name="parameters"></a><span data-ttu-id="496ae-638">Parameter:</span><span class="sxs-lookup"><span data-stu-id="496ae-638">Parameters:</span></span>

|<span data-ttu-id="496ae-639">Name</span><span class="sxs-lookup"><span data-stu-id="496ae-639">Name</span></span>| <span data-ttu-id="496ae-640">Typ</span><span class="sxs-lookup"><span data-stu-id="496ae-640">Type</span></span>| <span data-ttu-id="496ae-641">Beschreibung</span><span class="sxs-lookup"><span data-stu-id="496ae-641">Description</span></span>|
|---|---|---|
|`formData`| <span data-ttu-id="496ae-642">Zeichenfolge &#124; Objekt</span><span class="sxs-lookup"><span data-stu-id="496ae-642">String &#124; Object</span></span>| |<span data-ttu-id="496ae-p139">Eine Zeichenfolge, die Text- und HTML-Code enthält, die den Hauptteil des Antwortformulars darstellen. Die Zeichenfolge ist auf 32 KB beschränkt.</span><span class="sxs-lookup"><span data-stu-id="496ae-p139">A string that contains text and HTML and that represents the body of the reply form. The string is limited to 32 KB.</span></span><br/><span data-ttu-id="496ae-645">**ODER**</span><span class="sxs-lookup"><span data-stu-id="496ae-645">**OR**</span></span><br/><span data-ttu-id="496ae-p140">Ein Objekt, das Text- oder Anlagendaten und eine Rückruffunktion enthält. Das Objekt ist wie folgt definiert:</span><span class="sxs-lookup"><span data-stu-id="496ae-p140">An object that contains body or attachment data and a callback function. The object is defined as follows.</span></span> |
| `formData.htmlBody` | <span data-ttu-id="496ae-648">Zeichenfolge</span><span class="sxs-lookup"><span data-stu-id="496ae-648">String</span></span> | <span data-ttu-id="496ae-649">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="496ae-649">&lt;optional&gt;</span></span> | <span data-ttu-id="496ae-p141">Eine Zeichenfolge, die Text- und HTML-Code enthält, die den Hauptteil des Antwortformulars darstellen. Die Zeichenfolge ist auf 32 KB beschränkt.</span><span class="sxs-lookup"><span data-stu-id="496ae-p141">A string that contains text and HTML and that represents the body of the reply form. The string is limited to 32 KB.</span></span>
| `formData.attachments` | <span data-ttu-id="496ae-652">Array.&lt;Object&gt;</span><span class="sxs-lookup"><span data-stu-id="496ae-652">Array.&lt;Object&gt;</span></span> | <span data-ttu-id="496ae-653">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="496ae-653">&lt;optional&gt;</span></span> | <span data-ttu-id="496ae-654">Ein Array mit JSON-Objekten, die Datei- oder Elementanlagen sind.</span><span class="sxs-lookup"><span data-stu-id="496ae-654">An array of JSON objects that are either file or item attachments.</span></span> |
| `formData.attachments.type` | <span data-ttu-id="496ae-655">Zeichenfolge</span><span class="sxs-lookup"><span data-stu-id="496ae-655">String</span></span> | | <span data-ttu-id="496ae-p142">Gibt den Typ der Anlage an. Muss `file` für eine Dateianlage oder `item` für eine Elementanlage sein.</span><span class="sxs-lookup"><span data-stu-id="496ae-p142">Indicates the type of attachment. Must be `file` for a file attachment or `item` for an item attachment.</span></span> |
| `formData.attachments.name` | <span data-ttu-id="496ae-658">Zeichenfolge</span><span class="sxs-lookup"><span data-stu-id="496ae-658">String</span></span> | | <span data-ttu-id="496ae-659">Eine Zeichenfolge, die den Namen der Anlage mit bis zu 255 Zeichen enthält.</span><span class="sxs-lookup"><span data-stu-id="496ae-659">A string that contains the name of the attachment, up to 255 characters in length.</span></span>|
| `formData.attachments.url` | <span data-ttu-id="496ae-660">Zeichenfolge</span><span class="sxs-lookup"><span data-stu-id="496ae-660">String</span></span> | | <span data-ttu-id="496ae-p143">Wird nur verwendet, wenn `type` auf `file` gesetzt ist. Der URI des Speicherorts für die Datei.</span><span class="sxs-lookup"><span data-stu-id="496ae-p143">Only used if `type` is set to `file`. The URI of the location for the file.</span></span> |
| `formData.attachments.itemId` | <span data-ttu-id="496ae-663">Zeichenfolge</span><span class="sxs-lookup"><span data-stu-id="496ae-663">String</span></span> | | <span data-ttu-id="496ae-p144">Wird nur verwendet, wenn `type` auf `item` gesetzt ist. Die EWS-Element-ID der Anlage. Dies ist eine Zeichenfolge bis zu 100 Zeichen.</span><span class="sxs-lookup"><span data-stu-id="496ae-p144">Only used if `type` is set to `item`. The EWS item id of the attachment. This is a string up to 100 characters.</span></span> |
| `callback` | <span data-ttu-id="496ae-667">Funktion</span><span class="sxs-lookup"><span data-stu-id="496ae-667">function</span></span> | <span data-ttu-id="496ae-668">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="496ae-668">&lt;optional&gt;</span></span> | <span data-ttu-id="496ae-669">Bei Abschluss der Methode wird die im Parameter `callback` übergebene Funktion mit einem einzigen Parameter aufgerufen: `asyncResult`. Dieser Parameter ist ein Objekt des Typs [AsyncResult](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="496ae-669">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [AsyncResult](/javascript/api/office/office.asyncresult) object.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="496ae-670">Anforderungen</span><span class="sxs-lookup"><span data-stu-id="496ae-670">Requirements</span></span>

|<span data-ttu-id="496ae-671">Anforderung</span><span class="sxs-lookup"><span data-stu-id="496ae-671">Requirement</span></span>| <span data-ttu-id="496ae-672">Wert</span><span class="sxs-lookup"><span data-stu-id="496ae-672">Value</span></span>|
|---|---|
|[<span data-ttu-id="496ae-673">Mindestversion des Postfachanforderungssatzes</span><span class="sxs-lookup"><span data-stu-id="496ae-673">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="496ae-674">1.0</span><span class="sxs-lookup"><span data-stu-id="496ae-674">1.0</span></span>|
|[<span data-ttu-id="496ae-675">Mindestberechtigungsstufe</span><span class="sxs-lookup"><span data-stu-id="496ae-675">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="496ae-676">ReadItem</span><span class="sxs-lookup"><span data-stu-id="496ae-676">ReadItem</span></span>|
|[<span data-ttu-id="496ae-677">Zutreffender Outlook-Modus</span><span class="sxs-lookup"><span data-stu-id="496ae-677">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="496ae-678">Lesen</span><span class="sxs-lookup"><span data-stu-id="496ae-678">Read</span></span>|

##### <a name="examples"></a><span data-ttu-id="496ae-679">Beispiele</span><span class="sxs-lookup"><span data-stu-id="496ae-679">Examples</span></span>

<span data-ttu-id="496ae-680">Im folgenden Code wird eine Zeichenfolge an die `displayReplyAllForm`-Funktion übergeben.</span><span class="sxs-lookup"><span data-stu-id="496ae-680">The following code passes a string to the `displayReplyAllForm` function.</span></span>

```
Office.context.mailbox.item.displayReplyAllForm('hello there');
Office.context.mailbox.item.displayReplyAllForm('<b>hello there</b>');
```

<span data-ttu-id="496ae-681">Antworten mit einem leeren Textkörper.</span><span class="sxs-lookup"><span data-stu-id="496ae-681">Reply with an empty body.</span></span>

```
Office.context.mailbox.item.displayReplyAllForm({});
```

<span data-ttu-id="496ae-682">Antworten nur einem Textkörper.</span><span class="sxs-lookup"><span data-stu-id="496ae-682">Reply with just a body.</span></span>

```
Office.context.mailbox.item.displayReplyAllForm(
{
  'htmlBody' : 'hi'
});
```

<span data-ttu-id="496ae-683">Antworten mit einem Textkörper und einer Dateianlage.</span><span class="sxs-lookup"><span data-stu-id="496ae-683">Reply with a body and a file attachment.</span></span>

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

<span data-ttu-id="496ae-684">Antworten mit einem Textkörper und einer Elementanlage.</span><span class="sxs-lookup"><span data-stu-id="496ae-684">Reply with a body and an item attachment.</span></span>

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

<span data-ttu-id="496ae-685">Antworten Sie mit einem Textkörper, einer Dateianlage, einer Elementanlage und einem Callback.</span><span class="sxs-lookup"><span data-stu-id="496ae-685">Reply with a body, file attachment, item attachment, and a callback.</span></span>

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

#### <a name="displayreplyformformdata"></a><span data-ttu-id="496ae-686">displayReplyForm(formData)</span><span class="sxs-lookup"><span data-stu-id="496ae-686">displayReplyForm(formData)</span></span>

<span data-ttu-id="496ae-687">Zeigt ein Antwortformular an, das nur den Absender der ausgewählten Nachricht oder den Organisator des ausgewählten Termins enthält.</span><span class="sxs-lookup"><span data-stu-id="496ae-687">Displays a reply form that includes only the sender of the selected message or the organizer of the selected appointment.</span></span>

> [!NOTE]
> <span data-ttu-id="496ae-688">Diese Methode wird in Outlook für iOS oder Outlook für Android nicht unterstützt.</span><span class="sxs-lookup"><span data-stu-id="496ae-688">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="496ae-689">In Outlook Web App wird das Antwortformular als Popupformular in der Dreispaltenansicht und als Popupformular in der Zwei- oder Einspaltenansicht angezeigt.</span><span class="sxs-lookup"><span data-stu-id="496ae-689">In Outlook Web App, the reply form is displayed as a pop-out form in the 3-column view and a pop-up form in the 2- or 1-column view.</span></span>

<span data-ttu-id="496ae-690">Wenn einer der Zeichenfolgenparameter seinen Grenzwert überschreitet, löst `displayReplyForm` eine Ausnahme aus.</span><span class="sxs-lookup"><span data-stu-id="496ae-690">If any of the string parameters exceed their limits, `displayReplyForm` throws an exception.</span></span>

<span data-ttu-id="496ae-p145">Wenn Anlagen im `formData.attachments`-Parameter angegeben werden, versuchen Outlook und Outlook Web App alle Anlagen herunterzuladen und sie an das Antwortformular anzufügen. Wenn Anlagen nicht hinzugefügt werden können, wird in der Formularbenutzeroberfläche ein Fehler ausgegeben. Wenn dies nicht möglich ist, wird keine Fehlermeldung ausgegeben.</span><span class="sxs-lookup"><span data-stu-id="496ae-p145">When attachments are specified in the `formData.attachments` parameter, Outlook and Outlook Web App attempt to download all attachments and attach them to the reply form. If any attachments fail to be added, an error is shown in the form UI. If this isn't possible, then no error message is thrown.</span></span>

##### <a name="parameters"></a><span data-ttu-id="496ae-694">Parameter:</span><span class="sxs-lookup"><span data-stu-id="496ae-694">Parameters:</span></span>

|<span data-ttu-id="496ae-695">Name</span><span class="sxs-lookup"><span data-stu-id="496ae-695">Name</span></span>| <span data-ttu-id="496ae-696">Typ</span><span class="sxs-lookup"><span data-stu-id="496ae-696">Type</span></span>| <span data-ttu-id="496ae-697">Beschreibung</span><span class="sxs-lookup"><span data-stu-id="496ae-697">Description</span></span>|
|---|---|---|
|`formData`| <span data-ttu-id="496ae-698">Zeichenfolge &#124; Objekt</span><span class="sxs-lookup"><span data-stu-id="496ae-698">String &#124; Object</span></span>| | <span data-ttu-id="496ae-p146">Eine Zeichenfolge, die Text- und HTML-Code enthält, die den Hauptteil des Antwortformulars darstellen. Die Zeichenfolge ist auf 32 KB beschränkt.</span><span class="sxs-lookup"><span data-stu-id="496ae-p146">A string that contains text and HTML and that represents the body of the reply form. The string is limited to 32 KB.</span></span><br/><span data-ttu-id="496ae-701">**ODER**</span><span class="sxs-lookup"><span data-stu-id="496ae-701">**OR**</span></span><br/><span data-ttu-id="496ae-p147">Ein Objekt, das Text- oder Anlagendaten und eine Rückruffunktion enthält. Das Objekt ist wie folgt definiert:</span><span class="sxs-lookup"><span data-stu-id="496ae-p147">An object that contains body or attachment data and a callback function. The object is defined as follows.</span></span> |
| `formData.htmlBody` | <span data-ttu-id="496ae-704">Zeichenfolge</span><span class="sxs-lookup"><span data-stu-id="496ae-704">String</span></span> | <span data-ttu-id="496ae-705">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="496ae-705">&lt;optional&gt;</span></span> | <span data-ttu-id="496ae-p148">Eine Zeichenfolge, die Text- und HTML-Code enthält, die den Hauptteil des Antwortformulars darstellen. Die Zeichenfolge ist auf 32 KB beschränkt.</span><span class="sxs-lookup"><span data-stu-id="496ae-p148">A string that contains text and HTML and that represents the body of the reply form. The string is limited to 32 KB.</span></span>
| `formData.attachments` | <span data-ttu-id="496ae-708">Array.&lt;Object&gt;</span><span class="sxs-lookup"><span data-stu-id="496ae-708">Array.&lt;Object&gt;</span></span> | <span data-ttu-id="496ae-709">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="496ae-709">&lt;optional&gt;</span></span> | <span data-ttu-id="496ae-710">Ein Array mit JSON-Objekten, die Datei- oder Elementanlagen sind.</span><span class="sxs-lookup"><span data-stu-id="496ae-710">An array of JSON objects that are either file or item attachments.</span></span> |
| `formData.attachments.type` | <span data-ttu-id="496ae-711">Zeichenfolge</span><span class="sxs-lookup"><span data-stu-id="496ae-711">String</span></span> | | <span data-ttu-id="496ae-p149">Gibt den Typ der Anlage an. Muss `file` für eine Dateianlage oder `item` für eine Elementanlage sein.</span><span class="sxs-lookup"><span data-stu-id="496ae-p149">Indicates the type of attachment. Must be `file` for a file attachment or `item` for an item attachment.</span></span> |
| `formData.attachments.name` | <span data-ttu-id="496ae-714">Zeichenfolge</span><span class="sxs-lookup"><span data-stu-id="496ae-714">String</span></span> | | <span data-ttu-id="496ae-715">Eine Zeichenfolge, die den Namen der Anlage mit bis zu 255 Zeichen enthält.</span><span class="sxs-lookup"><span data-stu-id="496ae-715">A string that contains the name of the attachment, up to 255 characters in length.</span></span>|
| `formData.attachments.url` | <span data-ttu-id="496ae-716">Zeichenfolge</span><span class="sxs-lookup"><span data-stu-id="496ae-716">String</span></span> | | <span data-ttu-id="496ae-p150">Wird nur verwendet, wenn `type` auf `file` gesetzt ist. Der URI des Speicherorts für die Datei.</span><span class="sxs-lookup"><span data-stu-id="496ae-p150">Only used if `type` is set to `file`. The URI of the location for the file.</span></span> |
| `formData.attachments.itemId` | <span data-ttu-id="496ae-719">Zeichenfolge</span><span class="sxs-lookup"><span data-stu-id="496ae-719">String</span></span> | | <span data-ttu-id="496ae-p151">Wird nur verwendet, wenn `type` auf `item` gesetzt ist. Die EWS-Element-ID der Anlage. Dies ist eine Zeichenfolge bis zu 100 Zeichen.</span><span class="sxs-lookup"><span data-stu-id="496ae-p151">Only used if `type` is set to `item`. The EWS item id of the attachment. This is a string up to 100 characters.</span></span> |
| `callback` | <span data-ttu-id="496ae-723">Funktion</span><span class="sxs-lookup"><span data-stu-id="496ae-723">function</span></span> | <span data-ttu-id="496ae-724">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="496ae-724">&lt;optional&gt;</span></span> | <span data-ttu-id="496ae-725">Bei Abschluss der Methode wird die im Parameter `callback` übergebene Funktion mit einem einzigen Parameter aufgerufen: `asyncResult`. Dieser Parameter ist ein Objekt des Typs [AsyncResult](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="496ae-725">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [AsyncResult](/javascript/api/office/office.asyncresult) object.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="496ae-726">Anforderungen</span><span class="sxs-lookup"><span data-stu-id="496ae-726">Requirements</span></span>

|<span data-ttu-id="496ae-727">Anforderung</span><span class="sxs-lookup"><span data-stu-id="496ae-727">Requirement</span></span>| <span data-ttu-id="496ae-728">Wert</span><span class="sxs-lookup"><span data-stu-id="496ae-728">Value</span></span>|
|---|---|
|[<span data-ttu-id="496ae-729">Mindestversion des Postfachanforderungssatzes</span><span class="sxs-lookup"><span data-stu-id="496ae-729">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="496ae-730">1.0</span><span class="sxs-lookup"><span data-stu-id="496ae-730">1.0</span></span>|
|[<span data-ttu-id="496ae-731">Mindestberechtigungsstufe</span><span class="sxs-lookup"><span data-stu-id="496ae-731">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="496ae-732">ReadItem</span><span class="sxs-lookup"><span data-stu-id="496ae-732">ReadItem</span></span>|
|[<span data-ttu-id="496ae-733">Zutreffender Outlook-Modus</span><span class="sxs-lookup"><span data-stu-id="496ae-733">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="496ae-734">Lesen</span><span class="sxs-lookup"><span data-stu-id="496ae-734">Read</span></span>|

##### <a name="examples"></a><span data-ttu-id="496ae-735">Beispiele</span><span class="sxs-lookup"><span data-stu-id="496ae-735">Examples</span></span>

<span data-ttu-id="496ae-736">Im folgenden Code wird eine Zeichenfolge an die `displayReplyForm`-Funktion übergeben.</span><span class="sxs-lookup"><span data-stu-id="496ae-736">The following code passes a string to the `displayReplyForm` function.</span></span>

```
Office.context.mailbox.item.displayReplyForm('hello there');
Office.context.mailbox.item.displayReplyForm('<b>hello there</b>');
```

<span data-ttu-id="496ae-737">Antworten mit einem leeren Textkörper.</span><span class="sxs-lookup"><span data-stu-id="496ae-737">Reply with an empty body.</span></span>

```
Office.context.mailbox.item.displayReplyForm({});
```

<span data-ttu-id="496ae-738">Antworten nur einem Textkörper.</span><span class="sxs-lookup"><span data-stu-id="496ae-738">Reply with just a body.</span></span>

```
Office.context.mailbox.item.displayReplyForm(
{
  'htmlBody' : 'hi'
});
```

<span data-ttu-id="496ae-739">Antworten mit einem Textkörper und einer Dateianlage.</span><span class="sxs-lookup"><span data-stu-id="496ae-739">Reply with a body and a file attachment.</span></span>

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

<span data-ttu-id="496ae-740">Antworten mit einem Textkörper und einer Elementanlage.</span><span class="sxs-lookup"><span data-stu-id="496ae-740">Reply with a body and an item attachment.</span></span>

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

<span data-ttu-id="496ae-741">Antworten Sie mit einem Textkörper, einer Dateianlage, einer Elementanlage und einem Callback.</span><span class="sxs-lookup"><span data-stu-id="496ae-741">Reply with a body, file attachment, item attachment, and a callback.</span></span>

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

#### <a name="getentities--entitiesjavascriptapioutlook13officeentities"></a><span data-ttu-id="496ae-742">getEntities() → {[Entities](/javascript/api/outlook_1_3/office.entities)}</span><span class="sxs-lookup"><span data-stu-id="496ae-742">getEntities() → {[Entities](/javascript/api/outlook_1_3/office.entities)}</span></span>

<span data-ttu-id="496ae-743">Ruft das ausgewählte Element Textkörper gefundenen Entitäten ab.</span><span class="sxs-lookup"><span data-stu-id="496ae-743">Gets the entities found in the selected item's body.</span></span>

> [!NOTE]
> <span data-ttu-id="496ae-744">Diese Methode wird in Outlook für iOS oder Outlook für Android nicht unterstützt.</span><span class="sxs-lookup"><span data-stu-id="496ae-744">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

##### <a name="requirements"></a><span data-ttu-id="496ae-745">Anforderungen</span><span class="sxs-lookup"><span data-stu-id="496ae-745">Requirements</span></span>

|<span data-ttu-id="496ae-746">Anforderung</span><span class="sxs-lookup"><span data-stu-id="496ae-746">Requirement</span></span>| <span data-ttu-id="496ae-747">Wert</span><span class="sxs-lookup"><span data-stu-id="496ae-747">Value</span></span>|
|---|---|
|[<span data-ttu-id="496ae-748">Mindestversion des Postfachanforderungssatzes</span><span class="sxs-lookup"><span data-stu-id="496ae-748">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="496ae-749">1.0</span><span class="sxs-lookup"><span data-stu-id="496ae-749">1.0</span></span>|
|[<span data-ttu-id="496ae-750">Mindestberechtigungsstufe</span><span class="sxs-lookup"><span data-stu-id="496ae-750">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="496ae-751">ReadItem</span><span class="sxs-lookup"><span data-stu-id="496ae-751">ReadItem</span></span>|
|[<span data-ttu-id="496ae-752">Zutreffender Outlook-Modus</span><span class="sxs-lookup"><span data-stu-id="496ae-752">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="496ae-753">Lesen</span><span class="sxs-lookup"><span data-stu-id="496ae-753">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="496ae-754">Gibt zurück:</span><span class="sxs-lookup"><span data-stu-id="496ae-754">Returns:</span></span>

<span data-ttu-id="496ae-755">Typ: [Entitäten](/javascript/api/outlook_1_3/office.entities)</span><span class="sxs-lookup"><span data-stu-id="496ae-755">Type: [Entities](/javascript/api/outlook_1_3/office.entities)</span></span>

##### <a name="example"></a><span data-ttu-id="496ae-756">Beispiel</span><span class="sxs-lookup"><span data-stu-id="496ae-756">Example</span></span>

<span data-ttu-id="496ae-757">Das folgende Beispiel greift auf die Kontakte Entitäten im Textkörper des aktuellen Elements.</span><span class="sxs-lookup"><span data-stu-id="496ae-757">The following example accesses the contacts entities in the current item's body.</span></span>

```
var contacts = Office.context.mailbox.item.getEntities().contacts;
```

#### <a name="getentitiesbytypeentitytype--nullable-arraystringcontactjavascriptapioutlook13officecontactmeetingsuggestionjavascriptapioutlook13officemeetingsuggestionphonenumberjavascriptapioutlook13officephonenumbertasksuggestionjavascriptapioutlook13officetasksuggestion"></a><span data-ttu-id="496ae-758">getEntitiesByType(entityType) → (nullable) {Array.<(String|[Contact](/javascript/api/outlook_1_3/office.contact)|[MeetingSuggestion](/javascript/api/outlook_1_3/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook_1_3/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook_1_3/office.tasksuggestion))>}</span><span class="sxs-lookup"><span data-stu-id="496ae-758">getEntitiesByType(entityType) → (nullable) {Array.<(String|[Contact](/javascript/api/outlook_1_3/office.contact)|[MeetingSuggestion](/javascript/api/outlook_1_3/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook_1_3/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook_1_3/office.tasksuggestion))>}</span></span>

<span data-ttu-id="496ae-759">Ruft ein Array aller Entitäten des angegebenen Entitätstyps im Hauptteil des ausgewählten Elements gefunden.</span><span class="sxs-lookup"><span data-stu-id="496ae-759">Gets an array of all the entities of the specified entity type found in the selected item's body.</span></span>

> [!NOTE]
> <span data-ttu-id="496ae-760">Diese Methode wird in Outlook für iOS oder Outlook für Android nicht unterstützt.</span><span class="sxs-lookup"><span data-stu-id="496ae-760">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

##### <a name="parameters"></a><span data-ttu-id="496ae-761">Parameter:</span><span class="sxs-lookup"><span data-stu-id="496ae-761">Parameters:</span></span>

|<span data-ttu-id="496ae-762">Name</span><span class="sxs-lookup"><span data-stu-id="496ae-762">Name</span></span>| <span data-ttu-id="496ae-763">Typ</span><span class="sxs-lookup"><span data-stu-id="496ae-763">Type</span></span>| <span data-ttu-id="496ae-764">Beschreibung</span><span class="sxs-lookup"><span data-stu-id="496ae-764">Description</span></span>|
|---|---|---|
|`entityType`| [<span data-ttu-id="496ae-765">Office.MailboxEnums.EntityType</span><span class="sxs-lookup"><span data-stu-id="496ae-765">Office.MailboxEnums.EntityType</span></span>](/javascript/api/outlook_1_3/office.mailboxenums.entitytype)|<span data-ttu-id="496ae-766">Einer der Werte der EntityType-Enumeration.</span><span class="sxs-lookup"><span data-stu-id="496ae-766">One of the EntityType enumeration values.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="496ae-767">Anforderungen</span><span class="sxs-lookup"><span data-stu-id="496ae-767">Requirements</span></span>

|<span data-ttu-id="496ae-768">Anforderung</span><span class="sxs-lookup"><span data-stu-id="496ae-768">Requirement</span></span>| <span data-ttu-id="496ae-769">Wert</span><span class="sxs-lookup"><span data-stu-id="496ae-769">Value</span></span>|
|---|---|
|[<span data-ttu-id="496ae-770">Mindestversion des Postfachanforderungssatzes</span><span class="sxs-lookup"><span data-stu-id="496ae-770">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="496ae-771">1.0</span><span class="sxs-lookup"><span data-stu-id="496ae-771">1.0</span></span>|
|[<span data-ttu-id="496ae-772">Mindestberechtigungsstufe</span><span class="sxs-lookup"><span data-stu-id="496ae-772">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="496ae-773">Eingeschränkt</span><span class="sxs-lookup"><span data-stu-id="496ae-773">Restricted</span></span>|
|[<span data-ttu-id="496ae-774">Zutreffender Outlook-Modus</span><span class="sxs-lookup"><span data-stu-id="496ae-774">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="496ae-775">Lesen</span><span class="sxs-lookup"><span data-stu-id="496ae-775">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="496ae-776">Gibt zurück:</span><span class="sxs-lookup"><span data-stu-id="496ae-776">Returns:</span></span>

<span data-ttu-id="496ae-777">Wenn der in `entityType` übergebene Wert kein gültiges Element der `EntityType`-Enumeration ist, gibt die Methode null zurück.</span><span class="sxs-lookup"><span data-stu-id="496ae-777">If the value passed in `entityType` is not a valid member of the `EntityType` enumeration, the method returns null.</span></span> <span data-ttu-id="496ae-778">Wenn keine Entitäten des angegebenen Typs im Textkörper des Elements vorhanden sind, gibt die Methode ein leeres Array zurück.</span><span class="sxs-lookup"><span data-stu-id="496ae-778">If no entities of the specified type are present in the item's body, the method returns an empty array.</span></span> <span data-ttu-id="496ae-779">Andernfalls hängt der Typ der Objekte im zurückgegebenen Array vom Typ der Entität ab, die im `entityType`-Parameter angefordert wurde.</span><span class="sxs-lookup"><span data-stu-id="496ae-779">Otherwise, the type of the objects in the returned array depends on the type of entity requested in the `entityType` parameter.</span></span>

<span data-ttu-id="496ae-780">Während Sie die minimale Berechtigungsstufe für diese Methode **Restricted** ist, erfordern einige Entitätstypen **ReadItem** für den Zugriff, wie in der folgenden Tabelle angegeben wird.</span><span class="sxs-lookup"><span data-stu-id="496ae-780">While the minimum permission level to use this method is **Restricted**, some entity types require **ReadItem** to access, as specified in the following table.</span></span>

| <span data-ttu-id="496ae-781">Wert von `entityType`</span><span class="sxs-lookup"><span data-stu-id="496ae-781">Value of `entityType`</span></span> | <span data-ttu-id="496ae-782">Typ der Objekte im zurückgegebenen Array</span><span class="sxs-lookup"><span data-stu-id="496ae-782">Type of objects in returned array</span></span> | <span data-ttu-id="496ae-783">Erforderliche Berechtigungsstufe</span><span class="sxs-lookup"><span data-stu-id="496ae-783">Required Permission Level</span></span> |
| --- | --- | --- |
| `Address` | <span data-ttu-id="496ae-784">Zeichenfolge</span><span class="sxs-lookup"><span data-stu-id="496ae-784">String</span></span> | <span data-ttu-id="496ae-785">**Restricted**</span><span class="sxs-lookup"><span data-stu-id="496ae-785">**Restricted**</span></span> |
| `Contact` | <span data-ttu-id="496ae-786">Contact</span><span class="sxs-lookup"><span data-stu-id="496ae-786">Contact</span></span> | <span data-ttu-id="496ae-787">**ReadItem**</span><span class="sxs-lookup"><span data-stu-id="496ae-787">**ReadItem**</span></span> |
| `EmailAddress` | <span data-ttu-id="496ae-788">Zeichenfolge</span><span class="sxs-lookup"><span data-stu-id="496ae-788">String</span></span> | <span data-ttu-id="496ae-789">**ReadItem**</span><span class="sxs-lookup"><span data-stu-id="496ae-789">**ReadItem**</span></span> |
| `MeetingSuggestion` | <span data-ttu-id="496ae-790">MeetingSuggestion</span><span class="sxs-lookup"><span data-stu-id="496ae-790">MeetingSuggestion</span></span> | <span data-ttu-id="496ae-791">**ReadItem**</span><span class="sxs-lookup"><span data-stu-id="496ae-791">**ReadItem**</span></span> |
| `PhoneNumber` | <span data-ttu-id="496ae-792">PhoneNumber</span><span class="sxs-lookup"><span data-stu-id="496ae-792">PhoneNumber</span></span> | <span data-ttu-id="496ae-793">**Restricted**</span><span class="sxs-lookup"><span data-stu-id="496ae-793">**Restricted**</span></span> |
| `TaskSuggestion` | <span data-ttu-id="496ae-794">TaskSuggestion</span><span class="sxs-lookup"><span data-stu-id="496ae-794">TaskSuggestion</span></span> | <span data-ttu-id="496ae-795">**ReadItem**</span><span class="sxs-lookup"><span data-stu-id="496ae-795">**ReadItem**</span></span> |
| `URL` | <span data-ttu-id="496ae-796">Zeichenfolge</span><span class="sxs-lookup"><span data-stu-id="496ae-796">String</span></span> | <span data-ttu-id="496ae-797">**Restricted**</span><span class="sxs-lookup"><span data-stu-id="496ae-797">**Restricted**</span></span> |

<span data-ttu-id="496ae-798">Typ: Array.<(String|[Contact](/javascript/api/outlook_1_3/office.contact)|[MeetingSuggestion](/javascript/api/outlook_1_3/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook_1_3/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook_1_3/office.tasksuggestion))></span><span class="sxs-lookup"><span data-stu-id="496ae-798">Type: Array.<(String|[Contact](/javascript/api/outlook_1_3/office.contact)|[MeetingSuggestion](/javascript/api/outlook_1_3/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook_1_3/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook_1_3/office.tasksuggestion))></span></span>

##### <a name="example"></a><span data-ttu-id="496ae-799">Beispiel</span><span class="sxs-lookup"><span data-stu-id="496ae-799">Example</span></span>

<span data-ttu-id="496ae-800">Im folgenden Beispiel wird veranschaulicht, wie auf ein Array von Zeichenfolgen, die Postadressen im Textkörper des aktuellen Elements darstellen.</span><span class="sxs-lookup"><span data-stu-id="496ae-800">The following example shows how to access an array of strings that represent postal addresses in the current item's body.</span></span>

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

#### <a name="getfilteredentitiesbynamename--nullable-arraystringcontactjavascriptapioutlook13officecontactmeetingsuggestionjavascriptapioutlook13officemeetingsuggestionphonenumberjavascriptapioutlook13officephonenumbertasksuggestionjavascriptapioutlook13officetasksuggestion"></a><span data-ttu-id="496ae-801">getFilteredEntitiesByName(name) → (nullable) {Array.<(String|[Contact](/javascript/api/outlook_1_3/office.contact)|[MeetingSuggestion](/javascript/api/outlook_1_3/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook_1_3/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook_1_3/office.tasksuggestion))>}</span><span class="sxs-lookup"><span data-stu-id="496ae-801">getFilteredEntitiesByName(name) → (nullable) {Array.<(String|[Contact](/javascript/api/outlook_1_3/office.contact)|[MeetingSuggestion](/javascript/api/outlook_1_3/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook_1_3/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook_1_3/office.tasksuggestion))>}</span></span>

<span data-ttu-id="496ae-802">Gibt bekannte Entitäten im ausgewählten Element zurück, die den in der XML-Manifestdatei definierten benannten Filter übergeben.</span><span class="sxs-lookup"><span data-stu-id="496ae-802">Returns well-known entities in the selected item that pass the named filter defined in the manifest XML file.</span></span>

> [!NOTE]
> <span data-ttu-id="496ae-803">Diese Methode wird in Outlook für iOS oder Outlook für Android nicht unterstützt.</span><span class="sxs-lookup"><span data-stu-id="496ae-803">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="496ae-804">Die `getFilteredEntitiesByName`-Methode gibt die Entitäten zurück, die dem im [ItemHasKnownEntity](/javascript/office/manifest/rule#itemhasknownentity-rule)-Regelelement der XML-Manifestdatei definierten regulären Ausdruck mit dem angegebenen `FilterName`-Elementwert entsprechen.</span><span class="sxs-lookup"><span data-stu-id="496ae-804">The `getFilteredEntitiesByName` method returns the entities that match the regular expression defined in the [ItemHasKnownEntity](/javascript/office/manifest/rule#itemhasknownentity-rule) rule element in the manifest XML file with the specified `FilterName` element value.</span></span>

##### <a name="parameters"></a><span data-ttu-id="496ae-805">Parameter:</span><span class="sxs-lookup"><span data-stu-id="496ae-805">Parameters:</span></span>

|<span data-ttu-id="496ae-806">Name</span><span class="sxs-lookup"><span data-stu-id="496ae-806">Name</span></span>| <span data-ttu-id="496ae-807">Typ</span><span class="sxs-lookup"><span data-stu-id="496ae-807">Type</span></span>| <span data-ttu-id="496ae-808">Beschreibung</span><span class="sxs-lookup"><span data-stu-id="496ae-808">Description</span></span>|
|---|---|---|
|`name`| <span data-ttu-id="496ae-809">Zeichenfolge</span><span class="sxs-lookup"><span data-stu-id="496ae-809">String</span></span>|<span data-ttu-id="496ae-810">Der Name des `ItemHasKnownEntity`-Regelelements, das den entsprechenden Filter definiert.</span><span class="sxs-lookup"><span data-stu-id="496ae-810">The name of the `ItemHasKnownEntity` rule element that defines the filter to match.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="496ae-811">Anforderungen</span><span class="sxs-lookup"><span data-stu-id="496ae-811">Requirements</span></span>

|<span data-ttu-id="496ae-812">Anforderung</span><span class="sxs-lookup"><span data-stu-id="496ae-812">Requirement</span></span>| <span data-ttu-id="496ae-813">Wert</span><span class="sxs-lookup"><span data-stu-id="496ae-813">Value</span></span>|
|---|---|
|[<span data-ttu-id="496ae-814">Mindestversion des Postfachanforderungssatzes</span><span class="sxs-lookup"><span data-stu-id="496ae-814">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="496ae-815">1.0</span><span class="sxs-lookup"><span data-stu-id="496ae-815">1.0</span></span>|
|[<span data-ttu-id="496ae-816">Mindestberechtigungsstufe</span><span class="sxs-lookup"><span data-stu-id="496ae-816">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="496ae-817">ReadItem</span><span class="sxs-lookup"><span data-stu-id="496ae-817">ReadItem</span></span>|
|[<span data-ttu-id="496ae-818">Zutreffender Outlook-Modus</span><span class="sxs-lookup"><span data-stu-id="496ae-818">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="496ae-819">Lesen</span><span class="sxs-lookup"><span data-stu-id="496ae-819">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="496ae-820">Gibt zurück:</span><span class="sxs-lookup"><span data-stu-id="496ae-820">Returns:</span></span>

<span data-ttu-id="496ae-p153">Befindet sich kein `ItemHasKnownEntity`-Element im Manifest mit einem `FilterName`-Elementwert, der dem `name`-Parameter entspricht, gibt die Methode `null` zurück. Wenn der `name`-Parameter einem `ItemHasKnownEntity`-Element im Manifest entspricht, es aber keine entsprechenden Entitäten im aktuellen Element gibt, gibt die Methode ein leeres Array zurück.</span><span class="sxs-lookup"><span data-stu-id="496ae-p153">If there is no `ItemHasKnownEntity` element in the manifest with a `FilterName` element value that matches the `name` parameter, the method returns `null`. If the `name` parameter does match an `ItemHasKnownEntity` element in the manifest, but there are no entities in the current item that match, the method return an empty array.</span></span>

<span data-ttu-id="496ae-823">Typ: Array.<(String|[Contact](/javascript/api/outlook_1_3/office.contact)|[MeetingSuggestion](/javascript/api/outlook_1_3/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook_1_3/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook_1_3/office.tasksuggestion))></span><span class="sxs-lookup"><span data-stu-id="496ae-823">Type: Array.<(String|[Contact](/javascript/api/outlook_1_3/office.contact)|[MeetingSuggestion](/javascript/api/outlook_1_3/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook_1_3/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook_1_3/office.tasksuggestion))></span></span>

#### <a name="getregexmatches--object"></a><span data-ttu-id="496ae-824">getRegExMatches() → {Object}</span><span class="sxs-lookup"><span data-stu-id="496ae-824">getRegExMatches() → {Object}</span></span>

<span data-ttu-id="496ae-825">Gibt Zeichenfolgenwerte im ausgewählten Element zurück, die den in der XML-Manifestdatei definierten regulären Ausdrücken entsprechen.</span><span class="sxs-lookup"><span data-stu-id="496ae-825">Returns string values in the selected item that match the regular expressions defined in the manifest XML file.</span></span>

> [!NOTE]
> <span data-ttu-id="496ae-826">Diese Methode wird in Outlook für iOS oder Outlook für Android nicht unterstützt.</span><span class="sxs-lookup"><span data-stu-id="496ae-826">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="496ae-p154">Die `getRegExMatches`-Methode gibt die Zeichenfolgen zurück, die dem im `ItemHasRegularExpressionMatch`- oder `ItemHasKnownEntity`-Regelelement der XML-Manifestdatei definierten regulären Ausdruck entsprechen. Bei einer `ItemHasRegularExpressionMatch`-Regel muss eine entsprechende Zeichenfolge in der Eigenschaft des Elements vorhanden sein, das von dieser Regel angegeben wird. Der einfache `PropertyName`-Typ definiert die unterstützten Eigenschaften.</span><span class="sxs-lookup"><span data-stu-id="496ae-p154">The `getRegExMatches` method returns the strings that match the regular expression defined in each `ItemHasRegularExpressionMatch` or `ItemHasKnownEntity` rule element in the manifest XML file. For an `ItemHasRegularExpressionMatch` rule, a matching string has to occur in the property of the item that is specified by that rule. The `PropertyName` simple type defines the supported properties.</span></span>

<span data-ttu-id="496ae-830">Nehmen Sie z. B. an, dass ein Add-In-Manifest das folgende `Rule`-Element aufweist:</span><span class="sxs-lookup"><span data-stu-id="496ae-830">For example, consider an add-in manifest has the following `Rule` element:</span></span>

```
<Rule xsi:type="RuleCollection" Mode="And">
  <Rule xsi:type="ItemIs" FormType="Read" ItemType="Message" />
  <Rule xsi:type="RuleCollection" Mode="Or">
    <Rule xsi:type="ItemHasRegularExpressionMatch" RegExName="fruits" RegExValue="apple|banana|coconut" PropertyName="BodyAsPlaintext" IgnoreCase="true" />
    <Rule xsi:type="ItemHasRegularExpressionMatch" RegExName="veggies" RegExValue="tomato|onion|spinach|broccoli" PropertyName="BodyAsPlaintext" IgnoreCase="true" />
  </Rule>
</Rule>
```

<span data-ttu-id="496ae-831">Das von `getRegExMatches` zurückgegebene Objekt hätte zwei Eigenschaften: `fruits` und `veggies`.</span><span class="sxs-lookup"><span data-stu-id="496ae-831">The object returned from `getRegExMatches` would have two properties: `fruits` and `veggies`.</span></span>

```
{
  'fruits': ['apple','banana','Banana','coconut'],
  'veggies': ['tomato','onion','spinach','broccoli']
}
```

<span data-ttu-id="496ae-p155">Wenn Sie eine `ItemHasRegularExpressionMatch`-Regel für die Textkörpereigenschaft eines Elements festlegen, sollte der reguläre Ausdruck den Textkörper weiter filtern und nicht versuchen, den gesamten Textkörper des Elements zurückzugeben. Wenn der gesamte Textkörper eines Elements mit einem regulären Ausdruck wie `.*` abgerufen wird, werden nicht immer die gewünschten Ergebnisse erzielt. Verwenden Sie stattdessen die [`Body.getAsync`](/javascript/api/outlook_1_3/office.body#getasync-coerciontype--options--callback-)-Methode, um den gesamten Textkörper abzurufen.</span><span class="sxs-lookup"><span data-stu-id="496ae-p155">If you specify an `ItemHasRegularExpressionMatch` rule on the body property of an item, the regular expression should further filter the body and should not attempt to return the entire body of the item. Using a regular expression such as `.*` to obtain the entire body of an item does not always return the expected results. Instead, use the [`Body.getAsync`](/javascript/api/outlook_1_3/office.body#getasync-coerciontype--options--callback-) method to retrieve the entire body.</span></span>

##### <a name="requirements"></a><span data-ttu-id="496ae-835">Anforderungen</span><span class="sxs-lookup"><span data-stu-id="496ae-835">Requirements</span></span>

|<span data-ttu-id="496ae-836">Anforderung</span><span class="sxs-lookup"><span data-stu-id="496ae-836">Requirement</span></span>| <span data-ttu-id="496ae-837">Wert</span><span class="sxs-lookup"><span data-stu-id="496ae-837">Value</span></span>|
|---|---|
|[<span data-ttu-id="496ae-838">Mindestversion des Postfachanforderungssatzes</span><span class="sxs-lookup"><span data-stu-id="496ae-838">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="496ae-839">1.0</span><span class="sxs-lookup"><span data-stu-id="496ae-839">1.0</span></span>|
|[<span data-ttu-id="496ae-840">Mindestberechtigungsstufe</span><span class="sxs-lookup"><span data-stu-id="496ae-840">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="496ae-841">ReadItem</span><span class="sxs-lookup"><span data-stu-id="496ae-841">ReadItem</span></span>|
|[<span data-ttu-id="496ae-842">Zutreffender Outlook-Modus</span><span class="sxs-lookup"><span data-stu-id="496ae-842">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="496ae-843">Lesen</span><span class="sxs-lookup"><span data-stu-id="496ae-843">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="496ae-844">Gibt zurück:</span><span class="sxs-lookup"><span data-stu-id="496ae-844">Returns:</span></span>

<span data-ttu-id="496ae-p156">Ein Objekt mit Arrays aus Zeichenfolgen, die den in der XML-Manifestdatei definierten regulären Ausdrücken entsprechen. Der Name der einzelnen Arrays ist gleich dem entsprechenden Wert des `RegExName`-Attributs der entsprechenden `ItemHasRegularExpressionMatch`-Regel oder des `FilterName`-Attributs der entsprechenden `ItemHasKnownEntity`-Regel.</span><span class="sxs-lookup"><span data-stu-id="496ae-p156">An object that contains arrays of strings that match the regular expressions defined in the manifest XML file. The name of each array is equal to the corresponding value of the `RegExName` attribute of the matching `ItemHasRegularExpressionMatch` rule or the `FilterName` attribute of the matching `ItemHasKnownEntity` rule.</span></span>

<dl class="param-type"><span data-ttu-id="496ae-847">

<dt>
Typ</dt>


</span><span class="sxs-lookup"><span data-stu-id="496ae-847">

<dt>Type</dt>

</span></span><dd><span data-ttu-id="496ae-848">Object</span><span class="sxs-lookup"><span data-stu-id="496ae-848">Object</span></span></dd>

</dl>

##### <a name="example"></a><span data-ttu-id="496ae-849">Beispiel</span><span class="sxs-lookup"><span data-stu-id="496ae-849">Example</span></span>

<span data-ttu-id="496ae-850">Das folgende Beispiel zeigt, wie Sie auf das Array von Übereinstimmungen für die regulären <rule>-Ausdruckselemente `fruits` und `veggies` zugreifen, die im Manifest angegeben sind.</rule></span><span class="sxs-lookup"><span data-stu-id="496ae-850">The following example shows how to access the array of matches for the regular expression <rule>elements `fruits` and `veggies`, which are specified in the manifest.</rule></span></span>

```
var allMatches = Office.context.mailbox.item.getRegExMatches();
var fruits = allMatches.fruits;
var veges = allMatches.veggies;
```

#### <a name="getregexmatchesbynamename--nullable-array-string-"></a><span data-ttu-id="496ae-851">getRegExMatchesByName(name) → (Nullwerte zulassen) {Array. < Zeichenfolge >}</span><span class="sxs-lookup"><span data-stu-id="496ae-851">getRegExMatchesByName(name) → (nullable) {Array.< String >}</span></span>

<span data-ttu-id="496ae-852">Gibt Zeichenfolgenwerte im ausgewählten Element zurück, die dem in der XML-Manifestdatei definierten benannten regulären Ausdruck entsprechen.</span><span class="sxs-lookup"><span data-stu-id="496ae-852">Returns string values in the selected item that match the named regular expression defined in the manifest XML file.</span></span>

> [!NOTE]
> <span data-ttu-id="496ae-853">Diese Methode wird in Outlook für iOS oder Outlook für Android nicht unterstützt.</span><span class="sxs-lookup"><span data-stu-id="496ae-853">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="496ae-854">Die `getRegExMatchesByName`-Methode gibt die Zeichenfolgen zurück, die dem im `ItemHasRegularExpressionMatch`-Regelelement der XML-Manifestdatei definierten regulären Ausdruck mit dem angegebenen `RegExName`-Elementwert entsprechen.</span><span class="sxs-lookup"><span data-stu-id="496ae-854">The `getRegExMatchesByName` method returns the strings that match the regular expression defined in the `ItemHasRegularExpressionMatch` rule element in the manifest XML file with the specified `RegExName` element value.</span></span>

<span data-ttu-id="496ae-p157">Wenn Sie eine `ItemHasRegularExpressionMatch`-Regel für die Textkörpereigenschaft eines Elements festlegen, sollte der reguläre Ausdruck den Textkörper weiter filtern und nicht versuchen, den gesamten Textkörper des Elements zurückzugeben. Wenn der gesamte Textkörper eines Elements mit einem regulären Ausdruck wie `.*` abgerufen wird, werden nicht immer die gewünschten Ergebnisse erzielt.</span><span class="sxs-lookup"><span data-stu-id="496ae-p157">If you specify an `ItemHasRegularExpressionMatch` rule on the body property of an item, the regular expression should further filter the body and should not attempt to return the entire body of the item. Using a regular expression such as `.*` to obtain the entire body of an item does not always return the expected results.</span></span>

##### <a name="parameters"></a><span data-ttu-id="496ae-857">Parameter:</span><span class="sxs-lookup"><span data-stu-id="496ae-857">Parameters:</span></span>

|<span data-ttu-id="496ae-858">Name</span><span class="sxs-lookup"><span data-stu-id="496ae-858">Name</span></span>| <span data-ttu-id="496ae-859">Typ</span><span class="sxs-lookup"><span data-stu-id="496ae-859">Type</span></span>| <span data-ttu-id="496ae-860">Beschreibung</span><span class="sxs-lookup"><span data-stu-id="496ae-860">Description</span></span>|
|---|---|---|
|`name`| <span data-ttu-id="496ae-861">Zeichenfolge</span><span class="sxs-lookup"><span data-stu-id="496ae-861">String</span></span>|<span data-ttu-id="496ae-862">Der Name des `ItemHasRegularExpressionMatch`-Regelelements, das den entsprechenden Filter definiert.</span><span class="sxs-lookup"><span data-stu-id="496ae-862">The name of the `ItemHasRegularExpressionMatch` rule element that defines the filter to match.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="496ae-863">Anforderungen</span><span class="sxs-lookup"><span data-stu-id="496ae-863">Requirements</span></span>

|<span data-ttu-id="496ae-864">Anforderung</span><span class="sxs-lookup"><span data-stu-id="496ae-864">Requirement</span></span>| <span data-ttu-id="496ae-865">Wert</span><span class="sxs-lookup"><span data-stu-id="496ae-865">Value</span></span>|
|---|---|
|[<span data-ttu-id="496ae-866">Mindestversion des Postfachanforderungssatzes</span><span class="sxs-lookup"><span data-stu-id="496ae-866">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="496ae-867">1.0</span><span class="sxs-lookup"><span data-stu-id="496ae-867">1.0</span></span>|
|[<span data-ttu-id="496ae-868">Mindestberechtigungsstufe</span><span class="sxs-lookup"><span data-stu-id="496ae-868">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="496ae-869">ReadItem</span><span class="sxs-lookup"><span data-stu-id="496ae-869">ReadItem</span></span>|
|[<span data-ttu-id="496ae-870">Zutreffender Outlook-Modus</span><span class="sxs-lookup"><span data-stu-id="496ae-870">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="496ae-871">Lesen</span><span class="sxs-lookup"><span data-stu-id="496ae-871">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="496ae-872">Gibt zurück:</span><span class="sxs-lookup"><span data-stu-id="496ae-872">Returns:</span></span>

<span data-ttu-id="496ae-873">Ein Array mit den Zeichenfolgen, die dem in der XML-Manifestdatei definierten regulären Ausdruck entsprechen.</span><span class="sxs-lookup"><span data-stu-id="496ae-873">An array that contains the strings that match the regular expression defined in the manifest XML file.</span></span>

<dl class="param-type"><span data-ttu-id="496ae-874">

<dt>Typ</dt>

</span><span class="sxs-lookup"><span data-stu-id="496ae-874">

<dt>Type</dt>

</span></span><dd><span data-ttu-id="496ae-875">Array. < Zeichenfolge ></span><span class="sxs-lookup"><span data-stu-id="496ae-875">Array.< String ></span></span></dd>

</dl>

##### <a name="example"></a><span data-ttu-id="496ae-876">Beispiel</span><span class="sxs-lookup"><span data-stu-id="496ae-876">Example</span></span>

```
var fruits = Office.context.mailbox.item.getRegExMatchesByName("fruits");
var veggies = Office.context.mailbox.item.getRegExMatchesByName("veggies");
```

####  <a name="getselecteddataasynccoerciontype-options-callback--string"></a><span data-ttu-id="496ae-877">getSelectedDataAsync(coercionType, [options], callback) → {String}</span><span class="sxs-lookup"><span data-stu-id="496ae-877">getSelectedDataAsync(coercionType, [options], callback) → {String}</span></span>

<span data-ttu-id="496ae-878">Gibt asynchron ausgewählte Daten aus dem Betreff oder Textkörper einer Nachricht zurück.</span><span class="sxs-lookup"><span data-stu-id="496ae-878">Asynchronously returns selected data from the subject or body of a message.</span></span>

<span data-ttu-id="496ae-p158">Wenn keine Auswahl vorhanden ist, aber der Cursor sich im Nachrichtentext oder Betreff befindet, gibt die Methode für die ausgewählten Daten NULL zurück. Wenn ein anderes Feld als der Textkörper oder Betreff ausgewählt ist, gibt die Methode den `InvalidSelection`-Fehler zurück.</span><span class="sxs-lookup"><span data-stu-id="496ae-p158">If there is no selection but the cursor is in the body or subject, the method returns null for the selected data. If a field other than the body or subject is selected, the method returns the `InvalidSelection` error.</span></span>

##### <a name="parameters"></a><span data-ttu-id="496ae-881">Parameter:</span><span class="sxs-lookup"><span data-stu-id="496ae-881">Parameters:</span></span>

|<span data-ttu-id="496ae-882">Name</span><span class="sxs-lookup"><span data-stu-id="496ae-882">Name</span></span>| <span data-ttu-id="496ae-883">Typ</span><span class="sxs-lookup"><span data-stu-id="496ae-883">Type</span></span>| <span data-ttu-id="496ae-884">Attribute</span><span class="sxs-lookup"><span data-stu-id="496ae-884">Attributes</span></span>| <span data-ttu-id="496ae-885">Beschreibung</span><span class="sxs-lookup"><span data-stu-id="496ae-885">Description</span></span>|
|---|---|---|---|
|`coercionType`| [<span data-ttu-id="496ae-886">Office.CoercionType</span><span class="sxs-lookup"><span data-stu-id="496ae-886">Office.CoercionType</span></span>](office.md#coerciontype-string)||<span data-ttu-id="496ae-p159">Fordert ein Format für die Daten an. Wenn es sich um Texthandelt, gibt die Methode den unformatierten Text als Zeichenfolge zurück und entfernt eventuell vorhandene HTML-Tags. Wenn es sich um HTML handelt, gibt die Methode den ausgewählten Text zurück, entweder als unformatierten Text oder als HTML.</span><span class="sxs-lookup"><span data-stu-id="496ae-p159">Requests a format for the data. If Text, the method returns the plain text as a string , removing any HTML tags present. If HTML, the method returns the selected text, whether it is plaintext or HTML.</span></span>|
|`options`| <span data-ttu-id="496ae-890">Object</span><span class="sxs-lookup"><span data-stu-id="496ae-890">Object</span></span>| <span data-ttu-id="496ae-891">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="496ae-891">&lt;optional&gt;</span></span>|<span data-ttu-id="496ae-892">Ein Objektliteral, das mindestens eine der folgenden Eigenschaften enthält.</span><span class="sxs-lookup"><span data-stu-id="496ae-892">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`| <span data-ttu-id="496ae-893">Object</span><span class="sxs-lookup"><span data-stu-id="496ae-893">Object</span></span>| <span data-ttu-id="496ae-894">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="496ae-894">&lt;optional&gt;</span></span>|<span data-ttu-id="496ae-895">Entwickler können ein Objekt bereitstellen, auf das sie in der Callbackmethode zugreifen möchten.</span><span class="sxs-lookup"><span data-stu-id="496ae-895">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`| <span data-ttu-id="496ae-896">Funktion</span><span class="sxs-lookup"><span data-stu-id="496ae-896">function</span></span>||<span data-ttu-id="496ae-897">Wenn die Methode abgeschlossen wird, wird die Funktion , die im `callback`-Parameter übergeben wird, mit einem einzelnen Parameter (`asyncResult`) aufgerufen, der ein [`AsyncResult`](/javascript/api/office/office.asyncresult)-Objekt darstellt.</span><span class="sxs-lookup"><span data-stu-id="496ae-897">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span><br/><br/><span data-ttu-id="496ae-898">Rufen Sie für den Zugriff auf die ausgewählten Daten aus der Rückrufmethode `asyncResult.value.data` auf.</span><span class="sxs-lookup"><span data-stu-id="496ae-898">To access the selected data from the callback method, call `asyncResult.value.data`.</span></span> <span data-ttu-id="496ae-899">Rufen Sie die Source-Eigenschaft für den Zugriff, die die Auswahl stammen, `asyncResult.value.sourceProperty`, die entweder sein `body` oder `subject`.</span><span class="sxs-lookup"><span data-stu-id="496ae-899">To access the source property that the selection comes from, call `asyncResult.value.sourceProperty`, which will be either `body` or `subject`.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="496ae-900">Anforderungen</span><span class="sxs-lookup"><span data-stu-id="496ae-900">Requirements</span></span>

|<span data-ttu-id="496ae-901">Anforderung</span><span class="sxs-lookup"><span data-stu-id="496ae-901">Requirement</span></span>| <span data-ttu-id="496ae-902">Wert</span><span class="sxs-lookup"><span data-stu-id="496ae-902">Value</span></span>|
|---|---|
|[<span data-ttu-id="496ae-903">Mindestversion des Postfachanforderungssatzes</span><span class="sxs-lookup"><span data-stu-id="496ae-903">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="496ae-904">1.2</span><span class="sxs-lookup"><span data-stu-id="496ae-904">1.2</span></span>|
|[<span data-ttu-id="496ae-905">Mindestberechtigungsstufe</span><span class="sxs-lookup"><span data-stu-id="496ae-905">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="496ae-906">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="496ae-906">ReadWriteItem</span></span>|
|[<span data-ttu-id="496ae-907">Zutreffender Outlook-Modus</span><span class="sxs-lookup"><span data-stu-id="496ae-907">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="496ae-908">Verfassen</span><span class="sxs-lookup"><span data-stu-id="496ae-908">Compose</span></span>|

##### <a name="returns"></a><span data-ttu-id="496ae-909">Gibt zurück:</span><span class="sxs-lookup"><span data-stu-id="496ae-909">Returns:</span></span>

<span data-ttu-id="496ae-910">Die ausgewählten Daten als Zeichenfolge mit dem durch `coercionType` bestimmten Format.</span><span class="sxs-lookup"><span data-stu-id="496ae-910">The selected data as a string with format determined by `coercionType`.</span></span>

<dl class="param-type"><span data-ttu-id="496ae-911">

<dt>
Typ</dt>


</span><span class="sxs-lookup"><span data-stu-id="496ae-911">

<dt>Type</dt>

</span></span><dd><span data-ttu-id="496ae-912">Zeichenfolge</span><span class="sxs-lookup"><span data-stu-id="496ae-912">String</span></span></dd>

</dl>

##### <a name="example"></a><span data-ttu-id="496ae-913">Beispiel</span><span class="sxs-lookup"><span data-stu-id="496ae-913">Example</span></span>

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

####  <a name="loadcustompropertiesasynccallback-usercontext"></a><span data-ttu-id="496ae-914">loadCustomPropertiesAsync(callback, [userContext])</span><span class="sxs-lookup"><span data-stu-id="496ae-914">loadCustomPropertiesAsync(callback, [userContext])</span></span>

<span data-ttu-id="496ae-915">Lädt asynchron benutzerdefinierte Eigenschaften für dieses Add-In für das ausgewählte Element.</span><span class="sxs-lookup"><span data-stu-id="496ae-915">Asynchronously loads custom properties for this add-in on the selected item.</span></span>

<span data-ttu-id="496ae-p161">Benutzerdefinierte Eigenschaften werden als Schlüssel-/Wert-Paare pro App und pro Element gespeichert. Diese Methode gibt ein `CustomProperties`-Objekt im Rückruf zurück, das Methoden für den Zugriff auf die benutzerdefinierten Eigenschaften für das aktuelle Element und das aktuelle Add-In bietet. Benutzerdefinierte Eigenschaften sind für das Element nicht verschlüsselt und sollten darum nicht als sicherer Speicher verwendet werden.</span><span class="sxs-lookup"><span data-stu-id="496ae-p161">Custom properties are stored as key/value pairs on a per-app, per-item basis. This method returns a `CustomProperties` object in the callback, which provides methods to access the custom properties specific to the current item and the current add-in. Custom properties are not encrypted on the item, so this should not be used as secure storage.</span></span>

##### <a name="parameters"></a><span data-ttu-id="496ae-919">Parameter:</span><span class="sxs-lookup"><span data-stu-id="496ae-919">Parameters:</span></span>

|<span data-ttu-id="496ae-920">Name</span><span class="sxs-lookup"><span data-stu-id="496ae-920">Name</span></span>| <span data-ttu-id="496ae-921">Typ</span><span class="sxs-lookup"><span data-stu-id="496ae-921">Type</span></span>| <span data-ttu-id="496ae-922">Attribute</span><span class="sxs-lookup"><span data-stu-id="496ae-922">Attributes</span></span>| <span data-ttu-id="496ae-923">Beschreibung</span><span class="sxs-lookup"><span data-stu-id="496ae-923">Description</span></span>|
|---|---|---|---|
|`callback`| <span data-ttu-id="496ae-924">Funktion</span><span class="sxs-lookup"><span data-stu-id="496ae-924">function</span></span>||<span data-ttu-id="496ae-925">Wenn die Methode abgeschlossen wird, wird die Funktion , die im `callback`-Parameter übergeben wird, mit einem einzelnen Parameter (`asyncResult`) aufgerufen, der ein [`AsyncResult`](/javascript/api/office/office.asyncresult)-Objekt darstellt.</span><span class="sxs-lookup"><span data-stu-id="496ae-925">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span><br/><br/><span data-ttu-id="496ae-926">Die benutzerdefinierten Eigenschaften werden als [`CustomProperties`](/javascript/api/outlook_1_3/office.customproperties)-Objekt in der `asyncResult.value`-Eigenschaft bereitgestellt.</span><span class="sxs-lookup"><span data-stu-id="496ae-926">The custom properties are provided as a [`CustomProperties`](/javascript/api/outlook_1_3/office.customproperties) object in the `asyncResult.value` property.</span></span> <span data-ttu-id="496ae-927">Dieses Objekt kann zum Abrufen, festlegen und Entfernen benutzerdefinierter Eigenschaften aus dem Element und speichern Sie Änderungen an den benutzerdefinierten Eigenschaftensatz zurück an den Server verwendet werden.</span><span class="sxs-lookup"><span data-stu-id="496ae-927">This object can be used to get, set, and remove custom properties from the item and save changes to the custom property set back to the server.</span></span>|
|`userContext`| <span data-ttu-id="496ae-928">Objekt</span><span class="sxs-lookup"><span data-stu-id="496ae-928">Object</span></span>| <span data-ttu-id="496ae-929">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="496ae-929">&lt;optional&gt;</span></span>|<span data-ttu-id="496ae-930">Entwickler können ein beliebiges Objekt bereitstellen, den sie in der Rückruffunktion zugreifen möchten.</span><span class="sxs-lookup"><span data-stu-id="496ae-930">Developers can provide any object they wish to access in the callback function.</span></span> <span data-ttu-id="496ae-931">Dieses Objekt zugegriffen werden kann, indem die `asyncResult.asyncContext` -Eigenschaft in der Callback-Funktion.</span><span class="sxs-lookup"><span data-stu-id="496ae-931">This object can be accessed by the `asyncResult.asyncContext` property in the callback function.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="496ae-932">Anforderungen</span><span class="sxs-lookup"><span data-stu-id="496ae-932">Requirements</span></span>

|<span data-ttu-id="496ae-933">Anforderung</span><span class="sxs-lookup"><span data-stu-id="496ae-933">Requirement</span></span>| <span data-ttu-id="496ae-934">Wert</span><span class="sxs-lookup"><span data-stu-id="496ae-934">Value</span></span>|
|---|---|
|[<span data-ttu-id="496ae-935">Mindestversion des Postfachanforderungssatzes</span><span class="sxs-lookup"><span data-stu-id="496ae-935">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="496ae-936">1.0</span><span class="sxs-lookup"><span data-stu-id="496ae-936">1.0</span></span>|
|[<span data-ttu-id="496ae-937">Mindestberechtigungsstufe</span><span class="sxs-lookup"><span data-stu-id="496ae-937">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="496ae-938">ReadItem</span><span class="sxs-lookup"><span data-stu-id="496ae-938">ReadItem</span></span>|
|[<span data-ttu-id="496ae-939">Zutreffender Outlook-Modus</span><span class="sxs-lookup"><span data-stu-id="496ae-939">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="496ae-940">Verfassen oder Lesen</span><span class="sxs-lookup"><span data-stu-id="496ae-940">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="496ae-941">Beispiel</span><span class="sxs-lookup"><span data-stu-id="496ae-941">Example</span></span>

<span data-ttu-id="496ae-p164">Das folgende Codebeispiel veranschaulicht die Verwendung der `loadCustomPropertiesAsync`-Methode zum asynchronen Laden der benutzerdefinierten Eigenschaften, die für das aktuelle Element spezifisch sind. Des Weiteren veranschaulicht das Beispiel die Verwendung der `CustomProperties.saveAsync`-Methode zum Speichern der Eigenschaften auf dem Server. Nach dem Laden der benutzerdefinierten Eigenschaften wird in dem Codebeispiel die `CustomProperties.get`-Methode dazu verwendet, die benutzerdefinierte `myProp`-Eigenschaft zu lesen. Die `CustomProperties.set`-Methode wird dazu verwendet, die benutzerdefinierte `otherProp`-Eigenschaft zu schreiben. Schließlich wird die `saveAsync`-Methode aufgerufen, um die benutzerdefinierten Eigenschaften zu speichern.</span><span class="sxs-lookup"><span data-stu-id="496ae-p164">The following code example shows how to use the `loadCustomPropertiesAsync` method to asynchronously load custom properties that are specific to the current item. The example also shows how to use the `CustomProperties.saveAsync` method to save these properties back to the server. After loading the custom properties, the code sample uses the `CustomProperties.get` method to read the custom property `myProp`, the `CustomProperties.set` method to write the custom property `otherProp`, and then finally calls the `saveAsync` method to save the custom properties.</span></span>

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

####  <a name="removeattachmentasyncattachmentid-options-callback"></a><span data-ttu-id="496ae-945">removeAttachmentAsync(attachmentId, [options], [callback])</span><span class="sxs-lookup"><span data-stu-id="496ae-945">removeAttachmentAsync(attachmentId, [options], [callback])</span></span>

<span data-ttu-id="496ae-946">Entfernt eine Anlage aus einer Nachricht oder einem Termin.</span><span class="sxs-lookup"><span data-stu-id="496ae-946">Removes an attachment from a message or appointment.</span></span>

<span data-ttu-id="496ae-p165">Die `removeAttachmentAsync`-Methode entfernt die Anlage mit dem angegebenen Bezeichner aus dem Element. Als bewährte Vorgehensweise sollten Sie den Anlagenbezeichner nur dann zum Entfernen einer Anlage verwenden, wenn die gleiche Mail-App die Anlage in der gleichen Sitzung hinzugefügt hat. In Outlook Web App und OWA für Geräte ist der Anlagenbezeichner nur innerhalb der gleichen Sitzung gültig. Eine Sitzung ist abgeschlossen, wenn der Benutzer die App schließt, oder wenn der Benutzer in einem eingebetteten Formular mit dem Verfassen beginnt und den Vorgang dann in einem separaten Fenster fortsetzt.</span><span class="sxs-lookup"><span data-stu-id="496ae-p165">The `removeAttachmentAsync` method removes the attachment with the specified identifier from the item. As a best practice, you should use the attachment identifier to remove an attachment only if the same mail app has added that attachment in the same session. In Outlook Web App and OWA for Devices, the attachment identifier is valid only within the same session. A session is over when the user closes the app, or if the user starts composing in an inline form and subsequently pops out the inline form to continue in a separate window.</span></span>

##### <a name="parameters"></a><span data-ttu-id="496ae-951">Parameter:</span><span class="sxs-lookup"><span data-stu-id="496ae-951">Parameters:</span></span>

|<span data-ttu-id="496ae-952">Name</span><span class="sxs-lookup"><span data-stu-id="496ae-952">Name</span></span>| <span data-ttu-id="496ae-953">Typ</span><span class="sxs-lookup"><span data-stu-id="496ae-953">Type</span></span>| <span data-ttu-id="496ae-954">Attribute</span><span class="sxs-lookup"><span data-stu-id="496ae-954">Attributes</span></span>| <span data-ttu-id="496ae-955">Beschreibung</span><span class="sxs-lookup"><span data-stu-id="496ae-955">Description</span></span>|
|---|---|---|---|
|`attachmentId`| <span data-ttu-id="496ae-956">Zeichenfolge</span><span class="sxs-lookup"><span data-stu-id="496ae-956">String</span></span>||<span data-ttu-id="496ae-p166">Der Bezeichner des Anhangs, der entfernt werden soll. Die maximale Länge der Zeichenfolge ist 100 Zeichen.</span><span class="sxs-lookup"><span data-stu-id="496ae-p166">The identifier of the attachment to remove. The maximum length of the string is 100 characters.</span></span>|
|`options`| <span data-ttu-id="496ae-959">Object</span><span class="sxs-lookup"><span data-stu-id="496ae-959">Object</span></span>| <span data-ttu-id="496ae-960">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="496ae-960">&lt;optional&gt;</span></span>|<span data-ttu-id="496ae-961">Ein Objektliteral, das mindestens eine der folgenden Eigenschaften enthält.</span><span class="sxs-lookup"><span data-stu-id="496ae-961">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`| <span data-ttu-id="496ae-962">Object</span><span class="sxs-lookup"><span data-stu-id="496ae-962">Object</span></span>| <span data-ttu-id="496ae-963">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="496ae-963">&lt;optional&gt;</span></span>|<span data-ttu-id="496ae-964">Entwickler können ein Objekt bereitstellen, auf das sie in der Callbackmethode zugreifen möchten.</span><span class="sxs-lookup"><span data-stu-id="496ae-964">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`| <span data-ttu-id="496ae-965">Funktion</span><span class="sxs-lookup"><span data-stu-id="496ae-965">function</span></span>| <span data-ttu-id="496ae-966">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="496ae-966">&lt;optional&gt;</span></span>|<span data-ttu-id="496ae-967">Wenn die Methode abgeschlossen wird, wird die Funktion , die im `callback`-Parameter übergeben wird, mit einem einzelnen Parameter (`asyncResult`) aufgerufen, der ein [`AsyncResult`](/javascript/api/office/office.asyncresult)-Objekt darstellt.</span><span class="sxs-lookup"><span data-stu-id="496ae-967">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span> <br/><span data-ttu-id="496ae-968">Wenn beim Entfernen der Anlage ein Fehler auftritt, enthält die Eigenschaft `asyncResult.error` einen Fehlercode mit dem Grund für den Fehler.</span><span class="sxs-lookup"><span data-stu-id="496ae-968">If removing the attachment fails, the `asyncResult.error` property will contain an error code with the reason for the failure.</span></span>|

##### <a name="errors"></a><span data-ttu-id="496ae-969">Fehler</span><span class="sxs-lookup"><span data-stu-id="496ae-969">Errors</span></span>

| <span data-ttu-id="496ae-970">Fehlercode</span><span class="sxs-lookup"><span data-stu-id="496ae-970">Error code</span></span> | <span data-ttu-id="496ae-971">Beschreibung</span><span class="sxs-lookup"><span data-stu-id="496ae-971">Description</span></span> |
|------------|-------------|
| `InvalidAttachmentId` | <span data-ttu-id="496ae-972">Der Bezeichner für die Anlage ist nicht vorhanden.</span><span class="sxs-lookup"><span data-stu-id="496ae-972">The attachment identifier does not exist.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="496ae-973">Anforderungen</span><span class="sxs-lookup"><span data-stu-id="496ae-973">Requirements</span></span>

|<span data-ttu-id="496ae-974">Anforderung</span><span class="sxs-lookup"><span data-stu-id="496ae-974">Requirement</span></span>| <span data-ttu-id="496ae-975">Wert</span><span class="sxs-lookup"><span data-stu-id="496ae-975">Value</span></span>|
|---|---|
|[<span data-ttu-id="496ae-976">Mindestversion des Postfachanforderungssatzes</span><span class="sxs-lookup"><span data-stu-id="496ae-976">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="496ae-977">1.1</span><span class="sxs-lookup"><span data-stu-id="496ae-977">1.1</span></span>|
|[<span data-ttu-id="496ae-978">Mindestberechtigungsstufe</span><span class="sxs-lookup"><span data-stu-id="496ae-978">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="496ae-979">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="496ae-979">ReadWriteItem</span></span>|
|[<span data-ttu-id="496ae-980">Zutreffender Outlook-Modus</span><span class="sxs-lookup"><span data-stu-id="496ae-980">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="496ae-981">Verfassen</span><span class="sxs-lookup"><span data-stu-id="496ae-981">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="496ae-982">Beispiel</span><span class="sxs-lookup"><span data-stu-id="496ae-982">Example</span></span>

<span data-ttu-id="496ae-983">Mit dem folgende Code wird eine Anlage mit dem Bezeichner "0" entfernt.</span><span class="sxs-lookup"><span data-stu-id="496ae-983">The following code removes an attachment with an identifier of '0'.</span></span>

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

####  <a name="saveasyncoptions-callback"></a><span data-ttu-id="496ae-984">saveAsync([options], callback)</span><span class="sxs-lookup"><span data-stu-id="496ae-984">saveAsync([options], callback)</span></span>

<span data-ttu-id="496ae-985">Speicher asynchron ein Element. .</span><span class="sxs-lookup"><span data-stu-id="496ae-985">Asynchronously saves an item.</span></span>

<span data-ttu-id="496ae-p167">Beim Aufrufen speichert diese Methode die aktuelle Nachricht als Entwurf und  gibt die Element-ID über die Callbackmethode zurück. In Outlook Web App oder Outlook im Onlinemodus wird das Element auf dem Server gespeichert. In Outlook im Cache-Modus wird das Element im lokalen Cache gespeichert.</span><span class="sxs-lookup"><span data-stu-id="496ae-p167">When invoked, this method saves the current message as a draft and returns the item id via the callback method. In Outlook Web App or Outlook in online mode, the item is saved to the server. In Outlook in cached mode, the item is saved to the local cache.</span></span>

> [!NOTE]
> <span data-ttu-id="496ae-989">Wenn Ihr Add-In ruft `saveAsync` für ein Element im Entwurfsmodus, um das Abrufen einer `itemId` um mit EWS oder REST-API verwenden, beachten Sie, dass wenn Outlook im Cache-Modus ist, es kann einige Zeit dauern, bevor das Element tatsächlich auf dem Server synchronisiert wird.</span><span class="sxs-lookup"><span data-stu-id="496ae-989">If your add-in calls `saveAsync` on an item in compose mode in order to get an `itemId` to use with EWS or the REST API, be aware that when Outlook is in cached mode, it may take some time before the item is actually synced to the server.</span></span> <span data-ttu-id="496ae-990">Bis das Element mit synchronisiert ist, die `itemId` wird ein Fehler zurückgegeben.</span><span class="sxs-lookup"><span data-stu-id="496ae-990">Until the item is synced, using the `itemId` will return an error.</span></span>

<span data-ttu-id="496ae-p169">Da Termine keinen Entwurfsstatus haben, wird das Element bei Aufruf von `saveAsync` für einen Termin im Verfassenmodus als normaler Termin im Kalender des Benutzers gespeichert. Für neue Termine, die noch nicht gespeichert wurden, wird keine Einladung gesendet. Beim Speichern eines vorhandenen Termins wird eine Aktualisierung an die hinzugefügten oder entfernten Teilnehmer gesendet.</span><span class="sxs-lookup"><span data-stu-id="496ae-p169">Since appointments have no draft state, if `saveAsync` is called on an appointment in compose mode, the item will be saved as a normal appointment on the user's calendar. For new appointments that have not been saved before, no invitation will be sent. Saving an existing appointment will send an update to added or removed attendees.</span></span>

> [!NOTE]
> <span data-ttu-id="496ae-994">Die folgenden Clients haben unterschiedlichem Verhalten für `saveAsync` Termine im Verfassenmodus:</span><span class="sxs-lookup"><span data-stu-id="496ae-994">The following clients have different behavior for `saveAsync` on appointments in compose mode:</span></span>
>
> - <span data-ttu-id="496ae-995">Outlook für Mac unterstützt keine `saveAsync` auf einer Besprechung im Verfassenmodus.</span><span class="sxs-lookup"><span data-stu-id="496ae-995">Mac Outlook does not support `saveAsync` on a meeting in compose mode.</span></span> <span data-ttu-id="496ae-996">Aufrufen von `saveAsync` auf einer Besprechung in Outlook für Mac wird ein Fehler zurückgegeben.</span><span class="sxs-lookup"><span data-stu-id="496ae-996">Calling `saveAsync` on a meeting in Mac Outlook will return an error.</span></span>
> - <span data-ttu-id="496ae-997">Outlook im Web immer sendet eine Einladung oder zu aktualisieren, wenn `saveAsync` für einen Termin aufgerufen wird, im Verfassenmodus.</span><span class="sxs-lookup"><span data-stu-id="496ae-997">Outlook on the web always sends an invitation or update when `saveAsync` is called on an appointment in compose mode.</span></span>

##### <a name="parameters"></a><span data-ttu-id="496ae-998">Parameter:</span><span class="sxs-lookup"><span data-stu-id="496ae-998">Parameters:</span></span>

|<span data-ttu-id="496ae-999">Name</span><span class="sxs-lookup"><span data-stu-id="496ae-999">Name</span></span>| <span data-ttu-id="496ae-1000">Typ</span><span class="sxs-lookup"><span data-stu-id="496ae-1000">Type</span></span>| <span data-ttu-id="496ae-1001">Attribute</span><span class="sxs-lookup"><span data-stu-id="496ae-1001">Attributes</span></span>| <span data-ttu-id="496ae-1002">Beschreibung</span><span class="sxs-lookup"><span data-stu-id="496ae-1002">Description</span></span>|
|---|---|---|---|
|`options`| <span data-ttu-id="496ae-1003">Objekt</span><span class="sxs-lookup"><span data-stu-id="496ae-1003">Object</span></span>| <span data-ttu-id="496ae-1004">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="496ae-1004">&lt;optional&gt;</span></span>|<span data-ttu-id="496ae-1005">Ein Objektliteral, das mindestens eine der folgenden Eigenschaften enthält.</span><span class="sxs-lookup"><span data-stu-id="496ae-1005">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`| <span data-ttu-id="496ae-1006">Object</span><span class="sxs-lookup"><span data-stu-id="496ae-1006">Object</span></span>| <span data-ttu-id="496ae-1007">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="496ae-1007">&lt;optional&gt;</span></span>|<span data-ttu-id="496ae-1008">Entwickler können ein Objekt bereitstellen, auf das sie in der Callbackmethode zugreifen möchten.</span><span class="sxs-lookup"><span data-stu-id="496ae-1008">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`| <span data-ttu-id="496ae-1009">Funktion</span><span class="sxs-lookup"><span data-stu-id="496ae-1009">function</span></span>||<span data-ttu-id="496ae-1010">Wenn die Methode abgeschlossen wird, wird die Funktion , die im `callback`-Parameter übergeben wird, mit einem einzelnen Parameter (`asyncResult`) aufgerufen, der ein [`AsyncResult`](/javascript/api/office/office.asyncresult)-Objekt darstellt.</span><span class="sxs-lookup"><span data-stu-id="496ae-1010">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span><br/><br/><span data-ttu-id="496ae-1011">Auf Erfolg, erfolgt die Element-ID der `asyncResult.value` Eigenschaft.</span><span class="sxs-lookup"><span data-stu-id="496ae-1011">On success, the item identifier is provided in the `asyncResult.value` property.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="496ae-1012">Anforderungen</span><span class="sxs-lookup"><span data-stu-id="496ae-1012">Requirements</span></span>

|<span data-ttu-id="496ae-1013">Anforderung</span><span class="sxs-lookup"><span data-stu-id="496ae-1013">Requirement</span></span>| <span data-ttu-id="496ae-1014">Wert</span><span class="sxs-lookup"><span data-stu-id="496ae-1014">Value</span></span>|
|---|---|
|[<span data-ttu-id="496ae-1015">Mindestversion des Postfachanforderungssatzes</span><span class="sxs-lookup"><span data-stu-id="496ae-1015">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="496ae-1016">1.3</span><span class="sxs-lookup"><span data-stu-id="496ae-1016">1.3</span></span>|
|[<span data-ttu-id="496ae-1017">Mindestberechtigungsstufe</span><span class="sxs-lookup"><span data-stu-id="496ae-1017">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="496ae-1018">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="496ae-1018">ReadWriteItem</span></span>|
|[<span data-ttu-id="496ae-1019">Zutreffender Outlook-Modus</span><span class="sxs-lookup"><span data-stu-id="496ae-1019">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="496ae-1020">Verfassen</span><span class="sxs-lookup"><span data-stu-id="496ae-1020">Compose</span></span>|

##### <a name="examples"></a><span data-ttu-id="496ae-1021">Beispiele</span><span class="sxs-lookup"><span data-stu-id="496ae-1021">Examples</span></span>

```
Office.context.mailbox.item.saveAsync(
  function callback(result) {
    // Process the result
  });
```

<span data-ttu-id="496ae-p171">Es folgt ein Beispiel für den `result`-Parameter, der an die Callbackfunktion übergeben wird. Die `value`-Eigenschaft enthält die Element-ID des Elements.</span><span class="sxs-lookup"><span data-stu-id="496ae-p171">The following is an example of the `result` parameter passed to the callback function. The `value` property contains the item ID of the item.</span></span>

```
{
  "value":"AAMkADI5...AAA=",
  "status":"succeeded"
}
```

####  <a name="setselecteddataasyncdata-options-callback"></a><span data-ttu-id="496ae-1024">setSelectedDataAsync(data, [options], callback)</span><span class="sxs-lookup"><span data-stu-id="496ae-1024">setSelectedDataAsync(data, [options], callback)</span></span>

<span data-ttu-id="496ae-1025">Fügt asynchron Daten in den Textkörper oder Betreff einer Nachricht ein.</span><span class="sxs-lookup"><span data-stu-id="496ae-1025">Asynchronously inserts data into the body or subject of a message.</span></span>

<span data-ttu-id="496ae-p172">Die `setSelectedDataAsync`-Methode fügt die angegebene Zeichenfolge an der Cursorposition im Betreff oder im Textkörper ein, oder, falls im Editor Text ausgewählt ist, ersetzt den markierten Text. Wenn sich der Cursor nicht im Textkörper oder im Betreffsfeld befindet, wird ein Fehler zurückgegeben. Nach dem Einfügen wird der Cursor am Ende der eingefügten Inhalte platziert.</span><span class="sxs-lookup"><span data-stu-id="496ae-p172">The `setSelectedDataAsync` method inserts the specified string at the cursor location in the subject or body of the item, or, if text is selected in the editor, it replaces the selected text. If the cursor is not in the body or subject field, an error is returned. After insertion, the cursor is placed at the end of the inserted content.</span></span>

##### <a name="parameters"></a><span data-ttu-id="496ae-1029">Parameter:</span><span class="sxs-lookup"><span data-stu-id="496ae-1029">Parameters:</span></span>

|<span data-ttu-id="496ae-1030">Name</span><span class="sxs-lookup"><span data-stu-id="496ae-1030">Name</span></span>| <span data-ttu-id="496ae-1031">Typ</span><span class="sxs-lookup"><span data-stu-id="496ae-1031">Type</span></span>| <span data-ttu-id="496ae-1032">Attribute</span><span class="sxs-lookup"><span data-stu-id="496ae-1032">Attributes</span></span>| <span data-ttu-id="496ae-1033">Beschreibung</span><span class="sxs-lookup"><span data-stu-id="496ae-1033">Description</span></span>|
|---|---|---|---|
|`data`| <span data-ttu-id="496ae-1034">Zeichenfolge</span><span class="sxs-lookup"><span data-stu-id="496ae-1034">String</span></span>||<span data-ttu-id="496ae-p173">Die einzufügenden Daten. Daten dürfen 1.000.000 Zeichen nicht überschreiten. Werden mehr als 1.000.000 Zeichen übergeben, wird eine `ArgumentOutOfRange`-Ausnahme ausgelöst.</span><span class="sxs-lookup"><span data-stu-id="496ae-p173">The data to be inserted. Data is not to exceed 1,000,000 characters. If more than 1,000,000 characters are passed in, an `ArgumentOutOfRange` exception is thrown.</span></span>|
|`options`| <span data-ttu-id="496ae-1038">Object</span><span class="sxs-lookup"><span data-stu-id="496ae-1038">Object</span></span>| <span data-ttu-id="496ae-1039">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="496ae-1039">&lt;optional&gt;</span></span>|<span data-ttu-id="496ae-1040">Ein Objektliteral, das mindestens eine der folgenden Eigenschaften enthält.</span><span class="sxs-lookup"><span data-stu-id="496ae-1040">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`| <span data-ttu-id="496ae-1041">Object</span><span class="sxs-lookup"><span data-stu-id="496ae-1041">Object</span></span>| <span data-ttu-id="496ae-1042">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="496ae-1042">&lt;optional&gt;</span></span>|<span data-ttu-id="496ae-1043">Entwickler können ein Objekt bereitstellen, auf das sie in der Rückrufmethode zugreifen möchten.</span><span class="sxs-lookup"><span data-stu-id="496ae-1043">Developers can provide any object they wish to access in the callback method.</span></span>|
|`options.coercionType`| [<span data-ttu-id="496ae-1044">Office.CoercionType</span><span class="sxs-lookup"><span data-stu-id="496ae-1044">Office.CoercionType</span></span>](office.md#coerciontype-string)| <span data-ttu-id="496ae-1045">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="496ae-1045">&lt;optional&gt;</span></span>|<span data-ttu-id="496ae-p174">Wenn `text`, wird das aktuelle Format in Outlook Web App und Outlook angewendet. Wenn das Feld ein HTML-Editor ist, werden nur die Textdaten eingefügt, selbst wenn es sich bei den Daten um HTML-Daten handelt.</span><span class="sxs-lookup"><span data-stu-id="496ae-p174">If `text`, the current style is applied in Outlook Web App and Outlook. If the field is an HTML editor, only the text data is inserted, even if the data is HTML.</span></span><br/><br/><span data-ttu-id="496ae-p175">Wenn `html` gesetzt ist und das Feld HTML unterstützt (der Betreff jedoch nicht), wird die aktuelle Formatvorlage in Outlook Web App angewendet und die Standardformatvorlage in Outlook. Ist das Feld ein Textfeld, wird ein Fehler des Typs `InvalidDataFormat` zurückgegeben.</span><span class="sxs-lookup"><span data-stu-id="496ae-p175">If `html` and the field supports HTML (the subject doesn't), the current style is applied in Outlook Web App and the default style is applied in Outlook. If the field is a text field, an `InvalidDataFormat` error is returned.</span></span><br/><br/><span data-ttu-id="496ae-1050">Wenn `coercionType` nicht festgelegt wird, hängt das Ergebnis vom Feld ab: Wenn das Feld HTML ist, wird HTML verwendet. Wenn das Feld Text ist, wird Nur-Text verwendet.</span><span class="sxs-lookup"><span data-stu-id="496ae-1050">If `coercionType` is not set, the result depends on the field: if the field is HTML then HTML is used; if the field is text, then plain text is used.</span></span>|
|`callback`| <span data-ttu-id="496ae-1051">function</span><span class="sxs-lookup"><span data-stu-id="496ae-1051">function</span></span>||<span data-ttu-id="496ae-1052">Wenn die Methode abgeschlossen wird, wird die Funktion , die im `callback`-Parameter übergeben wird, mit einem einzelnen Parameter (`asyncResult`) aufgerufen, der ein [`AsyncResult`](/javascript/api/office/office.asyncresult)-Objekt darstellt.</span><span class="sxs-lookup"><span data-stu-id="496ae-1052">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="496ae-1053">Anforderungen</span><span class="sxs-lookup"><span data-stu-id="496ae-1053">Requirements</span></span>

|<span data-ttu-id="496ae-1054">Anforderung</span><span class="sxs-lookup"><span data-stu-id="496ae-1054">Requirement</span></span>| <span data-ttu-id="496ae-1055">Wert</span><span class="sxs-lookup"><span data-stu-id="496ae-1055">Value</span></span>|
|---|---|
|[<span data-ttu-id="496ae-1056">Mindestversion des Postfachanforderungssatzes</span><span class="sxs-lookup"><span data-stu-id="496ae-1056">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="496ae-1057">1.2</span><span class="sxs-lookup"><span data-stu-id="496ae-1057">1.2</span></span>|
|[<span data-ttu-id="496ae-1058">Mindestberechtigungsstufe</span><span class="sxs-lookup"><span data-stu-id="496ae-1058">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="496ae-1059">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="496ae-1059">ReadWriteItem</span></span>|
|[<span data-ttu-id="496ae-1060">Zutreffender Outlook-Modus</span><span class="sxs-lookup"><span data-stu-id="496ae-1060">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="496ae-1061">Verfassen</span><span class="sxs-lookup"><span data-stu-id="496ae-1061">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="496ae-1062">Beispiel</span><span class="sxs-lookup"><span data-stu-id="496ae-1062">Example</span></span>

```
Office.context.mailbox.item.setSelectedDataAsync("Hello World!");
Office.context.mailbox.item.setSelectedDataAsync("<b>Hello World!</b>", { coercionType : "html" });
```