
# <a name="item"></a><span data-ttu-id="739f5-101">item</span><span class="sxs-lookup"><span data-stu-id="739f5-101">item</span></span>

### <span data-ttu-id="739f5-p101">[Office](office.md)[.context](office.context.md)[.mailbox](office.context.mailbox.md). item</span><span class="sxs-lookup"><span data-stu-id="739f5-p101">[Office](office.md)[.context](office.context.md)[.mailbox](office.context.mailbox.md). item</span></span>

<span data-ttu-id="739f5-p102">Der `item`-Namespace wird für den Zugriff auf die aktuell ausgewählte Nachricht, Besprechungsanfrage oder den aktuell ausgewählten Termin verwendet. Sie können den Typ von `item` mithilfe der [itemType](#itemtype-officemailboxenumsitemtypejavascriptapioutlook11officemailboxenumsitemtype)-Eigenschaft bestimmen.</span><span class="sxs-lookup"><span data-stu-id="739f5-p102">The `item` namespace is used to access the currently selected message, meeting request, or appointment. You can determine the type of the `item` by using the [itemType](#itemtype-officemailboxenumsitemtypejavascriptapioutlook11officemailboxenumsitemtype) property.</span></span>

##### <a name="requirements"></a><span data-ttu-id="739f5-106">Anforderungen</span><span class="sxs-lookup"><span data-stu-id="739f5-106">Requirements</span></span>

|<span data-ttu-id="739f5-107">Anforderung</span><span class="sxs-lookup"><span data-stu-id="739f5-107">Requirement</span></span>| <span data-ttu-id="739f5-108">Wert</span><span class="sxs-lookup"><span data-stu-id="739f5-108">Value</span></span>|
|---|---|
|[<span data-ttu-id="739f5-109">Mindestversion des Postfachanforderungssatzes</span><span class="sxs-lookup"><span data-stu-id="739f5-109">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="739f5-110">1.0</span><span class="sxs-lookup"><span data-stu-id="739f5-110">1.0</span></span>|
|[<span data-ttu-id="739f5-111">Mindestberechtigungsstufe</span><span class="sxs-lookup"><span data-stu-id="739f5-111">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="739f5-112">Eingeschränkt</span><span class="sxs-lookup"><span data-stu-id="739f5-112">Restricted</span></span>|
|[<span data-ttu-id="739f5-113">Zutreffender Outlook-Modus</span><span class="sxs-lookup"><span data-stu-id="739f5-113">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="739f5-114">Verfassen oder Lesen</span><span class="sxs-lookup"><span data-stu-id="739f5-114">Compose or read</span></span>|

### <a name="example"></a><span data-ttu-id="739f5-115">Beispiel</span><span class="sxs-lookup"><span data-stu-id="739f5-115">Example</span></span>

<span data-ttu-id="739f5-116">Im folgenden JavaScript-Codebeispiel wird der Zugriff auf die `subject`-Eigenschaft des aktuellen Elements in Outlook veranschaulicht.</span><span class="sxs-lookup"><span data-stu-id="739f5-116">The following JavaScript code example shows how to access the `subject` property of the current item in Outlook.</span></span>

```JavaScript
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

### <a name="members"></a><span data-ttu-id="739f5-117">Elemente</span><span class="sxs-lookup"><span data-stu-id="739f5-117">Members</span></span>

#### <a name="attachments-arrayattachmentdetailsjavascriptapioutlook11officeattachmentdetails"></a><span data-ttu-id="739f5-118">attachments :Array.<[AttachmentDetails](/javascript/api/outlook_1_1/office.attachmentdetails)></span><span class="sxs-lookup"><span data-stu-id="739f5-118">attachments :Array.<[AttachmentDetails](/javascript/api/outlook_1_1/office.attachmentdetails)></span></span>

<span data-ttu-id="739f5-p103">Ruft ein Array mit Anlagen für das Element ab. Nur Lesemodus.</span><span class="sxs-lookup"><span data-stu-id="739f5-p103">Gets an array of attachments for the item. Read mode only.</span></span>

> [!NOTE]
> <span data-ttu-id="739f5-121">Bestimmte Dateitypen werden aufgrund von Sicherheitsproblemen von Outlook blockiert und daher nicht zurückgegeben werden.</span><span class="sxs-lookup"><span data-stu-id="739f5-121">Certain types of files are blocked by Outlook due to potential security issues and are therefore not returned.</span></span> <span data-ttu-id="739f5-122">Weitere Informationen finden Sie unter [Blockierte Anlagen in Outlook](https://support.office.com/article/Blocked-attachments-in-Outlook-434752E1-02D3-4E90-9124-8B81E49A8519).</span><span class="sxs-lookup"><span data-stu-id="739f5-122">For more information, see [Blocked attachments in Outlook](https://support.office.com/article/Blocked-attachments-in-Outlook-434752E1-02D3-4E90-9124-8B81E49A8519).</span></span>

##### <a name="type"></a><span data-ttu-id="739f5-123">Typ:</span><span class="sxs-lookup"><span data-stu-id="739f5-123">Type:</span></span>

*   <span data-ttu-id="739f5-124">Array.<[AttachmentDetails](/javascript/api/outlook_1_1/office.attachmentdetails)></span><span class="sxs-lookup"><span data-stu-id="739f5-124">Array.<[AttachmentDetails](/javascript/api/outlook_1_1/office.attachmentdetails)></span></span>

##### <a name="requirements"></a><span data-ttu-id="739f5-125">Anforderungen</span><span class="sxs-lookup"><span data-stu-id="739f5-125">Requirements</span></span>

|<span data-ttu-id="739f5-126">Anforderung</span><span class="sxs-lookup"><span data-stu-id="739f5-126">Requirement</span></span>| <span data-ttu-id="739f5-127">Wert</span><span class="sxs-lookup"><span data-stu-id="739f5-127">Value</span></span>|
|---|---|
|[<span data-ttu-id="739f5-128">Mindestversion des Postfachanforderungssatzes</span><span class="sxs-lookup"><span data-stu-id="739f5-128">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="739f5-129">1.0</span><span class="sxs-lookup"><span data-stu-id="739f5-129">1.0</span></span>|
|[<span data-ttu-id="739f5-130">Mindestberechtigungsstufe</span><span class="sxs-lookup"><span data-stu-id="739f5-130">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="739f5-131">ReadItem</span><span class="sxs-lookup"><span data-stu-id="739f5-131">ReadItem</span></span>|
|[<span data-ttu-id="739f5-132">Zutreffender Outlook-Modus</span><span class="sxs-lookup"><span data-stu-id="739f5-132">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="739f5-133">Lesen</span><span class="sxs-lookup"><span data-stu-id="739f5-133">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="739f5-134">Beispiel</span><span class="sxs-lookup"><span data-stu-id="739f5-134">Example</span></span>

<span data-ttu-id="739f5-135">Mit dem folgende Code wird eine HTML-Zeichenfolge mit Details aller Anlagen für das aktuelle Element erstellt.</span><span class="sxs-lookup"><span data-stu-id="739f5-135">The following code builds an HTML string with details of all attachments on the current item.</span></span>

```JavaScript
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

####  <a name="bcc-recipientsjavascriptapioutlook11officerecipients"></a><span data-ttu-id="739f5-136">bcc :[Recipients](/javascript/api/outlook_1_1/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="739f5-136">bcc :[Recipients](/javascript/api/outlook_1_1/office.recipients)</span></span>

<span data-ttu-id="739f5-137">Ruft ein Objekt, das Methoden zum Abrufen oder aktualisieren die Empfänger in der Zeile Bcc (blind Carbon Copy, Blindkopie) einer Nachricht bereitstellt.</span><span class="sxs-lookup"><span data-stu-id="739f5-137">Gets an object that provides methods to get or update the recipients on the Bcc (blind carbon copy) line of a message.</span></span> <span data-ttu-id="739f5-138">Nur Verfassenmodus.</span><span class="sxs-lookup"><span data-stu-id="739f5-138">Compose mode only.</span></span>

##### <a name="type"></a><span data-ttu-id="739f5-139">Typ:</span><span class="sxs-lookup"><span data-stu-id="739f5-139">Type:</span></span>

*   [<span data-ttu-id="739f5-140">Recipients</span><span class="sxs-lookup"><span data-stu-id="739f5-140">Recipients</span></span>](/javascript/api/outlook_1_1/office.recipients)

##### <a name="requirements"></a><span data-ttu-id="739f5-141">Anforderungen</span><span class="sxs-lookup"><span data-stu-id="739f5-141">Requirements</span></span>

|<span data-ttu-id="739f5-142">Anforderung</span><span class="sxs-lookup"><span data-stu-id="739f5-142">Requirement</span></span>| <span data-ttu-id="739f5-143">Wert</span><span class="sxs-lookup"><span data-stu-id="739f5-143">Value</span></span>|
|---|---|
|[<span data-ttu-id="739f5-144">Mindestversion des Postfachanforderungssatzes</span><span class="sxs-lookup"><span data-stu-id="739f5-144">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="739f5-145">1.1</span><span class="sxs-lookup"><span data-stu-id="739f5-145">1.1</span></span>|
|[<span data-ttu-id="739f5-146">Mindestberechtigungsstufe</span><span class="sxs-lookup"><span data-stu-id="739f5-146">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="739f5-147">ReadItem</span><span class="sxs-lookup"><span data-stu-id="739f5-147">ReadItem</span></span>|
|[<span data-ttu-id="739f5-148">Zutreffender Outlook-Modus</span><span class="sxs-lookup"><span data-stu-id="739f5-148">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="739f5-149">Verfassen</span><span class="sxs-lookup"><span data-stu-id="739f5-149">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="739f5-150">Beispiel</span><span class="sxs-lookup"><span data-stu-id="739f5-150">Example</span></span>

```JavaScript
Office.context.mailbox.item.bcc.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.bcc.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.bcc.getAsync(callback);

function callback(asyncResult) {
  var arrayOfBccRecipients = asyncResult.value;
}
```

####  <a name="body-bodyjavascriptapioutlook11officebody"></a><span data-ttu-id="739f5-151">body :[Body](/javascript/api/outlook_1_1/office.body)</span><span class="sxs-lookup"><span data-stu-id="739f5-151">body :[Body](/javascript/api/outlook_1_1/office.body)</span></span>

<span data-ttu-id="739f5-152">Ruft ein Objekt ab, das Methoden zum Bearbeiten des Textkörpers eines Elements bereitstellt.</span><span class="sxs-lookup"><span data-stu-id="739f5-152">Gets an object that provides methods for manipulating the body of an item.</span></span>

##### <a name="type"></a><span data-ttu-id="739f5-153">Typ:</span><span class="sxs-lookup"><span data-stu-id="739f5-153">Type:</span></span>

*   [<span data-ttu-id="739f5-154">Body</span><span class="sxs-lookup"><span data-stu-id="739f5-154">Body</span></span>](/javascript/api/outlook_1_1/office.body)

##### <a name="requirements"></a><span data-ttu-id="739f5-155">Anforderungen</span><span class="sxs-lookup"><span data-stu-id="739f5-155">Requirements</span></span>

|<span data-ttu-id="739f5-156">Anforderung</span><span class="sxs-lookup"><span data-stu-id="739f5-156">Requirement</span></span>| <span data-ttu-id="739f5-157">Wert</span><span class="sxs-lookup"><span data-stu-id="739f5-157">Value</span></span>|
|---|---|
|[<span data-ttu-id="739f5-158">Mindestversion des Postfachanforderungssatzes</span><span class="sxs-lookup"><span data-stu-id="739f5-158">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="739f5-159">1.1</span><span class="sxs-lookup"><span data-stu-id="739f5-159">1.1</span></span>|
|[<span data-ttu-id="739f5-160">Mindestberechtigungsstufe</span><span class="sxs-lookup"><span data-stu-id="739f5-160">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="739f5-161">ReadItem</span><span class="sxs-lookup"><span data-stu-id="739f5-161">ReadItem</span></span>|
|[<span data-ttu-id="739f5-162">Zutreffender Outlook-Modus</span><span class="sxs-lookup"><span data-stu-id="739f5-162">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="739f5-163">Verfassen oder Lesen</span><span class="sxs-lookup"><span data-stu-id="739f5-163">Compose or read</span></span>|

####  <a name="cc-arrayemailaddressdetailsjavascriptapioutlook11officeemailaddressdetailsrecipientsjavascriptapioutlook11officerecipients"></a><span data-ttu-id="739f5-164">cc: Array. <[EmailAddressDetails](/javascript/api/outlook_1_1/office.emailaddressdetails)>|[Empfänger](/javascript/api/outlook_1_1/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="739f5-164">cc :Array.<[EmailAddressDetails](/javascript/api/outlook_1_1/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook_1_1/office.recipients)</span></span>

<span data-ttu-id="739f5-165">Ermöglicht den Zugriff auf die (Carbon Copy, Blindkopie) Cc-Empfänger einer Nachricht.</span><span class="sxs-lookup"><span data-stu-id="739f5-165">Provides access to the Cc (carbon copy) recipients of a message.</span></span> <span data-ttu-id="739f5-166">Der Typ des Objekts und die Zugriffsebene, hängt von den Modus des aktuellen Elements ab.</span><span class="sxs-lookup"><span data-stu-id="739f5-166">The type of object and level of access depends on the mode of the current item.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="739f5-167">Lesemodus</span><span class="sxs-lookup"><span data-stu-id="739f5-167">Read mode</span></span>

<span data-ttu-id="739f5-p107">Die `cc`-Eigenschaft gibt ein Array mit einem `EmailAddressDetails`-Objekt für jeden Empfänger in der **Cc**-Zeile der Nachricht zurück. Die Auflistung ist auf höchstens 100 Elemente beschränkt.</span><span class="sxs-lookup"><span data-stu-id="739f5-p107">The `cc` property returns an array that contains an `EmailAddressDetails` object for each recipient listed on the **Cc** line of the message. The collection is limited to a maximum of 100 members.</span></span>

##### <a name="compose-mode"></a><span data-ttu-id="739f5-170">Verfassenmodus</span><span class="sxs-lookup"><span data-stu-id="739f5-170">Compose mode</span></span>

<span data-ttu-id="739f5-171">Die `cc` -Eigenschaft gibt eine `Recipients` -Objekt, das Methoden zum Abrufen oder aktualisieren die Empfänger in der Zeile **Cc** der Nachricht bereitstellt.</span><span class="sxs-lookup"><span data-stu-id="739f5-171">The `cc` property returns a `Recipients` object that provides methods to get or update the recipients on the **Cc** line of the message.</span></span>

##### <a name="type"></a><span data-ttu-id="739f5-172">Typ:</span><span class="sxs-lookup"><span data-stu-id="739f5-172">Type:</span></span>

*   <span data-ttu-id="739f5-173">Array.<[EmailAddressDetails](/javascript/api/outlook_1_1/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook_1_1/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="739f5-173">Array.<[EmailAddressDetails](/javascript/api/outlook_1_1/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook_1_1/office.recipients)</span></span>

##### <a name="requirements"></a><span data-ttu-id="739f5-174">Anforderungen</span><span class="sxs-lookup"><span data-stu-id="739f5-174">Requirements</span></span>

|<span data-ttu-id="739f5-175">Anforderung</span><span class="sxs-lookup"><span data-stu-id="739f5-175">Requirement</span></span>| <span data-ttu-id="739f5-176">Wert</span><span class="sxs-lookup"><span data-stu-id="739f5-176">Value</span></span>|
|---|---|
|[<span data-ttu-id="739f5-177">Mindestversion des Postfachanforderungssatzes</span><span class="sxs-lookup"><span data-stu-id="739f5-177">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="739f5-178">1.0</span><span class="sxs-lookup"><span data-stu-id="739f5-178">1.0</span></span>|
|[<span data-ttu-id="739f5-179">Mindestberechtigungsstufe</span><span class="sxs-lookup"><span data-stu-id="739f5-179">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="739f5-180">ReadItem</span><span class="sxs-lookup"><span data-stu-id="739f5-180">ReadItem</span></span>|
|[<span data-ttu-id="739f5-181">Zutreffender Outlook-Modus</span><span class="sxs-lookup"><span data-stu-id="739f5-181">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="739f5-182">Verfassen oder Lesen</span><span class="sxs-lookup"><span data-stu-id="739f5-182">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="739f5-183">Beispiel</span><span class="sxs-lookup"><span data-stu-id="739f5-183">Example</span></span>

```JavaScript
Office.context.mailbox.item.cc.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.cc.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.cc.getAsync(callback);

function callback(asyncResult) {
  var arrayOfCcRecipients = asyncResult.value;
}
```

####  <a name="nullable-conversationid-string"></a><span data-ttu-id="739f5-184">(nullable) conversationId :String</span><span class="sxs-lookup"><span data-stu-id="739f5-184">(nullable) conversationId :String</span></span>

<span data-ttu-id="739f5-185">Ruft einen Bezeichner für die E-Mail-Unterhaltung ab, in der eine bestimmte Nachricht enthalten ist.</span><span class="sxs-lookup"><span data-stu-id="739f5-185">Gets an identifier for the email conversation that contains a particular message.</span></span>

<span data-ttu-id="739f5-p108">Sie können für diese Eigenschaft eine ganze Zahl abrufen, wenn Ihre Mail-App in Formularen zum Lesen oder Antworten in Formularen zum Verfassen aktiviert wird. Wenn der Benutzer den Betreff der Antwortnachricht ändert, ändert sich beim Versenden die Konversations-ID für die entsprechende Nachricht, und der Wert, den Sie vorher bezogen haben, trifft nicht länger zu.</span><span class="sxs-lookup"><span data-stu-id="739f5-p108">You can get an integer for this property if your mail app is activated in read forms or responses in compose forms. If subsequently the user changes the subject of the reply message, upon sending the reply, the conversation ID for that message will change and that value you obtained earlier will no longer apply.</span></span>

<span data-ttu-id="739f5-p109">Sie erhalten in einem Formular zum Verfassen für diese Eigenschaft für ein neues Element null. Wenn der Benutzer einen Betreff festlegt und das Element speichert, gibt die `conversationId`-Eigenschaft einen Wert zurück.</span><span class="sxs-lookup"><span data-stu-id="739f5-p109">You get null for this property for a new item in a compose form. If the user sets a subject and saves the item, the `conversationId` property will return a value.</span></span>

##### <a name="type"></a><span data-ttu-id="739f5-190">Typ:</span><span class="sxs-lookup"><span data-stu-id="739f5-190">Type:</span></span>

*   <span data-ttu-id="739f5-191">String</span><span class="sxs-lookup"><span data-stu-id="739f5-191">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="739f5-192">Anforderungen</span><span class="sxs-lookup"><span data-stu-id="739f5-192">Requirements</span></span>

|<span data-ttu-id="739f5-193">Anforderung</span><span class="sxs-lookup"><span data-stu-id="739f5-193">Requirement</span></span>| <span data-ttu-id="739f5-194">Wert</span><span class="sxs-lookup"><span data-stu-id="739f5-194">Value</span></span>|
|---|---|
|[<span data-ttu-id="739f5-195">Mindestversion des Postfachanforderungssatzes</span><span class="sxs-lookup"><span data-stu-id="739f5-195">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="739f5-196">1.0</span><span class="sxs-lookup"><span data-stu-id="739f5-196">1.0</span></span>|
|[<span data-ttu-id="739f5-197">Mindestberechtigungsstufe</span><span class="sxs-lookup"><span data-stu-id="739f5-197">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="739f5-198">ReadItem</span><span class="sxs-lookup"><span data-stu-id="739f5-198">ReadItem</span></span>|
|[<span data-ttu-id="739f5-199">Zutreffender Outlook-Modus</span><span class="sxs-lookup"><span data-stu-id="739f5-199">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="739f5-200">Verfassen oder Lesen</span><span class="sxs-lookup"><span data-stu-id="739f5-200">Compose or read</span></span>|

#### <a name="datetimecreated-date"></a><span data-ttu-id="739f5-201">dateTimeCreated :Date</span><span class="sxs-lookup"><span data-stu-id="739f5-201">dateTimeCreated :Date</span></span>

<span data-ttu-id="739f5-p110">Ruft das Datum und die Uhrzeit der Erstellung eines Elements ab. Nur Lesemodus.</span><span class="sxs-lookup"><span data-stu-id="739f5-p110">Gets the date and time that an item was created. Read mode only.</span></span>

##### <a name="type"></a><span data-ttu-id="739f5-204">Typ:</span><span class="sxs-lookup"><span data-stu-id="739f5-204">Type:</span></span>

*   <span data-ttu-id="739f5-205">Datum</span><span class="sxs-lookup"><span data-stu-id="739f5-205">Date</span></span>

##### <a name="requirements"></a><span data-ttu-id="739f5-206">Anforderungen</span><span class="sxs-lookup"><span data-stu-id="739f5-206">Requirements</span></span>

|<span data-ttu-id="739f5-207">Anforderung</span><span class="sxs-lookup"><span data-stu-id="739f5-207">Requirement</span></span>| <span data-ttu-id="739f5-208">Wert</span><span class="sxs-lookup"><span data-stu-id="739f5-208">Value</span></span>|
|---|---|
|[<span data-ttu-id="739f5-209">Mindestversion des Postfachanforderungssatzes</span><span class="sxs-lookup"><span data-stu-id="739f5-209">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="739f5-210">1.0</span><span class="sxs-lookup"><span data-stu-id="739f5-210">1.0</span></span>|
|[<span data-ttu-id="739f5-211">Mindestberechtigungsstufe</span><span class="sxs-lookup"><span data-stu-id="739f5-211">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="739f5-212">ReadItem</span><span class="sxs-lookup"><span data-stu-id="739f5-212">ReadItem</span></span>|
|[<span data-ttu-id="739f5-213">Zutreffender Outlook-Modus</span><span class="sxs-lookup"><span data-stu-id="739f5-213">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="739f5-214">Lesen</span><span class="sxs-lookup"><span data-stu-id="739f5-214">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="739f5-215">Beispiel</span><span class="sxs-lookup"><span data-stu-id="739f5-215">Example</span></span>

```JavaScript
var created = Office.context.mailbox.item.dateTimeCreated;
```

#### <a name="datetimemodified-date"></a><span data-ttu-id="739f5-216">dateTimeModified :Date</span><span class="sxs-lookup"><span data-stu-id="739f5-216">dateTimeModified :Date</span></span>

<span data-ttu-id="739f5-p111">Ruft das Datum und die Uhrzeit der letzten Änderung eines Elements ab. Nur Lesemodus.</span><span class="sxs-lookup"><span data-stu-id="739f5-p111">Gets the date and time that an item was last modified. Read mode only.</span></span>

> [!NOTE]
> <span data-ttu-id="739f5-219">Dieser Member wird in Outlook für iOS oder Outlook für Android nicht unterstützt.</span><span class="sxs-lookup"><span data-stu-id="739f5-219">This member is not supported in Outlook for iOS or Outlook for Android.</span></span>

##### <a name="type"></a><span data-ttu-id="739f5-220">Typ:</span><span class="sxs-lookup"><span data-stu-id="739f5-220">Type:</span></span>

*   <span data-ttu-id="739f5-221">Datum</span><span class="sxs-lookup"><span data-stu-id="739f5-221">Date</span></span>

##### <a name="requirements"></a><span data-ttu-id="739f5-222">Anforderungen</span><span class="sxs-lookup"><span data-stu-id="739f5-222">Requirements</span></span>

|<span data-ttu-id="739f5-223">Anforderung</span><span class="sxs-lookup"><span data-stu-id="739f5-223">Requirement</span></span>| <span data-ttu-id="739f5-224">Wert</span><span class="sxs-lookup"><span data-stu-id="739f5-224">Value</span></span>|
|---|---|
|[<span data-ttu-id="739f5-225">Mindestversion des Postfachanforderungssatzes</span><span class="sxs-lookup"><span data-stu-id="739f5-225">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="739f5-226">1.0</span><span class="sxs-lookup"><span data-stu-id="739f5-226">1.0</span></span>|
|[<span data-ttu-id="739f5-227">Mindestberechtigungsstufe</span><span class="sxs-lookup"><span data-stu-id="739f5-227">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="739f5-228">ReadItem</span><span class="sxs-lookup"><span data-stu-id="739f5-228">ReadItem</span></span>|
|[<span data-ttu-id="739f5-229">Zutreffender Outlook-Modus</span><span class="sxs-lookup"><span data-stu-id="739f5-229">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="739f5-230">Lesen</span><span class="sxs-lookup"><span data-stu-id="739f5-230">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="739f5-231">Beispiel</span><span class="sxs-lookup"><span data-stu-id="739f5-231">Example</span></span>

```JavaScript
var modified = Office.context.mailbox.item.dateTimeModified;
```

####  <a name="end-datetimejavascriptapioutlook11officetime"></a><span data-ttu-id="739f5-232">end :Date|[Time](/javascript/api/outlook_1_1/office.time)</span><span class="sxs-lookup"><span data-stu-id="739f5-232">end :Date|[Time](/javascript/api/outlook_1_1/office.time)</span></span>

<span data-ttu-id="739f5-233">Ruft Datum und Zeit für das Ende des Termins ab oder legt diese fest.</span><span class="sxs-lookup"><span data-stu-id="739f5-233">Gets or sets the date and time that the appointment is to end.</span></span>

<span data-ttu-id="739f5-p112">Die `end`-Eigenschaft wird als Datums- und Uhrzeitwert in UTC ausgedrückt. Sie können die [`convertToLocalClientTime`](office.context.mailbox.md#converttolocalclienttimetimevalue--localclienttimejavascriptapioutlook11officelocalclienttime)-Methode verwenden, um den Endeigenschaftswert in den lokalen Uhrzeit- und Datumswert des Clients umzuwandeln.</span><span class="sxs-lookup"><span data-stu-id="739f5-p112">The `end` property is expressed as a Coordinated Universal Time (UTC) date and time value. You can use the [`convertToLocalClientTime`](office.context.mailbox.md#converttolocalclienttimetimevalue--localclienttimejavascriptapioutlook11officelocalclienttime) method to convert the end property value to the client’s local date and time.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="739f5-236">Lesemodus</span><span class="sxs-lookup"><span data-stu-id="739f5-236">Read mode</span></span>

<span data-ttu-id="739f5-237">Die `end`-Eigenschaft gibt ein `Date`-Objekt zurück.</span><span class="sxs-lookup"><span data-stu-id="739f5-237">The `end` property returns a `Date` object.</span></span>

##### <a name="compose-mode"></a><span data-ttu-id="739f5-238">Verfassenmodus</span><span class="sxs-lookup"><span data-stu-id="739f5-238">Compose mode</span></span>

<span data-ttu-id="739f5-239">Die `end`-Eigenschaft gibt ein `Time`-Objekt zurück.</span><span class="sxs-lookup"><span data-stu-id="739f5-239">The `end` property returns a `Time` object.</span></span>

<span data-ttu-id="739f5-240">Wenn Sie die [`Time.setAsync`](/javascript/api/outlook_1_1/office.time#setasync-datetime--options--callback-)-Methode verwenden, um die Endzeit im Verfassenmodus festzulegen, sollten Sie die [`convertToUtcClientTime`](office.context.mailbox.md#converttoutcclienttimeinput--date)-Methode verwenden, um die Ortszeit auf dem Client für den Server in UTC umzuwandeln.</span><span class="sxs-lookup"><span data-stu-id="739f5-240">When you use the [`Time.setAsync`](/javascript/api/outlook_1_1/office.time#setasync-datetime--options--callback-) method to set the end time, you should use the [`convertToUtcClientTime`](office.context.mailbox.md#converttoutcclienttimeinput--date) method to convert the local time on the client to UTC for the server.</span></span>

##### <a name="type"></a><span data-ttu-id="739f5-241">Typ:</span><span class="sxs-lookup"><span data-stu-id="739f5-241">Type:</span></span>

*   <span data-ttu-id="739f5-242">Date | [Time](/javascript/api/outlook_1_1/office.time)</span><span class="sxs-lookup"><span data-stu-id="739f5-242">Date | [Time](/javascript/api/outlook_1_1/office.time)</span></span>

##### <a name="requirements"></a><span data-ttu-id="739f5-243">Anforderungen</span><span class="sxs-lookup"><span data-stu-id="739f5-243">Requirements</span></span>

|<span data-ttu-id="739f5-244">Anforderung</span><span class="sxs-lookup"><span data-stu-id="739f5-244">Requirement</span></span>| <span data-ttu-id="739f5-245">Wert</span><span class="sxs-lookup"><span data-stu-id="739f5-245">Value</span></span>|
|---|---|
|[<span data-ttu-id="739f5-246">Mindestversion des Postfachanforderungssatzes</span><span class="sxs-lookup"><span data-stu-id="739f5-246">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="739f5-247">1.0</span><span class="sxs-lookup"><span data-stu-id="739f5-247">1.0</span></span>|
|[<span data-ttu-id="739f5-248">Mindestberechtigungsstufe</span><span class="sxs-lookup"><span data-stu-id="739f5-248">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="739f5-249">ReadItem</span><span class="sxs-lookup"><span data-stu-id="739f5-249">ReadItem</span></span>|
|[<span data-ttu-id="739f5-250">Zutreffender Outlook-Modus</span><span class="sxs-lookup"><span data-stu-id="739f5-250">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="739f5-251">Verfassen oder Lesen</span><span class="sxs-lookup"><span data-stu-id="739f5-251">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="739f5-252">Beispiel</span><span class="sxs-lookup"><span data-stu-id="739f5-252">Example</span></span>

<span data-ttu-id="739f5-253">Im folgenden Beispiel wird die Endzeit eines Termins im Verfassenmodus mithilfe der [`setAsync`](/javascript/api/outlook_1_1/office.time#setasync-datetime--options--callback-)-Methode des `Time`-Objekts festgelegt.</span><span class="sxs-lookup"><span data-stu-id="739f5-253">The following example sets the end time of an appointment in compose mode by using the [`setAsync`](/javascript/api/outlook_1_1/office.time#setasync-datetime--options--callback-) method of the `Time` object.</span></span>

```JavaScript
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

#### <a name="from-emailaddressdetailsjavascriptapioutlook11officeemailaddressdetails"></a><span data-ttu-id="739f5-254">from :[EmailAddressDetails](/javascript/api/outlook_1_1/office.emailaddressdetails)</span><span class="sxs-lookup"><span data-stu-id="739f5-254">from :[EmailAddressDetails](/javascript/api/outlook_1_1/office.emailaddressdetails)</span></span>

<span data-ttu-id="739f5-p113">Ruft die E-Mail-Adresse des Absenders einer Nachricht ab. Nur Lesemodus.</span><span class="sxs-lookup"><span data-stu-id="739f5-p113">Gets the email address of the sender of a message. Read mode only.</span></span>

<span data-ttu-id="739f5-p114">Die `from`- und [`sender`](#sender-emailaddressdetailsjavascriptapioutlook11officeemailaddressdetails)-Eigenschaften stellen dieselbe Person dar, außer die Nachricht wurde von einem Delegaten gesendet. In diesem Fall stellt die `from`-Eigenschaft die Stellvertretung dar, und die sender-Eigenschaft stellt den Delegaten dar.</span><span class="sxs-lookup"><span data-stu-id="739f5-p114">The `from` and [`sender`](#sender-emailaddressdetailsjavascriptapioutlook11officeemailaddressdetails) properties represent the same person unless the message is sent by a delegate. In that case, the `from` property represents the delegator, and the sender property represents the delegate.</span></span>

> [!NOTE]
> <span data-ttu-id="739f5-259">Die `recipientType` -Eigenschaft des der `EmailAddressDetails` -Objekts der `from` -Eigenschaft ist `undefined`.</span><span class="sxs-lookup"><span data-stu-id="739f5-259">The `recipientType` property of the `EmailAddressDetails` object in the `from` property is `undefined`.</span></span>

##### <a name="type"></a><span data-ttu-id="739f5-260">Typ:</span><span class="sxs-lookup"><span data-stu-id="739f5-260">Type:</span></span>

*   [<span data-ttu-id="739f5-261">EmailAddressDetails</span><span class="sxs-lookup"><span data-stu-id="739f5-261">EmailAddressDetails</span></span>](/javascript/api/outlook_1_1/office.emailaddressdetails)

##### <a name="requirements"></a><span data-ttu-id="739f5-262">Anforderungen</span><span class="sxs-lookup"><span data-stu-id="739f5-262">Requirements</span></span>

|<span data-ttu-id="739f5-263">Anforderung</span><span class="sxs-lookup"><span data-stu-id="739f5-263">Requirement</span></span>| <span data-ttu-id="739f5-264">Wert</span><span class="sxs-lookup"><span data-stu-id="739f5-264">Value</span></span>|
|---|---|
|[<span data-ttu-id="739f5-265">Mindestversion des Postfachanforderungssatzes</span><span class="sxs-lookup"><span data-stu-id="739f5-265">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="739f5-266">1.0</span><span class="sxs-lookup"><span data-stu-id="739f5-266">1.0</span></span>|
|[<span data-ttu-id="739f5-267">Mindestberechtigungsstufe</span><span class="sxs-lookup"><span data-stu-id="739f5-267">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="739f5-268">ReadItem</span><span class="sxs-lookup"><span data-stu-id="739f5-268">ReadItem</span></span>|
|[<span data-ttu-id="739f5-269">Zutreffender Outlook-Modus</span><span class="sxs-lookup"><span data-stu-id="739f5-269">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="739f5-270">Lesen</span><span class="sxs-lookup"><span data-stu-id="739f5-270">Read</span></span>|

#### <a name="internetmessageid-string"></a><span data-ttu-id="739f5-271">internetMessageId :String</span><span class="sxs-lookup"><span data-stu-id="739f5-271">internetMessageId :String</span></span>

<span data-ttu-id="739f5-p115">Ruft die Internetnachricht-ID einer E-Mail-Nachricht ab. Nur Lesemodus.</span><span class="sxs-lookup"><span data-stu-id="739f5-p115">Gets the Internet message identifier for an email message. Read mode only.</span></span>

##### <a name="type"></a><span data-ttu-id="739f5-274">Typ:</span><span class="sxs-lookup"><span data-stu-id="739f5-274">Type:</span></span>

*   <span data-ttu-id="739f5-275">String</span><span class="sxs-lookup"><span data-stu-id="739f5-275">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="739f5-276">Anforderungen</span><span class="sxs-lookup"><span data-stu-id="739f5-276">Requirements</span></span>

|<span data-ttu-id="739f5-277">Anforderung</span><span class="sxs-lookup"><span data-stu-id="739f5-277">Requirement</span></span>| <span data-ttu-id="739f5-278">Wert</span><span class="sxs-lookup"><span data-stu-id="739f5-278">Value</span></span>|
|---|---|
|[<span data-ttu-id="739f5-279">Mindestversion des Postfachanforderungssatzes</span><span class="sxs-lookup"><span data-stu-id="739f5-279">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="739f5-280">1.0</span><span class="sxs-lookup"><span data-stu-id="739f5-280">1.0</span></span>|
|[<span data-ttu-id="739f5-281">Mindestberechtigungsstufe</span><span class="sxs-lookup"><span data-stu-id="739f5-281">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="739f5-282">ReadItem</span><span class="sxs-lookup"><span data-stu-id="739f5-282">ReadItem</span></span>|
|[<span data-ttu-id="739f5-283">Zutreffender Outlook-Modus</span><span class="sxs-lookup"><span data-stu-id="739f5-283">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="739f5-284">Lesen</span><span class="sxs-lookup"><span data-stu-id="739f5-284">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="739f5-285">Beispiel</span><span class="sxs-lookup"><span data-stu-id="739f5-285">Example</span></span>

```JavaScript
var internetMessageId = Office.context.mailbox.item.internetMessageId;
```

#### <a name="itemclass-string"></a><span data-ttu-id="739f5-286">itemClass :String</span><span class="sxs-lookup"><span data-stu-id="739f5-286">itemClass :String</span></span>

<span data-ttu-id="739f5-p116">Ruft die Exchange-Webdienste-Elementklasse des ausgewählten Elements ab. Nur Lesemodus.</span><span class="sxs-lookup"><span data-stu-id="739f5-p116">Gets the Exchange Web Services item class of the selected item. Read mode only.</span></span>

<span data-ttu-id="739f5-p117">Die `itemClass`-Eigenschaft gibt die Nachrichtenklasse des ausgewählten Elements an. Folgende Standardnachrichtenklassen für das Nachrichten- oder Terminelement sind vorhanden:</span><span class="sxs-lookup"><span data-stu-id="739f5-p117">The `itemClass` property specifies the message class of the selected item. The following are the default message classes for the message or appointment item.</span></span>

| <span data-ttu-id="739f5-291">Typ</span><span class="sxs-lookup"><span data-stu-id="739f5-291">Type</span></span> | <span data-ttu-id="739f5-292">Beschreibung</span><span class="sxs-lookup"><span data-stu-id="739f5-292">Description</span></span> | <span data-ttu-id="739f5-293">Elementklasse</span><span class="sxs-lookup"><span data-stu-id="739f5-293">item class</span></span> |
| --- | --- | --- |
| <span data-ttu-id="739f5-294">Terminelemente</span><span class="sxs-lookup"><span data-stu-id="739f5-294">Appointment items</span></span> | <span data-ttu-id="739f5-295">Dies sind Kalenderelemente der Elementklasse `IPM.Appointment` oder `IPM.Appointment.Occurence`.</span><span class="sxs-lookup"><span data-stu-id="739f5-295">These are calendar items of the item class `IPM.Appointment` or `IPM.Appointment.Occurence`.</span></span> | `IPM.Appointment`<br />`IPM.Appointment.Occurence` |
| <span data-ttu-id="739f5-296">Nachrichtenelemente</span><span class="sxs-lookup"><span data-stu-id="739f5-296">Message items</span></span> | <span data-ttu-id="739f5-297">Diese Elemente umfassen E-Mail-Nachrichten der Standardnachrichtenklasse `IPM.Note` sowie Besprechungsanfragen, -antworten und -absagen, die `IPM.Schedule.Meeting` als Basisnachrichtenklasse verwenden.</span><span class="sxs-lookup"><span data-stu-id="739f5-297">These include email messages that have the default message class `IPM.Note`, and meeting requests, responses, and cancellations, that use `IPM.Schedule.Meeting` as the base message class.</span></span> | `IPM.Note`<br />`IPM.Schedule.Meeting.Request`<br />`IPM.Schedule.Meeting.Neg`<br />`IPM.Schedule.Meeting.Pos`<br />`IPM.Schedule.Meeting.Tent`<br />`IPM.Schedule.Meeting.Canceled` |

<span data-ttu-id="739f5-298">Sie können benutzerdefinierte Nachrichtenklassen zum Erweitern einer Standardnachrichtenklasse erstellen, z. B. eine benutzerdefinierte Terminnachrichtenklasse wie `IPM.Appointment.Contoso`.</span><span class="sxs-lookup"><span data-stu-id="739f5-298">You can create custom message classes that extends a default message class, for example, a custom appointment message class `IPM.Appointment.Contoso`.</span></span>

##### <a name="type"></a><span data-ttu-id="739f5-299">Typ:</span><span class="sxs-lookup"><span data-stu-id="739f5-299">Type:</span></span>

*   <span data-ttu-id="739f5-300">String</span><span class="sxs-lookup"><span data-stu-id="739f5-300">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="739f5-301">Anforderungen</span><span class="sxs-lookup"><span data-stu-id="739f5-301">Requirements</span></span>

|<span data-ttu-id="739f5-302">Anforderung</span><span class="sxs-lookup"><span data-stu-id="739f5-302">Requirement</span></span>| <span data-ttu-id="739f5-303">Wert</span><span class="sxs-lookup"><span data-stu-id="739f5-303">Value</span></span>|
|---|---|
|[<span data-ttu-id="739f5-304">Mindestversion des Postfachanforderungssatzes</span><span class="sxs-lookup"><span data-stu-id="739f5-304">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="739f5-305">1.0</span><span class="sxs-lookup"><span data-stu-id="739f5-305">1.0</span></span>|
|[<span data-ttu-id="739f5-306">Mindestberechtigungsstufe</span><span class="sxs-lookup"><span data-stu-id="739f5-306">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="739f5-307">ReadItem</span><span class="sxs-lookup"><span data-stu-id="739f5-307">ReadItem</span></span>|
|[<span data-ttu-id="739f5-308">Zutreffender Outlook-Modus</span><span class="sxs-lookup"><span data-stu-id="739f5-308">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="739f5-309">Lesen</span><span class="sxs-lookup"><span data-stu-id="739f5-309">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="739f5-310">Beispiel</span><span class="sxs-lookup"><span data-stu-id="739f5-310">Example</span></span>

```JavaScript
var itemClass = Office.context.mailbox.item.itemClass;
```

#### <a name="nullable-itemid-string"></a><span data-ttu-id="739f5-311">(nullable) itemId :String</span><span class="sxs-lookup"><span data-stu-id="739f5-311">(nullable) itemId :String</span></span>

<span data-ttu-id="739f5-p118">Ruft den Exchange-Webdienste-Elementbezeichner für das aktuelle Element ab. Nur Lesemodus.</span><span class="sxs-lookup"><span data-stu-id="739f5-p118">Gets the Exchange Web Services item identifier for the current item. Read mode only.</span></span>

> [!NOTE]
> <span data-ttu-id="739f5-314">Der Bezeichner, der von der `itemId`-Eigenschaft zurückgegeben wird, ist der gleiche Bezeichner wie der Exchange-Webdienste-Elementbezeichner.</span><span class="sxs-lookup"><span data-stu-id="739f5-314">The identifier returned by the `itemId` property is the same as the Exchange Web Services item identifier.</span></span> <span data-ttu-id="739f5-315">Die `itemId` -Eigenschaft ist nicht identisch mit der Outlook-Eintrags-ID oder die ID von der Outlook-REST-API verwendet.</span><span class="sxs-lookup"><span data-stu-id="739f5-315">The `itemId` property is not identical to the Outlook Entry ID or the ID used by the Outlook REST API.</span></span> <span data-ttu-id="739f5-316">Vor dem REST API-Aufrufe mit diesem Wert vornehmen, es konvertiert werden soll mit `Office.context.mailbox.convertToRestId`, die verfügbar und beginnen bei Anforderung ist 1.3 festgelegt.</span><span class="sxs-lookup"><span data-stu-id="739f5-316">Before making REST API calls using this value, it should be converted using `Office.context.mailbox.convertToRestId`, which is available starting in requirement set 1.3.</span></span> <span data-ttu-id="739f5-317">Weitere Informationen finden Sie unter [Verwenden der Outlook-REST-APIs aus einem Outlook-add-in](https://docs.microsoft.com/outlook/add-ins/use-rest-api#get-the-item-id).</span><span class="sxs-lookup"><span data-stu-id="739f5-317">For more details, see [Use the Outlook REST APIs from an Outlook add-in](https://docs.microsoft.com/outlook/add-ins/use-rest-api#get-the-item-id).</span></span>

##### <a name="type"></a><span data-ttu-id="739f5-318">Typ:</span><span class="sxs-lookup"><span data-stu-id="739f5-318">Type:</span></span>

*   <span data-ttu-id="739f5-319">String</span><span class="sxs-lookup"><span data-stu-id="739f5-319">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="739f5-320">Anforderungen</span><span class="sxs-lookup"><span data-stu-id="739f5-320">Requirements</span></span>

|<span data-ttu-id="739f5-321">Anforderung</span><span class="sxs-lookup"><span data-stu-id="739f5-321">Requirement</span></span>| <span data-ttu-id="739f5-322">Wert</span><span class="sxs-lookup"><span data-stu-id="739f5-322">Value</span></span>|
|---|---|
|[<span data-ttu-id="739f5-323">Mindestversion des Postfachanforderungssatzes</span><span class="sxs-lookup"><span data-stu-id="739f5-323">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="739f5-324">1.0</span><span class="sxs-lookup"><span data-stu-id="739f5-324">1.0</span></span>|
|[<span data-ttu-id="739f5-325">Mindestberechtigungsstufe</span><span class="sxs-lookup"><span data-stu-id="739f5-325">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="739f5-326">ReadItem</span><span class="sxs-lookup"><span data-stu-id="739f5-326">ReadItem</span></span>|
|[<span data-ttu-id="739f5-327">Zutreffender Outlook-Modus</span><span class="sxs-lookup"><span data-stu-id="739f5-327">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="739f5-328">Lesen</span><span class="sxs-lookup"><span data-stu-id="739f5-328">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="739f5-329">Beispiel</span><span class="sxs-lookup"><span data-stu-id="739f5-329">Example</span></span>

<span data-ttu-id="739f5-p120">Mit dem folgende Code wird das Vorhandensein eines Elementbezeichners überprüft. Wenn die `itemId`-Eigenschaft `null` oder `undefined` zurückgibt, speichert sie das Element und ruft den Elementbezeichner aus dem asynchronen Ergebnis ab.</span><span class="sxs-lookup"><span data-stu-id="739f5-p120">The following code checks for the presence of an item identifier. If the `itemId` property returns `null` or `undefined`, it saves the item to the store and gets the item identifier from the asynchronous result.</span></span>

```JavaScript
var itemId = Office.context.mailbox.item.itemId;
if (itemId === null || itemId == undefined) {
  Office.context.mailbox.item.saveAsync(function(result){
    itemId = result.value;
  });
}
```

####  <a name="itemtype-officemailboxenumsitemtypejavascriptapioutlook11officemailboxenumsitemtype"></a><span data-ttu-id="739f5-332">itemType :[Office.MailboxEnums.ItemType](/javascript/api/outlook_1_1/office.mailboxenums.itemtype)</span><span class="sxs-lookup"><span data-stu-id="739f5-332">itemType :[Office.MailboxEnums.ItemType](/javascript/api/outlook_1_1/office.mailboxenums.itemtype)</span></span>

<span data-ttu-id="739f5-333">Ruft den Typ des Elements ab, das eine Instanz darstellt.</span><span class="sxs-lookup"><span data-stu-id="739f5-333">Gets the type of item that an instance represents.</span></span>

<span data-ttu-id="739f5-334">Die `itemType`-Eigenschaft gibt einen der Werte der `ItemType`-Enumeration zurück, der angibt, ob es sich bei der `item`-Objektinstanz um eine Nachricht oder einen Termin handelt.</span><span class="sxs-lookup"><span data-stu-id="739f5-334">The `itemType` property returns one of the `ItemType` enumeration values, indicating whether the `item` object instance is a message or an appointment.</span></span>

##### <a name="type"></a><span data-ttu-id="739f5-335">Typ:</span><span class="sxs-lookup"><span data-stu-id="739f5-335">Type:</span></span>

*   [<span data-ttu-id="739f5-336">Office.MailboxEnums.ItemType</span><span class="sxs-lookup"><span data-stu-id="739f5-336">Office.MailboxEnums.ItemType</span></span>](/javascript/api/outlook_1_1/office.mailboxenums.itemtype)

##### <a name="requirements"></a><span data-ttu-id="739f5-337">Anforderungen</span><span class="sxs-lookup"><span data-stu-id="739f5-337">Requirements</span></span>

|<span data-ttu-id="739f5-338">Anforderung</span><span class="sxs-lookup"><span data-stu-id="739f5-338">Requirement</span></span>| <span data-ttu-id="739f5-339">Wert</span><span class="sxs-lookup"><span data-stu-id="739f5-339">Value</span></span>|
|---|---|
|[<span data-ttu-id="739f5-340">Mindestversion des Postfachanforderungssatzes</span><span class="sxs-lookup"><span data-stu-id="739f5-340">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="739f5-341">1.0</span><span class="sxs-lookup"><span data-stu-id="739f5-341">1.0</span></span>|
|[<span data-ttu-id="739f5-342">Mindestberechtigungsstufe</span><span class="sxs-lookup"><span data-stu-id="739f5-342">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="739f5-343">ReadItem</span><span class="sxs-lookup"><span data-stu-id="739f5-343">ReadItem</span></span>|
|[<span data-ttu-id="739f5-344">Zutreffender Outlook-Modus</span><span class="sxs-lookup"><span data-stu-id="739f5-344">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="739f5-345">Verfassen oder Lesen</span><span class="sxs-lookup"><span data-stu-id="739f5-345">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="739f5-346">Beispiel</span><span class="sxs-lookup"><span data-stu-id="739f5-346">Example</span></span>

```JavaScript
if (Office.context.mailbox.item.itemType == Office.MailboxEnums.ItemType.Message)
  // do something
else
  // do something else
```

####  <a name="location-stringlocationjavascriptapioutlook11officelocation"></a><span data-ttu-id="739f5-347">location :String|[Location](/javascript/api/outlook_1_1/office.location)</span><span class="sxs-lookup"><span data-stu-id="739f5-347">location :String|[Location](/javascript/api/outlook_1_1/office.location)</span></span>

<span data-ttu-id="739f5-348">Ruft den Ort eines Termins ab bzw. legt ihn fest.</span><span class="sxs-lookup"><span data-stu-id="739f5-348">Gets or sets the location of an appointment.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="739f5-349">Lesemodus</span><span class="sxs-lookup"><span data-stu-id="739f5-349">Read mode</span></span>

<span data-ttu-id="739f5-350">Die `location`-Eigenschaft gibt eine Zeichenfolge zurück, die den Ort des Termins enthält.</span><span class="sxs-lookup"><span data-stu-id="739f5-350">The `location` property returns a string that contains the location of the appointment.</span></span>

##### <a name="compose-mode"></a><span data-ttu-id="739f5-351">Verfassenmodus</span><span class="sxs-lookup"><span data-stu-id="739f5-351">Compose mode</span></span>

<span data-ttu-id="739f5-352">Die `location`-Eigenschaft gibt ein `Location`-Objekt zurück, das Methoden zum Abrufen und Festlegen des Orts für einen Termin bereitstellt.</span><span class="sxs-lookup"><span data-stu-id="739f5-352">The `location` property returns a `Location` object that provides methods that are used to get and set the location of the appointment.</span></span>

##### <a name="type"></a><span data-ttu-id="739f5-353">Typ:</span><span class="sxs-lookup"><span data-stu-id="739f5-353">Type:</span></span>

*   <span data-ttu-id="739f5-354">String | [Location](/javascript/api/outlook_1_1/office.location)</span><span class="sxs-lookup"><span data-stu-id="739f5-354">String | [Location](/javascript/api/outlook_1_1/office.location)</span></span>

##### <a name="requirements"></a><span data-ttu-id="739f5-355">Anforderungen</span><span class="sxs-lookup"><span data-stu-id="739f5-355">Requirements</span></span>

|<span data-ttu-id="739f5-356">Anforderung</span><span class="sxs-lookup"><span data-stu-id="739f5-356">Requirement</span></span>| <span data-ttu-id="739f5-357">Wert</span><span class="sxs-lookup"><span data-stu-id="739f5-357">Value</span></span>|
|---|---|
|[<span data-ttu-id="739f5-358">Mindestversion des Postfachanforderungssatzes</span><span class="sxs-lookup"><span data-stu-id="739f5-358">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="739f5-359">1.0</span><span class="sxs-lookup"><span data-stu-id="739f5-359">1.0</span></span>|
|[<span data-ttu-id="739f5-360">Mindestberechtigungsstufe</span><span class="sxs-lookup"><span data-stu-id="739f5-360">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="739f5-361">ReadItem</span><span class="sxs-lookup"><span data-stu-id="739f5-361">ReadItem</span></span>|
|[<span data-ttu-id="739f5-362">Zutreffender Outlook-Modus</span><span class="sxs-lookup"><span data-stu-id="739f5-362">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="739f5-363">Verfassen oder Lesen</span><span class="sxs-lookup"><span data-stu-id="739f5-363">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="739f5-364">Beispiel</span><span class="sxs-lookup"><span data-stu-id="739f5-364">Example</span></span>

```JavaScript
var userContext = { value : 1 };
Office.context.mailbox.item.location.getAsync( { context: userContext}, callback);

function callback(asyncResult) {
  var context = asyncResult.context;
  var location = asyncResult.value;
}
```

#### <a name="normalizedsubject-string"></a><span data-ttu-id="739f5-365">normalizedSubject :String</span><span class="sxs-lookup"><span data-stu-id="739f5-365">normalizedSubject :String</span></span>

<span data-ttu-id="739f5-p121">Ruft den Betreff für ein Element ab, wobei alle Präfixe entfernt werden (einschließlich `RE:` und `FWD:`). Nur Lesemodus.</span><span class="sxs-lookup"><span data-stu-id="739f5-p121">Gets the subject of an item, with all prefixes removed (including `RE:` and `FWD:`). Read mode only.</span></span>

<span data-ttu-id="739f5-p122">Die normalizedSubject-Eigenschaft ruft den Betreff des Elements mit allen Standardpräfixen (z. B. `RE:` und `FW:`) ab, die von E-Mail-Programmen hinzugefügt werden. Wenn der Betreff des Elements mit vollständigen Präfixen abgerufen werden soll, verwenden Sie die [`subject`](#subject-stringsubjectjavascriptapioutlook11officesubject)-Eigenschaft.</span><span class="sxs-lookup"><span data-stu-id="739f5-p122">The normalizedSubject property gets the subject of the item, with any standard prefixes (such as `RE:` and `FW:`) that are added by email programs. To get the subject of the item with the prefixes intact, use the [`subject`](#subject-stringsubjectjavascriptapioutlook11officesubject) property.</span></span>

##### <a name="type"></a><span data-ttu-id="739f5-370">Typ:</span><span class="sxs-lookup"><span data-stu-id="739f5-370">Type:</span></span>

*   <span data-ttu-id="739f5-371">String</span><span class="sxs-lookup"><span data-stu-id="739f5-371">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="739f5-372">Anforderungen</span><span class="sxs-lookup"><span data-stu-id="739f5-372">Requirements</span></span>

|<span data-ttu-id="739f5-373">Anforderung</span><span class="sxs-lookup"><span data-stu-id="739f5-373">Requirement</span></span>| <span data-ttu-id="739f5-374">Wert</span><span class="sxs-lookup"><span data-stu-id="739f5-374">Value</span></span>|
|---|---|
|[<span data-ttu-id="739f5-375">Mindestversion des Postfachanforderungssatzes</span><span class="sxs-lookup"><span data-stu-id="739f5-375">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="739f5-376">1.0</span><span class="sxs-lookup"><span data-stu-id="739f5-376">1.0</span></span>|
|[<span data-ttu-id="739f5-377">Mindestberechtigungsstufe</span><span class="sxs-lookup"><span data-stu-id="739f5-377">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="739f5-378">ReadItem</span><span class="sxs-lookup"><span data-stu-id="739f5-378">ReadItem</span></span>|
|[<span data-ttu-id="739f5-379">Zutreffender Outlook-Modus</span><span class="sxs-lookup"><span data-stu-id="739f5-379">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="739f5-380">Lesen</span><span class="sxs-lookup"><span data-stu-id="739f5-380">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="739f5-381">Beispiel</span><span class="sxs-lookup"><span data-stu-id="739f5-381">Example</span></span>

```JavaScript
var normalizedSubject = Office.context.mailbox.item.normalizedSubject;
```

####  <a name="optionalattendees-arrayemailaddressdetailsjavascriptapioutlook11officeemailaddressdetailsrecipientsjavascriptapioutlook11officerecipients"></a><span data-ttu-id="739f5-382">optionalAttendees :Array.<[EmailAddressDetails](/javascript/api/outlook_1_1/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook_1_1/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="739f5-382">optionalAttendees :Array.<[EmailAddressDetails](/javascript/api/outlook_1_1/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook_1_1/office.recipients)</span></span>

<span data-ttu-id="739f5-383">Ermöglicht den Zugriff auf die optionalen Teilnehmer eines Ereignisses.</span><span class="sxs-lookup"><span data-stu-id="739f5-383">Provides access to the optional attendees of an event.</span></span> <span data-ttu-id="739f5-384">Der Typ des Objekts und die Zugriffsebene, hängt von den Modus des aktuellen Elements ab.</span><span class="sxs-lookup"><span data-stu-id="739f5-384">The type of object and level of access depends on the mode of the current item.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="739f5-385">Lesemodus</span><span class="sxs-lookup"><span data-stu-id="739f5-385">Read mode</span></span>

<span data-ttu-id="739f5-386">Die `optionalAttendees`-Eigenschaft gibt ein Array mit einem `EmailAddressDetails`-Objekt für jeden optionalen Teilnehmer an der Besprechung zurück.</span><span class="sxs-lookup"><span data-stu-id="739f5-386">The `optionalAttendees` property returns an array that contains an `EmailAddressDetails` object for each optional attendee to the meeting.</span></span>

##### <a name="compose-mode"></a><span data-ttu-id="739f5-387">Verfassenmodus</span><span class="sxs-lookup"><span data-stu-id="739f5-387">Compose mode</span></span>

<span data-ttu-id="739f5-388">Die `optionalAttendees` -Eigenschaft gibt eine `Recipients` -Objekt, das Methoden zum Abrufen oder Aktualisieren der optionalen Teilnehmer für eine Besprechung bereitstellt.</span><span class="sxs-lookup"><span data-stu-id="739f5-388">The `optionalAttendees` property returns a `Recipients` object that provides methods to get or update the optional attendees for a meeting.</span></span>

##### <a name="type"></a><span data-ttu-id="739f5-389">Typ:</span><span class="sxs-lookup"><span data-stu-id="739f5-389">Type:</span></span>

*   <span data-ttu-id="739f5-390">Array.<[EmailAddressDetails](/javascript/api/outlook_1_1/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook_1_1/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="739f5-390">Array.<[EmailAddressDetails](/javascript/api/outlook_1_1/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook_1_1/office.recipients)</span></span>

##### <a name="requirements"></a><span data-ttu-id="739f5-391">Anforderungen</span><span class="sxs-lookup"><span data-stu-id="739f5-391">Requirements</span></span>

|<span data-ttu-id="739f5-392">Anforderung</span><span class="sxs-lookup"><span data-stu-id="739f5-392">Requirement</span></span>| <span data-ttu-id="739f5-393">Wert</span><span class="sxs-lookup"><span data-stu-id="739f5-393">Value</span></span>|
|---|---|
|[<span data-ttu-id="739f5-394">Mindestversion des Postfachanforderungssatzes</span><span class="sxs-lookup"><span data-stu-id="739f5-394">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="739f5-395">1.0</span><span class="sxs-lookup"><span data-stu-id="739f5-395">1.0</span></span>|
|[<span data-ttu-id="739f5-396">Mindestberechtigungsstufe</span><span class="sxs-lookup"><span data-stu-id="739f5-396">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="739f5-397">ReadItem</span><span class="sxs-lookup"><span data-stu-id="739f5-397">ReadItem</span></span>|
|[<span data-ttu-id="739f5-398">Zutreffender Outlook-Modus</span><span class="sxs-lookup"><span data-stu-id="739f5-398">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="739f5-399">Verfassen oder Lesen</span><span class="sxs-lookup"><span data-stu-id="739f5-399">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="739f5-400">Beispiel</span><span class="sxs-lookup"><span data-stu-id="739f5-400">Example</span></span>

```JavaScript
Office.context.mailbox.item.optionalAttendees.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.optionalAttendees.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.optionalAttendees.getAsync(callback);

function callback(asyncResult) {
  var arrayOfOptionalAttendeesRecipients = asyncResult.value;
}
```

#### <a name="organizer-emailaddressdetailsjavascriptapioutlook11officeemailaddressdetails"></a><span data-ttu-id="739f5-401">organizer :[EmailAddressDetails](/javascript/api/outlook_1_1/office.emailaddressdetails)</span><span class="sxs-lookup"><span data-stu-id="739f5-401">organizer :[EmailAddressDetails](/javascript/api/outlook_1_1/office.emailaddressdetails)</span></span>

<span data-ttu-id="739f5-p124">Ruft für eine angegebene Besprechung die E-Mail-Adresse des Organisators der Besprechung ab. Nur Lesemodus.</span><span class="sxs-lookup"><span data-stu-id="739f5-p124">Gets the email address of the meeting organizer for a specified meeting. Read mode only.</span></span>

##### <a name="type"></a><span data-ttu-id="739f5-404">Typ:</span><span class="sxs-lookup"><span data-stu-id="739f5-404">Type:</span></span>

*   [<span data-ttu-id="739f5-405">EmailAddressDetails</span><span class="sxs-lookup"><span data-stu-id="739f5-405">EmailAddressDetails</span></span>](/javascript/api/outlook_1_1/office.emailaddressdetails)

##### <a name="requirements"></a><span data-ttu-id="739f5-406">Anforderungen</span><span class="sxs-lookup"><span data-stu-id="739f5-406">Requirements</span></span>

|<span data-ttu-id="739f5-407">Anforderung</span><span class="sxs-lookup"><span data-stu-id="739f5-407">Requirement</span></span>| <span data-ttu-id="739f5-408">Wert</span><span class="sxs-lookup"><span data-stu-id="739f5-408">Value</span></span>|
|---|---|
|[<span data-ttu-id="739f5-409">Mindestversion des Postfachanforderungssatzes</span><span class="sxs-lookup"><span data-stu-id="739f5-409">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="739f5-410">1.0</span><span class="sxs-lookup"><span data-stu-id="739f5-410">1.0</span></span>|
|[<span data-ttu-id="739f5-411">Mindestberechtigungsstufe</span><span class="sxs-lookup"><span data-stu-id="739f5-411">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="739f5-412">ReadItem</span><span class="sxs-lookup"><span data-stu-id="739f5-412">ReadItem</span></span>|
|[<span data-ttu-id="739f5-413">Zutreffender Outlook-Modus</span><span class="sxs-lookup"><span data-stu-id="739f5-413">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="739f5-414">Lesen</span><span class="sxs-lookup"><span data-stu-id="739f5-414">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="739f5-415">Beispiel</span><span class="sxs-lookup"><span data-stu-id="739f5-415">Example</span></span>

```JavaScript
var organizerName = Office.context.mailbox.item.organizer.displayName;
var organizerAddress = Office.context.mailbox.item.organizer.emailAddress;
```

####  <a name="requiredattendees-arrayemailaddressdetailsjavascriptapioutlook11officeemailaddressdetailsrecipientsjavascriptapioutlook11officerecipients"></a><span data-ttu-id="739f5-416">requiredAttendees :Array.<[EmailAddressDetails](/javascript/api/outlook_1_1/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook_1_1/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="739f5-416">requiredAttendees :Array.<[EmailAddressDetails](/javascript/api/outlook_1_1/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook_1_1/office.recipients)</span></span>

<span data-ttu-id="739f5-417">Ermöglicht den Zugriff auf die erforderlichen Teilnehmer eines Ereignisses.</span><span class="sxs-lookup"><span data-stu-id="739f5-417">Provides access to the required attendees of an event.</span></span> <span data-ttu-id="739f5-418">Der Typ des Objekts und die Zugriffsebene, hängt von den Modus des aktuellen Elements ab.</span><span class="sxs-lookup"><span data-stu-id="739f5-418">The type of object and level of access depends on the mode of the current item.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="739f5-419">Lesemodus</span><span class="sxs-lookup"><span data-stu-id="739f5-419">Read mode</span></span>

<span data-ttu-id="739f5-420">Die `requiredAttendees`-Eigenschaft gibt ein Array mit einem `EmailAddressDetails`-Objekt für jeden erforderlichen Teilnehmer an der Besprechung zurück.</span><span class="sxs-lookup"><span data-stu-id="739f5-420">The `requiredAttendees` property returns an array that contains an `EmailAddressDetails` object for each required attendee to the meeting.</span></span>

##### <a name="compose-mode"></a><span data-ttu-id="739f5-421">Verfassenmodus</span><span class="sxs-lookup"><span data-stu-id="739f5-421">Compose mode</span></span>

<span data-ttu-id="739f5-422">Die `requiredAttendees` -Eigenschaft gibt eine `Recipients` -Objekt, das Methoden zum Abrufen oder Aktualisieren der erforderlichen Teilnehmer für eine Besprechung bereitstellt.</span><span class="sxs-lookup"><span data-stu-id="739f5-422">The `requiredAttendees` property returns a `Recipients` object that provides methods to get or update the required attendees for a meeting.</span></span>

##### <a name="type"></a><span data-ttu-id="739f5-423">Typ:</span><span class="sxs-lookup"><span data-stu-id="739f5-423">Type:</span></span>

*   <span data-ttu-id="739f5-424">Array.<[EmailAddressDetails](/javascript/api/outlook_1_1/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook_1_1/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="739f5-424">Array.<[EmailAddressDetails](/javascript/api/outlook_1_1/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook_1_1/office.recipients)</span></span>

##### <a name="requirements"></a><span data-ttu-id="739f5-425">Anforderungen</span><span class="sxs-lookup"><span data-stu-id="739f5-425">Requirements</span></span>

|<span data-ttu-id="739f5-426">Anforderung</span><span class="sxs-lookup"><span data-stu-id="739f5-426">Requirement</span></span>| <span data-ttu-id="739f5-427">Wert</span><span class="sxs-lookup"><span data-stu-id="739f5-427">Value</span></span>|
|---|---|
|[<span data-ttu-id="739f5-428">Mindestversion des Postfachanforderungssatzes</span><span class="sxs-lookup"><span data-stu-id="739f5-428">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="739f5-429">1.0</span><span class="sxs-lookup"><span data-stu-id="739f5-429">1.0</span></span>|
|[<span data-ttu-id="739f5-430">Mindestberechtigungsstufe</span><span class="sxs-lookup"><span data-stu-id="739f5-430">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="739f5-431">ReadItem</span><span class="sxs-lookup"><span data-stu-id="739f5-431">ReadItem</span></span>|
|[<span data-ttu-id="739f5-432">Zutreffender Outlook-Modus</span><span class="sxs-lookup"><span data-stu-id="739f5-432">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="739f5-433">Verfassen oder Lesen</span><span class="sxs-lookup"><span data-stu-id="739f5-433">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="739f5-434">Beispiel</span><span class="sxs-lookup"><span data-stu-id="739f5-434">Example</span></span>

```JavaScript
Office.context.mailbox.item.requiredAttendees.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.requiredAttendees.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.requiredAttendees.getAsync(callback);

function callback(asyncResult) {
  var arrayOfRequiredAttendeesRecipients = asyncResult.value;
}
```

#### <a name="sender-emailaddressdetailsjavascriptapioutlook11officeemailaddressdetails"></a><span data-ttu-id="739f5-435">sender :[EmailAddressDetails](/javascript/api/outlook_1_1/office.emailaddressdetails)</span><span class="sxs-lookup"><span data-stu-id="739f5-435">sender :[EmailAddressDetails](/javascript/api/outlook_1_1/office.emailaddressdetails)</span></span>

<span data-ttu-id="739f5-p126">Ruft die E-Mail-Adresse des Absenders einer E-Mail-Nachricht ab. Nur Lesemodus.</span><span class="sxs-lookup"><span data-stu-id="739f5-p126">Gets the email address of the sender of an email message. Read mode only.</span></span>

<span data-ttu-id="739f5-p127">Die [`from`](#from-emailaddressdetailsjavascriptapioutlook11officeemailaddressdetails)- und `sender`-Eigenschaften stellen dieselbe Person dar, außer die Nachricht wurde von einem Delegaten gesendet. In diesem Fall stellt die `from`-Eigenschaft die Stellvertretung dar, und die sender-Eigenschaft stellt den Delegaten dar.</span><span class="sxs-lookup"><span data-stu-id="739f5-p127">The [`from`](#from-emailaddressdetailsjavascriptapioutlook11officeemailaddressdetails) and `sender` properties represent the same person unless the message is sent by a delegate. In that case, the `from` property represents the delegator, and the sender property represents the delegate.</span></span>

> [!NOTE]
> <span data-ttu-id="739f5-440">Die `recipientType` -Eigenschaft des der `EmailAddressDetails` -Objekts der `from` -Eigenschaft ist `undefined`.</span><span class="sxs-lookup"><span data-stu-id="739f5-440">The `recipientType` property of the `EmailAddressDetails` object in the `from` property is `undefined`.</span></span>

##### <a name="type"></a><span data-ttu-id="739f5-441">Typ:</span><span class="sxs-lookup"><span data-stu-id="739f5-441">Type:</span></span>

*   [<span data-ttu-id="739f5-442">EmailAddressDetails</span><span class="sxs-lookup"><span data-stu-id="739f5-442">EmailAddressDetails</span></span>](/javascript/api/outlook_1_1/office.emailaddressdetails)

##### <a name="requirements"></a><span data-ttu-id="739f5-443">Anforderungen</span><span class="sxs-lookup"><span data-stu-id="739f5-443">Requirements</span></span>

|<span data-ttu-id="739f5-444">Anforderung</span><span class="sxs-lookup"><span data-stu-id="739f5-444">Requirement</span></span>| <span data-ttu-id="739f5-445">Wert</span><span class="sxs-lookup"><span data-stu-id="739f5-445">Value</span></span>|
|---|---|
|[<span data-ttu-id="739f5-446">Mindestversion des Postfachanforderungssatzes</span><span class="sxs-lookup"><span data-stu-id="739f5-446">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="739f5-447">1.0</span><span class="sxs-lookup"><span data-stu-id="739f5-447">1.0</span></span>|
|[<span data-ttu-id="739f5-448">Mindestberechtigungsstufe</span><span class="sxs-lookup"><span data-stu-id="739f5-448">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="739f5-449">ReadItem</span><span class="sxs-lookup"><span data-stu-id="739f5-449">ReadItem</span></span>|
|[<span data-ttu-id="739f5-450">Zutreffender Outlook-Modus</span><span class="sxs-lookup"><span data-stu-id="739f5-450">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="739f5-451">Lesen</span><span class="sxs-lookup"><span data-stu-id="739f5-451">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="739f5-452">Beispiel</span><span class="sxs-lookup"><span data-stu-id="739f5-452">Example</span></span>

```JavaScript
var senderName = Office.context.mailbox.item.sender.displayName;
var senderAddress = Office.context.mailbox.item.sender.emailAddress;
```

####  <a name="start-datetimejavascriptapioutlook11officetime"></a><span data-ttu-id="739f5-453">start :Date|[Time](/javascript/api/outlook_1_1/office.time)</span><span class="sxs-lookup"><span data-stu-id="739f5-453">start :Date|[Time](/javascript/api/outlook_1_1/office.time)</span></span>

<span data-ttu-id="739f5-454">Ruft Datum und Zeit für den Beginn des Termins ab oder legt Datum und Uhrzeit fest.</span><span class="sxs-lookup"><span data-stu-id="739f5-454">Gets or sets the date and time that the appointment is to begin.</span></span>

<span data-ttu-id="739f5-p128">Die `start`-Eigenschaft wird als Datums- und Uhrzeitwert in UTC ausgedrückt. Sie können die [`convertToLocalClientTime`](office.context.mailbox.md#converttolocalclienttimetimevalue--localclienttimejavascriptapioutlook11officelocalclienttime)-Methode verwenden, um den Wert in den lokalen Uhrzeit- und Datumswert des Clients umzuwandeln.</span><span class="sxs-lookup"><span data-stu-id="739f5-p128">The `start` property is expressed as a Coordinated Universal Time (UTC) date and time value. You can use the [`convertToLocalClientTime`](office.context.mailbox.md#converttolocalclienttimetimevalue--localclienttimejavascriptapioutlook11officelocalclienttime) method to convert the value to the client’s local date and time.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="739f5-457">Lesemodus</span><span class="sxs-lookup"><span data-stu-id="739f5-457">Read mode</span></span>

<span data-ttu-id="739f5-458">Die `start`-Eigenschaft gibt ein `Date`-Objekt zurück.</span><span class="sxs-lookup"><span data-stu-id="739f5-458">The `start` property returns a `Date` object.</span></span>

##### <a name="compose-mode"></a><span data-ttu-id="739f5-459">Verfassenmodus</span><span class="sxs-lookup"><span data-stu-id="739f5-459">Compose mode</span></span>

<span data-ttu-id="739f5-460">Die `start`-Eigenschaft gibt ein `Time`-Objekt zurück.</span><span class="sxs-lookup"><span data-stu-id="739f5-460">The `start` property returns a `Time` object.</span></span>

<span data-ttu-id="739f5-461">Wenn Sie die [`Time.setAsync`](/javascript/api/outlook_1_1/office.time#setasync-datetime--options--callback-)-Methode verwenden, um die Startzeit im Verfassenmodus festzulegen, sollten Sie die [`convertToUtcClientTime`](office.context.mailbox.md#converttoutcclienttimeinput--date)-Methode verwenden, um die Ortszeit auf dem Client für den Server in UTC umzuwandeln.</span><span class="sxs-lookup"><span data-stu-id="739f5-461">When you use the [`Time.setAsync`](/javascript/api/outlook_1_1/office.time#setasync-datetime--options--callback-) method to set the start time, you should use the [`convertToUtcClientTime`](office.context.mailbox.md#converttoutcclienttimeinput--date) method to convert the local time on the client to UTC for the server.</span></span>

##### <a name="type"></a><span data-ttu-id="739f5-462">Typ:</span><span class="sxs-lookup"><span data-stu-id="739f5-462">Type:</span></span>

*   <span data-ttu-id="739f5-463">Date | [Time](/javascript/api/outlook_1_1/office.time)</span><span class="sxs-lookup"><span data-stu-id="739f5-463">Date | [Time](/javascript/api/outlook_1_1/office.time)</span></span>

##### <a name="requirements"></a><span data-ttu-id="739f5-464">Anforderungen</span><span class="sxs-lookup"><span data-stu-id="739f5-464">Requirements</span></span>

|<span data-ttu-id="739f5-465">Anforderung</span><span class="sxs-lookup"><span data-stu-id="739f5-465">Requirement</span></span>| <span data-ttu-id="739f5-466">Wert</span><span class="sxs-lookup"><span data-stu-id="739f5-466">Value</span></span>|
|---|---|
|[<span data-ttu-id="739f5-467">Mindestversion des Postfachanforderungssatzes</span><span class="sxs-lookup"><span data-stu-id="739f5-467">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="739f5-468">1.0</span><span class="sxs-lookup"><span data-stu-id="739f5-468">1.0</span></span>|
|[<span data-ttu-id="739f5-469">Mindestberechtigungsstufe</span><span class="sxs-lookup"><span data-stu-id="739f5-469">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="739f5-470">ReadItem</span><span class="sxs-lookup"><span data-stu-id="739f5-470">ReadItem</span></span>|
|[<span data-ttu-id="739f5-471">Zutreffender Outlook-Modus</span><span class="sxs-lookup"><span data-stu-id="739f5-471">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="739f5-472">Verfassen oder Lesen</span><span class="sxs-lookup"><span data-stu-id="739f5-472">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="739f5-473">Beispiel</span><span class="sxs-lookup"><span data-stu-id="739f5-473">Example</span></span>

<span data-ttu-id="739f5-474">Im folgenden Beispiel wird die Startzeit eines Termins im Verfassenmodus mithilfe der [`setAsync`](/javascript/api/outlook_1_1/office.time#setasync-datetime--options--callback-)-Methode des `Time`-Objekts festgelegt.</span><span class="sxs-lookup"><span data-stu-id="739f5-474">The following example sets the start time of an appointment in compose mode by using the [`setAsync`](/javascript/api/outlook_1_1/office.time#setasync-datetime--options--callback-) method of the `Time` object.</span></span>

```JavaScript
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

####  <a name="subject-stringsubjectjavascriptapioutlook11officesubject"></a><span data-ttu-id="739f5-475">subject :String|[Subject](/javascript/api/outlook_1_1/office.subject)</span><span class="sxs-lookup"><span data-stu-id="739f5-475">subject :String|[Subject](/javascript/api/outlook_1_1/office.subject)</span></span>

<span data-ttu-id="739f5-476">Ruft die Beschreibung ab, die im Betrefffeld eines Elements angezeigt wird, oder legt sie fest.</span><span class="sxs-lookup"><span data-stu-id="739f5-476">Gets or sets the description that appears in the subject field of an item.</span></span>

<span data-ttu-id="739f5-477">Die `subject`-Eigenschaft ruft den gesamten Betreff des Elements ab oder legt ihn fest – so, wie er vom E-Mail-Server gesendet wird.</span><span class="sxs-lookup"><span data-stu-id="739f5-477">The `subject` property gets or sets the entire subject of the item, as sent by the email server.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="739f5-478">Lesemodus</span><span class="sxs-lookup"><span data-stu-id="739f5-478">Read mode</span></span>

<span data-ttu-id="739f5-p129">Die `subject`-Eigenschaft gibt eine Zeichenfolge zurück. Verwenden Sie die [`normalizedSubject`](#normalizedsubject-string)-Eigenschaft, um den Betreff ohne führende Präfixe wie `RE:` und `FW:` abzurufen.</span><span class="sxs-lookup"><span data-stu-id="739f5-p129">The `subject` property returns a string. Use the [`normalizedSubject`](#normalizedsubject-string) property to get the subject minus any leading prefixes such as `RE:` and `FW:`.</span></span>

```
var subject = Office.context.mailbox.item.subject;
```

##### <a name="compose-mode"></a><span data-ttu-id="739f5-481">Verfassenmodus</span><span class="sxs-lookup"><span data-stu-id="739f5-481">Compose mode</span></span>

<span data-ttu-id="739f5-482">Die `subject`-Eigenschaft gibt ein `Subject`-Objekt zurück, das Methoden zum Abrufen und Festlegen des Betreffs bereitstellt.</span><span class="sxs-lookup"><span data-stu-id="739f5-482">The `subject` property returns a `Subject` object that provides methods to get and set the subject.</span></span>

```JavaScript
Office.context.mailbox.item.subject.getAsync(callback);

function callback(asyncResult) {
  var subject = asyncResult.value;
}
```

##### <a name="type"></a><span data-ttu-id="739f5-483">Typ:</span><span class="sxs-lookup"><span data-stu-id="739f5-483">Type:</span></span>

*   <span data-ttu-id="739f5-484">String | [Subject](/javascript/api/outlook_1_1/office.subject)</span><span class="sxs-lookup"><span data-stu-id="739f5-484">String | [Subject](/javascript/api/outlook_1_1/office.subject)</span></span>

##### <a name="requirements"></a><span data-ttu-id="739f5-485">Anforderungen</span><span class="sxs-lookup"><span data-stu-id="739f5-485">Requirements</span></span>

|<span data-ttu-id="739f5-486">Anforderung</span><span class="sxs-lookup"><span data-stu-id="739f5-486">Requirement</span></span>| <span data-ttu-id="739f5-487">Wert</span><span class="sxs-lookup"><span data-stu-id="739f5-487">Value</span></span>|
|---|---|
|[<span data-ttu-id="739f5-488">Mindestversion des Postfachanforderungssatzes</span><span class="sxs-lookup"><span data-stu-id="739f5-488">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="739f5-489">1.0</span><span class="sxs-lookup"><span data-stu-id="739f5-489">1.0</span></span>|
|[<span data-ttu-id="739f5-490">Mindestberechtigungsstufe</span><span class="sxs-lookup"><span data-stu-id="739f5-490">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="739f5-491">ReadItem</span><span class="sxs-lookup"><span data-stu-id="739f5-491">ReadItem</span></span>|
|[<span data-ttu-id="739f5-492">Zutreffender Outlook-Modus</span><span class="sxs-lookup"><span data-stu-id="739f5-492">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="739f5-493">Verfassen oder Lesen</span><span class="sxs-lookup"><span data-stu-id="739f5-493">Compose or read</span></span>|

####  <a name="to-arrayemailaddressdetailsjavascriptapioutlook11officeemailaddressdetailsrecipientsjavascriptapioutlook11officerecipients"></a><span data-ttu-id="739f5-494">an: Array. <[EmailAddressDetails](/javascript/api/outlook_1_1/office.emailaddressdetails)>|[Empfänger](/javascript/api/outlook_1_1/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="739f5-494">to :Array.<[EmailAddressDetails](/javascript/api/outlook_1_1/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook_1_1/office.recipients)</span></span>

<span data-ttu-id="739f5-495">Ermöglicht den Zugriff auf die Empfänger in der Zeile **an** einer Nachricht.</span><span class="sxs-lookup"><span data-stu-id="739f5-495">Provides access to the recipients on the **To** line of a message.</span></span> <span data-ttu-id="739f5-496">Der Typ des Objekts und die Zugriffsebene, hängt von den Modus des aktuellen Elements ab.</span><span class="sxs-lookup"><span data-stu-id="739f5-496">The type of object and level of access depends on the mode of the current item.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="739f5-497">Lesemodus</span><span class="sxs-lookup"><span data-stu-id="739f5-497">Read mode</span></span>

<span data-ttu-id="739f5-p131">Die `to`-Eigenschaft gibt ein Array mit einem `EmailAddressDetails`-Objekt für jeden Empfänger in der**An**-Zeile der Nachricht zurück. Die Auflistung ist auf höchstens 100 Elemente beschränkt.</span><span class="sxs-lookup"><span data-stu-id="739f5-p131">The `to` property returns an array that contains an `EmailAddressDetails` object for each recipient listed on the **To** line of the message. The collection is limited to a maximum of 100 members.</span></span>

##### <a name="compose-mode"></a><span data-ttu-id="739f5-500">Verfassenmodus</span><span class="sxs-lookup"><span data-stu-id="739f5-500">Compose mode</span></span>

<span data-ttu-id="739f5-501">Die `to` -Eigenschaft gibt eine `Recipients` -Objekt, das Methoden zum Abrufen oder aktualisieren die Empfänger in der Zeile **an** der Nachricht bereitstellt.</span><span class="sxs-lookup"><span data-stu-id="739f5-501">The `to` property returns a `Recipients` object that provides methods to get or update the recipients on the **To** line of the message.</span></span>

##### <a name="type"></a><span data-ttu-id="739f5-502">Typ:</span><span class="sxs-lookup"><span data-stu-id="739f5-502">Type:</span></span>

*   <span data-ttu-id="739f5-503">Array.<[EmailAddressDetails](/javascript/api/outlook_1_1/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook_1_1/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="739f5-503">Array.<[EmailAddressDetails](/javascript/api/outlook_1_1/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook_1_1/office.recipients)</span></span>

##### <a name="requirements"></a><span data-ttu-id="739f5-504">Anforderungen</span><span class="sxs-lookup"><span data-stu-id="739f5-504">Requirements</span></span>

|<span data-ttu-id="739f5-505">Anforderung</span><span class="sxs-lookup"><span data-stu-id="739f5-505">Requirement</span></span>| <span data-ttu-id="739f5-506">Wert</span><span class="sxs-lookup"><span data-stu-id="739f5-506">Value</span></span>|
|---|---|
|[<span data-ttu-id="739f5-507">Mindestversion des Postfachanforderungssatzes</span><span class="sxs-lookup"><span data-stu-id="739f5-507">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="739f5-508">1.0</span><span class="sxs-lookup"><span data-stu-id="739f5-508">1.0</span></span>|
|[<span data-ttu-id="739f5-509">Mindestberechtigungsstufe</span><span class="sxs-lookup"><span data-stu-id="739f5-509">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="739f5-510">ReadItem</span><span class="sxs-lookup"><span data-stu-id="739f5-510">ReadItem</span></span>|
|[<span data-ttu-id="739f5-511">Zutreffender Outlook-Modus</span><span class="sxs-lookup"><span data-stu-id="739f5-511">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="739f5-512">Verfassen oder Lesen</span><span class="sxs-lookup"><span data-stu-id="739f5-512">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="739f5-513">Beispiel</span><span class="sxs-lookup"><span data-stu-id="739f5-513">Example</span></span>

```JavaScript
Office.context.mailbox.item.to.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.to.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.to.getAsync(callback);

function callback(asyncResult) {
  var arrayOfToRecipients = asyncResult.value;
}
```

### <a name="methods"></a><span data-ttu-id="739f5-514">Methoden</span><span class="sxs-lookup"><span data-stu-id="739f5-514">Methods</span></span>

####  <a name="addfileattachmentasyncuri-attachmentname-options-callback"></a><span data-ttu-id="739f5-515">addFileAttachmentAsync(uri, attachmentName, [options], [callback])</span><span class="sxs-lookup"><span data-stu-id="739f5-515">addFileAttachmentAsync(uri, attachmentName, [options], [callback])</span></span>

<span data-ttu-id="739f5-516">Fügt eine Datei zu einer Nachricht oder einem Termin als Anlage hinzu.</span><span class="sxs-lookup"><span data-stu-id="739f5-516">Adds a file to a message or appointment as an attachment.</span></span>

<span data-ttu-id="739f5-517">Die `addFileAttachmentAsync`-Methode lädt die Datei am angegebenen URI hoch und fügt sie an das Element im Verfassenformular an.</span><span class="sxs-lookup"><span data-stu-id="739f5-517">The `addFileAttachmentAsync` method uploads the file at the specified URI and attaches it to the item in the compose form.</span></span>

<span data-ttu-id="739f5-518">Anschließend können Sie den Bezeichner mit der [`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback)-Methode in der gleichen Sitzung zum Entfernen der Anlage verwenden.</span><span class="sxs-lookup"><span data-stu-id="739f5-518">You can subsequently use the identifier with the [`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback) method to remove the attachment in the same session.</span></span>

##### <a name="parameters"></a><span data-ttu-id="739f5-519">Parameter:</span><span class="sxs-lookup"><span data-stu-id="739f5-519">Parameters:</span></span>

|<span data-ttu-id="739f5-520">Name</span><span class="sxs-lookup"><span data-stu-id="739f5-520">Name</span></span>| <span data-ttu-id="739f5-521">Typ</span><span class="sxs-lookup"><span data-stu-id="739f5-521">Type</span></span>| <span data-ttu-id="739f5-522">Attribute</span><span class="sxs-lookup"><span data-stu-id="739f5-522">Attributes</span></span>| <span data-ttu-id="739f5-523">Beschreibung</span><span class="sxs-lookup"><span data-stu-id="739f5-523">Description</span></span>|
|---|---|---|---|
|`uri`| <span data-ttu-id="739f5-524">Zeichenfolge</span><span class="sxs-lookup"><span data-stu-id="739f5-524">String</span></span>||<span data-ttu-id="739f5-p132">Der URI, der den Speicherort der an die Nachricht oder den Termin anzuhängenden Datei angibt. Die maximale Länge ist 2048 Zeichen.</span><span class="sxs-lookup"><span data-stu-id="739f5-p132">The URI that provides the location of the file to attach to the message or appointment. The maximum length is 2048 characters.</span></span>|
|`attachmentName`| <span data-ttu-id="739f5-527">Zeichenfolge</span><span class="sxs-lookup"><span data-stu-id="739f5-527">String</span></span>||<span data-ttu-id="739f5-p133">Der Name der Anlage, der beim Hochladen der Anlage angezeigt wird. Die maximale Länge ist 255 Zeichen.</span><span class="sxs-lookup"><span data-stu-id="739f5-p133">The name of the attachment that is shown while the attachment is uploading. The maximum length is 255 characters.</span></span>|
|`options`| <span data-ttu-id="739f5-530">Object</span><span class="sxs-lookup"><span data-stu-id="739f5-530">Object</span></span>| <span data-ttu-id="739f5-531">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="739f5-531">&lt;optional&gt;</span></span>|<span data-ttu-id="739f5-532">Ein Objektliteral, das mindestens eine der folgenden Eigenschaften enthält.</span><span class="sxs-lookup"><span data-stu-id="739f5-532">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`| <span data-ttu-id="739f5-533">Object</span><span class="sxs-lookup"><span data-stu-id="739f5-533">Object</span></span>| <span data-ttu-id="739f5-534">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="739f5-534">&lt;optional&gt;</span></span>|<span data-ttu-id="739f5-535">Entwickler können ein Objekt bereitstellen, auf das sie in der Callbackmethode zugreifen möchten.</span><span class="sxs-lookup"><span data-stu-id="739f5-535">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`| <span data-ttu-id="739f5-536">Funktion</span><span class="sxs-lookup"><span data-stu-id="739f5-536">function</span></span>| <span data-ttu-id="739f5-537">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="739f5-537">&lt;optional&gt;</span></span>|<span data-ttu-id="739f5-538">Wenn die Methode abgeschlossen wird, wird die Funktion , die im `callback`-Parameter übergeben wird, mit einem einzelnen Parameter (`asyncResult`) aufgerufen, der ein [`AsyncResult`](/javascript/api/office/office.asyncresult)-Objekt darstellt.</span><span class="sxs-lookup"><span data-stu-id="739f5-538">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span> <br/><span data-ttu-id="739f5-539">Bei Erfolg wird der Anlagenbezeichner in der `asyncResult.value`-Eigenschaft bereitgestellt.</span><span class="sxs-lookup"><span data-stu-id="739f5-539">On success, the attachment identifier will be provided in the `asyncResult.value` property.</span></span><br/><span data-ttu-id="739f5-540">Wenn beim Hochladen der Anlage ein Fehler auftritt, enthält das `asyncResult`-Objekt ein `Error`-Objekt mit einer Beschreibung des Fehlers.</span><span class="sxs-lookup"><span data-stu-id="739f5-540">If uploading the attachment fails, the `asyncResult` object will contain an `Error` object that provides a description of the error.</span></span>|

##### <a name="errors"></a><span data-ttu-id="739f5-541">Fehler</span><span class="sxs-lookup"><span data-stu-id="739f5-541">Errors</span></span>

| <span data-ttu-id="739f5-542">Fehlercode</span><span class="sxs-lookup"><span data-stu-id="739f5-542">Error code</span></span> | <span data-ttu-id="739f5-543">Beschreibung</span><span class="sxs-lookup"><span data-stu-id="739f5-543">Description</span></span> |
|------------|-------------|
| `AttachmentSizeExceeded` | <span data-ttu-id="739f5-544">Die Anlage ist zu groß.</span><span class="sxs-lookup"><span data-stu-id="739f5-544">The attachment is larger than allowed.</span></span> |
| `FileTypeNotSupported` | <span data-ttu-id="739f5-545">Die Anlage enthält eine Erweiterung, die nicht zulässig ist.</span><span class="sxs-lookup"><span data-stu-id="739f5-545">The attachment has an extension that is not allowed.</span></span> |
| `NumberOfAttachmentsExceeded` | <span data-ttu-id="739f5-546">Die Nachricht oder der Termin enthält zu viele Anlagen.</span><span class="sxs-lookup"><span data-stu-id="739f5-546">The message or appointment has too many attachments.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="739f5-547">Anforderungen</span><span class="sxs-lookup"><span data-stu-id="739f5-547">Requirements</span></span>

|<span data-ttu-id="739f5-548">Anforderung</span><span class="sxs-lookup"><span data-stu-id="739f5-548">Requirement</span></span>| <span data-ttu-id="739f5-549">Wert</span><span class="sxs-lookup"><span data-stu-id="739f5-549">Value</span></span>|
|---|---|
|[<span data-ttu-id="739f5-550">Mindestversion des Postfachanforderungssatzes</span><span class="sxs-lookup"><span data-stu-id="739f5-550">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="739f5-551">1.1</span><span class="sxs-lookup"><span data-stu-id="739f5-551">1.1</span></span>|
|[<span data-ttu-id="739f5-552">Mindestberechtigungsstufe</span><span class="sxs-lookup"><span data-stu-id="739f5-552">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="739f5-553">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="739f5-553">ReadWriteItem</span></span>|
|[<span data-ttu-id="739f5-554">Zutreffender Outlook-Modus</span><span class="sxs-lookup"><span data-stu-id="739f5-554">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="739f5-555">Verfassen</span><span class="sxs-lookup"><span data-stu-id="739f5-555">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="739f5-556">Beispiel</span><span class="sxs-lookup"><span data-stu-id="739f5-556">Example</span></span>

```JavaScript
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

####  <a name="additemattachmentasyncitemid-attachmentname-options-callback"></a><span data-ttu-id="739f5-557">addItemAttachmentAsync(itemId, attachmentName, [options], [callback])</span><span class="sxs-lookup"><span data-stu-id="739f5-557">addItemAttachmentAsync(itemId, attachmentName, [options], [callback])</span></span>

<span data-ttu-id="739f5-558">Fügt der Nachricht oder dem Termin ein Exchange-Objekt, wie z. B. eine Nachricht, als Anhang hinzu.</span><span class="sxs-lookup"><span data-stu-id="739f5-558">Adds an Exchange item, such as a message, as an attachment to the message or appointment.</span></span>

<span data-ttu-id="739f5-p134">Die `addItemAttachmentAsync`-Methode fügt das Element mit dem angegebenen Exchange-Bezeichner an das Element im Formular zum Verfassen an. Wenn Sie eine Rückrufmethode angeben, wird die Methode mit einem `asyncResult`-Parameter aufgerufen, der entweder den Anlagenbezeichner oder einen Code enthält, der etwaige Fehler angibt, die beim Anfügen des Objekts aufgetreten sind. Sie können ggf. den `options`-Parameter verwenden, um Statusinformationen an die Rückrufmethode zu übergeben.</span><span class="sxs-lookup"><span data-stu-id="739f5-p134">The `addItemAttachmentAsync` method attaches the item with the specified Exchange identifier to the item in the compose form. If you specify a callback method, the method is called with one parameter, `asyncResult`, which contains either the attachment identifier or a code that indicates any error that occurred while attaching the item. You can use the `options` parameter to pass state information to the callback method, if needed.</span></span>

<span data-ttu-id="739f5-562">Anschließend können Sie den Bezeichner mit der [`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback)-Methode in der gleichen Sitzung zum Entfernen der Anlage verwenden.</span><span class="sxs-lookup"><span data-stu-id="739f5-562">You can subsequently use the identifier with the [`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback) method to remove the attachment in the same session.</span></span>

<span data-ttu-id="739f5-563">Wenn Ihre Office-Add-Ins in Outlook Web App ausgeführt wird die `addItemAttachmentAsync` -Methode kann Anfügen von Elementen in der Elemente, die Sie bearbeiten; Allerdings Dies wird nicht unterstützt und wird nicht empfohlen.</span><span class="sxs-lookup"><span data-stu-id="739f5-563">If your Office Add-in is running in Outlook Web App, the `addItemAttachmentAsync` method can attach items to items other than the item that you are editing; however, this is not supported and is not recommended.</span></span>

##### <a name="parameters"></a><span data-ttu-id="739f5-564">Parameter:</span><span class="sxs-lookup"><span data-stu-id="739f5-564">Parameters:</span></span>

|<span data-ttu-id="739f5-565">Name</span><span class="sxs-lookup"><span data-stu-id="739f5-565">Name</span></span>| <span data-ttu-id="739f5-566">Typ</span><span class="sxs-lookup"><span data-stu-id="739f5-566">Type</span></span>| <span data-ttu-id="739f5-567">Attribute</span><span class="sxs-lookup"><span data-stu-id="739f5-567">Attributes</span></span>| <span data-ttu-id="739f5-568">Beschreibung</span><span class="sxs-lookup"><span data-stu-id="739f5-568">Description</span></span>|
|---|---|---|---|
|`itemId`| <span data-ttu-id="739f5-569">Zeichenfolge</span><span class="sxs-lookup"><span data-stu-id="739f5-569">String</span></span>||<span data-ttu-id="739f5-p135">Der Exchange-Bezeichner des Objekts, das angehängt werden soll. Die maximale Länge beträgt 100 Zeichen.</span><span class="sxs-lookup"><span data-stu-id="739f5-p135">The Exchange identifier of the item to attach. The maximum length is 100 characters.</span></span>|
|`attachmentName`| <span data-ttu-id="739f5-572">Zeichenfolge</span><span class="sxs-lookup"><span data-stu-id="739f5-572">String</span></span>||<span data-ttu-id="739f5-p136">Der Betreff des Elements, das angehängt werden soll. Die maximale Länge ist 255 Zeichen.</span><span class="sxs-lookup"><span data-stu-id="739f5-p136">The sujbect of the item to be attached. The maximum length is 255 characters.</span></span>|
|`options`| <span data-ttu-id="739f5-575">Object</span><span class="sxs-lookup"><span data-stu-id="739f5-575">Object</span></span>| <span data-ttu-id="739f5-576">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="739f5-576">&lt;optional&gt;</span></span>|<span data-ttu-id="739f5-577">Ein Objektliteral, das mindestens eine der folgenden Eigenschaften enthält.</span><span class="sxs-lookup"><span data-stu-id="739f5-577">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`| <span data-ttu-id="739f5-578">Object</span><span class="sxs-lookup"><span data-stu-id="739f5-578">Object</span></span>| <span data-ttu-id="739f5-579">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="739f5-579">&lt;optional&gt;</span></span>|<span data-ttu-id="739f5-580">Entwickler können ein Objekt bereitstellen, auf das sie in der Callbackmethode zugreifen möchten.</span><span class="sxs-lookup"><span data-stu-id="739f5-580">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`| <span data-ttu-id="739f5-581">Funktion</span><span class="sxs-lookup"><span data-stu-id="739f5-581">function</span></span>| <span data-ttu-id="739f5-582">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="739f5-582">&lt;optional&gt;</span></span>|<span data-ttu-id="739f5-583">Wenn die Methode abgeschlossen wird, wird die Funktion , die im `callback`-Parameter übergeben wird, mit einem einzelnen Parameter (`asyncResult`) aufgerufen, der ein [`AsyncResult`](/javascript/api/office/office.asyncresult)-Objekt darstellt.</span><span class="sxs-lookup"><span data-stu-id="739f5-583">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span> <br/><span data-ttu-id="739f5-584">Bei Erfolg wird der Anlagenbezeichner in der `asyncResult.value`-Eigenschaft bereitgestellt.</span><span class="sxs-lookup"><span data-stu-id="739f5-584">On success, the attachment identifier will be provided in the `asyncResult.value` property.</span></span><br/><span data-ttu-id="739f5-585">Wenn beim Hinzufügen der Anlage ein Fehler auftritt, enthält das `asyncResult`-Objekt ein `Error`-Objekt mit einer Beschreibung des Fehlers.</span><span class="sxs-lookup"><span data-stu-id="739f5-585">If adding the attachment fails, the `asyncResult` object will contain an `Error` object that provides a description of the error.</span></span>|

##### <a name="errors"></a><span data-ttu-id="739f5-586">Fehler</span><span class="sxs-lookup"><span data-stu-id="739f5-586">Errors</span></span>

| <span data-ttu-id="739f5-587">Fehlercode</span><span class="sxs-lookup"><span data-stu-id="739f5-587">Error code</span></span> | <span data-ttu-id="739f5-588">Beschreibung</span><span class="sxs-lookup"><span data-stu-id="739f5-588">Description</span></span> |
|------------|-------------|
| `NumberOfAttachmentsExceeded` | <span data-ttu-id="739f5-589">Die Nachricht oder der Termin enthält zu viele Anlagen.</span><span class="sxs-lookup"><span data-stu-id="739f5-589">The message or appointment has too many attachments.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="739f5-590">Anforderungen</span><span class="sxs-lookup"><span data-stu-id="739f5-590">Requirements</span></span>

|<span data-ttu-id="739f5-591">Anforderung</span><span class="sxs-lookup"><span data-stu-id="739f5-591">Requirement</span></span>| <span data-ttu-id="739f5-592">Wert</span><span class="sxs-lookup"><span data-stu-id="739f5-592">Value</span></span>|
|---|---|
|[<span data-ttu-id="739f5-593">Mindestversion des Postfachanforderungssatzes</span><span class="sxs-lookup"><span data-stu-id="739f5-593">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="739f5-594">1.1</span><span class="sxs-lookup"><span data-stu-id="739f5-594">1.1</span></span>|
|[<span data-ttu-id="739f5-595">Mindestberechtigungsstufe</span><span class="sxs-lookup"><span data-stu-id="739f5-595">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="739f5-596">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="739f5-596">ReadWriteItem</span></span>|
|[<span data-ttu-id="739f5-597">Zutreffender Outlook-Modus</span><span class="sxs-lookup"><span data-stu-id="739f5-597">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="739f5-598">Verfassen</span><span class="sxs-lookup"><span data-stu-id="739f5-598">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="739f5-599">Beispiel</span><span class="sxs-lookup"><span data-stu-id="739f5-599">Example</span></span>

<span data-ttu-id="739f5-600">Im folgenden Beispiel wird ein vorhandenes Outlook-Element als Anlage mit dem Namen `My Attachment` hinzugefügt.</span><span class="sxs-lookup"><span data-stu-id="739f5-600">The following example adds an existing Outlook item as an attachment with the name `My Attachment`.</span></span>

```JavaScript
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

#### <a name="displayreplyallformformdata"></a><span data-ttu-id="739f5-601">displayReplyAllForm(formData)</span><span class="sxs-lookup"><span data-stu-id="739f5-601">displayReplyAllForm(formData)</span></span>

<span data-ttu-id="739f5-602">Zeigt ein Antwortformular an, das den Absender und alle Empfänger der ausgewählten Nachricht oder des Organisators und alle Teilnehmer des ausgewählten Termins enthält.</span><span class="sxs-lookup"><span data-stu-id="739f5-602">Displays a reply form that includes the sender and all recipients of the selected message or the organizer and all attendees of the selected appointment.</span></span>

> [!NOTE]
> <span data-ttu-id="739f5-603">Diese Methode wird in Outlook für iOS oder Outlook für Android nicht unterstützt.</span><span class="sxs-lookup"><span data-stu-id="739f5-603">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="739f5-604">In Outlook Web App wird das Antwortformular als Popupformular in der Dreispaltenansicht und als Popupformular in der Zwei- oder Einspaltenansicht angezeigt.</span><span class="sxs-lookup"><span data-stu-id="739f5-604">In Outlook Web App, the reply form is displayed as a pop-out form in the 3-column view and a pop-up form in the 2- or 1-column view.</span></span>

<span data-ttu-id="739f5-605">Wenn einer der Zeichenfolgenparameter seinen Grenzwert überschreitet, löst `displayReplyAllForm` eine Ausnahme aus.</span><span class="sxs-lookup"><span data-stu-id="739f5-605">If any of the string parameters exceed their limits, `displayReplyAllForm` throws an exception.</span></span>

> [!NOTE]
> <span data-ttu-id="739f5-606">Die Möglichkeit zum Einschließen von Anlagen in den Anruf an `displayReplyAllForm` Anforderungssatz 1.1 wird nicht unterstützt.</span><span class="sxs-lookup"><span data-stu-id="739f5-606">The ability to include attachments in the call to `displayReplyAllForm` is not supported in requirement set 1.1.</span></span> <span data-ttu-id="739f5-607">Die Anlagenunterstützung wurde zu `displayReplyAllForm` im Anforderungssatz 1.2 und höher hinzugefügt.</span><span class="sxs-lookup"><span data-stu-id="739f5-607">Attachment support was added to `displayReplyAllForm` in requirement set 1.2 and above.</span></span>

##### <a name="parameters"></a><span data-ttu-id="739f5-608">Parameter:</span><span class="sxs-lookup"><span data-stu-id="739f5-608">Parameters:</span></span>

|<span data-ttu-id="739f5-609">Name</span><span class="sxs-lookup"><span data-stu-id="739f5-609">Name</span></span>| <span data-ttu-id="739f5-610">Typ</span><span class="sxs-lookup"><span data-stu-id="739f5-610">Type</span></span>| <span data-ttu-id="739f5-611">Beschreibung</span><span class="sxs-lookup"><span data-stu-id="739f5-611">Description</span></span>|
|---|---|---|
|`formData`| <span data-ttu-id="739f5-612">Zeichenfolge &#124; Objekt</span><span class="sxs-lookup"><span data-stu-id="739f5-612">String &#124; Object</span></span>| |<span data-ttu-id="739f5-p138">Eine Zeichenfolge, die Text- und HTML-Code enthält, die den Hauptteil des Antwortformulars darstellen. Die Zeichenfolge ist auf 32 KB beschränkt.</span><span class="sxs-lookup"><span data-stu-id="739f5-p138">A string that contains text and HTML and that represents the body of the reply form. The string is limited to 32 KB.</span></span><br/><span data-ttu-id="739f5-615">**ODER**</span><span class="sxs-lookup"><span data-stu-id="739f5-615">**OR**</span></span><br/><span data-ttu-id="739f5-p139">Ein Objekt, das Text- oder Anlagendaten und eine Rückruffunktion enthält. Das Objekt ist wie folgt definiert:</span><span class="sxs-lookup"><span data-stu-id="739f5-p139">An object that contains body or attachment data and a callback function. The object is defined as follows.</span></span> |
| `formData.htmlBody` | <span data-ttu-id="739f5-618">Zeichenfolge</span><span class="sxs-lookup"><span data-stu-id="739f5-618">String</span></span> | <span data-ttu-id="739f5-619">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="739f5-619">&lt;optional&gt;</span></span> | <span data-ttu-id="739f5-p140">Eine Zeichenfolge, die Text- und HTML-Code enthält, die den Hauptteil des Antwortformulars darstellen. Die Zeichenfolge ist auf 32 KB beschränkt.</span><span class="sxs-lookup"><span data-stu-id="739f5-p140">A string that contains text and HTML and that represents the body of the reply form. The string is limited to 32 KB.</span></span>
| `callback` | <span data-ttu-id="739f5-622">Funktion</span><span class="sxs-lookup"><span data-stu-id="739f5-622">function</span></span> | <span data-ttu-id="739f5-623">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="739f5-623">&lt;optional&gt;</span></span> | <span data-ttu-id="739f5-624">Bei Abschluss der Methode wird die im Parameter `callback` übergebene Funktion mit einem einzigen Parameter aufgerufen: `asyncResult`. Dieser Parameter ist ein Objekt des Typs [AsyncResult](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="739f5-624">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [AsyncResult](/javascript/api/office/office.asyncresult) object.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="739f5-625">Anforderungen</span><span class="sxs-lookup"><span data-stu-id="739f5-625">Requirements</span></span>

|<span data-ttu-id="739f5-626">Anforderung</span><span class="sxs-lookup"><span data-stu-id="739f5-626">Requirement</span></span>| <span data-ttu-id="739f5-627">Wert</span><span class="sxs-lookup"><span data-stu-id="739f5-627">Value</span></span>|
|---|---|
|[<span data-ttu-id="739f5-628">Mindestversion des Postfachanforderungssatzes</span><span class="sxs-lookup"><span data-stu-id="739f5-628">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="739f5-629">1.0</span><span class="sxs-lookup"><span data-stu-id="739f5-629">1.0</span></span>|
|[<span data-ttu-id="739f5-630">Mindestberechtigungsstufe</span><span class="sxs-lookup"><span data-stu-id="739f5-630">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="739f5-631">ReadItem</span><span class="sxs-lookup"><span data-stu-id="739f5-631">ReadItem</span></span>|
|[<span data-ttu-id="739f5-632">Zutreffender Outlook-Modus</span><span class="sxs-lookup"><span data-stu-id="739f5-632">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="739f5-633">Lesen</span><span class="sxs-lookup"><span data-stu-id="739f5-633">Read</span></span>|

##### <a name="examples"></a><span data-ttu-id="739f5-634">Beispiele</span><span class="sxs-lookup"><span data-stu-id="739f5-634">Examples</span></span>

<span data-ttu-id="739f5-635">Im folgenden Code wird eine Zeichenfolge an die `displayReplyAllForm`-Funktion übergeben.</span><span class="sxs-lookup"><span data-stu-id="739f5-635">The following code passes a string to the `displayReplyAllForm` function.</span></span>

```JavaScript
Office.context.mailbox.item.displayReplyAllForm('hello there');
Office.context.mailbox.item.displayReplyAllForm('<b>hello there</b>');
```

<span data-ttu-id="739f5-636">Antworten mit einem leeren Textkörper.</span><span class="sxs-lookup"><span data-stu-id="739f5-636">Reply with an empty body.</span></span>

```JavaScript
Office.context.mailbox.item.displayReplyAllForm({});
```

<span data-ttu-id="739f5-637">Antworten mit einem Textkörper.</span><span class="sxs-lookup"><span data-stu-id="739f5-637">Reply with just a body.</span></span>

```JavaScript
Office.context.mailbox.item.displayReplyAllForm(
{
  'htmlBody' : 'hi'
});
```

<span data-ttu-id="739f5-638">Antworten mit einem Textkörper und einem Rückruf.</span><span class="sxs-lookup"><span data-stu-id="739f5-638">Reply with a body and a callback.</span></span>

```JavaScript
Office.context.mailbox.item.displayReplyAllForm(
{
  'htmlBody' : 'hi',
  'callback' : function(asyncResult)
  {
    console.log(asyncResult.value);
  }
});
```

#### <a name="displayreplyformformdata"></a><span data-ttu-id="739f5-639">displayReplyForm(formData)</span><span class="sxs-lookup"><span data-stu-id="739f5-639">displayReplyForm(formData)</span></span>

<span data-ttu-id="739f5-640">Zeigt ein Antwortformular an, das nur den Absender der ausgewählten Nachricht oder den Organisator des ausgewählten Termins enthält.</span><span class="sxs-lookup"><span data-stu-id="739f5-640">Displays a reply form that includes only the sender of the selected message or the organizer of the selected appointment.</span></span>

> [!NOTE]
> <span data-ttu-id="739f5-641">Diese Methode wird in Outlook für iOS oder Outlook für Android nicht unterstützt.</span><span class="sxs-lookup"><span data-stu-id="739f5-641">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="739f5-642">In Outlook Web App wird das Antwortformular als Popupformular in der Dreispaltenansicht und als Popupformular in der Zwei- oder Einspaltenansicht angezeigt.</span><span class="sxs-lookup"><span data-stu-id="739f5-642">In Outlook Web App, the reply form is displayed as a pop-out form in the 3-column view and a pop-up form in the 2- or 1-column view.</span></span>

<span data-ttu-id="739f5-643">Wenn einer der Zeichenfolgenparameter seinen Grenzwert überschreitet, löst `displayReplyForm` eine Ausnahme aus.</span><span class="sxs-lookup"><span data-stu-id="739f5-643">If any of the string parameters exceed their limits, `displayReplyForm` throws an exception.</span></span>

> [!NOTE]
> <span data-ttu-id="739f5-644">Die Möglichkeit zum Einschließen von Anlagen in den Anruf an `displayReplyForm` Anforderungssatz 1.1 wird nicht unterstützt.</span><span class="sxs-lookup"><span data-stu-id="739f5-644">The ability to include attachments in the call to `displayReplyForm` is not supported in requirement set 1.1.</span></span> <span data-ttu-id="739f5-645">Die Anlagenunterstützung wurde zu `displayReplyForm` im Anforderungssatz 1.2 und höher hinzugefügt.</span><span class="sxs-lookup"><span data-stu-id="739f5-645">Attachment support was added to `displayReplyForm` in requirement set 1.2 and above.</span></span>

##### <a name="parameters"></a><span data-ttu-id="739f5-646">Parameter:</span><span class="sxs-lookup"><span data-stu-id="739f5-646">Parameters:</span></span>

|<span data-ttu-id="739f5-647">Name</span><span class="sxs-lookup"><span data-stu-id="739f5-647">Name</span></span>| <span data-ttu-id="739f5-648">Typ</span><span class="sxs-lookup"><span data-stu-id="739f5-648">Type</span></span>| <span data-ttu-id="739f5-649">Beschreibung</span><span class="sxs-lookup"><span data-stu-id="739f5-649">Description</span></span>|
|---|---|---|
|`formData`| <span data-ttu-id="739f5-650">Zeichenfolge &#124; Objekt</span><span class="sxs-lookup"><span data-stu-id="739f5-650">String &#124; Object</span></span>| | <span data-ttu-id="739f5-p142">Eine Zeichenfolge, die Text- und HTML-Code enthält, die den Hauptteil des Antwortformulars darstellen. Die Zeichenfolge ist auf 32 KB beschränkt.</span><span class="sxs-lookup"><span data-stu-id="739f5-p142">A string that contains text and HTML and that represents the body of the reply form. The string is limited to 32 KB.</span></span><br/><span data-ttu-id="739f5-653">**ODER**</span><span class="sxs-lookup"><span data-stu-id="739f5-653">**OR**</span></span><br/><span data-ttu-id="739f5-p143">Ein Objekt, das Text- oder Anlagendaten und eine Rückruffunktion enthält. Das Objekt ist wie folgt definiert:</span><span class="sxs-lookup"><span data-stu-id="739f5-p143">An object that contains body or attachment data and a callback function. The object is defined as follows.</span></span> |
| `formData.htmlBody` | <span data-ttu-id="739f5-656">Zeichenfolge</span><span class="sxs-lookup"><span data-stu-id="739f5-656">String</span></span> | <span data-ttu-id="739f5-657">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="739f5-657">&lt;optional&gt;</span></span> | <span data-ttu-id="739f5-p144">Eine Zeichenfolge, die Text- und HTML-Code enthält, die den Hauptteil des Antwortformulars darstellen. Die Zeichenfolge ist auf 32 KB beschränkt.</span><span class="sxs-lookup"><span data-stu-id="739f5-p144">A string that contains text and HTML and that represents the body of the reply form. The string is limited to 32 KB.</span></span>
| `callback` | <span data-ttu-id="739f5-660">Funktion</span><span class="sxs-lookup"><span data-stu-id="739f5-660">function</span></span> | <span data-ttu-id="739f5-661">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="739f5-661">&lt;optional&gt;</span></span> | <span data-ttu-id="739f5-662">Bei Abschluss der Methode wird die im Parameter `callback` übergebene Funktion mit einem einzigen Parameter aufgerufen: `asyncResult`. Dieser Parameter ist ein Objekt des Typs [AsyncResult](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="739f5-662">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [AsyncResult](/javascript/api/office/office.asyncresult) object.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="739f5-663">Anforderungen</span><span class="sxs-lookup"><span data-stu-id="739f5-663">Requirements</span></span>

|<span data-ttu-id="739f5-664">Anforderung</span><span class="sxs-lookup"><span data-stu-id="739f5-664">Requirement</span></span>| <span data-ttu-id="739f5-665">Wert</span><span class="sxs-lookup"><span data-stu-id="739f5-665">Value</span></span>|
|---|---|
|[<span data-ttu-id="739f5-666">Mindestversion des Postfachanforderungssatzes</span><span class="sxs-lookup"><span data-stu-id="739f5-666">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="739f5-667">1.0</span><span class="sxs-lookup"><span data-stu-id="739f5-667">1.0</span></span>|
|[<span data-ttu-id="739f5-668">Mindestberechtigungsstufe</span><span class="sxs-lookup"><span data-stu-id="739f5-668">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="739f5-669">ReadItem</span><span class="sxs-lookup"><span data-stu-id="739f5-669">ReadItem</span></span>|
|[<span data-ttu-id="739f5-670">Zutreffender Outlook-Modus</span><span class="sxs-lookup"><span data-stu-id="739f5-670">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="739f5-671">Lesen</span><span class="sxs-lookup"><span data-stu-id="739f5-671">Read</span></span>|

##### <a name="examples"></a><span data-ttu-id="739f5-672">Beispiele</span><span class="sxs-lookup"><span data-stu-id="739f5-672">Examples</span></span>

<span data-ttu-id="739f5-673">Im folgenden Code wird eine Zeichenfolge an die `displayReplyForm`-Funktion übergeben.</span><span class="sxs-lookup"><span data-stu-id="739f5-673">The following code passes a string to the `displayReplyForm` function.</span></span>

```JavaScript
Office.context.mailbox.item.displayReplyForm('hello there');
Office.context.mailbox.item.displayReplyForm('<b>hello there</b>');
```

<span data-ttu-id="739f5-674">Antworten mit einem leeren Textkörper.</span><span class="sxs-lookup"><span data-stu-id="739f5-674">Reply with an empty body.</span></span>

```JavaScript
Office.context.mailbox.item.displayReplyForm({});
```

<span data-ttu-id="739f5-675">Antworten mit einem Textkörper.</span><span class="sxs-lookup"><span data-stu-id="739f5-675">Reply with just a body.</span></span>

```JavaScript
Office.context.mailbox.item.displayReplyForm(
{
  'htmlBody' : 'hi'
});
```

<span data-ttu-id="739f5-676">Antworten mit einem Textkörper und einem Rückruf.</span><span class="sxs-lookup"><span data-stu-id="739f5-676">Reply with a body and a callback.</span></span>

```JavaScript
Office.context.mailbox.item.displayReplyForm(
{
  'htmlBody' : 'hi',
  'callback' : function(asyncResult)
  {
    console.log(asyncResult.value);
  }
});
```

#### <a name="getentities--entitiesjavascriptapioutlook11officeentities"></a><span data-ttu-id="739f5-677">getEntities() → {[Entities](/javascript/api/outlook_1_1/office.entities)}</span><span class="sxs-lookup"><span data-stu-id="739f5-677">getEntities() → {[Entities](/javascript/api/outlook_1_1/office.entities)}</span></span>

<span data-ttu-id="739f5-678">Ruft das ausgewählte Element Textkörper gefundenen Entitäten ab.</span><span class="sxs-lookup"><span data-stu-id="739f5-678">Gets the entities found in the selected item's body.</span></span>

> [!NOTE]
> <span data-ttu-id="739f5-679">Diese Methode wird in Outlook für iOS oder Outlook für Android nicht unterstützt.</span><span class="sxs-lookup"><span data-stu-id="739f5-679">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

##### <a name="requirements"></a><span data-ttu-id="739f5-680">Anforderungen</span><span class="sxs-lookup"><span data-stu-id="739f5-680">Requirements</span></span>

|<span data-ttu-id="739f5-681">Anforderung</span><span class="sxs-lookup"><span data-stu-id="739f5-681">Requirement</span></span>| <span data-ttu-id="739f5-682">Wert</span><span class="sxs-lookup"><span data-stu-id="739f5-682">Value</span></span>|
|---|---|
|[<span data-ttu-id="739f5-683">Mindestversion des Postfachanforderungssatzes</span><span class="sxs-lookup"><span data-stu-id="739f5-683">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="739f5-684">1.0</span><span class="sxs-lookup"><span data-stu-id="739f5-684">1.0</span></span>|
|[<span data-ttu-id="739f5-685">Mindestberechtigungsstufe</span><span class="sxs-lookup"><span data-stu-id="739f5-685">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="739f5-686">ReadItem</span><span class="sxs-lookup"><span data-stu-id="739f5-686">ReadItem</span></span>|
|[<span data-ttu-id="739f5-687">Zutreffender Outlook-Modus</span><span class="sxs-lookup"><span data-stu-id="739f5-687">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="739f5-688">Lesen</span><span class="sxs-lookup"><span data-stu-id="739f5-688">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="739f5-689">Gibt zurück:</span><span class="sxs-lookup"><span data-stu-id="739f5-689">Returns:</span></span>

<span data-ttu-id="739f5-690">Typ: [Entitäten](/javascript/api/outlook_1_1/office.entities)</span><span class="sxs-lookup"><span data-stu-id="739f5-690">Type: [Entities](/javascript/api/outlook_1_1/office.entities)</span></span>

##### <a name="example"></a><span data-ttu-id="739f5-691">Beispiel</span><span class="sxs-lookup"><span data-stu-id="739f5-691">Example</span></span>

<span data-ttu-id="739f5-692">Das folgende Beispiel greift auf die Kontakte Entitäten im Textkörper des aktuellen Elements.</span><span class="sxs-lookup"><span data-stu-id="739f5-692">The following example accesses the contacts entities in the current item's body.</span></span>

```JavaScript
var contacts = Office.context.mailbox.item.getEntities().contacts;
```

#### <a name="getentitiesbytypeentitytype--nullable-arraystringcontactjavascriptapioutlook11officecontactmeetingsuggestionjavascriptapioutlook11officemeetingsuggestionphonenumberjavascriptapioutlook11officephonenumbertasksuggestionjavascriptapioutlook11officetasksuggestion"></a><span data-ttu-id="739f5-693">getEntitiesByType(entityType) → (nullable) {Array.<(String|[Contact](/javascript/api/outlook_1_1/office.contact)|[MeetingSuggestion](/javascript/api/outlook_1_1/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook_1_1/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook_1_1/office.tasksuggestion))>}</span><span class="sxs-lookup"><span data-stu-id="739f5-693">getEntitiesByType(entityType) → (nullable) {Array.<(String|[Contact](/javascript/api/outlook_1_1/office.contact)|[MeetingSuggestion](/javascript/api/outlook_1_1/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook_1_1/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook_1_1/office.tasksuggestion))>}</span></span>

<span data-ttu-id="739f5-694">Ruft ein Array aller Entitäten des angegebenen Entitätstyps im Hauptteil des ausgewählten Elements gefunden.</span><span class="sxs-lookup"><span data-stu-id="739f5-694">Gets an array of all the entities of the specified entity type found in the selected item's body.</span></span>

> [!NOTE]
> <span data-ttu-id="739f5-695">Diese Methode wird in Outlook für iOS oder Outlook für Android nicht unterstützt.</span><span class="sxs-lookup"><span data-stu-id="739f5-695">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

##### <a name="parameters"></a><span data-ttu-id="739f5-696">Parameter:</span><span class="sxs-lookup"><span data-stu-id="739f5-696">Parameters:</span></span>

|<span data-ttu-id="739f5-697">Name</span><span class="sxs-lookup"><span data-stu-id="739f5-697">Name</span></span>| <span data-ttu-id="739f5-698">Typ</span><span class="sxs-lookup"><span data-stu-id="739f5-698">Type</span></span>| <span data-ttu-id="739f5-699">Beschreibung</span><span class="sxs-lookup"><span data-stu-id="739f5-699">Description</span></span>|
|---|---|---|
|`entityType`| [<span data-ttu-id="739f5-700">Office.MailboxEnums.EntityType</span><span class="sxs-lookup"><span data-stu-id="739f5-700">Office.MailboxEnums.EntityType</span></span>](/javascript/api/outlook_1_1/office.MailboxEnums.entitytype)|<span data-ttu-id="739f5-701">Einer der Werte der EntityType-Enumeration.</span><span class="sxs-lookup"><span data-stu-id="739f5-701">One of the EntityType enumeration values.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="739f5-702">Anforderungen</span><span class="sxs-lookup"><span data-stu-id="739f5-702">Requirements</span></span>

|<span data-ttu-id="739f5-703">Anforderung</span><span class="sxs-lookup"><span data-stu-id="739f5-703">Requirement</span></span>| <span data-ttu-id="739f5-704">Wert</span><span class="sxs-lookup"><span data-stu-id="739f5-704">Value</span></span>|
|---|---|
|[<span data-ttu-id="739f5-705">Mindestversion des Postfachanforderungssatzes</span><span class="sxs-lookup"><span data-stu-id="739f5-705">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="739f5-706">1.0</span><span class="sxs-lookup"><span data-stu-id="739f5-706">1.0</span></span>|
|[<span data-ttu-id="739f5-707">Mindestberechtigungsstufe</span><span class="sxs-lookup"><span data-stu-id="739f5-707">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="739f5-708">Eingeschränkt</span><span class="sxs-lookup"><span data-stu-id="739f5-708">Restricted</span></span>|
|[<span data-ttu-id="739f5-709">Zutreffender Outlook-Modus</span><span class="sxs-lookup"><span data-stu-id="739f5-709">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="739f5-710">Lesen</span><span class="sxs-lookup"><span data-stu-id="739f5-710">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="739f5-711">Gibt zurück:</span><span class="sxs-lookup"><span data-stu-id="739f5-711">Returns:</span></span>

<span data-ttu-id="739f5-712">Wenn der in `entityType` übergebene Wert kein gültiges Element der `EntityType`-Enumeration ist, gibt die Methode null zurück.</span><span class="sxs-lookup"><span data-stu-id="739f5-712">If the value passed in `entityType` is not a valid member of the `EntityType` enumeration, the method returns null.</span></span> <span data-ttu-id="739f5-713">Wenn keine Entitäten des angegebenen Typs im Textkörper des Elements vorhanden sind, gibt die Methode ein leeres Array zurück.</span><span class="sxs-lookup"><span data-stu-id="739f5-713">If no entities of the specified type are present in the item's body, the method returns an empty array.</span></span> <span data-ttu-id="739f5-714">Andernfalls hängt der Typ der Objekte im zurückgegebenen Array vom Typ der Entität ab, die im `entityType`-Parameter angefordert wurde.</span><span class="sxs-lookup"><span data-stu-id="739f5-714">Otherwise, the type of the objects in the returned array depends on the type of entity requested in the `entityType` parameter.</span></span>

<span data-ttu-id="739f5-715">Während Sie die minimale Berechtigungsstufe für diese Methode **Restricted** ist, erfordern einige Entitätstypen **ReadItem** für den Zugriff, wie in der folgenden Tabelle angegeben wird.</span><span class="sxs-lookup"><span data-stu-id="739f5-715">While the minimum permission level to use this method is **Restricted**, some entity types require **ReadItem** to access, as specified in the following table.</span></span>

| <span data-ttu-id="739f5-716">Wert von `entityType`</span><span class="sxs-lookup"><span data-stu-id="739f5-716">Value of `entityType`</span></span> | <span data-ttu-id="739f5-717">Typ der Objekte im zurückgegebenen Array</span><span class="sxs-lookup"><span data-stu-id="739f5-717">Type of objects in returned array</span></span> | <span data-ttu-id="739f5-718">Erforderliche Berechtigungsstufe</span><span class="sxs-lookup"><span data-stu-id="739f5-718">Required Permission Level</span></span> |
| --- | --- | --- |
| `Address` | <span data-ttu-id="739f5-719">Zeichenfolge</span><span class="sxs-lookup"><span data-stu-id="739f5-719">String</span></span> | <span data-ttu-id="739f5-720">**Restricted**</span><span class="sxs-lookup"><span data-stu-id="739f5-720">**Restricted**</span></span> |
| `Contact` | <span data-ttu-id="739f5-721">Contact</span><span class="sxs-lookup"><span data-stu-id="739f5-721">Contact</span></span> | <span data-ttu-id="739f5-722">**ReadItem**</span><span class="sxs-lookup"><span data-stu-id="739f5-722">**ReadItem**</span></span> |
| `EmailAddress` | <span data-ttu-id="739f5-723">Zeichenfolge</span><span class="sxs-lookup"><span data-stu-id="739f5-723">String</span></span> | <span data-ttu-id="739f5-724">**ReadItem**</span><span class="sxs-lookup"><span data-stu-id="739f5-724">**ReadItem**</span></span> |
| `MeetingSuggestion` | <span data-ttu-id="739f5-725">MeetingSuggestion</span><span class="sxs-lookup"><span data-stu-id="739f5-725">MeetingSuggestion</span></span> | <span data-ttu-id="739f5-726">**ReadItem**</span><span class="sxs-lookup"><span data-stu-id="739f5-726">**ReadItem**</span></span> |
| `PhoneNumber` | <span data-ttu-id="739f5-727">PhoneNumber</span><span class="sxs-lookup"><span data-stu-id="739f5-727">PhoneNumber</span></span> | <span data-ttu-id="739f5-728">**Restricted**</span><span class="sxs-lookup"><span data-stu-id="739f5-728">**Restricted**</span></span> |
| `TaskSuggestion` | <span data-ttu-id="739f5-729">TaskSuggestion</span><span class="sxs-lookup"><span data-stu-id="739f5-729">TaskSuggestion</span></span> | <span data-ttu-id="739f5-730">**ReadItem**</span><span class="sxs-lookup"><span data-stu-id="739f5-730">**ReadItem**</span></span> |
| `URL` | <span data-ttu-id="739f5-731">Zeichenfolge</span><span class="sxs-lookup"><span data-stu-id="739f5-731">String</span></span> | <span data-ttu-id="739f5-732">**Restricted**</span><span class="sxs-lookup"><span data-stu-id="739f5-732">**Restricted**</span></span> |

<span data-ttu-id="739f5-733">Typ:  Array.<(String|[Contact](/javascript/api/outlook_1_1/office.contact)|[MeetingSuggestion](/javascript/api/outlook_1_1/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook_1_1/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook_1_1/office.tasksuggestion))></span><span class="sxs-lookup"><span data-stu-id="739f5-733">Type:  Array.<(String|[Contact](/javascript/api/outlook_1_1/office.contact)|[MeetingSuggestion](/javascript/api/outlook_1_1/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook_1_1/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook_1_1/office.tasksuggestion))></span></span>


##### <a name="example"></a><span data-ttu-id="739f5-734">Beispiel</span><span class="sxs-lookup"><span data-stu-id="739f5-734">Example</span></span>

<span data-ttu-id="739f5-735">Im folgenden Beispiel wird veranschaulicht, wie auf ein Array von Zeichenfolgen, die Postadressen im Textkörper des aktuellen Elements darstellen.</span><span class="sxs-lookup"><span data-stu-id="739f5-735">The following example shows how to access an array of strings that represent postal addresses in the current item's body.</span></span>

```JavaScript
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

#### <a name="getfilteredentitiesbynamename--nullable-arraystringcontactjavascriptapioutlook11officecontactmeetingsuggestionjavascriptapioutlook11officemeetingsuggestionphonenumberjavascriptapioutlook11officephonenumbertasksuggestionjavascriptapioutlook11officetasksuggestion"></a><span data-ttu-id="739f5-736">getFilteredEntitiesByName(name) → (nullable) {Array.<(String|[Contact](/javascript/api/outlook_1_1/office.contact)|[MeetingSuggestion](/javascript/api/outlook_1_1/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook_1_1/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook_1_1/office.tasksuggestion))>}</span><span class="sxs-lookup"><span data-stu-id="739f5-736">getFilteredEntitiesByName(name) → (nullable) {Array.<(String|[Contact](/javascript/api/outlook_1_1/office.contact)|[MeetingSuggestion](/javascript/api/outlook_1_1/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook_1_1/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook_1_1/office.tasksuggestion))>}</span></span>

<span data-ttu-id="739f5-737">Gibt bekannte Entitäten im ausgewählten Element zurück, die den in der XML-Manifestdatei definierten benannten Filter übergeben.</span><span class="sxs-lookup"><span data-stu-id="739f5-737">Returns well-known entities in the selected item that pass the named filter defined in the manifest XML file.</span></span>

> [!NOTE]
> <span data-ttu-id="739f5-738">Diese Methode wird in Outlook für iOS oder Outlook für Android nicht unterstützt.</span><span class="sxs-lookup"><span data-stu-id="739f5-738">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="739f5-739">Die `getFilteredEntitiesByName`-Methode gibt die Entitäten zurück, die dem im [ItemHasKnownEntity](/javascript/office/manifest/rule#itemhasknownentity-rule)-Regelelement der XML-Manifestdatei definierten regulären Ausdruck mit dem angegebenen `FilterName`-Elementwert entsprechen.</span><span class="sxs-lookup"><span data-stu-id="739f5-739">The `getFilteredEntitiesByName` method returns the entities that match the regular expression defined in the [ItemHasKnownEntity](/javascript/office/manifest/rule#itemhasknownentity-rule) rule element in the manifest XML file with the specified `FilterName` element value.</span></span>

##### <a name="parameters"></a><span data-ttu-id="739f5-740">Parameter:</span><span class="sxs-lookup"><span data-stu-id="739f5-740">Parameters:</span></span>

|<span data-ttu-id="739f5-741">Name</span><span class="sxs-lookup"><span data-stu-id="739f5-741">Name</span></span>| <span data-ttu-id="739f5-742">Typ</span><span class="sxs-lookup"><span data-stu-id="739f5-742">Type</span></span>| <span data-ttu-id="739f5-743">Beschreibung</span><span class="sxs-lookup"><span data-stu-id="739f5-743">Description</span></span>|
|---|---|---|
|`name`| <span data-ttu-id="739f5-744">Zeichenfolge</span><span class="sxs-lookup"><span data-stu-id="739f5-744">String</span></span>|<span data-ttu-id="739f5-745">Der Name des `ItemHasKnownEntity`-Regelelements, das den entsprechenden Filter definiert.</span><span class="sxs-lookup"><span data-stu-id="739f5-745">The name of the `ItemHasKnownEntity` rule element that defines the filter to match.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="739f5-746">Anforderungen</span><span class="sxs-lookup"><span data-stu-id="739f5-746">Requirements</span></span>

|<span data-ttu-id="739f5-747">Anforderung</span><span class="sxs-lookup"><span data-stu-id="739f5-747">Requirement</span></span>| <span data-ttu-id="739f5-748">Wert</span><span class="sxs-lookup"><span data-stu-id="739f5-748">Value</span></span>|
|---|---|
|[<span data-ttu-id="739f5-749">Mindestversion des Postfachanforderungssatzes</span><span class="sxs-lookup"><span data-stu-id="739f5-749">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="739f5-750">1.0</span><span class="sxs-lookup"><span data-stu-id="739f5-750">1.0</span></span>|
|[<span data-ttu-id="739f5-751">Mindestberechtigungsstufe</span><span class="sxs-lookup"><span data-stu-id="739f5-751">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="739f5-752">ReadItem</span><span class="sxs-lookup"><span data-stu-id="739f5-752">ReadItem</span></span>|
|[<span data-ttu-id="739f5-753">Zutreffender Outlook-Modus</span><span class="sxs-lookup"><span data-stu-id="739f5-753">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="739f5-754">Lesen</span><span class="sxs-lookup"><span data-stu-id="739f5-754">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="739f5-755">Gibt zurück:</span><span class="sxs-lookup"><span data-stu-id="739f5-755">Returns:</span></span>

<span data-ttu-id="739f5-p146">Befindet sich kein `ItemHasKnownEntity`-Element im Manifest mit einem `FilterName`-Elementwert, der dem `name`-Parameter entspricht, gibt die Methode `null` zurück. Wenn der `name`-Parameter einem `ItemHasKnownEntity`-Element im Manifest entspricht, es aber keine entsprechenden Entitäten im aktuellen Element gibt, gibt die Methode ein leeres Array zurück.</span><span class="sxs-lookup"><span data-stu-id="739f5-p146">If there is no `ItemHasKnownEntity` element in the manifest with a `FilterName` element value that matches the `name` parameter, the method returns `null`. If the `name` parameter does match an `ItemHasKnownEntity` element in the manifest, but there are no entities in the current item that match, the method return an empty array.</span></span>


<span data-ttu-id="739f5-758">Typ: Array.<(String|[Contact](/javascript/api/outlook_1_1/office.contact)|[MeetingSuggestion](/javascript/api/outlook_1_1/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook_1_1/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook_1_1/office.tasksuggestion))></span><span class="sxs-lookup"><span data-stu-id="739f5-758">Type: Array.<(String|[Contact](/javascript/api/outlook_1_1/office.contact)|[MeetingSuggestion](/javascript/api/outlook_1_1/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook_1_1/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook_1_1/office.tasksuggestion))></span></span>


#### <a name="getregexmatches--object"></a><span data-ttu-id="739f5-759">getRegExMatches() → {Object}</span><span class="sxs-lookup"><span data-stu-id="739f5-759">getRegExMatches() → {Object}</span></span>

<span data-ttu-id="739f5-760">Gibt Zeichenfolgenwerte im ausgewählten Element zurück, die den in der XML-Manifestdatei definierten regulären Ausdrücken entsprechen.</span><span class="sxs-lookup"><span data-stu-id="739f5-760">Returns string values in the selected item that match the regular expressions defined in the manifest XML file.</span></span>

> [!NOTE]
> <span data-ttu-id="739f5-761">Diese Methode wird in Outlook für iOS oder Outlook für Android nicht unterstützt.</span><span class="sxs-lookup"><span data-stu-id="739f5-761">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="739f5-p147">Die `getRegExMatches`-Methode gibt die Zeichenfolgen zurück, die dem im `ItemHasRegularExpressionMatch`- oder `ItemHasKnownEntity`-Regelelement der XML-Manifestdatei definierten regulären Ausdruck entsprechen. Bei einer `ItemHasRegularExpressionMatch`-Regel muss eine entsprechende Zeichenfolge in der Eigenschaft des Elements vorhanden sein, das von dieser Regel angegeben wird. Der einfache `PropertyName`-Typ definiert die unterstützten Eigenschaften.</span><span class="sxs-lookup"><span data-stu-id="739f5-p147">The `getRegExMatches` method returns the strings that match the regular expression defined in each `ItemHasRegularExpressionMatch` or `ItemHasKnownEntity` rule element in the manifest XML file. For an `ItemHasRegularExpressionMatch` rule, a matching string has to occur in the property of the item that is specified by that rule. The `PropertyName` simple type defines the supported properties.</span></span>

<span data-ttu-id="739f5-765">Nehmen Sie z. B. an, dass ein Add-In-Manifest das folgende `Rule`-Element aufweist:</span><span class="sxs-lookup"><span data-stu-id="739f5-765">For example, consider an add-in manifest has the following `Rule` element:</span></span>

```JavaScript
<Rule xsi:type="RuleCollection" Mode="And">
  <Rule xsi:type="ItemIs" FormType="Read" ItemType="Message" />
  <Rule xsi:type="RuleCollection" Mode="Or">
    <Rule xsi:type="ItemHasRegularExpressionMatch" RegExName="fruits" RegExValue="apple|banana|coconut" PropertyName="BodyAsPlaintext" IgnoreCase="true" />
    <Rule xsi:type="ItemHasRegularExpressionMatch" RegExName="veggies" RegExValue="tomato|onion|spinach|broccoli" PropertyName="BodyAsPlaintext" IgnoreCase="true" />
  </Rule>
</Rule>
```

<span data-ttu-id="739f5-766">Das von `getRegExMatches` zurückgegebene Objekt hätte zwei Eigenschaften: `fruits` und `veggies`.</span><span class="sxs-lookup"><span data-stu-id="739f5-766">The object returned from `getRegExMatches` would have two properties: `fruits` and `veggies`.</span></span>

```JavaScript
{
  'fruits': ['apple','banana','Banana','coconut'],
  'veggies': ['tomato','onion','spinach','broccoli']
}
```

> [!NOTE]
> <span data-ttu-id="739f5-p148">Wenn Sie eine `ItemHasRegularExpressionMatch`-Regel für die Textkörpereigenschaft eines Elements festlegen, sollte der reguläre Ausdruck den Textkörper weiter filtern und nicht versuchen, den gesamten Textkörper des Elements zurückzugeben. Wenn der gesamte Textkörper eines Elements mit einem regulären Ausdruck wie `.*` abgerufen wird, werden nicht immer die gewünschten Ergebnisse erzielt.</span><span class="sxs-lookup"><span data-stu-id="739f5-p148">If you specify an `ItemHasRegularExpressionMatch` rule on the body property of an item, the regular expression should further filter the body and should not attempt to return the entire body of the item. Using a regular expression such as `.*` to obtain the entire body of an item does not always return the expected results.</span></span>

##### <a name="requirements"></a><span data-ttu-id="739f5-769">Anforderungen</span><span class="sxs-lookup"><span data-stu-id="739f5-769">Requirements</span></span>

|<span data-ttu-id="739f5-770">Anforderung</span><span class="sxs-lookup"><span data-stu-id="739f5-770">Requirement</span></span>| <span data-ttu-id="739f5-771">Wert</span><span class="sxs-lookup"><span data-stu-id="739f5-771">Value</span></span>|
|---|---|
|[<span data-ttu-id="739f5-772">Mindestversion des Postfachanforderungssatzes</span><span class="sxs-lookup"><span data-stu-id="739f5-772">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="739f5-773">1.0</span><span class="sxs-lookup"><span data-stu-id="739f5-773">1.0</span></span>|
|[<span data-ttu-id="739f5-774">Mindestberechtigungsstufe</span><span class="sxs-lookup"><span data-stu-id="739f5-774">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="739f5-775">ReadItem</span><span class="sxs-lookup"><span data-stu-id="739f5-775">ReadItem</span></span>|
|[<span data-ttu-id="739f5-776">Zutreffender Outlook-Modus</span><span class="sxs-lookup"><span data-stu-id="739f5-776">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="739f5-777">Lesen</span><span class="sxs-lookup"><span data-stu-id="739f5-777">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="739f5-778">Gibt zurück:</span><span class="sxs-lookup"><span data-stu-id="739f5-778">Returns:</span></span>

<span data-ttu-id="739f5-p149">Ein Objekt mit Arrays aus Zeichenfolgen, die den in der XML-Manifestdatei definierten regulären Ausdrücken entsprechen. Der Name der einzelnen Arrays ist gleich dem entsprechenden Wert des `RegExName`-Attributs der entsprechenden `ItemHasRegularExpressionMatch`-Regel oder des `FilterName`-Attributs der entsprechenden `ItemHasKnownEntity`-Regel.</span><span class="sxs-lookup"><span data-stu-id="739f5-p149">An object that contains arrays of strings that match the regular expressions defined in the manifest XML file. The name of each array is equal to the corresponding value of the `RegExName` attribute of the matching `ItemHasRegularExpressionMatch` rule or the `FilterName` attribute of the matching `ItemHasKnownEntity` rule.</span></span>

<dl class="param-type"><span data-ttu-id="739f5-781">

<dt>
Typ</dt>


</span><span class="sxs-lookup"><span data-stu-id="739f5-781">

<dt>Type</dt>

</span></span><dd><span data-ttu-id="739f5-782">Object</span><span class="sxs-lookup"><span data-stu-id="739f5-782">Object</span></span></dd>

</dl>

##### <a name="example"></a><span data-ttu-id="739f5-783">Beispiel</span><span class="sxs-lookup"><span data-stu-id="739f5-783">Example</span></span>

<span data-ttu-id="739f5-784">Das folgende Beispiel zeigt, wie Sie auf das Array von Übereinstimmungen für die regulären <rule>-Ausdruckselemente `fruits` und `veggies` zugreifen, die im Manifest angegeben sind.</rule></span><span class="sxs-lookup"><span data-stu-id="739f5-784">The following example shows how to access the array of matches for the regular expression <rule>elements `fruits` and `veggies`, which are specified in the manifest.</rule></span></span>

```JavaScript
var allMatches = Office.context.mailbox.item.getRegExMatches();
var fruits = allMatches.fruits;
var veges = allMatches.veggies;
```

#### <a name="getregexmatchesbynamename--nullable-array-string-"></a><span data-ttu-id="739f5-785">getRegExMatchesByName(name) → (Nullwerte zulassen) {Array. < Zeichenfolge >}</span><span class="sxs-lookup"><span data-stu-id="739f5-785">getRegExMatchesByName(name) → (nullable) {Array.< String >}</span></span>

<span data-ttu-id="739f5-786">Gibt Zeichenfolgenwerte im ausgewählten Element zurück, die dem in der XML-Manifestdatei definierten benannten regulären Ausdruck entsprechen.</span><span class="sxs-lookup"><span data-stu-id="739f5-786">Returns string values in the selected item that match the named regular expression defined in the manifest XML file.</span></span>

> [!NOTE]
> <span data-ttu-id="739f5-787">Diese Methode wird in Outlook für iOS oder Outlook für Android nicht unterstützt.</span><span class="sxs-lookup"><span data-stu-id="739f5-787">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="739f5-788">Die `getRegExMatchesByName`-Methode gibt die Zeichenfolgen zurück, die dem im `ItemHasRegularExpressionMatch`-Regelelement der XML-Manifestdatei definierten regulären Ausdruck mit dem angegebenen `RegExName`-Elementwert entsprechen.</span><span class="sxs-lookup"><span data-stu-id="739f5-788">The `getRegExMatchesByName` method returns the strings that match the regular expression defined in the `ItemHasRegularExpressionMatch` rule element in the manifest XML file with the specified `RegExName` element value.</span></span>

<span data-ttu-id="739f5-p150">Wenn Sie eine `ItemHasRegularExpressionMatch`-Regel für die Textkörpereigenschaft eines Elements festlegen, sollte der reguläre Ausdruck den Textkörper weiter filtern und nicht versuchen, den gesamten Textkörper des Elements zurückzugeben. Wenn der gesamte Textkörper eines Elements mit einem regulären Ausdruck wie `.*` abgerufen wird, werden nicht immer die gewünschten Ergebnisse erzielt.</span><span class="sxs-lookup"><span data-stu-id="739f5-p150">If you specify an `ItemHasRegularExpressionMatch` rule on the body property of an item, the regular expression should further filter the body and should not attempt to return the entire body of the item. Using a regular expression such as `.*` to obtain the entire body of an item does not always return the expected results.</span></span>

##### <a name="parameters"></a><span data-ttu-id="739f5-791">Parameter:</span><span class="sxs-lookup"><span data-stu-id="739f5-791">Parameters:</span></span>

|<span data-ttu-id="739f5-792">Name</span><span class="sxs-lookup"><span data-stu-id="739f5-792">Name</span></span>| <span data-ttu-id="739f5-793">Typ</span><span class="sxs-lookup"><span data-stu-id="739f5-793">Type</span></span>| <span data-ttu-id="739f5-794">Beschreibung</span><span class="sxs-lookup"><span data-stu-id="739f5-794">Description</span></span>|
|---|---|---|
|`name`| <span data-ttu-id="739f5-795">Zeichenfolge</span><span class="sxs-lookup"><span data-stu-id="739f5-795">String</span></span>|<span data-ttu-id="739f5-796">Der Name des `ItemHasRegularExpressionMatch`-Regelelements, das den entsprechenden Filter definiert.</span><span class="sxs-lookup"><span data-stu-id="739f5-796">The name of the `ItemHasRegularExpressionMatch` rule element that defines the filter to match.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="739f5-797">Anforderungen</span><span class="sxs-lookup"><span data-stu-id="739f5-797">Requirements</span></span>

|<span data-ttu-id="739f5-798">Anforderung</span><span class="sxs-lookup"><span data-stu-id="739f5-798">Requirement</span></span>| <span data-ttu-id="739f5-799">Wert</span><span class="sxs-lookup"><span data-stu-id="739f5-799">Value</span></span>|
|---|---|
|[<span data-ttu-id="739f5-800">Mindestversion des Postfachanforderungssatzes</span><span class="sxs-lookup"><span data-stu-id="739f5-800">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="739f5-801">1.0</span><span class="sxs-lookup"><span data-stu-id="739f5-801">1.0</span></span>|
|[<span data-ttu-id="739f5-802">Mindestberechtigungsstufe</span><span class="sxs-lookup"><span data-stu-id="739f5-802">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="739f5-803">ReadItem</span><span class="sxs-lookup"><span data-stu-id="739f5-803">ReadItem</span></span>|
|[<span data-ttu-id="739f5-804">Zutreffender Outlook-Modus</span><span class="sxs-lookup"><span data-stu-id="739f5-804">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="739f5-805">Lesen</span><span class="sxs-lookup"><span data-stu-id="739f5-805">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="739f5-806">Gibt zurück:</span><span class="sxs-lookup"><span data-stu-id="739f5-806">Returns:</span></span>

<span data-ttu-id="739f5-807">Ein Array mit den Zeichenfolgen, die dem in der XML-Manifestdatei definierten regulären Ausdruck entsprechen.</span><span class="sxs-lookup"><span data-stu-id="739f5-807">An array that contains the strings that match the regular expression defined in the manifest XML file.</span></span>

<dl class="param-type"><span data-ttu-id="739f5-808">

<dt>Typ</dt>

</span><span class="sxs-lookup"><span data-stu-id="739f5-808">

<dt>Type</dt>

</span></span><dd><span data-ttu-id="739f5-809">Array. < Zeichenfolge ></span><span class="sxs-lookup"><span data-stu-id="739f5-809">Array.< String ></span></span></dd>

</dl>

##### <a name="example"></a><span data-ttu-id="739f5-810">Beispiel</span><span class="sxs-lookup"><span data-stu-id="739f5-810">Example</span></span>

```JavaScript
var fruits = Office.context.mailbox.item.getRegExMatchesByName("fruits");
var veggies = Office.context.mailbox.item.getRegExMatchesByName("veggies");
```

####  <a name="loadcustompropertiesasynccallback-usercontext"></a><span data-ttu-id="739f5-811">loadCustomPropertiesAsync(callback, [userContext])</span><span class="sxs-lookup"><span data-stu-id="739f5-811">loadCustomPropertiesAsync(callback, [userContext])</span></span>

<span data-ttu-id="739f5-812">Lädt asynchron benutzerdefinierte Eigenschaften für dieses Add-In für das ausgewählte Element.</span><span class="sxs-lookup"><span data-stu-id="739f5-812">Asynchronously loads custom properties for this add-in on the selected item.</span></span>

<span data-ttu-id="739f5-p151">Benutzerdefinierte Eigenschaften werden als Schlüssel-/Wert-Paare pro App und pro Element gespeichert. Diese Methode gibt ein `CustomProperties`-Objekt im Rückruf zurück, das Methoden für den Zugriff auf die benutzerdefinierten Eigenschaften für das aktuelle Element und das aktuelle Add-In bietet. Benutzerdefinierte Eigenschaften sind für das Element nicht verschlüsselt und sollten darum nicht als sicherer Speicher verwendet werden.</span><span class="sxs-lookup"><span data-stu-id="739f5-p151">Custom properties are stored as key/value pairs on a per-app, per-item basis. This method returns a `CustomProperties` object in the callback, which provides methods to access the custom properties specific to the current item and the current add-in. Custom properties are not encrypted on the item, so this should not be used as secure storage.</span></span>

##### <a name="parameters"></a><span data-ttu-id="739f5-816">Parameter:</span><span class="sxs-lookup"><span data-stu-id="739f5-816">Parameters:</span></span>

|<span data-ttu-id="739f5-817">Name</span><span class="sxs-lookup"><span data-stu-id="739f5-817">Name</span></span>| <span data-ttu-id="739f5-818">Typ</span><span class="sxs-lookup"><span data-stu-id="739f5-818">Type</span></span>| <span data-ttu-id="739f5-819">Attribute</span><span class="sxs-lookup"><span data-stu-id="739f5-819">Attributes</span></span>| <span data-ttu-id="739f5-820">Beschreibung</span><span class="sxs-lookup"><span data-stu-id="739f5-820">Description</span></span>|
|---|---|---|---|
|`callback`| <span data-ttu-id="739f5-821">Funktion</span><span class="sxs-lookup"><span data-stu-id="739f5-821">function</span></span>||<span data-ttu-id="739f5-822">Wenn die Methode abgeschlossen wird, wird die Funktion , die im `callback`-Parameter übergeben wird, mit einem einzelnen Parameter (`asyncResult`) aufgerufen, der ein [`AsyncResult`](/javascript/api/office/office.asyncresult)-Objekt darstellt.</span><span class="sxs-lookup"><span data-stu-id="739f5-822">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span><br/><br/><span data-ttu-id="739f5-823">Die benutzerdefinierten Eigenschaften werden als [`CustomProperties`](/javascript/api/outlook_1_1/office.customproperties)-Objekt in der `asyncResult.value`-Eigenschaft bereitgestellt.</span><span class="sxs-lookup"><span data-stu-id="739f5-823">The custom properties are provided as a [`CustomProperties`](/javascript/api/outlook_1_1/office.customproperties) object in the `asyncResult.value` property.</span></span> <span data-ttu-id="739f5-824">Dieses Objekt kann zum Abrufen, festlegen und Entfernen benutzerdefinierter Eigenschaften aus dem Element und speichern Sie Änderungen an den benutzerdefinierten Eigenschaftensatz zurück an den Server verwendet werden.</span><span class="sxs-lookup"><span data-stu-id="739f5-824">This object can be used to get, set, and remove custom properties from the item and save changes to the custom property set back to the server.</span></span>|
|`userContext`| <span data-ttu-id="739f5-825">Objekt</span><span class="sxs-lookup"><span data-stu-id="739f5-825">Object</span></span>| <span data-ttu-id="739f5-826">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="739f5-826">&lt;optional&gt;</span></span>|<span data-ttu-id="739f5-827">Entwickler können ein beliebiges Objekt bereitstellen, den sie in der Rückruffunktion zugreifen möchten.</span><span class="sxs-lookup"><span data-stu-id="739f5-827">Developers can provide any object they wish to access in the callback function.</span></span> <span data-ttu-id="739f5-828">Dieses Objekt zugegriffen werden kann, indem die `asyncResult.asyncContext` -Eigenschaft in der Callback-Funktion.</span><span class="sxs-lookup"><span data-stu-id="739f5-828">This object can be accessed by the `asyncResult.asyncContext` property in the callback function.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="739f5-829">Anforderungen</span><span class="sxs-lookup"><span data-stu-id="739f5-829">Requirements</span></span>

|<span data-ttu-id="739f5-830">Anforderung</span><span class="sxs-lookup"><span data-stu-id="739f5-830">Requirement</span></span>| <span data-ttu-id="739f5-831">Wert</span><span class="sxs-lookup"><span data-stu-id="739f5-831">Value</span></span>|
|---|---|
|[<span data-ttu-id="739f5-832">Mindestversion des Postfachanforderungssatzes</span><span class="sxs-lookup"><span data-stu-id="739f5-832">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="739f5-833">1.0</span><span class="sxs-lookup"><span data-stu-id="739f5-833">1.0</span></span>|
|[<span data-ttu-id="739f5-834">Mindestberechtigungsstufe</span><span class="sxs-lookup"><span data-stu-id="739f5-834">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="739f5-835">ReadItem</span><span class="sxs-lookup"><span data-stu-id="739f5-835">ReadItem</span></span>|
|[<span data-ttu-id="739f5-836">Zutreffender Outlook-Modus</span><span class="sxs-lookup"><span data-stu-id="739f5-836">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="739f5-837">Verfassen oder Lesen</span><span class="sxs-lookup"><span data-stu-id="739f5-837">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="739f5-838">Beispiel</span><span class="sxs-lookup"><span data-stu-id="739f5-838">Example</span></span>

<span data-ttu-id="739f5-p154">Das folgende Codebeispiel veranschaulicht die Verwendung der `loadCustomPropertiesAsync`-Methode zum asynchronen Laden der benutzerdefinierten Eigenschaften, die für das aktuelle Element spezifisch sind. Des Weiteren veranschaulicht das Beispiel die Verwendung der `CustomProperties.saveAsync`-Methode zum Speichern der Eigenschaften auf dem Server. Nach dem Laden der benutzerdefinierten Eigenschaften wird in dem Codebeispiel die `CustomProperties.get`-Methode dazu verwendet, die benutzerdefinierte `myProp`-Eigenschaft zu lesen. Die `CustomProperties.set`-Methode wird dazu verwendet, die benutzerdefinierte `otherProp`-Eigenschaft zu schreiben. Schließlich wird die `saveAsync`-Methode aufgerufen, um die benutzerdefinierten Eigenschaften zu speichern.</span><span class="sxs-lookup"><span data-stu-id="739f5-p154">The following code example shows how to use the `loadCustomPropertiesAsync` method to asynchronously load custom properties that are specific to the current item. The example also shows how to use the `CustomProperties.saveAsync` method to save these properties back to the server. After loading the custom properties, the code sample uses the `CustomProperties.get` method to read the custom property `myProp`, the `CustomProperties.set` method to write the custom property `otherProp`, and then finally calls the `saveAsync` method to save the custom properties.</span></span>

```JavaScript
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

####  <a name="removeattachmentasyncattachmentid-options-callback"></a><span data-ttu-id="739f5-842">removeAttachmentAsync(attachmentId, [options], [callback])</span><span class="sxs-lookup"><span data-stu-id="739f5-842">removeAttachmentAsync(attachmentId, [options], [callback])</span></span>

<span data-ttu-id="739f5-843">Entfernt eine Anlage aus einer Nachricht oder einem Termin.</span><span class="sxs-lookup"><span data-stu-id="739f5-843">Removes an attachment from a message or appointment.</span></span>

<span data-ttu-id="739f5-p155">Die `removeAttachmentAsync`-Methode entfernt die Anlage mit dem angegebenen Bezeichner aus dem Element. Als bewährte Vorgehensweise sollten Sie den Anlagenbezeichner nur dann zum Entfernen einer Anlage verwenden, wenn die gleiche Mail-App die Anlage in der gleichen Sitzung hinzugefügt hat. In Outlook Web App und OWA für Geräte ist der Anlagenbezeichner nur innerhalb der gleichen Sitzung gültig. Eine Sitzung ist abgeschlossen, wenn der Benutzer die App schließt, oder wenn der Benutzer in einem eingebetteten Formular mit dem Verfassen beginnt und den Vorgang dann in einem separaten Fenster fortsetzt.</span><span class="sxs-lookup"><span data-stu-id="739f5-p155">The `removeAttachmentAsync` method removes the attachment with the specified identifier from the item. As a best practice, you should use the attachment identifier to remove an attachment only if the same mail app has added that attachment in the same session. In Outlook Web App and OWA for Devices, the attachment identifier is valid only within the same session. A session is over when the user closes the app, or if the user starts composing in an inline form and subsequently pops out the inline form to continue in a separate window.</span></span>

##### <a name="parameters"></a><span data-ttu-id="739f5-848">Parameter:</span><span class="sxs-lookup"><span data-stu-id="739f5-848">Parameters:</span></span>

|<span data-ttu-id="739f5-849">Name</span><span class="sxs-lookup"><span data-stu-id="739f5-849">Name</span></span>| <span data-ttu-id="739f5-850">Typ</span><span class="sxs-lookup"><span data-stu-id="739f5-850">Type</span></span>| <span data-ttu-id="739f5-851">Attribute</span><span class="sxs-lookup"><span data-stu-id="739f5-851">Attributes</span></span>| <span data-ttu-id="739f5-852">Beschreibung</span><span class="sxs-lookup"><span data-stu-id="739f5-852">Description</span></span>|
|---|---|---|---|
|`attachmentId`| <span data-ttu-id="739f5-853">Zeichenfolge</span><span class="sxs-lookup"><span data-stu-id="739f5-853">String</span></span>||<span data-ttu-id="739f5-p156">Der Bezeichner des Anhangs, der entfernt werden soll. Die maximale Länge der Zeichenfolge ist 100 Zeichen.</span><span class="sxs-lookup"><span data-stu-id="739f5-p156">The identifier of the attachment to remove. The maximum length of the string is 100 characters.</span></span>|
|`options`| <span data-ttu-id="739f5-856">Object</span><span class="sxs-lookup"><span data-stu-id="739f5-856">Object</span></span>| <span data-ttu-id="739f5-857">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="739f5-857">&lt;optional&gt;</span></span>|<span data-ttu-id="739f5-858">Ein Objektliteral, das mindestens eine der folgenden Eigenschaften enthält.</span><span class="sxs-lookup"><span data-stu-id="739f5-858">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`| <span data-ttu-id="739f5-859">Object</span><span class="sxs-lookup"><span data-stu-id="739f5-859">Object</span></span>| <span data-ttu-id="739f5-860">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="739f5-860">&lt;optional&gt;</span></span>|<span data-ttu-id="739f5-861">Entwickler können ein Objekt bereitstellen, auf das sie in der Callbackmethode zugreifen möchten.</span><span class="sxs-lookup"><span data-stu-id="739f5-861">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`| <span data-ttu-id="739f5-862">Funktion</span><span class="sxs-lookup"><span data-stu-id="739f5-862">function</span></span>| <span data-ttu-id="739f5-863">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="739f5-863">&lt;optional&gt;</span></span>|<span data-ttu-id="739f5-864">Wenn die Methode abgeschlossen wird, wird die Funktion , die im `callback`-Parameter übergeben wird, mit einem einzelnen Parameter (`asyncResult`) aufgerufen, der ein [`AsyncResult`](/javascript/api/office/office.asyncresult)-Objekt darstellt.</span><span class="sxs-lookup"><span data-stu-id="739f5-864">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span> <br/><span data-ttu-id="739f5-865">Wenn beim Entfernen der Anlage ein Fehler auftritt, enthält die Eigenschaft `asyncResult.error` einen Fehlercode mit dem Grund für den Fehler.</span><span class="sxs-lookup"><span data-stu-id="739f5-865">If removing the attachment fails, the `asyncResult.error` property will contain an error code with the reason for the failure.</span></span>|

##### <a name="errors"></a><span data-ttu-id="739f5-866">Fehler</span><span class="sxs-lookup"><span data-stu-id="739f5-866">Errors</span></span>

| <span data-ttu-id="739f5-867">Fehlercode</span><span class="sxs-lookup"><span data-stu-id="739f5-867">Error code</span></span> | <span data-ttu-id="739f5-868">Beschreibung</span><span class="sxs-lookup"><span data-stu-id="739f5-868">Description</span></span> |
|------------|-------------|
| `InvalidAttachmentId` | <span data-ttu-id="739f5-869">Der Bezeichner für die Anlage ist nicht vorhanden.</span><span class="sxs-lookup"><span data-stu-id="739f5-869">The attachment identifier does not exist.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="739f5-870">Anforderungen</span><span class="sxs-lookup"><span data-stu-id="739f5-870">Requirements</span></span>

|<span data-ttu-id="739f5-871">Anforderung</span><span class="sxs-lookup"><span data-stu-id="739f5-871">Requirement</span></span>| <span data-ttu-id="739f5-872">Wert</span><span class="sxs-lookup"><span data-stu-id="739f5-872">Value</span></span>|
|---|---|
|[<span data-ttu-id="739f5-873">Mindestversion des Postfachanforderungssatzes</span><span class="sxs-lookup"><span data-stu-id="739f5-873">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="739f5-874">1.1</span><span class="sxs-lookup"><span data-stu-id="739f5-874">1.1</span></span>|
|[<span data-ttu-id="739f5-875">Mindestberechtigungsstufe</span><span class="sxs-lookup"><span data-stu-id="739f5-875">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="739f5-876">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="739f5-876">ReadWriteItem</span></span>|
|[<span data-ttu-id="739f5-877">Zutreffender Outlook-Modus</span><span class="sxs-lookup"><span data-stu-id="739f5-877">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="739f5-878">Verfassen</span><span class="sxs-lookup"><span data-stu-id="739f5-878">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="739f5-879">Beispiel</span><span class="sxs-lookup"><span data-stu-id="739f5-879">Example</span></span>

<span data-ttu-id="739f5-880">Mit dem folgende Code wird eine Anlage mit dem Bezeichner "0" entfernt.</span><span class="sxs-lookup"><span data-stu-id="739f5-880">The following code removes an attachment with an identifier of '0'.</span></span>

```JavaScript
Office.context.mailbox.item.removeAttachmentAsync(
  '0',
  { asyncContext : null },
  function (asyncResult)
  {
    console.log(asyncResult.status);
  }
);
```