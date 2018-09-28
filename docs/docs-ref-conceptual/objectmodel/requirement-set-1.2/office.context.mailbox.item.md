
# <a name="item"></a><span data-ttu-id="b89d7-101">item</span><span class="sxs-lookup"><span data-stu-id="b89d7-101">item</span></span>

### <span data-ttu-id="b89d7-p101">[Office](office.md)[.context](office.context.md)[.mailbox](office.context.mailbox.md). item</span><span class="sxs-lookup"><span data-stu-id="b89d7-p101">[Office](office.md)[.context](office.context.md)[.mailbox](office.context.mailbox.md). item</span></span>

<span data-ttu-id="b89d7-p102">Der `item`-Namespace wird für den Zugriff auf die aktuell ausgewählte Nachricht, Besprechungsanfrage oder den aktuell ausgewählten Termin verwendet. Sie können den Typ von `item` mithilfe der [itemType](#itemtype-officemailboxenumsitemtypejavascriptapioutlook12officemailboxenumsitemtype)-Eigenschaft bestimmen.</span><span class="sxs-lookup"><span data-stu-id="b89d7-p102">The `item` namespace is used to access the currently selected message, meeting request, or appointment. You can determine the type of the `item` by using the [itemType](#itemtype-officemailboxenumsitemtypejavascriptapioutlook12officemailboxenumsitemtype) property.</span></span>

##### <a name="requirements"></a><span data-ttu-id="b89d7-106">Anforderungen</span><span class="sxs-lookup"><span data-stu-id="b89d7-106">Requirements</span></span>

|<span data-ttu-id="b89d7-107">Anforderung</span><span class="sxs-lookup"><span data-stu-id="b89d7-107">Requirement</span></span>| <span data-ttu-id="b89d7-108">Wert</span><span class="sxs-lookup"><span data-stu-id="b89d7-108">Value</span></span>|
|---|---|
|[<span data-ttu-id="b89d7-109">Mindestversion des Postfachanforderungssatzes</span><span class="sxs-lookup"><span data-stu-id="b89d7-109">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="b89d7-110">1.0</span><span class="sxs-lookup"><span data-stu-id="b89d7-110">1.0</span></span>|
|[<span data-ttu-id="b89d7-111">Mindestberechtigungsstufe</span><span class="sxs-lookup"><span data-stu-id="b89d7-111">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="b89d7-112">Eingeschränkt</span><span class="sxs-lookup"><span data-stu-id="b89d7-112">Restricted</span></span>|
|[<span data-ttu-id="b89d7-113">Zutreffender Outlook-Modus</span><span class="sxs-lookup"><span data-stu-id="b89d7-113">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="b89d7-114">Verfassen oder Lesen</span><span class="sxs-lookup"><span data-stu-id="b89d7-114">Compose or read</span></span>|

### <a name="example"></a><span data-ttu-id="b89d7-115">Beispiel</span><span class="sxs-lookup"><span data-stu-id="b89d7-115">Example</span></span>

<span data-ttu-id="b89d7-116">Im folgenden JavaScript-Codebeispiel wird der Zugriff auf die `subject`-Eigenschaft des aktuellen Elements in Outlook veranschaulicht.</span><span class="sxs-lookup"><span data-stu-id="b89d7-116">The following JavaScript code example shows how to access the `subject` property of the current item in Outlook.</span></span>

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

### <a name="members"></a><span data-ttu-id="b89d7-117">Elemente</span><span class="sxs-lookup"><span data-stu-id="b89d7-117">Members</span></span>

#### <a name="attachments-arrayattachmentdetailsjavascriptapioutlook12officeattachmentdetails"></a><span data-ttu-id="b89d7-118">attachments :Array.<[AttachmentDetails](/javascript/api/outlook_1_2/office.attachmentdetails)></span><span class="sxs-lookup"><span data-stu-id="b89d7-118">attachments :Array.<[AttachmentDetails](/javascript/api/outlook_1_2/office.attachmentdetails)></span></span>

<span data-ttu-id="b89d7-p103">Ruft ein Array mit Anlagen für das Element ab. Nur Lesemodus.</span><span class="sxs-lookup"><span data-stu-id="b89d7-p103">Gets an array of attachments for the item. Read mode only.</span></span>

> [!NOTE]
> <span data-ttu-id="b89d7-121">Bestimmte Dateitypen werden aufgrund von Sicherheitsproblemen von Outlook blockiert und daher nicht zurückgegeben werden.</span><span class="sxs-lookup"><span data-stu-id="b89d7-121">Certain types of files are blocked by Outlook due to potential security issues and are therefore not returned.</span></span> <span data-ttu-id="b89d7-122">Weitere Informationen finden Sie unter [Blockierte Anlagen in Outlook](https://support.office.com/article/Blocked-attachments-in-Outlook-434752E1-02D3-4E90-9124-8B81E49A8519).</span><span class="sxs-lookup"><span data-stu-id="b89d7-122">For more information, see [Blocked attachments in Outlook](https://support.office.com/article/Blocked-attachments-in-Outlook-434752E1-02D3-4E90-9124-8B81E49A8519).</span></span>

##### <a name="type"></a><span data-ttu-id="b89d7-123">Typ:</span><span class="sxs-lookup"><span data-stu-id="b89d7-123">Type:</span></span>

*   <span data-ttu-id="b89d7-124">Array.<[AttachmentDetails](/javascript/api/outlook_1_2/office.attachmentdetails)></span><span class="sxs-lookup"><span data-stu-id="b89d7-124">Array.<[AttachmentDetails](/javascript/api/outlook_1_2/office.attachmentdetails)></span></span>

##### <a name="requirements"></a><span data-ttu-id="b89d7-125">Anforderungen</span><span class="sxs-lookup"><span data-stu-id="b89d7-125">Requirements</span></span>

|<span data-ttu-id="b89d7-126">Anforderung</span><span class="sxs-lookup"><span data-stu-id="b89d7-126">Requirement</span></span>| <span data-ttu-id="b89d7-127">Wert</span><span class="sxs-lookup"><span data-stu-id="b89d7-127">Value</span></span>|
|---|---|
|[<span data-ttu-id="b89d7-128">Mindestversion des Postfachanforderungssatzes</span><span class="sxs-lookup"><span data-stu-id="b89d7-128">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="b89d7-129">1.0</span><span class="sxs-lookup"><span data-stu-id="b89d7-129">1.0</span></span>|
|[<span data-ttu-id="b89d7-130">Mindestberechtigungsstufe</span><span class="sxs-lookup"><span data-stu-id="b89d7-130">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="b89d7-131">ReadItem</span><span class="sxs-lookup"><span data-stu-id="b89d7-131">ReadItem</span></span>|
|[<span data-ttu-id="b89d7-132">Zutreffender Outlook-Modus</span><span class="sxs-lookup"><span data-stu-id="b89d7-132">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="b89d7-133">Lesen</span><span class="sxs-lookup"><span data-stu-id="b89d7-133">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="b89d7-134">Beispiel</span><span class="sxs-lookup"><span data-stu-id="b89d7-134">Example</span></span>

<span data-ttu-id="b89d7-135">Mit dem folgende Code wird eine HTML-Zeichenfolge mit Details aller Anlagen für das aktuelle Element erstellt.</span><span class="sxs-lookup"><span data-stu-id="b89d7-135">The following code builds an HTML string with details of all attachments on the current item.</span></span>

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

####  <a name="bcc-recipientsjavascriptapioutlook12officerecipients"></a><span data-ttu-id="b89d7-136">bcc :[Recipients](/javascript/api/outlook_1_2/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="b89d7-136">bcc :[Recipients](/javascript/api/outlook_1_2/office.recipients)</span></span>

<span data-ttu-id="b89d7-137">Ruft ein Objekt, das Methoden zum Abrufen oder aktualisieren die Empfänger in der Zeile Bcc (blind Carbon Copy, Blindkopie) einer Nachricht bereitstellt.</span><span class="sxs-lookup"><span data-stu-id="b89d7-137">Gets an object that provides methods to get or update the recipients on the Bcc (blind carbon copy) line of a message.</span></span> <span data-ttu-id="b89d7-138">Nur Verfassenmodus.</span><span class="sxs-lookup"><span data-stu-id="b89d7-138">Compose mode only.</span></span>

##### <a name="type"></a><span data-ttu-id="b89d7-139">Typ:</span><span class="sxs-lookup"><span data-stu-id="b89d7-139">Type:</span></span>

*   [<span data-ttu-id="b89d7-140">Recipients</span><span class="sxs-lookup"><span data-stu-id="b89d7-140">Recipients</span></span>](/javascript/api/outlook_1_2/office.recipients)

##### <a name="requirements"></a><span data-ttu-id="b89d7-141">Anforderungen</span><span class="sxs-lookup"><span data-stu-id="b89d7-141">Requirements</span></span>

|<span data-ttu-id="b89d7-142">Anforderung</span><span class="sxs-lookup"><span data-stu-id="b89d7-142">Requirement</span></span>| <span data-ttu-id="b89d7-143">Wert</span><span class="sxs-lookup"><span data-stu-id="b89d7-143">Value</span></span>|
|---|---|
|[<span data-ttu-id="b89d7-144">Mindestversion des Postfachanforderungssatzes</span><span class="sxs-lookup"><span data-stu-id="b89d7-144">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="b89d7-145">1.1</span><span class="sxs-lookup"><span data-stu-id="b89d7-145">1.1</span></span>|
|[<span data-ttu-id="b89d7-146">Mindestberechtigungsstufe</span><span class="sxs-lookup"><span data-stu-id="b89d7-146">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="b89d7-147">ReadItem</span><span class="sxs-lookup"><span data-stu-id="b89d7-147">ReadItem</span></span>|
|[<span data-ttu-id="b89d7-148">Zutreffender Outlook-Modus</span><span class="sxs-lookup"><span data-stu-id="b89d7-148">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="b89d7-149">Verfassen</span><span class="sxs-lookup"><span data-stu-id="b89d7-149">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="b89d7-150">Beispiel</span><span class="sxs-lookup"><span data-stu-id="b89d7-150">Example</span></span>

```JavaScript
Office.context.mailbox.item.bcc.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.bcc.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.bcc.getAsync(callback);

function callback(asyncResult) {
  var arrayOfBccRecipients = asyncResult.value;
}
```

####  <a name="body-bodyjavascriptapioutlook12officebody"></a><span data-ttu-id="b89d7-151">body :[Body](/javascript/api/outlook_1_2/office.body)</span><span class="sxs-lookup"><span data-stu-id="b89d7-151">body :[Body](/javascript/api/outlook_1_2/office.body)</span></span>

<span data-ttu-id="b89d7-152">Ruft ein Objekt ab, das Methoden zum Bearbeiten des Textkörpers eines Elements bereitstellt.</span><span class="sxs-lookup"><span data-stu-id="b89d7-152">Gets an object that provides methods for manipulating the body of an item.</span></span>

##### <a name="type"></a><span data-ttu-id="b89d7-153">Typ:</span><span class="sxs-lookup"><span data-stu-id="b89d7-153">Type:</span></span>

*   [<span data-ttu-id="b89d7-154">Body</span><span class="sxs-lookup"><span data-stu-id="b89d7-154">Body</span></span>](/javascript/api/outlook_1_2/office.body)

##### <a name="requirements"></a><span data-ttu-id="b89d7-155">Anforderungen</span><span class="sxs-lookup"><span data-stu-id="b89d7-155">Requirements</span></span>

|<span data-ttu-id="b89d7-156">Anforderung</span><span class="sxs-lookup"><span data-stu-id="b89d7-156">Requirement</span></span>| <span data-ttu-id="b89d7-157">Wert</span><span class="sxs-lookup"><span data-stu-id="b89d7-157">Value</span></span>|
|---|---|
|[<span data-ttu-id="b89d7-158">Mindestversion des Postfachanforderungssatzes</span><span class="sxs-lookup"><span data-stu-id="b89d7-158">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="b89d7-159">1.1</span><span class="sxs-lookup"><span data-stu-id="b89d7-159">1.1</span></span>|
|[<span data-ttu-id="b89d7-160">Mindestberechtigungsstufe</span><span class="sxs-lookup"><span data-stu-id="b89d7-160">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="b89d7-161">ReadItem</span><span class="sxs-lookup"><span data-stu-id="b89d7-161">ReadItem</span></span>|
|[<span data-ttu-id="b89d7-162">Zutreffender Outlook-Modus</span><span class="sxs-lookup"><span data-stu-id="b89d7-162">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="b89d7-163">Verfassen oder Lesen</span><span class="sxs-lookup"><span data-stu-id="b89d7-163">Compose or read</span></span>|

####  <a name="cc-arrayemailaddressdetailsjavascriptapioutlook12officeemailaddressdetailsrecipientsjavascriptapioutlook12officerecipients"></a><span data-ttu-id="b89d7-164">cc: Array. <[EmailAddressDetails](/javascript/api/outlook_1_2/office.emailaddressdetails)>|[Empfänger](/javascript/api/outlook_1_2/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="b89d7-164">cc :Array.<[EmailAddressDetails](/javascript/api/outlook_1_2/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook_1_2/office.recipients)</span></span>

<span data-ttu-id="b89d7-165">Ermöglicht den Zugriff auf die (Carbon Copy, Blindkopie) Cc-Empfänger einer Nachricht.</span><span class="sxs-lookup"><span data-stu-id="b89d7-165">Provides access to the Cc (carbon copy) recipients of a message.</span></span> <span data-ttu-id="b89d7-166">Der Typ des Objekts und die Zugriffsebene, hängt von den Modus des aktuellen Elements ab.</span><span class="sxs-lookup"><span data-stu-id="b89d7-166">The type of object and level of access depends on the mode of the current item.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="b89d7-167">Lesemodus</span><span class="sxs-lookup"><span data-stu-id="b89d7-167">Read mode</span></span>

<span data-ttu-id="b89d7-p107">Die `cc`-Eigenschaft gibt ein Array mit einem `EmailAddressDetails`-Objekt für jeden Empfänger in der **Cc**-Zeile der Nachricht zurück. Die Auflistung ist auf höchstens 100 Elemente beschränkt.</span><span class="sxs-lookup"><span data-stu-id="b89d7-p107">The `cc` property returns an array that contains an `EmailAddressDetails` object for each recipient listed on the **Cc** line of the message. The collection is limited to a maximum of 100 members.</span></span>

##### <a name="compose-mode"></a><span data-ttu-id="b89d7-170">Verfassenmodus</span><span class="sxs-lookup"><span data-stu-id="b89d7-170">Compose mode</span></span>

<span data-ttu-id="b89d7-171">Die `cc` -Eigenschaft gibt eine `Recipients` -Objekt, das Methoden zum Abrufen oder aktualisieren die Empfänger in der Zeile **Cc** der Nachricht bereitstellt.</span><span class="sxs-lookup"><span data-stu-id="b89d7-171">The `cc` property returns a `Recipients` object that provides methods to get or update the recipients on the **Cc** line of the message.</span></span>

##### <a name="type"></a><span data-ttu-id="b89d7-172">Typ:</span><span class="sxs-lookup"><span data-stu-id="b89d7-172">Type:</span></span>

*   <span data-ttu-id="b89d7-173">Array.<[EmailAddressDetails](/javascript/api/outlook_1_2/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook_1_2/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="b89d7-173">Array.<[EmailAddressDetails](/javascript/api/outlook_1_2/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook_1_2/office.recipients)</span></span>

##### <a name="requirements"></a><span data-ttu-id="b89d7-174">Anforderungen</span><span class="sxs-lookup"><span data-stu-id="b89d7-174">Requirements</span></span>

|<span data-ttu-id="b89d7-175">Anforderung</span><span class="sxs-lookup"><span data-stu-id="b89d7-175">Requirement</span></span>| <span data-ttu-id="b89d7-176">Wert</span><span class="sxs-lookup"><span data-stu-id="b89d7-176">Value</span></span>|
|---|---|
|[<span data-ttu-id="b89d7-177">Mindestversion des Postfachanforderungssatzes</span><span class="sxs-lookup"><span data-stu-id="b89d7-177">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="b89d7-178">1.0</span><span class="sxs-lookup"><span data-stu-id="b89d7-178">1.0</span></span>|
|[<span data-ttu-id="b89d7-179">Mindestberechtigungsstufe</span><span class="sxs-lookup"><span data-stu-id="b89d7-179">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="b89d7-180">ReadItem</span><span class="sxs-lookup"><span data-stu-id="b89d7-180">ReadItem</span></span>|
|[<span data-ttu-id="b89d7-181">Zutreffender Outlook-Modus</span><span class="sxs-lookup"><span data-stu-id="b89d7-181">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="b89d7-182">Verfassen oder Lesen</span><span class="sxs-lookup"><span data-stu-id="b89d7-182">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="b89d7-183">Beispiel</span><span class="sxs-lookup"><span data-stu-id="b89d7-183">Example</span></span>

```JavaScript
Office.context.mailbox.item.cc.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.cc.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.cc.getAsync(callback);

function callback(asyncResult) {
  var arrayOfCcRecipients = asyncResult.value;
}
```

####  <a name="nullable-conversationid-string"></a><span data-ttu-id="b89d7-184">(nullable) conversationId :String</span><span class="sxs-lookup"><span data-stu-id="b89d7-184">(nullable) conversationId :String</span></span>

<span data-ttu-id="b89d7-185">Ruft einen Bezeichner für die E-Mail-Unterhaltung ab, in der eine bestimmte Nachricht enthalten ist.</span><span class="sxs-lookup"><span data-stu-id="b89d7-185">Gets an identifier for the email conversation that contains a particular message.</span></span>

<span data-ttu-id="b89d7-p108">Sie können für diese Eigenschaft eine ganze Zahl abrufen, wenn Ihre Mail-App in Formularen zum Lesen oder Antworten in Formularen zum Verfassen aktiviert wird. Wenn der Benutzer den Betreff der Antwortnachricht ändert, ändert sich beim Versenden die Konversations-ID für die entsprechende Nachricht, und der Wert, den Sie vorher bezogen haben, trifft nicht länger zu.</span><span class="sxs-lookup"><span data-stu-id="b89d7-p108">You can get an integer for this property if your mail app is activated in read forms or responses in compose forms. If subsequently the user changes the subject of the reply message, upon sending the reply, the conversation ID for that message will change and that value you obtained earlier will no longer apply.</span></span>

<span data-ttu-id="b89d7-p109">Sie erhalten in einem Formular zum Verfassen für diese Eigenschaft für ein neues Element null. Wenn der Benutzer einen Betreff festlegt und das Element speichert, gibt die `conversationId`-Eigenschaft einen Wert zurück.</span><span class="sxs-lookup"><span data-stu-id="b89d7-p109">You get null for this property for a new item in a compose form. If the user sets a subject and saves the item, the `conversationId` property will return a value.</span></span>

##### <a name="type"></a><span data-ttu-id="b89d7-190">Typ:</span><span class="sxs-lookup"><span data-stu-id="b89d7-190">Type:</span></span>

*   <span data-ttu-id="b89d7-191">String</span><span class="sxs-lookup"><span data-stu-id="b89d7-191">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="b89d7-192">Anforderungen</span><span class="sxs-lookup"><span data-stu-id="b89d7-192">Requirements</span></span>

|<span data-ttu-id="b89d7-193">Anforderung</span><span class="sxs-lookup"><span data-stu-id="b89d7-193">Requirement</span></span>| <span data-ttu-id="b89d7-194">Wert</span><span class="sxs-lookup"><span data-stu-id="b89d7-194">Value</span></span>|
|---|---|
|[<span data-ttu-id="b89d7-195">Mindestversion des Postfachanforderungssatzes</span><span class="sxs-lookup"><span data-stu-id="b89d7-195">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="b89d7-196">1.0</span><span class="sxs-lookup"><span data-stu-id="b89d7-196">1.0</span></span>|
|[<span data-ttu-id="b89d7-197">Mindestberechtigungsstufe</span><span class="sxs-lookup"><span data-stu-id="b89d7-197">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="b89d7-198">ReadItem</span><span class="sxs-lookup"><span data-stu-id="b89d7-198">ReadItem</span></span>|
|[<span data-ttu-id="b89d7-199">Zutreffender Outlook-Modus</span><span class="sxs-lookup"><span data-stu-id="b89d7-199">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="b89d7-200">Verfassen oder Lesen</span><span class="sxs-lookup"><span data-stu-id="b89d7-200">Compose or read</span></span>|

#### <a name="datetimecreated-date"></a><span data-ttu-id="b89d7-201">dateTimeCreated :Date</span><span class="sxs-lookup"><span data-stu-id="b89d7-201">dateTimeCreated :Date</span></span>

<span data-ttu-id="b89d7-p110">Ruft das Datum und die Uhrzeit der Erstellung eines Elements ab. Nur Lesemodus.</span><span class="sxs-lookup"><span data-stu-id="b89d7-p110">Gets the date and time that an item was created. Read mode only.</span></span>

##### <a name="type"></a><span data-ttu-id="b89d7-204">Typ:</span><span class="sxs-lookup"><span data-stu-id="b89d7-204">Type:</span></span>

*   <span data-ttu-id="b89d7-205">Datum</span><span class="sxs-lookup"><span data-stu-id="b89d7-205">Date</span></span>

##### <a name="requirements"></a><span data-ttu-id="b89d7-206">Anforderungen</span><span class="sxs-lookup"><span data-stu-id="b89d7-206">Requirements</span></span>

|<span data-ttu-id="b89d7-207">Anforderung</span><span class="sxs-lookup"><span data-stu-id="b89d7-207">Requirement</span></span>| <span data-ttu-id="b89d7-208">Wert</span><span class="sxs-lookup"><span data-stu-id="b89d7-208">Value</span></span>|
|---|---|
|[<span data-ttu-id="b89d7-209">Mindestversion des Postfachanforderungssatzes</span><span class="sxs-lookup"><span data-stu-id="b89d7-209">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="b89d7-210">1.0</span><span class="sxs-lookup"><span data-stu-id="b89d7-210">1.0</span></span>|
|[<span data-ttu-id="b89d7-211">Mindestberechtigungsstufe</span><span class="sxs-lookup"><span data-stu-id="b89d7-211">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="b89d7-212">ReadItem</span><span class="sxs-lookup"><span data-stu-id="b89d7-212">ReadItem</span></span>|
|[<span data-ttu-id="b89d7-213">Zutreffender Outlook-Modus</span><span class="sxs-lookup"><span data-stu-id="b89d7-213">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="b89d7-214">Lesen</span><span class="sxs-lookup"><span data-stu-id="b89d7-214">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="b89d7-215">Beispiel</span><span class="sxs-lookup"><span data-stu-id="b89d7-215">Example</span></span>

```JavaScript
var created = Office.context.mailbox.item.dateTimeCreated;
```

#### <a name="datetimemodified-date"></a><span data-ttu-id="b89d7-216">dateTimeModified :Date</span><span class="sxs-lookup"><span data-stu-id="b89d7-216">dateTimeModified :Date</span></span>

<span data-ttu-id="b89d7-p111">Ruft das Datum und die Uhrzeit der letzten Änderung eines Elements ab. Nur Lesemodus.</span><span class="sxs-lookup"><span data-stu-id="b89d7-p111">Gets the date and time that an item was last modified. Read mode only.</span></span>

> [!NOTE]
> <span data-ttu-id="b89d7-219">Dieser Member wird in Outlook für iOS oder Outlook für Android nicht unterstützt.</span><span class="sxs-lookup"><span data-stu-id="b89d7-219">This member is not supported in Outlook for iOS or Outlook for Android.</span></span>

##### <a name="type"></a><span data-ttu-id="b89d7-220">Typ:</span><span class="sxs-lookup"><span data-stu-id="b89d7-220">Type:</span></span>

*   <span data-ttu-id="b89d7-221">Datum</span><span class="sxs-lookup"><span data-stu-id="b89d7-221">Date</span></span>

##### <a name="requirements"></a><span data-ttu-id="b89d7-222">Anforderungen</span><span class="sxs-lookup"><span data-stu-id="b89d7-222">Requirements</span></span>

|<span data-ttu-id="b89d7-223">Anforderung</span><span class="sxs-lookup"><span data-stu-id="b89d7-223">Requirement</span></span>| <span data-ttu-id="b89d7-224">Wert</span><span class="sxs-lookup"><span data-stu-id="b89d7-224">Value</span></span>|
|---|---|
|[<span data-ttu-id="b89d7-225">Mindestversion des Postfachanforderungssatzes</span><span class="sxs-lookup"><span data-stu-id="b89d7-225">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="b89d7-226">1.0</span><span class="sxs-lookup"><span data-stu-id="b89d7-226">1.0</span></span>|
|[<span data-ttu-id="b89d7-227">Mindestberechtigungsstufe</span><span class="sxs-lookup"><span data-stu-id="b89d7-227">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="b89d7-228">ReadItem</span><span class="sxs-lookup"><span data-stu-id="b89d7-228">ReadItem</span></span>|
|[<span data-ttu-id="b89d7-229">Zutreffender Outlook-Modus</span><span class="sxs-lookup"><span data-stu-id="b89d7-229">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="b89d7-230">Lesen</span><span class="sxs-lookup"><span data-stu-id="b89d7-230">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="b89d7-231">Beispiel</span><span class="sxs-lookup"><span data-stu-id="b89d7-231">Example</span></span>

```JavaScript
var modified = Office.context.mailbox.item.dateTimeModified;
```

####  <a name="end-datetimejavascriptapioutlook12officetime"></a><span data-ttu-id="b89d7-232">end :Date|[Time](/javascript/api/outlook_1_2/office.time)</span><span class="sxs-lookup"><span data-stu-id="b89d7-232">end :Date|[Time](/javascript/api/outlook_1_2/office.time)</span></span>

<span data-ttu-id="b89d7-233">Ruft Datum und Zeit für das Ende des Termins ab oder legt diese fest.</span><span class="sxs-lookup"><span data-stu-id="b89d7-233">Gets or sets the date and time that the appointment is to end.</span></span>

<span data-ttu-id="b89d7-p112">Die `end`-Eigenschaft wird als Datums- und Uhrzeitwert in UTC ausgedrückt. Sie können die [`convertToLocalClientTime`](office.context.mailbox.md#converttolocalclienttimetimevalue--localclienttimejavascriptapioutlook12officelocalclienttime)-Methode verwenden, um den Endeigenschaftswert in den lokalen Uhrzeit- und Datumswert des Clients umzuwandeln.</span><span class="sxs-lookup"><span data-stu-id="b89d7-p112">The `end` property is expressed as a Coordinated Universal Time (UTC) date and time value. You can use the [`convertToLocalClientTime`](office.context.mailbox.md#converttolocalclienttimetimevalue--localclienttimejavascriptapioutlook12officelocalclienttime) method to convert the end property value to the client’s local date and time.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="b89d7-236">Lesemodus</span><span class="sxs-lookup"><span data-stu-id="b89d7-236">Read mode</span></span>

<span data-ttu-id="b89d7-237">Die `end`-Eigenschaft gibt ein `Date`-Objekt zurück.</span><span class="sxs-lookup"><span data-stu-id="b89d7-237">The `end` property returns a `Date` object.</span></span>

##### <a name="compose-mode"></a><span data-ttu-id="b89d7-238">Verfassenmodus</span><span class="sxs-lookup"><span data-stu-id="b89d7-238">Compose mode</span></span>

<span data-ttu-id="b89d7-239">Die `end`-Eigenschaft gibt ein `Time`-Objekt zurück.</span><span class="sxs-lookup"><span data-stu-id="b89d7-239">The `end` property returns a `Time` object.</span></span>

<span data-ttu-id="b89d7-240">Wenn Sie die [`Time.setAsync`](/javascript/api/outlook_1_2/office.time#setasync-datetime--options--callback-)-Methode verwenden, um die Endzeit im Verfassenmodus festzulegen, sollten Sie die [`convertToUtcClientTime`](office.context.mailbox.md#converttoutcclienttimeinput--date)-Methode verwenden, um die Ortszeit auf dem Client für den Server in UTC umzuwandeln.</span><span class="sxs-lookup"><span data-stu-id="b89d7-240">When you use the [`Time.setAsync`](/javascript/api/outlook_1_2/office.time#setasync-datetime--options--callback-) method to set the end time, you should use the [`convertToUtcClientTime`](office.context.mailbox.md#converttoutcclienttimeinput--date) method to convert the local time on the client to UTC for the server.</span></span>

##### <a name="type"></a><span data-ttu-id="b89d7-241">Typ:</span><span class="sxs-lookup"><span data-stu-id="b89d7-241">Type:</span></span>

*   <span data-ttu-id="b89d7-242">Date | [Time](/javascript/api/outlook_1_2/office.time)</span><span class="sxs-lookup"><span data-stu-id="b89d7-242">Date | [Time](/javascript/api/outlook_1_2/office.time)</span></span>

##### <a name="requirements"></a><span data-ttu-id="b89d7-243">Anforderungen</span><span class="sxs-lookup"><span data-stu-id="b89d7-243">Requirements</span></span>

|<span data-ttu-id="b89d7-244">Anforderung</span><span class="sxs-lookup"><span data-stu-id="b89d7-244">Requirement</span></span>| <span data-ttu-id="b89d7-245">Wert</span><span class="sxs-lookup"><span data-stu-id="b89d7-245">Value</span></span>|
|---|---|
|[<span data-ttu-id="b89d7-246">Mindestversion des Postfachanforderungssatzes</span><span class="sxs-lookup"><span data-stu-id="b89d7-246">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="b89d7-247">1.0</span><span class="sxs-lookup"><span data-stu-id="b89d7-247">1.0</span></span>|
|[<span data-ttu-id="b89d7-248">Mindestberechtigungsstufe</span><span class="sxs-lookup"><span data-stu-id="b89d7-248">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="b89d7-249">ReadItem</span><span class="sxs-lookup"><span data-stu-id="b89d7-249">ReadItem</span></span>|
|[<span data-ttu-id="b89d7-250">Zutreffender Outlook-Modus</span><span class="sxs-lookup"><span data-stu-id="b89d7-250">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="b89d7-251">Verfassen oder Lesen</span><span class="sxs-lookup"><span data-stu-id="b89d7-251">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="b89d7-252">Beispiel</span><span class="sxs-lookup"><span data-stu-id="b89d7-252">Example</span></span>

<span data-ttu-id="b89d7-253">Im folgenden Beispiel wird die Endzeit eines Termins im Verfassenmodus mithilfe der [`setAsync`](/javascript/api/outlook_1_2/office.time#setasync-datetime--options--callback-)-Methode des `Time`-Objekts festgelegt.</span><span class="sxs-lookup"><span data-stu-id="b89d7-253">The following example sets the end time of an appointment in compose mode by using the [`setAsync`](/javascript/api/outlook_1_2/office.time#setasync-datetime--options--callback-) method of the `Time` object.</span></span>

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

#### <a name="from-emailaddressdetailsjavascriptapioutlook12officeemailaddressdetails"></a><span data-ttu-id="b89d7-254">from :[EmailAddressDetails](/javascript/api/outlook_1_2/office.emailaddressdetails)</span><span class="sxs-lookup"><span data-stu-id="b89d7-254">from :[EmailAddressDetails](/javascript/api/outlook_1_2/office.emailaddressdetails)</span></span>

<span data-ttu-id="b89d7-p113">Ruft die E-Mail-Adresse des Absenders einer Nachricht ab. Nur Lesemodus.</span><span class="sxs-lookup"><span data-stu-id="b89d7-p113">Gets the email address of the sender of a message. Read mode only.</span></span>

<span data-ttu-id="b89d7-p114">Die `from`- und [`sender`](#sender-emailaddressdetailsjavascriptapioutlook12officeemailaddressdetails)-Eigenschaften stellen dieselbe Person dar, außer die Nachricht wurde von einem Delegaten gesendet. In diesem Fall stellt die `from`-Eigenschaft die Stellvertretung dar, und die sender-Eigenschaft stellt den Delegaten dar.</span><span class="sxs-lookup"><span data-stu-id="b89d7-p114">The `from` and [`sender`](#sender-emailaddressdetailsjavascriptapioutlook12officeemailaddressdetails) properties represent the same person unless the message is sent by a delegate. In that case, the `from` property represents the delegator, and the sender property represents the delegate.</span></span>

> [!NOTE]
> <span data-ttu-id="b89d7-259">Die `recipientType` -Eigenschaft des der `EmailAddressDetails` -Objekts der `from` -Eigenschaft ist `undefined`.</span><span class="sxs-lookup"><span data-stu-id="b89d7-259">The `recipientType` property of the `EmailAddressDetails` object in the `from` property is `undefined`.</span></span>

##### <a name="type"></a><span data-ttu-id="b89d7-260">Typ:</span><span class="sxs-lookup"><span data-stu-id="b89d7-260">Type:</span></span>

*   [<span data-ttu-id="b89d7-261">EmailAddressDetails</span><span class="sxs-lookup"><span data-stu-id="b89d7-261">EmailAddressDetails</span></span>](/javascript/api/outlook_1_2/office.emailaddressdetails)

##### <a name="requirements"></a><span data-ttu-id="b89d7-262">Anforderungen</span><span class="sxs-lookup"><span data-stu-id="b89d7-262">Requirements</span></span>

|<span data-ttu-id="b89d7-263">Anforderung</span><span class="sxs-lookup"><span data-stu-id="b89d7-263">Requirement</span></span>| <span data-ttu-id="b89d7-264">Wert</span><span class="sxs-lookup"><span data-stu-id="b89d7-264">Value</span></span>|
|---|---|
|[<span data-ttu-id="b89d7-265">Mindestversion des Postfachanforderungssatzes</span><span class="sxs-lookup"><span data-stu-id="b89d7-265">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="b89d7-266">1.0</span><span class="sxs-lookup"><span data-stu-id="b89d7-266">1.0</span></span>|
|[<span data-ttu-id="b89d7-267">Mindestberechtigungsstufe</span><span class="sxs-lookup"><span data-stu-id="b89d7-267">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="b89d7-268">ReadItem</span><span class="sxs-lookup"><span data-stu-id="b89d7-268">ReadItem</span></span>|
|[<span data-ttu-id="b89d7-269">Zutreffender Outlook-Modus</span><span class="sxs-lookup"><span data-stu-id="b89d7-269">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="b89d7-270">Lesen</span><span class="sxs-lookup"><span data-stu-id="b89d7-270">Read</span></span>|

#### <a name="internetmessageid-string"></a><span data-ttu-id="b89d7-271">internetMessageId :String</span><span class="sxs-lookup"><span data-stu-id="b89d7-271">internetMessageId :String</span></span>

<span data-ttu-id="b89d7-p115">Ruft die Internetnachricht-ID einer E-Mail-Nachricht ab. Nur Lesemodus.</span><span class="sxs-lookup"><span data-stu-id="b89d7-p115">Gets the Internet message identifier for an email message. Read mode only.</span></span>

##### <a name="type"></a><span data-ttu-id="b89d7-274">Typ:</span><span class="sxs-lookup"><span data-stu-id="b89d7-274">Type:</span></span>

*   <span data-ttu-id="b89d7-275">String</span><span class="sxs-lookup"><span data-stu-id="b89d7-275">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="b89d7-276">Anforderungen</span><span class="sxs-lookup"><span data-stu-id="b89d7-276">Requirements</span></span>

|<span data-ttu-id="b89d7-277">Anforderung</span><span class="sxs-lookup"><span data-stu-id="b89d7-277">Requirement</span></span>| <span data-ttu-id="b89d7-278">Wert</span><span class="sxs-lookup"><span data-stu-id="b89d7-278">Value</span></span>|
|---|---|
|[<span data-ttu-id="b89d7-279">Mindestversion des Postfachanforderungssatzes</span><span class="sxs-lookup"><span data-stu-id="b89d7-279">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="b89d7-280">1.0</span><span class="sxs-lookup"><span data-stu-id="b89d7-280">1.0</span></span>|
|[<span data-ttu-id="b89d7-281">Mindestberechtigungsstufe</span><span class="sxs-lookup"><span data-stu-id="b89d7-281">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="b89d7-282">ReadItem</span><span class="sxs-lookup"><span data-stu-id="b89d7-282">ReadItem</span></span>|
|[<span data-ttu-id="b89d7-283">Zutreffender Outlook-Modus</span><span class="sxs-lookup"><span data-stu-id="b89d7-283">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="b89d7-284">Lesen</span><span class="sxs-lookup"><span data-stu-id="b89d7-284">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="b89d7-285">Beispiel</span><span class="sxs-lookup"><span data-stu-id="b89d7-285">Example</span></span>

```JavaScript
var internetMessageId = Office.context.mailbox.item.internetMessageId;
```

#### <a name="itemclass-string"></a><span data-ttu-id="b89d7-286">itemClass :String</span><span class="sxs-lookup"><span data-stu-id="b89d7-286">itemClass :String</span></span>

<span data-ttu-id="b89d7-p116">Ruft die Exchange-Webdienste-Elementklasse des ausgewählten Elements ab. Nur Lesemodus.</span><span class="sxs-lookup"><span data-stu-id="b89d7-p116">Gets the Exchange Web Services item class of the selected item. Read mode only.</span></span>

<span data-ttu-id="b89d7-p117">Die `itemClass`-Eigenschaft gibt die Nachrichtenklasse des ausgewählten Elements an. Folgende Standardnachrichtenklassen für das Nachrichten- oder Terminelement sind vorhanden:</span><span class="sxs-lookup"><span data-stu-id="b89d7-p117">The `itemClass` property specifies the message class of the selected item. The following are the default message classes for the message or appointment item.</span></span>

| <span data-ttu-id="b89d7-291">Typ</span><span class="sxs-lookup"><span data-stu-id="b89d7-291">Type</span></span> | <span data-ttu-id="b89d7-292">Beschreibung</span><span class="sxs-lookup"><span data-stu-id="b89d7-292">Description</span></span> | <span data-ttu-id="b89d7-293">Elementklasse</span><span class="sxs-lookup"><span data-stu-id="b89d7-293">item class</span></span> |
| --- | --- | --- |
| <span data-ttu-id="b89d7-294">Terminelemente</span><span class="sxs-lookup"><span data-stu-id="b89d7-294">Appointment items</span></span> | <span data-ttu-id="b89d7-295">Dies sind Kalenderelemente der Elementklasse `IPM.Appointment` oder `IPM.Appointment.Occurence`.</span><span class="sxs-lookup"><span data-stu-id="b89d7-295">These are calendar items of the item class `IPM.Appointment` or `IPM.Appointment.Occurence`.</span></span> | `IPM.Appointment`<br />`IPM.Appointment.Occurence` |
| <span data-ttu-id="b89d7-296">Nachrichtenelemente</span><span class="sxs-lookup"><span data-stu-id="b89d7-296">Message items</span></span> | <span data-ttu-id="b89d7-297">Diese Elemente umfassen E-Mail-Nachrichten der Standardnachrichtenklasse `IPM.Note` sowie Besprechungsanfragen, -antworten und -absagen, die `IPM.Schedule.Meeting` als Basisnachrichtenklasse verwenden.</span><span class="sxs-lookup"><span data-stu-id="b89d7-297">These include email messages that have the default message class `IPM.Note`, and meeting requests, responses, and cancellations, that use `IPM.Schedule.Meeting` as the base message class.</span></span> | `IPM.Note`<br />`IPM.Schedule.Meeting.Request`<br />`IPM.Schedule.Meeting.Neg`<br />`IPM.Schedule.Meeting.Pos`<br />`IPM.Schedule.Meeting.Tent`<br />`IPM.Schedule.Meeting.Canceled` |

<span data-ttu-id="b89d7-298">Sie können benutzerdefinierte Nachrichtenklassen zum Erweitern einer Standardnachrichtenklasse erstellen, z. B. eine benutzerdefinierte Terminnachrichtenklasse wie `IPM.Appointment.Contoso`.</span><span class="sxs-lookup"><span data-stu-id="b89d7-298">You can create custom message classes that extends a default message class, for example, a custom appointment message class `IPM.Appointment.Contoso`.</span></span>

##### <a name="type"></a><span data-ttu-id="b89d7-299">Typ:</span><span class="sxs-lookup"><span data-stu-id="b89d7-299">Type:</span></span>

*   <span data-ttu-id="b89d7-300">String</span><span class="sxs-lookup"><span data-stu-id="b89d7-300">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="b89d7-301">Anforderungen</span><span class="sxs-lookup"><span data-stu-id="b89d7-301">Requirements</span></span>

|<span data-ttu-id="b89d7-302">Anforderung</span><span class="sxs-lookup"><span data-stu-id="b89d7-302">Requirement</span></span>| <span data-ttu-id="b89d7-303">Wert</span><span class="sxs-lookup"><span data-stu-id="b89d7-303">Value</span></span>|
|---|---|
|[<span data-ttu-id="b89d7-304">Mindestversion des Postfachanforderungssatzes</span><span class="sxs-lookup"><span data-stu-id="b89d7-304">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="b89d7-305">1.0</span><span class="sxs-lookup"><span data-stu-id="b89d7-305">1.0</span></span>|
|[<span data-ttu-id="b89d7-306">Mindestberechtigungsstufe</span><span class="sxs-lookup"><span data-stu-id="b89d7-306">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="b89d7-307">ReadItem</span><span class="sxs-lookup"><span data-stu-id="b89d7-307">ReadItem</span></span>|
|[<span data-ttu-id="b89d7-308">Zutreffender Outlook-Modus</span><span class="sxs-lookup"><span data-stu-id="b89d7-308">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="b89d7-309">Lesen</span><span class="sxs-lookup"><span data-stu-id="b89d7-309">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="b89d7-310">Beispiel</span><span class="sxs-lookup"><span data-stu-id="b89d7-310">Example</span></span>

```JavaScript
var itemClass = Office.context.mailbox.item.itemClass;
```

#### <a name="nullable-itemid-string"></a><span data-ttu-id="b89d7-311">(nullable) itemId :String</span><span class="sxs-lookup"><span data-stu-id="b89d7-311">(nullable) itemId :String</span></span>

<span data-ttu-id="b89d7-p118">Ruft den Exchange-Webdienste-Elementbezeichner für das aktuelle Element ab. Nur Lesemodus.</span><span class="sxs-lookup"><span data-stu-id="b89d7-p118">Gets the Exchange Web Services item identifier for the current item. Read mode only.</span></span>

> [!NOTE]
> <span data-ttu-id="b89d7-314">Der Bezeichner, der von der `itemId`-Eigenschaft zurückgegeben wird, ist der gleiche Bezeichner wie der Exchange-Webdienste-Elementbezeichner.</span><span class="sxs-lookup"><span data-stu-id="b89d7-314">The identifier returned by the `itemId` property is the same as the Exchange Web Services item identifier.</span></span> <span data-ttu-id="b89d7-315">Die `itemId` -Eigenschaft ist nicht identisch mit der Outlook-Eintrags-ID oder die ID von der Outlook-REST-API verwendet.</span><span class="sxs-lookup"><span data-stu-id="b89d7-315">The `itemId` property is not identical to the Outlook Entry ID or the ID used by the Outlook REST API.</span></span> <span data-ttu-id="b89d7-316">Vor dem REST API-Aufrufe mit diesem Wert vornehmen, es konvertiert werden soll mit `Office.context.mailbox.convertToRestId`, die verfügbar und beginnen bei Anforderung ist 1.3 festgelegt.</span><span class="sxs-lookup"><span data-stu-id="b89d7-316">Before making REST API calls using this value, it should be converted using `Office.context.mailbox.convertToRestId`, which is available starting in requirement set 1.3.</span></span> <span data-ttu-id="b89d7-317">Weitere Informationen finden Sie unter [Verwenden der Outlook-REST-APIs aus einem Outlook-add-in](https://docs.microsoft.com/outlook/add-ins/use-rest-api#get-the-item-id).</span><span class="sxs-lookup"><span data-stu-id="b89d7-317">For more details, see [Use the Outlook REST APIs from an Outlook add-in](https://docs.microsoft.com/outlook/add-ins/use-rest-api#get-the-item-id).</span></span>

##### <a name="type"></a><span data-ttu-id="b89d7-318">Typ:</span><span class="sxs-lookup"><span data-stu-id="b89d7-318">Type:</span></span>

*   <span data-ttu-id="b89d7-319">String</span><span class="sxs-lookup"><span data-stu-id="b89d7-319">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="b89d7-320">Anforderungen</span><span class="sxs-lookup"><span data-stu-id="b89d7-320">Requirements</span></span>

|<span data-ttu-id="b89d7-321">Anforderung</span><span class="sxs-lookup"><span data-stu-id="b89d7-321">Requirement</span></span>| <span data-ttu-id="b89d7-322">Wert</span><span class="sxs-lookup"><span data-stu-id="b89d7-322">Value</span></span>|
|---|---|
|[<span data-ttu-id="b89d7-323">Mindestversion des Postfachanforderungssatzes</span><span class="sxs-lookup"><span data-stu-id="b89d7-323">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="b89d7-324">1.0</span><span class="sxs-lookup"><span data-stu-id="b89d7-324">1.0</span></span>|
|[<span data-ttu-id="b89d7-325">Mindestberechtigungsstufe</span><span class="sxs-lookup"><span data-stu-id="b89d7-325">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="b89d7-326">ReadItem</span><span class="sxs-lookup"><span data-stu-id="b89d7-326">ReadItem</span></span>|
|[<span data-ttu-id="b89d7-327">Zutreffender Outlook-Modus</span><span class="sxs-lookup"><span data-stu-id="b89d7-327">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="b89d7-328">Lesen</span><span class="sxs-lookup"><span data-stu-id="b89d7-328">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="b89d7-329">Beispiel</span><span class="sxs-lookup"><span data-stu-id="b89d7-329">Example</span></span>

<span data-ttu-id="b89d7-p120">Mit dem folgende Code wird das Vorhandensein eines Elementbezeichners überprüft. Wenn die `itemId`-Eigenschaft `null` oder `undefined` zurückgibt, speichert sie das Element und ruft den Elementbezeichner aus dem asynchronen Ergebnis ab.</span><span class="sxs-lookup"><span data-stu-id="b89d7-p120">The following code checks for the presence of an item identifier. If the `itemId` property returns `null` or `undefined`, it saves the item to the store and gets the item identifier from the asynchronous result.</span></span>

```JavaScript
var itemId = Office.context.mailbox.item.itemId;
if (itemId === null || itemId == undefined) {
  Office.context.mailbox.item.saveAsync(function(result){
    itemId = result.value;
  });
}
```

####  <a name="itemtype-officemailboxenumsitemtypejavascriptapioutlook12officemailboxenumsitemtype"></a><span data-ttu-id="b89d7-332">itemType :[Office.MailboxEnums.ItemType](/javascript/api/outlook_1_2/office.mailboxenums.itemtype)</span><span class="sxs-lookup"><span data-stu-id="b89d7-332">itemType :[Office.MailboxEnums.ItemType](/javascript/api/outlook_1_2/office.mailboxenums.itemtype)</span></span>

<span data-ttu-id="b89d7-333">Ruft den Typ des Elements ab, das eine Instanz darstellt.</span><span class="sxs-lookup"><span data-stu-id="b89d7-333">Gets the type of item that an instance represents.</span></span>

<span data-ttu-id="b89d7-334">Die `itemType`-Eigenschaft gibt einen der Werte der `ItemType`-Enumeration zurück, der angibt, ob es sich bei der `item`-Objektinstanz um eine Nachricht oder einen Termin handelt.</span><span class="sxs-lookup"><span data-stu-id="b89d7-334">The `itemType` property returns one of the `ItemType` enumeration values, indicating whether the `item` object instance is a message or an appointment.</span></span>

##### <a name="type"></a><span data-ttu-id="b89d7-335">Typ:</span><span class="sxs-lookup"><span data-stu-id="b89d7-335">Type:</span></span>

*   [<span data-ttu-id="b89d7-336">Office.MailboxEnums.ItemType</span><span class="sxs-lookup"><span data-stu-id="b89d7-336">Office.MailboxEnums.ItemType</span></span>](/javascript/api/outlook_1_2/office.mailboxenums.itemtype)

##### <a name="requirements"></a><span data-ttu-id="b89d7-337">Anforderungen</span><span class="sxs-lookup"><span data-stu-id="b89d7-337">Requirements</span></span>

|<span data-ttu-id="b89d7-338">Anforderung</span><span class="sxs-lookup"><span data-stu-id="b89d7-338">Requirement</span></span>| <span data-ttu-id="b89d7-339">Wert</span><span class="sxs-lookup"><span data-stu-id="b89d7-339">Value</span></span>|
|---|---|
|[<span data-ttu-id="b89d7-340">Mindestversion des Postfachanforderungssatzes</span><span class="sxs-lookup"><span data-stu-id="b89d7-340">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="b89d7-341">1.0</span><span class="sxs-lookup"><span data-stu-id="b89d7-341">1.0</span></span>|
|[<span data-ttu-id="b89d7-342">Mindestberechtigungsstufe</span><span class="sxs-lookup"><span data-stu-id="b89d7-342">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="b89d7-343">ReadItem</span><span class="sxs-lookup"><span data-stu-id="b89d7-343">ReadItem</span></span>|
|[<span data-ttu-id="b89d7-344">Zutreffender Outlook-Modus</span><span class="sxs-lookup"><span data-stu-id="b89d7-344">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="b89d7-345">Verfassen oder Lesen</span><span class="sxs-lookup"><span data-stu-id="b89d7-345">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="b89d7-346">Beispiel</span><span class="sxs-lookup"><span data-stu-id="b89d7-346">Example</span></span>

```JavaScript
if (Office.context.mailbox.item.itemType == Office.MailboxEnums.ItemType.Message)
  // do something
else
  // do something else
```

####  <a name="location-stringlocationjavascriptapioutlook12officelocation"></a><span data-ttu-id="b89d7-347">location :String|[Location](/javascript/api/outlook_1_2/office.location)</span><span class="sxs-lookup"><span data-stu-id="b89d7-347">location :String|[Location](/javascript/api/outlook_1_2/office.location)</span></span>

<span data-ttu-id="b89d7-348">Ruft den Ort eines Termins ab bzw. legt ihn fest.</span><span class="sxs-lookup"><span data-stu-id="b89d7-348">Gets or sets the location of an appointment.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="b89d7-349">Lesemodus</span><span class="sxs-lookup"><span data-stu-id="b89d7-349">Read mode</span></span>

<span data-ttu-id="b89d7-350">Die `location`-Eigenschaft gibt eine Zeichenfolge zurück, die den Ort des Termins enthält.</span><span class="sxs-lookup"><span data-stu-id="b89d7-350">The `location` property returns a string that contains the location of the appointment.</span></span>

##### <a name="compose-mode"></a><span data-ttu-id="b89d7-351">Verfassenmodus</span><span class="sxs-lookup"><span data-stu-id="b89d7-351">Compose mode</span></span>

<span data-ttu-id="b89d7-352">Die `location`-Eigenschaft gibt ein `Location`-Objekt zurück, das Methoden zum Abrufen und Festlegen des Orts für einen Termin bereitstellt.</span><span class="sxs-lookup"><span data-stu-id="b89d7-352">The `location` property returns a `Location` object that provides methods that are used to get and set the location of the appointment.</span></span>

##### <a name="type"></a><span data-ttu-id="b89d7-353">Typ:</span><span class="sxs-lookup"><span data-stu-id="b89d7-353">Type:</span></span>

*   <span data-ttu-id="b89d7-354">String | [Location](/javascript/api/outlook_1_2/office.location)</span><span class="sxs-lookup"><span data-stu-id="b89d7-354">String | [Location](/javascript/api/outlook_1_2/office.location)</span></span>

##### <a name="requirements"></a><span data-ttu-id="b89d7-355">Anforderungen</span><span class="sxs-lookup"><span data-stu-id="b89d7-355">Requirements</span></span>

|<span data-ttu-id="b89d7-356">Anforderung</span><span class="sxs-lookup"><span data-stu-id="b89d7-356">Requirement</span></span>| <span data-ttu-id="b89d7-357">Wert</span><span class="sxs-lookup"><span data-stu-id="b89d7-357">Value</span></span>|
|---|---|
|[<span data-ttu-id="b89d7-358">Mindestversion des Postfachanforderungssatzes</span><span class="sxs-lookup"><span data-stu-id="b89d7-358">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="b89d7-359">1.0</span><span class="sxs-lookup"><span data-stu-id="b89d7-359">1.0</span></span>|
|[<span data-ttu-id="b89d7-360">Mindestberechtigungsstufe</span><span class="sxs-lookup"><span data-stu-id="b89d7-360">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="b89d7-361">ReadItem</span><span class="sxs-lookup"><span data-stu-id="b89d7-361">ReadItem</span></span>|
|[<span data-ttu-id="b89d7-362">Zutreffender Outlook-Modus</span><span class="sxs-lookup"><span data-stu-id="b89d7-362">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="b89d7-363">Verfassen oder Lesen</span><span class="sxs-lookup"><span data-stu-id="b89d7-363">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="b89d7-364">Beispiel</span><span class="sxs-lookup"><span data-stu-id="b89d7-364">Example</span></span>

```JavaScript
var userContext = { value : 1 };
Office.context.mailbox.item.location.getAsync( { context: userContext}, callback);

function callback(asyncResult) {
  var context = asyncResult.context;
  var location = asyncResult.value;
}
```

#### <a name="normalizedsubject-string"></a><span data-ttu-id="b89d7-365">normalizedSubject :String</span><span class="sxs-lookup"><span data-stu-id="b89d7-365">normalizedSubject :String</span></span>

<span data-ttu-id="b89d7-p121">Ruft den Betreff für ein Element ab, wobei alle Präfixe entfernt werden (einschließlich `RE:` und `FWD:`). Nur Lesemodus.</span><span class="sxs-lookup"><span data-stu-id="b89d7-p121">Gets the subject of an item, with all prefixes removed (including `RE:` and `FWD:`). Read mode only.</span></span>

<span data-ttu-id="b89d7-p122">Die normalizedSubject-Eigenschaft ruft den Betreff des Elements mit allen Standardpräfixen (z. B. `RE:` und `FW:`) ab, die von E-Mail-Programmen hinzugefügt werden. Wenn der Betreff des Elements mit vollständigen Präfixen abgerufen werden soll, verwenden Sie die [`subject`](#subject-stringsubjectjavascriptapioutlook12officesubject)-Eigenschaft.</span><span class="sxs-lookup"><span data-stu-id="b89d7-p122">The normalizedSubject property gets the subject of the item, with any standard prefixes (such as `RE:` and `FW:`) that are added by email programs. To get the subject of the item with the prefixes intact, use the [`subject`](#subject-stringsubjectjavascriptapioutlook12officesubject) property.</span></span>

##### <a name="type"></a><span data-ttu-id="b89d7-370">Typ:</span><span class="sxs-lookup"><span data-stu-id="b89d7-370">Type:</span></span>

*   <span data-ttu-id="b89d7-371">String</span><span class="sxs-lookup"><span data-stu-id="b89d7-371">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="b89d7-372">Anforderungen</span><span class="sxs-lookup"><span data-stu-id="b89d7-372">Requirements</span></span>

|<span data-ttu-id="b89d7-373">Anforderung</span><span class="sxs-lookup"><span data-stu-id="b89d7-373">Requirement</span></span>| <span data-ttu-id="b89d7-374">Wert</span><span class="sxs-lookup"><span data-stu-id="b89d7-374">Value</span></span>|
|---|---|
|[<span data-ttu-id="b89d7-375">Mindestversion des Postfachanforderungssatzes</span><span class="sxs-lookup"><span data-stu-id="b89d7-375">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="b89d7-376">1.0</span><span class="sxs-lookup"><span data-stu-id="b89d7-376">1.0</span></span>|
|[<span data-ttu-id="b89d7-377">Mindestberechtigungsstufe</span><span class="sxs-lookup"><span data-stu-id="b89d7-377">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="b89d7-378">ReadItem</span><span class="sxs-lookup"><span data-stu-id="b89d7-378">ReadItem</span></span>|
|[<span data-ttu-id="b89d7-379">Zutreffender Outlook-Modus</span><span class="sxs-lookup"><span data-stu-id="b89d7-379">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="b89d7-380">Lesen</span><span class="sxs-lookup"><span data-stu-id="b89d7-380">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="b89d7-381">Beispiel</span><span class="sxs-lookup"><span data-stu-id="b89d7-381">Example</span></span>

```JavaScript
var normalizedSubject = Office.context.mailbox.item.normalizedSubject;
```

####  <a name="optionalattendees-arrayemailaddressdetailsjavascriptapioutlook12officeemailaddressdetailsrecipientsjavascriptapioutlook12officerecipients"></a><span data-ttu-id="b89d7-382">optionalAttendees :Array.<[EmailAddressDetails](/javascript/api/outlook_1_2/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook_1_2/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="b89d7-382">optionalAttendees :Array.<[EmailAddressDetails](/javascript/api/outlook_1_2/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook_1_2/office.recipients)</span></span>

<span data-ttu-id="b89d7-383">Ermöglicht den Zugriff auf die optionalen Teilnehmer eines Ereignisses.</span><span class="sxs-lookup"><span data-stu-id="b89d7-383">Provides access to the optional attendees of an event.</span></span> <span data-ttu-id="b89d7-384">Der Typ des Objekts und die Zugriffsebene, hängt von den Modus des aktuellen Elements ab.</span><span class="sxs-lookup"><span data-stu-id="b89d7-384">The type of object and level of access depends on the mode of the current item.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="b89d7-385">Lesemodus</span><span class="sxs-lookup"><span data-stu-id="b89d7-385">Read mode</span></span>

<span data-ttu-id="b89d7-386">Die `optionalAttendees`-Eigenschaft gibt ein Array mit einem `EmailAddressDetails`-Objekt für jeden optionalen Teilnehmer an der Besprechung zurück.</span><span class="sxs-lookup"><span data-stu-id="b89d7-386">The `optionalAttendees` property returns an array that contains an `EmailAddressDetails` object for each optional attendee to the meeting.</span></span>

##### <a name="compose-mode"></a><span data-ttu-id="b89d7-387">Verfassenmodus</span><span class="sxs-lookup"><span data-stu-id="b89d7-387">Compose mode</span></span>

<span data-ttu-id="b89d7-388">Die `optionalAttendees` -Eigenschaft gibt eine `Recipients` -Objekt, das Methoden zum Abrufen oder Aktualisieren der optionalen Teilnehmer für eine Besprechung bereitstellt.</span><span class="sxs-lookup"><span data-stu-id="b89d7-388">The `optionalAttendees` property returns a `Recipients` object that provides methods to get or update the optional attendees for a meeting.</span></span>

##### <a name="type"></a><span data-ttu-id="b89d7-389">Typ:</span><span class="sxs-lookup"><span data-stu-id="b89d7-389">Type:</span></span>

*   <span data-ttu-id="b89d7-390">Array.<[EmailAddressDetails](/javascript/api/outlook_1_2/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook_1_2/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="b89d7-390">Array.<[EmailAddressDetails](/javascript/api/outlook_1_2/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook_1_2/office.recipients)</span></span>

##### <a name="requirements"></a><span data-ttu-id="b89d7-391">Anforderungen</span><span class="sxs-lookup"><span data-stu-id="b89d7-391">Requirements</span></span>

|<span data-ttu-id="b89d7-392">Anforderung</span><span class="sxs-lookup"><span data-stu-id="b89d7-392">Requirement</span></span>| <span data-ttu-id="b89d7-393">Wert</span><span class="sxs-lookup"><span data-stu-id="b89d7-393">Value</span></span>|
|---|---|
|[<span data-ttu-id="b89d7-394">Mindestversion des Postfachanforderungssatzes</span><span class="sxs-lookup"><span data-stu-id="b89d7-394">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="b89d7-395">1.0</span><span class="sxs-lookup"><span data-stu-id="b89d7-395">1.0</span></span>|
|[<span data-ttu-id="b89d7-396">Mindestberechtigungsstufe</span><span class="sxs-lookup"><span data-stu-id="b89d7-396">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="b89d7-397">ReadItem</span><span class="sxs-lookup"><span data-stu-id="b89d7-397">ReadItem</span></span>|
|[<span data-ttu-id="b89d7-398">Zutreffender Outlook-Modus</span><span class="sxs-lookup"><span data-stu-id="b89d7-398">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="b89d7-399">Verfassen oder Lesen</span><span class="sxs-lookup"><span data-stu-id="b89d7-399">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="b89d7-400">Beispiel</span><span class="sxs-lookup"><span data-stu-id="b89d7-400">Example</span></span>

```JavaScript
Office.context.mailbox.item.optionalAttendees.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.optionalAttendees.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.optionalAttendees.getAsync(callback);

function callback(asyncResult) {
  var arrayOfOptionalAttendeesRecipients = asyncResult.value;
}
```

#### <a name="organizer-emailaddressdetailsjavascriptapioutlook12officeemailaddressdetails"></a><span data-ttu-id="b89d7-401">organizer :[EmailAddressDetails](/javascript/api/outlook_1_2/office.emailaddressdetails)</span><span class="sxs-lookup"><span data-stu-id="b89d7-401">organizer :[EmailAddressDetails](/javascript/api/outlook_1_2/office.emailaddressdetails)</span></span>

<span data-ttu-id="b89d7-p124">Ruft für eine angegebene Besprechung die E-Mail-Adresse des Organisators der Besprechung ab. Nur Lesemodus.</span><span class="sxs-lookup"><span data-stu-id="b89d7-p124">Gets the email address of the meeting organizer for a specified meeting. Read mode only.</span></span>

##### <a name="type"></a><span data-ttu-id="b89d7-404">Typ:</span><span class="sxs-lookup"><span data-stu-id="b89d7-404">Type:</span></span>

*   [<span data-ttu-id="b89d7-405">EmailAddressDetails</span><span class="sxs-lookup"><span data-stu-id="b89d7-405">EmailAddressDetails</span></span>](/javascript/api/outlook_1_2/office.emailaddressdetails)

##### <a name="requirements"></a><span data-ttu-id="b89d7-406">Anforderungen</span><span class="sxs-lookup"><span data-stu-id="b89d7-406">Requirements</span></span>

|<span data-ttu-id="b89d7-407">Anforderung</span><span class="sxs-lookup"><span data-stu-id="b89d7-407">Requirement</span></span>| <span data-ttu-id="b89d7-408">Wert</span><span class="sxs-lookup"><span data-stu-id="b89d7-408">Value</span></span>|
|---|---|
|[<span data-ttu-id="b89d7-409">Mindestversion des Postfachanforderungssatzes</span><span class="sxs-lookup"><span data-stu-id="b89d7-409">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="b89d7-410">1.0</span><span class="sxs-lookup"><span data-stu-id="b89d7-410">1.0</span></span>|
|[<span data-ttu-id="b89d7-411">Mindestberechtigungsstufe</span><span class="sxs-lookup"><span data-stu-id="b89d7-411">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="b89d7-412">ReadItem</span><span class="sxs-lookup"><span data-stu-id="b89d7-412">ReadItem</span></span>|
|[<span data-ttu-id="b89d7-413">Zutreffender Outlook-Modus</span><span class="sxs-lookup"><span data-stu-id="b89d7-413">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="b89d7-414">Lesen</span><span class="sxs-lookup"><span data-stu-id="b89d7-414">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="b89d7-415">Beispiel</span><span class="sxs-lookup"><span data-stu-id="b89d7-415">Example</span></span>

```JavaScript
var organizerName = Office.context.mailbox.item.organizer.displayName;
var organizerAddress = Office.context.mailbox.item.organizer.emailAddress;
```

####  <a name="requiredattendees-arrayemailaddressdetailsjavascriptapioutlook12officeemailaddressdetailsrecipientsjavascriptapioutlook12officerecipients"></a><span data-ttu-id="b89d7-416">requiredAttendees :Array.<[EmailAddressDetails](/javascript/api/outlook_1_2/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook_1_2/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="b89d7-416">requiredAttendees :Array.<[EmailAddressDetails](/javascript/api/outlook_1_2/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook_1_2/office.recipients)</span></span>

<span data-ttu-id="b89d7-417">Ermöglicht den Zugriff auf die erforderlichen Teilnehmer eines Ereignisses.</span><span class="sxs-lookup"><span data-stu-id="b89d7-417">Provides access to the required attendees of an event.</span></span> <span data-ttu-id="b89d7-418">Der Typ des Objekts und die Zugriffsebene, hängt von den Modus des aktuellen Elements ab.</span><span class="sxs-lookup"><span data-stu-id="b89d7-418">The type of object and level of access depends on the mode of the current item.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="b89d7-419">Lesemodus</span><span class="sxs-lookup"><span data-stu-id="b89d7-419">Read mode</span></span>

<span data-ttu-id="b89d7-420">Die `requiredAttendees`-Eigenschaft gibt ein Array mit einem `EmailAddressDetails`-Objekt für jeden erforderlichen Teilnehmer an der Besprechung zurück.</span><span class="sxs-lookup"><span data-stu-id="b89d7-420">The `requiredAttendees` property returns an array that contains an `EmailAddressDetails` object for each required attendee to the meeting.</span></span>

##### <a name="compose-mode"></a><span data-ttu-id="b89d7-421">Verfassenmodus</span><span class="sxs-lookup"><span data-stu-id="b89d7-421">Compose mode</span></span>

<span data-ttu-id="b89d7-422">Die `requiredAttendees` -Eigenschaft gibt eine `Recipients` -Objekt, das Methoden zum Abrufen oder Aktualisieren der erforderlichen Teilnehmer für eine Besprechung bereitstellt.</span><span class="sxs-lookup"><span data-stu-id="b89d7-422">The `requiredAttendees` property returns a `Recipients` object that provides methods to get or update the required attendees for a meeting.</span></span>

##### <a name="type"></a><span data-ttu-id="b89d7-423">Typ:</span><span class="sxs-lookup"><span data-stu-id="b89d7-423">Type:</span></span>

*   <span data-ttu-id="b89d7-424">Array.<[EmailAddressDetails](/javascript/api/outlook_1_2/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook_1_2/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="b89d7-424">Array.<[EmailAddressDetails](/javascript/api/outlook_1_2/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook_1_2/office.recipients)</span></span>

##### <a name="requirements"></a><span data-ttu-id="b89d7-425">Anforderungen</span><span class="sxs-lookup"><span data-stu-id="b89d7-425">Requirements</span></span>

|<span data-ttu-id="b89d7-426">Anforderung</span><span class="sxs-lookup"><span data-stu-id="b89d7-426">Requirement</span></span>| <span data-ttu-id="b89d7-427">Wert</span><span class="sxs-lookup"><span data-stu-id="b89d7-427">Value</span></span>|
|---|---|
|[<span data-ttu-id="b89d7-428">Mindestversion des Postfachanforderungssatzes</span><span class="sxs-lookup"><span data-stu-id="b89d7-428">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="b89d7-429">1.0</span><span class="sxs-lookup"><span data-stu-id="b89d7-429">1.0</span></span>|
|[<span data-ttu-id="b89d7-430">Mindestberechtigungsstufe</span><span class="sxs-lookup"><span data-stu-id="b89d7-430">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="b89d7-431">ReadItem</span><span class="sxs-lookup"><span data-stu-id="b89d7-431">ReadItem</span></span>|
|[<span data-ttu-id="b89d7-432">Zutreffender Outlook-Modus</span><span class="sxs-lookup"><span data-stu-id="b89d7-432">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="b89d7-433">Verfassen oder Lesen</span><span class="sxs-lookup"><span data-stu-id="b89d7-433">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="b89d7-434">Beispiel</span><span class="sxs-lookup"><span data-stu-id="b89d7-434">Example</span></span>

```JavaScript
Office.context.mailbox.item.requiredAttendees.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.requiredAttendees.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.requiredAttendees.getAsync(callback);

function callback(asyncResult) {
  var arrayOfRequiredAttendeesRecipients = asyncResult.value;
}
```

#### <a name="sender-emailaddressdetailsjavascriptapioutlook12officeemailaddressdetails"></a><span data-ttu-id="b89d7-435">sender :[EmailAddressDetails](/javascript/api/outlook_1_2/office.emailaddressdetails)</span><span class="sxs-lookup"><span data-stu-id="b89d7-435">sender :[EmailAddressDetails](/javascript/api/outlook_1_2/office.emailaddressdetails)</span></span>

<span data-ttu-id="b89d7-p126">Ruft die E-Mail-Adresse des Absenders einer E-Mail-Nachricht ab. Nur Lesemodus.</span><span class="sxs-lookup"><span data-stu-id="b89d7-p126">Gets the email address of the sender of an email message. Read mode only.</span></span>

<span data-ttu-id="b89d7-p127">Die [`from`](#from-emailaddressdetailsjavascriptapioutlook12officeemailaddressdetails)- und `sender`-Eigenschaften stellen dieselbe Person dar, außer die Nachricht wurde von einem Delegaten gesendet. In diesem Fall stellt die `from`-Eigenschaft die Stellvertretung dar, und die sender-Eigenschaft stellt den Delegaten dar.</span><span class="sxs-lookup"><span data-stu-id="b89d7-p127">The [`from`](#from-emailaddressdetailsjavascriptapioutlook12officeemailaddressdetails) and `sender` properties represent the same person unless the message is sent by a delegate. In that case, the `from` property represents the delegator, and the sender property represents the delegate.</span></span>

> [!NOTE]
> <span data-ttu-id="b89d7-440">Die `recipientType` -Eigenschaft des der `EmailAddressDetails` -Objekts der `sender` -Eigenschaft ist `undefined`.</span><span class="sxs-lookup"><span data-stu-id="b89d7-440">The `recipientType` property of the `EmailAddressDetails` object in the `sender` property is `undefined`.</span></span>

##### <a name="type"></a><span data-ttu-id="b89d7-441">Typ:</span><span class="sxs-lookup"><span data-stu-id="b89d7-441">Type:</span></span>

*   [<span data-ttu-id="b89d7-442">EmailAddressDetails</span><span class="sxs-lookup"><span data-stu-id="b89d7-442">EmailAddressDetails</span></span>](/javascript/api/outlook_1_2/office.emailaddressdetails)

##### <a name="requirements"></a><span data-ttu-id="b89d7-443">Anforderungen</span><span class="sxs-lookup"><span data-stu-id="b89d7-443">Requirements</span></span>

|<span data-ttu-id="b89d7-444">Anforderung</span><span class="sxs-lookup"><span data-stu-id="b89d7-444">Requirement</span></span>| <span data-ttu-id="b89d7-445">Wert</span><span class="sxs-lookup"><span data-stu-id="b89d7-445">Value</span></span>|
|---|---|
|[<span data-ttu-id="b89d7-446">Mindestversion des Postfachanforderungssatzes</span><span class="sxs-lookup"><span data-stu-id="b89d7-446">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="b89d7-447">1.0</span><span class="sxs-lookup"><span data-stu-id="b89d7-447">1.0</span></span>|
|[<span data-ttu-id="b89d7-448">Mindestberechtigungsstufe</span><span class="sxs-lookup"><span data-stu-id="b89d7-448">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="b89d7-449">ReadItem</span><span class="sxs-lookup"><span data-stu-id="b89d7-449">ReadItem</span></span>|
|[<span data-ttu-id="b89d7-450">Zutreffender Outlook-Modus</span><span class="sxs-lookup"><span data-stu-id="b89d7-450">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="b89d7-451">Lesen</span><span class="sxs-lookup"><span data-stu-id="b89d7-451">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="b89d7-452">Beispiel</span><span class="sxs-lookup"><span data-stu-id="b89d7-452">Example</span></span>

```JavaScript
var senderName = Office.context.mailbox.item.sender.displayName;
var senderAddress = Office.context.mailbox.item.sender.emailAddress;
```

####  <a name="start-datetimejavascriptapioutlook12officetime"></a><span data-ttu-id="b89d7-453">start :Date|[Time](/javascript/api/outlook_1_2/office.time)</span><span class="sxs-lookup"><span data-stu-id="b89d7-453">start :Date|[Time](/javascript/api/outlook_1_2/office.time)</span></span>

<span data-ttu-id="b89d7-454">Ruft Datum und Zeit für den Beginn des Termins ab oder legt Datum und Uhrzeit fest.</span><span class="sxs-lookup"><span data-stu-id="b89d7-454">Gets or sets the date and time that the appointment is to begin.</span></span>

<span data-ttu-id="b89d7-p128">Die `start`-Eigenschaft wird als Datums- und Uhrzeitwert in UTC ausgedrückt. Sie können die [`convertToLocalClientTime`](office.context.mailbox.md#converttolocalclienttimetimevalue--localclienttimejavascriptapioutlook12officelocalclienttime)-Methode verwenden, um den Wert in den lokalen Uhrzeit- und Datumswert des Clients umzuwandeln.</span><span class="sxs-lookup"><span data-stu-id="b89d7-p128">The `start` property is expressed as a Coordinated Universal Time (UTC) date and time value. You can use the [`convertToLocalClientTime`](office.context.mailbox.md#converttolocalclienttimetimevalue--localclienttimejavascriptapioutlook12officelocalclienttime) method to convert the value to the client’s local date and time.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="b89d7-457">Lesemodus</span><span class="sxs-lookup"><span data-stu-id="b89d7-457">Read mode</span></span>

<span data-ttu-id="b89d7-458">Die `start`-Eigenschaft gibt ein `Date`-Objekt zurück.</span><span class="sxs-lookup"><span data-stu-id="b89d7-458">The `start` property returns a `Date` object.</span></span>

##### <a name="compose-mode"></a><span data-ttu-id="b89d7-459">Verfassenmodus</span><span class="sxs-lookup"><span data-stu-id="b89d7-459">Compose mode</span></span>

<span data-ttu-id="b89d7-460">Die `start`-Eigenschaft gibt ein `Time`-Objekt zurück.</span><span class="sxs-lookup"><span data-stu-id="b89d7-460">The `start` property returns a `Time` object.</span></span>

<span data-ttu-id="b89d7-461">Wenn Sie die [`Time.setAsync`](/javascript/api/outlook_1_2/office.time#setasync-datetime--options--callback-)-Methode verwenden, um die Startzeit im Verfassenmodus festzulegen, sollten Sie die [`convertToUtcClientTime`](office.context.mailbox.md#converttoutcclienttimeinput--date)-Methode verwenden, um die Ortszeit auf dem Client für den Server in UTC umzuwandeln.</span><span class="sxs-lookup"><span data-stu-id="b89d7-461">When you use the [`Time.setAsync`](/javascript/api/outlook_1_2/office.time#setasync-datetime--options--callback-) method to set the start time, you should use the [`convertToUtcClientTime`](office.context.mailbox.md#converttoutcclienttimeinput--date) method to convert the local time on the client to UTC for the server.</span></span>

##### <a name="type"></a><span data-ttu-id="b89d7-462">Typ:</span><span class="sxs-lookup"><span data-stu-id="b89d7-462">Type:</span></span>

*   <span data-ttu-id="b89d7-463">Date | [Time](/javascript/api/outlook_1_2/office.time)</span><span class="sxs-lookup"><span data-stu-id="b89d7-463">Date | [Time](/javascript/api/outlook_1_2/office.time)</span></span>

##### <a name="requirements"></a><span data-ttu-id="b89d7-464">Anforderungen</span><span class="sxs-lookup"><span data-stu-id="b89d7-464">Requirements</span></span>

|<span data-ttu-id="b89d7-465">Anforderung</span><span class="sxs-lookup"><span data-stu-id="b89d7-465">Requirement</span></span>| <span data-ttu-id="b89d7-466">Wert</span><span class="sxs-lookup"><span data-stu-id="b89d7-466">Value</span></span>|
|---|---|
|[<span data-ttu-id="b89d7-467">Mindestversion des Postfachanforderungssatzes</span><span class="sxs-lookup"><span data-stu-id="b89d7-467">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="b89d7-468">1.0</span><span class="sxs-lookup"><span data-stu-id="b89d7-468">1.0</span></span>|
|[<span data-ttu-id="b89d7-469">Mindestberechtigungsstufe</span><span class="sxs-lookup"><span data-stu-id="b89d7-469">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="b89d7-470">ReadItem</span><span class="sxs-lookup"><span data-stu-id="b89d7-470">ReadItem</span></span>|
|[<span data-ttu-id="b89d7-471">Zutreffender Outlook-Modus</span><span class="sxs-lookup"><span data-stu-id="b89d7-471">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="b89d7-472">Verfassen oder Lesen</span><span class="sxs-lookup"><span data-stu-id="b89d7-472">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="b89d7-473">Beispiel</span><span class="sxs-lookup"><span data-stu-id="b89d7-473">Example</span></span>

<span data-ttu-id="b89d7-474">Im folgenden Beispiel wird die Startzeit eines Termins im Verfassenmodus mithilfe der [`setAsync`](/javascript/api/outlook_1_2/office.time#setasync-datetime--options--callback-)-Methode des `Time`-Objekts festgelegt.</span><span class="sxs-lookup"><span data-stu-id="b89d7-474">The following example sets the start time of an appointment in compose mode by using the [`setAsync`](/javascript/api/outlook_1_2/office.time#setasync-datetime--options--callback-) method of the `Time` object.</span></span>

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

####  <a name="subject-stringsubjectjavascriptapioutlook12officesubject"></a><span data-ttu-id="b89d7-475">subject :String|[Subject](/javascript/api/outlook_1_2/office.subject)</span><span class="sxs-lookup"><span data-stu-id="b89d7-475">subject :String|[Subject](/javascript/api/outlook_1_2/office.subject)</span></span>

<span data-ttu-id="b89d7-476">Ruft die Beschreibung ab, die im Betrefffeld eines Elements angezeigt wird, oder legt sie fest.</span><span class="sxs-lookup"><span data-stu-id="b89d7-476">Gets or sets the description that appears in the subject field of an item.</span></span>

<span data-ttu-id="b89d7-477">Die `subject`-Eigenschaft ruft den gesamten Betreff des Elements ab oder legt ihn fest – so, wie er vom E-Mail-Server gesendet wird.</span><span class="sxs-lookup"><span data-stu-id="b89d7-477">The `subject` property gets or sets the entire subject of the item, as sent by the email server.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="b89d7-478">Lesemodus</span><span class="sxs-lookup"><span data-stu-id="b89d7-478">Read mode</span></span>

<span data-ttu-id="b89d7-p129">Die `subject`-Eigenschaft gibt eine Zeichenfolge zurück. Verwenden Sie die [`normalizedSubject`](#normalizedsubject-string)-Eigenschaft, um den Betreff ohne führende Präfixe wie `RE:` und `FW:` abzurufen.</span><span class="sxs-lookup"><span data-stu-id="b89d7-p129">The `subject` property returns a string. Use the [`normalizedSubject`](#normalizedsubject-string) property to get the subject minus any leading prefixes such as `RE:` and `FW:`.</span></span>

```
var subject = Office.context.mailbox.item.subject;
```

##### <a name="compose-mode"></a><span data-ttu-id="b89d7-481">Verfassenmodus</span><span class="sxs-lookup"><span data-stu-id="b89d7-481">Compose mode</span></span>

<span data-ttu-id="b89d7-482">Die `subject`-Eigenschaft gibt ein `Subject`-Objekt zurück, das Methoden zum Abrufen und Festlegen des Betreffs bereitstellt.</span><span class="sxs-lookup"><span data-stu-id="b89d7-482">The `subject` property returns a `Subject` object that provides methods to get and set the subject.</span></span>

```JavaScript
Office.context.mailbox.item.subject.getAsync(callback);

function callback(asyncResult) {
  var subject = asyncResult.value;
}
```

##### <a name="type"></a><span data-ttu-id="b89d7-483">Typ:</span><span class="sxs-lookup"><span data-stu-id="b89d7-483">Type:</span></span>

*   <span data-ttu-id="b89d7-484">String | [Subject](/javascript/api/outlook_1_2/office.subject)</span><span class="sxs-lookup"><span data-stu-id="b89d7-484">String | [Subject](/javascript/api/outlook_1_2/office.subject)</span></span>

##### <a name="requirements"></a><span data-ttu-id="b89d7-485">Anforderungen</span><span class="sxs-lookup"><span data-stu-id="b89d7-485">Requirements</span></span>

|<span data-ttu-id="b89d7-486">Anforderung</span><span class="sxs-lookup"><span data-stu-id="b89d7-486">Requirement</span></span>| <span data-ttu-id="b89d7-487">Wert</span><span class="sxs-lookup"><span data-stu-id="b89d7-487">Value</span></span>|
|---|---|
|[<span data-ttu-id="b89d7-488">Mindestversion des Postfachanforderungssatzes</span><span class="sxs-lookup"><span data-stu-id="b89d7-488">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="b89d7-489">1.0</span><span class="sxs-lookup"><span data-stu-id="b89d7-489">1.0</span></span>|
|[<span data-ttu-id="b89d7-490">Mindestberechtigungsstufe</span><span class="sxs-lookup"><span data-stu-id="b89d7-490">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="b89d7-491">ReadItem</span><span class="sxs-lookup"><span data-stu-id="b89d7-491">ReadItem</span></span>|
|[<span data-ttu-id="b89d7-492">Zutreffender Outlook-Modus</span><span class="sxs-lookup"><span data-stu-id="b89d7-492">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="b89d7-493">Verfassen oder Lesen</span><span class="sxs-lookup"><span data-stu-id="b89d7-493">Compose or read</span></span>|

####  <a name="to-arrayemailaddressdetailsjavascriptapioutlook12officeemailaddressdetailsrecipientsjavascriptapioutlook12officerecipients"></a><span data-ttu-id="b89d7-494">an: Array. <[EmailAddressDetails](/javascript/api/outlook_1_2/office.emailaddressdetails)>|[Empfänger](/javascript/api/outlook_1_2/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="b89d7-494">to :Array.<[EmailAddressDetails](/javascript/api/outlook_1_2/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook_1_2/office.recipients)</span></span>

<span data-ttu-id="b89d7-495">Ermöglicht den Zugriff auf die Empfänger in der Zeile **an** einer Nachricht.</span><span class="sxs-lookup"><span data-stu-id="b89d7-495">Provides access to the recipients on the **To** line of a message.</span></span> <span data-ttu-id="b89d7-496">Der Typ des Objekts und die Zugriffsebene, hängt von den Modus des aktuellen Elements ab.</span><span class="sxs-lookup"><span data-stu-id="b89d7-496">The type of object and level of access depends on the mode of the current item.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="b89d7-497">Lesemodus</span><span class="sxs-lookup"><span data-stu-id="b89d7-497">Read mode</span></span>

<span data-ttu-id="b89d7-p131">Die `to`-Eigenschaft gibt ein Array mit einem `EmailAddressDetails`-Objekt für jeden Empfänger in der**An**-Zeile der Nachricht zurück. Die Auflistung ist auf höchstens 100 Elemente beschränkt.</span><span class="sxs-lookup"><span data-stu-id="b89d7-p131">The `to` property returns an array that contains an `EmailAddressDetails` object for each recipient listed on the **To** line of the message. The collection is limited to a maximum of 100 members.</span></span>

##### <a name="compose-mode"></a><span data-ttu-id="b89d7-500">Verfassenmodus</span><span class="sxs-lookup"><span data-stu-id="b89d7-500">Compose mode</span></span>

<span data-ttu-id="b89d7-501">Die `to` -Eigenschaft gibt eine `Recipients` -Objekt, das Methoden zum Abrufen oder aktualisieren die Empfänger in der Zeile **an** der Nachricht bereitstellt.</span><span class="sxs-lookup"><span data-stu-id="b89d7-501">The `to` property returns a `Recipients` object that provides methods to get or update the recipients on the **To** line of the message.</span></span>

##### <a name="type"></a><span data-ttu-id="b89d7-502">Typ:</span><span class="sxs-lookup"><span data-stu-id="b89d7-502">Type:</span></span>

*   <span data-ttu-id="b89d7-503">Array.<[EmailAddressDetails](/javascript/api/outlook_1_2/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook_1_2/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="b89d7-503">Array.<[EmailAddressDetails](/javascript/api/outlook_1_2/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook_1_2/office.recipients)</span></span>

##### <a name="requirements"></a><span data-ttu-id="b89d7-504">Anforderungen</span><span class="sxs-lookup"><span data-stu-id="b89d7-504">Requirements</span></span>

|<span data-ttu-id="b89d7-505">Anforderung</span><span class="sxs-lookup"><span data-stu-id="b89d7-505">Requirement</span></span>| <span data-ttu-id="b89d7-506">Wert</span><span class="sxs-lookup"><span data-stu-id="b89d7-506">Value</span></span>|
|---|---|
|[<span data-ttu-id="b89d7-507">Mindestversion des Postfachanforderungssatzes</span><span class="sxs-lookup"><span data-stu-id="b89d7-507">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="b89d7-508">1.0</span><span class="sxs-lookup"><span data-stu-id="b89d7-508">1.0</span></span>|
|[<span data-ttu-id="b89d7-509">Mindestberechtigungsstufe</span><span class="sxs-lookup"><span data-stu-id="b89d7-509">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="b89d7-510">ReadItem</span><span class="sxs-lookup"><span data-stu-id="b89d7-510">ReadItem</span></span>|
|[<span data-ttu-id="b89d7-511">Zutreffender Outlook-Modus</span><span class="sxs-lookup"><span data-stu-id="b89d7-511">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="b89d7-512">Verfassen oder Lesen</span><span class="sxs-lookup"><span data-stu-id="b89d7-512">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="b89d7-513">Beispiel</span><span class="sxs-lookup"><span data-stu-id="b89d7-513">Example</span></span>

```JavaScript
Office.context.mailbox.item.to.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.to.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.to.getAsync(callback);

function callback(asyncResult) {
  var arrayOfToRecipients = asyncResult.value;
}
```

### <a name="methods"></a><span data-ttu-id="b89d7-514">Methoden</span><span class="sxs-lookup"><span data-stu-id="b89d7-514">Methods</span></span>

####  <a name="addfileattachmentasyncuri-attachmentname-options-callback"></a><span data-ttu-id="b89d7-515">addFileAttachmentAsync(uri, attachmentName, [options], [callback])</span><span class="sxs-lookup"><span data-stu-id="b89d7-515">addFileAttachmentAsync(uri, attachmentName, [options], [callback])</span></span>

<span data-ttu-id="b89d7-516">Fügt eine Datei zu einer Nachricht oder einem Termin als Anlage hinzu.</span><span class="sxs-lookup"><span data-stu-id="b89d7-516">Adds a file to a message or appointment as an attachment.</span></span>

<span data-ttu-id="b89d7-517">Die `addFileAttachmentAsync`-Methode lädt die Datei am angegebenen URI hoch und fügt sie an das Element im Verfassenformular an.</span><span class="sxs-lookup"><span data-stu-id="b89d7-517">The `addFileAttachmentAsync` method uploads the file at the specified URI and attaches it to the item in the compose form.</span></span>

<span data-ttu-id="b89d7-518">Anschließend können Sie den Bezeichner mit der [`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback)-Methode in der gleichen Sitzung zum Entfernen der Anlage verwenden.</span><span class="sxs-lookup"><span data-stu-id="b89d7-518">You can subsequently use the identifier with the [`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback) method to remove the attachment in the same session.</span></span>

##### <a name="parameters"></a><span data-ttu-id="b89d7-519">Parameter:</span><span class="sxs-lookup"><span data-stu-id="b89d7-519">Parameters:</span></span>

|<span data-ttu-id="b89d7-520">Name</span><span class="sxs-lookup"><span data-stu-id="b89d7-520">Name</span></span>| <span data-ttu-id="b89d7-521">Typ</span><span class="sxs-lookup"><span data-stu-id="b89d7-521">Type</span></span>| <span data-ttu-id="b89d7-522">Attribute</span><span class="sxs-lookup"><span data-stu-id="b89d7-522">Attributes</span></span>| <span data-ttu-id="b89d7-523">Beschreibung</span><span class="sxs-lookup"><span data-stu-id="b89d7-523">Description</span></span>|
|---|---|---|---|
|`uri`| <span data-ttu-id="b89d7-524">Zeichenfolge</span><span class="sxs-lookup"><span data-stu-id="b89d7-524">String</span></span>||<span data-ttu-id="b89d7-p132">Der URI, der den Speicherort der an die Nachricht oder den Termin anzuhängenden Datei angibt. Die maximale Länge ist 2048 Zeichen.</span><span class="sxs-lookup"><span data-stu-id="b89d7-p132">The URI that provides the location of the file to attach to the message or appointment. The maximum length is 2048 characters.</span></span>|
|`attachmentName`| <span data-ttu-id="b89d7-527">Zeichenfolge</span><span class="sxs-lookup"><span data-stu-id="b89d7-527">String</span></span>||<span data-ttu-id="b89d7-p133">Der Name der Anlage, der beim Hochladen der Anlage angezeigt wird. Die maximale Länge ist 255 Zeichen.</span><span class="sxs-lookup"><span data-stu-id="b89d7-p133">The name of the attachment that is shown while the attachment is uploading. The maximum length is 255 characters.</span></span>|
|`options`| <span data-ttu-id="b89d7-530">Object</span><span class="sxs-lookup"><span data-stu-id="b89d7-530">Object</span></span>| <span data-ttu-id="b89d7-531">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="b89d7-531">&lt;optional&gt;</span></span>|<span data-ttu-id="b89d7-532">Ein Objektliteral, das mindestens eine der folgenden Eigenschaften enthält.</span><span class="sxs-lookup"><span data-stu-id="b89d7-532">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`| <span data-ttu-id="b89d7-533">Object</span><span class="sxs-lookup"><span data-stu-id="b89d7-533">Object</span></span>| <span data-ttu-id="b89d7-534">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="b89d7-534">&lt;optional&gt;</span></span>|<span data-ttu-id="b89d7-535">Entwickler können ein Objekt bereitstellen, auf das sie in der Callbackmethode zugreifen möchten.</span><span class="sxs-lookup"><span data-stu-id="b89d7-535">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`| <span data-ttu-id="b89d7-536">Funktion</span><span class="sxs-lookup"><span data-stu-id="b89d7-536">function</span></span>| <span data-ttu-id="b89d7-537">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="b89d7-537">&lt;optional&gt;</span></span>|<span data-ttu-id="b89d7-538">Wenn die Methode abgeschlossen wird, wird die Funktion , die im `callback`-Parameter übergeben wird, mit einem einzelnen Parameter (`asyncResult`) aufgerufen, der ein [`AsyncResult`](/javascript/api/office/office.asyncresult)-Objekt darstellt.</span><span class="sxs-lookup"><span data-stu-id="b89d7-538">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span> <br/><span data-ttu-id="b89d7-539">Bei Erfolg wird der Anlagenbezeichner in der `asyncResult.value`-Eigenschaft bereitgestellt.</span><span class="sxs-lookup"><span data-stu-id="b89d7-539">On success, the attachment identifier will be provided in the `asyncResult.value` property.</span></span><br/><span data-ttu-id="b89d7-540">Wenn beim Hochladen der Anlage ein Fehler auftritt, enthält das `asyncResult`-Objekt ein `Error`-Objekt mit einer Beschreibung des Fehlers.</span><span class="sxs-lookup"><span data-stu-id="b89d7-540">If uploading the attachment fails, the `asyncResult` object will contain an `Error` object that provides a description of the error.</span></span>|

##### <a name="errors"></a><span data-ttu-id="b89d7-541">Fehler</span><span class="sxs-lookup"><span data-stu-id="b89d7-541">Errors</span></span>

| <span data-ttu-id="b89d7-542">Fehlercode</span><span class="sxs-lookup"><span data-stu-id="b89d7-542">Error code</span></span> | <span data-ttu-id="b89d7-543">Beschreibung</span><span class="sxs-lookup"><span data-stu-id="b89d7-543">Description</span></span> |
|------------|-------------|
| `AttachmentSizeExceeded` | <span data-ttu-id="b89d7-544">Die Anlage ist zu groß.</span><span class="sxs-lookup"><span data-stu-id="b89d7-544">The attachment is larger than allowed.</span></span> |
| `FileTypeNotSupported` | <span data-ttu-id="b89d7-545">Die Anlage enthält eine Erweiterung, die nicht zulässig ist.</span><span class="sxs-lookup"><span data-stu-id="b89d7-545">The attachment has an extension that is not allowed.</span></span> |
| `NumberOfAttachmentsExceeded` | <span data-ttu-id="b89d7-546">Die Nachricht oder der Termin enthält zu viele Anlagen.</span><span class="sxs-lookup"><span data-stu-id="b89d7-546">The message or appointment has too many attachments.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="b89d7-547">Anforderungen</span><span class="sxs-lookup"><span data-stu-id="b89d7-547">Requirements</span></span>

|<span data-ttu-id="b89d7-548">Anforderung</span><span class="sxs-lookup"><span data-stu-id="b89d7-548">Requirement</span></span>| <span data-ttu-id="b89d7-549">Wert</span><span class="sxs-lookup"><span data-stu-id="b89d7-549">Value</span></span>|
|---|---|
|[<span data-ttu-id="b89d7-550">Mindestversion des Postfachanforderungssatzes</span><span class="sxs-lookup"><span data-stu-id="b89d7-550">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="b89d7-551">1.1</span><span class="sxs-lookup"><span data-stu-id="b89d7-551">1.1</span></span>|
|[<span data-ttu-id="b89d7-552">Mindestberechtigungsstufe</span><span class="sxs-lookup"><span data-stu-id="b89d7-552">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="b89d7-553">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="b89d7-553">ReadWriteItem</span></span>|
|[<span data-ttu-id="b89d7-554">Zutreffender Outlook-Modus</span><span class="sxs-lookup"><span data-stu-id="b89d7-554">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="b89d7-555">Verfassen</span><span class="sxs-lookup"><span data-stu-id="b89d7-555">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="b89d7-556">Beispiel</span><span class="sxs-lookup"><span data-stu-id="b89d7-556">Example</span></span>

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

####  <a name="additemattachmentasyncitemid-attachmentname-options-callback"></a><span data-ttu-id="b89d7-557">addItemAttachmentAsync(itemId, attachmentName, [options], [callback])</span><span class="sxs-lookup"><span data-stu-id="b89d7-557">addItemAttachmentAsync(itemId, attachmentName, [options], [callback])</span></span>

<span data-ttu-id="b89d7-558">Fügt der Nachricht oder dem Termin ein Exchange-Objekt, wie z. B. eine Nachricht, als Anhang hinzu.</span><span class="sxs-lookup"><span data-stu-id="b89d7-558">Adds an Exchange item, such as a message, as an attachment to the message or appointment.</span></span>

<span data-ttu-id="b89d7-p134">Die `addItemAttachmentAsync`-Methode fügt das Element mit dem angegebenen Exchange-Bezeichner an das Element im Formular zum Verfassen an. Wenn Sie eine Rückrufmethode angeben, wird die Methode mit einem `asyncResult`-Parameter aufgerufen, der entweder den Anlagenbezeichner oder einen Code enthält, der etwaige Fehler angibt, die beim Anfügen des Objekts aufgetreten sind. Sie können ggf. den `options`-Parameter verwenden, um Statusinformationen an die Rückrufmethode zu übergeben.</span><span class="sxs-lookup"><span data-stu-id="b89d7-p134">The `addItemAttachmentAsync` method attaches the item with the specified Exchange identifier to the item in the compose form. If you specify a callback method, the method is called with one parameter, `asyncResult`, which contains either the attachment identifier or a code that indicates any error that occurred while attaching the item. You can use the `options` parameter to pass state information to the callback method, if needed.</span></span>

<span data-ttu-id="b89d7-562">Anschließend können Sie den Bezeichner mit der [`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback)-Methode in der gleichen Sitzung zum Entfernen der Anlage verwenden.</span><span class="sxs-lookup"><span data-stu-id="b89d7-562">You can subsequently use the identifier with the [`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback) method to remove the attachment in the same session.</span></span>

<span data-ttu-id="b89d7-563">Wenn Ihre Office-Add-Ins in Outlook Web App ausgeführt wird die `addItemAttachmentAsync` -Methode kann Anfügen von Elementen in der Elemente, die Sie bearbeiten; Allerdings Dies wird nicht unterstützt und wird nicht empfohlen.</span><span class="sxs-lookup"><span data-stu-id="b89d7-563">If your Office Add-in is running in Outlook Web App, the `addItemAttachmentAsync` method can attach items to items other than the item that you are editing; however, this is not supported and is not recommended.</span></span>

##### <a name="parameters"></a><span data-ttu-id="b89d7-564">Parameter:</span><span class="sxs-lookup"><span data-stu-id="b89d7-564">Parameters:</span></span>

|<span data-ttu-id="b89d7-565">Name</span><span class="sxs-lookup"><span data-stu-id="b89d7-565">Name</span></span>| <span data-ttu-id="b89d7-566">Typ</span><span class="sxs-lookup"><span data-stu-id="b89d7-566">Type</span></span>| <span data-ttu-id="b89d7-567">Attribute</span><span class="sxs-lookup"><span data-stu-id="b89d7-567">Attributes</span></span>| <span data-ttu-id="b89d7-568">Beschreibung</span><span class="sxs-lookup"><span data-stu-id="b89d7-568">Description</span></span>|
|---|---|---|---|
|`itemId`| <span data-ttu-id="b89d7-569">Zeichenfolge</span><span class="sxs-lookup"><span data-stu-id="b89d7-569">String</span></span>||<span data-ttu-id="b89d7-p135">Der Exchange-Bezeichner des Objekts, das angehängt werden soll. Die maximale Länge beträgt 100 Zeichen.</span><span class="sxs-lookup"><span data-stu-id="b89d7-p135">The Exchange identifier of the item to attach. The maximum length is 100 characters.</span></span>|
|`attachmentName`| <span data-ttu-id="b89d7-572">Zeichenfolge</span><span class="sxs-lookup"><span data-stu-id="b89d7-572">String</span></span>||<span data-ttu-id="b89d7-p136">Der Betreff des Elements, das angehängt werden soll. Die maximale Länge ist 255 Zeichen.</span><span class="sxs-lookup"><span data-stu-id="b89d7-p136">The sujbect of the item to be attached. The maximum length is 255 characters.</span></span>|
|`options`| <span data-ttu-id="b89d7-575">Object</span><span class="sxs-lookup"><span data-stu-id="b89d7-575">Object</span></span>| <span data-ttu-id="b89d7-576">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="b89d7-576">&lt;optional&gt;</span></span>|<span data-ttu-id="b89d7-577">Ein Objektliteral, das mindestens eine der folgenden Eigenschaften enthält.</span><span class="sxs-lookup"><span data-stu-id="b89d7-577">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`| <span data-ttu-id="b89d7-578">Object</span><span class="sxs-lookup"><span data-stu-id="b89d7-578">Object</span></span>| <span data-ttu-id="b89d7-579">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="b89d7-579">&lt;optional&gt;</span></span>|<span data-ttu-id="b89d7-580">Entwickler können ein Objekt bereitstellen, auf das sie in der Callbackmethode zugreifen möchten.</span><span class="sxs-lookup"><span data-stu-id="b89d7-580">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`| <span data-ttu-id="b89d7-581">Funktion</span><span class="sxs-lookup"><span data-stu-id="b89d7-581">function</span></span>| <span data-ttu-id="b89d7-582">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="b89d7-582">&lt;optional&gt;</span></span>|<span data-ttu-id="b89d7-583">Wenn die Methode abgeschlossen wird, wird die Funktion , die im `callback`-Parameter übergeben wird, mit einem einzelnen Parameter (`asyncResult`) aufgerufen, der ein [`AsyncResult`](/javascript/api/office/office.asyncresult)-Objekt darstellt.</span><span class="sxs-lookup"><span data-stu-id="b89d7-583">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span> <br/><span data-ttu-id="b89d7-584">Bei Erfolg wird der Anlagenbezeichner in der `asyncResult.value`-Eigenschaft bereitgestellt.</span><span class="sxs-lookup"><span data-stu-id="b89d7-584">On success, the attachment identifier will be provided in the `asyncResult.value` property.</span></span><br/><span data-ttu-id="b89d7-585">Wenn beim Hinzufügen der Anlage ein Fehler auftritt, enthält das `asyncResult`-Objekt ein `Error`-Objekt mit einer Beschreibung des Fehlers.</span><span class="sxs-lookup"><span data-stu-id="b89d7-585">If adding the attachment fails, the `asyncResult` object will contain an `Error` object that provides a description of the error.</span></span>|

##### <a name="errors"></a><span data-ttu-id="b89d7-586">Fehler</span><span class="sxs-lookup"><span data-stu-id="b89d7-586">Errors</span></span>

| <span data-ttu-id="b89d7-587">Fehlercode</span><span class="sxs-lookup"><span data-stu-id="b89d7-587">Error code</span></span> | <span data-ttu-id="b89d7-588">Beschreibung</span><span class="sxs-lookup"><span data-stu-id="b89d7-588">Description</span></span> |
|------------|-------------|
| `NumberOfAttachmentsExceeded` | <span data-ttu-id="b89d7-589">Die Nachricht oder der Termin enthält zu viele Anlagen.</span><span class="sxs-lookup"><span data-stu-id="b89d7-589">The message or appointment has too many attachments.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="b89d7-590">Anforderungen</span><span class="sxs-lookup"><span data-stu-id="b89d7-590">Requirements</span></span>

|<span data-ttu-id="b89d7-591">Anforderung</span><span class="sxs-lookup"><span data-stu-id="b89d7-591">Requirement</span></span>| <span data-ttu-id="b89d7-592">Wert</span><span class="sxs-lookup"><span data-stu-id="b89d7-592">Value</span></span>|
|---|---|
|[<span data-ttu-id="b89d7-593">Mindestversion des Postfachanforderungssatzes</span><span class="sxs-lookup"><span data-stu-id="b89d7-593">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="b89d7-594">1.1</span><span class="sxs-lookup"><span data-stu-id="b89d7-594">1.1</span></span>|
|[<span data-ttu-id="b89d7-595">Mindestberechtigungsstufe</span><span class="sxs-lookup"><span data-stu-id="b89d7-595">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="b89d7-596">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="b89d7-596">ReadWriteItem</span></span>|
|[<span data-ttu-id="b89d7-597">Zutreffender Outlook-Modus</span><span class="sxs-lookup"><span data-stu-id="b89d7-597">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="b89d7-598">Verfassen</span><span class="sxs-lookup"><span data-stu-id="b89d7-598">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="b89d7-599">Beispiel</span><span class="sxs-lookup"><span data-stu-id="b89d7-599">Example</span></span>

<span data-ttu-id="b89d7-600">Im folgenden Beispiel wird ein vorhandenes Outlook-Element als Anlage mit dem Namen `My Attachment` hinzugefügt.</span><span class="sxs-lookup"><span data-stu-id="b89d7-600">The following example adds an existing Outlook item as an attachment with the name `My Attachment`.</span></span>

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

#### <a name="displayreplyallformformdata"></a><span data-ttu-id="b89d7-601">displayReplyAllForm(formData)</span><span class="sxs-lookup"><span data-stu-id="b89d7-601">displayReplyAllForm(formData)</span></span>

<span data-ttu-id="b89d7-602">Zeigt ein Antwortformular an, das den Absender und alle Empfänger der ausgewählten Nachricht oder des Organisators und alle Teilnehmer des ausgewählten Termins enthält.</span><span class="sxs-lookup"><span data-stu-id="b89d7-602">Displays a reply form that includes the sender and all recipients of the selected message or the organizer and all attendees of the selected appointment.</span></span>

> [!NOTE]
> <span data-ttu-id="b89d7-603">Diese Methode wird in Outlook für iOS oder Outlook für Android nicht unterstützt.</span><span class="sxs-lookup"><span data-stu-id="b89d7-603">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="b89d7-604">In Outlook Web App wird das Antwortformular als Popupformular in der Dreispaltenansicht und als Popupformular in der Zwei- oder Einspaltenansicht angezeigt.</span><span class="sxs-lookup"><span data-stu-id="b89d7-604">In Outlook Web App, the reply form is displayed as a pop-out form in the 3-column view and a pop-up form in the 2- or 1-column view.</span></span>

<span data-ttu-id="b89d7-605">Wenn einer der Zeichenfolgenparameter seinen Grenzwert überschreitet, löst `displayReplyAllForm` eine Ausnahme aus.</span><span class="sxs-lookup"><span data-stu-id="b89d7-605">If any of the string parameters exceed their limits, `displayReplyAllForm` throws an exception.</span></span>

<span data-ttu-id="b89d7-p137">Wenn Anlagen im `formData.attachments`-Parameter angegeben werden, versuchen Outlook und Outlook Web App alle Anlagen herunterzuladen und sie an das Antwortformular anzufügen. Wenn Anlagen nicht hinzugefügt werden können, wird in der Formularbenutzeroberfläche ein Fehler ausgegeben. Wenn dies nicht möglich ist, wird keine Fehlermeldung ausgegeben.</span><span class="sxs-lookup"><span data-stu-id="b89d7-p137">When attachments are specified in the `formData.attachments` parameter, Outlook and Outlook Web App attempt to download all attachments and attach them to the reply form. If any attachments fail to be added, an error is shown in the form UI. If this isn't possible, then no error message is thrown.</span></span>

##### <a name="parameters"></a><span data-ttu-id="b89d7-609">Parameter:</span><span class="sxs-lookup"><span data-stu-id="b89d7-609">Parameters:</span></span>

|<span data-ttu-id="b89d7-610">Name</span><span class="sxs-lookup"><span data-stu-id="b89d7-610">Name</span></span>| <span data-ttu-id="b89d7-611">Typ</span><span class="sxs-lookup"><span data-stu-id="b89d7-611">Type</span></span>| <span data-ttu-id="b89d7-612">Beschreibung</span><span class="sxs-lookup"><span data-stu-id="b89d7-612">Description</span></span>|
|---|---|---|
|`formData`| <span data-ttu-id="b89d7-613">Zeichenfolge &#124; Objekt</span><span class="sxs-lookup"><span data-stu-id="b89d7-613">String &#124; Object</span></span>| |<span data-ttu-id="b89d7-p138">Eine Zeichenfolge, die Text- und HTML-Code enthält, die den Hauptteil des Antwortformulars darstellen. Die Zeichenfolge ist auf 32 KB beschränkt.</span><span class="sxs-lookup"><span data-stu-id="b89d7-p138">A string that contains text and HTML and that represents the body of the reply form. The string is limited to 32 KB.</span></span><br/><span data-ttu-id="b89d7-616">**ODER**</span><span class="sxs-lookup"><span data-stu-id="b89d7-616">**OR**</span></span><br/><span data-ttu-id="b89d7-p139">Ein Objekt, das Text- oder Anlagendaten und eine Rückruffunktion enthält. Das Objekt ist wie folgt definiert:</span><span class="sxs-lookup"><span data-stu-id="b89d7-p139">An object that contains body or attachment data and a callback function. The object is defined as follows.</span></span> |
| `formData.htmlBody` | <span data-ttu-id="b89d7-619">Zeichenfolge</span><span class="sxs-lookup"><span data-stu-id="b89d7-619">String</span></span> | <span data-ttu-id="b89d7-620">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="b89d7-620">&lt;optional&gt;</span></span> | <span data-ttu-id="b89d7-p140">Eine Zeichenfolge, die Text- und HTML-Code enthält, die den Hauptteil des Antwortformulars darstellen. Die Zeichenfolge ist auf 32 KB beschränkt.</span><span class="sxs-lookup"><span data-stu-id="b89d7-p140">A string that contains text and HTML and that represents the body of the reply form. The string is limited to 32 KB.</span></span>
| `formData.attachments` | <span data-ttu-id="b89d7-623">Array.&lt;Object&gt;</span><span class="sxs-lookup"><span data-stu-id="b89d7-623">Array.&lt;Object&gt;</span></span> | <span data-ttu-id="b89d7-624">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="b89d7-624">&lt;optional&gt;</span></span> | <span data-ttu-id="b89d7-625">Ein Array mit JSON-Objekten, die Datei- oder Elementanlagen sind.</span><span class="sxs-lookup"><span data-stu-id="b89d7-625">An array of JSON objects that are either file or item attachments.</span></span> |
| `formData.attachments.type` | <span data-ttu-id="b89d7-626">Zeichenfolge</span><span class="sxs-lookup"><span data-stu-id="b89d7-626">String</span></span> | | <span data-ttu-id="b89d7-p141">Gibt den Typ der Anlage an. Muss `file` für eine Dateianlage oder `item` für eine Elementanlage sein.</span><span class="sxs-lookup"><span data-stu-id="b89d7-p141">Indicates the type of attachment. Must be `file` for a file attachment or `item` for an item attachment.</span></span> |
| `formData.attachments.name` | <span data-ttu-id="b89d7-629">Zeichenfolge</span><span class="sxs-lookup"><span data-stu-id="b89d7-629">String</span></span> | | <span data-ttu-id="b89d7-630">Eine Zeichenfolge, die den Namen der Anlage mit bis zu 255 Zeichen enthält.</span><span class="sxs-lookup"><span data-stu-id="b89d7-630">A string that contains the name of the attachment, up to 255 characters in length.</span></span>|
| `formData.attachments.url` | <span data-ttu-id="b89d7-631">Zeichenfolge</span><span class="sxs-lookup"><span data-stu-id="b89d7-631">String</span></span> | | <span data-ttu-id="b89d7-p142">Wird nur verwendet, wenn `type` auf `file` gesetzt ist. Der URI des Speicherorts für die Datei.</span><span class="sxs-lookup"><span data-stu-id="b89d7-p142">Only used if `type` is set to `file`. The URI of the location for the file.</span></span> |
| `formData.attachments.itemId` | <span data-ttu-id="b89d7-634">Zeichenfolge</span><span class="sxs-lookup"><span data-stu-id="b89d7-634">String</span></span> | | <span data-ttu-id="b89d7-p143">Wird nur verwendet, wenn `type` auf `item` gesetzt ist. Die EWS-Element-ID der Anlage. Dies ist eine Zeichenfolge bis zu 100 Zeichen.</span><span class="sxs-lookup"><span data-stu-id="b89d7-p143">Only used if `type` is set to `item`. The EWS item id of the attachment. This is a string up to 100 characters.</span></span> |
| `callback` | <span data-ttu-id="b89d7-638">Funktion</span><span class="sxs-lookup"><span data-stu-id="b89d7-638">function</span></span> | <span data-ttu-id="b89d7-639">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="b89d7-639">&lt;optional&gt;</span></span> | <span data-ttu-id="b89d7-640">Bei Abschluss der Methode wird die im Parameter `callback` übergebene Funktion mit einem einzigen Parameter aufgerufen: `asyncResult`. Dieser Parameter ist ein Objekt des Typs [AsyncResult](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="b89d7-640">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [AsyncResult](/javascript/api/office/office.asyncresult) object.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="b89d7-641">Anforderungen</span><span class="sxs-lookup"><span data-stu-id="b89d7-641">Requirements</span></span>

|<span data-ttu-id="b89d7-642">Anforderung</span><span class="sxs-lookup"><span data-stu-id="b89d7-642">Requirement</span></span>| <span data-ttu-id="b89d7-643">Wert</span><span class="sxs-lookup"><span data-stu-id="b89d7-643">Value</span></span>|
|---|---|
|[<span data-ttu-id="b89d7-644">Mindestversion des Postfachanforderungssatzes</span><span class="sxs-lookup"><span data-stu-id="b89d7-644">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="b89d7-645">1.0</span><span class="sxs-lookup"><span data-stu-id="b89d7-645">1.0</span></span>|
|[<span data-ttu-id="b89d7-646">Mindestberechtigungsstufe</span><span class="sxs-lookup"><span data-stu-id="b89d7-646">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="b89d7-647">ReadItem</span><span class="sxs-lookup"><span data-stu-id="b89d7-647">ReadItem</span></span>|
|[<span data-ttu-id="b89d7-648">Zutreffender Outlook-Modus</span><span class="sxs-lookup"><span data-stu-id="b89d7-648">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="b89d7-649">Lesen</span><span class="sxs-lookup"><span data-stu-id="b89d7-649">Read</span></span>|

##### <a name="examples"></a><span data-ttu-id="b89d7-650">Beispiele</span><span class="sxs-lookup"><span data-stu-id="b89d7-650">Examples</span></span>

<span data-ttu-id="b89d7-651">Im folgenden Code wird eine Zeichenfolge an die `displayReplyAllForm`-Funktion übergeben.</span><span class="sxs-lookup"><span data-stu-id="b89d7-651">The following code passes a string to the `displayReplyAllForm` function.</span></span>

```JavaScript
Office.context.mailbox.item.displayReplyAllForm('hello there');
Office.context.mailbox.item.displayReplyAllForm('<b>hello there</b>');
```

<span data-ttu-id="b89d7-652">Antworten mit einem leeren Textkörper.</span><span class="sxs-lookup"><span data-stu-id="b89d7-652">Reply with an empty body.</span></span>

```JavaScript
Office.context.mailbox.item.displayReplyAllForm({});
```

<span data-ttu-id="b89d7-653">Antworten nur einem Textkörper.</span><span class="sxs-lookup"><span data-stu-id="b89d7-653">Reply with just a body.</span></span>

```JavaScript
Office.context.mailbox.item.displayReplyAllForm(
{
  'htmlBody' : 'hi'
});
```

<span data-ttu-id="b89d7-654">Antworten mit einem Textkörper und einer Dateianlage.</span><span class="sxs-lookup"><span data-stu-id="b89d7-654">Reply with a body and a file attachment.</span></span>

```JavaScript
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

<span data-ttu-id="b89d7-655">Antworten mit einem Textkörper und einer Elementanlage.</span><span class="sxs-lookup"><span data-stu-id="b89d7-655">Reply with a body and an item attachment.</span></span>

```JavaScript
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

<span data-ttu-id="b89d7-656">Antworten Sie mit einem Textkörper, einer Dateianlage, einer Elementanlage und einem Callback.</span><span class="sxs-lookup"><span data-stu-id="b89d7-656">Reply with a body, file attachment, item attachment, and a callback.</span></span>

```JavaScript
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

#### <a name="displayreplyformformdata"></a><span data-ttu-id="b89d7-657">displayReplyForm(formData)</span><span class="sxs-lookup"><span data-stu-id="b89d7-657">displayReplyForm(formData)</span></span>

<span data-ttu-id="b89d7-658">Zeigt ein Antwortformular an, das nur den Absender der ausgewählten Nachricht oder den Organisator des ausgewählten Termins enthält.</span><span class="sxs-lookup"><span data-stu-id="b89d7-658">Displays a reply form that includes only the sender of the selected message or the organizer of the selected appointment.</span></span>

> [!NOTE]
> <span data-ttu-id="b89d7-659">Diese Methode wird in Outlook für iOS oder Outlook für Android nicht unterstützt.</span><span class="sxs-lookup"><span data-stu-id="b89d7-659">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="b89d7-660">In Outlook Web App wird das Antwortformular als Popupformular in der Dreispaltenansicht und als Popupformular in der Zwei- oder Einspaltenansicht angezeigt.</span><span class="sxs-lookup"><span data-stu-id="b89d7-660">In Outlook Web App, the reply form is displayed as a pop-out form in the 3-column view and a pop-up form in the 2- or 1-column view.</span></span>

<span data-ttu-id="b89d7-661">Wenn einer der Zeichenfolgenparameter seinen Grenzwert überschreitet, löst `displayReplyForm` eine Ausnahme aus.</span><span class="sxs-lookup"><span data-stu-id="b89d7-661">If any of the string parameters exceed their limits, `displayReplyForm` throws an exception.</span></span>

<span data-ttu-id="b89d7-p144">Wenn Anlagen im `formData.attachments`-Parameter angegeben werden, versuchen Outlook und Outlook Web App alle Anlagen herunterzuladen und sie an das Antwortformular anzufügen. Wenn Anlagen nicht hinzugefügt werden können, wird in der Formularbenutzeroberfläche ein Fehler ausgegeben. Wenn dies nicht möglich ist, wird keine Fehlermeldung ausgegeben.</span><span class="sxs-lookup"><span data-stu-id="b89d7-p144">When attachments are specified in the `formData.attachments` parameter, Outlook and Outlook Web App attempt to download all attachments and attach them to the reply form. If any attachments fail to be added, an error is shown in the form UI. If this isn't possible, then no error message is thrown.</span></span>

##### <a name="parameters"></a><span data-ttu-id="b89d7-665">Parameter:</span><span class="sxs-lookup"><span data-stu-id="b89d7-665">Parameters:</span></span>

|<span data-ttu-id="b89d7-666">Name</span><span class="sxs-lookup"><span data-stu-id="b89d7-666">Name</span></span>| <span data-ttu-id="b89d7-667">Typ</span><span class="sxs-lookup"><span data-stu-id="b89d7-667">Type</span></span>| <span data-ttu-id="b89d7-668">Beschreibung</span><span class="sxs-lookup"><span data-stu-id="b89d7-668">Description</span></span>|
|---|---|---|
|`formData`| <span data-ttu-id="b89d7-669">Zeichenfolge &#124; Objekt</span><span class="sxs-lookup"><span data-stu-id="b89d7-669">String &#124; Object</span></span>| | <span data-ttu-id="b89d7-p145">Eine Zeichenfolge, die Text- und HTML-Code enthält, die den Hauptteil des Antwortformulars darstellen. Die Zeichenfolge ist auf 32 KB beschränkt.</span><span class="sxs-lookup"><span data-stu-id="b89d7-p145">A string that contains text and HTML and that represents the body of the reply form. The string is limited to 32 KB.</span></span><br/><span data-ttu-id="b89d7-672">**ODER**</span><span class="sxs-lookup"><span data-stu-id="b89d7-672">**OR**</span></span><br/><span data-ttu-id="b89d7-p146">Ein Objekt, das Text- oder Anlagendaten und eine Rückruffunktion enthält. Das Objekt ist wie folgt definiert:</span><span class="sxs-lookup"><span data-stu-id="b89d7-p146">An object that contains body or attachment data and a callback function. The object is defined as follows.</span></span> |
| `formData.htmlBody` | <span data-ttu-id="b89d7-675">Zeichenfolge</span><span class="sxs-lookup"><span data-stu-id="b89d7-675">String</span></span> | <span data-ttu-id="b89d7-676">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="b89d7-676">&lt;optional&gt;</span></span> | <span data-ttu-id="b89d7-p147">Eine Zeichenfolge, die Text- und HTML-Code enthält, die den Hauptteil des Antwortformulars darstellen. Die Zeichenfolge ist auf 32 KB beschränkt.</span><span class="sxs-lookup"><span data-stu-id="b89d7-p147">A string that contains text and HTML and that represents the body of the reply form. The string is limited to 32 KB.</span></span>
| `formData.attachments` | <span data-ttu-id="b89d7-679">Array.&lt;Object&gt;</span><span class="sxs-lookup"><span data-stu-id="b89d7-679">Array.&lt;Object&gt;</span></span> | <span data-ttu-id="b89d7-680">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="b89d7-680">&lt;optional&gt;</span></span> | <span data-ttu-id="b89d7-681">Ein Array mit JSON-Objekten, die Datei- oder Elementanlagen sind.</span><span class="sxs-lookup"><span data-stu-id="b89d7-681">An array of JSON objects that are either file or item attachments.</span></span> |
| `formData.attachments.type` | <span data-ttu-id="b89d7-682">Zeichenfolge</span><span class="sxs-lookup"><span data-stu-id="b89d7-682">String</span></span> | | <span data-ttu-id="b89d7-p148">Gibt den Typ der Anlage an. Muss `file` für eine Dateianlage oder `item` für eine Elementanlage sein.</span><span class="sxs-lookup"><span data-stu-id="b89d7-p148">Indicates the type of attachment. Must be `file` for a file attachment or `item` for an item attachment.</span></span> |
| `formData.attachments.name` | <span data-ttu-id="b89d7-685">Zeichenfolge</span><span class="sxs-lookup"><span data-stu-id="b89d7-685">String</span></span> | | <span data-ttu-id="b89d7-686">Eine Zeichenfolge, die den Namen der Anlage mit bis zu 255 Zeichen enthält.</span><span class="sxs-lookup"><span data-stu-id="b89d7-686">A string that contains the name of the attachment, up to 255 characters in length.</span></span>|
| `formData.attachments.url` | <span data-ttu-id="b89d7-687">Zeichenfolge</span><span class="sxs-lookup"><span data-stu-id="b89d7-687">String</span></span> | | <span data-ttu-id="b89d7-p149">Wird nur verwendet, wenn `type` auf `file` gesetzt ist. Der URI des Speicherorts für die Datei.</span><span class="sxs-lookup"><span data-stu-id="b89d7-p149">Only used if `type` is set to `file`. The URI of the location for the file.</span></span> |
| `formData.attachments.itemId` | <span data-ttu-id="b89d7-690">Zeichenfolge</span><span class="sxs-lookup"><span data-stu-id="b89d7-690">String</span></span> | | <span data-ttu-id="b89d7-p150">Wird nur verwendet, wenn `type` auf `item` gesetzt ist. Die EWS-Element-ID der Anlage. Dies ist eine Zeichenfolge bis zu 100 Zeichen.</span><span class="sxs-lookup"><span data-stu-id="b89d7-p150">Only used if `type` is set to `item`. The EWS item id of the attachment. This is a string up to 100 characters.</span></span> |
| `callback` | <span data-ttu-id="b89d7-694">Funktion</span><span class="sxs-lookup"><span data-stu-id="b89d7-694">function</span></span> | <span data-ttu-id="b89d7-695">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="b89d7-695">&lt;optional&gt;</span></span> | <span data-ttu-id="b89d7-696">Bei Abschluss der Methode wird die im Parameter `callback` übergebene Funktion mit einem einzigen Parameter aufgerufen: `asyncResult`. Dieser Parameter ist ein Objekt des Typs [AsyncResult](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="b89d7-696">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [AsyncResult](/javascript/api/office/office.asyncresult) object.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="b89d7-697">Anforderungen</span><span class="sxs-lookup"><span data-stu-id="b89d7-697">Requirements</span></span>

|<span data-ttu-id="b89d7-698">Anforderung</span><span class="sxs-lookup"><span data-stu-id="b89d7-698">Requirement</span></span>| <span data-ttu-id="b89d7-699">Wert</span><span class="sxs-lookup"><span data-stu-id="b89d7-699">Value</span></span>|
|---|---|
|[<span data-ttu-id="b89d7-700">Mindestversion des Postfachanforderungssatzes</span><span class="sxs-lookup"><span data-stu-id="b89d7-700">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="b89d7-701">1.0</span><span class="sxs-lookup"><span data-stu-id="b89d7-701">1.0</span></span>|
|[<span data-ttu-id="b89d7-702">Mindestberechtigungsstufe</span><span class="sxs-lookup"><span data-stu-id="b89d7-702">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="b89d7-703">ReadItem</span><span class="sxs-lookup"><span data-stu-id="b89d7-703">ReadItem</span></span>|
|[<span data-ttu-id="b89d7-704">Zutreffender Outlook-Modus</span><span class="sxs-lookup"><span data-stu-id="b89d7-704">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="b89d7-705">Lesen</span><span class="sxs-lookup"><span data-stu-id="b89d7-705">Read</span></span>|

##### <a name="examples"></a><span data-ttu-id="b89d7-706">Beispiele</span><span class="sxs-lookup"><span data-stu-id="b89d7-706">Examples</span></span>

<span data-ttu-id="b89d7-707">Im folgenden Code wird eine Zeichenfolge an die `displayReplyForm`-Funktion übergeben.</span><span class="sxs-lookup"><span data-stu-id="b89d7-707">The following code passes a string to the `displayReplyForm` function.</span></span>

```JavaScript
Office.context.mailbox.item.displayReplyForm('hello there');
Office.context.mailbox.item.displayReplyForm('<b>hello there</b>');
```

<span data-ttu-id="b89d7-708">Antworten mit einem leeren Textkörper.</span><span class="sxs-lookup"><span data-stu-id="b89d7-708">Reply with an empty body.</span></span>

```JavaScript
Office.context.mailbox.item.displayReplyForm({});
```

<span data-ttu-id="b89d7-709">Antworten nur einem Textkörper.</span><span class="sxs-lookup"><span data-stu-id="b89d7-709">Reply with just a body.</span></span>

```JavaScript
Office.context.mailbox.item.displayReplyForm(
{
  'htmlBody' : 'hi'
});
```

<span data-ttu-id="b89d7-710">Antworten mit einem Textkörper und einer Dateianlage.</span><span class="sxs-lookup"><span data-stu-id="b89d7-710">Reply with a body and a file attachment.</span></span>

```JavaScript
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

<span data-ttu-id="b89d7-711">Antworten mit einem Textkörper und einer Elementanlage.</span><span class="sxs-lookup"><span data-stu-id="b89d7-711">Reply with a body and an item attachment.</span></span>

```JavaScript
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

<span data-ttu-id="b89d7-712">Antworten Sie mit einem Textkörper, einer Dateianlage, einer Elementanlage und einem Callback.</span><span class="sxs-lookup"><span data-stu-id="b89d7-712">Reply with a body, file attachment, item attachment, and a callback.</span></span>

```JavaScript
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

#### <a name="getentities--entitiesjavascriptapioutlook12officeentities"></a><span data-ttu-id="b89d7-713">getEntities() → {[Entities](/javascript/api/outlook_1_2/office.entities)}</span><span class="sxs-lookup"><span data-stu-id="b89d7-713">getEntities() → {[Entities](/javascript/api/outlook_1_2/office.entities)}</span></span>

<span data-ttu-id="b89d7-714">Ruft das ausgewählte Element Textkörper gefundenen Entitäten ab.</span><span class="sxs-lookup"><span data-stu-id="b89d7-714">Gets the entities found in the selected item's body.</span></span>

> [!NOTE]
> <span data-ttu-id="b89d7-715">Diese Methode wird in Outlook für iOS oder Outlook für Android nicht unterstützt.</span><span class="sxs-lookup"><span data-stu-id="b89d7-715">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

##### <a name="requirements"></a><span data-ttu-id="b89d7-716">Anforderungen</span><span class="sxs-lookup"><span data-stu-id="b89d7-716">Requirements</span></span>

|<span data-ttu-id="b89d7-717">Anforderung</span><span class="sxs-lookup"><span data-stu-id="b89d7-717">Requirement</span></span>| <span data-ttu-id="b89d7-718">Wert</span><span class="sxs-lookup"><span data-stu-id="b89d7-718">Value</span></span>|
|---|---|
|[<span data-ttu-id="b89d7-719">Mindestversion des Postfachanforderungssatzes</span><span class="sxs-lookup"><span data-stu-id="b89d7-719">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="b89d7-720">1.0</span><span class="sxs-lookup"><span data-stu-id="b89d7-720">1.0</span></span>|
|[<span data-ttu-id="b89d7-721">Mindestberechtigungsstufe</span><span class="sxs-lookup"><span data-stu-id="b89d7-721">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="b89d7-722">ReadItem</span><span class="sxs-lookup"><span data-stu-id="b89d7-722">ReadItem</span></span>|
|[<span data-ttu-id="b89d7-723">Zutreffender Outlook-Modus</span><span class="sxs-lookup"><span data-stu-id="b89d7-723">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="b89d7-724">Lesen</span><span class="sxs-lookup"><span data-stu-id="b89d7-724">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="b89d7-725">Gibt zurück:</span><span class="sxs-lookup"><span data-stu-id="b89d7-725">Returns:</span></span>

<span data-ttu-id="b89d7-726">Typ: [Entitäten](/javascript/api/outlook_1_2/office.entities)</span><span class="sxs-lookup"><span data-stu-id="b89d7-726">Type: [Entities](/javascript/api/outlook_1_2/office.entities)</span></span>

##### <a name="example"></a><span data-ttu-id="b89d7-727">Beispiel</span><span class="sxs-lookup"><span data-stu-id="b89d7-727">Example</span></span>

<span data-ttu-id="b89d7-728">Das folgende Beispiel greift auf die Kontakte Entitäten im Textkörper des aktuellen Elements.</span><span class="sxs-lookup"><span data-stu-id="b89d7-728">The following example accesses the contacts entities in the current item's body.</span></span>

```
var contacts = Office.context.mailbox.item.getEntities().contacts;
```

#### <a name="getentitiesbytypeentitytype--nullable-arraystringcontactjavascriptapioutlook12officecontactmeetingsuggestionjavascriptapioutlook12officemeetingsuggestionphonenumberjavascriptapioutlook12officephonenumbertasksuggestionjavascriptapioutlook12officetasksuggestion"></a><span data-ttu-id="b89d7-729">getEntitiesByType(entityType) → (nullable) {Array.<(String|[Contact](/javascript/api/outlook_1_2/office.contact)|[MeetingSuggestion](/javascript/api/outlook_1_2/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook_1_2/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook_1_2/office.tasksuggestion))>}</span><span class="sxs-lookup"><span data-stu-id="b89d7-729">getEntitiesByType(entityType) → (nullable) {Array.<(String|[Contact](/javascript/api/outlook_1_2/office.contact)|[MeetingSuggestion](/javascript/api/outlook_1_2/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook_1_2/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook_1_2/office.tasksuggestion))>}</span></span>

<span data-ttu-id="b89d7-730">Ruft ein Array aller Entitäten des angegebenen Entitätstyps im Hauptteil des ausgewählten Elements gefunden.</span><span class="sxs-lookup"><span data-stu-id="b89d7-730">Gets an array of all the entities of the specified entity type found in the selected item's body.</span></span>

> [!NOTE]
> <span data-ttu-id="b89d7-731">Diese Methode wird in Outlook für iOS oder Outlook für Android nicht unterstützt.</span><span class="sxs-lookup"><span data-stu-id="b89d7-731">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

##### <a name="parameters"></a><span data-ttu-id="b89d7-732">Parameter:</span><span class="sxs-lookup"><span data-stu-id="b89d7-732">Parameters:</span></span>

|<span data-ttu-id="b89d7-733">Name</span><span class="sxs-lookup"><span data-stu-id="b89d7-733">Name</span></span>| <span data-ttu-id="b89d7-734">Typ</span><span class="sxs-lookup"><span data-stu-id="b89d7-734">Type</span></span>| <span data-ttu-id="b89d7-735">Beschreibung</span><span class="sxs-lookup"><span data-stu-id="b89d7-735">Description</span></span>|
|---|---|---|
|`entityType`| [<span data-ttu-id="b89d7-736">Office.MailboxEnums.EntityType</span><span class="sxs-lookup"><span data-stu-id="b89d7-736">Office.MailboxEnums.EntityType</span></span>](/javascript/api/outlook_1_2/office.mailboxenums.entitytype)|<span data-ttu-id="b89d7-737">Einer der Werte der EntityType-Enumeration.</span><span class="sxs-lookup"><span data-stu-id="b89d7-737">One of the EntityType enumeration values.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="b89d7-738">Anforderungen</span><span class="sxs-lookup"><span data-stu-id="b89d7-738">Requirements</span></span>

|<span data-ttu-id="b89d7-739">Anforderung</span><span class="sxs-lookup"><span data-stu-id="b89d7-739">Requirement</span></span>| <span data-ttu-id="b89d7-740">Wert</span><span class="sxs-lookup"><span data-stu-id="b89d7-740">Value</span></span>|
|---|---|
|[<span data-ttu-id="b89d7-741">Mindestversion des Postfachanforderungssatzes</span><span class="sxs-lookup"><span data-stu-id="b89d7-741">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="b89d7-742">1.0</span><span class="sxs-lookup"><span data-stu-id="b89d7-742">1.0</span></span>|
|[<span data-ttu-id="b89d7-743">Mindestberechtigungsstufe</span><span class="sxs-lookup"><span data-stu-id="b89d7-743">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="b89d7-744">Eingeschränkt</span><span class="sxs-lookup"><span data-stu-id="b89d7-744">Restricted</span></span>|
|[<span data-ttu-id="b89d7-745">Zutreffender Outlook-Modus</span><span class="sxs-lookup"><span data-stu-id="b89d7-745">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="b89d7-746">Lesen</span><span class="sxs-lookup"><span data-stu-id="b89d7-746">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="b89d7-747">Gibt zurück:</span><span class="sxs-lookup"><span data-stu-id="b89d7-747">Returns:</span></span>

<span data-ttu-id="b89d7-748">Wenn der in `entityType` übergebene Wert kein gültiges Element der `EntityType`-Enumeration ist, gibt die Methode null zurück.</span><span class="sxs-lookup"><span data-stu-id="b89d7-748">If the value passed in `entityType` is not a valid member of the `EntityType` enumeration, the method returns null.</span></span> <span data-ttu-id="b89d7-749">Wenn keine Entitäten des angegebenen Typs im Textkörper des Elements vorhanden sind, gibt die Methode ein leeres Array zurück.</span><span class="sxs-lookup"><span data-stu-id="b89d7-749">If no entities of the specified type are present in the item's body, the method returns an empty array.</span></span> <span data-ttu-id="b89d7-750">Andernfalls hängt der Typ der Objekte im zurückgegebenen Array vom Typ der Entität ab, die im `entityType`-Parameter angefordert wurde.</span><span class="sxs-lookup"><span data-stu-id="b89d7-750">Otherwise, the type of the objects in the returned array depends on the type of entity requested in the `entityType` parameter.</span></span>

<span data-ttu-id="b89d7-751">Während Sie die minimale Berechtigungsstufe für diese Methode **Restricted** ist, erfordern einige Entitätstypen **ReadItem** für den Zugriff, wie in der folgenden Tabelle angegeben wird.</span><span class="sxs-lookup"><span data-stu-id="b89d7-751">While the minimum permission level to use this method is **Restricted**, some entity types require **ReadItem** to access, as specified in the following table.</span></span>

| <span data-ttu-id="b89d7-752">Wert von `entityType`</span><span class="sxs-lookup"><span data-stu-id="b89d7-752">Value of `entityType`</span></span> | <span data-ttu-id="b89d7-753">Typ der Objekte im zurückgegebenen Array</span><span class="sxs-lookup"><span data-stu-id="b89d7-753">Type of objects in returned array</span></span> | <span data-ttu-id="b89d7-754">Erforderliche Berechtigungsstufe</span><span class="sxs-lookup"><span data-stu-id="b89d7-754">Required Permission Level</span></span> |
| --- | --- | --- |
| `Address` | <span data-ttu-id="b89d7-755">Zeichenfolge</span><span class="sxs-lookup"><span data-stu-id="b89d7-755">String</span></span> | <span data-ttu-id="b89d7-756">**Restricted**</span><span class="sxs-lookup"><span data-stu-id="b89d7-756">**Restricted**</span></span> |
| `Contact` | <span data-ttu-id="b89d7-757">Contact</span><span class="sxs-lookup"><span data-stu-id="b89d7-757">Contact</span></span> | <span data-ttu-id="b89d7-758">**ReadItem**</span><span class="sxs-lookup"><span data-stu-id="b89d7-758">**ReadItem**</span></span> |
| `EmailAddress` | <span data-ttu-id="b89d7-759">Zeichenfolge</span><span class="sxs-lookup"><span data-stu-id="b89d7-759">String</span></span> | <span data-ttu-id="b89d7-760">**ReadItem**</span><span class="sxs-lookup"><span data-stu-id="b89d7-760">**ReadItem**</span></span> |
| `MeetingSuggestion` | <span data-ttu-id="b89d7-761">MeetingSuggestion</span><span class="sxs-lookup"><span data-stu-id="b89d7-761">MeetingSuggestion</span></span> | <span data-ttu-id="b89d7-762">**ReadItem**</span><span class="sxs-lookup"><span data-stu-id="b89d7-762">**ReadItem**</span></span> |
| `PhoneNumber` | <span data-ttu-id="b89d7-763">PhoneNumber</span><span class="sxs-lookup"><span data-stu-id="b89d7-763">PhoneNumber</span></span> | <span data-ttu-id="b89d7-764">**Restricted**</span><span class="sxs-lookup"><span data-stu-id="b89d7-764">**Restricted**</span></span> |
| `TaskSuggestion` | <span data-ttu-id="b89d7-765">TaskSuggestion</span><span class="sxs-lookup"><span data-stu-id="b89d7-765">TaskSuggestion</span></span> | <span data-ttu-id="b89d7-766">**ReadItem**</span><span class="sxs-lookup"><span data-stu-id="b89d7-766">**ReadItem**</span></span> |
| `URL` | <span data-ttu-id="b89d7-767">Zeichenfolge</span><span class="sxs-lookup"><span data-stu-id="b89d7-767">String</span></span> | <span data-ttu-id="b89d7-768">**Restricted**</span><span class="sxs-lookup"><span data-stu-id="b89d7-768">**Restricted**</span></span> |

<span data-ttu-id="b89d7-769">Typ: Array.<(String|[Contact](/javascript/api/outlook_1_2/office.contact)|[MeetingSuggestion](/javascript/api/outlook_1_2/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook_1_2/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook_1_2/office.tasksuggestion))></span><span class="sxs-lookup"><span data-stu-id="b89d7-769">Type: Array.<(String|[Contact](/javascript/api/outlook_1_2/office.contact)|[MeetingSuggestion](/javascript/api/outlook_1_2/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook_1_2/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook_1_2/office.tasksuggestion))></span></span>

##### <a name="example"></a><span data-ttu-id="b89d7-770">Beispiel</span><span class="sxs-lookup"><span data-stu-id="b89d7-770">Example</span></span>

<span data-ttu-id="b89d7-771">Im folgenden Beispiel wird veranschaulicht, wie auf ein Array von Zeichenfolgen, die Postadressen im Textkörper des aktuellen Elements darstellen.</span><span class="sxs-lookup"><span data-stu-id="b89d7-771">The following example shows how to access an array of strings that represent postal addresses in the current item's body.</span></span>

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

#### <a name="getfilteredentitiesbynamename--nullable-arraystringcontactjavascriptapioutlook12officecontactmeetingsuggestionjavascriptapioutlook12officemeetingsuggestionphonenumberjavascriptapioutlook12officephonenumbertasksuggestionjavascriptapioutlook12officetasksuggestion"></a><span data-ttu-id="b89d7-772">getFilteredEntitiesByName(name) → (nullable) {Array.<(String|[Contact](/javascript/api/outlook_1_2/office.contact)|[MeetingSuggestion](/javascript/api/outlook_1_2/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook_1_2/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook_1_2/office.tasksuggestion))>}</span><span class="sxs-lookup"><span data-stu-id="b89d7-772">getFilteredEntitiesByName(name) → (nullable) {Array.<(String|[Contact](/javascript/api/outlook_1_2/office.contact)|[MeetingSuggestion](/javascript/api/outlook_1_2/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook_1_2/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook_1_2/office.tasksuggestion))>}</span></span>

<span data-ttu-id="b89d7-773">Gibt bekannte Entitäten im ausgewählten Element zurück, die den in der XML-Manifestdatei definierten benannten Filter übergeben.</span><span class="sxs-lookup"><span data-stu-id="b89d7-773">Returns well-known entities in the selected item that pass the named filter defined in the manifest XML file.</span></span>

> [!NOTE]
> <span data-ttu-id="b89d7-774">Diese Methode wird in Outlook für iOS oder Outlook für Android nicht unterstützt.</span><span class="sxs-lookup"><span data-stu-id="b89d7-774">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="b89d7-775">Die `getFilteredEntitiesByName`-Methode gibt die Entitäten zurück, die dem im [ItemHasKnownEntity](/javascript/office/manifest/rule#itemhasknownentity-rule)-Regelelement der XML-Manifestdatei definierten regulären Ausdruck mit dem angegebenen `FilterName`-Elementwert entsprechen.</span><span class="sxs-lookup"><span data-stu-id="b89d7-775">The `getFilteredEntitiesByName` method returns the entities that match the regular expression defined in the [ItemHasKnownEntity](/javascript/office/manifest/rule#itemhasknownentity-rule) rule element in the manifest XML file with the specified `FilterName` element value.</span></span>

##### <a name="parameters"></a><span data-ttu-id="b89d7-776">Parameter:</span><span class="sxs-lookup"><span data-stu-id="b89d7-776">Parameters:</span></span>

|<span data-ttu-id="b89d7-777">Name</span><span class="sxs-lookup"><span data-stu-id="b89d7-777">Name</span></span>| <span data-ttu-id="b89d7-778">Typ</span><span class="sxs-lookup"><span data-stu-id="b89d7-778">Type</span></span>| <span data-ttu-id="b89d7-779">Beschreibung</span><span class="sxs-lookup"><span data-stu-id="b89d7-779">Description</span></span>|
|---|---|---|
|`name`| <span data-ttu-id="b89d7-780">Zeichenfolge</span><span class="sxs-lookup"><span data-stu-id="b89d7-780">String</span></span>|<span data-ttu-id="b89d7-781">Der Name des `ItemHasKnownEntity`-Regelelements, das den entsprechenden Filter definiert.</span><span class="sxs-lookup"><span data-stu-id="b89d7-781">The name of the `ItemHasKnownEntity` rule element that defines the filter to match.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="b89d7-782">Anforderungen</span><span class="sxs-lookup"><span data-stu-id="b89d7-782">Requirements</span></span>

|<span data-ttu-id="b89d7-783">Anforderung</span><span class="sxs-lookup"><span data-stu-id="b89d7-783">Requirement</span></span>| <span data-ttu-id="b89d7-784">Wert</span><span class="sxs-lookup"><span data-stu-id="b89d7-784">Value</span></span>|
|---|---|
|[<span data-ttu-id="b89d7-785">Mindestversion des Postfachanforderungssatzes</span><span class="sxs-lookup"><span data-stu-id="b89d7-785">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="b89d7-786">1.0</span><span class="sxs-lookup"><span data-stu-id="b89d7-786">1.0</span></span>|
|[<span data-ttu-id="b89d7-787">Mindestberechtigungsstufe</span><span class="sxs-lookup"><span data-stu-id="b89d7-787">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="b89d7-788">ReadItem</span><span class="sxs-lookup"><span data-stu-id="b89d7-788">ReadItem</span></span>|
|[<span data-ttu-id="b89d7-789">Zutreffender Outlook-Modus</span><span class="sxs-lookup"><span data-stu-id="b89d7-789">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="b89d7-790">Lesen</span><span class="sxs-lookup"><span data-stu-id="b89d7-790">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="b89d7-791">Gibt zurück:</span><span class="sxs-lookup"><span data-stu-id="b89d7-791">Returns:</span></span>

<span data-ttu-id="b89d7-p152">Befindet sich kein `ItemHasKnownEntity`-Element im Manifest mit einem `FilterName`-Elementwert, der dem `name`-Parameter entspricht, gibt die Methode `null` zurück. Wenn der `name`-Parameter einem `ItemHasKnownEntity`-Element im Manifest entspricht, es aber keine entsprechenden Entitäten im aktuellen Element gibt, gibt die Methode ein leeres Array zurück.</span><span class="sxs-lookup"><span data-stu-id="b89d7-p152">If there is no `ItemHasKnownEntity` element in the manifest with a `FilterName` element value that matches the `name` parameter, the method returns `null`. If the `name` parameter does match an `ItemHasKnownEntity` element in the manifest, but there are no entities in the current item that match, the method return an empty array.</span></span>

<span data-ttu-id="b89d7-794">Typ: Array.<(String|[Contact](/javascript/api/outlook_1_2/office.contact)|[MeetingSuggestion](/javascript/api/outlook_1_2/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook_1_2/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook_1_2/office.tasksuggestion))></span><span class="sxs-lookup"><span data-stu-id="b89d7-794">Type: Array.<(String|[Contact](/javascript/api/outlook_1_2/office.contact)|[MeetingSuggestion](/javascript/api/outlook_1_2/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook_1_2/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook_1_2/office.tasksuggestion))></span></span>

#### <a name="getregexmatches--object"></a><span data-ttu-id="b89d7-795">getRegExMatches() → {Object}</span><span class="sxs-lookup"><span data-stu-id="b89d7-795">getRegExMatches() → {Object}</span></span>

<span data-ttu-id="b89d7-796">Gibt Zeichenfolgenwerte im ausgewählten Element zurück, die den in der XML-Manifestdatei definierten regulären Ausdrücken entsprechen.</span><span class="sxs-lookup"><span data-stu-id="b89d7-796">Returns string values in the selected item that match the regular expressions defined in the manifest XML file.</span></span>

> [!NOTE]
> <span data-ttu-id="b89d7-797">Diese Methode wird in Outlook für iOS oder Outlook für Android nicht unterstützt.</span><span class="sxs-lookup"><span data-stu-id="b89d7-797">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="b89d7-p153">Die `getRegExMatches`-Methode gibt die Zeichenfolgen zurück, die dem im `ItemHasRegularExpressionMatch`- oder `ItemHasKnownEntity`-Regelelement der XML-Manifestdatei definierten regulären Ausdruck entsprechen. Bei einer `ItemHasRegularExpressionMatch`-Regel muss eine entsprechende Zeichenfolge in der Eigenschaft des Elements vorhanden sein, das von dieser Regel angegeben wird. Der einfache `PropertyName`-Typ definiert die unterstützten Eigenschaften.</span><span class="sxs-lookup"><span data-stu-id="b89d7-p153">The `getRegExMatches` method returns the strings that match the regular expression defined in each `ItemHasRegularExpressionMatch` or `ItemHasKnownEntity` rule element in the manifest XML file. For an `ItemHasRegularExpressionMatch` rule, a matching string has to occur in the property of the item that is specified by that rule. The `PropertyName` simple type defines the supported properties.</span></span>

<span data-ttu-id="b89d7-801">Nehmen Sie z. B. an, dass ein Add-In-Manifest das folgende `Rule`-Element aufweist:</span><span class="sxs-lookup"><span data-stu-id="b89d7-801">For example, consider an add-in manifest has the following `Rule` element:</span></span>

```JavaScript
<Rule xsi:type="RuleCollection" Mode="And">
  <Rule xsi:type="ItemIs" FormType="Read" ItemType="Message" />
  <Rule xsi:type="RuleCollection" Mode="Or">
    <Rule xsi:type="ItemHasRegularExpressionMatch" RegExName="fruits" RegExValue="apple|banana|coconut" PropertyName="BodyAsPlaintext" IgnoreCase="true" />
    <Rule xsi:type="ItemHasRegularExpressionMatch" RegExName="veggies" RegExValue="tomato|onion|spinach|broccoli" PropertyName="BodyAsPlaintext" IgnoreCase="true" />
  </Rule>
</Rule>
```

<span data-ttu-id="b89d7-802">Das von `getRegExMatches` zurückgegebene Objekt hätte zwei Eigenschaften: `fruits` und `veggies`.</span><span class="sxs-lookup"><span data-stu-id="b89d7-802">The object returned from `getRegExMatches` would have two properties: `fruits` and `veggies`.</span></span>

```JavaScript
{
  'fruits': ['apple','banana','Banana','coconut'],
  'veggies': ['tomato','onion','spinach','broccoli']
}
```

> [!NOTE]
> <span data-ttu-id="b89d7-p154">Wenn Sie eine `ItemHasRegularExpressionMatch`-Regel für die Textkörpereigenschaft eines Elements festlegen, sollte der reguläre Ausdruck den Textkörper weiter filtern und nicht versuchen, den gesamten Textkörper des Elements zurückzugeben. Wenn der gesamte Textkörper eines Elements mit einem regulären Ausdruck wie `.*` abgerufen wird, werden nicht immer die gewünschten Ergebnisse erzielt.</span><span class="sxs-lookup"><span data-stu-id="b89d7-p154">If you specify an `ItemHasRegularExpressionMatch` rule on the body property of an item, the regular expression should further filter the body and should not attempt to return the entire body of the item. Using a regular expression such as `.*` to obtain the entire body of an item does not always return the expected results.</span></span>

##### <a name="requirements"></a><span data-ttu-id="b89d7-805">Anforderungen</span><span class="sxs-lookup"><span data-stu-id="b89d7-805">Requirements</span></span>

|<span data-ttu-id="b89d7-806">Anforderung</span><span class="sxs-lookup"><span data-stu-id="b89d7-806">Requirement</span></span>| <span data-ttu-id="b89d7-807">Wert</span><span class="sxs-lookup"><span data-stu-id="b89d7-807">Value</span></span>|
|---|---|
|[<span data-ttu-id="b89d7-808">Mindestversion des Postfachanforderungssatzes</span><span class="sxs-lookup"><span data-stu-id="b89d7-808">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="b89d7-809">1.0</span><span class="sxs-lookup"><span data-stu-id="b89d7-809">1.0</span></span>|
|[<span data-ttu-id="b89d7-810">Mindestberechtigungsstufe</span><span class="sxs-lookup"><span data-stu-id="b89d7-810">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="b89d7-811">ReadItem</span><span class="sxs-lookup"><span data-stu-id="b89d7-811">ReadItem</span></span>|
|[<span data-ttu-id="b89d7-812">Zutreffender Outlook-Modus</span><span class="sxs-lookup"><span data-stu-id="b89d7-812">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="b89d7-813">Lesen</span><span class="sxs-lookup"><span data-stu-id="b89d7-813">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="b89d7-814">Gibt zurück:</span><span class="sxs-lookup"><span data-stu-id="b89d7-814">Returns:</span></span>

<span data-ttu-id="b89d7-p155">Ein Objekt mit Arrays aus Zeichenfolgen, die den in der XML-Manifestdatei definierten regulären Ausdrücken entsprechen. Der Name der einzelnen Arrays ist gleich dem entsprechenden Wert des `RegExName`-Attributs der entsprechenden `ItemHasRegularExpressionMatch`-Regel oder des `FilterName`-Attributs der entsprechenden `ItemHasKnownEntity`-Regel.</span><span class="sxs-lookup"><span data-stu-id="b89d7-p155">An object that contains arrays of strings that match the regular expressions defined in the manifest XML file. The name of each array is equal to the corresponding value of the `RegExName` attribute of the matching `ItemHasRegularExpressionMatch` rule or the `FilterName` attribute of the matching `ItemHasKnownEntity` rule.</span></span>

<dl class="param-type"><span data-ttu-id="b89d7-817">

<dt>
Typ</dt>


</span><span class="sxs-lookup"><span data-stu-id="b89d7-817">

<dt>Type</dt>

</span></span><dd><span data-ttu-id="b89d7-818">Object</span><span class="sxs-lookup"><span data-stu-id="b89d7-818">Object</span></span></dd>

</dl>

##### <a name="example"></a><span data-ttu-id="b89d7-819">Beispiel</span><span class="sxs-lookup"><span data-stu-id="b89d7-819">Example</span></span>

<span data-ttu-id="b89d7-820">Das folgende Beispiel zeigt, wie Sie auf das Array von Übereinstimmungen für die regulären <rule>-Ausdruckselemente `fruits` und `veggies` zugreifen, die im Manifest angegeben sind.</rule></span><span class="sxs-lookup"><span data-stu-id="b89d7-820">The following example shows how to access the array of matches for the regular expression <rule>elements `fruits` and `veggies`, which are specified in the manifest.</rule></span></span>

```JavaScript
var allMatches = Office.context.mailbox.item.getRegExMatches();
var fruits = allMatches.fruits;
var veges = allMatches.veggies;
```

#### <a name="getregexmatchesbynamename--nullable-array-string-"></a><span data-ttu-id="b89d7-821">getRegExMatchesByName(name) → (Nullwerte zulassen) {Array. < Zeichenfolge >}</span><span class="sxs-lookup"><span data-stu-id="b89d7-821">getRegExMatchesByName(name) → (nullable) {Array.< String >}</span></span>

<span data-ttu-id="b89d7-822">Gibt Zeichenfolgenwerte im ausgewählten Element zurück, die dem in der XML-Manifestdatei definierten benannten regulären Ausdruck entsprechen.</span><span class="sxs-lookup"><span data-stu-id="b89d7-822">Returns string values in the selected item that match the named regular expression defined in the manifest XML file.</span></span>

> [!NOTE]
> <span data-ttu-id="b89d7-823">Diese Methode wird in Outlook für iOS oder Outlook für Android nicht unterstützt.</span><span class="sxs-lookup"><span data-stu-id="b89d7-823">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="b89d7-824">Die `getRegExMatchesByName`-Methode gibt die Zeichenfolgen zurück, die dem im `ItemHasRegularExpressionMatch`-Regelelement der XML-Manifestdatei definierten regulären Ausdruck mit dem angegebenen `RegExName`-Elementwert entsprechen.</span><span class="sxs-lookup"><span data-stu-id="b89d7-824">The `getRegExMatchesByName` method returns the strings that match the regular expression defined in the `ItemHasRegularExpressionMatch` rule element in the manifest XML file with the specified `RegExName` element value.</span></span>

<span data-ttu-id="b89d7-p156">Wenn Sie eine `ItemHasRegularExpressionMatch`-Regel für die Textkörpereigenschaft eines Elements festlegen, sollte der reguläre Ausdruck den Textkörper weiter filtern und nicht versuchen, den gesamten Textkörper des Elements zurückzugeben. Wenn der gesamte Textkörper eines Elements mit einem regulären Ausdruck wie `.*` abgerufen wird, werden nicht immer die gewünschten Ergebnisse erzielt.</span><span class="sxs-lookup"><span data-stu-id="b89d7-p156">If you specify an `ItemHasRegularExpressionMatch` rule on the body property of an item, the regular expression should further filter the body and should not attempt to return the entire body of the item. Using a regular expression such as `.*` to obtain the entire body of an item does not always return the expected results.</span></span>

##### <a name="parameters"></a><span data-ttu-id="b89d7-827">Parameter:</span><span class="sxs-lookup"><span data-stu-id="b89d7-827">Parameters:</span></span>

|<span data-ttu-id="b89d7-828">Name</span><span class="sxs-lookup"><span data-stu-id="b89d7-828">Name</span></span>| <span data-ttu-id="b89d7-829">Typ</span><span class="sxs-lookup"><span data-stu-id="b89d7-829">Type</span></span>| <span data-ttu-id="b89d7-830">Beschreibung</span><span class="sxs-lookup"><span data-stu-id="b89d7-830">Description</span></span>|
|---|---|---|
|`name`| <span data-ttu-id="b89d7-831">Zeichenfolge</span><span class="sxs-lookup"><span data-stu-id="b89d7-831">String</span></span>|<span data-ttu-id="b89d7-832">Der Name des `ItemHasRegularExpressionMatch`-Regelelements, das den entsprechenden Filter definiert.</span><span class="sxs-lookup"><span data-stu-id="b89d7-832">The name of the `ItemHasRegularExpressionMatch` rule element that defines the filter to match.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="b89d7-833">Anforderungen</span><span class="sxs-lookup"><span data-stu-id="b89d7-833">Requirements</span></span>

|<span data-ttu-id="b89d7-834">Anforderung</span><span class="sxs-lookup"><span data-stu-id="b89d7-834">Requirement</span></span>| <span data-ttu-id="b89d7-835">Wert</span><span class="sxs-lookup"><span data-stu-id="b89d7-835">Value</span></span>|
|---|---|
|[<span data-ttu-id="b89d7-836">Mindestversion des Postfachanforderungssatzes</span><span class="sxs-lookup"><span data-stu-id="b89d7-836">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="b89d7-837">1.0</span><span class="sxs-lookup"><span data-stu-id="b89d7-837">1.0</span></span>|
|[<span data-ttu-id="b89d7-838">Mindestberechtigungsstufe</span><span class="sxs-lookup"><span data-stu-id="b89d7-838">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="b89d7-839">ReadItem</span><span class="sxs-lookup"><span data-stu-id="b89d7-839">ReadItem</span></span>|
|[<span data-ttu-id="b89d7-840">Zutreffender Outlook-Modus</span><span class="sxs-lookup"><span data-stu-id="b89d7-840">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="b89d7-841">Lesen</span><span class="sxs-lookup"><span data-stu-id="b89d7-841">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="b89d7-842">Gibt zurück:</span><span class="sxs-lookup"><span data-stu-id="b89d7-842">Returns:</span></span>

<span data-ttu-id="b89d7-843">Ein Array mit den Zeichenfolgen, die dem in der XML-Manifestdatei definierten regulären Ausdruck entsprechen.</span><span class="sxs-lookup"><span data-stu-id="b89d7-843">An array that contains the strings that match the regular expression defined in the manifest XML file.</span></span>

<dl class="param-type"><span data-ttu-id="b89d7-844">

<dt>Typ</dt>

</span><span class="sxs-lookup"><span data-stu-id="b89d7-844">

<dt>Type</dt>

</span></span><dd><span data-ttu-id="b89d7-845">Array. < Zeichenfolge ></span><span class="sxs-lookup"><span data-stu-id="b89d7-845">Array.< String ></span></span></dd>

</dl>

##### <a name="example"></a><span data-ttu-id="b89d7-846">Beispiel</span><span class="sxs-lookup"><span data-stu-id="b89d7-846">Example</span></span>

```JavaScript
var fruits = Office.context.mailbox.item.getRegExMatchesByName("fruits");
var veggies = Office.context.mailbox.item.getRegExMatchesByName("veggies");
```

####  <a name="getselecteddataasynccoerciontype-options-callback--string"></a><span data-ttu-id="b89d7-847">getSelectedDataAsync(coercionType, [options], callback) → {String}</span><span class="sxs-lookup"><span data-stu-id="b89d7-847">getSelectedDataAsync(coercionType, [options], callback) → {String}</span></span>

<span data-ttu-id="b89d7-848">Gibt asynchron ausgewählte Daten aus dem Betreff oder Textkörper einer Nachricht zurück.</span><span class="sxs-lookup"><span data-stu-id="b89d7-848">Asynchronously returns selected data from the subject or body of a message.</span></span>

<span data-ttu-id="b89d7-p157">Wenn keine Auswahl vorhanden ist, aber der Cursor sich im Nachrichtentext oder Betreff befindet, gibt die Methode für die ausgewählten Daten NULL zurück. Wenn ein anderes Feld als der Textkörper oder Betreff ausgewählt ist, gibt die Methode den `InvalidSelection`-Fehler zurück.</span><span class="sxs-lookup"><span data-stu-id="b89d7-p157">If there is no selection but the cursor is in the body or subject, the method returns null for the selected data. If a field other than the body or subject is selected, the method returns the `InvalidSelection` error.</span></span>

##### <a name="parameters"></a><span data-ttu-id="b89d7-851">Parameter:</span><span class="sxs-lookup"><span data-stu-id="b89d7-851">Parameters:</span></span>

|<span data-ttu-id="b89d7-852">Name</span><span class="sxs-lookup"><span data-stu-id="b89d7-852">Name</span></span>| <span data-ttu-id="b89d7-853">Typ</span><span class="sxs-lookup"><span data-stu-id="b89d7-853">Type</span></span>| <span data-ttu-id="b89d7-854">Attribute</span><span class="sxs-lookup"><span data-stu-id="b89d7-854">Attributes</span></span>| <span data-ttu-id="b89d7-855">Beschreibung</span><span class="sxs-lookup"><span data-stu-id="b89d7-855">Description</span></span>|
|---|---|---|---|
|`coercionType`| [<span data-ttu-id="b89d7-856">Office.CoercionType</span><span class="sxs-lookup"><span data-stu-id="b89d7-856">Office.CoercionType</span></span>](office.md#coerciontype-string)||<span data-ttu-id="b89d7-p158">Fordert ein Format für die Daten an. Wenn es sich um Texthandelt, gibt die Methode den unformatierten Text als Zeichenfolge zurück und entfernt eventuell vorhandene HTML-Tags. Wenn es sich um HTML handelt, gibt die Methode den ausgewählten Text zurück, entweder als unformatierten Text oder als HTML.</span><span class="sxs-lookup"><span data-stu-id="b89d7-p158">Requests a format for the data. If Text, the method returns the plain text as a string , removing any HTML tags present. If HTML, the method returns the selected text, whether it is plaintext or HTML.</span></span>|
|`options`| <span data-ttu-id="b89d7-860">Object</span><span class="sxs-lookup"><span data-stu-id="b89d7-860">Object</span></span>| <span data-ttu-id="b89d7-861">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="b89d7-861">&lt;optional&gt;</span></span>|<span data-ttu-id="b89d7-862">Ein Objektliteral, das mindestens eine der folgenden Eigenschaften enthält.</span><span class="sxs-lookup"><span data-stu-id="b89d7-862">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`| <span data-ttu-id="b89d7-863">Object</span><span class="sxs-lookup"><span data-stu-id="b89d7-863">Object</span></span>| <span data-ttu-id="b89d7-864">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="b89d7-864">&lt;optional&gt;</span></span>|<span data-ttu-id="b89d7-865">Entwickler können ein Objekt bereitstellen, auf das sie in der Callbackmethode zugreifen möchten.</span><span class="sxs-lookup"><span data-stu-id="b89d7-865">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`| <span data-ttu-id="b89d7-866">Funktion</span><span class="sxs-lookup"><span data-stu-id="b89d7-866">function</span></span>||<span data-ttu-id="b89d7-867">Wenn die Methode abgeschlossen wird, wird die Funktion , die im `callback`-Parameter übergeben wird, mit einem einzelnen Parameter (`asyncResult`) aufgerufen, der ein [`AsyncResult`](/javascript/api/office/office.asyncresult)-Objekt darstellt.</span><span class="sxs-lookup"><span data-stu-id="b89d7-867">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span><br/><br/><span data-ttu-id="b89d7-868">Rufen Sie für den Zugriff auf die ausgewählten Daten aus der Rückrufmethode `asyncResult.value.data` auf.</span><span class="sxs-lookup"><span data-stu-id="b89d7-868">To access the selected data from the callback method, call `asyncResult.value.data`.</span></span> <span data-ttu-id="b89d7-869">Rufen Sie die Source-Eigenschaft für den Zugriff, die die Auswahl stammen, `asyncResult.value.sourceProperty`, die entweder sein `body` oder `subject`.</span><span class="sxs-lookup"><span data-stu-id="b89d7-869">To access the source property that the selection comes from, call `asyncResult.value.sourceProperty`, which will be either `body` or `subject`.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="b89d7-870">Anforderungen</span><span class="sxs-lookup"><span data-stu-id="b89d7-870">Requirements</span></span>

|<span data-ttu-id="b89d7-871">Anforderung</span><span class="sxs-lookup"><span data-stu-id="b89d7-871">Requirement</span></span>| <span data-ttu-id="b89d7-872">Wert</span><span class="sxs-lookup"><span data-stu-id="b89d7-872">Value</span></span>|
|---|---|
|[<span data-ttu-id="b89d7-873">Mindestversion des Postfachanforderungssatzes</span><span class="sxs-lookup"><span data-stu-id="b89d7-873">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="b89d7-874">1.2</span><span class="sxs-lookup"><span data-stu-id="b89d7-874">1.2</span></span>|
|[<span data-ttu-id="b89d7-875">Mindestberechtigungsstufe</span><span class="sxs-lookup"><span data-stu-id="b89d7-875">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="b89d7-876">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="b89d7-876">ReadWriteItem</span></span>|
|[<span data-ttu-id="b89d7-877">Zutreffender Outlook-Modus</span><span class="sxs-lookup"><span data-stu-id="b89d7-877">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="b89d7-878">Verfassen</span><span class="sxs-lookup"><span data-stu-id="b89d7-878">Compose</span></span>|

##### <a name="returns"></a><span data-ttu-id="b89d7-879">Gibt zurück:</span><span class="sxs-lookup"><span data-stu-id="b89d7-879">Returns:</span></span>

<span data-ttu-id="b89d7-880">Die ausgewählten Daten als Zeichenfolge mit dem durch `coercionType` bestimmten Format.</span><span class="sxs-lookup"><span data-stu-id="b89d7-880">The selected data as a string with format determined by `coercionType`.</span></span>

<dl class="param-type"><span data-ttu-id="b89d7-881">

<dt>
Typ</dt>


</span><span class="sxs-lookup"><span data-stu-id="b89d7-881">

<dt>Type</dt>

</span></span><dd><span data-ttu-id="b89d7-882">Zeichenfolge</span><span class="sxs-lookup"><span data-stu-id="b89d7-882">String</span></span></dd>

</dl>

##### <a name="example"></a><span data-ttu-id="b89d7-883">Beispiel</span><span class="sxs-lookup"><span data-stu-id="b89d7-883">Example</span></span>

```JavaScript
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

####  <a name="loadcustompropertiesasynccallback-usercontext"></a><span data-ttu-id="b89d7-884">loadCustomPropertiesAsync(callback, [userContext])</span><span class="sxs-lookup"><span data-stu-id="b89d7-884">loadCustomPropertiesAsync(callback, [userContext])</span></span>

<span data-ttu-id="b89d7-885">Lädt asynchron benutzerdefinierte Eigenschaften für dieses Add-In für das ausgewählte Element.</span><span class="sxs-lookup"><span data-stu-id="b89d7-885">Asynchronously loads custom properties for this add-in on the selected item.</span></span>

<span data-ttu-id="b89d7-p160">Benutzerdefinierte Eigenschaften werden als Schlüssel-/Wert-Paare pro App und pro Element gespeichert. Diese Methode gibt ein `CustomProperties`-Objekt im Rückruf zurück, das Methoden für den Zugriff auf die benutzerdefinierten Eigenschaften für das aktuelle Element und das aktuelle Add-In bietet. Benutzerdefinierte Eigenschaften sind für das Element nicht verschlüsselt und sollten darum nicht als sicherer Speicher verwendet werden.</span><span class="sxs-lookup"><span data-stu-id="b89d7-p160">Custom properties are stored as key/value pairs on a per-app, per-item basis. This method returns a `CustomProperties` object in the callback, which provides methods to access the custom properties specific to the current item and the current add-in. Custom properties are not encrypted on the item, so this should not be used as secure storage.</span></span>

##### <a name="parameters"></a><span data-ttu-id="b89d7-889">Parameter:</span><span class="sxs-lookup"><span data-stu-id="b89d7-889">Parameters:</span></span>

|<span data-ttu-id="b89d7-890">Name</span><span class="sxs-lookup"><span data-stu-id="b89d7-890">Name</span></span>| <span data-ttu-id="b89d7-891">Typ</span><span class="sxs-lookup"><span data-stu-id="b89d7-891">Type</span></span>| <span data-ttu-id="b89d7-892">Attribute</span><span class="sxs-lookup"><span data-stu-id="b89d7-892">Attributes</span></span>| <span data-ttu-id="b89d7-893">Beschreibung</span><span class="sxs-lookup"><span data-stu-id="b89d7-893">Description</span></span>|
|---|---|---|---|
|`callback`| <span data-ttu-id="b89d7-894">Funktion</span><span class="sxs-lookup"><span data-stu-id="b89d7-894">function</span></span>||<span data-ttu-id="b89d7-895">Wenn die Methode abgeschlossen wird, wird die Funktion , die im `callback`-Parameter übergeben wird, mit einem einzelnen Parameter (`asyncResult`) aufgerufen, der ein [`AsyncResult`](/javascript/api/office/office.asyncresult)-Objekt darstellt.</span><span class="sxs-lookup"><span data-stu-id="b89d7-895">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span><br/><br/><span data-ttu-id="b89d7-896">Die benutzerdefinierten Eigenschaften werden als [`CustomProperties`](/javascript/api/outlook_1_2/office.customproperties)-Objekt in der `asyncResult.value`-Eigenschaft bereitgestellt.</span><span class="sxs-lookup"><span data-stu-id="b89d7-896">The custom properties are provided as a [`CustomProperties`](/javascript/api/outlook_1_2/office.customproperties) object in the `asyncResult.value` property.</span></span> <span data-ttu-id="b89d7-897">Dieses Objekt kann zum Abrufen, festlegen und Entfernen benutzerdefinierter Eigenschaften aus dem Element und speichern Sie Änderungen an den benutzerdefinierten Eigenschaftensatz zurück an den Server verwendet werden.</span><span class="sxs-lookup"><span data-stu-id="b89d7-897">This object can be used to get, set, and remove custom properties from the item and save changes to the custom property set back to the server.</span></span>|
|`userContext`| <span data-ttu-id="b89d7-898">Objekt</span><span class="sxs-lookup"><span data-stu-id="b89d7-898">Object</span></span>| <span data-ttu-id="b89d7-899">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="b89d7-899">&lt;optional&gt;</span></span>|<span data-ttu-id="b89d7-900">Entwickler können ein beliebiges Objekt bereitstellen, den sie in der Rückruffunktion zugreifen möchten.</span><span class="sxs-lookup"><span data-stu-id="b89d7-900">Developers can provide any object they wish to access in the callback function.</span></span> <span data-ttu-id="b89d7-901">Dieses Objekt zugegriffen werden kann, indem die `asyncResult.asyncContext` -Eigenschaft in der Callback-Funktion.</span><span class="sxs-lookup"><span data-stu-id="b89d7-901">This object can be accessed by the `asyncResult.asyncContext` property in the callback function.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="b89d7-902">Anforderungen</span><span class="sxs-lookup"><span data-stu-id="b89d7-902">Requirements</span></span>

|<span data-ttu-id="b89d7-903">Anforderung</span><span class="sxs-lookup"><span data-stu-id="b89d7-903">Requirement</span></span>| <span data-ttu-id="b89d7-904">Wert</span><span class="sxs-lookup"><span data-stu-id="b89d7-904">Value</span></span>|
|---|---|
|[<span data-ttu-id="b89d7-905">Mindestversion des Postfachanforderungssatzes</span><span class="sxs-lookup"><span data-stu-id="b89d7-905">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="b89d7-906">1.0</span><span class="sxs-lookup"><span data-stu-id="b89d7-906">1.0</span></span>|
|[<span data-ttu-id="b89d7-907">Mindestberechtigungsstufe</span><span class="sxs-lookup"><span data-stu-id="b89d7-907">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="b89d7-908">ReadItem</span><span class="sxs-lookup"><span data-stu-id="b89d7-908">ReadItem</span></span>|
|[<span data-ttu-id="b89d7-909">Zutreffender Outlook-Modus</span><span class="sxs-lookup"><span data-stu-id="b89d7-909">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="b89d7-910">Verfassen oder Lesen</span><span class="sxs-lookup"><span data-stu-id="b89d7-910">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="b89d7-911">Beispiel</span><span class="sxs-lookup"><span data-stu-id="b89d7-911">Example</span></span>

<span data-ttu-id="b89d7-p163">Das folgende Codebeispiel veranschaulicht die Verwendung der `loadCustomPropertiesAsync`-Methode zum asynchronen Laden der benutzerdefinierten Eigenschaften, die für das aktuelle Element spezifisch sind. Des Weiteren veranschaulicht das Beispiel die Verwendung der `CustomProperties.saveAsync`-Methode zum Speichern der Eigenschaften auf dem Server. Nach dem Laden der benutzerdefinierten Eigenschaften wird in dem Codebeispiel die `CustomProperties.get`-Methode dazu verwendet, die benutzerdefinierte `myProp`-Eigenschaft zu lesen. Die `CustomProperties.set`-Methode wird dazu verwendet, die benutzerdefinierte `otherProp`-Eigenschaft zu schreiben. Schließlich wird die `saveAsync`-Methode aufgerufen, um die benutzerdefinierten Eigenschaften zu speichern.</span><span class="sxs-lookup"><span data-stu-id="b89d7-p163">The following code example shows how to use the `loadCustomPropertiesAsync` method to asynchronously load custom properties that are specific to the current item. The example also shows how to use the `CustomProperties.saveAsync` method to save these properties back to the server. After loading the custom properties, the code sample uses the `CustomProperties.get` method to read the custom property `myProp`, the `CustomProperties.set` method to write the custom property `otherProp`, and then finally calls the `saveAsync` method to save the custom properties.</span></span>

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

####  <a name="removeattachmentasyncattachmentid-options-callback"></a><span data-ttu-id="b89d7-915">removeAttachmentAsync(attachmentId, [options], [callback])</span><span class="sxs-lookup"><span data-stu-id="b89d7-915">removeAttachmentAsync(attachmentId, [options], [callback])</span></span>

<span data-ttu-id="b89d7-916">Entfernt eine Anlage aus einer Nachricht oder einem Termin.</span><span class="sxs-lookup"><span data-stu-id="b89d7-916">Removes an attachment from a message or appointment.</span></span>

<span data-ttu-id="b89d7-p164">Die `removeAttachmentAsync`-Methode entfernt die Anlage mit dem angegebenen Bezeichner aus dem Element. Als bewährte Vorgehensweise sollten Sie den Anlagenbezeichner nur dann zum Entfernen einer Anlage verwenden, wenn die gleiche Mail-App die Anlage in der gleichen Sitzung hinzugefügt hat. In Outlook Web App und OWA für Geräte ist der Anlagenbezeichner nur innerhalb der gleichen Sitzung gültig. Eine Sitzung ist abgeschlossen, wenn der Benutzer die App schließt, oder wenn der Benutzer in einem eingebetteten Formular mit dem Verfassen beginnt und den Vorgang dann in einem separaten Fenster fortsetzt.</span><span class="sxs-lookup"><span data-stu-id="b89d7-p164">The `removeAttachmentAsync` method removes the attachment with the specified identifier from the item. As a best practice, you should use the attachment identifier to remove an attachment only if the same mail app has added that attachment in the same session. In Outlook Web App and OWA for Devices, the attachment identifier is valid only within the same session. A session is over when the user closes the app, or if the user starts composing in an inline form and subsequently pops out the inline form to continue in a separate window.</span></span>

##### <a name="parameters"></a><span data-ttu-id="b89d7-921">Parameter:</span><span class="sxs-lookup"><span data-stu-id="b89d7-921">Parameters:</span></span>

|<span data-ttu-id="b89d7-922">Name</span><span class="sxs-lookup"><span data-stu-id="b89d7-922">Name</span></span>| <span data-ttu-id="b89d7-923">Typ</span><span class="sxs-lookup"><span data-stu-id="b89d7-923">Type</span></span>| <span data-ttu-id="b89d7-924">Attribute</span><span class="sxs-lookup"><span data-stu-id="b89d7-924">Attributes</span></span>| <span data-ttu-id="b89d7-925">Beschreibung</span><span class="sxs-lookup"><span data-stu-id="b89d7-925">Description</span></span>|
|---|---|---|---|
|`attachmentId`| <span data-ttu-id="b89d7-926">Zeichenfolge</span><span class="sxs-lookup"><span data-stu-id="b89d7-926">String</span></span>||<span data-ttu-id="b89d7-p165">Der Bezeichner des Anhangs, der entfernt werden soll. Die maximale Länge der Zeichenfolge ist 100 Zeichen.</span><span class="sxs-lookup"><span data-stu-id="b89d7-p165">The identifier of the attachment to remove. The maximum length of the string is 100 characters.</span></span>|
|`options`| <span data-ttu-id="b89d7-929">Object</span><span class="sxs-lookup"><span data-stu-id="b89d7-929">Object</span></span>| <span data-ttu-id="b89d7-930">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="b89d7-930">&lt;optional&gt;</span></span>|<span data-ttu-id="b89d7-931">Ein Objektliteral, das mindestens eine der folgenden Eigenschaften enthält.</span><span class="sxs-lookup"><span data-stu-id="b89d7-931">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`| <span data-ttu-id="b89d7-932">Object</span><span class="sxs-lookup"><span data-stu-id="b89d7-932">Object</span></span>| <span data-ttu-id="b89d7-933">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="b89d7-933">&lt;optional&gt;</span></span>|<span data-ttu-id="b89d7-934">Entwickler können ein Objekt bereitstellen, auf das sie in der Callbackmethode zugreifen möchten.</span><span class="sxs-lookup"><span data-stu-id="b89d7-934">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`| <span data-ttu-id="b89d7-935">Funktion</span><span class="sxs-lookup"><span data-stu-id="b89d7-935">function</span></span>| <span data-ttu-id="b89d7-936">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="b89d7-936">&lt;optional&gt;</span></span>|<span data-ttu-id="b89d7-937">Wenn die Methode abgeschlossen wird, wird die Funktion , die im `callback`-Parameter übergeben wird, mit einem einzelnen Parameter (`asyncResult`) aufgerufen, der ein [`AsyncResult`](/javascript/api/office/office.asyncresult)-Objekt darstellt.</span><span class="sxs-lookup"><span data-stu-id="b89d7-937">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span> <br/><span data-ttu-id="b89d7-938">Wenn beim Entfernen der Anlage ein Fehler auftritt, enthält die Eigenschaft `asyncResult.error` einen Fehlercode mit dem Grund für den Fehler.</span><span class="sxs-lookup"><span data-stu-id="b89d7-938">If removing the attachment fails, the `asyncResult.error` property will contain an error code with the reason for the failure.</span></span>|

##### <a name="errors"></a><span data-ttu-id="b89d7-939">Fehler</span><span class="sxs-lookup"><span data-stu-id="b89d7-939">Errors</span></span>

| <span data-ttu-id="b89d7-940">Fehlercode</span><span class="sxs-lookup"><span data-stu-id="b89d7-940">Error code</span></span> | <span data-ttu-id="b89d7-941">Beschreibung</span><span class="sxs-lookup"><span data-stu-id="b89d7-941">Description</span></span> |
|------------|-------------|
| `InvalidAttachmentId` | <span data-ttu-id="b89d7-942">Der Bezeichner für die Anlage ist nicht vorhanden.</span><span class="sxs-lookup"><span data-stu-id="b89d7-942">The attachment identifier does not exist.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="b89d7-943">Anforderungen</span><span class="sxs-lookup"><span data-stu-id="b89d7-943">Requirements</span></span>

|<span data-ttu-id="b89d7-944">Anforderung</span><span class="sxs-lookup"><span data-stu-id="b89d7-944">Requirement</span></span>| <span data-ttu-id="b89d7-945">Wert</span><span class="sxs-lookup"><span data-stu-id="b89d7-945">Value</span></span>|
|---|---|
|[<span data-ttu-id="b89d7-946">Mindestversion des Postfachanforderungssatzes</span><span class="sxs-lookup"><span data-stu-id="b89d7-946">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="b89d7-947">1.1</span><span class="sxs-lookup"><span data-stu-id="b89d7-947">1.1</span></span>|
|[<span data-ttu-id="b89d7-948">Mindestberechtigungsstufe</span><span class="sxs-lookup"><span data-stu-id="b89d7-948">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="b89d7-949">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="b89d7-949">ReadWriteItem</span></span>|
|[<span data-ttu-id="b89d7-950">Zutreffender Outlook-Modus</span><span class="sxs-lookup"><span data-stu-id="b89d7-950">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="b89d7-951">Verfassen</span><span class="sxs-lookup"><span data-stu-id="b89d7-951">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="b89d7-952">Beispiel</span><span class="sxs-lookup"><span data-stu-id="b89d7-952">Example</span></span>

<span data-ttu-id="b89d7-953">Mit dem folgende Code wird eine Anlage mit dem Bezeichner "0" entfernt.</span><span class="sxs-lookup"><span data-stu-id="b89d7-953">The following code removes an attachment with an identifier of '0'.</span></span>

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

####  <a name="setselecteddataasyncdata-options-callback"></a><span data-ttu-id="b89d7-954">setSelectedDataAsync(data, [options], callback)</span><span class="sxs-lookup"><span data-stu-id="b89d7-954">setSelectedDataAsync(data, [options], callback)</span></span>

<span data-ttu-id="b89d7-955">Fügt asynchron Daten in den Textkörper oder Betreff einer Nachricht ein.</span><span class="sxs-lookup"><span data-stu-id="b89d7-955">Asynchronously inserts data into the body or subject of a message.</span></span>

<span data-ttu-id="b89d7-p166">Die `setSelectedDataAsync`-Methode fügt die angegebene Zeichenfolge an der Cursorposition im Betreff oder im Textkörper ein, oder, falls im Editor Text ausgewählt ist, ersetzt den markierten Text. Wenn sich der Cursor nicht im Textkörper oder im Betreffsfeld befindet, wird ein Fehler zurückgegeben. Nach dem Einfügen wird der Cursor am Ende der eingefügten Inhalte platziert.</span><span class="sxs-lookup"><span data-stu-id="b89d7-p166">The `setSelectedDataAsync` method inserts the specified string at the cursor location in the subject or body of the item, or, if text is selected in the editor, it replaces the selected text. If the cursor is not in the body or subject field, an error is returned. After insertion, the cursor is placed at the end of the inserted content.</span></span>

##### <a name="parameters"></a><span data-ttu-id="b89d7-959">Parameter:</span><span class="sxs-lookup"><span data-stu-id="b89d7-959">Parameters:</span></span>

|<span data-ttu-id="b89d7-960">Name</span><span class="sxs-lookup"><span data-stu-id="b89d7-960">Name</span></span>| <span data-ttu-id="b89d7-961">Typ</span><span class="sxs-lookup"><span data-stu-id="b89d7-961">Type</span></span>| <span data-ttu-id="b89d7-962">Attribute</span><span class="sxs-lookup"><span data-stu-id="b89d7-962">Attributes</span></span>| <span data-ttu-id="b89d7-963">Beschreibung</span><span class="sxs-lookup"><span data-stu-id="b89d7-963">Description</span></span>|
|---|---|---|---|
|`data`| <span data-ttu-id="b89d7-964">Zeichenfolge</span><span class="sxs-lookup"><span data-stu-id="b89d7-964">String</span></span>||<span data-ttu-id="b89d7-p167">Die einzufügenden Daten. Daten dürfen 1.000.000 Zeichen nicht überschreiten. Werden mehr als 1.000.000 Zeichen übergeben, wird eine `ArgumentOutOfRange`-Ausnahme ausgelöst.</span><span class="sxs-lookup"><span data-stu-id="b89d7-p167">The data to be inserted. Data is not to exceed 1,000,000 characters. If more than 1,000,000 characters are passed in, an `ArgumentOutOfRange` exception is thrown.</span></span>|
|`options`| <span data-ttu-id="b89d7-968">Object</span><span class="sxs-lookup"><span data-stu-id="b89d7-968">Object</span></span>| <span data-ttu-id="b89d7-969">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="b89d7-969">&lt;optional&gt;</span></span>|<span data-ttu-id="b89d7-970">Ein Objektliteral, das mindestens eine der folgenden Eigenschaften enthält.</span><span class="sxs-lookup"><span data-stu-id="b89d7-970">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`| <span data-ttu-id="b89d7-971">Object</span><span class="sxs-lookup"><span data-stu-id="b89d7-971">Object</span></span>| <span data-ttu-id="b89d7-972">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="b89d7-972">&lt;optional&gt;</span></span>|<span data-ttu-id="b89d7-973">Entwickler können ein Objekt bereitstellen, auf das sie in der Rückrufmethode zugreifen möchten.</span><span class="sxs-lookup"><span data-stu-id="b89d7-973">Developers can provide any object they wish to access in the callback method.</span></span>|
|`options.coercionType`| [<span data-ttu-id="b89d7-974">Office.CoercionType</span><span class="sxs-lookup"><span data-stu-id="b89d7-974">Office.CoercionType</span></span>](office.md#coerciontype-string)| <span data-ttu-id="b89d7-975">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="b89d7-975">&lt;optional&gt;</span></span>|<span data-ttu-id="b89d7-p168">Wenn `text`, wird das aktuelle Format in Outlook Web App und Outlook angewendet. Wenn das Feld ein HTML-Editor ist, werden nur die Textdaten eingefügt, selbst wenn es sich bei den Daten um HTML-Daten handelt.</span><span class="sxs-lookup"><span data-stu-id="b89d7-p168">If `text`, the current style is applied in Outlook Web App and Outlook. If the field is an HTML editor, only the text data is inserted, even if the data is HTML.</span></span><br/><br/><span data-ttu-id="b89d7-p169">Wenn `html` gesetzt ist und das Feld HTML unterstützt (der Betreff jedoch nicht), wird die aktuelle Formatvorlage in Outlook Web App angewendet und die Standardformatvorlage in Outlook. Ist das Feld ein Textfeld, wird ein Fehler des Typs `InvalidDataFormat` zurückgegeben.</span><span class="sxs-lookup"><span data-stu-id="b89d7-p169">If `html` and the field supports HTML (the subject doesn't), the current style is applied in Outlook Web App and the default style is applied in Outlook. If the field is a text field, an `InvalidDataFormat` error is returned.</span></span><br/><br/><span data-ttu-id="b89d7-980">Wenn `coercionType` nicht festgelegt wird, hängt das Ergebnis vom Feld ab: Wenn das Feld HTML ist, wird HTML verwendet. Wenn das Feld Text ist, wird Nur-Text verwendet.</span><span class="sxs-lookup"><span data-stu-id="b89d7-980">If `coercionType` is not set, the result depends on the field: if the field is HTML then HTML is used; if the field is text, then plain text is used.</span></span>|
|`callback`| <span data-ttu-id="b89d7-981">function</span><span class="sxs-lookup"><span data-stu-id="b89d7-981">function</span></span>||<span data-ttu-id="b89d7-982">Wenn die Methode abgeschlossen wird, wird die Funktion , die im `callback`-Parameter übergeben wird, mit einem einzelnen Parameter (`asyncResult`) aufgerufen, der ein [`AsyncResult`](/javascript/api/office/office.asyncresult)-Objekt darstellt.</span><span class="sxs-lookup"><span data-stu-id="b89d7-982">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="b89d7-983">Anforderungen</span><span class="sxs-lookup"><span data-stu-id="b89d7-983">Requirements</span></span>

|<span data-ttu-id="b89d7-984">Anforderung</span><span class="sxs-lookup"><span data-stu-id="b89d7-984">Requirement</span></span>| <span data-ttu-id="b89d7-985">Wert</span><span class="sxs-lookup"><span data-stu-id="b89d7-985">Value</span></span>|
|---|---|
|[<span data-ttu-id="b89d7-986">Mindestversion des Postfachanforderungssatzes</span><span class="sxs-lookup"><span data-stu-id="b89d7-986">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="b89d7-987">1.2</span><span class="sxs-lookup"><span data-stu-id="b89d7-987">1.2</span></span>|
|[<span data-ttu-id="b89d7-988">Mindestberechtigungsstufe</span><span class="sxs-lookup"><span data-stu-id="b89d7-988">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="b89d7-989">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="b89d7-989">ReadWriteItem</span></span>|
|[<span data-ttu-id="b89d7-990">Zutreffender Outlook-Modus</span><span class="sxs-lookup"><span data-stu-id="b89d7-990">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="b89d7-991">Verfassen</span><span class="sxs-lookup"><span data-stu-id="b89d7-991">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="b89d7-992">Beispiel</span><span class="sxs-lookup"><span data-stu-id="b89d7-992">Example</span></span>

```JavaScript
Office.context.mailbox.item.setSelectedDataAsync("Hello World!");
Office.context.mailbox.item.setSelectedDataAsync("<b>Hello World!</b>", { coercionType : "html" });
```