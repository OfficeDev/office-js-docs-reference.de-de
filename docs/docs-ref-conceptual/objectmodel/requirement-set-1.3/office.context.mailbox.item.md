
# <a name="item"></a><span data-ttu-id="c91b5-101">item</span><span class="sxs-lookup"><span data-stu-id="c91b5-101">item</span></span>

### <a name="officeofficemdcontextofficecontextmdmailboxofficecontextmailboxmditem"></a><span data-ttu-id="c91b5-102">[Office](office.md)[.context](office.context.md)[.mailbox](office.context.mailbox.md).item</span><span class="sxs-lookup"><span data-stu-id="c91b5-102">[Office](office.md)[.context](office.context.md)[.mailbox](office.context.mailbox.md).item</span></span>

<span data-ttu-id="c91b5-p101">Der `item`-Namespace wird für den Zugriff auf die aktuell ausgewählte Nachricht, Besprechungsanfrage oder den aktuell ausgewählten Termin verwendet. Sie können den Typ von `item` mithilfe der [itemType](#itemtype-officemailboxenumsitemtypejavascriptapioutlook13officemailboxenumsitemtype)-Eigenschaft bestimmen.</span><span class="sxs-lookup"><span data-stu-id="c91b5-p101">The `item` namespace is used to access the currently selected message, meeting request, or appointment. You can determine the type of the `item` by using the [itemType](#itemtype-officemailboxenumsitemtypejavascriptapioutlook13officemailboxenumsitemtype) property.</span></span>

##### <a name="requirements"></a><span data-ttu-id="c91b5-105">Anforderungen</span><span class="sxs-lookup"><span data-stu-id="c91b5-105">Requirements</span></span>

|<span data-ttu-id="c91b5-106">Anforderung</span><span class="sxs-lookup"><span data-stu-id="c91b5-106">Requirement</span></span>| <span data-ttu-id="c91b5-107">Wert</span><span class="sxs-lookup"><span data-stu-id="c91b5-107">Value</span></span>|
|---|---|
|[<span data-ttu-id="c91b5-108">Mindestversion des Postfachanforderungssatzes</span><span class="sxs-lookup"><span data-stu-id="c91b5-108">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="c91b5-109">1.0</span><span class="sxs-lookup"><span data-stu-id="c91b5-109">1.0</span></span>|
|[<span data-ttu-id="c91b5-110">Mindestberechtigungsstufe</span><span class="sxs-lookup"><span data-stu-id="c91b5-110">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="c91b5-111">Eingeschränkt</span><span class="sxs-lookup"><span data-stu-id="c91b5-111">Restricted</span></span>|
|[<span data-ttu-id="c91b5-112">Zutreffender Outlook-Modus</span><span class="sxs-lookup"><span data-stu-id="c91b5-112">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="c91b5-113">Verfassen oder Lesen</span><span class="sxs-lookup"><span data-stu-id="c91b5-113">Compose or read</span></span>|

### <a name="example"></a><span data-ttu-id="c91b5-114">Beispiel</span><span class="sxs-lookup"><span data-stu-id="c91b5-114">Example</span></span>

<span data-ttu-id="c91b5-115">Im folgenden JavaScript-Codebeispiel wird der Zugriff auf die `subject`-Eigenschaft des aktuellen Elements in Outlook veranschaulicht.</span><span class="sxs-lookup"><span data-stu-id="c91b5-115">The following JavaScript code example shows how to access the `subject` property of the current item in Outlook.</span></span>

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

### <a name="members"></a><span data-ttu-id="c91b5-116">Elemente</span><span class="sxs-lookup"><span data-stu-id="c91b5-116">Members</span></span>

#### <a name="attachments-arrayattachmentdetailsjavascriptapioutlook13officeattachmentdetails"></a><span data-ttu-id="c91b5-117">attachments :Array.<[AttachmentDetails](/javascript/api/outlook_1_3/office.attachmentdetails)></span><span class="sxs-lookup"><span data-stu-id="c91b5-117">attachments :Array.<[AttachmentDetails](/javascript/api/outlook_1_3/office.attachmentdetails)></span></span>

<span data-ttu-id="c91b5-p102">Ruft ein Array mit Anlagen für das Element ab. Nur Lesemodus.</span><span class="sxs-lookup"><span data-stu-id="c91b5-p102">Gets an array of attachments for the item. Read mode only.</span></span>

> [!NOTE]
> <span data-ttu-id="c91b5-120">Bestimmte Dateitypen werden aufgrund von Sicherheitsproblemen von Outlook blockiert und daher nicht zurückgegeben werden.</span><span class="sxs-lookup"><span data-stu-id="c91b5-120">Certain types of files are blocked by Outlook due to potential security issues and are therefore not returned.</span></span> <span data-ttu-id="c91b5-121">Weitere Informationen finden Sie unter [Blockierte Anlagen in Outlook](https://support.office.com/article/Blocked-attachments-in-Outlook-434752E1-02D3-4E90-9124-8B81E49A8519).</span><span class="sxs-lookup"><span data-stu-id="c91b5-121">For more information, see [Blocked attachments in Outlook](https://support.office.com/article/Blocked-attachments-in-Outlook-434752E1-02D3-4E90-9124-8B81E49A8519).</span></span>

##### <a name="type"></a><span data-ttu-id="c91b5-122">Typ:</span><span class="sxs-lookup"><span data-stu-id="c91b5-122">Type:</span></span>

*   <span data-ttu-id="c91b5-123">Array.<[AttachmentDetails](/javascript/api/outlook_1_3/office.attachmentdetails)></span><span class="sxs-lookup"><span data-stu-id="c91b5-123">Array.<[AttachmentDetails](/javascript/api/outlook_1_3/office.attachmentdetails)></span></span>

##### <a name="requirements"></a><span data-ttu-id="c91b5-124">Anforderungen</span><span class="sxs-lookup"><span data-stu-id="c91b5-124">Requirements</span></span>

|<span data-ttu-id="c91b5-125">Anforderung</span><span class="sxs-lookup"><span data-stu-id="c91b5-125">Requirement</span></span>| <span data-ttu-id="c91b5-126">Wert</span><span class="sxs-lookup"><span data-stu-id="c91b5-126">Value</span></span>|
|---|---|
|[<span data-ttu-id="c91b5-127">Mindestversion des Postfachanforderungssatzes</span><span class="sxs-lookup"><span data-stu-id="c91b5-127">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="c91b5-128">1.0</span><span class="sxs-lookup"><span data-stu-id="c91b5-128">1.0</span></span>|
|[<span data-ttu-id="c91b5-129">Mindestberechtigungsstufe</span><span class="sxs-lookup"><span data-stu-id="c91b5-129">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="c91b5-130">ReadItem</span><span class="sxs-lookup"><span data-stu-id="c91b5-130">ReadItem</span></span>|
|[<span data-ttu-id="c91b5-131">Zutreffender Outlook-Modus</span><span class="sxs-lookup"><span data-stu-id="c91b5-131">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="c91b5-132">Lesen</span><span class="sxs-lookup"><span data-stu-id="c91b5-132">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="c91b5-133">Beispiel</span><span class="sxs-lookup"><span data-stu-id="c91b5-133">Example</span></span>

<span data-ttu-id="c91b5-134">Mit dem folgende Code wird eine HTML-Zeichenfolge mit Details aller Anlagen für das aktuelle Element erstellt.</span><span class="sxs-lookup"><span data-stu-id="c91b5-134">The following code builds an HTML string with details of all attachments on the current item.</span></span>

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

####  <a name="bcc-recipientsjavascriptapioutlook13officerecipients"></a><span data-ttu-id="c91b5-135">bcc :[Recipients](/javascript/api/outlook_1_3/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="c91b5-135">bcc :[Recipients](/javascript/api/outlook_1_3/office.recipients)</span></span>

<span data-ttu-id="c91b5-136">Ruft ein Objekt, das Methoden zum Abrufen oder aktualisieren die Empfänger in der Zeile Bcc (blind Carbon Copy, Blindkopie) einer Nachricht bereitstellt.</span><span class="sxs-lookup"><span data-stu-id="c91b5-136">Gets an object that provides methods to get or update the recipients on the Bcc (blind carbon copy) line of a message.</span></span> <span data-ttu-id="c91b5-137">Nur Verfassenmodus.</span><span class="sxs-lookup"><span data-stu-id="c91b5-137">Compose mode only.</span></span>

##### <a name="type"></a><span data-ttu-id="c91b5-138">Typ:</span><span class="sxs-lookup"><span data-stu-id="c91b5-138">Type:</span></span>

*   [<span data-ttu-id="c91b5-139">Recipients</span><span class="sxs-lookup"><span data-stu-id="c91b5-139">Recipients</span></span>](/javascript/api/outlook_1_3/office.recipients)

##### <a name="requirements"></a><span data-ttu-id="c91b5-140">Anforderungen</span><span class="sxs-lookup"><span data-stu-id="c91b5-140">Requirements</span></span>

|<span data-ttu-id="c91b5-141">Anforderung</span><span class="sxs-lookup"><span data-stu-id="c91b5-141">Requirement</span></span>| <span data-ttu-id="c91b5-142">Wert</span><span class="sxs-lookup"><span data-stu-id="c91b5-142">Value</span></span>|
|---|---|
|[<span data-ttu-id="c91b5-143">Mindestversion des Postfachanforderungssatzes</span><span class="sxs-lookup"><span data-stu-id="c91b5-143">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="c91b5-144">1.1</span><span class="sxs-lookup"><span data-stu-id="c91b5-144">1.1</span></span>|
|[<span data-ttu-id="c91b5-145">Mindestberechtigungsstufe</span><span class="sxs-lookup"><span data-stu-id="c91b5-145">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="c91b5-146">ReadItem</span><span class="sxs-lookup"><span data-stu-id="c91b5-146">ReadItem</span></span>|
|[<span data-ttu-id="c91b5-147">Zutreffender Outlook-Modus</span><span class="sxs-lookup"><span data-stu-id="c91b5-147">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="c91b5-148">Verfassen</span><span class="sxs-lookup"><span data-stu-id="c91b5-148">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="c91b5-149">Beispiel</span><span class="sxs-lookup"><span data-stu-id="c91b5-149">Example</span></span>

```
Office.context.mailbox.item.bcc.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.bcc.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.bcc.getAsync(callback);

function callback(asyncResult) {
  var arrayOfBccRecipients = asyncResult.value;
}
```

####  <a name="body-bodyjavascriptapioutlook13officebody"></a><span data-ttu-id="c91b5-150">body :[Body](/javascript/api/outlook_1_3/office.body)</span><span class="sxs-lookup"><span data-stu-id="c91b5-150">body :[Body](/javascript/api/outlook_1_3/office.body)</span></span>

<span data-ttu-id="c91b5-151">Ruft ein Objekt ab, das Methoden zum Bearbeiten des Textkörpers eines Elements bereitstellt.</span><span class="sxs-lookup"><span data-stu-id="c91b5-151">Gets an object that provides methods for manipulating the body of an item.</span></span>

##### <a name="type"></a><span data-ttu-id="c91b5-152">Typ:</span><span class="sxs-lookup"><span data-stu-id="c91b5-152">Type:</span></span>

*   [<span data-ttu-id="c91b5-153">Body</span><span class="sxs-lookup"><span data-stu-id="c91b5-153">Body</span></span>](/javascript/api/outlook_1_3/office.body)

##### <a name="requirements"></a><span data-ttu-id="c91b5-154">Anforderungen</span><span class="sxs-lookup"><span data-stu-id="c91b5-154">Requirements</span></span>

|<span data-ttu-id="c91b5-155">Anforderung</span><span class="sxs-lookup"><span data-stu-id="c91b5-155">Requirement</span></span>| <span data-ttu-id="c91b5-156">Wert</span><span class="sxs-lookup"><span data-stu-id="c91b5-156">Value</span></span>|
|---|---|
|[<span data-ttu-id="c91b5-157">Mindestversion des Postfachanforderungssatzes</span><span class="sxs-lookup"><span data-stu-id="c91b5-157">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="c91b5-158">1.1</span><span class="sxs-lookup"><span data-stu-id="c91b5-158">1.1</span></span>|
|[<span data-ttu-id="c91b5-159">Mindestberechtigungsstufe</span><span class="sxs-lookup"><span data-stu-id="c91b5-159">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="c91b5-160">ReadItem</span><span class="sxs-lookup"><span data-stu-id="c91b5-160">ReadItem</span></span>|
|[<span data-ttu-id="c91b5-161">Zutreffender Outlook-Modus</span><span class="sxs-lookup"><span data-stu-id="c91b5-161">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="c91b5-162">Verfassen oder Lesen</span><span class="sxs-lookup"><span data-stu-id="c91b5-162">Compose or read</span></span>|

####  <a name="cc-arrayemailaddressdetailsjavascriptapioutlook13officeemailaddressdetailsrecipientsjavascriptapioutlook13officerecipients"></a><span data-ttu-id="c91b5-163">cc: Array. <[EmailAddressDetails](/javascript/api/outlook_1_3/office.emailaddressdetails)>|[Empfänger](/javascript/api/outlook_1_3/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="c91b5-163">cc :Array.<[EmailAddressDetails](/javascript/api/outlook_1_3/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook_1_3/office.recipients)</span></span>

<span data-ttu-id="c91b5-164">Ermöglicht den Zugriff auf die (Carbon Copy, Blindkopie) Cc-Empfänger einer Nachricht.</span><span class="sxs-lookup"><span data-stu-id="c91b5-164">Provides access to the Cc (carbon copy) recipients of a message.</span></span> <span data-ttu-id="c91b5-165">Der Typ des Objekts und die Zugriffsebene, hängt von den Modus des aktuellen Elements ab.</span><span class="sxs-lookup"><span data-stu-id="c91b5-165">The type of object and level of access depends on the mode of the current item.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="c91b5-166">Lesemodus</span><span class="sxs-lookup"><span data-stu-id="c91b5-166">Read mode</span></span>

<span data-ttu-id="c91b5-p106">Die `cc`-Eigenschaft gibt ein Array mit einem `EmailAddressDetails`-Objekt für jeden Empfänger in der **Cc**-Zeile der Nachricht zurück. Die Auflistung ist auf höchstens 100 Elemente beschränkt.</span><span class="sxs-lookup"><span data-stu-id="c91b5-p106">The `cc` property returns an array that contains an `EmailAddressDetails` object for each recipient listed on the **Cc** line of the message. The collection is limited to a maximum of 100 members.</span></span>

##### <a name="compose-mode"></a><span data-ttu-id="c91b5-169">Verfassenmodus</span><span class="sxs-lookup"><span data-stu-id="c91b5-169">Compose mode</span></span>

<span data-ttu-id="c91b5-170">Die `cc` -Eigenschaft gibt eine `Recipients` -Objekt, das Methoden zum Abrufen oder aktualisieren die Empfänger in der Zeile **Cc** der Nachricht bereitstellt.</span><span class="sxs-lookup"><span data-stu-id="c91b5-170">The `cc` property returns a `Recipients` object that provides methods to get or update the recipients on the **Cc** line of the message.</span></span>

##### <a name="type"></a><span data-ttu-id="c91b5-171">Typ:</span><span class="sxs-lookup"><span data-stu-id="c91b5-171">Type:</span></span>

*   <span data-ttu-id="c91b5-172">Array.<[EmailAddressDetails](/javascript/api/outlook_1_3/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook_1_3/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="c91b5-172">Array.<[EmailAddressDetails](/javascript/api/outlook_1_3/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook_1_3/office.recipients)</span></span>

##### <a name="requirements"></a><span data-ttu-id="c91b5-173">Anforderungen</span><span class="sxs-lookup"><span data-stu-id="c91b5-173">Requirements</span></span>

|<span data-ttu-id="c91b5-174">Anforderung</span><span class="sxs-lookup"><span data-stu-id="c91b5-174">Requirement</span></span>| <span data-ttu-id="c91b5-175">Wert</span><span class="sxs-lookup"><span data-stu-id="c91b5-175">Value</span></span>|
|---|---|
|[<span data-ttu-id="c91b5-176">Mindestversion des Postfachanforderungssatzes</span><span class="sxs-lookup"><span data-stu-id="c91b5-176">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="c91b5-177">1.0</span><span class="sxs-lookup"><span data-stu-id="c91b5-177">1.0</span></span>|
|[<span data-ttu-id="c91b5-178">Mindestberechtigungsstufe</span><span class="sxs-lookup"><span data-stu-id="c91b5-178">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="c91b5-179">ReadItem</span><span class="sxs-lookup"><span data-stu-id="c91b5-179">ReadItem</span></span>|
|[<span data-ttu-id="c91b5-180">Zutreffender Outlook-Modus</span><span class="sxs-lookup"><span data-stu-id="c91b5-180">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="c91b5-181">Verfassen oder Lesen</span><span class="sxs-lookup"><span data-stu-id="c91b5-181">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="c91b5-182">Beispiel</span><span class="sxs-lookup"><span data-stu-id="c91b5-182">Example</span></span>

```
Office.context.mailbox.item.cc.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.cc.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.cc.getAsync(callback);

function callback(asyncResult) {
  var arrayOfCcRecipients = asyncResult.value;
}
```

####  <a name="nullable-conversationid-string"></a><span data-ttu-id="c91b5-183">(nullable) conversationId :String</span><span class="sxs-lookup"><span data-stu-id="c91b5-183">(nullable) conversationId :String</span></span>

<span data-ttu-id="c91b5-184">Ruft einen Bezeichner für die E-Mail-Unterhaltung ab, in der eine bestimmte Nachricht enthalten ist.</span><span class="sxs-lookup"><span data-stu-id="c91b5-184">Gets an identifier for the email conversation that contains a particular message.</span></span>

<span data-ttu-id="c91b5-p107">Sie können für diese Eigenschaft eine ganze Zahl abrufen, wenn Ihre Mail-App in Formularen zum Lesen oder Antworten in Formularen zum Verfassen aktiviert wird. Wenn der Benutzer den Betreff der Antwortnachricht ändert, ändert sich beim Versenden die Konversations-ID für die entsprechende Nachricht, und der Wert, den Sie vorher bezogen haben, trifft nicht länger zu.</span><span class="sxs-lookup"><span data-stu-id="c91b5-p107">You can get an integer for this property if your mail app is activated in read forms or responses in compose forms. If subsequently the user changes the subject of the reply message, upon sending the reply, the conversation ID for that message will change and that value you obtained earlier will no longer apply.</span></span>

<span data-ttu-id="c91b5-p108">Sie erhalten in einem Formular zum Verfassen für diese Eigenschaft für ein neues Element null. Wenn der Benutzer einen Betreff festlegt und das Element speichert, gibt die `conversationId`-Eigenschaft einen Wert zurück.</span><span class="sxs-lookup"><span data-stu-id="c91b5-p108">You get null for this property for a new item in a compose form. If the user sets a subject and saves the item, the `conversationId` property will return a value.</span></span>

##### <a name="type"></a><span data-ttu-id="c91b5-189">Typ:</span><span class="sxs-lookup"><span data-stu-id="c91b5-189">Type:</span></span>

*   <span data-ttu-id="c91b5-190">String</span><span class="sxs-lookup"><span data-stu-id="c91b5-190">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="c91b5-191">Anforderungen</span><span class="sxs-lookup"><span data-stu-id="c91b5-191">Requirements</span></span>

|<span data-ttu-id="c91b5-192">Anforderung</span><span class="sxs-lookup"><span data-stu-id="c91b5-192">Requirement</span></span>| <span data-ttu-id="c91b5-193">Wert</span><span class="sxs-lookup"><span data-stu-id="c91b5-193">Value</span></span>|
|---|---|
|[<span data-ttu-id="c91b5-194">Mindestversion des Postfachanforderungssatzes</span><span class="sxs-lookup"><span data-stu-id="c91b5-194">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="c91b5-195">1.0</span><span class="sxs-lookup"><span data-stu-id="c91b5-195">1.0</span></span>|
|[<span data-ttu-id="c91b5-196">Mindestberechtigungsstufe</span><span class="sxs-lookup"><span data-stu-id="c91b5-196">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="c91b5-197">ReadItem</span><span class="sxs-lookup"><span data-stu-id="c91b5-197">ReadItem</span></span>|
|[<span data-ttu-id="c91b5-198">Zutreffender Outlook-Modus</span><span class="sxs-lookup"><span data-stu-id="c91b5-198">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="c91b5-199">Verfassen oder Lesen</span><span class="sxs-lookup"><span data-stu-id="c91b5-199">Compose or read</span></span>|

#### <a name="datetimecreated-date"></a><span data-ttu-id="c91b5-200">dateTimeCreated :Date</span><span class="sxs-lookup"><span data-stu-id="c91b5-200">dateTimeCreated :Date</span></span>

<span data-ttu-id="c91b5-p109">Ruft das Datum und die Uhrzeit der Erstellung eines Elements ab. Nur Lesemodus.</span><span class="sxs-lookup"><span data-stu-id="c91b5-p109">Gets the date and time that an item was created. Read mode only.</span></span>

##### <a name="type"></a><span data-ttu-id="c91b5-203">Typ:</span><span class="sxs-lookup"><span data-stu-id="c91b5-203">Type:</span></span>

*   <span data-ttu-id="c91b5-204">Datum</span><span class="sxs-lookup"><span data-stu-id="c91b5-204">Date</span></span>

##### <a name="requirements"></a><span data-ttu-id="c91b5-205">Anforderungen</span><span class="sxs-lookup"><span data-stu-id="c91b5-205">Requirements</span></span>

|<span data-ttu-id="c91b5-206">Anforderung</span><span class="sxs-lookup"><span data-stu-id="c91b5-206">Requirement</span></span>| <span data-ttu-id="c91b5-207">Wert</span><span class="sxs-lookup"><span data-stu-id="c91b5-207">Value</span></span>|
|---|---|
|[<span data-ttu-id="c91b5-208">Mindestversion des Postfachanforderungssatzes</span><span class="sxs-lookup"><span data-stu-id="c91b5-208">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="c91b5-209">1.0</span><span class="sxs-lookup"><span data-stu-id="c91b5-209">1.0</span></span>|
|[<span data-ttu-id="c91b5-210">Mindestberechtigungsstufe</span><span class="sxs-lookup"><span data-stu-id="c91b5-210">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="c91b5-211">ReadItem</span><span class="sxs-lookup"><span data-stu-id="c91b5-211">ReadItem</span></span>|
|[<span data-ttu-id="c91b5-212">Zutreffender Outlook-Modus</span><span class="sxs-lookup"><span data-stu-id="c91b5-212">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="c91b5-213">Lesen</span><span class="sxs-lookup"><span data-stu-id="c91b5-213">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="c91b5-214">Beispiel</span><span class="sxs-lookup"><span data-stu-id="c91b5-214">Example</span></span>

```
var created = Office.context.mailbox.item.dateTimeCreated;
```

#### <a name="datetimemodified-date"></a><span data-ttu-id="c91b5-215">dateTimeModified :Date</span><span class="sxs-lookup"><span data-stu-id="c91b5-215">dateTimeModified :Date</span></span>

<span data-ttu-id="c91b5-p110">Ruft das Datum und die Uhrzeit der letzten Änderung eines Elements ab. Nur Lesemodus.</span><span class="sxs-lookup"><span data-stu-id="c91b5-p110">Gets the date and time that an item was last modified. Read mode only.</span></span>

> [!NOTE]
> <span data-ttu-id="c91b5-218">Dieser Member wird in Outlook für iOS oder Outlook für Android nicht unterstützt.</span><span class="sxs-lookup"><span data-stu-id="c91b5-218">This member is not supported in Outlook for iOS or Outlook for Android.</span></span>

##### <a name="type"></a><span data-ttu-id="c91b5-219">Typ:</span><span class="sxs-lookup"><span data-stu-id="c91b5-219">Type:</span></span>

*   <span data-ttu-id="c91b5-220">Datum</span><span class="sxs-lookup"><span data-stu-id="c91b5-220">Date</span></span>

##### <a name="requirements"></a><span data-ttu-id="c91b5-221">Anforderungen</span><span class="sxs-lookup"><span data-stu-id="c91b5-221">Requirements</span></span>

|<span data-ttu-id="c91b5-222">Anforderung</span><span class="sxs-lookup"><span data-stu-id="c91b5-222">Requirement</span></span>| <span data-ttu-id="c91b5-223">Wert</span><span class="sxs-lookup"><span data-stu-id="c91b5-223">Value</span></span>|
|---|---|
|[<span data-ttu-id="c91b5-224">Mindestversion des Postfachanforderungssatzes</span><span class="sxs-lookup"><span data-stu-id="c91b5-224">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="c91b5-225">1.0</span><span class="sxs-lookup"><span data-stu-id="c91b5-225">1.0</span></span>|
|[<span data-ttu-id="c91b5-226">Mindestberechtigungsstufe</span><span class="sxs-lookup"><span data-stu-id="c91b5-226">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="c91b5-227">ReadItem</span><span class="sxs-lookup"><span data-stu-id="c91b5-227">ReadItem</span></span>|
|[<span data-ttu-id="c91b5-228">Zutreffender Outlook-Modus</span><span class="sxs-lookup"><span data-stu-id="c91b5-228">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="c91b5-229">Lesen</span><span class="sxs-lookup"><span data-stu-id="c91b5-229">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="c91b5-230">Beispiel</span><span class="sxs-lookup"><span data-stu-id="c91b5-230">Example</span></span>

```
var modified = Office.context.mailbox.item.dateTimeModified;
```

####  <a name="end-datetimejavascriptapioutlook13officetime"></a><span data-ttu-id="c91b5-231">end :Date|[Time](/javascript/api/outlook_1_3/office.time)</span><span class="sxs-lookup"><span data-stu-id="c91b5-231">end :Date|[Time](/javascript/api/outlook_1_3/office.time)</span></span>

<span data-ttu-id="c91b5-232">Ruft Datum und Zeit für das Ende des Termins ab oder legt diese fest.</span><span class="sxs-lookup"><span data-stu-id="c91b5-232">Gets or sets the date and time that the appointment is to end.</span></span>

<span data-ttu-id="c91b5-p111">Die `end`-Eigenschaft wird als Datums- und Uhrzeitwert in UTC ausgedrückt. Sie können die [`convertToLocalClientTime`](office.context.mailbox.md#converttolocalclienttimetimevalue--localclienttimejavascriptapioutlook13officelocalclienttime)-Methode verwenden, um den Endeigenschaftswert in den lokalen Uhrzeit- und Datumswert des Clients umzuwandeln.</span><span class="sxs-lookup"><span data-stu-id="c91b5-p111">The `end` property is expressed as a Coordinated Universal Time (UTC) date and time value. You can use the [`convertToLocalClientTime`](office.context.mailbox.md#converttolocalclienttimetimevalue--localclienttimejavascriptapioutlook13officelocalclienttime) method to convert the end property value to the client’s local date and time.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="c91b5-235">Lesemodus</span><span class="sxs-lookup"><span data-stu-id="c91b5-235">Read mode</span></span>

<span data-ttu-id="c91b5-236">Die `end`-Eigenschaft gibt ein `Date`-Objekt zurück.</span><span class="sxs-lookup"><span data-stu-id="c91b5-236">The `end` property returns a `Date` object.</span></span>

##### <a name="compose-mode"></a><span data-ttu-id="c91b5-237">Verfassenmodus</span><span class="sxs-lookup"><span data-stu-id="c91b5-237">Compose mode</span></span>

<span data-ttu-id="c91b5-238">Die `end`-Eigenschaft gibt ein `Time`-Objekt zurück.</span><span class="sxs-lookup"><span data-stu-id="c91b5-238">The `end` property returns a `Time` object.</span></span>

<span data-ttu-id="c91b5-239">Wenn Sie die [`Time.setAsync`](/javascript/api/outlook_1_3/office.time#setasync-datetime--options--callback-)-Methode verwenden, um die Endzeit im Verfassenmodus festzulegen, sollten Sie die [`convertToUtcClientTime`](office.context.mailbox.md#converttoutcclienttimeinput--date)-Methode verwenden, um die Ortszeit auf dem Client für den Server in UTC umzuwandeln.</span><span class="sxs-lookup"><span data-stu-id="c91b5-239">When you use the [`Time.setAsync`](/javascript/api/outlook_1_3/office.time#setasync-datetime--options--callback-) method to set the end time, you should use the [`convertToUtcClientTime`](office.context.mailbox.md#converttoutcclienttimeinput--date) method to convert the local time on the client to UTC for the server.</span></span>

##### <a name="type"></a><span data-ttu-id="c91b5-240">Typ:</span><span class="sxs-lookup"><span data-stu-id="c91b5-240">Type:</span></span>

*   <span data-ttu-id="c91b5-241">Date | [Time](/javascript/api/outlook_1_3/office.time)</span><span class="sxs-lookup"><span data-stu-id="c91b5-241">Date | [Time](/javascript/api/outlook_1_3/office.time)</span></span>

##### <a name="requirements"></a><span data-ttu-id="c91b5-242">Anforderungen</span><span class="sxs-lookup"><span data-stu-id="c91b5-242">Requirements</span></span>

|<span data-ttu-id="c91b5-243">Anforderung</span><span class="sxs-lookup"><span data-stu-id="c91b5-243">Requirement</span></span>| <span data-ttu-id="c91b5-244">Wert</span><span class="sxs-lookup"><span data-stu-id="c91b5-244">Value</span></span>|
|---|---|
|[<span data-ttu-id="c91b5-245">Mindestversion des Postfachanforderungssatzes</span><span class="sxs-lookup"><span data-stu-id="c91b5-245">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="c91b5-246">1.0</span><span class="sxs-lookup"><span data-stu-id="c91b5-246">1.0</span></span>|
|[<span data-ttu-id="c91b5-247">Mindestberechtigungsstufe</span><span class="sxs-lookup"><span data-stu-id="c91b5-247">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="c91b5-248">ReadItem</span><span class="sxs-lookup"><span data-stu-id="c91b5-248">ReadItem</span></span>|
|[<span data-ttu-id="c91b5-249">Zutreffender Outlook-Modus</span><span class="sxs-lookup"><span data-stu-id="c91b5-249">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="c91b5-250">Verfassen oder Lesen</span><span class="sxs-lookup"><span data-stu-id="c91b5-250">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="c91b5-251">Beispiel</span><span class="sxs-lookup"><span data-stu-id="c91b5-251">Example</span></span>

<span data-ttu-id="c91b5-252">Im folgenden Beispiel wird die Endzeit eines Termins im Verfassenmodus mithilfe der [`setAsync`](/javascript/api/outlook_1_3/office.time#setasync-datetime--options--callback-)-Methode des `Time`-Objekts festgelegt.</span><span class="sxs-lookup"><span data-stu-id="c91b5-252">The following example sets the end time of an appointment in compose mode by using the [`setAsync`](/javascript/api/outlook_1_3/office.time#setasync-datetime--options--callback-) method of the `Time` object.</span></span>

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

#### <a name="from-emailaddressdetailsjavascriptapioutlook13officeemailaddressdetails"></a><span data-ttu-id="c91b5-253">from :[EmailAddressDetails](/javascript/api/outlook_1_3/office.emailaddressdetails)</span><span class="sxs-lookup"><span data-stu-id="c91b5-253">from :[EmailAddressDetails](/javascript/api/outlook_1_3/office.emailaddressdetails)</span></span>

<span data-ttu-id="c91b5-p112">Ruft die E-Mail-Adresse des Absenders einer Nachricht ab. Nur Lesemodus.</span><span class="sxs-lookup"><span data-stu-id="c91b5-p112">Gets the email address of the sender of a message. Read mode only.</span></span>

<span data-ttu-id="c91b5-p113">Die `from`- und [`sender`](#sender-emailaddressdetailsjavascriptapioutlook13officeemailaddressdetails)-Eigenschaften stellen dieselbe Person dar, außer die Nachricht wurde von einem Delegaten gesendet. In diesem Fall stellt die `from`-Eigenschaft die Stellvertretung dar, und die sender-Eigenschaft stellt den Delegaten dar.</span><span class="sxs-lookup"><span data-stu-id="c91b5-p113">The `from` and [`sender`](#sender-emailaddressdetailsjavascriptapioutlook13officeemailaddressdetails) properties represent the same person unless the message is sent by a delegate. In that case, the `from` property represents the delegator, and the sender property represents the delegate.</span></span>

> [!NOTE]
> <span data-ttu-id="c91b5-258">Die `recipientType` -Eigenschaft des der `EmailAddressDetails` -Objekts der `from` -Eigenschaft ist `undefined`.</span><span class="sxs-lookup"><span data-stu-id="c91b5-258">The `recipientType` property of the `EmailAddressDetails` object in the `from` property is `undefined`.</span></span>

##### <a name="type"></a><span data-ttu-id="c91b5-259">Typ:</span><span class="sxs-lookup"><span data-stu-id="c91b5-259">Type:</span></span>

*   [<span data-ttu-id="c91b5-260">EmailAddressDetails</span><span class="sxs-lookup"><span data-stu-id="c91b5-260">EmailAddressDetails</span></span>](/javascript/api/outlook_1_3/office.emailaddressdetails)

##### <a name="requirements"></a><span data-ttu-id="c91b5-261">Anforderungen</span><span class="sxs-lookup"><span data-stu-id="c91b5-261">Requirements</span></span>

|<span data-ttu-id="c91b5-262">Anforderung</span><span class="sxs-lookup"><span data-stu-id="c91b5-262">Requirement</span></span>| <span data-ttu-id="c91b5-263">Wert</span><span class="sxs-lookup"><span data-stu-id="c91b5-263">Value</span></span>|
|---|---|
|[<span data-ttu-id="c91b5-264">Mindestversion des Postfachanforderungssatzes</span><span class="sxs-lookup"><span data-stu-id="c91b5-264">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="c91b5-265">1.0</span><span class="sxs-lookup"><span data-stu-id="c91b5-265">1.0</span></span>|
|[<span data-ttu-id="c91b5-266">Mindestberechtigungsstufe</span><span class="sxs-lookup"><span data-stu-id="c91b5-266">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="c91b5-267">ReadItem</span><span class="sxs-lookup"><span data-stu-id="c91b5-267">ReadItem</span></span>|
|[<span data-ttu-id="c91b5-268">Zutreffender Outlook-Modus</span><span class="sxs-lookup"><span data-stu-id="c91b5-268">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="c91b5-269">Lesen</span><span class="sxs-lookup"><span data-stu-id="c91b5-269">Read</span></span>|

#### <a name="internetmessageid-string"></a><span data-ttu-id="c91b5-270">internetMessageId :String</span><span class="sxs-lookup"><span data-stu-id="c91b5-270">internetMessageId :String</span></span>

<span data-ttu-id="c91b5-p114">Ruft die Internetnachricht-ID einer E-Mail-Nachricht ab. Nur Lesemodus.</span><span class="sxs-lookup"><span data-stu-id="c91b5-p114">Gets the Internet message identifier for an email message. Read mode only.</span></span>

##### <a name="type"></a><span data-ttu-id="c91b5-273">Typ:</span><span class="sxs-lookup"><span data-stu-id="c91b5-273">Type:</span></span>

*   <span data-ttu-id="c91b5-274">String</span><span class="sxs-lookup"><span data-stu-id="c91b5-274">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="c91b5-275">Anforderungen</span><span class="sxs-lookup"><span data-stu-id="c91b5-275">Requirements</span></span>

|<span data-ttu-id="c91b5-276">Anforderung</span><span class="sxs-lookup"><span data-stu-id="c91b5-276">Requirement</span></span>| <span data-ttu-id="c91b5-277">Wert</span><span class="sxs-lookup"><span data-stu-id="c91b5-277">Value</span></span>|
|---|---|
|[<span data-ttu-id="c91b5-278">Mindestversion des Postfachanforderungssatzes</span><span class="sxs-lookup"><span data-stu-id="c91b5-278">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="c91b5-279">1.0</span><span class="sxs-lookup"><span data-stu-id="c91b5-279">1.0</span></span>|
|[<span data-ttu-id="c91b5-280">Mindestberechtigungsstufe</span><span class="sxs-lookup"><span data-stu-id="c91b5-280">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="c91b5-281">ReadItem</span><span class="sxs-lookup"><span data-stu-id="c91b5-281">ReadItem</span></span>|
|[<span data-ttu-id="c91b5-282">Zutreffender Outlook-Modus</span><span class="sxs-lookup"><span data-stu-id="c91b5-282">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="c91b5-283">Lesen</span><span class="sxs-lookup"><span data-stu-id="c91b5-283">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="c91b5-284">Beispiel</span><span class="sxs-lookup"><span data-stu-id="c91b5-284">Example</span></span>

```
var internetMessageId = Office.context.mailbox.item.internetMessageId;
```

#### <a name="itemclass-string"></a><span data-ttu-id="c91b5-285">itemClass :String</span><span class="sxs-lookup"><span data-stu-id="c91b5-285">itemClass :String</span></span>

<span data-ttu-id="c91b5-p115">Ruft die Exchange-Webdienste-Elementklasse des ausgewählten Elements ab. Nur Lesemodus.</span><span class="sxs-lookup"><span data-stu-id="c91b5-p115">Gets the Exchange Web Services item class of the selected item. Read mode only.</span></span>

<span data-ttu-id="c91b5-p116">Die `itemClass`-Eigenschaft gibt die Nachrichtenklasse des ausgewählten Elements an. Folgende Standardnachrichtenklassen für das Nachrichten- oder Terminelement sind vorhanden:</span><span class="sxs-lookup"><span data-stu-id="c91b5-p116">The `itemClass` property specifies the message class of the selected item. The following are the default message classes for the message or appointment item.</span></span>

| <span data-ttu-id="c91b5-290">Typ</span><span class="sxs-lookup"><span data-stu-id="c91b5-290">Type</span></span> | <span data-ttu-id="c91b5-291">Beschreibung</span><span class="sxs-lookup"><span data-stu-id="c91b5-291">Description</span></span> | <span data-ttu-id="c91b5-292">Elementklasse</span><span class="sxs-lookup"><span data-stu-id="c91b5-292">item class</span></span> |
| --- | --- | --- |
| <span data-ttu-id="c91b5-293">Terminelemente</span><span class="sxs-lookup"><span data-stu-id="c91b5-293">Appointment items</span></span> | <span data-ttu-id="c91b5-294">Dies sind Kalenderelemente der Elementklasse `IPM.Appointment` oder `IPM.Appointment.Occurence`.</span><span class="sxs-lookup"><span data-stu-id="c91b5-294">These are calendar items of the item class `IPM.Appointment` or `IPM.Appointment.Occurence`.</span></span> | `IPM.Appointment`<br />`IPM.Appointment.Occurence` |
| <span data-ttu-id="c91b5-295">Nachrichtenelemente</span><span class="sxs-lookup"><span data-stu-id="c91b5-295">Message items</span></span> | <span data-ttu-id="c91b5-296">Diese Elemente umfassen E-Mail-Nachrichten der Standardnachrichtenklasse `IPM.Note` sowie Besprechungsanfragen, -antworten und -absagen, die `IPM.Schedule.Meeting` als Basisnachrichtenklasse verwenden.</span><span class="sxs-lookup"><span data-stu-id="c91b5-296">These include email messages that have the default message class `IPM.Note`, and meeting requests, responses, and cancellations, that use `IPM.Schedule.Meeting` as the base message class.</span></span> | `IPM.Note`<br />`IPM.Schedule.Meeting.Request`<br />`IPM.Schedule.Meeting.Neg`<br />`IPM.Schedule.Meeting.Pos`<br />`IPM.Schedule.Meeting.Tent`<br />`IPM.Schedule.Meeting.Canceled` |

<span data-ttu-id="c91b5-297">Sie können benutzerdefinierte Nachrichtenklassen zum Erweitern einer Standardnachrichtenklasse erstellen, z. B. eine benutzerdefinierte Terminnachrichtenklasse wie `IPM.Appointment.Contoso`.</span><span class="sxs-lookup"><span data-stu-id="c91b5-297">You can create custom message classes that extends a default message class, for example, a custom appointment message class `IPM.Appointment.Contoso`.</span></span>

##### <a name="type"></a><span data-ttu-id="c91b5-298">Typ:</span><span class="sxs-lookup"><span data-stu-id="c91b5-298">Type:</span></span>

*   <span data-ttu-id="c91b5-299">String</span><span class="sxs-lookup"><span data-stu-id="c91b5-299">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="c91b5-300">Anforderungen</span><span class="sxs-lookup"><span data-stu-id="c91b5-300">Requirements</span></span>

|<span data-ttu-id="c91b5-301">Anforderung</span><span class="sxs-lookup"><span data-stu-id="c91b5-301">Requirement</span></span>| <span data-ttu-id="c91b5-302">Wert</span><span class="sxs-lookup"><span data-stu-id="c91b5-302">Value</span></span>|
|---|---|
|[<span data-ttu-id="c91b5-303">Mindestversion des Postfachanforderungssatzes</span><span class="sxs-lookup"><span data-stu-id="c91b5-303">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="c91b5-304">1.0</span><span class="sxs-lookup"><span data-stu-id="c91b5-304">1.0</span></span>|
|[<span data-ttu-id="c91b5-305">Mindestberechtigungsstufe</span><span class="sxs-lookup"><span data-stu-id="c91b5-305">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="c91b5-306">ReadItem</span><span class="sxs-lookup"><span data-stu-id="c91b5-306">ReadItem</span></span>|
|[<span data-ttu-id="c91b5-307">Zutreffender Outlook-Modus</span><span class="sxs-lookup"><span data-stu-id="c91b5-307">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="c91b5-308">Lesen</span><span class="sxs-lookup"><span data-stu-id="c91b5-308">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="c91b5-309">Beispiel</span><span class="sxs-lookup"><span data-stu-id="c91b5-309">Example</span></span>

```
var itemClass = Office.context.mailbox.item.itemClass;
```

#### <a name="nullable-itemid-string"></a><span data-ttu-id="c91b5-310">(nullable) itemId :String</span><span class="sxs-lookup"><span data-stu-id="c91b5-310">(nullable) itemId :String</span></span>

<span data-ttu-id="c91b5-p117">Ruft den Exchange-Webdienste-Elementbezeichner für das aktuelle Element ab. Nur Lesemodus.</span><span class="sxs-lookup"><span data-stu-id="c91b5-p117">Gets the Exchange Web Services item identifier for the current item. Read mode only.</span></span>

> [!NOTE]
> <span data-ttu-id="c91b5-313">Der Bezeichner, der von der `itemId`-Eigenschaft zurückgegeben wird, ist der gleiche Bezeichner wie der Exchange-Webdienste-Elementbezeichner.</span><span class="sxs-lookup"><span data-stu-id="c91b5-313">The identifier returned by the `itemId` property is the same as the Exchange Web Services item identifier.</span></span> <span data-ttu-id="c91b5-314">Die `itemId` -Eigenschaft ist nicht identisch mit der Outlook-Eintrags-ID oder die ID von der Outlook-REST-API verwendet.</span><span class="sxs-lookup"><span data-stu-id="c91b5-314">The `itemId` property is not identical to the Outlook Entry ID or the ID used by the Outlook REST API.</span></span> <span data-ttu-id="c91b5-315">REST-API-Aufrufe dieser Wert verwenden, sollten sie konvertiert werden, bevor [Office.context.mailbox.convertToRestId](office.context.mailbox.md#converttorestiditemid-restversion--string)verwenden.</span><span class="sxs-lookup"><span data-stu-id="c91b5-315">Before making REST API calls using this value, it should be converted using [Office.context.mailbox.convertToRestId](office.context.mailbox.md#converttorestiditemid-restversion--string).</span></span> <span data-ttu-id="c91b5-316">Weitere Informationen finden Sie unter [Verwenden der Outlook-REST-APIs aus einem Outlook-add-in](https://docs.microsoft.com/outlook/add-ins/use-rest-api#get-the-item-id).</span><span class="sxs-lookup"><span data-stu-id="c91b5-316">For more details, see [Use the Outlook REST APIs from an Outlook add-in](https://docs.microsoft.com/outlook/add-ins/use-rest-api#get-the-item-id).</span></span>

<span data-ttu-id="c91b5-p119">Die Eigenschaft `itemId` ist im Modus „Verfassen“ nicht verfügbar. Wenn ein Elementbezeichner erforderlich ist, kann die [`saveAsync`](#saveasyncoptions-callback)-Methode verwendet werden, um das Element zu speichern. Dadurch wird der Elementbezeichner im [`AsyncResult.value`](/javascript/api/office/office.asyncresult)-Parameter in der Rückruffunktion zurückgegeben.</span><span class="sxs-lookup"><span data-stu-id="c91b5-p119">The `itemId` property is not available in compose mode. If an item identifier is required, the [`saveAsync`](#saveasyncoptions-callback) method can be used to save the item to the store, which will return the item identifier in the [`AsyncResult.value`](/javascript/api/office/office.asyncresult) parameter in the callback function.</span></span>

##### <a name="type"></a><span data-ttu-id="c91b5-319">Typ:</span><span class="sxs-lookup"><span data-stu-id="c91b5-319">Type:</span></span>

*   <span data-ttu-id="c91b5-320">String</span><span class="sxs-lookup"><span data-stu-id="c91b5-320">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="c91b5-321">Anforderungen</span><span class="sxs-lookup"><span data-stu-id="c91b5-321">Requirements</span></span>

|<span data-ttu-id="c91b5-322">Anforderung</span><span class="sxs-lookup"><span data-stu-id="c91b5-322">Requirement</span></span>| <span data-ttu-id="c91b5-323">Wert</span><span class="sxs-lookup"><span data-stu-id="c91b5-323">Value</span></span>|
|---|---|
|[<span data-ttu-id="c91b5-324">Mindestversion des Postfachanforderungssatzes</span><span class="sxs-lookup"><span data-stu-id="c91b5-324">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="c91b5-325">1.0</span><span class="sxs-lookup"><span data-stu-id="c91b5-325">1.0</span></span>|
|[<span data-ttu-id="c91b5-326">Mindestberechtigungsstufe</span><span class="sxs-lookup"><span data-stu-id="c91b5-326">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="c91b5-327">ReadItem</span><span class="sxs-lookup"><span data-stu-id="c91b5-327">ReadItem</span></span>|
|[<span data-ttu-id="c91b5-328">Zutreffender Outlook-Modus</span><span class="sxs-lookup"><span data-stu-id="c91b5-328">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="c91b5-329">Lesen</span><span class="sxs-lookup"><span data-stu-id="c91b5-329">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="c91b5-330">Beispiel</span><span class="sxs-lookup"><span data-stu-id="c91b5-330">Example</span></span>

<span data-ttu-id="c91b5-p120">Mit dem folgende Code wird das Vorhandensein eines Elementbezeichners überprüft. Wenn die `itemId`-Eigenschaft `null` oder `undefined` zurückgibt, speichert sie das Element und ruft den Elementbezeichner aus dem asynchronen Ergebnis ab.</span><span class="sxs-lookup"><span data-stu-id="c91b5-p120">The following code checks for the presence of an item identifier. If the `itemId` property returns `null` or `undefined`, it saves the item to the store and gets the item identifier from the asynchronous result.</span></span>

```
var itemId = Office.context.mailbox.item.itemId;
if (itemId === null || itemId == undefined) {
  Office.context.mailbox.item.saveAsync(function(result){
    itemId = result.value;
  });
}
```

####  <a name="itemtype-officemailboxenumsitemtypejavascriptapioutlook13officemailboxenumsitemtype"></a><span data-ttu-id="c91b5-333">itemType :[Office.MailboxEnums.ItemType](/javascript/api/outlook_1_3/office.mailboxenums.itemtype)</span><span class="sxs-lookup"><span data-stu-id="c91b5-333">itemType :[Office.MailboxEnums.ItemType](/javascript/api/outlook_1_3/office.mailboxenums.itemtype)</span></span>

<span data-ttu-id="c91b5-334">Ruft den Typ des Elements ab, das eine Instanz darstellt.</span><span class="sxs-lookup"><span data-stu-id="c91b5-334">Gets the type of item that an instance represents.</span></span>

<span data-ttu-id="c91b5-335">Die `itemType`-Eigenschaft gibt einen der Werte der `ItemType`-Enumeration zurück, der angibt, ob es sich bei der `item`-Objektinstanz um eine Nachricht oder einen Termin handelt.</span><span class="sxs-lookup"><span data-stu-id="c91b5-335">The `itemType` property returns one of the `ItemType` enumeration values, indicating whether the `item` object instance is a message or an appointment.</span></span>

##### <a name="type"></a><span data-ttu-id="c91b5-336">Typ:</span><span class="sxs-lookup"><span data-stu-id="c91b5-336">Type:</span></span>

*   [<span data-ttu-id="c91b5-337">Office.MailboxEnums.ItemType</span><span class="sxs-lookup"><span data-stu-id="c91b5-337">Office.MailboxEnums.ItemType</span></span>](/javascript/api/outlook_1_3/office.mailboxenums.itemtype)

##### <a name="requirements"></a><span data-ttu-id="c91b5-338">Anforderungen</span><span class="sxs-lookup"><span data-stu-id="c91b5-338">Requirements</span></span>

|<span data-ttu-id="c91b5-339">Anforderung</span><span class="sxs-lookup"><span data-stu-id="c91b5-339">Requirement</span></span>| <span data-ttu-id="c91b5-340">Wert</span><span class="sxs-lookup"><span data-stu-id="c91b5-340">Value</span></span>|
|---|---|
|[<span data-ttu-id="c91b5-341">Mindestversion des Postfachanforderungssatzes</span><span class="sxs-lookup"><span data-stu-id="c91b5-341">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="c91b5-342">1.0</span><span class="sxs-lookup"><span data-stu-id="c91b5-342">1.0</span></span>|
|[<span data-ttu-id="c91b5-343">Mindestberechtigungsstufe</span><span class="sxs-lookup"><span data-stu-id="c91b5-343">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="c91b5-344">ReadItem</span><span class="sxs-lookup"><span data-stu-id="c91b5-344">ReadItem</span></span>|
|[<span data-ttu-id="c91b5-345">Zutreffender Outlook-Modus</span><span class="sxs-lookup"><span data-stu-id="c91b5-345">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="c91b5-346">Verfassen oder Lesen</span><span class="sxs-lookup"><span data-stu-id="c91b5-346">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="c91b5-347">Beispiel</span><span class="sxs-lookup"><span data-stu-id="c91b5-347">Example</span></span>

```
if (Office.context.mailbox.item.itemType == Office.MailboxEnums.ItemType.Message)
  // do something
else
  // do something else
```

####  <a name="location-stringlocationjavascriptapioutlook13officelocation"></a><span data-ttu-id="c91b5-348">location :String|[Location](/javascript/api/outlook_1_3/office.location)</span><span class="sxs-lookup"><span data-stu-id="c91b5-348">location :String|[Location](/javascript/api/outlook_1_3/office.location)</span></span>

<span data-ttu-id="c91b5-349">Ruft den Ort eines Termins ab bzw. legt ihn fest.</span><span class="sxs-lookup"><span data-stu-id="c91b5-349">Gets or sets the location of an appointment.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="c91b5-350">Lesemodus</span><span class="sxs-lookup"><span data-stu-id="c91b5-350">Read mode</span></span>

<span data-ttu-id="c91b5-351">Die `location`-Eigenschaft gibt eine Zeichenfolge zurück, die den Ort des Termins enthält.</span><span class="sxs-lookup"><span data-stu-id="c91b5-351">The `location` property returns a string that contains the location of the appointment.</span></span>

##### <a name="compose-mode"></a><span data-ttu-id="c91b5-352">Verfassenmodus</span><span class="sxs-lookup"><span data-stu-id="c91b5-352">Compose mode</span></span>

<span data-ttu-id="c91b5-353">Die `location`-Eigenschaft gibt ein `Location`-Objekt zurück, das Methoden zum Abrufen und Festlegen des Orts für einen Termin bereitstellt.</span><span class="sxs-lookup"><span data-stu-id="c91b5-353">The `location` property returns a `Location` object that provides methods that are used to get and set the location of the appointment.</span></span>

##### <a name="type"></a><span data-ttu-id="c91b5-354">Typ:</span><span class="sxs-lookup"><span data-stu-id="c91b5-354">Type:</span></span>

*   <span data-ttu-id="c91b5-355">String | [Location](/javascript/api/outlook_1_3/office.location)</span><span class="sxs-lookup"><span data-stu-id="c91b5-355">String | [Location](/javascript/api/outlook_1_3/office.location)</span></span>

##### <a name="requirements"></a><span data-ttu-id="c91b5-356">Anforderungen</span><span class="sxs-lookup"><span data-stu-id="c91b5-356">Requirements</span></span>

|<span data-ttu-id="c91b5-357">Anforderung</span><span class="sxs-lookup"><span data-stu-id="c91b5-357">Requirement</span></span>| <span data-ttu-id="c91b5-358">Wert</span><span class="sxs-lookup"><span data-stu-id="c91b5-358">Value</span></span>|
|---|---|
|[<span data-ttu-id="c91b5-359">Mindestversion des Postfachanforderungssatzes</span><span class="sxs-lookup"><span data-stu-id="c91b5-359">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="c91b5-360">1.0</span><span class="sxs-lookup"><span data-stu-id="c91b5-360">1.0</span></span>|
|[<span data-ttu-id="c91b5-361">Mindestberechtigungsstufe</span><span class="sxs-lookup"><span data-stu-id="c91b5-361">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="c91b5-362">ReadItem</span><span class="sxs-lookup"><span data-stu-id="c91b5-362">ReadItem</span></span>|
|[<span data-ttu-id="c91b5-363">Zutreffender Outlook-Modus</span><span class="sxs-lookup"><span data-stu-id="c91b5-363">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="c91b5-364">Verfassen oder Lesen</span><span class="sxs-lookup"><span data-stu-id="c91b5-364">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="c91b5-365">Beispiel</span><span class="sxs-lookup"><span data-stu-id="c91b5-365">Example</span></span>

```
var userContext = { value : 1 };
Office.context.mailbox.item.location.getAsync( { context: userContext}, callback);

function callback(asyncResult) {
  var context = asyncResult.context;
  var location = asyncResult.value;
}
```

#### <a name="normalizedsubject-string"></a><span data-ttu-id="c91b5-366">normalizedSubject :String</span><span class="sxs-lookup"><span data-stu-id="c91b5-366">normalizedSubject :String</span></span>

<span data-ttu-id="c91b5-p121">Ruft den Betreff für ein Element ab, wobei alle Präfixe entfernt werden (einschließlich `RE:` und `FWD:`). Nur Lesemodus.</span><span class="sxs-lookup"><span data-stu-id="c91b5-p121">Gets the subject of an item, with all prefixes removed (including `RE:` and `FWD:`). Read mode only.</span></span>

<span data-ttu-id="c91b5-p122">Die normalizedSubject-Eigenschaft ruft den Betreff des Elements mit allen Standardpräfixen (z. B. `RE:` und `FW:`) ab, die von E-Mail-Programmen hinzugefügt werden. Wenn der Betreff des Elements mit vollständigen Präfixen abgerufen werden soll, verwenden Sie die [`subject`](#subject-stringsubjectjavascriptapioutlook13officesubject)-Eigenschaft.</span><span class="sxs-lookup"><span data-stu-id="c91b5-p122">The normalizedSubject property gets the subject of the item, with any standard prefixes (such as `RE:` and `FW:`) that are added by email programs. To get the subject of the item with the prefixes intact, use the [`subject`](#subject-stringsubjectjavascriptapioutlook13officesubject) property.</span></span>

##### <a name="type"></a><span data-ttu-id="c91b5-371">Typ:</span><span class="sxs-lookup"><span data-stu-id="c91b5-371">Type:</span></span>

*   <span data-ttu-id="c91b5-372">String</span><span class="sxs-lookup"><span data-stu-id="c91b5-372">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="c91b5-373">Anforderungen</span><span class="sxs-lookup"><span data-stu-id="c91b5-373">Requirements</span></span>

|<span data-ttu-id="c91b5-374">Anforderung</span><span class="sxs-lookup"><span data-stu-id="c91b5-374">Requirement</span></span>| <span data-ttu-id="c91b5-375">Wert</span><span class="sxs-lookup"><span data-stu-id="c91b5-375">Value</span></span>|
|---|---|
|[<span data-ttu-id="c91b5-376">Mindestversion des Postfachanforderungssatzes</span><span class="sxs-lookup"><span data-stu-id="c91b5-376">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="c91b5-377">1.0</span><span class="sxs-lookup"><span data-stu-id="c91b5-377">1.0</span></span>|
|[<span data-ttu-id="c91b5-378">Mindestberechtigungsstufe</span><span class="sxs-lookup"><span data-stu-id="c91b5-378">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="c91b5-379">ReadItem</span><span class="sxs-lookup"><span data-stu-id="c91b5-379">ReadItem</span></span>|
|[<span data-ttu-id="c91b5-380">Zutreffender Outlook-Modus</span><span class="sxs-lookup"><span data-stu-id="c91b5-380">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="c91b5-381">Lesen</span><span class="sxs-lookup"><span data-stu-id="c91b5-381">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="c91b5-382">Beispiel</span><span class="sxs-lookup"><span data-stu-id="c91b5-382">Example</span></span>

```
var normalizedSubject = Office.context.mailbox.item.normalizedSubject;
```

####  <a name="notificationmessages-notificationmessagesjavascriptapioutlook13officenotificationmessages"></a><span data-ttu-id="c91b5-383">notificationMessages :[NotificationMessages](/javascript/api/outlook_1_3/office.notificationmessages)</span><span class="sxs-lookup"><span data-stu-id="c91b5-383">notificationMessages :[NotificationMessages](/javascript/api/outlook_1_3/office.notificationmessages)</span></span>

<span data-ttu-id="c91b5-384">Ruft die Benachrichtigungen für ein Element ab.</span><span class="sxs-lookup"><span data-stu-id="c91b5-384">Gets the notification messages for an item.</span></span>

##### <a name="type"></a><span data-ttu-id="c91b5-385">Typ:</span><span class="sxs-lookup"><span data-stu-id="c91b5-385">Type:</span></span>

*   [<span data-ttu-id="c91b5-386">NotificationMessages</span><span class="sxs-lookup"><span data-stu-id="c91b5-386">NotificationMessages</span></span>](/javascript/api/outlook_1_3/office.notificationmessages)

##### <a name="requirements"></a><span data-ttu-id="c91b5-387">Anforderungen</span><span class="sxs-lookup"><span data-stu-id="c91b5-387">Requirements</span></span>

|<span data-ttu-id="c91b5-388">Anforderung</span><span class="sxs-lookup"><span data-stu-id="c91b5-388">Requirement</span></span>| <span data-ttu-id="c91b5-389">Wert</span><span class="sxs-lookup"><span data-stu-id="c91b5-389">Value</span></span>|
|---|---|
|[<span data-ttu-id="c91b5-390">Mindestversion des Postfachanforderungssatzes</span><span class="sxs-lookup"><span data-stu-id="c91b5-390">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="c91b5-391">1.3</span><span class="sxs-lookup"><span data-stu-id="c91b5-391">1.3</span></span>|
|[<span data-ttu-id="c91b5-392">Mindestberechtigungsstufe</span><span class="sxs-lookup"><span data-stu-id="c91b5-392">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="c91b5-393">ReadItem</span><span class="sxs-lookup"><span data-stu-id="c91b5-393">ReadItem</span></span>|
|[<span data-ttu-id="c91b5-394">Zutreffender Outlook-Modus</span><span class="sxs-lookup"><span data-stu-id="c91b5-394">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="c91b5-395">Verfassen oder Lesen</span><span class="sxs-lookup"><span data-stu-id="c91b5-395">Compose or read</span></span>|

####  <a name="optionalattendees-arrayemailaddressdetailsjavascriptapioutlook13officeemailaddressdetailsrecipientsjavascriptapioutlook13officerecipients"></a><span data-ttu-id="c91b5-396">optionalAttendees :Array.<[EmailAddressDetails](/javascript/api/outlook_1_3/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook_1_3/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="c91b5-396">optionalAttendees :Array.<[EmailAddressDetails](/javascript/api/outlook_1_3/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook_1_3/office.recipients)</span></span>

<span data-ttu-id="c91b5-397">Ermöglicht den Zugriff auf die optionalen Teilnehmer eines Ereignisses.</span><span class="sxs-lookup"><span data-stu-id="c91b5-397">Provides access to the optional attendees of an event.</span></span> <span data-ttu-id="c91b5-398">Der Typ des Objekts und die Zugriffsebene, hängt von den Modus des aktuellen Elements ab.</span><span class="sxs-lookup"><span data-stu-id="c91b5-398">The type of object and level of access depends on the mode of the current item.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="c91b5-399">Lesemodus</span><span class="sxs-lookup"><span data-stu-id="c91b5-399">Read mode</span></span>

<span data-ttu-id="c91b5-400">Die `optionalAttendees`-Eigenschaft gibt ein Array mit einem `EmailAddressDetails`-Objekt für jeden optionalen Teilnehmer an der Besprechung zurück.</span><span class="sxs-lookup"><span data-stu-id="c91b5-400">The `optionalAttendees` property returns an array that contains an `EmailAddressDetails` object for each optional attendee to the meeting.</span></span>

##### <a name="compose-mode"></a><span data-ttu-id="c91b5-401">Verfassenmodus</span><span class="sxs-lookup"><span data-stu-id="c91b5-401">Compose mode</span></span>

<span data-ttu-id="c91b5-402">Die `optionalAttendees` -Eigenschaft gibt eine `Recipients` -Objekt, das Methoden zum Abrufen oder Aktualisieren der optionalen Teilnehmer für eine Besprechung bereitstellt.</span><span class="sxs-lookup"><span data-stu-id="c91b5-402">The `optionalAttendees` property returns a `Recipients` object that provides methods to get or update the optional attendees for a meeting.</span></span>

##### <a name="type"></a><span data-ttu-id="c91b5-403">Typ:</span><span class="sxs-lookup"><span data-stu-id="c91b5-403">Type:</span></span>

*   <span data-ttu-id="c91b5-404">Array.<[EmailAddressDetails](/javascript/api/outlook_1_3/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook_1_3/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="c91b5-404">Array.<[EmailAddressDetails](/javascript/api/outlook_1_3/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook_1_3/office.recipients)</span></span>

##### <a name="requirements"></a><span data-ttu-id="c91b5-405">Anforderungen</span><span class="sxs-lookup"><span data-stu-id="c91b5-405">Requirements</span></span>

|<span data-ttu-id="c91b5-406">Anforderung</span><span class="sxs-lookup"><span data-stu-id="c91b5-406">Requirement</span></span>| <span data-ttu-id="c91b5-407">Wert</span><span class="sxs-lookup"><span data-stu-id="c91b5-407">Value</span></span>|
|---|---|
|[<span data-ttu-id="c91b5-408">Mindestversion des Postfachanforderungssatzes</span><span class="sxs-lookup"><span data-stu-id="c91b5-408">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="c91b5-409">1.0</span><span class="sxs-lookup"><span data-stu-id="c91b5-409">1.0</span></span>|
|[<span data-ttu-id="c91b5-410">Mindestberechtigungsstufe</span><span class="sxs-lookup"><span data-stu-id="c91b5-410">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="c91b5-411">ReadItem</span><span class="sxs-lookup"><span data-stu-id="c91b5-411">ReadItem</span></span>|
|[<span data-ttu-id="c91b5-412">Zutreffender Outlook-Modus</span><span class="sxs-lookup"><span data-stu-id="c91b5-412">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="c91b5-413">Verfassen oder Lesen</span><span class="sxs-lookup"><span data-stu-id="c91b5-413">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="c91b5-414">Beispiel</span><span class="sxs-lookup"><span data-stu-id="c91b5-414">Example</span></span>

```
Office.context.mailbox.item.optionalAttendees.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.optionalAttendees.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.optionalAttendees.getAsync(callback);

function callback(asyncResult) {
  var arrayOfOptionalAttendeesRecipients = asyncResult.value;
}
```

#### <a name="organizer-emailaddressdetailsjavascriptapioutlook13officeemailaddressdetails"></a><span data-ttu-id="c91b5-415">organizer :[EmailAddressDetails](/javascript/api/outlook_1_3/office.emailaddressdetails)</span><span class="sxs-lookup"><span data-stu-id="c91b5-415">organizer :[EmailAddressDetails](/javascript/api/outlook_1_3/office.emailaddressdetails)</span></span>

<span data-ttu-id="c91b5-p124">Ruft für eine angegebene Besprechung die E-Mail-Adresse des Organisators der Besprechung ab. Nur Lesemodus.</span><span class="sxs-lookup"><span data-stu-id="c91b5-p124">Gets the email address of the meeting organizer for a specified meeting. Read mode only.</span></span>

##### <a name="type"></a><span data-ttu-id="c91b5-418">Typ:</span><span class="sxs-lookup"><span data-stu-id="c91b5-418">Type:</span></span>

*   [<span data-ttu-id="c91b5-419">EmailAddressDetails</span><span class="sxs-lookup"><span data-stu-id="c91b5-419">EmailAddressDetails</span></span>](/javascript/api/outlook_1_3/office.emailaddressdetails)

##### <a name="requirements"></a><span data-ttu-id="c91b5-420">Anforderungen</span><span class="sxs-lookup"><span data-stu-id="c91b5-420">Requirements</span></span>

|<span data-ttu-id="c91b5-421">Anforderung</span><span class="sxs-lookup"><span data-stu-id="c91b5-421">Requirement</span></span>| <span data-ttu-id="c91b5-422">Wert</span><span class="sxs-lookup"><span data-stu-id="c91b5-422">Value</span></span>|
|---|---|
|[<span data-ttu-id="c91b5-423">Mindestversion des Postfachanforderungssatzes</span><span class="sxs-lookup"><span data-stu-id="c91b5-423">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="c91b5-424">1.0</span><span class="sxs-lookup"><span data-stu-id="c91b5-424">1.0</span></span>|
|[<span data-ttu-id="c91b5-425">Mindestberechtigungsstufe</span><span class="sxs-lookup"><span data-stu-id="c91b5-425">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="c91b5-426">ReadItem</span><span class="sxs-lookup"><span data-stu-id="c91b5-426">ReadItem</span></span>|
|[<span data-ttu-id="c91b5-427">Zutreffender Outlook-Modus</span><span class="sxs-lookup"><span data-stu-id="c91b5-427">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="c91b5-428">Lesen</span><span class="sxs-lookup"><span data-stu-id="c91b5-428">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="c91b5-429">Beispiel</span><span class="sxs-lookup"><span data-stu-id="c91b5-429">Example</span></span>

```
var organizerName = Office.context.mailbox.item.organizer.displayName;
var organizerAddress = Office.context.mailbox.item.organizer.emailAddress;
```

####  <a name="requiredattendees-arrayemailaddressdetailsjavascriptapioutlook13officeemailaddressdetailsrecipientsjavascriptapioutlook13officerecipients"></a><span data-ttu-id="c91b5-430">requiredAttendees :Array.<[EmailAddressDetails](/javascript/api/outlook_1_3/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook_1_3/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="c91b5-430">requiredAttendees :Array.<[EmailAddressDetails](/javascript/api/outlook_1_3/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook_1_3/office.recipients)</span></span>

<span data-ttu-id="c91b5-431">Ermöglicht den Zugriff auf die erforderlichen Teilnehmer eines Ereignisses.</span><span class="sxs-lookup"><span data-stu-id="c91b5-431">Provides access to the required attendees of an event.</span></span> <span data-ttu-id="c91b5-432">Der Typ des Objekts und die Zugriffsebene, hängt von den Modus des aktuellen Elements ab.</span><span class="sxs-lookup"><span data-stu-id="c91b5-432">The type of object and level of access depends on the mode of the current item.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="c91b5-433">Lesemodus</span><span class="sxs-lookup"><span data-stu-id="c91b5-433">Read mode</span></span>

<span data-ttu-id="c91b5-434">Die `requiredAttendees`-Eigenschaft gibt ein Array mit einem `EmailAddressDetails`-Objekt für jeden erforderlichen Teilnehmer an der Besprechung zurück.</span><span class="sxs-lookup"><span data-stu-id="c91b5-434">The `requiredAttendees` property returns an array that contains an `EmailAddressDetails` object for each required attendee to the meeting.</span></span>

##### <a name="compose-mode"></a><span data-ttu-id="c91b5-435">Verfassenmodus</span><span class="sxs-lookup"><span data-stu-id="c91b5-435">Compose mode</span></span>

<span data-ttu-id="c91b5-436">Die `requiredAttendees` -Eigenschaft gibt eine `Recipients` -Objekt, das Methoden zum Abrufen oder Aktualisieren der erforderlichen Teilnehmer für eine Besprechung bereitstellt.</span><span class="sxs-lookup"><span data-stu-id="c91b5-436">The `requiredAttendees` property returns a `Recipients` object that provides methods to get or update the required attendees for a meeting.</span></span>

##### <a name="type"></a><span data-ttu-id="c91b5-437">Typ:</span><span class="sxs-lookup"><span data-stu-id="c91b5-437">Type:</span></span>

*   <span data-ttu-id="c91b5-438">Array.<[EmailAddressDetails](/javascript/api/outlook_1_3/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook_1_3/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="c91b5-438">Array.<[EmailAddressDetails](/javascript/api/outlook_1_3/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook_1_3/office.recipients)</span></span>

##### <a name="requirements"></a><span data-ttu-id="c91b5-439">Anforderungen</span><span class="sxs-lookup"><span data-stu-id="c91b5-439">Requirements</span></span>

|<span data-ttu-id="c91b5-440">Anforderung</span><span class="sxs-lookup"><span data-stu-id="c91b5-440">Requirement</span></span>| <span data-ttu-id="c91b5-441">Wert</span><span class="sxs-lookup"><span data-stu-id="c91b5-441">Value</span></span>|
|---|---|
|[<span data-ttu-id="c91b5-442">Mindestversion des Postfachanforderungssatzes</span><span class="sxs-lookup"><span data-stu-id="c91b5-442">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="c91b5-443">1.0</span><span class="sxs-lookup"><span data-stu-id="c91b5-443">1.0</span></span>|
|[<span data-ttu-id="c91b5-444">Mindestberechtigungsstufe</span><span class="sxs-lookup"><span data-stu-id="c91b5-444">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="c91b5-445">ReadItem</span><span class="sxs-lookup"><span data-stu-id="c91b5-445">ReadItem</span></span>|
|[<span data-ttu-id="c91b5-446">Zutreffender Outlook-Modus</span><span class="sxs-lookup"><span data-stu-id="c91b5-446">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="c91b5-447">Verfassen oder Lesen</span><span class="sxs-lookup"><span data-stu-id="c91b5-447">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="c91b5-448">Beispiel</span><span class="sxs-lookup"><span data-stu-id="c91b5-448">Example</span></span>

```
Office.context.mailbox.item.requiredAttendees.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.requiredAttendees.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.requiredAttendees.getAsync(callback);

function callback(asyncResult) {
  var arrayOfRequiredAttendeesRecipients = asyncResult.value;
}
```

#### <a name="sender-emailaddressdetailsjavascriptapioutlook13officeemailaddressdetails"></a><span data-ttu-id="c91b5-449">sender :[EmailAddressDetails](/javascript/api/outlook_1_3/office.emailaddressdetails)</span><span class="sxs-lookup"><span data-stu-id="c91b5-449">sender :[EmailAddressDetails](/javascript/api/outlook_1_3/office.emailaddressdetails)</span></span>

<span data-ttu-id="c91b5-p126">Ruft die E-Mail-Adresse des Absenders einer E-Mail-Nachricht ab. Nur Lesemodus.</span><span class="sxs-lookup"><span data-stu-id="c91b5-p126">Gets the email address of the sender of an email message. Read mode only.</span></span>

<span data-ttu-id="c91b5-p127">Die [`from`](#from-emailaddressdetailsjavascriptapioutlook13officeemailaddressdetails)- und `sender`-Eigenschaften stellen dieselbe Person dar, außer die Nachricht wurde von einem Delegaten gesendet. In diesem Fall stellt die `from`-Eigenschaft die Stellvertretung dar, und die sender-Eigenschaft stellt den Delegaten dar.</span><span class="sxs-lookup"><span data-stu-id="c91b5-p127">The [`from`](#from-emailaddressdetailsjavascriptapioutlook13officeemailaddressdetails) and `sender` properties represent the same person unless the message is sent by a delegate. In that case, the `from` property represents the delegator, and the sender property represents the delegate.</span></span>

> [!NOTE]
> <span data-ttu-id="c91b5-454">Die `recipientType` -Eigenschaft des der `EmailAddressDetails` -Objekts der `sender` -Eigenschaft ist `undefined`.</span><span class="sxs-lookup"><span data-stu-id="c91b5-454">The `recipientType` property of the `EmailAddressDetails` object in the `sender` property is `undefined`.</span></span>

##### <a name="type"></a><span data-ttu-id="c91b5-455">Typ:</span><span class="sxs-lookup"><span data-stu-id="c91b5-455">Type:</span></span>

*   [<span data-ttu-id="c91b5-456">EmailAddressDetails</span><span class="sxs-lookup"><span data-stu-id="c91b5-456">EmailAddressDetails</span></span>](/javascript/api/outlook_1_3/office.emailaddressdetails)

##### <a name="requirements"></a><span data-ttu-id="c91b5-457">Anforderungen</span><span class="sxs-lookup"><span data-stu-id="c91b5-457">Requirements</span></span>

|<span data-ttu-id="c91b5-458">Anforderung</span><span class="sxs-lookup"><span data-stu-id="c91b5-458">Requirement</span></span>| <span data-ttu-id="c91b5-459">Wert</span><span class="sxs-lookup"><span data-stu-id="c91b5-459">Value</span></span>|
|---|---|
|[<span data-ttu-id="c91b5-460">Mindestversion des Postfachanforderungssatzes</span><span class="sxs-lookup"><span data-stu-id="c91b5-460">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="c91b5-461">1.0</span><span class="sxs-lookup"><span data-stu-id="c91b5-461">1.0</span></span>|
|[<span data-ttu-id="c91b5-462">Mindestberechtigungsstufe</span><span class="sxs-lookup"><span data-stu-id="c91b5-462">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="c91b5-463">ReadItem</span><span class="sxs-lookup"><span data-stu-id="c91b5-463">ReadItem</span></span>|
|[<span data-ttu-id="c91b5-464">Zutreffender Outlook-Modus</span><span class="sxs-lookup"><span data-stu-id="c91b5-464">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="c91b5-465">Lesen</span><span class="sxs-lookup"><span data-stu-id="c91b5-465">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="c91b5-466">Beispiel</span><span class="sxs-lookup"><span data-stu-id="c91b5-466">Example</span></span>

```
var senderName = Office.context.mailbox.item.sender.displayName;
var senderAddress = Office.context.mailbox.item.sender.emailAddress;
```

####  <a name="start-datetimejavascriptapioutlook13officetime"></a><span data-ttu-id="c91b5-467">start :Date|[Time](/javascript/api/outlook_1_3/office.time)</span><span class="sxs-lookup"><span data-stu-id="c91b5-467">start :Date|[Time](/javascript/api/outlook_1_3/office.time)</span></span>

<span data-ttu-id="c91b5-468">Ruft Datum und Zeit für den Beginn des Termins ab oder legt Datum und Uhrzeit fest.</span><span class="sxs-lookup"><span data-stu-id="c91b5-468">Gets or sets the date and time that the appointment is to begin.</span></span>

<span data-ttu-id="c91b5-p128">Die `start`-Eigenschaft wird als Datums- und Uhrzeitwert in UTC ausgedrückt. Sie können die [`convertToLocalClientTime`](office.context.mailbox.md#converttolocalclienttimetimevalue--localclienttimejavascriptapioutlook13officelocalclienttime)-Methode verwenden, um den Wert in den lokalen Uhrzeit- und Datumswert des Clients umzuwandeln.</span><span class="sxs-lookup"><span data-stu-id="c91b5-p128">The `start` property is expressed as a Coordinated Universal Time (UTC) date and time value. You can use the [`convertToLocalClientTime`](office.context.mailbox.md#converttolocalclienttimetimevalue--localclienttimejavascriptapioutlook13officelocalclienttime) method to convert the value to the client’s local date and time.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="c91b5-471">Lesemodus</span><span class="sxs-lookup"><span data-stu-id="c91b5-471">Read mode</span></span>

<span data-ttu-id="c91b5-472">Die `start`-Eigenschaft gibt ein `Date`-Objekt zurück.</span><span class="sxs-lookup"><span data-stu-id="c91b5-472">The `start` property returns a `Date` object.</span></span>

##### <a name="compose-mode"></a><span data-ttu-id="c91b5-473">Verfassenmodus</span><span class="sxs-lookup"><span data-stu-id="c91b5-473">Compose mode</span></span>

<span data-ttu-id="c91b5-474">Die `start`-Eigenschaft gibt ein `Time`-Objekt zurück.</span><span class="sxs-lookup"><span data-stu-id="c91b5-474">The `start` property returns a `Time` object.</span></span>

<span data-ttu-id="c91b5-475">Wenn Sie die [`Time.setAsync`](/javascript/api/outlook_1_3/office.time#setasync-datetime--options--callback-)-Methode verwenden, um die Startzeit im Verfassenmodus festzulegen, sollten Sie die [`convertToUtcClientTime`](office.context.mailbox.md#converttoutcclienttimeinput--date)-Methode verwenden, um die Ortszeit auf dem Client für den Server in UTC umzuwandeln.</span><span class="sxs-lookup"><span data-stu-id="c91b5-475">When you use the [`Time.setAsync`](/javascript/api/outlook_1_3/office.time#setasync-datetime--options--callback-) method to set the start time, you should use the [`convertToUtcClientTime`](office.context.mailbox.md#converttoutcclienttimeinput--date) method to convert the local time on the client to UTC for the server.</span></span>

##### <a name="type"></a><span data-ttu-id="c91b5-476">Typ:</span><span class="sxs-lookup"><span data-stu-id="c91b5-476">Type:</span></span>

*   <span data-ttu-id="c91b5-477">Date | [Time](/javascript/api/outlook_1_3/office.time)</span><span class="sxs-lookup"><span data-stu-id="c91b5-477">Date | [Time](/javascript/api/outlook_1_3/office.time)</span></span>

##### <a name="requirements"></a><span data-ttu-id="c91b5-478">Anforderungen</span><span class="sxs-lookup"><span data-stu-id="c91b5-478">Requirements</span></span>

|<span data-ttu-id="c91b5-479">Anforderung</span><span class="sxs-lookup"><span data-stu-id="c91b5-479">Requirement</span></span>| <span data-ttu-id="c91b5-480">Wert</span><span class="sxs-lookup"><span data-stu-id="c91b5-480">Value</span></span>|
|---|---|
|[<span data-ttu-id="c91b5-481">Mindestversion des Postfachanforderungssatzes</span><span class="sxs-lookup"><span data-stu-id="c91b5-481">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="c91b5-482">1.0</span><span class="sxs-lookup"><span data-stu-id="c91b5-482">1.0</span></span>|
|[<span data-ttu-id="c91b5-483">Mindestberechtigungsstufe</span><span class="sxs-lookup"><span data-stu-id="c91b5-483">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="c91b5-484">ReadItem</span><span class="sxs-lookup"><span data-stu-id="c91b5-484">ReadItem</span></span>|
|[<span data-ttu-id="c91b5-485">Zutreffender Outlook-Modus</span><span class="sxs-lookup"><span data-stu-id="c91b5-485">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="c91b5-486">Verfassen oder Lesen</span><span class="sxs-lookup"><span data-stu-id="c91b5-486">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="c91b5-487">Beispiel</span><span class="sxs-lookup"><span data-stu-id="c91b5-487">Example</span></span>

<span data-ttu-id="c91b5-488">Im folgenden Beispiel wird die Startzeit eines Termins im Verfassenmodus mithilfe der [`setAsync`](/javascript/api/outlook_1_3/office.time#setasync-datetime--options--callback-)-Methode des `Time`-Objekts festgelegt.</span><span class="sxs-lookup"><span data-stu-id="c91b5-488">The following example sets the start time of an appointment in compose mode by using the [`setAsync`](/javascript/api/outlook_1_3/office.time#setasync-datetime--options--callback-) method of the `Time` object.</span></span>

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

####  <a name="subject-stringsubjectjavascriptapioutlook13officesubject"></a><span data-ttu-id="c91b5-489">subject :String|[Subject](/javascript/api/outlook_1_3/office.subject)</span><span class="sxs-lookup"><span data-stu-id="c91b5-489">subject :String|[Subject](/javascript/api/outlook_1_3/office.subject)</span></span>

<span data-ttu-id="c91b5-490">Ruft die Beschreibung ab, die im Betrefffeld eines Elements angezeigt wird, oder legt sie fest.</span><span class="sxs-lookup"><span data-stu-id="c91b5-490">Gets or sets the description that appears in the subject field of an item.</span></span>

<span data-ttu-id="c91b5-491">Die `subject`-Eigenschaft ruft den gesamten Betreff des Elements ab oder legt ihn fest – so, wie er vom E-Mail-Server gesendet wird.</span><span class="sxs-lookup"><span data-stu-id="c91b5-491">The `subject` property gets or sets the entire subject of the item, as sent by the email server.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="c91b5-492">Lesemodus</span><span class="sxs-lookup"><span data-stu-id="c91b5-492">Read mode</span></span>

<span data-ttu-id="c91b5-p129">Die `subject`-Eigenschaft gibt eine Zeichenfolge zurück. Verwenden Sie die [`normalizedSubject`](#normalizedsubject-string)-Eigenschaft, um den Betreff ohne führende Präfixe wie `RE:` und `FW:` abzurufen.</span><span class="sxs-lookup"><span data-stu-id="c91b5-p129">The `subject` property returns a string. Use the [`normalizedSubject`](#normalizedsubject-string) property to get the subject minus any leading prefixes such as `RE:` and `FW:`.</span></span>

```
var subject = Office.context.mailbox.item.subject;
```

##### <a name="compose-mode"></a><span data-ttu-id="c91b5-495">Verfassenmodus</span><span class="sxs-lookup"><span data-stu-id="c91b5-495">Compose mode</span></span>

<span data-ttu-id="c91b5-496">Die `subject`-Eigenschaft gibt ein `Subject`-Objekt zurück, das Methoden zum Abrufen und Festlegen des Betreffs bereitstellt.</span><span class="sxs-lookup"><span data-stu-id="c91b5-496">The `subject` property returns a `Subject` object that provides methods to get and set the subject.</span></span>

```
Office.context.mailbox.item.subject.getAsync(callback);

function callback(asyncResult) {
  var subject = asyncResult.value;
}
```

##### <a name="type"></a><span data-ttu-id="c91b5-497">Typ:</span><span class="sxs-lookup"><span data-stu-id="c91b5-497">Type:</span></span>

*   <span data-ttu-id="c91b5-498">String | [Subject](/javascript/api/outlook_1_3/office.subject)</span><span class="sxs-lookup"><span data-stu-id="c91b5-498">String | [Subject](/javascript/api/outlook_1_3/office.subject)</span></span>

##### <a name="requirements"></a><span data-ttu-id="c91b5-499">Anforderungen</span><span class="sxs-lookup"><span data-stu-id="c91b5-499">Requirements</span></span>

|<span data-ttu-id="c91b5-500">Anforderung</span><span class="sxs-lookup"><span data-stu-id="c91b5-500">Requirement</span></span>| <span data-ttu-id="c91b5-501">Wert</span><span class="sxs-lookup"><span data-stu-id="c91b5-501">Value</span></span>|
|---|---|
|[<span data-ttu-id="c91b5-502">Mindestversion des Postfachanforderungssatzes</span><span class="sxs-lookup"><span data-stu-id="c91b5-502">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="c91b5-503">1.0</span><span class="sxs-lookup"><span data-stu-id="c91b5-503">1.0</span></span>|
|[<span data-ttu-id="c91b5-504">Mindestberechtigungsstufe</span><span class="sxs-lookup"><span data-stu-id="c91b5-504">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="c91b5-505">ReadItem</span><span class="sxs-lookup"><span data-stu-id="c91b5-505">ReadItem</span></span>|
|[<span data-ttu-id="c91b5-506">Zutreffender Outlook-Modus</span><span class="sxs-lookup"><span data-stu-id="c91b5-506">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="c91b5-507">Verfassen oder Lesen</span><span class="sxs-lookup"><span data-stu-id="c91b5-507">Compose or read</span></span>|

####  <a name="to-arrayemailaddressdetailsjavascriptapioutlook13officeemailaddressdetailsrecipientsjavascriptapioutlook13officerecipients"></a><span data-ttu-id="c91b5-508">an: Array. <[EmailAddressDetails](/javascript/api/outlook_1_3/office.emailaddressdetails)>|[Empfänger](/javascript/api/outlook_1_3/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="c91b5-508">to :Array.<[EmailAddressDetails](/javascript/api/outlook_1_3/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook_1_3/office.recipients)</span></span>

<span data-ttu-id="c91b5-509">Ermöglicht den Zugriff auf die Empfänger in der Zeile **an** einer Nachricht.</span><span class="sxs-lookup"><span data-stu-id="c91b5-509">Provides access to the recipients on the **To** line of a message.</span></span> <span data-ttu-id="c91b5-510">Der Typ des Objekts und die Zugriffsebene, hängt von den Modus des aktuellen Elements ab.</span><span class="sxs-lookup"><span data-stu-id="c91b5-510">The type of object and level of access depends on the mode of the current item.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="c91b5-511">Lesemodus</span><span class="sxs-lookup"><span data-stu-id="c91b5-511">Read mode</span></span>

<span data-ttu-id="c91b5-p131">Die `to`-Eigenschaft gibt ein Array mit einem `EmailAddressDetails`-Objekt für jeden Empfänger in der**An**-Zeile der Nachricht zurück. Die Auflistung ist auf höchstens 100 Elemente beschränkt.</span><span class="sxs-lookup"><span data-stu-id="c91b5-p131">The `to` property returns an array that contains an `EmailAddressDetails` object for each recipient listed on the **To** line of the message. The collection is limited to a maximum of 100 members.</span></span>

##### <a name="compose-mode"></a><span data-ttu-id="c91b5-514">Verfassenmodus</span><span class="sxs-lookup"><span data-stu-id="c91b5-514">Compose mode</span></span>

<span data-ttu-id="c91b5-515">Die `to` -Eigenschaft gibt eine `Recipients` -Objekt, das Methoden zum Abrufen oder aktualisieren die Empfänger in der Zeile **an** der Nachricht bereitstellt.</span><span class="sxs-lookup"><span data-stu-id="c91b5-515">The `to` property returns a `Recipients` object that provides methods to get or update the recipients on the **To** line of the message.</span></span>

##### <a name="type"></a><span data-ttu-id="c91b5-516">Typ:</span><span class="sxs-lookup"><span data-stu-id="c91b5-516">Type:</span></span>

*   <span data-ttu-id="c91b5-517">Array.<[EmailAddressDetails](/javascript/api/outlook_1_3/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook_1_3/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="c91b5-517">Array.<[EmailAddressDetails](/javascript/api/outlook_1_3/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook_1_3/office.recipients)</span></span>

##### <a name="requirements"></a><span data-ttu-id="c91b5-518">Anforderungen</span><span class="sxs-lookup"><span data-stu-id="c91b5-518">Requirements</span></span>

|<span data-ttu-id="c91b5-519">Anforderung</span><span class="sxs-lookup"><span data-stu-id="c91b5-519">Requirement</span></span>| <span data-ttu-id="c91b5-520">Wert</span><span class="sxs-lookup"><span data-stu-id="c91b5-520">Value</span></span>|
|---|---|
|[<span data-ttu-id="c91b5-521">Mindestversion des Postfachanforderungssatzes</span><span class="sxs-lookup"><span data-stu-id="c91b5-521">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="c91b5-522">1.0</span><span class="sxs-lookup"><span data-stu-id="c91b5-522">1.0</span></span>|
|[<span data-ttu-id="c91b5-523">Mindestberechtigungsstufe</span><span class="sxs-lookup"><span data-stu-id="c91b5-523">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="c91b5-524">ReadItem</span><span class="sxs-lookup"><span data-stu-id="c91b5-524">ReadItem</span></span>|
|[<span data-ttu-id="c91b5-525">Zutreffender Outlook-Modus</span><span class="sxs-lookup"><span data-stu-id="c91b5-525">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="c91b5-526">Verfassen oder Lesen</span><span class="sxs-lookup"><span data-stu-id="c91b5-526">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="c91b5-527">Beispiel</span><span class="sxs-lookup"><span data-stu-id="c91b5-527">Example</span></span>

```
Office.context.mailbox.item.to.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.to.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.to.getAsync(callback);

function callback(asyncResult) {
  var arrayOfToRecipients = asyncResult.value;
}
```

### <a name="methods"></a><span data-ttu-id="c91b5-528">Methoden</span><span class="sxs-lookup"><span data-stu-id="c91b5-528">Methods</span></span>

####  <a name="addfileattachmentasyncuri-attachmentname-options-callback"></a><span data-ttu-id="c91b5-529">addFileAttachmentAsync(uri, attachmentName, [options], [callback])</span><span class="sxs-lookup"><span data-stu-id="c91b5-529">addFileAttachmentAsync(uri, attachmentName, [options], [callback])</span></span>

<span data-ttu-id="c91b5-530">Fügt eine Datei zu einer Nachricht oder einem Termin als Anlage hinzu.</span><span class="sxs-lookup"><span data-stu-id="c91b5-530">Adds a file to a message or appointment as an attachment.</span></span>

<span data-ttu-id="c91b5-531">Die `addFileAttachmentAsync`-Methode lädt die Datei am angegebenen URI hoch und fügt sie an das Element im Verfassenformular an.</span><span class="sxs-lookup"><span data-stu-id="c91b5-531">The `addFileAttachmentAsync` method uploads the file at the specified URI and attaches it to the item in the compose form.</span></span>

<span data-ttu-id="c91b5-532">Anschließend können Sie den Bezeichner mit der [`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback)-Methode in der gleichen Sitzung zum Entfernen der Anlage verwenden.</span><span class="sxs-lookup"><span data-stu-id="c91b5-532">You can subsequently use the identifier with the [`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback) method to remove the attachment in the same session.</span></span>

##### <a name="parameters"></a><span data-ttu-id="c91b5-533">Parameter:</span><span class="sxs-lookup"><span data-stu-id="c91b5-533">Parameters:</span></span>

|<span data-ttu-id="c91b5-534">Name</span><span class="sxs-lookup"><span data-stu-id="c91b5-534">Name</span></span>| <span data-ttu-id="c91b5-535">Typ</span><span class="sxs-lookup"><span data-stu-id="c91b5-535">Type</span></span>| <span data-ttu-id="c91b5-536">Attribute</span><span class="sxs-lookup"><span data-stu-id="c91b5-536">Attributes</span></span>| <span data-ttu-id="c91b5-537">Beschreibung</span><span class="sxs-lookup"><span data-stu-id="c91b5-537">Description</span></span>|
|---|---|---|---|
|`uri`| <span data-ttu-id="c91b5-538">Zeichenfolge</span><span class="sxs-lookup"><span data-stu-id="c91b5-538">String</span></span>||<span data-ttu-id="c91b5-p132">Der URI, der den Speicherort der an die Nachricht oder den Termin anzuhängenden Datei angibt. Die maximale Länge ist 2048 Zeichen.</span><span class="sxs-lookup"><span data-stu-id="c91b5-p132">The URI that provides the location of the file to attach to the message or appointment. The maximum length is 2048 characters.</span></span>|
|`attachmentName`| <span data-ttu-id="c91b5-541">Zeichenfolge</span><span class="sxs-lookup"><span data-stu-id="c91b5-541">String</span></span>||<span data-ttu-id="c91b5-p133">Der Name der Anlage, der beim Hochladen der Anlage angezeigt wird. Die maximale Länge ist 255 Zeichen.</span><span class="sxs-lookup"><span data-stu-id="c91b5-p133">The name of the attachment that is shown while the attachment is uploading. The maximum length is 255 characters.</span></span>|
|`options`| <span data-ttu-id="c91b5-544">Object</span><span class="sxs-lookup"><span data-stu-id="c91b5-544">Object</span></span>| <span data-ttu-id="c91b5-545">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="c91b5-545">&lt;optional&gt;</span></span>|<span data-ttu-id="c91b5-546">Ein Objektliteral, das mindestens eine der folgenden Eigenschaften enthält.</span><span class="sxs-lookup"><span data-stu-id="c91b5-546">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`| <span data-ttu-id="c91b5-547">Object</span><span class="sxs-lookup"><span data-stu-id="c91b5-547">Object</span></span>| <span data-ttu-id="c91b5-548">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="c91b5-548">&lt;optional&gt;</span></span>|<span data-ttu-id="c91b5-549">Entwickler können ein Objekt bereitstellen, auf das sie in der Callbackmethode zugreifen möchten.</span><span class="sxs-lookup"><span data-stu-id="c91b5-549">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`| <span data-ttu-id="c91b5-550">Funktion</span><span class="sxs-lookup"><span data-stu-id="c91b5-550">function</span></span>| <span data-ttu-id="c91b5-551">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="c91b5-551">&lt;optional&gt;</span></span>|<span data-ttu-id="c91b5-552">Wenn die Methode abgeschlossen wird, wird die Funktion , die im `callback`-Parameter übergeben wird, mit einem einzelnen Parameter (`asyncResult`) aufgerufen, der ein [`AsyncResult`](/javascript/api/office/office.asyncresult)-Objekt darstellt.</span><span class="sxs-lookup"><span data-stu-id="c91b5-552">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span> <br/><span data-ttu-id="c91b5-553">Bei Erfolg wird der Anlagenbezeichner in der `asyncResult.value`-Eigenschaft bereitgestellt.</span><span class="sxs-lookup"><span data-stu-id="c91b5-553">On success, the attachment identifier will be provided in the `asyncResult.value` property.</span></span><br/><span data-ttu-id="c91b5-554">Wenn beim Hochladen der Anlage ein Fehler auftritt, enthält das `asyncResult`-Objekt ein `Error`-Objekt mit einer Beschreibung des Fehlers.</span><span class="sxs-lookup"><span data-stu-id="c91b5-554">If uploading the attachment fails, the `asyncResult` object will contain an `Error` object that provides a description of the error.</span></span>|

##### <a name="errors"></a><span data-ttu-id="c91b5-555">Fehler</span><span class="sxs-lookup"><span data-stu-id="c91b5-555">Errors</span></span>

| <span data-ttu-id="c91b5-556">Fehlercode</span><span class="sxs-lookup"><span data-stu-id="c91b5-556">Error code</span></span> | <span data-ttu-id="c91b5-557">Beschreibung</span><span class="sxs-lookup"><span data-stu-id="c91b5-557">Description</span></span> |
|------------|-------------|
| `AttachmentSizeExceeded` | <span data-ttu-id="c91b5-558">Die Anlage ist zu groß.</span><span class="sxs-lookup"><span data-stu-id="c91b5-558">The attachment is larger than allowed.</span></span> |
| `FileTypeNotSupported` | <span data-ttu-id="c91b5-559">Die Anlage enthält eine Erweiterung, die nicht zulässig ist.</span><span class="sxs-lookup"><span data-stu-id="c91b5-559">The attachment has an extension that is not allowed.</span></span> |
| `NumberOfAttachmentsExceeded` | <span data-ttu-id="c91b5-560">Die Nachricht oder der Termin enthält zu viele Anlagen.</span><span class="sxs-lookup"><span data-stu-id="c91b5-560">The message or appointment has too many attachments.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="c91b5-561">Anforderungen</span><span class="sxs-lookup"><span data-stu-id="c91b5-561">Requirements</span></span>

|<span data-ttu-id="c91b5-562">Anforderung</span><span class="sxs-lookup"><span data-stu-id="c91b5-562">Requirement</span></span>| <span data-ttu-id="c91b5-563">Wert</span><span class="sxs-lookup"><span data-stu-id="c91b5-563">Value</span></span>|
|---|---|
|[<span data-ttu-id="c91b5-564">Mindestversion des Postfachanforderungssatzes</span><span class="sxs-lookup"><span data-stu-id="c91b5-564">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="c91b5-565">1.1</span><span class="sxs-lookup"><span data-stu-id="c91b5-565">1.1</span></span>|
|[<span data-ttu-id="c91b5-566">Mindestberechtigungsstufe</span><span class="sxs-lookup"><span data-stu-id="c91b5-566">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="c91b5-567">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="c91b5-567">ReadWriteItem</span></span>|
|[<span data-ttu-id="c91b5-568">Zutreffender Outlook-Modus</span><span class="sxs-lookup"><span data-stu-id="c91b5-568">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="c91b5-569">Verfassen</span><span class="sxs-lookup"><span data-stu-id="c91b5-569">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="c91b5-570">Beispiel</span><span class="sxs-lookup"><span data-stu-id="c91b5-570">Example</span></span>

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

####  <a name="additemattachmentasyncitemid-attachmentname-options-callback"></a><span data-ttu-id="c91b5-571">addItemAttachmentAsync(itemId, attachmentName, [options], [callback])</span><span class="sxs-lookup"><span data-stu-id="c91b5-571">addItemAttachmentAsync(itemId, attachmentName, [options], [callback])</span></span>

<span data-ttu-id="c91b5-572">Fügt der Nachricht oder dem Termin ein Exchange-Objekt, wie z. B. eine Nachricht, als Anhang hinzu.</span><span class="sxs-lookup"><span data-stu-id="c91b5-572">Adds an Exchange item, such as a message, as an attachment to the message or appointment.</span></span>

<span data-ttu-id="c91b5-p134">Die `addItemAttachmentAsync`-Methode fügt das Element mit dem angegebenen Exchange-Bezeichner an das Element im Formular zum Verfassen an. Wenn Sie eine Rückrufmethode angeben, wird die Methode mit einem `asyncResult`-Parameter aufgerufen, der entweder den Anlagenbezeichner oder einen Code enthält, der etwaige Fehler angibt, die beim Anfügen des Objekts aufgetreten sind. Sie können ggf. den `options`-Parameter verwenden, um Statusinformationen an die Rückrufmethode zu übergeben.</span><span class="sxs-lookup"><span data-stu-id="c91b5-p134">The `addItemAttachmentAsync` method attaches the item with the specified Exchange identifier to the item in the compose form. If you specify a callback method, the method is called with one parameter, `asyncResult`, which contains either the attachment identifier or a code that indicates any error that occurred while attaching the item. You can use the `options` parameter to pass state information to the callback method, if needed.</span></span>

<span data-ttu-id="c91b5-576">Anschließend können Sie den Bezeichner mit der [`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback)-Methode in der gleichen Sitzung zum Entfernen der Anlage verwenden.</span><span class="sxs-lookup"><span data-stu-id="c91b5-576">You can subsequently use the identifier with the [`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback) method to remove the attachment in the same session.</span></span>

<span data-ttu-id="c91b5-577">Wenn Ihre Office-Add-Ins in Outlook Web App ausgeführt wird die `addItemAttachmentAsync` -Methode kann Anfügen von Elementen in der Elemente, die Sie bearbeiten; Allerdings Dies wird nicht unterstützt und wird nicht empfohlen.</span><span class="sxs-lookup"><span data-stu-id="c91b5-577">If your Office Add-in is running in Outlook Web App, the `addItemAttachmentAsync` method can attach items to items other than the item that you are editing; however, this is not supported and is not recommended.</span></span>

##### <a name="parameters"></a><span data-ttu-id="c91b5-578">Parameter:</span><span class="sxs-lookup"><span data-stu-id="c91b5-578">Parameters:</span></span>

|<span data-ttu-id="c91b5-579">Name</span><span class="sxs-lookup"><span data-stu-id="c91b5-579">Name</span></span>| <span data-ttu-id="c91b5-580">Typ</span><span class="sxs-lookup"><span data-stu-id="c91b5-580">Type</span></span>| <span data-ttu-id="c91b5-581">Attribute</span><span class="sxs-lookup"><span data-stu-id="c91b5-581">Attributes</span></span>| <span data-ttu-id="c91b5-582">Beschreibung</span><span class="sxs-lookup"><span data-stu-id="c91b5-582">Description</span></span>|
|---|---|---|---|
|`itemId`| <span data-ttu-id="c91b5-583">Zeichenfolge</span><span class="sxs-lookup"><span data-stu-id="c91b5-583">String</span></span>||<span data-ttu-id="c91b5-p135">Der Exchange-Bezeichner des Objekts, das angehängt werden soll. Die maximale Länge beträgt 100 Zeichen.</span><span class="sxs-lookup"><span data-stu-id="c91b5-p135">The Exchange identifier of the item to attach. The maximum length is 100 characters.</span></span>|
|`attachmentName`| <span data-ttu-id="c91b5-586">Zeichenfolge</span><span class="sxs-lookup"><span data-stu-id="c91b5-586">String</span></span>||<span data-ttu-id="c91b5-p136">Der Betreff des Elements, das angehängt werden soll. Die maximale Länge ist 255 Zeichen.</span><span class="sxs-lookup"><span data-stu-id="c91b5-p136">The sujbect of the item to be attached. The maximum length is 255 characters.</span></span>|
|`options`| <span data-ttu-id="c91b5-589">Object</span><span class="sxs-lookup"><span data-stu-id="c91b5-589">Object</span></span>| <span data-ttu-id="c91b5-590">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="c91b5-590">&lt;optional&gt;</span></span>|<span data-ttu-id="c91b5-591">Ein Objektliteral, das mindestens eine der folgenden Eigenschaften enthält.</span><span class="sxs-lookup"><span data-stu-id="c91b5-591">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`| <span data-ttu-id="c91b5-592">Object</span><span class="sxs-lookup"><span data-stu-id="c91b5-592">Object</span></span>| <span data-ttu-id="c91b5-593">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="c91b5-593">&lt;optional&gt;</span></span>|<span data-ttu-id="c91b5-594">Entwickler können ein Objekt bereitstellen, auf das sie in der Callbackmethode zugreifen möchten.</span><span class="sxs-lookup"><span data-stu-id="c91b5-594">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`| <span data-ttu-id="c91b5-595">Funktion</span><span class="sxs-lookup"><span data-stu-id="c91b5-595">function</span></span>| <span data-ttu-id="c91b5-596">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="c91b5-596">&lt;optional&gt;</span></span>|<span data-ttu-id="c91b5-597">Wenn die Methode abgeschlossen wird, wird die Funktion , die im `callback`-Parameter übergeben wird, mit einem einzelnen Parameter (`asyncResult`) aufgerufen, der ein [`AsyncResult`](/javascript/api/office/office.asyncresult)-Objekt darstellt.</span><span class="sxs-lookup"><span data-stu-id="c91b5-597">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span> <br/><span data-ttu-id="c91b5-598">Bei Erfolg wird der Anlagenbezeichner in der `asyncResult.value`-Eigenschaft bereitgestellt.</span><span class="sxs-lookup"><span data-stu-id="c91b5-598">On success, the attachment identifier will be provided in the `asyncResult.value` property.</span></span><br/><span data-ttu-id="c91b5-599">Wenn beim Hinzufügen der Anlage ein Fehler auftritt, enthält das `asyncResult`-Objekt ein `Error`-Objekt mit einer Beschreibung des Fehlers.</span><span class="sxs-lookup"><span data-stu-id="c91b5-599">If adding the attachment fails, the `asyncResult` object will contain an `Error` object that provides a description of the error.</span></span>|

##### <a name="errors"></a><span data-ttu-id="c91b5-600">Fehler</span><span class="sxs-lookup"><span data-stu-id="c91b5-600">Errors</span></span>

| <span data-ttu-id="c91b5-601">Fehlercode</span><span class="sxs-lookup"><span data-stu-id="c91b5-601">Error code</span></span> | <span data-ttu-id="c91b5-602">Beschreibung</span><span class="sxs-lookup"><span data-stu-id="c91b5-602">Description</span></span> |
|------------|-------------|
| `NumberOfAttachmentsExceeded` | <span data-ttu-id="c91b5-603">Die Nachricht oder der Termin enthält zu viele Anlagen.</span><span class="sxs-lookup"><span data-stu-id="c91b5-603">The message or appointment has too many attachments.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="c91b5-604">Anforderungen</span><span class="sxs-lookup"><span data-stu-id="c91b5-604">Requirements</span></span>

|<span data-ttu-id="c91b5-605">Anforderung</span><span class="sxs-lookup"><span data-stu-id="c91b5-605">Requirement</span></span>| <span data-ttu-id="c91b5-606">Wert</span><span class="sxs-lookup"><span data-stu-id="c91b5-606">Value</span></span>|
|---|---|
|[<span data-ttu-id="c91b5-607">Mindestversion des Postfachanforderungssatzes</span><span class="sxs-lookup"><span data-stu-id="c91b5-607">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="c91b5-608">1.1</span><span class="sxs-lookup"><span data-stu-id="c91b5-608">1.1</span></span>|
|[<span data-ttu-id="c91b5-609">Mindestberechtigungsstufe</span><span class="sxs-lookup"><span data-stu-id="c91b5-609">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="c91b5-610">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="c91b5-610">ReadWriteItem</span></span>|
|[<span data-ttu-id="c91b5-611">Zutreffender Outlook-Modus</span><span class="sxs-lookup"><span data-stu-id="c91b5-611">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="c91b5-612">Verfassen</span><span class="sxs-lookup"><span data-stu-id="c91b5-612">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="c91b5-613">Beispiel</span><span class="sxs-lookup"><span data-stu-id="c91b5-613">Example</span></span>

<span data-ttu-id="c91b5-614">Im folgenden Beispiel wird ein vorhandenes Outlook-Element als Anlage mit dem Namen `My Attachment` hinzugefügt.</span><span class="sxs-lookup"><span data-stu-id="c91b5-614">The following example adds an existing Outlook item as an attachment with the name `My Attachment`.</span></span>

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

####  <a name="close"></a><span data-ttu-id="c91b5-615">close()</span><span class="sxs-lookup"><span data-stu-id="c91b5-615">close()</span></span>

<span data-ttu-id="c91b5-616">Schließt das aktuelle Element, das gerade verfasst wird.</span><span class="sxs-lookup"><span data-stu-id="c91b5-616">Closes the current item that is being composed.</span></span>

<span data-ttu-id="c91b5-p137">Das Verhalten der `close`-Methode hängt vom aktuellen Status des verfassten Elements ab. Wenn das Element über nicht gespeicherte Änderungen verfügt, fordert der Client den Benutzer auf, die Schließenaktion zu speichern, zu verwerfen oder abzubrechen.</span><span class="sxs-lookup"><span data-stu-id="c91b5-p137">The behavior of the `close` method depends on the current state of the item being composed. If the item has unsaved changes, the client prompts the user to save, discard, or cancel the close action.</span></span>

> [!NOTE]
> <span data-ttu-id="c91b5-619">In Outlook im Web, wenn das Element einen Termin handelt, und es wurde bereits zuvor gespeichert wurde mit `saveAsync`, der Benutzer wird aufgefordert, speichern, löschen oder Abbrechen, auch wenn keine Änderungen aufgetreten sind, da das Element zuletzt gespeichert.</span><span class="sxs-lookup"><span data-stu-id="c91b5-619">In Outlook on the web, if the item is an appointment and it has previously been saved using `saveAsync`, the user is prompted to save, discard, or cancel even if no changes have occurred since the item was last saved.</span></span>

<span data-ttu-id="c91b5-620">Wenn die Nachricht im Outlook-Desktopclient eine Inlineantwort ist, hat die `close`-Methode keine Auswirkung.</span><span class="sxs-lookup"><span data-stu-id="c91b5-620">In the Outlook desktop client, if the message is an inline reply, the `close` method has no effect.</span></span>

##### <a name="requirements"></a><span data-ttu-id="c91b5-621">Anforderungen</span><span class="sxs-lookup"><span data-stu-id="c91b5-621">Requirements</span></span>

|<span data-ttu-id="c91b5-622">Anforderung</span><span class="sxs-lookup"><span data-stu-id="c91b5-622">Requirement</span></span>| <span data-ttu-id="c91b5-623">Wert</span><span class="sxs-lookup"><span data-stu-id="c91b5-623">Value</span></span>|
|---|---|
|[<span data-ttu-id="c91b5-624">Mindestversion des Postfachanforderungssatzes</span><span class="sxs-lookup"><span data-stu-id="c91b5-624">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="c91b5-625">1.3</span><span class="sxs-lookup"><span data-stu-id="c91b5-625">1.3</span></span>|
|[<span data-ttu-id="c91b5-626">Mindestberechtigungsstufe</span><span class="sxs-lookup"><span data-stu-id="c91b5-626">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="c91b5-627">Eingeschränkt</span><span class="sxs-lookup"><span data-stu-id="c91b5-627">Restricted</span></span>|
|[<span data-ttu-id="c91b5-628">Zutreffender Outlook-Modus</span><span class="sxs-lookup"><span data-stu-id="c91b5-628">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="c91b5-629">Verfassen</span><span class="sxs-lookup"><span data-stu-id="c91b5-629">Compose</span></span>|

#### <a name="displayreplyallformformdata"></a><span data-ttu-id="c91b5-630">displayReplyAllForm(formData)</span><span class="sxs-lookup"><span data-stu-id="c91b5-630">displayReplyAllForm(formData)</span></span>

<span data-ttu-id="c91b5-631">Zeigt ein Antwortformular an, das den Absender und alle Empfänger der ausgewählten Nachricht oder des Organisators und alle Teilnehmer des ausgewählten Termins enthält.</span><span class="sxs-lookup"><span data-stu-id="c91b5-631">Displays a reply form that includes the sender and all recipients of the selected message or the organizer and all attendees of the selected appointment.</span></span>

> [!NOTE]
> <span data-ttu-id="c91b5-632">Diese Methode wird in Outlook für iOS oder Outlook für Android nicht unterstützt.</span><span class="sxs-lookup"><span data-stu-id="c91b5-632">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="c91b5-633">In Outlook Web App wird das Antwortformular als Popupformular in der Dreispaltenansicht und als Popupformular in der Zwei- oder Einspaltenansicht angezeigt.</span><span class="sxs-lookup"><span data-stu-id="c91b5-633">In Outlook Web App, the reply form is displayed as a pop-out form in the 3-column view and a pop-up form in the 2- or 1-column view.</span></span>

<span data-ttu-id="c91b5-634">Wenn einer der Zeichenfolgenparameter seinen Grenzwert überschreitet, löst `displayReplyAllForm` eine Ausnahme aus.</span><span class="sxs-lookup"><span data-stu-id="c91b5-634">If any of the string parameters exceed their limits, `displayReplyAllForm` throws an exception.</span></span>

<span data-ttu-id="c91b5-p138">Wenn Anlagen im `formData.attachments`-Parameter angegeben werden, versuchen Outlook und Outlook Web App alle Anlagen herunterzuladen und sie an das Antwortformular anzufügen. Wenn Anlagen nicht hinzugefügt werden können, wird in der Formularbenutzeroberfläche ein Fehler ausgegeben. Wenn dies nicht möglich ist, wird keine Fehlermeldung ausgegeben.</span><span class="sxs-lookup"><span data-stu-id="c91b5-p138">When attachments are specified in the `formData.attachments` parameter, Outlook and Outlook Web App attempt to download all attachments and attach them to the reply form. If any attachments fail to be added, an error is shown in the form UI. If this isn't possible, then no error message is thrown.</span></span>

##### <a name="parameters"></a><span data-ttu-id="c91b5-638">Parameter:</span><span class="sxs-lookup"><span data-stu-id="c91b5-638">Parameters:</span></span>

|<span data-ttu-id="c91b5-639">Name</span><span class="sxs-lookup"><span data-stu-id="c91b5-639">Name</span></span>| <span data-ttu-id="c91b5-640">Typ</span><span class="sxs-lookup"><span data-stu-id="c91b5-640">Type</span></span>| <span data-ttu-id="c91b5-641">Beschreibung</span><span class="sxs-lookup"><span data-stu-id="c91b5-641">Description</span></span>|
|---|---|---|
|`formData`| <span data-ttu-id="c91b5-642">Zeichenfolge &#124; Objekt</span><span class="sxs-lookup"><span data-stu-id="c91b5-642">String &#124; Object</span></span>| |<span data-ttu-id="c91b5-p139">Eine Zeichenfolge, die Text- und HTML-Code enthält, die den Hauptteil des Antwortformulars darstellen. Die Zeichenfolge ist auf 32 KB beschränkt.</span><span class="sxs-lookup"><span data-stu-id="c91b5-p139">A string that contains text and HTML and that represents the body of the reply form. The string is limited to 32 KB.</span></span><br/><span data-ttu-id="c91b5-645">**ODER**</span><span class="sxs-lookup"><span data-stu-id="c91b5-645">**OR**</span></span><br/><span data-ttu-id="c91b5-p140">Ein Objekt, das Text- oder Anlagendaten und eine Rückruffunktion enthält. Das Objekt ist wie folgt definiert:</span><span class="sxs-lookup"><span data-stu-id="c91b5-p140">An object that contains body or attachment data and a callback function. The object is defined as follows.</span></span> |
| `formData.htmlBody` | <span data-ttu-id="c91b5-648">Zeichenfolge</span><span class="sxs-lookup"><span data-stu-id="c91b5-648">String</span></span> | <span data-ttu-id="c91b5-649">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="c91b5-649">&lt;optional&gt;</span></span> | <span data-ttu-id="c91b5-p141">Eine Zeichenfolge, die Text- und HTML-Code enthält, die den Hauptteil des Antwortformulars darstellen. Die Zeichenfolge ist auf 32 KB beschränkt.</span><span class="sxs-lookup"><span data-stu-id="c91b5-p141">A string that contains text and HTML and that represents the body of the reply form. The string is limited to 32 KB.</span></span>
| `formData.attachments` | <span data-ttu-id="c91b5-652">Array.&lt;Object&gt;</span><span class="sxs-lookup"><span data-stu-id="c91b5-652">Array.&lt;Object&gt;</span></span> | <span data-ttu-id="c91b5-653">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="c91b5-653">&lt;optional&gt;</span></span> | <span data-ttu-id="c91b5-654">Ein Array mit JSON-Objekten, die Datei- oder Elementanlagen sind.</span><span class="sxs-lookup"><span data-stu-id="c91b5-654">An array of JSON objects that are either file or item attachments.</span></span> |
| `formData.attachments.type` | <span data-ttu-id="c91b5-655">Zeichenfolge</span><span class="sxs-lookup"><span data-stu-id="c91b5-655">String</span></span> | | <span data-ttu-id="c91b5-p142">Gibt den Typ der Anlage an. Muss `file` für eine Dateianlage oder `item` für eine Elementanlage sein.</span><span class="sxs-lookup"><span data-stu-id="c91b5-p142">Indicates the type of attachment. Must be `file` for a file attachment or `item` for an item attachment.</span></span> |
| `formData.attachments.name` | <span data-ttu-id="c91b5-658">Zeichenfolge</span><span class="sxs-lookup"><span data-stu-id="c91b5-658">String</span></span> | | <span data-ttu-id="c91b5-659">Eine Zeichenfolge, die den Namen der Anlage mit bis zu 255 Zeichen enthält.</span><span class="sxs-lookup"><span data-stu-id="c91b5-659">A string that contains the name of the attachment, up to 255 characters in length.</span></span>|
| `formData.attachments.url` | <span data-ttu-id="c91b5-660">Zeichenfolge</span><span class="sxs-lookup"><span data-stu-id="c91b5-660">String</span></span> | | <span data-ttu-id="c91b5-p143">Wird nur verwendet, wenn `type` auf `file` gesetzt ist. Der URI des Speicherorts für die Datei.</span><span class="sxs-lookup"><span data-stu-id="c91b5-p143">Only used if `type` is set to `file`. The URI of the location for the file.</span></span> |
| `formData.attachments.itemId` | <span data-ttu-id="c91b5-663">Zeichenfolge</span><span class="sxs-lookup"><span data-stu-id="c91b5-663">String</span></span> | | <span data-ttu-id="c91b5-p144">Wird nur verwendet, wenn `type` auf `item` gesetzt ist. Die EWS-Element-ID der Anlage. Dies ist eine Zeichenfolge bis zu 100 Zeichen.</span><span class="sxs-lookup"><span data-stu-id="c91b5-p144">Only used if `type` is set to `item`. The EWS item id of the attachment. This is a string up to 100 characters.</span></span> |
| `callback` | <span data-ttu-id="c91b5-667">Funktion</span><span class="sxs-lookup"><span data-stu-id="c91b5-667">function</span></span> | <span data-ttu-id="c91b5-668">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="c91b5-668">&lt;optional&gt;</span></span> | <span data-ttu-id="c91b5-669">Bei Abschluss der Methode wird die im Parameter `callback` übergebene Funktion mit einem einzigen Parameter aufgerufen: `asyncResult`. Dieser Parameter ist ein Objekt des Typs [AsyncResult](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="c91b5-669">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [AsyncResult](/javascript/api/office/office.asyncresult) object.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="c91b5-670">Anforderungen</span><span class="sxs-lookup"><span data-stu-id="c91b5-670">Requirements</span></span>

|<span data-ttu-id="c91b5-671">Anforderung</span><span class="sxs-lookup"><span data-stu-id="c91b5-671">Requirement</span></span>| <span data-ttu-id="c91b5-672">Wert</span><span class="sxs-lookup"><span data-stu-id="c91b5-672">Value</span></span>|
|---|---|
|[<span data-ttu-id="c91b5-673">Mindestversion des Postfachanforderungssatzes</span><span class="sxs-lookup"><span data-stu-id="c91b5-673">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="c91b5-674">1.0</span><span class="sxs-lookup"><span data-stu-id="c91b5-674">1.0</span></span>|
|[<span data-ttu-id="c91b5-675">Mindestberechtigungsstufe</span><span class="sxs-lookup"><span data-stu-id="c91b5-675">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="c91b5-676">ReadItem</span><span class="sxs-lookup"><span data-stu-id="c91b5-676">ReadItem</span></span>|
|[<span data-ttu-id="c91b5-677">Zutreffender Outlook-Modus</span><span class="sxs-lookup"><span data-stu-id="c91b5-677">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="c91b5-678">Lesen</span><span class="sxs-lookup"><span data-stu-id="c91b5-678">Read</span></span>|

##### <a name="examples"></a><span data-ttu-id="c91b5-679">Beispiele</span><span class="sxs-lookup"><span data-stu-id="c91b5-679">Examples</span></span>

<span data-ttu-id="c91b5-680">Im folgenden Code wird eine Zeichenfolge an die `displayReplyAllForm`-Funktion übergeben.</span><span class="sxs-lookup"><span data-stu-id="c91b5-680">The following code passes a string to the `displayReplyAllForm` function.</span></span>

```
Office.context.mailbox.item.displayReplyAllForm('hello there');
Office.context.mailbox.item.displayReplyAllForm('<b>hello there</b>');
```

<span data-ttu-id="c91b5-681">Antworten mit einem leeren Textkörper.</span><span class="sxs-lookup"><span data-stu-id="c91b5-681">Reply with an empty body.</span></span>

```
Office.context.mailbox.item.displayReplyAllForm({});
```

<span data-ttu-id="c91b5-682">Antworten nur einem Textkörper.</span><span class="sxs-lookup"><span data-stu-id="c91b5-682">Reply with just a body.</span></span>

```
Office.context.mailbox.item.displayReplyAllForm(
{
  'htmlBody' : 'hi'
});
```

<span data-ttu-id="c91b5-683">Antworten mit einem Textkörper und einer Dateianlage.</span><span class="sxs-lookup"><span data-stu-id="c91b5-683">Reply with a body and a file attachment.</span></span>

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

<span data-ttu-id="c91b5-684">Antworten mit einem Textkörper und einer Elementanlage.</span><span class="sxs-lookup"><span data-stu-id="c91b5-684">Reply with a body and an item attachment.</span></span>

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

<span data-ttu-id="c91b5-685">Antworten Sie mit einem Textkörper, einer Dateianlage, einer Elementanlage und einem Callback.</span><span class="sxs-lookup"><span data-stu-id="c91b5-685">Reply with a body, file attachment, item attachment, and a callback.</span></span>

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

#### <a name="displayreplyformformdata"></a><span data-ttu-id="c91b5-686">displayReplyForm(formData)</span><span class="sxs-lookup"><span data-stu-id="c91b5-686">displayReplyForm(formData)</span></span>

<span data-ttu-id="c91b5-687">Zeigt ein Antwortformular an, das nur den Absender der ausgewählten Nachricht oder den Organisator des ausgewählten Termins enthält.</span><span class="sxs-lookup"><span data-stu-id="c91b5-687">Displays a reply form that includes only the sender of the selected message or the organizer of the selected appointment.</span></span>

> [!NOTE]
> <span data-ttu-id="c91b5-688">Diese Methode wird in Outlook für iOS oder Outlook für Android nicht unterstützt.</span><span class="sxs-lookup"><span data-stu-id="c91b5-688">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="c91b5-689">In Outlook Web App wird das Antwortformular als Popupformular in der Dreispaltenansicht und als Popupformular in der Zwei- oder Einspaltenansicht angezeigt.</span><span class="sxs-lookup"><span data-stu-id="c91b5-689">In Outlook Web App, the reply form is displayed as a pop-out form in the 3-column view and a pop-up form in the 2- or 1-column view.</span></span>

<span data-ttu-id="c91b5-690">Wenn einer der Zeichenfolgenparameter seinen Grenzwert überschreitet, löst `displayReplyForm` eine Ausnahme aus.</span><span class="sxs-lookup"><span data-stu-id="c91b5-690">If any of the string parameters exceed their limits, `displayReplyForm` throws an exception.</span></span>

<span data-ttu-id="c91b5-p145">Wenn Anlagen im `formData.attachments`-Parameter angegeben werden, versuchen Outlook und Outlook Web App alle Anlagen herunterzuladen und sie an das Antwortformular anzufügen. Wenn Anlagen nicht hinzugefügt werden können, wird in der Formularbenutzeroberfläche ein Fehler ausgegeben. Wenn dies nicht möglich ist, wird keine Fehlermeldung ausgegeben.</span><span class="sxs-lookup"><span data-stu-id="c91b5-p145">When attachments are specified in the `formData.attachments` parameter, Outlook and Outlook Web App attempt to download all attachments and attach them to the reply form. If any attachments fail to be added, an error is shown in the form UI. If this isn't possible, then no error message is thrown.</span></span>

##### <a name="parameters"></a><span data-ttu-id="c91b5-694">Parameter:</span><span class="sxs-lookup"><span data-stu-id="c91b5-694">Parameters:</span></span>

|<span data-ttu-id="c91b5-695">Name</span><span class="sxs-lookup"><span data-stu-id="c91b5-695">Name</span></span>| <span data-ttu-id="c91b5-696">Typ</span><span class="sxs-lookup"><span data-stu-id="c91b5-696">Type</span></span>| <span data-ttu-id="c91b5-697">Beschreibung</span><span class="sxs-lookup"><span data-stu-id="c91b5-697">Description</span></span>|
|---|---|---|
|`formData`| <span data-ttu-id="c91b5-698">Zeichenfolge &#124; Objekt</span><span class="sxs-lookup"><span data-stu-id="c91b5-698">String &#124; Object</span></span>| | <span data-ttu-id="c91b5-p146">Eine Zeichenfolge, die Text- und HTML-Code enthält, die den Hauptteil des Antwortformulars darstellen. Die Zeichenfolge ist auf 32 KB beschränkt.</span><span class="sxs-lookup"><span data-stu-id="c91b5-p146">A string that contains text and HTML and that represents the body of the reply form. The string is limited to 32 KB.</span></span><br/><span data-ttu-id="c91b5-701">**ODER**</span><span class="sxs-lookup"><span data-stu-id="c91b5-701">**OR**</span></span><br/><span data-ttu-id="c91b5-p147">Ein Objekt, das Text- oder Anlagendaten und eine Rückruffunktion enthält. Das Objekt ist wie folgt definiert:</span><span class="sxs-lookup"><span data-stu-id="c91b5-p147">An object that contains body or attachment data and a callback function. The object is defined as follows.</span></span> |
| `formData.htmlBody` | <span data-ttu-id="c91b5-704">Zeichenfolge</span><span class="sxs-lookup"><span data-stu-id="c91b5-704">String</span></span> | <span data-ttu-id="c91b5-705">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="c91b5-705">&lt;optional&gt;</span></span> | <span data-ttu-id="c91b5-p148">Eine Zeichenfolge, die Text- und HTML-Code enthält, die den Hauptteil des Antwortformulars darstellen. Die Zeichenfolge ist auf 32 KB beschränkt.</span><span class="sxs-lookup"><span data-stu-id="c91b5-p148">A string that contains text and HTML and that represents the body of the reply form. The string is limited to 32 KB.</span></span>
| `formData.attachments` | <span data-ttu-id="c91b5-708">Array.&lt;Object&gt;</span><span class="sxs-lookup"><span data-stu-id="c91b5-708">Array.&lt;Object&gt;</span></span> | <span data-ttu-id="c91b5-709">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="c91b5-709">&lt;optional&gt;</span></span> | <span data-ttu-id="c91b5-710">Ein Array mit JSON-Objekten, die Datei- oder Elementanlagen sind.</span><span class="sxs-lookup"><span data-stu-id="c91b5-710">An array of JSON objects that are either file or item attachments.</span></span> |
| `formData.attachments.type` | <span data-ttu-id="c91b5-711">Zeichenfolge</span><span class="sxs-lookup"><span data-stu-id="c91b5-711">String</span></span> | | <span data-ttu-id="c91b5-p149">Gibt den Typ der Anlage an. Muss `file` für eine Dateianlage oder `item` für eine Elementanlage sein.</span><span class="sxs-lookup"><span data-stu-id="c91b5-p149">Indicates the type of attachment. Must be `file` for a file attachment or `item` for an item attachment.</span></span> |
| `formData.attachments.name` | <span data-ttu-id="c91b5-714">Zeichenfolge</span><span class="sxs-lookup"><span data-stu-id="c91b5-714">String</span></span> | | <span data-ttu-id="c91b5-715">Eine Zeichenfolge, die den Namen der Anlage mit bis zu 255 Zeichen enthält.</span><span class="sxs-lookup"><span data-stu-id="c91b5-715">A string that contains the name of the attachment, up to 255 characters in length.</span></span>|
| `formData.attachments.url` | <span data-ttu-id="c91b5-716">Zeichenfolge</span><span class="sxs-lookup"><span data-stu-id="c91b5-716">String</span></span> | | <span data-ttu-id="c91b5-p150">Wird nur verwendet, wenn `type` auf `file` gesetzt ist. Der URI des Speicherorts für die Datei.</span><span class="sxs-lookup"><span data-stu-id="c91b5-p150">Only used if `type` is set to `file`. The URI of the location for the file.</span></span> |
| `formData.attachments.itemId` | <span data-ttu-id="c91b5-719">Zeichenfolge</span><span class="sxs-lookup"><span data-stu-id="c91b5-719">String</span></span> | | <span data-ttu-id="c91b5-p151">Wird nur verwendet, wenn `type` auf `item` gesetzt ist. Die EWS-Element-ID der Anlage. Dies ist eine Zeichenfolge bis zu 100 Zeichen.</span><span class="sxs-lookup"><span data-stu-id="c91b5-p151">Only used if `type` is set to `item`. The EWS item id of the attachment. This is a string up to 100 characters.</span></span> |
| `callback` | <span data-ttu-id="c91b5-723">Funktion</span><span class="sxs-lookup"><span data-stu-id="c91b5-723">function</span></span> | <span data-ttu-id="c91b5-724">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="c91b5-724">&lt;optional&gt;</span></span> | <span data-ttu-id="c91b5-725">Bei Abschluss der Methode wird die im Parameter `callback` übergebene Funktion mit einem einzigen Parameter aufgerufen: `asyncResult`. Dieser Parameter ist ein Objekt des Typs [AsyncResult](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="c91b5-725">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [AsyncResult](/javascript/api/office/office.asyncresult) object.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="c91b5-726">Anforderungen</span><span class="sxs-lookup"><span data-stu-id="c91b5-726">Requirements</span></span>

|<span data-ttu-id="c91b5-727">Anforderung</span><span class="sxs-lookup"><span data-stu-id="c91b5-727">Requirement</span></span>| <span data-ttu-id="c91b5-728">Wert</span><span class="sxs-lookup"><span data-stu-id="c91b5-728">Value</span></span>|
|---|---|
|[<span data-ttu-id="c91b5-729">Mindestversion des Postfachanforderungssatzes</span><span class="sxs-lookup"><span data-stu-id="c91b5-729">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="c91b5-730">1.0</span><span class="sxs-lookup"><span data-stu-id="c91b5-730">1.0</span></span>|
|[<span data-ttu-id="c91b5-731">Mindestberechtigungsstufe</span><span class="sxs-lookup"><span data-stu-id="c91b5-731">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="c91b5-732">ReadItem</span><span class="sxs-lookup"><span data-stu-id="c91b5-732">ReadItem</span></span>|
|[<span data-ttu-id="c91b5-733">Zutreffender Outlook-Modus</span><span class="sxs-lookup"><span data-stu-id="c91b5-733">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="c91b5-734">Lesen</span><span class="sxs-lookup"><span data-stu-id="c91b5-734">Read</span></span>|

##### <a name="examples"></a><span data-ttu-id="c91b5-735">Beispiele</span><span class="sxs-lookup"><span data-stu-id="c91b5-735">Examples</span></span>

<span data-ttu-id="c91b5-736">Im folgenden Code wird eine Zeichenfolge an die `displayReplyForm`-Funktion übergeben.</span><span class="sxs-lookup"><span data-stu-id="c91b5-736">The following code passes a string to the `displayReplyForm` function.</span></span>

```
Office.context.mailbox.item.displayReplyForm('hello there');
Office.context.mailbox.item.displayReplyForm('<b>hello there</b>');
```

<span data-ttu-id="c91b5-737">Antworten mit einem leeren Textkörper.</span><span class="sxs-lookup"><span data-stu-id="c91b5-737">Reply with an empty body.</span></span>

```
Office.context.mailbox.item.displayReplyForm({});
```

<span data-ttu-id="c91b5-738">Antworten nur einem Textkörper.</span><span class="sxs-lookup"><span data-stu-id="c91b5-738">Reply with just a body.</span></span>

```
Office.context.mailbox.item.displayReplyForm(
{
  'htmlBody' : 'hi'
});
```

<span data-ttu-id="c91b5-739">Antworten mit einem Textkörper und einer Dateianlage.</span><span class="sxs-lookup"><span data-stu-id="c91b5-739">Reply with a body and a file attachment.</span></span>

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

<span data-ttu-id="c91b5-740">Antworten mit einem Textkörper und einer Elementanlage.</span><span class="sxs-lookup"><span data-stu-id="c91b5-740">Reply with a body and an item attachment.</span></span>

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

<span data-ttu-id="c91b5-741">Antworten Sie mit einem Textkörper, einer Dateianlage, einer Elementanlage und einem Callback.</span><span class="sxs-lookup"><span data-stu-id="c91b5-741">Reply with a body, file attachment, item attachment, and a callback.</span></span>

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

#### <a name="getentities--entitiesjavascriptapioutlook13officeentities"></a><span data-ttu-id="c91b5-742">getEntities() → {[Entities](/javascript/api/outlook_1_3/office.entities)}</span><span class="sxs-lookup"><span data-stu-id="c91b5-742">getEntities() → {[Entities](/javascript/api/outlook_1_3/office.entities)}</span></span>

<span data-ttu-id="c91b5-743">Ruft das ausgewählte Element Textkörper gefundenen Entitäten ab.</span><span class="sxs-lookup"><span data-stu-id="c91b5-743">Gets the entities found in the selected item's body.</span></span>

> [!NOTE]
> <span data-ttu-id="c91b5-744">Diese Methode wird in Outlook für iOS oder Outlook für Android nicht unterstützt.</span><span class="sxs-lookup"><span data-stu-id="c91b5-744">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

##### <a name="requirements"></a><span data-ttu-id="c91b5-745">Anforderungen</span><span class="sxs-lookup"><span data-stu-id="c91b5-745">Requirements</span></span>

|<span data-ttu-id="c91b5-746">Anforderung</span><span class="sxs-lookup"><span data-stu-id="c91b5-746">Requirement</span></span>| <span data-ttu-id="c91b5-747">Wert</span><span class="sxs-lookup"><span data-stu-id="c91b5-747">Value</span></span>|
|---|---|
|[<span data-ttu-id="c91b5-748">Mindestversion des Postfachanforderungssatzes</span><span class="sxs-lookup"><span data-stu-id="c91b5-748">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="c91b5-749">1.0</span><span class="sxs-lookup"><span data-stu-id="c91b5-749">1.0</span></span>|
|[<span data-ttu-id="c91b5-750">Mindestberechtigungsstufe</span><span class="sxs-lookup"><span data-stu-id="c91b5-750">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="c91b5-751">ReadItem</span><span class="sxs-lookup"><span data-stu-id="c91b5-751">ReadItem</span></span>|
|[<span data-ttu-id="c91b5-752">Zutreffender Outlook-Modus</span><span class="sxs-lookup"><span data-stu-id="c91b5-752">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="c91b5-753">Lesen</span><span class="sxs-lookup"><span data-stu-id="c91b5-753">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="c91b5-754">Gibt zurück:</span><span class="sxs-lookup"><span data-stu-id="c91b5-754">Returns:</span></span>

<span data-ttu-id="c91b5-755">Typ: [Entitäten](/javascript/api/outlook_1_3/office.entities)</span><span class="sxs-lookup"><span data-stu-id="c91b5-755">Type: [Entities](/javascript/api/outlook_1_3/office.entities)</span></span>

##### <a name="example"></a><span data-ttu-id="c91b5-756">Beispiel</span><span class="sxs-lookup"><span data-stu-id="c91b5-756">Example</span></span>

<span data-ttu-id="c91b5-757">Das folgende Beispiel greift auf die Kontakte Entitäten im Textkörper des aktuellen Elements.</span><span class="sxs-lookup"><span data-stu-id="c91b5-757">The following example accesses the contacts entities in the current item's body.</span></span>

```
var contacts = Office.context.mailbox.item.getEntities().contacts;
```

#### <a name="getentitiesbytypeentitytype--nullable-arraystringcontactjavascriptapioutlook13officecontactmeetingsuggestionjavascriptapioutlook13officemeetingsuggestionphonenumberjavascriptapioutlook13officephonenumbertasksuggestionjavascriptapioutlook13officetasksuggestion"></a><span data-ttu-id="c91b5-758">getEntitiesByType(entityType) → (nullable) {Array.<(String|[Contact](/javascript/api/outlook_1_3/office.contact)|[MeetingSuggestion](/javascript/api/outlook_1_3/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook_1_3/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook_1_3/office.tasksuggestion))>}</span><span class="sxs-lookup"><span data-stu-id="c91b5-758">getEntitiesByType(entityType) → (nullable) {Array.<(String|[Contact](/javascript/api/outlook_1_3/office.contact)|[MeetingSuggestion](/javascript/api/outlook_1_3/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook_1_3/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook_1_3/office.tasksuggestion))>}</span></span>

<span data-ttu-id="c91b5-759">Ruft ein Array aller Entitäten des angegebenen Entitätstyps im Hauptteil des ausgewählten Elements gefunden.</span><span class="sxs-lookup"><span data-stu-id="c91b5-759">Gets an array of all the entities of the specified entity type found in the selected item's body.</span></span>

> [!NOTE]
> <span data-ttu-id="c91b5-760">Diese Methode wird in Outlook für iOS oder Outlook für Android nicht unterstützt.</span><span class="sxs-lookup"><span data-stu-id="c91b5-760">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

##### <a name="parameters"></a><span data-ttu-id="c91b5-761">Parameter:</span><span class="sxs-lookup"><span data-stu-id="c91b5-761">Parameters:</span></span>

|<span data-ttu-id="c91b5-762">Name</span><span class="sxs-lookup"><span data-stu-id="c91b5-762">Name</span></span>| <span data-ttu-id="c91b5-763">Typ</span><span class="sxs-lookup"><span data-stu-id="c91b5-763">Type</span></span>| <span data-ttu-id="c91b5-764">Beschreibung</span><span class="sxs-lookup"><span data-stu-id="c91b5-764">Description</span></span>|
|---|---|---|
|`entityType`| [<span data-ttu-id="c91b5-765">Office.MailboxEnums.EntityType</span><span class="sxs-lookup"><span data-stu-id="c91b5-765">Office.MailboxEnums.EntityType</span></span>](/javascript/api/outlook_1_3/office.mailboxenums.entitytype)|<span data-ttu-id="c91b5-766">Einer der Werte der EntityType-Enumeration.</span><span class="sxs-lookup"><span data-stu-id="c91b5-766">One of the EntityType enumeration values.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="c91b5-767">Anforderungen</span><span class="sxs-lookup"><span data-stu-id="c91b5-767">Requirements</span></span>

|<span data-ttu-id="c91b5-768">Anforderung</span><span class="sxs-lookup"><span data-stu-id="c91b5-768">Requirement</span></span>| <span data-ttu-id="c91b5-769">Wert</span><span class="sxs-lookup"><span data-stu-id="c91b5-769">Value</span></span>|
|---|---|
|[<span data-ttu-id="c91b5-770">Mindestversion des Postfachanforderungssatzes</span><span class="sxs-lookup"><span data-stu-id="c91b5-770">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="c91b5-771">1.0</span><span class="sxs-lookup"><span data-stu-id="c91b5-771">1.0</span></span>|
|[<span data-ttu-id="c91b5-772">Mindestberechtigungsstufe</span><span class="sxs-lookup"><span data-stu-id="c91b5-772">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="c91b5-773">Eingeschränkt</span><span class="sxs-lookup"><span data-stu-id="c91b5-773">Restricted</span></span>|
|[<span data-ttu-id="c91b5-774">Zutreffender Outlook-Modus</span><span class="sxs-lookup"><span data-stu-id="c91b5-774">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="c91b5-775">Lesen</span><span class="sxs-lookup"><span data-stu-id="c91b5-775">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="c91b5-776">Gibt zurück:</span><span class="sxs-lookup"><span data-stu-id="c91b5-776">Returns:</span></span>

<span data-ttu-id="c91b5-777">Wenn der in `entityType` übergebene Wert kein gültiges Element der `EntityType`-Enumeration ist, gibt die Methode null zurück.</span><span class="sxs-lookup"><span data-stu-id="c91b5-777">If the value passed in `entityType` is not a valid member of the `EntityType` enumeration, the method returns null.</span></span> <span data-ttu-id="c91b5-778">Wenn keine Entitäten des angegebenen Typs im Textkörper des Elements vorhanden sind, gibt die Methode ein leeres Array zurück.</span><span class="sxs-lookup"><span data-stu-id="c91b5-778">If no entities of the specified type are present in the item's body, the method returns an empty array.</span></span> <span data-ttu-id="c91b5-779">Andernfalls hängt der Typ der Objekte im zurückgegebenen Array vom Typ der Entität ab, die im `entityType`-Parameter angefordert wurde.</span><span class="sxs-lookup"><span data-stu-id="c91b5-779">Otherwise, the type of the objects in the returned array depends on the type of entity requested in the `entityType` parameter.</span></span>

<span data-ttu-id="c91b5-780">Während Sie die minimale Berechtigungsstufe für diese Methode **Restricted** ist, erfordern einige Entitätstypen **ReadItem** für den Zugriff, wie in der folgenden Tabelle angegeben wird.</span><span class="sxs-lookup"><span data-stu-id="c91b5-780">While the minimum permission level to use this method is **Restricted**, some entity types require **ReadItem** to access, as specified in the following table.</span></span>

| <span data-ttu-id="c91b5-781">Wert von `entityType`</span><span class="sxs-lookup"><span data-stu-id="c91b5-781">Value of `entityType`</span></span> | <span data-ttu-id="c91b5-782">Typ der Objekte im zurückgegebenen Array</span><span class="sxs-lookup"><span data-stu-id="c91b5-782">Type of objects in returned array</span></span> | <span data-ttu-id="c91b5-783">Erforderliche Berechtigungsstufe</span><span class="sxs-lookup"><span data-stu-id="c91b5-783">Required Permission Level</span></span> |
| --- | --- | --- |
| `Address` | <span data-ttu-id="c91b5-784">Zeichenfolge</span><span class="sxs-lookup"><span data-stu-id="c91b5-784">String</span></span> | <span data-ttu-id="c91b5-785">**Restricted**</span><span class="sxs-lookup"><span data-stu-id="c91b5-785">**Restricted**</span></span> |
| `Contact` | <span data-ttu-id="c91b5-786">Contact</span><span class="sxs-lookup"><span data-stu-id="c91b5-786">Contact</span></span> | <span data-ttu-id="c91b5-787">**ReadItem**</span><span class="sxs-lookup"><span data-stu-id="c91b5-787">**ReadItem**</span></span> |
| `EmailAddress` | <span data-ttu-id="c91b5-788">Zeichenfolge</span><span class="sxs-lookup"><span data-stu-id="c91b5-788">String</span></span> | <span data-ttu-id="c91b5-789">**ReadItem**</span><span class="sxs-lookup"><span data-stu-id="c91b5-789">**ReadItem**</span></span> |
| `MeetingSuggestion` | <span data-ttu-id="c91b5-790">MeetingSuggestion</span><span class="sxs-lookup"><span data-stu-id="c91b5-790">MeetingSuggestion</span></span> | <span data-ttu-id="c91b5-791">**ReadItem**</span><span class="sxs-lookup"><span data-stu-id="c91b5-791">**ReadItem**</span></span> |
| `PhoneNumber` | <span data-ttu-id="c91b5-792">PhoneNumber</span><span class="sxs-lookup"><span data-stu-id="c91b5-792">PhoneNumber</span></span> | <span data-ttu-id="c91b5-793">**Restricted**</span><span class="sxs-lookup"><span data-stu-id="c91b5-793">**Restricted**</span></span> |
| `TaskSuggestion` | <span data-ttu-id="c91b5-794">TaskSuggestion</span><span class="sxs-lookup"><span data-stu-id="c91b5-794">TaskSuggestion</span></span> | <span data-ttu-id="c91b5-795">**ReadItem**</span><span class="sxs-lookup"><span data-stu-id="c91b5-795">**ReadItem**</span></span> |
| `URL` | <span data-ttu-id="c91b5-796">Zeichenfolge</span><span class="sxs-lookup"><span data-stu-id="c91b5-796">String</span></span> | <span data-ttu-id="c91b5-797">**Restricted**</span><span class="sxs-lookup"><span data-stu-id="c91b5-797">**Restricted**</span></span> |

<span data-ttu-id="c91b5-798">Typ: Array.<(String|[Contact](/javascript/api/outlook_1_3/office.contact)|[MeetingSuggestion](/javascript/api/outlook_1_3/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook_1_3/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook_1_3/office.tasksuggestion))></span><span class="sxs-lookup"><span data-stu-id="c91b5-798">Type: Array.<(String|[Contact](/javascript/api/outlook_1_3/office.contact)|[MeetingSuggestion](/javascript/api/outlook_1_3/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook_1_3/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook_1_3/office.tasksuggestion))></span></span>

##### <a name="example"></a><span data-ttu-id="c91b5-799">Beispiel</span><span class="sxs-lookup"><span data-stu-id="c91b5-799">Example</span></span>

<span data-ttu-id="c91b5-800">Im folgenden Beispiel wird veranschaulicht, wie auf ein Array von Zeichenfolgen, die Postadressen im Textkörper des aktuellen Elements darstellen.</span><span class="sxs-lookup"><span data-stu-id="c91b5-800">The following example shows how to access an array of strings that represent postal addresses in the current item's body.</span></span>

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

#### <a name="getfilteredentitiesbynamename--nullable-arraystringcontactjavascriptapioutlook13officecontactmeetingsuggestionjavascriptapioutlook13officemeetingsuggestionphonenumberjavascriptapioutlook13officephonenumbertasksuggestionjavascriptapioutlook13officetasksuggestion"></a><span data-ttu-id="c91b5-801">getFilteredEntitiesByName(name) → (nullable) {Array.<(String|[Contact](/javascript/api/outlook_1_3/office.contact)|[MeetingSuggestion](/javascript/api/outlook_1_3/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook_1_3/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook_1_3/office.tasksuggestion))>}</span><span class="sxs-lookup"><span data-stu-id="c91b5-801">getFilteredEntitiesByName(name) → (nullable) {Array.<(String|[Contact](/javascript/api/outlook_1_3/office.contact)|[MeetingSuggestion](/javascript/api/outlook_1_3/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook_1_3/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook_1_3/office.tasksuggestion))>}</span></span>

<span data-ttu-id="c91b5-802">Gibt bekannte Entitäten im ausgewählten Element zurück, die den in der XML-Manifestdatei definierten benannten Filter übergeben.</span><span class="sxs-lookup"><span data-stu-id="c91b5-802">Returns well-known entities in the selected item that pass the named filter defined in the manifest XML file.</span></span>

> [!NOTE]
> <span data-ttu-id="c91b5-803">Diese Methode wird in Outlook für iOS oder Outlook für Android nicht unterstützt.</span><span class="sxs-lookup"><span data-stu-id="c91b5-803">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="c91b5-804">Die `getFilteredEntitiesByName`-Methode gibt die Entitäten zurück, die dem im [ItemHasKnownEntity](/javascript/office/manifest/rule#itemhasknownentity-rule)-Regelelement der XML-Manifestdatei definierten regulären Ausdruck mit dem angegebenen `FilterName`-Elementwert entsprechen.</span><span class="sxs-lookup"><span data-stu-id="c91b5-804">The `getFilteredEntitiesByName` method returns the entities that match the regular expression defined in the [ItemHasKnownEntity](/javascript/office/manifest/rule#itemhasknownentity-rule) rule element in the manifest XML file with the specified `FilterName` element value.</span></span>

##### <a name="parameters"></a><span data-ttu-id="c91b5-805">Parameter:</span><span class="sxs-lookup"><span data-stu-id="c91b5-805">Parameters:</span></span>

|<span data-ttu-id="c91b5-806">Name</span><span class="sxs-lookup"><span data-stu-id="c91b5-806">Name</span></span>| <span data-ttu-id="c91b5-807">Typ</span><span class="sxs-lookup"><span data-stu-id="c91b5-807">Type</span></span>| <span data-ttu-id="c91b5-808">Beschreibung</span><span class="sxs-lookup"><span data-stu-id="c91b5-808">Description</span></span>|
|---|---|---|
|`name`| <span data-ttu-id="c91b5-809">Zeichenfolge</span><span class="sxs-lookup"><span data-stu-id="c91b5-809">String</span></span>|<span data-ttu-id="c91b5-810">Der Name des `ItemHasKnownEntity`-Regelelements, das den entsprechenden Filter definiert.</span><span class="sxs-lookup"><span data-stu-id="c91b5-810">The name of the `ItemHasKnownEntity` rule element that defines the filter to match.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="c91b5-811">Anforderungen</span><span class="sxs-lookup"><span data-stu-id="c91b5-811">Requirements</span></span>

|<span data-ttu-id="c91b5-812">Anforderung</span><span class="sxs-lookup"><span data-stu-id="c91b5-812">Requirement</span></span>| <span data-ttu-id="c91b5-813">Wert</span><span class="sxs-lookup"><span data-stu-id="c91b5-813">Value</span></span>|
|---|---|
|[<span data-ttu-id="c91b5-814">Mindestversion des Postfachanforderungssatzes</span><span class="sxs-lookup"><span data-stu-id="c91b5-814">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="c91b5-815">1.0</span><span class="sxs-lookup"><span data-stu-id="c91b5-815">1.0</span></span>|
|[<span data-ttu-id="c91b5-816">Mindestberechtigungsstufe</span><span class="sxs-lookup"><span data-stu-id="c91b5-816">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="c91b5-817">ReadItem</span><span class="sxs-lookup"><span data-stu-id="c91b5-817">ReadItem</span></span>|
|[<span data-ttu-id="c91b5-818">Zutreffender Outlook-Modus</span><span class="sxs-lookup"><span data-stu-id="c91b5-818">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="c91b5-819">Lesen</span><span class="sxs-lookup"><span data-stu-id="c91b5-819">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="c91b5-820">Gibt zurück:</span><span class="sxs-lookup"><span data-stu-id="c91b5-820">Returns:</span></span>

<span data-ttu-id="c91b5-p153">Befindet sich kein `ItemHasKnownEntity`-Element im Manifest mit einem `FilterName`-Elementwert, der dem `name`-Parameter entspricht, gibt die Methode `null` zurück. Wenn der `name`-Parameter einem `ItemHasKnownEntity`-Element im Manifest entspricht, es aber keine entsprechenden Entitäten im aktuellen Element gibt, gibt die Methode ein leeres Array zurück.</span><span class="sxs-lookup"><span data-stu-id="c91b5-p153">If there is no `ItemHasKnownEntity` element in the manifest with a `FilterName` element value that matches the `name` parameter, the method returns `null`. If the `name` parameter does match an `ItemHasKnownEntity` element in the manifest, but there are no entities in the current item that match, the method return an empty array.</span></span>

<span data-ttu-id="c91b5-823">Typ: Array.<(String|[Contact](/javascript/api/outlook_1_3/office.contact)|[MeetingSuggestion](/javascript/api/outlook_1_3/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook_1_3/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook_1_3/office.tasksuggestion))></span><span class="sxs-lookup"><span data-stu-id="c91b5-823">Type: Array.<(String|[Contact](/javascript/api/outlook_1_3/office.contact)|[MeetingSuggestion](/javascript/api/outlook_1_3/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook_1_3/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook_1_3/office.tasksuggestion))></span></span>

#### <a name="getregexmatches--object"></a><span data-ttu-id="c91b5-824">getRegExMatches() → {Object}</span><span class="sxs-lookup"><span data-stu-id="c91b5-824">getRegExMatches() → {Object}</span></span>

<span data-ttu-id="c91b5-825">Gibt Zeichenfolgenwerte im ausgewählten Element zurück, die den in der XML-Manifestdatei definierten regulären Ausdrücken entsprechen.</span><span class="sxs-lookup"><span data-stu-id="c91b5-825">Returns string values in the selected item that match the regular expressions defined in the manifest XML file.</span></span>

> [!NOTE]
> <span data-ttu-id="c91b5-826">Diese Methode wird in Outlook für iOS oder Outlook für Android nicht unterstützt.</span><span class="sxs-lookup"><span data-stu-id="c91b5-826">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="c91b5-p154">Die `getRegExMatches`-Methode gibt die Zeichenfolgen zurück, die dem im `ItemHasRegularExpressionMatch`- oder `ItemHasKnownEntity`-Regelelement der XML-Manifestdatei definierten regulären Ausdruck entsprechen. Bei einer `ItemHasRegularExpressionMatch`-Regel muss eine entsprechende Zeichenfolge in der Eigenschaft des Elements vorhanden sein, das von dieser Regel angegeben wird. Der einfache `PropertyName`-Typ definiert die unterstützten Eigenschaften.</span><span class="sxs-lookup"><span data-stu-id="c91b5-p154">The `getRegExMatches` method returns the strings that match the regular expression defined in each `ItemHasRegularExpressionMatch` or `ItemHasKnownEntity` rule element in the manifest XML file. For an `ItemHasRegularExpressionMatch` rule, a matching string has to occur in the property of the item that is specified by that rule. The `PropertyName` simple type defines the supported properties.</span></span>

<span data-ttu-id="c91b5-830">Nehmen Sie z. B. an, dass ein Add-In-Manifest das folgende `Rule`-Element aufweist:</span><span class="sxs-lookup"><span data-stu-id="c91b5-830">For example, consider an add-in manifest has the following `Rule` element:</span></span>

```
<Rule xsi:type="RuleCollection" Mode="And">
  <Rule xsi:type="ItemIs" FormType="Read" ItemType="Message" />
  <Rule xsi:type="RuleCollection" Mode="Or">
    <Rule xsi:type="ItemHasRegularExpressionMatch" RegExName="fruits" RegExValue="apple|banana|coconut" PropertyName="BodyAsPlaintext" IgnoreCase="true" />
    <Rule xsi:type="ItemHasRegularExpressionMatch" RegExName="veggies" RegExValue="tomato|onion|spinach|broccoli" PropertyName="BodyAsPlaintext" IgnoreCase="true" />
  </Rule>
</Rule>
```

<span data-ttu-id="c91b5-831">Das von `getRegExMatches` zurückgegebene Objekt hätte zwei Eigenschaften: `fruits` und `veggies`.</span><span class="sxs-lookup"><span data-stu-id="c91b5-831">The object returned from `getRegExMatches` would have two properties: `fruits` and `veggies`.</span></span>

```
{
  'fruits': ['apple','banana','Banana','coconut'],
  'veggies': ['tomato','onion','spinach','broccoli']
}
```

<span data-ttu-id="c91b5-p155">Wenn Sie eine `ItemHasRegularExpressionMatch`-Regel für die Textkörpereigenschaft eines Elements festlegen, sollte der reguläre Ausdruck den Textkörper weiter filtern und nicht versuchen, den gesamten Textkörper des Elements zurückzugeben. Wenn der gesamte Textkörper eines Elements mit einem regulären Ausdruck wie `.*` abgerufen wird, werden nicht immer die gewünschten Ergebnisse erzielt. Verwenden Sie stattdessen die [`Body.getAsync`](/javascript/api/outlook_1_3/office.body#getasync-coerciontype--options--callback-)-Methode, um den gesamten Textkörper abzurufen.</span><span class="sxs-lookup"><span data-stu-id="c91b5-p155">If you specify an `ItemHasRegularExpressionMatch` rule on the body property of an item, the regular expression should further filter the body and should not attempt to return the entire body of the item. Using a regular expression such as `.*` to obtain the entire body of an item does not always return the expected results. Instead, use the [`Body.getAsync`](/javascript/api/outlook_1_3/office.body#getasync-coerciontype--options--callback-) method to retrieve the entire body.</span></span>

##### <a name="requirements"></a><span data-ttu-id="c91b5-835">Anforderungen</span><span class="sxs-lookup"><span data-stu-id="c91b5-835">Requirements</span></span>

|<span data-ttu-id="c91b5-836">Anforderung</span><span class="sxs-lookup"><span data-stu-id="c91b5-836">Requirement</span></span>| <span data-ttu-id="c91b5-837">Wert</span><span class="sxs-lookup"><span data-stu-id="c91b5-837">Value</span></span>|
|---|---|
|[<span data-ttu-id="c91b5-838">Mindestversion des Postfachanforderungssatzes</span><span class="sxs-lookup"><span data-stu-id="c91b5-838">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="c91b5-839">1.0</span><span class="sxs-lookup"><span data-stu-id="c91b5-839">1.0</span></span>|
|[<span data-ttu-id="c91b5-840">Mindestberechtigungsstufe</span><span class="sxs-lookup"><span data-stu-id="c91b5-840">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="c91b5-841">ReadItem</span><span class="sxs-lookup"><span data-stu-id="c91b5-841">ReadItem</span></span>|
|[<span data-ttu-id="c91b5-842">Zutreffender Outlook-Modus</span><span class="sxs-lookup"><span data-stu-id="c91b5-842">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="c91b5-843">Lesen</span><span class="sxs-lookup"><span data-stu-id="c91b5-843">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="c91b5-844">Gibt zurück:</span><span class="sxs-lookup"><span data-stu-id="c91b5-844">Returns:</span></span>

<span data-ttu-id="c91b5-p156">Ein Objekt mit Arrays aus Zeichenfolgen, die den in der XML-Manifestdatei definierten regulären Ausdrücken entsprechen. Der Name der einzelnen Arrays ist gleich dem entsprechenden Wert des `RegExName`-Attributs der entsprechenden `ItemHasRegularExpressionMatch`-Regel oder des `FilterName`-Attributs der entsprechenden `ItemHasKnownEntity`-Regel.</span><span class="sxs-lookup"><span data-stu-id="c91b5-p156">An object that contains arrays of strings that match the regular expressions defined in the manifest XML file. The name of each array is equal to the corresponding value of the `RegExName` attribute of the matching `ItemHasRegularExpressionMatch` rule or the `FilterName` attribute of the matching `ItemHasKnownEntity` rule.</span></span>

<dl class="param-type"><span data-ttu-id="c91b5-847">

<dt>
Typ</dt>


</span><span class="sxs-lookup"><span data-stu-id="c91b5-847">

<dt>Type</dt>

</span></span><dd><span data-ttu-id="c91b5-848">Object</span><span class="sxs-lookup"><span data-stu-id="c91b5-848">Object</span></span></dd>

</dl>

##### <a name="example"></a><span data-ttu-id="c91b5-849">Beispiel</span><span class="sxs-lookup"><span data-stu-id="c91b5-849">Example</span></span>

<span data-ttu-id="c91b5-850">Das folgende Beispiel zeigt, wie Sie auf das Array von Übereinstimmungen für die regulären <rule>-Ausdruckselemente `fruits` und `veggies` zugreifen, die im Manifest angegeben sind.</rule></span><span class="sxs-lookup"><span data-stu-id="c91b5-850">The following example shows how to access the array of matches for the regular expression <rule>elements `fruits` and `veggies`, which are specified in the manifest.</rule></span></span>

```
var allMatches = Office.context.mailbox.item.getRegExMatches();
var fruits = allMatches.fruits;
var veges = allMatches.veggies;
```

#### <a name="getregexmatchesbynamename--nullable-array-string-"></a><span data-ttu-id="c91b5-851">getRegExMatchesByName(name) → (Nullwerte zulassen) {Array. < Zeichenfolge >}</span><span class="sxs-lookup"><span data-stu-id="c91b5-851">getRegExMatchesByName(name) → (nullable) {Array.< String >}</span></span>

<span data-ttu-id="c91b5-852">Gibt Zeichenfolgenwerte im ausgewählten Element zurück, die dem in der XML-Manifestdatei definierten benannten regulären Ausdruck entsprechen.</span><span class="sxs-lookup"><span data-stu-id="c91b5-852">Returns string values in the selected item that match the named regular expression defined in the manifest XML file.</span></span>

> [!NOTE]
> <span data-ttu-id="c91b5-853">Diese Methode wird in Outlook für iOS oder Outlook für Android nicht unterstützt.</span><span class="sxs-lookup"><span data-stu-id="c91b5-853">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="c91b5-854">Die `getRegExMatchesByName`-Methode gibt die Zeichenfolgen zurück, die dem im `ItemHasRegularExpressionMatch`-Regelelement der XML-Manifestdatei definierten regulären Ausdruck mit dem angegebenen `RegExName`-Elementwert entsprechen.</span><span class="sxs-lookup"><span data-stu-id="c91b5-854">The `getRegExMatchesByName` method returns the strings that match the regular expression defined in the `ItemHasRegularExpressionMatch` rule element in the manifest XML file with the specified `RegExName` element value.</span></span>

<span data-ttu-id="c91b5-p157">Wenn Sie eine `ItemHasRegularExpressionMatch`-Regel für die Textkörpereigenschaft eines Elements festlegen, sollte der reguläre Ausdruck den Textkörper weiter filtern und nicht versuchen, den gesamten Textkörper des Elements zurückzugeben. Wenn der gesamte Textkörper eines Elements mit einem regulären Ausdruck wie `.*` abgerufen wird, werden nicht immer die gewünschten Ergebnisse erzielt.</span><span class="sxs-lookup"><span data-stu-id="c91b5-p157">If you specify an `ItemHasRegularExpressionMatch` rule on the body property of an item, the regular expression should further filter the body and should not attempt to return the entire body of the item. Using a regular expression such as `.*` to obtain the entire body of an item does not always return the expected results.</span></span>

##### <a name="parameters"></a><span data-ttu-id="c91b5-857">Parameter:</span><span class="sxs-lookup"><span data-stu-id="c91b5-857">Parameters:</span></span>

|<span data-ttu-id="c91b5-858">Name</span><span class="sxs-lookup"><span data-stu-id="c91b5-858">Name</span></span>| <span data-ttu-id="c91b5-859">Typ</span><span class="sxs-lookup"><span data-stu-id="c91b5-859">Type</span></span>| <span data-ttu-id="c91b5-860">Beschreibung</span><span class="sxs-lookup"><span data-stu-id="c91b5-860">Description</span></span>|
|---|---|---|
|`name`| <span data-ttu-id="c91b5-861">Zeichenfolge</span><span class="sxs-lookup"><span data-stu-id="c91b5-861">String</span></span>|<span data-ttu-id="c91b5-862">Der Name des `ItemHasRegularExpressionMatch`-Regelelements, das den entsprechenden Filter definiert.</span><span class="sxs-lookup"><span data-stu-id="c91b5-862">The name of the `ItemHasRegularExpressionMatch` rule element that defines the filter to match.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="c91b5-863">Anforderungen</span><span class="sxs-lookup"><span data-stu-id="c91b5-863">Requirements</span></span>

|<span data-ttu-id="c91b5-864">Anforderung</span><span class="sxs-lookup"><span data-stu-id="c91b5-864">Requirement</span></span>| <span data-ttu-id="c91b5-865">Wert</span><span class="sxs-lookup"><span data-stu-id="c91b5-865">Value</span></span>|
|---|---|
|[<span data-ttu-id="c91b5-866">Mindestversion des Postfachanforderungssatzes</span><span class="sxs-lookup"><span data-stu-id="c91b5-866">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="c91b5-867">1.0</span><span class="sxs-lookup"><span data-stu-id="c91b5-867">1.0</span></span>|
|[<span data-ttu-id="c91b5-868">Mindestberechtigungsstufe</span><span class="sxs-lookup"><span data-stu-id="c91b5-868">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="c91b5-869">ReadItem</span><span class="sxs-lookup"><span data-stu-id="c91b5-869">ReadItem</span></span>|
|[<span data-ttu-id="c91b5-870">Zutreffender Outlook-Modus</span><span class="sxs-lookup"><span data-stu-id="c91b5-870">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="c91b5-871">Lesen</span><span class="sxs-lookup"><span data-stu-id="c91b5-871">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="c91b5-872">Gibt zurück:</span><span class="sxs-lookup"><span data-stu-id="c91b5-872">Returns:</span></span>

<span data-ttu-id="c91b5-873">Ein Array mit den Zeichenfolgen, die dem in der XML-Manifestdatei definierten regulären Ausdruck entsprechen.</span><span class="sxs-lookup"><span data-stu-id="c91b5-873">An array that contains the strings that match the regular expression defined in the manifest XML file.</span></span>

<dl class="param-type"><span data-ttu-id="c91b5-874">

<dt>Typ</dt>

</span><span class="sxs-lookup"><span data-stu-id="c91b5-874">

<dt>Type</dt>

</span></span><dd><span data-ttu-id="c91b5-875">Array. < Zeichenfolge ></span><span class="sxs-lookup"><span data-stu-id="c91b5-875">Array.< String ></span></span></dd>

</dl>

##### <a name="example"></a><span data-ttu-id="c91b5-876">Beispiel</span><span class="sxs-lookup"><span data-stu-id="c91b5-876">Example</span></span>

```
var fruits = Office.context.mailbox.item.getRegExMatchesByName("fruits");
var veggies = Office.context.mailbox.item.getRegExMatchesByName("veggies");
```

####  <a name="getselecteddataasynccoerciontype-options-callback--string"></a><span data-ttu-id="c91b5-877">getSelectedDataAsync(coercionType, [options], callback) → {String}</span><span class="sxs-lookup"><span data-stu-id="c91b5-877">getSelectedDataAsync(coercionType, [options], callback) → {String}</span></span>

<span data-ttu-id="c91b5-878">Gibt asynchron ausgewählte Daten aus dem Betreff oder Textkörper einer Nachricht zurück.</span><span class="sxs-lookup"><span data-stu-id="c91b5-878">Asynchronously returns selected data from the subject or body of a message.</span></span>

<span data-ttu-id="c91b5-p158">Wenn keine Auswahl vorhanden ist, aber der Cursor sich im Nachrichtentext oder Betreff befindet, gibt die Methode für die ausgewählten Daten NULL zurück. Wenn ein anderes Feld als der Textkörper oder Betreff ausgewählt ist, gibt die Methode den `InvalidSelection`-Fehler zurück.</span><span class="sxs-lookup"><span data-stu-id="c91b5-p158">If there is no selection but the cursor is in the body or subject, the method returns null for the selected data. If a field other than the body or subject is selected, the method returns the `InvalidSelection` error.</span></span>

##### <a name="parameters"></a><span data-ttu-id="c91b5-881">Parameter:</span><span class="sxs-lookup"><span data-stu-id="c91b5-881">Parameters:</span></span>

|<span data-ttu-id="c91b5-882">Name</span><span class="sxs-lookup"><span data-stu-id="c91b5-882">Name</span></span>| <span data-ttu-id="c91b5-883">Typ</span><span class="sxs-lookup"><span data-stu-id="c91b5-883">Type</span></span>| <span data-ttu-id="c91b5-884">Attribute</span><span class="sxs-lookup"><span data-stu-id="c91b5-884">Attributes</span></span>| <span data-ttu-id="c91b5-885">Beschreibung</span><span class="sxs-lookup"><span data-stu-id="c91b5-885">Description</span></span>|
|---|---|---|---|
|`coercionType`| [<span data-ttu-id="c91b5-886">Office.CoercionType</span><span class="sxs-lookup"><span data-stu-id="c91b5-886">Office.CoercionType</span></span>](office.md#coerciontype-string)||<span data-ttu-id="c91b5-p159">Fordert ein Format für die Daten an. Wenn es sich um Texthandelt, gibt die Methode den unformatierten Text als Zeichenfolge zurück und entfernt eventuell vorhandene HTML-Tags. Wenn es sich um HTML handelt, gibt die Methode den ausgewählten Text zurück, entweder als unformatierten Text oder als HTML.</span><span class="sxs-lookup"><span data-stu-id="c91b5-p159">Requests a format for the data. If Text, the method returns the plain text as a string , removing any HTML tags present. If HTML, the method returns the selected text, whether it is plaintext or HTML.</span></span>|
|`options`| <span data-ttu-id="c91b5-890">Object</span><span class="sxs-lookup"><span data-stu-id="c91b5-890">Object</span></span>| <span data-ttu-id="c91b5-891">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="c91b5-891">&lt;optional&gt;</span></span>|<span data-ttu-id="c91b5-892">Ein Objektliteral, das mindestens eine der folgenden Eigenschaften enthält.</span><span class="sxs-lookup"><span data-stu-id="c91b5-892">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`| <span data-ttu-id="c91b5-893">Object</span><span class="sxs-lookup"><span data-stu-id="c91b5-893">Object</span></span>| <span data-ttu-id="c91b5-894">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="c91b5-894">&lt;optional&gt;</span></span>|<span data-ttu-id="c91b5-895">Entwickler können ein Objekt bereitstellen, auf das sie in der Callbackmethode zugreifen möchten.</span><span class="sxs-lookup"><span data-stu-id="c91b5-895">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`| <span data-ttu-id="c91b5-896">Funktion</span><span class="sxs-lookup"><span data-stu-id="c91b5-896">function</span></span>||<span data-ttu-id="c91b5-897">Wenn die Methode abgeschlossen wird, wird die Funktion , die im `callback`-Parameter übergeben wird, mit einem einzelnen Parameter (`asyncResult`) aufgerufen, der ein [`AsyncResult`](/javascript/api/office/office.asyncresult)-Objekt darstellt.</span><span class="sxs-lookup"><span data-stu-id="c91b5-897">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span><br/><br/><span data-ttu-id="c91b5-898">Rufen Sie für den Zugriff auf die ausgewählten Daten aus der Rückrufmethode `asyncResult.value.data` auf.</span><span class="sxs-lookup"><span data-stu-id="c91b5-898">To access the selected data from the callback method, call `asyncResult.value.data`.</span></span> <span data-ttu-id="c91b5-899">Rufen Sie die Source-Eigenschaft für den Zugriff, die die Auswahl stammen, `asyncResult.value.sourceProperty`, die entweder sein `body` oder `subject`.</span><span class="sxs-lookup"><span data-stu-id="c91b5-899">To access the source property that the selection comes from, call `asyncResult.value.sourceProperty`, which will be either `body` or `subject`.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="c91b5-900">Anforderungen</span><span class="sxs-lookup"><span data-stu-id="c91b5-900">Requirements</span></span>

|<span data-ttu-id="c91b5-901">Anforderung</span><span class="sxs-lookup"><span data-stu-id="c91b5-901">Requirement</span></span>| <span data-ttu-id="c91b5-902">Wert</span><span class="sxs-lookup"><span data-stu-id="c91b5-902">Value</span></span>|
|---|---|
|[<span data-ttu-id="c91b5-903">Mindestversion des Postfachanforderungssatzes</span><span class="sxs-lookup"><span data-stu-id="c91b5-903">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="c91b5-904">1.2</span><span class="sxs-lookup"><span data-stu-id="c91b5-904">1.2</span></span>|
|[<span data-ttu-id="c91b5-905">Mindestberechtigungsstufe</span><span class="sxs-lookup"><span data-stu-id="c91b5-905">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="c91b5-906">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="c91b5-906">ReadWriteItem</span></span>|
|[<span data-ttu-id="c91b5-907">Zutreffender Outlook-Modus</span><span class="sxs-lookup"><span data-stu-id="c91b5-907">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="c91b5-908">Verfassen</span><span class="sxs-lookup"><span data-stu-id="c91b5-908">Compose</span></span>|

##### <a name="returns"></a><span data-ttu-id="c91b5-909">Gibt zurück:</span><span class="sxs-lookup"><span data-stu-id="c91b5-909">Returns:</span></span>

<span data-ttu-id="c91b5-910">Die ausgewählten Daten als Zeichenfolge mit dem durch `coercionType` bestimmten Format.</span><span class="sxs-lookup"><span data-stu-id="c91b5-910">The selected data as a string with format determined by `coercionType`.</span></span>

<dl class="param-type"><span data-ttu-id="c91b5-911">

<dt>
Typ</dt>


</span><span class="sxs-lookup"><span data-stu-id="c91b5-911">

<dt>Type</dt>

</span></span><dd><span data-ttu-id="c91b5-912">Zeichenfolge</span><span class="sxs-lookup"><span data-stu-id="c91b5-912">String</span></span></dd>

</dl>

##### <a name="example"></a><span data-ttu-id="c91b5-913">Beispiel</span><span class="sxs-lookup"><span data-stu-id="c91b5-913">Example</span></span>

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

####  <a name="loadcustompropertiesasynccallback-usercontext"></a><span data-ttu-id="c91b5-914">loadCustomPropertiesAsync(callback, [userContext])</span><span class="sxs-lookup"><span data-stu-id="c91b5-914">loadCustomPropertiesAsync(callback, [userContext])</span></span>

<span data-ttu-id="c91b5-915">Lädt asynchron benutzerdefinierte Eigenschaften für dieses Add-In für das ausgewählte Element.</span><span class="sxs-lookup"><span data-stu-id="c91b5-915">Asynchronously loads custom properties for this add-in on the selected item.</span></span>

<span data-ttu-id="c91b5-p161">Benutzerdefinierte Eigenschaften werden als Schlüssel-/Wert-Paare pro App und pro Element gespeichert. Diese Methode gibt ein `CustomProperties`-Objekt im Rückruf zurück, das Methoden für den Zugriff auf die benutzerdefinierten Eigenschaften für das aktuelle Element und das aktuelle Add-In bietet. Benutzerdefinierte Eigenschaften sind für das Element nicht verschlüsselt und sollten darum nicht als sicherer Speicher verwendet werden.</span><span class="sxs-lookup"><span data-stu-id="c91b5-p161">Custom properties are stored as key/value pairs on a per-app, per-item basis. This method returns a `CustomProperties` object in the callback, which provides methods to access the custom properties specific to the current item and the current add-in. Custom properties are not encrypted on the item, so this should not be used as secure storage.</span></span>

##### <a name="parameters"></a><span data-ttu-id="c91b5-919">Parameter:</span><span class="sxs-lookup"><span data-stu-id="c91b5-919">Parameters:</span></span>

|<span data-ttu-id="c91b5-920">Name</span><span class="sxs-lookup"><span data-stu-id="c91b5-920">Name</span></span>| <span data-ttu-id="c91b5-921">Typ</span><span class="sxs-lookup"><span data-stu-id="c91b5-921">Type</span></span>| <span data-ttu-id="c91b5-922">Attribute</span><span class="sxs-lookup"><span data-stu-id="c91b5-922">Attributes</span></span>| <span data-ttu-id="c91b5-923">Beschreibung</span><span class="sxs-lookup"><span data-stu-id="c91b5-923">Description</span></span>|
|---|---|---|---|
|`callback`| <span data-ttu-id="c91b5-924">Funktion</span><span class="sxs-lookup"><span data-stu-id="c91b5-924">function</span></span>||<span data-ttu-id="c91b5-925">Wenn die Methode abgeschlossen wird, wird die Funktion , die im `callback`-Parameter übergeben wird, mit einem einzelnen Parameter (`asyncResult`) aufgerufen, der ein [`AsyncResult`](/javascript/api/office/office.asyncresult)-Objekt darstellt.</span><span class="sxs-lookup"><span data-stu-id="c91b5-925">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span><br/><br/><span data-ttu-id="c91b5-926">Die benutzerdefinierten Eigenschaften werden als [`CustomProperties`](/javascript/api/outlook_1_3/office.customproperties)-Objekt in der `asyncResult.value`-Eigenschaft bereitgestellt.</span><span class="sxs-lookup"><span data-stu-id="c91b5-926">The custom properties are provided as a [`CustomProperties`](/javascript/api/outlook_1_3/office.customproperties) object in the `asyncResult.value` property.</span></span> <span data-ttu-id="c91b5-927">Dieses Objekt kann zum Abrufen, festlegen und Entfernen benutzerdefinierter Eigenschaften aus dem Element und speichern Sie Änderungen an den benutzerdefinierten Eigenschaftensatz zurück an den Server verwendet werden.</span><span class="sxs-lookup"><span data-stu-id="c91b5-927">This object can be used to get, set, and remove custom properties from the item and save changes to the custom property set back to the server.</span></span>|
|`userContext`| <span data-ttu-id="c91b5-928">Objekt</span><span class="sxs-lookup"><span data-stu-id="c91b5-928">Object</span></span>| <span data-ttu-id="c91b5-929">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="c91b5-929">&lt;optional&gt;</span></span>|<span data-ttu-id="c91b5-930">Entwickler können ein beliebiges Objekt bereitstellen, den sie in der Rückruffunktion zugreifen möchten.</span><span class="sxs-lookup"><span data-stu-id="c91b5-930">Developers can provide any object they wish to access in the callback function.</span></span> <span data-ttu-id="c91b5-931">Dieses Objekt zugegriffen werden kann, indem die `asyncResult.asyncContext` -Eigenschaft in der Callback-Funktion.</span><span class="sxs-lookup"><span data-stu-id="c91b5-931">This object can be accessed by the `asyncResult.asyncContext` property in the callback function.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="c91b5-932">Anforderungen</span><span class="sxs-lookup"><span data-stu-id="c91b5-932">Requirements</span></span>

|<span data-ttu-id="c91b5-933">Anforderung</span><span class="sxs-lookup"><span data-stu-id="c91b5-933">Requirement</span></span>| <span data-ttu-id="c91b5-934">Wert</span><span class="sxs-lookup"><span data-stu-id="c91b5-934">Value</span></span>|
|---|---|
|[<span data-ttu-id="c91b5-935">Mindestversion des Postfachanforderungssatzes</span><span class="sxs-lookup"><span data-stu-id="c91b5-935">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="c91b5-936">1.0</span><span class="sxs-lookup"><span data-stu-id="c91b5-936">1.0</span></span>|
|[<span data-ttu-id="c91b5-937">Mindestberechtigungsstufe</span><span class="sxs-lookup"><span data-stu-id="c91b5-937">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="c91b5-938">ReadItem</span><span class="sxs-lookup"><span data-stu-id="c91b5-938">ReadItem</span></span>|
|[<span data-ttu-id="c91b5-939">Zutreffender Outlook-Modus</span><span class="sxs-lookup"><span data-stu-id="c91b5-939">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="c91b5-940">Verfassen oder Lesen</span><span class="sxs-lookup"><span data-stu-id="c91b5-940">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="c91b5-941">Beispiel</span><span class="sxs-lookup"><span data-stu-id="c91b5-941">Example</span></span>

<span data-ttu-id="c91b5-p164">Das folgende Codebeispiel veranschaulicht die Verwendung der `loadCustomPropertiesAsync`-Methode zum asynchronen Laden der benutzerdefinierten Eigenschaften, die für das aktuelle Element spezifisch sind. Des Weiteren veranschaulicht das Beispiel die Verwendung der `CustomProperties.saveAsync`-Methode zum Speichern der Eigenschaften auf dem Server. Nach dem Laden der benutzerdefinierten Eigenschaften wird in dem Codebeispiel die `CustomProperties.get`-Methode dazu verwendet, die benutzerdefinierte `myProp`-Eigenschaft zu lesen. Die `CustomProperties.set`-Methode wird dazu verwendet, die benutzerdefinierte `otherProp`-Eigenschaft zu schreiben. Schließlich wird die `saveAsync`-Methode aufgerufen, um die benutzerdefinierten Eigenschaften zu speichern.</span><span class="sxs-lookup"><span data-stu-id="c91b5-p164">The following code example shows how to use the `loadCustomPropertiesAsync` method to asynchronously load custom properties that are specific to the current item. The example also shows how to use the `CustomProperties.saveAsync` method to save these properties back to the server. After loading the custom properties, the code sample uses the `CustomProperties.get` method to read the custom property `myProp`, the `CustomProperties.set` method to write the custom property `otherProp`, and then finally calls the `saveAsync` method to save the custom properties.</span></span>

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

####  <a name="removeattachmentasyncattachmentid-options-callback"></a><span data-ttu-id="c91b5-945">removeAttachmentAsync(attachmentId, [options], [callback])</span><span class="sxs-lookup"><span data-stu-id="c91b5-945">removeAttachmentAsync(attachmentId, [options], [callback])</span></span>

<span data-ttu-id="c91b5-946">Entfernt eine Anlage aus einer Nachricht oder einem Termin.</span><span class="sxs-lookup"><span data-stu-id="c91b5-946">Removes an attachment from a message or appointment.</span></span>

<span data-ttu-id="c91b5-p165">Die `removeAttachmentAsync`-Methode entfernt die Anlage mit dem angegebenen Bezeichner aus dem Element. Als bewährte Vorgehensweise sollten Sie den Anlagenbezeichner nur dann zum Entfernen einer Anlage verwenden, wenn die gleiche Mail-App die Anlage in der gleichen Sitzung hinzugefügt hat. In Outlook Web App und OWA für Geräte ist der Anlagenbezeichner nur innerhalb der gleichen Sitzung gültig. Eine Sitzung ist abgeschlossen, wenn der Benutzer die App schließt, oder wenn der Benutzer in einem eingebetteten Formular mit dem Verfassen beginnt und den Vorgang dann in einem separaten Fenster fortsetzt.</span><span class="sxs-lookup"><span data-stu-id="c91b5-p165">The `removeAttachmentAsync` method removes the attachment with the specified identifier from the item. As a best practice, you should use the attachment identifier to remove an attachment only if the same mail app has added that attachment in the same session. In Outlook Web App and OWA for Devices, the attachment identifier is valid only within the same session. A session is over when the user closes the app, or if the user starts composing in an inline form and subsequently pops out the inline form to continue in a separate window.</span></span>

##### <a name="parameters"></a><span data-ttu-id="c91b5-951">Parameter:</span><span class="sxs-lookup"><span data-stu-id="c91b5-951">Parameters:</span></span>

|<span data-ttu-id="c91b5-952">Name</span><span class="sxs-lookup"><span data-stu-id="c91b5-952">Name</span></span>| <span data-ttu-id="c91b5-953">Typ</span><span class="sxs-lookup"><span data-stu-id="c91b5-953">Type</span></span>| <span data-ttu-id="c91b5-954">Attribute</span><span class="sxs-lookup"><span data-stu-id="c91b5-954">Attributes</span></span>| <span data-ttu-id="c91b5-955">Beschreibung</span><span class="sxs-lookup"><span data-stu-id="c91b5-955">Description</span></span>|
|---|---|---|---|
|`attachmentId`| <span data-ttu-id="c91b5-956">Zeichenfolge</span><span class="sxs-lookup"><span data-stu-id="c91b5-956">String</span></span>||<span data-ttu-id="c91b5-p166">Der Bezeichner des Anhangs, der entfernt werden soll. Die maximale Länge der Zeichenfolge ist 100 Zeichen.</span><span class="sxs-lookup"><span data-stu-id="c91b5-p166">The identifier of the attachment to remove. The maximum length of the string is 100 characters.</span></span>|
|`options`| <span data-ttu-id="c91b5-959">Object</span><span class="sxs-lookup"><span data-stu-id="c91b5-959">Object</span></span>| <span data-ttu-id="c91b5-960">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="c91b5-960">&lt;optional&gt;</span></span>|<span data-ttu-id="c91b5-961">Ein Objektliteral, das mindestens eine der folgenden Eigenschaften enthält.</span><span class="sxs-lookup"><span data-stu-id="c91b5-961">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`| <span data-ttu-id="c91b5-962">Object</span><span class="sxs-lookup"><span data-stu-id="c91b5-962">Object</span></span>| <span data-ttu-id="c91b5-963">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="c91b5-963">&lt;optional&gt;</span></span>|<span data-ttu-id="c91b5-964">Entwickler können ein Objekt bereitstellen, auf das sie in der Callbackmethode zugreifen möchten.</span><span class="sxs-lookup"><span data-stu-id="c91b5-964">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`| <span data-ttu-id="c91b5-965">Funktion</span><span class="sxs-lookup"><span data-stu-id="c91b5-965">function</span></span>| <span data-ttu-id="c91b5-966">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="c91b5-966">&lt;optional&gt;</span></span>|<span data-ttu-id="c91b5-967">Wenn die Methode abgeschlossen wird, wird die Funktion , die im `callback`-Parameter übergeben wird, mit einem einzelnen Parameter (`asyncResult`) aufgerufen, der ein [`AsyncResult`](/javascript/api/office/office.asyncresult)-Objekt darstellt.</span><span class="sxs-lookup"><span data-stu-id="c91b5-967">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span> <br/><span data-ttu-id="c91b5-968">Wenn beim Entfernen der Anlage ein Fehler auftritt, enthält die Eigenschaft `asyncResult.error` einen Fehlercode mit dem Grund für den Fehler.</span><span class="sxs-lookup"><span data-stu-id="c91b5-968">If removing the attachment fails, the `asyncResult.error` property will contain an error code with the reason for the failure.</span></span>|

##### <a name="errors"></a><span data-ttu-id="c91b5-969">Fehler</span><span class="sxs-lookup"><span data-stu-id="c91b5-969">Errors</span></span>

| <span data-ttu-id="c91b5-970">Fehlercode</span><span class="sxs-lookup"><span data-stu-id="c91b5-970">Error code</span></span> | <span data-ttu-id="c91b5-971">Beschreibung</span><span class="sxs-lookup"><span data-stu-id="c91b5-971">Description</span></span> |
|------------|-------------|
| `InvalidAttachmentId` | <span data-ttu-id="c91b5-972">Der Bezeichner für die Anlage ist nicht vorhanden.</span><span class="sxs-lookup"><span data-stu-id="c91b5-972">The attachment identifier does not exist.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="c91b5-973">Anforderungen</span><span class="sxs-lookup"><span data-stu-id="c91b5-973">Requirements</span></span>

|<span data-ttu-id="c91b5-974">Anforderung</span><span class="sxs-lookup"><span data-stu-id="c91b5-974">Requirement</span></span>| <span data-ttu-id="c91b5-975">Wert</span><span class="sxs-lookup"><span data-stu-id="c91b5-975">Value</span></span>|
|---|---|
|[<span data-ttu-id="c91b5-976">Mindestversion des Postfachanforderungssatzes</span><span class="sxs-lookup"><span data-stu-id="c91b5-976">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="c91b5-977">1.1</span><span class="sxs-lookup"><span data-stu-id="c91b5-977">1.1</span></span>|
|[<span data-ttu-id="c91b5-978">Mindestberechtigungsstufe</span><span class="sxs-lookup"><span data-stu-id="c91b5-978">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="c91b5-979">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="c91b5-979">ReadWriteItem</span></span>|
|[<span data-ttu-id="c91b5-980">Zutreffender Outlook-Modus</span><span class="sxs-lookup"><span data-stu-id="c91b5-980">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="c91b5-981">Verfassen</span><span class="sxs-lookup"><span data-stu-id="c91b5-981">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="c91b5-982">Beispiel</span><span class="sxs-lookup"><span data-stu-id="c91b5-982">Example</span></span>

<span data-ttu-id="c91b5-983">Mit dem folgende Code wird eine Anlage mit dem Bezeichner "0" entfernt.</span><span class="sxs-lookup"><span data-stu-id="c91b5-983">The following code removes an attachment with an identifier of '0'.</span></span>

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

####  <a name="saveasyncoptions-callback"></a><span data-ttu-id="c91b5-984">saveAsync([options], callback)</span><span class="sxs-lookup"><span data-stu-id="c91b5-984">saveAsync([options], callback)</span></span>

<span data-ttu-id="c91b5-985">Speicher asynchron ein Element. .</span><span class="sxs-lookup"><span data-stu-id="c91b5-985">Asynchronously saves an item.</span></span>

<span data-ttu-id="c91b5-p167">Beim Aufrufen speichert diese Methode die aktuelle Nachricht als Entwurf und  gibt die Element-ID über die Callbackmethode zurück. In Outlook Web App oder Outlook im Onlinemodus wird das Element auf dem Server gespeichert. In Outlook im Cache-Modus wird das Element im lokalen Cache gespeichert.</span><span class="sxs-lookup"><span data-stu-id="c91b5-p167">When invoked, this method saves the current message as a draft and returns the item id via the callback method. In Outlook Web App or Outlook in online mode, the item is saved to the server. In Outlook in cached mode, the item is saved to the local cache.</span></span>

> [!NOTE]
> <span data-ttu-id="c91b5-989">Wenn Ihr Add-In ruft `saveAsync` für ein Element im Entwurfsmodus, um das Abrufen einer `itemId` um mit EWS oder REST-API verwenden, beachten Sie, dass wenn Outlook im Cache-Modus ist, es kann einige Zeit dauern, bevor das Element tatsächlich auf dem Server synchronisiert wird.</span><span class="sxs-lookup"><span data-stu-id="c91b5-989">If your add-in calls `saveAsync` on an item in compose mode in order to get an `itemId` to use with EWS or the REST API, be aware that when Outlook is in cached mode, it may take some time before the item is actually synced to the server.</span></span> <span data-ttu-id="c91b5-990">Bis das Element mit synchronisiert ist, die `itemId` wird ein Fehler zurückgegeben.</span><span class="sxs-lookup"><span data-stu-id="c91b5-990">Until the item is synced, using the `itemId` will return an error.</span></span>

<span data-ttu-id="c91b5-p169">Da Termine keinen Entwurfsstatus haben, wird das Element bei Aufruf von `saveAsync` für einen Termin im Verfassenmodus als normaler Termin im Kalender des Benutzers gespeichert. Für neue Termine, die noch nicht gespeichert wurden, wird keine Einladung gesendet. Beim Speichern eines vorhandenen Termins wird eine Aktualisierung an die hinzugefügten oder entfernten Teilnehmer gesendet.</span><span class="sxs-lookup"><span data-stu-id="c91b5-p169">Since appointments have no draft state, if `saveAsync` is called on an appointment in compose mode, the item will be saved as a normal appointment on the user's calendar. For new appointments that have not been saved before, no invitation will be sent. Saving an existing appointment will send an update to added or removed attendees.</span></span>

> [!NOTE]
> <span data-ttu-id="c91b5-994">Die folgenden Clients haben unterschiedlichem Verhalten für `saveAsync` Termine im Verfassenmodus:</span><span class="sxs-lookup"><span data-stu-id="c91b5-994">The following clients have different behavior for `saveAsync` on appointments in compose mode:</span></span>
>
> - <span data-ttu-id="c91b5-995">Outlook für Mac unterstützt keine `saveAsync` auf einer Besprechung im Verfassenmodus.</span><span class="sxs-lookup"><span data-stu-id="c91b5-995">Mac Outlook does not support `saveAsync` on a meeting in compose mode.</span></span> <span data-ttu-id="c91b5-996">Aufrufen von `saveAsync` auf einer Besprechung in Outlook für Mac wird ein Fehler zurückgegeben.</span><span class="sxs-lookup"><span data-stu-id="c91b5-996">Calling `saveAsync` on a meeting in Mac Outlook will return an error.</span></span>
> - <span data-ttu-id="c91b5-997">Outlook im Web immer sendet eine Einladung oder zu aktualisieren, wenn `saveAsync` für einen Termin aufgerufen wird, im Verfassenmodus.</span><span class="sxs-lookup"><span data-stu-id="c91b5-997">Outlook on the web always sends an invitation or update when `saveAsync` is called on an appointment in compose mode.</span></span>

##### <a name="parameters"></a><span data-ttu-id="c91b5-998">Parameter:</span><span class="sxs-lookup"><span data-stu-id="c91b5-998">Parameters:</span></span>

|<span data-ttu-id="c91b5-999">Name</span><span class="sxs-lookup"><span data-stu-id="c91b5-999">Name</span></span>| <span data-ttu-id="c91b5-1000">Typ</span><span class="sxs-lookup"><span data-stu-id="c91b5-1000">Type</span></span>| <span data-ttu-id="c91b5-1001">Attribute</span><span class="sxs-lookup"><span data-stu-id="c91b5-1001">Attributes</span></span>| <span data-ttu-id="c91b5-1002">Beschreibung</span><span class="sxs-lookup"><span data-stu-id="c91b5-1002">Description</span></span>|
|---|---|---|---|
|`options`| <span data-ttu-id="c91b5-1003">Objekt</span><span class="sxs-lookup"><span data-stu-id="c91b5-1003">Object</span></span>| <span data-ttu-id="c91b5-1004">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="c91b5-1004">&lt;optional&gt;</span></span>|<span data-ttu-id="c91b5-1005">Ein Objektliteral, das mindestens eine der folgenden Eigenschaften enthält.</span><span class="sxs-lookup"><span data-stu-id="c91b5-1005">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`| <span data-ttu-id="c91b5-1006">Object</span><span class="sxs-lookup"><span data-stu-id="c91b5-1006">Object</span></span>| <span data-ttu-id="c91b5-1007">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="c91b5-1007">&lt;optional&gt;</span></span>|<span data-ttu-id="c91b5-1008">Entwickler können ein Objekt bereitstellen, auf das sie in der Callbackmethode zugreifen möchten.</span><span class="sxs-lookup"><span data-stu-id="c91b5-1008">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`| <span data-ttu-id="c91b5-1009">Funktion</span><span class="sxs-lookup"><span data-stu-id="c91b5-1009">function</span></span>||<span data-ttu-id="c91b5-1010">Wenn die Methode abgeschlossen wird, wird die Funktion , die im `callback`-Parameter übergeben wird, mit einem einzelnen Parameter (`asyncResult`) aufgerufen, der ein [`AsyncResult`](/javascript/api/office/office.asyncresult)-Objekt darstellt.</span><span class="sxs-lookup"><span data-stu-id="c91b5-1010">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span><br/><br/><span data-ttu-id="c91b5-1011">Auf Erfolg, erfolgt die Element-ID der `asyncResult.value` Eigenschaft.</span><span class="sxs-lookup"><span data-stu-id="c91b5-1011">On success, the item identifier is provided in the `asyncResult.value` property.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="c91b5-1012">Anforderungen</span><span class="sxs-lookup"><span data-stu-id="c91b5-1012">Requirements</span></span>

|<span data-ttu-id="c91b5-1013">Anforderung</span><span class="sxs-lookup"><span data-stu-id="c91b5-1013">Requirement</span></span>| <span data-ttu-id="c91b5-1014">Wert</span><span class="sxs-lookup"><span data-stu-id="c91b5-1014">Value</span></span>|
|---|---|
|[<span data-ttu-id="c91b5-1015">Mindestversion des Postfachanforderungssatzes</span><span class="sxs-lookup"><span data-stu-id="c91b5-1015">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="c91b5-1016">1.3</span><span class="sxs-lookup"><span data-stu-id="c91b5-1016">1.3</span></span>|
|[<span data-ttu-id="c91b5-1017">Mindestberechtigungsstufe</span><span class="sxs-lookup"><span data-stu-id="c91b5-1017">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="c91b5-1018">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="c91b5-1018">ReadWriteItem</span></span>|
|[<span data-ttu-id="c91b5-1019">Zutreffender Outlook-Modus</span><span class="sxs-lookup"><span data-stu-id="c91b5-1019">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="c91b5-1020">Verfassen</span><span class="sxs-lookup"><span data-stu-id="c91b5-1020">Compose</span></span>|

##### <a name="examples"></a><span data-ttu-id="c91b5-1021">Beispiele</span><span class="sxs-lookup"><span data-stu-id="c91b5-1021">Examples</span></span>

```
Office.context.mailbox.item.saveAsync(
  function callback(result) {
    // Process the result
  });
```

<span data-ttu-id="c91b5-p171">Es folgt ein Beispiel für den `result`-Parameter, der an die Callbackfunktion übergeben wird. Die `value`-Eigenschaft enthält die Element-ID des Elements.</span><span class="sxs-lookup"><span data-stu-id="c91b5-p171">The following is an example of the `result` parameter passed to the callback function. The `value` property contains the item ID of the item.</span></span>

```
{
  "value":"AAMkADI5...AAA=",
  "status":"succeeded"
}
```

####  <a name="setselecteddataasyncdata-options-callback"></a><span data-ttu-id="c91b5-1024">setSelectedDataAsync(data, [options], callback)</span><span class="sxs-lookup"><span data-stu-id="c91b5-1024">setSelectedDataAsync(data, [options], callback)</span></span>

<span data-ttu-id="c91b5-1025">Fügt asynchron Daten in den Textkörper oder Betreff einer Nachricht ein.</span><span class="sxs-lookup"><span data-stu-id="c91b5-1025">Asynchronously inserts data into the body or subject of a message.</span></span>

<span data-ttu-id="c91b5-p172">Die `setSelectedDataAsync`-Methode fügt die angegebene Zeichenfolge an der Cursorposition im Betreff oder im Textkörper ein, oder, falls im Editor Text ausgewählt ist, ersetzt den markierten Text. Wenn sich der Cursor nicht im Textkörper oder im Betreffsfeld befindet, wird ein Fehler zurückgegeben. Nach dem Einfügen wird der Cursor am Ende der eingefügten Inhalte platziert.</span><span class="sxs-lookup"><span data-stu-id="c91b5-p172">The `setSelectedDataAsync` method inserts the specified string at the cursor location in the subject or body of the item, or, if text is selected in the editor, it replaces the selected text. If the cursor is not in the body or subject field, an error is returned. After insertion, the cursor is placed at the end of the inserted content.</span></span>

##### <a name="parameters"></a><span data-ttu-id="c91b5-1029">Parameter:</span><span class="sxs-lookup"><span data-stu-id="c91b5-1029">Parameters:</span></span>

|<span data-ttu-id="c91b5-1030">Name</span><span class="sxs-lookup"><span data-stu-id="c91b5-1030">Name</span></span>| <span data-ttu-id="c91b5-1031">Typ</span><span class="sxs-lookup"><span data-stu-id="c91b5-1031">Type</span></span>| <span data-ttu-id="c91b5-1032">Attribute</span><span class="sxs-lookup"><span data-stu-id="c91b5-1032">Attributes</span></span>| <span data-ttu-id="c91b5-1033">Beschreibung</span><span class="sxs-lookup"><span data-stu-id="c91b5-1033">Description</span></span>|
|---|---|---|---|
|`data`| <span data-ttu-id="c91b5-1034">Zeichenfolge</span><span class="sxs-lookup"><span data-stu-id="c91b5-1034">String</span></span>||<span data-ttu-id="c91b5-p173">Die einzufügenden Daten. Daten dürfen 1.000.000 Zeichen nicht überschreiten. Werden mehr als 1.000.000 Zeichen übergeben, wird eine `ArgumentOutOfRange`-Ausnahme ausgelöst.</span><span class="sxs-lookup"><span data-stu-id="c91b5-p173">The data to be inserted. Data is not to exceed 1,000,000 characters. If more than 1,000,000 characters are passed in, an `ArgumentOutOfRange` exception is thrown.</span></span>|
|`options`| <span data-ttu-id="c91b5-1038">Object</span><span class="sxs-lookup"><span data-stu-id="c91b5-1038">Object</span></span>| <span data-ttu-id="c91b5-1039">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="c91b5-1039">&lt;optional&gt;</span></span>|<span data-ttu-id="c91b5-1040">Ein Objektliteral, das mindestens eine der folgenden Eigenschaften enthält.</span><span class="sxs-lookup"><span data-stu-id="c91b5-1040">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`| <span data-ttu-id="c91b5-1041">Object</span><span class="sxs-lookup"><span data-stu-id="c91b5-1041">Object</span></span>| <span data-ttu-id="c91b5-1042">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="c91b5-1042">&lt;optional&gt;</span></span>|<span data-ttu-id="c91b5-1043">Entwickler können ein Objekt bereitstellen, auf das sie in der Rückrufmethode zugreifen möchten.</span><span class="sxs-lookup"><span data-stu-id="c91b5-1043">Developers can provide any object they wish to access in the callback method.</span></span>|
|`options.coercionType`| [<span data-ttu-id="c91b5-1044">Office.CoercionType</span><span class="sxs-lookup"><span data-stu-id="c91b5-1044">Office.CoercionType</span></span>](office.md#coerciontype-string)| <span data-ttu-id="c91b5-1045">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="c91b5-1045">&lt;optional&gt;</span></span>|<span data-ttu-id="c91b5-p174">Wenn `text`, wird das aktuelle Format in Outlook Web App und Outlook angewendet. Wenn das Feld ein HTML-Editor ist, werden nur die Textdaten eingefügt, selbst wenn es sich bei den Daten um HTML-Daten handelt.</span><span class="sxs-lookup"><span data-stu-id="c91b5-p174">If `text`, the current style is applied in Outlook Web App and Outlook. If the field is an HTML editor, only the text data is inserted, even if the data is HTML.</span></span><br/><br/><span data-ttu-id="c91b5-p175">Wenn `html` gesetzt ist und das Feld HTML unterstützt (der Betreff jedoch nicht), wird die aktuelle Formatvorlage in Outlook Web App angewendet und die Standardformatvorlage in Outlook. Ist das Feld ein Textfeld, wird ein Fehler des Typs `InvalidDataFormat` zurückgegeben.</span><span class="sxs-lookup"><span data-stu-id="c91b5-p175">If `html` and the field supports HTML (the subject doesn't), the current style is applied in Outlook Web App and the default style is applied in Outlook. If the field is a text field, an `InvalidDataFormat` error is returned.</span></span><br/><br/><span data-ttu-id="c91b5-1050">Wenn `coercionType` nicht festgelegt wird, hängt das Ergebnis vom Feld ab: Wenn das Feld HTML ist, wird HTML verwendet. Wenn das Feld Text ist, wird Nur-Text verwendet.</span><span class="sxs-lookup"><span data-stu-id="c91b5-1050">If `coercionType` is not set, the result depends on the field: if the field is HTML then HTML is used; if the field is text, then plain text is used.</span></span>|
|`callback`| <span data-ttu-id="c91b5-1051">function</span><span class="sxs-lookup"><span data-stu-id="c91b5-1051">function</span></span>||<span data-ttu-id="c91b5-1052">Wenn die Methode abgeschlossen wird, wird die Funktion , die im `callback`-Parameter übergeben wird, mit einem einzelnen Parameter (`asyncResult`) aufgerufen, der ein [`AsyncResult`](/javascript/api/office/office.asyncresult)-Objekt darstellt.</span><span class="sxs-lookup"><span data-stu-id="c91b5-1052">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="c91b5-1053">Anforderungen</span><span class="sxs-lookup"><span data-stu-id="c91b5-1053">Requirements</span></span>

|<span data-ttu-id="c91b5-1054">Anforderung</span><span class="sxs-lookup"><span data-stu-id="c91b5-1054">Requirement</span></span>| <span data-ttu-id="c91b5-1055">Wert</span><span class="sxs-lookup"><span data-stu-id="c91b5-1055">Value</span></span>|
|---|---|
|[<span data-ttu-id="c91b5-1056">Mindestversion des Postfachanforderungssatzes</span><span class="sxs-lookup"><span data-stu-id="c91b5-1056">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="c91b5-1057">1.2</span><span class="sxs-lookup"><span data-stu-id="c91b5-1057">1.2</span></span>|
|[<span data-ttu-id="c91b5-1058">Mindestberechtigungsstufe</span><span class="sxs-lookup"><span data-stu-id="c91b5-1058">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="c91b5-1059">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="c91b5-1059">ReadWriteItem</span></span>|
|[<span data-ttu-id="c91b5-1060">Zutreffender Outlook-Modus</span><span class="sxs-lookup"><span data-stu-id="c91b5-1060">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="c91b5-1061">Verfassen</span><span class="sxs-lookup"><span data-stu-id="c91b5-1061">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="c91b5-1062">Beispiel</span><span class="sxs-lookup"><span data-stu-id="c91b5-1062">Example</span></span>

```
Office.context.mailbox.item.setSelectedDataAsync("Hello World!");
Office.context.mailbox.item.setSelectedDataAsync("<b>Hello World!</b>", { coercionType : "html" });
```