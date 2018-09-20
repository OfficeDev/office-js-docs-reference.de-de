
# <a name="item"></a>item

### <a name="officeofficemdcontextofficecontextmdmailboxofficecontextmailboxmditem"></a>[Office](office.md)[.context](office.context.md)[.mailbox](office.context.mailbox.md).item

Der `item`-Namespace wird für den Zugriff auf die aktuell ausgewählte Nachricht, Besprechungsanfrage oder den aktuell ausgewählten Termin verwendet. Sie können den Typ von `item` mithilfe der [itemType](#itemtype-officemailboxenumsitemtypejavascriptapioutlook16officemailboxenumsitemtype)-Eigenschaft bestimmen.

##### <a name="requirements"></a>Anforderungen

|Anforderung| Wert|
|---|---|
|[Mindestversion des Postfachanforderungssatzes](/javascript/office/requirement-sets/outlook-api-requirement-sets)| 1.0|
|[Mindestberechtigungsstufe](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| Eingeschränkt|
|[Zutreffender Outlook-Modus](https://docs.microsoft.com/outlook/add-ins/#extension-points)| Verfassen oder Lesen|

##### <a name="members-and-methods"></a>Elemente und Methoden

| Element | Typ |
|--------|------|
| [attachments](#attachments-arrayattachmentdetailsjavascriptapioutlook16officeattachmentdetails) | Element |
| [bcc](#bcc-recipientsjavascriptapioutlook16officerecipients) | Element |
| [body](#body-bodyjavascriptapioutlook16officebody) | Element |
| [cc](#cc-arrayemailaddressdetailsjavascriptapioutlook16officeemailaddressdetailsrecipientsjavascriptapioutlook16officerecipients) | Element |
| [conversationId](#nullable-conversationid-string) | Element |
| [dateTimeCreated](#datetimecreated-date) | Element |
| [dateTimeModified](#datetimemodified-date) | Element |
| [end](#end-datetimejavascriptapioutlook16officetime) | Element |
| [from](#from-emailaddressdetailsjavascriptapioutlook16officeemailaddressdetails) | Element |
| [internetMessageId](#internetmessageid-string) | Element |
| [itemClass](#itemclass-string) | Element |
| [itemId](#nullable-itemid-string) | Element |
| [itemType](#itemtype-officemailboxenumsitemtypejavascriptapioutlook16officemailboxenumsitemtype) | Element |
| [location](#location-stringlocationjavascriptapioutlook16officelocation) | Element |
| [normalizedSubject](#normalizedsubject-string) | Element |
| [notificationMessages](#notificationmessages-notificationmessagesjavascriptapioutlook16officenotificationmessages) | Element |
| [optionalAttendees](#optionalattendees-arrayemailaddressdetailsjavascriptapioutlook16officeemailaddressdetailsrecipientsjavascriptapioutlook16officerecipients) | Element |
| [organizer](#organizer-emailaddressdetailsjavascriptapioutlook16officeemailaddressdetails) | Element |
| [requiredAttendees](#requiredattendees-arrayemailaddressdetailsjavascriptapioutlook16officeemailaddressdetailsrecipientsjavascriptapioutlook16officerecipients) | Element |
| [sender](#sender-emailaddressdetailsjavascriptapioutlook16officeemailaddressdetails) | Element |
| [start](#start-datetimejavascriptapioutlook16officetime) | Element |
| [subject](#subject-stringsubjectjavascriptapioutlook16officesubject) | Element |
| [to](#to-arrayemailaddressdetailsjavascriptapioutlook16officeemailaddressdetailsrecipientsjavascriptapioutlook16officerecipients) | Element |
| [addFileAttachmentAsync](#addfileattachmentasyncuri-attachmentname-options-callback) | Methode |
| [addItemAttachmentAsync](#additemattachmentasyncitemid-attachmentname-options-callback) | Methode |
| [close](#close) | Methode |
| [displayReplyAllForm](#displayreplyallformformdata) | Methode |
| [displayReplyForm](#displayreplyformformdata) | Methode |
| [getEntities](#getentities--entitiesjavascriptapioutlook16officeentities) | Methode |
| [getEntitiesByType](#getentitiesbytypeentitytype--nullable-arraystringcontactjavascriptapioutlook16officecontactmeetingsuggestionjavascriptapioutlook16officemeetingsuggestionphonenumberjavascriptapioutlook16officephonenumbertasksuggestionjavascriptapioutlook16officetasksuggestion) | Methode |
| [getFilteredEntitiesByName](#getfilteredentitiesbynamename--nullable-arraystringcontactjavascriptapioutlook16officecontactmeetingsuggestionjavascriptapioutlook16officemeetingsuggestionphonenumberjavascriptapioutlook16officephonenumbertasksuggestionjavascriptapioutlook16officetasksuggestion) | Methode |
| [getRegExMatches](#getregexmatches--object) | Methode |
| [getRegExMatchesByName](#getregexmatchesbynamename--nullable-array-string-) | Methode |
| [getSelectedDataAsync](#getselecteddataasynccoerciontype-options-callback--string) | Methode |
| [getSelectedEntities](#getselectedentities--entitiesjavascriptapioutlook16officeentities) | Methode |
| [getSelectedRegExMatches](#getselectedregexmatches--object) | Methode |
| [loadCustomPropertiesAsync](#loadcustompropertiesasynccallback-usercontext) | Methode |
| [removeAttachmentAsync](#removeattachmentasyncattachmentid-options-callback) | Methode |
| [saveAsync](#saveasyncoptions-callback) | Methode |
| [setSelectedDataAsync](#setselecteddataasyncdata-options-callback) | Methode |

### <a name="example"></a>Beispiel

Im folgenden JavaScript-Codebeispiel wird der Zugriff auf die `subject`-Eigenschaft des aktuellen Elements in Outlook veranschaulicht.

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

### <a name="members"></a>Elemente

#### <a name="attachments-arrayattachmentdetailsjavascriptapioutlook16officeattachmentdetails"></a>attachments :Array.<[AttachmentDetails](/javascript/api/outlook_1_6/office.attachmentdetails)>

Ruft ein Array mit Anlagen für das Element ab. Nur Lesemodus.

> [!NOTE]
> Bestimmte Dateitypen werden aufgrund von Sicherheitsproblemen von Outlook blockiert und daher nicht zurückgegeben werden. Weitere Informationen finden Sie unter [Blockierte Anlagen in Outlook](https://support.office.com/article/Blocked-attachments-in-Outlook-434752E1-02D3-4E90-9124-8B81E49A8519).

##### <a name="type"></a>Typ:

*   Array.<[AttachmentDetails](/javascript/api/outlook_1_6/office.attachmentdetails)>

##### <a name="requirements"></a>Anforderungen

|Anforderung| Wert|
|---|---|
|[Mindestversion des Postfachanforderungssatzes](/javascript/office/requirement-sets/outlook-api-requirement-sets)| 1.0|
|[Mindestberechtigungsstufe](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| ReadItem|
|[Zutreffender Outlook-Modus](https://docs.microsoft.com/outlook/add-ins/#extension-points)| Lesen|

##### <a name="example"></a>Beispiel

Mit dem folgende Code wird eine HTML-Zeichenfolge mit Details aller Anlagen für das aktuelle Element erstellt.

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

####  <a name="bcc-recipientsjavascriptapioutlook16officerecipients"></a>bcc :[Recipients](/javascript/api/outlook_1_6/office.recipients)

Ruft ein Objekt, das Methoden zum Abrufen oder aktualisieren die Empfänger in der Zeile Bcc (blind Carbon Copy, Blindkopie) einer Nachricht bereitstellt. Nur Verfassenmodus.

##### <a name="type"></a>Typ:

*   [Recipients](/javascript/api/outlook_1_6/office.recipients)

##### <a name="requirements"></a>Anforderungen

|Anforderung| Wert|
|---|---|
|[Mindestversion des Postfachanforderungssatzes](/javascript/office/requirement-sets/outlook-api-requirement-sets)| 1.1|
|[Mindestberechtigungsstufe](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| ReadItem|
|[Zutreffender Outlook-Modus](https://docs.microsoft.com/outlook/add-ins/#extension-points)| Verfassen|

##### <a name="example"></a>Beispiel

```
Office.context.mailbox.item.bcc.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.bcc.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.bcc.getAsync(callback);

function callback(asyncResult) {
  var arrayOfBccRecipients = asyncResult.value;
}
```

####  <a name="body-bodyjavascriptapioutlook16officebody"></a>body :[Body](/javascript/api/outlook_1_6/office.body)

Ruft ein Objekt ab, das Methoden zum Bearbeiten des Textkörpers eines Elements bereitstellt.

##### <a name="type"></a>Typ:

*   [Body](/javascript/api/outlook_1_6/office.body)

##### <a name="requirements"></a>Anforderungen

|Anforderung| Wert|
|---|---|
|[Mindestversion des Postfachanforderungssatzes](/javascript/office/requirement-sets/outlook-api-requirement-sets)| 1.1|
|[Mindestberechtigungsstufe](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| ReadItem|
|[Zutreffender Outlook-Modus](https://docs.microsoft.com/outlook/add-ins/#extension-points)| Verfassen oder Lesen|

####  <a name="cc-arrayemailaddressdetailsjavascriptapioutlook16officeemailaddressdetailsrecipientsjavascriptapioutlook16officerecipients"></a>cc: Array. <[EmailAddressDetails](/javascript/api/outlook_1_6/office.emailaddressdetails)>|[Empfänger](/javascript/api/outlook_1_6/office.recipients)

Ermöglicht den Zugriff auf die (Carbon Copy, Blindkopie) Cc-Empfänger einer Nachricht. Der Typ des Objekts und die Zugriffsebene, hängt von den Modus des aktuellen Elements ab.

##### <a name="read-mode"></a>Lesemodus

Die `cc`-Eigenschaft gibt ein Array mit einem `EmailAddressDetails`-Objekt für jeden Empfänger in der **Cc**-Zeile der Nachricht zurück. Die Auflistung ist auf höchstens 100 Elemente beschränkt.

##### <a name="compose-mode"></a>Verfassenmodus

Die `cc` -Eigenschaft gibt eine `Recipients` -Objekt, das Methoden zum Abrufen oder aktualisieren die Empfänger in der Zeile **Cc** der Nachricht bereitstellt.

##### <a name="type"></a>Typ:

*   Array.<[EmailAddressDetails](/javascript/api/outlook_1_6/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook_1_6/office.recipients)

##### <a name="requirements"></a>Anforderungen

|Anforderung| Wert|
|---|---|
|[Mindestversion des Postfachanforderungssatzes](/javascript/office/requirement-sets/outlook-api-requirement-sets)| 1.0|
|[Mindestberechtigungsstufe](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| ReadItem|
|[Zutreffender Outlook-Modus](https://docs.microsoft.com/outlook/add-ins/#extension-points)| Verfassen oder Lesen|

##### <a name="example"></a>Beispiel

```
Office.context.mailbox.item.cc.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.cc.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.cc.getAsync(callback);

function callback(asyncResult) {
  var arrayOfCcRecipients = asyncResult.value;
}
```

####  <a name="nullable-conversationid-string"></a>(nullable) conversationId :String

Ruft einen Bezeichner für die E-Mail-Unterhaltung ab, in der eine bestimmte Nachricht enthalten ist.

Sie können für diese Eigenschaft eine ganze Zahl abrufen, wenn Ihre Mail-App in Formularen zum Lesen oder Antworten in Formularen zum Verfassen aktiviert wird. Wenn der Benutzer den Betreff der Antwortnachricht ändert, ändert sich beim Versenden die Konversations-ID für die entsprechende Nachricht, und der Wert, den Sie vorher bezogen haben, trifft nicht länger zu.

Sie erhalten in einem Formular zum Verfassen für diese Eigenschaft für ein neues Element null. Wenn der Benutzer einen Betreff festlegt und das Element speichert, gibt die `conversationId`-Eigenschaft einen Wert zurück.

##### <a name="type"></a>Typ:

*   String

##### <a name="requirements"></a>Anforderungen

|Anforderung| Wert|
|---|---|
|[Mindestversion des Postfachanforderungssatzes](/javascript/office/requirement-sets/outlook-api-requirement-sets)| 1.0|
|[Mindestberechtigungsstufe](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| ReadItem|
|[Zutreffender Outlook-Modus](https://docs.microsoft.com/outlook/add-ins/#extension-points)| Verfassen oder Lesen|

#### <a name="datetimecreated-date"></a>dateTimeCreated :Date

Ruft das Datum und die Uhrzeit der Erstellung eines Elements ab. Nur Lesemodus.

##### <a name="type"></a>Typ:

*   Datum

##### <a name="requirements"></a>Anforderungen

|Anforderung| Wert|
|---|---|
|[Mindestversion des Postfachanforderungssatzes](/javascript/office/requirement-sets/outlook-api-requirement-sets)| 1.0|
|[Mindestberechtigungsstufe](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| ReadItem|
|[Zutreffender Outlook-Modus](https://docs.microsoft.com/outlook/add-ins/#extension-points)| Lesen|

##### <a name="example"></a>Beispiel

```
var created = Office.context.mailbox.item.dateTimeCreated;
```

#### <a name="datetimemodified-date"></a>dateTimeModified :Date

Ruft das Datum und die Uhrzeit der letzten Änderung eines Elements ab. Nur Lesemodus.

> [!NOTE]
> Dieser Member wird in Outlook für iOS oder Outlook für Android nicht unterstützt.

##### <a name="type"></a>Typ:

*   Datum

##### <a name="requirements"></a>Anforderungen

|Anforderung| Wert|
|---|---|
|[Mindestversion des Postfachanforderungssatzes](/javascript/office/requirement-sets/outlook-api-requirement-sets)| 1.0|
|[Mindestberechtigungsstufe](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| ReadItem|
|[Zutreffender Outlook-Modus](https://docs.microsoft.com/outlook/add-ins/#extension-points)| Lesen|

##### <a name="example"></a>Beispiel

```
var modified = Office.context.mailbox.item.dateTimeModified;
```

####  <a name="end-datetimejavascriptapioutlook16officetime"></a>end :Date|[Time](/javascript/api/outlook_1_6/office.time)

Ruft Datum und Zeit für das Ende des Termins ab oder legt diese fest.

Die `end`-Eigenschaft wird als Datums- und Uhrzeitwert in UTC ausgedrückt. Sie können die [`convertToLocalClientTime`](office.context.mailbox.md#converttolocalclienttimetimevalue--localclienttimejavascriptapioutlook16officelocalclienttime)-Methode verwenden, um den Endeigenschaftswert in den lokalen Uhrzeit- und Datumswert des Clients umzuwandeln.

##### <a name="read-mode"></a>Lesemodus

Die `end`-Eigenschaft gibt ein `Date`-Objekt zurück.

##### <a name="compose-mode"></a>Verfassenmodus

Die `end`-Eigenschaft gibt ein `Time`-Objekt zurück.

Wenn Sie die [`Time.setAsync`](/javascript/api/outlook_1_6/office.time#setasync-datetime--options--callback-)-Methode verwenden, um die Endzeit im Verfassenmodus festzulegen, sollten Sie die [`convertToUtcClientTime`](office.context.mailbox.md#converttoutcclienttimeinput--date)-Methode verwenden, um die Ortszeit auf dem Client für den Server in UTC umzuwandeln.

##### <a name="type"></a>Typ:

*   Date | [Time](/javascript/api/outlook_1_6/office.time)

##### <a name="requirements"></a>Anforderungen

|Anforderung| Wert|
|---|---|
|[Mindestversion des Postfachanforderungssatzes](/javascript/office/requirement-sets/outlook-api-requirement-sets)| 1.0|
|[Mindestberechtigungsstufe](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| ReadItem|
|[Zutreffender Outlook-Modus](https://docs.microsoft.com/outlook/add-ins/#extension-points)| Verfassen oder Lesen|

##### <a name="example"></a>Beispiel

Im folgenden Beispiel wird die Endzeit eines Termins im Verfassenmodus mithilfe der [`setAsync`](/javascript/api/outlook_1_6/office.time#setasync-datetime--options--callback-)-Methode des `Time`-Objekts festgelegt.

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

#### <a name="from-emailaddressdetailsjavascriptapioutlook16officeemailaddressdetails"></a>from :[EmailAddressDetails](/javascript/api/outlook_1_6/office.emailaddressdetails)

Ruft die E-Mail-Adresse des Absenders einer Nachricht ab. Nur Lesemodus.

Die `from`- und [`sender`](#sender-emailaddressdetailsjavascriptapioutlook16officeemailaddressdetails)-Eigenschaften stellen dieselbe Person dar, außer die Nachricht wurde von einem Delegaten gesendet. In diesem Fall stellt die `from`-Eigenschaft die Stellvertretung dar, und die sender-Eigenschaft stellt den Delegaten dar.

> [!NOTE]
> Die `recipientType` -Eigenschaft des der `EmailAddressDetails` -Objekts der `from` -Eigenschaft ist `undefined`.

##### <a name="type"></a>Typ:

*   [EmailAddressDetails](/javascript/api/outlook_1_6/office.emailaddressdetails)

##### <a name="requirements"></a>Anforderungen

|Anforderung| Wert|
|---|---|
|[Mindestversion des Postfachanforderungssatzes](/javascript/office/requirement-sets/outlook-api-requirement-sets)| 1.0|
|[Mindestberechtigungsstufe](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| ReadItem|
|[Zutreffender Outlook-Modus](https://docs.microsoft.com/outlook/add-ins/#extension-points)| Lesen|

#### <a name="internetmessageid-string"></a>internetMessageId :String

Ruft die Internetnachricht-ID einer E-Mail-Nachricht ab. Nur Lesemodus.

##### <a name="type"></a>Typ:

*   String

##### <a name="requirements"></a>Anforderungen

|Anforderung| Wert|
|---|---|
|[Mindestversion des Postfachanforderungssatzes](/javascript/office/requirement-sets/outlook-api-requirement-sets)| 1.0|
|[Mindestberechtigungsstufe](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| ReadItem|
|[Zutreffender Outlook-Modus](https://docs.microsoft.com/outlook/add-ins/#extension-points)| Lesen|

##### <a name="example"></a>Beispiel

```
var internetMessageId = Office.context.mailbox.item.internetMessageId;
```

#### <a name="itemclass-string"></a>itemClass :String

Ruft die Exchange-Webdienste-Elementklasse des ausgewählten Elements ab. Nur Lesemodus.

Die `itemClass`-Eigenschaft gibt die Nachrichtenklasse des ausgewählten Elements an. Folgende Standardnachrichtenklassen für das Nachrichten- oder Terminelement sind vorhanden:

| Typ | Beschreibung | Elementklasse |
| --- | --- | --- |
| Terminelemente | Dies sind Kalenderelemente der Elementklasse `IPM.Appointment` oder `IPM.Appointment.Occurence`. | `IPM.Appointment`<br />`IPM.Appointment.Occurence` |
| Nachrichtenelemente | Diese Elemente umfassen E-Mail-Nachrichten der Standardnachrichtenklasse `IPM.Note` sowie Besprechungsanfragen, -antworten und -absagen, die `IPM.Schedule.Meeting` als Basisnachrichtenklasse verwenden. | `IPM.Note`<br />`IPM.Schedule.Meeting.Request`<br />`IPM.Schedule.Meeting.Neg`<br />`IPM.Schedule.Meeting.Pos`<br />`IPM.Schedule.Meeting.Tent`<br />`IPM.Schedule.Meeting.Canceled` |

Sie können benutzerdefinierte Nachrichtenklassen zum Erweitern einer Standardnachrichtenklasse erstellen, z. B. eine benutzerdefinierte Terminnachrichtenklasse wie `IPM.Appointment.Contoso`.

##### <a name="type"></a>Typ:

*   String

##### <a name="requirements"></a>Anforderungen

|Anforderung| Wert|
|---|---|
|[Mindestversion des Postfachanforderungssatzes](/javascript/office/requirement-sets/outlook-api-requirement-sets)| 1.0|
|[Mindestberechtigungsstufe](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| ReadItem|
|[Zutreffender Outlook-Modus](https://docs.microsoft.com/outlook/add-ins/#extension-points)| Lesen|

##### <a name="example"></a>Beispiel

```
var itemClass = Office.context.mailbox.item.itemClass;
```

#### <a name="nullable-itemid-string"></a>(nullable) itemId :String

Ruft den Exchange-Webdienste-Elementbezeichner für das aktuelle Element ab. Nur Lesemodus.

> [!NOTE]
> Der Bezeichner, der von der `itemId`-Eigenschaft zurückgegeben wird, ist der gleiche Bezeichner wie der Exchange-Webdienste-Elementbezeichner. Die `itemId` -Eigenschaft ist nicht identisch mit der Outlook-Eintrags-ID oder die ID von der Outlook-REST-API verwendet. REST-API-Aufrufe dieser Wert verwenden, sollten sie konvertiert werden, bevor [Office.context.mailbox.convertToRestId](office.context.mailbox.md#converttorestiditemid-restversion--string)verwenden. Weitere Informationen finden Sie unter [Verwenden der Outlook-REST-APIs aus einem Outlook-add-in](https://docs.microsoft.com/outlook/add-ins/use-rest-api#get-the-item-id).

Die Eigenschaft `itemId` ist im Modus „Verfassen“ nicht verfügbar. Wenn ein Elementbezeichner erforderlich ist, kann die [`saveAsync`](#saveasyncoptions-callback)-Methode verwendet werden, um das Element zu speichern. Dadurch wird der Elementbezeichner im [`AsyncResult.value`](/javascript/api/office/office.asyncresult)-Parameter in der Rückruffunktion zurückgegeben.

##### <a name="type"></a>Typ:

*   String

##### <a name="requirements"></a>Anforderungen

|Anforderung| Wert|
|---|---|
|[Mindestversion des Postfachanforderungssatzes](/javascript/office/requirement-sets/outlook-api-requirement-sets)| 1.0|
|[Mindestberechtigungsstufe](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| ReadItem|
|[Zutreffender Outlook-Modus](https://docs.microsoft.com/outlook/add-ins/#extension-points)| Lesen|

##### <a name="example"></a>Beispiel

Mit dem folgende Code wird das Vorhandensein eines Elementbezeichners überprüft. Wenn die `itemId`-Eigenschaft `null` oder `undefined` zurückgibt, speichert sie das Element und ruft den Elementbezeichner aus dem asynchronen Ergebnis ab.

```
var itemId = Office.context.mailbox.item.itemId;
if (itemId === null || itemId == undefined) {
  Office.context.mailbox.item.saveAsync(function(result){
    itemId = result.value;
  });
}
```

####  <a name="itemtype-officemailboxenumsitemtypejavascriptapioutlook16officemailboxenumsitemtype"></a>itemType :[Office.MailboxEnums.ItemType](/javascript/api/outlook_1_6/office.mailboxenums.itemtype)

Ruft den Typ des Elements ab, das eine Instanz darstellt.

Die `itemType`-Eigenschaft gibt einen der Werte der `ItemType`-Enumeration zurück, der angibt, ob es sich bei der `item`-Objektinstanz um eine Nachricht oder einen Termin handelt.

##### <a name="type"></a>Typ:

*   [Office.MailboxEnums.ItemType](/javascript/api/outlook_1_6/office.mailboxenums.itemtype)

##### <a name="requirements"></a>Anforderungen

|Anforderung| Wert|
|---|---|
|[Mindestversion des Postfachanforderungssatzes](/javascript/office/requirement-sets/outlook-api-requirement-sets)| 1.0|
|[Mindestberechtigungsstufe](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| ReadItem|
|[Zutreffender Outlook-Modus](https://docs.microsoft.com/outlook/add-ins/#extension-points)| Verfassen oder Lesen|

##### <a name="example"></a>Beispiel

```
if (Office.context.mailbox.item.itemType == Office.MailboxEnums.ItemType.Message)
  // do something
else
  // do something else
```

####  <a name="location-stringlocationjavascriptapioutlook16officelocation"></a>location :String|[Location](/javascript/api/outlook_1_6/office.location)

Ruft den Ort eines Termins ab bzw. legt ihn fest.

##### <a name="read-mode"></a>Lesemodus

Die `location`-Eigenschaft gibt eine Zeichenfolge zurück, die den Ort des Termins enthält.

##### <a name="compose-mode"></a>Verfassenmodus

Die `location`-Eigenschaft gibt ein `Location`-Objekt zurück, das Methoden zum Abrufen und Festlegen des Orts für einen Termin bereitstellt.

##### <a name="type"></a>Typ:

*   String | [Location](/javascript/api/outlook_1_6/office.location)

##### <a name="requirements"></a>Anforderungen

|Anforderung| Wert|
|---|---|
|[Mindestversion des Postfachanforderungssatzes](/javascript/office/requirement-sets/outlook-api-requirement-sets)| 1.0|
|[Mindestberechtigungsstufe](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| ReadItem|
|[Zutreffender Outlook-Modus](https://docs.microsoft.com/outlook/add-ins/#extension-points)| Verfassen oder Lesen|

##### <a name="example"></a>Beispiel

```
var userContext = { value : 1 };
Office.context.mailbox.item.location.getAsync( { context: userContext}, callback);

function callback(asyncResult) {
  var context = asyncResult.context;
  var location = asyncResult.value;
}
```

#### <a name="normalizedsubject-string"></a>normalizedSubject :String

Ruft den Betreff für ein Element ab, wobei alle Präfixe entfernt werden (einschließlich `RE:` und `FWD:`). Nur Lesemodus.

Die normalizedSubject-Eigenschaft ruft den Betreff des Elements mit allen Standardpräfixen (z. B. `RE:` und `FW:`) ab, die von E-Mail-Programmen hinzugefügt werden. Wenn der Betreff des Elements mit vollständigen Präfixen abgerufen werden soll, verwenden Sie die [`subject`](#subject-stringsubjectjavascriptapioutlook16officesubject)-Eigenschaft.

##### <a name="type"></a>Typ:

*   String

##### <a name="requirements"></a>Anforderungen

|Anforderung| Wert|
|---|---|
|[Mindestversion des Postfachanforderungssatzes](/javascript/office/requirement-sets/outlook-api-requirement-sets)| 1.0|
|[Mindestberechtigungsstufe](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| ReadItem|
|[Zutreffender Outlook-Modus](https://docs.microsoft.com/outlook/add-ins/#extension-points)| Lesen|

##### <a name="example"></a>Beispiel

```
var normalizedSubject = Office.context.mailbox.item.normalizedSubject;
```

####  <a name="notificationmessages-notificationmessagesjavascriptapioutlook16officenotificationmessages"></a>notificationMessages :[NotificationMessages](/javascript/api/outlook_1_6/office.notificationmessages)

Ruft die Benachrichtigungen für ein Element ab.

##### <a name="type"></a>Typ:

*   [NotificationMessages](/javascript/api/outlook_1_6/office.notificationmessages)

##### <a name="requirements"></a>Anforderungen

|Anforderung| Wert|
|---|---|
|[Mindestversion des Postfachanforderungssatzes](/javascript/office/requirement-sets/outlook-api-requirement-sets)| 1.3|
|[Mindestberechtigungsstufe](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| ReadItem|
|[Zutreffender Outlook-Modus](https://docs.microsoft.com/outlook/add-ins/#extension-points)| Verfassen oder Lesen|

####  <a name="optionalattendees-arrayemailaddressdetailsjavascriptapioutlook16officeemailaddressdetailsrecipientsjavascriptapioutlook16officerecipients"></a>optionalAttendees :Array.<[EmailAddressDetails](/javascript/api/outlook_1_6/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook_1_6/office.recipients)

Ermöglicht den Zugriff auf die optionalen Teilnehmer eines Ereignisses. Der Typ des Objekts und die Zugriffsebene, hängt von den Modus des aktuellen Elements ab.

##### <a name="read-mode"></a>Lesemodus

Die `optionalAttendees`-Eigenschaft gibt ein Array mit einem `EmailAddressDetails`-Objekt für jeden optionalen Teilnehmer an der Besprechung zurück.

##### <a name="compose-mode"></a>Verfassenmodus

Die `optionalAttendees` -Eigenschaft gibt eine `Recipients` -Objekt, das Methoden zum Abrufen oder Aktualisieren der optionalen Teilnehmer für eine Besprechung bereitstellt.

##### <a name="type"></a>Typ:

*   Array.<[EmailAddressDetails](/javascript/api/outlook_1_6/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook/office.recipients)

##### <a name="requirements"></a>Anforderungen

|Anforderung| Wert|
|---|---|
|[Mindestversion des Postfachanforderungssatzes](/javascript/office/requirement-sets/outlook-api-requirement-sets)| 1.0|
|[Mindestberechtigungsstufe](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| ReadItem|
|[Zutreffender Outlook-Modus](https://docs.microsoft.com/outlook/add-ins/#extension-points)| Verfassen oder Lesen|

##### <a name="example"></a>Beispiel

```
Office.context.mailbox.item.optionalAttendees.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.optionalAttendees.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.optionalAttendees.getAsync(callback);

function callback(asyncResult) {
  var arrayOfOptionalAttendeesRecipients = asyncResult.value;
}
```

#### <a name="organizer-emailaddressdetailsjavascriptapioutlook16officeemailaddressdetails"></a>organizer :[EmailAddressDetails](/javascript/api/outlook_1_6/office.emailaddressdetails)

Ruft für eine angegebene Besprechung die E-Mail-Adresse des Organisators der Besprechung ab. Nur Lesemodus.

##### <a name="type"></a>Typ:

*   [EmailAddressDetails](/javascript/api/outlook_1_6/office.emailaddressdetails)

##### <a name="requirements"></a>Anforderungen

|Anforderung| Wert|
|---|---|
|[Mindestversion des Postfachanforderungssatzes](/javascript/office/requirement-sets/outlook-api-requirement-sets)| 1.0|
|[Mindestberechtigungsstufe](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| ReadItem|
|[Zutreffender Outlook-Modus](https://docs.microsoft.com/outlook/add-ins/#extension-points)| Lesen|

##### <a name="example"></a>Beispiel

```
var organizerName = Office.context.mailbox.item.organizer.displayName;
var organizerAddress = Office.context.mailbox.item.organizer.emailAddress;
```

####  <a name="requiredattendees-arrayemailaddressdetailsjavascriptapioutlook16officeemailaddressdetailsrecipientsjavascriptapioutlook16officerecipients"></a>requiredAttendees :Array.<[EmailAddressDetails](/javascript/api/outlook_1_6/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook_1_6/office.recipients)

Ermöglicht den Zugriff auf die erforderlichen Teilnehmer eines Ereignisses. Der Typ des Objekts und die Zugriffsebene, hängt von den Modus des aktuellen Elements ab.

##### <a name="read-mode"></a>Lesemodus

Die `requiredAttendees`-Eigenschaft gibt ein Array mit einem `EmailAddressDetails`-Objekt für jeden erforderlichen Teilnehmer an der Besprechung zurück.

##### <a name="compose-mode"></a>Verfassenmodus

Die `requiredAttendees` -Eigenschaft gibt eine `Recipients` -Objekt, das Methoden zum Abrufen oder Aktualisieren der erforderlichen Teilnehmer für eine Besprechung bereitstellt.

##### <a name="type"></a>Typ:

*   Array.<[EmailAddressDetails](/javascript/api/outlook_1_6/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook_1_6/office.recipients)

##### <a name="requirements"></a>Anforderungen

|Anforderung| Wert|
|---|---|
|[Mindestversion des Postfachanforderungssatzes](/javascript/office/requirement-sets/outlook-api-requirement-sets)| 1.0|
|[Mindestberechtigungsstufe](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| ReadItem|
|[Zutreffender Outlook-Modus](https://docs.microsoft.com/outlook/add-ins/#extension-points)| Verfassen oder Lesen|

##### <a name="example"></a>Beispiel

```
Office.context.mailbox.item.requiredAttendees.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.requiredAttendees.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.requiredAttendees.getAsync(callback);

function callback(asyncResult) {
  var arrayOfRequiredAttendeesRecipients = asyncResult.value;
}
```

#### <a name="sender-emailaddressdetailsjavascriptapioutlook16officeemailaddressdetails"></a>sender :[EmailAddressDetails](/javascript/api/outlook_1_6/office.emailaddressdetails)

Ruft die E-Mail-Adresse des Absenders einer E-Mail-Nachricht ab. Nur Lesemodus.

Die [`from`](#from-emailaddressdetailsjavascriptapioutlook16officeemailaddressdetails)- und `sender`-Eigenschaften stellen dieselbe Person dar, außer die Nachricht wurde von einem Delegaten gesendet. In diesem Fall stellt die `from`-Eigenschaft die Stellvertretung dar, und die sender-Eigenschaft stellt den Delegaten dar.

> [!NOTE]
> Die `recipientType` -Eigenschaft des der `EmailAddressDetails` -Objekts der `sender` -Eigenschaft ist `undefined`.

##### <a name="type"></a>Typ:

*   [EmailAddressDetails](/javascript/api/outlook_1_6/office.emailaddressdetails)

##### <a name="requirements"></a>Anforderungen

|Anforderung| Wert|
|---|---|
|[Mindestversion des Postfachanforderungssatzes](/javascript/office/requirement-sets/outlook-api-requirement-sets)| 1.0|
|[Mindestberechtigungsstufe](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| ReadItem|
|[Zutreffender Outlook-Modus](https://docs.microsoft.com/outlook/add-ins/#extension-points)| Lesen|

##### <a name="example"></a>Beispiel

```
var senderName = Office.context.mailbox.item.sender.displayName;
var senderAddress = Office.context.mailbox.item.sender.emailAddress;
```

####  <a name="start-datetimejavascriptapioutlook16officetime"></a>start :Date|[Time](/javascript/api/outlook_1_6/office.time)

Ruft Datum und Zeit für den Beginn des Termins ab oder legt Datum und Uhrzeit fest.

Die `start`-Eigenschaft wird als Datums- und Uhrzeitwert in UTC ausgedrückt. Sie können die [`convertToLocalClientTime`](office.context.mailbox.md#converttolocalclienttimetimevalue--localclienttimejavascriptapioutlook16officelocalclienttime)-Methode verwenden, um den Wert in den lokalen Uhrzeit- und Datumswert des Clients umzuwandeln.

##### <a name="read-mode"></a>Lesemodus

Die `start`-Eigenschaft gibt ein `Date`-Objekt zurück.

##### <a name="compose-mode"></a>Verfassenmodus

Die `start`-Eigenschaft gibt ein `Time`-Objekt zurück.

Wenn Sie die [`Time.setAsync`](/javascript/api/outlook_1_6/office.time#setasync-datetime--options--callback-)-Methode verwenden, um die Startzeit im Verfassenmodus festzulegen, sollten Sie die [`convertToUtcClientTime`](office.context.mailbox.md#converttoutcclienttimeinput--date)-Methode verwenden, um die Ortszeit auf dem Client für den Server in UTC umzuwandeln.

##### <a name="type"></a>Typ:

*   Date | [Time](/javascript/api/outlook_1_6/office.time)

##### <a name="requirements"></a>Anforderungen

|Anforderung| Wert|
|---|---|
|[Mindestversion des Postfachanforderungssatzes](/javascript/office/requirement-sets/outlook-api-requirement-sets)| 1.0|
|[Mindestberechtigungsstufe](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| ReadItem|
|[Zutreffender Outlook-Modus](https://docs.microsoft.com/outlook/add-ins/#extension-points)| Verfassen oder Lesen|

##### <a name="example"></a>Beispiel

Im folgenden Beispiel wird die Startzeit eines Termins im Verfassenmodus mithilfe der [`setAsync`](/javascript/api/outlook_1_6/office.time#setasync-datetime--options--callback-)-Methode des `Time`-Objekts festgelegt.

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

####  <a name="subject-stringsubjectjavascriptapioutlook16officesubject"></a>subject :String|[Subject](/javascript/api/outlook_1_6/office.subject)

Ruft die Beschreibung ab, die im Betrefffeld eines Elements angezeigt wird, oder legt sie fest.

Die `subject`-Eigenschaft ruft den gesamten Betreff des Elements ab oder legt ihn fest – so, wie er vom E-Mail-Server gesendet wird.

##### <a name="read-mode"></a>Lesemodus

Die `subject`-Eigenschaft gibt eine Zeichenfolge zurück. Verwenden Sie die [`normalizedSubject`](#normalizedsubject-string)-Eigenschaft, um den Betreff ohne führende Präfixe wie `RE:` und `FW:` abzurufen.

```
var subject = Office.context.mailbox.item.subject;
```

##### <a name="compose-mode"></a>Verfassenmodus

Die `subject`-Eigenschaft gibt ein `Subject`-Objekt zurück, das Methoden zum Abrufen und Festlegen des Betreffs bereitstellt.

```
Office.context.mailbox.item.subject.getAsync(callback);

function callback(asyncResult) {
  var subject = asyncResult.value;
}
```

##### <a name="type"></a>Typ:

*   String | [Subject](/javascript/api/outlook_1_6/office.subject)

##### <a name="requirements"></a>Anforderungen

|Anforderung| Wert|
|---|---|
|[Mindestversion des Postfachanforderungssatzes](/javascript/office/requirement-sets/outlook-api-requirement-sets)| 1.0|
|[Mindestberechtigungsstufe](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| ReadItem|
|[Zutreffender Outlook-Modus](https://docs.microsoft.com/outlook/add-ins/#extension-points)| Verfassen oder Lesen|

####  <a name="to-arrayemailaddressdetailsjavascriptapioutlook16officeemailaddressdetailsrecipientsjavascriptapioutlook16officerecipients"></a>an: Array. <[EmailAddressDetails](/javascript/api/outlook_1_6/office.emailaddressdetails)>|[Empfänger](/javascript/api/outlook_1_6/office.recipients)

Ermöglicht den Zugriff auf die Empfänger in der Zeile **an** einer Nachricht. Der Typ des Objekts und die Zugriffsebene, hängt von den Modus des aktuellen Elements ab.

##### <a name="read-mode"></a>Lesemodus

Die `to`-Eigenschaft gibt ein Array mit einem `EmailAddressDetails`-Objekt für jeden Empfänger in der**An**-Zeile der Nachricht zurück. Die Auflistung ist auf höchstens 100 Elemente beschränkt.

##### <a name="compose-mode"></a>Verfassenmodus

Die `to` -Eigenschaft gibt eine `Recipients` -Objekt, das Methoden zum Abrufen oder aktualisieren die Empfänger in der Zeile **an** der Nachricht bereitstellt.

##### <a name="type"></a>Typ:

*   Array.<[EmailAddressDetails](/javascript/api/outlook_1_6/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook_1_6/office.recipients)

##### <a name="requirements"></a>Anforderungen

|Anforderung| Wert|
|---|---|
|[Mindestversion des Postfachanforderungssatzes](/javascript/office/requirement-sets/outlook-api-requirement-sets)| 1.0|
|[Mindestberechtigungsstufe](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| ReadItem|
|[Zutreffender Outlook-Modus](https://docs.microsoft.com/outlook/add-ins/#extension-points)| Verfassen oder Lesen|

##### <a name="example"></a>Beispiel

```
Office.context.mailbox.item.to.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.to.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.to.getAsync(callback);

function callback(asyncResult) {
  var arrayOfToRecipients = asyncResult.value;
}
```

### <a name="methods"></a>Methoden

####  <a name="addfileattachmentasyncuri-attachmentname-options-callback"></a>addFileAttachmentAsync(uri, attachmentName, [options], [callback])

Fügt eine Datei zu einer Nachricht oder einem Termin als Anlage hinzu.

Die `addFileAttachmentAsync`-Methode lädt die Datei am angegebenen URI hoch und fügt sie an das Element im Verfassenformular an.

Anschließend können Sie den Bezeichner mit der [`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback)-Methode in der gleichen Sitzung zum Entfernen der Anlage verwenden.

##### <a name="parameters"></a>Parameter:

|Name| Typ| Attribute| Beschreibung|
|---|---|---|---|
|`uri`| Zeichenfolge||Der URI, der den Speicherort der an die Nachricht oder den Termin anzuhängenden Datei angibt. Die maximale Länge ist 2048 Zeichen.|
|`attachmentName`| Zeichenfolge||Der Name der Anlage, der beim Hochladen der Anlage angezeigt wird. Die maximale Länge ist 255 Zeichen.|
|`options`| Object| &lt;optional&gt;|Ein Objektliteral, das mindestens eine der folgenden Eigenschaften enthält.|
| `options.asyncContext` | Object | &lt;optional&gt; | Entwickler können jedes Objekt angeben, auf das sie in der Rückrufmethode zugreifen möchten. |
| `options.isInline` | Boolean | &lt;optional&gt; | Wenn `true`: Zeigt an, dass die Anlage inline im Nachrichtentext angezeigt wird und nicht in der Anlagenlisten angezeigt werden soll. |
|`callback`| function| &lt;optional&gt;|Wenn die Methode abgeschlossen wird, wird die Funktion , die im `callback`-Parameter übergeben wird, mit einem einzelnen Parameter (`asyncResult`) aufgerufen, der ein [`AsyncResult`](/javascript/api/office/office.asyncresult)-Objekt darstellt. <br/>Bei Erfolg wird der Anlagenbezeichner in der `asyncResult.value`-Eigenschaft bereitgestellt.<br/>Wenn beim Hochladen der Anlage ein Fehler auftritt, enthält das `asyncResult`-Objekt ein `Error`-Objekt mit einer Beschreibung des Fehlers.|

##### <a name="errors"></a>Fehler

| Fehlercode | Beschreibung |
|------------|-------------|
| `AttachmentSizeExceeded` | Die Anlage ist zu groß. |
| `FileTypeNotSupported` | Die Anlage enthält eine Erweiterung, die nicht zulässig ist. |
| `NumberOfAttachmentsExceeded` | Die Nachricht oder der Termin enthält zu viele Anlagen. |

##### <a name="requirements"></a>Anforderungen

|Anforderung| Wert|
|---|---|
|[Mindestversion des Postfachanforderungssatzes](/javascript/office/requirement-sets/outlook-api-requirement-sets)| 1.1|
|[Mindestberechtigungsstufe](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| ReadWriteItem|
|[Zutreffender Outlook-Modus](https://docs.microsoft.com/outlook/add-ins/#extension-points)| Verfassen|

##### <a name="examples"></a>Beispiele

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

Im folgenden Beispiel wird eine Bilddatei als Inlineanlage hinzugefügt, und die Anlage wird im Nachrichtentext referenziert.

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

####  <a name="additemattachmentasyncitemid-attachmentname-options-callback"></a>addItemAttachmentAsync(itemId, attachmentName, [options], [callback])

Fügt der Nachricht oder dem Termin ein Exchange-Objekt, wie z. B. eine Nachricht, als Anhang hinzu.

Die `addItemAttachmentAsync`-Methode fügt das Element mit dem angegebenen Exchange-Bezeichner an das Element im Formular zum Verfassen an. Wenn Sie eine Rückrufmethode angeben, wird die Methode mit einem `asyncResult`-Parameter aufgerufen, der entweder den Anlagenbezeichner oder einen Code enthält, der etwaige Fehler angibt, die beim Anfügen des Objekts aufgetreten sind. Sie können ggf. den `options`-Parameter verwenden, um Statusinformationen an die Rückrufmethode zu übergeben.

Anschließend können Sie den Bezeichner mit der [`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback)-Methode in der gleichen Sitzung zum Entfernen der Anlage verwenden.

Wenn Ihre Office-Add-Ins in Outlook Web App ausgeführt wird die `addItemAttachmentAsync` -Methode kann Anfügen von Elementen in der Elemente, die Sie bearbeiten; Allerdings Dies wird nicht unterstützt und wird nicht empfohlen.

##### <a name="parameters"></a>Parameter:

|Name| Typ| Attribute| Beschreibung|
|---|---|---|---|
|`itemId`| Zeichenfolge||Der Exchange-Bezeichner des Objekts, das angehängt werden soll. Die maximale Länge beträgt 100 Zeichen.|
|`attachmentName`| Zeichenfolge||Der Betreff des Elements, das angehängt werden soll. Die maximale Länge ist 255 Zeichen.|
|`options`| Object| &lt;optional&gt;|Ein Objektliteral, das mindestens eine der folgenden Eigenschaften enthält.|
|`options.asyncContext`| Object| &lt;optional&gt;|Entwickler können ein Objekt bereitstellen, auf das sie in der Callbackmethode zugreifen möchten.|
|`callback`| Funktion| &lt;optional&gt;|Wenn die Methode abgeschlossen wird, wird die Funktion , die im `callback`-Parameter übergeben wird, mit einem einzelnen Parameter (`asyncResult`) aufgerufen, der ein [`AsyncResult`](/javascript/api/office/office.asyncresult)-Objekt darstellt. <br/>Bei Erfolg wird der Anlagenbezeichner in der `asyncResult.value`-Eigenschaft bereitgestellt.<br/>Wenn beim Hinzufügen der Anlage ein Fehler auftritt, enthält das `asyncResult`-Objekt ein `Error`-Objekt mit einer Beschreibung des Fehlers.|

##### <a name="errors"></a>Fehler

| Fehlercode | Beschreibung |
|------------|-------------|
| `NumberOfAttachmentsExceeded` | Die Nachricht oder der Termin enthält zu viele Anlagen. |

##### <a name="requirements"></a>Anforderungen

|Anforderung| Wert|
|---|---|
|[Mindestversion des Postfachanforderungssatzes](/javascript/office/requirement-sets/outlook-api-requirement-sets)| 1.1|
|[Mindestberechtigungsstufe](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| ReadWriteItem|
|[Zutreffender Outlook-Modus](https://docs.microsoft.com/outlook/add-ins/#extension-points)| Verfassen|

##### <a name="example"></a>Beispiel

Im folgenden Beispiel wird ein vorhandenes Outlook-Element als Anlage mit dem Namen `My Attachment` hinzugefügt.

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

####  <a name="close"></a>close()

Schließt das aktuelle Element, das gerade verfasst wird.

Das Verhalten der `close`-Methode hängt vom aktuellen Status des verfassten Elements ab. Wenn das Element über nicht gespeicherte Änderungen verfügt, fordert der Client den Benutzer auf, die Schließenaktion zu speichern, zu verwerfen oder abzubrechen.

> [!NOTE]
> In Outlook im Web, wenn das Element einen Termin handelt, und es wurde bereits zuvor gespeichert wurde mit `saveAsync`, der Benutzer wird aufgefordert, speichern, löschen oder Abbrechen, auch wenn keine Änderungen aufgetreten sind, da das Element zuletzt gespeichert.

Wenn die Nachricht im Outlook-Desktopclient eine Inlineantwort ist, hat die `close`-Methode keine Auswirkung.

##### <a name="requirements"></a>Anforderungen

|Anforderung| Wert|
|---|---|
|[Mindestversion des Postfachanforderungssatzes](/javascript/office/requirement-sets/outlook-api-requirement-sets)| 1.3|
|[Mindestberechtigungsstufe](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| Eingeschränkt|
|[Zutreffender Outlook-Modus](https://docs.microsoft.com/outlook/add-ins/#extension-points)| Verfassen|

#### <a name="displayreplyallformformdata"></a>displayReplyAllForm(formData)

Zeigt ein Antwortformular an, das den Absender und alle Empfänger der ausgewählten Nachricht oder des Organisators und alle Teilnehmer des ausgewählten Termins enthält.

> [!NOTE]
> Diese Methode wird in Outlook für iOS oder Outlook für Android nicht unterstützt.

In Outlook Web App wird das Antwortformular als Popupformular in der Dreispaltenansicht und als Popupformular in der Zwei- oder Einspaltenansicht angezeigt.

Wenn einer der Zeichenfolgenparameter seinen Grenzwert überschreitet, löst `displayReplyAllForm` eine Ausnahme aus.

Wenn Anlagen im `formData.attachments`-Parameter angegeben werden, versuchen Outlook und Outlook Web App alle Anlagen herunterzuladen und sie an das Antwortformular anzufügen. Wenn Anlagen nicht hinzugefügt werden können, wird in der Formularbenutzeroberfläche ein Fehler ausgegeben. Wenn dies nicht möglich ist, wird keine Fehlermeldung ausgegeben.

##### <a name="parameters"></a>Parameter:

| Name | Typ | Attribute | Beschreibung |
|---|---|---|---|
|`formData`| Zeichenfolge &#124; Objekt| |Eine Zeichenfolge, die Text- und HTML-Code enthält, die den Hauptteil des Antwortformulars darstellen. Die Zeichenfolge ist auf 32 KB beschränkt.<br/>**ODER**<br/>Ein Objekt, das Text- oder Anlagendaten und eine Rückruffunktion enthält. Das Objekt ist wie folgt definiert: |
| `formData.htmlBody` | Zeichenfolge | &lt;optional&gt; | Eine Zeichenfolge, die Text- und HTML-Code enthält, die den Hauptteil des Antwortformulars darstellen. Die Zeichenfolge ist auf 32 KB beschränkt.
| `formData.attachments` | Array.&lt;Object&gt; | &lt;optional&gt; | Ein Array mit JSON-Objekten, die Datei- oder Elementanlagen sind. |
| `formData.attachments.type` | Zeichenfolge | | Gibt den Typ der Anlage an. Muss `file` für eine Dateianlage oder `item` für eine Elementanlage sein. |
| `formData.attachments.name` | Zeichenfolge | | Eine Zeichenfolge, die den Namen der Anlage mit bis zu 255 Zeichen enthält.|
| `formData.attachments.url` | Zeichenfolge | | Wird nur verwendet, wenn `type` auf `file` gesetzt ist. Der URI des Speicherorts für die Datei. |
| `formData.attachments.isInline` | Boolean | | Wird nur verwendet, wenn `type` auf `file` gesetzt ist. Wenn `true`: Zeigt an, dass die Anlage inline im Nachrichtentext angezeigt wird und nicht in der Anlagenlisten angezeigt werden soll. |
| `formData.attachments.itemId` | Zeichenfolge | | Wird nur verwendet, wenn `type` auf `item` gesetzt ist. Die EWS-Element-ID der Anlage. Dies ist eine Zeichenfolge bis zu 100 Zeichen. |
| `callback` | Funktion | &lt;optional&gt; | Bei Abschluss der Methode wird die im Parameter `callback` übergebene Funktion mit einem einzigen Parameter aufgerufen: `asyncResult`. Dieser Parameter ist ein Objekt des Typs [AsyncResult](/javascript/api/office/office.asyncresult). |

##### <a name="requirements"></a>Anforderungen

|Anforderung| Wert|
|---|---|
|[Mindestversion des Postfachanforderungssatzes](/javascript/office/requirement-sets/outlook-api-requirement-sets)| 1.0|
|[Mindestberechtigungsstufe](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| ReadItem|
|[Zutreffender Outlook-Modus](https://docs.microsoft.com/outlook/add-ins/#extension-points)| Lesen|

##### <a name="examples"></a>Beispiele

Im folgenden Code wird eine Zeichenfolge an die `displayReplyAllForm`-Funktion übergeben.

```
Office.context.mailbox.item.displayReplyAllForm('hello there');
Office.context.mailbox.item.displayReplyAllForm('<b>hello there</b>');
```

Antworten mit einem leeren Textkörper.

```
Office.context.mailbox.item.displayReplyAllForm({});
```

Antworten nur einem Textkörper.

```
Office.context.mailbox.item.displayReplyAllForm(
{
  'htmlBody' : 'hi'
});
```

Antworten mit einem Textkörper und einer Dateianlage.

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

Antworten mit einem Textkörper und einer Elementanlage.

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

Antworten Sie mit einem Textkörper, einer Dateianlage, einer Elementanlage und einem Callback.

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

#### <a name="displayreplyformformdata"></a>displayReplyForm(formData)

Zeigt ein Antwortformular an, das nur den Absender der ausgewählten Nachricht oder den Organisator des ausgewählten Termins enthält.

> [!NOTE]
> Diese Methode wird in Outlook für iOS oder Outlook für Android nicht unterstützt.

In Outlook Web App wird das Antwortformular als Popupformular in der Dreispaltenansicht und als Popupformular in der Zwei- oder Einspaltenansicht angezeigt.

Wenn einer der Zeichenfolgenparameter seinen Grenzwert überschreitet, löst `displayReplyForm` eine Ausnahme aus.

Wenn Anlagen im `formData.attachments`-Parameter angegeben werden, versuchen Outlook und Outlook Web App alle Anlagen herunterzuladen und sie an das Antwortformular anzufügen. Wenn Anlagen nicht hinzugefügt werden können, wird in der Formularbenutzeroberfläche ein Fehler ausgegeben. Wenn dies nicht möglich ist, wird keine Fehlermeldung ausgegeben.

##### <a name="parameters"></a>Parameter:

| Name | Typ | Attribute | Beschreibung |
|---|---|---|---|
|`formData`| Zeichenfolge &#124; Objekt| | Eine Zeichenfolge, die Text- und HTML-Code enthält, die den Hauptteil des Antwortformulars darstellen. Die Zeichenfolge ist auf 32 KB beschränkt.<br/>**ODER**<br/>Ein Objekt, das Text- oder Anlagendaten und eine Rückruffunktion enthält. Das Objekt ist wie folgt definiert: |
| `formData.htmlBody` | Zeichenfolge | &lt;optional&gt; | Eine Zeichenfolge, die Text- und HTML-Code enthält, die den Hauptteil des Antwortformulars darstellen. Die Zeichenfolge ist auf 32 KB beschränkt.
| `formData.attachments` | Array.&lt;Object&gt; | &lt;optional&gt; | Ein Array mit JSON-Objekten, die Datei- oder Elementanlagen sind. |
| `formData.attachments.type` | Zeichenfolge | | Gibt den Typ der Anlage an. Muss `file` für eine Dateianlage oder `item` für eine Elementanlage sein. |
| `formData.attachments.name` | Zeichenfolge | | Eine Zeichenfolge, die den Namen der Anlage mit bis zu 255 Zeichen enthält.|
| `formData.attachments.url` | Zeichenfolge | | Wird nur verwendet, wenn `type` auf `file` gesetzt ist. Der URI des Speicherorts für die Datei. |
| `formData.attachments.isInline` | Boolean | | Wird nur verwendet, wenn `type` auf `file` gesetzt ist. Wenn `true`: Zeigt an, dass die Anlage inline im Nachrichtentext angezeigt wird und nicht in der Anlagenlisten angezeigt werden soll. |
| `formData.attachments.itemId` | Zeichenfolge | | Wird nur verwendet, wenn `type` auf `item` gesetzt ist. Die EWS-Element-ID der Anlage. Dies ist eine Zeichenfolge bis zu 100 Zeichen. |
| `callback` | Funktion | &lt;optional&gt; | Bei Abschluss der Methode wird die im Parameter `callback` übergebene Funktion mit einem einzigen Parameter aufgerufen: `asyncResult`. Dieser Parameter ist ein Objekt des Typs [AsyncResult](/javascript/api/office/office.asyncresult). |

##### <a name="requirements"></a>Anforderungen

|Anforderung| Wert|
|---|---|
|[Mindestversion des Postfachanforderungssatzes](/javascript/office/requirement-sets/outlook-api-requirement-sets)| 1.0|
|[Mindestberechtigungsstufe](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| ReadItem|
|[Zutreffender Outlook-Modus](https://docs.microsoft.com/outlook/add-ins/#extension-points)| Lesen|

##### <a name="examples"></a>Beispiele

Im folgenden Code wird eine Zeichenfolge an die `displayReplyForm`-Funktion übergeben.

```
Office.context.mailbox.item.displayReplyForm('hello there');
Office.context.mailbox.item.displayReplyForm('<b>hello there</b>');
```

Antworten mit einem leeren Textkörper.

```
Office.context.mailbox.item.displayReplyForm({});
```

Antworten nur einem Textkörper.

```
Office.context.mailbox.item.displayReplyForm(
{
  'htmlBody' : 'hi'
});
```

Antworten mit einem Textkörper und einer Dateianlage.

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

Antworten mit einem Textkörper und einer Elementanlage.

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

Antworten Sie mit einem Textkörper, einer Dateianlage, einer Elementanlage und einem Callback.

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

#### <a name="getentities--entitiesjavascriptapioutlook16officeentities"></a>getEntities() → {[Entities](/javascript/api/outlook_1_6/office.entities)}

Ruft das ausgewählte Element Textkörper gefundenen Entitäten ab.

> [!NOTE]
> Diese Methode wird in Outlook für iOS oder Outlook für Android nicht unterstützt.

##### <a name="requirements"></a>Anforderungen

|Anforderung| Wert|
|---|---|
|[Mindestversion des Postfachanforderungssatzes](/javascript/office/requirement-sets/outlook-api-requirement-sets)| 1.0|
|[Mindestberechtigungsstufe](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| ReadItem|
|[Zutreffender Outlook-Modus](https://docs.microsoft.com/outlook/add-ins/#extension-points)| Lesen|

##### <a name="returns"></a>Gibt zurück:

Typ: [Entitäten](/javascript/api/outlook_1_6/office.entities)

##### <a name="example"></a>Beispiel

Das folgende Beispiel greift auf die Kontakte Entitäten im Textkörper des aktuellen Elements.

```
var contacts = Office.context.mailbox.item.getEntities().contacts;
```

#### <a name="getentitiesbytypeentitytype--nullable-arraystringcontactjavascriptapioutlook16officecontactmeetingsuggestionjavascriptapioutlook16officemeetingsuggestionphonenumberjavascriptapioutlook16officephonenumbertasksuggestionjavascriptapioutlook16officetasksuggestion"></a>getEntitiesByType(entityType) → (nullable) {Array.<(String|[Contact](/javascript/api/outlook_1_6/office.contact)|[MeetingSuggestion](/javascript/api/outlook_1_6/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook_1_6/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook_1_6/office.tasksuggestion))>}

Ruft ein Array aller Entitäten des angegebenen Entitätstyps im Hauptteil des ausgewählten Elements gefunden.

> [!NOTE]
> Diese Methode wird in Outlook für iOS oder Outlook für Android nicht unterstützt.

##### <a name="parameters"></a>Parameter:

|Name| Typ| Beschreibung|
|---|---|---|
|`entityType`| [Office.MailboxEnums.EntityType](/javascript/api/outlook_1_6/office.mailboxenums.entitytype)|Einer der Werte der EntityType-Enumeration.|

##### <a name="requirements"></a>Anforderungen

|Anforderung| Wert|
|---|---|
|[Mindestversion des Postfachanforderungssatzes](/javascript/office/requirement-sets/outlook-api-requirement-sets)| 1.0|
|[Mindestberechtigungsstufe](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| Eingeschränkt|
|[Zutreffender Outlook-Modus](https://docs.microsoft.com/outlook/add-ins/#extension-points)| Lesen|

##### <a name="returns"></a>Gibt zurück:

Wenn der in `entityType` übergebene Wert kein gültiges Element der `EntityType`-Enumeration ist, gibt die Methode null zurück. Wenn keine Entitäten des angegebenen Typs im Textkörper des Elements vorhanden sind, gibt die Methode ein leeres Array zurück. Andernfalls hängt der Typ der Objekte im zurückgegebenen Array vom Typ der Entität ab, die im `entityType`-Parameter angefordert wurde.

Während Sie die minimale Berechtigungsstufe für diese Methode **Restricted** ist, erfordern einige Entitätstypen **ReadItem** für den Zugriff, wie in der folgenden Tabelle angegeben wird.

| Wert von `entityType` | Typ der Objekte im zurückgegebenen Array | Erforderliche Berechtigungsstufe |
| --- | --- | --- |
| `Address` | Zeichenfolge | **Restricted** |
| `Contact` | Contact | **ReadItem** |
| `EmailAddress` | Zeichenfolge | **ReadItem** |
| `MeetingSuggestion` | MeetingSuggestion | **ReadItem** |
| `PhoneNumber` | PhoneNumber | **Restricted** |
| `TaskSuggestion` | TaskSuggestion | **ReadItem** |
| `URL` | Zeichenfolge | **Restricted** |

Typ: Array.<(String|[Contact](/javascript/api/outlook_1_6/office.contact)|[MeetingSuggestion](/javascript/api/outlook_1_6/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook_1_6/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook_1_6/office.tasksuggestion))>

##### <a name="example"></a>Beispiel

Im folgenden Beispiel wird veranschaulicht, wie auf ein Array von Zeichenfolgen, die Postadressen im Textkörper des aktuellen Elements darstellen.

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

#### <a name="getfilteredentitiesbynamename--nullable-arraystringcontactjavascriptapioutlook16officecontactmeetingsuggestionjavascriptapioutlook16officemeetingsuggestionphonenumberjavascriptapioutlook16officephonenumbertasksuggestionjavascriptapioutlook16officetasksuggestion"></a>getFilteredEntitiesByName(name) → (nullable) {Array.<(String|[Contact](/javascript/api/outlook_1_6/office.contact)|[MeetingSuggestion](/javascript/api/outlook_1_6/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook_1_6/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook_1_6/office.tasksuggestion))>}

Gibt bekannte Entitäten im ausgewählten Element zurück, die den in der XML-Manifestdatei definierten benannten Filter übergeben.

> [!NOTE]
> Diese Methode wird in Outlook für iOS oder Outlook für Android nicht unterstützt.

Die `getFilteredEntitiesByName`-Methode gibt die Entitäten zurück, die dem im [ItemHasKnownEntity](/javascript/office/manifest/rule#itemhasknownentity-rule)-Regelelement der XML-Manifestdatei definierten regulären Ausdruck mit dem angegebenen `FilterName`-Elementwert entsprechen.

##### <a name="parameters"></a>Parameter:

|Name| Typ| Beschreibung|
|---|---|---|
|`name`| Zeichenfolge|Der Name des `ItemHasKnownEntity`-Regelelements, das den entsprechenden Filter definiert.|

##### <a name="requirements"></a>Anforderungen

|Anforderung| Wert|
|---|---|
|[Mindestversion des Postfachanforderungssatzes](/javascript/office/requirement-sets/outlook-api-requirement-sets)| 1.0|
|[Mindestberechtigungsstufe](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| ReadItem|
|[Zutreffender Outlook-Modus](https://docs.microsoft.com/outlook/add-ins/#extension-points)| Lesen|

##### <a name="returns"></a>Gibt zurück:

Befindet sich kein `ItemHasKnownEntity`-Element im Manifest mit einem `FilterName`-Elementwert, der dem `name`-Parameter entspricht, gibt die Methode `null` zurück. Wenn der `name`-Parameter einem `ItemHasKnownEntity`-Element im Manifest entspricht, es aber keine entsprechenden Entitäten im aktuellen Element gibt, gibt die Methode ein leeres Array zurück.

Typ: Array.<(String|[Contact](/javascript/api/outlook_1_6/office.contact)|[MeetingSuggestion](/javascript/api/outlook_1_6/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook_1_6/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook_1_6/office.tasksuggestion))>

#### <a name="getregexmatches--object"></a>getRegExMatches() → {Object}

Gibt Zeichenfolgenwerte im ausgewählten Element zurück, die den in der XML-Manifestdatei definierten regulären Ausdrücken entsprechen.

> [!NOTE]
> Diese Methode wird in Outlook für iOS oder Outlook für Android nicht unterstützt.

Die `getRegExMatches`-Methode gibt die Zeichenfolgen zurück, die dem im `ItemHasRegularExpressionMatch`- oder `ItemHasKnownEntity`-Regelelement der XML-Manifestdatei definierten regulären Ausdruck entsprechen. Bei einer `ItemHasRegularExpressionMatch`-Regel muss eine entsprechende Zeichenfolge in der Eigenschaft des Elements vorhanden sein, das von dieser Regel angegeben wird. Der einfache `PropertyName`-Typ definiert die unterstützten Eigenschaften.

Nehmen Sie z. B. an, dass ein Add-In-Manifest das folgende `Rule`-Element aufweist:

```
<Rule xsi:type="RuleCollection" Mode="And">
  <Rule xsi:type="ItemIs" FormType="Read" ItemType="Message" />
  <Rule xsi:type="RuleCollection" Mode="Or">
    <Rule xsi:type="ItemHasRegularExpressionMatch" RegExName="fruits" RegExValue="apple|banana|coconut" PropertyName="BodyAsPlaintext" IgnoreCase="true" />
    <Rule xsi:type="ItemHasRegularExpressionMatch" RegExName="veggies" RegExValue="tomato|onion|spinach|broccoli" PropertyName="BodyAsPlaintext" IgnoreCase="true" />
  </Rule>
</Rule>
```

Das von `getRegExMatches` zurückgegebene Objekt hätte zwei Eigenschaften: `fruits` und `veggies`.

```
{
  'fruits': ['apple','banana','Banana','coconut'],
  'veggies': ['tomato','onion','spinach','broccoli']
}
```

Wenn Sie eine `ItemHasRegularExpressionMatch`-Regel für die Textkörpereigenschaft eines Elements festlegen, sollte der reguläre Ausdruck den Textkörper weiter filtern und nicht versuchen, den gesamten Textkörper des Elements zurückzugeben. Wenn der gesamte Textkörper eines Elements mit einem regulären Ausdruck wie `.*` abgerufen wird, werden nicht immer die gewünschten Ergebnisse erzielt. Verwenden Sie stattdessen die [`Body.getAsync`](/javascript/api/outlook_1_6/office.body#getasync-coerciontype--options--callback-)-Methode, um den gesamten Textkörper abzurufen.

##### <a name="requirements"></a>Anforderungen

|Anforderung| Wert|
|---|---|
|[Mindestversion des Postfachanforderungssatzes](/javascript/office/requirement-sets/outlook-api-requirement-sets)| 1.0|
|[Mindestberechtigungsstufe](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| ReadItem|
|[Zutreffender Outlook-Modus](https://docs.microsoft.com/outlook/add-ins/#extension-points)| Lesen|

##### <a name="returns"></a>Gibt zurück:

Ein Objekt mit Arrays aus Zeichenfolgen, die den in der XML-Manifestdatei definierten regulären Ausdrücken entsprechen. Der Name der einzelnen Arrays ist gleich dem entsprechenden Wert des `RegExName`-Attributs der entsprechenden `ItemHasRegularExpressionMatch`-Regel oder des `FilterName`-Attributs der entsprechenden `ItemHasKnownEntity`-Regel.

<dl class="param-type">

<dt>
Typ</dt>


<dd>Object</dd>

</dl>

##### <a name="example"></a>Beispiel

Das folgende Beispiel zeigt, wie Sie auf das Array von Übereinstimmungen für die regulären Ausdrucksregelelemente `fruits` und `veggies` zugreifen, die im Manifest angegeben sind.

```
var allMatches = Office.context.mailbox.item.getRegExMatches();
var fruits = allMatches.fruits;
var veges = allMatches.veggies;
```

#### <a name="getregexmatchesbynamename--nullable-array-string-"></a>getRegExMatchesByName(name) → (Nullwerte zulassen) {Array. < Zeichenfolge >}

Gibt Zeichenfolgenwerte im ausgewählten Element zurück, die dem in der XML-Manifestdatei definierten benannten regulären Ausdruck entsprechen.

> [!NOTE]
> Diese Methode wird in Outlook für iOS oder Outlook für Android nicht unterstützt.

Die `getRegExMatchesByName`-Methode gibt die Zeichenfolgen zurück, die dem im `ItemHasRegularExpressionMatch`-Regelelement der XML-Manifestdatei definierten regulären Ausdruck mit dem angegebenen `RegExName`-Elementwert entsprechen.

Wenn Sie eine `ItemHasRegularExpressionMatch`-Regel für die Textkörpereigenschaft eines Elements festlegen, sollte der reguläre Ausdruck den Textkörper weiter filtern und nicht versuchen, den gesamten Textkörper des Elements zurückzugeben. Wenn der gesamte Textkörper eines Elements mit einem regulären Ausdruck wie `.*` abgerufen wird, werden nicht immer die gewünschten Ergebnisse erzielt.

##### <a name="parameters"></a>Parameter:

|Name| Typ| Beschreibung|
|---|---|---|
|`name`| Zeichenfolge|Der Name des `ItemHasRegularExpressionMatch`-Regelelements, das den entsprechenden Filter definiert.|

##### <a name="requirements"></a>Anforderungen

|Anforderung| Wert|
|---|---|
|[Mindestversion des Postfachanforderungssatzes](/javascript/office/requirement-sets/outlook-api-requirement-sets)| 1.0|
|[Mindestberechtigungsstufe](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| ReadItem|
|[Zutreffender Outlook-Modus](https://docs.microsoft.com/outlook/add-ins/#extension-points)| Lesen|

##### <a name="returns"></a>Gibt zurück:

Ein Array mit den Zeichenfolgen, die dem in der XML-Manifestdatei definierten regulären Ausdruck entsprechen.

<dl class="param-type">

<dt>Typ</dt>

<dd>Array. < Zeichenfolge ></dd>

</dl>

##### <a name="example"></a>Beispiel

```
var fruits = Office.context.mailbox.item.getRegExMatchesByName("fruits");
var veggies = Office.context.mailbox.item.getRegExMatchesByName("veggies");
```

####  <a name="getselecteddataasynccoerciontype-options-callback--string"></a>getSelectedDataAsync(coercionType, [options], callback) → {String}

Gibt asynchron ausgewählte Daten aus dem Betreff oder Textkörper einer Nachricht zurück.

Wenn keine Auswahl vorhanden ist, aber der Cursor sich im Nachrichtentext oder Betreff befindet, gibt die Methode für die ausgewählten Daten NULL zurück. Wenn ein anderes Feld als der Textkörper oder Betreff ausgewählt ist, gibt die Methode den `InvalidSelection`-Fehler zurück.

##### <a name="parameters"></a>Parameter:

|Name| Typ| Attribute| Beschreibung|
|---|---|---|---|
|`coercionType`| [Office.CoercionType](office.md#coerciontype-string)||Fordert ein Format für die Daten an. Wenn es sich um Texthandelt, gibt die Methode den unformatierten Text als Zeichenfolge zurück und entfernt eventuell vorhandene HTML-Tags. Wenn es sich um HTML handelt, gibt die Methode den ausgewählten Text zurück, entweder als unformatierten Text oder als HTML.|
|`options`| Object| &lt;optional&gt;|Ein Objektliteral, das mindestens eine der folgenden Eigenschaften enthält.|
|`options.asyncContext`| Object| &lt;optional&gt;|Entwickler können ein Objekt bereitstellen, auf das sie in der Callbackmethode zugreifen möchten.|
|`callback`| Funktion||Wenn die Methode abgeschlossen wird, wird die Funktion , die im `callback`-Parameter übergeben wird, mit einem einzelnen Parameter (`asyncResult`) aufgerufen, der ein [`AsyncResult`](/javascript/api/office/office.asyncresult)-Objekt darstellt.<br/><br/>Rufen Sie für den Zugriff auf die ausgewählten Daten aus der Rückrufmethode `asyncResult.value.data` auf. Rufen Sie die Source-Eigenschaft für den Zugriff, die die Auswahl stammen, `asyncResult.value.sourceProperty`, die entweder sein `body` oder `subject`.|

##### <a name="requirements"></a>Anforderungen

|Anforderung| Wert|
|---|---|
|[Mindestversion des Postfachanforderungssatzes](/javascript/office/requirement-sets/outlook-api-requirement-sets)| 1.2|
|[Mindestberechtigungsstufe](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| ReadWriteItem|
|[Zutreffender Outlook-Modus](https://docs.microsoft.com/outlook/add-ins/#extension-points)| Verfassen|

##### <a name="returns"></a>Gibt zurück:

Die ausgewählten Daten als Zeichenfolge mit dem durch `coercionType` bestimmten Format.

<dl class="param-type">

<dt>
Typ</dt>


<dd>Zeichenfolge</dd>

</dl>

##### <a name="example"></a>Beispiel

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

#### <a name="getselectedentities--entitiesjavascriptapioutlook16officeentities"></a>getSelectedEntities() → {[Entitäten](/javascript/api/outlook_1_6/office.entities)}

Ruft die Entitäten ab, die in einer hervorgehobenen Übereinstimmung gefunden werden, die ein Benutzer ausgewählt hat. Hervorgehobene Übereinstimmungen gelten für [Kontext-Add-Ins](https://docs.microsoft.com/outlook/add-ins/contextual-outlook-add-ins).

> [!NOTE]
> Diese Methode wird in Outlook für iOS oder Outlook für Android nicht unterstützt.

##### <a name="requirements"></a>Anforderungen

|Anforderung| Wert|
|---|---|
|[Mindestversion des Postfachanforderungssatzes](/javascript/office/requirement-sets/outlook-api-requirement-sets)| 1.6 |
|[Mindestberechtigungsstufe](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| ReadItem|
|[Zutreffender Outlook-Modus](https://docs.microsoft.com/outlook/add-ins/#extension-points)| Lesen|

##### <a name="returns"></a>Gibt zurück:

Typ: [Entitäten](/javascript/api/outlook_1_6/office.entities)

##### <a name="example"></a>Beispiel

Im folgenden Beispiel wird auf die Adressentitäten in der hervorgehobenen Übereinstimmung zugegriffen, die der Benutzer ausgewählt hat.

```
var contacts = Office.context.mailbox.item.getSelectedEntities().addresses;
```

#### <a name="getselectedregexmatches--object"></a>getSelectedRegExMatches() → {Object}

Gibt Zeichenfolgenwerte in einer hervorgehobenen Übereinstimmung zurück, die den in der XML-Manifestdatei definierten regulären Ausdrücken entsprechen. Hervorgehobene Übereinstimmungen gelten für [Kontext-Add-Ins](https://docs.microsoft.com/outlook/add-ins/contextual-outlook-add-ins).

> [!NOTE]
> Diese Methode wird in Outlook für iOS oder Outlook für Android nicht unterstützt.

Die `getSelectedRegExMatches`-Methode gibt die Zeichenfolgen zurück, die dem im `ItemHasRegularExpressionMatch`- oder `ItemHasKnownEntity`-Regelelement der XML-Manifestdatei definierten regulären Ausdruck entsprechen. Bei einer `ItemHasRegularExpressionMatch`-Regel muss eine entsprechende Zeichenfolge in der Eigenschaft des Elements vorhanden sein, das von dieser Regel angegeben wird. Der einfache `PropertyName`-Typ definiert die unterstützten Eigenschaften.

Nehmen Sie z. B. an, dass ein Add-In-Manifest das folgende `Rule`-Element aufweist:

```
<Rule xsi:type="RuleCollection" Mode="And">
  <Rule xsi:type="ItemIs" FormType="Read" ItemType="Message" />
  <Rule xsi:type="RuleCollection" Mode="Or">
    <Rule xsi:type="ItemHasRegularExpressionMatch" RegExName="fruits" RegExValue="apple|banana|coconut" PropertyName="BodyAsPlaintext" IgnoreCase="true" />
    <Rule xsi:type="ItemHasRegularExpressionMatch" RegExName="veggies" RegExValue="tomato|onion|spinach|broccoli" PropertyName="BodyAsPlaintext" IgnoreCase="true" />
  </Rule>
</Rule>
```

Das von `getRegExMatches` zurückgegebene Objekt hätte zwei Eigenschaften: `fruits` und `veggies`.

```
{
  'fruits': ['apple','banana','Banana','coconut'],
  'veggies': ['tomato','onion','spinach','broccoli']
}
```

Wenn Sie eine `ItemHasRegularExpressionMatch`-Regel für die Textkörpereigenschaft eines Elements festlegen, sollte der reguläre Ausdruck den Textkörper weiter filtern und nicht versuchen, den gesamten Textkörper des Elements zurückzugeben. Wenn der gesamte Textkörper eines Elements mit einem regulären Ausdruck wie `.*` abgerufen wird, werden nicht immer die gewünschten Ergebnisse erzielt. Verwenden Sie stattdessen die [`Body.getAsync`](/javascript/api/outlook_1_6/office.body#getasync-coerciontype--options--callback-)-Methode, um den gesamten Textkörper abzurufen.

##### <a name="requirements"></a>Anforderungen

|Anforderung| Wert|
|---|---|
|[Mindestversion des Postfachanforderungssatzes](/javascript/office/requirement-sets/outlook-api-requirement-sets)| 1.6 |
|[Mindestberechtigungsstufe](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| ReadItem|
|[Zutreffender Outlook-Modus](https://docs.microsoft.com/outlook/add-ins/#extension-points)| Lesen|

##### <a name="returns"></a>Gibt zurück:

Ein Objekt mit Arrays aus Zeichenfolgen, die den in der XML-Manifestdatei definierten regulären Ausdrücken entsprechen. Der Name der einzelnen Arrays ist gleich dem entsprechenden Wert des `RegExName`-Attributs der entsprechenden `ItemHasRegularExpressionMatch`-Regel oder des `FilterName`-Attributs der entsprechenden `ItemHasKnownEntity`-Regel.

##### <a name="example"></a>Beispiel

Das folgende Beispiel zeigt, wie Sie auf das Array von Übereinstimmungen für die regulären Ausdrucksregelelemente `fruits` und `veggies` zugreifen, die im Manifest angegeben sind.

```
var selectedMatches = Office.context.mailbox.item.getSelectedRegExMatches();
var fruits = selectedMatches.fruits;
var veggies = selectedMatches.veggies;
```

####  <a name="loadcustompropertiesasynccallback-usercontext"></a>loadCustomPropertiesAsync(callback, [userContext])

Lädt asynchron benutzerdefinierte Eigenschaften für dieses Add-In für das ausgewählte Element.

Benutzerdefinierte Eigenschaften werden als Schlüssel-/Wert-Paare pro App und pro Element gespeichert. Diese Methode gibt ein `CustomProperties`-Objekt im Rückruf zurück, das Methoden für den Zugriff auf die benutzerdefinierten Eigenschaften für das aktuelle Element und das aktuelle Add-In bietet. Benutzerdefinierte Eigenschaften sind für das Element nicht verschlüsselt und sollten darum nicht als sicherer Speicher verwendet werden.

##### <a name="parameters"></a>Parameter:

|Name| Typ| Attribute| Beschreibung|
|---|---|---|---|
|`callback`| Funktion||Wenn die Methode abgeschlossen wird, wird die Funktion , die im `callback`-Parameter übergeben wird, mit einem einzelnen Parameter (`asyncResult`) aufgerufen, der ein [`AsyncResult`](/javascript/api/office/office.asyncresult)-Objekt darstellt.<br/><br/>Die benutzerdefinierten Eigenschaften werden als [`CustomProperties`](/javascript/api/outlook_1_6/office.customproperties)-Objekt in der `asyncResult.value`-Eigenschaft bereitgestellt. Dieses Objekt kann zum Abrufen, festlegen und Entfernen benutzerdefinierter Eigenschaften aus dem Element und speichern Sie Änderungen an den benutzerdefinierten Eigenschaftensatz zurück an den Server verwendet werden.|
|`userContext`| Objekt| &lt;optional&gt;|Entwickler können ein beliebiges Objekt bereitstellen, den sie in der Rückruffunktion zugreifen möchten. Dieses Objekt zugegriffen werden kann, indem die `asyncResult.asyncContext` -Eigenschaft in der Callback-Funktion.|

##### <a name="requirements"></a>Anforderungen

|Anforderung| Wert|
|---|---|
|[Mindestversion des Postfachanforderungssatzes](/javascript/office/requirement-sets/outlook-api-requirement-sets)| 1.0|
|[Mindestberechtigungsstufe](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| ReadItem|
|[Zutreffender Outlook-Modus](https://docs.microsoft.com/outlook/add-ins/#extension-points)| Verfassen oder Lesen|

##### <a name="example"></a>Beispiel

Das folgende Codebeispiel veranschaulicht die Verwendung der `loadCustomPropertiesAsync`-Methode zum asynchronen Laden der benutzerdefinierten Eigenschaften, die für das aktuelle Element spezifisch sind. Des Weiteren veranschaulicht das Beispiel die Verwendung der `CustomProperties.saveAsync`-Methode zum Speichern der Eigenschaften auf dem Server. Nach dem Laden der benutzerdefinierten Eigenschaften wird in dem Codebeispiel die `CustomProperties.get`-Methode dazu verwendet, die benutzerdefinierte `myProp`-Eigenschaft zu lesen. Die `CustomProperties.set`-Methode wird dazu verwendet, die benutzerdefinierte `otherProp`-Eigenschaft zu schreiben. Schließlich wird die `saveAsync`-Methode aufgerufen, um die benutzerdefinierten Eigenschaften zu speichern.

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

####  <a name="removeattachmentasyncattachmentid-options-callback"></a>removeAttachmentAsync(attachmentId, [options], [callback])

Entfernt eine Anlage aus einer Nachricht oder einem Termin.

Die `removeAttachmentAsync`-Methode entfernt die Anlage mit dem angegebenen Bezeichner aus dem Element. Als bewährte Vorgehensweise sollten Sie den Anlagenbezeichner nur dann zum Entfernen einer Anlage verwenden, wenn die gleiche Mail-App die Anlage in der gleichen Sitzung hinzugefügt hat. In Outlook Web App und OWA für Geräte ist der Anlagenbezeichner nur innerhalb der gleichen Sitzung gültig. Eine Sitzung ist abgeschlossen, wenn der Benutzer die App schließt, oder wenn der Benutzer in einem eingebetteten Formular mit dem Verfassen beginnt und den Vorgang dann in einem separaten Fenster fortsetzt.

##### <a name="parameters"></a>Parameter:

|Name| Typ| Attribute| Beschreibung|
|---|---|---|---|
|`attachmentId`| Zeichenfolge||Der Bezeichner des Anhangs, der entfernt werden soll. Die maximale Länge der Zeichenfolge ist 100 Zeichen.|
|`options`| Object| &lt;optional&gt;|Ein Objektliteral, das mindestens eine der folgenden Eigenschaften enthält.|
|`options.asyncContext`| Object| &lt;optional&gt;|Entwickler können ein Objekt bereitstellen, auf das sie in der Callbackmethode zugreifen möchten.|
|`callback`| Funktion| &lt;optional&gt;|Wenn die Methode abgeschlossen wird, wird die Funktion , die im `callback`-Parameter übergeben wird, mit einem einzelnen Parameter (`asyncResult`) aufgerufen, der ein [`AsyncResult`](/javascript/api/office/office.asyncresult)-Objekt darstellt. <br/>Wenn beim Entfernen der Anlage ein Fehler auftritt, enthält die Eigenschaft `asyncResult.error` einen Fehlercode mit dem Grund für den Fehler.|

##### <a name="errors"></a>Fehler

| Fehlercode | Beschreibung |
|------------|-------------|
| `InvalidAttachmentId` | Der Bezeichner für die Anlage ist nicht vorhanden. |

##### <a name="requirements"></a>Anforderungen

|Anforderung| Wert|
|---|---|
|[Mindestversion des Postfachanforderungssatzes](/javascript/office/requirement-sets/outlook-api-requirement-sets)| 1.1|
|[Mindestberechtigungsstufe](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| ReadWriteItem|
|[Zutreffender Outlook-Modus](https://docs.microsoft.com/outlook/add-ins/#extension-points)| Verfassen|

##### <a name="example"></a>Beispiel

Mit dem folgende Code wird eine Anlage mit dem Bezeichner "0" entfernt.

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

####  <a name="saveasyncoptions-callback"></a>saveAsync([options], callback)

Speicher asynchron ein Element. .

Beim Aufrufen speichert diese Methode die aktuelle Nachricht als Entwurf und  gibt die Element-ID über die Callbackmethode zurück. In Outlook Web App oder Outlook im Onlinemodus wird das Element auf dem Server gespeichert. In Outlook im Cache-Modus wird das Element im lokalen Cache gespeichert.

> [!NOTE]
> Wenn Ihr Add-In ruft `saveAsync` für ein Element im Entwurfsmodus, um das Abrufen einer `itemId` um mit EWS oder REST-API verwenden, beachten Sie, dass wenn Outlook im Cache-Modus ist, es kann einige Zeit dauern, bevor das Element tatsächlich auf dem Server synchronisiert wird. Bis das Element mit synchronisiert ist, die `itemId` wird ein Fehler zurückgegeben.

Da Termine keinen Entwurfsstatus haben, wird das Element bei Aufruf von `saveAsync` für einen Termin im Verfassenmodus als normaler Termin im Kalender des Benutzers gespeichert. Für neue Termine, die noch nicht gespeichert wurden, wird keine Einladung gesendet. Beim Speichern eines vorhandenen Termins wird eine Aktualisierung an die hinzugefügten oder entfernten Teilnehmer gesendet.

> [!NOTE]
> Die folgenden Clients haben unterschiedlichem Verhalten für `saveAsync` Termine im Verfassenmodus:
>
> - Outlook für Mac unterstützt keine `saveAsync` auf einer Besprechung im Verfassenmodus. Aufrufen von `saveAsync` auf einer Besprechung in Outlook für Mac wird ein Fehler zurückgegeben.
> - Outlook im Web immer sendet eine Einladung oder zu aktualisieren, wenn `saveAsync` für einen Termin aufgerufen wird, im Verfassenmodus.

##### <a name="parameters"></a>Parameter:

|Name| Typ| Attribute| Beschreibung|
|---|---|---|---|
|`options`| Objekt| &lt;optional&gt;|Ein Objektliteral, das mindestens eine der folgenden Eigenschaften enthält.|
|`options.asyncContext`| Object| &lt;optional&gt;|Entwickler können ein Objekt bereitstellen, auf das sie in der Callbackmethode zugreifen möchten.|
|`callback`| Funktion||Wenn die Methode abgeschlossen wird, wird die Funktion , die im `callback`-Parameter übergeben wird, mit einem einzelnen Parameter (`asyncResult`) aufgerufen, der ein [`AsyncResult`](/javascript/api/office/office.asyncresult)-Objekt darstellt.<br/><br/>Auf Erfolg, erfolgt die Element-ID der `asyncResult.value` Eigenschaft.|

##### <a name="requirements"></a>Anforderungen

|Anforderung| Wert|
|---|---|
|[Mindestversion des Postfachanforderungssatzes](/javascript/office/requirement-sets/outlook-api-requirement-sets)| 1.3|
|[Mindestberechtigungsstufe](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| ReadWriteItem|
|[Zutreffender Outlook-Modus](https://docs.microsoft.com/outlook/add-ins/#extension-points)| Verfassen|

##### <a name="examples"></a>Beispiele

```
Office.context.mailbox.item.saveAsync(
  function callback(result) {
    // Process the result
  });
```

Es folgt ein Beispiel für den `result`-Parameter, der an die Callbackfunktion übergeben wird. Die `value`-Eigenschaft enthält die Element-ID des Elements.

```
{
  "value":"AAMkADI5...AAA=",
  "status":"succeeded"
}
```

####  <a name="setselecteddataasyncdata-options-callback"></a>setSelectedDataAsync(data, [options], callback)

Fügt asynchron Daten in den Textkörper oder Betreff einer Nachricht ein.

Die `setSelectedDataAsync`-Methode fügt die angegebene Zeichenfolge an der Cursorposition im Betreff oder im Textkörper ein, oder, falls im Editor Text ausgewählt ist, ersetzt den markierten Text. Wenn sich der Cursor nicht im Textkörper oder im Betreffsfeld befindet, wird ein Fehler zurückgegeben. Nach dem Einfügen wird der Cursor am Ende der eingefügten Inhalte platziert.

##### <a name="parameters"></a>Parameter:

|Name| Typ| Attribute| Beschreibung|
|---|---|---|---|
|`data`| Zeichenfolge||Die einzufügenden Daten. Daten dürfen 1.000.000 Zeichen nicht überschreiten. Werden mehr als 1.000.000 Zeichen übergeben, wird eine `ArgumentOutOfRange`-Ausnahme ausgelöst.|
|`options`| Object| &lt;optional&gt;|Ein Objektliteral, das mindestens eine der folgenden Eigenschaften enthält.|
|`options.asyncContext`| Object| &lt;optional&gt;|Entwickler können ein Objekt bereitstellen, auf das sie in der Rückrufmethode zugreifen möchten.|
|`options.coercionType`| [Office.CoercionType](office.md#coerciontype-string)| &lt;optional&gt;|Wenn `text`, wird das aktuelle Format in Outlook Web App und Outlook angewendet. Wenn das Feld ein HTML-Editor ist, werden nur die Textdaten eingefügt, selbst wenn es sich bei den Daten um HTML-Daten handelt.<br/><br/>Wenn `html` gesetzt ist und das Feld HTML unterstützt (der Betreff jedoch nicht), wird die aktuelle Formatvorlage in Outlook Web App angewendet und die Standardformatvorlage in Outlook. Ist das Feld ein Textfeld, wird ein Fehler des Typs `InvalidDataFormat` zurückgegeben.<br/><br/>Wenn `coercionType` nicht festgelegt wird, hängt das Ergebnis vom Feld ab: Wenn das Feld HTML ist, wird HTML verwendet. Wenn das Feld Text ist, wird Nur-Text verwendet.|
|`callback`| function||Wenn die Methode abgeschlossen wird, wird die Funktion , die im `callback`-Parameter übergeben wird, mit einem einzelnen Parameter (`asyncResult`) aufgerufen, der ein [`AsyncResult`](/javascript/api/office/office.asyncresult)-Objekt darstellt. |

##### <a name="requirements"></a>Anforderungen

|Anforderung| Wert|
|---|---|
|[Mindestversion des Postfachanforderungssatzes](/javascript/office/requirement-sets/outlook-api-requirement-sets)| 1.2|
|[Mindestberechtigungsstufe](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| ReadWriteItem|
|[Zutreffender Outlook-Modus](https://docs.microsoft.com/outlook/add-ins/#extension-points)| Verfassen|

##### <a name="example"></a>Beispiel

```
Office.context.mailbox.item.setSelectedDataAsync("Hello World!");
Office.context.mailbox.item.setSelectedDataAsync("<b>Hello World!</b>", { coercionType : "html" });
```