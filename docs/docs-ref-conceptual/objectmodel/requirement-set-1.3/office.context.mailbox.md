
# <a name="mailbox"></a><span data-ttu-id="6c144-101">mailbox</span><span class="sxs-lookup"><span data-stu-id="6c144-101">mailbox</span></span>

### <span data-ttu-id="6c144-p101">[Office](Office.md)[.context](Office.context.md). mailbox</span><span class="sxs-lookup"><span data-stu-id="6c144-p101">[Office](Office.md)[.context](Office.context.md). mailbox</span></span>

<span data-ttu-id="6c144-104">Ermöglicht den Zugriff auf das Objektmodell von Outlook-add-in für Microsoft Outlook und Microsoft Outlook im Web.</span><span class="sxs-lookup"><span data-stu-id="6c144-104">Provides access to the Outlook add-in object model for Microsoft Outlook and Microsoft Outlook on the web.</span></span>

##### <a name="requirements"></a><span data-ttu-id="6c144-105">Anforderungen</span><span class="sxs-lookup"><span data-stu-id="6c144-105">Requirements</span></span>

|<span data-ttu-id="6c144-106">Anforderung</span><span class="sxs-lookup"><span data-stu-id="6c144-106">Requirement</span></span>| <span data-ttu-id="6c144-107">Wert</span><span class="sxs-lookup"><span data-stu-id="6c144-107">Value</span></span>|
|---|---|
|[<span data-ttu-id="6c144-108">Mindestversion des Postfachanforderungssatzes</span><span class="sxs-lookup"><span data-stu-id="6c144-108">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="6c144-109">1.0</span><span class="sxs-lookup"><span data-stu-id="6c144-109">1.0</span></span>|
|[<span data-ttu-id="6c144-110">Mindestberechtigungsstufe</span><span class="sxs-lookup"><span data-stu-id="6c144-110">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="6c144-111">Eingeschränkt</span><span class="sxs-lookup"><span data-stu-id="6c144-111">Restricted</span></span>|
|[<span data-ttu-id="6c144-112">Zutreffender Outlook-Modus</span><span class="sxs-lookup"><span data-stu-id="6c144-112">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="6c144-113">Verfassen oder Lesen</span><span class="sxs-lookup"><span data-stu-id="6c144-113">Compose or read</span></span>|

### <a name="namespaces"></a><span data-ttu-id="6c144-114">Namespaces</span><span class="sxs-lookup"><span data-stu-id="6c144-114">Namespaces</span></span>

<span data-ttu-id="6c144-115">[diagnostics](Office.context.mailbox.diagnostics.md): Stellt einem Outlook-Add-In Diagnoseinformationen bereit.</span><span class="sxs-lookup"><span data-stu-id="6c144-115">[diagnostics](Office.context.mailbox.diagnostics.md): Provides diagnostic information to an Outlook add-in.</span></span>

<span data-ttu-id="6c144-116">[item](Office.context.mailbox.item.md): Stellt Methoden und Eigenschaften für den Zugriff auf eine Nachricht oder einen Termins in einem Outlook-Add-In bereit.</span><span class="sxs-lookup"><span data-stu-id="6c144-116">[item](Office.context.mailbox.item.md): Provides methods and properties for accessing a message or appointment in an Outlook add-in.</span></span>

<span data-ttu-id="6c144-117">[userProfile](Office.context.mailbox.userProfile.md): Stellt Informationen zum Benutzer in einem Outlook-Add-In bereit.</span><span class="sxs-lookup"><span data-stu-id="6c144-117">[userProfile](Office.context.mailbox.userProfile.md): Provides information about the user in an Outlook add-in.</span></span>

### <a name="members"></a><span data-ttu-id="6c144-118">Elemente</span><span class="sxs-lookup"><span data-stu-id="6c144-118">Members</span></span>

#### <a name="ewsurl-string"></a><span data-ttu-id="6c144-119">ewsUrl :String</span><span class="sxs-lookup"><span data-stu-id="6c144-119">ewsUrl :String</span></span>

<span data-ttu-id="6c144-p102">Ruft die URL des EWS-Endpunkts (Exchange Web Services) für dieses E-Mail-Konto ab. Nur Lesemodus.</span><span class="sxs-lookup"><span data-stu-id="6c144-p102">Gets the URL of the Exchange Web Services (EWS) endpoint for this email account. Read mode only.</span></span>

> [!NOTE]
> <span data-ttu-id="6c144-122">Dieser Member wird in Outlook für iOS oder Outlook für Android nicht unterstützt.</span><span class="sxs-lookup"><span data-stu-id="6c144-122">This member is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="6c144-p103">Der `ewsUrl`-Wert kann durch einen Remotedienst zum Ausführen von EWS-Aufrufen an das Postfach des Benutzers verwendet werden. Sie können z. B. einen Remotedienst erstellen, um [Anlagen](https://docs.microsoft.com/outlook/add-ins/get-attachments-of-an-outlook-item) aus dem ausgewählten Element abzurufen.</span><span class="sxs-lookup"><span data-stu-id="6c144-p103">The `ewsUrl` value can be used by a remote service to make EWS calls to the user's mailbox. For example, you can create a remote service to [get attachments from the selected item](https://docs.microsoft.com/outlook/add-ins/get-attachments-of-an-outlook-item).</span></span>

<span data-ttu-id="6c144-125">Ihre App muss über die im Manifest angegebene **ReadItem**-Berechtigung verfügen, um das `ewsUrl`-Element im Lesemodus aufzurufen.</span><span class="sxs-lookup"><span data-stu-id="6c144-125">Your app must have the **ReadItem** permission specified in its manifest to call the `ewsUrl` member in read mode.</span></span>

<span data-ttu-id="6c144-p104">Im Verfassenmodus müssen Sie die [`saveAsync`](Office.context.mailbox.item.md#saveasyncoptions-callback)-Methode aufrufen, bevor Sie das `ewsUrl`-Element verwenden können. Die App muss über **ReadWriteItem**-Berechtigungen verfügen, um die `saveAsync`-Methode aufzurufen.</span><span class="sxs-lookup"><span data-stu-id="6c144-p104">In compose mode you must call the [`saveAsync`](Office.context.mailbox.item.md#saveasyncoptions-callback) method before you can use the `ewsUrl` member. Your app must have **ReadWriteItem** permissions to call the `saveAsync` method.</span></span>

##### <a name="type"></a><span data-ttu-id="6c144-128">Typ:</span><span class="sxs-lookup"><span data-stu-id="6c144-128">Type:</span></span>

*   <span data-ttu-id="6c144-129">String</span><span class="sxs-lookup"><span data-stu-id="6c144-129">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="6c144-130">Anforderungen</span><span class="sxs-lookup"><span data-stu-id="6c144-130">Requirements</span></span>

|<span data-ttu-id="6c144-131">Anforderung</span><span class="sxs-lookup"><span data-stu-id="6c144-131">Requirement</span></span>| <span data-ttu-id="6c144-132">Wert</span><span class="sxs-lookup"><span data-stu-id="6c144-132">Value</span></span>|
|---|---|
|[<span data-ttu-id="6c144-133">Mindestversion des Postfachanforderungssatzes</span><span class="sxs-lookup"><span data-stu-id="6c144-133">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="6c144-134">1.0</span><span class="sxs-lookup"><span data-stu-id="6c144-134">1.0</span></span>|
|[<span data-ttu-id="6c144-135">Mindestberechtigungsstufe</span><span class="sxs-lookup"><span data-stu-id="6c144-135">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="6c144-136">ReadItem</span><span class="sxs-lookup"><span data-stu-id="6c144-136">ReadItem</span></span>|
|[<span data-ttu-id="6c144-137">Zutreffender Outlook-Modus</span><span class="sxs-lookup"><span data-stu-id="6c144-137">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="6c144-138">Verfassen oder Lesen</span><span class="sxs-lookup"><span data-stu-id="6c144-138">Compose or read</span></span>|

### <a name="methods"></a><span data-ttu-id="6c144-139">Methoden</span><span class="sxs-lookup"><span data-stu-id="6c144-139">Methods</span></span>

####  <a name="converttoewsiditemid-restversion--string"></a><span data-ttu-id="6c144-140">convertToEwsId(itemId, restVersion) → {String}</span><span class="sxs-lookup"><span data-stu-id="6c144-140">convertToEwsId(itemId, restVersion) → {String}</span></span>

<span data-ttu-id="6c144-141">Wandelt eine für REST formatierte Element-ID ins EWS-Format um.</span><span class="sxs-lookup"><span data-stu-id="6c144-141">Converts an item ID formatted for REST into EWS format.</span></span>

> [!NOTE]
> <span data-ttu-id="6c144-142">Diese Methode wird in Outlook für iOS oder Outlook für Android nicht unterstützt.</span><span class="sxs-lookup"><span data-stu-id="6c144-142">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="6c144-p105">Über eine REST-API abgerufene Element-IDs (z. B. die [Outlook Mail-API](https://docs.microsoft.com/previous-versions/office/office-365-api/api/version-2.0/mail-rest-operations) oder [Microsoft Graph](http://graph.microsoft.io/)) verwenden ein anderes Format, als das von Exchange-Webdiensten (EWS) verwendete. Die `convertToEwsId`-Methode wandelt eine für REST formatierte ID in das entsprechende EWS-Format um.</span><span class="sxs-lookup"><span data-stu-id="6c144-p105">Item IDs retrieved via a REST API (such as the [Outlook Mail API](https://docs.microsoft.com/previous-versions/office/office-365-api/api/version-2.0/mail-rest-operations) or the [Microsoft Graph](http://graph.microsoft.io/)) use a different format than the format used by Exchange Web Services (EWS). The `convertToEwsId` method converts a REST-formatted ID into the proper format for EWS.</span></span>

##### <a name="parameters"></a><span data-ttu-id="6c144-145">Parameter:</span><span class="sxs-lookup"><span data-stu-id="6c144-145">Parameters:</span></span>

|<span data-ttu-id="6c144-146">Name</span><span class="sxs-lookup"><span data-stu-id="6c144-146">Name</span></span>| <span data-ttu-id="6c144-147">Typ</span><span class="sxs-lookup"><span data-stu-id="6c144-147">Type</span></span>| <span data-ttu-id="6c144-148">Beschreibung</span><span class="sxs-lookup"><span data-stu-id="6c144-148">Description</span></span>|
|---|---|---|
|`itemId`| <span data-ttu-id="6c144-149">Zeichenfolge</span><span class="sxs-lookup"><span data-stu-id="6c144-149">String</span></span>|<span data-ttu-id="6c144-150">Eine für Outlook-REST-APIs formatierte Element-ID</span><span class="sxs-lookup"><span data-stu-id="6c144-150">An item ID formatted for the Outlook REST APIs</span></span>|
|`restVersion`| [<span data-ttu-id="6c144-151">Office.MailboxEnums.RestVersion</span><span class="sxs-lookup"><span data-stu-id="6c144-151">Office.MailboxEnums.RestVersion</span></span>](/javascript/api/outlook_1_3/office.mailboxenums.restversion)|<span data-ttu-id="6c144-152">Ein Wert, der die Version der Outlook-REST-API angibt, die zum Abrufen der Element-ID verwendet wurde.</span><span class="sxs-lookup"><span data-stu-id="6c144-152">A value indicating the version of the Outlook REST API used to retrieve the item ID.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="6c144-153">Anforderungen</span><span class="sxs-lookup"><span data-stu-id="6c144-153">Requirements</span></span>

|<span data-ttu-id="6c144-154">Anforderung</span><span class="sxs-lookup"><span data-stu-id="6c144-154">Requirement</span></span>| <span data-ttu-id="6c144-155">Wert</span><span class="sxs-lookup"><span data-stu-id="6c144-155">Value</span></span>|
|---|---|
|[<span data-ttu-id="6c144-156">Mindestversion des Postfachanforderungssatzes</span><span class="sxs-lookup"><span data-stu-id="6c144-156">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="6c144-157">1.3</span><span class="sxs-lookup"><span data-stu-id="6c144-157">1.3</span></span>|
|[<span data-ttu-id="6c144-158">Mindestberechtigungsstufe</span><span class="sxs-lookup"><span data-stu-id="6c144-158">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="6c144-159">Eingeschränkt</span><span class="sxs-lookup"><span data-stu-id="6c144-159">Restricted</span></span>|
|[<span data-ttu-id="6c144-160">Zutreffender Outlook-Modus</span><span class="sxs-lookup"><span data-stu-id="6c144-160">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="6c144-161">Verfassen oder Lesen</span><span class="sxs-lookup"><span data-stu-id="6c144-161">Compose or read</span></span>|

##### <a name="returns"></a><span data-ttu-id="6c144-162">Gibt zurück:</span><span class="sxs-lookup"><span data-stu-id="6c144-162">Returns:</span></span>

<span data-ttu-id="6c144-163">Typ: Zeichenfolge</span><span class="sxs-lookup"><span data-stu-id="6c144-163">Type: String</span></span>

##### <a name="example"></a><span data-ttu-id="6c144-164">Beispiel</span><span class="sxs-lookup"><span data-stu-id="6c144-164">Example</span></span>

```
// Get an item's ID from a REST API
var restId = 'AAMkAGVlOTZjNTM3LW...';

// Treat restId as coming from the v2.0 version of the
// Outlook Mail API
var ewsId = Office.context.mailbox.convertToEwsId(restId, Office.MailboxEnums.RestVersion.v2_0);
```

####  <a name="converttolocalclienttimetimevalue--localclienttimejavascriptapioutlook13officelocalclienttime"></a><span data-ttu-id="6c144-165">convertToLocalClientTime(timeValue) → {[LocalClientTime](/javascript/api/outlook_1_3/office.LocalClientTime)}</span><span class="sxs-lookup"><span data-stu-id="6c144-165">convertToLocalClientTime(timeValue) → {[LocalClientTime](/javascript/api/outlook_1_3/office.LocalClientTime)}</span></span>

<span data-ttu-id="6c144-166">Ruft ein Wörterbuch mit Uhrzeitinformationen basierend auf der Zeiteinstellung des lokalen Clients ab.</span><span class="sxs-lookup"><span data-stu-id="6c144-166">Gets a dictionary containing time information in local client time.</span></span>

<span data-ttu-id="6c144-p106">Die in einer Mail-App für Outlook oder Outlook Web App verwendeten Daten und Uhrzeiten stammen u. U. aus verschiedenen Zeitzonen. Outlook verwendet die Zeitzone des Client-Computers; Outlook Web App verwendet die im Exchange Admin Center (EAC) festgelegte Zeitzone. Sie sollten Datums- und Uhrzeitwerte bearbeiten, damit die auf der Benutzeroberfläche angezeigten Werte immer den von Benutzer erwarteten Zeitzonen entsprechen.</span><span class="sxs-lookup"><span data-stu-id="6c144-p106">The dates and times used by a mail app for Outlook or Outlook Web App can use different time zones. Outlook uses the client computer time zone; Outlook Web App uses the time zone set on the Exchange Admin Center (EAC). You should handle date and time values so that the values you display on the user interface are always consistent with the time zone that the user expects.</span></span>

<span data-ttu-id="6c144-p107">Wird die Mail-App in Outlook ausgeführt, wird mit der `convertToLocalClientTime`-Methode ein Wörterbuchobjekt zurückgegeben, dessen Werte auf die Zeitzone des Clientcomputers festgelegt sind. Wird die Mail-App in Outlook ausgeführt, wird mit der `convertToLocalClientTime`-Methode ein Wörterbuchobjekt zurückgegeben, dessen Werte auf die Zeitzone des Clientcomputers festgelegt sind.</span><span class="sxs-lookup"><span data-stu-id="6c144-p107">If the mail app is running in Outlook, the `convertToLocalClientTime` method will return a dictionary object with the values set to the client computer time zone. If the mail app is running in Outlook Web App, the `convertToLocalClientTime` method will return a dictionary object with the values set to the time zone specified in the EAC.</span></span>

##### <a name="parameters"></a><span data-ttu-id="6c144-172">Parameter:</span><span class="sxs-lookup"><span data-stu-id="6c144-172">Parameters:</span></span>

|<span data-ttu-id="6c144-173">Name</span><span class="sxs-lookup"><span data-stu-id="6c144-173">Name</span></span>| <span data-ttu-id="6c144-174">Typ</span><span class="sxs-lookup"><span data-stu-id="6c144-174">Type</span></span>| <span data-ttu-id="6c144-175">Beschreibung</span><span class="sxs-lookup"><span data-stu-id="6c144-175">Description</span></span>|
|---|---|---|
|`timeValue`| <span data-ttu-id="6c144-176">Datum</span><span class="sxs-lookup"><span data-stu-id="6c144-176">Date</span></span>|<span data-ttu-id="6c144-177">Ein Date-Objekt</span><span class="sxs-lookup"><span data-stu-id="6c144-177">A Date object</span></span>|

##### <a name="requirements"></a><span data-ttu-id="6c144-178">Anforderungen</span><span class="sxs-lookup"><span data-stu-id="6c144-178">Requirements</span></span>

|<span data-ttu-id="6c144-179">Anforderung</span><span class="sxs-lookup"><span data-stu-id="6c144-179">Requirement</span></span>| <span data-ttu-id="6c144-180">Wert</span><span class="sxs-lookup"><span data-stu-id="6c144-180">Value</span></span>|
|---|---|
|[<span data-ttu-id="6c144-181">Mindestversion des Postfachanforderungssatzes</span><span class="sxs-lookup"><span data-stu-id="6c144-181">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="6c144-182">1.0</span><span class="sxs-lookup"><span data-stu-id="6c144-182">1.0</span></span>|
|[<span data-ttu-id="6c144-183">Mindestberechtigungsstufe</span><span class="sxs-lookup"><span data-stu-id="6c144-183">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="6c144-184">ReadItem</span><span class="sxs-lookup"><span data-stu-id="6c144-184">ReadItem</span></span>|
|[<span data-ttu-id="6c144-185">Zutreffender Outlook-Modus</span><span class="sxs-lookup"><span data-stu-id="6c144-185">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="6c144-186">Verfassen oder Lesen</span><span class="sxs-lookup"><span data-stu-id="6c144-186">Compose or read</span></span>|

##### <a name="returns"></a><span data-ttu-id="6c144-187">Gibt zurück:</span><span class="sxs-lookup"><span data-stu-id="6c144-187">Returns:</span></span>

<span data-ttu-id="6c144-188">Typ: [LocalClientTime](/javascript/api/outlook_1_3/office.LocalClientTime)</span><span class="sxs-lookup"><span data-stu-id="6c144-188">Type: [LocalClientTime](/javascript/api/outlook_1_3/office.LocalClientTime)</span></span>

####  <a name="converttorestiditemid-restversion--string"></a><span data-ttu-id="6c144-189">convertToRestId(itemId, restVersion) → {String}</span><span class="sxs-lookup"><span data-stu-id="6c144-189">convertToRestId(itemId, restVersion) → {String}</span></span>

<span data-ttu-id="6c144-190">Wandelt eine für EWS formatierte Element-ID ins REST-Format um.</span><span class="sxs-lookup"><span data-stu-id="6c144-190">Converts an item ID formatted for EWS into REST format.</span></span>

> [!NOTE]
> <span data-ttu-id="6c144-191">Diese Methode wird in Outlook für iOS oder Outlook für Android nicht unterstützt.</span><span class="sxs-lookup"><span data-stu-id="6c144-191">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="6c144-p108">Über EWS oder die `itemId` abgerufene Element-IDs verwenden ein anderes Format als die über REST-APIs (z. B. die [Outlook Mail-API](https://docs.microsoft.com/previous-versions/office/office-365-api/api/version-2.0/mail-rest-operations) oder [Microsoft Graph](http://graph.microsoft.io/)) abgerufenen Element-IDs. Die `convertToRestId`-Methode wandelt eine für EWS formatierte ID in das entsprechende REST-Format um.</span><span class="sxs-lookup"><span data-stu-id="6c144-p108">Item IDs retrieved via EWS or via the `itemId` property use a different format than the format used by REST APIs (such as the [Outlook Mail API](https://docs.microsoft.com/previous-versions/office/office-365-api/api/version-2.0/mail-rest-operations) or the [Microsoft Graph](http://graph.microsoft.io/)). The `convertToRestId` method converts an EWS-formatted ID into the proper format for REST.</span></span>

##### <a name="parameters"></a><span data-ttu-id="6c144-194">Parameter:</span><span class="sxs-lookup"><span data-stu-id="6c144-194">Parameters:</span></span>

|<span data-ttu-id="6c144-195">Name</span><span class="sxs-lookup"><span data-stu-id="6c144-195">Name</span></span>| <span data-ttu-id="6c144-196">Typ</span><span class="sxs-lookup"><span data-stu-id="6c144-196">Type</span></span>| <span data-ttu-id="6c144-197">Beschreibung</span><span class="sxs-lookup"><span data-stu-id="6c144-197">Description</span></span>|
|---|---|---|
|`itemId`| <span data-ttu-id="6c144-198">Zeichenfolge</span><span class="sxs-lookup"><span data-stu-id="6c144-198">String</span></span>|<span data-ttu-id="6c144-199">Eine für Exchange-Webdienste (EWS) formatierte Element-ID</span><span class="sxs-lookup"><span data-stu-id="6c144-199">An item ID formatted for Exchange Web Services (EWS)</span></span>|
|`restVersion`| [<span data-ttu-id="6c144-200">Office.MailboxEnums.RestVersion</span><span class="sxs-lookup"><span data-stu-id="6c144-200">Office.MailboxEnums.RestVersion</span></span>](/javascript/api/outlook_1_3/office.mailboxenums.restversion)|<span data-ttu-id="6c144-201">Ein Wert, der die Version der Outlook-REST-API angibt, mit der die konvertierte ID verwendet wird.</span><span class="sxs-lookup"><span data-stu-id="6c144-201">A value indicating the version of the Outlook REST API that the converted ID will be used with.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="6c144-202">Anforderungen</span><span class="sxs-lookup"><span data-stu-id="6c144-202">Requirements</span></span>

|<span data-ttu-id="6c144-203">Anforderung</span><span class="sxs-lookup"><span data-stu-id="6c144-203">Requirement</span></span>| <span data-ttu-id="6c144-204">Wert</span><span class="sxs-lookup"><span data-stu-id="6c144-204">Value</span></span>|
|---|---|
|[<span data-ttu-id="6c144-205">Mindestversion des Postfachanforderungssatzes</span><span class="sxs-lookup"><span data-stu-id="6c144-205">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="6c144-206">1.3</span><span class="sxs-lookup"><span data-stu-id="6c144-206">1.3</span></span>|
|[<span data-ttu-id="6c144-207">Mindestberechtigungsstufe</span><span class="sxs-lookup"><span data-stu-id="6c144-207">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="6c144-208">Eingeschränkt</span><span class="sxs-lookup"><span data-stu-id="6c144-208">Restricted</span></span>|
|[<span data-ttu-id="6c144-209">Zutreffender Outlook-Modus</span><span class="sxs-lookup"><span data-stu-id="6c144-209">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="6c144-210">Verfassen oder Lesen</span><span class="sxs-lookup"><span data-stu-id="6c144-210">Compose or read</span></span>|

##### <a name="returns"></a><span data-ttu-id="6c144-211">Gibt zurück:</span><span class="sxs-lookup"><span data-stu-id="6c144-211">Returns:</span></span>

<span data-ttu-id="6c144-212">Typ: Zeichenfolge</span><span class="sxs-lookup"><span data-stu-id="6c144-212">Type: String</span></span>

##### <a name="example"></a><span data-ttu-id="6c144-213">Beispiel</span><span class="sxs-lookup"><span data-stu-id="6c144-213">Example</span></span>

```
// Get the currently selected item's ID
var ewsId = Office.context.mailbox.item.itemId;

// Convert to a REST ID for the v2.0 version of the
// Outlook Mail API
var restId = Office.context.mailbox.convertToRestId(ewsId, Office.MailboxEnums.RestVersion.v2_0);
```

####  <a name="converttoutcclienttimeinput--date"></a><span data-ttu-id="6c144-214">convertToUtcClientTime(input) → {Date}</span><span class="sxs-lookup"><span data-stu-id="6c144-214">convertToUtcClientTime(input) → {Date}</span></span>

<span data-ttu-id="6c144-215">Ruft ein Date-Objekt aus einem Wörterbuch mit Uhrzeitinformationen ab.</span><span class="sxs-lookup"><span data-stu-id="6c144-215">Gets a Date object from a dictionary containing time information.</span></span>

<span data-ttu-id="6c144-216">Mit der `convertToUtcClientTime`-Methode wird ein Wörterbuch mit lokalem Datum und lokaler Uhrzeit in ein Date-Objekt mit den richtigen Werten für das lokale Datum und die lokale Uhrzeit konvertiert.</span><span class="sxs-lookup"><span data-stu-id="6c144-216">The `convertToUtcClientTime` method converts a dictionary containing a local date and time to a Date object with the correct values for the local date and time.</span></span>

##### <a name="parameters"></a><span data-ttu-id="6c144-217">Parameter:</span><span class="sxs-lookup"><span data-stu-id="6c144-217">Parameters:</span></span>

|<span data-ttu-id="6c144-218">Name</span><span class="sxs-lookup"><span data-stu-id="6c144-218">Name</span></span>| <span data-ttu-id="6c144-219">Typ</span><span class="sxs-lookup"><span data-stu-id="6c144-219">Type</span></span>| <span data-ttu-id="6c144-220">Beschreibung</span><span class="sxs-lookup"><span data-stu-id="6c144-220">Description</span></span>|
|---|---|---|
|`input`| [<span data-ttu-id="6c144-221">LocalClientTime</span><span class="sxs-lookup"><span data-stu-id="6c144-221">LocalClientTime</span></span>](/javascript/api/outlook_1_3/office.LocalClientTime)|<span data-ttu-id="6c144-222">Der zu konvertierende Wert für die lokale Uhrzeit.</span><span class="sxs-lookup"><span data-stu-id="6c144-222">The local time value to convert.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="6c144-223">Anforderungen</span><span class="sxs-lookup"><span data-stu-id="6c144-223">Requirements</span></span>

|<span data-ttu-id="6c144-224">Anforderung</span><span class="sxs-lookup"><span data-stu-id="6c144-224">Requirement</span></span>| <span data-ttu-id="6c144-225">Wert</span><span class="sxs-lookup"><span data-stu-id="6c144-225">Value</span></span>|
|---|---|
|[<span data-ttu-id="6c144-226">Mindestversion des Postfachanforderungssatzes</span><span class="sxs-lookup"><span data-stu-id="6c144-226">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="6c144-227">1.0</span><span class="sxs-lookup"><span data-stu-id="6c144-227">1.0</span></span>|
|[<span data-ttu-id="6c144-228">Mindestberechtigungsstufe</span><span class="sxs-lookup"><span data-stu-id="6c144-228">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="6c144-229">ReadItem</span><span class="sxs-lookup"><span data-stu-id="6c144-229">ReadItem</span></span>|
|[<span data-ttu-id="6c144-230">Zutreffender Outlook-Modus</span><span class="sxs-lookup"><span data-stu-id="6c144-230">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="6c144-231">Verfassen oder Lesen</span><span class="sxs-lookup"><span data-stu-id="6c144-231">Compose or read</span></span>|

##### <a name="returns"></a><span data-ttu-id="6c144-232">Gibt zurück:</span><span class="sxs-lookup"><span data-stu-id="6c144-232">Returns:</span></span>

<span data-ttu-id="6c144-233">Ein Date-Objekt der Uhrzeit in UTC.</span><span class="sxs-lookup"><span data-stu-id="6c144-233">A Date object with the time expressed in UTC.</span></span>

<dl class="param-type"><span data-ttu-id="6c144-234">

<dt>
Typ</dt>


</span><span class="sxs-lookup"><span data-stu-id="6c144-234">

<dt>Type</dt>

</span></span><dd><span data-ttu-id="6c144-235">Datum</span><span class="sxs-lookup"><span data-stu-id="6c144-235">Date</span></span></dd>

</dl>

####  <a name="displayappointmentformitemid"></a><span data-ttu-id="6c144-236">displayAppointmentForm(itemId)</span><span class="sxs-lookup"><span data-stu-id="6c144-236">displayAppointmentForm(itemId)</span></span>

<span data-ttu-id="6c144-237">Zeigt einen bestehenden Kalendertermin an.</span><span class="sxs-lookup"><span data-stu-id="6c144-237">Displays an existing calendar appointment.</span></span>

> [!NOTE]
> <span data-ttu-id="6c144-238">Diese Methode wird in Outlook für iOS oder Outlook für Android nicht unterstützt.</span><span class="sxs-lookup"><span data-stu-id="6c144-238">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="6c144-239">Mit der `displayAppointmentForm`-Methode wird ein vorhandener Kalendertermin auf dem Desktop in einem neuen Fenster oder auf Mobilgeräten in einem Dialogfeld geöffnet.</span><span class="sxs-lookup"><span data-stu-id="6c144-239">The `displayAppointmentForm` method opens an existing calendar appointment in a new window on the desktop or in a dialog box on mobile devices.</span></span>

<span data-ttu-id="6c144-p109">In Outlook for Mac können Sie diese Methode verwenden, um einen einzelnen Termin anzuzeigen, der nicht Teil einer Terminserie ist, oder den Mastertermin einer Terminserie, jedoch keine Instanz der Terminserie. Dies liegt daran, dass Sie in Outlook for Mac nicht auf die Eigenschaften (einschließlich der Element-ID) von Instanzen einer Terminserie zugreifen können.</span><span class="sxs-lookup"><span data-stu-id="6c144-p109">In Outlook for Mac, you can use this method to display a single appointment that is not part of a recurring series, or the master appointment of a recurring series, but you cannot display an instance of the series. This is because in Outlook for Mac, you cannot access the properties (including the item ID) of instances of a recurring series.</span></span>

<span data-ttu-id="6c144-242">In Outlook Web App öffnet diese Methode das angegebene Formular nur, wenn der Textkörper des Formulars nicht größer ist als 32 KB.</span><span class="sxs-lookup"><span data-stu-id="6c144-242">In Outlook Web App, this method opens the specified form only if the body of the form is less than or equal to 32KB number of characters.</span></span>

<span data-ttu-id="6c144-243">Wenn der angegebene Elementbezeichner keinen vorhandenen Termin identifiziert, wird auf dem Clientcomputer oder Gerät ein leerer Bereich geöffnet, und es wird keine Fehlermeldung zurückgegeben.</span><span class="sxs-lookup"><span data-stu-id="6c144-243">If the specified item identifier does not identify an existing appointment, a blank pane opens on the client computer or device, and no error message will be returned.</span></span>

##### <a name="parameters"></a><span data-ttu-id="6c144-244">Parameter:</span><span class="sxs-lookup"><span data-stu-id="6c144-244">Parameters:</span></span>

|<span data-ttu-id="6c144-245">Name</span><span class="sxs-lookup"><span data-stu-id="6c144-245">Name</span></span>| <span data-ttu-id="6c144-246">Typ</span><span class="sxs-lookup"><span data-stu-id="6c144-246">Type</span></span>| <span data-ttu-id="6c144-247">Beschreibung</span><span class="sxs-lookup"><span data-stu-id="6c144-247">Description</span></span>|
|---|---|---|
|`itemId`| <span data-ttu-id="6c144-248">Zeichenfolge</span><span class="sxs-lookup"><span data-stu-id="6c144-248">String</span></span>|<span data-ttu-id="6c144-249">Der EWS-Bezeichner (Exchange Web Services, Exchange-Webdienste) für einen vorhandenen Kalendertermin.</span><span class="sxs-lookup"><span data-stu-id="6c144-249">The Exchange Web Services (EWS) identifier for an existing calendar appointment.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="6c144-250">Anforderungen</span><span class="sxs-lookup"><span data-stu-id="6c144-250">Requirements</span></span>

|<span data-ttu-id="6c144-251">Anforderung</span><span class="sxs-lookup"><span data-stu-id="6c144-251">Requirement</span></span>| <span data-ttu-id="6c144-252">Wert</span><span class="sxs-lookup"><span data-stu-id="6c144-252">Value</span></span>|
|---|---|
|[<span data-ttu-id="6c144-253">Mindestversion des Postfachanforderungssatzes</span><span class="sxs-lookup"><span data-stu-id="6c144-253">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="6c144-254">1.0</span><span class="sxs-lookup"><span data-stu-id="6c144-254">1.0</span></span>|
|[<span data-ttu-id="6c144-255">Mindestberechtigungsstufe</span><span class="sxs-lookup"><span data-stu-id="6c144-255">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="6c144-256">ReadItem</span><span class="sxs-lookup"><span data-stu-id="6c144-256">ReadItem</span></span>|
|[<span data-ttu-id="6c144-257">Zutreffender Outlook-Modus</span><span class="sxs-lookup"><span data-stu-id="6c144-257">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="6c144-258">Verfassen oder Lesen</span><span class="sxs-lookup"><span data-stu-id="6c144-258">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="6c144-259">Beispiel</span><span class="sxs-lookup"><span data-stu-id="6c144-259">Example</span></span>

```
Office.context.mailbox.displayAppointmentForm(appointmentId);
```

####  <a name="displaymessageformitemid"></a><span data-ttu-id="6c144-260">displayMessageForm(itemId)</span><span class="sxs-lookup"><span data-stu-id="6c144-260">displayMessageForm(itemId)</span></span>

<span data-ttu-id="6c144-261">Zeigt eine vorhandene Nachricht an.</span><span class="sxs-lookup"><span data-stu-id="6c144-261">Displays an existing message.</span></span>

> [!NOTE]
> <span data-ttu-id="6c144-262">Diese Methode wird in Outlook für iOS oder Outlook für Android nicht unterstützt.</span><span class="sxs-lookup"><span data-stu-id="6c144-262">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="6c144-263">Die `displayMessageForm`-Methode öffnet eine vorhandene Nachricht in einem neuen Fenster auf dem Desktop bzw. in einem Dialogfeld auf Mobilgeräten.</span><span class="sxs-lookup"><span data-stu-id="6c144-263">The `displayMessageForm` method opens an existing message in a new window on the desktop or in a dialog box on mobile devices.</span></span>

<span data-ttu-id="6c144-264">In Outlook Web App wird mit dieser Methode das angegebene Formular nur dann geöffnet, wenn der Textkörper des Formular Zeichen im Umfang vom maximal 32 KB umfasst.</span><span class="sxs-lookup"><span data-stu-id="6c144-264">In Outlook Web App, this method opens the specified form only if the body of the form is less than or equal to 32 KB number of characters.</span></span>

<span data-ttu-id="6c144-265">Wenn der angegebene Elementbezeichner keine vorhandenen Nachrichten erkennt, wird auf dem Client-Computer keine Nachricht angezeigt, und es werden keine Fehlermeldungen zurückgegeben.</span><span class="sxs-lookup"><span data-stu-id="6c144-265">If the specified item identifier does not identify an existing message, no message will be displayed on the client computer, and no error message will be returned.</span></span>

<span data-ttu-id="6c144-p110">Verwenden Sie `displayMessageForm` nicht mit einem `itemId`-Objekt, das einen Termin darstellt. Verwenden Sie die `displayAppointmentForm`-Methode, um einen vorhandenen Termin anzuzeigen, und `displayNewAppointmentForm`, um ein Formular zum Erstellen eines neuen Termins anzuzeigen.</span><span class="sxs-lookup"><span data-stu-id="6c144-p110">Do not use the `displayMessageForm` with an `itemId` that represents an appointment. Use the `displayAppointmentForm` method to display an existing appointment, and `displayNewAppointmentForm` to display a form to create a new appointment.</span></span>

##### <a name="parameters"></a><span data-ttu-id="6c144-268">Parameter:</span><span class="sxs-lookup"><span data-stu-id="6c144-268">Parameters:</span></span>

|<span data-ttu-id="6c144-269">Name</span><span class="sxs-lookup"><span data-stu-id="6c144-269">Name</span></span>| <span data-ttu-id="6c144-270">Typ</span><span class="sxs-lookup"><span data-stu-id="6c144-270">Type</span></span>| <span data-ttu-id="6c144-271">Beschreibung</span><span class="sxs-lookup"><span data-stu-id="6c144-271">Description</span></span>|
|---|---|---|
|`itemId`| <span data-ttu-id="6c144-272">Zeichenfolge</span><span class="sxs-lookup"><span data-stu-id="6c144-272">String</span></span>|<span data-ttu-id="6c144-273">Der Exchange-Webdienste (EWS) für eine vorhandene Nachricht.</span><span class="sxs-lookup"><span data-stu-id="6c144-273">The Exchange Web Services (EWS) identifier for an existing message.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="6c144-274">Anforderungen</span><span class="sxs-lookup"><span data-stu-id="6c144-274">Requirements</span></span>

|<span data-ttu-id="6c144-275">Anforderung</span><span class="sxs-lookup"><span data-stu-id="6c144-275">Requirement</span></span>| <span data-ttu-id="6c144-276">Wert</span><span class="sxs-lookup"><span data-stu-id="6c144-276">Value</span></span>|
|---|---|
|[<span data-ttu-id="6c144-277">Mindestversion des Postfachanforderungssatzes</span><span class="sxs-lookup"><span data-stu-id="6c144-277">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="6c144-278">1.0</span><span class="sxs-lookup"><span data-stu-id="6c144-278">1.0</span></span>|
|[<span data-ttu-id="6c144-279">Mindestberechtigungsstufe</span><span class="sxs-lookup"><span data-stu-id="6c144-279">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="6c144-280">ReadItem</span><span class="sxs-lookup"><span data-stu-id="6c144-280">ReadItem</span></span>|
|[<span data-ttu-id="6c144-281">Zutreffender Outlook-Modus</span><span class="sxs-lookup"><span data-stu-id="6c144-281">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="6c144-282">Verfassen oder Lesen</span><span class="sxs-lookup"><span data-stu-id="6c144-282">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="6c144-283">Beispiel</span><span class="sxs-lookup"><span data-stu-id="6c144-283">Example</span></span>

```
Office.context.mailbox.displayMessageForm(messageId);
```

#### <a name="displaynewappointmentformparameters"></a><span data-ttu-id="6c144-284">displayNewAppointmentForm(parameters)</span><span class="sxs-lookup"><span data-stu-id="6c144-284">displayNewAppointmentForm(parameters)</span></span>

<span data-ttu-id="6c144-285">Zeigt ein Formular zum Erstellen eines neuen Kalendertermins an.</span><span class="sxs-lookup"><span data-stu-id="6c144-285">Displays a form for creating a new calendar appointment.</span></span>

> [!NOTE]
> <span data-ttu-id="6c144-286">Diese Methode wird in Outlook für iOS oder Outlook für Android nicht unterstützt.</span><span class="sxs-lookup"><span data-stu-id="6c144-286">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="6c144-p111">Mit der `displayNewAppointmentForm`-Methode wird ein Formular geöffnet, mit dem der Benutzer einen neuen Termin oder eine Besprechung erstellen kann. Wenn Parameter angegeben wurden, werden die Felder im Terminformular automatisch mit dem Inhalt der Parameter ausgefüllt.</span><span class="sxs-lookup"><span data-stu-id="6c144-p111">The `displayNewAppointmentForm` method opens a form that enables the user to create a new appointment or meeting. If parameters are specified, the appointment form fields are automatically populated with the contents of the parameters.</span></span>

<span data-ttu-id="6c144-p112">In Outlook Web App und OWA for Devices zeigt diese Methode immer ein Formular mit einem Teilnehmerfeld an. Wenn Sie keine Teilnehmer als Eingabeargumente angeben, zeigt die Methode ein Formular mit einer Schaltfläche **Speichern** an. Wenn Sie Teilnehmer angegeben haben, enthält das Formular die Teilnehmer und eine Schaltfläche **Senden**.</span><span class="sxs-lookup"><span data-stu-id="6c144-p112">In Outlook Web App and OWA for Devices, this method always displays a form with an attendees field. If you do not specify any attendees as input arguments, the method displays a form with a **Save** button. If you have specified attendees, the form would include the attendees and a **Send** button.</span></span>

<span data-ttu-id="6c144-p113">Wenn Teilnehmer oder Ressourcen im `requiredAttendees`-, `optionalAttendees`- oder `resources`-Parameter festgelegt werden, wird mit dieser Methode im Outlook Rich Client und Outlook RT ein Besprechungsformular mit der Schaltfläche **Senden** angezeigt. Wenn keine Teilnehmer angegeben werden, wird mit dieser Methode ein Terminformular mit der Schaltfläche **Speichern & schließen** angezeigt.</span><span class="sxs-lookup"><span data-stu-id="6c144-p113">In the Outlook rich client and Outlook RT, if you specify any attendees or resources in the `requiredAttendees`, `optionalAttendees`, or `resources` parameter, this method displays a meeting form with a **Send** button. If you don't specify any recipients, this method displays an appointment form with a **Save & Close** button.</span></span>

<span data-ttu-id="6c144-294">Wenn einer der Parameter die angegebenen Größenbeschränkungen überschreitet oder wenn ein unbekannter Parametername angegeben wird, wird eine Ausnahme ausgelöst.</span><span class="sxs-lookup"><span data-stu-id="6c144-294">If any of the parameters exceed the specified size limits, or if an unknown parameter name is specified, an exception is thrown.</span></span>

##### <a name="parameters"></a><span data-ttu-id="6c144-295">Parameter:</span><span class="sxs-lookup"><span data-stu-id="6c144-295">Parameters:</span></span>

|<span data-ttu-id="6c144-296">Name</span><span class="sxs-lookup"><span data-stu-id="6c144-296">Name</span></span>| <span data-ttu-id="6c144-297">Typ</span><span class="sxs-lookup"><span data-stu-id="6c144-297">Type</span></span>| <span data-ttu-id="6c144-298">Beschreibung</span><span class="sxs-lookup"><span data-stu-id="6c144-298">Description</span></span>|
|---|---|---|
| `parameters` | <span data-ttu-id="6c144-299">Object</span><span class="sxs-lookup"><span data-stu-id="6c144-299">Object</span></span> | <span data-ttu-id="6c144-300">Ein Wörterbuch mit Parametern, die den neuen Termin beschreiben</span><span class="sxs-lookup"><span data-stu-id="6c144-300">A dictionary of parameters describing the new appointment.</span></span> |
| `parameters.requiredAttendees` | <span data-ttu-id="6c144-301">Array.&lt;String&gt; &#124; Array.&lt;[EmailAddressDetails](/javascript/api/outlook_1_3/office.emailaddressdetails)&gt;</span><span class="sxs-lookup"><span data-stu-id="6c144-301">Array.&lt;String&gt; &#124; Array.&lt;[EmailAddressDetails](/javascript/api/outlook_1_3/office.emailaddressdetails)&gt;</span></span> | <span data-ttu-id="6c144-p114">Ein Array aus Zeichenfolgen mit den E-Mail-Adressen oder ein Array mit einem `EmailAddressDetails`-Objekt für jeden der erforderlichen Teilnehmer für den Termin. Das Array darf maximal 100 Einträge enthalten.</span><span class="sxs-lookup"><span data-stu-id="6c144-p114">An array of strings containing the email addresses or an array containing an `EmailAddressDetails` object for each of the required attendees for the appointment. The array is limited to a maximum of 100 entries.</span></span> |
| `parameters.optionalAttendees` | <span data-ttu-id="6c144-304">Array.&lt;String&gt; &#124; Array.&lt;[EmailAddressDetails](/javascript/api/outlook_1_3/office.emailaddressdetails)&gt;</span><span class="sxs-lookup"><span data-stu-id="6c144-304">Array.&lt;String&gt; &#124; Array.&lt;[EmailAddressDetails](/javascript/api/outlook_1_3/office.emailaddressdetails)&gt;</span></span> | <span data-ttu-id="6c144-p115">Ein Array aus Zeichenfolgen mit den E-Mail-Adressen oder ein Array mit einem Objekt des Typs `EmailAddressDetails` für jeden der optionalen Teilnehmer des Termins. Das Array darf maximal 100 Einträge enthalten.</span><span class="sxs-lookup"><span data-stu-id="6c144-p115">An array of strings containing the email addresses or an array containing an `EmailAddressDetails` object for each of the optional attendees for the appointment. The array is limited to a maximum of 100 entries.</span></span> |
| `parameters.start` | <span data-ttu-id="6c144-307">Datum</span><span class="sxs-lookup"><span data-stu-id="6c144-307">Date</span></span> | <span data-ttu-id="6c144-308">Ein Objekt des Typs `Date`. Gibt das Startdatum und den Beginn des Termins an.</span><span class="sxs-lookup"><span data-stu-id="6c144-308">A `Date` object specifying the start date and time of the appointment.</span></span> |
| `parameters.end` | <span data-ttu-id="6c144-309">Date</span><span class="sxs-lookup"><span data-stu-id="6c144-309">Date</span></span> | <span data-ttu-id="6c144-310">Ein Objekt des Typs `Date`. Gibt das Enddatum und das Ende des Termins an.</span><span class="sxs-lookup"><span data-stu-id="6c144-310">A `Date` object specifying the end date and time of the appointment.</span></span> |
| `parameters.location` | <span data-ttu-id="6c144-311">String</span><span class="sxs-lookup"><span data-stu-id="6c144-311">String</span></span> | <span data-ttu-id="6c144-p116">Eine Zeichenfolge mit dem Ort für den Termin. Die Zeichenfolge ist auf maximal 255 Zeichen beschränkt.</span><span class="sxs-lookup"><span data-stu-id="6c144-p116">A string containing the location of the appointment. The string is limited to a maximum of 255 characters.</span></span> |
| `parameters.resources` | <span data-ttu-id="6c144-314">Array.&lt;String&gt;</span><span class="sxs-lookup"><span data-stu-id="6c144-314">Array.&lt;String&gt;</span></span> | <span data-ttu-id="6c144-p117">Ein Array aus Zeichenfolgen, die die für den Termin erforderlichen Ressourcen enthalten. Das Array darf maximal 100 Einträge enthalten.</span><span class="sxs-lookup"><span data-stu-id="6c144-p117">An array of strings containing the resources required for the appointment. The array is limited to a maximum of 100 entries.</span></span> |
| `parameters.subject` | <span data-ttu-id="6c144-317">Zeichenfolge</span><span class="sxs-lookup"><span data-stu-id="6c144-317">String</span></span> | <span data-ttu-id="6c144-p118">Eine Zeichenfolge mit dem Betreff für den Termin. Die Zeichenfolge ist auf maximal 255 Zeichen beschränkt.</span><span class="sxs-lookup"><span data-stu-id="6c144-p118">A string containing the subject of the appointment. The string is limited to a maximum of 255 characters.</span></span> |
| `parameters.body` | <span data-ttu-id="6c144-320">Zeichenfolge</span><span class="sxs-lookup"><span data-stu-id="6c144-320">String</span></span> | <span data-ttu-id="6c144-p119">Der Text des Termins. Der Textinhalt darf maximal 32 KB umfassen.</span><span class="sxs-lookup"><span data-stu-id="6c144-p119">The body of the appointment. The body content is limited to a maximum size of 32 KB.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="6c144-323">Anforderungen</span><span class="sxs-lookup"><span data-stu-id="6c144-323">Requirements</span></span>

|<span data-ttu-id="6c144-324">Anforderung</span><span class="sxs-lookup"><span data-stu-id="6c144-324">Requirement</span></span>| <span data-ttu-id="6c144-325">Wert</span><span class="sxs-lookup"><span data-stu-id="6c144-325">Value</span></span>|
|---|---|
|[<span data-ttu-id="6c144-326">Mindestversion des Postfachanforderungssatzes</span><span class="sxs-lookup"><span data-stu-id="6c144-326">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="6c144-327">1.0</span><span class="sxs-lookup"><span data-stu-id="6c144-327">1.0</span></span>|
|[<span data-ttu-id="6c144-328">Mindestberechtigungsstufe</span><span class="sxs-lookup"><span data-stu-id="6c144-328">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="6c144-329">ReadItem</span><span class="sxs-lookup"><span data-stu-id="6c144-329">ReadItem</span></span>|
|[<span data-ttu-id="6c144-330">Zutreffender Outlook-Modus</span><span class="sxs-lookup"><span data-stu-id="6c144-330">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="6c144-331">Lesen</span><span class="sxs-lookup"><span data-stu-id="6c144-331">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="6c144-332">Beispiel</span><span class="sxs-lookup"><span data-stu-id="6c144-332">Example</span></span>

```
var start = new Date();
var end = new Date();
end.setHours(start.getHours() + 1);

Office.context.mailbox.displayNewAppointmentForm(
  {
    requiredAttendees: ['bob@contoso.com'],
    optionalAttendees: ['sam@contoso.com'],
    start: start,
    end: end,
    location: 'Home',
    resources: ['projector@contoso.com'],
    subject: 'meeting',
    body: 'Hello World!'
  });
```

#### <a name="getcallbacktokenasynccallback-usercontext"></a><span data-ttu-id="6c144-333">getCallbackTokenAsync(callback, [userContext])</span><span class="sxs-lookup"><span data-stu-id="6c144-333">getCallbackTokenAsync(callback, [userContext])</span></span>

<span data-ttu-id="6c144-334">Ruft eine Zeichenfolge ab, die einen Token enthält, der verwendet wird, um eine Anlage oder ein Element von einem Exchange Server abzurufen.</span><span class="sxs-lookup"><span data-stu-id="6c144-334">Gets a string that contains a token used to get an attachment or item from an Exchange Server.</span></span>

<span data-ttu-id="6c144-p120">Die `getCallbackTokenAsync`-Methode führt einen asynchronen Aufruf zum Abruf eines nicht transparenten Tokens vom Exchange-Server aus, der das Postfach des Benutzers hostet. Die Gültigkeitsdauer des Rückruftokens beträgt 5 Minuten.</span><span class="sxs-lookup"><span data-stu-id="6c144-p120">The `getCallbackTokenAsync` method makes an asynchronous call to get an opaque token from the Exchange Server that hosts the user's mailbox. The lifetime of the callback token is 5 minutes.</span></span>

<span data-ttu-id="6c144-p121">Sie können das Token und einen Anlagen- oder einen Elementbezeichner an ein Drittanbietersystem weitergeben. Das Drittanbietersystem verwendet das Token als Trägerautorisierungstoken, um den EWS-Vorgang (Exchange-Webdienste) [GetAttachment](https://docs.microsoft.com/exchange/client-developer/web-service-reference/getattachment-operation) oder [GetItem](https://docs.microsoft.com/exchange/client-developer/web-service-reference/getitem-operation) aufzurufen und eine Anlage oder ein Element zurückzugeben. Sie können z. B. einen Remotedienst erstellen, um [Anlagen aus dem ausgewählten Element abzurufen](https://docs.microsoft.com/outlook/add-ins/get-attachments-of-an-outlook-item).</span><span class="sxs-lookup"><span data-stu-id="6c144-p121">You can pass the token and an attachment identifier or item identifier to a third-party system. The third-party system uses the token as a bearer authorization token to call the Exchange Web Services (EWS) [GetAttachment](https://docs.microsoft.com/exchange/client-developer/web-service-reference/getattachment-operation) or [GetItem](https://docs.microsoft.com/exchange/client-developer/web-service-reference/getitem-operation) operation to return an attachment or item. For example, you can create a remote service to [get attachments from the selected item](https://docs.microsoft.com/outlook/add-ins/get-attachments-of-an-outlook-item).</span></span>

<span data-ttu-id="6c144-340">Ihre App muss über die im Manifest angegebene **ReadItem**-Berechtigung verfügen, um die `getCallbackTokenAsync`-Methode im Lesemodus aufzurufen.</span><span class="sxs-lookup"><span data-stu-id="6c144-340">Your app must have the **ReadItem** permission specified in its manifest to call the `getCallbackTokenAsync` method in read mode.</span></span>

<span data-ttu-id="6c144-p122">Im Verfassenmodus müssen Sie die [`saveAsync`](Office.context.mailbox.item.md#saveasyncoptions-callback)-Methode aufrufen, um einen Elementbezeichner an die `getCallbackTokenAsync`-Methode zu übergeben. Die App muss über **ReadWriteItem**-Berechtigungen verfügen, um die `saveAsync`-Methode aufzurufen.</span><span class="sxs-lookup"><span data-stu-id="6c144-p122">In compose mode you must call the [`saveAsync`](Office.context.mailbox.item.md#saveasyncoptions-callback) method to get an item identifier to pass to the `getCallbackTokenAsync` method. Your app must have **ReadWriteItem** permissions to call the `saveAsync` method.</span></span>

##### <a name="parameters"></a><span data-ttu-id="6c144-343">Parameter:</span><span class="sxs-lookup"><span data-stu-id="6c144-343">Parameters:</span></span>

|<span data-ttu-id="6c144-344">Name</span><span class="sxs-lookup"><span data-stu-id="6c144-344">Name</span></span>| <span data-ttu-id="6c144-345">Typ</span><span class="sxs-lookup"><span data-stu-id="6c144-345">Type</span></span>| <span data-ttu-id="6c144-346">Attribute</span><span class="sxs-lookup"><span data-stu-id="6c144-346">Attributes</span></span>| <span data-ttu-id="6c144-347">Beschreibung</span><span class="sxs-lookup"><span data-stu-id="6c144-347">Description</span></span>|
|---|---|---|---|
|`callback`| <span data-ttu-id="6c144-348">Funktion</span><span class="sxs-lookup"><span data-stu-id="6c144-348">function</span></span>||<span data-ttu-id="6c144-p123">Wenn die Methode abgeschlossen wird, wird die Funktion , die im `callback`-Parameter übergeben wird, mit einem einzelnen Parameter (`asyncResult`) aufgerufen, der ein [`AsyncResult`](/javascript/api/office/office.asyncresult)-Objekt darstellt. Das Token wird als Zeichenfolge in der `asyncResult.value`-Eigenschaft bereitgestellt.</span><span class="sxs-lookup"><span data-stu-id="6c144-p123">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object. The token is provided as a string in the `asyncResult.value` property.</span></span>|
|`userContext`| <span data-ttu-id="6c144-351">Objekt</span><span class="sxs-lookup"><span data-stu-id="6c144-351">Object</span></span>| <span data-ttu-id="6c144-352">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="6c144-352">&lt;optional&gt;</span></span>|<span data-ttu-id="6c144-353">Jegliche Zustandsdaten, die an die asynchrone Methode übergeben werden.</span><span class="sxs-lookup"><span data-stu-id="6c144-353">Any state data that is passed to the asynchronous method.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="6c144-354">Anforderungen</span><span class="sxs-lookup"><span data-stu-id="6c144-354">Requirements</span></span>

|<span data-ttu-id="6c144-355">Anforderung</span><span class="sxs-lookup"><span data-stu-id="6c144-355">Requirement</span></span>| <span data-ttu-id="6c144-356">Wert</span><span class="sxs-lookup"><span data-stu-id="6c144-356">Value</span></span>|
|---|---|
|[<span data-ttu-id="6c144-357">Mindestversion des Postfachanforderungssatzes</span><span class="sxs-lookup"><span data-stu-id="6c144-357">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="6c144-358">1.3</span><span class="sxs-lookup"><span data-stu-id="6c144-358">1.3</span></span>|
|[<span data-ttu-id="6c144-359">Mindestberechtigungsstufe</span><span class="sxs-lookup"><span data-stu-id="6c144-359">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="6c144-360">ReadItem</span><span class="sxs-lookup"><span data-stu-id="6c144-360">ReadItem</span></span>|
|[<span data-ttu-id="6c144-361">Zutreffender Outlook-Modus</span><span class="sxs-lookup"><span data-stu-id="6c144-361">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="6c144-362">Verfassenmodus und Lesemodus</span><span class="sxs-lookup"><span data-stu-id="6c144-362">Compose and read</span></span>|

##### <a name="example"></a><span data-ttu-id="6c144-363">Beispiel</span><span class="sxs-lookup"><span data-stu-id="6c144-363">Example</span></span>

```js
function getCallbackToken() {
  Office.context.mailbox.getCallbackTokenAsync(cb);
}

function cb(asyncResult) {
  var token = asyncResult.value;
}
```

####  <a name="getuseridentitytokenasynccallback-usercontext"></a><span data-ttu-id="6c144-364">getUserIdentityTokenAsync(callback, [userContext])</span><span class="sxs-lookup"><span data-stu-id="6c144-364">getUserIdentityTokenAsync(callback, [userContext])</span></span>

<span data-ttu-id="6c144-365">Ruft ein Token ab, das den Benutzer und das Office-Add-In identifiziert.</span><span class="sxs-lookup"><span data-stu-id="6c144-365">Gets a token identifying the user and the Office Add-in.</span></span>

<span data-ttu-id="6c144-366">Die `getUserIdentityTokenAsync`-Methode gibt ein Token zurück, das Sie dazu verwenden können, um das [Add-In und den Benutzer mit einem Drittanbietersystem zu identifizieren und zu authentifizieren](https://docs.microsoft.com/outlook/add-ins/authentication).</span><span class="sxs-lookup"><span data-stu-id="6c144-366">The `getUserIdentityTokenAsync` method returns a token that you can use to identify and [authenticate the add-in and user with a third-party system](https://docs.microsoft.com/outlook/add-ins/authentication).</span></span>

##### <a name="parameters"></a><span data-ttu-id="6c144-367">Parameter:</span><span class="sxs-lookup"><span data-stu-id="6c144-367">Parameters:</span></span>

|<span data-ttu-id="6c144-368">Name</span><span class="sxs-lookup"><span data-stu-id="6c144-368">Name</span></span>| <span data-ttu-id="6c144-369">Typ</span><span class="sxs-lookup"><span data-stu-id="6c144-369">Type</span></span>| <span data-ttu-id="6c144-370">Attribute</span><span class="sxs-lookup"><span data-stu-id="6c144-370">Attributes</span></span>| <span data-ttu-id="6c144-371">Beschreibung</span><span class="sxs-lookup"><span data-stu-id="6c144-371">Description</span></span>|
|---|---|---|---|
|`callback`| <span data-ttu-id="6c144-372">Funktion</span><span class="sxs-lookup"><span data-stu-id="6c144-372">function</span></span>||<span data-ttu-id="6c144-373">Wenn die Methode abgeschlossen wird, wird die Funktion , die im `callback`-Parameter übergeben wird, mit einem einzelnen Parameter (`asyncResult`) aufgerufen, der ein [`AsyncResult`](/javascript/api/office/office.asyncresult)-Objekt darstellt.</span><span class="sxs-lookup"><span data-stu-id="6c144-373">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span><br/><br/><span data-ttu-id="6c144-374">Das Token wird als Zeichenfolge in der `asyncResult.value`-Eigenschaft bereitgestellt.</span><span class="sxs-lookup"><span data-stu-id="6c144-374">The token is provided as a string in the `asyncResult.value` property.</span></span>|
|`userContext`| <span data-ttu-id="6c144-375">Objekt</span><span class="sxs-lookup"><span data-stu-id="6c144-375">Object</span></span>| <span data-ttu-id="6c144-376">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="6c144-376">&lt;optional&gt;</span></span>|<span data-ttu-id="6c144-377">Jegliche Zustandsdaten, die an die asynchrone Methode übergeben werden.</span><span class="sxs-lookup"><span data-stu-id="6c144-377">Any state data that is passed to the asynchronous method.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="6c144-378">Anforderungen</span><span class="sxs-lookup"><span data-stu-id="6c144-378">Requirements</span></span>

|<span data-ttu-id="6c144-379">Anforderung</span><span class="sxs-lookup"><span data-stu-id="6c144-379">Requirement</span></span>| <span data-ttu-id="6c144-380">Wert</span><span class="sxs-lookup"><span data-stu-id="6c144-380">Value</span></span>|
|---|---|
|[<span data-ttu-id="6c144-381">Mindestversion des Postfachanforderungssatzes</span><span class="sxs-lookup"><span data-stu-id="6c144-381">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="6c144-382">1.0</span><span class="sxs-lookup"><span data-stu-id="6c144-382">1.0</span></span>|
|[<span data-ttu-id="6c144-383">Mindestberechtigungsstufe</span><span class="sxs-lookup"><span data-stu-id="6c144-383">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="6c144-384">ReadItem</span><span class="sxs-lookup"><span data-stu-id="6c144-384">ReadItem</span></span>|
|[<span data-ttu-id="6c144-385">Zutreffender Outlook-Modus</span><span class="sxs-lookup"><span data-stu-id="6c144-385">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="6c144-386">Verfassen oder Lesen</span><span class="sxs-lookup"><span data-stu-id="6c144-386">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="6c144-387">Beispiel</span><span class="sxs-lookup"><span data-stu-id="6c144-387">Example</span></span>

```js
function getIdentityToken() {
  Office.context.mailbox.getUserIdentityTokenAsync(cb);
}

function cb(asyncResult) {
  var token = asyncResult.value;
}
```

####  <a name="makeewsrequestasyncdata-callback-usercontext"></a><span data-ttu-id="6c144-388">makeEwsRequestAsync(data, callback, [userContext])</span><span class="sxs-lookup"><span data-stu-id="6c144-388">makeEwsRequestAsync(data, callback, [userContext])</span></span>

<span data-ttu-id="6c144-389">Richtet eine asynchrone Anforderung an einen EWS-Dienst (Exchange Web Services) auf dem Exchange-Server, der das Postfach des Benutzers hostet.</span><span class="sxs-lookup"><span data-stu-id="6c144-389">Makes an asynchronous request to an Exchange Web Services (EWS) service on the Exchange server that hosts the user’s mailbox.</span></span>

> [!NOTE]
> <span data-ttu-id="6c144-390">Diese Methode wird in den folgenden Szenarien nicht unterstützt.</span><span class="sxs-lookup"><span data-stu-id="6c144-390">This method is not supported in the following scenarios.</span></span>
> - <span data-ttu-id="6c144-391">In Outlook für iOS oder Outlook für Android</span><span class="sxs-lookup"><span data-stu-id="6c144-391">In Outlook for iOS or Outlook for Android</span></span>
> - <span data-ttu-id="6c144-392">Wenn das Add-In geladen wird, in einem Postfach Google Mail</span><span class="sxs-lookup"><span data-stu-id="6c144-392">When the add-in is loaded in a Gmail mailbox</span></span>
> 
> <span data-ttu-id="6c144-393">In diesen Fällen sollte-add-ins [mithilfe von REST-APIs](https://docs.microsoft.com/outlook/add-ins/use-rest-api) , stattdessen Zugriff auf das Postfach des Benutzers.</span><span class="sxs-lookup"><span data-stu-id="6c144-393">In these cases, add-ins should [use REST APIs](https://docs.microsoft.com/outlook/add-ins/use-rest-api) to access the user's mailbox instead.</span></span>

<span data-ttu-id="6c144-394">Die `makeEwsRequestAsync`-Methode sendet eine EWS-Anforderung für das Add-In zu Exchange.</span><span class="sxs-lookup"><span data-stu-id="6c144-394">The `makeEwsRequestAsync` method sends an EWS request on behalf of the add-in to Exchange.</span></span> <span data-ttu-id="6c144-395">Eine Liste der unterstützten EWS-Vorgänge finden Sie unter [Aufrufen von Webdiensten aus einem Outlook-add-in](https://docs.microsoft.com/outlook/add-ins/web-services#ews-operations-that-add-ins-support) .</span><span class="sxs-lookup"><span data-stu-id="6c144-395">See [Call web services from an Outlook add-in](https://docs.microsoft.com/outlook/add-ins/web-services#ews-operations-that-add-ins-support) for a list of the supported EWS operations.</span></span>

<span data-ttu-id="6c144-396">Sie können keine Elemente, die Ordnern zugeordnet sind, mit der `makeEwsRequestAsync`-Methode anfordern.</span><span class="sxs-lookup"><span data-stu-id="6c144-396">You cannot request Folder Associated Items with the `makeEwsRequestAsync` method.</span></span>

<span data-ttu-id="6c144-397">Die XML-Anfrage muss UTF-8-Codierung angeben.</span><span class="sxs-lookup"><span data-stu-id="6c144-397">The XML request must specify UTF-8 encoding.</span></span>

```
<?xml version="1.0" encoding="utf-8"?>
```

<span data-ttu-id="6c144-p125">Das Add-In muss über die **ReadWriteMailbox**-Berechtigung verfügen, um die `makeEwsRequestAsync`-Methode zu verwenden. Informationen zur **ReadWriteMailbox**-Berechtigung und zu den EWS-Vorgängen, die Sie mit der `makeEwsRequestAsync`-Methode aufrufen können, finden Sie unter [Angeben von Berechtigungen für den E-Mail-Add-In-Zugriff auf das Benutzerpostfach](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions).</span><span class="sxs-lookup"><span data-stu-id="6c144-p125">Your add-in must have the **ReadWriteMailbox** permission to use the `makeEwsRequestAsync` method. For information about using the **ReadWriteMailbox** permission and the EWS operations that you can call with the `makeEwsRequestAsync` method, see [Specify permissions for mail add-in access to the user's mailbox](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions).</span></span>

> [!NOTE]
> <span data-ttu-id="6c144-400">Der Serveradministrator muss festgelegt `OAuthAuthentication` auf "true" auf dem Client Access Server EWS-Verzeichnis so aktivieren Sie die `makeEwsRequestAsync` -Methode, um EWS stellen anfordert.</span><span class="sxs-lookup"><span data-stu-id="6c144-400">The server administrator must set `OAuthAuthentication` to true on the Client Access Server EWS directory to enable the `makeEwsRequestAsync` method to make EWS requests.</span></span>

##### <a name="version-differences"></a><span data-ttu-id="6c144-401">Versionsunterschiede</span><span class="sxs-lookup"><span data-stu-id="6c144-401">Version differences</span></span>

<span data-ttu-id="6c144-402">Wenn Sie die `makeEwsRequestAsync`-Methode in Mail-Apps verwenden, die in älteren Outlook-Versionen als Version 15.0.4535.1004 ausgeführt werden, sollten Sie den Codierungswert auf `ISO-8859-1` festlegen.</span><span class="sxs-lookup"><span data-stu-id="6c144-402">When you use the `makeEwsRequestAsync` method in mail apps running in Outlook versions earlier than version 15.0.4535.1004, you should set the encoding value to `ISO-8859-1`.</span></span>

```
<?xml version="1.0" encoding="iso-8859-1"?>
```

<span data-ttu-id="6c144-p126">Sie müssen den Codierungswert nicht festlegen, wenn Ihre Mail-App in Outlook im Web ausgeführt wird. Sie können mithilfe der mailbox.diagnostics.hostName-Eigenschaft ermitteln, ob Ihre Mail-App in Outlook oder Outlook im Web ausgeführt wird. Sie können mithilfe der mailbox.diagnostics.hostVersion-Eigenschaft ermitteln, welche Version von Outlook ausgeführt wird.</span><span class="sxs-lookup"><span data-stu-id="6c144-p126">You do not need to set the encoding value when your mail app is running in Outlook on the web. You can determine whether your mail app is running in Outlook or Outlook on the web by using the mailbox.diagnostics.hostName property. You can determine what version of Outlook is running by using the mailbox.diagnostics.hostVersion property.</span></span>

##### <a name="parameters"></a><span data-ttu-id="6c144-406">Parameter:</span><span class="sxs-lookup"><span data-stu-id="6c144-406">Parameters:</span></span>

|<span data-ttu-id="6c144-407">Name</span><span class="sxs-lookup"><span data-stu-id="6c144-407">Name</span></span>| <span data-ttu-id="6c144-408">Typ</span><span class="sxs-lookup"><span data-stu-id="6c144-408">Type</span></span>| <span data-ttu-id="6c144-409">Attribute</span><span class="sxs-lookup"><span data-stu-id="6c144-409">Attributes</span></span>| <span data-ttu-id="6c144-410">Beschreibung</span><span class="sxs-lookup"><span data-stu-id="6c144-410">Description</span></span>|
|---|---|---|---|
|`data`| <span data-ttu-id="6c144-411">Zeichenfolge</span><span class="sxs-lookup"><span data-stu-id="6c144-411">String</span></span>||<span data-ttu-id="6c144-412">Die EWS-Anforderung.</span><span class="sxs-lookup"><span data-stu-id="6c144-412">The EWS request.</span></span>|
|`callback`| <span data-ttu-id="6c144-413">Funktion</span><span class="sxs-lookup"><span data-stu-id="6c144-413">function</span></span>||<span data-ttu-id="6c144-414">Wenn die Methode abgeschlossen wird, wird die Funktion , die im `callback`-Parameter übergeben wird, mit einem einzelnen Parameter (`asyncResult`) aufgerufen, der ein [`AsyncResult`](/javascript/api/office/office.asyncresult)-Objekt darstellt.</span><span class="sxs-lookup"><span data-stu-id="6c144-414">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span><br/><br/><span data-ttu-id="6c144-415">Das XML-Ergebnis des EWS-Aufrufs wird als Zeichenfolge in der `asyncResult.value`-Eigenschaft bereitgestellt.</span><span class="sxs-lookup"><span data-stu-id="6c144-415">The XML result of the EWS call is provided as a string in the `asyncResult.value` property.</span></span> <span data-ttu-id="6c144-416">Wenn das Ergebnis 1 MB Größe überschreitet, wird stattdessen eine Fehlermeldung zurückgegeben.</span><span class="sxs-lookup"><span data-stu-id="6c144-416">If the result exceeds 1 MB in size, an error message is returned instead.</span></span>|
|`userContext`| <span data-ttu-id="6c144-417">Objekt</span><span class="sxs-lookup"><span data-stu-id="6c144-417">Object</span></span>| <span data-ttu-id="6c144-418">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="6c144-418">&lt;optional&gt;</span></span>|<span data-ttu-id="6c144-419">Jegliche Zustandsdaten, die an die asynchrone Methode übergeben werden.</span><span class="sxs-lookup"><span data-stu-id="6c144-419">Any state data that is passed to the asynchronous method.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="6c144-420">Anforderungen</span><span class="sxs-lookup"><span data-stu-id="6c144-420">Requirements</span></span>

|<span data-ttu-id="6c144-421">Anforderung</span><span class="sxs-lookup"><span data-stu-id="6c144-421">Requirement</span></span>| <span data-ttu-id="6c144-422">Wert</span><span class="sxs-lookup"><span data-stu-id="6c144-422">Value</span></span>|
|---|---|
|[<span data-ttu-id="6c144-423">Mindestversion des Postfachanforderungssatzes</span><span class="sxs-lookup"><span data-stu-id="6c144-423">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="6c144-424">1.0</span><span class="sxs-lookup"><span data-stu-id="6c144-424">1.0</span></span>|
|[<span data-ttu-id="6c144-425">Mindestberechtigungsstufe</span><span class="sxs-lookup"><span data-stu-id="6c144-425">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="6c144-426">ReadWriteMailbox</span><span class="sxs-lookup"><span data-stu-id="6c144-426">ReadWriteMailbox</span></span>|
|[<span data-ttu-id="6c144-427">Zutreffender Outlook-Modus</span><span class="sxs-lookup"><span data-stu-id="6c144-427">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="6c144-428">Verfassen oder Lesen</span><span class="sxs-lookup"><span data-stu-id="6c144-428">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="6c144-429">Beispiel</span><span class="sxs-lookup"><span data-stu-id="6c144-429">Example</span></span>

<span data-ttu-id="6c144-430">Das folgende Beispiel ruft `makeEwsRequestAsync` zum Verwenden des `GetItem`-Vorgangs auf, um den Betreff eines Elements abzurufen.</span><span class="sxs-lookup"><span data-stu-id="6c144-430">The following example calls `makeEwsRequestAsync` to use the `GetItem` operation to get the subject of an item.</span></span>

```js
function getSubjectRequest(id) {
   // Return a GetItem operation request for the subject of the specified item.
   var request =
    '<?xml version="1.0" encoding="utf-8"?>' +
    '<soap:Envelope xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance"' +
    '               xmlns:xsd="http://www.w3.org/2001/XMLSchema"' +
    '               xmlns:soap="http://schemas.xmlsoap.org/soap/envelope/"' +
    '               xmlns:t="http://schemas.microsoft.com/exchange/services/2006/types">' +
    '  <soap:Header>' +
    '    <RequestServerVersion Version="Exchange2013" xmlns="http://schemas.microsoft.com/exchange/services/2006/types" soap:mustUnderstand="0" />' +
    '  </soap:Header>' +
    '  <soap:Body>' +
    '    <GetItem xmlns="http://schemas.microsoft.com/exchange/services/2006/messages">' +
    '      <ItemShape>' +
    '        <t:BaseShape>IdOnly</t:BaseShape>' +
    '        <t:AdditionalProperties>' +
    '            <t:FieldURI FieldURI="item:Subject"/>' +
    '        </t:AdditionalProperties>' +
    '      </ItemShape>' +
    '      <ItemIds><t:ItemId Id="' + id + '"/></ItemIds>' +
    '    </GetItem>' +
    '  </soap:Body>' +
    '</soap:Envelope>';

   return request;
}

function sendRequest() {
   // Create a local variable that contains the mailbox.
   Office.context.mailbox.makeEwsRequestAsync(
    getSubjectRequest(mailbox.item.itemId), callback);
}

function callback(asyncResult)  {
   var result = asyncResult.value;
   var context = asyncResult.asyncContext;

   // Process the returned response here.
}
```