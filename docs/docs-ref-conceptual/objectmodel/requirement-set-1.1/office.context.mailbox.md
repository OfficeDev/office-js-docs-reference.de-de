
# <a name="mailbox"></a><span data-ttu-id="83f86-101">mailbox</span><span class="sxs-lookup"><span data-stu-id="83f86-101">mailbox</span></span>

### <span data-ttu-id="83f86-p101">[Office](Office.md)[.context](Office.context.md). mailbox</span><span class="sxs-lookup"><span data-stu-id="83f86-p101">[Office](Office.md)[.context](Office.context.md). mailbox</span></span>

<span data-ttu-id="83f86-104">Ermöglicht den Zugriff auf das Objektmodell von Outlook-add-in für Microsoft Outlook und Microsoft Outlook im Web.</span><span class="sxs-lookup"><span data-stu-id="83f86-104">Provides access to the Outlook add-in object model for Microsoft Outlook and Microsoft Outlook on the web.</span></span>

##### <a name="requirements"></a><span data-ttu-id="83f86-105">Anforderungen</span><span class="sxs-lookup"><span data-stu-id="83f86-105">Requirements</span></span>

|<span data-ttu-id="83f86-106">Anforderung</span><span class="sxs-lookup"><span data-stu-id="83f86-106">Requirement</span></span>| <span data-ttu-id="83f86-107">Wert</span><span class="sxs-lookup"><span data-stu-id="83f86-107">Value</span></span>|
|---|---|
|[<span data-ttu-id="83f86-108">Mindestversion des Postfachanforderungssatzes</span><span class="sxs-lookup"><span data-stu-id="83f86-108">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="83f86-109">1.0</span><span class="sxs-lookup"><span data-stu-id="83f86-109">1.0</span></span>|
|[<span data-ttu-id="83f86-110">Mindestberechtigungsstufe</span><span class="sxs-lookup"><span data-stu-id="83f86-110">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="83f86-111">Eingeschränkt</span><span class="sxs-lookup"><span data-stu-id="83f86-111">Restricted</span></span>|
|[<span data-ttu-id="83f86-112">Zutreffender Outlook-Modus</span><span class="sxs-lookup"><span data-stu-id="83f86-112">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="83f86-113">Verfassen oder Lesen</span><span class="sxs-lookup"><span data-stu-id="83f86-113">Compose or read</span></span>|

### <a name="namespaces"></a><span data-ttu-id="83f86-114">Namespaces</span><span class="sxs-lookup"><span data-stu-id="83f86-114">Namespaces</span></span>

<span data-ttu-id="83f86-115">[diagnostics](Office.context.mailbox.diagnostics.md): Stellt einem Outlook-Add-In Diagnoseinformationen bereit.</span><span class="sxs-lookup"><span data-stu-id="83f86-115">[diagnostics](Office.context.mailbox.diagnostics.md): Provides diagnostic information to an Outlook add-in.</span></span>

<span data-ttu-id="83f86-116">[item](Office.context.mailbox.item.md): Stellt Methoden und Eigenschaften für den Zugriff auf eine Nachricht oder einen Termins in einem Outlook-Add-In bereit.</span><span class="sxs-lookup"><span data-stu-id="83f86-116">[item](Office.context.mailbox.item.md): Provides methods and properties for accessing a message or appointment in an Outlook add-in.</span></span>

<span data-ttu-id="83f86-117">[userProfile](Office.context.mailbox.userProfile.md): Stellt Informationen zum Benutzer in einem Outlook-Add-In bereit.</span><span class="sxs-lookup"><span data-stu-id="83f86-117">[userProfile](Office.context.mailbox.userProfile.md): Provides information about the user in an Outlook add-in.</span></span>

### <a name="members"></a><span data-ttu-id="83f86-118">Elemente</span><span class="sxs-lookup"><span data-stu-id="83f86-118">Members</span></span>

#### <a name="ewsurl-string"></a><span data-ttu-id="83f86-119">ewsUrl :String</span><span class="sxs-lookup"><span data-stu-id="83f86-119">ewsUrl :String</span></span>

<span data-ttu-id="83f86-p102">Ruft die URL des EWS-Endpunkts (Exchange Web Services) für dieses E-Mail-Konto ab. Nur Lesemodus.</span><span class="sxs-lookup"><span data-stu-id="83f86-p102">Gets the URL of the Exchange Web Services (EWS) endpoint for this email account. Read mode only.</span></span>

> [!NOTE]
> <span data-ttu-id="83f86-122">Dieser Member wird in Outlook für iOS oder Outlook für Android nicht unterstützt.</span><span class="sxs-lookup"><span data-stu-id="83f86-122">This member is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="83f86-p103">Der `ewsUrl`-Wert kann durch einen Remotedienst zum Ausführen von EWS-Aufrufen an das Postfach des Benutzers verwendet werden. Sie können z. B. einen Remotedienst erstellen, um [Anlagen](https://docs.microsoft.com/outlook/add-ins/get-attachments-of-an-outlook-item) aus dem ausgewählten Element abzurufen.</span><span class="sxs-lookup"><span data-stu-id="83f86-p103">The `ewsUrl` value can be used by a remote service to make EWS calls to the user's mailbox. For example, you can create a remote service to [get attachments from the selected item](https://docs.microsoft.com/outlook/add-ins/get-attachments-of-an-outlook-item).</span></span>

##### <a name="type"></a><span data-ttu-id="83f86-125">Typ:</span><span class="sxs-lookup"><span data-stu-id="83f86-125">Type:</span></span>

*   <span data-ttu-id="83f86-126">String</span><span class="sxs-lookup"><span data-stu-id="83f86-126">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="83f86-127">Anforderungen</span><span class="sxs-lookup"><span data-stu-id="83f86-127">Requirements</span></span>

|<span data-ttu-id="83f86-128">Anforderung</span><span class="sxs-lookup"><span data-stu-id="83f86-128">Requirement</span></span>| <span data-ttu-id="83f86-129">Wert</span><span class="sxs-lookup"><span data-stu-id="83f86-129">Value</span></span>|
|---|---|
|[<span data-ttu-id="83f86-130">Mindestversion des Postfachanforderungssatzes</span><span class="sxs-lookup"><span data-stu-id="83f86-130">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="83f86-131">1.0</span><span class="sxs-lookup"><span data-stu-id="83f86-131">1.0</span></span>|
|[<span data-ttu-id="83f86-132">Mindestberechtigungsstufe</span><span class="sxs-lookup"><span data-stu-id="83f86-132">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="83f86-133">ReadItem</span><span class="sxs-lookup"><span data-stu-id="83f86-133">ReadItem</span></span>|
|[<span data-ttu-id="83f86-134">Zutreffender Outlook-Modus</span><span class="sxs-lookup"><span data-stu-id="83f86-134">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="83f86-135">Lesen</span><span class="sxs-lookup"><span data-stu-id="83f86-135">Read</span></span>|

### <a name="methods"></a><span data-ttu-id="83f86-136">Methoden</span><span class="sxs-lookup"><span data-stu-id="83f86-136">Methods</span></span>

####  <a name="converttolocalclienttimetimevalue--localclienttimejavascriptapioutlook11officelocalclienttime"></a><span data-ttu-id="83f86-137">convertToLocalClientTime(timeValue) → {[LocalClientTime](/javascript/api/outlook_1_1/office.LocalClientTime)}</span><span class="sxs-lookup"><span data-stu-id="83f86-137">convertToLocalClientTime(timeValue) → {[LocalClientTime](/javascript/api/outlook_1_1/office.LocalClientTime)}</span></span>

<span data-ttu-id="83f86-138">Ruft ein Wörterbuch mit Uhrzeitinformationen basierend auf der Zeiteinstellung des lokalen Clients ab.</span><span class="sxs-lookup"><span data-stu-id="83f86-138">Gets a dictionary containing time information in local client time.</span></span>

<span data-ttu-id="83f86-p104">Die in einer Mail-App für Outlook oder Outlook Web App verwendeten Daten und Uhrzeiten stammen u. U. aus verschiedenen Zeitzonen. Outlook verwendet die Zeitzone des Client-Computers; Outlook Web App verwendet die im Exchange Admin Center (EAC) festgelegte Zeitzone. Sie sollten Datums- und Uhrzeitwerte bearbeiten, damit die auf der Benutzeroberfläche angezeigten Werte immer den von Benutzer erwarteten Zeitzonen entsprechen.</span><span class="sxs-lookup"><span data-stu-id="83f86-p104">The dates and times used by a mail app for Outlook or Outlook Web App can use different time zones. Outlook uses the client computer time zone; Outlook Web App uses the time zone set on the Exchange Admin Center (EAC). You should handle date and time values so that the values you display on the user interface are always consistent with the time zone that the user expects.</span></span>

<span data-ttu-id="83f86-p105">Wird die Mail-App in Outlook ausgeführt, wird mit der `convertToLocalClientTime`-Methode ein Wörterbuchobjekt zurückgegeben, dessen Werte auf die Zeitzone des Clientcomputers festgelegt sind. Wird die Mail-App in Outlook ausgeführt, wird mit der `convertToLocalClientTime`-Methode ein Wörterbuchobjekt zurückgegeben, dessen Werte auf die Zeitzone des Clientcomputers festgelegt sind.</span><span class="sxs-lookup"><span data-stu-id="83f86-p105">If the mail app is running in Outlook, the `convertToLocalClientTime` method will return a dictionary object with the values set to the client computer time zone. If the mail app is running in Outlook Web App, the `convertToLocalClientTime` method will return a dictionary object with the values set to the time zone specified in the EAC.</span></span>

##### <a name="parameters"></a><span data-ttu-id="83f86-144">Parameter:</span><span class="sxs-lookup"><span data-stu-id="83f86-144">Parameters:</span></span>

|<span data-ttu-id="83f86-145">Name</span><span class="sxs-lookup"><span data-stu-id="83f86-145">Name</span></span>| <span data-ttu-id="83f86-146">Typ</span><span class="sxs-lookup"><span data-stu-id="83f86-146">Type</span></span>| <span data-ttu-id="83f86-147">Beschreibung</span><span class="sxs-lookup"><span data-stu-id="83f86-147">Description</span></span>|
|---|---|---|
|`timeValue`| <span data-ttu-id="83f86-148">Datum</span><span class="sxs-lookup"><span data-stu-id="83f86-148">Date</span></span>|<span data-ttu-id="83f86-149">Ein Date-Objekt</span><span class="sxs-lookup"><span data-stu-id="83f86-149">A Date object</span></span>|

##### <a name="requirements"></a><span data-ttu-id="83f86-150">Anforderungen</span><span class="sxs-lookup"><span data-stu-id="83f86-150">Requirements</span></span>

|<span data-ttu-id="83f86-151">Anforderung</span><span class="sxs-lookup"><span data-stu-id="83f86-151">Requirement</span></span>| <span data-ttu-id="83f86-152">Wert</span><span class="sxs-lookup"><span data-stu-id="83f86-152">Value</span></span>|
|---|---|
|[<span data-ttu-id="83f86-153">Mindestversion des Postfachanforderungssatzes</span><span class="sxs-lookup"><span data-stu-id="83f86-153">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="83f86-154">1.0</span><span class="sxs-lookup"><span data-stu-id="83f86-154">1.0</span></span>|
|[<span data-ttu-id="83f86-155">Mindestberechtigungsstufe</span><span class="sxs-lookup"><span data-stu-id="83f86-155">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="83f86-156">ReadItem</span><span class="sxs-lookup"><span data-stu-id="83f86-156">ReadItem</span></span>|
|[<span data-ttu-id="83f86-157">Zutreffender Outlook-Modus</span><span class="sxs-lookup"><span data-stu-id="83f86-157">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="83f86-158">Verfassen oder Lesen</span><span class="sxs-lookup"><span data-stu-id="83f86-158">Compose or read</span></span>|

##### <a name="returns"></a><span data-ttu-id="83f86-159">Gibt zurück:</span><span class="sxs-lookup"><span data-stu-id="83f86-159">Returns:</span></span>

<span data-ttu-id="83f86-160">Typ: [LocalClientTime](/javascript/api/outlook_1_1/office.LocalClientTime)</span><span class="sxs-lookup"><span data-stu-id="83f86-160">Type: [LocalClientTime](/javascript/api/outlook_1_1/office.LocalClientTime)</span></span>

####  <a name="converttoutcclienttimeinput--date"></a><span data-ttu-id="83f86-161">convertToUtcClientTime(input) → {Date}</span><span class="sxs-lookup"><span data-stu-id="83f86-161">convertToUtcClientTime(input) → {Date}</span></span>

<span data-ttu-id="83f86-162">Ruft ein Date-Objekt aus einem Wörterbuch mit Uhrzeitinformationen ab.</span><span class="sxs-lookup"><span data-stu-id="83f86-162">Gets a Date object from a dictionary containing time information.</span></span>

<span data-ttu-id="83f86-163">Mit der `convertToUtcClientTime`-Methode wird ein Wörterbuch mit lokalem Datum und lokaler Uhrzeit in ein Date-Objekt mit den richtigen Werten für das lokale Datum und die lokale Uhrzeit konvertiert.</span><span class="sxs-lookup"><span data-stu-id="83f86-163">The `convertToUtcClientTime` method converts a dictionary containing a local date and time to a Date object with the correct values for the local date and time.</span></span>

##### <a name="parameters"></a><span data-ttu-id="83f86-164">Parameter:</span><span class="sxs-lookup"><span data-stu-id="83f86-164">Parameters:</span></span>

|<span data-ttu-id="83f86-165">Name</span><span class="sxs-lookup"><span data-stu-id="83f86-165">Name</span></span>| <span data-ttu-id="83f86-166">Typ</span><span class="sxs-lookup"><span data-stu-id="83f86-166">Type</span></span>| <span data-ttu-id="83f86-167">Beschreibung</span><span class="sxs-lookup"><span data-stu-id="83f86-167">Description</span></span>|
|---|---|---|
|`input`| [<span data-ttu-id="83f86-168">LocalClientTime</span><span class="sxs-lookup"><span data-stu-id="83f86-168">LocalClientTime</span></span>](/javascript/api/outlook_1_1/office.LocalClientTime)|<span data-ttu-id="83f86-169">Der zu konvertierende Wert für die lokale Uhrzeit.</span><span class="sxs-lookup"><span data-stu-id="83f86-169">The local time value to convert.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="83f86-170">Anforderungen</span><span class="sxs-lookup"><span data-stu-id="83f86-170">Requirements</span></span>

|<span data-ttu-id="83f86-171">Anforderung</span><span class="sxs-lookup"><span data-stu-id="83f86-171">Requirement</span></span>| <span data-ttu-id="83f86-172">Wert</span><span class="sxs-lookup"><span data-stu-id="83f86-172">Value</span></span>|
|---|---|
|[<span data-ttu-id="83f86-173">Mindestversion des Postfachanforderungssatzes</span><span class="sxs-lookup"><span data-stu-id="83f86-173">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="83f86-174">1.0</span><span class="sxs-lookup"><span data-stu-id="83f86-174">1.0</span></span>|
|[<span data-ttu-id="83f86-175">Mindestberechtigungsstufe</span><span class="sxs-lookup"><span data-stu-id="83f86-175">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="83f86-176">ReadItem</span><span class="sxs-lookup"><span data-stu-id="83f86-176">ReadItem</span></span>|
|[<span data-ttu-id="83f86-177">Zutreffender Outlook-Modus</span><span class="sxs-lookup"><span data-stu-id="83f86-177">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="83f86-178">Verfassen oder Lesen</span><span class="sxs-lookup"><span data-stu-id="83f86-178">Compose or read</span></span>|

##### <a name="returns"></a><span data-ttu-id="83f86-179">Gibt zurück:</span><span class="sxs-lookup"><span data-stu-id="83f86-179">Returns:</span></span>

<span data-ttu-id="83f86-180">Ein Date-Objekt der Uhrzeit in UTC.</span><span class="sxs-lookup"><span data-stu-id="83f86-180">A Date object with the time expressed in UTC.</span></span>

<dl class="param-type"><span data-ttu-id="83f86-181">

<dt>
Typ</dt>


</span><span class="sxs-lookup"><span data-stu-id="83f86-181">

<dt>Type</dt>

</span></span><dd><span data-ttu-id="83f86-182">Datum</span><span class="sxs-lookup"><span data-stu-id="83f86-182">Date</span></span></dd>

</dl>

####  <a name="displayappointmentformitemid"></a><span data-ttu-id="83f86-183">displayAppointmentForm(itemId)</span><span class="sxs-lookup"><span data-stu-id="83f86-183">displayAppointmentForm(itemId)</span></span>

<span data-ttu-id="83f86-184">Zeigt einen bestehenden Kalendertermin an.</span><span class="sxs-lookup"><span data-stu-id="83f86-184">Displays an existing calendar appointment.</span></span>

> [!NOTE]
> <span data-ttu-id="83f86-185">Diese Methode wird in Outlook für iOS oder Outlook für Android nicht unterstützt.</span><span class="sxs-lookup"><span data-stu-id="83f86-185">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="83f86-186">Mit der `displayAppointmentForm`-Methode wird ein vorhandener Kalendertermin auf dem Desktop in einem neuen Fenster oder auf Mobilgeräten in einem Dialogfeld geöffnet.</span><span class="sxs-lookup"><span data-stu-id="83f86-186">The `displayAppointmentForm` method opens an existing calendar appointment in a new window on the desktop or in a dialog box on mobile devices.</span></span>

<span data-ttu-id="83f86-p106">In Outlook for Mac können Sie diese Methode verwenden, um einen einzelnen Termin anzuzeigen, der nicht Teil einer Terminserie ist, oder den Mastertermin einer Terminserie, jedoch keine Instanz der Terminserie. Dies liegt daran, dass Sie in Outlook for Mac nicht auf die Eigenschaften (einschließlich der Element-ID) von Instanzen einer Terminserie zugreifen können.</span><span class="sxs-lookup"><span data-stu-id="83f86-p106">In Outlook for Mac, you can use this method to display a single appointment that is not part of a recurring series, or the master appointment of a recurring series, but you cannot display an instance of the series. This is because in Outlook for Mac, you cannot access the properties (including the item ID) of instances of a recurring series.</span></span>

<span data-ttu-id="83f86-189">In Outlook Web App öffnet diese Methode das angegebene Formular nur, wenn der Textkörper des Formulars nicht größer ist als 32 KB.</span><span class="sxs-lookup"><span data-stu-id="83f86-189">In Outlook Web App, this method opens the specified form only if the body of the form is less than or equal to 32KB number of characters.</span></span>

<span data-ttu-id="83f86-190">Wenn der angegebene Elementbezeichner keinen vorhandenen Termin identifiziert, wird auf dem Clientcomputer oder Gerät ein leerer Bereich geöffnet, und es wird keine Fehlermeldung zurückgegeben.</span><span class="sxs-lookup"><span data-stu-id="83f86-190">If the specified item identifier does not identify an existing appointment, a blank pane opens on the client computer or device, and no error message will be returned.</span></span>

##### <a name="parameters"></a><span data-ttu-id="83f86-191">Parameter:</span><span class="sxs-lookup"><span data-stu-id="83f86-191">Parameters:</span></span>

|<span data-ttu-id="83f86-192">Name</span><span class="sxs-lookup"><span data-stu-id="83f86-192">Name</span></span>| <span data-ttu-id="83f86-193">Typ</span><span class="sxs-lookup"><span data-stu-id="83f86-193">Type</span></span>| <span data-ttu-id="83f86-194">Beschreibung</span><span class="sxs-lookup"><span data-stu-id="83f86-194">Description</span></span>|
|---|---|---|
|`itemId`| <span data-ttu-id="83f86-195">Zeichenfolge</span><span class="sxs-lookup"><span data-stu-id="83f86-195">String</span></span>|<span data-ttu-id="83f86-196">Der EWS-Bezeichner (Exchange Web Services, Exchange-Webdienste) für einen vorhandenen Kalendertermin.</span><span class="sxs-lookup"><span data-stu-id="83f86-196">The Exchange Web Services (EWS) identifier for an existing calendar appointment.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="83f86-197">Anforderungen</span><span class="sxs-lookup"><span data-stu-id="83f86-197">Requirements</span></span>

|<span data-ttu-id="83f86-198">Anforderung</span><span class="sxs-lookup"><span data-stu-id="83f86-198">Requirement</span></span>| <span data-ttu-id="83f86-199">Wert</span><span class="sxs-lookup"><span data-stu-id="83f86-199">Value</span></span>|
|---|---|
|[<span data-ttu-id="83f86-200">Mindestversion des Postfachanforderungssatzes</span><span class="sxs-lookup"><span data-stu-id="83f86-200">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="83f86-201">1.0</span><span class="sxs-lookup"><span data-stu-id="83f86-201">1.0</span></span>|
|[<span data-ttu-id="83f86-202">Mindestberechtigungsstufe</span><span class="sxs-lookup"><span data-stu-id="83f86-202">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="83f86-203">ReadItem</span><span class="sxs-lookup"><span data-stu-id="83f86-203">ReadItem</span></span>|
|[<span data-ttu-id="83f86-204">Zutreffender Outlook-Modus</span><span class="sxs-lookup"><span data-stu-id="83f86-204">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="83f86-205">Verfassen oder Lesen</span><span class="sxs-lookup"><span data-stu-id="83f86-205">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="83f86-206">Beispiel</span><span class="sxs-lookup"><span data-stu-id="83f86-206">Example</span></span>

```
Office.context.mailbox.displayAppointmentForm(appointmentId);
```

####  <a name="displaymessageformitemid"></a><span data-ttu-id="83f86-207">displayMessageForm(itemId)</span><span class="sxs-lookup"><span data-stu-id="83f86-207">displayMessageForm(itemId)</span></span>

<span data-ttu-id="83f86-208">Zeigt eine vorhandene Nachricht an.</span><span class="sxs-lookup"><span data-stu-id="83f86-208">Displays an existing message.</span></span>

> [!NOTE]
> <span data-ttu-id="83f86-209">Diese Methode wird in Outlook für iOS oder Outlook für Android nicht unterstützt.</span><span class="sxs-lookup"><span data-stu-id="83f86-209">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="83f86-210">Die `displayMessageForm`-Methode öffnet eine vorhandene Nachricht in einem neuen Fenster auf dem Desktop bzw. in einem Dialogfeld auf Mobilgeräten.</span><span class="sxs-lookup"><span data-stu-id="83f86-210">The `displayMessageForm` method opens an existing message in a new window on the desktop or in a dialog box on mobile devices.</span></span>

<span data-ttu-id="83f86-211">In Outlook Web App wird mit dieser Methode das angegebene Formular nur dann geöffnet, wenn der Textkörper des Formular Zeichen im Umfang vom maximal 32 KB umfasst.</span><span class="sxs-lookup"><span data-stu-id="83f86-211">In Outlook Web App, this method opens the specified form only if the body of the form is less than or equal to 32 KB number of characters.</span></span>

<span data-ttu-id="83f86-212">Wenn der angegebene Elementbezeichner keine vorhandenen Nachrichten erkennt, wird auf dem Client-Computer keine Nachricht angezeigt, und es werden keine Fehlermeldungen zurückgegeben.</span><span class="sxs-lookup"><span data-stu-id="83f86-212">If the specified item identifier does not identify an existing message, no message will be displayed on the client computer, and no error message will be returned.</span></span>

<span data-ttu-id="83f86-p107">Verwenden Sie `displayMessageForm` nicht mit einem `itemId`-Objekt, das einen Termin darstellt. Verwenden Sie die `displayAppointmentForm`-Methode, um einen vorhandenen Termin anzuzeigen, und `displayNewAppointmentForm`, um ein Formular zum Erstellen eines neuen Termins anzuzeigen.</span><span class="sxs-lookup"><span data-stu-id="83f86-p107">Do not use the `displayMessageForm` with an `itemId` that represents an appointment. Use the `displayAppointmentForm` method to display an existing appointment, and `displayNewAppointmentForm` to display a form to create a new appointment.</span></span>

##### <a name="parameters"></a><span data-ttu-id="83f86-215">Parameter:</span><span class="sxs-lookup"><span data-stu-id="83f86-215">Parameters:</span></span>

|<span data-ttu-id="83f86-216">Name</span><span class="sxs-lookup"><span data-stu-id="83f86-216">Name</span></span>| <span data-ttu-id="83f86-217">Typ</span><span class="sxs-lookup"><span data-stu-id="83f86-217">Type</span></span>| <span data-ttu-id="83f86-218">Beschreibung</span><span class="sxs-lookup"><span data-stu-id="83f86-218">Description</span></span>|
|---|---|---|
|`itemId`| <span data-ttu-id="83f86-219">Zeichenfolge</span><span class="sxs-lookup"><span data-stu-id="83f86-219">String</span></span>|<span data-ttu-id="83f86-220">Der Exchange-Webdienste (EWS) für eine vorhandene Nachricht.</span><span class="sxs-lookup"><span data-stu-id="83f86-220">The Exchange Web Services (EWS) identifier for an existing message.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="83f86-221">Anforderungen</span><span class="sxs-lookup"><span data-stu-id="83f86-221">Requirements</span></span>

|<span data-ttu-id="83f86-222">Anforderung</span><span class="sxs-lookup"><span data-stu-id="83f86-222">Requirement</span></span>| <span data-ttu-id="83f86-223">Wert</span><span class="sxs-lookup"><span data-stu-id="83f86-223">Value</span></span>|
|---|---|
|[<span data-ttu-id="83f86-224">Mindestversion des Postfachanforderungssatzes</span><span class="sxs-lookup"><span data-stu-id="83f86-224">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="83f86-225">1.0</span><span class="sxs-lookup"><span data-stu-id="83f86-225">1.0</span></span>|
|[<span data-ttu-id="83f86-226">Mindestberechtigungsstufe</span><span class="sxs-lookup"><span data-stu-id="83f86-226">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="83f86-227">ReadItem</span><span class="sxs-lookup"><span data-stu-id="83f86-227">ReadItem</span></span>|
|[<span data-ttu-id="83f86-228">Zutreffender Outlook-Modus</span><span class="sxs-lookup"><span data-stu-id="83f86-228">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="83f86-229">Verfassen oder Lesen</span><span class="sxs-lookup"><span data-stu-id="83f86-229">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="83f86-230">Beispiel</span><span class="sxs-lookup"><span data-stu-id="83f86-230">Example</span></span>

```
Office.context.mailbox.displayMessageForm(messageId);
```

#### <a name="displaynewappointmentformparameters"></a><span data-ttu-id="83f86-231">displayNewAppointmentForm(parameters)</span><span class="sxs-lookup"><span data-stu-id="83f86-231">displayNewAppointmentForm(parameters)</span></span>

<span data-ttu-id="83f86-232">Zeigt ein Formular zum Erstellen eines neuen Kalendertermins an.</span><span class="sxs-lookup"><span data-stu-id="83f86-232">Displays a form for creating a new calendar appointment.</span></span>

> [!NOTE]
> <span data-ttu-id="83f86-233">Diese Methode wird in Outlook für iOS oder Outlook für Android nicht unterstützt.</span><span class="sxs-lookup"><span data-stu-id="83f86-233">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="83f86-p108">Mit der `displayNewAppointmentForm`-Methode wird ein Formular geöffnet, mit dem der Benutzer einen neuen Termin oder eine Besprechung erstellen kann. Wenn Parameter angegeben wurden, werden die Felder im Terminformular automatisch mit dem Inhalt der Parameter ausgefüllt.</span><span class="sxs-lookup"><span data-stu-id="83f86-p108">The `displayNewAppointmentForm` method opens a form that enables the user to create a new appointment or meeting. If parameters are specified, the appointment form fields are automatically populated with the contents of the parameters.</span></span>

<span data-ttu-id="83f86-p109">In Outlook Web App und OWA for Devices zeigt diese Methode immer ein Formular mit einem Teilnehmerfeld an. Wenn Sie keine Teilnehmer als Eingabeargumente angeben, zeigt die Methode ein Formular mit einer Schaltfläche **Speichern** an. Wenn Sie Teilnehmer angegeben haben, enthält das Formular die Teilnehmer und eine Schaltfläche **Senden**.</span><span class="sxs-lookup"><span data-stu-id="83f86-p109">In Outlook Web App and OWA for Devices, this method always displays a form with an attendees field. If you do not specify any attendees as input arguments, the method displays a form with a **Save** button. If you have specified attendees, the form would include the attendees and a **Send** button.</span></span>

<span data-ttu-id="83f86-p110">Wenn Teilnehmer oder Ressourcen im `requiredAttendees`-, `optionalAttendees`- oder `resources`-Parameter festgelegt werden, wird mit dieser Methode im Outlook Rich Client und Outlook RT ein Besprechungsformular mit der Schaltfläche **Senden** angezeigt. Wenn keine Teilnehmer angegeben werden, wird mit dieser Methode ein Terminformular mit der Schaltfläche **Speichern & schließen** angezeigt.</span><span class="sxs-lookup"><span data-stu-id="83f86-p110">In the Outlook rich client and Outlook RT, if you specify any attendees or resources in the `requiredAttendees`, `optionalAttendees`, or `resources` parameter, this method displays a meeting form with a **Send** button. If you don't specify any recipients, this method displays an appointment form with a **Save & Close** button.</span></span>

<span data-ttu-id="83f86-241">Wenn einer der Parameter die angegebenen Größenbeschränkungen überschreitet oder wenn ein unbekannter Parametername angegeben wird, wird eine Ausnahme ausgelöst.</span><span class="sxs-lookup"><span data-stu-id="83f86-241">If any of the parameters exceed the specified size limits, or if an unknown parameter name is specified, an exception is thrown.</span></span>

##### <a name="parameters"></a><span data-ttu-id="83f86-242">Parameter:</span><span class="sxs-lookup"><span data-stu-id="83f86-242">Parameters:</span></span>

|<span data-ttu-id="83f86-243">Name</span><span class="sxs-lookup"><span data-stu-id="83f86-243">Name</span></span>| <span data-ttu-id="83f86-244">Typ</span><span class="sxs-lookup"><span data-stu-id="83f86-244">Type</span></span>| <span data-ttu-id="83f86-245">Beschreibung</span><span class="sxs-lookup"><span data-stu-id="83f86-245">Description</span></span>|
|---|---|---|
| `parameters` | <span data-ttu-id="83f86-246">Object</span><span class="sxs-lookup"><span data-stu-id="83f86-246">Object</span></span> | <span data-ttu-id="83f86-247">Ein Wörterbuch mit Parametern, die den neuen Termin beschreiben</span><span class="sxs-lookup"><span data-stu-id="83f86-247">A dictionary of parameters describing the new appointment.</span></span> |
| `parameters.requiredAttendees` | <span data-ttu-id="83f86-248">Array.&lt;String&gt; &#124; Array.&lt;[EmailAddressDetails](/javascript/api/outlook_1_1/office.emailaddressdetails)&gt;</span><span class="sxs-lookup"><span data-stu-id="83f86-248">Array.&lt;String&gt; &#124; Array.&lt;[EmailAddressDetails](/javascript/api/outlook_1_1/office.emailaddressdetails)&gt;</span></span> | <span data-ttu-id="83f86-p111">Ein Array aus Zeichenfolgen mit den E-Mail-Adressen oder ein Array mit einem `EmailAddressDetails`-Objekt für jeden der erforderlichen Teilnehmer für den Termin. Das Array darf maximal 100 Einträge enthalten.</span><span class="sxs-lookup"><span data-stu-id="83f86-p111">An array of strings containing the email addresses or an array containing an `EmailAddressDetails` object for each of the required attendees for the appointment. The array is limited to a maximum of 100 entries.</span></span> |
| `parameters.optionalAttendees` | <span data-ttu-id="83f86-251">Array.&lt;String&gt; &#124; Array.&lt;[EmailAddressDetails](/javascript/api/outlook_1_1/office.emailaddressdetails)&gt;</span><span class="sxs-lookup"><span data-stu-id="83f86-251">Array.&lt;String&gt; &#124; Array.&lt;[EmailAddressDetails](/javascript/api/outlook_1_1/office.emailaddressdetails)&gt;</span></span> | <span data-ttu-id="83f86-p112">Ein Array aus Zeichenfolgen mit den E-Mail-Adressen oder ein Array mit einem Objekt des Typs `EmailAddressDetails` für jeden der optionalen Teilnehmer des Termins. Das Array darf maximal 100 Einträge enthalten.</span><span class="sxs-lookup"><span data-stu-id="83f86-p112">An array of strings containing the email addresses or an array containing an `EmailAddressDetails` object for each of the optional attendees for the appointment. The array is limited to a maximum of 100 entries.</span></span> |
| `parameters.start` | <span data-ttu-id="83f86-254">Datum</span><span class="sxs-lookup"><span data-stu-id="83f86-254">Date</span></span> | <span data-ttu-id="83f86-255">Ein Objekt des Typs `Date`. Gibt das Startdatum und den Beginn des Termins an.</span><span class="sxs-lookup"><span data-stu-id="83f86-255">A `Date` object specifying the start date and time of the appointment.</span></span> |
| `parameters.end` | <span data-ttu-id="83f86-256">Date</span><span class="sxs-lookup"><span data-stu-id="83f86-256">Date</span></span> | <span data-ttu-id="83f86-257">Ein Objekt des Typs `Date`. Gibt das Enddatum und das Ende des Termins an.</span><span class="sxs-lookup"><span data-stu-id="83f86-257">A `Date` object specifying the end date and time of the appointment.</span></span> |
| `parameters.location` | <span data-ttu-id="83f86-258">String</span><span class="sxs-lookup"><span data-stu-id="83f86-258">String</span></span> | <span data-ttu-id="83f86-p113">Eine Zeichenfolge mit dem Ort für den Termin. Die Zeichenfolge ist auf maximal 255 Zeichen beschränkt.</span><span class="sxs-lookup"><span data-stu-id="83f86-p113">A string containing the location of the appointment. The string is limited to a maximum of 255 characters.</span></span> |
| `parameters.resources` | <span data-ttu-id="83f86-261">Array.&lt;String&gt;</span><span class="sxs-lookup"><span data-stu-id="83f86-261">Array.&lt;String&gt;</span></span> | <span data-ttu-id="83f86-p114">Ein Array aus Zeichenfolgen, die die für den Termin erforderlichen Ressourcen enthalten. Das Array darf maximal 100 Einträge enthalten.</span><span class="sxs-lookup"><span data-stu-id="83f86-p114">An array of strings containing the resources required for the appointment. The array is limited to a maximum of 100 entries.</span></span> |
| `parameters.subject` | <span data-ttu-id="83f86-264">Zeichenfolge</span><span class="sxs-lookup"><span data-stu-id="83f86-264">String</span></span> | <span data-ttu-id="83f86-p115">Eine Zeichenfolge mit dem Betreff für den Termin. Die Zeichenfolge ist auf maximal 255 Zeichen beschränkt.</span><span class="sxs-lookup"><span data-stu-id="83f86-p115">A string containing the subject of the appointment. The string is limited to a maximum of 255 characters.</span></span> |
| `parameters.body` | <span data-ttu-id="83f86-267">Zeichenfolge</span><span class="sxs-lookup"><span data-stu-id="83f86-267">String</span></span> | <span data-ttu-id="83f86-p116">Der Text des Termins. Der Textinhalt darf maximal 32 KB umfassen.</span><span class="sxs-lookup"><span data-stu-id="83f86-p116">The body of the appointment. The body content is limited to a maximum size of 32 KB.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="83f86-270">Anforderungen</span><span class="sxs-lookup"><span data-stu-id="83f86-270">Requirements</span></span>

|<span data-ttu-id="83f86-271">Anforderung</span><span class="sxs-lookup"><span data-stu-id="83f86-271">Requirement</span></span>| <span data-ttu-id="83f86-272">Wert</span><span class="sxs-lookup"><span data-stu-id="83f86-272">Value</span></span>|
|---|---|
|[<span data-ttu-id="83f86-273">Mindestversion des Postfachanforderungssatzes</span><span class="sxs-lookup"><span data-stu-id="83f86-273">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="83f86-274">1.0</span><span class="sxs-lookup"><span data-stu-id="83f86-274">1.0</span></span>|
|[<span data-ttu-id="83f86-275">Mindestberechtigungsstufe</span><span class="sxs-lookup"><span data-stu-id="83f86-275">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="83f86-276">ReadItem</span><span class="sxs-lookup"><span data-stu-id="83f86-276">ReadItem</span></span>|
|[<span data-ttu-id="83f86-277">Zutreffender Outlook-Modus</span><span class="sxs-lookup"><span data-stu-id="83f86-277">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="83f86-278">Lesen</span><span class="sxs-lookup"><span data-stu-id="83f86-278">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="83f86-279">Beispiel</span><span class="sxs-lookup"><span data-stu-id="83f86-279">Example</span></span>

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

#### <a name="getcallbacktokenasynccallback-usercontext"></a><span data-ttu-id="83f86-280">getCallbackTokenAsync(callback, [userContext])</span><span class="sxs-lookup"><span data-stu-id="83f86-280">getCallbackTokenAsync(callback, [userContext])</span></span>

<span data-ttu-id="83f86-281">Ruft eine Zeichenfolge ab, die einen Token enthält, der verwendet wird, um eine Anlage oder ein Element von einem Exchange Server abzurufen.</span><span class="sxs-lookup"><span data-stu-id="83f86-281">Gets a string that contains a token used to get an attachment or item from an Exchange Server.</span></span>

<span data-ttu-id="83f86-p117">Die `getCallbackTokenAsync`-Methode führt einen asynchronen Aufruf zum Abruf eines nicht transparenten Tokens vom Exchange-Server aus, der das Postfach des Benutzers hostet. Die Gültigkeitsdauer des Rückruftokens beträgt 5 Minuten.</span><span class="sxs-lookup"><span data-stu-id="83f86-p117">The `getCallbackTokenAsync` method makes an asynchronous call to get an opaque token from the Exchange Server that hosts the user's mailbox. The lifetime of the callback token is 5 minutes.</span></span>

<span data-ttu-id="83f86-p118">Sie können das Token und einen Anlagen- oder einen Elementbezeichner an ein Drittanbietersystem weitergeben. Das Drittanbietersystem verwendet das Token als Trägerautorisierungstoken, um den EWS-Vorgang (Exchange-Webdienste) [GetAttachment](https://docs.microsoft.com/exchange/client-developer/web-service-reference/getattachment-operation) oder [GetItem](https://docs.microsoft.com/exchange/client-developer/web-service-reference/getitem-operation) aufzurufen und eine Anlage oder ein Element zurückzugeben. Sie können z. B. einen Remotedienst erstellen, um [Anlagen aus dem ausgewählten Element abzurufen](https://docs.microsoft.com/outlook/add-ins/get-attachments-of-an-outlook-item).</span><span class="sxs-lookup"><span data-stu-id="83f86-p118">You can pass the token and an attachment identifier or item identifier to a third-party system. The third-party system uses the token as a bearer authorization token to call the Exchange Web Services (EWS) [GetAttachment](https://docs.microsoft.com/exchange/client-developer/web-service-reference/getattachment-operation) or [GetItem](https://docs.microsoft.com/exchange/client-developer/web-service-reference/getitem-operation) operation to return an attachment or item. For example, you can create a remote service to [get attachments from the selected item](https://docs.microsoft.com/outlook/add-ins/get-attachments-of-an-outlook-item).</span></span>

<span data-ttu-id="83f86-287">Ihre App muss über die im Manifest angegebene **ReadItem**-Berechtigung verfügen, um die `getCallbackTokenAsync`-Methode aufzurufen.</span><span class="sxs-lookup"><span data-stu-id="83f86-287">Your app must have the **ReadItem** permission specified in its manifest to call the `getCallbackTokenAsync` method.</span></span>

##### <a name="parameters"></a><span data-ttu-id="83f86-288">Parameter:</span><span class="sxs-lookup"><span data-stu-id="83f86-288">Parameters:</span></span>

|<span data-ttu-id="83f86-289">Name</span><span class="sxs-lookup"><span data-stu-id="83f86-289">Name</span></span>| <span data-ttu-id="83f86-290">Typ</span><span class="sxs-lookup"><span data-stu-id="83f86-290">Type</span></span>| <span data-ttu-id="83f86-291">Attribute</span><span class="sxs-lookup"><span data-stu-id="83f86-291">Attributes</span></span>| <span data-ttu-id="83f86-292">Beschreibung</span><span class="sxs-lookup"><span data-stu-id="83f86-292">Description</span></span>|
|---|---|---|---|
|`callback`| <span data-ttu-id="83f86-293">Funktion</span><span class="sxs-lookup"><span data-stu-id="83f86-293">function</span></span>||<span data-ttu-id="83f86-294">Wenn die Methode abgeschlossen wird, wird die Funktion , die im `callback`-Parameter übergeben wird, mit einem einzelnen Parameter (`asyncResult`) aufgerufen, der ein [`AsyncResult`](/javascript/api/office/office.asyncresult)-Objekt darstellt.</span><span class="sxs-lookup"><span data-stu-id="83f86-294">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span><br/><br/><span data-ttu-id="83f86-295">Das Token wird als Zeichenfolge in der `asyncResult.value`-Eigenschaft bereitgestellt.</span><span class="sxs-lookup"><span data-stu-id="83f86-295">The token is provided as a string in the `asyncResult.value` property.</span></span>|
|`userContext`| <span data-ttu-id="83f86-296">Objekt</span><span class="sxs-lookup"><span data-stu-id="83f86-296">Object</span></span>| <span data-ttu-id="83f86-297">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="83f86-297">&lt;optional&gt;</span></span>|<span data-ttu-id="83f86-298">Jegliche Zustandsdaten, die an die asynchrone Methode übergeben werden.</span><span class="sxs-lookup"><span data-stu-id="83f86-298">Any state data that is passed to the asynchronous method.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="83f86-299">Anforderungen</span><span class="sxs-lookup"><span data-stu-id="83f86-299">Requirements</span></span>

|<span data-ttu-id="83f86-300">Anforderung</span><span class="sxs-lookup"><span data-stu-id="83f86-300">Requirement</span></span>| <span data-ttu-id="83f86-301">Wert</span><span class="sxs-lookup"><span data-stu-id="83f86-301">Value</span></span>|
|---|---|
|[<span data-ttu-id="83f86-302">Mindestversion des Postfachanforderungssatzes</span><span class="sxs-lookup"><span data-stu-id="83f86-302">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="83f86-303">1.0</span><span class="sxs-lookup"><span data-stu-id="83f86-303">1.0</span></span>|
|[<span data-ttu-id="83f86-304">Mindestberechtigungsstufe</span><span class="sxs-lookup"><span data-stu-id="83f86-304">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="83f86-305">ReadItem</span><span class="sxs-lookup"><span data-stu-id="83f86-305">ReadItem</span></span>|
|[<span data-ttu-id="83f86-306">Zutreffender Outlook-Modus</span><span class="sxs-lookup"><span data-stu-id="83f86-306">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="83f86-307">Lesen</span><span class="sxs-lookup"><span data-stu-id="83f86-307">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="83f86-308">Beispiel</span><span class="sxs-lookup"><span data-stu-id="83f86-308">Example</span></span>

```js
function getCallbackToken() {
  Office.context.mailbox.getCallbackTokenAsync(cb);
}

function cb(asyncResult) {
  var token = asyncResult.value;
}
```

####  <a name="getuseridentitytokenasynccallback-usercontext"></a><span data-ttu-id="83f86-309">getUserIdentityTokenAsync(callback, [userContext])</span><span class="sxs-lookup"><span data-stu-id="83f86-309">getUserIdentityTokenAsync(callback, [userContext])</span></span>

<span data-ttu-id="83f86-310">Ruft ein Token ab, das den Benutzer und das Office-Add-In identifiziert.</span><span class="sxs-lookup"><span data-stu-id="83f86-310">Gets a token identifying the user and the Office Add-in.</span></span>

<span data-ttu-id="83f86-311">Die `getUserIdentityTokenAsync`-Methode gibt ein Token zurück, das Sie dazu verwenden können, um das [Add-In und den Benutzer mit einem Drittanbietersystem zu identifizieren und zu authentifizieren](https://docs.microsoft.com/outlook/add-ins/authentication).</span><span class="sxs-lookup"><span data-stu-id="83f86-311">The `getUserIdentityTokenAsync` method returns a token that you can use to identify and [authenticate the add-in and user with a third-party system](https://docs.microsoft.com/outlook/add-ins/authentication).</span></span>

##### <a name="parameters"></a><span data-ttu-id="83f86-312">Parameter:</span><span class="sxs-lookup"><span data-stu-id="83f86-312">Parameters:</span></span>

|<span data-ttu-id="83f86-313">Name</span><span class="sxs-lookup"><span data-stu-id="83f86-313">Name</span></span>| <span data-ttu-id="83f86-314">Typ</span><span class="sxs-lookup"><span data-stu-id="83f86-314">Type</span></span>| <span data-ttu-id="83f86-315">Attribute</span><span class="sxs-lookup"><span data-stu-id="83f86-315">Attributes</span></span>| <span data-ttu-id="83f86-316">Beschreibung</span><span class="sxs-lookup"><span data-stu-id="83f86-316">Description</span></span>|
|---|---|---|---|
|`callback`| <span data-ttu-id="83f86-317">Funktion</span><span class="sxs-lookup"><span data-stu-id="83f86-317">function</span></span>||<span data-ttu-id="83f86-318">Wenn die Methode abgeschlossen wird, wird die Funktion , die im `callback`-Parameter übergeben wird, mit einem einzelnen Parameter (`asyncResult`) aufgerufen, der ein [`AsyncResult`](/javascript/api/office/office.asyncresult)-Objekt darstellt.</span><span class="sxs-lookup"><span data-stu-id="83f86-318">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span><br/><br/><span data-ttu-id="83f86-319">Das Token wird als Zeichenfolge in der `asyncResult.value`-Eigenschaft bereitgestellt.</span><span class="sxs-lookup"><span data-stu-id="83f86-319">The token is provided as a string in the `asyncResult.value` property.</span></span>|
|`userContext`| <span data-ttu-id="83f86-320">Objekt</span><span class="sxs-lookup"><span data-stu-id="83f86-320">Object</span></span>| <span data-ttu-id="83f86-321">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="83f86-321">&lt;optional&gt;</span></span>|<span data-ttu-id="83f86-322">Jegliche Zustandsdaten, die an die asynchrone Methode übergeben werden.</span><span class="sxs-lookup"><span data-stu-id="83f86-322">Any state data that is passed to the asynchronous method.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="83f86-323">Anforderungen</span><span class="sxs-lookup"><span data-stu-id="83f86-323">Requirements</span></span>

|<span data-ttu-id="83f86-324">Anforderung</span><span class="sxs-lookup"><span data-stu-id="83f86-324">Requirement</span></span>| <span data-ttu-id="83f86-325">Wert</span><span class="sxs-lookup"><span data-stu-id="83f86-325">Value</span></span>|
|---|---|
|[<span data-ttu-id="83f86-326">Mindestversion des Postfachanforderungssatzes</span><span class="sxs-lookup"><span data-stu-id="83f86-326">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="83f86-327">1.0</span><span class="sxs-lookup"><span data-stu-id="83f86-327">1.0</span></span>|
|[<span data-ttu-id="83f86-328">Mindestberechtigungsstufe</span><span class="sxs-lookup"><span data-stu-id="83f86-328">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="83f86-329">ReadItem</span><span class="sxs-lookup"><span data-stu-id="83f86-329">ReadItem</span></span>|
|[<span data-ttu-id="83f86-330">Zutreffender Outlook-Modus</span><span class="sxs-lookup"><span data-stu-id="83f86-330">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="83f86-331">Verfassen oder Lesen</span><span class="sxs-lookup"><span data-stu-id="83f86-331">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="83f86-332">Beispiel</span><span class="sxs-lookup"><span data-stu-id="83f86-332">Example</span></span>

```js
function getIdentityToken() {
  Office.context.mailbox.getUserIdentityTokenAsync(cb);
}

function cb(asyncResult) {
  var token = asyncResult.value;
}
```

####  <a name="makeewsrequestasyncdata-callback-usercontext"></a><span data-ttu-id="83f86-333">makeEwsRequestAsync(data, callback, [userContext])</span><span class="sxs-lookup"><span data-stu-id="83f86-333">makeEwsRequestAsync(data, callback, [userContext])</span></span>

<span data-ttu-id="83f86-334">Richtet eine asynchrone Anforderung an einen EWS-Dienst (Exchange Web Services) auf dem Exchange-Server, der das Postfach des Benutzers hostet.</span><span class="sxs-lookup"><span data-stu-id="83f86-334">Makes an asynchronous request to an Exchange Web Services (EWS) service on the Exchange server that hosts the user’s mailbox.</span></span>

> [!NOTE]
> <span data-ttu-id="83f86-335">Diese Methode wird in den folgenden Szenarien nicht unterstützt.</span><span class="sxs-lookup"><span data-stu-id="83f86-335">This method is not supported in the following scenarios.</span></span>
> - <span data-ttu-id="83f86-336">In Outlook für iOS oder Outlook für Android</span><span class="sxs-lookup"><span data-stu-id="83f86-336">In Outlook for iOS or Outlook for Android</span></span>
> - <span data-ttu-id="83f86-337">Wenn das Add-In geladen wird, in einem Postfach Google Mail</span><span class="sxs-lookup"><span data-stu-id="83f86-337">When the add-in is loaded in a Gmail mailbox</span></span>
> 
> <span data-ttu-id="83f86-338">In diesen Fällen sollte-add-ins [mithilfe von REST-APIs](https://docs.microsoft.com/outlook/add-ins/use-rest-api) , stattdessen Zugriff auf das Postfach des Benutzers.</span><span class="sxs-lookup"><span data-stu-id="83f86-338">In these cases, add-ins should [use REST APIs](https://docs.microsoft.com/outlook/add-ins/use-rest-api) to access the user's mailbox instead.</span></span>

<span data-ttu-id="83f86-339">Die `makeEwsRequestAsync`-Methode sendet eine EWS-Anforderung für das Add-In zu Exchange.</span><span class="sxs-lookup"><span data-stu-id="83f86-339">The `makeEwsRequestAsync` method sends an EWS request on behalf of the add-in to Exchange.</span></span> <span data-ttu-id="83f86-340">Eine Liste der unterstützten EWS-Vorgänge finden Sie unter [Aufrufen von Webdiensten aus einem Outlook-add-in](https://docs.microsoft.com/outlook/add-ins/web-services#ews-operations-that-add-ins-support) .</span><span class="sxs-lookup"><span data-stu-id="83f86-340">See [Call web services from an Outlook add-in](https://docs.microsoft.com/outlook/add-ins/web-services#ews-operations-that-add-ins-support) for a list of the supported EWS operations.</span></span>

<span data-ttu-id="83f86-341">Sie können keine Elemente, die Ordnern zugeordnet sind, mit der `makeEwsRequestAsync`-Methode anfordern.</span><span class="sxs-lookup"><span data-stu-id="83f86-341">You cannot request Folder Associated Items with the `makeEwsRequestAsync` method.</span></span>

<span data-ttu-id="83f86-342">Die XML-Anfrage muss UTF-8-Codierung angeben.</span><span class="sxs-lookup"><span data-stu-id="83f86-342">The XML request must specify UTF-8 encoding.</span></span>

```
<?xml version="1.0" encoding="utf-8"?>
```

<span data-ttu-id="83f86-p120">Das Add-In muss über die **ReadWriteMailbox**-Berechtigung verfügen, um die `makeEwsRequestAsync`-Methode zu verwenden. Informationen zur **ReadWriteMailbox**-Berechtigung und zu den EWS-Vorgängen, die Sie mit der `makeEwsRequestAsync`-Methode aufrufen können, finden Sie unter [Angeben von Berechtigungen für den E-Mail-Add-In-Zugriff auf das Benutzerpostfach](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions).</span><span class="sxs-lookup"><span data-stu-id="83f86-p120">Your add-in must have the **ReadWriteMailbox** permission to use the `makeEwsRequestAsync` method. For information about using the **ReadWriteMailbox** permission and the EWS operations that you can call with the `makeEwsRequestAsync` method, see [Specify permissions for mail add-in access to the user's mailbox](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions).</span></span>

> [!NOTE]
> <span data-ttu-id="83f86-345">Der Serveradministrator muss festgelegt `OAuthAuthentication` auf "true" auf dem Client Access Server EWS-Verzeichnis so aktivieren Sie die `makeEwsRequestAsync` -Methode, um EWS stellen anfordert.</span><span class="sxs-lookup"><span data-stu-id="83f86-345">The server administrator must set `OAuthAuthentication` to true on the Client Access Server EWS directory to enable the `makeEwsRequestAsync` method to make EWS requests.</span></span>

##### <a name="version-differences"></a><span data-ttu-id="83f86-346">Versionsunterschiede</span><span class="sxs-lookup"><span data-stu-id="83f86-346">Version differences</span></span>

<span data-ttu-id="83f86-347">Wenn Sie die `makeEwsRequestAsync`-Methode in Mail-Apps verwenden, die in älteren Outlook-Versionen als Version 15.0.4535.1004 ausgeführt werden, sollten Sie den Codierungswert auf `ISO-8859-1` festlegen.</span><span class="sxs-lookup"><span data-stu-id="83f86-347">When you use the `makeEwsRequestAsync` method in mail apps running in Outlook versions earlier than version 15.0.4535.1004, you should set the encoding value to `ISO-8859-1`.</span></span>

```
<?xml version="1.0" encoding="iso-8859-1"?>
```

<span data-ttu-id="83f86-p121">Sie müssen den Codierungswert nicht festlegen, wenn Ihre Mail-App in Outlook im Web ausgeführt wird. Sie können mithilfe der mailbox.diagnostics.hostName-Eigenschaft ermitteln, ob Ihre Mail-App in Outlook oder Outlook im Web ausgeführt wird. Sie können mithilfe der mailbox.diagnostics.hostVersion-Eigenschaft ermitteln, welche Version von Outlook ausgeführt wird.</span><span class="sxs-lookup"><span data-stu-id="83f86-p121">You do not need to set the encoding value when your mail app is running in Outlook on the web. You can determine whether your mail app is running in Outlook or Outlook on the web by using the mailbox.diagnostics.hostName property. You can determine what version of Outlook is running by using the mailbox.diagnostics.hostVersion property.</span></span>

##### <a name="parameters"></a><span data-ttu-id="83f86-351">Parameter:</span><span class="sxs-lookup"><span data-stu-id="83f86-351">Parameters:</span></span>

|<span data-ttu-id="83f86-352">Name</span><span class="sxs-lookup"><span data-stu-id="83f86-352">Name</span></span>| <span data-ttu-id="83f86-353">Typ</span><span class="sxs-lookup"><span data-stu-id="83f86-353">Type</span></span>| <span data-ttu-id="83f86-354">Attribute</span><span class="sxs-lookup"><span data-stu-id="83f86-354">Attributes</span></span>| <span data-ttu-id="83f86-355">Beschreibung</span><span class="sxs-lookup"><span data-stu-id="83f86-355">Description</span></span>|
|---|---|---|---|
|`data`| <span data-ttu-id="83f86-356">Zeichenfolge</span><span class="sxs-lookup"><span data-stu-id="83f86-356">String</span></span>||<span data-ttu-id="83f86-357">Die EWS-Anforderung.</span><span class="sxs-lookup"><span data-stu-id="83f86-357">The EWS request.</span></span>|
|`callback`| <span data-ttu-id="83f86-358">Funktion</span><span class="sxs-lookup"><span data-stu-id="83f86-358">function</span></span>||<span data-ttu-id="83f86-359">Wenn die Methode abgeschlossen wird, wird die Funktion , die im `callback`-Parameter übergeben wird, mit einem einzelnen Parameter (`asyncResult`) aufgerufen, der ein [`AsyncResult`](/javascript/api/office/office.asyncresult)-Objekt darstellt.</span><span class="sxs-lookup"><span data-stu-id="83f86-359">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span><br/><br/><span data-ttu-id="83f86-360">Das XML-Ergebnis des EWS-Aufrufs wird als Zeichenfolge in der `asyncResult.value`-Eigenschaft bereitgestellt.</span><span class="sxs-lookup"><span data-stu-id="83f86-360">The XML result of the EWS call is provided as a string in the `asyncResult.value` property.</span></span> <span data-ttu-id="83f86-361">Wenn das Ergebnis 1 MB Größe überschreitet, wird stattdessen eine Fehlermeldung zurückgegeben.</span><span class="sxs-lookup"><span data-stu-id="83f86-361">If the result exceeds 1 MB in size, an error message is returned instead.</span></span>|
|`userContext`| <span data-ttu-id="83f86-362">Objekt</span><span class="sxs-lookup"><span data-stu-id="83f86-362">Object</span></span>| <span data-ttu-id="83f86-363">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="83f86-363">&lt;optional&gt;</span></span>|<span data-ttu-id="83f86-364">Jegliche Zustandsdaten, die an die asynchrone Methode übergeben werden.</span><span class="sxs-lookup"><span data-stu-id="83f86-364">Any state data that is passed to the asynchronous method.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="83f86-365">Anforderungen</span><span class="sxs-lookup"><span data-stu-id="83f86-365">Requirements</span></span>

|<span data-ttu-id="83f86-366">Anforderung</span><span class="sxs-lookup"><span data-stu-id="83f86-366">Requirement</span></span>| <span data-ttu-id="83f86-367">Wert</span><span class="sxs-lookup"><span data-stu-id="83f86-367">Value</span></span>|
|---|---|
|[<span data-ttu-id="83f86-368">Mindestversion des Postfachanforderungssatzes</span><span class="sxs-lookup"><span data-stu-id="83f86-368">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="83f86-369">1.0</span><span class="sxs-lookup"><span data-stu-id="83f86-369">1.0</span></span>|
|[<span data-ttu-id="83f86-370">Mindestberechtigungsstufe</span><span class="sxs-lookup"><span data-stu-id="83f86-370">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="83f86-371">ReadWriteMailbox</span><span class="sxs-lookup"><span data-stu-id="83f86-371">ReadWriteMailbox</span></span>|
|[<span data-ttu-id="83f86-372">Zutreffender Outlook-Modus</span><span class="sxs-lookup"><span data-stu-id="83f86-372">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="83f86-373">Verfassen oder Lesen</span><span class="sxs-lookup"><span data-stu-id="83f86-373">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="83f86-374">Beispiel</span><span class="sxs-lookup"><span data-stu-id="83f86-374">Example</span></span>

<span data-ttu-id="83f86-375">Das folgende Beispiel ruft `makeEwsRequestAsync` zum Verwenden des `GetItem`-Vorgangs auf, um den Betreff eines Elements abzurufen.</span><span class="sxs-lookup"><span data-stu-id="83f86-375">The following example calls `makeEwsRequestAsync` to use the `GetItem` operation to get the subject of an item.</span></span>

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