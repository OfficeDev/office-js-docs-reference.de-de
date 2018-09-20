 

# <a name="office"></a><span data-ttu-id="69065-101">Büro</span><span class="sxs-lookup"><span data-stu-id="69065-101">Office</span></span>

<span data-ttu-id="69065-p101">Der Office-Namespace stellt gemeinsam genutzte Oberflächen bereit, die von Add-Ins in allen Office-Apps verwendet werden. Diese Auflistung dokumentiert nur die Schnittstellen, die von Outlook-Add-Ins verwendet werden. Eine vollständige Auflistung des Office-Namespaces finden Sie im Thema zur [freigegebenen API](/javascript/api/office).</span><span class="sxs-lookup"><span data-stu-id="69065-p101">The Office namespace provides shared interfaces that are used by add-ins in all of the Office apps. This listing documents only those interfaces that are used by Outlook add-ins. For a full listing of the Office namespace, see the [Shared API](/javascript/api/office).</span></span>

##### <a name="requirements"></a><span data-ttu-id="69065-104">Anforderungen</span><span class="sxs-lookup"><span data-stu-id="69065-104">Requirements</span></span>

|<span data-ttu-id="69065-105">Anforderung</span><span class="sxs-lookup"><span data-stu-id="69065-105">Requirement</span></span>| <span data-ttu-id="69065-106">Wert</span><span class="sxs-lookup"><span data-stu-id="69065-106">Value</span></span>|
|---|---|
|[<span data-ttu-id="69065-107">Mindestversion des Postfachanforderungssatzes</span><span class="sxs-lookup"><span data-stu-id="69065-107">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="69065-108">1.0</span><span class="sxs-lookup"><span data-stu-id="69065-108">1.0</span></span>|
|[<span data-ttu-id="69065-109">Zutreffender Outlook-Modus</span><span class="sxs-lookup"><span data-stu-id="69065-109">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="69065-110">Verfassen oder Lesen</span><span class="sxs-lookup"><span data-stu-id="69065-110">Compose or read</span></span>|

##### <a name="members-and-methods"></a><span data-ttu-id="69065-111">Elemente und Methoden</span><span class="sxs-lookup"><span data-stu-id="69065-111">Members and methods</span></span>

| <span data-ttu-id="69065-112">Element</span><span class="sxs-lookup"><span data-stu-id="69065-112">Member</span></span> | <span data-ttu-id="69065-113">Typ</span><span class="sxs-lookup"><span data-stu-id="69065-113">Type</span></span> |
|--------|------|
| [<span data-ttu-id="69065-114">AsyncResultStatus</span><span class="sxs-lookup"><span data-stu-id="69065-114">AsyncResultStatus</span></span>](#asyncresultstatus-string) | <span data-ttu-id="69065-115">Element</span><span class="sxs-lookup"><span data-stu-id="69065-115">Member</span></span> |
| [<span data-ttu-id="69065-116">CoercionType</span><span class="sxs-lookup"><span data-stu-id="69065-116">CoercionType</span></span>](#coerciontype-string) | <span data-ttu-id="69065-117">Element</span><span class="sxs-lookup"><span data-stu-id="69065-117">Member</span></span> |
| [<span data-ttu-id="69065-118">EventType</span><span class="sxs-lookup"><span data-stu-id="69065-118">EventType</span></span>](#eventtype-string) | <span data-ttu-id="69065-119">Element</span><span class="sxs-lookup"><span data-stu-id="69065-119">Member</span></span> |
| [<span data-ttu-id="69065-120">SourceProperty</span><span class="sxs-lookup"><span data-stu-id="69065-120">SourceProperty</span></span>](#sourceproperty-string) | <span data-ttu-id="69065-121">Element</span><span class="sxs-lookup"><span data-stu-id="69065-121">Member</span></span> |

### <a name="namespaces"></a><span data-ttu-id="69065-122">Namespaces</span><span class="sxs-lookup"><span data-stu-id="69065-122">Namespaces</span></span>

<span data-ttu-id="69065-123">[Kontextmenü](office.context.md): gemeinsam genutzten Oberflächen aus dem Office-Add-ins-API-Kontext-Namespace für die Verwendung in der Outlook-add-in-API bietet.</span><span class="sxs-lookup"><span data-stu-id="69065-123">[context](office.context.md): Provides shared interfaces from the Office Add-ins API's context namespace for use in the Outlook add-in API.</span></span>

<span data-ttu-id="69065-124">[MailboxEnums](/javascript/api/outlook/office.mailboxenums.attachmenttype): Enthält die Enumerationen ItemType, EntityType, AttachmentType, RecipientType, ResponseType und ItemNotificationMessageType.</span><span class="sxs-lookup"><span data-stu-id="69065-124">[MailboxEnums](/javascript/api/outlook/office.mailboxenums.attachmenttype): Includes the ItemType, EntityType, AttachmentType, RecipientType, ResponseType, and ItemNotificationMessageType enumerations.</span></span>

### <a name="members"></a><span data-ttu-id="69065-125">Elemente</span><span class="sxs-lookup"><span data-stu-id="69065-125">Members</span></span>

####  <a name="asyncresultstatus-string"></a><span data-ttu-id="69065-126">AsyncResultStatus :String</span><span class="sxs-lookup"><span data-stu-id="69065-126">AsyncResultStatus :String</span></span>

<span data-ttu-id="69065-127">Gibt das Ergebnis eines asynchronen Aufrufs an.</span><span class="sxs-lookup"><span data-stu-id="69065-127">Specifies the result of an asynchronous call.</span></span>

##### <a name="type"></a><span data-ttu-id="69065-128">Typ:</span><span class="sxs-lookup"><span data-stu-id="69065-128">Type:</span></span>

*   <span data-ttu-id="69065-129">Zeichenfolge</span><span class="sxs-lookup"><span data-stu-id="69065-129">String</span></span>

##### <a name="properties"></a><span data-ttu-id="69065-130">Eigenschaften:</span><span class="sxs-lookup"><span data-stu-id="69065-130">Properties:</span></span>

|<span data-ttu-id="69065-131">Name</span><span class="sxs-lookup"><span data-stu-id="69065-131">Name</span></span>| <span data-ttu-id="69065-132">Typ</span><span class="sxs-lookup"><span data-stu-id="69065-132">Type</span></span>| <span data-ttu-id="69065-133">Beschreibung</span><span class="sxs-lookup"><span data-stu-id="69065-133">Description</span></span>|
|---|---|---|
|`Succeeded`| <span data-ttu-id="69065-134">Zeichenfolge</span><span class="sxs-lookup"><span data-stu-id="69065-134">String</span></span>|<span data-ttu-id="69065-135">Der Aufruf war erfolgreich.</span><span class="sxs-lookup"><span data-stu-id="69065-135">The call succeeded.</span></span>|
|`Failed`| <span data-ttu-id="69065-136">Zeichenfolge</span><span class="sxs-lookup"><span data-stu-id="69065-136">String</span></span>|<span data-ttu-id="69065-137">Der Aufruf ist fehlerhaft.</span><span class="sxs-lookup"><span data-stu-id="69065-137">The call failed.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="69065-138">Anforderungen</span><span class="sxs-lookup"><span data-stu-id="69065-138">Requirements</span></span>

|<span data-ttu-id="69065-139">Anforderung</span><span class="sxs-lookup"><span data-stu-id="69065-139">Requirement</span></span>| <span data-ttu-id="69065-140">Wert</span><span class="sxs-lookup"><span data-stu-id="69065-140">Value</span></span>|
|---|---|
|[<span data-ttu-id="69065-141">Mindestversion des Postfachanforderungssatzes</span><span class="sxs-lookup"><span data-stu-id="69065-141">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="69065-142">1.0</span><span class="sxs-lookup"><span data-stu-id="69065-142">1.0</span></span>|
|[<span data-ttu-id="69065-143">Zutreffender Outlook-Modus</span><span class="sxs-lookup"><span data-stu-id="69065-143">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="69065-144">Verfassen oder Lesen</span><span class="sxs-lookup"><span data-stu-id="69065-144">Compose or read</span></span>|

---

####  <a name="coerciontype-string"></a><span data-ttu-id="69065-145">CoercionType :String</span><span class="sxs-lookup"><span data-stu-id="69065-145">CoercionType :String</span></span>

<span data-ttu-id="69065-146">Gibt an, wie Daten umgewandelt werden sollen, die von der aufgerufenen Methode zurückgegeben oder festgelegt wurden.</span><span class="sxs-lookup"><span data-stu-id="69065-146">Specifies how to coerce data returned or set by the invoked method.</span></span>

##### <a name="type"></a><span data-ttu-id="69065-147">Typ:</span><span class="sxs-lookup"><span data-stu-id="69065-147">Type:</span></span>

*   <span data-ttu-id="69065-148">Zeichenfolge</span><span class="sxs-lookup"><span data-stu-id="69065-148">String</span></span>

##### <a name="properties"></a><span data-ttu-id="69065-149">Eigenschaften:</span><span class="sxs-lookup"><span data-stu-id="69065-149">Properties:</span></span>

|<span data-ttu-id="69065-150">Name</span><span class="sxs-lookup"><span data-stu-id="69065-150">Name</span></span>| <span data-ttu-id="69065-151">Typ</span><span class="sxs-lookup"><span data-stu-id="69065-151">Type</span></span>| <span data-ttu-id="69065-152">Beschreibung</span><span class="sxs-lookup"><span data-stu-id="69065-152">Description</span></span>|
|---|---|---|
|`Html`| <span data-ttu-id="69065-153">Zeichenfolge</span><span class="sxs-lookup"><span data-stu-id="69065-153">String</span></span>|<span data-ttu-id="69065-154">Fordert die Rückgabe der Daten im HTML-Format an.</span><span class="sxs-lookup"><span data-stu-id="69065-154">Requests the data be returned in HTML format.</span></span>|
|`Text`| <span data-ttu-id="69065-155">Zeichenfolge</span><span class="sxs-lookup"><span data-stu-id="69065-155">String</span></span>|<span data-ttu-id="69065-156">Fordert die Rückgabe der Daten im Textformat an.</span><span class="sxs-lookup"><span data-stu-id="69065-156">Requests the data be returned in text format.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="69065-157">Anforderungen</span><span class="sxs-lookup"><span data-stu-id="69065-157">Requirements</span></span>

|<span data-ttu-id="69065-158">Anforderung</span><span class="sxs-lookup"><span data-stu-id="69065-158">Requirement</span></span>| <span data-ttu-id="69065-159">Wert</span><span class="sxs-lookup"><span data-stu-id="69065-159">Value</span></span>|
|---|---|
|[<span data-ttu-id="69065-160">Mindestversion des Postfachanforderungssatzes</span><span class="sxs-lookup"><span data-stu-id="69065-160">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="69065-161">1.0</span><span class="sxs-lookup"><span data-stu-id="69065-161">1.0</span></span>|
|[<span data-ttu-id="69065-162">Zutreffender Outlook-Modus</span><span class="sxs-lookup"><span data-stu-id="69065-162">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="69065-163">Verfassen oder Lesen</span><span class="sxs-lookup"><span data-stu-id="69065-163">Compose or read</span></span>|

---

####  <a name="eventtype-string"></a><span data-ttu-id="69065-164">EventType :String</span><span class="sxs-lookup"><span data-stu-id="69065-164">EventType :String</span></span>

<span data-ttu-id="69065-165">Gibt das einem Ereignishandler zugeordnete Ereignis an.</span><span class="sxs-lookup"><span data-stu-id="69065-165">Specifies the event associated with an event handler.</span></span>

##### <a name="type"></a><span data-ttu-id="69065-166">Typ:</span><span class="sxs-lookup"><span data-stu-id="69065-166">Type:</span></span>

*   <span data-ttu-id="69065-167">Zeichenfolge</span><span class="sxs-lookup"><span data-stu-id="69065-167">String</span></span>

##### <a name="properties"></a><span data-ttu-id="69065-168">Eigenschaften:</span><span class="sxs-lookup"><span data-stu-id="69065-168">Properties:</span></span>

| <span data-ttu-id="69065-169">Name</span><span class="sxs-lookup"><span data-stu-id="69065-169">Name</span></span> | <span data-ttu-id="69065-170">Typ</span><span class="sxs-lookup"><span data-stu-id="69065-170">Type</span></span> | <span data-ttu-id="69065-171">Beschreibung</span><span class="sxs-lookup"><span data-stu-id="69065-171">Description</span></span> | <span data-ttu-id="69065-172">Mindestanforderung set</span><span class="sxs-lookup"><span data-stu-id="69065-172">Minimum requirement set</span></span> |
|---|---|---|---|
|`AppointmentTimeChanged`| <span data-ttu-id="69065-173">Zeichenfolge</span><span class="sxs-lookup"><span data-stu-id="69065-173">String</span></span> | <span data-ttu-id="69065-174">Der Termin Datums- oder Zeitwerte für die ausgewählte Datenreihe geändert wurde.</span><span class="sxs-lookup"><span data-stu-id="69065-174">The appointment date or time of the selected series has changed.</span></span> | <span data-ttu-id="69065-175">Vorschau</span><span class="sxs-lookup"><span data-stu-id="69065-175">Preview</span></span> |
|`ItemChanged`| <span data-ttu-id="69065-176">String</span><span class="sxs-lookup"><span data-stu-id="69065-176">String</span></span> | <span data-ttu-id="69065-177">Das ausgewählte Element wurde geändert.</span><span class="sxs-lookup"><span data-stu-id="69065-177">The selected item has changed.</span></span> | <span data-ttu-id="69065-178">1,5</span><span class="sxs-lookup"><span data-stu-id="69065-178">1.5</span></span> |
|`OfficeThemeChanged`| <span data-ttu-id="69065-179">String</span><span class="sxs-lookup"><span data-stu-id="69065-179">String</span></span> | <span data-ttu-id="69065-180">Das ausgewählte Element wurde geändert.</span><span class="sxs-lookup"><span data-stu-id="69065-180">The selected item has changed.</span></span> | <span data-ttu-id="69065-181">Vorschau</span><span class="sxs-lookup"><span data-stu-id="69065-181">Preview</span></span> |
|`RecipientsChanged`| <span data-ttu-id="69065-182">Zeichenfolge</span><span class="sxs-lookup"><span data-stu-id="69065-182">String</span></span> | <span data-ttu-id="69065-183">Die Empfängerliste des ausgewählten Elements wurde geändert.</span><span class="sxs-lookup"><span data-stu-id="69065-183">The recipient list of the selected item has changed.</span></span> | <span data-ttu-id="69065-184">Vorschau</span><span class="sxs-lookup"><span data-stu-id="69065-184">Preview</span></span> |
|`RecurrencePatternChanged`| <span data-ttu-id="69065-185">Zeichenfolge</span><span class="sxs-lookup"><span data-stu-id="69065-185">String</span></span> | <span data-ttu-id="69065-186">Das Serienmuster der ausgewählten Datenreihe geändert wurde.</span><span class="sxs-lookup"><span data-stu-id="69065-186">The recurrence pattern of the selected series has changed.</span></span> | <span data-ttu-id="69065-187">Vorschau</span><span class="sxs-lookup"><span data-stu-id="69065-187">Preview</span></span> |

##### <a name="requirements"></a><span data-ttu-id="69065-188">Anforderungen</span><span class="sxs-lookup"><span data-stu-id="69065-188">Requirements</span></span>

|<span data-ttu-id="69065-189">Anforderung</span><span class="sxs-lookup"><span data-stu-id="69065-189">Requirement</span></span>| <span data-ttu-id="69065-190">Wert</span><span class="sxs-lookup"><span data-stu-id="69065-190">Value</span></span>|
|---|---|
|[<span data-ttu-id="69065-191">Mindestversion des Postfachanforderungssatzes</span><span class="sxs-lookup"><span data-stu-id="69065-191">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="69065-192">1,5</span><span class="sxs-lookup"><span data-stu-id="69065-192">1.5</span></span> |
|[<span data-ttu-id="69065-193">Zutreffender Outlook-Modus</span><span class="sxs-lookup"><span data-stu-id="69065-193">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="69065-194">Verfassen oder Lesen</span><span class="sxs-lookup"><span data-stu-id="69065-194">Compose or read</span></span> |

---

####  <a name="sourceproperty-string"></a><span data-ttu-id="69065-195">SourceProperty :String</span><span class="sxs-lookup"><span data-stu-id="69065-195">SourceProperty :String</span></span>

<span data-ttu-id="69065-196">Gibt die Quelle der Daten an, die von der aufgerufenen Methode zurückgegeben werden.</span><span class="sxs-lookup"><span data-stu-id="69065-196">Specifies the source of the data returned by the invoked method.</span></span>

##### <a name="type"></a><span data-ttu-id="69065-197">Typ:</span><span class="sxs-lookup"><span data-stu-id="69065-197">Type:</span></span>

*   <span data-ttu-id="69065-198">Zeichenfolge</span><span class="sxs-lookup"><span data-stu-id="69065-198">String</span></span>

##### <a name="properties"></a><span data-ttu-id="69065-199">Eigenschaften:</span><span class="sxs-lookup"><span data-stu-id="69065-199">Properties:</span></span>

|<span data-ttu-id="69065-200">Name</span><span class="sxs-lookup"><span data-stu-id="69065-200">Name</span></span>| <span data-ttu-id="69065-201">Typ</span><span class="sxs-lookup"><span data-stu-id="69065-201">Type</span></span>| <span data-ttu-id="69065-202">Beschreibung</span><span class="sxs-lookup"><span data-stu-id="69065-202">Description</span></span>|
|---|---|---|
|`Body`| <span data-ttu-id="69065-203">Zeichenfolge</span><span class="sxs-lookup"><span data-stu-id="69065-203">String</span></span>|<span data-ttu-id="69065-204">Die Quelle der Daten stammt aus dem Textkörper einer Nachricht.</span><span class="sxs-lookup"><span data-stu-id="69065-204">The source of the data is from the body of a message.</span></span>|
|`Subject`| <span data-ttu-id="69065-205">Zeichenfolge</span><span class="sxs-lookup"><span data-stu-id="69065-205">String</span></span>|<span data-ttu-id="69065-206">Die Quelle der Daten stammt aus dem Betreff einer Nachricht.</span><span class="sxs-lookup"><span data-stu-id="69065-206">The source of the data is from the subject of a message.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="69065-207">Anforderungen</span><span class="sxs-lookup"><span data-stu-id="69065-207">Requirements</span></span>

|<span data-ttu-id="69065-208">Anforderung</span><span class="sxs-lookup"><span data-stu-id="69065-208">Requirement</span></span>| <span data-ttu-id="69065-209">Wert</span><span class="sxs-lookup"><span data-stu-id="69065-209">Value</span></span>|
|---|---|
|[<span data-ttu-id="69065-210">Mindestversion des Postfachanforderungssatzes</span><span class="sxs-lookup"><span data-stu-id="69065-210">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="69065-211">1.0</span><span class="sxs-lookup"><span data-stu-id="69065-211">1.0</span></span>|
|[<span data-ttu-id="69065-212">Zutreffender Outlook-Modus</span><span class="sxs-lookup"><span data-stu-id="69065-212">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="69065-213">Verfassen oder Lesen</span><span class="sxs-lookup"><span data-stu-id="69065-213">Compose or read</span></span>|