 

# <a name="office"></a><span data-ttu-id="7ee72-101">Büro</span><span class="sxs-lookup"><span data-stu-id="7ee72-101">Office</span></span>

<span data-ttu-id="7ee72-p101">Der Office-Namespace stellt gemeinsam genutzte Oberflächen bereit, die von Add-Ins in allen Office-Apps verwendet werden. Diese Auflistung dokumentiert nur die Schnittstellen, die von Outlook-Add-Ins verwendet werden. Eine vollständige Auflistung des Office-Namespaces finden Sie im Thema zur [freigegebenen API](/javascript/api/office).</span><span class="sxs-lookup"><span data-stu-id="7ee72-p101">The Office namespace provides shared interfaces that are used by add-ins in all of the Office apps. This listing documents only those interfaces that are used by Outlook add-ins. For a full listing of the Office namespace, see the [Shared API](/javascript/api/office).</span></span>

##### <a name="requirements"></a><span data-ttu-id="7ee72-104">Anforderungen</span><span class="sxs-lookup"><span data-stu-id="7ee72-104">Requirements</span></span>

|<span data-ttu-id="7ee72-105">Anforderung</span><span class="sxs-lookup"><span data-stu-id="7ee72-105">Requirement</span></span>| <span data-ttu-id="7ee72-106">Wert</span><span class="sxs-lookup"><span data-stu-id="7ee72-106">Value</span></span>|
|---|---|
|[<span data-ttu-id="7ee72-107">Mindestversion des Postfachanforderungssatzes</span><span class="sxs-lookup"><span data-stu-id="7ee72-107">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="7ee72-108">1.0</span><span class="sxs-lookup"><span data-stu-id="7ee72-108">1.0</span></span>|
|[<span data-ttu-id="7ee72-109">Zutreffender Outlook-Modus</span><span class="sxs-lookup"><span data-stu-id="7ee72-109">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="7ee72-110">Verfassen oder Lesen</span><span class="sxs-lookup"><span data-stu-id="7ee72-110">Compose or read</span></span>|

##### <a name="members-and-methods"></a><span data-ttu-id="7ee72-111">Elemente und Methoden</span><span class="sxs-lookup"><span data-stu-id="7ee72-111">Members and methods</span></span>

| <span data-ttu-id="7ee72-112">Element</span><span class="sxs-lookup"><span data-stu-id="7ee72-112">Member</span></span> | <span data-ttu-id="7ee72-113">Typ</span><span class="sxs-lookup"><span data-stu-id="7ee72-113">Type</span></span> |
|--------|------|
| [<span data-ttu-id="7ee72-114">AsyncResultStatus</span><span class="sxs-lookup"><span data-stu-id="7ee72-114">AsyncResultStatus</span></span>](#asyncresultstatus-string) | <span data-ttu-id="7ee72-115">Element</span><span class="sxs-lookup"><span data-stu-id="7ee72-115">Member</span></span> |
| [<span data-ttu-id="7ee72-116">CoercionType</span><span class="sxs-lookup"><span data-stu-id="7ee72-116">CoercionType</span></span>](#coerciontype-string) | <span data-ttu-id="7ee72-117">Element</span><span class="sxs-lookup"><span data-stu-id="7ee72-117">Member</span></span> |
| [<span data-ttu-id="7ee72-118">EventType</span><span class="sxs-lookup"><span data-stu-id="7ee72-118">EventType</span></span>](#eventtype-string) | <span data-ttu-id="7ee72-119">Element</span><span class="sxs-lookup"><span data-stu-id="7ee72-119">Member</span></span> |
| [<span data-ttu-id="7ee72-120">SourceProperty</span><span class="sxs-lookup"><span data-stu-id="7ee72-120">SourceProperty</span></span>](#sourceproperty-string) | <span data-ttu-id="7ee72-121">Element</span><span class="sxs-lookup"><span data-stu-id="7ee72-121">Member</span></span> |

### <a name="namespaces"></a><span data-ttu-id="7ee72-122">Namespaces</span><span class="sxs-lookup"><span data-stu-id="7ee72-122">Namespaces</span></span>

<span data-ttu-id="7ee72-123">[Kontextmenü](office.context.md): gemeinsam genutzten Oberflächen aus dem Office-Add-ins-API-Kontext-Namespace für die Verwendung in der Outlook-add-in-API bietet.</span><span class="sxs-lookup"><span data-stu-id="7ee72-123">[context](office.context.md): Provides shared interfaces from the Office Add-ins API's context namespace for use in the Outlook add-in API.</span></span>

<span data-ttu-id="7ee72-124">[MailboxEnums](/javascript/api/outlook/office.mailboxenums.attachmenttype): Enthält die Enumerationen ItemType, EntityType, AttachmentType, RecipientType, ResponseType und ItemNotificationMessageType.</span><span class="sxs-lookup"><span data-stu-id="7ee72-124">[MailboxEnums](/javascript/api/outlook/office.mailboxenums.attachmenttype): Includes the ItemType, EntityType, AttachmentType, RecipientType, ResponseType, and ItemNotificationMessageType enumerations.</span></span>

### <a name="members"></a><span data-ttu-id="7ee72-125">Elemente</span><span class="sxs-lookup"><span data-stu-id="7ee72-125">Members</span></span>

####  <a name="asyncresultstatus-string"></a><span data-ttu-id="7ee72-126">AsyncResultStatus :String</span><span class="sxs-lookup"><span data-stu-id="7ee72-126">AsyncResultStatus :String</span></span>

<span data-ttu-id="7ee72-127">Gibt das Ergebnis eines asynchronen Aufrufs an.</span><span class="sxs-lookup"><span data-stu-id="7ee72-127">Specifies the result of an asynchronous call.</span></span>

##### <a name="type"></a><span data-ttu-id="7ee72-128">Typ:</span><span class="sxs-lookup"><span data-stu-id="7ee72-128">Type:</span></span>

*   <span data-ttu-id="7ee72-129">Zeichenfolge</span><span class="sxs-lookup"><span data-stu-id="7ee72-129">String</span></span>

##### <a name="properties"></a><span data-ttu-id="7ee72-130">Eigenschaften:</span><span class="sxs-lookup"><span data-stu-id="7ee72-130">Properties:</span></span>

|<span data-ttu-id="7ee72-131">Name</span><span class="sxs-lookup"><span data-stu-id="7ee72-131">Name</span></span>| <span data-ttu-id="7ee72-132">Typ</span><span class="sxs-lookup"><span data-stu-id="7ee72-132">Type</span></span>| <span data-ttu-id="7ee72-133">Beschreibung</span><span class="sxs-lookup"><span data-stu-id="7ee72-133">Description</span></span>|
|---|---|---|
|`Succeeded`| <span data-ttu-id="7ee72-134">Zeichenfolge</span><span class="sxs-lookup"><span data-stu-id="7ee72-134">String</span></span>|<span data-ttu-id="7ee72-135">Der Aufruf war erfolgreich.</span><span class="sxs-lookup"><span data-stu-id="7ee72-135">The call succeeded.</span></span>|
|`Failed`| <span data-ttu-id="7ee72-136">Zeichenfolge</span><span class="sxs-lookup"><span data-stu-id="7ee72-136">String</span></span>|<span data-ttu-id="7ee72-137">Der Aufruf ist fehlerhaft.</span><span class="sxs-lookup"><span data-stu-id="7ee72-137">The call failed.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="7ee72-138">Anforderungen</span><span class="sxs-lookup"><span data-stu-id="7ee72-138">Requirements</span></span>

|<span data-ttu-id="7ee72-139">Anforderung</span><span class="sxs-lookup"><span data-stu-id="7ee72-139">Requirement</span></span>| <span data-ttu-id="7ee72-140">Wert</span><span class="sxs-lookup"><span data-stu-id="7ee72-140">Value</span></span>|
|---|---|
|[<span data-ttu-id="7ee72-141">Mindestversion des Postfachanforderungssatzes</span><span class="sxs-lookup"><span data-stu-id="7ee72-141">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="7ee72-142">1.0</span><span class="sxs-lookup"><span data-stu-id="7ee72-142">1.0</span></span>|
|[<span data-ttu-id="7ee72-143">Zutreffender Outlook-Modus</span><span class="sxs-lookup"><span data-stu-id="7ee72-143">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="7ee72-144">Verfassen oder Lesen</span><span class="sxs-lookup"><span data-stu-id="7ee72-144">Compose or read</span></span>|

---

####  <a name="coerciontype-string"></a><span data-ttu-id="7ee72-145">CoercionType :String</span><span class="sxs-lookup"><span data-stu-id="7ee72-145">CoercionType :String</span></span>

<span data-ttu-id="7ee72-146">Gibt an, wie Daten umgewandelt werden sollen, die von der aufgerufenen Methode zurückgegeben oder festgelegt wurden.</span><span class="sxs-lookup"><span data-stu-id="7ee72-146">Specifies how to coerce data returned or set by the invoked method.</span></span>

##### <a name="type"></a><span data-ttu-id="7ee72-147">Typ:</span><span class="sxs-lookup"><span data-stu-id="7ee72-147">Type:</span></span>

*   <span data-ttu-id="7ee72-148">Zeichenfolge</span><span class="sxs-lookup"><span data-stu-id="7ee72-148">String</span></span>

##### <a name="properties"></a><span data-ttu-id="7ee72-149">Eigenschaften:</span><span class="sxs-lookup"><span data-stu-id="7ee72-149">Properties:</span></span>

|<span data-ttu-id="7ee72-150">Name</span><span class="sxs-lookup"><span data-stu-id="7ee72-150">Name</span></span>| <span data-ttu-id="7ee72-151">Typ</span><span class="sxs-lookup"><span data-stu-id="7ee72-151">Type</span></span>| <span data-ttu-id="7ee72-152">Beschreibung</span><span class="sxs-lookup"><span data-stu-id="7ee72-152">Description</span></span>|
|---|---|---|
|`Html`| <span data-ttu-id="7ee72-153">Zeichenfolge</span><span class="sxs-lookup"><span data-stu-id="7ee72-153">String</span></span>|<span data-ttu-id="7ee72-154">Fordert die Rückgabe der Daten im HTML-Format an.</span><span class="sxs-lookup"><span data-stu-id="7ee72-154">Requests the data be returned in HTML format.</span></span>|
|`Text`| <span data-ttu-id="7ee72-155">Zeichenfolge</span><span class="sxs-lookup"><span data-stu-id="7ee72-155">String</span></span>|<span data-ttu-id="7ee72-156">Fordert die Rückgabe der Daten im Textformat an.</span><span class="sxs-lookup"><span data-stu-id="7ee72-156">Requests the data be returned in text format.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="7ee72-157">Anforderungen</span><span class="sxs-lookup"><span data-stu-id="7ee72-157">Requirements</span></span>

|<span data-ttu-id="7ee72-158">Anforderung</span><span class="sxs-lookup"><span data-stu-id="7ee72-158">Requirement</span></span>| <span data-ttu-id="7ee72-159">Wert</span><span class="sxs-lookup"><span data-stu-id="7ee72-159">Value</span></span>|
|---|---|
|[<span data-ttu-id="7ee72-160">Mindestversion des Postfachanforderungssatzes</span><span class="sxs-lookup"><span data-stu-id="7ee72-160">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="7ee72-161">1.0</span><span class="sxs-lookup"><span data-stu-id="7ee72-161">1.0</span></span>|
|[<span data-ttu-id="7ee72-162">Zutreffender Outlook-Modus</span><span class="sxs-lookup"><span data-stu-id="7ee72-162">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="7ee72-163">Verfassen oder Lesen</span><span class="sxs-lookup"><span data-stu-id="7ee72-163">Compose or read</span></span>|

---

####  <a name="eventtype-string"></a><span data-ttu-id="7ee72-164">EventType :String</span><span class="sxs-lookup"><span data-stu-id="7ee72-164">EventType :String</span></span>

<span data-ttu-id="7ee72-165">Gibt das einem Ereignishandler zugeordnete Ereignis an.</span><span class="sxs-lookup"><span data-stu-id="7ee72-165">Specifies the event associated with an event handler.</span></span>

##### <a name="type"></a><span data-ttu-id="7ee72-166">Typ:</span><span class="sxs-lookup"><span data-stu-id="7ee72-166">Type:</span></span>

*   <span data-ttu-id="7ee72-167">Zeichenfolge</span><span class="sxs-lookup"><span data-stu-id="7ee72-167">String</span></span>

##### <a name="properties"></a><span data-ttu-id="7ee72-168">Eigenschaften:</span><span class="sxs-lookup"><span data-stu-id="7ee72-168">Properties:</span></span>

| <span data-ttu-id="7ee72-169">Name</span><span class="sxs-lookup"><span data-stu-id="7ee72-169">Name</span></span> | <span data-ttu-id="7ee72-170">Typ</span><span class="sxs-lookup"><span data-stu-id="7ee72-170">Type</span></span> | <span data-ttu-id="7ee72-171">Beschreibung</span><span class="sxs-lookup"><span data-stu-id="7ee72-171">Description</span></span> | <span data-ttu-id="7ee72-172">Mindestanforderung set</span><span class="sxs-lookup"><span data-stu-id="7ee72-172">Minimum requirement set</span></span> |
|---|---|---|---|
|`AppointmentTimeChanged`| <span data-ttu-id="7ee72-173">Zeichenfolge</span><span class="sxs-lookup"><span data-stu-id="7ee72-173">String</span></span> | <span data-ttu-id="7ee72-174">Datum und Uhrzeit der ausgewählten Termins oder einer Datenreihe geändert wurde.</span><span class="sxs-lookup"><span data-stu-id="7ee72-174">The date or time of the selected appointment or series has changed.</span></span> | <span data-ttu-id="7ee72-175">1.7</span><span class="sxs-lookup"><span data-stu-id="7ee72-175">1.7</span></span> |
|`ItemChanged`| <span data-ttu-id="7ee72-176">String</span><span class="sxs-lookup"><span data-stu-id="7ee72-176">String</span></span> | <span data-ttu-id="7ee72-177">Das ausgewählte Element wurde geändert.</span><span class="sxs-lookup"><span data-stu-id="7ee72-177">The selected item has changed.</span></span> | <span data-ttu-id="7ee72-178">1,5</span><span class="sxs-lookup"><span data-stu-id="7ee72-178">1.5</span></span> |
|`OfficeThemeChanged`| <span data-ttu-id="7ee72-179">String</span><span class="sxs-lookup"><span data-stu-id="7ee72-179">String</span></span> | <span data-ttu-id="7ee72-180">Das ausgewählte Element wurde geändert.</span><span class="sxs-lookup"><span data-stu-id="7ee72-180">The selected item has changed.</span></span> | <span data-ttu-id="7ee72-181">Vorschau</span><span class="sxs-lookup"><span data-stu-id="7ee72-181">Preview</span></span> |
|`RecipientsChanged`| <span data-ttu-id="7ee72-182">Zeichenfolge</span><span class="sxs-lookup"><span data-stu-id="7ee72-182">String</span></span> | <span data-ttu-id="7ee72-183">Die Empfängerliste des ausgewählten Elements oder eines Termins Speicherorts wurde geändert.</span><span class="sxs-lookup"><span data-stu-id="7ee72-183">The recipient list of the selected item or appointment location has changed.</span></span> | <span data-ttu-id="7ee72-184">1.7</span><span class="sxs-lookup"><span data-stu-id="7ee72-184">1.7</span></span> |
|`RecurrenceChanged`| <span data-ttu-id="7ee72-185">Zeichenfolge</span><span class="sxs-lookup"><span data-stu-id="7ee72-185">String</span></span> | <span data-ttu-id="7ee72-186">Das Serienmuster der ausgewählten Datenreihe geändert wurde.</span><span class="sxs-lookup"><span data-stu-id="7ee72-186">The recurrence pattern of the selected series has changed.</span></span> | <span data-ttu-id="7ee72-187">1.7</span><span class="sxs-lookup"><span data-stu-id="7ee72-187">1.7</span></span> |

##### <a name="requirements"></a><span data-ttu-id="7ee72-188">Anforderungen</span><span class="sxs-lookup"><span data-stu-id="7ee72-188">Requirements</span></span>

|<span data-ttu-id="7ee72-189">Anforderung</span><span class="sxs-lookup"><span data-stu-id="7ee72-189">Requirement</span></span>| <span data-ttu-id="7ee72-190">Wert</span><span class="sxs-lookup"><span data-stu-id="7ee72-190">Value</span></span>|
|---|---|
|[<span data-ttu-id="7ee72-191">Mindestversion des Postfachanforderungssatzes</span><span class="sxs-lookup"><span data-stu-id="7ee72-191">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="7ee72-192">1,5</span><span class="sxs-lookup"><span data-stu-id="7ee72-192">1.5</span></span> |
|[<span data-ttu-id="7ee72-193">Zutreffender Outlook-Modus</span><span class="sxs-lookup"><span data-stu-id="7ee72-193">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="7ee72-194">Verfassen oder Lesen</span><span class="sxs-lookup"><span data-stu-id="7ee72-194">Compose or read</span></span> |

---

####  <a name="sourceproperty-string"></a><span data-ttu-id="7ee72-195">SourceProperty :String</span><span class="sxs-lookup"><span data-stu-id="7ee72-195">SourceProperty :String</span></span>

<span data-ttu-id="7ee72-196">Gibt die Quelle der Daten an, die von der aufgerufenen Methode zurückgegeben werden.</span><span class="sxs-lookup"><span data-stu-id="7ee72-196">Specifies the source of the data returned by the invoked method.</span></span>

##### <a name="type"></a><span data-ttu-id="7ee72-197">Typ:</span><span class="sxs-lookup"><span data-stu-id="7ee72-197">Type:</span></span>

*   <span data-ttu-id="7ee72-198">Zeichenfolge</span><span class="sxs-lookup"><span data-stu-id="7ee72-198">String</span></span>

##### <a name="properties"></a><span data-ttu-id="7ee72-199">Eigenschaften:</span><span class="sxs-lookup"><span data-stu-id="7ee72-199">Properties:</span></span>

|<span data-ttu-id="7ee72-200">Name</span><span class="sxs-lookup"><span data-stu-id="7ee72-200">Name</span></span>| <span data-ttu-id="7ee72-201">Typ</span><span class="sxs-lookup"><span data-stu-id="7ee72-201">Type</span></span>| <span data-ttu-id="7ee72-202">Beschreibung</span><span class="sxs-lookup"><span data-stu-id="7ee72-202">Description</span></span>|
|---|---|---|
|`Body`| <span data-ttu-id="7ee72-203">Zeichenfolge</span><span class="sxs-lookup"><span data-stu-id="7ee72-203">String</span></span>|<span data-ttu-id="7ee72-204">Die Quelle der Daten stammt aus dem Textkörper einer Nachricht.</span><span class="sxs-lookup"><span data-stu-id="7ee72-204">The source of the data is from the body of a message.</span></span>|
|`Subject`| <span data-ttu-id="7ee72-205">Zeichenfolge</span><span class="sxs-lookup"><span data-stu-id="7ee72-205">String</span></span>|<span data-ttu-id="7ee72-206">Die Quelle der Daten stammt aus dem Betreff einer Nachricht.</span><span class="sxs-lookup"><span data-stu-id="7ee72-206">The source of the data is from the subject of a message.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="7ee72-207">Anforderungen</span><span class="sxs-lookup"><span data-stu-id="7ee72-207">Requirements</span></span>

|<span data-ttu-id="7ee72-208">Anforderung</span><span class="sxs-lookup"><span data-stu-id="7ee72-208">Requirement</span></span>| <span data-ttu-id="7ee72-209">Wert</span><span class="sxs-lookup"><span data-stu-id="7ee72-209">Value</span></span>|
|---|---|
|[<span data-ttu-id="7ee72-210">Mindestversion des Postfachanforderungssatzes</span><span class="sxs-lookup"><span data-stu-id="7ee72-210">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="7ee72-211">1.0</span><span class="sxs-lookup"><span data-stu-id="7ee72-211">1.0</span></span>|
|[<span data-ttu-id="7ee72-212">Zutreffender Outlook-Modus</span><span class="sxs-lookup"><span data-stu-id="7ee72-212">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="7ee72-213">Verfassen oder Lesen</span><span class="sxs-lookup"><span data-stu-id="7ee72-213">Compose or read</span></span>|