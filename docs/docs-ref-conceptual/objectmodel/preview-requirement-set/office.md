 

# <a name="office"></a><span data-ttu-id="121aa-101">Büro</span><span class="sxs-lookup"><span data-stu-id="121aa-101">Office</span></span>

<span data-ttu-id="121aa-p101">Der Office-Namespace stellt gemeinsam genutzte Oberflächen bereit, die von Add-Ins in allen Office-Apps verwendet werden. Diese Auflistung dokumentiert nur die Schnittstellen, die von Outlook-Add-Ins verwendet werden. Eine vollständige Auflistung des Office-Namespaces finden Sie im Thema zur [freigegebenen API](/javascript/api/office).</span><span class="sxs-lookup"><span data-stu-id="121aa-p101">The Office namespace provides shared interfaces that are used by add-ins in all of the Office apps. This listing documents only those interfaces that are used by Outlook add-ins. For a full listing of the Office namespace, see the [Shared API](/javascript/api/office).</span></span>

##### <a name="requirements"></a><span data-ttu-id="121aa-104">Anforderungen</span><span class="sxs-lookup"><span data-stu-id="121aa-104">Requirements</span></span>

|<span data-ttu-id="121aa-105">Anforderung</span><span class="sxs-lookup"><span data-stu-id="121aa-105">Requirement</span></span>| <span data-ttu-id="121aa-106">Wert</span><span class="sxs-lookup"><span data-stu-id="121aa-106">Value</span></span>|
|---|---|
|[<span data-ttu-id="121aa-107">Mindestversion des Postfachanforderungssatzes</span><span class="sxs-lookup"><span data-stu-id="121aa-107">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="121aa-108">1.0</span><span class="sxs-lookup"><span data-stu-id="121aa-108">1.0</span></span>|
|[<span data-ttu-id="121aa-109">Zutreffender Outlook-Modus</span><span class="sxs-lookup"><span data-stu-id="121aa-109">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="121aa-110">Verfassen oder Lesen</span><span class="sxs-lookup"><span data-stu-id="121aa-110">Compose or read</span></span>|

##### <a name="members-and-methods"></a><span data-ttu-id="121aa-111">Elemente und Methoden</span><span class="sxs-lookup"><span data-stu-id="121aa-111">Members and methods</span></span>

| <span data-ttu-id="121aa-112">Element</span><span class="sxs-lookup"><span data-stu-id="121aa-112">Member</span></span> | <span data-ttu-id="121aa-113">Typ</span><span class="sxs-lookup"><span data-stu-id="121aa-113">Type</span></span> |
|--------|------|
| [<span data-ttu-id="121aa-114">AsyncResultStatus</span><span class="sxs-lookup"><span data-stu-id="121aa-114">AsyncResultStatus</span></span>](#asyncresultstatus-string) | <span data-ttu-id="121aa-115">Element</span><span class="sxs-lookup"><span data-stu-id="121aa-115">Member</span></span> |
| [<span data-ttu-id="121aa-116">CoercionType</span><span class="sxs-lookup"><span data-stu-id="121aa-116">CoercionType</span></span>](#coerciontype-string) | <span data-ttu-id="121aa-117">Element</span><span class="sxs-lookup"><span data-stu-id="121aa-117">Member</span></span> |
| [<span data-ttu-id="121aa-118">EventType</span><span class="sxs-lookup"><span data-stu-id="121aa-118">EventType</span></span>](#eventtype-string) | <span data-ttu-id="121aa-119">Element</span><span class="sxs-lookup"><span data-stu-id="121aa-119">Member</span></span> |
| [<span data-ttu-id="121aa-120">SourceProperty</span><span class="sxs-lookup"><span data-stu-id="121aa-120">SourceProperty</span></span>](#sourceproperty-string) | <span data-ttu-id="121aa-121">Element</span><span class="sxs-lookup"><span data-stu-id="121aa-121">Member</span></span> |

### <a name="namespaces"></a><span data-ttu-id="121aa-122">Namespaces</span><span class="sxs-lookup"><span data-stu-id="121aa-122">Namespaces</span></span>

<span data-ttu-id="121aa-123">[Kontextmenü](office.context.md): gemeinsam genutzten Oberflächen aus dem Office-Add-ins-API-Kontext-Namespace für die Verwendung in der Outlook-add-in-API bietet.</span><span class="sxs-lookup"><span data-stu-id="121aa-123">[context](office.context.md): Provides shared interfaces from the Office Add-ins API's context namespace for use in the Outlook add-in API.</span></span>

<span data-ttu-id="121aa-124">[MailboxEnums](/javascript/api/outlook/office.mailboxenums.attachmenttype): Enthält die Enumerationen ItemType, EntityType, AttachmentType, RecipientType, ResponseType und ItemNotificationMessageType.</span><span class="sxs-lookup"><span data-stu-id="121aa-124">[MailboxEnums](/javascript/api/outlook/office.mailboxenums.attachmenttype): Includes the ItemType, EntityType, AttachmentType, RecipientType, ResponseType, and ItemNotificationMessageType enumerations.</span></span>

### <a name="members"></a><span data-ttu-id="121aa-125">Elemente</span><span class="sxs-lookup"><span data-stu-id="121aa-125">Members</span></span>

####  <a name="asyncresultstatus-string"></a><span data-ttu-id="121aa-126">AsyncResultStatus :String</span><span class="sxs-lookup"><span data-stu-id="121aa-126">AsyncResultStatus :String</span></span>

<span data-ttu-id="121aa-127">Gibt das Ergebnis eines asynchronen Aufrufs an.</span><span class="sxs-lookup"><span data-stu-id="121aa-127">Specifies the result of an asynchronous call.</span></span>

##### <a name="type"></a><span data-ttu-id="121aa-128">Typ:</span><span class="sxs-lookup"><span data-stu-id="121aa-128">Type:</span></span>

*   <span data-ttu-id="121aa-129">Zeichenfolge</span><span class="sxs-lookup"><span data-stu-id="121aa-129">String</span></span>

##### <a name="properties"></a><span data-ttu-id="121aa-130">Eigenschaften:</span><span class="sxs-lookup"><span data-stu-id="121aa-130">Properties:</span></span>

|<span data-ttu-id="121aa-131">Name</span><span class="sxs-lookup"><span data-stu-id="121aa-131">Name</span></span>| <span data-ttu-id="121aa-132">Typ</span><span class="sxs-lookup"><span data-stu-id="121aa-132">Type</span></span>| <span data-ttu-id="121aa-133">Beschreibung</span><span class="sxs-lookup"><span data-stu-id="121aa-133">Description</span></span>|
|---|---|---|
|`Succeeded`| <span data-ttu-id="121aa-134">Zeichenfolge</span><span class="sxs-lookup"><span data-stu-id="121aa-134">String</span></span>|<span data-ttu-id="121aa-135">Der Aufruf war erfolgreich.</span><span class="sxs-lookup"><span data-stu-id="121aa-135">The call succeeded.</span></span>|
|`Failed`| <span data-ttu-id="121aa-136">Zeichenfolge</span><span class="sxs-lookup"><span data-stu-id="121aa-136">String</span></span>|<span data-ttu-id="121aa-137">Der Aufruf ist fehlerhaft.</span><span class="sxs-lookup"><span data-stu-id="121aa-137">The call failed.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="121aa-138">Anforderungen</span><span class="sxs-lookup"><span data-stu-id="121aa-138">Requirements</span></span>

|<span data-ttu-id="121aa-139">Anforderung</span><span class="sxs-lookup"><span data-stu-id="121aa-139">Requirement</span></span>| <span data-ttu-id="121aa-140">Wert</span><span class="sxs-lookup"><span data-stu-id="121aa-140">Value</span></span>|
|---|---|
|[<span data-ttu-id="121aa-141">Mindestversion des Postfachanforderungssatzes</span><span class="sxs-lookup"><span data-stu-id="121aa-141">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="121aa-142">1.0</span><span class="sxs-lookup"><span data-stu-id="121aa-142">1.0</span></span>|
|[<span data-ttu-id="121aa-143">Zutreffender Outlook-Modus</span><span class="sxs-lookup"><span data-stu-id="121aa-143">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="121aa-144">Verfassen oder Lesen</span><span class="sxs-lookup"><span data-stu-id="121aa-144">Compose or read</span></span>|

---

####  <a name="coerciontype-string"></a><span data-ttu-id="121aa-145">CoercionType :String</span><span class="sxs-lookup"><span data-stu-id="121aa-145">CoercionType :String</span></span>

<span data-ttu-id="121aa-146">Gibt an, wie Daten umgewandelt werden sollen, die von der aufgerufenen Methode zurückgegeben oder festgelegt wurden.</span><span class="sxs-lookup"><span data-stu-id="121aa-146">Specifies how to coerce data returned or set by the invoked method.</span></span>

##### <a name="type"></a><span data-ttu-id="121aa-147">Typ:</span><span class="sxs-lookup"><span data-stu-id="121aa-147">Type:</span></span>

*   <span data-ttu-id="121aa-148">Zeichenfolge</span><span class="sxs-lookup"><span data-stu-id="121aa-148">String</span></span>

##### <a name="properties"></a><span data-ttu-id="121aa-149">Eigenschaften:</span><span class="sxs-lookup"><span data-stu-id="121aa-149">Properties:</span></span>

|<span data-ttu-id="121aa-150">Name</span><span class="sxs-lookup"><span data-stu-id="121aa-150">Name</span></span>| <span data-ttu-id="121aa-151">Typ</span><span class="sxs-lookup"><span data-stu-id="121aa-151">Type</span></span>| <span data-ttu-id="121aa-152">Beschreibung</span><span class="sxs-lookup"><span data-stu-id="121aa-152">Description</span></span>|
|---|---|---|
|`Html`| <span data-ttu-id="121aa-153">Zeichenfolge</span><span class="sxs-lookup"><span data-stu-id="121aa-153">String</span></span>|<span data-ttu-id="121aa-154">Fordert die Rückgabe der Daten im HTML-Format an.</span><span class="sxs-lookup"><span data-stu-id="121aa-154">Requests the data be returned in HTML format.</span></span>|
|`Text`| <span data-ttu-id="121aa-155">Zeichenfolge</span><span class="sxs-lookup"><span data-stu-id="121aa-155">String</span></span>|<span data-ttu-id="121aa-156">Fordert die Rückgabe der Daten im Textformat an.</span><span class="sxs-lookup"><span data-stu-id="121aa-156">Requests the data be returned in text format.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="121aa-157">Anforderungen</span><span class="sxs-lookup"><span data-stu-id="121aa-157">Requirements</span></span>

|<span data-ttu-id="121aa-158">Anforderung</span><span class="sxs-lookup"><span data-stu-id="121aa-158">Requirement</span></span>| <span data-ttu-id="121aa-159">Wert</span><span class="sxs-lookup"><span data-stu-id="121aa-159">Value</span></span>|
|---|---|
|[<span data-ttu-id="121aa-160">Mindestversion des Postfachanforderungssatzes</span><span class="sxs-lookup"><span data-stu-id="121aa-160">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="121aa-161">1.0</span><span class="sxs-lookup"><span data-stu-id="121aa-161">1.0</span></span>|
|[<span data-ttu-id="121aa-162">Zutreffender Outlook-Modus</span><span class="sxs-lookup"><span data-stu-id="121aa-162">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="121aa-163">Verfassen oder Lesen</span><span class="sxs-lookup"><span data-stu-id="121aa-163">Compose or read</span></span>|

---

####  <a name="eventtype-string"></a><span data-ttu-id="121aa-164">EventType :String</span><span class="sxs-lookup"><span data-stu-id="121aa-164">EventType :String</span></span>

<span data-ttu-id="121aa-165">Gibt das einem Ereignishandler zugeordnete Ereignis an.</span><span class="sxs-lookup"><span data-stu-id="121aa-165">Specifies the event associated with an event handler.</span></span>

##### <a name="type"></a><span data-ttu-id="121aa-166">Typ:</span><span class="sxs-lookup"><span data-stu-id="121aa-166">Type:</span></span>

*   <span data-ttu-id="121aa-167">Zeichenfolge</span><span class="sxs-lookup"><span data-stu-id="121aa-167">String</span></span>

##### <a name="properties"></a><span data-ttu-id="121aa-168">Eigenschaften:</span><span class="sxs-lookup"><span data-stu-id="121aa-168">Properties:</span></span>

| <span data-ttu-id="121aa-169">Name</span><span class="sxs-lookup"><span data-stu-id="121aa-169">Name</span></span> | <span data-ttu-id="121aa-170">Typ</span><span class="sxs-lookup"><span data-stu-id="121aa-170">Type</span></span> | <span data-ttu-id="121aa-171">Beschreibung</span><span class="sxs-lookup"><span data-stu-id="121aa-171">Description</span></span> | <span data-ttu-id="121aa-172">Mindestanforderung set</span><span class="sxs-lookup"><span data-stu-id="121aa-172">Minimum requirement set</span></span> |
|---|---|---|---|
|`AppointmentTimeChanged`| <span data-ttu-id="121aa-173">Zeichenfolge</span><span class="sxs-lookup"><span data-stu-id="121aa-173">String</span></span> | <span data-ttu-id="121aa-174">Datum und Uhrzeit der ausgewählten Termins oder einer Datenreihe geändert wurde.</span><span class="sxs-lookup"><span data-stu-id="121aa-174">The date or time of the selected appointment or series has changed.</span></span> | <span data-ttu-id="121aa-175">1.7</span><span class="sxs-lookup"><span data-stu-id="121aa-175">1.7</span></span> |
|`ItemChanged`| <span data-ttu-id="121aa-176">String</span><span class="sxs-lookup"><span data-stu-id="121aa-176">String</span></span> | <span data-ttu-id="121aa-177">Das ausgewählte Element wurde geändert.</span><span class="sxs-lookup"><span data-stu-id="121aa-177">The selected item has changed.</span></span> | <span data-ttu-id="121aa-178">1,5</span><span class="sxs-lookup"><span data-stu-id="121aa-178">1.5</span></span> |
|`OfficeThemeChanged`| <span data-ttu-id="121aa-179">String</span><span class="sxs-lookup"><span data-stu-id="121aa-179">String</span></span> | <span data-ttu-id="121aa-180">Das ausgewählte Element wurde geändert.</span><span class="sxs-lookup"><span data-stu-id="121aa-180">The selected item has changed.</span></span> | <span data-ttu-id="121aa-181">Vorschau</span><span class="sxs-lookup"><span data-stu-id="121aa-181">Preview</span></span> |
|`RecipientsChanged`| <span data-ttu-id="121aa-182">Zeichenfolge</span><span class="sxs-lookup"><span data-stu-id="121aa-182">String</span></span> | <span data-ttu-id="121aa-183">Die Empfängerliste des ausgewählten Elements oder eines Termins Speicherorts wurde geändert.</span><span class="sxs-lookup"><span data-stu-id="121aa-183">The recipient list of the selected item or appointment location has changed.</span></span> | <span data-ttu-id="121aa-184">1.7</span><span class="sxs-lookup"><span data-stu-id="121aa-184">1.7</span></span> |
|`RecurrenceChanged`| <span data-ttu-id="121aa-185">Zeichenfolge</span><span class="sxs-lookup"><span data-stu-id="121aa-185">String</span></span> | <span data-ttu-id="121aa-186">Das Serienmuster der ausgewählten Datenreihe geändert wurde.</span><span class="sxs-lookup"><span data-stu-id="121aa-186">The recurrence pattern of the selected series has changed.</span></span> | <span data-ttu-id="121aa-187">1.7</span><span class="sxs-lookup"><span data-stu-id="121aa-187">1.7</span></span> |

##### <a name="requirements"></a><span data-ttu-id="121aa-188">Anforderungen</span><span class="sxs-lookup"><span data-stu-id="121aa-188">Requirements</span></span>

|<span data-ttu-id="121aa-189">Anforderung</span><span class="sxs-lookup"><span data-stu-id="121aa-189">Requirement</span></span>| <span data-ttu-id="121aa-190">Wert</span><span class="sxs-lookup"><span data-stu-id="121aa-190">Value</span></span>|
|---|---|
|[<span data-ttu-id="121aa-191">Mindestversion des Postfachanforderungssatzes</span><span class="sxs-lookup"><span data-stu-id="121aa-191">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="121aa-192">1,5</span><span class="sxs-lookup"><span data-stu-id="121aa-192">1.5</span></span> |
|[<span data-ttu-id="121aa-193">Zutreffender Outlook-Modus</span><span class="sxs-lookup"><span data-stu-id="121aa-193">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="121aa-194">Verfassen oder Lesen</span><span class="sxs-lookup"><span data-stu-id="121aa-194">Compose or read</span></span> |

---

####  <a name="sourceproperty-string"></a><span data-ttu-id="121aa-195">SourceProperty :String</span><span class="sxs-lookup"><span data-stu-id="121aa-195">SourceProperty :String</span></span>

<span data-ttu-id="121aa-196">Gibt die Quelle der Daten an, die von der aufgerufenen Methode zurückgegeben werden.</span><span class="sxs-lookup"><span data-stu-id="121aa-196">Specifies the source of the data returned by the invoked method.</span></span>

##### <a name="type"></a><span data-ttu-id="121aa-197">Typ:</span><span class="sxs-lookup"><span data-stu-id="121aa-197">Type:</span></span>

*   <span data-ttu-id="121aa-198">Zeichenfolge</span><span class="sxs-lookup"><span data-stu-id="121aa-198">String</span></span>

##### <a name="properties"></a><span data-ttu-id="121aa-199">Eigenschaften:</span><span class="sxs-lookup"><span data-stu-id="121aa-199">Properties:</span></span>

|<span data-ttu-id="121aa-200">Name</span><span class="sxs-lookup"><span data-stu-id="121aa-200">Name</span></span>| <span data-ttu-id="121aa-201">Typ</span><span class="sxs-lookup"><span data-stu-id="121aa-201">Type</span></span>| <span data-ttu-id="121aa-202">Beschreibung</span><span class="sxs-lookup"><span data-stu-id="121aa-202">Description</span></span>|
|---|---|---|
|`Body`| <span data-ttu-id="121aa-203">Zeichenfolge</span><span class="sxs-lookup"><span data-stu-id="121aa-203">String</span></span>|<span data-ttu-id="121aa-204">Die Quelle der Daten stammt aus dem Textkörper einer Nachricht.</span><span class="sxs-lookup"><span data-stu-id="121aa-204">The source of the data is from the body of a message.</span></span>|
|`Subject`| <span data-ttu-id="121aa-205">Zeichenfolge</span><span class="sxs-lookup"><span data-stu-id="121aa-205">String</span></span>|<span data-ttu-id="121aa-206">Die Quelle der Daten stammt aus dem Betreff einer Nachricht.</span><span class="sxs-lookup"><span data-stu-id="121aa-206">The source of the data is from the subject of a message.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="121aa-207">Anforderungen</span><span class="sxs-lookup"><span data-stu-id="121aa-207">Requirements</span></span>

|<span data-ttu-id="121aa-208">Anforderung</span><span class="sxs-lookup"><span data-stu-id="121aa-208">Requirement</span></span>| <span data-ttu-id="121aa-209">Wert</span><span class="sxs-lookup"><span data-stu-id="121aa-209">Value</span></span>|
|---|---|
|[<span data-ttu-id="121aa-210">Mindestversion des Postfachanforderungssatzes</span><span class="sxs-lookup"><span data-stu-id="121aa-210">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="121aa-211">1.0</span><span class="sxs-lookup"><span data-stu-id="121aa-211">1.0</span></span>|
|[<span data-ttu-id="121aa-212">Zutreffender Outlook-Modus</span><span class="sxs-lookup"><span data-stu-id="121aa-212">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="121aa-213">Verfassen oder Lesen</span><span class="sxs-lookup"><span data-stu-id="121aa-213">Compose or read</span></span>|