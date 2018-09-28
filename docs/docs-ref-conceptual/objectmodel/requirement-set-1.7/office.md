 

# <a name="office"></a><span data-ttu-id="124e4-101">Büro</span><span class="sxs-lookup"><span data-stu-id="124e4-101">Office</span></span>

<span data-ttu-id="124e4-p101">Der Office-Namespace stellt gemeinsam genutzte Oberflächen bereit, die von Add-Ins in allen Office-Apps verwendet werden. Diese Auflistung dokumentiert nur die Schnittstellen, die von Outlook-Add-Ins verwendet werden. Eine vollständige Auflistung des Office-Namespaces finden Sie im Thema zur [freigegebenen API](/javascript/api/office).</span><span class="sxs-lookup"><span data-stu-id="124e4-p101">The Office namespace provides shared interfaces that are used by add-ins in all of the Office apps. This listing documents only those interfaces that are used by Outlook add-ins. For a full listing of the Office namespace, see the [Shared API](/javascript/api/office).</span></span>

##### <a name="requirements"></a><span data-ttu-id="124e4-104">Anforderungen</span><span class="sxs-lookup"><span data-stu-id="124e4-104">Requirements</span></span>

|<span data-ttu-id="124e4-105">Anforderung</span><span class="sxs-lookup"><span data-stu-id="124e4-105">Requirement</span></span>| <span data-ttu-id="124e4-106">Wert</span><span class="sxs-lookup"><span data-stu-id="124e4-106">Value</span></span>|
|---|---|
|[<span data-ttu-id="124e4-107">Mindestversion des Postfachanforderungssatzes</span><span class="sxs-lookup"><span data-stu-id="124e4-107">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="124e4-108">1.0</span><span class="sxs-lookup"><span data-stu-id="124e4-108">1.0</span></span>|
|[<span data-ttu-id="124e4-109">Zutreffender Outlook-Modus</span><span class="sxs-lookup"><span data-stu-id="124e4-109">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="124e4-110">Verfassen oder Lesen</span><span class="sxs-lookup"><span data-stu-id="124e4-110">Compose or read</span></span>|

##### <a name="members-and-methods"></a><span data-ttu-id="124e4-111">Elemente und Methoden</span><span class="sxs-lookup"><span data-stu-id="124e4-111">Members and methods</span></span>

| <span data-ttu-id="124e4-112">Element</span><span class="sxs-lookup"><span data-stu-id="124e4-112">Member</span></span> | <span data-ttu-id="124e4-113">Typ</span><span class="sxs-lookup"><span data-stu-id="124e4-113">Type</span></span> |
|--------|------|
| [<span data-ttu-id="124e4-114">AsyncResultStatus</span><span class="sxs-lookup"><span data-stu-id="124e4-114">AsyncResultStatus</span></span>](#asyncresultstatus-string) | <span data-ttu-id="124e4-115">Element</span><span class="sxs-lookup"><span data-stu-id="124e4-115">Member</span></span> |
| [<span data-ttu-id="124e4-116">CoercionType</span><span class="sxs-lookup"><span data-stu-id="124e4-116">CoercionType</span></span>](#coerciontype-string) | <span data-ttu-id="124e4-117">Element</span><span class="sxs-lookup"><span data-stu-id="124e4-117">Member</span></span> |
| [<span data-ttu-id="124e4-118">EventType</span><span class="sxs-lookup"><span data-stu-id="124e4-118">EventType</span></span>](#eventtype-string) | <span data-ttu-id="124e4-119">Element</span><span class="sxs-lookup"><span data-stu-id="124e4-119">Member</span></span> |
| [<span data-ttu-id="124e4-120">SourceProperty</span><span class="sxs-lookup"><span data-stu-id="124e4-120">SourceProperty</span></span>](#sourceproperty-string) | <span data-ttu-id="124e4-121">Element</span><span class="sxs-lookup"><span data-stu-id="124e4-121">Member</span></span> |

### <a name="namespaces"></a><span data-ttu-id="124e4-122">Namespaces</span><span class="sxs-lookup"><span data-stu-id="124e4-122">Namespaces</span></span>

<span data-ttu-id="124e4-123">[Kontextmenü](office.context.md): gemeinsam genutzten Oberflächen aus dem Office-Add-ins-API-Kontext-Namespace für die Verwendung in der Outlook-add-in-API bietet.</span><span class="sxs-lookup"><span data-stu-id="124e4-123">[context](office.context.md): Provides shared interfaces from the Office Add-ins API's context namespace for use in the Outlook add-in API.</span></span>

<span data-ttu-id="124e4-124">[MailboxEnums](/javascript/api/outlook_1_7/office.mailboxenums.attachmenttype): Enthält die Enumerationen ItemType, EntityType, AttachmentType, RecipientType, ResponseType und ItemNotificationMessageType.</span><span class="sxs-lookup"><span data-stu-id="124e4-124">[MailboxEnums](/javascript/api/outlook_1_7/office.mailboxenums.attachmenttype): Includes the ItemType, EntityType, AttachmentType, RecipientType, ResponseType, and ItemNotificationMessageType enumerations.</span></span>

### <a name="members"></a><span data-ttu-id="124e4-125">Elemente</span><span class="sxs-lookup"><span data-stu-id="124e4-125">Members</span></span>

####  <a name="asyncresultstatus-string"></a><span data-ttu-id="124e4-126">AsyncResultStatus :String</span><span class="sxs-lookup"><span data-stu-id="124e4-126">AsyncResultStatus :String</span></span>

<span data-ttu-id="124e4-127">Gibt das Ergebnis eines asynchronen Aufrufs an.</span><span class="sxs-lookup"><span data-stu-id="124e4-127">Specifies the result of an asynchronous call.</span></span>

##### <a name="type"></a><span data-ttu-id="124e4-128">Typ:</span><span class="sxs-lookup"><span data-stu-id="124e4-128">Type:</span></span>

*   <span data-ttu-id="124e4-129">Zeichenfolge</span><span class="sxs-lookup"><span data-stu-id="124e4-129">String</span></span>

##### <a name="properties"></a><span data-ttu-id="124e4-130">Eigenschaften:</span><span class="sxs-lookup"><span data-stu-id="124e4-130">Properties:</span></span>

|<span data-ttu-id="124e4-131">Name</span><span class="sxs-lookup"><span data-stu-id="124e4-131">Name</span></span>| <span data-ttu-id="124e4-132">Typ</span><span class="sxs-lookup"><span data-stu-id="124e4-132">Type</span></span>| <span data-ttu-id="124e4-133">Beschreibung</span><span class="sxs-lookup"><span data-stu-id="124e4-133">Description</span></span>|
|---|---|---|
|`Succeeded`| <span data-ttu-id="124e4-134">Zeichenfolge</span><span class="sxs-lookup"><span data-stu-id="124e4-134">String</span></span>|<span data-ttu-id="124e4-135">Der Aufruf war erfolgreich.</span><span class="sxs-lookup"><span data-stu-id="124e4-135">The call succeeded.</span></span>|
|`Failed`| <span data-ttu-id="124e4-136">Zeichenfolge</span><span class="sxs-lookup"><span data-stu-id="124e4-136">String</span></span>|<span data-ttu-id="124e4-137">Der Aufruf ist fehlerhaft.</span><span class="sxs-lookup"><span data-stu-id="124e4-137">The call failed.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="124e4-138">Anforderungen</span><span class="sxs-lookup"><span data-stu-id="124e4-138">Requirements</span></span>

|<span data-ttu-id="124e4-139">Anforderung</span><span class="sxs-lookup"><span data-stu-id="124e4-139">Requirement</span></span>| <span data-ttu-id="124e4-140">Wert</span><span class="sxs-lookup"><span data-stu-id="124e4-140">Value</span></span>|
|---|---|
|[<span data-ttu-id="124e4-141">Mindestversion des Postfachanforderungssatzes</span><span class="sxs-lookup"><span data-stu-id="124e4-141">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="124e4-142">1.0</span><span class="sxs-lookup"><span data-stu-id="124e4-142">1.0</span></span>|
|[<span data-ttu-id="124e4-143">Zutreffender Outlook-Modus</span><span class="sxs-lookup"><span data-stu-id="124e4-143">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="124e4-144">Verfassen oder Lesen</span><span class="sxs-lookup"><span data-stu-id="124e4-144">Compose or read</span></span>|

---

####  <a name="coerciontype-string"></a><span data-ttu-id="124e4-145">CoercionType :String</span><span class="sxs-lookup"><span data-stu-id="124e4-145">CoercionType :String</span></span>

<span data-ttu-id="124e4-146">Gibt an, wie Daten umgewandelt werden sollen, die von der aufgerufenen Methode zurückgegeben oder festgelegt wurden.</span><span class="sxs-lookup"><span data-stu-id="124e4-146">Specifies how to coerce data returned or set by the invoked method.</span></span>

##### <a name="type"></a><span data-ttu-id="124e4-147">Typ:</span><span class="sxs-lookup"><span data-stu-id="124e4-147">Type:</span></span>

*   <span data-ttu-id="124e4-148">Zeichenfolge</span><span class="sxs-lookup"><span data-stu-id="124e4-148">String</span></span>

##### <a name="properties"></a><span data-ttu-id="124e4-149">Eigenschaften:</span><span class="sxs-lookup"><span data-stu-id="124e4-149">Properties:</span></span>

|<span data-ttu-id="124e4-150">Name</span><span class="sxs-lookup"><span data-stu-id="124e4-150">Name</span></span>| <span data-ttu-id="124e4-151">Typ</span><span class="sxs-lookup"><span data-stu-id="124e4-151">Type</span></span>| <span data-ttu-id="124e4-152">Beschreibung</span><span class="sxs-lookup"><span data-stu-id="124e4-152">Description</span></span>|
|---|---|---|
|`Html`| <span data-ttu-id="124e4-153">Zeichenfolge</span><span class="sxs-lookup"><span data-stu-id="124e4-153">String</span></span>|<span data-ttu-id="124e4-154">Fordert die Rückgabe der Daten im HTML-Format an.</span><span class="sxs-lookup"><span data-stu-id="124e4-154">Requests the data be returned in HTML format.</span></span>|
|`Text`| <span data-ttu-id="124e4-155">Zeichenfolge</span><span class="sxs-lookup"><span data-stu-id="124e4-155">String</span></span>|<span data-ttu-id="124e4-156">Fordert die Rückgabe der Daten im Textformat an.</span><span class="sxs-lookup"><span data-stu-id="124e4-156">Requests the data be returned in text format.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="124e4-157">Anforderungen</span><span class="sxs-lookup"><span data-stu-id="124e4-157">Requirements</span></span>

|<span data-ttu-id="124e4-158">Anforderung</span><span class="sxs-lookup"><span data-stu-id="124e4-158">Requirement</span></span>| <span data-ttu-id="124e4-159">Wert</span><span class="sxs-lookup"><span data-stu-id="124e4-159">Value</span></span>|
|---|---|
|[<span data-ttu-id="124e4-160">Mindestversion des Postfachanforderungssatzes</span><span class="sxs-lookup"><span data-stu-id="124e4-160">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="124e4-161">1.0</span><span class="sxs-lookup"><span data-stu-id="124e4-161">1.0</span></span>|
|[<span data-ttu-id="124e4-162">Zutreffender Outlook-Modus</span><span class="sxs-lookup"><span data-stu-id="124e4-162">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="124e4-163">Verfassen oder Lesen</span><span class="sxs-lookup"><span data-stu-id="124e4-163">Compose or read</span></span>|

---

####  <a name="eventtype-string"></a><span data-ttu-id="124e4-164">EventType :String</span><span class="sxs-lookup"><span data-stu-id="124e4-164">EventType :String</span></span>

<span data-ttu-id="124e4-165">Gibt das einem Ereignishandler zugeordnete Ereignis an.</span><span class="sxs-lookup"><span data-stu-id="124e4-165">Specifies the event associated with an event handler.</span></span>

##### <a name="type"></a><span data-ttu-id="124e4-166">Typ:</span><span class="sxs-lookup"><span data-stu-id="124e4-166">Type:</span></span>

*   <span data-ttu-id="124e4-167">Zeichenfolge</span><span class="sxs-lookup"><span data-stu-id="124e4-167">String</span></span>

##### <a name="properties"></a><span data-ttu-id="124e4-168">Eigenschaften:</span><span class="sxs-lookup"><span data-stu-id="124e4-168">Properties:</span></span>

| <span data-ttu-id="124e4-169">Name</span><span class="sxs-lookup"><span data-stu-id="124e4-169">Name</span></span> | <span data-ttu-id="124e4-170">Typ</span><span class="sxs-lookup"><span data-stu-id="124e4-170">Type</span></span> | <span data-ttu-id="124e4-171">Beschreibung</span><span class="sxs-lookup"><span data-stu-id="124e4-171">Description</span></span> | <span data-ttu-id="124e4-172">Mindestanforderung set</span><span class="sxs-lookup"><span data-stu-id="124e4-172">Minimum requirement set</span></span> |
|---|---|---|---|
|`AppointmentTimeChanged`| <span data-ttu-id="124e4-173">Zeichenfolge</span><span class="sxs-lookup"><span data-stu-id="124e4-173">String</span></span> | <span data-ttu-id="124e4-174">Datum und Uhrzeit der ausgewählten Termins oder einer Datenreihe geändert wurde.</span><span class="sxs-lookup"><span data-stu-id="124e4-174">The date or time of the selected appointment or series has changed.</span></span> | <span data-ttu-id="124e4-175">1.7</span><span class="sxs-lookup"><span data-stu-id="124e4-175">1.7</span></span> |
|`ItemChanged`| <span data-ttu-id="124e4-176">String</span><span class="sxs-lookup"><span data-stu-id="124e4-176">String</span></span> | <span data-ttu-id="124e4-177">Das ausgewählte Element wurde geändert.</span><span class="sxs-lookup"><span data-stu-id="124e4-177">The selected item has changed.</span></span> | <span data-ttu-id="124e4-178">1,5</span><span class="sxs-lookup"><span data-stu-id="124e4-178">1.5</span></span> |
|`RecipientsChanged`| <span data-ttu-id="124e4-179">Zeichenfolge</span><span class="sxs-lookup"><span data-stu-id="124e4-179">String</span></span> | <span data-ttu-id="124e4-180">Die Empfängerliste des ausgewählten Elements oder eines Termins Speicherorts wurde geändert.</span><span class="sxs-lookup"><span data-stu-id="124e4-180">The recipient list of the selected item or appointment location has changed.</span></span> | <span data-ttu-id="124e4-181">1.7</span><span class="sxs-lookup"><span data-stu-id="124e4-181">1.7</span></span> |
|`RecurrenceChanged`| <span data-ttu-id="124e4-182">Zeichenfolge</span><span class="sxs-lookup"><span data-stu-id="124e4-182">String</span></span> | <span data-ttu-id="124e4-183">Das Serienmuster der ausgewählten Datenreihe geändert wurde.</span><span class="sxs-lookup"><span data-stu-id="124e4-183">The recurrence pattern of the selected series has changed.</span></span> | <span data-ttu-id="124e4-184">1.7</span><span class="sxs-lookup"><span data-stu-id="124e4-184">1.7</span></span> |

##### <a name="requirements"></a><span data-ttu-id="124e4-185">Anforderungen</span><span class="sxs-lookup"><span data-stu-id="124e4-185">Requirements</span></span>

|<span data-ttu-id="124e4-186">Anforderung</span><span class="sxs-lookup"><span data-stu-id="124e4-186">Requirement</span></span>| <span data-ttu-id="124e4-187">Wert</span><span class="sxs-lookup"><span data-stu-id="124e4-187">Value</span></span>|
|---|---|
|[<span data-ttu-id="124e4-188">Mindestversion des Postfachanforderungssatzes</span><span class="sxs-lookup"><span data-stu-id="124e4-188">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="124e4-189">1,5</span><span class="sxs-lookup"><span data-stu-id="124e4-189">1.5</span></span> |
|[<span data-ttu-id="124e4-190">Zutreffender Outlook-Modus</span><span class="sxs-lookup"><span data-stu-id="124e4-190">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="124e4-191">Verfassen oder Lesen</span><span class="sxs-lookup"><span data-stu-id="124e4-191">Compose or read</span></span> |

---

####  <a name="sourceproperty-string"></a><span data-ttu-id="124e4-192">SourceProperty :String</span><span class="sxs-lookup"><span data-stu-id="124e4-192">SourceProperty :String</span></span>

<span data-ttu-id="124e4-193">Gibt die Quelle der Daten an, die von der aufgerufenen Methode zurückgegeben werden.</span><span class="sxs-lookup"><span data-stu-id="124e4-193">Specifies the source of the data returned by the invoked method.</span></span>

##### <a name="type"></a><span data-ttu-id="124e4-194">Typ:</span><span class="sxs-lookup"><span data-stu-id="124e4-194">Type:</span></span>

*   <span data-ttu-id="124e4-195">Zeichenfolge</span><span class="sxs-lookup"><span data-stu-id="124e4-195">String</span></span>

##### <a name="properties"></a><span data-ttu-id="124e4-196">Eigenschaften:</span><span class="sxs-lookup"><span data-stu-id="124e4-196">Properties:</span></span>

|<span data-ttu-id="124e4-197">Name</span><span class="sxs-lookup"><span data-stu-id="124e4-197">Name</span></span>| <span data-ttu-id="124e4-198">Typ</span><span class="sxs-lookup"><span data-stu-id="124e4-198">Type</span></span>| <span data-ttu-id="124e4-199">Beschreibung</span><span class="sxs-lookup"><span data-stu-id="124e4-199">Description</span></span>|
|---|---|---|
|`Body`| <span data-ttu-id="124e4-200">Zeichenfolge</span><span class="sxs-lookup"><span data-stu-id="124e4-200">String</span></span>|<span data-ttu-id="124e4-201">Die Quelle der Daten stammt aus dem Textkörper einer Nachricht.</span><span class="sxs-lookup"><span data-stu-id="124e4-201">The source of the data is from the body of a message.</span></span>|
|`Subject`| <span data-ttu-id="124e4-202">Zeichenfolge</span><span class="sxs-lookup"><span data-stu-id="124e4-202">String</span></span>|<span data-ttu-id="124e4-203">Die Quelle der Daten stammt aus dem Betreff einer Nachricht.</span><span class="sxs-lookup"><span data-stu-id="124e4-203">The source of the data is from the subject of a message.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="124e4-204">Anforderungen</span><span class="sxs-lookup"><span data-stu-id="124e4-204">Requirements</span></span>

|<span data-ttu-id="124e4-205">Anforderung</span><span class="sxs-lookup"><span data-stu-id="124e4-205">Requirement</span></span>| <span data-ttu-id="124e4-206">Wert</span><span class="sxs-lookup"><span data-stu-id="124e4-206">Value</span></span>|
|---|---|
|[<span data-ttu-id="124e4-207">Mindestversion des Postfachanforderungssatzes</span><span class="sxs-lookup"><span data-stu-id="124e4-207">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="124e4-208">1.0</span><span class="sxs-lookup"><span data-stu-id="124e4-208">1.0</span></span>|
|[<span data-ttu-id="124e4-209">Zutreffender Outlook-Modus</span><span class="sxs-lookup"><span data-stu-id="124e4-209">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="124e4-210">Verfassen oder Lesen</span><span class="sxs-lookup"><span data-stu-id="124e4-210">Compose or read</span></span>|