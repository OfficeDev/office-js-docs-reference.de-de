 

# <a name="office"></a><span data-ttu-id="c7323-101">Büro</span><span class="sxs-lookup"><span data-stu-id="c7323-101">Office</span></span>

<span data-ttu-id="c7323-p101">Der Office-Namespace stellt gemeinsam genutzte Oberflächen bereit, die von Add-Ins in allen Office-Apps verwendet werden. Diese Auflistung dokumentiert nur die Schnittstellen, die von Outlook-Add-Ins verwendet werden. Eine vollständige Auflistung des Office-Namespaces finden Sie im Thema zur [freigegebenen API](/javascript/api/office).</span><span class="sxs-lookup"><span data-stu-id="c7323-p101">The Office namespace provides shared interfaces that are used by add-ins in all of the Office apps. This listing documents only those interfaces that are used by Outlook add-ins. For a full listing of the Office namespace, see the [Shared API](/javascript/api/office).</span></span>

##### <a name="requirements"></a><span data-ttu-id="c7323-104">Anforderungen</span><span class="sxs-lookup"><span data-stu-id="c7323-104">Requirements</span></span>

|<span data-ttu-id="c7323-105">Anforderung</span><span class="sxs-lookup"><span data-stu-id="c7323-105">Requirement</span></span>| <span data-ttu-id="c7323-106">Wert</span><span class="sxs-lookup"><span data-stu-id="c7323-106">Value</span></span>|
|---|---|
|[<span data-ttu-id="c7323-107">Mindestversion des Postfachanforderungssatzes</span><span class="sxs-lookup"><span data-stu-id="c7323-107">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="c7323-108">1.0</span><span class="sxs-lookup"><span data-stu-id="c7323-108">1.0</span></span>|
|[<span data-ttu-id="c7323-109">Zutreffender Outlook-Modus</span><span class="sxs-lookup"><span data-stu-id="c7323-109">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="c7323-110">Verfassen oder Lesen</span><span class="sxs-lookup"><span data-stu-id="c7323-110">Compose or read</span></span>|

##### <a name="members-and-methods"></a><span data-ttu-id="c7323-111">Elemente und Methoden</span><span class="sxs-lookup"><span data-stu-id="c7323-111">Members and methods</span></span>

| <span data-ttu-id="c7323-112">Element</span><span class="sxs-lookup"><span data-stu-id="c7323-112">Member</span></span> | <span data-ttu-id="c7323-113">Typ</span><span class="sxs-lookup"><span data-stu-id="c7323-113">Type</span></span> |
|--------|------|
| [<span data-ttu-id="c7323-114">AsyncResultStatus</span><span class="sxs-lookup"><span data-stu-id="c7323-114">AsyncResultStatus</span></span>](#asyncresultstatus-string) | <span data-ttu-id="c7323-115">Element</span><span class="sxs-lookup"><span data-stu-id="c7323-115">Member</span></span> |
| [<span data-ttu-id="c7323-116">CoercionType</span><span class="sxs-lookup"><span data-stu-id="c7323-116">CoercionType</span></span>](#coerciontype-string) | <span data-ttu-id="c7323-117">Element</span><span class="sxs-lookup"><span data-stu-id="c7323-117">Member</span></span> |
| [<span data-ttu-id="c7323-118">EventType</span><span class="sxs-lookup"><span data-stu-id="c7323-118">EventType</span></span>](#eventtype-string) | <span data-ttu-id="c7323-119">Element</span><span class="sxs-lookup"><span data-stu-id="c7323-119">Member</span></span> |
| [<span data-ttu-id="c7323-120">SourceProperty</span><span class="sxs-lookup"><span data-stu-id="c7323-120">SourceProperty</span></span>](#sourceproperty-string) | <span data-ttu-id="c7323-121">Element</span><span class="sxs-lookup"><span data-stu-id="c7323-121">Member</span></span> |

### <a name="namespaces"></a><span data-ttu-id="c7323-122">Namespaces</span><span class="sxs-lookup"><span data-stu-id="c7323-122">Namespaces</span></span>

<span data-ttu-id="c7323-123">[Kontextmenü](office.context.md): gemeinsam genutzten Oberflächen aus dem Office-Add-ins-API-Kontext-Namespace für die Verwendung in der Outlook-add-in-API bietet.</span><span class="sxs-lookup"><span data-stu-id="c7323-123">[context](office.context.md): Provides shared interfaces from the Office Add-ins API's context namespace for use in the Outlook add-in API.</span></span>

<span data-ttu-id="c7323-124">[MailboxEnums](/javascript/api/outlook_1_7/office.mailboxenums.attachmenttype): Enthält die Enumerationen ItemType, EntityType, AttachmentType, RecipientType, ResponseType und ItemNotificationMessageType.</span><span class="sxs-lookup"><span data-stu-id="c7323-124">[MailboxEnums](/javascript/api/outlook_1_7/office.mailboxenums.attachmenttype): Includes the ItemType, EntityType, AttachmentType, RecipientType, ResponseType, and ItemNotificationMessageType enumerations.</span></span>

### <a name="members"></a><span data-ttu-id="c7323-125">Elemente</span><span class="sxs-lookup"><span data-stu-id="c7323-125">Members</span></span>

####  <a name="asyncresultstatus-string"></a><span data-ttu-id="c7323-126">AsyncResultStatus :String</span><span class="sxs-lookup"><span data-stu-id="c7323-126">AsyncResultStatus :String</span></span>

<span data-ttu-id="c7323-127">Gibt das Ergebnis eines asynchronen Aufrufs an.</span><span class="sxs-lookup"><span data-stu-id="c7323-127">Specifies the result of an asynchronous call.</span></span>

##### <a name="type"></a><span data-ttu-id="c7323-128">Typ:</span><span class="sxs-lookup"><span data-stu-id="c7323-128">Type:</span></span>

*   <span data-ttu-id="c7323-129">Zeichenfolge</span><span class="sxs-lookup"><span data-stu-id="c7323-129">String</span></span>

##### <a name="properties"></a><span data-ttu-id="c7323-130">Eigenschaften:</span><span class="sxs-lookup"><span data-stu-id="c7323-130">Properties:</span></span>

|<span data-ttu-id="c7323-131">Name</span><span class="sxs-lookup"><span data-stu-id="c7323-131">Name</span></span>| <span data-ttu-id="c7323-132">Typ</span><span class="sxs-lookup"><span data-stu-id="c7323-132">Type</span></span>| <span data-ttu-id="c7323-133">Beschreibung</span><span class="sxs-lookup"><span data-stu-id="c7323-133">Description</span></span>|
|---|---|---|
|`Succeeded`| <span data-ttu-id="c7323-134">Zeichenfolge</span><span class="sxs-lookup"><span data-stu-id="c7323-134">String</span></span>|<span data-ttu-id="c7323-135">Der Aufruf war erfolgreich.</span><span class="sxs-lookup"><span data-stu-id="c7323-135">The call succeeded.</span></span>|
|`Failed`| <span data-ttu-id="c7323-136">Zeichenfolge</span><span class="sxs-lookup"><span data-stu-id="c7323-136">String</span></span>|<span data-ttu-id="c7323-137">Der Aufruf ist fehlerhaft.</span><span class="sxs-lookup"><span data-stu-id="c7323-137">The call failed.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="c7323-138">Anforderungen</span><span class="sxs-lookup"><span data-stu-id="c7323-138">Requirements</span></span>

|<span data-ttu-id="c7323-139">Anforderung</span><span class="sxs-lookup"><span data-stu-id="c7323-139">Requirement</span></span>| <span data-ttu-id="c7323-140">Wert</span><span class="sxs-lookup"><span data-stu-id="c7323-140">Value</span></span>|
|---|---|
|[<span data-ttu-id="c7323-141">Mindestversion des Postfachanforderungssatzes</span><span class="sxs-lookup"><span data-stu-id="c7323-141">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="c7323-142">1.0</span><span class="sxs-lookup"><span data-stu-id="c7323-142">1.0</span></span>|
|[<span data-ttu-id="c7323-143">Zutreffender Outlook-Modus</span><span class="sxs-lookup"><span data-stu-id="c7323-143">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="c7323-144">Verfassen oder Lesen</span><span class="sxs-lookup"><span data-stu-id="c7323-144">Compose or read</span></span>|

---

####  <a name="coerciontype-string"></a><span data-ttu-id="c7323-145">CoercionType :String</span><span class="sxs-lookup"><span data-stu-id="c7323-145">CoercionType :String</span></span>

<span data-ttu-id="c7323-146">Gibt an, wie Daten umgewandelt werden sollen, die von der aufgerufenen Methode zurückgegeben oder festgelegt wurden.</span><span class="sxs-lookup"><span data-stu-id="c7323-146">Specifies how to coerce data returned or set by the invoked method.</span></span>

##### <a name="type"></a><span data-ttu-id="c7323-147">Typ:</span><span class="sxs-lookup"><span data-stu-id="c7323-147">Type:</span></span>

*   <span data-ttu-id="c7323-148">Zeichenfolge</span><span class="sxs-lookup"><span data-stu-id="c7323-148">String</span></span>

##### <a name="properties"></a><span data-ttu-id="c7323-149">Eigenschaften:</span><span class="sxs-lookup"><span data-stu-id="c7323-149">Properties:</span></span>

|<span data-ttu-id="c7323-150">Name</span><span class="sxs-lookup"><span data-stu-id="c7323-150">Name</span></span>| <span data-ttu-id="c7323-151">Typ</span><span class="sxs-lookup"><span data-stu-id="c7323-151">Type</span></span>| <span data-ttu-id="c7323-152">Beschreibung</span><span class="sxs-lookup"><span data-stu-id="c7323-152">Description</span></span>|
|---|---|---|
|`Html`| <span data-ttu-id="c7323-153">Zeichenfolge</span><span class="sxs-lookup"><span data-stu-id="c7323-153">String</span></span>|<span data-ttu-id="c7323-154">Fordert die Rückgabe der Daten im HTML-Format an.</span><span class="sxs-lookup"><span data-stu-id="c7323-154">Requests the data be returned in HTML format.</span></span>|
|`Text`| <span data-ttu-id="c7323-155">Zeichenfolge</span><span class="sxs-lookup"><span data-stu-id="c7323-155">String</span></span>|<span data-ttu-id="c7323-156">Fordert die Rückgabe der Daten im Textformat an.</span><span class="sxs-lookup"><span data-stu-id="c7323-156">Requests the data be returned in text format.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="c7323-157">Anforderungen</span><span class="sxs-lookup"><span data-stu-id="c7323-157">Requirements</span></span>

|<span data-ttu-id="c7323-158">Anforderung</span><span class="sxs-lookup"><span data-stu-id="c7323-158">Requirement</span></span>| <span data-ttu-id="c7323-159">Wert</span><span class="sxs-lookup"><span data-stu-id="c7323-159">Value</span></span>|
|---|---|
|[<span data-ttu-id="c7323-160">Mindestversion des Postfachanforderungssatzes</span><span class="sxs-lookup"><span data-stu-id="c7323-160">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="c7323-161">1.0</span><span class="sxs-lookup"><span data-stu-id="c7323-161">1.0</span></span>|
|[<span data-ttu-id="c7323-162">Zutreffender Outlook-Modus</span><span class="sxs-lookup"><span data-stu-id="c7323-162">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="c7323-163">Verfassen oder Lesen</span><span class="sxs-lookup"><span data-stu-id="c7323-163">Compose or read</span></span>|

---

####  <a name="eventtype-string"></a><span data-ttu-id="c7323-164">EventType :String</span><span class="sxs-lookup"><span data-stu-id="c7323-164">EventType :String</span></span>

<span data-ttu-id="c7323-165">Gibt das einem Ereignishandler zugeordnete Ereignis an.</span><span class="sxs-lookup"><span data-stu-id="c7323-165">Specifies the event associated with an event handler.</span></span>

##### <a name="type"></a><span data-ttu-id="c7323-166">Typ:</span><span class="sxs-lookup"><span data-stu-id="c7323-166">Type:</span></span>

*   <span data-ttu-id="c7323-167">Zeichenfolge</span><span class="sxs-lookup"><span data-stu-id="c7323-167">String</span></span>

##### <a name="properties"></a><span data-ttu-id="c7323-168">Eigenschaften:</span><span class="sxs-lookup"><span data-stu-id="c7323-168">Properties:</span></span>

| <span data-ttu-id="c7323-169">Name</span><span class="sxs-lookup"><span data-stu-id="c7323-169">Name</span></span> | <span data-ttu-id="c7323-170">Typ</span><span class="sxs-lookup"><span data-stu-id="c7323-170">Type</span></span> | <span data-ttu-id="c7323-171">Beschreibung</span><span class="sxs-lookup"><span data-stu-id="c7323-171">Description</span></span> | <span data-ttu-id="c7323-172">Mindestanforderung set</span><span class="sxs-lookup"><span data-stu-id="c7323-172">Minimum requirement set</span></span> |
|---|---|---|---|
|`AppointmentTimeChanged`| <span data-ttu-id="c7323-173">Zeichenfolge</span><span class="sxs-lookup"><span data-stu-id="c7323-173">String</span></span> | <span data-ttu-id="c7323-174">Datum und Uhrzeit der ausgewählten Termins oder einer Datenreihe geändert wurde.</span><span class="sxs-lookup"><span data-stu-id="c7323-174">The date or time of the selected appointment or series has changed.</span></span> | <span data-ttu-id="c7323-175">1.7</span><span class="sxs-lookup"><span data-stu-id="c7323-175">1.7</span></span> |
|`ItemChanged`| <span data-ttu-id="c7323-176">String</span><span class="sxs-lookup"><span data-stu-id="c7323-176">String</span></span> | <span data-ttu-id="c7323-177">Das ausgewählte Element wurde geändert.</span><span class="sxs-lookup"><span data-stu-id="c7323-177">The selected item has changed.</span></span> | <span data-ttu-id="c7323-178">1,5</span><span class="sxs-lookup"><span data-stu-id="c7323-178">1.5</span></span> |
|`RecipientsChanged`| <span data-ttu-id="c7323-179">Zeichenfolge</span><span class="sxs-lookup"><span data-stu-id="c7323-179">String</span></span> | <span data-ttu-id="c7323-180">Die Empfängerliste des ausgewählten Elements oder eines Termins Speicherorts wurde geändert.</span><span class="sxs-lookup"><span data-stu-id="c7323-180">The recipient list of the selected item or appointment location has changed.</span></span> | <span data-ttu-id="c7323-181">1.7</span><span class="sxs-lookup"><span data-stu-id="c7323-181">1.7</span></span> |
|`RecurrenceChanged`| <span data-ttu-id="c7323-182">Zeichenfolge</span><span class="sxs-lookup"><span data-stu-id="c7323-182">String</span></span> | <span data-ttu-id="c7323-183">Das Serienmuster der ausgewählten Datenreihe geändert wurde.</span><span class="sxs-lookup"><span data-stu-id="c7323-183">The recurrence pattern of the selected series has changed.</span></span> | <span data-ttu-id="c7323-184">1.7</span><span class="sxs-lookup"><span data-stu-id="c7323-184">1.7</span></span> |

##### <a name="requirements"></a><span data-ttu-id="c7323-185">Anforderungen</span><span class="sxs-lookup"><span data-stu-id="c7323-185">Requirements</span></span>

|<span data-ttu-id="c7323-186">Anforderung</span><span class="sxs-lookup"><span data-stu-id="c7323-186">Requirement</span></span>| <span data-ttu-id="c7323-187">Wert</span><span class="sxs-lookup"><span data-stu-id="c7323-187">Value</span></span>|
|---|---|
|[<span data-ttu-id="c7323-188">Mindestversion des Postfachanforderungssatzes</span><span class="sxs-lookup"><span data-stu-id="c7323-188">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="c7323-189">1,5</span><span class="sxs-lookup"><span data-stu-id="c7323-189">1.5</span></span> |
|[<span data-ttu-id="c7323-190">Zutreffender Outlook-Modus</span><span class="sxs-lookup"><span data-stu-id="c7323-190">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="c7323-191">Verfassen oder Lesen</span><span class="sxs-lookup"><span data-stu-id="c7323-191">Compose or read</span></span> |

---

####  <a name="sourceproperty-string"></a><span data-ttu-id="c7323-192">SourceProperty :String</span><span class="sxs-lookup"><span data-stu-id="c7323-192">SourceProperty :String</span></span>

<span data-ttu-id="c7323-193">Gibt die Quelle der Daten an, die von der aufgerufenen Methode zurückgegeben werden.</span><span class="sxs-lookup"><span data-stu-id="c7323-193">Specifies the source of the data returned by the invoked method.</span></span>

##### <a name="type"></a><span data-ttu-id="c7323-194">Typ:</span><span class="sxs-lookup"><span data-stu-id="c7323-194">Type:</span></span>

*   <span data-ttu-id="c7323-195">Zeichenfolge</span><span class="sxs-lookup"><span data-stu-id="c7323-195">String</span></span>

##### <a name="properties"></a><span data-ttu-id="c7323-196">Eigenschaften:</span><span class="sxs-lookup"><span data-stu-id="c7323-196">Properties:</span></span>

|<span data-ttu-id="c7323-197">Name</span><span class="sxs-lookup"><span data-stu-id="c7323-197">Name</span></span>| <span data-ttu-id="c7323-198">Typ</span><span class="sxs-lookup"><span data-stu-id="c7323-198">Type</span></span>| <span data-ttu-id="c7323-199">Beschreibung</span><span class="sxs-lookup"><span data-stu-id="c7323-199">Description</span></span>|
|---|---|---|
|`Body`| <span data-ttu-id="c7323-200">Zeichenfolge</span><span class="sxs-lookup"><span data-stu-id="c7323-200">String</span></span>|<span data-ttu-id="c7323-201">Die Quelle der Daten stammt aus dem Textkörper einer Nachricht.</span><span class="sxs-lookup"><span data-stu-id="c7323-201">The source of the data is from the body of a message.</span></span>|
|`Subject`| <span data-ttu-id="c7323-202">Zeichenfolge</span><span class="sxs-lookup"><span data-stu-id="c7323-202">String</span></span>|<span data-ttu-id="c7323-203">Die Quelle der Daten stammt aus dem Betreff einer Nachricht.</span><span class="sxs-lookup"><span data-stu-id="c7323-203">The source of the data is from the subject of a message.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="c7323-204">Anforderungen</span><span class="sxs-lookup"><span data-stu-id="c7323-204">Requirements</span></span>

|<span data-ttu-id="c7323-205">Anforderung</span><span class="sxs-lookup"><span data-stu-id="c7323-205">Requirement</span></span>| <span data-ttu-id="c7323-206">Wert</span><span class="sxs-lookup"><span data-stu-id="c7323-206">Value</span></span>|
|---|---|
|[<span data-ttu-id="c7323-207">Mindestversion des Postfachanforderungssatzes</span><span class="sxs-lookup"><span data-stu-id="c7323-207">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="c7323-208">1.0</span><span class="sxs-lookup"><span data-stu-id="c7323-208">1.0</span></span>|
|[<span data-ttu-id="c7323-209">Zutreffender Outlook-Modus</span><span class="sxs-lookup"><span data-stu-id="c7323-209">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="c7323-210">Verfassen oder Lesen</span><span class="sxs-lookup"><span data-stu-id="c7323-210">Compose or read</span></span>|