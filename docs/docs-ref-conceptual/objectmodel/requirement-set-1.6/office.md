 

# <a name="office"></a><span data-ttu-id="18f32-101">Büro</span><span class="sxs-lookup"><span data-stu-id="18f32-101">Office</span></span>

<span data-ttu-id="18f32-p101">Der Office-Namespace stellt gemeinsam genutzte Oberflächen bereit, die von Add-Ins in allen Office-Apps verwendet werden. Diese Auflistung dokumentiert nur die Schnittstellen, die von Outlook-Add-Ins verwendet werden. Eine vollständige Auflistung des Office-Namespaces finden Sie im Thema zur [freigegebenen API](/javascript/api/office).</span><span class="sxs-lookup"><span data-stu-id="18f32-p101">The Office namespace provides shared interfaces that are used by add-ins in all of the Office apps. This listing documents only those interfaces that are used by Outlook add-ins. For a full listing of the Office namespace, see the [Shared API](/javascript/api/office).</span></span>

##### <a name="requirements"></a><span data-ttu-id="18f32-104">Anforderungen</span><span class="sxs-lookup"><span data-stu-id="18f32-104">Requirements</span></span>

|<span data-ttu-id="18f32-105">Anforderung</span><span class="sxs-lookup"><span data-stu-id="18f32-105">Requirement</span></span>| <span data-ttu-id="18f32-106">Wert</span><span class="sxs-lookup"><span data-stu-id="18f32-106">Value</span></span>|
|---|---|
|[<span data-ttu-id="18f32-107">Mindestversion des Postfachanforderungssatzes</span><span class="sxs-lookup"><span data-stu-id="18f32-107">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="18f32-108">1.0</span><span class="sxs-lookup"><span data-stu-id="18f32-108">1.0</span></span>|
|[<span data-ttu-id="18f32-109">Zutreffender Outlook-Modus</span><span class="sxs-lookup"><span data-stu-id="18f32-109">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="18f32-110">Verfassen oder Lesen</span><span class="sxs-lookup"><span data-stu-id="18f32-110">Compose or read</span></span>|

##### <a name="members-and-methods"></a><span data-ttu-id="18f32-111">Elemente und Methoden</span><span class="sxs-lookup"><span data-stu-id="18f32-111">Members and methods</span></span>

| <span data-ttu-id="18f32-112">Element</span><span class="sxs-lookup"><span data-stu-id="18f32-112">Member</span></span> | <span data-ttu-id="18f32-113">Typ</span><span class="sxs-lookup"><span data-stu-id="18f32-113">Type</span></span> |
|--------|------|
| [<span data-ttu-id="18f32-114">AsyncResultStatus</span><span class="sxs-lookup"><span data-stu-id="18f32-114">AsyncResultStatus</span></span>](#asyncresultstatus-string) | <span data-ttu-id="18f32-115">Element</span><span class="sxs-lookup"><span data-stu-id="18f32-115">Member</span></span> |
| [<span data-ttu-id="18f32-116">CoercionType</span><span class="sxs-lookup"><span data-stu-id="18f32-116">CoercionType</span></span>](#coerciontype-string) | <span data-ttu-id="18f32-117">Element</span><span class="sxs-lookup"><span data-stu-id="18f32-117">Member</span></span> |
| [<span data-ttu-id="18f32-118">EventType</span><span class="sxs-lookup"><span data-stu-id="18f32-118">EventType</span></span>](#eventtype-string) | <span data-ttu-id="18f32-119">Element</span><span class="sxs-lookup"><span data-stu-id="18f32-119">Member</span></span> |
| [<span data-ttu-id="18f32-120">SourceProperty</span><span class="sxs-lookup"><span data-stu-id="18f32-120">SourceProperty</span></span>](#sourceproperty-string) | <span data-ttu-id="18f32-121">Element</span><span class="sxs-lookup"><span data-stu-id="18f32-121">Member</span></span> |

### <a name="namespaces"></a><span data-ttu-id="18f32-122">Namespaces</span><span class="sxs-lookup"><span data-stu-id="18f32-122">Namespaces</span></span>

<span data-ttu-id="18f32-123">[Kontextmenü](office.context.md): gemeinsam genutzten Oberflächen aus dem Office-Add-ins-API-Kontext-Namespace für die Verwendung in der Outlook-add-in-API bietet.</span><span class="sxs-lookup"><span data-stu-id="18f32-123">[context](office.context.md): Provides shared interfaces from the Office Add-ins API's context namespace for use in the Outlook add-in API.</span></span>

<span data-ttu-id="18f32-124">[MailboxEnums](/javascript/api/outlook/office.mailboxenums.attachmenttype): Enthält die Enumerationen ItemType, EntityType, AttachmentType, RecipientType, ResponseType und ItemNotificationMessageType.</span><span class="sxs-lookup"><span data-stu-id="18f32-124">[MailboxEnums](/javascript/api/outlook/office.mailboxenums.attachmenttype): Includes the ItemType, EntityType, AttachmentType, RecipientType, ResponseType, and ItemNotificationMessageType enumerations.</span></span>

### <a name="members"></a><span data-ttu-id="18f32-125">Elemente</span><span class="sxs-lookup"><span data-stu-id="18f32-125">Members</span></span>

####  <a name="asyncresultstatus-string"></a><span data-ttu-id="18f32-126">AsyncResultStatus :String</span><span class="sxs-lookup"><span data-stu-id="18f32-126">AsyncResultStatus :String</span></span>

<span data-ttu-id="18f32-127">Gibt das Ergebnis eines asynchronen Aufrufs an.</span><span class="sxs-lookup"><span data-stu-id="18f32-127">Specifies the result of an asynchronous call.</span></span>

##### <a name="type"></a><span data-ttu-id="18f32-128">Typ:</span><span class="sxs-lookup"><span data-stu-id="18f32-128">Type:</span></span>

*   <span data-ttu-id="18f32-129">Zeichenfolge</span><span class="sxs-lookup"><span data-stu-id="18f32-129">String</span></span>

##### <a name="properties"></a><span data-ttu-id="18f32-130">Eigenschaften:</span><span class="sxs-lookup"><span data-stu-id="18f32-130">Properties:</span></span>

|<span data-ttu-id="18f32-131">Name</span><span class="sxs-lookup"><span data-stu-id="18f32-131">Name</span></span>| <span data-ttu-id="18f32-132">Typ</span><span class="sxs-lookup"><span data-stu-id="18f32-132">Type</span></span>| <span data-ttu-id="18f32-133">Beschreibung</span><span class="sxs-lookup"><span data-stu-id="18f32-133">Description</span></span>|
|---|---|---|
|`Succeeded`| <span data-ttu-id="18f32-134">Zeichenfolge</span><span class="sxs-lookup"><span data-stu-id="18f32-134">String</span></span>|<span data-ttu-id="18f32-135">Der Aufruf war erfolgreich.</span><span class="sxs-lookup"><span data-stu-id="18f32-135">The call succeeded.</span></span>|
|`Failed`| <span data-ttu-id="18f32-136">Zeichenfolge</span><span class="sxs-lookup"><span data-stu-id="18f32-136">String</span></span>|<span data-ttu-id="18f32-137">Der Aufruf ist fehlerhaft.</span><span class="sxs-lookup"><span data-stu-id="18f32-137">The call failed.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="18f32-138">Anforderungen</span><span class="sxs-lookup"><span data-stu-id="18f32-138">Requirements</span></span>

|<span data-ttu-id="18f32-139">Anforderung</span><span class="sxs-lookup"><span data-stu-id="18f32-139">Requirement</span></span>| <span data-ttu-id="18f32-140">Wert</span><span class="sxs-lookup"><span data-stu-id="18f32-140">Value</span></span>|
|---|---|
|[<span data-ttu-id="18f32-141">Mindestversion des Postfachanforderungssatzes</span><span class="sxs-lookup"><span data-stu-id="18f32-141">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="18f32-142">1.0</span><span class="sxs-lookup"><span data-stu-id="18f32-142">1.0</span></span>|
|[<span data-ttu-id="18f32-143">Zutreffender Outlook-Modus</span><span class="sxs-lookup"><span data-stu-id="18f32-143">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="18f32-144">Verfassen oder Lesen</span><span class="sxs-lookup"><span data-stu-id="18f32-144">Compose or read</span></span>|

---

####  <a name="coerciontype-string"></a><span data-ttu-id="18f32-145">CoercionType :String</span><span class="sxs-lookup"><span data-stu-id="18f32-145">CoercionType :String</span></span>

<span data-ttu-id="18f32-146">Gibt an, wie Daten umgewandelt werden sollen, die von der aufgerufenen Methode zurückgegeben oder festgelegt wurden.</span><span class="sxs-lookup"><span data-stu-id="18f32-146">Specifies how to coerce data returned or set by the invoked method.</span></span>

##### <a name="type"></a><span data-ttu-id="18f32-147">Typ:</span><span class="sxs-lookup"><span data-stu-id="18f32-147">Type:</span></span>

*   <span data-ttu-id="18f32-148">Zeichenfolge</span><span class="sxs-lookup"><span data-stu-id="18f32-148">String</span></span>

##### <a name="properties"></a><span data-ttu-id="18f32-149">Eigenschaften:</span><span class="sxs-lookup"><span data-stu-id="18f32-149">Properties:</span></span>

|<span data-ttu-id="18f32-150">Name</span><span class="sxs-lookup"><span data-stu-id="18f32-150">Name</span></span>| <span data-ttu-id="18f32-151">Typ</span><span class="sxs-lookup"><span data-stu-id="18f32-151">Type</span></span>| <span data-ttu-id="18f32-152">Beschreibung</span><span class="sxs-lookup"><span data-stu-id="18f32-152">Description</span></span>|
|---|---|---|
|`Html`| <span data-ttu-id="18f32-153">Zeichenfolge</span><span class="sxs-lookup"><span data-stu-id="18f32-153">String</span></span>|<span data-ttu-id="18f32-154">Fordert die Rückgabe der Daten im HTML-Format an.</span><span class="sxs-lookup"><span data-stu-id="18f32-154">Requests the data be returned in HTML format.</span></span>|
|`Text`| <span data-ttu-id="18f32-155">Zeichenfolge</span><span class="sxs-lookup"><span data-stu-id="18f32-155">String</span></span>|<span data-ttu-id="18f32-156">Fordert die Rückgabe der Daten im Textformat an.</span><span class="sxs-lookup"><span data-stu-id="18f32-156">Requests the data be returned in text format.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="18f32-157">Anforderungen</span><span class="sxs-lookup"><span data-stu-id="18f32-157">Requirements</span></span>

|<span data-ttu-id="18f32-158">Anforderung</span><span class="sxs-lookup"><span data-stu-id="18f32-158">Requirement</span></span>| <span data-ttu-id="18f32-159">Wert</span><span class="sxs-lookup"><span data-stu-id="18f32-159">Value</span></span>|
|---|---|
|[<span data-ttu-id="18f32-160">Mindestversion des Postfachanforderungssatzes</span><span class="sxs-lookup"><span data-stu-id="18f32-160">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="18f32-161">1.0</span><span class="sxs-lookup"><span data-stu-id="18f32-161">1.0</span></span>|
|[<span data-ttu-id="18f32-162">Zutreffender Outlook-Modus</span><span class="sxs-lookup"><span data-stu-id="18f32-162">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="18f32-163">Verfassen oder Lesen</span><span class="sxs-lookup"><span data-stu-id="18f32-163">Compose or read</span></span>|

---

####  <a name="eventtype-string"></a><span data-ttu-id="18f32-164">EventType :String</span><span class="sxs-lookup"><span data-stu-id="18f32-164">EventType :String</span></span>

<span data-ttu-id="18f32-165">Gibt das einem Ereignishandler zugeordnete Ereignis an.</span><span class="sxs-lookup"><span data-stu-id="18f32-165">Specifies the event associated with an event handler.</span></span>

##### <a name="type"></a><span data-ttu-id="18f32-166">Typ:</span><span class="sxs-lookup"><span data-stu-id="18f32-166">Type:</span></span>

*   <span data-ttu-id="18f32-167">Zeichenfolge</span><span class="sxs-lookup"><span data-stu-id="18f32-167">String</span></span>

##### <a name="properties"></a><span data-ttu-id="18f32-168">Eigenschaften:</span><span class="sxs-lookup"><span data-stu-id="18f32-168">Properties:</span></span>

| <span data-ttu-id="18f32-169">Name</span><span class="sxs-lookup"><span data-stu-id="18f32-169">Name</span></span> | <span data-ttu-id="18f32-170">Typ</span><span class="sxs-lookup"><span data-stu-id="18f32-170">Type</span></span> | <span data-ttu-id="18f32-171">Beschreibung</span><span class="sxs-lookup"><span data-stu-id="18f32-171">Description</span></span> |
|---|---|---|
|`ItemChanged`| <span data-ttu-id="18f32-172">String</span><span class="sxs-lookup"><span data-stu-id="18f32-172">String</span></span> | <span data-ttu-id="18f32-173">Das ausgewählte Element wurde geändert.</span><span class="sxs-lookup"><span data-stu-id="18f32-173">The selected item has changed.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="18f32-174">Anforderungen</span><span class="sxs-lookup"><span data-stu-id="18f32-174">Requirements</span></span>

|<span data-ttu-id="18f32-175">Anforderung</span><span class="sxs-lookup"><span data-stu-id="18f32-175">Requirement</span></span>| <span data-ttu-id="18f32-176">Wert</span><span class="sxs-lookup"><span data-stu-id="18f32-176">Value</span></span>|
|---|---|
|[<span data-ttu-id="18f32-177">Mindestversion des Postfachanforderungssatzes</span><span class="sxs-lookup"><span data-stu-id="18f32-177">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="18f32-178">1,5</span><span class="sxs-lookup"><span data-stu-id="18f32-178">1.5</span></span> |
|[<span data-ttu-id="18f32-179">Zutreffender Outlook-Modus</span><span class="sxs-lookup"><span data-stu-id="18f32-179">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="18f32-180">Verfassen oder Lesen</span><span class="sxs-lookup"><span data-stu-id="18f32-180">Compose or read</span></span> |

---

####  <a name="sourceproperty-string"></a><span data-ttu-id="18f32-181">SourceProperty :String</span><span class="sxs-lookup"><span data-stu-id="18f32-181">SourceProperty :String</span></span>

<span data-ttu-id="18f32-182">Gibt die Quelle der Daten an, die von der aufgerufenen Methode zurückgegeben werden.</span><span class="sxs-lookup"><span data-stu-id="18f32-182">Specifies the source of the data returned by the invoked method.</span></span>

##### <a name="type"></a><span data-ttu-id="18f32-183">Typ:</span><span class="sxs-lookup"><span data-stu-id="18f32-183">Type:</span></span>

*   <span data-ttu-id="18f32-184">Zeichenfolge</span><span class="sxs-lookup"><span data-stu-id="18f32-184">String</span></span>

##### <a name="properties"></a><span data-ttu-id="18f32-185">Eigenschaften:</span><span class="sxs-lookup"><span data-stu-id="18f32-185">Properties:</span></span>

|<span data-ttu-id="18f32-186">Name</span><span class="sxs-lookup"><span data-stu-id="18f32-186">Name</span></span>| <span data-ttu-id="18f32-187">Typ</span><span class="sxs-lookup"><span data-stu-id="18f32-187">Type</span></span>| <span data-ttu-id="18f32-188">Beschreibung</span><span class="sxs-lookup"><span data-stu-id="18f32-188">Description</span></span>|
|---|---|---|
|`Body`| <span data-ttu-id="18f32-189">Zeichenfolge</span><span class="sxs-lookup"><span data-stu-id="18f32-189">String</span></span>|<span data-ttu-id="18f32-190">Die Quelle der Daten stammt aus dem Textkörper einer Nachricht.</span><span class="sxs-lookup"><span data-stu-id="18f32-190">The source of the data is from the body of a message.</span></span>|
|`Subject`| <span data-ttu-id="18f32-191">Zeichenfolge</span><span class="sxs-lookup"><span data-stu-id="18f32-191">String</span></span>|<span data-ttu-id="18f32-192">Die Quelle der Daten stammt aus dem Betreff einer Nachricht.</span><span class="sxs-lookup"><span data-stu-id="18f32-192">The source of the data is from the subject of a message.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="18f32-193">Anforderungen</span><span class="sxs-lookup"><span data-stu-id="18f32-193">Requirements</span></span>

|<span data-ttu-id="18f32-194">Anforderung</span><span class="sxs-lookup"><span data-stu-id="18f32-194">Requirement</span></span>| <span data-ttu-id="18f32-195">Wert</span><span class="sxs-lookup"><span data-stu-id="18f32-195">Value</span></span>|
|---|---|
|[<span data-ttu-id="18f32-196">Mindestversion des Postfachanforderungssatzes</span><span class="sxs-lookup"><span data-stu-id="18f32-196">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="18f32-197">1.0</span><span class="sxs-lookup"><span data-stu-id="18f32-197">1.0</span></span>|
|[<span data-ttu-id="18f32-198">Zutreffender Outlook-Modus</span><span class="sxs-lookup"><span data-stu-id="18f32-198">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="18f32-199">Verfassen oder Lesen</span><span class="sxs-lookup"><span data-stu-id="18f32-199">Compose or read</span></span>|