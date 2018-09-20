# <a name="office"></a><span data-ttu-id="4fa79-101">Büro</span><span class="sxs-lookup"><span data-stu-id="4fa79-101">Office</span></span>

<span data-ttu-id="4fa79-p101">Der Office-Namespace stellt gemeinsam genutzte Oberflächen bereit, die von Add-Ins in allen Office-Apps verwendet werden. Diese Auflistung dokumentiert nur die Schnittstellen, die von Outlook-Add-Ins verwendet werden. Eine vollständige Auflistung des Office-Namespaces finden Sie im Thema zur [freigegebenen API](/javascript/api/office).</span><span class="sxs-lookup"><span data-stu-id="4fa79-p101">The Office namespace provides shared interfaces that are used by add-ins in all of the Office apps. This listing documents only those interfaces that are used by Outlook add-ins. For a full listing of the Office namespace, see the [Shared API](/javascript/api/office).</span></span>

##### <a name="requirements"></a><span data-ttu-id="4fa79-104">Anforderungen</span><span class="sxs-lookup"><span data-stu-id="4fa79-104">Requirements</span></span>

|<span data-ttu-id="4fa79-105">Anforderung</span><span class="sxs-lookup"><span data-stu-id="4fa79-105">Requirement</span></span>| <span data-ttu-id="4fa79-106">Wert</span><span class="sxs-lookup"><span data-stu-id="4fa79-106">Value</span></span>|
|---|---|
|[<span data-ttu-id="4fa79-107">Mindestversion des Postfachanforderungssatzes</span><span class="sxs-lookup"><span data-stu-id="4fa79-107">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="4fa79-108">1.0</span><span class="sxs-lookup"><span data-stu-id="4fa79-108">1.0</span></span>|
|[<span data-ttu-id="4fa79-109">Zutreffender Outlook-Modus</span><span class="sxs-lookup"><span data-stu-id="4fa79-109">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="4fa79-110">Verfassen oder Lesen</span><span class="sxs-lookup"><span data-stu-id="4fa79-110">Compose or read</span></span>|

##### <a name="members-and-methods"></a><span data-ttu-id="4fa79-111">Elemente und Methoden</span><span class="sxs-lookup"><span data-stu-id="4fa79-111">Members and methods</span></span>

| <span data-ttu-id="4fa79-112">Element</span><span class="sxs-lookup"><span data-stu-id="4fa79-112">Member</span></span> | <span data-ttu-id="4fa79-113">Typ</span><span class="sxs-lookup"><span data-stu-id="4fa79-113">Type</span></span> |
|--------|------|
| [<span data-ttu-id="4fa79-114">AsyncResultStatus</span><span class="sxs-lookup"><span data-stu-id="4fa79-114">AsyncResultStatus</span></span>](#asyncresultstatus-string) | <span data-ttu-id="4fa79-115">Element</span><span class="sxs-lookup"><span data-stu-id="4fa79-115">Member</span></span> |
| [<span data-ttu-id="4fa79-116">CoercionType</span><span class="sxs-lookup"><span data-stu-id="4fa79-116">CoercionType</span></span>](#coerciontype-string) | <span data-ttu-id="4fa79-117">Element</span><span class="sxs-lookup"><span data-stu-id="4fa79-117">Member</span></span> |
| [<span data-ttu-id="4fa79-118">EventType</span><span class="sxs-lookup"><span data-stu-id="4fa79-118">EventType</span></span>](#eventtype-string) | <span data-ttu-id="4fa79-119">Element</span><span class="sxs-lookup"><span data-stu-id="4fa79-119">Member</span></span> |
| [<span data-ttu-id="4fa79-120">SourceProperty</span><span class="sxs-lookup"><span data-stu-id="4fa79-120">SourceProperty</span></span>](#sourceproperty-string) | <span data-ttu-id="4fa79-121">Element</span><span class="sxs-lookup"><span data-stu-id="4fa79-121">Member</span></span> |

### <a name="namespaces"></a><span data-ttu-id="4fa79-122">Namespaces</span><span class="sxs-lookup"><span data-stu-id="4fa79-122">Namespaces</span></span>

<span data-ttu-id="4fa79-123">[Kontextmenü](office.context.md): gemeinsam genutzten Oberflächen aus dem Office-Add-ins-API-Kontext-Namespace für die Verwendung in der Outlook-add-in-API bietet.</span><span class="sxs-lookup"><span data-stu-id="4fa79-123">[context](office.context.md): Provides shared interfaces from the Office Add-ins API's context namespace for use in the Outlook add-in API.</span></span>

<span data-ttu-id="4fa79-124">[MailboxEnums](/javascript/api/outlook/office.mailboxenums.attachmenttype): Enthält die Enumerationen ItemType, EntityType, AttachmentType, RecipientType, ResponseType und ItemNotificationMessageType.</span><span class="sxs-lookup"><span data-stu-id="4fa79-124">[MailboxEnums](/javascript/api/outlook/office.mailboxenums.attachmenttype): Includes the ItemType, EntityType, AttachmentType, RecipientType, ResponseType, and ItemNotificationMessageType enumerations.</span></span>

### <a name="members"></a><span data-ttu-id="4fa79-125">Elemente</span><span class="sxs-lookup"><span data-stu-id="4fa79-125">Members</span></span>

####  <a name="asyncresultstatus-string"></a><span data-ttu-id="4fa79-126">AsyncResultStatus :String</span><span class="sxs-lookup"><span data-stu-id="4fa79-126">AsyncResultStatus :String</span></span>

<span data-ttu-id="4fa79-127">Gibt das Ergebnis eines asynchronen Aufrufs an.</span><span class="sxs-lookup"><span data-stu-id="4fa79-127">Specifies the result of an asynchronous call.</span></span>

##### <a name="type"></a><span data-ttu-id="4fa79-128">Typ:</span><span class="sxs-lookup"><span data-stu-id="4fa79-128">Type:</span></span>

*   <span data-ttu-id="4fa79-129">Zeichenfolge</span><span class="sxs-lookup"><span data-stu-id="4fa79-129">String</span></span>

##### <a name="properties"></a><span data-ttu-id="4fa79-130">Eigenschaften:</span><span class="sxs-lookup"><span data-stu-id="4fa79-130">Properties:</span></span>

|<span data-ttu-id="4fa79-131">Name</span><span class="sxs-lookup"><span data-stu-id="4fa79-131">Name</span></span>| <span data-ttu-id="4fa79-132">Typ</span><span class="sxs-lookup"><span data-stu-id="4fa79-132">Type</span></span>| <span data-ttu-id="4fa79-133">Beschreibung</span><span class="sxs-lookup"><span data-stu-id="4fa79-133">Description</span></span>|
|---|---|---|
|`Succeeded`| <span data-ttu-id="4fa79-134">Zeichenfolge</span><span class="sxs-lookup"><span data-stu-id="4fa79-134">String</span></span>|<span data-ttu-id="4fa79-135">Der Aufruf war erfolgreich.</span><span class="sxs-lookup"><span data-stu-id="4fa79-135">The call succeeded.</span></span>|
|`Failed`| <span data-ttu-id="4fa79-136">Zeichenfolge</span><span class="sxs-lookup"><span data-stu-id="4fa79-136">String</span></span>|<span data-ttu-id="4fa79-137">Der Aufruf ist fehlerhaft.</span><span class="sxs-lookup"><span data-stu-id="4fa79-137">The call failed.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="4fa79-138">Anforderungen</span><span class="sxs-lookup"><span data-stu-id="4fa79-138">Requirements</span></span>

|<span data-ttu-id="4fa79-139">Anforderung</span><span class="sxs-lookup"><span data-stu-id="4fa79-139">Requirement</span></span>| <span data-ttu-id="4fa79-140">Wert</span><span class="sxs-lookup"><span data-stu-id="4fa79-140">Value</span></span>|
|---|---|
|[<span data-ttu-id="4fa79-141">Mindestversion des Postfachanforderungssatzes</span><span class="sxs-lookup"><span data-stu-id="4fa79-141">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="4fa79-142">1.0</span><span class="sxs-lookup"><span data-stu-id="4fa79-142">1.0</span></span>|
|[<span data-ttu-id="4fa79-143">Zutreffender Outlook-Modus</span><span class="sxs-lookup"><span data-stu-id="4fa79-143">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="4fa79-144">Verfassen oder Lesen</span><span class="sxs-lookup"><span data-stu-id="4fa79-144">Compose or read</span></span>|

---

####  <a name="coerciontype-string"></a><span data-ttu-id="4fa79-145">CoercionType :String</span><span class="sxs-lookup"><span data-stu-id="4fa79-145">CoercionType :String</span></span>

<span data-ttu-id="4fa79-146">Gibt an, wie Daten umgewandelt werden sollen, die von der aufgerufenen Methode zurückgegeben oder festgelegt wurden.</span><span class="sxs-lookup"><span data-stu-id="4fa79-146">Specifies how to coerce data returned or set by the invoked method.</span></span>

##### <a name="type"></a><span data-ttu-id="4fa79-147">Typ:</span><span class="sxs-lookup"><span data-stu-id="4fa79-147">Type:</span></span>

*   <span data-ttu-id="4fa79-148">Zeichenfolge</span><span class="sxs-lookup"><span data-stu-id="4fa79-148">String</span></span>

##### <a name="properties"></a><span data-ttu-id="4fa79-149">Eigenschaften:</span><span class="sxs-lookup"><span data-stu-id="4fa79-149">Properties:</span></span>

|<span data-ttu-id="4fa79-150">Name</span><span class="sxs-lookup"><span data-stu-id="4fa79-150">Name</span></span>| <span data-ttu-id="4fa79-151">Typ</span><span class="sxs-lookup"><span data-stu-id="4fa79-151">Type</span></span>| <span data-ttu-id="4fa79-152">Beschreibung</span><span class="sxs-lookup"><span data-stu-id="4fa79-152">Description</span></span>|
|---|---|---|
|`Html`| <span data-ttu-id="4fa79-153">Zeichenfolge</span><span class="sxs-lookup"><span data-stu-id="4fa79-153">String</span></span>|<span data-ttu-id="4fa79-154">Fordert die Rückgabe der Daten im HTML-Format an.</span><span class="sxs-lookup"><span data-stu-id="4fa79-154">Requests the data be returned in HTML format.</span></span>|
|`Text`| <span data-ttu-id="4fa79-155">Zeichenfolge</span><span class="sxs-lookup"><span data-stu-id="4fa79-155">String</span></span>|<span data-ttu-id="4fa79-156">Fordert die Rückgabe der Daten im Textformat an.</span><span class="sxs-lookup"><span data-stu-id="4fa79-156">Requests the data be returned in text format.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="4fa79-157">Anforderungen</span><span class="sxs-lookup"><span data-stu-id="4fa79-157">Requirements</span></span>

|<span data-ttu-id="4fa79-158">Anforderung</span><span class="sxs-lookup"><span data-stu-id="4fa79-158">Requirement</span></span>| <span data-ttu-id="4fa79-159">Wert</span><span class="sxs-lookup"><span data-stu-id="4fa79-159">Value</span></span>|
|---|---|
|[<span data-ttu-id="4fa79-160">Mindestversion des Postfachanforderungssatzes</span><span class="sxs-lookup"><span data-stu-id="4fa79-160">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="4fa79-161">1.0</span><span class="sxs-lookup"><span data-stu-id="4fa79-161">1.0</span></span>|
|[<span data-ttu-id="4fa79-162">Zutreffender Outlook-Modus</span><span class="sxs-lookup"><span data-stu-id="4fa79-162">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="4fa79-163">Verfassen oder Lesen</span><span class="sxs-lookup"><span data-stu-id="4fa79-163">Compose or read</span></span>|

---

####  <a name="eventtype-string"></a><span data-ttu-id="4fa79-164">EventType :String</span><span class="sxs-lookup"><span data-stu-id="4fa79-164">EventType :String</span></span>

<span data-ttu-id="4fa79-165">Gibt das einem Ereignishandler zugeordnete Ereignis an.</span><span class="sxs-lookup"><span data-stu-id="4fa79-165">Specifies the event associated with an event handler.</span></span>

##### <a name="type"></a><span data-ttu-id="4fa79-166">Typ:</span><span class="sxs-lookup"><span data-stu-id="4fa79-166">Type:</span></span>

*   <span data-ttu-id="4fa79-167">Zeichenfolge</span><span class="sxs-lookup"><span data-stu-id="4fa79-167">String</span></span>

##### <a name="properties"></a><span data-ttu-id="4fa79-168">Eigenschaften:</span><span class="sxs-lookup"><span data-stu-id="4fa79-168">Properties:</span></span>

| <span data-ttu-id="4fa79-169">Name</span><span class="sxs-lookup"><span data-stu-id="4fa79-169">Name</span></span> | <span data-ttu-id="4fa79-170">Typ</span><span class="sxs-lookup"><span data-stu-id="4fa79-170">Type</span></span> | <span data-ttu-id="4fa79-171">Beschreibung</span><span class="sxs-lookup"><span data-stu-id="4fa79-171">Description</span></span> |
|---|---|---|
|`ItemChanged`| <span data-ttu-id="4fa79-172">String</span><span class="sxs-lookup"><span data-stu-id="4fa79-172">String</span></span> | <span data-ttu-id="4fa79-173">Das ausgewählte Element wurde geändert.</span><span class="sxs-lookup"><span data-stu-id="4fa79-173">The selected item has changed.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="4fa79-174">Anforderungen</span><span class="sxs-lookup"><span data-stu-id="4fa79-174">Requirements</span></span>

|<span data-ttu-id="4fa79-175">Anforderung</span><span class="sxs-lookup"><span data-stu-id="4fa79-175">Requirement</span></span>| <span data-ttu-id="4fa79-176">Wert</span><span class="sxs-lookup"><span data-stu-id="4fa79-176">Value</span></span>|
|---|---|
|[<span data-ttu-id="4fa79-177">Mindestversion des Postfachanforderungssatzes</span><span class="sxs-lookup"><span data-stu-id="4fa79-177">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="4fa79-178">1,5</span><span class="sxs-lookup"><span data-stu-id="4fa79-178">1.5</span></span> |
|[<span data-ttu-id="4fa79-179">Zutreffender Outlook-Modus</span><span class="sxs-lookup"><span data-stu-id="4fa79-179">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="4fa79-180">Verfassen oder Lesen</span><span class="sxs-lookup"><span data-stu-id="4fa79-180">Compose or read</span></span> |

---

####  <a name="sourceproperty-string"></a><span data-ttu-id="4fa79-181">SourceProperty :String</span><span class="sxs-lookup"><span data-stu-id="4fa79-181">SourceProperty :String</span></span>

<span data-ttu-id="4fa79-182">Gibt die Quelle der Daten an, die von der aufgerufenen Methode zurückgegeben werden.</span><span class="sxs-lookup"><span data-stu-id="4fa79-182">Specifies the source of the data returned by the invoked method.</span></span>

##### <a name="type"></a><span data-ttu-id="4fa79-183">Typ:</span><span class="sxs-lookup"><span data-stu-id="4fa79-183">Type:</span></span>

*   <span data-ttu-id="4fa79-184">Zeichenfolge</span><span class="sxs-lookup"><span data-stu-id="4fa79-184">String</span></span>

##### <a name="properties"></a><span data-ttu-id="4fa79-185">Eigenschaften:</span><span class="sxs-lookup"><span data-stu-id="4fa79-185">Properties:</span></span>

|<span data-ttu-id="4fa79-186">Name</span><span class="sxs-lookup"><span data-stu-id="4fa79-186">Name</span></span>| <span data-ttu-id="4fa79-187">Typ</span><span class="sxs-lookup"><span data-stu-id="4fa79-187">Type</span></span>| <span data-ttu-id="4fa79-188">Beschreibung</span><span class="sxs-lookup"><span data-stu-id="4fa79-188">Description</span></span>|
|---|---|---|
|`Body`| <span data-ttu-id="4fa79-189">Zeichenfolge</span><span class="sxs-lookup"><span data-stu-id="4fa79-189">String</span></span>|<span data-ttu-id="4fa79-190">Die Quelle der Daten stammt aus dem Textkörper einer Nachricht.</span><span class="sxs-lookup"><span data-stu-id="4fa79-190">The source of the data is from the body of a message.</span></span>|
|`Subject`| <span data-ttu-id="4fa79-191">Zeichenfolge</span><span class="sxs-lookup"><span data-stu-id="4fa79-191">String</span></span>|<span data-ttu-id="4fa79-192">Die Quelle der Daten stammt aus dem Betreff einer Nachricht.</span><span class="sxs-lookup"><span data-stu-id="4fa79-192">The source of the data is from the subject of a message.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="4fa79-193">Anforderungen</span><span class="sxs-lookup"><span data-stu-id="4fa79-193">Requirements</span></span>

|<span data-ttu-id="4fa79-194">Anforderung</span><span class="sxs-lookup"><span data-stu-id="4fa79-194">Requirement</span></span>| <span data-ttu-id="4fa79-195">Wert</span><span class="sxs-lookup"><span data-stu-id="4fa79-195">Value</span></span>|
|---|---|
|[<span data-ttu-id="4fa79-196">Mindestversion des Postfachanforderungssatzes</span><span class="sxs-lookup"><span data-stu-id="4fa79-196">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="4fa79-197">1.0</span><span class="sxs-lookup"><span data-stu-id="4fa79-197">1.0</span></span>|
|[<span data-ttu-id="4fa79-198">Zutreffender Outlook-Modus</span><span class="sxs-lookup"><span data-stu-id="4fa79-198">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="4fa79-199">Verfassen oder Lesen</span><span class="sxs-lookup"><span data-stu-id="4fa79-199">Compose or read</span></span>|