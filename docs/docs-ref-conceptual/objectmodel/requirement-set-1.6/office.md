 

# <a name="office"></a><span data-ttu-id="ee980-101">Büro</span><span class="sxs-lookup"><span data-stu-id="ee980-101">Office</span></span>

<span data-ttu-id="ee980-p101">Der Office-Namespace stellt gemeinsam genutzte Oberflächen bereit, die von Add-Ins in allen Office-Apps verwendet werden. Diese Auflistung dokumentiert nur die Schnittstellen, die von Outlook-Add-Ins verwendet werden. Eine vollständige Auflistung des Office-Namespaces finden Sie im Thema zur [freigegebenen API](/javascript/api/office).</span><span class="sxs-lookup"><span data-stu-id="ee980-p101">The Office namespace provides shared interfaces that are used by add-ins in all of the Office apps. This listing documents only those interfaces that are used by Outlook add-ins. For a full listing of the Office namespace, see the [Shared API](/javascript/api/office).</span></span>

##### <a name="requirements"></a><span data-ttu-id="ee980-104">Anforderungen</span><span class="sxs-lookup"><span data-stu-id="ee980-104">Requirements</span></span>

|<span data-ttu-id="ee980-105">Anforderung</span><span class="sxs-lookup"><span data-stu-id="ee980-105">Requirement</span></span>| <span data-ttu-id="ee980-106">Wert</span><span class="sxs-lookup"><span data-stu-id="ee980-106">Value</span></span>|
|---|---|
|[<span data-ttu-id="ee980-107">Mindestversion des Postfachanforderungssatzes</span><span class="sxs-lookup"><span data-stu-id="ee980-107">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="ee980-108">1.0</span><span class="sxs-lookup"><span data-stu-id="ee980-108">1.0</span></span>|
|[<span data-ttu-id="ee980-109">Zutreffender Outlook-Modus</span><span class="sxs-lookup"><span data-stu-id="ee980-109">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="ee980-110">Verfassen oder Lesen</span><span class="sxs-lookup"><span data-stu-id="ee980-110">Compose or read</span></span>|

##### <a name="members-and-methods"></a><span data-ttu-id="ee980-111">Elemente und Methoden</span><span class="sxs-lookup"><span data-stu-id="ee980-111">Members and methods</span></span>

| <span data-ttu-id="ee980-112">Element</span><span class="sxs-lookup"><span data-stu-id="ee980-112">Member</span></span> | <span data-ttu-id="ee980-113">Typ</span><span class="sxs-lookup"><span data-stu-id="ee980-113">Type</span></span> |
|--------|------|
| [<span data-ttu-id="ee980-114">AsyncResultStatus</span><span class="sxs-lookup"><span data-stu-id="ee980-114">AsyncResultStatus</span></span>](#asyncresultstatus-string) | <span data-ttu-id="ee980-115">Element</span><span class="sxs-lookup"><span data-stu-id="ee980-115">Member</span></span> |
| [<span data-ttu-id="ee980-116">CoercionType</span><span class="sxs-lookup"><span data-stu-id="ee980-116">CoercionType</span></span>](#coerciontype-string) | <span data-ttu-id="ee980-117">Element</span><span class="sxs-lookup"><span data-stu-id="ee980-117">Member</span></span> |
| [<span data-ttu-id="ee980-118">EventType</span><span class="sxs-lookup"><span data-stu-id="ee980-118">EventType</span></span>](#eventtype-string) | <span data-ttu-id="ee980-119">Element</span><span class="sxs-lookup"><span data-stu-id="ee980-119">Member</span></span> |
| [<span data-ttu-id="ee980-120">SourceProperty</span><span class="sxs-lookup"><span data-stu-id="ee980-120">SourceProperty</span></span>](#sourceproperty-string) | <span data-ttu-id="ee980-121">Element</span><span class="sxs-lookup"><span data-stu-id="ee980-121">Member</span></span> |

### <a name="namespaces"></a><span data-ttu-id="ee980-122">Namespaces</span><span class="sxs-lookup"><span data-stu-id="ee980-122">Namespaces</span></span>

<span data-ttu-id="ee980-123">[Kontextmenü](office.context.md): gemeinsam genutzten Oberflächen aus dem Office-Add-ins-API-Kontext-Namespace für die Verwendung in der Outlook-add-in-API bietet.</span><span class="sxs-lookup"><span data-stu-id="ee980-123">[context](office.context.md): Provides shared interfaces from the Office Add-ins API's context namespace for use in the Outlook add-in API.</span></span>

<span data-ttu-id="ee980-124">[MailboxEnums](/javascript/api/outlook/office.mailboxenums.attachmenttype): Enthält die Enumerationen ItemType, EntityType, AttachmentType, RecipientType, ResponseType und ItemNotificationMessageType.</span><span class="sxs-lookup"><span data-stu-id="ee980-124">[MailboxEnums](/javascript/api/outlook/office.mailboxenums.attachmenttype): Includes the ItemType, EntityType, AttachmentType, RecipientType, ResponseType, and ItemNotificationMessageType enumerations.</span></span>

### <a name="members"></a><span data-ttu-id="ee980-125">Elemente</span><span class="sxs-lookup"><span data-stu-id="ee980-125">Members</span></span>

####  <a name="asyncresultstatus-string"></a><span data-ttu-id="ee980-126">AsyncResultStatus :String</span><span class="sxs-lookup"><span data-stu-id="ee980-126">AsyncResultStatus :String</span></span>

<span data-ttu-id="ee980-127">Gibt das Ergebnis eines asynchronen Aufrufs an.</span><span class="sxs-lookup"><span data-stu-id="ee980-127">Specifies the result of an asynchronous call.</span></span>

##### <a name="type"></a><span data-ttu-id="ee980-128">Typ:</span><span class="sxs-lookup"><span data-stu-id="ee980-128">Type:</span></span>

*   <span data-ttu-id="ee980-129">Zeichenfolge</span><span class="sxs-lookup"><span data-stu-id="ee980-129">String</span></span>

##### <a name="properties"></a><span data-ttu-id="ee980-130">Eigenschaften:</span><span class="sxs-lookup"><span data-stu-id="ee980-130">Properties:</span></span>

|<span data-ttu-id="ee980-131">Name</span><span class="sxs-lookup"><span data-stu-id="ee980-131">Name</span></span>| <span data-ttu-id="ee980-132">Typ</span><span class="sxs-lookup"><span data-stu-id="ee980-132">Type</span></span>| <span data-ttu-id="ee980-133">Beschreibung</span><span class="sxs-lookup"><span data-stu-id="ee980-133">Description</span></span>|
|---|---|---|
|`Succeeded`| <span data-ttu-id="ee980-134">Zeichenfolge</span><span class="sxs-lookup"><span data-stu-id="ee980-134">String</span></span>|<span data-ttu-id="ee980-135">Der Aufruf war erfolgreich.</span><span class="sxs-lookup"><span data-stu-id="ee980-135">The call succeeded.</span></span>|
|`Failed`| <span data-ttu-id="ee980-136">Zeichenfolge</span><span class="sxs-lookup"><span data-stu-id="ee980-136">String</span></span>|<span data-ttu-id="ee980-137">Der Aufruf ist fehlerhaft.</span><span class="sxs-lookup"><span data-stu-id="ee980-137">The call failed.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="ee980-138">Anforderungen</span><span class="sxs-lookup"><span data-stu-id="ee980-138">Requirements</span></span>

|<span data-ttu-id="ee980-139">Anforderung</span><span class="sxs-lookup"><span data-stu-id="ee980-139">Requirement</span></span>| <span data-ttu-id="ee980-140">Wert</span><span class="sxs-lookup"><span data-stu-id="ee980-140">Value</span></span>|
|---|---|
|[<span data-ttu-id="ee980-141">Mindestversion des Postfachanforderungssatzes</span><span class="sxs-lookup"><span data-stu-id="ee980-141">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="ee980-142">1.0</span><span class="sxs-lookup"><span data-stu-id="ee980-142">1.0</span></span>|
|[<span data-ttu-id="ee980-143">Zutreffender Outlook-Modus</span><span class="sxs-lookup"><span data-stu-id="ee980-143">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="ee980-144">Verfassen oder Lesen</span><span class="sxs-lookup"><span data-stu-id="ee980-144">Compose or read</span></span>|

---

####  <a name="coerciontype-string"></a><span data-ttu-id="ee980-145">CoercionType :String</span><span class="sxs-lookup"><span data-stu-id="ee980-145">CoercionType :String</span></span>

<span data-ttu-id="ee980-146">Gibt an, wie Daten umgewandelt werden sollen, die von der aufgerufenen Methode zurückgegeben oder festgelegt wurden.</span><span class="sxs-lookup"><span data-stu-id="ee980-146">Specifies how to coerce data returned or set by the invoked method.</span></span>

##### <a name="type"></a><span data-ttu-id="ee980-147">Typ:</span><span class="sxs-lookup"><span data-stu-id="ee980-147">Type:</span></span>

*   <span data-ttu-id="ee980-148">Zeichenfolge</span><span class="sxs-lookup"><span data-stu-id="ee980-148">String</span></span>

##### <a name="properties"></a><span data-ttu-id="ee980-149">Eigenschaften:</span><span class="sxs-lookup"><span data-stu-id="ee980-149">Properties:</span></span>

|<span data-ttu-id="ee980-150">Name</span><span class="sxs-lookup"><span data-stu-id="ee980-150">Name</span></span>| <span data-ttu-id="ee980-151">Typ</span><span class="sxs-lookup"><span data-stu-id="ee980-151">Type</span></span>| <span data-ttu-id="ee980-152">Beschreibung</span><span class="sxs-lookup"><span data-stu-id="ee980-152">Description</span></span>|
|---|---|---|
|`Html`| <span data-ttu-id="ee980-153">Zeichenfolge</span><span class="sxs-lookup"><span data-stu-id="ee980-153">String</span></span>|<span data-ttu-id="ee980-154">Fordert die Rückgabe der Daten im HTML-Format an.</span><span class="sxs-lookup"><span data-stu-id="ee980-154">Requests the data be returned in HTML format.</span></span>|
|`Text`| <span data-ttu-id="ee980-155">Zeichenfolge</span><span class="sxs-lookup"><span data-stu-id="ee980-155">String</span></span>|<span data-ttu-id="ee980-156">Fordert die Rückgabe der Daten im Textformat an.</span><span class="sxs-lookup"><span data-stu-id="ee980-156">Requests the data be returned in text format.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="ee980-157">Anforderungen</span><span class="sxs-lookup"><span data-stu-id="ee980-157">Requirements</span></span>

|<span data-ttu-id="ee980-158">Anforderung</span><span class="sxs-lookup"><span data-stu-id="ee980-158">Requirement</span></span>| <span data-ttu-id="ee980-159">Wert</span><span class="sxs-lookup"><span data-stu-id="ee980-159">Value</span></span>|
|---|---|
|[<span data-ttu-id="ee980-160">Mindestversion des Postfachanforderungssatzes</span><span class="sxs-lookup"><span data-stu-id="ee980-160">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="ee980-161">1.0</span><span class="sxs-lookup"><span data-stu-id="ee980-161">1.0</span></span>|
|[<span data-ttu-id="ee980-162">Zutreffender Outlook-Modus</span><span class="sxs-lookup"><span data-stu-id="ee980-162">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="ee980-163">Verfassen oder Lesen</span><span class="sxs-lookup"><span data-stu-id="ee980-163">Compose or read</span></span>|

---

####  <a name="eventtype-string"></a><span data-ttu-id="ee980-164">EventType :String</span><span class="sxs-lookup"><span data-stu-id="ee980-164">EventType :String</span></span>

<span data-ttu-id="ee980-165">Gibt das einem Ereignishandler zugeordnete Ereignis an.</span><span class="sxs-lookup"><span data-stu-id="ee980-165">Specifies the event associated with an event handler.</span></span>

##### <a name="type"></a><span data-ttu-id="ee980-166">Typ:</span><span class="sxs-lookup"><span data-stu-id="ee980-166">Type:</span></span>

*   <span data-ttu-id="ee980-167">Zeichenfolge</span><span class="sxs-lookup"><span data-stu-id="ee980-167">String</span></span>

##### <a name="properties"></a><span data-ttu-id="ee980-168">Eigenschaften:</span><span class="sxs-lookup"><span data-stu-id="ee980-168">Properties:</span></span>

| <span data-ttu-id="ee980-169">Name</span><span class="sxs-lookup"><span data-stu-id="ee980-169">Name</span></span> | <span data-ttu-id="ee980-170">Typ</span><span class="sxs-lookup"><span data-stu-id="ee980-170">Type</span></span> | <span data-ttu-id="ee980-171">Beschreibung</span><span class="sxs-lookup"><span data-stu-id="ee980-171">Description</span></span> |
|---|---|---|
|`ItemChanged`| <span data-ttu-id="ee980-172">String</span><span class="sxs-lookup"><span data-stu-id="ee980-172">String</span></span> | <span data-ttu-id="ee980-173">Das ausgewählte Element wurde geändert.</span><span class="sxs-lookup"><span data-stu-id="ee980-173">The selected item has changed.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="ee980-174">Anforderungen</span><span class="sxs-lookup"><span data-stu-id="ee980-174">Requirements</span></span>

|<span data-ttu-id="ee980-175">Anforderung</span><span class="sxs-lookup"><span data-stu-id="ee980-175">Requirement</span></span>| <span data-ttu-id="ee980-176">Wert</span><span class="sxs-lookup"><span data-stu-id="ee980-176">Value</span></span>|
|---|---|
|[<span data-ttu-id="ee980-177">Mindestversion des Postfachanforderungssatzes</span><span class="sxs-lookup"><span data-stu-id="ee980-177">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="ee980-178">1,5</span><span class="sxs-lookup"><span data-stu-id="ee980-178">1.5</span></span> |
|[<span data-ttu-id="ee980-179">Zutreffender Outlook-Modus</span><span class="sxs-lookup"><span data-stu-id="ee980-179">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="ee980-180">Verfassen oder Lesen</span><span class="sxs-lookup"><span data-stu-id="ee980-180">Compose or read</span></span> |

---

####  <a name="sourceproperty-string"></a><span data-ttu-id="ee980-181">SourceProperty :String</span><span class="sxs-lookup"><span data-stu-id="ee980-181">SourceProperty :String</span></span>

<span data-ttu-id="ee980-182">Gibt die Quelle der Daten an, die von der aufgerufenen Methode zurückgegeben werden.</span><span class="sxs-lookup"><span data-stu-id="ee980-182">Specifies the source of the data returned by the invoked method.</span></span>

##### <a name="type"></a><span data-ttu-id="ee980-183">Typ:</span><span class="sxs-lookup"><span data-stu-id="ee980-183">Type:</span></span>

*   <span data-ttu-id="ee980-184">Zeichenfolge</span><span class="sxs-lookup"><span data-stu-id="ee980-184">String</span></span>

##### <a name="properties"></a><span data-ttu-id="ee980-185">Eigenschaften:</span><span class="sxs-lookup"><span data-stu-id="ee980-185">Properties:</span></span>

|<span data-ttu-id="ee980-186">Name</span><span class="sxs-lookup"><span data-stu-id="ee980-186">Name</span></span>| <span data-ttu-id="ee980-187">Typ</span><span class="sxs-lookup"><span data-stu-id="ee980-187">Type</span></span>| <span data-ttu-id="ee980-188">Beschreibung</span><span class="sxs-lookup"><span data-stu-id="ee980-188">Description</span></span>|
|---|---|---|
|`Body`| <span data-ttu-id="ee980-189">Zeichenfolge</span><span class="sxs-lookup"><span data-stu-id="ee980-189">String</span></span>|<span data-ttu-id="ee980-190">Die Quelle der Daten stammt aus dem Textkörper einer Nachricht.</span><span class="sxs-lookup"><span data-stu-id="ee980-190">The source of the data is from the body of a message.</span></span>|
|`Subject`| <span data-ttu-id="ee980-191">Zeichenfolge</span><span class="sxs-lookup"><span data-stu-id="ee980-191">String</span></span>|<span data-ttu-id="ee980-192">Die Quelle der Daten stammt aus dem Betreff einer Nachricht.</span><span class="sxs-lookup"><span data-stu-id="ee980-192">The source of the data is from the subject of a message.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="ee980-193">Anforderungen</span><span class="sxs-lookup"><span data-stu-id="ee980-193">Requirements</span></span>

|<span data-ttu-id="ee980-194">Anforderung</span><span class="sxs-lookup"><span data-stu-id="ee980-194">Requirement</span></span>| <span data-ttu-id="ee980-195">Wert</span><span class="sxs-lookup"><span data-stu-id="ee980-195">Value</span></span>|
|---|---|
|[<span data-ttu-id="ee980-196">Mindestversion des Postfachanforderungssatzes</span><span class="sxs-lookup"><span data-stu-id="ee980-196">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="ee980-197">1.0</span><span class="sxs-lookup"><span data-stu-id="ee980-197">1.0</span></span>|
|[<span data-ttu-id="ee980-198">Zutreffender Outlook-Modus</span><span class="sxs-lookup"><span data-stu-id="ee980-198">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="ee980-199">Verfassen oder Lesen</span><span class="sxs-lookup"><span data-stu-id="ee980-199">Compose or read</span></span>|