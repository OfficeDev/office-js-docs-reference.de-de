 

# <a name="office"></a><span data-ttu-id="9f226-101">Büro</span><span class="sxs-lookup"><span data-stu-id="9f226-101">Office</span></span>

<span data-ttu-id="9f226-p101">Der Office-Namespace stellt gemeinsam genutzte Oberflächen bereit, die von Add-Ins in allen Office-Apps verwendet werden. Diese Auflistung dokumentiert nur die Schnittstellen, die von Outlook-Add-Ins verwendet werden. Eine vollständige Auflistung des Office-Namespaces finden Sie im Thema zur [freigegebenen API](/javascript/api/office).</span><span class="sxs-lookup"><span data-stu-id="9f226-p101">The Office namespace provides shared interfaces that are used by add-ins in all of the Office apps. This listing documents only those interfaces that are used by Outlook add-ins. For a full listing of the Office namespace, see the [Shared API](/javascript/api/office).</span></span>

##### <a name="requirements"></a><span data-ttu-id="9f226-104">Anforderungen</span><span class="sxs-lookup"><span data-stu-id="9f226-104">Requirements</span></span>

|<span data-ttu-id="9f226-105">Anforderung</span><span class="sxs-lookup"><span data-stu-id="9f226-105">Requirement</span></span>| <span data-ttu-id="9f226-106">Wert</span><span class="sxs-lookup"><span data-stu-id="9f226-106">Value</span></span>|
|---|---|
|[<span data-ttu-id="9f226-107">Mindestversion des Postfachanforderungssatzes</span><span class="sxs-lookup"><span data-stu-id="9f226-107">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="9f226-108">1.0</span><span class="sxs-lookup"><span data-stu-id="9f226-108">1.0</span></span>|
|[<span data-ttu-id="9f226-109">Zutreffender Outlook-Modus</span><span class="sxs-lookup"><span data-stu-id="9f226-109">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="9f226-110">Verfassen oder Lesen</span><span class="sxs-lookup"><span data-stu-id="9f226-110">Compose or read</span></span>|

### <a name="namespaces"></a><span data-ttu-id="9f226-111">Namespaces</span><span class="sxs-lookup"><span data-stu-id="9f226-111">Namespaces</span></span>

<span data-ttu-id="9f226-112">[Kontextmenü](Office.context.md): gemeinsam genutzten Oberflächen aus dem Office-Add-ins-API-Kontext-Namespace für die Verwendung in der Outlook-add-in-API bietet.</span><span class="sxs-lookup"><span data-stu-id="9f226-112">[context](Office.context.md): Provides shared interfaces from the Office Add-ins API's context namespace for use in the Outlook add-in API.</span></span>

<span data-ttu-id="9f226-113">[MailboxEnums](/javascript/api/outlook/office.mailboxenums.attachmenttype): Enthält die Enumerationen ItemType, EntityType, AttachmentType, RecipientType, ResponseType und ItemNotificationMessageType.</span><span class="sxs-lookup"><span data-stu-id="9f226-113">[MailboxEnums](/javascript/api/outlook/office.mailboxenums.attachmenttype): Includes the ItemType, EntityType, AttachmentType, RecipientType, ResponseType, and ItemNotificationMessageType enumerations.</span></span>

### <a name="members"></a><span data-ttu-id="9f226-114">Elemente</span><span class="sxs-lookup"><span data-stu-id="9f226-114">Members</span></span>

####  <a name="asyncresultstatus-string"></a><span data-ttu-id="9f226-115">AsyncResultStatus :String</span><span class="sxs-lookup"><span data-stu-id="9f226-115">AsyncResultStatus :String</span></span>

<span data-ttu-id="9f226-116">Gibt das Ergebnis eines asynchronen Aufrufs an.</span><span class="sxs-lookup"><span data-stu-id="9f226-116">Specifies the result of an asynchronous call.</span></span>

##### <a name="type"></a><span data-ttu-id="9f226-117">Typ:</span><span class="sxs-lookup"><span data-stu-id="9f226-117">Type:</span></span>

*   <span data-ttu-id="9f226-118">Zeichenfolge</span><span class="sxs-lookup"><span data-stu-id="9f226-118">String</span></span>

##### <a name="properties"></a><span data-ttu-id="9f226-119">Eigenschaften:</span><span class="sxs-lookup"><span data-stu-id="9f226-119">Properties:</span></span>

|<span data-ttu-id="9f226-120">Name</span><span class="sxs-lookup"><span data-stu-id="9f226-120">Name</span></span>| <span data-ttu-id="9f226-121">Typ</span><span class="sxs-lookup"><span data-stu-id="9f226-121">Type</span></span>| <span data-ttu-id="9f226-122">Beschreibung</span><span class="sxs-lookup"><span data-stu-id="9f226-122">Description</span></span>|
|---|---|---|
|`Succeeded`| <span data-ttu-id="9f226-123">Zeichenfolge</span><span class="sxs-lookup"><span data-stu-id="9f226-123">String</span></span>|<span data-ttu-id="9f226-124">Der Aufruf war erfolgreich.</span><span class="sxs-lookup"><span data-stu-id="9f226-124">The call succeeded.</span></span>|
|`Failed`| <span data-ttu-id="9f226-125">Zeichenfolge</span><span class="sxs-lookup"><span data-stu-id="9f226-125">String</span></span>|<span data-ttu-id="9f226-126">Der Aufruf ist fehlerhaft.</span><span class="sxs-lookup"><span data-stu-id="9f226-126">The call failed.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="9f226-127">Anforderungen</span><span class="sxs-lookup"><span data-stu-id="9f226-127">Requirements</span></span>

|<span data-ttu-id="9f226-128">Anforderung</span><span class="sxs-lookup"><span data-stu-id="9f226-128">Requirement</span></span>| <span data-ttu-id="9f226-129">Wert</span><span class="sxs-lookup"><span data-stu-id="9f226-129">Value</span></span>|
|---|---|
|[<span data-ttu-id="9f226-130">Mindestversion des Postfachanforderungssatzes</span><span class="sxs-lookup"><span data-stu-id="9f226-130">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="9f226-131">1.0</span><span class="sxs-lookup"><span data-stu-id="9f226-131">1.0</span></span>|
|[<span data-ttu-id="9f226-132">Zutreffender Outlook-Modus</span><span class="sxs-lookup"><span data-stu-id="9f226-132">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="9f226-133">Verfassen oder Lesen</span><span class="sxs-lookup"><span data-stu-id="9f226-133">Compose or read</span></span>|
####  <a name="coerciontype-string"></a><span data-ttu-id="9f226-134">CoercionType :String</span><span class="sxs-lookup"><span data-stu-id="9f226-134">CoercionType :String</span></span>

<span data-ttu-id="9f226-135">Gibt an, wie Daten umgewandelt werden sollen, die von der aufgerufenen Methode zurückgegeben oder festgelegt wurden.</span><span class="sxs-lookup"><span data-stu-id="9f226-135">Specifies how to coerce data returned or set by the invoked method.</span></span>

##### <a name="type"></a><span data-ttu-id="9f226-136">Typ:</span><span class="sxs-lookup"><span data-stu-id="9f226-136">Type:</span></span>

*   <span data-ttu-id="9f226-137">Zeichenfolge</span><span class="sxs-lookup"><span data-stu-id="9f226-137">String</span></span>

##### <a name="properties"></a><span data-ttu-id="9f226-138">Eigenschaften:</span><span class="sxs-lookup"><span data-stu-id="9f226-138">Properties:</span></span>

|<span data-ttu-id="9f226-139">Name</span><span class="sxs-lookup"><span data-stu-id="9f226-139">Name</span></span>| <span data-ttu-id="9f226-140">Typ</span><span class="sxs-lookup"><span data-stu-id="9f226-140">Type</span></span>| <span data-ttu-id="9f226-141">Beschreibung</span><span class="sxs-lookup"><span data-stu-id="9f226-141">Description</span></span>|
|---|---|---|
|`Html`| <span data-ttu-id="9f226-142">Zeichenfolge</span><span class="sxs-lookup"><span data-stu-id="9f226-142">String</span></span>|<span data-ttu-id="9f226-143">Fordert die Rückgabe der Daten im HTML-Format an.</span><span class="sxs-lookup"><span data-stu-id="9f226-143">Requests the data be returned in HTML format.</span></span>|
|`Text`| <span data-ttu-id="9f226-144">Zeichenfolge</span><span class="sxs-lookup"><span data-stu-id="9f226-144">String</span></span>|<span data-ttu-id="9f226-145">Fordert die Rückgabe der Daten im Textformat an.</span><span class="sxs-lookup"><span data-stu-id="9f226-145">Requests the data be returned in text format.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="9f226-146">Anforderungen</span><span class="sxs-lookup"><span data-stu-id="9f226-146">Requirements</span></span>

|<span data-ttu-id="9f226-147">Anforderung</span><span class="sxs-lookup"><span data-stu-id="9f226-147">Requirement</span></span>| <span data-ttu-id="9f226-148">Wert</span><span class="sxs-lookup"><span data-stu-id="9f226-148">Value</span></span>|
|---|---|
|[<span data-ttu-id="9f226-149">Mindestversion des Postfachanforderungssatzes</span><span class="sxs-lookup"><span data-stu-id="9f226-149">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="9f226-150">1.0</span><span class="sxs-lookup"><span data-stu-id="9f226-150">1.0</span></span>|
|[<span data-ttu-id="9f226-151">Zutreffender Outlook-Modus</span><span class="sxs-lookup"><span data-stu-id="9f226-151">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="9f226-152">Verfassen oder Lesen</span><span class="sxs-lookup"><span data-stu-id="9f226-152">Compose or read</span></span>|
####  <a name="sourceproperty-string"></a><span data-ttu-id="9f226-153">SourceProperty :String</span><span class="sxs-lookup"><span data-stu-id="9f226-153">SourceProperty :String</span></span>

<span data-ttu-id="9f226-154">Gibt die Quelle der Daten an, die von der aufgerufenen Methode zurückgegeben werden.</span><span class="sxs-lookup"><span data-stu-id="9f226-154">Specifies the source of the data returned by the invoked method.</span></span>

##### <a name="type"></a><span data-ttu-id="9f226-155">Typ:</span><span class="sxs-lookup"><span data-stu-id="9f226-155">Type:</span></span>

*   <span data-ttu-id="9f226-156">Zeichenfolge</span><span class="sxs-lookup"><span data-stu-id="9f226-156">String</span></span>

##### <a name="properties"></a><span data-ttu-id="9f226-157">Eigenschaften:</span><span class="sxs-lookup"><span data-stu-id="9f226-157">Properties:</span></span>

|<span data-ttu-id="9f226-158">Name</span><span class="sxs-lookup"><span data-stu-id="9f226-158">Name</span></span>| <span data-ttu-id="9f226-159">Typ</span><span class="sxs-lookup"><span data-stu-id="9f226-159">Type</span></span>| <span data-ttu-id="9f226-160">Beschreibung</span><span class="sxs-lookup"><span data-stu-id="9f226-160">Description</span></span>|
|---|---|---|
|`Body`| <span data-ttu-id="9f226-161">Zeichenfolge</span><span class="sxs-lookup"><span data-stu-id="9f226-161">String</span></span>|<span data-ttu-id="9f226-162">Die Quelle der Daten stammt aus dem Textkörper einer Nachricht.</span><span class="sxs-lookup"><span data-stu-id="9f226-162">The source of the data is from the body of a message.</span></span>|
|`Subject`| <span data-ttu-id="9f226-163">Zeichenfolge</span><span class="sxs-lookup"><span data-stu-id="9f226-163">String</span></span>|<span data-ttu-id="9f226-164">Die Quelle der Daten stammt aus dem Betreff einer Nachricht.</span><span class="sxs-lookup"><span data-stu-id="9f226-164">The source of the data is from the subject of a message.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="9f226-165">Anforderungen</span><span class="sxs-lookup"><span data-stu-id="9f226-165">Requirements</span></span>

|<span data-ttu-id="9f226-166">Anforderung</span><span class="sxs-lookup"><span data-stu-id="9f226-166">Requirement</span></span>| <span data-ttu-id="9f226-167">Wert</span><span class="sxs-lookup"><span data-stu-id="9f226-167">Value</span></span>|
|---|---|
|[<span data-ttu-id="9f226-168">Mindestversion des Postfachanforderungssatzes</span><span class="sxs-lookup"><span data-stu-id="9f226-168">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="9f226-169">1.0</span><span class="sxs-lookup"><span data-stu-id="9f226-169">1.0</span></span>|
|[<span data-ttu-id="9f226-170">Zutreffender Outlook-Modus</span><span class="sxs-lookup"><span data-stu-id="9f226-170">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="9f226-171">Verfassen oder Lesen</span><span class="sxs-lookup"><span data-stu-id="9f226-171">Compose or read</span></span>|