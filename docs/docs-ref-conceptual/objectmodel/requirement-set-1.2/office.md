 

# <a name="office"></a><span data-ttu-id="0477c-101">Büro</span><span class="sxs-lookup"><span data-stu-id="0477c-101">Office</span></span>

<span data-ttu-id="0477c-p101">Der Office-Namespace stellt gemeinsam genutzte Oberflächen bereit, die von Add-Ins in allen Office-Apps verwendet werden. Diese Auflistung dokumentiert nur die Schnittstellen, die von Outlook-Add-Ins verwendet werden. Eine vollständige Auflistung des Office-Namespaces finden Sie im Thema zur [freigegebenen API](/javascript/api/office).</span><span class="sxs-lookup"><span data-stu-id="0477c-p101">The Office namespace provides shared interfaces that are used by add-ins in all of the Office apps. This listing documents only those interfaces that are used by Outlook add-ins. For a full listing of the Office namespace, see the [Shared API](/javascript/api/office).</span></span>

##### <a name="requirements"></a><span data-ttu-id="0477c-104">Anforderungen</span><span class="sxs-lookup"><span data-stu-id="0477c-104">Requirements</span></span>

|<span data-ttu-id="0477c-105">Anforderung</span><span class="sxs-lookup"><span data-stu-id="0477c-105">Requirement</span></span>| <span data-ttu-id="0477c-106">Wert</span><span class="sxs-lookup"><span data-stu-id="0477c-106">Value</span></span>|
|---|---|
|[<span data-ttu-id="0477c-107">Mindestversion des Postfachanforderungssatzes</span><span class="sxs-lookup"><span data-stu-id="0477c-107">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="0477c-108">1.0</span><span class="sxs-lookup"><span data-stu-id="0477c-108">1.0</span></span>|
|[<span data-ttu-id="0477c-109">Zutreffender Outlook-Modus</span><span class="sxs-lookup"><span data-stu-id="0477c-109">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="0477c-110">Verfassen oder Lesen</span><span class="sxs-lookup"><span data-stu-id="0477c-110">Compose or read</span></span>|

### <a name="namespaces"></a><span data-ttu-id="0477c-111">Namespaces</span><span class="sxs-lookup"><span data-stu-id="0477c-111">Namespaces</span></span>

<span data-ttu-id="0477c-112">[Kontextmenü](office.context.md): gemeinsam genutzten Oberflächen aus dem Office-Add-ins-API-Kontext-Namespace für die Verwendung in der Outlook-add-in-API bietet.</span><span class="sxs-lookup"><span data-stu-id="0477c-112">[context](office.context.md): Provides shared interfaces from the Office Add-ins API's context namespace for use in the Outlook add-in API.</span></span>

<span data-ttu-id="0477c-113">[MailboxEnums](/javascript/api/outlook/office.mailboxenums.attachmenttype): Enthält die Enumerationen ItemType, EntityType, AttachmentType, RecipientType, ResponseType und ItemNotificationMessageType.</span><span class="sxs-lookup"><span data-stu-id="0477c-113">[MailboxEnums](/javascript/api/outlook/office.mailboxenums.attachmenttype): Includes the ItemType, EntityType, AttachmentType, RecipientType, ResponseType, and ItemNotificationMessageType enumerations.</span></span>

### <a name="members"></a><span data-ttu-id="0477c-114">Elemente</span><span class="sxs-lookup"><span data-stu-id="0477c-114">Members</span></span>

####  <a name="asyncresultstatus-string"></a><span data-ttu-id="0477c-115">AsyncResultStatus :String</span><span class="sxs-lookup"><span data-stu-id="0477c-115">AsyncResultStatus :String</span></span>

<span data-ttu-id="0477c-116">Gibt das Ergebnis eines asynchronen Aufrufs an.</span><span class="sxs-lookup"><span data-stu-id="0477c-116">Specifies the result of an asynchronous call.</span></span>

##### <a name="type"></a><span data-ttu-id="0477c-117">Typ:</span><span class="sxs-lookup"><span data-stu-id="0477c-117">Type:</span></span>

*   <span data-ttu-id="0477c-118">Zeichenfolge</span><span class="sxs-lookup"><span data-stu-id="0477c-118">String</span></span>

##### <a name="properties"></a><span data-ttu-id="0477c-119">Eigenschaften:</span><span class="sxs-lookup"><span data-stu-id="0477c-119">Properties:</span></span>

|<span data-ttu-id="0477c-120">Name</span><span class="sxs-lookup"><span data-stu-id="0477c-120">Name</span></span>| <span data-ttu-id="0477c-121">Typ</span><span class="sxs-lookup"><span data-stu-id="0477c-121">Type</span></span>| <span data-ttu-id="0477c-122">Beschreibung</span><span class="sxs-lookup"><span data-stu-id="0477c-122">Description</span></span>|
|---|---|---|
|`Succeeded`| <span data-ttu-id="0477c-123">Zeichenfolge</span><span class="sxs-lookup"><span data-stu-id="0477c-123">String</span></span>|<span data-ttu-id="0477c-124">Der Aufruf war erfolgreich.</span><span class="sxs-lookup"><span data-stu-id="0477c-124">The call succeeded.</span></span>|
|`Failed`| <span data-ttu-id="0477c-125">Zeichenfolge</span><span class="sxs-lookup"><span data-stu-id="0477c-125">String</span></span>|<span data-ttu-id="0477c-126">Der Aufruf ist fehlerhaft.</span><span class="sxs-lookup"><span data-stu-id="0477c-126">The call failed.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="0477c-127">Anforderungen</span><span class="sxs-lookup"><span data-stu-id="0477c-127">Requirements</span></span>

|<span data-ttu-id="0477c-128">Anforderung</span><span class="sxs-lookup"><span data-stu-id="0477c-128">Requirement</span></span>| <span data-ttu-id="0477c-129">Wert</span><span class="sxs-lookup"><span data-stu-id="0477c-129">Value</span></span>|
|---|---|
|[<span data-ttu-id="0477c-130">Mindestversion des Postfachanforderungssatzes</span><span class="sxs-lookup"><span data-stu-id="0477c-130">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="0477c-131">1.0</span><span class="sxs-lookup"><span data-stu-id="0477c-131">1.0</span></span>|
|[<span data-ttu-id="0477c-132">Zutreffender Outlook-Modus</span><span class="sxs-lookup"><span data-stu-id="0477c-132">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="0477c-133">Verfassen oder Lesen</span><span class="sxs-lookup"><span data-stu-id="0477c-133">Compose or read</span></span>|
####  <a name="coerciontype-string"></a><span data-ttu-id="0477c-134">CoercionType :String</span><span class="sxs-lookup"><span data-stu-id="0477c-134">CoercionType :String</span></span>

<span data-ttu-id="0477c-135">Gibt an, wie Daten umgewandelt werden sollen, die von der aufgerufenen Methode zurückgegeben oder festgelegt wurden.</span><span class="sxs-lookup"><span data-stu-id="0477c-135">Specifies how to coerce data returned or set by the invoked method.</span></span>

##### <a name="type"></a><span data-ttu-id="0477c-136">Typ:</span><span class="sxs-lookup"><span data-stu-id="0477c-136">Type:</span></span>

*   <span data-ttu-id="0477c-137">Zeichenfolge</span><span class="sxs-lookup"><span data-stu-id="0477c-137">String</span></span>

##### <a name="properties"></a><span data-ttu-id="0477c-138">Eigenschaften:</span><span class="sxs-lookup"><span data-stu-id="0477c-138">Properties:</span></span>

|<span data-ttu-id="0477c-139">Name</span><span class="sxs-lookup"><span data-stu-id="0477c-139">Name</span></span>| <span data-ttu-id="0477c-140">Typ</span><span class="sxs-lookup"><span data-stu-id="0477c-140">Type</span></span>| <span data-ttu-id="0477c-141">Beschreibung</span><span class="sxs-lookup"><span data-stu-id="0477c-141">Description</span></span>|
|---|---|---|
|`Html`| <span data-ttu-id="0477c-142">Zeichenfolge</span><span class="sxs-lookup"><span data-stu-id="0477c-142">String</span></span>|<span data-ttu-id="0477c-143">Fordert die Rückgabe der Daten im HTML-Format an.</span><span class="sxs-lookup"><span data-stu-id="0477c-143">Requests the data be returned in HTML format.</span></span>|
|`Text`| <span data-ttu-id="0477c-144">Zeichenfolge</span><span class="sxs-lookup"><span data-stu-id="0477c-144">String</span></span>|<span data-ttu-id="0477c-145">Fordert die Rückgabe der Daten im Textformat an.</span><span class="sxs-lookup"><span data-stu-id="0477c-145">Requests the data be returned in text format.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="0477c-146">Anforderungen</span><span class="sxs-lookup"><span data-stu-id="0477c-146">Requirements</span></span>

|<span data-ttu-id="0477c-147">Anforderung</span><span class="sxs-lookup"><span data-stu-id="0477c-147">Requirement</span></span>| <span data-ttu-id="0477c-148">Wert</span><span class="sxs-lookup"><span data-stu-id="0477c-148">Value</span></span>|
|---|---|
|[<span data-ttu-id="0477c-149">Mindestversion des Postfachanforderungssatzes</span><span class="sxs-lookup"><span data-stu-id="0477c-149">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="0477c-150">1.0</span><span class="sxs-lookup"><span data-stu-id="0477c-150">1.0</span></span>|
|[<span data-ttu-id="0477c-151">Zutreffender Outlook-Modus</span><span class="sxs-lookup"><span data-stu-id="0477c-151">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="0477c-152">Verfassen oder Lesen</span><span class="sxs-lookup"><span data-stu-id="0477c-152">Compose or read</span></span>|
####  <a name="sourceproperty-string"></a><span data-ttu-id="0477c-153">SourceProperty :String</span><span class="sxs-lookup"><span data-stu-id="0477c-153">SourceProperty :String</span></span>

<span data-ttu-id="0477c-154">Gibt die Quelle der Daten an, die von der aufgerufenen Methode zurückgegeben werden.</span><span class="sxs-lookup"><span data-stu-id="0477c-154">Specifies the source of the data returned by the invoked method.</span></span>

##### <a name="type"></a><span data-ttu-id="0477c-155">Typ:</span><span class="sxs-lookup"><span data-stu-id="0477c-155">Type:</span></span>

*   <span data-ttu-id="0477c-156">Zeichenfolge</span><span class="sxs-lookup"><span data-stu-id="0477c-156">String</span></span>

##### <a name="properties"></a><span data-ttu-id="0477c-157">Eigenschaften:</span><span class="sxs-lookup"><span data-stu-id="0477c-157">Properties:</span></span>

|<span data-ttu-id="0477c-158">Name</span><span class="sxs-lookup"><span data-stu-id="0477c-158">Name</span></span>| <span data-ttu-id="0477c-159">Typ</span><span class="sxs-lookup"><span data-stu-id="0477c-159">Type</span></span>| <span data-ttu-id="0477c-160">Beschreibung</span><span class="sxs-lookup"><span data-stu-id="0477c-160">Description</span></span>|
|---|---|---|
|`Body`| <span data-ttu-id="0477c-161">Zeichenfolge</span><span class="sxs-lookup"><span data-stu-id="0477c-161">String</span></span>|<span data-ttu-id="0477c-162">Die Quelle der Daten stammt aus dem Textkörper einer Nachricht.</span><span class="sxs-lookup"><span data-stu-id="0477c-162">The source of the data is from the body of a message.</span></span>|
|`Subject`| <span data-ttu-id="0477c-163">Zeichenfolge</span><span class="sxs-lookup"><span data-stu-id="0477c-163">String</span></span>|<span data-ttu-id="0477c-164">Die Quelle der Daten stammt aus dem Betreff einer Nachricht.</span><span class="sxs-lookup"><span data-stu-id="0477c-164">The source of the data is from the subject of a message.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="0477c-165">Anforderungen</span><span class="sxs-lookup"><span data-stu-id="0477c-165">Requirements</span></span>

|<span data-ttu-id="0477c-166">Anforderung</span><span class="sxs-lookup"><span data-stu-id="0477c-166">Requirement</span></span>| <span data-ttu-id="0477c-167">Wert</span><span class="sxs-lookup"><span data-stu-id="0477c-167">Value</span></span>|
|---|---|
|[<span data-ttu-id="0477c-168">Mindestversion des Postfachanforderungssatzes</span><span class="sxs-lookup"><span data-stu-id="0477c-168">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="0477c-169">1.0</span><span class="sxs-lookup"><span data-stu-id="0477c-169">1.0</span></span>|
|[<span data-ttu-id="0477c-170">Zutreffender Outlook-Modus</span><span class="sxs-lookup"><span data-stu-id="0477c-170">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="0477c-171">Verfassen oder Lesen</span><span class="sxs-lookup"><span data-stu-id="0477c-171">Compose or read</span></span>|