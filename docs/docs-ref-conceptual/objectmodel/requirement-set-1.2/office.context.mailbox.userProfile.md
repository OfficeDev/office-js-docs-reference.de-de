
# <a name="userprofile"></a><span data-ttu-id="4c7b4-101">userProfile</span><span class="sxs-lookup"><span data-stu-id="4c7b4-101">userProfile</span></span>

### <span data-ttu-id="4c7b4-p101">[Office](Office.md)[.context](Office.context.md)[.mailbox](Office.context.mailbox.md). userProfile</span><span class="sxs-lookup"><span data-stu-id="4c7b4-p101">[Office](Office.md)[.context](Office.context.md)[.mailbox](Office.context.mailbox.md). userProfile</span></span>

##### <a name="requirements"></a><span data-ttu-id="4c7b4-104">Anforderungen</span><span class="sxs-lookup"><span data-stu-id="4c7b4-104">Requirements</span></span>

|<span data-ttu-id="4c7b4-105">Anforderung</span><span class="sxs-lookup"><span data-stu-id="4c7b4-105">Requirement</span></span>| <span data-ttu-id="4c7b4-106">Wert</span><span class="sxs-lookup"><span data-stu-id="4c7b4-106">Value</span></span>|
|---|---|
|[<span data-ttu-id="4c7b4-107">Mindestversion des Postfachanforderungssatzes</span><span class="sxs-lookup"><span data-stu-id="4c7b4-107">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="4c7b4-108">1.0</span><span class="sxs-lookup"><span data-stu-id="4c7b4-108">1.0</span></span>|
|[<span data-ttu-id="4c7b4-109">Mindestberechtigungsstufe</span><span class="sxs-lookup"><span data-stu-id="4c7b4-109">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="4c7b4-110">ReadItem</span><span class="sxs-lookup"><span data-stu-id="4c7b4-110">ReadItem</span></span>|
|[<span data-ttu-id="4c7b4-111">Zutreffender Outlook-Modus</span><span class="sxs-lookup"><span data-stu-id="4c7b4-111">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="4c7b4-112">Verfassen oder Lesen</span><span class="sxs-lookup"><span data-stu-id="4c7b4-112">Compose or read</span></span>|

### <a name="members"></a><span data-ttu-id="4c7b4-113">Elemente</span><span class="sxs-lookup"><span data-stu-id="4c7b4-113">Members</span></span>

####  <a name="displayname-string"></a><span data-ttu-id="4c7b4-114">displayName :String</span><span class="sxs-lookup"><span data-stu-id="4c7b4-114">displayName :String</span></span>

<span data-ttu-id="4c7b4-115">Ruft den Anzeigenamen des aktuellen Benutzers ab.</span><span class="sxs-lookup"><span data-stu-id="4c7b4-115">Gets the user's display name.</span></span>

##### <a name="type"></a><span data-ttu-id="4c7b4-116">Typ:</span><span class="sxs-lookup"><span data-stu-id="4c7b4-116">Type:</span></span>

*   <span data-ttu-id="4c7b4-117">String</span><span class="sxs-lookup"><span data-stu-id="4c7b4-117">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="4c7b4-118">Anforderungen</span><span class="sxs-lookup"><span data-stu-id="4c7b4-118">Requirements</span></span>

|<span data-ttu-id="4c7b4-119">Anforderung</span><span class="sxs-lookup"><span data-stu-id="4c7b4-119">Requirement</span></span>| <span data-ttu-id="4c7b4-120">Wert</span><span class="sxs-lookup"><span data-stu-id="4c7b4-120">Value</span></span>|
|---|---|
|[<span data-ttu-id="4c7b4-121">Mindestversion des Postfachanforderungssatzes</span><span class="sxs-lookup"><span data-stu-id="4c7b4-121">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="4c7b4-122">1.0</span><span class="sxs-lookup"><span data-stu-id="4c7b4-122">1.0</span></span>|
|[<span data-ttu-id="4c7b4-123">Mindestberechtigungsstufe</span><span class="sxs-lookup"><span data-stu-id="4c7b4-123">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="4c7b4-124">ReadItem</span><span class="sxs-lookup"><span data-stu-id="4c7b4-124">ReadItem</span></span>|
|[<span data-ttu-id="4c7b4-125">Zutreffender Outlook-Modus</span><span class="sxs-lookup"><span data-stu-id="4c7b4-125">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="4c7b4-126">Verfassen oder Lesen</span><span class="sxs-lookup"><span data-stu-id="4c7b4-126">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="4c7b4-127">Beispiel</span><span class="sxs-lookup"><span data-stu-id="4c7b4-127">Example</span></span>

```
// Example: Allie Bellew
console.log(Office.context.mailbox.userProfile.displayName);
```

####  <a name="emailaddress-string"></a><span data-ttu-id="4c7b4-128">emailAddress :String</span><span class="sxs-lookup"><span data-stu-id="4c7b4-128">emailAddress :String</span></span>

<span data-ttu-id="4c7b4-129">Ruft die SMTP-E-Mail-Adresse des Benutzers ab.</span><span class="sxs-lookup"><span data-stu-id="4c7b4-129">Gets the user's SMTP email address.</span></span>

##### <a name="type"></a><span data-ttu-id="4c7b4-130">Typ:</span><span class="sxs-lookup"><span data-stu-id="4c7b4-130">Type:</span></span>

*   <span data-ttu-id="4c7b4-131">String</span><span class="sxs-lookup"><span data-stu-id="4c7b4-131">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="4c7b4-132">Anforderungen</span><span class="sxs-lookup"><span data-stu-id="4c7b4-132">Requirements</span></span>

|<span data-ttu-id="4c7b4-133">Anforderung</span><span class="sxs-lookup"><span data-stu-id="4c7b4-133">Requirement</span></span>| <span data-ttu-id="4c7b4-134">Wert</span><span class="sxs-lookup"><span data-stu-id="4c7b4-134">Value</span></span>|
|---|---|
|[<span data-ttu-id="4c7b4-135">Mindestversion des Postfachanforderungssatzes</span><span class="sxs-lookup"><span data-stu-id="4c7b4-135">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="4c7b4-136">1.0</span><span class="sxs-lookup"><span data-stu-id="4c7b4-136">1.0</span></span>|
|[<span data-ttu-id="4c7b4-137">Mindestberechtigungsstufe</span><span class="sxs-lookup"><span data-stu-id="4c7b4-137">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="4c7b4-138">ReadItem</span><span class="sxs-lookup"><span data-stu-id="4c7b4-138">ReadItem</span></span>|
|[<span data-ttu-id="4c7b4-139">Zutreffender Outlook-Modus</span><span class="sxs-lookup"><span data-stu-id="4c7b4-139">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="4c7b4-140">Verfassen oder Lesen</span><span class="sxs-lookup"><span data-stu-id="4c7b4-140">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="4c7b4-141">Beispiel</span><span class="sxs-lookup"><span data-stu-id="4c7b4-141">Example</span></span>

```
// Example: allieb@contoso.com
console.log(Office.context.mailbox.userProfile.emailAddress);
```

####  <a name="timezone-string"></a><span data-ttu-id="4c7b4-142">timeZone :String</span><span class="sxs-lookup"><span data-stu-id="4c7b4-142">timeZone :String</span></span>

<span data-ttu-id="4c7b4-143">Ruft die Standardzeitzone des Benutzers ab.</span><span class="sxs-lookup"><span data-stu-id="4c7b4-143">Gets the user's default time zone.</span></span>

##### <a name="type"></a><span data-ttu-id="4c7b4-144">Typ:</span><span class="sxs-lookup"><span data-stu-id="4c7b4-144">Type:</span></span>

*   <span data-ttu-id="4c7b4-145">String</span><span class="sxs-lookup"><span data-stu-id="4c7b4-145">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="4c7b4-146">Anforderungen</span><span class="sxs-lookup"><span data-stu-id="4c7b4-146">Requirements</span></span>

|<span data-ttu-id="4c7b4-147">Anforderung</span><span class="sxs-lookup"><span data-stu-id="4c7b4-147">Requirement</span></span>| <span data-ttu-id="4c7b4-148">Wert</span><span class="sxs-lookup"><span data-stu-id="4c7b4-148">Value</span></span>|
|---|---|
|[<span data-ttu-id="4c7b4-149">Mindestversion des Postfachanforderungssatzes</span><span class="sxs-lookup"><span data-stu-id="4c7b4-149">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="4c7b4-150">1.0</span><span class="sxs-lookup"><span data-stu-id="4c7b4-150">1.0</span></span>|
|[<span data-ttu-id="4c7b4-151">Mindestberechtigungsstufe</span><span class="sxs-lookup"><span data-stu-id="4c7b4-151">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="4c7b4-152">ReadItem</span><span class="sxs-lookup"><span data-stu-id="4c7b4-152">ReadItem</span></span>|
|[<span data-ttu-id="4c7b4-153">Zutreffender Outlook-Modus</span><span class="sxs-lookup"><span data-stu-id="4c7b4-153">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="4c7b4-154">Verfassen oder Lesen</span><span class="sxs-lookup"><span data-stu-id="4c7b4-154">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="4c7b4-155">Beispiel</span><span class="sxs-lookup"><span data-stu-id="4c7b4-155">Example</span></span>

```
// Example: Pacific Standard Time
console.log(Office.context.mailbox.userProfile.timeZone);
```