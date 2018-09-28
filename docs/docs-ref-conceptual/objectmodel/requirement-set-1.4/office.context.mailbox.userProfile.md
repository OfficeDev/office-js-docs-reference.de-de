
# <a name="userprofile"></a><span data-ttu-id="17bfa-101">userProfile</span><span class="sxs-lookup"><span data-stu-id="17bfa-101">userProfile</span></span>

### <span data-ttu-id="17bfa-p101">[Office](Office.md)[.context](Office.context.md)[.mailbox](Office.context.mailbox.md). userProfile</span><span class="sxs-lookup"><span data-stu-id="17bfa-p101">[Office](Office.md)[.context](Office.context.md)[.mailbox](Office.context.mailbox.md). userProfile</span></span>

##### <a name="requirements"></a><span data-ttu-id="17bfa-104">Anforderungen</span><span class="sxs-lookup"><span data-stu-id="17bfa-104">Requirements</span></span>

|<span data-ttu-id="17bfa-105">Anforderung</span><span class="sxs-lookup"><span data-stu-id="17bfa-105">Requirement</span></span>| <span data-ttu-id="17bfa-106">Wert</span><span class="sxs-lookup"><span data-stu-id="17bfa-106">Value</span></span>|
|---|---|
|[<span data-ttu-id="17bfa-107">Mindestversion des Postfachanforderungssatzes</span><span class="sxs-lookup"><span data-stu-id="17bfa-107">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="17bfa-108">1.0</span><span class="sxs-lookup"><span data-stu-id="17bfa-108">1.0</span></span>|
|[<span data-ttu-id="17bfa-109">Mindestberechtigungsstufe</span><span class="sxs-lookup"><span data-stu-id="17bfa-109">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="17bfa-110">ReadItem</span><span class="sxs-lookup"><span data-stu-id="17bfa-110">ReadItem</span></span>|
|[<span data-ttu-id="17bfa-111">Zutreffender Outlook-Modus</span><span class="sxs-lookup"><span data-stu-id="17bfa-111">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="17bfa-112">Verfassen oder Lesen</span><span class="sxs-lookup"><span data-stu-id="17bfa-112">Compose or read</span></span>|

### <a name="members"></a><span data-ttu-id="17bfa-113">Elemente</span><span class="sxs-lookup"><span data-stu-id="17bfa-113">Members</span></span>

####  <a name="displayname-string"></a><span data-ttu-id="17bfa-114">displayName :String</span><span class="sxs-lookup"><span data-stu-id="17bfa-114">displayName :String</span></span>

<span data-ttu-id="17bfa-115">Ruft den Anzeigenamen des aktuellen Benutzers ab.</span><span class="sxs-lookup"><span data-stu-id="17bfa-115">Gets the user's display name.</span></span>

##### <a name="type"></a><span data-ttu-id="17bfa-116">Typ:</span><span class="sxs-lookup"><span data-stu-id="17bfa-116">Type:</span></span>

*   <span data-ttu-id="17bfa-117">String</span><span class="sxs-lookup"><span data-stu-id="17bfa-117">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="17bfa-118">Anforderungen</span><span class="sxs-lookup"><span data-stu-id="17bfa-118">Requirements</span></span>

|<span data-ttu-id="17bfa-119">Anforderung</span><span class="sxs-lookup"><span data-stu-id="17bfa-119">Requirement</span></span>| <span data-ttu-id="17bfa-120">Wert</span><span class="sxs-lookup"><span data-stu-id="17bfa-120">Value</span></span>|
|---|---|
|[<span data-ttu-id="17bfa-121">Mindestversion des Postfachanforderungssatzes</span><span class="sxs-lookup"><span data-stu-id="17bfa-121">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="17bfa-122">1.0</span><span class="sxs-lookup"><span data-stu-id="17bfa-122">1.0</span></span>|
|[<span data-ttu-id="17bfa-123">Mindestberechtigungsstufe</span><span class="sxs-lookup"><span data-stu-id="17bfa-123">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="17bfa-124">ReadItem</span><span class="sxs-lookup"><span data-stu-id="17bfa-124">ReadItem</span></span>|
|[<span data-ttu-id="17bfa-125">Zutreffender Outlook-Modus</span><span class="sxs-lookup"><span data-stu-id="17bfa-125">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="17bfa-126">Verfassen oder Lesen</span><span class="sxs-lookup"><span data-stu-id="17bfa-126">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="17bfa-127">Beispiel</span><span class="sxs-lookup"><span data-stu-id="17bfa-127">Example</span></span>

```
// Example: Allie Bellew
console.log(Office.context.mailbox.userProfile.displayName);
```

####  <a name="emailaddress-string"></a><span data-ttu-id="17bfa-128">emailAddress :String</span><span class="sxs-lookup"><span data-stu-id="17bfa-128">emailAddress :String</span></span>

<span data-ttu-id="17bfa-129">Ruft die SMTP-E-Mail-Adresse des Benutzers ab.</span><span class="sxs-lookup"><span data-stu-id="17bfa-129">Gets the user's SMTP email address.</span></span>

##### <a name="type"></a><span data-ttu-id="17bfa-130">Typ:</span><span class="sxs-lookup"><span data-stu-id="17bfa-130">Type:</span></span>

*   <span data-ttu-id="17bfa-131">String</span><span class="sxs-lookup"><span data-stu-id="17bfa-131">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="17bfa-132">Anforderungen</span><span class="sxs-lookup"><span data-stu-id="17bfa-132">Requirements</span></span>

|<span data-ttu-id="17bfa-133">Anforderung</span><span class="sxs-lookup"><span data-stu-id="17bfa-133">Requirement</span></span>| <span data-ttu-id="17bfa-134">Wert</span><span class="sxs-lookup"><span data-stu-id="17bfa-134">Value</span></span>|
|---|---|
|[<span data-ttu-id="17bfa-135">Mindestversion des Postfachanforderungssatzes</span><span class="sxs-lookup"><span data-stu-id="17bfa-135">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="17bfa-136">1.0</span><span class="sxs-lookup"><span data-stu-id="17bfa-136">1.0</span></span>|
|[<span data-ttu-id="17bfa-137">Mindestberechtigungsstufe</span><span class="sxs-lookup"><span data-stu-id="17bfa-137">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="17bfa-138">ReadItem</span><span class="sxs-lookup"><span data-stu-id="17bfa-138">ReadItem</span></span>|
|[<span data-ttu-id="17bfa-139">Zutreffender Outlook-Modus</span><span class="sxs-lookup"><span data-stu-id="17bfa-139">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="17bfa-140">Verfassen oder Lesen</span><span class="sxs-lookup"><span data-stu-id="17bfa-140">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="17bfa-141">Beispiel</span><span class="sxs-lookup"><span data-stu-id="17bfa-141">Example</span></span>

```
// Example: allieb@contoso.com
console.log(Office.context.mailbox.userProfile.emailAddress);
```

####  <a name="timezone-string"></a><span data-ttu-id="17bfa-142">timeZone :String</span><span class="sxs-lookup"><span data-stu-id="17bfa-142">timeZone :String</span></span>

<span data-ttu-id="17bfa-143">Ruft die Standardzeitzone des Benutzers ab.</span><span class="sxs-lookup"><span data-stu-id="17bfa-143">Gets the user's default time zone.</span></span>

##### <a name="type"></a><span data-ttu-id="17bfa-144">Typ:</span><span class="sxs-lookup"><span data-stu-id="17bfa-144">Type:</span></span>

*   <span data-ttu-id="17bfa-145">String</span><span class="sxs-lookup"><span data-stu-id="17bfa-145">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="17bfa-146">Anforderungen</span><span class="sxs-lookup"><span data-stu-id="17bfa-146">Requirements</span></span>

|<span data-ttu-id="17bfa-147">Anforderung</span><span class="sxs-lookup"><span data-stu-id="17bfa-147">Requirement</span></span>| <span data-ttu-id="17bfa-148">Wert</span><span class="sxs-lookup"><span data-stu-id="17bfa-148">Value</span></span>|
|---|---|
|[<span data-ttu-id="17bfa-149">Mindestversion des Postfachanforderungssatzes</span><span class="sxs-lookup"><span data-stu-id="17bfa-149">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="17bfa-150">1.0</span><span class="sxs-lookup"><span data-stu-id="17bfa-150">1.0</span></span>|
|[<span data-ttu-id="17bfa-151">Mindestberechtigungsstufe</span><span class="sxs-lookup"><span data-stu-id="17bfa-151">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="17bfa-152">ReadItem</span><span class="sxs-lookup"><span data-stu-id="17bfa-152">ReadItem</span></span>|
|[<span data-ttu-id="17bfa-153">Zutreffender Outlook-Modus</span><span class="sxs-lookup"><span data-stu-id="17bfa-153">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="17bfa-154">Verfassen oder Lesen</span><span class="sxs-lookup"><span data-stu-id="17bfa-154">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="17bfa-155">Beispiel</span><span class="sxs-lookup"><span data-stu-id="17bfa-155">Example</span></span>

```
// Example: Pacific Standard Time
console.log(Office.context.mailbox.userProfile.timeZone);
```