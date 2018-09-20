# <a name="userprofile"></a><span data-ttu-id="e725b-101">userProfile</span><span class="sxs-lookup"><span data-stu-id="e725b-101">userProfile</span></span>

### <span data-ttu-id="e725b-p101">[Office](Office.md)[.context](Office.context.md)[.mailbox](Office.context.mailbox.md). userProfile</span><span class="sxs-lookup"><span data-stu-id="e725b-p101">[Office](Office.md)[.context](Office.context.md)[.mailbox](Office.context.mailbox.md). userProfile</span></span>

##### <a name="requirements"></a><span data-ttu-id="e725b-104">Anforderungen</span><span class="sxs-lookup"><span data-stu-id="e725b-104">Requirements</span></span>

|<span data-ttu-id="e725b-105">Anforderung</span><span class="sxs-lookup"><span data-stu-id="e725b-105">Requirement</span></span>| <span data-ttu-id="e725b-106">Wert</span><span class="sxs-lookup"><span data-stu-id="e725b-106">Value</span></span>|
|---|---|
|[<span data-ttu-id="e725b-107">Mindestversion des Postfachanforderungssatzes</span><span class="sxs-lookup"><span data-stu-id="e725b-107">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="e725b-108">1.0</span><span class="sxs-lookup"><span data-stu-id="e725b-108">1.0</span></span>|
|[<span data-ttu-id="e725b-109">Mindestberechtigungsstufe</span><span class="sxs-lookup"><span data-stu-id="e725b-109">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="e725b-110">ReadItem</span><span class="sxs-lookup"><span data-stu-id="e725b-110">ReadItem</span></span>|
|[<span data-ttu-id="e725b-111">Zutreffender Outlook-Modus</span><span class="sxs-lookup"><span data-stu-id="e725b-111">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="e725b-112">Verfassen oder Lesen</span><span class="sxs-lookup"><span data-stu-id="e725b-112">Compose or read</span></span>|

##### <a name="members-and-methods"></a><span data-ttu-id="e725b-113">Elemente und Methoden</span><span class="sxs-lookup"><span data-stu-id="e725b-113">Members and methods</span></span>

| <span data-ttu-id="e725b-114">Element</span><span class="sxs-lookup"><span data-stu-id="e725b-114">Member</span></span> | <span data-ttu-id="e725b-115">Typ</span><span class="sxs-lookup"><span data-stu-id="e725b-115">Type</span></span> |
|--------|------|
| [<span data-ttu-id="e725b-116">displayName</span><span class="sxs-lookup"><span data-stu-id="e725b-116">displayName</span></span>](#displayname-string) | <span data-ttu-id="e725b-117">Element</span><span class="sxs-lookup"><span data-stu-id="e725b-117">Member</span></span> |
| [<span data-ttu-id="e725b-118">emailAddress</span><span class="sxs-lookup"><span data-stu-id="e725b-118">emailAddress</span></span>](#emailaddress-string) | <span data-ttu-id="e725b-119">Element</span><span class="sxs-lookup"><span data-stu-id="e725b-119">Member</span></span> |
| [<span data-ttu-id="e725b-120">timeZone</span><span class="sxs-lookup"><span data-stu-id="e725b-120">timeZone</span></span>](#timezone-string) | <span data-ttu-id="e725b-121">Element</span><span class="sxs-lookup"><span data-stu-id="e725b-121">Member</span></span> |

### <a name="members"></a><span data-ttu-id="e725b-122">Elemente</span><span class="sxs-lookup"><span data-stu-id="e725b-122">Members</span></span>

####  <a name="displayname-string"></a><span data-ttu-id="e725b-123">displayName :String</span><span class="sxs-lookup"><span data-stu-id="e725b-123">displayName :String</span></span>

<span data-ttu-id="e725b-124">Ruft den Anzeigenamen des aktuellen Benutzers ab.</span><span class="sxs-lookup"><span data-stu-id="e725b-124">Gets the user's display name.</span></span>

##### <a name="type"></a><span data-ttu-id="e725b-125">Typ:</span><span class="sxs-lookup"><span data-stu-id="e725b-125">Type:</span></span>

*   <span data-ttu-id="e725b-126">String</span><span class="sxs-lookup"><span data-stu-id="e725b-126">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="e725b-127">Anforderungen</span><span class="sxs-lookup"><span data-stu-id="e725b-127">Requirements</span></span>

|<span data-ttu-id="e725b-128">Anforderung</span><span class="sxs-lookup"><span data-stu-id="e725b-128">Requirement</span></span>| <span data-ttu-id="e725b-129">Wert</span><span class="sxs-lookup"><span data-stu-id="e725b-129">Value</span></span>|
|---|---|
|[<span data-ttu-id="e725b-130">Mindestversion des Postfachanforderungssatzes</span><span class="sxs-lookup"><span data-stu-id="e725b-130">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="e725b-131">1.0</span><span class="sxs-lookup"><span data-stu-id="e725b-131">1.0</span></span>|
|[<span data-ttu-id="e725b-132">Mindestberechtigungsstufe</span><span class="sxs-lookup"><span data-stu-id="e725b-132">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="e725b-133">ReadItem</span><span class="sxs-lookup"><span data-stu-id="e725b-133">ReadItem</span></span>|
|[<span data-ttu-id="e725b-134">Zutreffender Outlook-Modus</span><span class="sxs-lookup"><span data-stu-id="e725b-134">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="e725b-135">Verfassen oder Lesen</span><span class="sxs-lookup"><span data-stu-id="e725b-135">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="e725b-136">Beispiel</span><span class="sxs-lookup"><span data-stu-id="e725b-136">Example</span></span>

```
// Example: Allie Bellew
console.log(Office.context.mailbox.userProfile.displayName);
```

####  <a name="emailaddress-string"></a><span data-ttu-id="e725b-137">emailAddress :String</span><span class="sxs-lookup"><span data-stu-id="e725b-137">emailAddress :String</span></span>

<span data-ttu-id="e725b-138">Ruft die SMTP-E-Mail-Adresse des Benutzers ab.</span><span class="sxs-lookup"><span data-stu-id="e725b-138">Gets the user's SMTP email address.</span></span>

##### <a name="type"></a><span data-ttu-id="e725b-139">Typ:</span><span class="sxs-lookup"><span data-stu-id="e725b-139">Type:</span></span>

*   <span data-ttu-id="e725b-140">String</span><span class="sxs-lookup"><span data-stu-id="e725b-140">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="e725b-141">Anforderungen</span><span class="sxs-lookup"><span data-stu-id="e725b-141">Requirements</span></span>

|<span data-ttu-id="e725b-142">Anforderung</span><span class="sxs-lookup"><span data-stu-id="e725b-142">Requirement</span></span>| <span data-ttu-id="e725b-143">Wert</span><span class="sxs-lookup"><span data-stu-id="e725b-143">Value</span></span>|
|---|---|
|[<span data-ttu-id="e725b-144">Mindestversion des Postfachanforderungssatzes</span><span class="sxs-lookup"><span data-stu-id="e725b-144">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="e725b-145">1.0</span><span class="sxs-lookup"><span data-stu-id="e725b-145">1.0</span></span>|
|[<span data-ttu-id="e725b-146">Mindestberechtigungsstufe</span><span class="sxs-lookup"><span data-stu-id="e725b-146">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="e725b-147">ReadItem</span><span class="sxs-lookup"><span data-stu-id="e725b-147">ReadItem</span></span>|
|[<span data-ttu-id="e725b-148">Zutreffender Outlook-Modus</span><span class="sxs-lookup"><span data-stu-id="e725b-148">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="e725b-149">Verfassen oder Lesen</span><span class="sxs-lookup"><span data-stu-id="e725b-149">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="e725b-150">Beispiel</span><span class="sxs-lookup"><span data-stu-id="e725b-150">Example</span></span>

```
// Example: allieb@contoso.com
console.log(Office.context.mailbox.userProfile.emailAddress);
```

####  <a name="timezone-string"></a><span data-ttu-id="e725b-151">timeZone :String</span><span class="sxs-lookup"><span data-stu-id="e725b-151">timeZone :String</span></span>

<span data-ttu-id="e725b-152">Ruft die Standardzeitzone des Benutzers ab.</span><span class="sxs-lookup"><span data-stu-id="e725b-152">Gets the user's default time zone.</span></span>

##### <a name="type"></a><span data-ttu-id="e725b-153">Typ:</span><span class="sxs-lookup"><span data-stu-id="e725b-153">Type:</span></span>

*   <span data-ttu-id="e725b-154">String</span><span class="sxs-lookup"><span data-stu-id="e725b-154">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="e725b-155">Anforderungen</span><span class="sxs-lookup"><span data-stu-id="e725b-155">Requirements</span></span>

|<span data-ttu-id="e725b-156">Anforderung</span><span class="sxs-lookup"><span data-stu-id="e725b-156">Requirement</span></span>| <span data-ttu-id="e725b-157">Wert</span><span class="sxs-lookup"><span data-stu-id="e725b-157">Value</span></span>|
|---|---|
|[<span data-ttu-id="e725b-158">Mindestversion des Postfachanforderungssatzes</span><span class="sxs-lookup"><span data-stu-id="e725b-158">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="e725b-159">1.0</span><span class="sxs-lookup"><span data-stu-id="e725b-159">1.0</span></span>|
|[<span data-ttu-id="e725b-160">Mindestberechtigungsstufe</span><span class="sxs-lookup"><span data-stu-id="e725b-160">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="e725b-161">ReadItem</span><span class="sxs-lookup"><span data-stu-id="e725b-161">ReadItem</span></span>|
|[<span data-ttu-id="e725b-162">Zutreffender Outlook-Modus</span><span class="sxs-lookup"><span data-stu-id="e725b-162">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="e725b-163">Verfassen oder Lesen</span><span class="sxs-lookup"><span data-stu-id="e725b-163">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="e725b-164">Beispiel</span><span class="sxs-lookup"><span data-stu-id="e725b-164">Example</span></span>

```
// Example: Pacific Standard Time
console.log(Office.context.mailbox.userProfile.timeZone);
```