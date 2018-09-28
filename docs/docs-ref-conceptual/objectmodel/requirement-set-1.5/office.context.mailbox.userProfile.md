# <a name="userprofile"></a><span data-ttu-id="aa6a2-101">userProfile</span><span class="sxs-lookup"><span data-stu-id="aa6a2-101">userProfile</span></span>

### <span data-ttu-id="aa6a2-p101">[Office](Office.md)[.context](Office.context.md)[.mailbox](Office.context.mailbox.md). userProfile</span><span class="sxs-lookup"><span data-stu-id="aa6a2-p101">[Office](Office.md)[.context](Office.context.md)[.mailbox](Office.context.mailbox.md). userProfile</span></span>

##### <a name="requirements"></a><span data-ttu-id="aa6a2-104">Anforderungen</span><span class="sxs-lookup"><span data-stu-id="aa6a2-104">Requirements</span></span>

|<span data-ttu-id="aa6a2-105">Anforderung</span><span class="sxs-lookup"><span data-stu-id="aa6a2-105">Requirement</span></span>| <span data-ttu-id="aa6a2-106">Wert</span><span class="sxs-lookup"><span data-stu-id="aa6a2-106">Value</span></span>|
|---|---|
|[<span data-ttu-id="aa6a2-107">Mindestversion des Postfachanforderungssatzes</span><span class="sxs-lookup"><span data-stu-id="aa6a2-107">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="aa6a2-108">1.0</span><span class="sxs-lookup"><span data-stu-id="aa6a2-108">1.0</span></span>|
|[<span data-ttu-id="aa6a2-109">Mindestberechtigungsstufe</span><span class="sxs-lookup"><span data-stu-id="aa6a2-109">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="aa6a2-110">ReadItem</span><span class="sxs-lookup"><span data-stu-id="aa6a2-110">ReadItem</span></span>|
|[<span data-ttu-id="aa6a2-111">Zutreffender Outlook-Modus</span><span class="sxs-lookup"><span data-stu-id="aa6a2-111">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="aa6a2-112">Verfassen oder Lesen</span><span class="sxs-lookup"><span data-stu-id="aa6a2-112">Compose or read</span></span>|

##### <a name="members-and-methods"></a><span data-ttu-id="aa6a2-113">Elemente und Methoden</span><span class="sxs-lookup"><span data-stu-id="aa6a2-113">Members and methods</span></span>

| <span data-ttu-id="aa6a2-114">Element</span><span class="sxs-lookup"><span data-stu-id="aa6a2-114">Member</span></span> | <span data-ttu-id="aa6a2-115">Typ</span><span class="sxs-lookup"><span data-stu-id="aa6a2-115">Type</span></span> |
|--------|------|
| [<span data-ttu-id="aa6a2-116">displayName</span><span class="sxs-lookup"><span data-stu-id="aa6a2-116">displayName</span></span>](#displayname-string) | <span data-ttu-id="aa6a2-117">Element</span><span class="sxs-lookup"><span data-stu-id="aa6a2-117">Member</span></span> |
| [<span data-ttu-id="aa6a2-118">emailAddress</span><span class="sxs-lookup"><span data-stu-id="aa6a2-118">emailAddress</span></span>](#emailaddress-string) | <span data-ttu-id="aa6a2-119">Element</span><span class="sxs-lookup"><span data-stu-id="aa6a2-119">Member</span></span> |
| [<span data-ttu-id="aa6a2-120">timeZone</span><span class="sxs-lookup"><span data-stu-id="aa6a2-120">timeZone</span></span>](#timezone-string) | <span data-ttu-id="aa6a2-121">Element</span><span class="sxs-lookup"><span data-stu-id="aa6a2-121">Member</span></span> |

### <a name="members"></a><span data-ttu-id="aa6a2-122">Elemente</span><span class="sxs-lookup"><span data-stu-id="aa6a2-122">Members</span></span>

####  <a name="displayname-string"></a><span data-ttu-id="aa6a2-123">displayName :String</span><span class="sxs-lookup"><span data-stu-id="aa6a2-123">displayName :String</span></span>

<span data-ttu-id="aa6a2-124">Ruft den Anzeigenamen des aktuellen Benutzers ab.</span><span class="sxs-lookup"><span data-stu-id="aa6a2-124">Gets the user's display name.</span></span>

##### <a name="type"></a><span data-ttu-id="aa6a2-125">Typ:</span><span class="sxs-lookup"><span data-stu-id="aa6a2-125">Type:</span></span>

*   <span data-ttu-id="aa6a2-126">String</span><span class="sxs-lookup"><span data-stu-id="aa6a2-126">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="aa6a2-127">Anforderungen</span><span class="sxs-lookup"><span data-stu-id="aa6a2-127">Requirements</span></span>

|<span data-ttu-id="aa6a2-128">Anforderung</span><span class="sxs-lookup"><span data-stu-id="aa6a2-128">Requirement</span></span>| <span data-ttu-id="aa6a2-129">Wert</span><span class="sxs-lookup"><span data-stu-id="aa6a2-129">Value</span></span>|
|---|---|
|[<span data-ttu-id="aa6a2-130">Mindestversion des Postfachanforderungssatzes</span><span class="sxs-lookup"><span data-stu-id="aa6a2-130">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="aa6a2-131">1.0</span><span class="sxs-lookup"><span data-stu-id="aa6a2-131">1.0</span></span>|
|[<span data-ttu-id="aa6a2-132">Mindestberechtigungsstufe</span><span class="sxs-lookup"><span data-stu-id="aa6a2-132">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="aa6a2-133">ReadItem</span><span class="sxs-lookup"><span data-stu-id="aa6a2-133">ReadItem</span></span>|
|[<span data-ttu-id="aa6a2-134">Zutreffender Outlook-Modus</span><span class="sxs-lookup"><span data-stu-id="aa6a2-134">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="aa6a2-135">Verfassen oder Lesen</span><span class="sxs-lookup"><span data-stu-id="aa6a2-135">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="aa6a2-136">Beispiel</span><span class="sxs-lookup"><span data-stu-id="aa6a2-136">Example</span></span>

```
// Example: Allie Bellew
console.log(Office.context.mailbox.userProfile.displayName);
```

####  <a name="emailaddress-string"></a><span data-ttu-id="aa6a2-137">emailAddress :String</span><span class="sxs-lookup"><span data-stu-id="aa6a2-137">emailAddress :String</span></span>

<span data-ttu-id="aa6a2-138">Ruft die SMTP-E-Mail-Adresse des Benutzers ab.</span><span class="sxs-lookup"><span data-stu-id="aa6a2-138">Gets the user's SMTP email address.</span></span>

##### <a name="type"></a><span data-ttu-id="aa6a2-139">Typ:</span><span class="sxs-lookup"><span data-stu-id="aa6a2-139">Type:</span></span>

*   <span data-ttu-id="aa6a2-140">String</span><span class="sxs-lookup"><span data-stu-id="aa6a2-140">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="aa6a2-141">Anforderungen</span><span class="sxs-lookup"><span data-stu-id="aa6a2-141">Requirements</span></span>

|<span data-ttu-id="aa6a2-142">Anforderung</span><span class="sxs-lookup"><span data-stu-id="aa6a2-142">Requirement</span></span>| <span data-ttu-id="aa6a2-143">Wert</span><span class="sxs-lookup"><span data-stu-id="aa6a2-143">Value</span></span>|
|---|---|
|[<span data-ttu-id="aa6a2-144">Mindestversion des Postfachanforderungssatzes</span><span class="sxs-lookup"><span data-stu-id="aa6a2-144">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="aa6a2-145">1.0</span><span class="sxs-lookup"><span data-stu-id="aa6a2-145">1.0</span></span>|
|[<span data-ttu-id="aa6a2-146">Mindestberechtigungsstufe</span><span class="sxs-lookup"><span data-stu-id="aa6a2-146">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="aa6a2-147">ReadItem</span><span class="sxs-lookup"><span data-stu-id="aa6a2-147">ReadItem</span></span>|
|[<span data-ttu-id="aa6a2-148">Zutreffender Outlook-Modus</span><span class="sxs-lookup"><span data-stu-id="aa6a2-148">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="aa6a2-149">Verfassen oder Lesen</span><span class="sxs-lookup"><span data-stu-id="aa6a2-149">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="aa6a2-150">Beispiel</span><span class="sxs-lookup"><span data-stu-id="aa6a2-150">Example</span></span>

```
// Example: allieb@contoso.com
console.log(Office.context.mailbox.userProfile.emailAddress);
```

####  <a name="timezone-string"></a><span data-ttu-id="aa6a2-151">timeZone :String</span><span class="sxs-lookup"><span data-stu-id="aa6a2-151">timeZone :String</span></span>

<span data-ttu-id="aa6a2-152">Ruft die Standardzeitzone des Benutzers ab.</span><span class="sxs-lookup"><span data-stu-id="aa6a2-152">Gets the user's default time zone.</span></span>

##### <a name="type"></a><span data-ttu-id="aa6a2-153">Typ:</span><span class="sxs-lookup"><span data-stu-id="aa6a2-153">Type:</span></span>

*   <span data-ttu-id="aa6a2-154">String</span><span class="sxs-lookup"><span data-stu-id="aa6a2-154">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="aa6a2-155">Anforderungen</span><span class="sxs-lookup"><span data-stu-id="aa6a2-155">Requirements</span></span>

|<span data-ttu-id="aa6a2-156">Anforderung</span><span class="sxs-lookup"><span data-stu-id="aa6a2-156">Requirement</span></span>| <span data-ttu-id="aa6a2-157">Wert</span><span class="sxs-lookup"><span data-stu-id="aa6a2-157">Value</span></span>|
|---|---|
|[<span data-ttu-id="aa6a2-158">Mindestversion des Postfachanforderungssatzes</span><span class="sxs-lookup"><span data-stu-id="aa6a2-158">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="aa6a2-159">1.0</span><span class="sxs-lookup"><span data-stu-id="aa6a2-159">1.0</span></span>|
|[<span data-ttu-id="aa6a2-160">Mindestberechtigungsstufe</span><span class="sxs-lookup"><span data-stu-id="aa6a2-160">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="aa6a2-161">ReadItem</span><span class="sxs-lookup"><span data-stu-id="aa6a2-161">ReadItem</span></span>|
|[<span data-ttu-id="aa6a2-162">Zutreffender Outlook-Modus</span><span class="sxs-lookup"><span data-stu-id="aa6a2-162">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="aa6a2-163">Verfassen oder Lesen</span><span class="sxs-lookup"><span data-stu-id="aa6a2-163">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="aa6a2-164">Beispiel</span><span class="sxs-lookup"><span data-stu-id="aa6a2-164">Example</span></span>

```
// Example: Pacific Standard Time
console.log(Office.context.mailbox.userProfile.timeZone);
```