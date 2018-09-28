
# <a name="userprofile"></a><span data-ttu-id="2361b-101">userProfile</span><span class="sxs-lookup"><span data-stu-id="2361b-101">userProfile</span></span>

### <span data-ttu-id="2361b-p101">[Office](Office.md)[.context](Office.context.md)[.mailbox](Office.context.mailbox.md). userProfile</span><span class="sxs-lookup"><span data-stu-id="2361b-p101">[Office](Office.md)[.context](Office.context.md)[.mailbox](Office.context.mailbox.md). userProfile</span></span>

##### <a name="requirements"></a><span data-ttu-id="2361b-104">Anforderungen</span><span class="sxs-lookup"><span data-stu-id="2361b-104">Requirements</span></span>

|<span data-ttu-id="2361b-105">Anforderung</span><span class="sxs-lookup"><span data-stu-id="2361b-105">Requirement</span></span>| <span data-ttu-id="2361b-106">Wert</span><span class="sxs-lookup"><span data-stu-id="2361b-106">Value</span></span>|
|---|---|
|[<span data-ttu-id="2361b-107">Mindestversion des Postfachanforderungssatzes</span><span class="sxs-lookup"><span data-stu-id="2361b-107">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="2361b-108">1.0</span><span class="sxs-lookup"><span data-stu-id="2361b-108">1.0</span></span>|
|[<span data-ttu-id="2361b-109">Mindestberechtigungsstufe</span><span class="sxs-lookup"><span data-stu-id="2361b-109">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="2361b-110">ReadItem</span><span class="sxs-lookup"><span data-stu-id="2361b-110">ReadItem</span></span>|
|[<span data-ttu-id="2361b-111">Zutreffender Outlook-Modus</span><span class="sxs-lookup"><span data-stu-id="2361b-111">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="2361b-112">Verfassen oder Lesen</span><span class="sxs-lookup"><span data-stu-id="2361b-112">Compose or read</span></span>|

##### <a name="members-and-methods"></a><span data-ttu-id="2361b-113">Elemente und Methoden</span><span class="sxs-lookup"><span data-stu-id="2361b-113">Members and methods</span></span>

| <span data-ttu-id="2361b-114">Element</span><span class="sxs-lookup"><span data-stu-id="2361b-114">Member</span></span> | <span data-ttu-id="2361b-115">Typ</span><span class="sxs-lookup"><span data-stu-id="2361b-115">Type</span></span> |
|--------|------|
| [<span data-ttu-id="2361b-116">accountType</span><span class="sxs-lookup"><span data-stu-id="2361b-116">accountType</span></span>](#accounttype-string) | <span data-ttu-id="2361b-117">Element</span><span class="sxs-lookup"><span data-stu-id="2361b-117">Member</span></span> |
| [<span data-ttu-id="2361b-118">displayName</span><span class="sxs-lookup"><span data-stu-id="2361b-118">displayName</span></span>](#displayname-string) | <span data-ttu-id="2361b-119">Element</span><span class="sxs-lookup"><span data-stu-id="2361b-119">Member</span></span> |
| [<span data-ttu-id="2361b-120">emailAddress</span><span class="sxs-lookup"><span data-stu-id="2361b-120">emailAddress</span></span>](#emailaddress-string) | <span data-ttu-id="2361b-121">Element</span><span class="sxs-lookup"><span data-stu-id="2361b-121">Member</span></span> |
| [<span data-ttu-id="2361b-122">timeZone</span><span class="sxs-lookup"><span data-stu-id="2361b-122">timeZone</span></span>](#timezone-string) | <span data-ttu-id="2361b-123">Element</span><span class="sxs-lookup"><span data-stu-id="2361b-123">Member</span></span> |

### <a name="members"></a><span data-ttu-id="2361b-124">Elemente</span><span class="sxs-lookup"><span data-stu-id="2361b-124">Members</span></span>

####  <a name="accounttype-string"></a><span data-ttu-id="2361b-125">AccountType: Zeichenfolge</span><span class="sxs-lookup"><span data-stu-id="2361b-125">accountType :String</span></span>

> [!NOTE]
> <span data-ttu-id="2361b-126">Dieser Member wird derzeit nur in Outlook 2016 oder unterstützt höher für Mac (build 16.9.1212 oder höher).</span><span class="sxs-lookup"><span data-stu-id="2361b-126">This member is currently only supported in Outlook 2016 or later for Mac (build 16.9.1212 or later).</span></span>

<span data-ttu-id="2361b-127">Ruft den Kontotyp des Benutzers, dem Postfach zugeordnete ab.</span><span class="sxs-lookup"><span data-stu-id="2361b-127">Gets the account type of the user associated with the mailbox.</span></span> <span data-ttu-id="2361b-128">In der folgenden Tabelle sind die möglichen Werte aufgeführt.</span><span class="sxs-lookup"><span data-stu-id="2361b-128">The possible values are listed in the following table.</span></span>

| <span data-ttu-id="2361b-129">Wert</span><span class="sxs-lookup"><span data-stu-id="2361b-129">Value</span></span> | <span data-ttu-id="2361b-130">Beschreibung</span><span class="sxs-lookup"><span data-stu-id="2361b-130">Description</span></span> |
|-------|-------------|
| `enterprise` | <span data-ttu-id="2361b-131">Das Postfach befindet sich auf einem lokalen Exchange-Server.</span><span class="sxs-lookup"><span data-stu-id="2361b-131">The mailbox is on an on-premises Exchange server.</span></span> |
| `gmail` | <span data-ttu-id="2361b-132">Das Postfach ist Google Mail-Konto zugeordnet.</span><span class="sxs-lookup"><span data-stu-id="2361b-132">The mailbox is associated with a Gmail account.</span></span> |
| `office365` | <span data-ttu-id="2361b-133">Das Postfach mit einer Office 365 arbeiten oder Schule Konto.</span><span class="sxs-lookup"><span data-stu-id="2361b-133">The mailbox is associated with an Office 365 work or school account.</span></span> |
| `outlookCom` | <span data-ttu-id="2361b-134">Das Postfach ist ein persönliches Outlook.com-Konto zugeordnet.</span><span class="sxs-lookup"><span data-stu-id="2361b-134">The mailbox is associated with a personal Outlook.com account.</span></span> |

##### <a name="type"></a><span data-ttu-id="2361b-135">Typ:</span><span class="sxs-lookup"><span data-stu-id="2361b-135">Type:</span></span>

*   <span data-ttu-id="2361b-136">String</span><span class="sxs-lookup"><span data-stu-id="2361b-136">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="2361b-137">Anforderungen</span><span class="sxs-lookup"><span data-stu-id="2361b-137">Requirements</span></span>

|<span data-ttu-id="2361b-138">Anforderung</span><span class="sxs-lookup"><span data-stu-id="2361b-138">Requirement</span></span>| <span data-ttu-id="2361b-139">Wert</span><span class="sxs-lookup"><span data-stu-id="2361b-139">Value</span></span>|
|---|---|
|[<span data-ttu-id="2361b-140">Mindestversion des Postfachanforderungssatzes</span><span class="sxs-lookup"><span data-stu-id="2361b-140">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="2361b-141">1.6</span><span class="sxs-lookup"><span data-stu-id="2361b-141">1.6</span></span> |
|[<span data-ttu-id="2361b-142">Mindestberechtigungsstufe</span><span class="sxs-lookup"><span data-stu-id="2361b-142">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="2361b-143">ReadItem</span><span class="sxs-lookup"><span data-stu-id="2361b-143">ReadItem</span></span>|
|[<span data-ttu-id="2361b-144">Zutreffender Outlook-Modus</span><span class="sxs-lookup"><span data-stu-id="2361b-144">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="2361b-145">Verfassen oder Lesen</span><span class="sxs-lookup"><span data-stu-id="2361b-145">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="2361b-146">Beispiel</span><span class="sxs-lookup"><span data-stu-id="2361b-146">Example</span></span>

```
console.log(Office.context.mailbox.userProfile.accountType);
```

####  <a name="displayname-string"></a><span data-ttu-id="2361b-147">displayName :String</span><span class="sxs-lookup"><span data-stu-id="2361b-147">displayName :String</span></span>

<span data-ttu-id="2361b-148">Ruft den Anzeigenamen des aktuellen Benutzers ab.</span><span class="sxs-lookup"><span data-stu-id="2361b-148">Gets the user's display name.</span></span>

##### <a name="type"></a><span data-ttu-id="2361b-149">Typ:</span><span class="sxs-lookup"><span data-stu-id="2361b-149">Type:</span></span>

*   <span data-ttu-id="2361b-150">String</span><span class="sxs-lookup"><span data-stu-id="2361b-150">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="2361b-151">Anforderungen</span><span class="sxs-lookup"><span data-stu-id="2361b-151">Requirements</span></span>

|<span data-ttu-id="2361b-152">Anforderung</span><span class="sxs-lookup"><span data-stu-id="2361b-152">Requirement</span></span>| <span data-ttu-id="2361b-153">Wert</span><span class="sxs-lookup"><span data-stu-id="2361b-153">Value</span></span>|
|---|---|
|[<span data-ttu-id="2361b-154">Mindestversion des Postfachanforderungssatzes</span><span class="sxs-lookup"><span data-stu-id="2361b-154">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="2361b-155">1.0</span><span class="sxs-lookup"><span data-stu-id="2361b-155">1.0</span></span>|
|[<span data-ttu-id="2361b-156">Mindestberechtigungsstufe</span><span class="sxs-lookup"><span data-stu-id="2361b-156">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="2361b-157">ReadItem</span><span class="sxs-lookup"><span data-stu-id="2361b-157">ReadItem</span></span>|
|[<span data-ttu-id="2361b-158">Zutreffender Outlook-Modus</span><span class="sxs-lookup"><span data-stu-id="2361b-158">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="2361b-159">Verfassen oder Lesen</span><span class="sxs-lookup"><span data-stu-id="2361b-159">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="2361b-160">Beispiel</span><span class="sxs-lookup"><span data-stu-id="2361b-160">Example</span></span>

```
// Example: Allie Bellew
console.log(Office.context.mailbox.userProfile.displayName);
```

####  <a name="emailaddress-string"></a><span data-ttu-id="2361b-161">emailAddress :String</span><span class="sxs-lookup"><span data-stu-id="2361b-161">emailAddress :String</span></span>

<span data-ttu-id="2361b-162">Ruft die SMTP-E-Mail-Adresse des Benutzers ab.</span><span class="sxs-lookup"><span data-stu-id="2361b-162">Gets the user's SMTP email address.</span></span>

##### <a name="type"></a><span data-ttu-id="2361b-163">Typ:</span><span class="sxs-lookup"><span data-stu-id="2361b-163">Type:</span></span>

*   <span data-ttu-id="2361b-164">String</span><span class="sxs-lookup"><span data-stu-id="2361b-164">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="2361b-165">Anforderungen</span><span class="sxs-lookup"><span data-stu-id="2361b-165">Requirements</span></span>

|<span data-ttu-id="2361b-166">Anforderung</span><span class="sxs-lookup"><span data-stu-id="2361b-166">Requirement</span></span>| <span data-ttu-id="2361b-167">Wert</span><span class="sxs-lookup"><span data-stu-id="2361b-167">Value</span></span>|
|---|---|
|[<span data-ttu-id="2361b-168">Mindestversion des Postfachanforderungssatzes</span><span class="sxs-lookup"><span data-stu-id="2361b-168">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="2361b-169">1.0</span><span class="sxs-lookup"><span data-stu-id="2361b-169">1.0</span></span>|
|[<span data-ttu-id="2361b-170">Mindestberechtigungsstufe</span><span class="sxs-lookup"><span data-stu-id="2361b-170">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="2361b-171">ReadItem</span><span class="sxs-lookup"><span data-stu-id="2361b-171">ReadItem</span></span>|
|[<span data-ttu-id="2361b-172">Zutreffender Outlook-Modus</span><span class="sxs-lookup"><span data-stu-id="2361b-172">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="2361b-173">Verfassen oder Lesen</span><span class="sxs-lookup"><span data-stu-id="2361b-173">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="2361b-174">Beispiel</span><span class="sxs-lookup"><span data-stu-id="2361b-174">Example</span></span>

```
// Example: allieb@contoso.com
console.log(Office.context.mailbox.userProfile.emailAddress);
```

####  <a name="timezone-string"></a><span data-ttu-id="2361b-175">timeZone :String</span><span class="sxs-lookup"><span data-stu-id="2361b-175">timeZone :String</span></span>

<span data-ttu-id="2361b-176">Ruft die Standardzeitzone des Benutzers ab.</span><span class="sxs-lookup"><span data-stu-id="2361b-176">Gets the user's default time zone.</span></span>

##### <a name="type"></a><span data-ttu-id="2361b-177">Typ:</span><span class="sxs-lookup"><span data-stu-id="2361b-177">Type:</span></span>

*   <span data-ttu-id="2361b-178">String</span><span class="sxs-lookup"><span data-stu-id="2361b-178">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="2361b-179">Anforderungen</span><span class="sxs-lookup"><span data-stu-id="2361b-179">Requirements</span></span>

|<span data-ttu-id="2361b-180">Anforderung</span><span class="sxs-lookup"><span data-stu-id="2361b-180">Requirement</span></span>| <span data-ttu-id="2361b-181">Wert</span><span class="sxs-lookup"><span data-stu-id="2361b-181">Value</span></span>|
|---|---|
|[<span data-ttu-id="2361b-182">Mindestversion des Postfachanforderungssatzes</span><span class="sxs-lookup"><span data-stu-id="2361b-182">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="2361b-183">1.0</span><span class="sxs-lookup"><span data-stu-id="2361b-183">1.0</span></span>|
|[<span data-ttu-id="2361b-184">Mindestberechtigungsstufe</span><span class="sxs-lookup"><span data-stu-id="2361b-184">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="2361b-185">ReadItem</span><span class="sxs-lookup"><span data-stu-id="2361b-185">ReadItem</span></span>|
|[<span data-ttu-id="2361b-186">Zutreffender Outlook-Modus</span><span class="sxs-lookup"><span data-stu-id="2361b-186">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="2361b-187">Verfassen oder Lesen</span><span class="sxs-lookup"><span data-stu-id="2361b-187">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="2361b-188">Beispiel</span><span class="sxs-lookup"><span data-stu-id="2361b-188">Example</span></span>

```
// Example: Pacific Standard Time
console.log(Office.context.mailbox.userProfile.timeZone);
```