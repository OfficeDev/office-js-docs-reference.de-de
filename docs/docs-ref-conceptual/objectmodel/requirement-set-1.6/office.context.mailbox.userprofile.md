
# <a name="userprofile"></a><span data-ttu-id="4a8e7-101">userProfile</span><span class="sxs-lookup"><span data-stu-id="4a8e7-101">userProfile</span></span>

### <span data-ttu-id="4a8e7-p101">[Office](Office.md)[.context](Office.context.md)[.mailbox](Office.context.mailbox.md). userProfile</span><span class="sxs-lookup"><span data-stu-id="4a8e7-p101">[Office](Office.md)[.context](Office.context.md)[.mailbox](Office.context.mailbox.md). userProfile</span></span>

##### <a name="requirements"></a><span data-ttu-id="4a8e7-104">Anforderungen</span><span class="sxs-lookup"><span data-stu-id="4a8e7-104">Requirements</span></span>

|<span data-ttu-id="4a8e7-105">Anforderung</span><span class="sxs-lookup"><span data-stu-id="4a8e7-105">Requirement</span></span>| <span data-ttu-id="4a8e7-106">Wert</span><span class="sxs-lookup"><span data-stu-id="4a8e7-106">Value</span></span>|
|---|---|
|[<span data-ttu-id="4a8e7-107">Mindestversion des Postfachanforderungssatzes</span><span class="sxs-lookup"><span data-stu-id="4a8e7-107">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="4a8e7-108">1.0</span><span class="sxs-lookup"><span data-stu-id="4a8e7-108">1.0</span></span>|
|[<span data-ttu-id="4a8e7-109">Mindestberechtigungsstufe</span><span class="sxs-lookup"><span data-stu-id="4a8e7-109">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="4a8e7-110">ReadItem</span><span class="sxs-lookup"><span data-stu-id="4a8e7-110">ReadItem</span></span>|
|[<span data-ttu-id="4a8e7-111">Zutreffender Outlook-Modus</span><span class="sxs-lookup"><span data-stu-id="4a8e7-111">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="4a8e7-112">Verfassen oder Lesen</span><span class="sxs-lookup"><span data-stu-id="4a8e7-112">Compose or read</span></span>|

##### <a name="members-and-methods"></a><span data-ttu-id="4a8e7-113">Elemente und Methoden</span><span class="sxs-lookup"><span data-stu-id="4a8e7-113">Members and methods</span></span>

| <span data-ttu-id="4a8e7-114">Element</span><span class="sxs-lookup"><span data-stu-id="4a8e7-114">Member</span></span> | <span data-ttu-id="4a8e7-115">Typ</span><span class="sxs-lookup"><span data-stu-id="4a8e7-115">Type</span></span> |
|--------|------|
| [<span data-ttu-id="4a8e7-116">accountType</span><span class="sxs-lookup"><span data-stu-id="4a8e7-116">accountType</span></span>](#accounttype-string) | <span data-ttu-id="4a8e7-117">Element</span><span class="sxs-lookup"><span data-stu-id="4a8e7-117">Member</span></span> |
| [<span data-ttu-id="4a8e7-118">displayName</span><span class="sxs-lookup"><span data-stu-id="4a8e7-118">displayName</span></span>](#displayname-string) | <span data-ttu-id="4a8e7-119">Element</span><span class="sxs-lookup"><span data-stu-id="4a8e7-119">Member</span></span> |
| [<span data-ttu-id="4a8e7-120">emailAddress</span><span class="sxs-lookup"><span data-stu-id="4a8e7-120">emailAddress</span></span>](#emailaddress-string) | <span data-ttu-id="4a8e7-121">Element</span><span class="sxs-lookup"><span data-stu-id="4a8e7-121">Member</span></span> |
| [<span data-ttu-id="4a8e7-122">timeZone</span><span class="sxs-lookup"><span data-stu-id="4a8e7-122">timeZone</span></span>](#timezone-string) | <span data-ttu-id="4a8e7-123">Element</span><span class="sxs-lookup"><span data-stu-id="4a8e7-123">Member</span></span> |

### <a name="members"></a><span data-ttu-id="4a8e7-124">Elemente</span><span class="sxs-lookup"><span data-stu-id="4a8e7-124">Members</span></span>

####  <a name="accounttype-string"></a><span data-ttu-id="4a8e7-125">AccountType: Zeichenfolge</span><span class="sxs-lookup"><span data-stu-id="4a8e7-125">accountType :String</span></span>

> [!NOTE]
> <span data-ttu-id="4a8e7-126">Dieser Member ist derzeit nur in unterstützten Outlook 2016 für Mac, build 16.9.1212 und höher.</span><span class="sxs-lookup"><span data-stu-id="4a8e7-126">This member is currently only supported in Outlook 2016 for Mac, build 16.9.1212 and greater.</span></span>

<span data-ttu-id="4a8e7-127">Ruft den Kontotyp des Benutzers, dem Postfach zugeordnete ab.</span><span class="sxs-lookup"><span data-stu-id="4a8e7-127">Gets the account type of the user associated with the mailbox.</span></span> <span data-ttu-id="4a8e7-128">In der folgenden Tabelle sind die möglichen Werte aufgeführt.</span><span class="sxs-lookup"><span data-stu-id="4a8e7-128">The possible values are listed in the following table.</span></span>

| <span data-ttu-id="4a8e7-129">Wert</span><span class="sxs-lookup"><span data-stu-id="4a8e7-129">Value</span></span> | <span data-ttu-id="4a8e7-130">Beschreibung</span><span class="sxs-lookup"><span data-stu-id="4a8e7-130">Description</span></span> |
|-------|-------------|
| `enterprise` | <span data-ttu-id="4a8e7-131">Das Postfach befindet sich auf einem lokalen Exchange-Server.</span><span class="sxs-lookup"><span data-stu-id="4a8e7-131">The mailbox is on an on-premises Exchange server.</span></span> |
| `gmail` | <span data-ttu-id="4a8e7-132">Das Postfach ist Google Mail-Konto zugeordnet.</span><span class="sxs-lookup"><span data-stu-id="4a8e7-132">The mailbox is associated with a Gmail account.</span></span> |
| `office365` | <span data-ttu-id="4a8e7-133">Das Postfach mit einer Office 365 arbeiten oder Schule Konto.</span><span class="sxs-lookup"><span data-stu-id="4a8e7-133">The mailbox is associated with an Office 365 work or school account.</span></span> |
| `outlookCom` | <span data-ttu-id="4a8e7-134">Das Postfach ist ein persönliches Outlook.com-Konto zugeordnet.</span><span class="sxs-lookup"><span data-stu-id="4a8e7-134">The mailbox is associated with a personal Outlook.com account.</span></span> |

##### <a name="type"></a><span data-ttu-id="4a8e7-135">Typ:</span><span class="sxs-lookup"><span data-stu-id="4a8e7-135">Type:</span></span>

*   <span data-ttu-id="4a8e7-136">String</span><span class="sxs-lookup"><span data-stu-id="4a8e7-136">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="4a8e7-137">Anforderungen</span><span class="sxs-lookup"><span data-stu-id="4a8e7-137">Requirements</span></span>

|<span data-ttu-id="4a8e7-138">Anforderung</span><span class="sxs-lookup"><span data-stu-id="4a8e7-138">Requirement</span></span>| <span data-ttu-id="4a8e7-139">Wert</span><span class="sxs-lookup"><span data-stu-id="4a8e7-139">Value</span></span>|
|---|---|
|[<span data-ttu-id="4a8e7-140">Mindestversion des Postfachanforderungssatzes</span><span class="sxs-lookup"><span data-stu-id="4a8e7-140">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="4a8e7-141">1.6</span><span class="sxs-lookup"><span data-stu-id="4a8e7-141">1.6</span></span> |
|[<span data-ttu-id="4a8e7-142">Mindestberechtigungsstufe</span><span class="sxs-lookup"><span data-stu-id="4a8e7-142">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="4a8e7-143">ReadItem</span><span class="sxs-lookup"><span data-stu-id="4a8e7-143">ReadItem</span></span>|
|[<span data-ttu-id="4a8e7-144">Zutreffender Outlook-Modus</span><span class="sxs-lookup"><span data-stu-id="4a8e7-144">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="4a8e7-145">Verfassen oder Lesen</span><span class="sxs-lookup"><span data-stu-id="4a8e7-145">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="4a8e7-146">Beispiel</span><span class="sxs-lookup"><span data-stu-id="4a8e7-146">Example</span></span>

```
console.log(Office.context.mailbox.userProfile.accountType);
```

####  <a name="displayname-string"></a><span data-ttu-id="4a8e7-147">displayName :String</span><span class="sxs-lookup"><span data-stu-id="4a8e7-147">displayName :String</span></span>

<span data-ttu-id="4a8e7-148">Ruft den Anzeigenamen des aktuellen Benutzers ab.</span><span class="sxs-lookup"><span data-stu-id="4a8e7-148">Gets the user's display name.</span></span>

##### <a name="type"></a><span data-ttu-id="4a8e7-149">Typ:</span><span class="sxs-lookup"><span data-stu-id="4a8e7-149">Type:</span></span>

*   <span data-ttu-id="4a8e7-150">String</span><span class="sxs-lookup"><span data-stu-id="4a8e7-150">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="4a8e7-151">Anforderungen</span><span class="sxs-lookup"><span data-stu-id="4a8e7-151">Requirements</span></span>

|<span data-ttu-id="4a8e7-152">Anforderung</span><span class="sxs-lookup"><span data-stu-id="4a8e7-152">Requirement</span></span>| <span data-ttu-id="4a8e7-153">Wert</span><span class="sxs-lookup"><span data-stu-id="4a8e7-153">Value</span></span>|
|---|---|
|[<span data-ttu-id="4a8e7-154">Mindestversion des Postfachanforderungssatzes</span><span class="sxs-lookup"><span data-stu-id="4a8e7-154">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="4a8e7-155">1.0</span><span class="sxs-lookup"><span data-stu-id="4a8e7-155">1.0</span></span>|
|[<span data-ttu-id="4a8e7-156">Mindestberechtigungsstufe</span><span class="sxs-lookup"><span data-stu-id="4a8e7-156">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="4a8e7-157">ReadItem</span><span class="sxs-lookup"><span data-stu-id="4a8e7-157">ReadItem</span></span>|
|[<span data-ttu-id="4a8e7-158">Zutreffender Outlook-Modus</span><span class="sxs-lookup"><span data-stu-id="4a8e7-158">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="4a8e7-159">Verfassen oder Lesen</span><span class="sxs-lookup"><span data-stu-id="4a8e7-159">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="4a8e7-160">Beispiel</span><span class="sxs-lookup"><span data-stu-id="4a8e7-160">Example</span></span>

```
// Example: Allie Bellew
console.log(Office.context.mailbox.userProfile.displayName);
```

####  <a name="emailaddress-string"></a><span data-ttu-id="4a8e7-161">emailAddress :String</span><span class="sxs-lookup"><span data-stu-id="4a8e7-161">emailAddress :String</span></span>

<span data-ttu-id="4a8e7-162">Ruft die SMTP-E-Mail-Adresse des Benutzers ab.</span><span class="sxs-lookup"><span data-stu-id="4a8e7-162">Gets the user's SMTP email address.</span></span>

##### <a name="type"></a><span data-ttu-id="4a8e7-163">Typ:</span><span class="sxs-lookup"><span data-stu-id="4a8e7-163">Type:</span></span>

*   <span data-ttu-id="4a8e7-164">String</span><span class="sxs-lookup"><span data-stu-id="4a8e7-164">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="4a8e7-165">Anforderungen</span><span class="sxs-lookup"><span data-stu-id="4a8e7-165">Requirements</span></span>

|<span data-ttu-id="4a8e7-166">Anforderung</span><span class="sxs-lookup"><span data-stu-id="4a8e7-166">Requirement</span></span>| <span data-ttu-id="4a8e7-167">Wert</span><span class="sxs-lookup"><span data-stu-id="4a8e7-167">Value</span></span>|
|---|---|
|[<span data-ttu-id="4a8e7-168">Mindestversion des Postfachanforderungssatzes</span><span class="sxs-lookup"><span data-stu-id="4a8e7-168">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="4a8e7-169">1.0</span><span class="sxs-lookup"><span data-stu-id="4a8e7-169">1.0</span></span>|
|[<span data-ttu-id="4a8e7-170">Mindestberechtigungsstufe</span><span class="sxs-lookup"><span data-stu-id="4a8e7-170">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="4a8e7-171">ReadItem</span><span class="sxs-lookup"><span data-stu-id="4a8e7-171">ReadItem</span></span>|
|[<span data-ttu-id="4a8e7-172">Zutreffender Outlook-Modus</span><span class="sxs-lookup"><span data-stu-id="4a8e7-172">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="4a8e7-173">Verfassen oder Lesen</span><span class="sxs-lookup"><span data-stu-id="4a8e7-173">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="4a8e7-174">Beispiel</span><span class="sxs-lookup"><span data-stu-id="4a8e7-174">Example</span></span>

```
// Example: allieb@contoso.com
console.log(Office.context.mailbox.userProfile.emailAddress);
```

####  <a name="timezone-string"></a><span data-ttu-id="4a8e7-175">timeZone :String</span><span class="sxs-lookup"><span data-stu-id="4a8e7-175">timeZone :String</span></span>

<span data-ttu-id="4a8e7-176">Ruft die Standardzeitzone des Benutzers ab.</span><span class="sxs-lookup"><span data-stu-id="4a8e7-176">Gets the user's default time zone.</span></span>

##### <a name="type"></a><span data-ttu-id="4a8e7-177">Typ:</span><span class="sxs-lookup"><span data-stu-id="4a8e7-177">Type:</span></span>

*   <span data-ttu-id="4a8e7-178">String</span><span class="sxs-lookup"><span data-stu-id="4a8e7-178">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="4a8e7-179">Anforderungen</span><span class="sxs-lookup"><span data-stu-id="4a8e7-179">Requirements</span></span>

|<span data-ttu-id="4a8e7-180">Anforderung</span><span class="sxs-lookup"><span data-stu-id="4a8e7-180">Requirement</span></span>| <span data-ttu-id="4a8e7-181">Wert</span><span class="sxs-lookup"><span data-stu-id="4a8e7-181">Value</span></span>|
|---|---|
|[<span data-ttu-id="4a8e7-182">Mindestversion des Postfachanforderungssatzes</span><span class="sxs-lookup"><span data-stu-id="4a8e7-182">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="4a8e7-183">1.0</span><span class="sxs-lookup"><span data-stu-id="4a8e7-183">1.0</span></span>|
|[<span data-ttu-id="4a8e7-184">Mindestberechtigungsstufe</span><span class="sxs-lookup"><span data-stu-id="4a8e7-184">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="4a8e7-185">ReadItem</span><span class="sxs-lookup"><span data-stu-id="4a8e7-185">ReadItem</span></span>|
|[<span data-ttu-id="4a8e7-186">Zutreffender Outlook-Modus</span><span class="sxs-lookup"><span data-stu-id="4a8e7-186">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="4a8e7-187">Verfassen oder Lesen</span><span class="sxs-lookup"><span data-stu-id="4a8e7-187">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="4a8e7-188">Beispiel</span><span class="sxs-lookup"><span data-stu-id="4a8e7-188">Example</span></span>

```
// Example: Pacific Standard Time
console.log(Office.context.mailbox.userProfile.timeZone);
```