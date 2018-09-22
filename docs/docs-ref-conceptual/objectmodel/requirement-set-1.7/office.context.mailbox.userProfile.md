
# <a name="userprofile"></a><span data-ttu-id="5d7d0-101">userProfile</span><span class="sxs-lookup"><span data-stu-id="5d7d0-101">userProfile</span></span>

### <span data-ttu-id="5d7d0-p101">[Office](Office.md)[.context](Office.context.md)[.mailbox](Office.context.mailbox.md). userProfile</span><span class="sxs-lookup"><span data-stu-id="5d7d0-p101">[Office](Office.md)[.context](Office.context.md)[.mailbox](Office.context.mailbox.md). userProfile</span></span>

##### <a name="requirements"></a><span data-ttu-id="5d7d0-104">Anforderungen</span><span class="sxs-lookup"><span data-stu-id="5d7d0-104">Requirements</span></span>

|<span data-ttu-id="5d7d0-105">Anforderung</span><span class="sxs-lookup"><span data-stu-id="5d7d0-105">Requirement</span></span>| <span data-ttu-id="5d7d0-106">Wert</span><span class="sxs-lookup"><span data-stu-id="5d7d0-106">Value</span></span>|
|---|---|
|[<span data-ttu-id="5d7d0-107">Mindestversion des Postfachanforderungssatzes</span><span class="sxs-lookup"><span data-stu-id="5d7d0-107">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="5d7d0-108">1.0</span><span class="sxs-lookup"><span data-stu-id="5d7d0-108">1.0</span></span>|
|[<span data-ttu-id="5d7d0-109">Mindestberechtigungsstufe</span><span class="sxs-lookup"><span data-stu-id="5d7d0-109">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="5d7d0-110">ReadItem</span><span class="sxs-lookup"><span data-stu-id="5d7d0-110">ReadItem</span></span>|
|[<span data-ttu-id="5d7d0-111">Zutreffender Outlook-Modus</span><span class="sxs-lookup"><span data-stu-id="5d7d0-111">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="5d7d0-112">Verfassen oder Lesen</span><span class="sxs-lookup"><span data-stu-id="5d7d0-112">Compose or read</span></span>|

##### <a name="members-and-methods"></a><span data-ttu-id="5d7d0-113">Elemente und Methoden</span><span class="sxs-lookup"><span data-stu-id="5d7d0-113">Members and methods</span></span>

| <span data-ttu-id="5d7d0-114">Element</span><span class="sxs-lookup"><span data-stu-id="5d7d0-114">Member</span></span> | <span data-ttu-id="5d7d0-115">Typ</span><span class="sxs-lookup"><span data-stu-id="5d7d0-115">Type</span></span> |
|--------|------|
| [<span data-ttu-id="5d7d0-116">accountType</span><span class="sxs-lookup"><span data-stu-id="5d7d0-116">accountType</span></span>](#accounttype-string) | <span data-ttu-id="5d7d0-117">Element</span><span class="sxs-lookup"><span data-stu-id="5d7d0-117">Member</span></span> |
| [<span data-ttu-id="5d7d0-118">displayName</span><span class="sxs-lookup"><span data-stu-id="5d7d0-118">displayName</span></span>](#displayname-string) | <span data-ttu-id="5d7d0-119">Element</span><span class="sxs-lookup"><span data-stu-id="5d7d0-119">Member</span></span> |
| [<span data-ttu-id="5d7d0-120">emailAddress</span><span class="sxs-lookup"><span data-stu-id="5d7d0-120">emailAddress</span></span>](#emailaddress-string) | <span data-ttu-id="5d7d0-121">Element</span><span class="sxs-lookup"><span data-stu-id="5d7d0-121">Member</span></span> |
| [<span data-ttu-id="5d7d0-122">timeZone</span><span class="sxs-lookup"><span data-stu-id="5d7d0-122">timeZone</span></span>](#timezone-string) | <span data-ttu-id="5d7d0-123">Element</span><span class="sxs-lookup"><span data-stu-id="5d7d0-123">Member</span></span> |

### <a name="members"></a><span data-ttu-id="5d7d0-124">Elemente</span><span class="sxs-lookup"><span data-stu-id="5d7d0-124">Members</span></span>

####  <a name="accounttype-string"></a><span data-ttu-id="5d7d0-125">AccountType: Zeichenfolge</span><span class="sxs-lookup"><span data-stu-id="5d7d0-125">accountType :String</span></span>

> [!NOTE]
> <span data-ttu-id="5d7d0-126">Dieser Member ist derzeit nur in unterstützten Outlook 2016 für Mac, build 16.9.1212 und höher.</span><span class="sxs-lookup"><span data-stu-id="5d7d0-126">This member is currently only supported in Outlook 2016 for Mac, build 16.9.1212 and greater.</span></span>

<span data-ttu-id="5d7d0-127">Ruft den Kontotyp des Benutzers, dem Postfach zugeordnete ab.</span><span class="sxs-lookup"><span data-stu-id="5d7d0-127">Gets the account type of the user associated with the mailbox.</span></span> <span data-ttu-id="5d7d0-128">In der folgenden Tabelle sind die möglichen Werte aufgeführt.</span><span class="sxs-lookup"><span data-stu-id="5d7d0-128">The possible values are listed in the following table.</span></span>

| <span data-ttu-id="5d7d0-129">Wert</span><span class="sxs-lookup"><span data-stu-id="5d7d0-129">Value</span></span> | <span data-ttu-id="5d7d0-130">Beschreibung</span><span class="sxs-lookup"><span data-stu-id="5d7d0-130">Description</span></span> |
|-------|-------------|
| `enterprise` | <span data-ttu-id="5d7d0-131">Das Postfach befindet sich auf einem lokalen Exchange-Server.</span><span class="sxs-lookup"><span data-stu-id="5d7d0-131">The mailbox is on an on-premises Exchange server.</span></span> |
| `gmail` | <span data-ttu-id="5d7d0-132">Das Postfach ist Google Mail-Konto zugeordnet.</span><span class="sxs-lookup"><span data-stu-id="5d7d0-132">The mailbox is associated with a Gmail account.</span></span> |
| `office365` | <span data-ttu-id="5d7d0-133">Das Postfach mit einer Office 365 arbeiten oder Schule Konto.</span><span class="sxs-lookup"><span data-stu-id="5d7d0-133">The mailbox is associated with an Office 365 work or school account.</span></span> |
| `outlookCom` | <span data-ttu-id="5d7d0-134">Das Postfach ist ein persönliches Outlook.com-Konto zugeordnet.</span><span class="sxs-lookup"><span data-stu-id="5d7d0-134">The mailbox is associated with a personal Outlook.com account.</span></span> |

##### <a name="type"></a><span data-ttu-id="5d7d0-135">Typ:</span><span class="sxs-lookup"><span data-stu-id="5d7d0-135">Type:</span></span>

*   <span data-ttu-id="5d7d0-136">String</span><span class="sxs-lookup"><span data-stu-id="5d7d0-136">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="5d7d0-137">Anforderungen</span><span class="sxs-lookup"><span data-stu-id="5d7d0-137">Requirements</span></span>

|<span data-ttu-id="5d7d0-138">Anforderung</span><span class="sxs-lookup"><span data-stu-id="5d7d0-138">Requirement</span></span>| <span data-ttu-id="5d7d0-139">Wert</span><span class="sxs-lookup"><span data-stu-id="5d7d0-139">Value</span></span>|
|---|---|
|[<span data-ttu-id="5d7d0-140">Mindestversion des Postfachanforderungssatzes</span><span class="sxs-lookup"><span data-stu-id="5d7d0-140">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="5d7d0-141">1.6</span><span class="sxs-lookup"><span data-stu-id="5d7d0-141">1.6</span></span> |
|[<span data-ttu-id="5d7d0-142">Mindestberechtigungsstufe</span><span class="sxs-lookup"><span data-stu-id="5d7d0-142">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="5d7d0-143">ReadItem</span><span class="sxs-lookup"><span data-stu-id="5d7d0-143">ReadItem</span></span>|
|[<span data-ttu-id="5d7d0-144">Zutreffender Outlook-Modus</span><span class="sxs-lookup"><span data-stu-id="5d7d0-144">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="5d7d0-145">Verfassen oder Lesen</span><span class="sxs-lookup"><span data-stu-id="5d7d0-145">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="5d7d0-146">Beispiel</span><span class="sxs-lookup"><span data-stu-id="5d7d0-146">Example</span></span>

```
console.log(Office.context.mailbox.userProfile.accountType);
```

####  <a name="displayname-string"></a><span data-ttu-id="5d7d0-147">displayName :String</span><span class="sxs-lookup"><span data-stu-id="5d7d0-147">displayName :String</span></span>

<span data-ttu-id="5d7d0-148">Ruft den Anzeigenamen des aktuellen Benutzers ab.</span><span class="sxs-lookup"><span data-stu-id="5d7d0-148">Gets the user's display name.</span></span>

##### <a name="type"></a><span data-ttu-id="5d7d0-149">Typ:</span><span class="sxs-lookup"><span data-stu-id="5d7d0-149">Type:</span></span>

*   <span data-ttu-id="5d7d0-150">String</span><span class="sxs-lookup"><span data-stu-id="5d7d0-150">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="5d7d0-151">Anforderungen</span><span class="sxs-lookup"><span data-stu-id="5d7d0-151">Requirements</span></span>

|<span data-ttu-id="5d7d0-152">Anforderung</span><span class="sxs-lookup"><span data-stu-id="5d7d0-152">Requirement</span></span>| <span data-ttu-id="5d7d0-153">Wert</span><span class="sxs-lookup"><span data-stu-id="5d7d0-153">Value</span></span>|
|---|---|
|[<span data-ttu-id="5d7d0-154">Mindestversion des Postfachanforderungssatzes</span><span class="sxs-lookup"><span data-stu-id="5d7d0-154">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="5d7d0-155">1.0</span><span class="sxs-lookup"><span data-stu-id="5d7d0-155">1.0</span></span>|
|[<span data-ttu-id="5d7d0-156">Mindestberechtigungsstufe</span><span class="sxs-lookup"><span data-stu-id="5d7d0-156">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="5d7d0-157">ReadItem</span><span class="sxs-lookup"><span data-stu-id="5d7d0-157">ReadItem</span></span>|
|[<span data-ttu-id="5d7d0-158">Zutreffender Outlook-Modus</span><span class="sxs-lookup"><span data-stu-id="5d7d0-158">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="5d7d0-159">Verfassen oder Lesen</span><span class="sxs-lookup"><span data-stu-id="5d7d0-159">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="5d7d0-160">Beispiel</span><span class="sxs-lookup"><span data-stu-id="5d7d0-160">Example</span></span>

```
// Example: Allie Bellew
console.log(Office.context.mailbox.userProfile.displayName);
```

####  <a name="emailaddress-string"></a><span data-ttu-id="5d7d0-161">emailAddress :String</span><span class="sxs-lookup"><span data-stu-id="5d7d0-161">emailAddress :String</span></span>

<span data-ttu-id="5d7d0-162">Ruft die SMTP-E-Mail-Adresse des Benutzers ab.</span><span class="sxs-lookup"><span data-stu-id="5d7d0-162">Gets the user's SMTP email address.</span></span>

##### <a name="type"></a><span data-ttu-id="5d7d0-163">Typ:</span><span class="sxs-lookup"><span data-stu-id="5d7d0-163">Type:</span></span>

*   <span data-ttu-id="5d7d0-164">String</span><span class="sxs-lookup"><span data-stu-id="5d7d0-164">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="5d7d0-165">Anforderungen</span><span class="sxs-lookup"><span data-stu-id="5d7d0-165">Requirements</span></span>

|<span data-ttu-id="5d7d0-166">Anforderung</span><span class="sxs-lookup"><span data-stu-id="5d7d0-166">Requirement</span></span>| <span data-ttu-id="5d7d0-167">Wert</span><span class="sxs-lookup"><span data-stu-id="5d7d0-167">Value</span></span>|
|---|---|
|[<span data-ttu-id="5d7d0-168">Mindestversion des Postfachanforderungssatzes</span><span class="sxs-lookup"><span data-stu-id="5d7d0-168">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="5d7d0-169">1.0</span><span class="sxs-lookup"><span data-stu-id="5d7d0-169">1.0</span></span>|
|[<span data-ttu-id="5d7d0-170">Mindestberechtigungsstufe</span><span class="sxs-lookup"><span data-stu-id="5d7d0-170">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="5d7d0-171">ReadItem</span><span class="sxs-lookup"><span data-stu-id="5d7d0-171">ReadItem</span></span>|
|[<span data-ttu-id="5d7d0-172">Zutreffender Outlook-Modus</span><span class="sxs-lookup"><span data-stu-id="5d7d0-172">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="5d7d0-173">Verfassen oder Lesen</span><span class="sxs-lookup"><span data-stu-id="5d7d0-173">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="5d7d0-174">Beispiel</span><span class="sxs-lookup"><span data-stu-id="5d7d0-174">Example</span></span>

```
// Example: allieb@contoso.com
console.log(Office.context.mailbox.userProfile.emailAddress);
```

####  <a name="timezone-string"></a><span data-ttu-id="5d7d0-175">timeZone :String</span><span class="sxs-lookup"><span data-stu-id="5d7d0-175">timeZone :String</span></span>

<span data-ttu-id="5d7d0-176">Ruft die Standardzeitzone des Benutzers ab.</span><span class="sxs-lookup"><span data-stu-id="5d7d0-176">Gets the user's default time zone.</span></span>

##### <a name="type"></a><span data-ttu-id="5d7d0-177">Typ:</span><span class="sxs-lookup"><span data-stu-id="5d7d0-177">Type:</span></span>

*   <span data-ttu-id="5d7d0-178">String</span><span class="sxs-lookup"><span data-stu-id="5d7d0-178">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="5d7d0-179">Anforderungen</span><span class="sxs-lookup"><span data-stu-id="5d7d0-179">Requirements</span></span>

|<span data-ttu-id="5d7d0-180">Anforderung</span><span class="sxs-lookup"><span data-stu-id="5d7d0-180">Requirement</span></span>| <span data-ttu-id="5d7d0-181">Wert</span><span class="sxs-lookup"><span data-stu-id="5d7d0-181">Value</span></span>|
|---|---|
|[<span data-ttu-id="5d7d0-182">Mindestversion des Postfachanforderungssatzes</span><span class="sxs-lookup"><span data-stu-id="5d7d0-182">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="5d7d0-183">1.0</span><span class="sxs-lookup"><span data-stu-id="5d7d0-183">1.0</span></span>|
|[<span data-ttu-id="5d7d0-184">Mindestberechtigungsstufe</span><span class="sxs-lookup"><span data-stu-id="5d7d0-184">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="5d7d0-185">ReadItem</span><span class="sxs-lookup"><span data-stu-id="5d7d0-185">ReadItem</span></span>|
|[<span data-ttu-id="5d7d0-186">Zutreffender Outlook-Modus</span><span class="sxs-lookup"><span data-stu-id="5d7d0-186">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="5d7d0-187">Verfassen oder Lesen</span><span class="sxs-lookup"><span data-stu-id="5d7d0-187">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="5d7d0-188">Beispiel</span><span class="sxs-lookup"><span data-stu-id="5d7d0-188">Example</span></span>

```
// Example: Pacific Standard Time
console.log(Office.context.mailbox.userProfile.timeZone);
```