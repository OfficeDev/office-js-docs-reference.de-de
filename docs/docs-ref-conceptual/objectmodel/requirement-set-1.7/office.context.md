
# <a name="context"></a><span data-ttu-id="534c3-101">context</span><span class="sxs-lookup"><span data-stu-id="534c3-101">context</span></span>

### <a name="officeofficemdcontext"></a><span data-ttu-id="534c3-102">[Office](Office.md).context</span><span class="sxs-lookup"><span data-stu-id="534c3-102">[Office](Office.md).context</span></span>

<span data-ttu-id="534c3-p101">Der Office.context-Namespace stellt gemeinsam genutzte Oberflächen bereit, die von Add-Ins in allen Office-Apps verwendet werden. Diese Auflistung dokumentiert nur die Schnittstellen, die von Outlook-Add-Ins verwendet werden. Eine vollständige Auflistung des Office.context-Namespaces finden Sie in der [Office.context-Referenz in der freigegebenen API](/javascript/api/office/office.context).</span><span class="sxs-lookup"><span data-stu-id="534c3-p101">The Office.context namespace provides shared interfaces that are used by add-ins in all of the Office apps. This listing documents only those interfaces that are used by Outlook add-ins. For a full listing of the Office.context namespace, see the [Office.context reference in the Shared API](/javascript/api/office/office.context).</span></span>

##### <a name="requirements"></a><span data-ttu-id="534c3-105">Anforderungen</span><span class="sxs-lookup"><span data-stu-id="534c3-105">Requirements</span></span>

|<span data-ttu-id="534c3-106">Anforderung</span><span class="sxs-lookup"><span data-stu-id="534c3-106">Requirement</span></span>| <span data-ttu-id="534c3-107">Wert</span><span class="sxs-lookup"><span data-stu-id="534c3-107">Value</span></span>|
|---|---|
|[<span data-ttu-id="534c3-108">Mindestversion des Postfachanforderungssatzes</span><span class="sxs-lookup"><span data-stu-id="534c3-108">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="534c3-109">1.0</span><span class="sxs-lookup"><span data-stu-id="534c3-109">1.0</span></span>|
|[<span data-ttu-id="534c3-110">Zutreffender Outlook-Modus</span><span class="sxs-lookup"><span data-stu-id="534c3-110">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="534c3-111">Verfassen oder Lesen</span><span class="sxs-lookup"><span data-stu-id="534c3-111">Compose or read</span></span>|

##### <a name="members-and-methods"></a><span data-ttu-id="534c3-112">Elemente und Methoden</span><span class="sxs-lookup"><span data-stu-id="534c3-112">Members and methods</span></span>

| <span data-ttu-id="534c3-113">Element</span><span class="sxs-lookup"><span data-stu-id="534c3-113">Member</span></span> | <span data-ttu-id="534c3-114">Typ</span><span class="sxs-lookup"><span data-stu-id="534c3-114">Type</span></span> |
|--------|------|
| [<span data-ttu-id="534c3-115">displayLanguage</span><span class="sxs-lookup"><span data-stu-id="534c3-115">displayLanguage</span></span>](#displaylanguage-string) | <span data-ttu-id="534c3-116">Element</span><span class="sxs-lookup"><span data-stu-id="534c3-116">Member</span></span> |
| [<span data-ttu-id="534c3-117">officeTheme</span><span class="sxs-lookup"><span data-stu-id="534c3-117">officeTheme</span></span>](#officetheme-object) | <span data-ttu-id="534c3-118">Element</span><span class="sxs-lookup"><span data-stu-id="534c3-118">Member</span></span> |
| [<span data-ttu-id="534c3-119">roamingSettings</span><span class="sxs-lookup"><span data-stu-id="534c3-119">roamingSettings</span></span>](#roamingsettings-roamingsettingsjavascriptapioutlook17officeroamingsettings) | <span data-ttu-id="534c3-120">Element</span><span class="sxs-lookup"><span data-stu-id="534c3-120">Member</span></span> |

### <a name="namespaces"></a><span data-ttu-id="534c3-121">Namespaces</span><span class="sxs-lookup"><span data-stu-id="534c3-121">Namespaces</span></span>

<span data-ttu-id="534c3-122">[Mailbox](office.context.mailbox.md): ermöglicht den Zugriff auf das Objektmodell von Outlook-add-in für Microsoft Outlook und Microsoft Outlook im Web.</span><span class="sxs-lookup"><span data-stu-id="534c3-122">[mailbox](office.context.mailbox.md): Provides access to the Outlook add-in object model for Microsoft Outlook and Microsoft Outlook on the web.</span></span>

### <a name="members"></a><span data-ttu-id="534c3-123">Elemente</span><span class="sxs-lookup"><span data-stu-id="534c3-123">Members</span></span>

####  <a name="displaylanguage-string"></a><span data-ttu-id="534c3-124">displayLanguage :String</span><span class="sxs-lookup"><span data-stu-id="534c3-124">displayLanguage :String</span></span>

<span data-ttu-id="534c3-125">Ruft das vom Benutzer im Sprachtagformat RFC 1766 angegebene Gebietsschema (Sprache) für die Benutzeroberfläche der Office-Hostanwendung ab.</span><span class="sxs-lookup"><span data-stu-id="534c3-125">Gets the locale (language) in RFC 1766 Language tag format specified by the user for the UI of the Office host application.</span></span>

<span data-ttu-id="534c3-126">Der `displayLanguage`-Wert gibt die aktuelle Einstellung **Anzeigesprache** wieder, die mit **Datei > Optionen > Sprache** in der Office-Hostanwendung angegeben wurde.</span><span class="sxs-lookup"><span data-stu-id="534c3-126">The `displayLanguage` value reflects the current **Display Language** setting specified with **File > Options > Language** in the Office host application.</span></span>

##### <a name="type"></a><span data-ttu-id="534c3-127">Typ:</span><span class="sxs-lookup"><span data-stu-id="534c3-127">Type:</span></span>

*   <span data-ttu-id="534c3-128">String</span><span class="sxs-lookup"><span data-stu-id="534c3-128">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="534c3-129">Anforderungen</span><span class="sxs-lookup"><span data-stu-id="534c3-129">Requirements</span></span>

|<span data-ttu-id="534c3-130">Anforderung</span><span class="sxs-lookup"><span data-stu-id="534c3-130">Requirement</span></span>| <span data-ttu-id="534c3-131">Wert</span><span class="sxs-lookup"><span data-stu-id="534c3-131">Value</span></span>|
|---|---|
|[<span data-ttu-id="534c3-132">Mindestversion des Postfachanforderungssatzes</span><span class="sxs-lookup"><span data-stu-id="534c3-132">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="534c3-133">1.0</span><span class="sxs-lookup"><span data-stu-id="534c3-133">1.0</span></span>|
|[<span data-ttu-id="534c3-134">Zutreffender Outlook-Modus</span><span class="sxs-lookup"><span data-stu-id="534c3-134">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="534c3-135">Verfassen oder Lesen</span><span class="sxs-lookup"><span data-stu-id="534c3-135">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="534c3-136">Beispiel</span><span class="sxs-lookup"><span data-stu-id="534c3-136">Example</span></span>

```js
function sayHelloWithDisplayLanguage() {
  var myDisplayLanguage = Office.context.displayLanguage;
  switch (myDisplayLanguage) {
    case 'en-US':
      write('Hello!');
      break;
    case 'en-NZ':
      write('G\'day mate!');
      break;
  }
}
// Function that writes to a div with id='message' on the page.
function write(message){
  document.getElementById('message').innerText += message;
}
```

####  <a name="officetheme-object"></a><span data-ttu-id="534c3-137">officeTheme :Object</span><span class="sxs-lookup"><span data-stu-id="534c3-137">officeTheme :Object</span></span>

<span data-ttu-id="534c3-138">Bietet Zugriff auf die Eigenschaften für Office-Farbdesigns.</span><span class="sxs-lookup"><span data-stu-id="534c3-138">Provides access to the properties for Office theme colors.</span></span>

> [!NOTE]
> <span data-ttu-id="534c3-139">Dieser Member wird in Outlook für iOS oder Outlook für Android nicht unterstützt.</span><span class="sxs-lookup"><span data-stu-id="534c3-139">This member is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="534c3-p102">Mit Office-Farbdesigns können Sie das Farbschema Ihres Add-Ins mit dem aktuellen Office-Design, das vom Benutzer mit **Datei > Office-Konto > Office-Design-UI** ausgewählt wurde, koordinieren. Das Farbschema wird auf alle Office-Hostanwendungen angewendet. Das Verwenden von Office-Designfarben eignet sich für E-Mail- und Aufgabenbereichs-Add-Ins.</span><span class="sxs-lookup"><span data-stu-id="534c3-p102">Using Office theme colors let's you coordinate the color scheme of your add-in with the current Office theme selected by the user with **File > Office Account > Office Theme UI**, which is applied across all Office host applications. Using Office theme colors is appropriate for mail and task pane add-ins.</span></span>

##### <a name="type"></a><span data-ttu-id="534c3-142">Typ:</span><span class="sxs-lookup"><span data-stu-id="534c3-142">Type:</span></span>

*   <span data-ttu-id="534c3-143">Object</span><span class="sxs-lookup"><span data-stu-id="534c3-143">Object</span></span>

##### <a name="properties"></a><span data-ttu-id="534c3-144">Eigenschaften:</span><span class="sxs-lookup"><span data-stu-id="534c3-144">Properties:</span></span>

|<span data-ttu-id="534c3-145">Name</span><span class="sxs-lookup"><span data-stu-id="534c3-145">Name</span></span>| <span data-ttu-id="534c3-146">Typ</span><span class="sxs-lookup"><span data-stu-id="534c3-146">Type</span></span>| <span data-ttu-id="534c3-147">Beschreibung</span><span class="sxs-lookup"><span data-stu-id="534c3-147">Description</span></span>|
|---|---|---|
|`bodyBackgroundColor`| <span data-ttu-id="534c3-148">Zeichenfolge</span><span class="sxs-lookup"><span data-stu-id="534c3-148">String</span></span>|<span data-ttu-id="534c3-149">Ruft die Hintergrundfarbe des Office-Designkörpers als hexadezimales Farbtriplet ab.</span><span class="sxs-lookup"><span data-stu-id="534c3-149">Gets the Office theme body background color as a hexadecimal color triplet.</span></span>|
|`bodyForegroundColor`| <span data-ttu-id="534c3-150">Zeichenfolge</span><span class="sxs-lookup"><span data-stu-id="534c3-150">String</span></span>|<span data-ttu-id="534c3-151">Ruft die Vordergrundfarbe des Office-Designkörpers als hexadezimales Farbtriplet ab.</span><span class="sxs-lookup"><span data-stu-id="534c3-151">Gets the Office theme body foreground color as a hexadecimal color triplet.</span></span>|
|`controlBackgroundColor`| <span data-ttu-id="534c3-152">Zeichenfolge</span><span class="sxs-lookup"><span data-stu-id="534c3-152">String</span></span>|<span data-ttu-id="534c3-153">Ruft die Hintergrundfarbe des Office-Designsteuerelements als hexadezimales Farbtriplet ab.</span><span class="sxs-lookup"><span data-stu-id="534c3-153">Gets the Office theme control background color as a hexadecimal color triplet.</span></span>|
|`controlForegroundColor`| <span data-ttu-id="534c3-154">Zeichenfolge</span><span class="sxs-lookup"><span data-stu-id="534c3-154">String</span></span>|<span data-ttu-id="534c3-155">Ruft die Vordergrundfarbe des Office-Designsteuerelements als hexadezimales Farbtriplet ab.</span><span class="sxs-lookup"><span data-stu-id="534c3-155">Gets the Office theme body control color as a hexadecimal color triplet.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="534c3-156">Anforderungen</span><span class="sxs-lookup"><span data-stu-id="534c3-156">Requirements</span></span>

|<span data-ttu-id="534c3-157">Anforderung</span><span class="sxs-lookup"><span data-stu-id="534c3-157">Requirement</span></span>| <span data-ttu-id="534c3-158">Wert</span><span class="sxs-lookup"><span data-stu-id="534c3-158">Value</span></span>|
|---|---|
|[<span data-ttu-id="534c3-159">Mindestversion des Postfachanforderungssatzes</span><span class="sxs-lookup"><span data-stu-id="534c3-159">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="534c3-160">1.3</span><span class="sxs-lookup"><span data-stu-id="534c3-160">1.3</span></span>|
|[<span data-ttu-id="534c3-161">Zutreffender Outlook-Modus</span><span class="sxs-lookup"><span data-stu-id="534c3-161">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="534c3-162">Verfassen oder Lesen</span><span class="sxs-lookup"><span data-stu-id="534c3-162">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="534c3-163">Beispiel</span><span class="sxs-lookup"><span data-stu-id="534c3-163">Example</span></span>

```js
function applyOfficeTheme(){
  // Get office theme colors.
  var bodyBackgroundColor = Office.context.officeTheme.bodyBackgroundColor;
  var bodyForegroundColor = Office.context.officeTheme.bodyForegroundColor;
  var controlBackgroundColor = Office.context.officeTheme.controlBackgroundColor
  var controlForegroundColor = Office.context.officeTheme.controlForegroundColor;

  // Apply body background color to a CSS class.
  $('.body').css('background-color', bodyBackgroundColor);
}
```

####  <a name="roamingsettings-roamingsettingsjavascriptapioutlook17officeroamingsettings"></a><span data-ttu-id="534c3-164">roamingSettings :[RoamingSettings](/javascript/api/outlook_1_7/office.RoamingSettings)</span><span class="sxs-lookup"><span data-stu-id="534c3-164">roamingSettings :[RoamingSettings](/javascript/api/outlook_1_7/office.RoamingSettings)</span></span>

<span data-ttu-id="534c3-165">Ruft ein Objekt ab, das die benutzerdefinierten Einstellungen oder den Status eines Mail-Add-Ins im Postfach eines Benutzers darstellt.</span><span class="sxs-lookup"><span data-stu-id="534c3-165">Gets an object that represents the custom settings or state of a mail add-in saved to a user's mailbox.</span></span>

<span data-ttu-id="534c3-166">Mit dem `RoamingSettings`-Objekt können Sie Daten für ein Mail-Add-In speichern und darauf zugreifen, die im Postfach eines Benutzers gespeichert sind, damit diese Daten für das Add-In verfügbar sind, wenn es über eine beliebige Hostclientanwendung zum Zugreifen auf das Postfach ausgeführt wird.</span><span class="sxs-lookup"><span data-stu-id="534c3-166">The `RoamingSettings` object lets you store and access data for a mail add-in that is stored in a user's mailbox, so that is available to that add-in when it is running from any host client application used to access that mailbox.</span></span>

##### <a name="type"></a><span data-ttu-id="534c3-167">Typ:</span><span class="sxs-lookup"><span data-stu-id="534c3-167">Type:</span></span>

*   [<span data-ttu-id="534c3-168">RoamingSettings</span><span class="sxs-lookup"><span data-stu-id="534c3-168">RoamingSettings</span></span>](/javascript/api/outlook_1_7/office.RoamingSettings)

##### <a name="requirements"></a><span data-ttu-id="534c3-169">Anforderungen</span><span class="sxs-lookup"><span data-stu-id="534c3-169">Requirements</span></span>

|<span data-ttu-id="534c3-170">Anforderung</span><span class="sxs-lookup"><span data-stu-id="534c3-170">Requirement</span></span>| <span data-ttu-id="534c3-171">Wert</span><span class="sxs-lookup"><span data-stu-id="534c3-171">Value</span></span>|
|---|---|
|[<span data-ttu-id="534c3-172">Mindestversion des Postfachanforderungssatzes</span><span class="sxs-lookup"><span data-stu-id="534c3-172">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="534c3-173">1.0</span><span class="sxs-lookup"><span data-stu-id="534c3-173">1.0</span></span>|
|[<span data-ttu-id="534c3-174">Mindestberechtigungsstufe</span><span class="sxs-lookup"><span data-stu-id="534c3-174">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="534c3-175">Eingeschränkt</span><span class="sxs-lookup"><span data-stu-id="534c3-175">Restricted</span></span>|
|[<span data-ttu-id="534c3-176">Zutreffender Outlook-Modus</span><span class="sxs-lookup"><span data-stu-id="534c3-176">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="534c3-177">Verfassen oder Lesen</span><span class="sxs-lookup"><span data-stu-id="534c3-177">Compose or read</span></span>|