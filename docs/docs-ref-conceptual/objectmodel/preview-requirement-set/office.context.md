
# <a name="context"></a>context

### <a name="officeofficemdcontext"></a>[Office](Office.md).context

Der Office.context-Namespace stellt gemeinsam genutzte Oberflächen bereit, die von Add-Ins in allen Office-Apps verwendet werden. Diese Auflistung dokumentiert nur die Schnittstellen, die von Outlook-Add-Ins verwendet werden. Eine vollständige Auflistung des Office.context-Namespaces finden Sie in der [Office.context-Referenz in der freigegebenen API](/javascript/api/office/office.context).

##### <a name="requirements"></a>Anforderungen

|Anforderung| Wert|
|---|---|
|[Mindestversion des Postfachanforderungssatzes](/javascript/office/requirement-sets/outlook-api-requirement-sets)| 1.0|
|[Zutreffender Outlook-Modus](https://docs.microsoft.com/outlook/add-ins/#extension-points)| Verfassen oder Lesen|

##### <a name="members-and-methods"></a>Elemente und Methoden

| Element | Typ |
|--------|------|
| [displayLanguage](#displaylanguage-string) | Element |
| [officeTheme](#officetheme-object) | Element |
| [roamingSettings](#roamingsettings-roamingsettingsjavascriptapioutlookofficeroamingsettings) | Element |

### <a name="namespaces"></a>Namespaces

[Mailbox](office.context.mailbox.md): ermöglicht den Zugriff auf das Objektmodell von Outlook-add-in für Microsoft Outlook und Microsoft Outlook im Web.

### <a name="members"></a>Elemente

####  <a name="displaylanguage-string"></a>displayLanguage :String

Ruft das vom Benutzer im Sprachtagformat RFC 1766 angegebene Gebietsschema (Sprache) für die Benutzeroberfläche der Office-Hostanwendung ab.

Der `displayLanguage`-Wert gibt die aktuelle Einstellung **Anzeigesprache** wieder, die mit **Datei > Optionen > Sprache** in der Office-Hostanwendung angegeben wurde.

##### <a name="type"></a>Typ:

*   String

##### <a name="requirements"></a>Anforderungen

|Anforderung| Wert|
|---|---|
|[Mindestversion des Postfachanforderungssatzes](/javascript/office/requirement-sets/outlook-api-requirement-sets)| 1.0|
|[Zutreffender Outlook-Modus](https://docs.microsoft.com/outlook/add-ins/#extension-points)| Verfassen oder Lesen|

##### <a name="example"></a>Beispiel

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

####  <a name="officetheme-object"></a>officeTheme :Object

Bietet Zugriff auf die Eigenschaften für Office-Farbdesigns.

> [!NOTE]
> Dieser Member wird in Outlook für iOS oder Outlook für Android nicht unterstützt.

Mit Office-Farbdesigns können Sie das Farbschema Ihres Add-Ins mit dem aktuellen Office-Design, das vom Benutzer mit **Datei > Office-Konto > Office-Design-UI** ausgewählt wurde, koordinieren. Das Farbschema wird auf alle Office-Hostanwendungen angewendet. Das Verwenden von Office-Designfarben eignet sich für E-Mail- und Aufgabenbereichs-Add-Ins.

##### <a name="type"></a>Typ:

*   Object

##### <a name="properties"></a>Eigenschaften:

|Name| Typ| Beschreibung|
|---|---|---|
|`bodyBackgroundColor`| Zeichenfolge|Ruft die Hintergrundfarbe des Office-Designkörpers als hexadezimales Farbtriplet ab.|
|`bodyForegroundColor`| Zeichenfolge|Ruft die Vordergrundfarbe des Office-Designkörpers als hexadezimales Farbtriplet ab.|
|`controlBackgroundColor`| Zeichenfolge|Ruft die Hintergrundfarbe des Office-Designsteuerelements als hexadezimales Farbtriplet ab.|
|`controlForegroundColor`| Zeichenfolge|Ruft die Vordergrundfarbe des Office-Designsteuerelements als hexadezimales Farbtriplet ab.|

##### <a name="requirements"></a>Anforderungen

|Anforderung| Wert|
|---|---|
|[Mindestversion des Postfachanforderungssatzes](/javascript/office/requirement-sets/outlook-api-requirement-sets)| 1.3|
|[Zutreffender Outlook-Modus](https://docs.microsoft.com/outlook/add-ins/#extension-points)| Verfassen oder Lesen|

##### <a name="example"></a>Beispiel

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

####  <a name="roamingsettings-roamingsettingsjavascriptapioutlookofficeroamingsettings"></a>roamingSettings :[RoamingSettings](/javascript/api/outlook/office.RoamingSettings)

Ruft ein Objekt ab, das die benutzerdefinierten Einstellungen oder den Status eines Mail-Add-Ins im Postfach eines Benutzers darstellt.

Mit dem `RoamingSettings`-Objekt können Sie Daten für ein Mail-Add-In speichern und darauf zugreifen, die im Postfach eines Benutzers gespeichert sind, damit diese Daten für das Add-In verfügbar sind, wenn es über eine beliebige Hostclientanwendung zum Zugreifen auf das Postfach ausgeführt wird.

##### <a name="type"></a>Typ:

*   [RoamingSettings](/javascript/api/outlook/office.RoamingSettings)

##### <a name="requirements"></a>Anforderungen

|Anforderung| Wert|
|---|---|
|[Mindestversion des Postfachanforderungssatzes](/javascript/office/requirement-sets/outlook-api-requirement-sets)| 1.0|
|[Mindestberechtigungsstufe](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| Eingeschränkt|
|[Zutreffender Outlook-Modus](https://docs.microsoft.com/outlook/add-ins/#extension-points)| Verfassen oder Lesen|