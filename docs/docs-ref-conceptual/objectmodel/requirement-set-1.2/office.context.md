
# <a name="context"></a>context

### [Office](Office.md).context

Der Office.context-Namespace stellt gemeinsam genutzte Oberflächen bereit, die von Add-Ins in allen Office-Apps verwendet werden. Diese Auflistung dokumentiert nur die Schnittstellen, die von Outlook-Add-Ins verwendet werden. Eine vollständige Auflistung des Office.context-Namespaces finden Sie in der [Office.context-Referenz in der freigegebenen API](/javascript/api/office/office.context).


##### <a name="requirements"></a>Anforderungen

|Anforderung| Wert|
|---|---|
|[Mindestversion des Postfachanforderungssatzes](/javascript/office/requirement-sets/outlook-api-requirement-sets)| 1.0|
|[Zutreffender Outlook-Modus](https://docs.microsoft.com/outlook/add-ins/#extension-points)| Verfassen oder Lesen|

### <a name="namespaces"></a>Namespaces

[Postfach](office.context.mailbox.md) – ermöglicht den Zugriff auf das Objektmodell von Outlook-add-in für Microsoft Outlook und Microsoft Outlook im Web.

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

####  <a name="roamingsettings-roamingsettingsjavascriptapioutlook12officeroamingsettings"></a>roamingSettings :[RoamingSettings](/javascript/api/outlook_1_2/office.RoamingSettings)

Ruft ein Objekt ab, das die benutzerdefinierten Einstellungen oder den Status eines Mail-Add-Ins im Postfach eines Benutzers darstellt.

Mit dem `RoamingSettings`-Objekt können Sie Daten für ein Mail-Add-In speichern und darauf zugreifen, die im Postfach eines Benutzers gespeichert sind, damit diese Daten für das Add-In verfügbar sind, wenn es über eine beliebige Hostclientanwendung zum Zugreifen auf das Postfach ausgeführt wird.

##### <a name="type"></a>Typ:

*   [RoamingSettings](/javascript/api/outlook_1_2/office.RoamingSettings)

##### <a name="requirements"></a>Anforderungen

|Anforderung| Wert|
|---|---|
|[Mindestversion des Postfachanforderungssatzes](/javascript/office/requirement-sets/outlook-api-requirement-sets)| 1.0|
|[Mindestberechtigungsstufe](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| Eingeschränkt|
|[Zutreffender Outlook-Modus](https://docs.microsoft.com/outlook/add-ins/#extension-points)| Verfassen oder Lesen|