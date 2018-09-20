
# <a name="diagnostics"></a>diagnostics

### [Office](Office.md)[.context](Office.context.md)[.mailbox](Office.context.mailbox.md). diagnostics

Stellt einem Outlook-Add-In Diagnoseinformationen bereit.

##### <a name="requirements"></a>Anforderungen

|Anforderung| Wert|
|---|---|
|[Mindestversion des Postfachanforderungssatzes](/javascript/office/requirement-sets/outlook-api-requirement-sets)| 1.0|
|[Mindestberechtigungsstufe](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| ReadItem|
|[Zutreffender Outlook-Modus](https://docs.microsoft.com/outlook/add-ins/#extension-points)| Verfassen oder Lesen|

##### <a name="members-and-methods"></a>Elemente und Methoden

| Element | Typ |
|--------|------|
| [hostName](#hostname-string) | Element |
| [hostVersion](#hostversion-string) | Element |
| [OWAView](#owaview-string) | Element |

### <a name="members"></a>Elemente

####  <a name="hostname-string"></a>hostName :String

Ruft eine Zeichenfolge ab, die den Namen der Hostanwendung darstellt.

Eine Zeichenfolge mit einem der folgenden Werte: `Outlook`, `Mac Outlook`, `OutlookIOS` oder `OutlookWebApp`.

##### <a name="type"></a>Typ:

*   String

##### <a name="requirements"></a>Anforderungen

|Anforderung| Wert|
|---|---|
|[Mindestversion des Postfachanforderungssatzes](/javascript/office/requirement-sets/outlook-api-requirement-sets)| 1.0|
|[Mindestberechtigungsstufe](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| ReadItem|
|[Zutreffender Outlook-Modus](https://docs.microsoft.com/outlook/add-ins/#extension-points)| Verfassen oder Lesen|

####  <a name="hostversion-string"></a>hostVersion :String

Ruft eine Zeichenfolge ab, die entweder die Version der Hostanwendung oder des Exchange-Servers darstellt.

Wenn das E-Mail-Add-In im Outlook-Desktopclient oder Outlook für iOS ausgeführt wird, gibt die `hostVersion`-Eigenschaft die Version der Hostanwendung zurück. In Outlook Web App gibt die Eigenschaft die Version von Exchange Server zurück. Ein Beispiel ist die Zeichenfolge `15.0.468.0`.

##### <a name="type"></a>Typ:

*   String

##### <a name="requirements"></a>Anforderungen

|Anforderung| Wert|
|---|---|
|[Mindestversion des Postfachanforderungssatzes](/javascript/office/requirement-sets/outlook-api-requirement-sets)| 1.0|
|[Mindestberechtigungsstufe](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| ReadItem|
|[Zutreffender Outlook-Modus](https://docs.microsoft.com/outlook/add-ins/#extension-points)| Verfassen oder Lesen|

####  <a name="owaview-string"></a>OWAView :String

Ruft eine Zeichenfolge ab, die die aktuelle Ansicht von Outlook Web App darstellt.

Die zurückgegebene Zeichenfolge kann einen der folgenden Werte haben: `OneColumn`, `TwoColumns` oder `ThreeColumns`.

Handelt es sich bei der Hostanwendung nicht um Outlook Web App, wird durch den Zugriff auf diese Eigenschaft `undefined` zurückgegeben.

Outlook Web App verfügt über drei Ansichten, die der Bildschirm- und Fensterbreite sowie mit der Anzahl der Spalten übereinstimmt, die angezeigt werden können:

*   `OneColumn` wird bei schmalem Bildschirm angezeigt. Outlook Web App verwendet dieses einspaltige Layout auf dem gesamten Bildschirm eines Smartphones.
*   `TwoColumns` wird bei etwas breiterem Bildschirm angezeigt. Outlook Web App verwendet diese Ansicht auf den meisten Tablet-PCs.
*   `ThreeColumns` wird bei breitem Bildschirm angezeigt. Outlook Web App verwendet diese Ansicht z. B. in einem Fenster mit voller Bildschirmgröße auf einem Desktopcomputer.

##### <a name="type"></a>Typ:

*   String

##### <a name="requirements"></a>Anforderungen

|Anforderung| Wert|
|---|---|
|[Mindestversion des Postfachanforderungssatzes](/javascript/office/requirement-sets/outlook-api-requirement-sets)| 1.0|
|[Mindestberechtigungsstufe](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| ReadItem|
|[Zutreffender Outlook-Modus](https://docs.microsoft.com/outlook/add-ins/#extension-points)| Verfassen oder Lesen|