# <a name="outlook-add-in-api-preview-requirement-set"></a>Preview-Anforderungssatz für die Outlook-Add-In-API

Die Outlook-add-in-API-Teilmenge der JavaScript-API für Office umfasst Objekte, Methoden, Eigenschaften und Ereignisse für die Verwendung in einer Outlook-add-Ins.

> [!NOTE]
> Diese Dokumentation ist für eine **Vorschau** [Anforderung festgelegt](/javascript/office/requirement-sets/outlook-api-requirement-sets). Dieser Anforderungssatz ist noch nicht vollständig implementiert. Daher melden Clients die Unterstützung für diesen Anforderungssatz nicht korrekt. Sie sollten diesen Anforderungssatz nicht im Manifest Ihres Add-Ins angeben. Methoden und Eigenschaften, die in diesem Anforderungssatz eingeführt werden, sollten vor der Verwendung einzeln auf ihre Verfügbarkeit getestet werden.

Die Preview-Anforderungssatz enthält alle Features von [Anforderung 1.7 festgelegt](../requirement-set-1.7/outlook-requirement-set-1.7.md).

## <a name="features-in-preview"></a>Preview-Funktionen

Die folgenden Funktionen haben aktuell noch Preview-Status:

- [SharedProperties](/javascript/api/outlook/office.sharedproperties) - hinzugefügt, ein neues Objekt, das die Eigenschaften eines Termins oder einer Nachricht Elements in einem freigegebenen Ordner, Kalender oder Postfach darstellt.
- [Event.Completed](/javascript/api/office/office.addincommands.event#completed-options-): ein neuer optionaler Parameter `options`, bei dem es sich um ein Wörterbuch mit einem einzigen gültigen Wert handelt, und zwar `allowEvent`. Dieser Wert wird verwendet, um die Ausführung eines Ereignisses abzubrechen.
- [Office.context.mailbox.item.addFileAttachmentFromBase64Async](office.context.mailbox.item.md#addfileattachmentfrombase64asyncbase64file-attachmentname-options-callback) - hinzugefügt, eine neue Methode, die eine Datei von base64-Codierung auf einer Nachricht oder eines Termins anfügt.
- [Office.context.mailbox.item.getInitializationContextAsync](office.context.mailbox.item.md#getinitializationcontextasyncoptions-callback) – Neue Funktion hinzugefügt, die Initialisierungsdaten zurückgibt, die übergeben werden, wenn das Add-In [durch eine Nachricht mit Aktionen aktiviert](https://docs.microsoft.com/outlook/actionable-messages/invoke-add-in-from-actionable-message) wird.
- [Office.context.mailbox.item.getSharedPropertiesAsync](office.context.mailbox.item.md#getsharedpropertiesasyncoptions-callback) - hinzugefügt, eine neue Methode, die ein Objekt abruft, die die SharedProperties eines Termins oder einer Nachrichtenelement darstellt.
- [Office.context.auth.getAccessTokenAsync](https://docs.microsoft.com/office/dev/add-ins/develop/sso-in-office-add-ins#sso-api-reference) – Zugriff auf `getAccessTokenAsync` wurde hinzugefügt, wodurch Add-Ins [ein Zugriffstoken für die Microsoft Graph-API erhalten](https://docs.microsoft.com/outlook/add-ins/authenticate-a-user-with-an-sso-token).
- [Office.MailboxEnums.DelegatePermissions](/javascript/api/outlook/office.mailboxenums.delegatepermissions) - hinzugefügt, eine neue Bit Flag-Enumeration, die die Berechtigungen der Stellvertretung angibt.
- [Office.EventType](/javascript/api/office/office.eventtype) - geändert zur Unterstützung von OfficeThemeChanged Ereignis durch Hinzufügen von `OfficeThemeChanged` Eintrag.

## <a name="see-also"></a>Siehe auch

- [Outlook-Add-Ins](https://docs.microsoft.com/outlook/add-ins/)
- [Codebeispiele zu Outlook-Add-Ins](https://developer.microsoft.com/outlook/gallery/?filterBy=Outlook,Samples,Add-ins)
- [Erste Schritte](https://docs.microsoft.com/outlook/add-ins/quick-start)