# <a name="outlook-add-in-api-requirement-set-11"></a>Outlook-Add-In-API-Anforderungssatz 1.1

Die Outlook-Add-In-API-Teilmenge der JavaScript-API für Office umfasst Objekte, Methoden, Eigenschaften und Ereignisse, die Sie in Outlook-Add-Ins verwenden können.

> [!NOTE]
> Diese Dokumentation ist für eine [Anforderung festgelegt](/javascript/office/requirement-sets/outlook-api-requirement-sets) als die aktuelle Anforderungssatz. 

## <a name="whats-new-in-11"></a>Neu in Version 1.1

Der Anforderungssatz 1.1 umfasst alle Funktionen von Anforderungssatz 1.0. Zusätzlich hat er eine Funktion implementiert, über die Add-Ins auf den Text von Nachrichten und Terminen zugreifen können, und eine Funktion, über die das aktuelle Element geändert werden kann.

### <a name="change-log"></a>Änderungsprotokoll

- Objekt [Body](/javascript/api/outlook_1_1/office.body) hinzugefügt: Stellt Methoden zum Hinzufügen und Aktualisieren des Inhalts eines Elements in einem Outlook-Add-In bereit.
- Objekt [Location](/javascript/api/outlook_1_1/office.location) hinzugefügt: Stellt Methoden zum Abrufen und Festlegen des Orts einer Besprechung in einem Outlook-Add-In bereit.
- Objekt [Recipients](/javascript/api/outlook_1_1/office.recipients) hinzugefügt: Stellt Methoden zum Abrufen und Festlegen der Empfänger eines Termins oder einer Nachricht in einem Outlook-Add-In bereit.
- Objekt [Subject](/javascript/api/outlook_1_1/office.subject) hinzugefügt: Stellt Methoden zum Abrufen und Festlegen des Betreffs eines Termins oder einer Nachricht in einem Outlook-Add-In bereit.
- Objekt [Time](/javascript/api/outlook_1_1/office.time) hinzugefügt: Stellt Methoden zum Abrufen und Festlegen der Startzeit oder der Endzeit einer Besprechung in einem Outlook-Add-In bereit.
- [Office.context.mailbox.item.addFileAttachmentAsync](office.context.mailbox.item.md#addfileattachmentasyncuri-attachmentname-options-callback) hinzugefügt: Fügt einer Nachricht oder einem Termin eine Datei als Anlage hinzu.
- [Office.context.mailbox.item.addItemAttachmentAsync](office.context.mailbox.item.md#additemattachmentasyncitemid-attachmentname-options-callback) hinzugefügt: Fügt einer Nachricht oder einem Termin ein Exchange-Element ans Anlage hinzu, z. B. eine Nachricht.
- [Office.context.mailbox.item.removeAttachmentAsync](office.context.mailbox.item.md#removeattachmentasyncattachmentid-options-callback) hinzugefügt: Entfernt Anlagen aus einer Nachricht oder einem Termin.
- [Office.context.mailbox.item.body](office.context.mailbox.item.md#body-bodyjavascriptapioutlook11officebody) hinzugefügt: Ruft ein Objekt ab, das Methoden zum Bearbeiten des Texts eines Elements bereitstellt.
- [Office.context.mailbox.item.bcc](office.context.mailbox.item.md#bcc-recipientsjavascriptapioutlook11officerecipients) hinzugefügt: Ruft die Empfänger in der Bcc-Zeile (Blind Carbon Copy) einer Nachricht ab, oder legt sie fest.
- [Office.MailboxEnums.RecipientType](/javascript/api/outlook_1_1/office.mailboxenums.recipienttype) hinzugefügt: Gibt den Empfängertyp für einen Termin an.

## <a name="see-also"></a>Siehe auch

- [Outlook-Add-Ins](https://docs.microsoft.com/outlook/add-ins/)
- [Codebeispiele zu Outlook-Add-Ins](https://developer.microsoft.com/outlook/gallery/?filterBy=Outlook,Samples,Add-ins)
- [Erste Schritte](https://docs.microsoft.com/outlook/add-ins/quick-start)