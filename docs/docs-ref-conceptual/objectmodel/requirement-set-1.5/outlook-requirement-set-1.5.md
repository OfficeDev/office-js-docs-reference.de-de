# <a name="outlook-add-in-api-requirement-set-15"></a>Outlook-Add-In-API-Anforderungssatz 1.5

Die Outlook-add-in-API-Teilmenge der JavaScript-API für Office umfasst Objekte, Methoden, Eigenschaften und Ereignisse für die Verwendung in einer Outlook-add-Ins.

> [!NOTE]
> Diese Dokumentation ist für eine [Anforderung festgelegt](/javascript/office/requirement-sets/outlook-api-requirement-sets) als die aktuelle Anforderungssatz.

## <a name="whats-new-in-15"></a>Neu in Version 1.5

Der Anforderungssatz 1.5 umfasst alle Funktionen von [Anforderungssatz 1.4](../requirement-set-1.4/outlook-requirement-set-1.4.md). Zusätzlich wurden die folgenden Funktionen implementiert:

- Es wurde Unterstützung für [anheftbare Aufgabenbereiche](https://docs.microsoft.com/outlook/add-ins/pinnable-taskpane) hinzugefügt.
- Es wurde Unterstützung für den Aufruf von [REST-APIs](https://docs.microsoft.com/outlook/add-ins/use-rest-api) hinzugefügt.
- Es wurde eine Funktion hinzugefügt, über die sich Anlagen als inline markieren lassen.
- Es wurde eine Funktion zum Schließen von Aufgabenbereichen und Dialogfeldern hinzugefügt.

### <a name="change-log"></a>Änderungsprotokoll

- [Office.context.mailbox.addHandlerAsync](office.context.mailbox.md#addhandlerasynceventtype-handler-options-callback) hinzugefügt: Fügt einen Ereignishandler für ein unterstütztes Ereignis hinzu.
- [Office.EventType](office.md#eventtype-string)hinzugefügt: Gibt an, das ein Ereignishandler zugeordnete Ereignis und bietet Unterstützung für ItemChanged-Ereignis.
- [Office.context.mailbox.restUrl](office.context.mailbox.md#resturl-string) hinzugefügt: Ruft die URL des REST-Endpunkts für das betreffende E-Mail-Konto ab.
- [Office.context.mailbox.getCallbackTokenAsync](office.context.mailbox.md#getcallbacktokenasyncoptions-callback) geändert: Es wurde eine neue Version dieser Methode mit einer neuen Signatur (`getCallbackTokenAsync([options], callback)`) hinzugefügt. Die ursprüngliche Version steht weiterhin in unveränderter Form zur Verfügung.
- Die [Office.context.ui.closeContainer](/javascript/api/office/office.ui#closecontainer--)hinzugefügt.
- [Office.context.mailbox.item.addFileAttachmentAsync](office.context.mailbox.item.md#addfileattachmentasyncuri-attachmentname-options-callback) geändert: Es wurde ein neuer Wert namens `isInline` im `options`-Wörterbuch hinzugefügt, über den festgelegt wird, dass ein Bild inline im Nachrichtentext verwendet werden soll.
- [Office.context.mailbox.item.displayReplyAllForm](office.context.mailbox.item.md#displayreplyallformformdata) geändert: Es wurde ein neuer Wert namens `isInline` im `formData.attachments`-Wörterbuch hinzugefügt, über den festgelegt wird, dass ein Bild inline im Nachrichtentext verwendet werden soll.
- [Office.context.mailbox.item.displayReplyForm](office.context.mailbox.item.md#displayreplyformformdata) geändert: Es wurde ein neuer Wert namens `isInline` im `formData.attachments`-Wörterbuch hinzugefügt, über den festgelegt wird, dass ein Bild inline im Nachrichtentext verwendet werden soll.

## <a name="see-also"></a>Siehe auch

- [Outlook-Add-Ins](https://docs.microsoft.com/outlook/add-ins/)
- [Codebeispiele zu Outlook-Add-Ins](https://developer.microsoft.com/outlook/gallery/?filterBy=Outlook,Samples,Add-ins)
- [Erste Schritte](https://docs.microsoft.com/outlook/add-ins/quick-start)