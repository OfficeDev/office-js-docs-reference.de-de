# <a name="outlook-add-in-api-requirement-set-12"></a>Outlook-Add-In-API-Anforderungssatz 1.2

Die Outlook-Add-In-API-Teilmenge der JavaScript-API für Office umfasst Objekte, Methoden, Eigenschaften und Ereignisse, die Sie in Outlook-Add-Ins verwenden können.

> [!NOTE]
> Diese Dokumentation ist für eine [Anforderung festgelegt](/javascript/office/requirement-sets/outlook-api-requirement-sets) als die aktuelle Anforderungssatz. 

## <a name="whats-new-in-12"></a>Neu in Version 1.2

Der Anforderungssatz 1.2 umfasst alle Funktionen von [Anforderungssatz 1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md). Zusätzlich hat er eine Funktion implementiert, über die Add-Ins Text am Cursor des Benutzers einfügen können, entweder im Betreff oder im Text der Nachricht.

### <a name="change-log"></a>Änderungsprotokoll

- [Office.context.mailbox.item.getSelectedDataAsync](office.context.mailbox.item.md#getselecteddataasynccoerciontype-options-callback--string)hinzugefügt: Gibt die ausgewählte Daten asynchron aus dem Betreff oder Nachrichtentext zurück.
- [Office.context.mailbox.item.setSelectedDataAsync](office.context.mailbox.item.md#setselecteddataasyncdata-options-callback) hinzugefügt: Fügt asynchron Daten in den Text oder den Betreff einer Nachricht ein.
- [Office.context.mailbox.item.displayReplyAllForm](office.context.mailbox.item.md#displayreplyallformformdata) geändert: Dem Parameter `formData` wurde die Eigenschaft `attachments` hinzugefügt.
- [Office.context.mailbox.item.displayReplyForm](office.context.mailbox.item.md#displayreplyformformdata) geändert: Dem Parameter `formData` wurde die Eigenschaft `attachments` hinzugefügt.

## <a name="see-also"></a>Siehe auch

- [Outlook-Add-Ins](https://docs.microsoft.com/outlook/add-ins/)
- [Codebeispiele zu Outlook-Add-Ins](https://developer.microsoft.com/outlook/gallery/?filterBy=Outlook,Samples,Add-ins)
- [Erste Schritte](https://docs.microsoft.com/outlook/add-ins/quick-start)