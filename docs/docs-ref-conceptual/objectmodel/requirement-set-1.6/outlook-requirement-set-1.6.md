# <a name="outlook-add-in-api-requirement-set-16"></a>Outlook-add-in-API-Anforderung festlegen 1.6

Die Outlook-add-in-API-Teilmenge der JavaScript-API für Office umfasst Objekte, Methoden, Eigenschaften und Ereignisse für die Verwendung in einer Outlook-add-Ins.

## <a name="whats-new-in-16"></a>Was ist neu in 1.6?

Anforderungssatz 1.6 enthält alle Features von [Anforderung 1,5 festgelegt](../requirement-set-1.5/outlook-requirement-set-1.5.md). Er hinzugefügt die folgenden Features.

- Hinzugefügte neue APIs für kontextbezogene-Add-ins abzurufenden Entität oder regulären Ausdruck übereinstimmen, dass der Benutzer ausgewählt werden, um das Add-in zu aktivieren.
- Eine neue API zum Öffnen ein Formulars für neue Nachrichten hinzugefügt.
- Die Möglichkeit für das Add-in zu bestimmen, die Kontenart an, der das Postfach des Benutzers hinzugefügt.

### <a name="change-log"></a>Änderungsprotokoll

- [Office.context.mailbox.item.getSelectedEntities](office.context.mailbox.item.md#getselectedentities--entitiesjavascriptapioutlook16officeentities)hinzugefügt: Fügt eine neue Funktion, die in eine hervorgehobene Übereinstimmung mit dem ein Benutzer ausgewählt hat gefundenen Entitäten abruft. Hervorgehobene Übereinstimmungen gelten für Kontext-Add-Ins.
- [Office.context.mailbox.item.getSelectedRegExMatches](office.context.mailbox.item.md#getselectedregexmatches--object)hinzugefügt: Fügt eine neue Funktion, die gibt Zeichenfolgenwerte zurück, in einem hervorgehobenen übereinstimmen, die die in der XML-Manifestdatei definierten regulären Ausdrücken entsprechen. Hervorgehobene Übereinstimmungen gelten für Kontext-Add-Ins.
- [Office.context.mailbox.displayNewMessageForm](office.context.mailbox.md#displaynewmessageformparameters)hinzugefügt: Fügt eine neue Funktion, die ein neues Nachrichtenformular öffnet.
- [Office.context.mailbox.userProfile.accountType](office.context.mailbox.userprofile.md#accounttype-string)hinzugefügt: Fügt ein neues Element in das Benutzerprofil, der die Art des Kontos des Benutzers angibt.

## <a name="see-also"></a>Siehe auch

- [Outlook-Add-Ins](https://docs.microsoft.com/outlook/add-ins/)
- [Codebeispiele zu Outlook-Add-Ins](https://developer.microsoft.com/outlook/gallery/?filterBy=Outlook,Samples,Add-ins)
- [Erste Schritte](https://docs.microsoft.com/outlook/add-ins/quick-start)