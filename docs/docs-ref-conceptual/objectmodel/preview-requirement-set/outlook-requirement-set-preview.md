# <a name="outlook-add-in-api-preview-requirement-set"></a>Preview-Anforderungssatz für die Outlook-Add-In-API

Die Outlook-Add-In-API-Teilmenge der JavaScript-API für Office umfasst Objekte, Methoden, Eigenschaften und Ereignisse, die Sie in Outlook-Add-Ins verwenden können.

> [!NOTE]
> Diese Dokumentation ist für eine **Vorschau** [Anforderung festgelegt](/javascript/office/requirement-sets/outlook-api-requirement-sets). Dieser Anforderungssatz ist noch nicht vollständig implementiert. Daher melden Clients die Unterstützung für diesen Anforderungssatz nicht korrekt. Sie sollten diesen Anforderungssatz nicht im Manifest Ihres Add-Ins angeben. Methoden und Eigenschaften, die in diesem Anforderungssatz eingeführt werden, sollten vor der Verwendung einzeln auf ihre Verfügbarkeit getestet werden.

Die Preview-Anforderungssatz enthält alle Features von [Anforderung 1.6 festgelegt](../requirement-set-1.6/outlook-requirement-set-1.6.md).

## <a name="features-in-preview"></a>Preview-Funktionen

Die folgenden Funktionen haben aktuell noch Preview-Status:

- [Von](/javascript/api/outlook/office.from) - hinzugefügt ein neues Objekt, das bietet eine Methode zum Abrufen der vom Wert.
- [Organizer](/javascript/api/outlook/office.organizer) - hinzugefügt, ein neues Objekt, das eine Methode zum Abrufen des Werts der Organisator bereitstellt.
- [Serie](/javascript/api/outlook/office.recurrence) - hinzugefügt, ein neues Objekt, das Methoden zum Abrufen und festlegen das Serienmuster von Terminen jedoch nur das Serienmuster von Nachrichten, die Anfragen meeting sind get bereitstellt.
- [SeriesTime](/javascript/api/outlook/office.seriestime) - hinzugefügt, ein neues Objekt, das Methoden zum Abrufen und Festlegen von Datum und Uhrzeit der Termine in einer Terminserie und zum Abrufen der Daten und Uhrzeiten von Besprechungsanfragen in einer Terminserie bereitstellt.
- [Event.Completed](/javascript/api/office/office.addincommands.event#completed-options-): ein neuer optionaler Parameter `options`, bei dem es sich um ein Wörterbuch mit einem einzigen gültigen Wert handelt, und zwar `allowEvent`. Dieser Wert wird verwendet, um die Ausführung eines Ereignisses abzubrechen.
- [Office.context.mailbox.item.addFileAttachmentFromBase64Async](office.context.mailbox.item.md#addfileattachmentfrombase64asyncbase64file-attachmentname-options-callback) - hinzugefügt, eine neue Methode, die eine Datei von base64-Codierung auf einer Nachricht oder eines Termins anfügt.
- [Office.context.mailbox.item.addHandlerAsync](office.context.mailbox.item.md#addhandlerasynceventtype-handler-options-callback) - hinzugefügt, eine neue Methode, die einen Ereignishandler für ein Ereignis unterstützte hinzufügt.
- [Office.context.mailbox.item.from](office.context.mailbox.item.md#from-emailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsfromjavascriptapioutlookofficefrom) - geändert, zum Abrufen der Wert im Modus "Verfassen".
- [Office.context.mailbox.item.getInitializationContextAsync](office.context.mailbox.item.md#getinitializationcontextasyncoptions-callback) – Neue Funktion hinzugefügt, die Initialisierungsdaten zurückgibt, die übergeben werden, wenn das Add-In [durch eine Nachricht mit Aktionen aktiviert](https://docs.microsoft.com/outlook/actionable-messages/invoke-add-in-from-actionable-message) wird.
- [Office.context.mailbox.item.organizer](office.context.mailbox.item.md#organizer-emailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsorganizerjavascriptapioutlookofficeorganizer) - so geändert, dass der Organisator Wert im Modus "Verfassen" abgerufen.
- [Office.context.mailbox.item.recurrence](office.context.mailbox.item.md#nullable-recurrence-recurrencejavascriptapioutlookofficerecurrence) - hinzugefügt, eine neue Eigenschaft, die Ruft ab oder legt ein Objekt, das Methoden zum Verwalten von des Serienmusters der ein Terminelement bereitstellt. Diese Eigenschaft kann auch verwendet werden, um das Serienmuster einer Besprechung abzurufen Element Anforderung.
- [Office.context.mailbox.item.removeHandlerAsync](office.context.mailbox.item.md#removehandlerasynceventtype-handler-options-callback) - hinzugefügt, eine neue Methode, die einen Ereignishandler entfernt wird.
- [Office.context.mailbox.item.seriesId](office.context.mailbox.item.md#nullable-seriesid-string) - hinzugefügt, dass eine neue Eigenschaft, die die Id der Datenreihe Auftreten eines Serientermins versehen gehört.
- [Office.context.auth.getAccessTokenAsync](https://docs.microsoft.com/office/dev/add-ins/develop/sso-in-office-add-ins#sso-api-reference) – Zugriff auf `getAccessTokenAsync` wurde hinzugefügt, wodurch Add-Ins [ein Zugriffstoken für die Microsoft Graph-API erhalten](https://docs.microsoft.com/outlook/add-ins/authenticate-a-user-with-an-sso-token).
- [Office.MailboxEnums.Days](/javascript/api/outlook/office.mailboxenums.days) - hinzugefügt, eine neue Enumeration, die den Tag der Woche oder Typ des Tages angibt.
- [Office.MailboxEnums.Month](/javascript/api/outlook/office.mailboxenums.month) - hinzugefügt, eine neue Enumeration, die den Monat angibt.
- [Office.MailboxEnums.RecurrenceTimeZone](/javascript/api/outlook/office.mailboxenums.recurrencetimezone) - hinzugefügt, eine neue Enumeration, die angibt, die Zeitzone auf die Serie angewendet wird.
- [Office.MailboxEnums.RecurrenceType](/javascript/api/outlook/office.mailboxenums.recurrencetype) - hinzugefügt, eine neue Enumeration, die den Typ des Serienmusters angibt.
- [Office.MailboxEnums.WeekNumber](/javascript/api/outlook/office.mailboxenums.weeknumber) - hinzugefügt, eine neue Enumeration, die die Woche des Monats angibt.
- [Office.EventType](/javascript/api/office/office.eventtype) - geändert zur Unterstützung von RecurrenceChanged, RecipientsChanged, AppointmentTimeChanged und OfficeThemeChanged Ereignisse durch Hinzufügen von `RecurrencePatternChanged`, `RecipientsChanged`, `AppointmentTimeChanged`, und `OfficeThemeChanged` Einträge fest.

## <a name="see-also"></a>Siehe auch

- [Outlook-Add-Ins](https://docs.microsoft.com/outlook/add-ins/)
- [Codebeispiele zu Outlook-Add-Ins](https://developer.microsoft.com/outlook/gallery/?filterBy=Outlook,Samples,Add-ins)
- [Erste Schritte](https://docs.microsoft.com/outlook/add-ins/quick-start)