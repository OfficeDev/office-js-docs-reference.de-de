# <a name="outlook-add-in-api-requirement-set-17"></a>Outlook-add-in-API-Anforderung festlegen 1.7

Die Outlook-Add-In-API-Teilmenge der JavaScript-API für Office umfasst Objekte, Methoden, Eigenschaften und Ereignisse, die Sie in Outlook-Add-Ins verwenden können.

## <a name="whats-new-in-17"></a>Was ist neu in 1.7?

Anforderungssatz 1.7 enthält alle Features von [Anforderung 1.6 festgelegt](../requirement-set-1.6/outlook-requirement-set-1.6.md). Er hinzugefügt die folgenden Features.

- Neue APIs über das Serienmuster für Termine und Nachrichten, die Anforderungen erfüllen werden hinzugefügt.
- Die item.from-Eigenschaft, um im Modus "Verfassen" ebenfalls verfügbar geändert.
- Unterstützung für Ereignisse RecurrenceChanged, RecipientsChanged und AppointmentTimeChanged hinzugefügt.

### <a name="change-log"></a>Änderungsprotokoll

- [Aus](/javascript/api/outlook_1_7/office.from)hinzugefügt: Fügt ein neues Objekt, das bietet eine Methode zum Abrufen der vom Wert.
- [Organizer](/javascript/api/outlook_1_7/office.organizer)hinzugefügt: Fügt ein neues Objekt, das eine Methode zum Abrufen des Werts der Organisator bereitstellt.
- [Serie](/javascript/api/outlook_1_7/office.recurrence)hinzugefügt: Fügt ein neues Objekt, das Methoden zum Abrufen und festlegen das Serienmuster von Terminen, jedoch nur dem Serienmuster von Nachrichten abzurufen, werden Besprechungsanfragen bereitstellt.
- [RecurrenceTimeZone](/javascript/api/outlook_1_7/office.recurrencetimezone)hinzugefügt: Fügt ein neues Objekt, das die Zeitzone Konfiguration des Serienmusters darstellt.
- [SeriesTime](/javascript/api/outlook_1_7/office.seriestime)hinzugefügt: Fügt ein neues Objekt, das Methoden zum Abrufen und Festlegen von Datum und Uhrzeit der Termine in einer Terminserie und zum Abrufen der Daten und Uhrzeiten von Besprechungsanfragen in einer Terminserie bereitstellt.
- [Office.context.mailbox.item.addHandlerAsync](office.context.mailbox.item.md#addhandlerasynceventtype-handler-options-callback)hinzugefügt: Fügt eine neue Methode, die einen Ereignishandler für ein Ereignis unterstützte hinzufügt.
- [Office.context.mailbox.item.from](office.context.mailbox.item.md#from-emailaddressdetailsjavascriptapioutlook17officeemailaddressdetailsfromjavascriptapioutlook17officefrom)geändert: ändert, wenn Sie möchten den Wert im Modus "Verfassen".
- Geänderte [Office.context.mailbox.item.organizer](office.context.mailbox.item.md#organizer-emailaddressdetailsjavascriptapioutlook17officeemailaddressdetailsorganizerjavascriptapioutlook17officeorganizer) - ändert, um den Organizer-Wert im Modus "Verfassen" abzurufen.
- [Office.context.mailbox.item.recurrence](office.context.mailbox.item.md#nullable-recurrence-recurrencejavascriptapioutlook17officerecurrence)hinzugefügt: Fügt eine neue Eigenschaft, die Ruft ab oder legt ein Objekt, das Methoden zum Verwalten von des Serienmusters der ein Terminelement bereitstellt. Diese Eigenschaft kann auch verwendet werden, um das Serienmuster einer Besprechung abzurufen Element Anforderung.
- [Office.context.mailbox.item.removeHandlerAsync](office.context.mailbox.item.md#removehandlerasynceventtype-handler-options-callback)hinzugefügt: Fügt eine neue Methode, die einen Ereignishandler entfernt.
- [Office.context.mailbox.item.seriesId](office.context.mailbox.item.md#nullable-seriesid-string)hinzugefügt: Fügt eine neue Eigenschaft, die der Id der Datenreihe auftreten ein Serientermins versehen gehört.
- [Office.MailboxEnums.Days](/javascript/api/outlook_1_7/office.mailboxenums.days)hinzugefügt: Fügt eine neue Enumeration, die den Tag der Woche oder Typ des Tages angibt.
- [Office.MailboxEnums.Month](/javascript/api/outlook_1_7/office.mailboxenums.month)hinzugefügt: Fügt eine neue Enumeration, die den Monat angibt.
- [Office.MailboxEnums.RecurrenceTimeZone](/javascript/api/outlook_1_7/office.mailboxenums.recurrencetimezone)hinzugefügt: Fügt eine neue Enumeration, die angibt, die Zeitzone auf die Serie angewendet.
- [Office.MailboxEnums.RecurrenceType](/javascript/api/outlook_1_7/office.mailboxenums.recurrencetype)hinzugefügt: Fügt eine neue Enumeration, die den Typ des Serienmusters angibt.
- [Office.MailboxEnums.WeekNumber](/javascript/api/outlook_1_7/office.mailboxenums.weeknumber)hinzugefügt: Fügt eine neue Enumeration, die die Woche des Monats angibt.
- [Office.EventType](/javascript/api/office/office.eventtype)geändert: ändert zur Unterstützung der Ereignisse RecurrenceChanged, RecipientsChanged und AppointmentTimeChanged durch Hinzufügen von `RecurrenceChanged`, `RecipientsChanged`, und `AppointmentTimeChanged` Einträge fest.

## <a name="see-also"></a>Siehe auch

- [Outlook-Add-Ins](https://docs.microsoft.com/outlook/add-ins/)
- [Codebeispiele zu Outlook-Add-Ins](https://developer.microsoft.com/outlook/gallery/?filterBy=Outlook,Samples,Add-ins)
- [Erste Schritte](https://docs.microsoft.com/outlook/add-ins/quick-start)