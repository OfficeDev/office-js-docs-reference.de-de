# <a name="outlook-add-in-api-requirement-set-13"></a>Outlook-Add-In-API-Anforderungssatz 1.3

Die Outlook-add-in-API-Teilmenge der JavaScript-API für Office umfasst Objekte, Methoden, Eigenschaften und Ereignisse für die Verwendung in einer Outlook-add-Ins.

> [!NOTE]
> Diese Dokumentation ist für eine [Anforderung festgelegt](/javascript/office/requirement-sets/outlook-api-requirement-sets) als die aktuelle Anforderungssatz. 

## <a name="whats-new-in-13"></a>Neu in Version 1.3

Der Anforderungssatz 1.3 umfasst alle Funktionen von [Anforderungssatz 1.2](../requirement-set-1.2/outlook-requirement-set-1.2.md). Zusätzlich wurden die folgenden Funktionen implementiert:

- Es wurde Unterstützung für [Add-In-Befehle](https://docs.microsoft.com/outlook/add-ins/add-in-commands-for-outlook) hinzugefügt.
- Es wurde eine Funktion hinzugefügt, über die sich Elemente während des Verfassens speichern oder schließen lassen.
- Das Objekt [Body](/javascript/api/outlook_1_3/office.body) wurde erweitert. Add-Ins können jetzt den gesamten Text abrufen oder festlegen.
- Es wurden Konvertierungsmethoden hinzugefügt, die IDs aus dem EWS-Format ins REST-Format konvertieren und umgekehrt.
- Es wurde eine Funktion hinzugefügt, über die der Infoleiste von Elementen Benachrichtigungen hinzugefügt werden können.

### <a name="change-log"></a>Änderungsprotokoll

- [Body.getAsync](/javascript/api/outlook_1_3/office.body#getasync-coerciontype--options--callback-) hinzugefügt: Gibt den aktuellen Text im angegebenen Format zurück.
- [Body.setAsync](/javascript/api/outlook_1_3/office.body#setasync-data--options--callback-) hinzugefügt: Ersetzt den gesamten Text durch den angegebenen Text.
- [Office.context.officeTheme](office.context.md#officetheme-object) hinzugefügt: Stellt Zugriff auf die Office-Farbdesigns bereit.
- Objekt [Event](/javascript/api/office/office.addincommands.event) hinzugefügt: Wird als Parameter an UI-lose Befehlsfunktionen in einem Outlook-Add-In übergeben. Wird verwendet, um den Abschluss der Verarbeitung zu signalisieren.
- [Office.context.mailbox.item.close](office.context.mailbox.item.md#close) hinzugefügt: Schließt das aktuelle Element, das gerade verfasst wird.
- [Office.context.mailbox.item.saveAsync](office.context.mailbox.item.md#saveasyncoptions-callback) hinzugefügt: Speichert ein Element asynchron.
- [Office.context.mailbox.item.notificationMessages](office.context.mailbox.item.md#notificationmessages-notificationmessagesjavascriptapioutlook13officenotificationmessages) hinzugefügt: Ruft die Benachrichtigungen für ein Element ab.
- [Office.context.mailbox.convertToEwsId](office.context.mailbox.md#converttoewsiditemid-restversion--string) hinzugefügt: Wandelt eine REST-formatierte Element-ID ins EWS-Format um.
- [Office.context.mailbox.convertToRestId](office.context.mailbox.md#converttorestiditemid-restversion--string) hinzugefügt: Wandelt eine EWS-formatierte Element-ID ins REST-Format um.
- [Office.MailboxEnums.ItemNotificationMessageType](/javascript/api/outlook_1_3/office.mailboxenums.itemnotificationmessagetype) hinzugefügt: Gibt den Benachrichtigungstyp für einen Termin oder eine Nachricht an.
- [Office.MailboxEnums.RestVersion](/javascript/api/outlook_1_3/office.mailboxenums.restversion) hinzugefügt: Gibt die Version der REST-API an, die einer REST-formatierten Element-ID entspricht.
- Objekt [NotificationMessages](/javascript/api/outlook_1_3/office.notificationmessages) hinzugefügt: Stellt Methoden für den Zugriff auf Benachrichtigungen in einem Outlook-Add-In bereit.
- Typ [NotificationMessageDetails](/javascript/api/outlook_1_3/office.notificationmessagedetails) hinzugefügt: Wird von der Methode `NotificationMessages.getAllAsync` zurückgegeben.

## <a name="see-also"></a>Siehe auch

- [Outlook-Add-Ins](https://docs.microsoft.com/outlook/add-ins/)
- [Codebeispiele zu Outlook-Add-Ins](https://developer.microsoft.com/outlook/gallery/?filterBy=Outlook,Samples,Add-ins)
- [Erste Schritte](https://docs.microsoft.com/outlook/add-ins/quick-start)