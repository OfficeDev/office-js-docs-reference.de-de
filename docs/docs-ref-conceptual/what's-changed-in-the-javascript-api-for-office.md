# <a name="whats-changed-in-the-javascript-api-for-office"></a>Änderungen in der JavaScript-API für Office

Die JavaScript-API für Office bietet neue und aktualisierte Objekte, Methoden, Eigenschaften, Ereignisse und Enumerationen, welche den Funktionsumfang Ihrer Office-Add-Ins erweitern. Verwenden Sie die Links unten, um die neuen und aktualisierten API-Mitglieder anzuzeigen.

Um Add-Ins mithilfe der neuen API-Mitglieder zu entwickeln, müssen Sie [die Dateien der JavaScript-API für Office in Ihrem Projekt aktualisieren](https://docs.microsoft.com/office/dev/add-ins/develop/update-your-javascript-api-for-office-and-manifest-schema-version).

Alle API-Mitglieder, einschließlich derjenigen, die aus vorherigen Updates unverändert übernommen wurden, finden Sie unter [JavaScript API for Office](javascript-api-for-office.md).

## <a name="new-and-updated-apis"></a>Neue und aktualisierte API

### <a name="new-and-updated-objects"></a>Neue und aktualisierte Objekte

|**Objekt**|**Beschreibung**|**Version hinzugefügt oder aktualisiert**|
|:-----|:-----|:-----|
|`Item`|Aktualisierungen und Ergänzungen:<br><ul><li><p>Die `getSelectedDataAsync` und `setSelectedDataAsync` Methoden zum Abrufen der Auswahl des Benutzers und in den Betreff und den Textkörper einer Nachricht oder eines Termins überschreiben unterstützen.</p></li><li><p>Die `displayReplyAllForm` und `displayReplyForm` Methoden zum Hinzufügen einer Anlage zu das Antwortformular eines Termins unterstützen.</p></li></ul>|Mailbox 1.2|
|`Item`|Mit Methoden und Feldern zum Erstellen von Verfassenmodus-Outlook-Add-Ins aktualisiert. |1.1|
|`Binding`|Unterstützt jetzt Tabellenbindung in Inhalts-Add-ins für Access.|1.1|
|`Bindings`|Unterstützt jetzt Tabellenbindung in Inhalts-Add-ins für Access.|1.1|
|`Body`|Wurde hinzugefügt, um das Erstellen und Bearbeiten des Textkörpers einer Nachricht oder eines Termins in Verfassenmodus-Outlook-Add-Ins zu ermöglichen.|1.1|
|`Document`|Aktualisierungen und Ergänzungen: <ul><li><p>Unterstützung der Eigenschaften <a href="https://docs.microsoft.com/javascript/api/office/office.document?view=office-js" target="_blank">mode</a>, <a href="https://docs.microsoft.com/javascript/api/office/office.document?view=office-js#settings" target="_blank">settings</a> und <a href="https://docs.microsoft.com/javascript/api/office/office.document?view=office-js" target="_blank">url</a> in Inhalts-Add-Ins für Access</p></li><li><p>Abrufen des Dokuments als PDF mit der Methode <a href="https://docs.microsoft.com/javascript/api/office/office.document?view=office-js#getfileasync-filetype--options--callback-" target="_blank">getFileAsync</a> in Add-Ins für PowerPoint und Word</p></li><li><p>Abrufen von Dateieigenschaften mit der Methode <a href="https://docs.microsoft.com/javascript/api/office/office.document?view=office-js#getfilepropertiesasync-options--callback-" target="_blank">getFileProperties</a> in Add-Ins für Excel, PowerPoint und Word</p></li><li><p>Navigieren zu Speicherorten und Objekten im Dokument mit der Methode <a href="https://docs.microsoft.com/javascript/api/office/office.document?view=office-js#gotobyidasync-id--gototype--options--callback-" target="_blank">goToByIdAsync</a> in Add-Ins für Excel und PowerPoint</p></li><li><p>Abrufen der ID, des Titels und des Indexes für ausgewählte Folien mit der Methode <a href="https://docs.microsoft.com/javascript/api/office/office.document?view=office-js#getselecteddataasync-coerciontype--options--callback-" target="_blank">getSelectedDataAsync</a> (wenn Sie die neue <span class="keyword">Office.CoercionType.SlideRange</span><a href="https://docs.microsoft.com/javascript/api/office/office.coerciontype?view=office-js" target="_blank">coercionType</a>-Enumeration angeben) in Add-Ins für PowerPoint</p></li></ul>|1.1|
|`Location`|Wurde hinzugefügt, um das Festlegen des Orts eines Termins in Verfassenmodus-Outlook-Add-Ins zu ermöglichen.|1.1|
|`Office`|Die Auswahlmethode unterstützt jetzt das Abrufen von Bindungen in Inhalts-Add-ins für Access.|1.1|
|`Recipients`|Wurde hinzugefügt, um das Abrufen und Festlegen der Empfänger einer Nachricht oder eines Termins im Verfassenmodus zu ermöglichen.|1.1|
|`Settings`|Unterstützt jetzt das Erstellen benutzerdefinierter Einstellungen in Inhalts-Add-ins für Access.|1.1|
|`Subject`|Wurde hinzugefügt, um das Abrufen und Festlegen des Betreffs einer Nachricht oder eines Termins in Verfassenmodus-Outlook-Add-Ins zu ermöglichen.|1.1|
|`Time`|Wurde hinzugefügt, um das Abrufen und Festlegen der Start- und Endzeit eines Termins in Verfassenmodus-Outlook-Add-Ins zu ermöglichen.|1.1|

### <a name="new-and-updated-enumerations"></a>Neue und aktualisierte Enumerationen

|**Objekt**|**Beschreibung**|**Version**|
|:-----|:-----|:-----|
|`ActiveView`|Gibt den Status der aktiven Ansicht des Dokuments an, z. B. ob der Benutzer das Dokument bearbeiten kann.Wurde hinzugefügt, damit Add-ins für PowerPoint ermitteln können, ob die Benutzer die Präsentation ( **Diaschau**) anzeigen oder Folien bearbeiten. |1.1|
|`CoercionType`|Mit  **Office.CoercionType.SlideRange** aktualisiert, um das Abrufen des ausgewählten Folienbereichs mit der Methode **getSelectedDataAsync** in Add-ins für PowerPoint zu unterstützen.|1.1|
|`EventType`|Mit dem neuen ActiveViewChanged-Ereignis aktualisiert.|1.1|
|`FileType`|Gibt jetzt Ausgaben in PDF-Format an.|1.1|
|`GoToType`|Wurde hinzugefügt, um die Stelle oder das Objekt im Dokument anzugeben, zu dem gewechselt werden soll.|1.1|

