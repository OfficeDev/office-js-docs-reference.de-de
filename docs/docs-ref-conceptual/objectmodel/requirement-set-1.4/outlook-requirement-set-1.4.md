# <a name="outlook-add-in-api-requirement-set-14"></a>Outlook-Add-In-API-Anforderungssatz 1.4

Die Outlook-Add-In-API-Teilmenge der JavaScript-API für Office umfasst Objekte, Methoden, Eigenschaften und Ereignisse, die Sie in Outlook-Add-Ins verwenden können.

> [!NOTE]
> Diese Dokumentation ist für eine [Anforderung festgelegt](/javascript/office/requirement-sets/outlook-api-requirement-sets) als die aktuelle Anforderungssatz.

## <a name="whats-new-in-14"></a>Neu in Version 1.4

Der Anforderungssatz 1.4 umfasst alle Funktionen von [Anforderungssatz 1.3](../requirement-set-1.3/outlook-requirement-set-1.3.md). Zusätzlich hat er Zugriff auf den Namespace `Office.ui` implementiert.

### <a name="change-log"></a>Änderungsprotokoll

- [Office.context.ui.displayDialogAsync](/javascript/api/office/office.ui#displaydialogasync-startaddress--options--callback-) hinzugefügt: Zeigt ein Dialogfeld in einem Office-Host an.
- [Office.context.ui.messageParent](/javascript/api/office/office.ui#messageparent-messageobject-) hinzugefügt: Übermittelt eine Nachricht von einem Dialogfeld an dessen übergeordnete Seite/Öffnerseite.
- [Dialog](/javascript/api/office/office.dialog) -Objekt hinzugefügt: das Objekt, das wird der [`displayDialogAsync`](/javascript/api/office/office.ui#displaydialogasync-startaddress--options--callback-) -Methode aufgerufen wird.

## <a name="see-also"></a>Siehe auch

- [Outlook-Add-Ins](https://docs.microsoft.com/outlook/add-ins/)
- [Codebeispiele zu Outlook-Add-Ins](https://developer.microsoft.com/outlook/gallery/?filterBy=Outlook,Samples,Add-ins)
- [Erste Schritte](https://docs.microsoft.com/outlook/add-ins/quick-start)