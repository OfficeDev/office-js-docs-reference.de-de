# <a name="javascript-api-for-office"></a>JavaScript-API für Office

Die JavaScript-API für Office ermöglicht Ihnen das Erstellen von Webanwendungen, die Interaktion mit den Objektmodellen in Office-hostanwendungen. Die Anwendung wird die Bibliothek "Office.js" verweisen, die ein Skript Loader ist. Die office.js-Bibliothek lädt die Objektmodelle, die für die Office-Anwendung anwendbar sind, die das Add-in ausgeführt wird. Sie können die folgenden JavaScript-Objektmodellen verwenden:

- **Allgemeine APIs** - APIs, die eingeführt wurden mit **Office 2013**. Dies ist für **alle Office-hostanwendungen** geladen und verbindet die Add-in-Anwendung mit der Office-Clientanwendung. Das Objektmodell enthält APIs, die speziell für Office-Clients und APIs, die für mehrere Office-Host-Clientanwendungen verfügbar sind. All diese Inhalte unter **API freigegeben**ist. 

  **Outlook** wird auch die allgemeine Syntax der API verwendet. Alle Websites unter dem Alias Office enthält Objekte, die Sie zum Schreiben von Skripts, die mit dem Inhalt im Office-Dokumenten, Arbeitsblättern, Präsentationen, e-Mail-Elemente und Projekte aus Ihrer Office-Add-ins interagieren verwenden können. Sie müssen diese allgemeine APIs verwenden, wenn Ihr Add-in in Office 2013 und höher konzipiert ist. Dieses Objektmodell verwendet Rückrufe.

- **Host-spezifische APIs** - APIs, die eingeführt wurden mit **Office 2016**. Dieses Objektmodell enthält Host-spezifische stark typisierten Objekte, die vertraute-Objekte entsprechen, die angezeigt werden, wenn Sie Office-Clients verwenden, und die Zukunft der Office JavaScript-APIs stellt. Die Host-spezifische APIs enthalten derzeit der JavaScript-API für Word und Excel JavaScript-API.

## <a name="supported-host-applications"></a>Unterstützte Hostanwendungen

- [Excel](overview/excel-add-ins-reference-overview.md)
- [OneNote](overview/onenote-add-ins-javascript-reference.md)
- [Outlook](requirement-sets/outlook-api-requirement-sets.md)
- [Visio](overview/visio-javascript-reference-overview.md)
- [Word](overview/word-add-ins-reference-overview.md)
- [Freigegebene API](requirement-sets/office-add-in-requirement-sets.md)

> [!NOTE] 
> [PowerPoint und Project](requirement-sets/powerpoint-and-project-note.md) Support-add-ins, die mit der JavaScript-API. Allerdings müssen sie derzeit Host-spezifische APIs keine. Interaktion mit dieser Hosts über die Shared-API.

Erfahren Sie mehr über [unterstützte Hosts und andere Anforderungen](https://docs.microsoft.com/office/dev/add-ins/concepts/requirements-for-running-office-add-ins).

## <a name="open-api-specifications"></a>Offene API-Spezifikationen

Während des Entwerfens und Entwickelns neuer APIs für Office-Add-Ins stellen wir diese zur Verfügung, damit Sie auf der Seite [Offene API-Spezifikationen](openspec.md) Ihr Feedback abgeben können. Erfahren Sie, welche neuen Funktionen geplant sind, und teilen Sie uns Ihre Meinung zu unseren Designspezifikationen mit.

## <a name="see-also"></a>Siehe auch

- [Office JavaScript-API-Referenz](https://docs.microsoft.com/javascript/api/overview/office?view=office-js)