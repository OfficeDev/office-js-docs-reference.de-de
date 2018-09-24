# <a name="office-common-api-requirement-sets"></a>Allgemeine Office-API-Anforderungssätze

Anforderungssätze sind benannte Gruppen von API-Mitgliedern. Office-Add-ins anforderungssätzen im Manifest angegebenen verwenden oder eine Überprüfung zur Laufzeit verwenden, um zu bestimmen, ob ein Office-Host APIs unterstützt, die ein Add-in muss. Weitere Informationen finden Sie unter [Office-Versionen und Anforderung festgelegt](https://docs.microsoft.com/office/dev/add-ins/develop/office-versions-and-requirement-sets).

Benötigen Sie Informationen darüber, wo Add-Ins vom Office-Host unterstützt werden? Finden Sie unter [Office-Add-in-Hosts und Plattform Verfügbarkeit](https://docs.microsoft.com/office/dev/add-ins/overview/office-add-in-availability).

Suchen Sie nach *hostspezifischen* API-Anforderungssätzen? Finden Sie die folgenden API-anforderungssätzen:
 
- [JavaScript-API-Anforderungss?tze f?r Excel](excel-api-requirement-sets.md) (ExcelApi)
- [JavaScript-API-Anforderungss?tze f?r Word](word-api-requirement-sets.md) (WordApi)
- [JavaScript-API-Anforderungss?tze f?r OneNote](onenote-api-requirement-sets.md) (OneNoteApi)
- [Grundlegendes zu Outlook-API-Anforderungss?tzen](outlook-api-requirement-sets.md) (MailBox)

> [!IMPORTANT]
> Es wird nicht mehr empfohlen, dass Sie erstellen und Verwenden von Access Web apps und Datenbanken in SharePoint. Als Alternative wird empfohlen, dass Sie [Microsoft PowerApps](https://powerapps.microsoft.com/) verwenden, ohne Code Business Solutions für Web- und mobile Geräte zu erstellen.

## <a name="common-api-requirement-sets"></a>Allgemeine API-Anforderungssätze

In der folgenden Tabelle sind die allgemeinen API-Anforderungssätze, die Methoden in den einzelnen Sätzen und die Office-Hostanwendungen aufgeführt, die den jeweiligen Anforderungssatz unterstützen. Alle diese API-Anforderungssätze haben Version 1.1.

|**Anforderungssatz**|**Office-Host**|**Methoden im Satz**|
|:-----|:-----|:-----|
| ActiveView | PowerPoint<br>PowerPoint Online<br>PowerPoint für iPad<br>PowerPoint für Mac|Document.getActiveViewAsync|
| AddInCommands | Finden Sie unter [Add-in-Befehl Anforderung festgelegt](add-in-commands-requirement-sets.md). | |
| BindingEvents  | Access Web Apps<br>Excel<br>Excel Online<br>Excel für iPad<br>Excel für Mac<br>Word 2013 und höher<br>Word 2016 für Mac und höher<br>Word Online<br>Word für iPad|Binding.addHanderAsync<br>Binding.removeHanderAsync|
| CompressedFile    | Excel<br>Excel Online<br>Excel für iPad<br>Excel für Mac<br>PowerPoint<br>PowerPoint Online<br>PowerPoint für iPad<br>PowerPoint für Mac<br>Word 2013 und höher<br>Word 2016 für Mac und höher<br>Word Online<br>Word für iPad|Unterstützt die Ausgabe in das Format Office Open XML (OOXML) als Byte-Array<br>(Office.FileType.Compressed), wenn die Methode Document.getFileAsync verwendet wird.|
| CustomXmlParts    | Word 2013 und höher<br>Word 2016 für Mac und höher<br>Word Online<br>Word für iPad|CustomXmlNode.getNodesAsync<br>CustomXmlNode.getNodeValueAsync<br>CustomXmlNode.getXmlAsync<br>CustomXmlNode.setNodeValueAsync<br>CustomXmlNode.setXmlAsync<br>CustomXmlPart.addHandlerAsync<br>CustomXmlPart.deleteAsync<br>CustomXmlPart.getNodesAsync<br>CustomXmlPart.getXmlAsync<br>CustomXmlPart.removeHandlerAsync<br>CustomXmlParts.addAsync<br>CustomXmlParts.getByIdAsync<br>CustomXmlParts.getByNamespaceAsync<br>CustomXmlPrefixMappings.addNamespaceAsync<br>CustomXmlPrefixMappings.getNamespaceAsync<br>CustomXmlPrefixMappings.getPrefixAsync|
| DialogApi | Finden Sie unter [Dialog API-anforderungssätzen](dialog-api-requirement-sets.md). | UI.messageParent<br>UI.displayDialogAsync<br>UI.closeContainer<br>UI.Dialog |
| DocumentEvents    | Excel<br>Excel Online<br>Excel für iPad<br>Excel für Mac<br>OneNote Online<br>PowerPoint<br>PowerPoint Online<br>PowerPoint für iPad<br>PowerPoint für Mac<br>Word 2013 und höher<br>Word 2016 für Mac und höher<br>Word Online<br>Word für iPad|Document.addHandlerAsync<br>Document.removeHandlerAsync|
| Datei  | Excel<br>Excel Online<br>Excel für iPad<br>Excel für Mac<br>PowerPoint<br>PowerPoint Online<br>PowerPoint für iPad<br>PowerPoint für Mac<br>Word 2013 und höher<br>Word 2016 für Mac und höher<br>Word Online<br>Word für iPad|Document.getFileAsync<br>File.closeAsync<br>File.getSliceAsync|
| HtmlCoercion  | OneNote Online<br>Word 2013 und höher<br>Word 2016 für Mac und höher<br>Word Online<br>Word für iPad|Unterstützt die Koersion ins HTML-Format (Office.CoercionType.Html) beim Lesen und Schreiben von Daten mithilfe der Methoden "Document.getSelectedDataAsync",<br>"Document.setSelectedDataAsync", "Binding.getDataAsync" oder "Binding.setDataAsync".|
| IdentityAPI | Finden Sie unter [Identity-API-anforderungssätzen](identity-api-requirement-sets.md). | Auth.getAccessTokenAsync |
| ImageCoercion | Excel<br>Excel für iPad<br>Excel für Mac<br>OneNote Online<br>PowerPoint<br>PowerPoint Online<br>PowerPoint für iPad<br>PowerPoint für Mac<br>Word 2013 und höher<br>Word 2016 für Mac und höher<br>Word Online<br>Word für iPad|Unterstützt beim Schreiben von Daten mithilfe der Methode "Document.setSelectedDataAsync" die Konvertierung in ein Bild (Office.CoercionType.Image).|
| Postfach   |Outlook für Windows<br>Outlook für Web<br>Outlook für Android<br>Outlook für Mac<br>Outlook Web App |Siehe [Grundlegendes zu Outlook-API-Anforderungssätzen](outlook-api-requirement-sets.md).|
| MatrixBindings    | Excel<br>Excel Online<br>Excel für iPad<br>Excel für Mac<br>Word<br>Word Online<br>Word für iPad<br>Word für Mac|Bindings.addFromNamedItemAsync<br>Bindings.addFromSelectionAsync<br>Bindings.getAllAsync<br>Bindings.getByIdAsync<br>Bindings.releaseByIdAsyncMatrix<br>Binding.getDataAsyncMatrix<br>Binding.setDataAsync|
| MatrixCoercion    | Excel<br>Excel Online<br>Excel für iPad<br>Excel für Mac<br>Word 2013 und höher<br>Word 2016 für Mac und höher<br>Word Online<br>Word für iPad|Unterstützt die Koersion in eine "Matrix"- (Array von Arrays-)Datenstruktur (Office.CoercionType.Matrix) beim Lesen und Schreiben von Daten mit der Methode "Document.getSelectedDataAsync", "Document.setSelectedDataAsync", "Binding.getDataAsync" oder "Binding.setDataAsync".|
| OoxmlCoercion | Word 2013 und höher<br>Word 2016 für Mac und höher<br>Word Online<br>Word für iPad|Unterstützt die Koersion in das Open Office XML-Format (Office.CoercionType.Ooxml) beim Lesen und Schreiben von Daten mit der Methode "Document.getSelectedDataAsync", "Document.setSelectedDataAsync", "Binding.getDataAsync" oder "Binding.setDataAsync".|
| PartialTableBindings  | Access Web Apps||
| PdfFile   | Excel für Mac<br>PowerPoint<br>PowerPoint Online<br>PowerPoint für iPad<br>PowerPoint für Mac<br>Word 2013 und höher<br>Word 2016 für Mac und höher<br>Word Online<br>Word für iPad|Unterstützt die Ausgabe im PDF-Format (Office.FileType.Pdf)<br>bei Verwendung der Document.getFileAsync-Methode|
| Selection | Excel<br>Excel Online<br>Excel für iPad<br>Excel für Mac<br>PowerPoint<br>PowerPoint Online<br>PowerPoint für iPad<br>PowerPoint für Mac<br>Project<br>Word 2013 und höher<br>Word 2016 für Mac und höher<br>Word Online<br>Word für iPad|Document.getSelectedDataAsync<br>Document.setSelectedDataAsync|
| Settings  | Access Web Apps<br>Excel<br>Excel Online<br>Excel für iPad<br>Excel für Mac<br>OneNote Online<br>PowerPoint<br>PowerPoint Online<br>PowerPoint für iPad<br>PowerPoint für Mac<br>Word 2013 und höher<br>Word 2016 für Mac und höher<br>Word Online<br>Word für iPad|Settings.get<br>Settings.remove<br>Settings.saveAsync<br>Settings.set|
| TableBindings | Access Web Apps<br>Excel<br>Excel Online<br>Excel für iPad<br>Excel für Mac<br>Word 2013 und höher<br>Word 2016 für Mac und höher<br>Word Online<br>Word für iPad|Bindings.addFromNamedItemAsync<br>Bindings.addFromSelectionAsync<br>Bindings.getAllAsync<br>Bindings.getByIdAsync<br>Bindings.releaseByIdAsyncTable<br>Binding.addColumnsAsyncTable<br>Binding.addRowsAsyncTable<br>Binding.deleteAllDataValuesAsyncTable<br>Binding.getDataAsyncTable<br>Binding.setDataAsync|
| TableCoercion | Access Web Apps<br>Excel<br>Excel Online<br>Excel für iPad<br>Excel für Mac<br>Word 2013 und höher<br>Word 2016 für Mac und höher<br>Word Online<br>Word für iPad|Unterstützt die Koersion in eine Tabellendatenstruktur (Office.CoercionType.Table) beim Lesen und Schreiben von Daten mit der Methode "Document.getSelectedDataAsync", "Document.setSelectedDataAsync", "Binding.getDataAsync" oder "Binding.setDataAsync".|
| TextBindings  | Excel<br>Excel Online<br>Excel für iPad<br>Word 2013 und höher<br>Word 2016 für Mac und höher<br>Word Online<br>Word für iPad|Bindings.addFromNamedItemAsync<br>Bindings.addFromSelectionAsync<br>Bindings.getAllAsync<br>Bindings.getByIdAsync<br>Bindings.releaseByIdAsyncText<br>Binding.getDataAsyncText<br>Binding.setDataAsync|
| TextCoercion  | Excel<br>Excel Online<br>Excel für iPad<br>OneNote Online<br>PowerPoint<br>PowerPoint Online<br>PowerPoint für iPad<br>PowerPoint für Mac<br>Project<br>Word 2013 und höher<br>Word 2016 für Mac und höher<br>Word Online<br>Word für iPad|Unterstützt die Koersion in das Textformat (Office.CoercionType.Text) beim Lesen und Schreiben von Daten mit der Methode "Document.getSelectedDataAsync", "Document.setSelectedDataAsync", "Binding.getDataAsync" oder "Binding.setDataAsync".|
| TextFile  | Word 2013 und höher<br>Word 2016 für Mac und höher<br>Word Online<br>Word für iPad|Unterstützt die Ausgabe in das Textformat (Office.FileType.Text), wenn Sie die Document.getFileAsync-Methode verwenden.|

## <a name="methods-that-arent-part-of-a-requirement-set"></a>Methoden, die nicht Teil eines Anforderungssatzes sind

Die folgenden Methoden in der JavaScript-API für Office gehören zu keinem Anforderungssatz. Wenn das Add-in einer dieser Methoden erforderlich sind, verwenden Sie die **Methoden** und **Methode** Elemente im Manifest das Add-in um zu deklarieren, dass sie erforderlich sind oder Ausführen der Common Language Runtime Kontrollkästchen mit einer `if` Anweisung. Weitere Informationen finden Sie unter [Angeben von Office-Hosts und API-Anforderungen](https://docs.microsoft.com/office/dev/add-ins/develop/specify-office-hosts-and-api-requirements).

|**Methodenname**|**Office-Hostunterstützung**|
|:-----|:-----|
|Bindings.addFromPromptAsync|Web-apps für Access, Excel, Excel Online und Excel für iPad|
|Document.getFilePropertiesAsync|Excel, Excel Online, Excel für iPad, Excel für Mac, PowerPoint, PowerPoint Online, PowerPoint für iPad, PowerPoint für Mac, Word, Word Online, Word für iPad und Word für Mac|
|Document.getProjectFieldAsync|Project Standard 2013 und Project Professional 2013|
|Document.getResourceFieldAsync|Project Standard 2013 und Project Professional 2013|
|Document.getSelectedResourceAsync|Project Standard 2013 und Project Professional 2013|
|Document.getSelectedTaskAsync|Project Standard 2013 und Project Professional 2013|
|Document.getSelectedViewAsync|Project Standard 2013 und Project Professional 2013|
|Document.getTaskAsync|Project Standard 2013 und Project Professional 2013|
|Document.getTaskFieldAsync|Project Standard 2013 und Project Professional 2013|
|Document.goToByIdAsync|Excel, Excel Online, Excel für iPad, Excel für Mac, PowerPoint, PowerPoint Online, PowerPoint für iPad, PowerPoint für Mac, Word, Word Online, Word für iPad und Word für Mac|
|Settings.addHandlerAsync|Zugriff auf Web-apps, Excel, Online Excel, PowerPoint, Online PowerPoint, Word und Word Online|
|Settings.refreshAsync|Zugriff auf Web-apps, Excel, Online Excel, PowerPoint, Online PowerPoint, Word und Word Online|
|Settings.removeHandlerAsync|Zugriff auf Web-apps, Excel, Online Excel, PowerPoint, Online PowerPoint, Word und Word Online|
|TableBinding.clearFormatsAsync|Excel, Excel Web App und Excel für Mac|
|TableBinding.setFormatsAsync|Excel, Excel Online, für iPad Excel und Excel für Mac|
|TableBinding.setTableOptionsAsync|Excel, Excel Online, für iPad Excel und Excel für Mac|

## <a name="see-also"></a>Siehe auch

- [Office-Versionen und Anforderungssätze](https://docs.microsoft.com/office/dev/add-ins/develop/office-versions-and-requirement-sets)
- 
  [Angeben von Office-Hosts und API-Anforderungen](https://docs.microsoft.com/office/dev/add-ins/develop/specify-office-hosts-and-api-requirements)
- [XML-Manifest für Office-Add-Ins](https://docs.microsoft.com/office/dev/add-ins/develop/add-in-manifests)
