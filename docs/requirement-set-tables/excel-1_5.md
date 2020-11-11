| Klasse | Felder | Beschreibung |
|:---|:---|:---|
|[CustomXmlPart](/javascript/api/excel/excel.customxmlpart)|[delete()](/javascript/api/excel/excel.customxmlpart#delete--)|Löscht die benutzerdefinierte XML-Komponente.|
||[getXml ()](/javascript/api/excel/excel.customxmlpart#getxml--)|Ruft den vollständigen XML-Inhalt der benutzerdefinierten XML-Komponente ab.|
||[id](/javascript/api/excel/excel.customxmlpart#id)|Die ID des benutzerdefinierten XML-Teils.|
||[NamespaceURI](/javascript/api/excel/excel.customxmlpart#namespaceuri)|Der Namespace-URI des benutzerdefinierten XML-Teils.|
||[setXml (XML: String)](/javascript/api/excel/excel.customxmlpart#setxml-xml-)|Legt den vollständigen XML-Inhalt der benutzerdefinierten XML-Komponente fest.|
|[CustomXmlPartCollection](/javascript/api/excel/excel.customxmlpartcollection)|[Add (XML: Zeichenfolge)](/javascript/api/excel/excel.customxmlpartcollection#add-xml-)|Fügt der Arbeitsmappe eine neue benutzerdefinierte XML-Komponente hinzu.|
||[getByNamespace (NamespaceURI: String)](/javascript/api/excel/excel.customxmlpartcollection#getbynamespace-namespaceuri-)|Ruft eine neue bereichsbezogene Sammlung von benutzerdefinierten XML-Komponenten ab, deren Namespaces dem angegebenen Namespace entsprechen.|
||[getCount()](/javascript/api/excel/excel.customxmlpartcollection#getcount--)|Ruft die Anzahl von CustomXml-Komponenten in der Sammlung ab.|
||[getItem(id: string)](/javascript/api/excel/excel.customxmlpartcollection#getitem-id-)|Ruft eine benutzerdefinierte XML-Komponente basierend auf ihrer ID ab.|
||[getItemOrNullObject(id: string)](/javascript/api/excel/excel.customxmlpartcollection#getitemornullobject-id-)|Ruft eine benutzerdefinierte XML-Komponente basierend auf ihrer ID ab.|
||[items](/javascript/api/excel/excel.customxmlpartcollection#items)|Ruft die geladenen untergeordneten Elemente in dieser Sammlung ab.|
|[CustomXmlPartScopedCollection](/javascript/api/excel/excel.customxmlpartscopedcollection)|[getCount()](/javascript/api/excel/excel.customxmlpartscopedcollection#getcount--)|Ruft die Anzahl von CustomXml-Komponenten in dieser Sammlung ab.|
||[getItem(id: string)](/javascript/api/excel/excel.customxmlpartscopedcollection#getitem-id-)|Ruft eine benutzerdefinierte XML-Komponente basierend auf ihrer ID ab.|
||[getItemOrNullObject(id: string)](/javascript/api/excel/excel.customxmlpartscopedcollection#getitemornullobject-id-)|Ruft eine benutzerdefinierte XML-Komponente basierend auf ihrer ID ab.|
||[getOnlyItem()](/javascript/api/excel/excel.customxmlpartscopedcollection#getonlyitem--)|Wenn die Sammlung genau ein Element enthält, gibt diese Methode es zurück.|
||[getOnlyItemOrNullObject()](/javascript/api/excel/excel.customxmlpartscopedcollection#getonlyitemornullobject--)|Wenn die Sammlung genau ein Element enthält, gibt diese Methode es zurück.|
||[items](/javascript/api/excel/excel.customxmlpartscopedcollection#items)|Ruft die geladenen untergeordneten Elemente in dieser Sammlung ab.|
|[PivotTable](/javascript/api/excel/excel.pivottable)|[id](/javascript/api/excel/excel.pivottable#id)|Die ID der PivotTable.|
|[RequestContext](/javascript/api/excel/excel.requestcontext)|[runtime](/javascript/api/excel/excel.requestcontext#runtime)|[API-Gruppe: ExcelApi 1,5]|
|[Runtime](/javascript/api/excel/excel.runtime)||[Workbook](/javascript/api/excel/excel.workbook)|[customXmlParts](/javascript/api/excel/excel.workbook#customxmlparts)|Stellt die Auflistung von benutzerdefinierten XML-Parts dar, die in dieser Arbeitsmappe enthalten sind.|
|[Arbeitsblatt](/javascript/api/excel/excel.worksheet)|[GetNext (visibleOnly?: Boolean)](/javascript/api/excel/excel.worksheet#getnext-visibleonly-)|Ruft das Arbeitsblatt ab, das diesem folgt.|
||[getNextOrNullObject (visibleOnly?: Boolean)](/javascript/api/excel/excel.worksheet#getnextornullobject-visibleonly-)|Ruft das Arbeitsblatt ab, das diesem folgt.|
||[GetPrevious (visibleOnly?: Boolean)](/javascript/api/excel/excel.worksheet#getprevious-visibleonly-)|Ruft das Arbeitsblatt ab, das diesem vorangestellt ist.|
||[getPreviousOrNullObject (visibleOnly?: Boolean)](/javascript/api/excel/excel.worksheet#getpreviousornullobject-visibleonly-)|Ruft das Arbeitsblatt ab, das diesem vorangestellt ist.|
|[WorksheetCollection](/javascript/api/excel/excel.worksheetcollection)|[GetFirst (visibleOnly?: Boolean)](/javascript/api/excel/excel.worksheetcollection#getfirst-visibleonly-)|Ruft das erste Arbeitsblatt in der Sammlung ab.|
||[GetLast (visibleOnly?: Boolean)](/javascript/api/excel/excel.worksheetcollection#getlast-visibleonly-)|Ruft das letzte Arbeitsblatt in der Sammlung ab.|
