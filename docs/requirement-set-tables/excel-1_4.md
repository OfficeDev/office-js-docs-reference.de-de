| Klasse | Felder | Beschreibung |
|:---|:---|:---|
|[BindingCollection](/javascript/api/excel/excel.bindingcollection)|[getCount()](/javascript/api/excel/excel.bindingcollection#getcount--)|Ruft die Anzahl der Bindungen in der Sammlung ab.|
||[getItemOrNullObject(id: string)](/javascript/api/excel/excel.bindingcollection#getitemornullobject-id-)|Ruft ein binding-Objekt anhand seiner ID ab.|
|[ChartCollection](/javascript/api/excel/excel.chartcollection)|[getCount()](/javascript/api/excel/excel.chartcollection#getcount--)|Gibt die Anzahl der Diagramme im Arbeitsblatt zurück.|
||[GetItemOrNullObject(name: string)](/javascript/api/excel/excel.chartcollection#getitemornullobject-name-)|Ruft ein Diagramm über seinen Namen ab.|
|[ChartPointsCollection](/javascript/api/excel/excel.chartpointscollection)|[getCount()](/javascript/api/excel/excel.chartpointscollection#getcount--)|Gibt die Anzahl der Diagrammpunkten in der Reihe zurück.|
|[ChartSeriesCollection](/javascript/api/excel/excel.chartseriescollection)|[getCount()](/javascript/api/excel/excel.chartseriescollection#getcount--)|Gibt die Anzahl der Datenreihen in der Sammlung zurück.|
|[NamedItem](/javascript/api/excel/excel.nameditem)|[comment](/javascript/api/excel/excel.nameditem#comment)|Gibt den diesem Namen zugeordneten Kommentar an.|
||[delete()](/javascript/api/excel/excel.nameditem#delete--)|Löscht den angegebenen Namen.|
||[getRangeOrNullObject()](/javascript/api/excel/excel.nameditem#getrangeornullobject--)|Ruft das Bereichsobjekt ab, das mit dem Namen verknüpft ist.|
||[scope](/javascript/api/excel/excel.nameditem#scope)|Gibt an, ob der Name auf die Arbeitsmappe oder auf ein bestimmtes Arbeitsblatt beschränkt ist.|
||[worksheet](/javascript/api/excel/excel.nameditem#worksheet)|Gibt das Arbeitsblatt zurück, auf dessen Bereich das benannte Element beschränkt ist.|
||[worksheetOrNullObject](/javascript/api/excel/excel.nameditem#worksheetornullobject)|Gibt das Arbeitsblatt zurück, auf dessen Bereich das benannte Element beschränkt ist.|
|[NamedItemCollection](/javascript/api/excel/excel.nameditemcollection)|[Add (Name: String, Reference: Range \| String, comment?: String)](/javascript/api/excel/excel.nameditemcollection#add-name--reference--comment-)|Fügt einen neuen Namen zur Auflistung des angegebenen Bereichs hinzu.|
||[addFormulaLocal (Name: String, Formula: String, comment?: String)](/javascript/api/excel/excel.nameditemcollection#addformulalocal-name--formula--comment-)|Fügt einen neuen Namen zu der Auflistung des angegebenen Bereichs unter Verwendung des Gebietsschemas des Benutzers für die Formel hinzu.|
||[getCount()](/javascript/api/excel/excel.nameditemcollection#getcount--)|Ruft die Anzahl von benannten Elementen in der Auflistung ab.|
||[GetItemOrNullObject(name: string)](/javascript/api/excel/excel.nameditemcollection#getitemornullobject-name-)|Ruft ein NamedItem-Objekt mit seinem Namen ab.|
|[PivotTableCollection](/javascript/api/excel/excel.pivottablecollection)|[getCount()](/javascript/api/excel/excel.pivottablecollection#getcount--)|Ruft die Anzahl von PivotTables in der Auflistung ab.|
||[GetItemOrNullObject(name: string)](/javascript/api/excel/excel.pivottablecollection#getitemornullobject-name-)|Ruft eine PivotTable anhand des Namens ab.|
|[Range](/javascript/api/excel/excel.range)|[getIntersectionOrNullObject (anotherRange: Bereichs \| Zeichenfolge)](/javascript/api/excel/excel.range#getintersectionornullobject-anotherrange-)|Ruft das Bereichsobjekt ab, das die rechteckige Schnittmenge der angegebenen Bereiche darstellt.|
||[getUsedRangeOrNullObject (valuesOnly?: Boolean)](/javascript/api/excel/excel.range#getusedrangeornullobject-valuesonly-)|Gibt den verwendeten Bereich des angegebenen Bereichsobjekts zurück.|
|[RangeViewCollection](/javascript/api/excel/excel.rangeviewcollection)|[getCount()](/javascript/api/excel/excel.rangeviewcollection#getcount--)|Gibt die Anzahl von RangeView-Objekten in der Auflistung zurück.|
|[Einstellung](/javascript/api/excel/excel.setting)|[delete()](/javascript/api/excel/excel.setting#delete--)|Löscht die Einstellung.|
||[key](/javascript/api/excel/excel.setting#key)|Der Schlüssel, der die ID der Einstellung darstellt.|
||[value](/javascript/api/excel/excel.setting#value)|Stellt den Wert dar, der für diese Einstellung gespeichert ist.|
|[SettingCollection](/javascript/api/excel/excel.settingcollection)|[Add (Key: String, Value: String \| Number \| Boolean \| Date \| array <any> \| any)](/javascript/api/excel/excel.settingcollection#add-key--value-)|Legt die angegebene Einstellung fest oder fügt sie zur Arbeitsmappe hinzu.|
||[getCount()](/javascript/api/excel/excel.settingcollection#getcount--)|Ruft die Anzahl der Einstellungen in der Sammlung ab.|
||[getItem(key: string)](/javascript/api/excel/excel.settingcollection#getitem-key-)|Ruft einen Setting-Eintrag über den Schlüssel ab.|
||[getItemOrNullObject(key: string)](/javascript/api/excel/excel.settingcollection#getitemornullobject-key-)|Ruft einen Setting-Eintrag über den Schlüssel ab.|
||[items](/javascript/api/excel/excel.settingcollection#items)|Ruft die geladenen untergeordneten Elemente in dieser Sammlung ab.|
||[onSettingsChanged](/javascript/api/excel/excel.settingcollection#onsettingschanged)|Tritt ein, wenn die Einstellungen im Dokument geändert werden.|
|[SettingsChangedEventArgs](/javascript/api/excel/excel.settingschangedeventargs)|[Einstellungen](/javascript/api/excel/excel.settingschangedeventargs#settings)|Ruft das Setting-Objekt ab, das die Bindung darstellt, durch die das SettingsChanged-Ereignis ausgelöst wurde.|
|[TableCollection](/javascript/api/excel/excel.tablecollection)|[getCount()](/javascript/api/excel/excel.tablecollection#getcount--)|Ruft die Anzahl von Tabellen in der Auflistung ab.|
||[getItemOrNullObject(key: string)](/javascript/api/excel/excel.tablecollection#getitemornullobject-key-)|Ruft eine Tabelle anhand des Namens oder der ID ab.|
|[TableColumnCollection](/javascript/api/excel/excel.tablecolumncollection)|[getCount()](/javascript/api/excel/excel.tablecolumncollection#getcount--)|Ruft die Anzahl der Spalten in der Tabelle ab.|
||[getItemOrNullObject (Key: Number \| String)](/javascript/api/excel/excel.tablecolumncollection#getitemornullobject-key-)|Ruft ein Spaltenobjekt nach Name oder ID ab.|
|[TableRowCollection](/javascript/api/excel/excel.tablerowcollection)|[getCount()](/javascript/api/excel/excel.tablerowcollection#getcount--)|Ruft die Anzahl der Zeilen in der Tabelle ab.|
|[Workbook](/javascript/api/excel/excel.workbook)|[Einstellungen](/javascript/api/excel/excel.workbook#settings)|Stellt eine Auflistung der mit der Arbeitsmappe verknüpften Einstellungen dar.|
|[Arbeitsblatt](/javascript/api/excel/excel.worksheet)|[getUsedRangeOrNullObject (valuesOnly?: Boolean)](/javascript/api/excel/excel.worksheet#getusedrangeornullobject-valuesonly-)|Der verwendete Bereich ist der kleinste Bereich, der mindestens eine der Zellen umfasst, die einen Wert enthalten oder denen eine Formatierung zugewiesen wurde.|
||[Namen](/javascript/api/excel/excel.worksheet#names)|Auflistung von Namen im Bereich des aktuellen Arbeitsblatts.|
|[WorksheetCollection](/javascript/api/excel/excel.worksheetcollection)|[GetCount (visibleOnly?: Boolean)](/javascript/api/excel/excel.worksheetcollection#getcount-visibleonly-)|Ruft die Anzahl von Arbeitsblättern in der Auflistung ab.|
||[getItemOrNullObject(key: string)](/javascript/api/excel/excel.worksheetcollection#getitemornullobject-key-)|Ruft das Arbeitsblattobjekt mithilfe des Namens oder der ID ab.|
