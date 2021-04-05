| Klasse | Felder | Beschreibung |
|:---|:---|:---|
|[BindingCollection](/javascript/api/excel/excel.bindingcollection)|[getCount()](/javascript/api/excel/excel.bindingcollection#getcount--)|Ruft die Anzahl der Bindungen in der Sammlung ab.|
||[getItemOrNullObject(id: string)](/javascript/api/excel/excel.bindingcollection#getitemornullobject-id-)|Ruft ein binding-Objekt anhand seiner ID ab.|
|[ChartCollection](/javascript/api/excel/excel.chartcollection)|[getCount()](/javascript/api/excel/excel.chartcollection#getcount--)|Gibt die Anzahl der Diagramme im Arbeitsblatt zurück.|
||[GetItemOrNullObject(name: string)](/javascript/api/excel/excel.chartcollection#getitemornullobject-name-)|Ruft ein Diagramm über seinen Namen ab.|
|[ChartPointsCollection](/javascript/api/excel/excel.chartpointscollection)|[getCount()](/javascript/api/excel/excel.chartpointscollection#getcount--)|Gibt die Anzahl der Diagrammpunkten in der Reihe zurück.|
|[ChartSeriesCollection](/javascript/api/excel/excel.chartseriescollection)|[getCount()](/javascript/api/excel/excel.chartseriescollection#getcount--)|Gibt die Anzahl der Datenreihen in der Sammlung zurück.|
|[NamedItem](/javascript/api/excel/excel.nameditem)|[comment](/javascript/api/excel/excel.nameditem#comment)|Gibt den Kommentar an, der diesem Namen zugeordnet ist.|
||[delete()](/javascript/api/excel/excel.nameditem#delete--)|Löscht den angegebenen Namen.|
||[getRangeOrNullObject()](/javascript/api/excel/excel.nameditem#getrangeornullobject--)|Ruft das Bereichsobjekt ab, das mit dem Namen verknüpft ist.|
||[scope](/javascript/api/excel/excel.nameditem#scope)|Gibt an, ob der Name auf die Arbeitsmappe oder ein bestimmtes Arbeitsblatt begrenzt ist.|
||[worksheet](/javascript/api/excel/excel.nameditem#worksheet)|Gibt das Arbeitsblatt zurück, auf dessen Bereich das benannte Element beschränkt ist.|
||[worksheetOrNullObject](/javascript/api/excel/excel.nameditem#worksheetornullobject)|Gibt das Arbeitsblatt zurück, auf das das benannte Element begrenzt ist.|
|[NamedItemCollection](/javascript/api/excel/excel.nameditemcollection)|[add(name: string, reference: Range \| string, comment?: string)](/javascript/api/excel/excel.nameditemcollection#add-name--reference--comment-)|Fügt einen neuen Namen zur Auflistung des angegebenen Bereichs hinzu.|
||[addFormulaLocal(name: string, formula: string, comment?: string)](/javascript/api/excel/excel.nameditemcollection#addformulalocal-name--formula--comment-)|Fügt einen neuen Namen zu der Auflistung des angegebenen Bereichs unter Verwendung des Gebietsschemas des Benutzers für die Formel hinzu.|
||[getCount()](/javascript/api/excel/excel.nameditemcollection#getcount--)|Ruft die Anzahl von benannten Elementen in der Auflistung ab.|
||[GetItemOrNullObject(name: string)](/javascript/api/excel/excel.nameditemcollection#getitemornullobject-name-)|Ruft ein `NamedItem` Objekt mit seinem Namen ab.|
|[PivotTableCollection](/javascript/api/excel/excel.pivottablecollection)|[getCount()](/javascript/api/excel/excel.pivottablecollection#getcount--)|Ruft die Anzahl von PivotTables in der Auflistung ab.|
||[GetItemOrNullObject(name: string)](/javascript/api/excel/excel.pivottablecollection#getitemornullobject-name-)|Ruft eine PivotTable anhand des Namens ab.|
|[Range](/javascript/api/excel/excel.range)|[getIntersectionOrNullObject(anotherRange: Range \| string)](/javascript/api/excel/excel.range#getintersectionornullobject-anotherrange-)|Ruft das Bereichsobjekt ab, das die rechteckige Schnittmenge der angegebenen Bereiche darstellt.|
||[getUsedRangeOrNullObject(valuesOnly?: boolean)](/javascript/api/excel/excel.range#getusedrangeornullobject-valuesonly-)|Gibt den verwendeten Bereich des angegebenen Bereichsobjekts zurück.|
|[RangeViewCollection](/javascript/api/excel/excel.rangeviewcollection)|[getCount()](/javascript/api/excel/excel.rangeviewcollection#getcount--)|Ruft die Anzahl der `RangeView` Objekte in der Auflistung ab.|
|[Einstellung](/javascript/api/excel/excel.setting)|[delete()](/javascript/api/excel/excel.setting#delete--)|Löscht die Einstellung.|
||[key](/javascript/api/excel/excel.setting#key)|Der Schlüssel, der die ID der Einstellung darstellt.|
||[value](/javascript/api/excel/excel.setting#value)|Stellt den Wert dar, der für diese Einstellung gespeichert ist.|
|[SettingCollection](/javascript/api/excel/excel.settingcollection)|[add(key: string, value: string \| number \| boolean \| Date Array \| <any> \| any)](/javascript/api/excel/excel.settingcollection#add-key--value-)|Legt die angegebene Einstellung fest oder fügt sie zur Arbeitsmappe hinzu.|
||[getCount()](/javascript/api/excel/excel.settingcollection#getcount--)|Ruft die Anzahl der Einstellungen in der Auflistung ab.|
||[getItem(key: string)](/javascript/api/excel/excel.settingcollection#getitem-key-)|Ruft einen Einstellungseintrag über den Schlüssel ab.|
||[getItemOrNullObject(key: string)](/javascript/api/excel/excel.settingcollection#getitemornullobject-key-)|Ruft einen Einstellungseintrag über den Schlüssel ab.|
||[items](/javascript/api/excel/excel.settingcollection#items)|Ruft die geladenen untergeordneten Elemente in dieser Sammlung ab.|
||[onSettingsChanged](/javascript/api/excel/excel.settingcollection#onsettingschanged)|Tritt auf, wenn die Einstellungen im Dokument geändert werden.|
|[SettingsChangedEventArgs](/javascript/api/excel/excel.settingschangedeventargs)|[Einstellungen](/javascript/api/excel/excel.settingschangedeventargs#settings)|Ruft das `Setting` Objekt ab, das die Bindung darstellt, die das Geänderte Einstellungsereignis ausgelöst hat.|
|[TableCollection](/javascript/api/excel/excel.tablecollection)|[getCount()](/javascript/api/excel/excel.tablecollection#getcount--)|Ruft die Anzahl von Tabellen in der Auflistung ab.|
||[getItemOrNullObject(key: string)](/javascript/api/excel/excel.tablecollection#getitemornullobject-key-)|Ruft eine Tabelle anhand des Namens oder der ID ab.|
|[TableColumnCollection](/javascript/api/excel/excel.tablecolumncollection)|[getCount()](/javascript/api/excel/excel.tablecolumncollection#getcount--)|Ruft die Anzahl der Spalten in der Tabelle ab.|
||[getItemOrNullObject(key: number \| string)](/javascript/api/excel/excel.tablecolumncollection#getitemornullobject-key-)|Ruft ein Spaltenobjekt nach Name oder ID ab.|
|[TableRowCollection](/javascript/api/excel/excel.tablerowcollection)|[getCount()](/javascript/api/excel/excel.tablerowcollection#getcount--)|Ruft die Anzahl der Zeilen in der Tabelle ab.|
|[Workbook](/javascript/api/excel/excel.workbook)|[Einstellungen](/javascript/api/excel/excel.workbook#settings)|Stellt eine Auflistung von Einstellungen dar, die der Arbeitsmappe zugeordnet sind.|
|[Arbeitsblatt](/javascript/api/excel/excel.worksheet)|[getUsedRangeOrNullObject(valuesOnly?: boolean)](/javascript/api/excel/excel.worksheet#getusedrangeornullobject-valuesonly-)|Der verwendete Bereich ist der kleinste Bereich, der mindestens eine der Zellen umfasst, die einen Wert enthalten oder denen eine Formatierung zugewiesen wurde.|
||[names](/javascript/api/excel/excel.worksheet#names)|Auflistung von Namen im Bereich des aktuellen Arbeitsblatts.|
|[WorksheetCollection](/javascript/api/excel/excel.worksheetcollection)|[getCount(visibleOnly?: boolean)](/javascript/api/excel/excel.worksheetcollection#getcount-visibleonly-)|Ruft die Anzahl von Arbeitsblättern in der Auflistung ab.|
||[getItemOrNullObject(key: string)](/javascript/api/excel/excel.worksheetcollection#getitemornullobject-key-)|Ruft das Arbeitsblattobjekt mithilfe des Namens oder der ID ab.|
