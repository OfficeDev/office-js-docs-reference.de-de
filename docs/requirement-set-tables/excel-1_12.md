| Klasse | Felder | Beschreibung |
|:---|:---|:---|
|[ChartAxisTitle](/javascript/api/excel/excel.chartaxistitle)|[textOrientation](/javascript/api/excel/excel.chartaxistitle#textorientation)|Gibt den Winkel an, an dem der Text für den Titel der Diagrammachse ausgerichtet ist.|
|[ChartSeries](/javascript/api/excel/excel.chartseries)|[getDimensionValues(dimension: Excel. ChartSeriesDimension)](/javascript/api/excel/excel.chartseries#getdimensionvalues-dimension-)|Ruft die Werte aus einer einzelnen Dimension der Diagrammreihe ab.|
|[Kommentar](/javascript/api/excel/excel.comment)|[contentType](/javascript/api/excel/excel.comment#contenttype)|Ruft den Inhaltstyp des Kommentars ab.|
|[CommentAddedEventArgs](/javascript/api/excel/excel.commentaddedeventargs)|[commentDetails](/javascript/api/excel/excel.commentaddedeventargs#commentdetails)|Ruft das `CommentDetail` Array ab, das die Kommentar-ID und die IDs der zugehörigen Antworten enthält.|
||[source](/javascript/api/excel/excel.commentaddedeventargs#source)|Gibt die Quelle des Ereignisses an.|
||[type](/javascript/api/excel/excel.commentaddedeventargs#type)|Ruft den Typ des Ereignisses ab.|
||[worksheetId](/javascript/api/excel/excel.commentaddedeventargs#worksheetid)|Ruft die ID des Arbeitsblatts ab, in dem das Ereignis aufgetreten ist.|
|[CommentChangedEventArgs](/javascript/api/excel/excel.commentchangedeventargs)|[changeType](/javascript/api/excel/excel.commentchangedeventargs#changetype)|Ruft den Änderungstyp ab, der angibt, wie das geänderte Ereignis ausgelöst wird.|
||[commentDetails](/javascript/api/excel/excel.commentchangedeventargs#commentdetails)|Ruft das `CommentDetail` Array ab, das die Kommentar-ID und die IDs der zugehörigen Antworten enthält.|
||[source](/javascript/api/excel/excel.commentchangedeventargs#source)|Gibt die Quelle des Ereignisses an.|
||[type](/javascript/api/excel/excel.commentchangedeventargs#type)|Ruft den Typ des Ereignisses ab.|
||[worksheetId](/javascript/api/excel/excel.commentchangedeventargs#worksheetid)|Ruft die ID des Arbeitsblatts ab, in dem das Ereignis aufgetreten ist.|
|[CommentCollection](/javascript/api/excel/excel.commentcollection)|[onAdded](/javascript/api/excel/excel.commentcollection#onadded)|Tritt auf, wenn die Kommentare hinzugefügt werden.|
||[onChanged](/javascript/api/excel/excel.commentcollection#onchanged)|Tritt auf, wenn Kommentare oder Antworten in einer Kommentarsammlung geändert werden, z. B. wenn Antworten gelöscht werden.|
||[onDeleted](/javascript/api/excel/excel.commentcollection#ondeleted)|Tritt auf, wenn Kommentare in der Kommentarsammlung gelöscht werden.|
|[CommentDeletedEventArgs](/javascript/api/excel/excel.commentdeletedeventargs)|[commentDetails](/javascript/api/excel/excel.commentdeletedeventargs#commentdetails)|Ruft das `CommentDetail` Array ab, das die Kommentar-ID und die IDs der zugehörigen Antworten enthält.|
||[source](/javascript/api/excel/excel.commentdeletedeventargs#source)|Gibt die Quelle des Ereignisses an.|
||[type](/javascript/api/excel/excel.commentdeletedeventargs#type)|Ruft den Typ des Ereignisses ab.|
||[worksheetId](/javascript/api/excel/excel.commentdeletedeventargs#worksheetid)|Ruft die ID des Arbeitsblatts ab, in dem das Ereignis aufgetreten ist.|
|[CommentDetail](/javascript/api/excel/excel.commentdetail)|[commentId](/javascript/api/excel/excel.commentdetail#commentid)|Stellt die ID des Kommentars dar.|
||[replyIds](/javascript/api/excel/excel.commentdetail#replyids)|Stellt die IDs der zugehörigen Antworten dar, die zum Kommentar gehören.|
|[CommentReply](/javascript/api/excel/excel.commentreply)|[contentType](/javascript/api/excel/excel.commentreply#contenttype)|Der Inhaltstyp der Antwort.|
|[CultureInfo](/javascript/api/excel/excel.cultureinfo)|[Datetimeformat](/javascript/api/excel/excel.cultureinfo#datetimeformat)|Definiert das kulturell geeignete Format für die Anzeige von Datum und Uhrzeit.|
|[DatetimeFormatInfo](/javascript/api/excel/excel.datetimeformatinfo)|[dateSeparator](/javascript/api/excel/excel.datetimeformatinfo#dateseparator)|Ruft die Zeichenfolge ab, die als Datumstrennzeichen verwendet wird.|
||[Longdatepattern](/javascript/api/excel/excel.datetimeformatinfo#longdatepattern)|Ruft die Formatzeichenfolge für einen langen Datumswert ab.|
||[longTimePattern](/javascript/api/excel/excel.datetimeformatinfo#longtimepattern)|Ruft die Formatzeichenfolge für einen langen Zeitwert ab.|
||[Shortdatepattern](/javascript/api/excel/excel.datetimeformatinfo#shortdatepattern)|Ruft die Formatzeichenfolge für einen kurzen Datumswert ab.|
||[timeSeparator](/javascript/api/excel/excel.datetimeformatinfo#timeseparator)|Ruft die Zeichenfolge ab, die als Zeittrennzeichen verwendet wird.|
|[PivotDateFilter](/javascript/api/excel/excel.pivotdatefilter)|[Komparator](/javascript/api/excel/excel.pivotdatefilter#comparator)|Der Vergleichswert ist der statische Wert, mit dem andere Werte verglichen werden.|
||[Bedingung](/javascript/api/excel/excel.pivotdatefilter#condition)|Gibt die Bedingung für den Filter an, der die erforderlichen Filterkriterien definiert.|
||[Exklusive](/javascript/api/excel/excel.pivotdatefilter#exclusive)|Wenn `true` , filter *excludes* items that meet criteria.|
||[Lowerbound](/javascript/api/excel/excel.pivotdatefilter#lowerbound)|Die untere Grenze des Bereichs für die `between` Filterbedingung.|
||[upperBound](/javascript/api/excel/excel.pivotdatefilter#upperbound)|Die obere Grenze des Bereichs für die `between` Filterbedingung.|
||[wholeDays](/javascript/api/excel/excel.pivotdatefilter#wholedays)|Gibt für `equals` , `before` und `after` `between` Filterbedingungen an, ob Vergleiche als ganze Tage durchgeführt werden sollen.|
|[PivotField](/javascript/api/excel/excel.pivotfield)|[applyFilter(filter: Excel. PivotFilters)](/javascript/api/excel/excel.pivotfield#applyfilter-filter-)|Legt einen oder mehrere der aktuellen PivotFilter des Felds fest und wendet sie auf das Feld an.|
||[clearAllFilters()](/javascript/api/excel/excel.pivotfield#clearallfilters--)|Löscht alle Kriterien aus allen Filtern des Felds.|
||[clearFilter(filterType: Excel. PivotFilterType)](/javascript/api/excel/excel.pivotfield#clearfilter-filtertype-)|Löscht alle vorhandenen Kriterien aus dem Feldfilter des angegebenen Typs (sofern derzeit eines angewendet wird).|
||[getFilters()](/javascript/api/excel/excel.pivotfield#getfilters--)|Ruft alle Filter ab, die derzeit auf das Feld angewendet werden.|
||[isFiltered(filterType?: Excel. PivotFilterType)](/javascript/api/excel/excel.pivotfield#isfiltered-filtertype-)|Überprüft, ob auf das Feld angewendete Filter vorhanden sind.|
|[PivotFilters](/javascript/api/excel/excel.pivotfilters)|[dateFilter](/javascript/api/excel/excel.pivotfilters#datefilter)|Der aktuell angewendete Datumsfilter des PivotFields.|
||[labelFilter](/javascript/api/excel/excel.pivotfilters#labelfilter)|Der aktuell angewendete Bezeichnungsfilter des PivotFields.|
||[manualFilter](/javascript/api/excel/excel.pivotfilters#manualfilter)|Der aktuell angewendete manuelle Filter des PivotFields.|
||[valueFilter](/javascript/api/excel/excel.pivotfilters#valuefilter)|Der aktuell angewendete Wertfilter des PivotFields.|
|[PivotLabelFilter](/javascript/api/excel/excel.pivotlabelfilter)|[Komparator](/javascript/api/excel/excel.pivotlabelfilter#comparator)|Der Vergleichswert ist der statische Wert, mit dem andere Werte verglichen werden.|
||[Bedingung](/javascript/api/excel/excel.pivotlabelfilter#condition)|Gibt die Bedingung für den Filter an, die die erforderlichen Filterkriterien definiert.|
||[Exklusive](/javascript/api/excel/excel.pivotlabelfilter#exclusive)|Wenn `true` , filter *excludes* items that meet criteria.|
||[Lowerbound](/javascript/api/excel/excel.pivotlabelfilter#lowerbound)|Die untere Grenze des Bereichs für die `between` Filterbedingung.|
||[substring](/javascript/api/excel/excel.pivotlabelfilter#substring)|Die teilzeichenfolge für `beginsWith` `endsWith` , und `contains` Filterbedingungen verwendet.|
||[upperBound](/javascript/api/excel/excel.pivotlabelfilter#upperbound)|Die obere Grenze des Bereichs für die `between` Filterbedingung.|
|[PivotManualFilter](/javascript/api/excel/excel.pivotmanualfilter)|[selectedItems](/javascript/api/excel/excel.pivotmanualfilter#selecteditems)|Eine Liste der ausgewählten Elemente, die manuell gefiltert werden sollen.|
|[PivotTable](/javascript/api/excel/excel.pivottable)|[allowMultipleFiltersPerField](/javascript/api/excel/excel.pivottable#allowmultiplefiltersperfield)|Gibt an, ob die PivotTable die Anwendung mehrerer PivotFilter auf einem bestimmten PivotField in der Tabelle zulässt.|
|[PivotTableScopedCollection](/javascript/api/excel/excel.pivottablescopedcollection)|[getCount()](/javascript/api/excel/excel.pivottablescopedcollection#getcount--)|Ruft die Anzahl der PivotTables in der Auflistung ab.|
||[getFirst()](/javascript/api/excel/excel.pivottablescopedcollection#getfirst--)|Ruft die erste PivotTable in der Auflistung ab.|
||[getItem(key: string)](/javascript/api/excel/excel.pivottablescopedcollection#getitem-key-)|Ruft eine PivotTable anhand des Namens ab.|
||[GetItemOrNullObject(name: string)](/javascript/api/excel/excel.pivottablescopedcollection#getitemornullobject-name-)|Ruft eine PivotTable anhand des Namens ab.|
||[items](/javascript/api/excel/excel.pivottablescopedcollection#items)|Ruft die geladenen untergeordneten Elemente in dieser Sammlung ab.|
|[PivotValueFilter](/javascript/api/excel/excel.pivotvaluefilter)|[Komparator](/javascript/api/excel/excel.pivotvaluefilter#comparator)|Der Vergleichswert ist der statische Wert, mit dem andere Werte verglichen werden.|
||[Bedingung](/javascript/api/excel/excel.pivotvaluefilter#condition)|Gibt die Bedingung für den Filter an, die die erforderlichen Filterkriterien definiert.|
||[Exklusive](/javascript/api/excel/excel.pivotvaluefilter#exclusive)|Wenn `true` , filter *excludes* items that meet criteria.|
||[Lowerbound](/javascript/api/excel/excel.pivotvaluefilter#lowerbound)|Die untere Grenze des Bereichs für die `between` Filterbedingung.|
||[selectionType](/javascript/api/excel/excel.pivotvaluefilter#selectiontype)|Gibt an, ob der Filter für die top/bottom N-Elemente, top/bottom N percent oder top/bottom N sum ist.|
||[Schwelle](/javascript/api/excel/excel.pivotvaluefilter#threshold)|Die N-Schwellenwertanzahl der Elemente, Prozent oder Summe, die nach einer Filterbedingung oben/unten gefiltert werden sollen.|
||[upperBound](/javascript/api/excel/excel.pivotvaluefilter#upperbound)|Die obere Grenze des Bereichs für die `between` Filterbedingung.|
||[value](/javascript/api/excel/excel.pivotvaluefilter#value)|Name des ausgewählten "Werts" im Feld, nach dem gefiltert werden soll.|
|[Range](/javascript/api/excel/excel.range)|[getDirectPrecedents()](/javascript/api/excel/excel.range#getdirectprecedents--)|Gibt ein `WorkbookRangeAreas` Objekt zurück, das den Bereich darstellt, der alle direkten Vorgänger einer Zelle im selben Arbeitsblatt oder in mehreren Arbeitsblättern enthält.|
||[getPivotTables(fullyContained?: boolean)](/javascript/api/excel/excel.range#getpivottables-fullycontained-)|Ruft eine bereichsbezogene Auflistung von PivotTables ab, die sich mit dem Bereich überlappen.|
||[getSpillParent()](/javascript/api/excel/excel.range#getspillparent--)|Ruft das Bereichsobjekt ab, das die Ankerzelle für eine Zelle enthält, in die ein Überlauf erfolgen kann.|
||[getSpillParentOrNullObject()](/javascript/api/excel/excel.range#getspillparentornullobject--)|Ruft das Bereichsobjekt ab, das die Ankerzelle für die Zelle enthält, in die überlauft wird.|
||[getSpillingToRange()](/javascript/api/excel/excel.range#getspillingtorange--)|Ruft beim Aufruf für eine Ankerzelle das Bereichsobjekt ab, das den Überlaufbereich enthält.|
||[getSpillingToRangeOrNullObject()](/javascript/api/excel/excel.range#getspillingtorangeornullobject--)|Ruft beim Aufruf für eine Ankerzelle das Bereichsobjekt ab, das den Überlaufbereich enthält.|
||[hasSpill](/javascript/api/excel/excel.range#hasspill)|Stellt dar, ob alle Zellen einen Überlaufrahmen aufweisen.|
||[numberFormatCategories](/javascript/api/excel/excel.range#numberformatcategories)|Stellt die Kategorie des Zahlenformats jeder Zelle dar.|
||[savedAsArray](/javascript/api/excel/excel.range#savedasarray)|Gibt an, ob alle Zellen als Arrayformel gespeichert würden.|
|[RangeAreasCollection](/javascript/api/excel/excel.rangeareascollection)|[getCount()](/javascript/api/excel/excel.rangeareascollection#getcount--)|Ruft die Anzahl der `RangeAreas` Objekte in dieser Auflistung ab.|
||[getItemAt(index: number)](/javascript/api/excel/excel.rangeareascollection#getitemat-index-)|Gibt das `RangeAreas` Objekt basierend auf der Position in der Auflistung zurück.|
||[items](/javascript/api/excel/excel.rangeareascollection#items)|Ruft die geladenen untergeordneten Elemente in dieser Sammlung ab.|
|[WorkbookRangeAreas](/javascript/api/excel/excel.workbookrangeareas)|[getRangeAreasBySheet(key: string)](/javascript/api/excel/excel.workbookrangeareas#getrangeareasbysheet-key-)|Gibt das Objekt basierend auf der `RangeAreas` Arbeitsblatt-ID oder dem Namen in der Auflistung zurück.|
||[getRangeAreasOrNullObjectBySheet(key: string)](/javascript/api/excel/excel.workbookrangeareas#getrangeareasornullobjectbysheet-key-)|Gibt das `RangeAreas` Objekt basierend auf dem Arbeitsblattnamen oder der ID in der Auflistung zurück.|
||[addresses](/javascript/api/excel/excel.workbookrangeareas#addresses)|Gibt ein Array von Adressen im A1-Format zurück.|
||[areas](/javascript/api/excel/excel.workbookrangeareas#areas)|Gibt das `RangeAreasCollection` Objekt zurück.|
||[Bereiche](/javascript/api/excel/excel.workbookrangeareas#ranges)|Gibt Bereiche zurück, aus denen dieses Objekt in einem `RangeCollection` Objekt besteht.|
|[Arbeitsblatt](/javascript/api/excel/excel.worksheet)|[Customproperties](/javascript/api/excel/excel.worksheet#customproperties)|Ruft eine Auflistung benutzerdefinierter Eigenschaften auf Arbeitsblattebene ab.|
|[WorksheetCustomProperty](/javascript/api/excel/excel.worksheetcustomproperty)|[delete()](/javascript/api/excel/excel.worksheetcustomproperty#delete--)|Löscht die benutzerdefinierte Eigenschaft.|
||[key](/javascript/api/excel/excel.worksheetcustomproperty#key)|Ruft den Schlüssel der benutzerdefinierten Eigenschaft ab.|
||[value](/javascript/api/excel/excel.worksheetcustomproperty#value)|Ruft den Wert der benutzerdefinierten Eigenschaft ab oder legt ihn fest.|
|[WorksheetCustomPropertyCollection](/javascript/api/excel/excel.worksheetcustompropertycollection)|[add(key: string, value: string)](/javascript/api/excel/excel.worksheetcustompropertycollection#add-key--value-)|Fügt eine neue benutzerdefinierte Eigenschaft hinzu, die dem bereitgestellten Schlüssel zugeordnet ist.|
||[getCount()](/javascript/api/excel/excel.worksheetcustompropertycollection#getcount--)|Ruft die Anzahl der benutzerdefinierten Eigenschaften in diesem Arbeitsblatt ab.|
||[getItem(key: string)](/javascript/api/excel/excel.worksheetcustompropertycollection#getitem-key-)|Ruft ein Objekt für benutzerdefinierte Eigenschaften über seinen Schlüssel ab, bei dem Groß-/Kleinschreibung nicht beachtet wird.|
||[getItemOrNullObject(key: string)](/javascript/api/excel/excel.worksheetcustompropertycollection#getitemornullobject-key-)|Ruft ein Objekt für benutzerdefinierte Eigenschaften über seinen Schlüssel ab, bei dem Groß-/Kleinschreibung nicht beachtet wird.|
||[items](/javascript/api/excel/excel.worksheetcustompropertycollection#items)|Ruft die geladenen untergeordneten Elemente in dieser Sammlung ab.|
