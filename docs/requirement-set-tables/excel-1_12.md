| Klasse | Felder | Beschreibung |
|:---|:---|:---|
|[ChartAxisTitle](/javascript/api/excel/excel.chartaxistitle)|[textOrientation](/javascript/api/excel/excel.chartaxistitle#textorientation)|Gibt den Winkel an, an dem der Text für den Titel der Diagrammachse ausgerichtet ist.|
|[ChartSeries](/javascript/api/excel/excel.chartseries)|[getDimensionValues(dimension: Excel.ChartSeriesDimension)](/javascript/api/excel/excel.chartseries#getdimensionvalues-dimension-)|Ruft die Werte aus einer einzelnen Dimension der Diagrammreihe ab.|
|[Kommentar](/javascript/api/excel/excel.comment)|[contentType](/javascript/api/excel/excel.comment#contenttype)|Ruft den Inhaltstyp des Kommentars ab.|
|[CommentAddedEventArgs](/javascript/api/excel/excel.commentaddedeventargs)|[commentDetails](/javascript/api/excel/excel.commentaddedeventargs#commentdetails)|Ruft das `CommentDetail` Array ab, das die Kommentar-ID und IDs der zugehörigen Antworten enthält.|
||[source](/javascript/api/excel/excel.commentaddedeventargs#source)|Gibt die Quelle des Ereignisses an.|
||[type](/javascript/api/excel/excel.commentaddedeventargs#type)|Ruft den Typ des Ereignisses ab.|
||[worksheetId](/javascript/api/excel/excel.commentaddedeventargs#worksheetid)|Ruft die ID des Arbeitsblatts ab, in dem das Ereignis passiert ist.|
|[CommentChangedEventArgs](/javascript/api/excel/excel.commentchangedeventargs)|[changeType](/javascript/api/excel/excel.commentchangedeventargs#changetype)|Ruft den Änderungstyp ab, der darstellt, wie das geänderte Ereignis ausgelöst wird.|
||[commentDetails](/javascript/api/excel/excel.commentchangedeventargs#commentdetails)|Get the `CommentDetail` array which contains the comment ID and IDs of its related replies.|
||[source](/javascript/api/excel/excel.commentchangedeventargs#source)|Gibt die Quelle des Ereignisses an.|
||[type](/javascript/api/excel/excel.commentchangedeventargs#type)|Ruft den Typ des Ereignisses ab.|
||[worksheetId](/javascript/api/excel/excel.commentchangedeventargs#worksheetid)|Ruft die ID des Arbeitsblatts ab, in dem das Ereignis passiert ist.|
|[CommentCollection](/javascript/api/excel/excel.commentcollection)|[onAdded](/javascript/api/excel/excel.commentcollection#onadded)|Tritt auf, wenn die Kommentare hinzugefügt werden.|
||[onChanged](/javascript/api/excel/excel.commentcollection#onchanged)|Tritt auf, wenn Kommentare oder Antworten in einer Kommentarsammlung geändert werden, auch wenn Antworten gelöscht werden.|
||[onDeleted](/javascript/api/excel/excel.commentcollection#ondeleted)|Tritt auf, wenn Kommentare in der Kommentarsammlung gelöscht werden.|
|[CommentDeletedEventArgs](/javascript/api/excel/excel.commentdeletedeventargs)|[commentDetails](/javascript/api/excel/excel.commentdeletedeventargs#commentdetails)|Ruft das `CommentDetail` Array ab, das die Kommentar-ID und IDs der zugehörigen Antworten enthält.|
||[source](/javascript/api/excel/excel.commentdeletedeventargs#source)|Gibt die Quelle des Ereignisses an.|
||[type](/javascript/api/excel/excel.commentdeletedeventargs#type)|Ruft den Typ des Ereignisses ab.|
||[worksheetId](/javascript/api/excel/excel.commentdeletedeventargs#worksheetid)|Ruft die ID des Arbeitsblatts ab, in dem das Ereignis passiert ist.|
|[CommentDetail](/javascript/api/excel/excel.commentdetail)|[commentId](/javascript/api/excel/excel.commentdetail#commentid)|Stellt die ID des Kommentars dar.|
||[replyIds](/javascript/api/excel/excel.commentdetail#replyids)|Stellt die IDs der zugehörigen Antworten dar, die zum Kommentar gehören.|
|[CommentReply](/javascript/api/excel/excel.commentreply)|[contentType](/javascript/api/excel/excel.commentreply#contenttype)|Der Inhaltstyp der Antwort.|
|[CultureInfo](/javascript/api/excel/excel.cultureinfo)|[datetimeFormat](/javascript/api/excel/excel.cultureinfo#datetimeformat)|Definiert das kulturell geeignete Format der Anzeige von Datum und Uhrzeit.|
|[DatetimeFormatInfo](/javascript/api/excel/excel.datetimeformatinfo)|[dateSeparator](/javascript/api/excel/excel.datetimeformatinfo#dateseparator)|Ruft die Zeichenfolge ab, die als Datumstrennzeichen verwendet wird.|
||[longDatePattern](/javascript/api/excel/excel.datetimeformatinfo#longdatepattern)|Ruft die Formatzeichenfolge für einen langen Datumswert ab.|
||[longTimePattern](/javascript/api/excel/excel.datetimeformatinfo#longtimepattern)|Ruft die Formatzeichenfolge für einen langfristigen Wert ab.|
||[shortDatePattern](/javascript/api/excel/excel.datetimeformatinfo#shortdatepattern)|Ruft die Formatzeichenfolge für einen kurzen Datumswert ab.|
||[timeSeparator](/javascript/api/excel/excel.datetimeformatinfo#timeseparator)|Ruft die Zeichenfolge ab, die als Zeittrennzeichen verwendet wird.|
|[PivotDateFilter](/javascript/api/excel/excel.pivotdatefilter)|[Comparator](/javascript/api/excel/excel.pivotdatefilter#comparator)|Der Vergleichswert ist der statische Wert, mit dem andere Werte verglichen werden.|
||[Bedingung](/javascript/api/excel/excel.pivotdatefilter#condition)|Gibt die Bedingung für den Filter an, der die erforderlichen Filterkriterien definiert.|
||[exclusive](/javascript/api/excel/excel.pivotdatefilter#exclusive)|Wenn `true` , filter schließt *Elemente* aus, die Kriterien erfüllen.|
||[lowerBound](/javascript/api/excel/excel.pivotdatefilter#lowerbound)|Die Untergrenze des Bereichs für die `between` Filterbedingung.|
||[upperBound](/javascript/api/excel/excel.pivotdatefilter#upperbound)|Die Obere Grenze des Bereichs für die `between` Filterbedingung.|
||[wholeDays](/javascript/api/excel/excel.pivotdatefilter#wholedays)|Für `equals` , , und `before` `after` Filterbedingungen gibt an, ob Vergleiche als ganze `between` Tage vorgenommen werden sollen.|
|[PivotField](/javascript/api/excel/excel.pivotfield)|[applyFilter(filter: Excel.PivotFilters)](/javascript/api/excel/excel.pivotfield#applyfilter-filter-)|Legt einen oder mehrere der aktuellen PivotFilter des Felds fest und wendet diese auf das Feld an.|
||[clearAllFilters()](/javascript/api/excel/excel.pivotfield#clearallfilters--)|Entfernt alle Kriterien aus allen Filtern des Felds.|
||[clearFilter(filterType: Excel.PivotFilterType)](/javascript/api/excel/excel.pivotfield#clearfilter-filtertype-)|Entfernt alle vorhandenen Kriterien aus dem Filter des Felds des angegebenen Typs (sofern aktuell angewendet).|
||[getFilters()](/javascript/api/excel/excel.pivotfield#getfilters--)|Ruft alle Filter ab, die derzeit auf das Feld angewendet werden.|
||[isFiltered(filterType?: Excel.PivotFilterType)](/javascript/api/excel/excel.pivotfield#isfiltered-filtertype-)|Überprüft, ob für das Feld Filter angewendet wurden.|
|[PivotFilters](/javascript/api/excel/excel.pivotfilters)|[dateFilter](/javascript/api/excel/excel.pivotfilters#datefilter)|Der aktuell angewendete Datumsfilter des PivotFields.|
||[labelFilter](/javascript/api/excel/excel.pivotfilters#labelfilter)|Der aktuell angewendete Bezeichnungsfilter des PivotFields.|
||[manualFilter](/javascript/api/excel/excel.pivotfilters#manualfilter)|Der aktuell angewendete manuelle Filter des PivotFields.|
||[valueFilter](/javascript/api/excel/excel.pivotfilters#valuefilter)|Der aktuell angewendete Wertfilter des PivotFields.|
|[PivotLabelFilter](/javascript/api/excel/excel.pivotlabelfilter)|[Comparator](/javascript/api/excel/excel.pivotlabelfilter#comparator)|Der Vergleichswert ist der statische Wert, mit dem andere Werte verglichen werden.|
||[Bedingung](/javascript/api/excel/excel.pivotlabelfilter#condition)|Gibt die Bedingung für den Filter an, der die erforderlichen Filterkriterien definiert.|
||[exclusive](/javascript/api/excel/excel.pivotlabelfilter#exclusive)|Wenn `true` , filter schließt *Elemente* aus, die Kriterien erfüllen.|
||[lowerBound](/javascript/api/excel/excel.pivotlabelfilter#lowerbound)|Die Untergrenze des Bereichs für die `between` Filterbedingung.|
||[substring](/javascript/api/excel/excel.pivotlabelfilter#substring)|Die Unterzeichenfolge, die für `beginsWith` , `endsWith` und `contains` Filterbedingungen verwendet wird.|
||[upperBound](/javascript/api/excel/excel.pivotlabelfilter#upperbound)|Die Obere Grenze des Bereichs für die `between` Filterbedingung.|
|[PivotManualFilter](/javascript/api/excel/excel.pivotmanualfilter)|[selectedItems](/javascript/api/excel/excel.pivotmanualfilter#selecteditems)|Eine Liste der ausgewählten Elemente, die manuell gefiltert werden müssen.|
|[PivotTable](/javascript/api/excel/excel.pivottable)|[allowMultipleFiltersPerField](/javascript/api/excel/excel.pivottable#allowmultiplefiltersperfield)|Gibt an, ob die PivotTable die Anwendung mehrerer PivotFilter auf ein bestimmtes PivotField in der Tabelle zulässt.|
|[PivotTableScopedCollection](/javascript/api/excel/excel.pivottablescopedcollection)|[getCount()](/javascript/api/excel/excel.pivottablescopedcollection#getcount--)|Ruft die Anzahl der PivotTables in der Auflistung ab.|
||[getFirst()](/javascript/api/excel/excel.pivottablescopedcollection#getfirst--)|Ruft die erste PivotTable in der Auflistung ab.|
||[getItem(key: string)](/javascript/api/excel/excel.pivottablescopedcollection#getitem-key-)|Ruft eine PivotTable anhand des Namens ab.|
||[GetItemOrNullObject(name: string)](/javascript/api/excel/excel.pivottablescopedcollection#getitemornullobject-name-)|Ruft eine PivotTable anhand des Namens ab.|
||[items](/javascript/api/excel/excel.pivottablescopedcollection#items)|Ruft die geladenen untergeordneten Elemente in dieser Sammlung ab.|
|[PivotValueFilter](/javascript/api/excel/excel.pivotvaluefilter)|[Comparator](/javascript/api/excel/excel.pivotvaluefilter#comparator)|Der Vergleichswert ist der statische Wert, mit dem andere Werte verglichen werden.|
||[Bedingung](/javascript/api/excel/excel.pivotvaluefilter#condition)|Gibt die Bedingung für den Filter an, der die erforderlichen Filterkriterien definiert.|
||[exclusive](/javascript/api/excel/excel.pivotvaluefilter#exclusive)|Wenn `true` , filter schließt *Elemente* aus, die Kriterien erfüllen.|
||[lowerBound](/javascript/api/excel/excel.pivotvaluefilter#lowerbound)|Die Untergrenze des Bereichs für die `between` Filterbedingung.|
||[selectionType](/javascript/api/excel/excel.pivotvaluefilter#selectiontype)|Gibt an, ob der Filter für die N-Elemente oben/unten, oben/unten N Prozent oder oben/unten N-Summe gilt.|
||[threshold](/javascript/api/excel/excel.pivotvaluefilter#threshold)|Die "N"-Schwellenwertanzahl von Elementen, Prozent oder Summen, die nach einer Filterbedingung vom oberen/unteren Rand gefiltert werden sollen.|
||[upperBound](/javascript/api/excel/excel.pivotvaluefilter#upperbound)|Die Obere Grenze des Bereichs für die `between` Filterbedingung.|
||[value](/javascript/api/excel/excel.pivotvaluefilter#value)|Name des ausgewählten Werts in dem Feld, nach dem gefiltert werden soll.|
|[Range](/javascript/api/excel/excel.range)|[getDirectPrecedents()](/javascript/api/excel/excel.range#getdirectprecedents--)|Gibt ein Objekt zurück, das den Bereich darstellt, der alle direkten Vorgänger einer Zelle im gleichen Arbeitsblatt oder `WorkbookRangeAreas` in mehreren Arbeitsblättern enthält.|
||[getPivotTables(fullyContained?: boolean)](/javascript/api/excel/excel.range#getpivottables-fullycontained-)|Ruft eine Bereichssammlung von PivotTables ab, die mit dem Bereich überlappen.|
||[getSpillParent()](/javascript/api/excel/excel.range#getspillparent--)|Ruft das Bereichsobjekt ab, das die Ankerzelle für eine Zelle enthält, in die ein Überlauf erfolgen kann.|
||[getSpillParentOrNullObject()](/javascript/api/excel/excel.range#getspillparentornullobject--)|Ruft das Range-Objekt ab, das die Ankerzelle für die Zelle enthält, in die übergelaufen wird.|
||[getSpillingToRange()](/javascript/api/excel/excel.range#getspillingtorange--)|Ruft beim Aufruf für eine Ankerzelle das Bereichsobjekt ab, das den Überlaufbereich enthält.|
||[getSpillingToRangeOrNullObject()](/javascript/api/excel/excel.range#getspillingtorangeornullobject--)|Ruft beim Aufruf für eine Ankerzelle das Bereichsobjekt ab, das den Überlaufbereich enthält.|
||[hasSpill](/javascript/api/excel/excel.range#hasspill)|Stellt dar, ob alle Zellen einen Überlaufrahmen aufweisen.|
||[numberFormatCategories](/javascript/api/excel/excel.range#numberformatcategories)|Stellt die Kategorie des Zahlenformats jeder Zelle dar.|
||[savedAsArray](/javascript/api/excel/excel.range#savedasarray)|Stellt dar, ob alle Zellen als Arrayformel gespeichert werden.|
|[RangeAreasCollection](/javascript/api/excel/excel.rangeareascollection)|[getCount()](/javascript/api/excel/excel.rangeareascollection#getcount--)|Ruft die Anzahl der `RangeAreas` Objekte in dieser Auflistung ab.|
||[getItemAt(index: number)](/javascript/api/excel/excel.rangeareascollection#getitemat-index-)|Gibt das `RangeAreas` Objekt basierend auf der Position in der Auflistung zurück.|
||[items](/javascript/api/excel/excel.rangeareascollection#items)|Ruft die geladenen untergeordneten Elemente in dieser Sammlung ab.|
|[WorkbookRangeAreas](/javascript/api/excel/excel.workbookrangeareas)|[getRangeAreasBySheet(key: string)](/javascript/api/excel/excel.workbookrangeareas#getrangeareasbysheet-key-)|Gibt das `RangeAreas` Objekt basierend auf der Arbeitsblatt-ID oder dem Namen in der Auflistung zurück.|
||[getRangeAreasOrNullObjectBySheet(key: string)](/javascript/api/excel/excel.workbookrangeareas#getrangeareasornullobjectbysheet-key-)|Gibt das `RangeAreas` Objekt basierend auf dem Arbeitsblattnamen oder der ID in der Auflistung zurück.|
||[addresses](/javascript/api/excel/excel.workbookrangeareas#addresses)|Gibt ein Array von Adressen im A1-Format zurück.|
||[areas](/javascript/api/excel/excel.workbookrangeareas#areas)|Gibt das Objekt `RangeAreasCollection` zurück.|
||[Bereiche](/javascript/api/excel/excel.workbookrangeareas#ranges)|Gibt Bereiche zurück, die dieses Objekt in einem Objekt  `RangeCollection`   enthalten.|
|[Arbeitsblatt](/javascript/api/excel/excel.worksheet)|[customProperties](/javascript/api/excel/excel.worksheet#customproperties)|Ruft eine Auflistung benutzerdefinierter Eigenschaften auf Arbeitsblattebene ab.|
|[WorksheetCustomProperty](/javascript/api/excel/excel.worksheetcustomproperty)|[delete()](/javascript/api/excel/excel.worksheetcustomproperty#delete--)|Löscht die benutzerdefinierte Eigenschaft.|
||[key](/javascript/api/excel/excel.worksheetcustomproperty#key)|Ruft den Schlüssel der benutzerdefinierten Eigenschaft ab.|
||[value](/javascript/api/excel/excel.worksheetcustomproperty#value)|Ruft den Wert der benutzerdefinierten Eigenschaft ab oder legt ihn fest.|
|[WorksheetCustomPropertyCollection](/javascript/api/excel/excel.worksheetcustompropertycollection)|[add(key: string, value: string)](/javascript/api/excel/excel.worksheetcustompropertycollection#add-key--value-)|Fügt eine neue benutzerdefinierte Eigenschaft hinzu, die dem bereitgestellten Schlüssel zufügt.|
||[getCount()](/javascript/api/excel/excel.worksheetcustompropertycollection#getcount--)|Ruft die Anzahl der benutzerdefinierten Eigenschaften in diesem Arbeitsblatt ab.|
||[getItem(key: string)](/javascript/api/excel/excel.worksheetcustompropertycollection#getitem-key-)|Ruft ein Objekt für benutzerdefinierte Eigenschaften über seinen Schlüssel ab, bei dem Groß-/Kleinschreibung nicht beachtet wird.|
||[getItemOrNullObject(key: string)](/javascript/api/excel/excel.worksheetcustompropertycollection#getitemornullobject-key-)|Ruft ein Objekt für benutzerdefinierte Eigenschaften über seinen Schlüssel ab, bei dem Groß-/Kleinschreibung nicht beachtet wird.|
||[items](/javascript/api/excel/excel.worksheetcustompropertycollection#items)|Ruft die geladenen untergeordneten Elemente in dieser Sammlung ab.|
