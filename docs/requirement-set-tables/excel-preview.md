| Klasse | Felder | Beschreibung |
|:---|:---|:---|
|[Kommentar](/javascript/api/excel/excel.comment)|[assignTask(email: string)](/javascript/api/excel/excel.comment#assigntask-email-)|Weist die dem Kommentar zugeordnete Aufgabe dem angegebenen Benutzer als alleinigen Zugewiesenen zu.|
||[getTask()](/javascript/api/excel/excel.comment#gettask--)|Ruft den Vorgang ab, der diesem Kommentar zugeordnet ist.|
||[getTaskOrNullObject()](/javascript/api/excel/excel.comment#gettaskornullobject--)|Ruft den Vorgang ab, der diesem Kommentar zugeordnet ist.|
|[CommentReply](/javascript/api/excel/excel.commentreply)|[assignTask(email: string)](/javascript/api/excel/excel.commentreply#assigntask-email-)|Weist die dem Kommentar zugeordnete Aufgabe dem angegebenen Benutzer als alleinigen Zugewiesenen zu.|
||[getTask()](/javascript/api/excel/excel.commentreply#gettask--)|Ruft den Vorgang ab, der diesem Kommentar zugeordnet ist.|
||[getTaskOrNullObject()](/javascript/api/excel/excel.commentreply#gettaskornullobject--)|Ruft den Vorgang ab, der diesem Kommentar zugeordnet ist.|
|[FormulaChangedEventDetail](/javascript/api/excel/excel.formulachangedeventdetail)|[cellAddress](/javascript/api/excel/excel.formulachangedeventdetail#celladdress)|Die Adresse der Zelle, die die geänderte Formel enthält.|
||[previousFormula](/javascript/api/excel/excel.formulachangedeventdetail#previousformula)|Stellt die vorherige Formel dar, bevor sie geändert wurde.|
|[LinkedDataType](/javascript/api/excel/excel.linkeddatatype)|[dataProvider](/javascript/api/excel/excel.linkeddatatype#dataprovider)|Der Name des Datenanbieters für den verknüpften Datentyp.|
||[lastRefreshed](/javascript/api/excel/excel.linkeddatatype#lastrefreshed)|Datum und Uhrzeit der lokalen Zeitzone seit dem Öffnen der Arbeitsmappe beim letzten Aktualisieren des verknüpften Datentyps.|
||[name](/javascript/api/excel/excel.linkeddatatype#name)|Der Name des verknüpften Datentyps.|
||[periodicRefreshInterval](/javascript/api/excel/excel.linkeddatatype#periodicrefreshinterval)|Die Häufigkeit in Sekunden, mit der der verknüpfte Datentyp aktualisiert wird, wenn `refreshMode` er auf "Periodisch" festgelegt ist.|
||[refreshMode](/javascript/api/excel/excel.linkeddatatype#refreshmode)|Der Mechanismus, mit dem die Daten für den verknüpften Datentyp abgerufen werden.|
||[ServiceID](/javascript/api/excel/excel.linkeddatatype#serviceid)|Die eindeutige ID des verknüpften Datentyps.|
||[supportedRefreshModes](/javascript/api/excel/excel.linkeddatatype#supportedrefreshmodes)|Gibt ein Array mit allen vom verknüpften Datentyp unterstützten Aktualisierungsmodi zurück.|
||[requestRefresh()](/javascript/api/excel/excel.linkeddatatype#requestrefresh--)|Stellt eine Anforderung zum Aktualisieren des verknüpften Datentyps.|
||[requestSetRefreshMode(refreshMode: Excel.LinkedDataTypeRefreshMode)](/javascript/api/excel/excel.linkeddatatype#requestsetrefreshmode-refreshmode-)|Stellt eine Anforderung zum Ändern des Aktualisierungsmodus für diesen verknüpften Datentyp.|
|[LinkedDataTypeAddedEventArgs](/javascript/api/excel/excel.linkeddatatypeaddedeventargs)|[ServiceID](/javascript/api/excel/excel.linkeddatatypeaddedeventargs#serviceid)|Die eindeutige ID des neuen verknüpften Datentyps.|
||[source](/javascript/api/excel/excel.linkeddatatypeaddedeventargs#source)|Ruft die Quelle des Ereignisses ab.|
||[Typ](/javascript/api/excel/excel.linkeddatatypeaddedeventargs#type)|Ruft den Typ des Ereignisses ab.|
|[LinkedDataTypeCollection](/javascript/api/excel/excel.linkeddatatypecollection)|[getCount()](/javascript/api/excel/excel.linkeddatatypecollection#getcount--)|Ruft die Anzahl der verknüpften Datentypen in der Auflistung ab.|
||[getItem(key: number)](/javascript/api/excel/excel.linkeddatatypecollection#getitem-key-)|Ruft einen verknüpften Datentyp nach Dienst-ID ab.|
||[getItemAt(index: number)](/javascript/api/excel/excel.linkeddatatypecollection#getitemat-index-)|Ruft einen verknüpften Datentyp nach seinem Index in der Auflistung ab.|
||[getItemOrNullObject(key: number)](/javascript/api/excel/excel.linkeddatatypecollection#getitemornullobject-key-)|Ruft einen verknüpften Datentyp nach ID ab.|
||[items](/javascript/api/excel/excel.linkeddatatypecollection#items)|Ruft die geladenen untergeordneten Elemente in dieser Sammlung ab.|
||[requestRefreshAll()](/javascript/api/excel/excel.linkeddatatypecollection#requestrefreshall--)|Stellt eine Anforderung zum Aktualisieren aller verknüpften Datentypen in der Auflistung.|
|[NamedSheetView](/javascript/api/excel/excel.namedsheetview)|[activate()](/javascript/api/excel/excel.namedsheetview#activate--)|Aktiviert diese Blattansicht.|
||[delete()](/javascript/api/excel/excel.namedsheetview#delete--)|Entfernt die Blattansicht aus dem Arbeitsblatt.|
||[duplicate(name?: string)](/javascript/api/excel/excel.namedsheetview#duplicate-name-)|Erstellt eine Kopie dieser Blattansicht.|
||[name](/javascript/api/excel/excel.namedsheetview#name)|Ruft den Namen der Blattansicht ab oder legt den Namen fest.|
|[NamedSheetViewCollection](/javascript/api/excel/excel.namedsheetviewcollection)|[add(name: string)](/javascript/api/excel/excel.namedsheetviewcollection#add-name-)|Erstellt eine neue Blattansicht mit dem angegebenen Namen.|
||[enterTemporary()](/javascript/api/excel/excel.namedsheetviewcollection#entertemporary--)|Erstellt und aktiviert eine neue temporäre Blattansicht.|
||[exit()](/javascript/api/excel/excel.namedsheetviewcollection#exit--)|Beendet die derzeit aktive Blattansicht.|
||[getActive()](/javascript/api/excel/excel.namedsheetviewcollection#getactive--)|Ruft die derzeit aktive Blattansicht des Arbeitsblatts ab.|
||[getCount()](/javascript/api/excel/excel.namedsheetviewcollection#getcount--)|Ruft die Anzahl der Blattansichten in diesem Arbeitsblatt ab.|
||[getItem(key: string)](/javascript/api/excel/excel.namedsheetviewcollection#getitem-key-)|Ruft eine Blattansicht mit ihrem Namen ab.|
||[getItemAt(index: number)](/javascript/api/excel/excel.namedsheetviewcollection#getitemat-index-)|Ruft eine Blattansicht nach ihrem Index in der Auflistung ab.|
||[items](/javascript/api/excel/excel.namedsheetviewcollection#items)|Ruft die geladenen untergeordneten Elemente in dieser Sammlung ab.|
|[PivotLayout](/javascript/api/excel/excel.pivotlayout)|[altTextDescription](/javascript/api/excel/excel.pivotlayout#alttextdescription)|Die Alternativtextbeschreibung der PivotTable.|
||[altTextTitle](/javascript/api/excel/excel.pivotlayout#alttexttitle)|Der Alternativtexttitel der PivotTable.|
||[displayBlankLineAfterEachItem(display: boolean)](/javascript/api/excel/excel.pivotlayout#displayblanklineaftereachitem-display-)|Legt fest, ob nach jedem Element eine leere Zeile angezeigt werden soll.|
||[emptyCellText](/javascript/api/excel/excel.pivotlayout#emptycelltext)|Der Text, der automatisch in eine beliebige leere Zelle in der PivotTable gefüllt wird, wenn `fillEmptyCells == true` .|
||[fillEmptyCells](/javascript/api/excel/excel.pivotlayout#fillemptycells)|Gibt an, ob leere Zellen in der PivotTable mit der aufgefüllt werden `emptyCellText` sollen.|
||[getCell(dataHierarchy: DataPivotHierarchy \| string, rowItems: Array<PivotItem \| string>, columnItems: Array<PivotItem \| string>)](/javascript/api/excel/excel.pivotlayout#getcell-datahierarchy--rowitems--columnitems-)|Ruft eine eindeutige Zelle in der PivotTable ab, die auf einer Datenhierarchie und den Zeilen- und Spaltenelementen ihrer jeweiligen Hierarchie basiert.|
||[pivotStyle](/javascript/api/excel/excel.pivotlayout#pivotstyle)|Die Formatvorlage, die auf die PivotTable angewendet wird.|
||[repeatAllItemLabels(repeatLabels: boolean)](/javascript/api/excel/excel.pivotlayout#repeatallitemlabels-repeatlabels-)|Legt die Einstellung "Alle Elementbeschriftungen wiederholen" für alle Felder in der PivotTable fest.|
||[setStyle(style: string \| PivotTableStyle \| BuiltInPivotTableStyle)](/javascript/api/excel/excel.pivotlayout#setstyle-style-)|Legt die Formatvorlage fest, die auf die PivotTable angewendet wird.|
||[showFieldHeaders](/javascript/api/excel/excel.pivotlayout#showfieldheaders)|Gibt an, ob in der PivotTable Feldkopfzeilen (Feldbeschriftungen und Filterd dropdowns) angezeigt werden.|
|[PivotTable](/javascript/api/excel/excel.pivottable)|[refreshOnOpen](/javascript/api/excel/excel.pivottable#refreshonopen)|Gibt an, ob die PivotTable beim Öffnen der Arbeitsmappe aktualisiert wird.|
|[Range](/javascript/api/excel/excel.range)|[getPrecedents()](/javascript/api/excel/excel.range#getprecedents--)|Gibt ein Objekt zurück, das den Bereich darstellt, der alle Vorgänger einer Zelle in einem Arbeitsblatt oder in mehreren `WorkbookRangeAreas` Arbeitsblättern enthält.|
|[RefreshModeChangedEventArgs](/javascript/api/excel/excel.refreshmodechangedeventargs)|[refreshMode](/javascript/api/excel/excel.refreshmodechangedeventargs#refreshmode)|Der Aktualisierungsmodus für verknüpfte Datentypen.|
||[ServiceID](/javascript/api/excel/excel.refreshmodechangedeventargs#serviceid)|Die eindeutige ID des Objekts, dessen Aktualisierungsmodus geändert wurde.|
||[source](/javascript/api/excel/excel.refreshmodechangedeventargs#source)|Ruft die Quelle des Ereignisses ab.|
||[Typ](/javascript/api/excel/excel.refreshmodechangedeventargs#type)|Ruft den Typ des Ereignisses ab.|
|[RefreshRequestCompletedEventArgs](/javascript/api/excel/excel.refreshrequestcompletedeventargs)|[aktualisiert](/javascript/api/excel/excel.refreshrequestcompletedeventargs#refreshed)|Gibt an, ob die Aktualisierungsanforderung erfolgreich war.|
||[ServiceID](/javascript/api/excel/excel.refreshrequestcompletedeventargs#serviceid)|Die eindeutige ID des Objekts, dessen Aktualisierungsanforderung abgeschlossen wurde.|
||[source](/javascript/api/excel/excel.refreshrequestcompletedeventargs#source)|Ruft die Quelle des Ereignisses ab.|
||[Typ](/javascript/api/excel/excel.refreshrequestcompletedeventargs#type)|Ruft den Typ des Ereignisses ab.|
||[warnings](/javascript/api/excel/excel.refreshrequestcompletedeventargs#warnings)|Ein Array, das alle von der Aktualisierungsanforderung generierten Warnungen enthält.|
|[ShapeCollection](/javascript/api/excel/excel.shapecollection)|[addSvg(xml: string)](/javascript/api/excel/excel.shapecollection#addsvg-xml-)|Erstellt eine skalierbare Vektorgrafik (SVG) aus einer XML-Zeichenfolge und fügt sie dem Arbeitsblatt hinzu.|
|[Datenschnitt](/javascript/api/excel/excel.slicer)|[nameInFormula](/javascript/api/excel/excel.slicer#nameinformula)|Stellt den in der Formel verwendeten Namen des Datenschnitts dar.|
||[slicerStyle](/javascript/api/excel/excel.slicer#slicerstyle)|Die Formatvorlage, die auf den Datenschnitt angewendet wird.|
||[setStyle(style: string \| SlicerStyle \| BuiltInSlicerStyle)](/javascript/api/excel/excel.slicer#setstyle-style-)|Legt die Formatvorlage fest, die auf den Datenschnitt angewendet wird.|
|[Table](/javascript/api/excel/excel.table)|[clearStyle()](/javascript/api/excel/excel.table#clearstyle--)|Ändert die Tabelle so, dass sie die Standard-Tabellenformatvorlage verwendet.|
||[onFiltered](/javascript/api/excel/excel.table#onfiltered)|Tritt ein, wenn ein Filter auf eine bestimmte Tabelle angewendet wird.|
||[tableStyle](/javascript/api/excel/excel.table#tablestyle)|Die Formatvorlage, die auf die Tabelle angewendet wird.|
||[setStyle(style: string \| TableStyle \| BuiltInTableStyle)](/javascript/api/excel/excel.table#setstyle-style-)|Legt die Formatvorlage fest, die auf die Tabelle angewendet wird.|
|[TableCollection](/javascript/api/excel/excel.tablecollection)|[onFiltered](/javascript/api/excel/excel.tablecollection#onfiltered)|Tritt ein, wenn ein Filter auf eine beliebige Tabelle in einer Arbeitsmappe oder auf einem Arbeitsblatt angewendet wird.|
|[TableFilteredEventArgs](/javascript/api/excel/excel.tablefilteredeventargs)|[tableId](/javascript/api/excel/excel.tablefilteredeventargs#tableid)|Ruft die ID der Tabelle ab, in der der Filter angewendet wird.|
||[Typ](/javascript/api/excel/excel.tablefilteredeventargs#type)|Ruft den Typ des Ereignisses ab.|
||[worksheetId](/javascript/api/excel/excel.tablefilteredeventargs#worksheetid)|Ruft die ID des Arbeitsblatts ab, das die Tabelle enthält.|
|[Aufgabe](/javascript/api/excel/excel.task)|[addAssignee(email: string)](/javascript/api/excel/excel.task#addassignee-email-)|Fügt der Aufgabe einen Zugewiesenen hinzu.|
||[applyChanges(taskChanges: Excel.TaskChanges)](/javascript/api/excel/excel.task#applychanges-taskchanges-)|Wendet die angegebenen Änderungen auf den Vorgang an.|
||[Assignees](/javascript/api/excel/excel.task#assignees)|Ruft die Benutzer ab, denen der Vorgang zugewiesen ist.|
||[comment](/javascript/api/excel/excel.task#comment)|Ruft den dem Vorgang zugeordneten Kommentar ab.|
||[dueDate](/javascript/api/excel/excel.task#duedate)|Ruft Datum und Uhrzeit der Fälligkeit des Vorgangs ab.|
||[historyRecords](/javascript/api/excel/excel.task#historyrecords)|Ruft die Verlaufseinträge des Vorgangs ab.|
||[id](/javascript/api/excel/excel.task#id)|Ruft die ID des Vorgangs ab.|
||[percentComplete](/javascript/api/excel/excel.task#percentcomplete)|Ruft den Fertigstellungsprozentsatz des Vorgangs ab.|
||[priority](/javascript/api/excel/excel.task#priority)|Ruft die Priorität des Vorgangs ab.|
||[startDate](/javascript/api/excel/excel.task#startdate)|Ruft das Datum und die Uhrzeit ab, zu der der Vorgang gestartet werden soll.|
||[title](/javascript/api/excel/excel.task#title)|Ruft den Titel des Vorgangs ab.|
||[removeAllAssignees()](/javascript/api/excel/excel.task#removeallassignees--)|Entfernt alle Zugewiesenen aus der Aufgabe.|
||[removeAssignee(email: string)](/javascript/api/excel/excel.task#removeassignee-email-)|Entfernt einen Zugewiesenen aus der Aufgabe.|
||[setPercentComplete(percentComplete: number)](/javascript/api/excel/excel.task#setpercentcomplete-percentcomplete-)|Ändert den Abschluss der Aufgabe.|
||[setPriority(priority: number)](/javascript/api/excel/excel.task#setpriority-priority-)|Ändert die Priorität des Vorgangs.|
||[setStartDateAndDueDate(startDate: Date, dueDate: Date)](/javascript/api/excel/excel.task#setstartdateandduedate-startdate--duedate-)|Ändert den Anfang und die Fälligkeitsdaten des Vorgangs.|
||[setTitle(title: string)](/javascript/api/excel/excel.task#settitle-title-)|Ändert den Titel der Aufgabe.|
|[TaskChanges](/javascript/api/excel/excel.taskchanges)|[dueDate](/javascript/api/excel/excel.taskchanges#duedate)|Legt ein neues Fälligkeitsdatum für den Vorgang in der Zeitzone UTC fest.|
||[emailsToAssign](/javascript/api/excel/excel.taskchanges#emailstoassign)|Legt die E-Mail-Adressen der Benutzer fest, die der Aufgabe zugewiesen werden.|
||[emailsToUnassign](/javascript/api/excel/excel.taskchanges#emailstounassign)|Legt die E-Mail-Adressen der Benutzer so fest, dass die Aufgabe nicht mehr zugewiesen wird.|
||[percentComplete](/javascript/api/excel/excel.taskchanges#percentcomplete)|Legt einen neuen Fertigstellungsprozentsatz für den Vorgang fest.|
||[priority](/javascript/api/excel/excel.taskchanges#priority)|Legt eine neue Priorität für den Vorgang fest.|
||[removeAllPreviousAssignees](/javascript/api/excel/excel.taskchanges#removeallpreviousassignees)|Legt fest, ob durch die Änderung alle vorherigen Zugewiesenen aus der Aufgabe entfernt werden sollen.|
||[startDate](/javascript/api/excel/excel.taskchanges#startdate)|Legt einen neuen Starttermin für den Vorgang in der Zeitzone UTC fest.|
||[title](/javascript/api/excel/excel.taskchanges#title)|Legt einen neuen Titel für die Aufgabe fest.|
|[TaskCollection](/javascript/api/excel/excel.taskcollection)|[getCount()](/javascript/api/excel/excel.taskcollection#getcount--)|Ruft die Anzahl der Vorgänge in der Auflistung ab.|
||[getItem(key: string)](/javascript/api/excel/excel.taskcollection#getitem-key-)|Ruft einen Vorgang mithilfe seiner ID ab.|
||[getItemAt(index: number)](/javascript/api/excel/excel.taskcollection#getitemat-index-)|Ruft einen Vorgang nach seinem Index in der Auflistung ab.|
||[getItemOrNullObject(key: string)](/javascript/api/excel/excel.taskcollection#getitemornullobject-key-)|Ruft einen Vorgang mithilfe seiner ID ab.|
||[items](/javascript/api/excel/excel.taskcollection#items)|Ruft die geladenen untergeordneten Elemente in dieser Sammlung ab.|
|[TaskHistoryRecord](/javascript/api/excel/excel.taskhistoryrecord)|[anchorId](/javascript/api/excel/excel.taskhistoryrecord#anchorid)|Stellt die ID des Objekts dar, an das der Vorgang verankert ist (z. B. commentId für Vorgänge, die Kommentaren zugeordnet sind).|
||[Assignee](/javascript/api/excel/excel.taskhistoryrecord#assignee)|Stellt den Benutzer dar, der der Aufgabe für einen Verlaufsdatensatztyp "Zuweisen" zugewiesen ist, oder den Benutzer, der die Zuweisung von der Aufgabe für einen Verlaufsdatensatztyp "Unassign" auflöss.|
||[attributionUser](/javascript/api/excel/excel.taskhistoryrecord#attributionuser)|Stellt den Benutzer dar, der die Aufgabe erstellt oder geändert hat.|
||[dueDate](/javascript/api/excel/excel.taskhistoryrecord#duedate)|Stellt das Fälligkeitsdatum des Vorgangs dar.|
||[historyRecordCreatedDate](/javascript/api/excel/excel.taskhistoryrecord#historyrecordcreateddate)|Stellt das Erstellungsdatum des Vorgangsverlaufsdatensatzes dar.|
||[id](/javascript/api/excel/excel.taskhistoryrecord#id)|ID für den Verlaufsdatensatz.|
||[percentComplete](/javascript/api/excel/excel.taskhistoryrecord#percentcomplete)|Stellt den Fertigstellungsprozentsatz des Vorgangs dar.|
||[priority](/javascript/api/excel/excel.taskhistoryrecord#priority)|Stellt die Priorität des Vorgangs dar.|
||[startDate](/javascript/api/excel/excel.taskhistoryrecord#startdate)|Stellt den Starttermin des Vorgangs dar.|
||[title](/javascript/api/excel/excel.taskhistoryrecord#title)|Stellt den Titel des Vorgangs dar.|
||[Typ](/javascript/api/excel/excel.taskhistoryrecord#type)|Stellt den Typ des Vorgangsverlaufsdatensatzs dar.|
||[undoHistoryId](/javascript/api/excel/excel.taskhistoryrecord#undohistoryid)|Stellt die TaskHistoryRecord.id dar, die für den Datensatztyp "Rückgängig" rückgängig gemacht wurde.|
|[TaskHistoryRecordCollection](/javascript/api/excel/excel.taskhistoryrecordcollection)|[getCount()](/javascript/api/excel/excel.taskhistoryrecordcollection#getcount--)|Ruft die Anzahl der Verlaufseinträge in der Auflistung für den Vorgang ab.|
||[getItemAt(index: number)](/javascript/api/excel/excel.taskhistoryrecordcollection#getitemat-index-)|Ruft einen Vorgangsverlaufsdatensatz mithilfe seines Indexes in der Auflistung ab.|
||[items](/javascript/api/excel/excel.taskhistoryrecordcollection#items)|Ruft die geladenen untergeordneten Elemente in dieser Sammlung ab.|
|[Benutzer](/javascript/api/excel/excel.user)|[displayName](/javascript/api/excel/excel.user#displayname)|Stellt den Anzeigenamen des Benutzers dar.|
||[email](/javascript/api/excel/excel.user#email)|Stellt die E-Mail-Adresse des Benutzers dar.|
||[uid](/javascript/api/excel/excel.user#uid)|Stellt die eindeutige ID des Benutzers dar.|
|[Workbook](/javascript/api/excel/excel.workbook)|[linkedDataTypes](/javascript/api/excel/excel.workbook#linkeddatatypes)|Gibt eine Auflistung verknüpfter Datentypen zurück, die Teil der Arbeitsmappe sind.|
||[tasks](/javascript/api/excel/excel.workbook#tasks)|Gibt eine Auflistung von Aufgaben zurück, die in der Arbeitsmappe vorhanden sind.|
||[showPivotFieldList](/javascript/api/excel/excel.workbook#showpivotfieldlist)|Gibt an, ob der Feldlistenbereich der PivotTable auf Arbeitsmappenebene angezeigt wird.|
||[use1904DateSystem](/javascript/api/excel/excel.workbook#use1904datesystem)|True, falls die Arbeitsmappe das 1904-Datumssystem verwendet.|
|[Arbeitsblatt](/javascript/api/excel/excel.worksheet)|[namedSheetViews](/javascript/api/excel/excel.worksheet#namedsheetviews)|Gibt eine Auflistung von Blattansichten zurück, die im Arbeitsblatt vorhanden sind.|
||[onFiltered](/javascript/api/excel/excel.worksheet#onfiltered)|Tritt ein, wenn ein Filter auf ein bestimmtes Arbeitsblatt angewendet wird.|
||[onFormulaChanged](/javascript/api/excel/excel.worksheet#onformulachanged)|Tritt auf, wenn eine oder mehrere Formeln in diesem Arbeitsblatt geändert werden.|
||[tasks](/javascript/api/excel/excel.worksheet#tasks)|Gibt eine Auflistung von Aufgaben zurück, die im Arbeitsblatt vorhanden sind.|
|[WorksheetCollection](/javascript/api/excel/excel.worksheetcollection)|[addFromBase64(base64File: string, sheetNamesToInsert?: string[], positionType?: Excel.WorksheetPositionType, relativeTo?: Worksheet \| string)](/javascript/api/excel/excel.worksheetcollection#addfrombase64-base64file--sheetnamestoinsert--positiontype--relativeto-)|Fügt die angegebenen Arbeitsblätter einer Arbeitsmappe in die aktuelle Arbeitsmappe ein.|
||[onFiltered](/javascript/api/excel/excel.worksheetcollection#onfiltered)|Tritt ein, wenn ein Filter eines beliebigen Arbeitsblatts in der Arbeitsmappe angewendet wird.|
||[onFormulaChanged](/javascript/api/excel/excel.worksheetcollection#onformulachanged)|Tritt auf, wenn eine oder mehrere Formeln in einem Beliebigen Arbeitsblatt dieser Auflistung geändert werden.|
|[WorksheetFilteredEventArgs](/javascript/api/excel/excel.worksheetfilteredeventargs)|[type](/javascript/api/excel/excel.worksheetfilteredeventargs#type)|Ruft den Typ des Ereignisses ab.|
||[worksheetId](/javascript/api/excel/excel.worksheetfilteredeventargs#worksheetid)|Ruft die ID des Arbeitsblatts ab, auf das der Filter angewendet wird.|
|[WorksheetFormulaChangedEventArgs](/javascript/api/excel/excel.worksheetformulachangedeventargs)|[formulaDetails](/javascript/api/excel/excel.worksheetformulachangedeventargs#formuladetails)|Ruft ein Array von FormulaChangedEventDetail -Objekten, die die Details zu allen geänderten Formeln enthalten.|
||[source](/javascript/api/excel/excel.worksheetformulachangedeventargs#source)|Die Quelle des Ereignisses.|
||[Typ](/javascript/api/excel/excel.worksheetformulachangedeventargs#type)|Ruft den Typ des Ereignisses ab.|
||[worksheetId](/javascript/api/excel/excel.worksheetformulachangedeventargs#worksheetid)|Ruft die ID des Arbeitsblatts ab, in dem die Formel geändert wurde.|
