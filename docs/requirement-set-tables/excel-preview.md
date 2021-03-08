| Klasse | Felder | Beschreibung |
|:---|:---|:---|
|[AutoFilter](/javascript/api/excel/excel.autofilter)|[clearColumnCriteria(columnIndex: number)](/javascript/api/excel/excel.autofilter#clearcolumncriteria-columnindex-)|Löscht die Filterkriterien von AutoFilter.|
|[Kommentar](/javascript/api/excel/excel.comment)|[assignTask(assigneee: Identity)](/javascript/api/excel/excel.comment#assigntask-assignee-)|Weist die dem Kommentar zugeordnete Aufgabe dem angegebenen Benutzer als Assignee zu.|
||[getTask()](/javascript/api/excel/excel.comment#gettask--)|Ruft die Aufgabe ab, die diesem Kommentar zugeordnet ist.|
||[getTaskOrNullObject()](/javascript/api/excel/excel.comment#gettaskornullobject--)|Ruft die Aufgabe ab, die diesem Kommentar zugeordnet ist.|
|[CommentCollection](/javascript/api/excel/excel.commentcollection)|[getItemOrNullObject(commentId: string)](/javascript/api/excel/excel.commentcollection#getitemornullobject-commentid-)|Ruft einen Kommentar aus der Sammlung basierend auf der ID ab.|
|[CommentReply](/javascript/api/excel/excel.commentreply)|[assignTask(assigneee: Identity)](/javascript/api/excel/excel.commentreply#assigntask-assignee-)|Weist die dem Kommentar zugeordnete Aufgabe dem angegebenen Benutzer als alleiniger Zuordnende zu.|
||[getTask()](/javascript/api/excel/excel.commentreply#gettask--)|Ruft die Aufgabe ab, die dem Thread dieser Kommentarantwort zugeordnet ist.|
||[getTaskOrNullObject()](/javascript/api/excel/excel.commentreply#gettaskornullobject--)|Ruft die Aufgabe ab, die dem Thread dieser Kommentarantwort zugeordnet ist.|
|[CommentReplyCollection](/javascript/api/excel/excel.commentreplycollection)|[getItemOrNullObject(commentReplyId: string)](/javascript/api/excel/excel.commentreplycollection#getitemornullobject-commentreplyid-)|Gibt eine Kommentarantwort zurück, die durch ihre ID angegeben ist.|
|[ConditionalFormatCollection](/javascript/api/excel/excel.conditionalformatcollection)|[getItemOrNullObject(id: string)](/javascript/api/excel/excel.conditionalformatcollection#getitemornullobject-id-)|Gibt ein bedingtes Format zurück, das durch seine ID identifiziert wird.|
|[DocumentTask](/javascript/api/excel/excel.documenttask)|[percentComplete](/javascript/api/excel/excel.documenttask#percentcomplete)|Gibt den Fertigstellungsprozentsatz des Vorgangs an.|
||[priority](/javascript/api/excel/excel.documenttask#priority)|Gibt die Priorität des Vorgangs an.|
||[assignees](/javascript/api/excel/excel.documenttask#assignees)|Gibt eine Auflistung von Assignees des Vorgangs zurück.|
||[änderungen](/javascript/api/excel/excel.documenttask#changes)|Ruft die Änderungseinträge des Vorgangs ab.|
||[comment](/javascript/api/excel/excel.documenttask#comment)|Ruft den Kommentar ab, der der Aufgabe zugeordnet ist.|
||[completedBy](/javascript/api/excel/excel.documenttask#completedby)|Ruft den letzten Benutzer ab, der die Aufgabe abgeschlossen hat.|
||[completedDateTime](/javascript/api/excel/excel.documenttask#completeddatetime)|Ruft Datum und Uhrzeit ab, zu der der Vorgang abgeschlossen wurde.|
||[createdBy](/javascript/api/excel/excel.documenttask#createdby)|Ruft den Benutzer ab, der die Aufgabe erstellt hat.|
||[createdDateTime](/javascript/api/excel/excel.documenttask#createddatetime)|Ruft Datum und Uhrzeit ab, zu der der Vorgang erstellt wurde.|
||[id](/javascript/api/excel/excel.documenttask#id)|Ruft die ID des Vorgangs ab.|
||[setStartAndDueDateTime(startDateTime: Date, dueDateTime: Date)](/javascript/api/excel/excel.documenttask#setstartandduedatetime-startdatetime--duedatetime-)|Ändert den Anfang und die Fälligkeitsdaten des Vorgangs.|
||[startAndDueDateTime](/javascript/api/excel/excel.documenttask#startandduedatetime)|Ruft das Datum und die Uhrzeit ab, zu der der Vorgang gestartet werden soll, oder legt diese fest.|
||[title](/javascript/api/excel/excel.documenttask#title)|Gibt den Titel des Vorgangs an.|
|[DocumentTaskChange](/javascript/api/excel/excel.documenttaskchange)|[Assignee](/javascript/api/excel/excel.documenttaskchange#assignee)|Represents the user assigned to the task for an `assign` change record type, or the user unassigned from the task for an `unassign` change record type.|
||[changedBy](/javascript/api/excel/excel.documenttaskchange#changedby)|Stellt den Benutzer dar, der die Aufgabe erstellt oder geändert hat.|
||[commentId](/javascript/api/excel/excel.documenttaskchange#commentid)|Represents the ID of the `Comment` or to which the task change is `CommentReply` anchored.|
||[createdDateTime](/javascript/api/excel/excel.documenttaskchange#createddatetime)|Stellt das Erstellungsdatum und die Uhrzeit des Vorgangsänderungsdatensatz dar.|
||[dueDateTime](/javascript/api/excel/excel.documenttaskchange#duedatetime)|Stellt das Fälligkeitsdatum und die Uhrzeit des Vorgangs in der UTC-Zeitzone dar.|
||[id](/javascript/api/excel/excel.documenttaskchange#id)|ID für den Vorgangsänderungsdatensatz.|
||[percentComplete](/javascript/api/excel/excel.documenttaskchange#percentcomplete)|Stellt den Fertigstellungsprozentsatz des Vorgangs dar.|
||[priority](/javascript/api/excel/excel.documenttaskchange#priority)|Stellt die Priorität des Vorgangs dar.|
||[startDateTime](/javascript/api/excel/excel.documenttaskchange#startdatetime)|Represents the task's start date and time, in UTC time zone.|
||[title](/javascript/api/excel/excel.documenttaskchange#title)|Stellt den Titel des Vorgangs dar.|
||[type](/javascript/api/excel/excel.documenttaskchange#type)|Stellt den Aktionstyp des Vorgangsänderungsdatensatz dar.|
||[undoHistoryId](/javascript/api/excel/excel.documenttaskchange#undohistoryid)|Stellt die `DocumentTaskChange.id` Eigenschaft dar, die für den Änderungsdatensatztyp `undo` rückgängig gemacht wurde.|
|[DocumentTaskChangeCollection](/javascript/api/excel/excel.documenttaskchangecollection)|[getCount()](/javascript/api/excel/excel.documenttaskchangecollection#getcount--)|Ruft die Anzahl der Änderungseinträge in der Auflistung für den Vorgang ab.|
||[getItemAt(index: number)](/javascript/api/excel/excel.documenttaskchangecollection#getitemat-index-)|Ruft einen Vorgangsänderungsdatensatz mithilfe seines Indexes in der Auflistung ab.|
||[items](/javascript/api/excel/excel.documenttaskchangecollection#items)|Ruft die geladenen untergeordneten Elemente in dieser Sammlung ab.|
|[DocumentTaskCollection](/javascript/api/excel/excel.documenttaskcollection)|[getCount()](/javascript/api/excel/excel.documenttaskcollection#getcount--)|Ruft die Anzahl der Vorgänge in der Auflistung ab.|
||[getItem(key: string)](/javascript/api/excel/excel.documenttaskcollection#getitem-key-)|Ruft einen Vorgang mithilfe seiner ID ab.|
||[getItemAt(index: number)](/javascript/api/excel/excel.documenttaskcollection#getitemat-index-)|Ruft einen Vorgang nach seinem Index in der Auflistung ab.|
||[getItemOrNullObject(key: string)](/javascript/api/excel/excel.documenttaskcollection#getitemornullobject-key-)|Ruft einen Vorgang mithilfe seiner ID ab.|
||[items](/javascript/api/excel/excel.documenttaskcollection#items)|Ruft die geladenen untergeordneten Elemente in dieser Sammlung ab.|
|[DocumentTaskSchedule](/javascript/api/excel/excel.documenttaskschedule)|[dueDateTime](/javascript/api/excel/excel.documenttaskschedule#duedatetime)|Ruft das Datum und die Uhrzeit ab, zu der der Vorgang fällig ist.|
||[startDateTime](/javascript/api/excel/excel.documenttaskschedule#startdatetime)|Ruft Datum und Uhrzeit ab, zu der der Vorgang gestartet werden soll.|
|[FormulaChangedEventDetail](/javascript/api/excel/excel.formulachangedeventdetail)|[cellAddress](/javascript/api/excel/excel.formulachangedeventdetail#celladdress)|Die Adresse der Zelle, die die geänderte Formel enthält.|
||[previousFormula](/javascript/api/excel/excel.formulachangedeventdetail#previousformula)|Stellt die vorherige Formel dar, bevor sie geändert wurde.|
|[GroupShapeCollection](/javascript/api/excel/excel.groupshapecollection)|[getItemOrNullObject(key: string)](/javascript/api/excel/excel.groupshapecollection#getitemornullobject-key-)|Ruft ein Shape mit seinem Namen oder seiner ID ab.|
|[Identität](/javascript/api/excel/excel.identity)|[displayName](/javascript/api/excel/excel.identity#displayname)|Stellt den Anzeigenamen des Benutzers dar.|
||[email](/javascript/api/excel/excel.identity#email)|Stellt die E-Mail-Adresse des Benutzers dar.|
||[id](/javascript/api/excel/excel.identity#id)|Stellt die eindeutige ID des Benutzers dar.|
|[IdentityCollection](/javascript/api/excel/excel.identitycollection)|[add(assignee: Identity)](/javascript/api/excel/excel.identitycollection#add-assignee-)|Fügt der Auflistung eine Benutzeridentität hinzu.|
||[clear()](/javascript/api/excel/excel.identitycollection#clear--)|Entfernt alle Benutzeridentitäten aus der Auflistung.|
||[getCount()](/javascript/api/excel/excel.identitycollection#getcount--)|Ruft die Anzahl der Elemente in der Auflistung ab.|
||[getItemAt(index: number)](/javascript/api/excel/excel.identitycollection#getitemat-index-)|Ruft eine Dokumentbenutzeridentität mithilfe des Index in der Auflistung ab.|
||[remove(assignee: Identity)](/javascript/api/excel/excel.identitycollection#remove-assignee-)|Entfernt eine Benutzeridentität aus der Auflistung.|
|[InsertWorksheetOptions](/javascript/api/excel/excel.insertworksheetoptions)|[positionType](/javascript/api/excel/excel.insertworksheetoptions#positiontype)|Die Einfügeposition der neuen Arbeitsblätter in der aktuellen Arbeitsmappe.|
||[relativeTo](/javascript/api/excel/excel.insertworksheetoptions#relativeto)|Das Arbeitsblatt in der aktuellen Arbeitsmappe, auf das für den Parameter verwiesen `WorksheetPositionType` wird.|
||[sheetNamesToInsert](/javascript/api/excel/excel.insertworksheetoptions#sheetnamestoinsert)|Die Namen der einzelnen arbeitsblätter, die eingefügt werden sollen.|
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
||[type](/javascript/api/excel/excel.linkeddatatypeaddedeventargs#type)|Ruft den Typ des Ereignisses ab.|
|[LinkedDataTypeCollection](/javascript/api/excel/excel.linkeddatatypecollection)|[getCount()](/javascript/api/excel/excel.linkeddatatypecollection#getcount--)|Ruft die Anzahl verknüpfter Datentypen in der Auflistung ab.|
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
||[exit()](/javascript/api/excel/excel.namedsheetviewcollection#exit--)|Beendet die aktuell aktive Blattansicht.|
||[getActive()](/javascript/api/excel/excel.namedsheetviewcollection#getactive--)|Ruft die aktuell aktive Blattansicht des Arbeitsblatts ab.|
||[getCount()](/javascript/api/excel/excel.namedsheetviewcollection#getcount--)|Ruft die Anzahl der Blattansichten in diesem Arbeitsblatt ab.|
||[getItem(key: string)](/javascript/api/excel/excel.namedsheetviewcollection#getitem-key-)|Ruft eine Blattansicht mit ihrem Namen ab.|
||[getItemAt(index: number)](/javascript/api/excel/excel.namedsheetviewcollection#getitemat-index-)|Ruft eine Blattansicht nach ihrem Index in der Auflistung ab.|
||[getItemOrNullObject(key: string)](/javascript/api/excel/excel.namedsheetviewcollection#getitemornullobject-key-)|Ruft eine Blattansicht mit ihrem Namen ab.|
||[items](/javascript/api/excel/excel.namedsheetviewcollection#items)|Ruft die geladenen untergeordneten Elemente in dieser Sammlung ab.|
|[PivotLayout](/javascript/api/excel/excel.pivotlayout)|[altTextDescription](/javascript/api/excel/excel.pivotlayout#alttextdescription)|Die Alttextbeschreibung der PivotTable.|
||[altTextTitle](/javascript/api/excel/excel.pivotlayout#alttexttitle)|Der Alttexttitel der PivotTable.|
||[displayBlankLineAfterEachItem(display: boolean)](/javascript/api/excel/excel.pivotlayout#displayblanklineaftereachitem-display-)|Legt fest, ob nach jedem Element eine leere Zeile angezeigt werden soll.|
||[emptyCellText](/javascript/api/excel/excel.pivotlayout#emptycelltext)|Der Text, der automatisch in eine beliebige leere Zelle in der PivotTable gefüllt wird, wenn `fillEmptyCells == true` .|
||[fillEmptyCells](/javascript/api/excel/excel.pivotlayout#fillemptycells)|Gibt an, ob leere Zellen in der PivotTable mit aufgefüllt werden `emptyCellText` sollen.|
||[getCell(dataHierarchy: DataPivotHierarchy \| string, rowItems: Array<PivotItem \| string>, columnItems: Array<PivotItem \| string>)](/javascript/api/excel/excel.pivotlayout#getcell-datahierarchy--rowitems--columnitems-)|Ruft eine eindeutige Zelle in der PivotTable ab, die auf einer Datenhierarchie und den Zeilen- und Spaltenelementen ihrer jeweiligen Hierarchie basiert.|
||[pivotStyle](/javascript/api/excel/excel.pivotlayout#pivotstyle)|Die auf die PivotTable angewendete Formatvorlage.|
||[repeatAllItemLabels(repeatLabels: boolean)](/javascript/api/excel/excel.pivotlayout#repeatallitemlabels-repeatlabels-)|Legt die Einstellung "Alle Elementbeschriftungen wiederholen" für alle Felder in der PivotTable fest.|
||[setStyle(style: string \| PivotTableStyle \| BuiltInPivotTableStyle)](/javascript/api/excel/excel.pivotlayout#setstyle-style-)|Legt die auf die PivotTable angewendete Formatvorlage fest.|
||[showFieldHeaders](/javascript/api/excel/excel.pivotlayout#showfieldheaders)|Gibt an, ob in der PivotTable Feldkopfzeilen (Feldbeschriftungen und Filterabdings) angezeigt werden.|
|[PivotTable](/javascript/api/excel/excel.pivottable)|[refreshOnOpen](/javascript/api/excel/excel.pivottable#refreshonopen)|Gibt an, ob die PivotTable aktualisiert wird, wenn die Arbeitsmappe geöffnet wird.|
|[PivotTableScopedCollection](/javascript/api/excel/excel.pivottablescopedcollection)|[getFirstOrNullObject()](/javascript/api/excel/excel.pivottablescopedcollection#getfirstornullobject--)|Ruft die erste PivotTable in der Auflistung ab.|
|[Range](/javascript/api/excel/excel.range)|[getDependents()](/javascript/api/excel/excel.range#getdependents--)|Gibt ein Objekt zurück, das den Bereich darstellt, der alle Abhängigen einer Zelle im gleichen Arbeitsblatt oder `WorkbookRangeAreas` in mehreren Arbeitsblättern enthält.|
||[getDirectDependents()](/javascript/api/excel/excel.range#getdirectdependents--)|Gibt ein Objekt zurück, das den Bereich darstellt, der alle direkten Abhängigkeiten einer Zelle im gleichen Arbeitsblatt oder `WorkbookRangeAreas` in mehreren Arbeitsblättern enthält.|
||[getExtendedRange(direction: Excel.KeyboardDirection, activeCell?: Range \| string)](/javascript/api/excel/excel.range#getextendedrange-direction--activecell-)|Gibt ein Range-Objekt zurück, das den aktuellen Bereich und bis zum Rand des Bereichs enthält, basierend auf der angegebenen Richtung.|
||[getMergedAreasOrNullObject()](/javascript/api/excel/excel.range#getmergedareasornullobject--)|Gibt ein RangeAreas-Objekt zurück, das die zusammengeführten Bereiche in diesem Bereich darstellt.|
||[getPrecedents()](/javascript/api/excel/excel.range#getprecedents--)|Gibt ein Objekt zurück, das den Bereich darstellt, der alle Vorgänger einer Zelle im gleichen Arbeitsblatt oder `WorkbookRangeAreas` in mehreren Arbeitsblättern enthält.|
||[getRangeEdge(direction: Excel.KeyboardDirection, activeCell?: Range \| string)](/javascript/api/excel/excel.range#getrangeedge-direction--activecell-)|Gibt ein Range-Objekt zurück, das die Edgezelle des Datenbereichs ist, die der angegebenen Richtung entspricht.|
|[RefreshModeChangedEventArgs](/javascript/api/excel/excel.refreshmodechangedeventargs)|[refreshMode](/javascript/api/excel/excel.refreshmodechangedeventargs#refreshmode)|Der verknüpfte Datentypaktualisierungsmodus.|
||[ServiceID](/javascript/api/excel/excel.refreshmodechangedeventargs#serviceid)|Die eindeutige ID des Objekts, dessen Aktualisierungsmodus geändert wurde.|
||[source](/javascript/api/excel/excel.refreshmodechangedeventargs#source)|Ruft die Quelle des Ereignisses ab.|
||[type](/javascript/api/excel/excel.refreshmodechangedeventargs#type)|Ruft den Typ des Ereignisses ab.|
|[RefreshRequestCompletedEventArgs](/javascript/api/excel/excel.refreshrequestcompletedeventargs)|[aktualisiert](/javascript/api/excel/excel.refreshrequestcompletedeventargs#refreshed)|Gibt an, ob die Aktualisierungsanforderung erfolgreich war.|
||[ServiceID](/javascript/api/excel/excel.refreshrequestcompletedeventargs#serviceid)|Die eindeutige ID des Objekts, dessen Aktualisierungsanforderung abgeschlossen wurde.|
||[source](/javascript/api/excel/excel.refreshrequestcompletedeventargs#source)|Ruft die Quelle des Ereignisses ab.|
||[type](/javascript/api/excel/excel.refreshrequestcompletedeventargs#type)|Ruft den Typ des Ereignisses ab.|
||[Warnungen](/javascript/api/excel/excel.refreshrequestcompletedeventargs#warnings)|Ein Array, das alle von der Aktualisierungsanforderung generierten Warnungen enthält.|
|[ShapeCollection](/javascript/api/excel/excel.shapecollection)|[addSvg(xml: string)](/javascript/api/excel/excel.shapecollection#addsvg-xml-)|Erstellt eine skalierbare Vektorgrafik (SVG) aus einer XML-Zeichenfolge und fügt sie dem Arbeitsblatt hinzu.|
||[getItemOrNullObject(key: string)](/javascript/api/excel/excel.shapecollection#getitemornullobject-key-)|Ruft ein Shape mit seinem Namen oder seiner ID ab.|
|[Datenschnitt](/javascript/api/excel/excel.slicer)|[nameInFormula](/javascript/api/excel/excel.slicer#nameinformula)|Stellt den in der Formel verwendeten Namen des Datenschnitts dar.|
||[slicerStyle](/javascript/api/excel/excel.slicer#slicerstyle)|Die auf den Datenschnitt angewendete Formatvorlage.|
||[setStyle(style: string \| SlicerStyle \| BuiltInSlicerStyle)](/javascript/api/excel/excel.slicer#setstyle-style-)|Legt die auf den Datenschnitt angewendete Formatvorlage fest.|
|[StyleCollection](/javascript/api/excel/excel.stylecollection)|[GetItemOrNullObject(name: string)](/javascript/api/excel/excel.stylecollection#getitemornullobject-name-)|Ruft eine Formatvorlage anhand des Namens ab.|
|[Table](/javascript/api/excel/excel.table)|[clearStyle()](/javascript/api/excel/excel.table#clearstyle--)|Ändert die Tabelle so, dass sie die Standard-Tabellenformatvorlage verwendet.|
||[onFiltered](/javascript/api/excel/excel.table#onfiltered)|Tritt auf, wenn ein Filter auf eine bestimmte Tabelle angewendet wird.|
||[tableStyle](/javascript/api/excel/excel.table#tablestyle)|Die auf die Tabelle angewendete Formatvorlage.|
||[resize(newRange: Range \| string)](/javascript/api/excel/excel.table#resize-newrange-)|Ändern Sie die Größe der Tabelle auf den neuen Bereich.|
||[setStyle(style: string \| TableStyle \| BuiltInTableStyle)](/javascript/api/excel/excel.table#setstyle-style-)|Legt die auf die Tabelle angewendete Formatvorlage fest.|
|[TableCollection](/javascript/api/excel/excel.tablecollection)|[onFiltered](/javascript/api/excel/excel.tablecollection#onfiltered)|Tritt auf, wenn ein Filter auf eine beliebige Tabelle in einer Arbeitsmappe oder auf ein Arbeitsblatt angewendet wird.|
|[TableFilteredEventArgs](/javascript/api/excel/excel.tablefilteredeventargs)|[tableId](/javascript/api/excel/excel.tablefilteredeventargs#tableid)|Ruft die ID der Tabelle ab, in der der Filter angewendet wird.|
||[type](/javascript/api/excel/excel.tablefilteredeventargs#type)|Ruft den Typ des Ereignisses ab.|
||[worksheetId](/javascript/api/excel/excel.tablefilteredeventargs#worksheetid)|Ruft die ID des Arbeitsblatts ab, das die Tabelle enthält.|
|[TableScopedCollection](/javascript/api/excel/excel.tablescopedcollection)|[getItemOrNullObject(key: string)](/javascript/api/excel/excel.tablescopedcollection#getitemornullobject-key-)|Ruft eine Tabelle anhand des Namens oder der ID ab.|
|[Workbook](/javascript/api/excel/excel.workbook)|[insertWorksheetsFromBase64(base64File: string, options?: Excel.InsertWorksheetOptions)](/javascript/api/excel/excel.workbook#insertworksheetsfrombase64-base64file--options-)|Fügt die angegebenen Arbeitsblätter aus einer Quellarbeitsmappe in die aktuelle Arbeitsmappe ein.|
||[linkedDataTypes](/javascript/api/excel/excel.workbook#linkeddatatypes)|Gibt eine Auflistung verknüpfter Datentypen zurück, die Teil der Arbeitsmappe sind.|
||[onActivated](/javascript/api/excel/excel.workbook#onactivated)|Tritt auf, wenn die Arbeitsmappe aktiviert wird.|
||[tasks](/javascript/api/excel/excel.workbook#tasks)|Gibt eine Auflistung von Aufgaben zurück, die in der Arbeitsmappe vorhanden sind.|
||[showPivotFieldList](/javascript/api/excel/excel.workbook#showpivotfieldlist)|Gibt an, ob der Feldlistenbereich der PivotTable auf Arbeitsmappenebene angezeigt wird.|
||[use1904DateSystem](/javascript/api/excel/excel.workbook#use1904datesystem)|True, falls die Arbeitsmappe das 1904-Datumssystem verwendet.|
|[WorkbookActivatedEventArgs](/javascript/api/excel/excel.workbookactivatedeventargs)|[type](/javascript/api/excel/excel.workbookactivatedeventargs#type)|Ruft den Typ des Ereignisses ab.|
|[Arbeitsblatt](/javascript/api/excel/excel.worksheet)|[namedSheetViews](/javascript/api/excel/excel.worksheet#namedsheetviews)|Gibt eine Auflistung von Blattansichten zurück, die im Arbeitsblatt vorhanden sind.|
||[onFiltered](/javascript/api/excel/excel.worksheet#onfiltered)|Tritt auf, wenn ein Filter auf ein bestimmtes Arbeitsblatt angewendet wird.|
||[onFormulaChanged](/javascript/api/excel/excel.worksheet#onformulachanged)|Tritt auf, wenn eine oder mehrere Formeln in diesem Arbeitsblatt geändert werden.|
||[tasks](/javascript/api/excel/excel.worksheet#tasks)|Gibt eine Auflistung von Aufgaben zurück, die im Arbeitsblatt vorhanden sind.|
|[WorksheetCollection](/javascript/api/excel/excel.worksheetcollection)|[addFromBase64(base64File: string, sheetNamesToInsert?: string[], positionType?: Excel.WorksheetPositionType, relativeTo?: Worksheet \| string)](/javascript/api/excel/excel.worksheetcollection#addfrombase64-base64file--sheetnamestoinsert--positiontype--relativeto-)|Fügt die angegebenen Arbeitsblätter einer Arbeitsmappe in die aktuelle Arbeitsmappe ein.|
||[onFiltered](/javascript/api/excel/excel.worksheetcollection#onfiltered)|Tritt ein, wenn ein Filter eines beliebigen Arbeitsblatts in der Arbeitsmappe angewendet wird.|
||[onFormulaChanged](/javascript/api/excel/excel.worksheetcollection#onformulachanged)|Tritt auf, wenn eine oder mehrere Formeln in einem Beliebigen Arbeitsblatt dieser Auflistung geändert werden.|
|[WorksheetFilteredEventArgs](/javascript/api/excel/excel.worksheetfilteredeventargs)|[type](/javascript/api/excel/excel.worksheetfilteredeventargs#type)|Ruft den Typ des Ereignisses ab.|
||[worksheetId](/javascript/api/excel/excel.worksheetfilteredeventargs#worksheetid)|Ruft die ID des Arbeitsblatts ab, in dem der Filter angewendet wird.|
|[WorksheetFormulaChangedEventArgs](/javascript/api/excel/excel.worksheetformulachangedeventargs)|[formulaDetails](/javascript/api/excel/excel.worksheetformulachangedeventargs#formuladetails)|Ruft ein Array von Objekten ab, das die Details zu allen geänderten `FormulaChangedEventDetail` Formeln enthält.|
||[source](/javascript/api/excel/excel.worksheetformulachangedeventargs#source)|Die Quelle des Ereignisses.|
||[type](/javascript/api/excel/excel.worksheetformulachangedeventargs#type)|Ruft den Typ des Ereignisses ab.|
||[worksheetId](/javascript/api/excel/excel.worksheetformulachangedeventargs#worksheetid)|Ruft die ID des Arbeitsblatts ab, in dem die Formel geändert wurde.|
