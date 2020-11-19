| Klasse | Felder | Beschreibung |
|:---|:---|:---|
|[Kommentar](/javascript/api/excel/excel.comment)|[assignTask (e-Mail: Zeichenfolge)](/javascript/api/excel/excel.comment#assigntask-email-)|Weist dem angegebenen Benutzer als alleinigem Empfänger die dem Kommentar zugeordnete Aufgabe zu.|
||[gettask ()](/javascript/api/excel/excel.comment#gettask--)|Ruft die Aufgabe ab, die diesem Kommentar zugeordnet ist.|
||[getTaskOrNullObject()](/javascript/api/excel/excel.comment#gettaskornullobject--)|Ruft die Aufgabe ab, die diesem Kommentar zugeordnet ist.|
|[CommentReply](/javascript/api/excel/excel.commentreply)|[assignTask (e-Mail: Zeichenfolge)](/javascript/api/excel/excel.commentreply#assigntask-email-)|Weist dem angegebenen Benutzer als alleinigem Empfänger die dem Kommentar zugeordnete Aufgabe zu.|
||[gettask ()](/javascript/api/excel/excel.commentreply#gettask--)|Ruft die Aufgabe ab, die diesem Kommentar zugeordnet ist.|
||[getTaskOrNullObject()](/javascript/api/excel/excel.commentreply#gettaskornullobject--)|Ruft die Aufgabe ab, die diesem Kommentar zugeordnet ist.|
|[LinkedDataType](/javascript/api/excel/excel.linkeddatatype)|[DataProvider](/javascript/api/excel/excel.linkeddatatype#dataprovider)|Der Name des Datenanbieters für den verknüpften Datentyp.|
||[lastRefreshed](/javascript/api/excel/excel.linkeddatatype#lastrefreshed)|Das Datum und die Uhrzeit der lokalen Zeitzone, seit die Arbeitsmappe geöffnet wurde, als der verknüpfte Datentyp zuletzt aktualisiert wurde.|
||[name](/javascript/api/excel/excel.linkeddatatype#name)|Der Name des verknüpften Daten Typs.|
||[periodicRefreshInterval](/javascript/api/excel/excel.linkeddatatype#periodicrefreshinterval)|Die Häufigkeit (in Sekunden), bei der der verknüpfte Datentyp aktualisiert wird, wenn `refreshMode` auf "periodisch" festgelegt ist.|
||[RefreshMode](/javascript/api/excel/excel.linkeddatatype#refreshmode)|Der Mechanismus, mit dem die Daten für den verknüpften Datentyp abgerufen werden.|
||[ServiceID](/javascript/api/excel/excel.linkeddatatype#serviceid)|Die eindeutige ID des verknüpften Datentyps.|
||[supportedRefreshModes](/javascript/api/excel/excel.linkeddatatype#supportedrefreshmodes)|Gibt ein Array mit allen vom verknüpften Datentyp unterstützten Aktualisierungsmodi zurück.|
||[requestRefresh()](/javascript/api/excel/excel.linkeddatatype#requestrefresh--)|Eine Anforderung zum Aktualisieren des verknüpften Datentyps wird erstellt.|
||[requestSetRefreshMode (RefreshMode: Excel. LinkedDataTypeRefreshMode)](/javascript/api/excel/excel.linkeddatatype#requestsetrefreshmode-refreshmode-)|Eine Anforderung zum Ändern des Aktualisierungsmodus für diesen verknüpften Datentyp wird angefordert.|
|[LinkedDataTypeAddedEventArgs](/javascript/api/excel/excel.linkeddatatypeaddedeventargs)|[ServiceID](/javascript/api/excel/excel.linkeddatatypeaddedeventargs#serviceid)|Die eindeutige ID des neuen verknüpften Daten Typs.|
||[source](/javascript/api/excel/excel.linkeddatatypeaddedeventargs#source)|Ruft die Quelle des Ereignisses ab.|
||[type](/javascript/api/excel/excel.linkeddatatypeaddedeventargs#type)|Ruft den Typ des Ereignisses ab.|
|[LinkedDataTypeCollection](/javascript/api/excel/excel.linkeddatatypecollection)|[getCount()](/javascript/api/excel/excel.linkeddatatypecollection#getcount--)|Ruft die Anzahl der verknüpften Datentypen in der Auflistung ab.|
||[GetItem (Schlüssel: Number)](/javascript/api/excel/excel.linkeddatatypecollection#getitem-key-)|Ruft einen verknüpften Datentyp nach Dienst-ID ab.|
||[getItemAt(index: number)](/javascript/api/excel/excel.linkeddatatypecollection#getitemat-index-)|Ruft einen verknüpften Datentyp anhand seines Index in der Auflistung ab.|
||[getItemOrNullObject (Key: Number)](/javascript/api/excel/excel.linkeddatatypecollection#getitemornullobject-key-)|Ruft einen verknüpften Datentyp anhand des ID-Typs ab.|
||[items](/javascript/api/excel/excel.linkeddatatypecollection#items)|Ruft die geladenen untergeordneten Elemente in dieser Sammlung ab.|
||[requestRefreshAll()](/javascript/api/excel/excel.linkeddatatypecollection#requestrefreshall--)|Erstellt eine Anforderung zum Aktualisieren aller verknüpften Datentypen in der Auflistung.|
|[NamedSheetView](/javascript/api/excel/excel.namedsheetview)|[activate()](/javascript/api/excel/excel.namedsheetview#activate--)|Aktiviert diese Blattansicht.|
||[delete()](/javascript/api/excel/excel.namedsheetview#delete--)|Entfernt die Blattansicht aus dem Arbeitsblatt.|
||[Duplikat (Name?: Zeichenfolge)](/javascript/api/excel/excel.namedsheetview#duplicate-name-)|Erstellt eine Kopie dieser Tabellenansicht.|
||[name](/javascript/api/excel/excel.namedsheetview#name)|Dient zum Abrufen oder Festlegen des Namens der Tabellenansicht.|
|[NamedSheetViewCollection](/javascript/api/excel/excel.namedsheetviewcollection)|[add(name: string)](/javascript/api/excel/excel.namedsheetviewcollection#add-name-)|Erstellt eine neue Tabellenansicht mit dem angegebenen Namen.|
||[enterTemporary()](/javascript/api/excel/excel.namedsheetviewcollection#entertemporary--)|Erstellt und aktiviert eine neue temporäre Arbeitsblattansicht.|
||[Exit ()](/javascript/api/excel/excel.namedsheetviewcollection#exit--)|Beendet die derzeit aktive Tabellenansicht.|
||[getactive ()](/javascript/api/excel/excel.namedsheetviewcollection#getactive--)|Ruft die derzeit aktive Tabellenansicht des Arbeitsblatts ab.|
||[getCount()](/javascript/api/excel/excel.namedsheetviewcollection#getcount--)|Ruft die Anzahl der Tabellen Ansichten in diesem Arbeitsblatt ab.|
||[getItem(key: string)](/javascript/api/excel/excel.namedsheetviewcollection#getitem-key-)|Ruft eine Tabellenansicht unter Verwendung des Namens ab.|
||[getItemAt(index: number)](/javascript/api/excel/excel.namedsheetviewcollection#getitemat-index-)|Ruft eine Blattansicht anhand ihres Index in der Auflistung ab.|
||[items](/javascript/api/excel/excel.namedsheetviewcollection#items)|Ruft die geladenen untergeordneten Elemente in dieser Sammlung ab.|
|[PivotLayout](/javascript/api/excel/excel.pivotlayout)|[altTextDescription](/javascript/api/excel/excel.pivotlayout#alttextdescription)|Die Alternativtext Beschreibung der PivotTable.|
||[altTextTitle](/javascript/api/excel/excel.pivotlayout#alttexttitle)|Der alt-Text Titel der PivotTable.|
||[displayBlankLineAfterEachItem (Display: Boolean)](/javascript/api/excel/excel.pivotlayout#displayblanklineaftereachitem-display-)|Legt fest, ob nach jedem Element eine leere Position angezeigt werden soll.|
||[emptyCellText](/javascript/api/excel/excel.pivotlayout#emptycelltext)|Der Text, der automatisch in eine beliebige leere Zelle in der PivotTable eingetragen wird, wenn `fillEmptyCells == true` .|
||[fillEmptyCells](/javascript/api/excel/excel.pivotlayout#fillemptycells)|Gibt an, ob leere Zellen in der PivotTable mit dem aufgefüllt werden sollen `emptyCellText` .|
||[getCell(dataHierarchy: DataPivotHierarchy \| string, rowItems: Array<PivotItem \| string>, columnItems: Array<PivotItem \| string>)](/javascript/api/excel/excel.pivotlayout#getcell-datahierarchy--rowitems--columnitems-)|Ruft eine eindeutige Zelle in der PivotTable ab, die auf einer Datenhierarchie und den Zeilen- und Spaltenelementen ihrer jeweiligen Hierarchie basiert.|
||[pivotstyle](/javascript/api/excel/excel.pivotlayout#pivotstyle)|Die Formatvorlage, die auf die PivotTable angewendet wird.|
||[repeatAllItemLabels (repeatLabels: Boolean)](/javascript/api/excel/excel.pivotlayout#repeatallitemlabels-repeatlabels-)|Legt die Einstellung "alle Elementbezeichnungen wiederholen" auf alle Felder in der PivotTable fest.|
||[SetStyle (Style: String \| pivottablestyle \| BuiltInPivotTableStyle)](/javascript/api/excel/excel.pivotlayout#setstyle-style-)|Legt die Formatvorlage fest, die auf die PivotTable angewendet wird.|
||[showFieldHeaders](/javascript/api/excel/excel.pivotlayout#showfieldheaders)|Gibt an, ob in der PivotTable Feld Header angezeigt werden (Feldbeschriftungen und Filter DropDowns).|
|[PivotTable](/javascript/api/excel/excel.pivottable)|[refreshOnOpen](/javascript/api/excel/excel.pivottable#refreshonopen)|Gibt an, ob die PivotTable beim Öffnen der Arbeitsmappe aktualisiert wird.|
|[Range](/javascript/api/excel/excel.range)|[getpräzedenzs ()](/javascript/api/excel/excel.range#getprecedents--)|Gibt ein `WorkbookRangeAreas` Objekt zurück, das den Bereich darstellt, der alle Vorgänger einer Zelle im gleichen Arbeitsblatt oder in mehreren Arbeitsblättern enthält.|
|[RefreshModeChangedEventArgs](/javascript/api/excel/excel.refreshmodechangedeventargs)|[RefreshMode](/javascript/api/excel/excel.refreshmodechangedeventargs#refreshmode)|Der Aktualisierungsmodus des verknüpften Daten Typs.|
||[ServiceID](/javascript/api/excel/excel.refreshmodechangedeventargs#serviceid)|Die eindeutige ID des Objekts, dessen Aktualisierungsmodus geändert wurde.|
||[source](/javascript/api/excel/excel.refreshmodechangedeventargs#source)|Ruft die Quelle des Ereignisses ab.|
||[type](/javascript/api/excel/excel.refreshmodechangedeventargs#type)|Ruft den Typ des Ereignisses ab.|
|[RefreshRequestCompletedEventArgs](/javascript/api/excel/excel.refreshrequestcompletedeventargs)|[Aktualisiert](/javascript/api/excel/excel.refreshrequestcompletedeventargs#refreshed)|Gibt an, ob die Anforderung zur Aktualisierung erfolgreich war.|
||[ServiceID](/javascript/api/excel/excel.refreshrequestcompletedeventargs#serviceid)|Die eindeutige ID des Objekts, dessen Aktualisierungsanforderung abgeschlossen wurde.|
||[source](/javascript/api/excel/excel.refreshrequestcompletedeventargs#source)|Ruft die Quelle des Ereignisses ab.|
||[type](/javascript/api/excel/excel.refreshrequestcompletedeventargs#type)|Ruft den Typ des Ereignisses ab.|
||[Warnungen](/javascript/api/excel/excel.refreshrequestcompletedeventargs#warnings)|Ein Array, das alle Warnungen enthält, die aus der Aktualisierungsanforderung generiert wurden.|
|[ShapeCollection](/javascript/api/excel/excel.shapecollection)|[addSvg(xml: string)](/javascript/api/excel/excel.shapecollection#addsvg-xml-)|Erstellt eine skalierbare Vektorgrafik (SVG) aus einer XML-Zeichenfolge und fügt sie dem Arbeitsblatt hinzu.|
|[Datenschnitt](/javascript/api/excel/excel.slicer)|[nameInFormula](/javascript/api/excel/excel.slicer#nameinformula)|Stellt den in der Formel verwendeten Namen des Datenschnitts dar.|
||[slicerStyle](/javascript/api/excel/excel.slicer#slicerstyle)|Die Formatvorlage, die auf den datenschnitt angewendet wird.|
||[SetStyle (Style: String \| SlicerStyle \| BuiltInSlicerStyle)](/javascript/api/excel/excel.slicer#setstyle-style-)|Legt die Formatvorlage fest, die auf den datenschnitt angewendet wird.|
|[Table](/javascript/api/excel/excel.table)|[clearStyle()](/javascript/api/excel/excel.table#clearstyle--)|Ändert die Tabelle so, dass sie die Standard-Tabellenformatvorlage verwendet.|
||[onFiltered](/javascript/api/excel/excel.table#onfiltered)|Tritt ein, wenn ein Filter auf eine bestimmte Tabelle angewendet wird.|
||[TableStyle](/javascript/api/excel/excel.table#tablestyle)|Die Formatvorlage, die auf die Tabelle angewendet wird.|
||[SetStyle (Style: String \| TableStyle \| BuiltInTableStyle)](/javascript/api/excel/excel.table#setstyle-style-)|Legt die Formatvorlage fest, die auf die Tabelle angewendet wird.|
|[TableCollection](/javascript/api/excel/excel.tablecollection)|[onFiltered](/javascript/api/excel/excel.tablecollection#onfiltered)|Tritt ein, wenn ein Filter auf eine beliebige Tabelle in einer Arbeitsmappe oder auf einem Arbeitsblatt angewendet wird.|
|[TableFilteredEventArgs](/javascript/api/excel/excel.tablefilteredeventargs)|[tableId](/javascript/api/excel/excel.tablefilteredeventargs#tableid)|Ruft die ID der Tabelle ab, in der der Filter angewendet wird.|
||[type](/javascript/api/excel/excel.tablefilteredeventargs#type)|Ruft den Typ des Ereignisses ab.|
||[worksheetId](/javascript/api/excel/excel.tablefilteredeventargs#worksheetid)|Ruft die ID des Arbeitsblatts ab, das die Tabelle enthält.|
|[Aufgabe](/javascript/api/excel/excel.task)|[addassignee (e-Mail: String)](/javascript/api/excel/excel.task#addassignee-email-)|Fügt dem Vorgang einen Empfänger hinzu.|
||[applyChanges (taskChanges: Excel. taskChanges)](/javascript/api/excel/excel.task#applychanges-taskchanges-)|Wendet die angegebenen Änderungen auf den Vorgang an.|
||[Empfängern](/javascript/api/excel/excel.task#assignees)|Ruft die Benutzer ab, denen die Aufgabe zugewiesen ist.|
||[comment](/javascript/api/excel/excel.task#comment)|Ruft den Kommentar ab, der dem Vorgang zugeordnet ist.|
||[DueDate](/javascript/api/excel/excel.task#duedate)|Ruft das Datum und die Uhrzeit der Fälligkeit des Vorgangs ab.|
||[historyRecords](/javascript/api/excel/excel.task#historyrecords)|Ruft die Verlaufsdatensätze des Vorgangs ab.|
||[id](/javascript/api/excel/excel.task#id)|Ruft die ID des Vorgangs ab.|
||[percentComplete](/javascript/api/excel/excel.task#percentcomplete)|Ruft den Abschlussprozentsatz des Vorgangs ab.|
||[priority](/javascript/api/excel/excel.task#priority)|Ruft die Priorität des Vorgangs ab.|
||[startDate](/javascript/api/excel/excel.task#startdate)|Ruft das Datum und die Uhrzeit ab, zu denen der Vorgang beginnen soll.|
||[title](/javascript/api/excel/excel.task#title)|Ruft den Titel der Aufgabe ab.|
||[removeAllAssignees()](/javascript/api/excel/excel.task#removeallassignees--)|Entfernt alle Empfänger aus dem Vorgang.|
||[removeAssignee (e-Mail: Zeichenfolge)](/javascript/api/excel/excel.task#removeassignee-email-)|Entfernt einen Empfänger aus dem Vorgang.|
||[setPercentComplete (PercentComplete: Number)](/javascript/api/excel/excel.task#setpercentcomplete-percentcomplete-)|Ändert den Abschluss der Aufgabe.|
||[setPriority (Priorität: Zahl)](/javascript/api/excel/excel.task#setpriority-priority-)|Ändert die Priorität des Vorgangs.|
||[setStartDateAndDueDate (StartDate: Date, DueDate: Date)](/javascript/api/excel/excel.task#setstartdateandduedate-startdate--duedate-)|Ändert das Start-und das Fälligkeitsdatum des Vorgangs.|
||[SetTitle (Title: String)](/javascript/api/excel/excel.task#settitle-title-)|Ändert den Titel des Vorgangs.|
|[TaskChanges](/javascript/api/excel/excel.taskchanges)|[DueDate](/javascript/api/excel/excel.taskchanges#duedate)|Legt ein neues Fälligkeitsdatum für den Vorgang in UTC-Zeitzone fest.|
||[emailsToAssign](/javascript/api/excel/excel.taskchanges#emailstoassign)|Legt die e-Mail-Adressen der Benutzer fest, die der Aufgabe zugewiesen werden sollen.|
||[emailsToUnassign](/javascript/api/excel/excel.taskchanges#emailstounassign)|Legt fest, dass die e-Mail-Adressen der Benutzer der Aufgabe zugewiesen werden sollen.|
||[percentComplete](/javascript/api/excel/excel.taskchanges#percentcomplete)|Legt einen neuen Fertigstellungsprozentsatz für den Vorgang fest.|
||[priority](/javascript/api/excel/excel.taskchanges#priority)|Legt eine neue Priorität für den Vorgang fest.|
||[removeAllPreviousAssignees](/javascript/api/excel/excel.taskchanges#removeallpreviousassignees)|Legt fest, ob die Änderung alle vorherigen Empfänger aus dem Vorgang entfernen soll.|
||[startDate](/javascript/api/excel/excel.taskchanges#startdate)|Legt in UTC-Zeitzone einen neuen Anfangstermin für den Vorgang fest.|
||[title](/javascript/api/excel/excel.taskchanges#title)|Legt einen neuen Titel für den Vorgang fest.|
|[Taskcollection](/javascript/api/excel/excel.taskcollection)|[getCount()](/javascript/api/excel/excel.taskcollection#getcount--)|Ruft die Anzahl der Aufgaben in der Auflistung ab.|
||[getItem(key: string)](/javascript/api/excel/excel.taskcollection#getitem-key-)|Ruft einen Vorgang mithilfe seiner ID ab.|
||[getItemAt(index: number)](/javascript/api/excel/excel.taskcollection#getitemat-index-)|Ruft eine Aufgabe nach dem Index in der Auflistung ab.|
||[getItemOrNullObject(key: string)](/javascript/api/excel/excel.taskcollection#getitemornullobject-key-)|Ruft einen Vorgang mithilfe seiner ID ab.|
||[items](/javascript/api/excel/excel.taskcollection#items)|Ruft die geladenen untergeordneten Elemente in dieser Sammlung ab.|
|[TaskHistoryRecord](/javascript/api/excel/excel.taskhistoryrecord)|[Verankerung](/javascript/api/excel/excel.taskhistoryrecord#anchorid)|Stellt die ID des Objekts dar, an dem der Vorgang verankert ist (z. b. Kommentierung für Aufgaben, die Kommentaren zugeordnet sind).|
||[Empfänger](/javascript/api/excel/excel.taskhistoryrecord#assignee)|Stellt den Benutzer dar, der dem Vorgang für einen Verlaufsdaten Satztyp "Assign" zugewiesen ist, oder der Benutzer, der die Zuweisung von der Aufgabe für den Verlaufsdaten Satztyp "Aufheben der Zuweisung" aufheben möchte.|
||[attributionUser](/javascript/api/excel/excel.taskhistoryrecord#attributionuser)|Stellt den Benutzer dar, der die Aufgabe erstellt oder geändert hat.|
||[DueDate](/javascript/api/excel/excel.taskhistoryrecord#duedate)|Stellt das Fälligkeitsdatum des Vorgangs dar.|
||[historyRecordCreatedDate](/javascript/api/excel/excel.taskhistoryrecord#historyrecordcreateddate)|Stellt das Erstellungsdatum des Vorgangs Verlaufsdaten Satzes dar.|
||[id](/javascript/api/excel/excel.taskhistoryrecord#id)|ID für den Verlaufsdatensatz.|
||[percentComplete](/javascript/api/excel/excel.taskhistoryrecord#percentcomplete)|Stellt den Fertigstellungsprozentsatz des Vorgangs dar.|
||[priority](/javascript/api/excel/excel.taskhistoryrecord#priority)|Stellt die Priorität des Vorgangs dar.|
||[startDate](/javascript/api/excel/excel.taskhistoryrecord#startdate)|Stellt den Anfangstermin des Vorgangs dar.|
||[title](/javascript/api/excel/excel.taskhistoryrecord#title)|Stellt den Titel des Vorgangs dar.|
||[type](/javascript/api/excel/excel.taskhistoryrecord#type)|Stellt den Typ des Aufgaben Verlaufsdaten Satzes dar.|
||[undoHistoryId](/javascript/api/excel/excel.taskhistoryrecord#undohistoryid)|Stellt die TaskHistoryRecord.ID-Eigenschaft dar, die für den Verlaufsdaten Satztyp "Undo" rückgängig gemacht wurde.|
|[TaskHistoryRecordCollection](/javascript/api/excel/excel.taskhistoryrecordcollection)|[getCount()](/javascript/api/excel/excel.taskhistoryrecordcollection#getcount--)|Ruft die Anzahl der Verlaufsdatensätze in der Auflistung für den Vorgang ab.|
||[getItemAt(index: number)](/javascript/api/excel/excel.taskhistoryrecordcollection#getitemat-index-)|Ruft einen Aufgaben Verlaufsdatensatz mithilfe seines Index in der Auflistung ab.|
||[items](/javascript/api/excel/excel.taskhistoryrecordcollection#items)|Ruft die geladenen untergeordneten Elemente in dieser Sammlung ab.|
|[Benutzende](/javascript/api/excel/excel.user)|[displayName](/javascript/api/excel/excel.user#displayname)|Stellt den Anzeigenamen des Benutzers dar.|
||[email](/javascript/api/excel/excel.user#email)|Stellt die E-Mail-Adresse des Benutzers dar.|
||[uid](/javascript/api/excel/excel.user#uid)|Stellt die eindeutige ID des Benutzers dar.|
|[Workbook](/javascript/api/excel/excel.workbook)|[linkedDataTypes](/javascript/api/excel/excel.workbook#linkeddatatypes)|Gibt eine Auflistung von verknüpften Datentypen zurück, die Teil der Arbeitsmappe sind.|
||[Aufgaben](/javascript/api/excel/excel.workbook#tasks)|Gibt eine Auflistung von Aufgaben zurück, die in der Arbeitsmappe vorhanden sind.|
||[showPivotFieldList](/javascript/api/excel/excel.workbook#showpivotfieldlist)|Gibt an, ob der Bereich Feldliste der PivotTable auf Arbeitsmappen Ebene angezeigt wird.|
||[use1904DateSystem](/javascript/api/excel/excel.workbook#use1904datesystem)|True, falls die Arbeitsmappe das 1904-Datumssystem verwendet.|
|[Arbeitsblatt](/javascript/api/excel/excel.worksheet)|[namedSheetViews](/javascript/api/excel/excel.worksheet#namedsheetviews)|Gibt eine Auflistung von Planansichten zurück, die im Arbeitsblatt vorhanden sind.|
||[onFiltered](/javascript/api/excel/excel.worksheet#onfiltered)|Tritt ein, wenn ein Filter auf ein bestimmtes Arbeitsblatt angewendet wird.|
||[Aufgaben](/javascript/api/excel/excel.worksheet#tasks)|Gibt eine Auflistung von Aufgaben zurück, die im Arbeitsblatt vorhanden sind.|
|[WorksheetCollection](/javascript/api/excel/excel.worksheetcollection)|[addFromBase64(base64File: string, sheetNamesToInsert?: string[], positionType?: Excel.WorksheetPositionType, relativeTo?: Worksheet \| string)](/javascript/api/excel/excel.worksheetcollection#addfrombase64-base64file--sheetnamestoinsert--positiontype--relativeto-)|Fügt die angegebenen Arbeitsblätter einer Arbeitsmappe in die aktuelle Arbeitsmappe ein.|
||[onFiltered](/javascript/api/excel/excel.worksheetcollection#onfiltered)|Tritt ein, wenn ein Filter eines beliebigen Arbeitsblatts in der Arbeitsmappe angewendet wird.|
|[WorksheetFilteredEventArgs](/javascript/api/excel/excel.worksheetfilteredeventargs)|[type](/javascript/api/excel/excel.worksheetfilteredeventargs#type)|Ruft den Typ des Ereignisses ab.|
||[worksheetId](/javascript/api/excel/excel.worksheetfilteredeventargs#worksheetid)|Ruft die ID des Arbeitsblatts ab, in dem der Filter angewendet wird.|
