| Klasse | Felder | Beschreibung |
|:---|:---|:---|
|[AutoFilter](/javascript/api/excel/excel.autofilter)|[clearColumnCriteria(columnIndex: number)](/javascript/api/excel/excel.autofilter#clearcolumncriteria-columnindex-)|Löscht die Filterkriterien von AutoFilter.|
|[ChangeDirectionState](/javascript/api/excel/excel.changedirectionstate)|[deleteShiftDirection](/javascript/api/excel/excel.changedirectionstate#deleteshiftdirection)|Stellt die Richtung (z. B. nach oben oder links) dar, in der die verbleibenden Zellen verschoben werden, wenn eine Zelle oder Zellen gelöscht werden.|
||[insertShiftDirection](/javascript/api/excel/excel.changedirectionstate#insertshiftdirection)|Stellt die Richtung (z. B. nach unten oder rechts) dar, in die sich die vorhandenen Zellen verschieben, wenn eine oder mehrere neue Zellen eingefügt werden.|
|[Kommentar](/javascript/api/excel/excel.comment)|[assignTask(assignee: Identity)](/javascript/api/excel/excel.comment#assigntask-assignee-)|Weist die Aufgabe, die dem Kommentar zugeordnet ist, dem angegebenen Benutzer als Zugewiesener zu.|
||[getTask()](/javascript/api/excel/excel.comment#gettask--)|Ruft den Vorgang ab, der diesem Kommentar zugeordnet ist.|
||[getTaskOrNullObject()](/javascript/api/excel/excel.comment#gettaskornullobject--)|Ruft den Vorgang ab, der diesem Kommentar zugeordnet ist.|
|[CommentCollection](/javascript/api/excel/excel.commentcollection)|[getItemOrNullObject(commentId: string)](/javascript/api/excel/excel.commentcollection#getitemornullobject-commentid-)|Ruft einen Kommentar aus der Sammlung basierend auf der ID ab.|
|[CommentReply](/javascript/api/excel/excel.commentreply)|[assignTask(assignee: Identity)](/javascript/api/excel/excel.commentreply#assigntask-assignee-)|Weist die dem Kommentar zugeordnete Aufgabe dem angegebenen Benutzer als alleiniger Zugewiesener zu.|
||[getTask()](/javascript/api/excel/excel.commentreply#gettask--)|Ruft den Vorgang ab, der dem Thread dieser Kommentarantwort zugeordnet ist.|
||[getTaskOrNullObject()](/javascript/api/excel/excel.commentreply#gettaskornullobject--)|Ruft den Vorgang ab, der dem Thread dieser Kommentarantwort zugeordnet ist.|
|[CommentReplyCollection](/javascript/api/excel/excel.commentreplycollection)|[getItemOrNullObject(commentReplyId: string)](/javascript/api/excel/excel.commentreplycollection#getitemornullobject-commentreplyid-)|Gibt eine Kommentarantwort zurück, die durch ihre ID angegeben ist.|
|[ConditionalFormatCollection](/javascript/api/excel/excel.conditionalformatcollection)|[getItemOrNullObject(id: string)](/javascript/api/excel/excel.conditionalformatcollection#getitemornullobject-id-)|Gibt ein bedingtes Format zurück, das durch seine ID identifiziert wird.|
|[DocumentTask](/javascript/api/excel/excel.documenttask)|[Percentcomplete](/javascript/api/excel/excel.documenttask#percentcomplete)|Gibt den Prozentsatz der Fertigstellung des Vorgangs an.|
||[Priorität](/javascript/api/excel/excel.documenttask#priority)|Gibt die Priorität des Vorgangs an.|
||[Erfüllungsgehilfen](/javascript/api/excel/excel.documenttask#assignees)|Gibt eine Auflistung von Zugewiesenen der Aufgabe zurück.|
||[Änderungen](/javascript/api/excel/excel.documenttask#changes)|Ruft die Änderungsdatensätze des Vorgangs ab.|
||[comment](/javascript/api/excel/excel.documenttask#comment)|Ruft den Kommentar ab, der dem Vorgang zugeordnet ist.|
||[completedBy](/javascript/api/excel/excel.documenttask#completedby)|Ruft den letzten Benutzer ab, der die Aufgabe abgeschlossen hat.|
||[completedDateTime](/javascript/api/excel/excel.documenttask#completeddatetime)|Ruft das Datum und die Uhrzeit der Fertigstellung des Vorgangs ab.|
||[createdBy](/javascript/api/excel/excel.documenttask#createdby)|Ruft den Benutzer ab, der die Aufgabe erstellt hat.|
||[createdDateTime](/javascript/api/excel/excel.documenttask#createddatetime)|Ruft das Datum und die Uhrzeit der Erstellung des Vorgangs ab.|
||[id](/javascript/api/excel/excel.documenttask#id)|Ruft die ID des Vorgangs ab.|
||[setStartAndDueDateTime(startDateTime: Date, dueDateTime: Date)](/javascript/api/excel/excel.documenttask#setstartandduedatetime-startdatetime--duedatetime-)|Ändert das Anfangs- und Fälligkeitsdatum des Vorgangs.|
||[startAndDueDateTime](/javascript/api/excel/excel.documenttask#startandduedatetime)|Dient zum Abrufen oder Festlegen des Datums und der Uhrzeit, zu der der Vorgang gestartet werden soll und fällig ist.|
||[title](/javascript/api/excel/excel.documenttask#title)|Gibt den Titel der Aufgabe an.|
|[DocumentTaskChange](/javascript/api/excel/excel.documenttaskchange)|[Zessionar](/javascript/api/excel/excel.documenttaskchange#assignee)|Stellt den Benutzer dar, der der Aufgabe für einen `assign` Änderungsdatensatztyp zugewiesen ist, oder der Benutzer, der der Aufgabe für einen Änderungsdatensatztyp nicht zugewiesen `unassign` wurde.|
||[changedBy](/javascript/api/excel/excel.documenttaskchange#changedby)|Stellt den Benutzer dar, der die Aufgabe erstellt oder geändert hat.|
||[commentId](/javascript/api/excel/excel.documenttaskchange#commentid)|Stellt die ID des Vorgangs `Comment` dar, mit `CommentReply` dem die Vorgangsänderung verankert ist.|
||[createdDateTime](/javascript/api/excel/excel.documenttaskchange#createddatetime)|Stellt das Erstellungsdatum und die Uhrzeit des Vorgangsänderungsdatensatzes dar.|
||[dueDateTime](/javascript/api/excel/excel.documenttaskchange#duedatetime)|Stellt das Fälligkeitsdatum und die Uhrzeit der Aufgabe in UTC-Zeitzone dar.|
||[id](/javascript/api/excel/excel.documenttaskchange#id)|ID für den Vorgangsänderungsdatensatz.|
||[Percentcomplete](/javascript/api/excel/excel.documenttaskchange#percentcomplete)|Stellt den Prozentsatz der Fertigstellung des Vorgangs dar.|
||[Priorität](/javascript/api/excel/excel.documenttaskchange#priority)|Stellt die Priorität des Vorgangs dar.|
||[startDateTime](/javascript/api/excel/excel.documenttaskchange#startdatetime)|Stellt das Startdatum und die Startzeit des Vorgangs in UTC-Zeitzone dar.|
||[title](/javascript/api/excel/excel.documenttaskchange#title)|Stellt den Titel des Vorgangs dar.|
||[type](/javascript/api/excel/excel.documenttaskchange#type)|Stellt den Aktionstyp des Vorgangsänderungsdatensatzes dar.|
||[undoHistoryId](/javascript/api/excel/excel.documenttaskchange#undohistoryid)|Stellt die `DocumentTaskChange.id` Eigenschaft dar, die für den Änderungsdatensatztyp rückgängig macht `undo` wurde.|
|[DocumentTaskChangeCollection](/javascript/api/excel/excel.documenttaskchangecollection)|[getCount()](/javascript/api/excel/excel.documenttaskchangecollection#getcount--)|Ruft die Anzahl der Änderungsdatensätze in der Auflistung für den Vorgang ab.|
||[getItemAt(index: number)](/javascript/api/excel/excel.documenttaskchangecollection#getitemat-index-)|Ruft einen Vorgangsänderungsdatensatz mithilfe seines Indexes in der Auflistung ab.|
||[items](/javascript/api/excel/excel.documenttaskchangecollection#items)|Ruft die geladenen untergeordneten Elemente in dieser Sammlung ab.|
|[DocumentTaskCollection](/javascript/api/excel/excel.documenttaskcollection)|[getCount()](/javascript/api/excel/excel.documenttaskcollection#getcount--)|Ruft die Anzahl der Vorgänge in der Auflistung ab.|
||[getItem(key: string)](/javascript/api/excel/excel.documenttaskcollection#getitem-key-)|Ruft einen Vorgang mithilfe seiner ID ab.|
||[getItemAt(index: number)](/javascript/api/excel/excel.documenttaskcollection#getitemat-index-)|Ruft einen Vorgang anhand seines Indexes in der Auflistung ab.|
||[getItemOrNullObject(key: string)](/javascript/api/excel/excel.documenttaskcollection#getitemornullobject-key-)|Ruft einen Vorgang mithilfe seiner ID ab.|
||[items](/javascript/api/excel/excel.documenttaskcollection#items)|Ruft die geladenen untergeordneten Elemente in dieser Sammlung ab.|
|[DocumentTaskSchedule](/javascript/api/excel/excel.documenttaskschedule)|[dueDateTime](/javascript/api/excel/excel.documenttaskschedule#duedatetime)|Ruft das Datum und die Uhrzeit der Fälligkeit des Vorgangs ab.|
||[startDateTime](/javascript/api/excel/excel.documenttaskschedule#startdatetime)|Ruft das Datum und die Uhrzeit ab, zu denen der Vorgang beginnen soll.|
|[GroupShapeCollection](/javascript/api/excel/excel.groupshapecollection)|[getItemOrNullObject(key: string)](/javascript/api/excel/excel.groupshapecollection#getitemornullobject-key-)|Ruft ein Shape anhand seines Namens oder seiner ID ab.|
|[Identity](/javascript/api/excel/excel.identity)|[displayName](/javascript/api/excel/excel.identity#displayname)|Stellt den Anzeigenamen des Benutzers dar.|
||[email](/javascript/api/excel/excel.identity#email)|Stellt die E-Mail-Adresse des Benutzers dar.|
||[id](/javascript/api/excel/excel.identity#id)|Stellt die eindeutige ID des Benutzers dar.|
|[IdentityCollection](/javascript/api/excel/excel.identitycollection)|[add(assignee: Identity)](/javascript/api/excel/excel.identitycollection#add-assignee-)|Fügt der Auflistung eine Benutzeridentität hinzu.|
||[clear()](/javascript/api/excel/excel.identitycollection#clear--)|Entfernt alle Benutzeridentitäten aus der Auflistung.|
||[getCount()](/javascript/api/excel/excel.identitycollection#getcount--)|Ruft die Anzahl der Elemente in der Auflistung ab.|
||[getItemAt(index: number)](/javascript/api/excel/excel.identitycollection#getitemat-index-)|Ruft eine Dokumentbenutzeridentität mithilfe des Indexes in der Auflistung ab.|
||[items](/javascript/api/excel/excel.identitycollection#items)|Ruft die geladenen untergeordneten Elemente in dieser Sammlung ab.|
||[remove(assignee: Identity)](/javascript/api/excel/excel.identitycollection#remove-assignee-)|Entfernt eine Benutzeridentität aus der Auflistung.|
|[IdentityEntity](/javascript/api/excel/excel.identityentity)|[displayName](/javascript/api/excel/excel.identityentity#displayname)|Stellt den Anzeigenamen des Benutzers dar.|
||[email](/javascript/api/excel/excel.identityentity#email)|Stellt die E-Mail-Adresse des Benutzers dar.|
||[id](/javascript/api/excel/excel.identityentity#id)|Stellt die eindeutige ID des Benutzers dar.|
|[LinkedDataType](/javascript/api/excel/excel.linkeddatatype)|[Dataprovider](/javascript/api/excel/excel.linkeddatatype#dataprovider)|Der Name des Datenanbieters für den verknüpften Datentyp.|
||[lastRefreshed](/javascript/api/excel/excel.linkeddatatype#lastrefreshed)|Datum und Uhrzeit der lokalen Zeitzone seit dem Öffnen der Arbeitsmappe, als der verknüpfte Datentyp zuletzt aktualisiert wurde.|
||[name](/javascript/api/excel/excel.linkeddatatype#name)|Der Name des verknüpften Datentyps.|
||[periodicRefreshInterval](/javascript/api/excel/excel.linkeddatatype#periodicrefreshinterval)|Die Häufigkeit in Sekunden, mit der der verknüpfte Datentyp aktualisiert wird, wenn `refreshMode` er auf "Periodisch" festgelegt ist.|
||[Refreshmode](/javascript/api/excel/excel.linkeddatatype#refreshmode)|Der Mechanismus, mit dem die Daten für den verknüpften Datentyp abgerufen werden.|
||[ServiceID](/javascript/api/excel/excel.linkeddatatype#serviceid)|Die eindeutige ID des verknüpften Datentyps.|
||[supportedRefreshModes](/javascript/api/excel/excel.linkeddatatype#supportedrefreshmodes)|Gibt ein Array mit allen Aktualisierungsmodi zurück, die vom verknüpften Datentyp unterstützt werden.|
||[requestRefresh()](/javascript/api/excel/excel.linkeddatatype#requestrefresh--)|Stellt eine Anforderung zum Aktualisieren des verknüpften Datentyps.|
||[requestSetRefreshMode(refreshMode: Excel. LinkedDataTypeRefreshMode)](/javascript/api/excel/excel.linkeddatatype#requestsetrefreshmode-refreshmode-)|Stellt eine Anforderung zum Ändern des Aktualisierungsmodus für diesen verknüpften Datentyp.|
|[LinkedDataTypeAddedEventArgs](/javascript/api/excel/excel.linkeddatatypeaddedeventargs)|[ServiceID](/javascript/api/excel/excel.linkeddatatypeaddedeventargs#serviceid)|Die eindeutige ID des neuen verknüpften Datentyps.|
||[source](/javascript/api/excel/excel.linkeddatatypeaddedeventargs#source)|Ruft die Quelle des Ereignisses ab.|
||[type](/javascript/api/excel/excel.linkeddatatypeaddedeventargs#type)|Ruft den Typ des Ereignisses ab.|
|[LinkedDataTypeCollection](/javascript/api/excel/excel.linkeddatatypecollection)|[getCount()](/javascript/api/excel/excel.linkeddatatypecollection#getcount--)|Ruft die Anzahl der verknüpften Datentypen in der Auflistung ab.|
||[getItem(key: number)](/javascript/api/excel/excel.linkeddatatypecollection#getitem-key-)|Ruft einen verknüpften Datentyp nach Dienst-ID ab.|
||[getItemAt(index: number)](/javascript/api/excel/excel.linkeddatatypecollection#getitemat-index-)|Ruft einen verknüpften Datentyp anhand seines Indexes in der Auflistung ab.|
||[getItemOrNullObject(key: number)](/javascript/api/excel/excel.linkeddatatypecollection#getitemornullobject-key-)|Ruft einen verknüpften Datentyp nach ID ab.|
||[items](/javascript/api/excel/excel.linkeddatatypecollection#items)|Ruft die geladenen untergeordneten Elemente in dieser Sammlung ab.|
||[requestRefreshAll()](/javascript/api/excel/excel.linkeddatatypecollection#requestrefreshall--)|Stellt eine Anforderung zum Aktualisieren aller verknüpften Datentypen in der Auflistung.|
|[LinkedWorkbook](/javascript/api/excel/excel.linkedworkbook)|[breakLinks()](/javascript/api/excel/excel.linkedworkbook#breaklinks--)|Stellt eine Anforderung, um die Verknüpfungen zu unterbrechen, die auf die verknüpfte Arbeitsmappe zeigen.|
||[id](/javascript/api/excel/excel.linkedworkbook#id)|Die ursprüngliche URL, die auf die verknüpfte Arbeitsmappe verweist.|
||[refresh()](/javascript/api/excel/excel.linkedworkbook#refresh--)|Stellt eine Anforderung zum Aktualisieren der aus der verknüpften Arbeitsmappe abgerufenen Daten.|
|[LinkedWorkbookCollection](/javascript/api/excel/excel.linkedworkbookcollection)|[breakAllLinks()](/javascript/api/excel/excel.linkedworkbookcollection#breakalllinks--)|Unterbricht alle Verknüpfungen zu den verknüpften Arbeitsmappen.|
||[getItem(key: string)](/javascript/api/excel/excel.linkedworkbookcollection#getitem-key-)|Ruft Informationen zu einer verknüpften Arbeitsmappe anhand ihrer URL ab.|
||[getItemOrNullObject(key: string)](/javascript/api/excel/excel.linkedworkbookcollection#getitemornullobject-key-)|Ruft Informationen zu einer verknüpften Arbeitsmappe anhand ihrer URL ab.|
||[items](/javascript/api/excel/excel.linkedworkbookcollection#items)|Ruft die geladenen untergeordneten Elemente in dieser Sammlung ab.|
||[refreshAll()](/javascript/api/excel/excel.linkedworkbookcollection#refreshall--)|Stellt eine Anforderung zum Aktualisieren aller Arbeitsmappenlinks.|
||[workbookLinksRefreshMode](/javascript/api/excel/excel.linkedworkbookcollection#workbooklinksrefreshmode)|Stellt den Aktualisierungsmodus der Arbeitsmappenlinks dar.|
|[NamedSheetViewCollection](/javascript/api/excel/excel.namedsheetviewcollection)|[getItemOrNullObject(key: string)](/javascript/api/excel/excel.namedsheetviewcollection#getitemornullobject-key-)|Ruft eine Blattansicht mit ihrem Namen ab.|
|[PivotLayout](/javascript/api/excel/excel.pivotlayout)|[getCell(dataHierarchy: DataPivotHierarchy \| string, rowItems: Array<PivotItem \| string>, columnItems: Array<PivotItem \| string>)](/javascript/api/excel/excel.pivotlayout#getcell-datahierarchy--rowitems--columnitems-)|Ruft eine eindeutige Zelle in der PivotTable ab, die auf einer Datenhierarchie und den Zeilen- und Spaltenelementen ihrer jeweiligen Hierarchie basiert.|
||[pivotStyle](/javascript/api/excel/excel.pivotlayout#pivotstyle)|Die auf die PivotTable angewendete Formatvorlage.|
||[setStyle(style: string \| PivotTableStyle \| BuiltInPivotTableStyle)](/javascript/api/excel/excel.pivotlayout#setstyle-style-)|Legt die Formatvorlage fest, die auf die PivotTable angewendet wird.|
|[PivotTableScopedCollection](/javascript/api/excel/excel.pivottablescopedcollection)|[getFirstOrNullObject()](/javascript/api/excel/excel.pivottablescopedcollection#getfirstornullobject--)|Ruft die erste PivotTable in der Auflistung ab.|
|[Range](/javascript/api/excel/excel.range)|[getDependents()](/javascript/api/excel/excel.range#getdependents--)|Gibt ein `WorkbookRangeAreas` Objekt zurück, das den Bereich darstellt, der alle Abhängigen einer Zelle im selben Arbeitsblatt oder in mehreren Arbeitsblättern enthält.|
||[getPrecedents()](/javascript/api/excel/excel.range#getprecedents--)|Gibt ein `WorkbookRangeAreas` Objekt zurück, das den Bereich darstellt, der alle Vorgänger einer Zelle im selben Arbeitsblatt oder in mehreren Arbeitsblättern enthält.|
|[RefreshModeChangedEventArgs](/javascript/api/excel/excel.refreshmodechangedeventargs)|[Refreshmode](/javascript/api/excel/excel.refreshmodechangedeventargs#refreshmode)|Der Aktualisierungsmodus für den verknüpften Datentyp.|
||[ServiceID](/javascript/api/excel/excel.refreshmodechangedeventargs#serviceid)|Die eindeutige ID des Objekts, dessen Aktualisierungsmodus geändert wurde.|
||[source](/javascript/api/excel/excel.refreshmodechangedeventargs#source)|Ruft die Quelle des Ereignisses ab.|
||[type](/javascript/api/excel/excel.refreshmodechangedeventargs#type)|Ruft den Typ des Ereignisses ab.|
|[RefreshRequestCompletedEventArgs](/javascript/api/excel/excel.refreshrequestcompletedeventargs)|[Aktualisiert](/javascript/api/excel/excel.refreshrequestcompletedeventargs#refreshed)|Gibt an, ob die Aktualisierungsanforderung erfolgreich war.|
||[ServiceID](/javascript/api/excel/excel.refreshrequestcompletedeventargs#serviceid)|Die eindeutige ID des Objekts, dessen Aktualisierungsanforderung abgeschlossen wurde.|
||[source](/javascript/api/excel/excel.refreshrequestcompletedeventargs#source)|Ruft die Quelle des Ereignisses ab.|
||[type](/javascript/api/excel/excel.refreshrequestcompletedeventargs#type)|Ruft den Typ des Ereignisses ab.|
||[Warnungen](/javascript/api/excel/excel.refreshrequestcompletedeventargs#warnings)|Ein Array, das alle Warnungen enthält, die aus der Aktualisierungsanforderung generiert werden.|
|[ShapeCollection](/javascript/api/excel/excel.shapecollection)|[addSvg(xml: string)](/javascript/api/excel/excel.shapecollection#addsvg-xml-)|Erstellt eine skalierbare Vektorgrafik (SVG) aus einer XML-Zeichenfolge und fügt sie dem Arbeitsblatt hinzu.|
||[getItemOrNullObject(key: string)](/javascript/api/excel/excel.shapecollection#getitemornullobject-key-)|Ruft ein Shape anhand seines Namens oder seiner ID ab.|
|[Datenschnitt](/javascript/api/excel/excel.slicer)|[nameInFormula](/javascript/api/excel/excel.slicer#nameinformula)|Stellt den in der Formel verwendeten Namen des Datenschnitts dar.|
||[slicerStyle](/javascript/api/excel/excel.slicer#slicerstyle)|Die auf den Datenschnitt angewendete Formatvorlage.|
||[setStyle(style: string \| SlicerStyle \| BuiltInSlicerStyle)](/javascript/api/excel/excel.slicer#setstyle-style-)|Legt die Formatvorlage fest, die auf den Datenschnitt angewendet wird.|
|[StyleCollection](/javascript/api/excel/excel.stylecollection)|[GetItemOrNullObject(name: string)](/javascript/api/excel/excel.stylecollection#getitemornullobject-name-)|Ruft eine Formatvorlage anhand des Namens ab.|
|[Table](/javascript/api/excel/excel.table)|[clearStyle()](/javascript/api/excel/excel.table#clearstyle--)|Ändert die Tabelle so, dass sie die Standard-Tabellenformatvorlage verwendet.|
||[onFiltered](/javascript/api/excel/excel.table#onfiltered)|Tritt auf, wenn ein Filter auf eine bestimmte Tabelle angewendet wird.|
||[Tablestyle](/javascript/api/excel/excel.table#tablestyle)|Die auf die Tabelle angewendete Formatvorlage.|
||[setStyle(style: string \| TableStyle \| BuiltInTableStyle)](/javascript/api/excel/excel.table#setstyle-style-)|Legt die Formatvorlage fest, die auf die Tabelle angewendet wird.|
|[TableCollection](/javascript/api/excel/excel.tablecollection)|[onFiltered](/javascript/api/excel/excel.tablecollection#onfiltered)|Tritt auf, wenn ein Filter auf eine beliebige Tabelle in einer Arbeitsmappe oder einem Arbeitsblatt angewendet wird.|
|[TableFilteredEventArgs](/javascript/api/excel/excel.tablefilteredeventargs)|[tableId](/javascript/api/excel/excel.tablefilteredeventargs#tableid)|Ruft die ID der Tabelle ab, in der der Filter angewendet wird.|
||[type](/javascript/api/excel/excel.tablefilteredeventargs#type)|Ruft den Typ des Ereignisses ab.|
||[worksheetId](/javascript/api/excel/excel.tablefilteredeventargs#worksheetid)|Ruft die ID des Arbeitsblatts ab, das die Tabelle enthält.|
|[TableScopedCollection](/javascript/api/excel/excel.tablescopedcollection)|[getItemOrNullObject(key: string)](/javascript/api/excel/excel.tablescopedcollection#getitemornullobject-key-)|Ruft eine Tabelle anhand des Namens oder der ID ab.|
|[Workbook](/javascript/api/excel/excel.workbook)|[linkedDataTypes](/javascript/api/excel/excel.workbook#linkeddatatypes)|Gibt eine Auflistung verknüpfter Datentypen zurück, die Teil der Arbeitsmappe sind.|
||[linkedWorkbooks](/javascript/api/excel/excel.workbook#linkedworkbooks)|Gibt eine Auflistung verknüpfter Arbeitsmappen zurück.|
||[Aufgaben](/javascript/api/excel/excel.workbook#tasks)|Gibt eine Auflistung von Aufgaben zurück, die in der Arbeitsmappe vorhanden sind.|
||[showPivotFieldList](/javascript/api/excel/excel.workbook#showpivotfieldlist)|Gibt an, ob der Feldlistenbereich der PivotTable auf Arbeitsmappenebene angezeigt wird.|
||[use1904DateSystem](/javascript/api/excel/excel.workbook#use1904datesystem)|True, falls die Arbeitsmappe das 1904-Datumssystem verwendet.|
|[Arbeitsblatt](/javascript/api/excel/excel.worksheet)|[onFiltered](/javascript/api/excel/excel.worksheet#onfiltered)|Tritt auf, wenn ein Filter auf ein bestimmtes Arbeitsblatt angewendet wird.|
||[onProtectionChanged](/javascript/api/excel/excel.worksheet#onprotectionchanged)|Tritt auf, wenn der Arbeitsblattschutzstatus geändert wird.|
||[tabId](/javascript/api/excel/excel.worksheet#tabid)|Gibt einen Wert zurück, der dieses Arbeitsblatt darstellt, das von Open Office XML gelesen werden kann.|
||[Aufgaben](/javascript/api/excel/excel.worksheet#tasks)|Gibt eine Auflistung von Aufgaben zurück, die im Arbeitsblatt vorhanden sind.|
|[WorksheetChangedEventArgs](/javascript/api/excel/excel.worksheetchangedeventargs)|[changeDirectionState](/javascript/api/excel/excel.worksheetchangedeventargs#changedirectionstate)|Stellt eine Änderung der Richtung dar, in der die Zellen in einem Arbeitsblatt verschoben werden, wenn eine Zelle oder Zellen gelöscht oder eingefügt werden.|
||[triggerSource](/javascript/api/excel/excel.worksheetchangedeventargs#triggersource)|Stellt die Triggerquelle des Ereignisses dar.|
|[WorksheetCollection](/javascript/api/excel/excel.worksheetcollection)|[addFromBase64(base64File: string, sheetNamesToInsert?: string[], positionType?: Excel.WorksheetPositionType, relativeTo?: Worksheet \| string)](/javascript/api/excel/excel.worksheetcollection#addfrombase64-base64file--sheetnamestoinsert--positiontype--relativeto-)|Fügt die angegebenen Arbeitsblätter einer Arbeitsmappe in die aktuelle Arbeitsmappe ein.|
||[onFiltered](/javascript/api/excel/excel.worksheetcollection#onfiltered)|Tritt ein, wenn ein Filter eines beliebigen Arbeitsblatts in der Arbeitsmappe angewendet wird.|
||[onProtectionChanged](/javascript/api/excel/excel.worksheetcollection#onprotectionchanged)|Tritt auf, wenn der Arbeitsblattschutzstatus geändert wird.|
|[WorksheetFilteredEventArgs](/javascript/api/excel/excel.worksheetfilteredeventargs)|[type](/javascript/api/excel/excel.worksheetfilteredeventargs#type)|Ruft den Typ des Ereignisses ab.|
||[worksheetId](/javascript/api/excel/excel.worksheetfilteredeventargs#worksheetid)|Ruft die ID des Arbeitsblatts ab, auf das der Filter angewendet wird.|
|[WorksheetProtectionChangedEventArgs](/javascript/api/excel/excel.worksheetprotectionchangedeventargs)|[isProtected](/javascript/api/excel/excel.worksheetprotectionchangedeventargs#isprotected)|Ruft den aktuellen Schutzstatus des Arbeitsblatts ab.|
||[source](/javascript/api/excel/excel.worksheetprotectionchangedeventargs#source)|Die Quelle des Ereignisses.|
||[type](/javascript/api/excel/excel.worksheetprotectionchangedeventargs#type)|Ruft den Typ des Ereignisses ab.|
||[worksheetId](/javascript/api/excel/excel.worksheetprotectionchangedeventargs#worksheetid)|Ruft die ID des Arbeitsblatts ab, in dem der Schutzstatus geändert wird.|
