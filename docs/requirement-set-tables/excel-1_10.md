| Klasse | Felder | Beschreibung |
|:---|:---|:---|
|[Kommentar](/javascript/api/excel/excel.comment)|[content](/javascript/api/excel/excel.comment#content)|Der Inhalt des Kommentars.|
||[delete()](/javascript/api/excel/excel.comment#delete--)|Löscht den Kommentar und alle verbundenen Antworten.|
||[getLocation()](/javascript/api/excel/excel.comment#getlocation--)|Ruft die Zelle ab, in der sich dieser Kommentar befindet.|
||[authorEmail](/javascript/api/excel/excel.comment#authoremail)|Ruft die E-Mail des Autors des Kommentars ab.|
||[authorName](/javascript/api/excel/excel.comment#authorname)|Ruft den Namen des Autors des Kommentars ab.|
||[creationDate](/javascript/api/excel/excel.comment#creationdate)|Ruft den Erstellungszeitpunkt des Kommentars ab.|
||[id](/javascript/api/excel/excel.comment#id)|Gibt den Kommentarbezeichner an.|
||[replies](/javascript/api/excel/excel.comment#replies)|Stellt eine Sammlung der Antwortobjekte dar, die dem Kommentar zugeordnet sind.|
|[CommentCollection](/javascript/api/excel/excel.commentcollection)|[add(cellAddress: Range \| string, content: string, contentType?: Excel.ContentType)](/javascript/api/excel/excel.commentcollection#add-celladdress--content--contenttype-)|Erstellt einen neuen Kommentar mit dem angegebenen Inhalt auf der angegebenen Zelle.|
||[getCount()](/javascript/api/excel/excel.commentcollection#getcount--)|Ruft die Anzahl der Kommentare in der Sammlung ab.|
||[getItem(commentId: string)](/javascript/api/excel/excel.commentcollection#getitem-commentid-)|Ruft einen Kommentar aus der Sammlung basierend auf der ID ab.|
||[getItemAt(index: number)](/javascript/api/excel/excel.commentcollection#getitemat-index-)|Ruft einen Kommentar aus der Sammlung basierend auf ihrer Position ab.|
||[getItemByCell(cellAddress: Range \| string)](/javascript/api/excel/excel.commentcollection#getitembycell-celladdress-)|Ruft den Kommentar aus der angegebenen Zelle ab.|
||[getItemByReplyId(replyId: string)](/javascript/api/excel/excel.commentcollection#getitembyreplyid-replyid-)|Ruft den Kommentar ab, mit dem die angegebene Antwort verbunden ist.|
||[items](/javascript/api/excel/excel.commentcollection#items)|Ruft die geladenen untergeordneten Elemente in dieser Sammlung ab.|
|[CommentReply](/javascript/api/excel/excel.commentreply)|[content](/javascript/api/excel/excel.commentreply#content)|Der Inhalt der Kommentarantwort.|
||[delete()](/javascript/api/excel/excel.commentreply#delete--)|Löscht die Kommentarantwort.|
||[getLocation()](/javascript/api/excel/excel.commentreply#getlocation--)|Ruft die Zelle ab, in der sich diese Kommentarantwort befindet.|
||[getParentComment()](/javascript/api/excel/excel.commentreply#getparentcomment--)|Ruft den übergeordneten Kommentar dieser Antwort ab.|
||[authorEmail](/javascript/api/excel/excel.commentreply#authoremail)|Ruft die E-Mail des Autors der Kommentarantwort ab.|
||[authorName](/javascript/api/excel/excel.commentreply#authorname)|Ruft den Namen des Autors der Kommentarantwort ab.|
||[creationDate](/javascript/api/excel/excel.commentreply#creationdate)|Ruft den Erstellungszeitpunkt der Kommentarantwort ab.|
||[id](/javascript/api/excel/excel.commentreply#id)|Gibt den Kommentarantwortbezeichner an.|
|[CommentReplyCollection](/javascript/api/excel/excel.commentreplycollection)|[add(content: string, contentType?: Excel.ContentType)](/javascript/api/excel/excel.commentreplycollection#add-content--contenttype-)|Erstellt eine Kommentarantwort für einen Kommentar.|
||[getCount()](/javascript/api/excel/excel.commentreplycollection#getcount--)|Ruft die Anzahl der Kommentarantworten in der Sammlung ab.|
||[getItem(commentReplyId: string)](/javascript/api/excel/excel.commentreplycollection#getitem-commentreplyid-)|Gibt eine Kommentarantwort zurück, die durch ihre ID angegeben ist.|
||[getItemAt(index: number)](/javascript/api/excel/excel.commentreplycollection#getitemat-index-)|Ruft eine Kommentarantwort basierend auf ihrer Position in der Sammlung ab.|
||[items](/javascript/api/excel/excel.commentreplycollection#items)|Ruft die geladenen untergeordneten Elemente in dieser Sammlung ab.|
|[PivotLayout](/javascript/api/excel/excel.pivotlayout)|[enableFieldList](/javascript/api/excel/excel.pivotlayout#enablefieldlist)|Gibt an, ob die Feldliste auf der Benutzeroberfläche angezeigt werden kann.|
|[PivotTableStyle](/javascript/api/excel/excel.pivottablestyle)|[delete()](/javascript/api/excel/excel.pivottablestyle#delete--)|Löscht die PivotTable-Formatvorlage.|
||[duplicate()](/javascript/api/excel/excel.pivottablestyle#duplicate--)|Erstellt ein Duplikat dieser PivotTable-Formatvorlage mit Kopien aller Formatvorlageelemente.|
||[name](/javascript/api/excel/excel.pivottablestyle#name)|Ruft den Namen der PivotTable-Formatvorlage ab.|
||[readOnly](/javascript/api/excel/excel.pivottablestyle#readonly)|Gibt an, ob `PivotTableStyle` dieses Objekt schreibgeschützt ist.|
|[PivotTableStyleCollection](/javascript/api/excel/excel.pivottablestylecollection)|[add(name: string, makeUniqueName?: boolean)](/javascript/api/excel/excel.pivottablestylecollection#add-name--makeuniquename-)|Erstellt ein `PivotTableStyle` Leeres mit dem angegebenen Namen.|
||[getCount()](/javascript/api/excel/excel.pivottablestylecollection#getcount--)|Ruft die Anzahl der PivotTable-Formatvorlagen in der Sammlung ab.|
||[getDefault()](/javascript/api/excel/excel.pivottablestylecollection#getdefault--)|Ruft die standardmäßige PivotTable-Formatvorlage für den Bereich des übergeordneten Objekts ab.|
||[getItem(name: string)](/javascript/api/excel/excel.pivottablestylecollection#getitem-name-)|Ruft einen `PivotTableStyle` nach Namen ab.|
||[GetItemOrNullObject(name: string)](/javascript/api/excel/excel.pivottablestylecollection#getitemornullobject-name-)|Ruft einen `PivotTableStyle` nach Namen ab.|
||[items](/javascript/api/excel/excel.pivottablestylecollection#items)|Ruft die geladenen untergeordneten Elemente in dieser Sammlung ab.|
||[setDefault(newDefaultStyle: PivotTableStyle \| string)](/javascript/api/excel/excel.pivottablestylecollection#setdefault-newdefaultstyle-)|Legt die standardmäßige PivotTable-Formatvorlage für die Verwendung im Bereich des übergeordneten Objekts fest.|
|[Range](/javascript/api/excel/excel.range)|[group(groupOption: Excel.GroupOption)](/javascript/api/excel/excel.range#group-groupoption-)|Gruppen von Spalten und Zeilen für eine Gliederung.|
||[hideGroupDetails(groupOption: Excel.GroupOption)](/javascript/api/excel/excel.range#hidegroupdetails-groupoption-)|Blendet die Details der Zeilen- oder Spaltengruppe aus.|
||[height](/javascript/api/excel/excel.range#height)|Gibt den Abstand in Punkt für 100 % Zoom vom oberen Rand des Bereichs zum unteren Rand des Bereichs zurück.|
||[left](/javascript/api/excel/excel.range#left)|Gibt den Abstand in Punkt für 100 % Zoom vom linken Rand des Arbeitsblatts zum linken Rand des Bereichs zurück.|
||[top](/javascript/api/excel/excel.range#top)|Gibt den Abstand in Punkt für 100 % Zoom vom oberen Rand des Arbeitsblatts zum oberen Rand des Bereichs zurück.|
||[width](/javascript/api/excel/excel.range#width)|Gibt den Abstand in Punkt für 100 % Zoom vom linken Rand des Bereichs zum rechten Rand des Bereichs zurück.|
||[showGroupDetails(groupOption: Excel.GroupOption)](/javascript/api/excel/excel.range#showgroupdetails-groupoption-)|Zeigt die Details der Zeilen- oder Spaltengruppe an.|
||[ungroup(groupOption: Excel.GroupOption)](/javascript/api/excel/excel.range#ungroup-groupoption-)|Aufheben der Gruppierung von Spalten und Zeilen für eine Gliederung.|
|[Shape](/javascript/api/excel/excel.shape)|[copyTo(destinationSheet?: Worksheet \| string)](/javascript/api/excel/excel.shape#copyto-destinationsheet-)|Kopiert und fügt ein Objekt `Shape` ein.|
||[placement](/javascript/api/excel/excel.shape#placement)|Stellt dar, wie das Objekt an die Zellen darunter angefügt ist.|
|[Datenschnitt](/javascript/api/excel/excel.slicer)|[caption](/javascript/api/excel/excel.slicer#caption)|Stellt die Beschriftung des Datenschnitts dar.|
||[clearFilters()](/javascript/api/excel/excel.slicer#clearfilters--)|Löscht alle Filter, die derzeit für den Datenschnitt verwendet werden.|
||[delete()](/javascript/api/excel/excel.slicer#delete--)|Löscht den Datenschnitt.|
||[getSelectedItems()](/javascript/api/excel/excel.slicer#getselecteditems--)|Gibt ein Array mit den Schlüsseln der ausgewählten Elemente zurück.|
||[height](/javascript/api/excel/excel.slicer#height)|Stellt die Höhe des Datenschnitts in typografischen Punkten dar.|
||[left](/javascript/api/excel/excel.slicer#left)|Stellt den Abstand in Punkt von der linken Seite des Datenschnitts zur linken Seite des Arbeitsblatts dar.|
||[name](/javascript/api/excel/excel.slicer#name)|Stellt den Namen des Datenschnitts dar.|
||[id](/javascript/api/excel/excel.slicer#id)|Stellt die eindeutige ID des Datenschnitts dar.|
||[isFilterCleared](/javascript/api/excel/excel.slicer#isfiltercleared)|Der Wert ist, wenn alle derzeit auf den Datenschnitt angewendeten `true` Filter geleert werden.|
||[slicerItems](/javascript/api/excel/excel.slicer#sliceritems)|Stellt die Auflistung von Datenschnittelementen dar, die Teil des Datenschnitts sind.|
||[worksheet](/javascript/api/excel/excel.slicer#worksheet)|Stellt das Arbeitsblatt dar, das den Datenschnitt enthält.|
||[selectItems(items?: string[])](/javascript/api/excel/excel.slicer#selectitems-items-)|Wählt Datenschnittelemente basierend auf ihren Schlüsseln aus.|
||[sortBy](/javascript/api/excel/excel.slicer#sortby)|Stellt die Sortierreihenfolge der Elemente im Datenschnitt dar.|
||[style](/javascript/api/excel/excel.slicer#style)|Konstantenwert, der die Datenschnittart darstellt.|
||[top](/javascript/api/excel/excel.slicer#top)|Stellt den Abstand in Punkt von der Oberkante des Datenschnitts zur Oberkante des Arbeitsblatts dar.|
||[width](/javascript/api/excel/excel.slicer#width)|Die Breite des Datenschnitts in Punkten.|
|[SlicerCollection](/javascript/api/excel/excel.slicercollection)|[add(slicerSource: string \| PivotTable \| Table, sourceField: string \| PivotField \| number \| TableColumn, slicerDestination?: string \| Worksheet)](/javascript/api/excel/excel.slicercollection#add-slicersource--sourcefield--slicerdestination-)|Fügt der Arbeitsmappe einen neuen Datenschnitt hinzu.|
||[getCount()](/javascript/api/excel/excel.slicercollection#getcount--)|Gibt die Anzahl der Datenschnitte in der Sammlung zurück.|
||[getItem(key: string)](/javascript/api/excel/excel.slicercollection#getitem-key-)|Ruft ein Datenschnittobjekt mit seinem Namen oder seiner ID ab.|
||[getItemAt(index: number)](/javascript/api/excel/excel.slicercollection#getitemat-index-)|Ruft einen Datenschnitt anhand seiner Position in der Auflistung ab.|
||[getItemOrNullObject(key: string)](/javascript/api/excel/excel.slicercollection#getitemornullobject-key-)|Ruft einen Datenschnitt mit seinem Namen oder seiner ID ab.|
||[items](/javascript/api/excel/excel.slicercollection#items)|Ruft die geladenen untergeordneten Elemente in dieser Sammlung ab.|
|[SlicerItem](/javascript/api/excel/excel.sliceritem)|[isSelected](/javascript/api/excel/excel.sliceritem#isselected)|Der Wert `true` ist, wenn das Datenschnittelement ausgewählt ist.|
||[hasData](/javascript/api/excel/excel.sliceritem#hasdata)|Der Wert `true` ist, wenn das Datenschnittelement Daten enthält.|
||[key](/javascript/api/excel/excel.sliceritem#key)|Stellt den eindeutigen Wert dar, der für das Datenschnittelement steht.|
||[name](/javascript/api/excel/excel.sliceritem#name)|Represents the title displayed in the Excel UI.|
|[SlicerItemCollection](/javascript/api/excel/excel.sliceritemcollection)|[getCount()](/javascript/api/excel/excel.sliceritemcollection#getcount--)|Gibt die Anzahl der Datenschnittelemente im Datenschnitt zurück.|
||[getItem(key: string)](/javascript/api/excel/excel.sliceritemcollection#getitem-key-)|Ruft ein Datenschnittelement-Objekt anhand seines Schlüssels oder Namens ab.|
||[getItemAt(index: number)](/javascript/api/excel/excel.sliceritemcollection#getitemat-index-)|Ruft ein Datenschnittelement anhand seiner Position in der Sammlung ab.|
||[getItemOrNullObject(key: string)](/javascript/api/excel/excel.sliceritemcollection#getitemornullobject-key-)|Ruft ein Datenschnittelement anhand seines Schlüssels oder Namens ab.|
||[items](/javascript/api/excel/excel.sliceritemcollection#items)|Ruft die geladenen untergeordneten Elemente in dieser Sammlung ab.|
|[SlicerStyle](/javascript/api/excel/excel.slicerstyle)|[delete()](/javascript/api/excel/excel.slicerstyle#delete--)|Löscht die Datenschnittart.|
||[duplicate()](/javascript/api/excel/excel.slicerstyle#duplicate--)|Erstellt ein Duplikat dieser Datenschnittformatvorlage mit Kopien aller Formatvorlageelemente.|
||[name](/javascript/api/excel/excel.slicerstyle#name)|Ruft den Namen des Datenschnittformats ab.|
||[readOnly](/javascript/api/excel/excel.slicerstyle#readonly)|Gibt an, ob `SlicerStyle` dieses Objekt schreibgeschützt ist.|
|[SlicerStyleCollection](/javascript/api/excel/excel.slicerstylecollection)|[add(name: string, makeUniqueName?: boolean)](/javascript/api/excel/excel.slicerstylecollection#add-name--makeuniquename-)|Erstellt ein leeres Datenschnittformat mit dem angegebenen Namen.|
||[getCount()](/javascript/api/excel/excel.slicerstylecollection#getcount--)|Ruft die Anzahl der Datenschnitt-Formatvorlagen in der Sammlung ab.|
||[getDefault()](/javascript/api/excel/excel.slicerstylecollection#getdefault--)|Ruft den Standardwert `SlicerStyle` für den Bereich des übergeordneten Objekts ab.|
||[getItem(name: string)](/javascript/api/excel/excel.slicerstylecollection#getitem-name-)|Ruft einen `SlicerStyle` nach Namen ab.|
||[GetItemOrNullObject(name: string)](/javascript/api/excel/excel.slicerstylecollection#getitemornullobject-name-)|Ruft einen `SlicerStyle` nach Namen ab.|
||[items](/javascript/api/excel/excel.slicerstylecollection#items)|Ruft die geladenen untergeordneten Elemente in dieser Sammlung ab.|
||[setDefault(newDefaultStyle: SlicerStyle \| string)](/javascript/api/excel/excel.slicerstylecollection#setdefault-newdefaultstyle-)|Legt die Standardschnittart für die Verwendung im Bereich des übergeordneten Objekts fest.|
|[TableStyle](/javascript/api/excel/excel.tablestyle)|[delete()](/javascript/api/excel/excel.tablestyle#delete--)|Löscht die Tabellenformatvorlage.|
||[duplicate()](/javascript/api/excel/excel.tablestyle#duplicate--)|Erstellt ein Duplikat dieser Tabellenformatvorlage mit Kopien aller Formatvorlageelemente.|
||[name](/javascript/api/excel/excel.tablestyle#name)|Ruft den Namen der Tabellenformatvorlage ab.|
||[readOnly](/javascript/api/excel/excel.tablestyle#readonly)|Gibt an, ob `TableStyle` dieses Objekt schreibgeschützt ist.|
|[TableStyleCollection](/javascript/api/excel/excel.tablestylecollection)|[add(name: string, makeUniqueName?: boolean)](/javascript/api/excel/excel.tablestylecollection#add-name--makeuniquename-)|Erstellt ein `TableStyle` Leeres mit dem angegebenen Namen.|
||[getCount()](/javascript/api/excel/excel.tablestylecollection#getcount--)|Ruft die Anzahl der Tabellenformatvorlagen in der Sammlung ab.|
||[getDefault()](/javascript/api/excel/excel.tablestylecollection#getdefault--)|Ruft die Standardformatvorlage für den Bereich des übergeordneten Objekts ab.|
||[getItem(name: string)](/javascript/api/excel/excel.tablestylecollection#getitem-name-)|Ruft einen `TableStyle` nach Namen ab.|
||[GetItemOrNullObject(name: string)](/javascript/api/excel/excel.tablestylecollection#getitemornullobject-name-)|Ruft einen `TableStyle` nach Namen ab.|
||[items](/javascript/api/excel/excel.tablestylecollection#items)|Ruft die geladenen untergeordneten Elemente in dieser Sammlung ab.|
||[setDefault(newDefaultStyle: TableStyle \| string)](/javascript/api/excel/excel.tablestylecollection#setdefault-newdefaultstyle-)|Legt die Standardtabelle für die Verwendung im Bereich des übergeordneten Objekts fest.|
|[TimelineStyle](/javascript/api/excel/excel.timelinestyle)|[delete()](/javascript/api/excel/excel.timelinestyle#delete--)|Löscht die Tabellenformatvorlage.|
||[duplicate()](/javascript/api/excel/excel.timelinestyle#duplicate--)|Erstellt ein Duplikat dieser Zeitachsenformatvorlage mit Kopien aller Formatvorlageelemente.|
||[name](/javascript/api/excel/excel.timelinestyle#name)|Ruft den Namen des Zeitachsenformats ab.|
||[readOnly](/javascript/api/excel/excel.timelinestyle#readonly)|Gibt an, ob `TimelineStyle` dieses Objekt schreibgeschützt ist.|
|[TimelineStyleCollection](/javascript/api/excel/excel.timelinestylecollection)|[add(name: string, makeUniqueName?: boolean)](/javascript/api/excel/excel.timelinestylecollection#add-name--makeuniquename-)|Erstellt ein `TimelineStyle` Leeres mit dem angegebenen Namen.|
||[getCount()](/javascript/api/excel/excel.timelinestylecollection#getcount--)|Ruft die Anzahl der Zeitachsen-Formatvorlagen in der Sammlung ab.|
||[getDefault()](/javascript/api/excel/excel.timelinestylecollection#getdefault--)|Ruft die Standardzeitachsenformatvorlage für den Bereich des übergeordneten Objekts ab.|
||[getItem(name: string)](/javascript/api/excel/excel.timelinestylecollection#getitem-name-)|Ruft einen `TimelineStyle` nach Namen ab.|
||[GetItemOrNullObject(name: string)](/javascript/api/excel/excel.timelinestylecollection#getitemornullobject-name-)|Ruft einen `TimelineStyle` nach Namen ab.|
||[items](/javascript/api/excel/excel.timelinestylecollection#items)|Ruft die geladenen untergeordneten Elemente in dieser Sammlung ab.|
||[setDefault(newDefaultStyle: TimelineStyle \| string)](/javascript/api/excel/excel.timelinestylecollection#setdefault-newdefaultstyle-)|Legt die Standardzeitachsenart für die Verwendung im Bereich des übergeordneten Objekts fest.|
|[Workbook](/javascript/api/excel/excel.workbook)|[getActiveSlicer()](/javascript/api/excel/excel.workbook#getactiveslicer--)|Ruft den derzeit aktiven Datenschnitt in der Arbeitsmappe ab.|
||[getActiveSlicerOrNullObject()](/javascript/api/excel/excel.workbook#getactiveslicerornullobject--)|Ruft den derzeit aktiven Datenschnitt in der Arbeitsmappe ab.|
||[comments](/javascript/api/excel/excel.workbook#comments)|Stellt eine Sammlung von Kommentaren dar, die der Arbeitsmappe zugeordnet sind.|
||[pivotTableStyles](/javascript/api/excel/excel.workbook#pivottablestyles)|Stellt eine Sammlung der mit der Arbeitsmappe verknüpften PivotTableStyles dar.|
||[slicerStyles](/javascript/api/excel/excel.workbook#slicerstyles)|Stellt eine Sammlung der mit der Arbeitsmappe verknüpften SlicerStyles dar.|
||[slicers](/javascript/api/excel/excel.workbook#slicers)|Stellt eine Sammlung von Datenschnitten dar, die der Arbeitsmappe zugeordnet sind.|
||[tableStyles](/javascript/api/excel/excel.workbook#tablestyles)|Stellt eine Auflistung der mit der Arbeitsmappe verknüpften TableStyles dar.|
||[timelineStyles](/javascript/api/excel/excel.workbook#timelinestyles)|Stellt eine Auflistung der mit der Arbeitsmappe verknüpften TimelineStyles dar.|
|[Arbeitsblatt](/javascript/api/excel/excel.worksheet)|[comments](/javascript/api/excel/excel.worksheet#comments)|Gibt eine Sammlung aller Kommentarobjekte auf dem Arbeitsblatt zurück.|
||[onColumnSorted](/javascript/api/excel/excel.worksheet#oncolumnsorted)|Tritt auf, wenn eine oder mehrere Spalten sortiert wurden.|
||[onRowSorted](/javascript/api/excel/excel.worksheet#onrowsorted)|Tritt auf, wenn eine oder mehrere Zeilen sortiert wurden.|
||[onSingleClicked](/javascript/api/excel/excel.worksheet#onsingleclicked)|Tritt auf, wenn im Arbeitsblatt eine mit der linken Maustaste geklickte/angetippte Aktion stattfindet.|
||[slicers](/javascript/api/excel/excel.worksheet#slicers)|Gibt eine Auflistung von Datenschnitten zurück, die Teil des Arbeitsblatts sind.|
||[showOutlineLevels(rowLevels: number, columnLevels: number)](/javascript/api/excel/excel.worksheet#showoutlinelevels-rowlevels--columnlevels-)|Zeigt Zeilen- oder Spaltengruppen nach ihren Gliederungsebenen an.|
|[WorksheetCollection](/javascript/api/excel/excel.worksheetcollection)|[onColumnSorted](/javascript/api/excel/excel.worksheetcollection#oncolumnsorted)|Tritt auf, wenn eine oder mehrere Spalten sortiert wurden.|
||[onRowSorted](/javascript/api/excel/excel.worksheetcollection#onrowsorted)|Tritt auf, wenn eine oder mehrere Zeilen sortiert wurden.|
||[onSingleClicked](/javascript/api/excel/excel.worksheetcollection#onsingleclicked)|Tritt auf, wenn in der Arbeitsblattsammlung mit der linken Maustaste geklickt/getippt wird.|
|[WorksheetColumnSortedEventArgs](/javascript/api/excel/excel.worksheetcolumnsortedeventargs)|[address](/javascript/api/excel/excel.worksheetcolumnsortedeventargs#address)|Ruft die Bereichsadresse ab, die den sortierten Bereich auf einem bestimmten Arbeitsblatt darstellt.|
||[source](/javascript/api/excel/excel.worksheetcolumnsortedeventargs#source)|Ruft die Quelle des Ereignisses ab.|
||[type](/javascript/api/excel/excel.worksheetcolumnsortedeventargs#type)|Ruft den Typ des Ereignisses ab.|
||[worksheetId](/javascript/api/excel/excel.worksheetcolumnsortedeventargs#worksheetid)|Ruft die ID des Arbeitsblatts ab, in dem die Sortierung passiert ist.|
|[WorksheetRowSortedEventArgs](/javascript/api/excel/excel.worksheetrowsortedeventargs)|[address](/javascript/api/excel/excel.worksheetrowsortedeventargs#address)|Ruft die Bereichsadresse ab, die den sortierten Bereich auf einem bestimmten Arbeitsblatt darstellt.|
||[source](/javascript/api/excel/excel.worksheetrowsortedeventargs#source)|Ruft die Quelle des Ereignisses ab.|
||[type](/javascript/api/excel/excel.worksheetrowsortedeventargs#type)|Ruft den Typ des Ereignisses ab.|
||[worksheetId](/javascript/api/excel/excel.worksheetrowsortedeventargs#worksheetid)|Ruft die ID des Arbeitsblatts ab, in dem die Sortierung passiert ist.|
|[WorksheetSingleClickedEventArgs](/javascript/api/excel/excel.worksheetsingleclickedeventargs)|[address](/javascript/api/excel/excel.worksheetsingleclickedeventargs#address)|Ruft die Adresse ab, welche die Zelle repräsentiert, die für ein bestimmtes Arbeitsblatt mit der linken Maustaste angeklickt/getippt wurde.|
||[offsetX](/javascript/api/excel/excel.worksheetsingleclickedeventargs#offsetx)|Der Abstand (in Punkt) vom linksgeklickten/angetippten Punkt zum linken Gitternetzrand (oder rechts für Rechts-nach-links-Sprachen) der Zelle mit links geklickter/getippter Zelle.|
||[offsetY](/javascript/api/excel/excel.worksheetsingleclickedeventargs#offsety)|Der Abstand, in Punkten, vom linksgeklickten/getippten Punkt bis zum oberen Rand der Gitterlinie der linksgeklickten/getippten Zelle.|
||[Typ](/javascript/api/excel/excel.worksheetsingleclickedeventargs#type)|Ruft den Typ des Ereignisses ab.|
||[worksheetId](/javascript/api/excel/excel.worksheetsingleclickedeventargs#worksheetid)|Ruft die ID des Arbeitsblatts ab, in dem auf die Zelle mit der linken Maustaste geklickt/getippt wurde.|
