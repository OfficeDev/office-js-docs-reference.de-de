| Klasse | Felder | Beschreibung |
|:---|:---|:---|
|[Binding](/javascript/api/excel/excel.binding)|[delete()](/javascript/api/excel/excel.binding#delete--)|Löscht die Bindung.|
|[BindingCollection](/javascript/api/excel/excel.bindingcollection)|[add(range: Range \| string, bindingType: Excel.BindingType, id: string)](/javascript/api/excel/excel.bindingcollection#add-range--bindingtype--id-)|Fügt eine neue Bindung an einen bestimmten Bereich hinzu.|
||[addFromNamedItem(name: string, bindingType: Excel.BindingType, id: string)](/javascript/api/excel/excel.bindingcollection#addfromnameditem-name--bindingtype--id-)|Fügt eine neue Bindung auf Grundlage eines benannten Elements in der Arbeitsmappe hinzu.|
||[addFromSelection(bindingType: Excel.BindingType, id: string)](/javascript/api/excel/excel.bindingcollection#addfromselection-bindingtype--id-)|Fügt eine neue Bindung basierend auf der aktuellen Auswahl hinzu.|
|[PivotTable](/javascript/api/excel/excel.pivottable)|[name](/javascript/api/excel/excel.pivottable#name)|Der Name der PivotTable.|
||[Arbeitsblatt](/javascript/api/excel/excel.pivottable#worksheet)|Das Arbeitsblatt, das die aktuelle PivotTable enthält.|
||[refresh()](/javascript/api/excel/excel.pivottable#refresh--)|Aktualisiert die PivotTable.|
|[PivotTableCollection](/javascript/api/excel/excel.pivottablecollection)|[getItem(name: string)](/javascript/api/excel/excel.pivottablecollection#getitem-name-)|Ruft eine PivotTable anhand des Namens ab.|
||[items](/javascript/api/excel/excel.pivottablecollection#items)|Ruft die geladenen untergeordneten Elemente in dieser Sammlung ab.|
||[refreshAll()](/javascript/api/excel/excel.pivottablecollection#refreshall--)|Aktualisiert alle PivotTables in der Sammlung.|
|[Range](/javascript/api/excel/excel.range)|[getVisibleView()](/javascript/api/excel/excel.range#getvisibleview--)|Stellt die sichtbaren Zeilen des aktuellen Bereichs dar.|
|[RangeView](/javascript/api/excel/excel.rangeview)|[formulas](/javascript/api/excel/excel.rangeview#formulas)|Stellt die Formel in der A1-Schreibweise dar.|
||[formulasLocal](/javascript/api/excel/excel.rangeview#formulaslocal)|Stellt die Formel in der A1-Schreibweise, Sprache des Benutzers und im Gebietsschema der Zahlenformatierung dar.|
||[formulasR1C1](/javascript/api/excel/excel.rangeview#formulasr1c1)|Stellt die Formel in der R1C1-Schreibweise dar.|
||[getRange()](/javascript/api/excel/excel.rangeview#getrange--)|Ruft den übergeordneten Bereich ab, der dem aktuellen zugeordnet `RangeView` ist.|
||[numberFormat](/javascript/api/excel/excel.rangeview#numberformat)|Stellt den Excel-Zahlenformatcode für die angegebene Zelle dar.|
||[cellAddresses](/javascript/api/excel/excel.rangeview#celladdresses)|Stellt die Zellenadressen des `RangeView` dar.|
||[columnCount](/javascript/api/excel/excel.rangeview#columncount)|Die Anzahl der sichtbaren Spalten.|
||[Index](/javascript/api/excel/excel.rangeview#index)|Gibt einen Wert zurück, der den Index des `RangeView` darstellt.|
||[rowCount](/javascript/api/excel/excel.rangeview#rowcount)|Die Anzahl der sichtbaren Zeilen.|
||[rows](/javascript/api/excel/excel.rangeview#rows)|Stellt eine Sammlung der mit dem Bereich verknüpften Bereichsansichten dar.|
||[text](/javascript/api/excel/excel.rangeview#text)|Textwerte des angegebenen Bereichs.|
||[valueTypes](/javascript/api/excel/excel.rangeview#valuetypes)|Stellt den Datentyp in jeder Zelle dar.|
||[values](/javascript/api/excel/excel.rangeview#values)|Stellt die Rohwerte der angegebenen Bereichsansicht dar.|
|[RangeViewCollection](/javascript/api/excel/excel.rangeviewcollection)|[getItemAt(index: number)](/javascript/api/excel/excel.rangeviewcollection#getitemat-index-)|Ruft eine `RangeView` Zeile über ihren Index ab.|
||[items](/javascript/api/excel/excel.rangeviewcollection#items)|Ruft die geladenen untergeordneten Elemente in dieser Sammlung ab.|
|[Table](/javascript/api/excel/excel.table)|[highlightFirstColumn](/javascript/api/excel/excel.table#highlightfirstcolumn)|Gibt an, ob die erste Spalte spezielle Formatierungen enthält.|
||[highlightLastColumn](/javascript/api/excel/excel.table#highlightlastcolumn)|Gibt an, ob die letzte Spalte spezielle Formatierungen enthält.|
||[showBandedColumns](/javascript/api/excel/excel.table#showbandedcolumns)|Gibt an, ob die Spalten eine gebänderte Formatierung anzeigen, bei der ungerade Spalten anders hervorgehoben werden als gleichmäßige Spalten, um das Lesen der Tabelle zu vereinfachen.|
||[showBandedRows](/javascript/api/excel/excel.table#showbandedrows)|Gibt an, ob in den Zeilen eine gebänderte Formatierung angezeigt wird, bei der ungerade Zeilen anders hervorgehoben werden als gleichmäßige Zeilen, um das Lesen der Tabelle zu vereinfachen.|
||[showFilterButton](/javascript/api/excel/excel.table#showfilterbutton)|Gibt an, ob die Filterschaltflächen am oberen Rand jeder Spaltenüberschrift angezeigt werden.|
|[Workbook](/javascript/api/excel/excel.workbook)|[pivotTables](/javascript/api/excel/excel.workbook#pivottables)|Stellt eine Auflistung der mit der Arbeitsmappe verknüpften PivotTables dar.|
|[Arbeitsblatt](/javascript/api/excel/excel.worksheet)|[pivotTables](/javascript/api/excel/excel.worksheet#pivottables)|Die Sammlung von PivotTables, die Teil des Arbeitsblatts sind.|
