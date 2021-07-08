| Klasse | Felder | Beschreibung |
|:---|:---|:---|
|[FormulaChangedEventDetail](/javascript/api/excel/excel.formulachangedeventdetail)|[cellAddress](/javascript/api/excel/excel.formulachangedeventdetail#celladdress)|Die Adresse der Zelle, die die geänderte Formel enthält.|
||[previousFormula](/javascript/api/excel/excel.formulachangedeventdetail#previousformula)|Stellt die vorherige Formel dar, bevor sie geändert wurde.|
|[InsertWorksheetOptions](/javascript/api/excel/excel.insertworksheetoptions)|[positionType](/javascript/api/excel/excel.insertworksheetoptions#positiontype)|Die Position zum Einfügen der neuen Arbeitsblätter in der aktuellen Arbeitsmappe.|
||[Relativeto](/javascript/api/excel/excel.insertworksheetoptions#relativeto)|Das Arbeitsblatt in der aktuellen Arbeitsmappe, auf das für den Parameter verwiesen `WorksheetPositionType` wird.|
||[sheetNamesToInsert](/javascript/api/excel/excel.insertworksheetoptions#sheetnamestoinsert)|Die Namen der einzelnen einzufügenden Arbeitsblätter.|
|[PivotLayout](/javascript/api/excel/excel.pivotlayout)|[altTextDescription](/javascript/api/excel/excel.pivotlayout#alttextdescription)|Die Alternativtextbeschreibung der PivotTable.|
||[altTextTitle](/javascript/api/excel/excel.pivotlayout#alttexttitle)|Der Alttexttitel der PivotTable.|
||[displayBlankLineAfterEachItem(display: boolean)](/javascript/api/excel/excel.pivotlayout#displayblanklineaftereachitem-display-)|Legt fest, ob nach jedem Element eine leere Zeile angezeigt werden soll.|
||[emptyCellText](/javascript/api/excel/excel.pivotlayout#emptycelltext)|Der Text, der automatisch in eine beliebige leere Zelle in der PivotTable gefüllt wird, wenn `fillEmptyCells == true` .|
||[fillEmptyCells](/javascript/api/excel/excel.pivotlayout#fillemptycells)|Gibt an, ob leere Zellen in der PivotTable mit der aufgefüllt werden `emptyCellText` sollen.|
||[repeatAllItemLabels(repeatLabels: boolean)](/javascript/api/excel/excel.pivotlayout#repeatallitemlabels-repeatlabels-)|Legt die Einstellung "Alle Elementbeschriftungen wiederholen" für alle Felder in der PivotTable fest.|
||[showFieldHeaders](/javascript/api/excel/excel.pivotlayout#showfieldheaders)|Gibt an, ob die PivotTable Feldüberschriften (Feldbeschriftungen und Filter-Dropdowns) anzeigt.|
|[PivotTable](/javascript/api/excel/excel.pivottable)|[refreshOnOpen](/javascript/api/excel/excel.pivottable#refreshonopen)|Gibt an, ob die PivotTable beim Öffnen der Arbeitsmappe aktualisiert wird.|
|[Range](/javascript/api/excel/excel.range)|[getDirectDependents()](/javascript/api/excel/excel.range#getdirectdependents--)|Gibt ein `WorkbookRangeAreas` Objekt zurück, das den Bereich darstellt, der alle direkten Nachfolger einer Zelle im selben Arbeitsblatt oder in mehreren Arbeitsblättern enthält.|
||[getExtendedRange(direction: Excel. KeyboardDirection, activeCell?: \| Bereichszeichenfolge)](/javascript/api/excel/excel.range#getextendedrange-direction--activecell-)|Gibt basierend auf der angegebenen Richtung ein Bereichsobjekt zurück, das den aktuellen Bereich und bis zum Rand des Bereichs enthält.|
||[getMergedAreasOrNullObject()](/javascript/api/excel/excel.range#getmergedareasornullobject--)|Gibt ein RangeAreas -Objekt zurück, das die zusammengeführten Bereiche in diesem Bereich darstellt.|
||[getRangeEdge(direction: Excel. KeyboardDirection, activeCell?: \| Bereichszeichenfolge)](/javascript/api/excel/excel.range#getrangeedge-direction--activecell-)|Gibt ein Bereichsobjekt zurück, bei dem es sich um die Edgezelle des Datenbereichs handelt, die der angegebenen Richtung entspricht.|
|[Table](/javascript/api/excel/excel.table)|[resize(newRange: Range \| string)](/javascript/api/excel/excel.table#resize-newrange-)|Ändern Sie die Größe der Tabelle in den neuen Bereich.|
|[Workbook](/javascript/api/excel/excel.workbook)|[insertWorksheetsFromBase64(base64File: string, options?: Excel. InsertWorksheetOptions)](/javascript/api/excel/excel.workbook#insertworksheetsfrombase64-base64file--options-)|Fügt die angegebenen Arbeitsblätter aus einer Quellarbeitsmappe in die aktuelle Arbeitsmappe ein.|
||[onActivated](/javascript/api/excel/excel.workbook#onactivated)|Tritt auf, wenn die Arbeitsmappe aktiviert wird.|
|[WorkbookActivatedEventArgs](/javascript/api/excel/excel.workbookactivatedeventargs)|[type](/javascript/api/excel/excel.workbookactivatedeventargs#type)|Ruft den Typ des Ereignisses ab.|
|[Arbeitsblatt](/javascript/api/excel/excel.worksheet)|[onFormulaChanged](/javascript/api/excel/excel.worksheet#onformulachanged)|Tritt auf, wenn eine oder mehrere Formeln in diesem Arbeitsblatt geändert werden.|
|[WorksheetCollection](/javascript/api/excel/excel.worksheetcollection)|[onFormulaChanged](/javascript/api/excel/excel.worksheetcollection#onformulachanged)|Tritt auf, wenn eine oder mehrere Formeln in einem Arbeitsblatt dieser Auflistung geändert werden.|
|[WorksheetFormulaChangedEventArgs](/javascript/api/excel/excel.worksheetformulachangedeventargs)|[formulaDetails](/javascript/api/excel/excel.worksheetformulachangedeventargs#formuladetails)|Ruft ein Array von `FormulaChangedEventDetail` Objekten ab, die die Details zu allen geänderten Formeln enthalten.|
||[source](/javascript/api/excel/excel.worksheetformulachangedeventargs#source)|Die Quelle des Ereignisses.|
||[type](/javascript/api/excel/excel.worksheetformulachangedeventargs#type)|Ruft den Typ des Ereignisses ab.|
||[worksheetId](/javascript/api/excel/excel.worksheetformulachangedeventargs#worksheetid)|Ruft die ID des Arbeitsblatts ab, in dem die Formel geändert wurde.|
