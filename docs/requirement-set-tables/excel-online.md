| Klasse | Felder | Beschreibung |
|:---|:---|:---|
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
||[items](/javascript/api/excel/excel.namedsheetviewcollection#items)|Ruft die geladenen untergeordneten Elemente in dieser Sammlung ab.|
|[Range](/javascript/api/excel/excel.range)|[getExtendedRange(direction: Excel.KeyboardDirection, activeCell?: Range \| string)](/javascript/api/excel/excel.range#getextendedrange-direction--activecell-)|Gibt ein Range-Objekt zurück, das den aktuellen Bereich und bis zum Rand des Bereichs enthält, basierend auf der angegebenen Richtung.|
||[getMergedAreas()](/javascript/api/excel/excel.range#getmergedareas--)|Gibt ein `RangeAreas` Objekt zurück, das die zusammengeführten Bereiche in diesem Bereich darstellt.|
||[getRangeEdge(direction: Excel.KeyboardDirection, activeCell?: Range \| string)](/javascript/api/excel/excel.range#getrangeedge-direction--activecell-)|Gibt ein Range-Objekt zurück, das die Edgezelle des Datenbereichs ist, die der angegebenen Richtung entspricht.|
|[Table](/javascript/api/excel/excel.table)|[resize(newRange: Range \| string)](/javascript/api/excel/excel.table#resize-newrange-)|Ändern Sie die Größe der Tabelle auf den neuen Bereich.|
|[Arbeitsblatt](/javascript/api/excel/excel.worksheet)|[namedSheetViews](/javascript/api/excel/excel.worksheet#namedsheetviews)|Gibt eine Auflistung von Blattansichten zurück, die im Arbeitsblatt vorhanden sind.|
