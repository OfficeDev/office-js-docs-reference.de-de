# <a name="excel-javascript-api-overview"></a>Excel-JavaScript-API (Übersicht)

Sie können die JavaScript-APIs verwenden, um Add-Ins für Excel 2016 zu erstellen. Die folgende Liste enthält die Excel-Objekte auf hoher Ebene, die in der API verfügbar sind. Jeder Seite-Verknüpfung enthält eine Beschreibung der Eigenschaften, Ereignisse und Methoden, die für das Objekt verfügbar sind. Erkunden Sie die Links aus dem Menü, um mehr zu erfahren.

Einige der zentralen Excel-Objekte sind zur einfacheren Handhabung unten aufgeführt: 

- [Arbeitsmappe](/javascript/api/excel/excel.workbook): Das Objekt auf oberster Ebene , das dazugehörige Arbeitsmappenobjekte wie z. B. Arbeitsblätter, Tabellen, Bereiche usw. enthält. Es kann auch zum Auflisten zugehöriger Verweise verwendet werden.

- [Worksheet](/javascript/api/excel/excel.worksheet): Stellt ein Arbeitsblatt in einer Arbeitsmappe dar. 
    - [WorksheetCollection](/javascript/api/excel/excel.worksheetcollection): Eine Sammlung von **Worksheet**-Objekten in einer Arbeitsmappe.

- [Range](/javascript/api/excel/excel.range): Stellt eine Zelle, Zeile, Spalte oder Auswahl von Zellen dar, die mindestens einen zusammenhängenden Zellenblock enthalten.

- [Table](/javascript/api/excel/excel.table): Stellt eine Aufstellung organisierter Zellen zur einfachen Datenverwaltung dar.
    - [TableCollection](/javascript/api/excel/excel.tablecollection): Eine Sammlung von Tabellen einer Arbeitsmappe oder eines Arbeitsblatts.
    - [TableColumnCollection](/javascript/api/excel/excel.tablecolumncollection): Eine Sammlung aller Spalten in einer Tabelle.
    - [TableRowCollection](/javascript/api/excel/excel.tablerowcollection): Eine Sammlung aller Zeilen in einer Tabelle.

- [Chart](/javascript/api/excel/excel.chart): Stelle ein Diagrammobjekt in einem Arbeitsblatt dar, das die zugrunde liegenden Daten visuell darstellt.
    - [ChartCollection](/javascript/api/excel/excel.chartcollection): Eine Sammlung von Diagrammen in einem Arbeitsblatt.

- [TableSort](/javascript/api/excel/excel.tablesort): Ein Objekt für das Verwalten von Sortiervorgängen für **Table**-Objekte.

- [rangeSort](/javascript/api/excel/excel.rangesort): Ein Objekt für das Verwalten von Sortiervorgängen für **Range**-Objekte.

- [Filter](/javascript/api/excel/excel.filter): Ein Objekt, das die Filter einer Tabellenspalte verwaltet.

- [WorksheetProtection](/javascript/api/excel/excel.worksheetprotection): Stellt den Schutz eines **Worksheet**-Objekts dar.

- [NamedItem](/javascript/api/excel/excel.nameditem): Stellt einen definierten Namen für einen Zellbereich oder einen Wert dar. 
    - [NamedItemCollection](/javascript/api/excel/excel.nameditemcollection): Eine Sammlung von **NamedItem**-Objekten in einer Arbeitsmappe.

- [Binding](/javascript/api/excel/excel.binding): Eine abstrakte Klasse, die eine Bindung an einen Abschnitt der Arbeitsmappe darstellt.
    - [BindingCollection](/javascript/api/excel/excel.bindingcollection): Eine Sammlung von **Binding**-Objekten in einer Arbeitsmappe.

## <a name="excel-javascript-api-open-specifications"></a>Offene Spezifikationen JavaScript-API für Excel

Wir entwerfen und Entwickeln von neuen APIs für Excel-add-ins, können wir ihnen für Ihr Feedback auf unserer Seite [Open API-Spezifikationen](../openspec.md) zur Verfügung. Erfahren Sie, was neue Funktionen sind in der Pipeline für die Excel-JavaScript-APIs und Ihre mitteilen, unsere Entwurfsspezifikationen.

## <a name="excel-javascript-api-reference"></a>JavaScript-API-Referenz für Excel

Ausführliche Informationen zu Excel-JavaScript-API finden Sie unter der [JavaScript-API für Excel-Referenzdokumentation](/javascript/api/excel).

## <a name="see-also"></a>Siehe auch

- [Übersicht über Excel-Add-Ins](https://docs.microsoft.com/office/dev/add-ins/excel/excel-add-ins-overview)
- [Office-Add-Ins-Plattformübersicht](https://docs.microsoft.com/office/dev/add-ins/overview/office-add-ins)
- [Excel-add-in-Beispiele in der GitHub](https://github.com/OfficeDev?utf8=%E2%9C%93&q=Excel)
