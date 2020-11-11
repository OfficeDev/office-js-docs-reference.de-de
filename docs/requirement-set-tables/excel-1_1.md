| Klasse | Felder | Beschreibung |
|:---|:---|:---|
|[Anwendung](/javascript/api/excel/excel.application)|[Calculate (CalculationType: Excel. CalculationType)](/javascript/api/excel/excel.application#calculate-calculationtype-)|Alle in Excel geöffnete Arbeitsmappen werden neu berechnet.|
||[calculationMode](/javascript/api/excel/excel.application#calculationmode)|Gibt den in der Arbeitsmappe verwendeten Berechnungsmodus gemäß der Definition durch die Konstanten in Excel. CalculationMode zurück.|
|[Binding](/javascript/api/excel/excel.binding)|[getRange()](/javascript/api/excel/excel.binding#getrange--)|Gibt den durch die Bindung dargestellten Bereich zurück.|
||[getTable()](/javascript/api/excel/excel.binding#gettable--)|Gibt die durch die Bindung dargestellte Tabelle zurück.|
||[getText()](/javascript/api/excel/excel.binding#gettext--)|Gibt den durch die Bindung dargestellten Text zurück.|
||[id](/javascript/api/excel/excel.binding#id)|Stellt die Bindungs-ID dar.|
||[Typ](/javascript/api/excel/excel.binding#type)|Gibt den Typ der Bindung an.|
|[BindingCollection](/javascript/api/excel/excel.bindingcollection)|[getItem(id: string)](/javascript/api/excel/excel.bindingcollection#getitem-id-)|Ruft ein Binding-Objekt anhand seiner ID ab.|
||[getItemAt(index: number)](/javascript/api/excel/excel.bindingcollection#getitemat-index-)|Ruft ein Binding-Objekt anhand seiner Position im Elementarray ab.|
||[count](/javascript/api/excel/excel.bindingcollection#count)|Gibt die Anzahl der Bindungen in der Sammlung zurück.|
||[items](/javascript/api/excel/excel.bindingcollection#items)|Ruft die geladenen untergeordneten Elemente in dieser Sammlung ab.|
|[Chart](/javascript/api/excel/excel.chart)|[delete()](/javascript/api/excel/excel.chart#delete--)|Löscht das Diagrammobjekt.|
||[height](/javascript/api/excel/excel.chart#height)|Gibt die Höhe des Chart-Objekts in Punkten an.|
||[left](/javascript/api/excel/excel.chart#left)|Der Abstand von der linken Seite des Diagramms zu dem Ursprung des Arbeitsblatts (in Punkten).|
||[name](/javascript/api/excel/excel.chart#name)|Gibt den Namen eines Chart-Objekts an.|
||[aXes](/javascript/api/excel/excel.chart#axes)|Die Achsen des Diagramms.|
||[dataLabels](/javascript/api/excel/excel.chart#datalabels)|Stellt die Datenbeschriftungen im Diagramm dar.|
||[format](/javascript/api/excel/excel.chart#format)|Kapselt die Formateigenschaften für den Diagrammbereich.|
||[Legende](/javascript/api/excel/excel.chart#legend)|Die Legende für das Diagramm.|
||[series](/javascript/api/excel/excel.chart#series)|Eine einzelne Datenreihe oder eine Sammlung von Datenreihen im Diagramm.|
||[title](/javascript/api/excel/excel.chart#title)|Gibt den Titel des angegebenen Diagramms an, einschließlich des Texts, der Sichtbarkeit, der Position und der Formatierung des Titels.|
||[SetData (SourceData: Range, seriesBy?: Excel. ChartSeriesBy)](/javascript/api/excel/excel.chart#setdata-sourcedata--seriesby-)|Setzt die Quelldaten für das Diagramm zurück.|
||[SetPosition (startCell: Bereichs \| Zeichenfolge, endCell?: Bereichs \| Zeichenfolge)](/javascript/api/excel/excel.chart#setposition-startcell--endcell-)|Positioniert das Diagramm im Verhältnis zu den Zellen im Arbeitsblatt.|
||[top](/javascript/api/excel/excel.chart#top)|Gibt den Abstand (in Punkt) vom oberen Rand des Objekts zum oberen Rand von Zeile 1 (auf einem Arbeitsblatt) oder zum oberen Rand des Diagrammbereichs (in einem Diagramm) an.|
||[width](/javascript/api/excel/excel.chart#width)|Gibt die Breite des Chart-Objekts in Punkten an.|
|[ChartAreaFormat](/javascript/api/excel/excel.chartareaformat)|[fill](/javascript/api/excel/excel.chartareaformat#fill)|Stellt die Füllung eines Objekts dar, einschließlich Informationen zur Hintergrundformatierung.|
||[font](/javascript/api/excel/excel.chartareaformat#font)|Stellt die Zeichenformatierung (Schriftart, Schriftgrad, Farbe usw.) für das aktuelle Objekt dar.|
|[ChartAxes](/javascript/api/excel/excel.chartaxes)|[categoryAxis](/javascript/api/excel/excel.chartaxes#categoryaxis)|Stellt die Rubrikenachse in einem Diagramm dar.|
||[seriesAxis](/javascript/api/excel/excel.chartaxes#seriesaxis)|Stellt die Reihenachse eines dreidimensionalen Diagramms dar.|
||[valueAxis](/javascript/api/excel/excel.chartaxes#valueaxis)|Stellt die Größenachse in einer Achse dar.|
|[ChartAxis](/javascript/api/excel/excel.chartaxis)|[majorUnit](/javascript/api/excel/excel.chartaxis#majorunit)|Stellt das Intervall zwischen zwei Hauptteilstrichen dar.|
||[maximum](/javascript/api/excel/excel.chartaxis#maximum)|Stellt den Maximalwert auf der Größenachse dar.|
||[minimum](/javascript/api/excel/excel.chartaxis#minimum)|Stellt den Mindestwert auf der Größenachse dar.|
||[minorUnit](/javascript/api/excel/excel.chartaxis#minorunit)|Stellt das Intervall zwischen zwei Hilfsteilstrichen dar.|
||[format](/javascript/api/excel/excel.chartaxis#format)|Stellt die Formatierung für ein Diagrammobjekt dar, einschließlich Linien- und Schriftartformatierung.|
||[majorGridlines](/javascript/api/excel/excel.chartaxis#majorgridlines)|Gibt ein Gitternetzlinien-Objekt zurück, das die Hauptgitternetzlinien für die angegebene Achse darstellt.|
||[minorGridlines](/javascript/api/excel/excel.chartaxis#minorgridlines)|Gibt ein Gitternetzlinien-Objekt zurück, das die Hilfsgitternetzlinien für die angegebene Achse darstellt.|
||[title](/javascript/api/excel/excel.chartaxis#title)|Stellt den Achsentitel dar.|
|[ChartAxisFormat](/javascript/api/excel/excel.chartaxisformat)|[font](/javascript/api/excel/excel.chartaxisformat#font)|Gibt die Schriftattribute (Schriftartname, Schriftgrad, Farbe usw.) für ein Diagrammachsen Element an.|
||[line](/javascript/api/excel/excel.chartaxisformat#line)|Gibt die Formatierung von Diagramm Leitungen an.|
|[ChartAxisTitle](/javascript/api/excel/excel.chartaxistitle)|[format](/javascript/api/excel/excel.chartaxistitle#format)|Gibt die Formatierung des Titels der Diagrammachse an.|
||[text](/javascript/api/excel/excel.chartaxistitle#text)|Gibt den Titel der Achse an.|
||[visible](/javascript/api/excel/excel.chartaxistitle#visible)|Gibt an, ob der Achsentitel visibile ist.|
|[ChartAxisTitleFormat](/javascript/api/excel/excel.chartaxistitleformat)|[font](/javascript/api/excel/excel.chartaxistitleformat#font)|Gibt die Schriftattribute des Diagrammachsen Titels an, beispielsweise Schriftartname, Schriftgrad, Farbe usw.|
|[ChartCollection](/javascript/api/excel/excel.chartcollection)|[Add (Typ: Excel. ChartType, SourceData: Range, seriesBy?: Excel. ChartSeriesBy)](/javascript/api/excel/excel.chartcollection#add-type--sourcedata--seriesby-)|Erstellt ein neues Diagramm.|
||[getItem(name: string)](/javascript/api/excel/excel.chartcollection#getitem-name-)|Ruft ein Diagramm über seinen Namen ab.|
||[getItemAt(index: number)](/javascript/api/excel/excel.chartcollection#getitemat-index-)|Ruft ein Diagramm anhand seiner Position in der Sammlung ab.|
||[count](/javascript/api/excel/excel.chartcollection#count)|Gibt die Anzahl der Diagramme im Arbeitsblatt zurück.|
||[items](/javascript/api/excel/excel.chartcollection#items)|Ruft die geladenen untergeordneten Elemente in dieser Sammlung ab.|
|[ChartDataLabelFormat](/javascript/api/excel/excel.chartdatalabelformat)|[fill](/javascript/api/excel/excel.chartdatalabelformat#fill)|Stellt die Füllung der aktuellen Diagrammdatenbeschriftung dar.|
||[font](/javascript/api/excel/excel.chartdatalabelformat#font)|Stellt die Zeichenformatierung (Schriftart, Schriftgrad, Farbe usw.) für eine Diagrammdatenbeschriftung dar.|
|[ChartDataLabels](/javascript/api/excel/excel.chartdatalabels)|[position](/javascript/api/excel/excel.chartdatalabels#position)|DataLabelPosition-Wert, der die Position der Datenbeschriftung darstellt.|
||[format](/javascript/api/excel/excel.chartdatalabels#format)|Gibt das Format von Diagrammdaten Beschriftungen an, das Füllungs-und Schriftartformatierungen enthält.|
||[Separator](/javascript/api/excel/excel.chartdatalabels#separator)|Zeichenfolge, die das Trennzeichen für die Datenbeschriftungen in einem Diagramm darstellt.|
||[showBubbleSize](/javascript/api/excel/excel.chartdatalabels#showbubblesize)|Gibt an, ob die Daten Beschriftungs Blasengröße sichtbar ist.|
||[showCategoryName](/javascript/api/excel/excel.chartdatalabels#showcategoryname)|Gibt an, ob der Name der Daten Beschriftungskategorie sichtbar ist.|
||[showLegendKey](/javascript/api/excel/excel.chartdatalabels#showlegendkey)|Gibt an, ob der Legendenschlüssel für die Datenbeschriftung sichtbar ist.|
||[showPercentage](/javascript/api/excel/excel.chartdatalabels#showpercentage)|Gibt an, ob der Prozentsatz der Datenbeschriftung sichtbar ist.|
||[showSeriesName](/javascript/api/excel/excel.chartdatalabels#showseriesname)|Gibt an, ob der Name der Daten Beschriftungs Reihe sichtbar ist.|
||[ShowValue](/javascript/api/excel/excel.chartdatalabels#showvalue)|Gibt an, ob der Wert der Datenbeschriftung sichtbar ist.|
|[ChartFill](/javascript/api/excel/excel.chartfill)|[clear()](/javascript/api/excel/excel.chartfill#clear--)|Entfernen der Füllfarbe eines Diagrammelements|
||[setSolidColor(color: string)](/javascript/api/excel/excel.chartfill#setsolidcolor-color-)|Legt die Füllung eines Diagrammelements auf einfarbige Füllung fest.|
|[ChartFont](/javascript/api/excel/excel.chartfont)|[bold](/javascript/api/excel/excel.chartfont#bold)|Stellt den Fett-Status der Schriftart dar.|
||[color](/javascript/api/excel/excel.chartfont#color)|HTML-Farb Codedarstellung der Textfarbe (z. b. #FF0000 steht für rot).|
||[italic](/javascript/api/excel/excel.chartfont#italic)|Stellt den Kursiv-Status der Schriftart dar.|
||[name](/javascript/api/excel/excel.chartfont#name)|Name der Schriftart (beispielsweise "Calibri")|
||[size](/javascript/api/excel/excel.chartfont#size)|Größe der Schriftart (z. b. 11)|
||[underline](/javascript/api/excel/excel.chartfont#underline)|Art der auf die Schriftart angewendeten Unterstreichung.|
|[ChartGridlines](/javascript/api/excel/excel.chartgridlines)|[format](/javascript/api/excel/excel.chartgridlines#format)|Stellt die Formatierung der Diagrammgitternetzlinien dar.|
||[visible](/javascript/api/excel/excel.chartgridlines#visible)|Gibt an, ob die Achsen Gitternetzlinien sichtbar sind.|
|[ChartGridlinesFormat](/javascript/api/excel/excel.chartgridlinesformat)|[line](/javascript/api/excel/excel.chartgridlinesformat#line)|Stellt die Formatierung der Diagrammlinien dar.|
|[ChartLegend](/javascript/api/excel/excel.chartlegend)|[Overlay](/javascript/api/excel/excel.chartlegend#overlay)|Gibt an, ob sich die Diagrammlegende mit dem Haupttextkörper des Diagramms überlappen soll.|
||[position](/javascript/api/excel/excel.chartlegend#position)|Gibt die Position der Legende im Diagramm an.|
||[format](/javascript/api/excel/excel.chartlegend#format)|Stellt die Formatierung für eine Diagrammlegende dar, einschließlich Füllung und Schriftartformatierung.|
||[visible](/javascript/api/excel/excel.chartlegend#visible)|Gibt an, ob die Chart Legend sichtbar ist.|
|[ChartLegendFormat](/javascript/api/excel/excel.chartlegendformat)|[fill](/javascript/api/excel/excel.chartlegendformat#fill)|Stellt die Füllung eines Objekts dar, einschließlich Informationen zur Hintergrundformatierung.|
||[font](/javascript/api/excel/excel.chartlegendformat#font)|Stellt die Schriftartattribute wie Schriftartname, Schriftgrad, Farbe usw. dar.|
|[ChartLineFormat](/javascript/api/excel/excel.chartlineformat)|[clear()](/javascript/api/excel/excel.chartlineformat#clear--)|Löschen der Linienformatierung eines Diagrammelements.|
||[color](/javascript/api/excel/excel.chartlineformat#color)|HTML-Farbcode, der die Farbe der Linien im Diagramm darstellt.|
|[ChartPoint](/javascript/api/excel/excel.chartpoint)|[format](/javascript/api/excel/excel.chartpoint#format)|Kapselt die Formateigenschaften eines Diagrammpunkts.|
||[value](/javascript/api/excel/excel.chartpoint#value)|Gibt den Wert eines Diagrammpunkts zurück.|
|[ChartPointFormat](/javascript/api/excel/excel.chartpointformat)|[fill](/javascript/api/excel/excel.chartpointformat#fill)|Stellt das Füllformat eines Diagramms dar, das Hintergrund Formatierungsinformationen enthält.|
|[ChartPointsCollection](/javascript/api/excel/excel.chartpointscollection)|[getItemAt(index: number)](/javascript/api/excel/excel.chartpointscollection#getitemat-index-)|Abrufen eines Punkts anhand seiner Position in der Datenreihe.|
||[count](/javascript/api/excel/excel.chartpointscollection#count)|Gibt die Anzahl der Diagrammpunkten in der Reihe zurück.|
||[items](/javascript/api/excel/excel.chartpointscollection#items)|Ruft die geladenen untergeordneten Elemente in dieser Sammlung ab.|
|[ChartSeries](/javascript/api/excel/excel.chartseries)|[name](/javascript/api/excel/excel.chartseries#name)|Gibt den Namen einer Datenreihe in einem Diagramm an.|
||[format](/javascript/api/excel/excel.chartseries#format)|Stellt die Formatierung für eine Diagrammdatenreihe dar, einschließlich Füllung und Linienformatierung.|
||[Punkt](/javascript/api/excel/excel.chartseries#points)|Gibt eine Auflistung aller Punkte in der Datenreihe zurück.|
|[ChartSeriesCollection](/javascript/api/excel/excel.chartseriescollection)|[getItemAt(index: number)](/javascript/api/excel/excel.chartseriescollection#getitemat-index-)|Ruft eine Datenreihe anhand ihrer Position in der Sammlung ab.|
||[count](/javascript/api/excel/excel.chartseriescollection#count)|Gibt die Anzahl der Datenreihen in der Sammlung zurück.|
||[items](/javascript/api/excel/excel.chartseriescollection#items)|Ruft die geladenen untergeordneten Elemente in dieser Sammlung ab.|
|[ChartSeriesFormat](/javascript/api/excel/excel.chartseriesformat)|[fill](/javascript/api/excel/excel.chartseriesformat#fill)|Stellt die Füllung einer Diagrammdatenreihe dar, einschließlich Informationen zur Hintergrundformatierung.|
||[line](/javascript/api/excel/excel.chartseriesformat#line)|Die Zeilenformatierung.|
|[ChartTitle](/javascript/api/excel/excel.charttitle)|[Overlay](/javascript/api/excel/excel.charttitle#overlay)|Gibt an, ob der Diagrammtitel im Diagramm überlagert wird.|
||[format](/javascript/api/excel/excel.charttitle#format)|Stellt die Formatierung für einen Diagrammtitel dar, einschließlich Füllung und Schriftartformatierung.|
||[text](/javascript/api/excel/excel.charttitle#text)|Gibt den Titel Text des Diagramms an.|
||[visible](/javascript/api/excel/excel.charttitle#visible)|Gibt an, ob der Diagrammtitel visibile ist.|
|[ChartTitleFormat](/javascript/api/excel/excel.charttitleformat)|[fill](/javascript/api/excel/excel.charttitleformat#fill)|Stellt die Füllung eines Objekts dar, einschließlich Informationen zur Hintergrundformatierung.|
||[font](/javascript/api/excel/excel.charttitleformat#font)|Stellt die Zeichenformatierung (Schriftart, Schriftgrad, Farbe usw.) für ein Objekt dar.|
|[NamedItem](/javascript/api/excel/excel.nameditem)|[getRange()](/javascript/api/excel/excel.nameditem#getrange--)|Ruft das Bereichsobjekt ab, das mit dem Namen verknüpft ist.|
||[name](/javascript/api/excel/excel.nameditem#name)|Der Name des Objekts.|
||[Typ](/javascript/api/excel/excel.nameditem#type)|Gibt den Typ des Werts an, der von der Formel des Namens zurückgegeben wird.|
||[value](/javascript/api/excel/excel.nameditem#value)|Stellt den Wert dar, der von der Formel des Namens berechnet wurde.|
||[visible](/javascript/api/excel/excel.nameditem#visible)|Gibt an, ob das Objekt sichtbar ist.|
|[NamedItemCollection](/javascript/api/excel/excel.nameditemcollection)|[getItem(name: string)](/javascript/api/excel/excel.nameditemcollection#getitem-name-)|Ruft ein NamedItem-Objekt mit seinem Namen ab.|
||[items](/javascript/api/excel/excel.nameditemcollection#items)|Ruft die geladenen untergeordneten Elemente in dieser Sammlung ab.|
|[Range](/javascript/api/excel/excel.range)|[clear(applyTo?: Excel.ClearApplyTo)](/javascript/api/excel/excel.range#clear-applyto-)|Löscht Bereichswerte, Format, Füllung, Rahmen usw.|
||[Delete (Shift: Excel. DeleteShiftDirection)](/javascript/api/excel/excel.range#delete-shift-)|Löscht die dem Bereich zugeordneten Zellen.|
||[formulas](/javascript/api/excel/excel.range#formulas)|Stellt die Formel in der A1-Schreibweise dar.|
||[formulasLocal](/javascript/api/excel/excel.range#formulaslocal)|Stellt die Formel in der A1-Schreibweise, Sprache des Benutzers und im Gebietsschema der Zahlenformatierung dar.|
||[getBoundingRect (anotherRange: Bereichs \| Zeichenfolge)](/javascript/api/excel/excel.range#getboundingrect-anotherrange-)|Ruft das kleinste Bereichsobjekt ab, das die angegebenen Bereiche umfasst.|
||[getCell(row: number, column: number)](/javascript/api/excel/excel.range#getcell-row--column-)|Ruft das Bereichsobjekt ab, das die einzelne Zelle basierend auf Zeilen- und Spaltenanzahl enthält.|
||[getColumn(column: number)](/javascript/api/excel/excel.range#getcolumn-column-)|Ruft eine Spalte ab, die im Bereich enthalten ist.|
||[getEntireColumn()](/javascript/api/excel/excel.range#getentirecolumn--)|Ruft ein Objekt ab, das die gesamte Spalte des Bereichs darstellt (beispielsweise, wenn der aktuelle Bereich die Zellen "B4: e11" darstellt, sein `getEntireColumn` Wert ist ein Bereich, der die Spalten "B:E" darstellt).|
||[getEntireRow()](/javascript/api/excel/excel.range#getentirerow--)|Ruft ein Objekt ab, das die gesamte Zeile des Bereichs darstellt (beispielsweise, wenn der aktuelle Bereich die Zellen "B4: e11" darstellt, sein `GetEntireRow` Wert ist ein Bereich, der die Zeilen "4:11") darstellt.|
||[getintersection (anotherRange: Bereichs \| Zeichenfolge)](/javascript/api/excel/excel.range#getintersection-anotherrange-)|Ruft das Bereichsobjekt ab, das die rechteckige Schnittmenge der angegebenen Bereiche darstellt.|
||[getLastCell()](/javascript/api/excel/excel.range#getlastcell--)|Ruft die letzte Zelle im Bereich ab.|
||[getLastColumn ()](/javascript/api/excel/excel.range#getlastcolumn--)|Ruft die letzte Spalte im Bereich ab.|
||[GetLastRow ()](/javascript/api/excel/excel.range#getlastrow--)|Ruft die letzte Zeile im Bereich ab.|
||[getOffsetRange(rowOffset: number, columnOffset: number)](/javascript/api/excel/excel.range#getoffsetrange-rowoffset--columnoffset-)|Ruft ein Objekt ab, das einen Bereich darstellt, der aus dem angegebenen Bereich versetzt ist.|
||[getRow(row: number)](/javascript/api/excel/excel.range#getrow-row-)|Ruft eine Zelle ab, die im Bereich enthalten ist.|
||[Insert (Shift: Excel. InsertShiftDirection)](/javascript/api/excel/excel.range#insert-shift-)|Fügt eine Zelle oder einen Zellbereich in das Arbeitsblatt anstelle dieses Bereichs ein, und verschiebt die anderen Zellen, um Platz zu schaffen.|
||[NumberFormat](/javascript/api/excel/excel.range#numberformat)|Stellt den Zahlenformatcode für Excel für den angegebenen Bereich dar.|
||[address](/javascript/api/excel/excel.range#address)|Gibt den Bereichsbezug im a1-Format an.|
||[addressLocal](/javascript/api/excel/excel.range#addresslocal)|Gibt den Bereichsverweis für den angegebenen Bereich in der Sprache des Benutzers an.|
||[cellCount](/javascript/api/excel/excel.range#cellcount)|Gibt die Anzahl der Zellen im Bereich an.|
||[ColumnCount](/javascript/api/excel/excel.range#columncount)|Gibt die Gesamtanzahl der Spalten im Bereich an.|
||[columnIndex](/javascript/api/excel/excel.range#columnindex)|Gibt die Spaltennummer der ersten Zelle im Bereich an.|
||[format](/javascript/api/excel/excel.range#format)|Gibt ein Formatobjekt zurück, das die Schriftart des Bereichs, Füllung, den Rahmen, die Ausrichtung und andere Eigenschaften verschachtelt.|
||[RowCount](/javascript/api/excel/excel.range#rowcount)|Gibt die Anzahl der Zeilen im Bereich zurück.|
||[rowIndex](/javascript/api/excel/excel.range#rowindex)|Gibt die Spaltenanzahl der ersten Zelle im Bereich zurück.|
||[text](/javascript/api/excel/excel.range#text)|Textwerte des angegebenen Bereichs.|
||[valueTypes](/javascript/api/excel/excel.range#valuetypes)|Gibt den Typ der Daten in jeder Zelle an.|
||[worksheet](/javascript/api/excel/excel.range#worksheet)|Das Arbeitsblatt, das den aktuellen Bereich enthält.|
||[select()](/javascript/api/excel/excel.range#select--)|Wählt den angegebenen Bereich in der Excel-Benutzeroberfläche aus.|
||[values](/javascript/api/excel/excel.range#values)|Stellt die Rohwerte des angegebenen Bereichs dar.|
|[RangeBorder](/javascript/api/excel/excel.rangeborder)|[color](/javascript/api/excel/excel.rangeborder#color)|HTML-Farbcode, der die Farbe der Rahmenlinie, des Formulars #RRGGBB (z. B.  "FFA500") oder als benannte HTML-Farbe (z. B. "orange") darstellt.|
||[sideIndex](/javascript/api/excel/excel.rangeborder#sideindex)|Konstanter Wert, der die bestimmte Seiten des Rahmens angibt.|
||[style](/javascript/api/excel/excel.rangeborder#style)|Einer der Konstanten der Linienart, die die Linienart für den Rahmen angibt.|
||[weight](/javascript/api/excel/excel.rangeborder#weight)|Gibt die Stärke des Rahmens um einen Bereich an.|
|[RangeBorderCollection](/javascript/api/excel/excel.rangebordercollection)|[GetItem (Index: Excel. BorderIndex)](/javascript/api/excel/excel.rangebordercollection#getitem-index-)|Ruft ein Rahmen-Objekt ab, das den Namen verwendet|
||[getItemAt(index: number)](/javascript/api/excel/excel.rangebordercollection#getitemat-index-)|Ruft ein Rahmen-Objekt ab, das den Namen verwendet|
||[count](/javascript/api/excel/excel.rangebordercollection#count)|Die Anzahl der Rahmen-Objekte in der Auflistung.|
||[items](/javascript/api/excel/excel.rangebordercollection#items)|Ruft die geladenen untergeordneten Elemente in dieser Sammlung ab.|
|[RangeFill](/javascript/api/excel/excel.rangefill)|[clear()](/javascript/api/excel/excel.rangefill#clear--)|Setzt den Hintergrund des Bereichs zurück.|
||[color](/javascript/api/excel/excel.rangefill#color)|HTML-Farbcode, der die Hintergrundfarbe des Formulars #RRGGBB (z. b. "FFA500") oder als benannte HTML-Farbe (beispielsweise "Orange") darstellt|
|[RangeFont](/javascript/api/excel/excel.rangefont)|[bold](/javascript/api/excel/excel.rangefont#bold)|Stellt den Fett-Status der Schriftart dar.|
||[color](/javascript/api/excel/excel.rangefont#color)|HTML-Farb Codedarstellung der Textfarbe (z. b. #FF0000 steht für rot).|
||[italic](/javascript/api/excel/excel.rangefont#italic)|Gibt den kursiven Status der Schriftart an.|
||[name](/javascript/api/excel/excel.rangefont#name)|Name der Schriftart (beispielsweise "Calibri").|
||[size](/javascript/api/excel/excel.rangefont#size)|Schriftgrad|
||[underline](/javascript/api/excel/excel.rangefont#underline)|Art der auf die Schriftart angewendeten Unterstreichung.|
|[RangeFormat](/javascript/api/excel/excel.rangeformat)|[horizontalAlignment](/javascript/api/excel/excel.rangeformat#horizontalalignment)|Stellt die horizontale Ausrichtung für das angegebene Objekt dar.|
||[borders](/javascript/api/excel/excel.rangeformat#borders)|Auflistung von Border-Objekten, die für den gesamten ausgewählten Bereich gelten.|
||[fill](/javascript/api/excel/excel.rangeformat#fill)|Gibt das Fill-Objekt an, das für den gesamten Bereich definiert ist.|
||[font](/javascript/api/excel/excel.rangeformat#font)|Gibt das Font-Objekt an, das für den gesamten Bereich definiert ist.|
||[verticalAlignment](/javascript/api/excel/excel.rangeformat#verticalalignment)|Stellt die vertikale Ausrichtung für das angegebene Objekt dar.|
||[wrapText](/javascript/api/excel/excel.rangeformat#wraptext)|Gibt an, ob der Text im Objekt von Excel umbrochen wird.|
|[Table](/javascript/api/excel/excel.table)|[delete()](/javascript/api/excel/excel.table#delete--)|Löscht die Tabelle.|
||[getDataBodyRange()](/javascript/api/excel/excel.table#getdatabodyrange--)|Ruft das Bereichsobjekt ab, das mit dem Datenteil der Tabelle verknüpft ist.|
||[getHeaderRowRange()](/javascript/api/excel/excel.table#getheaderrowrange--)|Ruft das Bereichsobjekt ab, das mit der Kopfzeile der Tabelle verknüpft ist.|
||[getRange()](/javascript/api/excel/excel.table#getrange--)|Ruft das Bereichsobjekt ab, das mit der gesamten Tabelle verknüpft ist.|
||[getTotalRowRange()](/javascript/api/excel/excel.table#gettotalrowrange--)|Ruft das Bereichsobjekt ab, das mit der Ergebniszeile der Tabelle verknüpft ist.|
||[name](/javascript/api/excel/excel.table#name)|Der Name der Tabelle.|
||[columns](/javascript/api/excel/excel.table#columns)|Stellt eine Auflistung aller Spalten in der Tabelle dar.|
||[id](/javascript/api/excel/excel.table#id)|Gibt einen Wert zurück, der das Arbeitsblatt in einer bestimmten Arbeitsmappe eindeutig identifiziert.|
||[rows](/javascript/api/excel/excel.table#rows)|Stellt eine Auflistung aller Zeilen in der Tabelle dar.|
||[Show Headers](/javascript/api/excel/excel.table#showheaders)|Gibt an, ob die Überschriftenzeile sichtbar ist.|
||[ShowTotals](/javascript/api/excel/excel.table#showtotals)|Gibt an, ob die Zeile Summe angezeigt wird.|
||[style](/javascript/api/excel/excel.table#style)|Konstanter Wert, der das Tabellenformat darstellt.|
|[TableCollection](/javascript/api/excel/excel.tablecollection)|[Add (Address: Bereichs \| Zeichenfolge, hasHeaders: Boolean)](/javascript/api/excel/excel.tablecollection#add-address--hasheaders-)|Erstellen Sie eine neue Tabelle.|
||[getItem(key: string)](/javascript/api/excel/excel.tablecollection#getitem-key-)|Ruft eine Tabelle anhand des Namens oder der ID ab.|
||[getItemAt(index: number)](/javascript/api/excel/excel.tablecollection#getitemat-index-)|Ruft eine Tabelle anhand ihrer Position in der Auflistung ab.|
||[count](/javascript/api/excel/excel.tablecollection#count)|Gibt die Anzahl der Tabellen in der Arbeitsmappe zurück.|
||[items](/javascript/api/excel/excel.tablecollection#items)|Ruft die geladenen untergeordneten Elemente in dieser Sammlung ab.|
|[TableColumn](/javascript/api/excel/excel.tablecolumn)|[delete()](/javascript/api/excel/excel.tablecolumn#delete--)|Löscht die Spalte aus der Tabelle.|
||[getDataBodyRange()](/javascript/api/excel/excel.tablecolumn#getdatabodyrange--)|Ruft das Bereichsobjekt ab, das mit dem Datenteil der Spalte verknüpft ist.|
||[getHeaderRowRange()](/javascript/api/excel/excel.tablecolumn#getheaderrowrange--)|Ruft das Bereichsobjekt ab, das mit der Überschriftenzeile der Spalte verknüpft ist.|
||[getRange()](/javascript/api/excel/excel.tablecolumn#getrange--)|Ruft das Bereichsobjekt ab, das mit der gesamten Spalte verknüpft ist.|
||[getTotalRowRange()](/javascript/api/excel/excel.tablecolumn#gettotalrowrange--)|Ruft das Bereichsobjekt ab, das mit der Ergebniszeile der Spalte verknüpft ist.|
||[name](/javascript/api/excel/excel.tablecolumn#name)|Gibt den Namen der Tabellenspalte an.|
||[id](/javascript/api/excel/excel.tablecolumn#id)|Gibt einen eindeutigen Schlüssel an, der die Spalte in der Tabelle angibt.|
||[Index](/javascript/api/excel/excel.tablecolumn#index)|Gibt die Indexnummer der Spalte in der Spaltenauflistung der Tabelle zurück.|
||[values](/javascript/api/excel/excel.tablecolumn#values)|Stellt die Rohwerte des angegebenen Bereichs dar.|
|[TableColumnCollection](/javascript/api/excel/excel.tablecolumncollection)|[Add (Index?: Number, Values?: Array<Array<Boolean \| String \| Number>> \| Boolean \| String \| Number, Name?: String)](/javascript/api/excel/excel.tablecolumncollection#add-index--values--name-)|Fügt der Tabelle eine neue Spalte hinzu.|
||[GetItem (Key: Number \| String)](/javascript/api/excel/excel.tablecolumncollection#getitem-key-)|Ruft ein Spaltenobjekt nach Name oder ID ab.|
||[getItemAt(index: number)](/javascript/api/excel/excel.tablecolumncollection#getitemat-index-)|Ruft eine Spalte anhand ihrer Position in der Auflistung ab.|
||[count](/javascript/api/excel/excel.tablecolumncollection#count)|Gibt die Anzahl der Spalten in der Tabelle an.|
||[items](/javascript/api/excel/excel.tablecolumncollection#items)|Ruft die geladenen untergeordneten Elemente in dieser Sammlung ab.|
|[TableRow](/javascript/api/excel/excel.tablerow)|[delete()](/javascript/api/excel/excel.tablerow#delete--)|Löscht die Zeile aus der Tabelle.|
||[getRange()](/javascript/api/excel/excel.tablerow#getrange--)|Ruft das Bereichsobjekt ab, das mit der gesamten Zeile verknüpft ist.|
||[Index](/javascript/api/excel/excel.tablerow#index)|Gibt die Indexnummer der Zeile in der Zeilenauflistung der Tabelle zurück.|
||[values](/javascript/api/excel/excel.tablerow#values)|Stellt die Rohwerte des angegebenen Bereichs dar.|
|[TableRowCollection](/javascript/api/excel/excel.tablerowcollection)|[Add (Index?: Number, Values?: Array<Array<Boolean \| String \| Number>> \| Boolean \| String \| Number)](/javascript/api/excel/excel.tablerowcollection#add-index--values-)|Fügt der Tabelle mindestens eine Zeile hinzu.|
||[getItemAt(index: number)](/javascript/api/excel/excel.tablerowcollection#getitemat-index-)|Ruft eine Zeile anhand ihrer Position in der Auflistung ab.|
||[count](/javascript/api/excel/excel.tablerowcollection#count)|Gibt die Anzahl der Zeilen in der Tabelle zurück.|
||[items](/javascript/api/excel/excel.tablerowcollection#items)|Ruft die geladenen untergeordneten Elemente in dieser Sammlung ab.|
|[Workbook](/javascript/api/excel/excel.workbook)|[Funktion getselectedrange ()](/javascript/api/excel/excel.workbook#getselectedrange--)|Ruft den aktuell ausgewählten einzelnen Bereich aus der Arbeitsmappe ab.|
||[application](/javascript/api/excel/excel.workbook#application)|Stellt die Excel-Anwendungsinstanz dar, die diese Arbeitsmappe enthält.|
||[bindings](/javascript/api/excel/excel.workbook#bindings)|Stellt eine Auflistung aller Bindungen dar, die Teil der Arbeitsmappe sind.|
||[Namen](/javascript/api/excel/excel.workbook#names)|Stellt eine Auflistung der benannten Elemente des Arbeitsmappenbereichs dar (benannte Bereiche und Konstanten).|
||[Tabellen](/javascript/api/excel/excel.workbook#tables)|Stellt eine Auflistung der mit der Arbeitsmappe verknüpften Tabellen dar.|
||[Arbeitsblätter](/javascript/api/excel/excel.workbook#worksheets)|Stellt eine Auflistung der mit der Arbeitsmappe verknüpften Arbeitsblätter dar.|
|[Arbeitsblatt](/javascript/api/excel/excel.worksheet)|[activate()](/javascript/api/excel/excel.worksheet#activate--)|Aktivieren Sie das Arbeitsblatt in der Excel-Benutzeroberfläche.|
||[delete()](/javascript/api/excel/excel.worksheet#delete--)|Löscht das Arbeitsblatt aus der Arbeitsmappe.|
||[getCell(row: number, column: number)](/javascript/api/excel/excel.worksheet#getcell-row--column-)|Ruft das Bereichsobjekt ab, das die einzelne Zelle basierend auf Zeilen- und Spaltenanzahl enthält.|
||[GetRange (Address?: String)](/javascript/api/excel/excel.worksheet#getrange-address-)|Ruft das Range-Objekt ab, das einen einzelnen rechteckigen Zellenblock darstellt, der durch die Adresse oder den Namen angegeben wird.|
||[name](/javascript/api/excel/excel.worksheet#name)|Der Anzeigename des Arbeitsblatts.|
||[position](/javascript/api/excel/excel.worksheet#position)|Die nullbasiert Position des Arbeitsblatts in der Arbeitsmappe.|
||[Diagramme](/javascript/api/excel/excel.worksheet#charts)|Gibt eine Auflistung von Diagrammen zurück, die Teil des Arbeitsblatts sind.|
||[id](/javascript/api/excel/excel.worksheet#id)|Gibt einen Wert zurück, der das Arbeitsblatt in einer bestimmten Arbeitsmappe eindeutig identifiziert.|
||[Tabellen](/javascript/api/excel/excel.worksheet#tables)|Gibt die Sammlung von Tabellen zurück, die Teil des Arbeitsblatts sind.|
||[visibility](/javascript/api/excel/excel.worksheet#visibility)|Die Sichtbarkeit des Arbeitsblatts.|
|[WorksheetCollection](/javascript/api/excel/excel.worksheetcollection)|[Add (Name?: Zeichenfolge)](/javascript/api/excel/excel.worksheetcollection#add-name-)|Fügt der Arbeitsmappe ein neues Arbeitsblatt hinzu.|
||[getActiveWorksheet()](/javascript/api/excel/excel.worksheetcollection#getactiveworksheet--)|Ruft das derzeit aktive Arbeitsblatt in der Arbeitsmappe ab.|
||[getItem(key: string)](/javascript/api/excel/excel.worksheetcollection#getitem-key-)|Ruft das Arbeitsblattobjekt mithilfe des Namens oder der ID ab.|
||[items](/javascript/api/excel/excel.worksheetcollection#items)|Ruft die geladenen untergeordneten Elemente in dieser Sammlung ab.|
