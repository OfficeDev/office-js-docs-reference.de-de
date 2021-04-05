| Klasse | Felder | Beschreibung |
|:---|:---|:---|
|[Anwendung](/javascript/api/excel/excel.application)|[calculationEngineVersion](/javascript/api/excel/excel.application#calculationengineversion)|Gibt die Version der Excel-Berechnungsmaschine zurück, die für die letzte vollständige Neuberechnung verwendet wurde.|
||[calculationState](/javascript/api/excel/excel.application#calculationstate)|Gibt den Berechnungszustand der Anwendung zurück.|
||[iterativeCalculation](/javascript/api/excel/excel.application#iterativecalculation)|Gibt die iterativen Berechnungseinstellungen zurück.|
||[suspendScreenUpdatingUntilNextSync()](/javascript/api/excel/excel.application#suspendscreenupdatinguntilnextsync--)|Die Bildschirmaktualisierung wird angehalten, bis die nächste `context.sync()` aufgerufen wird.|
|[AutoFilter](/javascript/api/excel/excel.autofilter)|[apply(range: Range \| string, columnIndex?: number, criteria?: Excel.FilterCriteria)](/javascript/api/excel/excel.autofilter#apply-range--columnindex--criteria-)|Wendet den AutoFilter auf einen Bereich an.|
||[clearCriteria()](/javascript/api/excel/excel.autofilter#clearcriteria--)|Löscht die Filterkriterien von AutoFilter.|
||[getRange()](/javascript/api/excel/excel.autofilter#getrange--)|Gibt das `Range` Objekt zurück, das den Bereich darstellt, auf den der AutoFilter angewendet wird.|
||[getRangeOrNullObject()](/javascript/api/excel/excel.autofilter#getrangeornullobject--)|Gibt das `Range` Objekt zurück, das den Bereich darstellt, auf den der AutoFilter angewendet wird.|
||[criteria](/javascript/api/excel/excel.autofilter#criteria)|Ein Array, das alle Filterkriterien in dem per AutoFilter gefilterten Bereich enthält.|
||[enabled](/javascript/api/excel/excel.autofilter#enabled)|Gibt an, ob der AutoFilter aktiviert ist.|
||[isDataFiltered](/javascript/api/excel/excel.autofilter#isdatafiltered)|Gibt an, ob der AutoFilter Filterkriterien enthält.|
||[reapply()](/javascript/api/excel/excel.autofilter#reapply--)|Wendet das angegebene Autofilter-Objekt an, das aktuell für den Bereich aktiv ist.|
||[remove()](/javascript/api/excel/excel.autofilter#remove--)|Entfernt den AutoFilter für den Bereich.|
|[CellBorder](/javascript/api/excel/excel.cellborder)|[color](/javascript/api/excel/excel.cellborder#color)|Stellt die `color` Eigenschaft eines einzelnen Rahmens dar.|
||[style](/javascript/api/excel/excel.cellborder#style)|Stellt die `style` Eigenschaft eines einzelnen Rahmens dar.|
||[tintAndShade](/javascript/api/excel/excel.cellborder#tintandshade)|Stellt die `tintAndShade` Eigenschaft eines einzelnen Rahmens dar.|
||[weight](/javascript/api/excel/excel.cellborder#weight)|Stellt die `weight` Eigenschaft eines einzelnen Rahmens dar.|
|[CellBorderCollection](/javascript/api/excel/excel.cellbordercollection)|[bottom](/javascript/api/excel/excel.cellbordercollection#bottom)|Stellt die `format.borders.bottom` Eigenschaft dar.|
||[diagonalDown](/javascript/api/excel/excel.cellbordercollection#diagonaldown)|Stellt die `format.borders.diagonalDown` Eigenschaft dar.|
||[diagonalUp](/javascript/api/excel/excel.cellbordercollection#diagonalup)|Stellt die `format.borders.diagonalUp` Eigenschaft dar.|
||[horizontal](/javascript/api/excel/excel.cellbordercollection#horizontal)|Stellt die `format.borders.horizontal` Eigenschaft dar.|
||[left](/javascript/api/excel/excel.cellbordercollection#left)|Stellt die `format.borders.left` Eigenschaft dar.|
||[right](/javascript/api/excel/excel.cellbordercollection#right)|Stellt die `format.borders.right` Eigenschaft dar.|
||[top](/javascript/api/excel/excel.cellbordercollection#top)|Stellt die `format.borders.top` Eigenschaft dar.|
||[vertical](/javascript/api/excel/excel.cellbordercollection#vertical)|Stellt die `format.borders.vertical` Eigenschaft dar.|
|[CellProperties](/javascript/api/excel/excel.cellproperties)|[address](/javascript/api/excel/excel.cellproperties#address)|Stellt die `address` Eigenschaft dar.|
||[addressLocal](/javascript/api/excel/excel.cellproperties#addresslocal)|Stellt die `addressLocal` Eigenschaft dar.|
||[hidden](/javascript/api/excel/excel.cellproperties#hidden)|Stellt die `hidden` Eigenschaft dar.|
|[CellPropertiesFill](/javascript/api/excel/excel.cellpropertiesfill)|[color](/javascript/api/excel/excel.cellpropertiesfill#color)|Stellt die `format.fill.color` Eigenschaft dar.|
||[pattern](/javascript/api/excel/excel.cellpropertiesfill#pattern)|Stellt die `format.fill.pattern` Eigenschaft dar.|
||[patternColor](/javascript/api/excel/excel.cellpropertiesfill#patterncolor)|Stellt die `format.fill.patternColor` Eigenschaft dar.|
||[patternTintAndShade](/javascript/api/excel/excel.cellpropertiesfill#patterntintandshade)|Stellt die `format.fill.patternTintAndShade` Eigenschaft dar.|
||[tintAndShade](/javascript/api/excel/excel.cellpropertiesfill#tintandshade)|Stellt die `format.fill.tintAndShade` Eigenschaft dar.|
|[CellPropertiesFont](/javascript/api/excel/excel.cellpropertiesfont)|[bold](/javascript/api/excel/excel.cellpropertiesfont#bold)|Stellt die `format.font.bold` Eigenschaft dar.|
||[color](/javascript/api/excel/excel.cellpropertiesfont#color)|Stellt die `format.font.color` Eigenschaft dar.|
||[italic](/javascript/api/excel/excel.cellpropertiesfont#italic)|Stellt die `format.font.italic` Eigenschaft dar.|
||[name](/javascript/api/excel/excel.cellpropertiesfont#name)|Stellt die `format.font.name` Eigenschaft dar.|
||[size](/javascript/api/excel/excel.cellpropertiesfont#size)|Stellt die `format.font.size` Eigenschaft dar.|
||[strikethrough](/javascript/api/excel/excel.cellpropertiesfont#strikethrough)|Stellt die `format.font.strikethrough` Eigenschaft dar.|
||[subscript](/javascript/api/excel/excel.cellpropertiesfont#subscript)|Stellt die `format.font.subscript` Eigenschaft dar.|
||[superscript](/javascript/api/excel/excel.cellpropertiesfont#superscript)|Stellt die `format.font.superscript` Eigenschaft dar.|
||[tintAndShade](/javascript/api/excel/excel.cellpropertiesfont#tintandshade)|Stellt die `format.font.tintAndShade` Eigenschaft dar.|
||[underline](/javascript/api/excel/excel.cellpropertiesfont#underline)|Stellt die `format.font.underline` Eigenschaft dar.|
|[CellPropertiesFormat](/javascript/api/excel/excel.cellpropertiesformat)|[autoIndent](/javascript/api/excel/excel.cellpropertiesformat#autoindent)|Stellt die `autoIndent` Eigenschaft dar.|
||[borders](/javascript/api/excel/excel.cellpropertiesformat#borders)|Stellt die `borders` Eigenschaft dar.|
||[fill](/javascript/api/excel/excel.cellpropertiesformat#fill)|Stellt die `fill` Eigenschaft dar.|
||[font](/javascript/api/excel/excel.cellpropertiesformat#font)|Stellt die `font` Eigenschaft dar.|
||[horizontalAlignment](/javascript/api/excel/excel.cellpropertiesformat#horizontalalignment)|Stellt die `horizontalAlignment` Eigenschaft dar.|
||[indentLevel](/javascript/api/excel/excel.cellpropertiesformat#indentlevel)|Stellt die `indentLevel` Eigenschaft dar.|
||[protection](/javascript/api/excel/excel.cellpropertiesformat#protection)|Stellt die `protection` Eigenschaft dar.|
||[readingOrder](/javascript/api/excel/excel.cellpropertiesformat#readingorder)|Stellt die `readingOrder` Eigenschaft dar.|
||[shrinkToFit](/javascript/api/excel/excel.cellpropertiesformat#shrinktofit)|Stellt die `shrinkToFit` Eigenschaft dar.|
||[textOrientation](/javascript/api/excel/excel.cellpropertiesformat#textorientation)|Stellt die `textOrientation` Eigenschaft dar.|
||[useStandardHeight](/javascript/api/excel/excel.cellpropertiesformat#usestandardheight)|Stellt die `useStandardHeight` Eigenschaft dar.|
||[useStandardWidth](/javascript/api/excel/excel.cellpropertiesformat#usestandardwidth)|Stellt die `useStandardWidth` Eigenschaft dar.|
||[verticalAlignment](/javascript/api/excel/excel.cellpropertiesformat#verticalalignment)|Stellt die `verticalAlignment` Eigenschaft dar.|
||[wrapText](/javascript/api/excel/excel.cellpropertiesformat#wraptext)|Stellt die `wrapText` Eigenschaft dar.|
|[CellPropertiesProtection](/javascript/api/excel/excel.cellpropertiesprotection)|[formulaHidden](/javascript/api/excel/excel.cellpropertiesprotection#formulahidden)|Stellt die `format.protection.formulaHidden` Eigenschaft dar.|
||[locked](/javascript/api/excel/excel.cellpropertiesprotection#locked)|Stellt die `format.protection.locked` Eigenschaft dar.|
|[ChangedEventDetail](/javascript/api/excel/excel.changedeventdetail)|[valueAfter](/javascript/api/excel/excel.changedeventdetail#valueafter)|Stellt den Wert nach der Änderung dar.|
||[valueBefore](/javascript/api/excel/excel.changedeventdetail#valuebefore)|Stellt den Wert vor der Änderung dar.|
||[valueTypeAfter](/javascript/api/excel/excel.changedeventdetail#valuetypeafter)|Stellt den Werttyp nach der Änderung dar.|
||[valueTypeBefore](/javascript/api/excel/excel.changedeventdetail#valuetypebefore)|Stellt den Werttyp vor der Änderung dar.|
|[Chart](/javascript/api/excel/excel.chart)|[activate()](/javascript/api/excel/excel.chart#activate--)|Aktiviert das Diagramm auf der Excel-Benutzeroberfläche.|
||[pivotOptions](/javascript/api/excel/excel.chart#pivotoptions)|Verkapselt die Optionen für ein PivotChart.|
|[ChartAreaFormat](/javascript/api/excel/excel.chartareaformat)|[colorScheme](/javascript/api/excel/excel.chartareaformat#colorscheme)|Gibt das Farbschema des Diagramms an.|
||[roundedCorners](/javascript/api/excel/excel.chartareaformat#roundedcorners)|Gibt an, ob der Diagrammbereich des Diagramms abgerundete Ecken hat.|
|[ChartAxis](/javascript/api/excel/excel.chartaxis)|[linkNumberFormat](/javascript/api/excel/excel.chartaxis#linknumberformat)|Gibt an, ob das Zahlenformat mit den Zellen verknüpft ist.|
|[ChartBinOptions](/javascript/api/excel/excel.chartbinoptions)|[allowOverflow](/javascript/api/excel/excel.chartbinoptions#allowoverflow)|Gibt an, ob der Bin overflow in einem Histogramm- oder Paretodiagramm aktiviert ist.|
||[allowUnderflow](/javascript/api/excel/excel.chartbinoptions#allowunderflow)|Gibt an, ob der Bin-Underflow in einem Histogramm- oder Paretodiagramm aktiviert ist.|
||[count](/javascript/api/excel/excel.chartbinoptions#count)|Gibt die Binanzahl eines Histogramms oder Eines Paretodiagramms an.|
||[overflowValue](/javascript/api/excel/excel.chartbinoptions#overflowvalue)|Gibt den Bin-Overflow-Wert eines Histogramms oder Paretodiagramms an.|
||[type](/javascript/api/excel/excel.chartbinoptions#type)|Gibt den Typ des Bins für ein Histogramm- oder Paretodiagramm an.|
||[underflowValue](/javascript/api/excel/excel.chartbinoptions#underflowvalue)|Gibt den Bin-Underflow-Wert eines Histogramm- oder Paretodiagramms an.|
||[width](/javascript/api/excel/excel.chartbinoptions#width)|Gibt den Wert der Binbreite eines Histogramms oder Eines Paretodiagramms an.|
|[ChartBoxwhiskerOptions](/javascript/api/excel/excel.chartboxwhiskeroptions)|[quartileCalculation](/javascript/api/excel/excel.chartboxwhiskeroptions#quartilecalculation)|Gibt an, ob der Quartilberechnungstyp eines Feld- und Whiskerdiagramms verwendet wird.|
||[showInnerPoints](/javascript/api/excel/excel.chartboxwhiskeroptions#showinnerpoints)|Gibt an, ob innere Punkte in einem Feld- und Whiskerdiagramm angezeigt werden.|
||[showMeanLine](/javascript/api/excel/excel.chartboxwhiskeroptions#showmeanline)|Gibt an, ob die mittlere Linie in einem Feld- und Whiskerdiagramm angezeigt wird.|
||[showMeanMarker](/javascript/api/excel/excel.chartboxwhiskeroptions#showmeanmarker)|Gibt an, ob die mittlere Markierung in einem Feld- und Whiskerdiagramm angezeigt wird.|
||[showOutlierPoints](/javascript/api/excel/excel.chartboxwhiskeroptions#showoutlierpoints)|Gibt an, ob Ausreißerpunkte in einem Feld- und Whiskerdiagramm angezeigt werden.|
|[ChartDataLabel](/javascript/api/excel/excel.chartdatalabel)|[linkNumberFormat](/javascript/api/excel/excel.chartdatalabel#linknumberformat)|Gibt an, ob das Zahlenformat mit den Zellen verknüpft ist (sodass sich das Zahlenformat in den Beschriftungen ändert, wenn es sich in den Zellen ändert).|
|[ChartDataLabels](/javascript/api/excel/excel.chartdatalabels)|[linkNumberFormat](/javascript/api/excel/excel.chartdatalabels#linknumberformat)|Gibt an, ob das Zahlenformat mit den Zellen verknüpft ist.|
|[ChartErrorBars](/javascript/api/excel/excel.charterrorbars)|[endStyleCap](/javascript/api/excel/excel.charterrorbars#endstylecap)|Gibt an, ob Fehlerleisten eine Endformatvorlage haben.|
||[include](/javascript/api/excel/excel.charterrorbars#include)|Gibt an, welche Teile der Fehlerindikatoren enthalten sein sollen.|
||[format](/javascript/api/excel/excel.charterrorbars#format)|Gibt den Formatierungstyp der Fehlerindikatoren an.|
||[Typ](/javascript/api/excel/excel.charterrorbars#type)|Der Bereichstyp, der durch die Fehlerindikatoren gekennzeichnet ist.|
||[visible](/javascript/api/excel/excel.charterrorbars#visible)|Gibt an, ob die Fehleranzeigen angezeigt werden.|
|[ChartErrorBarsFormat](/javascript/api/excel/excel.charterrorbarsformat)|[line](/javascript/api/excel/excel.charterrorbarsformat#line)|Stellt die Formatierung der Diagrammlinie dar.|
|[ChartMapOptions](/javascript/api/excel/excel.chartmapoptions)|[labelStrategy](/javascript/api/excel/excel.chartmapoptions#labelstrategy)|Gibt die Strategie für Serienkartenbeschriftungen eines Regionenkartendiagramms an.|
||[level](/javascript/api/excel/excel.chartmapoptions#level)|Gibt die Datenreihenzuordnungsebene eines Regionenkartendiagramms an.|
||[projectionType](/javascript/api/excel/excel.chartmapoptions#projectiontype)|Gibt den Datenreihenprojektionstyp eines Regionenkartendiagramms an.|
|[ChartPivotOptions](/javascript/api/excel/excel.chartpivotoptions)|[showAxisFieldButtons](/javascript/api/excel/excel.chartpivotoptions#showaxisfieldbuttons)|Gibt an, ob die Achsenfeldschaltflächen in einem PivotChart angezeigt werden.|
||[showLegendFieldButtons](/javascript/api/excel/excel.chartpivotoptions#showlegendfieldbuttons)|Gibt an, ob die Legendenfeldschaltflächen in einem PivotChart angezeigt werden.|
||[showReportFilterFieldButtons](/javascript/api/excel/excel.chartpivotoptions#showreportfilterfieldbuttons)|Gibt an, ob die Berichtsfilterfeldschaltflächen in einem PivotChart angezeigt werden sollen.|
||[showValueFieldButtons](/javascript/api/excel/excel.chartpivotoptions#showvaluefieldbuttons)|Gibt an, ob die Schaltflächen des Anzeigewertfelds in einem PivotChart angezeigt werden.|
|[ChartSeries](/javascript/api/excel/excel.chartseries)|[bubbleScale](/javascript/api/excel/excel.chartseries#bubblescale)|Dies kann ein ganzzahliger Wert von 0 (null) bis 300 sein, der einem Prozentsatz der Standardgröße darstellt.|
||[gradientMaximumColor](/javascript/api/excel/excel.chartseries#gradientmaximumcolor)|Gibt die Farbe für den maximal zulässigen Wert einer Diagrammreihe einer Region an.|
||[gradientMaximumType](/javascript/api/excel/excel.chartseries#gradientmaximumtype)|Gibt den Typ für den maximal zulässigen Wert einer Diagrammreihe für Regionen an.|
||[gradientMaximumValue](/javascript/api/excel/excel.chartseries#gradientmaximumvalue)|Gibt den maximal zulässigen Wert einer Diagrammreihe für Regionen an.|
||[gradientMidpointColor](/javascript/api/excel/excel.chartseries#gradientmidpointcolor)|Gibt die Farbe für den Mittelpunktwert einer Diagrammreihe für Regionenkarten an.|
||[gradientMidpointType](/javascript/api/excel/excel.chartseries#gradientmidpointtype)|Gibt den Typ für den Mittelpunktwert einer Kartenreihe für Regionen an.|
||[gradientMidpointValue](/javascript/api/excel/excel.chartseries#gradientmidpointvalue)|Gibt den Mittelpunktwert einer Diagrammreihe für Regionenkarten an.|
||[gradientMinimumColor](/javascript/api/excel/excel.chartseries#gradientminimumcolor)|Gibt die Farbe für den Minimalwert einer Kartenreihe für Regionen an.|
||[gradientMinimumType](/javascript/api/excel/excel.chartseries#gradientminimumtype)|Gibt den Typ für den Minimalwert einer Kartenreihe für Regionen an.|
||[gradientMinimumValue](/javascript/api/excel/excel.chartseries#gradientminimumvalue)|Gibt den Minimalwert einer Diagrammreihe für Regionen an.|
||[gradientStyle](/javascript/api/excel/excel.chartseries#gradientstyle)|Gibt die Farbverlaufsart der Datenreihe eines Regionendiagramms an.|
||[invertColor](/javascript/api/excel/excel.chartseries#invertcolor)|Gibt die Füllfarbe für negative Datenpunkte in einer Datenreihe an.|
||[parentLabelStrategy](/javascript/api/excel/excel.chartseries#parentlabelstrategy)|Gibt den Strategiebereich der übergeordneten Beschriftung der Datenreihe für ein Strukturkartendiagramm an.|
||[binOptions](/javascript/api/excel/excel.chartseries#binoptions)|Verkapselt die Intervalloptionen für Histogramme und Pareto-Diagramme.|
||[boxwhiskerOptions](/javascript/api/excel/excel.chartseries#boxwhiskeroptions)|Verkapselt die Optionen für Kastengrafikdiagramme.|
||[mapOptions](/javascript/api/excel/excel.chartseries#mapoptions)|Verkapselt die Optionen für ein Bereichsdiagramm.|
||[xErrorBars](/javascript/api/excel/excel.chartseries#xerrorbars)|Stellt das Fehlerbalkenobjekt für eine Datenreihe dar.|
||[yErrorBars](/javascript/api/excel/excel.chartseries#yerrorbars)|Stellt das Fehlerbalkenobjekt für eine Datenreihe dar.|
||[showConnectorLines](/javascript/api/excel/excel.chartseries#showconnectorlines)|Gibt an, ob Verbindungslinien in Wasserfalldiagrammen angezeigt werden.|
||[showLeaderLines](/javascript/api/excel/excel.chartseries#showleaderlines)|Gibt an, ob Führungslinien für jede Datenbeschriftung in der Datenreihe angezeigt werden.|
||[splitValue](/javascript/api/excel/excel.chartseries#splitvalue)|Gibt den Schwellenwert an, der zwei Abschnitte eines Kreis-aus-Kreis-Diagramms oder eines Balken-aus-Kreis-Diagramms trennt.|
|[ChartTrendlineLabel](/javascript/api/excel/excel.charttrendlinelabel)|[linkNumberFormat](/javascript/api/excel/excel.charttrendlinelabel#linknumberformat)|Gibt an, ob das Zahlenformat mit den Zellen verknüpft ist (sodass sich das Zahlenformat in den Beschriftungen ändert, wenn es sich in den Zellen ändert).|
|[ColumnProperties](/javascript/api/excel/excel.columnproperties)|[address](/javascript/api/excel/excel.columnproperties#address)|Stellt die `address` Eigenschaft dar.|
||[addressLocal](/javascript/api/excel/excel.columnproperties#addresslocal)|Stellt die `addressLocal` Eigenschaft dar.|
||[columnIndex](/javascript/api/excel/excel.columnproperties#columnindex)|Stellt die `columnIndex` Eigenschaft dar.|
|[ConditionalFormat](/javascript/api/excel/excel.conditionalformat)|[getRanges()](/javascript/api/excel/excel.conditionalformat#getranges--)|Gibt den zurück, der einen oder mehrere rechteckige Bereiche umfasst, auf die das `RangeAreas` konditonale Format angewendet wird.|
|[DataValidation](/javascript/api/excel/excel.datavalidation)|[getInvalidCells()](/javascript/api/excel/excel.datavalidation#getinvalidcells--)|Gibt ein `RangeAreas` Objekt mit einem oder mehreren rechteckigen Bereichen mit ungültigen Zellwerten zurück.|
||[getInvalidCellsOrNullObject()](/javascript/api/excel/excel.datavalidation#getinvalidcellsornullobject--)|Gibt ein `RangeAreas` Objekt mit einem oder mehreren rechteckigen Bereichen mit ungültigen Zellwerten zurück.|
|[FilterCriteria](/javascript/api/excel/excel.filtercriteria)|[subField](/javascript/api/excel/excel.filtercriteria#subfield)|Die Eigenschaft, die vom Filter verwendet wird, um einen Rich-Filter auf Rich-Werte zu verwenden.|
|[GeometricShape](/javascript/api/excel/excel.geometricshape)|[id](/javascript/api/excel/excel.geometricshape#id)|Gibt die ID der Form zurück.|
||[shape](/javascript/api/excel/excel.geometricshape#shape)|Gibt das `Shape` Objekt für die geometrische Form zurück.|
|[GroupShapeCollection](/javascript/api/excel/excel.groupshapecollection)|[getCount()](/javascript/api/excel/excel.groupshapecollection#getcount--)|Gibt die Anzahl der Formen in der Formgruppe zurück.|
||[getItem(key: string)](/javascript/api/excel/excel.groupshapecollection#getitem-key-)|Ruft ein Shape mit seinem Namen oder seiner ID ab.|
||[getItemAt(index: number)](/javascript/api/excel/excel.groupshapecollection#getitemat-index-)|Ruft eine Form anhand ihrer Position in der Auflistung ab.|
||[items](/javascript/api/excel/excel.groupshapecollection#items)|Ruft die geladenen untergeordneten Elemente in dieser Sammlung ab.|
|[HeaderFooter](/javascript/api/excel/excel.headerfooter)|[centerFooter](/javascript/api/excel/excel.headerfooter#centerfooter)|Die zentriert Fußzeile des Arbeitsblatts.|
||[centerHeader](/javascript/api/excel/excel.headerfooter#centerheader)|Die zentriert Kopfzeile des Arbeitsblatts.|
||[leftFooter](/javascript/api/excel/excel.headerfooter#leftfooter)|Die linke Fußzeile des Arbeitsblatts.|
||[leftHeader](/javascript/api/excel/excel.headerfooter#leftheader)|Die linke Kopfzeile des Arbeitsblatts.|
||[rightFooter](/javascript/api/excel/excel.headerfooter#rightfooter)|Die rechte Fußzeile des Arbeitsblatts.|
||[rightHeader](/javascript/api/excel/excel.headerfooter#rightheader)|Die rechte Kopfzeile des Arbeitsblatts.|
|[HeaderFooterGroup](/javascript/api/excel/excel.headerfootergroup)|[defaultForAllPages](/javascript/api/excel/excel.headerfootergroup#defaultforallpages)|Die allgemeine Kopf-/Fußzeile, die für alle Seiten verwendet wird, sofern nicht gerade/ungerade oder erste Seite angegeben ist.|
||[evenPages](/javascript/api/excel/excel.headerfootergroup#evenpages)|Die für gerade Seiten zu verwendende Kopf-/Fußzeile, für ungerade Seiten muss die ungerade Kopf-/Fußzeile angegeben werden.|
||[firstPage](/javascript/api/excel/excel.headerfootergroup#firstpage)|Die Kopf-/Fußzeile der ersten Seite, für alle anderen Seiten werden die Einstellungen für die allgemeine oder gerade/ungerade Kopf-/Fußzeile verwendet.|
||[oddPages](/javascript/api/excel/excel.headerfootergroup#oddpages)|Die für ungerade Seiten zu verwendende Kopf-/Fußzeile, für gerade Seiten muss die gerade Kopf-/Fußzeile angegeben werden.|
||[state](/javascript/api/excel/excel.headerfootergroup#state)|Der Status, nach dem Kopf- und Fußzeilen festgelegt werden.|
||[useSheetMargins](/javascript/api/excel/excel.headerfootergroup#usesheetmargins)|Ruft eine Kennzeichnung ab, ob die Kopf-/Fußzeilen an den Seitenrändern in den Optionen für das Seitenlayout des Arbeitsblatts ausgerichtet sind, oder legt diese fest.|
||[useSheetScale](/javascript/api/excel/excel.headerfootergroup#usesheetscale)|Ruft eine Kennzeichnung ab, die angibt, ob die Kopf-/Fußzeilen mit dem Prozentsatz für die Skalierung der Seite skaliert werden sollen, die in den Optionen für das Seitenlayout des Arbeitsblatts festgelegt ist, oder legt diese fest.|
|[Image](/javascript/api/excel/excel.image)|[format](/javascript/api/excel/excel.image#format)|Gibt das Format des Bilds zurück.|
||[id](/javascript/api/excel/excel.image#id)|Gibt die Shape-ID für das Bildobjekt an.|
||[shape](/javascript/api/excel/excel.image#shape)|Gibt das `Shape` dem Bild zugeordnete Objekt zurück.|
|[IterativeCalculation](/javascript/api/excel/excel.iterativecalculation)|[enabled](/javascript/api/excel/excel.iterativecalculation#enabled)|True, wenn Zirkelbezüge in Excel durch Iteration aufgelöst werden.|
||[maxChange](/javascript/api/excel/excel.iterativecalculation#maxchange)|Gibt den maximalen Änderungsbetrag zwischen den einzelnen Iterationen an, wenn Excel Zirkelverweise aufhebt.|
||[maxIteration](/javascript/api/excel/excel.iterativecalculation#maxiteration)|Gibt die maximale Anzahl von Iterationen an, die Excel zum Auflösen eines Zirkelverweises verwenden kann.|
|[Line](/javascript/api/excel/excel.line)|[beginArrowheadLength](/javascript/api/excel/excel.line#beginarrowheadlength)|Stellt die Länge der Pfeilspitze am Anfang der angegebenen Linie dar.|
||[beginArrowheadStyle](/javascript/api/excel/excel.line#beginarrowheadstyle)|Stellt das Format der Pfeilspitze am Anfang der angegebenen Linie dar.|
||[beginArrowheadWidth](/javascript/api/excel/excel.line#beginarrowheadwidth)|Stellt die Breite der Pfeilspitze am Anfang der angegebenen Linie dar.|
||[connectBeginShape(shape: Excel.Shape, connectionSite: number)](/javascript/api/excel/excel.line#connectbeginshape-shape--connectionsite-)|Fügt den Anfang der angegebenen Verbindung an eine angegebene Form an.|
||[connectEndShape(shape: Excel.Shape, connectionSite: number)](/javascript/api/excel/excel.line#connectendshape-shape--connectionsite-)|Fügt das Ende der angegebenen Verbindung an eine angegebene Form an.|
||[connectorType](/javascript/api/excel/excel.line#connectortype)|Stellt den Verbindertyp für die Linie dar.|
||[disconnectBeginShape()](/javascript/api/excel/excel.line#disconnectbeginshape--)|Löst den Anfang der angegebenen Verbindung von einer Form.|
||[disconnectEndShape()](/javascript/api/excel/excel.line#disconnectendshape--)|Löst das Ende der angegebenen Verbindung von einer Form.|
||[endArrowheadLength](/javascript/api/excel/excel.line#endarrowheadlength)|Stellt die Länge der Pfeilspitze am Ende der angegebenen Linie dar.|
||[endArrowheadStyle](/javascript/api/excel/excel.line#endarrowheadstyle)|Stellt das Format der Pfeilspitze am Ende der angegebenen Linie dar.|
||[endArrowheadWidth](/javascript/api/excel/excel.line#endarrowheadwidth)|Stellt die Breite der Pfeilspitze am Ende der angegebenen Linie dar.|
||[beginConnectedShape](/javascript/api/excel/excel.line#beginconnectedshape)|Stellt die Form dar, an die der Anfang der angegebenen Linie angefügt ist.|
||[beginConnectedSite](/javascript/api/excel/excel.line#beginconnectedsite)|Stellt die Verbindungsseite dar, mit der der Anfang einer Verbindung verbunden ist.|
||[endConnectedShape](/javascript/api/excel/excel.line#endconnectedshape)|Stellt die Form dar, an die das Ende der angegebenen Linie angefügt ist.|
||[endConnectedSite](/javascript/api/excel/excel.line#endconnectedsite)|Stellt die Verbindungsseite dar, mit der das Ende einer Verbindung verbunden ist.|
||[id](/javascript/api/excel/excel.line#id)|Gibt den Formbezeichner an.|
||[isBeginConnected](/javascript/api/excel/excel.line#isbeginconnected)|Gibt an, ob der Anfang der angegebenen Linie mit einer Form verbunden ist.|
||[isEndConnected](/javascript/api/excel/excel.line#isendconnected)|Gibt an, ob das Ende der angegebenen Linie mit einer Form verbunden ist.|
||[shape](/javascript/api/excel/excel.line#shape)|Gibt das `Shape` objekt zurück, das der Zeile zugeordnet ist.|
|[PageBreak](/javascript/api/excel/excel.pagebreak)|[delete()](/javascript/api/excel/excel.pagebreak#delete--)|Löscht ein Seitenumbruchobjekt.|
||[getCellAfterBreak()](/javascript/api/excel/excel.pagebreak#getcellafterbreak--)|Ruft die erste Zelle hinter dem Seitenumbruch ab.|
||[columnIndex](/javascript/api/excel/excel.pagebreak#columnindex)|Gibt den Spaltenindex für den Seitenumbruch an.|
||[rowIndex](/javascript/api/excel/excel.pagebreak#rowindex)|Gibt den Zeilenindex für den Seitenwechsel an.|
|[PageBreakCollection](/javascript/api/excel/excel.pagebreakcollection)|[add(pageBreakRange: Range \| string)](/javascript/api/excel/excel.pagebreakcollection#add-pagebreakrange-)|Fügt vor der oberen linken Zelle des angegebenen Bereichs einen Seitenumbruch hinzu.|
||[getCount()](/javascript/api/excel/excel.pagebreakcollection#getcount--)|Ruft die Anzahl der Seitenumbrüche in der Sammlung ab.|
||[getItem(index: number)](/javascript/api/excel/excel.pagebreakcollection#getitem-index-)|Ruft ein Seitenumbruchobjekt über den Index ab.|
||[items](/javascript/api/excel/excel.pagebreakcollection#items)|Ruft die geladenen untergeordneten Elemente in dieser Sammlung ab.|
||[removePageBreaks()](/javascript/api/excel/excel.pagebreakcollection#removepagebreaks--)|Setzt alle manuellen Seitenumbrüche in der Sammlung zurück.|
|[PageLayout](/javascript/api/excel/excel.pagelayout)|[blackAndWhite](/javascript/api/excel/excel.pagelayout#blackandwhite)|Die Schwarzweißdruckoption des Arbeitsblatts.|
||[bottomMargin](/javascript/api/excel/excel.pagelayout#bottommargin)|Der untere Seitenrand des Arbeitsblatts, der für das Drucken in Punkten verwendet werden soll.|
||[centerHorizontally](/javascript/api/excel/excel.pagelayout#centerhorizontally)|Das zentriert angeordnete Horizontale Flag des Arbeitsblatts.|
||[centerVertically](/javascript/api/excel/excel.pagelayout#centervertically)|Vertikales Zentrieren des Arbeitsblatts.|
||[draftMode](/javascript/api/excel/excel.pagelayout#draftmode)|Die Entwurfsmodusoption des Arbeitsblatts.|
||[firstPageNumber](/javascript/api/excel/excel.pagelayout#firstpagenumber)|Die erste seitenzahl des Arbeitsblatts, die gedruckt werden soll.|
||[footerMargin](/javascript/api/excel/excel.pagelayout#footermargin)|Der Fußzeilenrand des Arbeitsblatts in Punkt zur Verwendung beim Drucken.|
||[getPrintArea()](/javascript/api/excel/excel.pagelayout#getprintarea--)|Ruft das `RangeAreas` Objekt ab, das einen oder mehrere rechteckige Bereiche umfasst, das den Druckbereich für das Arbeitsblatt darstellt.|
||[getPrintAreaOrNullObject()](/javascript/api/excel/excel.pagelayout#getprintareaornullobject--)|Ruft das `RangeAreas` Objekt ab, das einen oder mehrere rechteckige Bereiche umfasst, das den Druckbereich für das Arbeitsblatt darstellt.|
||[getPrintTitleColumns()](/javascript/api/excel/excel.pagelayout#getprinttitlecolumns--)|Ruft das Bereichsobjekt ab, das die Titelspalten darstellt.|
||[getPrintTitleColumnsOrNullObject()](/javascript/api/excel/excel.pagelayout#getprinttitlecolumnsornullobject--)|Ruft das Bereichsobjekt ab, das die Titelspalten darstellt.|
||[getPrintTitleRows()](/javascript/api/excel/excel.pagelayout#getprinttitlerows--)|Ruft das Bereichsobjekt ab, das die Titelzeilen darstellt.|
||[getPrintTitleRowsOrNullObject()](/javascript/api/excel/excel.pagelayout#getprinttitlerowsornullobject--)|Ruft das Bereichsobjekt ab, das die Titelzeilen darstellt.|
||[headerMargin](/javascript/api/excel/excel.pagelayout#headermargin)|Der Kopfzeilenrand des Arbeitsblatts in Punkten, der beim Drucken verwendet werden kann.|
||[leftMargin](/javascript/api/excel/excel.pagelayout#leftmargin)|Der linke Rand des Arbeitsblatts in Punkten, der beim Drucken verwendet werden kann.|
||[orientation](/javascript/api/excel/excel.pagelayout#orientation)|Die Ausrichtung des Arbeitsblatts auf der Seite.|
||[paperSize](/javascript/api/excel/excel.pagelayout#papersize)|Das Papierformat des Arbeitsblatts der Seite.|
||[printComments](/javascript/api/excel/excel.pagelayout#printcomments)|Gibt an, ob die Kommentare des Arbeitsblatts beim Drucken angezeigt werden sollen.|
||[printErrors](/javascript/api/excel/excel.pagelayout#printerrors)|Die Druckfehleroption des Arbeitsblatts.|
||[printGridlines](/javascript/api/excel/excel.pagelayout#printgridlines)|Gibt an, ob die Gitternetzlinien des Arbeitsblatts gedruckt werden.|
||[printHeadings](/javascript/api/excel/excel.pagelayout#printheadings)|Gibt an, ob die Überschriften des Arbeitsblatts gedruckt werden.|
||[printOrder](/javascript/api/excel/excel.pagelayout#printorder)|Die Seitendruckreihenfolge des Arbeitsblatts.|
||[headersFooters](/javascript/api/excel/excel.pagelayout#headersfooters)|Kopf- und Fußzeilenkonfiguration für das Arbeitsblatt.|
||[rightMargin](/javascript/api/excel/excel.pagelayout#rightmargin)|Der rechte Rand des Arbeitsblatts in Punkt, der beim Drucken verwendet werden kann.|
||[setPrintArea(printArea: Range \| RangeAreas \| string)](/javascript/api/excel/excel.pagelayout#setprintarea-printarea-)|Legt den Druckbereich des Arbeitsblatts fest.|
||[setPrintMargins(unit: Excel.PrintMarginUnit, marginOptions: Excel.PageLayoutMarginOptions)](/javascript/api/excel/excel.pagelayout#setprintmargins-unit--marginoptions-)|Legt die Seitenränder des Arbeitsblatts mit Einheiten fest.|
||[setPrintTitleColumns(printTitleColumns: Range \| string)](/javascript/api/excel/excel.pagelayout#setprinttitlecolumns-printtitlecolumns-)|Legt die Spalten fest, die die links auf jeder Seite des Arbeitsblatts im Druck zu wiederholenden Zellen enthalten.|
||[setPrintTitleRows(printTitleRows: Range \| string)](/javascript/api/excel/excel.pagelayout#setprinttitlerows-printtitlerows-)|Legt die Zeilen fest, die die oben auf jeder Seite des Arbeitsblatts im Druck zu wiederholenden Zellen enthalten.|
||[topMargin](/javascript/api/excel/excel.pagelayout#topmargin)|Der obere Rand des Arbeitsblatts in Punkten, der beim Drucken verwendet werden kann.|
||[zoom](/javascript/api/excel/excel.pagelayout#zoom)|Die Druckzoomoptionen des Arbeitsblatts.|
|[PageLayoutMarginOptions](/javascript/api/excel/excel.pagelayoutmarginoptions)|[bottom](/javascript/api/excel/excel.pagelayoutmarginoptions#bottom)|Gibt den unteren Seitenrand des Seitenlayouts in der einheit an, die zum Drucken verwendet werden soll.|
||[footer](/javascript/api/excel/excel.pagelayoutmarginoptions#footer)|Gibt den Seitenlayoutfußrand in der zum Drucken angegebenen Einheit an.|
||[header](/javascript/api/excel/excel.pagelayoutmarginoptions#header)|Gibt den Seitenlayoutheaderrand in der einheit an, die zum Drucken verwendet werden soll.|
||[left](/javascript/api/excel/excel.pagelayoutmarginoptions#left)|Gibt den linken Seitenrand in der zum Drucken angegebenen Einheit an.|
||[right](/javascript/api/excel/excel.pagelayoutmarginoptions#right)|Gibt den rechten Seitenrand des Seitenlayouts in der einheit an, die zum Drucken verwendet werden soll.|
||[top](/javascript/api/excel/excel.pagelayoutmarginoptions#top)|Gibt den oberen Seitenrand des Seitenlayouts in der einheit an, die zum Drucken verwendet werden soll.|
|[PageLayoutZoomOptions](/javascript/api/excel/excel.pagelayoutzoomoptions)|[horizontalFitToPages](/javascript/api/excel/excel.pagelayoutzoomoptions#horizontalfittopages)|Anzahl der horizontal einzupassenden Seiten.|
||[scale](/javascript/api/excel/excel.pagelayoutzoomoptions#scale)|Der Skalierungswert für die Druckseite kann zwischen 10 und 400 liegen.|
||[verticalFitToPages](/javascript/api/excel/excel.pagelayoutzoomoptions#verticalfittopages)|Anzahl der vertikal einzupassenden Seiten.|
|[PivotField](/javascript/api/excel/excel.pivotfield)|[sortByValues(sortBy: Excel.SortBy, valuesHierarchy: Excel.DataPivotHierarchy, pivotItemScope?: Array<PivotItem \| string>)](/javascript/api/excel/excel.pivotfield#sortbyvalues-sortby--valueshierarchy--pivotitemscope-)|Sortiert das PivotField in einem bestimmten Bereich nach den angegebenen Werten.|
|[PivotLayout](/javascript/api/excel/excel.pivotlayout)|[autoFormat](/javascript/api/excel/excel.pivotlayout#autoformat)|Gibt an, ob die Formatierung automatisch formatiert wird, wenn sie aktualisiert wird oder wenn Felder verschoben werden.|
||[getDataHierarchy(cell: Range \| string)](/javascript/api/excel/excel.pivotlayout#getdatahierarchy-cell-)|Ruft die DataHierarchy ab, die zum Berechnen des Werts in einem angegebenen Bereich innerhalb der PivotTable verwendet wird.|
||[getPivotItems(axis: Excel.PivotAxis, cell: Range \| string)](/javascript/api/excel/excel.pivotlayout#getpivotitems-axis--cell-)|Ruft die PivotItems von einer Achse ab, die den Wert in einem angegebenen Bereich innerhalb der PivotTable bilden.|
||[preserveFormatting](/javascript/api/excel/excel.pivotlayout#preserveformatting)|Gibt an, ob die Formatierung beibehalten wird, wenn der Bericht durch Vorgänge wie das Pivotieren, Sortieren oder Ändern von Seitenfeldelementen aktualisiert oder neu berechnet wird.|
||[setAutoSortOnCell(cell: Range \| string, sortBy: Excel.SortBy)](/javascript/api/excel/excel.pivotlayout#setautosortoncell-cell--sortby-)|Legt fest, dass die PivotTable automatisch nach der angegebenen Zelle sortiert, um automatisch alle notwendigen Kriterien und den Kontext auszuwählen.|
|[PivotTable](/javascript/api/excel/excel.pivottable)|[enableDataValueEditing](/javascript/api/excel/excel.pivottable#enabledatavalueediting)|Gibt an, ob die PivotTable die Bearbeitung von Werten im Datentext durch den Benutzer zulässt.|
||[useCustomSortLists](/javascript/api/excel/excel.pivottable#usecustomsortlists)|Gibt an, ob die PivotTable beim Sortieren benutzerdefinierte Listen verwendet.|
|[Range](/javascript/api/excel/excel.range)|[autoFill(destinationRange?: \| Bereichszeichenfolge, autoFillType?: Excel.AutoFillType)](/javascript/api/excel/excel.range#autofill-destinationrange--autofilltype-)|Füllt den Bereich vom aktuellen Bereich bis zum Zielbereich mithilfe der angegebenen AutoFill-Logik.|
||[convertDataTypeToText()](/javascript/api/excel/excel.range#convertdatatypetotext--)|Konvertiert die Bereichszellen mit Datentypen in Text.|
||[convertToLinkedDataType(serviceID: number, languageCulture: string)](/javascript/api/excel/excel.range#converttolinkeddatatype-serviceid--languageculture-)|Konvertiert die Bereichszellen in verknüpfte Datentypen im Arbeitsblatt.|
||[copyFrom(sourceRange: Range \| RangeAreas \| string, copyType?: Excel.RangeCopyType, skipBlanks?: boolean, transpose?: boolean)](/javascript/api/excel/excel.range#copyfrom-sourcerange--copytype--skipblanks--transpose-)|Kopiert Zelldaten oder Formatierungen aus dem Quellbereich oder `RangeAreas` in den aktuellen Bereich.|
||[find(text: string, criteria: Excel.SearchCriteria)](/javascript/api/excel/excel.range#find-text--criteria-)|Sucht die angegebene Zeichenfolge anhand der angegebenen Kriterien.|
||[findOrNullObject(text: string, criteria: Excel.SearchCriteria)](/javascript/api/excel/excel.range#findornullobject-text--criteria-)|Sucht die angegebene Zeichenfolge anhand der angegebenen Kriterien.|
||[flashFill()](/javascript/api/excel/excel.range#flashfill--)|Führt eine Blitzfüllung für den aktuellen Bereich aus.|
||[getCellProperties(cellPropertiesLoadOptions: CellPropertiesLoadOptions)](/javascript/api/excel/excel.range#getcellproperties-cellpropertiesloadoptions-)|Gibt ein 2D-Array zurück, das die Daten für die Schriftart, die Füllung, den Rahmen, die Ausrichtung und andere Eigenschaften jeder Zelle kapselt.|
||[getColumnProperties(columnPropertiesLoadOptions: ColumnPropertiesLoadOptions)](/javascript/api/excel/excel.range#getcolumnproperties-columnpropertiesloadoptions-)|Gibt ein eindimensionales Array zurück, das die Daten für die Schriftart, die Füllung, den Rahmen, die Ausrichtung und andere Eigenschaften jeder Spalte kapselt.|
||[getRowProperties(rowPropertiesLoadOptions: RowPropertiesLoadOptions)](/javascript/api/excel/excel.range#getrowproperties-rowpropertiesloadoptions-)|Gibt ein eindimensionales Array zurück, das die Daten für die Schriftart, die Füllung, den Rahmen, die Ausrichtung und andere Eigenschaften jeder Zeile kapselt.|
||[getSpecialCells(cellType: Excel.SpecialCellType, cellValueType?: Excel.SpecialCellValueType)](/javascript/api/excel/excel.range#getspecialcells-celltype--cellvaluetype-)|Ruft das Objekt ab, das einen oder mehrere rechteckige Bereiche umfasst, das alle Zellen darstellt, die mit dem angegebenen Typ und `RangeAreas` Wert übereinstimmen.|
||[getSpecialCellsOrNullObject(cellType: Excel.SpecialCellType, cellValueType?: Excel.SpecialCellValueType)](/javascript/api/excel/excel.range#getspecialcellsornullobject-celltype--cellvaluetype-)|Ruft das Objekt ab, das einen oder mehrere Bereiche umfasst, das alle Zellen darstellt, die mit dem angegebenen Typ und `RangeAreas` Wert übereinstimmen.|
||[getTables(fullyContained?: boolean)](/javascript/api/excel/excel.range#gettables-fullycontained-)|Ruft eine bereichsbezogene Sammlung von Tabellen ab, die sich mit dem Bereich überschneidet.|
||[linkedDataTypeState](/javascript/api/excel/excel.range#linkeddatatypestate)|Stellt den Datentypstatus der einzelnen Zellen dar.|
||[removeDuplicates(columns: number[], includesHeader: boolean)](/javascript/api/excel/excel.range#removeduplicates-columns--includesheader-)|Entfernt doppelte Werte aus dem durch die Spalten angegebenen Bereich.|
||[replaceAll(text: string, replacement: string, criteria: Excel.ReplaceCriteria)](/javascript/api/excel/excel.range#replaceall-text--replacement--criteria-)|Sucht und ersetzt die angegebene Zeichenfolge auf der Grundlage der im aktuellen Bereich angegebenen Kriterien.|
||[setCellProperties(cellPropertiesData: SettableCellProperties[][])](/javascript/api/excel/excel.range#setcellproperties-cellpropertiesdata-)|Aktualisiert den Bereich basierend auf einem 2D-Array von Zelleigenschaften und kapselt Dinge wie Schriftart, Füllung, Rahmen und Ausrichtung ein.|
||[setColumnProperties(columnPropertiesData: SettableColumnProperties[])](/javascript/api/excel/excel.range#setcolumnproperties-columnpropertiesdata-)|Aktualisiert den Bereich basierend auf einem eindimensionalen Array von Spalteneigenschaften und kapselt Elemente wie Schriftart, Füllung, Rahmen und Ausrichtung ein.|
||[setDirty()](/javascript/api/excel/excel.range#setdirty--)|Legt für einen Bereich Neuberechnung bei der nächsten auszuführenden Neuberechnung fest.|
||[setRowProperties(rowPropertiesData: SettableRowProperties[])](/javascript/api/excel/excel.range#setrowproperties-rowpropertiesdata-)|Aktualisiert den Bereich basierend auf einem eindimensionalen Array von Zeileneigenschaften und kapselt Dinge wie Schriftart, Füllung, Rahmen und Ausrichtung ein.|
|[RangeAreas](/javascript/api/excel/excel.rangeareas)|[calculate()](/javascript/api/excel/excel.rangeareas#calculate--)|Berechnet alle Zellen in `RangeAreas` der .|
||[clear(applyTo?: Excel.ClearApplyTo)](/javascript/api/excel/excel.rangeareas#clear-applyto-)|Entfernt Werte, Format, Füllung, Rahmen und andere Eigenschaften für jeden Bereich, aus dem dieses Objekt `RangeAreas` besteht.|
||[convertDataTypeToText()](/javascript/api/excel/excel.rangeareas#convertdatatypetotext--)|Konvertiert alle Zellen in der `RangeAreas` mit Datentypen in Text.|
||[convertToLinkedDataType(serviceID: number, languageCulture: string)](/javascript/api/excel/excel.rangeareas#converttolinkeddatatype-serviceid--languageculture-)|Konvertiert alle Zellen in den `RangeAreas` verknüpften Datentypen.|
||[copyFrom(sourceRange: Range \| RangeAreas \| string, copyType?: Excel.RangeCopyType, skipBlanks?: boolean, transpose?: boolean)](/javascript/api/excel/excel.rangeareas#copyfrom-sourcerange--copytype--skipblanks--transpose-)|Kopiert Zellendaten oder Formatierungen aus dem Quellbereich oder `RangeAreas` in die aktuelle `RangeAreas` .|
||[getEntireColumn()](/javascript/api/excel/excel.rangeareas#getentirecolumn--)|Gibt ein Objekt zurück, das die gesamten Spalten des darstellt (wenn der Aktuelle z. B. die Zellen `RangeAreas` `RangeAreas` `RangeAreas` "B4:E11, H2" darstellt, wird ein zurückgegeben, das die Spalten `RangeAreas` "B:E, H:H") darstellt.|
||[getEntireRow()](/javascript/api/excel/excel.rangeareas#getentirerow--)|Gibt ein Objekt zurück, das die gesamten Zeilen des darstellt (wenn der Aktuelle z. B. die Zellen `RangeAreas` `RangeAreas` "B4:E11" darstellt, wird ein zurückgegeben, das die Zeilen `RangeAreas` `RangeAreas` "4:11") darstellt.|
||[getIntersection(anotherRange: Range \| RangeAreas \| string)](/javascript/api/excel/excel.rangeareas#getintersection-anotherrange-)|Gibt das `RangeAreas` Objekt zurück, das die Schnittmenge der angegebenen Bereiche oder `RangeAreas` darstellt.|
||[getIntersectionOrNullObject(anotherRange: Range \| RangeAreas \| string)](/javascript/api/excel/excel.rangeareas#getintersectionornullobject-anotherrange-)|Gibt das `RangeAreas` Objekt zurück, das die Schnittmenge der angegebenen Bereiche oder `RangeAreas` darstellt.|
||[getOffsetRangeAreas(rowOffset: number, columnOffset: number)](/javascript/api/excel/excel.rangeareas#getoffsetrangeareas-rowoffset--columnoffset-)|Gibt ein `RangeAreas` Objekt zurück, das um den bestimmten Zeilen- und Spaltenversatz verschoben wird.|
||[getSpecialCells(cellType: Excel.SpecialCellType, cellValueType?: Excel.SpecialCellValueType)](/javascript/api/excel/excel.rangeareas#getspecialcells-celltype--cellvaluetype-)|Gibt ein `RangeAreas` Objekt zurück, das alle Zellen darstellt, die mit dem angegebenen Typ und Wert übereinstimmen.|
||[getSpecialCellsOrNullObject(cellType: Excel.SpecialCellType, cellValueType?: Excel.SpecialCellValueType)](/javascript/api/excel/excel.rangeareas#getspecialcellsornullobject-celltype--cellvaluetype-)|Gibt ein `RangeAreas` Objekt zurück, das alle Zellen darstellt, die mit dem angegebenen Typ und Wert übereinstimmen.|
||[getTables(fullyContained?: boolean)](/javascript/api/excel/excel.rangeareas#gettables-fullycontained-)|Gibt eine Bereichssammlung von Tabellen zurück, die mit einem beliebigen Bereich in diesem Objekt `RangeAreas` überlappen.|
||[getUsedRangeAreas(valuesOnly?: boolean)](/javascript/api/excel/excel.rangeareas#getusedrangeareas-valuesonly-)|Gibt den `RangeAreas` verwendeten Wert zurück, der alle verwendeten Bereiche einzelner rechteckiger Bereiche im Objekt `RangeAreas` umfasst.|
||[getUsedRangeAreasOrNullObject(valuesOnly?: boolean)](/javascript/api/excel/excel.rangeareas#getusedrangeareasornullobject-valuesonly-)|Gibt den `RangeAreas` verwendeten Wert zurück, der alle verwendeten Bereiche einzelner rechteckiger Bereiche im Objekt `RangeAreas` umfasst.|
||[address](/javascript/api/excel/excel.rangeareas#address)|Gibt den `RangeAreas` Verweis im A1-Format zurück.|
||[addressLocal](/javascript/api/excel/excel.rangeareas#addresslocal)|Gibt den `RangeAreas` Verweis im Benutzer-Locale zurück.|
||[areaCount](/javascript/api/excel/excel.rangeareas#areacount)|Gibt die Anzahl rechteckiger Bereiche zurück, aus denen dieses Objekt `RangeAreas` besteht.|
||[areas](/javascript/api/excel/excel.rangeareas#areas)|Gibt eine Auflistung rechteckiger Bereiche zurück, die dieses Objekt `RangeAreas` umfassen.|
||[cellCount](/javascript/api/excel/excel.rangeareas#cellcount)|Gibt die Anzahl der Zellen im Objekt zurück und summiert die Zellanzahl aller einzelnen `RangeAreas` rechteckigen Bereiche.|
||[conditionalFormats](/javascript/api/excel/excel.rangeareas#conditionalformats)|Gibt eine Auflistung bedingter Formate zurück, die sich mit beliebigen Zellen in diesem Objekt `RangeAreas` überschneiden.|
||[dataValidation](/javascript/api/excel/excel.rangeareas#datavalidation)|Gibt ein Datenüberprüfungsobjekt für alle Bereiche in der `RangeAreas` zurück.|
||[format](/javascript/api/excel/excel.rangeareas#format)|Gibt ein Objekt zurück, das die Schriftart, die Füllung, die Rahmen, die Ausrichtung und andere Eigenschaften für alle Bereiche `RangeFormat` im Objekt `RangeAreas` kapselt.|
||[isEntireColumn](/javascript/api/excel/excel.rangeareas#isentirecolumn)|Gibt an, ob alle Bereiche dieses Objekts ganze Spalten `RangeAreas` darstellen (z. B. "A:C, Q:Z").|
||[isEntireRow](/javascript/api/excel/excel.rangeareas#isentirerow)|Gibt an, ob alle Bereiche dieses Objekts ganze Zeilen `RangeAreas` darstellen (z. B. "1:3, 5:7").|
||[worksheet](/javascript/api/excel/excel.rangeareas#worksheet)|Gibt das Arbeitsblatt für das aktuelle `RangeAreas` zurück.|
||[setDirty()](/javascript/api/excel/excel.rangeareas#setdirty--)|Legt den `RangeAreas` fest, der neu berechnet werden soll, wenn die nächste Neuberechnung erfolgt.|
||[style](/javascript/api/excel/excel.rangeareas#style)|Stellt die Formatvorlage für alle Bereiche in diesem Objekt `RangeAreas` dar.|
|[RangeBorder](/javascript/api/excel/excel.rangeborder)|[tintAndShade](/javascript/api/excel/excel.rangeborder#tintandshade)|Gibt ein Double an, das eine Farbe für den Bereichsrand aufhellt oder dunkler macht, der Wert liegt zwischen -1 (dunkelster) und 1 (hellster Wert), mit 0 für die Ursprüngliche Farbe.|
|[RangeBorderCollection](/javascript/api/excel/excel.rangebordercollection)|[tintAndShade](/javascript/api/excel/excel.rangebordercollection#tintandshade)|Gibt ein Double an, das eine Farbe für Bereichsränder hellt oder abdunkriert.|
|[RangeCollection](/javascript/api/excel/excel.rangecollection)|[getCount()](/javascript/api/excel/excel.rangecollection#getcount--)|Gibt die Anzahl der Bereiche in `RangeCollection` zurück.|
||[getItemAt(index: number)](/javascript/api/excel/excel.rangecollection#getitemat-index-)|Gibt das Range-Objekt basierend auf seiner Position in `RangeCollection` zurück.|
||[items](/javascript/api/excel/excel.rangecollection#items)|Ruft die geladenen untergeordneten Elemente in dieser Sammlung ab.|
|[RangeFill](/javascript/api/excel/excel.rangefill)|[pattern](/javascript/api/excel/excel.rangefill#pattern)|Das Muster eines Bereichs.|
||[patternColor](/javascript/api/excel/excel.rangefill#patterncolor)|Der HTML-Farbcode, der die Farbe des Bereichsmusters in der Form #RRGGBB (z. B. "FFA500") oder als benannte #A0 (z. B. "Orange") darstellt.|
||[patternTintAndShade](/javascript/api/excel/excel.rangefill#patterntintandshade)|Gibt ein Double an, das eine Musterfarbe für die Bereichsfüllung aufleuchtet oder abdunkriert.|
||[tintAndShade](/javascript/api/excel/excel.rangefill#tintandshade)|Gibt einen Double-Wert an, der eine Farbe für die Bereichsfüllung hellt oder abdunkriert.|
|[RangeFont](/javascript/api/excel/excel.rangefont)|[strikethrough](/javascript/api/excel/excel.rangefont#strikethrough)|Gibt den Durchschlagsstatus der Schriftart an.|
||[subscript](/javascript/api/excel/excel.rangefont#subscript)|Gibt den Subscriptstatus der Schriftart an.|
||[superscript](/javascript/api/excel/excel.rangefont#superscript)|Gibt den Hochgestellten status der Schriftart an.|
||[tintAndShade](/javascript/api/excel/excel.rangefont#tintandshade)|Gibt ein Double an, das eine Farbe für die Bereichsschriftart hellt oder abdunkriert.|
|[RangeFormat](/javascript/api/excel/excel.rangeformat)|[autoIndent](/javascript/api/excel/excel.rangeformat#autoindent)|Gibt an, ob Text automatisch eingezogen wird, wenn die Textausrichtung auf die gleiche Verteilung festgelegt ist.|
||[indentLevel](/javascript/api/excel/excel.rangeformat#indentlevel)|Eine ganze Zahl zwischen 0 und 250, die die Einzugsebene angibt.|
||[readingOrder](/javascript/api/excel/excel.rangeformat#readingorder)|Die Leserichtung für den Bereich.|
||[shrinkToFit](/javascript/api/excel/excel.rangeformat#shrinktofit)|Gibt an, ob Text automatisch verkleinert wird, um in die verfügbare Spaltenbreite zu passen.|
|[RemoveDuplicatesResult](/javascript/api/excel/excel.removeduplicatesresult)|[removed](/javascript/api/excel/excel.removeduplicatesresult#removed)|Die Anzahl der vom Vorgang entfernten doppelten Zeilen.|
||[uniqueRemaining](/javascript/api/excel/excel.removeduplicatesresult#uniqueremaining)|Die Anzahl der verbleibenden eindeutigen Zeilen im Ergebnisbereich.|
|[ReplaceCriteria](/javascript/api/excel/excel.replacecriteria)|[completeMatch](/javascript/api/excel/excel.replacecriteria#completematch)|Gibt an, ob die Übereinstimmung vollständig oder teilweise sein muss.|
||[matchCase](/javascript/api/excel/excel.replacecriteria#matchcase)|Gibt an, ob bei der Übereinstimmung die Zwischenfalle beachtet wird.|
|[RowProperties](/javascript/api/excel/excel.rowproperties)|[address](/javascript/api/excel/excel.rowproperties#address)|Stellt die `address` Eigenschaft dar.|
||[addressLocal](/javascript/api/excel/excel.rowproperties#addresslocal)|Stellt die `addressLocal` Eigenschaft dar.|
||[rowIndex](/javascript/api/excel/excel.rowproperties#rowindex)|Stellt die `rowIndex` Eigenschaft dar.|
|[SearchCriteria](/javascript/api/excel/excel.searchcriteria)|[completeMatch](/javascript/api/excel/excel.searchcriteria#completematch)|Gibt an, ob die Übereinstimmung vollständig oder teilweise sein muss.|
||[matchCase](/javascript/api/excel/excel.searchcriteria#matchcase)|Gibt an, ob bei der Übereinstimmung die Zwischenfalle beachtet wird.|
||[searchDirection](/javascript/api/excel/excel.searchcriteria#searchdirection)|Gibt die Suchrichtung an.|
|[SettableCellProperties](/javascript/api/excel/excel.settablecellproperties)|[format](/javascript/api/excel/excel.settablecellproperties#format)|Stellt die `format` Eigenschaft dar.|
||[hyperlink](/javascript/api/excel/excel.settablecellproperties#hyperlink)|Stellt die `hyperlink` Eigenschaft dar.|
||[style](/javascript/api/excel/excel.settablecellproperties#style)|Stellt die `style` Eigenschaft dar.|
|[SettableColumnProperties](/javascript/api/excel/excel.settablecolumnproperties)|[columnHidden](/javascript/api/excel/excel.settablecolumnproperties#columnhidden)|Stellt die `columnHidden` Eigenschaft dar.|
||[columnWidth](/javascript/api/excel/excel.settablecolumnproperties#columnwidth)||
||[format: Excel.CellPropertiesFormat & { columnWidth?](/javascript/api/excel/excel.settablecolumnproperties#format)|Stellt die `format` Eigenschaft dar.|
|[SettableRowProperties](/javascript/api/excel/excel.settablerowproperties)|[format: Excel.CellPropertiesFormat & { rowHeight?](/javascript/api/excel/excel.settablerowproperties#format)|Stellt die `format` Eigenschaft dar.|
||[rowHeight](/javascript/api/excel/excel.settablerowproperties#rowheight)||
||[rowHidden](/javascript/api/excel/excel.settablerowproperties#rowhidden)|Stellt die `rowHidden` Eigenschaft dar.|
|[Shape](/javascript/api/excel/excel.shape)|[altTextDescription](/javascript/api/excel/excel.shape#alttextdescription)|Gibt den alternativen Beschreibungstext für ein Objekt `Shape` an.|
||[altTextTitle](/javascript/api/excel/excel.shape#alttexttitle)|Gibt den alternativen Titeltext für ein Objekt `Shape` an.|
||[delete()](/javascript/api/excel/excel.shape#delete--)|Entfernt die Form aus dem Arbeitsblatt.|
||[geometricShapeType](/javascript/api/excel/excel.shape#geometricshapetype)|Gibt den geometrischen Formtyp dieser geometrischen Form an.|
||[getAsImage(format: Excel.PictureFormat)](/javascript/api/excel/excel.shape#getasimage-format-)|Konvertiert die Form in ein Bild und gibt das Bild als base64-codierte Zeichenfolge zurück.|
||[height](/javascript/api/excel/excel.shape#height)|Gibt die Höhe der Form in Punkt an.|
||[incrementLeft(increment: number)](/javascript/api/excel/excel.shape#incrementleft-increment-)|Verschiebt die Form horizontal um die angegebene Punktanzahl.|
||[incrementRotation(increment: number)](/javascript/api/excel/excel.shape#incrementrotation-increment-)|Dreht die Form um die angegebene Gradzahl um die Z-Achse.|
||[incrementTop(increment: number)](/javascript/api/excel/excel.shape#incrementtop-increment-)|Verschiebt die Form vertikal um die angegebene Anzahl von Punkten.|
||[left](/javascript/api/excel/excel.shape#left)|Der Abstand in Punkten von der linken Seite der Form zur linken Seite des Arbeitsblatts.|
||[lockAspectRatio](/javascript/api/excel/excel.shape#lockaspectratio)|Gibt an, ob das Seitenverhältnis dieser Form gesperrt ist.|
||[name](/javascript/api/excel/excel.shape#name)|Gibt den Namen der Form an.|
||[connectionSiteCount](/javascript/api/excel/excel.shape#connectionsitecount)|Gibt die Anzahl der Verbindungsseiten für diese Form zurück.|
||[fill](/javascript/api/excel/excel.shape#fill)|Gibt die Füllungsformatierung dieser Form zurück.|
||[geometricShape](/javascript/api/excel/excel.shape#geometricshape)|Gibt die der Form zugeordnete geometrische Form zurück.|
||[group](/javascript/api/excel/excel.shape#group)|Gibt die der Form zugeordnete Formgruppe zurück.|
||[id](/javascript/api/excel/excel.shape#id)|Gibt den Formbezeichner an.|
||[image](/javascript/api/excel/excel.shape#image)|Gibt das Bild zurück, das der Form zugeordnet ist.|
||[level](/javascript/api/excel/excel.shape#level)|Gibt die Ebene der angegebenen Form an.|
||[line](/javascript/api/excel/excel.shape#line)|Gibt die Linie zurück, die der Form zugeordnet ist.|
||[lineFormat](/javascript/api/excel/excel.shape#lineformat)|Gibt die Linienformatierung dieser Form zurück.|
||[onActivated](/javascript/api/excel/excel.shape#onactivated)|Tritt ein, wenn die Form aktiviert wird.|
||[onDeactivated](/javascript/api/excel/excel.shape#ondeactivated)|Tritt ein, wenn die Form deaktiviert wird.|
||[parentGroup](/javascript/api/excel/excel.shape#parentgroup)|Gibt die übergeordnete Gruppe dieser Form an.|
||[textFrame](/javascript/api/excel/excel.shape#textframe)|Gibt das textFrame-Objekt dieser Form zurück.|
||[type](/javascript/api/excel/excel.shape#type)|Gibt den Typ dieser Form zurück.|
||[zOrderPosition](/javascript/api/excel/excel.shape#zorderposition)|Gibt die Position der angegebenen Form in der Z-Reihenfolge an, wobei 0 den Boden des Reihenfolgestapels darstellt.|
||[rotation](/javascript/api/excel/excel.shape#rotation)|Gibt die Drehung der Form in Grad an.|
||[scaleHeight(scaleFactor: number, scaleType: Excel.ShapeScaleType, scaleFrom?: Excel.ShapeScaleFrom)](/javascript/api/excel/excel.shape#scaleheight-scalefactor--scaletype--scalefrom-)|Skaliert die Höhe der Form anhand eines angegebenen Faktors.|
||[scaleWidth(scaleFactor: number, scaleType: Excel.ShapeScaleType, scaleFrom?: Excel.ShapeScaleFrom)](/javascript/api/excel/excel.shape#scalewidth-scalefactor--scaletype--scalefrom-)|Skaliert die Breite der Form anhand eines angegebenen Faktors.|
||[setZOrder(position: Excel.ShapeZOrder)](/javascript/api/excel/excel.shape#setzorder-position-)|Verschiebt de angegebene Form in der Z-Reihenfolge der Sammlung nach oben oder unten, wodurch sie vor oder hinter anderen Formen zu liegen kommt.|
||[top](/javascript/api/excel/excel.shape#top)|Der Abstand in Punkten zwischen der oberen Kante der Form und der oberen Kante der Arbeitsmappe.|
||[visible](/javascript/api/excel/excel.shape#visible)|Gibt an, ob die Form sichtbar ist.|
||[width](/javascript/api/excel/excel.shape#width)|Gibt die Breite der Form in Punkt an.|
|[ShapeActivatedEventArgs](/javascript/api/excel/excel.shapeactivatedeventargs)|[shapeId](/javascript/api/excel/excel.shapeactivatedeventargs#shapeid)|Ruft die ID des aktivierten Shapes ab.|
||[type](/javascript/api/excel/excel.shapeactivatedeventargs#type)|Ruft den Typ des Ereignisses ab.|
||[worksheetId](/javascript/api/excel/excel.shapeactivatedeventargs#worksheetid)|Ruft die ID des Arbeitsblatts ab, in dem das Shape aktiviert wird.|
|[ShapeCollection](/javascript/api/excel/excel.shapecollection)|[addGeometricShape(geometricShapeType: Excel.GeometricShapeType)](/javascript/api/excel/excel.shapecollection#addgeometricshape-geometricshapetype-)|Fügt dem Arbeitsblatt eine geometrische Form hinzu.|
||[addGroup(values: Array<string \| Shape>)](/javascript/api/excel/excel.shapecollection#addgroup-values-)|Gruppiert eine Teilmenge von Formen auf dem Arbeitsblatt dieser Sammlung.|
||[addImage(base64ImageString: string)](/javascript/api/excel/excel.shapecollection#addimage-base64imagestring-)|Erstellt ein Bild aus einer base64-codierten Zeichenfolge und fügt es dem Arbeitsblatt hinzu.|
||[addLine(startLeft: number, startTop: number, endLeft: number, endTop: number, connectorType?: Excel.ConnectorType)](/javascript/api/excel/excel.shapecollection#addline-startleft--starttop--endleft--endtop--connectortype-)|Fügt einem Arbeitsblatt eine Linie hinzu.|
||[addTextBox(text?: string)](/javascript/api/excel/excel.shapecollection#addtextbox-text-)|Fügt dem Arbeitsblatt ein Textfeld mit dem angegebenen Text als Inhalt hinzu.|
||[getCount()](/javascript/api/excel/excel.shapecollection#getcount--)|Gibt die Anzahl der Formen auf dem Arbeitsblatt zurück.|
||[getItem(key: string)](/javascript/api/excel/excel.shapecollection#getitem-key-)|Ruft ein Shape mit seinem Namen oder seiner ID ab.|
||[getItemAt(index: number)](/javascript/api/excel/excel.shapecollection#getitemat-index-)|Ruft eine Form anhand ihrer Position in der Sammlung ab.|
||[items](/javascript/api/excel/excel.shapecollection#items)|Ruft die geladenen untergeordneten Elemente in dieser Sammlung ab.|
|[ShapeDeactivatedEventArgs](/javascript/api/excel/excel.shapedeactivatedeventargs)|[shapeId](/javascript/api/excel/excel.shapedeactivatedeventargs#shapeid)|Ruft die ID der deaktivierten Form ab.|
||[type](/javascript/api/excel/excel.shapedeactivatedeventargs#type)|Ruft den Typ des Ereignisses ab.|
||[worksheetId](/javascript/api/excel/excel.shapedeactivatedeventargs#worksheetid)|Ruft die ID des Arbeitsblatts ab, in dem das Shape deaktiviert ist.|
|[ShapeFill](/javascript/api/excel/excel.shapefill)|[clear()](/javascript/api/excel/excel.shapefill#clear--)|Löscht die Füllungsformatierung dieser Form.|
||[foregroundColor](/javascript/api/excel/excel.shapefill#foregroundcolor)|Stellt die Vordergrundfarbe der Formfüllung im HTML-Farbformat im Format #RRGGBB (z. B. "FFA500") oder als benannte HTML-Farbe (z. B. "Orange") dar.|
||[type](/javascript/api/excel/excel.shapefill#type)|Gibt den Füllungstyp der Form zurück.|
||[setSolidColor(color: string)](/javascript/api/excel/excel.shapefill#setsolidcolor-color-)|Legt die Füllungsformatierung der Form auf einfarbige Füllung fest.|
||[transparency](/javascript/api/excel/excel.shapefill#transparency)|Gibt den Transparenzprozentsatz der Füllung als Wert zwischen 0,0 (undurchsichtig) und 1,0 (klar) an.|
|[ShapeFont](/javascript/api/excel/excel.shapefont)|[bold](/javascript/api/excel/excel.shapefont#bold)|Stellt den Fett-Status der Schriftart dar.|
||[color](/javascript/api/excel/excel.shapefont#color)|HTML-Farbcodedarstellung der Textfarbe (z. B. "#FF0000" steht für Rot).|
||[italic](/javascript/api/excel/excel.shapefont#italic)|Stellt den Kursiv-Status der Schriftart dar.|
||[name](/javascript/api/excel/excel.shapefont#name)|Stellt den Schriftartennamen dar (z. B. "Calibri").|
||[size](/javascript/api/excel/excel.shapefont#size)|Stellt die Schriftgröße in Punkt dar (z. B. 11).|
||[underline](/javascript/api/excel/excel.shapefont#underline)|Art der auf die Schriftart angewendeten Unterstreichung.|
|[ShapeGroup](/javascript/api/excel/excel.shapegroup)|[id](/javascript/api/excel/excel.shapegroup#id)|Gibt den Formbezeichner an.|
||[shape](/javascript/api/excel/excel.shapegroup#shape)|Gibt das `Shape` objekt zurück, das der Gruppe zugeordnet ist.|
||[shapes](/javascript/api/excel/excel.shapegroup#shapes)|Gibt die Auflistung von Objekten `Shape` zurück.|
||[ungroup()](/javascript/api/excel/excel.shapegroup#ungroup--)|Hebt die Gruppierung von gruppierten Formen in der angegebenen Formgruppe auf.|
|[ShapeLineFormat](/javascript/api/excel/excel.shapelineformat)|[color](/javascript/api/excel/excel.shapelineformat#color)|Stellt die Linienfarbe im HTML-Farbformat, im Format #RRGGBB (z. B. "FFA500") oder als benannte HTML-Farbe (z. B. "Orange") dar.|
||[dashStyle](/javascript/api/excel/excel.shapelineformat#dashstyle)|Steht für die Linienart der Form.|
||[style](/javascript/api/excel/excel.shapelineformat#style)|Stellt die Linienart der Form dar.|
||[transparency](/javascript/api/excel/excel.shapelineformat#transparency)|Stellt den Deckungsgrad der angegebenen Linie als Wert von 0,0 (undurchsichtig) bis 1,0 (transparent) dar.|
||[visible](/javascript/api/excel/excel.shapelineformat#visible)|Gibt an, ob die Linienformatierung eines Shape-Elements sichtbar ist.|
||[weight](/javascript/api/excel/excel.shapelineformat#weight)|Stellt die Stärke der Linie in Punkt dar.|
|[SortField](/javascript/api/excel/excel.sortfield)|[subField](/javascript/api/excel/excel.sortfield#subfield)|Gibt das Unterfeld an, das der Zieleigenschaftsname eines rich-Werts ist, nach dem sortiert werden soll.|
|[StyleCollection](/javascript/api/excel/excel.stylecollection)|[getCount()](/javascript/api/excel/excel.stylecollection#getcount--)|Ruft die Anzahl von Formatvorlagen in der Sammlung ab.|
||[getItemAt(index: number)](/javascript/api/excel/excel.stylecollection#getitemat-index-)|Ruft eine Formatvorlage anhand ihrer Position in der Sammlung ab.|
|[Table](/javascript/api/excel/excel.table)|[autoFilter](/javascript/api/excel/excel.table#autofilter)|Represents the `AutoFilter` object of the table.|
|[TableAddedEventArgs](/javascript/api/excel/excel.tableaddedeventargs)|[source](/javascript/api/excel/excel.tableaddedeventargs#source)|Ruft die Quelle des Ereignisses ab.|
||[tableId](/javascript/api/excel/excel.tableaddedeventargs#tableid)|Ruft die ID der hinzugefügten Tabelle ab.|
||[type](/javascript/api/excel/excel.tableaddedeventargs#type)|Ruft den Typ des Ereignisses ab.|
||[worksheetId](/javascript/api/excel/excel.tableaddedeventargs#worksheetid)|Ruft die ID des Arbeitsblatts ab, in dem die Tabelle hinzugefügt wird.|
|[TableChangedEventArgs](/javascript/api/excel/excel.tablechangedeventargs)|[details](/javascript/api/excel/excel.tablechangedeventargs#details)|Ruft die Informationen zum Änderungsdetail ab.|
|[TableCollection](/javascript/api/excel/excel.tablecollection)|[onAdded](/javascript/api/excel/excel.tablecollection#onadded)|Tritt auf, wenn eine neue Tabelle in einer Arbeitsmappe hinzugefügt wird.|
||[onDeleted](/javascript/api/excel/excel.tablecollection#ondeleted)|Tritt ein, wenn in einer Arbeitsmappe die angegebene Tabelle gelöscht wird.|
|[TableDeletedEventArgs](/javascript/api/excel/excel.tabledeletedeventargs)|[source](/javascript/api/excel/excel.tabledeletedeventargs#source)|Ruft die Quelle des Ereignisses ab.|
||[tableId](/javascript/api/excel/excel.tabledeletedeventargs#tableid)|Ruft die ID der gelöschten Tabelle ab.|
||[tableName](/javascript/api/excel/excel.tabledeletedeventargs#tablename)|Ruft den Namen der gelöschten Tabelle ab.|
||[type](/javascript/api/excel/excel.tabledeletedeventargs#type)|Ruft den Typ des Ereignisses ab.|
||[worksheetId](/javascript/api/excel/excel.tabledeletedeventargs#worksheetid)|Ruft die ID des Arbeitsblatts ab, in dem die Tabelle gelöscht wird.|
|[TableScopedCollection](/javascript/api/excel/excel.tablescopedcollection)|[getCount()](/javascript/api/excel/excel.tablescopedcollection#getcount--)|Ruft die Anzahl von Tabellen in der Auflistung ab.|
||[getFirst()](/javascript/api/excel/excel.tablescopedcollection#getfirst--)|Ruft die erste Tabelle in der Sammlung ab.|
||[getItem(key: string)](/javascript/api/excel/excel.tablescopedcollection#getitem-key-)|Ruft eine Tabelle anhand des Namens oder der ID ab.|
||[items](/javascript/api/excel/excel.tablescopedcollection#items)|Ruft die geladenen untergeordneten Elemente in dieser Sammlung ab.|
|[TextFrame](/javascript/api/excel/excel.textframe)|[autoSizeSetting](/javascript/api/excel/excel.textframe#autosizesetting)|Die Einstellungen für die automatische Größenanpassung für den Textrahmen.|
||[bottomMargin](/javascript/api/excel/excel.textframe#bottommargin)|Stellt den unteren Rand des Textrahmens in Punkt dar.|
||[deleteText()](/javascript/api/excel/excel.textframe#deletetext--)|Löscht den gesamten Text im Textrahmen.|
||[horizontalAlignment](/javascript/api/excel/excel.textframe#horizontalalignment)|Stellt die horizontale Ausrichtung des Textrahmens dar.|
||[horizontalOverflow](/javascript/api/excel/excel.textframe#horizontaloverflow)|Stellt das horizontale Überlaufverhalten des Textrahmens dar.|
||[leftMargin](/javascript/api/excel/excel.textframe#leftmargin)|Stellt den linken Rand des Textrahmens in Punkt dar.|
||[orientation](/javascript/api/excel/excel.textframe#orientation)|Stellt den Winkel dar, an dem der Text für den Textrahmen ausgerichtet ist.|
||[readingOrder](/javascript/api/excel/excel.textframe#readingorder)|Stellt die Lesereihenfolge des Textrahmens dar, entweder von links nach rechts oder von rechts nach links.|
||[hasText](/javascript/api/excel/excel.textframe#hastext)|Gibt an, ob der Textrahmen Text enthält.|
||[textRange](/javascript/api/excel/excel.textframe#textrange)|Stellt den Text dar, der an eine Form im Textrahmen angefügt ist, sowie Eigenschaften und Methoden für das Bearbeiten des Texts.|
||[rightMargin](/javascript/api/excel/excel.textframe#rightmargin)|Stellt den rechten Rand des Textrahmens in Punkt dar.|
||[topMargin](/javascript/api/excel/excel.textframe#topmargin)|Stellt den oberen Rand des Textrahmens in Punkt dar.|
||[verticalAlignment](/javascript/api/excel/excel.textframe#verticalalignment)|Stellt die vertikale Ausrichtung des Textrahmens dar.|
||[verticalOverflow](/javascript/api/excel/excel.textframe#verticaloverflow)|Stellt das vertikale Überlaufverhalten des Textrahmens dar.|
|[TextRange](/javascript/api/excel/excel.textrange)|[getSubstring(start: number, length?: number)](/javascript/api/excel/excel.textrange#getsubstring-start--length-)|Gibt ein TextRange-Objekt für die Teilzeichenfolge im angegebenen Bereich zurück.|
||[font](/javascript/api/excel/excel.textrange#font)|Gibt ein `ShapeFont` Objekt zurück, das die Schriftartattribute für den Textbereich darstellt.|
||[text](/javascript/api/excel/excel.textrange#text)|Stellt den unformatierten Textinhalt des Textbereichs dar.|
|[Arbeitsmappe](/javascript/api/excel/excel.workbook)|[chartDataPointTrack](/javascript/api/excel/excel.workbook#chartdatapointtrack)|True, wenn alle Diagramme in der Arbeitsmappe die tatsächlichen Datenpunkte nachverfolgen, mit denen sie verbunden sind.|
||[getActiveChart()](/javascript/api/excel/excel.workbook#getactivechart--)|Ruft das derzeit aktive Diagramm in der Arbeitsmappe ab.|
||[getActiveChartOrNullObject()](/javascript/api/excel/excel.workbook#getactivechartornullobject--)|Ruft das derzeit aktive Diagramm in der Arbeitsmappe ab.|
||[getIsActiveCollabSession()](/javascript/api/excel/excel.workbook#getisactivecollabsession--)|Gibt `true` zurück, ob die Arbeitsmappe von mehreren Benutzern bearbeitet wird (durch gemeinsamen Erstellen).|
||[getSelectedRanges()](/javascript/api/excel/excel.workbook#getselectedranges--)|Ruft die aktuell ausgewählten Bereiche (einen oder mehrere) aus der Arbeitsmappe ab.|
||[isDirty](/javascript/api/excel/excel.workbook#isdirty)|Gibt an, ob seit dem letzten Speichern der Arbeitsmappe Änderungen vorgenommen wurden.|
||[autoSave](/javascript/api/excel/excel.workbook#autosave)|Gibt an, ob sich die Arbeitsmappe im AutoSave-Modus befindet.|
||[calculationEngineVersion](/javascript/api/excel/excel.workbook#calculationengineversion)|Gibt eine Nummer zur Version des Excel-Berechnungsmoduls zurück.|
||[onAutoSaveSettingChanged](/javascript/api/excel/excel.workbook#onautosavesettingchanged)|Tritt auf, wenn die Einstellung AutoSave in der Arbeitsmappe geändert wird.|
||[previouslySaved](/javascript/api/excel/excel.workbook#previouslysaved)|Gibt an, ob die Arbeitsmappe jemals lokal oder online gespeichert wurde.|
||[usePrecisionAsDisplayed](/javascript/api/excel/excel.workbook#useprecisionasdisplayed)|True, wenn die Berechnungen in dieser Arbeitsmappe nur mit der Genauigkeit durchgeführt werden, mit der die Zahlen angezeigt werden.|
|[WorkbookAutoSaveSettingChangedEventArgs](/javascript/api/excel/excel.workbookautosavesettingchangedeventargs)|[type](/javascript/api/excel/excel.workbookautosavesettingchangedeventargs#type)|Ruft den Typ des Ereignisses ab.|
|[Arbeitsblatt](/javascript/api/excel/excel.worksheet)|[enableCalculation](/javascript/api/excel/excel.worksheet#enablecalculation)|Bestimmt, ob Excel das Arbeitsblatt bei Bedarf neu berechnen soll.|
||[findAll(text: string, criteria: Excel.WorksheetSearchCriteria)](/javascript/api/excel/excel.worksheet#findall-text--criteria-)|Sucht alle Vorkommen der angegebenen Zeichenfolge basierend auf den angegebenen Kriterien und gibt sie als Objekt zurück, das einen `RangeAreas` oder mehrere rechteckige Bereiche umfasst.|
||[findAllOrNullObject(text: string, criteria: Excel.WorksheetSearchCriteria)](/javascript/api/excel/excel.worksheet#findallornullobject-text--criteria-)|Sucht alle Vorkommen der angegebenen Zeichenfolge basierend auf den angegebenen Kriterien und gibt sie als Objekt zurück, das einen `RangeAreas` oder mehrere rechteckige Bereiche umfasst.|
||[getRanges(address?: string)](/javascript/api/excel/excel.worksheet#getranges-address-)|Ruft das `RangeAreas` Objekt ab, das einen oder mehrere Blöcke rechteckiger Bereiche darstellt, die durch die Adresse oder den Namen angegeben werden.|
||[autoFilter](/javascript/api/excel/excel.worksheet#autofilter)|Represents the `AutoFilter` object of the worksheet.|
||[horizontalPageBreaks](/javascript/api/excel/excel.worksheet#horizontalpagebreaks)|Ruft die Sammlung der horizontalen Seitenumbrüche für das Arbeitsblatt ab.|
||[onFormatChanged](/javascript/api/excel/excel.worksheet#onformatchanged)|Tritt ein, wenn das Format für ein bestimmtes Arbeitsblatt geändert wird.|
||[pageLayout](/javascript/api/excel/excel.worksheet#pagelayout)|Ruft das `PageLayout` Objekt des Arbeitsblatts ab.|
||[shapes](/javascript/api/excel/excel.worksheet#shapes)|Gibt die Sammlung aller Formobjekte auf dem Arbeitsblatt zurück.|
||[verticalPageBreaks](/javascript/api/excel/excel.worksheet#verticalpagebreaks)|Ruft die Sammlung der vertikalen Seitenumbrüche für das Arbeitsblatt ab.|
||[replaceAll(text: string, replacement: string, criteria: Excel.ReplaceCriteria)](/javascript/api/excel/excel.worksheet#replaceall-text--replacement--criteria-)|Sucht und ersetzt die angegebene Zeichenfolge auf der Grundlage der auf dem aktuellen Arbeitsblatt angegebenen Kriterien.|
|[WorksheetChangedEventArgs](/javascript/api/excel/excel.worksheetchangedeventargs)|[details](/javascript/api/excel/excel.worksheetchangedeventargs#details)|Stellt die Informationen zum Änderungsdetail dar.|
|[WorksheetCollection](/javascript/api/excel/excel.worksheetcollection)|[onChanged](/javascript/api/excel/excel.worksheetcollection#onchanged)|Tritt ein, wenn eines der Arbeitsblätter in der Arbeitsmappe geändert wird.|
||[onFormatChanged](/javascript/api/excel/excel.worksheetcollection#onformatchanged)|Tritt auf, wenn ein Arbeitsblatt in der Arbeitsmappe ein Format geändert hat.|
||[onSelectionChanged](/javascript/api/excel/excel.worksheetcollection#onselectionchanged)|Tritt ein, wenn sich die Auswahl auf einem beliebigen Arbeitsblatt ändert.|
|[WorksheetFormatChangedEventArgs](/javascript/api/excel/excel.worksheetformatchangedeventargs)|[address](/javascript/api/excel/excel.worksheetformatchangedeventargs#address)|Ruft die Bereichsadresse ab, die den geänderten Bereich eines bestimmten Arbeitsblatts darstellt.|
||[getRange(ctx: Excel.RequestContext)](/javascript/api/excel/excel.worksheetformatchangedeventargs#getrange-ctx-)|Ruft den Bereich ab, der den geänderten Bereich eines bestimmten Arbeitsblatts darstellt.|
||[getRangeOrNullObject(ctx: Excel.RequestContext)](/javascript/api/excel/excel.worksheetformatchangedeventargs#getrangeornullobject-ctx-)|Ruft den Bereich ab, der den geänderten Bereich eines bestimmten Arbeitsblatts darstellt.|
||[source](/javascript/api/excel/excel.worksheetformatchangedeventargs#source)|Ruft die Quelle des Ereignisses ab.|
||[type](/javascript/api/excel/excel.worksheetformatchangedeventargs#type)|Ruft den Typ des Ereignisses ab.|
||[worksheetId](/javascript/api/excel/excel.worksheetformatchangedeventargs#worksheetid)|Ruft die ID des Arbeitsblatts ab, in dem sich die Daten geändert haben.|
|[WorksheetSearchCriteria](/javascript/api/excel/excel.worksheetsearchcriteria)|[completeMatch](/javascript/api/excel/excel.worksheetsearchcriteria#completematch)|Gibt an, ob die Übereinstimmung vollständig oder teilweise sein muss.|
||[matchCase](/javascript/api/excel/excel.worksheetsearchcriteria#matchcase)|Gibt an, ob bei der Übereinstimmung die Zwischenfalle beachtet wird.|
