| Klasse | Felder | Beschreibung |
|:---|:---|:---|
|[Chart](/javascript/api/excel/excel.chart)|[Charttype](/javascript/api/excel/excel.chart#charttype)|Gibt den Typ des Diagramms an.|
||[id](/javascript/api/excel/excel.chart#id)|Die eindeutige ID des Diagramms.|
||[showAllFieldButtons](/javascript/api/excel/excel.chart#showallfieldbuttons)|Gibt an, ob alle Feldschaltflächen in einem PivotChart angezeigt werden sollen.|
|[ChartAreaFormat](/javascript/api/excel/excel.chartareaformat)|[Grenze](/javascript/api/excel/excel.chartareaformat#border)|Stellt das Rahmenformat des Diagrammbereichs dar, einschließlich Farbe, Linienart und Gewichtung.|
|[ChartAxes](/javascript/api/excel/excel.chartaxes)|[getItem(type: Excel. ChartAxisType, group?: Excel. ChartAxisGroup)](/javascript/api/excel/excel.chartaxes#getitem-type--group-)|Gibt eine bestimmte Achse anhand des Typs und der Gruppe zurück.|
|[ChartAxis](/javascript/api/excel/excel.chartaxis)|[baseTimeUnit](/javascript/api/excel/excel.chartaxis#basetimeunit)|Gibt die Basiseinheit für die angegebene Rubrikenachse an.|
||[Categorytype](/javascript/api/excel/excel.chartaxis#categorytype)|Gibt den Typ der Rubrikenachse an.|
||[displayUnit](/javascript/api/excel/excel.chartaxis#displayunit)|Die Anzeigeeinheit für die Achse.|
||[logBase](/javascript/api/excel/excel.chartaxis#logbase)|Gibt die Basis des Logarithmus bei Verwendung logarithmischer Skalierungen an.|
||[majorTickMark](/javascript/api/excel/excel.chartaxis#majortickmark)|Gibt den Typ der Haupt-Teilstriche für die angegebene Achse an.|
||[majorTimeUnitScale](/javascript/api/excel/excel.chartaxis#majortimeunitscale)|Gibt den Skalierungswert der Haupteinheit für die Rubrikenachse an, wenn die `categoryType` Eigenschaft auf festgelegt `dateAxis` ist.|
||[minorTickMark](/javascript/api/excel/excel.chartaxis#minortickmark)|Gibt den Typ der Hilfs-Teilstriche für die angegebene Achse an.|
||[minorTimeUnitScale](/javascript/api/excel/excel.chartaxis#minortimeunitscale)|Gibt den Wert der Hilfseinheitsskala für die Rubrikenachse an, wenn die `categoryType` Eigenschaft auf festgelegt `dateAxis` ist.|
||[axisGroup](/javascript/api/excel/excel.chartaxis#axisgroup)|Gibt die Gruppe für die angegebene Achse an.|
||[customDisplayUnit](/javascript/api/excel/excel.chartaxis#customdisplayunit)|Gibt den Wert der benutzerdefinierten Anzeigeeinheit für die Achse an.|
||[height](/javascript/api/excel/excel.chartaxis#height)|Gibt die Höhe der Diagrammachse in Punkt an.|
||[left](/javascript/api/excel/excel.chartaxis#left)|Gibt den Abstand zwischen dem linken Rand der Achse und dem linken Rand des Diagrammbereichs in Punkten an.|
||[top](/javascript/api/excel/excel.chartaxis#top)|Gibt den Abstand zwischen dem oberen Rand der Achse und dem oberen Rand des Diagrammbereichs in Punkt an.|
||[type](/javascript/api/excel/excel.chartaxis#type)|Gibt den Achsentyp an.|
||[width](/javascript/api/excel/excel.chartaxis#width)|Gibt die Breite der Diagrammachse in Punkt an.|
||[reversePlotOrder](/javascript/api/excel/excel.chartaxis#reverseplotorder)|Gibt an, ob Excel Datenpunkte vom letzten bis zum ersten gezeichnet.|
||[scaleType](/javascript/api/excel/excel.chartaxis#scaletype)|Gibt den Skalierungstyp der Größenachse an.|
||[setCategoryNames(sourceData: Range)](/javascript/api/excel/excel.chartaxis#setcategorynames-sourcedata-)|Legt alle Kategorienamen für die angegebene Achse fest.|
||[setCustomDisplayUnit(value: number)](/javascript/api/excel/excel.chartaxis#setcustomdisplayunit-value-)|Legt für die Anzeigeeinheit der Achse einen benutzerdefinierten Wert fest.|
||[showDisplayUnitLabel](/javascript/api/excel/excel.chartaxis#showdisplayunitlabel)|Gibt an, ob die Achsenanzeigeeinheitsbeschriftung sichtbar ist.|
||[tickLabelPosition](/javascript/api/excel/excel.chartaxis#ticklabelposition)|Gibt die Position der Teilstrichbeschriftungen der angegebenen Achse an.|
||[tickLabelSpacing](/javascript/api/excel/excel.chartaxis#ticklabelspacing)|Gibt die Anzahl der Kategorien oder Datenreihen zwischen Teilstrichbeschriftungen an.|
||[tickMarkSpacing](/javascript/api/excel/excel.chartaxis#tickmarkspacing)|Gibt die Anzahl der Kategorien oder Datenreihen zwischen Teilstrichen an.|
||[visible](/javascript/api/excel/excel.chartaxis#visible)|Gibt an, ob die Achse sichtbar ist.|
|[ChartBorder](/javascript/api/excel/excel.chartborder)|[color](/javascript/api/excel/excel.chartborder#color)|HTML-Farbcode, der die Farbe der Rahmen im Diagramm darstellt.|
||[Linestyle](/javascript/api/excel/excel.chartborder#linestyle)|Steht für die Linienart des Rahmens.|
||[weight](/javascript/api/excel/excel.chartborder#weight)|Stärke des Rahmens in Punkt angegeben.|
|[ChartDataLabel](/javascript/api/excel/excel.chartdatalabel)|[position](/javascript/api/excel/excel.chartdatalabel#position)|Wert, der die Position der Datenbeschriftung darstellt.|
||[Trennzeichen](/javascript/api/excel/excel.chartdatalabel#separator)|Zeichenfolge, die das Trennzeichen für die Datenbeschriftungen in einem Diagramm darstellt.|
||[showBubbleSize](/javascript/api/excel/excel.chartdatalabel#showbubblesize)|Gibt an, ob die Blasengröße der Datenbeschriftung sichtbar ist.|
||[showCategoryName](/javascript/api/excel/excel.chartdatalabel#showcategoryname)|Gibt an, ob der Kategoriename der Datenbeschriftung angezeigt wird.|
||[showLegendKey](/javascript/api/excel/excel.chartdatalabel#showlegendkey)|Gibt an, ob der Legendenschlüssel der Datenbeschriftung sichtbar ist.|
||[showPercentage](/javascript/api/excel/excel.chartdatalabel#showpercentage)|Gibt an, ob der Prozentsatz der Datenbeschriftung sichtbar ist.|
||[showSeriesName](/javascript/api/excel/excel.chartdatalabel#showseriesname)|Gibt an, ob der Datenbeschriftungsreihenname sichtbar ist.|
||[showValue](/javascript/api/excel/excel.chartdatalabel#showvalue)|Gibt an, ob der Datenbeschriftungswert sichtbar ist.|
|[ChartFormatString](/javascript/api/excel/excel.chartformatstring)|[font](/javascript/api/excel/excel.chartformatstring#font)|Stellt die Schriftartattribute dar, z. B. Schriftartname, Schriftgrad und Farbe eines Diagrammzeichenobjekts.|
|[ChartLegend](/javascript/api/excel/excel.chartlegend)|[height](/javascript/api/excel/excel.chartlegend#height)|Gibt die Höhe der Legende im Diagramm in Punkt an.|
||[left](/javascript/api/excel/excel.chartlegend#left)|Gibt den linken Wert der Legende im Diagramm in Punkt an.|
||[legendEntries](/javascript/api/excel/excel.chartlegend#legendentries)|Stellt eine Sammlung von LegendEntries in der Legende dar.|
||[showShadow](/javascript/api/excel/excel.chartlegend#showshadow)|Gibt an, ob die Legende einen Schatten auf dem Diagramm hat.|
||[top](/javascript/api/excel/excel.chartlegend#top)|Gibt den oberen Rand einer Diagrammlegende an.|
||[width](/javascript/api/excel/excel.chartlegend#width)|Gibt die Breite der Legende im Diagramm in Punkt an.|
|[ChartLegendEntry](/javascript/api/excel/excel.chartlegendentry)|[visible](/javascript/api/excel/excel.chartlegendentry#visible)|Stellt die Sichtbarkeit eines Diagrammlegendeeintrags dar.|
|[ChartLegendEntryCollection](/javascript/api/excel/excel.chartlegendentrycollection)|[getCount()](/javascript/api/excel/excel.chartlegendentrycollection#getcount--)|Gibt die Anzahl der Legendeneinträge in der Auflistung zurück.|
||[getItemAt(index: number)](/javascript/api/excel/excel.chartlegendentrycollection#getitemat-index-)|Gibt einen Legendeneintrag am angegebenen Index zurück.|
||[items](/javascript/api/excel/excel.chartlegendentrycollection#items)|Ruft die geladenen untergeordneten Elemente in dieser Sammlung ab.|
|[ChartLineFormat](/javascript/api/excel/excel.chartlineformat)|[Linestyle](/javascript/api/excel/excel.chartlineformat#linestyle)|Stellt die Linienart dar.|
||[weight](/javascript/api/excel/excel.chartlineformat#weight)|Stärke der Linie, angegeben in Punkt.|
|[ChartPoint](/javascript/api/excel/excel.chartpoint)|[hasDataLabel](/javascript/api/excel/excel.chartpoint#hasdatalabel)|Gibt an, ob ein Datenpunkt über eine Datenbeschriftung verfügt.|
||[markerBackgroundColor](/javascript/api/excel/excel.chartpoint#markerbackgroundcolor)|HTML-Farbcodedarstellung der Hintergrundfarbe des Markers eines Datenpunkts (z. B. #FF0000 stellt Rot dar).|
||[markerForegroundColor](/javascript/api/excel/excel.chartpoint#markerforegroundcolor)|HTML-Farbcodedarstellung der Vordergrundfarbe der Markierung eines Datenpunkts (z. B. #FF0000 stellt Rot dar).|
||[markerSize](/javascript/api/excel/excel.chartpoint#markersize)|Stellt die Markierungsgröße eines Datenpunkts dar.|
||[Markerstyle](/javascript/api/excel/excel.chartpoint#markerstyle)|Stellt den Markerstil eines Diagrammdatenpunkts dar.|
||[dataLabel](/javascript/api/excel/excel.chartpoint#datalabel)|Gibt die Datenbeschriftung eines Diagrammpunkts zurück.|
|[ChartPointFormat](/javascript/api/excel/excel.chartpointformat)|[Grenze](/javascript/api/excel/excel.chartpointformat#border)|Stellt das Rahmenformat eines Diagrammdatenpunkts dar, das Farb-, Format- und Gewichtungsinformationen enthält.|
|[ChartSeries](/javascript/api/excel/excel.chartseries)|[Charttype](/javascript/api/excel/excel.chartseries#charttype)|Stellt den Diagrammtyp einer Reihe dar.|
||[delete()](/javascript/api/excel/excel.chartseries#delete--)|Löscht die Diagrammreihen.|
||[doughnutHoleSize](/javascript/api/excel/excel.chartseries#doughnutholesize)|Stellt die Innenringgröße einer Diagrammreihe dar.|
||[Gefiltert](/javascript/api/excel/excel.chartseries#filtered)|Gibt an, ob die Datenreihe gefiltert wird.|
||[gapWidth](/javascript/api/excel/excel.chartseries#gapwidth)|Stellt die Abstandsbreite einer Diagrammreihe dar.|
||[hasDataLabels](/javascript/api/excel/excel.chartseries#hasdatalabels)|Gibt an, ob die Datenreihe Datenbeschriftungen enthält.|
||[markerBackgroundColor](/javascript/api/excel/excel.chartseries#markerbackgroundcolor)|Gibt die Hintergrundfarbe des Markers einer Diagrammreihe an.|
||[markerForegroundColor](/javascript/api/excel/excel.chartseries#markerforegroundcolor)|Gibt die Vordergrundfarbe der Markierung einer Diagrammreihe an.|
||[markerSize](/javascript/api/excel/excel.chartseries#markersize)|Gibt die Markierungsgröße einer Diagrammreihe an.|
||[Markerstyle](/javascript/api/excel/excel.chartseries#markerstyle)|Gibt die Markierungsart einer Diagrammreihe an.|
||[plotOrder](/javascript/api/excel/excel.chartseries#plotorder)|Gibt die Zeichnungsreihenfolge einer Diagrammreihe innerhalb der Diagrammgruppe an.|
||[Trendlinien](/javascript/api/excel/excel.chartseries#trendlines)|Die Sammlung von Trendlinien in der Datenreihe.|
||[setBubbleSizes(sourceData: Range)](/javascript/api/excel/excel.chartseries#setbubblesizes-sourcedata-)|Legt die Blasengrößen für eine Diagrammreihe fest.|
||[setValues(sourceData: Range)](/javascript/api/excel/excel.chartseries#setvalues-sourcedata-)|Legt die Werte für eine Diagrammreihe fest.|
||[setXAxisValues(sourceData: Range)](/javascript/api/excel/excel.chartseries#setxaxisvalues-sourcedata-)|Legt die Werte der X-Achse für eine Diagrammreihe fest.|
||[showShadow](/javascript/api/excel/excel.chartseries#showshadow)|Gibt an, ob die Datenreihe einen Schatten hat.|
||[Glatt](/javascript/api/excel/excel.chartseries#smooth)|Gibt an, ob die Datenreihe reibungslos verläuft.|
|[ChartSeriesCollection](/javascript/api/excel/excel.chartseriescollection)|[add(name?: string, index?: number)](/javascript/api/excel/excel.chartseriescollection#add-name--index-)|Fügt der Sammlung eine neue Reihe hinzu.|
|[ChartTitle](/javascript/api/excel/excel.charttitle)|[getSubstring(start: number, length: number)](/javascript/api/excel/excel.charttitle#getsubstring-start--length-)|Ruft die Teilzeichenfolge eines Diagrammtitels ab.|
||[horizontalAlignment](/javascript/api/excel/excel.charttitle#horizontalalignment)|Gibt die horizontale Ausrichtung für den Diagrammtitel an.|
||[left](/javascript/api/excel/excel.charttitle#left)|Gibt den Abstand zwischen dem linken Rand des Diagrammtitels und dem linken Rand des Diagrammbereichs in Punkt an.|
||[position](/javascript/api/excel/excel.charttitle#position)|Gibt die Position des Diagrammtitels an.|
||[height](/javascript/api/excel/excel.charttitle#height)|Gibt die Höhe des Diagrammtitels in Punkten zurück.|
||[width](/javascript/api/excel/excel.charttitle#width)|Gibt die Breite des Diagrammtitels in Punkt an.|
||[setFormula(formula: string)](/javascript/api/excel/excel.charttitle#setformula-formula-)|Legt einen Zeichenfolgenwert fest, der die Formel des Diagrammtitels unter Verwendung der A1-Schreibweise angibt.|
||[showShadow](/javascript/api/excel/excel.charttitle#showshadow)|Stellt einen booleschen Wert dar, der bestimmt, ob der Diagrammtitel über eine Schattierung verfügt.|
||[textOrientation](/javascript/api/excel/excel.charttitle#textorientation)|Gibt den Winkel an, an dem der Text für den Diagrammtitel ausgerichtet ist.|
||[top](/javascript/api/excel/excel.charttitle#top)|Gibt den Abstand zwischen dem oberen Rand des Diagrammtitels und dem oberen Rand des Diagrammbereichs in Punkt an.|
||[verticalAlignment](/javascript/api/excel/excel.charttitle#verticalalignment)|Gibt die vertikale Ausrichtung des Diagrammtitels an.|
|[ChartTitleFormat](/javascript/api/excel/excel.charttitleformat)|[Grenze](/javascript/api/excel/excel.charttitleformat#border)|Stellt das Rahmenformat des Diagrammtitels dar, einschließlich Farbe, Linienart und Gewichtung.|
|[ChartTrendline](/javascript/api/excel/excel.charttrendline)|[delete()](/javascript/api/excel/excel.charttrendline#delete--)|Löschen des Trendline-Objekts.|
||[Abfangen](/javascript/api/excel/excel.charttrendline#intercept)|Stellt den Wert der y-Schnittpunkt der Trendlinie dar.|
||[movingAveragePeriod](/javascript/api/excel/excel.charttrendline#movingaverageperiod)|Stellt den Zeitraum einer Diagrammtrendlinie dar.|
||[name](/javascript/api/excel/excel.charttrendline#name)|Gibt den Namen der Trendlinie zurück.|
||[polynomialOrder](/javascript/api/excel/excel.charttrendline#polynomialorder)|Stellt die Reihenfolge einer Diagrammtrendlinie dar.|
||[format](/javascript/api/excel/excel.charttrendline#format)|Stellt die Formatierung der Diagrammtrendlinie dar.|
||[type](/javascript/api/excel/excel.charttrendline#type)|Stellt die Beschriftung einer Diagrammtrendlinie dar.|
|[ChartTrendlineCollection](/javascript/api/excel/excel.charttrendlinecollection)|[add(type?: Excel. ChartTrendlineType)](/javascript/api/excel/excel.charttrendlinecollection#add-type-)|Fügt der Trendliniensammlung eine neue Trendlinie hinzu.|
||[getCount()](/javascript/api/excel/excel.charttrendlinecollection#getcount--)|Gibt die Anzahl der Trendlinien in der Sammlung zurück.|
||[getItem(index: number)](/javascript/api/excel/excel.charttrendlinecollection#getitem-index-)|Ruft ein Trendlinienobjekt nach Index ab, bei dem es sich um die Einfügereihenfolge im Elementarray handelt.|
||[items](/javascript/api/excel/excel.charttrendlinecollection#items)|Ruft die geladenen untergeordneten Elemente in dieser Sammlung ab.|
|[ChartTrendlineFormat](/javascript/api/excel/excel.charttrendlineformat)|[line](/javascript/api/excel/excel.charttrendlineformat#line)|Stellt die Formatierung der Diagrammlinien dar.|
|[CustomProperty](/javascript/api/excel/excel.customproperty)|[delete()](/javascript/api/excel/excel.customproperty#delete--)|Löscht die benutzerdefinierte Eigenschaft.|
||[key](/javascript/api/excel/excel.customproperty#key)|Der Schlüssel der benutzerdefinierten Eigenschaft.|
||[type](/javascript/api/excel/excel.customproperty#type)|Der Typ des Werts, der für die benutzerdefinierte Eigenschaft verwendet wird.|
||[value](/javascript/api/excel/excel.customproperty#value)|Der Wert der benutzerdefinierten Eigenschaft.|
|[CustomPropertyCollection](/javascript/api/excel/excel.custompropertycollection)|[add(key: string, value: any)](/javascript/api/excel/excel.custompropertycollection#add-key--value-)|Erstellt eine neue benutzerdefinierte Eigenschaft oder legt eine vorhandene fest.|
||[deleteAll()](/javascript/api/excel/excel.custompropertycollection#deleteall--)|Löscht alle benutzerdefinierten Eigenschaften in dieser Sammlung.|
||[getCount()](/javascript/api/excel/excel.custompropertycollection#getcount--)|Ruft die Anzahl von benutzerdefinierten Eigenschaften ab.|
||[getItem(key: string)](/javascript/api/excel/excel.custompropertycollection#getitem-key-)|Ruft ein Objekt für benutzerdefinierte Eigenschaften über seinen Schlüssel ab, bei dem Groß-/Kleinschreibung nicht beachtet wird.|
||[getItemOrNullObject(key: string)](/javascript/api/excel/excel.custompropertycollection#getitemornullobject-key-)|Ruft ein Objekt für benutzerdefinierte Eigenschaften über seinen Schlüssel ab, bei dem Groß-/Kleinschreibung nicht beachtet wird.|
||[items](/javascript/api/excel/excel.custompropertycollection#items)|Ruft die geladenen untergeordneten Elemente in dieser Sammlung ab.|
|[DataConnectionCollection](/javascript/api/excel/excel.dataconnectioncollection)|[refreshAll()](/javascript/api/excel/excel.dataconnectioncollection#refreshall--)|Aktualisiert alle Datenverbindungen in der Auflistung.|
|[DocumentProperties](/javascript/api/excel/excel.documentproperties)|[Autor](/javascript/api/excel/excel.documentproperties#author)|Der Autor der Arbeitsmappe.|
||[Kategorie](/javascript/api/excel/excel.documentproperties#category)|Die Kategorie der Arbeitsmappe.|
||[comments](/javascript/api/excel/excel.documentproperties#comments)|Die Kommentare der Arbeitsmappe.|
||[company](/javascript/api/excel/excel.documentproperties#company)|Das Unternehmen der Arbeitsmappe.|
||[Schlüsselwörter](/javascript/api/excel/excel.documentproperties#keywords)|Die Schlüsselwörter der Arbeitsmappe.|
||[Manager](/javascript/api/excel/excel.documentproperties#manager)|Der Vorgesetzte der Arbeitsmappe.|
||[creationDate](/javascript/api/excel/excel.documentproperties#creationdate)|Ruft das Erstellungsdatum der Arbeitsmappe ab.|
||[Benutzerdefinierte](/javascript/api/excel/excel.documentproperties#custom)|Ruft die Sammlung der benutzerdefinierten Eigenschaften der Arbeitsmappe ab.|
||[lastAuthor](/javascript/api/excel/excel.documentproperties#lastauthor)|Ruft den letzten Autor der Arbeitsmappe ab.|
||[Revisionnumber](/javascript/api/excel/excel.documentproperties#revisionnumber)|Ruft die Revisionsnummer der Arbeitsmappe ab.|
||[Betreff](/javascript/api/excel/excel.documentproperties#subject)|Der Betreff der Arbeitsmappe.|
||[title](/javascript/api/excel/excel.documentproperties#title)|Der Titel der Arbeitsmappe.|
|[NamedItem](/javascript/api/excel/excel.nameditem)|[formula](/javascript/api/excel/excel.nameditem#formula)|Die Formel des benannten Elements.|
||[arrayValues](/javascript/api/excel/excel.nameditem#arrayvalues)|Gibt ein Objekt mit Werten und Typen des benannten Elements zurück.|
|[NamedItemArrayValues](/javascript/api/excel/excel.nameditemarrayvalues)|[types](/javascript/api/excel/excel.nameditemarrayvalues#types)|Stellt die Typen für jedes Element im benannten Elementarray dar.|
||[values](/javascript/api/excel/excel.nameditemarrayvalues#values)|Stellt die Werte der einzelnen Elemente im benannten Elementarray dar.|
|[Range](/javascript/api/excel/excel.range)|[getAbsoluteResizedRange(numRows: number, numColumns: number)](/javascript/api/excel/excel.range#getabsoluteresizedrange-numrows--numcolumns-)|Ruft ein `Range` Objekt mit derselben zelle oben links wie das aktuelle `Range` Objekt ab, jedoch mit der angegebenen Anzahl von Zeilen und Spalten.|
||[getImage()](/javascript/api/excel/excel.range#getimage--)|Rendert den Bereich als base64-codiertes PNG-Bild.|
||[getSurroundingRegion()](/javascript/api/excel/excel.range#getsurroundingregion--)|Gibt ein `Range` Objekt zurück, das den umgebenden Bereich für die obere linke Zelle in diesem Bereich darstellt.|
||[hyperlink](/javascript/api/excel/excel.range#hyperlink)|Stellt den Hyperlink für den aktuellen Bereich dar.|
||[numberFormatLocal](/javascript/api/excel/excel.range#numberformatlocal)|Stellt den Zahlenformatcode Excel für den angegebenen Bereich basierend auf den Spracheinstellungen des Benutzers dar.|
||[isEntireColumn](/javascript/api/excel/excel.range#isentirecolumn)|Gibt an, ob der angegebene Bereich eine ganze Spalte ist.|
||[isEntireRow](/javascript/api/excel/excel.range#isentirerow)|Gibt an, ob der angegebene Bereich eine ganze Zeile ist.|
||[showCard()](/javascript/api/excel/excel.range#showcard--)|Zeigt die Karte für eine aktive Zelle an, wenn sie einen hohen Wertinhalt hat.|
||[style](/javascript/api/excel/excel.range#style)|Stellt die Formatvorlage des aktuellen Bereichs dar.|
|[RangeFormat](/javascript/api/excel/excel.rangeformat)|[textOrientation](/javascript/api/excel/excel.rangeformat#textorientation)|Die Textausrichtung aller Zellen innerhalb des Bereichs.|
||[useStandardHeight](/javascript/api/excel/excel.rangeformat#usestandardheight)|Bestimmt, ob die Zeilenhöhe des `Range` Objekts der Standardhöhe des Blatts entspricht.|
||[useStandardWidth](/javascript/api/excel/excel.rangeformat#usestandardwidth)|Gibt an, ob die Spaltenbreite des `Range` Objekts der Standardbreite des Blatts entspricht.|
|[RangeHyperlink](/javascript/api/excel/excel.rangehyperlink)|[address](/javascript/api/excel/excel.rangehyperlink#address)|Stellt das URL-Ziel für den Hyperlink dar.|
||[documentReference](/javascript/api/excel/excel.rangehyperlink#documentreference)|Stellt das Dokumentverweisziel für den Hyperlink dar.|
||[Quickinfo](/javascript/api/excel/excel.rangehyperlink#screentip)|Stellt die Zeichenfolge dar, die angezeigt wird, wenn mit dem Mauszeiger über den Link gefahren wird.|
||[textToDisplay](/javascript/api/excel/excel.rangehyperlink#texttodisplay)|Stellt die Zeichenfolge dar, die in der oberen linken Zelle des Bereichs angezeigt wird.|
|[Style](/javascript/api/excel/excel.style)|[delete()](/javascript/api/excel/excel.style#delete--)|Löscht diese Formatvorlage.|
||[formulaHidden](/javascript/api/excel/excel.style#formulahidden)|Gibt an, ob die Formel ausgeblendet wird, wenn das Arbeitsblatt geschützt ist.|
||[horizontalAlignment](/javascript/api/excel/excel.style#horizontalalignment)|Stellt die horizontale Ausrichtung für den Stil dar.|
||[includeAlignment](/javascript/api/excel/excel.style#includealignment)|Gibt an, ob die Formatvorlage die Eigenschaften autoeinzug, horizontale Ausrichtung, vertikale Ausrichtung, Umbruchtext, Einzugsebene und Textausrichtung enthält.|
||[includeBorder](/javascript/api/excel/excel.style#includeborder)|Gibt an, ob die Formatvorlage die Eigenschaften Farbe, Farbindex, Linienart und Stärkerahmen enthält.|
||[includeFont](/javascript/api/excel/excel.style#includefont)|Gibt an, ob die Formatvorlage die Eigenschaften "Background", "bold", "color", "color index", "font style", "kursiv", "name", "size", "strikethrough", "subscript", "superscript" und "underline" enthält.|
||[includeNumber](/javascript/api/excel/excel.style#includenumber)|Gibt an, ob die Formatvorlage die Eigenschaft für das Zahlenformat enthält.|
||[includePatterns](/javascript/api/excel/excel.style#includepatterns)|Gibt an, ob die Formatvorlage die Innenbereichseigenschaften Farbe, Farbindex, Invertierung , Muster, Musterfarbe und Musterfarbindex enthält.|
||[includeProtection](/javascript/api/excel/excel.style#includeprotection)|Gibt an, ob die Formatvorlage die Eigenschaften für ausgeblendeten und gesperrten Schutz enthält.|
||[indentLevel](/javascript/api/excel/excel.style#indentlevel)|Eine ganze Zahl zwischen 0 und 250, die die Einzugsebene für die Formatvorlage angibt.|
||[locked](/javascript/api/excel/excel.style#locked)|Gibt an, ob das Objekt gesperrt ist, wenn das Arbeitsblatt geschützt ist.|
||[Numberformat](/javascript/api/excel/excel.style#numberformat)|Der Formatcode des Zahlenformats für die Formatvorlage.|
||[numberFormatLocal](/javascript/api/excel/excel.style#numberformatlocal)|Der lokalisierte Formatcode des Zahlenformats für die Formatvorlage.|
||[readingOrder](/javascript/api/excel/excel.style#readingorder)|Die Leserichtung für die Formatvorlage.|
||[borders](/javascript/api/excel/excel.style#borders)|Eine Auflistung von vier Rahmenobjekten, die die Formatvorlage der vier Rahmen darstellen.|
||[Builtin](/javascript/api/excel/excel.style#builtin)|Gibt an, ob es sich bei der Formatvorlage um eine integrierte Formatvorlage handelt.|
||[fill](/javascript/api/excel/excel.style#fill)|Die Füllung der Formatvorlage.|
||[font](/javascript/api/excel/excel.style#font)|Ein `Font` Objekt, das die Schriftart der Formatvorlage darstellt.|
||[name](/javascript/api/excel/excel.style#name)|Der Name der Formatvorlage.|
||[shrinkToFit](/javascript/api/excel/excel.style#shrinktofit)|Gibt an, ob Text automatisch so verkleinert wird, dass er in die verfügbare Spaltenbreite passt.|
||[verticalAlignment](/javascript/api/excel/excel.style#verticalalignment)|Gibt die vertikale Ausrichtung für die Formatvorlage an.|
||[wrapText](/javascript/api/excel/excel.style#wraptext)|Gibt an, ob Excel den Text im Objekt umschließt.|
|[StyleCollection](/javascript/api/excel/excel.stylecollection)|[add(name: string)](/javascript/api/excel/excel.stylecollection#add-name-)|Fügt der Sammlung eine neue Formatvorlage hinzu.|
||[getItem(name: string)](/javascript/api/excel/excel.stylecollection#getitem-name-)|Ruft einen `Style` anhand des Namens ab.|
||[items](/javascript/api/excel/excel.stylecollection#items)|Ruft die geladenen untergeordneten Elemente in dieser Sammlung ab.|
|[Table](/javascript/api/excel/excel.table)|[onChanged](/javascript/api/excel/excel.table#onchanged)|Tritt auf, wenn Daten in Zellen in einer bestimmten Tabelle geändert werden.|
||[onSelectionChanged](/javascript/api/excel/excel.table#onselectionchanged)|Tritt auf, wenn sich die Auswahl in einer bestimmten Tabelle ändert.|
|[TableChangedEventArgs](/javascript/api/excel/excel.tablechangedeventargs)|[address](/javascript/api/excel/excel.tablechangedeventargs#address)|Ruft die Adresse ab, die den geänderten Bereich der Tabelle auf einem bestimmten Arbeitsblatt darstellt.|
||[changeType](/javascript/api/excel/excel.tablechangedeventargs#changetype)|Ruft den Änderungstyp ab, der angibt, wie das geänderte Ereignis ausgelöst wird.|
||[source](/javascript/api/excel/excel.tablechangedeventargs#source)|Ruft die Quelle des Ereignisses ab.|
||[tableId](/javascript/api/excel/excel.tablechangedeventargs#tableid)|Ruft die ID der Tabelle ab, in der die Daten geändert wurden.|
||[type](/javascript/api/excel/excel.tablechangedeventargs#type)|Ruft den Typ des Ereignisses ab.|
||[worksheetId](/javascript/api/excel/excel.tablechangedeventargs#worksheetid)|Ruft die ID des Arbeitsblatts ab, in dem die Daten geändert wurden.|
|[TableCollection](/javascript/api/excel/excel.tablecollection)|[onChanged](/javascript/api/excel/excel.tablecollection#onchanged)|Tritt auf, wenn Daten in einer beliebigen Tabelle in einer Arbeitsmappe oder einem Arbeitsblatt geändert werden.|
|[TableSelectionChangedEventArgs](/javascript/api/excel/excel.tableselectionchangedeventargs)|[address](/javascript/api/excel/excel.tableselectionchangedeventargs#address)|Ruft die Bereichsadresse ab, die den ausgewählten Bereich der Tabelle auf einem bestimmten Arbeitsblatt darstellt.|
||[isInsideTable](/javascript/api/excel/excel.tableselectionchangedeventargs#isinsidetable)|Gibt an, ob sich die Auswahl innerhalb einer Tabelle befindet.|
||[tableId](/javascript/api/excel/excel.tableselectionchangedeventargs#tableid)|Ruft die ID der Tabelle ab, in der sich die Auswahl geändert hat.|
||[type](/javascript/api/excel/excel.tableselectionchangedeventargs#type)|Ruft den Typ des Ereignisses ab.|
||[worksheetId](/javascript/api/excel/excel.tableselectionchangedeventargs#worksheetid)|Ruft die ID des Arbeitsblatts ab, in dem sich die Auswahl geändert hat.|
|[Workbook](/javascript/api/excel/excel.workbook)|[getActiveCell()](/javascript/api/excel/excel.workbook#getactivecell--)|Ruft die derzeit aktive Zelle aus der Arbeitsmappe ab.|
||[Dataconnections](/javascript/api/excel/excel.workbook#dataconnections)|Stellt alle Datenverbindungen in der Arbeitsmappe dar.|
||[name](/javascript/api/excel/excel.workbook#name)|Ruft den Namen der Arbeitsmappe ab.|
||[properties](/javascript/api/excel/excel.workbook#properties)|Ruft die Arbeitsmappeneigenschaften ab.|
||[protection](/javascript/api/excel/excel.workbook#protection)|Gibt das Schutzobjekt für eine Arbeitsmappe zurück.|
||[Stile](/javascript/api/excel/excel.workbook#styles)|Stellt eine Auflistung der mit der Arbeitsmappe verknüpften Formatvorlagen dar.|
|[WorkbookProtection](/javascript/api/excel/excel.workbookprotection)|[protect(password?: string)](/javascript/api/excel/excel.workbookprotection#protect-password-)|Schützt ein Arbeitsblatt.|
||[Geschützt](/javascript/api/excel/excel.workbookprotection#protected)|Gibt an, ob die Arbeitsmappe geschützt ist.|
||[unprotect(password?: string)](/javascript/api/excel/excel.workbookprotection#unprotect-password-)|Schützt eine Arbeitsmappe.|
|[Arbeitsblatt](/javascript/api/excel/excel.worksheet)|[copy(positionType?: Excel. WorksheetPositionType, relativeTo?: Excel. Arbeitsblatt)](/javascript/api/excel/excel.worksheet#copy-positiontype--relativeto-)|Kopiert ein Arbeitsblatt und platziert es an der angegebenen Position.|
||[getRangeByIndexes(startRow: number, startColumn: number, rowCount: number, columnCount: number)](/javascript/api/excel/excel.worksheet#getrangebyindexes-startrow--startcolumn--rowcount--columncount-)|Ruft das Objekt ab, das `Range` mit einem bestimmten Zeilen- und Spaltenindex beginnt und eine bestimmte Anzahl von Zeilen und Spalten umfasst.|
||[freezePanes](/javascript/api/excel/excel.worksheet#freezepanes)|Ruft ein Objekt ab, mit dem fixierte Fensterausschnitte auf dem Arbeitsblatt bearbeitet werden können.|
||[onActivated](/javascript/api/excel/excel.worksheet#onactivated)|Tritt auf, wenn das Arbeitsblatt aktiviert wird.|
||[onChanged](/javascript/api/excel/excel.worksheet#onchanged)|Tritt auf, wenn Daten in einem bestimmten Arbeitsblatt geändert werden.|
||[onDeactivated](/javascript/api/excel/excel.worksheet#ondeactivated)|Tritt auf, wenn das Arbeitsblatt deaktiviert wird.|
||[onSelectionChanged](/javascript/api/excel/excel.worksheet#onselectionchanged)|Tritt auf, wenn sich die Auswahl in einem bestimmten Arbeitsblatt ändert.|
||[standardHeight](/javascript/api/excel/excel.worksheet#standardheight)|Gibt die Standardhöhe (Standard) aller Zeilen in der Arbeitsmappe in Punkt zurück.|
||[standardWidth](/javascript/api/excel/excel.worksheet#standardwidth)|Gibt die Standardbreite (Standard) aller Spalten im Arbeitsblatt an.|
||[tabColor](/javascript/api/excel/excel.worksheet#tabcolor)|Die Registerkartenfarbe des Arbeitsblatts.|
|[WorksheetActivatedEventArgs](/javascript/api/excel/excel.worksheetactivatedeventargs)|[type](/javascript/api/excel/excel.worksheetactivatedeventargs#type)|Ruft den Typ des Ereignisses ab.|
||[worksheetId](/javascript/api/excel/excel.worksheetactivatedeventargs#worksheetid)|Ruft die ID des Arbeitsblatts ab, das aktiviert ist.|
|[WorksheetAddedEventArgs](/javascript/api/excel/excel.worksheetaddedeventargs)|[source](/javascript/api/excel/excel.worksheetaddedeventargs#source)|Ruft die Quelle des Ereignisses ab.|
||[type](/javascript/api/excel/excel.worksheetaddedeventargs#type)|Ruft den Typ des Ereignisses ab.|
||[worksheetId](/javascript/api/excel/excel.worksheetaddedeventargs#worksheetid)|Ruft die ID des Arbeitsblatts ab, das der Arbeitsmappe hinzugefügt wird.|
|[WorksheetChangedEventArgs](/javascript/api/excel/excel.worksheetchangedeventargs)|[address](/javascript/api/excel/excel.worksheetchangedeventargs#address)|Ruft die Bereichsadresse ab, die den geänderten Bereich eines bestimmten Arbeitsblatts darstellt.|
||[changeType](/javascript/api/excel/excel.worksheetchangedeventargs#changetype)|Ruft den Änderungstyp ab, der angibt, wie das geänderte Ereignis ausgelöst wird.|
||[source](/javascript/api/excel/excel.worksheetchangedeventargs#source)|Ruft die Quelle des Ereignisses ab.|
||[type](/javascript/api/excel/excel.worksheetchangedeventargs#type)|Ruft den Typ des Ereignisses ab.|
||[worksheetId](/javascript/api/excel/excel.worksheetchangedeventargs#worksheetid)|Ruft die ID des Arbeitsblatts ab, in dem die Daten geändert wurden.|
|[WorksheetCollection](/javascript/api/excel/excel.worksheetcollection)|[onActivated](/javascript/api/excel/excel.worksheetcollection#onactivated)|Tritt auf, wenn ein beliebiges Arbeitsblatt in der Arbeitsmappe aktiviert wird.|
||[onAdded](/javascript/api/excel/excel.worksheetcollection#onadded)|Tritt auf, wenn der Arbeitsmappe ein neues Arbeitsblatt hinzugefügt wird.|
||[onDeactivated](/javascript/api/excel/excel.worksheetcollection#ondeactivated)|Tritt auf, wenn ein beliebiges Arbeitsblatt in der Arbeitsmappe deaktiviert wird.|
||[onDeleted](/javascript/api/excel/excel.worksheetcollection#ondeleted)|Tritt auf, wenn ein Arbeitsblatt aus der Arbeitsmappe gelöscht wird.|
|[WorksheetDeactivatedEventArgs](/javascript/api/excel/excel.worksheetdeactivatedeventargs)|[type](/javascript/api/excel/excel.worksheetdeactivatedeventargs#type)|Ruft den Typ des Ereignisses ab.|
||[worksheetId](/javascript/api/excel/excel.worksheetdeactivatedeventargs#worksheetid)|Ruft die ID des Arbeitsblatts ab, das deaktiviert ist.|
|[WorksheetDeletedEventArgs](/javascript/api/excel/excel.worksheetdeletedeventargs)|[source](/javascript/api/excel/excel.worksheetdeletedeventargs#source)|Ruft die Quelle des Ereignisses ab.|
||[type](/javascript/api/excel/excel.worksheetdeletedeventargs#type)|Ruft den Typ des Ereignisses ab.|
||[worksheetId](/javascript/api/excel/excel.worksheetdeletedeventargs#worksheetid)|Ruft die ID des Arbeitsblatts ab, das aus der Arbeitsmappe gelöscht wird.|
|[WorksheetFreezePanes](/javascript/api/excel/excel.worksheetfreezepanes)|[freezeAt(frozenRange: Range \| string)](/javascript/api/excel/excel.worksheetfreezepanes#freezeat-frozenrange-)|Legt die fixierten Zellen in der Ansicht des aktiven Arbeitsblatts fest.|
||[freezeColumns(count?: number)](/javascript/api/excel/excel.worksheetfreezepanes#freezecolumns-count-)|Fixieren sie die ersten Spalten des Arbeitsblatts.|
||[freezeRows(count?: number)](/javascript/api/excel/excel.worksheetfreezepanes#freezerows-count-)|Fixieren Sie die oberen Zeilen des Arbeitsblatts.|
||[getLocation()](/javascript/api/excel/excel.worksheetfreezepanes#getlocation--)|Ruft den Bereich ab, der die fixierten Zellen in der aktiven Ansicht des Arbeitsblatts beschreibt.|
||[getLocationOrNullObject()](/javascript/api/excel/excel.worksheetfreezepanes#getlocationornullobject--)|Ruft den Bereich ab, der die fixierten Zellen in der aktiven Ansicht des Arbeitsblatts beschreibt.|
||[unfreeze()](/javascript/api/excel/excel.worksheetfreezepanes#unfreeze--)|Entfernt alle fixierten Bereiche auf dem Arbeitsblatt.|
|[WorksheetProtection](/javascript/api/excel/excel.worksheetprotection)|[unprotect(password?: string)](/javascript/api/excel/excel.worksheetprotection#unprotect-password-)|Hebt den Schutz eines Arbeitsblatts auf.|
|[WorksheetProtectionOptions](/javascript/api/excel/excel.worksheetprotectionoptions)|[allowEditObjects](/javascript/api/excel/excel.worksheetprotectionoptions#alloweditobjects)|Stellt die Arbeitsblatt-Schutzoption dar, die das Bearbeiten von Objekten zulässt.|
||[allowEditScenarios](/javascript/api/excel/excel.worksheetprotectionoptions#alloweditscenarios)|Stellt die Arbeitsblatt-Schutzoption dar, die das Bearbeiten von Szenarien ermöglicht.|
||[Selectionmode](/javascript/api/excel/excel.worksheetprotectionoptions#selectionmode)|Stellt die Arbeitsblatt-Schutzoption zum Zulassen des Auswahlmodus dar.|
|[WorksheetSelectionChangedEventArgs](/javascript/api/excel/excel.worksheetselectionchangedeventargs)|[address](/javascript/api/excel/excel.worksheetselectionchangedeventargs#address)|Ruft die Bereichsadresse ab, die den ausgewählten Bereich auf einem bestimmten Arbeitsblatt darstellt.|
||[type](/javascript/api/excel/excel.worksheetselectionchangedeventargs#type)|Ruft den Typ des Ereignisses ab.|
||[worksheetId](/javascript/api/excel/excel.worksheetselectionchangedeventargs#worksheetid)|Ruft die ID des Arbeitsblatts ab, in dem sich die Auswahl geändert hat.|
