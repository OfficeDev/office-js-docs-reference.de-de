| Klasse | Felder | Beschreibung |
|:---|:---|:---|
|[BasicDataValidation](/javascript/api/excel/excel.basicdatavalidation)|[formula1](/javascript/api/excel/excel.basicdatavalidation#formula1)|Gibt den rechten Operanden an, wenn die Operatoreigenschaft auf einen binären Operator wie GreaterThan festgelegt ist (der linke Operand ist der Wert, den der Benutzer in die Zelle einzugeben versucht).|
||[formula2](/javascript/api/excel/excel.basicdatavalidation#formula2)|Gibt mit den ternären Operatoren Between und NotBetween den Operanden an, der an die Obergrenze gebunden ist.|
||[operator](/javascript/api/excel/excel.basicdatavalidation#operator)|Der Operator, der zum Überprüfen der Daten verwendet wird.|
|[Chart](/javascript/api/excel/excel.chart)|[categoryLabelLevel](/javascript/api/excel/excel.chart#categorylabellevel)|Gibt eine Enumerationskonstante auf Diagrammkategoriebeschriftungsebene an, die auf die Ebene der Quellkategoriebeschriftungen verweist.|
||[displayBlanksAs](/javascript/api/excel/excel.chart#displayblanksas)|Gibt an, wie leere Zellen in einem Diagramm gezeichnet werden.|
||[plotBy](/javascript/api/excel/excel.chart#plotby)|Gibt an, wie Spalten oder Zeilen als Datenreihen im Diagramm verwendet werden.|
||[plotVisibleOnly](/javascript/api/excel/excel.chart#plotvisibleonly)|„True“, wenn nur sichtbare Zellen dargestellt werden.|
||[onActivated](/javascript/api/excel/excel.chart#onactivated)|Tritt auf, wenn das Diagramm aktiviert wird.|
||[onDeactivated](/javascript/api/excel/excel.chart#ondeactivated)|Tritt auf, wenn das Diagramm deaktiviert wird.|
||[plotArea](/javascript/api/excel/excel.chart#plotarea)|Stellt die Zeichnungsfläche für das Diagramm dar.|
||[seriesNameLevel](/javascript/api/excel/excel.chart#seriesnamelevel)|Gibt eine Enumerationskonstante auf Der Namensebene der Diagrammreihe an, die auf die Ebene der Quelldatenreihennamen verweist.|
||[showDataLabelsOverMaximum](/javascript/api/excel/excel.chart#showdatalabelsovermaximum)|Gibt an, ob die Datenbeschriftungen angezeigt werden sollen, wenn der Wert größer als der Maximalwert auf der Größenachse ist.|
||[style](/javascript/api/excel/excel.chart#style)|Gibt die Diagrammformatvorlage für das Diagramm an.|
|[ChartActivatedEventArgs](/javascript/api/excel/excel.chartactivatedeventargs)|[chartId](/javascript/api/excel/excel.chartactivatedeventargs#chartid)|Ruft die ID des Diagramms ab, das aktiviert ist.|
||[type](/javascript/api/excel/excel.chartactivatedeventargs#type)|Ruft den Typ des Ereignisses ab.|
||[worksheetId](/javascript/api/excel/excel.chartactivatedeventargs#worksheetid)|Ruft die ID des Arbeitsblatts ab, in dem das Diagramm aktiviert wird.|
|[ChartAddedEventArgs](/javascript/api/excel/excel.chartaddedeventargs)|[chartId](/javascript/api/excel/excel.chartaddedeventargs#chartid)|Ruft die ID des Diagramms ab, das dem Arbeitsblatt hinzugefügt wird.|
||[source](/javascript/api/excel/excel.chartaddedeventargs#source)|Ruft die Quelle des Ereignisses ab.|
||[type](/javascript/api/excel/excel.chartaddedeventargs#type)|Ruft den Typ des Ereignisses ab.|
||[worksheetId](/javascript/api/excel/excel.chartaddedeventargs#worksheetid)|Ruft die ID des Arbeitsblatts ab, in dem das Diagramm hinzugefügt wird.|
|[ChartAxis](/javascript/api/excel/excel.chartaxis)|[Ausrichtung](/javascript/api/excel/excel.chartaxis#alignment)|Gibt die Ausrichtung für die angegebene Achsen-Teilstrichbeschriftung an.|
||[isBetweenCategories](/javascript/api/excel/excel.chartaxis#isbetweencategories)|Gibt an, ob die Größenachse die Rubrikenachse zwischen den Kategorien schneidet.|
||[Mehrstufige](/javascript/api/excel/excel.chartaxis#multilevel)|Gibt an, ob eine Achse über mehrere Ebenen verfügt.|
||[Numberformat](/javascript/api/excel/excel.chartaxis#numberformat)|Gibt den Formatcode für die Achsen-Teilstrichbeschriftung an.|
||[Offset](/javascript/api/excel/excel.chartaxis#offset)|Gibt den Abstand zwischen den Beschriftungsebenen und den Abstand zwischen der ersten Ebene und der Achsenlinie an.|
||[position](/javascript/api/excel/excel.chartaxis#position)|Gibt die angegebene Achsenposition an, an der sich die andere Achse schneidet.|
||[positionAt](/javascript/api/excel/excel.chartaxis#positionat)|Gibt die Achsenposition an, an der sich die andere Achse schneidet.|
||[setPositionAt(value: number)](/javascript/api/excel/excel.chartaxis#setpositionat-value-)|Legt die angegebene Achsenposition fest, an der sich die andere Achse schneidet.|
||[textOrientation](/javascript/api/excel/excel.chartaxis#textorientation)|Gibt den Winkel an, an dem der Text für die Teilstrichbeschriftung der Diagrammachse ausgerichtet ist.|
|[ChartAxisFormat](/javascript/api/excel/excel.chartaxisformat)|[fill](/javascript/api/excel/excel.chartaxisformat#fill)|Gibt die Formatierung der Diagrammfüllung an.|
|[ChartAxisTitle](/javascript/api/excel/excel.chartaxistitle)|[setFormula(formula: string)](/javascript/api/excel/excel.chartaxistitle#setformula-formula-)|Ein Zeichenfolgenwert, der die Formel des Diagrammachseltitels unter Verwendung der A1-Schreibweise angibt.|
|[ChartAxisTitleFormat](/javascript/api/excel/excel.chartaxistitleformat)|[Grenze](/javascript/api/excel/excel.chartaxistitleformat#border)|Gibt das Rahmenformat der Diagrammachse an, einschließlich Farbe, Linienart und Gewichtung.|
||[fill](/javascript/api/excel/excel.chartaxistitleformat#fill)|Gibt die Füllformatierung des Diagrammachsentitels an.|
|[ChartBorder](/javascript/api/excel/excel.chartborder)|[clear()](/javascript/api/excel/excel.chartborder#clear--)|Dient zum Löschen der Rahmenformatierung eines Diagrammelements.|
|[ChartCollection](/javascript/api/excel/excel.chartcollection)|[onActivated](/javascript/api/excel/excel.chartcollection#onactivated)|Tritt auf, wenn ein Diagramm aktiviert wird.|
||[onAdded](/javascript/api/excel/excel.chartcollection#onadded)|Tritt auf, wenn dem Arbeitsblatt ein neues Diagramm hinzugefügt wird.|
||[onDeactivated](/javascript/api/excel/excel.chartcollection#ondeactivated)|Tritt auf, wenn ein Diagramm deaktiviert wird.|
||[onDeleted](/javascript/api/excel/excel.chartcollection#ondeleted)|Tritt auf, wenn ein Diagramm gelöscht wird.|
|[ChartDataLabel](/javascript/api/excel/excel.chartdatalabel)|[Autotext](/javascript/api/excel/excel.chartdatalabel#autotext)|Gibt an, ob die Datenbeschriftung automatisch den entsprechenden Text basierend auf dem Kontext generiert.|
||[formula](/javascript/api/excel/excel.chartdatalabel#formula)|Ein Zeichenfolgenwert, der die Formel der Diagrammdatenbeschriftung unter Verwendung der A1-Schreibweise angibt.|
||[horizontalAlignment](/javascript/api/excel/excel.chartdatalabel#horizontalalignment)|Stellt die horizontale Ausrichtung für die Diagrammdatenbeschriftung dar.|
||[left](/javascript/api/excel/excel.chartdatalabel#left)|Gibt den Abstand zwischen dem linken Rand der Diagrammdatenbeschriftung und dem linken Rand des Diagrammbereichs in Punkten an.|
||[Numberformat](/javascript/api/excel/excel.chartdatalabel#numberformat)|Zeichenfolgenwert, der den Formatcode für die Datenbeschriftung angibt.|
||[format](/javascript/api/excel/excel.chartdatalabel#format)|Das Format der Diagrammdatenbeschriftung.|
||[height](/javascript/api/excel/excel.chartdatalabel#height)|Gibt die Höhe der Diagrammdatenbeschriftung in Punkten zurück.|
||[width](/javascript/api/excel/excel.chartdatalabel#width)|Gibt die Breite der Diagrammdatenbeschriftung in Punkten zurück.|
||[text](/javascript/api/excel/excel.chartdatalabel#text)|Eine Zeichenfolge, die den Text der Datenbeschriftung in einem Diagramm darstellt.|
||[textOrientation](/javascript/api/excel/excel.chartdatalabel#textorientation)|Stellt den Winkel dar, an dem der Text für die Diagrammdatenbeschriftung ausgerichtet ist.|
||[top](/javascript/api/excel/excel.chartdatalabel#top)|Gibt den Abstand zwischen dem oberen Rand der Diagrammdatenbeschriftung und dem oberen Rand des Diagrammbereichs in Punkten an.|
||[verticalAlignment](/javascript/api/excel/excel.chartdatalabel#verticalalignment)|Stellt die vertikale Ausrichtung der Diagrammdatenbeschriftung dar.|
|[ChartDataLabelFormat](/javascript/api/excel/excel.chartdatalabelformat)|[Grenze](/javascript/api/excel/excel.chartdatalabelformat#border)|Gibt das Rahmenformat einschließlich Farbe, Linienart und Stärke an.|
|[ChartDataLabels](/javascript/api/excel/excel.chartdatalabels)|[Autotext](/javascript/api/excel/excel.chartdatalabels#autotext)|Gibt an, ob Datenbeschriftungen basierend auf dem Kontext automatisch entsprechenden Text generieren.|
||[horizontalAlignment](/javascript/api/excel/excel.chartdatalabels#horizontalalignment)|Gibt die horizontale Ausrichtung für die Diagrammdatenbeschriftung an.|
||[Numberformat](/javascript/api/excel/excel.chartdatalabels#numberformat)|Gibt den Formatcode für Datenbeschriftungen an.|
||[textOrientation](/javascript/api/excel/excel.chartdatalabels#textorientation)|Stellt den Winkel dar, an dem der Text für Datenbeschriftungen ausgerichtet ist.|
||[verticalAlignment](/javascript/api/excel/excel.chartdatalabels#verticalalignment)|Stellt die vertikale Ausrichtung der Diagrammdatenbeschriftung dar.|
|[ChartDeactivatedEventArgs](/javascript/api/excel/excel.chartdeactivatedeventargs)|[chartId](/javascript/api/excel/excel.chartdeactivatedeventargs#chartid)|Ruft die ID des Diagramms ab, das deaktiviert wird.|
||[type](/javascript/api/excel/excel.chartdeactivatedeventargs#type)|Ruft den Typ des Ereignisses ab.|
||[worksheetId](/javascript/api/excel/excel.chartdeactivatedeventargs#worksheetid)|Ruft die ID des Arbeitsblatts ab, in dem das Diagramm deaktiviert wird.|
|[ChartDeletedEventArgs](/javascript/api/excel/excel.chartdeletedeventargs)|[chartId](/javascript/api/excel/excel.chartdeletedeventargs#chartid)|Ruft die ID des Diagramms ab, das aus dem Arbeitsblatt gelöscht wird.|
||[source](/javascript/api/excel/excel.chartdeletedeventargs#source)|Ruft die Quelle des Ereignisses ab.|
||[type](/javascript/api/excel/excel.chartdeletedeventargs#type)|Ruft den Typ des Ereignisses ab.|
||[worksheetId](/javascript/api/excel/excel.chartdeletedeventargs#worksheetid)|Ruft die ID des Arbeitsblatts ab, in dem das Diagramm gelöscht wird.|
|[ChartLegendEntry](/javascript/api/excel/excel.chartlegendentry)|[height](/javascript/api/excel/excel.chartlegendentry#height)|Gibt die Höhe des Legendeneintrags in der Diagrammlegende an.|
||[Index](/javascript/api/excel/excel.chartlegendentry#index)|Gibt den Index des Legendeneintrags in der Diagrammlegende an.|
||[left](/javascript/api/excel/excel.chartlegendentry#left)|Gibt den linken Wert eines Diagrammlegendeeintrags an.|
||[top](/javascript/api/excel/excel.chartlegendentry#top)|Gibt den oberen Rand eines Diagrammlegendeeintrags an.|
||[width](/javascript/api/excel/excel.chartlegendentry#width)|Stellt die Breite des Legendeneintrags im Diagramm Legende dar.|
|[ChartLegendFormat](/javascript/api/excel/excel.chartlegendformat)|[Grenze](/javascript/api/excel/excel.chartlegendformat#border)|Gibt das Rahmenformat einschließlich Farbe, Linienart und Stärke an.|
|[ChartPlotArea](/javascript/api/excel/excel.chartplotarea)|[height](/javascript/api/excel/excel.chartplotarea#height)|Gibt den Höhenwert einer Zeichnungsfläche an.|
||[insideHeight](/javascript/api/excel/excel.chartplotarea#insideheight)|Gibt den Innenhöhenwert einer Zeichnungsfläche an.|
||[insideLeft](/javascript/api/excel/excel.chartplotarea#insideleft)|Gibt den inneren linken Wert einer Zeichnungsfläche an.|
||[insideTop](/javascript/api/excel/excel.chartplotarea#insidetop)|Gibt den inneren oberen Wert einer Zeichnungsfläche an.|
||[insideWidth](/javascript/api/excel/excel.chartplotarea#insidewidth)|Gibt den Innenbreitenwert einer Zeichnungsfläche an.|
||[left](/javascript/api/excel/excel.chartplotarea#left)|Gibt den linken Wert einer Zeichnungsfläche an.|
||[position](/javascript/api/excel/excel.chartplotarea#position)|Gibt die Position einer Zeichnungsfläche an.|
||[format](/javascript/api/excel/excel.chartplotarea#format)|Gibt die Formatierung einer Diagrammdiagrammfläche an.|
||[top](/javascript/api/excel/excel.chartplotarea#top)|Gibt den oberen Wert einer Zeichnungsfläche an.|
||[width](/javascript/api/excel/excel.chartplotarea#width)|Gibt den Breitenwert einer Zeichnungsfläche an.|
|[ChartPlotAreaFormat](/javascript/api/excel/excel.chartplotareaformat)|[Grenze](/javascript/api/excel/excel.chartplotareaformat#border)|Gibt die Rahmenattribute einer Diagrammdiagrammfläche an.|
||[fill](/javascript/api/excel/excel.chartplotareaformat#fill)|Gibt das Füllformat eines Objekts an, das Informationen zur Hintergrundformatierung enthält.|
|[ChartSeries](/javascript/api/excel/excel.chartseries)|[axisGroup](/javascript/api/excel/excel.chartseries#axisgroup)|Gibt die Gruppe für die angegebene Datenreihe an.|
||[Explosion](/javascript/api/excel/excel.chartseries#explosion)|Gibt den Explosionswert für ein Kreis- oder Ringdiagrammsegment an.|
||[firstSliceAngle](/javascript/api/excel/excel.chartseries#firstsliceangle)|Gibt den Winkel des ersten Kreisdiagramm- oder Ringdiagrammsegments in Grad (von vertikal im Uhrzeigersinn) an.|
||[invertIfNegative](/javascript/api/excel/excel.chartseries#invertifnegative)|True, wenn Excel das Muster im Element umkehrt, wenn es einer negativen Zahl entspricht.|
||[Überlappen](/javascript/api/excel/excel.chartseries#overlap)|Gibt an, wie Balken und Spalten angeordnet sind.|
||[dataLabels](/javascript/api/excel/excel.chartseries#datalabels)|Stellt eine Auflistung aller Datenbeschriftungen in der Datenreihe dar.|
||[secondPlotSize](/javascript/api/excel/excel.chartseries#secondplotsize)|Gibt die Größe des sekundären Abschnitts eines Kreis-aus-Kreis-Diagramms oder eines Balken-aus-Kreis-Diagramms als Prozentsatz der Größe des primären Kreises an.|
||[splitType](/javascript/api/excel/excel.chartseries#splittype)|Gibt an, wie die beiden Abschnitte eines Kreis-aus-Kreis-Diagramms oder eines Balken-aus-Kreis-Diagramms geteilt werden.|
||[varyByCategories](/javascript/api/excel/excel.chartseries#varybycategories)|True, wenn Excel jeder Datenpunktmarkierung eine andere Farbe oder ein anderes Muster zuweist.|
|[ChartTrendline](/javascript/api/excel/excel.charttrendline)|[backwardPeriod](/javascript/api/excel/excel.charttrendline#backwardperiod)|Die Anzahl der Punkte, über die sich eine Trendlinie zurück erstreckt.|
||[forwardPeriod](/javascript/api/excel/excel.charttrendline#forwardperiod)|Die Anzahl der Punkte, über die sich eine Trendlinie vorwärts erstreckt.|
||[Etikett](/javascript/api/excel/excel.charttrendline#label)|Die Beschriftung einer Diagrammtrendlinie.|
||[showEquation](/javascript/api/excel/excel.charttrendline#showequation)|True, wenn die Formel für die Trendlinie im Diagramm angezeigt wird.|
||[showRSquared](/javascript/api/excel/excel.charttrendline#showrsquared)|True, wenn der r-quadratische Wert für die Trendlinie im Diagramm angezeigt wird.|
|[ChartTrendlineLabel](/javascript/api/excel/excel.charttrendlinelabel)|[Autotext](/javascript/api/excel/excel.charttrendlinelabel#autotext)|Gibt an, ob die Trendlinienbeschriftung automatisch den entsprechenden Text basierend auf dem Kontext generiert.|
||[formula](/javascript/api/excel/excel.charttrendlinelabel#formula)|Zeichenfolgenwert, der die Formel der Diagrammtrendlinienbeschriftung in der A1-Schreibweise darstellt.|
||[horizontalAlignment](/javascript/api/excel/excel.charttrendlinelabel#horizontalalignment)|Stellt die horizontale Ausrichtung der Diagrammtrendlinienbeschriftung dar.|
||[left](/javascript/api/excel/excel.charttrendlinelabel#left)|Stellt den Abstand zwischen dem linken Rand der Diagrammtrendlinienbeschriftung und dem linken Rand des Diagrammbereichs in Punkt dar.|
||[Numberformat](/javascript/api/excel/excel.charttrendlinelabel#numberformat)|Zeichenfolgenwert, der den Formatcode für die Trendlinienbeschriftung darstellt.|
||[format](/javascript/api/excel/excel.charttrendlinelabel#format)|Das Format der Diagrammtrendlinienbeschriftung.|
||[height](/javascript/api/excel/excel.charttrendlinelabel#height)|Gibt die Höhe der Diagrammtrendlinienbeschriftung in Punkten zurück.|
||[width](/javascript/api/excel/excel.charttrendlinelabel#width)|Gibt die Breite der Diagrammtrendlinienbeschriftung in Punkten zurück.|
||[text](/javascript/api/excel/excel.charttrendlinelabel#text)|Eine Zeichenfolge, die den Text der Trendlinienbeschriftung in einem Diagramm darstellt.|
||[textOrientation](/javascript/api/excel/excel.charttrendlinelabel#textorientation)|Stellt den Winkel dar, an dem der Text für die Diagrammtrendlinienbeschriftung ausgerichtet ist.|
||[top](/javascript/api/excel/excel.charttrendlinelabel#top)|Stellt den Abstand zwischen dem oberen Rand der Diagrammtrendlinienbeschriftung und dem oberen Rand des Diagrammbereichs in Punkt dar.|
||[verticalAlignment](/javascript/api/excel/excel.charttrendlinelabel#verticalalignment)|Stellt die vertikale Ausrichtung der Diagrammtrendlinienbeschriftung dar.|
|[ChartTrendlineLabelFormat](/javascript/api/excel/excel.charttrendlinelabelformat)|[Grenze](/javascript/api/excel/excel.charttrendlinelabelformat#border)|Gibt das Rahmenformat an, das Farbe, Linienart und Stärke umfasst.|
||[fill](/javascript/api/excel/excel.charttrendlinelabelformat#fill)|Gibt das Füllformat der aktuellen Diagrammtrendlinienbeschriftung an.|
||[font](/javascript/api/excel/excel.charttrendlinelabelformat#font)|Gibt die Schriftattribute (z. B. Schriftartname, Schriftgrad und Farbe) für eine Diagrammtrendlinienbeschriftung an.|
|[CustomDataValidation](/javascript/api/excel/excel.customdatavalidation)|[formula](/javascript/api/excel/excel.customdatavalidation#formula)|Eine benutzerdefinierte Formel für die Datenüberprüfung.|
|[DataPivotHierarchy](/javascript/api/excel/excel.datapivothierarchy)|[name](/javascript/api/excel/excel.datapivothierarchy#name)|Der Name des DataPivotHierarchy-Objekts.|
||[Numberformat](/javascript/api/excel/excel.datapivothierarchy#numberformat)|Das Zahlenformat des DataPivotHierarchy-Objekts.|
||[position](/javascript/api/excel/excel.datapivothierarchy#position)|Die Position des DataPivotHierarchy-Objekts.|
||[Feld](/javascript/api/excel/excel.datapivothierarchy#field)|Gibt die PivotFields-Objekte zurück, die dem DataPivotHierarchy-Objekt zugeordnet sind.|
||[id](/javascript/api/excel/excel.datapivothierarchy#id)|DIE ID des DataPivotHierarchy-Objekts.|
||[setToDefault()](/javascript/api/excel/excel.datapivothierarchy#settodefault--)|Setzt das DataPivotHierarchy-Objekt auf die Standardwerte zurück.|
||[showAs](/javascript/api/excel/excel.datapivothierarchy#showas)|Gibt an, ob die Daten als eine bestimmte Zusammenfassungsberechnung angezeigt werden sollen.|
||[summarizeBy](/javascript/api/excel/excel.datapivothierarchy#summarizeby)|Gibt an, ob alle Elemente des DataPivotHierarchy-Objekts angezeigt werden.|
|[DataPivotHierarchyCollection](/javascript/api/excel/excel.datapivothierarchycollection)|[add(pivotHierarchy: Excel. PivotHierarchy)](/javascript/api/excel/excel.datapivothierarchycollection#add-pivothierarchy-)|Fügt der aktuellen Achse das PivotHierarchy-Objekt hinzu.|
||[getCount()](/javascript/api/excel/excel.datapivothierarchycollection#getcount--)|Ruft die Anzahl der Pivot-Hierarchien in der Sammlung ab.|
||[getItem(name: string)](/javascript/api/excel/excel.datapivothierarchycollection#getitem-name-)|Ruft ein DataPivotHierarchy-Objekt anhand seines Namens oder seiner ID ab.|
||[GetItemOrNullObject(name: string)](/javascript/api/excel/excel.datapivothierarchycollection#getitemornullobject-name-)|Ruft das DataPivotHierarchy-Objekt anhand des Namens ab.|
||[items](/javascript/api/excel/excel.datapivothierarchycollection#items)|Ruft die geladenen untergeordneten Elemente in dieser Sammlung ab.|
||[remove(DataPivotHierarchy: Excel. DataPivotHierarchy)](/javascript/api/excel/excel.datapivothierarchycollection#remove-datapivothierarchy-)|Entfernt das PivotHierarchy-Objekt von der aktuellen Achse.|
|[DataValidation](/javascript/api/excel/excel.datavalidation)|[clear()](/javascript/api/excel/excel.datavalidation#clear--)|Löscht die Datenüberprüfung aus dem aktuellen Bereich.|
||[errorAlert](/javascript/api/excel/excel.datavalidation#erroralert)|Fehlermeldung, wenn Benutzer ungültige Daten eingibt.|
||[ignoreBlanks](/javascript/api/excel/excel.datavalidation#ignoreblanks)|Gibt an, ob die Datenüberprüfung für leere Zellen durchgeführt wird.|
||[prompt](/javascript/api/excel/excel.datavalidation#prompt)|Eingabeaufforderung, wenn Benutzer eine Zelle auswählen.|
||[type](/javascript/api/excel/excel.datavalidation#type)|Den Typ der Datenüberprüfung finden Sie unter `Excel.DataValidationType` Details .|
||[Gültig](/javascript/api/excel/excel.datavalidation#valid)|Gibt an, ob alle Zellwerte entsprechend den Datenüberprüfungsregeln gültig sind.|
||[Regel](/javascript/api/excel/excel.datavalidation#rule)|Datenüberprüfungsregel, die unterschiedliche Typen von Datenüberprüfungskriterien enthält.|
|[DataValidationErrorAlert](/javascript/api/excel/excel.datavalidationerroralert)|[meldung](/javascript/api/excel/excel.datavalidationerroralert#message)|Stellt die Fehlermeldung dar.|
||[showAlert](/javascript/api/excel/excel.datavalidationerroralert#showalert)|Gibt an, ob ein Fehlerwarnungsdialogfeld angezeigt werden soll, wenn ein Benutzer ungültige Daten eingibt.|
||[style](/javascript/api/excel/excel.datavalidationerroralert#style)|Den Warnungstyp der Datenüberprüfung finden Sie `Excel.DataValidationAlertStyle` unter "Details".|
||[title](/javascript/api/excel/excel.datavalidationerroralert#title)|Stellt den Titel des Dialogfelds "Fehlerwarnung" dar.|
|[DataValidationPrompt](/javascript/api/excel/excel.datavalidationprompt)|[meldung](/javascript/api/excel/excel.datavalidationprompt#message)|Gibt die Meldung der Eingabeaufforderung an.|
||[showPrompt](/javascript/api/excel/excel.datavalidationprompt#showprompt)|Gibt an, ob eine Eingabeaufforderung angezeigt wird, wenn ein Benutzer eine Zelle mit Datenüberprüfung auswählt.|
||[title](/javascript/api/excel/excel.datavalidationprompt#title)|Gibt den Titel für die Eingabeaufforderung an.|
|[DataValidationRule](/javascript/api/excel/excel.datavalidationrule)|[Benutzerdefinierte](/javascript/api/excel/excel.datavalidationrule#custom)|Kriterien für eine benutzerdefinierte Datenüberprüfung.|
||[date](/javascript/api/excel/excel.datavalidationrule#date)|Kriterien für die Datenüberprüfung.|
||[Decimal](/javascript/api/excel/excel.datavalidationrule#decimal)|Kriterien für die dezimale Datenüberprüfung.|
||[list](/javascript/api/excel/excel.datavalidationrule#list)|Führt die Kriterien für die Datenüberprüfung auf.|
||[textLength](/javascript/api/excel/excel.datavalidationrule#textlength)|Kriterien für die Gültigkeitsprüfung von Textlängen.|
||[time](/javascript/api/excel/excel.datavalidationrule#time)|Kriterien für die Zeitdatenüberprüfung.|
||[wholeNumber](/javascript/api/excel/excel.datavalidationrule#wholenumber)|Kriterien für die Gültigkeitsprüfung ganzer Zahlen.|
|[DateTimeDataValidation](/javascript/api/excel/excel.datetimedatavalidation)|[formula1](/javascript/api/excel/excel.datetimedatavalidation#formula1)|Gibt den rechten Operanden an, wenn die Operatoreigenschaft auf einen binären Operator wie GreaterThan festgelegt ist (der linke Operand ist der Wert, den der Benutzer in die Zelle einzugeben versucht).|
||[formula2](/javascript/api/excel/excel.datetimedatavalidation#formula2)|Gibt mit den ternären Operatoren Between und NotBetween den Operanden an, der an die Obergrenze gebunden ist.|
||[operator](/javascript/api/excel/excel.datetimedatavalidation#operator)|Der Operator, der zum Überprüfen der Daten verwendet wird.|
|[FilterPivotHierarchy](/javascript/api/excel/excel.filterpivothierarchy)|[enableMultipleFilterItems](/javascript/api/excel/excel.filterpivothierarchy#enablemultiplefilteritems)|Gibt an, ob mehrere Filterelemente zulässig sind.|
||[name](/javascript/api/excel/excel.filterpivothierarchy#name)|Der Name des FilterPivotHierarchy-Objekts|
||[position](/javascript/api/excel/excel.filterpivothierarchy#position)|Die Position des FilterPivotHierarchy-Objekts|
||[fields](/javascript/api/excel/excel.filterpivothierarchy#fields)|Gibt die PivotFields-Objekte zurück, die dem FilterPivotHierarchy-Objekt zugeordnet sind.|
||[id](/javascript/api/excel/excel.filterpivothierarchy#id)|DIE ID des FilterPivotHierarchy-Objekts.|
||[setToDefault()](/javascript/api/excel/excel.filterpivothierarchy#settodefault--)|Setzt das FilterPivotHierarchy-Objekt auf die Standardwerte zurück.|
|[FilterPivotHierarchyCollection](/javascript/api/excel/excel.filterpivothierarchycollection)|[add(pivotHierarchy: Excel. PivotHierarchy)](/javascript/api/excel/excel.filterpivothierarchycollection#add-pivothierarchy-)|Fügt der aktuellen Achse das PivotHierarchy-Objekt hinzu.|
||[getCount()](/javascript/api/excel/excel.filterpivothierarchycollection#getcount--)|Ruft die Anzahl der Pivot-Hierarchien in der Sammlung ab.|
||[getItem(name: string)](/javascript/api/excel/excel.filterpivothierarchycollection#getitem-name-)|Ruft ein FilterPivotHierarchy anhand des Namens oder der ID ab.|
||[GetItemOrNullObject(name: string)](/javascript/api/excel/excel.filterpivothierarchycollection#getitemornullobject-name-)|Ruft das FilterPivotHierarchy-Objekt anhand des Namens ab.|
||[items](/javascript/api/excel/excel.filterpivothierarchycollection#items)|Ruft die geladenen untergeordneten Elemente in dieser Sammlung ab.|
||[remove(filterPivotHierarchy: Excel. FilterPivotHierarchy)](/javascript/api/excel/excel.filterpivothierarchycollection#remove-filterpivothierarchy-)|Entfernt das PivotHierarchy-Objekt von der aktuellen Achse.|
|[ListDataValidation](/javascript/api/excel/excel.listdatavalidation)|[inCellDropDown](/javascript/api/excel/excel.listdatavalidation#incelldropdown)|Gibt an, ob die Liste in einer Zell-Dropdownliste angezeigt werden soll.|
||[source](/javascript/api/excel/excel.listdatavalidation#source)|Quelle der Liste für die Datenüberprüfung|
|[PivotField](/javascript/api/excel/excel.pivotfield)|[name](/javascript/api/excel/excel.pivotfield#name)|Der Name von PivotField.|
||[id](/javascript/api/excel/excel.pivotfield#id)|DIE ID des PivotFields.|
||[items](/javascript/api/excel/excel.pivotfield#items)|Gibt die PivotFields-Objekte zurück, die dem PivotField-Objekt zugeordnet sind.|
||[showAllItems](/javascript/api/excel/excel.pivotfield#showallitems)|Legt fest, ob alle Elemente des PivotField-Objekts angezeigt werden.|
||[sortByLabels(sortBy: SortBy)](/javascript/api/excel/excel.pivotfield#sortbylabels-sortby-)|Sortiert das PivotField-Objekt.|
||[subtotals](/javascript/api/excel/excel.pivotfield#subtotals)|Zwischensumme von PivotField|
|[PivotFieldCollection](/javascript/api/excel/excel.pivotfieldcollection)|[getCount()](/javascript/api/excel/excel.pivotfieldcollection#getcount--)|Ruft die Anzahl der Pivot-Felder in der Auflistung ab.|
||[getItem(name: string)](/javascript/api/excel/excel.pivotfieldcollection#getitem-name-)|Ruft ein PivotField anhand seines Namens oder seiner ID ab.|
||[GetItemOrNullObject(name: string)](/javascript/api/excel/excel.pivotfieldcollection#getitemornullobject-name-)|Ruft ein PivotField anhand des Namens ab.|
||[items](/javascript/api/excel/excel.pivotfieldcollection#items)|Ruft die geladenen untergeordneten Elemente in dieser Sammlung ab.|
|[PivotHierarchy](/javascript/api/excel/excel.pivothierarchy)|[name](/javascript/api/excel/excel.pivothierarchy#name)|Der Name des PivotHierarchy-Objekts|
||[fields](/javascript/api/excel/excel.pivothierarchy#fields)|Gibt die PivotFields-Objekte zurück, die dem PivotHierarchy-Objekt zugeordnet sind.|
||[id](/javascript/api/excel/excel.pivothierarchy#id)|DIE ID des PivotHierarchy-Objekts.|
|[PivotHierarchyCollection](/javascript/api/excel/excel.pivothierarchycollection)|[getCount()](/javascript/api/excel/excel.pivothierarchycollection#getcount--)|Ruft die Anzahl der Pivot-Hierarchien in der Sammlung ab.|
||[getItem(name: string)](/javascript/api/excel/excel.pivothierarchycollection#getitem-name-)|Ruft ein PivotHierarchy-Objekt anhand seines Namens oder seiner ID ab.|
||[GetItemOrNullObject(name: string)](/javascript/api/excel/excel.pivothierarchycollection#getitemornullobject-name-)|Ruft das PivotHierarchy-Objekt anhand des Namens ab.|
||[items](/javascript/api/excel/excel.pivothierarchycollection#items)|Ruft die geladenen untergeordneten Elemente in dieser Sammlung ab.|
|[PivotItem](/javascript/api/excel/excel.pivotitem)|[Isexpanded](/javascript/api/excel/excel.pivotitem#isexpanded)|Gibt an, ob das Element zum Anzeigen untergeordneter Elemente erweitert wird oder untergeordnete Elemente ausgeblendet werden.|
||[name](/javascript/api/excel/excel.pivotitem#name)|Der Name von PivotItem|
||[id](/javascript/api/excel/excel.pivotitem#id)|DIE ID des PivotItem-Steuerelements.|
||[visible](/javascript/api/excel/excel.pivotitem#visible)|Gibt an, ob das PivotItem-Objekt sichtbar ist.|
|[PivotItemCollection](/javascript/api/excel/excel.pivotitemcollection)|[getCount()](/javascript/api/excel/excel.pivotitemcollection#getcount--)|Ruft die Anzahl der PivotItems in der Auflistung ab.|
||[getItem(name: string)](/javascript/api/excel/excel.pivotitemcollection#getitem-name-)|Ruft ein PivotItem anhand seines Namens oder seiner ID ab.|
||[GetItemOrNullObject(name: string)](/javascript/api/excel/excel.pivotitemcollection#getitemornullobject-name-)|Ruft ein PivotItem anhand des Namens ab.|
||[items](/javascript/api/excel/excel.pivotitemcollection#items)|Ruft die geladenen untergeordneten Elemente in dieser Sammlung ab.|
|[PivotLayout](/javascript/api/excel/excel.pivotlayout)|[getColumnLabelRange()](/javascript/api/excel/excel.pivotlayout#getcolumnlabelrange--)|Gibt den Bereich zurück, in dem sich die Spaltenbeschriftungen in PivotTable befinden.|
||[getDataBodyRange()](/javascript/api/excel/excel.pivotlayout#getdatabodyrange--)|Gibt den Bereich zurück, in dem sich die Datenwerte in PivotTable befinden.|
||[getFilterAxisRange()](/javascript/api/excel/excel.pivotlayout#getfilteraxisrange--)|Gibt den Bereich des Filterbereichs von PivotTable zurück.|
||[getRange()](/javascript/api/excel/excel.pivotlayout#getrange--)|Gibt den Bereich an, in dem PivotTable vorhanden ist, mit Ausnahme des Filterbereichs.|
||[getRowLabelRange()](/javascript/api/excel/excel.pivotlayout#getrowlabelrange--)|Gibt den Bereich zurück, in dem sich die Zeilenbeschriftungen in PivotTable befinden.|
||[layoutType](/javascript/api/excel/excel.pivotlayout#layouttype)|Diese Eigenschaft gibt das PivotLayoutType-Objekt aller Felder in PivotTable an.|
||[showColumnGrandTotals](/javascript/api/excel/excel.pivotlayout#showcolumngrandtotals)|Gibt an, ob im PivotTable-Bericht Gesamtsummen für Spalten angezeigt werden.|
||[showRowGrandTotals](/javascript/api/excel/excel.pivotlayout#showrowgrandtotals)|Gibt an, ob im PivotTable-Bericht Gesamtsummen für Zeilen angezeigt werden.|
||[subtotalLocation](/javascript/api/excel/excel.pivotlayout#subtotallocation)|Diese Eigenschaft gibt die `SubtotalLocationType` aller Felder in der PivotTable an.|
|[PivotTable](/javascript/api/excel/excel.pivottable)|[delete()](/javascript/api/excel/excel.pivottable#delete--)|Aktualisiert PivotTable|
||[columnHierarchies](/javascript/api/excel/excel.pivottable#columnhierarchies)|Die Pivot-Hierarchien der Spalten von PivotTable.|
||[dataHierarchies](/javascript/api/excel/excel.pivottable#datahierarchies)|Die Pivot-Hierarchien der Daten von PivotTable.|
||[filterHierarchies](/javascript/api/excel/excel.pivottable#filterhierarchies)|Die Pivot-Hierarchien der Filter von PivotTable.|
||[Hierarchien](/javascript/api/excel/excel.pivottable#hierarchies)|Die Pivot-Hierarchien von PivotTable.|
||[Layout](/javascript/api/excel/excel.pivottable#layout)|Das PivotLayout-Objekt, das das Layout und die visuelle Struktur von PivotTable beschreibt.|
||[rowHierarchies](/javascript/api/excel/excel.pivottable#rowhierarchies)|Die Pivot-Hierarchien der Zeilen von PivotTable.|
|[PivotTableCollection](/javascript/api/excel/excel.pivottablecollection)|[add(name: string, source: Range \| string \| Table, destination: Range \| string)](/javascript/api/excel/excel.pivottablecollection#add-name--source--destination-)|Fügen Sie eine PivotTable basierend auf den angegebenen Quelldaten hinzu, und fügen Sie sie in der oberen linken Zelle des Zielbereichs ein.|
|[Range](/javascript/api/excel/excel.range)|[dataValidation](/javascript/api/excel/excel.range#datavalidation)|Gibt ein Datenüberprüfungsobjekt zurück.|
|[RowColumnPivotHierarchy](/javascript/api/excel/excel.rowcolumnpivothierarchy)|[name](/javascript/api/excel/excel.rowcolumnpivothierarchy#name)|Der Name von RowColumnPivotHierarchy.|
||[position](/javascript/api/excel/excel.rowcolumnpivothierarchy#position)|Die Position von RowColumnPivotHierarchy|
||[fields](/javascript/api/excel/excel.rowcolumnpivothierarchy#fields)|Gibt die PivotFields-Objekte zurück, die dem RowColumnPivotHierarchy-Objekt zugeordnet sind.|
||[id](/javascript/api/excel/excel.rowcolumnpivothierarchy#id)|ID des RowColumnPivotHierarchy-Objekts.|
||[setToDefault()](/javascript/api/excel/excel.rowcolumnpivothierarchy#settodefault--)|Setzt das RowColumnPivotHierarchy-Objekt auf die Standardwerte zurück.|
|[RowColumnPivotHierarchyCollection](/javascript/api/excel/excel.rowcolumnpivothierarchycollection)|[add(pivotHierarchy: Excel. PivotHierarchy)](/javascript/api/excel/excel.rowcolumnpivothierarchycollection#add-pivothierarchy-)|Fügt der aktuellen Achse das PivotHierarchy-Objekt hinzu.|
||[getCount()](/javascript/api/excel/excel.rowcolumnpivothierarchycollection#getcount--)|Ruft die Anzahl der Pivot-Hierarchien in der Sammlung ab.|
||[getItem(name: string)](/javascript/api/excel/excel.rowcolumnpivothierarchycollection#getitem-name-)|Ruft ein RowColumnPivotHierarchy-Objekt anhand seines Namens oder seiner ID ab.|
||[GetItemOrNullObject(name: string)](/javascript/api/excel/excel.rowcolumnpivothierarchycollection#getitemornullobject-name-)|Ruft das RowColumnPivotHierarchy-Objekt anhand des Namens ab.|
||[items](/javascript/api/excel/excel.rowcolumnpivothierarchycollection#items)|Ruft die geladenen untergeordneten Elemente in dieser Sammlung ab.|
||[remove(rowColumnPivotHierarchy: Excel. RowColumnPivotHierarchy)](/javascript/api/excel/excel.rowcolumnpivothierarchycollection#remove-rowcolumnpivothierarchy-)|Entfernt das PivotHierarchy-Objekt von der aktuellen Achse.|
|[Laufzeit](/javascript/api/excel/excel.runtime)|[Enableevents](/javascript/api/excel/excel.runtime#enableevents)|Umschalten von JavaScript-Ereignissen im aktuellen Aufgabenbereich oder Inhalts-Add-In.|
|[ShowAsRule](/javascript/api/excel/excel.showasrule)|[baseField](/javascript/api/excel/excel.showasrule#basefield)|Das PivotField, auf dem die Berechnung basieren `ShowAs` soll, falls zutreffend, entsprechend dem `ShowAsCalculation` Typ, andernfalls. `null`|
||[baseItem](/javascript/api/excel/excel.showasrule#baseitem)|Das Element, auf dem die Berechnung basieren `ShowAs` soll, falls zutreffend, entsprechend dem `ShowAsCalculation` Typ, sonst `null` .|
||[Berechnung](/javascript/api/excel/excel.showasrule#calculation)|Die `ShowAs` für das PivotField zu verwendende Berechnung.|
|[Style](/javascript/api/excel/excel.style)|[autoIndent](/javascript/api/excel/excel.style#autoindent)|Gibt an, ob Text automatisch eingezogen wird, wenn die Textausrichtung in einer Zelle auf eine gleiche Verteilung festgelegt ist.|
||[textOrientation](/javascript/api/excel/excel.style#textorientation)|Die Textausrichtung für die Formatvorlage.|
|[Subtotals](/javascript/api/excel/excel.subtotals)|[automatic](/javascript/api/excel/excel.subtotals#automatic)|Wenn `Automatic` dies auf festgelegt `true` ist, werden alle anderen Werte beim Festlegen der `Subtotals` ignoriert.|
||[Durchschnittliche](/javascript/api/excel/excel.subtotals#average)||
||[count](/javascript/api/excel/excel.subtotals#count)||
||[countNumbers](/javascript/api/excel/excel.subtotals#countnumbers)||
||[Max](/javascript/api/excel/excel.subtotals#max)||
||[Min](/javascript/api/excel/excel.subtotals#min)||
||[Produkt](/javascript/api/excel/excel.subtotals#product)||
||[standardDeviation](/javascript/api/excel/excel.subtotals#standarddeviation)||
||[standardDeviationP](/javascript/api/excel/excel.subtotals#standarddeviationp)||
||[Summe](/javascript/api/excel/excel.subtotals#sum)||
||[Varianz](/javascript/api/excel/excel.subtotals#variance)||
||[varianz](/javascript/api/excel/excel.subtotals#variancep)||
|[Table](/javascript/api/excel/excel.table)|[legacyId](/javascript/api/excel/excel.table#legacyid)|Gibt eine numerische ID zurück.|
|[TableChangedEventArgs](/javascript/api/excel/excel.tablechangedeventargs)|[getRange(ctx: Excel.RequestContext)](/javascript/api/excel/excel.tablechangedeventargs#getrange-ctx-)|Ruft den Bereich ab, der den geänderten Bereich einer Tabelle in einem bestimmten Arbeitsblatt darstellt.|
||[getRangeOrNullObject(ctx: Excel.RequestContext)](/javascript/api/excel/excel.tablechangedeventargs#getrangeornullobject-ctx-)|Ruft den Bereich ab, der den geänderten Bereich einer Tabelle in einem bestimmten Arbeitsblatt darstellt.|
|[Workbook](/javascript/api/excel/excel.workbook)|[readOnly](/javascript/api/excel/excel.workbook#readonly)|Gibt `true` zurück, ob die Arbeitsmappe im schreibgeschützten Modus geöffnet ist.|
|[WorkbookCreated](/javascript/api/excel/excel.workbookcreated)||[Arbeitsblatt](/javascript/api/excel/excel.worksheet)|[onCalculated](/javascript/api/excel/excel.worksheet#oncalculated)|Tritt auf, wenn das Arbeitsblatt berechnet wird.|
||[Showgridlines](/javascript/api/excel/excel.worksheet#showgridlines)|Gibt an, ob Gitternetzlinien für den Benutzer sichtbar sind.|
||[showHeadings](/javascript/api/excel/excel.worksheet#showheadings)|Gibt an, ob Überschriften für den Benutzer sichtbar sind.|
|[WorksheetCalculatedEventArgs](/javascript/api/excel/excel.worksheetcalculatedeventargs)|[type](/javascript/api/excel/excel.worksheetcalculatedeventargs#type)|Ruft den Typ des Ereignisses ab.|
||[worksheetId](/javascript/api/excel/excel.worksheetcalculatedeventargs#worksheetid)|Ruft die ID des Arbeitsblatts ab, in dem die Berechnung erfolgt ist.|
|[WorksheetChangedEventArgs](/javascript/api/excel/excel.worksheetchangedeventargs)|[getRange(ctx: Excel.RequestContext)](/javascript/api/excel/excel.worksheetchangedeventargs#getrange-ctx-)|Ruft den Bereich ab, der den geänderten Bereich eines bestimmten Arbeitsblatts darstellt.|
||[getRangeOrNullObject(ctx: Excel.RequestContext)](/javascript/api/excel/excel.worksheetchangedeventargs#getrangeornullobject-ctx-)|Ruft den Bereich ab, der den geänderten Bereich eines bestimmten Arbeitsblatts darstellt.|
|[WorksheetCollection](/javascript/api/excel/excel.worksheetcollection)|[onCalculated](/javascript/api/excel/excel.worksheetcollection#oncalculated)|Tritt auf, wenn ein beliebiges Arbeitsblatt in der Arbeitsmappe berechnet wird.|
