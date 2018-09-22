# <a name="excel-javascript-api-requirement-sets"></a>JavaScript-API-Anforderungssätze für Excel

Anforderungssätze sind benannte Gruppen von API-Mitgliedern. Office-Add-ins anforderungssätzen im Manifest angegebenen verwenden oder eine Überprüfung zur Laufzeit verwenden, um zu bestimmen, ob ein Office-Host APIs unterstützt, die ein Add-in muss. Weitere Informationen finden Sie unter [Office-Versionen und Anforderung festgelegt](https://docs.microsoft.com/office/dev/add-ins/develop/office-versions-and-requirement-sets).

Führen Sie Excel-add-ins in mehreren Versionen von Office, einschließlich Office 2016 für Windows, Office für iPad, Office für Mac und Office Online. Die folgende Tabelle enthält die Excel-anforderungssätze, Office-hostanwendungen, die jede Anforderungssatz und die Buildversionen unterstützt oder-Nummer Anwendungen.

> [!NOTE]
> Jeder API, die als **Beta** markiert ist, ist nicht für die Endbenutzer Produktion bereit. Wir stellen ihnen für Entwickler versucht, sie Auschecken in Umgebungen für Entwicklung und Test zur Verfügung. Es ist nicht vorgesehen, gegen Produktion/Business critical Dokumente verwendet werden soll.
> 
> Für die anforderungssätze, die als **Beta**gekennzeichnet sind, verwenden Sie die angegebene (oder höhere) Version von Office-Software und Verwenden der Beta-Bibliothek auf dem CDN: https://appsforoffice.microsoft.com/lib/beta/hosted/office.js. Einträge, die nicht als **Beta** sind im Allgemeinen verfügbar und können Produktion Bibliothek auf dem CDN: https://appsforoffice.microsoft.com/lib/1/hosted/office.js.

|  Anforderungssatz  |  Office 365 für Windows\*  |  Office 365 für iPad  |  Office 365 für Mac  | Office Online  |  Office Online Server  |
|:-----|-----|:-----|:-----|:-----|:-----|
| Beta  | Bitte [Besuchen Sie unsere JavaScript-API für Excel open-Spezifikation Seite](https://github.com/OfficeDev/office-js-docs/tree/ExcelJs_OpenSpec)! |
| ExcelApi1.8  | Version 1808 (Build 10730.20102) oder höher | 2,17 oder höher | 16,17 oder höher | September 2018 | Bald verfügbar |
| ExcelApi1.7  | Version 1801 (Build 9001.2171) oder höher   | 2,9 oder höher | 16,9 oder höher | April 2018 | Bald verfügbar |
| ExcelApi1.6  | Version 1704 (Build 8201.2001) oder höher   | 2.2 oder höher |15.36 oder höher| April 2017 | Bald verfügbar|
| ExcelApi1.5  | Version 1703 (Build 8067.2070) oder höher   | 2.2 oder höher |15.36 oder höher| März 2017 | Bald verfügbar|
| ExcelApi1.4  | Version 1701 (Build 7870.2024) oder höher   | 2.2 oder höher |15.36 oder höher| Januar 2017 | Bald verfügbar|
| ExcelApi1.3  | Version 1608 (Build 7369.2055) oder höher | 1.27 oder höher |  15.27 oder höher| September 2016 | Version 1608 (Build 7601.6800) oder höher|
| ExcelApi1.2  | Version 1601 (Build 6741.2088) oder höher | 1.21 oder höher | 15.22 oder höher| Januar 2016 ||
| ExcelApi1.1  | Version 1509 (Build 4266.1001) oder höher | 1.19 oder höher | 15.20 oder höher| Januar 2016 ||

> [!NOTE]
> Die Buildnummer für Office 2016 über MSI installiert ist 16.0.4266.1001. Diese Version enthält nur die ExcelApi 1.1 Anforderungssatz.

Weitere Informationen zu Versionen, Build-Nummern und Office Online Server finden Sie unter:

- [Versions- und Buildnummern der Updatekanalversionen für Office 365-Clients](https://support.office.com/article/version-and-build-numbers-of-update-channel-releases-ae942449-1fca-4484-898b-a933ea23def7)
- [Welche Version von Office verwende ich?](https://support.office.com/article/What-version-of-Office-am-I-using-932788b8-a3ce-44bf-bb09-e334518b8b19)
- [So finden Sie die Versions- und Bildnummer für eine Office 365-Clientanwendung](https://support.office.com/article/version-and-build-numbers-of-update-channel-releases-ae942449-1fca-4484-898b-a933ea23def7)
- 
  [Office Online Server-Übersicht](https://docs.microsoft.com/officeonlineserver/office-online-server-overview)

## <a name="whats-new-in-excel-javascript-api-18"></a>Was ist neu in Excel JavaScript-API 1,8

Die JavaScript-API für Excel-Anforderung Set 1,8 Features enthalten APIs für PivotTables, die datenüberprüfung, Diagramme, Ereignisse für Diagramme, Leistungsoptionen und Arbeitsmappe erstellen.

### <a name="pivottable"></a>PivotTable

Welle 2 der PivotTable-APIs können-add-ins die Hierarchien einer PivotTable festgelegt. Jetzt können Sie steuern, die Daten und wie es aggregiert werden. Unsere [PivotTable Artikel](https://docs.microsoft.com/office/dev/add-ins/excel/excel-add-ins-pivottables) mehr über die neue PivotTable-Funktionen.

### <a name="data-validation"></a>Datenüberprüfung

Data Validation-können, die Sie steuern, welche eines Benutzers in einem Arbeitsblatt eingibt. Sie können Zellen auf vordefinierte Antwort wird beschränken oder Popup Warnungen zu unerwünschten Eingabe übergeben. Erfahren Sie mehr über das [Hinzufügen von datenüberprüfung auf Bereiche](https://docs.microsoft.com/office/dev/add-ins/excel/excel-add-ins-data-validation) heute.

### <a name="charts"></a>Diagramme

Eine neue Runde des Diagramms APIs werden noch größere programmgesteuerte Kontrolle über Diagrammelemente. Sie haben nun größer Zugriff auf die Legende, Achsen, Trendlinie und Zeichnungsfläche.

### <a name="events"></a>Events

Weitere [Ereignisse](https://docs.microsoft.com/office/dev/add-ins/excel/excel-add-ins-events) wurden für Diagramme hinzugefügt. Haben Sie Ihre reagieren-add-in für Benutzer, die Interaktion mit dem Diagramm. Sie können auch die [Umschaltfläche Ereignisse](https://docs.microsoft.com/office/dev/add-ins/excel/performance#enable-and-disable-events) auslösen, über die gesamte Arbeitsmappe.


|Objekt| Neuerungen| Beschreibung|Anforderungssatz|
|:----|:----|:----|:----|
|[application](/javascript/api/excel/excel.application)|_Methode_ > [createWorkbook(base64File: string)](/javascript/api/excel/excel.application)|Erstellt eine neue Arbeitsmappe ausgeblendete mit optionalen base64-codierten XLSX-Dateien.|1,8|
|[basicDataValidation](/javascript/api/excel/excel.basicdatavalidation)|_Eigenschaft_ > formula1|Ruft ab oder legt diesen fest Formula1, d. h. Mindestwert oder Wert des Operators abhängig.|1,8|
|[basicDataValidation](/javascript/api/excel/excel.basicdatavalidation)|_Eigenschaft_ > formula2|Ruft ab oder legt diesen fest Formula2, d. h. Höchstwert oder Wert des Operators abhängig.|1,8|
|[basicDataValidation](/javascript/api/excel/excel.basicdatavalidation)|_Beziehung_ > Operator|Der Operator für die Validierung der Daten verwendet.|1,8|
|[chart](/javascript/api/excel/excel.chart)|_Eigenschaft_ > CategoryLabelLevel|Zurückgeben oder festlegen eine ChartCategoryLabelLevel-Enumeration-Konstante, die in Bezug auf die Ebene der, in dem die Kategorie Beschriftungen aus einer Führungslinie basieren wird sind. Lese-/Schreibzugriff.|1,8|
|[chart](/javascript/api/excel/excel.chart)|_Eigenschaft_ > PlotVisibleOnly|True, Wenn nur sichtbare Zellen gezeichnet werden. False, Wenn sichtbare und ausgeblendete Zellen gezeichnet werden. ReadWrite.|1,8|
|[chart](/javascript/api/excel/excel.chart)|_Eigenschaft_ > SeriesNameLevel|Zurückgeben oder festlegen eine ChartSeriesNameLevel-Enumeration-Konstante, die in Bezug auf die Ebene der, in dem die Namen von Datenreihen aus Quelle wird werden. Lese-/Schreibzugriff.|1,8|
|[chart](/javascript/api/excel/excel.chart)|_Eigenschaft_ > ShowDataLabelsOverMaximum|Stellt dar, ob für die datenbeschriftungen angezeigt, wenn der Wert größer als der maximale Wert für die Größenachse ist.|1,8|
|[chart](/javascript/api/excel/excel.chart)|_Eigenschaft_ > style|Zurückgeben oder Festlegen des Diagrammformats für das Diagramm. ReadWrite.|1,8|
|[chart](/javascript/api/excel/excel.chart)|_Beziehung_ > DisplayBlanksAs|Gibt zurück oder legt fest, wie leere Zellen in einem Diagramm gezeichnet werden. ReadWrite.|1,8|
|[chart](/javascript/api/excel/excel.chart)|_Beziehung_ > PlotArea|Die PlotArea für das Diagramm darstellt. Schreibgeschützt.|1,8|
|[chart](/javascript/api/excel/excel.chart)|_Beziehung_ > PlotBy|Gibt zurück oder legt fest, wie die Spalten und Zeilen als Datenreihen im Diagramm verwendet werden. ReadWrite.|1,8|
|[chartActivatedEventArgs](/javascript/api/excel/excel.chartactivatedeventargs)|_Eigenschaft_ > ChartId|Ruft die Id des Diagramms, das aktiviert wird.|1,8|
|[chartActivatedEventArgs](/javascript/api/excel/excel.chartactivatedeventargs)|_Eigenschaft_ > Typ|Ruft den Typ des Ereignisses ab.|1,8|
|[chartActivatedEventArgs](/javascript/api/excel/excel.chartactivatedeventargs)|_Eigenschaft_ > WorksheetId|Ruft die Id des Arbeitsblatts, in dem das Diagramm aktiviert wird.|1,8|
|[chartAddedEventArgs](/javascript/api/excel/excel.chartaddedeventargs)|_Eigenschaft_ > ChartId|Ruft die Id des Diagramms, das dem Arbeitsblatt hinzugefügt wird.|1,8|
|[chartAddedEventArgs](/javascript/api/excel/excel.chartaddedeventargs)|_Eigenschaft_ > Typ|Ruft den Typ des Ereignisses ab.|1,8|
|[chartAddedEventArgs](/javascript/api/excel/excel.chartaddedeventargs)|_Eigenschaft_ > WorksheetId|Ruft die Id des Arbeitsblatts in der das Diagramm hinzugefügt wird.|1,8|
|[chartAddedEventArgs](/javascript/api/excel/excel.chartaddedeventargs)|_Beziehung_ > Quelle|Ruft die Quelle des Ereignisses ab.|1,8|
|[chartAxis](/javascript/api/excel/excel.chartaxis)|_Eigenschaft_ > IsBetweenCategories|Stellt dar, ob die Größenachse die Rubrikenachse zwischen den Rubriken schneidet.|1,8|
|[chartAxis](/javascript/api/excel/excel.chartaxis)|_Eigenschaft_ > mit mehreren Ebenen|Stellt dar, ob eine Achse mit mehreren Ebenen oder nicht ist.|1,8|
|[chartAxis](/javascript/api/excel/excel.chartaxis)|_Eigenschaft_ > numberFormat|Stellt den Formatierungscode für den Abstand zwischen Achsenbeschriftung dar.|1,8|
|[chartAxis](/javascript/api/excel/excel.chartaxis)|_Eigenschaft_ > Offset|Den Abstand zwischen den Beschriftungsebenen und den Abstand zwischen der ersten Ebene und der Achsenlinie darstellt. Der Wert muss eine ganze Zahl zwischen 0 und 1000 sein.|1,8|
|[chartAxis](/javascript/api/excel/excel.chartaxis)|_Eigenschaft_ > PositionAt|Stellt die angegebene Achsenposition Schnittpunkt mit die anderen Achse an. Sie sollten die SetPositionAt(double)-Methode verwenden, um diese Eigenschaft festzulegen. Schreibgeschützt.|1,8|
|[chartAxis](/javascript/api/excel/excel.chartaxis)|_Eigenschaft_ > TextOrientation|Stellt die Ausrichtung des Texts der Achse Teilstrich Beschriftung. Der Wert muss eine ganze Zahl sein entweder von-90 bis 90 oder 180 für Text vertikal ausgerichtet.|1,8|
|[chartAxis](/javascript/api/excel/excel.chartaxis)|_Beziehung_ > Ausrichtung|Stellt die Ausrichtung für die Beschriftung der angegebenen Achse Teilstrich dar.|1,8|
|[chartAxis](/javascript/api/excel/excel.chartaxis)|_Beziehung_ > Position|Stellt die angegebene Achsenposition Schnittpunkt mit die anderen Achse an.|1,8|
|[chartAxis](/javascript/api/excel/excel.chartaxis)|_Methode_ > [SetPositionAt(value: double)](/javascript/api/excel/excel.chartaxis)|Festlegen der angegebenen Achsenposition Schnittpunkt mit die anderen Achse an.|1,8|
|[chartAxisFormat](/javascript/api/excel/excel.chartaxisformat)|_Beziehung_ > Füllung|Stellt die Formatierung Diagramm ausfüllen. Schreibgeschützt.|1,8|
|[chartAxisTitle](/javascript/api/excel/excel.chartaxistitle)|_Methode_ > [SetFormula(formula: string)](/javascript/api/excel/excel.chartaxistitle)|Ein String-Wert, der die Formel des Diagrammtitels Achse unter Verwendung der A1-Schreibweise darstellt.|1,8|
|[chartAxisTitleFormat](/javascript/api/excel/excel.chartaxistitleformat)|_Beziehung_ > Rahmen|Stellt das Rahmenformat, das Color, Linestyle und Weight enthält. Schreibgeschützt.|1,8|
|[chartAxisTitleFormat](/javascript/api/excel/excel.chartaxistitleformat)|_Beziehung_ > Füllung|Stellt die Formatierung Diagramm ausfüllen. Schreibgeschützt.|1,8|
|[chartBorder](/javascript/api/excel/excel.chartborder)|_Methode_ > [clear()](/javascript/api/excel/excel.chartborder)|Deaktivieren Sie das Rahmenformat eines Diagrammelements.|1,8|
|[chartDataLabel](/javascript/api/excel/excel.chartdatalabel)|_Eigenschaft_ > AutoText|Boolescher Wert, wenn Daten automatisch bezeichnen generiert basierend auf Kontext angemessenen Text.|1,8|
|[chartDataLabel](/javascript/api/excel/excel.chartdatalabel)|_Eigenschaft_ > Formel|String-Wert, der die Formel des Diagramms Datenbeschriftung unter Verwendung der A1-Schreibweise darstellt.|1,8|
|[chartDataLabel](/javascript/api/excel/excel.chartdatalabel)|_Eigenschaft_ > height|Gibt die Höhe des Datenbeschriftung Diagramms in Punkt zurück. Schreibgeschützt. NULL, wenn das Diagramm Datenbeschriftung nicht sichtbar ist. Schreibgeschützt.|1,8|
|[chartDataLabel](/javascript/api/excel/excel.chartdatalabel)|_Eigenschaft_ > linken|Stellt den Abstand in Punkt vom linken Rand des Diagramms Datenbeschriftung zum linken Rand des Diagrammbereichs dar. NULL, wenn das Diagramm Datenbeschriftung nicht sichtbar ist.|1,8|
|[chartDataLabel](/javascript/api/excel/excel.chartdatalabel)|_Eigenschaft_ > numberFormat|String-Wert, der den Formatierungscode für Datenbeschriftung darstellt.|1,8|
|[chartDataLabel](/javascript/api/excel/excel.chartdatalabel)|_Eigenschaft_ > text|Eine Zeichenfolge, die den Text der Datenbeschriftung in einem Diagramm darstellt.|1,8|
|[chartDataLabel](/javascript/api/excel/excel.chartdatalabel)|_Eigenschaft_ > TextOrientation|Stellt die Ausrichtung von Text der Datenbeschriftung Diagramm dar. Der Wert muss eine ganze Zahl sein entweder von-90 bis 90 oder 180 für Text vertikal ausgerichtet.|1,8|
|[chartDataLabel](/javascript/api/excel/excel.chartdatalabel)|_Eigenschaft_ > oben|Stellt den Abstand in Punkt vom oberen Rand der Datenbeschriftung an den Anfang der Diagrammfläche des Diagramms dar. NULL, wenn das Diagramm Datenbeschriftung nicht sichtbar ist.|1,8|
|[chartDataLabel](/javascript/api/excel/excel.chartdatalabel)|_Eigenschaft_ > width|Gibt die Breite in Punkt, der Datenbeschriftung Diagramm zurück. Schreibgeschützt. NULL, wenn das Diagramm Datenbeschriftung nicht sichtbar ist. Schreibgeschützt.|1,8|
|[chartDataLabel](/javascript/api/excel/excel.chartdatalabel)|_Beziehung_ > Format|Stellt das Format des Diagramms Datenbeschriftung. Schreibgeschützt.|1,8|
|[chartDataLabel](/javascript/api/excel/excel.chartdatalabel)|_Beziehung_ > horizontalAlignment|Die horizontale Ausrichtung für die Datenbeschriftung Diagramm darstellt.|1,8|
|[chartDataLabel](/javascript/api/excel/excel.chartdatalabel)|_Beziehung_ > verticalAlignment|Stellt die vertikale Ausrichtung des Diagramms Datenbeschriftung an.|1,8|
|[chartDataLabelFormat](/javascript/api/excel/excel.chartdatalabelformat)|_Beziehung_ > Rahmen|Stellt das Rahmenformat, das Color, Linestyle und Weight enthält. Schreibgeschützt.|1,8|
|[chartDataLabels](/javascript/api/excel/excel.chartdatalabels)|_Eigenschaft_ > AutoText|Stellt dar, ob datenbeschriftungen automatisch basierend auf Kontext angemessenen Text generiert.|1,8|
|[chartDataLabels](/javascript/api/excel/excel.chartdatalabels)|_Eigenschaft_ > numberFormat|Stellt den Formatierungscode für datenbeschriftungen an.|1,8|
|[chartDataLabels](/javascript/api/excel/excel.chartdatalabels)|_Eigenschaft_ > TextOrientation|Stellt die Ausrichtung von Text Datenbeschriftungstyp. Der Wert muss eine ganze Zahl sein, entweder von-90 bis 90 oder 0 bis 180 für Text vertikal ausgerichtet.|1,8|
|[chartDataLabels](/javascript/api/excel/excel.chartdatalabels)|_Beziehung_ > horizontalAlignment|Die horizontale Ausrichtung für die Datenbeschriftung Diagramm darstellt.|1,8|
|[chartDataLabels](/javascript/api/excel/excel.chartdatalabels)|_Beziehung_ > verticalAlignment|Stellt die vertikale Ausrichtung des Diagramms Datenbeschriftung an.|1,8|
|[chartDeactivatedEventArgs](/javascript/api/excel/excel.chartdeactivatedeventargs)|_Eigenschaft_ > ChartId|Ruft die Id des Diagramms, das deaktiviert wird.|1,8|
|[chartDeactivatedEventArgs](/javascript/api/excel/excel.chartdeactivatedeventargs)|_Eigenschaft_ > Typ|Ruft den Typ des Ereignisses ab.|1,8|
|[chartDeactivatedEventArgs](/javascript/api/excel/excel.chartdeactivatedeventargs)|_Eigenschaft_ > WorksheetId|Ruft die Id des Arbeitsblatts, in dem das Diagramm deaktiviert wird.|1,8|
|[chartDeletedEventArgs](/javascript/api/excel/excel.chartdeletedeventargs)|_Eigenschaft_ > ChartId|Ruft die Id des Diagramms, der aus dem Arbeitsblatt gelöscht wird.|1,8|
|[chartDeletedEventArgs](/javascript/api/excel/excel.chartdeletedeventargs)|_Eigenschaft_ > Typ|Ruft den Typ des Ereignisses ab.|1,8|
|[chartDeletedEventArgs](/javascript/api/excel/excel.chartdeletedeventargs)|_Eigenschaft_ > WorksheetId|Ruft die Id des Arbeitsblatts, in dem das Diagramm gelöscht wird.|1,8|
|[chartDeletedEventArgs](/javascript/api/excel/excel.chartdeletedeventargs)|_Beziehung_ > Quelle|Ruft die Quelle des Ereignisses ab.|1,8|
|[chartLegendEntry](/javascript/api/excel/excel.chartlegendentry)|_Eigenschaft_ > height|Stellt die Höhe des die LegendEntry auf die Diagrammlegende dar. Schreibgeschützt.|1,8|
|[chartLegendEntry](/javascript/api/excel/excel.chartlegendentry)|_Eigenschaft_ > index|Stellt den Index der LegendEntry der Diagrammlegende dar. Schreibgeschützt.|1,8|
|[chartLegendEntry](/javascript/api/excel/excel.chartlegendentry)|_Eigenschaft_ > linken|Links vom ein LegendEntry Diagramm darstellt. Schreibgeschützt.|1,8|
|[chartLegendEntry](/javascript/api/excel/excel.chartlegendentry)|_Eigenschaft_ > oben|Im oberen Bereich des ein LegendEntry Diagramm darstellt. Schreibgeschützt.|1,8|
|[chartLegendEntry](/javascript/api/excel/excel.chartlegendentry)|_Eigenschaft_ > width|Die Breite des der LegendEntry im Diagramm Legende darstellt. Schreibgeschützt.|1,8|
|[chartLegendFormat](/javascript/api/excel/excel.chartlegendformat)|_Beziehung_ > Rahmen|Stellt das Rahmenformat, das Color, Linestyle und Weight enthält. Schreibgeschützt.|1,8|
|[chartPlotArea](/javascript/api/excel/excel.chartplotarea)|_Eigenschaft_ > height|Stellt den Wert für die Höhe des PlotArea dar.|1,8|
|[chartPlotArea](/javascript/api/excel/excel.chartplotarea)|_Eigenschaft_ > InsideHeight|Stellt den Wert InsideHeight PlotArea dar.|1,8|
|[chartPlotArea](/javascript/api/excel/excel.chartplotarea)|_Eigenschaft_ > InsideLeft|Stellt den Wert InsideLeft PlotArea dar.|1,8|
|[chartPlotArea](/javascript/api/excel/excel.chartplotarea)|_Eigenschaft_ > InsideTop|Stellt den Wert InsideTop PlotArea dar.|1,8|
|[chartPlotArea](/javascript/api/excel/excel.chartplotarea)|_Eigenschaft_ > InsideWidth|Stellt den Wert InsideWidth PlotArea dar.|1,8|
|[chartPlotArea](/javascript/api/excel/excel.chartplotarea)|_Eigenschaft_ > linken|Stellt den linken PlotArea-Wert.|1,8|
|[chartPlotArea](/javascript/api/excel/excel.chartplotarea)|_Eigenschaft_ > oben|Stellt den oberen Wert PlotArea dar.|1,8|
|[chartPlotArea](/javascript/api/excel/excel.chartplotarea)|_Eigenschaft_ > width|Den Breitenwert des PlotArea darstellt.|1,8|
|[chartPlotArea](/javascript/api/excel/excel.chartplotarea)|_Beziehung_ > Format|Stellt die Formatierung der ein Diagramm PlotArea. Schreibgeschützt.|1,8|
|[chartPlotArea](/javascript/api/excel/excel.chartplotarea)|_Beziehung_ > Position|Stellt die Position des PlotArea dar.|1,8|
|[chartPlotAreaFormat](/javascript/api/excel/excel.chartplotareaformat)|_Beziehung_ > Rahmen|Stellt die Rahmenattribute einer PlotArea Diagramm dar. Schreibgeschützt.|1,8|
|[chartPlotAreaFormat](/javascript/api/excel/excel.chartplotareaformat)|_Beziehung_ > Füllung|Stellt die Füllung eines Objekts dar, einschließlich Informationen zur Hintergrundformatierung. Schreibgeschützt.|1,8|
|[chartSeries](/javascript/api/excel/excel.chartseries)|_Eigenschaft_ > Explosion|Zurückgeben oder festlegen den Explosionswert für Kreis- oder Ringsegments Segments. Gibt 0 (null) zurück, wenn keine Explosion vorhanden ist (die Spitze des Segments ist in der Mitte des Diagramms). ReadWrite.|1,8|
|[chartSeries](/javascript/api/excel/excel.chartseries)|_Eigenschaft_ > FirstSliceAngle|Zurückgeben oder festlegen den Winkel des ersten Kreis- oder Ringsegments Segments in Grad (vertikal im Uhrzeigersinn). Gilt nur für Kreis-, 3-d Kreis- und Ringdiagrammen. Ein Wert zwischen 0 und 360 kann sein. ReadWrite|1,8|
|[chartSeries](/javascript/api/excel/excel.chartseries)|_Eigenschaft_ > InvertIfNegative|True, wenn Microsoft Excel das Muster im Element invertiert, wenn es eine negative Zahl entspricht. ReadWrite.|1,8|
|[chartSeries](/javascript/api/excel/excel.chartseries)|_Eigenschaft_ > überlappen|Gibt an, wie Balken und Spalten angeordnet sind. Ein Wert zwischen-100 und 100 kann sein. Gilt nur für 2D-Diagramme Balken und 2 2D-Säulendiagramme. ReadWrite.|1,8|
|[chartSeries](/javascript/api/excel/excel.chartseries)|_Eigenschaft_ > SecondPlotSize|Zurückgeben oder Festlegen der Größe des sekundären Abschnitts entweder ein Kreis-aus-Kreis oder Balken-aus-Kreis-Diagramms als Prozentsatz der Größe des primären Kreises. Ein Wert zwischen 5 und 200 kann sein. ReadWrite.|1,8|
|[chartSeries](/javascript/api/excel/excel.chartseries)|_Eigenschaft_ > VaryByCategories|True, wenn Microsoft Excel jeden Datenpunkt eine andere Farbe oder ein Muster zuweist. Das Diagramm darf nur eine Datenreihe enthalten. ReadWrite.|1,8|
|[chartSeries](/javascript/api/excel/excel.chartseries)|_Beziehung_ > AxisGroup|Zurückgeben oder festlegen die Gruppe für die angegebene Datenreihe. ReadWrite|1,8|
|[chartSeries](/javascript/api/excel/excel.chartseries)|_Beziehung_ > DataLabels|Eine Auflistung aller DataLabel in der Datenreihe darstellt. Schreibgeschützt.|1,8|
|[chartSeries](/javascript/api/excel/excel.chartseries)|_Beziehung_ > SplitType|Gibt zurück oder legt fest, die ein Kreis-aus-Kreis oder Balken-aus-Kreis-Diagramms geteilt werden wie. ReadWrite.|1,8|
|[chartTrendline](/javascript/api/excel/excel.charttrendline)|_Eigenschaft_ > BackwardPeriod|Stellt die Anzahl der Zeiträume, die die sich eine Trendlinie zurück erstreckt.|1,8|
|[chartTrendline](/javascript/api/excel/excel.charttrendline)|_Eigenschaft_ > ForwardPeriod|Stellt die Anzahl der Zeiträume, die die sich eine Trendlinie vorwärts erstreckt.|1,8|
|[chartTrendline](/javascript/api/excel/excel.charttrendline)|_Eigenschaft_ > ShowEquation|True, wenn die Formel für die sich eine Trendlinie im Diagramm angezeigt wird.|1,8|
|[chartTrendline](/javascript/api/excel/excel.charttrendline)|_Eigenschaft_ > ShowRSquared|True, wenn der R-quadrierten für die sich eine Trendlinie im Diagramm angezeigt wird.|1,8|
|[chartTrendline](/javascript/api/excel/excel.charttrendline)|_Beziehung_ > Beschriftung|Repräsentiert die Beschriftung einer Trendlinie Diagramm. Schreibgeschützt.|1,8|
|[chartTrendlineLabel](/javascript/api/excel/excel.charttrendlinelabel)|_Eigenschaft_ > AutoText|Boolescher Wert, wenn Trendlinie Bezeichnung automatisch basierend auf Kontext angemessenen Text generiert.|1,8|
|[chartTrendlineLabel](/javascript/api/excel/excel.charttrendlinelabel)|_Eigenschaft_ > Formel|String-Wert, der die Formel des Diagramms Trendlinie Bezeichnung unter Verwendung der A1-Schreibweise darstellt.|1,8|
|[chartTrendlineLabel](/javascript/api/excel/excel.charttrendlinelabel)|_Eigenschaft_ > height|Gibt die Höhe des Diagramms Trendlinie Beschriftung in Punkt zurück. Schreibgeschützt. NULL, wenn das Diagramm Trendlinie Bezeichnung nicht sichtbar ist. Schreibgeschützt.|1,8|
|[chartTrendlineLabel](/javascript/api/excel/excel.charttrendlinelabel)|_Eigenschaft_ > linken|Stellt den Abstand in Punkt vom linken Rand des Diagramms Trendlinie Beschriftung zum linken Rand des Diagrammbereichs dar. NULL, wenn das Diagramm Trendlinie Bezeichnung nicht sichtbar ist.|1,8|
|[chartTrendlineLabel](/javascript/api/excel/excel.charttrendlinelabel)|_Eigenschaft_ > numberFormat|String-Wert, der den Formatierungscode für Bezeichnung Trendlinie darstellt.|1,8|
|[chartTrendlineLabel](/javascript/api/excel/excel.charttrendlinelabel)|_Eigenschaft_ > text|Eine Zeichenfolge, die den Text der Beschriftung Trendlinie in einem Diagramm darstellt.|1,8|
|[chartTrendlineLabel](/javascript/api/excel/excel.charttrendlinelabel)|_Eigenschaft_ > TextOrientation|Stellt die Ausrichtung von Text des Diagramms Trendlinie Beschriftung. Der Wert muss eine ganze Zahl sein entweder von-90 bis 90 oder 180 für Text vertikal ausgerichtet.|1,8|
|[chartTrendlineLabel](/javascript/api/excel/excel.charttrendlinelabel)|_Eigenschaft_ > oben|Stellt den Abstand in Punkt vom oberen Rand des Diagramms Trendlinie Beschriftung an den Anfang der Diagrammfläche. NULL, wenn das Diagramm Trendlinie Bezeichnung nicht sichtbar ist.|1,8|
|[chartTrendlineLabel](/javascript/api/excel/excel.charttrendlinelabel)|_Eigenschaft_ > width|Gibt die Breite des Diagramms Trendlinie Beschriftung in Punkt zurück. Schreibgeschützt. NULL, wenn das Diagramm Trendlinie Bezeichnung nicht sichtbar ist. Schreibgeschützt.|1,8|
|[chartTrendlineLabel](/javascript/api/excel/excel.charttrendlinelabel)|_Beziehung_ > Format|Stellt das Format des Diagramms Trendlinie Beschriftung. Schreibgeschützt.|1,8|
|[chartTrendlineLabel](/javascript/api/excel/excel.charttrendlinelabel)|_Beziehung_ > horizontalAlignment|Die horizontale Ausrichtung für Diagramm Trendlinie Beschriftung repräsentiert.|1,8|
|[chartTrendlineLabel](/javascript/api/excel/excel.charttrendlinelabel)|_Beziehung_ > verticalAlignment|Stellt die vertikale Ausrichtung des Diagramms Trendlinie Beschriftung.|1,8|
|[chartTrendlineLabelFormat](/javascript/api/excel/excel.charttrendlinelabelformat)|_Beziehung_ > Rahmen|Stellt das Rahmenformat, das Color, Linestyle und Weight enthält. Schreibgeschützt.|1,8|
|[chartTrendlineLabelFormat](/javascript/api/excel/excel.charttrendlinelabelformat)|_Beziehung_ > Füllung|Stellt das Füllformat des aktuellen Texts Diagramm Trendlinie dar. Schreibgeschützt.|1,8|
|[chartTrendlineLabelFormat](/javascript/api/excel/excel.charttrendlinelabelformat)|_Beziehung_ > font|Stellt die Schriftattribute (Name der Schriftart, Schriftgrad, Farbe usw.) für ein Diagramm Trendlinie Label. Schreibgeschützt.|1,8|
|[createWorkbookPostProcessAction](/javascript/api/excel/excel.createworkbookpostprocessaction)|_Eigenschaft_ > FakeFileId|Überträgt zusätzliche Daten an den Client-seitigen, z. B. WorksheetId für TableSelectionChangedEvent.|1,8|
|[createWorkbookPostProcessAction](/javascript/api/excel/excel.createworkbookpostprocessaction)|_Eigenschaft_ > fileBase64|Überträgt zusätzliche Daten an den Client-seitigen, z. B. WorksheetId für TableSelectionChangedEvent.|1,8|
|[createWorkbookPostProcessAction](/javascript/api/excel/excel.createworkbookpostprocessaction)|_Beziehung_ > ActionType|Überträgt zusätzliche Daten an den Client-seitigen, z. B. WorksheetId für TableSelectionChangedEvent.|1,8|
|[customDataValidation](/javascript/api/excel/excel.customdatavalidation)|_Eigenschaft_ > Formel| Eine benutzerdefinierte Daten Validation-Formel. Dadurch wird die spezielle Eingaberegeln, z. B. verhindern, dass Duplikate oder Einschränken der Summe in einen Bereich von Tabellenzellen erstellt.|1,8|
|[dataPivotHierarchy](/javascript/api/excel/excel.datapivothierarchy)|_Eigenschaft_ > id|ID des der DataPivotHierarchy. Schreibgeschützt.|1,8|
|[dataPivotHierarchy](/javascript/api/excel/excel.datapivothierarchy)|_Eigenschaft_ > name|Name der DataPivotHierarchy.|1,8|
|[dataPivotHierarchy](/javascript/api/excel/excel.datapivothierarchy)|_Eigenschaft_ > numberFormat|Von der DataPivotHierarchy Zahlenformat.|1,8|
|[dataPivotHierarchy](/javascript/api/excel/excel.datapivothierarchy)|_Eigenschaft_ > Position|Die Position für das DataPivotHierarchy.|1,8|
|[dataPivotHierarchy](/javascript/api/excel/excel.datapivothierarchy)|_Beziehung_ > Feld|Gibt die PivotFields der DataPivotHierarchy zugeordnet. Schreibgeschützt.|1,8|
|[dataPivotHierarchy](/javascript/api/excel/excel.datapivothierarchy)|_Beziehung_ > ShowAs|Bestimmt, ob die Daten als eine bestimmte zusammenfassende Berechnung soll oder nicht angezeigt werden.|1,8|
|[dataPivotHierarchy](/javascript/api/excel/excel.datapivothierarchy)|_Beziehung_ > SummarizeBy|Bestimmt, ob alle Elemente der DataPivotHierarchy anzeigen.|1,8|
|[dataPivotHierarchy](/javascript/api/excel/excel.datapivothierarchy)|_Methode_ > [setToDefault()](/javascript/api/excel/excel.datapivothierarchy#settodefault)|Die DataPivotHierarchy auf deren Standardwerte zurückgesetzt.|1,8|
|[dataPivotHierarchyCollection](/javascript/api/excel/excel.datapivothierarchycollection)|_Eigenschaft_ > items|Eine Auflistung von DataPivotHierarchy-Objekten. Schreibgeschützt.|1,8|
|[dataPivotHierarchyCollection](/javascript/api/excel/excel.datapivothierarchycollection)|_Methode_ > [Add(pivotHierarchy: PivotHierarchy)](/javascript/api/excel/excel.datapivothierarchycollection)|Die aktuelle Achse hinzugefügt der PivotHierarchy.|1,8|
|[dataPivotHierarchyCollection](/javascript/api/excel/excel.datapivothierarchycollection)|_Methode_ > [getCount()](/javascript/api/excel/excel.datapivothierarchycollection)|Ruft die Anzahl der Pivot-Hierarchien in der Auflistung ab.|1,8|
|[dataPivotHierarchyCollection](/javascript/api/excel/excel.datapivothierarchycollection)|_Methode_ > [getItem(name: string)](/javascript/api/excel/excel.datapivothierarchycollection)|Ruft eine DataPivotHierarchy nach dessen Name oder die Id ab.|1,8|
|[dataPivotHierarchyCollection](/javascript/api/excel/excel.datapivothierarchycollection)|_Methode_ > [getItemOrNullObject(name: string)](/javascript/api/excel/excel.datapivothierarchycollection)|Ruft eine DataPivotHierarchy nach Namen ab. Wenn die DataPivotHierarchy nicht vorhanden ist, wird ein null-Objekt zurückzugeben.|1,8|
|[dataPivotHierarchyCollection](/javascript/api/excel/excel.datapivothierarchycollection)|_Methode_ > [Remove(DataPivotHierarchy: DataPivotHierarchy)](/javascript/api/excel/excel.datapivothierarchycollection)|Entfernt die PivotHierarchy aus der aktuellen Achse an.|1,8|
|[dataValidation](/javascript/api/excel/excel.datavalidation)|_Eigenschaft_ > IgnoreBlanks|Leerzeichen ignorieren: keine datenüberprüfung auf leere Zellen erfolgt, wird standardmäßig auf "true".|1,8|
|[dataValidation](/javascript/api/excel/excel.datavalidation)|_Eigenschaft_ > ungültig|Stellt dar, wenn alle Zellwerte entsprechend der Regeln für die Validierung gültig sind. Schreibgeschützt.|1,8|
|[dataValidation](/javascript/api/excel/excel.datavalidation)|_Beziehung_ > ErrorAlert|Fehlermeldung, wenn Benutzer ungültige Daten eingibt.|1,8|
|[dataValidation](/javascript/api/excel/excel.datavalidation)|_Beziehung_ > Aufforderung|Aufforderung, wenn Benutzer eine Zelle auswählt.|1,8|
|[dataValidation](/javascript/api/excel/excel.datavalidation)|_Beziehung_ > Regel|Daten Überprüfungsregel, die verschiedene Typen von datenüberprüfungskriterien enthält.|1,8|
|[dataValidation](/javascript/api/excel/excel.datavalidation)|_Beziehung_ > type|Geben Sie die Überprüfung, [Excel.DataValidationType](/javascript/api/excel/excel.datavalidationtype) Einzelheiten finden Sie unter. Schreibgeschützt.|1,8|
|[dataValidation](/javascript/api/excel/excel.datavalidation)|_Methode_ > [clear()](/javascript/api/excel/excel.datavalidation)|Löscht die Überprüfung aus dem aktuellen Bereich.|1,8|
|[dataValidationErrorAlert](/javascript/api/excel/excel.datavalidationerroralert)|_Eigenschaft_ > Nachricht|Fehlermeldung darstellt.|1,8|
|[dataValidationErrorAlert](/javascript/api/excel/excel.datavalidationerroralert)|_Eigenschaft_ > ShowAlert|Bestimmt, ob für eine Warnung Fehlerdialogfeld angezeigt werden oder nicht, wenn der Benutzer ungültige Daten eingibt. Der Standardwert ist true.|1,8|
|[dataValidationErrorAlert](/javascript/api/excel/excel.datavalidationerroralert)|_Eigenschaft_ > title|Fehler, Warnung Dialogfeldtitel darstellt.|1,8|
|[dataValidationErrorAlert](/javascript/api/excel/excel.datavalidationerroralert)|_Beziehung_ > Formatvorlage|Stellt die datenüberprüfung Benachrichtigungstyp, finden Sie ausführliche [Excel.DataValidationAlertStyle](/javascript/api/excel/excel.datavalidationalertstyle) .|1,8|
|[dataValidationPrompt](/javascript/api/excel/excel.datavalidationprompt)|_Eigenschaft_ > Nachricht|Stellt die Nachricht der Ansage dar.|1,8|
|[dataValidationPrompt](/javascript/api/excel/excel.datavalidationprompt)|_Eigenschaft_ > ShowPrompt|Bestimmt, ob die Meldung wird angezeigt, wenn Benutzer eine Zelle mit einer Validierung auswählt.|1,8|
|[dataValidationPrompt](/javascript/api/excel/excel.datavalidationprompt)|_Eigenschaft_ > title|Stellt den Titel für die Aufforderung zur dar.|1,8|
|[dataValidationRule](/javascript/api/excel/excel.datavalidationrule)|_Beziehung_ > benutzerdefinierte|Benutzerdefinierte datenüberprüfungskriterien.|1,8|
|[dataValidationRule](/javascript/api/excel/excel.datavalidationrule)|_Beziehung_ > Datum|Datum datenüberprüfungskriterien.|1,8|
|[dataValidationRule](/javascript/api/excel/excel.datavalidationrule)|_Beziehung_ > decimal|Gültigkeitskriterien Decimal-Datentyp.|1,8|
|[dataValidationRule](/javascript/api/excel/excel.datavalidationrule)|_Beziehung_ > list|Liste datenüberprüfungskriterien.|1,8|
|[dataValidationRule](/javascript/api/excel/excel.datavalidationrule)|_Beziehung_ > TextLength|TextLength datenüberprüfungskriterien.|1,8|
|[dataValidationRule](/javascript/api/excel/excel.datavalidationrule)|_Beziehung_ > Zeit|Zeit datenüberprüfungskriterien.|1,8|
|[dataValidationRule](/javascript/api/excel/excel.datavalidationrule)|_Beziehung_ > WholeNumber|WholeNumber datenüberprüfungskriterien.|1,8|
|[dateTimeDataValidation](/javascript/api/excel/excel.datetimedatavalidation)|_Eigenschaft_ > formula1|Ruft ab oder legt diesen fest Formula1, d. h. Mindestwert oder Wert je nach den Operator.|1,8|
|[dateTimeDataValidation](/javascript/api/excel/excel.datetimedatavalidation)|_Eigenschaft_ > formula2|Ruft ab oder legt diesen fest Formula2, d. h. Höchstwert oder Wert je nach den Operator.|1,8|
|[dateTimeDataValidation](/javascript/api/excel/excel.datetimedatavalidation)|_Beziehung_ > Operator|Der Operator für die Validierung der Daten verwendet.|1,8|
|[enableEventsPostProcessAction](/javascript/api/excel/excel.enableeventspostprocessaction)|_Eigenschaft_ > IsEnableEvents {|Überträgt zusätzliche Daten an den Client-seitigen, z. B. WorksheetId für TableSelectionChangedEvent.|1,8|
|[enableEventsPostProcessAction](/javascript/api/excel/excel.enableeventspostprocessaction)|_Beziehung_ > ActionType|Überträgt zusätzliche Daten an den Client-seitigen, z. B. WorksheetId für TableSelectionChangedEvent.|1,8|
|[enableEventsPostProcessAction](/javascript/api/excel/excel.enableeventspostprocessaction)|_Beziehung_ > ControlId|Überträgt zusätzliche Daten an den Client-seitigen, z. B. WorksheetId für TableSelectionChangedEvent.|1,8|
|[filterPivotHierarchy](/javascript/api/excel/excel.filterpivothierarchy)|_Eigenschaft_ > EnableMultipleFilterItems|Bestimmt, ob mehrere Filtern von Elementen zu ermöglichen.|1,8|
|[filterPivotHierarchy](/javascript/api/excel/excel.filterpivothierarchy)|_Eigenschaft_ > id|ID des der FilterPivotHierarchy. Schreibgeschützt.|1,8|
|[filterPivotHierarchy](/javascript/api/excel/excel.filterpivothierarchy)|_Eigenschaft_ > name|Name der FilterPivotHierarchy.|1,8|
|[filterPivotHierarchy](/javascript/api/excel/excel.filterpivothierarchy)|_Eigenschaft_ > Position|Die Position für das FilterPivotHierarchy.|1,8|
|[filterPivotHierarchy](/javascript/api/excel/excel.filterpivothierarchy)|_Beziehung_ > fields|Gibt die PivotFields der FilterPivotHierarchy zugeordnet. Schreibgeschützt.|1,8|
|[filterPivotHierarchy](/javascript/api/excel/excel.filterpivothierarchy)|_Methode_ > [setToDefault()](/javascript/api/excel/excel.filterpivothierarchy)|Die FilterPivotHierarchy auf deren Standardwerte zurückgesetzt.|1,8|
|[filterPivotHierarchyCollection](/javascript/api/excel/excel.filterpivothierarchycollection)|_Eigenschaft_ > items|Eine Auflistung von FilterPivotHierarchy-Objekten. Schreibgeschützt.|1,8|
|[filterPivotHierarchyCollection](/javascript/api/excel/excel.filterpivothierarchycollection)|_Methode_ > [Add(pivotHierarchy: PivotHierarchy)](/javascript/api/excel/excel.filterpivothierarchycollection)|Die aktuelle Achse hinzugefügt der PivotHierarchy. Wenn die Hierarchie an anderer Stelle auf der Zeilen-, Spalten- oder Filterachse vorhanden ist, wird es aus diesem Speicherort entfernt werden.|1,8|
|[filterPivotHierarchyCollection](/javascript/api/excel/excel.filterpivothierarchycollection)|_Methode_ > [getCount()](/javascript/api/excel/excel.filterpivothierarchycollection)|Ruft die Anzahl der Pivot-Hierarchien in der Auflistung ab.|1,8|
|[filterPivotHierarchyCollection](/javascript/api/excel/excel.filterpivothierarchycollection)|_Methode_ > [getItem(name: string)](/javascript/api/excel/excel.filterpivothierarchycollection)|Ruft eine FilterPivotHierarchy nach dessen Name oder die Id ab.|1,8|
|[filterPivotHierarchyCollection](/javascript/api/excel/excel.filterpivothierarchycollection)|_Methode_ > [getItemOrNullObject(name: string)](/javascript/api/excel/excel.filterpivothierarchycollection)|Ruft eine FilterPivotHierarchy nach Namen ab. Wenn die FilterPivotHierarchy nicht vorhanden ist, wird ein null-Objekt zurückzugeben.|1,8|
|[filterPivotHierarchyCollection](/javascript/api/excel/excel.filterpivothierarchycollection)|_Methode_ > [Remove(filterPivotHierarchy: FilterPivotHierarchy)](/javascript/api/excel/excel.filterpivothierarchycollection)|Entfernt die PivotHierarchy aus der aktuellen Achse an.|1,8|
|[listDataValidation](/javascript/api/excel/excel.listdatavalidation)|_Eigenschaft_ > InCellDropDown|Zeigt die Liste in Zelle Dropdown-Liste oder nicht, wird standardmäßig auf "true".|1,8|
|[listDataValidation](/javascript/api/excel/excel.listdatavalidation)|_Eigenschaft_ > Quelle|Quelle der Liste für die datenüberprüfung|1,8|
|[openWorkbookPostProcessAction](/javascript/api/excel/excel.openworkbookpostprocessaction)|_Eigenschaft_ > FakeFileId|Überträgt zusätzliche Daten an den Client-seitigen, z. B. WorksheetId für TableSelectionChangedEvent.|1,8|
|[openWorkbookPostProcessAction](/javascript/api/excel/excel.openworkbookpostprocessaction)|_Beziehung_ > ActionType|Überträgt zusätzliche Daten an den Client-seitigen, z. B. WorksheetId für TableSelectionChangedEvent.|1,8|
|[pivotField](/javascript/api/excel/excel.pivotfield)|_Eigenschaft_ > id|ID des PivotFields. Schreibgeschützt.|1,8|
|[pivotField](/javascript/api/excel/excel.pivotfield)|_Eigenschaft_ > name|Name des PivotFields.|1,8|
|[pivotField](/javascript/api/excel/excel.pivotfield)|_Eigenschaft_ > ShowAllItems|Bestimmt, ob alle Elemente des PivotFields anzuzeigen.|1,8|
|[pivotField](/javascript/api/excel/excel.pivotfield)|_Beziehung_ > Elemente|Gibt die PivotFields PivotFields zugeordnet. Schreibgeschützt.|1,8|
|[pivotField](/javascript/api/excel/excel.pivotfield)|_Beziehung_ > Teilergebnisse|Teilergebnisse des PivotFields.|1,8|
|[pivotField](/javascript/api/excel/excel.pivotfield)|_Methode_ > [SortByLabels(sortby: SortBy)](/javascript/api/excel/excel.pivotfield)|Sortiert das PivotField. Wenn eine DataPivotHierarchy angegeben wird, klicken Sie dann sortieren angewendet wird basierend auf dem, wenn nicht Sortierung basiert auf das PivotField selbst.|1,8|
|[pivotFieldCollection](/javascript/api/excel/excel.pivotfieldcollection)|_Eigenschaft_ > items|Eine Auflistung von PivotField-Objekten. Schreibgeschützt.|1,8|
|[pivotFieldCollection](/javascript/api/excel/excel.pivotfieldcollection)|_Methode_ > [getCount()](/javascript/api/excel/excel.pivotfieldcollection)|Ruft die Anzahl der Pivot-Hierarchien in der Auflistung ab.|1,8|
|[pivotFieldCollection](/javascript/api/excel/excel.pivotfieldcollection)|_Methode_ > [getItem(name: string)](/javascript/api/excel/excel.pivotfieldcollection)|Ruft eine PivotHierarchy nach dessen Name oder die Id ab.|1,8|
|[pivotFieldCollection](/javascript/api/excel/excel.pivotfieldcollection)|_Methode_ > [getItemOrNullObject(name: string)](/javascript/api/excel/excel.pivotfieldcollection)|Ruft eine PivotHierarchy nach Namen ab. Wenn die PivotHierarchy nicht vorhanden ist, wird ein null-Objekt zurückzugeben.|1,8|
|[pivotHierarchy](/javascript/api/excel/excel.pivothierarchy)|_Eigenschaft_ > id|ID des der PivotHierarchy. Schreibgeschützt.|1,8|
|[pivotHierarchy](/javascript/api/excel/excel.pivothierarchy)|_Eigenschaft_ > name|Name der PivotHierarchy.|1,8|
|[pivotHierarchy](/javascript/api/excel/excel.pivothierarchy)|_Beziehung_ > fields|Gibt die PivotFields der PivotHierarchy zugeordnet. Schreibgeschützt.|1,8|
|[pivotHierarchyCollection](/javascript/api/excel/excel.pivothierarchycollection)|_Eigenschaft_ > items|Eine Auflistung von PivotHierarchy-Objekten. Schreibgeschützt.|1,8|
|[pivotHierarchyCollection](/javascript/api/excel/excel.pivothierarchycollection)|_Methode_ > [getCount()](/javascript/api/excel/excel.pivothierarchycollection)|Ruft die Anzahl der Pivot-Hierarchien in der Auflistung ab.|1,8|
|[pivotHierarchyCollection](/javascript/api/excel/excel.pivothierarchycollection)|_Methode_ > [getItem(name: string)](/javascript/api/excel/excel.pivothierarchycollection)|Ruft eine PivotHierarchy nach dessen Name oder die Id ab.|1,8|
|[pivotHierarchyCollection](/javascript/api/excel/excel.pivothierarchycollection)|_Methode_ > [getItemOrNullObject(name: string)](/javascript/api/excel/excel.pivothierarchycollection)|Ruft eine PivotHierarchy nach Namen ab. Wenn die PivotHierarchy nicht vorhanden ist, wird ein null-Objekt zurückzugeben.|1,8|
|[pivotItem](/javascript/api/excel/excel.pivotitem)|_Eigenschaft_ > id|ID des PivotItem-Objekts. Schreibgeschützt.|1,8|
|[pivotItem](/javascript/api/excel/excel.pivotitem)|_Eigenschaft_ > IsExpanded|Bestimmt, ob das Element erweitert wird, um untergeordnete Elemente anzuzeigen, oder wenn es ausgeblendet ist und untergeordnete Elemente werden ausgeblendet.|1,8|
|[pivotItem](/javascript/api/excel/excel.pivotitem)|_Eigenschaft_ > name|Name des PivotItem-Objekts.|1,8|
|[pivotItem](/javascript/api/excel/excel.pivotitem)|_Eigenschaft_ > sichtbar|Bestimmt, ob PivotItem-Objekts angezeigt wird.|1,8|
|[pivotItemCollection](/javascript/api/excel/excel.pivotitemcollection)|_Eigenschaft_ > items|Eine Auflistung von PivotItem-Objekte. Schreibgeschützt.|1,8|
|[pivotItemCollection](/javascript/api/excel/excel.pivotitemcollection)|_Methode_ > [getCount()](/javascript/api/excel/excel.pivotitemcollection)|Ruft die Anzahl der Pivot-Hierarchien in der Auflistung ab.|1,8|
|[pivotItemCollection](/javascript/api/excel/excel.pivotitemcollection)|_Methode_ > [getItem(name: string)](/javascript/api/excel/excel.pivotitemcollection)|Ruft eine PivotHierarchy nach dessen Name oder die Id ab.|1,8|
|[pivotItemCollection](/javascript/api/excel/excel.pivotitemcollection)|_Methode_ > [getItemOrNullObject(name: string)](/javascript/api/excel/excel.pivotitemcollection)|Ruft eine PivotHierarchy nach Namen ab. Wenn die PivotHierarchy nicht vorhanden ist, wird ein null-Objekt zurückzugeben.|1,8|
|[pivotLayout](/javascript/api/excel/excel.pivotlayout)|_Eigenschaft_ > ShowColumnGrandTotals|True, wenn der PivotTable-Bericht Gesamtergebnisse für Spalten summiert.|1,8|
|[pivotLayout](/javascript/api/excel/excel.pivotlayout)|_Eigenschaft_ > ShowRowGrandTotals|True, wenn der PivotTable-Bericht Gesamtergebnisse für Zeilen angezeigt.|1,8|
|[pivotLayout](/javascript/api/excel/excel.pivotlayout)|_Eigenschaft_ > SubtotalLocation|Diese Eigenschaft gibt die SubtotalLocationType aller Felder in der PivotTable an. Wenn Felder unterschiedlichen Zustände haben, wird dieser null sein. Mögliche Werte sind: AtTop, AtBottom.|1,8|
|[pivotLayout](/javascript/api/excel/excel.pivotlayout)|_Beziehung_ > LayoutType|Diese Eigenschaft gibt die PivotLayoutType aller Felder in der PivotTable an. Wenn Felder unterschiedlichen Zustände haben, wird dieser null sein.|1,8|
|[pivotLayout](/javascript/api/excel/excel.pivotlayout)|_Methode_ > [getColumnLabelRange()](/javascript/api/excel/excel.pivotlayout)|Gibt den Bereich, in denen die PivotTable Spaltenüberschriften befinden.|1,8|
|[pivotLayout](/javascript/api/excel/excel.pivotlayout)|_Methode_ > [getDataBodyRange()](/javascript/api/excel/excel.pivotlayout)|Gibt den Bereich, in denen die PivotTable Datenwerte befinden.|1,8|
|[pivotLayout](/javascript/api/excel/excel.pivotlayout.md)|_Methode_ > [getFilterAxisRange()](/javascript/api/excel/excel.pivotlayout)|Gibt den Bereich der Filterbereich der PivotTable zurück.|1,8|
|[pivotLayout](/javascript/api/excel/excel.pivotlayout)|_Methode_ > [getRange()](/javascript/api/excel/excel.pivotlayout)|Gibt den Bereich die PivotTable, vorhanden ist, ausgenommen Filterbereich zurück.|1,8|
|[pivotLayout](/javascript/api/excel/excel.pivotlayout)|_Methode_ > [getRowLabelRange()](/javascript/api/excel/excel.pivotlayout)|Gibt den Bereich, in denen die PivotTable zeilenbeschriftungen befinden.|1,8|
|[pivotTable](/javascript/api/excel/excel.pivottable)|_Beziehung_ > ColumnHierarchies|Die Spalte Pivot Hierarchien der PivotTable. Schreibgeschützt.|1,8|
|[pivotTable](/javascript/api/excel/excel.pivottable)|_Beziehung_ > DataHierarchies|Die Daten Pivot Hierarchien der PivotTable. Schreibgeschützt.|1,8|
|[pivotTable](/javascript/api/excel/excel.pivottable)|_Beziehung_ > FilterHierarchies|Die Filter Pivot Hierarchien der PivotTable. Schreibgeschützt.|1,8|
|[pivotTable](/javascript/api/excel/excel.pivottable)|_Beziehung_ > Hierarchien|Die Pivot-Hierarchien der PivotTable. Schreibgeschützt.|1,8|
|[pivotTable](/javascript/api/excel/excel.pivottable)|_Beziehung_ > Layout|Die PivotLayout beschreiben das Layout und die visuelle Struktur der PivotTable. Schreibgeschützt.|1,8|
|[pivotTable](/javascript/api/excel/excel.pivottable)|_Beziehung_ > RowHierarchies|Die Zeile Pivot Hierarchien der PivotTable. Schreibgeschützt.|1,8|
|[pivotTable](/javascript/api/excel/excel.pivottable)|_Methode_ > [delete()](/javascript/api/excel/excel.pivottable)|Löscht die PivotTable.|1,8|
|[pivotTableCollection](/javascript/api/excel/excel.pivottablecollection)|_Methode_ > [Add(name: string, source: object, destination: object)](/javascript/api/excel/excel.pivottablecollection)|Hinzufügen einer Pivottable basierend auf den angegebenen Quelldaten und fügt ihn an der linken oberen Zelle des Zielbereichs.|1,8|
|[range](/javascript/api/excel/excel.range)|_Beziehung_ > DataValidation|Gibt ein Data Validation-Objekt zurück. Schreibgeschützt.|1,8|
|[rowColumnPivotHierarchy](/javascript/api/excel/excel.rowcolumnpivothierarchy)|_Eigenschaft_ > id|ID des der RowColumnPivotHierarchy. Schreibgeschützt.|1,8|
|[rowColumnPivotHierarchy](/javascript/api/excel/excel.rowcolumnpivothierarchy)|_Eigenschaft_ > name|Name der RowColumnPivotHierarchy.|1,8|
|[rowColumnPivotHierarchy](/javascript/api/excel/excel.rowcolumnpivothierarchy)|_Eigenschaft_ > Position|Die Position für das RowColumnPivotHierarchy.|1,8|
|[rowColumnPivotHierarchy](/javascript/api/excel/excel.rowcolumnpivothierarchy)|_Beziehung_ > fields|Gibt die PivotFields der RowColumnPivotHierarchy zugeordnet. Schreibgeschützt.|1,8|
|[rowColumnPivotHierarchy](/javascript/api/excel/excel.rowcolumnpivothierarchy)|_Methode_ > [setToDefault()](/javascript/api/excel/excel.rowcolumnpivothierarchy)|Die RowColumnPivotHierarchy auf deren Standardwerte zurückgesetzt.|1,8|
|[rowColumnPivotHierarchyCollection](/javascript/api/excel/excel.rowcolumnpivothierarchycollection)|_Eigenschaft_ > items|Eine Auflistung von RowColumnPivotHierarchy-Objekten. Schreibgeschützt.|1,8|
|[rowColumnPivotHierarchyCollection](/javascript/api/excel/excel.rowcolumnpivothierarchycollection)|_Methode_ > [Add(pivotHierarchy: PivotHierarchy)](/javascript/api/excel/excel.rowcolumnpivothierarchycollection)|Die aktuelle Achse hinzugefügt der PivotHierarchy. Ist die Hierarchie vorhanden ist, an anderer Stelle in der Zeile, Spalte|1,8|
|[rowColumnPivotHierarchyCollection](/javascript/api/excel/excel.rowcolumnpivothierarchycollection)|_Methode_ > [getCount()](/javascript/api/excel/excel.rowcolumnpivothierarchycollection)|Ruft die Anzahl der Pivot-Hierarchien in der Auflistung ab.|1,8|
|[rowColumnPivotHierarchyCollection](/javascript/api/excel/excel.rowcolumnpivothierarchycollection)|_Methode_ > [getItem(name: string)](/javascript/api/excel/excel.rowcolumnpivothierarchycollection)|Ruft eine RowColumnPivotHierarchy nach dessen Name oder die Id ab.|1,8|
|[rowColumnPivotHierarchyCollection](/javascript/api/excel/excel.rowcolumnpivothierarchycollection)|_Methode_ > [getItemOrNullObject(name: string)](/javascript/api/excel/excel.rowcolumnpivothierarchycollection)|Ruft eine RowColumnPivotHierarchy nach Namen ab. Wenn die RowColumnPivotHierarchy nicht vorhanden ist, wird ein null-Objekt zurückzugeben.|1,8|
|[rowColumnPivotHierarchyCollection](/javascript/api/excel/excel.rowcolumnpivothierarchycollection)|_Methode_ > [Remove(rowColumnPivotHierarchy: RowColumnPivotHierarchy)](/javascript/api/excel/excel.rowcolumnpivothierarchycollection)|Entfernt die PivotHierarchy aus der aktuellen Achse an.|1,8|
|[Laufzeit](/javascript/api/excel/excel.runtime)|_Eigenschaft_ > EnableEvents|Umschalten von JavaScript-Ereignisse in der aktuellen Taskpane oder Content-add-in.|1,8|
|[showAsRule](/javascript/api/excel/excel.showasrule)|_Beziehung_ > BaseField|Das base PivotField die Berechnung ShowAs basieren soll, falls zutreffend basierend auf dem ShowAsCalculation, andernfalls null.|1,8|
|[showAsRule](/javascript/api/excel/excel.showasrule)|_Beziehung_ > BaseItem|Basiselements auf die Berechnung ShowAs, falls zutreffend basieren Grundlage auf der ShowAsCalculation, andernfalls null.|1,8|
|[showAsRule](/javascript/api/excel/excel.showasrule)|_Beziehung_ > Berechnung|Die ShowAs Berechnung, die für die Daten-PivotField verwendet.|1,8|
|[Formatvorlage](/javascript/api/excel/excel.style)|_Eigenschaft_ > AutoIndent|Gibt an, ob Text automatisch eingezogen wird, wenn die Ausrichtung des Texts in einer Zelle auf gleiche Verteilung festgelegt ist.|1,8|
|[Formatvorlage](/javascript/api/excel/excel.style)|_Eigenschaft_ > TextOrientation|Die Ausrichtung von Text für die Formatvorlage.|1,8|
|[Teilergebnisse](/javascript/api/excel/excel.subtotals)|_Eigenschaft_ > Automatische|Wenn automatisch auf "true", und klicken Sie dann alle anderen Werte festgelegt ist wird ignoriert werden, wenn die Teilergebnisse festlegen.|1,8|
|[Teilergebnisse](/javascript/api/excel/excel.subtotals)|_Eigenschaft_ > durchschnittliche| |1,8|
|[Teilergebnisse](/javascript/api/excel/excel.subtotals)|_Eigenschaft_ > Count| |1,8|
|[Teilergebnisse](/javascript/api/excel/excel.subtotals)|_Eigenschaft_ > CountNumbers| |1,8|
|[Teilergebnisse](/javascript/api/excel/excel.subtotals)|_Eigenschaft_ > max| |1,8|
|[Teilergebnisse](/javascript/api/excel/excel.subtotals)|_Eigenschaft_ > min| |1,8|
|[Teilergebnisse](/javascript/api/excel/excel.subtotals)|_Eigenschaft_ > Produkt| |1,8|
|[Teilergebnisse](/javascript/api/excel/excel.subtotals)|_Eigenschaft_ > StandardDeviation| |1,8|
|[Teilergebnisse](/javascript/api/excel/excel.subtotals)|_Eigenschaft_ > StandardDeviationP| |1,8|
|[Teilergebnisse](/javascript/api/excel/excel.subtotals)|_Eigenschaft_ > Summe| |1,8|
|[Teilergebnisse](/javascript/api/excel/excel.subtotals)|_Eigenschaft_ > Abweichung| |1,8|
|[Teilergebnisse](/javascript/api/excel/excel.subtotals)|_Eigenschaft_ > VarianceP| |1,8|
|[Tabelle](/javascript/api/excel/excel.table)|_Eigenschaft_ > LegacyId|Gibt eine numerische Id zurück. Schreibgeschützt.|1,8|
|[workbook](/javascript/api/excel/excel.workbook)|_Eigenschaft_ > ReadOnly|True, wenn die Arbeitsmappe im schreibgeschützten Modus geöffnet ist. Schreibgeschützt.|1,8|
|[workbookCreated](/javascript/api/excel/excel.workbookcreated)|_Eigenschaft_ > id|Gibt einen Wert, der das WorkbookCreated-Objekt eindeutig identifiziert. Schreibgeschützt.|1,8|
|[workbookCreated](/javascript/api/excel/excel.workbookcreated)|_Methode_ > [open()](/javascript/api/excel/excel.workbookcreated)|Öffnen Sie die Arbeitsmappe.|1,8|
|[worksheet](/javascript/api/excel/excel.worksheet)|_Eigenschaft_ > ShowGridlines|Ruft ab oder legt diesen fest im Arbeitsblatt Gitternetzlinien Kennzeichnung.|1,8|
|[worksheet](/javascript/api/excel/excel.worksheet)|_Eigenschaft_ > ShowHeadings|Ruft ab oder legt diesen fest im Arbeitsblatt Überschriften Kennzeichnung.|1,8|
|[worksheetCalculatedEventArgs](/javascript/api/excel/excel.worksheetcalculatedeventargs)|_Eigenschaft_ > Typ|Ruft den Typ des Ereignisses ab.|1,8|
|[worksheetCalculatedEventArgs](/javascript/api/excel/excel.worksheetcalculatedeventargs)|_Eigenschaft_ > WorksheetId|Ruft die Id des Arbeitsblatts, das berechnet wird.|1,8|

## <a name="whats-new-in-excel-javascript-api-17"></a>Was ist neu in Excel JavaScript-API 1.7

Die JavaScript-API für Excel-Anforderung Set 1.7 Features enthalten APIs für Diagramme, Ereignisse, Arbeitsblättern, Bereiche, Dokumenteigenschaften, benannte Elemente, die Schutzoptionen für und Formatvorlagen.

### <a name="customize-charts"></a>Anpassen von Diagrammen

Mit dem neuen Diagramm APIs können Sie zusätzliche Diagrammtypen erstellen, ein Diagramm eine Datenreihe hinzugefügt, des Diagrammtitels festgelegt, einen Achsentitel hinzufügen, hinzufügen Anzeigeeinheit, -Durchschnitt Trendlinie hinzufügen, ändern eine lineare Trendlinie und vieles mehr. Es folgen einige Beispiele:

* Diagramm der Achse - abzurufen, festzulegen, formatieren und entfernen Achse Einheit, Label und Title in einem Diagramm.
* Diagrammdatenreihen - hinzufügen, festlegen und löschen eine Datenreihe in einem Diagramm.  Ändern Sie Serie Markierungen, Zeichnungsfläche Orders und Größe.
* Diagramm-Trendlinien - hinzufügen, abrufen und Formatieren von Trendlinien in einem Diagramm.
* Diagrammlegende - Format Legendenschriftart in einem Diagramm.
* Diagramm-Punkt - Farbe des Diagramms festlegen.
* Diagramm der Titel Teilzeichenfolge - abrufen und Festlegen der Titel Teilzeichenfolge für ein Diagramm.
* Diagrammtyp - Option, um weitere Diagrammtypen erstellen.

### <a name="events"></a>Events

Tritt auf Excel-Ereignisse, die APIs eine Vielzahl von Ereignishandler bereitstellen, mit die Ihr Add-in automatisch eine festgelegte Funktion, wenn ein bestimmtes Ereignis ausführen können. Sie k?nnen diese Funktion so anpassen, dass die f?r Ihr Szenario erforderlichen Aktionen ausgef?hrt werden. Eine Liste der Ereignisse, die derzeit verfügbar sind, finden Sie unter [Arbeiten mit Ereignissen mithilfe der JavaScript-API für Excel](https://docs.microsoft.com/office/dev/add-ins/excel/excel-add-ins-events).

### <a name="customize-the-appearance-of-worksheets-and-ranges"></a>Anpassen der Darstellung von Arbeitsblättern und Bereichen

Die neuen APIs können Sie die Darstellung von Arbeitsblättern in verschiedene Arten anpassen:

* Fixieren von Fensterbereichen beibehalten bestimmte Zeilen oder Spalten angezeigt, wenn Sie im Arbeitsblatt einen Bildlauf durchführen. Beispielsweise, wenn die erste Zeile im Arbeitsblatt Kopfzeilen enthält, können Sie diese Zeile einfrieren, sodass die Spaltenüberschriften bleiben sichtbar, wie Sie einen Bildlauf nach unten im Arbeitsblatt durchführen.
* Ändern Sie die Arbeitsblatt Registerkarte Farbe.
* Arbeitsblatt Überschriften hinzufügen.


Sie können die Darstellung von Bereichen auf verschiedene Weise anpassen:

* Festlegen des Formats Zelle für einen Bereich, um sicherzustellen, dass Sie sicher, dass alle Zellen im Bereich konsistente Formatierung haben. Ein Zellenformat ist eine definierte Reihe von Formatierung Merkmale wie Schriftarten und Schriftgrade, Zahlenformate, Zellenrahmen und Zelle Schattierung. Verwenden Sie eine der integrierten Zellenformatvorlagen Excel, oder erstellen Sie eine eigene benutzerdefinierte Zellenformatvorlage.
* Legen Sie die Ausrichtung von Text für einen Bereich.
* Hinzufügen oder Ändern eines Hyperlinks auf einen Bereich, der an eine andere Position in der Arbeitsmappe oder auf einem externen Standort verknüpft.

### <a name="manage-document-properties"></a>Verwalten von Dokumenteigenschaften

Verwenden die Dokumenteigenschaften APIs, können Sie Zugriff auf die integrierten Dokumenteigenschaften und auch erstellen und Verwalten benutzerdefinierter Dokumenteigenschaften, um den Zustand der Arbeitsmappe und Laufwerk Workflow und Geschäftslogik zu speichern.

### <a name="copy-worksheets"></a>Kopieren von Arbeitsblättern

Die Kopie Arbeitsblatt APIs können Sie kopieren die Daten und das Format aus einem Arbeitsblatt in einem neuen Arbeitsblatt in derselben Arbeitsmappe und Reduzierung des Datenübertragung erforderlich.

### <a name="handle-ranges-with-ease"></a>Verarbeiten von Bereichen ganz einfach

Den verschiedenen Bereich APIs können Sie Aktionen wie Get den umgebenden Bereich, einen geänderten Bereich und vieles mehr erhalten möchten. Diese APIs sollten Aufgaben wie das Bereich Bearbeitung und Adressierung sehr viel effizienter.

Weitere Schritte:

* Arbeitsmappen und Arbeitsblätter Schutzoptionen - verwenden Sie diese APIs zum Schützen von Daten in einem Arbeitsblatt und die Struktur der Arbeitsmappe.
* Aktualisieren Sie ein benanntes Element – verwenden Sie diese API, um ein benanntes Element zu aktualisieren.
* Abrufen der aktiven Zelle – verwenden Sie diese API zum Abrufen der aktiven Zelle in einer Arbeitsmappe.

|Objekt| Neuigkeiten| Beschreibung|Anforderungssatz|
|:----|:----|:----|:----|
|[chart](/javascript/api/excel/excel.chart)|_Eigenschaft_ > ChartType|Stellt den Typ des Diagramms dar. Mögliche Werte sind: ColumnClustered, Gestapelte, ColumnStacked100, BarClustered, BarStacked, BarStacked100, LineStacked, LineStacked100, LineMarkers, LineMarkersStacked, LineMarkersStacked100, PieOfPie, usw....|1.7|
|[chart](/javascript/api/excel/excel.chart)|_Eigenschaft_ > id|Die eindeutige Id des Diagramms. Schreibgeschützt.|1.7|
|[chart](/javascript/api/excel/excel.chart)|_Eigenschaft_ > ShowAllFieldButtons|Stellt dar, ob in einer PivotChart alle Feldschaltflächen angezeigt.|1.7|
|[chartAreaFormat](/javascript/api/excel/excel.chartareaformat)|_Beziehung_ > Rahmen|Stellt das Rahmenformat der Diagrammfläche die Farbe, Linestyle und Weight enthält. Schreibgeschützt.|1.7|
|[chartAxes](/javascript/api/excel/excel.chartaxes)|_Methode_ > GetItem (Typ: string, gruppieren: Zeichenfolge)|Gibt die bestimmte Achse vom Typ und Gruppe zurück.|1.7|
|[chartAxis](/javascript/api/excel/excel.chartaxis)|_Eigenschaft_ > AxisBetweenCategories|Stellt dar, ob die Größenachse die Rubrikenachse zwischen den Rubriken schneidet.|1.7|
|[chartAxis](/javascript/api/excel/excel.chartaxis)|_Eigenschaft_ > AxisGroup|Stellt die Gruppe für die angegebene Achse dar. Schreibgeschützt. Mögliche Werte sind: primären, sekundären.|1.7|
|[chartAxis](/javascript/api/excel/excel.chartaxis)|_Eigenschaft_ > CategoryType|Gibt den Typ der Rubrikenachse zurück oder legt ihn fest. Mögliche Werte sind: automatisch, TextAxis, DateAxis.|1.7|
|[chartAxis](/javascript/api/excel/excel.chartaxis)|_Eigenschaft_ > schneidet|Die angegebene Achse den Schnittpunkt die anderen Achse darstellt. Mögliche Werte sind: automatische, das Maximum, Mindest- und benutzerdefinierten.|1.7|
|[chartAxis](/javascript/api/excel/excel.chartaxis)|_Eigenschaft_ > CrossesAt|Stellt die angegebene Achse Schnittpunkt mit die anderen Achse an. Schreibgeschützt. Auf diese Eigenschaft festgelegt ist, sollten SetCrossesAt(double)-Methode verwenden. Schreibgeschützt.|1.7|
|[chartAxis](/javascript/api/excel/excel.chartaxis)|_Eigenschaft_ > CustomDisplayUnit|Der Einheitenwert für benutzerdefinierte Achse Anzeige darstellt. Schreibgeschützt. Um diese Eigenschaft festzulegen, verwenden Sie die SetCustomDisplayUnit(double)-Methode. Schreibgeschützt.|1.7|
|[chartAxis](/javascript/api/excel/excel.chartaxis)|_Eigenschaft_ > DisplayUnit|Die Anzeigeeinheit der Achse darstellt. Mögliche Werte sind: None, Hunderte, Tausende, TenThousands, HundredThousands, Millionen, TenMillions, HundredMillions, Milliarden, Billionen, Benutzerdefiniert.|1.7|
|[chartAxis](/javascript/api/excel/excel.chartaxis)|_Eigenschaft_ > height|Stellt die Höhe in Punkt, der die Diagrammachse dar. NULL, wenn die Achse nicht sichtbar's. Schreibgeschützt.|1.7|
|[chartAxis](/javascript/api/excel/excel.chartaxis)|_Eigenschaft_ > linken|Stellt den Abstand in Punkt vom linken Rand der Achse auf der linken Seite des Diagrammbereichs. NULL, wenn die Achse nicht sichtbar's. Schreibgeschützt.|1.7|
|[chartAxis](/javascript/api/excel/excel.chartaxis)|_Eigenschaft_ > LogBase|Stellt die Basis des Logarithmus wenn logarithmische Skalen verwenden.|1.7|
|[chartAxis](/javascript/api/excel/excel.chartaxis)|_Eigenschaft_ > ReversePlotOrder|Stellt dar, ob Microsoft Excel Datenpunkte vom letzten zum ersten darstellt.|1.7|
|[chartAxis](/javascript/api/excel/excel.chartaxis)|_Eigenschaft_ > ScaleType|Stellt den Typ der Achse Skalierung Wert dar. Mögliche Werte sind: Linear, logarithmische.|1.7|
|[chartAxis](/javascript/api/excel/excel.chartaxis)|_Eigenschaft_ > ShowDisplayUnitLabel|Stellt dar, ob die Achse anzeigeeinheitenbeschriftung sichtbar ist.|1.7|
|[chartAxis](/javascript/api/excel/excel.chartaxis)|_Eigenschaft_ > TickLabelSpacing|Stellt die Anzahl der Rubriken oder Datenreihen zwischen teilstrichbeschriftungen dar. Ein Wert von 1 bis 31999 oder eine leere Zeichenfolge für die Einstellung für automatische kann sein. Der zurückgegebene Wert ist immer eine Zahl.|1.7|
|[chartAxis](/javascript/api/excel/excel.chartaxis)|_Eigenschaft_ > TickMarkSpacing|Stellt die Anzahl der Rubriken oder Datenreihen zwischen Teilstrichen.|1.7|
|[chartAxis](/javascript/api/excel/excel.chartaxis)|_Eigenschaft_ > oben|Stellt den Abstand in Punkt vom oberen Rand der Achse an den Anfang der Diagrammfläche. NULL, wenn die Achse nicht sichtbar's. Schreibgeschützt.|1.7|
|[chartAxis](/javascript/api/excel/excel.chartaxis)|_Eigenschaft_ > Typ|Stellt den Typ der Achse an. Schreibgeschützt. Mögliche Werte sind: Ungültige, Kategorie, Wert, Datenreihe.|1.7|
|[chartAxis](/javascript/api/excel/excel.chartaxis)|_Eigenschaft_ > sichtbar|Ein boolescher Wert stellt die Sichtbarkeit der Achse.|1.7|
|[chartAxis](/javascript/api/excel/excel.chartaxis)|_Eigenschaft_ > width|Stellt die Breite in Punkt, der die Diagrammachse dar. NULL, wenn die Achse nicht sichtbar's. Schreibgeschützt.|1.7|
|[chartAxis](/javascript/api/excel/excel.chartaxis)|_Beziehung_ > BaseTimeUnit|Gibt die Basiseinheit für die angegebene Rubrikenachse zurück oder legt diese Einheit fest.|1.7|
|[chartAxis](/javascript/api/excel/excel.chartaxis)|_Beziehung_ > MajorTickMark|Stellt den Typ der wichtigsten Teilstrich der angegebenen Achse an.|1.7|
|[chartAxis](/javascript/api/excel/excel.chartaxis)|_Beziehung_ > MajorTimeUnitScale|Zurückgeben oder Festlegen der wichtigsten Einheitenwert des Skalierung für die Rubrikenachse, wenn die CategoryType-Eigenschaft auf Zeitskala festgelegt ist.|1.7|
|[chartAxis](/javascript/api/excel/excel.chartaxis)|_Beziehung_ > MinorTickMark|Stellt den Typ des Hilfsteilstriche für die angegebene Achse markieren.|1.7|
|[chartAxis](/javascript/api/excel/excel.chartaxis)|_Beziehung_ > MinorTimeUnitScale|Gibt an, oder der kleinere Einheit Scale-Wert für die Rubrikenachse festgelegt, wenn die CategoryType-Eigenschaft auf Zeitskala festgelegt ist.|1.7|
|[chartAxis](/javascript/api/excel/excel.chartaxis)|_Beziehung_ > TickLabelPosition|Stellt die Position der teilstrichbeschriftungen auf der angegebenen Achse an.|1.7|
|[chartAxis](/javascript/api/excel/excel.chartaxis)|_Methode_ > SetCategoryNames(sourceData: Range)|Alle Rubrikennamen für die angegebene Achse festgelegt.|1.7|
|[chartAxis](/javascript/api/excel/excel.chartaxis)|_Methode_ > SetCrossesAt(value: double)|Festlegen der angegebenen Achse Schnittpunkt mit die anderen Achse an.|1.7|
|[chartAxis](/javascript/api/excel/excel.chartaxis)|_Methode_ > SetCustomDisplayUnit(value: double)|Die Achse Anzeigeeinheit festgelegt auf einen benutzerdefinierten Wert.|1.7|
|[chartBorder](/javascript/api/excel/excel.chartborder)|_Eigenschaft_ > color|HTML-Farbcode, das die Farbe des Rahmens im Diagramm darstellt.|1.7|
|[chartBorder](/javascript/api/excel/excel.chartborder)|_Eigenschaft_ > Gewicht|Stellt die Stärke des Rahmens, in Punkt.|1.7|
|[chartBorder](/javascript/api/excel/excel.chartborder)|_Beziehung_ > LineStyle|Die Linienart des Rahmens darstellt.|1.7|
|[chartDataLabel](/javascript/api/excel/excel.chartdatalabel)|_Eigenschaft_ > Position|DataLabelPosition-Wert, der die Position der Datenbeschriftung darstellt. Die folgenden Werte sind möglich: None, Center, InsideEnd, InsideBase, OutsideEnd, Left, Right, Top, Bottom, BestFit, Callout.|1.7|
|[chartDataLabel](/javascript/api/excel/excel.chartdatalabel)|_Eigenschaft_ > Trennzeichen|Eine Zeichenfolge, die das Trennzeichen für die Datenbeschriftung in einem Diagramm darstellt.|1.7|
|[chartDataLabel](/javascript/api/excel/excel.chartdatalabel)|_Eigenschaft_ > ShowBubbleSize|Boolescher Wert, der angibt, ob die Größe der Datenbeschriftungs-Sprechblase angezeigt wird.|1.7|
|[chartDataLabel](/javascript/api/excel/excel.chartdatalabel)|_Eigenschaft_ > ShowCategoryName|Boolescher Wert, der angibt, ob der Kategoriename der Datenbeschriftung angezeigt wird.|1.7|
|[chartDataLabel](/javascript/api/excel/excel.chartdatalabel)|_Eigenschaft_ > ShowLegendKey|Boolescher Wert, der angibt, ob das Legendensymbol der Datenbeschriftung angezeigt wird.|1.7|
|[chartDataLabel](/javascript/api/excel/excel.chartdatalabel)|_Eigenschaft_ > ShowPercentage|Boolescher Wert, der angibt, ob der Prozentsatz der Datenbeschriftung angezeigt wird.|1.7|
|[chartDataLabel](/javascript/api/excel/excel.chartdatalabel)|_Eigenschaft_ > ShowSeriesName|Boolescher Wert, der angibt, ob der Name der Datenbeschriftungsreihe angezeigt wird.|1.7|
|[chartDataLabel](/javascript/api/excel/excel.chartdatalabel)|_Eigenschaft_ > ShowValue|Boolescher Wert, der angibt, ob der Datenbeschriftungswert angezeigt wird.|1.7|
|[chartLegend](/javascript/api/excel/excel.chartlegend)|_Eigenschaft_ > height|Die Höhe der Legende im Diagramm darstellt.|1.7|
|[chartLegend](/javascript/api/excel/excel.chartlegend)|_Eigenschaft_ > linken|Stellt Links neben einer Diagrammlegende dar.|1.7|
|[chartLegend](/javascript/api/excel/excel.chartlegend)|_Eigenschaft_ > ShowShadow|Stellt dar, ob die Legende im Diagramm Schatten aufweist.|1.7|
|[chartLegend](/javascript/api/excel/excel.chartlegend)|_Eigenschaft_ > oben|Stellt im oberen Bereich des einer Diagrammlegende dar.|1.7|
|[chartLegend](/javascript/api/excel/excel.chartlegend)|_Eigenschaft_ > width|Die Breite der Legende im Diagramm darstellt.|1.7|
|[chartLegend](/javascript/api/excel/excel.chartlegend)|_Beziehung_ > LegendEntries|Stellt eine Auflistung von LegendEntries in der Legende dar. Schreibgeschützt.|1.7|
|[chartLegendEntry](/javascript/api/excel/excel.chartlegendentry)|_Eigenschaft_ > sichtbar|Stellt die sichtbaren eines Diagramms Legendeneintrags.|1.7|
|[chartLegendEntryCollection](/javascript/api/excel/excel.chartlegendentrycollection)|_Eigenschaft_ > items|Eine Auflistung von ChartLegendEntry-Objekten. Schreibgeschützt.|1.7|
|[chartLegendEntryCollection](/javascript/api/excel/excel.chartlegendentrycollection)|_Methode_ > getCount()|Gibt die Anzahl der LegendEntry in der Auflistung zurück.|1.7|
|[chartLegendEntryCollection](/javascript/api/excel/excel.chartlegendentrycollection)|_Methode_ > GetItemAt(index: number)|Gibt ein LegendEntry am angegebenen Index zurück.|1.7|
|[chartPoint](/javascript/api/excel/excel.chartpoint)|_Eigenschaft_ > HasDataLabel|Stellt dar, ob ein Datenpunkt Datalabel verfügt. Nicht zutreffend für Fläche Diagramme.|1.7|
|[chartPoint](/javascript/api/excel/excel.chartpoint)|_Eigenschaft_ > MarkerBackgroundColor|Zeigen Sie HTML-Farbe Code Darstellung der Hintergrundfarbe der Markierung von Daten. E.g. # FF0000 für Rot.|1.7|
|[chartPoint](/javascript/api/excel/excel.chartpoint)|_Eigenschaft_ > MarkerForegroundColor|Zeigen Sie HTML-Farbe Code Darstellung der Markierung Vordergrundfarbe der Daten. E.g. # FF0000 für Rot.|1.7|
|[chartPoint](/javascript/api/excel/excel.chartpoint)|_Eigenschaft_ > MarkerSize|Stellt die Größe des Datenpunkts.|1.7|
|[chartPoint](/javascript/api/excel/excel.chartpoint)|_Eigenschaft_ > MarkerStyle|Stellt die Art der Markierung eines Diagrammdatenpunkts dar. Mögliche Werte sind: Ungültige, automatisch, keine, Quadrat, Raute, Dreieck, X, Stern, Dot, Strich, Circle sowie, Bild.|1.7|
|[chartPoint](/javascript/api/excel/excel.chartpoint)|_Beziehung_ > DataLabel|Gibt die Beschriftung eines Diagramms Punkts zurück. Schreibgeschützt.|1.7|
|[chartPointFormat](/javascript/api/excel/excel.chartpointformat)|_Beziehung_ > Rahmen|Stellt das Rahmenformat eines Diagrammdatenpunkts, die Farbe, Stil und Weight Informationen enthält. Schreibgeschützt.|1.7|
|[chartSeries](/javascript/api/excel/excel.chartseries)|_Eigenschaft_ > ChartType|Stellt den Diagrammtyp einer Reihe an. Mögliche Werte sind: ColumnClustered, Gestapelte, ColumnStacked100, BarClustered, BarStacked, BarStacked100, LineStacked, LineStacked100, LineMarkers, LineMarkersStacked, LineMarkersStacked100, PieOfPie, usw....|1.7|
|[chartSeries](/javascript/api/excel/excel.chartseries)|_Eigenschaft_ > DoughnutHoleSize|Stellt die Innenringgröße von einer Diagrammdatenreihe dar.  Gilt nur für Diagramme Ring und DoughnutExploded.|1.7|
|[chartSeries](/javascript/api/excel/excel.chartseries)|_Eigenschaft_ > gefiltert|Boolescher Wert, wenn die Datenreihe oder nicht gefiltert wird. Nicht zutreffend für Fläche Diagramme.|1.7|
|[chartSeries](/javascript/api/excel/excel.chartseries)|_Eigenschaft_ > GapWidth|Stellt den Abstand von einer Diagrammdatenreihe dar.  Nur gültig für häufig verwendete Hyperlinks und Spalte Diagramme, sowie|1.7|
|[chartSeries](/javascript/api/excel/excel.chartseries)|_Eigenschaft_ > HasDataLabels|Boolescher Wert, wenn die Datenreihe datenbeschriftungen oder nicht aufweist.|1.7|
|[chartSeries](/javascript/api/excel/excel.chartseries)|_Eigenschaft_ > MarkerBackgroundColor|Stellt Markierungen Hintergrundfarbe einer Diagrammdatenreihe dar.|1.7|
|[chartSeries](/javascript/api/excel/excel.chartseries)|_Eigenschaft_ > MarkerForegroundColor|Markierungen Vordergrundfarbe einer diagrammserie darstellt.|1.7|
|[chartSeries](/javascript/api/excel/excel.chartseries)|_Eigenschaft_ > MarkerSize|Stellt die Größe von einer Diagrammdatenreihe dar.|1.7|
|[chartSeries](/javascript/api/excel/excel.chartseries)|_Eigenschaft_ > MarkerStyle|Stellt die Art der Markierung von einer Diagrammdatenreihe dar. Mögliche Werte sind: Ungültige, automatisch, keine, Quadrat, Raute, Dreieck, X, Stern, Dot, Strich, Circle sowie, Bild.|1.7|
|[chartSeries](/javascript/api/excel/excel.chartseries)|_Eigenschaft_ > PlotOrder|Die Reihenfolge der Zeichnungsflächen einer diagrammserie in der Diagrammgruppe darstellt.|1.7|
|[chartSeries](/javascript/api/excel/excel.chartseries)|_Eigenschaft_ > ShowShadow|Boolescher Wert, wenn die Datenreihe Schatten oder nicht aufweist.|1.7|
|[chartSeries](/javascript/api/excel/excel.chartseries)|_Eigenschaft_ > reibungslose|Boolescher Wert, wenn die Datenreihe reibungslose oder nicht ist. Nur für Linien- und Punkt (XY)-Diagramme.|1.7|
|[chartSeries](/javascript/api/excel/excel.chartseries)|_Beziehung_ > DataLabels|Eine Auflistung aller DataLabel in der Datenreihe darstellt. Schreibgeschützt.|ApiSet.InProgressFeatures.ChartingAPI|
|[chartSeries](/javascript/api/excel/excel.chartseries)|_Beziehung_ > Trendlinien|Stellt eine Auflistung von Trendlinien der Datenreihe. Schreibgeschützt.|1.7|
|[chartSeries](/javascript/api/excel/excel.chartseries)|_Methode_ > delete()|Löscht die Diagrammdatenreihe.|1.7|
|[chartSeries](/javascript/api/excel/excel.chartseries)|_Methode_ > SetBubbleSizes(sourceData: Range)|Legen Sie die Größe der Blasen für einer Diagrammdatenreihe dar. Funktioniert nur für Blasendiagramme.|1.7|
|[chartSeries](/javascript/api/excel/excel.chartseries)|_Methode_ > SetValues(sourceData: Range)|Legen Sie Werte für einer Diagrammdatenreihe dar. Für Punkt (XY)-Diagramm bedeutet dies Y-Achsen-Werten.|1.7|
|[chartSeries](/javascript/api/excel/excel.chartseries)|_Methode_ > SetXAxisValues(sourceData: Range)|Legen Sie die Werte der X-Achse für einer Diagrammdatenreihe dar. Funktioniert nur für Punkt (XY)-Diagramme.|1.7|
|[chartSeriesCollection](/javascript/api/excel/excel.chartseriescollection)|_Methode_ > hinzufügen (Name: string, index: Anzahl)|Der Auflistung eine neue Datenreihe hinzugefügt.|1.7|
|[chartTitle](/javascript/api/excel/excel.charttitle)|_Eigenschaft_ > height|Gibt die Höhe des des Diagrammtitels in Punkt zurück. Schreibgeschützt. "NULL" Wenn Diagrammtitel nicht sichtbar's. Schreibgeschützt.|1.7|
|[chartTitle](/javascript/api/excel/excel.charttitle)|_Eigenschaft_ > HorizontalAlignment|Stellt die horizontale Ausrichtung für den Diagrammtitel dar. Mögliche Werte sind: nach rechts Mitte, links, Justify, verteilt.|1.7|
|[chartTitle](/javascript/api/excel/excel.charttitle)|_Eigenschaft_ > linken|Stellt den Abstand in Punkt vom linken Rand des Diagrammtitels zum linken Rand des Diagrammbereichs dar. "NULL" Wenn Diagrammtitel nicht sichtbar's.|1.7|
|[chartTitle](/javascript/api/excel/excel.charttitle)|_Eigenschaft_ > Position|Die Position des Diagrammtitels darstellt. Mögliche Werte sind: oben, automatisch, unten, rechts, links.|1.7|
|[chartTitle](/javascript/api/excel/excel.charttitle)|_Eigenschaft_ > ShowShadow|Stellt einen booleschen Wert, der bestimmt, ob der Diagrammtitel einen Schatten aufweist.|1.7|
|[chartTitle](/javascript/api/excel/excel.charttitle)|_Eigenschaft_ > TextOrientation|Die Ausrichtung von Text des Diagrammtitels darstellt. Der Wert muss eine ganze Zahl sein entweder von-90 bis 90 oder 180 für Text vertikal ausgerichtet.|1.7|
|[chartTitle](/javascript/api/excel/excel.charttitle)|_Eigenschaft_ > oben|Stellt den Abstand in Punkt vom oberen Rand des Diagrammtitels am Anfang der Diagrammfläche. "NULL" Wenn Diagrammtitel nicht sichtbar's.|1.7|
|[chartTitle](/javascript/api/excel/excel.charttitle)|_Eigenschaft_ > VerticalAlignment|Die vertikale Ausrichtung des Diagrammtitels darstellt. Mögliche Werte sind: Center, unten, oben, Justify, verteilt.|1.7|
|[chartTitle](/javascript/api/excel/excel.charttitle)|_Eigenschaft_ > width|Gibt die Breite in Punkt, der den Diagrammtitel zurück. Schreibgeschützt. "NULL" Wenn Diagrammtitel nicht sichtbar's. Schreibgeschützt.|1.7|
|[chartTitle](/javascript/api/excel/excel.charttitle)|_Methode_ > SetFormula(formula: string)|Legt einen String-Wert, der die Formel des Diagrammtitels unter Verwendung der A1-Schreibweise darstellt.|1.7|
|[chartTitleFormat](/javascript/api/excel/excel.charttitleformat)|_Beziehung_ > Rahmen|Stellt das Rahmenformat des Diagrammtitels, einschließlich Color, Linestyle und Weight. Schreibgeschützt.|1.7|
|[chartTrendline](/javascript/api/excel/excel.charttrendline)|_Eigenschaft_ > rückwärts|Stellt die Anzahl der Zeiträume, die die sich eine Trendlinie zurück erstreckt.|1.7|
|[chartTrendline](/javascript/api/excel/excel.charttrendline)|_Eigenschaft_ > DisplayEquation|True, wenn die Formel für die sich eine Trendlinie im Diagramm angezeigt wird.|1.7|
|[chartTrendline](/javascript/api/excel/excel.charttrendline)|_Eigenschaft_ > DisplayRSquared|True, wenn der R-quadrierten für die sich eine Trendlinie im Diagramm angezeigt wird.|1.7|
|[chartTrendline](/javascript/api/excel/excel.charttrendline)|_Eigenschaft_ > weiterleiten|Stellt die Anzahl der Zeiträume, die die sich eine Trendlinie vorwärts erstreckt.|1.7|
|[chartTrendline](/javascript/api/excel/excel.charttrendline)|_Eigenschaft_ > Intercept|Den Intercept-Wert der Trendlinie darstellt. Kann auf einen numerischen Wert oder eine leere Zeichenfolge (für automatische Werte) festgelegt werden. Der zurückgegebene Wert ist immer eine Zahl.|1.7|
|[chartTrendline](/javascript/api/excel/excel.charttrendline)|_Eigenschaft_ > MovingAveragePeriod|Stellt den Zeitraum einer Trendlinie Diagramm nur für Trendlinie mit MovingAverage Typ dar.|1.7|
|[chartTrendline](/javascript/api/excel/excel.charttrendline)|_Eigenschaft_ > name|Stellt den Namen der Trendlinie dar. Kann auf einen String-Wert festgelegt werden, oder auf null-Wert stellt automatische Werte festgelegt werden kann. Der zurückgegebene Wert ist immer eine Zeichenfolge|1.7|
|[chartTrendline](/javascript/api/excel/excel.charttrendline)|_Eigenschaft_ > PolynomialOrder|Die Reihenfolge der eine Trendlinie Diagramm nur für Trendlinie mit Polynomial Typ darstellt.|1.7|
|[chartTrendline](/javascript/api/excel/excel.charttrendline)|_Eigenschaft_ > Typ|Stellt den Typ einer Trendlinie Diagramm dar. Mögliche Werte sind: Linear, logarithmische, exponentielle MovingAverage, Polynomial, Power.|1.7|
|[chartTrendline](/javascript/api/excel/excel.charttrendline)|_Beziehung_ > Format|Stellt die Formatierung einer Trendlinie Diagramm dar. Schreibgeschützt.|1.7|
|[chartTrendline](/javascript/api/excel/excel.charttrendline)|_Methode_ > delete()|Das Trendline-Objekt zu löschen.|1.7|
|[chartTrendlineCollection](/javascript/api/excel/excel.charttrendlinecollection)|_Eigenschaft_ > items|Eine Auflistung von ChartTrendline-Objekten. Schreibgeschützt.|1.7|
|[chartTrendlineCollection](/javascript/api/excel/excel.charttrendlinecollection)|_Methode_ > Add(type: string)|Trendline-Auflistung hinzugefügt eine neue Trendlinie.|1.7|
|[chartTrendlineCollection](/javascript/api/excel/excel.charttrendlinecollection)|_Methode_ > getCount()|Gibt die Anzahl der Trendlinien in der Auflistung zurück.|1.7|
|[chartTrendlineCollection](/javascript/api/excel/excel.charttrendlinecollection)|_Methode_ > GetItem(index: number)|Rufen Sie Trendline-Objekt ab, indem Index, der die Reihenfolge der Einfügemarke im Items-Array ist.|1.7|
|[chartTrendlineFormat](/javascript/api/excel/excel.charttrendlineformat)|_Beziehung_ > Linie|Stellt die Formatierung der Diagrammlinien dar. Schreibgeschützt.|1.7|
|[customProperty](/javascript/api/excel/excel.customproperty)|_Eigenschaft_ > key|Ruft den Schlüssel der benutzerdefinierten Eigenschaft ab. Schreibgeschützt. Schreibgeschützt.|1.7|
|[customProperty](/javascript/api/excel/excel.customproperty)|_Eigenschaft_ > Typ|Ruft den Werttyp der benutzerdefinierten Eigenschaft. Schreibgeschützt. Schreibgeschützt. Mögliche Werte sind: Anzahl, Boolean, Date, String, Float.|1.7|
|[customProperty](/javascript/api/excel/excel.customproperty)|_Eigenschaft_ > value|Ruft den Wert der benutzerdefinierten Eigenschaft ab oder legt ihn fest.|1.7|
|[customProperty](/javascript/api/excel/excel.customproperty)|_Methode_ > delete()|Löscht die benutzerdefinierte Eigenschaft.|1.7|
|[customPropertyCollection](/javascript/api/excel/excel.custompropertycollection)|_Eigenschaft_ > items|Eine Sammlung von customProperty-Objekten. Schreibgeschützt.|1.7|
|[customPropertyCollection](/javascript/api/excel/excel.custompropertycollection)|_Methode_ > hinzufügen (Schlüssel: string-Wert: Objekt)|Erstellt eine neue benutzerdefinierte Eigenschaft oder legt eine vorhandene fest.|1.7|
|[customPropertyCollection](/javascript/api/excel/excel.custompropertycollection)|_Methode_ > deleteAll()|Löscht alle benutzerdefinierten Eigenschaften in dieser Sammlung.|1.7|
|[customPropertyCollection](/javascript/api/excel/excel.custompropertycollection)|_Methode_ > getCount()|Ruft die Anzahl von benutzerdefinierten Eigenschaften ab.|1.7|
|[customPropertyCollection](/javascript/api/excel/excel.custompropertycollection)|_Methode_ > GetItem(key: string)|Ruft ein Objekt für benutzerdefinierte Eigenschaften über seinen Schlüssel ab, bei dem Groß-/Kleinschreibung nicht beachtet wird. Wird ausgelöst, wenn die benutzerdefinierte Eigenschaft nicht vorhanden ist.|1.7|
|[customPropertyCollection](/javascript/api/excel/excel.custompropertycollection)|_Methode_ > GetItemOrNullObject(key: string)|Ruft ein Objekt für benutzerdefinierte Eigenschaften über seinen Schlüssel ab, bei dem Groß-/Kleinschreibung nicht beachtet wird. Gibt ein NULL-Objekt zurück, wenn die benutzerdefinierte Eigenschaft nicht vorhanden ist.|1.7|
|[dataConnectionCollection](/javascript/api/excel/excel.dataconnectioncollection)|_Eigenschaft_ > items|Eine Auflistung von DataConnection-Objekten. Schreibgeschützt.|1.7|
|[dataConnectionCollection](/javascript/api/excel/excel.dataconnectioncollection)|_Methode_ > refreshAll()|Aktualisiert die Datenverbindungen in der Auflistung.|1.7|
|[documentProperties](/javascript/api/excel/excel.documentproperties)|_Eigenschaft_ > author|Dient zum Abrufen oder festlegen den Autor der Arbeitsmappe.|1.7|
|[documentProperties](/javascript/api/excel/excel.documentproperties)|_Eigenschaft_ > category|Ruft ab oder legt die Kategorie der Arbeitsmappe fest.|1.7|
|[documentProperties](/javascript/api/excel/excel.documentproperties)|_Eigenschaft_ > comments|Dient zum Abrufen oder die Kommentare der Arbeitsmappe festgelegt.|1.7|
|[documentProperties](/javascript/api/excel/excel.documentproperties)|_Eigenschaft_ > company|Dient zum Abrufen oder Festlegen des Unternehmens der Arbeitsmappe.|1.7|
|[documentProperties](/javascript/api/excel/excel.documentproperties)|_Eigenschaft_ > keywords|Dient zum Abrufen oder die Schlüsselwörter der Arbeitsmappe festgelegt.|1.7|
|[documentProperties](/javascript/api/excel/excel.documentproperties)|_Eigenschaft_ > lastAuthor|Ruft den letzten Autor der Arbeitsmappe ab. Schreibgeschützt. Schreibgeschützt.|1.7|
|[documentProperties](/javascript/api/excel/excel.documentproperties)|_Eigenschaft_ > manager|Dient zum Abrufen oder Festlegen der Vorgesetzte des der Arbeitsmappe.|1.7|
|[documentProperties](/javascript/api/excel/excel.documentproperties)|_Eigenschaft_ > revisionNumber|Ruft die Revisionsnummer der Arbeitsmappe ab. Schreibgeschützt.|1.7|
|[documentProperties](/javascript/api/excel/excel.documentproperties)|_Eigenschaft_ > subject|Dient zum Abrufen oder Festlegen des Betreffs der Arbeitsmappe.|1.7|
|[documentProperties](/javascript/api/excel/excel.documentproperties)|_Eigenschaft_ > title|Dient zum Abrufen oder Festlegen des Titels der Arbeitsmappe.|1.7|
|[documentProperties](/javascript/api/excel/excel.documentproperties)|_Beziehung_ > creationDate|Ruft das Erstellungsdatum der Arbeitsmappe ab. Schreibgeschützt. Schreibgeschützt.|1.7|
|[documentProperties](/javascript/api/excel/excel.documentproperties)|_Beziehung_ > benutzerdefinierte|Ruft die Auflistung benutzerdefinierter Eigenschaften der Arbeitsmappe ab. Schreibgeschützt. Schreibgeschützt.|1.7|
|[namedItem](/javascript/api/excel/excel.nameditem)|_Eigenschaft_ > Formel|Ruft ab oder legt die Formel für das benannte Element.  Formel beginnt immer mit einem "=" anmelden.|1.7|
|[namedItem](/javascript/api/excel/excel.nameditem)|_Beziehung_ > ArrayValues|Gibt ein Objekt, das Werte und Typen des benannten Elements enthält. Schreibgeschützt.|1.7|
|[namedItemArrayValues](/javascript/api/excel/excel.nameditemarrayvalues)|_Eigenschaft_ > Typen|Stellt die Typen für jedes Element im Array benanntes Element schreibgeschützt. Mögliche Werte sind: unbekannt, leer, String, Integer, Double, Boolean, Fehler.|1.7|
|[namedItemArrayValues](/javascript/api/excel/excel.nameditemarrayvalues)|_Eigenschaft_ > values|Die Werte der einzelnen Elemente im Array benanntes Element darstellt. Schreibgeschützt.|1.7|
|[range](/javascript/api/excel/excel.range)|_Eigenschaft_ > IsEntireColumn|Stellt dar, wenn der aktuelle Bereich eine ganze Spalte ist. Schreibgeschützt.|1.7|
|[range](/javascript/api/excel/excel.range)|_Eigenschaft_ > IsEntireRow|Stellt dar, wenn der aktuelle Bereich eine ganze Zeile ist. Schreibgeschützt.|1.7|
|[range](/javascript/api/excel/excel.range)|_Eigenschaft_ > NumberFormatLocal|Excel Zahlenformat für den angegebenen Bereich als eine Zeichenfolge in der Sprache des Benutzers dargestellt wird.|1.7|
|[range](/javascript/api/excel/excel.range)|_Eigenschaft_ > style|Das Format des aktuellen Bereichs darstellt. Diese null oder eine Zeichenfolge zurück.|1.7|
|[range](/javascript/api/excel/excel.range)|_Methode_ > GetAbsoluteResizedRange (NumRows: number, NumColumns: Anzahl)|Ruft ein Range-Objekt mit der gleichen linke obere Zelle als das aktuelle Range-Objekt, wobei jedoch die angegebene Anzahl von Zeilen und Spalten ab.|1.7|
|[range](/javascript/api/excel/excel.range)|_Methode_ > getImage()|Gibt den Bereich als base64-codierten Bild wieder.|1.7|
|[range](/javascript/api/excel/excel.range)|_Methode_ > getSurroundingRegion()|Gibt ein Range-Objekt, das die umgebende Region für die linke obere Zelle in diesem Bereich darstellt. Ein umgebenden Bereich ist ein Bereich, der durch eine beliebige Kombination leere Zeilen und leere Spalten relativ zu diesem Bereich begrenzt.|1.7|
|[range](/javascript/api/excel/excel.range)|_Methode_ > showCard()|Zeigt die Visitenkarte für eine aktive Zelle rich Inhalt verfügt.|1.7|
|[rangeFormat](/javascript/api/excel/excel.rangeformat)|_Eigenschaft_ > TextOrientation|Ruft ab oder legt die Ausrichtung von Text aller Zellen innerhalb des Bereichs.|1.7|
|[rangeFormat](/javascript/api/excel/excel.rangeformat)|_Eigenschaft_ > UseStandardHeight|Bestimmt, ob die Zeilenhöhe des Range-Objekts die Standardhöhe für das Blatt übereinstimmt.|1.7|
|[rangeFormat](/javascript/api/excel/excel.rangeformat)|_Eigenschaft_ > UseStandardWidth|Bestimmt, ob der Columnwidth des Range-Objekts der Standardbreite des Blatts entspricht.|1.7|
|[rangeHyperlink](/javascript/api/excel/excel.rangehyperlink)|_Eigenschaft_ > address|Das Url-Ziel des Hyperlinks darstellt.|1.7|
|[rangeHyperlink](/javascript/api/excel/excel.rangehyperlink)|_Eigenschaft_ > Dokument.|Das Dokument darstellt. Ziel für den Hyperlink.|1.7|
|[rangeHyperlink](/javascript/api/excel/excel.rangehyperlink)|_Eigenschaft_ > QuickInfo|Stellt die Zeichenfolge, die angezeigt werden, wenn der Mauszeiger über dem Hyperlink dar.|1.7|
|[rangeHyperlink](/javascript/api/excel/excel.rangehyperlink)|_Eigenschaft_ > TextToDisplay|Stellt die Zeichenfolge, die in den meisten Zelle im Bereich von links oben angezeigt wird.|1.7|
|[Formatvorlage](/javascript/api/excel/excel.style)|_Eigenschaft_ > AddIndent|Gibt an, ob Text automatisch eingezogen wird, wenn die Ausrichtung des Texts in einer Zelle auf gleiche Verteilung festgelegt ist.|1.7|
|[Formatvorlage](/javascript/api/excel/excel.style)|_Eigenschaft_ > AutoIndent|Gibt an, ob Text automatisch eingezogen wird, wenn die Ausrichtung des Texts in einer Zelle auf gleiche Verteilung festgelegt ist.|1.7|
|[Formatvorlage](/javascript/api/excel/excel.style)|_Eigenschaft_ > BuiltIn|Gibt an, ob die Formatvorlage um eine integrierte Formatvorlage handelt. Schreibgeschützt.|1.7|
|[Formatvorlage](/javascript/api/excel/excel.style)|_Eigenschaft_ > formulaHidden|Gibt an, ob die Formel ausgeblendet ist, wenn das Arbeitsblatt geschützt ist.|1.7|
|[Formatvorlage](/javascript/api/excel/excel.style)|_Eigenschaft_ > HorizontalAlignment|Die horizontale Ausrichtung für die Formatvorlage repräsentiert. Mögliche Werte sind: Allgemein, links, Mitte, rechts, Füllung, Justify, CenterAcrossSelection, verteilt.|1.7|
|[Formatvorlage](/javascript/api/excel/excel.style)|_Eigenschaft_ > IncludeAlignment|Gibt an, ob die Formatvorlage die Eigenschaften AutoIndent, HorizontalAlignment, VerticalAlignment, WrapText, IndentLevel und TextOrientation enthält.|1.7|
|[Formatvorlage](/javascript/api/excel/excel.style)|_Eigenschaft_ > IncludeBorder|Gibt an, ob die Formatvorlage die Rahmeneigenschaften Color, ColorIndex, LineStyle und Weight enthält.|1.7|
|[Formatvorlage](/javascript/api/excel/excel.style)|_Eigenschaft_ > IncludeFont|Gibt an, ob die Formatvorlage die Schriftarteigenschaften Hintergrund, fett, Farbe, ColorIndex, FontStyle, kursiv, Name, Größe, durchgestrichen, tiefgestellt, hochgestellt und Unterstreichung enthält.|1.7|
|[Formatvorlage](/javascript/api/excel/excel.style)|_Eigenschaft_ > IncludeNumber|Gibt an, ob die Formatvorlage die NumberFormat-Eigenschaft enthält.|1.7|
|[Formatvorlage](/javascript/api/excel/excel.style)|_Eigenschaft_ > IncludePatterns|Gibt an, ob die Formatvorlage der Farbe, ColorIndex, InvertIfNegative, Pattern, PatternColor und PatternColorIndex einschließt.|1.7|
|[Formatvorlage](/javascript/api/excel/excel.style)|_Eigenschaft_ > IncludeProtection|Gibt an, ob die Formatvorlage die FormulaHidden und Locked Protection-Eigenschaften enthält.|1.7|
|[Formatvorlage](/javascript/api/excel/excel.style)|_Eigenschaft_ > IndentLevel|Eine ganze Zahl von 0 bis zu 250, die die Einzugsebene für die Formatvorlage angibt.|1.7|
|[Formatvorlage](/javascript/api/excel/excel.style)|_Eigenschaft_ > locked|Gibt an, ob das Objekt gesperrt ist, wenn das Arbeitsblatt geschützt ist.|1.7|
|[Formatvorlage](/javascript/api/excel/excel.style)|_Eigenschaft_ > name|Der Name der Formatvorlage. Schreibgeschützt.|1.7|
|[Formatvorlage](/javascript/api/excel/excel.style)|_Eigenschaft_ > numberFormat|Den Formatierungscode für das Zahlenformat für die Formatvorlage.|1.7|
|[Formatvorlage](/javascript/api/excel/excel.style)|_Eigenschaft_ > NumberFormatLocal|Die lokalisierten Formatierungscode für das Zahlenformat für die Formatvorlage.|1.7|
|[Formatvorlage](/javascript/api/excel/excel.style)|_Eigenschaft_ > Ausrichtung|Die Ausrichtung von Text für die Formatvorlage.|1.7|
|[Formatvorlage](/javascript/api/excel/excel.style)|_Eigenschaft_ > ReadingOrder|Die Lesereihenfolge für die Formatvorlage. Mögliche Werte sind: Kontext LeftToRight, RightToLeft.|1.7|
|[Formatvorlage](/javascript/api/excel/excel.style)|_Eigenschaft_ > ShrinkToFit|Gibt an, ob Text automatisch an die verfügbare Spaltenbreite angepasst wird.|1.7|
|[Formatvorlage](/javascript/api/excel/excel.style)|_Eigenschaft_ > TextOrientation|Die Ausrichtung von Text für die Formatvorlage.|1.7|
|[Formatvorlage](/javascript/api/excel/excel.style)|_Eigenschaft_ > VerticalAlignment|Die vertikale Ausrichtung für die Formatvorlage darstellt. Mögliche Werte sind: oben, zentriert, unten, Justify, verteilt.|1.7|
|[Formatvorlage](/javascript/api/excel/excel.style)|_Eigenschaft_ > WrapText|Gibt an, ob den Text im Objekt Microsoft Excel umbrochen wird.|1.7|
|[Formatvorlage](/javascript/api/excel/excel.style)|_Beziehung_ > Rahmen|Eine Auflistung der Rahmen von vier Border-Objekten, die das Format der vier Rahmen darstellen. Schreibgeschützt.|1.7|
|[Formatvorlage](/javascript/api/excel/excel.style)|_Beziehung_ > Füllung|Die Füllung der Formatvorlage. Schreibgeschützt.|1.7|
|[Formatvorlage](/javascript/api/excel/excel.style)|_Beziehung_ > font|Ein Font-Objekt, das die Schriftart der Formatvorlage darstellt. Schreibgeschützt.|1.7|
|[Formatvorlage](/javascript/api/excel/excel.style)|_Methode_ > delete()|Löscht diese Formatvorlage.|1.7|
|[styleCollection](/javascript/api/excel/excel.stylecollection)|_Eigenschaft_ > items|Eine Auflistung von Style-Objekten. Schreibgeschützt.|1.7|
|[styleCollection](/javascript/api/excel/excel.stylecollection)|_Methode_ > Add(name: string)]|Der Auflistung hinzugefügt eine neue Formatvorlage.|1.7|
|[styleCollection](/javascript/api/excel/excel.stylecollection)|_Methode_ > GetItem(name: string)|Ruft eine Formatvorlage nach Namen ab.|1.7|
|[tableChangedEventArgs](/javascript/api/excel/excel.tablechangedeventargs)|_Eigenschaft_ > address|Ruft die Adresse, die den geänderten Bereich einer Tabelle auf ein bestimmtes Arbeitsblatt darstellt.|1.7|
|[tableChangedEventArgs](/javascript/api/excel/excel.tablechangedeventargs)|_Eigenschaft_ > ChangeType|Ruft die Art der Überarbeitung, die darstellt, wie das Changed-Ereignis ausgelöst wird. Mögliche Werte sind: andere, RangeEdited, RowInserted, RowDeleted, ColumnInserted, ColumnDeleted, CellInserted, CellDeleted.|1.7|
|[tableChangedEventArgs](/javascript/api/excel/excel.tablechangedeventargs)|_Eigenschaft_ > Quelle|Ruft die Quelle des Ereignisses ab. Mögliche Werte sind: lokale, Remote.|1.7|
|[tableChangedEventArgs](/javascript/api/excel/excel.tablechangedeventargs)|_Eigenschaft_ > TableId|Ruft die Id der Tabelle, in dem die Daten geändert.|1.7|
|[tableChangedEventArgs](/javascript/api/excel/excel.tablechangedeventargs)|_Eigenschaft_ > Typ|Ruft den Typ des Ereignisses ab. Mögliche Werte sind: WorksheetDataChanged WorksheetSelectionChanged WorksheetAdded WorksheetActivated WorksheetDeactivated TableDataChanged TableSelectionChanged, WorksheetDeleted.|1.7|
|[tableChangedEventArgs](/javascript/api/excel/excel.tablechangedeventargs)|_Eigenschaft_ > WorksheetId|Ruft die Id des Arbeitsblatts, in dem die Daten geändert.|1.7|
|[tableSelectionChangedEventArgs](/javascript/api/excel/excel.tableselectionchangedeventargs)|_Eigenschaft_ > address|Dient zum Abrufen der Bereichsadresse, die den ausgewählten Bereich der Tabelle auf ein bestimmtes Arbeitsblatt darstellt.|1.7|
|[tableSelectionChangedEventArgs](/javascript/api/excel/excel.tableselectionchangedeventargs)|_Eigenschaft_ > IsInsideTable|Gibt an, wenn die Auswahl in einer Tabelle befindet, Adresse werden nicht verwendbar, wenn IsInsideTable auf false festgelegt ist.|1.7|
|[tableSelectionChangedEventArgs](/javascript/api/excel/excel.tableselectionchangedeventargs)|_Eigenschaft_ > TableId|Ruft die Id der Tabelle, in dem die Auswahl geändert.|1.7|
|[tableSelectionChangedEventArgs](/javascript/api/excel/excel.tableselectionchangedeventargs)|_Eigenschaft_ > Typ|Ruft den Typ des Ereignisses ab. Mögliche Werte sind: WorksheetDataChanged WorksheetSelectionChanged WorksheetAdded WorksheetActivated WorksheetDeactivated TableDataChanged TableSelectionChanged, WorksheetDeleted.|1.7|
|[tableSelectionChangedEventArgs](/javascript/api/excel/excel.tableselectionchangedeventargs)|_Eigenschaft_ > WorksheetId|Ruft die Id des Arbeitsblatts, in dem die Auswahl geändert.|1.7|
|[workbook](/javascript/api/excel/excel.workbook)|_Eigenschaft_ > name|Ruft den Namen der Arbeitsmappe ab. Schreibgeschützt.|1.7|
|[workbook](/javascript/api/excel/excel.workbook)|_Beziehung_ > DataConnections|Alle datenverbindungen in der Arbeitsmappe aktualisiert. Schreibgeschützt.|1.7|
|[workbook](/javascript/api/excel/excel.workbook)|_Beziehung_ > properties|Ruft die Eigenschaften der Arbeitsmappe ab. Schreibgeschützt.|1.7|
|[workbook](/javascript/api/excel/excel.workbook)|_Beziehung_ > protection|Gibt die Arbeitsmappe Protection-Objekt für eine Arbeitsmappe zurück. Schreibgeschützt.|1.7|
|[workbook](/javascript/api/excel/excel.workbook)|_Beziehung_ > Formatvorlagen|Stellt eine Sammlung von Formatvorlagen mit der Arbeitsmappe verbundene. Schreibgeschützt.|1.7|
|[workbook](/javascript/api/excel/excel.workbook)|_Methode_ > getActiveCell()|Ruft die derzeit aktive Zelle aus der Arbeitsmappe ab.|1.7|
|[workbookProtection](/javascript/api/excel/excel.workbookprotection)|_Eigenschaft_ > protected|Gibt an, ob die Arbeitsmappe geschützt ist. Schreibgeschützt. Schreibgeschützt.|1.7|
|[workbookProtection](/javascript/api/excel/excel.workbookprotection)|_Methode_ > Protect(password: string)|Schützt ein Arbeitsblatt. Schlägt fehl, wenn die Arbeitsmappe geschützt ist.|1.7|
|[workbookProtection](/javascript/api/excel/excel.workbookprotection)|_Methode_ > Unprotect(password: string)|Hebt den Schutz einer Arbeitsmappe.|1.7|
|[worksheet](/javascript/api/excel/excel.worksheet)|_Eigenschaft_ > Gitternetzlinien|Ruft ab oder legt diesen fest im Arbeitsblatt Gitternetzlinien Kennzeichnung.|1.7|
|[worksheet](/javascript/api/excel/excel.worksheet)|_Eigenschaft_ > Überschriften|Ruft ab oder legt diesen fest im Arbeitsblatt Überschriften Kennzeichnung.|1.7|
|[worksheet](/javascript/api/excel/excel.worksheet)|_Eigenschaft_ > ShowHeadings|Ruft ab oder legt diesen fest im Arbeitsblatt Überschriften Kennzeichnung.|1.7|
|[worksheet](/javascript/api/excel/excel.worksheet)|_Eigenschaft_ > StandardHeight|Gibt die Höhe Standard (Standard) aller Zeilen im Arbeitsblatt in Punkt zurück. Schreibgeschützt.|1.7|
|[worksheet](/javascript/api/excel/excel.worksheet)|_Eigenschaft_ > StandardWidth|Gibt an, oder die Standard (Standard) Breite aller Spalten im Arbeitsblatt festgelegt.|1.7|
|[worksheet](/javascript/api/excel/excel.worksheet)|_Eigenschaft_ > TabColor|Ruft ab oder legt die Farbe des Arbeitsblatts Registerkarte.|1.7|
|[worksheet](/javascript/api/excel/excel.worksheet)|_Beziehung_ > FreezePanes|Ruft ein Objekt, das verwendet werden kann, fixierte Bereiche auf dem Arbeitsblatt bearbeiten schreibgeschützt.|1.7|
|[worksheet](/javascript/api/excel/excel.worksheet)|_Methode_ > Kopie (PositionType: WorksheetPositionType RelativeTo: Arbeitsblatt)|Kopieren eines Arbeitsblatts, und platzieren Sie es an der angegebenen Position. Das kopierte Arbeitsblatt zurück.|1.7|
|[worksheet](/javascript/api/excel/excel.worksheet)|_Methode_ > GetRangeByIndexes (StartRow-Parameter: Number, StartColumn: Number, RowCount: Number, ColumnCount: Anzahl)| Ruft das Bereichsobjekt, beginnend an einem bestimmten Zeilen- und Spaltenindex und eine bestimmte Anzahl von Zeilen und Spalten umfassend ab.|1.7|
|[worksheetActivatedEventArgs](/javascript/api/excel/excel.worksheetactivatedeventargs)|_Eigenschaft_ > Typ|Ruft den Typ des Ereignisses ab. Mögliche Werte sind: WorksheetDataChanged WorksheetSelectionChanged WorksheetAdded WorksheetActivated WorksheetDeactivated TableDataChanged TableSelectionChanged, WorksheetDeleted.|1.7|
|[worksheetActivatedEventArgs](/javascript/api/excel/excel.worksheetactivatedeventargs)|_Eigenschaft_ > WorksheetId|Ruft die Id des Arbeitsblatts, das aktiviert wird.|1.7|
|[worksheetAddedEventArgs](/javascript/api/excel/excel.worksheetaddedeventargs)|_Eigenschaft_ > Quelle|Ruft die Quelle des Ereignisses ab. Mögliche Werte sind: lokale, Remote.|1.7|
|[worksheetAddedEventArgs](/javascript/api/excel/excel.worksheetaddedeventargs)|_Eigenschaft_ > Typ|Ruft den Typ des Ereignisses ab. Mögliche Werte sind: WorksheetDataChanged WorksheetSelectionChanged WorksheetAdded WorksheetActivated WorksheetDeactivated TableDataChanged TableSelectionChanged, WorksheetDeleted.|1.7|
|[worksheetAddedEventArgs](/javascript/api/excel/excel.worksheetaddedeventargs)|_Eigenschaft_ > WorksheetId|Ruft die Id des Arbeitsblatts, das die Arbeitsmappe hinzugefügt wird.|1.7|
|[worksheetChangedEventArgs](/javascript/api/excel/excel.worksheetchangedeventargs)|_Eigenschaft_ > address|Dient zum Abrufen der Bereichsadresse übernommen, die das geänderte Bereich, der ein bestimmtes Arbeitsblatt darstellt.|1.7|
|[worksheetChangedEventArgs](/javascript/api/excel/excel.worksheetchangedeventargs)|_Eigenschaft_ > ChangeType|Ruft die Art der Überarbeitung, die darstellt, wie das Changed-Ereignis ausgelöst wird. Mögliche Werte sind: andere, RangeEdited, RowInserted, RowDeleted, ColumnInserted, ColumnDeleted, CellInserted, CellDeleted.|1.7|
|[worksheetChangedEventArgs](/javascript/api/excel/excel.worksheetchangedeventargs)|_Eigenschaft_ > Quelle|Ruft die Quelle des Ereignisses ab. Mögliche Werte sind: lokale, Remote.|1.7|
|[worksheetChangedEventArgs](/javascript/api/excel/excel.worksheetchangedeventargs)|_Eigenschaft_ > Typ|Ruft den Typ des Ereignisses ab. Mögliche Werte sind: WorksheetDataChanged WorksheetSelectionChanged WorksheetAdded WorksheetActivated WorksheetDeactivated TableDataChanged TableSelectionChanged, WorksheetDeleted.|1.7|
|[worksheetChangedEventArgs](/javascript/api/excel/excel.worksheetchangedeventargs)|_Eigenschaft_ > WorksheetId|Ruft die Id des Arbeitsblatts, in dem die Daten geändert.|1.7|
|[worksheetDeactivatedEventArgs](/javascript/api/excel/excel.worksheetdeactivatedeventargs)|_Eigenschaft_ > Typ|Ruft den Typ des Ereignisses ab. Mögliche Werte sind: WorksheetDataChanged WorksheetSelectionChanged WorksheetAdded WorksheetActivated WorksheetDeactivated TableDataChanged TableSelectionChanged, WorksheetDeleted.|1.7|
|[worksheetDeactivatedEventArgs](/javascript/api/excel/excel.worksheetdeactivatedeventargs)|_Eigenschaft_ > WorksheetId|Ruft die Id des Arbeitsblatts, das deaktiviert wird.|1.7|
|[worksheetDeletedEventArgs](/javascript/api/excel/excel.worksheetdeletedeventargs)|_Eigenschaft_ > Quelle|Ruft die Quelle des Ereignisses ab. Mögliche Werte sind: lokale, Remote.|1.7|
|[worksheetDeletedEventArgs](/javascript/api/excel/excel.worksheetdeletedeventargs)|_Eigenschaft_ > Typ|Ruft den Typ des Ereignisses ab. Mögliche Werte sind: WorksheetDataChanged WorksheetSelectionChanged WorksheetAdded WorksheetActivated WorksheetDeactivated TableDataChanged TableSelectionChanged, WorksheetDeleted.|1.7|
|[worksheetDeletedEventArgs](/javascript/api/excel/excel.worksheetdeletedeventargs)|_Eigenschaft_ > WorksheetId|Ruft die Id des Arbeitsblatts, das aus der Arbeitsmappe gelöscht wird.|1.7|
|[worksheetFreezePanes](/javascript/api/excel/excel.worksheetfreezepanes)|_Methode_ > FreezeAt (FrozenRange: Bereich oder Zeichenfolge)|Der fixierten Zellen festgelegt in der Ansicht des aktiven Arbeitsblatts.|1.7|
|[worksheetFreezePanes](/javascript/api/excel/excel.worksheetfreezepanes)|_Methode_ > FreezeColumns(count: number)|Fixieren Sie die ersten Spalten des Arbeitsblatts direkten.|1.7|
|[worksheetFreezePanes](/javascript/api/excel/excel.worksheetfreezepanes)|_Methode_ > FreezeRows(count: number)|Fixieren Sie die oberen Zeilen des Arbeitsblatts direkten.|1.7|
|[worksheetFreezePanes](/javascript/api/excel/excel.worksheetfreezepanes)|_Methode_ > getLocation()|Dient zum Abrufen eines Bereichs, das die fixierten Zellen in der Ansicht des aktiven Arbeitsblatts beschreibt.|1.7|
|[worksheetFreezePanes](/javascript/api/excel/excel.worksheetfreezepanes)|_Methode_ > getLocationOrNullObject()|Dient zum Abrufen eines Bereichs, das die fixierten Zellen in der Ansicht des aktiven Arbeitsblatts beschreibt.|1.7|
|[worksheetFreezePanes](/javascript/api/excel/excel.worksheetfreezepanes)|_Methode_ > unfreeze()|Entfernt alle fixierten Fensterbereiche in das Arbeitsblatt an.|1.7|
|[worksheetProtectionOptions](/javascript/api/excel/excel.worksheetprotectionoptions)|_Eigenschaft_ > AllowEditObjects|Stellt die Option Arbeitsblatt Protection Bearbeiten von Objekten zu ermöglichen.|1.7|
|[worksheetProtectionOptions](/javascript/api/excel/excel.worksheetprotectionoptions)|_Eigenschaft_ > AllowEditScenarios|Stellt die Option Arbeitsblatt Protection bearbeiten Szenarien zu ermöglichen.|1.7|
|[worksheetProtectionOptions](/javascript/api/excel/excel.worksheetprotectionoptions)|_Beziehung_ > SelectionMode|Stellt die Option Arbeitsblatt Schutz der Auswahlmodus an.|1.7|
|[worksheetSelectionChangedEventArgs](/javascript/api/excel/excel.worksheetselectionchangedeventargs)|_Eigenschaft_ > address|Dient zum Abrufen der Bereichsadresse übernommen, die im markierte Bereich des ein bestimmtes Arbeitsblatt darstellt.|1.7|
|[worksheetSelectionChangedEventArgs](/javascript/api/excel/excel.worksheetselectionchangedeventargs)|_Eigenschaft_ > Typ|Ruft den Typ des Ereignisses ab. Mögliche Werte sind: WorksheetDataChanged WorksheetSelectionChanged WorksheetAdded WorksheetActivated WorksheetDeactivated TableDataChanged TableSelectionChanged, WorksheetDeleted.|1.7|
|[worksheetSelectionChangedEventArgs](/javascript/api/excel/excel.worksheetselectionchangedeventargs)|_Eigenschaft_ > WorksheetId|Ruft die Id des Arbeitsblatts, in dem die Auswahl geändert.|1.7|


## <a name="whats-new-in-excel-javascript-api-16"></a>Was ist neu in Excel JavaScript-API 1.6 

### <a name="conditional-formatting"></a>Bedingte Formatierung

Führt die bedingte Formatierung eines Bereichs. Die folgenden Typen von bedingter Formatierung ermöglicht:

* Farbskala
* Datenbalken
* Symbolsatz
* Benutzerdefiniert

Weitere Schritte:

* Gibt den Bereich an, auf dem die bedingte Formatierung angewendet wird. 
* Entfernen der bedingten Formatierung. 
* Priorität und StopIfTrue-Funktion enthält. 
* Abrufen der Auflistung der gesamten bedingten Formatierung für einen bestimmten Bereich. 
* Löscht alle bedingten Formate, die im aktuellen angegebenen Bereich aktiv sind. 

|Objekt| Neuigkeiten| Beschreibung|Anforderungssatz|
|:----|:----|:----|:----|
|[application](/javascript/api/excel/excel.application)|_Methode_ > suspendApiCalculationUntilNextSync()|Hält die Berechnung an, bis die nächste „context.sync()“ aufgerufen wird. Nachdem dies festgelegt wurde, liegt es in der Verantwortung des Entwicklers, die Arbeitsmappe neu zu berechnen, um sicherzustellen, dass alle Abhängigkeiten verteilt werden.|1.6|
|[cellValueConditionalFormat](/javascript/api/excel/excel.cellvalueconditionalformat)|_Beziehung_ > Format|Gibt ein Formatobjekt zurück, das die Schriftart, die Füllung, den Rahmen und andere Eigenschaften bedingter Formate kapselt. Schreibgeschützt.|1.6|
|[cellValueConditionalFormat](/javascript/api/excel/excel.cellvalueconditionalformat)|_Beziehung_ > Regel|Stellt das Regelobjekt in diesem bedingte Format dar.|1.6|
|[colorScaleConditionalFormat](/javascript/api/excel/excel.colorscaleconditionalformat)|_Eigenschaft_ > ThreeColorScale|Wenn „true“, hat die Farbskala drei Punkte (Minimum, Mittelpunkt, Maximum), andernfalls hat sie zwei (Minimum, Maximum). Schreibgeschützt.|1.6|
|[colorScaleConditionalFormat](/javascript/api/excel/excel.colorscaleconditionalformat)|_Beziehung_ > criteria|Die Kriterien der Farbskala. Der Mittelpunkt ist bei Verwendung einer 2-Punkte-Farbskala optional.|1.6|
|[conditionalCellValueRule](/javascript/api/excel/excel.conditionalcellvaluerule)|_Eigenschaft_ > formula1|Die Formel, sofern erforderlich, nach der die bedingte Formatregel ausgewertet werden soll.|1.6|
|[conditionalCellValueRule](/javascript/api/excel/excel.conditionalcellvaluerule)|_Eigenschaft_ > formula2|Die Formel, sofern erforderlich, nach der die bedingte Formatregel ausgewertet werden soll.|1.6|
|[conditionalCellValueRule](/javascript/api/excel/excel.conditionalcellvaluerule)|_Eigenschaft_ > operator|Der Operator des bedingte Textformats. Die folgenden Werte sind möglich: Invalid, Between, NotBetween, EqualTo, NotEqualTo, GreaterThan, LessThan, GreaterThanOrEqual, LessThanOrEqual.|1.6|
|[conditionalColorScaleCriteria](/javascript/api/excel/excel.conditionalcolorscalecriteria)|_Beziehung_ > maximale|Das Maximalpunktkriterium der Farbskala.|1.6|
|[conditionalColorScaleCriteria](/javascript/api/excel/excel.conditionalcolorscalecriteria)|_Beziehung_ > Mittelpunkt|Das Mittelpunktkriterium der Farbskala im Fall einer 3-Punkte-Farbskala.|1.6|
|[conditionalColorScaleCriteria](/javascript/api/excel/excel.conditionalcolorscalecriteria)|_Beziehung_ > minimale|Das Minimumpunktkriterium der Farbskala.|1.6|
|[conditionalColorScaleCriterion](/javascript/api/excel/excel.conditionalcolorscalecriterion)|_Eigenschaft_ > color|HTML-Farbcodedarstellung der Farbskalafarbe. #ff0000 stellt beispielsweise Rot dar.|1.6|
|[conditionalColorScaleCriterion](/javascript/api/excel/excel.conditionalcolorscalecriterion)|_Eigenschaft_ > Formel|Eine Zahl, eine Formel oder NULL (für Typ „LowestValue“).|1.6|
|[conditionalColorScaleCriterion](/javascript/api/excel/excel.conditionalcolorscalecriterion)|_Eigenschaft_ > Typ|Die Basis für die bedingte Symbolformel. Die folgenden Werte sind möglich: Invalid, LowestValue, HighestValue, Number, Percent, Formula, Percentile.|1.6|
|[conditionalDataBarNegativeFormat](/javascript/api/excel/excel.conditionaldatabarnegativeformat)|_Eigenschaft_ > BorderColor-Eigenschaft|HTML-Farbcode, der die Farbe der Rahmenlinie, des Formulars #RRGGBB (z. B.  „FFA500“) oder als benannte HTML-Farbe (z. B. „orange“) darstellt.|1.6|
|[conditionalDataBarNegativeFormat](/javascript/api/excel/excel.conditionaldatabarnegativeformat)|_Eigenschaft_ > FillColor|HTML-Farbcode, der die Füllfarbe im Format #RRGGBB (z. B.  „FFA500“) oder als benannte HTML-Farbe (z. B. „orange“) darstellt.|1.6|
|[conditionalDataBarNegativeFormat](/javascript/api/excel/excel.conditionaldatabarnegativeformat)|_Eigenschaft_ > MatchPositiveBorderColor|Boolesche Darstellung zur Angabe, ob der negative Datenbalken dieselbe Rahmenfarbe wie der Datenbalken hat.|1.6|
|[conditionalDataBarNegativeFormat](/javascript/api/excel/excel.conditionaldatabarnegativeformat)|_Eigenschaft_ > MatchPositiveFillColor|Boolesche Darstellung zur Angabe, ob der negative Datenbalken dieselbe Füllfarbe wie der Datenbalken hat.|1.6|
|[conditionalDataBarPositiveFormat](/javascript/api/excel/excel.conditionaldatabarpositiveformat)|_Eigenschaft_ > BorderColor-Eigenschaft|HTML-Farbcode, der die Farbe der Rahmenlinie, des Formulars #RRGGBB (z. B.  „FFA500“) oder als benannte HTML-Farbe (z. B. „orange“) darstellt.|1.6|
|[conditionalDataBarPositiveFormat](/javascript/api/excel/excel.conditionaldatabarpositiveformat)|_Eigenschaft_ > FillColor|HTML-Farbcode, der die Füllfarbe im Format #RRGGBB (z. B.  „FFA500“) oder als benannte HTML-Farbe (z. B. „orange“) darstellt.|1.6|
|[conditionalDataBarPositiveFormat](/javascript/api/excel/excel.conditionaldatabarpositiveformat)|_Eigenschaft_ > GradientFill|Boolesche Darstellung der Angabe, ob der Datenbalken einen Farbverlauf hat.|1.6|
|[conditionalDataBarRule](/javascript/api/excel/excel.conditionaldatabarrule)|_Eigenschaft_ > Formel|Die Formel, sofern erforderlich, nach der die Datenbalkenregel ausgewertet werden soll.|1.6|
|[conditionalDataBarRule](/javascript/api/excel/excel.conditionaldatabarrule)|_Eigenschaft_ > Typ|Der Typ der Regel für den Datenbalken. Die folgenden Werte sind möglich: LowestValue, HighestValue, Number, Percent, Formula, Percentile, Automatic.|1.6|
|[Das conditionalFormat](/javascript/api/excel/excel.conditionalformat)|_Eigenschaft_ > id|Die Priorität der bedingten Formatierung in der aktuellen ConditionalFormatCollection. Schreibgeschützt.|1.6|
|[Das conditionalFormat](/javascript/api/excel/excel.conditionalformat)|_Eigenschaft_ > Priorität|Die Priorität (oder der Index) in der bedingten Formatsammlung, in der dieses bedingte Format derzeit enthalten ist. Bei einer Änderung|1.6|
|[Das conditionalFormat](/javascript/api/excel/excel.conditionalformat)|_Eigenschaft_ > StopIfTrue|Wenn die Bedingungen dieses bedingten Formats erfüllt sind, werden keine Formate niedrigerer Priorität für diese Zelle wirksam.|1.6|
|[Das conditionalFormat](/javascript/api/excel/excel.conditionalformat)|_Eigenschaft_ > Typ|Ein bedingter Formattyp. Es kann nur jeweils ein Typ festgelegt werden. Schreibgeschützt. Schreibgeschützt. Die folgenden Werte sind möglich: Custom, DataBar, ColorScale, IconSet.|1.6|
|[Das conditionalFormat](/javascript/api/excel/excel.conditionalformat)|_Beziehung_ > CellValue|Gibt die Zellenwerteigenschaften des bedingten Formats zurück, wenn das aktuelle bedingte Format einen CellValue-Typ aufweist. Schreibgeschützt.|1.6|
|[Das conditionalFormat](/javascript/api/excel/excel.conditionalformat)|_Beziehung_ > CellValueOrNullObject|Gibt die Zellenwerteigenschaften des bedingten Formats zurück, wenn das aktuelle bedingte Format einen CellValue-Typ aufweist. Schreibgeschützt.|1.6|
|[Das conditionalFormat](/javascript/api/excel/excel.conditionalformat)|_Beziehung_ > ColorScale|Gibt die ColorScale-Eigenschaften des bedingten Formats zurück, wenn das aktuelle bedingte Format einen ColorScale-Typ aufweist. Schreibgeschützt.|1.6|
|[Das conditionalFormat](/javascript/api/excel/excel.conditionalformat)|_Beziehung_ > ColorScaleOrNullObject|Gibt die ColorScale-Eigenschaften des bedingten Formats zurück, wenn das aktuelle bedingte Format einen ColorScale-Typ aufweist. Schreibgeschützt.|1.6|
|[Das conditionalFormat](/javascript/api/excel/excel.conditionalformat)|_Beziehung_ > benutzerdefinierte|Gibt die benutzerdefinierten Eigenschaften des bedingten Formats zurück, wenn das aktuelle bedingte Format einen custom-Typ aufweist. Schreibgeschützt.|1.6|
|[Das conditionalFormat](/javascript/api/excel/excel.conditionalformat)|_Beziehung_ > CustomOrNullObject|Gibt die benutzerdefinierten Eigenschaften des bedingten Formats zurück, wenn das aktuelle bedingte Format einen custom-Typ aufweist. Schreibgeschützt.|1.6|
|[Das conditionalFormat](/javascript/api/excel/excel.conditionalformat)|_Beziehung_ > DataBar|Gibt die Datenbalkeneigenschaften zurück, wenn das aktuelle bedingte Format ein Datenbalken ist. Schreibgeschützt.|1.6|
|[Das conditionalFormat](/javascript/api/excel/excel.conditionalformat)|_Beziehung_ > DataBarOrNullObject|Gibt die Datenbalkeneigenschaften zurück, wenn das aktuelle bedingte Format ein Datenbalken ist. Schreibgeschützt.|1.6|
|[Das conditionalFormat](/javascript/api/excel/excel.conditionalformat)|_Beziehung_ > IconSet|Gibt die IconSet-Eigenschaften des bedingten Formats zurück, wenn das aktuelle bedingte Format einen IconSet-Typ aufweist. Schreibgeschützt.|1.6|
|[Das conditionalFormat](/javascript/api/excel/excel.conditionalformat)|_Beziehung_ > IconSetOrNullObject|Gibt die IconSet-Eigenschaften des bedingten Formats zurück, wenn das aktuelle bedingte Format einen IconSet-Typ aufweist. Schreibgeschützt.|1.6|
|[Das conditionalFormat](/javascript/api/excel/excel.conditionalformat)|_Beziehung_ > voreingestellte|Gibt die voreingestellten Kriterien für das bedingte Format zurück, wie die Eigenschaften „above average“, „below average“, „unique values“, „contains blank“, „nonblank“, „error“, „noerror“. Schreibgeschützt.|1.6|
|[Das conditionalFormat](/javascript/api/excel/excel.conditionalformat)|_Beziehung_ > PresetOrNullObject|Gibt die voreingestellten Kriterien für das bedingte Format zurück, wie die Eigenschaften „above average“, „below average“, „unique values“, „contains blank“, „nonblank“, „error“, „noerror“. Schreibgeschützt.|1.6|
|[Das conditionalFormat](/javascript/api/excel/excel.conditionalformat)|_Beziehung_ > TextComparison|Gibt die spezifischen Texteigenschaften des bedingten Formats zurück, wenn das aktuelle bedingte Format einen Text-Typ aufweist. Schreibgeschützt.|1.6|
|[Das conditionalFormat](/javascript/api/excel/excel.conditionalformat)|_Beziehung_ > TextComparisonOrNullObject|Gibt die spezifischen Texteigenschaften des bedingten Formats zurück, wenn das aktuelle bedingte Format einen Text-Typ aufweist. Schreibgeschützt.|1.6|
|[Das conditionalFormat](/javascript/api/excel/excel.conditionalformat)|_Beziehung_ > TopBottom|Gibt die TopBottom-Eigenschaften des bedingten Formats zurück, wenn das aktuelle bedingte Format einen TopBottom-Typ aufweist. Schreibgeschützt.|1.6|
|[Das conditionalFormat](/javascript/api/excel/excel.conditionalformat)|_Beziehung_ > TopBottomOrNullObject|Gibt die TopBottom-Eigenschaften des bedingten Formats zurück, wenn das aktuelle bedingte Format einen TopBottom-Typ aufweist. Schreibgeschützt.|1.6|
|[Das conditionalFormat](/javascript/api/excel/excel.conditionalformat)|_Methode_ > delete()|Löscht dieses bedingte Format.|1.6|
|[Das conditionalFormat](/javascript/api/excel/excel.conditionalformat)|_Methode_ > getRange()|Gibt den Bereich zurück, auf den das bedingte Format angewendet wird, oder ein NULL-Objekt, wenn der Bereich nicht zusammenhängend ist. Schreibgeschützt.|1.6|
|[Das conditionalFormat](/javascript/api/excel/excel.conditionalformat)|_Methode_ > getRangeOrNullObject()|Gibt den Bereich zurück, auf den das bedingte Format angewendet wird, oder ein NULL-Objekt, wenn der Bereich nicht zusammenhängend ist. Schreibgeschützt.|1.6|
|[conditionalFormatCollection](/javascript/api/excel/excel.conditionalformatcollection)|_Eigenschaft_ > items|Eine Sammlung von conditionalFormat-Objekten. Schreibgeschützt.|1.6|
|[conditionalFormatCollection](/javascript/api/excel/excel.conditionalformatcollection)|_Methode_ > Add(type: string)|Fügt ein neues bedingtes Format zur Sammlung an der Firsttop-Priorität hinzu.|1.6|
|[conditionalFormatCollection](/javascript/api/excel/excel.conditionalformatcollection)|_Methode_ > clearAll()|Löscht alle bedingten Formate, die im aktuellen angegebenen Bereich aktiv sind.|1.6|
|[conditionalFormatCollection](/javascript/api/excel/excel.conditionalformatcollection)|_Methode_ > getCount()|Gibt die Anzahl der bedingten Formate in der Arbeitsmappe zurück. Schreibgeschützt.|1.6|
|[conditionalFormatCollection](/javascript/api/excel/excel.conditionalformatcollection)|_Methode_ > GetItem(id: string)|Gibt ein bedingtes Format für die angegebene ID|1.6|
|[conditionalFormatCollection](/javascript/api/excel/excel.conditionalformatcollection)|_Methode_ > GetItemAt(index: number)|Gibt ein bedingtes Format am angegebenen Index zurück.|1.6|
|[conditionalFormatRule](/javascript/api/excel/excel.conditionalformatrule)|_Eigenschaft_ > Formel|Die Formel, sofern erforderlich, nach der die bedingte Formatregel ausgewertet werden soll.|1.6|
|[conditionalFormatRule](/javascript/api/excel/excel.conditionalformatrule)|_Eigenschaft_ > FormulaLocal|Die Formel in der Benutzersprache, sofern erforderlich, nach der die bedingte Formatregel ausgewertet werden soll.|1.6|
|[conditionalFormatRule](/javascript/api/excel/excel.conditionalformatrule)|_Eigenschaft_ > formulaR1C1|Die Formel in R1C1-Notation, sofern erforderlich, nach der die bedingte Formatregel ausgewertet werden soll.|1.6|
|[conditionalIconCriterion](/javascript/api/excel/excel.conditionaliconcriterion)|_Eigenschaft_ > Formel|Eine Zahl oder eine Formel, je nach Typ.|1.6|
|[conditionalIconCriterion](/javascript/api/excel/excel.conditionaliconcriterion)|_Eigenschaft_ > operator|GreaterThan oder GreaterThanOrEqual für jeden der Regeltypen für das bedingte Symbolformat. Die folgenden Werte sind möglich: Invalid, GreaterThan, GreaterThanOrEqual.|1.6|
|[conditionalIconCriterion](/javascript/api/excel/excel.conditionaliconcriterion)|_Beziehung_ > CustomIcon|Das benutzerdefinierte Symbol für das aktuelle Kriterium, sofern verschieden vom Standard-IconSet; andernfalls wird NULL zurückgegeben.|1.6|
|[conditionalIconCriterion](/javascript/api/excel/excel.conditionaliconcriterion)|_Beziehung_ > type|Die Basis für die bedingte Symbolformel.|1.6|
|[conditionalPresetCriteriaRule](/javascript/api/excel/excel.conditionalpresetcriteriarule)|_Eigenschaft_ > Kriterium|Das Kriterium des bedingten Formats. Die folgenden Werte sind möglich: Invalid, Blanks, NonBlanks, Errors, NonErrors, Yesterday, Today, Tomorrow, LastSevenDays, LastWeek, ThisWeek, NextWeek, LastMonth, ThisMonth, NextMonth, AboveAverage, BelowAverage, EqualOrAboveAverage, EqualOrBelowAverage, OneStdDevAboveAverage, OneStdDevBelowAverage, TwoStdDevAboveAverage, TwoStdDevBelowAverage, ThreeStdDevAboveAverage, ThreeStdDevBelowAverage, UniqueValues, DuplicateValues.|1.6|
|[conditionalRangeBorder](/javascript/api/excel/excel.conditionalrangeborder)|_Eigenschaft_ > color|HTML-Farbcode, der die Farbe der Rahmenlinie, des Formulars #RRGGBB (z. B.  „FFA500“) oder als benannte HTML-Farbe (z. B. „orange“) darstellt.|1.6|
|[conditionalRangeBorder](/javascript/api/excel/excel.conditionalrangeborder)|_Eigenschaft_ > id|Stellt die Rahmen-ID dar. Schreibgeschützt. Die folgenden Werte sind möglich: EdgeTop, EdgeBottom, EdgeLeft, EdgeRight.|1.6|
|[conditionalRangeBorder](/javascript/api/excel/excel.conditionalrangeborder)|_Eigenschaft_ > SideIndex|Konstanter Wert, der die bestimmte Seiten des Rahmens angibt. Schreibgeschützt. Die folgenden Werte sind möglich: EdgeTop, EdgeBottom, EdgeLeft, EdgeRight.|1.6|
|[conditionalRangeBorder](/javascript/api/excel/excel.conditionalrangeborder)|_Eigenschaft_ > style|Einer der Konstanten der Linienart, die die Linienart für den Rahmen angibt. Die folgenden Werte sind möglich: None, Continuous, Dash, DashDot, DashDotDot, Dot, Double, SlantDashDot.|1.6|
|[conditionalRangeBorderCollection](/javascript/api/excel/excel.conditionalrangebordercollection)|_Eigenschaft_ > Count|Die Anzahl der Rahmen-Objekte in der Auflistung. Schreibgeschützt.|1.6|
|[conditionalRangeBorderCollection](/javascript/api/excel/excel.conditionalrangebordercollection)|_Eigenschaft_ > items|Eine Sammlung von conditionalRangeBorder-Objekten. Schreibgeschützt.|1.6|
|[conditionalRangeBorderCollection](/javascript/api/excel/excel.conditionalrangebordercollection)|_Beziehung_ > unten|Ruft den oberen Rahmen ab. Schreibgeschützt.|1.6|
|[conditionalRangeBorderCollection](/javascript/api/excel/excel.conditionalrangebordercollection)|_Beziehung_ > linken|Ruft den oberen Rahmen ab. Schreibgeschützt.|1.6|
|[conditionalRangeBorderCollection](/javascript/api/excel/excel.conditionalrangebordercollection)|_Beziehung_ > rechten|Ruft den oberen Rahmen ab. Schreibgeschützt.|1.6|
|[conditionalRangeBorderCollection](/javascript/api/excel/excel.conditionalrangebordercollection)|_Beziehung_ > oben|Ruft den oberen Rahmen ab. Schreibgeschützt.|1.6|
|[conditionalRangeBorderCollection](/javascript/api/excel/excel.conditionalrangebordercollection)|_Methode_ > GetItem(index: string)|Ruft ein Rahmen-Objekt ab, das den Namen verwendet|1.6|
|[conditionalRangeBorderCollection](/javascript/api/excel/excel.conditionalrangebordercollection)|_Methode_ > GetItemAt(index: number)|Dient zum Abrufen eines Rahmenobjekts mithilfe seines Index.|1.6|
|[conditionalRangeFill](/javascript/api/excel/excel.conditionalrangefill)|_Eigenschaft_ > color|HTML-Farbcode, der die Füllfarbe im Format #RRGGBB (z. B.  „FFA500“) oder als benannte HTML-Farbe (z. B. „orange“) darstellt.|1.6|
|[conditionalRangeFill](/javascript/api/excel/excel.conditionalrangefill)|_Methode_ > Löschvorgang()|Setzt die Füllung zurück.|1.6|
|[conditionalRangeFont](/javascript/api/excel/excel.conditionalrangefont)|_Eigenschaft_ > fett|Stellt den Fett-Status der Schriftart dar.|1.6|
|[conditionalRangeFont](/javascript/api/excel/excel.conditionalrangefont)|_Eigenschaft_ > color|HTML-Farbcodedarstellung der Textfarbe. #ff0000 stellt beispielsweise Rot dar.|1.6|
|[conditionalRangeFont](/javascript/api/excel/excel.conditionalrangefont)|_Eigenschaft_ > kursiv|Stellt den Kursiv-Status der Schriftart dar.|1.6|
|[conditionalRangeFont](/javascript/api/excel/excel.conditionalrangefont)|_Eigenschaft_ > durchgestrichen|Stellt den Durchgestrichen-Status der Schriftart dar.|1.6|
|[conditionalRangeFont](/javascript/api/excel/excel.conditionalrangefont)|_Eigenschaft_ > Unterstreichung|Art der auf die Schriftart angewendete Unterstreichung. Die folgenden Werte sind möglich: None, Single, Double.|1.6|
|[conditionalRangeFont](/javascript/api/excel/excel.conditionalrangefont)|_Methode_ > Löschvorgang()|Setzt die Schriftartformate zurück.|1.6|
|[conditionalRangeFormat](/javascript/api/excel/excel.conditionalrangeformat)|_Eigenschaft_ > numberFormat|Stellt den Excel-Zahlenformatcode für den angegebenen Bereich dar. Deaktiviert, wenn NULL übergeben wird.|1.6|
|[conditionalRangeFormat](/javascript/api/excel/excel.conditionalrangeformat)|_Beziehung_ > Rahmen|Sammlung von Rahmenobjekten, die für den gesamten bedingten Formatbereich gelten. Schreibgeschützt.|1.6|
|[conditionalRangeFormat](/javascript/api/excel/excel.conditionalrangeformat)|_Beziehung_ > Füllung|Gibt das Füllungsobjekt zurück, das für den gesamten bedingten Formatbereich definiert ist. Schreibgeschützt.|1.6|
|[conditionalRangeFormat](/javascript/api/excel/excel.conditionalrangeformat)|_Beziehung_ > font|Gibt das Schriftartobjekt zurück, das für den gesamten bedingten Formatbereich definiert ist. Schreibgeschützt.|1.6|
|[conditionalTextComparisonRule](/javascript/api/excel/excel.conditionaltextcomparisonrule)|_Eigenschaft_ > operator|Der Operator des bedingte Textformats. Die folgenden Werte sind möglich: Invalid, Contains, NotContains, BeginsWith, EndsWith.|1.6|
|[conditionalTextComparisonRule](/javascript/api/excel/excel.conditionaltextcomparisonrule)|_Eigenschaft_ > text|Der Textwert des bedingten Formats.|1.6|
|[conditionalTopBottomRule](/javascript/api/excel/excel.conditionaltopbottomrule)|_Eigenschaft_ > Rang|Der Rang zwischen 1 und 1000 für numerische Ränge oder 1 und 100 als Prozentränge.|1.6|
|[conditionalTopBottomRule](/javascript/api/excel/excel.conditionaltopbottomrule)|_Eigenschaft_ > Typ|Formatieren von Werten basierend auf dem oberen oder unteren Rang. Die folgenden Werte sind möglich: Invalid, TopItems, TopPercent, BottomItems, BottomPercent.|1.6|
|[customConditionalFormat](/javascript/api/excel/excel.customconditionalformat)|_Beziehung_ > Format|Gibt ein Formatobjekt zurück, das die Schriftart, die Füllung, den Rahmen und andere Eigenschaften bedingter Formate kapselt. Schreibgeschützt.|1.6|
|[customConditionalFormat](/javascript/api/excel/excel.customconditionalformat)|_Beziehung_ > Regel|Stellt das Regelobjekt in diesem bedingte Format dar. Schreibgeschützt.|1.6|
|[dataBarConditionalFormat](/javascript/api/excel/excel.databarconditionalformat)|_Eigenschaft_ > AxisColor|HTML-Farbcode, der die Farbe der Achsenlinie im Format #RRGGBB (z. B.  „FFA500“) oder als benannte HTML-Farbe (z. B. „orange“) darstellt.|1.6|
|[dataBarConditionalFormat](/javascript/api/excel/excel.databarconditionalformat)|_Eigenschaft_ > AxisFormat|Darstellung, wie die Achse für einen Excel-Datenbalken bestimmt wird. Die folgenden Werte sind möglich: Automatic, None, CellMidPoint.|1.6|
|[dataBarConditionalFormat](/javascript/api/excel/excel.databarconditionalformat)|_Eigenschaft_ > BarDirection|Stellt die Richtung dar, auf der die Datenbalkengrafik basieren soll. Die folgenden Werte sind möglich: Context, LeftToRight, RightToLeft.|1.6|
|[dataBarConditionalFormat](/javascript/api/excel/excel.databarconditionalformat)|_Eigenschaft_ > ShowDataBarOnly|Wenn „true“, werden die Werte in den Zellen ausgeblendet, auf die der Datenbalken angewendet wird.|1.6|
|[dataBarConditionalFormat](/javascript/api/excel/excel.databarconditionalformat)|_Beziehung_ > LowerBoundRule|Die Regel, die bestimmt, was die untere Grenze für einen Datenbalken darstellt (und wie diese ggf. berechnet wird).|1.6|
|[dataBarConditionalFormat](/javascript/api/excel/excel.databarconditionalformat)|_Beziehung_ > NegativeFormat|Darstellung aller Werte auf der linken Seite der Achse in einem Excel-Datenbalken. Schreibgeschützt.|1.6|
|[dataBarConditionalFormat](/javascript/api/excel/excel.databarconditionalformat)|_Beziehung_ > PositiveFormat|Darstellung aller Werte auf der rechten Seite der Achse in einem Excel-Datenbalken. Schreibgeschützt.|1.6|
|[dataBarConditionalFormat](/javascript/api/excel/excel.databarconditionalformat)|_Beziehung_ > UpperBoundRule|Die Regel, die bestimmt, was die obere Grenze für einen Datenbalken darstellt (und wie diese ggf. berechnet wird).|1.6|
|[iconSetConditionalFormat](/javascript/api/excel/excel.iconsetconditionalformat)|_Eigenschaft_ > ReverseIconOrder|Wenn „true“, wird die Symbolreihenfolgt für das IconSet umgekehrt. Beachten Sie, dass dies bei Verwendung benutzerdefinierter Symbole nicht festgelegt werden kann.|1.6|
|[iconSetConditionalFormat](/javascript/api/excel/excel.iconsetconditionalformat)|_Eigenschaft_ > ShowIconOnly|Wenn „true“, werden die Werte ausgeblendet und nur die Symbole angezeigt.|1.6|
|[iconSetConditionalFormat](/javascript/api/excel/excel.iconsetconditionalformat)|_Eigenschaft_ > style|Wenn festgelegt, wird die IconSet-Option für das bedingte Format angezeigt. Die folgenden Werte sind möglich: Invalid, ThreeArrows, ThreeArrowsGray, ThreeFlags, ThreeTrafficLights1, ThreeTrafficLights2, ThreeSigns, ThreeSymbols, ThreeSymbols2, FourArrows, FourArrowsGray, FourRedToBlack, FourRating, FourTrafficLights, FiveArrows, FiveArrowsGray, FiveRating, FiveQuarters, ThreeStars, ThreeTriangles, FiveBoxes.|1.6|
|[iconSetConditionalFormat](/javascript/api/excel/excel.iconsetconditionalformat)|_Beziehung_ > criteria|Ein Array von Kriterien und IconSets für die Regeln und potenziellen benutzerdefinierte Symbole für bedingte Symbole. Beachten Sie, dass für das erste Kriterium nur das benutzerdefinierte Symbol geändert werden kann, während Typ, Formel und Operator ignoriert werden, wenn festgelegt.|1.6|
|[presetCriteriaConditionalFormat](/javascript/api/excel/excel.presetcriteriaconditionalformat)|_Beziehung_ > Format|Gibt ein Formatobjekt zurück, das die Schriftart, die Füllung, den Rahmen und andere Eigenschaften bedingter Formate kapselt. Schreibgeschützt.|1.6|
|[presetCriteriaConditionalFormat](/javascript/api/excel/excel.presetcriteriaconditionalformat)|_Beziehung_ > Regel|Die Regel des bedingten Formats.|1.6|
|[range](/javascript/api/excel/excel.range)|_Beziehung_ > ConditionalFormats|Sammlung von ConditionalFormats, die den Bereich überschneiden. Schreibgeschützt.|1.6|
|[range](/javascript/api/excel/excel.range)|_Methode_ > calculate()|Berechnet einen Zellbereich auf einem Arbeitsblatt.|1.6|
|[textConditionalFormat](/javascript/api/excel/excel.textconditionalformat)|_Beziehung_ > Format|Gibt ein Formatobjekt zurück, das die Schriftart, die Füllung, den Rahmen und andere Eigenschaften bedingter Formate kapselt. Schreibgeschützt.|1.6|
|[textConditionalFormat](/javascript/api/excel/excel.textconditionalformat)|_Beziehung_ > Regel|Die Regel des bedingten Formats.|1.6|
|[topBottomConditionalFormat](/javascript/api/excel/excel.topbottomconditionalformat)|_Beziehung_ > Format|Gibt ein Formatobjekt zurück, das die Schriftart, die Füllung, den Rahmen und andere Eigenschaften bedingter Formate kapselt. Schreibgeschützt.|1.6|
|[topBottomConditionalFormat](/javascript/api/excel/excel.topbottomconditionalformat)|_Beziehung_ > Regel|Die Kriterien des bedingten Formats „topbottom“.|1.6|
|[workbook](/javascript/api/excel/excel.workbook)|_Beziehung_ > InternalTest|Nur für internen Gebrauch. Schreibgeschützt.|1.6|
|[worksheet](/javascript/api/excel/excel.worksheet)|_Methode_ > Calculate(markAllDirty: bool)|Berechnet alle Zellen auf einem Arbeitsblatt.|1.6|

##  <a name="whats-new-in-excel-javascript-api-15"></a>Was ist neu in Excel JavaScript-API 1,5

### <a name="custom-xml-part"></a>Benutzerdefinierte XML-Komponente

* Ergänzung durch benutzerdefinierte XML-Komponentensammlung für Arbeitsmappenobjekt.
* Abrufen der benutzerdefinierten XML-Komponente anhand der ID
* Ruft eine bereichsbezogene Sammlung von benutzerdefinierten XML-Komponenten ab, deren Namespaces dem angegebenen Namespace entsprechen.
* Ruft die einer Komponente zugewiesene XML-Zeichenfolge ab.
* Stellt die ID und den Namespace einer Komponente bereit.
* Fügt der Arbeitsmappe eine neue benutzerdefinierte XML-Komponente hinzu.
* Legt die gesamte XML-Komponente fest.
* Löscht eine benutzerdefinierte XML-Komponente.
* Löscht ein Attribut mit dem angegebenen Namen aus dem Element, das durch xpath identifiziert wird.
* Fragt den XML-Inhalt anhand von xpath ab.
* Aktualisiert und löscht das Attribut und fügt es ein.

**Referenzimplementierung:** [Hier](https://github.com/mandren/Excel-CustomXMLPart-Demo) finden Sie eine Referenzimplementierung, die zeigt, wie benutzerdefinierte XML-Komponenten in einem Add-In verwendet werden können.

### <a name="others"></a>Sonstige
* `range.getSurroundingRegion()` Gibt ein Range-Objekt zurück, das die umgebenden Region für diesen Bereich darstellt. Eine umgebende Region ist ein Bereich, der von einer Kombination von leeren Zeilen und leeren Spalten relativ zu diesem Bereich begrenzt wird.
* `getNextColumn()` und `getPreviousColumn()`, `getLast() für Tabellenspalte.
* `getActiveWorksheet()` in der Arbeitsmappe.
* `getRange(address: string)` aus der Arbeitsmappe.
* `getBoundingRange(ranges: )` Ruft das kleinste Bereichsobjekt ab, das die angegebenen Bereiche umfasst. Der umgebende Bereich zwischen „B2:C5“ und „D10:E15“ lautet beispielsweise „B2:E15.“
* `getCount()` in verschiedenen Sammlungen z. B. benanntes Element, Arbeitsblatt, Tabelle usw., um die Anzahl von Elementen in einer Sammlung abzurufen. `workbook.worksheets.getCount()`
* `getFirst()` und `getLast()` und letztes Element abrufen in verschiedenen Sammlungen, z. B. Arbeitsblatt, Tabellenspalte, Diagrammpunkte, Bereichsansichtsammlung.
* `getNext()` und `getPrevious()` auf dem Arbeitsblatt, Tabellenspaltensammlung.
* `getRangeR1C1()` Ruft das Bereichsobjekt, beginnend an einem bestimmten Zeilen- und Spaltenindex und eine bestimmte Anzahl von Zeilen und Spalten umfassend ab.

|Objekt| Neuigkeiten| Beschreibung|Anforderungssatz|
|:----|:----|:----|:----|
|[CustomXMLPart-Objekt](/javascript/api/excel/excel.customxmlpart)|_Eigenschaft_ > id|Die ID der benutzerdefinierten XML-Komponente. Schreibgeschützt.|1,5|
|[CustomXMLPart-Objekt](/javascript/api/excel/excel.customxmlpart)|_Eigenschaft_ > NamespaceUri|Der Namespace-URI der benutzerdefinierten XML-Komponente. Schreibgeschützt.|1,5|
|[CustomXMLPart-Objekt](/javascript/api/excel/excel.customxmlpart)|_Methode_ > delete()|Löscht die benutzerdefinierte XML-Komponente.|1,5|
|[CustomXMLPart-Objekt](/javascript/api/excel/excel.customxmlpart)|_Methode_ > getXml()|Ruft den vollständigen XML-Inhalt der benutzerdefinierten XML-Komponente ab.|1,5|
|[CustomXMLPart-Objekt](/javascript/api/excel/excel.customxmlpart)|_Methode_ > SetXml(xml: string)|Legt den vollständigen XML-Inhalt der benutzerdefinierten XML-Komponente fest.|1,5|
|[customXmlPartCollection](/javascript/api/excel/excel.customxmlpartcollection)|_Eigenschaft_ > items|Eine Sammlung von customXmlPart-Objekten. Schreibgeschützt.|1,5|
|[customXmlPartCollection](/javascript/api/excel/excel.customxmlpartcollection)|_Methode_ > Add(xml: string)|Fügt der Arbeitsmappe eine neue benutzerdefinierte XML-Komponente hinzu.|1,5|
|[customXmlPartCollection](/javascript/api/excel/excel.customxmlpartcollection)|_Methode_ > GetByNamespace(namespaceUri: string)|Ruft eine neue bereichsbezogene Sammlung von benutzerdefinierten XML-Komponenten ab, deren Namespaces dem angegebenen Namespace entsprechen.|1,5|
|[customXmlPartCollection](/javascript/api/excel/excel.customxmlpartcollection)|_Methode_ > getCount()|Ruft die Anzahl von CustomXml-Komponenten in der Sammlung ab.|1,5|
|[customXmlPartCollection](/javascript/api/excel/excel.customxmlpartcollection)|_Methode_ > GetItem(id: string)|Ruft eine benutzerdefinierte XML-Komponente basierend auf ihrer ID ab.|1,5|
|[customXmlPartCollection](/javascript/api/excel/excel.customxmlpartcollection)|_Methode_ > GetItemOrNullObject(id: string)|Ruft eine benutzerdefinierte XML-Komponente basierend auf ihrer ID ab.|1,5|
|[customXmlPartScopedCollection](/javascript/api/excel/excel.customxmlpartscopedcollection)|_Eigenschaft_ > items|Eine Sammlung von customXmlPartScoped-Objekten. Schreibgeschützt.|1,5|
|[customXmlPartScopedCollection](/javascript/api/excel/excel.customxmlpartscopedcollection)|_Methode_ > getCount()|Ruft die Anzahl von CustomXml-Komponenten in dieser Sammlung ab.|1,5|
|[customXmlPartScopedCollection](/javascript/api/excel/excel.customxmlpartscopedcollection)|_Methode_ > GetItem(id: string)|Ruft eine benutzerdefinierte XML-Komponente basierend auf ihrer ID ab.|1,5|
|[customXmlPartScopedCollection](/javascript/api/excel/excel.customxmlpartscopedcollection)|_Methode_ > GetItemOrNullObject(id: string)|Ruft eine benutzerdefinierte XML-Komponente basierend auf ihrer ID ab.|1,5|
|[customXmlPartScopedCollection](/javascript/api/excel/excel.customxmlpartscopedcollection)|_Methode_ > getOnlyItem()|Wenn die Sammlung genau ein Element enthält, gibt diese Methode es zurück.|1,5|
|[customXmlPartScopedCollection](/javascript/api/excel/excel.customxmlpartscopedcollection)|_Methode_ > getOnlyItemOrNullObject()|Wenn die Sammlung genau ein Element enthält, gibt diese Methode es zurück.|1,5|
|[workbook](/javascript/api/excel/excel.workbook)|_Beziehung_ > CustomXmlParts|Stellt die Sammlung von benutzerdefinierten XML-Komponenten dar, die in dieser Arbeitsmappe enthalten sind. Schreibgeschützt.|1,5|
|[worksheet](/javascript/api/excel/excel.worksheet)|_Methode_ > GetNext(visibleOnly: bool)|Ruft das Arbeitsblatt nach dem aktuellen ab. Wenn es keine nachfolgenden Arbeitsblätter gibt, löst diese Methode einen Fehler aus.|1,5|
|[worksheet](/javascript/api/excel/excel.worksheet)|_Methode_ > GetNextOrNullObject(visibleOnly: bool)|Ruft das Arbeitsblatt nach dem aktuellen ab. Wenn es keine nachfolgenden Arbeitsblätter gibt, gibt diese Methode ein NULL-Objekt zurück.|1,5|
|[worksheet](/javascript/api/excel/excel.worksheet)|_Methode_ > GetPrevious(visibleOnly: bool)|Ruft das Arbeitsblatt vor dem aktuellen ab. Wenn es keine vorhergehenden Arbeitsblätter gibt, löst diese Methode einen Fehler aus.|1,5|
|[worksheet](/javascript/api/excel/excel.worksheet)|_Methode_ > GetPreviousOrNullObject(visibleOnly: bool)|Ruft das Arbeitsblatt vor dem aktuellen ab. Wenn es keine vorhergehenden Arbeitsblätter gibt, gibt diese Methode ein NULL-Objekt zurück.|1,5|
|[worksheetCollection](/javascript/api/excel/excel.worksheetcollection)|_Methode_ > GetFirst(visibleOnly: bool)|Ruft das erste Arbeitsblatt in der Sammlung ab.|1,5|
|[worksheetCollection](/javascript/api/excel/excel.worksheetcollection)|_Methode_ > GetLast(visibleOnly: bool)|Ruft das letzte Arbeitsblatt in der Sammlung ab.|1,5|

## <a name="whats-new-in-excel-javascript-api-14"></a>Neuigkeiten in der Excel-JavaScript-API 1.4
Im folgenden werden, dass die neuen Ergänzungen der Excel-JavaScript-APIs in Anforderung 1.4 festgelegt.

### <a name="named-item-add-and-new-properties"></a>Benanntes Element wurde hinzugefügt und neue Eigenschaften

Neue Eigenschaften:

* `comment`
* `scope` Auf den Bereich des Arbeitsblatt oder der Arbeitsmappe ausgelegte Elemente
* `worksheet` gibt das Arbeitsblatt zurück, in dessen Bereich sich das benannte Element befindet.

Neue Methoden:

* `add(name: string, reference: Range or string, comment: string)`Fügt der Sammlung des angegebenen Bereichs einen neuen Namen hinzu.
* `addFormulaLocal(name: string, formula: string, comment: string)`Fügt einen neuen Namen zu der Auflistung des angegebenen Bereichs unter Verwendung des Gebietsschemas des Benutzers für die Formel hinzu.

### <a name="settings-api-in-in-excel-namespace"></a>Einstellungs-API im Excel-Namespace

Das [Setting](/javascript/api/excel/excel.setting)-Objekt stellt ein Schlüssel/Wert-Paar einer dauerhaften Einstellung für das Dokument dar. Es wurden nun Einstellungen im Zusammenhang mit APIs unter dem Excel-Namespace hinzugefügt. Dadurch werden keine neuen Funktionen angeboten, aber es wird erleichtert, in der zusagebasierten Batch-API-Syntax zu verbleiben, um die Abhängigkeit von allgemeinen Aufgaben im Zusammenhang mit der API für Excel zu reduzieren.

APIs enthalten `getItem()`, um den Setting-Eintrag über den Schlüssel abzurufen, und `add()`, um das angegebene Schlüssel:Wert-Einstellungspaar zu der Arbeitsmappe hinzuzufügen.

### <a name="others"></a>Sonstige

* Hinzufügen des Tabellenspaltennamens (in der früheren Version war nur das Lesen möglich).
* Hinzufügen einer Tabellenspalte am Ende der Tabelle (in der früheren Versionen war dies überall, außer an der letzten Stelle möglich).
* Gleichzeitiges Hinzufügen mehrerer Zeilen zu einer Tabelle (in der früheren Version war jeweils nur 1 Zeile möglich).
* `range.getColumnsAfter(count: number)` und `range.getColumnsBefore(count: number)` zum Abrufen einer bestimmten Anzahl von Spalten rechts/links vom aktuellen Bereichsobjekt.
* Element oder Null-Objektfunktion abrufen: Diese Funktionalität ermöglicht das Abrufen eines Objekts unter Verwendung eines Schlüssels. Wenn das Objekt nicht vorhanden ist, hat die isNullObject-Eigenschaft des zurückgegebenen Objekts den Wert „True“. Auf diese Weise können Entwickler überprüfen, ob ein Objekt vorhanden ist oder nicht, ohne dies über die Ausnahmebehandlung zu verarbeiten. Verfügbar im Arbeitsblatt, im benannten Element, in der Bindung, in Datenreihen usw.

    ```javascript
    worksheet.GetItemOrNullObject()
    ```

|Objekt| Neuigkeiten| Beschreibung|Anforderungssatz|
|:----|:----|:----|:----|
|[bindingCollection](/javascript/api/excel/excel.bindingcollection)|_Methode_ > getCount()|Ruft die Anzahl der Bindungen in der Sammlung ab.|1.4|
|[bindingCollection](/javascript/api/excel/excel.bindingcollection)|_Methode_ > GetItemOrNullObject(id: string)|Ruft ein Binding-Objekt anhand seiner ID ab. Wenn das Binding-Objekt nicht vorhanden ist, wird ein NULL-Objekt zurückgegeben.|1.4|
|[chartCollection](/javascript/api/excel/excel.chartcollection)|_Methode_ > getCount()|Gibt die Anzahl der Diagramme im Arbeitsblatt zurück.|1.4|
|[chartCollection](/javascript/api/excel/excel.chartcollection)|_Methode_ > GetItemOrNullObject(name: string)|Ruft ein Diagramm über seinen Namen ab. Wenn mehrere Diagramme mit demselben Namen vorhanden sind, wird das erste zurückgegeben.|1.4|
|[chartPointsCollection](/javascript/api/excel/excel.chartpointscollection)|_Methode_ > getCount()|Gibt die Anzahl der Diagrammpunkten in der Reihe zurück.|1.4|
|[chartSeriesCollection](/javascript/api/excel/excel.chartseriescollection)|_Methode_ > getCount()|Gibt die Anzahl der Datenreihen in der Sammlung zurück.|1.4|
|[namedItem](/javascript/api/excel/excel.nameditem)|_Eigenschaft_ > comment|Stellt den Kommentar dar, der mit diesem Namen verknüpft ist|1.4|
|[namedItem](/javascript/api/excel/excel.nameditem)|_Eigenschaft_ > scope|Gibt an, ob der Name im Bereich der Arbeitsmappe oder im Bereich eines bestimmten Arbeitsblatts liegt. Schreibgeschützt. Die folgenden Werte sind möglich: Equal, Greater, GreaterEqual, Less, LessEqual, NotEqual.|1.4|
|[namedItem](/javascript/api/excel/excel.nameditem)|_Beziehung_ > worksheet|Gibt das Arbeitsblatt zurück, auf dessen Bereich das benannte Element beschränkt ist. Löst einen Fehler aus, wenn die Elemente stattdessen auf den Bereich der Arbeitsmappe beschränkt sind. Schreibgeschützt.|1.4|
|[namedItem](/javascript/api/excel/excel.nameditem)|_Beziehung_ > worksheetOrNullObject|Gibt das Arbeitsblatt zurück, auf dessen Bereich das benannte Element beschränkt ist. Gibt ein NULL-Objekt zurück, wenn das Element stattdessen auf den Bereich der Arbeitsmappe beschränkt ist. Schreibgeschützt.|1.4|
|[namedItem](/javascript/api/excel/excel.nameditem)|_Methode_ > delete()|Löscht den angegebenen Namen.|1.4|
|[namedItem](/javascript/api/excel/excel.nameditem)|_Methode_ > getRangeOrNullObject()|Ruft das Bereichsobjekt ab, das mit dem Namen verknüpft ist. Gibt ein NULL-Objekt zurück, wenn der Typ des benannten Elements kein Bereich ist.|1.4|
|[namedItemCollection](/javascript/api/excel/excel.nameditemcollection)|_Methode_ > hinzufügen (Name: string, Referenz: Bereich oder eine Zeichenfolge, Kommentar: Zeichenfolge)|Fügt einen neuen Namen zu der Auflistung des angegebenen Bereichs hinzu.|1.4|
|[namedItemCollection](/javascript/api/excel/excel.nameditemcollection)|_Methode_ > AddFormulaLocal (Name: string, Formel: string, Kommentar: Zeichenfolge)|Fügt einen neuen Namen zu der Auflistung des angegebenen Bereichs unter Verwendung des Gebietsschemas des Benutzers für die Formel hinzu.|1.4|
|[namedItemCollection](/javascript/api/excel/excel.nameditemcollection)|_Methode_ > getCount()|Ruft die Anzahl von benannten Elementen in der Auflistung ab.|1.4|
|[namedItemCollection](/javascript/api/excel/excel.nameditemcollection)|_Methode_ > GetItemOrNullObject(name: string)|Ruft ein nameditem-Objekt über den Namen ab. Wenn das nameditem-Objekt nicht vorhanden ist, wird ein NULL-Objekt zurückgegeben.|1.4|
|[pivotTableCollection](/javascript/api/excel/excel.pivottablecollection)|_Methode_ > getCount()|Ruft die Anzahl von PivotTables in der Auflistung ab.|1.4|
|[pivotTableCollection](/javascript/api/excel/excel.pivottablecollection)|_Methode_ > GetItemOrNullObject(name: string)|Ruft eine PivotTable anhand des Namens ab. Wenn die PivotTable nicht vorhanden ist, wird ein NULL-Objekt zurückgegeben.|1.4|
|[range](/javascript/api/excel/excel.range)|_Methode_ > GetIntersectionOrNullObject (AnotherRange: Bereich oder Zeichenfolge)|Ruft das Bereichsobjekt ab, das die rechteckige Schnittmenge der angegebenen Bereiche darstellt. Wenn keine Schnittmenge gefunden wird, wird ein null-Objekt zurückgegeben.|1.4|
|[range](/javascript/api/excel/excel.range)|_Methode_ > GetUsedRangeOrNullObject(valuesOnly: bool)|Gibt den verwendeten Bereich des angegebenen Bereichsobjekts zurück. Wenn keine verwendeten Zellen innerhalb des Bereichs vorhanden sind, gibt diese Funktion ein NULL-Objekt zurück.|1.4|
|[rangeViewCollection](/javascript/api/excel/excel.rangeviewcollection)|_Methode_ > getCount()|Gibt die Anzahl von RangeView-Objekten in der Auflistung zurück.|1.4|
|[setting](/javascript/api/excel/excel.setting)|_Eigenschaft_ > key|Gibt den Schlüssel zurück, der die ID der Einstellung darstellt. Schreibgeschützt.|1.4|
|[setting](/javascript/api/excel/excel.setting)|_Eigenschaft_ > value|Stellt den Wert dar, der für diese Einstellung gespeichert ist.|1.4|
|[setting](/javascript/api/excel/excel.setting)|_Methode_ > delete()|Löscht die Einstellung.|1.4|
|[settingCollection](/javascript/api/excel/excel.settingcollection)|_Eigenschaft_ > items|Eine Sammlung von setting-Objekten. Schreibgeschützt.|1.4|
|[settingCollection](/javascript/api/excel/excel.settingcollection)|_Methode_ > hinzufügen (Schlüssel: string-Wert: (alle))|Legt die angegebene Einstellung fest oder fügt sie zur Arbeitsmappe hinzu.|1.4|
|[settingCollection](/javascript/api/excel/excel.settingcollection)|_Methode_ > getCount()|Ruft die Anzahl der Einstellungen in der Sammlung ab.|1.4|
|[settingCollection](/javascript/api/excel/excel.settingcollection)|_Methode_ > GetItem(key: string)|Ruft einen Setting-Eintrag über den Schlüssel ab.|1.4|
|[settingCollection](/javascript/api/excel/excel.settingcollection)|_Methode_ > GetItemOrNullObject(key: string)|Ruft einen Setting-Eintrag über den Schlüssel ab. Wenn das Setting-Objekt nicht vorhanden ist, wird ein NULL-Objekt zurückgegeben.|1.4|
|[settingsChangedEventArgs](/javascript/api/excel/excel.settingschangedeventargs)|_Beziehung_ > settings|Ruft das Setting-Objekt ab, das die Bindung darstellt, durch die das SettingsChanged-Ereignis ausgelöst wurde.|1.4|
|[tableCollection](/javascript/api/excel/excel.tablecollection)|_Methode_ > getCount()]|Ruft die Anzahl von Tabellen in der Auflistung ab.|1.4|
|[tableCollection](/javascript/api/excel/excel.tablecollection)|_Methode_ > GetItemOrNullObject (Schlüssel: Zahl oder eine Zeichenfolge)|Ruft eine Tabelle anhand des Namens oder der ID ab. Wenn die Tabelle nicht vorhanden ist, wird ein NULL-Objekt zurückgegeben.|1.4|
|[tableColumnCollection](/javascript/api/excel/excel.tablecolumncollection)|_Methode_ > getCount()|Ruft die Anzahl der Spalten in der Tabelle ab.|1.4|
|[tableColumnCollection](/javascript/api/excel/excel.tablecolumncollection)|_Methode_ > GetItemOrNullObject (Schlüssel: Zahl oder eine Zeichenfolge)|Ruft ein Spaltenobjekt nach Name oder ID ab. Wenn die Spalte nicht vorhanden ist, wird ein NULL-Objekt zurückgegeben.|1.4|
|[tableRowCollection](/javascript/api/excel/excel.tablerowcollection)|_Methode_ > getCount()|Ruft die Anzahl der Zeilen in der Tabelle ab.|1.4|
|[workbook](/javascript/api/excel/excel.workbook)|_Beziehung_ > settings|Stellt eine Auflistung der mit der Arbeitsmappe verknüpften Einstellungen dar. Schreibgeschützt.|1.4|
|[worksheet](/javascript/api/excel/excel.worksheet)|_Beziehung_ > names|Auflistung von Namen im Bereich des aktuellen Arbeitsblatts. Schreibgeschützt.|1.4|
|[worksheet](/javascript/api/excel/excel.worksheet)|_Methode_ > GetUsedRangeOrNullObject(valuesOnly: bool)|Der verwendete Bereich ist der kleinste Bereich, der mindestens eine der Zellen umfasst, die einen Wert enthalten oder denen eine Formatierung zugewiesen wurde. Wenn das gesamte Arbeitsblatt leer ist, gibt diese Funktion ein NULL-Objekt zurück.|1.4|
|[worksheetCollection](/javascript/api/excel/excel.worksheetcollection)|_Methode_ > GetCount(visibleOnly: bool)|Ruft die Anzahl von Arbeitsblättern in der Auflistung ab.|1.4|
|[worksheetCollection](/javascript/api/excel/excel.worksheetcollection)|_Methode_ > GetItemOrNullObject(key: string)|Ruft das Arbeitsblattobjekt mithilfe des Namens oder der ID ab. Wenn das Arbeitsblatt nicht vorhanden ist, wird ein NULL-Objekt zurückgegeben.|1.4|

## <a name="whats-new-in-excel-javascript-api-13"></a>Neuigkeiten in der Excel-JavaScript-API 1.3

Im Folgenden werden die neuen Elemente der Excel-JavaScript-APIs in Anforderungssatz 1.3 aufgeführt.

|Objekt| Neuerungen| Beschreibung|Anforderungssatz|
|:----|:----|:----|:----|
|[binding](/javascript/api/excel/excel.binding)|_Methode_ > delete()|Löscht die Bindung.|1.3|
|[bindingCollection](/javascript/api/excel/excel.bindingcollection)|_Methode_ > hinzufügen (Bereich: Bereich oder die Zeichenfolge, die "BindingType": String, -Id: Zeichenfolge)|Fügt eine neue Bindung an einen bestimmten Bereich hinzu.|1.3|
|[bindingCollection](/javascript/api/excel/excel.bindingcollection)|_Methode_ > AddFromNamedItem (Name: String, "BindingType": String, -Id: Zeichenfolge)|Fügt eine neue Bindung auf Grundlage eines benannten Elements in der Arbeitsmappe hinzu.|1.3|
|[bindingCollection](/javascript/api/excel/excel.bindingcollection)|_Methode_ > AddFromSelection ("BindingType": String, -Id: Zeichenfolge)|Fügt eine neue Bindung basierend auf der aktuellen Auswahl hinzu.|1.3|
|[bindingCollection](/javascript/api/excel/excel.bindingcollection)|_Methode_ > GetItemOrNull(id: string)|Ruft ein binding-Objekt anhand seiner ID ab. Wenn das binding-Objekt nicht existiert, hat die isNull-Eigenschaft des Rückgabeobjekts den Wert True.|1.3|
|[chartCollection](/javascript/api/excel/excel.chartcollection)|_Methode_ > GetItemOrNull(name: string)|Ruft ein Diagramm über seinen Namen ab. Wenn mehrere Diagramme mit demselben Namen vorhanden sind, wird das erste zurückgegeben.|1.3|
|[namedItemCollection](/javascript/api/excel/excel.nameditemcollection)|_Methode_ > GetItemOrNull(name: string)|Ruft ein nameditem-Objekt über den Namen ab. Wenn das nameditem-Objekt nicht existiert, hat die isNull-Eigenschaft des zurückgegebenen Objekts den Wert True.|1.3|
|[pivotTable](/javascript/api/excel/excel.pivottable)|_Eigenschaft_ > name|Der Name der PivotTable.|1.3|
|[pivotTable](/javascript/api/excel/excel.pivottable)|_Beziehung_ > worksheet|Das Arbeitsblatt, das die aktuelle PivotTable enthält. Schreibgeschützt.|1.3|
|[pivotTable](/javascript/api/excel/excel.pivottable)|_Methode_ > refresh()|Aktualisiert die PivotTable.|1.3|
|[pivotTableCollection](/javascript/api/excel/excel.pivottablecollection)|_Eigenschaft_ > items|Eine Sammlung von pivotTable-Objekten. Schreibgeschützt.|1.3|
|[pivotTableCollection](/javascript/api/excel/excel.pivottablecollection)|_Methode_ > GetItem(name: string)|Ruft eine PivotTable anhand des Namens ab.|1.3|
|[pivotTableCollection](/javascript/api/excel/excel.pivottablecollection)|_Methode_ > GetItemOrNull(name: string)|Ruft eine PivotTable anhand des Namens ab. Wenn die PivotTable nicht existiert, hat die isNull-Eigenschaft des Rückgabeobjekts den Wert True.|1.3|
|[range](/javascript/api/excel/excel.range)|_Methode_ > GetIntersectionOrNull (AnotherRange: Bereich oder Zeichenfolge)|Ruft das Bereichsobjekt ab, das die rechteckige Schnittmenge der angegebenen Bereiche darstellt. Wenn keine Schnittmenge gefunden wird, wird ein null-Objekt zurückgegeben.|1.3|
|[range](/javascript/api/excel/excel.range)|_Methode_ > getVisibleView()|Stellt die sichtbaren Zeilen des aktuellen Bereichs dar.|1.3|
|[rangeView](/javascript/api/excel/excel.rangeview)|_Eigenschaft_ > cellAddresses|Stellt die Zelladressen des RangeView dar. Schreibgeschützt.|1.3|
|[rangeView](/javascript/api/excel/excel.rangeview)|_Eigenschaft_ > columnCount|Gibt die Anzahl der sichtbaren Spalten zurück. Schreibgeschützt.|1.3|
|[rangeView](/javascript/api/excel/excel.rangeview)|_Eigenschaft_ > formulas|Stellt die Formel in der A1-Schreibweise dar.|1.3|
|[rangeView](/javascript/api/excel/excel.rangeview)|_Eigenschaft_ > formulasLocal|Stellt die Formel in der A1-Schreibweise, Sprache des Benutzers und im Gebietsschema der Zahlenformatierung dar.  Beispielsweise würde die englische Formel "=SUM(A1, introduced in 1.5)" in Deutsch "=SUMME(A1; 1,5) " werden.|1.3|
|[rangeView](/javascript/api/excel/excel.rangeview)|_Eigenschaft_ > formulasR1C1|Stellt die Formel in der R1C1-Schreibweise dar.|1.3|
|[rangeView](/javascript/api/excel/excel.rangeview)|_Eigenschaft_ > index|Gibt einen Wert zurück, der den Index des RangeView darstellt. Schreibgeschützt.|1.3|
|[rangeView](/javascript/api/excel/excel.rangeview)|_Eigenschaft_ > numberFormat|Stellt den Excel-Zahlenformatcode für die angegebene Zelle dar.|1.3|
|[rangeView](/javascript/api/excel/excel.rangeview)|_Eigenschaft_ > rowCount|Gibt die Anzahl der sichtbaren Zeilen zurück. Schreibgeschützt.|1.3|
|[rangeView](/javascript/api/excel/excel.rangeview)|_Eigenschaft_ > text|Textwerte des angegebenen Bereichs. Der Textwert hängt nicht von der Zellenbreite ab. Die Ersetzung des #-Zeichens, die in der Excel-Benutzeroberfläche passiert, wirkt sich nicht auf den von der API zurückgegebenen Textwert aus. Schreibgeschützt.|1.3|
|[rangeView](/javascript/api/excel/excel.rangeview)|_Eigenschaft_ > valueTypes|Stellt den Datentyp in jeder Zelle dar. Schreibgeschützt. Die folgenden Werte sind möglich: Unbekannt, leer, Zeichenfolge, Ganzzahl, Doppelwort, Boolesch, Fehler.|1.3|
|[rangeView](/javascript/api/excel/excel.rangeview)|_Eigenschaft_ > values|Stellt die Rohwerte der angegebenen Bereichsansicht dar. Die zurückgegebenen Daten können vom Typ Zeichenfolge, Zahl oder ein boolescher Wert sein. Zellen, die einen Fehler enthalten, geben die Fehlerzeichenfolge zurück.|1.3|
|[rangeView](/javascript/api/excel/excel.rangeview)|_Beziehung_ > rows|Stellt eine Sammlung der mit dem Bereich verknüpften Bereichsansichten dar. Schreibgeschützt.|1.3|
|[rangeView](/javascript/api/excel/excel.rangeview)|_Methode_ > getRange()|Ruft den übergeordneten Bereich zum aktuellen RangeView ab.|1.3|
|[rangeViewCollection](/javascript/api/excel/excel.rangeviewcollection)|_Eigenschaft_ > items|Eine Sammlung von rangeView-Objekten. Schreibgeschützt.|1.3|
|[rangeViewCollection](/javascript/api/excel/excel.rangeviewcollection)|_Methode_ > GetItemAt(index: number)|Ruft eine RangeView-Zeile über der Index ab. Nullindiziert.|1.3|
|[setting](/javascript/api/excel/excel.setting)|_Eigenschaft_ > key|Gibt den Schlüssel zurück, der die ID der Einstellung darstellt. Schreibgeschützt.|1.3|
|[setting](/javascript/api/excel/excel.setting)|_Methode_ > delete()|Löscht die Einstellung.|1.3|
|[settingCollection](/javascript/api/excel/excel.settingcollection)|_Eigenschaft_ > items|Eine Sammlung von setting-Objekten. Schreibgeschützt.|1.3|
|[settingCollection](/javascript/api/excel/excel.settingcollection)|_Methode_ > GetItem(key: string)|Ruft einen Setting-Eintrag über den Schlüssel ab.|1.3|
|[settingCollection](/javascript/api/excel/excel.settingcollection)|_Methode_ > GetItemOrNull(key: string)|Ruft einen Setting-Eintrag über den Schlüssel ab. Wenn das Setting-Objekt nicht existiert, hat die isNull-Eigenschaft des zurückgegebenen Objekts den Wert True.|1.3|
|[settingCollection](/javascript/api/excel/excel.settingcollection)|_Methode_ > festgelegt (Schlüssel: string-Wert: Zeichenfolge)|Legt die angegebene Einstellung fest oder fügt sie zur Arbeitsmappe hinzu.|1.3|
|[settingsChangedEventArgs](/javascript/api/excel/excel.settingschangedeventargs)|_Beziehung_ > settingCollection|Ruft das Setting-Objekt ab, das die Bindung darstellt, durch die das SettingsChanged-Ereignis ausgelöst wurde.|1.3|
|[table](/javascript/api/excel/excel.table)|_Eigenschaft_ > highlightFirstColumn|Gibt an, ob die erste Spalte spezielle Formatierung enthält.|1.3|
|[table](/javascript/api/excel/excel.table)|_Property_ > highlightLastColumn|Gibt an, ob die letzte Spalte spezielle Formatierung enthält.|1.3|
|[table](/javascript/api/excel/excel.table)|_Eigenschaft_ > showBandedColumns|Gibt an, ob die Spalten gebänderte Formatierung aufweisen, wobei ungerade Spalten anders hervorgehoben werden als gerade, um die Tabelle leichter lesbar zu machen.|1.3|
|[table](/javascript/api/excel/excel.table)|_Eigenschaft_ > showBandedRows|Gibt an, ob die Zeilen gebänderte Formatierung aufweisen, wobei ungerade Zeilen anders hervorgehoben werden als gerade, um die Tabelle leichter lesbar zu machen.|1.3|
|[table](/javascript/api/excel/excel.table)|_Eigenschaft_ > showFilterButton|Gibt an, ob die Filterschaltflächen am oberen Rand jedes Spaltenheaders sichtbar sind. Diese Einstellung ist nur zulässig, wenn die Tabelle eine Kopfzeile enthält.|1.3|
|[tableCollection](/javascript/api/excel/excel.tablecollection)|_Methode_ > GetItemOrNull (Schlüssel: Zahl oder eine Zeichenfolge)|Ruft eine Tabelle anhand des Namens oder der ID ab. Wenn die Tabelle nicht existiert, hat die isNull-Eigenschaft des Rückgabeobjekts den Wert True.|1.3|
|[tableColumnCollection](/javascript/api/excel/excel.tablecolumncollection)|_Methode_ > GetItemOrNull (Schlüssel: Zahl oder eine Zeichenfolge)|Ruft ein Spaltenobjekt nach Name oder ID ab. Wenn die Spalte nicht existiert, hat die isNull-Eigenschaft des zurückgegebenen Objekts den Wert True.|1.3|
|[workbook](/javascript/api/excel/excel.workbook)|_Beziehung_ > pivotTables|Stellt eine Auflistung der mit der Arbeitsmappe verknüpften PivotTables dar. Schreibgeschützt.|1.3|
|[workbook](/javascript/api/excel/excel.workbook)|_Beziehung_ > settings|Stellt eine Auflistung der mit der Arbeitsmappe verknüpften Einstellungen dar. Schreibgeschützt.|1.3|
|[worksheet](/javascript/api/excel/excel.worksheet)|_Beziehung_ > pivotTables|Die Sammlung von PivotTables, die Teil des Arbeitsblatts sind. Schreibgeschützt.|1.3|

## <a name="whats-new-in-excel-javascript-api-12"></a>Neuigkeiten in der Excel-JavaScript-API 1.2

Im Folgenden werden die neuen Elemente der Excel-JavaScript-APIs in Anforderungssatz 1.2 aufgeführt.

|Objekt| Neuerungen| Beschreibung|Anforderungssatz|
|:----|:----|:----|:----|
|[chart](/javascript/api/excel/excel.chart)|_Eigenschaft_ > id|Ruft ein Diagramm anhand seiner Position in der Sammlung ab. Schreibgeschützt.|1.2|
|[chart](/javascript/api/excel/excel.chart)|_Beziehung_ > worksheet|Das Arbeitsblatt, das das aktuelle Diagramm enthält. Schreibgeschützt.|1.2|
|[chart](/javascript/api/excel/excel.chart)|_Methode_ > GetImage (Höhe: number, Breite: Number, FittingMode: Zeichenfolge)|Rendert das Diagramm als base64-codiertes Bild durch Skalierung, um es an die angegebenen Maße anzupassen.|1.2|
|[filter](/javascript/api/excel/excel.filter)|_Beziehung_ > criteria|Der aktuell angewendete Filter in der angegebenen Spalte. Schreibgeschützt.|1.2|
|[filter](/javascript/api/excel/excel.filter)|_Methode_ > Apply(criteria: FilterCriteria)|Wendet die angegebenen Filterkriterien in der angegebenen Spalte an.|1.2|
|[filter](/javascript/api/excel/excel.filter)|_Methode_ > ApplyBottomItemsFilter(count: number)|Wendet den Filter "Bottom Item" auf die Spalte für die angegebene Anzahl von Elementen an.|1.2|
|[filter](/javascript/api/excel/excel.filter)|_Methode_ > ApplyBottomPercentFilter(percent: number)]|Wendet den Filter "Bottom Percent" auf die Spalte für den angegebenen Prozentsatz von Elementen an.|1.2|
|[filter](/javascript/api/excel/excel.filter)|_Methode_ > ApplyCellColorFilter(color: string)|Wendet den Filter "Cell Color" auf die Spalte für die angegebene Farbe an.|1.2|
|[filter](/javascript/api/excel/excel.filter)|_Methode_ > ApplyCustomFilter (criteria1: String, criteria2: String, Operator: Zeichenfolge)|Wendet den Filter "Icon" auf die Spalte für die angegebenen Kriterienzeichenfolgen an.|1.2|
|[filter](/javascript/api/excel/excel.filter)|_Methode_ > ApplyDynamicFilter(criteria: string)|Wendet den Filter "Dynamic" auf die Spalte an.|1.2|
|[filter](/javascript/api/excel/excel.filter)|_Methode_ > ApplyFontColorFilter(color: string)|Wendet den Filter "Font Color" auf die Spalte für die angegebene Farbe an.|1.2|
|[filter](/javascript/api/excel/excel.filter)|_Methode_ > ApplyIconFilter(icon: Icon)|Wendet den Filter "Icon" auf die Spalte für das angegebene Symbol an.|1.2|
|[filter](/javascript/api/excel/excel.filter)|_Methode_ > ApplyTopItemsFilter(count: number)|Wendet den Filter "Top Item" auf die Spalte für die angegebene Anzahl von Elementen an.|1.2|
|[filter](/javascript/api/excel/excel.filter)|_Methode_ > ApplyTopPercentFilter(percent: number)|Wendet den Filter "Top Percent" auf die Spalte für den angegebenen Prozentsatz von Elementen an.|1.2|
|[filter](/javascript/api/excel/excel.filter)|_Methode_ > ApplyValuesFilter (Werte: ())|Wendet den Filter "Values" auf die Spalte für die angegebenen Werte an.|1.2|
|[filter](/javascript/api/excel/excel.filter)|_Methode_ > Löschvorgang()|Deaktiviert den Filter für die angegebene Spalte.|1.2|
|[filterCriteria](/javascript/api/excel/excel.filtercriteria)|_Eigenschaft_ > color|Die HTML-Farbzeichenfolge, die zum Filtern von Zellen verwendet wird. Wird mit "cellColor"- und "fontColor"-Filterung verwendet.|1.2|
|[filterCriteria](/javascript/api/excel/excel.filtercriteria)|_Eigenschaft_ > criterion1|Das erste verwendete Kriterium zum Filtern von Daten. Wird als Operator bei "custom"-Filtern verwendet.|1.2|
|[filterCriteria](/javascript/api/excel/excel.filtercriteria)|_Eigenschaft_ > criterion2|Das zweite verwendete Kriterium zum Filtern von Daten. Wird nur als Operator bei "custom"-Filtern verwendet.|1.2|
|[filterCriteria](/javascript/api/excel/excel.filtercriteria)|_Eigenschaft_ > dynamicCriteria|Die dynamischen Kriterien aus dem Satz "Excel.DynamicFilterCriteria", die auf diese Spalte angewendet werden. Wird mit "dynamic"-Filtern verwendet. Die folgenden Werte sind möglich: Unknown, AboveAverage, AllDatesInPeriodApril, AllDatesInPeriodAugust, AllDatesInPeriodDecember, AllDatesInPeriodFebruray, AllDatesInPeriodJanuary, AllDatesInPeriodJuly, AllDatesInPeriodJune, AllDatesInPeriodMarch, AllDatesInPeriodMay, AllDatesInPeriodNovember, AllDatesInPeriodOctober, AllDatesInPeriodQuarter1, AllDatesInPeriodQuarter2, AllDatesInPeriodQuarter3, AllDatesInPeriodQuarter4, AllDatesInPeriodSeptember, BelowAverage, LastMonth, LastQuarter, LastWeek, LastYear, NextMonth, NextQuarter, NextWeek, NextYear, ThisMonth, ThisQuarter, ThisWeek, ThisYear, Today, Tomorrow, YearToDate, Yesterday.|1.2|
|[filterCriteria](/javascript/api/excel/excel.filtercriteria)|_Eigenschaft_ > filterOn|Die Eigenschaft, die vom Filter verwendet wird, um zu bestimmen, ob die Werte sichtbar bleiben sollen. Die folgenden Werte sind möglich: BottomItems, BottomPercent, CellColor, Dynamic, FontColor, Values, TopItems, TopPercent, Icon, Custom.|1.2|
|[filterCriteria](/javascript/api/excel/excel.filtercriteria)|_Eigenschaft_ > operator|Der Operator, mit dem Kriterium 1 und 2 bei Verwendung von "Custom"-Filtern kombiniert werden. Die folgenden Werte sind möglich: And, Or.|1.2|
|[filterCriteria](/javascript/api/excel/excel.filtercriteria)|_Eigenschaft_ > values|Der Satz von Werten, der als Teil von "values"-Filtern verwendet werden soll.|1.2|
|[filterCriteria](/javascript/api/excel/excel.filtercriteria)|_Beziehung_ > icon|Das zum Filtern von Zellen verwendete Symbol. Wird mit "icon"-Filtern verwendet.|1.2|
|[filterDatetime](/javascript/api/excel/excel.filterdatetime)|_Eigenschaft_ > date|Das Datum im ISO8601-Format zum Filtern von Daten.|1.2|
|[filterDatetime](/javascript/api/excel/excel.filterdatetime)|_Eigenschaft_ > specificity|Genauigkeit des Datums zum Beibehalten von Daten. Wenn z. B. das Datum 2005-04-02 ist und die Spezifität auf „Monat“ festgelegt ist, werden beim Filtervorgang alle Zeilen mit einem Datum im Monat April 2009 beibehalten. Die folgenden Werte sind möglich: Jahr, Monat, Tag, Stunde, Minute und Sekunde.|1.2|
|[formatProtection](/javascript/api/excel/excel.formatprotection)|_Eigenschaft_ > formulaHidden|Zeigt an, ob Excel die Formel für die Zellen im Bereich ausblendet. Ein Nullwert gibt an, dass der gesamte Bereich keine einheitliche Einstellung zur Formelausblendung hat.|1.2|
|[formatProtection](/javascript/api/excel/excel.formatprotection)|_Eigenschaft_ > locked|Gibt an, ob Excel die Zellen im Objekt sperrt. Ein Nullwert gibt an, dass der gesamte Bereich keine einheitliche Einstellung zum Sperren hat.|1.2|
|[icon](/javascript/api/excel/excel.icon)|_Eigenschaft_ > index|Stellt den Index des Symbols im angegebenen Satz dar.|1.2|
|[icon](/javascript/api/excel/excel.icon)|_Eigenschaft_ > set|Stellt den Satz dar, zu dem das Symbol gehört. Die folgenden Werte sind möglich: Invalid, ThreeArrows, ThreeArrowsGray, ThreeFlags, ThreeTrafficLights1, ThreeTrafficLights2, ThreeSigns, ThreeSymbols, ThreeSymbols2, FourArrows, FourArrowsGray, FourRedToBlack, FourRating, FourTrafficLights, FiveArrows, FiveArrowsGray, FiveRating, FiveQuarters, ThreeStars, ThreeTriangles, FiveBoxes.|1.2|
|[range](/javascript/api/excel/excel.range)|_Eigenschaft_ > columnHidden|Stellt dar, ob alle Spalten des aktuellen Bereichs ausgeblendet sind.|1.2|
|[range](/javascript/api/excel/excel.range)|_Eigenschaft_ > formulasR1C1|Stellt die Formel in der R1C1-Schreibweise dar.|1.2|
|[range](/javascript/api/excel/excel.range)|_Eigenschaft_ > hidden|Stellt dar, ob alle Zellen des aktuellen Bereichs ausgeblendet sind. Schreibgeschützt.|1.2|
|[range](/javascript/api/excel/excel.range)|_Eigenschaft_ > rowHidden|Stellt dar, ob alle Zeilen des aktuellen Bereichs ausgeblendet sind.|1.2|
|[range](/javascript/api/excel/excel.range)|_Beziehung_ > sort|Stellt die Bereichssortierung des aktuellen Bereichs dar. Schreibgeschützt.|1.2|
|[range](/javascript/api/excel/excel.range)|_Methode_ > Merge(across: bool)|Führt die Zellen des Bereichs in eine Region im Arbeitsblatt zusammen.|1.2|
|[range](/javascript/api/excel/excel.range)|_Methode_ > unmerge()|Hebt den Zellverbund des Bereichs in einzelne Zellen auf.|1.2|
|[rangeFormat](/javascript/api/excel/excel.rangeformat)|_Eigenschaft_ > columnWidth|Ruft die Breite aller Spalten innerhalb des Bereichs ab oder legt diese fest. Wenn die Breite der Spalten nicht gleichmäßig ist, wird Null zurückgegeben.|1.2|
|[rangeFormat](/javascript/api/excel/excel.rangeformat)|_Eigenschaft_ > rowHeight|Ruft die Höhe aller Zeilen des Bereichs ab oder legt diese fest. Wenn die Höhe der Zeilen nicht gleichmäßig ist, wird Null zurückgegeben.|1.2|
|[rangeFormat](/javascript/api/excel/excel.rangeformat)|_Beziehung_ > protection|Gibt das Formatschutz-Objekt für einen Bereich zurück. Schreibgeschützt.|1.2|
|[rangeFormat](/javascript/api/excel/excel.rangeformat)|_Methode_ > autofitColumns()|Ändert die Breite der Spalten des aktuellen Bereichs, um basierend auf den aktuellen Daten in den Spalten die optimale Breite zu erzielen.|1.2|
|[rangeFormat](/javascript/api/excel/excel.rangeformat)|_Methode_ > autofitRows()|Ändert die Höhe der Zeilen des aktuellen Bereichs, um basierend auf den aktuellen Daten in den Zeilen die optimale Höhe zu erzielen.|1.2|
|[rangeReference](/javascript/api/excel/excel.rangereference)|_Eigenschaft_ > address|Stellt die sichtbaren Zeilen des aktuellen Bereichs dar.|1.2|
|[rangeSort](/javascript/api/excel/excel.rangesort)|_Methode_ > anwenden (Felder: SortField, MatchCase: Bool HasHeaders: Bool, Ausrichtung: string,-Methode: Zeichenfolge)|Führt einen Sortiervorgang aus.|1.2|
|[sortField](/javascript/api/excel/excel.sortfield)|_Eigenschaft_ > ascending|Stellt dar, ob die Sortierung in aufsteigender Reihenfolge ausgeführt wird.|1.2|
|[sortField](/javascript/api/excel/excel.sortfield)|_Eigenschaft_ > color|Stellt die Farbe dar, die das Ziel der Bedingung ist, wenn die Sortierung für die Schrift- oder Zellenfarbe gilt.|1.2|
|[sortField](/javascript/api/excel/excel.sortfield)|_Eigenschaft_ > dataOption|Stellt weitere Sortieroptionen für dieses Feld dar. Die folgenden Werte sind möglich: Normal, TextAsNumber.|1.2|
|[sortField](/javascript/api/excel/excel.sortfield)|_Eigenschaft_ > key|Stellt die Spalte (oder Zeile, je nach Sortierausrichtung) dar, für die die Bedingung gilt. Wird als Offset von der ersten Spalte (oder Zeile) dargestellt.|1.2|
|[sortField](/javascript/api/excel/excel.sortfield)|_Eigenschaft_ > sortOn|Stellt den Typ der Sortierung dieser Bedingung dar. Die folgenden Werte sind möglich: Value, CellColor, FontColor, Icon.|1.2|
|[sortField](/javascript/api/excel/excel.sortfield)|_Beziehung_ > icon|Stellt das Symbol dar, auf das sich die Bedingung bezieht, wenn die Sortierung für das Symbol der Zelle gilt.|1.2|
|[table](/javascript/api/excel/excel.table)|_Beziehung_ > sort|Stellt die Sortierung für die Tabelle dar. Schreibgeschützt.|1.2|
|[table](/javascript/api/excel/excel.table)|_Beziehung_ > worksheet|Das Arbeitsblatt, das die aktuelle Tabelle enthält. Schreibgeschützt.|1.2|
|[table](/javascript/api/excel/excel.table)|_Methode_ > clearFilters()|Löscht alle Filter, die derzeit in der Tabelle verwendet werden.|1.2|
|[table](/javascript/api/excel/excel.table)|_Methode_ > convertToRange()|Wandelt die Tabelle in einen normalen Bereich von Zellen um. Alle Daten werden beibehalten.|1.2|
|[table](/javascript/api/excel/excel.table)|_Methode_ > reapplyFilters()|Wendet alle Filter erneut an, die derzeit in der Tabelle vorhanden sind.|1.2|
|[tableColumn](/javascript/api/excel/excel.tablecolumn)|_Beziehung_ > filter|Ruft den auf die Salte angewendeten Filter ab. Schreibgeschützt.|1.2|
|[tableSort](/javascript/api/excel/excel.tablesort)|_Eigenschaft_ > matchCase|Stellt dar, ob die Groß-/Kleinschreibung den letzten Sortiervorgang der Tabelle beeinflusst hat. Schreibgeschützt.|1.2|
|[tableSort](/javascript/api/excel/excel.tablesort)|_Eigenschaft_ > method|Stellt die chinesische Zeichensortiermethode dar, mit der die Tabelle zuletzt sortiert wurde. Schreibgeschützt. Die folgenden Werte sind möglich: PinYin, StrokeCount.|1.2|
|[tableSort](/javascript/api/excel/excel.tablesort)|_Beziehung_ > fields|Stellt die aktuellen Bedingungen dar, die zuletzt zum Sortieren der Tabelle verwendet wurden. Schreibgeschützt.|1.2|
|[tableSort](/javascript/api/excel/excel.tablesort)|_Methode_ > anwenden (Felder: SortField, MatchCase: Bool Methode: Zeichenfolge)|Führt einen Sortiervorgang aus.|1.2|
|[tableSort](/javascript/api/excel/excel.tablesort)|_Methode_ > Löschvorgang()|Löscht die Sortierung, die derzeit in der Tabelle enthalten ist. Dies ändert nicht die Sortierung der Tabelle, löscht jedoch den Zustand der Kopfzeilen-Schaltflächen.|1.2|
|[tableSort](/javascript/api/excel/excel.tablesort)|_Methode_ > reapply()|Wendet die aktuellen Sortierparameter erneut auf die Tabelle an.|1.2|
|[workbook](/javascript/api/excel/excel.workbook)|_Beziehung_ > functions|Stellt die Instanz der Excel-Anwendung dar, die diese Arbeitsmappe enthält. Schreibgeschützt.|1.2|
|[worksheet](/javascript/api/excel/excel.worksheet)|_Beziehung_ > protection|Gibt das Arbeitsblattschutz-Objekt für ein Arbeitsblatt zurück. Schreibgeschützt.|1.2|
|[worksheetProtection](/javascript/api/excel/excel.worksheetprotection)|_Eigenschaft_ > protected|Zeigt an, ob das Arbeitsblatt geschützt ist. Schreibgeschützt. Schreibgeschützt.|1.2|
|[worksheetProtection](/javascript/api/excel/excel.worksheetprotection)|_Beziehung_ > options|Optionen für den Arbeitsblattschutz. Schreibgeschützt.|1.2|
|[worksheetProtection](/javascript/api/excel/excel.worksheetprotection)|_Methode_ > Protect(options: WorksheetProtectionOptions)|Schützt ein Arbeitsblatt. Liefert einen Fehler, wenn das Arbeitsblatt geschützt ist.|1.2|
|[worksheetProtection](/javascript/api/excel/excel.worksheetprotection)|_Methode_ > unprotect()|Hebt den Schutz eines Arbeitsblatts auf.|1.2|
|[worksheetProtectionOptions](/javascript/api/excel/excel.worksheetprotectionoptions)|_Eigenschaft_ > allowAutoFilter|Stellt die Arbeitsblatt-Schutzoption zum Zulassen der Verwendung der automatischen Filterfunktion dar.|1.2|
|[worksheetProtectionOptions](/javascript/api/excel/excel.worksheetprotectionoptions)|_Eigenschaft_ > allowDeleteColumns|Stellt die Arbeitsblatt-Schutzoption zum Zulassen des Löschens von Spalten dar.|1.2|
|[worksheetProtectionOptions](/javascript/api/excel/excel.worksheetprotectionoptions)|_Eigenschaft_ > allowDeleteRows|Stellt die Arbeitsblatt-Schutzoption zum Zulassen des Löschens von Zeilen dar.|1.2|
|[worksheetProtectionOptions](/javascript/api/excel/excel.worksheetprotectionoptions)|_Eigenschaft_ > allowFormatCells|Stellt die Arbeitsblatt-Schutzoption zum Zulassen des Formatierens von Zellen dar.|1.2|
|[worksheetProtectionOptions](/javascript/api/excel/excel.worksheetprotectionoptions)|_Eigenschaft_ > allowFormatColumns|Stellt die Arbeitsblatt-Schutzoption zum Zulassen des Formatierens von Spalten dar.|1.2|
|[worksheetProtectionOptions](/javascript/api/excel/excel.worksheetprotectionoptions)|_Eigenschaft_ > allowFormatRows|Stellt die Arbeitsblatt-Schutzoption zum Zulassen des Formatierens von Zeilen dar.|1.2|
|[worksheetProtectionOptions](/javascript/api/excel/excel.worksheetprotectionoptions)|_Eigenschaft_ > allowInsertColumns|Stellt die Arbeitsblatt-Schutzoption zum Zulassen des Einfügens von Spalten dar.|1.2|
|[worksheetProtectionOptions](/javascript/api/excel/excel.worksheetprotectionoptions)|_Eigenschaft_ > allowInsertHyperlinks|Stellt die Arbeitsblatt-Schutzoption zum Zulassen des Einfügens von Links dar.|1.2|
|[worksheetProtectionOptions](/javascript/api/excel/excel.worksheetprotectionoptions)|_Eigenschaft_ > allowInsertRows|Stellt die Arbeitsblatt-Schutzoption zum Zulassen des Einfügens von Zeilen dar.|1.2|
|[worksheetProtectionOptions](/javascript/api/excel/excel.worksheetprotectionoptions)|_Eigenschaft_ > allowPivotTables|Stellt die Arbeitsblatt-Schutzoption zum Zulassen der Verwendung des PivotTable-Features dar.|1.2|
|[worksheetProtectionOptions](/javascript/api/excel/excel.worksheetprotectionoptions)|_Eigenschaft_ > allowSort|Stellt die Arbeitsblatt-Schutzoption zum Zulassen der Verwendung der Sortierfunktion dar.|1.2|

## <a name="excel-javascript-api-11"></a>Excel-JavaScript-API 1.1

Excel JavaScript-API 1.1 ist die erste Version der API. Weitere Informationen zur API finden Sie unter den Referenzthemen zur [JavaScript-API für Excel](/javascript/api/excel) .

## <a name="see-also"></a>Siehe auch

- [Office-Versionen und Anforderungssätze](https://docs.microsoft.com/office/dev/add-ins/develop/office-versions-and-requirement-sets)
- 
  [Angeben von Office-Hosts und API-Anforderungen](https://docs.microsoft.com/office/dev/add-ins/develop/specify-office-hosts-and-api-requirements)
- [XML-Manifest für Office-Add-Ins](https://docs.microsoft.com/office/dev/add-ins/develop/add-in-manifests)
