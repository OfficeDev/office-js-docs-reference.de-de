| Klasse | Felder | Beschreibung |
|:---|:---|:---|
|[Anwendung](/javascript/api/excel/excel.application)|[suspendApiCalculationUntilNextSync()](/javascript/api/excel/excel.application#suspendapicalculationuntilnextsync--)|Die Berechnung wird angehalten, bis die nächste `context.sync()` aufgerufen wird.|
|[CellValueConditionalFormat](/javascript/api/excel/excel.cellvalueconditionalformat)|[format](/javascript/api/excel/excel.cellvalueconditionalformat#format)|Gibt ein Formatobjekt zurück, das die Eigenschaften Schriftart, Füllung, Rahmen und andere bedingte Formate kapselt.|
||[rule](/javascript/api/excel/excel.cellvalueconditionalformat#rule)|Gibt das Regelobjekt für dieses bedingte Format an.|
|[ColorScaleConditionalFormat](/javascript/api/excel/excel.colorscaleconditionalformat)|[criteria](/javascript/api/excel/excel.colorscaleconditionalformat#criteria)|Die Kriterien der Farbskala.|
||[threeColorScale](/javascript/api/excel/excel.colorscaleconditionalformat#threecolorscale)|Wenn `true` , hat die Farbskala drei Punkte (Minimum, Midpoint, Maximum), andernfalls zwei (Minimum, Maximum).|
|[ConditionalCellValueRule](/javascript/api/excel/excel.conditionalcellvaluerule)|[formula1](/javascript/api/excel/excel.conditionalcellvaluerule#formula1)|Die Formel, falls erforderlich, für die die Regel für das bedingte Format ausgewertet werden soll.|
||[formula2](/javascript/api/excel/excel.conditionalcellvaluerule#formula2)|Die Formel, falls erforderlich, für die die Regel für das bedingte Format ausgewertet werden soll.|
||[operator](/javascript/api/excel/excel.conditionalcellvaluerule#operator)|Der Operator des bedingten Zellwertformats.|
|[ConditionalColorScaleCriteria](/javascript/api/excel/excel.conditionalcolorscalecriteria)|[maximum](/javascript/api/excel/excel.conditionalcolorscalecriteria#maximum)|Der maximale Punkt des Farbskalenkriteriums.|
||[Midpoint](/javascript/api/excel/excel.conditionalcolorscalecriteria#midpoint)|Der Mittelpunkt des Farbskalenkriteriums, wenn es sich bei der Farbskala um eine 3-Farbskala handelt.|
||[minimum](/javascript/api/excel/excel.conditionalcolorscalecriteria#minimum)|Der Mindestpunkt des Farbskalakriteriums.|
|[ConditionalColorScaleCriterion](/javascript/api/excel/excel.conditionalcolorscalecriterion)|[color](/javascript/api/excel/excel.conditionalcolorscalecriterion#color)|HTML-Farbcodedarstellung der Farbskalafarbe (z. B. #FF0000 rot).|
||[formula](/javascript/api/excel/excel.conditionalcolorscalecriterion#formula)|Eine Zahl, eine Formel oder `null` (wenn `type` ist `lowestValue` ).|
||[type](/javascript/api/excel/excel.conditionalcolorscalecriterion#type)|Worauf die Bedingungsformel des Kriteriums basieren sollte.|
|[ConditionalDataBarNegativeFormat](/javascript/api/excel/excel.conditionaldatabarnegativeformat)|[borderColor](/javascript/api/excel/excel.conditionaldatabarnegativeformat#bordercolor)|HTML-Farbcode, der die Farbe der Rahmenlinie darstellt, in form #RRGGBB (z. B. "FFA500") oder als benannte HTML-Farbe (z. B. "Orange").|
||[fillColor](/javascript/api/excel/excel.conditionaldatabarnegativeformat#fillcolor)|HTML-Farbcode, der die Füllfarbe darstellt, im Format #RRGGBB (z. B. "FFA500") oder als benannte HTML-Farbe (z. B. "orange").|
||[matchPositiveBorderColor](/javascript/api/excel/excel.conditionaldatabarnegativeformat#matchpositivebordercolor)|Gibt an, ob die negative Datenleiste die gleiche Rahmenfarbe wie die positive Datenleiste hat.|
||[matchPositiveFillColor](/javascript/api/excel/excel.conditionaldatabarnegativeformat#matchpositivefillcolor)|Gibt an, ob die negative Datenleiste die gleiche Füllfarbe wie die positive Datenleiste hat.|
|[ConditionalDataBarPositiveFormat](/javascript/api/excel/excel.conditionaldatabarpositiveformat)|[borderColor](/javascript/api/excel/excel.conditionaldatabarpositiveformat#bordercolor)|HTML-Farbcode, der die Farbe der Rahmenlinie darstellt, in form #RRGGBB (z. B. "FFA500") oder als benannte HTML-Farbe (z. B. "Orange").|
||[fillColor](/javascript/api/excel/excel.conditionaldatabarpositiveformat#fillcolor)|HTML-Farbcode, der die Füllfarbe darstellt, im Format #RRGGBB (z. B. "FFA500") oder als benannte HTML-Farbe (z. B. "orange").|
||[gradientFill](/javascript/api/excel/excel.conditionaldatabarpositiveformat#gradientfill)|Gibt an, ob die Datenleiste über einen Farbverlauf verfügt.|
|[ConditionalDataBarRule](/javascript/api/excel/excel.conditionaldatabarrule)|[formula](/javascript/api/excel/excel.conditionaldatabarrule#formula)|Die Formel, falls erforderlich, für die die Datenbalkenregel ausgewertet werden soll.|
||[type](/javascript/api/excel/excel.conditionaldatabarrule#type)|Der Typ der Regel für die Datenleiste.|
|[ConditionalFormat](/javascript/api/excel/excel.conditionalformat)|[delete()](/javascript/api/excel/excel.conditionalformat#delete--)|Löscht dieses bedingte Format.|
||[getRange()](/javascript/api/excel/excel.conditionalformat#getrange--)|Gibt den Bereich zurück, auf den das bedingte Format angewendet wird.|
||[getRangeOrNullObject()](/javascript/api/excel/excel.conditionalformat#getrangeornullobject--)|Gibt den Bereich zurück, auf den das konditonale Format angewendet wird.|
||[priorität](/javascript/api/excel/excel.conditionalformat#priority)|Die Priorität (oder der Index) innerhalb der bedingten Formatsammlung, in der dieses bedingte Format derzeit vorhanden ist.|
||[cellValue](/javascript/api/excel/excel.conditionalformat#cellvalue)|Gibt die Eigenschaften des bedingten Zellenwerts zurück, wenn das aktuelle bedingte Format ein Typ `CellValue` ist.|
||[cellValueOrNullObject](/javascript/api/excel/excel.conditionalformat#cellvalueornullobject)|Gibt die Eigenschaften des bedingten Zellenwerts zurück, wenn das aktuelle bedingte Format ein Typ `CellValue` ist.|
||[colorScale](/javascript/api/excel/excel.conditionalformat#colorscale)|Gibt die Eigenschaften des bedingten Farbskalenformats zurück, wenn es sich bei dem aktuellen bedingten Format um einen Typ `ColorScale` handelt.|
||[colorScaleOrNullObject](/javascript/api/excel/excel.conditionalformat#colorscaleornullobject)|Gibt die Eigenschaften des bedingten Farbskalenformats zurück, wenn es sich bei dem aktuellen bedingten Format um einen Typ `ColorScale` handelt.|
||[custom](/javascript/api/excel/excel.conditionalformat#custom)|Gibt die benutzerdefinierten Eigenschaften des bedingten Formats zurück, wenn es sich bei dem aktuellen bedingten Format um einen benutzerdefinierten Typ handelt.|
||[customOrNullObject](/javascript/api/excel/excel.conditionalformat#customornullobject)|Gibt die benutzerdefinierten Eigenschaften des bedingten Formats zurück, wenn es sich bei dem aktuellen bedingten Format um einen benutzerdefinierten Typ handelt.|
||[dataBar](/javascript/api/excel/excel.conditionalformat#databar)|Gibt die Eigenschaften der Datenleiste zurück, wenn das aktuelle bedingte Format eine Datenleiste ist.|
||[dataBarOrNullObject](/javascript/api/excel/excel.conditionalformat#databarornullobject)|Gibt die Eigenschaften der Datenleiste zurück, wenn das aktuelle bedingte Format eine Datenleiste ist.|
||[iconSet](/javascript/api/excel/excel.conditionalformat#iconset)|Gibt die Eigenschaften für das bedingte Format des Symbolsatzs zurück, wenn es sich bei dem aktuellen bedingten Format um einen Typ `IconSet` handelt.|
||[iconSetOrNullObject](/javascript/api/excel/excel.conditionalformat#iconsetornullobject)|Gibt die Eigenschaften für das bedingte Format des Symbolsatzs zurück, wenn es sich bei dem aktuellen bedingten Format um einen Typ `IconSet` handelt.|
||[id](/javascript/api/excel/excel.conditionalformat#id)|Die Priorität des bedingten Formats im aktuellen `ConditionalFormatCollection` .|
||[preset](/javascript/api/excel/excel.conditionalformat#preset)|Gibt das voreingestellte bedingte Format für Kriterien zurück.|
||[presetOrNullObject](/javascript/api/excel/excel.conditionalformat#presetornullobject)|Gibt das voreingestellte bedingte Format für Kriterien zurück.|
||[textComparison](/javascript/api/excel/excel.conditionalformat#textcomparison)|Gibt die spezifischen Eigenschaften des bedingten Textformats zurück, wenn es sich bei dem aktuellen bedingten Format um einen Texttyp handelt.|
||[textComparisonOrNullObject](/javascript/api/excel/excel.conditionalformat#textcomparisonornullobject)|Gibt die spezifischen Eigenschaften des bedingten Textformats zurück, wenn es sich bei dem aktuellen bedingten Format um einen Texttyp handelt.|
||[topBottom](/javascript/api/excel/excel.conditionalformat#topbottom)|Gibt die Eigenschaften des bedingten Formats oben/unten zurück, wenn es sich bei dem aktuellen bedingten Format um einen Typ `TopBottom` handelt.|
||[topBottomOrNullObject](/javascript/api/excel/excel.conditionalformat#topbottomornullobject)|Gibt die Eigenschaften des bedingten Formats oben/unten zurück, wenn es sich bei dem aktuellen bedingten Format um einen Typ `TopBottom` handelt.|
||[type](/javascript/api/excel/excel.conditionalformat#type)|Ein Typ des bedingten Formats.|
||[stopIfTrue](/javascript/api/excel/excel.conditionalformat#stopiftrue)|Wenn die Bedingungen dieses bedingten Formats erfüllt sind, werden keine Formate niedrigerer Priorität für diese Zelle wirksam.|
|[ConditionalFormatCollection](/javascript/api/excel/excel.conditionalformatcollection)|[add(type: Excel.ConditionalFormatType)](/javascript/api/excel/excel.conditionalformatcollection#add-type-)|Fügt der Auflistung ein neues bedingtes Format mit der ersten/obersten Priorität hinzu.|
||[clearAll()](/javascript/api/excel/excel.conditionalformatcollection#clearall--)|Löscht alle bedingten Formate, die im aktuellen angegebenen Bereich aktiv sind.|
||[getCount()](/javascript/api/excel/excel.conditionalformatcollection#getcount--)|Gibt die Anzahl der bedingten Formate in der Arbeitsmappe zurück.|
||[getItem(id: string)](/javascript/api/excel/excel.conditionalformatcollection#getitem-id-)|Gibt ein bedingtes Format für die angegebene ID zurück.|
||[getItemAt(index: number)](/javascript/api/excel/excel.conditionalformatcollection#getitemat-index-)|Gibt ein bedingtes Format am angegebenen Index zurück.|
||[items](/javascript/api/excel/excel.conditionalformatcollection#items)|Ruft die geladenen untergeordneten Elemente in dieser Sammlung ab.|
|[ConditionalFormatRule](/javascript/api/excel/excel.conditionalformatrule)|[formula](/javascript/api/excel/excel.conditionalformatrule#formula)|Die Formel, falls erforderlich, für die die Regel für das bedingte Format ausgewertet werden soll.|
||[formulaLocal](/javascript/api/excel/excel.conditionalformatrule#formulalocal)|Die Formel, falls erforderlich, für die die Regel für das bedingte Format in der Sprache des Benutzers ausgewertet werden soll.|
||[formulaR1C1](/javascript/api/excel/excel.conditionalformatrule#formular1c1)|Die Formel, falls erforderlich, für die die Regel für das bedingte Format in der R1C1-Formatvorlage ausgewertet werden soll.|
|[ConditionalIconCriterion](/javascript/api/excel/excel.conditionaliconcriterion)|[customIcon](/javascript/api/excel/excel.conditionaliconcriterion#customicon)|Das benutzerdefinierte Symbol für das aktuelle Kriterium wird zurückgegeben, wenn es sich vom Standardsymbolsatz `null` unterscheiden sollte.|
||[formula](/javascript/api/excel/excel.conditionaliconcriterion#formula)|Eine Zahl oder eine Formel, je nach Typ.|
||[operator](/javascript/api/excel/excel.conditionaliconcriterion#operator)|`greaterThan` oder `greaterThanOrEqual` für jeden Regeltyp für das bedingte Symbolformat.|
||[type](/javascript/api/excel/excel.conditionaliconcriterion#type)|Die Basis für die bedingte Symbolformel.|
|[ConditionalPresetCriteriaRule](/javascript/api/excel/excel.conditionalpresetcriteriarule)|[Kriterium](/javascript/api/excel/excel.conditionalpresetcriteriarule#criterion)|Das Kriterium des bedingten Formats.|
|[ConditionalRangeBorder](/javascript/api/excel/excel.conditionalrangeborder)|[color](/javascript/api/excel/excel.conditionalrangeborder#color)|HTML-Farbcode, der die Farbe der Rahmenlinie darstellt, in form #RRGGBB (z. B. "FFA500") oder als benannte HTML-Farbe (z. B. "Orange").|
||[sideIndex](/javascript/api/excel/excel.conditionalrangeborder#sideindex)|Konstanter Wert, der die bestimmte Seiten des Rahmens angibt.|
||[style](/javascript/api/excel/excel.conditionalrangeborder#style)|Einer der Konstanten der Linienart, die die Linienart für den Rahmen angibt.|
|[ConditionalRangeBorderCollection](/javascript/api/excel/excel.conditionalrangebordercollection)|[getItem(index: Excel.ConditionalRangeBorderIndex)](/javascript/api/excel/excel.conditionalrangebordercollection#getitem-index-)|Ruft ein Rahmen-Objekt ab, das den Namen verwendet|
||[getItemAt(index: number)](/javascript/api/excel/excel.conditionalrangebordercollection#getitemat-index-)|Ruft ein Rahmen-Objekt ab, das den Namen verwendet|
||[bottom](/javascript/api/excel/excel.conditionalrangebordercollection#bottom)|Ruft den unteren Rahmen ab.|
||[count](/javascript/api/excel/excel.conditionalrangebordercollection#count)|Die Anzahl der Rahmen-Objekte in der Auflistung.|
||[items](/javascript/api/excel/excel.conditionalrangebordercollection#items)|Ruft die geladenen untergeordneten Elemente in dieser Sammlung ab.|
||[left](/javascript/api/excel/excel.conditionalrangebordercollection#left)|Ruft den linken Rahmen ab.|
||[right](/javascript/api/excel/excel.conditionalrangebordercollection#right)|Ruft den rechten Rahmen ab.|
||[top](/javascript/api/excel/excel.conditionalrangebordercollection#top)|Ruft den oberen Rahmen ab.|
|[ConditionalRangeFill](/javascript/api/excel/excel.conditionalrangefill)|[clear()](/javascript/api/excel/excel.conditionalrangefill#clear--)|Setzt die Füllung zurück.|
||[color](/javascript/api/excel/excel.conditionalrangefill#color)|HTML-Farbcode, der die Farbe der Füllung darstellt, im Format #RRGGBB (z. B. "FFA500") oder als benannte HTML-Farbe (z. B. "Orange").|
|[ConditionalRangeFont](/javascript/api/excel/excel.conditionalrangefont)|[bold](/javascript/api/excel/excel.conditionalrangefont#bold)|Gibt an, ob die Schriftart fett formatiert ist.|
||[clear()](/javascript/api/excel/excel.conditionalrangefont#clear--)|Setzt die Schriftartformate zurück.|
||[color](/javascript/api/excel/excel.conditionalrangefont#color)|HTML-Farbcodedarstellung der Textfarbe (z. B. #FF0000 rot).|
||[italic](/javascript/api/excel/excel.conditionalrangefont#italic)|Gibt an, ob die Schriftart italisch ist.|
||[strikethrough](/javascript/api/excel/excel.conditionalrangefont#strikethrough)|Gibt den Durchschlagsstatus der Schriftart an.|
||[underline](/javascript/api/excel/excel.conditionalrangefont#underline)|Der Unterstreichungstyp, der auf die Schriftart angewendet wird.|
|[ConditionalRangeFormat](/javascript/api/excel/excel.conditionalrangeformat)|[numberFormat](/javascript/api/excel/excel.conditionalrangeformat#numberformat)|Stellt den Zahlenformatcode von Excel für den angegebenen Bereich dar.|
||[borders](/javascript/api/excel/excel.conditionalrangeformat#borders)|Auflistung von Rahmenobjekten, die für den gesamten bereich des bedingten Formats gelten.|
||[fill](/javascript/api/excel/excel.conditionalrangeformat#fill)|Gibt das fill-Objekt zurück, das im bereich des allgemeinen bedingten Formats definiert ist.|
||[font](/javascript/api/excel/excel.conditionalrangeformat#font)|Gibt das font-Objekt zurück, das im bereich des allgemeinen bedingten Formats definiert ist.|
|[ConditionalTextComparisonRule](/javascript/api/excel/excel.conditionaltextcomparisonrule)|[operator](/javascript/api/excel/excel.conditionaltextcomparisonrule#operator)|Der Operator des bedingten Textformats.|
||[text](/javascript/api/excel/excel.conditionaltextcomparisonrule#text)|Der Textwert des bedingten Formats.|
|[ConditionalTopBottomRule](/javascript/api/excel/excel.conditionaltopbottomrule)|[rank](/javascript/api/excel/excel.conditionaltopbottomrule#rank)|Der Rang zwischen 1 und 1000 für numerische Ränge oder 1 und 100 als Prozentränge.|
||[type](/javascript/api/excel/excel.conditionaltopbottomrule#type)|Formatieren Sie Werte basierend auf dem oberen oder unteren Rang.|
|[CustomConditionalFormat](/javascript/api/excel/excel.customconditionalformat)|[format](/javascript/api/excel/excel.customconditionalformat#format)|Gibt ein Formatobjekt zurück, das die Eigenschaften Schriftart, Füllung, Rahmen und andere bedingte Formate kapselt.|
||[rule](/javascript/api/excel/excel.customconditionalformat#rule)|Gibt das Objekt `Rule` in diesem bedingten Format an.|
|[DataBarConditionalFormat](/javascript/api/excel/excel.databarconditionalformat)|[axisColor](/javascript/api/excel/excel.databarconditionalformat#axiscolor)|HTML-Farbcode, der die Farbe der Achsenlinie im Format #RRGGBB (z. B. "FFA500") oder als benannte HTML-Farbe (z. B. "Orange") darstellt.|
||[axisFormat](/javascript/api/excel/excel.databarconditionalformat#axisformat)|Darstellung, wie die Achse für eine Excel-Datenleiste bestimmt wird.|
||[barDirection](/javascript/api/excel/excel.databarconditionalformat#bardirection)|Gibt die Richtung an, auf der die Datenbalkengrafik basieren soll.|
||[lowerBoundRule](/javascript/api/excel/excel.databarconditionalformat#lowerboundrule)|Die Regel, die bestimmt, was die untere Grenze für einen Datenbalken darstellt (und wie diese ggf. berechnet wird).|
||[negativeFormat](/javascript/api/excel/excel.databarconditionalformat#negativeformat)|Darstellung aller Werte links neben der Achse in einer Excel-Datenleiste.|
||[positiveFormat](/javascript/api/excel/excel.databarconditionalformat#positiveformat)|Darstellung aller Werte rechts neben der Achse in einer Excel-Datenleiste.|
||[showDataBarOnly](/javascript/api/excel/excel.databarconditionalformat#showdatabaronly)|Wenn `true` , blendet die Werte aus den Zellen aus, in denen die Datenleiste angewendet wird.|
||[upperBoundRule](/javascript/api/excel/excel.databarconditionalformat#upperboundrule)|Die Regel, die bestimmt, was die obere Grenze für einen Datenbalken darstellt (und wie diese ggf. berechnet wird).|
|[IconSetConditionalFormat](/javascript/api/excel/excel.iconsetconditionalformat)|[criteria](/javascript/api/excel/excel.iconsetconditionalformat#criteria)|Ein Array von Kriterien und Symbolsätzen für die Regeln und potenziellen benutzerdefinierten Symbole für bedingte Symbole.|
||[reverseIconOrder](/javascript/api/excel/excel.iconsetconditionalformat#reverseiconorder)|Wenn `true` , wird die Symbolbestellung für den Symbolsatz umgekehrt.|
||[showIconOnly](/javascript/api/excel/excel.iconsetconditionalformat#showicononly)|Wenn `true` , blendet die Werte aus und zeigt nur Symbole an.|
||[style](/javascript/api/excel/excel.iconsetconditionalformat#style)|Wenn festgelegt, wird die Symbolsatzoption für das bedingte Format angezeigt.|
|[PresetCriteriaConditionalFormat](/javascript/api/excel/excel.presetcriteriaconditionalformat)|[format](/javascript/api/excel/excel.presetcriteriaconditionalformat#format)|Gibt ein Formatobjekt zurück, das die Eigenschaften Schriftart, Füllung, Rahmen und andere bedingte Formate kapselt.|
||[rule](/javascript/api/excel/excel.presetcriteriaconditionalformat#rule)|Die Regel des bedingten Formats.|
|[Range](/javascript/api/excel/excel.range)|[calculate()](/javascript/api/excel/excel.range#calculate--)|Berechnet einen Zellbereich auf einem Arbeitsblatt.|
||[conditionalFormats](/javascript/api/excel/excel.range#conditionalformats)|Die Auflistung davon `ConditionalFormats` überschneidet den Bereich.|
|[TextConditionalFormat](/javascript/api/excel/excel.textconditionalformat)|[format](/javascript/api/excel/excel.textconditionalformat#format)|Gibt ein Formatobjekt zurück, das die Schriftart, die Füllung, die Rahmen und andere Eigenschaften des bedingten Formats kapselt.|
||[rule](/javascript/api/excel/excel.textconditionalformat#rule)|Die Regel des bedingten Formats.|
|[TopBottomConditionalFormat](/javascript/api/excel/excel.topbottomconditionalformat)|[format](/javascript/api/excel/excel.topbottomconditionalformat#format)|Gibt ein Formatobjekt zurück, das die Schriftart, die Füllung, die Rahmen und andere Eigenschaften des bedingten Formats kapselt.|
||[rule](/javascript/api/excel/excel.topbottomconditionalformat#rule)|Die Kriterien des bedingten Formats oben/unten.|
|[Arbeitsblatt](/javascript/api/excel/excel.worksheet)|[calculate(markAllDirty: boolean)](/javascript/api/excel/excel.worksheet#calculate-markalldirty-)|Berechnet alle Zellen auf einem Arbeitsblatt.|
