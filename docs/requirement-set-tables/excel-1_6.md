| Klasse | Felder | Beschreibung |
|:---|:---|:---|
|[Anwendung](/javascript/api/excel/excel.application)|[suspendApiCalculationUntilNextSync()](/javascript/api/excel/excel.application#suspendapicalculationuntilnextsync--)|Hält die Berechnung an, bis die nächste „context.sync()“ aufgerufen wird.|
|[CellValueConditionalFormat](/javascript/api/excel/excel.cellvalueconditionalformat)|[format](/javascript/api/excel/excel.cellvalueconditionalformat#format)|Gibt ein Format-Objekt zurück, das die Schriftart, Füllung, Rahmen und andere Eigenschaften für bedingte Formate kapselt.|
||[Regel](/javascript/api/excel/excel.cellvalueconditionalformat#rule)|Gibt das Rule-Objekt für dieses bedingte Format an.|
|[ColorScaleConditionalFormat](/javascript/api/excel/excel.colorscaleconditionalformat)|[criteria](/javascript/api/excel/excel.colorscaleconditionalformat#criteria)|Die Kriterien der Farb Skala.|
||[threeColorScale](/javascript/api/excel/excel.colorscaleconditionalformat#threecolorscale)|Bei true wird die Farbskala mit drei Punkten (Minimum, Mittelpunkt, Maximum) verwendet, andernfalls hat Sie zwei (minimal, Maximum).|
|[ConditionalCellValueRule](/javascript/api/excel/excel.conditionalcellvaluerule)|[Formula1](/javascript/api/excel/excel.conditionalcellvaluerule#formula1)|Die Formel, sofern erforderlich, nach der die bedingte Formatregel ausgewertet werden soll.|
||[Formula2](/javascript/api/excel/excel.conditionalcellvaluerule#formula2)|Die Formel, sofern erforderlich, nach der die bedingte Formatregel ausgewertet werden soll.|
||[operator](/javascript/api/excel/excel.conditionalcellvaluerule#operator)|Der Operator des bedingten Formats von Zellenwert.|
|[ConditionalColorScaleCriteria](/javascript/api/excel/excel.conditionalcolorscalecriteria)|[maximum](/javascript/api/excel/excel.conditionalcolorscalecriteria#maximum)|Das Maximalpunktkriterium der Farbskala.|
||[Mitte](/javascript/api/excel/excel.conditionalcolorscalecriteria#midpoint)|Das Mittelpunktkriterium der Farbskala im Fall einer 3-Punkte-Farbskala.|
||[minimum](/javascript/api/excel/excel.conditionalcolorscalecriteria#minimum)|Das Minimumpunktkriterium der Farbskala.|
|[ConditionalColorScaleCriterion](/javascript/api/excel/excel.conditionalcolorscalecriterion)|[color](/javascript/api/excel/excel.conditionalcolorscalecriterion#color)|HTML-Farb Codedarstellung der Farbskalen Farbe (beispielsweise #FF0000 rot).|
||[formula](/javascript/api/excel/excel.conditionalcolorscalecriterion#formula)|Eine Zahl, eine Formel oder NULL (für Typ „LowestValue“).|
||[type](/javascript/api/excel/excel.conditionalcolorscalecriterion#type)|Worauf das Kriterium bedingte Formel basieren soll.|
|[ConditionalDataBarNegativeFormat](/javascript/api/excel/excel.conditionaldatabarnegativeformat)|[BorderColor](/javascript/api/excel/excel.conditionaldatabarnegativeformat#bordercolor)|HTML-Farbcode, der die Farbe der Rahmenlinie, des Formulars #RRGGBB (z. B.  "FFA500") oder als benannte HTML-Farbe (z. B. "orange") darstellt.|
||[FillColor](/javascript/api/excel/excel.conditionaldatabarnegativeformat#fillcolor)|HTML-Farbcode, der die Füllfarbe darstellt, des Formulars #RRGGBB (beispielsweise "FFA500") oder als benannte HTML-Farbe (beispielsweise "Orange").|
||[matchPositiveBorderColor](/javascript/api/excel/excel.conditionaldatabarnegativeformat#matchpositivebordercolor)|Gibt an, ob der negative Datenleiste dieselbe Rahmenfarbe wie der positive Datenbereich aufweist.|
||[matchPositiveFillColor](/javascript/api/excel/excel.conditionaldatabarnegativeformat#matchpositivefillcolor)|Gibt an, ob der negative Datenbereich dieselbe Füllfarbe wie der positive Datenbereich aufweist.|
|[ConditionalDataBarPositiveFormat](/javascript/api/excel/excel.conditionaldatabarpositiveformat)|[BorderColor](/javascript/api/excel/excel.conditionaldatabarpositiveformat#bordercolor)|HTML-Farbcode, der die Farbe der Rahmenlinie, des Formulars #RRGGBB (z. B.  "FFA500") oder als benannte HTML-Farbe (z. B. "orange") darstellt.|
||[FillColor](/javascript/api/excel/excel.conditionaldatabarpositiveformat#fillcolor)|HTML-Farbcode, der die Füllfarbe darstellt, des Formulars #RRGGBB (beispielsweise "FFA500") oder als benannte HTML-Farbe (beispielsweise "Orange").|
||[gradientFill](/javascript/api/excel/excel.conditionaldatabarpositiveformat#gradientfill)|Gibt an, ob der Datenbalken einen Farbverlauf aufweist.|
|[ConditionalDataBarRule](/javascript/api/excel/excel.conditionaldatabarrule)|[formula](/javascript/api/excel/excel.conditionaldatabarrule#formula)|Die Formel, sofern erforderlich, nach der die Datenbalkenregel ausgewertet werden soll.|
||[type](/javascript/api/excel/excel.conditionaldatabarrule#type)|Der Typ der Regel für den datenbar.|
|[ConditionalFormat](/javascript/api/excel/excel.conditionalformat)|[delete()](/javascript/api/excel/excel.conditionalformat#delete--)|Löscht dieses bedingte Format.|
||[getRange()](/javascript/api/excel/excel.conditionalformat#getrange--)|Gibt den Bereich zurück, auf den das bedingte Format angewendet wird.|
||[getRangeOrNullObject()](/javascript/api/excel/excel.conditionalformat#getrangeornullobject--)|Gibt den Bereich zurück, auf den das bedingte-Format angewendet wird, oder ein NULL-Objekt, wenn das bedingte Format auf mehrere Bereiche angewendet wird.|
||[priority](/javascript/api/excel/excel.conditionalformat#priority)|Die Priorität (oder der Index) innerhalb der bedingten Format Auflistung, in der sich dieses bedingte Format derzeit befindet.|
||[cellValue](/javascript/api/excel/excel.conditionalformat#cellvalue)|Gibt die bedingten Formateigenschaften des Zellenwerts zurück, wenn es sich bei dem aktuellen bedingten Format um einen cellvalue-Typ handelt.|
||[cellValueOrNullObject](/javascript/api/excel/excel.conditionalformat#cellvalueornullobject)|Gibt die bedingten Formateigenschaften des Zellenwerts zurück, wenn es sich bei dem aktuellen bedingten Format um einen cellvalue-Typ handelt.|
||[colorScale](/javascript/api/excel/excel.conditionalformat#colorscale)|Gibt die ColorScale bedingten Formateigenschaften zurück, wenn das aktuelle bedingte Format ein ColorScale-Typ ist.|
||[colorScaleOrNullObject](/javascript/api/excel/excel.conditionalformat#colorscaleornullobject)|Gibt die ColorScale bedingten Formateigenschaften zurück, wenn das aktuelle bedingte Format ein ColorScale-Typ ist.|
||[Custom](/javascript/api/excel/excel.conditionalformat#custom)|Gibt die benutzerdefinierten bedingten Formateigenschaften zurück, wenn das aktuelle bedingte Format ein benutzerdefinierter Typ ist.|
||[customOrNullObject](/javascript/api/excel/excel.conditionalformat#customornullobject)|Gibt die benutzerdefinierten bedingten Formateigenschaften zurück, wenn das aktuelle bedingte Format ein benutzerdefinierter Typ ist.|
||[dataBar](/javascript/api/excel/excel.conditionalformat#databar)|Gibt die Eigenschaften des Datenbalkens zurück, wenn es sich bei dem aktuellen bedingten Format um einen Datenbalken handelt.|
||[dataBarOrNullObject](/javascript/api/excel/excel.conditionalformat#databarornullobject)|Gibt die Eigenschaften des Datenbalkens zurück, wenn es sich bei dem aktuellen bedingten Format um einen Datenbalken handelt.|
||[iconSet](/javascript/api/excel/excel.conditionalformat#iconset)|Gibt die bedingten Formateigenschaften von Iconset zurück, wenn es sich bei dem aktuellen bedingten Format um einen Iconset-Typ handelt.|
||[iconSetOrNullObject](/javascript/api/excel/excel.conditionalformat#iconsetornullobject)|Gibt die bedingten Formateigenschaften von Iconset zurück, wenn es sich bei dem aktuellen bedingten Format um einen Iconset-Typ handelt.|
||[id](/javascript/api/excel/excel.conditionalformat#id)|Die Priorität des bedingten Formats in der aktuellen ConditionalFormatCollection.|
||[voreingestellten](/javascript/api/excel/excel.conditionalformat#preset)|Gibt das bedingte Format der voreingestellten Kriterien zurück.|
||[presetOrNullObject](/javascript/api/excel/excel.conditionalformat#presetornullobject)|Gibt das bedingte Format der voreingestellten Kriterien zurück.|
||[Textvergleich](/javascript/api/excel/excel.conditionalformat#textcomparison)|Gibt die spezifischen Text bedingten Formateigenschaften zurück, wenn es sich bei dem aktuellen bedingten Format um einen Texttyp handelt.|
||[textComparisonOrNullObject](/javascript/api/excel/excel.conditionalformat#textcomparisonornullobject)|Gibt die spezifischen Text bedingten Formateigenschaften zurück, wenn es sich bei dem aktuellen bedingten Format um einen Texttyp handelt.|
||[topBottom](/javascript/api/excel/excel.conditionalformat#topbottom)|Gibt die oberen/unteren bedingten Formateigenschaften zurück, wenn es sich bei dem aktuellen bedingten Format um einen untersten Typ handelt.|
||[topBottomOrNullObject](/javascript/api/excel/excel.conditionalformat#topbottomornullobject)|Gibt die oberen/unteren bedingten Formateigenschaften zurück, wenn es sich bei dem aktuellen bedingten Format um einen untersten Typ handelt.|
||[type](/javascript/api/excel/excel.conditionalformat#type)|Ein Typ von bedingter Formatierung.|
||[stopIfTrue](/javascript/api/excel/excel.conditionalformat#stopiftrue)|Wenn die Bedingungen dieses bedingten Formats erfüllt sind, werden keine Formate niedrigerer Priorität für diese Zelle wirksam.|
|[ConditionalFormatCollection](/javascript/api/excel/excel.conditionalformatcollection)|[Add (Typ: Excel. ConditionalFormatType)](/javascript/api/excel/excel.conditionalformatcollection#add-type-)|Fügt der Auflistung ein neues bedingtes Format an der ersten/obersten Priorität hinzu.|
||[clearAll ()](/javascript/api/excel/excel.conditionalformatcollection#clearall--)|Löscht alle bedingten Formate, die im aktuellen angegebenen Bereich aktiv sind.|
||[getCount()](/javascript/api/excel/excel.conditionalformatcollection#getcount--)|Gibt die Anzahl der bedingten Formate in der Arbeitsmappe zurück.|
||[getItem(id: string)](/javascript/api/excel/excel.conditionalformatcollection#getitem-id-)|Gibt ein bedingtes Format für die angegebene ID zurück.|
||[getItemAt(index: number)](/javascript/api/excel/excel.conditionalformatcollection#getitemat-index-)|Gibt ein bedingtes Format am angegebenen Index zurück.|
||[items](/javascript/api/excel/excel.conditionalformatcollection#items)|Ruft die geladenen untergeordneten Elemente in dieser Sammlung ab.|
|[ConditionalFormatRule](/javascript/api/excel/excel.conditionalformatrule)|[formula](/javascript/api/excel/excel.conditionalformatrule#formula)|Die Formel, sofern erforderlich, nach der die bedingte Formatregel ausgewertet werden soll.|
||[formulaLocal](/javascript/api/excel/excel.conditionalformatrule#formulalocal)|Die Formel in der Benutzersprache, sofern erforderlich, nach der die bedingte Formatregel ausgewertet werden soll.|
||[formulaR1C1](/javascript/api/excel/excel.conditionalformatrule#formular1c1)|Die Formel in R1C1-Notation, sofern erforderlich, nach der die bedingte Formatregel ausgewertet werden soll.|
|[ConditionalIconCriterion](/javascript/api/excel/excel.conditionaliconcriterion)|[customIcon](/javascript/api/excel/excel.conditionaliconcriterion#customicon)|Das benutzerdefinierte Symbol für das aktuelle Kriterium, sofern verschieden vom Standard-IconSet; andernfalls wird NULL zurückgegeben.|
||[formula](/javascript/api/excel/excel.conditionaliconcriterion#formula)|Eine Zahl oder eine Formel, je nach Typ.|
||[operator](/javascript/api/excel/excel.conditionaliconcriterion#operator)|GreaterThan oder GreaterThanOrEqual für jeden Typ des Regel Typs für das bedingte Symbol Format.|
||[type](/javascript/api/excel/excel.conditionaliconcriterion#type)|Die Basis für die bedingte Symbolformel.|
|[ConditionalPresetCriteriaRule](/javascript/api/excel/excel.conditionalpresetcriteriarule)|[Kriterium](/javascript/api/excel/excel.conditionalpresetcriteriarule#criterion)|Das Kriterium des bedingten Formats.|
|[ConditionalRangeBorder](/javascript/api/excel/excel.conditionalrangeborder)|[color](/javascript/api/excel/excel.conditionalrangeborder#color)|HTML-Farbcode, der die Farbe der Rahmenlinie, des Formulars #RRGGBB (z. B.  "FFA500") oder als benannte HTML-Farbe (z. B. "orange") darstellt.|
||[sideIndex](/javascript/api/excel/excel.conditionalrangeborder#sideindex)|Konstanter Wert, der die bestimmte Seiten des Rahmens angibt.|
||[style](/javascript/api/excel/excel.conditionalrangeborder#style)|Einer der Konstanten der Linienart, die die Linienart für den Rahmen angibt.|
|[ConditionalRangeBorderCollection](/javascript/api/excel/excel.conditionalrangebordercollection)|[GetItem (Index: Excel. ConditionalRangeBorderIndex)](/javascript/api/excel/excel.conditionalrangebordercollection#getitem-index-)|Ruft ein Rahmen-Objekt ab, das den Namen verwendet|
||[getItemAt(index: number)](/javascript/api/excel/excel.conditionalrangebordercollection#getitemat-index-)|Ruft ein Rahmen-Objekt ab, das den Namen verwendet|
||[bottom](/javascript/api/excel/excel.conditionalrangebordercollection#bottom)|Ruft die untere Rahmenlinie ab.|
||[count](/javascript/api/excel/excel.conditionalrangebordercollection#count)|Die Anzahl der Rahmen-Objekte in der Auflistung.|
||[items](/javascript/api/excel/excel.conditionalrangebordercollection#items)|Ruft die geladenen untergeordneten Elemente in dieser Sammlung ab.|
||[left](/javascript/api/excel/excel.conditionalrangebordercollection#left)|Ruft den linken Rahmen ab.|
||[right](/javascript/api/excel/excel.conditionalrangebordercollection#right)|Ruft den rechten Rahmen ab.|
||[top](/javascript/api/excel/excel.conditionalrangebordercollection#top)|Ruft den oberen Rahmen ab.|
|[ConditionalRangeFill](/javascript/api/excel/excel.conditionalrangefill)|[clear()](/javascript/api/excel/excel.conditionalrangefill#clear--)|Setzt die Füllung zurück.|
||[color](/javascript/api/excel/excel.conditionalrangefill#color)|HTML-Farbcode, der die Farbe der Füllung, des Formulars #RRGGBB (beispielsweise "FFA500") oder als benannte HTML-Farbe (beispielsweise "Orange") darstellt.|
|[ConditionalRangeFont](/javascript/api/excel/excel.conditionalrangefont)|[bold](/javascript/api/excel/excel.conditionalrangefont#bold)|Gibt an, ob die Schriftart fett formatiert ist.|
||[clear()](/javascript/api/excel/excel.conditionalrangefont#clear--)|Setzt die Schriftartformate zurück.|
||[color](/javascript/api/excel/excel.conditionalrangefont#color)|HTML-Farb Codedarstellung der Textfarbe (z. b. #FF0000 steht für rot).|
||[italic](/javascript/api/excel/excel.conditionalrangefont#italic)|Gibt an, ob die Schriftart kursiv formatiert ist.|
||[strikethrough](/javascript/api/excel/excel.conditionalrangefont#strikethrough)|Gibt den durchgestrichen Status der Schriftart an.|
||[underline](/javascript/api/excel/excel.conditionalrangefont#underline)|Der Typ der Unterstreichung, der auf die Schriftart angewendet wird.|
|[ConditionalRangeFormat](/javascript/api/excel/excel.conditionalrangeformat)|[NumberFormat](/javascript/api/excel/excel.conditionalrangeformat#numberformat)|Stellt den Zahlenformatcode für Excel für den angegebenen Bereich dar.|
||[borders](/javascript/api/excel/excel.conditionalrangeformat#borders)|Auflistung von Border-Objekten, die auf den allgemeinen bedingten Formatbereich angewendet werden.|
||[fill](/javascript/api/excel/excel.conditionalrangeformat#fill)|Gibt das Fill-Objekt zurück, das für den allgemeinen bedingten Formatierungsbereich definiert ist.|
||[font](/javascript/api/excel/excel.conditionalrangeformat#font)|Gibt das Font-Objekt zurück, das im allgemeinen bedingten Formatbereich definiert ist.|
|[ConditionalTextComparisonRule](/javascript/api/excel/excel.conditionaltextcomparisonrule)|[operator](/javascript/api/excel/excel.conditionaltextcomparisonrule#operator)|Der Operator des bedingten Textformats.|
||[text](/javascript/api/excel/excel.conditionaltextcomparisonrule#text)|Der Textwert des bedingten Formats.|
|[ConditionalTopBottomRule](/javascript/api/excel/excel.conditionaltopbottomrule)|[rank](/javascript/api/excel/excel.conditionaltopbottomrule#rank)|Der Rang zwischen 1 und 1000 für numerische Ränge oder 1 und 100 als Prozentränge.|
||[type](/javascript/api/excel/excel.conditionaltopbottomrule#type)|Formatieren Sie Werte basierend auf dem oberen oder unteren Rang.|
|[CustomConditionalFormat](/javascript/api/excel/excel.customconditionalformat)|[format](/javascript/api/excel/excel.customconditionalformat#format)|Gibt ein Format-Objekt zurück, das die Schriftart, Füllung, Rahmen und andere Eigenschaften für bedingte Formate kapselt.|
||[Regel](/javascript/api/excel/excel.customconditionalformat#rule)|Gibt das Rule-Objekt für dieses bedingte Format an.|
|[DataBarConditionalFormat](/javascript/api/excel/excel.databarconditionalformat)|[axisColor](/javascript/api/excel/excel.databarconditionalformat#axiscolor)|HTML-Farbcode, der die Farbe der Achsenlinie, des Formulars #RRGGBB (z. b. "FFA500") oder als benannte HTML-Farbe (beispielsweise "Orange") darstellt.|
||[axisFormat](/javascript/api/excel/excel.databarconditionalformat#axisformat)|Darstellung der Bestimmung der Achse für einen Excel-Datenbalken.|
||[barDirection](/javascript/api/excel/excel.databarconditionalformat#bardirection)|Gibt die Richtung an, auf der die Datenbalken Grafik basieren soll.|
||[lowerBoundRule](/javascript/api/excel/excel.databarconditionalformat#lowerboundrule)|Die Regel, die bestimmt, was die untere Grenze für einen Datenbalken darstellt (und wie diese ggf. berechnet wird).|
||[negativeFormat](/javascript/api/excel/excel.databarconditionalformat#negativeformat)|Darstellung aller Werte Links von der Achse in einem Excel-Datenbalken.|
||[positiveFormat](/javascript/api/excel/excel.databarconditionalformat#positiveformat)|Darstellung aller Werte rechts von der Achse in einem Excel-Datenbalken.|
||[showDataBarOnly](/javascript/api/excel/excel.databarconditionalformat#showdatabaronly)|Wenn „true“, werden die Werte in den Zellen ausgeblendet, auf die der Datenbalken angewendet wird.|
||[upperBoundRule](/javascript/api/excel/excel.databarconditionalformat#upperboundrule)|Die Regel, die bestimmt, was die obere Grenze für einen Datenbalken darstellt (und wie diese ggf. berechnet wird).|
|[IconSetConditionalFormat](/javascript/api/excel/excel.iconsetconditionalformat)|[criteria](/javascript/api/excel/excel.iconsetconditionalformat#criteria)|Ein Array von Kriterien und IconSets für die Regeln und potenziellen benutzerdefinierten Symbole für bedingte Symbole.|
||[reverseIconOrder](/javascript/api/excel/excel.iconsetconditionalformat#reverseiconorder)|Wenn true, werden die Symbol Reihenfolgen für das Iconset umgekehrt.|
||[showIconOnly](/javascript/api/excel/excel.iconsetconditionalformat#showicononly)|Wenn „true“, werden die Werte ausgeblendet und nur die Symbole angezeigt.|
||[style](/javascript/api/excel/excel.iconsetconditionalformat#style)|Wenn dieser Wert festgelegt ist, wird die Iconset-Option für das bedingte Format angezeigt.|
|[PresetCriteriaConditionalFormat](/javascript/api/excel/excel.presetcriteriaconditionalformat)|[format](/javascript/api/excel/excel.presetcriteriaconditionalformat#format)|Gibt ein Format-Objekt zurück, das die Schriftart, Füllung, Rahmen und andere Eigenschaften für bedingte Formate kapselt.|
||[Regel](/javascript/api/excel/excel.presetcriteriaconditionalformat#rule)|Die Regel des bedingten Formats.|
|[Range](/javascript/api/excel/excel.range)|[calculate()](/javascript/api/excel/excel.range#calculate--)|Berechnet einen Zellbereich auf einem Arbeitsblatt.|
||[conditionalFormats](/javascript/api/excel/excel.range#conditionalformats)|Die ConditionalFormats-Auflistung, die den Bereich schneidet.|
|[TextConditionalFormat](/javascript/api/excel/excel.textconditionalformat)|[format](/javascript/api/excel/excel.textconditionalformat#format)|Gibt ein Format-Objekt zurück, das die Schriftart, Füllung, Rahmen und andere Eigenschaften der bedingten Formatierung kapselt.|
||[Regel](/javascript/api/excel/excel.textconditionalformat#rule)|Die Regel des bedingten Formats.|
|[TopBottomConditionalFormat](/javascript/api/excel/excel.topbottomconditionalformat)|[format](/javascript/api/excel/excel.topbottomconditionalformat#format)|Gibt ein Format-Objekt zurück, das die Schriftart, Füllung, Rahmen und andere Eigenschaften der bedingten Formatierung kapselt.|
||[Regel](/javascript/api/excel/excel.topbottomconditionalformat#rule)|Die Kriterien des bedingten Formats oben/unten.|
|[Arbeitsblatt](/javascript/api/excel/excel.worksheet)|[Calculate (markAllDirty: Boolean)](/javascript/api/excel/excel.worksheet#calculate-markalldirty-)|Berechnet alle Zellen auf einem Arbeitsblatt.|
