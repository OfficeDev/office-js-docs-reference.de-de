| Klasse | Felder | Beschreibung |
|:---|:---|:---|
|[Anwendung](/javascript/api/excel/excel.application)|[CultureInfo](/javascript/api/excel/excel.application#cultureinfo)|Enthält Informationen basierend auf den aktuellen Systemkultur Einstellungen.|
||[decimalSeparator](/javascript/api/excel/excel.application#decimalseparator)|Ruft die Zeichenfolge ab, die als Dezimaltrennzeichen für numerische Werte verwendet wird.|
||[thousandsSeparator](/javascript/api/excel/excel.application#thousandsseparator)|Ruft die Zeichenfolge ab, die zum Trennen von Zifferngruppen links vom Dezimaltrennzeichen für numerische Werte verwendet wird.|
||[useSystemSeparators](/javascript/api/excel/excel.application#usesystemseparators)|Gibt an, ob die Systemtrennzeichen von Excel aktiviert sind.|
|[Kommentar](/javascript/api/excel/excel.comment)|[Erwähnungen](/javascript/api/excel/excel.comment#mentions)|Ruft die Entitäten (z. b. Personen) ab, die in Kommentaren erwähnt werden.|
||[richContent](/javascript/api/excel/excel.comment#richcontent)|Ruft den umfangreichen Kommentar Inhalt ab (beispielsweise Erwähnungen in Kommentaren).|
||[gelöst](/javascript/api/excel/excel.comment#resolved)|Der Kommentar Threadstatus.|
||[updateMentions (contentWithMentions: Excel. CommentRichContent)](/javascript/api/excel/excel.comment#updatementions-contentwithmentions-)|Aktualisiert den Kommentar Inhalt mit einer speziell formatierten Zeichenfolge und einer Liste von Erwähnungen.|
|[CommentCollection](/javascript/api/excel/excel.commentcollection)|[Add (cellAddress: Range \| String, Content: CommentRichContent \| String, ContentType?: Excel. ContentType)](/javascript/api/excel/excel.commentcollection#add-celladdress--content--contenttype-)|Erstellt einen neuen Kommentar mit dem angegebenen Inhalt auf der angegebenen Zelle.|
|[CommentMention](/javascript/api/excel/excel.commentmention)|[email](/javascript/api/excel/excel.commentmention#email)|Die e-Mail-Adresse der Entität, die in Comment erwähnt wird.|
||[id](/javascript/api/excel/excel.commentmention#id)|Die ID der Entität.|
||[name](/javascript/api/excel/excel.commentmention#name)|Der Name der Entität, die in Comment erwähnt wird.|
|[CommentReply](/javascript/api/excel/excel.commentreply)|[Erwähnungen](/javascript/api/excel/excel.commentreply#mentions)|Die in Kommentaren erwähnten Entitäten (beispielsweise Personen).|
||[gelöst](/javascript/api/excel/excel.commentreply#resolved)|Der Kommentar Antwortstatus.|
||[richContent](/javascript/api/excel/excel.commentreply#richcontent)|Der umfangreiche Kommentar Inhalt (beispielsweise Erwähnungen in Kommentaren).|
||[updateMentions (contentWithMentions: Excel. CommentRichContent)](/javascript/api/excel/excel.commentreply#updatementions-contentwithmentions-)|Aktualisiert den Kommentar Inhalt mit einer speziell formatierten Zeichenfolge und einer Liste von Erwähnungen.|
|[CommentReplyCollection](/javascript/api/excel/excel.commentreplycollection)|[Add (Content: CommentRichContent \| String, ContentType?: Excel. ContentType)](/javascript/api/excel/excel.commentreplycollection#add-content--contenttype-)|Erstellt eine Kommentarantwort für einen Kommentar.|
|[CommentRichContent](/javascript/api/excel/excel.commentrichcontent)|[Erwähnungen](/javascript/api/excel/excel.commentrichcontent#mentions)|Ein Array mit allen Entitäten (z. b. Personen), die im Kommentar erwähnt werden.|
||[richContent](/javascript/api/excel/excel.commentrichcontent#richcontent)|Gibt den umfangreichen Inhalt des Kommentars an (beispielsweise Kommentar Inhalt mit Erwähnungen, die erste erwähnte Entität verfügt über das ID-Attribut 0, und die zweite erwähnte Entität verfügt über das ID-Attribut 1).|
|[CultureInfo](/javascript/api/excel/excel.cultureinfo)|[name](/javascript/api/excel/excel.cultureinfo#name)|Ruft den Namen der Kultur im Format languagecode2-Country/regioncode2 (z. b., "zh-cn" oder "en-US") ab.|
||[NumberFormat](/javascript/api/excel/excel.cultureinfo#numberformat)|Definiert das kulturell geeignete Format für die Anzeige von Zahlen.|
|[NumberFormatInfo](/javascript/api/excel/excel.numberformatinfo)|[numberDecimalSeparator](/javascript/api/excel/excel.numberformatinfo#numberdecimalseparator)|Ruft die Zeichenfolge ab, die als Dezimaltrennzeichen für numerische Werte verwendet wird.|
||[numberGroupSeparator](/javascript/api/excel/excel.numberformatinfo#numbergroupseparator)|Ruft die Zeichenfolge ab, die zum Trennen von Zifferngruppen links vom Dezimaltrennzeichen für numerische Werte verwendet wird.|
|[Range](/javascript/api/excel/excel.range)|[MoveTo (destinationRange: Bereichs \| Zeichenfolge)](/javascript/api/excel/excel.range#moveto-destinationrange-)|Verschiebt Zellenwerte, Formatierungen und Formeln aus dem aktuellen Bereich in den Zielbereich, wobei die alten Informationen in diesen Zellen ersetzt werden.|
|[RangeFormat](/javascript/api/excel/excel.rangeformat)|[adjustIndent (Betrag: Zahl)](/javascript/api/excel/excel.rangeformat#adjustindent-amount-)|Passt den Einzug der Bereichs Formatierung an.|
|[Workbook](/javascript/api/excel/excel.workbook)|[close(closeBehavior?: Excel.CloseBehavior)](/javascript/api/excel/excel.workbook#close-closebehavior-)|Aktuelle Arbeitsmappe schließen.|
||[save(saveBehavior?: Excel.SaveBehavior)](/javascript/api/excel/excel.workbook#save-savebehavior-)|Aktuelle Arbeitsmappe speichern.|
|[Arbeitsblatt](/javascript/api/excel/excel.worksheet)|[onRowHiddenChanged](/javascript/api/excel/excel.worksheet#onrowhiddenchanged)|Tritt auf, wenn sich der ausgeblendete Status einer oder mehrerer Zeilen in einem bestimmten Arbeitsblatt geändert hat.|
|[WorksheetCalculatedEventArgs](/javascript/api/excel/excel.worksheetcalculatedeventargs)|[address](/javascript/api/excel/excel.worksheetcalculatedeventargs#address)|Die Adresse des Bereichs, der die Berechnung abgeschlossen hat.|
|[WorksheetCollection](/javascript/api/excel/excel.worksheetcollection)|[onRowHiddenChanged](/javascript/api/excel/excel.worksheetcollection#onrowhiddenchanged)|Tritt auf, wenn sich der ausgeblendete Status einer oder mehrerer Zeilen in einem bestimmten Arbeitsblatt geändert hat.|
|[WorksheetRowHiddenChangedEventArgs](/javascript/api/excel/excel.worksheetrowhiddenchangedeventargs)|[address](/javascript/api/excel/excel.worksheetrowhiddenchangedeventargs#address)|Ruft die Bereichsadresse ab, die den geänderten Bereich eines bestimmten Arbeitsblatts darstellt.|
||[changeType](/javascript/api/excel/excel.worksheetrowhiddenchangedeventargs#changetype)|Ruft den Typ der Änderung ab, der angibt, wie das Ereignis ausgelöst wurde.|
||[source](/javascript/api/excel/excel.worksheetrowhiddenchangedeventargs#source)|Ruft die Quelle des Ereignisses ab.|
||[Typ](/javascript/api/excel/excel.worksheetrowhiddenchangedeventargs#type)|Ruft den Typ des Ereignisses ab.|
||[worksheetId](/javascript/api/excel/excel.worksheetrowhiddenchangedeventargs#worksheetid)|Ruft die ID des Arbeitsblatts ab, auf dem die Daten geändert wurden.|
