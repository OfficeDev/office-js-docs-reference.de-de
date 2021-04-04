| Klasse | Felder | Beschreibung |
|:---|:---|:---|
|[Anwendung](/javascript/api/excel/excel.application)|[cultureInfo](/javascript/api/excel/excel.application#cultureinfo)|Stellt Informationen basierend auf den aktuellen Systemkultureinstellungen zur Verfügung.|
||[decimalSeparator](/javascript/api/excel/excel.application#decimalseparator)|Ruft die Zeichenfolge ab, die als Dezimaltrennzeichen für numerische Werte verwendet wird.|
||[thousandsSeparator](/javascript/api/excel/excel.application#thousandsseparator)|Ruft die Zeichenfolge ab, die zum Trennen von Zifferngruppen links vom Dezimaltrennzeichen für numerische Werte verwendet wird.|
||[useSystemSeparators](/javascript/api/excel/excel.application#usesystemseparators)|Gibt an, ob die Systemtrennzeichen von Excel aktiviert sind.|
|[Kommentar](/javascript/api/excel/excel.comment)|[Erwähnungen](/javascript/api/excel/excel.comment#mentions)|Ruft die Entitäten (z. B. Personen) ab, die in Kommentaren erwähnt werden.|
||[richContent](/javascript/api/excel/excel.comment#richcontent)|Ruft den reichhaltigen Kommentarinhalt ab (z. B. Erwähnungen in Kommentaren).|
||[aufgelöst](/javascript/api/excel/excel.comment#resolved)|Der Kommentarthreadstatus.|
||[updateMentions(contentWithMentions: Excel.CommentRichContent)](/javascript/api/excel/excel.comment#updatementions-contentwithmentions-)|Aktualisiert den Kommentarinhalt mit einer speziell formatierten Zeichenfolge und einer Liste von Erwähnungen.|
|[CommentCollection](/javascript/api/excel/excel.commentcollection)|[add(cellAddress: Range \| string, content: CommentRichContent \| string, contentType?: Excel.ContentType)](/javascript/api/excel/excel.commentcollection#add-celladdress--content--contenttype-)|Erstellt einen neuen Kommentar mit dem angegebenen Inhalt auf der angegebenen Zelle.|
|[CommentMention](/javascript/api/excel/excel.commentmention)|[email](/javascript/api/excel/excel.commentmention#email)|Die E-Mail-Adresse der Entität, die in einem Kommentar erwähnt wird.|
||[id](/javascript/api/excel/excel.commentmention#id)|Die ID der Entität.|
||[name](/javascript/api/excel/excel.commentmention#name)|Der Name der Entität, die in einem Kommentar erwähnt wird.|
|[CommentReply](/javascript/api/excel/excel.commentreply)|[Erwähnungen](/javascript/api/excel/excel.commentreply#mentions)|Die Entitäten (z. B. Personen), die in Kommentaren erwähnt werden.|
||[aufgelöst](/javascript/api/excel/excel.commentreply#resolved)|Der Kommentarantwortstatus.|
||[richContent](/javascript/api/excel/excel.commentreply#richcontent)|Der umfangreiche Kommentarinhalt (z. B. Erwähnungen in Kommentaren).|
||[updateMentions(contentWithMentions: Excel.CommentRichContent)](/javascript/api/excel/excel.commentreply#updatementions-contentwithmentions-)|Aktualisiert den Kommentarinhalt mit einer speziell formatierten Zeichenfolge und einer Liste von Erwähnungen.|
|[CommentReplyCollection](/javascript/api/excel/excel.commentreplycollection)|[add(content: CommentRichContent \| string, contentType?: Excel.ContentType)](/javascript/api/excel/excel.commentreplycollection#add-content--contenttype-)|Erstellt eine Kommentarantwort für einen Kommentar.|
|[CommentRichContent](/javascript/api/excel/excel.commentrichcontent)|[Erwähnungen](/javascript/api/excel/excel.commentrichcontent#mentions)|Ein Array mit allen Entitäten (z. B. Personen), die im Kommentar erwähnt werden.|
||[richContent](/javascript/api/excel/excel.commentrichcontent#richcontent)|Gibt den reichhaltigen Inhalt des Kommentars an (z. B. Kommentarinhalt mit Erwähnungen, die erste erwähnte Entität hat das ID-Attribut 0, und die zweite erwähnte Entität hat das ID-Attribut 1).|
|[CultureInfo](/javascript/api/excel/excel.cultureinfo)|[name](/javascript/api/excel/excel.cultureinfo#name)|Ruft den Kulturnamen im Format languagecode2-country/regioncode2 ab (z. B. "zh-cn" oder "en-us").|
||[numberFormat](/javascript/api/excel/excel.cultureinfo#numberformat)|Definiert das kulturell geeignete Format der Anzeige von Zahlen.|
|[NumberFormatInfo](/javascript/api/excel/excel.numberformatinfo)|[numberDecimalSeparator](/javascript/api/excel/excel.numberformatinfo#numberdecimalseparator)|Ruft die Zeichenfolge ab, die als Dezimaltrennzeichen für numerische Werte verwendet wird.|
||[numberGroupSeparator](/javascript/api/excel/excel.numberformatinfo#numbergroupseparator)|Ruft die Zeichenfolge ab, die zum Trennen von Zifferngruppen links vom Dezimaltrennzeichen für numerische Werte verwendet wird.|
|[Range](/javascript/api/excel/excel.range)|[moveTo(destinationRange: Range \| string)](/javascript/api/excel/excel.range#moveto-destinationrange-)|Verschiebt Zellenwerte, Formatierungen und Formeln aus dem aktuellen Bereich in den Zielbereich und ersetzt dabei die alten Informationen in diesen Zellen.|
|[RangeFormat](/javascript/api/excel/excel.rangeformat)|[adjustIndent(amount: number)](/javascript/api/excel/excel.rangeformat#adjustindent-amount-)|Passt den Einzug der Bereichsformatierung an.|
|[Workbook](/javascript/api/excel/excel.workbook)|[close(closeBehavior?: Excel.CloseBehavior)](/javascript/api/excel/excel.workbook#close-closebehavior-)|Aktuelle Arbeitsmappe schließen.|
||[save(saveBehavior?: Excel.SaveBehavior)](/javascript/api/excel/excel.workbook#save-savebehavior-)|Aktuelle Arbeitsmappe speichern.|
|[Arbeitsblatt](/javascript/api/excel/excel.worksheet)|[onRowHiddenChanged](/javascript/api/excel/excel.worksheet#onrowhiddenchanged)|Tritt auf, wenn sich der ausgeblendete Zustand einer oder mehreren Zeilen in einem bestimmten Arbeitsblatt geändert hat.|
|[WorksheetCalculatedEventArgs](/javascript/api/excel/excel.worksheetcalculatedeventargs)|[address](/javascript/api/excel/excel.worksheetcalculatedeventargs#address)|Die Adresse des Bereichs, der die Berechnung abgeschlossen hat.|
|[WorksheetCollection](/javascript/api/excel/excel.worksheetcollection)|[onRowHiddenChanged](/javascript/api/excel/excel.worksheetcollection#onrowhiddenchanged)|Tritt auf, wenn sich der ausgeblendete Zustand einer oder mehreren Zeilen in einem bestimmten Arbeitsblatt geändert hat.|
|[WorksheetRowHiddenChangedEventArgs](/javascript/api/excel/excel.worksheetrowhiddenchangedeventargs)|[address](/javascript/api/excel/excel.worksheetrowhiddenchangedeventargs#address)|Ruft die Bereichsadresse ab, die den geänderten Bereich eines bestimmten Arbeitsblatts darstellt.|
||[changeType](/javascript/api/excel/excel.worksheetrowhiddenchangedeventargs#changetype)|Ruft den Typ der Änderung ab, der darstellt, wie das Ereignis ausgelöst wurde.|
||[source](/javascript/api/excel/excel.worksheetrowhiddenchangedeventargs#source)|Ruft die Quelle des Ereignisses ab.|
||[type](/javascript/api/excel/excel.worksheetrowhiddenchangedeventargs#type)|Ruft den Typ des Ereignisses ab.|
||[worksheetId](/javascript/api/excel/excel.worksheetrowhiddenchangedeventargs#worksheetid)|Ruft die ID des Arbeitsblatts ab, in dem sich die Daten geändert haben.|
