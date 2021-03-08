| Klasse | Felder | Beschreibung |
|:---|:---|:---|
|[Anwendung](/javascript/api/word/word.application)|[createDocument(base64File?: string)](/javascript/api/word/word.application#createdocument-base64file-)|Erstellt ein neues Dokument mithilfe einer optionalen base64-codierten DOCX-Datei.|
|[Body](/javascript/api/word/word.body)|[getRange(rangeLocation?: Word.RangeLocation)](/javascript/api/word/word.body#getrange-rangelocation-)|Ruft den gesamten Text oder den Start- bzw. Endpunkt des Textkörpers als Bereich ab.|
||[insertTable(rowCount: number, columnCount: number, insertLocation: Word.InsertLocation, values?: string[][])](/javascript/api/word/word.body#inserttable-rowcount--columncount--insertlocation--values-)|Fügt eine Tabelle mit der angegebenen Anzahl von Zeilen und Spalten ein.|
||[Listen](/javascript/api/word/word.body#lists)|Ruft die Sammlung von Listenobjekten im Textkörper ab.|
||[parentBody](/javascript/api/word/word.body#parentbody)|Ruft den übergeordneten Text des Texts ab.|
||[parentBodyOrNullObject](/javascript/api/word/word.body#parentbodyornullobject)|Ruft den übergeordneten Text des Texts ab.|
||[parentContentControlOrNullObject](/javascript/api/word/word.body#parentcontentcontrolornullobject)|Ruft das Inhaltssteuerelement ab, das den Textkörper enthält.|
||[parentSection](/javascript/api/word/word.body#parentsection)|Ruft den übergeordneten Abschnitt des Texts ab.|
||[parentSectionOrNullObject](/javascript/api/word/word.body#parentsectionornullobject)|Ruft den übergeordneten Abschnitt des Texts ab.|
||[Tabellen](/javascript/api/word/word.body#tables)|Ruft die Sammlung von Tabellenobjekten im Text ab.|
||[type](/javascript/api/word/word.body#type)|Ruft den Typ des Texts ab.|
||[styleBuiltIn](/javascript/api/word/word.body#stylebuiltin)|Ruft den Namen der integrierten Formatvorlage für den Text ab oder legt ihn fest.|
|[ContentControl](/javascript/api/word/word.contentcontrol)|[getRange(rangeLocation?: Word.RangeLocation)](/javascript/api/word/word.contentcontrol#getrange-rangelocation-)|Ruft das gesamte Inhaltssteuerelement oder den Start- bzw. Endpunkt des Inhaltssteuerelements als Bereich ab.|
||[getTextRanges(endingMarks: string[], trimSpacing?: boolean)](/javascript/api/word/word.contentcontrol#gettextranges-endingmarks--trimspacing-)|Ruft die Textbereiche im Inhaltssteuerelement mithilfe von Satzzeichen und/oder anderen Endzeichen ab.|
||[insertTable(rowCount: number, columnCount: number, insertLocation: Word.InsertLocation, values?: string[][])](/javascript/api/word/word.contentcontrol#inserttable-rowcount--columncount--insertlocation--values-)|Fügt eine Tabelle mit der angegebenen Anzahl von Zeilen und Spalten in oder neben ein Inhaltssteuerelement ein.|
||[Listen](/javascript/api/word/word.contentcontrol#lists)|Ruft die Sammlung von Listenobjekten im Inhaltssteuerelement ab.|
||[parentBody](/javascript/api/word/word.contentcontrol#parentbody)|Ruft den übergeordneten Text des Inhaltssteuerelements ab.|
||[parentContentControlOrNullObject](/javascript/api/word/word.contentcontrol#parentcontentcontrolornullobject)|Ruft das Inhaltssteuerelement ab, das das Inhaltssteuerelement enthält.|
||[parentTable](/javascript/api/word/word.contentcontrol#parenttable)|Ruft die Tabelle ab, die das Inhaltssteuerelement enthält.|
||[parentTableCell](/javascript/api/word/word.contentcontrol#parenttablecell)|Ruft die Tabellenzelle ab, die das Inhaltssteuerelement enthält.|
||[parentTableCellOrNullObject](/javascript/api/word/word.contentcontrol#parenttablecellornullobject)|Ruft die Tabellenzelle ab, die das Inhaltssteuerelement enthält.|
||[parentTableOrNullObject](/javascript/api/word/word.contentcontrol#parenttableornullobject)|Ruft die Tabelle ab, die das Inhaltssteuerelement enthält.|
||[subtype](/javascript/api/word/word.contentcontrol#subtype)|Ruft den Untertyp des Inhaltssteuerelements ab.|
||[Tabellen](/javascript/api/word/word.contentcontrol#tables)|Ruft die Sammlung von Tabellenobjekten im Inhaltssteuerelement ab.|
||[split(delimiters: string[], multiParagraphs?: boolean, trimDelimiters?: boolean, trimSpacing?: boolean)](/javascript/api/word/word.contentcontrol#split-delimiters--multiparagraphs--trimdelimiters--trimspacing-)|Teilt das Inhaltssteuerelement mithilfe von Trennzeichen in untergeordnete Bereiche.|
||[styleBuiltIn](/javascript/api/word/word.contentcontrol#stylebuiltin)|Ruft den Namen der integrierten Formatvorlage für das Inhaltssteuerelement ab oder legt ihn fest.|
|[ContentControlCollection](/javascript/api/word/word.contentcontrolcollection)|[getByIdOrNullObject(id: number)](/javascript/api/word/word.contentcontrolcollection#getbyidornullobject-id-)|Ruft ein Inhaltssteuerelement mithilfe der ID ab.|
||[getByTypes(types: Word.ContentControlType[])](/javascript/api/word/word.contentcontrolcollection#getbytypes-types-)|Ruft die Inhaltssteuerelemente ab, die die angegebenen Typen und/oder Untertypen enthalten.|
||[getFirst()](/javascript/api/word/word.contentcontrolcollection#getfirst--)|Ruft das erste Inhaltssteuerelement in dieser Sammlung ab.|
||[getFirstOrNullObject()](/javascript/api/word/word.contentcontrolcollection#getfirstornullobject--)|Ruft das erste Inhaltssteuerelement in dieser Sammlung ab.|
|[CustomProperty](/javascript/api/word/word.customproperty)|[delete()](/javascript/api/word/word.customproperty#delete--)|Löscht die benutzerdefinierte Eigenschaft.|
||[key](/javascript/api/word/word.customproperty#key)|Ruft den Schlüssel der benutzerdefinierten Eigenschaft ab.|
||[type](/javascript/api/word/word.customproperty#type)|Ruft den Wertetyp der benutzerdefinierten Eigenschaft ab.|
||[value](/javascript/api/word/word.customproperty#value)|Ruft den Wert der benutzerdefinierten Eigenschaft ab oder legt ihn fest.|
|[CustomPropertyCollection](/javascript/api/word/word.custompropertycollection)|[add(key: string, value: any)](/javascript/api/word/word.custompropertycollection#add-key--value-)|Erstellt eine neue benutzerdefinierte Eigenschaft oder legt eine vorhandene fest.|
||[deleteAll()](/javascript/api/word/word.custompropertycollection#deleteall--)|Löscht alle benutzerdefinierten Eigenschaften in dieser Sammlung.|
||[getCount()](/javascript/api/word/word.custompropertycollection#getcount--)|Ruft die Anzahl von benutzerdefinierten Eigenschaften ab.|
||[getItem(key: string)](/javascript/api/word/word.custompropertycollection#getitem-key-)|Ruft ein Objekt für benutzerdefinierte Eigenschaften über seinen Schlüssel ab, bei dem Groß-/Kleinschreibung nicht beachtet wird.|
||[getItemOrNullObject(key: string)](/javascript/api/word/word.custompropertycollection#getitemornullobject-key-)|Ruft ein Objekt für benutzerdefinierte Eigenschaften über seinen Schlüssel ab, bei dem Groß-/Kleinschreibung nicht beachtet wird.|
||[items](/javascript/api/word/word.custompropertycollection#items)|Ruft die geladenen untergeordneten Elemente in dieser Sammlung ab.|
|[Dokument](/javascript/api/word/word.document)|[properties](/javascript/api/word/word.document#properties)|Ruft die Eigenschaften des Dokuments ab.|
|[DocumentCreated](/javascript/api/word/word.documentcreated)|[open()](/javascript/api/word/word.documentcreated#open--)|Öffnet das Dokument.|
||[Text](/javascript/api/word/word.documentcreated#body)|Ruft das body-Objekt des Dokuments ab.|
||[contentControls](/javascript/api/word/word.documentcreated#contentcontrols)|Ruft die Auflistung von Inhaltssteuerelementobjekten im Dokument ab.|
||[properties](/javascript/api/word/word.documentcreated#properties)|Ruft die Eigenschaften des Dokuments ab.|
||[gespeichert](/javascript/api/word/word.documentcreated#saved)|Gibt an, ob die Änderungen im Dokument gespeichert wurden.|
||[sections](/javascript/api/word/word.documentcreated#sections)|Ruft die Auflistung von Abschnittsobjekten im Dokument ab.|
||[save()](/javascript/api/word/word.documentcreated#save--)|Speichert das Dokument.|
|[DocumentProperties](/javascript/api/word/word.documentproperties)|[author](/javascript/api/word/word.documentproperties#author)|Ruft den Autor des Dokuments ab oder legt ihn fest.|
||[category](/javascript/api/word/word.documentproperties#category)|Ruft die Kategorie des Dokuments ab oder legt sie fest.|
||[comments](/javascript/api/word/word.documentproperties#comments)|Ruft die Kommentare des Dokuments ab oder legt sie fest.|
||[company](/javascript/api/word/word.documentproperties#company)|Ruft das Unternehmen des Dokuments ab oder legt es fest.|
||[format](/javascript/api/word/word.documentproperties#format)|Ruft das Format des Dokuments ab oder legt es fest.|
||[Keywords](/javascript/api/word/word.documentproperties#keywords)|Ruft die Schlüsselwörter des Dokuments ab oder legt sie fest.|
||[manager](/javascript/api/word/word.documentproperties#manager)|Ruft den Manager des Dokuments ab oder legt ihn fest.|
||[applicationName](/javascript/api/word/word.documentproperties#applicationname)|Ruft den Anwendungsnamen des Dokuments ab.|
||[creationDate](/javascript/api/word/word.documentproperties#creationdate)|Ruft das Erstellungsdatum des Dokuments ab.|
||[customProperties](/javascript/api/word/word.documentproperties#customproperties)|Ruft die Sammlung der benutzerdefinierten Eigenschaften des Dokuments ab.|
||[lastAuthor](/javascript/api/word/word.documentproperties#lastauthor)|Ruft den letzten Autor des Dokuments ab.|
||[lastPrintDate](/javascript/api/word/word.documentproperties#lastprintdate)|Ruft das letzte Druckdatum des Dokuments ab.|
||[lastSaveTime](/javascript/api/word/word.documentproperties#lastsavetime)|Ruft die Zeit der letzten Speicherung des Dokuments ab.|
||[revisionNumber](/javascript/api/word/word.documentproperties#revisionnumber)|Ruft die Revisionsnummer des Dokuments ab.|
||[Sicherheit](/javascript/api/word/word.documentproperties#security)|Ruft Sicherheitseinstellungen des Dokuments ab.|
||[template](/javascript/api/word/word.documentproperties#template)|Ruft die Vorlage des Dokuments ab.|
||[Betreff](/javascript/api/word/word.documentproperties#subject)|Ruft den Betreff des Dokuments ab oder legt ihn fest.|
||[title](/javascript/api/word/word.documentproperties#title)|Ruft den Titel des Dokuments ab oder legt ihn fest.|
|[InlinePicture](/javascript/api/word/word.inlinepicture)|[getNext()](/javascript/api/word/word.inlinepicture#getnext--)|Ruft das nächste Inlinebild ab.|
||[getNextOrNullObject()](/javascript/api/word/word.inlinepicture#getnextornullobject--)|Ruft das nächste Inlinebild ab.|
||[getRange(rangeLocation?: Word.RangeLocation)](/javascript/api/word/word.inlinepicture#getrange-rangelocation-)|Ruft das Bild oder den Start- bzw. Endpunkt des Bilds als Bereich ab.|
||[parentContentControlOrNullObject](/javascript/api/word/word.inlinepicture#parentcontentcontrolornullobject)|Ruft das Inhaltssteuerelement ab, das das eingebundene Bild enthält.|
||[parentTable](/javascript/api/word/word.inlinepicture#parenttable)|Ruft die Tabelle ab, die das Inlinebild enthält.|
||[parentTableCell](/javascript/api/word/word.inlinepicture#parenttablecell)|Ruft die Tabellenzelle ab, die das Inlinebild enthält.|
||[parentTableCellOrNullObject](/javascript/api/word/word.inlinepicture#parenttablecellornullobject)|Ruft die Tabellenzelle ab, die das Inlinebild enthält.|
||[parentTableOrNullObject](/javascript/api/word/word.inlinepicture#parenttableornullobject)|Ruft die Tabelle ab, die das Inlinebild enthält.|
|[InlinePictureCollection](/javascript/api/word/word.inlinepicturecollection)|[getFirst()](/javascript/api/word/word.inlinepicturecollection#getfirst--)|Ruft das erste Inlinebild in dieser Sammlung ab.|
||[getFirstOrNullObject()](/javascript/api/word/word.inlinepicturecollection#getfirstornullobject--)|Ruft das erste Inlinebild in dieser Sammlung ab.|
|[List](/javascript/api/word/word.list)|[getLevelParagraphs(level: number)](/javascript/api/word/word.list#getlevelparagraphs-level-)|Ruft die Absätze auf der angegebenen Ebene in der Liste ab.|
||[getLevelString(level: number)](/javascript/api/word/word.list#getlevelstring-level-)|Ruft das Aufzählungszeichen, die Nummer oder das Bild auf der angegebenen Ebene als Zeichenfolge ab.|
||[insertParagraph(paragraphText: string, insertLocation: Word.InsertLocation)](/javascript/api/word/word.list#insertparagraph-paragraphtext--insertlocation-)|Fügt an der angegebenen Position einen Absatz ein.|
||[id](/javascript/api/word/word.list#id)|Ruft die ID der Liste ab.|
||[levelExistences](/javascript/api/word/word.list#levelexistences)|Überprüft, ob jede der 9 Ebenen in der Liste vorhanden ist.|
||[levelTypes](/javascript/api/word/word.list#leveltypes)|Ruft die Typen aller 9 Ebenen in der Liste ab.|
||[paragraphs](/javascript/api/word/word.list#paragraphs)|Ruft die Absätze in der Liste ab.|
||[setLevelAlignment(level: number, alignment: Word.Alignment)](/javascript/api/word/word.list#setlevelalignment-level--alignment-)|Legt die Ausrichtung des Aufzählungszeichens, der Nummer oder des Bilds auf der angegebenen Ebene in der Liste fest.|
||[setLevelBullet(level: number, listBullet: Word.ListBullet, charCode?: number, fontName?: string)](/javascript/api/word/word.list#setlevelbullet-level--listbullet--charcode--fontname-)|Legt das Format des Aufzählungszeichens auf der angegebenen Ebene in der Liste fest.|
||[setLevelIndents(level: number, textIndent: number, bulletNumberPictureIndent: number)](/javascript/api/word/word.list#setlevelindents-level--textindent--bulletnumberpictureindent-)|Legt die zwei Einzüge der angegebenen Ebene in der Liste fest.|
||[setLevelNumbering(level: number, listNumbering: Word.ListNumbering, formatString?: Array<\| string number>)](/javascript/api/word/word.list#setlevelnumbering-level--listnumbering--formatstring-)|Legt das Nummerierungsformat auf der angegebenen Ebene in der Liste fest.|
||[setLevelStartingNumber(level: number, startingNumber: number)](/javascript/api/word/word.list#setlevelstartingnumber-level--startingnumber-)|Legt die Anfangsnummer auf der angegebenen Ebene in der Liste fest.|
|[listCollection](/javascript/api/word/word.listcollection)|[getById(id: number)](/javascript/api/word/word.listcollection#getbyid-id-)|Ruft eine Liste über ihren Bezeichner ab.|
||[getByIdOrNullObject(id: number)](/javascript/api/word/word.listcollection#getbyidornullobject-id-)|Ruft eine Liste über ihren Bezeichner ab.|
||[getFirst()](/javascript/api/word/word.listcollection#getfirst--)|Ruft die erste Liste in dieser Sammlung ab.|
||[getFirstOrNullObject()](/javascript/api/word/word.listcollection#getfirstornullobject--)|Ruft die erste Liste in dieser Sammlung ab.|
||[getItem(index: number)](/javascript/api/word/word.listcollection#getitem-index-)|Ruft ein Listenobjekt nach dem Index in der Sammlung ab.|
||[items](/javascript/api/word/word.listcollection#items)|Ruft die geladenen untergeordneten Elemente in dieser Sammlung ab.|
|[ListItem](/javascript/api/word/word.listitem)|[getAncestor(parentOnly?: boolean)](/javascript/api/word/word.listitem#getancestor-parentonly-)|Ruft das übergeordnete Element der Liste bzw. den nächsten Vorgänger ab, wenn das übergeordnete Element nicht vorhanden ist.|
||[getAncestorOrNullObject(parentOnly?: boolean)](/javascript/api/word/word.listitem#getancestorornullobject-parentonly-)|Ruft das übergeordnete Element der Liste bzw. den nächsten Vorgänger ab, wenn das übergeordnete Element nicht vorhanden ist.|
||[getDescendants(directChildrenOnly?: boolean)](/javascript/api/word/word.listitem#getdescendants-directchildrenonly-)|Ruft alle untergeordneten Listenelemente des Listenelements ab.|
||[level](/javascript/api/word/word.listitem#level)|Ruft die Ebene des Elements in der Liste ab oder legt sie fest.|
||[listString](/javascript/api/word/word.listitem#liststring)|Ruft das Aufzählungszeichen, die Nummer oder das Bild des Listenelements als Zeichenfolge ab.|
||[siblingIndex](/javascript/api/word/word.listitem#siblingindex)|Ruft die Ordnungsnummer des Listenelements im Verhältnis zu seinen gleichgeordneten Elementen ab.|
|[Paragraph](/javascript/api/word/word.paragraph)|[attachToList(listId: number, level: number)](/javascript/api/word/word.paragraph#attachtolist-listid--level-)|Fügt den Absatz zu einer vorhandenen Liste auf der angegebenen Ebene hinzu.|
||[detachFromList()](/javascript/api/word/word.paragraph#detachfromlist--)|Verschiebt diesen Absatz aus der Liste, falls der Absatz ein Listenelement ist.|
||[getNext()](/javascript/api/word/word.paragraph#getnext--)|Ruft den nächsten Absatz ab.|
||[getNextOrNullObject()](/javascript/api/word/word.paragraph#getnextornullobject--)|Ruft den nächsten Absatz ab.|
||[getPrevious()](/javascript/api/word/word.paragraph#getprevious--)|Ruft den vorherigen Absatz ab.|
||[getPreviousOrNullObject()](/javascript/api/word/word.paragraph#getpreviousornullobject--)|Ruft den vorherigen Absatz ab.|
||[getRange(rangeLocation?: Word.RangeLocation)](/javascript/api/word/word.paragraph#getrange-rangelocation-)|Ruft den gesamten Absatz oder den Start- bzw. Endpunkt des Absatzes als Bereich ab.|
||[getTextRanges(endingMarks: string[], trimSpacing?: boolean)](/javascript/api/word/word.paragraph#gettextranges-endingmarks--trimspacing-)|Ruft die Textbereiche im Absatz mithilfe von Satzzeichen und/oder anderen Endzeichen ab.|
||[insertTable(rowCount: number, columnCount: number, insertLocation: Word.InsertLocation, values?: string[][])](/javascript/api/word/word.paragraph#inserttable-rowcount--columncount--insertlocation--values-)|Fügt eine Tabelle mit der angegebenen Anzahl von Zeilen und Spalten ein.|
||[isLastParagraph](/javascript/api/word/word.paragraph#islastparagraph)|Gibt an, dass es sich bei dem Absatz um den letzten innerhalb seines übergeordneten Texts handelt.|
||[isListItem](/javascript/api/word/word.paragraph#islistitem)|Überprüft, ob der Absatz ein Listenelement ist.|
||[list](/javascript/api/word/word.paragraph#list)|Ruft die Liste ab, zu der dieser Absatz gehört.|
||[listItem](/javascript/api/word/word.paragraph#listitem)|Ruft das ListItem für den Absatz ab.|
||[listItemOrNullObject](/javascript/api/word/word.paragraph#listitemornullobject)|Ruft das ListItem für den Absatz ab.|
||[listOrNullObject](/javascript/api/word/word.paragraph#listornullobject)|Ruft die Liste ab, zu der dieser Absatz gehört.|
||[parentBody](/javascript/api/word/word.paragraph#parentbody)|Ruft den übergeordneten Text des Absatzes ab.|
||[parentContentControlOrNullObject](/javascript/api/word/word.paragraph#parentcontentcontrolornullobject)|Ruft das Inhaltssteuerelement ab, das den Abschnitt enthält.|
||[parentTable](/javascript/api/word/word.paragraph#parenttable)|Ruft die Tabelle ab, die den Absatz enthält.|
||[parentTableCell](/javascript/api/word/word.paragraph#parenttablecell)|Ruft die Tabellenzelle ab, die den Absatz enthält.|
||[parentTableCellOrNullObject](/javascript/api/word/word.paragraph#parenttablecellornullobject)|Ruft die Tabellenzelle ab, die den Absatz enthält.|
||[parentTableOrNullObject](/javascript/api/word/word.paragraph#parenttableornullobject)|Ruft die Tabelle ab, die den Absatz enthält.|
||[tableNestingLevel](/javascript/api/word/word.paragraph#tablenestinglevel)|Ruft die Ebene der Tabelle des Absatzes ab.|
||[split(delimiters: string[], trimDelimiters?: boolean, trimSpacing?: boolean)](/javascript/api/word/word.paragraph#split-delimiters--trimdelimiters--trimspacing-)|Teilt den Absatz mithilfe von Trennzeichen in untergeordnete Bereiche.|
||[startNewList()](/javascript/api/word/word.paragraph#startnewlist--)|Beginnt eine neue Liste mit diesem Absatz.|
||[styleBuiltIn](/javascript/api/word/word.paragraph#stylebuiltin)|Ruft den Namen der integrierten Formatvorlage für den Absatz ab oder legt ihn fest.|
|[ParagraphCollection](/javascript/api/word/word.paragraphcollection)|[getFirst()](/javascript/api/word/word.paragraphcollection#getfirst--)|Ruft den ersten Absatz in dieser Sammlung ab.|
||[getFirstOrNullObject()](/javascript/api/word/word.paragraphcollection#getfirstornullobject--)|Ruft den ersten Absatz in dieser Sammlung ab.|
||[getLast()](/javascript/api/word/word.paragraphcollection#getlast--)|Ruft den letzten Absatz in dieser Sammlung ab.|
||[getLastOrNullObject()](/javascript/api/word/word.paragraphcollection#getlastornullobject--)|Ruft den letzten Absatz in dieser Sammlung ab.|
|[Range](/javascript/api/word/word.range)|[compareLocationWith(range: Word.Range)](/javascript/api/word/word.range#comparelocationwith-range-)|Vergleicht die Position dieses Bereichs mit der eines anderen Bereichs.|
||[expandTo(range: Word.Range)](/javascript/api/word/word.range#expandto-range-)|Gibt einen neuen Bereich zurück, der diesen Bereich in beide Richtungen erweitert, um einen anderen Bereich zu überdecken.|
||[expandToOrNullObject(range: Word.Range)](/javascript/api/word/word.range#expandtoornullobject-range-)|Gibt einen neuen Bereich zurück, der diesen Bereich in beide Richtungen erweitert, um einen anderen Bereich zu überdecken.|
||[getHyperlinkRanges()](/javascript/api/word/word.range#gethyperlinkranges--)|Ruft untergeordnete Linkbereiche innerhalb des Bereichs ab.|
||[getNextTextRange(endingMarks: string[], trimSpacing?: boolean)](/javascript/api/word/word.range#getnexttextrange-endingmarks--trimspacing-)|Ruft den nächsten Textbereich mithilfe von Satzzeichen und/oder anderen Endzeichen ab.|
||[getNextTextRangeOrNullObject(endingMarks: string[], trimSpacing?: boolean)](/javascript/api/word/word.range#getnexttextrangeornullobject-endingmarks--trimspacing-)|Ruft den nächsten Textbereich mithilfe von Satzzeichen und/oder anderen Endzeichen ab.|
||[getRange(rangeLocation?: Word.RangeLocation)](/javascript/api/word/word.range#getrange-rangelocation-)|Klont den Bereich oder ruft den Anfangs- bzw. Endpunkt des Bereichs als neuen Bereich ab.|
||[getTextRanges(endingMarks: string[], trimSpacing?: boolean)](/javascript/api/word/word.range#gettextranges-endingmarks--trimspacing-)|Ruft die untergeordneten Textbereiche im Bereich mithilfe von Satzzeichen und/oder anderen Endzeichen ab.|
||[hyperlink](/javascript/api/word/word.range#hyperlink)|Ruft den ersten Link in dem Bereich ab oder legt einen Link für den Bereich fest.|
||[insertTable(rowCount: number, columnCount: number, insertLocation: Word.InsertLocation, values?: string[][])](/javascript/api/word/word.range#inserttable-rowcount--columncount--insertlocation--values-)|Fügt eine Tabelle mit der angegebenen Anzahl von Zeilen und Spalten ein.|
||[intersectWith(range: Word.Range)](/javascript/api/word/word.range#intersectwith-range-)|Gibt einen neuen Bereich als Schnittmenge dieses Bereichs mit einem anderen Bereich zurück.|
||[intersectWithOrNullObject(range: Word.Range)](/javascript/api/word/word.range#intersectwithornullobject-range-)|Gibt einen neuen Bereich als Schnittmenge dieses Bereichs mit einem anderen Bereich zurück.|
||[isEmpty](/javascript/api/word/word.range#isempty)|Überprüft, ob die Länge des Bereichs 0 ist.|
||[Listen](/javascript/api/word/word.range#lists)|Ruft die Sammlung von Listenobjekten im Bereich ab.|
||[parentBody](/javascript/api/word/word.range#parentbody)|Ruft den übergeordneten Text des Bereichs ab.|
||[parentContentControlOrNullObject](/javascript/api/word/word.range#parentcontentcontrolornullobject)|Ruft das Inhaltssteuerelement ab, das den Bereich enthält.|
||[parentTable](/javascript/api/word/word.range#parenttable)|Ruft die Tabelle ab, die den Bereich enthält.|
||[parentTableCell](/javascript/api/word/word.range#parenttablecell)|Ruft die Tabellenzelle ab, die den Bereich enthält.|
||[parentTableCellOrNullObject](/javascript/api/word/word.range#parenttablecellornullobject)|Ruft die Tabellenzelle ab, die den Bereich enthält.|
||[parentTableOrNullObject](/javascript/api/word/word.range#parenttableornullobject)|Ruft die Tabelle ab, die den Bereich enthält.|
||[Tabellen](/javascript/api/word/word.range#tables)|Ruft die Sammlung von Tabellenobjekten im Bereich ab.|
||[split(delimiters: string[], multiParagraphs?: boolean, trimDelimiters?: boolean, trimSpacing?: boolean)](/javascript/api/word/word.range#split-delimiters--multiparagraphs--trimdelimiters--trimspacing-)|Teilt den Bereich mithilfe von Trennzeichen in untergeordnete Bereiche.|
||[styleBuiltIn](/javascript/api/word/word.range#stylebuiltin)|Ruft den Namen der integrierten Formatvorlage für den Bereich ab oder legt ihn fest.|
|[RangeCollection](/javascript/api/word/word.rangecollection)|[getFirst()](/javascript/api/word/word.rangecollection#getfirst--)|Ruft den ersten Bereich in dieser Sammlung ab.|
||[getFirstOrNullObject()](/javascript/api/word/word.rangecollection#getfirstornullobject--)|Ruft den ersten Bereich in dieser Sammlung ab.|
|[RequestContext](/javascript/api/word/word.requestcontext)|[application](/javascript/api/word/word.requestcontext#application)|[Api-Satz: WordApi 1.3] *|
|[Section](/javascript/api/word/word.section)|[getNext()](/javascript/api/word/word.section#getnext--)|Ruft den nächsten Abschnitt ab.|
||[getNextOrNullObject()](/javascript/api/word/word.section#getnextornullobject--)|Ruft den nächsten Abschnitt ab.|
|[SectionCollection](/javascript/api/word/word.sectioncollection)|[getFirst()](/javascript/api/word/word.sectioncollection#getfirst--)|Ruft den ersten Abschnitt in dieser Sammlung ab.|
||[getFirstOrNullObject()](/javascript/api/word/word.sectioncollection#getfirstornullobject--)|Ruft den ersten Abschnitt in dieser Sammlung ab.|
|[Table](/javascript/api/word/word.table)|[addColumns(insertLocation: Word.InsertLocation, columnCount: number, values?: string[][])](/javascript/api/word/word.table#addcolumns-insertlocation--columncount--values-)|Fügt am Anfang oder Ende der Tabelle Spalten hinzu, wobei die erste oder letzte vorhandene Spalte als Vorlage verwendet wird.|
||[addRows(insertLocation: Word.InsertLocation, rowCount: number, values?: string[][])](/javascript/api/word/word.table#addrows-insertlocation--rowcount--values-)|Fügt am Anfang oder Ende der Tabelle Zeilen hinzu, wobei die erste oder letzte vorhandene Zeile als Vorlage verwendet wird.|
||[Ausrichtung](/javascript/api/word/word.table#alignment)|Ruft die Ausrichtung der Tabelle mit der Seitenspalte ab oder legt sie fest.|
||[autoFitWindow()](/javascript/api/word/word.table#autofitwindow--)|Passt die Tabellenspalten automatisch an die Breite des Fensters an.|
||[clear()](/javascript/api/word/word.table#clear--)|Löscht den Inhalt der Tabelle.|
||[delete()](/javascript/api/word/word.table#delete--)|Löscht die gesamte Tabelle.|
||[deleteColumns(columnIndex: number, columnCount?: number)](/javascript/api/word/word.table#deletecolumns-columnindex--columncount-)|Löscht bestimmte Spalten.|
||[deleteRows(rowIndex: number, rowCount?: number)](/javascript/api/word/word.table#deleterows-rowindex--rowcount-)|Löscht bestimmte Zeilen.|
||[distributeColumns()](/javascript/api/word/word.table#distributecolumns--)|Verteilt die Breite der Spalten gleichmäßig.|
||[getBorder(borderLocation: Word.BorderLocation)](/javascript/api/word/word.table#getborder-borderlocation-)|Ruft die Rahmenart für den angegebenen Rahmen ab.|
||[getCell(rowIndex: number, cellIndex: number)](/javascript/api/word/word.table#getcell-rowindex--cellindex-)|Ruft die Tabellenzelle in einer angegebenen Zeile und Spalte ab.|
||[getCellOrNullObject(rowIndex: number, cellIndex: number)](/javascript/api/word/word.table#getcellornullobject-rowindex--cellindex-)|Ruft die Tabellenzelle in einer angegebenen Zeile und Spalte ab.|
||[getCellPadding(cellPaddingLocation: Word.CellPaddingLocation)](/javascript/api/word/word.table#getcellpadding-cellpaddinglocation-)|Ruft den Textabstand in Punkt ab.|
||[getNext()](/javascript/api/word/word.table#getnext--)|Ruft die nächste Tabelle ab.|
||[getNextOrNullObject()](/javascript/api/word/word.table#getnextornullobject--)|Ruft die nächste Tabelle ab.|
||[getParagraphAfter()](/javascript/api/word/word.table#getparagraphafter--)|Ruft den Absatz nach der Tabelle ab.|
||[getParagraphAfterOrNullObject()](/javascript/api/word/word.table#getparagraphafterornullobject--)|Ruft den Absatz nach der Tabelle ab.|
||[getParagraphBefore()](/javascript/api/word/word.table#getparagraphbefore--)|Ruft den Absatz vor der Tabelle ab.|
||[getParagraphBeforeOrNullObject()](/javascript/api/word/word.table#getparagraphbeforeornullobject--)|Ruft den Absatz vor der Tabelle ab.|
||[getRange(rangeLocation?: Word.RangeLocation)](/javascript/api/word/word.table#getrange-rangelocation-)|Ruft den Bereich ab, der diese Tabelle enthält, oder den Bereich am Anfang bzw. Ende der Tabelle.|
||[headerRowCount](/javascript/api/word/word.table#headerrowcount)|Ruft die Anzahl von Kopfzeilen ab oder legt sie fest.|
||[horizontalAlignment](/javascript/api/word/word.table#horizontalalignment)|Ruft die horizontale Ausrichtung jeder Zelle in der Tabelle ab und legt sie fest.|
||[ignorePunct](/javascript/api/word/word.table#ignorepunct)||
||[ignoreSpace](/javascript/api/word/word.table#ignorespace)||
||[insertContentControl()](/javascript/api/word/word.table#insertcontentcontrol--)|Fügt ein Inhaltssteuerelement in der Tabelle ein.|
||[insertParagraph(paragraphText: string, insertLocation: Word.InsertLocation)](/javascript/api/word/word.table#insertparagraph-paragraphtext--insertlocation-)|Fügt an der angegebenen Position einen Absatz ein.|
||[insertTable(rowCount: number, columnCount: number, insertLocation: Word.InsertLocation, values?: string[][])](/javascript/api/word/word.table#inserttable-rowcount--columncount--insertlocation--values-)|Fügt eine Tabelle mit der angegebenen Anzahl von Zeilen und Spalten ein.|
||[matchCase](/javascript/api/word/word.table#matchcase)||
||[matchPrefix](/javascript/api/word/word.table#matchprefix)||
||[matchSuffix](/javascript/api/word/word.table#matchsuffix)||
||[matchWholeWord](/javascript/api/word/word.table#matchwholeword)||
||[matchWildcards](/javascript/api/word/word.table#matchwildcards)||
||[font](/javascript/api/word/word.table#font)|Ruft die Schriftart ab.|
||[isUniform](/javascript/api/word/word.table#isuniform)|Gibt an, ob alle Zeilen der Tabelle uniform sind.|
||[nestingLevel](/javascript/api/word/word.table#nestinglevel)|Ruft die Schachtelungsebene der Tabelle ab.|
||[parentBody](/javascript/api/word/word.table#parentbody)|Ruft den übergeordneten Text der Tabelle ab.|
||[parentContentControl](/javascript/api/word/word.table#parentcontentcontrol)|Ruft das Inhaltssteuerelement ab, das die Tabelle enthält.|
||[parentContentControlOrNullObject](/javascript/api/word/word.table#parentcontentcontrolornullobject)|Ruft das Inhaltssteuerelement ab, das die Tabelle enthält.|
||[parentTable](/javascript/api/word/word.table#parenttable)|Ruft die Tabelle ab, die diese Tabelle enthält.|
||[parentTableCell](/javascript/api/word/word.table#parenttablecell)|Ruft die Tabellenzelle ab, die diese Tabelle enthält.|
||[parentTableCellOrNullObject](/javascript/api/word/word.table#parenttablecellornullobject)|Ruft die Tabellenzelle ab, die diese Tabelle enthält.|
||[parentTableOrNullObject](/javascript/api/word/word.table#parenttableornullobject)|Ruft die Tabelle ab, die diese Tabelle enthält.|
||[rowCount](/javascript/api/word/word.table#rowcount)|Ruft die Anzahl der Zeilen in der Tabelle ab.|
||[rows](/javascript/api/word/word.table#rows)|Ruft alle Zeilen der Tabelle ab.|
||[Tabellen](/javascript/api/word/word.table#tables)|Ruft die untergeordneten Tabellen ab, die eine Ebene tiefer verschachtelt sind.|
||[search(searchText: string, searchOptions?: Word.SearchOptions \| {
            ignorePunct?: boolean
            ignoreSpace?: boolean
            matchCase?: boolean
            matchPrefix?: boolean
            matchSuffix?: boolean
            matchWholeWord?: boolean
            matchWildcards?: boolean
        })](/javascript/api/word/word.table#search-searchtext--searchoptions--ignorepunct--ignorespace--matchcase--matchprefix--matchsuffix--matchwholeword--matchwildcards-)|Performs a search with the specified SearchOptions on the scope of the table object.|
||[select(selectionMode?: Word.SelectionMode)](/javascript/api/word/word.table#select-selectionmode-)| Wählt die Tabelle oder die Position am Anfang oder Ende der Tabelle aus und navigiert über die Benutzeroberfläche von Word zu ihr.| || [setCellPadding(cellPaddingLocation: Word.CellPaddingLocation, cellPadding: number)](/javascript/api/word/word.table#setcellpadding-cellpaddinglocation--cellpadding-)| Legt den Zellabstand in punkt.| || [shadingColor |](/javascript/api/word/word.table#shadingcolor) Ruft die Schattierungsfarbe ab und legt | || [Formatvorlage](/javascript/api/word/word.table#style)| Ruft den Formatvorlagenamen für die Tabelle ab, oder legt | || [styleBandedColumns](/javascript/api/word/word.table#stylebandedcolumns)| Ruft ab und legt fest, ob die Tabelle spalten.| || [styleBandedRows](/javascript/api/word/word.table#stylebandedrows)| Ruft ab und legt fest, ob die Tabelle Überbandzeilen | || [styleBuiltIn](/javascript/api/word/word.table#stylebuiltin)| Ruft den integrierten Formatvorlagenamen für die Tabelle ab, oder legt | || [styleFirstColumn](/javascript/api/word/word.table#stylefirstcolumn)| Ruft ab und legt fest, ob die Tabelle eine erste Spalte mit einer speziellen Formatvorlage | || [styleLastColumn](/javascript/api/word/word.table#stylelastcolumn)| Ruft ab und legt fest, ob die Tabelle eine letzte Spalte mit einer speziellen Formatvorlage | || [styleTotalRow](/javascript/api/word/word.table#styletotalrow)| Ruft ab und legt fest, ob die Tabelle eine Gesamtzeile (letzte) mit einer speziellen Formatvorlage | || [Werte](/javascript/api/word/word.table#values)| Ruft die Textwerte in der Tabelle als 2D-Javascript-Array ab und legt | || [verticalAlignment](/javascript/api/word/word.table#verticalalignment)| Ruft die vertikale Ausrichtung jeder Zelle in der Tabelle ab und legt sie | || [width](/javascript/api/word/word.table#width)| Ruft die Breite der Tabelle in points.| | [TableBorder](/javascript/api/word/word.tableborder) | [farblich](/javascript/api/word/word.tableborder#color)| Ruft die Tabellenrandfarbe ab oder legt sie fest. | || [Typ](/javascript/api/word/word.tableborder#type)| Ruft den Typ des Tabellenrands ab oder legt diesen | || [width](/javascript/api/word/word.tableborder#width)| Ruft die Breite des Tabellenrands in Punkt ab, oder legt diese fest.| | [TableCell](/javascript/api/word/word.tablecell) | [columnWidth](/javascript/api/word/word.tablecell#columnwidth)| Ruft die Breite der Spalte der Zelle in points.| || [deleteColumn()](/javascript/api/word/word.tablecell#deletecolumn--)| Löscht die Spalte mit dieser Zelle.| || [deleteRow()](/javascript/api/word/word.tablecell#deleterow--)| Löscht die Zeile mit dieser Zelle.| || [getBorder(borderLocation: Word.BorderLocation)](/javascript/api/word/word.tablecell#getborder-borderlocation-)| Ruft die Rahmenformatvorlage für den angegebenen Rahmen ab.| || [getCellPadding(cellPaddingLocation: Word.CellPaddingLocation)](/javascript/api/word/word.tablecell#getcellpadding-cellpaddinglocation-)| Ruft zellenabstand in points.| || [getNext()](/javascript/api/word/word.tablecell#getnext--)| Ruft die nächste Zelle ab.| || [getNextOrNullObject()](/javascript/api/word/word.tablecell#getnextornullobject--)| Ruft die nächste Zelle ab.| || [horizontalAlignment](/javascript/api/word/word.tablecell#horizontalalignment)| Ruft die horizontale Ausrichtung der Zelle ab und legt | || [insertColumns(insertLocation: Word.InsertLocation, columnCount: number, values?: string[][])](/javascript/api/word/word.tablecell#insertcolumns-insertlocation--columncount--values-)| Fügt links oder rechts der Zelle Spalten hinzu, indem die Spalte der Zelle als Vorlage verwendet | || [insertRows(insertLocation: Word.InsertLocation, rowCount: number, values?: string[][])](/javascript/api/word/word.tablecell#insertrows-insertlocation--rowcount--values-)| Fügt Zeilen oberhalb oder unterhalb der Zelle ein, indem die Zeile der Zelle als Vorlage verwendet | || [body](/javascript/api/word/word.tablecell#body)| Ruft das body-Objekt der Zelle ab.| || [cellIndex](/javascript/api/word/word.tablecell#cellindex)| Ruft den Index der Zelle in ihrer Zeile ab.| || [parentRow-|](/javascript/api/word/word.tablecell#parentrow) Ruft die übergeordnete Zeile der Zelle ab.| || [parentTable-|](/javascript/api/word/word.tablecell#parenttable) Ruft die übergeordnete Tabelle der Zelle ab.| || [rowIndex](/javascript/api/word/word.tablecell#rowindex)| Ruft den Index der Zeile der Zelle in der Tabelle ab.| || [width](/javascript/api/word/word.tablecell#width)| Ruft die Breite der Zelle in points.| || [setCellPadding(cellPaddingLocation: Word.CellPaddingLocation, cellPadding: number)](/javascript/api/word/word.tablecell#setcellpadding-cellpaddinglocation--cellpadding-)| Legt den Zellabstand in punkt.| || [shadingColor |](/javascript/api/word/word.tablecell#shadingcolor) Ruft die Schattierungsfarbe der Zelle ab, oder legt sie fest.| || [Wert](/javascript/api/word/word.tablecell#value)| Ruft den Text der Zelle ab und legt | || [verticalAlignment](/javascript/api/word/word.tablecell#verticalalignment)| Ruft die vertikale Ausrichtung der Zelle ab und legt | | [TableCellCollection](/javascript/api/word/word.tablecellcollection) | [getFirst()](/javascript/api/word/word.tablecellcollection#getfirst--)| Ruft die erste Tabellenzelle in dieser Auflistung ab.| || [getFirstOrNullObject()](/javascript/api/word/word.tablecellcollection#getfirstornullobject--)| Ruft die erste Tabellenzelle in dieser Auflistung ab.| || [Elemente](/javascript/api/word/word.tablecellcollection#items)| Ruft die geladenen untergeordneten Elemente in dieser Auflistung ab.| | [TableCollection](/javascript/api/word/word.tablecollection) | [getFirst()](/javascript/api/word/word.tablecollection#getfirst--)| Ruft die erste Tabelle in dieser Auflistung ab.| || [getFirstOrNullObject()](/javascript/api/word/word.tablecollection#getfirstornullobject--)| Ruft die erste Tabelle in dieser Auflistung ab.| || [Elemente](/javascript/api/word/word.tablecollection#items)| Ruft die geladenen untergeordneten Elemente in dieser Auflistung ab.| | [TableRow](/javascript/api/word/word.tablerow) | [clear()](/javascript/api/word/word.tablerow#clear--)| Der Inhalt der Zeile wird | || [delete()](/javascript/api/word/word.tablerow#delete--)| Löscht die gesamte Zeile.| || [getBorder(borderLocation: Word.BorderLocation)](/javascript/api/word/word.tablerow#getborder-borderlocation-)| Ruft die Rahmenformatvorlage der Zellen in der Zeile ab.| || [getCellPadding(cellPaddingLocation: Word.CellPaddingLocation)](/javascript/api/word/word.tablerow#getcellpadding-cellpaddinglocation-)| Ruft zellenabstand in points.| || [getNext()](/javascript/api/word/word.tablerow#getnext--)| Ruft die nächste Zeile ab.| || [getNextOrNullObject()](/javascript/api/word/word.tablerow#getnextornullobject--)| Ruft die nächste Zeile ab.| || [horizontalAlignment](/javascript/api/word/word.tablerow#horizontalalignment)| Ruft die horizontale Ausrichtung jeder Zelle in der Zeile ab und legt | || [ignorePunct](/javascript/api/word/word.tablerow#ignorepunct)|| || [ignoreSpace](/javascript/api/word/word.tablerow#ignorespace)|| || [insertRows(insertLocation: Word.InsertLocation, rowCount: number, values?: string[][])](/javascript/api/word/word.tablerow#insertrows-insertlocation--rowcount--values-)| Fügt Zeilen ein, die diese Zeile als Vorlage verwenden.| || [matchCase](/javascript/api/word/word.tablerow#matchcase)|| || [matchPrefix](/javascript/api/word/word.tablerow#matchprefix)|| || [matchSuffix](/javascript/api/word/word.tablerow#matchsuffix)|| || [matchWholeWord](/javascript/api/word/word.tablerow#matchwholeword)|| || [matchWildcards](/javascript/api/word/word.tablerow#matchwildcards)|| || [preferredHeight](/javascript/api/word/word.tablerow#preferredheight)| Ruft die bevorzugte Höhe der Zeile in punkt.| || [cellCount](/javascript/api/word/word.tablerow#cellcount)| Ruft die Anzahl der Zellen in der Zeile ab.| || [Zellen |](/javascript/api/word/word.tablerow#cells) Ruft cells.| || [schriftart](/javascript/api/word/word.tablerow#font)| Ruft die schriftart.| || [isHeader](/javascript/api/word/word.tablerow#isheader)| Überprüft, ob es sich bei der Zeile um eine Kopfzeile | || [parentTable-|](/javascript/api/word/word.tablerow#parenttable) Ruft übergeordnete Tabelle ab.| || [rowIndex](/javascript/api/word/word.tablerow#rowindex)| Ruft den Index der Zeile in der übergeordneten Tabelle ab.| || [search(searchText: string, searchOptions?: Word.SearchOptions \| { ignorePunct?: boolean ignoreSpace?: boolean matchCase?: boolean matchPrefix?: boolean matchSuffix?: boolean matchWholeWord?: boolean matchWildcards?: boolean })](/javascript/api/word/word.tablerow#search-searchtext--searchoptions--ignorepunct--ignorespace--matchcase--matchprefix--matchsuffix--matchwholeword--matchwildcards-)| Führt eine Suche mit den angegebenen SearchOptions im Bereich der Zeile aus.| || [select(selectionMode?: Word.SelectionMode)](/javascript/api/word/word.tablerow#select-selectionmode-)| Wählt die Zeile aus und navigiert die Word-Benutzeroberfläche zu ihr.| || [setCellPadding(cellPaddingLocation: Word.CellPaddingLocation, cellPadding: number)](/javascript/api/word/word.tablerow#setcellpadding-cellpaddinglocation--cellpadding-)| Legt den Zellabstand in punkt.| || [shadingColor |](/javascript/api/word/word.tablerow#shadingcolor) Ruft die Schattierungsfarbe ab und legt | || [Werte](/javascript/api/word/word.tablerow#values)| Ruft die Textwerte in der Zeile als 2D-Javascript-Array ab und legt | || [verticalAlignment](/javascript/api/word/word.tablerow#verticalalignment)| Ruft die vertikale Ausrichtung der Zellen in der Zeile ab und legt sie | | [TableRowCollection](/javascript/api/word/word.tablerowcollection) | [getFirst()](/javascript/api/word/word.tablerowcollection#getfirst--)| Ruft die erste Zeile in dieser Auflistung ab.| || [getFirstOrNullObject()](/javascript/api/word/word.tablerowcollection#getfirstornullobject--)| Ruft die erste Zeile in dieser Auflistung ab.| || [Elemente](/javascript/api/word/word.tablerowcollection#items)| Ruft die geladenen untergeordneten Elemente in dieser Auflistung ab.|
