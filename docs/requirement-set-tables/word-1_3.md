| Klasse | Felder | Beschreibung |
|:---|:---|:---|
|[Anwendung](/javascript/api/word/word.application)|[createDocument (Base64-Datei?: String)](/javascript/api/word/word.application#createdocument-base64file-)|Erstellt ein neues Dokument mithilfe einer optionalen Base64-codierten docx-Datei.|
|[Body](/javascript/api/word/word.body)|[GetRange (rangeLocation?: Word. rangeLocation)](/javascript/api/word/word.body#getrange-rangelocation-)|Ruft den gesamten Text oder den Start- bzw. Endpunkt des Textkörpers als Bereich ab.|
||[Insertable (ROWCOUNT: Number, ColumnCount: Number, insertLocation: Word. insertLocation, Values?: String [] [])](/javascript/api/word/word.body#inserttable-rowcount--columncount--insertlocation--values-)|Fügt eine Tabelle mit der angegebenen Anzahl von Zeilen und Spalten ein.|
||[Listen](/javascript/api/word/word.body#lists)|Ruft die Sammlung von Listenobjekten im Textkörper ab.|
||[parentBody](/javascript/api/word/word.body#parentbody)|Ruft den übergeordneten Text des Texts ab.|
||[parentBodyOrNullObject](/javascript/api/word/word.body#parentbodyornullobject)|Ruft den übergeordneten Text des Texts ab.|
||[parentContentControlOrNullObject](/javascript/api/word/word.body#parentcontentcontrolornullobject)|Ruft das Inhaltssteuerelement ab, das den Textkörper enthält.|
||[parentSection](/javascript/api/word/word.body#parentsection)|Ruft den übergeordneten Abschnitt des Texts ab.|
||[parentSectionOrNullObject](/javascript/api/word/word.body#parentsectionornullobject)|Ruft den übergeordneten Abschnitt des Texts ab.|
||[Tabellen](/javascript/api/word/word.body#tables)|Ruft die Sammlung von Tabellenobjekten im Text ab.|
||[Typ](/javascript/api/word/word.body#type)|Ruft den Typ des Texts ab.|
||[styleBuiltIn](/javascript/api/word/word.body#stylebuiltin)|Ruft den Namen der integrierten Formatvorlage für den Text ab oder legt ihn fest.|
|[ContentControl](/javascript/api/word/word.contentcontrol)|[GetRange (rangeLocation?: Word. rangeLocation)](/javascript/api/word/word.contentcontrol#getrange-rangelocation-)|Ruft das gesamte Inhaltssteuerelement oder den Start- bzw. Endpunkt des Inhaltssteuerelements als Bereich ab.|
||[getTextRanges (endingMarks: String [], trimSpacing?: Boolean)](/javascript/api/word/word.contentcontrol#gettextranges-endingmarks--trimspacing-)|Ruft die Textbereiche im Inhaltssteuerelement mithilfe von Satzzeichen und/oder anderen Endmarkierungen ab.|
||[Insertable (ROWCOUNT: Number, ColumnCount: Number, insertLocation: Word. insertLocation, Values?: String [] [])](/javascript/api/word/word.contentcontrol#inserttable-rowcount--columncount--insertlocation--values-)|Fügt eine Tabelle mit der angegebenen Anzahl von Zeilen und Spalten in oder neben ein Inhaltssteuerelement ein.|
||[Listen](/javascript/api/word/word.contentcontrol#lists)|Ruft die Sammlung von Listenobjekten im Inhaltssteuerelement ab.|
||[parentBody](/javascript/api/word/word.contentcontrol#parentbody)|Ruft den übergeordneten Text des Inhaltssteuerelements ab.|
||[parentContentControlOrNullObject](/javascript/api/word/word.contentcontrol#parentcontentcontrolornullobject)|Ruft das Inhaltssteuerelement ab, das das Inhaltssteuerelement enthält.|
||[parentTable](/javascript/api/word/word.contentcontrol#parenttable)|Ruft die Tabelle ab, die das Inhaltssteuerelement enthält.|
||[parentTableCell](/javascript/api/word/word.contentcontrol#parenttablecell)|Ruft die Tabellenzelle ab, die das Inhaltssteuerelement enthält.|
||[parentTableCellOrNullObject](/javascript/api/word/word.contentcontrol#parenttablecellornullobject)|Ruft die Tabellenzelle ab, die das Inhaltssteuerelement enthält.|
||[parentTableOrNullObject](/javascript/api/word/word.contentcontrol#parenttableornullobject)|Ruft die Tabelle ab, die das Inhaltssteuerelement enthält.|
||[Untertyp](/javascript/api/word/word.contentcontrol#subtype)|Ruft den Untertyp des Inhaltssteuerelements ab.|
||[Tabellen](/javascript/api/word/word.contentcontrol#tables)|Ruft die Sammlung von Tabellenobjekten im Inhaltssteuerelement ab.|
||[Split (Trennzeichen: String [], multiparagraphs?: Boolean, trimDelimiters?: Boolean, trimSpacing?: Boolean)](/javascript/api/word/word.contentcontrol#split-delimiters--multiparagraphs--trimdelimiters--trimspacing-)|Teilt das Inhaltssteuerelement mithilfe von Trennzeichen in untergeordnete Bereiche.|
||[styleBuiltIn](/javascript/api/word/word.contentcontrol#stylebuiltin)|Ruft den Namen der integrierten Formatvorlage für das Inhaltssteuerelement ab oder legt ihn fest.|
|[ContentControlCollection](/javascript/api/word/word.contentcontrolcollection)|[getByIdOrNullObject (ID: Number)](/javascript/api/word/word.contentcontrolcollection#getbyidornullobject-id-)|Ruft ein Inhaltssteuerelement mithilfe der ID ab.|
||[getByTypes (Types: Word. contentcontroltype [])](/javascript/api/word/word.contentcontrolcollection#getbytypes-types-)|Ruft die Inhaltssteuerelemente mit den angegebenen Typen und/oder Untertypen ab.|
||[getFirst()](/javascript/api/word/word.contentcontrolcollection#getfirst--)|Ruft das erste Inhaltssteuerelement in dieser Sammlung ab.|
||[getFirstOrNullObject()](/javascript/api/word/word.contentcontrolcollection#getfirstornullobject--)|Ruft das erste Inhaltssteuerelement in dieser Sammlung ab.|
|[CustomProperty](/javascript/api/word/word.customproperty)|[delete()](/javascript/api/word/word.customproperty#delete--)|Löscht die benutzerdefinierte Eigenschaft.|
||[key](/javascript/api/word/word.customproperty#key)|Ruft den Schlüssel der benutzerdefinierten Eigenschaft ab.|
||[Typ](/javascript/api/word/word.customproperty#type)|Ruft den Wertetyp der benutzerdefinierten Eigenschaft ab.|
||[value](/javascript/api/word/word.customproperty#value)|Ruft den Wert der benutzerdefinierten Eigenschaft ab oder legt ihn fest.|
|[CustomPropertyCollection](/javascript/api/word/word.custompropertycollection)|[Add (Key: String, Value: any)](/javascript/api/word/word.custompropertycollection#add-key--value-)|Erstellt eine neue benutzerdefinierte Eigenschaft oder legt eine vorhandene fest.|
||[DeleteAll ()](/javascript/api/word/word.custompropertycollection#deleteall--)|Löscht alle benutzerdefinierten Eigenschaften in dieser Sammlung.|
||[getCount()](/javascript/api/word/word.custompropertycollection#getcount--)|Ruft die Anzahl von benutzerdefinierten Eigenschaften ab.|
||[getItem(key: string)](/javascript/api/word/word.custompropertycollection#getitem-key-)|Ruft ein Objekt für benutzerdefinierte Eigenschaften über seinen Schlüssel ab, bei dem Groß-/Kleinschreibung nicht beachtet wird.|
||[getItemOrNullObject(key: string)](/javascript/api/word/word.custompropertycollection#getitemornullobject-key-)|Ruft ein Objekt für benutzerdefinierte Eigenschaften über seinen Schlüssel ab, bei dem Groß-/Kleinschreibung nicht beachtet wird.|
||[items](/javascript/api/word/word.custompropertycollection#items)|Ruft die geladenen untergeordneten Elemente in dieser Sammlung ab.|
|[Document](/javascript/api/word/word.document)|[properties](/javascript/api/word/word.document#properties)|Ruft die Eigenschaften des Dokuments ab.|
|[DocumentCreated](/javascript/api/word/word.documentcreated)|[Open ()](/javascript/api/word/word.documentcreated#open--)|Öffnet das Dokument.|
||[Text](/javascript/api/word/word.documentcreated#body)|Ruft das Body-Objekt des Dokuments ab.|
||[contentControls](/javascript/api/word/word.documentcreated#contentcontrols)|Ruft die Auflistung der Inhaltssteuerelement-Objekte im Dokument ab.|
||[properties](/javascript/api/word/word.documentcreated#properties)|Ruft die Eigenschaften des Dokuments ab.|
||[gespeichert](/javascript/api/word/word.documentcreated#saved)|Gibt an, ob die Änderungen im Dokument gespeichert wurden.|
||[sections](/javascript/api/word/word.documentcreated#sections)|Ruft die Auflistung von Section-Objekten im Dokument ab.|
||[save()](/javascript/api/word/word.documentcreated#save--)|Speichert das Dokument.|
|[DocumentProperties](/javascript/api/word/word.documentproperties)|[Autor](/javascript/api/word/word.documentproperties#author)|Ruft den Autor des Dokuments ab oder legt ihn fest.|
||[Kategorie](/javascript/api/word/word.documentproperties#category)|Ruft die Kategorie des Dokuments ab oder legt sie fest.|
||[comments](/javascript/api/word/word.documentproperties#comments)|Ruft die Kommentare des Dokuments ab oder legt sie fest.|
||[company](/javascript/api/word/word.documentproperties#company)|Ruft das Unternehmen des Dokuments ab oder legt es fest.|
||[format](/javascript/api/word/word.documentproperties#format)|Ruft das Format des Dokuments ab oder legt es fest.|
||[Schlüsselwörter](/javascript/api/word/word.documentproperties#keywords)|Ruft die Schlüsselwörter des Dokuments ab oder legt sie fest.|
||[manager](/javascript/api/word/word.documentproperties#manager)|Ruft den Manager des Dokuments ab oder legt ihn fest.|
||[applicationName](/javascript/api/word/word.documentproperties#applicationname)|Ruft den Anwendungsnamen des Dokuments ab.|
||[creationDate](/javascript/api/word/word.documentproperties#creationdate)|Ruft das Erstellungsdatum des Dokuments ab.|
||[customProperties](/javascript/api/word/word.documentproperties#customproperties)|Ruft die Sammlung der benutzerdefinierten Eigenschaften des Dokuments ab.|
||[lastAuthor](/javascript/api/word/word.documentproperties#lastauthor)|Ruft den letzten Autor des Dokuments ab.|
||[Last Print date](/javascript/api/word/word.documentproperties#lastprintdate)|Ruft das letzte Druckdatum des Dokuments ab.|
||[lastSaveTime](/javascript/api/word/word.documentproperties#lastsavetime)|Ruft die Zeit der letzten Speicherung des Dokuments ab.|
||[revisionNumber](/javascript/api/word/word.documentproperties#revisionnumber)|Ruft die Revisionsnummer des Dokuments ab.|
||[Sicherheits](/javascript/api/word/word.documentproperties#security)|Ruft Sicherheitseinstellungen des Dokuments ab.|
||[template](/javascript/api/word/word.documentproperties#template)|Ruft die Vorlage des Dokuments ab.|
||[Betreff](/javascript/api/word/word.documentproperties#subject)|Ruft den Betreff des Dokuments ab oder legt ihn fest.|
||[title](/javascript/api/word/word.documentproperties#title)|Ruft den Titel des Dokuments ab oder legt ihn fest.|
|[InlinePicture](/javascript/api/word/word.inlinepicture)|[GetNext ()](/javascript/api/word/word.inlinepicture#getnext--)|Ruft das nächste Inlinebild ab.|
||[getNextOrNullObject()](/javascript/api/word/word.inlinepicture#getnextornullobject--)|Ruft das nächste Inlinebild ab.|
||[GetRange (rangeLocation?: Word. rangeLocation)](/javascript/api/word/word.inlinepicture#getrange-rangelocation-)|Ruft das Bild oder den Start- bzw. Endpunkt des Bilds als Bereich ab.|
||[parentContentControlOrNullObject](/javascript/api/word/word.inlinepicture#parentcontentcontrolornullobject)|Ruft das Inhaltssteuerelement ab, das das eingebundene Bild enthält.|
||[parentTable](/javascript/api/word/word.inlinepicture#parenttable)|Ruft die Tabelle ab, die das Inlinebild enthält.|
||[parentTableCell](/javascript/api/word/word.inlinepicture#parenttablecell)|Ruft die Tabellenzelle ab, die das Inlinebild enthält.|
||[parentTableCellOrNullObject](/javascript/api/word/word.inlinepicture#parenttablecellornullobject)|Ruft die Tabellenzelle ab, die das Inlinebild enthält.|
||[parentTableOrNullObject](/javascript/api/word/word.inlinepicture#parenttableornullobject)|Ruft die Tabelle ab, die das Inlinebild enthält.|
|[InlinePictureCollection](/javascript/api/word/word.inlinepicturecollection)|[getFirst()](/javascript/api/word/word.inlinepicturecollection#getfirst--)|Ruft das erste Inlinebild in dieser Sammlung ab.|
||[getFirstOrNullObject()](/javascript/api/word/word.inlinepicturecollection#getfirstornullobject--)|Ruft das erste Inlinebild in dieser Sammlung ab.|
|[List](/javascript/api/word/word.list)|[getLevelParagraphs (Level: Number)](/javascript/api/word/word.list#getlevelparagraphs-level-)|Ruft die Absätze auf der angegebenen Ebene in der Liste ab.|
||[getlevelling (Level: Number)](/javascript/api/word/word.list#getlevelstring-level-)|Ruft das Aufzählungszeichen, die Zahl oder das Bild auf der angegebenen Ebene als Zeichenfolge ab.|
||[insertParagraph (ParagraphText: String, insertLocation: Word. insertLocation)](/javascript/api/word/word.list#insertparagraph-paragraphtext--insertlocation-)|Fügt an der angegebenen Position einen Absatz ein.|
||[id](/javascript/api/word/word.list#id)|Ruft die ID der Liste ab.|
||[levelExistences](/javascript/api/word/word.list#levelexistences)|Überprüft, ob jede der 9 Ebenen in der Liste vorhanden ist.|
||[levelTypes](/javascript/api/word/word.list#leveltypes)|Ruft die Typen aller 9 Ebenen in der Liste ab.|
||[paragraphs](/javascript/api/word/word.list#paragraphs)|Ruft die Absätze in der Liste ab.|
||[setLevelAlignment (Level: Number, Alignment: Word. Alignment)](/javascript/api/word/word.list#setlevelalignment-level--alignment-)|Legt die Ausrichtung des Aufzählungszeichens, der Zahl oder der Grafik auf der angegebenen Ebene in der Liste fest.|
||[setLevelBullet (Level: Number, listBullet: Word. listBullet, CharCode?: Number, FontName?: String)](/javascript/api/word/word.list#setlevelbullet-level--listbullet--charcode--fontname-)|Legt das Format des Aufzählungszeichens auf der angegebenen Ebene in der Liste fest.|
||[setLevelIndents (Level: Number, textIndent: Number, bulletNumberPictureIndent: Number)](/javascript/api/word/word.list#setlevelindents-level--textindent--bulletnumberpictureindent-)|Legt die zwei Einzüge der angegebenen Ebene in der Liste fest.|
||[setLevelNumbering (Level: Number, listNumbering: Word. listNumbering, Format String?: Array<String \| Number>)](/javascript/api/word/word.list#setlevelnumbering-level--listnumbering--formatstring-)|Legt das Nummerierungsformat auf der angegebenen Ebene in der Liste fest.|
||[setLevelStartingNumber (Level: Number, startingNumber: Number)](/javascript/api/word/word.list#setlevelstartingnumber-level--startingnumber-)|Legt die Anfangsnummer auf der angegebenen Ebene in der Liste fest.|
|[listCollection](/javascript/api/word/word.listcollection)|[getById(id: number)](/javascript/api/word/word.listcollection#getbyid-id-)|Ruft eine Liste über ihren Bezeichner ab.|
||[getByIdOrNullObject (ID: Number)](/javascript/api/word/word.listcollection#getbyidornullobject-id-)|Ruft eine Liste über ihren Bezeichner ab.|
||[getFirst()](/javascript/api/word/word.listcollection#getfirst--)|Ruft die erste Liste in dieser Sammlung ab.|
||[getFirstOrNullObject()](/javascript/api/word/word.listcollection#getfirstornullobject--)|Ruft die erste Liste in dieser Sammlung ab.|
||[getItem(index: number)](/javascript/api/word/word.listcollection#getitem-index-)|Ruft ein Listenobjekt nach dem Index in der Sammlung ab.|
||[items](/javascript/api/word/word.listcollection#items)|Ruft die geladenen untergeordneten Elemente in dieser Sammlung ab.|
|[ListItem](/javascript/api/word/word.listitem)|[GetAncestor (parentOnly?: Boolean)](/javascript/api/word/word.listitem#getancestor-parentonly-)|Ruft das übergeordnete Element der Liste bzw. den nächsten Vorgänger ab, wenn das übergeordnete Element nicht vorhanden ist.|
||[getAncestorOrNullObject (parentOnly?: Boolean)](/javascript/api/word/word.listitem#getancestorornullobject-parentonly-)|Ruft das übergeordnete Element der Liste bzw. den nächsten Vorgänger ab, wenn das übergeordnete Element nicht vorhanden ist.|
||[getdescendants (directChildrenOnly?: Boolean)](/javascript/api/word/word.listitem#getdescendants-directchildrenonly-)|Ruft alle untergeordneten Listenelemente des Listenelements ab.|
||[level](/javascript/api/word/word.listitem#level)|Ruft die Ebene des Elements in der Liste ab oder legt sie fest.|
||[ListString](/javascript/api/word/word.listitem#liststring)|Ruft das Listenelement Aufzählungszeichen, Zahl oder Grafik als Zeichenfolge ab.|
||[siblingIndex](/javascript/api/word/word.listitem#siblingindex)|Ruft die Ordnungsnummer des Listenelements im Verhältnis zu seinen gleichgeordneten Elementen ab.|
|[Paragraph](/javascript/api/word/word.paragraph)|[attachToList (Listen-ID: Number, Level: Number)](/javascript/api/word/word.paragraph#attachtolist-listid--level-)|Fügt den Absatz zu einer vorhandenen Liste auf der angegebenen Ebene hinzu.|
||[detachFromList()](/javascript/api/word/word.paragraph#detachfromlist--)|Verschiebt diesen Absatz aus der Liste, falls der Absatz ein Listenelement ist.|
||[GetNext ()](/javascript/api/word/word.paragraph#getnext--)|Ruft den nächsten Absatz ab.|
||[getNextOrNullObject()](/javascript/api/word/word.paragraph#getnextornullobject--)|Ruft den nächsten Absatz ab.|
||[GetPrevious ()](/javascript/api/word/word.paragraph#getprevious--)|Ruft den vorherigen Absatz ab.|
||[getPreviousOrNullObject()](/javascript/api/word/word.paragraph#getpreviousornullobject--)|Ruft den vorherigen Absatz ab.|
||[GetRange (rangeLocation?: Word. rangeLocation)](/javascript/api/word/word.paragraph#getrange-rangelocation-)|Ruft den gesamten Absatz oder den Start- bzw. Endpunkt des Absatzes als Bereich ab.|
||[getTextRanges (endingMarks: String [], trimSpacing?: Boolean)](/javascript/api/word/word.paragraph#gettextranges-endingmarks--trimspacing-)|Ruft die Textbereiche im Absatz mithilfe von Satzzeichen und/oder anderen Endmarkierungen ab.|
||[Insertable (ROWCOUNT: Number, ColumnCount: Number, insertLocation: Word. insertLocation, Values?: String [] [])](/javascript/api/word/word.paragraph#inserttable-rowcount--columncount--insertlocation--values-)|Fügt eine Tabelle mit der angegebenen Anzahl von Zeilen und Spalten ein.|
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
||[Split (Trennzeichen: String [], trimDelimiters?: Boolean, trimSpacing?: Boolean)](/javascript/api/word/word.paragraph#split-delimiters--trimdelimiters--trimspacing-)|Teilt den Absatz mithilfe von Trennzeichen in untergeordnete Bereiche.|
||[startNewList()](/javascript/api/word/word.paragraph#startnewlist--)|Beginnt eine neue Liste mit diesem Absatz.|
||[styleBuiltIn](/javascript/api/word/word.paragraph#stylebuiltin)|Ruft den Namen der integrierten Formatvorlage für den Absatz ab oder legt ihn fest.|
|[ParagraphCollection](/javascript/api/word/word.paragraphcollection)|[getFirst()](/javascript/api/word/word.paragraphcollection#getfirst--)|Ruft den ersten Absatz in dieser Sammlung ab.|
||[getFirstOrNullObject()](/javascript/api/word/word.paragraphcollection#getfirstornullobject--)|Ruft den ersten Absatz in dieser Sammlung ab.|
||[GetLast ()](/javascript/api/word/word.paragraphcollection#getlast--)|Ruft den letzten Absatz in dieser Sammlung ab.|
||[getLastOrNullObject()](/javascript/api/word/word.paragraphcollection#getlastornullobject--)|Ruft den letzten Absatz in dieser Sammlung ab.|
|[Range](/javascript/api/word/word.range)|[compareLocationWith (Range: Word. Range)](/javascript/api/word/word.range#comparelocationwith-range-)|Vergleicht die Position dieses Bereichs mit der eines anderen Bereichs.|
||[expandTo (Range: Word. Range)](/javascript/api/word/word.range#expandto-range-)|Gibt einen neuen Bereich zurück, der diesen Bereich in beide Richtungen erweitert, um einen anderen Bereich zu überdecken.|
||[expandToOrNullObject (Range: Word. Range)](/javascript/api/word/word.range#expandtoornullobject-range-)|Gibt einen neuen Bereich zurück, der diesen Bereich in beide Richtungen erweitert, um einen anderen Bereich zu überdecken.|
||[getHyperlinkRanges()](/javascript/api/word/word.range#gethyperlinkranges--)|Ruft untergeordnete Linkbereiche innerhalb des Bereichs ab.|
||[getNextTextRange (endingMarks: String [], trimSpacing?: Boolean)](/javascript/api/word/word.range#getnexttextrange-endingmarks--trimspacing-)|Ruft den nächsten Textbereich mithilfe von Satzzeichen und/oder anderen Endmarkierungen ab.|
||[getNextTextRangeOrNullObject (endingMarks: String [], trimSpacing?: Boolean)](/javascript/api/word/word.range#getnexttextrangeornullobject-endingmarks--trimspacing-)|Ruft den nächsten Textbereich mithilfe von Satzzeichen und/oder anderen Endmarkierungen ab.|
||[GetRange (rangeLocation?: Word. rangeLocation)](/javascript/api/word/word.range#getrange-rangelocation-)|Klont den Bereich oder ruft den Anfangs- bzw. Endpunkt des Bereichs als neuen Bereich ab.|
||[getTextRanges (endingMarks: String [], trimSpacing?: Boolean)](/javascript/api/word/word.range#gettextranges-endingmarks--trimspacing-)|Ruft die untergeordneten Textbereiche im Bereich mithilfe von Satzzeichen und/oder anderen Endmarkierungen ab.|
||[hyperlink](/javascript/api/word/word.range#hyperlink)|Ruft den ersten Link in dem Bereich ab oder legt einen Link für den Bereich fest.|
||[Insertable (ROWCOUNT: Number, ColumnCount: Number, insertLocation: Word. insertLocation, Values?: String [] [])](/javascript/api/word/word.range#inserttable-rowcount--columncount--insertlocation--values-)|Fügt eine Tabelle mit der angegebenen Anzahl von Zeilen und Spalten ein.|
||[intersectWith (Range: Word. Range)](/javascript/api/word/word.range#intersectwith-range-)|Gibt einen neuen Bereich als Schnittmenge dieses Bereichs mit einem anderen Bereich zurück.|
||[intersectWithOrNullObject (Range: Word. Range)](/javascript/api/word/word.range#intersectwithornullobject-range-)|Gibt einen neuen Bereich als Schnittmenge dieses Bereichs mit einem anderen Bereich zurück.|
||[IsEmpty](/javascript/api/word/word.range#isempty)|Überprüft, ob die Länge des Bereichs 0 ist.|
||[Listen](/javascript/api/word/word.range#lists)|Ruft die Sammlung von Listenobjekten im Bereich ab.|
||[parentBody](/javascript/api/word/word.range#parentbody)|Ruft den übergeordneten Text des Bereichs ab.|
||[parentContentControlOrNullObject](/javascript/api/word/word.range#parentcontentcontrolornullobject)|Ruft das Inhaltssteuerelement ab, das den Bereich enthält.|
||[parentTable](/javascript/api/word/word.range#parenttable)|Ruft die Tabelle ab, die den Bereich enthält.|
||[parentTableCell](/javascript/api/word/word.range#parenttablecell)|Ruft die Tabellenzelle ab, die den Bereich enthält.|
||[parentTableCellOrNullObject](/javascript/api/word/word.range#parenttablecellornullobject)|Ruft die Tabellenzelle ab, die den Bereich enthält.|
||[parentTableOrNullObject](/javascript/api/word/word.range#parenttableornullobject)|Ruft die Tabelle ab, die den Bereich enthält.|
||[Tabellen](/javascript/api/word/word.range#tables)|Ruft die Sammlung von Tabellenobjekten im Bereich ab.|
||[Split (Trennzeichen: String [], multiparagraphs?: Boolean, trimDelimiters?: Boolean, trimSpacing?: Boolean)](/javascript/api/word/word.range#split-delimiters--multiparagraphs--trimdelimiters--trimspacing-)|Teilt den Bereich mithilfe von Trennzeichen in untergeordnete Bereiche.|
||[styleBuiltIn](/javascript/api/word/word.range#stylebuiltin)|Ruft den Namen der integrierten Formatvorlage für den Bereich ab oder legt ihn fest.|
|[RangeCollection](/javascript/api/word/word.rangecollection)|[getFirst()](/javascript/api/word/word.rangecollection#getfirst--)|Ruft den ersten Bereich in dieser Sammlung ab.|
||[getFirstOrNullObject()](/javascript/api/word/word.rangecollection#getfirstornullobject--)|Ruft den ersten Bereich in dieser Sammlung ab.|
|[RequestContext](/javascript/api/word/word.requestcontext)|[application](/javascript/api/word/word.requestcontext#application)|[API-Gruppe: WordApi 1,3] *|
|[Section](/javascript/api/word/word.section)|[GetNext ()](/javascript/api/word/word.section#getnext--)|Ruft den nächsten Abschnitt ab.|
||[getNextOrNullObject()](/javascript/api/word/word.section#getnextornullobject--)|Ruft den nächsten Abschnitt ab.|
|[SectionCollection](/javascript/api/word/word.sectioncollection)|[getFirst()](/javascript/api/word/word.sectioncollection#getfirst--)|Ruft den ersten Abschnitt in dieser Sammlung ab.|
||[getFirstOrNullObject()](/javascript/api/word/word.sectioncollection#getfirstornullobject--)|Ruft den ersten Abschnitt in dieser Sammlung ab.|
|[Table](/javascript/api/word/word.table)|[AddColumns (insertLocation: Word. insertLocation, ColumnCount: Number, Values?: String [] [])](/javascript/api/word/word.table#addcolumns-insertlocation--columncount--values-)|Fügt am Anfang oder Ende der Tabelle Spalten hinzu, wobei die erste oder letzte vorhandene Spalte als Vorlage verwendet wird.|
||[AddRows (insertLocation: Word. insertLocation, ROWCOUNT: Number, Values?: String [] [])](/javascript/api/word/word.table#addrows-insertlocation--rowcount--values-)|Fügt am Anfang oder Ende der Tabelle Zeilen hinzu, wobei die erste oder letzte vorhandene Zeile als Vorlage verwendet wird.|
||[Ausrichtung](/javascript/api/word/word.table#alignment)|Dient zum Abrufen oder Festlegen der Ausrichtung der Tabelle für die Seitenspalte.|
||[autoFitWindow()](/javascript/api/word/word.table#autofitwindow--)|Passt die Tabellenspalten automatisch an die Breite des Fensters an.|
||[clear()](/javascript/api/word/word.table#clear--)|Löscht den Inhalt der Tabelle.|
||[delete()](/javascript/api/word/word.table#delete--)|Löscht die gesamte Tabelle.|
||[deleteColumns (columnIndex: Number, ColumnCount?: Number)](/javascript/api/word/word.table#deletecolumns-columnindex--columncount-)|Löscht bestimmte Spalten.|
||[deleteRows (rowIndex: Number, ROWCOUNT?: Number)](/javascript/api/word/word.table#deleterows-rowindex--rowcount-)|Löscht bestimmte Zeilen.|
||[distributeColumns()](/javascript/api/word/word.table#distributecolumns--)|Verteilt die Breite der Spalten gleichmäßig.|
||[GetBorder (borderLocation: Word. borderLocation)](/javascript/api/word/word.table#getborder-borderlocation-)|Ruft die Rahmenart für den angegebenen Rahmen ab.|
||[getCell(rowIndex: number, cellIndex: number)](/javascript/api/word/word.table#getcell-rowindex--cellindex-)|Ruft die Tabellenzelle in einer angegebenen Zeile und Spalte ab.|
||[getCellOrNullObject (rowIndex: Number, cellIndex: Number)](/javascript/api/word/word.table#getcellornullobject-rowindex--cellindex-)|Ruft die Tabellenzelle in einer angegebenen Zeile und Spalte ab.|
||[getCellPadding (cellPaddingLocation: Word. cellPaddingLocation)](/javascript/api/word/word.table#getcellpadding-cellpaddinglocation-)|Ruft den Textabstand in Punkt ab.|
||[GetNext ()](/javascript/api/word/word.table#getnext--)|Ruft die nächste Tabelle ab.|
||[getNextOrNullObject()](/javascript/api/word/word.table#getnextornullobject--)|Ruft die nächste Tabelle ab.|
||[getParagraphAfter()](/javascript/api/word/word.table#getparagraphafter--)|Ruft den Absatz nach der Tabelle ab.|
||[getParagraphAfterOrNullObject()](/javascript/api/word/word.table#getparagraphafterornullobject--)|Ruft den Absatz nach der Tabelle ab.|
||[getParagraphBefore()](/javascript/api/word/word.table#getparagraphbefore--)|Ruft den Absatz vor der Tabelle ab.|
||[getParagraphBeforeOrNullObject()](/javascript/api/word/word.table#getparagraphbeforeornullobject--)|Ruft den Absatz vor der Tabelle ab.|
||[GetRange (rangeLocation?: Word. rangeLocation)](/javascript/api/word/word.table#getrange-rangelocation-)|Ruft den Bereich ab, der diese Tabelle enthält, oder den Bereich am Anfang bzw. Ende der Tabelle.|
||[headerRowCount](/javascript/api/word/word.table#headerrowcount)|Ruft die Anzahl von Kopfzeilen ab oder legt sie fest.|
||[horizontalAlignment](/javascript/api/word/word.table#horizontalalignment)|Ruft die horizontale Ausrichtung jeder Zelle in der Tabelle ab und legt sie fest.|
||[ignorePunct](/javascript/api/word/word.table#ignorepunct)||
||[ignoreSpace](/javascript/api/word/word.table#ignorespace)||
||[insertContentControl()](/javascript/api/word/word.table#insertcontentcontrol--)|Fügt ein Inhaltssteuerelement in der Tabelle ein.|
||[insertParagraph (ParagraphText: String, insertLocation: Word. insertLocation)](/javascript/api/word/word.table#insertparagraph-paragraphtext--insertlocation-)|Fügt an der angegebenen Position einen Absatz ein.|
||[Insertable (ROWCOUNT: Number, ColumnCount: Number, insertLocation: Word. insertLocation, Values?: String [] [])](/javascript/api/word/word.table#inserttable-rowcount--columncount--insertlocation--values-)|Fügt eine Tabelle mit der angegebenen Anzahl von Zeilen und Spalten ein.|
||[matchCase](/javascript/api/word/word.table#matchcase)||
||[matchPrefix](/javascript/api/word/word.table#matchprefix)||
||[matchSuffix](/javascript/api/word/word.table#matchsuffix)||
||[matchWholeWord](/javascript/api/word/word.table#matchwholeword)||
||[MatchWildcards](/javascript/api/word/word.table#matchwildcards)||
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
||[RowCount](/javascript/api/word/word.table#rowcount)|Ruft die Anzahl der Zeilen in der Tabelle ab.|
||[rows](/javascript/api/word/word.table#rows)|Ruft alle Zeilen der Tabelle ab.|
||[Tabellen](/javascript/api/word/word.table#tables)|Ruft die untergeordneten Tabellen ab, die eine Ebene tiefer verschachtelt sind.|
||[Suche (searchText: String, SearchOptions?: Word. SearchOptions \| {ignorePunct?: Boolean ignorespace?: Boolean MatchCase?: Boolean matchPrefix?: Boolean matchSuffix?: Boolean matchWholeWord?: Boolean MatchWildcards?: Boolean})](/javascript/api/word/word.table#search-searchtext--searchoptions--ignorepunct--ignorespace--matchcase--matchprefix--matchsuffix--matchwholeword--matchwildcards-)|Führt eine Suche mit den angegebenen SearchOptions für den Bereich des Table-Objekts aus.|
||[auswählen (SelectionMode?: Word. SelectionMode)](/javascript/api/word/word.table#select-selectionmode-)|Wählt die Tabelle oder die Position am Anfang oder Ende der Tabelle aus und navigiert die Word-Benutzeroberfläche an diese Position.|
||[setCellPadding (cellPaddingLocation: Word. cellPaddingLocation, CellPadding: Number)](/javascript/api/word/word.table#setcellpadding-cellpaddinglocation--cellpadding-)|Legt den Textabstand in Punkt fest.|
||[shadingColor](/javascript/api/word/word.table#shadingcolor)|Ruft die Schattierungsfarbe ab und legt diese fest.|
||[style](/javascript/api/word/word.table#style)|Ruft den Namen der Formatvorlage für die Tabelle ab oder legt ihn fest.|
||[styleBandedColumns](/javascript/api/word/word.table#stylebandedcolumns)|Ruft ab und legt fest, ob die Tabelle gebänderte Spalten enthält.|
||[styleBandedRows](/javascript/api/word/word.table#stylebandedrows)|Ruft ab und legt fest, ob die Tabelle gebänderte Zeilen enthält.|
||[styleBuiltIn](/javascript/api/word/word.table#stylebuiltin)|Ruft den Namen der integrierten Formatvorlage für die Tabelle ab oder legt ihn fest.|
||[styleFirstColumn](/javascript/api/word/word.table#stylefirstcolumn)|Ruft ab und legt fest, ob die Tabelle eine erste Spalte mit einer speziellen Formatvorlage aufweist.|
||[styleLastColumn](/javascript/api/word/word.table#stylelastcolumn)|Ruft ab und legt fest, ob die Tabelle eine letzte Spalte mit einer speziellen Formatvorlage aufweist.|
||[styleTotalRow](/javascript/api/word/word.table#styletotalrow)|Ruft ab und legt fest, ob die Tabelle eine Ergebniszeile (letzte Zeile) mit einer speziellen Formatvorlage aufweist.|
||[values](/javascript/api/word/word.table#values)|Ruft die Textwerte in der Tabelle als 2D-Javascript-Array ab und legt sie fest.|
||[verticalAlignment](/javascript/api/word/word.table#verticalalignment)|Ruft die vertikale Ausrichtung jeder Zelle in der Tabelle ab und legt sie fest.|
||[width](/javascript/api/word/word.table#width)|Ruft die Breite der Tabelle in Punkt ab und legt sie fest.|
|[TableBorder](/javascript/api/word/word.tableborder)|[color](/javascript/api/word/word.tableborder#color)|Ruft die Farbe des Tabellenrahmens ab oder legt Sie fest.|
||[Typ](/javascript/api/word/word.tableborder#type)|Ruft den Typ des Tabellenrahmens ab oder legt ihn fest.|
||[width](/javascript/api/word/word.tableborder#width)|Ruft die Breite des Tabellenrahmens in Punkt ab oder legt sie fest.|
|[TableCell](/javascript/api/word/word.tablecell)|[columnWidth](/javascript/api/word/word.tablecell#columnwidth)|Ruft die Breite der Spalte der Zelle in Punkt ab und legt sie fest.|
||[deleteColumn()](/javascript/api/word/word.tablecell#deletecolumn--)|Löscht die Spalte, die diese Zelle enthält.|
||[deleteRow ()](/javascript/api/word/word.tablecell#deleterow--)|Löscht die Zeile, die diese Zelle enthält.|
||[GetBorder (borderLocation: Word. borderLocation)](/javascript/api/word/word.tablecell#getborder-borderlocation-)|Ruft die Rahmenart für den angegebenen Rahmen ab.|
||[getCellPadding (cellPaddingLocation: Word. cellPaddingLocation)](/javascript/api/word/word.tablecell#getcellpadding-cellpaddinglocation-)|Ruft den Textabstand in Punkt ab.|
||[GetNext ()](/javascript/api/word/word.tablecell#getnext--)|Ruft die nächste Zelle ab.|
||[getNextOrNullObject()](/javascript/api/word/word.tablecell#getnextornullobject--)|Ruft die nächste Zelle ab.|
||[horizontalAlignment](/javascript/api/word/word.tablecell#horizontalalignment)|Ruft die horizontale Ausrichtung der Zelle ab und legt sie fest.|
||[insertColumns (insertLocation: Word. insertLocation, ColumnCount: Number, Values?: String [] [])](/javascript/api/word/word.tablecell#insertcolumns-insertlocation--columncount--values-)|Fügt links oder rechts von der Zelle Spalten hinzu, wobei die Spalte der Zelle als Vorlage verwendet wird.|
||[insertRows (insertLocation: Word. insertLocation, ROWCOUNT: Number, Values?: String [] [])](/javascript/api/word/word.tablecell#insertrows-insertlocation--rowcount--values-)|Fügt oberhalb oder unterhalb der Zelle Zeilen hinzu, wobei die Zeile der Zelle als Vorlage verwendet wird.|
||[Text](/javascript/api/word/word.tablecell#body)|Ruft das body-Objekt der Zelle ab.|
||[cellIndex](/javascript/api/word/word.tablecell#cellindex)|Ruft den Index der Zelle in der Zeile ab.|
||[ParentRow](/javascript/api/word/word.tablecell#parentrow)|Ruft die übergeordnete Zeile der Zelle ab.|
||[parentTable](/javascript/api/word/word.tablecell#parenttable)|Ruft die übergeordnete Tabelle der Zelle ab.|
||[rowIndex](/javascript/api/word/word.tablecell#rowindex)|Ruft den Index der Zeile mit der Zelle in der Tabelle ab.|
||[width](/javascript/api/word/word.tablecell#width)|Ruft die Breite der Zelle in Punkt ab.|
||[setCellPadding (cellPaddingLocation: Word. cellPaddingLocation, CellPadding: Number)](/javascript/api/word/word.tablecell#setcellpadding-cellpaddinglocation--cellpadding-)|Legt den Textabstand in Punkt fest.|
||[shadingColor](/javascript/api/word/word.tablecell#shadingcolor)|Ruft die Schattierungsfarbe der Zelle ab oder legt diese fest.|
||[value](/javascript/api/word/word.tablecell#value)|Ruft den Text der Zelle ab und legt ihn fest.|
||[verticalAlignment](/javascript/api/word/word.tablecell#verticalalignment)|Ruft die vertikale Ausrichtung der Zelle ab und legt sie fest.|
|[TableCellCollection](/javascript/api/word/word.tablecellcollection)|[getFirst()](/javascript/api/word/word.tablecellcollection#getfirst--)|Ruft die erste Tabellenzelle in dieser Sammlung ab.|
||[getFirstOrNullObject()](/javascript/api/word/word.tablecellcollection#getfirstornullobject--)|Ruft die erste Tabellenzelle in dieser Sammlung ab.|
||[items](/javascript/api/word/word.tablecellcollection#items)|Ruft die geladenen untergeordneten Elemente in dieser Sammlung ab.|
|[TableCollection](/javascript/api/word/word.tablecollection)|[getFirst()](/javascript/api/word/word.tablecollection#getfirst--)|Ruft die erste Tabelle in dieser Sammlung ab.|
||[getFirstOrNullObject()](/javascript/api/word/word.tablecollection#getfirstornullobject--)|Ruft die erste Tabelle in dieser Sammlung ab.|
||[items](/javascript/api/word/word.tablecollection#items)|Ruft die geladenen untergeordneten Elemente in dieser Sammlung ab.|
|[TableRow](/javascript/api/word/word.tablerow)|[clear()](/javascript/api/word/word.tablerow#clear--)|Löscht den Inhalt der Zeile.|
||[delete()](/javascript/api/word/word.tablerow#delete--)|Löscht die gesamte Zeile.|
||[GetBorder (borderLocation: Word. borderLocation)](/javascript/api/word/word.tablerow#getborder-borderlocation-)|Ruft die Rahmenart der Zellen in der Zeile ab.|
||[getCellPadding (cellPaddingLocation: Word. cellPaddingLocation)](/javascript/api/word/word.tablerow#getcellpadding-cellpaddinglocation-)|Ruft den Textabstand in Punkt ab.|
||[GetNext ()](/javascript/api/word/word.tablerow#getnext--)|Ruft die nächste Zeile ab.|
||[getNextOrNullObject()](/javascript/api/word/word.tablerow#getnextornullobject--)|Ruft die nächste Zeile ab.|
||[horizontalAlignment](/javascript/api/word/word.tablerow#horizontalalignment)|Ruft die horizontale Ausrichtung jeder Zelle in der Zeile ab und legt sie fest.|
||[ignorePunct](/javascript/api/word/word.tablerow#ignorepunct)||
||[ignoreSpace](/javascript/api/word/word.tablerow#ignorespace)||
||[insertRows (insertLocation: Word. insertLocation, ROWCOUNT: Number, Values?: String [] [])](/javascript/api/word/word.tablerow#insertrows-insertlocation--rowcount--values-)|Fügt Zeilen mit dieser Zeile als Vorlage ein.|
||[matchCase](/javascript/api/word/word.tablerow#matchcase)||
||[matchPrefix](/javascript/api/word/word.tablerow#matchprefix)||
||[matchSuffix](/javascript/api/word/word.tablerow#matchsuffix)||
||[matchWholeWord](/javascript/api/word/word.tablerow#matchwholeword)||
||[MatchWildcards](/javascript/api/word/word.tablerow#matchwildcards)||
||[preferredHeight](/javascript/api/word/word.tablerow#preferredheight)|Ruft die bevorzugte Höhe der Zeile in Punkt ab und legt sie fest.|
||[cellCount](/javascript/api/word/word.tablerow#cellcount)|Ruft die Anzahl der Zellen in der Zeile ab.|
||[Zellen](/javascript/api/word/word.tablerow#cells)|Ruft Zellen ab.|
||[font](/javascript/api/word/word.tablerow#font)|Ruft die Schriftart ab.|
||[isHeader](/javascript/api/word/word.tablerow#isheader)|Überprüft, ob die Zeile eine Kopfzeile ist.|
||[parentTable](/javascript/api/word/word.tablerow#parenttable)|Ruft die übergeordnete Tabelle ab.|
||[rowIndex](/javascript/api/word/word.tablerow#rowindex)|Ruft den Index der Zeile in der übergeordneten Tabelle ab.|
||[Suche (searchText: String, SearchOptions?: Word. SearchOptions \| {ignorePunct?: Boolean ignorespace?: Boolean MatchCase?: Boolean matchPrefix?: Boolean matchSuffix?: Boolean matchWholeWord?: Boolean MatchWildcards?: Boolean})](/javascript/api/word/word.tablerow#search-searchtext--searchoptions--ignorepunct--ignorespace--matchcase--matchprefix--matchsuffix--matchwholeword--matchwildcards-)|Führt eine Suche mit den angegebenen SearchOptions für den Bereich der Zeile aus.|
||[auswählen (SelectionMode?: Word. SelectionMode)](/javascript/api/word/word.tablerow#select-selectionmode-)|Wählt die Zeile aus und navigiert in der Word-Benutzeroberfläche zur dieser Position.|
||[setCellPadding (cellPaddingLocation: Word. cellPaddingLocation, CellPadding: Number)](/javascript/api/word/word.tablerow#setcellpadding-cellpaddinglocation--cellpadding-)|Legt den Textabstand in Punkt fest.|
||[shadingColor](/javascript/api/word/word.tablerow#shadingcolor)|Ruft die Schattierungsfarbe ab und legt diese fest.|
||[values](/javascript/api/word/word.tablerow#values)|Ruft die Textwerte in der Zeile als 2D-JavaScript-Array ab und legt Sie fest.|
||[verticalAlignment](/javascript/api/word/word.tablerow#verticalalignment)|Ruft die vertikale Ausrichtung der Zellen in der Zeile ab und legt sie fest.|
|[TableRowCollection](/javascript/api/word/word.tablerowcollection)|[getFirst()](/javascript/api/word/word.tablerowcollection#getfirst--)|Ruft die erste Zeile in dieser Sammlung ab.|
||[getFirstOrNullObject()](/javascript/api/word/word.tablerowcollection#getfirstornullobject--)|Ruft die erste Zeile in dieser Sammlung ab.|
||[items](/javascript/api/word/word.tablerowcollection#items)|Ruft die geladenen untergeordneten Elemente in dieser Sammlung ab.|
