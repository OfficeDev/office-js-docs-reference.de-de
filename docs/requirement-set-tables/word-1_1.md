| Klasse | Felder | Beschreibung |
|:---|:---|:---|
|[Body](/javascript/api/word/word.body)|[clear()](/javascript/api/word/word.body#clear--)|Löscht den Inhalt des Body-Objekts.|
||[getHtml()](/javascript/api/word/word.body#gethtml--)|Ruft eine HTML-Darstellung des Body-Objekts ab.|
||[getOoxml()](/javascript/api/word/word.body#getooxml--)|Ruft die OOXML (Office Open XML)-Darstellung des Body-Objekts ab.|
||[ignorePunct](/javascript/api/word/word.body#ignorepunct)||
||[ignoreSpace](/javascript/api/word/word.body#ignorespace)||
||[insertBreak (breaktype: Word. breaktype, insertLocation: Word. insertLocation)](/javascript/api/word/word.body#insertbreak-breaktype--insertlocation-)|Fügt an der angegebenen Position im Hauptdokument einen Umbruch ein.|
||[insertContentControl()](/javascript/api/word/word.body#insertcontentcontrol--)|Umschließt das Body-Objekt mit einem Rich-Text-Inhaltssteuerelement.|
||[insertFileFromBase64 (Base64-Datei: String, insertLocation: Word. insertLocation)](/javascript/api/word/word.body#insertfilefrombase64-base64file--insertlocation-)|Fügt in den Textkörper an der angegebenen Position ein Dokument ein.|
||[insertHtml (HTML: String, insertLocation: Word. insertLocation)](/javascript/api/word/word.body#inserthtml-html--insertlocation-)|Fügt an der angegebenen Position HTML-Code ein.|
||[insertOoxml (OOXML: String, insertLocation: Word. insertLocation)](/javascript/api/word/word.body#insertooxml-ooxml--insertlocation-)|Fügt an der angegebenen Position OOXML-Code ein.|
||[insertParagraph (ParagraphText: String, insertLocation: Word. insertLocation)](/javascript/api/word/word.body#insertparagraph-paragraphtext--insertlocation-)|Fügt an der angegebenen Position einen Absatz ein.|
||[InsertText (Text: String, insertLocation: Word. insertLocation)](/javascript/api/word/word.body#inserttext-text--insertlocation-)|Fügt in den Textkörper an der angegebenen Position Text ein.|
||[matchCase](/javascript/api/word/word.body#matchcase)||
||[matchPrefix](/javascript/api/word/word.body#matchprefix)||
||[matchSuffix](/javascript/api/word/word.body#matchsuffix)||
||[matchWholeWord](/javascript/api/word/word.body#matchwholeword)||
||[MatchWildcards](/javascript/api/word/word.body#matchwildcards)||
||[contentControls](/javascript/api/word/word.body#contentcontrols)|Ruft die Auflistung der Rich-Text-Inhaltssteuerelemente im Textkörper ab.|
||[font](/javascript/api/word/word.body#font)|Ruft das Textformat des Textkörpers ab.|
||[inlinePictures](/javascript/api/word/word.body#inlinepictures)|Ruft die Auflistung der Inline Picture-Objekte im Textkörper ab.|
||[paragraphs](/javascript/api/word/word.body#paragraphs)|Ruft die Auflistung von Paragraph-Objekten im Textkörper ab.|
||[parentContentControl](/javascript/api/word/word.body#parentcontentcontrol)|Ruft das Inhaltssteuerelement ab, das den Textkörper enthält.|
||[text](/javascript/api/word/word.body#text)|Ruft den Textkörper ab.|
||[Suche (searchText: String, SearchOptions?: Word. SearchOptions \| {ignorePunct?: Boolean ignorespace?: Boolean MatchCase?: Boolean matchPrefix?: Boolean matchSuffix?: Boolean matchWholeWord?: Boolean MatchWildcards?: Boolean})](/javascript/api/word/word.body#search-searchtext--searchoptions--ignorepunct--ignorespace--matchcase--matchprefix--matchsuffix--matchwholeword--matchwildcards-)|Führt eine Suche mit den angegebenen SearchOptions für den Bereich des Body-Objekts aus.|
||[auswählen (SelectionMode?: Word. SelectionMode)](/javascript/api/word/word.body#select-selectionmode-)|Wählt den Textkörper und navigiert die Word-Benutzeroberfläche an diese Position.|
||[style](/javascript/api/word/word.body#style)|Dient zum Abrufen oder Festlegen des Formatvorlagen namens für den Textkörper.|
|[ContentControl](/javascript/api/word/word.contentcontrol)|[Aussehen](/javascript/api/word/word.contentcontrol#appearance)|Ruft die Darstellung des Inhaltssteuerelements ab oder fest sie fest.|
||[können](/javascript/api/word/word.contentcontrol#cannotdelete)|Ruft einen Wert ab oder legt ihn fest, der angibt, ob der Benutzer das Inhaltssteuerelement löschen kann.|
||[cannotEdit](/javascript/api/word/word.contentcontrol#cannotedit)|Ruft einen Wert ab oder legt ihn fest, der angibt, ob der Benutzer das Inhaltssteuerelement bearbeiten kann.|
||[clear()](/javascript/api/word/word.contentcontrol#clear--)|Löscht den Inhalt des Inhaltssteuerelements.|
||[color](/javascript/api/word/word.contentcontrol#color)|Ruft die Farbe des Inhaltssteuerelements ab oder legt sie fest.|
||[Delete (keepContent: Boolean)](/javascript/api/word/word.contentcontrol#delete-keepcontent-)|Löscht das Inhaltssteuerelement und dessen Inhalt.|
||[getHtml()](/javascript/api/word/word.contentcontrol#gethtml--)|Ruft eine HTML-Darstellung des Inhaltssteuerelement-Objekts ab.|
||[getOoxml()](/javascript/api/word/word.contentcontrol#getooxml--)|Ruft die Office Open XML (OOXML)-Darstellung des Inhaltssteuerelement-Objekts ab.|
||[ignorePunct](/javascript/api/word/word.contentcontrol#ignorepunct)||
||[ignoreSpace](/javascript/api/word/word.contentcontrol#ignorespace)||
||[insertBreak (breaktype: Word. breaktype, insertLocation: Word. insertLocation)](/javascript/api/word/word.contentcontrol#insertbreak-breaktype--insertlocation-)|Fügt an der angegebenen Position im Hauptdokument einen Umbruch ein.|
||[insertFileFromBase64 (Base64-Datei: String, insertLocation: Word. insertLocation)](/javascript/api/word/word.contentcontrol#insertfilefrombase64-base64file--insertlocation-)|Fügt ein Dokument in das Inhaltssteuerelement an der angegebenen Position ein.|
||[insertHtml (HTML: String, insertLocation: Word. insertLocation)](/javascript/api/word/word.contentcontrol#inserthtml-html--insertlocation-)|Fügt an der angegebenen Position HTML ein.|
||[insertOoxml (OOXML: String, insertLocation: Word. insertLocation)](/javascript/api/word/word.contentcontrol#insertooxml-ooxml--insertlocation-)|Fügt OOXML an der angegebenen Position in das Inhaltssteuerelement ein.|
||[insertParagraph (ParagraphText: String, insertLocation: Word. insertLocation)](/javascript/api/word/word.contentcontrol#insertparagraph-paragraphtext--insertlocation-)|Fügt an der angegebenen Position einen Absatz ein.|
||[InsertText (Text: String, insertLocation: Word. insertLocation)](/javascript/api/word/word.contentcontrol#inserttext-text--insertlocation-)|Fügt an der angegebenen Position Text in das Inhaltssteuerelement ein.|
||[matchCase](/javascript/api/word/word.contentcontrol#matchcase)||
||[matchPrefix](/javascript/api/word/word.contentcontrol#matchprefix)||
||[matchSuffix](/javascript/api/word/word.contentcontrol#matchsuffix)||
||[matchWholeWord](/javascript/api/word/word.contentcontrol#matchwholeword)||
||[MatchWildcards](/javascript/api/word/word.contentcontrol#matchwildcards)||
||[placeholderText](/javascript/api/word/word.contentcontrol#placeholdertext)|Ruft den Platzhaltertext des Inhaltssteuerelements ab oder legt ihn fest.|
||[contentControls](/javascript/api/word/word.contentcontrol#contentcontrols)|Ruft die Auflistung von Inhaltssteuerelement-Objekten im Inhaltssteuerelement ab.|
||[font](/javascript/api/word/word.contentcontrol#font)|Ruft das Textformat des Inhaltssteuerelements ab.|
||[id](/javascript/api/word/word.contentcontrol#id)|Ruft eine Ganzzahl ab, die den Inhaltssteuerelement-Bezeichner darstellt.|
||[inlinePictures](/javascript/api/word/word.contentcontrol#inlinepictures)|Ruft die Auflistung von inlinePicture-Objekten im Inhaltssteuerelement ab.|
||[paragraphs](/javascript/api/word/word.contentcontrol#paragraphs)|Ruft die Auflistung von Abschnittsobjekten im Inhaltssteuerelement ab.|
||[parentContentControl](/javascript/api/word/word.contentcontrol#parentcontentcontrol)|Ruft das Inhaltssteuerelement ab, das das Inhaltssteuerelement enthält.|
||[text](/javascript/api/word/word.contentcontrol#text)|Ruft den Text des Inhaltssteuerelements ab.|
||[Typ](/javascript/api/word/word.contentcontrol#type)|Ruft den Typen des Inhaltssteuerelements ab.|
||[removeWhenEdited](/javascript/api/word/word.contentcontrol#removewhenedited)|Ruft einen Wert ab oder legt ihn fest, der angibt, ob der Benutzer das Inhaltssteuerelement nach der Bearbeitung entfernt wird.|
||[Suche (searchText: String, SearchOptions?: Word. SearchOptions \| {ignorePunct?: Boolean ignorespace?: Boolean MatchCase?: Boolean matchPrefix?: Boolean matchSuffix?: Boolean matchWholeWord?: Boolean MatchWildcards?: Boolean})](/javascript/api/word/word.contentcontrol#search-searchtext--searchoptions--ignorepunct--ignorespace--matchcase--matchprefix--matchsuffix--matchwholeword--matchwildcards-)|Führt eine Suche mit den angegebenen SearchOptions für den Bereich des Inhaltssteuerelement-Objekts aus.|
||[auswählen (SelectionMode?: Word. SelectionMode)](/javascript/api/word/word.contentcontrol#select-selectionmode-)|Wählt das Inhaltssteuerelement aus.|
||[style](/javascript/api/word/word.contentcontrol#style)|Dient zum Abrufen oder Festlegen des Formatvorlagen namens für das Inhaltssteuerelement.|
||[Tag](/javascript/api/word/word.contentcontrol#tag)|Ruft ein Tag ab oder legt es fest, um ein Inhaltssteuerelement zu identifizieren.|
||[title](/javascript/api/word/word.contentcontrol#title)|Ruft den Titel eines Inhaltssteuerelements ab oder legt ihn fest.|
|[ContentControlCollection](/javascript/api/word/word.contentcontrolcollection)|[getById(id: number)](/javascript/api/word/word.contentcontrolcollection#getbyid-id-)|Ruft ein Inhaltssteuerelement mithilfe der ID ab.|
||[getByTag(tag: string)](/javascript/api/word/word.contentcontrolcollection#getbytag-tag-)|Ruft die Inhaltssteuerelemente ab, die das angegebene Tag enthalten.|
||[getByTitle(title: string)](/javascript/api/word/word.contentcontrolcollection#getbytitle-title-)|Ruft die Inhaltssteuerelemente ab, die den angegebenen Titel enthalten.|
||[getItem(index: number)](/javascript/api/word/word.contentcontrolcollection#getitem-index-)|Ruft ein Inhaltssteuerelement nach dem Index in der Auflistung ab.|
||[items](/javascript/api/word/word.contentcontrolcollection#items)|Ruft die geladenen untergeordneten Elemente in dieser Sammlung ab.|
|[Document](/javascript/api/word/word.document)|[GetSelection ()](/javascript/api/word/word.document#getselection--)|Ruft die aktuelle Auswahl des Dokuments ab.|
||[Text](/javascript/api/word/word.document#body)|Ruft das Body-Objekt des Dokuments ab.|
||[contentControls](/javascript/api/word/word.document#contentcontrols)|Ruft die Auflistung der Inhaltssteuerelement-Objekte im Dokument ab.|
||[gespeichert](/javascript/api/word/word.document#saved)|Gibt an, ob die Änderungen im Dokument gespeichert wurden.|
||[sections](/javascript/api/word/word.document#sections)|Ruft die Auflistung von Section-Objekten im Dokument ab.|
||[save()](/javascript/api/word/word.document#save--)|Speichert das Dokument.|
|[Font](/javascript/api/word/word.font)|[bold](/javascript/api/word/word.font#bold)|Ruft einen Wert ab, der angibt, ob die Schriftart fett formatiert ist, oder legt diesen fest.|
||[color](/javascript/api/word/word.font#color)|Ruft die Farbe für die angegebene Schriftart ab, oder legt diese fest.|
||[doubleStrikeThrough](/javascript/api/word/word.font#doublestrikethrough)|Dient zum Abrufen oder Festlegen eines Werts, der angibt, ob die Schriftart doppelt durchgestrichen wurde.|
||[HighlightColor](/javascript/api/word/word.font#highlightcolor)|Ruft die Hervorhebungsfarbe ab oder legt Sie fest.|
||[italic](/javascript/api/word/word.font#italic)|Ruft einen Wert ab, der angibt, ob die Schriftart kursiv ist, oder legt diesen fest.|
||[name](/javascript/api/word/word.font#name)|Ruft einen Wert ab, der den Namen der Schriftart angibt, oder legt diesen fest.|
||[size](/javascript/api/word/word.font#size)|Ruft einen Wert ab, der den Schriftgrad in Punkten angibt, oder legt diesen fest.|
||[durchgestrichen](/javascript/api/word/word.font#strikethrough)|Dient zum Abrufen oder Festlegen eines Werts, der angibt, ob die Schriftart durchgestrichen wurde.|
||[subscript](/javascript/api/word/word.font#subscript)|Ruft einen Wert ab, der angibt, ob die Schriftart tiefgestellt ist, oder legt diesen fest.|
||[superscript](/javascript/api/word/word.font#superscript)|Ruft einen Wert ab, der angibt, ob die Schriftart hochgestellt ist, oder legt diesen fest.|
||[underline](/javascript/api/word/word.font#underline)|Ruft einen Wert ab, der den auf die Schriftart angewendete Art der Unterstreichung angibt, oder legt diesen fest.|
|[InlinePicture](/javascript/api/word/word.inlinepicture)|[altTextDescription](/javascript/api/word/word.inlinepicture#alttextdescription)|Dient zum Abrufen oder Festlegen einer Zeichenfolge, die den mit dem eingebundenen Bild verknüpften alternativen Text darstellt.|
||[altTextTitle](/javascript/api/word/word.inlinepicture#alttexttitle)|Ruft eine Zeichenfolge ab, die den Titel für das eingebundene Bild enthält, oder legt diese fest.|
||[getBase64ImageSrc()](/javascript/api/word/word.inlinepicture#getbase64imagesrc--)|Ruft die Base64-verschlüsselte Zeichenfolgendarstellung des eingebundenen Bilds ab.|
||[height](/javascript/api/word/word.inlinepicture#height)|Ruft eine Zahl ab, die die Höhe des eingebundenen Bilds beschreibt, oder legt diese fest.|
||[hyperlink](/javascript/api/word/word.inlinepicture#hyperlink)|Dient zum Abrufen oder Festlegen eines Hyperlinks für das Bild.|
||[insertContentControl()](/javascript/api/word/word.inlinepicture#insertcontentcontrol--)|Umschließt das eingebundene Bild mit einem Rich-Text-Inhaltssteuerelement.|
||[lockAspectRatio](/javascript/api/word/word.inlinepicture#lockaspectratio)|Ruft einen Wert ab, der angibt, ob das eingebundene Bild seine ursprünglichen Proportionen beibehält, wenn Sie seine Größe ändern, oder legt diesen fest.|
||[parentContentControl](/javascript/api/word/word.inlinepicture#parentcontentcontrol)|Ruft das Inhaltssteuerelement ab, das das eingebundene Bild enthält.|
||[width](/javascript/api/word/word.inlinepicture#width)|Ruft eine Zahl ab, die die Breite des eingebundenen Bilds beschreibt, oder legt diese fest.|
|[InlinePictureCollection](/javascript/api/word/word.inlinepicturecollection)|[items](/javascript/api/word/word.inlinepicturecollection#items)|Ruft die geladenen untergeordneten Elemente in dieser Sammlung ab.|
|[Paragraph](/javascript/api/word/word.paragraph)|[Ausrichtung](/javascript/api/word/word.paragraph#alignment)|Ruft die Ausrichtung für einen Absatz ab oder legt sie fest.|
||[clear()](/javascript/api/word/word.paragraph#clear--)|Löscht den Inhalt des Abschnittsobjekts.|
||[delete()](/javascript/api/word/word.paragraph#delete--)|Löscht den Absatz und seinen Inhalt aus dem Dokument.|
||[firstLineIndent](/javascript/api/word/word.paragraph#firstlineindent)|Ruft den Wert in Punkten für einen Erstzeileneinzug oder einen hängenden Einzug ab, oder legt ihn fest.|
||[getHtml()](/javascript/api/word/word.paragraph#gethtml--)|Ruft eine HTML-Darstellung des Paragraph-Objekts ab.|
||[getOoxml()](/javascript/api/word/word.paragraph#getooxml--)|Ruft die Office Open XML (OOXML)-Darstellung des Abschnittsobjekts ab.|
||[ignorePunct](/javascript/api/word/word.paragraph#ignorepunct)||
||[ignoreSpace](/javascript/api/word/word.paragraph#ignorespace)||
||[insertBreak (breaktype: Word. breaktype, insertLocation: Word. insertLocation)](/javascript/api/word/word.paragraph#insertbreak-breaktype--insertlocation-)|Fügt an der angegebenen Position im Hauptdokument einen Umbruch ein.|
||[insertContentControl()](/javascript/api/word/word.paragraph#insertcontentcontrol--)|Umschließt das Abschnittsobjekt mit einem Rich-Text-Inhaltssteuerelement.|
||[insertFileFromBase64 (Base64-Datei: String, insertLocation: Word. insertLocation)](/javascript/api/word/word.paragraph#insertfilefrombase64-base64file--insertlocation-)|Fügt ein Dokument in den Absatz an der angegebenen Position ein.|
||[insertHtml (HTML: String, insertLocation: Word. insertLocation)](/javascript/api/word/word.paragraph#inserthtml-html--insertlocation-)|Fügt in den aktuellen Abschnitt an der angegebenen Position HTML-Code ein.|
||[insertInlinePictureFromBase64 (base64EncodedImage: String, insertLocation: Word. insertLocation)](/javascript/api/word/word.paragraph#insertinlinepicturefrombase64-base64encodedimage--insertlocation-)|Fügt in den aktuellen Abschnitt an der angegebenen Position ein Bild ein.|
||[insertOoxml (OOXML: String, insertLocation: Word. insertLocation)](/javascript/api/word/word.paragraph#insertooxml-ooxml--insertlocation-)|Fügt OOXML an der angegebenen Position in den Absatz ein.|
||[insertParagraph (ParagraphText: String, insertLocation: Word. insertLocation)](/javascript/api/word/word.paragraph#insertparagraph-paragraphtext--insertlocation-)|Fügt an der angegebenen Position einen Absatz ein.|
||[InsertText (Text: String, insertLocation: Word. insertLocation)](/javascript/api/word/word.paragraph#inserttext-text--insertlocation-)|Fügt in den Abschnitt an der angegebenen Position Text ein.|
||[LeftIndent](/javascript/api/word/word.paragraph#leftindent)|Ruft den linken Einzugswert für den Abschnitt in Punkten ab oder legt ihn fest.|
||[LineSpacing](/javascript/api/word/word.paragraph#linespacing)|Ruft den Zeilenabstand für den angegebenen Abschnitt in Punkten ab, oder legt ihn fest.|
||[lineUnitAfter](/javascript/api/word/word.paragraph#lineunitafter)|Ruft den Abstand nach dem Absatz in Rasterlinien ab oder legt ihn fest.|
||[lineUnitBefore](/javascript/api/word/word.paragraph#lineunitbefore)|Ruft den Abstand vor dem Abschnitt in Rasterlinien ab, oder legt ihn fest.|
||[matchCase](/javascript/api/word/word.paragraph#matchcase)||
||[matchPrefix](/javascript/api/word/word.paragraph#matchprefix)||
||[matchSuffix](/javascript/api/word/word.paragraph#matchsuffix)||
||[matchWholeWord](/javascript/api/word/word.paragraph#matchwholeword)||
||[MatchWildcards](/javascript/api/word/word.paragraph#matchwildcards)||
||[outlineLevel](/javascript/api/word/word.paragraph#outlinelevel)|Ruft die Gliederungsebene für den Absatz fest oder ruft sie ab.|
||[contentControls](/javascript/api/word/word.paragraph#contentcontrols)|Ruft die Auflistung der Inhaltssteuerelement-Objekte im Absatz ab.|
||[font](/javascript/api/word/word.paragraph#font)|Ruft das Textformat des Abschnitts ab.|
||[inlinePictures](/javascript/api/word/word.paragraph#inlinepictures)|Ruft die Auflistung der Inline Picture-Objekte im Absatz ab.|
||[parentContentControl](/javascript/api/word/word.paragraph#parentcontentcontrol)|Ruft das Inhaltssteuerelement ab, das den Abschnitt enthält.|
||[text](/javascript/api/word/word.paragraph#text)|Ruft den Text des Absatzes ab.|
||[rightIndent](/javascript/api/word/word.paragraph#rightindent)|Ruft den rechten Einzugswert für den Abschnitt in Punkten ab oder legt ihn fest.|
||[Suche (searchText: String, SearchOptions?: Word. SearchOptions \| {ignorePunct?: Boolean ignorespace?: Boolean MatchCase?: Boolean matchPrefix?: Boolean matchSuffix?: Boolean matchWholeWord?: Boolean MatchWildcards?: Boolean})](/javascript/api/word/word.paragraph#search-searchtext--searchoptions--ignorepunct--ignorespace--matchcase--matchprefix--matchsuffix--matchwholeword--matchwildcards-)|Führt eine Suche mit den angegebenen SearchOptions für den Bereich des Paragraph-Objekts aus.|
||[auswählen (SelectionMode?: Word. SelectionMode)](/javascript/api/word/word.paragraph#select-selectionmode-)|Wählt und navigiert die Word-Benutzeroberfläche zu diesem Absatz.|
||[spaceAfter](/javascript/api/word/word.paragraph#spaceafter)|Ruft den Abstand nach dem Abschnitt in Punkten ab.|
||[spaceBefore](/javascript/api/word/word.paragraph#spacebefore)|Ruft den Abstand vor dem Abschnitt in Punkten ab.|
||[style](/javascript/api/word/word.paragraph#style)|Ruft den Formatvorlagennamen für den Absatz ab oder legt ihn fest.|
|[ParagraphCollection](/javascript/api/word/word.paragraphcollection)|[items](/javascript/api/word/word.paragraphcollection#items)|Ruft die geladenen untergeordneten Elemente in dieser Sammlung ab.|
|[Range](/javascript/api/word/word.range)|[clear()](/javascript/api/word/word.range#clear--)|Löscht den Inhalt des Bereichsobjekts.|
||[delete()](/javascript/api/word/word.range#delete--)|Löscht den Bereich und seinen Inhalt aus dem Dokument.|
||[getHtml()](/javascript/api/word/word.range#gethtml--)|Ruft eine HTML-Darstellung des Range-Objekts ab.|
||[getOoxml()](/javascript/api/word/word.range#getooxml--)|Ruft die OOXML-Darstellung des Bereichsobjekts ab.|
||[ignorePunct](/javascript/api/word/word.range#ignorepunct)||
||[ignoreSpace](/javascript/api/word/word.range#ignorespace)||
||[insertBreak (breaktype: Word. breaktype, insertLocation: Word. insertLocation)](/javascript/api/word/word.range#insertbreak-breaktype--insertlocation-)|Fügt an der angegebenen Position im Hauptdokument einen Umbruch ein.|
||[insertContentControl()](/javascript/api/word/word.range#insertcontentcontrol--)|Umschließt das Bereichsobjekt mit einem Rich-Text-Inhaltssteuerelement.|
||[insertFileFromBase64 (Base64-Datei: String, insertLocation: Word. insertLocation)](/javascript/api/word/word.range#insertfilefrombase64-base64file--insertlocation-)|Fügt an der angegebenen Position ein Dokument ein.|
||[insertHtml (HTML: String, insertLocation: Word. insertLocation)](/javascript/api/word/word.range#inserthtml-html--insertlocation-)|Fügt an der angegebenen Position HTML-Code ein.|
||[insertOoxml (OOXML: String, insertLocation: Word. insertLocation)](/javascript/api/word/word.range#insertooxml-ooxml--insertlocation-)|Fügt an der angegebenen Position OOXML-Code ein.|
||[insertParagraph (ParagraphText: String, insertLocation: Word. insertLocation)](/javascript/api/word/word.range#insertparagraph-paragraphtext--insertlocation-)|Fügt an der angegebenen Position einen Absatz ein.|
||[InsertText (Text: String, insertLocation: Word. insertLocation)](/javascript/api/word/word.range#inserttext-text--insertlocation-)|Fügt an der angegebenen Position Text ein.|
||[matchCase](/javascript/api/word/word.range#matchcase)||
||[matchPrefix](/javascript/api/word/word.range#matchprefix)||
||[matchSuffix](/javascript/api/word/word.range#matchsuffix)||
||[matchWholeWord](/javascript/api/word/word.range#matchwholeword)||
||[MatchWildcards](/javascript/api/word/word.range#matchwildcards)||
||[contentControls](/javascript/api/word/word.range#contentcontrols)|Ruft die Auflistung der Inhaltssteuerelement-Objekte im Bereich ab.|
||[font](/javascript/api/word/word.range#font)|Ruft das Textformat des Bereichs ab.|
||[paragraphs](/javascript/api/word/word.range#paragraphs)|Ruft die Auflistung von Paragraph-Objekten im Bereich ab.|
||[parentContentControl](/javascript/api/word/word.range#parentcontentcontrol)|Ruft das Inhaltssteuerelement ab, das den Bereich enthält.|
||[text](/javascript/api/word/word.range#text)|Ruft den Text des Bereichs ab.|
||[Suche (searchText: String, SearchOptions?: Word. SearchOptions \| {ignorePunct?: Boolean ignorespace?: Boolean MatchCase?: Boolean matchPrefix?: Boolean matchSuffix?: Boolean matchWholeWord?: Boolean MatchWildcards?: Boolean})](/javascript/api/word/word.range#search-searchtext--searchoptions--ignorepunct--ignorespace--matchcase--matchprefix--matchsuffix--matchwholeword--matchwildcards-)|Führt eine Suche mit den angegebenen SearchOptions für den Bereich des Range-Objekts aus.|
||[auswählen (SelectionMode?: Word. SelectionMode)](/javascript/api/word/word.range#select-selectionmode-)|Wählt und navigiert die Word-Benutzeroberfläche zu diesem Bereich.|
||[style](/javascript/api/word/word.range#style)|Dient zum Abrufen oder Festlegen des Formatvorlagen namens für den Bereich.|
|[RangeCollection](/javascript/api/word/word.rangecollection)|[items](/javascript/api/word/word.rangecollection#items)|Ruft die geladenen untergeordneten Elemente in dieser Sammlung ab.|
|[SearchOptions](/javascript/api/word/word.searchoptions)|[ignorePunct](/javascript/api/word/word.searchoptions#ignorepunct)|Ruft einen Wert ab oder legt ihn fest, um anzugeben, dass alle Interpunktionszeichen zwischen Wörtern ignoriert werden sollen.|
||[ignoreSpace](/javascript/api/word/word.searchoptions#ignorespace)|Dient zum Abrufen oder Festlegen eines Werts, der angibt, ob alle Leerzeichen zwischen Wörtern ignoriert werden sollen.|
||[matchCase](/javascript/api/word/word.searchoptions#matchcase)|Ruft einen Wert ab oder legt ihn fest, um anzugeben, dass die Suche mit Berücksichtigung von Groß-/Kleinschreibung durchgeführt werden soll.|
||[matchPrefix](/javascript/api/word/word.searchoptions#matchprefix)|Ruft einen Wert ab oder legt ihn fest, um anzugeben, ob Wörter beachtet werden sollen, die mit der Suchzeichenfolge beginnen.|
||[matchSuffix](/javascript/api/word/word.searchoptions#matchsuffix)|Ruft einen Wert ab oder legt ihn fest, um anzugeben, ob Wörter beachtet werden sollen, die mit der Suchzeichenfolge enden.|
||[matchWholeWord](/javascript/api/word/word.searchoptions#matchwholeword)|Ruft einen Wert ab oder legt ihn fest, um anzugeben, dass nur ganze Wörter und nicht nach Textelementen in längeren Wörtern gesucht werden soll.|
||[MatchWildcards](/javascript/api/word/word.searchoptions#matchwildcards)|Ruft einen Wert ab oder legt ihn fest, um anzugeben, ob die Suche mithilfe von speziellen Suchoperatoren ausgeführt wird.|
|[Section](/javascript/api/word/word.section)|[getfooter (Typ: Word. HeaderFooterType)](/javascript/api/word/word.section#getfooter-type-)|Ruft eine der Fußzeilen des Abschnitts ab.|
||[GetHeader (Typ: Word. HeaderFooterType)](/javascript/api/word/word.section#getheader-type-)|Ruft eine der Kopfzeilen des Abschnitts ab.|
||[Text](/javascript/api/word/word.section#body)|Ruft das Body-Objekt des Abschnitts ab.|
|[SectionCollection](/javascript/api/word/word.sectioncollection)|[items](/javascript/api/word/word.sectioncollection#items)|Ruft die geladenen untergeordneten Elemente in dieser Sammlung ab.|
