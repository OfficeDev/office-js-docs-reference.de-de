# <a name="word-javascript-api-requirement-sets"></a>JavaScript-API-Anforderungssätze für Word

Anforderungssätze sind benannte Gruppen von API-Mitgliedern. Office-Add-ins anforderungssätzen im Manifest angegebenen verwenden oder eine Überprüfung zur Laufzeit verwenden, um zu bestimmen, ob ein Office-Host APIs unterstützt, die ein Add-in muss. Weitere Informationen finden Sie unter [Office-Versionen und Anforderung festgelegt](https://docs.microsoft.com/office/dev/add-ins/develop/office-versions-and-requirement-sets).

Word-add-ins führen Sie über mehrere Versionen von Office, einschließlich Office 2016 oder höher für Windows, Office für iPad, Office für Mac und Office Online. Die folgende Tabelle enthält die Word anforderungssätze, die Office-hostanwendungen, die Anforderung festgelegt, und der Build oder Version Numbers für Anwendungsbereiche unterstützen.

> [!NOTE]
> Für die anforderungssätze, die als Beta gekennzeichnet sind, verwenden Sie die angegebene (oder höhere) Version von Office-Software und die Beta-Bibliothek mit dem CDN verwenden: https://appsforoffice.microsoft.com/lib/beta/hosted/office.js.
> 
> Einträge, die nicht als Beta im Allgemeinen verfügbar sind, und Sie auch weiterhin können mithilfe Produktion CDN-Bibliothek aufgeführt:https://appsforoffice.microsoft.com/lib/1/hosted/office.js

|  Anforderungssatz  |   Office 365 für Windows\*  |  Office 365 für iPad  |  Office 365 für Mac  | Office Online  | Office Online Server  |
|:-----|-----|:-----|:-----|:-----|:-----|
| WordApi 1.3 | Version 1612 (Build 7668.1000) oder höher| März 2017, 2.22 oder höher | März 2017, 15.32 oder höher| März 2017 ||
| WordApi 1.2  | Update vom Dezember 2015, Version 1601 (Build 6568.1000) oder höher | Januar 2016, 1.18 oder höher | Januar 2016, 15.19 oder höher| September 2016 | |
| WordApi 1.1  | Version 1509 (Build 4266.1001) oder höher| Januar 2016, 1.18 oder höher | Januar 2016, 15.19 oder höher| September 2016 | |

> [!NOTE]
> Die Buildnummer für Office 2016 über MSI installiert ist 16.0.4266.1001. Diese Version enthält nur die WordApi 1.1 Anforderungssatz.

Weitere Informationen über Versionen, Buildnummern und Office Online Server finden Sie unter:

- [Versions- und Buildnummern der Updatekanalversionen für Office 365-Clients](https://support.office.com/article/version-and-build-numbers-of-update-channel-releases-ae942449-1fca-4484-898b-a933ea23def7)
- [Welche Version von Office verwende ich?](https://support.office.com/article/What-version-of-Office-am-I-using-932788b8-a3ce-44bf-bb09-e334518b8b19)
- [So finden Sie die Versions- und Bildnummer für eine Office 365-Clientanwendung](https://support.office.com/article/version-and-build-numbers-of-update-channel-releases-ae942449-1fca-4484-898b-a933ea23def7)
- 
  [Office Online Server-Übersicht](https://docs.microsoft.com/officeonlineserver/office-online-server-overview)

## <a name="office-common-api-requirement-sets"></a>Allgemeine Office-API-Anforderungssätze

Informationen zu allgemeinen API-Anforderungssätzen finden Sie unter [Allgemeine Office-API-Anforderungssätze](office-add-in-requirement-sets.md).

## <a name="whats-new-in-word-javascript-api-13"></a>Neuigkeiten in der Word-JavaScript-API 1.3 

Im Folgenden werden die neuen Elemente der Word-JavaScript-APIs in Anforderungssatz 1.3 aufgeführt. 

|Objekt| Neuerungen| Beschreibung|Anforderungssatz| 
|:-----|-----|:----|:----| 
|[application](/javascript/api/word/word.application)|_Methode_ > createDocument(base64File: string) | Erstellt ein neues Dokument mithilfe einer base64-codierten DOCX-Datei. Schreibgeschützt.|1.3|
|[body](/javascript/api/word/word.body)|_Beziehung_ > lists|Ruft die Sammlung von Listenobjekten im Textkörper ab. Schreibgeschützt.|1.3|
|[body](/javascript/api/word/word.body)|_Beziehung_ > parentBody|Ruft den übergeordneten Text des Texts ab. Beispielsweise könnte der übergeordnete Text des Texts einer Tabellenzelle ein Header sein. Schreibgeschützt.|1.3|
|[body](/javascript/api/word/word.body)|_Beziehung_ > parentSection|Ruft den übergeordneten Abschnitt des Texts ab. Schreibgeschützt.|1.3|
|[body](/javascript/api/word/word.body)|_Beziehung_ > styleBuiltIn|Ruft den Namen der integrierten Formatvorlage für den Text ab oder legt ihn fest. Verwenden Sie diese Eigenschaft für integrierte Formatvorlagen, die zwischen Gebietsschemas portabel sind. Informationen zum Verwenden benutzerdefinierter Formatvorlagen oder lokalisierter Namen finden Sie unter der Eigenschaft "style".|1.3|
|[body](/javascript/api/word/word.body)|_Beziehung_ > tables|Ruft die Sammlung von Tabellenobjekten im Text ab. Schreibgeschützt.|1.3|
|[body](/javascript/api/word/word.body)|_Beziehung_ > type|Ruft den Typ des Texts ab. Der Typ kann "MainDoc", "Section", "Header", "Footer" oder "TableCell" sein. Schreibgeschützt.|1.3|
|[body](/javascript/api/word/word.body)|_Methode_ > GetRange(rangeLocation: RangeLocation)|Ruft den gesamten Text oder den Start- bzw. Endpunkt des Texts als Bereich ab.|1.3|
|[body](/javascript/api/word/word.body)|_Methode_ > InsertTable (RowCount: Number, ColumnCount: Number, InsertLocation: InsertLocation Werte: Zeichenfolge)|Fügt eine Tabelle mit der angegebenen Anzahl von Zeilen und Spalten ein. Der insertLocation-Wert kann "Start" oder "End" lauten.|1.3|
|[breaktype](/javascript/api/word/word.breaktype)|_Beziehung_ > Breaks|Gibt die Form einer Pause: Zeile, eine Seite oder im Abschnittstyp. Schreibgeschützt.|1.3|
|[contentControl](/javascript/api/word/word.contentcontrol)|_Beziehung_ > lists|Ruft die Sammlung von Listenobjekten im Inhaltssteuerelement ab. Schreibgeschützt.|1.3|
|[contentControl](/javascript/api/word/word.contentcontrol)|_Beziehung_ > parentBody|Ruft den übergeordneten Text des Inhaltssteuerelements ab. Schreibgeschützt.|1.3|
|[contentControl](/javascript/api/word/word.contentcontrol)|_Beziehung_ > parentTable|Ruft die Tabelle ab, die das Inhaltssteuerelement enthält. Gibt ein NULL-Objekt zurück, wenn es nicht in einer Tabelle enthalten ist. Schreibgeschützt.|1.3|
|[contentControl](/javascript/api/word/word.contentcontrol)|_Beziehung_ > parentTableCell|Ruft die Tabellenzelle ab, die das Inhaltssteuerelement enthält. Gibt ein NULL-Objekt zurück, wenn es nicht in einer Tabellenzelle enthalten ist. Schreibgeschützt.|1.3|
|[contentControl](/javascript/api/word/word.contentcontrol)|_Beziehung_ > styleBuiltIn|Ruft den Namen der integrierten Formatvorlage für das Inhaltssteuerelement ab oder legt ihn fest. Verwenden Sie diese Eigenschaft für integrierte Formatvorlagen, die zwischen Gebietsschemas portabel sind. Informationen zum Verwenden benutzerdefinierter Formatvorlagen oder lokalisierter Namen finden Sie unter der Eigenschaft "style".|1.3|
|[contentControl](/javascript/api/word/word.contentcontrol)|_Beziehung_ > subtype|Ruft den Untertyp des Inhaltssteuerelements ab. Der Untertyp kann für Rich-Text-Inhaltssteuerelemente "RichTextInline", "RichTextParagraphs", "RichTextTableCell", "RichTextTableRow" und "RichTextTable" sein. Schreibgeschützt.|1.3|
|[contentControl](/javascript/api/word/word.contentcontrol)|_Beziehung_ > tables|Ruft die Sammlung von Tabellenobjekten im Inhaltssteuerelement ab. Schreibgeschützt.|1.3|
|[contentControl](/javascript/api/word/word.contentcontrol)|_Methode_ > GetRange(rangeLocation: RangeLocation)|Ruft das gesamte Inhaltssteuerelement oder den Start- bzw. Endpunkt des Inhaltssteuerelements als Bereich ab.|1.3|
|[contentControl](/javascript/api/word/word.contentcontrol)|_Methode_ > GetTextRanges (EndingMarks: String, TrimSpacing: Bool)|Ruft die Textbereiche im Inhaltssteuerelement mithilfe von Satzzeichen und/oder anderen Endmarkierungen ab.|1.3|
|[contentControl](/javascript/api/word/word.contentcontrol)|_Methode_ > InsertTable (RowCount: Number, ColumnCount: Number, InsertLocation: InsertLocation Werte: Zeichenfolge)|Fügt eine Tabelle mit der angegebenen Anzahl von Zeilen und Spalten in oder neben ein Inhaltssteuerelement ein. Der insertLocation-Wert kann "Start", "End", "Before" oder "After" lauten.|1.3|
|[contentControl](/javascript/api/word/word.contentcontrol)|_Methode_ > Teilen (Trennzeichen: string, MultiParagraphs: Bool TrimDelimiters: Bool TrimSpacing: Bool)|Teilt das Inhaltssteuerelement mithilfe von Trennzeichen in untergeordnete Bereiche.|1.3|
|[contentControlCollection](/javascript/api/word/word.contentcontrolcollection)|_Methode_ > GetByTypes(types: ContentControlType)|Ruft die Inhaltssteuerelemente mit den angegebenen Typen und/oder Untertypen ab.|1.3|
|[contentControlCollection](/javascript/api/word/word.contentcontrolcollection)|_Methode_ > getFirst()|Ruft das erste Inhaltssteuerelement in dieser Sammlung ab.|1.3|
|[customProperty](/javascript/api/word/word.customproperty)|_Eigenschaft_ > key|Ruft den Schlüssel der benutzerdefinierten Eigenschaft. Schreibgeschützt. |1.3|
|[customProperty](/javascript/api/word/word.customproperty)|_Eigenschaft_ > value|Ruft den Wert der benutzerdefinierten Eigenschaft ab oder legt ihn fest.|1.3|
|[customProperty](/javascript/api/word/word.customproperty)|_Beziehung_ > type|Ruft den Werttyp der benutzerdefinierten Eigenschaft. Schreibgeschützt.|1.3|
|[customProperty](/javascript/api/word/word.customproperty)|_Methode_ > delete()|Löscht die benutzerdefinierte Eigenschaft.|1.3|
|[customPropertyCollection](/javascript/api/word/word.custompropertycollection)|_Eigenschaft_ > items|Eine Sammlung von customProperty-Objekten. Schreibgeschützt.|1.3|
|[customPropertyCollection](/javascript/api/word/word.custompropertycollection)|_Methode_ > deleteAll()|Löscht alle benutzerdefinierten Eigenschaften in dieser Sammlung.|1.3|
|[customPropertyCollection](/javascript/api/word/word.custompropertycollection)|_Methode_ > getCount()|Ruft die Anzahl von benutzerdefinierten Eigenschaften ab.|1.3|
|[customPropertyCollection](/javascript/api/word/word.custompropertycollection)|_Methode_ > GetItem(key: string)|Ruft ein Objekt für benutzerdefinierte Eigenschaften über seinen Schlüssel ab, bei dem Groß-/Kleinschreibung nicht beachtet wird.|1.3|
|[customPropertyCollection](/javascript/api/word/word.custompropertycollection)|_Methode_ > festgelegt (Schlüssel: string-Wert: Objekt)|Erstellt eine benutzerdefinierte Eigenschaft oder legt sie fest.|1.3|
|[document](/javascript/api/word/word.document)|_Beziehung_ > properties|Ruft die Eigenschaften des aktuellen Dokuments ab. Schreibgeschützt.|1.3|
|[document](/javascript/api/word/word.document)|_Methode_ > open()|Öffnet das Dokument.|1.3|
|[documentProperties](/javascript/api/word/word.documentproperties)|_Eigenschaft_ > applicationName|Ruft den Anwendungsnamen des Dokuments ab. Schreibgeschützt.|1.3|
|[documentProperties](/javascript/api/word/word.documentproperties)|_Eigenschaft_ > author|Ruft den Autor des Dokuments ab oder legt ihn fest.|1.3|
|[documentProperties](/javascript/api/word/word.documentproperties)|_Eigenschaft_ > category|Ruft die Kategorie des Dokuments ab oder legt sie fest.|1.3|
|[documentProperties](/javascript/api/word/word.documentproperties)|_Eigenschaft_ > comments|Ruft die Kommentare des Dokuments ab oder legt sie fest.|1.3|
|[documentProperties](/javascript/api/word/word.documentproperties)|_Eigenschaft_ > company|Ruft das Unternehmen des Dokuments ab oder legt es fest.|1.3|
|[documentProperties](/javascript/api/word/word.documentproperties)|_Eigenschaft_ > format|Ruft das Format des Dokuments ab oder legt es fest.|1.3|
|[documentProperties](/javascript/api/word/word.documentproperties)|_Eigenschaft_ > keywords|Ruft die Schlüsselwörter des Dokuments ab oder legt sie fest.|1.3|
|[documentProperties](/javascript/api/word/word.documentproperties)|_Eigenschaft_ > lastAuthor|Ruft den letzten Autor des Dokuments ab oder legt ihn fest.|1.3|
|[documentProperties](/javascript/api/word/word.documentproperties)|_Eigenschaft_ > manager|Ruft den Manager des Dokuments ab oder legt ihn fest.|1.3|
|[documentProperties](/javascript/api/word/word.documentproperties)|_Eigenschaft_ > revisionNumber|Ruft die Revisionsnummer des Dokuments ab. Schreibgeschützt.|1.3|
|[documentProperties](/javascript/api/word/word.documentproperties)|_Eigenschaft_ > security|Ruft die Sicherheit des Dokuments ab. Schreibgeschützt.|1.3|
|[documentProperties](/javascript/api/word/word.documentproperties)|_Eigenschaft_ > subject|Ruft den Betreff des Dokuments ab oder legt ihn fest.|1.3|
|[documentProperties](/javascript/api/word/word.documentproperties)|_Eigenschaft_ > template|Ruft die Vorlage des Dokuments ab. Schreibgeschützt.|1.3|
|[documentProperties](/javascript/api/word/word.documentproperties)|_Eigenschaft_ > title|Ruft den Titel des Dokuments ab oder legt ihn fest.|1.3|
|[documentProperties](/javascript/api/word/word.documentproperties)|_Beziehung_ > creationDate|Ruft das Erstellungsdatum des Dokuments ab. Schreibgeschützt.|1.3|
|[documentProperties](/javascript/api/word/word.documentproperties)|_Beziehung_ > customProperties|Ruft die Auflistung der benutzerdefinierten Eigenschaften des Dokuments schreibgeschützt.|1.3|
|[documentProperties](/javascript/api/word/word.documentproperties)|_Beziehung_ > lastPrintDate|Ruft das Datum der letzten Drucken des Dokuments ab. Schreibgeschützt.|1.3|
|[documentProperties](/javascript/api/word/word.documentproperties)|_Beziehung_ > lastSaveTime|Ruft ab dem letzten Speichern des Dokuments. Schreibgeschützt.|1.3|
|[inlinePicture](/javascript/api/word/word.inlinepicture)|_Beziehung_ > parentTable|Ruft die Tabelle ab, die das Inlinebild enthält. Gibt ein NULL-Objekt zurück, wenn es nicht in einer Tabelle enthalten ist. Schreibgeschützt.|1.3|
|[inlinePicture](/javascript/api/word/word.inlinepicture)|_Beziehung_ > parentTableCell|Ruft die Tabellenzelle ab, die das Inlinebild enthält. Gibt ein NULL-Objekt zurück, wenn es nicht in einer Tabellenzelle enthalten ist. Schreibgeschützt.|1.3|
|[inlinePicture](/javascript/api/word/word.inlinepicture)|_Methode_ > getNext()|Ruft das nächste Inlinebild ab.|1.3|
|[inlinePicture](/javascript/api/word/word.inlinepicture)|_Methode_ > GetRange(rangeLocation: RangeLocation)|Ruft das gesamte Bild oder den Start- bzw. Endpunkt des Bilds als Bereich ab.|1.3|
|[inlinePictureCollection](/javascript/api/word/word.inlinepicturecollection)|_Methode_ > getFirst()|Ruft das erste Inlinebild in dieser Sammlung ab.|1.3|
|[list](/javascript/api/word/word.list)|_Eigenschaft_ > id|Ruft die ID der Liste ab. Schreibgeschützt.|1.3|
|[list](/javascript/api/word/word.list)|_Eigenschaft_ > levelExistences|Überprüft, ob jede der 9 Ebenen in der Liste vorhanden ist. Der Wert true gibt an, dass die Ebene vorhanden ist, was bedeutet, dass mindestens ein Listenelement auf dieser Ebene vorhanden ist. Schreibgeschützt.|1.3|
|[list](/javascript/api/word/word.list)|_Beziehung_ > levelTypes|Ruft die Typen aller 9 Ebenen in der Liste ab. Jeder Typ kann "Bullet", "Number" oder "Picture" sein. Schreibgeschützt.|1.3|
|[list](/javascript/api/word/word.list)|_Beziehung_ > paragraphs|Ruft die Absätze in der Liste ab. Schreibgeschützt.|1.3|
|[list](/javascript/api/word/word.list)|_Methode_ > GetLevelParagraphs(level: number)|Ruft die Absätze ab, die auf der angegebenen Ebene in der Liste vorkommen.|1.3|
|[list](/javascript/api/word/word.list)|_Methode_ > GetLevelString(level: number)|Ruft das Aufzählungszeichen, die Zahl oder das Bild auf der angegebenen Ebene als Zeichenfolge ab.|1.3|
|[list](/javascript/api/word/word.list)|_Methode_ > InsertParagraph (ParagraphText: String, InsertLocation: InsertLocation)|Fügt an der angegebenen Position einen Absatz ein. Der insertLocation-Wert kann "Start", "End", "Before" oder "After" lauten.|1.3|
|[list](/javascript/api/word/word.list)|_Methode_ > SetLevelAlignment (Ebene: number, Ausrichtung: Ausrichtung)|Legt die Ausrichtung des Aufzählungszeichens, der Zahl oder des Bilds auf der angegebenen Ebene fest.|1.3|
|[list](/javascript/api/word/word.list)|_Methode_ > SetLevelBullet (Level: Anzahl, ListBullet: ListBullet Zeichencode: Number, FontName: Zeichenfolge)|Legt das Format des Aufzählungszeichens auf der angegebenen Ebene in der Liste fest. Wenn das Aufzählungszeichen vom Typ "Custom" ist, ist charCode erforderlich.|1.3|
|[list](/javascript/api/word/word.list)|_Methode_ > SetLevelIndents (Level: Anzahl, TextIndent: Float, TextIndent: Float)|Legt die zwei Einzüge der angegebenen Ebene in der Liste fest.|1.3|
|[list](/javascript/api/word/word.list)|_Methode_ > SetLevelNumbering (Level: Anzahl, ListNumbering: ListNumbering FormatString: Objekt)|Legt das Format der Nummerierung auf der angegebenen Ebene in der Liste fest.|1.3|
|[list](/javascript/api/word/word.list)|_Methode_ > SetLevelStartingNumber (Level: Anzahl, StartingNumber: Anzahl)|Legt die Anfangsnummer auf der angegebenen Ebene in der Liste fest. Der Standardwert ist 1.|1.3|
|[listCollection](/javascript/api/word/word.listcollection)|_Eigenschaft_ > items|Eine Sammlung von Listenobjekten. Schreibgeschützt.|1.3|
|[listCollection](/javascript/api/word/word.listcollection)|_Methode_ > GetById(id: number)|Ruft eine Liste über ihren Bezeichner ab.|1.3|
|[listCollection](/javascript/api/word/word.listcollection)|_Methode_ > getFirst()|Ruft die erste Liste in dieser Sammlung ab.|1.3|
|[listCollection](/javascript/api/word/word.listcollection)|_Methode_ > GetItem(index: number)|Ruft ein Listenobjekt nach dem Index in der Auflistung ab.|1.3|
|[listItem](/javascript/api/word/word.listitem)|_Eigenschaft_ > level|Ruft die Ebene des Elements in der Liste ab oder legt sie fest.|1.3|
|[listItem](/javascript/api/word/word.listitem)|_Eigenschaft_ > listString|Ruft das Aufzählungszeichen, die Zahl oder das Bild des Listenelements als Zeichenfolge ab. Schreibgeschützt.|1.3|
|[listItem](/javascript/api/word/word.listitem)|_Eigenschaft_ > siblingIndex|Ruft die Ordnungsnummer des Listenelements im Verhältnis zu seinen gleichgeordneten Elementen ab. Schreibgeschützt.|1.3|
|[listItem](/javascript/api/word/word.listitem)|_Methode_ > GetAncestor(parentOnly: bool)|Ruft das übergeordnete Element der Liste bzw. den nächsten Vorgänger ab, wenn das übergeordnete Element nicht vorhanden ist.|1.3|
|[listItem](/javascript/api/word/word.listitem)|_Methode_ > GetDescendants(directChildrenOnly: bool)|Ruft alle untergeordneten Listenelemente des Listenelements ab.|1.3|
|[paragraph](/javascript/api/word/word.paragraph)|_Eigenschaft_ > isLastParagraph|Gibt an, dass es sich bei dem Absatz um den letzten innerhalb seines übergeordneten Texts handelt. Schreibgeschützt.|1.3|
|[paragraph](/javascript/api/word/word.paragraph)|_Eigenschaft_ > isListItem|Überprüft, ob der Absatz ein Listenelement ist. Schreibgeschützt.|1.3|
|[paragraph](/javascript/api/word/word.paragraph)|_Eigenschaft_ > tableNestingLevel|Ruft die Ebene der Tabelle des Absatzes ab. Gibt 0 zurück, wenn sich der Absatz nicht in einer Tabelle befindet. Schreibgeschützt.|1.3|
|[paragraph](/javascript/api/word/word.paragraph)|_Beziehung_ > list|Ruft die Liste ab, zu der dieser Absatz gehört. Gibt ein NULL-Objekt zurück, wenn sich der Absatz nicht in einer Liste befindet. Schreibgeschützt.|1.3|
|[paragraph](/javascript/api/word/word.paragraph)|_Beziehung_ > listItem|Ruft das ListItem für den Absatz ab. Gibt ein NULL-Objekt zurück, wenn der Absatz nicht Teil einer Liste ist. Schreibgeschützt.|1.3|
|[paragraph](/javascript/api/word/word.paragraph)|_Beziehung_ > parentBody|Ruft den übergeordneten Text des Absatzes ab. Schreibgeschützt.|1.3|
|[paragraph](/javascript/api/word/word.paragraph)|_Beziehung_ > parentTable|Ruft die Tabelle ab, die den Absatz enthält. Gibt ein NULL-Objekt zurück, wenn es nicht in einer Tabelle enthalten ist. Schreibgeschützt.|1.3|
|[paragraph](/javascript/api/word/word.paragraph)|_Beziehung_ > parentTableCell|Ruft die Tabellenzelle ab, die den Absatz enthält. Gibt ein NULL-Objekt zurück, wenn es nicht in einer Tabellenzelle enthalten ist. Schreibgeschützt.|1.3|
|[paragraph](/javascript/api/word/word.paragraph)|_Beziehung_ > styleBuiltIn|Ruft den Namen der integrierten Formatvorlage für den Absatz ab oder legt ihn fest. Verwenden Sie diese Eigenschaft für integrierte Formatvorlagen, die zwischen Gebietsschemas portabel sind. Informationen zum Verwenden benutzerdefinierter Formatvorlagen oder lokalisierter Namen finden Sie unter der Eigenschaft "style".|1.3|
|[paragraph](/javascript/api/word/word.paragraph)|_Methode_ > AttachToList (ListId: number, Level: Anzahl)|Fügt den Absatz zu einer vorhandenen Liste auf der angegebenen Ebene hinzu. Führt zu einem Fehler, wenn der Absatz nicht in die Liste aufgenommen werden kann oder der Absatz bereits ein Listenelement ist.|1.3|
|[paragraph](/javascript/api/word/word.paragraph)|_Methode_ > detachFromList()|Verschiebt diesen Absatz aus der Liste, falls der Absatz ein Listenelement ist.|1.3|
|[paragraph](/javascript/api/word/word.paragraph)|_Methode_ > getNext()|Ruft den nächsten Absatz ab.|1.3|
|[paragraph](/javascript/api/word/word.paragraph)|_Methode_ > getPrevious()|Ruft den vorherigen Absatz ab.|1.3|
|[paragraph](/javascript/api/word/word.paragraph)|_Methode_ > GetRange(rangeLocation: RangeLocation)|Ruft den gesamten Absatz oder den Start- bzw. Endpunkt des Absatzes als Bereich ab.|1.3|
|[paragraph](/javascript/api/word/word.paragraph)|_Methode_ > GetTextRanges (EndingMarks: String, TrimSpacing: Bool)|Ruft die Textbereiche im Absatz mithilfe von Satzzeichen und/oder anderen Endmarkierungen ab.|1.3|
|[paragraph](/javascript/api/word/word.paragraph)|_Methode_ > InsertTable (RowCount: Number, ColumnCount: Number, InsertLocation: InsertLocation Werte: Zeichenfolge)|Fügt eine Tabelle mit der angegebenen Anzahl von Zeilen und Spalten ein. Der insertLocation-Wert kann "Before" oder "After" lauten.|1.3|
|[paragraph](/javascript/api/word/word.paragraph)|_Methode_ > Teilen (Trennzeichen: string, TrimDelimiters: Bool TrimSpacing: Bool)|Teilt den Absatz mithilfe von Trennzeichen in untergeordnete Bereiche.|1.3|
|[paragraph](/javascript/api/word/word.paragraph)|_Methode_ > startNewList()|Beginnt eine neue Liste mit diesem Absatz. Führt zu einem Fehler, wenn der Absatz bereits ein Listenelement ist.|1.3|
|[paragraphCollection](/javascript/api/word/word.paragraphcollection)|_Methode_ > getFirst()|Ruft den ersten Absatz in dieser Sammlung ab.|1.3|
|[paragraphCollection](/javascript/api/word/word.paragraphcollection)|_Methode_ > getLast()|Ruft den letzten Absatz in dieser Sammlung ab.|1.3|
|[range](/javascript/api/word/word.range)|_Eigenschaft_ > hyperlink|Ruft den ersten Link in dem Bereich ab oder legt einen Link für den Bereich fest. Wenn Sie einen neuen Link für den Bereich festlegen, werden alle Links im Bereich gelöscht. Verwenden Sie ein Zeilenumbruchzeichen ('\n'), um den Adressteil vom optionalen Speicherortsteil zu trennen.|1.3|
|[range](/javascript/api/word/word.range)|_Eigenschaft_ > isEmpty|Überprüft, ob die Länge des Bereichs 0 ist. Schreibgeschützt.|1.3|
|[range](/javascript/api/word/word.range)|_Beziehung_ > lists|Ruft die Sammlung von Listenobjekten im Bereich ab. Schreibgeschützt.|1.3|
|[range](/javascript/api/word/word.range)|_Beziehung_ > parentBody|Ruft den übergeordneten Text des Bereichs ab. Schreibgeschützt.|1.3|
|[range](/javascript/api/word/word.range)|_Beziehung_ > parentTable|Ruft die Tabelle ab, die den Bereich enthält. Gibt NULL zurück, wenn er nicht in einer Tabelle enthalten ist. Schreibgeschützt.|1.3|
|[range](/javascript/api/word/word.range)|_Beziehung_ > parentTableCell|Ruft die Tabellenzelle ab, die den Bereich enthält. Gibt ein NULL-Objekt zurück, wenn es nicht in einer Tabellenzelle enthalten ist. Schreibgeschützt.|1.3|
|[range](/javascript/api/word/word.range)|_Beziehung_ > styleBuiltIn|Ruft den Namen der integrierten Formatvorlage für den Bereich ab oder legt ihn fest. Verwenden Sie diese Eigenschaft für integrierte Formatvorlagen, die zwischen Gebietsschemas portabel sind. Informationen zum Verwenden benutzerdefinierter Formatvorlagen oder lokalisierter Namen finden Sie unter der Eigenschaft "style".|1.3|
|[range](/javascript/api/word/word.range)|_Beziehung_ > tables|Ruft die Sammlung von Tabellenobjekten im Bereich ab. Schreibgeschützt.|1.3|
|[range](/javascript/api/word/word.range)|_Methode_ > CompareLocationWith(range: Range)|Vergleicht die Position dieses Bereichs mit der eines anderen Bereichs.|1.3|
|[range](/javascript/api/word/word.range)|_Methode_ > ExpandTo(range: Range)|Gibt einen neuen Bereich zurück, der diesen Bereich in beide Richtungen erweitert, um einen anderen Bereich zu überdecken. Dieser Bereich wird nicht geändert.|1.3|
|[range](/javascript/api/word/word.range)|_Methode_ > getHyperlinkRanges()|Ruft untergeordnete Linkbereiche innerhalb des Bereichs ab.|1.3|
|[range](/javascript/api/word/word.range)|_Methode_ > GetNextTextRange (EndingMarks: String, TrimSpacing: Bool)|Ruft die nächsten Textbereiche mithilfe von Satzzeichen und/oder anderen Endmarkierungen ab.|1.3|
|[range](/javascript/api/word/word.range)|_Methode_ > GetRange(rangeLocation: RangeLocation)|Klont den Bereich oder ruft den Anfangs- bzw. Endpunkt des Bereichs als einen neuen Bereich ab.|1.3|
|[range](/javascript/api/word/word.range)|_Methode_ > GetTextRanges (EndingMarks: String, TrimSpacing: Bool)|Ruft die untergeordneten Textbereiche im Bereich mithilfe von Satzzeichen und/oder anderen Endmarkierungen ab.|1.3|
|[range](/javascript/api/word/word.range)|_Methode_ > InsertTable (RowCount: Number, ColumnCount: Number, InsertLocation: InsertLocation Werte: Zeichenfolge)|Fügt eine Tabelle mit der angegebenen Anzahl von Zeilen und Spalten ein. Der insertLocation-Wert kann "Before" oder "After" lauten.|1.3|
|[range](/javascript/api/word/word.range)|_Methode_ > IntersectWith(range: Range)|Gibt einen neuen Bereich als Schnittmenge dieses Bereichs mit einem anderen Bereich zurück. Dieser Bereich wird nicht geändert.|1.3|
|[range](/javascript/api/word/word.range)|_Methode_ > Teilen (Trennzeichen: string, MultiParagraphs: Bool TrimDelimiters: Bool TrimSpacing: Bool)|Teilt den Bereich mithilfe von Trennzeichen in untergeordnete Bereiche.|1.3|
|[rangeCollection](/javascript/api/word/word.rangecollection)|_Eigenschaft_ > items|Eine Sammlung von range-Objekten. Schreibgeschützt.|1.3|
|[rangeCollection](/javascript/api/word/word.rangecollection)|_Methode_ > getFirst()|Ruft den ersten Bereich in dieser Sammlung ab.|1.3|
|[rangeCollection](/javascript/api/word/word.rangecollection)|_Methode_ > GetItem(index: number)|Ruft ein Bereichsobjekt nach dem Index in der Auflistung ab.|1.3|
|[requestContext](/javascript/api/word/word.requestcontext)|_Methode_ > Laden (Objekt:-Objekts, option: Objekt)|Füllt das auf der JavaScript-Ebene erstellte Proxyobjekt mit der im Parameter angegebenen Eigenschaft und den Optionen. |1.3|
|[requestContext](/javascript/api/word/word.requestcontext)|_Methode_ > sync()|Sendet die Anforderungswarteschlange an Word, und gibt ein Zusageobjekt zurück, das zum Verketten weiterer Aktionen verwendet werden kann.|1.3|
|[section](/javascript/api/word/word.section)|_Methode_ > getNext()|Ruft den nächsten Abschnitt ab.|1.3|
|[sectionCollection](/javascript/api/word/word.sectioncollection)|_Methode_ > getFirst()|Ruft den ersten Abschnitt in dieser Sammlung ab.|1.3|
|[table](/javascript/api/word/word.table)|_Eigenschaft_ > headerRowCount|Ruft die Anzahl von Kopfzeilen ab oder legt sie fest.|1.3|
|[table](/javascript/api/word/word.table)|_Eigenschaft_ > height|Ruft die Höhe der Tabelle in Punkt ab. Schreibgeschützt.|1.3|
|[table](/javascript/api/word/word.table)|_Eigenschaft_ > isUniform|Gibt an, ob alle Zeilen der Tabelle uniform sind. Schreibgeschützt.|1.3|
|[table](/javascript/api/word/word.table)|_Eigenschaft_ > nestingLevel|Ruft die Schachtelungsebene der Tabelle ab. Tabellen der obersten Ebene haben Ebene 1. Schreibgeschützt.|1.3|
|[table](/javascript/api/word/word.table)|_Eigenschaft_ > rowCount|Ruft die Anzahl der Zeilen in der Tabelle ab. Schreibgeschützt.|1.3|
|[table](/javascript/api/word/word.table)|_Eigenschaft_ > shadingColor|Ruft die Schattierungsfarbe ab und legt diese fest.|1.3|
|[table](/javascript/api/word/word.table)|_Eigenschaft_ > style|Ruft den Namen der Formatvorlage für die Tabelle ab oder legt ihn fest. Verwenden Sie diese Eigenschaft für benutzerdefinierte Formatvorlagen und lokalisierte Formatvorlagennamen. Informationen zur Verwendung der integrierten Formatvorlagen, die zwischen Gebietsschemas portabel sind, finden Sie unter der Eigenschaft "styleBuiltIn".|1.3|
|[table](/javascript/api/word/word.table)|_Eigenschaft_ > styleBandedColumns|Ruft ab und legt fest, ob die Tabelle gebänderte Spalten enthält.|1.3|
|[table](/javascript/api/word/word.table)|_Eigenschaft_ > styleBandedRows|Ruft ab und legt fest, ob die Tabelle gebänderte Zeilen enthält.|1.3|
|[table](/javascript/api/word/word.table)|_Eigenschaft_ > styleFirstColumn|Ruft ab und legt fest, ob die Tabelle eine erste Spalte mit einer speziellen Formatvorlage aufweist.|1.3|
|[table](/javascript/api/word/word.table)|_Eigenschaft_ > styleLastColumn|Ruft ab und legt fest, ob die Tabelle eine letzte Spalte mit einer speziellen Formatvorlage aufweist.|1.3|
|[table](/javascript/api/word/word.table)|_Eigenschaft_ > styleTotalRow|Ruft ab und legt fest, ob die Tabelle eine Ergebniszeile (letzte Zeile) mit einer speziellen Formatvorlage aufweist.|1.3|
|[table](/javascript/api/word/word.table)|_Eigenschaft_ > values|Ruft die Textwerte in der Tabelle als 2D-Javascript-Array ab und legt sie fest.|1.3|
|[table](/javascript/api/word/word.table)|_Eigenschaft_ > width|Ruft die Breite der Tabelle in Punkt ab und legt sie fest.|1.3|
|[table](/javascript/api/word/word.table)|_Beziehung_ > font|Ruft die Schriftart ab. Verwenden Sie diese Option zum Abrufen und Festlegen des Schriftartnamens, der Größe, der Farbe und anderer Eigenschaften. Schreibgeschützt.|1.3|
|[table](/javascript/api/word/word.table)|_Beziehung_ > horizontalAlignment|Ruft die horizontale Ausrichtung jeder Zelle in der Tabelle ab und legt sie fest. Der Wert kann "left", "centered", "right" oder "justified" lauten.|1.3|
|[table](/javascript/api/word/word.table)|_Beziehung_ > paragraphAfter|Ruft den Absatz nach der Tabelle ab. Schreibgeschützt.|1.3|
|[table](/javascript/api/word/word.table)|_Beziehung_ > paragraphBefore|Ruft den Absatz vor der Tabelle ab. Schreibgeschützt.|1.3|
|[table](/javascript/api/word/word.table)|_Beziehung_ > parentBody|Ruft den übergeordneten Text der Tabelle ab. Schreibgeschützt.|1.3|
|[table](/javascript/api/word/word.table)|_Beziehung_ > parentContentControl|Ruft das Inhaltssteuerelement ab, das die Tabelle enthält. Schreibgeschützt.|1.3|
|[table](/javascript/api/word/word.table)|_Beziehung_ > parentTable|Ruft die Tabelle ab, die diese Tabelle enthält. Gibt ein NULL-Objekt zurück, wenn es nicht in einer Tabelle enthalten ist. Schreibgeschützt.|1.3|
|[table](/javascript/api/word/word.table)|_Beziehung_ > parentTableCell|Ruft die Tabellenzelle ab, die diese Tabelle enthält. Gibt ein NULL-Objekt zurück, wenn es nicht in einer Tabellenzelle enthalten ist. Schreibgeschützt.|1.3|
|[table](/javascript/api/word/word.table)|_Beziehung_ > rows|Ruft alle Zeilen der Tabelle ab. Schreibgeschützt.|1.3|
|[table](/javascript/api/word/word.table)|_Beziehung_ > styleBuiltIn|Ruft den Namen der integrierten Formatvorlage für die Tabelle ab oder legt ihn fest. Verwenden Sie diese Eigenschaft für integrierte Formatvorlagen, die zwischen Gebietsschemas portabel sind. Informationen zum Verwenden benutzerdefinierter Formatvorlagen oder lokalisierter Namen finden Sie unter der Eigenschaft "style".|1.3|
|[table](/javascript/api/word/word.table)|_Beziehung_ > tables|Ruft die untergeordneten Tabellen ab, die eine Ebene tiefer verschachtelt sind. Schreibgeschützt.|1.3|
|[table](/javascript/api/word/word.table)|_Beziehung_ > verticalAlignment|Ruft die vertikale Ausrichtung jeder Zelle in der Tabelle ab und legt sie fest. Der Wert kann "top", "center" oder "bottom" lauten.|1.3|
|[table](/javascript/api/word/word.table)|_Methode_ > AddColumns (InsertLocation: InsertLocation, ColumnCount:, Zahlenwerte: Zeichenfolge)|Fügt am Anfang oder Ende der Tabelle Spalten hinzu, wobei die erste oder letzte vorhandene Spalte als Vorlage verwendet wird. Dies gilt für einheitliche Tabellen. Falls Zeichenfolgenwerte angegeben werden, werden diese in den neu eingefügten Zeilen festgelegt.|1.3|
|[table](/javascript/api/word/word.table)|_Methode_ > AddRows (InsertLocation: InsertLocation, RowCount:, Zahlenwerte: Zeichenfolge)|Fügt am Anfang oder Ende der Tabelle Zeilen hinzu, wobei die erste oder letzte vorhandene Zeile als Vorlage verwendet wird. Falls Zeichenfolgenwerte angegeben werden, werden diese in den neu eingefügten Zeilen festgelegt.|1.3|
|[table](/javascript/api/word/word.table)|_Methode_ > autoFitContents()|Passt die Tabellenspalten automatisch der Breite ihrer Inhalte an.|1.3|
|[table](/javascript/api/word/word.table)|_Methode_ > autoFitWindow()|Passt die Tabellenspalten automatisch der Breite des Fensters an.|1.3|
|[table](/javascript/api/word/word.table)|_Methode_ > Löschvorgang()|Löscht den Inhalt der Tabelle.|1.3|
|[table](/javascript/api/word/word.table)|_Methode_ > delete()|Löscht die gesamte Tabelle.|1.3|
|[table](/javascript/api/word/word.table)|_Methode_ > DeleteColumns (ColumnIndex: Number, ColumnCount: Anzahl)|Löscht bestimmte Spalten. Dies gilt für einheitliche Tabellen.|1.3|
|[table](/javascript/api/word/word.table)|_Methode_ > DeleteRows (RowIndex: Number, RowCount: Anzahl)|Löscht bestimmte Zeilen.|1.3|
|[table](/javascript/api/word/word.table)|_Methode_ > distributeColumns()|Verteilt die Breite der Spalten gleichmäßig.|1.3|
|[table](/javascript/api/word/word.table)|_Methode_ > distributeRows()|Verteilt die Höhe der Zeilen gleichmäßig.|1.3|
|[table](/javascript/api/word/word.table)|_Methode_ > GetBorder(borderLocation: BorderLocation)|Ruft die Rahmenart für den angegebenen Rahmen ab.|1.3|
|[table](/javascript/api/word/word.table)|_Methode_ > GetCell (RowIndex: Number, CellIndex: Anzahl)|Ruft die Tabellenzelle in einer angegebenen Zeile und Spalte ab.|1.3|
|[table](/javascript/api/word/word.table)|_Methode_ > GetCellPadding(cellPaddingLocation: CellPaddingLocation)|Ruft den Textabstand in Punkt ab.|1.3|
|[table](/javascript/api/word/word.table)|_Methode_ > getNext()|Ruft die nächste Tabelle ab.|1.3|
|[table](/javascript/api/word/word.table)|_Methode_ > GetRange(rangeLocation: RangeLocation)|Ruft den Bereich ab, der diese Tabelle enthält, oder den Bereich am Anfang bzw. Ende der Tabelle.|1.3|
|[table](/javascript/api/word/word.table)|_Methode_ > insertContentControl()|Fügt ein Inhaltssteuerelement in der Tabelle ein.|1.3|
|[table](/javascript/api/word/word.table)|_Methode_ > InsertParagraph (ParagraphText: String, InsertLocation: InsertLocation)|Fügt an der angegebenen Position einen Absatz ein. Der insertLocation-Wert kann 'Before' oder 'After' lauten.|1.3|
|[table](/javascript/api/word/word.table)|_Methode_ > InsertTable (RowCount: Number, ColumnCount: Number, InsertLocation: InsertLocation Werte: Zeichenfolge)|Fügt eine Tabelle mit der angegebenen Anzahl von Zeilen und Spalten ein. Der insertLocation-Wert kann "Before" oder "After" lauten.|1.3|
|[table](/javascript/api/word/word.table)|_Methode_ > Search (SearchText: string, SearchOptions: ParamTypeStrings.SearchOptions)|Führt eine Suche mit den angegebenen SearchOptions im Suchbereich des Tabellenobjekts aus. Die Suchergebnisse sind eine Sammlung von Bereichsobjekten.|1.3|
|[table](/javascript/api/word/word.table)|_Methode_ > Select(selectionMode: SelectionMode)|Wählt die Tabelle oder die Position am Anfang oder Ende der Tabelle aus und navigiert die Word-Benutzeroberfläche an diese Position.|1.3|
|[table](/javascript/api/word/word.table)|_Methode_ > SetCellPadding (CellPaddingLocation: CellPaddingLocation CellPadding: Float)|Legt den Textabstand in Punkt fest.|1.3|
|[tableBorder](/javascript/api/word/word.tableborder)|_Eigenschaft_ > color|Ruft die Farbe des Tabellenrahmens als Hexadezimalwert oder Namen ab oder legt sie fest.|1.3|
|[tableBorder](/javascript/api/word/word.tableborder)|_Eigenschaft_ > width|Ruft die Breite des Tabellenrahmens in Punkt ab oder legt sie fest. Nicht anwendbar auf Tabellenrahmentypen mit fester Breite.|1.3|
|[tableBorder](/javascript/api/word/word.tableborder)|_Beziehung_ > type|Ruft den Typ des Tabellenrahmens ab oder legt ihn fest.|1.3|
|[tableCell](/javascript/api/word/word.tablecell)|_Eigenschaft_ > cellIndex|Ruft den Index der Zelle in der Zeile ab. Schreibgeschützt.|1.3|
|[tableCell](/javascript/api/word/word.tablecell)|_Eigenschaft_ > columnWidth|Ruft die Breite der Spalte der Zelle in Punkt ab und legt sie fest. Dies gilt für einheitliche Tabellen.|1.3|
|[tableCell](/javascript/api/word/word.tablecell)|_Eigenschaft_ > rowIndex|Ruft den Index der Zeile mit der Zelle in der Tabelle ab. Schreibgeschützt.|1.3|
|[tableCell](/javascript/api/word/word.tablecell)|_Eigenschaft_ > shadingColor|Ruft die Schattierungsfarbe der Zelle ab oder legt diese fest. Die Farbe wird im Format "#RRGGBB" oder mit dem Namen der Farbe angegeben.|1.3|
|[tableCell](/javascript/api/word/word.tablecell)|_Eigenschaft_ > value|Ruft den Text der Zelle ab und legt ihn fest.|1.3|
|[tableCell](/javascript/api/word/word.tablecell)|_Eigenschaft_ > width|Ruft die Breite der Zelle in Punkt ab. Schreibgeschützt.|1.3|
|[tableCell](/javascript/api/word/word.tablecell)|_Beziehung_ > body|Ruft das body-Objekt der Zelle ab. Schreibgeschützt.|1.3|
|[tableCell](/javascript/api/word/word.tablecell)|_Beziehung_ > horizontalAlignment|Ruft die horizontale Ausrichtung der Zelle ab und legt sie fest. Der Wert kann "left", "centered", "right" oder "justified" lauten.|1.3|
|[tableCell](/javascript/api/word/word.tablecell)|_Beziehung_ > parentRow|Ruft die übergeordnete Zeile der Zelle ab. Schreibgeschützt.|1.3|
|[tableCell](/javascript/api/word/word.tablecell)|_Beziehung_ > parentTable|Ruft die übergeordnete Tabelle der Zelle ab. Schreibgeschützt.|1.3|
|[tableCell](/javascript/api/word/word.tablecell)|_Beziehung_ > verticalAlignment|Ruft die vertikale Ausrichtung der Zelle ab und legt sie fest. Der Wert kann "top", "center" oder "bottom" lauten.|1.3|
|[tableCell](/javascript/api/word/word.tablecell)|_Methode_ > deleteColumn()|Löscht die Spalte, die diese Zelle enthält. Dies gilt für einheitliche Tabellen.|1.3|
|[tableCell](/javascript/api/word/word.tablecell)|_Methode_ > deleteRow()|Löscht die Zeile, die diese Zelle enthält.|1.3|
|[tableCell](/javascript/api/word/word.tablecell)|_Methode_ > GetBorder(borderLocation: BorderLocation)|Ruft die Rahmenart für den angegebenen Rahmen ab.|1.3|
|[tableCell](/javascript/api/word/word.tablecell)|_Methode_ > GetCellPadding(cellPaddingLocation: CellPaddingLocation)|Ruft den Textabstand in Punkt ab.|1.3|
|[tableCell](/javascript/api/word/word.tablecell)|_Methode_ > getNext()|Ruft die nächste Zelle ab.|1.3|
|[tableCell](/javascript/api/word/word.tablecell)|_Methode_ > InsertColumns (InsertLocation: InsertLocation, ColumnCount:, Zahlenwerte: Zeichenfolge)|Fügt links oder rechts von der Zelle Spalten hinzu, wobei die Spalte der Zelle als Vorlage verwendet wird. Dies gilt für einheitliche Tabellen. Falls Zeichenfolgenwerte angegeben werden, werden diese in den neu eingefügten Zeilen festgelegt.|1.3|
|[tableCell](/javascript/api/word/word.tablecell)|_Methode_ > InsertRows (InsertLocation: InsertLocation, RowCount:, Zahlenwerte: Zeichenfolge)|Fügt oberhalb oder unterhalb der Zelle Zeilen hinzu, wobei die Zeile der Zelle als Vorlage verwendet wird. Falls Zeichenfolgenwerte angegeben werden, werden diese in den neu eingefügten Zeilen festgelegt.|1.3|
|[tableCell](/javascript/api/word/word.tablecell)|_Methode_ > SetCellPadding (CellPaddingLocation: CellPaddingLocation CellPadding: Float)|Legt den Textabstand in Punkt fest.|1.3|
|[tableCellCollection](/javascript/api/word/word.tablecellcollection)|_Eigenschaft_ > items|Eine Sammlung von tableCell-Objekten. Schreibgeschützt.|1.3|
|[tableCellCollection](/javascript/api/word/word.tablecellcollection)|_Methode_ > getFirst()|Ruft die erste Tabellenzelle in dieser Sammlung ab.|1.3|
|[tableCellCollection](/javascript/api/word/word.tablecellcollection)|_Methode_ > GetItem(index: number)|Ruft ein Tabellenzellenobjekt nach seinem Index in der Sammlung ab.|1.3|
|[tableCollection](/javascript/api/word/word.tablecollection)|_Eigenschaft_ > items|Eine Auflistung von Tabellenobjekten. Schreibgeschützt.|1.3|
|[tableCollection](/javascript/api/word/word.tablecollection)|_Methode_ > getFirst()|Ruft die erste Tabelle in dieser Sammlung ab.|1.3|
|[tableCollection](/javascript/api/word/word.tablecollection)|_Methode_ > GetItem(index: number)|Ruft ein Tabellenobjekt nach seinem Index in der Sammlung ab.|1.3|
|[tableRow](/javascript/api/word/word.tablerow)|_Eigenschaft_ > cellCount|Ruft die Anzahl der Zellen in der Zeile ab. Schreibgeschützt.|1.3|
|[tableRow](/javascript/api/word/word.tablerow)|_Eigenschaft_ > isHeader|Überprüft, ob die Zeile eine Kopfzeile ist. Schreibgeschützt. Zum Festlegen der Anzahl der Überschriftenzeilen verwenden Sie HeaderRowCount im Table-Objekt. Schreibgeschützt.|1.3|
|[tableRow](/javascript/api/word/word.tablerow)|_Eigenschaft_ > preferredHeight|Ruft die bevorzugte Höhe der Zeile in Punkt ab und legt sie fest.|1.3|
|[tableRow](/javascript/api/word/word.tablerow)|_Eigenschaft_ > rowIndex|Ruft den Index der Zeile in der übergeordneten Tabelle ab. Schreibgeschützt.|1.3|
|[tableRow](/javascript/api/word/word.tablerow)|_Eigenschaft_ > shadingColor|Ruft die Schattierungsfarbe ab und legt diese fest.|1.3|
|[tableRow](/javascript/api/word/word.tablerow)|_Eigenschaft_ > values|Ruft die Textwerte in der Zeile als 1D-Javascript-Array ab und legt sie fest.|1.3|
|[tableRow](/javascript/api/word/word.tablerow)|_Beziehung_ > cells|Ruft Zellen ab. Schreibgeschützt.|1.3|
|[tableRow](/javascript/api/word/word.tablerow)|_Beziehung_ > font|Ruft die Schriftart ab. Verwenden Sie diese Option zum Abrufen und Festlegen des Schriftartnamens, der Größe, der Farbe und anderer Eigenschaften. Schreibgeschützt.|1.3|
|[tableRow](/javascript/api/word/word.tablerow)|_Beziehung_ > horizontalAlignment|Ruft die horizontale Ausrichtung jeder Zelle in der Zeile ab und legt sie fest. Der Wert kann "left", "centered", "right" oder "justified" lauten.|1.3|
|[tableRow](/javascript/api/word/word.tablerow)|_Beziehung_ > parentTable|Ruft die übergeordnete Tabelle ab. Schreibgeschützt.|1.3|
|[tableRow](/javascript/api/word/word.tablerow)|_Beziehung_ > verticalAlignment|Ruft die vertikale Ausrichtung der Zellen in der Zeile ab und legt sie fest. Der Wert kann "top", "center" oder "bottom" lauten.|1.3|
|[tableRow](/javascript/api/word/word.tablerow)|_Methode_ > Löschvorgang()|Löscht den Inhalt der Zeile.|1.3|
|[tableRow](/javascript/api/word/word.tablerow)|_Methode_ > delete()|Löscht die gesamte Zeile.|1.3|
|[tableRow](/javascript/api/word/word.tablerow)|_Methode_ > GetBorder(borderLocation: BorderLocation)|Ruft die Rahmenart der Zellen in der Zeile ab.|1.3|
|[tableRow](/javascript/api/word/word.tablerow)|_Methode_ > GetCellPadding(cellPaddingLocation: CellPaddingLocation)|Ruft den Textabstand in Punkt ab.|1.3|
|[tableRow](/javascript/api/word/word.tablerow)|_Methode_ > getNext()|Ruft die nächste Zeile ab.|1.3|
|[tableRow](/javascript/api/word/word.tablerow)|_Methode_ > InsertRows (InsertLocation: InsertLocation, RowCount:, Zahlenwerte: Zeichenfolge)|Fügt Zeilen mit dieser Zeile als Vorlage ein. Wenn Werte angegeben sind, werden die Werte in die neuen Zeilen eingefügt.|1.3|
|[tableRow](/javascript/api/word/word.tablerow)|_Methode_ > Search (SearchText: string, SearchOptions: ParamTypeStrings.SearchOptions)|Führt eine Suche mit den angegebenen searchOptions im Suchbereich der Zeile aus. Die Suchergebnisse sind eine Sammlung von Bereichsobjekten.|1.3|
|[tableRow](/javascript/api/word/word.tablerow)|_Methode_ > Select(selectionMode: SelectionMode)|Wählt die Zeile aus und navigiert die Word-Benutzeroberfläche an diese Position.|1.3|
|[tableRow](/javascript/api/word/word.tablerow)|_Methode_ > SetCellPadding (CellPaddingLocation: CellPaddingLocation CellPadding: Float)|Legt den Textabstand in Punkt fest.|1.3|
|[tableRowCollection](/javascript/api/word/word.tablerowcollection)|_Eigenschaft_ > items|Eine Sammlung von tableRow-Objekten. Schreibgeschützt.|1.3|
|[tableRowCollection](/javascript/api/word/word.tablerowcollection)|_Methode_ > getFirst()|Ruft die erste Zeile in dieser Sammlung ab.|1.3|
|[tableRowCollection](/javascript/api/word/word.tablerowcollection)|_Methode_ > GetItem(index: number)|Ruft ein Tabellenzeilenobjekt nach seinem Index in der Sammlung ab.|1.3|


## <a name="whats-new-in-word-javascript-api-12"></a>Neuigkeiten in der Word-JavaScript-API 1.2

Im Folgenden werden die neuen Elemente der Word-JavaScript-APIs in Anforderungssatz 1.2 aufgeführt. 

|Objekt| Neuerungen| Beschreibung|Anforderungssatz|
|:-----|-----|:----|:----|
|[contentControl](/javascript/api/word/word.contentcontrol)|_Methode_ > insertInlinePictureFromBase64 (base64EncodedImage: String, InsertLocation: InsertLocation)|Fügt an der angegebenen Position im Inhaltssteuerelement ein Inlinebild ein. Der insertLocation-Wert kann "Replace", "Start" oder "End" lauten.|1.2|
|[inlinePicture](/javascript/api/word/word.inlinepicture)|_Beziehung_ > paragraph|Ruft den übergeordneten Absatz ab, der das Inlinebild enthält. Schreibgeschützt.|1.2|
|[inlinePicture](/javascript/api/word/word.inlinepicture)|_Methode_ > delete()|Löscht das Inlinebild aus dem Dokument.|1.2|
|[inlinePicture](/javascript/api/word/word.inlinepicture)|_Methode_ > InsertBreak (BreakType: BreakType InsertLocation: InsertLocation)|Fügt an der angegebenen Position im Hauptdokument einen Umbruch ein. Der insertLocation-Wert kann "Before" oder "After" lauten.|1.2|
|[inlinePicture](/javascript/api/word/word.inlinepicture)|_Methode_ > insertFileFromBase64 (base64File: String, InsertLocation: InsertLocation)|Fügt an der angegebenen Position ein Dokument ein. Der insertLocation-Wert kann "Before" oder "After" lauten.|1.2|
|[inlinePicture](/javascript/api/word/word.inlinepicture)|_Methode_ > InsertHtml (html: String, InsertLocation: InsertLocation)|Fügt an der angegebenen Position HTML-Code ein. Der insertLocation-Wert kann "Before" oder "After" lauten.|1.2|
|[inlinePicture](/javascript/api/word/word.inlinepicture)|_Methode_ > insertInlinePictureFromBase64 (base64EncodedImage: String, InsertLocation: InsertLocation|Fügt an der angegebenen Position ein Inlinebild ein. Der insertLocation-Wert kann "Replace", "Before" oder "After" lauten.|1.2|
|[inlinePicture](/javascript/api/word/word.inlinepicture)|_Methode_ > InsertOoxml (Ooxml: String, InsertLocation: InsertLocation)|Fügt an der angegebenen Position OOXML-Code ein.  Der insertLocation-Wert kann "Before" oder "After" lauten.|1.2|
|[inlinePicture](/javascript/api/word/word.inlinepicture)|_Methode_ > InsertParagraph (ParagraphText: String, InsertLocation: InsertLocation)|Fügt an der angegebenen Position einen Absatz ein. Der insertLocation-Wert kann 'Before' oder 'After' lauten.|1.2|
|[inlinePicture](/javascript/api/word/word.inlinepicture)|_Methode_ > InsertText (Text: String, InsertLocation: InsertLocation)|Fügt an der angegebenen Position Text ein. Der insertLocation-Wert kann "Before" oder "After" lauten.|1.2|
|[inlinePicture](/javascript/api/word/word.inlinepicture)|_Methode_ > Select(selectionMode: SelectionMode)|Wählt das Inlinebild aus. Dies bewirkt, dass Word einen Bildlauf zur Auswahl durchführt.|1.2|
|[range](/javascript/api/word/word.range)|_Beziehung_ > inlinePictures|Ruft die Sammlung von Inlinebildobjekten im Bereich ab. Schreibgeschützt.|1.2|
|[range](/javascript/api/word/word.range)|_Methode_ > insertInlinePictureFromBase64 (base64EncodedImage: String, InsertLocation: InsertLocation)|Fügt an der angegebenen Position ein Bild ein. Der insertLocation-Wert kann "Replace", "Start", "End", "Before" oder "After" lauten.|1.2|

## <a name="word-javascript-api-11"></a>Word-JavaScript-API 1.1

Die Word-JavaScript-API 1.1 ist die erste Version der API. Details zu der API finden Sie in den Referenzthemen zur [JavaScript-API für Word](/javascript/api/word). 

## <a name="see-also"></a>Siehe auch

- [Office-Versionen und Anforderungssätze](https://docs.microsoft.com/office/dev/add-ins/develop/office-versions-and-requirement-sets)
- 
  [Angeben von Office-Hosts und API-Anforderungen](https://docs.microsoft.com/office/dev/add-ins/develop/specify-office-hosts-and-api-requirements)
- [XML-Manifest für Office-Add-Ins](https://docs.microsoft.com/office/dev/add-ins/develop/add-in-manifests)
