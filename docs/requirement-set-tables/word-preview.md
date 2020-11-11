| Klasse | Felder | Beschreibung |
|:---|:---|:---|
|[ContentControl](/javascript/api/word/word.contentcontrol)|[onDataChanged](/javascript/api/word/word.contentcontrol#ondatachanged)|Tritt auf, wenn Daten innerhalb des Inhaltssteuerelements geändert werden.|
||[onDeleted](/javascript/api/word/word.contentcontrol#ondeleted)|Tritt auf, wenn das Inhaltssteuerelement gelöscht wird.|
||[onSelectionChanged](/javascript/api/word/word.contentcontrol#onselectionchanged)|Tritt auf, wenn die Auswahl innerhalb des Inhaltssteuerelements geändert wird.|
|[ContentControlEventArgs](/javascript/api/word/word.contentcontroleventargs)|[contentControl](/javascript/api/word/word.contentcontroleventargs#contentcontrol)|Das Objekt, das das Ereignis ausgelöst hat.|
||[eventType](/javascript/api/word/word.contentcontroleventargs#eventtype)|Der Ereignistyp.|
|[CustomXmlPart](/javascript/api/word/word.customxmlpart)|[delete()](/javascript/api/word/word.customxmlpart#delete--)|Löscht die benutzerdefinierte XML-Komponente.|
||[DeleteAttribute (XPath: String, namespaceMappings: any, Name: String)](/javascript/api/word/word.customxmlpart#deleteattribute-xpath--namespacemappings--name-)|Löscht ein Attribut mit dem angegebenen Namen aus dem Element, das durch XPath identifiziert wird.|
||[DeleteElement (XPath: String, namespaceMappings: any)](/javascript/api/word/word.customxmlpart#deleteelement-xpath--namespacemappings-)|Löscht das durch XPath identifizierte Element.|
||[getXml ()](/javascript/api/word/word.customxmlpart#getxml--)|Ruft den vollständigen XML-Inhalt der benutzerdefinierten XML-Komponente ab.|
||[InsertAttribute (XPath: String, namespaceMappings: any, Name: String, Value: String)](/javascript/api/word/word.customxmlpart#insertattribute-xpath--namespacemappings--name--value-)|Fügt ein Attribut mit dem angegebenen Namen und Wert in das durch XPath identifizierte Element ein.|
||[InsertElement (XPath: String, XML: String, namespaceMappings: any, Index?: Number)](/javascript/api/word/word.customxmlpart#insertelement-xpath--xml--namespacemappings--index-)|Fügt die angegebene XML-Code unter dem übergeordneten Element durch XPath an untergeordneten Position Index identifiziert.|
||[Query (XPath: String, namespaceMappings: any)](/javascript/api/word/word.customxmlpart#query-xpath--namespacemappings-)|Fragt den XML-Inhalt der benutzerdefinierten XML-Komponente ab.|
||[id](/javascript/api/word/word.customxmlpart#id)|Ruft die ID der benutzerdefinierten XML-Komponente ab.|
||[NamespaceURI](/javascript/api/word/word.customxmlpart#namespaceuri)|Ruft den Namespace-URI der benutzerdefinierten XML-Komponente ab.|
||[setXml (XML: String)](/javascript/api/word/word.customxmlpart#setxml-xml-)|Legt den vollständigen XML-Inhalt der benutzerdefinierten XML-Komponente fest.|
||[UpdateAttribute (XPath: String, namespaceMappings: any, Name: String, Value: String)](/javascript/api/word/word.customxmlpart#updateattribute-xpath--namespacemappings--name--value-)|Aktualisiert den Wert eines Attributs mit dem angegebenen Namen des Elements, das durch XPath identifiziert wird.|
||[updateElement (XPath: String, XML: String, namespaceMappings: any)](/javascript/api/word/word.customxmlpart#updateelement-xpath--xml--namespacemappings-)|Aktualisiert die XML-Code des Elements, das durch XPath identifiziert wird.|
|[CustomXmlPartCollection](/javascript/api/word/word.customxmlpartcollection)|[Add (XML: Zeichenfolge)](/javascript/api/word/word.customxmlpartcollection#add-xml-)|Fügt dem Dokument eine neue benutzerdefinierte XML-Komponente hinzu.|
||[getByNamespace (NamespaceURI: String)](/javascript/api/word/word.customxmlpartcollection#getbynamespace-namespaceuri-)|Ruft eine neue bereichsbezogene Sammlung von benutzerdefinierten XML-Komponenten ab, deren Namespaces dem angegebenen Namespace entsprechen.|
||[getCount()](/javascript/api/word/word.customxmlpartcollection#getcount--)|Ruft die Anzahl der Elemente in der Auflistung ab.|
||[getItem(id: string)](/javascript/api/word/word.customxmlpartcollection#getitem-id-)|Ruft eine benutzerdefinierte XML-Komponente basierend auf ihrer ID ab.|
||[getItemOrNullObject(id: string)](/javascript/api/word/word.customxmlpartcollection#getitemornullobject-id-)|Ruft eine benutzerdefinierte XML-Komponente basierend auf ihrer ID ab.|
||[items](/javascript/api/word/word.customxmlpartcollection#items)|Ruft die geladenen untergeordneten Elemente in dieser Sammlung ab.|
|[CustomXmlPartScopedCollection](/javascript/api/word/word.customxmlpartscopedcollection)|[getCount()](/javascript/api/word/word.customxmlpartscopedcollection#getcount--)|Ruft die Anzahl der Elemente in der Auflistung ab.|
||[getItem(id: string)](/javascript/api/word/word.customxmlpartscopedcollection#getitem-id-)|Ruft eine benutzerdefinierte XML-Komponente basierend auf ihrer ID ab.|
||[getItemOrNullObject(id: string)](/javascript/api/word/word.customxmlpartscopedcollection#getitemornullobject-id-)|Ruft eine benutzerdefinierte XML-Komponente basierend auf ihrer ID ab.|
||[getOnlyItem()](/javascript/api/word/word.customxmlpartscopedcollection#getonlyitem--)|Wenn die Sammlung genau ein Element enthält, gibt diese Methode es zurück.|
||[getOnlyItemOrNullObject()](/javascript/api/word/word.customxmlpartscopedcollection#getonlyitemornullobject--)|Wenn die Sammlung genau ein Element enthält, gibt diese Methode es zurück.|
||[items](/javascript/api/word/word.customxmlpartscopedcollection#items)|Ruft die geladenen untergeordneten Elemente in dieser Sammlung ab.|
|[Document](/javascript/api/word/word.document)|[deleteBookmark (Name: String)](/javascript/api/word/word.document#deletebookmark-name-)|Löscht eine Textmarke (sofern vorhanden) aus dem Dokument.|
||[getBookmarkRange (Name: String)](/javascript/api/word/word.document#getbookmarkrange-name-)|Ruft den Bereich eines Lesezeichens ab.|
||[getBookmarkRangeOrNullObject (Name: String)](/javascript/api/word/word.document#getbookmarkrangeornullobject-name-)|Ruft den Bereich eines Lesezeichens ab.|
||[customXmlParts](/javascript/api/word/word.document#customxmlparts)|Ruft die benutzerdefinierten XML-Komponenten im Dokument ab.|
||[onContentControlAdded](/javascript/api/word/word.document#oncontentcontroladded)|Tritt ein, wenn ein Inhaltssteuerelement hinzugefügt wird.|
||[Einstellungen](/javascript/api/word/word.document#settings)|Ruft die Einstellungen des Add-Ins im Dokument ab.|
|[DocumentCreated](/javascript/api/word/word.documentcreated)|[deleteBookmark (Name: String)](/javascript/api/word/word.documentcreated#deletebookmark-name-)|Löscht eine Textmarke (sofern vorhanden) aus dem Dokument.|
||[getBookmarkRange (Name: String)](/javascript/api/word/word.documentcreated#getbookmarkrange-name-)|Ruft den Bereich eines Lesezeichens ab.|
||[getBookmarkRangeOrNullObject (Name: String)](/javascript/api/word/word.documentcreated#getbookmarkrangeornullobject-name-)|Ruft den Bereich eines Lesezeichens ab.|
||[customXmlParts](/javascript/api/word/word.documentcreated#customxmlparts)|Ruft die benutzerdefinierten XML-Komponenten im Dokument ab.|
||[Einstellungen](/javascript/api/word/word.documentcreated#settings)|Ruft die Einstellungen des Add-Ins im Dokument ab.|
|[InlinePicture](/javascript/api/word/word.inlinepicture)|[Image Format](/javascript/api/word/word.inlinepicture#imageformat)|Ruft das Format des Inline Bilds ab.|
|[List](/javascript/api/word/word.list)|[getLevelFont (Level: Number)](/javascript/api/word/word.list#getlevelfont-level-)|Ruft die Schriftart des Aufzählungszeichens, der Zahl oder der Grafik auf der angegebenen Ebene in der Liste ab.|
||[getLevelPicture (Level: Number)](/javascript/api/word/word.list#getlevelpicture-level-)|Ruft die Base64-codierte Zeichenfolgendarstellung des Bilds auf der angegebenen Ebene in der Liste ab.|
||[resetLevelFont (Level: Number, resetFontName?: Boolean)](/javascript/api/word/word.list#resetlevelfont-level--resetfontname-)|Setzt die Schriftart des Aufzählungszeichens, der Zahl oder der Grafik auf der angegebenen Ebene in der Liste zurück.|
||[setLevelPicture (Level: Number, base64EncodedImage?: String)](/javascript/api/word/word.list#setlevelpicture-level--base64encodedimage-)|Legt das Bild auf der angegebenen Ebene in der Liste fest.|
|[Range](/javascript/api/word/word.range)|[GetBookmarks (includeHidden?: Boolean, includeAdjacent?: Boolean)](/javascript/api/word/word.range#getbookmarks-includehidden--includeadjacent-)|Ruft die Namen aller Lesezeichen in oder überlappenden des Bereichs ab.|
||[insertBookmark (Name: String)](/javascript/api/word/word.range#insertbookmark-name-)|Fügt eine Textmarke in den Bereich ein.|
|[Einstellung](/javascript/api/word/word.setting)|[delete()](/javascript/api/word/word.setting#delete--)|Löscht die Einstellung.|
||[key](/javascript/api/word/word.setting#key)|Ruft den Schlüssel der Einstellung ab.|
||[value](/javascript/api/word/word.setting#value)|Dient zum Abrufen oder Festlegen des Werts der Einstellung.|
|[SettingCollection](/javascript/api/word/word.settingcollection)|[Add (Key: String, Value: any)](/javascript/api/word/word.settingcollection#add-key--value-)|Erstellt eine neue Einstellung oder legt eine vorhandene Einstellung fest.|
||[DeleteAll ()](/javascript/api/word/word.settingcollection#deleteall--)|Löscht alle Einstellungen in diesem Add-in.|
||[getCount()](/javascript/api/word/word.settingcollection#getcount--)|Ruft die Anzahl der Einstellungen ab.|
||[getItem(key: string)](/javascript/api/word/word.settingcollection#getitem-key-)|Ruft ein Einstellungsobjekt anhand des Schlüssels ab, bei dem die Groß-/Kleinschreibung beachtet wird.|
||[getItemOrNullObject(key: string)](/javascript/api/word/word.settingcollection#getitemornullobject-key-)|Ruft ein Einstellungsobjekt anhand des Schlüssels ab, bei dem die Groß-/Kleinschreibung beachtet wird.|
||[items](/javascript/api/word/word.settingcollection#items)|Ruft die geladenen untergeordneten Elemente in dieser Sammlung ab.|
|[Table](/javascript/api/word/word.table)|[mergeCells (topRow: Number, firstCell: Number, bottomRow: Number, lastCell: Number)](/javascript/api/word/word.table#mergecells-toprow--firstcell--bottomrow--lastcell-)|Verbindet die eingebundenen Zellen mit einer ersten und letzten Zelle.|
|[TableCell](/javascript/api/word/word.tablecell)|[Split (ROWCOUNT: Number, ColumnCount: Number)](/javascript/api/word/word.tablecell#split-rowcount--columncount-)|Teilt die Zelle in die angegebene Anzahl von Zeilen und Spalten auf.|
|[TableRow](/javascript/api/word/word.tablerow)|[insertContentControl()](/javascript/api/word/word.tablerow#insertcontentcontrol--)|Fügt ein Inhaltssteuerelement in die Zeile ein.|
||[Merge ()](/javascript/api/word/word.tablerow#merge--)|Führt die Zeile in einer Zelle zusammen.|
