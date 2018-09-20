# <a name="onenote-javascript-api-overview"></a>OneNote-JavaScript-API (Übersicht)

Gilt für: OneNote Online

Die folgenden Links zeigen die in der API verfügbaren OneNote-Objekte auf hoher Ebene. Jeder Seite-Verknüpfung enthält eine Beschreibung der Eigenschaften, Ereignisse und Methoden für das Objekt verfügbar. Durchsuchen Sie diese Links, um mehr zu erfahren. 
    
- [Application](/javascript/api/onenote/onenote.application): Das Objekt der obersten Ebene, das für den Zugriff auf alle global adressierbaren OneNote-Objekte verwendet wird, z. B. das aktive Notizbuch und den aktiven Abschnitt.

- [Notebook](/javascript/api/onenote/onenote.notebook): Ein Notizbuch. Notizbücher können Abschnittsgruppen und Abschnitte enthalten.
    - [NotebookCollection](/javascript/api/onenote/onenote.notebookcollection): Eine Auflistung von Notizbüchern.

- [SectionGroup](/javascript/api/onenote/onenote.sectiongroup): Eine Abschnittsgruppe. Abschnittsgruppen enthalten Abschnittsgruppen und Abschnitte.
    - [SectionGroupCollection](/javascript/api/onenote/onenote.sectiongroupcollection): Eine Auflistung von Abschnittsgruppen.

- [Section](/javascript/api/onenote/onenote.section): Ein Abschnitt. Abschnitte enthalten Seiten.
    - [SectionCollection](/javascript/api/onenote/onenote.sectioncollection): Eine Auflistung von Abschnitten.

- [Page](/javascript/api/onenote/onenote.page): Eine Seite. Seiten enthalten PageContent-Objekte.
    - [PageCollection](/javascript/api/onenote/onenote.pagecollection): Eine Auflistung von Seiten.

- [PageContent](/javascript/api/onenote/onenote.pagecontent): Ein Bereich auf oberster Ebene auf einer Seite, der Inhaltstypen enthält, z. B. Outline oder Image. Ein PageContent-Objekt kann einer Position auf der Seite zugewiesen werden.
    - [PageContentCollection](/javascript/api/onenote/onenote.pagecontentcollection): Eine Auflistung von PageContent-Objekten, die die Inhalte einer Seite darstellt.

- [Outline](/javascript/api/onenote/onenote.outline): Ein Container für Paragraph-Objekte. Eine Gliederung ist ein direkt untergeordnetes Element eines PageContent-Objekts.

- [Image](/javascript/api/onenote/onenote.image): Ein Bildobjekt. Ein Bild kann ein direkt untergeordnetes Element eines PageContent-Objekts oder eines Paragraph-Objekts sein.

- [Paragraph](/javascript/api/onenote/onenote.paragraph): Ein Container für den sichtbaren Inhalt auf einer Seite. Ein Absatz ist ein direkt untergeordnetes Element einer Gliederung.
    - [ParagraphCollection](/javascript/api/onenote/onenote.paragraphcollection): Eine Auflistung von Paragraph-Objekten in einer Gliederung.

- [RichText](/javascript/api/onenote/onenote.richtext): Ein RichText-Objekt.

- [Table](/javascript/api/onenote/onenote.table): Ein Container für TableRow-Objekte.

- [TableRow](/javascript/api/onenote/onenote.tablerow): Ein Container für TableCell-Objekte.
    - [TableRowCollection](/javascript/api/onenote/onenote.tablerowcollection): Eine Auflistung von TableRow-Objekten in einer Tabelle.
 
- [TableCell](/javascript/api/onenote/onenote.tablecell): Ein Container für Paragraph-Objekte.
    - [TableCellCollection](/javascript/api/onenote/onenote.tablecellcollection): Eine Auflistung von TableCell-Objekten in einer TableRow.

## <a name="onenote-javascript-api-reference"></a>JavaScript-API-Referenz für OneNote

Ausführliche Informationen zu OneNote JavaScript-API finden Sie in der [Referenzdokumentation für die OneNote-JavaScript-API](/javascript/api/onenote).

## <a name="see-also"></a>Siehe auch

- [Übersicht über die JavaScript-API-Programmierung für OneNote](https://docs.microsoft.com/office/dev/add-ins/onenote/onenote-add-ins-programming-overview)
- [Erstellen Ihres ersten OneNote-Add-Ins](https://docs.microsoft.com/office/dev/add-ins/onenote/onenote-add-ins-getting-started)
- [Rubric Grader-Beispiel](https://github.com/OfficeDev/OneNote-Add-in-Rubric-Grader)
- [Office-Add-Ins-Plattformübersicht](https://docs.microsoft.com/office/dev/add-ins/overview/office-add-ins)
