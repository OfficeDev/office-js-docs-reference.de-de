# <a name="how-the-office-javascript-api-documentation-is-generated"></a>So wird die Office-JavaScript-API-Dokumentation generiert

Die Office JavaScript-Referenz Dokumentationsseiten werden aus typdefinitions Dateien und Beispielcode Ausschnitten generiert. Dieser Prozess verwendet eine Mischung aus Open Source-Tools und Repository-spezifischen Skripts. Dieses Dokument zielt darauf ab, die Prozesse in diesem Repository transparent zu machen, damit die Community besser von diesen Inhalten profitieren und dazu beitragen kann.

## <a name="content-sources"></a>Inhaltsquellen

Zwei Arten von Inhalten werden zusammengefasst, um die Office-js-Referenzdokumentation zu erstellen: Typdefinitionen und Codeausschnitte. Dadurch wird eine vollständige API-Abdeckung gewährleistet und kleine Inlinecode Beispiele gegeben.

### <a name="type-definition-files"></a>Typen Definitionsdateien

Die typdefinitions Dateien auf [eindeutig typisiert](https://github.com/DefinitelyTyped/DefinitelyTyped) sind die einzige Quelle der Wahrheit für die Dokumentation. Alle Office-Add-in, bei denen die Ausführung von schreiben mithilfe dieser typdefinitions Dateien kompiliert wird. Diese geben auch JavaScript-und Scripting-Entwickler-IntelliSense-Funktionen. Durch die Erstellung der Referenzdokumentation aus diesen Definitionen werden genauere Informationen bereitgestellt.

Es gibt vier relevante d. TS-Dateien, die Quellinhalte für verschiedene Unterabschnitte der Dokumente bereitstellen.

- [Office-js/Index. d. TS](https://raw.githubusercontent.com/DefinitelyTyped/DefinitelyTyped/master/types/office-js/index.d.ts) (die Veröffentlichungs Definitionen)
  - [Excel (Version)](https://docs.microsoft.com/javascript/api/excel_release)
  - [OneNote](https://docs.microsoft.com/javascript/api/onenote)
  - [PowerPoint](https://docs.microsoft.com/javascript/api/powerpoint)
  - [Visio](https://docs.microsoft.com/javascript/api/visio)
  - [Word (Version)](https://docs.microsoft.com/javascript/api/word_release)
  - [OfficeExtensions-unter Abschnitt der allgemeinen API](https://docs.microsoft.com/javascript/api/office)
- [Office-js-Preview/Index. d. TS](https://raw.githubusercontent.com/DefinitelyTyped/DefinitelyTyped/master/types/office-js-preview/index.d.ts) (die Vorschau Definitionen.)
  - [Excel (Vorschau)](https://docs.microsoft.com/javascript/api/excel)
  - [Outlook (Vorschau)](https://docs.microsoft.com/javascript/api/outlook)
  - [Word (Vorschau)](https://docs.microsoft.com/javascript/api/word)
  - [Allgemeine API](https://docs.microsoft.com/javascript/api/office)
- [Custom-Functions-Runtime/Index. d. TS](https://github.com/DefinitelyTyped/DefinitelyTyped/blob/master/types/custom-functions-runtime/index.d.ts) (die CLR-Definitionen für benutzerdefinierte Excel-Funktionen)
  - [Benutzerdefinierte Funktionen](https://docs.microsoft.com/javascript/api/custom-functions-runtime)
- [Office-Runtime/Index. d. TS](https://github.com/DefinitelyTyped/DefinitelyTyped/blob/master/types/office-runtime/index.d.ts) (die Office-Lauf Zeitdefinitionen für die Plattform für benutzerdefinierte Funktionen)
  - [Office-Laufzeit](https://docs.microsoft.com/javascript/api/office-runtime)

Ältere Versionen der APIs haben eigene d. TS-Dateien. Diese werden beibehalten, wenn ein neuer API-Anforderungs Satzes freigegeben wird. Sie können auch mithilfe des [Version Remover-Tools](https://github.com/OfficeDev/office-js-docs-reference/blob/master/generate-docs/tools/VersionRemover.ts)generiert werden. Diese alten d. TS-Dateien werden beibehalten, sodass im Ereignis-APIs gepatcht oder geändert werden, das ursprüngliche Verhalten ist weiterhin dokumentiert. Dies ist hilfreich, wenn Sie eine ältere Version der API als Ziel haben.

#### <a name="testing-type-definition-file-changes"></a>Testen von Änderungen an der Typen Definitionsdatei

Alle Änderungen an der Dokumentation für die Office-JavaScript-API werden durch Bearbeiten der oben erwähnten vier d. TS-Dateien vorgenommen. Sie können jedoch eine Änderung testen, bevor Sie eine PR an DefinitelyTyped übermitteln (wenn Sie beispielsweise testen möchten, wie ihre Formatierung in ein Abschlag übersetzt wird), indem Sie die entsprechende Datei in [generieren-docs/Script-Inputs](https://github.com/OfficeDev/office-js-docs-reference/tree/master/generate-docs/script-inputs) bearbeiten und [GenerateDocs. cmd](https://github.com/OfficeDev/office-js-docs-reference/blob/master/generate-docs/GenerateDocs.cmd)ausführen. Wenn Sie dazu aufgefordert werden, wählen Sie die Option "lokale Dateien" aus.

Durch das Pushen von Änderungen an einem Remote Zweig dieses Repo wird die docs.Microsoft.com-Plattform zum Erstellen einer Test Verzweigung veranlasst. Diese Verzweigung wird auf review.docs.Microsoft.com gerendert, auf die nur interne Microsoft-Mitarbeiter zugreifen können. Jeder, der Ihre PR überprüft, überprüft die Website überprüfen auf Richtigkeit.

### <a name="code-snippets"></a>Codeausschnitte

Code Beispiel Ausschnitte werden den Verweisseiten aus zwei Quellen hinzugefügt:

- [Skript laborbeispiele](https://github.com/OfficeDev/office-js-snippets)
- [Lokale Code Ausschnitte](https://github.com/OfficeDev/office-js-docs-reference/tree/master/docs/code-snippets)

Die lokalen Codeausschnitte befinden sich in hostspezifischen YAML-Dateien. Ihr Inhalt ist nach Klasse und Feld organisiert, sodass er der entsprechenden Stelle in einer Referenzseite zugeordnet werden kann. Die Sprache des Codeausschnitts (JavaScript oder Manuskript) wird durch die Verwendung von await-Anweisungen hergeleitet.

Die Skript Labor Codeausschnitte werden aus Arbeitsbeispielen abgerufen. Derzeit werden Excel-, Outlook-und Word-Beispiele den Referenzdokumenten Abschnitten durch [Zuordnungsdateien](https://github.com/OfficeDev/office-js-snippets/tree/master/snippet-extractor-metadata)zugeordnet. Diese stimmen mit einzelnen Beispiel Methoden mit Eigenschaften oder Methoden in der API überein. Beim Ausführen des Office-js-Snippets-Repositorys `yarn start` wird [eine YAML-Datei](https://github.com/OfficeDev/office-js-snippets/blob/master/snippet-extractor-output/snippets.yaml) erstellt, die alle zugeordneten Codeausschnitte enthält. Diese YAML-Datei ist die Eingabe in das Referenz Dokumentationstool.

## <a name="tooling-pipeline"></a>Tooling-Pipeline

![Ein Bild, das die Ablaufsteuerung von eindeutig typisiert, an den Präprozessor, den API-Extraktionsprozess, den Präprozessor, den API-Dokumenter und bis zum Postprocessor zeigt.](ToolingPipeline.png)

Zwischen den Inhaltsquellen und den letzten Seiten werden im Dokumentations Inhalt fünf Tools beschrieben:

1. [Präprozessor-Skript](https://github.com/OfficeDev/office-js-docs-reference/blob/master/generate-docs/scripts/preprocessor.ts)
1. [API-Extraktion](https://api-extractor.com/)
1. [Präprozessor-Skript](https://github.com/OfficeDev/office-js-docs-reference/blob/master/generate-docs/scripts/midprocessor.ts)
1. [API-Dokumentation](https://github.com/microsoft/rushstack/blob/master/apps/api-documenter/README.md)
1. [Postprocessor-Skript](https://github.com/OfficeDev/office-js-docs-reference/blob/master/generate-docs/scripts/postprocessor.ts)

Der Präprozessor verwendet die d. TS-Dateien und teilt Sie in hostspezifische Abschnitte auf. Sie führt alle erforderlichen Aufräumvorgänge aus, damit die nachfolgenden Tools die Daten ordnungsgemäß verarbeiten können.

API Extractor konvertiert die d. TS-Dateien in JSON-Daten. Dadurch werden alle Typdaten löst, sodass eine einfachere Analyse möglich ist.

Der Mittel Prozessor Ruft die Codeausschnitte ab und verbindet Sie mit den richtigen Hosts und bereinigt die Vernetzung zwischen Outlook und Common API-Objekten.

API-Dokumentierer wandelt die JSON-Daten in yml-Dateien um. Die yml-Dateien werden vom geöffneten Veröffentlichungs System, das unsere Dokumente in docs.Microsoft.com veröffentlicht, in einen Abschlag umgewandelt. API-Dokumentierer enthält auch eine Office-spezifische Erweiterung, die unsere Codeausschnitte einfügt.

Der postprocesser bereinigt das Inhaltsverzeichnis und verschiebt die yml-Dateien in den [Veröffentlichungsordner](https://github.com/OfficeDev/office-js-docs-reference/tree/master/docs/docs-ref-autogen).

Alle fünf Schritte werden ausgeführt, wenn [GenerateDocs. cmd](https://github.com/OfficeDev/office-js-docs-reference/blob/master/generate-docs/GenerateDocs.cmd) ausgeführt wird. Dieses Skript verarbeitet auch die Installation von Knoten Modulen, bereinigt alte Dateisätze und typdefinitions Dateien für Versionstypen für jeden Anforderungssatz.
