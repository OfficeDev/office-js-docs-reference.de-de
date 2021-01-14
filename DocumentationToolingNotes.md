# <a name="how-the-office-javascript-api-documentation-is-generated"></a>Wie die JavaScript-API-Dokumentation für Office generiert wird

Die Office-JavaScript-Referenzdokumentationsseiten werden aus Typdefinitionsdateien und Beispielausschnitten generiert. Dieser Prozess verwendet eine Mischung aus Open Source-Tools und repositoryspezifischen Skripts. Dieses Dokument soll die Prozesse dieses Repositorys transparent machen, damit die Community von diesen Inhalten besser profitieren und zu diesen beitragen kann.

## <a name="content-sources"></a>Inhaltsquellen

Zwei Inhaltstypen werden kombiniert, um die Office-JS-Referenzdokumentation zu erstellen: Typdefinitionen und Codeausschnitte. Diese stellen eine vollständige API-Abdeckung sicher und geben kleine Inlinecodebeispiele.

### <a name="type-definition-files"></a>Typdefinitionsdateien

Die Typdefinitionsdateien auf ["Definitely Typed"](https://github.com/DefinitelyTyped/DefinitelyTyped) sind die einzige Wahrheitsquelle für die Dokumentation. Jedes Office-Add-In, das TypeScript verwendet, kompiliert mithilfe dieser Typdefinitionsdateien. Diese bieten JavaScript- und TypeScript-Entwicklern IntelliSense Funktionen. Durch erstellen der Referenzdokumentation aus diesen Definitionen stellen wir genauere Informationen zur Verfügung.

Es gibt vier relevante d.ts-Dateien, die Quellinhalte für verschiedene Unterabschnitte der Dokumente bereitstellen.

- [office-js/index.d.ts](https://raw.githubusercontent.com/DefinitelyTyped/DefinitelyTyped/master/types/office-js/index.d.ts) (Die Releasedefinitionen.)
  - [Excel (Release)](https://docs.microsoft.com/javascript/api/excel_release)
  - [OneNote](https://docs.microsoft.com/javascript/api/onenote)
  - [PowerPoint](https://docs.microsoft.com/javascript/api/powerpoint)
  - [Visio](https://docs.microsoft.com/javascript/api/visio)
  - [Word (Release)](https://docs.microsoft.com/javascript/api/word_release)
  - [Unterabschnitt "OfficeExtensions" der allgemeinen API](https://docs.microsoft.com/javascript/api/office)
- [office-js-preview/index.d.ts](https://raw.githubusercontent.com/DefinitelyTyped/DefinitelyTyped/master/types/office-js-preview/index.d.ts) (Die Vorschaudefinitionen.)
  - [Excel (Vorschau)](https://docs.microsoft.com/javascript/api/excel)
  - [Outlook (Vorschau)](https://docs.microsoft.com/javascript/api/outlook)
  - [Word (Vorschau)](https://docs.microsoft.com/javascript/api/word)
  - [Allgemeine API](https://docs.microsoft.com/javascript/api/office)
- [custom-functions-runtime/index.d.ts (Die](https://github.com/DefinitelyTyped/DefinitelyTyped/blob/master/types/custom-functions-runtime/index.d.ts) Excel Custom Functions-Laufzeitdefinitionen).)
  - [Benutzerdefinierte Funktionen](https://docs.microsoft.com/javascript/api/custom-functions-runtime)
- [office-runtime/index.d.ts](https://github.com/DefinitelyTyped/DefinitelyTyped/blob/master/types/office-runtime/index.d.ts) (Die Office-Laufzeitdefinitionen für die Plattform für benutzerdefinierte Funktionen.)
  - [Office Runtime](https://docs.microsoft.com/javascript/api/office-runtime)

Ältere Versionen der APIs verfügen über eigene d.ts-Dateien. Diese werden beibehalten, wenn ein neuer API-Anforderungssatz veröffentlicht wird. Sie können auch mit dem Tool ["Versionsentferner" generiert werden.](https://github.com/OfficeDev/office-js-docs-reference/blob/master/generate-docs/tools/VersionRemover.ts) Diese alten d.ts-Dateien werden so beibehalten, dass das ursprüngliche Verhalten weiterhin dokumentiert wird, wenn APIs gepatcht oder geändert werden. Dies ist hilfreich, wenn Sie eine ältere Version der API als Ziel verwenden müssen.

#### <a name="testing-type-definition-file-changes"></a>Testen von Änderungen der Typdefinitionsdatei

Alle Dokumentationsänderungen für die Office-JavaScript-API werden durch Bearbeiten der vier oben genannten d.ts-Dateien vorgenommen. Sie können jedoch eine Änderung testen, bevor Sie eine PR an DefinitelyTyped übermitteln (wenn Sie z. B. testen müssen, wie Ihre Formatierung in Markdown übersetzt wird), indem Sie die entsprechende Datei in [generate-docs/script-inputs](https://github.com/OfficeDev/office-js-docs-reference/tree/master/generate-docs/script-inputs) bearbeiten und [GenerateDocs.cmd](https://github.com/OfficeDev/office-js-docs-reference/blob/master/generate-docs/GenerateDocs.cmd)ausführen. Wenn Sie dazu aufgefordert werden, wählen Sie die Option "Lokale Dateien" aus.

Das Pushen von Änderungen an einen Remotezweig dieses Repositorys bewirkt, dass die docs.microsoft.com A0 einen Testzweig aufbaut. Diese Verzweigung wird auf einem review.docs.microsoft.com gerendert, auf den nur interne Mitarbeiter von Microsoft zugriffen können. Jeder, der Ihre PR überprüft, überprüft die Überprüfungswebsite auf Genauigkeit.

### <a name="code-snippets"></a>Codeausschnitte

Codebeispielausschnitte werden den Referenzseiten aus zwei Quellen hinzugefügt:

- [Script Lab Samples](https://github.com/OfficeDev/office-js-snippets)
- [Lokale Codeausschnitte](https://github.com/OfficeDev/office-js-docs-reference/tree/master/docs/code-snippets)

Die lokalen Codeausschnitte befinden sich in hostspezifischen Yaml-Dateien. Der Inhalt ist nach Klasse und Feld organisiert, sodass er dem entsprechenden Ort auf einer Referenzseite zugeordnet werden kann. Die Sprache des Codeausschnitts (JavaScript oder TypeScript) wird durch die Verwendung von await-Anweisungen abgeleitet.

Die Script Lab-Codeausschnitte werden aus Arbeitsbeispielen entnommen. Derzeit werden Excel-, Outlook-, PowerPoint- und Word-Beispiele so zugeordnet, dass sie über Zuordnungsdateien auf [Dokumentabschnitte verweisen.](https://github.com/OfficeDev/office-js-snippets/tree/prod/snippet-extractor-metadata) Diese entsprechen einzelnen Beispielmethoden Eigenschaften oder Methoden in der API. Wenn das Office-js-snippets-Repository ausgeführt wird, wird eine Yaml-Datei mit allen zugeordneten `yarn start` Codeausschnitten [](https://github.com/OfficeDev/office-js-snippets/blob/prod/snippet-extractor-output/snippets.yaml) erstellt. Diese Yaml-Datei ist die Eingabe in das Referenzdokumentationstool.

## <a name="tooling-pipeline"></a>Toolpipeline

![Abbildung des Steuerungsflusses von "Definitely Typed" zum Präprozessor, API Extractor, Midprocessor, API Documenter und zum Postprocessor.](ToolingPipeline.png)

Zwischen den Inhaltsquellen und den endgültigen Seiten durchgeht der Dokumentationsinhalt fünf Toolschritte:

1. [Präprozessorskript](https://github.com/OfficeDev/office-js-docs-reference/blob/master/generate-docs/scripts/preprocessor.ts)
1. [API Extractor](https://api-extractor.com/)
1. [Midprocessorskript](https://github.com/OfficeDev/office-js-docs-reference/blob/master/generate-docs/scripts/midprocessor.ts)
1. [API Documenter](https://github.com/microsoft/rushstack/blob/master/apps/api-documenter/README.md)
1. [Postprocessorskript](https://github.com/OfficeDev/office-js-docs-reference/blob/master/generate-docs/scripts/postprocessor.ts)

Der Präprozessor übernimmt die d.ts-Dateien und teilt sie in hostspezifische Abschnitte auf. Sie führt alle Bereinigungen durch, die für die nachfolgenden Tools erforderlich sind, um die Daten ordnungsgemäß zu verarbeiten.

Der API Extractor konvertiert die d.ts-Dateien in JSON-Daten. Dadurch werden alle Typdaten tokenisiert, was eine einfachere Analyse ermöglicht.

Der Midprocessor ruft die Codeausschnitte ab und paart sie mit den richtigen Hosts und bereinigt die Verknüpfung zwischen Outlook- und allgemeinen API-Objekten.

API Documenter konvertiert die JSON-Daten in YML-Dateien. Die .yml-Dateien werden vom Open Publishing System, das unsere Dokumente in der Datei veröffentlicht, in Markdown docs.microsoft.com. DER API Documenter enthält auch eine Office-spezifische Erweiterung, die unsere Codeausschnitte einf?en.

Der Postprozessor bereinigt das Inhaltsverzeichnis und verschiebt die YML-Dateien in den [Veröffentlichungsordner.](https://github.com/OfficeDev/office-js-docs-reference/tree/master/docs/docs-ref-autogen)

Alle fünf dieser Schritte werden ausgeführt, wenn [GenerateDocs.cmd](https://github.com/OfficeDev/office-js-docs-reference/blob/master/generate-docs/GenerateDocs.cmd) ausgeführt wird. Dieses Skript verarbeitet auch die Knotenmodulinstallation, bereinigt alte Dateisätze und Versionstypdefinitionsdateien für jeden Anforderungssatz.
