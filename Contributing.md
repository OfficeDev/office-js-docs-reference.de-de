# <a name="contribute-to-this-documentation"></a>An dieser Dokumentation mitwirken

Vielen Dank für Ihr Interesse an unserer Dokumentation!

* [Möglichkeiten zur Mitwirkung](#ways-to-contribute)
* [Mit GitHub mitwirken](#contribute-using-github)
* [Mit Git mitwirken](#contribute-using-git)
* [Verwenden von Markdown zum Formatieren Ihres Themas](#how-to-use-markdown-to-format-your-topic)
* [Häufig gestellte Fragen](#faq)
* [Weitere Ressourcen](#more-resources)

## <a name="ways-to-contribute"></a>Möglichkeiten zur Mitwirkung

Nachfolgend finden Sie einige Möglichkeiten, wie Sie zu dieser Dokumentation beitragen können:

* Lesen Sie die Informationen unter [Mit GitHub mitwirken](#contribute-using-github), um kleine Änderungen an einem Artikel vorzunehmen.
* Lesen Sie die Informationen unter [Mit Git mitwirken](#contribute-using-git), um umfangreiche Änderungen vorzunehmen bzw. Änderungen vorzunehmen, die mit Code zu tun haben.
* Melden von Dokumentationsfehlern über GitHub-Probleme.
* Fordern Sie eine neue Dokumentation auf [der Microsoft 365 on Q&A-Website](https://docs.microsoft.com/answers/products/m365) an.

## <a name="contribute-using-github"></a>Mit GitHub mitwirken

> **Wichtig:** Der Referenzinhalt in diesem Repository wird automatisch generiert. Lesen Sie die [Dokumentationstools, bevor](https://github.com/OfficeDev/office-js-docs-reference/blob/master/DocumentationToolingNotes.md) Sie Dateien bearbeiten.

Verwenden Sie GitHub, um an dieser Dokumentation mitzuwirken, ohne dass Sie hierfür das Repo auf Ihrem Desktop klonen. Dies ist die einfachste Möglichkeit zum Erstellen einer Pullanforderung in diesem Repository. Verwenden Sie diese Methode, um eine kleinere Änderung vorzunehmen, mit der keine Codeänderungen einhergehen.

Mit dieser Methode können Sie zu einem Artikel gleichzeitig beitragen.

### <a name="to-contribute-using-github"></a>So können Sie mit GitHub mitwirken

1. Suchen Sie den Artikel, den Sie bei GitHub beisteuern möchten.
2. Wenn Sie sich im Artikel in GitHub befinden, melden Sie sich bei GitHub an (erhalten Sie ein kostenloses Konto [Join GitHub](https://github.com/join)).
3. Wählen Sie das **Stiftsymbol** aus (bearbeiten Sie die Datei in Ihrer Verzweigung des Projekts), und nehmen Sie Ihre Änderungen im Fenster **<>Edit file** vor.
4. Führen Sie einen Bildlauf nach unten durch, und geben Sie eine Beschreibung ein.
5. Wählen Sie **Propose file change**>**Create pull request** aus.

Sie haben nun erfolgreich eine Pullanforderung übermittelt. Pullanforderungen werden in der Regel innerhalb von 10 Werktagen bearbeitet.

## <a name="contribute-using-git"></a>Mit Git mitwirken

Verwenden Sie Git, um umfangreiche Änderungen vorzunehmen, z. B.:

* Beisteuern von Code
* Beisteuern von Änderungen, die Auswirkungen auf die Bedeutung haben
* Beisteuern von umfangreichen Änderungen an Text
* Hinzufügen neuer Themen

### <a name="to-contribute-using-git"></a>So können Sie mit Git mitwirken

1. Wenn Sie nicht über ein GitHub-Konto verfügen, richten Sie eines auf [GitHub](https://github.com/join) ein.
2. Nachdem Sie ein Konto erstellt haben, installieren Sie Git auf Ihrem Computer. Führen Sie die Schritte im [Lernprogramm Einrichten von Git] aus.
3. Um eine Pullanforderung mit Git zu übermitteln, führen Sie die Schritte unter [Verwenden von GitHub, Git und diesem Repository](#use-github-git-and-this-repository) durch.
4. Sie werden aufgefordert, den Lizenzvertrag für Mitwirkende zu unterzeichnen, wenn Sie:

    * Mitglied der Microsoft Open Technologies-Gruppe sind.
    * Ein Mitwirkender, der nicht für Microsoft arbeitet.

Als Mitglied der Community müssen Sie den Lizenzvertrag für Mitwirkende (CLA) unterzeichnen, bevor Sie umfangreiche Beiträge zu einem Projekt leisten können. Sie müssen die Dokumentation nur einmal ausfüllen und übermitteln. Überprüfen Sie das Dokument sorgfältig. Möglicherweise muss Ihr Arbeitgeber das Dokument unterzeichnen.

Durch das Unterzeichnen des Lizenzvertrags für Mitwirkende (CLA) erhalten Sie keine Zugriffsrechte auf das Hauptrepository, aber es bedeutet, dass die Office Developer- und Office Developer Content Publishing-Teams in der Lage sind, Ihre Beiträge zu prüfen und überdenken. Sie werden für Ihre Übermittlungen guthaben.

Pullanforderungen werden in der Regel innerhalb von 10 Werktagen bearbeitet.

## <a name="use-github-git-and-this-repository"></a>Verwenden von GitHub, Git und diesem Repository

**Hinweis:** Die meisten der Informationen in diesem Abschnitt finden Sie in [GitHub-Hilfe]-Artikeln.  Wenn Sie mit Git und GitHub vertraut sind, fahren Sie mit dem Abschnitt zum **Einsenden und Bearbeiten von Inhalten** mit speziellen Informationen zum Code-/Inhaltsfluss dieses Repositorys fort.

### <a name="to-set-up-your-fork-of-the-repository"></a>So richten Sie Ihre Verzweigung des Repositorys ein

1. Richten Sie ein GitHub-Konto ein, damit Sie an diesem Projekt mitwirken können. Falls noch nicht geschehen, wechseln Sie zu [GitHub](https://github.com/join), und richten Sie jetzt ein Konto ein.
2. Installieren Sie Git auf Ihrem Computer. Führen Sie die Schritte im [Lernprogramm Einrichten von Git] aus.
3. Erstellen Sie Ihre eigene Verzweigung dieses Repositorys. Klicken Sie dazu oben auf der Seite auf die Schaltfläche **Verzweigung**.
4. Kopieren Sie die Verzweigung auf Ihren Computer. Öffnen Sie hierzu Git Bash. Geben Sie an der Eingabeaufforderung Folgendes ein:

        git clone https://github.com/<your user name>/<repo name>.git

    Erstellen Sie danach einen Verweis auf das Stammrepository durch Eingabe der folgenden Befehle:

        cd <repo name>
        git remote add upstream https://github.com/OfficeDev/<repo name>.git
        git fetch upstream

Ausgezeichnet. Sie haben nun Ihr Repository eingerichtet. Sie müssen diese Schritte nicht erneut ausführen.

### <a name="contribute-and-edit-content"></a>Einsenden und Bearbeiten von Inhalten

Um die Mitwirkung für Sie so einfach wie möglich zu machen, führen Sie die folgenden Schritte aus.

#### <a name="to-contribute-and-edit-content"></a>So senden Sie Inhalte ein und bearbeiten diese

1. Erstellen Sie eine neue Verzweigung.
2. Fügen Sie neue Inhalte hinzu, oder bearbeiten Sie vorhandene Inhalte.
3. Senden Sie eine Pull-Anforderung an das Hauptrepository.
4. Löschen Sie die Verzweigung.

**Wichtig** Begrenzen Sie jede Verzweigung auf ein einzelnes Konzept bzw. einen einzelnen Artikel, um den Workflow zu optimieren und das Risiko von Zusammenführungskonflikten zu verringern. Zu geeigneten Inhalten für eine neue Verzweigung gehören:

* Ein neuer Artikel.
* Bearbeitung von Rechtschreibung und Grammatik.
* Vornehmen einer einzelnen Formatierungsänderung für eine große Gruppe von Artikeln (z. B. Übernehmen einer neuen Copyright-Fußzeile).

#### <a name="to-create-a-new-branch"></a>So erstellen Sie eine neue Verzweigung

1. Öffnen Sie Git Bash.
2. Geben Sie an der Git Bash-Eingabeaufforderung `git pull upstream master:<new branch name>` ein. Dadurch wird eine neue Verzweigung lokal erstellt, die aus der neuesten OfficeDev-Hauptverzweigung kopiert wird.
3. Geben Sie an der Git Bash-Eingabeaufforderung `git push origin <new branch name>` ein. Dadurch wird GitHub auf die neue Verzweigung hingewiesen. Die neue Verzweigung sollte nun in Ihrer Verzweigung des Repository auf GitHub angezeigt werden.
4. Geben Sie in der Git Bash-Befehlszeile `git checkout <new branch name>` ein, um zu Ihrer neuen Verzweigung umzuschalten.

#### <a name="add-new-content-or-edit-existing-content"></a>Hinzufügen neuer Inhalte oder Bearbeiten vorhandener Inhalte

Sie navigieren mithilfe des Datei-Explorers zu dem Repository auf dem Computer. Die Repositorydateien befinden sich in `C:\Users\<yourusername>\<repo name>`.

Zum Bearbeiten von Dateien öffnen Sie sie in einem Editor Ihrer Wahl und ändern sie. Um eine neue Datei zu erstellen, verwenden Sie den Editor Ihrer Wahl, und speichern Sie die neue Datei im entsprechenden Speicherort in Ihrer lokalen Kopie des Repositorys. Speichern Sie während der Bearbeitung häufig.

Die Dateien in `C:\Users\<yourusername>\<repo name>` sind eine Arbeitskopie der neuen Verzweigung, die Sie in Ihrem lokalen Repository erstellt haben. Änderungen in diesem Ordner wirken sich erst auf das lokale Repository aus, wenn Sie eine Änderung übernehmen. Um eine Änderung am lokalen Repository zu übernehmen, geben Sie die folgenden Befehle in GitBash ein:

    git add .
    git commit -v -a -m "<Describe the changes made in this commit>"

Der `add`-Befehl fügt Ihre Änderungen einem Stagingbereich als Vorbereitung für die Aufnahme im Repository hinzu. Der Punkt nach dem `add`-Befehl gibt an, dass Sie alle Dateien, die Sie hinzugefügt oder geändert haben, übernehmen möchten und dass Unterordner rekursiv überprüft werden sollen. (Wenn Sie nicht alle Änderungen übernehmen möchten, können Sie bestimmte Dateien hinzufügen. Sie können den Vorgang auch rückgängig machen. Falls Sie Hilfe benötigen, geben Sie `git add -help` oder `git status` ein.)

Der Befehl `commit` wendet die bereitgestellten Änderungen auf das Repository an. Der Schalter `-m` bedeutet, dass Sie den Commit-Kommentar in der Befehlszeile angeben. Die Optionen -v und -a können weggelassen werden. Die Option -v ist für die ausführliche Ausgabe des Befehls, und -a entspricht in ihrer Funktion den bereits verwendeten Befehl zum Hinzufügen.

Sie können während der Arbeit mehrfach übernehmen oder warten und erst nach Abschluss der Arbeit übernehmen.

#### <a name="submit-a-pull-request-to-the-main-repository"></a>Einsenden einer Pull-Anforderung an das Hauptrepository

Wenn Sie Ihre Arbeit abgeschlossen haben und bereit sind, sie mit dem zentralen Repository zusammenzuführen, gehen Sie folgendermaßen vor.

#### <a name="to-submit-a-pull-request-to-the-main-repository"></a>So übermitteln Sie eine Pullanforderung an das Hauptrepository

1. Geben Sie an der Git Bash-Eingabeaufforderung `git push origin <new branch name>` ein. Im lokalen Repository bezieht sich `origin` auf Ihr GitHub-Repository, aus dem Sie das lokale Repository geklont haben. Mit diesem Befehl wird der aktuelle Status Ihrer neuen Verzweigung, einschließlich aller im vorherigen Schritt vorgenommenen Commits, in Ihre GitHub-Verzweigung übertragen.
2. Navigieren Sie auf der GitHub-Website in Ihrer Verzweigung zur neuen Verzweigung.
3. Klicken Sie oben auf der Seite auf die Schaltfläche **Pullanforderung**.
4. Stellen Sie sicher, dass die untere Verzweigung `OfficeDev/<repo name>@master` und die obere Verzweigung `<your username>/<repo name>@<branch name>` lautet.
5. Klicken Sie auf die Schaltfläche **Commitbereich aktualisieren**.
6. Geben Sie Ihrer Pullanforderung einen Titel, und beschreiben Sie alle Änderungen, die Sie vornehmen.
7. Senden Sie die Pullanforderung.

Einer der Websiteadministratoren wird nun Ihre Pullanforderung verarbeiten. Ihre Pullanforderung wird auf der OfficeDev/<repo name>-Website unter „Probleme“ angezeigt. Wenn die Pullanforderung angenommen wird, wird das Problem gelöst.

#### <a name="create-a-new-branch-after-merge"></a>Erstellen einer neuen Verzweigung nach dem Zusammenführen

Nachdem eine Verzweigung erfolgreich zusammengeführt wurde (d. h. Ihre Pullanforderung wurde angenommen), arbeiten Sie nicht in der lokalen Verzweigung weiter. Dies kann zu Zusammenführungskonflikten führen, wenn Sie eine weitere Pullanforderung einsenden. Wenn Sie eine weitere Aktualisierung vornehmen möchten, erstellen Sie eine neue lokale Verzweigung aus der erfolgreich zusammengeführten übergeordneten Verzweigung, und löschen Sie dann die anfängliche lokale Verzweigung.

Wenn Ihre lokale Verzweigung X beispielsweise erfolgreich mit der Hauptverzweigung „OfficeDev/microsoftgraph/microsoft-graph-docs“ zusammengeführt wurde und Sie die zusammengeführten Inhalte weiter aktualisieren möchten. Erstellen Sie eine neue lokale Verzweigung X2 aus der Hauptverzweigung „OfficeDev/microsoftgraph/microsoft-graph-docs“. Öffnen Sie dazu GitBash, und führen Sie die folgenden Befehle aus:

    cd microsoft-graph-docs
    git pull upstream master:X2
    git push origin X2

Sie haben nun lokale Kopien (in einer neuen lokalen Verzweigung) der Arbeit, die Sie in Verzweigung X übermittelt haben. Die X2-Verzweigung enthält auch alle Arbeiten, die andere Autoren zusammengeführt haben. Wenn Ihre Arbeit also von der anderer Personen abhängt (z. B. freigegebene Bilder), ist sie in der neuen Verzweigung verfügbar. Sie können feststellen, ob sich Ihre vorherige Arbeit (und die anderer Personen) in der Verzweigung befindet, indem Sie die neue Verzweigung überprüfen ...

    git checkout X2

... und den Inhalt verifizieren. (Der `checkout`-Befehl aktualisiert die Dateien in `C:\Users\<yourusername>\microsoft-graph-docs` auf den aktuellen Status der X2-Verzweigung.) Sobald Sie die neue Verzweigung überprüft haben, können Sie Aktualisierungen am Inhalt vornehmen und wie gewohnt übernehmen. Um aber zu vermeiden, dass Sie versehentlich in der zusammengeführten Verzweigung (X) arbeiten, empfiehlt es sich, sie zu löschen (Informationen finden Sie im folgenden Abschnitt **Löschen einer Verzweigung**).

#### <a name="delete-a-branch"></a>Löschen einer Verzweigung

Nachdem Ihre Änderungen erfolgreich mit dem Hauptrepository zusammengeführt wurden, können Sie die verwendete Verzweigung löschen, da Sie sie nicht mehr benötigen.  Zusätzliche Aufgaben sollten in einer neuen Verzweigung vorgenommen werden.  

#### <a name="to-delete-a-branch"></a>So löschen Sie eine Verzweigung

1. Geben Sie an der Git Bash-Eingabeaufforderung `git checkout master` ein. Dadurch wird sichergestellt, dass Sie sich nicht in der zu löschenden Verzweigung befinden (das ist nicht zulässig).
2. Geben Sie an der Eingabeaufforderung `git branch -d <branch name>` ein: Dadurch wird die Verzweigung auf dem Computer nur dann gelöscht, wenn sie erfolgreich mit dem übergeordneten Repository zusammengeführt wurde. (Sie können dieses Verhalten mit dem `–D`-Flag außer Kraft setzen, allerdings sollten Sie sich in diesem Falle wirklich sicher sein.)
3. Geben Sie schließlich `git push origin :<branch name>` an der Befehlszeile ein (ein Leerzeichen vor dem Doppelpunkt, kein Leerzeichen dahinter).  Dadurch wird die Verzweigung in Ihrer Github-Verzweigung gelöscht.  

Herzlichen Glückwunsch, Sie haben erfolgreich am Projekt mitgewirkt.

## <a name="how-to-use-markdown-to-format-your-topic"></a>Verwenden von Markdown zum Formatieren Ihres Themas

### <a name="markdown"></a>Markdown

Alle Artikel in diesem Repository verwenden Markdown. Eine vollständige Einführung (und eine Auflistung aller Syntax) finden Sie unter [Daring Fireball - Markdown].

## <a name="faq"></a>Häufig gestellte Fragen

### <a name="how-do-i-get-a-github-account"></a>Wie erhalte ich ein GitHub-Konto?

Füllen Sie das Formular auf der Seite [Join GitHub](https://github.com/join) aus, um ein kostenloses GitHub-Konto zu eröffnen.

### <a name="where-do-i-get-a-contributors-license-agreement"></a>Wo erhalte ich einen Lizenzvertrag für Mitwirkende? 

Ihnen wird automatisch eine Benachrichtigung gesendet, dass Sie den Lizenzvertrag für Mitwirkende (Contributor's License Agreement, CLA) unterzeichnen müssen, wenn Ihre Pullanforderung dies erfordert.

Als Mitglied der Community **müssen Sie den Lizenzvertrag für Mitwirkende (CLA) unterzeichnen, bevor Sie umfangreiche Beiträge zu diesem Projekt leisten können**. Sie müssen die Dokumentation nur einmal ausfüllen und übermitteln. Überprüfen Sie das Dokument sorgfältig. Möglicherweise muss Ihr Arbeitgeber das Dokument unterzeichnen.

### <a name="what-happens-with-my-contributions"></a>Was geschieht mit meinen Beiträgen?

Wenn Sie Ihre Änderungen über eine Pullanforderung übermitteln, wird unser Team benachrichtigt und ihre Pullanforderung überprüfen. Sie erhalten Benachrichtigungen zu Ihrer Pullanforderung von GitHub; Sie werden möglicherweise auch von einer Person aus unserem Team benachrichtigt, wenn wir weitere Informationen benötigen. Wenn Ihre Pullanforderung genehmigt wurde, aktualisieren wir die Dokumentation. Wir behalten uns das Recht vor, Ihre Übermittlung aus Rechtlichen, Formaten, Klarheit oder anderen Gründen zu bearbeiten.

### <a name="can-i-become-an-approver-for-this-repositorys-github-pull-requests"></a>Kann ich ein Genehmiger für die Pullanforderungen in GitHub für dieses Repository werden?

Derzeit können externe Mitwirkende keine Pullanforderungen in diesem Repository genehmigen.

### <a name="how-soon-will-i-get-a-response-about-my-change-request"></a>Wie bald wird eine Antwort zu meiner Änderungsanforderung erhalten?

Pullanforderungen werden in der Regel innerhalb von 10 Werktagen bearbeitet.

## <a name="more-resources"></a>Weitere Ressourcen

* Weitere Informationen zu Markdown finden Sie auf der Website des Markdown-Erstellers [Daring Fireball].
* Weitere Informationen zur Verwendung von Git und GitHub finden Sie in der [GitHub-Hilfe.]

[GitHub Home]: http://github.com
[GitHub-Hilfe]: http://help.github.com/
[Einrichten von Git]: https://help.github.com/articles/set-up-git/
[Daring Fireball – Markdown]: http://daringfireball.net/projects/markdown/
[Daring Fireball]: http://daringfireball.net/
