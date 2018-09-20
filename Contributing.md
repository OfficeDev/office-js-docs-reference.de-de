# <a name="contribute-to-this-documentation"></a>An dieser Dokumentation mitwirken

Vielen Dank f?r Ihr Interesse an unserer Dokumentation!

* [Möglichkeiten zur Mitwirkung](#ways-to-contribute)
* [Mit GitHub mitwirken](#contribute-using-github)
* [Mit Git mitwirken](#contribute-using-git)
* [Verwenden von Markdown zum Formatieren Ihres Themas](#how-to-use-markdown-to-format-your-topic)
* [H?ufig gestellte Fragen](#faq)
* [Weitere Ressourcen](#more-resources)

## <a name="ways-to-contribute"></a>Möglichkeiten zur Mitwirkung

Nachfolgend finden Sie einige M?glichkeiten, wie Sie zu dieser Dokumentation beitragen k?nnen:

* Lesen Sie die Informationen unter [Mit GitHub mitwirken](#contribute-using-github), um kleine ?nderungen an einem Artikel vorzunehmen.
* Lesen Sie die Informationen unter [Mit Git mitwirken](#contribute-using-git), um umfangreiche ?nderungen vorzunehmen bzw. ?nderungen vorzunehmen, die mit Code zu tun haben.
* Melden von Fehlern Dokumentation über GitHub Probleme.
* Fordern Sie die neue Dokumentation auf der Website [UserVoice der Office Developer-Plattform](http://officespdev.uservoice.com) an.

## <a name="contribute-using-github"></a>Mit GitHub mitwirken

Verwenden Sie GitHub, um an dieser Dokumentation mitzuwirken, ohne dass Sie hierf?r das Repo auf Ihrem Desktop klonen. Dies ist die einfachste M?glichkeit zum Erstellen einer Pullanforderung in diesem Repository. Verwenden Sie diese Methode, um eine kleinere ?nderung vorzunehmen, mit der keine Code?nderungen einhergehen. 

**Hinweis** Mit dieser Methode k?nnen Sie jeweils zu einem einzelnen Artikel beitragen.

### <a name="to-contribute-using-github"></a>So k?nnen Sie mit GitHub mitwirken

1. Suchen Sie den Artikel, den Sie bei GitHub beisteuern m?chten.
2. Wenn Sie sich in dem Artikel bei GitHub befinden, melden Sie sich bei GitHub an (eröffnen Sie ein kostenloses Konto über [Join GitHub](https://github.com/join)).
3. Wählen Sie das **Stiftsymbol** aus (bearbeiten Sie die Datei in Ihrer Verzweigung des Projekts), und nehmen Sie Ihre Änderungen im Fenster **<>Edit file** vor. 
4. F?hren Sie einen Bildlauf nach unten durch, und geben Sie eine Beschreibung ein.
5. W?hlen Sie **Propose file change**>**Create pull request** aus.

Sie haben nun erfolgreich eine Pullanforderung ?bermittelt. Pullanforderungen werden in der Regel innerhalb von 10 Werktagen bearbeitet. 


## <a name="contribute-using-git"></a>Mit Git mitwirken

Verwenden Sie Git, um umfangreiche ?nderungen vorzunehmen, z. B.:

* Beisteuern von Code
* Beisteuern von ?nderungen, die Auswirkungen auf die Bedeutung haben
* Beisteuern von umfangreichen ?nderungen an Text
* Hinzuf?gen neuer Themen

### <a name="to-contribute-using-git"></a>So k?nnen Sie mit Git mitwirken

1. Wenn Sie nicht ?ber ein GitHub-Konto verf?gen, richten Sie eines auf [GitHub](https://github.com/join) ein. 
2. Nachdem Sie ein Konto erstellt haben, installieren Sie Git auf Ihrem Computer. Führen Sie die Schritte in diesem Lernprogramm [Git einrichten] .
3. Um eine Pullanforderung mit Git zu übermitteln, führen Sie die Schritte unter [Verwenden von GitHub, Git und diesem Repository](#use-github-git-and-this-repository) durch.
4. Sie werden aufgefordert, den Lizenzvertrag f?r Mitwirkende zu unterzeichnen, wenn Sie:

    * Mitglied der Microsoft Open Technologies-Gruppe sind.
    * Ein Mitwirkender sind, der nicht für Microsoft arbeitet.

Als Mitglied der Community müssen Sie den Lizenzvertrag für Mitwirkende (CLA) unterzeichnen, bevor Sie umfangreiche Beiträge zu einem Projekt leisten können. Sie müssen die Dokumentation nur einmal ausfüllen und übermitteln. Überprüfen Sie das Dokument sorgfältig. Möglicherweise muss Ihr Arbeitgeber das Dokument unterzeichnen.

Signieren der CLA gewährt Ihnen Rechte zum Haupt-Repository commit für, aber es bedeutet, dass die Office-Entwickler und Office Developer Content Publishing-Teams können, überprüfen und genehmigen Ihre Beiträge. Sie werden Ihre Übermittlungen angerechnet.

Pullanforderungen werden in der Regel innerhalb von 10 Werktagen bearbeitet.

## <a name="use-github-git-and-this-repository"></a>Verwenden von GitHub, Git und diesem Repository

**Hinweis:** Die meisten der Informationen in diesem Abschnitt finden Sie in [GitHub-Hilfe]-Artikeln.  Wenn Sie mit Git und GitHub vertraut sind, fahren Sie mit dem Abschnitt zum **Einsenden und Bearbeiten von Inhalten** mit speziellen Informationen zum Code-/Inhaltsfluss dieses Repositorys fort.

### <a name="to-set-up-your-fork-of-the-repository"></a>So richten Sie Ihre Verzweigung des Repositorys ein

1.  Richten Sie ein GitHub-Konto ein, damit Sie an diesem Projekt mitwirken können. Falls noch nicht geschehen, wechseln Sie zu [GitHub](https://github.com/join), und richten Sie jetzt ein Konto ein.
2.  Installieren Sie Git auf Ihrem Computer. Führen Sie die Schritte in diesem Lernprogramm [Git einrichten] .
3.  Erstellen Sie Ihre eigene Verzweigung dieses Repositorys. Wählen Sie dazu am oberen Rand der Seite die Schaltfläche **Verzweigung** .
4.  Kopieren Sie die Verzweigung auf Ihren Computer. Öffnen Sie hierzu Git Bash. Geben Sie an der Eingabeaufforderung Folgendes ein:

        git clone https://github.com/<your user name>/<repo name>.git

    Erstellen Sie danach einen Verweis auf das Stammrepository durch Eingabe der folgenden Befehle:

        cd <repo name>
        git remote add upstream https://github.com/OfficeDev/<repo name>.git
        git fetch upstream

Ausgezeichnet. Sie haben nun Ihr Repository eingerichtet. Sie müssen diese Schritte nicht erneut ausführen.

### <a name="contribute-and-edit-content"></a>Einsenden und Bearbeiten von Inhalten

Um die Mitwirkung f?r Sie so einfach wie m?glich zu machen, f?hren Sie die folgenden Schritte aus.

#### <a name="to-contribute-and-edit-content"></a>So senden Sie Inhalte ein und bearbeiten diese

1. Erstellen Sie eine neue Verzweigung.
2. Fügen Sie neue Inhalte hinzu, oder bearbeiten Sie vorhandene Inhalte.
3. Senden Sie eine Pull-Anforderung an das Hauptrepository.
4. Löschen Sie die Verzweigung.

**Wichtig** Begrenzen Sie jede Verzweigung auf ein einzelnes Konzept bzw. einen einzelnen Artikel, um den Workflow zu optimieren und das Risiko von Zusammenführungskonflikten zu verringern. Zu geeigneten Inhalten für eine neue Verzweigung gehören:

* Ein neuer Artikel.
* Bearbeitung von Rechtschreibung und Grammatik.
* Vornehmen einer einzelnen Formatierungs?nderung f?r eine gro?e Gruppe von Artikeln (z. B. ?bernehmen einer neuen Copyright-Fu?zeile).

#### <a name="to-create-a-new-branch"></a>So erstellen Sie eine neue Verzweigung

1.  ?ffnen Sie Git Bash.
2.  Geben Sie an der Git Bash-Eingabeaufforderung `git pull upstream master:<new branch name>` ein. Dadurch wird eine neue Verzweigung lokal erstellt, die aus der neuesten OfficeDev-Hauptverzweigung kopiert wird.
3.  Geben Sie an der Git Bash-Eingabeaufforderung `git push origin <new branch name>` ein. Dadurch wird GitHub auf die neue Verzweigung hingewiesen. Die neue Verzweigung sollte nun in Ihrer Verzweigung des Repository auf GitHub angezeigt werden.
4.  Geben Sie in der Git Bash-Befehlszeile `git checkout <new branch name>` ein, um zu Ihrer neuen Verzweigung umzuschalten.

#### <a name="add-new-content-or-edit-existing-content"></a>Hinzufügen neuer Inhalte oder Bearbeiten vorhandener Inhalte

Sie navigieren mithilfe des Datei-Explorers zu dem Repository auf dem Computer. Die Repositorydateien befinden sich in `C:\Users\<yourusername>\<repo name>`.

Zum Bearbeiten von Dateien ?ffnen Sie sie in einem Editor Ihrer Wahl und ?ndern sie. Um eine neue Datei zu erstellen, verwenden Sie den Editor Ihrer Wahl, und speichern Sie die neue Datei im entsprechenden Speicherort in Ihrer lokalen Kopie des Repositorys. Speichern Sie w?hrend der Bearbeitung h?ufig.

Die Dateien in `C:\Users\<yourusername>\<repo name>` sind eine Arbeitskopie der neuen Verzweigung, die Sie in Ihrem lokalen Repository erstellt haben. Änderungen in diesem Ordner wirken sich erst auf das lokale Repository aus, wenn Sie eine Änderung übernehmen. Um eine Änderung am lokalen Repository zu übernehmen, geben Sie die folgenden Befehle in GitBash ein:

    git add .
    git commit -v -a -m "<Describe the changes made in this commit>"

Der `add`-Befehl fügt Ihre Änderungen einem Stagingbereich als Vorbereitung für die Aufnahme im Repository hinzu. Der Punkt nach dem `add`-Befehl gibt an, dass Sie alle Dateien, die Sie hinzugefügt oder geändert haben, übernehmen möchten und dass Unterordner rekursiv überprüft werden sollen. (Wenn Sie nicht alle Änderungen übernehmen möchten, können Sie bestimmte Dateien hinzufügen. Sie können den Vorgang auch rückgängig machen. Falls Sie Hilfe benötigen, geben Sie `git add -help` oder `git status` ein.)

Der Befehl `commit` wendet die bereitgestellten Änderungen auf das Repository an. Der Schalter `-m` bedeutet, dass Sie den Commit-Kommentar in der Befehlszeile angeben. Die Optionen -v und -a können weggelassen werden. Die Option -v ist für die ausführliche Ausgabe des Befehls, und -a entspricht in ihrer Funktion den bereits verwendeten Befehl zum Hinzufügen.

Sie k?nnen w?hrend der Arbeit mehrfach ?bernehmen oder warten und erst nach Abschluss der Arbeit ?bernehmen.

#### <a name="submit-a-pull-request-to-the-main-repository"></a>?bermitteln einer Pullanforderung an das Hauptrepository

Wenn Sie Ihre Arbeit abgeschlossen haben und bereit sind, sie mit dem zentralen Repository zusammenzuf?hren, gehen Sie folgenderma?en vor.

#### <a name="to-submit-a-pull-request-to-the-main-repository"></a>So ?bermitteln Sie eine Pullanforderung an das Hauptrepository

1.  Geben Sie an der Git Bash-Eingabeaufforderung `git push origin <new branch name>` ein. Im lokalen Repository bezieht sich `origin` auf Ihr GitHub-Repository, aus dem Sie das lokale Repository geklont haben. Mit diesem Befehl wird der aktuelle Status Ihrer neuen Verzweigung, einschließlich aller im vorherigen Schritt vorgenommenen Commits, in Ihre GitHub-Verzweigung übertragen.
2.  Navigieren Sie auf der GitHub-Website in Ihrer Verzweigung zur neuen Verzweigung.
3.  Klicken Sie oben auf der Seite auf die Schaltfl?che **Pullanforderung**.
4.  Stellen Sie sicher, dass die untere Verzweigung `OfficeDev/<repo name>@master` und die obere Verzweigung `<your username>/<repo name>@<branch name>` lautet.
5.  Klicken Sie auf die Schaltfl?che **Commitbereich aktualisieren**.
6.  Geben Sie Ihrer Pullanforderung einen Titel, und beschreiben Sie alle ?nderungen, die Sie vornehmen.
7.  Senden Sie die Pull-Anforderung.

Einer der Websiteadministratoren wird nun Ihre Pullanforderung verarbeiten. Ihre Pullanforderung wird auf der OfficeDev/<repo name>-Website unter „Probleme“ angezeigt. Wenn die Pullanforderung angenommen wird, wird das Problem gelöst.

#### <a name="create-a-new-branch-after-merge"></a>Erstellen einer neuen Verzweigung nach dem Zusammenführen

Nachdem eine Verzweigung erfolgreich zusammengef?hrt wurde (d. h. Ihre Pullanforderung wurde angenommen), arbeiten Sie nicht in der lokalen Verzweigung weiter. Dies kann zu Zusammenf?hrungskonflikten f?hren, wenn Sie eine weitere Pullanforderung einsenden. Wenn Sie eine weitere Aktualisierung vornehmen m?chten, erstellen Sie eine neue lokale Verzweigung aus der erfolgreich zusammengef?hrten ?bergeordneten Verzweigung, und l?schen Sie dann die anf?ngliche lokale Verzweigung.

Wenn Ihre lokale Verzweigung X beispielsweise erfolgreich mit der Hauptverzweigung ?OfficeDev/microsoftgraph/microsoft-graph-docs? zusammengef?hrt wurde und Sie die zusammengef?hrten Inhalte weiter aktualisieren m?chten. Erstellen Sie eine neue lokale Verzweigung X2 aus der Hauptverzweigung ?OfficeDev/microsoftgraph/microsoft-graph-docs?. ?ffnen Sie dazu GitBash, und f?hren Sie die folgenden Befehle aus:

    cd microsoft-graph-docs
    git pull upstream master:X2
    git push origin X2

Sie haben nun lokale Kopien (in einer neuen lokalen Verzweigung) der Arbeit, die Sie in Verzweigung X ?bermittelt haben. Die X2-Verzweigung enth?lt auch alle Arbeiten, die andere Autoren zusammengef?hrt haben. Wenn Ihre Arbeit also von der anderer Personen abh?ngt (z. B. freigegebene Bilder), ist sie in der neuen Verzweigung verf?gbar. Sie k?nnen feststellen, ob sich Ihre vorherige Arbeit (und die anderer Personen) in der Verzweigung befindet, indem Sie die neue Verzweigung ?berpr?fen ...

    git checkout X2

... und den Inhalt verifizieren. (Der `checkout`-Befehl aktualisiert die Dateien in `C:\Users\<yourusername>\microsoft-graph-docs` auf den aktuellen Status der X2-Verzweigung.) Sobald Sie die neue Verzweigung überprüft haben, können Sie Aktualisierungen am Inhalt vornehmen und wie gewohnt übernehmen. Um aber zu vermeiden, dass Sie versehentlich in der zusammengeführten Verzweigung (X) arbeiten, empfiehlt es sich, sie zu löschen (Informationen finden Sie im folgenden Abschnitt **Löschen einer Verzweigung**).

#### <a name="delete-a-branch"></a>Löschen einer Verzweigung

Nachdem Ihre ?nderungen erfolgreich mit dem Hauptrepository zusammengef?hrt wurden, k?nnen Sie die verwendete Verzweigung l?schen, da Sie sie nicht mehr ben?tigen.  Zus?tzliche Aufgaben sollten in einer neuen Verzweigung vorgenommen werden.  

#### <a name="to-delete-a-branch"></a>So l?schen Sie eine Verzweigung

1.  Geben Sie an der Git Bash-Eingabeaufforderung `git checkout master` ein. Dadurch wird sichergestellt, dass Sie sich nicht in der zu löschenden Verzweigung befinden (das ist nicht zulässig).
2.  Geben Sie an der Eingabeaufforderung `git branch -d <branch name>` ein: Dadurch wird die Verzweigung auf dem Computer nur dann gelöscht, wenn sie erfolgreich mit dem übergeordneten Repository zusammengeführt wurde. (Sie können dieses Verhalten mit dem `–D`-Flag außer Kraft setzen, allerdings sollten Sie sich in diesem Falle wirklich sicher sein.)
3.  Geben Sie schließlich `git push origin :<branch name>` an der Befehlszeile ein (ein Leerzeichen vor dem Doppelpunkt, kein Leerzeichen dahinter).  Dadurch wird die Verzweigung in Ihrer Github-Verzweigung gelöscht.  

Herzlichen Gl?ckwunsch, Sie haben erfolgreich am Projekt mitgewirkt.

## <a name="how-to-use-markdown-to-format-your-topic"></a>Verwenden von Markdown zum Formatieren Ihres Themas

### <a name="markdown"></a>Markdown

Alle Artikel in diesem Repository verwenden Markdown. Eine vollständige Einführung (und die Liste der alle Syntax) finden Sie unter [Daring Fireball - Abzugsverteilung(en)].
 
## <a name="faq"></a>H?ufig gestellte Fragen

### <a name="how-do-i-get-a-github-account"></a>Wie erhalte ich ein GitHub-Konto?

F?llen Sie das Formular auf der Seite [Join GitHub](https://github.com/join) aus, um ein kostenloses GitHub-Konto zu er?ffnen. 

### <a name="where-do-i-get-a-contributors-license-agreement"></a>Wo erhalte ich einen Lizenzvertrag f?r Mitwirkende? 

Ihnen wird automatisch eine Benachrichtigung gesendet, dass Sie den Lizenzvertrag f?r Mitwirkende (Contributor's License Agreement, CLA) unterzeichnen m?ssen, wenn Ihre Pullanforderung dies erfordert. 

Als Mitglied der Community **müssen Sie den Lizenzvertrag für Mitwirkende (CLA) unterzeichnen, bevor Sie umfangreiche Beiträge zu diesem Projekt leisten können**. Sie müssen die Dokumentation nur einmal ausfüllen und übermitteln. Überprüfen Sie das Dokument sorgfältig. Möglicherweise muss Ihr Arbeitgeber das Dokument unterzeichnen.

### <a name="what-happens-with-my-contributions"></a>Was geschieht mit meinen Beitr?gen?

Wenn Sie ?nderungen ?ber eine Pullanforderung ?bermitteln, wird unser Team benachrichtigt und Ihre Pullanforderung wird von diesem Team bearbeitet. Sie erhalten Benachrichtigungen zu Ihrer Pullanforderung von GitHub. M?glicherweise werden Sie auch von einer Person in unserem Team benachrichtigt, wenn Sie weitere Informationen ben?tigen. Wenn Ihre Pullanforderung genehmigt wird, wird die Dokumentation aktualisiert. Wir behalten uns das Recht vor, Ihre ?bermittlung im Hinblick auf rechtliche und stilistische Aspekte sowie Klarheit oder andere Probleme zu bearbeiten.

### <a name="can-i-become-an-approver-for-this-repositorys-github-pull-requests"></a>Kann ich ein Genehmiger f?r die Pullanforderungen in GitHub f?r dieses Repository werden?

Derzeit k?nnen externe Mitwirkende keine Pullanforderungen in diesem Repository genehmigen.

### <a name="how-soon-will-i-get-a-response-about-my-change-request"></a>Wie schnell erhalte ich eine Antwort zu meiner ?nderungsanforderung?

Pullanforderungen werden in der Regel innerhalb von 10 Werktagen bearbeitet.


## <a name="more-resources"></a>Weitere Ressourcen

* Weitere Informationen zum Abzugsverteilung(en) finden Sie in den Abzugsverteilung(en) Ersteller Website [Daring Fireball].
* Weitere Informationen zum Verwenden von Git und GitHub überprüfen Sie zuerst die [GitHub Hilfe].

[GitHub Home]: http://github.com
[GitHub-Hilfe]: http://help.github.com/
[Einrichten von Git]: https://help.github.com/articles/set-up-git/
[Daring Fireball - Abzugsverteilung(en)]: http://daringfireball.net/projects/markdown/
[Daring Fireball]: http://daringfireball.net/
