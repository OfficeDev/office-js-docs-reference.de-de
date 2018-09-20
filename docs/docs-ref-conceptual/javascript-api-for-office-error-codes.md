# <a name="javascript-api-for-office-error-codes"></a>JavaScript API für Office – Fehlercodes

In diesem Artikel werden die Fehlermeldungen dokumentiert, die bei Verwendung der JavaScript-API für Office (Office.js) auftreten können.

**Gilt für:** Office-Add-Ins | SharePoint-Add-Ins | Excel | Outlook | PowerPoint | Project | Word

## <a name="error-codes"></a>Fehlercodes

In der folgenden Tabelle sind die Fehlercodes, Namen und Meldungen sowie die Bedingungen aufgeführt, auf die sie hinweisen.

|**Error.code**|**Error.name**|**Error.message**|**Condition**|
|:-----|:-----|:-----|:-----|
|1000|Ungültiger Koersionstyp|Der angegebene Koersionstyp wird nicht unterstützt.|Der Koersionstyp wird in der Hostanwendung nicht unterstützt. (Z. B. werden die Koersionstypen OOXML und HTML in Excel nicht unterstützt.)|
|1001|Fehler beim Lesen von Daten|Die aktuelle Auswahl wird nicht unterstützt.|Die aktuelle Auswahl des Benutzers wird nicht unterstützt. (D. h., sie entspricht nicht den unterstützten Koersionstypen.)|
|1002|Ungültiger Koersionstyp|Der angegebene Koersionstyp ist für diesen Bindungstyp nicht kompatibel.|Der Lösungsentwickler hat eine inkompatible Kombination von Koersionstyp und Bindungstyp angegeben.|
|1003|Fehler beim Lesen von Daten|Der angegebene rowCount- oder columnCount-Wert ist ungültig.|Der Benutzer gibt ungültige Spalten- oder Zeilenzahlen an.|
|1004|Fehler beim Lesen von Daten|Die aktuelle Auswahl ist nicht mit dem angegebenen Koersionstyp kompatibel.|Die aktuelle Auswahl wird von dieser Anwendung für den angegebenen Koersionstyp nicht unterstützt.|
|1005|Fehler beim Lesen von Daten|Der angegebene startRow- oder startColumn-Wert ist ungültig.|Der Benutzer gibt ungültige startRow- oder startCol-Werte an.|
|1006|Fehler beim Lesen von Daten|Koordinatenparameter können nicht mit dem Koersionstyp "Tabelle" verwendet werden, wenn die Tabelle verbundene Zellen enthält.|Der Benutzer versucht, partielle Daten aus einer nicht einheitlichen Tabelle (einer Tabelle mit verbundenen Zellen) abzurufen. |
|1007|Fehler beim Lesen von Daten|Das Dokument ist zu groß.|Der Benutzer versucht ein Dokument abzurufen, das größer als die derzeit unterstützte Größe ist.|
|1008|Fehler beim Lesen von Daten|Das angeforderte Dataset ist zu groß.|Der Benutzer fordert das Lesen von Daten jenseits der von den Host-Add-ins definierten Datenlimits an.|
|1009|Fehler beim Lesen von Daten|Der angegebene Dateityp wird nicht unterstützt.|Der Benutzer sendet einen ungültigen Dateityp.|
|2000|Fehler beim Schreiben von Daten|Der angegebene Datenobjekttyp wird nicht unterstützt. |Es wird ein nicht unterstütztes Datenobjekt angegeben.|
|2001|Fehler beim Schreiben von Daten|In die aktuelle Auswahl kann nicht geschrieben werden.|Die aktuelle Auswahl des Benutzers wird für einen Schreibvorgang nicht unterstützt (z. B. wenn der Benutzer ein Bild auswählt).|
|2002|Fehler beim Schreiben von Daten|Das angegebene Datenobjekt ist nicht kompatibel mit der Form oder den Dimensionen der aktuellen Auswahl.|Mehrere Zellen sind ausgewählt (und die Form der Auswahl stimmt nicht mit der Form der Daten überein).Mehrere Zellen sind ausgewählt (und die Dimensionen der Auswahl stimmen nicht mit den Dimensionen der Daten überein).|
|2003|Fehler beim Schreiben von Daten|Der "set"-Vorgang verursacht einen Fehler, weil das angegebene Datenobjekt Daten überschreibt.|Eine einzelne Zelle ist ausgewählt, und das angegebene Datenobjekt überschreibt Daten in der Arbeitsmappe.|
|2004|Fehler beim Schreiben von Daten|Das angegebene Datenobjekt entspricht nicht der Größe der aktuellen Auswahl.|Der Benutzer gibt ein Objekt an, das größer als die aktuelle Auswahl ist.|
|2005|Fehler beim Schreiben von Daten|Der angegebene startRow- oder startColumn-Wert ist ungültig.|Der Benutzer gibt ungültige startRow- oder startCol-Werte an.|
|2006|Fehler: Ungültiges Format|Das Format des angegebenen Datenobjekts ist nicht gültig.|Der Lösungsentwickler gibt eine ungültige HTML- oder OOXML-Zeichenfolge bzw. eine falsch formatierte HTML-Zeichenfolge an.|
|2007|Ungültiges Datenobjekt|Der Typ des angegebenen Datenobjekts ist nicht kompatibel mit der aktuellen Auswahl.|Der Lösungsentwickler gibt ein Datenobjekt an, das nicht kompatibel mit dem angegebenen Koersionstyp ist.|
|2008|Fehler beim Schreiben von Daten|Noch nicht festgelegt|Noch nicht festgelegt|
|2009|Fehler beim Schreiben von Daten|Das angegebene Datenobjekt ist zu groß.|Der Benutzer versucht, Daten jenseits der von Host-Add-ins definierten Datenlimits festzulegen.|
|2010|Fehler beim Schreiben von Daten|Koordinatenparameter können nicht mit dem Koersionstyp "Tabelle" verwendet werden, wenn die Tabelle verbundene Zellen enthält.|Der Benutzer versucht, partielle Daten aus einer nicht einheitlichen Tabelle (einer Tabelle mit verbundenen Zellen) festzulegen.|
|3000|Fehler beim Erstellen der Bindung|Eine Bindung an die aktuelle Auswahl ist nicht möglich.|Die Auswahl des Benutzers wird nicht zur Bindung unterstützt. (Der Benutzer wählt z. B. ein Bild oder ein anderes nicht unterstütztes Objekt aus.)|
|3001|Fehler beim Erstellen der Bindung|Noch nicht festgelegt|Noch nicht festgelegt|
|3002|Fehler: Ungültige Bindung|Die angegebene Bindung ist nicht vorhanden.|Der Entwickler versucht, eine Bindung an eine nicht vorhandene oder entfernte Bindung herzustellen.|
|3003|Fehler beim Erstellen der Bindung|Eine nicht zusammenhängende Auswahl wird nicht unterstützt.|Der Benutzer trifft eine mehrfache Auswahl.|
|3004|Fehler beim Erstellen der Bindung|Eine Bindung kann nicht mit der aktuellen Auswahl und dem angegebenen Bindungstyp erstellt werden.|Dies kann in verschiedenen Situationen auftreten. Weitere Informationen finden Sie unter "Fehlerbedingungen beim Erstellen der Bindung" weiter unten in diesem Artikel.|
|3005|Ungültiger Bindungsvorgang|Der Vorgang wird für diesen Bindungstyp nicht unterstützt.|Der Entwickler sendet einen Vorgang zum Hinzufügen einer Zeile oder Spalte für einen anderen Bindungstyp als  _Tabelle_.|
|3006|Fehler beim Erstellen der Bindung|Das benannte Element ist nicht vorhanden.|Das benannte Element wurde nicht gefunden. Es ist kein Inhaltssteuerelement und keine Tabelle mit diesem Namen vorhanden.|
|3007|Fehler beim Erstellen der Bindung|Es wurden mehrere Objekte mit dem gleichen Namen gefunden.|Konfliktfehler: Es sind mehrere Inhaltssteuerelemente mit dem gleichen Namen vorhanden, und die Fehlerbedingung bei Konflikt ist auf  **true** gesetzt.|
|3008|Fehler beim Erstellen der Bindung|Der angegebene Bindungstyp ist mit dem angegebenen benannten Element nicht kompatibel.|Das benannte Element kann nicht an den Typ gebunden werden. Beispielsweise enthält ein Inhaltssteuerelement Text, aber der Entwickler versuchte, eine Bindung an den Koersionstyp  _Tabelle_ herzustellen.|
|3009|Ungültiger Bindungsvorgang|Der Bindungstyp wird nicht unterstützt.|Dient der Abwärtskompatibilität.|
|3010|Nicht unterstützter Bindungsvorgang|Die ausgewählten Inhalte müssen im Tabellenformat vorliegen. Formatieren Sie die Daten als Tabelle, und versuchen Sie dann erneut.|Der Entwickler versucht, die Methoden  **addRowsAsynch** oder **deleteAllDataValuesAsynch** des **TableBinding**-Objekts für Daten vom Koersionstyp _matrix_ zu verwenden.|
|4000|Fehler beim Lesen von Einstellungen|Der angegebene Einstellungsname ist nicht vorhanden.|Es wird ein nicht vorhandener Einstellungsname angegeben.|
|4001|Fehler beim Speichern von Einstellungen|Die Einstellungen konnten nicht gespeichert werden.|Einstellungen konnten nicht gespeichert werden.|
|4002|Fehler: Veraltete Einstellungen|Einstellungen konnten nicht gespeichert werden, da sie veraltet sind.|Einstellungen sind veraltet, und der Entwickler hat angegeben, dass Einstellungen nicht überschrieben werden sollen.|
|5000|Fehler: Veraltete Einstellungen|Der Vorgang wird nicht unterstützt.|Der Vorgang wird vom aktuellen Host nicht unterstützt. Z. B. wird  **document.getSelectionAsync** von Outlook aus aufgerufen.|
|5001|Interner Fehler|Ein interner Fehler ist aufgetreten.|Verweist auf eine interne Fehlerbedingung, die aus einem der folgenden Gründe auftreten kann:<br/><table><tr><td>Ein Add-in, das von einem anderen Benutzer verwendet wird, der das Arbeitsblatt freigegeben hat, hat ungefähr zum gleichen Zeitpunkt eine Bindung erstellt, sodass Ihr Add-in versuchen muss, die Bindung erneut zu erstellen.</tr></td><tr><td>Es ist ein unbekannter Fehler aufgetreten.</tr></td><tr><td>Dieser Vorgang ist fehlgeschlagen.</tr></td><tr><td>Zugriff verweigert, da der Benutzer nicht Mitglied einer autorisierten Rolle ist.</tr></td><tr><td>Zugriff verweigert, da eine sichere, verschlüsselte Kommunikation erforderlich ist.</tr></td><tr><td>Die Daten sind veraltet und der Benutzer muss die Aktivierung der Abfragen bestätigen, um sie zu aktualisieren.</tr></td><tr><td>Die CPU-Auslastung der Websitesammlung wurde überschritten.</tr></td><tr><td>Das Speicherkontingent der Websitesammlung wurde überschritten.</tr></td><tr><td>Das Sitzungsspeicherkontingent wurde überschritten.</tr></td><tr><td>Die Arbeitsmappe weist einen ungültigen Status auf, sodass der Vorgang nicht durchgeführt werden kann.</tr></td><tr><td>In der Sitzung ist aufgrund von Inaktivität ein Timeout aufgetreten, der Benutzer muss die Arbeitsmappe neu laden.</tr></td><tr><td>Die maximale Anzahl zulässiger Sitzungen pro Benutzer wurde überschritten.</tr></td><tr><td>Der Vorgang wurde durch den Benutzer abgebrochen.</tr></td><tr><td>Der Vorgang konnte nicht abgeschlossen werden, da er zu lange dauert.</tr></td><tr><td>Die Anforderung konnte nicht abgeschlossen werden und muss erneut ausgeführt werden.</tr></td><tr><td>Der Testzeitraum des Produkts ist abgelaufen.</tr></td><tr><td>In der Sitzung ist aufgrund von Inaktivität ein Timeout aufgetreten.</tr></td><tr><td>Der Benutzer verfügt über keine Berechtigung, um den Vorgang für den angegebenen Bereich durchzuführen.</tr></td><tr><td>Die regionalen Einstellungen des Benutzers stimmen nicht mit der aktuellen Zusammenarbeitssitzung überein.</tr></td><tr><td>Der Benutzer ist nicht mehr verbunden, er muss die Arbeitsmappe aktualisieren oder erneut öffnen.</tr></td><tr><td>Der angeforderte Bereich ist in dem Arbeitsblatt nicht vorhanden.</tr></td><tr><td>Der Benutzer verfügt nicht über die Berechtigung, die Arbeitsmappe zu bearbeiten.</tr></td><tr><td>Die Arbeitsmappe kann nicht bearbeitet werden, da sie gesperrt ist.</tr></td><tr><td>Die Sitzung kann die Arbeitsmappe nicht automatisch speichern.</tr></td><tr><td>Die Sitzung kann die Sperre für die Arbeitsmappe nicht aktualisieren.</tr></td><tr><td>Die Anforderung kann nicht verarbeitet werden und muss erneut ausgeführt werden.</tr></td><tr><td>Die Anmeldeinformationen des Benutzers konnten nicht geprüft werden und müssen erneut eingegeben werden.</tr></td><tr><td>Dem Benutzer wurde der Zugriff verweigert.</tr></td><tr><td>Die freigegebene Arbeitsmappe muss aktualisiert werden.</tr></td></table>|
|5002|Berechtigung verweigert|Der angeforderte Vorgang ist im aktuellen Dokumentmodus nicht zulässig.|Der Lösungsentwickler sendet einen set-Vorgang, aber das Dokument befindet sich in einem Modus, der keine Änderungen zulässt, z. B. 'Bearbeitung einschränken'.|
|5003|Ereignis-Registrierungsfehler|Der angegebene Ereignistyp wird vom aktuellen Objekt nicht unterstützt.|Der Lösungsentwickler versucht, einen Handler für ein nicht vorhandenes Ereignis zu registrieren oder dessen Registrierung aufzuheben.|
|5004|Ungültiger API-Aufruf|Ungültiger API-Aufruf im aktuellen Kontext.|Es erfolgte ein für den Kontext ungültiger API-Aufruf (Z. B. wurde versucht, ein  **CustomXMLPart** -Objekt in Excel zu verwenden).|
|5005|Daten veraltet|Der Vorgang lieferte einen Fehler, da die Daten auf dem Server veraltet sind.|Die Daten auf dem Server müssen aktualisiert werden.|
|5006|Sitzungstimeout|In der Dokumentsitzung ist ein Timeout aufgetreten. Laden Sie das Dokument neu. |In der Sitzung ist ein Timeout aufgetreten.|
|5007|Ungültiger API-Aufruf|Die Aufzählung wird im aktuellen Kontext nicht unterstützt.|Die Aufzählung wird im aktuellen Kontext nicht unterstützt.|
|5009|Berechtigung verweigert|Zugriff verweigert.|Das Add-in verfügt nicht über die Berechtigung zum Aufruf der bestimmten API.|
|6000|Ungültiger Knoten|Der angegebene Knoten wurde nicht gefunden.|Der  **CustomXmlPart**-Knoten wurde nicht gefunden.|
|6100|Benutzerdefiniertes XML-Fehler|Benutzerdefiniertes XML-Fehler|Ungültiger API-Aufruf|
|7000|Ungültige ID|Die angegebene ID ist nicht vorhanden.|Ungültige ID|
|7001|Ungültige Navigation|Das Objekt befindet sich an einem Ort, an dem Navigation nicht unterstützt wird.|Der Benutzer kann das Objekt finden, aber nicht dorthin navigieren. (Z. B. erfolgt die Bindung in Word an die Kopfzeile, die Fußzeile oder einen Kommentar.)|
|7002|Ungültige Navigation|Das Objekt ist gesperrt oder geschützt.|Der Benutzer versucht, in einen gesperrten oder geschützten Bereich zu navigieren.|
|7004|Ungültige Navigation|Bei dem Vorgang ist ein Fehler aufgetreten, weil sich der Index außerhalb des gültigen Bereichs befindet.|Der Benutzer versucht, zu einem Index zu navigieren, der sich außerhalb des zulässigen Bereichs befindet.|
|8000|Fehlender Parameter|Die Tabellenzelle konnte nicht formatiert werden, da Parameterwerte fehlen. Überprüfen Sie die Parameter, und versuchen Sie es erneut.|Bei der cellFormat-Methode fehlen Parameter. Beispielsweise fehlen Parameter für "cells", "format" oder "tableOptions".|
|8010|Ungültiger Wert|Mindestens einer der "cells"-Parameter besitzt einen unzulässigen Wert. Überprüfen Sie die Werte, und versuchen Sie es erneut.|Die Referenzaufzählung der gemeinsamen Zellen ist nicht definiert. Beispiel: All, Data, Headers.|
|8011|Ungültiger Wert|Mindestens einer der "tableOptions"-Parameter besitzt einen unzulässigen Wert. Überprüfen Sie die Werte, und versuchen Sie es erneut.|Einer der Werte in „tableOptions“ ist ungültig.|
|8012|Ungültiger Wert|Mindestens einer der "format"-Parameter besitzt einen unzulässigen Wert. Überprüfen Sie die Werte, und versuchen Sie es erneut.|Einer der Werte im Format ist ungültig.|
|8020|Bereichsüberschreitung|Der Wert für den Zeilenindex liegt außerhalb des zulässigen Bereichs. Verwenden Sie einen positiven Wert (0 oder höher), der kleiner als die Anzahl der Zeilen ist.|Der Zeilenindex ist größer als der größte Zeilenindex der Tabelle oder kleiner als 0.|
|8021|Bereichsüberschreitung|Der Wert für den Spaltenindex liegt außerhalb des zulässigen Bereichs. Verwenden Sie einen positiven Wert (0 oder höher), der kleiner als die Anzahl der Spalten ist.|Der Spaltenindex ist größer als der größte Spaltenindex der Tabelle oder kleiner als 0.|
|8022|Bereichsüberschreitung|Der Wert liegt außerhalb des zulässigen Bereichs.|Einige Werte im Format liegen außerhalb des unterstützten Bereichs.|
|9016|Berechtigung verweigert|Berechtigung verweigert|Der Zugriff wurde verweigert.|
|12002|||Eine der folgenden Varianten:<br> – An der URL, die an`displayDialogAsync` übergeben wurde, ist keine Seite vorhanden.<br> – Die Seite, die an `displayDialogAsync` übergeben wurde, wurde geladen, aber das Dialogfeld wurde an eine Seite weitergeleitet, die nicht gefunden oder geladen werden konnte, oder es wurde an eine URL mit ungültiger Syntax weitergeleitet. Im Dialogfeld und Trigges ausgelöst eine `DialogEventReceived` Ereignis in der Hostseite.|
|12003|||Das Dialogfeld wurde an eine URL mit dem HTTP-Protokoll weitergeleitet. HTTPS ist erforderlich. Im Dialogfeld und Trigges ausgelöst eine `DialogEventReceived` Ereignis in der Hostseite.|
|12004|||Die Domäne der URL übergeben `displayDialogAsync` ist nicht vertrauenswürdig. Die Domäne muss derselben Domäne wie der Hostseite (einschließlich Protokoll-und Anschlussnummer). Wird ausgelöst, durch Aufruf von `displayDialogAsync`.|
|12005|||Die URL übergeben `displayDialogAsync` das HTTP-Protokoll verwendet. HTTPS ist erforderlich. Wird ausgelöst, durch Aufruf von `displayDialogAsync`. (In einigen Versionen von Office ist die Fehlermeldung zurückgegeben, die mit 12005 identisch für 12004 zurückgegebene.)|
|12006|||Das Dialogfeld wurde geschlossen, in der Regel, weil der Benutzer die Schaltfläche **X** auswählt. Im Dialogfeld und Trigges ausgelöst eine `DialogEventReceived` Ereignis in der Hostseite.|
|12007|||In diesem Hostfenster ist bereits ein Dialogfeld geöffnet. Ein Hostfenster, wie etwa einer Aufgabenbereich-App kann nur ein Dialogfeld gleichzeitig geöffnet haben. Wird ausgelöst, durch Aufruf von `displayDialogAsync`.|
|13000 - 13010|||Finden Sie unter [Ursachen und Behandlung von GetAccessTokenAsync](https://docs.microsoft.com/office/dev/add-ins/develop/troubleshoot-sso-in-office-add-ins#causes-and-handling-of-errors-from-getaccesstokenasync).|

## <a name="binding-creation-error-conditions"></a>Fehlerbedingungen beim Erstellen der Bindung

Wenn eine Bindung in der-API erstellt wird, geben Sie den Bindungstyp an, den Sie verwenden möchten. In den folgenden Tabellen sind die Bindungstypen und die daraus resultierenden Bindungsverhaltensweisen aufgeführt, die zu erwarten sind.

### <a name="behavior-in-excel"></a>Verhalten in Excel

In der folgenden Tabelle ist das Bindungsverhalten in Excel zusammengefasst.

|**Angegebener Bindungstyp**|**Tatsächliche Auswahl**|**Verhalten**|
|:-----|:-----|:-----|
|Matrix|Bereich von Zellen (innerhalb einer Tabelle und einzelne Zelle)|Eine Bindung vom Typ _matrix_ wird für die ausgewählten Zellen erstellt. Es wird keine Änderung im Dokument erwartet.|
|Matrix|Ausgewählter Text in der Zelle|Eine Bindung vom Typ _matrix_ wird für die gesamte Zelle erstellt. Es wird keine Änderung im Dokument erwartet.|
|Matrix|Mehrfachauswahl/ungültige Auswahl (z. B. wählt der Benutzer ein Bild, ein Objekt oder WordArt aus)|Die Bindung kann nicht erstellt werden.|
|Tabelle|Bereich von Zellen (schließt einzelne Zelle ein)|Die Bindung kann nicht erstellt werden.|
|Tabelle|Bereich von Zellen innerhalb einer Tabelle (schließt einzelne Zelle in einer Tabelle oder die gesamte Tabelle oder Text in einer Zelle einer Tabelle ein)|In der gesamten Tabelle wird eine Bindung erstellt.|
|Tabelle|Auswahl zur Hälfte innerhalb und zur Hälfte außerhalb der Tabelle|Die Bindung kann nicht erstellt werden.|
|Tabelle|Ausgewählter Text in der Zelle (nicht in der Tabelle)|Die Bindung kann nicht erstellt werden.|
|Tabelle|Mehrfachauswahl/ungültige Auswahl (Z. B. wählt der Benutzer ein Bild, ein Objekt, WordArt usw. aus)|Die Bindung kann nicht erstellt werden.|
|Text|Bereich von Zellen|Die Bindung kann nicht erstellt werden.|
|Text|Bereich von Zellen in einer Tabelle|Die Bindung kann nicht erstellt werden.|
|Text|Einzelne Zelle|Eine Bindung vom Typ  _text_ wird erstellt.|
|Text|Einzelne Zelle in einer Tabelle|Eine Bindung vom Typ  _text_ wird erstellt.|
|Text|Ausgewählter Text in der Zelle|Eine Bindung vom Typ  _text_ in der gesamten Zelle wird erstellt.|

### <a name="behavior-in-word"></a>Verhalten in Word

In der folgenden Tabelle ist das Bindungsverhalten in Word zusammengefasst.

|**Angegebener Bindungstyp**|**Tatsächliche Auswahl**|**Verhalten**|
|:-----|:-----|:-----|
|Matrix|Text|Die Bindung kann nicht erstellt werden.|
|Matrix|Gesamte Tabelle|Eine Bindung vom Typ  _matrix_ wird erstellt.Das Dokument wird geändert, und die Tabelle muss in ein Inhaltssteuerelement eingeschlossen werden. |
|Matrix|Bereich in einer Tabelle|Die Bindung kann nicht erstellt werden.|
|Matrix|Ungültige Auswahl (z. B. mehrere, ungültige Objekte usw.)|Die Bindung kann nicht erstellt werden.|
|Tabelle|Text|Die Bindung kann nicht erstellt werden.|
|Tabelle|Gesamte Tabelle|Eine Bindung vom Typ  _text_ wird erstellt.|
|Tabelle|Bereich in einer Tabelle|Die Bindung kann nicht erstellt werden.|
|Tabelle|Ungültige Auswahl (z. B. mehrere, ungültige Objekte usw.)|Die Bindung kann nicht erstellt werden.|
|Text|Gesamte Tabelle|Eine Bindung vom Typ  _text_ wird erstellt.|
|Text|Bereich in einer Tabelle|Die Bindung kann nicht erstellt werden.|
|Text|Mehrfachauswahl|Die letzte Auswahl wird in ein Inhaltssteuerelement und eine Bindung an dieses Steuerelement eingeschlossen. Ein Inhaltssteuerelement vom Typ  _text_ wird erstellt.|
|Text|Ungültige Auswahl (z. B. mehrere, ungültige Objekte usw.)|Die Bindung kann nicht erstellt werden.|

## <a name="see-also"></a>Siehe auch
   
- [Entwicklungslebenszyklus von Office-Add-Ins](https://docs.microsoft.com/office/dev/add-ins/concepts/add-in-development-lifecycle)
    
