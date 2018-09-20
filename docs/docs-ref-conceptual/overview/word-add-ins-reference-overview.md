# <a name="word-javascript-api-overview"></a>Word-JavaScript-API (Übersicht)

Word bietet eine umfangreiche Sammlung von APIs, die Sie zum Erstellen von Add-Ins verwenden können, die mit Dokumentinhalten und Metadaten interagieren. Verwenden Sie diese APIs, um überzeugende Oberflächen zu erstellen, die mit Word interagieren und Word erweitern. Sie können Inhalte importieren und exportieren, neue Dokumente aus unterschiedlichen Datenquellen zusammenstellen und mit Dokumentworkflows interagieren, um benutzerdefinierte Dokumentlösungen zu erstellen.

Sie können zwei JavaScript-APIs verwenden, um mit den Objekten und Metadaten in einem Word-Dokument zu interagieren:

- JavaScript-API für Word - In Office 2016 eingeführt.
- [JavaScript-API für Office](../javascript-api-for-office.md) (Office.js) - In Office 2013 eingeführt.

## <a name="word-javascript-api"></a>JavaScript-API für Word

Die JavaScript-API für Word wird von Office.js geladen. Sie ändert die Art der Interaktion mit Objekten wie Dokumenten und Absätzen. Statt einzelne asynchrone APIs zum Abrufen und Aktualisieren der einzelnen Objekte bereitzustellen, bietet die JavaScript-API für Word „Proxy“-JavaScript-Objekte, die den realen Objekten entsprechen, die in Word ausgeführt werden. Sie können mit diesen Proxyobjekten interagieren, indem Sie ihre Eigenschaften synchron lesen und synchrone Methoden aufrufen, um dazu Vorgänge durchzuführen. Diese Interaktionen mit Proxyobjekten werden nicht sofort im ausgeführten Skript umgesetzt, deshalb stellen wir eine Methode zum Kontext mit dem Namen sync() bereit. Die **context.sync**-Methode synchronisiert den Status zwischen ausgeführten JavaScript-Objekten und den realen Objekten in Office, indem in die Warteschlange eingereihte Anweisungen ausgeführt und Eigenschaften von geladenen Word-Objekten für die Verwendung in Ihrem Skript abgerufen werden.

## <a name="javascript-api-for-office"></a>JavaScript-API für Office

Sie können von den folgenden Speicherorten auf Office.js verweisen:

* https://appsforoffice.microsoft.com/lib/1/hosted/office.js-Verwenden Sie diese Ressource für Produktion-add-ins.
* https://appsforoffice.microsoft.com/lib/beta/hosted/office.js-Verwenden Sie diese Ressource aus, wenn Sie die Preview-Features möchten.

Wenn Sie [Visual Studio](https://www.visualstudio.com/products/free-developer-offers-vs) verwenden, können Sie die [Office-Entwicklertools](https://www.visualstudio.com/features/office-tools-vs.aspx) herunterladen, um Projektvorlagen zu erhalten, die Office.js enthalten.  Sie können auch [nuget verwenden, um Office.js zu erhalten](https://www.nuget.org/packages/Microsoft.Office.js/).

Wenn Sie TypeScript verwenden und über npm verfügen, können Sie die TypeScript-Definitionen durch Eingeben des folgenden Texts in der Befehlszeilenoberfläche abrufen: `typings install office-js --ambient`.

## <a name="running-word-add-ins"></a>Ausführen von Word-Add-Ins

Verwenden Sie einen Office.initialize-Ereignishandler, um Ihr Add-In auszuführen. Unter [Grundlegendes zur API](https://docs.microsoft.com/office/dev/add-ins/develop/understanding-the-javascript-api-for-office) finden Sie weitere Informationen zur Add-In-Initialisierung.

Add-Ins für Word 2016 werden durch Übergeben einer Funktion an die **Word.run()**-Methode ausgeführt . Die Funktion in der **run**-Methode muss ein Kontextargument besitzen. Dieses [Kontextobjekt](/javascript/api/word/word.requestcontext) unterscheidet sich dem Kontextobjekt, das Sie aus dem Office-Objekt abrufen, obwohl es ebenfalls für die Interaktion mit der Word-Laufzeitumgebung verwendet wird. Das folgenden Beispiel zeigt, wie ein Word-Add-In mithilfe der **Word.run()**-Methode initialisiert und ausgeführt wird.

```js
(function () {
    "use strict";

    // The initialize event handler must be run on each page to initialize Office JS.
    // You can add optional custom initialization code that will run after OfficeJS
    // has initialized.
    Office.initialize = function (reason) {
        // The reason object tells how the add-in was initialized. The values can be:
        // inserted - the add-in was inserted to an open document.
        // documentOpened - the add-in was already inserted in to the document and the document was opened.

        // Checks for the DOM to load using the jQuery ready function.
        $(document).ready(function () {
            // Set your optional initialization code.
            // You can also load saved settings from the Office object.
        });
    };

    // Run a batch operation against the Word JavaScript API object model.
    // Use the context argument to get access to the Word document.
    Word.run(function (context) {

        // Create a proxy object for the document.
        var thisDocument = context.document;
        // ...
    })
})();
```

### <a name="synchronizing-word-documents-with-word-javascript-api-proxy-objects"></a>Synchronisieren von Word-Dokumenten mit Proxyobjekten der Word-JavaScript-API

Das Objektmodell der Word-JavaScript-API ist lose mit den Objekten in Word verknüpft. Die Objekte der Word--JavaScript-API sind Proxys für Objekte in einem Word-Dokument. Aktionen, die für Proxyobjekte ausgeführt werden, werden in Word erst umgesetzt, wenn der Dokumentstatus synchronisiert wurde. Der Status des Word-Dokuments wird in den Proxyobjekten jedoch erst umgesetzt, wenn der Status des Dokuments synchronisiert wurde. Um den Dokumentstatus zu synchronisieren, führen Sie die **context.sync()**-Methode aus. Das folgende Beispiel veranschaulicht die Erstellung eines Proxytextobjekts und einen Befehl in der Warteschlange, um die Texteigenschaft zu dem Proxytextobjekt zu laden, und verwendet die **context.sync()**-Methode, um den Text im Word-Dokument mit dem Textproxyobjekt zu synchronisieren.

```js
// Run a batch operation against the Word object model.
Word.run(function (context) {

    // Create a proxy object for the document body.
    // The body object hasn't been set with any property values.
    var body = context.document.body;

    // Queue a command to load the text property for the proxy document body object.
    context.load(body, 'text');

    // Synchronize the document state by executing the queued commands,
    // and return a promise to indicate task completion.
    return context.sync().then(function () {
        console.log("Body contents: " + body.text);
    });
})
```

### <a name="executing-a-batch-of-commands"></a>Ausführen einer Reihe von Befehlen

Die Word-Proxyobjekte besitzen Methoden für den Zugriff und die Aktualisierung des Objektmodells. Diese Methoden werden sequenziell in der Reihenfolge ausgeführt, in der sie alle im Batch in die Warteschlange gestellt wurden. Alle Befehle in der Warteschlange im Batch werden ausgeführt, wenn context.sync() aufgerufen wird.

Im folgenden Beispiel wird gezeigt, wie die Warteschlange von Befehlen funktioniert. Wenn **context.sync()** aufgerufen wird, wird der Befehl zum Laden des Textkörpers in Word ausgeführt. Dann wird der Befehl zum Einfügen von Text in Word ausgeführt. Die Ergebnisse werden im Textproxyobjekt zurückgegeben. Der Wert der **body.text**-Eigenschaft in der Word-JavaScript-API ist der Wert des Word-Dokumenttexts <u>bevor</u> der Text in das Word-Dokument eingefügt wurde.


```js
// Run a batch operation against the Word JavaScript API.
Word.run(function (context) {

    // Create a proxy object for the document body.
    var body = context.document.body;

    // Queue a command to load the text property of the proxy body object.
    context.load(body, 'text');

    // Queue a command to insert text into the end of the Word document body.
    body.insertText('This is text inserted after loading the body.text property',
                    Word.InsertLocation.end);

    // Synchronize the document state by executing the queued commands,
    // and return a promise to indicate task completion.
    return context.sync().then(function () {
        console.log("Body contents: " + body.text);
    });
})
```

## <a name="word-javascript-api-open-specifications"></a>Offene Spezifikationen Word JavaScript-API

Während des Entwerfens und Entwickelns neuer APIs für Word-Add-Ins stellen wir diese zur Verfügung, damit Sie auf der Seite [Offene API-Spezifikationen](../openspec.md) Ihr Feedback abgeben können. Erfahren Sie, welche neuen Funktionen für die Word-JavaScript-APIs geplant sind, und teilen Sie uns Ihre Meinung zu unseren Designspezifikationen mit.

## <a name="word-javascript-api-reference"></a>JavaScript-API-Referenz für Word

Ausführliche Informationen zu der JavaScript-API für Word finden Sie in der [Referenzdokumentation für die Word-JavaScript-API](/javascript/api/word).

## <a name="see-also"></a>Siehe auch

* [?bersicht ?ber Word-Add-Ins](https://docs.microsoft.com/office/dev/add-ins/word/word-add-ins-programming-overview)
* [Office-Add-Ins-Plattformübersicht](https://docs.microsoft.com/office/dev/add-ins/overview/office-add-ins)
* [Word-Add-In-Beispiele auf GitHub](https://github.com/OfficeDev?utf8=%E2%9C%93&q=Word)
