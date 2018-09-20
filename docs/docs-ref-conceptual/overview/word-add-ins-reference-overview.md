# <a name="word-javascript-api-overview"></a><span data-ttu-id="db6bd-101">Word-JavaScript-API (Übersicht)</span><span class="sxs-lookup"><span data-stu-id="db6bd-101">Word JavaScript API overview</span></span>

<span data-ttu-id="db6bd-p101">Word bietet eine umfangreiche Sammlung von APIs, die Sie zum Erstellen von Add-Ins verwenden können, die mit Dokumentinhalten und Metadaten interagieren. Verwenden Sie diese APIs, um überzeugende Oberflächen zu erstellen, die mit Word interagieren und Word erweitern. Sie können Inhalte importieren und exportieren, neue Dokumente aus unterschiedlichen Datenquellen zusammenstellen und mit Dokumentworkflows interagieren, um benutzerdefinierte Dokumentlösungen zu erstellen.</span><span class="sxs-lookup"><span data-stu-id="db6bd-p101">Word provides a rich set of APIs that you can use to create add-ins that interact with document content and metadata. Use these APIs to create compelling experiences that integrate with and extend Word. You can import and export content, assemble new documents from different data sources, and integrate with document workflows to create custom document solutions.</span></span>

<span data-ttu-id="db6bd-105">Sie können zwei JavaScript-APIs verwenden, um mit den Objekten und Metadaten in einem Word-Dokument zu interagieren:</span><span class="sxs-lookup"><span data-stu-id="db6bd-105">You can use two JavaScript APIs to interact with the objects and metadata in a Word document:</span></span>

- <span data-ttu-id="db6bd-106">JavaScript-API für Word - In Office 2016 eingeführt.</span><span class="sxs-lookup"><span data-stu-id="db6bd-106">Word JavaScript API - Introduced in Office 2016.</span></span>
- <span data-ttu-id="db6bd-107">[JavaScript-API für Office](../javascript-api-for-office.md) (Office.js) - In Office 2013 eingeführt.</span><span class="sxs-lookup"><span data-stu-id="db6bd-107">[JavaScript API for Office](../javascript-api-for-office.md) (Office.js) - Introduced in Office 2013.</span></span>

## <a name="word-javascript-api"></a><span data-ttu-id="db6bd-108">JavaScript-API für Word</span><span class="sxs-lookup"><span data-stu-id="db6bd-108">Word JavaScript API</span></span>

<span data-ttu-id="db6bd-p102">Die JavaScript-API für Word wird von Office.js geladen. Sie ändert die Art der Interaktion mit Objekten wie Dokumenten und Absätzen. Statt einzelne asynchrone APIs zum Abrufen und Aktualisieren der einzelnen Objekte bereitzustellen, bietet die JavaScript-API für Word „Proxy“-JavaScript-Objekte, die den realen Objekten entsprechen, die in Word ausgeführt werden. Sie können mit diesen Proxyobjekten interagieren, indem Sie ihre Eigenschaften synchron lesen und synchrone Methoden aufrufen, um dazu Vorgänge durchzuführen. Diese Interaktionen mit Proxyobjekten werden nicht sofort im ausgeführten Skript umgesetzt, deshalb stellen wir eine Methode zum Kontext mit dem Namen sync() bereit. Die **context.sync**-Methode synchronisiert den Status zwischen ausgeführten JavaScript-Objekten und den realen Objekten in Office, indem in die Warteschlange eingereihte Anweisungen ausgeführt und Eigenschaften von geladenen Word-Objekten für die Verwendung in Ihrem Skript abgerufen werden.</span><span class="sxs-lookup"><span data-stu-id="db6bd-p102">The Word JavaScript API is loaded by Office.js. The Word JavaScript API changes the way that you can interact with objects like documents and paragraphs. Rather than providing individual asynchronous APIs for retrieving and updating each of these objects, the Word JavaScript API provides “proxy” JavaScript objects that correspond to the real objects running in Word. You can interact with these proxy objects by synchronously reading and writing their properties and calling synchronous methods to perform operations on them. These interactions with proxy objects aren’t immediately realized in the running script. The **context.sync** method synchronizes the state between your running JavaScript and the real objects in Office by executing queued instructions and retrieving properties of loaded Word objects for use in your script.</span></span>

## <a name="javascript-api-for-office"></a><span data-ttu-id="db6bd-115">JavaScript-API für Office</span><span class="sxs-lookup"><span data-stu-id="db6bd-115">JavaScript API for Office</span></span>

<span data-ttu-id="db6bd-116">Sie können von den folgenden Speicherorten auf Office.js verweisen:</span><span class="sxs-lookup"><span data-stu-id="db6bd-116">You can reference Office.js from the following locations:</span></span>

* <span data-ttu-id="db6bd-117">https://appsforoffice.microsoft.com/lib/1/hosted/office.js-Verwenden Sie diese Ressource für Produktion-add-ins.</span><span class="sxs-lookup"><span data-stu-id="db6bd-117">https://appsforoffice.microsoft.com/lib/1/hosted/office.js - use this resource for production add-ins.</span></span>
* <span data-ttu-id="db6bd-118">https://appsforoffice.microsoft.com/lib/beta/hosted/office.js-Verwenden Sie diese Ressource aus, wenn Sie die Preview-Features möchten.</span><span class="sxs-lookup"><span data-stu-id="db6bd-118">https://appsforoffice.microsoft.com/lib/beta/hosted/office.js - use this resource when you're trying out preview features.</span></span>

<span data-ttu-id="db6bd-p103">Wenn Sie [Visual Studio](https://www.visualstudio.com/products/free-developer-offers-vs) verwenden, können Sie die [Office-Entwicklertools](https://www.visualstudio.com/features/office-tools-vs.aspx) herunterladen, um Projektvorlagen zu erhalten, die Office.js enthalten.  Sie können auch [nuget verwenden, um Office.js zu erhalten](https://www.nuget.org/packages/Microsoft.Office.js/).</span><span class="sxs-lookup"><span data-stu-id="db6bd-p103">If you're using [Visual Studio](https://www.visualstudio.com/products/free-developer-offers-vs), you can download the [Office Developer Tools](https://www.visualstudio.com/features/office-tools-vs.aspx) to get project templates that include Office.js.  You can also use [nuget to get Office.js](https://www.nuget.org/packages/Microsoft.Office.js/).</span></span>

<span data-ttu-id="db6bd-121">Wenn Sie TypeScript verwenden und über npm verfügen, können Sie die TypeScript-Definitionen durch Eingeben des folgenden Texts in der Befehlszeilenoberfläche abrufen: `typings install office-js --ambient`.</span><span class="sxs-lookup"><span data-stu-id="db6bd-121">If you use TypeScript and have npm, you can get the the TypeScript definitions by typing this in your command line interface: `typings install office-js --ambient`.</span></span>

## <a name="running-word-add-ins"></a><span data-ttu-id="db6bd-122">Ausführen von Word-Add-Ins</span><span class="sxs-lookup"><span data-stu-id="db6bd-122">Running Word add-ins</span></span>

<span data-ttu-id="db6bd-p104">Verwenden Sie einen Office.initialize-Ereignishandler, um Ihr Add-In auszuführen. Unter [Grundlegendes zur API](https://docs.microsoft.com/office/dev/add-ins/develop/understanding-the-javascript-api-for-office) finden Sie weitere Informationen zur Add-In-Initialisierung.</span><span class="sxs-lookup"><span data-stu-id="db6bd-p104">To run your add-in, use an Office.initialize event handler. For more information about add-in initialization, see [Understanding the API](https://docs.microsoft.com/office/dev/add-ins/develop/understanding-the-javascript-api-for-office) .</span></span>

<span data-ttu-id="db6bd-p105">Add-Ins für Word 2016 werden durch Übergeben einer Funktion an die **Word.run()**-Methode ausgeführt . Die Funktion in der **run**-Methode muss ein Kontextargument besitzen. Dieses [Kontextobjekt](/javascript/api/word/word.requestcontext) unterscheidet sich dem Kontextobjekt, das Sie aus dem Office-Objekt abrufen, obwohl es ebenfalls für die Interaktion mit der Word-Laufzeitumgebung verwendet wird. Das folgenden Beispiel zeigt, wie ein Word-Add-In mithilfe der **Word.run()**-Methode initialisiert und ausgeführt wird.</span><span class="sxs-lookup"><span data-stu-id="db6bd-p105">Add-ins that target Word 2016 execute by passing a function into the **Word.run()** method. The function passed into the **run** method must have a context argument. This [context object](/javascript/api/word/word.requestcontext) is different than the context object you get from the Office object, but it is also used to interact with the Word runtime environment. The context object provides access to the Word JavaScript API object model. The following example shows how to initialize and execute a Word add-in by using the **Word.run()** method.</span></span>

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

### <a name="synchronizing-word-documents-with-word-javascript-api-proxy-objects"></a><span data-ttu-id="db6bd-130">Synchronisieren von Word-Dokumenten mit Proxyobjekten der Word-JavaScript-API</span><span class="sxs-lookup"><span data-stu-id="db6bd-130">Synchronizing Word documents with Word JavaScript API proxy objects</span></span>

<span data-ttu-id="db6bd-p106">Das Objektmodell der Word-JavaScript-API ist lose mit den Objekten in Word verknüpft. Die Objekte der Word--JavaScript-API sind Proxys für Objekte in einem Word-Dokument. Aktionen, die für Proxyobjekte ausgeführt werden, werden in Word erst umgesetzt, wenn der Dokumentstatus synchronisiert wurde. Der Status des Word-Dokuments wird in den Proxyobjekten jedoch erst umgesetzt, wenn der Status des Dokuments synchronisiert wurde. Um den Dokumentstatus zu synchronisieren, führen Sie die **context.sync()**-Methode aus. Das folgende Beispiel veranschaulicht die Erstellung eines Proxytextobjekts und einen Befehl in der Warteschlange, um die Texteigenschaft zu dem Proxytextobjekt zu laden, und verwendet die **context.sync()**-Methode, um den Text im Word-Dokument mit dem Textproxyobjekt zu synchronisieren.</span><span class="sxs-lookup"><span data-stu-id="db6bd-p106">The Word JavaScript API object model is loosely coupled with the objects in Word. Word JavaScript API objects are proxies for objects in a Word document. Actions taken on proxy objects are not realized in Word until the document state has been synchronized. Conversely, the state of the Word document is not realized in the proxy objects until the document state has been synchronized. To synchronize the document state, you run the **context.sync()** method. The following example creates a proxy body object and a queued command to load the text property on the proxy body object, and uses the **context.sync()** method to synchronize the body of the Word document with the body proxy object.</span></span>

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

### <a name="executing-a-batch-of-commands"></a><span data-ttu-id="db6bd-137">Ausführen einer Reihe von Befehlen</span><span class="sxs-lookup"><span data-stu-id="db6bd-137">Executing a batch of commands</span></span>

<span data-ttu-id="db6bd-p107">Die Word-Proxyobjekte besitzen Methoden für den Zugriff und die Aktualisierung des Objektmodells. Diese Methoden werden sequenziell in der Reihenfolge ausgeführt, in der sie alle im Batch in die Warteschlange gestellt wurden. Alle Befehle in der Warteschlange im Batch werden ausgeführt, wenn context.sync() aufgerufen wird.</span><span class="sxs-lookup"><span data-stu-id="db6bd-p107">The Word proxy objects have methods for accessing and updating the object model. These methods are executed sequentially in the order in which they were queued in the batch. All of the commands that are queued in the batch are executed when context.sync() is called.</span></span>

<span data-ttu-id="db6bd-p108">Im folgenden Beispiel wird gezeigt, wie die Warteschlange von Befehlen funktioniert. Wenn **context.sync()** aufgerufen wird, wird der Befehl zum Laden des Textkörpers in Word ausgeführt. Dann wird der Befehl zum Einfügen von Text in Word ausgeführt. Die Ergebnisse werden im Textproxyobjekt zurückgegeben. Der Wert der **body.text**-Eigenschaft in der Word-JavaScript-API ist der Wert des Word-Dokumenttexts <u>bevor</u> der Text in das Word-Dokument eingefügt wurde.</span><span class="sxs-lookup"><span data-stu-id="db6bd-p108">The following example shows how the command queue works. When **context.sync()** is called, the command to load the body text is executed in Word. Then, the command to insert text into the body in Word occurs. The results are then returned to the body proxy object. The value of the **body.text** property in the Word JavaScript API is the value of the Word document body <u>before</u> the text was inserted into Word document.</span></span>


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

## <a name="word-javascript-api-open-specifications"></a><span data-ttu-id="db6bd-146">Offene Spezifikationen Word JavaScript-API</span><span class="sxs-lookup"><span data-stu-id="db6bd-146">Word JavaScript API open specifications</span></span>

<span data-ttu-id="db6bd-p109">Während des Entwerfens und Entwickelns neuer APIs für Word-Add-Ins stellen wir diese zur Verfügung, damit Sie auf der Seite [Offene API-Spezifikationen](../openspec.md) Ihr Feedback abgeben können. Erfahren Sie, welche neuen Funktionen für die Word-JavaScript-APIs geplant sind, und teilen Sie uns Ihre Meinung zu unseren Designspezifikationen mit.</span><span class="sxs-lookup"><span data-stu-id="db6bd-p109">As we design and develop new APIs for Word add-ins, we'll make them available for your feedback on our [Open API specifications](../openspec.md) page. Find out what new features are in the pipeline for the Word JavaScript APIs, and provide your input on our design specifications.</span></span>

## <a name="word-javascript-api-reference"></a><span data-ttu-id="db6bd-149">JavaScript-API-Referenz für Word</span><span class="sxs-lookup"><span data-stu-id="db6bd-149">Word JavaScript API reference</span></span>

<span data-ttu-id="db6bd-150">Ausführliche Informationen zu der JavaScript-API für Word finden Sie in der [Referenzdokumentation für die Word-JavaScript-API](/javascript/api/word).</span><span class="sxs-lookup"><span data-stu-id="db6bd-150">For detailed information about the Word JavaScript API, see the [Word JavaScript API reference documentation](/javascript/api/word).</span></span>

## <a name="see-also"></a><span data-ttu-id="db6bd-151">Siehe auch</span><span class="sxs-lookup"><span data-stu-id="db6bd-151">See also</span></span>

* [<span data-ttu-id="db6bd-152">?bersicht ?ber Word-Add-Ins</span><span class="sxs-lookup"><span data-stu-id="db6bd-152">Word add-ins overview</span></span>](https://docs.microsoft.com/office/dev/add-ins/word/word-add-ins-programming-overview)
* [<span data-ttu-id="db6bd-153">Office-Add-Ins-Plattformübersicht</span><span class="sxs-lookup"><span data-stu-id="db6bd-153">Office Add-ins platform overview</span></span>](https://docs.microsoft.com/office/dev/add-ins/overview/office-add-ins)
* [<span data-ttu-id="db6bd-154">Word-Add-In-Beispiele auf GitHub</span><span class="sxs-lookup"><span data-stu-id="db6bd-154">Word add-in samples on GitHub</span></span>](https://github.com/OfficeDev?utf8=%E2%9C%93&q=Word)
