# <a name="visio-javascript-api-overview"></a><span data-ttu-id="148a6-101">Visio-JavaScript-API (Übersicht)</span><span class="sxs-lookup"><span data-stu-id="148a6-101">Visio JavaScript API overview</span></span>

<span data-ttu-id="148a6-102">Sie können die JavaScript-APIs für Visio zum Einbetten von Visio-Diagrammen in SharePoint Online verwenden.</span><span class="sxs-lookup"><span data-stu-id="148a6-102">You can use the Visio JavaScript APIs to embed Visio diagrams in SharePoint Online.</span></span> <span data-ttu-id="148a6-103">Ein eingebettetes Visio-Diagramm ist ein Diagramm, das in einer SharePoint-Dokumentbibliothek gespeichert und auf einer SharePoint-Seite angezeigt wird.</span><span class="sxs-lookup"><span data-stu-id="148a6-103">An embedded Visio diagram is a diagram that is stored in a SharePoint document library and displayed on a SharePoint page.</span></span> <span data-ttu-id="148a6-104">Wenn Sie ein Visio-Diagramm einbetten, anzeigen in einer HTML `<iframe>` Element.</span><span class="sxs-lookup"><span data-stu-id="148a6-104">To embed a Visio diagram, display it in an HTML `<iframe>` element.</span></span> <span data-ttu-id="148a6-105">Anschließend können Sie die JavaScript-APIs für Visio verwenden, um programmgesteuert mit dem eingebetteten Diagramm zu arbeiten.</span><span class="sxs-lookup"><span data-stu-id="148a6-105">Then you can use Visio JavaScript APIs to programmatically work with the embedded diagram.</span></span>

![Visio-Diagramm in iframe auf SharePoint-Seite zusammen mit Skript-Editor-Webpart](/javascript/api/docs-ref-conceptual/images/visio-api-block-diagram.png)


<span data-ttu-id="148a6-107">Sie können die Visio-JavaScript-APIs für Folgendes verwenden:</span><span class="sxs-lookup"><span data-stu-id="148a6-107">You can use the Visio JavaScript APIs to:</span></span>

* <span data-ttu-id="148a6-108">Interagieren Sie mit Visio-Diagrammelemente wie Seiten und Shapes.</span><span class="sxs-lookup"><span data-stu-id="148a6-108">Interact with Visio diagram elements like pages and shapes.</span></span>
* <span data-ttu-id="148a6-109">Erstellen von Markup mit visual auf Visio-Diagramm Zeichenbereichs ab.</span><span class="sxs-lookup"><span data-stu-id="148a6-109">Create visual markup on the Visio diagram canvas.</span></span>
* <span data-ttu-id="148a6-110">Schreiben von benutzerdefinierten Handlern für Mausereignisse innerhalb der Zeichnung.</span><span class="sxs-lookup"><span data-stu-id="148a6-110">Write custom handlers for mouse events within the drawing.</span></span>
* <span data-ttu-id="148a6-111">Verfügbarmachen von Diagrammdaten, z. B. Shapetext, Shapedaten und Hyperlinks in Ihrer Lösung</span><span class="sxs-lookup"><span data-stu-id="148a6-111">Expose diagram data, such as shape text, shape data, and hyperlinks, to your solution.</span></span>

<span data-ttu-id="148a6-p102">In diesem Artikel wird beschrieben, wie die JavaScript-APIs für Visio mit Visio Online zum Erstellen von Lösungen für SharePoint Online verwendet werden. Der Artikel enthält eine Einführung in die wichtigsten Konzepte zur Verwendung der APIs, wie z. B. **EmbeddedContext**, **RequestContext** sowie zu JavaScript-Proxyobjekten und den **sync()**-, **Visio.run()**- und **load()**-Methoden. In den Codebeispielen wird gezeigt, wie Sie diese Konzepte anwenden.</span><span class="sxs-lookup"><span data-stu-id="148a6-p102">This article describes how to use the Visio JavaScript APIs with Visio Online to build your solutions for SharePoint Online. It introduces key concepts that are fundamental to using the APIs, such as **EmbeddedSession**, **RequestContext**, and JavaScript proxy objects, and the **sync()**, **Visio.run()**, and **load()** methods. The code examples show you how to apply these concepts.</span></span>

## <a name="embeddedsession"></a><span data-ttu-id="148a6-115">EmbeddedSession</span><span class="sxs-lookup"><span data-stu-id="148a6-115">EmbeddedSession</span></span>

<span data-ttu-id="148a6-116">Das EmbeddedSession-Objekt initialisiert die Kommunikation zwischen dem Entwickler- und dem Visio Online-Frame.</span><span class="sxs-lookup"><span data-stu-id="148a6-116">The EmbeddedSession object initializes communication between the developer frame and the Visio Online frame.</span></span>

```js
var session = new OfficeExtension.EmbeddedSession(url, { id: "embed-iframe",container: document.getElementById("iframeHost") });
session.init().then(function () {
    window.console.log("Session successfully initialized");
});
```

## <a name="visiorunsession-functioncontext--batch-"></a><span data-ttu-id="148a6-117">Visio.Run (Sitzung function(context) {Batch})</span><span class="sxs-lookup"><span data-stu-id="148a6-117">Visio.run(session, function(context) { batch })</span></span>

<span data-ttu-id="148a6-118">**Visio.run()** führt ein Batch-Skript aus, das Aktionen zum Visio-Objektmodell durchführt.</span><span class="sxs-lookup"><span data-stu-id="148a6-118">**Visio.run()** executes a batch script that performs actions on the Visio object model.</span></span> <span data-ttu-id="148a6-119">Die Batchbefehle enthalten Definitionen lokaler JavaScript-Proxyobjekte und **sync()**-Methoden, die den Status zwischen den lokalen und Visio-Objekten und der Zusage-Auflösung synchronisieren.</span><span class="sxs-lookup"><span data-stu-id="148a6-119">The batch commands include definitions of local JavaScript proxy objects and **sync()** methods that synchronize the state between local and Visio objects and promise resolution.</span></span> <span data-ttu-id="148a6-120">Der Vorteil von Batch-Anforderungen in **Visio.run()** ist, dass bei der Zusage-Auflösung aller nachverfolgten Bereichsobjekte, die während der Ausführung zugeordnet wurden, automatisch freigegeben werden.</span><span class="sxs-lookup"><span data-stu-id="148a6-120">The advantage of batching requests in **Visio.run()** is that when the promise is resolved, any tracked page objects that were allocated during the execution will be automatically released.</span></span>

<span data-ttu-id="148a6-121">Die run-Methode ist in der Sitzung und RequestContext-Objekt und gibt eine Zusage zurück (nur in der Regel das Ergebnis der **context.sync()**).</span><span class="sxs-lookup"><span data-stu-id="148a6-121">The run method takes in session and RequestContext object and returns a promise (typically, just the result of **context.sync()**).</span></span> <span data-ttu-id="148a6-122">Es ist möglich, den Batchvorgang außerhalb der **Visio.run()** auszuführen.</span><span class="sxs-lookup"><span data-stu-id="148a6-122">It is possible to run the batch operation outside of the **Visio.run()**.</span></span> <span data-ttu-id="148a6-123">In einem solchen Szenario müssen die Seitenobjektverweise jedoch manuell nachverfolgt und verwaltet werden.</span><span class="sxs-lookup"><span data-stu-id="148a6-123">However, in such a scenario, any page object references needs to be manually tracked and managed.</span></span>

## <a name="requestcontext"></a><span data-ttu-id="148a6-124">RequestContext</span><span class="sxs-lookup"><span data-stu-id="148a6-124">RequestContext</span></span>

<span data-ttu-id="148a6-125">Das Objekt RequestContext vereinfacht die Anforderungen an die Visio-Anwendung.</span><span class="sxs-lookup"><span data-stu-id="148a6-125">The RequestContext object facilitates requests to the Visio application.</span></span> <span data-ttu-id="148a6-126">Da der Rahmen für Entwickler und die Visio-Online-Anwendung in zwei verschiedenen Iframes ausführen, ist das RequestContext-Objekt (Kontext im nächsten Beispiel) erforderlich, um Zugriff auf Visio und damit zusammenhängende Objekte wie Seiten und -Shapes aus der Rahmen für Entwickler zu erhalten.</span><span class="sxs-lookup"><span data-stu-id="148a6-126">Because the developer frame and the Visio Online application run in two different iframes, the RequestContext object (context in next example) is required to get access to Visio and related objects such as pages and shapes, from the developer frame.</span></span>

```js
function hideToolbars() {
    Visio.run(session, function(context){
        var app = context.document.application;
        app.showToolbars = false;
        return context.sync().then(function () {
            window.console.log("Toolbars Hidden");
        });
    }).catch(function(error)
    {
        window.console.log("Error: " + error);
    });
};
```

## <a name="proxy-objects"></a><span data-ttu-id="148a6-127">Proxyobjekte</span><span class="sxs-lookup"><span data-stu-id="148a6-127">Proxy objects</span></span>

<span data-ttu-id="148a6-p106">Die in einem Add-In deklarierten und verwendeten Visio-JavaScript-Objekte sind Proxyobjekte für die realen Objekte in einem Visio-Dokument. Alle Aktionen zu Proxyobjekten werden in Visio nicht realisiert, und der Status des Visio-Dokuments wird in den Proxyobjekten erst umgesetzt, wenn der Status des Dokuments synchronisiert wurde. Der Dokumentstatus wird synchronisiert, wenn `context.sync()` ausgeführt wird.</span><span class="sxs-lookup"><span data-stu-id="148a6-p106">The Visio JavaScript objects declared and used in an add-in are proxy objects for the real objects in a Visio document. All actions taken on proxy objects are not realized in Visio, and the state of the Visio document is not realized in the proxy objects until the document state has been synchronized. The document state is synchronized when `context.sync()` is run.</span></span>

<span data-ttu-id="148a6-131">Verweisen auf die ausgewählte Seite ist beispielsweise lokale JavaScript-Objekt GetActivePage deklariert.</span><span class="sxs-lookup"><span data-stu-id="148a6-131">For example, the local JavaScript object getActivePage is declared to reference the selected page.</span></span> <span data-ttu-id="148a6-132">Dies kann die Einstellung der Eigenschaften und Methoden aufrufen in die Warteschlange verwendet werden.</span><span class="sxs-lookup"><span data-stu-id="148a6-132">This can be used to queue the setting of its properties and invoking methods.</span></span> <span data-ttu-id="148a6-133">Die Aktionen für solche Objekte werden nicht realisiert, bis die **sync()** -Methode ausgeführt wird.</span><span class="sxs-lookup"><span data-stu-id="148a6-133">The actions on such objects are not realized until the **sync()** method is run.</span></span>

```js
var activePage = context.document.getActivePage();
```

## <a name="sync"></a><span data-ttu-id="148a6-134">sync()</span><span class="sxs-lookup"><span data-stu-id="148a6-134">sync()</span></span>

<span data-ttu-id="148a6-135">Die **sync()** -Methode den Status zwischen JavaScript Proxyobjekte synchronisiert und reale Objekte in Visio durch Ausführen von Anweisungen in der Warteschlange auf dem Kontext und Abrufen von Eigenschaften des Office-Objekte für die Verwendung in Ihrem Code geladen.</span><span class="sxs-lookup"><span data-stu-id="148a6-135">The **sync()** method synchronizes the state between JavaScript proxy objects and real objects in Visio by executing instructions queued on the context and retrieving properties of loaded Office objects for use in your code.</span></span> <span data-ttu-id="148a6-136">Diese Methode gibt eine Zusage zurück, die nach Abschluss der Synchronisierung aufgelöst wird.</span><span class="sxs-lookup"><span data-stu-id="148a6-136">This method returns a promise, which is resolved when synchronization is complete.</span></span> 

## <a name="load"></a><span data-ttu-id="148a6-137">load()</span><span class="sxs-lookup"><span data-stu-id="148a6-137">load()</span></span>

<span data-ttu-id="148a6-p109">Die **load()**-Methode dient zum Auffüllen der Proxyobjekte, die auf der JavaScript-Ebene im Add-In erstellt werden. Beim Abrufen eines Objekts, z. B. eines Dokuments, wird ein lokales Proxyobjekt zunächst auf der JavaScript-Ebene erstellt. Damit kann die Einstellung der Eigenschaften und das Abrufen von Methoden in der Warteschlange verwendet werden. Zum Lesen von Objekteigenschaften oder Beziehungen müssen die **load()**- und **sync()**-Methoden zuerst aufgerufen werden. Die load()-Methode nimmt die Eigenschaften und Beziehungen auf, die beim Aufrufen der **sync()**-Methode geladen werden müssen.</span><span class="sxs-lookup"><span data-stu-id="148a6-p109">The **load()** method is used to fill in the proxy objects created in the add-in JavaScript layer. When trying to retrieve an object such as a document, a local proxy object is created first in the JavaScript layer. Such an object can be used to queue the setting of its properties and invoking methods. However, for reading object properties or relations, the **load()** and **sync()** methods need to be invoked first. The load() method takes in the properties and relations that need to be loaded when the **sync()** method is called.</span></span>

<span data-ttu-id="148a6-143">Nachfolgend ist die Syntax für die **load()**-Methode veranschaulicht.</span><span class="sxs-lookup"><span data-stu-id="148a6-143">The following shows the syntax for the **load()** method.</span></span>

```js
object.load(string: properties); //or object.load(array: properties); //or object.load({loadOption});
```

1. <span data-ttu-id="148a6-144">**Eigenschaften** ist die Liste der Eigenschaftennamen geladen, als angegebenen durch Trennzeichen getrennte Zeichenfolgen sein oder Array von Namen.</span><span class="sxs-lookup"><span data-stu-id="148a6-144">**properties** is the list of property names to be loaded, specified as comma-delimited strings or array of names.</span></span> <span data-ttu-id="148a6-145">Weitere Informationen dazu finden Sie in den **.load()**-Methoden unter den einzelnen Objekten.</span><span class="sxs-lookup"><span data-stu-id="148a6-145">See **.load()** methods under each object for details.</span></span>

2. <span data-ttu-id="148a6-p111">**loadOption** gibt ein Objekt an, das die Optionen für Auswahl, Erweiterung, oben und Überspringen beschreibt. Weitere Informationen finden Sie im Objekt [Ladeoptionen](/javascript/api/office/officeextension.loadoption).</span><span class="sxs-lookup"><span data-stu-id="148a6-p111">**loadOption** specifies an object that describes the selection, expansion, top, and skip options. See object load [options](/javascript/api/office/officeextension.loadoption) for details.</span></span>

## <a name="example-printing-all-shapes-text-in-active-page"></a><span data-ttu-id="148a6-148">Beispiel: Drucken des gesamten Shapetexts in der aktiven Seite</span><span class="sxs-lookup"><span data-stu-id="148a6-148">Example: Printing all shapes text in active page</span></span>

<span data-ttu-id="148a6-149">Das folgende Beispiel zeigt, wie der Shapetextwert aus einem Arrayshapeobjekt gedruckt wird.</span><span class="sxs-lookup"><span data-stu-id="148a6-149">The following example shows you how to print shape text value from an array shapes object.</span></span>
<span data-ttu-id="148a6-150">Die **Visio.run()**-Methode enthält eine Reihe von Anweisungen.</span><span class="sxs-lookup"><span data-stu-id="148a6-150">The **Visio.run()** method contains a batch of instructions.</span></span> <span data-ttu-id="148a6-151">Als Teil dieses Batches wird ein Proxyobjekt erstellt, das auf Shapes im aktiven Dokument verweist.</span><span class="sxs-lookup"><span data-stu-id="148a6-151">As part of this batch, a proxy object is created that references shapes on the active document.</span></span>

<span data-ttu-id="148a6-152">Alle diese Befehle werden in die Warteschlange gestellt und ausgeführt werden, wenn **context.sync()** aufgerufen wird.</span><span class="sxs-lookup"><span data-stu-id="148a6-152">All these commands are queued and run when **context.sync()** is called.</span></span> <span data-ttu-id="148a6-153">Die**sync()**-Methode gibt eine Zusage zurück, mit der sie mit anderen Vorgängen verkettet werden kann.</span><span class="sxs-lookup"><span data-stu-id="148a6-153">The **sync()** method returns a promise that can be used to chain it with other operations.</span></span>

```js
Visio.run(session, function (context) {
    var page = context.document.getActivePage();
    var shapes = page.shapes;
    shapes.load();
    return context.sync().then(function () {
        for(var i=0; i<shapes.items.length;i++) {
            var shape = shapes.items[i];
            window.console.log("Shape Text: " + shape.text );
        }
    });
}).catch(function(error) {
    window.console.log("Error: " + error);
    if (error instanceof OfficeExtension.Error) {
        window.console.log ("Debug info: " + JSON.stringify(error.debugInfo));
    }
});
```

## <a name="error-messages"></a><span data-ttu-id="148a6-154">Fehlermeldungen</span><span class="sxs-lookup"><span data-stu-id="148a6-154">Error messages</span></span>

<span data-ttu-id="148a6-p114">Fehler werden mithilfe eines Fehlerobjekts zurückgegeben, das aus einem Code und einer Nachricht besteht. Die folgende Tabelle enthält eine Liste möglicher Fehlerzustände, die auftreten können.</span><span class="sxs-lookup"><span data-stu-id="148a6-p114">Errors are returned using an error object that consists of a code and a message. The following table provides a list of possible error conditions that can occur.</span></span>

| <span data-ttu-id="148a6-157">error.code</span><span class="sxs-lookup"><span data-stu-id="148a6-157">error.code</span></span>            | <span data-ttu-id="148a6-158">error.message</span><span class="sxs-lookup"><span data-stu-id="148a6-158">error.message</span></span> |
|-----------------------|----------------------------------------------------------------|
| <span data-ttu-id="148a6-159">InvalidArgument</span><span class="sxs-lookup"><span data-stu-id="148a6-159">InvalidArgument</span></span>       | <span data-ttu-id="148a6-160">Das Argument ist ungültig oder fehlt oder weist ein falsches Format auf.</span><span class="sxs-lookup"><span data-stu-id="148a6-160">The argument is invalid or missing or has an incorrect format.</span></span> |
| <span data-ttu-id="148a6-161">GeneralException</span><span class="sxs-lookup"><span data-stu-id="148a6-161">GeneralException</span></span>      | <span data-ttu-id="148a6-162">Beim Verarbeiten der Anforderung ist ein interner Fehler aufgetreten.</span><span class="sxs-lookup"><span data-stu-id="148a6-162">There was an internal error while processing the request.</span></span> |
| <span data-ttu-id="148a6-163">NotImplemented</span><span class="sxs-lookup"><span data-stu-id="148a6-163">NotImplemented</span></span>        | <span data-ttu-id="148a6-164">Das angeforderte Feature ist nicht implementiert.</span><span class="sxs-lookup"><span data-stu-id="148a6-164">The requested feature isn't implemented.</span></span>  |
| <span data-ttu-id="148a6-165">UnsupportedOperation</span><span class="sxs-lookup"><span data-stu-id="148a6-165">UnsupportedOperation</span></span>  | <span data-ttu-id="148a6-166">Dieser Vorgang wird nicht unterstützt.</span><span class="sxs-lookup"><span data-stu-id="148a6-166">The operation being attempted is not supported.</span></span> |
| <span data-ttu-id="148a6-167">AccessDenied</span><span class="sxs-lookup"><span data-stu-id="148a6-167">AccessDenied</span></span>          | <span data-ttu-id="148a6-168">Sie können den angeforderten Vorgang nicht durchzuführen.</span><span class="sxs-lookup"><span data-stu-id="148a6-168">You cannot perform the requested operation.</span></span> |
| <span data-ttu-id="148a6-169">ItemNotFound</span><span class="sxs-lookup"><span data-stu-id="148a6-169">ItemNotFound</span></span>          | <span data-ttu-id="148a6-170">Die angeforderte Ressource ist nicht vorhanden.</span><span class="sxs-lookup"><span data-stu-id="148a6-170">The requested resource doesn't exist.</span></span> |

## <a name="get-started"></a><span data-ttu-id="148a6-171">Erste Schritte</span><span class="sxs-lookup"><span data-stu-id="148a6-171">Get started</span></span>

<span data-ttu-id="148a6-172">Das Beispiel können in diesem Abschnitt Sie die ersten Schritte beim.</span><span class="sxs-lookup"><span data-stu-id="148a6-172">You can use the example in this section to get started.</span></span> <span data-ttu-id="148a6-173">Dieses Beispiel zeigt, wie den Shape-Text des ausgewählten Shapes in einem Visio-Diagramm programmgesteuert angezeigt.</span><span class="sxs-lookup"><span data-stu-id="148a6-173">This example shows you how to programmatically display the shape text of the selected shape in a Visio diagram.</span></span> <span data-ttu-id="148a6-174">Zunächst erstellen Sie eine klassische Seite in SharePoint Online, oder bearbeiten Sie eine vorhandene Seite.</span><span class="sxs-lookup"><span data-stu-id="148a6-174">To begin, create a classic page in SharePoint Online or edit an existing page.</span></span> <span data-ttu-id="148a6-175">Fügen Sie ein Skript-Editor-Webpart auf der Seite und den folgenden Code kopieren und einfügen.</span><span class="sxs-lookup"><span data-stu-id="148a6-175">Add a script editor webpart on the page and copy-paste the following code.</span></span>

```js
<script src='https://appsforoffice.microsoft.com/embedded/1.0/visio-web-embedded.js' type='text/javascript'></script>

Enter Visio File Url:<br/>
<script language="javascript">
document.write("<input type='text' id='fileUrl' size='120'/>");
document.write("<input type='button' value='InitEmbeddedFrame' onclick='initEmbeddedFrame()' />");
document.write("<br />");
document.write("<input type='button' value='SelectedShapeText' onclick='getSelectedShapeText()' />");
document.write("<textarea id='ResultOutput' style='width:350px;height:60px'> </textarea>");
document.write("<div id='iframeHost' />");

let session; // Global variable to store the session and pass it afterwards in Visio.run()
var textArea;
// Loads the Visio application and Initializes communication between developer frame and Visio online frame
function initEmbeddedFrame() {
    textArea = document.getElementById('ResultOutput');
    var url = document.getElementById('fileUrl').value;
    if (!url) {
        window.alert("File URL should not be empty");
    }
    // APIs are enabled for EmbedView action only.
    url = url.replace("action=view","action=embedview");
    url = url.replace("action=interactivepreview","action=embedview");
    url = url.replace("action=default","action=embedview");
    url = url.replace("action=edit","action=embedview");
  
    session = new OfficeExtension.EmbeddedSession(url, { id: "embed-iframe",container: document.getElementById("iframeHost") });
    return session.init().then(function () {
        // Initialization is successful
        textArea.value  = "Initialization is successful";
    });
}

// Code for getting selected Shape Text using the shapes collection object
function getSelectedShapeText() {
    Visio.run(session, function (context) {
        var page = context.document.getActivePage();
        var shapes = page.shapes;
        shapes.load();
        return context.sync().then(function () {
            textArea.value = "Please select a Shape in the Diagram";
            for(var i=0; i<shapes.items.length;i++) {
                var shape = shapes.items[i];
                if ( shape.select == true) {
                    textArea.value = shape.text;
                    return;
                }
            }
        });
    }).catch(function(error) {
        textArea.value = "Error: ";
        if (error instanceof OfficeExtension.Error) {
            textArea.value += "Debug info: " + JSON.stringify(error.debugInfo);
        }
    });
}
</script>
```

<span data-ttu-id="148a6-176">Nach Ausführung dieses brauchen Sie ist die URL eines Visio-Diagramms, dem Sie arbeiten möchten.</span><span class="sxs-lookup"><span data-stu-id="148a6-176">After that, all you need is the URL of a Visio diagram that you want to work with.</span></span> <span data-ttu-id="148a6-177">Gerade Laden Sie das Visio-Diagramm in SharePoint Online, und öffnen Sie es in Visio Online.</span><span class="sxs-lookup"><span data-stu-id="148a6-177">Just upload the Visio diagram to SharePoint Online and open it in Visio Online.</span></span> <span data-ttu-id="148a6-178">Öffnen Sie von dort das Dialogfeld "einbetten", und verwenden Sie die URL einbetten im obigen Beispiel.</span><span class="sxs-lookup"><span data-stu-id="148a6-178">From there, open the Embed dialog and use the Embed URL in the above example.</span></span>

![Kopieren Sie die URL der Visio-Datei aus Embed-Dialogfeld](/javascript/api/docs-ref-conceptual/images/Visio-embed-url.png)

<span data-ttu-id="148a6-180">Wenn Sie Online Visio im Bearbeitungsmodus verwenden, öffnen Sie im Dialogfeld einbetten durch Auswählen von **Datei** > **Freigeben** > **einbetten**.</span><span class="sxs-lookup"><span data-stu-id="148a6-180">If you are using Visio Online in Edit mode, open the Embed dialog by choosing **File** > **Share** > **Embed**.</span></span> <span data-ttu-id="148a6-181">Wenn Sie Online Visio im Ansichtsmodus verwenden, öffnen Sie im Dialogfeld einbetten durch Auswählen von '...' und klicken Sie dann auf **einbetten**.</span><span class="sxs-lookup"><span data-stu-id="148a6-181">If you are using Visio Online in View mode, open the Embed dialog by choosing '...' and then **Embed**.</span></span>

## <a name="open-api-specifications"></a><span data-ttu-id="148a6-182">Offene API-Spezifikationen</span><span class="sxs-lookup"><span data-stu-id="148a6-182">Open API specifications</span></span>

<span data-ttu-id="148a6-p118">Während des Entwerfens und Entwickelns neuer APIs stellen wir diese zur Verfügung, damit Sie auf der Seite [Offene API-Spezifikationen](../openspec.md) Ihr Feedback abgeben können. Erfahren Sie, welche neuen Funktionen geplant sind, und teilen Sie uns Ihre Meinung zu unseren Designspezifikationen mit.</span><span class="sxs-lookup"><span data-stu-id="148a6-p118">As we design and develop new APIs, we'll make them available for your feedback on our [Open API specifications](../openspec.md) page. Find out what new features are in the pipeline, and provide your input on our design specifications.</span></span>

## <a name="visio-javascript-api-reference"></a><span data-ttu-id="148a6-185">Visio-JavaScript-API-Referenz</span><span class="sxs-lookup"><span data-stu-id="148a6-185">Visio JavaScript API reference</span></span>

<span data-ttu-id="148a6-186">Ausführliche Informationen zu Visio-JavaScript-API finden Sie in der [Visio-JavaScript-API-Referenzdokumentation](/javascript/api/visio).</span><span class="sxs-lookup"><span data-stu-id="148a6-186">For detailed information about Visio JavaScript API, see the [Visio JavaScript API reference documentation](/javascript/api/visio).</span></span>
