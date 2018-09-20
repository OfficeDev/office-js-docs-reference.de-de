# <a name="functionfile-element"></a><span data-ttu-id="06cd1-101">FunctionFile-Element</span><span class="sxs-lookup"><span data-stu-id="06cd1-101">FunctionFile element</span></span>

<span data-ttu-id="06cd1-p101">Gibt die Quellcodedatei für Vorgänge an, die ein Add-In durch Add-In-Befehle verfügbar macht, die eine JavaScript-Funktion ausführen, anstatt die Benutzeroberfläche anzuzeigen. Das **FunctionFile**-Element ist ein untergeordnetes Element von [DesktopFormFactor](desktopformfactor.md) oder [MobileFormFactor](mobileformfactor.md). Das **resid**-Attribut des **FunctionFile**-Elements wird auf den Wert des **id**-Attributs eines **Url**-Elements im **Resources**-Element festgelegt, das die URL einer HTML-Datei enthält, die alle von Add-In-Befehlsschaltflächen ohne Benutzeroberfläche verwendeten JavaScript-Funktionen entsprechend der Definition im [Control](control.md)-Element enthält oder lädt.</span><span class="sxs-lookup"><span data-stu-id="06cd1-p101">Specifies the source code file for operations that an add-in exposes through add-in commands that execute a JavaScript function instead of displaying UI. The  **FunctionFile** element is a child element of [DesktopFormFactor](desktopformfactor.md) or [MobileFormFactor](mobileformfactor.md). The **resid** attribute of the **FunctionFile** element is set to the value of the **id** attribute of a **Url** element in the **Resources** element that contains the URL to an HTML file that contains or loads all  the JavaScript functions used by UI-less add-in command buttons, as defined by the [Control element](control.md).</span></span>

<span data-ttu-id="06cd1-105">Es folgt ein Beispiel für das **FunctionFile** -Element.</span><span class="sxs-lookup"><span data-stu-id="06cd1-105">The following is an example of the  **FunctionFile** element.</span></span>

```XML
<DesktopFormFactor>
  <FunctionFile resid="residDesktopFuncUrl" />
  <ExtensionPoint xsi:type="PrimaryCommandSurface">
    <!-- information about this extension point -->
  </ExtensionPoint>

  <!-- You can define more than one ExtensionPoint element as needed -->

</DesktopFormFactor>
```

<span data-ttu-id="06cd1-106">Das JavaScript in der HTML-Datei durch das **FunctionFile** -Element angegebenen muss Aufrufen `Office.initialize` und definieren benannte Funktionen, die einen einzelnen Parameter: `event`.</span><span class="sxs-lookup"><span data-stu-id="06cd1-106">The JavaScript in the HTML file indicated by the  **FunctionFile** element must call `Office.initialize` and define named functions that take a single parameter: `event`.</span></span> <span data-ttu-id="06cd1-107">Die Funktionen verwenden, sollten die `item.notificationMessages` -API, um den Fortschritt, Erfolg oder Fehler für den Benutzer anzuzeigen.</span><span class="sxs-lookup"><span data-stu-id="06cd1-107">The functions should use the `item.notificationMessages` API to indicate progress, success, or failure to the user.</span></span> <span data-ttu-id="06cd1-108">Sie sollten außerdem aufrufen `event.completed` Wenn sie die Ausführung beendet wurde.</span><span class="sxs-lookup"><span data-stu-id="06cd1-108">It should also call `event.completed` when it has finished execution.</span></span> <span data-ttu-id="06cd1-109">Der Name der Funktionen sind im Element **Funktionsname** für UI-weniger Schaltflächen verwendet.</span><span class="sxs-lookup"><span data-stu-id="06cd1-109">The name of the functions are used in the **FunctionName** element for UI-less buttons.</span></span>

<span data-ttu-id="06cd1-110">Nachfolgend ein Beispiel für eine HTML-Datei, die eine **trackMessage**-Funktion definiert.</span><span class="sxs-lookup"><span data-stu-id="06cd1-110">The following is an example of an HTML file defining a **trackMessage** function.</span></span>

```js
Office.initialize = function () {
    doAuth();
}

function trackMessage (event) {
    var buttonId = event.source.id;    
    var itemId = Office.context.mailbox.item.id;
    // save this message
    event.completed();
}
```

<span data-ttu-id="06cd1-111">Im folgenden Code wird gezeigt, wie die von  **FunctionName** verwendete Funktion implementiert wird.</span><span class="sxs-lookup"><span data-stu-id="06cd1-111">The following code shows how to implement the function used by  **FunctionName**.</span></span>

```js
// The initialize function must be run each time a new page is loaded.
(function () {
    Office.initialize = function (reason) {
        // If you need to initialize something you can do so here.
    };
})();

// Your function must be in the global namespace.
function writeText(event) {

    // Implement your custom code here. The following code is a simple example.

    Office.context.document.setSelectedDataAsync("ExecuteFunction works. Button ID=" + event.source.id,
        function (asyncResult) {
            var error = asyncResult.error;
            if (asyncResult.status === "failed") {
                // Show error message.
            }
            else {
                // Show success message.
            }
        });
    // Calling event.completed is required. event.completed lets the platform know that processing has completed.
    event.completed();
}
```

> [!IMPORTANT]
> <span data-ttu-id="06cd1-112">Der Aufruf von **event.completed** signalisiert, dass Sie das Ereignis erfolgreich behandelt haben.</span><span class="sxs-lookup"><span data-stu-id="06cd1-112">The call to  **event.completed** signals that you have successfully handled the event.</span></span> <span data-ttu-id="06cd1-113">Wird eine Funktion mehrmals aufgerufen, beispielsweise durch mehrere Klicks auf denselben Add-In-Befehl, werden alle Ereignisse automatisch in die Warteschlange gestellt.</span><span class="sxs-lookup"><span data-stu-id="06cd1-113">When a function is called multiple times, such as multiple clicks on the same add-in command, all events are automatically queued.</span></span> <span data-ttu-id="06cd1-114">Das erste Ereignis wird automatisch ausgeführt, während die anderen Ereignisse in der Warteschlange verbleiben.</span><span class="sxs-lookup"><span data-stu-id="06cd1-114">The first event runs automatically, while the other events remain on the queue.</span></span> <span data-ttu-id="06cd1-115">Wenn Ihre Funktion **event.completed** aufruft, wird der nächste Aufruf für diese Funktion in der Warteschlange ausgeführt.</span><span class="sxs-lookup"><span data-stu-id="06cd1-115">When your function calls **event.completed**, the next queued call to that function runs.</span></span> <span data-ttu-id="06cd1-116">Rufen Sie **event.completed**auf. Andernfalls wird die Funktion nicht ausgeführt.</span><span class="sxs-lookup"><span data-stu-id="06cd1-116">You must call **event.completed**; otherwise your function will not run.</span></span>