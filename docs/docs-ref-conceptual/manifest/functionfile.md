# <a name="functionfile-element"></a>FunctionFile-Element

Gibt die Quellcodedatei für Vorgänge an, die ein Add-In durch Add-In-Befehle verfügbar macht, die eine JavaScript-Funktion ausführen, anstatt die Benutzeroberfläche anzuzeigen. Das **FunctionFile**-Element ist ein untergeordnetes Element von [DesktopFormFactor](desktopformfactor.md) oder [MobileFormFactor](mobileformfactor.md). Das **resid**-Attribut des **FunctionFile**-Elements wird auf den Wert des **id**-Attributs eines **Url**-Elements im **Resources**-Element festgelegt, das die URL einer HTML-Datei enthält, die alle von Add-In-Befehlsschaltflächen ohne Benutzeroberfläche verwendeten JavaScript-Funktionen entsprechend der Definition im [Control](control.md)-Element enthält oder lädt.

Es folgt ein Beispiel für das **FunctionFile** -Element.

```XML
<DesktopFormFactor>
  <FunctionFile resid="residDesktopFuncUrl" />
  <ExtensionPoint xsi:type="PrimaryCommandSurface">
    <!-- information about this extension point -->
  </ExtensionPoint>

  <!-- You can define more than one ExtensionPoint element as needed -->

</DesktopFormFactor>
```

Das JavaScript in der HTML-Datei durch das **FunctionFile** -Element angegebenen muss Aufrufen `Office.initialize` und definieren benannte Funktionen, die einen einzelnen Parameter: `event`. Die Funktionen verwenden, sollten die `item.notificationMessages` -API, um den Fortschritt, Erfolg oder Fehler für den Benutzer anzuzeigen. Sie sollten außerdem aufrufen `event.completed` Wenn sie die Ausführung beendet wurde. Der Name der Funktionen sind im Element **Funktionsname** für UI-weniger Schaltflächen verwendet.

Nachfolgend ein Beispiel für eine HTML-Datei, die eine **trackMessage**-Funktion definiert.

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

Im folgenden Code wird gezeigt, wie die von  **FunctionName** verwendete Funktion implementiert wird.

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
> Der Aufruf von **event.completed** signalisiert, dass Sie das Ereignis erfolgreich behandelt haben. Wird eine Funktion mehrmals aufgerufen, beispielsweise durch mehrere Klicks auf denselben Add-In-Befehl, werden alle Ereignisse automatisch in die Warteschlange gestellt. Das erste Ereignis wird automatisch ausgeführt, während die anderen Ereignisse in der Warteschlange verbleiben. Wenn Ihre Funktion **event.completed** aufruft, wird der nächste Aufruf für diese Funktion in der Warteschlange ausgeführt. Rufen Sie **event.completed**auf. Andernfalls wird die Funktion nicht ausgeführt.