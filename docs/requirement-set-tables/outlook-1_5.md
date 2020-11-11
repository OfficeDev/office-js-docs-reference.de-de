| Klasse | Felder | Beschreibung |
|:---|:---|:---|
|[Postfach](/javascript/api/outlook/outlook.mailbox)|[addHandlerAsync (EventType: Office. EventType \| String, Handler: (Type: Office. EventType) => void, Options?: Office. AsyncContextOptions, Callback?: (asyncResult: Office. AsyncResult <void> ) => void)](/javascript/api/outlook/outlook.mailbox#addhandlerasync-eventtype--handler--type-)|Fügt einen Ereignishandler für ein unterstütztes Ereignis hinzu.|
||[getCallbackTokenAsync (Options: Office. AsyncContextOptions & {isrest?: Boolean}, Callback: (asyncResult: Office. AsyncResult <string> ) => void)](/javascript/api/outlook/outlook.mailbox#getcallbacktokenasync-options--isrest--callback--asyncresult-)|Ruft eine Zeichenfolge ab, die ein Token enthält, das zum Aufrufen von Rest-APIs oder Exchange-Webdienste verwendet wird.|
||[isrest](/javascript/api/outlook/outlook.mailbox#isrest)||
||[removeHandlerAsync (EventType: Office. EventType \| String, Options?: Office. AsyncContextOptions, Callback?: (asyncResult: Office. AsyncResult <void> ) => void)](/javascript/api/outlook/outlook.mailbox#removehandlerasync-eventtype--options--callback--asyncresult-)|Entfernt die Ereignishandler für einen unterstützten Ereignistyp.|
||[restUrl](/javascript/api/outlook/outlook.mailbox#resturl)|Ruft die URL des REST-Endpunkts für das betreffende E-Mail-Konto ab.|
