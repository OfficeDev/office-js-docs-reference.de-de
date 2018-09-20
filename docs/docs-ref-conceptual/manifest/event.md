# <a name="event-element"></a>Element „Event“

Dieses Element definiert einen Ereignishandler in einem Add-In.

> [!NOTE] 
> Die `Event` Element ist derzeit nur von Outlook im Web in Office 365 unterstützt.

## <a name="attributes"></a>Attribute

|  Attribut  |  Erforderlich  |  Beschreibung  |
|:-----|:-----|:-----|
|  [Typ](#type-attribute)  |  Ja  | Gibt das zu behandelnde Ereignis an. |
|  [FunctionExecution](#functionexecution-attribute)  |  Ja  | Gibt den Ausführungsstil für den Ereignishandler an: asynchron oder synchron. Zurzeit werden nur synchrone Ereignishandler unterstützt. |
|  [FunctionName](#functionname-attribute)  |  Ja  | Gibt den Funktionsnamen des Ereignishandlers an. |

### <a name="type-attribute"></a>Attribut „Type“

Dieses Attribut ist erforderlich. Es gibt an, welches Ereignis den Ereignishandler aufrufen wird. Die möglichen Werte für dieses Attribut sind in der nachfolgenden Tabelle aufgeführt.

|  Ereignistyp  |  Beschreibung  |
|:-----|:-----|
|  `ItemSend`  |  Der Ereignishandler wird aufgerufen, wenn der Benutzer eine Nachricht oder eine Besprechungseinladung sendet.  |

### <a name="functionexecution-attribute"></a>Attribut „FunctionExecution“

Dieses Attribut ist erforderlich. Es MUSS auf `synchronous` gesetzt werden.

### <a name="functionname-attribute"></a>Attribut „FunctionName“

Dieses Attribut ist erforderlich. Es gibt den Funktionsnamen des Ereignishandlers an. Dieser Wert muss einem Funktionsnamen in der [Funktionsdatei](functionfile.md) des Add-Ins entsprechen.

```xml
<Event Type="ItemSend" FunctionExecution="synchronous" FunctionName="itemSendHandler" /> 
```