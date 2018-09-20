# <a name="action-element"></a>Action-Element

Gibt die Aktion an, die durchgeführt werden soll, wenn der Benutzer ein [Schaltflächen](control.md#button-control)- oder [Menü](control.md#menu-dropdown-button-controls)steuerelement auswählt.
 
## <a name="attributes"></a>Attribute

|  Attribut  |  Erforderlich  |  Beschreibung  |
|:-----|:-----|:-----|
|  [xsi:type](#xsitype)  |  Ja  | Die Art der durchzuführenden Aktion|

## <a name="child-elements"></a>Untergeordnete Elemente

|  Element |  Beschreibung  |
|:-----|:-----|
|  [FunctionName](#functionname) |    Gibt den Namen der auszuführenden Funktion an. |
|  [SourceLocation](#sourcelocation) |    Gibt den Speicherort der Quelldatei für diese Aktion an. |
|  [TaskpaneId](#taskpaneid) | Gibt die ID des Aufgabenbereichscontainers an.|
|  [Title](#title) | Gibt den benutzerdefinierten Titel für den Aufgabenbereich an.|
|  [SupportsPinning](#supportspinning) | Gibt an, dass ein Aufgabenbereich das Anheften unterstützt, sodass der Aufgabenbereich geöffnet bleibt, wenn der Benutzer die Auswahl ändert.|
  

## <a name="xsitype"></a>xsi:type

Dieses Attribut gibt die Art von Aktion an, die durchgeführt wird, wenn der Benutzer auf die Schaltfläche klickt. Das kann eine der folgenden Aktionen sein:

- `ExecuteFunction`
- `ShowTaskpane`

## <a name="functionname"></a>FunctionName

Erforderliches Element, wenn **xsi:type** „ExecuteFunction“ ist. Gibt den Namen der auszuführenden Funktion an. Die Funktion ist in der Datei enthalten, die im Element [FunctionFile](functionfile.md) angegeben ist.

```xml
<Action xsi:type="ExecuteFunction">
  <FunctionName>getSubject</FunctionName>
</Action>
```

## <a name="sourcelocation"></a>SourceLocation

Erforderliches Element, wenn  **xsi:type** "ShowTaskpane" ist. Gibt den Speicherort der Quelldatei für diese Aktion an. Das Attribut **resid** muss auf den Wert des **id**-Attributs eines  **Url**-Elements im  **Urls**-Element im  [Resources](resources.md) -Element festgelegt werden.

```xml
<Action xsi:type="ShowTaskpane">
  <SourceLocation resid="readTaskPaneUrl" />
</Action>
```  

## <a name="taskpaneid"></a>TaskpaneId

Optionales Element, wenn  **xsi:type** den Wert "ShowTaskpane" hat. Gibt die ID des Aufgabenbereichscontainers an. Wenn Sie über mehrere "ShowTaskpane"-Aktionen verfügen, verwenden Sie eine andere **TaskpaneId**einen unabhängigen Bereich für jede benötigen. Verwenden Sie dieselbe **TaskpaneId** für verschiedene Aktionen, die denselben Bereich nutzen. Wenn Benutzer Befehle mit derselben **TaskpaneId** verwenden, bleibt der Bereichscontainer geöffnet, aber der Inhalt des Bereichs wird durch die entsprechende Aktion "SourceLocation" ersetzt. 

> [!NOTE]
> Dieses Element wird in Outlook nicht unterstützt.

Das folgende Beispiel zeigt zwei Aktionen, die dieselbe **TaskpaneId** verwenden. 

```xml
<Action xsi:type="ShowTaskpane">
  <TaskpaneId>MyPane</TaskpaneId>
  <SourceLocation resid="aTaskPaneUrl" />
</Action>

<Action xsi:type="ShowTaskpane">
  <TaskpaneId>MyPane</TaskpaneId>
  <SourceLocation resid="anotherTaskPaneUrl" />
</Action>
```  

Das folgende Beispiel zeigt zwei Aktionen, die eine andere **TaskpaneId** verwenden. Um diese Beispiele im Kontext zu sehen finden Sie weitere Informationen unter [Beispiel für einfache Add-In-Befehle](https://github.com/OfficeDev/Office-Add-in-Commands-Samples/blob/master/Simple/Manifest/SimpleAddin.xml).

```xml
<Action xsi:type="ShowTaskpane">
   <TaskpaneId>MyTaskPaneID1</TaskpaneId>
   <SourceLocation resid="Contoso.Taskpane1.Url" />
</Action>

<Action xsi:type="ShowTaskpane">
   <TaskpaneId>MyTaskPaneID2</TaskpaneId>
   <SourceLocation resid="Contoso.Taskpane2.Url" />
</Action>
```  

```xml
<bt:Urls>
   <bt:Url id="Contoso.Taskpane1.Url" DefaultValue="https://commandsimple.azurewebsites.net/Taskpane.html" />
   <bt:Url id="Contoso.Taskpane2.Url" DefaultValue="https://commandsimple.azurewebsites.net/Taskpane2.html" />
</bt:Urls>
```  

## <a name="title"></a>Titel
Optionales Element, wenn  **xsi:type** den Wert „ShowTaskpane“ hat. Gibt den benutzerdefinierten Titel für den Aufgabenbereich für diese Aktion an. 

Das folgende Beispiel zeigt zwei verschiedene Aktionen, die das **Title**-Element verwenden.

```xml
<Action xsi:type="ShowTaskpane">
<TaskpaneId>Office.AutoShowTaskpaneWithDocument</TaskpaneId>
<SourceLocation resid="PG.Code.Url" />
<Title resid="PG.CodeCommand.Title" />
</Action>
``` 

```xml
<Action xsi:type="ShowTaskpane">
<SourceLocation resid="PG.Run.Url" />
<Title resid="PG.RunCommand.Title" />
</Action>
``` 

```xml
<bt:Urls>
<bt:Url id="PG.Code.Url" DefaultValue="https://localhost:3000?commands=1" />
<bt:Url id="PG.Run.Url" DefaultValue="https://localhost:3000/run.html" />
</bt:Urls>
``` 

```xml
<bt:ShortStrings>
<bt:String id="PG.CodeCommand.Title" DefaultValue="Code" />
<bt:String id="PG.RunCommand.Title" DefaultValue="Run" />
</bt:ShortStrings>
``` 

## <a name="supportspinning"></a>SupportsPinning

Optionales Element, wenn **xsi:type** den Wert „ShowTaskpane“ hat. Die enthaltenden [VersionOverrides](versionoverrides.md)-Elemente müssen einen `xsi:type`-Attributwert von `VersionOverridesV1_1` aufweisen. Schließen Sie dieses Element in den Wert `true`, um das Anheften des Aufgabenbereichs zu unterstützen. Der Benutzer kann dann den Aufgabenbereich „anheften“, sodass dieser geöffnet bleibt, wenn die Auswahl geändert wird. Weitere Informationen finden Sie unter [Implementieren eines anheftbaren Aufgabenbereichs in Outlook](https://docs.microsoft.com/outlook/add-ins/pinnable-taskpane).

> [!NOTE]
> SupportsPinning derzeit nur unterstützt von Outlook 2016 für Windows (Build 7628.1000 oder höher).

```xml
<Action xsi:type="ShowTaskpane">
  <SourceLocation resid="readTaskPaneUrl" />
  <SupportsPinning>true</SupportsPinning>
</Action>
```


