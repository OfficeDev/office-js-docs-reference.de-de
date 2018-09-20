# <a name="action-element"></a><span data-ttu-id="8163f-101">Action-Element</span><span class="sxs-lookup"><span data-stu-id="8163f-101">Action element</span></span>

<span data-ttu-id="8163f-102">Gibt die Aktion an, die durchgeführt werden soll, wenn der Benutzer ein [Schaltflächen](control.md#button-control)- oder [Menü](control.md#menu-dropdown-button-controls)steuerelement auswählt.</span><span class="sxs-lookup"><span data-stu-id="8163f-102">Specifies the action to perform when the user selects a  [Button](control.md#button-control) or [Menu](control.md#menu-dropdown-button-controls) controls.</span></span>
 
## <a name="attributes"></a><span data-ttu-id="8163f-103">Attribute</span><span class="sxs-lookup"><span data-stu-id="8163f-103">Attributes</span></span>

|  <span data-ttu-id="8163f-104">Attribut</span><span class="sxs-lookup"><span data-stu-id="8163f-104">Attribute</span></span>  |  <span data-ttu-id="8163f-105">Erforderlich</span><span class="sxs-lookup"><span data-stu-id="8163f-105">Required</span></span>  |  <span data-ttu-id="8163f-106">Beschreibung</span><span class="sxs-lookup"><span data-stu-id="8163f-106">Description</span></span>  |
|:-----|:-----|:-----|
|  [<span data-ttu-id="8163f-107">xsi:type</span><span class="sxs-lookup"><span data-stu-id="8163f-107">xsi:type</span></span>](#xsitype)  |  <span data-ttu-id="8163f-108">Ja</span><span class="sxs-lookup"><span data-stu-id="8163f-108">Yes</span></span>  | <span data-ttu-id="8163f-109">Die Art der durchzuführenden Aktion</span><span class="sxs-lookup"><span data-stu-id="8163f-109">Action type to take</span></span>|

## <a name="child-elements"></a><span data-ttu-id="8163f-110">Untergeordnete Elemente</span><span class="sxs-lookup"><span data-stu-id="8163f-110">Child elements</span></span>

|  <span data-ttu-id="8163f-111">Element</span><span class="sxs-lookup"><span data-stu-id="8163f-111">Element</span></span> |  <span data-ttu-id="8163f-112">Beschreibung</span><span class="sxs-lookup"><span data-stu-id="8163f-112">Description</span></span>  |
|:-----|:-----|
|  [<span data-ttu-id="8163f-113">FunctionName</span><span class="sxs-lookup"><span data-stu-id="8163f-113">FunctionName</span></span>](#functionname) |    <span data-ttu-id="8163f-114">Gibt den Namen der auszuführenden Funktion an.</span><span class="sxs-lookup"><span data-stu-id="8163f-114">Specifies the name of the function to execute.</span></span> |
|  [<span data-ttu-id="8163f-115">SourceLocation</span><span class="sxs-lookup"><span data-stu-id="8163f-115">SourceLocation</span></span>](#sourcelocation) |    <span data-ttu-id="8163f-116">Gibt den Speicherort der Quelldatei für diese Aktion an.</span><span class="sxs-lookup"><span data-stu-id="8163f-116">Specifies the source file location for this action.</span></span> |
|  [<span data-ttu-id="8163f-117">TaskpaneId</span><span class="sxs-lookup"><span data-stu-id="8163f-117">TaskpaneId</span></span>](#taskpaneid) | <span data-ttu-id="8163f-118">Gibt die ID des Aufgabenbereichscontainers an.</span><span class="sxs-lookup"><span data-stu-id="8163f-118">Specifies the ID of the task pane container.</span></span>|
|  [<span data-ttu-id="8163f-119">Title</span><span class="sxs-lookup"><span data-stu-id="8163f-119">Title</span></span>](#title) | <span data-ttu-id="8163f-120">Gibt den benutzerdefinierten Titel für den Aufgabenbereich an.</span><span class="sxs-lookup"><span data-stu-id="8163f-120">Specifies the custom title for the task pane.</span></span>|
|  [<span data-ttu-id="8163f-121">SupportsPinning</span><span class="sxs-lookup"><span data-stu-id="8163f-121">SupportsPinning</span></span>](#supportspinning) | <span data-ttu-id="8163f-122">Gibt an, dass ein Aufgabenbereich das Anheften unterstützt, sodass der Aufgabenbereich geöffnet bleibt, wenn der Benutzer die Auswahl ändert.</span><span class="sxs-lookup"><span data-stu-id="8163f-122">Specifies that a task pane supports pinning, which keeps the task pane open when the user changes the selection.</span></span>|
  

## <a name="xsitype"></a><span data-ttu-id="8163f-123">xsi:type</span><span class="sxs-lookup"><span data-stu-id="8163f-123">xsi:type</span></span>

<span data-ttu-id="8163f-p101">Dieses Attribut gibt die Art von Aktion an, die durchgeführt wird, wenn der Benutzer auf die Schaltfläche klickt. Das kann eine der folgenden Aktionen sein:</span><span class="sxs-lookup"><span data-stu-id="8163f-p101">This attribute specifies the kind of action performed when the user selects the button. It can be one of the following:</span></span>

- `ExecuteFunction`
- `ShowTaskpane`

## <a name="functionname"></a><span data-ttu-id="8163f-126">FunctionName</span><span class="sxs-lookup"><span data-stu-id="8163f-126">FunctionName</span></span>

<span data-ttu-id="8163f-p102">Erforderliches Element, wenn **xsi:type** „ExecuteFunction“ ist. Gibt den Namen der auszuführenden Funktion an. Die Funktion ist in der Datei enthalten, die im Element [FunctionFile](functionfile.md) angegeben ist.</span><span class="sxs-lookup"><span data-stu-id="8163f-p102">Required element when **xsi:type** is "ExecuteFunction". Specifies the name of the function to execute. The function is contained in the file specified in the [FunctionFile](functionfile.md) element.</span></span>

```xml
<Action xsi:type="ExecuteFunction">
  <FunctionName>getSubject</FunctionName>
</Action>
```

## <a name="sourcelocation"></a><span data-ttu-id="8163f-130">SourceLocation</span><span class="sxs-lookup"><span data-stu-id="8163f-130">SourceLocation</span></span>

<span data-ttu-id="8163f-p103">Erforderliches Element, wenn  **xsi:type** "ShowTaskpane" ist. Gibt den Speicherort der Quelldatei für diese Aktion an. Das Attribut **resid** muss auf den Wert des **id**-Attributs eines  **Url**-Elements im  **Urls**-Element im  [Resources](resources.md) -Element festgelegt werden.</span><span class="sxs-lookup"><span data-stu-id="8163f-p103">Required element when  **xsi:type** is "ShowTaskpane". Specifies the source file location for this action. The **resid** attribute must be set to the value of the **id** attribute of a **Url** element in the **Urls** element in the [Resources](resources.md) element.</span></span>

```xml
<Action xsi:type="ShowTaskpane">
  <SourceLocation resid="readTaskPaneUrl" />
</Action>
```  

## <a name="taskpaneid"></a><span data-ttu-id="8163f-134">TaskpaneId</span><span class="sxs-lookup"><span data-stu-id="8163f-134">TaskpaneId</span></span>

<span data-ttu-id="8163f-p104">Optionales Element, wenn  **xsi:type** den Wert "ShowTaskpane" hat. Gibt die ID des Aufgabenbereichscontainers an. Wenn Sie über mehrere "ShowTaskpane"-Aktionen verfügen, verwenden Sie eine andere **TaskpaneId**einen unabhängigen Bereich für jede benötigen. Verwenden Sie dieselbe **TaskpaneId** für verschiedene Aktionen, die denselben Bereich nutzen. Wenn Benutzer Befehle mit derselben **TaskpaneId** verwenden, bleibt der Bereichscontainer geöffnet, aber der Inhalt des Bereichs wird durch die entsprechende Aktion "SourceLocation" ersetzt.</span><span class="sxs-lookup"><span data-stu-id="8163f-p104">Optional element when  **xsi:type** is "ShowTaskpane". Specifies the ID of the task pane container. When you have multiple "ShowTaskpane" actions, use a different **TaskpaneId** if you want an independent pane for each. Use the same **TaskpaneId** for  different actions that share the same pane. When users choose commands that share the same **TaskpaneId**, the pane container will remain open but the contents of the pane will be replaced with the corresponding Action "SourceLocation".</span></span> 

> [!NOTE]
> <span data-ttu-id="8163f-140">Dieses Element wird in Outlook nicht unterstützt.</span><span class="sxs-lookup"><span data-stu-id="8163f-140">This element is not supported in Outlook.</span></span>

<span data-ttu-id="8163f-141">Das folgende Beispiel zeigt zwei Aktionen, die dieselbe **TaskpaneId** verwenden.</span><span class="sxs-lookup"><span data-stu-id="8163f-141">The following example shows two actions that share the same **TaskpaneId**.</span></span> 

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

<span data-ttu-id="8163f-p105">Das folgende Beispiel zeigt zwei Aktionen, die eine andere **TaskpaneId** verwenden. Um diese Beispiele im Kontext zu sehen finden Sie weitere Informationen unter [Beispiel für einfache Add-In-Befehle](https://github.com/OfficeDev/Office-Add-in-Commands-Samples/blob/master/Simple/Manifest/SimpleAddin.xml).</span><span class="sxs-lookup"><span data-stu-id="8163f-p105">The following examples show two actions that use a different **TaskpaneId**. To see these examples in context, see [Simple Add-in Commands Sample](https://github.com/OfficeDev/Office-Add-in-Commands-Samples/blob/master/Simple/Manifest/SimpleAddin.xml).</span></span>

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

## <a name="title"></a><span data-ttu-id="8163f-144">Titel</span><span class="sxs-lookup"><span data-stu-id="8163f-144">Title</span></span>
<span data-ttu-id="8163f-p106">Optionales Element, wenn  **xsi:type** den Wert „ShowTaskpane“ hat. Gibt den benutzerdefinierten Titel für den Aufgabenbereich für diese Aktion an.</span><span class="sxs-lookup"><span data-stu-id="8163f-p106">Optional element when  **xsi:type** is "ShowTaskpane". Specifies the custom title for the task pane for this action.</span></span> 

<span data-ttu-id="8163f-147">Das folgende Beispiel zeigt zwei verschiedene Aktionen, die das **Title**-Element verwenden.</span><span class="sxs-lookup"><span data-stu-id="8163f-147">The following examples show two different actions that use the **Title** element.</span></span>

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

## <a name="supportspinning"></a><span data-ttu-id="8163f-148">SupportsPinning</span><span class="sxs-lookup"><span data-stu-id="8163f-148">SupportsPinning</span></span>

<span data-ttu-id="8163f-p107">Optionales Element, wenn **xsi:type** den Wert „ShowTaskpane“ hat. Die enthaltenden [VersionOverrides](versionoverrides.md)-Elemente müssen einen `xsi:type`-Attributwert von `VersionOverridesV1_1` aufweisen. Schließen Sie dieses Element in den Wert `true`, um das Anheften des Aufgabenbereichs zu unterstützen. Der Benutzer kann dann den Aufgabenbereich „anheften“, sodass dieser geöffnet bleibt, wenn die Auswahl geändert wird. Weitere Informationen finden Sie unter [Implementieren eines anheftbaren Aufgabenbereichs in Outlook](https://docs.microsoft.com/outlook/add-ins/pinnable-taskpane).</span><span class="sxs-lookup"><span data-stu-id="8163f-p107">Optional element when **xsi:type** is "ShowTaskpane". The containing [VersionOverrides](versionoverrides.md) elements must have an `xsi:type` attribute value of `VersionOverridesV1_1`. Include this element with a value of `true` to support taskpane pinning. The user will be able to "pin" the taskpane, causing it to stay open when changing the selection. For more information, see [Implement a pinnable taskpane in Outlook](https://docs.microsoft.com/outlook/add-ins/pinnable-taskpane).</span></span>

> [!NOTE]
> <span data-ttu-id="8163f-154">SupportsPinning derzeit nur unterstützt von Outlook 2016 für Windows (Build 7628.1000 oder höher).</span><span class="sxs-lookup"><span data-stu-id="8163f-154">SupportsPinning currently only supported by Outlook 2016 for Windows (build 7628.1000 or later).</span></span>

```xml
<Action xsi:type="ShowTaskpane">
  <SourceLocation resid="readTaskPaneUrl" />
  <SupportsPinning>true</SupportsPinning>
</Action>
```


