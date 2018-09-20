# <a name="control-element"></a><span data-ttu-id="bc1f3-101">Control-Element</span><span class="sxs-lookup"><span data-stu-id="bc1f3-101">Control element</span></span>

<span data-ttu-id="bc1f3-p101">Definiert eine JavaScript-Funktion, die eine Aktion ausführt oder einen Aufgabenbereich startet. Ein **Control**-Element kann entweder eine Schaltfläche oder eine Menüoption sein. Mindestens ein **Control**-Element muss im [Group](group.md)-Element enthalten sein.</span><span class="sxs-lookup"><span data-stu-id="bc1f3-p101">Defines a JavaScript function that executes an action or launches a task pane. A **Control** element can be either a button or a menu option. At least one **Control** must be included in a [Group](group.md) element.</span></span>

## <a name="attributes"></a><span data-ttu-id="bc1f3-105">Attribute</span><span class="sxs-lookup"><span data-stu-id="bc1f3-105">Attributes</span></span>

|  <span data-ttu-id="bc1f3-106">Attribut</span><span class="sxs-lookup"><span data-stu-id="bc1f3-106">Attribute</span></span>  |  <span data-ttu-id="bc1f3-107">Erforderlich</span><span class="sxs-lookup"><span data-stu-id="bc1f3-107">Required</span></span>  |  <span data-ttu-id="bc1f3-108">Beschreibung</span><span class="sxs-lookup"><span data-stu-id="bc1f3-108">Description</span></span>  |
|:-----|:-----|:-----|
|<span data-ttu-id="bc1f3-109">**xsi:type**</span><span class="sxs-lookup"><span data-stu-id="bc1f3-109">**xsi:type**</span></span>|<span data-ttu-id="bc1f3-110">Ja</span><span class="sxs-lookup"><span data-stu-id="bc1f3-110">Yes</span></span>|<span data-ttu-id="bc1f3-p102">Der Typ des Steuerelements, das definiert wird. Kann `Button`, `Menu` oder `MobileButton` sein.</span><span class="sxs-lookup"><span data-stu-id="bc1f3-p102">The type of control being defined. Can be either `Button`, `Menu`, or `MobileButton`.</span></span> |
|<span data-ttu-id="bc1f3-113">**id**</span><span class="sxs-lookup"><span data-stu-id="bc1f3-113">**id**</span></span>|<span data-ttu-id="bc1f3-114">Nein</span><span class="sxs-lookup"><span data-stu-id="bc1f3-114">No</span></span>|<span data-ttu-id="bc1f3-p103">Die ID des Steuerelements. Darf maximal 125 Zeichen lang sein.</span><span class="sxs-lookup"><span data-stu-id="bc1f3-p103">The ID of the control element. Can be a maximum of 125 characters.</span></span>|

> [!NOTE]
> <span data-ttu-id="bc1f3-117">Die `MobileButton` -Wert für **xsi: Type** im VersionOverrides-Schema 1.1 definiert ist.</span><span class="sxs-lookup"><span data-stu-id="bc1f3-117">The `MobileButton` value for **xsi:type** is defined in VersionOverrides schema 1.1.</span></span> <span data-ttu-id="bc1f3-118">Dies gilt nur für die **Steuerelement** -Elemente, die innerhalb eines [MobileFormFactor](mobileformfactor.md) -Elements enthalten sind.</span><span class="sxs-lookup"><span data-stu-id="bc1f3-118">It only applies to the **Control** elements contained within a [MobileFormFactor](mobileformfactor.md) element.</span></span>

## <a name="button-control"></a><span data-ttu-id="bc1f3-119">Schaltflächensteuerelement</span><span class="sxs-lookup"><span data-stu-id="bc1f3-119">Button control</span></span>

<span data-ttu-id="bc1f3-p105">Eine Schaltfläche führt bei Auswahl durch den Benutzer eine einzelne Aktion aus. Sie kann eine Funktion ausführen oder einen Aufgabenbereich anzeigen. Jedes Schaltflächen-Steuerelement muss eine für das Manifest eindeutige `id` aufweisen.</span><span class="sxs-lookup"><span data-stu-id="bc1f3-p105">A button performs a single action when the user selects it. It can either execute a function or show a task pane. Each button control must have an `id` unique to the manifest.</span></span> 

### <a name="child-elements"></a><span data-ttu-id="bc1f3-123">Untergeordnete Elemente</span><span class="sxs-lookup"><span data-stu-id="bc1f3-123">Child elements</span></span>
|  <span data-ttu-id="bc1f3-124">Element</span><span class="sxs-lookup"><span data-stu-id="bc1f3-124">Element</span></span> |  <span data-ttu-id="bc1f3-125">Erforderlich</span><span class="sxs-lookup"><span data-stu-id="bc1f3-125">Required</span></span>  |  <span data-ttu-id="bc1f3-126">Beschreibung</span><span class="sxs-lookup"><span data-stu-id="bc1f3-126">Description</span></span>  |
|:-----|:-----|:-----|
|  <span data-ttu-id="bc1f3-127">**Label**</span><span class="sxs-lookup"><span data-stu-id="bc1f3-127">**Label**</span></span>     | <span data-ttu-id="bc1f3-128">Ja</span><span class="sxs-lookup"><span data-stu-id="bc1f3-128">Yes</span></span> |  <span data-ttu-id="bc1f3-p106">Der Text für die Schaltfläche. Das Attribut  **resid** muss auf den Wert des Attributs **id** eines **String**-Elements im  **ShortStrings**-Element im  [Resources](resources.md)-Element festgelegt werden.</span><span class="sxs-lookup"><span data-stu-id="bc1f3-p106">The text for the button. The  **resid** attribute must be set to the value of the **id** attribute of a **String** element in the **ShortStrings** element in the [Resources](resources.md)  element.</span></span>        |
|  <span data-ttu-id="bc1f3-131">**ToolTip**</span><span class="sxs-lookup"><span data-stu-id="bc1f3-131">**ToolTip**</span></span>  |<span data-ttu-id="bc1f3-132">Nein</span><span class="sxs-lookup"><span data-stu-id="bc1f3-132">No</span></span>|<span data-ttu-id="bc1f3-p107">Die QuickInfo für die Schaltfläche. Das Attribut **resid**  muss auf den Wert des Attributs **id** eines **String**-Elements festgelegt sein. Das Element **String** ist ein untergeordnetes Element des Elements **LongStrings**, und dieses ist ein untergeordnetes Element des Elements [Resources](resources.md).</span><span class="sxs-lookup"><span data-stu-id="bc1f3-p107">The tooltip for the button. The  **resid** attribute must be set to the value of the **id** attribute of a **String** element. The **String** element is a child of the **LongStrings** element, which is a child of the [Resources](resources.md) element.</span></span>|        
|  [<span data-ttu-id="bc1f3-136">Supertip</span><span class="sxs-lookup"><span data-stu-id="bc1f3-136">Supertip</span></span>](supertip.md)  | <span data-ttu-id="bc1f3-137">Ja</span><span class="sxs-lookup"><span data-stu-id="bc1f3-137">Yes</span></span> |  <span data-ttu-id="bc1f3-138">Die Multiinfo für die Schaltfläche.</span><span class="sxs-lookup"><span data-stu-id="bc1f3-138">The supertip for the button.</span></span>    |
|  [<span data-ttu-id="bc1f3-139">Icon</span><span class="sxs-lookup"><span data-stu-id="bc1f3-139">Icon</span></span>](icon.md)      | <span data-ttu-id="bc1f3-140">Ja</span><span class="sxs-lookup"><span data-stu-id="bc1f3-140">Yes</span></span> |  <span data-ttu-id="bc1f3-141">Ein Bild für die Schaltfläche.</span><span class="sxs-lookup"><span data-stu-id="bc1f3-141">An image for the button.</span></span>         |
|  [<span data-ttu-id="bc1f3-142">Action</span><span class="sxs-lookup"><span data-stu-id="bc1f3-142">Action</span></span>](action.md)    | <span data-ttu-id="bc1f3-143">Ja</span><span class="sxs-lookup"><span data-stu-id="bc1f3-143">Yes</span></span> |  <span data-ttu-id="bc1f3-144">Gibt die auszuführende Aktion an.</span><span class="sxs-lookup"><span data-stu-id="bc1f3-144">Specifies the action to perform.</span></span>  |

### <a name="executefunction-button-example"></a><span data-ttu-id="bc1f3-145">Beispiel für ExecuteFunction-Schaltfläche</span><span class="sxs-lookup"><span data-stu-id="bc1f3-145">ExecuteFunction button example</span></span>

```xml
<Control xsi:type="Button" id="msgReadFunctionButton">
  <Label resid="funcReadButtonLabel" />
  <Supertip>
    <Title resid="funcReadSuperTipTitle" />
    <Description resid="funcReadSuperTipDescription" />
  </Supertip>
  <Icon>
    <bt:Image size="16" resid="blue-icon-16" />
    <bt:Image size="32" resid="blue-icon-32" />
    <bt:Image size="80" resid="blue-icon-80" />
  </Icon>
  <Action xsi:type="ExecuteFunction">
    <FunctionName>getSubject</FunctionName>
  </Action>
</Control>
```

### <a name="showtaskpane-button-example"></a><span data-ttu-id="bc1f3-146">Beispiel für ShowTaskpane-Schaltfläche</span><span class="sxs-lookup"><span data-stu-id="bc1f3-146">ShowTaskpane button example</span></span>

```xml
<Control xsi:type="Button" id="msgReadOpenPaneButton">
  <Label resid="paneReadButtonLabel" />
  <Supertip>
    <Title resid="paneReadSuperTipTitle" />
    <Description resid="paneReadSuperTipDescription" />
  </Supertip>
  <Icon>
    <bt:Image size="16" resid="green-icon-16" />
    <bt:Image size="32" resid="green-icon-32" />
    <bt:Image size="80" resid="green-icon-80" />
  </Icon>
  <Action xsi:type="ShowTaskpane">
    <SourceLocation resid="readTaskPaneUrl" />
  </Action>
</Control>
```

## <a name="menu-dropdown-button-controls"></a><span data-ttu-id="bc1f3-147">Steuerelemente für Menüs (Dropdownschaltfläche)</span><span class="sxs-lookup"><span data-stu-id="bc1f3-147">Menu (dropdown button) controls</span></span>

<span data-ttu-id="bc1f3-p108">Ein Menü definiert eine statische Liste mit Optionen. Jedes Menüelement führt eine Funktion aus oder zeigt einen Aufgabenbereich an. Untermenüs werden nicht unterstützt.</span><span class="sxs-lookup"><span data-stu-id="bc1f3-p108">A menu defines a static list of options. Each menu item either executes a function or shows a task pane. Submenus are not supported.</span></span> 

<span data-ttu-id="bc1f3-151">Bei Verwendung mit einem **PrimaryCommandSurface**- oder **ContextMenu**[-Erweiterungspunkt](extensionpoint.md) definiert das Steuerelement Folgendes:</span><span class="sxs-lookup"><span data-stu-id="bc1f3-151">When used with a **PrimaryCommandSurface** or **ContextMenu** [extension point](extensionpoint.md), the menu control defines:</span></span>

- <span data-ttu-id="bc1f3-152">Ein Menüelement auf der Stammebene.</span><span class="sxs-lookup"><span data-stu-id="bc1f3-152">A root-level menu item.</span></span>

- <span data-ttu-id="bc1f3-153">Eine Liste der Untermenüelemente.</span><span class="sxs-lookup"><span data-stu-id="bc1f3-153">A list of submenu items.</span></span>

<span data-ttu-id="bc1f3-p109">Bei Verwendung mit  **PrimaryCommandSurface** zeigt das Menüelement auf der Stammebene eine Schaltfläche auf dem Menüband an. Wenn die Schaltfläche ausgewählt wird, wird das Untermenü als eine Dropdownliste angezeigt. Bei Verwendung mit **ContextMenu** wird ein Menüelement mit einem Untermenü im Kontextmenü eingefügt. In beiden Fällen können einzelne Untermenüelemente entweder eine JavaScript-Funktion ausführen oder einen Aufgabenbereich anzeigen. Derzeit wird nur eine Ebene von Untermenüs unterstützt.</span><span class="sxs-lookup"><span data-stu-id="bc1f3-p109">When used with **PrimaryCommandSurface**, the root menu item displays as a button on the ribbon. When the button is selected, the submenu displays as a drop-down list. When used with  **ContextMenu**, a menu item with a submenu is inserted on the context menu. In both cases, individual submenu items can either execute a JavaScript function or show a task pane. Only one level of submenus is supported at this time.</span></span>

<span data-ttu-id="bc1f3-p110">Das folgende Beispiel zeigt, wie ein Menüelement mit zwei Untermenüelementen definiert wird. Das erste Untermenüelement zeigt einen Aufgabenbereich an, das zweite Untermenüelement führt eine JavaScript-Funktion aus.</span><span class="sxs-lookup"><span data-stu-id="bc1f3-p110">The following example shows how to define a menu item with two submenu items. The first submenu item shows a task pane, and the second submenu item runs a JavaScript function.</span></span>

```xml
<Control xsi:type="Menu" id="TestMenu2">
  <Label resid="residLabel3" />
  <Tooltip resid="residToolTip" />
  <Supertip>
    <Title resid="residLabel" />
    <Description resid="residToolTip" />
  </Supertip>
  <Icon>
    <bt:Image size="16" resid="icon1_32x32" />
    <bt:Image size="32" resid="icon1_32x32" />
    <bt:Image size="80" resid="icon1_32x32" />
  </Icon>
  <Items>
    <Item id="showGallery2">
      <Label resid="residLabel3"/>
      <Supertip>
        <Title resid="residLabel" />
        <Description resid="residToolTip" />
      </Supertip>
      <Icon>
        <bt:Image size="16" resid="icon1_32x32" />
        <bt:Image size="32" resid="icon1_32x32" />
        <bt:Image size="80" resid="icon1_32x32" />
      </Icon>
      <Action xsi:type="ShowTaskpane">
        <TaskpaneId>MyTaskPaneID1</TaskpaneId>
        <SourceLocation resid="residUnitConverterUrl" />
      </Action>
    </Item>
    <Item id="showGallery3">
      <Label resid="residLabel5"/>
      <Supertip>
        <Title resid="residLabel" />
        <Description resid="residToolTip" />
      </Supertip>
      <Icon>
        <bt:Image size="16" resid="icon4_32x32" />
        <bt:Image size="32" resid="icon4_32x32" />
        <bt:Image size="80" resid="icon4_32x32" />
      </Icon>
      <Action xsi:type="ExecuteFunction">
        <FunctionName>getButton</FunctionName>
      </Action>
    </Item>
  </Items>
</Control>

```

### <a name="child-elements"></a><span data-ttu-id="bc1f3-161">Untergeordnete Elemente</span><span class="sxs-lookup"><span data-stu-id="bc1f3-161">Child elements</span></span>

|  <span data-ttu-id="bc1f3-162">Element</span><span class="sxs-lookup"><span data-stu-id="bc1f3-162">Element</span></span> |  <span data-ttu-id="bc1f3-163">Erforderlich</span><span class="sxs-lookup"><span data-stu-id="bc1f3-163">Required</span></span>  |  <span data-ttu-id="bc1f3-164">Beschreibung</span><span class="sxs-lookup"><span data-stu-id="bc1f3-164">Description</span></span>  |
|:-----|:-----|:-----|
|  <span data-ttu-id="bc1f3-165">**Label**</span><span class="sxs-lookup"><span data-stu-id="bc1f3-165">**Label**</span></span>     | <span data-ttu-id="bc1f3-166">Ja</span><span class="sxs-lookup"><span data-stu-id="bc1f3-166">Yes</span></span> |  <span data-ttu-id="bc1f3-p111">Der Text für die Schaltfläche. Das Attribut  **resid** muss auf den Wert des Attributs **id** eines **String**-Elements im  **ShortStrings**-Element im  [Resources](resources.md)-Element festgelegt werden.</span><span class="sxs-lookup"><span data-stu-id="bc1f3-p111">The text for the button. The  **resid** attribute must be set to the value of the **id** attribute of a **String** element in the **ShortStrings** element in the [Resources](resources.md) element.</span></span>      |
|  <span data-ttu-id="bc1f3-169">**ToolTip**</span><span class="sxs-lookup"><span data-stu-id="bc1f3-169">**ToolTip**</span></span>  |<span data-ttu-id="bc1f3-170">Nein</span><span class="sxs-lookup"><span data-stu-id="bc1f3-170">No</span></span>|<span data-ttu-id="bc1f3-p112">Die QuickInfo für die Schaltfläche. Das Attribut **resid**  muss auf den Wert des Attributs **id** eines **String**-Elements festgelegt sein. Das Element **String** ist ein untergeordnetes Element des Elements **LongStrings**, und dieses ist ein untergeordnetes Element des Elements [Resources](resources.md).</span><span class="sxs-lookup"><span data-stu-id="bc1f3-p112">The tooltip for the button. The  **resid** attribute must be set to the value of the **id** attribute of a **String** element. The **String** element is a child of the **LongStrings** element, which is a child of the [Resources](resources.md) element.</span></span>|        
|  [<span data-ttu-id="bc1f3-174">Supertip</span><span class="sxs-lookup"><span data-stu-id="bc1f3-174">Supertip</span></span>](supertip.md)  | <span data-ttu-id="bc1f3-175">Ja</span><span class="sxs-lookup"><span data-stu-id="bc1f3-175">Yes</span></span> |  <span data-ttu-id="bc1f3-176">Die Multiinfo für diese Schaltfläche.</span><span class="sxs-lookup"><span data-stu-id="bc1f3-176">The supertip for this button.</span></span>    |
|  [<span data-ttu-id="bc1f3-177">Icon</span><span class="sxs-lookup"><span data-stu-id="bc1f3-177">Icon</span></span>](icon.md)      | <span data-ttu-id="bc1f3-178">Ja</span><span class="sxs-lookup"><span data-stu-id="bc1f3-178">Yes</span></span> |  <span data-ttu-id="bc1f3-179">Ein Bild für die Schaltfläche.</span><span class="sxs-lookup"><span data-stu-id="bc1f3-179">An image for the button.</span></span>         |
|  <span data-ttu-id="bc1f3-180">**Items**</span><span class="sxs-lookup"><span data-stu-id="bc1f3-180">**Items**</span></span>     | <span data-ttu-id="bc1f3-181">Ja</span><span class="sxs-lookup"><span data-stu-id="bc1f3-181">Yes</span></span> |  <span data-ttu-id="bc1f3-p113">Eine Auflistung von Schaltflächen, die innerhalb des Menüs angezeigt werden sollen. Enthält die  **Item**-Elemente für jedes Untermenüelement. Jedes **Item**-Element enthält die untergeordneten Elemente eines [Schaltflächensteuerelements](#button-control).</span><span class="sxs-lookup"><span data-stu-id="bc1f3-p113">A collection of Buttons to display within the menu. Contains the  **Item** elements for each submenu item. Each **Item** element contains the  child elements of the [Button control](#button-control).</span></span>|

### <a name="menu-control-examples"></a><span data-ttu-id="bc1f3-185">Beispiele für Menüsteuerelemente</span><span class="sxs-lookup"><span data-stu-id="bc1f3-185">Menu control examples</span></span>

```xml
<Control xsi:type="Menu" id="TestMenu2">
  <Label resid="residLabel3" />
  <Tooltip resid="residToolTip" />
  <Supertip>
    <Title resid="residLabel" />
    <Description resid="residToolTip" />
  </Supertip>
  <Icon>
    <bt:Image size="16" resid="icon1_32x32" />
    <bt:Image size="32" resid="icon1_32x32" />
    <bt:Image size="80" resid="icon1_32x32" />
  </Icon>
  <Items>
    <Item id="showGallery2">
      <Label resid="residLabel3"/>
      <Supertip>
        <Title resid="residLabel" />
        <Description resid="residToolTip" />
      </Supertip>
      <Icon>
        <bt:Image size="16" resid="icon1_32x32" />
        <bt:Image size="32" resid="icon1_32x32" />
        <bt:Image size="80" resid="icon1_32x32" />
      </Icon>
      <Action xsi:type="ShowTaskpane">
        <TaskpaneId>MyTaskPaneID1</TaskpaneId>
        <SourceLocation resid="residUnitConverterUrl" />
      </Action>
    </Item>
    <Item id="showGallery3">
      <Label resid="residLabel5"/>
      <Supertip>
        <Title resid="residLabel" />
        <Description resid="residToolTip" />
      </Supertip>
      <Icon>
        <bt:Image size="16" resid="icon4_32x32" />
        <bt:Image size="32" resid="icon4_32x32" />
        <bt:Image size="80" resid="icon4_32x32" />
      </Icon>
      <Action xsi:type="ExecuteFunction">
        <FunctionName>getButton</FunctionName>
      </Action>
    </Item>
  </Items>
</Control>

```

```xml
<Control xsi:type="Menu" id="msgReadMenuButton">
  <Label resid="menuReadButtonLabel" />
  <Supertip>
    <Title resid="menuReadSuperTipTitle" />
    <Description resid="menuReadSuperTipDescription" />
  </Supertip>
  <Icon>
    <bt:Image size="16" resid="red-icon-16" />
    <bt:Image size="32" resid="red-icon-32" />
    <bt:Image size="80" resid="red-icon-80" />
  </Icon>
  <Items>
    <Item id="msgReadMenuItem1">
      <Label resid="menuItem1ReadLabel" />
      <Supertip>
        <Title resid="menuItem1ReadLabel" />
        <Description resid="menuItem1ReadTip" />
      </Supertip>
      <Icon>
        <bt:Image size="16" resid="red-icon-16" />
        <bt:Image size="32" resid="red-icon-32" />
        <bt:Image size="80" resid="red-icon-80" />
      </Icon>
      <Action xsi:type="ExecuteFunction">
        <FunctionName>getItemClass</FunctionName>
      </Action>
    </Item>
  </Items>
</Control>
```

## <a name="mobilebutton-control"></a><span data-ttu-id="bc1f3-186">MobileButton-Steuerelement</span><span class="sxs-lookup"><span data-stu-id="bc1f3-186">MobileButton control</span></span>

<span data-ttu-id="bc1f3-p114">Eine mobile Schaltfläche führt bei Auswahl durch den Benutzer eine einzelne Aktion aus. Sie kann eine Funktion ausführen oder einen Aufgabenbereich anzeigen. Jedes mobile Schaltflächensteuerelement muss eine für das Manifest eindeutige `id` aufweisen.</span><span class="sxs-lookup"><span data-stu-id="bc1f3-p114">A mobile button performs a single action when the user selects it. It can either execute a function or show a task pane. Each mobile button control must have an `id` unique to the manifest.</span></span>

<span data-ttu-id="bc1f3-p115">Der `MobileButton`-Wert für **xsi:type** ist im VersionOverrides-Schema 1.1 definiert. Das enthaltende [VersionOverrides](versionoverrides.md)-Element muss einen `xsi:type`-Attributwert von `VersionOverridesV1_1` aufweisen.</span><span class="sxs-lookup"><span data-stu-id="bc1f3-p115">The `MobileButton` value for **xsi:type** is defined in VersionOverrides schema 1.1. The containing [VersionOverrides](versionoverrides.md) element must have an `xsi:type` attribute value of `VersionOverridesV1_1`.</span></span>

### <a name="child-elements"></a><span data-ttu-id="bc1f3-192">Untergeordnete Elemente</span><span class="sxs-lookup"><span data-stu-id="bc1f3-192">Child elements</span></span>
|  <span data-ttu-id="bc1f3-193">Element</span><span class="sxs-lookup"><span data-stu-id="bc1f3-193">Element</span></span> |  <span data-ttu-id="bc1f3-194">Erforderlich</span><span class="sxs-lookup"><span data-stu-id="bc1f3-194">Required</span></span>  |  <span data-ttu-id="bc1f3-195">Beschreibung</span><span class="sxs-lookup"><span data-stu-id="bc1f3-195">Description</span></span>  |
|:-----|:-----|:-----|
|  <span data-ttu-id="bc1f3-196">**Label**</span><span class="sxs-lookup"><span data-stu-id="bc1f3-196">**Label**</span></span>     | <span data-ttu-id="bc1f3-197">Ja</span><span class="sxs-lookup"><span data-stu-id="bc1f3-197">Yes</span></span> |  <span data-ttu-id="bc1f3-p116">Der Text für die Schaltfläche. Das Attribut  **resid** muss auf den Wert des Attributs **id** eines **String**-Elements im  **ShortStrings**-Element im  [Resources](resources.md)-Element festgelegt werden.</span><span class="sxs-lookup"><span data-stu-id="bc1f3-p116">The text for the button. The  **resid** attribute must be set to the value of the **id** attribute of a **String** element in the **ShortStrings** element in the [Resources](resources.md)  element.</span></span>        |
|  [<span data-ttu-id="bc1f3-200">Icon</span><span class="sxs-lookup"><span data-stu-id="bc1f3-200">Icon</span></span>](icon.md)      | <span data-ttu-id="bc1f3-201">Ja</span><span class="sxs-lookup"><span data-stu-id="bc1f3-201">Yes</span></span> |  <span data-ttu-id="bc1f3-202">Ein Bild für die Schaltfläche.</span><span class="sxs-lookup"><span data-stu-id="bc1f3-202">An image for the button.</span></span>         |
|  [<span data-ttu-id="bc1f3-203">Action</span><span class="sxs-lookup"><span data-stu-id="bc1f3-203">Action</span></span>](action.md)    | <span data-ttu-id="bc1f3-204">Ja</span><span class="sxs-lookup"><span data-stu-id="bc1f3-204">Yes</span></span> |  <span data-ttu-id="bc1f3-205">Gibt die auszuführende Aktion an.</span><span class="sxs-lookup"><span data-stu-id="bc1f3-205">Specifies the action to perform.</span></span>  |

### <a name="executefunction-mobile-button-example"></a><span data-ttu-id="bc1f3-206">Beispiel für mobile ExecuteFunction-Schaltfläche</span><span class="sxs-lookup"><span data-stu-id="bc1f3-206">ExecuteFunction mobile button example</span></span>

```xml
<Control xsi:type="MobileButton" id="msgReadFunctionButton">
  <Label resid="funcReadButtonLabel" />
  <Icon>
    <bt:Image resid="blue-icon-16-1" size="25" scale="1" />
    <bt:Image resid="blue-icon-16-2" size="25" scale="2" />
    <bt:Image resid="blue-icon-16-3" size="25" scale="3" />
    <bt:Image resid="blue-icon-32-1" size="32" scale="1" />
    <bt:Image resid="blue-icon-32-2" size="32" scale="2" />
    <bt:Image resid="blue-icon-32-3" size="32" scale="3" />
    <bt:Image resid="blue-icon-80-1" size="48" scale="1" />
    <bt:Image resid="blue-icon-80-2" size="48" scale="2" />
    <bt:Image resid="blue-icon-80-3" size="48" scale="3" />
  </Icon>
  <Action xsi:type="ExecuteFunction">
    <FunctionName>getSubject</FunctionName>
  </Action>
</Control>
```

### <a name="showtaskpane-mobile-button-example"></a><span data-ttu-id="bc1f3-207">Beispiel für mobile ShowTaskpane-Schaltfläche</span><span class="sxs-lookup"><span data-stu-id="bc1f3-207">ShowTaskpane mobile button example</span></span>

```xml
<Control xsi:type="MobileButton" id="msgReadOpenPaneButton">
  <Label resid="paneReadButtonLabel" />
  <Icon>
    <bt:Image resid="blue-icon-16-1" size="25" scale="1" />
    <bt:Image resid="blue-icon-16-2" size="25" scale="2" />
    <bt:Image resid="blue-icon-16-3" size="25" scale="3" />
    <bt:Image resid="blue-icon-32-1" size="32" scale="1" />
    <bt:Image resid="blue-icon-32-2" size="32" scale="2" />
    <bt:Image resid="blue-icon-32-3" size="32" scale="3" />
    <bt:Image resid="blue-icon-80-1" size="48" scale="1" />
    <bt:Image resid="blue-icon-80-2" size="48" scale="2" />
    <bt:Image resid="blue-icon-80-3" size="48" scale="3" />
  </Icon>
  <Action xsi:type="ShowTaskpane">
    <SourceLocation resid="readTaskPaneUrl" />
  </Action>
</Control>
```