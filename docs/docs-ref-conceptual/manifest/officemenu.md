# <a name="officemenu-element"></a><span data-ttu-id="9d2ef-101">OfficeMenu-Element</span><span class="sxs-lookup"><span data-stu-id="9d2ef-101">OfficeMenu element</span></span>

<span data-ttu-id="9d2ef-p101">Definiert eine Sammlung von Steuerelementen, die dem Office-Kontextmenü hinzugefügt werden sollen. Gilt für Word-, Excel-, PowerPoint- und OneNote-Add-Ins.</span><span class="sxs-lookup"><span data-stu-id="9d2ef-p101">Defines a collection of controls to be added to the Office context menu. Applies to Word, Excel, PowerPoint, and OneNote add-ins.</span></span>

## <a name="attributes"></a><span data-ttu-id="9d2ef-104">Attribute</span><span class="sxs-lookup"><span data-stu-id="9d2ef-104">Attributes</span></span>

| <span data-ttu-id="9d2ef-105">Attribut</span><span class="sxs-lookup"><span data-stu-id="9d2ef-105">Attribute</span></span>            | <span data-ttu-id="9d2ef-106">Erforderlich</span><span class="sxs-lookup"><span data-stu-id="9d2ef-106">Required</span></span> | <span data-ttu-id="9d2ef-107">Beschreibung</span><span class="sxs-lookup"><span data-stu-id="9d2ef-107">Description</span></span>                          |
|:---------------------|:--------:|:-------------------------------------|
| [<span data-ttu-id="9d2ef-108">xsi:type</span><span class="sxs-lookup"><span data-stu-id="9d2ef-108">xsi:type</span></span>](#xsitype) | <span data-ttu-id="9d2ef-109">Ja</span><span class="sxs-lookup"><span data-stu-id="9d2ef-109">Yes</span></span>      | <span data-ttu-id="9d2ef-110">Der Typ des OfficeMenu-Elements, das definiert wird.</span><span class="sxs-lookup"><span data-stu-id="9d2ef-110">The type of OfficeMenu being defined.</span></span>|

## <a name="child-elements"></a><span data-ttu-id="9d2ef-111">Untergeordnete Elemente</span><span class="sxs-lookup"><span data-stu-id="9d2ef-111">Child elements</span></span>

|  <span data-ttu-id="9d2ef-112">Element</span><span class="sxs-lookup"><span data-stu-id="9d2ef-112">Element</span></span> |  <span data-ttu-id="9d2ef-113">Erforderlich</span><span class="sxs-lookup"><span data-stu-id="9d2ef-113">Required</span></span>  |  <span data-ttu-id="9d2ef-114">Beschreibung</span><span class="sxs-lookup"><span data-stu-id="9d2ef-114">Description</span></span>  |
|:-----|:-----|:-----|
|  [<span data-ttu-id="9d2ef-115">Control</span><span class="sxs-lookup"><span data-stu-id="9d2ef-115">Control</span></span>](#control)    | <span data-ttu-id="9d2ef-116">Ja</span><span class="sxs-lookup"><span data-stu-id="9d2ef-116">Yes</span></span> |  <span data-ttu-id="9d2ef-117">Eine Auflistung eines oder mehrerer Control-Objekte.</span><span class="sxs-lookup"><span data-stu-id="9d2ef-117">A collection of one or more Control objects.</span></span>  |

## <a name="xsitype"></a><span data-ttu-id="9d2ef-118">xsi:type</span><span class="sxs-lookup"><span data-stu-id="9d2ef-118">xsi:type</span></span>

<span data-ttu-id="9d2ef-119">Gibt ein integriertes Menü der Office-Clientanwendung an, zu dem dieses Office-Add-in hinzugefügt werden soll.</span><span class="sxs-lookup"><span data-stu-id="9d2ef-119">Specifies a built-in menu of the Office client application on which to add this Office Add-in.</span></span>

- <span data-ttu-id="9d2ef-p102">`ContextMenuText` -  Zeigt das Element im Kontextmenü an, wenn Text ausgewählt wird und der Benutzer dann mit der rechten Maustaste auf den ausgewählten Text klickt. Gilt für Word, Excel, PowerPoint und OneNote.</span><span class="sxs-lookup"><span data-stu-id="9d2ef-p102">`ContextMenuText` -  Displays the item on the context menu when text is selected and the user opens the context menu (right-clicks) on the selected text. Applies to Word, Excel, PowerPoint, and OneNote.</span></span>
- <span data-ttu-id="9d2ef-p103">`ContextMenuCell` - Zeigt das Element im Kontextmenü an, wenn der Benutzer mit der rechten Maustaste in eine Zelle in dem Arbeitsblatt klickt. Gilt für Excel.</span><span class="sxs-lookup"><span data-stu-id="9d2ef-p103">`ContextMenuCell` -  Displays the item on the context menu when the user opens the context menu (right-clicks) on a cell on the spreadsheet. Applies to Excel.</span></span> 

## <a name="control"></a><span data-ttu-id="9d2ef-124">Steuerelement</span><span class="sxs-lookup"><span data-stu-id="9d2ef-124">Control</span></span>

<span data-ttu-id="9d2ef-125">Jedes **OfficeMenu**-Element erfordert mindestens ein oder mehrere [menu](control.md#menu-dropdown-button-controls)- Steuerelemente.</span><span class="sxs-lookup"><span data-stu-id="9d2ef-125">Each **OfficeMenu** element requires at one or more [menu](control.md#menu-dropdown-button-controls) controls.</span></span> 

## <a name="example"></a><span data-ttu-id="9d2ef-126">Beispiel</span><span class="sxs-lookup"><span data-stu-id="9d2ef-126">Example</span></span>

```xml
<OfficeMenu id="ContextMenuCell">
    <Control xsi:type="Menu" id="myMenuID">
      <Label resid="residLabel3" />
      <Supertip>
          <Title resid="residLabel" />
          <Description resid="residToolTip" />
      </Supertip>   
      <Icon>
        <bt:Image size="16" resid="icon1_16x16" />
        <bt:Image size="32" resid="icon1_32x32" />
        <bt:Image size="80" resid="icon1_80x80" />
      </Icon>    
      <Items>
        <Item id="myMenuItemID">
          <Label resid="residLabel3"/>
          <Supertip>
            <Title resid="residLabel" />
            <Description resid="residToolTip" />
          </Supertip>
          <Icon>
            <bt:Image size="16" resid="icon1_16x16" />
            <bt:Image size="32" resid="icon1_32x32" />
            <bt:Image size="80" resid="icon1_80x80" />
          </Icon>    
          <Action xsi:type="ShowTaskpane">
            <SourceLocation resid="residTaskpaneUrl2" />    
          </Action>    
        </Item>
      </Items>
    </Control>   
</OfficeMenu>
```
