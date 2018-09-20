# <a name="customtab-element"></a><span data-ttu-id="7ecdc-101">CustomTab-Element</span><span class="sxs-lookup"><span data-stu-id="7ecdc-101">CustomTab element</span></span>

<span data-ttu-id="7ecdc-p101">Im Menüband geben Sie Registerkarte und Gruppe für die Add-In-Befehle an. Dies kann entweder in der Standardregisterkarte sein (  **Home**,  **Message** oder  **Meeting**), oder in einer durch das Add-In definierten benutzerdefinierten Registerkarte.</span><span class="sxs-lookup"><span data-stu-id="7ecdc-p101">On the ribbon, you specify which tab and group for their add-in commands. This can either be on the default tab (either  **Home**,  **Message**, or  **Meeting**), or on a custom tab defined by the add-in.</span></span>

<span data-ttu-id="7ecdc-p102">In benutzerdefinierten Registerkarten kann das Add-In bis zu 10 Gruppen erstellen. Jede Gruppe ist auf 6 Steuerelemente begrenzt, unabhängig davon, in welcher Registerkarte es angezeigt wird. Add-Ins sind auf eine benutzerdefinierte Registerkarte begrenzt.</span><span class="sxs-lookup"><span data-stu-id="7ecdc-p102">On custom tabs, the add-in can create up to 10 groups. Each group is limited to 6 controls, regardless of which tab it appears on. Add-ins are limited to one custom tab.</span></span>

<span data-ttu-id="7ecdc-107">Das Attribut  **id** muss im Manifest eindeutig sein.</span><span class="sxs-lookup"><span data-stu-id="7ecdc-107">The  **id** attribute must be unique within the manifest.</span></span>

## <a name="child-elements"></a><span data-ttu-id="7ecdc-108">Untergeordnete Elemente</span><span class="sxs-lookup"><span data-stu-id="7ecdc-108">Child elements</span></span>

|  <span data-ttu-id="7ecdc-109">Element</span><span class="sxs-lookup"><span data-stu-id="7ecdc-109">Element</span></span> |  <span data-ttu-id="7ecdc-110">Erforderlich</span><span class="sxs-lookup"><span data-stu-id="7ecdc-110">Required</span></span>  |  <span data-ttu-id="7ecdc-111">Beschreibung</span><span class="sxs-lookup"><span data-stu-id="7ecdc-111">Description</span></span>  |
|:-----|:-----|:-----|
|  [<span data-ttu-id="7ecdc-112">Group</span><span class="sxs-lookup"><span data-stu-id="7ecdc-112">Group</span></span>](group.md)      | <span data-ttu-id="7ecdc-113">Ja</span><span class="sxs-lookup"><span data-stu-id="7ecdc-113">Yes</span></span> |  <span data-ttu-id="7ecdc-114">Definiert eine Gruppe von Befehlen.</span><span class="sxs-lookup"><span data-stu-id="7ecdc-114">Defines a Group of commands.</span></span>  |
|  [<span data-ttu-id="7ecdc-115">Label</span><span class="sxs-lookup"><span data-stu-id="7ecdc-115">Label</span></span>](#label-tab)      | <span data-ttu-id="7ecdc-116">Ja</span><span class="sxs-lookup"><span data-stu-id="7ecdc-116">Yes</span></span> |  <span data-ttu-id="7ecdc-117">Die Bezeichnung für CustomTab oder eine Gruppe</span><span class="sxs-lookup"><span data-stu-id="7ecdc-117">The label for the CustomTab or a Group.</span></span>  |
|  [<span data-ttu-id="7ecdc-118">Control</span><span class="sxs-lookup"><span data-stu-id="7ecdc-118">Control</span></span>](control.md)    | <span data-ttu-id="7ecdc-119">Ja</span><span class="sxs-lookup"><span data-stu-id="7ecdc-119">Yes</span></span> |  <span data-ttu-id="7ecdc-120">Eine Auflistung von einem oder mehreren Control-Objekten.</span><span class="sxs-lookup"><span data-stu-id="7ecdc-120">A collection of one or more Control objects.</span></span>  |

### <a name="group"></a><span data-ttu-id="7ecdc-121">Group</span><span class="sxs-lookup"><span data-stu-id="7ecdc-121">Group</span></span>

<span data-ttu-id="7ecdc-p103">Erforderlich. Siehe [Group-Element](group.md).</span><span class="sxs-lookup"><span data-stu-id="7ecdc-p103">Required. See [Group element](group.md).</span></span>

### <a name="label-tab"></a><span data-ttu-id="7ecdc-124">Label (Tab)</span><span class="sxs-lookup"><span data-stu-id="7ecdc-124">Label (Tab)</span></span>

<span data-ttu-id="7ecdc-p104">Erforderlich. Die Beschriftung der benutzerdefinierten Registerkarte. Das Attribut  **resid** muss auf den Wert des Attributs **id** eines **String**-Elements im  **ShortStrings**-Element im  [Resources](resources.md) -Element festgelegt werden.</span><span class="sxs-lookup"><span data-stu-id="7ecdc-p104">Required. The label of the custom tab. The  **resid** attribute must be set to the value of the **id** attribute of a **String** element in the **ShortStrings** element in the [Resources](resources.md) element.</span></span>


## <a name="customtab-example"></a><span data-ttu-id="7ecdc-127">CustomTab-Beispiel</span><span class="sxs-lookup"><span data-stu-id="7ecdc-127">CustomTab example</span></span>

```xml
<ExtensionPoint xsi:type="MessageReadCommandSurface">
  <CustomTab id="TabCustom1">
    <Group id="msgreadCustomTab.grp1">
    </Group>
    <Label resid="customTabLabel1"/>
  </CustomTab>
</ExtensionPoint>
```