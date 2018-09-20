# <a name="group-element"></a><span data-ttu-id="53bce-101">Group-Element</span><span class="sxs-lookup"><span data-stu-id="53bce-101">Group element</span></span>

<span data-ttu-id="53bce-p101">Definiert eine Gruppe von Benutzeroberflächensteuerelemente auf einer Registerkarte.  In benutzerdefinierten Registerkarten kann das Add-In bis zu 10 Gruppen erstellen. Jede Gruppe ist auf 6 Steuerelemente begrenzt, unabhängig davon, in welcher Registerkarte es angezeigt wird. Add-Ins sind auf eine benutzerdefinierte Registerkarte begrenzt.</span><span class="sxs-lookup"><span data-stu-id="53bce-p101">Defines a group of UI controls in a tab.  On custom tabs, the add-in can create up to 10 groups. Each group is limited to 6 controls, regardless of which tab it appears on. Add-ins are limited to one custom tab.</span></span>

## <a name="attributes"></a><span data-ttu-id="53bce-105">Attribute</span><span class="sxs-lookup"><span data-stu-id="53bce-105">Attributes</span></span>

|  <span data-ttu-id="53bce-106">Attribut</span><span class="sxs-lookup"><span data-stu-id="53bce-106">Attribute</span></span>  |  <span data-ttu-id="53bce-107">Erforderlich</span><span class="sxs-lookup"><span data-stu-id="53bce-107">Required</span></span>  |  <span data-ttu-id="53bce-108">Beschreibung</span><span class="sxs-lookup"><span data-stu-id="53bce-108">Description</span></span>  |
|:-----|:-----|:-----|
|  [<span data-ttu-id="53bce-109">id</span><span class="sxs-lookup"><span data-stu-id="53bce-109">id</span></span>](#id-attribute)  |  <span data-ttu-id="53bce-110">Ja</span><span class="sxs-lookup"><span data-stu-id="53bce-110">Yes</span></span>  | <span data-ttu-id="53bce-111">Eine eindeutige ID für die Gruppe.</span><span class="sxs-lookup"><span data-stu-id="53bce-111">A unique ID for the group.</span></span>|

### <a name="id-attribute"></a><span data-ttu-id="53bce-112">id-Attribut</span><span class="sxs-lookup"><span data-stu-id="53bce-112">id attribute</span></span>

<span data-ttu-id="53bce-p102">Erforderlich. Eindeutiger Bezeichner für die Gruppe. Es ist eine Zeichenfolge mit maximal 125 Zeichen. Dies muss im Manifest eindeutig sein, andernfalls wird die Gruppe nicht gerendert.</span><span class="sxs-lookup"><span data-stu-id="53bce-p102">Required. Unique identifier for the group. It is a string with a maximum of 125 characters. This must be unique within the manifest or the group will fail to render.</span></span>

## <a name="child-elements"></a><span data-ttu-id="53bce-117">Untergeordnete Elemente</span><span class="sxs-lookup"><span data-stu-id="53bce-117">Child elements</span></span>
|  <span data-ttu-id="53bce-118">Element</span><span class="sxs-lookup"><span data-stu-id="53bce-118">Element</span></span> |  <span data-ttu-id="53bce-119">Erforderlich</span><span class="sxs-lookup"><span data-stu-id="53bce-119">Required</span></span>  |  <span data-ttu-id="53bce-120">Beschreibung</span><span class="sxs-lookup"><span data-stu-id="53bce-120">Description</span></span>  |
|:-----|:-----|:-----|
|  [<span data-ttu-id="53bce-121">Label</span><span class="sxs-lookup"><span data-stu-id="53bce-121">Label</span></span>](#label)      | <span data-ttu-id="53bce-122">Ja</span><span class="sxs-lookup"><span data-stu-id="53bce-122">Yes</span></span> |  <span data-ttu-id="53bce-123">Die Bezeichnung für CustomTab oder eine Gruppe</span><span class="sxs-lookup"><span data-stu-id="53bce-123">The label for the CustomTab or a group.</span></span>  |
|  [<span data-ttu-id="53bce-124">Control</span><span class="sxs-lookup"><span data-stu-id="53bce-124">Control</span></span>](#control)    | <span data-ttu-id="53bce-125">Ja</span><span class="sxs-lookup"><span data-stu-id="53bce-125">Yes</span></span> |  <span data-ttu-id="53bce-126">Auflistung von einem oder mehreren Control-Objekten.</span><span class="sxs-lookup"><span data-stu-id="53bce-126">Collection of one or more Control objects.</span></span>  |

### <a name="label"></a><span data-ttu-id="53bce-127">Label</span><span class="sxs-lookup"><span data-stu-id="53bce-127">Label</span></span> 

<span data-ttu-id="53bce-p103">Erforderlich. Die Beschriftung der Gruppe. Das Attribut  **resid** muss auf den Wert des Attributs **id** eines **String**-Elements im  **ShortStrings**-Element im  [Resources](resources.md) -Element festgelegt werden.</span><span class="sxs-lookup"><span data-stu-id="53bce-p103">Required. The label of the group. The  **resid** attribute must be set to the value of the **id** attribute of a **String** element in the **ShortStrings** element in the [Resources](resources.md) element.</span></span>

### <a name="control"></a><span data-ttu-id="53bce-131">Control</span><span class="sxs-lookup"><span data-stu-id="53bce-131">Control</span></span>
<span data-ttu-id="53bce-132">Eine Gruppe erfordert mindestens ein Steuerelement.</span><span class="sxs-lookup"><span data-stu-id="53bce-132">A group requires at least one control.</span></span>

```xml
<Group id="msgreadCustomTab.grp1">
    <Label resid="residCustomTabGroupLabel"/>
    <Control xsi:type="Button" id="Button2">
    <!-- information on the control -->
    </Control>
    <!-- other controls, as needed -->
</Group>
```