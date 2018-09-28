# <a name="supertip"></a><span data-ttu-id="b81e6-101">Multiinfo</span><span class="sxs-lookup"><span data-stu-id="b81e6-101">Supertip</span></span>

<span data-ttu-id="b81e6-p101">Definiert eine umfangreiche QuickInfo (Titel und Beschreibung). Es wird von [Button](control.md#button-control)- oder [Menu](control.md#menu-dropdown-button-controls)Steuerelementen verwendet.</span><span class="sxs-lookup"><span data-stu-id="b81e6-p101">Defines a rich tooltip (both Title and Description). It is used by both [Button](control.md#button-control) or [Menu](control.md#menu-dropdown-button-controls)  controls.</span></span>

## <a name="child-elements"></a><span data-ttu-id="b81e6-104">Untergeordnete Elemente</span><span class="sxs-lookup"><span data-stu-id="b81e6-104">Child elements</span></span>

|  <span data-ttu-id="b81e6-105">Element</span><span class="sxs-lookup"><span data-stu-id="b81e6-105">Element</span></span> |  <span data-ttu-id="b81e6-106">Erforderlich</span><span class="sxs-lookup"><span data-stu-id="b81e6-106">Required</span></span>  |  <span data-ttu-id="b81e6-107">Beschreibung</span><span class="sxs-lookup"><span data-stu-id="b81e6-107">Description</span></span>  |
|:-----|:-----|:-----|
|  [<span data-ttu-id="b81e6-108">Titel</span><span class="sxs-lookup"><span data-stu-id="b81e6-108">Title</span></span>](#title)        | <span data-ttu-id="b81e6-109">Ja</span><span class="sxs-lookup"><span data-stu-id="b81e6-109">Yes</span></span> |   <span data-ttu-id="b81e6-110">Der Text für die Multiinfo.</span><span class="sxs-lookup"><span data-stu-id="b81e6-110">The text for the supertip.</span></span>         |
|  [<span data-ttu-id="b81e6-111">Beschreibung</span><span class="sxs-lookup"><span data-stu-id="b81e6-111">Description</span></span>](#description)  | <span data-ttu-id="b81e6-112">Ja</span><span class="sxs-lookup"><span data-stu-id="b81e6-112">Yes</span></span> |  <span data-ttu-id="b81e6-113">Die Beschreibung für die Multiinfo.</span><span class="sxs-lookup"><span data-stu-id="b81e6-113">The description for the supertip.</span></span>    |

### <a name="title"></a><span data-ttu-id="b81e6-114">Title</span><span class="sxs-lookup"><span data-stu-id="b81e6-114">Title</span></span>

<span data-ttu-id="b81e6-115">Erforderlich.</span><span class="sxs-lookup"><span data-stu-id="b81e6-115">Required.</span></span> <span data-ttu-id="b81e6-116">Der Text für die Multiinfo.</span><span class="sxs-lookup"><span data-stu-id="b81e6-116">The text for the supertip.</span></span> <span data-ttu-id="b81e6-117">Das Attribut  **resid** muss auf den Wert des Attributs **id** eines **String**-Elements im  **ShortStrings**-Element im  [Resources](resources.md)-Element festgelegt werden.</span><span class="sxs-lookup"><span data-stu-id="b81e6-117">The  **resid** attribute must be set to the value of the **id** attribute of a **String** element in the **ShortStrings** element in the [Resources](resources.md) element.</span></span>

### <a name="description"></a><span data-ttu-id="b81e6-118">Beschreibung</span><span class="sxs-lookup"><span data-stu-id="b81e6-118">Description</span></span>

<span data-ttu-id="b81e6-119">Erforderlich.</span><span class="sxs-lookup"><span data-stu-id="b81e6-119">Required.</span></span> <span data-ttu-id="b81e6-120">Die Beschreibung für die Multiinfo.</span><span class="sxs-lookup"><span data-stu-id="b81e6-120">The description for the supertip.</span></span> <span data-ttu-id="b81e6-121">Das Attribut **Resid** muss auf den Wert des **Id** -Attributs ein **String** -Element im **LongStrings** -Element im [Ressourcen](resources.md) -Element festgelegt werden.</span><span class="sxs-lookup"><span data-stu-id="b81e6-121">The  **resid** attribute must be set to the value of the **id** attribute of a **String** element in the **LongStrings** element in the [Resources](resources.md) element.</span></span>

## <a name="example"></a><span data-ttu-id="b81e6-122">Beispiel</span><span class="sxs-lookup"><span data-stu-id="b81e6-122">Example</span></span>

```xml
<Supertip>
    <Title resid="funcReadSuperTipTitle" />
    <Description resid="funcReadSuperTipDescription" />
</Supertip>
```
