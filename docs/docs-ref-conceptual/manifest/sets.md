# <a name="sets-element"></a><span data-ttu-id="bd880-101">Sets-Element</span><span class="sxs-lookup"><span data-stu-id="bd880-101">Sets element</span></span>

<span data-ttu-id="bd880-102">Gibt die minimale Untergruppe der JavaScript-API für Office an, die Ihr Office-Add-In zur Aktivierung benötigt.</span><span class="sxs-lookup"><span data-stu-id="bd880-102">Specifies the minimum subset of the JavaScript API for Office that your Office Add-in requires in order to activate.</span></span>

<span data-ttu-id="bd880-103">**Add-In-Typ:** Inhalt, Aufgabenbereich, E-Mail</span><span class="sxs-lookup"><span data-stu-id="bd880-103">**Add-in type:** Content, Task pane, Mail</span></span>

## <a name="syntax"></a><span data-ttu-id="bd880-104">Syntax</span><span class="sxs-lookup"><span data-stu-id="bd880-104">Syntax</span></span>

```XML
<Sets DefaultMinVersion="n .n ">
   ...
</Sets>
```

## <a name="contained-in"></a><span data-ttu-id="bd880-105">Enthalten in</span><span class="sxs-lookup"><span data-stu-id="bd880-105">Contained in</span></span>

[<span data-ttu-id="bd880-106">Requirements</span><span class="sxs-lookup"><span data-stu-id="bd880-106">Requirements</span></span>](requirements.md)

## <a name="can-contain"></a><span data-ttu-id="bd880-107">Kann enthalten</span><span class="sxs-lookup"><span data-stu-id="bd880-107">Can contain</span></span>

[<span data-ttu-id="bd880-108">Set</span><span class="sxs-lookup"><span data-stu-id="bd880-108">Set</span></span>](set.md)

## <a name="attributes"></a><span data-ttu-id="bd880-109">Attribute</span><span class="sxs-lookup"><span data-stu-id="bd880-109">Attributes</span></span>

|<span data-ttu-id="bd880-110">**Attribut**</span><span class="sxs-lookup"><span data-stu-id="bd880-110">**Attribute**</span></span>|<span data-ttu-id="bd880-111">**Typ**</span><span class="sxs-lookup"><span data-stu-id="bd880-111">**Type**</span></span>|<span data-ttu-id="bd880-112">**Erforderlich**</span><span class="sxs-lookup"><span data-stu-id="bd880-112">**Required**</span></span>|<span data-ttu-id="bd880-113">**Beschreibung**</span><span class="sxs-lookup"><span data-stu-id="bd880-113">**Description**</span></span>|
|:-----|:-----|:-----|:-----|
|<span data-ttu-id="bd880-114">DefaultMinVersion</span><span class="sxs-lookup"><span data-stu-id="bd880-114">DefaultMinVersion</span></span>|<span data-ttu-id="bd880-115">string</span><span class="sxs-lookup"><span data-stu-id="bd880-115">string</span></span>|<span data-ttu-id="bd880-116">Optional</span><span class="sxs-lookup"><span data-stu-id="bd880-116">optional</span></span>|<span data-ttu-id="bd880-p101">Gibt den Standardattributwert  **MinVersion** für alle untergeordneten [Set](set.md)-Elemente an. Der Standardwert lautet „1.1“.</span><span class="sxs-lookup"><span data-stu-id="bd880-p101">Specifies the default  **MinVersion** attribute value for all child [Set](set.md) elements. The default value is "1.1".</span></span>|

## <a name="remarks"></a><span data-ttu-id="bd880-119">Hinweise</span><span class="sxs-lookup"><span data-stu-id="bd880-119">Remarks</span></span>

<span data-ttu-id="bd880-120">Weitere Informationen zu anforderungssätze finden Sie unter [Office-Versionen und Anforderung festgelegt](https://docs.microsoft.com/office/dev/add-ins/develop/office-versions-and-requirement-sets).</span><span class="sxs-lookup"><span data-stu-id="bd880-120">For more information about requirement sets, see [Office versions and requirement sets](https://docs.microsoft.com/office/dev/add-ins/develop/office-versions-and-requirement-sets).</span></span>

<span data-ttu-id="bd880-121">Weitere Informationen zum **MinVersion**-Attribut des **Set**-Elements und zum **DefaultMinVersion**-Attribut des **Sets**-Elements finden Sie unter [Festlegen des Requirements-Element im Manifest](https://docs.microsoft.com/office/dev/add-ins/develop/specify-office-hosts-and-api-requirements#set-the-requirements-element-in-the-manifest).</span><span class="sxs-lookup"><span data-stu-id="bd880-121">For more information about the  **MinVersion** attribute of the **Set** element and the **DefaultMinVersion** attribute of the **Sets** element, see [Set the Requirements element in the manifest](https://docs.microsoft.com/office/dev/add-ins/develop/specify-office-hosts-and-api-requirements#set-the-requirements-element-in-the-manifest).</span></span>

