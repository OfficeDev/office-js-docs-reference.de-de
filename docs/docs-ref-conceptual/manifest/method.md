# <a name="method-element"></a><span data-ttu-id="bb637-101">Method-Element</span><span class="sxs-lookup"><span data-stu-id="bb637-101">Method element</span></span>

<span data-ttu-id="bb637-102">Gibt eine individuelle Methode aus der JavaScript-API für Office an, die Ihr Office-Add-In zum Aktivieren benötigt.</span><span class="sxs-lookup"><span data-stu-id="bb637-102">Specifies an individual method from the JavaScript API for Office that your Office Add-in requires in order to activate.</span></span>

<span data-ttu-id="bb637-103">**Add-In-Typ:** Inhalt, Aufgabenbereich</span><span class="sxs-lookup"><span data-stu-id="bb637-103">**Add-in type:** Content, Task pane</span></span>

## <a name="syntax"></a><span data-ttu-id="bb637-104">Syntax</span><span class="sxs-lookup"><span data-stu-id="bb637-104">Syntax</span></span>

```XML
<Method Name="string"/>
```

## <a name="contained-in"></a><span data-ttu-id="bb637-105">Enthalten in</span><span class="sxs-lookup"><span data-stu-id="bb637-105">Contained in</span></span>

[<span data-ttu-id="bb637-106">Methoden</span><span class="sxs-lookup"><span data-stu-id="bb637-106">Methods</span></span>](methods.md)

## <a name="attributes"></a><span data-ttu-id="bb637-107">Attribute</span><span class="sxs-lookup"><span data-stu-id="bb637-107">Attributes</span></span>

|<span data-ttu-id="bb637-108">**Attribut**</span><span class="sxs-lookup"><span data-stu-id="bb637-108">**Attribute**</span></span>|<span data-ttu-id="bb637-109">**Typ**</span><span class="sxs-lookup"><span data-stu-id="bb637-109">**Type**</span></span>|<span data-ttu-id="bb637-110">**Erforderlich**</span><span class="sxs-lookup"><span data-stu-id="bb637-110">**Required**</span></span>|<span data-ttu-id="bb637-111">**Beschreibung**</span><span class="sxs-lookup"><span data-stu-id="bb637-111">**Description**</span></span>|
|:-----|:-----|:-----|:-----|
|<span data-ttu-id="bb637-112">Name</span><span class="sxs-lookup"><span data-stu-id="bb637-112">Name</span></span>|<span data-ttu-id="bb637-113">string</span><span class="sxs-lookup"><span data-stu-id="bb637-113">string</span></span>|<span data-ttu-id="bb637-114">erforderlich</span><span class="sxs-lookup"><span data-stu-id="bb637-114">required</span></span>|<span data-ttu-id="bb637-p101">Gibt den Namen der erforderlichen Methode qualifiziert mit dem übergeordneten Objekt an. Beispiel: Um die **getSelectedDataAsync**-Methode anzugeben, müssen Sie `"Document.getSelectedDataAsync"` angeben.</span><span class="sxs-lookup"><span data-stu-id="bb637-p101">Specifies the name of the required method qualified with its parent object. For example, to specify the  **getSelectedDataAsync** method, you must specify `"Document.getSelectedDataAsync"`.</span></span>|

## <a name="remarks"></a><span data-ttu-id="bb637-117">Hinweise</span><span class="sxs-lookup"><span data-stu-id="bb637-117">Remarks</span></span>

<span data-ttu-id="bb637-118">Die **Methoden** und **Methode** Elemente werden von Mail-add-ins nicht unterstützt. Weitere Informationen zu anforderungssätze finden Sie unter [Office-Versionen und Anforderung festgelegt](https://docs.microsoft.com/office/dev/add-ins/develop/office-versions-and-requirement-sets).</span><span class="sxs-lookup"><span data-stu-id="bb637-118">The  **Methods** and **Method** elements aren't supported by mail add-ins. For more information about requirement sets, see [Office versions and requirement sets](https://docs.microsoft.com/office/dev/add-ins/develop/office-versions-and-requirement-sets).</span></span>

> [!IMPORTANT] 
> <span data-ttu-id="bb637-119">Da es nicht möglich ist, die Mindestversion-Anforderung für die einzelnen Methoden angeben, sollten, um sicherzustellen, dass eine Methode zur Laufzeit verfügbar ist Sie auch eine **if** -Anweisung verwenden beim Aufrufen dieser Methode im Skript Ihr Add-in.</span><span class="sxs-lookup"><span data-stu-id="bb637-119">Because there is no way to specify the minimum version requirement for individual methods, to make sure that a method is available at runtime, you should also use an **if** statement when calling that method in the script of your add-in.</span></span> <span data-ttu-id="bb637-120">Weitere Informationen hierzu finden Sie unter [Grundlegendes zur JavaScript-API für Office](https://docs.microsoft.com/office/dev/add-ins/develop/understanding-the-javascript-api-for-office).</span><span class="sxs-lookup"><span data-stu-id="bb637-120">For more information about how to do this, see [Understanding the JavaScript API for Office](https://docs.microsoft.com/office/dev/add-ins/develop/understanding-the-javascript-api-for-office).</span></span>

