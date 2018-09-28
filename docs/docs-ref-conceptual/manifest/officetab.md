# <a name="officetab-element"></a><span data-ttu-id="60ed0-101">OfficeTab-Element</span><span class="sxs-lookup"><span data-stu-id="60ed0-101">OfficeTab element</span></span>

<span data-ttu-id="60ed0-p101">Definiert die Registerkarte des Menübands, auf der Ihr Add-In-Befehl angezeigt wird. Dies kann entweder die Standardregisterkarte sein (**Startseite**,  **Nachricht** oder  **Besprechung**), oder eine benutzerdefinierte Registerkarte, die vom Add-In definiert wird. Dieses Element ist erforderlich.</span><span class="sxs-lookup"><span data-stu-id="60ed0-p101">Defines the ribbon tab on which your add-in command appears. This can either be the default tab (either  **Home**,  **Message**, or  **Meeting**), or a custom tab defined by the add-in. This element is required.</span></span>

## <a name="child-elements"></a><span data-ttu-id="60ed0-105">Untergeordnete Elemente</span><span class="sxs-lookup"><span data-stu-id="60ed0-105">Child elements</span></span>

|  <span data-ttu-id="60ed0-106">Element</span><span class="sxs-lookup"><span data-stu-id="60ed0-106">Element</span></span> |  <span data-ttu-id="60ed0-107">Erforderlich</span><span class="sxs-lookup"><span data-stu-id="60ed0-107">Required</span></span>  |  <span data-ttu-id="60ed0-108">Beschreibung</span><span class="sxs-lookup"><span data-stu-id="60ed0-108">Description</span></span>  |
|:-----|:-----|:-----|
|  <span data-ttu-id="60ed0-109">Group</span><span class="sxs-lookup"><span data-stu-id="60ed0-109">Group</span></span>      | <span data-ttu-id="60ed0-110">Ja</span><span class="sxs-lookup"><span data-stu-id="60ed0-110">Yes</span></span> |  <span data-ttu-id="60ed0-p102">Definiert eine Gruppe von Befehlen. Sie können nur eine Gruppe pro Add-In zu der Standardregisterkarte hinzufügen.</span><span class="sxs-lookup"><span data-stu-id="60ed0-p102">Defines a group of commands. You can add only one group per add-in to the default tab.</span></span>  |

<span data-ttu-id="60ed0-113">Nachfolgend finden Sie die gültige `id`-Werte nach Host.</span><span class="sxs-lookup"><span data-stu-id="60ed0-113">The following are valid tab `id` values by host.</span></span> <span data-ttu-id="60ed0-114">Werte in **Fettschrift** werden in Desktop- und online (beispielsweise Word 2016 oder höher für Windows und Word Online) unterstützt.</span><span class="sxs-lookup"><span data-stu-id="60ed0-114">Values in **bold** are supported in both desktop and online (for example, Word 2016 or later for Windows and Word Online).</span></span>

### <a name="outlook"></a><span data-ttu-id="60ed0-115">Outlook</span><span class="sxs-lookup"><span data-stu-id="60ed0-115">Outlook</span></span>

- <span data-ttu-id="60ed0-116">**TabDefault**</span><span class="sxs-lookup"><span data-stu-id="60ed0-116">**TabDefault**</span></span>

### <a name="word"></a><span data-ttu-id="60ed0-117">Word</span><span class="sxs-lookup"><span data-stu-id="60ed0-117">Word</span></span>

- <span data-ttu-id="60ed0-118">**TabHome**</span><span class="sxs-lookup"><span data-stu-id="60ed0-118">**TabHome**</span></span>
- <span data-ttu-id="60ed0-119">**TabInsert**</span><span class="sxs-lookup"><span data-stu-id="60ed0-119">**TabInsert**</span></span>
- <span data-ttu-id="60ed0-120">TabWordDesign</span><span class="sxs-lookup"><span data-stu-id="60ed0-120">TabWordDesign</span></span>
- <span data-ttu-id="60ed0-121">**TabPageLayoutWord**</span><span class="sxs-lookup"><span data-stu-id="60ed0-121">**TabPageLayoutWord**</span></span>
- <span data-ttu-id="60ed0-122">TabReferences</span><span class="sxs-lookup"><span data-stu-id="60ed0-122">TabReferences</span></span>
- <span data-ttu-id="60ed0-123">TabMailings</span><span class="sxs-lookup"><span data-stu-id="60ed0-123">TabMailings</span></span>
- <span data-ttu-id="60ed0-124">TabReviewWord</span><span class="sxs-lookup"><span data-stu-id="60ed0-124">TabReviewWord</span></span>
- <span data-ttu-id="60ed0-125">**TabView**</span><span class="sxs-lookup"><span data-stu-id="60ed0-125">**TabView**</span></span>
- <span data-ttu-id="60ed0-126">TabDeveloper</span><span class="sxs-lookup"><span data-stu-id="60ed0-126">TabDeveloper</span></span>
- <span data-ttu-id="60ed0-127">TabAddIns</span><span class="sxs-lookup"><span data-stu-id="60ed0-127">TabAddIns</span></span>
- <span data-ttu-id="60ed0-128">TabBlogPost</span><span class="sxs-lookup"><span data-stu-id="60ed0-128">TabBlogPost</span></span>
- <span data-ttu-id="60ed0-129">TabBlogInsert</span><span class="sxs-lookup"><span data-stu-id="60ed0-129">TabBlogInsert</span></span>
- <span data-ttu-id="60ed0-130">TabPrintPreview</span><span class="sxs-lookup"><span data-stu-id="60ed0-130">TabPrintPreview</span></span>
- <span data-ttu-id="60ed0-131">TabOutlining</span><span class="sxs-lookup"><span data-stu-id="60ed0-131">TabOutlining</span></span>
- <span data-ttu-id="60ed0-132">TabConflicts</span><span class="sxs-lookup"><span data-stu-id="60ed0-132">TabConflicts</span></span>
- <span data-ttu-id="60ed0-133">TabBackgroundRemoval</span><span class="sxs-lookup"><span data-stu-id="60ed0-133">TabBackgroundRemoval</span></span>
- <span data-ttu-id="60ed0-134">TabBroadcastPresentation</span><span class="sxs-lookup"><span data-stu-id="60ed0-134">TabBroadcastPresentation</span></span>

### <a name="excel"></a><span data-ttu-id="60ed0-135">Excel</span><span class="sxs-lookup"><span data-stu-id="60ed0-135">Excel</span></span>

- <span data-ttu-id="60ed0-136">**TabHome**</span><span class="sxs-lookup"><span data-stu-id="60ed0-136">**TabHome**</span></span>
- <span data-ttu-id="60ed0-137">**TabInsert**</span><span class="sxs-lookup"><span data-stu-id="60ed0-137">**TabInsert**</span></span>
- <span data-ttu-id="60ed0-138">TabPageLayoutExcel</span><span class="sxs-lookup"><span data-stu-id="60ed0-138">TabPageLayoutExcel</span></span>
- <span data-ttu-id="60ed0-139">TabFormulas</span><span class="sxs-lookup"><span data-stu-id="60ed0-139">TabFormulas</span></span>
- <span data-ttu-id="60ed0-140">**TabData**</span><span class="sxs-lookup"><span data-stu-id="60ed0-140">**TabData**</span></span>
- <span data-ttu-id="60ed0-141">**TabReview**</span><span class="sxs-lookup"><span data-stu-id="60ed0-141">**TabReview**</span></span>
- <span data-ttu-id="60ed0-142">**TabView**</span><span class="sxs-lookup"><span data-stu-id="60ed0-142">**TabView**</span></span>
- <span data-ttu-id="60ed0-143">TabDeveloper</span><span class="sxs-lookup"><span data-stu-id="60ed0-143">TabDeveloper</span></span>
- <span data-ttu-id="60ed0-144">TabAddIns</span><span class="sxs-lookup"><span data-stu-id="60ed0-144">TabAddIns</span></span>
- <span data-ttu-id="60ed0-145">TabPrintPreview</span><span class="sxs-lookup"><span data-stu-id="60ed0-145">TabPrintPreview</span></span>
- <span data-ttu-id="60ed0-146">TabBackgroundRemoval</span><span class="sxs-lookup"><span data-stu-id="60ed0-146">TabBackgroundRemoval</span></span> 

### <a name="powerpoint"></a><span data-ttu-id="60ed0-147">PowerPoint</span><span class="sxs-lookup"><span data-stu-id="60ed0-147">PowerPoint</span></span>

- <span data-ttu-id="60ed0-148">**TabHome**</span><span class="sxs-lookup"><span data-stu-id="60ed0-148">**TabHome**</span></span>
- <span data-ttu-id="60ed0-149">**TabInsert**</span><span class="sxs-lookup"><span data-stu-id="60ed0-149">**TabInsert**</span></span>
- <span data-ttu-id="60ed0-150">**TabDesign**</span><span class="sxs-lookup"><span data-stu-id="60ed0-150">**TabDesign**</span></span>
- <span data-ttu-id="60ed0-151">**TabTransitions**</span><span class="sxs-lookup"><span data-stu-id="60ed0-151">**TabTransitions**</span></span>
- <span data-ttu-id="60ed0-152">**TabAnimations**</span><span class="sxs-lookup"><span data-stu-id="60ed0-152">**TabAnimations**</span></span>
- <span data-ttu-id="60ed0-153">TabSlideShow</span><span class="sxs-lookup"><span data-stu-id="60ed0-153">TabSlideShow</span></span>
- <span data-ttu-id="60ed0-154">TabReview</span><span class="sxs-lookup"><span data-stu-id="60ed0-154">TabReview</span></span>
- <span data-ttu-id="60ed0-155">**TabView**</span><span class="sxs-lookup"><span data-stu-id="60ed0-155">**TabView**</span></span>
- <span data-ttu-id="60ed0-156">TabDeveloper</span><span class="sxs-lookup"><span data-stu-id="60ed0-156">TabDeveloper</span></span>
- <span data-ttu-id="60ed0-157">TabAddIns</span><span class="sxs-lookup"><span data-stu-id="60ed0-157">TabAddIns</span></span>
- <span data-ttu-id="60ed0-158">TabPrintPreview</span><span class="sxs-lookup"><span data-stu-id="60ed0-158">TabPrintPreview</span></span>
- <span data-ttu-id="60ed0-159">TabMerge</span><span class="sxs-lookup"><span data-stu-id="60ed0-159">TabMerge</span></span>
- <span data-ttu-id="60ed0-160">TabGrayscale</span><span class="sxs-lookup"><span data-stu-id="60ed0-160">TabGrayscale</span></span>
- <span data-ttu-id="60ed0-161">TabBlackAndWhite</span><span class="sxs-lookup"><span data-stu-id="60ed0-161">TabBlackAndWhite</span></span>
- <span data-ttu-id="60ed0-162">TabBroadcastPresentation</span><span class="sxs-lookup"><span data-stu-id="60ed0-162">TabBroadcastPresentation</span></span>
- <span data-ttu-id="60ed0-163">TabSlideMaster</span><span class="sxs-lookup"><span data-stu-id="60ed0-163">TabSlideMaster</span></span>
- <span data-ttu-id="60ed0-164">TabHandoutMaster</span><span class="sxs-lookup"><span data-stu-id="60ed0-164">TabHandoutMaster</span></span>
- <span data-ttu-id="60ed0-165">TabNotesMaster</span><span class="sxs-lookup"><span data-stu-id="60ed0-165">TabNotesMaster</span></span>
- <span data-ttu-id="60ed0-166">TabBackgroundRemoval</span><span class="sxs-lookup"><span data-stu-id="60ed0-166">TabBackgroundRemoval</span></span>
- <span data-ttu-id="60ed0-167">TabSlideMasterHome</span><span class="sxs-lookup"><span data-stu-id="60ed0-167">TabSlideMasterHome</span></span>

### <a name="onenote"></a><span data-ttu-id="60ed0-168">OneNote</span><span class="sxs-lookup"><span data-stu-id="60ed0-168">OneNote</span></span>

- <span data-ttu-id="60ed0-169">**TabHome**</span><span class="sxs-lookup"><span data-stu-id="60ed0-169">**TabHome**</span></span>
- <span data-ttu-id="60ed0-170">**TabInsert**</span><span class="sxs-lookup"><span data-stu-id="60ed0-170">**TabInsert**</span></span>
- <span data-ttu-id="60ed0-171">**TabView**</span><span class="sxs-lookup"><span data-stu-id="60ed0-171">**TabView**</span></span>
- <span data-ttu-id="60ed0-172">TabDeveloper</span><span class="sxs-lookup"><span data-stu-id="60ed0-172">TabDeveloper</span></span>
- <span data-ttu-id="60ed0-173">TabAddIns</span><span class="sxs-lookup"><span data-stu-id="60ed0-173">TabAddIns</span></span>

## <a name="group"></a><span data-ttu-id="60ed0-174">Group</span><span class="sxs-lookup"><span data-stu-id="60ed0-174">Group</span></span>

<span data-ttu-id="60ed0-p104">Eine Gruppe von UI-Erweiterungspunkten auf einer Registerkarte. Eine Gruppe kann bis zu sechs Steuerelemente enthalten. Das **id**-Attribut ist erforderlich, und jede **id** muss im Manifest eindeutig sein. Die **id** ist eine Zeichenfolge mit maximal 125 Zeichen. Siehe [Group-Element](group.md).</span><span class="sxs-lookup"><span data-stu-id="60ed0-p104">A group of UI extension points in a tab. A group can have up to six controls. The  **id** attribute is required and each **id** must be unique within the manifest. The **id** is a string with a maximum of 125 characters. See [Group element](group.md).</span></span>

## <a name="officetab-example"></a><span data-ttu-id="60ed0-179">OfficeTab-Beispiel </span><span class="sxs-lookup"><span data-stu-id="60ed0-179">OfficeTab example</span></span>

```xml
<ExtensionPoint xsi:type="MessageReadCommandSurface">
  <OfficeTab id="TabDefault">
    <Group id="msgreadTabMessage.grp1">
        <!-- Group Definition -->
    </Group>
  </OfficeTab>
</ExtensionPoint>
```
