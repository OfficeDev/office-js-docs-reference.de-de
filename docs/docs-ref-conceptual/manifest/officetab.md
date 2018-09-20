# <a name="officetab-element"></a><span data-ttu-id="d22b6-101">OfficeTab-Element</span><span class="sxs-lookup"><span data-stu-id="d22b6-101">OfficeTab element</span></span>

<span data-ttu-id="d22b6-p101">Definiert die Registerkarte des Menübands, auf der Ihr Add-In-Befehl angezeigt wird. Dies kann entweder die Standardregisterkarte sein (**Startseite**,  **Nachricht** oder  **Besprechung**), oder eine benutzerdefinierte Registerkarte, die vom Add-In definiert wird. Dieses Element ist erforderlich.</span><span class="sxs-lookup"><span data-stu-id="d22b6-p101">Defines the ribbon tab on which your add-in command appears. This can either be the default tab (either  **Home**,  **Message**, or  **Meeting**), or a custom tab defined by the add-in. This element is required.</span></span>

## <a name="child-elements"></a><span data-ttu-id="d22b6-105">Untergeordnete Elemente</span><span class="sxs-lookup"><span data-stu-id="d22b6-105">Child elements</span></span>

|  <span data-ttu-id="d22b6-106">Element</span><span class="sxs-lookup"><span data-stu-id="d22b6-106">Element</span></span> |  <span data-ttu-id="d22b6-107">Erforderlich</span><span class="sxs-lookup"><span data-stu-id="d22b6-107">Required</span></span>  |  <span data-ttu-id="d22b6-108">Beschreibung</span><span class="sxs-lookup"><span data-stu-id="d22b6-108">Description</span></span>  |
|:-----|:-----|:-----|
|  <span data-ttu-id="d22b6-109">Group</span><span class="sxs-lookup"><span data-stu-id="d22b6-109">Group</span></span>      | <span data-ttu-id="d22b6-110">Ja</span><span class="sxs-lookup"><span data-stu-id="d22b6-110">Yes</span></span> |  <span data-ttu-id="d22b6-p102">Definiert eine Gruppe von Befehlen. Sie können nur eine Gruppe pro Add-In zu der Standardregisterkarte hinzufügen.</span><span class="sxs-lookup"><span data-stu-id="d22b6-p102">Defines a group of commands. You can add only one group per add-in to the default tab.</span></span>  |

<span data-ttu-id="d22b6-p103">Nachfolgend finden Sie die gültigen `id`-Werte nach Host. **Fett formatierte** Werte werden sowohl auf dem Desktop als auch online unterstützt (z. B. Word 2016 für Windows und Word Online).</span><span class="sxs-lookup"><span data-stu-id="d22b6-p103">The following are valid tab `id` values by host. Values in **bold** are supported in both desktop and online (for example, Word 2016 for Windows and Word Online).</span></span> 

### <a name="outlook"></a><span data-ttu-id="d22b6-115">Outlook</span><span class="sxs-lookup"><span data-stu-id="d22b6-115">Outlook</span></span> 

- <span data-ttu-id="d22b6-116">**TabDefault**</span><span class="sxs-lookup"><span data-stu-id="d22b6-116">**TabDefault**</span></span>

### <a name="word"></a><span data-ttu-id="d22b6-117">Word</span><span class="sxs-lookup"><span data-stu-id="d22b6-117">Word</span></span>

- <span data-ttu-id="d22b6-118">**TabHome**</span><span class="sxs-lookup"><span data-stu-id="d22b6-118">**TabHome**</span></span>
- <span data-ttu-id="d22b6-119">**TabInsert**</span><span class="sxs-lookup"><span data-stu-id="d22b6-119">**TabInsert**</span></span>
- <span data-ttu-id="d22b6-120">TabWordDesign</span><span class="sxs-lookup"><span data-stu-id="d22b6-120">TabWordDesign</span></span>
- <span data-ttu-id="d22b6-121">**TabPageLayoutWord**</span><span class="sxs-lookup"><span data-stu-id="d22b6-121">**TabPageLayoutWord**</span></span>
- <span data-ttu-id="d22b6-122">TabReferences</span><span class="sxs-lookup"><span data-stu-id="d22b6-122">TabReferences</span></span>
- <span data-ttu-id="d22b6-123">TabMailings</span><span class="sxs-lookup"><span data-stu-id="d22b6-123">TabMailings</span></span>
- <span data-ttu-id="d22b6-124">TabReviewWord</span><span class="sxs-lookup"><span data-stu-id="d22b6-124">TabReviewWord</span></span>
- <span data-ttu-id="d22b6-125">**TabView**</span><span class="sxs-lookup"><span data-stu-id="d22b6-125">**TabView**</span></span>
- <span data-ttu-id="d22b6-126">TabDeveloper</span><span class="sxs-lookup"><span data-stu-id="d22b6-126">TabDeveloper</span></span>
- <span data-ttu-id="d22b6-127">TabAddIns</span><span class="sxs-lookup"><span data-stu-id="d22b6-127">TabAddIns</span></span>
- <span data-ttu-id="d22b6-128">TabBlogPost</span><span class="sxs-lookup"><span data-stu-id="d22b6-128">TabBlogPost</span></span>
- <span data-ttu-id="d22b6-129">TabBlogInsert</span><span class="sxs-lookup"><span data-stu-id="d22b6-129">TabBlogInsert</span></span>
- <span data-ttu-id="d22b6-130">TabPrintPreview</span><span class="sxs-lookup"><span data-stu-id="d22b6-130">TabPrintPreview</span></span>
- <span data-ttu-id="d22b6-131">TabOutlining</span><span class="sxs-lookup"><span data-stu-id="d22b6-131">TabOutlining</span></span>
- <span data-ttu-id="d22b6-132">TabConflicts</span><span class="sxs-lookup"><span data-stu-id="d22b6-132">TabConflicts</span></span>
- <span data-ttu-id="d22b6-133">TabBackgroundRemoval</span><span class="sxs-lookup"><span data-stu-id="d22b6-133">TabBackgroundRemoval</span></span>
- <span data-ttu-id="d22b6-134">TabBroadcastPresentation</span><span class="sxs-lookup"><span data-stu-id="d22b6-134">TabBroadcastPresentation</span></span>

### <a name="excel"></a><span data-ttu-id="d22b6-135">Excel</span><span class="sxs-lookup"><span data-stu-id="d22b6-135">Excel</span></span>

- <span data-ttu-id="d22b6-136">**TabHome**</span><span class="sxs-lookup"><span data-stu-id="d22b6-136">**TabHome**</span></span>
- <span data-ttu-id="d22b6-137">**TabInsert**</span><span class="sxs-lookup"><span data-stu-id="d22b6-137">**TabInsert**</span></span>
- <span data-ttu-id="d22b6-138">TabPageLayoutExcel</span><span class="sxs-lookup"><span data-stu-id="d22b6-138">TabPageLayoutExcel</span></span>
- <span data-ttu-id="d22b6-139">TabFormulas</span><span class="sxs-lookup"><span data-stu-id="d22b6-139">TabFormulas</span></span>
- <span data-ttu-id="d22b6-140">**TabData**</span><span class="sxs-lookup"><span data-stu-id="d22b6-140">**TabData**</span></span>
- <span data-ttu-id="d22b6-141">**TabReview**</span><span class="sxs-lookup"><span data-stu-id="d22b6-141">**TabReview**</span></span>
- <span data-ttu-id="d22b6-142">**TabView**</span><span class="sxs-lookup"><span data-stu-id="d22b6-142">**TabView**</span></span>
- <span data-ttu-id="d22b6-143">TabDeveloper</span><span class="sxs-lookup"><span data-stu-id="d22b6-143">TabDeveloper</span></span>
- <span data-ttu-id="d22b6-144">TabAddIns</span><span class="sxs-lookup"><span data-stu-id="d22b6-144">TabAddIns</span></span>
- <span data-ttu-id="d22b6-145">TabPrintPreview</span><span class="sxs-lookup"><span data-stu-id="d22b6-145">TabPrintPreview</span></span>
- <span data-ttu-id="d22b6-146">TabBackgroundRemoval</span><span class="sxs-lookup"><span data-stu-id="d22b6-146">TabBackgroundRemoval</span></span> 

### <a name="powerpoint"></a><span data-ttu-id="d22b6-147">PowerPoint</span><span class="sxs-lookup"><span data-stu-id="d22b6-147">PowerPoint</span></span>

- <span data-ttu-id="d22b6-148">**TabHome**</span><span class="sxs-lookup"><span data-stu-id="d22b6-148">**TabHome**</span></span>
- <span data-ttu-id="d22b6-149">**TabInsert**</span><span class="sxs-lookup"><span data-stu-id="d22b6-149">**TabInsert**</span></span>
- <span data-ttu-id="d22b6-150">**TabDesign**</span><span class="sxs-lookup"><span data-stu-id="d22b6-150">**TabDesign**</span></span>
- <span data-ttu-id="d22b6-151">**TabTransitions**</span><span class="sxs-lookup"><span data-stu-id="d22b6-151">**TabTransitions**</span></span>
- <span data-ttu-id="d22b6-152">**TabAnimations**</span><span class="sxs-lookup"><span data-stu-id="d22b6-152">**TabAnimations**</span></span>
- <span data-ttu-id="d22b6-153">TabSlideShow</span><span class="sxs-lookup"><span data-stu-id="d22b6-153">TabSlideShow</span></span>
- <span data-ttu-id="d22b6-154">TabReview</span><span class="sxs-lookup"><span data-stu-id="d22b6-154">TabReview</span></span>
- <span data-ttu-id="d22b6-155">**TabView**</span><span class="sxs-lookup"><span data-stu-id="d22b6-155">**TabView**</span></span>
- <span data-ttu-id="d22b6-156">TabDeveloper</span><span class="sxs-lookup"><span data-stu-id="d22b6-156">TabDeveloper</span></span>
- <span data-ttu-id="d22b6-157">TabAddIns</span><span class="sxs-lookup"><span data-stu-id="d22b6-157">TabAddIns</span></span>
- <span data-ttu-id="d22b6-158">TabPrintPreview</span><span class="sxs-lookup"><span data-stu-id="d22b6-158">TabPrintPreview</span></span>
- <span data-ttu-id="d22b6-159">TabMerge</span><span class="sxs-lookup"><span data-stu-id="d22b6-159">TabMerge</span></span>
- <span data-ttu-id="d22b6-160">TabGrayscale</span><span class="sxs-lookup"><span data-stu-id="d22b6-160">TabGrayscale</span></span>
- <span data-ttu-id="d22b6-161">TabBlackAndWhite</span><span class="sxs-lookup"><span data-stu-id="d22b6-161">TabBlackAndWhite</span></span>
- <span data-ttu-id="d22b6-162">TabBroadcastPresentation</span><span class="sxs-lookup"><span data-stu-id="d22b6-162">TabBroadcastPresentation</span></span>
- <span data-ttu-id="d22b6-163">TabSlideMaster</span><span class="sxs-lookup"><span data-stu-id="d22b6-163">TabSlideMaster</span></span>
- <span data-ttu-id="d22b6-164">TabHandoutMaster</span><span class="sxs-lookup"><span data-stu-id="d22b6-164">TabHandoutMaster</span></span>
- <span data-ttu-id="d22b6-165">TabNotesMaster</span><span class="sxs-lookup"><span data-stu-id="d22b6-165">TabNotesMaster</span></span>
- <span data-ttu-id="d22b6-166">TabBackgroundRemoval</span><span class="sxs-lookup"><span data-stu-id="d22b6-166">TabBackgroundRemoval</span></span>
- <span data-ttu-id="d22b6-167">TabSlideMasterHome</span><span class="sxs-lookup"><span data-stu-id="d22b6-167">TabSlideMasterHome</span></span>

### <a name="onenote"></a><span data-ttu-id="d22b6-168">OneNote</span><span class="sxs-lookup"><span data-stu-id="d22b6-168">OneNote</span></span>

- <span data-ttu-id="d22b6-169">**TabHome**</span><span class="sxs-lookup"><span data-stu-id="d22b6-169">**TabHome**</span></span>
- <span data-ttu-id="d22b6-170">**TabInsert**</span><span class="sxs-lookup"><span data-stu-id="d22b6-170">**TabInsert**</span></span>
- <span data-ttu-id="d22b6-171">**TabView**</span><span class="sxs-lookup"><span data-stu-id="d22b6-171">**TabView**</span></span>
- <span data-ttu-id="d22b6-172">TabDeveloper</span><span class="sxs-lookup"><span data-stu-id="d22b6-172">TabDeveloper</span></span>
- <span data-ttu-id="d22b6-173">TabAddIns</span><span class="sxs-lookup"><span data-stu-id="d22b6-173">TabAddIns</span></span>

## <a name="group"></a><span data-ttu-id="d22b6-174">Group</span><span class="sxs-lookup"><span data-stu-id="d22b6-174">Group</span></span>

<span data-ttu-id="d22b6-p104">Eine Gruppe von UI-Erweiterungspunkten auf einer Registerkarte. Eine Gruppe kann bis zu sechs Steuerelemente enthalten. Das **id**-Attribut ist erforderlich, und jede **id** muss im Manifest eindeutig sein. Die **id** ist eine Zeichenfolge mit maximal 125 Zeichen. Siehe [Group-Element](group.md).</span><span class="sxs-lookup"><span data-stu-id="d22b6-p104">A group of UI extension points in a tab. A group can have up to six controls. The  **id** attribute is required and each **id** must be unique within the manifest. The **id** is a string with a maximum of 125 characters. See [Group element](group.md).</span></span>

## <a name="officetab-example"></a><span data-ttu-id="d22b6-179">OfficeTab-Beispiel </span><span class="sxs-lookup"><span data-stu-id="d22b6-179">OfficeTab example</span></span>

```xml
<ExtensionPoint xsi:type="MessageReadCommandSurface">
  <OfficeTab id="TabDefault">
    <Group id="msgreadTabMessage.grp1">
        <!-- Group Definition -->
    </Group>
  </OfficeTab>
</ExtensionPoint>
```
