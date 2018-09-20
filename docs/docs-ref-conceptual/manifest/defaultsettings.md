# <a name="defaultsettings-element"></a><span data-ttu-id="0c212-101">DefaultSettings-Element</span><span class="sxs-lookup"><span data-stu-id="0c212-101">DefaultSettings element</span></span>

<span data-ttu-id="0c212-102">Gibt den Standardspeicherort für die Quell- und anderen Standardeinstellungen für den Inhalt oder task Pane-add-in.</span><span class="sxs-lookup"><span data-stu-id="0c212-102">Specifies the default source location and other default settings for your content or task pane add-in.</span></span>

<span data-ttu-id="0c212-103">**Add-In-Typ:** Inhalt, Aufgabenbereich</span><span class="sxs-lookup"><span data-stu-id="0c212-103">**Add-in type:** Content, Task pane</span></span>

## <a name="syntax"></a><span data-ttu-id="0c212-104">Syntax</span><span class="sxs-lookup"><span data-stu-id="0c212-104">Syntax</span></span>

```XML
<DefaultSettings>
  ...
</DefaultSettings>
```

## <a name="contained-in"></a><span data-ttu-id="0c212-105">Enthalten in</span><span class="sxs-lookup"><span data-stu-id="0c212-105">Contained in</span></span>

[<span data-ttu-id="0c212-106">OfficeApp</span><span class="sxs-lookup"><span data-stu-id="0c212-106">OfficeApp</span></span>](officeapp.md)

## <a name="can-contain"></a><span data-ttu-id="0c212-107">Kann enthalten</span><span class="sxs-lookup"><span data-stu-id="0c212-107">Can contain</span></span>

|<span data-ttu-id="0c212-108">**Element**</span><span class="sxs-lookup"><span data-stu-id="0c212-108">**Element**</span></span>|<span data-ttu-id="0c212-109">**Inhalt**</span><span class="sxs-lookup"><span data-stu-id="0c212-109">**Content**</span></span>|<span data-ttu-id="0c212-110">**E-Mail**</span><span class="sxs-lookup"><span data-stu-id="0c212-110">**Mail**</span></span>|<span data-ttu-id="0c212-111">**TaskPane**</span><span class="sxs-lookup"><span data-stu-id="0c212-111">**TaskPane**</span></span>|
|:-----|:-----|:-----|:-----|
|[<span data-ttu-id="0c212-112">SourceLocation</span><span class="sxs-lookup"><span data-stu-id="0c212-112">SourceLocation</span></span>](sourcelocation.md)|<span data-ttu-id="0c212-113">x</span><span class="sxs-lookup"><span data-stu-id="0c212-113">x</span></span>||<span data-ttu-id="0c212-114">x</span><span class="sxs-lookup"><span data-stu-id="0c212-114">x</span></span>|
|[<span data-ttu-id="0c212-115">RequestedWidth</span><span class="sxs-lookup"><span data-stu-id="0c212-115">RequestedWidth</span></span>](requestedwidth.md)|<span data-ttu-id="0c212-116">x</span><span class="sxs-lookup"><span data-stu-id="0c212-116">x</span></span>|||
|[<span data-ttu-id="0c212-117">RequestedHeight</span><span class="sxs-lookup"><span data-stu-id="0c212-117">RequestedHeight</span></span>](requestedheight.md)|<span data-ttu-id="0c212-118">x</span><span class="sxs-lookup"><span data-stu-id="0c212-118">x</span></span>|||

## <a name="remarks"></a><span data-ttu-id="0c212-119">Bemerkungen</span><span class="sxs-lookup"><span data-stu-id="0c212-119">Remarks</span></span>

<span data-ttu-id="0c212-120">Der Quellspeicherort und andere Einstellungen im **DefaultSettings**-Element gelten nur für Inhalts- und Aufgabenbereich-Add-Ins. Bei E-Mail-Add-Ins geben Sie die Standardspeicherorte für Quelldateien und andere Standardeinstellungen im [FormSettings](formsettings.md)-Element an.</span><span class="sxs-lookup"><span data-stu-id="0c212-120">The source location and other settings in the  **DefaultSettings** element apply only to content and task pane add-ins. For mail add-ins, you specify the default locations for source files and other default settings in the [FormSettings](formsettings.md) element.</span></span>

