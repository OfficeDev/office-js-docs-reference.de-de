# <a name="sourcelocation-element"></a><span data-ttu-id="69d94-101">SourceLocation-Element</span><span class="sxs-lookup"><span data-stu-id="69d94-101">SourceLocation element</span></span>

<span data-ttu-id="69d94-p101">Gibt die Quelldateispeicherorte für Ihr Office-Add-In als URL an, die zwischen 1 und 2018 Zeichen lang ist. Der Quellspeicherort muss eine HTTPS-Adresse, kein Dateipfad sein.</span><span class="sxs-lookup"><span data-stu-id="69d94-p101">Specifies the source file location(s) for your Office Add-in as a URL between 1 and 2018 characters long. The source location must be an HTTPS address, not a file path.</span></span>

<span data-ttu-id="69d94-104">**Add-In-Typ:** Inhalt, Aufgabenbereich, E-Mail</span><span class="sxs-lookup"><span data-stu-id="69d94-104">**Add-in type:** Content, Task pane, Mail</span></span>

## <a name="syntax"></a><span data-ttu-id="69d94-105">Syntax</span><span class="sxs-lookup"><span data-stu-id="69d94-105">Syntax</span></span>

```XML
<SourceLocation DefaultValue="string" />
```

## <a name="contained-in"></a><span data-ttu-id="69d94-106">Enthalten in</span><span class="sxs-lookup"><span data-stu-id="69d94-106">Contained in</span></span>

- <span data-ttu-id="69d94-107">[DefaultSettings](defaultsettings.md) (Inhalts- und Aufgabenbereich-Add-Ins)</span><span class="sxs-lookup"><span data-stu-id="69d94-107">[DefaultSettings](defaultsettings.md) (Content and task pane add-ins)</span></span>
- <span data-ttu-id="69d94-108">[FormSettings](formsettings.md) (E-Mail-Add-Ins)</span><span class="sxs-lookup"><span data-stu-id="69d94-108">[FormSettings](formsettings.md) (Mail add-ins)</span></span>
- <span data-ttu-id="69d94-109">[ExtensionPoint](extensionpoint.md) (E-Mail-Kontext-Add-Ins)</span><span class="sxs-lookup"><span data-stu-id="69d94-109">[ExtensionPoint](extensionpoint.md) (Contextual mail add-ins)</span></span>

## <a name="can-contain"></a><span data-ttu-id="69d94-110">Kann enthalten</span><span class="sxs-lookup"><span data-stu-id="69d94-110">Can contain</span></span>

[<span data-ttu-id="69d94-111">Override</span><span class="sxs-lookup"><span data-stu-id="69d94-111">Override</span></span>](override.md)

## <a name="attributes"></a><span data-ttu-id="69d94-112">Attribute</span><span class="sxs-lookup"><span data-stu-id="69d94-112">Attributes</span></span>

|<span data-ttu-id="69d94-113">**Attribut**</span><span class="sxs-lookup"><span data-stu-id="69d94-113">**Attribute**</span></span>|<span data-ttu-id="69d94-114">**Typ**</span><span class="sxs-lookup"><span data-stu-id="69d94-114">**Type**</span></span>|<span data-ttu-id="69d94-115">**Erforderlich**</span><span class="sxs-lookup"><span data-stu-id="69d94-115">**Required**</span></span>|<span data-ttu-id="69d94-116">**Beschreibung**</span><span class="sxs-lookup"><span data-stu-id="69d94-116">**Description**</span></span>|
|:-----|:-----|:-----|:-----|
|<span data-ttu-id="69d94-117">DefaultValue</span><span class="sxs-lookup"><span data-stu-id="69d94-117">DefaultValue</span></span>|<span data-ttu-id="69d94-118">URL</span><span class="sxs-lookup"><span data-stu-id="69d94-118">URL</span></span>|<span data-ttu-id="69d94-119">erforderlich</span><span class="sxs-lookup"><span data-stu-id="69d94-119">required</span></span>|<span data-ttu-id="69d94-120">Gibt den Standardwert für diese Einstellung für das im [DefaultLocale](defaultlocale.md)-Element angegebene Gebietsschema an.</span><span class="sxs-lookup"><span data-stu-id="69d94-120">Specifies the default value for this setting for the locale specified in the [DefaultLocale](defaultlocale.md) element.</span></span>|
