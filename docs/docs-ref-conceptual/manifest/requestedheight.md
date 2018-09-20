# <a name="requestedheight-element"></a><span data-ttu-id="da221-101">RequestedHeight-Element</span><span class="sxs-lookup"><span data-stu-id="da221-101">RequestedHeight element</span></span>

<span data-ttu-id="da221-102">Gibt die ursprüngliche Höhe (in Pixel) eines Content-add-in oder e-Mail-add-in.</span><span class="sxs-lookup"><span data-stu-id="da221-102">Specifies the initial height (in pixels) of a content add-in or mail add-in.</span></span> 

<span data-ttu-id="da221-103">**-Add-in-Typ:** Inhalte, e-Mail-Nachrichten</span><span class="sxs-lookup"><span data-stu-id="da221-103">**Add-in type:** Content, Mail</span></span>

## <a name="syntax"></a><span data-ttu-id="da221-104">Syntax</span><span class="sxs-lookup"><span data-stu-id="da221-104">Syntax</span></span>

```XML
<RequestedHeight>integer</RequestedHeight>
```

## <a name="contained-in"></a><span data-ttu-id="da221-105">Enthalten in</span><span class="sxs-lookup"><span data-stu-id="da221-105">Contained in</span></span>

- <span data-ttu-id="da221-106">[DefaultSettings](defaultsettings.md) (Content-add-ins) mit einem Wert, der zwischen 32 und 1000 sein kann</span><span class="sxs-lookup"><span data-stu-id="da221-106">[DefaultSettings](defaultsettings.md) (Content add-ins) with a value that can be between 32 and 1000</span></span>
- <span data-ttu-id="da221-107">[DesktopSettings](desktopsettings.md) und [TabletSettings](tabletsettings.md) (Mail-add-ins) mit einem Wert, der zwischen 32 und 450 werden kann</span><span class="sxs-lookup"><span data-stu-id="da221-107">[DesktopSettings](desktopsettings.md) and [TabletSettings](tabletsettings.md) (Mail add-ins) with a value that can be between 32 and 450</span></span>
- <span data-ttu-id="da221-108">[ExtensionPoint](extensionpoint.md) (Kontextbezogene Mail-add-ins) mit einem Wert, der zwischen 140 und 450 für die **DetectedEntity** Erweiterungspunkt sowie zwischen 32 und 450 für die **CustomPane** Erweiterungspunkt werden kann</span><span class="sxs-lookup"><span data-stu-id="da221-108">[ExtensionPoint](extensionpoint.md) (Contextual mail add-ins) with a value that can be between 140 and 450 for the **DetectedEntity** extension point and between 32 and 450 for the **CustomPane** extension point</span></span>