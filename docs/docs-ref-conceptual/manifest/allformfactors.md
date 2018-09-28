# <a name="allformfactors-element"></a><span data-ttu-id="d0851-101">AllFormFactors-Element</span><span class="sxs-lookup"><span data-stu-id="d0851-101">AllFormFactors element</span></span>

<span data-ttu-id="d0851-102">Gibt die Einstellungen für ein Add-In für alle Formfaktoren an.</span><span class="sxs-lookup"><span data-stu-id="d0851-102">Specifies the settings for an add-in for all form factors.</span></span> <span data-ttu-id="d0851-103">Derzeit ist das einzige Feature **AllFormFactors** mit benutzerdefinierten Funktionen.</span><span class="sxs-lookup"><span data-stu-id="d0851-103">Currently, the only feature using **AllFormFactors** is custom functions.</span></span> <span data-ttu-id="d0851-104">**AllFormFactors** ist ein erforderliches Element aus, wenn Sie benutzerdefinierte Funktionen verwenden.</span><span class="sxs-lookup"><span data-stu-id="d0851-104">**AllFormFactors** is a required element when using custom functions.</span></span>

## <a name="child-elements"></a><span data-ttu-id="d0851-105">Untergeordnete Elemente</span><span class="sxs-lookup"><span data-stu-id="d0851-105">Child elements</span></span>

|  <span data-ttu-id="d0851-106">Element</span><span class="sxs-lookup"><span data-stu-id="d0851-106">Element</span></span> |  <span data-ttu-id="d0851-107">Erforderlich</span><span class="sxs-lookup"><span data-stu-id="d0851-107">Required</span></span>  |  <span data-ttu-id="d0851-108">Beschreibung</span><span class="sxs-lookup"><span data-stu-id="d0851-108">Description</span></span>  |
|:-----|:-----|:-----|
|  [<span data-ttu-id="d0851-109">ExtensionPoint</span><span class="sxs-lookup"><span data-stu-id="d0851-109">ExtensionPoint</span></span>](extensionpoint.md) |  <span data-ttu-id="d0851-110">Ja</span><span class="sxs-lookup"><span data-stu-id="d0851-110">Yes</span></span> |  <span data-ttu-id="d0851-111">Definiert, wo ein Add-In Funktionen verfügbar macht.</span><span class="sxs-lookup"><span data-stu-id="d0851-111">Defines where an add-in exposes functionality.</span></span> |

## <a name="allformfactors-example"></a><span data-ttu-id="d0851-112">AllFormFactors-Beispiel</span><span class="sxs-lookup"><span data-stu-id="d0851-112">AllFormFactors example</span></span>

```xml
<Hosts>
    <Host xsi:type="Workbook">
        <AllFormFactors>
            <ExtensionPoint xsi:type="CustomFunctions">
                    <!-- Information on this extension point -->
            </ExtensionPoint>
        </AllFormFactors>
    </Host>
</Hosts>
```
