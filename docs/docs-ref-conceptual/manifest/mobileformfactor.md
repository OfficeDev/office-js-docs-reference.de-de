# <a name="mobileformfactor-element"></a><span data-ttu-id="cc1ee-101">MobileFormFactor-Element</span><span class="sxs-lookup"><span data-stu-id="cc1ee-101">MobileFormFactor element</span></span>

<span data-ttu-id="cc1ee-p101">Gibt die Einstellungen für ein Add-In für den mobilen Formfaktor an. Es enthält alle Add-In-Informationen für den mobilen Formfaktor mit Ausnahme des **Resources**-Knotens.</span><span class="sxs-lookup"><span data-stu-id="cc1ee-p101">Specifies the settings for an add-in for the mobile form factor. It contains all the add-in information for the mobile form factor except for the **Resources** node.</span></span>

<span data-ttu-id="cc1ee-p102">Jede **MobileFormFactor**-Definition enthält das **FunctionFile**-Element und mindestens ein **ExtensionPoint**-Elemente. Weitere Informationen finden Sie unter [FunctionFile-Element](functionfile.md) und [ExtensionPoint-Element](extensionpoint.md)</span><span class="sxs-lookup"><span data-stu-id="cc1ee-p102">Each **MobileFormFactor** definition contains the  **FunctionFile** element and one or more **ExtensionPoint** elements. For more information, see [FunctionFile element](functionfile.md) and [ExtensionPoint element](extensionpoint.md).</span></span>

<span data-ttu-id="cc1ee-p103">Das **MobileFormFactor**-Element ist im VersionOverrides Schema 1.1 definiert. Das enthaltende [VersionOverrides](versionoverrides.md)-Element muss einen `xsi:type`-Attributwert von `VersionOverridesV1_1` aufweisen.</span><span class="sxs-lookup"><span data-stu-id="cc1ee-p103">The **MobileFormFactor** element is defined in VersionOverrides schema 1.1. The containing [VersionOverrides](versionoverrides.md) element must have an `xsi:type` attribute value of `VersionOverridesV1_1`.</span></span>

## <a name="child-elements"></a><span data-ttu-id="cc1ee-108">Untergeordnete Elemente</span><span class="sxs-lookup"><span data-stu-id="cc1ee-108">Child elements</span></span>

| <span data-ttu-id="cc1ee-109">Element</span><span class="sxs-lookup"><span data-stu-id="cc1ee-109">Element</span></span>                               | <span data-ttu-id="cc1ee-110">Erforderlich</span><span class="sxs-lookup"><span data-stu-id="cc1ee-110">Required</span></span> | <span data-ttu-id="cc1ee-111">Beschreibung</span><span class="sxs-lookup"><span data-stu-id="cc1ee-111">Description</span></span>  |
|:--------------------------------------|:--------:|:-------------|
| [<span data-ttu-id="cc1ee-112">ExtensionPoint</span><span class="sxs-lookup"><span data-stu-id="cc1ee-112">ExtensionPoint</span></span>](extensionpoint.md) | <span data-ttu-id="cc1ee-113">Ja</span><span class="sxs-lookup"><span data-stu-id="cc1ee-113">Yes</span></span>      | <span data-ttu-id="cc1ee-114">Definiert, wo ein Add-In Funktionen verfügbar macht.</span><span class="sxs-lookup"><span data-stu-id="cc1ee-114">Defines where an add-in exposes functionality.</span></span> |
| [<span data-ttu-id="cc1ee-115">FunctionFile</span><span class="sxs-lookup"><span data-stu-id="cc1ee-115">FunctionFile</span></span>](functionfile.md)     | <span data-ttu-id="cc1ee-116">Ja</span><span class="sxs-lookup"><span data-stu-id="cc1ee-116">Yes</span></span>      | <span data-ttu-id="cc1ee-117">Eine URL zu einer Datei, die JavaScript-Funktionen enthält.</span><span class="sxs-lookup"><span data-stu-id="cc1ee-117">A URL to a file that contains JavaScript functions.</span></span>|

## <a name="mobileformfactor-example"></a><span data-ttu-id="cc1ee-118">MobileFormFactor-Beispiel</span><span class="sxs-lookup"><span data-stu-id="cc1ee-118">MobileFormFactor example</span></span>

```xml
...
<Hosts>
  <Host xsi:type="MailHost">
    ...
    <MobileFormFactor>
      <FunctionFile resid="residUILessFunctionFileUrl" />
      <ExtensionPoint xsi:type="MobileMessageReadCommandSurface">
        <!-- information on this extension point -->
      </ExtensionPoint> 
      <!-- possibly more ExtensionPoint elements -->
    </MobileFormFactor>
  </Host>
</Hosts>
...
```
