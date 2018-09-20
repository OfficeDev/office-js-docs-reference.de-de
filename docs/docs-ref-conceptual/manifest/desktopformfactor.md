# <a name="desktopformfactor-element"></a><span data-ttu-id="cd3f5-101">DesktopFormFactor-Element</span><span class="sxs-lookup"><span data-stu-id="cd3f5-101">DesktopFormFactor element</span></span>

<span data-ttu-id="cd3f5-p101">Gibt die Einstellungen für ein Add-In für den  Desktopformfaktor an. Der Desktopformfaktor umfasst Office für Windows, Office für Mac und Office Online. Es enthält alle Add-In-Informationen für den Desktopformfaktor mit Ausnahme des **Resources**-Knotens.</span><span class="sxs-lookup"><span data-stu-id="cd3f5-p101">Specifies the settings for an add-in for the desktop form factor. The desktop form factor includes Office for Windows, Office for Mac, and Office Online. It contains all the add-in information for the desktop form factor except for the  **Resources** node.</span></span>

<span data-ttu-id="cd3f5-p102">Jede DesktopFormFactor-Definition enthält das **FunctionFile**-Element und mindestens ein **ExtensionPoint**-Element. Weitere Informationen finden Sie unter [FunctionFile-Element](functionfile.md) und [ExtensionPoint-Element](extensionpoint.md).</span><span class="sxs-lookup"><span data-stu-id="cd3f5-p102">Each DesktopFormFactor definition contains the  **FunctionFile** element and one or more **ExtensionPoint** elements. For more information, see [FunctionFile element](functionfile.md) and [ExtensionPoint element](extensionpoint.md).</span></span> 

## <a name="child-elements"></a><span data-ttu-id="cd3f5-107">Untergeordnete Elemente</span><span class="sxs-lookup"><span data-stu-id="cd3f5-107">Child elements</span></span>

| <span data-ttu-id="cd3f5-108">Element</span><span class="sxs-lookup"><span data-stu-id="cd3f5-108">Element</span></span>                               | <span data-ttu-id="cd3f5-109">Erforderlich</span><span class="sxs-lookup"><span data-stu-id="cd3f5-109">Required</span></span> | <span data-ttu-id="cd3f5-110">Beschreibung</span><span class="sxs-lookup"><span data-stu-id="cd3f5-110">Description</span></span>  |
|:--------------------------------------|:--------:|:-------------|
| [<span data-ttu-id="cd3f5-111">ExtensionPoint</span><span class="sxs-lookup"><span data-stu-id="cd3f5-111">ExtensionPoint</span></span>](extensionpoint.md) | <span data-ttu-id="cd3f5-112">Ja</span><span class="sxs-lookup"><span data-stu-id="cd3f5-112">Yes</span></span>      | <span data-ttu-id="cd3f5-113">Definiert, wo ein Add-In Funktionen verfügbar macht.</span><span class="sxs-lookup"><span data-stu-id="cd3f5-113">Defines where an add-in exposes functionality.</span></span> |
| [<span data-ttu-id="cd3f5-114">FunctionFile</span><span class="sxs-lookup"><span data-stu-id="cd3f5-114">FunctionFile</span></span>](functionfile.md)     | <span data-ttu-id="cd3f5-115">Ja</span><span class="sxs-lookup"><span data-stu-id="cd3f5-115">Yes</span></span>      | <span data-ttu-id="cd3f5-116">Eine URL zu einer Datei, die JavaScript-Funktionen enthält.</span><span class="sxs-lookup"><span data-stu-id="cd3f5-116">A URL to a file that contains JavaScript functions.</span></span>|
| [<span data-ttu-id="cd3f5-117">GetStarted</span><span class="sxs-lookup"><span data-stu-id="cd3f5-117">GetStarted</span></span>](getstarted.md)         | <span data-ttu-id="cd3f5-118">Nein</span><span class="sxs-lookup"><span data-stu-id="cd3f5-118">No</span></span>       | <span data-ttu-id="cd3f5-119">Definiert die Beschriftung, die angezeigt wird, wenn Sie das Add-In in Word-, Excel- oder PowerPoint-Hosts installieren.</span><span class="sxs-lookup"><span data-stu-id="cd3f5-119">Defines the callout that appears when installing the add-in in Word, Excel, or PowerPoint hosts.</span></span> |

## <a name="desktopformfactor-example"></a><span data-ttu-id="cd3f5-120">DesktopFormFactor-Beispiel</span><span class="sxs-lookup"><span data-stu-id="cd3f5-120">DesktopFormFactor example</span></span>

```xml
...
<Hosts>
  <Host xsi:type="Presentation">
    <DesktopFormFactor>
      <FunctionFile resid="residDesktopFuncUrl" />
      <GetStarted>
        <!-- GetStarted callout -->
      </GetStarted>
      <ExtensionPoint xsi:type="PrimaryCommandSurface">
        <!-- information on this extension point -->
      </ExtensionPoint> 
      <!-- possibly more ExtensionPoint elements -->
    </DesktopFormFactor>
  </Host>
</Hosts>
...
```
