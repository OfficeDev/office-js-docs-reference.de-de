# <a name="script-element"></a><span data-ttu-id="15ff7-101">Script-Element</span><span class="sxs-lookup"><span data-stu-id="15ff7-101">Script element</span></span>

<span data-ttu-id="15ff7-102">Definiert Skripteinstellungen, die von einer benutzerdefinierten Funktion in Excel verwendet werden.</span><span class="sxs-lookup"><span data-stu-id="15ff7-102">Defines script settings used by a custom function in Excel.</span></span>

## <a name="attributes"></a><span data-ttu-id="15ff7-103">Attribute</span><span class="sxs-lookup"><span data-stu-id="15ff7-103">Attributes</span></span>

<span data-ttu-id="15ff7-104">Keine</span><span class="sxs-lookup"><span data-stu-id="15ff7-104">None</span></span>

## <a name="child-elements"></a><span data-ttu-id="15ff7-105">Untergeordnete Elemente</span><span class="sxs-lookup"><span data-stu-id="15ff7-105">Child elements</span></span>

|<span data-ttu-id="15ff7-106">Elemente</span><span class="sxs-lookup"><span data-stu-id="15ff7-106">Elements</span></span>  |  <span data-ttu-id="15ff7-107">Erforderlich</span><span class="sxs-lookup"><span data-stu-id="15ff7-107">Required</span></span>  |  <span data-ttu-id="15ff7-108">Beschreibung</span><span class="sxs-lookup"><span data-stu-id="15ff7-108">Description</span></span>  |
|:-----|:-----|:-----|
|  [<span data-ttu-id="15ff7-109">SourceLocation</span><span class="sxs-lookup"><span data-stu-id="15ff7-109">SourceLocation</span></span>](customfunctionssourcelocation.md)  |  <span data-ttu-id="15ff7-110">Ja</span><span class="sxs-lookup"><span data-stu-id="15ff7-110">Yes</span></span>  | <span data-ttu-id="15ff7-111">Zeichenfolge mit Ressourcen-ID der JavaScript-Datei, die von benutzerdefinierten Funktionen verwendet wird.</span><span class="sxs-lookup"><span data-stu-id="15ff7-111">String with resource id of the JavaScript file used by custom functions.</span></span>|

## <a name="example"></a><span data-ttu-id="15ff7-112">Beispiel</span><span class="sxs-lookup"><span data-stu-id="15ff7-112">Example</span></span>

```xml
<Script>
    <SourceLocation resid="scriptURL" />
    <!-- The Script element is not used in the Developer Preview. -->
</Script>
```
