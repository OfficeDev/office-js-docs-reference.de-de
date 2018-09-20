# <a name="sourcelocation-element"></a><span data-ttu-id="ee969-101">SourceLocation-Element</span><span class="sxs-lookup"><span data-stu-id="ee969-101">SourceLocation element</span></span>

<span data-ttu-id="ee969-102">Definiert den Speicherort einer Ressource, die für Script- oder Page-Elemente erforderlich ist, die von benutzerdefinierten Funktionen in Excel verwendet werden.</span><span class="sxs-lookup"><span data-stu-id="ee969-102">Defines the location of a resource needed by the Script or Page elements used by custom functions in Excel.</span></span>

## <a name="attributes"></a><span data-ttu-id="ee969-103">Attribute</span><span class="sxs-lookup"><span data-stu-id="ee969-103">Attributes</span></span>

| <span data-ttu-id="ee969-104">**Attribut**</span><span class="sxs-lookup"><span data-stu-id="ee969-104">**Attribute**</span></span> | <span data-ttu-id="ee969-105">**Erforderlich**</span><span class="sxs-lookup"><span data-stu-id="ee969-105">**Required**</span></span> | <span data-ttu-id="ee969-106">**Beschreibung**</span><span class="sxs-lookup"><span data-stu-id="ee969-106">**Description**</span></span>                                                                      |
|---------------|--------------|--------------------------------------------------------------------------------------|
| <span data-ttu-id="ee969-107">resid</span><span class="sxs-lookup"><span data-stu-id="ee969-107">resid</span></span>         | <span data-ttu-id="ee969-108">Ja</span><span class="sxs-lookup"><span data-stu-id="ee969-108">Yes</span></span>          | <span data-ttu-id="ee969-109">Der Name einer URL-Ressource, die im Abschnitt &lt;Ressourcen&gt; des Manifests definiert ist.</span><span class="sxs-lookup"><span data-stu-id="ee969-109">The name of a URL resource defined in the &lt;Resources&gt; section of the manifest.</span></span> |

## <a name="child-elements"></a><span data-ttu-id="ee969-110">Untergeordnete Elemente</span><span class="sxs-lookup"><span data-stu-id="ee969-110">Child elements</span></span>

<span data-ttu-id="ee969-111">Keine</span><span class="sxs-lookup"><span data-stu-id="ee969-111">None</span></span>

## <a name="example"></a><span data-ttu-id="ee969-112">Beispiel</span><span class="sxs-lookup"><span data-stu-id="ee969-112">Example</span></span>

```xml
<SourceLocation resid="pageURL"/>
```