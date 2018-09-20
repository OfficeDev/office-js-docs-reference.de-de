# <a name="page-element"></a><span data-ttu-id="e81c7-101">Page-Element</span><span class="sxs-lookup"><span data-stu-id="e81c7-101">Page element</span></span>

<span data-ttu-id="e81c7-102">Definiert HTML-Seiteneinstellungen, die von einer benutzerdefinierten Funktion in Excel verwendet werden.</span><span class="sxs-lookup"><span data-stu-id="e81c7-102">Defines HTML page settings used by a custom function in Excel.</span></span>

## <a name="attributes"></a><span data-ttu-id="e81c7-103">Attribute</span><span class="sxs-lookup"><span data-stu-id="e81c7-103">Attributes</span></span>

<span data-ttu-id="e81c7-104">Keine</span><span class="sxs-lookup"><span data-stu-id="e81c7-104">None</span></span>

## <a name="child-elements"></a><span data-ttu-id="e81c7-105">Untergeordnete Elemente</span><span class="sxs-lookup"><span data-stu-id="e81c7-105">Child elements</span></span>

|  <span data-ttu-id="e81c7-106">Element</span><span class="sxs-lookup"><span data-stu-id="e81c7-106">Element</span></span>  |  <span data-ttu-id="e81c7-107">Erforderlich</span><span class="sxs-lookup"><span data-stu-id="e81c7-107">Required</span></span>  |  <span data-ttu-id="e81c7-108">Beschreibung</span><span class="sxs-lookup"><span data-stu-id="e81c7-108">Description</span></span>  |
|:-----|:-----|:-----|
|  [<span data-ttu-id="e81c7-109">SourceLocation</span><span class="sxs-lookup"><span data-stu-id="e81c7-109">SourceLocation</span></span>](customfunctionssourcelocation.md)  |  <span data-ttu-id="e81c7-110">Ja</span><span class="sxs-lookup"><span data-stu-id="e81c7-110">Yes</span></span>  | <span data-ttu-id="e81c7-111">Zeichenfolge mit Ressourcen-ID der JavaScript-Datei, die von benutzerdefinierten Funktionen verwendet wird.</span><span class="sxs-lookup"><span data-stu-id="e81c7-111">String with the resource id of the HTML file used by custom functions.</span></span> |

## <a name="example"></a><span data-ttu-id="e81c7-112">Beispiel</span><span class="sxs-lookup"><span data-stu-id="e81c7-112">Example</span></span>

```xml
<Page>
    <SourceLocation resid="pageURL"/>
</Page>
```
