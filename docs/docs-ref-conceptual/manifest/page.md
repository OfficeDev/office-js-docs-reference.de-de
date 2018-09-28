# <a name="page-element"></a><span data-ttu-id="a5091-101">Page-Element</span><span class="sxs-lookup"><span data-stu-id="a5091-101">Page element</span></span>

<span data-ttu-id="a5091-102">Definiert HTML-Seiteneinstellungen, die von einer benutzerdefinierten Funktion in Excel verwendet werden.</span><span class="sxs-lookup"><span data-stu-id="a5091-102">Defines HTML page settings used by a custom function in Excel.</span></span>

## <a name="attributes"></a><span data-ttu-id="a5091-103">Attribute</span><span class="sxs-lookup"><span data-stu-id="a5091-103">Attributes</span></span>

<span data-ttu-id="a5091-104">Keine</span><span class="sxs-lookup"><span data-stu-id="a5091-104">None</span></span>

## <a name="child-elements"></a><span data-ttu-id="a5091-105">Untergeordnete Elemente</span><span class="sxs-lookup"><span data-stu-id="a5091-105">Child elements</span></span>

|  <span data-ttu-id="a5091-106">Element</span><span class="sxs-lookup"><span data-stu-id="a5091-106">Element</span></span>  |  <span data-ttu-id="a5091-107">Erforderlich</span><span class="sxs-lookup"><span data-stu-id="a5091-107">Required</span></span>  |  <span data-ttu-id="a5091-108">Beschreibung</span><span class="sxs-lookup"><span data-stu-id="a5091-108">Description</span></span>  |
|:-----|:-----|:-----|
|  [<span data-ttu-id="a5091-109">SourceLocation</span><span class="sxs-lookup"><span data-stu-id="a5091-109">SourceLocation</span></span>](customfunctionssourcelocation.md)  |  <span data-ttu-id="a5091-110">Ja</span><span class="sxs-lookup"><span data-stu-id="a5091-110">Yes</span></span>  | <span data-ttu-id="a5091-111">Zeichenfolge mit Ressourcen-ID der JavaScript-Datei, die von benutzerdefinierten Funktionen verwendet wird.</span><span class="sxs-lookup"><span data-stu-id="a5091-111">String with the resource id of the HTML file used by custom functions.</span></span> |

## <a name="example"></a><span data-ttu-id="a5091-112">Beispiel</span><span class="sxs-lookup"><span data-stu-id="a5091-112">Example</span></span>

```xml
<Page>
    <SourceLocation resid="pageURL"/>
</Page>
```
