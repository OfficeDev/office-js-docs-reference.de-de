# <a name="metadata-element"></a><span data-ttu-id="44ea0-101">Metadata-element</span><span class="sxs-lookup"><span data-stu-id="44ea0-101">Metadata element</span></span>

<span data-ttu-id="44ea0-102">Definiert die Metadateneinstellungen, die durch eine benutzerdefinierte Funktion in Excel verwendet.</span><span class="sxs-lookup"><span data-stu-id="44ea0-102">Defines the metadata settings used by a custom function in Excel.</span></span>

## <a name="attributes"></a><span data-ttu-id="44ea0-103">Attribute</span><span class="sxs-lookup"><span data-stu-id="44ea0-103">Attributes</span></span>

<span data-ttu-id="44ea0-104">Keine</span><span class="sxs-lookup"><span data-stu-id="44ea0-104">None</span></span>

## <a name="child-elements"></a><span data-ttu-id="44ea0-105">Untergeordnete Elemente</span><span class="sxs-lookup"><span data-stu-id="44ea0-105">Child elements</span></span>

|  <span data-ttu-id="44ea0-106">Element</span><span class="sxs-lookup"><span data-stu-id="44ea0-106">Element</span></span>  |  <span data-ttu-id="44ea0-107">Erforderlich</span><span class="sxs-lookup"><span data-stu-id="44ea0-107">Required</span></span>  |  <span data-ttu-id="44ea0-108">Beschreibung</span><span class="sxs-lookup"><span data-stu-id="44ea0-108">Description</span></span>  |
|:-----|:-----|:-----|
|  [<span data-ttu-id="44ea0-109">SourceLocation</span><span class="sxs-lookup"><span data-stu-id="44ea0-109">SourceLocation</span></span>](customfunctionssourcelocation.md)  |  <span data-ttu-id="44ea0-110">Ja</span><span class="sxs-lookup"><span data-stu-id="44ea0-110">Yes</span></span>  | <span data-ttu-id="44ea0-111">Eine Zeichenfolge mit der Ressourcen-Id der JSON-Datei, die von benutzerdefinierten Funktionen.</span><span class="sxs-lookup"><span data-stu-id="44ea0-111">String with the resource id of the JSON file used by custom functions.</span></span> |

## <a name="example"></a><span data-ttu-id="44ea0-112">Beispiel</span><span class="sxs-lookup"><span data-stu-id="44ea0-112">Example</span></span>

```xml
<Metadata>
    <SourceLocation resid="JSON-URL" />
</Metadata>
```
