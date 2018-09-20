# <a name="metadata-element"></a><span data-ttu-id="1b6dd-101">Metadata-element</span><span class="sxs-lookup"><span data-stu-id="1b6dd-101">Metadata element</span></span>

<span data-ttu-id="1b6dd-102">Definiert die Metadateneinstellungen, die durch eine benutzerdefinierte Funktion in Excel verwendet.</span><span class="sxs-lookup"><span data-stu-id="1b6dd-102">Defines the metadata settings used by a custom function in Excel.</span></span>

## <a name="attributes"></a><span data-ttu-id="1b6dd-103">Attribute</span><span class="sxs-lookup"><span data-stu-id="1b6dd-103">Attributes</span></span>

<span data-ttu-id="1b6dd-104">Keine</span><span class="sxs-lookup"><span data-stu-id="1b6dd-104">None</span></span>

## <a name="child-elements"></a><span data-ttu-id="1b6dd-105">Untergeordnete Elemente</span><span class="sxs-lookup"><span data-stu-id="1b6dd-105">Child elements</span></span>

|  <span data-ttu-id="1b6dd-106">Element</span><span class="sxs-lookup"><span data-stu-id="1b6dd-106">Element</span></span>  |  <span data-ttu-id="1b6dd-107">Erforderlich</span><span class="sxs-lookup"><span data-stu-id="1b6dd-107">Required</span></span>  |  <span data-ttu-id="1b6dd-108">Beschreibung</span><span class="sxs-lookup"><span data-stu-id="1b6dd-108">Description</span></span>  |
|:-----|:-----|:-----|
|  [<span data-ttu-id="1b6dd-109">SourceLocation</span><span class="sxs-lookup"><span data-stu-id="1b6dd-109">SourceLocation</span></span>](customfunctionssourcelocation.md)  |  <span data-ttu-id="1b6dd-110">Ja</span><span class="sxs-lookup"><span data-stu-id="1b6dd-110">Yes</span></span>  | <span data-ttu-id="1b6dd-111">Eine Zeichenfolge mit der Ressourcen-Id der JSON-Datei, die von benutzerdefinierten Funktionen.</span><span class="sxs-lookup"><span data-stu-id="1b6dd-111">String with the resource id of the JSON file used by custom functions.</span></span> |

## <a name="example"></a><span data-ttu-id="1b6dd-112">Beispiel</span><span class="sxs-lookup"><span data-stu-id="1b6dd-112">Example</span></span>

```xml
<Metadata>
    <SourceLocation resid="JSON-URL" />
</Metadata>
```
