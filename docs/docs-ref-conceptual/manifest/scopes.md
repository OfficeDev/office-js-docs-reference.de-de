# <a name="scopes-element"></a><span data-ttu-id="fb06e-101">Element „Scopes“</span><span class="sxs-lookup"><span data-stu-id="fb06e-101">Scopes element</span></span>

<span data-ttu-id="fb06e-102">Dieses Element enthält Berechtigungen für Microsoft Graph, die das Add-In benötigt.</span><span class="sxs-lookup"><span data-stu-id="fb06e-102">Contains permissions to Microsoft Graph that the add-in needs.</span></span> <span data-ttu-id="fb06e-103">Es wird vom Office Store verwendet, um ein Zustimmungsdialogfeld zu erstellen.</span><span class="sxs-lookup"><span data-stu-id="fb06e-103">The Office Store uses the Scopes element to create a consent dialog box.</span></span> <span data-ttu-id="fb06e-104">Bei der Installation des Add-Ins aus dem Store wird der Benutzer aufgefordert, dem Add-In die angegebenen Berechtigungen für seine Microsoft Graph-Daten zu erteilen.</span><span class="sxs-lookup"><span data-stu-id="fb06e-104">When users install the add-in from the Store, they are prompted to grant the add-in the specified permissions to the user's Microsoft Graph data.</span></span>

## <a name="child-elements"></a><span data-ttu-id="fb06e-105">Untergeordnete Elemente</span><span class="sxs-lookup"><span data-stu-id="fb06e-105">Child elements</span></span>

|  <span data-ttu-id="fb06e-106">Element</span><span class="sxs-lookup"><span data-stu-id="fb06e-106">Element</span></span> |  <span data-ttu-id="fb06e-107">Typ</span><span class="sxs-lookup"><span data-stu-id="fb06e-107">Type</span></span>  |  <span data-ttu-id="fb06e-108">Beschreibung</span><span class="sxs-lookup"><span data-stu-id="fb06e-108">Description</span></span>  |
|:-----|:-----|:-----|
|  <span data-ttu-id="fb06e-109">**Scope**</span><span class="sxs-lookup"><span data-stu-id="fb06e-109">**Scope**</span></span>                |  <span data-ttu-id="fb06e-110">string</span><span class="sxs-lookup"><span data-stu-id="fb06e-110">string</span></span>     |   <span data-ttu-id="fb06e-111">Der Name einer Berechtigung für Microsoft Graph, zum Beispiel „Files.Read.All“</span><span class="sxs-lookup"><span data-stu-id="fb06e-111">The name of a permission to Microsoft Graph; for example, Files.Read.All.</span></span> |

## <a name="example"></a><span data-ttu-id="fb06e-112">Beispiel</span><span class="sxs-lookup"><span data-stu-id="fb06e-112">Example</span></span>

```xml
<OfficeApp>
...
  <VersionOverrides xmlns="http://schemas.microsoft.com/office/mailappversionoverrides" xsi:type="VersionOverridesV1_0">
    ...
    <WebApplicationInfo>
      <Id>12345678-abcd-1234-efab-123456789abc</Id>
      <Resource>api://myDomain.com/12345678-abcd-1234-efab-123456789abc<Resource>
      <Scopes>
        <Scope>Files.Read.All</Scope>
        <Scope>offline_access</Scope>
        <Scope>openid</Scope>
        <Scope>profile</Scope>
      </Scopes>
    </WebApplicationInfo>
  </VersionOverrides>
...
</OfficeApp>
```
