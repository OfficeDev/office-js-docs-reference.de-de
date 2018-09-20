# <a name="webapplicationinfo-element"></a><span data-ttu-id="1122a-101">Element „WebApplicationInfo“</span><span class="sxs-lookup"><span data-stu-id="1122a-101">WebApplicationInfo element</span></span>

<span data-ttu-id="1122a-102">Dieses Element unterstützt einmaliges Anmelden (Single Sign-On, SSO) in Office-Add-Ins. Es enthält Informationen zu dem Add-In als:</span><span class="sxs-lookup"><span data-stu-id="1122a-102">Supports single sign-on (SSO) in Office Add-ins. This element contains information for the add-in as both:</span></span>

- <span data-ttu-id="1122a-103">OAuth 2.0-*Ressource*, für die die Office-Hostanwendung möglicherweise Berechtigungen benötigt</span><span class="sxs-lookup"><span data-stu-id="1122a-103">An OAuth 2.0 *resource* to which the Office host application might need permissions.</span></span>
- <span data-ttu-id="1122a-104">OAuth 2.0-*Client*, der möglicherweise Berechtigungen für Microsoft Graph benötigt</span><span class="sxs-lookup"><span data-stu-id="1122a-104">An OAuth 2.0 *client* that might need permissions to Microsoft Graph.</span></span>

<span data-ttu-id="1122a-105">**WebApplicationInfo** ist ein untergeordnetes Element des Elements [VersionOverrides](versionoverrides.md) im Manifest.</span><span class="sxs-lookup"><span data-stu-id="1122a-105">**WebApplicationInfo** is a child element of the [VersionOverrides](versionoverrides.md) element in the manifest.</span></span>  

## <a name="child-elements"></a><span data-ttu-id="1122a-106">Untergeordnete Elemente</span><span class="sxs-lookup"><span data-stu-id="1122a-106">Child elements</span></span>

|  <span data-ttu-id="1122a-107">Element</span><span class="sxs-lookup"><span data-stu-id="1122a-107">Element</span></span> |  <span data-ttu-id="1122a-108">Erforderlich</span><span class="sxs-lookup"><span data-stu-id="1122a-108">Required</span></span>  |  <span data-ttu-id="1122a-109">Beschreibung</span><span class="sxs-lookup"><span data-stu-id="1122a-109">Description</span></span>  |
|:-----|:-----|:-----|
|  <span data-ttu-id="1122a-110">**Id**</span><span class="sxs-lookup"><span data-stu-id="1122a-110">**Id**</span></span>    |  <span data-ttu-id="1122a-111">Ja</span><span class="sxs-lookup"><span data-stu-id="1122a-111">Yes</span></span>   |  <span data-ttu-id="1122a-112">Die **Anwendungs-ID** des zugeordneten Dienstes des Add-Ins, der im Azure Active Directory v 2.0-Endpunkt registriert ist.</span><span class="sxs-lookup"><span data-stu-id="1122a-112">The **Application Id** of the add-in's associated service as registered in the Azure Active Directory v 2.0 endpoint.</span></span>|
|  <span data-ttu-id="1122a-113">**Ressource**</span><span class="sxs-lookup"><span data-stu-id="1122a-113">**Resource**</span></span>  |  <span data-ttu-id="1122a-114">Ja</span><span class="sxs-lookup"><span data-stu-id="1122a-114">Yes</span></span>   |  <span data-ttu-id="1122a-115">Gibt den **Anwendungs-ID-URI** des Add-Ins an, der im Azure Active Directory v2.2.0-Endpunkt registriert ist.</span><span class="sxs-lookup"><span data-stu-id="1122a-115">Specifies the **Application ID URI** of the add-in as registered in the Azure Active Directory v 2.0 endpoint.</span></span>|
|  [<span data-ttu-id="1122a-116">Scopes</span><span class="sxs-lookup"><span data-stu-id="1122a-116">Scopes</span></span>](scopes.md)                |  <span data-ttu-id="1122a-117">Nein</span><span class="sxs-lookup"><span data-stu-id="1122a-117">No</span></span>  |  <span data-ttu-id="1122a-118">Gibt die Berechtigungen für Microsoft Graph an, die das Add-In benötigt.</span><span class="sxs-lookup"><span data-stu-id="1122a-118">Specifies the permissions that the add-in needs to Microsoft Graph.</span></span>  |

> [!NOTE] 
> <span data-ttu-id="1122a-119">Derzeit ist es erforderlich, dass Ihr Add-in Ressource Host übereinstimmt.</span><span class="sxs-lookup"><span data-stu-id="1122a-119">Currently, it's necessary that your add-in's Resource matches its Host.</span></span> <span data-ttu-id="1122a-120">Office fordert kein Token für ein Add-In an, es sein denn, der Besitz wird nachgewiesen. Heutzutage erfolgt dies durch Hosten des Add-Ins unter dem vollqualifizierten Domänennamen der Ressource.</span><span class="sxs-lookup"><span data-stu-id="1122a-120">Office will not request a Token for an add-in unless it can prove ownership, and today this is done by hosting the add-in under the Resource's fully-qualified domain name.</span></span>

## <a name="webapplicationinfo-example"></a><span data-ttu-id="1122a-121">Beispiel für das Element „WebApplicationInfo“</span><span class="sxs-lookup"><span data-stu-id="1122a-121">WebApplicationInfo example</span></span>

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
