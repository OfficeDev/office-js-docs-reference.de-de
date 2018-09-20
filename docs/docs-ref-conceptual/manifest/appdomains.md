# <a name="appdomains-element"></a><span data-ttu-id="c0a81-101">AppDomains-Element</span><span class="sxs-lookup"><span data-stu-id="c0a81-101">AppDomains element</span></span>

<span data-ttu-id="c0a81-p101">Listet neben der im SourceLocation-Element angegebenen Domäne alle Domänen auf, die das Office-Add-In zum Laden von Seiten verwendet. Geben Sie für jede zusätzliche Domäne ein AppDomain-Element an.</span><span class="sxs-lookup"><span data-stu-id="c0a81-p101">Lists any domains in addition to the domain specified in the SourceLocation element that your Office Add-in will use to load pages. For each additional domain, specify an AppDomain element.</span></span>

 <span data-ttu-id="c0a81-104">**Add-In-Typ:** Inhalt, Aufgabenbereich, E-Mail</span><span class="sxs-lookup"><span data-stu-id="c0a81-104">**Add-in type:** Content, Task pane, Mail</span></span>

## <a name="syntax"></a><span data-ttu-id="c0a81-105">Syntax</span><span class="sxs-lookup"><span data-stu-id="c0a81-105">Syntax</span></span>

```XML
<AppDomains>
    <AppDomain>AppDomain1</AppDomain>
    <AppDomain>AppDomain2</AppDomain>
</AppDomains>
```

## <a name="contained-in"></a><span data-ttu-id="c0a81-106">Enthalten in</span><span class="sxs-lookup"><span data-stu-id="c0a81-106">Contained in</span></span>

[<span data-ttu-id="c0a81-107">OfficeApp</span><span class="sxs-lookup"><span data-stu-id="c0a81-107">OfficeApp</span></span>](officeapp.md)

## <a name="can-contain"></a><span data-ttu-id="c0a81-108">Kann enthalten</span><span class="sxs-lookup"><span data-stu-id="c0a81-108">Can contain</span></span>

[<span data-ttu-id="c0a81-109">AppDomain</span><span class="sxs-lookup"><span data-stu-id="c0a81-109">AppDomain</span></span>](appdomain.md)

## <a name="remarks"></a><span data-ttu-id="c0a81-110">Hinweise</span><span class="sxs-lookup"><span data-stu-id="c0a81-110">Remarks</span></span>

<span data-ttu-id="c0a81-p102">Standardmäßig kann Ihr Add-In alle Seiten laden, die sich in derselben Domäne wie der im **SourceLocation**-Element angegebene Ort befinden. Geben Sie zum Laden von Seiten, die sich nicht in derselben Domäne wie das Add-In befinden, die Domänen mithilfe der **AppDomains**- und **AppDomain**- Elemente an. Dieses Element kann nicht leer sein.</span><span class="sxs-lookup"><span data-stu-id="c0a81-p102">By default, your add-in can load any page that is in the same domain as the location specified in the **SourceLocation** element. To load pages that are not in the same domain as the add-in, specify the domains by using the **AppDomains** and **AppDomain** elements. This element can't be empty.</span></span> 
