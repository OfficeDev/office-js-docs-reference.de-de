# <a name="appdomains-element"></a>AppDomains-Element

Listet neben der im SourceLocation-Element angegebenen Domäne alle Domänen auf, die das Office-Add-In zum Laden von Seiten verwendet. Geben Sie für jede zusätzliche Domäne ein AppDomain-Element an.

 **Add-In-Typ:** Inhalt, Aufgabenbereich, E-Mail

## <a name="syntax"></a>Syntax

```XML
<AppDomains>
    <AppDomain>AppDomain1</AppDomain>
    <AppDomain>AppDomain2</AppDomain>
</AppDomains>
```

## <a name="contained-in"></a>Enthalten in

[OfficeApp](officeapp.md)

## <a name="can-contain"></a>Kann enthalten

[AppDomain](appdomain.md)

## <a name="remarks"></a>Hinweise

Standardmäßig kann Ihr Add-In alle Seiten laden, die sich in derselben Domäne wie der im **SourceLocation**-Element angegebene Ort befinden. Geben Sie zum Laden von Seiten, die sich nicht in derselben Domäne wie das Add-In befinden, die Domänen mithilfe der **AppDomains**- und **AppDomain**- Elemente an. Dieses Element kann nicht leer sein. 
