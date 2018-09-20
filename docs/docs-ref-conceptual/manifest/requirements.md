# <a name="requirements-element"></a>Requirements-Element

Gibt den minimalen Satz von JavaScript API für Office-Anforderungen ([Anforderungssätze](https://docs.microsoft.com/office/dev/add-ins/develop/office-versions-and-requirement-sets#specify-office-hosts-and-requirement-sets) und/oder Methoden) an, die Ihr Office-Add-In aktivieren muss.

**Add-In-Typ:** Inhalt, Aufgabenbereich, E-Mail

## <a name="syntax"></a>Syntax

```XML
<Requirements>
   ...
</Requirements>
```

## <a name="contained-in"></a>Enthalten in

[OfficeApp](officeapp.md)

## <a name="can-contain"></a>Kann enthalten

|**Element**|**Inhalt**|**E-Mail**|**TaskPane**|
|:-----|:-----|:-----|:-----|
|[Sätzen](sets.md)|x|x|x|
|[Methoden](methods.md)|x||x|

## <a name="remarks"></a>Bemerkungen

Weitere Informationen zu anforderungssätze finden Sie unter [Office-Versionen und Anforderung festgelegt](https://docs.microsoft.com/office/dev/add-ins/develop/office-versions-and-requirement-sets).

