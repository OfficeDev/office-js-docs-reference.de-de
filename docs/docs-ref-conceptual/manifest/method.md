# <a name="method-element"></a>Method-Element

Gibt eine individuelle Methode aus der JavaScript-API für Office an, die Ihr Office-Add-In zum Aktivieren benötigt.

**Add-In-Typ:** Inhalt, Aufgabenbereich

## <a name="syntax"></a>Syntax

```XML
<Method Name="string"/>
```

## <a name="contained-in"></a>Enthalten in

[Methoden](methods.md)

## <a name="attributes"></a>Attribute

|**Attribut**|**Typ**|**Erforderlich**|**Beschreibung**|
|:-----|:-----|:-----|:-----|
|Name|string|erforderlich|Gibt den Namen der erforderlichen Methode qualifiziert mit dem übergeordneten Objekt an. Beispiel: Um die **getSelectedDataAsync**-Methode anzugeben, müssen Sie `"Document.getSelectedDataAsync"` angeben.|

## <a name="remarks"></a>Hinweise

Die **Methoden** und **Methode** Elemente werden von Mail-add-ins nicht unterstützt. Weitere Informationen zu anforderungssätze finden Sie unter [Office-Versionen und Anforderung festgelegt](https://docs.microsoft.com/office/dev/add-ins/develop/office-versions-and-requirement-sets).

> [!IMPORTANT] 
> Da es nicht möglich ist, die Mindestversion-Anforderung für die einzelnen Methoden angeben, sollten, um sicherzustellen, dass eine Methode zur Laufzeit verfügbar ist Sie auch eine **if** -Anweisung verwenden beim Aufrufen dieser Methode im Skript Ihr Add-in. Weitere Informationen hierzu finden Sie unter [Grundlegendes zur JavaScript-API für Office](https://docs.microsoft.com/office/dev/add-ins/develop/understanding-the-javascript-api-for-office).

