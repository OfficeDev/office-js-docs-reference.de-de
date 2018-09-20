# <a name="sets-element"></a>Sets-Element

Gibt die minimale Untergruppe der JavaScript-API für Office an, die Ihr Office-Add-In zur Aktivierung benötigt.

**Add-In-Typ:** Inhalt, Aufgabenbereich, E-Mail

## <a name="syntax"></a>Syntax

```XML
<Sets DefaultMinVersion="n .n ">
   ...
</Sets>
```

## <a name="contained-in"></a>Enthalten in

[Requirements](requirements.md)

## <a name="can-contain"></a>Kann enthalten

[Set](set.md)

## <a name="attributes"></a>Attribute

|**Attribut**|**Typ**|**Erforderlich**|**Beschreibung**|
|:-----|:-----|:-----|:-----|
|DefaultMinVersion|string|Optional|Gibt den Standardattributwert  **MinVersion** für alle untergeordneten [Set](set.md)-Elemente an. Der Standardwert lautet „1.1“.|

## <a name="remarks"></a>Hinweise

Weitere Informationen zu anforderungssätze finden Sie unter [Office-Versionen und Anforderung festgelegt](https://docs.microsoft.com/office/dev/add-ins/develop/office-versions-and-requirement-sets).

Weitere Informationen zum **MinVersion**-Attribut des **Set**-Elements und zum **DefaultMinVersion**-Attribut des **Sets**-Elements finden Sie unter [Festlegen des Requirements-Element im Manifest](https://docs.microsoft.com/office/dev/add-ins/develop/specify-office-hosts-and-api-requirements#set-the-requirements-element-in-the-manifest).

