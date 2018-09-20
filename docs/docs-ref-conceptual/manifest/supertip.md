# <a name="supertip"></a>Multiinfo

Definiert eine umfangreiche QuickInfo (Titel und Beschreibung). Es wird von [Button](control.md#button-control)- oder [Menu](control.md#menu-dropdown-button-controls)Steuerelementen verwendet.

## <a name="child-elements"></a>Untergeordnete Elemente

|  Element |  Erforderlich  |  Beschreibung  |
|:-----|:-----|:-----|
|  [Titel](#title)        | Ja |   Der Text für die Multiinfo.         |
|  [Beschreibung](#description)  | Ja |  Die Beschreibung für die Multiinfo.    |

### <a name="title"></a>Title

Erforderlich. Der Text für die Multiinfo. Das Attribut  **resid** muss auf den Wert des Attributs **id** eines **String**-Elements im  **ShortStrings**-Element im  [Resources](resources.md)-Element festgelegt werden.

### <a name="description"></a>Beschreibung

Erforderlich. Die Beschreibung für die Multiinfo. Das Attribut **Resid** muss auf den Wert des **Id** -Attributs ein **String** -Element im **LongStrings** -Element im [Ressourcen](resources.md) -Element festgelegt werden.

## <a name="example"></a>Beispiel

```xml
<Supertip>
    <Title resid="funcReadSuperTipTitle" />
    <Description resid="funcReadSuperTipDescription" />
</Supertip>
```
