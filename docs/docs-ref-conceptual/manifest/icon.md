# <a name="icon-element"></a>Icon-Element

Definiert **Image**-Elemente für [Button](control.md#button-control)- oder [Menu](control.md#menu-dropdown-button-controls)-Steuerelemente.

## <a name="attributes"></a>Attribute

|  Attribut  |  Erforderlich  |  Beschreibung  |
|:-----|:-----|:-----|
|  **xsi:type**  |  Nein  | Der Typ des Symbols, das definiert wird. Dies gilt nur für Symbole in mobilen Formularfaktoren. Bei **Icon**-Elementen, die in einem [MobileFormFactor](mobileformfactor.md)-Element enthalten sind, muss dieses Attribut auf `bt:MobileIconList` festgelegt sein. |

## <a name="child-elements"></a>Untergeordnete Elemente

|  Element |  Erforderlich  |  Beschreibung  |
|:-----|:-----|:-----|
|  [Image](#image)        | Ja |   resid-Element eines zu verwendenden Bilds         |

### <a name="image"></a>Bild

Ein Bild für die Schaltfläche. Das Attribut  **resid** muss auf einen Wert des **id**-Attributs eines  **Image**-Elements im  **Images**-Element im  [Resources](resources.md) -Element festgelegt werden. Das Attribut **size** gibt die Größe des Bilds in Pixel an. Drei Bildgrößen sind erforderlich (16, 32 und 80 Pixel), es werden aber fünf weitere Größen unterstützt (20, 24, 40, 48 und 64 Pixel).|

```xml
<Icon>
  <bt:Image size="16" resid="blue-icon-16" />
  <bt:Image size="32" resid="blue-icon-32" />
  <bt:Image size="80" resid="blue-icon-80" />
</Icon>
```

## <a name="additional-requirements-for-mobile-form-factors"></a>Weitere Anforderungen für mobile Formularfaktoren

Wenn das übergeordnete **Icon**-Element ist ein Nachfolger eines [MobileFormFactor](mobileformfactor.md)-Elements ist, sind die mindestens erforderlichen Größen geringfügig anders. Das Manifest muss mindestens 25, 32 und 48 Pixelgrößen bereitstellen. Jede bereitgestellte Größe muss mindestens dreimal angezeigt werden, wobei das `scale`-Attribut auf `1`, `2` oder `3` festgelegt ist.

```xml
<Icon xsi:type="bt:MobileIconList">
  <bt:Image resid="blue-icon-16-1" size="25" scale="1" />
  <bt:Image resid="blue-icon-16-2" size="25" scale="2" />
  <bt:Image resid="blue-icon-16-3" size="25" scale="3" />
  <bt:Image resid="blue-icon-32-1" size="32" scale="1" />
  <bt:Image resid="blue-icon-32-2" size="32" scale="2" />
  <bt:Image resid="blue-icon-32-3" size="32" scale="3" />
  <bt:Image resid="blue-icon-80-1" size="48" scale="1" />
  <bt:Image resid="blue-icon-80-2" size="48" scale="2" />
  <bt:Image resid="blue-icon-80-3" size="48" scale="3" />
</Icon>
```