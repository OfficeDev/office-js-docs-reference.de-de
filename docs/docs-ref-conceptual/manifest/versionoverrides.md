# <a name="versionoverrides-element"></a>VersionOverrides-Element

Das Stammelement, das Informationen für die Add-In-Befehle enthält, die durch das Add-In implementiert werden. **VersionOverrides** ist ein untergeordnetes Element des [OfficeApp](./officeapp.md)-Elements im Manifest. Dieses Element wird in Manifestschema v1. 1 und höher unterstützt, ist jedoch im VersionOverrides-Schema v1.0 oder v1.1 definiert.

## <a name="attributes"></a>Attribute

|  Attribut  |  Erforderlich  |  Beschreibung  |
|:-----|:-----|:-----|
|  **xmlns**       |  Ja  |  Der Schemaspeicherort, der `http://schemas.microsoft.com/office/mailappversionoverrides` sein muss, wenn `xsi:type` `VersionOverridesV1_0` ist, und `http://schemas.microsoft.com/office/mailappversionoverrides/1.1`, wenn `xsi:type` `VersionOverridesV1_1` ist.|
|  **xsi:type**  |  Ja  | Die Schemaversion. Zu diesem Zeitpunkt sind die einzigen zulässigen Werte `VersionOverridesV1_0` und `VersionOverridesV1_1`. |

> [!NOTE]
> Derzeit nur Outlook 2016 Schema v1. 1 VersionOverrides unterstützt und die `VersionOverridesV1_1` Typ.

## <a name="child-elements"></a>Untergeordnete Elemente

|  Element |  Erforderlich  |  Beschreibung  |
|:-----|:-----|:-----|
|  **Beschreibung**    |  Nein   |  Beschreibt das Add-In. Dies überschreibt ggf. das `Description`-Element im übergeordneten Teil des Manifests. Der Text der Beschreibung ist in einem untergeordneten Element des **LongString**-Elements enthalten, das im [Ressourcen](./resources.md)Element enthalten ist. Das `resid`-Attribut des **Description**-Elements wird auf den Wert des `id`-Attributs des `String`-Elements festgelegt, das den Text enthält.|
|  **Anforderungen**  |  Nein   |  Gibt die Mindestanforderungen und die Version von Office.js an, die das Add-In erfordert. Dies überschreibt das `Requirements`-Element im übergeordneten Teil des Manifests.|
|  [Hosts](./hosts.md)                |  Ja  |  Gibt eine Auflistung von Office-Hosts an. Das untergeordnete Hosts-Element überschreibt das Hosts-Element im übergeordneten Teil des Manifests.  |
|  [Ressourcen](./resources.md)    |  Ja  | Definiert eine Auflistung von Ressourcen (Zeichenfolgen, URLs und Bilder), auf die von anderen Elementen des Manifests verwiesen wird.|
|  **VersionOverrides**    |  Nein  | Definiert Add-In-Befehle unter einer neueren Schemaversion. Einzelheiten finden Sie unter [Implementieren mehrerer Versionen](#implementing-multiple-versions). |
|  **WebApplicationInfo**    |  Nein  | Gibt Details über die zugeordnete Web-Anwendung des Add-ins an. |



### <a name="versionoverrides-example"></a>„VersionOverrides“-Beispiel
```xml
<OfficeApp>
...
  <VersionOverrides xmlns="http://schemas.microsoft.com/office/mailappversionoverrides" xsi:type="VersionOverridesV1_0">
    <Description resid="residDescription" />
    <Requirements>
      <!-- add information on requirements -->
    </Requirements>
    <Hosts>
      <Host xsi:type="MailHost">
        <!-- add information on form factors -->
      </Host>
    </Hosts>
    <Resources>
      <!-- add information on resources -->
    </Resources>
  </VersionOverrides>
...
</OfficeApp>
```

## <a name="implementing-multiple-versions"></a>Implementieren mehrerer Versionen

Ein Manifest kann mehrere Versionen des `VersionOverrides`-Elements implementieren, das unterschiedliche Versionen des VersionOverrides-Schemas unterstützt. Auf diese Weise können neue Features in einem neueren Schema unterstützt werden, während gleichzeitig ältere Clients unterstützt werden, die die neuen Features nicht unterstützen.

Um mehrere Versionen zu implementieren, muss das `VersionOverrides`-Element für die neuere Version ein untergeordnetes Element des `VersionOverrides`-Elements für die ältere Version sein. Das untergeordnete Element `VersionOverrides` erbt keine Werte vom übergeordneten Element.

Um sowohl das VersionOverrides-Schema v1.0 als auch v1.1 zu implementieren, würde das Manifest wie im folgenden Beispiel aussehen:

```xml
<OfficeApp>
...
  <VersionOverrides xmlns="http://schemas.microsoft.com/office/mailappversionoverrides" xsi:type="VersionOverridesV1_0">
    <Description resid="residDescription" />
    <Requirements>
      <!-- add information on requirements -->
    </Requirements>
    <Hosts>
      <Host xsi:type="MailHost">
        <!-- add information on form factors -->
      </Host>
    </Hosts>
    <Resources>
      <!-- add information on resources -->
    </Resources>

    <VersionOverrides xmlns="http://schemas.microsoft.com/office/mailappversionoverrides/1.1" xsi:type="VersionOverridesV1_1">
      <Description resid="residDescription" />
      <Requirements>
        <!-- add information on requirements -->
      </Requirements>
      <Hosts>
        <Host xsi:type="MailHost">
          <!-- add information on form factors -->
        </Host>
      </Hosts>
      <Resources>
        <!-- add information on resources -->
      </Resources>
    </VersionOverrides>  
...
</OfficeApp>
```
