# <a name="versionoverrides-element"></a><span data-ttu-id="08a71-101">VersionOverrides-Element</span><span class="sxs-lookup"><span data-stu-id="08a71-101">VersionOverrides element</span></span>

<span data-ttu-id="08a71-p101">Das Stammelement, das Informationen für die Add-In-Befehle enthält, die durch das Add-In implementiert werden. **VersionOverrides** ist ein untergeordnetes Element des [OfficeApp](./officeapp.md)-Elements im Manifest. Dieses Element wird in Manifestschema v1. 1 und höher unterstützt, ist jedoch im VersionOverrides-Schema v1.0 oder v1.1 definiert.</span><span class="sxs-lookup"><span data-stu-id="08a71-p101">The root element that contains information for the add-in commands implemented by the add-in. **VersionOverrides** is a child element of the [OfficeApp](./officeapp.md) element in the manifest. This element is supported in manifest schema v1.1 and later but is defined in the VersionOverrides v1.0 or v1.1 schema.</span></span>

## <a name="attributes"></a><span data-ttu-id="08a71-105">Attribute</span><span class="sxs-lookup"><span data-stu-id="08a71-105">Attributes</span></span>

|  <span data-ttu-id="08a71-106">Attribut</span><span class="sxs-lookup"><span data-stu-id="08a71-106">Attribute</span></span>  |  <span data-ttu-id="08a71-107">Erforderlich</span><span class="sxs-lookup"><span data-stu-id="08a71-107">Required</span></span>  |  <span data-ttu-id="08a71-108">Beschreibung</span><span class="sxs-lookup"><span data-stu-id="08a71-108">Description</span></span>  |
|:-----|:-----|:-----|
|  <span data-ttu-id="08a71-109">**xmlns**</span><span class="sxs-lookup"><span data-stu-id="08a71-109">**xmlns**</span></span>       |  <span data-ttu-id="08a71-110">Ja</span><span class="sxs-lookup"><span data-stu-id="08a71-110">Yes</span></span>  |  <span data-ttu-id="08a71-111">Der Schemaspeicherort, der `http://schemas.microsoft.com/office/mailappversionoverrides` sein muss, wenn `xsi:type` `VersionOverridesV1_0` ist, und `http://schemas.microsoft.com/office/mailappversionoverrides/1.1`, wenn `xsi:type` `VersionOverridesV1_1` ist.</span><span class="sxs-lookup"><span data-stu-id="08a71-111">The schema location, which must be `http://schemas.microsoft.com/office/mailappversionoverrides` when `xsi:type` is `VersionOverridesV1_0`, and `http://schemas.microsoft.com/office/mailappversionoverrides/1.1` when `xsi:type` is `VersionOverridesV1_1`.</span></span>|
|  <span data-ttu-id="08a71-112">**xsi:type**</span><span class="sxs-lookup"><span data-stu-id="08a71-112">**xsi:type**</span></span>  |  <span data-ttu-id="08a71-113">Ja</span><span class="sxs-lookup"><span data-stu-id="08a71-113">Yes</span></span>  | <span data-ttu-id="08a71-p102">Die Schemaversion. Zu diesem Zeitpunkt sind die einzigen zulässigen Werte `VersionOverridesV1_0` und `VersionOverridesV1_1`.</span><span class="sxs-lookup"><span data-stu-id="08a71-p102">The schema version. At this time, the only valid values are `VersionOverridesV1_0` and `VersionOverridesV1_1`.</span></span> |

> [!NOTE]
> <span data-ttu-id="08a71-116">Derzeit nur Outlook 2016 Schema v1. 1 VersionOverrides unterstützt und die `VersionOverridesV1_1` Typ.</span><span class="sxs-lookup"><span data-stu-id="08a71-116">Currently only Outlook 2016 supports the VersionOverrides v1.1 schema and the `VersionOverridesV1_1` type.</span></span>

## <a name="child-elements"></a><span data-ttu-id="08a71-117">Untergeordnete Elemente</span><span class="sxs-lookup"><span data-stu-id="08a71-117">Child elements</span></span>

|  <span data-ttu-id="08a71-118">Element</span><span class="sxs-lookup"><span data-stu-id="08a71-118">Element</span></span> |  <span data-ttu-id="08a71-119">Erforderlich</span><span class="sxs-lookup"><span data-stu-id="08a71-119">Required</span></span>  |  <span data-ttu-id="08a71-120">Beschreibung</span><span class="sxs-lookup"><span data-stu-id="08a71-120">Description</span></span>  |
|:-----|:-----|:-----|
|  <span data-ttu-id="08a71-121">**Beschreibung**</span><span class="sxs-lookup"><span data-stu-id="08a71-121">**Description**</span></span>    |  <span data-ttu-id="08a71-122">Nein</span><span class="sxs-lookup"><span data-stu-id="08a71-122">No</span></span>   |  <span data-ttu-id="08a71-p103">Beschreibt das Add-In. Dies überschreibt ggf. das `Description`-Element im übergeordneten Teil des Manifests. Der Text der Beschreibung ist in einem untergeordneten Element des **LongString**-Elements enthalten, das im [Ressourcen](./resources.md)Element enthalten ist. Das `resid`-Attribut des **Description**-Elements wird auf den Wert des `id`-Attributs des `String`-Elements festgelegt, das den Text enthält.</span><span class="sxs-lookup"><span data-stu-id="08a71-p103">Describes the add-in. This overrides the `Description` element in any parent portion of the manifest. The text of the description is contained in a child element of the **LongString** element contained in the [Resources](./resources.md) element. The `resid` attribute of the **Description** element is set to the value of the `id` attribute of the `String` element that contains the text.</span></span>|
|  <span data-ttu-id="08a71-127">**Anforderungen**</span><span class="sxs-lookup"><span data-stu-id="08a71-127">**Requirements**</span></span>  |  <span data-ttu-id="08a71-128">Nein</span><span class="sxs-lookup"><span data-stu-id="08a71-128">No</span></span>   |  <span data-ttu-id="08a71-p104">Gibt die Mindestanforderungen und die Version von Office.js an, die das Add-In erfordert. Dies überschreibt das `Requirements`-Element im übergeordneten Teil des Manifests.</span><span class="sxs-lookup"><span data-stu-id="08a71-p104">Specifies the minimum requirement set and version of Office.js that the add-in requires. This overrides the  `Requirements` element in the parent portion of the manifest.</span></span>|
|  [<span data-ttu-id="08a71-131">Hosts</span><span class="sxs-lookup"><span data-stu-id="08a71-131">Hosts</span></span>](./hosts.md)                |  <span data-ttu-id="08a71-132">Ja</span><span class="sxs-lookup"><span data-stu-id="08a71-132">Yes</span></span>  |  <span data-ttu-id="08a71-p105">Gibt eine Auflistung von Office-Hosts an. Das untergeordnete Hosts-Element überschreibt das Hosts-Element im übergeordneten Teil des Manifests.</span><span class="sxs-lookup"><span data-stu-id="08a71-p105">Specifies a collection of Office hosts. The child  Hosts element overrides the Hosts element in the parent portion of the manifest.</span></span>  |
|  [<span data-ttu-id="08a71-135">Ressourcen</span><span class="sxs-lookup"><span data-stu-id="08a71-135">Resources</span></span>](./resources.md)    |  <span data-ttu-id="08a71-136">Ja</span><span class="sxs-lookup"><span data-stu-id="08a71-136">Yes</span></span>  | <span data-ttu-id="08a71-137">Definiert eine Auflistung von Ressourcen (Zeichenfolgen, URLs und Bilder), auf die von anderen Elementen des Manifests verwiesen wird.</span><span class="sxs-lookup"><span data-stu-id="08a71-137">Defines a collection of resources (strings, URLs, and images) that other manifest elements reference.</span></span>|
|  <span data-ttu-id="08a71-138">**VersionOverrides**</span><span class="sxs-lookup"><span data-stu-id="08a71-138">**VersionOverrides**</span></span>    |  <span data-ttu-id="08a71-139">Nein</span><span class="sxs-lookup"><span data-stu-id="08a71-139">No</span></span>  | <span data-ttu-id="08a71-p106">Definiert Add-In-Befehle unter einer neueren Schemaversion. Einzelheiten finden Sie unter [Implementieren mehrerer Versionen](#implementing-multiple-versions).</span><span class="sxs-lookup"><span data-stu-id="08a71-p106">Defines add-in commands under a newer schema version. See [Implementing multiple versions](#implementing-multiple-versions) for details.</span></span> |
|  <span data-ttu-id="08a71-142">**WebApplicationInfo**</span><span class="sxs-lookup"><span data-stu-id="08a71-142">**WebApplicationInfo**</span></span>    |  <span data-ttu-id="08a71-143">Nein</span><span class="sxs-lookup"><span data-stu-id="08a71-143">No</span></span>  | <span data-ttu-id="08a71-144">Gibt Details über die zugeordnete Web-Anwendung des Add-ins an.</span><span class="sxs-lookup"><span data-stu-id="08a71-144">Specifies details about the add-in's associated Web application.</span></span> |



### <a name="versionoverrides-example"></a><span data-ttu-id="08a71-145">„VersionOverrides“-Beispiel</span><span class="sxs-lookup"><span data-stu-id="08a71-145">VersionOverrides example</span></span>
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

## <a name="implementing-multiple-versions"></a><span data-ttu-id="08a71-146">Implementieren mehrerer Versionen</span><span class="sxs-lookup"><span data-stu-id="08a71-146">Implementing multiple versions</span></span>

<span data-ttu-id="08a71-p107">Ein Manifest kann mehrere Versionen des `VersionOverrides`-Elements implementieren, das unterschiedliche Versionen des VersionOverrides-Schemas unterstützt. Auf diese Weise können neue Features in einem neueren Schema unterstützt werden, während gleichzeitig ältere Clients unterstützt werden, die die neuen Features nicht unterstützen.</span><span class="sxs-lookup"><span data-stu-id="08a71-p107">A manifest can implement multiple versions of the `VersionOverrides` element which support different versions of the VersionOverrides schema. This can be done to optionally support new features in a newer schema while still supporting older clients that do not support the new features.</span></span>

<span data-ttu-id="08a71-149">Um mehrere Versionen zu implementieren, muss das `VersionOverrides`-Element für die neuere Version ein untergeordnetes Element des `VersionOverrides`-Elements für die ältere Version sein.</span><span class="sxs-lookup"><span data-stu-id="08a71-149">In order to implement multiple versions, the `VersionOverrides` element for the newer version must be a child of the `VersionOverrides` element for the older version.</span></span> <span data-ttu-id="08a71-150">Das untergeordnete Element `VersionOverrides` erbt keine Werte vom übergeordneten Element.</span><span class="sxs-lookup"><span data-stu-id="08a71-150">The child `VersionOverrides` element doesn't inherit any values from the parent.</span></span>

<span data-ttu-id="08a71-151">Um sowohl das VersionOverrides-Schema v1.0 als auch v1.1 zu implementieren, würde das Manifest wie im folgenden Beispiel aussehen:</span><span class="sxs-lookup"><span data-stu-id="08a71-151">To implement both the VersionOverrides v1.0 and v1.1 schema, the manifest would look similar to the following example:</span></span>

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
