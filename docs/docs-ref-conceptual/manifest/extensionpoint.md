# <a name="extensionpoint-element"></a><span data-ttu-id="a241e-101">ExtensionPoint-Element</span><span class="sxs-lookup"><span data-stu-id="a241e-101">ExtensionPoint element</span></span>

 <span data-ttu-id="a241e-102">Definiert, wo ein Add-In Funktionen in der Office-Benutzeroberfläche verfügbar macht.</span><span class="sxs-lookup"><span data-stu-id="a241e-102">Defines where an add-in exposes functionality in the Office UI.</span></span> <span data-ttu-id="a241e-103">Das **ExtensionPoint**-Element ist ein untergeordnetes Element von [AllFormFactor](allformfactors.md), [DesktopFormFactor](desktopformfactor.md) oder [MobileFormFactor](mobileformfactor.md).</span><span class="sxs-lookup"><span data-stu-id="a241e-103">The **ExtensionPoint** element is a child element of [AllFormFactors](allformfactors.md), [DesktopFormFactor](desktopformfactor.md) or [MobileFormFactor](mobileformfactor.md).</span></span> 

## <a name="attributes"></a><span data-ttu-id="a241e-104">Attribute</span><span class="sxs-lookup"><span data-stu-id="a241e-104">Attributes</span></span>

|  <span data-ttu-id="a241e-105">Attribut</span><span class="sxs-lookup"><span data-stu-id="a241e-105">Attribute</span></span>  |  <span data-ttu-id="a241e-106">Erforderlich</span><span class="sxs-lookup"><span data-stu-id="a241e-106">Required</span></span>  |  <span data-ttu-id="a241e-107">Beschreibung</span><span class="sxs-lookup"><span data-stu-id="a241e-107">Description</span></span>  |
|:-----|:-----|:-----|
|  <span data-ttu-id="a241e-108">**xsi:type**</span><span class="sxs-lookup"><span data-stu-id="a241e-108">**xsi:type**</span></span>  |  <span data-ttu-id="a241e-109">Ja</span><span class="sxs-lookup"><span data-stu-id="a241e-109">Yes</span></span>  | <span data-ttu-id="a241e-110">Der Typ des Erweiterungspunkts, der definiert wird.</span><span class="sxs-lookup"><span data-stu-id="a241e-110">The type of extension point being defined.</span></span>|

## <a name="extension-points-for-excel-only"></a><span data-ttu-id="a241e-111">Nur für Excel verfügbare Erweiterungspunkte</span><span class="sxs-lookup"><span data-stu-id="a241e-111">Extension points for Excel only</span></span>

- <span data-ttu-id="a241e-112">**CustomFunctions** - eine benutzerdefinierte Funktion in JavaScript für Excel geschrieben.</span><span class="sxs-lookup"><span data-stu-id="a241e-112">**CustomFunctions** - A custom function written in JavaScript for Excel.</span></span>

<span data-ttu-id="a241e-113">[Dieses XML-Codebeispiel](https://github.com/OfficeDev/Excel-Custom-Functions/blob/master/customfunctions.xml) veranschaulicht, wie das **ExtensionPoint** -Element mit der Attributwert **CustomFunctions** und die untergeordneten Elemente verwenden, um verwendet werden.</span><span class="sxs-lookup"><span data-stu-id="a241e-113">[This XML code sample](https://github.com/OfficeDev/Excel-Custom-Functions/blob/master/customfunctions.xml) shows how to use the **ExtensionPoint** element with the **CustomFunctions** attribute value, and the child elements to be used.</span></span>

## <a name="extension-points-for-word-excel-powerpoint-and-onenote-add-in-commands"></a><span data-ttu-id="a241e-114">Erweiterungspunkte für Word-, Excel-, PowerPoint und OneNote-Add-In-Befehle</span><span class="sxs-lookup"><span data-stu-id="a241e-114">Extension points for Word, Excel, PowerPoint, and OneNote add-in commands</span></span>

- <span data-ttu-id="a241e-115">**PrimaryCommandSurface** Das Menüband in Office.</span><span class="sxs-lookup"><span data-stu-id="a241e-115">**PrimaryCommandSurface** - The ribbon in Office.</span></span>
- <span data-ttu-id="a241e-116">**ContextMenu** Das Kontextmenü, das angezeigt wird, wenn Sie mit der rechten Maustaste auf die Office-Benutzeroberfläche klicken.</span><span class="sxs-lookup"><span data-stu-id="a241e-116">**ContextMenu** - The shortcut menu that appears when you right-click in the Office UI.</span></span>

<span data-ttu-id="a241e-117">In den folgenden Beispielen wird gezeigt, wie Sie das  **ExtensionPoint**-Element mit  **PrimaryCommandSurface**- und  **ContextMenu**-Attributwerten verwenden, und Sie sehen die untergeordneten Elemente, die für jedes Element verwendet werden sollten.</span><span class="sxs-lookup"><span data-stu-id="a241e-117">The following examples show how to use the  **ExtensionPoint** element with **PrimaryCommandSurface** and **ContextMenu** attribute values, and the child elements that should be used with each.</span></span>

> [!IMPORTANT] 
> <span data-ttu-id="a241e-118">Stellen Sie für Elemente, die das ID-Attribut enthalten sicher, dass Sie eine eindeutige ID bereitstellen</span><span class="sxs-lookup"><span data-stu-id="a241e-118">For elements that contain an ID attribute, make sure you provide a unique ID.</span></span> <span data-ttu-id="a241e-119">Es wird empfohlen, den Namen Ihres Unternehmens zusammen mit Ihrer ID zu verwenden.</span><span class="sxs-lookup"><span data-stu-id="a241e-119">We recommend that you use your company's name along with your ID.</span></span> <span data-ttu-id="a241e-120">Verwenden Sie beispielsweise das folgende Format.</span><span class="sxs-lookup"><span data-stu-id="a241e-120">For example, use the following format.</span></span> <CustomTab id="mycompanyname.mygroupname">

```XML
<ExtensionPoint xsi:type="PrimaryCommandSurface">
          <CustomTab id="Contoso Tab">
          <!-- If you want to use a default tab that comes with Office, remove the above CustomTab element, and then uncomment the following OfficeTab element -->
            <!-- <OfficeTab id="TabData"> -->
            <Label resid="residLabel4" />
            <Group id="Group1Id12">
              <Label resid="residLabel4" />
              <Icon>
                <bt:Image size="16" resid="icon1_32x32" />
                <bt:Image size="32" resid="icon1_32x32" />
                <bt:Image size="80" resid="icon1_32x32" />
              </Icon>
              <Tooltip resid="residToolTip" />
              <Control xsi:type="Button" id="Button1Id1">

                  <!-- information about the control -->
              </Control>
              <!-- other controls, as needed -->
            </Group>
          </CustomTab>
        </ExtensionPoint>

      <ExtensionPoint xsi:type="ContextMenu">
        <OfficeMenu id="ContextMenuCell">
          <Control xsi:type="Menu" id="ContextMenu2">
                  <!-- information about the control -->
          </Control>
          <!-- other controls, as needed -->
        </OfficeMenu>
        </ExtensionPoint>
```

#### <a name="child-elements"></a><span data-ttu-id="a241e-121">Untergeordnete Elemente</span><span class="sxs-lookup"><span data-stu-id="a241e-121">Child elements</span></span>
 
|<span data-ttu-id="a241e-122">**Element**</span><span class="sxs-lookup"><span data-stu-id="a241e-122">**Element**</span></span>|<span data-ttu-id="a241e-123">**Beschreibung**</span><span class="sxs-lookup"><span data-stu-id="a241e-123">**Description**</span></span>|
|:-----|:-----|
|<span data-ttu-id="a241e-124">**CustomTab**</span><span class="sxs-lookup"><span data-stu-id="a241e-124">**CustomTab**</span></span>|<span data-ttu-id="a241e-p103">Erforderlich, wenn Sie eine benutzerdefinierte Registerkarte (unter Verwendung von  **PrimaryCommandSurface**) zum Menüband hinzufügen möchten. Wenn Sie das  **CustomTab**-Element verwenden, können Sie das  **OfficeTab**-Element nicht verwenden. Das  **ID** -Attribut ist erforderlich.</span><span class="sxs-lookup"><span data-stu-id="a241e-p103">Required if you want to add a custom tab to the ribbon (using  **PrimaryCommandSurface**). If you use the  **CustomTab** element, you can't use the **OfficeTab** element. The **id** attribute is required.</span></span>|
|<span data-ttu-id="a241e-128">**OfficeTab**</span><span class="sxs-lookup"><span data-stu-id="a241e-128">**OfficeTab**</span></span>|<span data-ttu-id="a241e-p104">Erforderlich, wenn Sie eine standardmäßige Office-Menübandregisterkarte (unter Verwendung von **PrimaryCommandSurface**) erweitern möchten. Wenn Sie das **OfficeTab**-Element verwenden, können Sie das **CustomTab**-Element nicht verwenden. Weitere Informationen finden Sie unter [OfficeTab](officetab.md)</span><span class="sxs-lookup"><span data-stu-id="a241e-p104">Required if you want to extend a default Office ribbon tab (using **PrimaryCommandSurface**). If you use the  **OfficeTab** element, you can't use the **CustomTab** element. For details, see [OfficeTab](officetab.md).</span></span>|
|<span data-ttu-id="a241e-132">**OfficeMenu**</span><span class="sxs-lookup"><span data-stu-id="a241e-132">**OfficeMenu**</span></span>|<span data-ttu-id="a241e-p105">Erforderlich, wenn Sie Add-In-Befehle zu einem Standardkontextmenü (unter Verwendung von **ContextMenu**) hinzufügen. Das **id**-Attribut muss festgelegt werden auf: </span><span class="sxs-lookup"><span data-stu-id="a241e-p105">Required if you're adding add-in commands to a default context menu (using  **ContextMenu**). The  **id** attribute must be set to: </span></span><br/> <span data-ttu-id="a241e-p106">- **ContextMenuText** für Excel oder Word. Zeigt das Element im Kontextmenü an, wenn Text ausgewählt wird und der Benutzer dann mit der rechten Maustaste auf den ausgewählten Text klickt. </span><span class="sxs-lookup"><span data-stu-id="a241e-p106">- **ContextMenuText** for Excel or Word. Displays the item on the context menu when text is selected and then the user right-clicks on the selected text. </span></span><br/> <span data-ttu-id="a241e-p107">- **ContextMenuCell** für Excel. Zeigt das Element im Kontextmenü an, wenn der Benutzer mit der rechten Maustaste in eine Zelle in dem Arbeitsblatt klickt.</span><span class="sxs-lookup"><span data-stu-id="a241e-p107">- **ContextMenuCell** for Excel. Displays the  item on the context menu when the user right-clicks on a cell on the spreadsheet.</span></span>|
|<span data-ttu-id="a241e-139">**Group**</span><span class="sxs-lookup"><span data-stu-id="a241e-139">**Group**</span></span>|<span data-ttu-id="a241e-p108">Eine Gruppe von Erweiterungspunkten der Benutzeroberfläche auf einer Registerkarte.Eine Gruppe kann über bis zu sechs Steuerelemente verfügen. Das **id** -Attribut ist erforderlich. Es handelt sich um eine Zeichenfolge mit maximal 125 Zeichen.</span><span class="sxs-lookup"><span data-stu-id="a241e-p108">A group of user interface extension points on a tab. A group can have up to six controls. The  **id** attribute is required. It's a string with a maximum of 125 characters.</span></span>|
|<span data-ttu-id="a241e-143">**Label**</span><span class="sxs-lookup"><span data-stu-id="a241e-143">**Label**</span></span>|<span data-ttu-id="a241e-p109">Erforderlich. Die Bezeichnung der Gruppe. Das  **resid** -Attribut muss auf den Wert des **id**-Attributs eines  **String**-Elements festegelegt werden. Das  **String**-Element ist ein untergeordnetes Element vom  **ShortStrings**-Element, das ein untergeordnetes Element vom  **Resources** -Element ist.</span><span class="sxs-lookup"><span data-stu-id="a241e-p109">Required. The label of the group. The  **resid** attribute must be set to the value of the **id** attribute of a **String** element. The **String** element is a child element of the **ShortStrings** element, which is a child element of the **Resources** element.</span></span>|
|<span data-ttu-id="a241e-148">**Icon**</span><span class="sxs-lookup"><span data-stu-id="a241e-148">**Icon**</span></span>|<span data-ttu-id="a241e-p110">Erforderlich. Gibt das Gruppensymbol an, das auf kleinen Geräten verwendet wird, oder wenn zu viele Schaltflächen angezeigt werden. Das  **resid** -Attribut muss auf den Wert des **id**-Attributs eines  **Image**-Elements festgelegt werden. Das  **Image**-Element ist ein untergeordnetes Element vom  **Images**-Element, das ein untergeordnetes Element vom  **Resources** -Element ist. Das **size**-Attribut gibt die Größe des Bilds in Pixeln an. Es sind drei Bildgrößen erforderlich: 16, 32 und 80. Es werden auch fünf optionale Größen unterstützt: 20, 24, 40, 48 und 64.</span><span class="sxs-lookup"><span data-stu-id="a241e-p110">Required. Specifies the group's icon to be used on small form factor devices, or when too many buttons are displayed. The  **resid** attribute must be set to the value of the **id** attribute of an **Image** element. The **Image** element is a child element of the **Images** element, which is a child element of the **Resources** element. The **size** attribute gives the size, in pixels, of the image. Three image sizes are required: 16, 32, and 80. Five optional sizes are also supported: 20, 24, 40, 48, and 64.</span></span>|
|<span data-ttu-id="a241e-156">**Tooltip**</span><span class="sxs-lookup"><span data-stu-id="a241e-156">**Tooltip**</span></span>|<span data-ttu-id="a241e-p111">Optional. Die QuickInfo der Gruppe. Das  **resid** -Attribut muss auf den Wert des **id**-Attributs eines  **String**-Elements festgelegt werden. Das  **String**-Element ist ein untergeordnetes Element vom  **LongStrings**-Element, das ein untergeordnetes Element vom  **Resources** -Element ist.</span><span class="sxs-lookup"><span data-stu-id="a241e-p111">Optional. The tooltip of the group. The  **resid** attribute must be set to the value of the **id** attribute of a **String** element. The **String** element is a child element of the **LongStrings** element, which is a child element of the **Resources** element.</span></span>|
|<span data-ttu-id="a241e-161">**Control**</span><span class="sxs-lookup"><span data-stu-id="a241e-161">**Control**</span></span>|<span data-ttu-id="a241e-162">Jede Gruppe erfordert mindestens ein Steuerelement.</span><span class="sxs-lookup"><span data-stu-id="a241e-162">Each group requires at least one control.</span></span> <span data-ttu-id="a241e-163">Ein **Control**-Element kann entweder ein **Button** oder **Menu** sein.</span><span class="sxs-lookup"><span data-stu-id="a241e-163">A  **Control** element can be either a **Button** or a **Menu**.</span></span> <span data-ttu-id="a241e-164">Verwenden Sie **Menu**, um eine Dropdownliste mit Steuerelementen für Schaltflächen anzugeben.</span><span class="sxs-lookup"><span data-stu-id="a241e-164">Use  **Menu** to specify a drop-down list of button controls.</span></span> <span data-ttu-id="a241e-165">Derzeit werden nur Schaltflächen und Menüs unterstützt.</span><span class="sxs-lookup"><span data-stu-id="a241e-165">Currently, only buttons and menus are supported.</span></span> <span data-ttu-id="a241e-166">Weitere Informationen finden Sie in den Abschnitten [Steuerelemente für Schaltflächen](control.md#button-control) und [Menüsteuerelemente](control.md#menu-dropdown-button-controls).</span><span class="sxs-lookup"><span data-stu-id="a241e-166">See the [Button controls](control.md#button-control) and [Menu controls](control.md#menu-dropdown-button-controls) sections for more information.</span></span><br/><span data-ttu-id="a241e-167">**Hinweis:**  So machen Sie die Problembehandlung, wird empfohlen, ein **Control** -Element und die zugehörigen **Ressourcen** untergeordneten Elemente jeweils einzeln hinzugefügt werden.</span><span class="sxs-lookup"><span data-stu-id="a241e-167">**Note:**  To make troubleshooting easier, we recommend that a  **Control** element and the related **Resources** child elements be added one at a time.</span></span>|
|<span data-ttu-id="a241e-168">**Script**</span><span class="sxs-lookup"><span data-stu-id="a241e-168">**Script**</span></span>|<span data-ttu-id="a241e-169">Enthält Links zur JavaScript-Datei mit der benutzerdefinierten Funktionsdefinition und dem Registrierungscode.</span><span class="sxs-lookup"><span data-stu-id="a241e-169">Links to the JavaScript file with the custom function definition and registration code.</span></span> <span data-ttu-id="a241e-170">Dieses Element wird nicht in der Vorschau für Entwickler verwendet.</span><span class="sxs-lookup"><span data-stu-id="a241e-170">This element is not used in the Developer Preview.</span></span> <span data-ttu-id="a241e-171">Stattdessen werden alle JavaScript-Dateien über die HTML-Seite geladen.</span><span class="sxs-lookup"><span data-stu-id="a241e-171">Instead, the HTML page is responsible for loading all JavaScript files.</span></span>|
|<span data-ttu-id="a241e-172">**Page**</span><span class="sxs-lookup"><span data-stu-id="a241e-172">**Page**</span></span>|<span data-ttu-id="a241e-173">Enthält Links zur HTML-Seite für Ihre benutzerdefinierten Funktionen.</span><span class="sxs-lookup"><span data-stu-id="a241e-173">Links to the HTML page for your custom functions.</span></span>|

## <a name="extension-points-for-outlook"></a><span data-ttu-id="a241e-174">Erweiterungspunkte für Outlook</span><span class="sxs-lookup"><span data-stu-id="a241e-174">Extension points for Outlook</span></span>

- [<span data-ttu-id="a241e-175">MessageReadCommandSurface</span><span class="sxs-lookup"><span data-stu-id="a241e-175">MessageReadCommandSurface</span></span>](#messagereadcommandsurface) 
- [<span data-ttu-id="a241e-176">MessageComposeCommandSurface</span><span class="sxs-lookup"><span data-stu-id="a241e-176">MessageComposeCommandSurface</span></span>](#messagecomposecommandsurface) 
- [<span data-ttu-id="a241e-177">AppointmentOrganizerCommandSurface</span><span class="sxs-lookup"><span data-stu-id="a241e-177">AppointmentOrganizerCommandSurface</span></span>](#appointmentorganizercommandsurface) 
- [<span data-ttu-id="a241e-178">AppointmentAttendeeCommandSurface</span><span class="sxs-lookup"><span data-stu-id="a241e-178">AppointmentAttendeeCommandSurface</span></span>](#appointmentattendeecommandsurface)
- <span data-ttu-id="a241e-179">[Module](#module) (Kann nur im [DesktopFormFactor](desktopformfactor.md) verwendet werden.)</span><span class="sxs-lookup"><span data-stu-id="a241e-179">[Module](#module) (Can only be used in the [DesktopFormFactor](desktopformfactor.md).)</span></span>
- [<span data-ttu-id="a241e-180">MobileMessageReadCommandSurface</span><span class="sxs-lookup"><span data-stu-id="a241e-180">MobileMessageReadCommandSurface</span></span>](#mobilemessagereadcommandsurface)
- [<span data-ttu-id="a241e-181">Events</span><span class="sxs-lookup"><span data-stu-id="a241e-181">Events</span></span>](#events)
- [<span data-ttu-id="a241e-182">DetectedEntity</span><span class="sxs-lookup"><span data-stu-id="a241e-182">DetectedEntity</span></span>](#detectedentity)

### <a name="messagereadcommandsurface"></a><span data-ttu-id="a241e-183">MessageReadCommandSurface</span><span class="sxs-lookup"><span data-stu-id="a241e-183">MessageReadCommandSurface</span></span>
<span data-ttu-id="a241e-p114">Mit diesem Erweiterungspunkt werden Schaltflächen für die Ansicht gelesener Mails auf der Befehlsoberfläche platziert. In Outlook Desktop wird das Element im Menüband angezeigt.</span><span class="sxs-lookup"><span data-stu-id="a241e-p114">This extension point puts buttons in the command surface for the mail read view. In Outlook desktop, this appears in the ribbon.</span></span>

#### <a name="child-elements"></a><span data-ttu-id="a241e-186">Untergeordnete Elemente</span><span class="sxs-lookup"><span data-stu-id="a241e-186">Child elements</span></span>

|  <span data-ttu-id="a241e-187">Element</span><span class="sxs-lookup"><span data-stu-id="a241e-187">Element</span></span> |  <span data-ttu-id="a241e-188">Beschreibung</span><span class="sxs-lookup"><span data-stu-id="a241e-188">Description</span></span>  |
|:-----|:-----|
|  [<span data-ttu-id="a241e-189">OfficeTab</span><span class="sxs-lookup"><span data-stu-id="a241e-189">OfficeTab</span></span>](officetab.md) |  <span data-ttu-id="a241e-190">Fügt die Befehle auf der Registerkarte des Menübands hinzu.</span><span class="sxs-lookup"><span data-stu-id="a241e-190">Adds the command(s) to the default ribbon tab.</span></span>  |
|  [<span data-ttu-id="a241e-191">CustomTab</span><span class="sxs-lookup"><span data-stu-id="a241e-191">CustomTab</span></span>](customtab.md) |  <span data-ttu-id="a241e-192">Fügt die Befehle auf der benutzerdefinierten Registerkarte des Menübands hinzu.</span><span class="sxs-lookup"><span data-stu-id="a241e-192">Adds the command(s) to the custom ribbon tab.</span></span>  |

#### <a name="officetab-example"></a><span data-ttu-id="a241e-193">OfficeTab-Beispiel </span><span class="sxs-lookup"><span data-stu-id="a241e-193">OfficeTab example</span></span>
```xml
<ExtensionPoint xsi:type="MessageReadCommandSurface">
  <OfficeTab id="TabDefault">
        <-- OfficeTab Definition -->
  </OfficeTab>
</ExtensionPoint>
```

#### <a name="customtab-example"></a><span data-ttu-id="a241e-194">CustomTab-Beispiel</span><span class="sxs-lookup"><span data-stu-id="a241e-194">CustomTab example</span></span>
```xml
<ExtensionPoint xsi:type="MessageReadCommandSurface">
  <CustomTab id="TabCustom1">
        <-- CustomTab Definition -->
  </CustomTab>
</ExtensionPoint>
```

### <a name="messagecomposecommandsurface"></a><span data-ttu-id="a241e-195">MessageComposeCommandSurface</span><span class="sxs-lookup"><span data-stu-id="a241e-195">MessageComposeCommandSurface</span></span>
<span data-ttu-id="a241e-196">Dieser Erweiterungspunkt platziert Schaltflächen für Add-Ins, die Mailformulare zum Verfassen verwenden, im Menüband.</span><span class="sxs-lookup"><span data-stu-id="a241e-196">This extension point puts buttons on the ribbon for add-ins using mail compose form.</span></span> 

#### <a name="child-elements"></a><span data-ttu-id="a241e-197">Untergeordnete Elemente</span><span class="sxs-lookup"><span data-stu-id="a241e-197">Child elements</span></span>

|  <span data-ttu-id="a241e-198">Element</span><span class="sxs-lookup"><span data-stu-id="a241e-198">Element</span></span> |  <span data-ttu-id="a241e-199">Beschreibung</span><span class="sxs-lookup"><span data-stu-id="a241e-199">Description</span></span>  |
|:-----|:-----|
|  [<span data-ttu-id="a241e-200">OfficeTab</span><span class="sxs-lookup"><span data-stu-id="a241e-200">OfficeTab</span></span>](officetab.md) |  <span data-ttu-id="a241e-201">Fügt die Befehle auf der Registerkarte des Menübands hinzu.</span><span class="sxs-lookup"><span data-stu-id="a241e-201">Adds the command(s) to the default ribbon tab.</span></span>  |
|  [<span data-ttu-id="a241e-202">CustomTab</span><span class="sxs-lookup"><span data-stu-id="a241e-202">CustomTab</span></span>](customtab.md) |  <span data-ttu-id="a241e-203">Fügt die Befehle auf der benutzerdefinierten Registerkarte des Menübands hinzu.</span><span class="sxs-lookup"><span data-stu-id="a241e-203">Adds the command(s) to the custom ribbon tab.</span></span>  |

#### <a name="officetab-example"></a><span data-ttu-id="a241e-204">OfficeTab-Beispiel </span><span class="sxs-lookup"><span data-stu-id="a241e-204">OfficeTab example</span></span>
```xml
<ExtensionPoint xsi:type="MessageComposeCommandSurface">
  <OfficeTab id="TabDefault">
        <-- OfficeTab Definition -->
  </OfficeTab>
</ExtensionPoint>
```

#### <a name="customtab-example"></a><span data-ttu-id="a241e-205">CustomTab-Beispiel</span><span class="sxs-lookup"><span data-stu-id="a241e-205">CustomTab example</span></span>

```xml
<ExtensionPoint xsi:type="MessageComposeCommandSurface">
  <CustomTab id="TabCustom1">
        <-- CustomTab Definition -->
  </CustomTab>
</ExtensionPoint>
```

### <a name="appointmentorganizercommandsurface"></a><span data-ttu-id="a241e-206">AppointmentOrganizerCommandSurface</span><span class="sxs-lookup"><span data-stu-id="a241e-206">AppointmentOrganizerCommandSurface</span></span>

<span data-ttu-id="a241e-207">Dieser Erweiterungspunkt platziert Schaltflächen für das Formular, das dem Organisator der Besprechung angezeigt wird, im Menüband.</span><span class="sxs-lookup"><span data-stu-id="a241e-207">This extension point puts buttons on the ribbon for the form that's displayed to the organizer of the meeting.</span></span> 

#### <a name="child-elements"></a><span data-ttu-id="a241e-208">Untergeordnete Elemente</span><span class="sxs-lookup"><span data-stu-id="a241e-208">Child elements</span></span>

|  <span data-ttu-id="a241e-209">Element</span><span class="sxs-lookup"><span data-stu-id="a241e-209">Element</span></span> |  <span data-ttu-id="a241e-210">Beschreibung</span><span class="sxs-lookup"><span data-stu-id="a241e-210">Description</span></span>  |
|:-----|:-----|
|  [<span data-ttu-id="a241e-211">OfficeTab</span><span class="sxs-lookup"><span data-stu-id="a241e-211">OfficeTab</span></span>](officetab.md) |  <span data-ttu-id="a241e-212">Fügt die Befehle auf der Registerkarte des Menübands hinzu.</span><span class="sxs-lookup"><span data-stu-id="a241e-212">Adds the command(s) to the default ribbon tab.</span></span>  |
|  [<span data-ttu-id="a241e-213">CustomTab</span><span class="sxs-lookup"><span data-stu-id="a241e-213">CustomTab</span></span>](customtab.md) |  <span data-ttu-id="a241e-214">Fügt die Befehle auf der benutzerdefinierten Registerkarte des Menübands hinzu.</span><span class="sxs-lookup"><span data-stu-id="a241e-214">Adds the command(s) to the custom ribbon tab.</span></span>  |

#### <a name="officetab-example"></a><span data-ttu-id="a241e-215">OfficeTab-Beispiel </span><span class="sxs-lookup"><span data-stu-id="a241e-215">OfficeTab example</span></span>
```xml
<ExtensionPoint xsi:type="AppointmentOrganizerCommandSurface">
  <OfficeTab id="TabDefault">
        <-- OfficeTab Definition -->
  </OfficeTab>
</ExtensionPoint>
```

#### <a name="customtab-example"></a><span data-ttu-id="a241e-216">CustomTab-Beispiel</span><span class="sxs-lookup"><span data-stu-id="a241e-216">CustomTab example</span></span>
```xml
<ExtensionPoint xsi:type="AppointmentOrganizerCommandSurface">
  <CustomTab id="TabCustom1">
        <-- CustomTab Definition -->
  </CustomTab>
</ExtensionPoint>
```

### <a name="appointmentattendeecommandsurface"></a><span data-ttu-id="a241e-217">AppointmentAttendeeCommandSurface</span><span class="sxs-lookup"><span data-stu-id="a241e-217">AppointmentAttendeeCommandSurface</span></span>

<span data-ttu-id="a241e-218">Dieser Erweiterungspunkt platziert Schaltflächen für das Formular, das dem Teilnehmer der Besprechung angezeigt wird, im Menüband.</span><span class="sxs-lookup"><span data-stu-id="a241e-218">This extension point puts buttons on the ribbon for the form that's displayed to the attendee of the meeting.</span></span> 

#### <a name="child-elements"></a><span data-ttu-id="a241e-219">Untergeordnete Elemente</span><span class="sxs-lookup"><span data-stu-id="a241e-219">Child elements</span></span>

|  <span data-ttu-id="a241e-220">Element</span><span class="sxs-lookup"><span data-stu-id="a241e-220">Element</span></span> |  <span data-ttu-id="a241e-221">Beschreibung</span><span class="sxs-lookup"><span data-stu-id="a241e-221">Description</span></span>  |
|:-----|:-----|
|  [<span data-ttu-id="a241e-222">OfficeTab</span><span class="sxs-lookup"><span data-stu-id="a241e-222">OfficeTab</span></span>](officetab.md) |  <span data-ttu-id="a241e-223">Fügt die Befehle auf der Registerkarte des Menübands hinzu.</span><span class="sxs-lookup"><span data-stu-id="a241e-223">Adds the command(s) to the default ribbon tab.</span></span>  |
|  [<span data-ttu-id="a241e-224">CustomTab</span><span class="sxs-lookup"><span data-stu-id="a241e-224">CustomTab</span></span>](customtab.md) |  <span data-ttu-id="a241e-225">Fügt die Befehle auf der benutzerdefinierten Registerkarte des Menübands hinzu.</span><span class="sxs-lookup"><span data-stu-id="a241e-225">Adds the command(s) to the custom ribbon tab.</span></span>  |

#### <a name="officetab-example"></a><span data-ttu-id="a241e-226">OfficeTab-Beispiel </span><span class="sxs-lookup"><span data-stu-id="a241e-226">OfficeTab example</span></span>
```xml
<ExtensionPoint xsi:type="AppointmentAttendeeCommandSurface">
  <OfficeTab id="TabDefault">
        <-- OfficeTab Definition -->
  </OfficeTab>
</ExtensionPoint>
```

#### <a name="customtab-example"></a><span data-ttu-id="a241e-227">CustomTab-Beispiel</span><span class="sxs-lookup"><span data-stu-id="a241e-227">CustomTab example</span></span>
```xml
<ExtensionPoint xsi:type="AppointmentAttendeeCommandSurface">
  <CustomTab id="TabCustom1">
        <-- CustomTab Definition -->
  </CustomTab>
</ExtensionPoint>
```

### <a name="module"></a><span data-ttu-id="a241e-228">Module</span><span class="sxs-lookup"><span data-stu-id="a241e-228">Module</span></span>

<span data-ttu-id="a241e-229">Dieser Erweiterungspunkt platziert Schaltflächen für die Modulerweiterung im Menüband.</span><span class="sxs-lookup"><span data-stu-id="a241e-229">This extension point puts buttons on the ribbon for the module extension.</span></span> 

#### <a name="child-elements"></a><span data-ttu-id="a241e-230">Untergeordnete Elemente</span><span class="sxs-lookup"><span data-stu-id="a241e-230">Child elements</span></span>

|  <span data-ttu-id="a241e-231">Element</span><span class="sxs-lookup"><span data-stu-id="a241e-231">Element</span></span> |  <span data-ttu-id="a241e-232">Beschreibung</span><span class="sxs-lookup"><span data-stu-id="a241e-232">Description</span></span>  |
|:-----|:-----|
|  [<span data-ttu-id="a241e-233">OfficeTab</span><span class="sxs-lookup"><span data-stu-id="a241e-233">OfficeTab</span></span>](officetab.md) |  <span data-ttu-id="a241e-234">Fügt die Befehle auf der Registerkarte des Menübands hinzu.</span><span class="sxs-lookup"><span data-stu-id="a241e-234">Adds the command(s) to the default ribbon tab.</span></span>  |
|  [<span data-ttu-id="a241e-235">CustomTab</span><span class="sxs-lookup"><span data-stu-id="a241e-235">CustomTab</span></span>](customtab.md) |  <span data-ttu-id="a241e-236">Fügt die Befehle auf der benutzerdefinierten Registerkarte des Menübands hinzu.</span><span class="sxs-lookup"><span data-stu-id="a241e-236">Adds the command(s) to the custom ribbon tab.</span></span>  |

### <a name="mobilemessagereadcommandsurface"></a><span data-ttu-id="a241e-237">MobileMessageReadCommandSurface</span><span class="sxs-lookup"><span data-stu-id="a241e-237">MobileMessageReadCommandSurface</span></span>
<span data-ttu-id="a241e-238">Mit diesem Erweiterungspunkt werden Schaltflächen für die Ansicht gelesener Mails auf der Befehlsoberfläche in dem mobilen Formfaktor platziert.</span><span class="sxs-lookup"><span data-stu-id="a241e-238">This extension point puts buttons in the command surface for the mail read view in the mobile form factor.</span></span>

#### <a name="child-elements"></a><span data-ttu-id="a241e-239">Untergeordnete Elemente</span><span class="sxs-lookup"><span data-stu-id="a241e-239">Child elements</span></span>

|  <span data-ttu-id="a241e-240">Element</span><span class="sxs-lookup"><span data-stu-id="a241e-240">Element</span></span> |  <span data-ttu-id="a241e-241">Beschreibung</span><span class="sxs-lookup"><span data-stu-id="a241e-241">Description</span></span>  |
|:-----|:-----|
|  [<span data-ttu-id="a241e-242">Group</span><span class="sxs-lookup"><span data-stu-id="a241e-242">Group</span></span>](group.md) |  <span data-ttu-id="a241e-243">Fügt eine Gruppe von Schaltflächen zu der Oberfläche mit Befehlen.</span><span class="sxs-lookup"><span data-stu-id="a241e-243">Adds a group of buttons to the command surface.</span></span>  |

<span data-ttu-id="a241e-244">**ExtensionPoint**-Elemente dieses Typs können nur ein untergeordnetes Element enthalten, ein **Group**-Element.</span><span class="sxs-lookup"><span data-stu-id="a241e-244">**ExtensionPoint** elements of this type can only have one child element: a **Group** element.</span></span>

<span data-ttu-id="a241e-245">Für **Control**-Elemente, die in diesem Erweiterungspunkt enthalten sind, muss das **xsi:type**-Attribut auf `MobileButton` festgelegt sein.</span><span class="sxs-lookup"><span data-stu-id="a241e-245">**Control** elements contained in this extension point must have the **xsi:type** attribute set to `MobileButton`.</span></span>

#### <a name="example"></a><span data-ttu-id="a241e-246">Beispiel</span><span class="sxs-lookup"><span data-stu-id="a241e-246">Example</span></span>
```xml
<ExtensionPoint xsi:type="MobileMessageReadCommandSurface">
  <Group id="mobileGroupID">
    <Label resid="residAppName"/>
      <Control id="mobileButton1" xsi:type="MobileButton">
        <!-- Control definition -->
      </Control>
  </Group>
</ExtensionPoint>
```

### <a name="events"></a><span data-ttu-id="a241e-247">Ereignisse</span><span class="sxs-lookup"><span data-stu-id="a241e-247">Events</span></span>

<span data-ttu-id="a241e-248">Dieser Erweiterungspunkt fügt einen Ereignishandler für ein spezifisches Ereignis hinzu.</span><span class="sxs-lookup"><span data-stu-id="a241e-248">This extension point adds an event handler for a specified event.</span></span>

> [!NOTE]
> <span data-ttu-id="a241e-249">Dieser Elementtyp wird nur von Outlook im Web in Office 365 unterstützt.</span><span class="sxs-lookup"><span data-stu-id="a241e-249">This element type is only supported by Outlook on the web in Office 365.</span></span>

| <span data-ttu-id="a241e-250">Element</span><span class="sxs-lookup"><span data-stu-id="a241e-250">Element</span></span> | <span data-ttu-id="a241e-251">Beschreibung</span><span class="sxs-lookup"><span data-stu-id="a241e-251">Description</span></span>  |
|:-----|:-----|
|  [<span data-ttu-id="a241e-252">Event</span><span class="sxs-lookup"><span data-stu-id="a241e-252">Event</span></span>](event.md) |  <span data-ttu-id="a241e-253">Gibt das Ereignis und die Ereignishandlerfunktion an.</span><span class="sxs-lookup"><span data-stu-id="a241e-253">Specifies the event and event handler function.</span></span>  |

#### <a name="itemsend-event-example"></a><span data-ttu-id="a241e-254">Beispiel für ein ItemSend-Ereignis</span><span class="sxs-lookup"><span data-stu-id="a241e-254">ItemSend event example</span></span>

```xml
<ExtensionPoint xsi:type="Events"> 
  <Event Type="ItemSend" FunctionExecution="synchronous" FunctionName="itemSendHandler" /> 
</ExtensionPoint> 
```

### <a name="detectedentity"></a><span data-ttu-id="a241e-255">DetectedEntity</span><span class="sxs-lookup"><span data-stu-id="a241e-255">DetectedEntity</span></span>

<span data-ttu-id="a241e-256">Dieser Erweiterungspunkt fügt eine Kontext-Add-In-Aktivierung für einen angegebenen Entitätstyp hinzu.</span><span class="sxs-lookup"><span data-stu-id="a241e-256">This extension point adds a contextual add-in activation on a specified entity type.</span></span>

<span data-ttu-id="a241e-257">Das enthaltende [VersionOverrides](versionoverrides.md)-Element muss einen `xsi:type`-Attributwert von `VersionOverridesV1_1` aufweisen.</span><span class="sxs-lookup"><span data-stu-id="a241e-257">The containing [VersionOverrides](versionoverrides.md) element must have an `xsi:type` attribute value of `VersionOverridesV1_1`.</span></span>

> [!NOTE]
> <span data-ttu-id="a241e-258">Dieser Elementtyp wird nur von Outlook im Web in Office 365 unterstützt.</span><span class="sxs-lookup"><span data-stu-id="a241e-258">This element type is only supported by Outlook on the web in Office 365.</span></span>

|  <span data-ttu-id="a241e-259">Element</span><span class="sxs-lookup"><span data-stu-id="a241e-259">Element</span></span> |  <span data-ttu-id="a241e-260">Beschreibung</span><span class="sxs-lookup"><span data-stu-id="a241e-260">Description</span></span>  |
|:-----|:-----|
|  [<span data-ttu-id="a241e-261">Label</span><span class="sxs-lookup"><span data-stu-id="a241e-261">Label</span></span>](#label) |  <span data-ttu-id="a241e-262">Gibt die Bezeichnung für das Add-In im Kontextfenster an.</span><span class="sxs-lookup"><span data-stu-id="a241e-262">Specifies the label for the add-in in the contextual window.</span></span>  |
|  [<span data-ttu-id="a241e-263">SourceLocation</span><span class="sxs-lookup"><span data-stu-id="a241e-263">SourceLocation</span></span>](sourcelocation.md) |  <span data-ttu-id="a241e-264">Gibt die URL für das Kontextfenster an.</span><span class="sxs-lookup"><span data-stu-id="a241e-264">Specifies the URL for the contextual window.</span></span>  |
|  [<span data-ttu-id="a241e-265">Rule</span><span class="sxs-lookup"><span data-stu-id="a241e-265">Rule</span></span>](rule.md) |  <span data-ttu-id="a241e-266">Gibt die Regel(n) an, die bestimmen, wann ein Add-In aktiviert wird.</span><span class="sxs-lookup"><span data-stu-id="a241e-266">Specifies the rule or rules that determine when an add-in activates.</span></span>  |

#### <a name="label"></a><span data-ttu-id="a241e-267">Label</span><span class="sxs-lookup"><span data-stu-id="a241e-267">Label</span></span>

<span data-ttu-id="a241e-p115">Erforderlich. Die Beschriftung der Gruppe. Das Attribut  **resid** muss auf den Wert des Attributs **id** eines **String**-Elements im  **ShortStrings**-Element im  [Resources](resources.md) -Element festgelegt werden.</span><span class="sxs-lookup"><span data-stu-id="a241e-p115">Required. The label of the group. The  **resid** attribute must be set to the value of the **id** attribute of a **String** element in the **ShortStrings** element in the [Resources](resources.md) element.</span></span>

#### <a name="highlight-requirements"></a><span data-ttu-id="a241e-271">Hervorhebungsanforderungen</span><span class="sxs-lookup"><span data-stu-id="a241e-271">Highlight requirements</span></span>

<span data-ttu-id="a241e-p116">Ein Benutzer kann ein Kontext-Add-In nur durch Interaktion mit einer hervorgehobenen Entität aktivieren. Mithilfe des Attributs `Highlight` des `Rule`-Elements für die Regeltypen `ItemHasKnownEntity` und `ItemHasRegularExpressionMatch` können Entwickler steuern, welche Entitäten hervorgehoben werden.</span><span class="sxs-lookup"><span data-stu-id="a241e-p116">The only way a user can activate a contextual add-in is to interact with a highlighted entity. Developers can control which entities are highlighted by using the `Highlight` attribute of the `Rule` element for `ItemHasKnownEntity` and `ItemHasRegularExpressionMatch` rule types.</span></span>

<span data-ttu-id="a241e-p117">Sie müssen jedoch einige Einschränkungen beachten. Mit diesen Einschränkungen soll sichergestellt werden, dass in anwendbaren Nachrichten oder Terminen immer eine hervorgehobene Entität vorhanden ist, damit der Benutzer das Add-In aktivieren kann.</span><span class="sxs-lookup"><span data-stu-id="a241e-p117">However, there are some limitations to be aware of. These limitations are in place to ensure that there will always be a highlighted entity in applicable messages or appointments to give the user a way to activate the add-in.</span></span>

- <span data-ttu-id="a241e-276">Die Entitätstypen `EmailAddress` und `Url` können nicht hervorgehoben und daher nicht verwendet werden, um ein Add-In zu aktivieren.</span><span class="sxs-lookup"><span data-stu-id="a241e-276">The `EmailAddress` and `Url` entity types cannot be highlighted, and therefore cannot be used to activate an add-in.</span></span>
- <span data-ttu-id="a241e-277">Wenn Sie eine einzelne Regel verwenden, MUSS `Highlight` auf `all` festgelegt werden.</span><span class="sxs-lookup"><span data-stu-id="a241e-277">If using a single rule, `Highlight` MUST be set to `all`.</span></span>
- <span data-ttu-id="a241e-278">Wenn Sie einen `RuleCollection`-Regeltyp mit `Mode="AND"` zum Kombinieren mehrerer Regeln verwenden, MUSS `Highlight` für mindestens eine der Regeln auf `all` festgelegt werden.</span><span class="sxs-lookup"><span data-stu-id="a241e-278">If using a `RuleCollection` rule type with `Mode="AND"` to combine multiple rules, at least one of the rules MUST have `Highlight` set to `all`.</span></span>
- <span data-ttu-id="a241e-279">Wenn Sie einen `RuleCollection`-Regeltyp mit `Mode="OR"` zum Kombinieren mehrerer Regeln verwenden, MUSS `Highlight` für alle Regeln auf `all` festgelegt werden.</span><span class="sxs-lookup"><span data-stu-id="a241e-279">If using a `RuleCollection` rule type with `Mode="OR"` to combine multiple rules, all of the rules MUST have `Highlight` set to `all`.</span></span>

#### <a name="detectedentity-event-example"></a><span data-ttu-id="a241e-280">Beispiel für DetectedEntity-Ereignis</span><span class="sxs-lookup"><span data-stu-id="a241e-280">DetectedEntity event example</span></span>

```xml
<ExtensionPoint xsi:type="DetectedEntity">
  <Label resid="residLabelName"/>
  <SourceLocation resid="residDetectedEntityURL" />
  <Rule xsi:type="RuleCollection" Mode="And">
    <Rule xsi:type="ItemIs" ItemType="Message" />
    <Rule xsi:type="ItemHasKnownEntity" EntityType="MeetingSuggestion" Highlight="all" />
    <Rule xsi:type="ItemHasKnownEntity" EntityType="Address" Highlight="none" />
  </Rule>
</ExtensionPoint> 
```