# <a name="extensionpoint-element"></a>ExtensionPoint-Element

 Definiert, wo ein Add-In Funktionen in der Office-Benutzeroberfläche verfügbar macht. Das **ExtensionPoint**-Element ist ein untergeordnetes Element von [AllFormFactor](allformfactors.md), [DesktopFormFactor](desktopformfactor.md) oder [MobileFormFactor](mobileformfactor.md). 

## <a name="attributes"></a>Attribute

|  Attribut  |  Erforderlich  |  Beschreibung  |
|:-----|:-----|:-----|
|  **xsi:type**  |  Ja  | Der Typ des Erweiterungspunkts, der definiert wird.|

## <a name="extension-points-for-excel-only"></a>Nur für Excel verfügbare Erweiterungspunkte

- **CustomFunctions** - eine benutzerdefinierte Funktion in JavaScript für Excel geschrieben.

[Dieses XML-Codebeispiel](https://github.com/OfficeDev/Excel-Custom-Functions/blob/master/customfunctions.xml) veranschaulicht, wie das **ExtensionPoint** -Element mit der Attributwert **CustomFunctions** und die untergeordneten Elemente verwenden, um verwendet werden.

## <a name="extension-points-for-word-excel-powerpoint-and-onenote-add-in-commands"></a>Erweiterungspunkte für Word-, Excel-, PowerPoint und OneNote-Add-In-Befehle

- **PrimaryCommandSurface** Das Menüband in Office.
- **ContextMenu** Das Kontextmenü, das angezeigt wird, wenn Sie mit der rechten Maustaste auf die Office-Benutzeroberfläche klicken.

In den folgenden Beispielen wird gezeigt, wie Sie das  **ExtensionPoint**-Element mit  **PrimaryCommandSurface**- und  **ContextMenu**-Attributwerten verwenden, und Sie sehen die untergeordneten Elemente, die für jedes Element verwendet werden sollten.

> [!IMPORTANT] 
> Stellen Sie für Elemente, die das ID-Attribut enthalten sicher, dass Sie eine eindeutige ID bereitstellen Es wird empfohlen, den Namen Ihres Unternehmens zusammen mit Ihrer ID zu verwenden. Verwenden Sie beispielsweise das folgende Format. <CustomTab id="mycompanyname.mygroupname">

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

#### <a name="child-elements"></a>Untergeordnete Elemente
 
|**Element**|**Beschreibung**|
|:-----|:-----|
|**CustomTab**|Erforderlich, wenn Sie eine benutzerdefinierte Registerkarte (unter Verwendung von  **PrimaryCommandSurface**) zum Menüband hinzufügen möchten. Wenn Sie das  **CustomTab**-Element verwenden, können Sie das  **OfficeTab**-Element nicht verwenden. Das  **ID** -Attribut ist erforderlich.|
|**OfficeTab**|Erforderlich, wenn Sie eine standardmäßige Office-Menübandregisterkarte (unter Verwendung von **PrimaryCommandSurface**) erweitern möchten. Wenn Sie das **OfficeTab**-Element verwenden, können Sie das **CustomTab**-Element nicht verwenden. Weitere Informationen finden Sie unter [OfficeTab](officetab.md)|
|**OfficeMenu**|Erforderlich, wenn Sie Add-In-Befehle zu einem Standardkontextmenü (unter Verwendung von **ContextMenu**) hinzufügen. Das **id**-Attribut muss festgelegt werden auf: <br/> - **ContextMenuText** für Excel oder Word. Zeigt das Element im Kontextmenü an, wenn Text ausgewählt wird und der Benutzer dann mit der rechten Maustaste auf den ausgewählten Text klickt. <br/> - **ContextMenuCell** für Excel. Zeigt das Element im Kontextmenü an, wenn der Benutzer mit der rechten Maustaste in eine Zelle in dem Arbeitsblatt klickt.|
|**Group**|Eine Gruppe von Erweiterungspunkten der Benutzeroberfläche auf einer Registerkarte.Eine Gruppe kann über bis zu sechs Steuerelemente verfügen. Das **id** -Attribut ist erforderlich. Es handelt sich um eine Zeichenfolge mit maximal 125 Zeichen.|
|**Label**|Erforderlich. Die Bezeichnung der Gruppe. Das  **resid** -Attribut muss auf den Wert des **id**-Attributs eines  **String**-Elements festegelegt werden. Das  **String**-Element ist ein untergeordnetes Element vom  **ShortStrings**-Element, das ein untergeordnetes Element vom  **Resources** -Element ist.|
|**Icon**|Erforderlich. Gibt das Gruppensymbol an, das auf kleinen Geräten verwendet wird, oder wenn zu viele Schaltflächen angezeigt werden. Das  **resid** -Attribut muss auf den Wert des **id**-Attributs eines  **Image**-Elements festgelegt werden. Das  **Image**-Element ist ein untergeordnetes Element vom  **Images**-Element, das ein untergeordnetes Element vom  **Resources** -Element ist. Das **size**-Attribut gibt die Größe des Bilds in Pixeln an. Es sind drei Bildgrößen erforderlich: 16, 32 und 80. Es werden auch fünf optionale Größen unterstützt: 20, 24, 40, 48 und 64.|
|**Tooltip**|Optional. Die QuickInfo der Gruppe. Das  **resid** -Attribut muss auf den Wert des **id**-Attributs eines  **String**-Elements festgelegt werden. Das  **String**-Element ist ein untergeordnetes Element vom  **LongStrings**-Element, das ein untergeordnetes Element vom  **Resources** -Element ist.|
|**Control**|Jede Gruppe erfordert mindestens ein Steuerelement. Ein **Control**-Element kann entweder ein **Button** oder **Menu** sein. Verwenden Sie **Menu**, um eine Dropdownliste mit Steuerelementen für Schaltflächen anzugeben. Derzeit werden nur Schaltflächen und Menüs unterstützt. Weitere Informationen finden Sie in den Abschnitten [Steuerelemente für Schaltflächen](control.md#button-control) und [Menüsteuerelemente](control.md#menu-dropdown-button-controls).<br/>**Hinweis:**  So machen Sie die Problembehandlung, wird empfohlen, ein **Control** -Element und die zugehörigen **Ressourcen** untergeordneten Elemente jeweils einzeln hinzugefügt werden.|
|**Script**|Enthält Links zur JavaScript-Datei mit der benutzerdefinierten Funktionsdefinition und dem Registrierungscode. Dieses Element wird nicht in der Vorschau für Entwickler verwendet. Stattdessen werden alle JavaScript-Dateien über die HTML-Seite geladen.|
|**Page**|Enthält Links zur HTML-Seite für Ihre benutzerdefinierten Funktionen.|

## <a name="extension-points-for-outlook"></a>Erweiterungspunkte für Outlook

- [MessageReadCommandSurface](#messagereadcommandsurface) 
- [MessageComposeCommandSurface](#messagecomposecommandsurface) 
- [AppointmentOrganizerCommandSurface](#appointmentorganizercommandsurface) 
- [AppointmentAttendeeCommandSurface](#appointmentattendeecommandsurface)
- [Module](#module) (Kann nur im [DesktopFormFactor](desktopformfactor.md) verwendet werden.)
- [MobileMessageReadCommandSurface](#mobilemessagereadcommandsurface)
- [Events](#events)
- [DetectedEntity](#detectedentity)

### <a name="messagereadcommandsurface"></a>MessageReadCommandSurface
Mit diesem Erweiterungspunkt werden Schaltflächen für die Ansicht gelesener Mails auf der Befehlsoberfläche platziert. In Outlook Desktop wird das Element im Menüband angezeigt.

#### <a name="child-elements"></a>Untergeordnete Elemente

|  Element |  Beschreibung  |
|:-----|:-----|
|  [OfficeTab](officetab.md) |  Fügt die Befehle auf der Registerkarte des Menübands hinzu.  |
|  [CustomTab](customtab.md) |  Fügt die Befehle auf der benutzerdefinierten Registerkarte des Menübands hinzu.  |

#### <a name="officetab-example"></a>OfficeTab-Beispiel 
```xml
<ExtensionPoint xsi:type="MessageReadCommandSurface">
  <OfficeTab id="TabDefault">
        <-- OfficeTab Definition -->
  </OfficeTab>
</ExtensionPoint>
```

#### <a name="customtab-example"></a>CustomTab-Beispiel
```xml
<ExtensionPoint xsi:type="MessageReadCommandSurface">
  <CustomTab id="TabCustom1">
        <-- CustomTab Definition -->
  </CustomTab>
</ExtensionPoint>
```

### <a name="messagecomposecommandsurface"></a>MessageComposeCommandSurface
Dieser Erweiterungspunkt platziert Schaltflächen für Add-Ins, die Mailformulare zum Verfassen verwenden, im Menüband. 

#### <a name="child-elements"></a>Untergeordnete Elemente

|  Element |  Beschreibung  |
|:-----|:-----|
|  [OfficeTab](officetab.md) |  Fügt die Befehle auf der Registerkarte des Menübands hinzu.  |
|  [CustomTab](customtab.md) |  Fügt die Befehle auf der benutzerdefinierten Registerkarte des Menübands hinzu.  |

#### <a name="officetab-example"></a>OfficeTab-Beispiel 
```xml
<ExtensionPoint xsi:type="MessageComposeCommandSurface">
  <OfficeTab id="TabDefault">
        <-- OfficeTab Definition -->
  </OfficeTab>
</ExtensionPoint>
```

#### <a name="customtab-example"></a>CustomTab-Beispiel

```xml
<ExtensionPoint xsi:type="MessageComposeCommandSurface">
  <CustomTab id="TabCustom1">
        <-- CustomTab Definition -->
  </CustomTab>
</ExtensionPoint>
```

### <a name="appointmentorganizercommandsurface"></a>AppointmentOrganizerCommandSurface

Dieser Erweiterungspunkt platziert Schaltflächen für das Formular, das dem Organisator der Besprechung angezeigt wird, im Menüband. 

#### <a name="child-elements"></a>Untergeordnete Elemente

|  Element |  Beschreibung  |
|:-----|:-----|
|  [OfficeTab](officetab.md) |  Fügt die Befehle auf der Registerkarte des Menübands hinzu.  |
|  [CustomTab](customtab.md) |  Fügt die Befehle auf der benutzerdefinierten Registerkarte des Menübands hinzu.  |

#### <a name="officetab-example"></a>OfficeTab-Beispiel 
```xml
<ExtensionPoint xsi:type="AppointmentOrganizerCommandSurface">
  <OfficeTab id="TabDefault">
        <-- OfficeTab Definition -->
  </OfficeTab>
</ExtensionPoint>
```

#### <a name="customtab-example"></a>CustomTab-Beispiel
```xml
<ExtensionPoint xsi:type="AppointmentOrganizerCommandSurface">
  <CustomTab id="TabCustom1">
        <-- CustomTab Definition -->
  </CustomTab>
</ExtensionPoint>
```

### <a name="appointmentattendeecommandsurface"></a>AppointmentAttendeeCommandSurface

Dieser Erweiterungspunkt platziert Schaltflächen für das Formular, das dem Teilnehmer der Besprechung angezeigt wird, im Menüband. 

#### <a name="child-elements"></a>Untergeordnete Elemente

|  Element |  Beschreibung  |
|:-----|:-----|
|  [OfficeTab](officetab.md) |  Fügt die Befehle auf der Registerkarte des Menübands hinzu.  |
|  [CustomTab](customtab.md) |  Fügt die Befehle auf der benutzerdefinierten Registerkarte des Menübands hinzu.  |

#### <a name="officetab-example"></a>OfficeTab-Beispiel 
```xml
<ExtensionPoint xsi:type="AppointmentAttendeeCommandSurface">
  <OfficeTab id="TabDefault">
        <-- OfficeTab Definition -->
  </OfficeTab>
</ExtensionPoint>
```

#### <a name="customtab-example"></a>CustomTab-Beispiel
```xml
<ExtensionPoint xsi:type="AppointmentAttendeeCommandSurface">
  <CustomTab id="TabCustom1">
        <-- CustomTab Definition -->
  </CustomTab>
</ExtensionPoint>
```

### <a name="module"></a>Module

Dieser Erweiterungspunkt platziert Schaltflächen für die Modulerweiterung im Menüband. 

#### <a name="child-elements"></a>Untergeordnete Elemente

|  Element |  Beschreibung  |
|:-----|:-----|
|  [OfficeTab](officetab.md) |  Fügt die Befehle auf der Registerkarte des Menübands hinzu.  |
|  [CustomTab](customtab.md) |  Fügt die Befehle auf der benutzerdefinierten Registerkarte des Menübands hinzu.  |

### <a name="mobilemessagereadcommandsurface"></a>MobileMessageReadCommandSurface
Mit diesem Erweiterungspunkt werden Schaltflächen für die Ansicht gelesener Mails auf der Befehlsoberfläche in dem mobilen Formfaktor platziert.

#### <a name="child-elements"></a>Untergeordnete Elemente

|  Element |  Beschreibung  |
|:-----|:-----|
|  [Group](group.md) |  Fügt eine Gruppe von Schaltflächen zu der Oberfläche mit Befehlen.  |

**ExtensionPoint**-Elemente dieses Typs können nur ein untergeordnetes Element enthalten, ein **Group**-Element.

Für **Control**-Elemente, die in diesem Erweiterungspunkt enthalten sind, muss das **xsi:type**-Attribut auf `MobileButton` festgelegt sein.

#### <a name="example"></a>Beispiel
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

### <a name="events"></a>Ereignisse

Dieser Erweiterungspunkt fügt einen Ereignishandler für ein spezifisches Ereignis hinzu.

> [!NOTE]
> Dieser Elementtyp wird nur von Outlook im Web in Office 365 unterstützt.

| Element | Beschreibung  |
|:-----|:-----|
|  [Event](event.md) |  Gibt das Ereignis und die Ereignishandlerfunktion an.  |

#### <a name="itemsend-event-example"></a>Beispiel für ein ItemSend-Ereignis

```xml
<ExtensionPoint xsi:type="Events"> 
  <Event Type="ItemSend" FunctionExecution="synchronous" FunctionName="itemSendHandler" /> 
</ExtensionPoint> 
```

### <a name="detectedentity"></a>DetectedEntity

Dieser Erweiterungspunkt fügt eine Kontext-Add-In-Aktivierung für einen angegebenen Entitätstyp hinzu.

Das enthaltende [VersionOverrides](versionoverrides.md)-Element muss einen `xsi:type`-Attributwert von `VersionOverridesV1_1` aufweisen.

> [!NOTE]
> Dieser Elementtyp wird nur von Outlook im Web in Office 365 unterstützt.

|  Element |  Beschreibung  |
|:-----|:-----|
|  [Label](#label) |  Gibt die Bezeichnung für das Add-In im Kontextfenster an.  |
|  [SourceLocation](sourcelocation.md) |  Gibt die URL für das Kontextfenster an.  |
|  [Rule](rule.md) |  Gibt die Regel(n) an, die bestimmen, wann ein Add-In aktiviert wird.  |

#### <a name="label"></a>Label

Erforderlich. Die Beschriftung der Gruppe. Das Attribut  **resid** muss auf den Wert des Attributs **id** eines **String**-Elements im  **ShortStrings**-Element im  [Resources](resources.md) -Element festgelegt werden.

#### <a name="highlight-requirements"></a>Hervorhebungsanforderungen

Ein Benutzer kann ein Kontext-Add-In nur durch Interaktion mit einer hervorgehobenen Entität aktivieren. Mithilfe des Attributs `Highlight` des `Rule`-Elements für die Regeltypen `ItemHasKnownEntity` und `ItemHasRegularExpressionMatch` können Entwickler steuern, welche Entitäten hervorgehoben werden.

Sie müssen jedoch einige Einschränkungen beachten. Mit diesen Einschränkungen soll sichergestellt werden, dass in anwendbaren Nachrichten oder Terminen immer eine hervorgehobene Entität vorhanden ist, damit der Benutzer das Add-In aktivieren kann.

- Die Entitätstypen `EmailAddress` und `Url` können nicht hervorgehoben und daher nicht verwendet werden, um ein Add-In zu aktivieren.
- Wenn Sie eine einzelne Regel verwenden, MUSS `Highlight` auf `all` festgelegt werden.
- Wenn Sie einen `RuleCollection`-Regeltyp mit `Mode="AND"` zum Kombinieren mehrerer Regeln verwenden, MUSS `Highlight` für mindestens eine der Regeln auf `all` festgelegt werden.
- Wenn Sie einen `RuleCollection`-Regeltyp mit `Mode="OR"` zum Kombinieren mehrerer Regeln verwenden, MUSS `Highlight` für alle Regeln auf `all` festgelegt werden.

#### <a name="detectedentity-event-example"></a>Beispiel für DetectedEntity-Ereignis

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