# <a name="officemenu-element"></a>OfficeMenu-Element

Definiert eine Sammlung von Steuerelementen, die dem Office-Kontextmenü hinzugefügt werden sollen. Gilt für Word-, Excel-, PowerPoint- und OneNote-Add-Ins.

## <a name="attributes"></a>Attribute

| Attribut            | Erforderlich | Beschreibung                          |
|:---------------------|:--------:|:-------------------------------------|
| [xsi:type](#xsitype) | Ja      | Der Typ des OfficeMenu-Elements, das definiert wird.|

## <a name="child-elements"></a>Untergeordnete Elemente

|  Element |  Erforderlich  |  Beschreibung  |
|:-----|:-----|:-----|
|  [Control](#control)    | Ja |  Eine Auflistung eines oder mehrerer Control-Objekte.  |

## <a name="xsitype"></a>xsi:type

Gibt ein integriertes Menü der Office-Clientanwendung an, zu dem dieses Office-Add-in hinzugefügt werden soll.

- `ContextMenuText` -  Zeigt das Element im Kontextmenü an, wenn Text ausgewählt wird und der Benutzer dann mit der rechten Maustaste auf den ausgewählten Text klickt. Gilt für Word, Excel, PowerPoint und OneNote.
- `ContextMenuCell` - Zeigt das Element im Kontextmenü an, wenn der Benutzer mit der rechten Maustaste in eine Zelle in dem Arbeitsblatt klickt. Gilt für Excel. 

## <a name="control"></a>Steuerelement

Jedes **OfficeMenu**-Element erfordert mindestens ein oder mehrere [menu](control.md#menu-dropdown-button-controls)- Steuerelemente. 

## <a name="example"></a>Beispiel

```xml
<OfficeMenu id="ContextMenuCell">
    <Control xsi:type="Menu" id="myMenuID">
      <Label resid="residLabel3" />
      <Supertip>
          <Title resid="residLabel" />
          <Description resid="residToolTip" />
      </Supertip>   
      <Icon>
        <bt:Image size="16" resid="icon1_16x16" />
        <bt:Image size="32" resid="icon1_32x32" />
        <bt:Image size="80" resid="icon1_80x80" />
      </Icon>    
      <Items>
        <Item id="myMenuItemID">
          <Label resid="residLabel3"/>
          <Supertip>
            <Title resid="residLabel" />
            <Description resid="residToolTip" />
          </Supertip>
          <Icon>
            <bt:Image size="16" resid="icon1_16x16" />
            <bt:Image size="32" resid="icon1_32x32" />
            <bt:Image size="80" resid="icon1_80x80" />
          </Icon>    
          <Action xsi:type="ShowTaskpane">
            <SourceLocation resid="residTaskpaneUrl2" />    
          </Action>    
        </Item>
      </Items>
    </Control>   
</OfficeMenu>
```
