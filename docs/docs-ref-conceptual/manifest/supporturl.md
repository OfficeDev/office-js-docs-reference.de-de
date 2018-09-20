# <a name="supporturl-element"></a>SupportUrl-Element

Gibt die URL einer Seite an, die Supportinformationen für Ihr Add-In enthält.

## <a name="syntax"></a>Syntax

```XML
<OfficeApp>
...
  <IconUrl DefaultValue="https://contoso.com/assets/icon-32.png" />
  <HighResolutionIconUrl DefaultValue="https://contoso.com/assets/hi-res-icon.png"/>
  
  
  <SupportUrl DefaultValue="https://contoso.com/support " />
  
  
  <AppDomains>
  ...
  </AppDomains>
...
</OfficeApp>
```

## <a name="contained-in"></a>Enthalten in

[OfficeApp](officeapp.md)

## <a name="can-contain"></a>Kann enthalten

|  Element | Erforderlich | Beschreibung  |
|:-----|:-----|:-----|
|  [Override](override.md)   | Nein | Gibt die Einstellung für zusätzliche Gebietsschema-URLS an. |

## <a name="attributes"></a>Attribute

|**Attribut**|**Typ**|**Erforderlich**|**Beschreibung**|
|:-----|:-----|:-----|:-----|
|DefaultValue|URL|erforderlich|Gibt den Standardwert für diese Einstellung an, der für das im [DefaultLocale](defaultlocale.md)-Element angegebene Gebietsschema ausgedrückt wird.|
