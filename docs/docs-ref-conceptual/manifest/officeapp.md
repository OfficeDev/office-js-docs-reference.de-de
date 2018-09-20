# <a name="officeapp-element"></a>OfficeApp-Element

Das Stammelement im Manifest eines Office-Add-Ins.

**Add-In-Typ:** Inhalt, Aufgabenbereich, E-Mail

## <a name="syntax"></a>Syntax

```XML
<OfficeApp 
  xmlns="http://schemas.microsoft.com/office/appforoffice/1.1" 
  xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance" 
  xsi:type= ["ContentApp" |"MailApp"| "TaskPaneApp"]>
  ...
</OfficeApp>
```

## <a name="contained-in"></a>Enthalten in

 _none_

## <a name="must-contain"></a>Muss enthalten

|**Element**|**Inhalt**|**E-Mail**|**TaskPane**|
|:-----|:-----|:-----|:-----|
|
  [ID](id.md)|x|x|x|
|[Version](version.md)|x|x|x|
|[ProviderName](providername.md)|x|x|x|
|[DefaultLocale](defaultlocale.md)|x|x|x|
|[DefaultSettings](defaultsettings.md)|x||x|
|[Anzeigename](displayname.md)|x|x|x|
|[Beschreibung](description.md)|x|x|x|
|[FormSettings](formsettings.md)||x||
|[Berechtigungen](permissions.md)|x||x|
|[Rule](rule.md)||x||

## <a name="can-contain"></a>Kann enthalten

|**Element**|**Inhalt**|**E-Mail**|**TaskPane**|
|:-----|:-----|:-----|:-----|
|[AlternateId](alternateid.md)|x|x|x|
|[IconUrl](iconurl.md)|x|x|x|
|[HighResolutionIconUrl](highresolutioniconurl.md)|x|x|x|
|[SupportUrl](supporturl.md)|x|x|x|
|[AppDomains](appdomains.md)|x|x|x|
|[Hosts](hosts.md)|x|x|x|
|[Requirements](requirements.md)|x|x|x|
|[AllowSnapshot](allowsnapshot.md)|x|||
|[Berechtigungen](permissions.md)||x||
|[DisableEntityHighlighting](disableentityhighlighting.md)||x||
|[Wörterbuch](dictionary.md)|||x|
|[VersionOverrides](versionoverrides.md)|X|X|X|

## <a name="attributes"></a>Attribute

|||
|:-----|:-----|
|xmlns|Definiert den Office-Add-In-Manifestnamespace und die Schemaversion. Dieses Attribut sollte immer auf `"http://schemas.microsoft.com/office/appforoffice/1.1"` festgelegt werden.|
|xmlns: xsi|Definiert die XML-Schemainstanz. Dieses Attribut sollte immer auf `"http://www.w3.org/2001/XMLSchema-instance"` festgelegt werden.|
|xsi:type|Definiert die Art des Office-Add-Ins. Dieses Attribut sollte immer auf `"ContentApp"`, `"MailApp"` oder `"TaskPaneApp"` festgelegt werden.|
