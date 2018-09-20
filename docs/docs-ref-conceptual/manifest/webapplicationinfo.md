# <a name="webapplicationinfo-element"></a>Element „WebApplicationInfo“

Dieses Element unterstützt einmaliges Anmelden (Single Sign-On, SSO) in Office-Add-Ins. Es enthält Informationen zu dem Add-In als:

- OAuth 2.0-*Ressource*, für die die Office-Hostanwendung möglicherweise Berechtigungen benötigt
- OAuth 2.0-*Client*, der möglicherweise Berechtigungen für Microsoft Graph benötigt

**WebApplicationInfo** ist ein untergeordnetes Element des Elements [VersionOverrides](versionoverrides.md) im Manifest.  

## <a name="child-elements"></a>Untergeordnete Elemente

|  Element |  Erforderlich  |  Beschreibung  |
|:-----|:-----|:-----|
|  **Id**    |  Ja   |  Die **Anwendungs-ID** des zugeordneten Dienstes des Add-Ins, der im Azure Active Directory v 2.0-Endpunkt registriert ist.|
|  **Ressource**  |  Ja   |  Gibt den **Anwendungs-ID-URI** des Add-Ins an, der im Azure Active Directory v2.2.0-Endpunkt registriert ist.|
|  [Scopes](scopes.md)                |  Nein  |  Gibt die Berechtigungen für Microsoft Graph an, die das Add-In benötigt.  |

> [!NOTE] 
> Derzeit ist es erforderlich, dass Ihr Add-in Ressource Host übereinstimmt. Office fordert kein Token für ein Add-In an, es sein denn, der Besitz wird nachgewiesen. Heutzutage erfolgt dies durch Hosten des Add-Ins unter dem vollqualifizierten Domänennamen der Ressource.

## <a name="webapplicationinfo-example"></a>Beispiel für das Element „WebApplicationInfo“

```xml
<OfficeApp>
...
  <VersionOverrides xmlns="http://schemas.microsoft.com/office/mailappversionoverrides" xsi:type="VersionOverridesV1_0">
    ...
    <WebApplicationInfo>
      <Id>12345678-abcd-1234-efab-123456789abc</Id>
      <Resource>api://myDomain.com/12345678-abcd-1234-efab-123456789abc<Resource>
      <Scopes>
        <Scope>Files.Read.All</Scope>
        <Scope>offline_access</Scope>
        <Scope>openid</Scope>
        <Scope>profile</Scope>        
      </Scopes>
    </WebApplicationInfo>
  </VersionOverrides>
...
</OfficeApp>
```
