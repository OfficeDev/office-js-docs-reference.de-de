# <a name="scopes-element"></a>Element „Scopes“

Dieses Element enthält Berechtigungen für Microsoft Graph, die das Add-In benötigt. Es wird vom Office Store verwendet, um ein Zustimmungsdialogfeld zu erstellen. Bei der Installation des Add-Ins aus dem Store wird der Benutzer aufgefordert, dem Add-In die angegebenen Berechtigungen für seine Microsoft Graph-Daten zu erteilen.

## <a name="child-elements"></a>Untergeordnete Elemente

|  Element |  Typ  |  Beschreibung  |
|:-----|:-----|:-----|
|  **Scope**                |  string     |   Der Name einer Berechtigung für Microsoft Graph, zum Beispiel „Files.Read.All“ |

## <a name="example"></a>Beispiel

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
