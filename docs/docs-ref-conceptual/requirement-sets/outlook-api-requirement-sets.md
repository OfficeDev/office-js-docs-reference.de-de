# <a name="outlook-javascript-api-requirement-sets"></a>Outlook-JavaScript-API-anforderungssätzen

Outlook-add-ins deklarieren, welche API-Versionen, die hierfür erforderlich sind, mithilfe von Element " [Requirements](/javascript/office/manifest/requirements) " in ihre [manifest](https://docs.microsoft.com/office/dev/add-ins/develop/add-in-manifests). Outlook-Add-Ins umfassen immer ein [Set](/javascript/office/manifest/set)-Element mit einem `Name`-Attribut, das auf `Mailbox` festgelegt ist, und einem `MinVersion`-Attribut, das auf den mindestens erforderlichen API-Anforderungssatz festgelegt ist, der das Add-In-Szenario unterstützt.

Der folgende Manifestcodeausschnitt gibt z. B. als Mindestanforderungssatz 1.1 an:

```xml
<Requirements>
  <Sets>
    <Set Name="MailBox" MinVersion="1.1" />
  </Sets>
</Requirements>
```

Alle Outlook-APIs gehören zum `Mailbox` [-Anforderungssatz](https://docs.microsoft.com/office/dev/add-ins/develop/specify-office-hosts-and-api-requirements). Der `Mailbox`-Anforderungssatz verfügt über Versionen, und jeder neuen Satz von APIs, den wir veröffentlichen, gehört zu einer höheren Version des Satzes. Nicht alle Outlook-Clients unterstützen den neuesten Satz von APIs, aber wenn ein Outlook-Client Unterstützung für einen Anforderungssatz deklariert, unterstützt er alle APIs in diesem Anforderungssatz.

Das Festlegen einer Mindestanforderungssatzversion im Manifest steuert, in welchem Outlook-Client das Add-In angezeigt wird. Wenn ein Client den Mindestanforderungssatz nicht unterstützt, wird das Add-In nicht geladen. Wenn z. B. Anforderungssatzversion 1.3 angegeben wird, bedeutet dies, dass das Add-In nicht in einem Outlook-Client angezeigt wird, der nicht mindestens 1.3 unterstützt.

## <a name="using-apis-from-later-requirement-sets"></a>Verwenden von APIs aus neueren Anforderungssätzen

Festlegen von keinem Anforderungssatz schränkt nicht verfügbare APIs, die das Add-in verwenden können. Angenommen, wenn das Add-in gibt Anforderung 1.1 festgelegt, aber es in einem Outlook-Client die 1.3 unterstützen ausgeführt wird, kann das Add-in APIs aus Anforderungssatz 1.3 verwenden.

Für die Verwendung neuer APIs können Entwickler nur für ihre Existenz überprüfen Sie mithilfe der Standardmethode JavaScript:

```js
if (item.somePropertyOrFunction !== undefined) {
  item.somePropertyOrFunction ...
}
```

Für APIs, die sich in der im Manifest angegebenen Version des Anforderungssatzes befinden, sind solche Überprüfungen nicht erforderlich.

## <a name="choosing-a-minimum-requirement-set"></a>Auswählen einen Mindestanforderungssatzes

Entwickler sollten den ältesten Anforderungssatz verwenden, der die kritischen APIs für ihr Szenario enthält, ohne die das Add-In nicht funktioniert.

## <a name="clients"></a>Clients

Die folgenden Clients unterstützen Outlook-Add-Ins.

| Client | Unterstützte API-Anforderungssätze |
| --- | --- |
| Outlook 2016 (Klick-und-Los) für Windows | 1.1, 1.2, 1.3, 1.4, 1,5, 1.6, 1.7 |
| Outlook 2016 (MSI) für Windows | 1.1, 1.2, 1.3, 1.4 |
| Outlook 2016 für Mac | 1.1, 1.2, 1.3, 1.4, 1,5, 1,6 |
| Outlook 2013 für Windows | 1.1, 1.2, 1.3, 1.4 |
| Outlook für iPhone | 1.1, 1.2, 1.3, 1.4, 1.5 |
| Outlook für Android | 1.1, 1.2, 1.3, 1.4, 1.5 |
| Outlook im Web (Office 365 und Outlook.com) | 1.1, 1.2, 1.3, 1.4, 1,5, 1,6 |
| Outlook Web App (Exchange 2013 lokal) | 1.1 |
| Outlook Web App (Exchange 2016 lokal) | 1.1, 1.2. 1.3 |

> [!NOTE]
> Unterstützung für 1.3 in Outlook 2013 wurde als Teil der [Dezember 8 2015, aktualisieren für Outlook 2013 (KB3114349)](https://support.microsoft.com/kb/3114349)hinzugefügt. Unterstützung für 1.4 in Outlook 2013 wurde als Teil der [September 13, 2016, aktualisieren für Outlook 2013 (KB3118280)](https://support.microsoft.com/help/3118280)hinzugefügt.
