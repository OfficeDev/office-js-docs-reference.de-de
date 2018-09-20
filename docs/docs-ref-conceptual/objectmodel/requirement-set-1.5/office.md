# <a name="office"></a>Büro

Der Office-Namespace stellt gemeinsam genutzte Oberflächen bereit, die von Add-Ins in allen Office-Apps verwendet werden. Diese Auflistung dokumentiert nur die Schnittstellen, die von Outlook-Add-Ins verwendet werden. Eine vollständige Auflistung des Office-Namespaces finden Sie im Thema zur [freigegebenen API](/javascript/api/office).

##### <a name="requirements"></a>Anforderungen

|Anforderung| Wert|
|---|---|
|[Mindestversion des Postfachanforderungssatzes](/javascript/office/requirement-sets/outlook-api-requirement-sets)| 1.0|
|[Zutreffender Outlook-Modus](https://docs.microsoft.com/outlook/add-ins/#extension-points)| Verfassen oder Lesen|

##### <a name="members-and-methods"></a>Elemente und Methoden

| Element | Typ |
|--------|------|
| [AsyncResultStatus](#asyncresultstatus-string) | Element |
| [CoercionType](#coerciontype-string) | Element |
| [EventType](#eventtype-string) | Element |
| [SourceProperty](#sourceproperty-string) | Element |

### <a name="namespaces"></a>Namespaces

[Kontextmenü](office.context.md): gemeinsam genutzten Oberflächen aus dem Office-Add-ins-API-Kontext-Namespace für die Verwendung in der Outlook-add-in-API bietet.

[MailboxEnums](/javascript/api/outlook/office.mailboxenums.attachmenttype): Enthält die Enumerationen ItemType, EntityType, AttachmentType, RecipientType, ResponseType und ItemNotificationMessageType.

### <a name="members"></a>Elemente

####  <a name="asyncresultstatus-string"></a>AsyncResultStatus :String

Gibt das Ergebnis eines asynchronen Aufrufs an.

##### <a name="type"></a>Typ:

*   Zeichenfolge

##### <a name="properties"></a>Eigenschaften:

|Name| Typ| Beschreibung|
|---|---|---|
|`Succeeded`| Zeichenfolge|Der Aufruf war erfolgreich.|
|`Failed`| Zeichenfolge|Der Aufruf ist fehlerhaft.|

##### <a name="requirements"></a>Anforderungen

|Anforderung| Wert|
|---|---|
|[Mindestversion des Postfachanforderungssatzes](/javascript/office/requirement-sets/outlook-api-requirement-sets)| 1.0|
|[Zutreffender Outlook-Modus](https://docs.microsoft.com/outlook/add-ins/#extension-points)| Verfassen oder Lesen|

---

####  <a name="coerciontype-string"></a>CoercionType :String

Gibt an, wie Daten umgewandelt werden sollen, die von der aufgerufenen Methode zurückgegeben oder festgelegt wurden.

##### <a name="type"></a>Typ:

*   Zeichenfolge

##### <a name="properties"></a>Eigenschaften:

|Name| Typ| Beschreibung|
|---|---|---|
|`Html`| Zeichenfolge|Fordert die Rückgabe der Daten im HTML-Format an.|
|`Text`| Zeichenfolge|Fordert die Rückgabe der Daten im Textformat an.|

##### <a name="requirements"></a>Anforderungen

|Anforderung| Wert|
|---|---|
|[Mindestversion des Postfachanforderungssatzes](/javascript/office/requirement-sets/outlook-api-requirement-sets)| 1.0|
|[Zutreffender Outlook-Modus](https://docs.microsoft.com/outlook/add-ins/#extension-points)| Verfassen oder Lesen|

---

####  <a name="eventtype-string"></a>EventType :String

Gibt das einem Ereignishandler zugeordnete Ereignis an.

##### <a name="type"></a>Typ:

*   Zeichenfolge

##### <a name="properties"></a>Eigenschaften:

| Name | Typ | Beschreibung |
|---|---|---|
|`ItemChanged`| String | Das ausgewählte Element wurde geändert. |

##### <a name="requirements"></a>Anforderungen

|Anforderung| Wert|
|---|---|
|[Mindestversion des Postfachanforderungssatzes](/javascript/office/requirement-sets/outlook-api-requirement-sets)| 1,5 |
|[Zutreffender Outlook-Modus](https://docs.microsoft.com/outlook/add-ins/#extension-points)| Verfassen oder Lesen |

---

####  <a name="sourceproperty-string"></a>SourceProperty :String

Gibt die Quelle der Daten an, die von der aufgerufenen Methode zurückgegeben werden.

##### <a name="type"></a>Typ:

*   Zeichenfolge

##### <a name="properties"></a>Eigenschaften:

|Name| Typ| Beschreibung|
|---|---|---|
|`Body`| Zeichenfolge|Die Quelle der Daten stammt aus dem Textkörper einer Nachricht.|
|`Subject`| Zeichenfolge|Die Quelle der Daten stammt aus dem Betreff einer Nachricht.|

##### <a name="requirements"></a>Anforderungen

|Anforderung| Wert|
|---|---|
|[Mindestversion des Postfachanforderungssatzes](/javascript/office/requirement-sets/outlook-api-requirement-sets)| 1.0|
|[Zutreffender Outlook-Modus](https://docs.microsoft.com/outlook/add-ins/#extension-points)| Verfassen oder Lesen|