
# <a name="userprofile"></a>userProfile

### [Office](Office.md)[.context](Office.context.md)[.mailbox](Office.context.mailbox.md). userProfile

##### <a name="requirements"></a>Anforderungen

|Anforderung| Wert|
|---|---|
|[Mindestversion des Postfachanforderungssatzes](/javascript/office/requirement-sets/outlook-api-requirement-sets)| 1.0|
|[Mindestberechtigungsstufe](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| ReadItem|
|[Zutreffender Outlook-Modus](https://docs.microsoft.com/outlook/add-ins/#extension-points)| Verfassen oder Lesen|

### <a name="members"></a>Elemente

####  <a name="displayname-string"></a>displayName :String

Ruft den Anzeigenamen des aktuellen Benutzers ab.

##### <a name="type"></a>Typ:

*   String

##### <a name="requirements"></a>Anforderungen

|Anforderung| Wert|
|---|---|
|[Mindestversion des Postfachanforderungssatzes](/javascript/office/requirement-sets/outlook-api-requirement-sets)| 1.0|
|[Mindestberechtigungsstufe](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| ReadItem|
|[Zutreffender Outlook-Modus](https://docs.microsoft.com/outlook/add-ins/#extension-points)| Verfassen oder Lesen|

##### <a name="example"></a>Beispiel

```
// Example: Allie Bellew
console.log(Office.context.mailbox.userProfile.displayName);
```

####  <a name="emailaddress-string"></a>emailAddress :String

Ruft die SMTP-E-Mail-Adresse des Benutzers ab.

##### <a name="type"></a>Typ:

*   String

##### <a name="requirements"></a>Anforderungen

|Anforderung| Wert|
|---|---|
|[Mindestversion des Postfachanforderungssatzes](/javascript/office/requirement-sets/outlook-api-requirement-sets)| 1.0|
|[Mindestberechtigungsstufe](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| ReadItem|
|[Zutreffender Outlook-Modus](https://docs.microsoft.com/outlook/add-ins/#extension-points)| Verfassen oder Lesen|

##### <a name="example"></a>Beispiel

```
// Example: allieb@contoso.com
console.log(Office.context.mailbox.userProfile.emailAddress);
```

####  <a name="timezone-string"></a>timeZone :String

Ruft die Standardzeitzone des Benutzers ab.

##### <a name="type"></a>Typ:

*   String

##### <a name="requirements"></a>Anforderungen

|Anforderung| Wert|
|---|---|
|[Mindestversion des Postfachanforderungssatzes](/javascript/office/requirement-sets/outlook-api-requirement-sets)| 1.0|
|[Mindestberechtigungsstufe](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| ReadItem|
|[Zutreffender Outlook-Modus](https://docs.microsoft.com/outlook/add-ins/#extension-points)| Verfassen oder Lesen|

##### <a name="example"></a>Beispiel

```
// Example: Pacific Standard Time
console.log(Office.context.mailbox.userProfile.timeZone);
```