# <a name="script-element"></a>Script-Element

Definiert Skripteinstellungen, die von einer benutzerdefinierten Funktion in Excel verwendet werden.

## <a name="attributes"></a>Attribute

Keine

## <a name="child-elements"></a>Untergeordnete Elemente

|Elemente  |  Erforderlich  |  Beschreibung  |
|:-----|:-----|:-----|
|  [SourceLocation](customfunctionssourcelocation.md)  |  Ja  | Zeichenfolge mit Ressourcen-ID der JavaScript-Datei, die von benutzerdefinierten Funktionen verwendet wird.|

## <a name="example"></a>Beispiel

```xml
<Script>
    <SourceLocation resid="scriptURL" />
    <!-- The Script element is not used in the Developer Preview. -->
</Script>
```
