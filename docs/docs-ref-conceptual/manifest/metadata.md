# <a name="metadata-element"></a>Metadata-element

Definiert die Metadateneinstellungen, die durch eine benutzerdefinierte Funktion in Excel verwendet.

## <a name="attributes"></a>Attribute

Keine

## <a name="child-elements"></a>Untergeordnete Elemente

|  Element  |  Erforderlich  |  Beschreibung  |
|:-----|:-----|:-----|
|  [SourceLocation](customfunctionssourcelocation.md)  |  Ja  | Eine Zeichenfolge mit der Ressourcen-Id der JSON-Datei, die von benutzerdefinierten Funktionen. |

## <a name="example"></a>Beispiel

```xml
<Metadata>
    <SourceLocation resid="JSON-URL" />
</Metadata>
```
