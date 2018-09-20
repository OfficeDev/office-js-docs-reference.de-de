# <a name="page-element"></a>Page-Element

Definiert HTML-Seiteneinstellungen, die von einer benutzerdefinierten Funktion in Excel verwendet werden.

## <a name="attributes"></a>Attribute

Keine

## <a name="child-elements"></a>Untergeordnete Elemente

|  Element  |  Erforderlich  |  Beschreibung  |
|:-----|:-----|:-----|
|  [SourceLocation](customfunctionssourcelocation.md)  |  Ja  | Zeichenfolge mit Ressourcen-ID der JavaScript-Datei, die von benutzerdefinierten Funktionen verwendet wird. |

## <a name="example"></a>Beispiel

```xml
<Page>
    <SourceLocation resid="pageURL"/>
</Page>
```
