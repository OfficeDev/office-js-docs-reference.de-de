# <a name="sourcelocation-element"></a>SourceLocation-Element

Definiert den Speicherort einer Ressource, die für Script- oder Page-Elemente erforderlich ist, die von benutzerdefinierten Funktionen in Excel verwendet werden.

## <a name="attributes"></a>Attribute

| **Attribut** | **Erforderlich** | **Beschreibung**                                                                      |
|---------------|--------------|--------------------------------------------------------------------------------------|
| resid         | Ja          | Der Name einer URL-Ressource, die im Abschnitt &lt;Ressourcen&gt; des Manifests definiert ist. |

## <a name="child-elements"></a>Untergeordnete Elemente

Keine

## <a name="example"></a>Beispiel

```xml
<SourceLocation resid="pageURL"/>
```