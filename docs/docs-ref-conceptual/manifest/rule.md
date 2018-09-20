# <a name="rule-element"></a>Rule-Element

Gibt die Aktivierungsregel(n) an, die für dieses E-Mail-Kontext-Add-In ausgewertet werden sollte(n).

**Add-In-Typ:** E-Mail-Kontext-Add-In

## <a name="contained-in"></a>Enthalten in

- [OfficeApp](officeapp.md)
- [ExtensionPoint](extensionpoint.md)

## <a name="attributes"></a>Attribute

| Attribut | Erforderlich | Beschreibung |
|:-----|:-----|:-----|
| **xsi:type** | Ja | Der Typ der Regel, die definiert wird. |

Folgende Regeltypen sind möglich:

- [ItemIs](#itemis-rule)
- [ItemHasAttachment](#itemhasattachment-rule)
- [ItemHasKnownEntity](#itemhasknownentity-rule)
- [ItemHasRegularExpressionMatch](#itemhasregularexpressionmatch-rule)
- [RuleCollection](#rulecollection)

## <a name="itemis-rule"></a>ItemIs-Regel

Definiert eine Regel, die „true“ ausgibt, wenn das ausgewählte Element den angegebenen Typ aufweist.

### <a name="attributes"></a>Attribute

| Attribut | Erforderlich | Beschreibung |
|:-----|:-----|:-----|
| **ItemType** | Ja | Gibt den zu findenden Elementtyp an. Kann `Message` oder `Appointment` sein. Der Elementtyp `Message` umfasst E-Mail, Besprechungsanfragen, Besprechungsantworten und Besprechungsabsagen. |
| **FormType** | Nein (innerhalb von [ExtensionPoint](extensionpoint.md)), Ja (innerhalb von [OfficeApp](officeapp.md)) | Gibt an, ob die App im Lese- oder Bearbeitungsformat für das Element angezeigt werden soll. Folgende Werte sind möglich: `Read`, `Edit`, `ReadOrEdit`. Bei Angabe für eine `Rule` innerhalb eines `ExtensionPoint` MUSS dieser Wert `Read` sein. |
| **ItemClass** | Nein | Gibt die zu findende benutzerdefinierte Nachrichtenklasse an. Weitere Informationen finden Sie unter [Aktivieren eines E-Mail-Add-Ins in Outlook für eine bestimmte Nachrichtenklasse](https://docs.microsoft.com/outlook/add-ins/activation-rules). |
| **IncludeSubClasses** | Nein | Gibt an, ob die Regel „true“ ausgeben soll, wenn das Element einer Unterklasse der angegebenen Nachrichtenklasse angehört; der Standardwert ist `false`. |

### <a name="example"></a>Beispiel

```XML
<Rule xsi:type="ItemIs" ItemType= "Message" />
```

## <a name="itemhasattachment-rule"></a>ItemHasAttachment-Regel

Definiert eine Regel, die „true“ ausgibt, wenn das Element eine Anlage enthält.

### <a name="example"></a>Beispiel

```XML
<Rule xsi:type="ItemHasAttachment" />
```

## <a name="itemhasknownentity-rule"></a>ItemHasKnownEntity-Regel

Definiert eine Regel, die „true“ ausgibt, wenn das Element im Betreff oder im Textkörper Text vom angegebenen Entitätstyp enthält.

### <a name="attributes"></a>Attribute

| Attribut | Erforderlich | Beschreibung |
|:-----|:-----|:-----|
| **EntityType** | Ja | Gibt den Entitätstyp an, der gefunden werden muss, damit die Regel „true“ ausgibt. Folgende Werte sind möglich: `MeetingSuggestion`, `TaskSuggestion`, `Address`, `Url`, `PhoneNumber`, `EmailAddress` oder `Contact`. |
| **RegExFilter** | Nein | Gibt einen regulären Ausdruck an, der zur Aktivierung für diese Entität ausgeführt wird. |
| **FilterName** | Nein | Gibt den Namen des regulären Ausdrucksfilters an, damit später im Code Ihres Add-Ins darauf verwiesen werden kann. |
| **IgnoreCase** | Nein | Legt fest, dass die Groß- und Kleinschreibung ignoriert wird, wenn der vom Attribut **RegExFilter** angegebene reguläre Ausdruck ausgeführt wird. |
| **Markieren** | Nein | **Hinweis:** Dies gilt nur für **Rule**-Elemente innerhalb von **ExtensionPoint**-Elementen. Gibt an, wie der Client übereinstimmende Entitäten hervorheben soll. Folgende Werte sind möglich: `all` oder `none`. Falls keine Angabe erfolgt, ist der Standardwert `all`. |

### <a name="example"></a>Beispiel

```XML
<Rule xsi:type="ItemHasKnownEntity" EntityType="EmailAddress" />
```

## <a name="itemhasregularexpressionmatch-rule"></a>ItemHasRegularExpressionMatch-Regel

Definiert eine Regel, die "true" ausgibt, wenn in der angegebenen Eigenschaft des Elements eine Übereinstimmung mit dem regulären Ausdruck vorhanden ist.

### <a name="attributes"></a>Attribute

| Attribut | Erforderlich | Beschreibung |
|:-----|:-----|:-----|
| **RegExName** | Ja | Gibt den Namen des regulären Ausdrucks an, damit Sie im Code Ihres Add-Ins auf den Ausdruck verweisen können. |
| **RegExValue** | Ja | Gibt den regulären Ausdruck an, der ausgewertet wird, um zu bestimmen, ob das E-Mail-Add-In angezeigt werden soll. |
| **PropertyName** | Ja | Gibt den Namen der Eigenschaft an, für die der reguläre Ausdruck ausgewertet wird. Folgende Werte sind möglich: `Subject`, `BodyAsPlaintext`, `BodyAsHtml` oder `SenderSTMPAddress`. |
| **IgnoreCase** | Nein | Gibt an, dass die Groß- und Kleinschreibung beim Ausführen des regulären Ausdrucks ignoriert wird. |
| **Markieren** | Nein | **Hinweis:** Dies gilt nur für **Rule**-Elemente innerhalb von **ExtensionPoint**-Elementen. Gibt an, wie der Client übereinstimmenden Text hervorheben soll. Folgende Werte sind möglich: `all` oder `none`. Falls keine Angabe erfolgt, ist der Standardwert `all`. |

### <a name="example"></a>Beispiel

```XML
<Rule xsi:type="ItemHasRegularExpressionMatch" RegExName="SupportArticleNumber" RegExValue="(\W|^)kb\d{6}(\W|$)" PropertyName="BodyAsHtml" IgnoreCase="true" />
```

## <a name="rulecollection"></a>RuleCollection

Definiert eine Sammlung von Regeln sowie den logischen Operator, der beim Auswerten der Regeln verwendet werden soll.

### <a name="attributes"></a>Attribute

| Attribut | Erforderlich | Beschreibung |
|:-----|:-----|:-----|
| **Mode** | Ja | Gibt den logischen Operator an, der beim Auswerten dieser Regelsammlung verwendet werden soll. Folgende Werte sind möglich: `And` oder `Or`. |

### <a name="example"></a>Beispiel

```XML
<Rule xsi:type="RuleCollection" Mode="And">
  <Rule xsi:type="ItemIs" ItemType="Message" />
  <Rule xsi:type="ItemHasKnownEntity" EntityType="MeetingSuggestion" />
  <Rule xsi:type="ItemHasKnownEntity" EntityType="Address" Highlight="none" />
</Rule>
```

## <a name="see-also"></a>Siehe auch

- [Aktivierungsregeln für Outlook-Add-Ins](https://docs.microsoft.com/outlook/add-ins/activation-rules)
- [Erkennen bestimmter Zeichenfolgen in einem Outlook-Element als bekannte Entitäten](https://docs.microsoft.com/outlook/add-ins/match-strings-in-an-item-as-well-known-entities)    
- [Verwenden regulärer Ausdrücke für Aktivierungsregeln zum Anzeigen eines Outlook-Add-Ins](https://docs.microsoft.com/outlook/add-ins/use-regular-expressions-to-show-an-outlook-add-in)