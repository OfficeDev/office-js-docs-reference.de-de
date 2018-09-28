# <a name="rule-element"></a><span data-ttu-id="150ec-101">Rule-Element</span><span class="sxs-lookup"><span data-stu-id="150ec-101">Rule element</span></span>

<span data-ttu-id="150ec-102">Gibt die Aktivierungsregel(n) an, die für dieses E-Mail-Kontext-Add-In ausgewertet werden sollte(n).</span><span class="sxs-lookup"><span data-stu-id="150ec-102">Specifies the activation rule(s) that should be evaluated for this contextual mail add-in.</span></span>

<span data-ttu-id="150ec-103">**Add-In-Typ:** E-Mail-Kontext-Add-In</span><span class="sxs-lookup"><span data-stu-id="150ec-103">**Add-in type:** Mail contextual add-in</span></span>

## <a name="contained-in"></a><span data-ttu-id="150ec-104">Enthalten in</span><span class="sxs-lookup"><span data-stu-id="150ec-104">Contained in</span></span>

- [<span data-ttu-id="150ec-105">OfficeApp</span><span class="sxs-lookup"><span data-stu-id="150ec-105">OfficeApp</span></span>](officeapp.md)
- [<span data-ttu-id="150ec-106">ExtensionPoint</span><span class="sxs-lookup"><span data-stu-id="150ec-106">ExtensionPoint</span></span>](extensionpoint.md)

## <a name="attributes"></a><span data-ttu-id="150ec-107">Attribute</span><span class="sxs-lookup"><span data-stu-id="150ec-107">Attributes</span></span>

| <span data-ttu-id="150ec-108">Attribut</span><span class="sxs-lookup"><span data-stu-id="150ec-108">Attribute</span></span> | <span data-ttu-id="150ec-109">Erforderlich</span><span class="sxs-lookup"><span data-stu-id="150ec-109">Required</span></span> | <span data-ttu-id="150ec-110">Beschreibung</span><span class="sxs-lookup"><span data-stu-id="150ec-110">Description</span></span> |
|:-----|:-----|:-----|
| <span data-ttu-id="150ec-111">**xsi:type**</span><span class="sxs-lookup"><span data-stu-id="150ec-111">**xsi:type**</span></span> | <span data-ttu-id="150ec-112">Ja</span><span class="sxs-lookup"><span data-stu-id="150ec-112">Yes</span></span> | <span data-ttu-id="150ec-113">Der Typ der Regel, die definiert wird.</span><span class="sxs-lookup"><span data-stu-id="150ec-113">The type of rule being defined.</span></span> |

<span data-ttu-id="150ec-114">Folgende Regeltypen sind möglich:</span><span class="sxs-lookup"><span data-stu-id="150ec-114">The type of rule can be one of the following.</span></span>

- [<span data-ttu-id="150ec-115">ItemIs</span><span class="sxs-lookup"><span data-stu-id="150ec-115">ItemIs</span></span>](#itemis-rule)
- [<span data-ttu-id="150ec-116">ItemHasAttachment</span><span class="sxs-lookup"><span data-stu-id="150ec-116">ItemHasAttachment</span></span>](#itemhasattachment-rule)
- [<span data-ttu-id="150ec-117">ItemHasKnownEntity</span><span class="sxs-lookup"><span data-stu-id="150ec-117">ItemHasKnownEntity</span></span>](#itemhasknownentity-rule)
- [<span data-ttu-id="150ec-118">ItemHasRegularExpressionMatch</span><span class="sxs-lookup"><span data-stu-id="150ec-118">ItemHasRegularExpressionMatch</span></span>](#itemhasregularexpressionmatch-rule)
- [<span data-ttu-id="150ec-119">RuleCollection</span><span class="sxs-lookup"><span data-stu-id="150ec-119">RuleCollection</span></span>](#rulecollection)

## <a name="itemis-rule"></a><span data-ttu-id="150ec-120">ItemIs-Regel</span><span class="sxs-lookup"><span data-stu-id="150ec-120">ItemIs rule</span></span>

<span data-ttu-id="150ec-121">Definiert eine Regel, die „true“ ausgibt, wenn das ausgewählte Element den angegebenen Typ aufweist.</span><span class="sxs-lookup"><span data-stu-id="150ec-121">Defines a rule that evaluates to true if the selected item is of the specified type.</span></span>

### <a name="attributes"></a><span data-ttu-id="150ec-122">Attribute</span><span class="sxs-lookup"><span data-stu-id="150ec-122">Attributes</span></span>

| <span data-ttu-id="150ec-123">Attribut</span><span class="sxs-lookup"><span data-stu-id="150ec-123">Attribute</span></span> | <span data-ttu-id="150ec-124">Erforderlich</span><span class="sxs-lookup"><span data-stu-id="150ec-124">Required</span></span> | <span data-ttu-id="150ec-125">Beschreibung</span><span class="sxs-lookup"><span data-stu-id="150ec-125">Description</span></span> |
|:-----|:-----|:-----|
| <span data-ttu-id="150ec-126">**ItemType**</span><span class="sxs-lookup"><span data-stu-id="150ec-126">**ItemType**</span></span> | <span data-ttu-id="150ec-127">Ja</span><span class="sxs-lookup"><span data-stu-id="150ec-127">Yes</span></span> | <span data-ttu-id="150ec-p101">Gibt den zu findenden Elementtyp an. Kann `Message` oder `Appointment` sein. Der Elementtyp `Message` umfasst E-Mail, Besprechungsanfragen, Besprechungsantworten und Besprechungsabsagen.</span><span class="sxs-lookup"><span data-stu-id="150ec-p101">Specifies the item type to match. Can be `Message` or `Appointment`. `Message` item type includes email, meeting requests, meeting responses, and meeting cancellations.</span></span> |
| <span data-ttu-id="150ec-131">**FormType**</span><span class="sxs-lookup"><span data-stu-id="150ec-131">**FormType**</span></span> | <span data-ttu-id="150ec-132">Nein (innerhalb von [ExtensionPoint](extensionpoint.md)), Ja (innerhalb von [OfficeApp](officeapp.md))</span><span class="sxs-lookup"><span data-stu-id="150ec-132">No (within [ExtensionPoint](extensionpoint.md)), Yes (within [OfficeApp](officeapp.md))</span></span> | <span data-ttu-id="150ec-p102">Gibt an, ob die App im Lese- oder Bearbeitungsformat für das Element angezeigt werden soll. Folgende Werte sind möglich: `Read`, `Edit`, `ReadOrEdit`. Bei Angabe für eine `Rule` innerhalb eines `ExtensionPoint` MUSS dieser Wert `Read` sein.</span><span class="sxs-lookup"><span data-stu-id="150ec-p102">Specifies whether the app should appear in read or edit form for the item. Can be one of the following: `Read`, `Edit`, `ReadOrEdit`. If specified on a `Rule` within an `ExtensionPoint`, this value MUST be `Read`.</span></span> |
| <span data-ttu-id="150ec-136">**ItemClass**</span><span class="sxs-lookup"><span data-stu-id="150ec-136">**ItemClass**</span></span> | <span data-ttu-id="150ec-137">Nein</span><span class="sxs-lookup"><span data-stu-id="150ec-137">No</span></span> | <span data-ttu-id="150ec-p103">Gibt die zu findende benutzerdefinierte Nachrichtenklasse an. Weitere Informationen finden Sie unter [Aktivieren eines E-Mail-Add-Ins in Outlook für eine bestimmte Nachrichtenklasse](https://docs.microsoft.com/outlook/add-ins/activation-rules).</span><span class="sxs-lookup"><span data-stu-id="150ec-p103">Specifies the custom message class to match. For more information, see [Activate a mail add-in in Outlook for a specific message class](https://docs.microsoft.com/outlook/add-ins/activation-rules).</span></span> |
| <span data-ttu-id="150ec-140">**IncludeSubClasses**</span><span class="sxs-lookup"><span data-stu-id="150ec-140">**IncludeSubClasses**</span></span> | <span data-ttu-id="150ec-141">Nein</span><span class="sxs-lookup"><span data-stu-id="150ec-141">No</span></span> | <span data-ttu-id="150ec-142">Gibt an, ob die Regel „true“ ausgeben soll, wenn das Element einer Unterklasse der angegebenen Nachrichtenklasse angehört; der Standardwert ist `false`.</span><span class="sxs-lookup"><span data-stu-id="150ec-142">Specifies whether the rule should evaluate to true if the item is of a subclass of the specified message class; the default is `false`.</span></span> |

### <a name="example"></a><span data-ttu-id="150ec-143">Beispiel</span><span class="sxs-lookup"><span data-stu-id="150ec-143">Example</span></span>

```XML
<Rule xsi:type="ItemIs" ItemType= "Message" />
```

## <a name="itemhasattachment-rule"></a><span data-ttu-id="150ec-144">ItemHasAttachment-Regel</span><span class="sxs-lookup"><span data-stu-id="150ec-144">ItemHasAttachment rule</span></span>

<span data-ttu-id="150ec-145">Definiert eine Regel, die „true“ ausgibt, wenn das Element eine Anlage enthält.</span><span class="sxs-lookup"><span data-stu-id="150ec-145">Defines a rule that evaluates to true if the item contains an attachment.</span></span>

### <a name="example"></a><span data-ttu-id="150ec-146">Beispiel</span><span class="sxs-lookup"><span data-stu-id="150ec-146">Example</span></span>

```XML
<Rule xsi:type="ItemHasAttachment" />
```

## <a name="itemhasknownentity-rule"></a><span data-ttu-id="150ec-147">ItemHasKnownEntity-Regel</span><span class="sxs-lookup"><span data-stu-id="150ec-147">ItemHasKnownEntity rule</span></span>

<span data-ttu-id="150ec-148">Definiert eine Regel, die „true“ ausgibt, wenn das Element im Betreff oder im Textkörper Text vom angegebenen Entitätstyp enthält.</span><span class="sxs-lookup"><span data-stu-id="150ec-148">Defines a rule that evaluates to true if the item contains text of the specified entity type in its subject or body.</span></span>

### <a name="attributes"></a><span data-ttu-id="150ec-149">Attribute</span><span class="sxs-lookup"><span data-stu-id="150ec-149">Attributes</span></span>

| <span data-ttu-id="150ec-150">Attribut</span><span class="sxs-lookup"><span data-stu-id="150ec-150">Attribute</span></span> | <span data-ttu-id="150ec-151">Erforderlich</span><span class="sxs-lookup"><span data-stu-id="150ec-151">Required</span></span> | <span data-ttu-id="150ec-152">Beschreibung</span><span class="sxs-lookup"><span data-stu-id="150ec-152">Description</span></span> |
|:-----|:-----|:-----|
| <span data-ttu-id="150ec-153">**EntityType**</span><span class="sxs-lookup"><span data-stu-id="150ec-153">**EntityType**</span></span> | <span data-ttu-id="150ec-154">Ja</span><span class="sxs-lookup"><span data-stu-id="150ec-154">Yes</span></span> | <span data-ttu-id="150ec-p104">Gibt den Entitätstyp an, der gefunden werden muss, damit die Regel „true“ ausgibt. Folgende Werte sind möglich: `MeetingSuggestion`, `TaskSuggestion`, `Address`, `Url`, `PhoneNumber`, `EmailAddress` oder `Contact`.</span><span class="sxs-lookup"><span data-stu-id="150ec-p104">Specifies the type of entity that must be found for the rule to evaluate to true. Can be one of the following: `MeetingSuggestion`, `TaskSuggestion`, `Address`, `Url`, `PhoneNumber`, `EmailAddress`, or `Contact`.</span></span> |
| <span data-ttu-id="150ec-157">**RegExFilter**</span><span class="sxs-lookup"><span data-stu-id="150ec-157">**RegExFilter**</span></span> | <span data-ttu-id="150ec-158">Nein</span><span class="sxs-lookup"><span data-stu-id="150ec-158">No</span></span> | <span data-ttu-id="150ec-159">Gibt einen regulären Ausdruck an, der zur Aktivierung für diese Entität ausgeführt wird.</span><span class="sxs-lookup"><span data-stu-id="150ec-159">Specifies a regular expression to run against this entity for activation.</span></span> |
| <span data-ttu-id="150ec-160">**FilterName**</span><span class="sxs-lookup"><span data-stu-id="150ec-160">**FilterName**</span></span> | <span data-ttu-id="150ec-161">Nein</span><span class="sxs-lookup"><span data-stu-id="150ec-161">No</span></span> | <span data-ttu-id="150ec-162">Gibt den Namen des regulären Ausdrucksfilters an, damit später im Code Ihres Add-Ins darauf verwiesen werden kann.</span><span class="sxs-lookup"><span data-stu-id="150ec-162">Specifies the name of the regular expression filter, so that it is subsequently possible to refer to it in your add-in's code.</span></span> |
| <span data-ttu-id="150ec-163">**IgnoreCase**</span><span class="sxs-lookup"><span data-stu-id="150ec-163">**IgnoreCase**</span></span> | <span data-ttu-id="150ec-164">Nein</span><span class="sxs-lookup"><span data-stu-id="150ec-164">No</span></span> | <span data-ttu-id="150ec-165">Legt fest, dass die Groß- und Kleinschreibung ignoriert wird, wenn der vom Attribut **RegExFilter** angegebene reguläre Ausdruck ausgeführt wird.</span><span class="sxs-lookup"><span data-stu-id="150ec-165">Specifies to ignore case when running the regular expression specified by the  **RegExFilter** attribute.</span></span> |
| <span data-ttu-id="150ec-166">**Markieren**</span><span class="sxs-lookup"><span data-stu-id="150ec-166">**Highlight**</span></span> | <span data-ttu-id="150ec-167">Nein</span><span class="sxs-lookup"><span data-stu-id="150ec-167">No</span></span> | <span data-ttu-id="150ec-p105">**Hinweis:** Dies gilt nur für **Rule**-Elemente innerhalb von **ExtensionPoint**-Elementen. Gibt an, wie der Client übereinstimmende Entitäten hervorheben soll. Folgende Werte sind möglich: `all` oder `none`. Falls keine Angabe erfolgt, ist der Standardwert `all`.</span><span class="sxs-lookup"><span data-stu-id="150ec-p105">**Note:** this only applies to **Rule** elements within **ExtensionPoint** elements. Specifies how the client should highlight matching entities. Can be one of the following: `all` or `none`. If not specified, the default value is `all`.</span></span> |

### <a name="example"></a><span data-ttu-id="150ec-172">Beispiel</span><span class="sxs-lookup"><span data-stu-id="150ec-172">Example</span></span>

```XML
<Rule xsi:type="ItemHasKnownEntity" EntityType="EmailAddress" />
```

## <a name="itemhasregularexpressionmatch-rule"></a><span data-ttu-id="150ec-173">ItemHasRegularExpressionMatch-Regel</span><span class="sxs-lookup"><span data-stu-id="150ec-173">ItemHasRegularExpressionMatch rule</span></span>

<span data-ttu-id="150ec-174">Definiert eine Regel, die "true" ausgibt, wenn in der angegebenen Eigenschaft des Elements eine Übereinstimmung mit dem regulären Ausdruck vorhanden ist.</span><span class="sxs-lookup"><span data-stu-id="150ec-174">Defines a rule that evaluates to true if a match for the specified regular expression can be found in the specified property of the item.</span></span>

### <a name="attributes"></a><span data-ttu-id="150ec-175">Attribute</span><span class="sxs-lookup"><span data-stu-id="150ec-175">Attributes</span></span>

| <span data-ttu-id="150ec-176">Attribut</span><span class="sxs-lookup"><span data-stu-id="150ec-176">Attribute</span></span> | <span data-ttu-id="150ec-177">Erforderlich</span><span class="sxs-lookup"><span data-stu-id="150ec-177">Required</span></span> | <span data-ttu-id="150ec-178">Beschreibung</span><span class="sxs-lookup"><span data-stu-id="150ec-178">Description</span></span> |
|:-----|:-----|:-----|
| <span data-ttu-id="150ec-179">**RegExName**</span><span class="sxs-lookup"><span data-stu-id="150ec-179">**RegExName**</span></span> | <span data-ttu-id="150ec-180">Ja</span><span class="sxs-lookup"><span data-stu-id="150ec-180">Yes</span></span> | <span data-ttu-id="150ec-181">Gibt den Namen des regulären Ausdrucks an, damit Sie im Code Ihres Add-Ins auf den Ausdruck verweisen können.</span><span class="sxs-lookup"><span data-stu-id="150ec-181">Specifies the name of the regular expression, so that you can refer to the expression in the code for your add-in.</span></span> |
| <span data-ttu-id="150ec-182">**RegExValue**</span><span class="sxs-lookup"><span data-stu-id="150ec-182">**RegExValue**</span></span> | <span data-ttu-id="150ec-183">Ja</span><span class="sxs-lookup"><span data-stu-id="150ec-183">Yes</span></span> | <span data-ttu-id="150ec-184">Gibt den regulären Ausdruck an, der ausgewertet wird, um zu bestimmen, ob das E-Mail-Add-In angezeigt werden soll.</span><span class="sxs-lookup"><span data-stu-id="150ec-184">Specifies the regular expression that will be evaluated to determine whether the mail add-in should be shown.</span></span> |
| <span data-ttu-id="150ec-185">**PropertyName**</span><span class="sxs-lookup"><span data-stu-id="150ec-185">**PropertyName**</span></span> | <span data-ttu-id="150ec-186">Ja</span><span class="sxs-lookup"><span data-stu-id="150ec-186">Yes</span></span> | <span data-ttu-id="150ec-p106">Gibt den Namen der Eigenschaft an, für die der reguläre Ausdruck ausgewertet wird. Folgende Werte sind möglich: `Subject`, `BodyAsPlaintext`, `BodyAsHtml` oder `SenderSTMPAddress`.</span><span class="sxs-lookup"><span data-stu-id="150ec-p106">Specifies the name of the property that the regular expression will be evaluated against. Can be one of the following: `Subject`, `BodyAsPlaintext`, `BodyAsHtml`, or `SenderSTMPAddress`.</span></span> |
| <span data-ttu-id="150ec-189">**IgnoreCase**</span><span class="sxs-lookup"><span data-stu-id="150ec-189">**IgnoreCase**</span></span> | <span data-ttu-id="150ec-190">Nein</span><span class="sxs-lookup"><span data-stu-id="150ec-190">No</span></span> | <span data-ttu-id="150ec-191">Gibt an, dass die Groß- und Kleinschreibung beim Ausführen des regulären Ausdrucks ignoriert wird.</span><span class="sxs-lookup"><span data-stu-id="150ec-191">Specifies to ignore the case when executing the regular expression.</span></span> |
| <span data-ttu-id="150ec-192">**Markieren**</span><span class="sxs-lookup"><span data-stu-id="150ec-192">**Highlight**</span></span> | <span data-ttu-id="150ec-193">Nein</span><span class="sxs-lookup"><span data-stu-id="150ec-193">No</span></span> | <span data-ttu-id="150ec-p107">**Hinweis:** Dies gilt nur für **Rule**-Elemente innerhalb von **ExtensionPoint**-Elementen. Gibt an, wie der Client übereinstimmenden Text hervorheben soll. Folgende Werte sind möglich: `all` oder `none`. Falls keine Angabe erfolgt, ist der Standardwert `all`.</span><span class="sxs-lookup"><span data-stu-id="150ec-p107">**Note:** this only applies to **Rule** elements within **ExtensionPoint** elements. Specifies how the client should highlight matching text. Can be one of the following: `all` or `none`. If not specified, the default value is `all`.</span></span> |

### <a name="example"></a><span data-ttu-id="150ec-198">Beispiel</span><span class="sxs-lookup"><span data-stu-id="150ec-198">Example</span></span>

```XML
<Rule xsi:type="ItemHasRegularExpressionMatch" RegExName="SupportArticleNumber" RegExValue="(\W|^)kb\d{6}(\W|$)" PropertyName="BodyAsHtml" IgnoreCase="true" />
```

## <a name="rulecollection"></a><span data-ttu-id="150ec-199">RuleCollection</span><span class="sxs-lookup"><span data-stu-id="150ec-199">RuleCollection</span></span>

<span data-ttu-id="150ec-200">Definiert eine Sammlung von Regeln sowie den logischen Operator, der beim Auswerten der Regeln verwendet werden soll.</span><span class="sxs-lookup"><span data-stu-id="150ec-200">Defines a collection of rules and the logical operator to use when evaluating them.</span></span>

### <a name="attributes"></a><span data-ttu-id="150ec-201">Attribute</span><span class="sxs-lookup"><span data-stu-id="150ec-201">Attributes</span></span>

| <span data-ttu-id="150ec-202">Attribut</span><span class="sxs-lookup"><span data-stu-id="150ec-202">Attribute</span></span> | <span data-ttu-id="150ec-203">Erforderlich</span><span class="sxs-lookup"><span data-stu-id="150ec-203">Required</span></span> | <span data-ttu-id="150ec-204">Beschreibung</span><span class="sxs-lookup"><span data-stu-id="150ec-204">Description</span></span> |
|:-----|:-----|:-----|
| <span data-ttu-id="150ec-205">**Mode**</span><span class="sxs-lookup"><span data-stu-id="150ec-205">**Mode**</span></span> | <span data-ttu-id="150ec-206">Ja</span><span class="sxs-lookup"><span data-stu-id="150ec-206">Yes</span></span> | <span data-ttu-id="150ec-p108">Gibt den logischen Operator an, der beim Auswerten dieser Regelsammlung verwendet werden soll. Folgende Werte sind möglich: `And` oder `Or`.</span><span class="sxs-lookup"><span data-stu-id="150ec-p108">Specifies the logical operator to use when evaluating this rule collection. Can be either: `And` or `Or`.</span></span> |

### <a name="example"></a><span data-ttu-id="150ec-209">Beispiel</span><span class="sxs-lookup"><span data-stu-id="150ec-209">Example</span></span>

```XML
<Rule xsi:type="RuleCollection" Mode="And">
  <Rule xsi:type="ItemIs" ItemType="Message" />
  <Rule xsi:type="ItemHasKnownEntity" EntityType="MeetingSuggestion" />
  <Rule xsi:type="ItemHasKnownEntity" EntityType="Address" Highlight="none" />
</Rule>
```

## <a name="see-also"></a><span data-ttu-id="150ec-210">Siehe auch</span><span class="sxs-lookup"><span data-stu-id="150ec-210">See also</span></span>

- [<span data-ttu-id="150ec-211">Aktivierungsregeln für Outlook-Add-Ins</span><span class="sxs-lookup"><span data-stu-id="150ec-211">Activation rules for Outlook add-ins</span></span>](https://docs.microsoft.com/outlook/add-ins/activation-rules)
- [<span data-ttu-id="150ec-212">Erkennen bestimmter Zeichenfolgen in einem Outlook-Element als bekannte Entitäten</span><span class="sxs-lookup"><span data-stu-id="150ec-212">Match strings in an Outlook item as well-known entities</span></span>](https://docs.microsoft.com/outlook/add-ins/match-strings-in-an-item-as-well-known-entities)    
- [<span data-ttu-id="150ec-213">Verwenden regulärer Ausdrücke für Aktivierungsregeln zum Anzeigen eines Outlook-Add-Ins</span><span class="sxs-lookup"><span data-stu-id="150ec-213">Use regular expression activation rules to show an Outlook add-in</span></span>](https://docs.microsoft.com/outlook/add-ins/use-regular-expressions-to-show-an-outlook-add-in)