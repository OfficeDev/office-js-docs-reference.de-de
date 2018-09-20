# <a name="getstarted-element"></a>GetStarted-Element

Stellt Informationen bereit, die von der Beschriftung verwendet werden, die angezeigt wird, wenn das Add-In auf Word-, Excel-, PowerPoint- und OneNote-Hosts installiert wird. Das **GetStarted**-Element ist ein untergeordnetes Element von [DesktopFormFactor](desktopformfactor.md).

## <a name="child-elements"></a>Untergeordnete Elemente

| Element                       | Erforderlich | Beschreibung                                        |
|:------------------------------|:--------:|:---------------------------------------------------|
| [Titel](#title)               | Ja      | Definiert, wo ein Add-In Funktionen verfügbar macht.     |
| [Beschreibung](#description)   | Ja      | Eine URL zu einer Datei, die JavaScript-Funktionen enthält.|
| [LearnMoreUrl](#learnmoreurl) | Nein       | Eine URL zu einer Seite, auf der das Add-In ausführlich erläutert wird.   |

### <a name="title"></a>Titel 

Erforderlich. Der Titel, der für den oberen Rand der Beschriftung verwendet wird. Das**resid**-Attribut verweist auf eine gültige ID im **ShortStrings**-Element im Abschnitt [Ressourcen](resources.md).

### <a name="description"></a>Beschreibung

Erforderlich. Die Beschreibung/der Textinhalt der Beschriftung. Das**resid**-Attribut verweist auf eine gültige ID im **LongStrings**-Element im Abschnitt [Ressourcen](resources.md).

### <a name="learnmoreurl"></a>LearnMoreUrl

Erforderlich. Die URL zu einer Seite, auf der der Benutzer mehr über das Add-In erfahren kann. Das**resid**-Attribut verweist auf eine gültige ID im **Urls**-Element im Abschnitt [Ressourcen](resources.md).

> [!NOTE]
> **LearnMoreUrl** wird derzeit nicht in Word, Excel oder PowerPoint-Clients gerendert. Es wird empfohlen, dass Sie diese URL für alle Clients hinzufügen, damit die URL dargestellt wird, sobald sie verfügbar ist. 

## <a name="see-also"></a>Siehe auch

In den folgenden Codebeispielen wird das **GetStarted**-Element verwendet:

* [Excel-Web-Add-In zum Bearbeiten der Tabellen- und Diagrammformatierung](https://github.com/OfficeDev/Excel-Add-in-JavaScript-SalesTracker)
* [Word Add-In JavaScript SpecKit](https://github.com/OfficeDev/Word-Add-in-JS-SpecKit)
* [Einfügen von Excel-Diagrammen mithilfe von Microsoft Graph in eine PowerPoint-add-in](https://github.com/OfficeDev/PowerPoint-Add-in-Microsoft-Graph-ASPNET-InsertChart)
