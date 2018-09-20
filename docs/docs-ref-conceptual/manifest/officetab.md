# <a name="officetab-element"></a>OfficeTab-Element

Definiert die Registerkarte des Menübands, auf der Ihr Add-In-Befehl angezeigt wird. Dies kann entweder die Standardregisterkarte sein (**Startseite**,  **Nachricht** oder  **Besprechung**), oder eine benutzerdefinierte Registerkarte, die vom Add-In definiert wird. Dieses Element ist erforderlich.

## <a name="child-elements"></a>Untergeordnete Elemente

|  Element |  Erforderlich  |  Beschreibung  |
|:-----|:-----|:-----|
|  Group      | Ja |  Definiert eine Gruppe von Befehlen. Sie können nur eine Gruppe pro Add-In zu der Standardregisterkarte hinzufügen.  |

Nachfolgend finden Sie die gültigen `id`-Werte nach Host. **Fett formatierte** Werte werden sowohl auf dem Desktop als auch online unterstützt (z. B. Word 2016 für Windows und Word Online). 

### <a name="outlook"></a>Outlook 

- **TabDefault**

### <a name="word"></a>Word

- **TabHome**
- **TabInsert**
- TabWordDesign
- **TabPageLayoutWord**
- TabReferences
- TabMailings
- TabReviewWord
- **TabView**
- TabDeveloper
- TabAddIns
- TabBlogPost
- TabBlogInsert
- TabPrintPreview
- TabOutlining
- TabConflicts
- TabBackgroundRemoval
- TabBroadcastPresentation

### <a name="excel"></a>Excel

- **TabHome**
- **TabInsert**
- TabPageLayoutExcel
- TabFormulas
- **TabData**
- **TabReview**
- **TabView**
- TabDeveloper
- TabAddIns
- TabPrintPreview
- TabBackgroundRemoval 

### <a name="powerpoint"></a>PowerPoint

- **TabHome**
- **TabInsert**
- **TabDesign**
- **TabTransitions**
- **TabAnimations**
- TabSlideShow
- TabReview
- **TabView**
- TabDeveloper
- TabAddIns
- TabPrintPreview
- TabMerge
- TabGrayscale
- TabBlackAndWhite
- TabBroadcastPresentation
- TabSlideMaster
- TabHandoutMaster
- TabNotesMaster
- TabBackgroundRemoval
- TabSlideMasterHome

### <a name="onenote"></a>OneNote

- **TabHome**
- **TabInsert**
- **TabView**
- TabDeveloper
- TabAddIns

## <a name="group"></a>Group

Eine Gruppe von UI-Erweiterungspunkten auf einer Registerkarte. Eine Gruppe kann bis zu sechs Steuerelemente enthalten. Das **id**-Attribut ist erforderlich, und jede **id** muss im Manifest eindeutig sein. Die **id** ist eine Zeichenfolge mit maximal 125 Zeichen. Siehe [Group-Element](group.md).

## <a name="officetab-example"></a>OfficeTab-Beispiel 

```xml
<ExtensionPoint xsi:type="MessageReadCommandSurface">
  <OfficeTab id="TabDefault">
    <Group id="msgreadTabMessage.grp1">
        <!-- Group Definition -->
    </Group>
  </OfficeTab>
</ExtensionPoint>
```
