# <a name="allowsnapshot-element"></a>AllowSnapshot-Element

Gibt an, ob eine Momentaufnahme Ihres Inhalts-Add-Ins mit dem Hostdokument gespeichert wird.

**Add-In-Typ:** Inhalt

## <a name="syntax"></a>Syntax

```XML
<AllowSnapshot> [true | false]</AllowSnapshot>
```

## <a name="contained-in"></a>Enthalten in

[OfficeApp](officeapp.md)

## <a name="remarks"></a>Hinweise

 > [!IMPORTANT]
 > **AllowSnapshot** ist `true` standardmäßig. Dadurch wird ein Bild des Add-Ins für Benutzer sichtbar, die das Dokument in einer Version der Hostanwendung öffnen, die Office-Add-Ins nicht unterstützt, oder es wird ein statisches Bild des Add-Ins bereitgestellt, wenn die Hostanwendung keine Verbindung zu dem Server herstellen kann, auf dem das Add-In gehostet wird. Dies bedeutet jedoch auch, dass auf möglicherweise vertrauliche Informationen, die in dem Add-in angezeigt werden, direkt aus dem Dokument, das das Add-In hostet, zugegriffen werden kann.

