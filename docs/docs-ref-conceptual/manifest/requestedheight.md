# <a name="requestedheight-element"></a>RequestedHeight-Element

Gibt die ursprüngliche Höhe (in Pixel) eines Content-add-in oder e-Mail-add-in. 

**-Add-in-Typ:** Inhalte, e-Mail-Nachrichten

## <a name="syntax"></a>Syntax

```XML
<RequestedHeight>integer</RequestedHeight>
```

## <a name="contained-in"></a>Enthalten in

- [DefaultSettings](defaultsettings.md) (Content-add-ins) mit einem Wert, der zwischen 32 und 1000 sein kann
- [DesktopSettings](desktopsettings.md) und [TabletSettings](tabletsettings.md) (Mail-add-ins) mit einem Wert, der zwischen 32 und 450 werden kann
- [ExtensionPoint](extensionpoint.md) (Kontextbezogene Mail-add-ins) mit einem Wert, der zwischen 140 und 450 für die **DetectedEntity** Erweiterungspunkt sowie zwischen 32 und 450 für die **CustomPane** Erweiterungspunkt werden kann