| Klasse | Felder | Beschreibung |
|:---|:---|:---|
|[InsertSlideOptions](/javascript/api/powerpoint/powerpoint.insertslideoptions)|[Formatierung](/javascript/api/powerpoint/powerpoint.insertslideoptions#formatting)|Gibt an, welche Formatierung beim Einfügen von Folien verwendet werden soll.|
||[sourceSlideIds](/javascript/api/powerpoint/powerpoint.insertslideoptions#sourceslideids)|Gibt die Folien aus der Quellpräsentation an, die in die aktuelle Präsentation eingefügt werden.|
||[targetSlideId](/javascript/api/powerpoint/powerpoint.insertslideoptions#targetslideid)|Gibt an, wo in der Präsentation die neuen Folien eingefügt werden.|
|[Presentation](/javascript/api/powerpoint/powerpoint.presentation)|[insertSlidesFromBase64 (Base64-Datei: String, Options?: PowerPoint. InsertSlideOptions)](/javascript/api/powerpoint/powerpoint.presentation#insertslidesfrombase64-base64file--options-)|Fügt die angegebenen Folien aus einer Präsentation in die aktuelle Präsentation ein.|
||[Folien](/javascript/api/powerpoint/powerpoint.presentation#slides)|Gibt eine geordnete Auflistung von Folien in der Präsentation zurück.|
|[Slide](/javascript/api/powerpoint/powerpoint.slide)|[delete()](/javascript/api/powerpoint/powerpoint.slide#delete--)|Löscht die Folie aus der Präsentation.|
||[id](/javascript/api/powerpoint/powerpoint.slide#id)|Ruft die eindeutige ID der Folie ab.|
|[SlideCollection](/javascript/api/powerpoint/powerpoint.slidecollection)|[getCount()](/javascript/api/powerpoint/powerpoint.slidecollection#getcount--)|Ruft die Anzahl der Folien in der Auflistung ab.|
||[getItem(key: string)](/javascript/api/powerpoint/powerpoint.slidecollection#getitem-key-)|Ruft eine Folie mit ihrer eindeutigen ID ab.|
||[getItemAt(index: number)](/javascript/api/powerpoint/powerpoint.slidecollection#getitemat-index-)|Ruft eine Folie mit dem nullbasierten Index in der Auflistung ab.|
||[getItemOrNullObject(id: string)](/javascript/api/powerpoint/powerpoint.slidecollection#getitemornullobject-id-)|Ruft eine Folie mit ihrer eindeutigen ID ab.|
||[items](/javascript/api/powerpoint/powerpoint.slidecollection#items)|Ruft die geladenen untergeordneten Elemente in dieser Sammlung ab.|
