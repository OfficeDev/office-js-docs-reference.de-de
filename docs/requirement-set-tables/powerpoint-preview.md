| Klasse | Felder | Beschreibung |
|:---|:---|:---|
|[AddSlideOptions](/javascript/api/powerpoint/powerpoint.addslideoptions)|[layoutId](/javascript/api/powerpoint/powerpoint.addslideoptions#layoutid)|Gibt die ID eines Folienlayouts an, das für die neue Folie verwendet werden soll.|
||[slideMasterId](/javascript/api/powerpoint/powerpoint.addslideoptions#slidemasterid)|Gibt die ID eines Folienmasters an, der für die neue Folie verwendet werden soll.|
|[Presentation](/javascript/api/powerpoint/powerpoint.presentation)|[slideMasters](/javascript/api/powerpoint/powerpoint.presentation#slidemasters)|Gibt die Auflistung der `SlideMaster` Objekte zurück, die sich in der Präsentation befinden.|
||[tags](/javascript/api/powerpoint/powerpoint.presentation#tags)|Gibt eine Auflistung von Tags zurück, die der Präsentation zugeordnet sind.|
|[Form](/javascript/api/powerpoint/powerpoint.shape)|[id](/javascript/api/powerpoint/powerpoint.shape#id)|Ruft die eindeutige ID des Shapes ab.|
||[tags](/javascript/api/powerpoint/powerpoint.shape#tags)|Gibt eine Auflistung von Tags in der Form zurück.|
|[ShapeCollection](/javascript/api/powerpoint/powerpoint.shapecollection)|[getCount()](/javascript/api/powerpoint/powerpoint.shapecollection#getcount--)|Ruft die Anzahl der Shapes in der Auflistung ab.|
||[getItem(key: string)](/javascript/api/powerpoint/powerpoint.shapecollection#getitem-key-)|Ruft ein Shape mit seiner eindeutigen ID ab.|
||[getItemAt(index: number)](/javascript/api/powerpoint/powerpoint.shapecollection#getitemat-index-)|Ruft ein Shape mit seinem nullbasierten Index in der Auflistung ab.|
||[getItemOrNullObject(id: string)](/javascript/api/powerpoint/powerpoint.shapecollection#getitemornullobject-id-)|Ruft ein Shape mit seiner eindeutigen ID ab.|
||[items](/javascript/api/powerpoint/powerpoint.shapecollection#items)|Ruft die geladenen untergeordneten Elemente in dieser Sammlung ab.|
|[Slide](/javascript/api/powerpoint/powerpoint.slide)|[layout](/javascript/api/powerpoint/powerpoint.slide#layout)|Ruft das Layout der Folie ab.|
||[shapes](/javascript/api/powerpoint/powerpoint.slide#shapes)|Gibt eine Auflistung von Formen auf der Folie zurück.|
||[slideMaster](/javascript/api/powerpoint/powerpoint.slide#slidemaster)|Ruft das `SlideMaster` Objekt ab, das den Standardinhalt der Folie darstellt.|
||[tags](/javascript/api/powerpoint/powerpoint.slide#tags)|Gibt eine Auflistung von Tags auf der Folie zurück.|
|[SlideCollection](/javascript/api/powerpoint/powerpoint.slidecollection)|[add(options?: PowerPoint.AddSlideOptions)](/javascript/api/powerpoint/powerpoint.slidecollection#add-options-)|Fügt am Ende der Auflistung eine neue Folie hinzu.|
|[SlideLayout](/javascript/api/powerpoint/powerpoint.slidelayout)|[id](/javascript/api/powerpoint/powerpoint.slidelayout#id)|Ruft die eindeutige ID des Folienlayouts ab.|
||[name](/javascript/api/powerpoint/powerpoint.slidelayout#name)|Ruft den Namen des Folienlayouts ab.|
|[SlideLayoutCollection](/javascript/api/powerpoint/powerpoint.slidelayoutcollection)|[getCount()](/javascript/api/powerpoint/powerpoint.slidelayoutcollection#getcount--)|Ruft die Anzahl der Layouts in der Auflistung ab.|
||[getItem(key: string)](/javascript/api/powerpoint/powerpoint.slidelayoutcollection#getitem-key-)|Ruft ein Layout mit seiner eindeutigen ID ab.|
||[getItemAt(index: number)](/javascript/api/powerpoint/powerpoint.slidelayoutcollection#getitemat-index-)|Ruft ein Layout mit seinem nullbasierten Index in der Auflistung ab.|
||[getItemOrNullObject(id: string)](/javascript/api/powerpoint/powerpoint.slidelayoutcollection#getitemornullobject-id-)|Ruft ein Layout mit seiner eindeutigen ID ab.|
||[items](/javascript/api/powerpoint/powerpoint.slidelayoutcollection#items)|Ruft die geladenen untergeordneten Elemente in dieser Sammlung ab.|
|[SlideMaster](/javascript/api/powerpoint/powerpoint.slidemaster)|[id](/javascript/api/powerpoint/powerpoint.slidemaster#id)|Ruft die eindeutige ID des Folienmasters ab.|
||[layouts](/javascript/api/powerpoint/powerpoint.slidemaster#layouts)|Ruft die Auflistung von Layouts ab, die vom Folienmaster für Folien bereitgestellt werden.|
||[name](/javascript/api/powerpoint/powerpoint.slidemaster#name)|Ruft den eindeutigen Namen des Folienmasters ab.|
|[SlideMasterCollection](/javascript/api/powerpoint/powerpoint.slidemastercollection)|[getCount()](/javascript/api/powerpoint/powerpoint.slidemastercollection#getcount--)|Ruft die Anzahl der Folienmaster in der Auflistung ab.|
||[getItem(key: string)](/javascript/api/powerpoint/powerpoint.slidemastercollection#getitem-key-)|Ruft einen Folienmaster mit seiner eindeutigen ID ab.|
||[getItemAt(index: number)](/javascript/api/powerpoint/powerpoint.slidemastercollection#getitemat-index-)|Ruft einen Folienmaster mit seinem nullbasierten Index in der Auflistung ab.|
||[getItemOrNullObject(id: string)](/javascript/api/powerpoint/powerpoint.slidemastercollection#getitemornullobject-id-)|Ruft einen Folienmaster mit seiner eindeutigen ID ab.|
||[items](/javascript/api/powerpoint/powerpoint.slidemastercollection#items)|Ruft die geladenen untergeordneten Elemente in dieser Sammlung ab.|
|[Tag](/javascript/api/powerpoint/powerpoint.tag)|[key](/javascript/api/powerpoint/powerpoint.tag#key)|Ruft die eindeutige ID des Tags ab.|
||[value](/javascript/api/powerpoint/powerpoint.tag#value)|Ruft den Wert des Tags ab.|
|[TagCollection](/javascript/api/powerpoint/powerpoint.tagcollection)|[add(key: string, value: string)](/javascript/api/powerpoint/powerpoint.tagcollection#add-key--value-)|Fügt am Ende der Auflistung ein neues Tag hinzu.|
||[delete(key: string)](/javascript/api/powerpoint/powerpoint.tagcollection#delete-key-)|Löscht das Tag mit dem in dieser `key` Auflistung angegebenen.|
||[getCount()](/javascript/api/powerpoint/powerpoint.tagcollection#getcount--)|Ruft die Anzahl der Tags in der Auflistung ab.|
||[getItem(key: string)](/javascript/api/powerpoint/powerpoint.tagcollection#getitem-key-)|Ruft ein Tag mit seiner eindeutigen ID ab.|
||[getItemAt(index: number)](/javascript/api/powerpoint/powerpoint.tagcollection#getitemat-index-)|Ruft ein Tag mit seinem nullbasierten Index in der Auflistung ab.|
||[getItemOrNullObject(key: string)](/javascript/api/powerpoint/powerpoint.tagcollection#getitemornullobject-key-)|Ruft ein Tag mit seiner eindeutigen ID ab.|
||[items](/javascript/api/powerpoint/powerpoint.tagcollection#items)|Ruft die geladenen untergeordneten Elemente in dieser Sammlung ab.|
