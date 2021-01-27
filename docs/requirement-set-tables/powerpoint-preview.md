| Klasse | Felder | Beschreibung |
|:---|:---|:---|
|[AddSlideOptions](/javascript/api/powerpoint/powerpoint.addslideoptions)|[layoutId](/javascript/api/powerpoint/powerpoint.addslideoptions#layoutid)|Gibt die ID eines Folienlayouts an, das für die neue Folie verwendet werden soll.|
||[slideMasterId](/javascript/api/powerpoint/powerpoint.addslideoptions#slidemasterid)|Gibt die ID eines Folienmasters an, der für die neue Folie verwendet werden soll.|
|[Presentation](/javascript/api/powerpoint/powerpoint.presentation)|[slideMasters](/javascript/api/powerpoint/powerpoint.presentation#slidemasters)|Gibt die Auflistung der `SlideMaster` Objekte zurück, die sich in der Präsentation befinden.|
|[Form](/javascript/api/powerpoint/powerpoint.shape)|[id](/javascript/api/powerpoint/powerpoint.shape#id)|Ruft die eindeutige ID des Shapes ab.|
|[ShapeCollection](/javascript/api/powerpoint/powerpoint.shapecollection)|[getCount()](/javascript/api/powerpoint/powerpoint.shapecollection#getcount--)|Ruft die Anzahl der Formen in der Auflistung ab.|
||[getItem(key: string)](/javascript/api/powerpoint/powerpoint.shapecollection#getitem-key-)|Ruft ein Shape mithilfe seiner eindeutigen ID ab.|
||[getItemAt(index: number)](/javascript/api/powerpoint/powerpoint.shapecollection#getitemat-index-)|Ruft ein Shape mithilfe seines nullbasierten Indexes in der Auflistung ab.|
||[getItemOrNullObject(id: string)](/javascript/api/powerpoint/powerpoint.shapecollection#getitemornullobject-id-)|Ruft ein Shape mithilfe seiner eindeutigen ID ab.|
||[items](/javascript/api/powerpoint/powerpoint.shapecollection#items)|Ruft die geladenen untergeordneten Elemente in dieser Sammlung ab.|
|[Slide](/javascript/api/powerpoint/powerpoint.slide)|[layout](/javascript/api/powerpoint/powerpoint.slide#layout)|Ruft das Layout der Folie ab.|
||[shapes](/javascript/api/powerpoint/powerpoint.slide#shapes)|Gibt eine Auflistung von Formen auf der Folie zurück.|
||[slideMaster](/javascript/api/powerpoint/powerpoint.slide#slidemaster)|Ruft das `SlideMaster` Objekt ab, das den Standardinhalt der Folie darstellt.|
|[SlideCollection](/javascript/api/powerpoint/powerpoint.slidecollection)|[add(options?: PowerPoint.AddSlideOptions)](/javascript/api/powerpoint/powerpoint.slidecollection#add-options-)|Fügt am Ende der Auflistung eine neue Folie hinzu.|
|[SlideLayout](/javascript/api/powerpoint/powerpoint.slidelayout)|[id](/javascript/api/powerpoint/powerpoint.slidelayout#id)|Ruft die eindeutige ID des Folienlayouts ab.|
||[name](/javascript/api/powerpoint/powerpoint.slidelayout#name)|Ruft den Namen des Folienlayouts ab.|
|[SlideLayoutCollection](/javascript/api/powerpoint/powerpoint.slidelayoutcollection)|[getCount()](/javascript/api/powerpoint/powerpoint.slidelayoutcollection#getcount--)|Ruft die Anzahl der Layouts in der Auflistung ab.|
||[getItem(key: string)](/javascript/api/powerpoint/powerpoint.slidelayoutcollection#getitem-key-)|Ruft ein Layout mithilfe seiner eindeutigen ID ab.|
||[getItemAt(index: number)](/javascript/api/powerpoint/powerpoint.slidelayoutcollection#getitemat-index-)|Ruft ein Layout mithilfe seines nullbasierten Indexes in der Auflistung ab.|
||[getItemOrNullObject(id: string)](/javascript/api/powerpoint/powerpoint.slidelayoutcollection#getitemornullobject-id-)|Ruft ein Layout mithilfe seiner eindeutigen ID ab.|
||[items](/javascript/api/powerpoint/powerpoint.slidelayoutcollection#items)|Ruft die geladenen untergeordneten Elemente in dieser Sammlung ab.|
|[SlideMaster](/javascript/api/powerpoint/powerpoint.slidemaster)|[id](/javascript/api/powerpoint/powerpoint.slidemaster#id)|Ruft die eindeutige ID des Folienmasters ab.|
||[layouts](/javascript/api/powerpoint/powerpoint.slidemaster#layouts)|Ruft die Auflistung der Layouts ab, die vom Folienmaster für Folien bereitgestellt werden.|
||[name](/javascript/api/powerpoint/powerpoint.slidemaster#name)|Ruft den eindeutigen Namen des Folienmasters ab.|
|[SlideMasterCollection](/javascript/api/powerpoint/powerpoint.slidemastercollection)|[getCount()](/javascript/api/powerpoint/powerpoint.slidemastercollection#getcount--)|Ruft die Anzahl der Folienmaster in der Auflistung ab.|
||[getItem(key: string)](/javascript/api/powerpoint/powerpoint.slidemastercollection#getitem-key-)|Ruft einen Folienmaster mithilfe seiner eindeutigen ID ab.|
||[getItemAt(index: number)](/javascript/api/powerpoint/powerpoint.slidemastercollection#getitemat-index-)|Ruft einen Folienmaster mithilfe seines nullbasierten Indexes in der Auflistung ab.|
||[getItemOrNullObject(id: string)](/javascript/api/powerpoint/powerpoint.slidemastercollection#getitemornullobject-id-)|Ruft einen Folienmaster mithilfe seiner eindeutigen ID ab.|
||[items](/javascript/api/powerpoint/powerpoint.slidemastercollection#items)|Ruft die geladenen untergeordneten Elemente in dieser Sammlung ab.|
