# <a name="onenote-javascript-api-requirement-sets"></a>JavaScript-API-Anforderungssätze für OneNote

Anforderungssätze sind benannte Gruppen von API-Mitgliedern. Office-Add-ins anforderungssätzen im Manifest angegebenen verwenden oder eine Überprüfung zur Laufzeit verwenden, um zu bestimmen, ob ein Office-Host APIs unterstützt, die ein Add-in muss. Weitere Informationen finden Sie unter [Office-Versionen und Anforderung festgelegt](https://docs.microsoft.com/office/dev/add-ins/develop/office-versions-and-requirement-sets).

In der folgenden Tabelle sind die OneNote-Anforderungssätze, die Office-Hostanwendungen, die diese Anforderungssätze unterstützen, und die Buildversionen oder das Verfügbarkeitsdatum aufgeführt.

|  Anforderungssatz  |  Office Online | 
|:-----|:-----|
| OneNoteApi 1.1  | September 2016 |  

## <a name="office-common-api-requirement-sets"></a>Allgemeine Office-API-Anforderungssätze

Informationen zu allgemeinen API-Anforderungssätzen finden Sie unter [Allgemeine Office-API-Anforderungssätze](office-add-in-requirement-sets.md).

## <a name="onenote-javascript-api-11"></a>OneNote-JavaScript-API 1.1 

Die OneNote-JavaScript-API 1.1 ist die erste Version der API. Weitere Informationen zur API finden Sie unter der [OneNote-JavaScript-API programming (Übersicht)](https://docs.microsoft.com/office/dev/add-ins/onenote/onenote-add-ins-programming-overview).

## <a name="runtime-requirement-support-check"></a>Überprüfung der Unterstützung einer Anforderung zur Laufzeit

Während der Laufzeit können Add-Ins überprüfen, ob ein bestimmter Host einen API-Anforderungssatz unterstützt, indem die folgende Überprüfung durchgeführt wird: 

```js
if (Office.context.requirements.isSetSupported('OneNoteApi', 1.1) === true) {
  /// perform actions
}
else {
  /// provide alternate flow/logic
}
```

## <a name="manifest-based-requirement-support-check"></a>Überprüfung der Unterstützung einer manifestbasierten Anforderung

Verwenden Sie das Requirements-Element im Add-In-Manifest, um wichtige Anforderungssätze oder API-Elemente anzugeben, die Ihr Add-In verwenden muss. Wenn der Office-Host oder die Plattform die im Requirements-Element angegebenen Anforderungssätze oder API-Elemente nicht unterstützt, wird das Add-In nicht auf diesem Host oder dieser Plattform ausgeführt und nicht unter „Meine Add-Ins“ angezeigt.

Das folgende Codebeispiel zeigt ein Add-In, das in allen Office-Hostanwendungen geladen wird, die den OneNoteApi-Anforderungssatz, Version 1.1, unterstützen.

```xml
<Requirements>
   <Sets DefaultMinVersion="1.1">
      <Set Name="OneNoteApi" MinVersion="1.1"/>
   </Sets>
</Requirements>
```

## <a name="see-also"></a>Siehe auch

- [Office-Versionen und Anforderungssätze](https://docs.microsoft.com/office/dev/add-ins/develop/office-versions-and-requirement-sets)
- 
  [Angeben von Office-Hosts und API-Anforderungen](https://docs.microsoft.com/office/dev/add-ins/develop/specify-office-hosts-and-api-requirements)
- [XML-Manifest für Office-Add-Ins](https://docs.microsoft.com/office/dev/add-ins/develop/add-in-manifests)
