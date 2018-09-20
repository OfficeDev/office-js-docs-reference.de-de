# <a name="dialog-api-requirement-sets"></a>Dialog-API-Anforderungssätze

Anforderungssätze sind benannte Gruppen von API-Mitgliedern. Office-Add-ins anforderungssätzen im Manifest angegebenen verwenden oder eine Überprüfung zur Laufzeit verwenden, um zu bestimmen, ob ein Office-Host APIs unterstützt, die ein Add-in muss. Weitere Informationen finden Sie unter [Office-Versionen und Anforderung festgelegt](https://docs.microsoft.com/office/dev/add-ins/develop/office-versions-and-requirement-sets).

Office-Add-Ins sind auf mehreren Versionen von Office lauffähig. In der folgenden Tabelle finden Sie Informationen über die Dialog-API-Anforderungssätze, die Office-Hostanwendungen, die den jeweiligen Anforderungssatz unterstützen, und die Build- oder Versionsnummern für die Office-Anwendung.

|  Anforderungssatz  | Office 2013 für Windows | Office 2016 für Windows (MSI-Installationen)   | Office 365 für Windows (C2R installiert)   |  Office 365 für iPad  |  Office 365 für Mac  | Office Online  |  Office Online Server  |
|:-----|-----|:-----|:-----|:-----|:-----|:-----|:-----|
| DialogApi 1.1  | Build 15.0.4855.1000 oder höher | Build 16.0.4390.1000 oder höher | Version 1602 (Build 6741.0000) oder höher | 1.22 oder höher | 15.20 oder höher| Januar 2017 | Version 1608 (Build 7601.6800) oder höher|

Weitere Informationen über Versionen, Buildnummern und Office Online Server finden Sie unter:

- [Versions- und Buildnummern der Updatekanalversionen für Office 365-Clients](https://support.office.com/article/version-and-build-numbers-of-update-channel-releases-ae942449-1fca-4484-898b-a933ea23def7)
- [Welche Version von Office verwende ich?](https://support.office.com/article/What-version-of-Office-am-I-using-932788b8-a3ce-44bf-bb09-e334518b8b19)
- [So finden Sie die Versions- und Bildnummer für eine Office 365-Clientanwendung](https://support.office.com/article/version-and-build-numbers-of-update-channel-releases-ae942449-1fca-4484-898b-a933ea23def7)
- [Office Online Server-Übersicht](https://docs.microsoft.com/officeonlineserver/office-online-server-overview)

## <a name="office-common-api-requirement-sets"></a>Allgemeine Office-API-Anforderungssätze

Informationen zu allgemeinen API-Anforderungssätzen finden Sie unter [Allgemeine Office-API-Anforderungssätze](office-add-in-requirement-sets.md).

## <a name="dialog-api-11"></a>Dialog-API 1.1 

Das Dialogfeld API 1.1 ist die erste Version der API. Weitere Informationen zur API finden Sie im [Dialogfeld API](/javascript/api/office/office.ui) -Referenzthema.

## <a name="see-also"></a>Siehe auch

- [Office-Versionen und Anforderungssätze](https://docs.microsoft.com/office/dev/add-ins/develop/office-versions-and-requirement-sets)
- 
  [Angeben von Office-Hosts und API-Anforderungen](https://docs.microsoft.com/office/dev/add-ins/develop/specify-office-hosts-and-api-requirements)
- [XML-Manifest für Office-Add-Ins](https://docs.microsoft.com/office/dev/add-ins/develop/add-in-manifests)
