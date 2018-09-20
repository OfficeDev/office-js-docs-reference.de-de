# <a name="identity-api-requirement-sets"></a>API-Anforderungssätze identifizieren

Anforderungssätze sind benannte Gruppen von API-Mitgliedern. Office-Add-ins anforderungssätzen im Manifest angegebenen verwenden oder eine Überprüfung zur Laufzeit verwenden, um zu bestimmen, ob ein Office-Host APIs unterstützt, die ein Add-in muss. Weitere Informationen finden Sie unter [Office-Versionen und Anforderung festgelegt](https://docs.microsoft.com/office/dev/add-ins/develop/office-versions-and-requirement-sets).

Führen Sie Office-Add-ins in mehreren Versionen von Office. Die folgende Tabelle enthält die Identität API-anforderungssätzen, die Office-hostanwendungen, die für die Office-Anwendung festgelegt, und der Build oder Version Numbers Anforderung unterstützen.

|  Anforderungssatz  | Office 2013 für Windows | Office 365 für Windows   |  Office 365 für iPad  |  Office 365 für Mac  | Office Online  | SharePoint Online | OneDrive.com |Outlook.com und Exchange Online|
|:-----|-----|:-----|:-----|:-----|:-----|:-----|:-----|:-----|
| IdentityAPI 1.1  | Nicht zutreffend | Vorschau **& #42;** | Bald verfügbar | Vorschau **& #42;**| Verfügbar | Verfügbar| Bald verfügbar | Bald verfügbar |

> **& #42;** Während der Phase der Vorschau ist die Identität API 2016 Windows und Mac nur für Benutzer in der Insider-Anwendung mit der Option Fast unterstützt. Um die Insider Programm teilnehmen, finden Sie unter [Office-Insider werden](https://products.office.com/office-insider?tab=tab-1). Um die Fast Track zu wechseln, finden Sie unter [Fast Insider](https://answers.microsoft.com/en-us/msoffice/forum/msoffice_officeinsider-mso_win10-msoinsider_reg/its-here-office-insider-fast-for-office-2016-on/dbe8e7bb-9523-44a4-948b-9436fedfd961).

Weitere Informationen über Versionen, Buildnummern und Office Online Server finden Sie unter:

- [Versions- und Buildnummern der Updatekanalversionen für Office 365-Clients](https://support.office.com/article/version-and-build-numbers-of-update-channel-releases-ae942449-1fca-4484-898b-a933ea23def7)
- [Welche Version von Office verwende ich?](https://support.office.com/article/What-version-of-Office-am-I-using-932788b8-a3ce-44bf-bb09-e334518b8b19)
- [So finden Sie die Versions- und Bildnummer für eine Office 365-Clientanwendung](https://support.office.com/article/version-and-build-numbers-of-update-channel-releases-ae942449-1fca-4484-898b-a933ea23def7)
- 
  [Office Online Server-Übersicht](https://docs.microsoft.com/officeonlineserver/office-online-server-overview)

## <a name="office-common-api-requirement-sets"></a>Allgemeine Office-API-Anforderungssätze

Informationen zu allgemeinen API-Anforderungssätzen finden Sie unter [Allgemeine Office-API-Anforderungssätze](office-add-in-requirement-sets.md).

## <a name="identityapi-11"></a>IdentityAPI 1.1 

Die Single Sign-On-IdentityAPI 1.1 ist die erste Version der API. Weitere Informationen zur API finden Sie unter der `getAccessTokenAsync` -Methode in der [Office.Auth](/javascript/api/office/office.auth) Referenzthema.

## <a name="see-also"></a>Siehe auch

- [Office-Versionen und Anforderungssätze](https://docs.microsoft.com/office/dev/add-ins/develop/office-versions-and-requirement-sets)
- 
  [Angeben von Office-Hosts und API-Anforderungen](https://docs.microsoft.com/office/dev/add-ins/develop/specify-office-hosts-and-api-requirements)
- [XML-Manifest für Office-Add-Ins](https://docs.microsoft.com/office/dev/add-ins/develop/add-in-manifests)
