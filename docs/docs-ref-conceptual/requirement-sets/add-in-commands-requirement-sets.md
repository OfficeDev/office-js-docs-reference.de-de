# <a name="add-in-commands-requirement-sets"></a>Add-In-Befehlsanforderungssätze

Anforderungssätze sind benannte Gruppen von API-Mitgliedern. Office-Add-ins anforderungssätzen im Manifest angegebenen verwenden oder eine Überprüfung zur Laufzeit verwenden, um zu bestimmen, ob ein Office-Host APIs unterstützt, die ein Add-in muss. Weitere Informationen finden Sie unter [Office-Versionen und Anforderung festgelegt](https://docs.microsoft.com/office/dev/add-ins/develop/office-versions-and-requirement-sets).

Add-In-Befehle sind Elemente der Benutzeroberfläche, die die Office-Benutzeroberfläche erweitern und Aktionen in Ihrem Add-In starten. Sie können Add-In-Befehle verwenden, um eine Schaltfläche zum Menüband oder ein Element zu einem Kontextmenü hinzuzufügen. Weitere Informationen finden Sie unter [Add-In-Befehle für Excel, Word und PowerPoint](https://docs.microsoft.com/office/dev/add-ins/design/add-in-commands) und [Add-In-Befehle für Outlook](https://docs.microsoft.com/outlook/add-ins/add-in-commands-for-outlook).

Die erste Version von Add-In-Befehlen hat noch keinen entsprechenden Anforderungssatz (d. h., es gibt keinen AddInCommands 1.0-Anforderungssatz). In der folgenden Tabelle werden die Office-Hostanwendungen aufgeführt, die die erste Version unterstützen, sowie die Buildversionen oder -nummern dieser Anwendungen.  

| Version   |  Office 2013 für Windows | Office 2016 für Windows (ohne Abonnement) | Office 365 für Windows   |  Office 365 für iPad  |  Office 365 für Mac  | Office Online  |  
|:-----|:-----|:-----|:-----|:-----|:-----|:-----|
| Add-In-Befehle (erste Version, kein Anforderungssatz) | – | 16.0.4678.1000 *unterstützt nur Outlook* |Version 1603 (Build 6769.0000) oder höher | Nicht zutreffend | 15.33 oder höher| Januar 2016 | |

Im Anforderungssatz für Add-In-Befehle der Version 1.1 wurde die Möglichkeit eingeführt, [einen Aufgabenbereich mit Dokumenten automatisch zu öffnen](https://docs.microsoft.com/office/dev/add-ins/develop/automatically-open-a-task-pane-with-a-document).

In der folgenden Tabelle sind der Anforderungssatz für Add-In-Befehle der Version 1.1 aufgeführt, die Office-Hostanwendungen, die diesen Anforderungssatz unterstützen, sowie die Build- oder Versionsnummern der Office-Anwendung. 

|  Anforderungssatz  |  Office 2013 für Windows | Office 2016 für Windows (ohne Abonnement) | Office 365 für Windows   |  Office 365 für iPad  |  Office 365 für Mac  | Office Online  |  
|:-----|:-----|:-----|:-----|:-----|:-----|:-----|
| AddInCommands 1.1  | Nicht zutreffend | 16.0.4678.1000 *unterstützt nur Outlook*  | Version 1705 (Build 8121.1000) oder höher | Nicht zutreffend | 15.34 oder höher| Mai 2017 | |

Weitere Informationen über Versionen, Buildnummern und Office Online Server finden Sie unter:

- [Versions- und Buildnummern der Updatekanalversionen für Office 365-Clients](https://support.office.com/article/version-and-build-numbers-of-update-channel-releases-ae942449-1fca-4484-898b-a933ea23def7)
- [Welche Version von Office verwende ich?](https://support.office.com/article/What-version-of-Office-am-I-using-932788b8-a3ce-44bf-bb09-e334518b8b19)
- [So finden Sie die Versions- und Bildnummer für eine Office 365-Clientanwendung](https://support.office.com/article/version-and-build-numbers-of-update-channel-releases-ae942449-1fca-4484-898b-a933ea23def7)
- 
  [Office Online Server-Übersicht](https://docs.microsoft.com/officeonlineserver/office-online-server-overview)

## <a name="office-common-api-requirement-sets"></a>Allgemeine Office-API-Anforderungssätze

Informationen zu allgemeinen API-Anforderungssätzen finden Sie unter [Allgemeine Office-API-Anforderungssätze](office-add-in-requirement-sets.md).

## <a name="see-also"></a>Siehe auch

- [Office-Versionen und Anforderungssätze](https://docs.microsoft.com/office/dev/add-ins/develop/office-versions-and-requirement-sets)
- 
  [Angeben von Office-Hosts und API-Anforderungen](https://docs.microsoft.com/office/dev/add-ins/develop/specify-office-hosts-and-api-requirements)
- [XML-Manifest für Office-Add-Ins](https://docs.microsoft.com/office/dev/add-ins/develop/add-in-manifests)
