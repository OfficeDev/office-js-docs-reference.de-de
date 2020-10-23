---
layout: LandingPage
ms.topic: landing-page
title: JavaScript-API-Referenz für Office
description: Die JavaScript-APIs für Office nach Host und Version.
author: o365devx
ms.author: o365devx
ms.prod: non-product-specific
localization_priority: Priority
ms.date: 10/19/2020
ms.openlocfilehash: 1bd13892aaa172d958f5e9fcbb0e63871fd2cdd7
ms.sourcegitcommit: d5885aa1eaab2bbe8ddba2e2fdc618ac99657ef3
ms.translationtype: HT
ms.contentlocale: de-DE
ms.lasthandoff: 10/23/2020
ms.locfileid: "48739818"
---
# <a name="office-add-ins-javascript-api-reference"></a>Office-Add-Ins JavaScript-API-Referenz

Die JavaScript-API für Office ermöglicht Ihnen, Webanwendungen zu erstellen, die mit den Objektmodellen in Office-Hostanwendungen interagieren. Verwenden Sie diesen Abschnitt, um mehr über die Klassen, Methoden und anderen Typen zu erfahren, die zum Erstellen von Office-Add-Ins verfügbar sind.

Die folgende Liste enthält APIs für die [unterstützten Office-Hostanwendungen](/office/dev/add-ins/overview/office-add-in-availability). Allgemeine API-Links umfassen alle APIs, die nicht einem bestimmten Host zugeordnet sind (wie im [Office Allgemeinen API-Anforderungssatz](/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets) beschrieben). Andere Elemente verweisen auf eine Version der API-Referenzdokumentation für diesen Host auf der Grundlage eines Anforderungssatzes. Die Referenzdokumentation wird so versioniert, dass sie alle APIs bis einschließlich dieses Anforderungssatzes enthält (z. B. ExcelApi 1.3 zeigt APIs in ExcelApi 1.1, 1.2, 1.3 sowie die Allgemeinen APIs).

`ExcelApiOnline 1.1` ist ein spezieller Anforderungssatz. Darin sind die neuesten APIs für Excel für das Web enthalten, die jedoch möglicherweise noch nicht auf allen Plattformen vollständig unterstützt werden. Weitere Informationen finden Sie unter [Excel JavaScript API online-only Anforderungssatz](/office/dev/add-ins/reference/requirement-sets/excel-api-online-requirement-set).

> [!TIP]
> Sie können die Version einer Referenzseite jederzeit über das Dropdownmenü "Filterauswahl" oberhalb des Inhaltsverzeichnisses ändern. Wenn die Seite in dieser bestimmten Version nicht vorhanden ist, werden Sie zur aktuellen Version zurückgeleitet.

<h2>Office-Hosts</h2>

<ul class="cardsK panelContent cols cols3">
    <li>
        <a class="card x-hidden-focus">
            <div class="cardImageOuter">
                <div class="cardImage">
                    <img src="/javascript/api/overview/images/logo-excel.svg" alt="Excel add-ins" />
                </div>
            </div>
            <div class="cardText">
                <h3>Excel-APIs</h3>
                <ul>
                    <li><a style="font-size: 1rem;" href="/javascript/api/excel?view=excel-js-preview">ExcelApi-Vorschau</a></li>
                    <li><a style="font-size: 1rem;" href="/javascript/api/excel?view=excel-js-online">ExcelApiOnline 1.1</a></li>
                    <li><a style="font-size: 1rem;" href="/javascript/api/excel?view=excel-js-1.12">ExcelApi 1.12</a></li>
                    <li><a style="font-size: 1rem;" href="/javascript/api/excel?view=excel-js-1.11">ExcelApi 1.11</a></li>
                    <li><a style="font-size: 1rem;" href="/javascript/api/excel?view=excel-js-1.10">ExcelApi 1.10</a></li>
                    <li><a style="font-size: 1rem;" href="/javascript/api/excel?view=excel-js-1.9">ExcelApi 1.9</a></li>
                    <li><a style="font-size: 1rem;" href="/javascript/api/excel?view=excel-js-1.8">ExcelApi 1.8</a></li>
                    <li><a style="font-size: 1rem;" href="/javascript/api/excel?view=excel-js-1.7">ExcelApi 1.7</a></li>
                    <li><a style="font-size: 1rem;" href="/javascript/api/excel?view=excel-js-1.6">ExcelApi 1.6</a></li>
                    <li><a style="font-size: 1rem;" href="/javascript/api/excel?view=excel-js-1.5">ExcelApi 1.5</a></li>
                    <li><a style="font-size: 1rem;" href="/javascript/api/excel?view=excel-js-1.4">ExcelApi 1.4</a></li>
                    <li><a style="font-size: 1rem;" href="/javascript/api/excel?view=excel-js-1.3">ExcelApi 1.3</a></li>
                    <li><a style="font-size: 1rem;" href="/javascript/api/excel?view=excel-js-1.2">ExcelApi 1.2</a></li>
                    <li><a style="font-size: 1rem;" href="/javascript/api/excel?view=excel-js-1.1">ExcelApi 1.1</a></li>
                    <li><a style="font-size: 1rem;" href="/javascript/api/office?view=excel-js-preview">Allgemeine APIs</a></li>
                </ul>
            </div>
        </a>
    </li>
    <li>
        <a class="card x-hidden-focus">
            <div class="cardImageOuter">
                <div class="cardImage">
                    <img src="/javascript/api/overview/images/logo-outlook.svg" alt="Outlook add-ins" />
                </div>
            </div>
            <div class="cardText">
                <h3>Outlook-APIs</h3>
                <ul>
                    <li><a style="font-size: 1rem;" href="/javascript/api/outlook?view=outlook-js-preview">Mailbox-Vorschau</a></li>
                    <li><a style="font-size: 1rem;" href="/javascript/api/outlook?view=outlook-js-1.9">Mailbox 1.9</a></li>
                    <li><a style="font-size: 1rem;" href="/javascript/api/outlook?view=outlook-js-1.8">Mailbox 1.8</a></li>
                    <li><a style="font-size: 1rem;" href="/javascript/api/outlook?view=outlook-js-1.7">Mailbox 1.7</a></li>
                    <li><a style="font-size: 1rem;" href="/javascript/api/outlook?view=outlook-js-1.6">Mailbox 1.6</a></li>
                    <li><a style="font-size: 1rem;" href="/javascript/api/outlook?view=outlook-js-1.5">Mailbox 1.5</a></li>
                    <li><a style="font-size: 1rem;" href="/javascript/api/outlook?view=outlook-js-1.4">Mailbox 1.4</a></li>
                    <li><a style="font-size: 1rem;" href="/javascript/api/outlook?view=outlook-js-1.3">Mailbox 1.3</a></li>
                    <li><a style="font-size: 1rem;" href="/javascript/api/outlook?view=outlook-js-1.2">Mailbox 1.2</a></li>
                    <li><a style="font-size: 1rem;" href="/javascript/api/outlook?view=outlook-js-1.1">Mailbox 1.1</a></li>
                    <li><a style="font-size: 1rem;" href="/javascript/api/office?view=outlook-js-preview">Allgemeine APIs</a></li>
                </ul>
            </div>
        </a>
    </li>
    <li>
        <a class="card x-hidden-focus">
            <div class="cardImageOuter">
                <div class="cardImage">
                    <img src="/javascript/api/overview/images/logo-word.svg" alt="Word add-ins" />
                </div>
            </div>
            <div class="cardText">
                <h3>Word-APIs</h3>
                <ul>
                    <li><a style="font-size: 1rem;" href="/javascript/api/word?view=word-js-preview">WordApi-Vorschau</a></li>
                    <li><a style="font-size: 1rem;" href="/javascript/api/word?view=word-js-1.3">WordApi 1.3</a></li>
                    <li><a style="font-size: 1rem;" href="/javascript/api/word?view=word-js-1.2">WordApi 1.2</a></li>
                    <li><a style="font-size: 1rem;" href="/javascript/api/word?view=word-js-1.1">WordApi 1.1</a></li>
                    <li><a style="font-size: 1rem;" href="/javascript/api/office?view=word-js-preview">Allgemeine APIs</a></li>
                </ul>
            </div>
        </a>
    </li>
    <li>
        <a class="card x-hidden-focus">
            <div class="cardImageOuter">
                <div class="cardImage">
                    <img src="/javascript/api/overview/images/logo-onenote.svg" alt="OneNote add-ins" />
                </div>
            </div>
            <div class="cardText">
                <h3>OneNote-APIs</h3>
                <ul>
                    <li><a style="font-size: 1rem;" href="/javascript/api/onenote?view=onenote-js-1.1">OneNoteApi 1.1</a></li>
                    <li><a style="font-size: 1rem;" href="/javascript/api/office?view=onenote-js-1.1">Allgemeine APIs</a></li>
                </ul>
            </div>
        </a>
    </li>
    <li>
        <a class="card x-hidden-focus">
            <div class="cardImageOuter">
                <div class="cardImage">
                    <img src="/javascript/api/overview/images/logo-visio.svg" alt="Visio add-ins" />
                </div>
            </div>
            <div class="cardText">
                <h3>Visio-APIs</h3>
                <ul>
                    <li><a style="font-size: 1rem;" href="/javascript/api/visio?view=visio-js-1.1">VisioApi 1.1</a></li>
                </ul>
            </div>
        </a>
    </li>
    <li>
        <a class="card x-hidden-focus">
            <div class="cardImageOuter">
                <div class="cardImage">
                    <img src="/javascript/api/overview/images/logo-powerpoint.svg" alt="PowerPoint add-ins" />
                </div>
            </div>
            <div class="cardText">
                <h3>PowerPoint-APIs</h3>
                <ul>
                    <li><a style="font-size: 1rem;" href="/javascript/api/powerpoint?view=powerpoint-js-preview">PowerPointApi Preview</a></li>
                    <li><a style="font-size: 1rem;" href="/javascript/api/powerpoint?view=powerpoint-js-1.1">PowerPointApi 1.1</a></li>
                    <li><a style="font-size: 1rem;" href="/javascript/api/office?view=powerpoint-js-preview">Allgemeine APIs</a></li>
                </ul>
            </div>
        </a>
    </li>
    <li>
        <a class="card x-hidden-focus">
            <div class="cardImageOuter">
                <div class="cardImage">
                    <img src="/javascript/api/overview/images/logo-project.svg" alt="Project add-ins" />
                </div>
            </div>
            <div class="cardText">
                <h3>Projekt-APIs</h3>
                <ul>
                    <li><a style="font-size: 1rem;" href="/javascript/api/office?view=common-js">Ausschließlich Allgemeine APIs</a></li>
                </ul>
            </div>
        </a>
    </li>
</ul>

> [!NOTE]
> Wenn Sie JavaScript-APIs zum Entwickeln von Office-Scripts suchen, gehen Sie auf [Referenzdokumentation zur Office Scripts-API](/javascript/api/office-scripts/overview).

## <a name="see-also"></a>Weitere Informationen:

- [Informationen zu Office-Add-Ins](/office/dev/add-ins/overview)
- [Host- und Plattformverfügbarkeit von Office-Add-Ins](/office/dev/add-ins/overview/office-add-in-availability)
- [Office-Versionen und Anforderungssätze](/office/dev/add-ins/develop/office-versions-and-requirement-sets)
- [Entdecken Sie die Office JavaScript-API mithilfe von Script Lab](/office/dev/add-ins/overview/explore-with-script-lab)
