# <a name="allowsnapshot-element"></a><span data-ttu-id="5583f-101">AllowSnapshot-Element</span><span class="sxs-lookup"><span data-stu-id="5583f-101">AllowSnapshot element</span></span>

<span data-ttu-id="5583f-102">Gibt an, ob eine Momentaufnahme Ihres Inhalts-Add-Ins mit dem Hostdokument gespeichert wird.</span><span class="sxs-lookup"><span data-stu-id="5583f-102">Specifies whether a snapshot image of your content add-in is saved with the host document.</span></span>

<span data-ttu-id="5583f-103">**Add-In-Typ:** Inhalt</span><span class="sxs-lookup"><span data-stu-id="5583f-103">**Add-in type:** Content</span></span>

## <a name="syntax"></a><span data-ttu-id="5583f-104">Syntax</span><span class="sxs-lookup"><span data-stu-id="5583f-104">Syntax</span></span>

```XML
<AllowSnapshot> [true | false]</AllowSnapshot>
```

## <a name="contained-in"></a><span data-ttu-id="5583f-105">Enthalten in</span><span class="sxs-lookup"><span data-stu-id="5583f-105">Contained in</span></span>

[<span data-ttu-id="5583f-106">OfficeApp</span><span class="sxs-lookup"><span data-stu-id="5583f-106">OfficeApp</span></span>](officeapp.md)

## <a name="remarks"></a><span data-ttu-id="5583f-107">Hinweise</span><span class="sxs-lookup"><span data-stu-id="5583f-107">Remarks</span></span>

 > [!IMPORTANT]
 > <span data-ttu-id="5583f-108">**AllowSnapshot** ist `true` standardmäßig.</span><span class="sxs-lookup"><span data-stu-id="5583f-108">**AllowSnapshot** is `true` by default.</span></span> <span data-ttu-id="5583f-109">Dadurch wird ein Bild des Add-Ins für Benutzer sichtbar, die das Dokument in einer Version der Hostanwendung öffnen, die Office-Add-Ins nicht unterstützt, oder es wird ein statisches Bild des Add-Ins bereitgestellt, wenn die Hostanwendung keine Verbindung zu dem Server herstellen kann, auf dem das Add-In gehostet wird.</span><span class="sxs-lookup"><span data-stu-id="5583f-109">This makes an image of the add-in visible for users that open the document in a version of the host application that doesn't support Office Add-ins, or provides a static image of the add-in if the host application can't connect to the server hosting the add-in.</span></span> <span data-ttu-id="5583f-110">Dies bedeutet jedoch auch, dass auf möglicherweise vertrauliche Informationen, die in dem Add-in angezeigt werden, direkt aus dem Dokument, das das Add-In hostet, zugegriffen werden kann.</span><span class="sxs-lookup"><span data-stu-id="5583f-110">However, this also means that potentially sensitive information displayed in the add-in can be accessed directly from the document hosting the add-in.</span></span>

