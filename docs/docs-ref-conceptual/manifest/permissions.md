# <a name="permissions-element"></a><span data-ttu-id="6315f-101">Permissions-Element</span><span class="sxs-lookup"><span data-stu-id="6315f-101">Permissions element</span></span>

<span data-ttu-id="6315f-102">Gibt die Ebene des API-Zugriffs für Ihr Office-Add-In an. Sie müssen Berechtigungen auf der Grundlage der „geringsten Rechte“ anfordern.</span><span class="sxs-lookup"><span data-stu-id="6315f-102">Specifies the level of API access for your Office Add-in; you should request permissions based on the principle of least privilege.</span></span>

<span data-ttu-id="6315f-103">**Add-In-Typ:** Inhalt, Aufgabenbereich, E-Mail</span><span class="sxs-lookup"><span data-stu-id="6315f-103">**Add-in type:** Content, Task pane, Mail</span></span>

## <a name="syntax"></a><span data-ttu-id="6315f-104">Syntax</span><span class="sxs-lookup"><span data-stu-id="6315f-104">Syntax</span></span>

<span data-ttu-id="6315f-105">Für Inhalts- und Aufgabenbereich-Add-Ins:</span><span class="sxs-lookup"><span data-stu-id="6315f-105">For content and task pane add-ins:</span></span>

```XML
 <Permissions> [Restricted | ReadDocument | ReadAllDocument | WriteDocument | ReadWriteDocument]</Permissions>
```

<span data-ttu-id="6315f-106">Für Mail-add-ins</span><span class="sxs-lookup"><span data-stu-id="6315f-106">For mail add-ins</span></span>

```XML
 <Permissions>[Restricted | ReadItem | ReadWriteItem | ReadWriteMailbox]</Permissions>
```

## <a name="contained-in"></a><span data-ttu-id="6315f-107">Enthalten in</span><span class="sxs-lookup"><span data-stu-id="6315f-107">Contained in</span></span>

[<span data-ttu-id="6315f-108">OfficeApp</span><span class="sxs-lookup"><span data-stu-id="6315f-108">OfficeApp</span></span>](officeapp.md)

## <a name="remarks"></a><span data-ttu-id="6315f-109">Hinweise</span><span class="sxs-lookup"><span data-stu-id="6315f-109">Remarks</span></span>

<span data-ttu-id="6315f-110">Weitere Informationen finden Sie unter [Anfordern von Berechtigungen für API-Verwendung in Inhalts- und Aufgabenbereich-Add-Ins](https://docs.microsoft.com/office/dev/add-ins/develop/requesting-permissions-for-api-use-in-content-and-task-pane-add-ins) und [Grundlegendes zu Outlook-Add-In-Berechtigungen](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions).</span><span class="sxs-lookup"><span data-stu-id="6315f-110">For more detail, see [Requesting permissions for API use in content and task pane add-ins](https://docs.microsoft.com/office/dev/add-ins/develop/requesting-permissions-for-api-use-in-content-and-task-pane-add-ins) and [Understanding Outlook add-in permissions](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions).</span></span>
