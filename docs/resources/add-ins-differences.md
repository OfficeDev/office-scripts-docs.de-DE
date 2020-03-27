---
title: Unterschiede zwischen Office-Skripts und Office-Add-ins
description: Das Verhalten und die API-Unterschiede zwischen Office-Skripts und Office-Add-Ins.
ms.date: 03/23/2020
localization_priority: Normal
ms.openlocfilehash: 2290d4e34b7a7286d67443de9e9c64bad4fcd4b7
ms.sourcegitcommit: d556aaefac80e55f53ac56b7f6ecbc657ebd426f
ms.translationtype: MT
ms.contentlocale: de-DE
ms.lasthandoff: 03/26/2020
ms.locfileid: "42978715"
---
# <a name="differences-between-office-scripts-and-office-add-ins"></a><span data-ttu-id="a5075-103">Unterschiede zwischen Office-Skripts und Office-Add-ins</span><span class="sxs-lookup"><span data-stu-id="a5075-103">Differences between Office Scripts and Office Add-ins</span></span>

<span data-ttu-id="a5075-104">Office-Add-Ins-und Office-Skripts haben viele Gemeinsamkeiten.</span><span class="sxs-lookup"><span data-stu-id="a5075-104">Office Add-ins and Office Scripts have a lot in common.</span></span> <span data-ttu-id="a5075-105">Beide bieten eine automatisierte Steuerung einer Excel-Arbeitsmappe über `Excel` den Namespace der Office-JavaScript-API.</span><span class="sxs-lookup"><span data-stu-id="a5075-105">They both offer automated control of an Excel workbook through the `Excel` namespace of the Office JavaScript API.</span></span> <span data-ttu-id="a5075-106">Office-Skripts sind jedoch in ihrem Umfang beschränkter.</span><span class="sxs-lookup"><span data-stu-id="a5075-106">However, Office Scripts are more limited in their scope.</span></span>

![Ein Diagramm mit vier Quadranten, in dem die Fokusbereiche für unterschiedliche Office-Erweiterbarkeits Lösungen angezeigt werden.](../images/office-programmability-diagram.png)

<span data-ttu-id="a5075-109">Office-Skripts werden bis zum Abschluss mit einer manuellen Tastendruck oder als Schritt in [Power automatisieren](https://flow.microsoft.com/)ausgeführt, während Office-Add-Ins bleiben, während Ihre Aufgabenbereiche geöffnet sind.</span><span class="sxs-lookup"><span data-stu-id="a5075-109">Office Scripts run to completion with a manual button press or as a step in [Power Automate](https://flow.microsoft.com/), whereas Office Add-ins persist while their task panes are open.</span></span> <span data-ttu-id="a5075-110">Dies bedeutet, dass die Add-Ins den Status während einer Sitzung beibehalten können, während Office-Skripts keinen internen Status zwischen Läufen aufrecht erhalten.</span><span class="sxs-lookup"><span data-stu-id="a5075-110">This means the add-ins can maintain state during a session, whereas Office Scripts do not maintain an internal state between runs.</span></span> <span data-ttu-id="a5075-111">Wenn Sie feststellen, dass Ihre Excel-Erweiterung die Funktionen der Skriptplattform überschreiten muss, lesen Sie die [Office-Add-ins Dokumentation](/office/dev/add-ins) , um mehr über Office-Add-Ins zu erfahren.</span><span class="sxs-lookup"><span data-stu-id="a5075-111">If you find that your Excel extension needs to exceed the scripting platform's capabilities, visit the [Office Add-ins documentation](/office/dev/add-ins) to learn more about Office Add-ins.</span></span>

<span data-ttu-id="a5075-112">Der Rest dieses Artikels beschreibt die Hauptunterschiede zwischen Office-Add-Ins-und Office-Skripts.</span><span class="sxs-lookup"><span data-stu-id="a5075-112">The rest of this article describes on the main differences between Office Add-ins and Office Scripts.</span></span>

## <a name="platform-support"></a><span data-ttu-id="a5075-113">Plattformunterstützung</span><span class="sxs-lookup"><span data-stu-id="a5075-113">Platform Support</span></span>

<span data-ttu-id="a5075-114">Office-Add-Ins sind plattformübergreifend.</span><span class="sxs-lookup"><span data-stu-id="a5075-114">Office Add-ins are cross-platform.</span></span> <span data-ttu-id="a5075-115">Sie funktionieren auf Windows-Desktop-, Mac-, IOS-und Webplattformen und bieten die gleiche Benutzeroberfläche.</span><span class="sxs-lookup"><span data-stu-id="a5075-115">They work across Windows desktop, Mac, iOS, and web platforms and provide the same experience on each.</span></span> <span data-ttu-id="a5075-116">Jede Ausnahme hiervon wird in der Dokumentation der einzelnen API notiert.</span><span class="sxs-lookup"><span data-stu-id="a5075-116">Any exception to this is noted in the documentation of the individual API.</span></span>

<span data-ttu-id="a5075-117">Office-Skripts werden derzeit nur von für Excel im Internet unterstützt.</span><span class="sxs-lookup"><span data-stu-id="a5075-117">Office Scripts are currently only supported by for Excel on the web.</span></span> <span data-ttu-id="a5075-118">Alle Aufzeichnung, Bearbeitung und Ausführung erfolgt auf der Webplattform.</span><span class="sxs-lookup"><span data-stu-id="a5075-118">All recording, editing, and running is done on the web platform.</span></span>

## <a name="apis"></a><span data-ttu-id="a5075-119">APIs</span><span class="sxs-lookup"><span data-stu-id="a5075-119">APIs</span></span>

<span data-ttu-id="a5075-120">Office-Skripts unterstützen die meisten der Excel-JavaScript-APIs, was bedeutet, dass es eine Vielzahl von Funktionsüberschneidungen zwischen den beiden Plattformen gibt.</span><span class="sxs-lookup"><span data-stu-id="a5075-120">Office Scripts support most of the Excel JavaScript APIs, which means there's  a lot of functionality overlap between the two platforms.</span></span> <span data-ttu-id="a5075-121">Es gibt zwei Ausnahmen: Ereignisse und allgemeine APIs.</span><span class="sxs-lookup"><span data-stu-id="a5075-121">There are two exceptions: events and Common APIs.</span></span>

### <a name="events"></a><span data-ttu-id="a5075-122">Ereignisse</span><span class="sxs-lookup"><span data-stu-id="a5075-122">Events</span></span>

<span data-ttu-id="a5075-123">Office-Skripts unterstützen keine [Ereignisse](/office/dev/add-ins/excel/excel-add-ins-events).</span><span class="sxs-lookup"><span data-stu-id="a5075-123">Office Scripts do not support [events](/office/dev/add-ins/excel/excel-add-ins-events).</span></span> <span data-ttu-id="a5075-124">Jedes Skript führt den Code in einer einzigen `main` Methode aus und endet dann.</span><span class="sxs-lookup"><span data-stu-id="a5075-124">Every script runs the code in a single `main` method, then ends.</span></span> <span data-ttu-id="a5075-125">Es wird nicht erneut aktiviert, wenn Ereignisse ausgelöst werden und daher keine Ereignisse registrieren können.</span><span class="sxs-lookup"><span data-stu-id="a5075-125">It does not reactivate when events are triggered, and thus, cannot register events.</span></span>

### <a name="common-apis"></a><span data-ttu-id="a5075-126">Allgemeine APIs</span><span class="sxs-lookup"><span data-stu-id="a5075-126">Common APIs</span></span>

<span data-ttu-id="a5075-127">Office-Skripts können keine [allgemeinen APIs](/javascript/api/office)verwenden.</span><span class="sxs-lookup"><span data-stu-id="a5075-127">Office Scripts cannot use [Common APIs](/javascript/api/office).</span></span> <span data-ttu-id="a5075-128">Wenn Sie Authentifizierung, Dialogfenster oder andere Features benötigen, die nur von gängigen APIs unterstützt werden, müssen Sie wahrscheinlich ein Office-Add-in anstelle eines Office-Skripts erstellen.</span><span class="sxs-lookup"><span data-stu-id="a5075-128">If you need authentication, dialog windows, or other features that are only supported by Common APIs, you'll likely need to create an Office Add-in instead of an Office Script.</span></span>

## <a name="see-also"></a><span data-ttu-id="a5075-129">Siehe auch</span><span class="sxs-lookup"><span data-stu-id="a5075-129">See also</span></span>

- [<span data-ttu-id="a5075-130">Office-Skripts in Excel im Web</span><span class="sxs-lookup"><span data-stu-id="a5075-130">Office Scripts in Excel on the web</span></span>](../overview/excel.md)
- [<span data-ttu-id="a5075-131">Unterschiede zwischen Office-Skripts und VBA-Makros</span><span class="sxs-lookup"><span data-stu-id="a5075-131">Differences between Office Scripts and VBA macros</span></span>](vba-differences.md)
- [<span data-ttu-id="a5075-132">Behandeln von Problemen mit Office-Skripts</span><span class="sxs-lookup"><span data-stu-id="a5075-132">Troubleshooting Office Scripts</span></span>](../testing/troubleshooting.md)
- [<span data-ttu-id="a5075-133">Erstellen eines Excel-Aufgabenbereich-Add-Ins</span><span class="sxs-lookup"><span data-stu-id="a5075-133">Build an Excel task pane add-in</span></span>](/office/dev/add-ins/quickstarts/excel-quickstart-jquery)
