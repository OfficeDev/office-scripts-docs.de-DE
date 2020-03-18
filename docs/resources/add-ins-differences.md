---
title: Unterschiede zwischen Office-Skripts und Office-Add-ins
description: Das Verhalten und die API-Unterschiede zwischen Office-Skripts und Office-Add-Ins.
ms.date: 12/12/2019
localization_priority: Normal
ms.openlocfilehash: 4626afb66b54c94a72f29b039c601435c089d64d
ms.sourcegitcommit: b075eed5a6f275274fbbf6d62633219eac416f26
ms.translationtype: MT
ms.contentlocale: de-DE
ms.lasthandoff: 03/10/2020
ms.locfileid: "42700230"
---
# <a name="differences-between-office-scripts-and-office-add-ins"></a><span data-ttu-id="b95f9-103">Unterschiede zwischen Office-Skripts und Office-Add-ins</span><span class="sxs-lookup"><span data-stu-id="b95f9-103">Differences between Office Scripts and Office Add-ins</span></span>

<span data-ttu-id="b95f9-104">Office-Add-Ins-und Office-Skripts haben viele Gemeinsamkeiten.</span><span class="sxs-lookup"><span data-stu-id="b95f9-104">Office Add-ins and Office Scripts have a lot in common.</span></span> <span data-ttu-id="b95f9-105">Beide bieten eine automatisierte Steuerung einer Excel-Arbeitsmappe über `Excel` den Namespace der Office-JavaScript-API.</span><span class="sxs-lookup"><span data-stu-id="b95f9-105">They both offer automated control of an Excel workbook through the `Excel` namespace of the Office JavaScript API.</span></span> <span data-ttu-id="b95f9-106">Office-Skripts sind jedoch in ihrem Umfang beschränkter.</span><span class="sxs-lookup"><span data-stu-id="b95f9-106">However, Office Scripts are more limited in their scope.</span></span>

<span data-ttu-id="b95f9-107">Office-Skripts werden bis zum Abschluss mit einer manuellen Schaltfläche drücken, während Office-Add-Ins auf Benutzerinteraktion basieren und dauerhaft bleiben, während die Arbeitsmappe verwendet wird.</span><span class="sxs-lookup"><span data-stu-id="b95f9-107">Office Scripts run to completion with a manual button press, whereas Office Add-ins rely on user interaction and persist while the workbook is in use.</span></span> <span data-ttu-id="b95f9-108">Wenn Sie feststellen, dass Ihre Excel-Erweiterung die Funktionen der Skriptplattform überschreiten muss, lesen Sie die [Office-Add-ins Dokumentation](/office/dev/add-ins) , um mehr über Office-Add-Ins zu erfahren.</span><span class="sxs-lookup"><span data-stu-id="b95f9-108">If you find that your Excel extension needs to exceed the scripting platform's capabilities, visit the [Office Add-ins documentation](/office/dev/add-ins) to learn more about Office Add-ins.</span></span>

<span data-ttu-id="b95f9-109">Der Rest dieses Artikels beschreibt die Hauptunterschiede zwischen Office-Add-Ins-und Office-Skripts.</span><span class="sxs-lookup"><span data-stu-id="b95f9-109">The rest of this article describes on the main differences between Office Add-ins and Office Scripts.</span></span>

## <a name="platform-support"></a><span data-ttu-id="b95f9-110">Plattformunterstützung</span><span class="sxs-lookup"><span data-stu-id="b95f9-110">Platform Support</span></span>

<span data-ttu-id="b95f9-111">Office-Add-Ins sind plattformübergreifend.</span><span class="sxs-lookup"><span data-stu-id="b95f9-111">Office Add-ins are cross-platform.</span></span> <span data-ttu-id="b95f9-112">Sie funktionieren auf Windows-Desktop-, Mac-, IOS-und Webplattformen und bieten die gleiche Benutzeroberfläche.</span><span class="sxs-lookup"><span data-stu-id="b95f9-112">They work across Windows desktop, Mac, iOS, and web platforms and provide the same experience on each.</span></span> <span data-ttu-id="b95f9-113">Jede Ausnahme hiervon wird in der Dokumentation der einzelnen API notiert.</span><span class="sxs-lookup"><span data-stu-id="b95f9-113">Any exception to this is noted in the documentation of the individual API.</span></span>

<span data-ttu-id="b95f9-114">Office-Skripts werden derzeit nur von für Excel im Internet unterstützt.</span><span class="sxs-lookup"><span data-stu-id="b95f9-114">Office Scripts are currently only supported by for Excel on the web.</span></span> <span data-ttu-id="b95f9-115">Alle Aufzeichnung, Bearbeitung und Ausführung erfolgt auf der Webplattform.</span><span class="sxs-lookup"><span data-stu-id="b95f9-115">All recording, editing, and running is done on the web platform.</span></span>

## <a name="apis"></a><span data-ttu-id="b95f9-116">APIs</span><span class="sxs-lookup"><span data-stu-id="b95f9-116">APIs</span></span>

<span data-ttu-id="b95f9-117">Office-Skripts unterstützen die meisten der Excel-JavaScript-APIs, was bedeutet, dass es eine Vielzahl von Funktionsüberschneidungen zwischen den beiden Plattformen gibt.</span><span class="sxs-lookup"><span data-stu-id="b95f9-117">Office Scripts support most of the Excel JavaScript APIs, which means there's  a lot of functionality overlap between the two platforms.</span></span> <span data-ttu-id="b95f9-118">Es gibt zwei Ausnahmen: Ereignisse und allgemeine APIs.</span><span class="sxs-lookup"><span data-stu-id="b95f9-118">There are two exceptions: events and Common APIs.</span></span>

### <a name="events"></a><span data-ttu-id="b95f9-119">Ereignisse</span><span class="sxs-lookup"><span data-stu-id="b95f9-119">Events</span></span>

<span data-ttu-id="b95f9-120">Office-Skripts unterstützen keine [Ereignisse](/office/dev/add-ins/excel/excel-add-ins-events).</span><span class="sxs-lookup"><span data-stu-id="b95f9-120">Office Scripts do not support [events](/office/dev/add-ins/excel/excel-add-ins-events).</span></span> <span data-ttu-id="b95f9-121">Jedes Skript führt den Code in einer einzigen `main` Methode aus und endet dann.</span><span class="sxs-lookup"><span data-stu-id="b95f9-121">Every script runs the code in a single `main` method, then ends.</span></span> <span data-ttu-id="b95f9-122">Es wird nicht erneut aktiviert, wenn Ereignisse ausgelöst werden und daher keine Ereignisse registrieren können.</span><span class="sxs-lookup"><span data-stu-id="b95f9-122">It does not reactivate when events are triggered, and thus, cannot register events.</span></span>

### <a name="common-apis"></a><span data-ttu-id="b95f9-123">Allgemeine APIs</span><span class="sxs-lookup"><span data-stu-id="b95f9-123">Common APIs</span></span>

<span data-ttu-id="b95f9-124">Office-Skripts können keine [allgemeinen APIs](/javascript/api/office)verwenden.</span><span class="sxs-lookup"><span data-stu-id="b95f9-124">Office Scripts cannot use [Common APIs](/javascript/api/office).</span></span> <span data-ttu-id="b95f9-125">Wenn Sie Authentifizierung, Dialogfenster oder andere Features benötigen, die nur von gängigen APIs unterstützt werden, müssen Sie wahrscheinlich ein Office-Add-in anstelle eines Office-Skripts erstellen.</span><span class="sxs-lookup"><span data-stu-id="b95f9-125">If you need authentication, dialog windows, or other features that are only supported by Common APIs, you'll likely need to create an Office Add-in instead of an Office Script.</span></span>

## <a name="see-also"></a><span data-ttu-id="b95f9-126">Siehe auch</span><span class="sxs-lookup"><span data-stu-id="b95f9-126">See also</span></span>

- [<span data-ttu-id="b95f9-127">Office-Skripts in Excel im Internet</span><span class="sxs-lookup"><span data-stu-id="b95f9-127">Office Scripts in Excel on the web</span></span>](../overview/excel.md)
- [<span data-ttu-id="b95f9-128">Problembehandlung bei Office-Skripts</span><span class="sxs-lookup"><span data-stu-id="b95f9-128">Troubleshooting Office Scripts</span></span>](../testing/troubleshooting.md)
- [<span data-ttu-id="b95f9-129">Erstellen eines Excel-Aufgabenbereich-Add-Ins</span><span class="sxs-lookup"><span data-stu-id="b95f9-129">Build an Excel task pane add-in</span></span>](/office/dev/add-ins/quickstarts/excel-quickstart-jquery)