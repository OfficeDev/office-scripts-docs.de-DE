---
title: Plattformbeschränkungen und-Anforderungen mit Office-Skripts
description: Ressourcengrenzwerte und Browserunterstützung für Office-Skripts bei Verwendung mit Excel im Internet
ms.date: 07/23/2020
localization_priority: Normal
ms.openlocfilehash: 6e297cba0b9f984f2d541cc3c441a666f9ebfcef
ms.sourcegitcommit: ff7fde04ce5a66d8df06ed505951c8111e2e9833
ms.translationtype: MT
ms.contentlocale: de-DE
ms.lasthandoff: 08/11/2020
ms.locfileid: "46618159"
---
# <a name="platform-limits-and-requirements-with-office-scripts"></a><span data-ttu-id="30102-103">Plattformbeschränkungen und-Anforderungen mit Office-Skripts</span><span class="sxs-lookup"><span data-stu-id="30102-103">Platform limits and requirements with Office Scripts</span></span>

<span data-ttu-id="30102-104">Es gibt einige Plattformeinschränkungen, die Sie beim Entwickeln von Office-Skripts beachten sollten.</span><span class="sxs-lookup"><span data-stu-id="30102-104">There are some platform limitations of which you should be aware when developing Office Scripts.</span></span> <span data-ttu-id="30102-105">In diesem Artikel werden die Browserunterstützung und Daten Grenzwerte für Office-Skripts für Excel im Internet erläutert.</span><span class="sxs-lookup"><span data-stu-id="30102-105">This article details the browser support and data limits for Office Scripts for Excel on the web.</span></span>

## <a name="browser-support"></a><span data-ttu-id="30102-106">Browserunterstützung</span><span class="sxs-lookup"><span data-stu-id="30102-106">Browser support</span></span>

<span data-ttu-id="30102-107">Office-Skripts funktionieren in jedem Browser, [der Office für das Internet unterstützt](https://support.microsoft.com/office/ad1303e0-a318-47aa-b409-d3a5eb44e452).</span><span class="sxs-lookup"><span data-stu-id="30102-107">Office Scripts work in any browser that [supports Office for the web](https://support.microsoft.com/office/ad1303e0-a318-47aa-b409-d3a5eb44e452).</span></span> <span data-ttu-id="30102-108">Einige JavaScript-Funktionen werden in Internet Explorer 11 jedoch nicht unterstützt (IE 11).</span><span class="sxs-lookup"><span data-stu-id="30102-108">However, some JavaScript features aren't supported in Internet Explorer 11 (IE 11).</span></span> <span data-ttu-id="30102-109">Alle in [ES6 oder höher](https://www.w3schools.com/Js/js_es6.asp) eingeführten Features funktionieren nicht mit IE 11.</span><span class="sxs-lookup"><span data-stu-id="30102-109">Any features introduced in [ES6 or later](https://www.w3schools.com/Js/js_es6.asp) won't work with IE 11.</span></span> <span data-ttu-id="30102-110">Wenn Personen in Ihrer Organisation diesen Browser weiterhin verwenden, müssen Sie Ihre Skripts in dieser Umgebung testen, wenn Sie Sie freigeben.</span><span class="sxs-lookup"><span data-stu-id="30102-110">If people in your organization still use that browser, be sure to test your scripts in that environment when sharing them.</span></span>

### <a name="third-party-cookies"></a><span data-ttu-id="30102-111">Drittanbieter-Cookies</span><span class="sxs-lookup"><span data-stu-id="30102-111">Third-party cookies</span></span>

<span data-ttu-id="30102-112">Ihr Browser benötigt Cookies von Drittanbietern, die zum Anzeigen der Registerkarte " **automatisieren** " in Excel im Internet aktiviert sind.</span><span class="sxs-lookup"><span data-stu-id="30102-112">Your browser needs third-party cookies enabled to show the **Automate** tab in Excel on the web.</span></span> <span data-ttu-id="30102-113">Überprüfen Sie Ihre Browsereinstellungen, wenn die Registerkarte nicht angezeigt wird.</span><span class="sxs-lookup"><span data-stu-id="30102-113">Check your browser settings if the tab isn't being displayed.</span></span> <span data-ttu-id="30102-114">Wenn Sie eine private Browsersitzung verwenden, müssen Sie diese Einstellung möglicherweise jedes Mal erneut aktivieren.</span><span class="sxs-lookup"><span data-stu-id="30102-114">If you're using a private browser session, you may need to re-enable this setting each time.</span></span>

> [!NOTE]
> <span data-ttu-id="30102-115">Einige Browser bezeichnen diese Einstellung als "alle Cookies" anstelle von "Drittanbieter-Cookies".</span><span class="sxs-lookup"><span data-stu-id="30102-115">Some browsers refer to this setting as "all cookies", instead of "third-party cookies".</span></span>

## <a name="data-limits"></a><span data-ttu-id="30102-116">Beschränkungen für Daten</span><span class="sxs-lookup"><span data-stu-id="30102-116">Data limits</span></span>

<span data-ttu-id="30102-117">Es gibt Grenzen dafür, wie viel Excel-Daten gleichzeitig übertragen werden können und wie viele einzelne Power-Automatisierungs Transaktionen durchgeführt werden können.</span><span class="sxs-lookup"><span data-stu-id="30102-117">There are limits on how much Excel data can be transferred at once and how many individual Power Automate transactions can be conducted.</span></span>

### <a name="excel"></a><span data-ttu-id="30102-118">Excel</span><span class="sxs-lookup"><span data-stu-id="30102-118">Excel</span></span>

<span data-ttu-id="30102-119">Excel für das Internet hat die folgenden Einschränkungen beim Ausführen von Aufrufen der Arbeitsmappe über ein Skript:</span><span class="sxs-lookup"><span data-stu-id="30102-119">Excel for the web has the following limitations when making calls to the workbook through a script:</span></span>

- <span data-ttu-id="30102-120">Anforderungen und Antworten sind auf **5MB**limitiert.</span><span class="sxs-lookup"><span data-stu-id="30102-120">Requests and responses are limited to **5MB**.</span></span>
- <span data-ttu-id="30102-121">Ein Bereich ist auf **5 Millionen Zellen**limitiert.</span><span class="sxs-lookup"><span data-stu-id="30102-121">A range is limited to **five million cells**.</span></span>

<span data-ttu-id="30102-122">Wenn beim Umgang mit großen Datasets Fehler auftreten, verwenden Sie mehrere kleinere Bereiche anstelle größerer Bereiche.</span><span class="sxs-lookup"><span data-stu-id="30102-122">If you're encountering errors when dealing with large datasets, try using multiple smaller ranges instead of larger ranges.</span></span> <span data-ttu-id="30102-123">Sie können auch APIs wie [Range. getSpecialCells](/javascript/api/office-scripts/excelscript/excelscript.range#getspecialcells-celltype--cellvaluetype-) , um bestimmte Zellen anstelle von großen Bereichen Ziel.</span><span class="sxs-lookup"><span data-stu-id="30102-123">You can also APIs like [Range.getSpecialCells](/javascript/api/office-scripts/excelscript/excelscript.range#getspecialcells-celltype--cellvaluetype-) to target specific cells instead of large ranges.</span></span>

### <a name="power-automate"></a><span data-ttu-id="30102-124">Power Automate</span><span class="sxs-lookup"><span data-stu-id="30102-124">Power Automate</span></span>

<span data-ttu-id="30102-125">Bei der Verwendung von Office-Skripts mit Power Automation sind Sie auf **200 Anrufe pro Tag**limitiert.</span><span class="sxs-lookup"><span data-stu-id="30102-125">When using Office Scripts with Power Automate, you're limited to **200 calls per day**.</span></span> <span data-ttu-id="30102-126">Dieses Limit wird auf 12:00 Uhr UTC zurückgesetzt.</span><span class="sxs-lookup"><span data-stu-id="30102-126">This limit resets at 12:00 AM UTC.</span></span>

<span data-ttu-id="30102-127">Die Power Automation-Plattform hat auch Nutzungseinschränkungen, die im Artikel [Limits and Configuration in Power Automation](/power-automate/limits-and-config)zu finden sind.</span><span class="sxs-lookup"><span data-stu-id="30102-127">The Power Automate platform also has usage limitations, which can be found in the article [Limits and configuration in Power Automate](/power-automate/limits-and-config).</span></span>

## <a name="see-also"></a><span data-ttu-id="30102-128">Siehe auch</span><span class="sxs-lookup"><span data-stu-id="30102-128">See also</span></span>

- [<span data-ttu-id="30102-129">Behandeln von Problemen mit Office-Skripts</span><span class="sxs-lookup"><span data-stu-id="30102-129">Troubleshooting Office Scripts</span></span>](troubleshooting.md)
- [<span data-ttu-id="30102-130">Rückgängigmachen der Auswirkungen eines Office-Skripts</span><span class="sxs-lookup"><span data-stu-id="30102-130">Undo the effects of an Office Script</span></span>](undo.md)
- [<span data-ttu-id="30102-131">Verbessern der Leistung Ihrer Office-Skripts</span><span class="sxs-lookup"><span data-stu-id="30102-131">Improve the performance of your Office Scripts</span></span>](../develop/web-client-performance.md)
- [<span data-ttu-id="30102-132">Grundlegendes zur Skripterstellung für Office-Skripts in Excel im Internet</span><span class="sxs-lookup"><span data-stu-id="30102-132">Scripting Fundamentals for Office Scripts in Excel on the web</span></span>](../develop/scripting-fundamentals.md)
