---
title: Plattformbeschränkungen und-Anforderungen mit Office-Skripts
description: Ressourcengrenzwerte und Browserunterstützung für Office-Skripts bei Verwendung mit Excel im Internet
ms.date: 10/23/2020
localization_priority: Normal
ms.openlocfilehash: 61f5c55be278ae056014d3b01e4176354d913f87
ms.sourcegitcommit: d3e7681e262bdccc281fcb7b3c719494202e846b
ms.translationtype: MT
ms.contentlocale: de-DE
ms.lasthandoff: 11/06/2020
ms.locfileid: "48930078"
---
# <a name="platform-limits-and-requirements-with-office-scripts"></a><span data-ttu-id="7fa42-103">Plattformbeschränkungen und-Anforderungen mit Office-Skripts</span><span class="sxs-lookup"><span data-stu-id="7fa42-103">Platform limits and requirements with Office Scripts</span></span>

<span data-ttu-id="7fa42-104">Es gibt einige Plattformeinschränkungen, die Sie beim Entwickeln von Office-Skripts beachten sollten.</span><span class="sxs-lookup"><span data-stu-id="7fa42-104">There are some platform limitations of which you should be aware when developing Office Scripts.</span></span> <span data-ttu-id="7fa42-105">In diesem Artikel werden die Browserunterstützung und Daten Grenzwerte für Office-Skripts für Excel im Internet erläutert.</span><span class="sxs-lookup"><span data-stu-id="7fa42-105">This article details the browser support and data limits for Office Scripts for Excel on the web.</span></span>

## <a name="browser-support"></a><span data-ttu-id="7fa42-106">Browserunterstützung</span><span class="sxs-lookup"><span data-stu-id="7fa42-106">Browser support</span></span>

<span data-ttu-id="7fa42-107">Office-Skripts funktionieren in jedem Browser, [der Office für das Internet unterstützt](https://support.microsoft.com/office/ad1303e0-a318-47aa-b409-d3a5eb44e452).</span><span class="sxs-lookup"><span data-stu-id="7fa42-107">Office Scripts work in any browser that [supports Office for the web](https://support.microsoft.com/office/ad1303e0-a318-47aa-b409-d3a5eb44e452).</span></span> <span data-ttu-id="7fa42-108">Einige JavaScript-Funktionen werden in Internet Explorer 11 jedoch nicht unterstützt (IE 11).</span><span class="sxs-lookup"><span data-stu-id="7fa42-108">However, some JavaScript features aren't supported in Internet Explorer 11 (IE 11).</span></span> <span data-ttu-id="7fa42-109">Alle in [ES6 oder höher](https://www.w3schools.com/Js/js_es6.asp) eingeführten Features funktionieren nicht mit IE 11.</span><span class="sxs-lookup"><span data-stu-id="7fa42-109">Any features introduced in [ES6 or later](https://www.w3schools.com/Js/js_es6.asp) won't work with IE 11.</span></span> <span data-ttu-id="7fa42-110">Wenn Personen in Ihrer Organisation diesen Browser weiterhin verwenden, müssen Sie Ihre Skripts in dieser Umgebung testen, wenn Sie Sie freigeben.</span><span class="sxs-lookup"><span data-stu-id="7fa42-110">If people in your organization still use that browser, be sure to test your scripts in that environment when sharing them.</span></span>

[!INCLUDE [Teams support note](../includes/teams-support-note.md)]

### <a name="third-party-cookies"></a><span data-ttu-id="7fa42-111">Drittanbieter-Cookies</span><span class="sxs-lookup"><span data-stu-id="7fa42-111">Third-party cookies</span></span>

<span data-ttu-id="7fa42-112">Ihr Browser benötigt Cookies von Drittanbietern, die zum Anzeigen der Registerkarte " **automatisieren** " in Excel im Internet aktiviert sind.</span><span class="sxs-lookup"><span data-stu-id="7fa42-112">Your browser needs third-party cookies enabled to show the **Automate** tab in Excel on the web.</span></span> <span data-ttu-id="7fa42-113">Überprüfen Sie Ihre Browsereinstellungen, wenn die Registerkarte nicht angezeigt wird.</span><span class="sxs-lookup"><span data-stu-id="7fa42-113">Check your browser settings if the tab isn't being displayed.</span></span> <span data-ttu-id="7fa42-114">Wenn Sie eine private Browsersitzung verwenden, müssen Sie diese Einstellung möglicherweise jedes Mal erneut aktivieren.</span><span class="sxs-lookup"><span data-stu-id="7fa42-114">If you're using a private browser session, you may need to re-enable this setting each time.</span></span>

> [!NOTE]
> <span data-ttu-id="7fa42-115">Einige Browser bezeichnen diese Einstellung als "alle Cookies" anstelle von "Drittanbieter-Cookies".</span><span class="sxs-lookup"><span data-stu-id="7fa42-115">Some browsers refer to this setting as "all cookies", instead of "third-party cookies".</span></span>

#### <a name="instructions-for-adjusting-cookie-settings-in-popular-browsers"></a><span data-ttu-id="7fa42-116">Anweisungen zum Anpassen von Cookieeinstellungen in gängigen Browsern</span><span class="sxs-lookup"><span data-stu-id="7fa42-116">Instructions for adjusting cookie settings in popular browsers</span></span>

- [<span data-ttu-id="7fa42-117">Chrome</span><span class="sxs-lookup"><span data-stu-id="7fa42-117">Chrome</span></span>](https://support.google.com/chrome/answer/95647)
- [<span data-ttu-id="7fa42-118">Microsoft Edge</span><span class="sxs-lookup"><span data-stu-id="7fa42-118">Edge</span></span>](https://support.microsoft.com/microsoft-edge/temporarily-allow-cookies-and-site-data-in-microsoft-edge-597f04f2-c0ce-f08c-7c2b-541086362bd2)
- [<span data-ttu-id="7fa42-119">Firefox</span><span class="sxs-lookup"><span data-stu-id="7fa42-119">Firefox</span></span>](https://support.mozilla.org/kb/disable-third-party-cookies)
- [<span data-ttu-id="7fa42-120">Safari</span><span class="sxs-lookup"><span data-stu-id="7fa42-120">Safari</span></span>](https://support.apple.com/guide/safari/manage-cookies-and-website-data-sfri11471/mac)

## <a name="data-limits"></a><span data-ttu-id="7fa42-121">Beschränkungen für Daten</span><span class="sxs-lookup"><span data-stu-id="7fa42-121">Data limits</span></span>

<span data-ttu-id="7fa42-122">Es gibt Grenzen dafür, wie viel Excel-Daten gleichzeitig übertragen werden können und wie viele einzelne Power-Automatisierungs Transaktionen durchgeführt werden können.</span><span class="sxs-lookup"><span data-stu-id="7fa42-122">There are limits on how much Excel data can be transferred at once and how many individual Power Automate transactions can be conducted.</span></span>

### <a name="excel"></a><span data-ttu-id="7fa42-123">Excel</span><span class="sxs-lookup"><span data-stu-id="7fa42-123">Excel</span></span>

<span data-ttu-id="7fa42-124">Excel für das Internet hat die folgenden Einschränkungen beim Ausführen von Aufrufen der Arbeitsmappe über ein Skript:</span><span class="sxs-lookup"><span data-stu-id="7fa42-124">Excel for the web has the following limitations when making calls to the workbook through a script:</span></span>

- <span data-ttu-id="7fa42-125">Anforderungen und Antworten sind auf **5MB** limitiert.</span><span class="sxs-lookup"><span data-stu-id="7fa42-125">Requests and responses are limited to **5MB**.</span></span>
- <span data-ttu-id="7fa42-126">Ein Bereich ist auf **5 Millionen Zellen** limitiert.</span><span class="sxs-lookup"><span data-stu-id="7fa42-126">A range is limited to **five million cells**.</span></span>

<span data-ttu-id="7fa42-127">Wenn beim Umgang mit großen Datasets Fehler auftreten, verwenden Sie mehrere kleinere Bereiche anstelle größerer Bereiche.</span><span class="sxs-lookup"><span data-stu-id="7fa42-127">If you're encountering errors when dealing with large datasets, try using multiple smaller ranges instead of larger ranges.</span></span> <span data-ttu-id="7fa42-128">Sie können auch APIs wie [Range. getSpecialCells](/javascript/api/office-scripts/excelscript/excelscript.range#getspecialcells-celltype--cellvaluetype-) , um bestimmte Zellen anstelle von großen Bereichen Ziel.</span><span class="sxs-lookup"><span data-stu-id="7fa42-128">You can also APIs like [Range.getSpecialCells](/javascript/api/office-scripts/excelscript/excelscript.range#getspecialcells-celltype--cellvaluetype-) to target specific cells instead of large ranges.</span></span>

### <a name="power-automate"></a><span data-ttu-id="7fa42-129">Power Automate</span><span class="sxs-lookup"><span data-stu-id="7fa42-129">Power Automate</span></span>

<span data-ttu-id="7fa42-130">Bei der Verwendung von Office-Skripts mit Power Automation sind Sie auf **200 Anrufe pro Tag** limitiert.</span><span class="sxs-lookup"><span data-stu-id="7fa42-130">When using Office Scripts with Power Automate, you're limited to **200 calls per day**.</span></span> <span data-ttu-id="7fa42-131">Dieses Limit wird auf 12:00 Uhr UTC zurückgesetzt.</span><span class="sxs-lookup"><span data-stu-id="7fa42-131">This limit resets at 12:00 AM UTC.</span></span>

<span data-ttu-id="7fa42-132">Die Power Automation-Plattform hat auch Nutzungseinschränkungen, die im Artikel [Limits and Configuration in Power Automation](/power-automate/limits-and-config)zu finden sind.</span><span class="sxs-lookup"><span data-stu-id="7fa42-132">The Power Automate platform also has usage limitations, which can be found in the article [Limits and configuration in Power Automate](/power-automate/limits-and-config).</span></span>

## <a name="see-also"></a><span data-ttu-id="7fa42-133">Siehe auch</span><span class="sxs-lookup"><span data-stu-id="7fa42-133">See also</span></span>

- [<span data-ttu-id="7fa42-134">Behandeln von Problemen mit Office-Skripts</span><span class="sxs-lookup"><span data-stu-id="7fa42-134">Troubleshooting Office Scripts</span></span>](troubleshooting.md)
- [<span data-ttu-id="7fa42-135">Rückgängigmachen der Auswirkungen eines Office-Skripts</span><span class="sxs-lookup"><span data-stu-id="7fa42-135">Undo the effects of an Office Script</span></span>](undo.md)
- [<span data-ttu-id="7fa42-136">Verbessern der Leistung Ihrer Office-Skripts</span><span class="sxs-lookup"><span data-stu-id="7fa42-136">Improve the performance of your Office Scripts</span></span>](../develop/web-client-performance.md)
- [<span data-ttu-id="7fa42-137">Grundlegendes zur Skripterstellung für Office-Skripts in Excel im Internet</span><span class="sxs-lookup"><span data-stu-id="7fa42-137">Scripting Fundamentals for Office Scripts in Excel on the web</span></span>](../develop/scripting-fundamentals.md)
