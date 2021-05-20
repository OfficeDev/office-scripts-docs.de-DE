---
title: Plattformlimits und -anforderungen mit Office Scripts
description: Ressourcenlimits und Browserunterstützung für Office Skripts bei Verwendung mit Excel im Web
ms.date: 05/17/2021
localization_priority: Normal
ms.openlocfilehash: 7e81aaf2f96faeb67c815814fe3b7f1795651318
ms.sourcegitcommit: 4687693f02fc90a57ba30c461f35046e02e6f5fb
ms.translationtype: MT
ms.contentlocale: de-DE
ms.lasthandoff: 05/19/2021
ms.locfileid: "52545581"
---
# <a name="platform-limits-and-requirements-with-office-scripts"></a><span data-ttu-id="10996-103">Plattformlimits und -anforderungen mit Office Scripts</span><span class="sxs-lookup"><span data-stu-id="10996-103">Platform limits and requirements with Office Scripts</span></span>

<span data-ttu-id="10996-104">Es gibt einige Plattformeinschränkungen, die Sie bei der Entwicklung Office Skripts beachten sollten.</span><span class="sxs-lookup"><span data-stu-id="10996-104">There are some platform limitations of which you should be aware when developing Office Scripts.</span></span> <span data-ttu-id="10996-105">In diesem Artikel werden die Browserunterstützung und die Datenlimits für Office Skripts für Excel im Web beschrieben.</span><span class="sxs-lookup"><span data-stu-id="10996-105">This article details the browser support and data limits for Office Scripts for Excel on the web.</span></span>

## <a name="browser-support"></a><span data-ttu-id="10996-106">Browserunterstützung</span><span class="sxs-lookup"><span data-stu-id="10996-106">Browser support</span></span>

<span data-ttu-id="10996-107">Office Skripts funktionieren in jedem Browser, der [Office für das Web unterstützt.](https://support.microsoft.com/office/ad1303e0-a318-47aa-b409-d3a5eb44e452)</span><span class="sxs-lookup"><span data-stu-id="10996-107">Office Scripts work in any browser that [supports Office for the web](https://support.microsoft.com/office/ad1303e0-a318-47aa-b409-d3a5eb44e452).</span></span> <span data-ttu-id="10996-108">Einige JavaScript-Funktionen werden in Internet Explorer 11 (IE 11) jedoch nicht unterstützt.</span><span class="sxs-lookup"><span data-stu-id="10996-108">However, some JavaScript features aren't supported in Internet Explorer 11 (IE 11).</span></span> <span data-ttu-id="10996-109">Alle Funktionen, die in [ES6 oder höher](https://www.w3schools.com/Js/js_es6.asp) eingeführt werden, funktionieren nicht mit IE 11.</span><span class="sxs-lookup"><span data-stu-id="10996-109">Any features introduced in [ES6 or later](https://www.w3schools.com/Js/js_es6.asp) won't work with IE 11.</span></span> <span data-ttu-id="10996-110">Wenn Personen in Ihrer Organisation diesen Browser weiterhin verwenden, sollten Sie Ihre Skripts in dieser Umgebung testen, wenn Sie sie freigeben.</span><span class="sxs-lookup"><span data-stu-id="10996-110">If people in your organization still use that browser, be sure to test your scripts in that environment when sharing them.</span></span>

[!INCLUDE [Teams support note](../includes/teams-support-note.md)]

### <a name="third-party-cookies"></a><span data-ttu-id="10996-111">Cookies von Drittanbietern</span><span class="sxs-lookup"><span data-stu-id="10996-111">Third-party cookies</span></span>

<span data-ttu-id="10996-112">Ihr Browser benötigt Cookies von Drittanbietern, die aktiviert sind, um die Registerkarte **Automatisieren** in Excel im Web anzuzeigen.</span><span class="sxs-lookup"><span data-stu-id="10996-112">Your browser needs third-party cookies enabled to show the **Automate** tab in Excel on the web.</span></span> <span data-ttu-id="10996-113">Überprüfen Sie die Browsereinstellungen, wenn die Registerkarte nicht angezeigt wird.</span><span class="sxs-lookup"><span data-stu-id="10996-113">Check your browser settings if the tab isn't being displayed.</span></span> <span data-ttu-id="10996-114">Wenn Sie eine private Browsersitzung verwenden, müssen Sie diese Einstellung möglicherweise jedes Mal erneut aktivieren.</span><span class="sxs-lookup"><span data-stu-id="10996-114">If you're using a private browser session, you may need to re-enable this setting each time.</span></span>

> [!NOTE]
> <span data-ttu-id="10996-115">Einige Browser bezeichnen diese Einstellung als "alle Cookies" anstelle von "Drittanbieter-Cookies".</span><span class="sxs-lookup"><span data-stu-id="10996-115">Some browsers refer to this setting as "all cookies", instead of "third-party cookies".</span></span>

#### <a name="instructions-for-adjusting-cookie-settings-in-popular-browsers"></a><span data-ttu-id="10996-116">Anweisungen zum Anpassen der Cookie-Einstellungen in gängigen Browsern</span><span class="sxs-lookup"><span data-stu-id="10996-116">Instructions for adjusting cookie settings in popular browsers</span></span>

- [<span data-ttu-id="10996-117">Chrome</span><span class="sxs-lookup"><span data-stu-id="10996-117">Chrome</span></span>](https://support.google.com/chrome/answer/95647)
- [<span data-ttu-id="10996-118">Microsoft Edge</span><span class="sxs-lookup"><span data-stu-id="10996-118">Edge</span></span>](https://support.microsoft.com/microsoft-edge/temporarily-allow-cookies-and-site-data-in-microsoft-edge-597f04f2-c0ce-f08c-7c2b-541086362bd2)
- [<span data-ttu-id="10996-119">Firefox</span><span class="sxs-lookup"><span data-stu-id="10996-119">Firefox</span></span>](https://support.mozilla.org/kb/disable-third-party-cookies)
- [<span data-ttu-id="10996-120">Safari</span><span class="sxs-lookup"><span data-stu-id="10996-120">Safari</span></span>](https://support.apple.com/guide/safari/manage-cookies-and-website-data-sfri11471/mac)

## <a name="data-limits"></a><span data-ttu-id="10996-121">Beschränkungen für Daten</span><span class="sxs-lookup"><span data-stu-id="10996-121">Data limits</span></span>

<span data-ttu-id="10996-122">Es gibt Grenzen, wie viele Excel Daten auf einmal übertragen werden können und wie viele einzelne Power Automate Transaktionen durchgeführt werden können.</span><span class="sxs-lookup"><span data-stu-id="10996-122">There are limits on how much Excel data can be transferred at once and how many individual Power Automate transactions can be conducted.</span></span>

### <a name="excel"></a><span data-ttu-id="10996-123">Excel</span><span class="sxs-lookup"><span data-stu-id="10996-123">Excel</span></span>

<span data-ttu-id="10996-124">Excel für das Web hat die folgenden Einschränkungen beim Aufrufen der Arbeitsmappe über ein Skript:</span><span class="sxs-lookup"><span data-stu-id="10996-124">Excel for the web has the following limitations when making calls to the workbook through a script:</span></span>

- <span data-ttu-id="10996-125">Anfragen und Antworten sind auf **5 MB** begrenzt.</span><span class="sxs-lookup"><span data-stu-id="10996-125">Requests and responses are limited to **5MB**.</span></span>
- <span data-ttu-id="10996-126">Ein Bereich ist auf **fünf Millionen Zellen** begrenzt.</span><span class="sxs-lookup"><span data-stu-id="10996-126">A range is limited to **five million cells**.</span></span>

<span data-ttu-id="10996-127">Wenn beim Umgang mit großen Datasets Fehler auftreten, verwenden Sie mehrere kleinere Bereiche anstelle größerer Bereiche.</span><span class="sxs-lookup"><span data-stu-id="10996-127">If you're encountering errors when dealing with large datasets, try using multiple smaller ranges instead of larger ranges.</span></span> <span data-ttu-id="10996-128">Ein Beispiel finden Sie unter Schreiben eines großen Dataset-Beispiels. [](../resources/samples/write-large-dataset.md)</span><span class="sxs-lookup"><span data-stu-id="10996-128">For an example, see the [Write a large dataset](../resources/samples/write-large-dataset.md) sample.</span></span> <span data-ttu-id="10996-129">Sie können auch APIs wie [Range.getSpecialCells](/javascript/api/office-scripts/excelscript/excelscript.range#getspecialcells-celltype--cellvaluetype-) verwenden, um bestimmte Zellen anstelle großer Bereiche zu zielen.</span><span class="sxs-lookup"><span data-stu-id="10996-129">You can also use APIs like [Range.getSpecialCells](/javascript/api/office-scripts/excelscript/excelscript.range#getspecialcells-celltype--cellvaluetype-) to target specific cells instead of large ranges.</span></span>

### <a name="power-automate"></a><span data-ttu-id="10996-130">Power Automate</span><span class="sxs-lookup"><span data-stu-id="10996-130">Power Automate</span></span>

<span data-ttu-id="10996-131">Bei Verwendung Office Skripts mit Power Automate ist jeder Benutzer auf **400 Aufrufe der Aktion Skript ausführen pro Tag** beschränkt.</span><span class="sxs-lookup"><span data-stu-id="10996-131">When using Office Scripts with Power Automate, each user is limited to **400 calls to the Run Script action per day**.</span></span> <span data-ttu-id="10996-132">Dieses Limit wird um 12:00 Uhr UTC zurückgesetzt.</span><span class="sxs-lookup"><span data-stu-id="10996-132">This limit resets at 12:00 AM UTC.</span></span>

<span data-ttu-id="10996-133">Die Power Automate Plattform hat auch Nutzungsbeschränkungen, die in den folgenden Artikeln zu finden sind:</span><span class="sxs-lookup"><span data-stu-id="10996-133">The Power Automate platform also has usage limitations, which can be found in the following articles:</span></span>

- [<span data-ttu-id="10996-134">Grenzwerte und Konfiguration in Power Automate</span><span class="sxs-lookup"><span data-stu-id="10996-134">Limits and configuration in Power Automate</span></span>](/power-automate/limits-and-config)
- [<span data-ttu-id="10996-135">Bekannte Probleme und Einschränkungen für den Excel Online(Business)-Connector</span><span class="sxs-lookup"><span data-stu-id="10996-135">Known issues and limitations for the Excel Online (Business) connector</span></span>](/connectors/excelonlinebusiness/#known-issues-and-limitations)

## <a name="see-also"></a><span data-ttu-id="10996-136">Siehe auch</span><span class="sxs-lookup"><span data-stu-id="10996-136">See also</span></span>

- [<span data-ttu-id="10996-137">Fehlerbehebung Office Skripts</span><span class="sxs-lookup"><span data-stu-id="10996-137">Troubleshoot Office Scripts</span></span>](troubleshooting.md)
- [<span data-ttu-id="10996-138">Auswirkungen von Office-Skripts rückgängig machen</span><span class="sxs-lookup"><span data-stu-id="10996-138">Undo the effects of Office Scripts</span></span>](undo.md)
- [<span data-ttu-id="10996-139">Verbessern Sie die Leistung Ihrer Office Scripts</span><span class="sxs-lookup"><span data-stu-id="10996-139">Improve the performance of your Office Scripts</span></span>](../develop/web-client-performance.md)
- [<span data-ttu-id="10996-140">Skriptgrundlagen für Office Skripts in Excel im Web</span><span class="sxs-lookup"><span data-stu-id="10996-140">Scripting Fundamentals for Office Scripts in Excel on the web</span></span>](../develop/scripting-fundamentals.md)
