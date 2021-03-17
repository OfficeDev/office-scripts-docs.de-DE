---
title: Plattformbeschränkungen und -anforderungen mit Office-Skripts
description: Ressourcenbeschränkungen und Browserunterstützung für Office-Skripts bei Verwendung mit Excel im Web
ms.date: 03/12/2021
localization_priority: Normal
ms.openlocfilehash: 93307b6204f409f26c77b5ead33188205d5c4b4d
ms.sourcegitcommit: 5bde455b06ee2ed007f3e462d8ad485b257774ef
ms.translationtype: MT
ms.contentlocale: de-DE
ms.lasthandoff: 03/17/2021
ms.locfileid: "50837265"
---
# <a name="platform-limits-and-requirements-with-office-scripts"></a><span data-ttu-id="3e094-103">Plattformbeschränkungen und -anforderungen mit Office-Skripts</span><span class="sxs-lookup"><span data-stu-id="3e094-103">Platform limits and requirements with Office Scripts</span></span>

<span data-ttu-id="3e094-104">Es gibt einige Plattformeinschränkungen, die Sie bei der Entwicklung von Office-Skripts beachten sollten.</span><span class="sxs-lookup"><span data-stu-id="3e094-104">There are some platform limitations of which you should be aware when developing Office Scripts.</span></span> <span data-ttu-id="3e094-105">In diesem Artikel werden die Browserunterstützung und Datenbeschränkungen für Office-Skripts für Excel im Web erläutert.</span><span class="sxs-lookup"><span data-stu-id="3e094-105">This article details the browser support and data limits for Office Scripts for Excel on the web.</span></span>

## <a name="browser-support"></a><span data-ttu-id="3e094-106">Browserunterstützung</span><span class="sxs-lookup"><span data-stu-id="3e094-106">Browser support</span></span>

<span data-ttu-id="3e094-107">Office-Skripts funktionieren in jedem Browser, [der Office für das Web unterstützt.](https://support.microsoft.com/office/ad1303e0-a318-47aa-b409-d3a5eb44e452)</span><span class="sxs-lookup"><span data-stu-id="3e094-107">Office Scripts work in any browser that [supports Office for the web](https://support.microsoft.com/office/ad1303e0-a318-47aa-b409-d3a5eb44e452).</span></span> <span data-ttu-id="3e094-108">Einige JavaScript-Features werden jedoch in Internet Explorer 11 (IE 11) nicht unterstützt.</span><span class="sxs-lookup"><span data-stu-id="3e094-108">However, some JavaScript features aren't supported in Internet Explorer 11 (IE 11).</span></span> <span data-ttu-id="3e094-109">In ES6 oder [höher eingeführte](https://www.w3schools.com/Js/js_es6.asp) Features funktionieren nicht mit IE 11.</span><span class="sxs-lookup"><span data-stu-id="3e094-109">Any features introduced in [ES6 or later](https://www.w3schools.com/Js/js_es6.asp) won't work with IE 11.</span></span> <span data-ttu-id="3e094-110">Wenn Personen in Ihrer Organisation diesen Browser weiterhin verwenden, sollten Sie ihre Skripts in dieser Umgebung testen, wenn Sie sie freigeben.</span><span class="sxs-lookup"><span data-stu-id="3e094-110">If people in your organization still use that browser, be sure to test your scripts in that environment when sharing them.</span></span>

[!INCLUDE [Teams support note](../includes/teams-support-note.md)]

### <a name="third-party-cookies"></a><span data-ttu-id="3e094-111">Cookies von Drittanbietern</span><span class="sxs-lookup"><span data-stu-id="3e094-111">Third-party cookies</span></span>

<span data-ttu-id="3e094-112">Ihr Browser benötigt Cookies von Drittanbietern, die aktiviert sind, um die Registerkarte **Automatisieren** in Excel im Web anzuzeigen.</span><span class="sxs-lookup"><span data-stu-id="3e094-112">Your browser needs third-party cookies enabled to show the **Automate** tab in Excel on the web.</span></span> <span data-ttu-id="3e094-113">Überprüfen Sie die Browsereinstellungen, wenn die Registerkarte nicht angezeigt wird.</span><span class="sxs-lookup"><span data-stu-id="3e094-113">Check your browser settings if the tab isn't being displayed.</span></span> <span data-ttu-id="3e094-114">Wenn Sie eine private Browsersitzung verwenden, müssen Sie diese Einstellung möglicherweise jedes Mal erneut aktivieren.</span><span class="sxs-lookup"><span data-stu-id="3e094-114">If you're using a private browser session, you may need to re-enable this setting each time.</span></span>

> [!NOTE]
> <span data-ttu-id="3e094-115">In einigen Browsern wird diese Einstellung als "alle Cookies" anstelle von "Cookies von Drittanbietern" bezeichnet.</span><span class="sxs-lookup"><span data-stu-id="3e094-115">Some browsers refer to this setting as "all cookies", instead of "third-party cookies".</span></span>

#### <a name="instructions-for-adjusting-cookie-settings-in-popular-browsers"></a><span data-ttu-id="3e094-116">Anweisungen zum Anpassen von Cookieeinstellungen in beliebten Browsern</span><span class="sxs-lookup"><span data-stu-id="3e094-116">Instructions for adjusting cookie settings in popular browsers</span></span>

- [<span data-ttu-id="3e094-117">Chrome</span><span class="sxs-lookup"><span data-stu-id="3e094-117">Chrome</span></span>](https://support.google.com/chrome/answer/95647)
- [<span data-ttu-id="3e094-118">Microsoft Edge</span><span class="sxs-lookup"><span data-stu-id="3e094-118">Edge</span></span>](https://support.microsoft.com/microsoft-edge/temporarily-allow-cookies-and-site-data-in-microsoft-edge-597f04f2-c0ce-f08c-7c2b-541086362bd2)
- [<span data-ttu-id="3e094-119">Firefox</span><span class="sxs-lookup"><span data-stu-id="3e094-119">Firefox</span></span>](https://support.mozilla.org/kb/disable-third-party-cookies)
- [<span data-ttu-id="3e094-120">Safari</span><span class="sxs-lookup"><span data-stu-id="3e094-120">Safari</span></span>](https://support.apple.com/guide/safari/manage-cookies-and-website-data-sfri11471/mac)

## <a name="data-limits"></a><span data-ttu-id="3e094-121">Beschränkungen für Daten</span><span class="sxs-lookup"><span data-stu-id="3e094-121">Data limits</span></span>

<span data-ttu-id="3e094-122">Es gibt Beschränkungen, wie viele Excel-Daten gleichzeitig übertragen werden können und wie viele einzelne Power Automate-Transaktionen durchgeführt werden können.</span><span class="sxs-lookup"><span data-stu-id="3e094-122">There are limits on how much Excel data can be transferred at once and how many individual Power Automate transactions can be conducted.</span></span>

### <a name="excel"></a><span data-ttu-id="3e094-123">Excel</span><span class="sxs-lookup"><span data-stu-id="3e094-123">Excel</span></span>

<span data-ttu-id="3e094-124">Excel für das Web hat die folgenden Einschränkungen, wenn Aufrufe der Arbeitsmappe über ein Skript ausgeführt werden:</span><span class="sxs-lookup"><span data-stu-id="3e094-124">Excel for the web has the following limitations when making calls to the workbook through a script:</span></span>

- <span data-ttu-id="3e094-125">Anforderungen und Antworten sind auf **5 MB beschränkt.**</span><span class="sxs-lookup"><span data-stu-id="3e094-125">Requests and responses are limited to **5MB**.</span></span>
- <span data-ttu-id="3e094-126">Ein Bereich ist auf **fünf Millionen Zellen beschränkt.**</span><span class="sxs-lookup"><span data-stu-id="3e094-126">A range is limited to **five million cells**.</span></span>

<span data-ttu-id="3e094-127">Wenn beim Umgang mit großen Datasets Fehler auftreten, versuchen Sie, mehrere kleinere Bereiche anstelle größerer Bereiche zu verwenden.</span><span class="sxs-lookup"><span data-stu-id="3e094-127">If you're encountering errors when dealing with large datasets, try using multiple smaller ranges instead of larger ranges.</span></span> <span data-ttu-id="3e094-128">Sie können apIs wie [Range.getSpecialCells](/javascript/api/office-scripts/excelscript/excelscript.range#getspecialcells-celltype--cellvaluetype-) auch auf bestimmte Zellen statt auf große Bereiche zielen.</span><span class="sxs-lookup"><span data-stu-id="3e094-128">You can also APIs like [Range.getSpecialCells](/javascript/api/office-scripts/excelscript/excelscript.range#getspecialcells-celltype--cellvaluetype-) to target specific cells instead of large ranges.</span></span>

### <a name="power-automate"></a><span data-ttu-id="3e094-129">Power Automate</span><span class="sxs-lookup"><span data-stu-id="3e094-129">Power Automate</span></span>

<span data-ttu-id="3e094-130">Bei verwendung von Office Scripts mit Power Automate ist jeder Benutzer auf **200** Anrufe pro Tag beschränkt.</span><span class="sxs-lookup"><span data-stu-id="3e094-130">When using Office Scripts with Power Automate, each user is limited to **200 calls per day**.</span></span> <span data-ttu-id="3e094-131">Dieser Grenzwert wird um 12:00 Uhr UTC zurückgesetzt.</span><span class="sxs-lookup"><span data-stu-id="3e094-131">This limit resets at 12:00 AM UTC.</span></span>

<span data-ttu-id="3e094-132">Die Power Automate-Plattform hat auch Nutzungseinschränkungen, die in den folgenden Artikeln zu finden sind:</span><span class="sxs-lookup"><span data-stu-id="3e094-132">The Power Automate platform also has usage limitations, which can be found in the following articles:</span></span>

- [<span data-ttu-id="3e094-133">Grenzwerte und Konfiguration in Power Automate</span><span class="sxs-lookup"><span data-stu-id="3e094-133">Limits and configuration in Power Automate</span></span>](/power-automate/limits-and-config)
- [<span data-ttu-id="3e094-134">Bekannte Probleme und Einschränkungen für den Excel Online (Business)-Connector</span><span class="sxs-lookup"><span data-stu-id="3e094-134">Known issues and limitations for the Excel Online (Business) connector</span></span>](/connectors/excelonlinebusiness/#known-issues-and-limitations)

## <a name="see-also"></a><span data-ttu-id="3e094-135">Siehe auch</span><span class="sxs-lookup"><span data-stu-id="3e094-135">See also</span></span>

- [<span data-ttu-id="3e094-136">Behandeln von Problemen mit Office-Skripts</span><span class="sxs-lookup"><span data-stu-id="3e094-136">Troubleshooting Office Scripts</span></span>](troubleshooting.md)
- [<span data-ttu-id="3e094-137">Rückgängigmachen der Auswirkungen eines Office-Skripts</span><span class="sxs-lookup"><span data-stu-id="3e094-137">Undo the effects of an Office Script</span></span>](undo.md)
- [<span data-ttu-id="3e094-138">Verbessern der Leistung Ihrer Office-Skripts</span><span class="sxs-lookup"><span data-stu-id="3e094-138">Improve the performance of your Office Scripts</span></span>](../develop/web-client-performance.md)
- [<span data-ttu-id="3e094-139">Skripting-Grundlagen für Office-Skripts in Excel im Web</span><span class="sxs-lookup"><span data-stu-id="3e094-139">Scripting Fundamentals for Office Scripts in Excel on the web</span></span>](../develop/scripting-fundamentals.md)
