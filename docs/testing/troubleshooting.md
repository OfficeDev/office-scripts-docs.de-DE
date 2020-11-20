---
title: Behandeln von Problemen mit Office-Skripts
description: Tipps und Techniken zum Debuggen von Office-Skripts sowie Hilferessourcen.
ms.date: 10/30/2020
localization_priority: Normal
ms.openlocfilehash: b45957bd336edce527397253cacec8cb09df715a
ms.sourcegitcommit: 82d3c0ef1e187bcdeceb2b5fc3411186674fe150
ms.translationtype: MT
ms.contentlocale: de-DE
ms.lasthandoff: 11/18/2020
ms.locfileid: "49342878"
---
# <a name="troubleshooting-office-scripts"></a><span data-ttu-id="19c0c-103">Behandeln von Problemen mit Office-Skripts</span><span class="sxs-lookup"><span data-stu-id="19c0c-103">Troubleshooting Office Scripts</span></span>

<span data-ttu-id="19c0c-104">Wenn Sie Office-Skripts entwickeln, können Sie Fehler machen.</span><span class="sxs-lookup"><span data-stu-id="19c0c-104">As you develop Office Scripts, you may make mistakes.</span></span> <span data-ttu-id="19c0c-105">Das ist okay.</span><span class="sxs-lookup"><span data-stu-id="19c0c-105">It's okay.</span></span> <span data-ttu-id="19c0c-106">Wir verfügen über Tools, mit denen Sie die Probleme finden und Ihre Skripts perfekt funktionieren lassen.</span><span class="sxs-lookup"><span data-stu-id="19c0c-106">We have tools that help find the problems and get your scripts working perfectly.</span></span>

## <a name="console-logs"></a><span data-ttu-id="19c0c-107">Konsolen Protokolle</span><span class="sxs-lookup"><span data-stu-id="19c0c-107">Console logs</span></span>

<span data-ttu-id="19c0c-108">Manchmal möchten Sie bei der Problembehandlung Nachrichten auf dem Bildschirm drucken.</span><span class="sxs-lookup"><span data-stu-id="19c0c-108">Sometimes while troubleshooting, you'll want to print messages to the screen.</span></span> <span data-ttu-id="19c0c-109">Diese können Ihnen den aktuellen Wert der Variablen oder die Code Pfade zeigen, die ausgelöst werden.</span><span class="sxs-lookup"><span data-stu-id="19c0c-109">These can show you the current value of variables or which code paths are being triggered.</span></span> <span data-ttu-id="19c0c-110">Protokollieren Sie dazu Text in der Konsole.</span><span class="sxs-lookup"><span data-stu-id="19c0c-110">To do this, log text to the console.</span></span>

```TypeScript
console.log("Logging myRange's address.");
console.log(myRange.getAddress());
```

<span data-ttu-id="19c0c-111">Zeichenfolgen, die an übergeben werden, werden `console.log` in der Protokollierungs Konsole des Code-Editors angezeigt.</span><span class="sxs-lookup"><span data-stu-id="19c0c-111">Strings passed to `console.log` will be displayed in the Code Editor's logging console.</span></span> <span data-ttu-id="19c0c-112">Klicken Sie zum Aktivieren der Konsole auf die Schaltfläche **Ellipsen** , und wählen Sie **Protokolle...**</span><span class="sxs-lookup"><span data-stu-id="19c0c-112">To turn on the console, press the **Ellipses** button and select **Logs...**</span></span>

<span data-ttu-id="19c0c-113">Protokolle wirken sich nicht auf die Arbeitsmappe aus.</span><span class="sxs-lookup"><span data-stu-id="19c0c-113">Logs do not affect the workbook.</span></span>

## <a name="error-messages"></a><span data-ttu-id="19c0c-114">Fehlermeldungen</span><span class="sxs-lookup"><span data-stu-id="19c0c-114">Error messages</span></span>

<span data-ttu-id="19c0c-115">Wenn Ihr Excel-Skript auf ein Problem stößt, wird ein Fehler ausgegeben.</span><span class="sxs-lookup"><span data-stu-id="19c0c-115">When your Excel Script encounters a problem running, it produces an error.</span></span> <span data-ttu-id="19c0c-116">Es wird eine Aufforderung angezeigt, in der Sie gefragt werden, ob Sie **Protokolle anzeigen** möchten.</span><span class="sxs-lookup"><span data-stu-id="19c0c-116">You'll see a prompt pop-up asking if you want to **View Logs**.</span></span> <span data-ttu-id="19c0c-117">Drücken Sie die-Taste, um die Konsole zu öffnen und Fehler anzuzeigen.</span><span class="sxs-lookup"><span data-stu-id="19c0c-117">Press that button to open the console and display any errors.</span></span>

## <a name="automate-tab-not-appearing-or-office-scripts-unavailable"></a><span data-ttu-id="19c0c-118">Die Registerkarte "automatisieren" wird nicht angezeigt oder Office-Skripts sind nicht verfügbar</span><span class="sxs-lookup"><span data-stu-id="19c0c-118">Automate tab not appearing or Office Scripts unavailable</span></span>

<span data-ttu-id="19c0c-119">Die folgenden Schritte sollten bei der Behandlung von Problemen im Zusammenhang mit der **automatischen** Registerkarte, die nicht in Excel im Internet angezeigt wird, helfen.</span><span class="sxs-lookup"><span data-stu-id="19c0c-119">The following steps should help troubleshoot any problems related to the **Automate** tab not appearing in Excel on the web.</span></span>

1. <span data-ttu-id="19c0c-120">[Stellen Sie sicher, dass Ihre Microsoft 365-Lizenz Office-Skripts enthält](../overview/excel.md#requirements).</span><span class="sxs-lookup"><span data-stu-id="19c0c-120">[Make sure your Microsoft 365 license includes Office Scripts](../overview/excel.md#requirements).</span></span>
1. <span data-ttu-id="19c0c-121">Stellen [Sie sicher, dass Ihr Browser unterstützt wird](platform-limits.md#browser-support).</span><span class="sxs-lookup"><span data-stu-id="19c0c-121">[Check that your browser is supported](platform-limits.md#browser-support).</span></span>
1. <span data-ttu-id="19c0c-122">[Stellen Sie sicher, dass Cookies von Drittanbietern aktiviert sind](platform-limits.md#third-party-cookies).</span><span class="sxs-lookup"><span data-stu-id="19c0c-122">[Ensure third-party cookies are enabled](platform-limits.md#third-party-cookies).</span></span>
1. <span data-ttu-id="19c0c-123">[Stellen Sie sicher, dass Office-Skripts im Microsoft 365 Admin Center von Ihrem Administrator nicht deaktiviert wurden](/microsoft-365/admin/manage/manage-office-scripts-settings).</span><span class="sxs-lookup"><span data-stu-id="19c0c-123">[Ensure that your admin has not disabled Office Scripts in the Microsoft 365 admin center](/microsoft-365/admin/manage/manage-office-scripts-settings).</span></span>

[!INCLUDE [Teams support note](../includes/teams-support-note.md)]

## <a name="help-resources"></a><span data-ttu-id="19c0c-124">Hilferessourcen</span><span class="sxs-lookup"><span data-stu-id="19c0c-124">Help resources</span></span>

<span data-ttu-id="19c0c-125">[Stack Overflow](https://stackoverflow.com/questions/tagged/office-scripts) ist eine Community von Entwicklern, die bereit sind, bei Codierungs Problemen zu helfen.</span><span class="sxs-lookup"><span data-stu-id="19c0c-125">[Stack Overflow](https://stackoverflow.com/questions/tagged/office-scripts) is a community of developers willing to help with coding problems.</span></span> <span data-ttu-id="19c0c-126">Häufig können Sie die Lösung für Ihr Problem in einer schnell Stapel-Überlauf Suche finden.</span><span class="sxs-lookup"><span data-stu-id="19c0c-126">Often, you'll be able to find the solution to your problem through a quick Stack Overflow search.</span></span> <span data-ttu-id="19c0c-127">Wenn nicht, stellen Sie Ihre Frage und markieren Sie Sie mit dem Tag "Office-Scripts".</span><span class="sxs-lookup"><span data-stu-id="19c0c-127">If not, ask your question and tag it with the "office-scripts" tag.</span></span> <span data-ttu-id="19c0c-128">Vergessen Sie nicht, dass Sie ein Office- *Skript* erstellen, kein Office *-Add-in*.</span><span class="sxs-lookup"><span data-stu-id="19c0c-128">Be sure to mention you're creating an Office *Script*, not an Office *Add-in*.</span></span>

<span data-ttu-id="19c0c-129">Wenn ein Problem mit der Office-JavaScript-API auftritt, erstellen Sie ein Problem im [officedev/Office-js-GitHub-](https://github.com/OfficeDev/office-js) Repository.</span><span class="sxs-lookup"><span data-stu-id="19c0c-129">If you encounter a problem with the Office JavaScript API, create an issue in the [OfficeDev/office-js](https://github.com/OfficeDev/office-js) GitHub repository.</span></span> <span data-ttu-id="19c0c-130">Mitglieder des Produktteams werden auf Probleme reagieren und weitere Unterstützung bereitstellen.</span><span class="sxs-lookup"><span data-stu-id="19c0c-130">Members of the product team will respond to issues and provide further assistance.</span></span> <span data-ttu-id="19c0c-131">Das Erstellen eines Problems im Repository **officedev/Office-js** zeigt, dass Sie einen Fehler in der Office-JavaScript-API-Bibliothek gefunden haben, die das Produktteam adressieren sollte.</span><span class="sxs-lookup"><span data-stu-id="19c0c-131">Creating an issue in the **OfficeDev/office-js** repository indicates you have found a flaw in the Office JavaScript API library that the product team should address.</span></span>

<span data-ttu-id="19c0c-132">Wenn ein Problem mit dem Aktions Recorder oder-Editor auftritt, senden Sie Feedback über die Schaltfläche **Hilfe > Feedback** in Excel.</span><span class="sxs-lookup"><span data-stu-id="19c0c-132">If there is a problem with the Action Recorder or Editor, send feedback through the **Help > Feedback** button in Excel.</span></span>

## <a name="see-also"></a><span data-ttu-id="19c0c-133">Siehe auch</span><span class="sxs-lookup"><span data-stu-id="19c0c-133">See also</span></span>

- [<span data-ttu-id="19c0c-134">Office-Skripts in Excel im Web</span><span class="sxs-lookup"><span data-stu-id="19c0c-134">Office Scripts in Excel on the web</span></span>](../overview/excel.md)
- [<span data-ttu-id="19c0c-135">Grundlegendes zur Skripterstellung für Office-Skripts in Excel im Internet</span><span class="sxs-lookup"><span data-stu-id="19c0c-135">Scripting Fundamentals for Office Scripts in Excel on the web</span></span>](../develop/scripting-fundamentals.md)
- [<span data-ttu-id="19c0c-136">Plattformbeschränkungen mit Office-Skripts</span><span class="sxs-lookup"><span data-stu-id="19c0c-136">Platform Limits with Office Scripts</span></span>](platform-limits.md)
- [<span data-ttu-id="19c0c-137">Verbessern der Leistung Ihrer Office-Skripts</span><span class="sxs-lookup"><span data-stu-id="19c0c-137">Improve the performance of your Office Scripts</span></span>](../develop/web-client-performance.md)
- [<span data-ttu-id="19c0c-138">Rückgängigmachen der Auswirkungen eines Office-Skripts</span><span class="sxs-lookup"><span data-stu-id="19c0c-138">Undo the effects of an Office Script</span></span>](undo.md)
