---
title: Behandeln von Problemen mit Office-Skripts
description: Tipps und Techniken zum Debuggen von Office-Skripts sowie Hilferessourcen.
ms.date: 05/14/2020
localization_priority: Normal
ms.openlocfilehash: 6448980eec45214a589444229db0fd781b9fea13
ms.sourcegitcommit: aec3c971c6640429f89b6bb99d2c95ea06725599
ms.translationtype: MT
ms.contentlocale: de-DE
ms.lasthandoff: 06/25/2020
ms.locfileid: "44878619"
---
# <a name="troubleshooting-office-scripts"></a><span data-ttu-id="3a940-103">Behandeln von Problemen mit Office-Skripts</span><span class="sxs-lookup"><span data-stu-id="3a940-103">Troubleshooting Office Scripts</span></span>

<span data-ttu-id="3a940-104">Wenn Sie Office-Skripts entwickeln, können Sie Fehler machen.</span><span class="sxs-lookup"><span data-stu-id="3a940-104">As you develop Office Scripts, you may make mistakes.</span></span> <span data-ttu-id="3a940-105">Das ist okay.</span><span class="sxs-lookup"><span data-stu-id="3a940-105">It's okay.</span></span> <span data-ttu-id="3a940-106">Wir verfügen über Tools, mit denen Sie die Probleme finden und Ihre Skripts perfekt funktionieren lassen.</span><span class="sxs-lookup"><span data-stu-id="3a940-106">We have tools that help find the problems and get your scripts working perfectly.</span></span>

## <a name="console-logs"></a><span data-ttu-id="3a940-107">Konsolen Protokolle</span><span class="sxs-lookup"><span data-stu-id="3a940-107">Console logs</span></span>

<span data-ttu-id="3a940-108">Manchmal möchten Sie bei der Problembehandlung Nachrichten auf dem Bildschirm drucken.</span><span class="sxs-lookup"><span data-stu-id="3a940-108">Sometimes while troubleshooting, you'll want to print messages to the screen.</span></span> <span data-ttu-id="3a940-109">Diese können Ihnen den aktuellen Wert der Variablen oder die Code Pfade zeigen, die ausgelöst werden.</span><span class="sxs-lookup"><span data-stu-id="3a940-109">These can show you the current value of variables or which code paths are being triggered.</span></span> <span data-ttu-id="3a940-110">Protokollieren Sie dazu Text in der Konsole.</span><span class="sxs-lookup"><span data-stu-id="3a940-110">To do this, log text to the console.</span></span>

```TypeScript
console.log("Logging myRange's address.");
console.log(myRange.getAddress());
```

<span data-ttu-id="3a940-111">Zeichenfolgen, die an übergeben werden, werden `console.log` in der Protokollierungs Konsole des Code-Editors angezeigt.</span><span class="sxs-lookup"><span data-stu-id="3a940-111">Strings passed to`console.log` will be displayed in the Code Editor's logging console.</span></span> <span data-ttu-id="3a940-112">Klicken Sie zum Aktivieren der Konsole auf die Schaltfläche **Ellipsen** , und wählen Sie **Protokolle...**</span><span class="sxs-lookup"><span data-stu-id="3a940-112">To turn on the console, press the **Ellipses** button and select **Logs...**</span></span>

<span data-ttu-id="3a940-113">Protokolle wirken sich nicht auf die Arbeitsmappe aus.</span><span class="sxs-lookup"><span data-stu-id="3a940-113">Logs do not affect the workbook.</span></span>

## <a name="error-messages"></a><span data-ttu-id="3a940-114">Fehlermeldungen</span><span class="sxs-lookup"><span data-stu-id="3a940-114">Error messages</span></span>

<span data-ttu-id="3a940-115">Wenn Ihr Excel-Skript auf ein Problem stößt, wird ein Fehler ausgegeben.</span><span class="sxs-lookup"><span data-stu-id="3a940-115">When your Excel Script encounters a problem running, it produces an error.</span></span> <span data-ttu-id="3a940-116">Es wird eine Aufforderung angezeigt, in der Sie gefragt werden, ob Sie **Protokolle anzeigen**möchten.</span><span class="sxs-lookup"><span data-stu-id="3a940-116">You'll see a prompt pop-up asking if you want to **View Logs**.</span></span> <span data-ttu-id="3a940-117">Drücken Sie die-Taste, um die Konsole zu öffnen und Fehler anzuzeigen.</span><span class="sxs-lookup"><span data-stu-id="3a940-117">Press that button to open the console and display any errors.</span></span>

## <a name="help-resources"></a><span data-ttu-id="3a940-118">Hilferessourcen</span><span class="sxs-lookup"><span data-stu-id="3a940-118">Help resources</span></span>

<span data-ttu-id="3a940-119">[Stack Overflow](https://stackoverflow.com/questions/tagged/office-scripts) ist eine Community von Entwicklern, die bereit sind, bei Codierungs Problemen zu helfen.</span><span class="sxs-lookup"><span data-stu-id="3a940-119">[Stack Overflow](https://stackoverflow.com/questions/tagged/office-scripts) is a community of developers willing to help with coding problems.</span></span> <span data-ttu-id="3a940-120">Häufig können Sie die Lösung für Ihr Problem in einer schnell Stapel-Überlauf Suche finden.</span><span class="sxs-lookup"><span data-stu-id="3a940-120">Often, you'll be able to find the solution to your problem through a quick Stack Overflow search.</span></span> <span data-ttu-id="3a940-121">Wenn nicht, stellen Sie Ihre Frage und markieren Sie Sie mit dem Tag "Office-Scripts".</span><span class="sxs-lookup"><span data-stu-id="3a940-121">If not, ask your question and tag it with the "office-scripts" tag.</span></span> <span data-ttu-id="3a940-122">Vergessen Sie nicht, dass Sie ein Office- *Skript*erstellen, kein Office *-Add-in*.</span><span class="sxs-lookup"><span data-stu-id="3a940-122">Be sure to mention you're creating an Office *Script*, not an Office *Add-in*.</span></span>

<span data-ttu-id="3a940-123">Wenn ein Problem mit der Office-JavaScript-API auftritt, erstellen Sie ein Problem im [officedev/Office-js-GitHub-](https://github.com/OfficeDev/office-js) Repository.</span><span class="sxs-lookup"><span data-stu-id="3a940-123">If you encounter a problem with the Office JavaScript API, create an issue in the [OfficeDev/office-js](https://github.com/OfficeDev/office-js) GitHub repository.</span></span> <span data-ttu-id="3a940-124">Mitglieder des Produktteams werden auf Probleme reagieren und weitere Unterstützung bereitstellen.</span><span class="sxs-lookup"><span data-stu-id="3a940-124">Members of the product team will respond to issues and provide further assistance.</span></span> <span data-ttu-id="3a940-125">Das Erstellen eines Problems im Repository **officedev/Office-js** zeigt, dass Sie einen Fehler in der Office-JavaScript-API-Bibliothek gefunden haben, die das Produktteam adressieren sollte.</span><span class="sxs-lookup"><span data-stu-id="3a940-125">Creating an issue in the **OfficeDev/office-js** repository indicates you have found a flaw in the Office JavaScript API library that the product team should address.</span></span>

<span data-ttu-id="3a940-126">Wenn ein Problem mit dem Aktions Recorder oder-Editor auftritt, senden Sie Feedback über die Schaltfläche **Hilfe > Feedback** in Excel.</span><span class="sxs-lookup"><span data-stu-id="3a940-126">If there is a problem with the Action Recorder or Editor, send feedback through the **Help > Feedback** button in Excel.</span></span>

## <a name="see-also"></a><span data-ttu-id="3a940-127">Siehe auch</span><span class="sxs-lookup"><span data-stu-id="3a940-127">See also</span></span>

- [<span data-ttu-id="3a940-128">Office-Skripts in Excel im Web</span><span class="sxs-lookup"><span data-stu-id="3a940-128">Office Scripts in Excel on the web</span></span>](../overview/excel.md)
- [<span data-ttu-id="3a940-129">Grundlegendes zur Skripterstellung für Office-Skripts in Excel im Internet</span><span class="sxs-lookup"><span data-stu-id="3a940-129">Scripting Fundamentals for Office Scripts in Excel on the web</span></span>](../develop/scripting-fundamentals.md)
- [<span data-ttu-id="3a940-130">Rückgängigmachen der Auswirkungen eines Office-Skripts</span><span class="sxs-lookup"><span data-stu-id="3a940-130">Undo the effects of an Office Script</span></span>](undo.md)
- [<span data-ttu-id="3a940-131">Verbessern der Leistung Ihrer Office-Skripts</span><span class="sxs-lookup"><span data-stu-id="3a940-131">Improve the performance of your Office Scripts</span></span>](../develop/web-client-performance.md)
