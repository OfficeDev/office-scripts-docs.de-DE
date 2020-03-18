---
title: Problembehandlung bei Office-Skripts
description: Tipps und Techniken zum Debuggen von Office-Skripts sowie Hilferessourcen.
ms.date: 12/13/2019
localization_priority: Normal
ms.openlocfilehash: 959faff875f342dc1b1ab158ad9ded24732b0894
ms.sourcegitcommit: b075eed5a6f275274fbbf6d62633219eac416f26
ms.translationtype: MT
ms.contentlocale: de-DE
ms.lasthandoff: 03/10/2020
ms.locfileid: "42700203"
---
# <a name="troubleshooting-office-scripts"></a><span data-ttu-id="102f4-103">Problembehandlung bei Office-Skripts</span><span class="sxs-lookup"><span data-stu-id="102f4-103">Troubleshooting Office Scripts</span></span>

<span data-ttu-id="102f4-104">Wenn Sie Office-Skripts entwickeln, können Sie Fehler machen.</span><span class="sxs-lookup"><span data-stu-id="102f4-104">As you develop Office Scripts, you may make mistakes.</span></span> <span data-ttu-id="102f4-105">Das ist okay.</span><span class="sxs-lookup"><span data-stu-id="102f4-105">It's okay.</span></span> <span data-ttu-id="102f4-106">Wir verfügen über Tools, mit denen Sie die Probleme finden und Ihre Skripts perfekt funktionieren lassen.</span><span class="sxs-lookup"><span data-stu-id="102f4-106">We have tools that help find the problems and get your scripts working perfectly.</span></span>

## <a name="console-logs"></a><span data-ttu-id="102f4-107">Konsolen Protokolle</span><span class="sxs-lookup"><span data-stu-id="102f4-107">Console logs</span></span>

<span data-ttu-id="102f4-108">Manchmal möchten Sie bei der Problembehandlung Nachrichten auf dem Bildschirm drucken.</span><span class="sxs-lookup"><span data-stu-id="102f4-108">Sometimes while troubleshooting, you'll want to print messages to the screen.</span></span> <span data-ttu-id="102f4-109">Diese können Ihnen den aktuellen Wert der Variablen oder die Code Pfade zeigen, die ausgelöst werden.</span><span class="sxs-lookup"><span data-stu-id="102f4-109">These can show you the current value of variables or which code paths are being triggered.</span></span> <span data-ttu-id="102f4-110">Protokollieren Sie dazu Text in der Konsole.</span><span class="sxs-lookup"><span data-stu-id="102f4-110">To do this, log text to the console.</span></span>

```TypeScript
console.log("Logging my range's address.");
myRange.load("address");
await context.sync();
console.log(myRange.address);
```

> [!IMPORTANT]
> <span data-ttu-id="102f4-111">Vergessen Sie nicht `load` , Arbeitsblatt `sync` Daten und mit der Arbeitsmappe vor dem Protokollieren von Objekteigenschaften.</span><span class="sxs-lookup"><span data-stu-id="102f4-111">Don't forget to `load` worksheet data and `sync` with the workbook before logging object properties.</span></span>

<span data-ttu-id="102f4-112">Zeichenfolgen,`console.log` die an übergeben werden, werden in der Protokollierungs Konsole des Code-Editors angezeigt.</span><span class="sxs-lookup"><span data-stu-id="102f4-112">Strings passed to`console.log` will be displayed in the Code Editor's logging console.</span></span> <span data-ttu-id="102f4-113">Klicken Sie zum Aktivieren der Konsole auf die Schaltfläche **Ellipsen** , und wählen Sie **Protokolle...**</span><span class="sxs-lookup"><span data-stu-id="102f4-113">To turn on the console, press the **Ellipses** button and select **Logs...**</span></span>

<span data-ttu-id="102f4-114">Protokolle wirken sich nicht auf die Arbeitsmappe aus.</span><span class="sxs-lookup"><span data-stu-id="102f4-114">Logs do not affect the workbook.</span></span>

## <a name="error-messages"></a><span data-ttu-id="102f4-115">Fehlermeldungen</span><span class="sxs-lookup"><span data-stu-id="102f4-115">Error messages</span></span>

<span data-ttu-id="102f4-116">Wenn Ihr Excel-Skript auf ein Problem stößt, wird ein Fehler ausgegeben.</span><span class="sxs-lookup"><span data-stu-id="102f4-116">When your Excel Script encounters a problem running, it produces an error.</span></span> <span data-ttu-id="102f4-117">Es wird eine Aufforderung angezeigt, in der Sie gefragt werden, ob Sie **Protokolle anzeigen**möchten.</span><span class="sxs-lookup"><span data-stu-id="102f4-117">You'll see a prompt pop-up asking if you want to **View Logs**.</span></span> <span data-ttu-id="102f4-118">Drücken Sie die-Taste, um die Konsole zu öffnen und Fehler anzuzeigen.</span><span class="sxs-lookup"><span data-stu-id="102f4-118">Press that button to open the console and display any errors.</span></span>

## <a name="help-resources"></a><span data-ttu-id="102f4-119">Hilferessourcen</span><span class="sxs-lookup"><span data-stu-id="102f4-119">Help resources</span></span>

<span data-ttu-id="102f4-120">[Stack Overflow](https://stackoverflow.com/questions/tagged/office-scripts) ist eine Community von Entwicklern, die bereit sind, bei Codierungs Problemen zu helfen.</span><span class="sxs-lookup"><span data-stu-id="102f4-120">[Stack Overflow](https://stackoverflow.com/questions/tagged/office-scripts) is a community of developers willing to help with coding problems.</span></span> <span data-ttu-id="102f4-121">Häufig können Sie die Lösung für Ihr Problem in einer schnell Stapel-Überlauf Suche finden.</span><span class="sxs-lookup"><span data-stu-id="102f4-121">Often, you'll be able to find the solution to your problem through a quick Stack Overflow search.</span></span> <span data-ttu-id="102f4-122">Wenn nicht, stellen Sie Ihre Frage und markieren Sie Sie mit dem Tag "Office-Scripts".</span><span class="sxs-lookup"><span data-stu-id="102f4-122">If not, ask your question and tag it with the "office-scripts" tag.</span></span> <span data-ttu-id="102f4-123">Vergessen Sie nicht, dass Sie ein Office- *Skript*erstellen, kein Office *-Add-in*.</span><span class="sxs-lookup"><span data-stu-id="102f4-123">Be sure to mention you're creating an Office *Script*, not an Office *Add-in*.</span></span>

<span data-ttu-id="102f4-124">Wenn ein Problem mit der Office-JavaScript-API auftritt, erstellen Sie ein Problem im [officedev/Office-js-GitHub-](https://github.com/OfficeDev/office-js) Repository.</span><span class="sxs-lookup"><span data-stu-id="102f4-124">If you encounter a problem with the Office JavaScript API, create an issue in the [OfficeDev/office-js](https://github.com/OfficeDev/office-js) GitHub repository.</span></span> <span data-ttu-id="102f4-125">Mitglieder des Produktteams werden auf Probleme reagieren und weitere Unterstützung bereitstellen.</span><span class="sxs-lookup"><span data-stu-id="102f4-125">Members of the product team will respond to issues and provide further assistance.</span></span> <span data-ttu-id="102f4-126">Das Erstellen eines Problems im Repository **officedev/Office-js** zeigt, dass Sie einen Fehler in der Office-JavaScript-API-Bibliothek gefunden haben, die das Produktteam adressieren sollte.</span><span class="sxs-lookup"><span data-stu-id="102f4-126">Creating an issue in the **OfficeDev/office-js** repository indicates you have found a flaw in the Office JavaScript API library that the product team should address.</span></span>

<span data-ttu-id="102f4-127">Wenn ein Problem mit dem Aktions Recorder oder-Editor auftritt, senden Sie Feedback über die Schaltfläche **Hilfe > Feedback** in Excel.</span><span class="sxs-lookup"><span data-stu-id="102f4-127">If there is a problem with the Action Recorder or Editor, send feedback through the **Help > Feedback** button in Excel.</span></span>

## <a name="see-also"></a><span data-ttu-id="102f4-128">Siehe auch</span><span class="sxs-lookup"><span data-stu-id="102f4-128">See also</span></span>

- [<span data-ttu-id="102f4-129">Office-Skripts in Excel im Internet</span><span class="sxs-lookup"><span data-stu-id="102f4-129">Office Scripts in Excel on the web</span></span>](../overview/excel.md)
- [<span data-ttu-id="102f4-130">Grundlegendes zur Skripterstellung für Office-Skripts in Excel im Internet</span><span class="sxs-lookup"><span data-stu-id="102f4-130">Scripting Fundamentals for Office Scripts in Excel on the web</span></span>](../develop/scripting-fundamentals.md)
- [<span data-ttu-id="102f4-131">Rückgängigmachen der Auswirkungen eines Office-Skripts</span><span class="sxs-lookup"><span data-stu-id="102f4-131">Undo the effects of an Office Script</span></span>](undo.md)
