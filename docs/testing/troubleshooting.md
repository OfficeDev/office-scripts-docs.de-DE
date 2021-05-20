---
title: Fehlerbehebung Office Skripts
description: Debuggen von Tipps und Techniken für Office Skripts sowie Hilferessourcen.
ms.date: 05/17/2021
localization_priority: Normal
ms.openlocfilehash: ff0ac1e63084c7c541d2a4925f1f011d16fa4992
ms.sourcegitcommit: 4687693f02fc90a57ba30c461f35046e02e6f5fb
ms.translationtype: MT
ms.contentlocale: de-DE
ms.lasthandoff: 05/19/2021
ms.locfileid: "52545556"
---
# <a name="troubleshoot-office-scripts"></a><span data-ttu-id="5af7a-103">Fehlerbehebung Office Skripts</span><span class="sxs-lookup"><span data-stu-id="5af7a-103">Troubleshoot Office Scripts</span></span>

<span data-ttu-id="5af7a-104">Wenn Sie Office-Skripts entwickeln, können Sie Fehler machen.</span><span class="sxs-lookup"><span data-stu-id="5af7a-104">As you develop Office Scripts, you may make mistakes.</span></span> <span data-ttu-id="5af7a-105">Das ist okay.</span><span class="sxs-lookup"><span data-stu-id="5af7a-105">It's okay.</span></span> <span data-ttu-id="5af7a-106">Sie haben die Tools, um die Probleme zu finden und Ihre Skripte perfekt funktionieren zu lassen.</span><span class="sxs-lookup"><span data-stu-id="5af7a-106">You have the tools to help find the problems and get your scripts working perfectly.</span></span>

## <a name="types-of-errors"></a><span data-ttu-id="5af7a-107">Arten von Fehlern</span><span class="sxs-lookup"><span data-stu-id="5af7a-107">Types of errors</span></span>

<span data-ttu-id="5af7a-108">Office Skriptfehler lassen sich in eine von zwei Kategorien einteilen:</span><span class="sxs-lookup"><span data-stu-id="5af7a-108">Office Scripts errors fall into one of two categories:</span></span>

* <span data-ttu-id="5af7a-109">Kompilierungszeitfehler oder Warnungen</span><span class="sxs-lookup"><span data-stu-id="5af7a-109">Compile-time errors or warnings</span></span>
* <span data-ttu-id="5af7a-110">Laufzeitfehler</span><span class="sxs-lookup"><span data-stu-id="5af7a-110">Runtime errors</span></span>

### <a name="compile-time-errors"></a><span data-ttu-id="5af7a-111">Kompilierungszeitfehler</span><span class="sxs-lookup"><span data-stu-id="5af7a-111">Compile-time errors</span></span>

<span data-ttu-id="5af7a-112">Kompilierzeitfehler und Warnungen werden zunächst im Code-Editor angezeigt.</span><span class="sxs-lookup"><span data-stu-id="5af7a-112">Compile-time errors and warnings are initially shown in the Code Editor.</span></span> <span data-ttu-id="5af7a-113">Diese zeigen die welligen roten Unterstreichungen im Editor.</span><span class="sxs-lookup"><span data-stu-id="5af7a-113">These are shown by the wavy red underlines in the editor.</span></span> <span data-ttu-id="5af7a-114">Sie werden auch unter der Registerkarte **Probleme** am unteren Rand des Aufgabenbereichs Code-Editor angezeigt.</span><span class="sxs-lookup"><span data-stu-id="5af7a-114">They are also displayed under the **Problems** tab at the bottom of the Code Editor task pane.</span></span> <span data-ttu-id="5af7a-115">Wenn Sie den Fehler auswählen, werden weitere Details zum Problem angezeigt und Lösungen vorgeschlagen.</span><span class="sxs-lookup"><span data-stu-id="5af7a-115">Selecting the error will give more details about the problem and suggest solutions.</span></span> <span data-ttu-id="5af7a-116">Kompilierungsfehler sollten vor dem Ausführen des Skripts behoben werden.</span><span class="sxs-lookup"><span data-stu-id="5af7a-116">Compile-time errors should be addressed before running the script.</span></span>

:::image type="content" source="../images/explicit-any-editor-message.png" alt-text="Ein Compilerfehler, der im Hovertext des Code-Editors angezeigt wird":::

<span data-ttu-id="5af7a-118">Möglicherweise werden auch orangefarbene Warnhinweise und graue Informationsmeldungen angezeigt.</span><span class="sxs-lookup"><span data-stu-id="5af7a-118">You may also see orange warning underlines and grey informational messages.</span></span> <span data-ttu-id="5af7a-119">Diese weisen auf Leistungsvorschläge oder andere Möglichkeiten hin, bei denen das Skript unbeabsichtigte Auswirkungen haben kann.</span><span class="sxs-lookup"><span data-stu-id="5af7a-119">These indicate performance suggestions or other possibilities where the script may have unintentional effects.</span></span> <span data-ttu-id="5af7a-120">Solche Warnungen sollten eingehend geprüft werden, bevor sie zurückgewiesen werden.</span><span class="sxs-lookup"><span data-stu-id="5af7a-120">Such warnings should be examined closely before dismissing them.</span></span>

### <a name="runtime-errors"></a><span data-ttu-id="5af7a-121">Laufzeitfehler</span><span class="sxs-lookup"><span data-stu-id="5af7a-121">Runtime errors</span></span>

<span data-ttu-id="5af7a-122">Laufzeitfehler treten aufgrund von Logikproblemen im Skript auf.</span><span class="sxs-lookup"><span data-stu-id="5af7a-122">Runtime errors happen because of logic issues in the script.</span></span> <span data-ttu-id="5af7a-123">Dies kann daran liegen, dass sich ein im Skript verwendetes Objekt nicht in der Arbeitsmappe befindet, eine Tabelle anders formatiert ist als erwartet, oder eine andere geringfügige Diskrepanz zwischen den Anforderungen des Skripts und der aktuellen Arbeitsmappe besteht.</span><span class="sxs-lookup"><span data-stu-id="5af7a-123">This could be because an object used in the script isn't in the workbook, a table is formatted differently than anticipated, or some other slight discrepancy between the script's requirements and the current workbook.</span></span> <span data-ttu-id="5af7a-124">Das folgende Skript generiert einen Fehler, wenn kein Arbeitsblatt mit dem Namen "TestSheet" vorhanden ist.</span><span class="sxs-lookup"><span data-stu-id="5af7a-124">The following script generates an error when a worksheet named "TestSheet" is not present.</span></span>

```TypeScript
function main(workbook: ExcelScript.Workbook) {
  let mySheet = workbook.getWorksheet('TestSheet');

  // This will throw an error if there is no "TestSheet".
  mySheet.getRange("A1");
}
```

### <a name="console-messages"></a><span data-ttu-id="5af7a-125">Konsolenmeldungen</span><span class="sxs-lookup"><span data-stu-id="5af7a-125">Console messages</span></span>

<span data-ttu-id="5af7a-126">Sowohl Kompilierungszeit- als auch Laufzeitfehler zeigen Fehlermeldungen in der Konsole an, wenn ein Skript ausgeführt wird.</span><span class="sxs-lookup"><span data-stu-id="5af7a-126">Both compile-time and runtime errors display error messages in the console when a script runs.</span></span> <span data-ttu-id="5af7a-127">Sie geben eine Zeilennummer an, bei der das Problem aufgetreten ist.</span><span class="sxs-lookup"><span data-stu-id="5af7a-127">They give a line number where the problem was encountered.</span></span> <span data-ttu-id="5af7a-128">Beachten Sie, dass die Ursache eines Problems möglicherweise eine andere Codezeile als in der Konsole angegeben ist.</span><span class="sxs-lookup"><span data-stu-id="5af7a-128">Keep in mind that the root cause of any issue may be a different line of code than what is indicated in the console.</span></span>

<span data-ttu-id="5af7a-129">Die folgende Abbildung zeigt die Konsolenausgabe für den [expliziten `any` ](../develop/typescript-restrictions.md) Compilerfehler.</span><span class="sxs-lookup"><span data-stu-id="5af7a-129">The following image shows the console output for the [explicit `any`](../develop/typescript-restrictions.md) compiler error.</span></span> <span data-ttu-id="5af7a-130">Beachten Sie den Text `[5, 16]` am Anfang der Fehlerzeichenfolge.</span><span class="sxs-lookup"><span data-stu-id="5af7a-130">Note the text `[5, 16]` at the beginning of the error string.</span></span> <span data-ttu-id="5af7a-131">Dies zeigt an, dass sich der Fehler in Zeile 5 befindet, beginnend mit Zeichen 16.</span><span class="sxs-lookup"><span data-stu-id="5af7a-131">This indicates the error is on line 5, starting at character 16.</span></span>
:::image type="content" source="../images/explicit-any-error-message.png" alt-text="Die Code-Editor-Konsole, die eine explizite &quot;any&quot;-Fehlermeldung anzeigt":::

<span data-ttu-id="5af7a-133">Das folgende Bild zeigt die Konsolenausgabe für einen Laufzeitfehler.</span><span class="sxs-lookup"><span data-stu-id="5af7a-133">The follow image shows the console output for a runtime error.</span></span> <span data-ttu-id="5af7a-134">Hier versucht das Skript, ein Arbeitsblatt mit dem Namen eines vorhandenen Arbeitsblatts hinzuzufügen.</span><span class="sxs-lookup"><span data-stu-id="5af7a-134">Here, the script tries to add a worksheet with a the name of an existing worksheet.</span></span> <span data-ttu-id="5af7a-135">Beachten Sie erneut die "Zeile 2" vor dem Fehler, um anzuzeigen, welche Zeile untersucht werden soll.</span><span class="sxs-lookup"><span data-stu-id="5af7a-135">Again, note the "Line 2" preceding the error to show which line to investigate.</span></span>
:::image type="content" source="../images/runtime-error-console.png" alt-text="Die Code-Editor-Konsole, die einen Fehler aus dem &quot;addWorksheet&quot;-Aufruf anzeigt":::

## <a name="console-logs"></a><span data-ttu-id="5af7a-137">Konsolenprotokolle</span><span class="sxs-lookup"><span data-stu-id="5af7a-137">Console logs</span></span>

<span data-ttu-id="5af7a-138">Drucken Sie Nachrichten mit der Anweisung auf den `console.log` Bildschirm.</span><span class="sxs-lookup"><span data-stu-id="5af7a-138">Print messages to the screen with the `console.log` statement.</span></span> <span data-ttu-id="5af7a-139">Diese Protokolle können Ihnen den aktuellen Wert von Variablen anzeigen oder zeigen, welche Codepfade ausgelöst werden.</span><span class="sxs-lookup"><span data-stu-id="5af7a-139">These logs can show you the current value of variables or which code paths are being triggered.</span></span> <span data-ttu-id="5af7a-140">Rufen Sie dazu `console.log` ein beliebiges Objekt als Parameter auf.</span><span class="sxs-lookup"><span data-stu-id="5af7a-140">To do this, call `console.log` with any object as a parameter.</span></span> <span data-ttu-id="5af7a-141">In der Regel ist a der einfachste Typ, der in der Konsole gelesen werden `string` kann.</span><span class="sxs-lookup"><span data-stu-id="5af7a-141">Usually, a `string` is the easiest type to read in the console.</span></span>

```TypeScript
console.log("Logging myRange's address.");
console.log(myRange.getAddress());
```

<span data-ttu-id="5af7a-142">An Strings, die übergeben `console.log` werden, werden in der Protokollierungskonsole des Code-Editors am unteren Rand des Aufgabenbereichs angezeigt.</span><span class="sxs-lookup"><span data-stu-id="5af7a-142">Strings passed to `console.log` are displayed in the Code Editor's logging console, at the bottom of the task pane.</span></span> <span data-ttu-id="5af7a-143">Protokolle werden auf der Registerkarte **Ausgabe** gefunden, obwohl die Registerkarte automatisch den Fokus erhält, wenn ein Protokoll geschrieben wird.</span><span class="sxs-lookup"><span data-stu-id="5af7a-143">Logs are found on the **Output** tab, though the tab automatically gains focus when a log is written.</span></span>

<span data-ttu-id="5af7a-144">Protokolle wirken sich nicht auf die Arbeitsmappe aus.</span><span class="sxs-lookup"><span data-stu-id="5af7a-144">Logs do not affect the workbook.</span></span>

## <a name="automate-tab-not-appearing-or-office-scripts-unavailable"></a><span data-ttu-id="5af7a-145">Registerkarte automatisieren wird nicht angezeigt oder Office Skripts nicht verfügbar</span><span class="sxs-lookup"><span data-stu-id="5af7a-145">Automate tab not appearing or Office Scripts unavailable</span></span>

<span data-ttu-id="5af7a-146">Die folgenden Schritte sollten zur Behebung von Problemen im Zusammenhang mit der Registerkarte **Automatisieren** helfen, die nicht in Excel im Web angezeigt werden.</span><span class="sxs-lookup"><span data-stu-id="5af7a-146">The following steps should help troubleshoot any problems related to the **Automate** tab not appearing in Excel on the web.</span></span>

1. <span data-ttu-id="5af7a-147">[Stellen Sie sicher, dass Ihre Microsoft 365 Lizenz Office Scripts enthält.](../overview/excel.md#requirements)</span><span class="sxs-lookup"><span data-stu-id="5af7a-147">[Make sure your Microsoft 365 license includes Office Scripts](../overview/excel.md#requirements).</span></span>
1. <span data-ttu-id="5af7a-148">[Überprüfen Sie, ob Ihr Browser unterstützt wird.](platform-limits.md#browser-support)</span><span class="sxs-lookup"><span data-stu-id="5af7a-148">[Check that your browser is supported](platform-limits.md#browser-support).</span></span>
1. <span data-ttu-id="5af7a-149">[Stellen Sie sicher, dass Cookies von Drittanbietern aktiviert sind](platform-limits.md#third-party-cookies).</span><span class="sxs-lookup"><span data-stu-id="5af7a-149">[Ensure third-party cookies are enabled](platform-limits.md#third-party-cookies).</span></span>
1. <span data-ttu-id="5af7a-150">[Stellen Sie sicher, dass Ihr Administrator Office Skripts im Microsoft 365 Admin Center nicht deaktiviert hat.](/microsoft-365/admin/manage/manage-office-scripts-settings)</span><span class="sxs-lookup"><span data-stu-id="5af7a-150">[Ensure that your admin has not disabled Office Scripts in the Microsoft 365 admin center](/microsoft-365/admin/manage/manage-office-scripts-settings).</span></span>

[!INCLUDE [Teams support note](../includes/teams-support-note.md)]

## <a name="troubleshoot-scripts-in-power-automate"></a><span data-ttu-id="5af7a-151">Fehlerbehebung von Skripts in Power Automate</span><span class="sxs-lookup"><span data-stu-id="5af7a-151">Troubleshoot scripts in Power Automate</span></span>

<span data-ttu-id="5af7a-152">Informationen zum Ausführen von Skripts durch Power Automate finden Sie unter [Problembehandlung Office Skripts, die in Power Automate ausgeführt werden.](power-automate-troubleshooting.md)</span><span class="sxs-lookup"><span data-stu-id="5af7a-152">For information specific to running scripts through Power Automate, see [Troubleshoot Office Scripts running in Power Automate](power-automate-troubleshooting.md).</span></span>

## <a name="help-resources"></a><span data-ttu-id="5af7a-153">Hilferessourcen</span><span class="sxs-lookup"><span data-stu-id="5af7a-153">Help resources</span></span>

<span data-ttu-id="5af7a-154">[Stack Overflow](https://stackoverflow.com/questions/tagged/office-scripts) ist eine Community von Entwicklern, die bereit sind, bei Codierungsproblemen zu helfen.</span><span class="sxs-lookup"><span data-stu-id="5af7a-154">[Stack Overflow](https://stackoverflow.com/questions/tagged/office-scripts) is a community of developers willing to help with coding problems.</span></span> <span data-ttu-id="5af7a-155">Häufig können Sie die Lösung für Ihr Problem durch eine schnelle Stack Overflow-Suche finden.</span><span class="sxs-lookup"><span data-stu-id="5af7a-155">Often, you'll be able to find the solution to your problem through a quick Stack Overflow search.</span></span> <span data-ttu-id="5af7a-156">Wenn nicht, stellen Sie Ihre Frage und markieren Sie sie mit dem Tag "office-scripts".</span><span class="sxs-lookup"><span data-stu-id="5af7a-156">If not, ask your question and tag it with the "office-scripts" tag.</span></span> <span data-ttu-id="5af7a-157">Erwähnen Sie, dass Sie ein Office *Skript* erstellen, nicht ein Office *Add-In*.</span><span class="sxs-lookup"><span data-stu-id="5af7a-157">Be sure to mention you're creating an Office *Script*, not an Office *Add-in*.</span></span>

<span data-ttu-id="5af7a-158">Wenn ein Problem mit der Office JavaScript-API auftritt, erstellen Sie ein Problem im [Repository OfficeDev/office-js](https://github.com/OfficeDev/office-js) GitHub.</span><span class="sxs-lookup"><span data-stu-id="5af7a-158">If you encounter a problem with the Office JavaScript API, create an issue in the [OfficeDev/office-js](https://github.com/OfficeDev/office-js) GitHub repository.</span></span> <span data-ttu-id="5af7a-159">Die Mitglieder des Produktteams werden auf Probleme reagieren und weitere Unterstützung leisten.</span><span class="sxs-lookup"><span data-stu-id="5af7a-159">Members of the product team will respond to issues and provide further assistance.</span></span> <span data-ttu-id="5af7a-160">Das Erstellen eines Problems im **OfficeDev/office-js-Repository** zeigt an, dass Sie einen Fehler in der Office JavaScript-API-Bibliothek gefunden haben, den das Produktteam beheben sollte.</span><span class="sxs-lookup"><span data-stu-id="5af7a-160">Creating an issue in the **OfficeDev/office-js** repository indicates you have found a flaw in the Office JavaScript API library that the product team should address.</span></span>

<span data-ttu-id="5af7a-161">Wenn ein Problem mit dem Aktionsrekorder oder -editor vorliegt, senden Sie Feedback über die Schaltfläche **Hilfe > Feedback** in Excel.</span><span class="sxs-lookup"><span data-stu-id="5af7a-161">If there is a problem with the Action Recorder or Editor, send feedback through the **Help > Feedback** button in Excel.</span></span>

## <a name="see-also"></a><span data-ttu-id="5af7a-162">Siehe auch</span><span class="sxs-lookup"><span data-stu-id="5af7a-162">See also</span></span>

- [<span data-ttu-id="5af7a-163">Bewährte Methoden in Office-Skripts</span><span class="sxs-lookup"><span data-stu-id="5af7a-163">Best practices in Office Scripts</span></span>](../develop/best-practices.md)
- [<span data-ttu-id="5af7a-164">Plattformlimits mit Office Scripts</span><span class="sxs-lookup"><span data-stu-id="5af7a-164">Platform limits with Office Scripts</span></span>](platform-limits.md)
- [<span data-ttu-id="5af7a-165">Verbessern Sie die Leistung Ihrer Office Scripts</span><span class="sxs-lookup"><span data-stu-id="5af7a-165">Improve the performance of your Office Scripts</span></span>](../develop/web-client-performance.md)
- [<span data-ttu-id="5af7a-166">Beheben Office Skripts, die in PowerAutomate ausgeführt werden</span><span class="sxs-lookup"><span data-stu-id="5af7a-166">Troubleshoot Office Scripts running in PowerAutomate</span></span>](power-automate-troubleshooting.md)
- [<span data-ttu-id="5af7a-167">Auswirkungen von Office-Skripts rückgängig machen</span><span class="sxs-lookup"><span data-stu-id="5af7a-167">Undo the effects of Office Scripts</span></span>](undo.md)
