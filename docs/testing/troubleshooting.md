---
title: Problembehandlung Office Skripts
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
# <a name="troubleshoot-office-scripts"></a><span data-ttu-id="5df35-103">Problembehandlung Office Skripts</span><span class="sxs-lookup"><span data-stu-id="5df35-103">Troubleshoot Office Scripts</span></span>

<span data-ttu-id="5df35-104">Wenn Sie skripts Office entwickeln, können Sie Fehler machen.</span><span class="sxs-lookup"><span data-stu-id="5df35-104">As you develop Office Scripts, you may make mistakes.</span></span> <span data-ttu-id="5df35-105">Das ist okay.</span><span class="sxs-lookup"><span data-stu-id="5df35-105">It's okay.</span></span> <span data-ttu-id="5df35-106">Sie verfügen über die Tools, um die Probleme zu finden und die Skripts optimal zu funktionieren.</span><span class="sxs-lookup"><span data-stu-id="5df35-106">You have the tools to help find the problems and get your scripts working perfectly.</span></span>

## <a name="types-of-errors"></a><span data-ttu-id="5df35-107">Fehlertypen</span><span class="sxs-lookup"><span data-stu-id="5df35-107">Types of errors</span></span>

<span data-ttu-id="5df35-108">Office Skriptfehler fallen in eine von zwei Kategorien:</span><span class="sxs-lookup"><span data-stu-id="5df35-108">Office Scripts errors fall into one of two categories:</span></span>

* <span data-ttu-id="5df35-109">Kompilierungszeitfehler oder -warnungen</span><span class="sxs-lookup"><span data-stu-id="5df35-109">Compile-time errors or warnings</span></span>
* <span data-ttu-id="5df35-110">Laufzeitfehler</span><span class="sxs-lookup"><span data-stu-id="5df35-110">Runtime errors</span></span>

### <a name="compile-time-errors"></a><span data-ttu-id="5df35-111">Kompilierungszeitfehler</span><span class="sxs-lookup"><span data-stu-id="5df35-111">Compile-time errors</span></span>

<span data-ttu-id="5df35-112">Kompilierungsfehler und Warnungen werden zunächst im Code-Editor angezeigt.</span><span class="sxs-lookup"><span data-stu-id="5df35-112">Compile-time errors and warnings are initially shown in the Code Editor.</span></span> <span data-ttu-id="5df35-113">Diese werden durch die wellenförmigen roten Unterstreichungen im Editor angezeigt.</span><span class="sxs-lookup"><span data-stu-id="5df35-113">These are shown by the wavy red underlines in the editor.</span></span> <span data-ttu-id="5df35-114">Sie werden auch unten  im Aufgabenbereich des Code-Editors unter der Registerkarte Probleme angezeigt.</span><span class="sxs-lookup"><span data-stu-id="5df35-114">They are also displayed under the **Problems** tab at the bottom of the Code Editor task pane.</span></span> <span data-ttu-id="5df35-115">Wenn Sie den Fehler auswählen, erhalten Sie weitere Details zum Problem und schlagen Lösungen vor.</span><span class="sxs-lookup"><span data-stu-id="5df35-115">Selecting the error will give more details about the problem and suggest solutions.</span></span> <span data-ttu-id="5df35-116">Kompilierungsfehler sollten vor dem Ausführen des Skripts behoben werden.</span><span class="sxs-lookup"><span data-stu-id="5df35-116">Compile-time errors should be addressed before running the script.</span></span>

:::image type="content" source="../images/explicit-any-editor-message.png" alt-text="Ein Compilerfehler, der im #A0 angezeigt wird":::

<span data-ttu-id="5df35-118">Möglicherweise werden auch orangefarbene Warnmeldungen und graue Informationsmeldungen angezeigt.</span><span class="sxs-lookup"><span data-stu-id="5df35-118">You may also see orange warning underlines and grey informational messages.</span></span> <span data-ttu-id="5df35-119">Diese zeigen Leistungsvorschläge oder andere Möglichkeiten an, bei denen das Skript unbeabsichtigte Auswirkungen haben kann.</span><span class="sxs-lookup"><span data-stu-id="5df35-119">These indicate performance suggestions or other possibilities where the script may have unintentional effects.</span></span> <span data-ttu-id="5df35-120">Solche Warnungen sollten vor dem Schließen genau geprüft werden.</span><span class="sxs-lookup"><span data-stu-id="5df35-120">Such warnings should be examined closely before dismissing them.</span></span>

### <a name="runtime-errors"></a><span data-ttu-id="5df35-121">Laufzeitfehler</span><span class="sxs-lookup"><span data-stu-id="5df35-121">Runtime errors</span></span>

<span data-ttu-id="5df35-122">Laufzeitfehler können aufgrund von Logikproblemen im Skript auftreten.</span><span class="sxs-lookup"><span data-stu-id="5df35-122">Runtime errors happen because of logic issues in the script.</span></span> <span data-ttu-id="5df35-123">Dies kann daran liegen, dass sich ein im Skript verwendetes Objekt nicht in der Arbeitsmappe befindet, eine Tabelle anders formatiert ist als erwartet oder eine andere geringfügige Diskrepanz zwischen den Anforderungen des Skripts und der aktuellen Arbeitsmappe besteht.</span><span class="sxs-lookup"><span data-stu-id="5df35-123">This could be because an object used in the script isn't in the workbook, a table is formatted differently than anticipated, or some other slight discrepancy between the script's requirements and the current workbook.</span></span> <span data-ttu-id="5df35-124">Das folgende Skript generiert einen Fehler, wenn ein Arbeitsblatt mit dem Namen "TestSheet" nicht vorhanden ist.</span><span class="sxs-lookup"><span data-stu-id="5df35-124">The following script generates an error when a worksheet named "TestSheet" is not present.</span></span>

```TypeScript
function main(workbook: ExcelScript.Workbook) {
  let mySheet = workbook.getWorksheet('TestSheet');

  // This will throw an error if there is no "TestSheet".
  mySheet.getRange("A1");
}
```

### <a name="console-messages"></a><span data-ttu-id="5df35-125">Konsolennachrichten</span><span class="sxs-lookup"><span data-stu-id="5df35-125">Console messages</span></span>

<span data-ttu-id="5df35-126">Sowohl Kompilierungszeit- als auch Laufzeitfehler zeigen Fehlermeldungen in der Konsole an, wenn ein Skript ausgeführt wird.</span><span class="sxs-lookup"><span data-stu-id="5df35-126">Both compile-time and runtime errors display error messages in the console when a script runs.</span></span> <span data-ttu-id="5df35-127">Sie geben eine Zeilennummer, in der das Problem aufgetreten ist.</span><span class="sxs-lookup"><span data-stu-id="5df35-127">They give a line number where the problem was encountered.</span></span> <span data-ttu-id="5df35-128">Denken Sie daran, dass die Hauptursache eines Problems möglicherweise eine andere Codezeile als die in der Konsole angegebenen ist.</span><span class="sxs-lookup"><span data-stu-id="5df35-128">Keep in mind that the root cause of any issue may be a different line of code than what is indicated in the console.</span></span>

<span data-ttu-id="5df35-129">Die folgende Abbildung zeigt die Konsolenausgabe für den [expliziten `any` Compilerfehler.](../develop/typescript-restrictions.md)</span><span class="sxs-lookup"><span data-stu-id="5df35-129">The following image shows the console output for the [explicit `any`](../develop/typescript-restrictions.md) compiler error.</span></span> <span data-ttu-id="5df35-130">Notieren Sie sich `[5, 16]` den Text am Anfang der Fehlerzeichenfolge.</span><span class="sxs-lookup"><span data-stu-id="5df35-130">Note the text `[5, 16]` at the beginning of the error string.</span></span> <span data-ttu-id="5df35-131">Dies gibt an, dass sich der Fehler in Zeile 5 befindet, beginnend mit Zeichen 16.</span><span class="sxs-lookup"><span data-stu-id="5df35-131">This indicates the error is on line 5, starting at character 16.</span></span>
:::image type="content" source="../images/explicit-any-error-message.png" alt-text="Die Code-Editor-Konsole, in der eine explizite Fehlermeldung &quot;any&quot; angezeigt wird":::

<span data-ttu-id="5df35-133">In der folgenden Abbildung wird die Konsolenausgabe für einen Laufzeitfehler angezeigt.</span><span class="sxs-lookup"><span data-stu-id="5df35-133">The follow image shows the console output for a runtime error.</span></span> <span data-ttu-id="5df35-134">Hier versucht das Skript, ein Arbeitsblatt mit dem Namen eines vorhandenen Arbeitsblatts hinzuzufügen.</span><span class="sxs-lookup"><span data-stu-id="5df35-134">Here, the script tries to add a worksheet with a the name of an existing worksheet.</span></span> <span data-ttu-id="5df35-135">Notieren Sie sich erneut die Zeile 2 vor dem Fehler, um zu zeigen, welche Zeile untersucht werden soll.</span><span class="sxs-lookup"><span data-stu-id="5df35-135">Again, note the "Line 2" preceding the error to show which line to investigate.</span></span>
:::image type="content" source="../images/runtime-error-console.png" alt-text="Die Code-Editor-Konsole, in der ein Fehler aus dem Aufruf &quot;addWorksheet&quot; angezeigt wird":::

## <a name="console-logs"></a><span data-ttu-id="5df35-137">Konsolenprotokolle</span><span class="sxs-lookup"><span data-stu-id="5df35-137">Console logs</span></span>

<span data-ttu-id="5df35-138">Drucken von Nachrichten auf dem Bildschirm mit der `console.log` Anweisung.</span><span class="sxs-lookup"><span data-stu-id="5df35-138">Print messages to the screen with the `console.log` statement.</span></span> <span data-ttu-id="5df35-139">In diesen Protokollen können Sie den aktuellen Wert von Variablen oder die Codepfade anzeigen, die ausgelöst werden.</span><span class="sxs-lookup"><span data-stu-id="5df35-139">These logs can show you the current value of variables or which code paths are being triggered.</span></span> <span data-ttu-id="5df35-140">Rufen Sie dazu ein `console.log` beliebiges Objekt als Parameter auf.</span><span class="sxs-lookup"><span data-stu-id="5df35-140">To do this, call `console.log` with any object as a parameter.</span></span> <span data-ttu-id="5df35-141">In der Regel `string` ist a der einfachste Typ, der in der Konsole gelesen werden kann.</span><span class="sxs-lookup"><span data-stu-id="5df35-141">Usually, a `string` is the easiest type to read in the console.</span></span>

```TypeScript
console.log("Logging myRange's address.");
console.log(myRange.getAddress());
```

<span data-ttu-id="5df35-142">An übergebene Zeichenfolgen werden in der Protokollierungskonsole des Code-Editors am unteren Rand `console.log` des Aufgabenbereichs angezeigt.</span><span class="sxs-lookup"><span data-stu-id="5df35-142">Strings passed to `console.log` are displayed in the Code Editor's logging console, at the bottom of the task pane.</span></span> <span data-ttu-id="5df35-143">Protokolle werden auf der Registerkarte **Ausgabe** gefunden, die Registerkarte erhält jedoch automatisch den Fokus, wenn ein Protokoll geschrieben wird.</span><span class="sxs-lookup"><span data-stu-id="5df35-143">Logs are found on the **Output** tab, though the tab automatically gains focus when a log is written.</span></span>

<span data-ttu-id="5df35-144">Protokolle wirken sich nicht auf die Arbeitsmappe aus.</span><span class="sxs-lookup"><span data-stu-id="5df35-144">Logs do not affect the workbook.</span></span>

## <a name="automate-tab-not-appearing-or-office-scripts-unavailable"></a><span data-ttu-id="5df35-145">Automatisieren der Registerkarte nicht angezeigt oder Office Skripts nicht verfügbar</span><span class="sxs-lookup"><span data-stu-id="5df35-145">Automate tab not appearing or Office Scripts unavailable</span></span>

<span data-ttu-id="5df35-146">Die folgenden Schritte sollten bei der Problembehandlung bei Problemen im Zusammenhang mit der Registerkarte **Automatisieren** helfen, die nicht in der Excel im Web.</span><span class="sxs-lookup"><span data-stu-id="5df35-146">The following steps should help troubleshoot any problems related to the **Automate** tab not appearing in Excel on the web.</span></span>

1. <span data-ttu-id="5df35-147">[Stellen Sie sicher, Microsoft 365 Ihre Lizenz skripts Office enthält.](../overview/excel.md#requirements)</span><span class="sxs-lookup"><span data-stu-id="5df35-147">[Make sure your Microsoft 365 license includes Office Scripts](../overview/excel.md#requirements).</span></span>
1. <span data-ttu-id="5df35-148">[Überprüfen Sie, ob Ihr Browser unterstützt wird.](platform-limits.md#browser-support)</span><span class="sxs-lookup"><span data-stu-id="5df35-148">[Check that your browser is supported](platform-limits.md#browser-support).</span></span>
1. <span data-ttu-id="5df35-149">[Stellen Sie sicher, dass Cookies von Drittanbietern aktiviert sind.](platform-limits.md#third-party-cookies)</span><span class="sxs-lookup"><span data-stu-id="5df35-149">[Ensure third-party cookies are enabled](platform-limits.md#third-party-cookies).</span></span>
1. <span data-ttu-id="5df35-150">[Stellen Sie sicher, dass Ihr Administrator die Skripts im Office Admin Center Microsoft 365 deaktiviert hat.](/microsoft-365/admin/manage/manage-office-scripts-settings)</span><span class="sxs-lookup"><span data-stu-id="5df35-150">[Ensure that your admin has not disabled Office Scripts in the Microsoft 365 admin center](/microsoft-365/admin/manage/manage-office-scripts-settings).</span></span>

[!INCLUDE [Teams support note](../includes/teams-support-note.md)]

## <a name="troubleshoot-scripts-in-power-automate"></a><span data-ttu-id="5df35-151">Behandeln von Skripts in Power Automate</span><span class="sxs-lookup"><span data-stu-id="5df35-151">Troubleshoot scripts in Power Automate</span></span>

<span data-ttu-id="5df35-152">Informationen zum Ausführen von Skripts über Power Automate finden Sie unter [Troubleshoot Office Scripts running in Power Automate](power-automate-troubleshooting.md).</span><span class="sxs-lookup"><span data-stu-id="5df35-152">For information specific to running scripts through Power Automate, see [Troubleshoot Office Scripts running in Power Automate](power-automate-troubleshooting.md).</span></span>

## <a name="help-resources"></a><span data-ttu-id="5df35-153">Hilferessourcen</span><span class="sxs-lookup"><span data-stu-id="5df35-153">Help resources</span></span>

<span data-ttu-id="5df35-154">[Stack Overflow](https://stackoverflow.com/questions/tagged/office-scripts) ist eine Community von Entwicklern, die bei Codierungsproblemen helfen.</span><span class="sxs-lookup"><span data-stu-id="5df35-154">[Stack Overflow](https://stackoverflow.com/questions/tagged/office-scripts) is a community of developers willing to help with coding problems.</span></span> <span data-ttu-id="5df35-155">Häufig können Sie die Lösung für Ihr Problem über eine schnelle Stack Overflow-Suche finden.</span><span class="sxs-lookup"><span data-stu-id="5df35-155">Often, you'll be able to find the solution to your problem through a quick Stack Overflow search.</span></span> <span data-ttu-id="5df35-156">Wenn nicht, stellen Sie Ihre Frage, und markieren Sie sie mit dem Tag "office-scripts".</span><span class="sxs-lookup"><span data-stu-id="5df35-156">If not, ask your question and tag it with the "office-scripts" tag.</span></span> <span data-ttu-id="5df35-157">Achten Sie darauf, zu erwähnen, dass Sie ein *Office-Skript* erstellen, nicht ein *Office-Add-In .*</span><span class="sxs-lookup"><span data-stu-id="5df35-157">Be sure to mention you're creating an Office *Script*, not an Office *Add-in*.</span></span>

<span data-ttu-id="5df35-158">Wenn ein Problem mit der Office JavaScript-API besteht, erstellen Sie ein Problem im [OfficeDev/office-js-GitHub](https://github.com/OfficeDev/office-js) Repository.</span><span class="sxs-lookup"><span data-stu-id="5df35-158">If you encounter a problem with the Office JavaScript API, create an issue in the [OfficeDev/office-js](https://github.com/OfficeDev/office-js) GitHub repository.</span></span> <span data-ttu-id="5df35-159">Mitglieder des Produktteams reagieren auf Probleme und bieten weitere Unterstützung.</span><span class="sxs-lookup"><span data-stu-id="5df35-159">Members of the product team will respond to issues and provide further assistance.</span></span> <span data-ttu-id="5df35-160">Das Erstellen eines Problems im **OfficeDev/office-js-Repository** weist darauf hin, dass sie einen Fehler in der Office JavaScript-API-Bibliothek gefunden haben, die das Produktteam beheben sollte.</span><span class="sxs-lookup"><span data-stu-id="5df35-160">Creating an issue in the **OfficeDev/office-js** repository indicates you have found a flaw in the Office JavaScript API library that the product team should address.</span></span>

<span data-ttu-id="5df35-161">Wenn ein Problem mit der Aktionsaufzeichnung oder dem Editor vor liegt, senden Sie Feedback über die Schaltfläche Hilfe **> Feedback** in Excel.</span><span class="sxs-lookup"><span data-stu-id="5df35-161">If there is a problem with the Action Recorder or Editor, send feedback through the **Help > Feedback** button in Excel.</span></span>

## <a name="see-also"></a><span data-ttu-id="5df35-162">Siehe auch</span><span class="sxs-lookup"><span data-stu-id="5df35-162">See also</span></span>

- [<span data-ttu-id="5df35-163">Bewährte Methoden in Office-Skripts</span><span class="sxs-lookup"><span data-stu-id="5df35-163">Best practices in Office Scripts</span></span>](../develop/best-practices.md)
- [<span data-ttu-id="5df35-164">Plattformbeschränkungen mit Office Skripts</span><span class="sxs-lookup"><span data-stu-id="5df35-164">Platform limits with Office Scripts</span></span>](platform-limits.md)
- [<span data-ttu-id="5df35-165">Verbessern der Leistung Ihrer Office Skripts</span><span class="sxs-lookup"><span data-stu-id="5df35-165">Improve the performance of your Office Scripts</span></span>](../develop/web-client-performance.md)
- [<span data-ttu-id="5df35-166">Problembehandlung Office in PowerAutomate ausgeführten Skripts</span><span class="sxs-lookup"><span data-stu-id="5df35-166">Troubleshoot Office Scripts running in PowerAutomate</span></span>](power-automate-troubleshooting.md)
- [<span data-ttu-id="5df35-167">Auswirkungen von Office-Skripts rückgängig machen</span><span class="sxs-lookup"><span data-stu-id="5df35-167">Undo the effects of Office Scripts</span></span>](undo.md)
