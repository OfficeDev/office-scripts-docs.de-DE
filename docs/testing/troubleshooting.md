---
title: Problembehandlung bei Office Skripts
description: Debugtipps und -techniken für Office Skripts sowie Hilferessourcen.
ms.date: 05/17/2021
localization_priority: Normal
ms.openlocfilehash: 251ad72588422a86c52c81666164c2c4bd79bdb5
ms.sourcegitcommit: 4693c8f79428ec74695328275703af0ba1bfea8f
ms.translationtype: MT
ms.contentlocale: de-DE
ms.lasthandoff: 06/23/2021
ms.locfileid: "53074648"
---
# <a name="troubleshoot-office-scripts"></a><span data-ttu-id="559c2-103">Problembehandlung bei Office Skripts</span><span class="sxs-lookup"><span data-stu-id="559c2-103">Troubleshoot Office Scripts</span></span>

<span data-ttu-id="559c2-104">Wenn Sie Office Skripts entwickeln, machen Sie möglicherweise Fehler.</span><span class="sxs-lookup"><span data-stu-id="559c2-104">As you develop Office Scripts, you may make mistakes.</span></span> <span data-ttu-id="559c2-105">Das ist okay.</span><span class="sxs-lookup"><span data-stu-id="559c2-105">It's okay.</span></span> <span data-ttu-id="559c2-106">Sie verfügen über die Tools, mit denen Sie die Probleme finden und Ihre Skripts einwandfrei funktionieren können.</span><span class="sxs-lookup"><span data-stu-id="559c2-106">You have the tools to help find the problems and get your scripts working perfectly.</span></span>

## <a name="types-of-errors"></a><span data-ttu-id="559c2-107">Fehlertypen</span><span class="sxs-lookup"><span data-stu-id="559c2-107">Types of errors</span></span>

<span data-ttu-id="559c2-108">Office Skriptfehler werden in eine von zwei Kategorien unterteilt:</span><span class="sxs-lookup"><span data-stu-id="559c2-108">Office Scripts errors fall into one of two categories:</span></span>

* <span data-ttu-id="559c2-109">Kompilierungszeitfehler oder Warnungen</span><span class="sxs-lookup"><span data-stu-id="559c2-109">Compile-time errors or warnings</span></span>
* <span data-ttu-id="559c2-110">Laufzeitfehler</span><span class="sxs-lookup"><span data-stu-id="559c2-110">Runtime errors</span></span>

### <a name="compile-time-errors"></a><span data-ttu-id="559c2-111">Kompilierungszeitfehler</span><span class="sxs-lookup"><span data-stu-id="559c2-111">Compile-time errors</span></span>

<span data-ttu-id="559c2-112">Fehler und Warnungen zur Kompilierungszeit werden zunächst im Code-Editor angezeigt.</span><span class="sxs-lookup"><span data-stu-id="559c2-112">Compile-time errors and warnings are initially shown in the Code Editor.</span></span> <span data-ttu-id="559c2-113">Diese werden von den wellenförmigen roten Unterstreichungen im Editor angezeigt.</span><span class="sxs-lookup"><span data-stu-id="559c2-113">These are shown by the wavy red underlines in the editor.</span></span> <span data-ttu-id="559c2-114">Sie werden auch unter der Registerkarte **"Probleme"** am unteren Rand des Code-Editor-Aufgabenbereichs angezeigt.</span><span class="sxs-lookup"><span data-stu-id="559c2-114">They are also displayed under the **Problems** tab at the bottom of the Code Editor task pane.</span></span> <span data-ttu-id="559c2-115">Wenn Sie den Fehler auswählen, erhalten Sie weitere Details zu dem Problem und schlagen Lösungen vor.</span><span class="sxs-lookup"><span data-stu-id="559c2-115">Selecting the error will give more details about the problem and suggest solutions.</span></span> <span data-ttu-id="559c2-116">Kompilierungszeitfehler sollten behoben werden, bevor das Skript ausgeführt wird.</span><span class="sxs-lookup"><span data-stu-id="559c2-116">Compile-time errors should be addressed before running the script.</span></span>

:::image type="content" source="../images/explicit-any-editor-message.png" alt-text="Ein Compilerfehler, der im Hovertext des Code-Editors angezeigt wird.":::

<span data-ttu-id="559c2-118">Möglicherweise werden auch orangefarbene Warnhinweise und graue Informationsmeldungen angezeigt.</span><span class="sxs-lookup"><span data-stu-id="559c2-118">You may also see orange warning underlines and grey informational messages.</span></span> <span data-ttu-id="559c2-119">Diese deuten auf Leistungsvorschläge oder andere Möglichkeiten hin, bei denen das Skript unbeabsichtigte Effekte haben kann.</span><span class="sxs-lookup"><span data-stu-id="559c2-119">These indicate performance suggestions or other possibilities where the script may have unintentional effects.</span></span> <span data-ttu-id="559c2-120">Solche Warnungen sollten sorgfältig geprüft werden, bevor sie geschlossen werden.</span><span class="sxs-lookup"><span data-stu-id="559c2-120">Such warnings should be examined closely before dismissing them.</span></span>

### <a name="runtime-errors"></a><span data-ttu-id="559c2-121">Laufzeitfehler</span><span class="sxs-lookup"><span data-stu-id="559c2-121">Runtime errors</span></span>

<span data-ttu-id="559c2-122">Laufzeitfehler treten aufgrund von Logikproblemen im Skript auf.</span><span class="sxs-lookup"><span data-stu-id="559c2-122">Runtime errors happen because of logic issues in the script.</span></span> <span data-ttu-id="559c2-123">Dies kann darauf zurückzuführen sein, dass sich ein im Skript verwendetes Objekt nicht in der Arbeitsmappe befindet, eine Tabelle anders formatiert ist als erwartet, oder eine andere geringfügige Abweichung zwischen den Anforderungen des Skripts und der aktuellen Arbeitsmappe.</span><span class="sxs-lookup"><span data-stu-id="559c2-123">This could be because an object used in the script isn't in the workbook, a table is formatted differently than anticipated, or some other slight discrepancy between the script's requirements and the current workbook.</span></span> <span data-ttu-id="559c2-124">Das folgende Skript generiert einen Fehler, wenn kein Arbeitsblatt mit dem Namen "TestSheet" vorhanden ist.</span><span class="sxs-lookup"><span data-stu-id="559c2-124">The following script generates an error when a worksheet named "TestSheet" is not present.</span></span>

```TypeScript
function main(workbook: ExcelScript.Workbook) {
  let mySheet = workbook.getWorksheet('TestSheet');

  // This will throw an error if there is no "TestSheet".
  mySheet.getRange("A1");
}
```

### <a name="console-messages"></a><span data-ttu-id="559c2-125">Konsolenmeldungen</span><span class="sxs-lookup"><span data-stu-id="559c2-125">Console messages</span></span>

<span data-ttu-id="559c2-126">Bei Kompilierungs- und Laufzeitfehlern werden Fehlermeldungen in der Konsole angezeigt, wenn ein Skript ausgeführt wird.</span><span class="sxs-lookup"><span data-stu-id="559c2-126">Both compile-time and runtime errors display error messages in the console when a script runs.</span></span> <span data-ttu-id="559c2-127">Sie geben eine Zeilennummer an, in der das Problem aufgetreten ist.</span><span class="sxs-lookup"><span data-stu-id="559c2-127">They give a line number where the problem was encountered.</span></span> <span data-ttu-id="559c2-128">Beachten Sie, dass die Ursache eines Problems möglicherweise eine andere Codezeile als die in der Konsole angegebene ist.</span><span class="sxs-lookup"><span data-stu-id="559c2-128">Keep in mind that the root cause of any issue may be a different line of code than what is indicated in the console.</span></span>

<span data-ttu-id="559c2-129">Die folgende Abbildung zeigt die Konsolenausgabe für den [expliziten `any` ](../develop/typescript-restrictions.md) Compilerfehler.</span><span class="sxs-lookup"><span data-stu-id="559c2-129">The following image shows the console output for the [explicit `any`](../develop/typescript-restrictions.md) compiler error.</span></span> <span data-ttu-id="559c2-130">Notieren Sie sich den Text `[5, 16]` am Anfang der Fehlerzeichenfolge.</span><span class="sxs-lookup"><span data-stu-id="559c2-130">Note the text `[5, 16]` at the beginning of the error string.</span></span> <span data-ttu-id="559c2-131">Dies weist darauf hin, dass sich der Fehler in Zeile 5 befindet, beginnend mit Zeichen 16.</span><span class="sxs-lookup"><span data-stu-id="559c2-131">This indicates the error is on line 5, starting at character 16.</span></span>
:::image type="content" source="../images/explicit-any-error-message.png" alt-text="In der Code-Editor-Konsole wird eine explizite Any-Fehlermeldung angezeigt.":::

<span data-ttu-id="559c2-133">Die folgende Abbildung zeigt die Konsolenausgabe für einen Laufzeitfehler.</span><span class="sxs-lookup"><span data-stu-id="559c2-133">The follow image shows the console output for a runtime error.</span></span> <span data-ttu-id="559c2-134">Hier versucht das Skript, ein Arbeitsblatt mit dem Namen eines vorhandenen Arbeitsblatts hinzuzufügen.</span><span class="sxs-lookup"><span data-stu-id="559c2-134">Here, the script tries to add a worksheet with a the name of an existing worksheet.</span></span> <span data-ttu-id="559c2-135">Beachten Sie erneut die "Zeile 2" vor dem Fehler, um anzuzeigen, welche Zeile untersucht werden soll.</span><span class="sxs-lookup"><span data-stu-id="559c2-135">Again, note the "Line 2" preceding the error to show which line to investigate.</span></span>
:::image type="content" source="../images/runtime-error-console.png" alt-text="In der Code-Editor-Konsole wird ein Fehler aus dem Aufruf von &quot;addWorksheet&quot; angezeigt.":::

## <a name="console-logs"></a><span data-ttu-id="559c2-137">Konsolenprotokolle</span><span class="sxs-lookup"><span data-stu-id="559c2-137">Console logs</span></span>

<span data-ttu-id="559c2-138">Drucken Sie Nachrichten mit der Anweisung auf dem `console.log` Bildschirm.</span><span class="sxs-lookup"><span data-stu-id="559c2-138">Print messages to the screen with the `console.log` statement.</span></span> <span data-ttu-id="559c2-139">In diesen Protokollen können Sie den aktuellen Wert von Variablen anzeigen oder welche Codepfade ausgelöst werden.</span><span class="sxs-lookup"><span data-stu-id="559c2-139">These logs can show you the current value of variables or which code paths are being triggered.</span></span> <span data-ttu-id="559c2-140">Rufen Sie dazu `console.log` ein beliebiges Objekt als Parameter auf.</span><span class="sxs-lookup"><span data-stu-id="559c2-140">To do this, call `console.log` with any object as a parameter.</span></span> <span data-ttu-id="559c2-141">In der Regel ist a `string` der einfachste Typ, der in der Konsole gelesen werden kann.</span><span class="sxs-lookup"><span data-stu-id="559c2-141">Usually, a `string` is the easiest type to read in the console.</span></span>

```TypeScript
console.log("Logging myRange's address.");
console.log(myRange.getAddress());
```

<span data-ttu-id="559c2-142">Übergebene Zeichenfolgen `console.log` werden in der Protokollierungskonsole des Code-Editors am unteren Rand des Aufgabenbereichs angezeigt.</span><span class="sxs-lookup"><span data-stu-id="559c2-142">Strings passed to `console.log` are displayed in the Code Editor's logging console, at the bottom of the task pane.</span></span> <span data-ttu-id="559c2-143">Protokolle werden auf der Registerkarte **"Ausgabe"** gefunden, obwohl die Registerkarte automatisch den Fokus erhält, wenn ein Protokoll geschrieben wird.</span><span class="sxs-lookup"><span data-stu-id="559c2-143">Logs are found on the **Output** tab, though the tab automatically gains focus when a log is written.</span></span>

<span data-ttu-id="559c2-144">Protokolle wirken sich nicht auf die Arbeitsmappe aus.</span><span class="sxs-lookup"><span data-stu-id="559c2-144">Logs do not affect the workbook.</span></span>

## <a name="automate-tab-not-appearing-or-office-scripts-unavailable"></a><span data-ttu-id="559c2-145">Registerkarte "Automatisieren" wird nicht angezeigt, oder Office Skripts nicht verfügbar</span><span class="sxs-lookup"><span data-stu-id="559c2-145">Automate tab not appearing or Office Scripts unavailable</span></span>

<span data-ttu-id="559c2-146">Die folgenden Schritte sollten ihnen helfen, Probleme im Zusammenhang mit der Registerkarte **"Automatisieren"** zu beheben, die nicht in Excel im Web angezeigt wird.</span><span class="sxs-lookup"><span data-stu-id="559c2-146">The following steps should help troubleshoot any problems related to the **Automate** tab not appearing in Excel on the web.</span></span>

1. <span data-ttu-id="559c2-147">[Stellen Sie sicher, dass Ihre Microsoft 365-Lizenz Office Skripts enthält.](../overview/excel.md#requirements)</span><span class="sxs-lookup"><span data-stu-id="559c2-147">[Make sure your Microsoft 365 license includes Office Scripts](../overview/excel.md#requirements).</span></span>
1. <span data-ttu-id="559c2-148">[Überprüfen Sie, ob Ihr Browser unterstützt wird.](platform-limits.md#browser-support)</span><span class="sxs-lookup"><span data-stu-id="559c2-148">[Check that your browser is supported](platform-limits.md#browser-support).</span></span>
1. <span data-ttu-id="559c2-149">[Stellen Sie sicher, dass Cookies von Drittanbietern aktiviert sind.](platform-limits.md#third-party-cookies)</span><span class="sxs-lookup"><span data-stu-id="559c2-149">[Ensure third-party cookies are enabled](platform-limits.md#third-party-cookies).</span></span>
1. <span data-ttu-id="559c2-150">[Stellen Sie sicher, dass Ihr Administrator Office Skripts im Microsoft 365 Admin Center nicht deaktiviert hat.](/microsoft-365/admin/manage/manage-office-scripts-settings)</span><span class="sxs-lookup"><span data-stu-id="559c2-150">[Ensure that your admin has not disabled Office Scripts in the Microsoft 365 admin center](/microsoft-365/admin/manage/manage-office-scripts-settings).</span></span>

[!INCLUDE [Teams support note](../includes/teams-support-note.md)]

## <a name="troubleshoot-scripts-in-power-automate"></a><span data-ttu-id="559c2-151">Problembehandlung bei Skripts in Power Automate</span><span class="sxs-lookup"><span data-stu-id="559c2-151">Troubleshoot scripts in Power Automate</span></span>

<span data-ttu-id="559c2-152">Informationen zum Ausführen von Skripts über Power Automate finden Sie unter [Problembehandlung bei Office Skripts,](power-automate-troubleshooting.md)die in Power Automate ausgeführt werden.</span><span class="sxs-lookup"><span data-stu-id="559c2-152">For information specific to running scripts through Power Automate, see [Troubleshoot Office Scripts running in Power Automate](power-automate-troubleshooting.md).</span></span>

## <a name="help-resources"></a><span data-ttu-id="559c2-153">Hilferessourcen</span><span class="sxs-lookup"><span data-stu-id="559c2-153">Help resources</span></span>

<span data-ttu-id="559c2-154">[Stack Overflow](https://stackoverflow.com/questions/tagged/office-scripts) ist eine Community von Entwicklern, die bei Codierungsproblemen helfen möchten.</span><span class="sxs-lookup"><span data-stu-id="559c2-154">[Stack Overflow](https://stackoverflow.com/questions/tagged/office-scripts) is a community of developers willing to help with coding problems.</span></span> <span data-ttu-id="559c2-155">Häufig können Sie die Lösung für Ihr Problem mithilfe einer schnellen Stack Overflow-Suche finden.</span><span class="sxs-lookup"><span data-stu-id="559c2-155">Often, you'll be able to find the solution to your problem through a quick Stack Overflow search.</span></span> <span data-ttu-id="559c2-156">Wenn nicht, stellen Sie Ihre Frage, und markieren Sie sie mit dem Tag "office-scripts".</span><span class="sxs-lookup"><span data-stu-id="559c2-156">If not, ask your question and tag it with the "office-scripts" tag.</span></span> <span data-ttu-id="559c2-157">Erwähnen Sie unbedingt, dass Sie ein Office *Skript* erstellen, nicht ein *Office-Add-In.*</span><span class="sxs-lookup"><span data-stu-id="559c2-157">Be sure to mention you're creating an Office *Script*, not an Office *Add-in*.</span></span>

<span data-ttu-id="559c2-158">Um eine Featureanforderung für Office Skripts zu übermitteln, veröffentlichen Sie Ihre Idee auf unserer [User Voice-Seite,](https://excel.uservoice.com/forums/274580-excel-for-the-web?category_id=143439)oder wenn die Featureanforderung bereits vorhanden ist, fügen Sie Ihre Stimme dafür hinzu.</span><span class="sxs-lookup"><span data-stu-id="559c2-158">To submit a feature request for Office Scripts, post your idea to our [User Voice page](https://excel.uservoice.com/forums/274580-excel-for-the-web?category_id=143439), or if the feature request already exists there, add your vote for it.</span></span> <span data-ttu-id="559c2-159">Stellen Sie sicher, dass Sie die Anforderung unter Excel für das Web in der Kategorie "Makros, Skripts und Add-Ins" ablegen.</span><span class="sxs-lookup"><span data-stu-id="559c2-159">Be sure to file the request under Excel for the web in the "Macros, Scripts and Add-ins" category.</span></span>

<span data-ttu-id="559c2-160">Wenn ein Problem mit dem Action Recorder oder Editor vorliegt, teilen Sie uns dies bitte mit.</span><span class="sxs-lookup"><span data-stu-id="559c2-160">If there is a problem with the Action Recorder or Editor, please let us know.</span></span> <span data-ttu-id="559c2-161">Wählen Sie im Aufgabenbereich des Code-Editors **im Menü ...** die Schaltfläche **"Feedback senden"** aus, um Probleme zu teilen.</span><span class="sxs-lookup"><span data-stu-id="559c2-161">In the Code Editor task pane's **...** menu, select the **Send feedback** button to share any issues.</span></span>

:::image type="content" source="../images/code-editor-feedback.png" alt-text="Das Code-Editor-Überlaufmenü mit der Schaltfläche &quot;Feedback senden&quot;.":::

## <a name="see-also"></a><span data-ttu-id="559c2-163">Siehe auch</span><span class="sxs-lookup"><span data-stu-id="559c2-163">See also</span></span>

- [<span data-ttu-id="559c2-164">Bewährte Methoden in Office-Skripts</span><span class="sxs-lookup"><span data-stu-id="559c2-164">Best practices in Office Scripts</span></span>](../develop/best-practices.md)
- [<span data-ttu-id="559c2-165">Plattformbeschränkungen mit Office Skripts</span><span class="sxs-lookup"><span data-stu-id="559c2-165">Platform limits with Office Scripts</span></span>](platform-limits.md)
- [<span data-ttu-id="559c2-166">Verbessern der Leistung Ihrer Office-Skripts</span><span class="sxs-lookup"><span data-stu-id="559c2-166">Improve the performance of your Office Scripts</span></span>](../develop/web-client-performance.md)
- [<span data-ttu-id="559c2-167">Problembehandlung bei Office Skripts, die in PowerAutomate ausgeführt werden</span><span class="sxs-lookup"><span data-stu-id="559c2-167">Troubleshoot Office Scripts running in PowerAutomate</span></span>](power-automate-troubleshooting.md)
- [<span data-ttu-id="559c2-168">Auswirkungen von Office-Skripts rückgängig machen</span><span class="sxs-lookup"><span data-stu-id="559c2-168">Undo the effects of Office Scripts</span></span>](undo.md)
