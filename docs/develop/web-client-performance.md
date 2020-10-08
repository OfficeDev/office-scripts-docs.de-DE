---
title: Verbessern der Leistung Ihrer Office-Skripts
description: Erstellen Sie schnellere Skripts, indem Sie die Kommunikation zwischen der Excel-Arbeitsmappe und Ihrem Skript verstehen.
ms.date: 06/15/2020
localization_priority: Normal
ms.openlocfilehash: 4d5b7c70f14e3fc598b95a6226e3ef8caf89f651
ms.sourcegitcommit: aec3c971c6640429f89b6bb99d2c95ea06725599
ms.translationtype: MT
ms.contentlocale: de-DE
ms.lasthandoff: 06/25/2020
ms.locfileid: "44878902"
---
# <a name="improve-the-performance-of-your-office-scripts"></a><span data-ttu-id="5c098-103">Verbessern der Leistung Ihrer Office-Skripts</span><span class="sxs-lookup"><span data-stu-id="5c098-103">Improve the performance of your Office Scripts</span></span>

<span data-ttu-id="5c098-104">Der Zweck von Office-Skripts besteht darin, häufig ausgeführte Aufgabenserien zu automatisieren, um Zeit zu sparen.</span><span class="sxs-lookup"><span data-stu-id="5c098-104">The purpose of Office Scripts is to automate commonly performed series of tasks to save you time.</span></span> <span data-ttu-id="5c098-105">Ein langsames Skript kann sich so fühlen, als würde es den Workflow nicht beschleunigen.</span><span class="sxs-lookup"><span data-stu-id="5c098-105">A slow script can feel like it doesn't speed up your workflow.</span></span> <span data-ttu-id="5c098-106">In den meisten Fällen ist Ihr Skript vollständig in Ordnung und wird wie erwartet ausgeführt.</span><span class="sxs-lookup"><span data-stu-id="5c098-106">Most of the time, your script will be perfectly fine and run as expected.</span></span> <span data-ttu-id="5c098-107">Es gibt jedoch einige vermeidbare Szenarien, die sich auf die Leistung auswirken können.</span><span class="sxs-lookup"><span data-stu-id="5c098-107">However, there are a few, avoidable scenarios that can affect performance.</span></span>

<span data-ttu-id="5c098-108">Der häufigste Grund für ein langsames Skript ist eine übermäßige Kommunikation mit der Arbeitsmappe.</span><span class="sxs-lookup"><span data-stu-id="5c098-108">The most common reason for a slow script is excessive communication with the workbook.</span></span> <span data-ttu-id="5c098-109">Ihr Skript wird auf dem lokalen Computer ausgeführt, während die Arbeitsmappe in der Cloud vorhanden ist.</span><span class="sxs-lookup"><span data-stu-id="5c098-109">Your script runs on your local machine, while the workbook exists in the cloud.</span></span> <span data-ttu-id="5c098-110">Zu bestimmten Zeiten synchronisiert Ihr Skript seine lokalen Daten mit dem der Arbeitsmappe.</span><span class="sxs-lookup"><span data-stu-id="5c098-110">At certain times, your script synchronizes its local data with that of the workbook.</span></span> <span data-ttu-id="5c098-111">Dies bedeutet, dass alle Schreibvorgänge (beispielsweise `workbook.addWorksheet()` ) nur dann auf die Arbeitsmappe angewendet werden, wenn diese hinter den Kulissen stattfindende Synchronisierung erfolgt.</span><span class="sxs-lookup"><span data-stu-id="5c098-111">This means that any write operations (such as `workbook.addWorksheet()`) are only applied to the workbook when this behind-the-scenes synchronization happens.</span></span> <span data-ttu-id="5c098-112">Ebenso erhalten alle Lesevorgänge (beispielsweise `myRange.getValues()` ) nur Daten aus der Arbeitsmappe für das Skript zu diesen Zeiten.</span><span class="sxs-lookup"><span data-stu-id="5c098-112">Likewise, any read operations (such as `myRange.getValues()`) only get data from the workbook for the script at those times.</span></span> <span data-ttu-id="5c098-113">In beiden Fällen ruft das Skript Informationen ab, bevor es auf die Daten zugreift.</span><span class="sxs-lookup"><span data-stu-id="5c098-113">In either case, the script fetches information before it acts on the data.</span></span> <span data-ttu-id="5c098-114">Mit dem folgenden Code wird beispielsweise die Anzahl der Zeilen im verwendeten Bereich genau protokolliert.</span><span class="sxs-lookup"><span data-stu-id="5c098-114">For example, the following code will accurately log the number of rows in the used range.</span></span>

```TypeScript
let usedRange = workbook.getActiveWorksheet().getUsedRange();
let rowCount = usedRange.getRowCount();
// The script will read the range and row count from
// the workbook before logging the information.
console.log(rowCount);
```

<span data-ttu-id="5c098-115">Office Scripts-APIs stellen sicher, dass alle Daten in der Arbeitsmappe oder im Skript korrekt sind und bei Bedarf auf dem neuesten Stand sind.</span><span class="sxs-lookup"><span data-stu-id="5c098-115">Office Scripts APIs ensure any data in the workbook or script is accurate and up-to-date when necessary.</span></span> <span data-ttu-id="5c098-116">Sie müssen sich keine Gedanken über diese Synchronisierungen machen, damit Ihr Skript ordnungsgemäß ausgeführt wird.</span><span class="sxs-lookup"><span data-stu-id="5c098-116">You don't need to worry about these synchronizations for your script to run correctly.</span></span> <span data-ttu-id="5c098-117">Das Bewusstsein für diese Kommunikation zwischen Skript und Wolke kann Ihnen jedoch helfen, nicht benötigte Netzwerk Anrufe zu vermeiden.</span><span class="sxs-lookup"><span data-stu-id="5c098-117">However, an awareness of this script-to-cloud communication can help you avoid unneeded network calls.</span></span>

## <a name="performance-optimizations"></a><span data-ttu-id="5c098-118">Leistungsoptimierungen</span><span class="sxs-lookup"><span data-stu-id="5c098-118">Performance optimizations</span></span>

<span data-ttu-id="5c098-119">Sie können einfache Techniken anwenden, um die Kommunikation mit der Cloud zu verringern.</span><span class="sxs-lookup"><span data-stu-id="5c098-119">You can apply simple techniques to help reduce the communication to the cloud.</span></span> <span data-ttu-id="5c098-120">Die folgenden Muster helfen Ihnen dabei, die Skripts zu beschleunigen.</span><span class="sxs-lookup"><span data-stu-id="5c098-120">The following patterns help speed up your scripts.</span></span>

- <span data-ttu-id="5c098-121">Lesen von Arbeitsmappendaten einmal statt wiederholt in einer Schleife.</span><span class="sxs-lookup"><span data-stu-id="5c098-121">Read workbook data once instead of repeatedly in a loop.</span></span>
- <span data-ttu-id="5c098-122">Entfernen Sie unnötige `console.log` Anweisungen.</span><span class="sxs-lookup"><span data-stu-id="5c098-122">Remove unnecessary `console.log` statements.</span></span>
- <span data-ttu-id="5c098-123">Vermeiden Sie die Verwendung von try/catch-Blöcken.</span><span class="sxs-lookup"><span data-stu-id="5c098-123">Avoid using try/catch blocks.</span></span>

### <a name="read-workbook-data-outside-of-a-loop"></a><span data-ttu-id="5c098-124">Arbeitsmappen-Daten außerhalb einer Schleife lesen</span><span class="sxs-lookup"><span data-stu-id="5c098-124">Read workbook data outside of a loop</span></span>

<span data-ttu-id="5c098-125">Jede Methode, die Daten aus der Arbeitsmappe abruft, kann einen Netzwerkaufruf auslösen.</span><span class="sxs-lookup"><span data-stu-id="5c098-125">Any method that gets data from the workbook can trigger a network call.</span></span> <span data-ttu-id="5c098-126">Anstatt wiederholt den gleichen Aufruf durchführen, sollten Sie Daten lokal speichern, wann immer möglich.</span><span class="sxs-lookup"><span data-stu-id="5c098-126">Rather than repeatedly making the same call, you should save data locally whenever possible.</span></span> <span data-ttu-id="5c098-127">Dies gilt insbesondere beim Umgang mit Schleifen.</span><span class="sxs-lookup"><span data-stu-id="5c098-127">This is especially true when dealing with loops.</span></span>

<span data-ttu-id="5c098-128">Verwenden Sie ein Skript, um die Anzahl der negativen Zahlen im verwendeten Bereich eines Arbeitsblatts abzurufen.</span><span class="sxs-lookup"><span data-stu-id="5c098-128">Consider a script to get the count of negative numbers in the used range of a worksheet.</span></span> <span data-ttu-id="5c098-129">Das Skript muss jede Zelle im verwendeten Bereich durchlaufen.</span><span class="sxs-lookup"><span data-stu-id="5c098-129">The script needs to iterate over every cell in the used range.</span></span> <span data-ttu-id="5c098-130">Dazu benötigt er den Bereich, die Anzahl der Zeilen und die Anzahl der Spalten.</span><span class="sxs-lookup"><span data-stu-id="5c098-130">To do that, it needs the range, the number of rows, and the number of columns.</span></span> <span data-ttu-id="5c098-131">Sie sollten diese als lokale Variablen speichern, bevor Sie mit der Schleife beginnen.</span><span class="sxs-lookup"><span data-stu-id="5c098-131">You should store those as local variables before starting the loop.</span></span> <span data-ttu-id="5c098-132">Andernfalls wird durch jede Iteration der Schleife eine Rückgabe an die Arbeitsmappe erzwungen.</span><span class="sxs-lookup"><span data-stu-id="5c098-132">Otherwise, each iteration of the loop will force a return to the workbook.</span></span>

```TypeScript
/**
 * This script provides the count of negative numbers that are present
 * in the used range of the current worksheet.
 */
function main(workbook: ExcelScript.Workbook) {
  // Get the working range.
  let usedRange = workbook.getActiveWorksheet().getUsedRange();

  // Save the values locally to avoid repeatedly asking the workbook.
  let usedRangeValues = usedRange.getValues();

  // Start the negative number counter.
  let negativeCount = 0;

  // Iterate over the entire range looking for negative numbers.
  for (let i = 0; i < usedRangeValues.length; i++) {
    for (let j = 0; j < usedRangeValues[i].length; j++) {
      if (usedRangeValues[i][j] < 0) {
        negativeCount++;
      }
    }
  }

  // Log the negative number count to the console.
  console.log(negativeCount);
}
```

> [!NOTE]
> <span data-ttu-id="5c098-133">Versuchen Sie als Experiment `usedRangeValues` in der Schleife mit zu ersetzen `usedRange.getValues()` .</span><span class="sxs-lookup"><span data-stu-id="5c098-133">As an experiment, try replacing `usedRangeValues` in the loop with `usedRange.getValues()`.</span></span> <span data-ttu-id="5c098-134">Sie können feststellen, dass das Skript beim Umgang mit großen Bereichen wesentlich länger ausgeführt wird.</span><span class="sxs-lookup"><span data-stu-id="5c098-134">You may notice the script takes considerably longer to run when dealing with large ranges.</span></span>

### <a name="remove-unnecessary-consolelog-statements"></a><span data-ttu-id="5c098-135">Entfernen unnötiger `console.log` Anweisungen</span><span class="sxs-lookup"><span data-stu-id="5c098-135">Remove unnecessary `console.log` statements</span></span>

<span data-ttu-id="5c098-136">Die Konsolenprotokollierung ist ein wichtiges Tool zum [Debuggen von Skripts](../testing/troubleshooting.md).</span><span class="sxs-lookup"><span data-stu-id="5c098-136">Console logging is a vital tool for [debugging your scripts](../testing/troubleshooting.md).</span></span> <span data-ttu-id="5c098-137">Es zwingt das Skript jedoch zur Synchronisierung mit der Arbeitsmappe, um sicherzustellen, dass die protokollierten Informationen auf dem neuesten Stand sind.</span><span class="sxs-lookup"><span data-stu-id="5c098-137">However, it does force the script to synchronize with the workbook to ensure the logged information is up-to-date.</span></span> <span data-ttu-id="5c098-138">Entfernen Sie nicht benötigte Protokollierungs Anweisungen (beispielsweise die zum Testen verwendeten), bevor Sie Ihr Skript freigeben.</span><span class="sxs-lookup"><span data-stu-id="5c098-138">Consider removing unnecessary logging statements (such as those used for testing) before sharing your script.</span></span> <span data-ttu-id="5c098-139">Dies führt in der Regel nicht zu einem spürbaren Leistungsproblem, es sei denn, die `console.log()` Anweisung befindet sich in einer Schleife.</span><span class="sxs-lookup"><span data-stu-id="5c098-139">This typically won't cause a noticeable performance issue, unless the `console.log()` statement is in a loop.</span></span>

### <a name="avoid-using-trycatch-blocks"></a><span data-ttu-id="5c098-140">Vermeiden der Verwendung von try/catch-Blöcken</span><span class="sxs-lookup"><span data-stu-id="5c098-140">Avoid using try/catch blocks</span></span>

<span data-ttu-id="5c098-141">Es wird nicht empfohlen, [ `try` / `catch` Blöcke](https://developer.mozilla.org/docs/Web/JavaScript/Reference/Statements/try...catch) als Teil der erwarteten Ablaufsteuerung eines Skripts zu verwenden.</span><span class="sxs-lookup"><span data-stu-id="5c098-141">We don't recommend using [`try`/`catch` blocks](https://developer.mozilla.org/docs/Web/JavaScript/Reference/Statements/try...catch) as part of a script's expected control flow.</span></span> <span data-ttu-id="5c098-142">Die meisten Fehler können vermieden werden, indem die von der Arbeitsmappe zurückgegebenen Objekte überprüft werden.</span><span class="sxs-lookup"><span data-stu-id="5c098-142">Most errors can be avoided by checking objects returned from the workbook.</span></span> <span data-ttu-id="5c098-143">Das folgende Skript überprüft beispielsweise, ob die von der Arbeitsmappe zurückgegebene Tabelle vorhanden ist, bevor versucht wird, eine Zeile hinzuzufügen.</span><span class="sxs-lookup"><span data-stu-id="5c098-143">For example, the following script checks that the table returned by the workbook exists before trying to add a row.</span></span>

```TypeScript
/**
 * This script adds a row to "MyTable", if that table is present.
 */
function main(workbook: ExcelScript.Workbook) {
  let table = workbook.getTable("MyTable");

  // Check if the table exists.
  if (table) {
    // Add the row.
    table.addRow(-1, ["2012", "Yes", "Maybe"]);
  } else {
    // Report the missing table.
    console.log("MyTable not found.");
  }
}
```

## <a name="case-by-case-help"></a><span data-ttu-id="5c098-144">Fall-für-Fall-Hilfe</span><span class="sxs-lookup"><span data-stu-id="5c098-144">Case-by-case help</span></span>

<span data-ttu-id="5c098-145">Wenn die Office-Skriptplattform erweitert wird, um mit [Power Automation](https://flow.microsoft.com/), [Adaptive Cards](https://docs.microsoft.com/adaptive-cards)und anderen produktübergreifenden Features zu arbeiten, werden die Details der Skript-Arbeitsmappen-Kommunikation komplizierter.</span><span class="sxs-lookup"><span data-stu-id="5c098-145">As the Office Scripts platform expands to work with [Power Automate](https://flow.microsoft.com/), [Adaptive Cards](https://docs.microsoft.com/adaptive-cards), and other cross-product features, the details of the script-workbook communication become more intricate.</span></span> <span data-ttu-id="5c098-146">Wenn Sie Hilfe beim Beschleunigen des Skripts benötigen, erreichen Sie den [Stapelüberlauf](https://stackoverflow.com/questions/tagged/office-scripts).</span><span class="sxs-lookup"><span data-stu-id="5c098-146">If you need help making your script run faster, please reach out through [Stack Overflow](https://stackoverflow.com/questions/tagged/office-scripts).</span></span> <span data-ttu-id="5c098-147">Achten Sie darauf, Ihre Frage mit "Office-Skripts" zu versehen, damit Experten Sie finden und helfen können.</span><span class="sxs-lookup"><span data-stu-id="5c098-147">Be sure to tag your question with "office-scripts" so experts can find it and help.</span></span>

## <a name="see-also"></a><span data-ttu-id="5c098-148">Siehe auch</span><span class="sxs-lookup"><span data-stu-id="5c098-148">See also</span></span>

- [<span data-ttu-id="5c098-149">Grundlagen der Skripterstellung für Office-Skripts in Excel im Web</span><span class="sxs-lookup"><span data-stu-id="5c098-149">Scripting fundamentals for Office Scripts in Excel on the web</span></span>](scripting-fundamentals.md)
- [<span data-ttu-id="5c098-150">MDN-Webdocs: Loops and Iteration</span><span class="sxs-lookup"><span data-stu-id="5c098-150">MDN web docs: Loops and iteration</span></span>](https://developer.mozilla.org/docs/Web/JavaScript/Guide/Loops_and_iteration)
