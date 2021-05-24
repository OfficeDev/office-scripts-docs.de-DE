---
title: Verbessern der Leistung Ihrer Office Skripts
description: Erstellen Sie schnellere Skripts, indem Sie die Kommunikation zwischen Excel Arbeitsmappe und Ihrem Skript verstehen.
ms.date: 05/17/2021
localization_priority: Normal
ms.openlocfilehash: 512e2108cb81cf9ac8ae98980951d5d01b3d2de9
ms.sourcegitcommit: 4687693f02fc90a57ba30c461f35046e02e6f5fb
ms.translationtype: MT
ms.contentlocale: de-DE
ms.lasthandoff: 05/19/2021
ms.locfileid: "52544991"
---
# <a name="improve-the-performance-of-your-office-scripts"></a><span data-ttu-id="f1542-103">Verbessern der Leistung Ihrer Office Skripts</span><span class="sxs-lookup"><span data-stu-id="f1542-103">Improve the performance of your Office Scripts</span></span>

<span data-ttu-id="f1542-104">Der Zweck Office Skripts besteht in der Automatisierung häufig ausgeführter Aufgabenreihen, um Zeit zu sparen.</span><span class="sxs-lookup"><span data-stu-id="f1542-104">The purpose of Office Scripts is to automate commonly performed series of tasks to save you time.</span></span> <span data-ttu-id="f1542-105">Ein langsames Skript kann den Workflow nicht beschleunigen.</span><span class="sxs-lookup"><span data-stu-id="f1542-105">A slow script can feel like it doesn't speed up your workflow.</span></span> <span data-ttu-id="f1542-106">In der meisten Zeit ist Ihr Skript vollkommen in Ordnung und wird wie erwartet ausgeführt.</span><span class="sxs-lookup"><span data-stu-id="f1542-106">Most of the time, your script will be perfectly fine and run as expected.</span></span> <span data-ttu-id="f1542-107">Es gibt jedoch einige vermeidbare Szenarien, die sich auf die Leistung auswirken können.</span><span class="sxs-lookup"><span data-stu-id="f1542-107">However, there are a few, avoidable scenarios that can affect performance.</span></span>

<span data-ttu-id="f1542-108">Der häufigste Grund für ein langsames Skript ist die übermäßige Kommunikation mit der Arbeitsmappe.</span><span class="sxs-lookup"><span data-stu-id="f1542-108">The most common reason for a slow script is excessive communication with the workbook.</span></span> <span data-ttu-id="f1542-109">Das Skript wird auf dem lokalen Computer ausgeführt, während die Arbeitsmappe in der Cloud vorhanden ist.</span><span class="sxs-lookup"><span data-stu-id="f1542-109">Your script runs on your local machine, while the workbook exists in the cloud.</span></span> <span data-ttu-id="f1542-110">Zu bestimmten Zeiten synchronisiert ihr Skript seine lokalen Daten mit den daten der Arbeitsmappe.</span><span class="sxs-lookup"><span data-stu-id="f1542-110">At certain times, your script synchronizes its local data with that of the workbook.</span></span> <span data-ttu-id="f1542-111">Dies bedeutet, dass schreibvorgänge (z. B. ) nur auf die Arbeitsmappe angewendet werden, wenn diese Synchronisierung im Hintergrund `workbook.addWorksheet()` erfolgt.</span><span class="sxs-lookup"><span data-stu-id="f1542-111">This means that any write operations (such as `workbook.addWorksheet()`) are only applied to the workbook when this behind-the-scenes synchronization happens.</span></span> <span data-ttu-id="f1542-112">Ebenso erhalten lesevorgänge (z. B. ) nur daten aus der `myRange.getValues()` Arbeitsmappe für das Skript zu diesen Zeiten.</span><span class="sxs-lookup"><span data-stu-id="f1542-112">Likewise, any read operations (such as `myRange.getValues()`) only get data from the workbook for the script at those times.</span></span> <span data-ttu-id="f1542-113">In beiden Fällen ruft das Skript Informationen ab, bevor es auf die Daten wirkt.</span><span class="sxs-lookup"><span data-stu-id="f1542-113">In either case, the script fetches information before it acts on the data.</span></span> <span data-ttu-id="f1542-114">Mit dem folgenden Code wird beispielsweise die Anzahl der Zeilen im verwendeten Bereich genau protokolliert.</span><span class="sxs-lookup"><span data-stu-id="f1542-114">For example, the following code will accurately log the number of rows in the used range.</span></span>

```TypeScript
let usedRange = workbook.getActiveWorksheet().getUsedRange();
let rowCount = usedRange.getRowCount();
// The script will read the range and row count from
// the workbook before logging the information.
console.log(rowCount);
```

<span data-ttu-id="f1542-115">Office Skript-APIs stellen sicher, dass alle Daten in der Arbeitsmappe oder dem Skript bei Bedarf korrekt und aktuell sind.</span><span class="sxs-lookup"><span data-stu-id="f1542-115">Office Scripts APIs ensure any data in the workbook or script is accurate and up-to-date when necessary.</span></span> <span data-ttu-id="f1542-116">Sie müssen sich keine Gedanken über diese Synchronisierungen machen, damit Ihr Skript ordnungsgemäß ausgeführt wird.</span><span class="sxs-lookup"><span data-stu-id="f1542-116">You don't need to worry about these synchronizations for your script to run correctly.</span></span> <span data-ttu-id="f1542-117">Die Kenntnis dieser Skript-zu-Cloud-Kommunikation kann Ihnen jedoch helfen, unnötige Netzwerkanrufe zu vermeiden.</span><span class="sxs-lookup"><span data-stu-id="f1542-117">However, an awareness of this script-to-cloud communication can help you avoid unneeded network calls.</span></span>

## <a name="performance-optimizations"></a><span data-ttu-id="f1542-118">Leistungsoptimierungen</span><span class="sxs-lookup"><span data-stu-id="f1542-118">Performance optimizations</span></span>

<span data-ttu-id="f1542-119">Sie können einfache Techniken anwenden, um die Kommunikation mit der Cloud zu reduzieren.</span><span class="sxs-lookup"><span data-stu-id="f1542-119">You can apply simple techniques to help reduce the communication to the cloud.</span></span> <span data-ttu-id="f1542-120">Die folgenden Muster beschleunigen Ihre Skripts.</span><span class="sxs-lookup"><span data-stu-id="f1542-120">The following patterns help speed up your scripts.</span></span>

- <span data-ttu-id="f1542-121">Lesen Von Arbeitsmappendaten einmal statt wiederholt in einer Schleife.</span><span class="sxs-lookup"><span data-stu-id="f1542-121">Read workbook data once instead of repeatedly in a loop.</span></span>
- <span data-ttu-id="f1542-122">Entfernen Sie unnötige `console.log` Anweisungen.</span><span class="sxs-lookup"><span data-stu-id="f1542-122">Remove unnecessary `console.log` statements.</span></span>
- <span data-ttu-id="f1542-123">Vermeiden Sie die Verwendung von try/catch-Blöcken.</span><span class="sxs-lookup"><span data-stu-id="f1542-123">Avoid using try/catch blocks.</span></span>

### <a name="read-workbook-data-outside-of-a-loop"></a><span data-ttu-id="f1542-124">Lesen von Arbeitsmappendaten außerhalb einer Schleife</span><span class="sxs-lookup"><span data-stu-id="f1542-124">Read workbook data outside of a loop</span></span>

<span data-ttu-id="f1542-125">Jede Methode, die Daten aus der Arbeitsmappe abruft, kann einen Netzwerkaufruf auslösen.</span><span class="sxs-lookup"><span data-stu-id="f1542-125">Any method that gets data from the workbook can trigger a network call.</span></span> <span data-ttu-id="f1542-126">Anstatt immer wieder denselben Anruf zu machen, sollten Sie Daten nach Möglichkeit lokal speichern.</span><span class="sxs-lookup"><span data-stu-id="f1542-126">Rather than repeatedly making the same call, you should save data locally whenever possible.</span></span> <span data-ttu-id="f1542-127">Dies gilt insbesondere beim Umgang mit Schleifen.</span><span class="sxs-lookup"><span data-stu-id="f1542-127">This is especially true when dealing with loops.</span></span>

<span data-ttu-id="f1542-128">Erwägen Sie ein Skript, um die Anzahl negativer Zahlen im verwendeten Bereich eines Arbeitsblatts zu erhalten.</span><span class="sxs-lookup"><span data-stu-id="f1542-128">Consider a script to get the count of negative numbers in the used range of a worksheet.</span></span> <span data-ttu-id="f1542-129">Das Skript muss über jede Zelle im verwendeten Bereich iterieren.</span><span class="sxs-lookup"><span data-stu-id="f1542-129">The script needs to iterate over every cell in the used range.</span></span> <span data-ttu-id="f1542-130">Dazu benötigt sie den Bereich, die Anzahl der Zeilen und die Anzahl der Spalten.</span><span class="sxs-lookup"><span data-stu-id="f1542-130">To do that, it needs the range, the number of rows, and the number of columns.</span></span> <span data-ttu-id="f1542-131">Sie sollten diese als lokale Variablen speichern, bevor Sie die Schleife starten.</span><span class="sxs-lookup"><span data-stu-id="f1542-131">You should store those as local variables before starting the loop.</span></span> <span data-ttu-id="f1542-132">Andernfalls wird bei jeder Iteration der Schleife eine Rückgabe zur Arbeitsmappe erzwingen.</span><span class="sxs-lookup"><span data-stu-id="f1542-132">Otherwise, each iteration of the loop will force a return to the workbook.</span></span>

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
> <span data-ttu-id="f1542-133">Versuchen Sie als Experiment, `usedRangeValues` in der Schleife durch zu `usedRange.getValues()` ersetzen.</span><span class="sxs-lookup"><span data-stu-id="f1542-133">As an experiment, try replacing `usedRangeValues` in the loop with `usedRange.getValues()`.</span></span> <span data-ttu-id="f1542-134">Möglicherweise dauert die Ausführung des Skripts erheblich länger, wenn große Bereiche verwendet werden.</span><span class="sxs-lookup"><span data-stu-id="f1542-134">You may notice the script takes considerably longer to run when dealing with large ranges.</span></span>

### <a name="avoid-using-trycatch-blocks-in-or-surrounding-loops"></a><span data-ttu-id="f1542-135">Vermeiden der `try...catch` Verwendung von Blöcken in oder umgebenden Schleifen</span><span class="sxs-lookup"><span data-stu-id="f1542-135">Avoid using `try...catch` blocks in or surrounding loops</span></span>

<span data-ttu-id="f1542-136">Es wird nicht empfohlen, Anweisungen [`try...catch`](https://developer.mozilla.org/docs/Web/JavaScript/Reference/Statements/try...catch) entweder in Schleifen oder umgebenden Schleifen zu verwenden.</span><span class="sxs-lookup"><span data-stu-id="f1542-136">We don't recommend using [`try...catch`](https://developer.mozilla.org/docs/Web/JavaScript/Reference/Statements/try...catch) statements either in loops or surrounding loops.</span></span> <span data-ttu-id="f1542-137">Aus demselben Grund sollten Sie das Lesen von Daten in einer Schleife vermeiden: Jede Iteration erzwingt die Synchronisierung des Skripts mit der Arbeitsmappe, um sicherzustellen, dass kein Fehler ausgelöst wurde.</span><span class="sxs-lookup"><span data-stu-id="f1542-137">This is for the same reason you should avoid reading data in a loop: each iteration forces the script to synchronize with the workbook to make sure no error has been thrown.</span></span> <span data-ttu-id="f1542-138">Die meisten Fehler können vermieden werden, indem Von der Arbeitsmappe zurückgegebene Objekte überprüft werden.</span><span class="sxs-lookup"><span data-stu-id="f1542-138">Most errors can be avoided by checking objects returned from the workbook.</span></span> <span data-ttu-id="f1542-139">Das folgende Skript überprüft beispielsweise, ob die von der Arbeitsmappe zurückgegebene Tabelle vorhanden ist, bevor versucht wird, eine Zeile hinzuzufügen.</span><span class="sxs-lookup"><span data-stu-id="f1542-139">For example, the following script checks that the table returned by the workbook exists before trying to add a row.</span></span>

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

### <a name="remove-unnecessary-consolelog-statements"></a><span data-ttu-id="f1542-140">Entfernen unnötiger `console.log` Anweisungen</span><span class="sxs-lookup"><span data-stu-id="f1542-140">Remove unnecessary `console.log` statements</span></span>

<span data-ttu-id="f1542-141">Die Konsolenprotokollierung ist ein wichtiges Tool zum [Debuggen Ihrer Skripts.](../testing/troubleshooting.md)</span><span class="sxs-lookup"><span data-stu-id="f1542-141">Console logging is a vital tool for [debugging your scripts](../testing/troubleshooting.md).</span></span> <span data-ttu-id="f1542-142">Das Skript muss jedoch mit der Arbeitsmappe synchronisiert werden, um sicherzustellen, dass die protokollierten Informationen auf dem neuesten Stand sind.</span><span class="sxs-lookup"><span data-stu-id="f1542-142">However, it does force the script to synchronize with the workbook to ensure the logged information is up-to-date.</span></span> <span data-ttu-id="f1542-143">Ziehen Sie das Entfernen unnötiger Protokollierungsanweisungen (z. B. zu Testzwecken) in Betracht, bevor Sie Ihr Skript freigeben.</span><span class="sxs-lookup"><span data-stu-id="f1542-143">Consider removing unnecessary logging statements (such as those used for testing) before sharing your script.</span></span> <span data-ttu-id="f1542-144">Dies verursacht normalerweise kein spürbares Leistungsproblem, es sei denn, die `console.log()` Anweisung befindet sich in einer Schleife.</span><span class="sxs-lookup"><span data-stu-id="f1542-144">This typically won't cause a noticeable performance issue, unless the `console.log()` statement is in a loop.</span></span>

## <a name="case-by-case-help"></a><span data-ttu-id="f1542-145">Fall-für-Fall-Hilfe</span><span class="sxs-lookup"><span data-stu-id="f1542-145">Case-by-case help</span></span>

<span data-ttu-id="f1542-146">Wenn die Office-Skriptplattform erweitert wird, um mit [Power Automate,](https://flow.microsoft.com/) [adaptiven](/adaptive-cards)Karten und anderen produktübergreifenden Features zu arbeiten, werden die Details der Skriptarbeitsmappenkommunikation komplizierter.</span><span class="sxs-lookup"><span data-stu-id="f1542-146">As the Office Scripts platform expands to work with [Power Automate](https://flow.microsoft.com/), [Adaptive Cards](/adaptive-cards), and other cross-product features, the details of the script-workbook communication become more intricate.</span></span> <span data-ttu-id="f1542-147">Wenn Sie Hilfe benötigen, damit Ihr Skript schneller ausgeführt wird, wenden Sie sich bitte an [Microsoft Q&A](/answers/topics/office-scripts-dev.html).</span><span class="sxs-lookup"><span data-stu-id="f1542-147">If you need help making your script run faster, please reach out through [Microsoft Q&A](/answers/topics/office-scripts-dev.html).</span></span> <span data-ttu-id="f1542-148">Achten Sie darauf, Ihre Frage mit "office-scripts-dev" zu kennzeichnen, damit Experten sie finden und helfen können.</span><span class="sxs-lookup"><span data-stu-id="f1542-148">Be sure to tag your question with "office-scripts-dev" so experts can find it and help.</span></span>

## <a name="see-also"></a><span data-ttu-id="f1542-149">Siehe auch</span><span class="sxs-lookup"><span data-stu-id="f1542-149">See also</span></span>

- [<span data-ttu-id="f1542-150">Grundlagen der Skripterstellung für Office-Skripts in Excel im Web</span><span class="sxs-lookup"><span data-stu-id="f1542-150">Scripting fundamentals for Office Scripts in Excel on the web</span></span>](scripting-fundamentals.md)
- [<span data-ttu-id="f1542-151">MDN-Web-Dokumente: Schleifen und Iteration</span><span class="sxs-lookup"><span data-stu-id="f1542-151">MDN web docs: Loops and iteration</span></span>](https://developer.mozilla.org/docs/Web/JavaScript/Guide/Loops_and_iteration)
