---
title: Verbessern der Leistung Ihrer Office-Skripts
description: Erstellen Sie schnellere Skripts, indem Sie die Kommunikation zwischen der Arbeitsmappe in Excel und Ihrem Skript verstehen.
ms.date: 06/15/2020
localization_priority: Normal
ms.openlocfilehash: ce50a6fd7ad02ddcd2dd304be8b4dd8fa3d0acf3
ms.sourcegitcommit: 7580dcb8f2f97974c2a9cce25ea30d6526730e28
ms.translationtype: MT
ms.contentlocale: de-DE
ms.lasthandoff: 01/14/2021
ms.locfileid: "49867870"
---
# <a name="improve-the-performance-of-your-office-scripts"></a><span data-ttu-id="1357f-103">Verbessern der Leistung Ihrer Office-Skripts</span><span class="sxs-lookup"><span data-stu-id="1357f-103">Improve the performance of your Office Scripts</span></span>

<span data-ttu-id="1357f-104">Der Zweck von Office Scripts besteht in der Automatisierung häufig ausgeführter Aufgabenreihen, um Zeit zu sparen.</span><span class="sxs-lookup"><span data-stu-id="1357f-104">The purpose of Office Scripts is to automate commonly performed series of tasks to save you time.</span></span> <span data-ttu-id="1357f-105">Ein langsames Skript kann den Workflow nicht beschleunigen.</span><span class="sxs-lookup"><span data-stu-id="1357f-105">A slow script can feel like it doesn't speed up your workflow.</span></span> <span data-ttu-id="1357f-106">In den meisten Zeiten ist Ihr Skript völlig in Ordnung und wird wie erwartet ausgeführt.</span><span class="sxs-lookup"><span data-stu-id="1357f-106">Most of the time, your script will be perfectly fine and run as expected.</span></span> <span data-ttu-id="1357f-107">Es gibt jedoch einige vermeidbare Szenarien, die sich auf die Leistung auswirken können.</span><span class="sxs-lookup"><span data-stu-id="1357f-107">However, there are a few, avoidable scenarios that can affect performance.</span></span>

<span data-ttu-id="1357f-108">Der häufigste Grund für ein langsames Skript ist die übermäßige Kommunikation mit der Arbeitsmappe.</span><span class="sxs-lookup"><span data-stu-id="1357f-108">The most common reason for a slow script is excessive communication with the workbook.</span></span> <span data-ttu-id="1357f-109">Ihr Skript wird auf Dem lokalen Computer ausgeführt, während die Arbeitsmappe in der Cloud vorhanden ist.</span><span class="sxs-lookup"><span data-stu-id="1357f-109">Your script runs on your local machine, while the workbook exists in the cloud.</span></span> <span data-ttu-id="1357f-110">Zu bestimmten Zeiten synchronisiert das Skript seine lokalen Daten mit dem der Arbeitsmappe.</span><span class="sxs-lookup"><span data-stu-id="1357f-110">At certain times, your script synchronizes its local data with that of the workbook.</span></span> <span data-ttu-id="1357f-111">Dies bedeutet, dass alle Schreibvorgänge (z. B. ) nur auf die Arbeitsmappe angewendet werden, wenn diese `workbook.addWorksheet()` Hintergrundsynchronisierung erfolgt.</span><span class="sxs-lookup"><span data-stu-id="1357f-111">This means that any write operations (such as `workbook.addWorksheet()`) are only applied to the workbook when this behind-the-scenes synchronization happens.</span></span> <span data-ttu-id="1357f-112">Ebenso erhalten alle Lesevorgänge (z. B. ) nur zu diesen Zeiten Daten aus der Arbeitsmappe für `myRange.getValues()` das Skript.</span><span class="sxs-lookup"><span data-stu-id="1357f-112">Likewise, any read operations (such as `myRange.getValues()`) only get data from the workbook for the script at those times.</span></span> <span data-ttu-id="1357f-113">In beiden Fällen ruft das Skript Informationen ab, bevor es für die Daten agiert.</span><span class="sxs-lookup"><span data-stu-id="1357f-113">In either case, the script fetches information before it acts on the data.</span></span> <span data-ttu-id="1357f-114">Mit dem folgenden Code wird beispielsweise die Anzahl der Zeilen im verwendeten Bereich genau protokolliert.</span><span class="sxs-lookup"><span data-stu-id="1357f-114">For example, the following code will accurately log the number of rows in the used range.</span></span>

```TypeScript
let usedRange = workbook.getActiveWorksheet().getUsedRange();
let rowCount = usedRange.getRowCount();
// The script will read the range and row count from
// the workbook before logging the information.
console.log(rowCount);
```

<span data-ttu-id="1357f-115">Office-Skript-APIs stellen sicher, dass alle Daten in der Arbeitsmappe oder dem Skript bei Bedarf korrekt und aktuell sind.</span><span class="sxs-lookup"><span data-stu-id="1357f-115">Office Scripts APIs ensure any data in the workbook or script is accurate and up-to-date when necessary.</span></span> <span data-ttu-id="1357f-116">Sie müssen sich keine Gedanken über diese Synchronisierungen machen, damit Ihr Skript ordnungsgemäß ausgeführt wird.</span><span class="sxs-lookup"><span data-stu-id="1357f-116">You don't need to worry about these synchronizations for your script to run correctly.</span></span> <span data-ttu-id="1357f-117">Die Kenntnis dieser Skript-zu-Cloud-Kommunikation kann Ihnen jedoch helfen, unnötige Netzwerkanrufe zu vermeiden.</span><span class="sxs-lookup"><span data-stu-id="1357f-117">However, an awareness of this script-to-cloud communication can help you avoid unneeded network calls.</span></span>

## <a name="performance-optimizations"></a><span data-ttu-id="1357f-118">Leistungsoptimierungen</span><span class="sxs-lookup"><span data-stu-id="1357f-118">Performance optimizations</span></span>

<span data-ttu-id="1357f-119">Sie können einfache Techniken anwenden, um die Kommunikation mit der Cloud zu reduzieren.</span><span class="sxs-lookup"><span data-stu-id="1357f-119">You can apply simple techniques to help reduce the communication to the cloud.</span></span> <span data-ttu-id="1357f-120">Die folgenden Muster beschleunigen Ihre Skripts.</span><span class="sxs-lookup"><span data-stu-id="1357f-120">The following patterns help speed up your scripts.</span></span>

- <span data-ttu-id="1357f-121">Lesen Sie Arbeitsmappendaten einmal, anstatt wiederholt in einer Schleife.</span><span class="sxs-lookup"><span data-stu-id="1357f-121">Read workbook data once instead of repeatedly in a loop.</span></span>
- <span data-ttu-id="1357f-122">Entfernen Sie unnötige `console.log` Anweisungen.</span><span class="sxs-lookup"><span data-stu-id="1357f-122">Remove unnecessary `console.log` statements.</span></span>
- <span data-ttu-id="1357f-123">Vermeiden Sie try/catch-Blöcke.</span><span class="sxs-lookup"><span data-stu-id="1357f-123">Avoid using try/catch blocks.</span></span>

### <a name="read-workbook-data-outside-of-a-loop"></a><span data-ttu-id="1357f-124">Lesen von Arbeitsmappendaten außerhalb einer Schleife</span><span class="sxs-lookup"><span data-stu-id="1357f-124">Read workbook data outside of a loop</span></span>

<span data-ttu-id="1357f-125">Jede Methode, die Daten aus der Arbeitsmappe abruft, kann einen Netzwerkaufruf auslösen.</span><span class="sxs-lookup"><span data-stu-id="1357f-125">Any method that gets data from the workbook can trigger a network call.</span></span> <span data-ttu-id="1357f-126">Anstatt denselben Aufruf wiederholt zu machen, sollten Sie Daten nach Möglichkeit lokal speichern.</span><span class="sxs-lookup"><span data-stu-id="1357f-126">Rather than repeatedly making the same call, you should save data locally whenever possible.</span></span> <span data-ttu-id="1357f-127">Dies gilt insbesondere für Schleifen.</span><span class="sxs-lookup"><span data-stu-id="1357f-127">This is especially true when dealing with loops.</span></span>

<span data-ttu-id="1357f-128">Stellen Sie sich ein Skript vor, um die Anzahl negativer Zahlen im verwendeten Bereich eines Arbeitsblatts zu erhalten.</span><span class="sxs-lookup"><span data-stu-id="1357f-128">Consider a script to get the count of negative numbers in the used range of a worksheet.</span></span> <span data-ttu-id="1357f-129">Das Skript muss jede Zelle im verwendeten Bereich durch iterieren.</span><span class="sxs-lookup"><span data-stu-id="1357f-129">The script needs to iterate over every cell in the used range.</span></span> <span data-ttu-id="1357f-130">Dazu sind der Bereich, die Anzahl der Zeilen und die Anzahl der Spalten benötigt.</span><span class="sxs-lookup"><span data-stu-id="1357f-130">To do that, it needs the range, the number of rows, and the number of columns.</span></span> <span data-ttu-id="1357f-131">Sie sollten diese als lokale Variablen speichern, bevor Sie die Schleife starten.</span><span class="sxs-lookup"><span data-stu-id="1357f-131">You should store those as local variables before starting the loop.</span></span> <span data-ttu-id="1357f-132">Andernfalls wird bei jeder Iteration der Schleife eine Rückgabe an die Arbeitsmappe erzwingen.</span><span class="sxs-lookup"><span data-stu-id="1357f-132">Otherwise, each iteration of the loop will force a return to the workbook.</span></span>

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
> <span data-ttu-id="1357f-133">Versuchen Sie als Experiment, `usedRangeValues` in der Schleife durch zu `usedRange.getValues()` ersetzen.</span><span class="sxs-lookup"><span data-stu-id="1357f-133">As an experiment, try replacing `usedRangeValues` in the loop with `usedRange.getValues()`.</span></span> <span data-ttu-id="1357f-134">Sie werden feststellen, dass die Ausführung des Skripts bei großen Bereichen erheblich länger dauert.</span><span class="sxs-lookup"><span data-stu-id="1357f-134">You may notice the script takes considerably longer to run when dealing with large ranges.</span></span>

### <a name="remove-unnecessary-consolelog-statements"></a><span data-ttu-id="1357f-135">Entfernen unnötiger `console.log` Anweisungen</span><span class="sxs-lookup"><span data-stu-id="1357f-135">Remove unnecessary `console.log` statements</span></span>

<span data-ttu-id="1357f-136">Die Konsolenprotokollierung ist ein wichtiges Tool zum [Debuggen Ihrer Skripts.](../testing/troubleshooting.md)</span><span class="sxs-lookup"><span data-stu-id="1357f-136">Console logging is a vital tool for [debugging your scripts](../testing/troubleshooting.md).</span></span> <span data-ttu-id="1357f-137">Es wird jedoch die Synchronisierung des Skripts mit der Arbeitsmappe erzwingen, um sicherzustellen, dass die protokollierten Informationen auf dem neuesten Stand sind.</span><span class="sxs-lookup"><span data-stu-id="1357f-137">However, it does force the script to synchronize with the workbook to ensure the logged information is up-to-date.</span></span> <span data-ttu-id="1357f-138">Ziehen Sie in Betracht, unnötige Protokollierungsanweisungen (z. B. zu Testzwecken) zu entfernen, bevor Sie ihr Skript freigeben.</span><span class="sxs-lookup"><span data-stu-id="1357f-138">Consider removing unnecessary logging statements (such as those used for testing) before sharing your script.</span></span> <span data-ttu-id="1357f-139">Dies verursacht in der Regel kein erkennbares Leistungsproblem, es sei denn, die `console.log()` Anweisung befindet sich in einer Schleife.</span><span class="sxs-lookup"><span data-stu-id="1357f-139">This typically won't cause a noticeable performance issue, unless the `console.log()` statement is in a loop.</span></span>

### <a name="avoid-using-trycatch-blocks"></a><span data-ttu-id="1357f-140">Vermeiden der Verwendung von try/catch-Blöcken</span><span class="sxs-lookup"><span data-stu-id="1357f-140">Avoid using try/catch blocks</span></span>

<span data-ttu-id="1357f-141">Es wird nicht empfohlen, Blöcke [ `try` / `catch` als](https://developer.mozilla.org/docs/Web/JavaScript/Reference/Statements/try...catch) Teil des erwarteten Steuerungsflusses eines Skripts zu verwenden.</span><span class="sxs-lookup"><span data-stu-id="1357f-141">We don't recommend using [`try`/`catch` blocks](https://developer.mozilla.org/docs/Web/JavaScript/Reference/Statements/try...catch) as part of a script's expected control flow.</span></span> <span data-ttu-id="1357f-142">Die meisten Fehler können vermieden werden, indem Objekte überprüft werden, die von der Arbeitsmappe zurückgegeben werden.</span><span class="sxs-lookup"><span data-stu-id="1357f-142">Most errors can be avoided by checking objects returned from the workbook.</span></span> <span data-ttu-id="1357f-143">Mit dem folgenden Skript wird beispielsweise überprüft, ob die von der Arbeitsmappe zurückgegebene Tabelle vorhanden ist, bevor versucht wird, eine Zeile hinzuzufügen.</span><span class="sxs-lookup"><span data-stu-id="1357f-143">For example, the following script checks that the table returned by the workbook exists before trying to add a row.</span></span>

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

## <a name="case-by-case-help"></a><span data-ttu-id="1357f-144">Fall-für-Fall-Hilfe</span><span class="sxs-lookup"><span data-stu-id="1357f-144">Case-by-case help</span></span>

<span data-ttu-id="1357f-145">Wenn die Office -Skript-Plattform erweitert wird, um [mit Power Automate,](https://flow.microsoft.com/) [adaptiven](/adaptive-cards)Karten und anderen produktübergreifenden Features zu arbeiten, werden die Details der Skriptarbeitsmappenkommunikation komplizierter.</span><span class="sxs-lookup"><span data-stu-id="1357f-145">As the Office Scripts platform expands to work with [Power Automate](https://flow.microsoft.com/), [Adaptive Cards](/adaptive-cards), and other cross-product features, the details of the script-workbook communication become more intricate.</span></span> <span data-ttu-id="1357f-146">Wenn Sie Hilfe benötigen, damit Ihr Skript schneller ausgeführt wird, wenden Sie sich an [Stack Overflow.](https://stackoverflow.com/questions/tagged/office-scripts)</span><span class="sxs-lookup"><span data-stu-id="1357f-146">If you need help making your script run faster, please reach out through [Stack Overflow](https://stackoverflow.com/questions/tagged/office-scripts).</span></span> <span data-ttu-id="1357f-147">Markieren Sie Ihre Frage unbedingt mit "Office-Skripts", damit Experten sie finden und Hilfe erhalten können.</span><span class="sxs-lookup"><span data-stu-id="1357f-147">Be sure to tag your question with "office-scripts" so experts can find it and help.</span></span>

## <a name="see-also"></a><span data-ttu-id="1357f-148">Siehe auch</span><span class="sxs-lookup"><span data-stu-id="1357f-148">See also</span></span>

- [<span data-ttu-id="1357f-149">Grundlagen der Skripterstellung für Office-Skripts in Excel im Web</span><span class="sxs-lookup"><span data-stu-id="1357f-149">Scripting fundamentals for Office Scripts in Excel on the web</span></span>](scripting-fundamentals.md)
- [<span data-ttu-id="1357f-150">MDN-Web-Dokumente: Schleifen und Iteration</span><span class="sxs-lookup"><span data-stu-id="1357f-150">MDN web docs: Loops and iteration</span></span>](https://developer.mozilla.org/docs/Web/JavaScript/Guide/Loops_and_iteration)