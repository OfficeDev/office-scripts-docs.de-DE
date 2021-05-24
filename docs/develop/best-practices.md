---
title: Bewährte Methoden in Office-Skripts
description: So verhindern Sie häufige Probleme und schreiben Office Skripts, die unerwartete Eingaben oder Daten verarbeiten können.
ms.date: 05/10/2021
localization_priority: Normal
ms.openlocfilehash: 0697e6fd1fa8f437a4a585d938254deb5a05f20c
ms.sourcegitcommit: 4687693f02fc90a57ba30c461f35046e02e6f5fb
ms.translationtype: MT
ms.contentlocale: de-DE
ms.lasthandoff: 05/19/2021
ms.locfileid: "52546025"
---
# <a name="best-practices-in-office-scripts"></a><span data-ttu-id="2f20f-103">Bewährte Methoden in Office-Skripts</span><span class="sxs-lookup"><span data-stu-id="2f20f-103">Best practices in Office Scripts</span></span>

<span data-ttu-id="2f20f-104">Diese Muster und Methoden sind so konzipiert, dass Ihre Skripts jedes Mal erfolgreich ausgeführt werden können.</span><span class="sxs-lookup"><span data-stu-id="2f20f-104">These patterns and practices are designed to help your scripts run successfully every time.</span></span> <span data-ttu-id="2f20f-105">Verwenden Sie sie, um häufige Fallstricke zu vermeiden, wenn Sie mit der Automatisierung Excel beginnen.</span><span class="sxs-lookup"><span data-stu-id="2f20f-105">Use them to avoid common pitfalls as you start automating your Excel workflow.</span></span>

## <a name="verify-an-object-is-present"></a><span data-ttu-id="2f20f-106">Überprüfen, ob ein Objekt vorhanden ist</span><span class="sxs-lookup"><span data-stu-id="2f20f-106">Verify an object is present</span></span>

<span data-ttu-id="2f20f-107">Skripts verlassen sich häufig darauf, dass ein bestimmtes Arbeitsblatt oder eine bestimmte Tabelle in der Arbeitsmappe vorhanden ist.</span><span class="sxs-lookup"><span data-stu-id="2f20f-107">Scripts often rely on a certain worksheet or table being present in the workbook.</span></span> <span data-ttu-id="2f20f-108">Sie werden jedoch möglicherweise zwischen Skriptläufen umbenannt oder entfernt.</span><span class="sxs-lookup"><span data-stu-id="2f20f-108">However, they might get renamed or removed between script runs.</span></span> <span data-ttu-id="2f20f-109">Indem Sie überprüfen, ob diese Tabellen oder Arbeitsblätter vorhanden sind, bevor Sie Methoden für sie aufrufen, können Sie sicherstellen, dass das Skript nicht abrupt endet.</span><span class="sxs-lookup"><span data-stu-id="2f20f-109">By checking if those tables or worksheets exist before calling methods on them, you can make sure the script doesn't end abruptly.</span></span>

<span data-ttu-id="2f20f-110">Der folgende Beispielcode überprüft, ob das Arbeitsblatt "Index" in der Arbeitsmappe vorhanden ist.</span><span class="sxs-lookup"><span data-stu-id="2f20f-110">The following sample code checks if the "Index" worksheet is present in the workbook.</span></span> <span data-ttu-id="2f20f-111">Wenn das Arbeitsblatt vorhanden ist, ruft das Skript einen Bereich ab und geht weiter.</span><span class="sxs-lookup"><span data-stu-id="2f20f-111">If the worksheet is present, the script gets a range and proceeds.</span></span> <span data-ttu-id="2f20f-112">Wenn es nicht vorhanden ist, protokolliert das Skript eine benutzerdefinierte Fehlermeldung.</span><span class="sxs-lookup"><span data-stu-id="2f20f-112">If it isn't present, the script logs a custom error message.</span></span>

```TypeScript
// Make sure the "Index" worksheet exists before using it.
let indexSheet = workbook.getWorksheet('Index');
if (indexSheet) {
  let range = indexSheet.getRange("A1");
  // Continue using the range...
} else {
  console.log("Index sheet not found.");
}
```

<span data-ttu-id="2f20f-113">Der `?` TypeScript-Operator überprüft, ob das Objekt vorhanden ist, bevor eine Methode aufruft.</span><span class="sxs-lookup"><span data-stu-id="2f20f-113">The TypeScript `?` operator checks if the object exists before calling a method.</span></span> <span data-ttu-id="2f20f-114">Dadurch kann Der Code optimiert werden, wenn Sie nichts Besonderes tun müssen, wenn das Objekt nicht vorhanden ist.</span><span class="sxs-lookup"><span data-stu-id="2f20f-114">This can make your code more streamlined if you don't need to do anything special when the object doesn't exist.</span></span>

```TypeScript
// The ? ensures that the delete() API is only called if the object exists.
workbook.getWorksheet('Index')?.delete();
```

## <a name="validate-data-and-workbook-state-first"></a><span data-ttu-id="2f20f-115">Überprüfen des Daten- und Arbeitsmappenstatus zuerst</span><span class="sxs-lookup"><span data-stu-id="2f20f-115">Validate data and workbook state first</span></span>

<span data-ttu-id="2f20f-116">Stellen Sie sicher, dass alle Arbeitsblätter, Tabellen, Formen und anderen Objekte vorhanden sind, bevor Sie mit den Daten arbeiten.</span><span class="sxs-lookup"><span data-stu-id="2f20f-116">Make sure all your worksheets, tables, shapes, and other objects are present before working on the data.</span></span> <span data-ttu-id="2f20f-117">Überprüfen Sie anhand des vorherigen Musters, ob sich alles in der Arbeitsmappe befindet und Ihren Erwartungen entspricht.</span><span class="sxs-lookup"><span data-stu-id="2f20f-117">Using the previous pattern, check to see if everything is in the workbook and matches your expectations.</span></span> <span data-ttu-id="2f20f-118">Wenn Sie dies tun, bevor Daten geschrieben werden, wird sichergestellt, dass das Skript die Arbeitsmappe nicht in einem Teilzustand be lässt.</span><span class="sxs-lookup"><span data-stu-id="2f20f-118">Doing this before any data is written ensures your script doesn't leave the workbook in a partial state.</span></span>

<span data-ttu-id="2f20f-119">Das folgende Skript erfordert, dass zwei Tabellen mit dem Namen "Table1" und "Table2" vorhanden sind.</span><span class="sxs-lookup"><span data-stu-id="2f20f-119">The following script requires two tables named "Table1" and "Table2" to be present.</span></span> <span data-ttu-id="2f20f-120">Das Skript überprüft zunächst, ob die Tabellen vorhanden sind, und endet dann mit der Anweisung und einer entsprechenden Meldung, falls `return` nicht.</span><span class="sxs-lookup"><span data-stu-id="2f20f-120">The script first checks if the tables are present and then ends with the `return` statement and an appropriate message if they're not.</span></span>

```TypeScript
function main(workbook: ExcelScript.Workbook) {
  // These tables must be in the workbook for the script.
  const TargetTableName = 'Table1';
  const SourceTableName = 'Table2';

  // Get the table objects.
  let targetTable = workbook.getTable(TargetTableName);
  let sourceTable = workbook.getTable(SourceTableName);

  // Check if the tables are there.
  if (!targetTable || !sourceTable) {
    console.log(`Required tables missing - Check that both the source (${TargetTableName}) and target (${SourceTableName}) tables are present before running the script.`);
    return;
  }

  // Continue....
}
```

<span data-ttu-id="2f20f-121">Wenn die Überprüfung in einer separaten Funktion ausgeführt wird, müssen Sie das Skript weiterhin beenden, indem Sie die `return` Anweisung aus der Funktion `main` ausgeben.</span><span class="sxs-lookup"><span data-stu-id="2f20f-121">If the verification is happening in a separate function, you still must end the script by issuing the `return` statement from the `main` function.</span></span> <span data-ttu-id="2f20f-122">Die Rückgabe von der Unterfunktion beendet das Skript nicht.</span><span class="sxs-lookup"><span data-stu-id="2f20f-122">Returning from the subfunction doesn't end the script.</span></span>

<span data-ttu-id="2f20f-123">Das folgende Skript hat dasselbe Verhalten wie das vorherige Skript.</span><span class="sxs-lookup"><span data-stu-id="2f20f-123">The following script has the same behavior as the previous one.</span></span> <span data-ttu-id="2f20f-124">Der Unterschied ist, dass `main` die Funktion die Funktion `inputPresent` aufruft, um alles zu überprüfen.</span><span class="sxs-lookup"><span data-stu-id="2f20f-124">The difference is that the `main` function calls the `inputPresent` function to verify everything.</span></span> <span data-ttu-id="2f20f-125">`inputPresent` gibt einen booleschen Wert ( `true` oder `false` ) zurück, um anzugeben, ob alle erforderlichen Eingaben vorhanden sind.</span><span class="sxs-lookup"><span data-stu-id="2f20f-125">`inputPresent` returns a boolean (`true` or `false`) to indicate whether all required inputs are present.</span></span> <span data-ttu-id="2f20f-126">Die `main` Funktion verwendet diesen booleschen Wert, um zu entscheiden, ob das Skript fortgesetzt oder beendet wird.</span><span class="sxs-lookup"><span data-stu-id="2f20f-126">The `main` function uses that boolean to decide on continuing or ending the script.</span></span>

```TypeScript
function main(workbook: ExcelScript.Workbook) {

  // Get the table objects.
  if (!inputPresent(workbook)) {
    return;
  }

  // Continue....
}

function inputPresent( workbook: ExcelScript.Workbook): boolean {
  // These tables must be in the workbook for the script.
  const TargetTableName = 'Table1';
  const SourceTableName = 'Table2';

  // Get the table objects.
  let targetTable = workbook.getTable(TargetTableName);
  let sourceTable = workbook.getTable(SourceTableName);

  // Check if the tables are there.
  if (!targetTable || !sourceTable) {
    console.log(`Required tables missing - Check that both the source (${TargetTableName}) and target (${SourceTableName}) tables are present before running the script.`);
    return false;
  }

  return true;
}
```

## <a name="when-to-use-a-throw-statement"></a><span data-ttu-id="2f20f-127">Verwendung einer `throw` Anweisung</span><span class="sxs-lookup"><span data-stu-id="2f20f-127">When to use a `throw` statement</span></span>

<span data-ttu-id="2f20f-128">Eine [`throw`](https://developer.mozilla.org/docs/web/javascript/reference/statements/throw) Anweisung gibt an, dass ein unerwarteter Fehler aufgetreten ist.</span><span class="sxs-lookup"><span data-stu-id="2f20f-128">A [`throw`](https://developer.mozilla.org/docs/web/javascript/reference/statements/throw) statement indicates an unexpected error has occurred.</span></span> <span data-ttu-id="2f20f-129">Der Code wird sofort beendet.</span><span class="sxs-lookup"><span data-stu-id="2f20f-129">It ends the code immediately.</span></span> <span data-ttu-id="2f20f-130">In den meisten Beispielen müssen Sie ihr Skript `throw` nicht verwenden.</span><span class="sxs-lookup"><span data-stu-id="2f20f-130">For the most part, you don't need to `throw` from your script.</span></span> <span data-ttu-id="2f20f-131">In der Regel informiert das Skript den Benutzer automatisch, dass das Skript aufgrund eines Problems nicht ausgeführt werden konnte.</span><span class="sxs-lookup"><span data-stu-id="2f20f-131">Usually, the script automatically informs the user that the script failed to run due to an issue.</span></span> <span data-ttu-id="2f20f-132">In den meisten Fällen reicht es aus, das Skript mit einer Fehlermeldung und einer Anweisung `return` aus der Funktion zu `main` beenden.</span><span class="sxs-lookup"><span data-stu-id="2f20f-132">In most cases, it's sufficient to end the script with an error message and a `return` statement from the `main` function.</span></span>

<span data-ttu-id="2f20f-133">Wenn Ihr Skript jedoch im Rahmen eines Power Automate ausgeführt wird, sollten Sie möglicherweise verhindern, dass der Fluss fortgesetzt wird.</span><span class="sxs-lookup"><span data-stu-id="2f20f-133">However, if your script is running as part of a Power Automate flow, you may want to stop the flow from continuing.</span></span> <span data-ttu-id="2f20f-134">Eine `throw` Anweisung stoppt das Skript und weist den Fluss an, ebenfalls zu beenden.</span><span class="sxs-lookup"><span data-stu-id="2f20f-134">A `throw` statement stops the script and tells the flow to stop as well.</span></span>

<span data-ttu-id="2f20f-135">Das folgende Skript zeigt, wie Sie die `throw` Anweisung in unserem Beispiel für die Tabellenüberprüfung verwenden.</span><span class="sxs-lookup"><span data-stu-id="2f20f-135">The following script shows how to use the `throw` statement in our table checking example.</span></span>

```TypeScript
function main(workbook: ExcelScript.Workbook) {
  // These tables must be in the workbook for the script.
  const TargetTableName = 'Table1';
  const SourceTableName = 'Table2';

  // Get the table objects.
  let targetTable = workbook.getTable(TargetTableName);
  let sourceTable = workbook.getTable(SourceTableName);

  // Check if the tables are there.
  if (!targetTable || !sourceTable) {
    // Immediately end the script with an error.
    throw `Required tables missing - Check that both the source (${TargetTableName}) and target (${SourceTableName}) tables are present before running the script.`;
  }
  
```

## <a name="when-to-use-a-trycatch-statement"></a><span data-ttu-id="2f20f-136">Verwendung einer `try...catch` Anweisung</span><span class="sxs-lookup"><span data-stu-id="2f20f-136">When to use a `try...catch` statement</span></span>

<span data-ttu-id="2f20f-137">Die [`try...catch`](https://developer.mozilla.org/docs/Web/JavaScript/Reference/Statements/try...catch) Anweisung ist eine Möglichkeit, um zu erkennen, ob ein API-Aufruf fehlschlägt, und das Skript weiter auszuführen.</span><span class="sxs-lookup"><span data-stu-id="2f20f-137">The [`try...catch`](https://developer.mozilla.org/docs/Web/JavaScript/Reference/Statements/try...catch) statement is a way to detect if an API call fails and continue running the script.</span></span>

<span data-ttu-id="2f20f-138">Betrachten Sie den folgenden Codeausschnitt, der eine große Datenaktualisierung für einen Bereich ausführt.</span><span class="sxs-lookup"><span data-stu-id="2f20f-138">Consider the following snippet that performs a large data update on a range.</span></span>

```TypeScript
range.setValues(someLargeValues);
```

<span data-ttu-id="2f20f-139">Wenn `someLargeValues` die Größe größer Excel das Web verarbeiten kann, schlägt der Aufruf `setValues()` fehl.</span><span class="sxs-lookup"><span data-stu-id="2f20f-139">If `someLargeValues` is larger than Excel for the web can handle, the `setValues()` call fails.</span></span> <span data-ttu-id="2f20f-140">Das Skript schlägt dann auch mit einem [Laufzeitfehler fehl.](../testing/troubleshooting.md#runtime-errors)</span><span class="sxs-lookup"><span data-stu-id="2f20f-140">The script then also fails with a [runtime error](../testing/troubleshooting.md#runtime-errors).</span></span> <span data-ttu-id="2f20f-141">Mit der Anweisung kann Ihr Skript diese Bedingung erkennen, ohne das Skript sofort zu beenden `try...catch` und den Standardfehler zu zeigen.</span><span class="sxs-lookup"><span data-stu-id="2f20f-141">The `try...catch` statement lets your script recognize this condition, without immediately ending the script and showing the default error.</span></span>

<span data-ttu-id="2f20f-142">Ein Ansatz, um dem Skriptbenutzer eine bessere Benutzererfahrung zu bieten, besteht in der Benutzerdefinierten Fehlermeldung.</span><span class="sxs-lookup"><span data-stu-id="2f20f-142">One approach for giving the script user a better experience is to present them a custom error message.</span></span> <span data-ttu-id="2f20f-143">Der folgende Codeausschnitt zeigt eine `try...catch` Anweisung, in der weitere Fehlerinformationen protokollieren werden, um dem Leser besser zu helfen.</span><span class="sxs-lookup"><span data-stu-id="2f20f-143">The following snippet shows a `try...catch` statement logging more error information to better help the reader.</span></span>

```TypeScript
try {
    range.setValues(someLargeValues);
} catch (error) {
    console.log(`The script failed to update the values at location ${range.getAddress()}. Please inspect and run again.`);
    console.log(error);
    return; // End the script (assuming this is in the main function).
}
```

<span data-ttu-id="2f20f-144">Ein weiterer Ansatz zum Umgang mit Fehlern ist das Fallbackverhalten, das den Fehlerfall behandelt.</span><span class="sxs-lookup"><span data-stu-id="2f20f-144">Another approach to dealing with errors is to have fallback behavior that handles the error case.</span></span> <span data-ttu-id="2f20f-145">Im folgenden Codeausschnitt wird der Block verwendet, um eine alternative Methode zu testen, um das Update in kleinere Teile aufteilen und `catch` den Fehler zu vermeiden.</span><span class="sxs-lookup"><span data-stu-id="2f20f-145">The following snippet uses the `catch` block to try an alternate method break up the update into smaller pieces and avoid the error.</span></span>

> [!TIP]
> <span data-ttu-id="2f20f-146">Ein vollständiges Beispiel zum Aktualisieren eines großen Bereichs finden Sie unter [Write a large dataset](../resources/samples/write-large-dataset.md).</span><span class="sxs-lookup"><span data-stu-id="2f20f-146">For a full example on how to update a large range, see [Write a large dataset](../resources/samples/write-large-dataset.md).</span></span>

```TypeScript
try {
    range.setValues(someLargeValues);
} catch (error) {
    console.log(`The script failed to update the values at location ${range.getAddress()}. Trying a different approach.`);
    handleUpdatesInSmallerBatches(someLargeValues);
}

// Continue...
}
```

> [!NOTE]
> <span data-ttu-id="2f20f-147">Die `try...catch` Verwendung innerhalb oder um eine Schleife verlangsamt Ihr Skript.</span><span class="sxs-lookup"><span data-stu-id="2f20f-147">Using `try...catch` inside or around a loop slows down your script.</span></span> <span data-ttu-id="2f20f-148">Weitere Leistungsinformationen finden Sie unter [Vermeiden der Verwendung von `try...catch` Blöcken](web-client-performance.md#avoid-using-trycatch-blocks-in-or-surrounding-loops).</span><span class="sxs-lookup"><span data-stu-id="2f20f-148">For more performance information, see [Avoid using `try...catch` blocks](web-client-performance.md#avoid-using-trycatch-blocks-in-or-surrounding-loops).</span></span>

## <a name="see-also"></a><span data-ttu-id="2f20f-149">Siehe auch</span><span class="sxs-lookup"><span data-stu-id="2f20f-149">See also</span></span>

- [<span data-ttu-id="2f20f-150">Behandeln von Problemen mit Office-Skripts</span><span class="sxs-lookup"><span data-stu-id="2f20f-150">Troubleshooting Office Scripts</span></span>](../testing/troubleshooting.md)
- [<span data-ttu-id="2f20f-151">Problembehandlungsinformationen für Power Automate mit Office Skripts</span><span class="sxs-lookup"><span data-stu-id="2f20f-151">Troubleshooting information for Power Automate with Office Scripts</span></span>](../testing/power-automate-troubleshooting.md)
- [<span data-ttu-id="2f20f-152">Plattformbeschränkungen mit Office Skripts</span><span class="sxs-lookup"><span data-stu-id="2f20f-152">Platform limits with Office Scripts</span></span>](../testing/platform-limits.md)
- [<span data-ttu-id="2f20f-153">Verbessern der Leistung Ihrer Office Skripts</span><span class="sxs-lookup"><span data-stu-id="2f20f-153">Improve the performance of your Office Scripts</span></span>](web-client-performance.md)
