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
# <a name="best-practices-in-office-scripts"></a>Bewährte Methoden in Office-Skripts

Diese Muster und Methoden sind so konzipiert, dass Ihre Skripts jedes Mal erfolgreich ausgeführt werden können. Verwenden Sie sie, um häufige Fallstricke zu vermeiden, wenn Sie mit der Automatisierung Excel beginnen.

## <a name="verify-an-object-is-present"></a>Überprüfen, ob ein Objekt vorhanden ist

Skripts verlassen sich häufig darauf, dass ein bestimmtes Arbeitsblatt oder eine bestimmte Tabelle in der Arbeitsmappe vorhanden ist. Sie werden jedoch möglicherweise zwischen Skriptläufen umbenannt oder entfernt. Indem Sie überprüfen, ob diese Tabellen oder Arbeitsblätter vorhanden sind, bevor Sie Methoden für sie aufrufen, können Sie sicherstellen, dass das Skript nicht abrupt endet.

Der folgende Beispielcode überprüft, ob das Arbeitsblatt "Index" in der Arbeitsmappe vorhanden ist. Wenn das Arbeitsblatt vorhanden ist, ruft das Skript einen Bereich ab und geht weiter. Wenn es nicht vorhanden ist, protokolliert das Skript eine benutzerdefinierte Fehlermeldung.

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

Der `?` TypeScript-Operator überprüft, ob das Objekt vorhanden ist, bevor eine Methode aufruft. Dadurch kann Der Code optimiert werden, wenn Sie nichts Besonderes tun müssen, wenn das Objekt nicht vorhanden ist.

```TypeScript
// The ? ensures that the delete() API is only called if the object exists.
workbook.getWorksheet('Index')?.delete();
```

## <a name="validate-data-and-workbook-state-first"></a>Überprüfen des Daten- und Arbeitsmappenstatus zuerst

Stellen Sie sicher, dass alle Arbeitsblätter, Tabellen, Formen und anderen Objekte vorhanden sind, bevor Sie mit den Daten arbeiten. Überprüfen Sie anhand des vorherigen Musters, ob sich alles in der Arbeitsmappe befindet und Ihren Erwartungen entspricht. Wenn Sie dies tun, bevor Daten geschrieben werden, wird sichergestellt, dass das Skript die Arbeitsmappe nicht in einem Teilzustand be lässt.

Das folgende Skript erfordert, dass zwei Tabellen mit dem Namen "Table1" und "Table2" vorhanden sind. Das Skript überprüft zunächst, ob die Tabellen vorhanden sind, und endet dann mit der Anweisung und einer entsprechenden Meldung, falls `return` nicht.

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

Wenn die Überprüfung in einer separaten Funktion ausgeführt wird, müssen Sie das Skript weiterhin beenden, indem Sie die `return` Anweisung aus der Funktion `main` ausgeben. Die Rückgabe von der Unterfunktion beendet das Skript nicht.

Das folgende Skript hat dasselbe Verhalten wie das vorherige Skript. Der Unterschied ist, dass `main` die Funktion die Funktion `inputPresent` aufruft, um alles zu überprüfen. `inputPresent` gibt einen booleschen Wert ( `true` oder `false` ) zurück, um anzugeben, ob alle erforderlichen Eingaben vorhanden sind. Die `main` Funktion verwendet diesen booleschen Wert, um zu entscheiden, ob das Skript fortgesetzt oder beendet wird.

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

## <a name="when-to-use-a-throw-statement"></a>Verwendung einer `throw` Anweisung

Eine [`throw`](https://developer.mozilla.org/docs/web/javascript/reference/statements/throw) Anweisung gibt an, dass ein unerwarteter Fehler aufgetreten ist. Der Code wird sofort beendet. In den meisten Beispielen müssen Sie ihr Skript `throw` nicht verwenden. In der Regel informiert das Skript den Benutzer automatisch, dass das Skript aufgrund eines Problems nicht ausgeführt werden konnte. In den meisten Fällen reicht es aus, das Skript mit einer Fehlermeldung und einer Anweisung `return` aus der Funktion zu `main` beenden.

Wenn Ihr Skript jedoch im Rahmen eines Power Automate ausgeführt wird, sollten Sie möglicherweise verhindern, dass der Fluss fortgesetzt wird. Eine `throw` Anweisung stoppt das Skript und weist den Fluss an, ebenfalls zu beenden.

Das folgende Skript zeigt, wie Sie die `throw` Anweisung in unserem Beispiel für die Tabellenüberprüfung verwenden.

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

## <a name="when-to-use-a-trycatch-statement"></a>Verwendung einer `try...catch` Anweisung

Die [`try...catch`](https://developer.mozilla.org/docs/Web/JavaScript/Reference/Statements/try...catch) Anweisung ist eine Möglichkeit, um zu erkennen, ob ein API-Aufruf fehlschlägt, und das Skript weiter auszuführen.

Betrachten Sie den folgenden Codeausschnitt, der eine große Datenaktualisierung für einen Bereich ausführt.

```TypeScript
range.setValues(someLargeValues);
```

Wenn `someLargeValues` die Größe größer Excel das Web verarbeiten kann, schlägt der Aufruf `setValues()` fehl. Das Skript schlägt dann auch mit einem [Laufzeitfehler fehl.](../testing/troubleshooting.md#runtime-errors) Mit der Anweisung kann Ihr Skript diese Bedingung erkennen, ohne das Skript sofort zu beenden `try...catch` und den Standardfehler zu zeigen.

Ein Ansatz, um dem Skriptbenutzer eine bessere Benutzererfahrung zu bieten, besteht in der Benutzerdefinierten Fehlermeldung. Der folgende Codeausschnitt zeigt eine `try...catch` Anweisung, in der weitere Fehlerinformationen protokollieren werden, um dem Leser besser zu helfen.

```TypeScript
try {
    range.setValues(someLargeValues);
} catch (error) {
    console.log(`The script failed to update the values at location ${range.getAddress()}. Please inspect and run again.`);
    console.log(error);
    return; // End the script (assuming this is in the main function).
}
```

Ein weiterer Ansatz zum Umgang mit Fehlern ist das Fallbackverhalten, das den Fehlerfall behandelt. Im folgenden Codeausschnitt wird der Block verwendet, um eine alternative Methode zu testen, um das Update in kleinere Teile aufteilen und `catch` den Fehler zu vermeiden.

> [!TIP]
> Ein vollständiges Beispiel zum Aktualisieren eines großen Bereichs finden Sie unter [Write a large dataset](../resources/samples/write-large-dataset.md).

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
> Die `try...catch` Verwendung innerhalb oder um eine Schleife verlangsamt Ihr Skript. Weitere Leistungsinformationen finden Sie unter [Vermeiden der Verwendung von `try...catch` Blöcken](web-client-performance.md#avoid-using-trycatch-blocks-in-or-surrounding-loops).

## <a name="see-also"></a>Siehe auch

- [Behandeln von Problemen mit Office-Skripts](../testing/troubleshooting.md)
- [Problembehandlungsinformationen für Power Automate mit Office Skripts](../testing/power-automate-troubleshooting.md)
- [Plattformbeschränkungen mit Office Skripts](../testing/platform-limits.md)
- [Verbessern der Leistung Ihrer Office Skripts](web-client-performance.md)
