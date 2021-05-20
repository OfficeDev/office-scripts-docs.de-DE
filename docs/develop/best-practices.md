---
title: Bewährte Methoden in Office-Skripts
description: So verhindern Sie häufige Probleme und schreiben robuste Office Skripts, die unerwartete Eingaben oder Daten verarbeiten können.
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

Diese Muster und Vorgehensweisen sollen Ihnen dabei helfen, Ihre Skripts jedes Mal erfolgreich auszuführen. Verwenden Sie sie, um häufige Fallstricke zu vermeiden, wenn Sie mit der Automatisierung Ihres Excel-Workflows beginnen.

## <a name="verify-an-object-is-present"></a>Überprüfen, ob ein Objekt vorhanden ist

Skripts basieren häufig darauf, dass ein bestimmtes Arbeitsblatt oder eine bestimmte Tabelle in der Arbeitsmappe vorhanden ist. Sie können jedoch zwischen Skriptausführungen umbenannt oder entfernt werden. Wenn Sie überprüfen, ob diese Tabellen oder Arbeitsblätter vorhanden sind, bevor Methoden aufgerufen werden, können Sie sicherstellen, dass das Skript nicht abrupt beendet wird.

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

Der TypeScript-Operator `?` überprüft, ob das Objekt vor dem Aufruf einer Methode vorhanden ist. Dies kann den Code schlanker machen, wenn Sie nichts Besonderes tun müssen, wenn das Objekt nicht vorhanden ist.

```TypeScript
// The ? ensures that the delete() API is only called if the object exists.
workbook.getWorksheet('Index')?.delete();
```

## <a name="validate-data-and-workbook-state-first"></a>Überprüfen des Daten- und Arbeitsmappenstatus zuerst

Stellen Sie sicher, dass alle Arbeitsblätter, Tabellen, Shapes und anderen Objekte vorhanden sind, bevor Sie an den Daten arbeiten. Überprüfen Sie anhand des vorherigen Musters, ob sich alles in der Arbeitsmappe befindet und Ihren Erwartungen entspricht. Wenn Sie dies tun, bevor Daten geschrieben werden, wird sichergestellt, dass das Skript die Arbeitsmappe nicht in einem Teilzustand belässt.

Das folgende Skript erfordert, dass zwei Tabellen mit den Namen "Table1" und "Table2" vorhanden sind. Das Skript überprüft zunächst, ob die Tabellen vorhanden sind, und endet dann mit der `return` Anweisung und einer entsprechenden Nachricht, falls dies nicht der Fall ist.

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

Wenn die Überprüfung in einer separaten Funktion erfolgt, müssen Sie das Skript dennoch beenden, indem Sie die `return` Anweisung aus der Funktion `main` ausgeben. Wenn Sie von der Unterfunktion zurückkehren, wird das Skript nicht beendet.

Das folgende Skript hat das gleiche Verhalten wie das vorherige. Der Unterschied besteht darin, dass die `main` Funktion die Funktion `inputPresent` aufruft, um alles zu überprüfen. `inputPresent` gibt einen booleschen ( `true` oder `false` ) zurück, um anzugeben, ob alle erforderlichen Eingaben vorhanden sind. Die `main` Funktion verwendet diese boolesche, um zu entscheiden, ob das Skript fortgesetzt oder beendet werden soll.

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

## <a name="when-to-use-a-throw-statement"></a>Wann eine Anweisung verwendet `throw` werden soll

Eine [`throw`](https://developer.mozilla.org/docs/web/javascript/reference/statements/throw) Anweisung weist darauf hin, dass ein unerwarteter Fehler aufgetreten ist. Der Code wird sofort beendet. Zum größten Teil müssen Sie nicht `throw` aus Ihrem Skript. Normalerweise informiert das Skript den Benutzer automatisch darüber, dass das Skript aufgrund eines Problems nicht ausgeführt werden konnte. In den meisten Fällen reicht es aus, das Skript mit einer Fehlermeldung und einer Anweisung der Funktion zu `return` `main` beenden.

Wenn Ihr Skript jedoch als Teil eines Power Automate-Flusses ausgeführt wird, sollten Sie verhindern, dass der Fluss fortgesetzt wird. Eine `throw` Anweisung stoppt das Skript und weist den Flow an, ebenfalls zu stoppen.

Das folgende Skript zeigt, wie die `throw` Anweisung in unserem Tabellenüberprüfungsbeispiel verwendet wird.

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

## <a name="when-to-use-a-trycatch-statement"></a>Wann eine Anweisung verwendet `try...catch` werden soll

Die [`try...catch`](https://developer.mozilla.org/docs/Web/JavaScript/Reference/Statements/try...catch) Anweisung ist eine Möglichkeit, zu erkennen, ob ein API-Aufruf fehlschlägt, und das Skript weiter auszuführen.

Betrachten Sie den folgenden Ausschnitt, der eine große Datenaktualisierung für einen Bereich durchführt.

```TypeScript
range.setValues(someLargeValues);
```

Wenn `someLargeValues` die Excel für das Web verarbeiten kann, schlägt der Aufruf `setValues()` fehl. Das Skript schlägt dann auch mit einem [Laufzeitfehler fehl.](../testing/troubleshooting.md#runtime-errors) Mit `try...catch` der Anweisung lässt Ihr Skript diese Bedingung erkennen, ohne das Skript sofort zu beenden und den Standardfehler anzuzeigen.

Ein Ansatz, um dem Skriptbenutzer eine bessere Benutzererfahrung zu bieten, besteht darin, ihm eine benutzerdefinierte Fehlermeldung zu präsentieren. Der folgende Ausschnitt zeigt eine Anweisung, die `try...catch` mehr Fehlerinformationen protokolliert, um dem Leser besser zu helfen.

```TypeScript
try {
    range.setValues(someLargeValues);
} catch (error) {
    console.log(`The script failed to update the values at location ${range.getAddress()}. Please inspect and run again.`);
    console.log(error);
    return; // End the script (assuming this is in the main function).
}
```

Ein weiterer Ansatz für den Umgang mit Fehlern ist ein Fallbackverhalten, das den Fehlerfall behandelt. Der folgende Ausschnitt verwendet den `catch` Block, um zu versuchen, eine alternative Methode zu versuchen, die Aktualisierung in kleinere Teile aufzuteilen und den Fehler zu vermeiden.

> [!TIP]
> Ein vollständiges Beispiel zum Aktualisieren eines großen Bereichs finden Sie unter [Schreiben eines großen Datasets](../resources/samples/write-large-dataset.md).

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
> Die Verwendung `try...catch` innerhalb oder um eine Schleife verlangsamt Ihr Skript. Weitere Leistungsinformationen finden Sie unter [Vermeiden der Verwendung von `try...catch` Blöcken](web-client-performance.md#avoid-using-trycatch-blocks-in-or-surrounding-loops).

## <a name="see-also"></a>Siehe auch

- [Behandeln von Problemen mit Office-Skripts](../testing/troubleshooting.md)
- [Fehlerbehebungsinformationen für Power Automate mit Office Skripts](../testing/power-automate-troubleshooting.md)
- [Plattformlimits mit Office Scripts](../testing/platform-limits.md)
- [Verbessern Sie die Leistung Ihrer Office Scripts](web-client-performance.md)
