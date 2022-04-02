---
title: Bewährte Methoden in Office-Skripts
description: So verhindern Sie häufige Probleme und schreiben robuste Office Skripts, die unerwartete Eingaben oder Daten verarbeiten können.
ms.date: 12/29/2021
ms.localizationpriority: medium
ms.openlocfilehash: 689196e1a0ca70c999ec8048de64190cbfe75581
ms.sourcegitcommit: 7023b9e23499806901a5ecf8ebc460b76887cca6
ms.translationtype: MT
ms.contentlocale: de-DE
ms.lasthandoff: 03/31/2022
ms.locfileid: "64585765"
---
# <a name="best-practices-in-office-scripts"></a>Bewährte Methoden in Office-Skripts

Diese Muster und Methoden sind so konzipiert, dass Ihre Skripts jedes Mal erfolgreich ausgeführt werden können. Verwenden Sie sie, um häufige Fallstricke zu vermeiden, wenn Sie mit der Automatisierung Ihres Excel Workflows beginnen.

## <a name="use-the-action-recorder-to-learn-new-features"></a>Verwenden des Action Recorders zum Erlernen neuer Features

Excel erledigt viele Dinge. Die meisten davon können in Skripts geschrieben werden. Der Action Recorder zeichnet Ihre Excel Aktionen auf und übersetzt sie in Code. Dies ist die einfachste Möglichkeit, um zu erfahren, wie verschiedene Features mit Office Skripts funktionieren. Wenn Sie Code für eine bestimmte Aktion benötigen, wechseln Sie zum Action Recorder, führen Sie die Aktionen aus, wählen **Sie "Als Code kopieren"** aus, und fügen Sie den resultierenden Code in Ihr Skript ein.

:::image type="content" source="../images/action-recorder-copy-code.png" alt-text="Der Aufgabenbereich „Aktionsrekorder“ mit hervorgehobener Schaltfläche „Als Code kopieren“.":::

## <a name="verify-an-object-is-present"></a>Überprüfen, ob ein Objekt vorhanden ist

Skripts basieren häufig darauf, dass ein bestimmtes Arbeitsblatt oder eine bestimmte Tabelle in der Arbeitsmappe vorhanden ist. Sie können jedoch zwischen Skriptläufen umbenannt oder entfernt werden. Indem Sie überprüfen, ob diese Tabellen oder Arbeitsblätter vorhanden sind, bevor Sie Methoden für sie aufrufen, können Sie sicherstellen, dass das Skript nicht abrupt beendet wird.

Im folgenden Beispielcode wird überprüft, ob das Arbeitsblatt "Index" in der Arbeitsmappe vorhanden ist. Wenn das Arbeitsblatt vorhanden ist, ruft das Skript einen Bereich ab und setzt den Vorgang fort. Wenn es nicht vorhanden ist, protokolliert das Skript eine benutzerdefinierte Fehlermeldung.

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

Der TypeScript-Operator `?` überprüft, ob das Objekt vorhanden ist, bevor eine Methode aufgerufen wird. Dadurch kann Ihr Code optimiert werden, wenn Sie nichts Besonderes tun müssen, wenn das Objekt nicht vorhanden ist.

```TypeScript
// The ? ensures that the delete() API is only called if the object exists.
workbook.getWorksheet('Index')?.delete();
```

## <a name="validate-data-and-workbook-state-first"></a>Überprüfen des Daten- und Arbeitsmappenstatus zuerst

Stellen Sie sicher, dass alle Arbeitsblätter, Tabellen, Shapes und anderen Objekte vorhanden sind, bevor Sie an den Daten arbeiten. Überprüfen Sie anhand des vorherigen Musters, ob sich alles in der Arbeitsmappe befindet und Ihren Erwartungen entspricht. Wenn Sie dies tun, bevor Daten geschrieben werden, wird sichergestellt, dass das Skript die Arbeitsmappe nicht in einem Teilstatus belässt.

Das folgende Skript erfordert, dass zwei Tabellen mit dem Namen "Table1" und "Table2" vorhanden sind. Das Skript überprüft zuerst, ob die Tabellen vorhanden sind, und endet dann mit der `return` Anweisung und einer entsprechenden Meldung, wenn dies nicht der Fall ist.

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

  // Continue...
}
```

Wenn die Überprüfung in einer separaten Funktion erfolgt, müssen Sie dennoch das Skript beenden, indem Sie die `return` Anweisung aus der `main` Funktion ausgeben. Wenn Sie von der Unterfunktion zurückkehren, wird das Skript nicht beendet.

Das folgende Skript hat das gleiche Verhalten wie das vorherige. Der Unterschied besteht darin, dass die `main` Funktion die `inputPresent` Funktion aufruft, um alles zu überprüfen. `inputPresent` gibt einen booleschen Wert (`true` oder `false`) zurück, um anzugeben, ob alle erforderlichen Eingaben vorhanden sind. Die `main` Funktion verwendet diesen booleschen Wert, um zu entscheiden, ob das Skript fortgesetzt oder beendet werden soll.

```TypeScript
function main(workbook: ExcelScript.Workbook) {

  // Get the table objects.
  if (!inputPresent(workbook)) {
    return;
  }

  // Continue...
}

function inputPresent(workbook: ExcelScript.Workbook): boolean {
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

## <a name="when-to-use-a-throw-statement"></a>Wann sollte eine `throw` Anweisung verwendet werden?

Eine [`throw`](https://developer.mozilla.org/docs/web/javascript/reference/statements/throw) Anweisung gibt an, dass ein unerwarteter Fehler aufgetreten ist. Der Code wird sofort beendet. Sie müssen `throw` in den meisten Fällen nicht aus Ihrem Skript stammen. In der Regel informiert das Skript den Benutzer automatisch, dass das Skript aufgrund eines Problems nicht ausgeführt werden konnte. In den meisten Fällen reicht es aus, das Skript mit einer Fehlermeldung und einer `return` Anweisung aus der `main` Funktion zu beenden.

Wenn Ihr Skript jedoch als Teil eines Power Automate-Flusses ausgeführt wird, sollten Sie verhindern, dass der Fluss fortgesetzt wird. Eine `throw` Anweisung stoppt das Skript und weist den Fluss an, auch zu beenden.

Das folgende Skript zeigt, wie Sie die Anweisung in unserem Beispiel für die `throw` Tabellenüberprüfung verwenden.

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

## <a name="when-to-use-a-trycatch-statement"></a>Wann sollte eine `try...catch` Anweisung verwendet werden?

Die [`try...catch`](https://developer.mozilla.org/docs/Web/JavaScript/Reference/Statements/try...catch) Anweisung ist eine Möglichkeit, um zu erkennen, ob ein API-Aufruf fehlschlägt, und das Skript weiter auszuführen.

Betrachten Sie den folgenden Codeausschnitt, der eine Aktualisierung großer Daten für einen Bereich durchführt.

```TypeScript
range.setValues(someLargeValues);
```

Wenn `someLargeValues` größer ist, als Excel für das Web verarbeiten kann, schlägt der `setValues()` Aufruf fehl. Das Skript schlägt dann auch mit einem [Laufzeitfehler](../testing/troubleshooting.md#runtime-errors) fehl. Mit `try...catch` der Anweisung kann das Skript diese Bedingung erkennen, ohne das Skript sofort zu beenden und den Standardfehler anzuzeigen.

Eine Möglichkeit, dem Skriptbenutzer eine bessere Benutzererfahrung zu bieten, besteht darin, ihm eine benutzerdefinierte Fehlermeldung anzuzeigen. Der folgende Codeausschnitt zeigt eine `try...catch` Anweisung, in der weitere Fehlerinformationen protokolliert werden, um dem Leser besser zu helfen.

```TypeScript
try {
    range.setValues(someLargeValues);
} catch (error) {
    console.log(`The script failed to update the values at location ${range.getAddress()}. Please inspect and run again.`);
    console.log(error);
    return; // End the script (assuming this is in the main function).
}
```

Ein weiterer Ansatz für den Umgang mit Fehlern ist ein Fallbackverhalten, das den Fehlerfall behandelt. Der folgende Codeausschnitt verwendet den `catch` Block, um eine alternative Methode auszuprobieren, um die Aktualisierung in kleinere Teile aufzuteilen und den Fehler zu vermeiden.

> [!TIP]
> Ein vollständiges Beispiel zum Aktualisieren eines großen Bereichs finden Sie unter ["Schreiben eines großen Datasets](../resources/samples/write-large-dataset.md)".

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
> Die Verwendung `try...catch` innerhalb oder um eine Schleife verlangsamt Ihr Skript. Weitere Informationen zur Leistung finden Sie unter ["Vermeiden der Verwendung von `try...catch` Blöcken](web-client-performance.md#avoid-using-trycatch-blocks-in-or-surrounding-loops)".

## <a name="see-also"></a>Weitere Informationen

- [Behandeln von Problemen mit Office-Skripts](../testing/troubleshooting.md)
- [Problembehandlungsinformationen für Power Automate mit Office Skripts](../testing/power-automate-troubleshooting.md)
- [Plattformbeschränkungen mit Office Skripts](../testing/platform-limits.md)
- [Verbessern der Leistung Ihrer Office-Skripts](web-client-performance.md)
