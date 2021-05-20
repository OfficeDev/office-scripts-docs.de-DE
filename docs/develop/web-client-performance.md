---
title: Verbessern Sie die Leistung Ihrer Office Scripts
description: Erstellen Sie schnellere Skripts, indem Sie die Kommunikation zwischen der arbeitsmappe Excel und Ihrem Skript verstehen.
ms.date: 05/17/2021
localization_priority: Normal
ms.openlocfilehash: 512e2108cb81cf9ac8ae98980951d5d01b3d2de9
ms.sourcegitcommit: 4687693f02fc90a57ba30c461f35046e02e6f5fb
ms.translationtype: MT
ms.contentlocale: de-DE
ms.lasthandoff: 05/19/2021
ms.locfileid: "52544991"
---
# <a name="improve-the-performance-of-your-office-scripts"></a>Verbessern Sie die Leistung Ihrer Office Scripts

Der Zweck von Office Scripts besteht darin, häufig ausgeführte Aufgabenserien zu automatisieren, um Ihnen Zeit zu sparen. Ein langsames Skript kann sich so anfühlen, als würde es Ihren Workflow nicht beschleunigen. Meistens ist Ihr Skript vollkommen in Ordnung und wird wie erwartet ausgeführt. Es gibt jedoch einige vermeidbare Szenarien, die sich auf die Leistung auswirken können.

Der häufigste Grund für ein langsames Skript ist eine übermäßige Kommunikation mit der Arbeitsmappe. Das Skript wird auf dem lokalen Computer ausgeführt, während die Arbeitsmappe in der Cloud vorhanden ist. Zu bestimmten Zeiten synchronisiert das Skript seine lokalen Daten mit denen der Arbeitsmappe. Dies bedeutet, dass alle Schreibvorgänge (z. B. `workbook.addWorksheet()` ) nur dann auf die Arbeitsmappe angewendet werden, wenn diese Synchronisierung hinter den Kulissen stattfindet. Ebenso erhalten alle Lesevorgänge (z. B. `myRange.getValues()` ) zu diesem Zeitpunkt nur Daten aus der Arbeitsmappe für das Skript. In beiden Fällen ruft das Skript Informationen ab, bevor es auf die Daten einwirkt. Der folgende Code protokolliert z. B. die Anzahl der Zeilen im verwendeten Bereich genau.

```TypeScript
let usedRange = workbook.getActiveWorksheet().getUsedRange();
let rowCount = usedRange.getRowCount();
// The script will read the range and row count from
// the workbook before logging the information.
console.log(rowCount);
```

Office Skript-APIs stellen sicher, dass alle Daten in der Arbeitsmappe oder im Skript korrekt und aktuell sind, wenn dies erforderlich ist. Sie müssen sich keine Gedanken über diese Synchronisierungen machen, damit Ihr Skript ordnungsgemäß ausgeführt werden kann. Ein Bewusstsein für diese Skript-zu-Cloud-Kommunikation kann Ihnen jedoch helfen, unnötige Netzwerkaufrufe zu vermeiden.

## <a name="performance-optimizations"></a>Leistungsoptimierungen

Sie können einfache Techniken anwenden, um die Kommunikation mit der Cloud zu reduzieren. Die folgenden Muster helfen Ihnen dabei, Ihre Skripts zu beschleunigen.

- Lesen Sie Arbeitsmappendaten einmal statt wiederholt in einer Schleife.
- Entfernen Sie unnötige `console.log` Anweisungen.
- Vermeiden Sie die Verwendung von try/catch-Blöcken.

### <a name="read-workbook-data-outside-of-a-loop"></a>Lesen von Arbeitsmappendaten außerhalb einer Schleife

Jede Methode, die Daten aus der Arbeitsmappe abruft, kann einen Netzwerkaufruf auslösen. Anstatt wiederholt denselben Aufruf zu tätigen, sollten Sie Daten nach Möglichkeit lokal speichern. Dies gilt insbesondere für den Umgang mit Schleifen.

Betrachten Sie ein Skript, um die Anzahl negativer Zahlen im verwendeten Bereich eines Arbeitsblatts abzubekommen. Das Skript muss über jede Zelle im verwendeten Bereich iterieren. Dazu benötigt es den Bereich, die Anzahl der Zeilen und die Anzahl der Spalten. Sie sollten diese als lokale Variablen speichern, bevor Sie die Schleife starten. Andernfalls erzwingt jede Iteration der Schleife eine Rückkehr zur Arbeitsmappe.

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
> Versuchen Sie als Experiment, `usedRangeValues` in der Schleife durch zu `usedRange.getValues()` ersetzen. Sie können feststellen, dass die Ausführung des Skripts erheblich länger dauert, wenn es um große Bereiche geht.

### <a name="avoid-using-trycatch-blocks-in-or-surrounding-loops"></a>Vermeiden Sie die Verwendung von `try...catch` Blöcken in oder umgebenden Schleifen

Es wird nicht empfohlen, Anweisungen entweder in Schleifen oder umgebenden Schleifen zu [`try...catch`](https://developer.mozilla.org/docs/Web/JavaScript/Reference/Statements/try...catch) verwenden. Dies ist aus dem gleichen Grund, aus dem Sie das Lesen von Daten in einer Schleife vermeiden sollten: Jede Iteration zwingt das Skript, mit der Arbeitsmappe zu synchronisieren, um sicherzustellen, dass kein Fehler ausgelöst wurde. Die meisten Fehler können vermieden werden, indem Objekte überprüft werden, die von der Arbeitsmappe zurückgegeben werden. Das folgende Skript überprüft beispielsweise, ob die von der Arbeitsmappe zurückgegebene Tabelle vorhanden ist, bevor versucht wird, eine Zeile hinzuzufügen.

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

### <a name="remove-unnecessary-consolelog-statements"></a>Entfernen unnötiger `console.log` Anweisungen

Die Konsolenprotokollierung ist ein wichtiges Tool zum [Debuggen Ihrer Skripts](../testing/troubleshooting.md). Es erzwingt jedoch, dass das Skript mit der Arbeitsmappe synchronisiert wird, um sicherzustellen, dass die protokollierten Informationen auf dem neuesten Stand sind. Entfernen Sie unnötige Protokollierungsanweisungen (z. B. zum Testen), bevor Sie Ihr Skript freigeben. Dies verursacht in der Regel kein auffälliges Leistungsproblem, es sei denn, die `console.log()` Anweisung befindet sich in einer Schleife.

## <a name="case-by-case-help"></a>Fall-für-Fall-Hilfe

Wenn die Office Scripts-Plattform erweitert wird, um mit [Power Automate](https://flow.microsoft.com/), [Adaptive Cards](/adaptive-cards)und anderen produktübergreifenden Funktionen zu arbeiten, werden die Details der Skriptarbeitsmappenkommunikation komplizierter. Wenn Sie Hilfe benötigen, damit Ihr Skript schneller ausgeführt wird, wenden Sie sich bitte an [Microsoft Q&A](/answers/topics/office-scripts-dev.html). Markieren Sie Ihre Frage mit "office-scripts-dev", damit Experten sie finden und helfen können.

## <a name="see-also"></a>Siehe auch

- [Grundlagen der Skripterstellung für Office-Skripts in Excel im Web](scripting-fundamentals.md)
- [MDN-Webdokumente: Schleifen und Iteration](https://developer.mozilla.org/docs/Web/JavaScript/Guide/Loops_and_iteration)
