---
title: Verbessern der Leistung Ihrer Office-Skripts
description: Erstellen Sie schnellere Skripts, indem Sie die Kommunikation zwischen der Excel Arbeitsmappe und Ihrem Skript verstehen.
ms.date: 05/17/2021
ms.localizationpriority: medium
ms.openlocfilehash: 169256bdae809c413c10f1f00240afc28be795f4
ms.sourcegitcommit: d3ed4bdeeba805d97c930394e172e8306a0cf484
ms.translationtype: MT
ms.contentlocale: de-DE
ms.lasthandoff: 09/15/2021
ms.locfileid: "59331164"
---
# <a name="improve-the-performance-of-your-office-scripts"></a>Verbessern der Leistung Ihrer Office-Skripts

Der Zweck von Office Skripts besteht darin, häufig ausgeführte Aufgabenreihen zu automatisieren, um Zeit zu sparen. Ein langsames Skript kann den Eindruck haben, dass es den Workflow nicht beschleunigt. In den meisten Jahren ist Ihr Skript einwandfrei und wird erwartungsgemäß ausgeführt. Es gibt jedoch einige verwendbare Szenarien, die sich auf die Leistung auswirken können.

Der häufigste Grund für ein langsames Skript ist eine übermäßige Kommunikation mit der Arbeitsmappe. Ihr Skript wird auf Ihrem lokalen Computer ausgeführt, während die Arbeitsmappe in der Cloud vorhanden ist. Zu bestimmten Zeiten synchronisiert das Skript die lokalen Daten mit den Daten der Arbeitsmappe. Dies bedeutet, dass alle Schreibvorgänge (z. `workbook.addWorksheet()` B. ) nur auf die Arbeitsmappe angewendet werden, wenn diese Synchronisierung im Hintergrund erfolgt. Ebenso rufen alle Lesevorgänge (z. `myRange.getValues()` B. ) nur zu diesen Zeiten Daten aus der Arbeitsmappe für das Skript ab. In beiden Fällen ruft das Skript Informationen ab, bevor es auf die Daten reagiert. Beispielsweise protokolliert der folgende Code die Anzahl der Zeilen im verwendeten Bereich genau.

```TypeScript
let usedRange = workbook.getActiveWorksheet().getUsedRange();
let rowCount = usedRange.getRowCount();
// The script will read the range and row count from
// the workbook before logging the information.
console.log(rowCount);
```

Office Skript-APIs stellen sicher, dass alle Daten in der Arbeitsmappe oder im Skript korrekt und bei Bedarf auf dem neuesten Stand sind. Sie müssen sich keine Gedanken über diese Synchronisierungen machen, damit Ihr Skript ordnungsgemäß ausgeführt wird. Ein Bewusstsein für diese Script-to-Cloud-Kommunikation kann Ihnen jedoch dabei helfen, nicht benötigte Netzwerkanrufe zu vermeiden.

## <a name="performance-optimizations"></a>Leistungsoptimierungen

Sie können einfache Techniken anwenden, um die Kommunikation mit der Cloud zu reduzieren. Mit den folgenden Mustern können Sie Ihre Skripts beschleunigen.

- Arbeitsmappendaten einmal lesen, anstatt sich wiederholt in einer Schleife zu befinden.
- Entfernen Sie unnötige `console.log` Anweisungen.
- Vermeiden Sie die Verwendung von Try/Catch-Blöcken.

### <a name="read-workbook-data-outside-of-a-loop"></a>Lesen von Arbeitsmappendaten außerhalb einer Schleife

Jede Methode, die Daten aus der Arbeitsmappe abruft, kann einen Netzwerkaufruf auslösen. Anstatt wiederholt denselben Aufruf durchzuführen, sollten Sie Daten nach Möglichkeit lokal speichern. Dies gilt insbesondere für Schleifen.

Erwägen Sie ein Skript, um die Anzahl negativer Zahlen im verwendeten Bereich eines Arbeitsblatts abzurufen. Das Skript muss jede Zelle im verwendeten Bereich durchlaufen. Dazu werden der Bereich, die Anzahl der Zeilen und die Anzahl der Spalten benötigt. Sie sollten diese als lokale Variablen speichern, bevor Sie die Schleife starten. Andernfalls erzrückt jede Iteration der Schleife eine Rückgabe an die Arbeitsmappe.

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
> Versuchen Sie als Experiment, `usedRangeValues` in der Schleife durch `usedRange.getValues()` . Sie werden feststellen, dass die Ausführung des Skripts bei großen Bereichen erheblich länger dauert.

### <a name="avoid-using-trycatch-blocks-in-or-surrounding-loops"></a>Vermeiden der Verwendung von `try...catch` Blöcken in oder umgebenden Schleifen

Es wird nicht empfohlen, [`try...catch`](https://developer.mozilla.org/docs/Web/JavaScript/Reference/Statements/try...catch) Anweisungen in Schleifen oder umgebenden Schleifen zu verwenden. Aus dem gleichen Grund sollten Sie das Lesen von Daten in einer Schleife vermeiden: Jede Iteration erzwingt die Synchronisierung des Skripts mit der Arbeitsmappe, um sicherzustellen, dass kein Fehler ausgelöst wurde. Die meisten Fehler können vermieden werden, indem von der Arbeitsmappe zurückgegebene Objekte überprüft werden. Das folgende Skript überprüft beispielsweise, ob die von der Arbeitsmappe zurückgegebene Tabelle vorhanden ist, bevor versucht wird, eine Zeile hinzuzufügen.

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

Die Konsolenprotokollierung ist ein wichtiges Tool zum [Debuggen ihrer Skripts.](../testing/troubleshooting.md) Allerdings wird erzwungen, dass das Skript mit der Arbeitsmappe synchronisiert wird, um sicherzustellen, dass die protokollierten Informationen auf dem neuesten Stand sind. Ziehen Sie in Erwägung, unnötige Protokollierungsanweisungen (z. B. die zu Testzwecken verwendeten) zu entfernen, bevor Sie Ihr Skript freigeben. Dies führt in der Regel nicht zu einem erheblichen Leistungsproblem, es sei denn, die `console.log()` Anweisung befindet sich in einer Schleife.

## <a name="case-by-case-help"></a>Fall-für-Fall-Hilfe

Wenn die Plattform Office Skripts erweitert wird, um mit [Power Automate,](https://flow.microsoft.com/) [adaptiven Karten](/adaptive-cards)und anderen produktübergreifenden Features zu arbeiten, werden die Details der Kommunikation zwischen Skripts und Arbeitsmappen komplizierter. Wenn Sie Hilfe benötigen, damit Ihr Skript schneller ausgeführt wird, wenden Sie sich an [Microsoft Q&A](/answers/topics/office-scripts-excel-dev.html). Achten Sie darauf, Ihre Frage mit "office-scripts-dev" zu markieren, damit Experten sie finden und Hilfe erhalten.

## <a name="see-also"></a>Siehe auch

- [Grundlagen der Skripterstellung für Office-Skripts in Excel im Web](scripting-fundamentals.md)
- [MDN-Webdokumentation: Schleifen und Iteration](https://developer.mozilla.org/docs/Web/JavaScript/Guide/Loops_and_iteration)
