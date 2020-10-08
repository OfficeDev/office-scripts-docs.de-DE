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
# <a name="improve-the-performance-of-your-office-scripts"></a>Verbessern der Leistung Ihrer Office-Skripts

Der Zweck von Office-Skripts besteht darin, häufig ausgeführte Aufgabenserien zu automatisieren, um Zeit zu sparen. Ein langsames Skript kann sich so fühlen, als würde es den Workflow nicht beschleunigen. In den meisten Fällen ist Ihr Skript vollständig in Ordnung und wird wie erwartet ausgeführt. Es gibt jedoch einige vermeidbare Szenarien, die sich auf die Leistung auswirken können.

Der häufigste Grund für ein langsames Skript ist eine übermäßige Kommunikation mit der Arbeitsmappe. Ihr Skript wird auf dem lokalen Computer ausgeführt, während die Arbeitsmappe in der Cloud vorhanden ist. Zu bestimmten Zeiten synchronisiert Ihr Skript seine lokalen Daten mit dem der Arbeitsmappe. Dies bedeutet, dass alle Schreibvorgänge (beispielsweise `workbook.addWorksheet()` ) nur dann auf die Arbeitsmappe angewendet werden, wenn diese hinter den Kulissen stattfindende Synchronisierung erfolgt. Ebenso erhalten alle Lesevorgänge (beispielsweise `myRange.getValues()` ) nur Daten aus der Arbeitsmappe für das Skript zu diesen Zeiten. In beiden Fällen ruft das Skript Informationen ab, bevor es auf die Daten zugreift. Mit dem folgenden Code wird beispielsweise die Anzahl der Zeilen im verwendeten Bereich genau protokolliert.

```TypeScript
let usedRange = workbook.getActiveWorksheet().getUsedRange();
let rowCount = usedRange.getRowCount();
// The script will read the range and row count from
// the workbook before logging the information.
console.log(rowCount);
```

Office Scripts-APIs stellen sicher, dass alle Daten in der Arbeitsmappe oder im Skript korrekt sind und bei Bedarf auf dem neuesten Stand sind. Sie müssen sich keine Gedanken über diese Synchronisierungen machen, damit Ihr Skript ordnungsgemäß ausgeführt wird. Das Bewusstsein für diese Kommunikation zwischen Skript und Wolke kann Ihnen jedoch helfen, nicht benötigte Netzwerk Anrufe zu vermeiden.

## <a name="performance-optimizations"></a>Leistungsoptimierungen

Sie können einfache Techniken anwenden, um die Kommunikation mit der Cloud zu verringern. Die folgenden Muster helfen Ihnen dabei, die Skripts zu beschleunigen.

- Lesen von Arbeitsmappendaten einmal statt wiederholt in einer Schleife.
- Entfernen Sie unnötige `console.log` Anweisungen.
- Vermeiden Sie die Verwendung von try/catch-Blöcken.

### <a name="read-workbook-data-outside-of-a-loop"></a>Arbeitsmappen-Daten außerhalb einer Schleife lesen

Jede Methode, die Daten aus der Arbeitsmappe abruft, kann einen Netzwerkaufruf auslösen. Anstatt wiederholt den gleichen Aufruf durchführen, sollten Sie Daten lokal speichern, wann immer möglich. Dies gilt insbesondere beim Umgang mit Schleifen.

Verwenden Sie ein Skript, um die Anzahl der negativen Zahlen im verwendeten Bereich eines Arbeitsblatts abzurufen. Das Skript muss jede Zelle im verwendeten Bereich durchlaufen. Dazu benötigt er den Bereich, die Anzahl der Zeilen und die Anzahl der Spalten. Sie sollten diese als lokale Variablen speichern, bevor Sie mit der Schleife beginnen. Andernfalls wird durch jede Iteration der Schleife eine Rückgabe an die Arbeitsmappe erzwungen.

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
> Versuchen Sie als Experiment `usedRangeValues` in der Schleife mit zu ersetzen `usedRange.getValues()` . Sie können feststellen, dass das Skript beim Umgang mit großen Bereichen wesentlich länger ausgeführt wird.

### <a name="remove-unnecessary-consolelog-statements"></a>Entfernen unnötiger `console.log` Anweisungen

Die Konsolenprotokollierung ist ein wichtiges Tool zum [Debuggen von Skripts](../testing/troubleshooting.md). Es zwingt das Skript jedoch zur Synchronisierung mit der Arbeitsmappe, um sicherzustellen, dass die protokollierten Informationen auf dem neuesten Stand sind. Entfernen Sie nicht benötigte Protokollierungs Anweisungen (beispielsweise die zum Testen verwendeten), bevor Sie Ihr Skript freigeben. Dies führt in der Regel nicht zu einem spürbaren Leistungsproblem, es sei denn, die `console.log()` Anweisung befindet sich in einer Schleife.

### <a name="avoid-using-trycatch-blocks"></a>Vermeiden der Verwendung von try/catch-Blöcken

Es wird nicht empfohlen, [ `try` / `catch` Blöcke](https://developer.mozilla.org/docs/Web/JavaScript/Reference/Statements/try...catch) als Teil der erwarteten Ablaufsteuerung eines Skripts zu verwenden. Die meisten Fehler können vermieden werden, indem die von der Arbeitsmappe zurückgegebenen Objekte überprüft werden. Das folgende Skript überprüft beispielsweise, ob die von der Arbeitsmappe zurückgegebene Tabelle vorhanden ist, bevor versucht wird, eine Zeile hinzuzufügen.

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

## <a name="case-by-case-help"></a>Fall-für-Fall-Hilfe

Wenn die Office-Skriptplattform erweitert wird, um mit [Power Automation](https://flow.microsoft.com/), [Adaptive Cards](https://docs.microsoft.com/adaptive-cards)und anderen produktübergreifenden Features zu arbeiten, werden die Details der Skript-Arbeitsmappen-Kommunikation komplizierter. Wenn Sie Hilfe beim Beschleunigen des Skripts benötigen, erreichen Sie den [Stapelüberlauf](https://stackoverflow.com/questions/tagged/office-scripts). Achten Sie darauf, Ihre Frage mit "Office-Skripts" zu versehen, damit Experten Sie finden und helfen können.

## <a name="see-also"></a>Siehe auch

- [Grundlagen der Skripterstellung für Office-Skripts in Excel im Web](scripting-fundamentals.md)
- [MDN-Webdocs: Loops and Iteration](https://developer.mozilla.org/docs/Web/JavaScript/Guide/Loops_and_iteration)
