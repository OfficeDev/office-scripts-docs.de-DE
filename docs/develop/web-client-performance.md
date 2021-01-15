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
# <a name="improve-the-performance-of-your-office-scripts"></a>Verbessern der Leistung Ihrer Office-Skripts

Der Zweck von Office Scripts besteht in der Automatisierung häufig ausgeführter Aufgabenreihen, um Zeit zu sparen. Ein langsames Skript kann den Workflow nicht beschleunigen. In den meisten Zeiten ist Ihr Skript völlig in Ordnung und wird wie erwartet ausgeführt. Es gibt jedoch einige vermeidbare Szenarien, die sich auf die Leistung auswirken können.

Der häufigste Grund für ein langsames Skript ist die übermäßige Kommunikation mit der Arbeitsmappe. Ihr Skript wird auf Dem lokalen Computer ausgeführt, während die Arbeitsmappe in der Cloud vorhanden ist. Zu bestimmten Zeiten synchronisiert das Skript seine lokalen Daten mit dem der Arbeitsmappe. Dies bedeutet, dass alle Schreibvorgänge (z. B. ) nur auf die Arbeitsmappe angewendet werden, wenn diese `workbook.addWorksheet()` Hintergrundsynchronisierung erfolgt. Ebenso erhalten alle Lesevorgänge (z. B. ) nur zu diesen Zeiten Daten aus der Arbeitsmappe für `myRange.getValues()` das Skript. In beiden Fällen ruft das Skript Informationen ab, bevor es für die Daten agiert. Mit dem folgenden Code wird beispielsweise die Anzahl der Zeilen im verwendeten Bereich genau protokolliert.

```TypeScript
let usedRange = workbook.getActiveWorksheet().getUsedRange();
let rowCount = usedRange.getRowCount();
// The script will read the range and row count from
// the workbook before logging the information.
console.log(rowCount);
```

Office-Skript-APIs stellen sicher, dass alle Daten in der Arbeitsmappe oder dem Skript bei Bedarf korrekt und aktuell sind. Sie müssen sich keine Gedanken über diese Synchronisierungen machen, damit Ihr Skript ordnungsgemäß ausgeführt wird. Die Kenntnis dieser Skript-zu-Cloud-Kommunikation kann Ihnen jedoch helfen, unnötige Netzwerkanrufe zu vermeiden.

## <a name="performance-optimizations"></a>Leistungsoptimierungen

Sie können einfache Techniken anwenden, um die Kommunikation mit der Cloud zu reduzieren. Die folgenden Muster beschleunigen Ihre Skripts.

- Lesen Sie Arbeitsmappendaten einmal, anstatt wiederholt in einer Schleife.
- Entfernen Sie unnötige `console.log` Anweisungen.
- Vermeiden Sie try/catch-Blöcke.

### <a name="read-workbook-data-outside-of-a-loop"></a>Lesen von Arbeitsmappendaten außerhalb einer Schleife

Jede Methode, die Daten aus der Arbeitsmappe abruft, kann einen Netzwerkaufruf auslösen. Anstatt denselben Aufruf wiederholt zu machen, sollten Sie Daten nach Möglichkeit lokal speichern. Dies gilt insbesondere für Schleifen.

Stellen Sie sich ein Skript vor, um die Anzahl negativer Zahlen im verwendeten Bereich eines Arbeitsblatts zu erhalten. Das Skript muss jede Zelle im verwendeten Bereich durch iterieren. Dazu sind der Bereich, die Anzahl der Zeilen und die Anzahl der Spalten benötigt. Sie sollten diese als lokale Variablen speichern, bevor Sie die Schleife starten. Andernfalls wird bei jeder Iteration der Schleife eine Rückgabe an die Arbeitsmappe erzwingen.

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
> Versuchen Sie als Experiment, `usedRangeValues` in der Schleife durch zu `usedRange.getValues()` ersetzen. Sie werden feststellen, dass die Ausführung des Skripts bei großen Bereichen erheblich länger dauert.

### <a name="remove-unnecessary-consolelog-statements"></a>Entfernen unnötiger `console.log` Anweisungen

Die Konsolenprotokollierung ist ein wichtiges Tool zum [Debuggen Ihrer Skripts.](../testing/troubleshooting.md) Es wird jedoch die Synchronisierung des Skripts mit der Arbeitsmappe erzwingen, um sicherzustellen, dass die protokollierten Informationen auf dem neuesten Stand sind. Ziehen Sie in Betracht, unnötige Protokollierungsanweisungen (z. B. zu Testzwecken) zu entfernen, bevor Sie ihr Skript freigeben. Dies verursacht in der Regel kein erkennbares Leistungsproblem, es sei denn, die `console.log()` Anweisung befindet sich in einer Schleife.

### <a name="avoid-using-trycatch-blocks"></a>Vermeiden der Verwendung von try/catch-Blöcken

Es wird nicht empfohlen, Blöcke [ `try` / `catch` als](https://developer.mozilla.org/docs/Web/JavaScript/Reference/Statements/try...catch) Teil des erwarteten Steuerungsflusses eines Skripts zu verwenden. Die meisten Fehler können vermieden werden, indem Objekte überprüft werden, die von der Arbeitsmappe zurückgegeben werden. Mit dem folgenden Skript wird beispielsweise überprüft, ob die von der Arbeitsmappe zurückgegebene Tabelle vorhanden ist, bevor versucht wird, eine Zeile hinzuzufügen.

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

Wenn die Office -Skript-Plattform erweitert wird, um [mit Power Automate,](https://flow.microsoft.com/) [adaptiven](/adaptive-cards)Karten und anderen produktübergreifenden Features zu arbeiten, werden die Details der Skriptarbeitsmappenkommunikation komplizierter. Wenn Sie Hilfe benötigen, damit Ihr Skript schneller ausgeführt wird, wenden Sie sich an [Stack Overflow.](https://stackoverflow.com/questions/tagged/office-scripts) Markieren Sie Ihre Frage unbedingt mit "Office-Skripts", damit Experten sie finden und Hilfe erhalten können.

## <a name="see-also"></a>Siehe auch

- [Grundlagen der Skripterstellung für Office-Skripts in Excel im Web](scripting-fundamentals.md)
- [MDN-Web-Dokumente: Schleifen und Iteration](https://developer.mozilla.org/docs/Web/JavaScript/Guide/Loops_and_iteration)