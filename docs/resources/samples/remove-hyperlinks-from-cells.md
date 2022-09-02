---
title: Entfernen von Links aus jeder Zelle in einem Excel-Arbeitsblatt
description: Erfahren Sie, wie Sie mithilfe von Office-Skripts Links aus jeder Zelle eines Excel-Arbeitsblatts entfernen.
ms.date: 06/29/2021
ms.localizationpriority: medium
ms.openlocfilehash: 1445988b1e6a85fcab8914ffeaaef80a07a52f5e
ms.sourcegitcommit: a6504f8b0d6b717457c6e0b5306c35ad3900914e
ms.translationtype: MT
ms.contentlocale: de-DE
ms.lasthandoff: 09/02/2022
ms.locfileid: "67572626"
---
# <a name="remove-hyperlinks-from-each-cell-in-an-excel-worksheet"></a>Entfernen von Links aus jeder Zelle in einem Excel-Arbeitsblatt

 In diesem Beispiel werden alle Hyperlinks aus dem aktuellen Arbeitsblatt gelöscht. Es durchläuft das Arbeitsblatt, und wenn der Zelle ein Hyperlink zugeordnet ist, wird der Hyperlink gelöscht, der Zellwert bleibt jedoch erhalten. Protokolliert außerdem die Zeit, die zum Durchlaufen benötigt wird.

> [!NOTE]
> Dies funktioniert nur, wenn die Zellenanzahl < 10k beträgt.

## <a name="sample-excel-file"></a>Excel-Beispieldatei

Laden Sie die Datei [remove-hyperlinks.xlsx](remove-hyperlinks.xlsx) für eine sofort einsatzbereite Arbeitsmappe herunter. Fügen Sie das folgende Skript hinzu, um das Beispiel selbst auszuprobieren!

## <a name="sample-code-remove-hyperlinks"></a>Beispielcode: Entfernen von Links

```TypeScript
function main(workbook: ExcelScript.Workbook, sheetName: string = 'Sheet1') {
  // Get the active worksheet. 
  let sheet = workbook.getWorksheet(sheetName);

  // Get the used range to operate on.
  // For large ranges (over 10000 entries), consider splitting the operation into batches for performance.
  const targetRange = sheet.getUsedRange(true);
  console.log(`Target Range to clear hyperlinks from: ${targetRange.getAddress()}`);

  const rowCount = targetRange.getRowCount();
  const colCount = targetRange.getColumnCount();
  console.log(`Searching for hyperlinks in ${targetRange.getAddress()} which contains ${(rowCount * colCount)} cells`);

  // Go through each individual cell looking for a hyperlink. 
  // This allows us to limit the formatting changes to only the cells with hyperlink formatting.
  let clearedCount = 0;
  for (let i = 0; i < rowCount; i++) {
    for (let j = 0; j < colCount; j++) {
      const cell = targetRange.getCell(i, j);
      const hyperlink = cell.getHyperlink();
      if (hyperlink) {
        cell.clear(ExcelScript.ClearApplyTo.hyperlinks);
        cell.getFormat().getFont().setUnderline(ExcelScript.RangeUnderlineStyle.none);
        cell.getFormat().getFont().setColor('Black');
        clearedCount++;
      }
    }
  }

  console.log(`Done. Cleared hyperlinks from ${clearedCount} cells`);
}
```

## <a name="training-video-remove-hyperlinks-from-each-cell-in-an-excel-worksheet"></a>Schulungsvideo: Entfernen von Links aus jeder Zelle in einem Excel-Arbeitsblatt

[Sehen Sie sich Sudhi Ramamurthy bei diesem Beispiel auf YouTube an](https://youtu.be/v20fdinxpHU).
