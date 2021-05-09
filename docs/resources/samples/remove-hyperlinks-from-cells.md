---
title: Entfernen von Hyperlinks aus jeder Zelle in Excel Arbeitsblatt
description: Erfahren Sie, wie Sie Office Skripts verwenden, um Hyperlinks aus jeder Zelle in einem arbeitsblatt Excel entfernen.
ms.date: 05/06/2021
localization_priority: Normal
ms.openlocfilehash: 048d01691377a7086bdba9ceb87ca98839cfa4d1
ms.sourcegitcommit: 763d341857bcb209b2f2c278a82fdb63d0e18f0a
ms.translationtype: MT
ms.contentlocale: de-DE
ms.lasthandoff: 05/08/2021
ms.locfileid: "52285801"
---
# <a name="remove-hyperlinks-from-each-cell-in-an-excel-worksheet"></a>Entfernen von Hyperlinks aus jeder Zelle in Excel Arbeitsblatt

 In diesem Beispiel werden alle Hyperlinks aus dem aktuellen Arbeitsblatt entfernt. Es durchläuft das Arbeitsblatt, und wenn der Zelle ein Hyperlink zugeordnet ist, wird der Hyperlink geräumt, der Zellwert wird jedoch beibehalten. Protokolliert außerdem die Zeit, die zum Abschließen der Durchquerung benötigt wird.

> [!NOTE]
> Dies funktioniert nur, wenn die Zellenanzahl < 10k beträgt.

## <a name="sample-code-remove-hyperlinks"></a>Beispielcode: Entfernen von Hyperlinks

Laden Sie die Datei <a href="remove-hyperlinks.xlsx">remove-hyperlinks.xlsx</a> in diesem Beispiel verwendeten herunter, und testen Sie sie selbst!

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

## <a name="training-video-remove-hyperlinks-from-each-cell-in-an-excel-worksheet"></a>Schulungsvideo: Entfernen von Hyperlinks aus jeder Zelle in einem Excel Arbeitsblatt

[Sehen Sie sich an, wie Sudhi Ramamurthy dieses Beispiel auf YouTube durchspazieren.](https://youtu.be/v20fdinxpHU)
