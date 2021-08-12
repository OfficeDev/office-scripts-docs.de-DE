---
title: Entfernen von Hyperlinks aus jeder Zelle in einem Excel Arbeitsblatt
description: Erfahren Sie, wie Sie Office Skripts verwenden, um Hyperlinks aus jeder Zelle in einem Excel Arbeitsblatt zu entfernen.
ms.date: 06/29/2021
localization_priority: Normal
ms.openlocfilehash: 498d55ea1ee7926ab124d00795825660005c5e38e73ed5d90fe8f9208a583908
ms.sourcegitcommit: 75f7ed8c2d23a104acc293f8ce29ea580b4fcdc5
ms.translationtype: MT
ms.contentlocale: de-DE
ms.lasthandoff: 08/11/2021
ms.locfileid: "57847429"
---
# <a name="remove-hyperlinks-from-each-cell-in-an-excel-worksheet"></a>Entfernen von Hyperlinks aus jeder Zelle in einem Excel Arbeitsblatt

 In diesem Beispiel werden alle Hyperlinks aus dem aktuellen Arbeitsblatt gelöscht. Es durchläuft das Arbeitsblatt, und wenn der Zelle ein Hyperlink zugeordnet ist, wird der Hyperlink gelöscht, der Zellwert wird jedoch beibehalten. Protokolliert außerdem die Zeit, die für die Durchführung der Durchquerung benötigt wird.

> [!NOTE]
> Dies funktioniert nur, wenn die Zellenanzahl < 10k beträgt.

## <a name="sample-excel-file"></a>Beispieldatei für Excel

Laden Sie die Datei <a href="remove-hyperlinks.xlsx">remove-hyperlinks.xlsx</a> für eine einsatzbereite Arbeitsmappe herunter. Fügen Sie das folgende Skript hinzu, um das Beispiel selbst auszuprobieren!

## <a name="sample-code-remove-hyperlinks"></a>Beispielcode: Entfernen von Hyperlinks

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

[Sehen Sie sich dieses Beispiel auf YouTube an.](https://youtu.be/v20fdinxpHU)
