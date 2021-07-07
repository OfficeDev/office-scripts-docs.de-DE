---
title: Entfernen von Hyperlinks aus jeder Zelle in einem Excel Arbeitsblatt
description: Erfahren Sie, wie Sie Office Skripts verwenden, um Hyperlinks aus jeder Zelle in einem Excel Arbeitsblatt zu entfernen.
ms.date: 06/29/2021
localization_priority: Normal
ms.openlocfilehash: dc33eb639edac8ada29824a53440031942e59179
ms.sourcegitcommit: 211c157ca746e266eeb079f5fa1925a1e35ab702
ms.translationtype: MT
ms.contentlocale: de-DE
ms.lasthandoff: 07/07/2021
ms.locfileid: "53313750"
---
# <a name="remove-hyperlinks-from-each-cell-in-an-excel-worksheet"></a><span data-ttu-id="2edac-103">Entfernen von Hyperlinks aus jeder Zelle in einem Excel Arbeitsblatt</span><span class="sxs-lookup"><span data-stu-id="2edac-103">Remove hyperlinks from each cell in an Excel worksheet</span></span>

 <span data-ttu-id="2edac-104">In diesem Beispiel werden alle Hyperlinks aus dem aktuellen Arbeitsblatt gelöscht.</span><span class="sxs-lookup"><span data-stu-id="2edac-104">This sample clears all of the hyperlinks from the current worksheet.</span></span> <span data-ttu-id="2edac-105">Es durchläuft das Arbeitsblatt, und wenn der Zelle ein Hyperlink zugeordnet ist, wird der Hyperlink gelöscht, der Zellwert wird jedoch beibehalten.</span><span class="sxs-lookup"><span data-stu-id="2edac-105">It traverses the worksheet and if there is any hyperlink associated with the cell, it clears the hyperlink yet retains the cell value as is.</span></span> <span data-ttu-id="2edac-106">Protokolliert außerdem die Zeit, die für die Durchführung der Durchquerung benötigt wird.</span><span class="sxs-lookup"><span data-stu-id="2edac-106">Also logs the time it takes to complete traversal.</span></span>

> [!NOTE]
> <span data-ttu-id="2edac-107">Dies funktioniert nur, wenn die Zellenanzahl < 10k beträgt.</span><span class="sxs-lookup"><span data-stu-id="2edac-107">This only works if the cell count is < 10k.</span></span>

## <a name="sample-excel-file"></a><span data-ttu-id="2edac-108">Beispieldatei für Excel</span><span class="sxs-lookup"><span data-stu-id="2edac-108">Sample Excel file</span></span>

<span data-ttu-id="2edac-109">Laden Sie die Datei <a href="remove-hyperlinks.xlsx">remove-hyperlinks.xlsx</a> für eine einsatzbereite Arbeitsmappe herunter.</span><span class="sxs-lookup"><span data-stu-id="2edac-109">Download the file <a href="remove-hyperlinks.xlsx">remove-hyperlinks.xlsx</a> for a ready-to-use workbook.</span></span> <span data-ttu-id="2edac-110">Fügen Sie das folgende Skript hinzu, um das Beispiel selbst auszuprobieren!</span><span class="sxs-lookup"><span data-stu-id="2edac-110">Add the following script to try the sample yourself!</span></span>

## <a name="sample-code-remove-hyperlinks"></a><span data-ttu-id="2edac-111">Beispielcode: Entfernen von Hyperlinks</span><span class="sxs-lookup"><span data-stu-id="2edac-111">Sample code: Remove hyperlinks</span></span>

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

## <a name="training-video-remove-hyperlinks-from-each-cell-in-an-excel-worksheet"></a><span data-ttu-id="2edac-112">Schulungsvideo: Entfernen von Hyperlinks aus jeder Zelle in einem Excel Arbeitsblatt</span><span class="sxs-lookup"><span data-stu-id="2edac-112">Training video: Remove hyperlinks from each cell in an Excel worksheet</span></span>

<span data-ttu-id="2edac-113">[Sehen Sie sich dieses Beispiel auf YouTube an.](https://youtu.be/v20fdinxpHU)</span><span class="sxs-lookup"><span data-stu-id="2edac-113">[Watch Sudhi Ramamurthy walk through this sample on YouTube](https://youtu.be/v20fdinxpHU).</span></span>
