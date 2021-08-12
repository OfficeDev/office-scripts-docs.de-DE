---
title: Erstellen eines großen Datasets
description: Erfahren Sie, wie Sie ein großes Dataset in kleinere Schreibvorgänge in Office Skripts aufteilen.
ms.date: 05/13/2021
localization_priority: Normal
ms.openlocfilehash: ec058bb1e3aabea52b0127c26b4c4a1352914fb575105ff4a90f82b222d6251d
ms.sourcegitcommit: 75f7ed8c2d23a104acc293f8ce29ea580b4fcdc5
ms.translationtype: MT
ms.contentlocale: de-DE
ms.lasthandoff: 08/11/2021
ms.locfileid: "57847584"
---
# <a name="write-a-large-dataset"></a>Erstellen eines großen Datasets

Die `Range.setValues()` API platziert Daten in einen Bereich. Diese API weist Einschränkungen in Abhängigkeit von verschiedenen Faktoren auf, z. B. Datengröße und Netzwerkeinstellungen. Dies bedeutet, dass Sie, wenn Sie versuchen, eine große Menge an Informationen in eine Arbeitsmappe als einen einzigen Vorgang zu schreiben, die Daten in kleinere Batches schreiben müssen, um einen [großen Bereich](../../testing/platform-limits.md)zuverlässig zu aktualisieren.

Informationen zu den Leistungsgrundlagen in Office Skripts finden Sie unter [Verbessern der Leistung Ihrer Office Skripts.](../../develop/web-client-performance.md)

## <a name="sample-code-write-a-large-dataset"></a>Beispielcode: Schreiben eines großen Datasets

Dieses Skript schreibt Zeilen eines Bereichs in kleinere Teile. Es wählt jeweils 1000 Zellen aus, die gleichzeitig geschrieben werden sollen. Führen Sie das Skript auf einem leeren Arbeitsblatt aus, um die Aktualisierungsbatches in Aktion zu sehen. Die Konsolenausgabe bietet einen weiteren Einblick in die Vorgänge.

> [!NOTE]
> Sie können die Anzahl der geschriebenen Zeilen ändern, indem Sie den Wert von `SAMPLE_ROWS` ändern. Sie können die Anzahl der zu schreibenden Zellen als einzelne Aktion ändern, indem Sie den Wert von `CELLS_IN_BATCH` ändern.

```TypeScript
function main(workbook: ExcelScript.Workbook) {
  const SAMPLE_ROWS = 100000;
  const CELLS_IN_BATCH = 10000;

  // Get the current worksheet.
  const sheet = workbook.getActiveWorksheet();

  console.log(`Generating data...`)
  let data: (string | number | boolean)[][] = [];
  // Generate six columns of random data per row. 
  for (let i = 0; i < SAMPLE_ROWS; i++) {
    data.push([i, ...[getRandomString(5), getRandomString(20), getRandomString(10), Math.random()], "Sample data"]);
  }

  console.log(`Calling update range function...`);
  const updated = updateRangeInBatches(sheet.getRange("B2"), data, CELLS_IN_BATCH);
  if (!updated) {
    console.log(`Update did not take place or complete. Check and run again.`);
  }
}

function updateRangeInBatches(
  startCell: ExcelScript.Range,
  values: (string | boolean | number)[][],
  cellsInBatch: number
): boolean {

  const startTime = new Date().getTime();
  console.log(`Cells per batch setting: ${cellsInBatch}`);

  // Determine the total number of cells to write.
  const totalCells = values.length * values[0].length;
  console.log(`Total cells to update in the target range: ${totalCells}`);
  if (totalCells <= cellsInBatch) {
    console.log(`No need to batch -- updating directly`);
    updateTargetRange(startCell, values);
    return true;
  }

  // Determine how many rows to write at once.
  const rowsPerBatch = Math.floor(cellsInBatch / values[0].length);
  console.log("Rows per batch: " + rowsPerBatch);
  let rowCount = 0;
  let totalRowsUpdated = 0;
  let batchCount = 0;

  // Write each batch of rows.
  for (let i = 0; i < values.length; i++) {
    rowCount++;
    if (rowCount === rowsPerBatch) {
      batchCount++;
      console.log(`Calling update next batch function. Batch#: ${batchCount}`);
      updateNextBatch(startCell, values, rowsPerBatch, totalRowsUpdated);

      // Write a completion percentage to help the user understand the progress.
      rowCount = 0;
      totalRowsUpdated += rowsPerBatch;
      console.log(`${((totalRowsUpdated / values.length) * 100).toFixed(1)}% Done`);
    }
  }
  
  console.log(`Updating remaining rows -- last batch: ${rowCount}`)
  if (rowCount > 0) {
    updateNextBatch(startCell, values, rowCount, totalRowsUpdated);
  }

  let endTime = new Date().getTime();
  console.log(`Completed ${totalCells} cells update. It took: ${((endTime - startTime) / 1000).toFixed(6)} seconds to complete. ${((((endTime  - startTime) / 1000)) / cellsInBatch).toFixed(8)} seconds per ${cellsInBatch} cells-batch.`);

  return true;
}

/**
 * A helper function that computes the target range and updates. 
 */
function updateNextBatch(
  startingCell: ExcelScript.Range,
  data: (string | boolean | number)[][],
  rowsPerBatch: number,
  totalRowsUpdated: number
) {
  const newStartCell = startingCell.getOffsetRange(totalRowsUpdated, 0);
  const targetRange = newStartCell.getResizedRange(rowsPerBatch - 1, data[0].length - 1);
  console.log(`Updating batch at range ${targetRange.getAddress()}`);
  const dataToUpdate = data.slice(totalRowsUpdated, totalRowsUpdated + rowsPerBatch);
  try {
    targetRange.setValues(dataToUpdate);
  } catch (e) {
    throw `Error while updating the batch range: ${JSON.stringify(e)}`;
  }
  return;
}

/**
 * A helper function that computes the target range given the target range's starting cell
 * and selected range and updates the values.
 */
function updateTargetRange(
  targetCell: ExcelScript.Range,
  values: (string | boolean | number)[][]
) {
  const targetRange = targetCell.getResizedRange(values.length - 1, values[0].length - 1);
  console.log(`Updating the range: ${targetRange.getAddress()}`);
  try {
    targetRange.setValues(values);
  } catch (e) {
    throw `Error while updating the whole range: ${JSON.stringify(e)}`;
  }
  return;
}

// Credit: https://www.codegrepper.com/code-examples/javascript/random+text+generator+javascript
function getRandomString(length: number): string {
  var randomChars = 'ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789';
  var result = '';
  for (var i = 0; i < length; i++) {
    result += randomChars.charAt(Math.floor(Math.random() * randomChars.length));
  }
  return result;
}
```

## <a name="training-video-write-a-large-dataset"></a>Schulungsvideo: Schreiben eines großen Datasets

[Sehen Sie sich dieses Beispiel auf YouTube an.](https://youtu.be/BP9Kp0Ltj7U)
