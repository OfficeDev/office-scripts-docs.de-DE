---
title: Leistungsoptimierung beim Schreiben eines großen Datasets
description: Erfahren Sie, wie Sie die Leistung beim Schreiben eines großen Datasets in skripts Office optimieren.
ms.date: 04/28/2021
localization_priority: Normal
ms.openlocfilehash: 9622494378a24db16ea43b5500d6efa156726ff8
ms.sourcegitcommit: 763d341857bcb209b2f2c278a82fdb63d0e18f0a
ms.translationtype: MT
ms.contentlocale: de-DE
ms.lasthandoff: 05/08/2021
ms.locfileid: "52285948"
---
# <a name="performance-optimization-when-writing-a-large-dataset"></a>Leistungsoptimierung beim Schreiben eines großen Datasets

## <a name="basic-performance-optimization"></a>Grundlegende Leistungsoptimierung

Grundlegende Informationen zur Leistung in Office Finden Sie im Abschnitt [Leistung](getting-started.md#basic-performance-considerations) im Artikel Erste Schritte.

## <a name="sample-code-optimize-performance-of-a-large-dataset"></a>Beispielcode: Optimieren der Leistung eines großen Datasets

Die `setValues()` Range-API ermöglicht das Festlegen der Werte eines Bereichs. Diese API hat Dateneinschränkungen in Abhängigkeit von verschiedenen Faktoren wie Datengröße, Netzwerkeinstellungen usw. Um einen großen Datenbereich zuverlässig zu aktualisieren, müssen Sie darüber nachdenken, Datenupdates in kleineren Abschnitten zu machen. Dieses Skript versucht, dies zu tun, und schreibt Zeilen eines Bereichs in Blöcke, sodass dies in kleineren Teilen möglich ist, wenn ein großer Bereich aktualisiert werden muss. **Warnung**: Es wurde nicht in verschiedenen Größen getestet, beachten Sie dies, wenn Sie dies in Ihrem Skript verwenden möchten. Da wir Gelegenheit zum Testen haben, werden wir mit Erkenntnissen darüber aktualisieren, wie sie für verschiedene Datengrößen funktioniert.

Dieses Skript wählt 1K-Zellen pro Block aus, sie können jedoch außer Kraft setzen, um zu testen, wie es für Sie funktioniert. Es aktualisiert 100-k-Zeilen mit 6 Datenspalten. Führen Sie dies auf einem leeren Blatt aus, das sie untersuchen soll.

```TypeScript
function main(workbook: ExcelScript.Workbook) {
  const sheet = workbook.getActiveWorksheet();

  let data: (string | number | boolean)[][] = [];
  // Number of rows in the random data (x 6 columns).
  const sampleRows = 100000;

  console.log(`Generating data...`)
  // Dynamically generate some random data for testing purpose. 
  for (let i = 0; i < sampleRows; i++) {
    data.push([i, ...[getRandomString(5), getRandomString(20), getRandomString(10), Math.random()], "Sample data"]);
  }

  console.log(`Calling update range function...`);
  const updated = updateRangeInChunks(sheet.getRange("B2"), data);
  if (!updated) {
    console.log(`Update did not take place or complete. Check and run again.`);
  }
}

function updateRangeInChunks(
  startCell: ExcelScript.Range,
  values: (string | boolean | number)[][],
  cellsInChunk: number = 10000
): boolean {

  const startTime = new Date().getTime();
  console.log(`Cells per chunk setting: ${cellsInChunk}`);
  if (!values) {
    console.log(`Invalid input values to update.`);
    return false;
  }
  if (values.length === 0 || values[0].length === 0) {
    console.log(`Empty data -- nothing to update.`);
    return true;
  }
  const totalCells = values.length * values[0].length;

  console.log(`Total cells to update in the target range: ${totalCells}`);
  if (totalCells <= cellsInChunk) {
    console.log(`No need to chunk -- updating directly`);
    updateTargetRange(startCell, values);
    return true;
  }

  const rowsPerChunk = Math.floor(cellsInChunk / values[0].length);
  console.log("Rows per chunk: " + rowsPerChunk);
  let rowCount = 0;
  let totalRowsUpdated = 0;
  let chunkCount = 0;

  for (let i = 0; i < values.length; i++) {
    rowCount++;
    if (rowCount === rowsPerChunk) {
      chunkCount++;
      console.log(`Calling update next chunk function. Chunk#: ${chunkCount}`);
      updateNextChunk(startCell, values, rowsPerChunk, totalRowsUpdated);
      rowCount = 0;
      totalRowsUpdated += rowsPerChunk;
      console.log(`${((totalRowsUpdated / values.length) * 100).toFixed(1)}% Done`);

    }
  }
  console.log(`Updating remaining rows -- last chunk: ${rowCount}`)
  if (rowCount > 0) {
    updateNextChunk(startCell, values, rowCount, totalRowsUpdated);
  }

  let endTime = new Date().getTime();
  console.log(`Completed ${totalCells} cells update. It took: ${((endTime - startTime) / 1000).toFixed(6)} seconds to complete. ${((((endTime  - startTime) / 1000)) / cellsInChunk).toFixed(8)} seconds per ${cellsInChunk} cells-chunk.`);

  return true;
}

/**
 * A helper function that computes the target range and updates. 
 */

function updateNextChunk(
  startingCell: ExcelScript.Range,
  data: (string | boolean | number)[][],
  rowsPerChunk: number,
  totalRowsUpdated: number
) {

  const newStartCell = startingCell.getOffsetRange(totalRowsUpdated, 0);
  const targetRange = newStartCell.getResizedRange(rowsPerChunk - 1, data[0].length - 1);
  console.log(`Updating chunk at range ${targetRange.getAddress()}`);
  const dataToUpdate = data.slice(totalRowsUpdated, totalRowsUpdated + rowsPerChunk);
  try {
    targetRange.setValues(dataToUpdate);
  } catch (e) {
    throw `Error while updating the chunk range: ${JSON.stringify(e)}`;
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

## <a name="training-video-optimize-performance-when-writing-a-large-dataset"></a>Schulungsvideo: Optimieren der Leistung beim Schreiben eines großen Datasets

[Sehen Sie sich an, wie Sudhi Ramamurthy dieses Beispiel auf YouTube durchspazieren.](https://youtu.be/BP9Kp0Ltj7U)
