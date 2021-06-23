---
title: Verschieben von Zeilen über Tabellen mithilfe von Office Skripts
description: Erfahren Sie, wie Sie Zeilen über Tabellen verschieben, indem Sie Filter speichern und dann die Filter verarbeiten und erneut anwenden.
ms.date: 05/06/2021
localization_priority: Normal
ms.openlocfilehash: c850ed055457f6733694027469a96a87e74ef66a
ms.sourcegitcommit: 4693c8f79428ec74695328275703af0ba1bfea8f
ms.translationtype: MT
ms.contentlocale: de-DE
ms.lasthandoff: 06/23/2021
ms.locfileid: "53074452"
---
# <a name="move-rows-across-tables-by-saving-filters-then-processing-and-reapplying-the-filters"></a><span data-ttu-id="08490-103">Verschieben von Zeilen über Tabellen durch Speichern von Filtern, anschließendes Verarbeiten und erneutes Anwenden der Filter</span><span class="sxs-lookup"><span data-stu-id="08490-103">Move rows across tables by saving filters, then processing and reapplying the filters</span></span>

<span data-ttu-id="08490-104">In diesem Skript werden folgende Schritte ausgeführt:</span><span class="sxs-lookup"><span data-stu-id="08490-104">This script does the following:</span></span>

* <span data-ttu-id="08490-105">Markiert Zeilen aus der Quelltabelle, in denen der Wert in einer Spalte gleich _einem Wert_ ist.</span><span class="sxs-lookup"><span data-stu-id="08490-105">Selects rows from the source table where the value in a column is equal to _some value_.</span></span>
* <span data-ttu-id="08490-106">Verschiebt alle markierten Zeilen in eine andere (Ziel-)Tabelle eines anderen Arbeitsblatts.</span><span class="sxs-lookup"><span data-stu-id="08490-106">Moves all selected rows into another (target) table on another worksheet.</span></span>
* <span data-ttu-id="08490-107">Wendet die relevanten Filter erneut auf die Quelltabelle an.</span><span class="sxs-lookup"><span data-stu-id="08490-107">Reapplies the relevant filters on the source table.</span></span>

:::image type="content" source="../../images/table-filter-before-after.png" alt-text="Screenshots der Arbeitsmappe vor und nach.":::

## <a name="sample-excel-file"></a><span data-ttu-id="08490-109">Beispieldatei für Excel</span><span class="sxs-lookup"><span data-stu-id="08490-109">Sample Excel file</span></span>

<span data-ttu-id="08490-110">Laden Sie die Datei <a href="input-table-filters.xlsx"> herunter, die </a> in dieser Lösung verwendetinput-table-filters.xlsx, um sie selbst auszuprobieren!</span><span class="sxs-lookup"><span data-stu-id="08490-110">Download the file <a href="input-table-filters.xlsx">input-table-filters.xlsx</a> used in this solution to try it out yourself!</span></span>

## <a name="sample-code-move-rows-using-range-values"></a><span data-ttu-id="08490-111">Beispielcode: Verschieben von Zeilen mithilfe von Bereichswerten</span><span class="sxs-lookup"><span data-stu-id="08490-111">Sample code: Move rows using range values</span></span>

```TypeScript
function main(workbook: ExcelScript.Workbook) {

  // You can change these names to match the data in your workbook.
  const TargetTableName = 'Table1';
  const SourceTableName = 'Table2';
  const IndexOfColumnToFilterOn = 1;
  const NameOfColumnToFilterOn = 'Category';
  const ValueToFilterOn = 'Clothing';

  // Get the Table objects.
  let targetTable = workbook.getTable(TargetTableName);
  let sourceTable = workbook.getTable(SourceTableName);

  // If either table is missing, report that information and stop the script.
  if (!targetTable || !sourceTable) {
    console.log(`Tables missing - Check to make sure both source (${TargetTableName}) and target table (${SourceTableName}) are present before running the script. `);
    return;
  }

  // Save the filter criteria.
  const tableFilters = {};
  // For each table column, collect the filter criteria on that column.
  sourceTable.getColumns().forEach((column) => {
    let colFilterCriteria = column.getFilter().getCriteria();
    if (colFilterCriteria) {
      tableFilters[column.getName()] = colFilterCriteria;
    }
  });

  // Get all the data from the table.
  const sourceRange = sourceTable.getRangeBetweenHeaderAndTotal();
  const dataRows: (number | string | boolean)[][] = sourceTable.getRangeBetweenHeaderAndTotal().getValues();

  // Create variables to hold the rows to be moved and their addresses.
  let rowsToMoveValues: (number | string | boolean)[][] = [];
  let rowAddressToRemove: string[] = [];

  // Get the data values from the source table.
  for (let i = 0; i < dataRows.length; i++) { 
    if (dataRows[i][IndexOfColumnToFilterOn] === ValueToFilterOn) {
      rowsToMoveValues.push(dataRows[i]);

      // Get the intersection between table address and the entire row where we found the match. This provides the address of the range to remove.
      let address = sourceRange.getIntersection(sourceRange.getCell(i,0).getEntireRow()).getAddress();
      rowAddressToRemove.push(address);
    }
  }

  // If there are no data rows to process, end the script.
  if (rowsToMoveValues.length < 1) {
    console.log('No rows selected from the source table match the filter criteria.');
    return;
  }

  console.log(`Adding ${rowsToMoveValues.length} rows to target table.`);

  // Insert rows at the end of target table.
  targetTable.addRows(-1, rowsToMoveValues)

  // Remove the rows from the source table.
  const sheet = sourceTable.getWorksheet();

  // Remove all filters before removing rows.
  sourceTable.getAutoFilter().clearCriteria();

  // Important: Remove the rows starting at the bottom of the table.
  // Otherwise, the lower rows change position before they are deleted.
  console.log(`Removing ${rowAddressToRemove.length} rows from the source table.`);
  rowAddressToRemove.reverse().forEach((address) => {
    sheet.getRange(address).delete(ExcelScript.DeleteShiftDirection.up);
  });

  // Reapply the original filters. 
  Object.keys(tableFilters).forEach((columnName) => {
      sourceTable.getColumnByName(columnName).getFilter().apply(tableFilters[columnName]);
    });
}
```

## <a name="training-video-move-rows-across-tables"></a><span data-ttu-id="08490-112">Schulungsvideo: Verschieben von Zeilen über Tabellen</span><span class="sxs-lookup"><span data-stu-id="08490-112">Training video: Move rows across tables</span></span>

<span data-ttu-id="08490-113">[Sehen Sie sich dieses Beispiel auf YouTube an.](https://youtu.be/_3t3Pk4i2L0)</span><span class="sxs-lookup"><span data-stu-id="08490-113">[Watch Sudhi Ramamurthy walk through this sample on YouTube](https://youtu.be/_3t3Pk4i2L0).</span></span> <span data-ttu-id="08490-114">In der Videolösung werden zwei Skripts angezeigt.</span><span class="sxs-lookup"><span data-stu-id="08490-114">There are two scripts shown in the video's solution.</span></span> <span data-ttu-id="08490-115">Der Hauptunterschied besteht darin, wie die Zeilen ausgewählt werden.</span><span class="sxs-lookup"><span data-stu-id="08490-115">The main difference is how the rows are selected.</span></span>

* <span data-ttu-id="08490-116">In der ersten Variante werden die Zeilen ausgewählt, indem der Tabellenfilter angewendet und der sichtbare Bereich gelesen wird.</span><span class="sxs-lookup"><span data-stu-id="08490-116">In the first variant, the rows are selected by applying the table filter and reading the visible range.</span></span>
* <span data-ttu-id="08490-117">In der zweiten werden die Zeilen ausgewählt, indem die Werte gelesen und die Zeilenwerte extrahiert werden (dies wird im Beispiel auf dieser Seite verwendet).</span><span class="sxs-lookup"><span data-stu-id="08490-117">In the second, the rows are selected by reading the values and extracting the row values (which is what the sample on this page uses).</span></span>
