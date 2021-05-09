---
title: Löschen des Tabellenspaltenfilters basierend auf der aktiven Zellenposition
description: Erfahren Sie, wie Sie den Tabellenspaltenfilter basierend auf der aktiven Zellenposition löschen.
ms.date: 05/03/2021
localization_priority: Normal
ms.openlocfilehash: 06fba191a79f4641d4d1017bda332c7559b50e6d
ms.sourcegitcommit: 763d341857bcb209b2f2c278a82fdb63d0e18f0a
ms.translationtype: MT
ms.contentlocale: de-DE
ms.lasthandoff: 05/08/2021
ms.locfileid: "52285927"
---
# <a name="clear-table-column-filter-based-on-active-cell-location"></a><span data-ttu-id="423aa-103">Löschen des Tabellenspaltenfilters basierend auf der aktiven Zellenposition</span><span class="sxs-lookup"><span data-stu-id="423aa-103">Clear table column filter based on active cell location</span></span>

<span data-ttu-id="423aa-104">In diesem Beispiel wird der Tabellenspaltenfilter basierend auf der aktiven Zellenposition entfernt.</span><span class="sxs-lookup"><span data-stu-id="423aa-104">This sample clears the table column filter based on the active cell location.</span></span> <span data-ttu-id="423aa-105">Das Skript erkennt, ob die Zelle Teil einer Tabelle ist, bestimmt die Tabellenspalte und filtert, die darauf angewendet werden.</span><span class="sxs-lookup"><span data-stu-id="423aa-105">The script detects if the cell is part of a table, determines the table column, and clears any filter that are applied on it.</span></span>

<span data-ttu-id="423aa-106">Wenn Sie mehr darüber erfahren möchten, wie Sie den Filter speichern, bevor Sie ihn löschen (und später erneut anwenden), finden Sie weitere Informationen unter [Move rows across tables by saving filters](move-rows-across-tables.md), a more advanced sample.</span><span class="sxs-lookup"><span data-stu-id="423aa-106">If you wish to learn more about how to save the filter prior to clearing it (and re-apply later), see [Move rows across tables by saving filters](move-rows-across-tables.md), a more advanced sample.</span></span>

<span data-ttu-id="423aa-107">_Vor dem Löschen des Spaltenfilters (beachten Sie die aktive Zelle)_</span><span class="sxs-lookup"><span data-stu-id="423aa-107">_Before clearing column filter (notice the active cell)_</span></span>

:::image type="content" source="../../images/before-filter-applied.png" alt-text="Eine aktive Zelle vor dem Löschen des Spaltenfilters":::

<span data-ttu-id="423aa-109">_Nach dem Löschen des Spaltenfilters_</span><span class="sxs-lookup"><span data-stu-id="423aa-109">_After clearing column filter_</span></span>

:::image type="content" source="../../images/after-filter-cleared.png" alt-text="Eine aktive Zelle nach dem Löschen des Spaltenfilters":::

## <a name="sample-code-clear-table-column-filter-based-on-active-cell"></a><span data-ttu-id="423aa-111">Beispielcode: Tabellenspaltenfilter basierend auf aktiver Zelle löschen</span><span class="sxs-lookup"><span data-stu-id="423aa-111">Sample code: Clear table column filter based on active cell</span></span>

<span data-ttu-id="423aa-112">Das folgende Skript entfernt den Tabellenspaltenfilter basierend auf der aktiven Zellenposition und kann auf jede Excel tabelle angewendet werden.</span><span class="sxs-lookup"><span data-stu-id="423aa-112">The following script clears the table column filter based on active cell location and can be applied to any Excel file with a table.</span></span> <span data-ttu-id="423aa-113">Zur Vereinfachung können Sie die Datei herunterladen und <a href="table-with-filter.xlsx">table-with-filter.xlsx. </a></span><span class="sxs-lookup"><span data-stu-id="423aa-113">For convenience, you can download and use <a href="table-with-filter.xlsx">table-with-filter.xlsx</a>.</span></span>

```TypeScript
function main(workbook: ExcelScript.Workbook) {
    // Get the active cell.
    const cell = workbook.getActiveCell();

    // Get all tables associated with that cell.
    const tables = cell.getTables();
    
    // If there is no table on the selection, end the script.
    if (tables.length !== 1) {
      console.log("The selection is not in a table.");
      return;
    }

    // Get the first table associated with the active cell.
    const currentTable = tables[0];

    // Log key information about the table.
    console.log(currentTable.getName());
    console.log(currentTable.getRange().getAddress());

    // Get the table header above the current cell by referencing its column.
    const entireColumn = cell.getEntireColumn();
    const intersect = entireColumn.getIntersection(currentTable.getRange());
    console.log(intersect.getAddress());

    const headerCellValue = intersect.getCell(0,0).getValue() as string;
    console.log(headerCellValue);

    // Get the TableColumn object matching that header.
    const tableColumn = currentTable.getColumnByName(headerCellValue);

    // Clear the filter on that table column.
    tableColumn.getFilter().clear();
}
```
