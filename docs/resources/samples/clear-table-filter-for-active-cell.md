---
title: Löschen des Tabellenspaltenfilters basierend auf der Position der aktiven Zelle
description: Erfahren Sie, wie Sie den Tabellenspaltenfilter basierend auf der aktiven Zellenposition löschen.
ms.date: 05/03/2021
localization_priority: Normal
ms.openlocfilehash: d6f267b433be9a0ddf44edf53ed92a136eb2ded6
ms.sourcegitcommit: 4693c8f79428ec74695328275703af0ba1bfea8f
ms.translationtype: MT
ms.contentlocale: de-DE
ms.lasthandoff: 06/23/2021
ms.locfileid: "53074438"
---
# <a name="clear-table-column-filter-based-on-active-cell-location"></a><span data-ttu-id="fce1b-103">Löschen des Tabellenspaltenfilters basierend auf der Position der aktiven Zelle</span><span class="sxs-lookup"><span data-stu-id="fce1b-103">Clear table column filter based on active cell location</span></span>

<span data-ttu-id="fce1b-104">In diesem Beispiel wird der Tabellenspaltenfilter basierend auf der position der aktiven Zelle gelöscht.</span><span class="sxs-lookup"><span data-stu-id="fce1b-104">This sample clears the table column filter based on the active cell location.</span></span> <span data-ttu-id="fce1b-105">Das Skript erkennt, ob die Zelle Teil einer Tabelle ist, bestimmt die Tabellenspalte und löscht alle Filter, die darauf angewendet werden.</span><span class="sxs-lookup"><span data-stu-id="fce1b-105">The script detects if the cell is part of a table, determines the table column, and clears any filter that are applied on it.</span></span>

<span data-ttu-id="fce1b-106">Wenn Sie mehr darüber erfahren möchten, wie Sie den Filter speichern, bevor Sie ihn löschen (und später erneut anwenden), finden Sie unter [Verschieben von Zeilen in Tabellen durch Speichern von Filtern](move-rows-across-tables.md)ein erweitertes Beispiel.</span><span class="sxs-lookup"><span data-stu-id="fce1b-106">If you wish to learn more about how to save the filter prior to clearing it (and re-apply later), see [Move rows across tables by saving filters](move-rows-across-tables.md), a more advanced sample.</span></span>

<span data-ttu-id="fce1b-107">_Vor dem Löschen des Spaltenfilters (beachten Sie die aktive Zelle)_</span><span class="sxs-lookup"><span data-stu-id="fce1b-107">_Before clearing column filter (notice the active cell)_</span></span>

:::image type="content" source="../../images/before-filter-applied.png" alt-text="Eine aktive Zelle vor dem Löschen des Spaltenfilters.":::

<span data-ttu-id="fce1b-109">_Nach dem Löschen des Spaltenfilters_</span><span class="sxs-lookup"><span data-stu-id="fce1b-109">_After clearing column filter_</span></span>

:::image type="content" source="../../images/after-filter-cleared.png" alt-text="Eine aktive Zelle nach dem Löschen des Spaltenfilters.":::

## <a name="sample-code-clear-table-column-filter-based-on-active-cell"></a><span data-ttu-id="fce1b-111">Beispielcode: Löschen des Tabellenspaltenfilters basierend auf der aktiven Zelle</span><span class="sxs-lookup"><span data-stu-id="fce1b-111">Sample code: Clear table column filter based on active cell</span></span>

<span data-ttu-id="fce1b-112">Das folgende Skript löscht den Tabellenspaltenfilter basierend auf der Position der aktiven Zelle und kann auf jede Excel Datei mit einer Tabelle angewendet werden.</span><span class="sxs-lookup"><span data-stu-id="fce1b-112">The following script clears the table column filter based on active cell location and can be applied to any Excel file with a table.</span></span> <span data-ttu-id="fce1b-113">Aus Gründen der Einfachheit können Sie <a href="table-with-filter.xlsx">table-with-filter.xlsx</a>herunterladen und verwenden.</span><span class="sxs-lookup"><span data-stu-id="fce1b-113">For convenience, you can download and use <a href="table-with-filter.xlsx">table-with-filter.xlsx</a>.</span></span>

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
