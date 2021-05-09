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
# <a name="clear-table-column-filter-based-on-active-cell-location"></a>Löschen des Tabellenspaltenfilters basierend auf der aktiven Zellenposition

In diesem Beispiel wird der Tabellenspaltenfilter basierend auf der aktiven Zellenposition entfernt. Das Skript erkennt, ob die Zelle Teil einer Tabelle ist, bestimmt die Tabellenspalte und filtert, die darauf angewendet werden.

Wenn Sie mehr darüber erfahren möchten, wie Sie den Filter speichern, bevor Sie ihn löschen (und später erneut anwenden), finden Sie weitere Informationen unter [Move rows across tables by saving filters](move-rows-across-tables.md), a more advanced sample.

_Vor dem Löschen des Spaltenfilters (beachten Sie die aktive Zelle)_

:::image type="content" source="../../images/before-filter-applied.png" alt-text="Eine aktive Zelle vor dem Löschen des Spaltenfilters":::

_Nach dem Löschen des Spaltenfilters_

:::image type="content" source="../../images/after-filter-cleared.png" alt-text="Eine aktive Zelle nach dem Löschen des Spaltenfilters":::

## <a name="sample-code-clear-table-column-filter-based-on-active-cell"></a>Beispielcode: Tabellenspaltenfilter basierend auf aktiver Zelle löschen

Das folgende Skript entfernt den Tabellenspaltenfilter basierend auf der aktiven Zellenposition und kann auf jede Excel tabelle angewendet werden. Zur Vereinfachung können Sie die Datei herunterladen und <a href="table-with-filter.xlsx">table-with-filter.xlsx. </a>

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
