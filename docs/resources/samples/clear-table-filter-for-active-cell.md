---
title: Löschen des Tabellenspaltenfilters basierend auf der aktiven Zellenposition
description: Erfahren Sie, wie Sie den Tabellenspaltenfilter basierend auf der aktiven Zellenposition löschen.
ms.date: 03/04/2021
localization_priority: Normal
ms.openlocfilehash: bbca4adce1de2cfade2c4f84273bf0bc06b5cc4b
ms.sourcegitcommit: f7a7aebfb687f2a35dbed07ed62ff352a114525a
ms.translationtype: MT
ms.contentlocale: de-DE
ms.lasthandoff: 05/06/2021
ms.locfileid: "52232501"
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
    // Get active cell.
    const cell = workbook.getActiveCell();

    // Get all tables associated with that cell.
    const tables = cell.getTables();
    
    // If there is no table on the selection, return/exit.
    if (tables.length !== 1) {
      console.log("The selection is not in a table.");
      return;
    }

    // Get table (since it is already determined that there is only
    // a single table part of the selection).
    const currentTable = tables[0];

    console.log(currentTable.getName());
    console.log(currentTable.getRange().getAddress());

    const entireCol = cell.getEntireColumn();
    const intersect = entireCol.getIntersection(currentTable.getRange());
    console.log(intersect.getAddress());

    const headerCellValue = intersect.getCell(0,0).getValue() as string;
    console.log(headerCellValue);

    // Get column.
    const col = currentTable.getColumnByName(headerCellValue);

    // Clear filter.
    col.getFilter().clear();
}
```
