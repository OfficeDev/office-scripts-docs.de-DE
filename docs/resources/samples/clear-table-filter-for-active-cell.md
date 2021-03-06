---
title: Löschen des Tabellenspaltenfilters basierend auf der Position der aktiven Zelle
description: Erfahren Sie, wie Sie den Tabellenspaltenfilter basierend auf der aktiven Zellenposition löschen.
ms.date: 06/29/2021
localization_priority: Normal
ms.openlocfilehash: f10e23b4ad948a28c5b749533ddedefe164d7142
ms.sourcegitcommit: 211c157ca746e266eeb079f5fa1925a1e35ab702
ms.translationtype: MT
ms.contentlocale: de-DE
ms.lasthandoff: 07/07/2021
ms.locfileid: "53313890"
---
# <a name="clear-table-column-filter-based-on-active-cell-location"></a>Löschen des Tabellenspaltenfilters basierend auf der Position der aktiven Zelle

In diesem Beispiel wird der Tabellenspaltenfilter basierend auf der position der aktiven Zelle gelöscht. Das Skript erkennt, ob die Zelle Teil einer Tabelle ist, bestimmt die Tabellenspalte und löscht alle Filter, die darauf angewendet werden.

Wenn Sie mehr darüber erfahren möchten, wie Sie den Filter speichern, bevor Sie ihn löschen (und später erneut anwenden), finden Sie unter [Verschieben von Zeilen in Tabellen durch Speichern von Filtern](move-rows-across-tables.md)ein erweitertes Beispiel.

_Vor dem Löschen des Spaltenfilters (beachten Sie die aktive Zelle)_

:::image type="content" source="../../images/before-filter-applied.png" alt-text="Eine aktive Zelle vor dem Löschen des Spaltenfilters.":::

_Nach dem Löschen des Spaltenfilters_

:::image type="content" source="../../images/after-filter-cleared.png" alt-text="Eine aktive Zelle nach dem Löschen des Spaltenfilters.":::

## <a name="sample-excel-file"></a>Beispieldatei für Excel

Laden Sie <a href="table-with-filter.xlsx">table-with-filter.xlsx</a> für eine einsatzbereite Arbeitsmappe herunter. Fügen Sie das folgende Skript hinzu, um das Beispiel selbst auszuprobieren!

## <a name="sample-code-clear-table-column-filter-based-on-active-cell"></a>Beispielcode: Löschen des Tabellenspaltenfilters basierend auf der aktiven Zelle

Das folgende Skript löscht den Tabellenspaltenfilter basierend auf der Position der aktiven Zelle und kann auf eine beliebige Excel Datei mit einer Tabelle angewendet werden.

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
