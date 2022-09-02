---
title: Tabellenspaltenfilter entfernen
description: Erfahren Sie, wie Sie den Tabellenspaltenfilter basierend auf der aktiven Zellenposition löschen.
ms.date: 07/15/2022
ms.localizationpriority: medium
ms.openlocfilehash: e016f7f2af9e7553229f3b3b19007e011879de8e
ms.sourcegitcommit: a6504f8b0d6b717457c6e0b5306c35ad3900914e
ms.translationtype: MT
ms.contentlocale: de-DE
ms.lasthandoff: 09/02/2022
ms.locfileid: "67572521"
---
# <a name="remove-table-column-filters"></a>Tabellenspaltenfilter entfernen

In diesem Beispiel werden die Filter basierend auf der aktiven Zellenposition aus einer Tabellenspalte entfernt. Das Skript erkennt, ob die Zelle Teil einer Tabelle ist, bestimmt die Tabellenspalte und löscht alle Filter, die darauf angewendet werden.

Wenn Sie mehr darüber erfahren möchten, wie Sie den Filter speichern, bevor Sie ihn löschen (und später erneut anwenden), lesen [Sie "Verschieben von Zeilen über Tabellen hinweg", indem Sie Filter speichern](move-rows-across-tables.md), ein erweitertes Beispiel.

## <a name="sample-excel-file"></a>Excel-Beispieldatei

Laden Sie [table-with-filter.xlsx](table-with-filter.xlsx) für eine sofort einsatzbereite Arbeitsmappe herunter. Fügen Sie das folgende Skript hinzu, um das Beispiel selbst auszuprobieren!

## <a name="sample-code-clear-table-column-filter-based-on-active-cell"></a>Beispielcode: Löschen des Tabellenspaltenfilters basierend auf der aktiven Zelle

Das folgende Skript löscht den Tabellenspaltenfilter basierend auf der aktiven Zellenposition und kann auf jede Excel-Datei mit einer Tabelle angewendet werden.

```TypeScript
function main(workbook: ExcelScript.Workbook) {
  // Get the active cell.
  const cell = workbook.getActiveCell();

  // Get the tables associated with that cell.
  // Since tables can't overlap, this will be one table at most.
  const currentTable = cell.getTables()[0];

  // If there is no table on the selection, end the script.
  if (!currentTable) {
    console.log("The selection is not in a table.");
    return;
  }

  // Get the table header above the current cell by referencing its column.
  const entireColumn = cell.getEntireColumn();
  const intersect = entireColumn.getIntersection(currentTable.getRange());
  const headerCellValue = intersect.getCell(0, 0).getValue() as string;

  // Get the TableColumn object matching that header.
  const tableColumn = currentTable.getColumnByName(headerCellValue);

  // Clear the filters on that table column.
  tableColumn.getFilter().clear();
}
```

## <a name="before-clearing-column-filter-notice-the-active-cell"></a>Vor dem Löschen des Spaltenfilters (beachten Sie die aktive Zelle)

:::image type="content" source="../../images/before-filter-applied.png" alt-text="Eine aktive Zelle vor dem Löschen des Spaltenfilters.":::

## <a name="after-clearing-column-filter"></a>Nach dem Löschen des Spaltenfilters

:::image type="content" source="../../images/after-filter-cleared.png" alt-text="Eine aktive Zelle nach dem Löschen des Spaltenfilters.":::
