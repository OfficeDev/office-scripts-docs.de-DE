---
title: Zählen leerer Zeilen in Blättern
description: Erfahren Sie, wie Sie Office-Skripts verwenden, um zu erkennen, ob leere Zeilen anstelle von Daten in Arbeitsblättern vorhanden sind, und melden Sie dann die Anzahl leerer Zeilen, die in einem Power Automate werden sollen.
ms.date: 03/31/2021
localization_priority: Normal
ms.openlocfilehash: db84f2446c168f867c325a05129fe982c9645731
ms.sourcegitcommit: f7a7aebfb687f2a35dbed07ed62ff352a114525a
ms.translationtype: MT
ms.contentlocale: de-DE
ms.lasthandoff: 05/06/2021
ms.locfileid: "52232585"
---
# <a name="count-blank-rows-on-sheets"></a>Zählen leerer Zeilen in Blättern

Dieses Projekt umfasst zwei Skripts:

* [Zählen Sie leere Zeilen auf einem bestimmten](#sample-code-count-blank-rows-on-a-given-sheet)Blatt: Durchläuft den verwendeten Bereich in einem bestimmten Arbeitsblatt und gibt eine leere Zeilenanzahl zurück.
* [Zählen Sie leere Zeilen in allen Blättern:](#sample-code-count-blank-rows-on-all-sheets)Durchläuft den verwendeten Bereich _für_ alle Arbeitsblätter und gibt eine leere Zeilenanzahl zurück.

> [!NOTE]
> Für unser Skript ist eine leere Zeile eine beliebige Zeile, in der keine Daten enthalten sind. Die Zeile kann Formatierung haben.

_Dieses Blatt gibt die Anzahl von 4 leeren Zeilen zurück._

:::image type="content" source="../../images/blank-rows.png" alt-text="Ein Arbeitsblatt mit Daten mit leeren Zeilen":::

_Dieses Blatt gibt die Anzahl von 0 leeren Zeilen zurück (alle Zeilen haben einige Daten)_

:::image type="content" source="../../images/no-blank-rows.png" alt-text="Ein Arbeitsblatt mit Daten ohne leere Zeilen":::

## <a name="sample-code-count-blank-rows-on-a-given-sheet"></a>Beispielcode: Zählen leerer Zeilen auf einem bestimmten Blatt

```TypeScript
function main(workbook: ExcelScript.Workbook): number
{
  const sheet = workbook.getWorksheet('Sheet1'); 
  // Getting the active worksheet is not suitable for a script used by Power Automate.
  // const sheet = workbook.getActiveWorksheet();
  
  const range = sheet.getUsedRange(true); // Get value only.
  if (!range) {
    console.log(`No data on this sheet. `);
    return;
  }
  console.log(`Used range for the worksheet: ${range.getAddress()}`);
  const values = range.getValues();
  let emptyRows = 0;
  for (let row of values) {
    let len = 0; 
    for (let cell of row) {
      len = len + cell.toString().length;
    }
    if (len === 0) { 
      emptyRows++;
    }
  }
  console.log(`Total empty row: ` + emptyRows);
  return emptyRows;
}
```

## <a name="sample-code-count-blank-rows-on-all-sheets"></a>Beispielcode: Zählen leerer Zeilen in allen Blättern

```TypeScript
function main(workbook: ExcelScript.Workbook): number
{
  const sheets = workbook.getWorksheets();
  let emptyRows = 0;
  for (let sheet of sheets) { 
    const range = sheet.getUsedRange(true); // Get value only.
    if (!range) {
      console.log(`No data on this sheet. `);
      continue;
    }
    console.log(`Used range for the worksheet ${sheet.getName()}: ${range.getAddress()}`);
    const values = range.getValues();

    for (let row of values) {
      let len = 0;
      for (let cell of row) {
        len = len + cell.toString().length;
      }
      if (len === 0) {
        emptyRows++;
      }
    }
  }
  console.log(`Total empty row: ` + emptyRows);
  return emptyRows;
}
```

## <a name="use-with-power-automate"></a>Verwenden mit Power Automate

:::image type="content" source="../../images/use-in-power-automate.png" alt-text="Ein Power Automate, der zeigt, wie Sie ein Skript Office einrichten":::
