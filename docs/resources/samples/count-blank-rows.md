---
title: Zählen leerer Zeilen in Blättern
description: In diesem Artikel erfahren Sie, wie Sie mithilfe von Office Skripts ermitteln, ob es leere Zeilen anstelle von Daten in Arbeitsblättern gibt, und dann die Anzahl leerer Zeilen melden, die in einem Power Automate-Fluss verwendet werden soll.
ms.date: 06/29/2021
ms.localizationpriority: medium
ms.openlocfilehash: 103d2f96c1780b47363dcb6caab82553dd556b80
ms.sourcegitcommit: d3ed4bdeeba805d97c930394e172e8306a0cf484
ms.translationtype: MT
ms.contentlocale: de-DE
ms.lasthandoff: 09/15/2021
ms.locfileid: "59332213"
---
# <a name="count-blank-rows-on-sheets"></a>Zählen leerer Zeilen in Blättern

Dieses Projekt enthält zwei Skripts:

* [Zählen leerer Zeilen auf einem bestimmten Blatt:](#sample-code-count-blank-rows-on-a-given-sheet)Durchläuft den verwendeten Bereich eines bestimmten Arbeitsblatts und gibt eine leere Zeilenanzahl zurück.
* [Leere Zeilen auf allen Blättern](#sample-code-count-blank-rows-on-all-sheets)zählen: Durchläuft den verwendeten Bereich _auf allen Arbeitsblättern_ und gibt eine leere Zeilenanzahl zurück.

> [!NOTE]
> Für unser Skript ist eine leere Zeile eine Zeile, in der keine Daten vorhanden sind. Die Zeile kann eine Formatierung aufweisen.

_Dieses Blatt gibt die Anzahl von 4 leeren Zeilen zurück._

:::image type="content" source="../../images/blank-rows.png" alt-text="Ein Arbeitsblatt mit Daten mit leeren Zeilen.":::

_Dieses Blatt gibt die Anzahl von 0 leeren Zeilen zurück (alle Zeilen haben einige Daten)_

:::image type="content" source="../../images/no-blank-rows.png" alt-text="Ein Arbeitsblatt mit Daten ohne leere Zeilen.":::

## <a name="sample-code-count-blank-rows-on-a-given-sheet"></a>Beispielcode: Zählen leerer Zeilen auf einem bestimmten Blatt

```TypeScript
function main(workbook: ExcelScript.Workbook): number
{
  // Get the worksheet named "Sheet1".
  const sheet = workbook.getWorksheet('Sheet1'); 
  
  // Get the entire data range.
  const range = sheet.getUsedRange(true);

  // If the used range is empty, end the script.
  if (!range) {
    console.log(`No data on this sheet.`);
    return;
  }
  
  // Log the address of the used range.
  console.log(`Used range for the worksheet: ${range.getAddress()}`);
    
  // Look through the values in the range for blank rows.
  const values = range.getValues();
  let emptyRows = 0;
  for (let row of values) {
    let emptyRow = true;
    
    // Look at every cell in the row for one with a value.
    for (let cell of row) {
      if (cell.toString().length > 0) {
        emptyRow = false
      }
    }

    // If no cell had a value, the row is empty.
    if (emptyRow) {
      emptyRows++;
    }
  }

  // Log the number of empty rows.
  console.log(`Total empty rows: ${emptyRows}`);

  // Return the number of empty rows for use in a Power Automate flow.
  return emptyRows;
}
```

## <a name="sample-code-count-blank-rows-on-all-sheets"></a>Beispielcode: Zählen leerer Zeilen auf allen Blättern

```TypeScript
function main(workbook: ExcelScript.Workbook): number
{
  // Loop through every worksheet in the workbook.
  const sheets = workbook.getWorksheets();
  let emptyRows = 0;
  for (let sheet of sheets) {     
    // Get the entire data range.
    const range = sheet.getUsedRange(true);
  
    // If the used range is empty, skip to the next worksheet.
    if (!range) {
      console.log(`No data on this sheet.`);
      continue;
    }
    
    // Log the address of the used range.
    console.log(`Used range for the worksheet: ${range.getAddress()}`);
      
    // Look through the values in the range for blank rows.
    const values = range.getValues();
    for (let row of values) {
      let emptyRow = true;
      
      // Look at every cell in the row for one with a value.
      for (let cell of row) {
        if (cell.toString().length > 0) {
          emptyRow = false
        }
      }
  
      // If no cell had a value, the row is empty.
      if (emptyRow) {
        emptyRows++;
      }
    }
  }

  // Log the number of empty rows.
  console.log(`Total empty rows: ${emptyRows}`);

  // Return the number of empty rows for use in a Power Automate flow.
  return emptyRows;
}
```
