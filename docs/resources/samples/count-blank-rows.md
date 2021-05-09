---
title: Zählen leerer Zeilen in Blättern
description: Erfahren Sie, wie Sie Office-Skripts verwenden, um zu erkennen, ob leere Zeilen anstelle von Daten in Arbeitsblättern vorhanden sind, und melden Sie dann die Anzahl leerer Zeilen, die in einem Power Automate werden sollen.
ms.date: 05/04/2021
localization_priority: Normal
ms.openlocfilehash: e636c9b1b24dedb73042cd9ee4d20688698ae8a7
ms.sourcegitcommit: 763d341857bcb209b2f2c278a82fdb63d0e18f0a
ms.translationtype: MT
ms.contentlocale: de-DE
ms.lasthandoff: 05/08/2021
ms.locfileid: "52285850"
---
# <a name="count-blank-rows-on-sheets"></a><span data-ttu-id="b1ed4-103">Zählen leerer Zeilen in Blättern</span><span class="sxs-lookup"><span data-stu-id="b1ed4-103">Count blank rows on sheets</span></span>

<span data-ttu-id="b1ed4-104">Dieses Projekt umfasst zwei Skripts:</span><span class="sxs-lookup"><span data-stu-id="b1ed4-104">This project includes two scripts:</span></span>

* <span data-ttu-id="b1ed4-105">[Zählen Sie leere Zeilen auf einem bestimmten](#sample-code-count-blank-rows-on-a-given-sheet)Blatt: Durchläuft den verwendeten Bereich in einem bestimmten Arbeitsblatt und gibt eine leere Zeilenanzahl zurück.</span><span class="sxs-lookup"><span data-stu-id="b1ed4-105">[Count blank rows on a given sheet](#sample-code-count-blank-rows-on-a-given-sheet): Traverses the used range on a given worksheet and returns a blank row count.</span></span>
* <span data-ttu-id="b1ed4-106">[Zählen Sie leere Zeilen in allen Blättern:](#sample-code-count-blank-rows-on-all-sheets)Durchläuft den verwendeten Bereich _für_ alle Arbeitsblätter und gibt eine leere Zeilenanzahl zurück.</span><span class="sxs-lookup"><span data-stu-id="b1ed4-106">[Count blank rows on all sheets](#sample-code-count-blank-rows-on-all-sheets): Traverses the used range on _all of the worksheets_ and returns a blank row count.</span></span>

> [!NOTE]
> <span data-ttu-id="b1ed4-107">Für unser Skript ist eine leere Zeile eine beliebige Zeile, in der keine Daten enthalten sind.</span><span class="sxs-lookup"><span data-stu-id="b1ed4-107">For our script, a blank row is any row where there's no data.</span></span> <span data-ttu-id="b1ed4-108">Die Zeile kann Formatierung haben.</span><span class="sxs-lookup"><span data-stu-id="b1ed4-108">The row can have formatting.</span></span>

<span data-ttu-id="b1ed4-109">_Dieses Blatt gibt die Anzahl von 4 leeren Zeilen zurück._</span><span class="sxs-lookup"><span data-stu-id="b1ed4-109">_This sheet returns count of 4 blank rows_</span></span>

:::image type="content" source="../../images/blank-rows.png" alt-text="Ein Arbeitsblatt mit Daten mit leeren Zeilen":::

<span data-ttu-id="b1ed4-111">_Dieses Blatt gibt die Anzahl von 0 leeren Zeilen zurück (alle Zeilen haben einige Daten)_</span><span class="sxs-lookup"><span data-stu-id="b1ed4-111">_This sheet returns count of 0 blank rows (all rows have some data)_</span></span>

:::image type="content" source="../../images/no-blank-rows.png" alt-text="Ein Arbeitsblatt mit Daten ohne leere Zeilen":::

## <a name="sample-code-count-blank-rows-on-a-given-sheet"></a><span data-ttu-id="b1ed4-113">Beispielcode: Zählen leerer Zeilen auf einem bestimmten Blatt</span><span class="sxs-lookup"><span data-stu-id="b1ed4-113">Sample code: Count blank rows on a given sheet</span></span>

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

## <a name="sample-code-count-blank-rows-on-all-sheets"></a><span data-ttu-id="b1ed4-114">Beispielcode: Zählen leerer Zeilen in allen Blättern</span><span class="sxs-lookup"><span data-stu-id="b1ed4-114">Sample code: Count blank rows on all sheets</span></span>

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

## <a name="use-with-power-automate"></a><span data-ttu-id="b1ed4-115">Verwenden mit Power Automate</span><span class="sxs-lookup"><span data-stu-id="b1ed4-115">Use with Power Automate</span></span>

:::image type="content" source="../../images/use-in-power-automate.png" alt-text="Ein Power Automate, der zeigt, wie Sie ein Skript Office einrichten":::
