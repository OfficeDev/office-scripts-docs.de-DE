---
title: Zählen leerer Zeilen in Blättern
description: In diesem Artikel erfahren Sie, wie Sie mithilfe von Office Skripts ermitteln, ob es leere Zeilen anstelle von Daten in Arbeitsblättern gibt, und dann die Anzahl leerer Zeilen melden, die in einem Power Automate-Fluss verwendet werden soll.
ms.date: 06/29/2021
localization_priority: Normal
ms.openlocfilehash: e5b60779d2ca2de5f4cf4e03ddd6ff7372515ad6
ms.sourcegitcommit: 211c157ca746e266eeb079f5fa1925a1e35ab702
ms.translationtype: MT
ms.contentlocale: de-DE
ms.lasthandoff: 07/07/2021
ms.locfileid: "53313806"
---
# <a name="count-blank-rows-on-sheets"></a><span data-ttu-id="8e562-103">Zählen leerer Zeilen in Blättern</span><span class="sxs-lookup"><span data-stu-id="8e562-103">Count blank rows on sheets</span></span>

<span data-ttu-id="8e562-104">Dieses Projekt enthält zwei Skripts:</span><span class="sxs-lookup"><span data-stu-id="8e562-104">This project includes two scripts:</span></span>

* <span data-ttu-id="8e562-105">[Zählen leerer Zeilen auf einem bestimmten Blatt:](#sample-code-count-blank-rows-on-a-given-sheet)Durchläuft den verwendeten Bereich eines bestimmten Arbeitsblatts und gibt eine leere Zeilenanzahl zurück.</span><span class="sxs-lookup"><span data-stu-id="8e562-105">[Count blank rows on a given sheet](#sample-code-count-blank-rows-on-a-given-sheet): Traverses the used range on a given worksheet and returns a blank row count.</span></span>
* <span data-ttu-id="8e562-106">[Leere Zeilen auf allen Blättern](#sample-code-count-blank-rows-on-all-sheets)zählen: Durchläuft den verwendeten Bereich _auf allen Arbeitsblättern_ und gibt eine leere Zeilenanzahl zurück.</span><span class="sxs-lookup"><span data-stu-id="8e562-106">[Count blank rows on all sheets](#sample-code-count-blank-rows-on-all-sheets): Traverses the used range on _all of the worksheets_ and returns a blank row count.</span></span>

> [!NOTE]
> <span data-ttu-id="8e562-107">Für unser Skript ist eine leere Zeile eine Zeile, in der keine Daten vorhanden sind.</span><span class="sxs-lookup"><span data-stu-id="8e562-107">For our script, a blank row is any row where there's no data.</span></span> <span data-ttu-id="8e562-108">Die Zeile kann eine Formatierung aufweisen.</span><span class="sxs-lookup"><span data-stu-id="8e562-108">The row can have formatting.</span></span>

<span data-ttu-id="8e562-109">_Dieses Blatt gibt die Anzahl von 4 leeren Zeilen zurück._</span><span class="sxs-lookup"><span data-stu-id="8e562-109">_This sheet returns count of 4 blank rows_</span></span>

:::image type="content" source="../../images/blank-rows.png" alt-text="Ein Arbeitsblatt mit Daten mit leeren Zeilen.":::

<span data-ttu-id="8e562-111">_Dieses Blatt gibt die Anzahl von 0 leeren Zeilen zurück (alle Zeilen haben einige Daten)_</span><span class="sxs-lookup"><span data-stu-id="8e562-111">_This sheet returns count of 0 blank rows (all rows have some data)_</span></span>

:::image type="content" source="../../images/no-blank-rows.png" alt-text="Ein Arbeitsblatt mit Daten ohne leere Zeilen.":::

## <a name="sample-code-count-blank-rows-on-a-given-sheet"></a><span data-ttu-id="8e562-113">Beispielcode: Zählen leerer Zeilen auf einem bestimmten Blatt</span><span class="sxs-lookup"><span data-stu-id="8e562-113">Sample code: Count blank rows on a given sheet</span></span>

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

## <a name="sample-code-count-blank-rows-on-all-sheets"></a><span data-ttu-id="8e562-114">Beispielcode: Zählen leerer Zeilen auf allen Blättern</span><span class="sxs-lookup"><span data-stu-id="8e562-114">Sample code: Count blank rows on all sheets</span></span>

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
