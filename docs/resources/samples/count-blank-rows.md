---
title: Zählen leerer Zeilen in Blättern
description: Erfahren Sie, wie Sie mithilfe von Office-Skripts ermitteln, ob leere Zeilen anstelle von Daten in Arbeitsblättern vorhanden sind, und dann die Anzahl leerer Zeilen melden, die in einem Power Automate-Fluss verwendet werden soll.
ms.date: 03/31/2021
localization_priority: Normal
ms.openlocfilehash: 088ab97c686484ca5c13c875b80431ac28d20736
ms.sourcegitcommit: 45ffe3dbd2c834b78592ad35928cf8096f5e80bc
ms.translationtype: MT
ms.contentlocale: de-DE
ms.lasthandoff: 04/14/2021
ms.locfileid: "51754831"
---
# <a name="count-blank-rows-on-sheets"></a><span data-ttu-id="431c6-103">Zählen leerer Zeilen in Blättern</span><span class="sxs-lookup"><span data-stu-id="431c6-103">Count blank rows on sheets</span></span>

<span data-ttu-id="431c6-104">Dieses Projekt umfasst zwei Skripts:</span><span class="sxs-lookup"><span data-stu-id="431c6-104">This project includes two scripts:</span></span>

* <span data-ttu-id="431c6-105">[Zählen Sie leere Zeilen auf einem bestimmten](#sample-code-count-blank-rows-on-a-given-sheet)Blatt: Durchläuft den verwendeten Bereich in einem bestimmten Arbeitsblatt und gibt eine leere Zeilenanzahl zurück.</span><span class="sxs-lookup"><span data-stu-id="431c6-105">[Count blank rows on a given sheet](#sample-code-count-blank-rows-on-a-given-sheet): Traverses the used range on a given worksheet and returns a blank row count.</span></span>
* <span data-ttu-id="431c6-106">[Zählen Sie leere Zeilen in allen Blättern:](#sample-code-count-blank-rows-on-all-sheets)Durchläuft den verwendeten Bereich _für_ alle Arbeitsblätter und gibt eine leere Zeilenanzahl zurück.</span><span class="sxs-lookup"><span data-stu-id="431c6-106">[Count blank rows on all sheets](#sample-code-count-blank-rows-on-all-sheets): Traverses the used range on _all of the worksheets_ and returns a blank row count.</span></span>

> [!NOTE]
> <span data-ttu-id="431c6-107">Für unser Skript ist eine leere Zeile eine beliebige Zeile, in der keine Daten enthalten sind.</span><span class="sxs-lookup"><span data-stu-id="431c6-107">For our script, a blank row is any row where there's no data.</span></span> <span data-ttu-id="431c6-108">Die Zeile kann Formatierung haben.</span><span class="sxs-lookup"><span data-stu-id="431c6-108">The row can have formatting.</span></span>

<span data-ttu-id="431c6-109">_Dieses Blatt gibt die Anzahl von 4 leeren Zeilen zurück._</span><span class="sxs-lookup"><span data-stu-id="431c6-109">_This sheet returns count of 4 blank rows_</span></span>

:::image type="content" source="../../images/blank-rows.png" alt-text="Ein Arbeitsblatt mit Daten mit leeren Zeilen.":::

<span data-ttu-id="431c6-111">_Dieses Blatt gibt die Anzahl von 0 leeren Zeilen zurück (alle Zeilen haben einige Daten)_</span><span class="sxs-lookup"><span data-stu-id="431c6-111">_This sheet returns count of 0 blank rows (all rows have some data)_</span></span>

:::image type="content" source="../../images/no-blank-rows.png" alt-text="Ein Arbeitsblatt mit Daten ohne leere Zeilen.":::

## <a name="sample-code-count-blank-rows-on-a-given-sheet"></a><span data-ttu-id="431c6-113">Beispielcode: Zählen leerer Zeilen auf einem bestimmten Blatt</span><span class="sxs-lookup"><span data-stu-id="431c6-113">Sample code: Count blank rows on a given sheet</span></span>

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

## <a name="sample-code-count-blank-rows-on-all-sheets"></a><span data-ttu-id="431c6-114">Beispielcode: Zählen leerer Zeilen in allen Blättern</span><span class="sxs-lookup"><span data-stu-id="431c6-114">Sample code: Count blank rows on all sheets</span></span>

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

## <a name="use-with-power-automate"></a><span data-ttu-id="431c6-115">Verwenden mit Power Automate</span><span class="sxs-lookup"><span data-stu-id="431c6-115">Use with Power Automate</span></span>

:::image type="content" source="../../images/use-in-power-automate.png" alt-text="Ein Power Automate-Fluss, der zeigt, wie Sie ein Office-Skript ausführen.":::
