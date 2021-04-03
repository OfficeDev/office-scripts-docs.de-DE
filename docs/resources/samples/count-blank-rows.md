---
title: Zählen leerer Zeilen in Blättern
description: Erfahren Sie, wie Sie mithilfe von Office-Skripts ermitteln, ob leere Zeilen anstelle von Daten in Arbeitsblättern vorhanden sind, und dann die Anzahl leerer Zeilen melden, die in einem Power Automate-Fluss verwendet werden soll.
ms.date: 03/31/2021
localization_priority: Normal
ms.openlocfilehash: 1f52b9c4d538d5d3e64dc61dae3e27d046b56862
ms.sourcegitcommit: 5d24e77df70aa2c1c982275d53213c2a9323ff86
ms.translationtype: MT
ms.contentlocale: de-DE
ms.lasthandoff: 04/02/2021
ms.locfileid: "51571400"
---
# <a name="count-blank-rows-on-sheets"></a>Zählen leerer Zeilen in Blättern

Dieses Projekt umfasst zwei Skripts:

* [Zählen Sie leere Zeilen auf einem bestimmten](#sample-code-count-blank-rows-on-a-given-sheet)Blatt: Durchläuft den verwendeten Bereich in einem bestimmten Arbeitsblatt und gibt eine leere Zeilenanzahl zurück.
* [Zählen Sie leere Zeilen in allen Blättern:](#sample-code-count-blank-rows-on-all-sheets)Durchläuft den verwendeten Bereich _für_ alle Arbeitsblätter und gibt eine leere Zeilenanzahl zurück.

> [!NOTE]
> Für unser Skript ist eine leere Zeile eine beliebige Zeile, in der keine Daten enthalten sind. Die Zeile kann Formatierung haben.

_Dieses Blatt gibt die Anzahl von 4 leeren Zeilen zurück._

![Daten mit leeren Zeilen](../../images/blank-rows.png)

_Dieses Blatt gibt die Anzahl von 0 leeren Zeilen zurück (alle Zeilen haben einige Daten)_

![Daten ohne leere Zeilen](../../images/no-blank-rows.png)

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

![Screenshot der Einrichtung in Power Automate](../../images/use-in-power-automate.png)
