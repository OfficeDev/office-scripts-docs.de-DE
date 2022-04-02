---
title: Inhaltsverzeichnis für eine Arbeitsmappe erstellen
description: Erfahren Sie, wie Sie ein Inhaltsverzeichnis mit Links zu den einzelnen Arbeitsblättern erstellen.
ms.date: 01/20/2022
ms.localizationpriority: medium
ms.openlocfilehash: b2d69609514c2e1e87f9c0590ea10152fc7d5e7d
ms.sourcegitcommit: 7023b9e23499806901a5ecf8ebc460b76887cca6
ms.translationtype: MT
ms.contentlocale: de-DE
ms.lasthandoff: 03/31/2022
ms.locfileid: "64585520"
---
# <a name="create-a-workbook-table-of-contents"></a>Inhaltsverzeichnis für eine Arbeitsmappe erstellen

In diesem Beispiel wird gezeigt, wie Sie ein Inhaltsverzeichnis für die Arbeitsmappe erstellen. Jeder Eintrag im Inhaltsverzeichnis ist ein Hyperlink zu einem der Arbeitsblätter in der Arbeitsmappe.

:::image type="content" source="../../images/table-of-contents-sample.png" alt-text="Das Inhaltsverzeichnis-Arbeitsblatt mit Links zu den anderen Arbeitsblättern.":::

## <a name="sample-excel-file"></a>Beispieldatei für Excel

Laden Sie <a href="table-of-contents.xlsx">table-of-contents.xlsx</a> für eine einsatzbereite Arbeitsmappe herunter. Fügen Sie das folgende Skript hinzu, und testen Sie das Beispiel selbst!

## <a name="sample-code-create-a-workbook-table-of-contents"></a>Beispielcode: Erstellen eines Arbeitsmappenverzeichniss

```TypeScript
function main(workbook: ExcelScript.Workbook) {
  // Insert a new worksheet at the beginning of the workbook.
  let tocSheet = workbook.addWorksheet();
  tocSheet.setPosition(0);
  tocSheet.setName("Table of Contents");

  // Give the worksheet a title in the sheet.
  tocSheet.getRange("A1").setValue("Table of Contents");
  tocSheet.getRange("A1").getFormat().getFont().setBold(true);

  // Create the table of contents headers.
  let tocRange = tocSheet.getRange("A2:B2")
  tocRange.setValues([["#", "Name"]]);

  // Get the range for the table of contents entries.
  let worksheets = workbook.getWorksheets();
  tocRange = tocRange.getResizedRange(worksheets.length, 0);

  // Loop through all worksheets in the workbook, except the first one.
  for (let i = 1; i < worksheets.length; i++) {
    // Create a row for each worksheet with its index and linked name.
    tocRange.getCell(i, 0).setValue(i);
    tocRange.getCell(i, 1).setHyperlink({
      textToDisplay: worksheets[i].getName(),
      documentReference: `'${worksheets[i].getName()}'!A1`
    });
  };

  // Activate the table of contents worksheet.
  tocSheet.activate();
}
```
