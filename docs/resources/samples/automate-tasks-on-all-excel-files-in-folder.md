---
title: Ausführen eines Skripts für alle Excel-Dateien in einem Ordner
description: 'Erfahren Sie, wie Sie ein Skript für alle #A0 in einem Ordner auf OneDrive for Business ausführen.'
ms.date: 03/31/2021
localization_priority: Normal
ms.openlocfilehash: a11876e8241a069a7c640bbcf2c36b4842d3bd90
ms.sourcegitcommit: 5d24e77df70aa2c1c982275d53213c2a9323ff86
ms.translationtype: MT
ms.contentlocale: de-DE
ms.lasthandoff: 04/02/2021
ms.locfileid: "51571490"
---
# <a name="run-a-script-on-all-excel-files-in-a-folder"></a><span data-ttu-id="04e1e-103">Ausführen eines Skripts für alle Excel-Dateien in einem Ordner</span><span class="sxs-lookup"><span data-stu-id="04e1e-103">Run a script on all Excel files in a folder</span></span>

<span data-ttu-id="04e1e-104">Dieses Projekt führt eine Reihe von Automatisierungsaufgaben für alle Dateien aus, die sich in einem Ordner auf OneDrive for Business befinden.</span><span class="sxs-lookup"><span data-stu-id="04e1e-104">This project performs a set of automation tasks on all files situated in a folder on OneDrive for Business.</span></span> <span data-ttu-id="04e1e-105">Es kann auch für einen SharePoint-Ordner verwendet werden.</span><span class="sxs-lookup"><span data-stu-id="04e1e-105">It could also be used on a SharePoint folder.</span></span>
<span data-ttu-id="04e1e-106">Es führt Berechnungen für die Excel-Dateien durch, fügt Formatierungen hinzu und fügt einen Kommentar ein, der [@mentions](https://support.microsoft.com/office/90701709-5dc1-41c7-aa48-b01d4a46e8c7) kollegen.</span><span class="sxs-lookup"><span data-stu-id="04e1e-106">It performs calculations on the Excel files, adds formatting, and inserts a comment that [@mentions](https://support.microsoft.com/office/90701709-5dc1-41c7-aa48-b01d4a46e8c7) a colleague.</span></span>

## <a name="sample-code-add-formatting-and-insert-comment"></a><span data-ttu-id="04e1e-107">Beispielcode: Hinzufügen von Formatierung und Einfügen eines Kommentars</span><span class="sxs-lookup"><span data-stu-id="04e1e-107">Sample code: Add formatting and insert comment</span></span>

<span data-ttu-id="04e1e-108">Laden Sie die Datei <a href="https://github.com/OfficeDev/office-scripts-docs/blob/master/docs/resources/samples/highlight-alert-excel-files.zip?raw=true">highlight-alert-excel-files.zip</a>, extrahieren Sie die Dateien in einen Ordner mit dem Titel **Sales,** der in diesem Beispiel verwendet wird, und testen Sie sie selbst!</span><span class="sxs-lookup"><span data-stu-id="04e1e-108">Download the file <a href="https://github.com/OfficeDev/office-scripts-docs/blob/master/docs/resources/samples/highlight-alert-excel-files.zip?raw=true">highlight-alert-excel-files.zip</a>, extract the files to a folder titled **Sales** used in this sample, and try it out yourself!</span></span>

```TypeScript
function main(workbook: ExcelScript.Workbook) {
  let table1 = workbook.getTable("Table1");
  const rowCount = table1.getRowCount();
  if (rowCount === 0) {
    return;
  }
  workbook.getApplication().calculate(ExcelScript.CalculationType.full);

  const amountDueCol = table1.getColumnByName('Amount Due');
  const amountDueValues = amountDueCol.getRangeBetweenHeaderAndTotal().getValues();

  let highestValue = amountDueValues[0][0];
  let row = 0;
  for (let i = 1; i < amountDueValues.length; i++) {
    if (amountDueValues[i][0] > highestValue) {
      highestValue = amountDueValues[i][0];
      row = i;
    }
  }
  // Set fill color to FFFF00 for range in table Table1 cell in row 0 on column "Amount due".
  table1.getColumn("Amount due")
    .getRangeBetweenHeaderAndTotal()
    .getRow(row)
    .getFormat()
    .getFill()
    .setColor("FFFF00");
  let selectedSheet = workbook.getActiveWorksheet();
  // Insert comment at cell InvoiceAmounts!F2.
  workbook.addComment(table1.getColumn("Amount due")
    .getRangeBetweenHeaderAndTotal()
    .getRow(row), {
    mentions: [{
      email: "AdeleV@M365x904181.OnMicrosoft.com",
      id: 0,
      name: "Adele Vance"
    }],
    richContent: "<at id=\"0\">Adele Vance</at> Please review this amount"
  }, ExcelScript.ContentType.mention);
}
```

## <a name="training-video-run-a-script-on-all-excel-files-in-a-folder"></a><span data-ttu-id="04e1e-109">Schulungsvideo: Ausführen eines Skripts für alle Excel-Dateien in einem Ordner</span><span class="sxs-lookup"><span data-stu-id="04e1e-109">Training video: Run a script on all Excel files in a folder</span></span>

<span data-ttu-id="04e1e-110">[Sehen Sie sich schritt-für-Schritt-Video](https://youtu.be/xMg711o7k6w) an, wie Sie ein Skript für alle #A0 in einem OneDrive for Business- oder #A1 ausführen.</span><span class="sxs-lookup"><span data-stu-id="04e1e-110">[Watch step-by-step video](https://youtu.be/xMg711o7k6w) on how to run a script on all Excel files in a OneDrive for Business or SharePoint folder.</span></span>
