---
title: Querverweis und Formatieren einer Excel-Datei
description: Erfahren Sie, wie Sie Office-Skripts und Power Automate zum Querverweisen und Formatieren einer Excel-Datei verwenden.
ms.date: 03/18/2021
localization_priority: Normal
ms.openlocfilehash: 287de604733b7e6a126d0c81cb4e23351e558c61
ms.sourcegitcommit: 5d24e77df70aa2c1c982275d53213c2a9323ff86
ms.translationtype: MT
ms.contentlocale: de-DE
ms.lasthandoff: 04/02/2021
ms.locfileid: "51571514"
---
# <a name="cross-reference-and-format-an-excel-file"></a><span data-ttu-id="80c37-103">Querverweis und Formatieren einer Excel-Datei</span><span class="sxs-lookup"><span data-stu-id="80c37-103">Cross-reference and format an Excel file</span></span>

<span data-ttu-id="80c37-104">Diese Lösung zeigt, wie zwei Excel-Dateien mithilfe von Office-Skripts und Power Automate querverweisen und formatiert werden können.</span><span class="sxs-lookup"><span data-stu-id="80c37-104">This solution shows how two Excel files can be cross-referenced and formatted using Office Scripts and Power Automate.</span></span>

<span data-ttu-id="80c37-105">Das Projekt erreicht Folgendes:</span><span class="sxs-lookup"><span data-stu-id="80c37-105">The project achieves the following:</span></span>

1. <span data-ttu-id="80c37-106">Extrahiert Ereignisdaten <a href="events.xlsx"> aus </a>events.xlsxmit einer Ausführungsskriptaktion.</span><span class="sxs-lookup"><span data-stu-id="80c37-106">Extracts event data from <a href="events.xlsx">events.xlsx</a> using one Run script action.</span></span>
1. <span data-ttu-id="80c37-107">Übergibt diese Daten an die zweite Excel-Datei, die Ereignistransaktionsdaten enthält, und verwendet diese Daten zur grundlegenden Überprüfung von Daten und formatierungen fehlender oder falscher Daten mithilfe von Office-Skripts.</span><span class="sxs-lookup"><span data-stu-id="80c37-107">Passes that data to the second Excel file containing event transaction data and uses that data to do basic validation of data and formatting of missing or incorrect data using Office Scripts.</span></span>
1. <span data-ttu-id="80c37-108">E-Mails des Ergebnisses an einen Prüfer.</span><span class="sxs-lookup"><span data-stu-id="80c37-108">Emails the result to a reviewer.</span></span>

<span data-ttu-id="80c37-109">Weitere Informationen finden Sie unter [Querverweis und Formatieren von zwei Excel-Dateien mithilfe von Office-Skripts.](https://powerusers.microsoft.com/t5/Power-Automate-Cookbook/Cross-Reference-and-formatting-two-Excel-files-using-Office/td-p/728535)</span><span class="sxs-lookup"><span data-stu-id="80c37-109">For further details, see [Cross Reference and formatting two Excel files using Office Scripts](https://powerusers.microsoft.com/t5/Power-Automate-Cookbook/Cross-Reference-and-formatting-two-Excel-files-using-Office/td-p/728535).</span></span>

## <a name="sample-excel-files"></a><span data-ttu-id="80c37-110">Beispiel-Excel-Dateien</span><span class="sxs-lookup"><span data-stu-id="80c37-110">Sample Excel files</span></span>

<span data-ttu-id="80c37-111">Laden Sie die folgenden Dateien herunter, die in dieser Lösung verwendet werden, um es selbst auszuprobieren.</span><span class="sxs-lookup"><span data-stu-id="80c37-111">Download the following files used in this solution to try it out yourself!</span></span>

1. <span data-ttu-id="80c37-112"><a href="events.xlsx">events.xlsx</a></span><span class="sxs-lookup"><span data-stu-id="80c37-112"><a href="events.xlsx">events.xlsx</a></span></span>
1. <span data-ttu-id="80c37-113"><a href="event-transactions.xlsx">event-transactions.xlsx</a></span><span class="sxs-lookup"><span data-stu-id="80c37-113"><a href="event-transactions.xlsx">event-transactions.xlsx</a></span></span>

## <a name="sample-code-get-event-data"></a><span data-ttu-id="80c37-114">Beispielcode: Get event data</span><span class="sxs-lookup"><span data-stu-id="80c37-114">Sample code: Get event data</span></span>

```TypeScript
function main(workbook: ExcelScript.Workbook): EventData[] {
    let table = workbook.getWorksheet('Keys').getTables()[0];
    let range = table.getRangeBetweenHeaderAndTotal();
    let rows = range.getValues();
    let records: EventData[] = [];
    for (let row of rows) {
        let [event, date, location, capacity] = row;
        records.push({
            event: event as string,
            date: date as number, 
            location: location as string,
            capacity: capacity as number
        })
    }
    console.log(JSON.stringify(records))
    return records;
}

interface EventData {
    event: string
    date: number
    location: string
    capacity: number
}
```

## <a name="sample-code-validate-event-transactions"></a><span data-ttu-id="80c37-115">Beispielcode: Überprüfen von Ereignistransaktionen</span><span class="sxs-lookup"><span data-stu-id="80c37-115">Sample code: Validate event transactions</span></span>

```TypeScript
function main(workbook: ExcelScript.Workbook, keys: string): string {
    let table = workbook.getWorksheet('Transactions').getTables()[0];
    let range = table.getRangeBetweenHeaderAndTotal();
    range.clear(ExcelScript.ClearApplyTo.formats);
  
    let overallMatch = true;
  
    table.getColumnByName('Date').getRangeBetweenHeaderAndTotal().setNumberFormatLocal("yyyy-mm-dd;@");
    table.getColumnByName('Capacity').getRangeBetweenHeaderAndTotal().getFormat()
      .setHorizontalAlignment(ExcelScript.HorizontalAlignment.center);
    let rows = range.getValues();
    let keysObject = JSON.parse(keys) as EventData[];
    for (let i=0; i < rows.length; i++){
      let row = rows[i];
      let [event, date, location, capacity] = row;
      let match = false;
      for (let keyObject of keysObject){
        if (keyObject.event === event) {
          match = true;
          if (keyObject.date !== date) {
            overallMatch = false;
            range.getCell(i, 1).getFormat()
              .getFill()
              .setColor("FFFF00");
          }
          if (keyObject.location !== location) {
            overallMatch = false;
            range.getCell(i, 2).getFormat()
              .getFill()
              .setColor("FFFF00");
          }
          if (keyObject.capacity !== capacity) {
            overallMatch = false;
            range.getCell(i, 3).getFormat()
              .getFill()
              .setColor("FFFF00");
          }   
          break;             
        }
      }
      if (!match) {
        overallMatch = false;
        range.getCell(i, 0).getFormat()
          .getFill()
          .setColor("FFFF00");      
      }
  
    }
    let returnString = "All the data is in the right order.";
    if (overallMatch === false) {
      returnString = "Mismatch found. Data requires your review.";
    }
    console.log("Returning: " + returnString);
    return returnString;
}

interface EventData {
event: string
date: number
location: string
capacity: number
}
```

## <a name="training-video-cross-reference-and-format-an-excel-file"></a><span data-ttu-id="80c37-116">Schulungsvideo: Querverweis und Formatieren einer Excel-Datei</span><span class="sxs-lookup"><span data-stu-id="80c37-116">Training video: Cross-reference and format an Excel file</span></span>

<span data-ttu-id="80c37-117">[![Schritt-für-Schritt-Video zum Querverweis und Formatieren einer Excel-Datei ansehen](../../images/cross-ref-tables-vid.jpg)](https://youtu.be/dVwqBf483qo "Schrittweises Video zum Querverweis und Formatieren einer Excel-Datei")</span><span class="sxs-lookup"><span data-stu-id="80c37-117">[![Watch step-by-step video on how to cross-reference and format an Excel file](../../images/cross-ref-tables-vid.jpg)](https://youtu.be/dVwqBf483qo "Step-by-step video on how to cross-reference and format an Excel file")</span></span>
