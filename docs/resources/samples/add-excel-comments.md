---
title: Hinzufügen von Kommentaren in Excel
description: Erfahren Sie, wie Sie Office Skripts verwenden, um Kommentare in einem Arbeitsblatt hinzuzufügen.
ms.date: 06/29/2021
localization_priority: Normal
ms.openlocfilehash: 77e308d020281c71751e2652f8dbaec00c263e44
ms.sourcegitcommit: 211c157ca746e266eeb079f5fa1925a1e35ab702
ms.translationtype: MT
ms.contentlocale: de-DE
ms.lasthandoff: 07/07/2021
ms.locfileid: "53313911"
---
# <a name="add-comments-in-excel"></a><span data-ttu-id="fa0d1-103">Hinzufügen von Kommentaren in Excel</span><span class="sxs-lookup"><span data-stu-id="fa0d1-103">Add comments in Excel</span></span>

<span data-ttu-id="fa0d1-104">In diesem Beispiel wird gezeigt, wie Sie einer Zelle, einschließlich [@mentioning](https://support.microsoft.com/office/90701709-5dc1-41c7-aa48-b01d4a46e8c7) eines Kollegen, Kommentare hinzufügen.</span><span class="sxs-lookup"><span data-stu-id="fa0d1-104">This sample shows how to add comments to a cell including [@mentioning](https://support.microsoft.com/office/90701709-5dc1-41c7-aa48-b01d4a46e8c7) a colleague.</span></span>

## <a name="example-scenario"></a><span data-ttu-id="fa0d1-105">Beispielszenario</span><span class="sxs-lookup"><span data-stu-id="fa0d1-105">Example scenario</span></span>

* <span data-ttu-id="fa0d1-106">Der Teamleiter behält den Schichtplan bei.</span><span class="sxs-lookup"><span data-stu-id="fa0d1-106">The team lead maintains the shift schedule.</span></span> <span data-ttu-id="fa0d1-107">Der Teamleiter weist dem Schichtdatensatz eine Mitarbeiter-ID zu.</span><span class="sxs-lookup"><span data-stu-id="fa0d1-107">The team lead assigns an employee ID to the shift record.</span></span>
* <span data-ttu-id="fa0d1-108">Das Team leitet den Wunsch, den Mitarbeiter zu benachrichtigen.</span><span class="sxs-lookup"><span data-stu-id="fa0d1-108">The team lead wishes to notify the employee.</span></span> <span data-ttu-id="fa0d1-109">Durch Hinzufügen eines Kommentars, der den Mitarbeiter @mentions, wird dem Mitarbeiter eine benutzerdefinierte Nachricht aus dem Arbeitsblatt per E-Mail gesendet.</span><span class="sxs-lookup"><span data-stu-id="fa0d1-109">By adding a comment that @mentions the employee, the employee is emailed with a custom message from the worksheet.</span></span>
* <span data-ttu-id="fa0d1-110">Anschließend kann der Mitarbeiter die Arbeitsmappe anzeigen und nach Belieben auf den Kommentar antworten.</span><span class="sxs-lookup"><span data-stu-id="fa0d1-110">Subsequently, the employee can view the workbook and respond to the comment at their convenience.</span></span>

## <a name="solution"></a><span data-ttu-id="fa0d1-111">Lösung</span><span class="sxs-lookup"><span data-stu-id="fa0d1-111">Solution</span></span>

1. <span data-ttu-id="fa0d1-112">Das Skript extrahiert Mitarbeiterinformationen aus dem Mitarbeiterarbeitsblatt.</span><span class="sxs-lookup"><span data-stu-id="fa0d1-112">The script extracts employee information from the employee worksheet.</span></span>
1. <span data-ttu-id="fa0d1-113">Das Skript fügt dann einen Kommentar (einschließlich der entsprechenden Mitarbeiter-E-Mail) zur entsprechenden Zelle im Schichtdatensatz hinzu.</span><span class="sxs-lookup"><span data-stu-id="fa0d1-113">The script then adds a comment (including the relevant employee email) to the appropriate cell in the shift record.</span></span>
1. <span data-ttu-id="fa0d1-114">Vorhandene Kommentare in der Zelle werden vor dem Hinzufügen des neuen Kommentars entfernt.</span><span class="sxs-lookup"><span data-stu-id="fa0d1-114">Existing comments in the cell are removed before adding the new comment.</span></span>

## <a name="sample-excel-file"></a><span data-ttu-id="fa0d1-115">Beispieldatei für Excel</span><span class="sxs-lookup"><span data-stu-id="fa0d1-115">Sample Excel file</span></span>

<span data-ttu-id="fa0d1-116">Laden Sie <a href="excel-comments.xlsx">excel-comments.xlsx</a> für eine einsatzbereite Arbeitsmappe herunter.</span><span class="sxs-lookup"><span data-stu-id="fa0d1-116">Download <a href="excel-comments.xlsx">excel-comments.xlsx</a> for a ready-to-use workbook.</span></span> <span data-ttu-id="fa0d1-117">Fügen Sie das folgende Skript hinzu, um das Beispiel selbst auszuprobieren!</span><span class="sxs-lookup"><span data-stu-id="fa0d1-117">Add the following script to try the sample yourself!</span></span>

## <a name="sample-code-add-comments"></a><span data-ttu-id="fa0d1-118">Beispielcode: Hinzufügen von Kommentaren</span><span class="sxs-lookup"><span data-stu-id="fa0d1-118">Sample code: Add comments</span></span>

```TypeScript
function main(workbook: ExcelScript.Workbook) {
  // Get the list of employees.
  const employees = workbook.getWorksheet('Employees').getUsedRange().getTexts();
  console.log(employees); 
  
  // Get the schedule information from the schedule table.
  const scheduleSheet = workbook.getWorksheet('Schedule');
  const table = scheduleSheet.getTables()[0];
  const range = table.getRangeBetweenHeaderAndTotal();
  const scheduleData = range.getTexts();

  // Look through the schedule for a matching employee.
  for (let i = 0; i < scheduleData.length; i++) {
    let employeeId = scheduleData[i][3];

    // Compare the employee ID in the schedule against the employee information table.
    let employeeInfo = employees.find(employeeRow => employeeRow[0] === employeeId);
    if (employeeInfo) {
      console.log("Found a match " + employeeInfo);
      let adminNotes = scheduleData[i][4];

      // Look for and delete old comments, so we avoid conflicts.
      let comment = workbook.getCommentByCell(range.getCell(i, 5));
      if (comment) {
        comment.delete();
      }

      // Add a comment using the admin notes as the text.
      workbook.addComment(range.getCell(i,5), {
        mentions: [{
          email: employeeInfo[1],
          id: 0, // This ID maps this mention to the `id=0` text in the comment.
          name: employeeInfo[2]
        }],
        richContent: `<at id=\"0\">${employeeInfo[2]}</at> ${adminNotes}`
      }, ExcelScript.ContentType.mention);        
      
    } else {
      console.log("No match for: " + employeeId);
    }
  }
}
```

## <a name="training-video-add-comments"></a><span data-ttu-id="fa0d1-119">Schulungsvideo: Hinzufügen von Kommentaren</span><span class="sxs-lookup"><span data-stu-id="fa0d1-119">Training video: Add comments</span></span>

<span data-ttu-id="fa0d1-120">[Sehen Sie sich dieses Beispiel auf YouTube an.](https://youtu.be/CpR78nkaOFw)</span><span class="sxs-lookup"><span data-stu-id="fa0d1-120">[Watch Sudhi Ramamurthy walk through this sample on YouTube](https://youtu.be/CpR78nkaOFw).</span></span>
