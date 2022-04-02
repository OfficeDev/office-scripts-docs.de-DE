---
title: Hinzufügen von Kommentaren in Excel
description: Erfahren Sie, wie Sie Office Skripts verwenden, um Kommentare in einem Arbeitsblatt hinzuzufügen.
ms.date: 06/29/2021
ms.localizationpriority: medium
ms.openlocfilehash: 101f07fd2f1abcd4120585162dc2b77b8aece91a
ms.sourcegitcommit: 7023b9e23499806901a5ecf8ebc460b76887cca6
ms.translationtype: MT
ms.contentlocale: de-DE
ms.lasthandoff: 03/31/2022
ms.locfileid: "64585597"
---
# <a name="add-comments-in-excel"></a>Hinzufügen von Kommentaren in Excel

In diesem Beispiel wird gezeigt, wie Sie einer Zelle, einschließlich [@mentioning](https://support.microsoft.com/office/90701709-5dc1-41c7-aa48-b01d4a46e8c7) eines Kollegen, Kommentare hinzufügen.

## <a name="example-scenario"></a>Beispielszenario

* Der Teamleiter behält den Schichtplan bei. Der Teamleiter weist dem Schichtdatensatz eine Mitarbeiter-ID zu.
* Das Team leitet den Wunsch, den Mitarbeiter zu benachrichtigen. Durch Hinzufügen eines Kommentars, der den Mitarbeiter @mentions, wird dem Mitarbeiter eine benutzerdefinierte Nachricht aus dem Arbeitsblatt per E-Mail gesendet.
* Anschließend kann der Mitarbeiter die Arbeitsmappe anzeigen und nach Belieben auf den Kommentar antworten.

## <a name="solution"></a>Lösung

1. Das Skript extrahiert Mitarbeiterinformationen aus dem Mitarbeiterarbeitsblatt.
1. Das Skript fügt dann einen Kommentar (einschließlich der entsprechenden Mitarbeiter-E-Mail) zur entsprechenden Zelle im Schichtdatensatz hinzu.
1. Vorhandene Kommentare in der Zelle werden vor dem Hinzufügen des neuen Kommentars entfernt.

## <a name="sample-excel-file"></a>Beispieldatei für Excel

Laden Sie <a href="excel-comments.xlsx">excel-comments.xlsx</a> für eine einsatzbereite Arbeitsmappe herunter. Fügen Sie das folgende Skript hinzu, um das Beispiel selbst auszuprobieren!

## <a name="sample-code-add-comments"></a>Beispielcode: Hinzufügen von Kommentaren

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

## <a name="training-video-add-comments"></a>Schulungsvideo: Hinzufügen von Kommentaren

[Sehen Sie sich an, wie Sie dieses Beispiel auf YouTube durchlaufen](https://youtu.be/CpR78nkaOFw).
