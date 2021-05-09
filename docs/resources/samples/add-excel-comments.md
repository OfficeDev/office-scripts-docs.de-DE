---
title: Hinzufügen von Kommentaren in Excel
description: Erfahren Sie, wie Sie Office Skripts zum Hinzufügen von Kommentaren in einem Arbeitsblatt verwenden.
ms.date: 05/03/2021
localization_priority: Normal
ms.openlocfilehash: e5e5d17c076eceaf06fddeea1a67d31ee3581f31
ms.sourcegitcommit: 763d341857bcb209b2f2c278a82fdb63d0e18f0a
ms.translationtype: MT
ms.contentlocale: de-DE
ms.lasthandoff: 05/08/2021
ms.locfileid: "52285934"
---
# <a name="add-comments-in-excel"></a>Hinzufügen von Kommentaren in Excel

In diesem Beispiel wird gezeigt, wie Sie einer Zelle Kommentare hinzufügen, [einschließlich @mentioning](https://support.microsoft.com/office/90701709-5dc1-41c7-aa48-b01d4a46e8c7) Kollegen.

## <a name="example-scenario"></a>Beispielszenario

* Der Teamleiter behält den Schichtplan bei. Der Teamleiter weist dem Schichtdatensatz eine Mitarbeiter-ID zu.
* Der Teamleiter möchte den Mitarbeiter benachrichtigen. Durch Hinzufügen eines Kommentars, @mentions Mitarbeiter hinzugefügt wird, wird dem Mitarbeiter eine benutzerdefinierte Nachricht aus dem Arbeitsblatt per E-Mail gesendet.
* Anschließend kann der Mitarbeiter die Arbeitsmappe anzeigen und nach Befreundlichkeit auf den Kommentar antworten.

## <a name="solution"></a>Lösung

1. Das Skript extrahiert Mitarbeiterinformationen aus dem Mitarbeiterarbeitsblatt.
1. Das Skript fügt dann der entsprechenden Zelle im Schichtdatensatz einen Kommentar (einschließlich der entsprechenden Mitarbeiter-E-Mail) hinzu.
1. Vorhandene Kommentare in der Zelle werden entfernt, bevor der neue Kommentar hinzugefügt wird.

## <a name="sample-code-add-comments"></a>Beispielcode: Hinzufügen von Kommentaren

Laden Sie die Datei <a href="excel-comments.xlsx">excel-comments.xlsx</a> in diesem Beispiel verwendeten herunter, und testen Sie sie selbst!

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

[Sehen Sie sich an, wie Sudhi Ramamurthy dieses Beispiel auf YouTube durchspazieren.](https://youtu.be/CpR78nkaOFw)
