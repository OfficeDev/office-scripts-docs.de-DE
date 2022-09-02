---
title: Hinzufügen von Kommentaren in Excel
description: Erfahren Sie, wie Sie Office-Skripts zum Hinzufügen von Kommentaren in einem Arbeitsblatt verwenden.
ms.date: 06/29/2021
ms.localizationpriority: medium
ms.openlocfilehash: 90f072805e6798a4f9d6e74889ccca15610c87bd
ms.sourcegitcommit: a6504f8b0d6b717457c6e0b5306c35ad3900914e
ms.translationtype: MT
ms.contentlocale: de-DE
ms.lasthandoff: 09/02/2022
ms.locfileid: "67572493"
---
# <a name="add-comments-in-excel"></a>Hinzufügen von Kommentaren in Excel

In diesem Beispiel wird gezeigt, wie Einer Zelle, einschließlich [@mentioning](https://support.microsoft.com/office/90701709-5dc1-41c7-aa48-b01d4a46e8c7) eines Kollegen, Kommentare hinzugefügt werden.

## <a name="example-scenario"></a>Beispielszenario

* Der Teamleiter verwaltet den Schichtplan. Der Teamleiter weist dem Schichtdatensatz eine Mitarbeiter-ID zu.
* Der Teamleiter möchte den Mitarbeiter benachrichtigen. Durch Hinzufügen eines Kommentars, der den Mitarbeiter @mentions, wird dem Mitarbeiter eine benutzerdefinierte Nachricht aus dem Arbeitsblatt per E-Mail übermittelt.
* Anschließend kann der Mitarbeiter die Arbeitsmappe anzeigen und nach belieben auf den Kommentar antworten.

## <a name="solution"></a>Lösung

1. Das Skript extrahiert Mitarbeiterinformationen aus dem Mitarbeiterarbeitsblatt.
1. Das Skript fügt dann einen Kommentar (einschließlich der entsprechenden Mitarbeiter-E-Mail) zur entsprechenden Zelle im Schichtdatensatz hinzu.
1. Vorhandene Kommentare in der Zelle werden entfernt, bevor der neue Kommentar hinzugefügt wird.

## <a name="sample-excel-file"></a>Excel-Beispieldatei

Laden Sie [excel-comments.xlsx](excel-comments.xlsx) für eine sofort einsatzbereite Arbeitsmappe herunter. Fügen Sie das folgende Skript hinzu, um das Beispiel selbst auszuprobieren!

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

[Sehen Sie sich Sudhi Ramamurthy bei diesem Beispiel auf YouTube an](https://youtu.be/CpR78nkaOFw).
