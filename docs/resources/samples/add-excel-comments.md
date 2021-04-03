---
title: Hinzufügen von Kommentaren in Excel
description: Erfahren Sie, wie Sie Office-Skripts zum Hinzufügen von Kommentaren in einem Arbeitsblatt verwenden.
ms.date: 03/29/2021
localization_priority: Normal
ms.openlocfilehash: aaaf26df6973bd081290b0fbb67edecad8627e53
ms.sourcegitcommit: 5d24e77df70aa2c1c982275d53213c2a9323ff86
ms.translationtype: MT
ms.contentlocale: de-DE
ms.lasthandoff: 04/02/2021
ms.locfileid: "51571532"
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
    const employees = workbook.getWorksheet('Employees').getUsedRange().getTexts();
    console.log(employees); 

    const scheduleSheet = workbook.getWorksheet('Schedule');
    const table = scheduleSheet.getTables()[0];
    const range = table.getRangeBetweenHeaderAndTotal();
    const scheduleData = range.getTexts();

    for (let i=0; i < scheduleData.length; i++) {
      let eId = scheduleData[i][3];

      let employeeInfo = employees.find(e => e[0] === eId);
      if (employeeInfo) {
        console.log("Found a match " + employeeInfo);
        let adminNotes = scheduleData[i][4];
        try { 
          let comment = workbook.getCommentByCell(range.getCell(i, 5));
          comment.delete();
        } catch {
            console.log("Ignore if there is no existing comment in the cell");
        }
        workbook.addComment(range.getCell(i,5), {
          mentions: [{
            email: employeeInfo[1],
            id: 0,
            name: employeeInfo[2]
          }],
          richContent: `<at id=\"0\">${employeeInfo[2]}</at> ${adminNotes}`
        }, ExcelScript.ContentType.mention);        
        
      } else {
        console.log("No match for: " + eId);
      }
    }
    return;
}
```

## <a name="training-video-add-comments"></a>Schulungsvideo: Hinzufügen von Kommentaren

[![Schritt-für-Schritt-Video zum Hinzufügen von Kommentaren in einer Excel-Datei ansehen](../../images/comments-vid.jpg)](https://youtu.be/CpR78nkaOFw "Schrittweises Video zum Hinzufügen von Kommentaren in einer Excel-Datei")
