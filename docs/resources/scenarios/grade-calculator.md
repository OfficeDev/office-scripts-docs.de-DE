---
title: 'Beispielszenario für Office-Skripts: Grade Calculator'
description: Ein Beispiel, das die Prozent-und Buchstaben Noten für eine Kursteilnehmer Klasse bestimmt.
ms.date: 02/20/2020
localization_priority: Normal
ms.openlocfilehash: 0db6f7c116594f7655bfc0adc8f5a79dbbf2a0af
ms.sourcegitcommit: b075eed5a6f275274fbbf6d62633219eac416f26
ms.translationtype: MT
ms.contentlocale: de-DE
ms.lasthandoff: 03/10/2020
ms.locfileid: "42700234"
---
# <a name="office-scripts-sample-scenario-grade-calculator"></a>Beispielszenario für Office-Skripts: Grade Calculator

In diesem Szenario sind Sie ein Kursleiter, der die Auszählung jedes Schülers angeht. Sie haben die Noten für ihre Zuweisungen und Tests eingegeben, während Sie gehen. Nun ist es an der Zeit, die Schicksale der Schüler zu bestimmen.

Sie entwickeln ein Skript, das die Noten für jede Punkte Kategorie summiert. Anschließend wird jedem Schüler basierend auf der Gesamtzahl ein Brief Grad zugewiesen. Um die Genauigkeit sicherzustellen, fügen Sie ein paar Schecks hinzu, um zu sehen, ob einzelne Noten zu niedrig oder hoch sind. Wenn die Punktzahl eines Schülers kleiner als 0 (null) oder größer als der mögliche Punktwert ist, wird die Zelle durch das Skript mit einer roten Füllung markiert und nicht mit der Summe der Punkte des Schülers. Dies ist ein klarer Hinweis darauf, welche Datensätze Sie doppelt überprüfen müssen. Sie fügen auch einige grundlegende Formatierungen zu den Noten hinzu, damit Sie den oberen und unteren Teil der Klasse schnell anzeigen können.

## <a name="scripting-skills-covered"></a>Abgedeckte Skript Fertigkeiten

- Zellenformatierung
- Fehlerüberprüfung
- Reguläre Ausdrücke

## <a name="setup-instructions"></a>Setup Anweisungen

1. Laden Sie <a href="grade-calculator.xlsx">Grade-Calculator. xlsx</a> in Ihre OneDrive herunter.

2. Öffnen Sie die Arbeitsmappe mit Excel für das Internet.

3. Öffnen Sie auf der Registerkarte **automatisieren** den **Code-Editor**.

4. Klicken Sie im Aufgabenbereich **Code-Editor** auf **Neues Skript** , und fügen Sie das folgende Skript in den Editor ein.

    ```TypeScript
    async function main(context: Excel.RequestContext) {
      // Get the number of student record rows.
      let sheet = context.workbook.worksheets.getActiveWorksheet();
      let studentsRange = sheet.getUsedRange().load("values, rowCount");
      await context.sync();
      console.log("Total students: " + (studentsRange.rowCount - 1));

      // Clean up any formatting from previous runs of the script.
      studentsRange.clear(Excel.ClearApplyTo.formats);
      studentsRange.getColumn(4).getCell(0, 0).getRowsBelow(studentsRange.rowCount - 1).clear(Excel.ClearApplyTo.all);
      studentsRange.getColumn(5).getCell(0, 0).getRowsBelow(studentsRange.rowCount - 1).clear(Excel.ClearApplyTo.all);
      await context.sync();

      // Parse the headers for the maximum possible scores for each category.
      // The format is `category (score)`.
      let assignmentsMax = studentsRange.values[0][1].match(/\d+/)[0];
      let midTermMax = studentsRange.values[0][2].match(/\d+/)[0];
      let finalsMax = studentsRange.values[0][3].match(/\d+/)[0];
      console.log("Assignments max score:" + assignmentsMax);
      console.log("Mid-term max score: " + midTermMax);
      console.log("Final max score: " + finalsMax);

      // Look at every student row.
      for (let i = 1; i < studentsRange.values.length; i++) {
        let row = studentsRange.values[i];
        let total = row[1] + row[2] + row[3];
        let valid = true;

        // Look for any records that are too low or too high.
        if (row[1] < 0 || row[1] > assignmentsMax) {
          studentsRange.getCell(i, 1).format.fill.color = "Red";
          valid = false;
        }
        if (row[2] < 0 || row[2] > midTermMax) {
          studentsRange.getCell(i, 2).format.fill.color = "Red";
          valid = false;
        }
        if (row[3] < 0 || row[3] > finalsMax) {
          studentsRange.getCell(i, 3).format.fill.color = "Red";
          valid = false;
        }

        // If the scores are valid, total that student's points and assign them a letter grade.
        if (valid) {
          let grade: string;
          switch (true) {
            case total < 60:
              grade = "E";
              break;
            case total < 70:
              grade = "D";
              break;
            case total < 80:
              grade = "C";
              break;
            case total < 90:
              grade = "B";
              break;
            default:
              grade = "A";
              break;
          }

          studentsRange.getCell(i, 4).values = [[total]];
          studentsRange.getCell(i, 5).values = [[grade]];

          // Highlight excellent students and those in need of attention.
          if (grade === "A") {
            studentsRange.getCell(i, 5).format.fill.color = "Green";
          } else if (grade === "E" || grade === "D") {
            studentsRange.getCell(i, 5).format.fill.color = "Orange";
          }
        }
      }

      studentsRange.getColumn(5).format.horizontalAlignment = "Center";
    }
    ```

5. Benennen Sie das Skript in den **Grade Calculator** um, und speichern Sie es.

## <a name="running-the-script"></a>Ausführen des Skripts

Führen Sie das **Grade Calculator** -Skript auf dem einzigen Arbeitsblatt aus. Das Skript summiert die Noten und weist jedem Schüler eine Note zu. Wenn einzelne Noten mehr Punkte aufweisen als die Zuordnung oder der Test Wert ist, wird die betroffene Note rot markiert, und die Summe wird nicht berechnet.

### <a name="before-running-the-script"></a>Vor dem Ausführen des Skripts

![Ein Arbeitsblatt, in dem Zeilen mit Partituren für Schüler angezeigt werden.](../../images/scenario-grade-calculator-before.png)

### <a name="after-running-the-script"></a>Nach dem Ausführen des Skripts

![Ein Arbeitsblatt, in dem die Kursteilnehmerdaten mit ungültigen Zellen in roten Gesamtzahlen für gültige Schüler Zeilen angezeigt werden.](../../images/scenario-grade-calculator-after.png)
