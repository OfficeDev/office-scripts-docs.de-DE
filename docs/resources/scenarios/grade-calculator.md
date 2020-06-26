---
title: 'Beispielszenario für Office-Skripts: Grade Calculator'
description: Ein Beispiel, das die Prozent-und Buchstaben Noten für eine Kursteilnehmer Klasse bestimmt.
ms.date: 06/01/2020
localization_priority: Normal
ms.openlocfilehash: 6f8e3db756c72cf1d0e2f774ccd819c041f0c42d
ms.sourcegitcommit: aec3c971c6640429f89b6bb99d2c95ea06725599
ms.translationtype: MT
ms.contentlocale: de-DE
ms.lasthandoff: 06/25/2020
ms.locfileid: "44878640"
---
# <a name="office-scripts-sample-scenario-grade-calculator"></a>Beispielszenario für Office-Skripts: Grade Calculator

In diesem Szenario sind Sie ein Kursleiter, der die Auszählung jedes Schülers angeht. Sie haben die Noten für ihre Zuweisungen und Tests eingegeben, während Sie gehen. Nun ist es an der Zeit, die Schicksale der Schüler zu bestimmen.

Sie entwickeln ein Skript, das die Noten für jede Punkte Kategorie summiert. Anschließend wird jedem Schüler basierend auf der Gesamtzahl ein Brief Grad zugewiesen. Um die Genauigkeit sicherzustellen, fügen Sie ein paar Schecks hinzu, um zu sehen, ob einzelne Noten zu niedrig oder hoch sind. Wenn die Punktzahl eines Schülers kleiner als 0 (null) oder größer als der mögliche Punktwert ist, wird die Zelle durch das Skript mit einer roten Füllung markiert und nicht mit der Summe der Punkte des Schülers. Dies ist ein klarer Hinweis darauf, welche Datensätze Sie doppelt überprüfen müssen. Sie fügen auch einige grundlegende Formatierungen zu den Noten hinzu, damit Sie den oberen und unteren Teil der Klasse schnell anzeigen können.

## <a name="scripting-skills-covered"></a>Abgedeckte Skript Fertigkeiten

- Zellenformatierung
- Fehlerüberprüfung
- Reguläre Ausdrücke
- Bedingte Formatierung

## <a name="setup-instructions"></a>Setup Anweisungen

1. Laden Sie <a href="grade-calculator.xlsx">grade-calculator.xlsx</a> auf Ihre OneDrive herunter.

2. Öffnen Sie die Arbeitsmappe mit Excel für das Internet.

3. Öffnen Sie auf der Registerkarte **automatisieren** den **Code-Editor**.

4. Klicken Sie im Aufgabenbereich **Code-Editor** auf **Neues Skript** , und fügen Sie das folgende Skript in den Editor ein.

    ```TypeScript
    function main(workbook: ExcelScript.Workbook) {
      // Get the worksheet and validate the data.
      let studentsRange = workbook.getActiveWorksheet().getUsedRange();
      if (studentsRange.getColumnCount() !== 6) {
        throw new Error(`The required columns are not present. Expected column headers: "Student ID | Assignment score | Mid-term | Final | Total | Grade"`);
      }

      let studentData = studentsRange.getValues();

      // Clear the total and grade columns.
      studentsRange.getColumn(4).getCell(1, 0).getAbsoluteResizedRange(studentData.length - 1, 2).clear();

      // Clear all conditional formatting.
      workbook.getActiveWorksheet().getUsedRange().clearAllConditionalFormats();

      // Use regular expressions to read the max score from the assignment, mid-term, and final scores columns.
      let maxScores: string[] = [];
      const assignmentMaxMatches = studentData[0][1].match(/\d+/);
      const midtermMaxMatches = studentData[0][2].match(/\d+/);
      const finalMaxMatches = studentData[0][3].match(/\d+/);

      // Check the matches happened before proceeding.
      if (!(assignmentMaxMatches && midtermMaxMatches && finalMaxMatches)) {
        throw new Error(`The scores are not present in the column headers. Expected format: "Assignments (n)|Mid-term (n)|Final (n)"`);
      }

      // Use the first (and only) match from the regular expressions as the max scores.
      maxScores = [assignmentMaxMatches[0], midtermMaxMatches[0], finalMaxMatches[0]];

      // Set conditional formatting for each of the assignment, mid-term, and final scores columns.
      maxScores.forEach((score, i) => {
        let range = studentsRange.getColumn(i + 1).getCell(0, 0).getRowsBelow(studentData.length - 1);
        setCellValueConditionalFormatting(
          score,
          range,
          "#9C0006",
          "#FFC7CE",
          ExcelScript.ConditionalCellValueOperator.greaterThan
        )
      });

      // Store the current range information to avoid calling the workbook in the loop.
      let studentsRangeFormulas = studentsRange.getColumn(4).getFormulasR1C1();
      let studentsRangeValues = studentsRange.getColumn(5).getValues();

      /* Iterate over each of the student rows and compute the total score and letter grade.
      * Note that iterator starts at index 1 to skip first (header) row.
      */
      for (let i = 1; i < studentData.length; i++) {
        // If any of the scores are invalid, skip processing it.
        if (studentData[i][1] > maxScores[0] ||
          studentData[i][2] > maxScores[1] ||
          studentData[i][3] > maxScores[2]) {
          continue;
        }
        const total = studentData[i][1] + studentData[i][2] + studentData[i][3];
        let grade: string;
        switch (true) {
          case total < 60:
            grade = "F";
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

        // Set total score formula.
        studentsRangeFormulas[i][0] = '=RC[-2]+RC[-1]';
        // Set grade cell.
        studentsRangeValues[i][0] = grade;
      }

      // Set the formulas and values outside the loop.
      studentsRange.getColumn(4).setFormulasR1C1(studentsRangeFormulas);
      studentsRange.getColumn(5).setValues(studentsRangeValues);

      // Put a conditional formatting on the grade column.
      let totalRange = studentsRange.getColumn(5).getCell(0, 0).getRowsBelow(studentData.length - 1);
      setCellValueConditionalFormatting(
        "A",
        totalRange,
        "#001600",
        "#C6EFCE",
        ExcelScript.ConditionalCellValueOperator.equalTo
      );
      ["D", "F"].forEach((grade) => {
        setCellValueConditionalFormatting(
          grade,
          totalRange,
          "#9C0006",
          "#FFC7CE",
          ExcelScript.ConditionalCellValueOperator.equalTo
        );
      })
      // Center the grade column.
      studentsRange.getColumn(5).getFormat().setHorizontalAlignment(ExcelScript.HorizontalAlignment.center);
    }

    /**
     * Helper function to apply conditional formatting.
     * @param value Cell value to use in conditional formatting formula1.
     * @param range Target range.
     * @param fontColor Font color to use.
     * @param fillColor Fill color to use.
     * @param operator Operator to use in conditional formatting.
     */
    function setCellValueConditionalFormatting(
      value: string,
      range: ExcelScript.Range,
      fontColor: string,
      fillColor: string,
      operator: ExcelScript.ConditionalCellValueOperator) {
      // Determine the formula1 based on the type of value parameter.
      let formula1: string;
      if (isNaN(Number(value))) {
        // For cell value equalTo rule, use this format: formula1: "=\"A\"",
        formula1 = `=\"${value}\"`;
      } else {
        // For number input (greater-than or less-than rules), just append '='.
        formula1 = `=${value}`;
      }

      // Apply conditional formatting.
      let conditionalFormatting : ExcelScript.ConditionalFormat;
      conditionalFormatting = range.addConditionalFormat(ExcelScript.ConditionalFormatType.cellValue);
      conditionalFormatting.getCellValue().getFormat().getFont().setColor(fontColor);
      conditionalFormatting.getCellValue().getFormat().getFill().setColor(fillColor);
      conditionalFormatting.getCellValue().setRule({formula1, operator});
    }
    ```

5. Benennen Sie das Skript in den **Grade Calculator** um, und speichern Sie es.

## <a name="running-the-script"></a>Ausführen des Skripts

Führen Sie das **Grade Calculator** -Skript auf dem einzigen Arbeitsblatt aus. Das Skript summiert die Noten und weist jedem Schüler eine Note zu. Wenn einzelne Noten mehr Punkte aufweisen als die Zuordnung oder der Test Wert ist, wird die betroffene Note rot markiert, und die Summe wird nicht berechnet.

### <a name="before-running-the-script"></a>Vor dem Ausführen des Skripts

![Ein Arbeitsblatt, in dem Zeilen mit Partituren für Schüler angezeigt werden.](../../images/scenario-grade-calculator-before.png)

### <a name="after-running-the-script"></a>Nach dem Ausführen des Skripts

![Ein Arbeitsblatt, in dem die Kursteilnehmerdaten mit ungültigen Zellen in roten Gesamtzahlen für gültige Schüler Zeilen angezeigt werden.](../../images/scenario-grade-calculator-after.png)
