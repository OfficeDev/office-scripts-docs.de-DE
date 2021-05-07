---
title: 'Office Skriptbeispielszenario: Notenrechner'
description: Ein Beispiel, das die Prozent- und Buchstabennoten für eine Schülerklasse bestimmt.
ms.date: 12/17/2020
localization_priority: Normal
ms.openlocfilehash: e2ef6e7522fc88219bf6ba40900a1ecceecb263b
ms.sourcegitcommit: f7a7aebfb687f2a35dbed07ed62ff352a114525a
ms.translationtype: MT
ms.contentlocale: de-DE
ms.lasthandoff: 05/06/2021
ms.locfileid: "52232697"
---
# <a name="office-scripts-sample-scenario-grade-calculator"></a><span data-ttu-id="0cd3b-103">Office Skriptbeispielszenario: Notenrechner</span><span class="sxs-lookup"><span data-stu-id="0cd3b-103">Office Scripts sample scenario: Grade calculator</span></span>

<span data-ttu-id="0cd3b-104">In diesem Szenario sind Sie ein Kursleiter, der die End-of-Term-Noten jedes Schülers abfüllt.</span><span class="sxs-lookup"><span data-stu-id="0cd3b-104">In this scenario, you're an instructor tallying every student's end-of-term grades.</span></span> <span data-ttu-id="0cd3b-105">Sie haben die Noten für ihre Zuordnungen und Tests während Ihres Wegs eingeben.</span><span class="sxs-lookup"><span data-stu-id="0cd3b-105">You've been entering the scores for their assignments and tests as you go.</span></span> <span data-ttu-id="0cd3b-106">Jetzt ist es an der Zeit, die Schicksale der Schüler zu bestimmen.</span><span class="sxs-lookup"><span data-stu-id="0cd3b-106">Now, it is time to determine the students' fates.</span></span>

<span data-ttu-id="0cd3b-107">Sie entwickeln ein Skript, das die Noten für jede Punktkategorie insgesamt auflisten kann.</span><span class="sxs-lookup"><span data-stu-id="0cd3b-107">You'll develop a script that totals the grades for each point category.</span></span> <span data-ttu-id="0cd3b-108">Anschließend wird jedem Schüler basierend auf der Gesamtsumme eine Briefnote zugewiesen.</span><span class="sxs-lookup"><span data-stu-id="0cd3b-108">It will then assign a letter grade to each student based on the total.</span></span> <span data-ttu-id="0cd3b-109">Um die Genauigkeit sicherzustellen, fügen Sie ein paar Prüfungen hinzu, um zu überprüfen, ob einzelne Bewertungen zu niedrig oder zu hoch sind.</span><span class="sxs-lookup"><span data-stu-id="0cd3b-109">To help ensure accuracy, you'll add a couple checks to see if any individual scores are too low or high.</span></span> <span data-ttu-id="0cd3b-110">Wenn die Punktzahl eines Schülers kleiner als null oder mehr als der mögliche Punktwert ist, markiert das Skript die Zelle mit einer roten Füllung und nicht die Punkte dieses Kursteilnehmers.</span><span class="sxs-lookup"><span data-stu-id="0cd3b-110">If a student's score is less than zero or more than the possible point value, the script will flag the cell with a red fill and not total that student's points.</span></span> <span data-ttu-id="0cd3b-111">Dies ist ein eindeutiger Hinweis darauf, welche Datensätze Sie überprüfen müssen.</span><span class="sxs-lookup"><span data-stu-id="0cd3b-111">This will be a clear indication of which records you need to double-check.</span></span> <span data-ttu-id="0cd3b-112">Sie fügen den Noten auch einige grundlegende Formatierungen hinzu, damit Sie den oberen und unteren Teil der Klasse schnell anzeigen können.</span><span class="sxs-lookup"><span data-stu-id="0cd3b-112">You'll also add some basic formatting to the grades so you can quickly view the top and bottom of the class.</span></span>

## <a name="scripting-skills-covered"></a><span data-ttu-id="0cd3b-113">Abgedeckte Skriptkenntnisse</span><span class="sxs-lookup"><span data-stu-id="0cd3b-113">Scripting skills covered</span></span>

- <span data-ttu-id="0cd3b-114">Zellenformatierung</span><span class="sxs-lookup"><span data-stu-id="0cd3b-114">Cell formatting</span></span>
- <span data-ttu-id="0cd3b-115">Fehlerüberprüfung</span><span class="sxs-lookup"><span data-stu-id="0cd3b-115">Error checking</span></span>
- <span data-ttu-id="0cd3b-116">Reguläre Ausdrücke</span><span class="sxs-lookup"><span data-stu-id="0cd3b-116">Regular expressions</span></span>
- <span data-ttu-id="0cd3b-117">Bedingte Formatierung</span><span class="sxs-lookup"><span data-stu-id="0cd3b-117">Conditional formatting</span></span>

## <a name="setup-instructions"></a><span data-ttu-id="0cd3b-118">Setupanweisungen</span><span class="sxs-lookup"><span data-stu-id="0cd3b-118">Setup instructions</span></span>

1. <span data-ttu-id="0cd3b-119">Laden <a href="grade-calculator.xlsx">grade-calculator.xlsx</a> auf Ihre OneDrive.</span><span class="sxs-lookup"><span data-stu-id="0cd3b-119">Download <a href="grade-calculator.xlsx">grade-calculator.xlsx</a> to your OneDrive.</span></span>

2. <span data-ttu-id="0cd3b-120">Öffnen Sie die Arbeitsmappe mit Excel für das Web.</span><span class="sxs-lookup"><span data-stu-id="0cd3b-120">Open the workbook with Excel for the web.</span></span>

3. <span data-ttu-id="0cd3b-121">Öffnen Sie **auf** der Registerkarte Automatisieren **alle Skripts**.</span><span class="sxs-lookup"><span data-stu-id="0cd3b-121">Under the **Automate** tab, open **All Scripts**.</span></span>

4. <span data-ttu-id="0cd3b-122">Drücken Sie **im Aufgabenbereich Code-Editor** die **Taste Neues Skript,** und fügen Sie das folgende Skript in den Editor ein.</span><span class="sxs-lookup"><span data-stu-id="0cd3b-122">In the **Code Editor** task pane, press **New Script** and paste the following script into the editor.</span></span>

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
      const assignmentMaxMatches = (studentData[0][1] as string).match(/\d+/);
      const midtermMaxMatches = (studentData[0][2] as string).match(/\d+/);
      const finalMaxMatches = (studentData[0][3] as string).match(/\d+/);

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
        const total = (studentData[i][1] as number) + (studentData[i][2] as number) + (studentData[i][3] as number);
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
          "#443300",
          "#FFEE22",
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
      let conditionalFormatting: ExcelScript.ConditionalFormat;
      conditionalFormatting = range.addConditionalFormat(ExcelScript.ConditionalFormatType.cellValue);
      conditionalFormatting.getCellValue().getFormat().getFont().setColor(fontColor);
      conditionalFormatting.getCellValue().getFormat().getFill().setColor(fillColor);
      conditionalFormatting.getCellValue().setRule({ formula1, operator });
    }
    ```

5. <span data-ttu-id="0cd3b-123">Benennen Sie das Skript in **Notenrechner um,** und speichern Sie es.</span><span class="sxs-lookup"><span data-stu-id="0cd3b-123">Rename the script to **Grade Calculator** and save it.</span></span>

## <a name="running-the-script"></a><span data-ttu-id="0cd3b-124">Ausführen des Skripts</span><span class="sxs-lookup"><span data-stu-id="0cd3b-124">Running the script</span></span>

<span data-ttu-id="0cd3b-125">Führen Sie das **Skript Notenrechner** im einzigen Arbeitsblatt aus.</span><span class="sxs-lookup"><span data-stu-id="0cd3b-125">Run the **Grade Calculator** script on the only worksheet.</span></span> <span data-ttu-id="0cd3b-126">Das Skript gesamt die Noten und weisen jedem Schüler eine Briefnote zu.</span><span class="sxs-lookup"><span data-stu-id="0cd3b-126">The script will total the grades and assign each student a letter grade.</span></span> <span data-ttu-id="0cd3b-127">Wenn einzelne Noten mehr Punkte haben, als die Zuordnung oder der Test wert ist, wird die beleidigungswürdige Note rot markiert, und die Summe wird nicht berechnet.</span><span class="sxs-lookup"><span data-stu-id="0cd3b-127">If any individual grades have more points than the assignment or test is worth, then the offending grade is marked red and the total is not calculated.</span></span> <span data-ttu-id="0cd3b-128">Außerdem werden alle "A"-Noten grün hervorgehoben, während die Noten "D" und "F" gelb hervorgehoben sind.</span><span class="sxs-lookup"><span data-stu-id="0cd3b-128">Also, any 'A' grades are highlighted in green, while 'D' and 'F' grades are highlighted in yellow.</span></span>

### <a name="before-running-the-script"></a><span data-ttu-id="0cd3b-129">Vor dem Ausführen des Skripts</span><span class="sxs-lookup"><span data-stu-id="0cd3b-129">Before running the script</span></span>

:::image type="content" source="../../images/scenario-grade-calculator-before.png" alt-text="Ein Arbeitsblatt, das Zeilen mit Noten für Schüler und Studenten zeigt":::

### <a name="after-running-the-script"></a><span data-ttu-id="0cd3b-131">Nach dem Ausführen des Skripts</span><span class="sxs-lookup"><span data-stu-id="0cd3b-131">After running the script</span></span>

:::image type="content" source="../../images/scenario-grade-calculator-after.png" alt-text="Ein Arbeitsblatt, das die Schülerergebnisdaten mit ungültigen Zellen in roten Summen für gültige Schülerzeilen zeigt":::
