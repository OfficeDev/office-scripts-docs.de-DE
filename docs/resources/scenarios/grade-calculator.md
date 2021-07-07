---
title: 'Office Skript-Beispielszenario: Notenrechner'
description: Ein Beispiel, das den Prozentsatz und die Buchstabennoten für einen Kurs von Schülern bestimmt.
ms.date: 06/29/2021
localization_priority: Normal
ms.openlocfilehash: 2d98e68f37418ade238a707cb74cc7ccf47e8f59
ms.sourcegitcommit: 211c157ca746e266eeb079f5fa1925a1e35ab702
ms.translationtype: MT
ms.contentlocale: de-DE
ms.lasthandoff: 07/07/2021
ms.locfileid: "53313792"
---
# <a name="office-scripts-sample-scenario-grade-calculator"></a><span data-ttu-id="458bd-103">Office Skript-Beispielszenario: Notenrechner</span><span class="sxs-lookup"><span data-stu-id="458bd-103">Office Scripts sample scenario: Grade calculator</span></span>

<span data-ttu-id="458bd-104">In diesem Szenario sind Sie Kursleiter, der die Abschlussnoten jedes Kursteilnehmers anhört.</span><span class="sxs-lookup"><span data-stu-id="458bd-104">In this scenario, you're an instructor tallying every student's end-of-term grades.</span></span> <span data-ttu-id="458bd-105">Sie haben die Bewertungen für ihre Aufgaben und Tests während Ihres Vorgangs eingegeben.</span><span class="sxs-lookup"><span data-stu-id="458bd-105">You've been entering the scores for their assignments and tests as you go.</span></span> <span data-ttu-id="458bd-106">Jetzt ist es an der Zeit, die Schüler und Studenten zu ermitteln.</span><span class="sxs-lookup"><span data-stu-id="458bd-106">Now, it is time to determine the students' fates.</span></span>

<span data-ttu-id="458bd-107">Sie entwickeln ein Skript, das die Noten für jede Punktekategorie summiert.</span><span class="sxs-lookup"><span data-stu-id="458bd-107">You'll develop a script that totals the grades for each point category.</span></span> <span data-ttu-id="458bd-108">Anschließend wird jedem Kursteilnehmer basierend auf der Summe eine Notenstufe zugewiesen.</span><span class="sxs-lookup"><span data-stu-id="458bd-108">It will then assign a letter grade to each student based on the total.</span></span> <span data-ttu-id="458bd-109">Um die Genauigkeit sicherzustellen, fügen Sie ein paar Prüfungen hinzu, um festzustellen, ob einzelne Bewertungen zu niedrig oder hoch sind.</span><span class="sxs-lookup"><span data-stu-id="458bd-109">To help ensure accuracy, you'll add a couple checks to see if any individual scores are too low or high.</span></span> <span data-ttu-id="458bd-110">Wenn die Bewertung eines Schülers kleiner als Null oder mehr als der mögliche Punktwert ist, kennzeichnet das Skript die Zelle mit einer roten Füllung und nicht mit der Summe der Punkte dieses Schülers.</span><span class="sxs-lookup"><span data-stu-id="458bd-110">If a student's score is less than zero or more than the possible point value, the script will flag the cell with a red fill and not total that student's points.</span></span> <span data-ttu-id="458bd-111">Dies ist ein eindeutiger Hinweis darauf, welche Datensätze Sie überprüfen müssen.</span><span class="sxs-lookup"><span data-stu-id="458bd-111">This will be a clear indication of which records you need to double-check.</span></span> <span data-ttu-id="458bd-112">Außerdem fügen Sie den Noten einige grundlegende Formatierungen hinzu, damit Sie schnell die obere und untere Ebene der Klasse anzeigen können.</span><span class="sxs-lookup"><span data-stu-id="458bd-112">You'll also add some basic formatting to the grades so you can quickly view the top and bottom of the class.</span></span>

## <a name="scripting-skills-covered"></a><span data-ttu-id="458bd-113">Abgedeckte Skriptfähigkeiten</span><span class="sxs-lookup"><span data-stu-id="458bd-113">Scripting skills covered</span></span>

- <span data-ttu-id="458bd-114">Zellenformatierung</span><span class="sxs-lookup"><span data-stu-id="458bd-114">Cell formatting</span></span>
- <span data-ttu-id="458bd-115">Fehlerüberprüfung</span><span class="sxs-lookup"><span data-stu-id="458bd-115">Error checking</span></span>
- <span data-ttu-id="458bd-116">Reguläre Ausdrücke</span><span class="sxs-lookup"><span data-stu-id="458bd-116">Regular expressions</span></span>
- <span data-ttu-id="458bd-117">Bedingte Formatierung</span><span class="sxs-lookup"><span data-stu-id="458bd-117">Conditional formatting</span></span>

## <a name="setup-instructions"></a><span data-ttu-id="458bd-118">Setup-Anweisungen</span><span class="sxs-lookup"><span data-stu-id="458bd-118">Setup instructions</span></span>

1. <span data-ttu-id="458bd-119">Laden Sie <a href="grade-calculator.xlsx">grade-calculator.xlsx</a> auf Ihre OneDrive herunter.</span><span class="sxs-lookup"><span data-stu-id="458bd-119">Download <a href="grade-calculator.xlsx">grade-calculator.xlsx</a> to your OneDrive.</span></span>

1. <span data-ttu-id="458bd-120">Öffnen Sie die Arbeitsmappe mit Excel für das Web.</span><span class="sxs-lookup"><span data-stu-id="458bd-120">Open the workbook with Excel for the web.</span></span>

1. <span data-ttu-id="458bd-121">Wählen Sie auf der Registerkarte **"Automatisieren"** die Option **"Neues Skript"** aus, und fügen Sie das folgende Skript in den Editor ein.</span><span class="sxs-lookup"><span data-stu-id="458bd-121">Under the **Automate** tab, select **New Script** and paste the following script into the editor.</span></span>

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

1. <span data-ttu-id="458bd-122">Benennen Sie das Skript in den **Notenrechner um,** und speichern Sie es.</span><span class="sxs-lookup"><span data-stu-id="458bd-122">Rename the script to **Grade Calculator** and save it.</span></span>

## <a name="running-the-script"></a><span data-ttu-id="458bd-123">Ausführen des Skripts</span><span class="sxs-lookup"><span data-stu-id="458bd-123">Running the script</span></span>

<span data-ttu-id="458bd-124">Führen Sie das Skript **"Notenrechner"** auf dem einzigen Arbeitsblatt aus.</span><span class="sxs-lookup"><span data-stu-id="458bd-124">Run the **Grade Calculator** script on the only worksheet.</span></span> <span data-ttu-id="458bd-125">Das Skript führt die Noten aus und weist jedem Schüler eine Buchstabennote zu.</span><span class="sxs-lookup"><span data-stu-id="458bd-125">The script will total the grades and assign each student a letter grade.</span></span> <span data-ttu-id="458bd-126">Wenn einzelne Noten mehr Punkte haben, als die Aufgabe oder der Test wert ist, wird die ungültige Noten als rot markiert, und die Summe wird nicht berechnet.</span><span class="sxs-lookup"><span data-stu-id="458bd-126">If any individual grades have more points than the assignment or test is worth, then the offending grade is marked red and the total is not calculated.</span></span> <span data-ttu-id="458bd-127">Außerdem werden alle "A"-Noten grün hervorgehoben, während die Noten "D" und "F" gelb hervorgehoben sind.</span><span class="sxs-lookup"><span data-stu-id="458bd-127">Also, any 'A' grades are highlighted in green, while 'D' and 'F' grades are highlighted in yellow.</span></span>

### <a name="before-running-the-script"></a><span data-ttu-id="458bd-128">Vor dem Ausführen des Skripts</span><span class="sxs-lookup"><span data-stu-id="458bd-128">Before running the script</span></span>

:::image type="content" source="../../images/scenario-grade-calculator-before.png" alt-text="Ein Arbeitsblatt mit Zeilen mit Bewertungen für Schüler/Studenten.":::

### <a name="after-running-the-script"></a><span data-ttu-id="458bd-130">Nach dem Ausführen des Skripts</span><span class="sxs-lookup"><span data-stu-id="458bd-130">After running the script</span></span>

:::image type="content" source="../../images/scenario-grade-calculator-after.png" alt-text="Ein Arbeitsblatt, das die Schülerbewertungsdaten mit ungültigen Zellen in roten Summen für gültige Schülerzeilen anzeigt.":::
