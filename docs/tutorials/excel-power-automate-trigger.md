---
title: Übergeben von Daten zu Skripts in einem automatisch ausgeführten Power Automate-Datenfluss
description: Ein Lernprogramm zum Ausführen von Office-Skripts für Excel im Web mithilfe von Power Automate, wenn E-Mails empfangen und Flussdaten an das Skript übergeben werden.
ms.date: 11/30/2020
localization_priority: Priority
ms.openlocfilehash: b73f40c70669fedbe8a0adcf346995cb20b62d37
ms.sourcegitcommit: af487756dffea0f8f0cd62710c586842cb08073c
ms.translationtype: HT
ms.contentlocale: de-DE
ms.lasthandoff: 12/04/2020
ms.locfileid: "49571479"
---
# <a name="pass-data-to-scripts-in-an-automatically-run-power-automate-flow-preview"></a><span data-ttu-id="21e51-103">Übergeben von Daten zu Skripts in einem automatisch ausgeführten Power Automate-Datenfluss (Vorschau)</span><span class="sxs-lookup"><span data-stu-id="21e51-103">Pass data to scripts in an automatically-run Power Automate flow (preview)</span></span>

<span data-ttu-id="21e51-104">In diesem Lernprogramm erfahren Sie, wie Sie ein Office-Skript für Excel im Web mit einem automatisierten [Power Automate](https://flow.microsoft.com)-Workflow verwenden.</span><span class="sxs-lookup"><span data-stu-id="21e51-104">This tutorial teaches you how to use an Office Script for Excel on the web with an automated [Power Automate](https://flow.microsoft.com) workflow.</span></span> <span data-ttu-id="21e51-105">Das Skript wird jedes Mal, wenn Sie eine E-Mail erhalten, automatisch ausgeführt, um Informationen aus der E-Mail in einer Excel-Arbeitsmappe aufzuzeichnen.</span><span class="sxs-lookup"><span data-stu-id="21e51-105">Your script will automatically run each time you receive an email, recording information from the email in an Excel workbook.</span></span> <span data-ttu-id="21e51-106">Die Möglichkeit, Daten aus anderen Anwendungen in ein Office-Skript zu übertragen, bietet Ihnen ein hohes Maß an Flexibilität und Freiheit für Ihre automatisierten Prozesse.</span><span class="sxs-lookup"><span data-stu-id="21e51-106">Being able to pass data from other applications into an Office Script gives you a great deal of flexibility and freedom in your automated processes.</span></span>

> [!TIP]
> <span data-ttu-id="21e51-107">Wenn Sie mit Office-Skripten noch nicht vertraut sind, empfehlen wir, mit dem [Aufzeichnen, Bearbeiten und Erstellen von Office-Skripten in Excel im Web](excel-tutorial.md)-Lernprogramm zu beginnen.</span><span class="sxs-lookup"><span data-stu-id="21e51-107">If you are new to Office Scripts, we recommend starting with the [Record, edit, and create Office Scripts in Excel on the web](excel-tutorial.md) tutorial.</span></span> <span data-ttu-id="21e51-108">Wenn Sie noch nicht mit Power Automate vertraut sind, empfehlen wir, dass Sie mit dem Lernprogramm [Aufrufen von Skripts aus einem manuellen Power Automate-Datenfluss](excel-power-automate-manual.md) beginnen.</span><span class="sxs-lookup"><span data-stu-id="21e51-108">If you are new to Power Automate, we recommend starting with the [Call scripts from a manual Power Automate flow](excel-power-automate-manual.md) tutorial.</span></span> <span data-ttu-id="21e51-109">[Office-Skripts verwenden TypeScript](../overview/code-editor-environment.md), und dieses Lernprogramm richten sich an Anfänger bis Fortgeschrittene mit JavaScript oder TypeScript.</span><span class="sxs-lookup"><span data-stu-id="21e51-109">[Office Scripts use TypeScript](../overview/code-editor-environment.md) and this tutorial is intended for people with beginner to intermediate-level knowledge of JavaScript or TypeScript.</span></span> <span data-ttu-id="21e51-110">Wenn Sie noch nicht mit JavaScript vertraut sind, empfehlen wir Ihnen, mit dem [Mozilla-JavaScript-Lernprogramm](https://developer.mozilla.org/docs/Web/JavaScript/Guide/Introduction) zu beginnen.</span><span class="sxs-lookup"><span data-stu-id="21e51-110">If you're new to JavaScript, we recommend starting with the [Mozilla JavaScript tutorial](https://developer.mozilla.org/docs/Web/JavaScript/Guide/Introduction).</span></span>

## <a name="prerequisites"></a><span data-ttu-id="21e51-111">Voraussetzungen</span><span class="sxs-lookup"><span data-stu-id="21e51-111">Prerequisites</span></span>

[!INCLUDE [Tutorial prerequisites](../includes/power-automate-tutorial-prerequisites.md)]

## <a name="prepare-the-workbook"></a><span data-ttu-id="21e51-112">Vorbereiten der Arbeitsmappe</span><span class="sxs-lookup"><span data-stu-id="21e51-112">Prepare the workbook</span></span>

<span data-ttu-id="21e51-113">Power Automation kann [relative Bezüge](../develop/power-automate-integration.md#avoid-using-relative-references) wie `Workbook.getActiveWorksheet` nicht verwenden, um auf Arbeitsmappenfunktionen zuzugreifen.</span><span class="sxs-lookup"><span data-stu-id="21e51-113">Power Automate can't use [relative references](../develop/power-automate-integration.md#avoid-using-relative-references) like `Workbook.getActiveWorksheet` to access workbook components.</span></span> <span data-ttu-id="21e51-114">Deshalb benötigen wir eine Arbeitsmappe und ein Arbeitsblatt mit konsistenten Namen, die Power Automate als Referenz verwenden kann.</span><span class="sxs-lookup"><span data-stu-id="21e51-114">So, we need a workbook and worksheet with consistent names for Power Automate to reference.</span></span>

1. <span data-ttu-id="21e51-115">Erstellen Sie eine neue Arbeitsmappe mit dem Namen **Mein Arbeitsblatt**.</span><span class="sxs-lookup"><span data-stu-id="21e51-115">Create a new workbook named **MyWorkbook**.</span></span>

2. <span data-ttu-id="21e51-116">Wechseln Sie zur Registerkarte **Automate**, und wählen Sie **Code Editor** aus.</span><span class="sxs-lookup"><span data-stu-id="21e51-116">Go to the **Automate** tab and select **Code Editor**.</span></span>

3. <span data-ttu-id="21e51-117">Wählen Sie **New Script** aus.</span><span class="sxs-lookup"><span data-stu-id="21e51-117">Select **New Script**.</span></span>

4. <span data-ttu-id="21e51-118">Ersetzen Sie den vorhandenen Code durch den folgenden Code, und klicken Sie auf **Run**.</span><span class="sxs-lookup"><span data-stu-id="21e51-118">Replace the existing code with the following script and press **Run**.</span></span> <span data-ttu-id="21e51-119">Dadurch wird die Arbeitsmappe mit konsistenten Namen für Arbeitsblatt, Tabelle und PivotTable eingerichtet.</span><span class="sxs-lookup"><span data-stu-id="21e51-119">This will setup the workbook with consistent worksheet, table, and PivotTable names.</span></span>

    ```TypeScript
    function main(workbook: ExcelScript.Workbook) {
      // Add a new worksheet to store our email table
      let emailsSheet = workbook.addWorksheet("Emails");

      // Add data and create a table
      emailsSheet.getRange("A1:D1").setValues([
        ["Date", "Day of the week", "Email address", "Subject"]
      ]);
      let newTable = workbook.addTable(emailsSheet.getRange("A1:D2"), true);
      newTable.setName("EmailTable");

      // Add a new PivotTable to a new worksheet
      let pivotWorksheet = workbook.addWorksheet("Subjects");
      let newPivotTable = workbook.addPivotTable("Pivot", "EmailTable", pivotWorksheet.getRange("A3:C20"));

      // Setup the pivot hierarchies
      newPivotTable.addRowHierarchy(newPivotTable.getHierarchy("Day of the week"));
      newPivotTable.addRowHierarchy(newPivotTable.getHierarchy("Email address"));
      newPivotTable.addDataHierarchy(newPivotTable.getHierarchy("Subject"));
    }
    ```

## <a name="create-an-office-script"></a><span data-ttu-id="21e51-120">Erstellen eines Office-Skripts</span><span class="sxs-lookup"><span data-stu-id="21e51-120">Create an Office Script</span></span>

<span data-ttu-id="21e51-121">Jetzt erstellen Sie ein Skript, das Informationen aus einer E-Mail protokolliert.</span><span class="sxs-lookup"><span data-stu-id="21e51-121">Let's create a script that logs information from an email.</span></span> <span data-ttu-id="21e51-122">Wir möchten wissen, an welchen Wochentagen die meisten E-Mails empfangen werden und wie viele eindeutige Absender diese E-Mails senden.</span><span class="sxs-lookup"><span data-stu-id="21e51-122">We want to know how which days of the week we receive the most mail and how many unique senders are sending that mail.</span></span> <span data-ttu-id="21e51-123">Die Arbeitsmappe enthält eine Tabelle mit den Spalten **Date**, **Day of the week**, **Email address** und **Subject**.</span><span class="sxs-lookup"><span data-stu-id="21e51-123">Our workbook has a table with **Date**, **Day of the week**, **Email address**, and **Subject** columns.</span></span> <span data-ttu-id="21e51-124">Unser Arbeitsblatt enthält außerdem eine PivotTable, die die Spalten **Day of the week** und **Email address** verwendet (dabei handelt es sich um die Zeilenhierarchien).</span><span class="sxs-lookup"><span data-stu-id="21e51-124">Our worksheet also has a PivotTable that is pivoting on the **Day of the week** and **Email address** (those are the row hierarchies).</span></span> <span data-ttu-id="21e51-125">Die Anzahl eindeutiger **Themen** entspricht den aggregierten Informationen, die angezeigt werden (die Datenhierarchie).</span><span class="sxs-lookup"><span data-stu-id="21e51-125">The count of unique **Subjects** is the aggregated information being displayed (the data hierarchy).</span></span> <span data-ttu-id="21e51-126">Nachdem die E-Mail-Tabelle aktualisiert wurde, wird auch das Skript aktualisiert.</span><span class="sxs-lookup"><span data-stu-id="21e51-126">We'll have our script refresh that PivotTable after updating the email table.</span></span>

1. <span data-ttu-id="21e51-127">Wählen Sie im **Code-Editor** die Option **New Script** aus.</span><span class="sxs-lookup"><span data-stu-id="21e51-127">From within the **Code Editor**, select **New Script**.</span></span>

2. <span data-ttu-id="21e51-128">Der Datenstrom, den wir später im Lernprogramm erstellen, sendet die Skriptinformationen zu jeder empfangenen E-Mail-Nachricht.</span><span class="sxs-lookup"><span data-stu-id="21e51-128">The flow that we'll create later in the tutorial will send our script information about each email that's received.</span></span> <span data-ttu-id="21e51-129">Das Skript muss diese Eingabe über Parameter in der `main`-Funktion akzeptieren.</span><span class="sxs-lookup"><span data-stu-id="21e51-129">The script needs to accept that input through parameters in the `main` function.</span></span> <span data-ttu-id="21e51-130">Ersetzen Sie das Standardskript durch das folgende Skript:</span><span class="sxs-lookup"><span data-stu-id="21e51-130">Replace the default script with the following script:</span></span>

    ```TypeScript
    function main(
      workbook: ExcelScript.Workbook,
      from: string,
      dateReceived: string,
      subject: string) {

    }
    ```

3. <span data-ttu-id="21e51-131">Das Skript benötigt Zugriff auf die Tabelle und die PivotTable der Arbeitsmappe.</span><span class="sxs-lookup"><span data-stu-id="21e51-131">The script needs access to the workbook's table and PivotTable.</span></span> <span data-ttu-id="21e51-132">Fügen Sie den folgenden Code dem Skripttext nach dem öffnenden `{` hinzu:</span><span class="sxs-lookup"><span data-stu-id="21e51-132">Add the following code to the body of the script, after the opening `{`:</span></span>

    ```TypeScript
    // Get the email table.
    let emailWorksheet = workbook.getWorksheet("Emails");
    let table = emailWorksheet.getTable("EmailTable");
  
    // Get the PivotTable.
    let pivotTableWorksheet = workbook.getWorksheet("Subjects");
    let pivotTable = pivotTableWorksheet.getPivotTable("Pivot");
    ```

4. <span data-ttu-id="21e51-133">Der `dateReceived`-Parameter ist vom Typ `string`.</span><span class="sxs-lookup"><span data-stu-id="21e51-133">The `dateReceived` parameter is of type `string`.</span></span> <span data-ttu-id="21e51-134">Dies wird jetzt ein [`Date`-Objekt](../develop/javascript-objects.md#date) konvertiert, damit wir den Wochentag ganz einfach abrufen können.</span><span class="sxs-lookup"><span data-stu-id="21e51-134">Let's convert that to a [`Date` object](../develop/javascript-objects.md#date) so we can easily get the day of the week.</span></span> <span data-ttu-id="21e51-135">Danach muss der Zahlenwert des Tages einer besser lesbaren Version zugeordnet werden.</span><span class="sxs-lookup"><span data-stu-id="21e51-135">After doing that, we'll need to map the day's number value to a more readable version.</span></span> <span data-ttu-id="21e51-136">Fügen Sie den folgenden Code am Ende des Skripts vor der schließenden `}` hinzu:</span><span class="sxs-lookup"><span data-stu-id="21e51-136">Add the following code to the end of your script, before the closing `}`:</span></span>

    ```TypeScript
      // Parse the received date string to determine the day of the week.
      let emailDate = new Date(dateReceived);
      let dayName = emailDate.toLocaleDateString("en-US", { weekday: 'long' });
    ```

5. <span data-ttu-id="21e51-137">Die `subject`-Zeichenfolge enthält möglicherweise das „RE:“-Antworttag.</span><span class="sxs-lookup"><span data-stu-id="21e51-137">The `subject` string may include the "RE:" reply tag.</span></span> <span data-ttu-id="21e51-138">Wir entfernen den Tag aus der Zeichenfolge, damit E-Mails im selben Thread den gleichen Betreff für die Tabelle aufweisen.</span><span class="sxs-lookup"><span data-stu-id="21e51-138">Let's remove that from the string so that emails in the same thread have the same subject for the table.</span></span> <span data-ttu-id="21e51-139">Fügen Sie den folgenden Code am Ende des Skripts vor der schließenden `}` hinzu:</span><span class="sxs-lookup"><span data-stu-id="21e51-139">Add the following code to the end of your script, before the closing `}`:</span></span>

    ```TypeScript
    // Remove the reply tag from the email subject to group emails on the same thread.
    let subjectText = subject.replace("Re: ", "");
    subjectText = subjectText.replace("RE: ", "");
    ```

6. <span data-ttu-id="21e51-140">Nachdem Sie die E-Mail-Daten nach Wunsch formatiert haben, fügen Sie eine Zeile zur E-Mail-Tabelle hinzu.</span><span class="sxs-lookup"><span data-stu-id="21e51-140">Now that the email data has been formatted to our liking, let's add a row to the email table.</span></span> <span data-ttu-id="21e51-141">Fügen Sie den folgenden Code am Ende des Skripts vor der schließenden `}` hinzu:</span><span class="sxs-lookup"><span data-stu-id="21e51-141">Add the following code to the end of your script, before the closing `}`:</span></span>

    ```TypeScript
    // Add the parsed text to the table.
    table.addRow(-1, [dateReceived, dayName, from, subjectText]);
    ```

7. <span data-ttu-id="21e51-142">Abschließend stellen Sie sicher, dass die PivotTable aktualisiert wird.</span><span class="sxs-lookup"><span data-stu-id="21e51-142">Finally, let's make sure the PivotTable is refreshed.</span></span> <span data-ttu-id="21e51-143">Fügen Sie den folgenden Code am Ende des Skripts vor der schließenden `}` hinzu:</span><span class="sxs-lookup"><span data-stu-id="21e51-143">Add the following code to the end of your script, before the closing `}`:</span></span>

    ```TypeScript
    // Refresh the PivotTable to include the new row.
    pivotTable.refresh();
    ```

8. <span data-ttu-id="21e51-144">Benennen Sie das Skript in **E-Mail aufzeichnen** um, und klicken Sie auf **Save Script**.</span><span class="sxs-lookup"><span data-stu-id="21e51-144">Rename your script **Record Email** and press **Save script**.</span></span>

<span data-ttu-id="21e51-145">Jetzt ist Ihr Skript bereit für einen Power Automate-Workflow.</span><span class="sxs-lookup"><span data-stu-id="21e51-145">Your script is now ready for a Power Automate workflow.</span></span> <span data-ttu-id="21e51-146">Es sollte wie das folgende Skript aussehen:</span><span class="sxs-lookup"><span data-stu-id="21e51-146">It should look like the following script:</span></span>

```TypeScript
function main(
  workbook: ExcelScript.Workbook,
  from: string,
  dateReceived: string,
  subject: string) {
  // Get the email table.
  let emailWorksheet = workbook.getWorksheet("Emails");
  let table = emailWorksheet.getTable("EmailTable");

  // Get the PivotTable.
  let pivotTableWorksheet = workbook.getWorksheet("Subjects");
  let pivotTable = pivotTableWorksheet.getPivotTable("Pivot");

  // Parse the received date string to determine the day of the week.
  let emailDate = new Date(dateReceived);
  let dayName = emailDate.toLocaleDateString("en-US", { weekday: 'long' });

  // Remove the reply tag from the email subject to group emails on the same thread.
  let subjectText = subject.replace("Re: ", "");
  subjectText = subjectText.replace("RE: ", "");

  // Add the parsed text to the table.
  table.addRow(-1, [dateReceived, dayName, from, subjectText]);

  // Refresh the PivotTable to include the new row.
  pivotTable.refresh();
}
```

## <a name="create-an-automated-workflow-with-power-automate"></a><span data-ttu-id="21e51-147">Erstellen eines automatisierten Workflows mit Power Automate</span><span class="sxs-lookup"><span data-stu-id="21e51-147">Create an automated workflow with Power Automate</span></span>

1. <span data-ttu-id="21e51-148">Melden Sie sich an der [Power Automate-Website](https://flow.microsoft.com) an.</span><span class="sxs-lookup"><span data-stu-id="21e51-148">Sign in to the [Power Automate site](https://flow.microsoft.com).</span></span>

2. <span data-ttu-id="21e51-149">Klicken Sie in dem Menü, das auf der linken Seite des Bildschirms angezeigt wird, auf **Create**.</span><span class="sxs-lookup"><span data-stu-id="21e51-149">In the menu that's displayed on the left side of the screen, press **Create**.</span></span> <span data-ttu-id="21e51-150">Damit gelangen Sie zur Liste der Möglichkeiten zum Erstellen neuer Workflows.</span><span class="sxs-lookup"><span data-stu-id="21e51-150">This brings you to list of ways to create new workflows.</span></span>

    ![Die Schaltfläche „Erstellen“ in Power Automate.](../images/power-automate-tutorial-1.png)

3. <span data-ttu-id="21e51-152">Wählen Sie im Abschnitt **Start from blank** die Option **Automated flow** aus.</span><span class="sxs-lookup"><span data-stu-id="21e51-152">In the **Start from blank** section, select **Automated flow**.</span></span> <span data-ttu-id="21e51-153">Dadurch wird ein Workflow erstellt, der von einem Ereignis ausgelöst wird, z. B. das Empfangen einer E-Mail.</span><span class="sxs-lookup"><span data-stu-id="21e51-153">This creates a workflow triggered by an event, such as receiving an email.</span></span>

    ![Die Option für dem automatisiertem Fluss in Power Automate.](../images/power-automate-params-tutorial-1.png)

4. <span data-ttu-id="21e51-155">Geben Sie im daraufhin angezeigten Dialogfenster einen Namen für den Fluss im Textfeld **Flow name** ein.</span><span class="sxs-lookup"><span data-stu-id="21e51-155">In the dialog window that appears, enter a name for your flow in the **Flow name** text box.</span></span> <span data-ttu-id="21e51-156">Wählen Sie dann **When a new email arrives** aus der Liste der Optionen unter **Choose your flow's trigger** aus.</span><span class="sxs-lookup"><span data-stu-id="21e51-156">Then select **When a new email arrives** from the list of options under **Choose your flow's trigger**.</span></span> <span data-ttu-id="21e51-157">Möglicherweise müssen Sie mithilfe des Suchfelds nach der Option suchen.</span><span class="sxs-lookup"><span data-stu-id="21e51-157">You may need to search for the option using the search box.</span></span> <span data-ttu-id="21e51-158">Klicken Sie abschließend **Create**.</span><span class="sxs-lookup"><span data-stu-id="21e51-158">Finally, press **Create**.</span></span>

    ![Ein Teil des Fensters zum Erstellen eines automatisierten Flusses in Power Automate, das die Option „Neue E-Mail trifft ein“ zeigt.](../images/power-automate-params-tutorial-2.png)

    > [!NOTE]
    > <span data-ttu-id="21e51-160">In diesem Lernprogramm wird Outlook verwendet.</span><span class="sxs-lookup"><span data-stu-id="21e51-160">This tutorial uses Outlook.</span></span> <span data-ttu-id="21e51-161">Sie können stattdessen Ihren bevorzugten E-Mail-Dienst verwenden, obwohl einige Optionen unterschiedlich sein können.</span><span class="sxs-lookup"><span data-stu-id="21e51-161">Feel free to use your preferred email service instead, though some options may be different.</span></span>

5. <span data-ttu-id="21e51-162">Klicken Sie auf **New step**.</span><span class="sxs-lookup"><span data-stu-id="21e51-162">Press **New step**.</span></span>

6. <span data-ttu-id="21e51-163">Wählen Sie die Registerkarte **Standard** aus, und wählen Sie dann **Excel Online (Business)** aus.</span><span class="sxs-lookup"><span data-stu-id="21e51-163">Select the **Standard** tab, then select **Excel Online (Business)**.</span></span>

    ![Die Power Automate-Option für Excel Online (Business).](../images/power-automate-tutorial-4.png)

7. <span data-ttu-id="21e51-165">Wählen Sie unter **Actions** die Option **Run script (preview)** aus.</span><span class="sxs-lookup"><span data-stu-id="21e51-165">Under **Actions**, select **Run script (preview)**.</span></span>

    ![Die Power Automate-Aktionsoption für „Run script (preview)“.](../images/power-automate-tutorial-5.png)

8. <span data-ttu-id="21e51-167">Als Nächstes wählen Sie die Arbeitsmappe, das Skript und die Eingabeargumente für das Skript aus, die im Datenfluss-Schritt verwendet werden sollen.</span><span class="sxs-lookup"><span data-stu-id="21e51-167">Next, you'll select the workbook, script, and script input arguments to use in the flow step.</span></span> <span data-ttu-id="21e51-168">In diesem Lernprogramm verwenden Sie die Arbeitsmappe, die Sie in Ihrem OneDrive erstellt haben. Sie könnten jedoch jede beliebige Arbeitsmappe auf einer OneDrive- oder SharePoint-Website verwenden.</span><span class="sxs-lookup"><span data-stu-id="21e51-168">For the tutorial, you'll use the workbook you created in your OneDrive, but you could use any workbook in a OneDrive or SharePoint site.</span></span> <span data-ttu-id="21e51-169">Geben Sie die folgenden Einstellungen für den Konnektor **Run script** an:</span><span class="sxs-lookup"><span data-stu-id="21e51-169">Specify the following settings for the **Run script** connector:</span></span>

    - <span data-ttu-id="21e51-170">**Location**: OneDrive for Business</span><span class="sxs-lookup"><span data-stu-id="21e51-170">**Location**: OneDrive for Business</span></span>
    - <span data-ttu-id="21e51-171">**Document Library**: OneDrive</span><span class="sxs-lookup"><span data-stu-id="21e51-171">**Document Library**: OneDrive</span></span>
    - <span data-ttu-id="21e51-172">**File**: MyWorkbook.xlsx *(Ausgewählt über den Dateibrowser)*</span><span class="sxs-lookup"><span data-stu-id="21e51-172">**File**: MyWorkbook.xlsx *(Chosen through the file browser)*</span></span>
    - <span data-ttu-id="21e51-173">**Script**: E-Mail aufzeichnen</span><span class="sxs-lookup"><span data-stu-id="21e51-173">**Script**: Record Email</span></span>
    - <span data-ttu-id="21e51-174">**from**: Von *(dynamischer Inhalt aus Outlook)*</span><span class="sxs-lookup"><span data-stu-id="21e51-174">**from**: From *(dynamic content from Outlook)*</span></span>
    - <span data-ttu-id="21e51-175">**dateReceived**: Uhrzeit des Empfangs *(dynamischer Inhalt aus Outlook)*</span><span class="sxs-lookup"><span data-stu-id="21e51-175">**dateReceived**: Received Time *(dynamic content from Outlook)*</span></span>
    - <span data-ttu-id="21e51-176">**subject**: Betreff *(dynamischer Inhalt aus Outlook)*</span><span class="sxs-lookup"><span data-stu-id="21e51-176">**subject**: Subject *(dynamic content from Outlook)*</span></span>

    <span data-ttu-id="21e51-177">*Beachten Sie, dass die Parameter für das Skript nur angezeigt werden, wenn das Skript ausgewählt wurde.*</span><span class="sxs-lookup"><span data-stu-id="21e51-177">*Note that the parameters for the script will only appear once the script is selected.*</span></span>

    ![Parameter der Aktionsoption "Power Automation" für "Skript ausführen" (Vorschau).](../images/power-automate-params-tutorial-3.png)

9. <span data-ttu-id="21e51-179">Klicken Sie auf **Save**.</span><span class="sxs-lookup"><span data-stu-id="21e51-179">Press **Save**.</span></span>

<span data-ttu-id="21e51-180">Der Fluss ist nun aktiviert.</span><span class="sxs-lookup"><span data-stu-id="21e51-180">Your flow is now enabled.</span></span> <span data-ttu-id="21e51-181">Er wird das Skript automatisch jedes Mal ausführen, wenn Sie eine E-Mail über Outlook erhalten.</span><span class="sxs-lookup"><span data-stu-id="21e51-181">It will automatically run your script each time you receive an email through Outlook.</span></span>

## <a name="manage-the-script-in-power-automate"></a><span data-ttu-id="21e51-182">Verwalten des Skripts in Power Automate</span><span class="sxs-lookup"><span data-stu-id="21e51-182">Manage the script in Power Automate</span></span>

1. <span data-ttu-id="21e51-183">Wählen Sie auf der Hauptseite der Power Automate-Seite **My Flows** aus.</span><span class="sxs-lookup"><span data-stu-id="21e51-183">From the main Power Automate page, select **My flows**.</span></span>

    ![Die Schaltfläche „My Flows“ in Power Automate.](../images/power-automate-tutorial-7.png)

2. <span data-ttu-id="21e51-185">Wählen Sie Ihren Flow aus.</span><span class="sxs-lookup"><span data-stu-id="21e51-185">Select your flow.</span></span> <span data-ttu-id="21e51-186">Hier sehen Sie den Ausführungsverlauf.</span><span class="sxs-lookup"><span data-stu-id="21e51-186">Here you can see the run history.</span></span> <span data-ttu-id="21e51-187">Sie können die Seite aktualisieren, oder Sie können auf die Schaltfläche **All runs** klicken, um den Verlauf zu aktualisieren.</span><span class="sxs-lookup"><span data-stu-id="21e51-187">You can refresh the page or press the refresh **All runs** button to update the history.</span></span> <span data-ttu-id="21e51-188">Der Flow wird kurz nach Empfang einer E-Mail ausgelöst.</span><span class="sxs-lookup"><span data-stu-id="21e51-188">The flow will trigger shortly after an email is received.</span></span> <span data-ttu-id="21e51-189">Testen Sie den Flow durch Senden von E-Mails.</span><span class="sxs-lookup"><span data-stu-id="21e51-189">Test the flow by sending yourself mail.</span></span>

<span data-ttu-id="21e51-190">Wenn der Flow ausgelöst und das Skript erfolgreich ausgeführt wird, sollten die Tabelle und die PivotTable der Arbeitsmappe aktualisiert werden.</span><span class="sxs-lookup"><span data-stu-id="21e51-190">When the flow is triggered and successfully runs your script, you should see the workbook's table and PivotTable update.</span></span>

![Die E-Mail-Tabelle nach dem Flow wurde mehrere Male ausgeführt.](../images/power-automate-params-tutorial-4.png)

![Die PivotTable nach dem Flow wurde mehrere Male ausgeführt.](../images/power-automate-params-tutorial-5.png)

## <a name="next-steps"></a><span data-ttu-id="21e51-193">Nächste Schritte</span><span class="sxs-lookup"><span data-stu-id="21e51-193">Next steps</span></span>

<span data-ttu-id="21e51-194">Besuchen Sie [Ausführen von Office-Skripts mit Power Automate](../develop/power-automate-integration.md), um mehr über das Verbinden von Office-Skripts mit Power Automate zu erfahren.</span><span class="sxs-lookup"><span data-stu-id="21e51-194">Visit [Run Office Scripts with Power Automate](../develop/power-automate-integration.md) to learn more about connecting Office Scripts with Power Automate.</span></span>

<span data-ttu-id="21e51-195">Sie können sich auch das Beispielszenario [Automatisierte Aufgabenerinnerungen](../resources/scenarios/task-reminders.md) ansehen, um zu erfahren, wie Sie Office-Skripts und Power Automate mit Teams Adaptive Cards kombinieren können.</span><span class="sxs-lookup"><span data-stu-id="21e51-195">You can also check out the [Automated task reminders sample scenario](../resources/scenarios/task-reminders.md) to learn how to combine Office Scripts and Power Automate with Teams Adaptive Cards.</span></span>
