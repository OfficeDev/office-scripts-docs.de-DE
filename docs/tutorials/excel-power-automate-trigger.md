---
title: Übergeben von Daten zu Skripts in einem automatisch ausgeführten Power Automate-Datenfluss
description: Ein Lernprogramm zum Ausführen von Office-Skripts für Excel im Web mithilfe von Power Automate, wenn E-Mails empfangen und Flussdaten an das Skript übergeben werden.
ms.date: 06/29/2021
localization_priority: Priority
ms.openlocfilehash: 27a028d3cc2af58ca158bb631b7b266cd2a3d488
ms.sourcegitcommit: 211c157ca746e266eeb079f5fa1925a1e35ab702
ms.translationtype: HT
ms.contentlocale: de-DE
ms.lasthandoff: 07/07/2021
ms.locfileid: "53313701"
---
# <a name="pass-data-to-scripts-in-an-automatically-run-power-automate-flow"></a><span data-ttu-id="c1ef5-103">Übergeben von Daten zu Skripts in einem automatisch ausgeführten Power Automate-Datenfluss</span><span class="sxs-lookup"><span data-stu-id="c1ef5-103">Pass data to scripts in an automatically-run Power Automate flow</span></span>

<span data-ttu-id="c1ef5-104">In diesem Lernprogramm erfahren Sie, wie Sie ein Office-Skript für Excel im Web mit einem automatisierten [Power Automate](https://flow.microsoft.com)-Workflow verwenden.</span><span class="sxs-lookup"><span data-stu-id="c1ef5-104">This tutorial teaches you how to use an Office Script for Excel on the web with an automated [Power Automate](https://flow.microsoft.com) workflow.</span></span> <span data-ttu-id="c1ef5-105">Das Skript wird jedes Mal, wenn Sie eine E-Mail erhalten, automatisch ausgeführt, um Informationen aus der E-Mail in einer Excel-Arbeitsmappe aufzuzeichnen.</span><span class="sxs-lookup"><span data-stu-id="c1ef5-105">Your script will automatically run each time you receive an email, recording information from the email in an Excel workbook.</span></span> <span data-ttu-id="c1ef5-106">Die Möglichkeit, Daten aus anderen Anwendungen in ein Office-Skript zu übertragen, bietet Ihnen ein hohes Maß an Flexibilität und Freiheit für Ihre automatisierten Prozesse.</span><span class="sxs-lookup"><span data-stu-id="c1ef5-106">Being able to pass data from other applications into an Office Script gives you a great deal of flexibility and freedom in your automated processes.</span></span>

> [!TIP]
> <span data-ttu-id="c1ef5-107">Wenn Sie mit Office-Skripten noch nicht vertraut sind, empfehlen wir, mit dem [Aufzeichnen, Bearbeiten und Erstellen von Office-Skripten in Excel im Web](excel-tutorial.md)-Lernprogramm zu beginnen.</span><span class="sxs-lookup"><span data-stu-id="c1ef5-107">If you are new to Office Scripts, we recommend starting with the [Record, edit, and create Office Scripts in Excel on the web](excel-tutorial.md) tutorial.</span></span> <span data-ttu-id="c1ef5-108">Wenn Sie noch nicht mit Power Automate vertraut sind, empfehlen wir, dass Sie mit dem Lernprogramm [Aufrufen von Skripts aus einem manuellen Power Automate-Datenfluss](excel-power-automate-manual.md) beginnen.</span><span class="sxs-lookup"><span data-stu-id="c1ef5-108">If you are new to Power Automate, we recommend starting with the [Call scripts from a manual Power Automate flow](excel-power-automate-manual.md) tutorial.</span></span> <span data-ttu-id="c1ef5-109">[Office-Skripts verwenden TypeScript](../overview/code-editor-environment.md), und dieses Lernprogramm richten sich an Anfänger bis Fortgeschrittene mit JavaScript oder TypeScript.</span><span class="sxs-lookup"><span data-stu-id="c1ef5-109">[Office Scripts use TypeScript](../overview/code-editor-environment.md) and this tutorial is intended for people with beginner to intermediate-level knowledge of JavaScript or TypeScript.</span></span> <span data-ttu-id="c1ef5-110">Wenn Sie noch nicht mit JavaScript vertraut sind, empfehlen wir Ihnen, mit dem [Mozilla-JavaScript-Lernprogramm](https://developer.mozilla.org/docs/Web/JavaScript/Guide/Introduction) zu beginnen.</span><span class="sxs-lookup"><span data-stu-id="c1ef5-110">If you're new to JavaScript, we recommend starting with the [Mozilla JavaScript tutorial](https://developer.mozilla.org/docs/Web/JavaScript/Guide/Introduction).</span></span>

## <a name="prerequisites"></a><span data-ttu-id="c1ef5-111">Voraussetzungen</span><span class="sxs-lookup"><span data-stu-id="c1ef5-111">Prerequisites</span></span>

[!INCLUDE [Tutorial prerequisites](../includes/power-automate-tutorial-prerequisites.md)]

## <a name="prepare-the-workbook"></a><span data-ttu-id="c1ef5-112">Vorbereiten der Arbeitsmappe</span><span class="sxs-lookup"><span data-stu-id="c1ef5-112">Prepare the workbook</span></span>

<span data-ttu-id="c1ef5-113">Power Automate sollte für den Zugriff auf Arbeitsmappenkomponenten keine [relativen Bezüge](../testing/power-automate-troubleshooting.md#avoid-relative-references) wie `Workbook.getActiveWorksheet` verwenden.</span><span class="sxs-lookup"><span data-stu-id="c1ef5-113">Power Automate shouldn't use [relative references](../testing/power-automate-troubleshooting.md#avoid-relative-references) like `Workbook.getActiveWorksheet` to access workbook components.</span></span> <span data-ttu-id="c1ef5-114">Deshalb benötigen wir eine Arbeitsmappe und ein Arbeitsblatt mit konsistenten Namen, die Power Automate als Referenz verwenden kann.</span><span class="sxs-lookup"><span data-stu-id="c1ef5-114">So, we need a workbook and worksheet with consistent names for Power Automate to reference.</span></span>

1. <span data-ttu-id="c1ef5-115">Erstellen Sie eine neue Arbeitsmappe mit dem Namen **Mein Arbeitsblatt**.</span><span class="sxs-lookup"><span data-stu-id="c1ef5-115">Create a new workbook named **MyWorkbook**.</span></span>

2. <span data-ttu-id="c1ef5-116">Wechseln Sie zur Registerkarte **Automatisieren**, und wählen Sie **Alle Skripts** aus.</span><span class="sxs-lookup"><span data-stu-id="c1ef5-116">Go to the **Automate** tab and select **All Scripts**.</span></span>

3. <span data-ttu-id="c1ef5-117">Wählen Sie **New Script** aus.</span><span class="sxs-lookup"><span data-stu-id="c1ef5-117">Select **New Script**.</span></span>

4. <span data-ttu-id="c1ef5-118">Ersetzen Sie den vorhandenen Code durch den folgenden Code, und wählen Sie **Ausführen** aus.</span><span class="sxs-lookup"><span data-stu-id="c1ef5-118">Replace the existing code with the following script and select **Run**.</span></span> <span data-ttu-id="c1ef5-119">Dadurch wird die Arbeitsmappe mit konsistenten Namen für Arbeitsblatt, Tabelle und PivotTable eingerichtet.</span><span class="sxs-lookup"><span data-stu-id="c1ef5-119">This will setup the workbook with consistent worksheet, table, and PivotTable names.</span></span>

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

## <a name="create-an-office-script"></a><span data-ttu-id="c1ef5-120">Erstellen eines Office-Skripts</span><span class="sxs-lookup"><span data-stu-id="c1ef5-120">Create an Office Script</span></span>

<span data-ttu-id="c1ef5-121">Jetzt erstellen Sie ein Skript, das Informationen aus einer E-Mail protokolliert.</span><span class="sxs-lookup"><span data-stu-id="c1ef5-121">Let's create a script that logs information from an email.</span></span> <span data-ttu-id="c1ef5-122">Wir möchten wissen, an welchen Wochentagen wir die meisten E-Mails empfangen, und wie viele eindeutige Absender diese E-Mails senden.</span><span class="sxs-lookup"><span data-stu-id="c1ef5-122">We want to know which days of the week we receive the most mail and how many unique senders are sending that mail.</span></span> <span data-ttu-id="c1ef5-123">Die Arbeitsmappe enthält eine Tabelle mit den Spalten **Date**, **Day of the week**, **Email address** und **Subject**.</span><span class="sxs-lookup"><span data-stu-id="c1ef5-123">Our workbook has a table with **Date**, **Day of the week**, **Email address**, and **Subject** columns.</span></span> <span data-ttu-id="c1ef5-124">Unser Arbeitsblatt enthält außerdem eine PivotTable, die die Spalten **Day of the week** und **Email address** verwendet (dabei handelt es sich um die Zeilenhierarchien).</span><span class="sxs-lookup"><span data-stu-id="c1ef5-124">Our worksheet also has a PivotTable that is pivoting on the **Day of the week** and **Email address** (those are the row hierarchies).</span></span> <span data-ttu-id="c1ef5-125">Die Anzahl eindeutiger **Themen** entspricht den aggregierten Informationen, die angezeigt werden (die Datenhierarchie).</span><span class="sxs-lookup"><span data-stu-id="c1ef5-125">The count of unique **Subjects** is the aggregated information being displayed (the data hierarchy).</span></span> <span data-ttu-id="c1ef5-126">Nachdem die E-Mail-Tabelle aktualisiert wurde, wird auch das Skript aktualisiert.</span><span class="sxs-lookup"><span data-stu-id="c1ef5-126">We'll have our script refresh that PivotTable after updating the email table.</span></span>

1. <span data-ttu-id="c1ef5-127">Wählen Sie im Aufgabenbereich Code-Editor die Option **Neues Skript** aus.</span><span class="sxs-lookup"><span data-stu-id="c1ef5-127">From within the Code Editor task pane, select **New Script**.</span></span>

2. <span data-ttu-id="c1ef5-128">Der Datenstrom, den wir später im Lernprogramm erstellen, sendet die Skriptinformationen zu jeder empfangenen E-Mail-Nachricht.</span><span class="sxs-lookup"><span data-stu-id="c1ef5-128">The flow that we'll create later in the tutorial will send our script information about each email that's received.</span></span> <span data-ttu-id="c1ef5-129">Das Skript muss diese Eingabe über Parameter in der `main`-Funktion akzeptieren.</span><span class="sxs-lookup"><span data-stu-id="c1ef5-129">The script needs to accept that input through parameters in the `main` function.</span></span> <span data-ttu-id="c1ef5-130">Ersetzen Sie das Standardskript durch das folgende Skript:</span><span class="sxs-lookup"><span data-stu-id="c1ef5-130">Replace the default script with the following script:</span></span>

    ```TypeScript
    function main(
      workbook: ExcelScript.Workbook,
      from: string,
      dateReceived: string,
      subject: string) {

    }
    ```

3. <span data-ttu-id="c1ef5-131">Das Skript benötigt Zugriff auf die Tabelle und die PivotTable der Arbeitsmappe.</span><span class="sxs-lookup"><span data-stu-id="c1ef5-131">The script needs access to the workbook's table and PivotTable.</span></span> <span data-ttu-id="c1ef5-132">Fügen Sie den folgenden Code dem Skripttext nach dem öffnenden `{` hinzu:</span><span class="sxs-lookup"><span data-stu-id="c1ef5-132">Add the following code to the body of the script, after the opening `{`:</span></span>

    ```TypeScript
    // Get the email table.
    let emailWorksheet = workbook.getWorksheet("Emails");
    let table = emailWorksheet.getTable("EmailTable");
  
    // Get the PivotTable.
    let pivotTableWorksheet = workbook.getWorksheet("Subjects");
    let pivotTable = pivotTableWorksheet.getPivotTable("Pivot");
    ```

4. <span data-ttu-id="c1ef5-133">Der `dateReceived`-Parameter ist vom Typ `string`.</span><span class="sxs-lookup"><span data-stu-id="c1ef5-133">The `dateReceived` parameter is of type `string`.</span></span> <span data-ttu-id="c1ef5-134">Dies wird jetzt ein [`Date`-Objekt](../develop/javascript-objects.md#date) konvertiert, damit wir den Wochentag ganz einfach abrufen können.</span><span class="sxs-lookup"><span data-stu-id="c1ef5-134">Let's convert that to a [`Date` object](../develop/javascript-objects.md#date) so we can easily get the day of the week.</span></span> <span data-ttu-id="c1ef5-135">Danach muss der Zahlenwert des Tages einer besser lesbaren Version zugeordnet werden.</span><span class="sxs-lookup"><span data-stu-id="c1ef5-135">After doing that, we'll need to map the day's number value to a more readable version.</span></span> <span data-ttu-id="c1ef5-136">Fügen Sie den folgenden Code am Ende des Skripts vor der schließenden `}` hinzu:</span><span class="sxs-lookup"><span data-stu-id="c1ef5-136">Add the following code to the end of your script, before the closing `}`:</span></span>

    ```TypeScript
      // Parse the received date string to determine the day of the week.
      let emailDate = new Date(dateReceived);
      let dayName = emailDate.toLocaleDateString("en-US", { weekday: 'long' });
    ```

5. <span data-ttu-id="c1ef5-137">Die `subject`-Zeichenfolge enthält möglicherweise das „RE:“-Antworttag.</span><span class="sxs-lookup"><span data-stu-id="c1ef5-137">The `subject` string may include the "RE:" reply tag.</span></span> <span data-ttu-id="c1ef5-138">Wir entfernen den Tag aus der Zeichenfolge, damit E-Mails im selben Thread den gleichen Betreff für die Tabelle aufweisen.</span><span class="sxs-lookup"><span data-stu-id="c1ef5-138">Let's remove that from the string so that emails in the same thread have the same subject for the table.</span></span> <span data-ttu-id="c1ef5-139">Fügen Sie den folgenden Code am Ende des Skripts vor der schließenden `}` hinzu:</span><span class="sxs-lookup"><span data-stu-id="c1ef5-139">Add the following code to the end of your script, before the closing `}`:</span></span>

    ```TypeScript
    // Remove the reply tag from the email subject to group emails on the same thread.
    let subjectText = subject.replace("Re: ", "");
    subjectText = subjectText.replace("RE: ", "");
    ```

6. <span data-ttu-id="c1ef5-140">Nachdem Sie die E-Mail-Daten nach Wunsch formatiert haben, fügen Sie eine Zeile zur E-Mail-Tabelle hinzu.</span><span class="sxs-lookup"><span data-stu-id="c1ef5-140">Now that the email data has been formatted to our liking, let's add a row to the email table.</span></span> <span data-ttu-id="c1ef5-141">Fügen Sie den folgenden Code am Ende des Skripts vor der schließenden `}` hinzu:</span><span class="sxs-lookup"><span data-stu-id="c1ef5-141">Add the following code to the end of your script, before the closing `}`:</span></span>

    ```TypeScript
    // Add the parsed text to the table.
    table.addRow(-1, [dateReceived, dayName, from, subjectText]);
    ```

7. <span data-ttu-id="c1ef5-142">Abschließend stellen Sie sicher, dass die PivotTable aktualisiert wird.</span><span class="sxs-lookup"><span data-stu-id="c1ef5-142">Finally, let's make sure the PivotTable is refreshed.</span></span> <span data-ttu-id="c1ef5-143">Fügen Sie den folgenden Code am Ende des Skripts vor der schließenden `}` hinzu:</span><span class="sxs-lookup"><span data-stu-id="c1ef5-143">Add the following code to the end of your script, before the closing `}`:</span></span>

    ```TypeScript
    // Refresh the PivotTable to include the new row.
    pivotTable.refresh();
    ```

8. <span data-ttu-id="c1ef5-144">Benennen Sie das Skript in **E-Mail aufzeichnen** um, und wählen Sie **Skript speichern** aus.</span><span class="sxs-lookup"><span data-stu-id="c1ef5-144">Rename your script **Record Email** and select **Save script**.</span></span>

<span data-ttu-id="c1ef5-145">Jetzt ist Ihr Skript bereit für einen Power Automate-Workflow.</span><span class="sxs-lookup"><span data-stu-id="c1ef5-145">Your script is now ready for a Power Automate workflow.</span></span> <span data-ttu-id="c1ef5-146">Es sollte wie das folgende Skript aussehen:</span><span class="sxs-lookup"><span data-stu-id="c1ef5-146">It should look like the following script:</span></span>

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

## <a name="create-an-automated-workflow-with-power-automate"></a><span data-ttu-id="c1ef5-147">Erstellen eines automatisierten Workflows mit Power Automate</span><span class="sxs-lookup"><span data-stu-id="c1ef5-147">Create an automated workflow with Power Automate</span></span>

1. <span data-ttu-id="c1ef5-148">Melden Sie sich an der [Power Automate-Website](https://flow.microsoft.com) an.</span><span class="sxs-lookup"><span data-stu-id="c1ef5-148">Sign in to the [Power Automate site](https://flow.microsoft.com).</span></span>

2. <span data-ttu-id="c1ef5-149">Wählen Sie in dem Menü, das auf der linken Seite des Bildschirms angezeigt wird, **Erstellen** aus.</span><span class="sxs-lookup"><span data-stu-id="c1ef5-149">In the menu that's displayed on the left side of the screen, select **Create**.</span></span> <span data-ttu-id="c1ef5-150">Damit gelangen Sie zur Liste der Möglichkeiten zum Erstellen neuer Workflows.</span><span class="sxs-lookup"><span data-stu-id="c1ef5-150">This brings you to list of ways to create new workflows.</span></span>

    :::image type="content" source="../images/power-automate-tutorial-1.png" alt-text="Die Power Automate-Schaltfläche „Erstellen“.":::

3. <span data-ttu-id="c1ef5-152">Wählen Sie im Abschnitt **Start from blank** die Option **Automated flow** aus.</span><span class="sxs-lookup"><span data-stu-id="c1ef5-152">In the **Start from blank** section, select **Automated flow**.</span></span> <span data-ttu-id="c1ef5-153">Dadurch wird ein Workflow erstellt, der von einem Ereignis ausgelöst wird, z. B. das Empfangen einer E-Mail.</span><span class="sxs-lookup"><span data-stu-id="c1ef5-153">This creates a workflow triggered by an event, such as receiving an email.</span></span>

    :::image type="content" source="../images/power-automate-params-tutorial-1.png" alt-text="Die Option für den automatisierten Fluss in Power Automate.":::

4. <span data-ttu-id="c1ef5-155">Geben Sie im daraufhin angezeigten Dialogfenster einen Namen für den Fluss im Textfeld **Flow name** ein.</span><span class="sxs-lookup"><span data-stu-id="c1ef5-155">In the dialog window that appears, enter a name for your flow in the **Flow name** text box.</span></span> <span data-ttu-id="c1ef5-156">Wählen Sie dann **When a new email arrives** aus der Liste der Optionen unter **Choose your flow's trigger** aus.</span><span class="sxs-lookup"><span data-stu-id="c1ef5-156">Then select **When a new email arrives** from the list of options under **Choose your flow's trigger**.</span></span> <span data-ttu-id="c1ef5-157">Möglicherweise müssen Sie mithilfe des Suchfelds nach der Option suchen.</span><span class="sxs-lookup"><span data-stu-id="c1ef5-157">You may need to search for the option using the search box.</span></span> <span data-ttu-id="c1ef5-158">Wählen Sie abschließend **Erstellen** aus.</span><span class="sxs-lookup"><span data-stu-id="c1ef5-158">Finally, select **Create**.</span></span>

    :::image type="content" source="../images/power-automate-params-tutorial-2.png" alt-text="Ein Teil des Power Automate-Flusses zeigt den ‚Flussnamen‘ und die Optionen zum Auswählen des Triggers für den Fluss. Der Flussname ist ‚E-Mail-Fluss-Datensatz‘, und der Trigger ist die Option ‚Wenn eine neue E-Mail in Outlook eintrifft‘.":::

    > [!NOTE]
    > <span data-ttu-id="c1ef5-p116">In diesem Tutorial wird Outlook verwendet. Sie können stattdessen auch Ihren bevorzugten E-Mail-Dienst verwenden, obwohl einige Optionen anders sein können.</span><span class="sxs-lookup"><span data-stu-id="c1ef5-p116">This tutorial uses Outlook. Feel free to use your preferred email service instead, though some options may be different.</span></span>

5. <span data-ttu-id="c1ef5-162">Wählen Sie **Neuer Schritt** aus.</span><span class="sxs-lookup"><span data-stu-id="c1ef5-162">Select **New step**.</span></span>

6. <span data-ttu-id="c1ef5-163">Wählen Sie die Registerkarte **Standard** aus, und wählen Sie dann **Excel Online (Business)** aus.</span><span class="sxs-lookup"><span data-stu-id="c1ef5-163">Select the **Standard** tab, then select **Excel Online (Business)**.</span></span>

    :::image type="content" source="../images/power-automate-tutorial-4.png" alt-text="Excel Online (Business)-Option in Power Automate.":::

7. <span data-ttu-id="c1ef5-165">Wählen Sie unter **Aktionen** die Option **Skript ausführen** aus.</span><span class="sxs-lookup"><span data-stu-id="c1ef5-165">Under **Actions**, select **Run script**.</span></span>

    :::image type="content" source="../images/power-automate-tutorial-5.png" alt-text="Aktionsoption „Skript ausführen“ in Power Automate.":::

8. <span data-ttu-id="c1ef5-167">Als Nächstes wählen Sie die Arbeitsmappe, das Skript und die Eingabeargumente für das Skript aus, die im Datenfluss-Schritt verwendet werden sollen.</span><span class="sxs-lookup"><span data-stu-id="c1ef5-167">Next, you'll select the workbook, script, and script input arguments to use in the flow step.</span></span> <span data-ttu-id="c1ef5-168">In diesem Lernprogramm verwenden Sie die Arbeitsmappe, die Sie in Ihrem OneDrive erstellt haben. Sie könnten jedoch jede beliebige Arbeitsmappe auf einer OneDrive- oder SharePoint-Website verwenden.</span><span class="sxs-lookup"><span data-stu-id="c1ef5-168">For the tutorial, you'll use the workbook you created in your OneDrive, but you could use any workbook in a OneDrive or SharePoint site.</span></span> <span data-ttu-id="c1ef5-169">Geben Sie die folgenden Einstellungen für den Konnektor **Run script** an:</span><span class="sxs-lookup"><span data-stu-id="c1ef5-169">Specify the following settings for the **Run script** connector:</span></span>

    - <span data-ttu-id="c1ef5-170">**Location**: OneDrive for Business</span><span class="sxs-lookup"><span data-stu-id="c1ef5-170">**Location**: OneDrive for Business</span></span>
    - <span data-ttu-id="c1ef5-171">**Document Library**: OneDrive</span><span class="sxs-lookup"><span data-stu-id="c1ef5-171">**Document Library**: OneDrive</span></span>
    - <span data-ttu-id="c1ef5-172">**File**: MyWorkbook.xlsx *(Ausgewählt über den Dateibrowser)*</span><span class="sxs-lookup"><span data-stu-id="c1ef5-172">**File**: MyWorkbook.xlsx *(Chosen through the file browser)*</span></span>
    - <span data-ttu-id="c1ef5-173">**Script**: E-Mail aufzeichnen</span><span class="sxs-lookup"><span data-stu-id="c1ef5-173">**Script**: Record Email</span></span>
    - <span data-ttu-id="c1ef5-174">**from**: Von *(dynamischer Inhalt aus Outlook)*</span><span class="sxs-lookup"><span data-stu-id="c1ef5-174">**from**: From *(dynamic content from Outlook)*</span></span>
    - <span data-ttu-id="c1ef5-175">**dateReceived**: Uhrzeit des Empfangs *(dynamischer Inhalt aus Outlook)*</span><span class="sxs-lookup"><span data-stu-id="c1ef5-175">**dateReceived**: Received Time *(dynamic content from Outlook)*</span></span>
    - <span data-ttu-id="c1ef5-176">**subject**: Betreff *(dynamischer Inhalt aus Outlook)*</span><span class="sxs-lookup"><span data-stu-id="c1ef5-176">**subject**: Subject *(dynamic content from Outlook)*</span></span>

    <span data-ttu-id="c1ef5-177">*Beachten Sie, dass die Parameter für das Skript nur angezeigt werden, wenn das Skript ausgewählt wurde.*</span><span class="sxs-lookup"><span data-stu-id="c1ef5-177">*Note that the parameters for the script will only appear once the script is selected.*</span></span>

    :::image type="content" source="../images/power-automate-params-tutorial-3.png" alt-text="Die Aktion zum Ausführen eines Skripts in PowerAutomate zeigt die Optionen an, die erscheinen, nachdem das Skript ausgewählt wurde.":::

9. <span data-ttu-id="c1ef5-179">Wählen Sie **Speichern** aus.</span><span class="sxs-lookup"><span data-stu-id="c1ef5-179">Select **Save**.</span></span>

<span data-ttu-id="c1ef5-p118">Der Fluss ist nun aktiviert. Er wird das Skript automatisch jedes Mal ausführen, wenn Sie eine E-Mail über Outlook erhalten.</span><span class="sxs-lookup"><span data-stu-id="c1ef5-p118">Your flow is now enabled. It will automatically run your script each time you receive an email through Outlook.</span></span>

## <a name="manage-the-script-in-power-automate"></a><span data-ttu-id="c1ef5-182">Verwalten des Skripts in Power Automate</span><span class="sxs-lookup"><span data-stu-id="c1ef5-182">Manage the script in Power Automate</span></span>

1. <span data-ttu-id="c1ef5-183">Wählen Sie auf der Hauptseite der Power Automate-Seite **My Flows** aus.</span><span class="sxs-lookup"><span data-stu-id="c1ef5-183">From the main Power Automate page, select **My flows**.</span></span>

    :::image type="content" source="../images/power-automate-tutorial-7.png" alt-text="Die Schaltfläche „Meine Flows“ in Power Automate.":::

2. <span data-ttu-id="c1ef5-185">Wählen Sie Ihren Flow aus.</span><span class="sxs-lookup"><span data-stu-id="c1ef5-185">Select your flow.</span></span> <span data-ttu-id="c1ef5-186">Hier sehen Sie den Ausführungsverlauf.</span><span class="sxs-lookup"><span data-stu-id="c1ef5-186">Here you can see the run history.</span></span> <span data-ttu-id="c1ef5-187">Sie können die Seite aktualisieren, oder Sie können die Schaltfläche **All runs** auswählen, um den Verlauf zu aktualisieren.</span><span class="sxs-lookup"><span data-stu-id="c1ef5-187">You can refresh the page or select the refresh **All runs** button to update the history.</span></span> <span data-ttu-id="c1ef5-188">Der Flow wird kurz nach Empfang einer E-Mail ausgelöst.</span><span class="sxs-lookup"><span data-stu-id="c1ef5-188">The flow will trigger shortly after an email is received.</span></span> <span data-ttu-id="c1ef5-189">Testen Sie den Flow durch Senden von E-Mails.</span><span class="sxs-lookup"><span data-stu-id="c1ef5-189">Test the flow by sending yourself mail.</span></span>

<span data-ttu-id="c1ef5-190">Wenn der Flow ausgelöst und das Skript erfolgreich ausgeführt wird, sollten die Tabelle und die PivotTable der Arbeitsmappe aktualisiert werden.</span><span class="sxs-lookup"><span data-stu-id="c1ef5-190">When the flow is triggered and successfully runs your script, you should see the workbook's table and PivotTable update.</span></span>

:::image type="content" source="../images/power-automate-params-tutorial-4.png" alt-text="Ein Arbeitsblatt, auf dem die E-Mail-Tabelle angezeigt wird, nachdem der Fluss dreimal ausgeführt wurde.":::

:::image type="content" source="../images/power-automate-params-tutorial-5.png" alt-text="Ein Arbeitsblatt, auf dem die PivotTable angezeigt wird, nachdem der Fluss dreimal ausgeführt wurde":::

## <a name="next-steps"></a><span data-ttu-id="c1ef5-193">Nächste Schritte</span><span class="sxs-lookup"><span data-stu-id="c1ef5-193">Next steps</span></span>

<span data-ttu-id="c1ef5-194">Schließen Sie das Tutorial[Zurückgeben von Daten aus einem Skript an einen automatisch ausgeführten Power Automate-Flow](excel-power-automate-returns.md) ab.</span><span class="sxs-lookup"><span data-stu-id="c1ef5-194">Complete the [Return data from a script to an automatically-run Power Automate flow](excel-power-automate-returns.md) tutorial.</span></span> <span data-ttu-id="c1ef5-195">Hier lernen Sie, wie Sie Daten aus einem Skript an den Flow zurückgeben.</span><span class="sxs-lookup"><span data-stu-id="c1ef5-195">It teaches you how to return data from a script to the flow.</span></span>

<span data-ttu-id="c1ef5-196">Sie können sich auch das Beispielszenario [Automatisierte Aufgabenerinnerungen](../resources/scenarios/task-reminders.md) ansehen, um zu erfahren, wie Sie Office-Skripts und Power Automate mit Teams Adaptive Cards kombinieren können.</span><span class="sxs-lookup"><span data-stu-id="c1ef5-196">You can also check out the [Automated task reminders sample scenario](../resources/scenarios/task-reminders.md) to learn how to combine Office Scripts and Power Automate with Teams Adaptive Cards.</span></span>
