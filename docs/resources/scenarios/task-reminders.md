---
title: 'Office Skript-Beispielszenario: Automatisierte Aufgabenerinnerungen'
description: Ein Beispiel, das Power Automate und adaptive Karten verwendet, automatisiert Aufgabenerinnerungen in einer Projektverwaltungstabelle.
ms.date: 06/29/2021
localization_priority: Normal
ms.openlocfilehash: cf25b81ad44bbe963083f6a8346c0fd59a514305
ms.sourcegitcommit: 211c157ca746e266eeb079f5fa1925a1e35ab702
ms.translationtype: MT
ms.contentlocale: de-DE
ms.lasthandoff: 07/07/2021
ms.locfileid: "53313981"
---
# <a name="office-scripts-sample-scenario-automated-task-reminders"></a><span data-ttu-id="6b9ed-103">Office Skript-Beispielszenario: Automatisierte Aufgabenerinnerungen</span><span class="sxs-lookup"><span data-stu-id="6b9ed-103">Office Scripts sample scenario: Automated task reminders</span></span>

<span data-ttu-id="6b9ed-104">In diesem Szenario verwalten Sie ein Projekt.</span><span class="sxs-lookup"><span data-stu-id="6b9ed-104">In this scenario you're managing a project.</span></span> <span data-ttu-id="6b9ed-105">Sie verwenden ein Excel Arbeitsblatt, um den Status Ihrer Mitarbeiter jeden Monat nachzuverfolgen.</span><span class="sxs-lookup"><span data-stu-id="6b9ed-105">You use an Excel worksheet to track your employees' status every month.</span></span> <span data-ttu-id="6b9ed-106">Sie müssen die Personen häufig daran erinnern, ihren Status auszufüllen, daher haben Sie sich entschieden, diesen Erinnerungsprozess zu automatisieren.</span><span class="sxs-lookup"><span data-stu-id="6b9ed-106">You often need to remind people to fill out their status, so you've decided to automate that reminder process.</span></span>

<span data-ttu-id="6b9ed-107">Sie erstellen einen Power Automate Fluss an Nachrichtenmitarbeiter mit fehlenden Statusfeldern und wenden deren Antworten auf das Arbeitsblatt an.</span><span class="sxs-lookup"><span data-stu-id="6b9ed-107">You'll create a Power Automate flow to message people with missing status fields and apply their responses to the spreadsheet.</span></span> <span data-ttu-id="6b9ed-108">Zu diesem Zweck entwickeln Sie ein Skriptpaar, um die Arbeit mit der Arbeitsmappe zu verarbeiten.</span><span class="sxs-lookup"><span data-stu-id="6b9ed-108">To do this, you'll develop a pair of scripts to handle the working with the workbook.</span></span> <span data-ttu-id="6b9ed-109">Das erste Skript ruft eine Liste von Personen mit leeren Status ab, und das zweite Skript fügt der rechten Zeile eine Statuszeichenfolge hinzu.</span><span class="sxs-lookup"><span data-stu-id="6b9ed-109">The first script gets a list of people with blank statuses and the second script adds a status string to the right row.</span></span> <span data-ttu-id="6b9ed-110">Sie verwenden auch [Teams adaptive Karten,](/microsoftteams/platform/task-modules-and-cards/what-are-cards) damit Mitarbeiter ihren Status direkt über die Benachrichtigung eingeben können.</span><span class="sxs-lookup"><span data-stu-id="6b9ed-110">You'll also make use of [Teams Adaptive Cards](/microsoftteams/platform/task-modules-and-cards/what-are-cards) to have employees enter their status directly from the notification.</span></span>

## <a name="scripting-skills-covered"></a><span data-ttu-id="6b9ed-111">Abgedeckte Skriptfähigkeiten</span><span class="sxs-lookup"><span data-stu-id="6b9ed-111">Scripting skills covered</span></span>

- <span data-ttu-id="6b9ed-112">Erstellen von Flüssen in Power Automate</span><span class="sxs-lookup"><span data-stu-id="6b9ed-112">Create flows in Power Automate</span></span>
- <span data-ttu-id="6b9ed-113">Übergeben von Daten an Skripts</span><span class="sxs-lookup"><span data-stu-id="6b9ed-113">Pass data to scripts</span></span>
- <span data-ttu-id="6b9ed-114">Zurückgeben von Daten aus Skripts</span><span class="sxs-lookup"><span data-stu-id="6b9ed-114">Return data from scripts</span></span>
- <span data-ttu-id="6b9ed-115">Teams Adaptive Karten</span><span class="sxs-lookup"><span data-stu-id="6b9ed-115">Teams Adaptive Cards</span></span>
- <span data-ttu-id="6b9ed-116">Tabellen</span><span class="sxs-lookup"><span data-stu-id="6b9ed-116">Tables</span></span>

## <a name="prerequisites"></a><span data-ttu-id="6b9ed-117">Voraussetzungen</span><span class="sxs-lookup"><span data-stu-id="6b9ed-117">Prerequisites</span></span>

<span data-ttu-id="6b9ed-118">In diesem Szenario [werden Power Automate](https://flow.microsoft.com) und [Microsoft Teams](https://www.microsoft.com/microsoft-365/microsoft-teams/group-chat-software)verwendet.</span><span class="sxs-lookup"><span data-stu-id="6b9ed-118">This scenario uses [Power Automate](https://flow.microsoft.com) and [Microsoft Teams](https://www.microsoft.com/microsoft-365/microsoft-teams/group-chat-software).</span></span> <span data-ttu-id="6b9ed-119">Sie benötigen beides, das mit dem Konto verknüpft ist, das Sie für die Entwicklung von Office Skripts verwenden.</span><span class="sxs-lookup"><span data-stu-id="6b9ed-119">You will need both associated with the account that you use for developing Office Scripts.</span></span> <span data-ttu-id="6b9ed-120">Für den kostenlosen Zugriff auf ein Microsoft Developer-Abonnement, um mehr über diese Anwendungen zu erfahren und mit diesen zu arbeiten, sollten Sie am [Microsoft 365 Developer Program](https://developer.microsoft.com/microsoft-365/dev-program)teilnehmen.</span><span class="sxs-lookup"><span data-stu-id="6b9ed-120">For free access to a Microsoft Developer subscription to learn about and work with these applications, consider joining the [Microsoft 365 Developer Program](https://developer.microsoft.com/microsoft-365/dev-program).</span></span>

## <a name="setup-instructions"></a><span data-ttu-id="6b9ed-121">Setup-Anweisungen</span><span class="sxs-lookup"><span data-stu-id="6b9ed-121">Setup instructions</span></span>

1. <span data-ttu-id="6b9ed-122">Laden Sie <a href="task-reminders.xlsx">task-reminders.xlsx</a> auf Ihre OneDrive herunter.</span><span class="sxs-lookup"><span data-stu-id="6b9ed-122">Download <a href="task-reminders.xlsx">task-reminders.xlsx</a> to your OneDrive.</span></span>

1. <span data-ttu-id="6b9ed-123">Öffnen Sie die Arbeitsmappe in Excel im Web.</span><span class="sxs-lookup"><span data-stu-id="6b9ed-123">Open the workbook in Excel on the web.</span></span>

1. <span data-ttu-id="6b9ed-124">Zunächst benötigen wir ein Skript, um alle Mitarbeiter mit Statusberichten zu erhalten, die in der Tabelle fehlen.</span><span class="sxs-lookup"><span data-stu-id="6b9ed-124">First, we need a script to get all the employees with status reports that are missing from the spreadsheet.</span></span> <span data-ttu-id="6b9ed-125">Wählen Sie auf der Registerkarte **"Automatisieren"** die Option **"Neues Skript"** aus, und fügen Sie das folgende Skript in den Editor ein.</span><span class="sxs-lookup"><span data-stu-id="6b9ed-125">Under the **Automate** tab, select **New Script** and paste the following script into the editor.</span></span>

    ```TypeScript
    /**
     * This script looks for missing status reports in a project management table.
     *
     * @returns An array of Employee objects (containing their names and emails).
     */
    function main(workbook: ExcelScript.Workbook): Employee[] {
      // Get the first worksheet and the first table on that worksheet.
      let sheet = workbook.getFirstWorksheet()
      let table = sheet.getTables()[0];

      // Give the column indices names matching their expected content.
      const NAME_INDEX = 0;
      const EMAIL_INDEX = 1;
      const STATUS_REPORT_INDEX = 2;

      // Get the data for the whole table.
      let bodyRangeValues = table.getRangeBetweenHeaderAndTotal().getValues();

      // Create the array of Employee objects to return.
      let people: Employee[] = [];

      // Loop through the table and check each row for completion.
      for (let i = 0; i < bodyRangeValues.length; i++) {
        let row = bodyRangeValues[i];
        if (row[STATUS_REPORT_INDEX] === "") {
          // Save the email to return.
          people.push({ name: row[NAME_INDEX].toString(), email: row[EMAIL_INDEX].toString() });
        }
      }

      // Log the array to verify we're getting the right rows.
      console.log(people);

      // Return the array of Employees.
      return people;
    }

    /**
     * An interface representing an employee.
     * An array of Employees will be returned from the script
     * for the Power Automate flow.
     */
    interface Employee {
      name: string;
      email: string;
    }
    ```

1. <span data-ttu-id="6b9ed-126">Speichern Sie das Skript mit dem Namen **"Personen abrufen".**</span><span class="sxs-lookup"><span data-stu-id="6b9ed-126">Save the script with the name **Get People**.</span></span>

1. <span data-ttu-id="6b9ed-127">Als Nächstes benötigen wir ein zweites Skript, um die Statusberichtskarten zu verarbeiten und die neuen Informationen in die Tabelle einzufügen.</span><span class="sxs-lookup"><span data-stu-id="6b9ed-127">Next, we need a second script to process the status report cards and put the new information in the spreadsheet.</span></span> <span data-ttu-id="6b9ed-128">Wählen Sie im Code-Editor-Aufgabenbereich **"Neues Skript"** aus, und fügen Sie das folgende Skript in den Editor ein.</span><span class="sxs-lookup"><span data-stu-id="6b9ed-128">In the Code Editor task pane, select **New Script** and paste the following script into the editor.</span></span>

    ```TypeScript
    /**
     * This script applies the results of a Teams Adaptive Card about
     * a status update to a project management table.
     *
     * @param senderEmail - The email address of the employee updating their status.
     * @param statusReportResponse - The employee's status report.
     */
    function main(workbook: ExcelScript.Workbook,
      senderEmail: string,
      statusReportResponse: string) {

      // Get the first worksheet and the first table in that worksheet.
      let sheet = workbook.getFirstWorksheet();
      let table = sheet.getTables()[0];

      // Give the column indices names matching their expected content.
      const NAME_INDEX = 0;
      const EMAIL_INDEX = 1;
      const STATUS_REPORT_INDEX = 2;

      // Get the range and data for the whole table.
      let bodyRange = table.getRangeBetweenHeaderAndTotal();
      let tableRowCount = bodyRange.getRowCount();
      let bodyRangeValues = bodyRange.getValues();

      // Create a flag to denote success.
      let statusAdded = false;

      // Loop through the table and check each row for a matching email address.
      for (let i = 0; i < tableRowCount && !statusAdded; i++) {
        let row = bodyRangeValues[i];

        // Check if the row's email address matches.
        if (row[EMAIL_INDEX] === senderEmail) {
          // Add the Teams Adaptive Card response to the table.
          bodyRange.getCell(i, STATUS_REPORT_INDEX).setValues([
            [statusReportResponse]
          ]);
          statusAdded = true;
        }
      }

      // If successful, log the status update.
      if (statusAdded) {
        console.log(
          `Successfully added status report for ${senderEmail} containing: ${statusReportResponse}`
        );
      }
    }
    ```

1. <span data-ttu-id="6b9ed-129">Speichern Sie das Skript mit dem Namen **"Save Status".**</span><span class="sxs-lookup"><span data-stu-id="6b9ed-129">Save the script with the name **Save Status**.</span></span>

1. <span data-ttu-id="6b9ed-130">Jetzt müssen wir den Fluss erstellen.</span><span class="sxs-lookup"><span data-stu-id="6b9ed-130">Now, we need to create the flow.</span></span> <span data-ttu-id="6b9ed-131">Öffnen [Sie Power Automate](https://flow.microsoft.com/).</span><span class="sxs-lookup"><span data-stu-id="6b9ed-131">Open [Power Automate](https://flow.microsoft.com/).</span></span>

    > [!TIP]
    > <span data-ttu-id="6b9ed-132">Wenn Sie noch keinen Flow erstellt haben, sehen Sie sich unser Lernprogramm ["Start using scripts with Power Automate"](../../tutorials/excel-power-automate-manual.md) an, um die Grundlagen zu erlernen.</span><span class="sxs-lookup"><span data-stu-id="6b9ed-132">If you haven't created a flow before, please check out our tutorial [Start using scripts with Power Automate](../../tutorials/excel-power-automate-manual.md) to learn the basics.</span></span>

1. <span data-ttu-id="6b9ed-133">Erstellen Sie einen neuen **Sofortablauf.**</span><span class="sxs-lookup"><span data-stu-id="6b9ed-133">Create a new **Instant flow**.</span></span>

1. <span data-ttu-id="6b9ed-134">Wählen Sie manuell einen Fluss aus den Optionen **aus,** und wählen Sie **Erstellen** aus.</span><span class="sxs-lookup"><span data-stu-id="6b9ed-134">Choose **Manually trigger a flow** from the options and select **Create**.</span></span>

1. <span data-ttu-id="6b9ed-135">Der Fluss muss das Skript **"Personen** abrufen" aufrufen, um alle Mitarbeiter mit leeren Statusfeldern abzurufen.</span><span class="sxs-lookup"><span data-stu-id="6b9ed-135">The flow needs to call the **Get People** script to get all the employees with empty status fields.</span></span> <span data-ttu-id="6b9ed-136">Wählen Sie **"Neu" aus,** und wählen Sie dann **Excel Online (Business)** aus.</span><span class="sxs-lookup"><span data-stu-id="6b9ed-136">Select **New step**, then select **Excel Online (Business)**.</span></span> <span data-ttu-id="6b9ed-137">Wählen Sie unter **Aktionen** die Option **Skript ausführen** aus.</span><span class="sxs-lookup"><span data-stu-id="6b9ed-137">Under **Actions**, select **Run script**.</span></span> <span data-ttu-id="6b9ed-138">Geben Sie die folgenden Einträge für den Flussschritt an:</span><span class="sxs-lookup"><span data-stu-id="6b9ed-138">Provide the following entries for the flow step:</span></span>

    - <span data-ttu-id="6b9ed-139">**Location**: OneDrive for Business</span><span class="sxs-lookup"><span data-stu-id="6b9ed-139">**Location**: OneDrive for Business</span></span>
    - <span data-ttu-id="6b9ed-140">**Document Library**: OneDrive</span><span class="sxs-lookup"><span data-stu-id="6b9ed-140">**Document Library**: OneDrive</span></span>
    - <span data-ttu-id="6b9ed-141">**Datei:** task-reminders.xlsx *(über den Dateibrowser ausgewählt)*</span><span class="sxs-lookup"><span data-stu-id="6b9ed-141">**File**: task-reminders.xlsx *(Chosen through the file browser)*</span></span>
    - <span data-ttu-id="6b9ed-142">**Skript:** Personen abrufen</span><span class="sxs-lookup"><span data-stu-id="6b9ed-142">**Script**: Get People</span></span>

    :::image type="content" source="../../images/scenario-task-reminders-first-flow-step.png" alt-text="Der Power Automate-Ablauf mit dem ersten Ausführungsskriptflussschritt.":::

1. <span data-ttu-id="6b9ed-144">Als Nächstes muss der Fluss jeden Mitarbeiter in dem array verarbeiten, das vom Skript zurückgegeben wird.</span><span class="sxs-lookup"><span data-stu-id="6b9ed-144">Next, the flow needs to process each Employee in the array returned by the script.</span></span> <span data-ttu-id="6b9ed-145">Wählen Sie **"Neuer Schritt"** aus, und wählen Sie dann **"Adaptive Karte an einen Teams Benutzer posten" aus, und warten Sie auf eine Antwort.**</span><span class="sxs-lookup"><span data-stu-id="6b9ed-145">Select **New step**, then choose **Post an Adaptive Card to a Teams user and wait for a response**.</span></span>

1. <span data-ttu-id="6b9ed-146">Fügen Sie für das **Feld "Empfänger"** **E-Mails** aus dem dynamischen Inhalt hinzu (die Auswahl enthält das Excel Logo).</span><span class="sxs-lookup"><span data-stu-id="6b9ed-146">For the **Recipient** field, add **email** from the dynamic content (the selection will have the Excel logo by it).</span></span> <span data-ttu-id="6b9ed-147">Durch das Hinzufügen von **E-Mails** wird der Flussschritt von einem **Apply-Element für jeden** Block umgeben.</span><span class="sxs-lookup"><span data-stu-id="6b9ed-147">Adding **email** causes the flow step to be surrounded by an **Apply to each** block.</span></span> <span data-ttu-id="6b9ed-148">Das bedeutet, dass das Array von Power Automate durchlaufen wird.</span><span class="sxs-lookup"><span data-stu-id="6b9ed-148">That means the array will be iterated over by Power Automate.</span></span>

1. <span data-ttu-id="6b9ed-149">Das Senden einer adaptiven Karte erfordert, dass der JSON-Code der Karte als **Nachricht** bereitgestellt wird.</span><span class="sxs-lookup"><span data-stu-id="6b9ed-149">Sending an Adaptive Card requires the card's JSON to be provided as the **Message**.</span></span> <span data-ttu-id="6b9ed-150">Sie können den Designer für [adaptive Karten](https://adaptivecards.io/designer/) verwenden, um benutzerdefinierte Karten zu erstellen.</span><span class="sxs-lookup"><span data-stu-id="6b9ed-150">You can use the [Adaptive Card Designer](https://adaptivecards.io/designer/) to create custom cards.</span></span> <span data-ttu-id="6b9ed-151">Verwenden Sie für dieses Beispiel den folgenden JSON-Code.</span><span class="sxs-lookup"><span data-stu-id="6b9ed-151">For this sample, use the following JSON.</span></span>  

    ```json
    {
      "$schema": "http://adaptivecards.io/schemas/adaptive-card.json",
      "type": "AdaptiveCard",
      "version": "1.0",
      "body": [
        {
          "type": "TextBlock",
          "size": "Medium",
          "weight": "Bolder",
          "text": "Update your Status Report"
        },
        {
          "type": "Image",
          "altText": "",
          "url": "https://i.imgur.com/f5RcuF3.png"
        },
        {
          "type": "TextBlock",
          "text": "This is a reminder to update your status report for this month's review. You can do so right here in this card, or by adding it directly to the spreadsheet.",
          "wrap": true
        },
        {
          "type": "Input.Text",
          "placeholder": "My status report for this month is...",
          "id": "response",
          "isMultiline": true
        }
      ],
      "actions": [
        {
          "type": "Action.Submit",
          "title": "Submit",
          "id": "submit"
        }
      ]
    }
    ```

1. <span data-ttu-id="6b9ed-152">Füllen Sie die verbleibenden Felder wie folgt aus:</span><span class="sxs-lookup"><span data-stu-id="6b9ed-152">Fill out the remaining fields as follows:</span></span>

    - <span data-ttu-id="6b9ed-153">**Nachricht aktualisieren:** Vielen Dank für die Übermittlung Ihres Statusberichts.</span><span class="sxs-lookup"><span data-stu-id="6b9ed-153">**Update message**: Thank you for submitting your status report.</span></span> <span data-ttu-id="6b9ed-154">Ihre Antwort wurde erfolgreich zur Kalkulationstabelle hinzugefügt.</span><span class="sxs-lookup"><span data-stu-id="6b9ed-154">Your response has been successfully added to the spreadsheet.</span></span>
    - <span data-ttu-id="6b9ed-155">**Sollte Karte aktualisieren:** Ja</span><span class="sxs-lookup"><span data-stu-id="6b9ed-155">**Should update card**: Yes</span></span>

1. <span data-ttu-id="6b9ed-156">Wählen Sie im Abschnitt **"Auf jeden** Block anwenden" nach dem **Posten einer adaptiven Karte für einen Teams Benutzer und Warten auf eine Antwort** eine Aktion **hinzufügen** aus.</span><span class="sxs-lookup"><span data-stu-id="6b9ed-156">In the **Apply to each** block, following the **Post an Adaptive Card to a Teams user and wait for a response**, select **Add an action**.</span></span> <span data-ttu-id="6b9ed-157">Wählen Sie **Excel Online (Business)** aus.</span><span class="sxs-lookup"><span data-stu-id="6b9ed-157">Select **Excel Online (Business)**.</span></span> <span data-ttu-id="6b9ed-158">Wählen Sie unter **Aktionen** die Option **Skript ausführen** aus.</span><span class="sxs-lookup"><span data-stu-id="6b9ed-158">Under **Actions**, select **Run script**.</span></span> <span data-ttu-id="6b9ed-159">Geben Sie die folgenden Einträge für den Flussschritt an:</span><span class="sxs-lookup"><span data-stu-id="6b9ed-159">Provide the following entries for the flow step:</span></span>

    - <span data-ttu-id="6b9ed-160">**Location**: OneDrive for Business</span><span class="sxs-lookup"><span data-stu-id="6b9ed-160">**Location**: OneDrive for Business</span></span>
    - <span data-ttu-id="6b9ed-161">**Document Library**: OneDrive</span><span class="sxs-lookup"><span data-stu-id="6b9ed-161">**Document Library**: OneDrive</span></span>
    - <span data-ttu-id="6b9ed-162">**Datei:** task-reminders.xlsx *(über den Dateibrowser ausgewählt)*</span><span class="sxs-lookup"><span data-stu-id="6b9ed-162">**File**: task-reminders.xlsx *(Chosen through the file browser)*</span></span>
    - <span data-ttu-id="6b9ed-163">**Skript:** Save Status</span><span class="sxs-lookup"><span data-stu-id="6b9ed-163">**Script**: Save Status</span></span>
    - <span data-ttu-id="6b9ed-164">**senderEmail**: E-Mail *(dynamischer Inhalt von Excel)*</span><span class="sxs-lookup"><span data-stu-id="6b9ed-164">**senderEmail**: email *(dynamic content from Excel)*</span></span>
    - <span data-ttu-id="6b9ed-165">**statusReportResponse:** response *(dynamic content from Teams)*</span><span class="sxs-lookup"><span data-stu-id="6b9ed-165">**statusReportResponse**: response *(dynamic content from Teams)*</span></span>

    :::image type="content" source="../../images/scenario-task-reminders-last-flow-step.png" alt-text="Der Power Automate-Ablauf, in dem die einzelnen Schritte angezeigt werden.":::

1. <span data-ttu-id="6b9ed-167">Speichern Sie den Fluss.</span><span class="sxs-lookup"><span data-stu-id="6b9ed-167">Save the flow.</span></span>

## <a name="running-the-flow"></a><span data-ttu-id="6b9ed-168">Ausführen des Flusses</span><span class="sxs-lookup"><span data-stu-id="6b9ed-168">Running the flow</span></span>

<span data-ttu-id="6b9ed-169">Um den Fluss zu testen, stellen Sie sicher, dass alle Tabellenzeilen mit leerem Status eine E-Mail-Adresse verwenden, die an ein Teams Konto gebunden ist (Sie sollten beim Testen wahrscheinlich Ihre eigene E-Mail-Adresse verwenden).</span><span class="sxs-lookup"><span data-stu-id="6b9ed-169">To test the flow, make sure any table rows with blank status use an email address tied to a Teams account (you should probably use your own email address while testing).</span></span> <span data-ttu-id="6b9ed-170">Verwenden Sie die Schaltfläche **"Test"** auf der Flow-Editor-Seite, oder führen Sie den Fluss über Ihre Registerkarte **"Meine Flüsse"** aus. Achten Sie darauf, den Zugriff zuzulassen, wenn Sie dazu aufgefordert werden.</span><span class="sxs-lookup"><span data-stu-id="6b9ed-170">Use the **Test** button on the flow editor page or run the flow through your **My flows** tab. Be sure to allow access when prompted.</span></span>

<span data-ttu-id="6b9ed-171">Sie sollten eine adaptive Karte von Power Automate bis Teams erhalten.</span><span class="sxs-lookup"><span data-stu-id="6b9ed-171">You should receive an Adaptive Card from Power Automate through Teams.</span></span> <span data-ttu-id="6b9ed-172">Nachdem Sie das Statusfeld in der Karte ausgefüllt haben, wird der Fluss fortgesetzt und die Tabelle mit dem von Ihnen angegebenen Status aktualisiert.</span><span class="sxs-lookup"><span data-stu-id="6b9ed-172">Once you fill out the status field in the card, the flow will continue and update the spreadsheet with the status you provide.</span></span>

### <a name="before-running-the-flow"></a><span data-ttu-id="6b9ed-173">Vor dem Ausführen des Flusses</span><span class="sxs-lookup"><span data-stu-id="6b9ed-173">Before running the flow</span></span>

:::image type="content" source="../../images/scenario-task-reminders-spreadsheet-before.png" alt-text="Ein Arbeitsblatt mit einem Statusbericht, der einen fehlenden Statuseintrag enthält.":::

### <a name="receiving-the-adaptive-card"></a><span data-ttu-id="6b9ed-175">Empfangen der adaptiven Karte</span><span class="sxs-lookup"><span data-stu-id="6b9ed-175">Receiving the Adaptive Card</span></span>

:::image type="content" source="../../images/scenario-task-reminders-adaptive-card.png" alt-text="Eine adaptive Karte in Teams, die den Mitarbeiter um eine Statusaktualisierung bittet.":::

### <a name="after-running-the-flow"></a><span data-ttu-id="6b9ed-177">Nach dem Ausführen des Flusses</span><span class="sxs-lookup"><span data-stu-id="6b9ed-177">After running the flow</span></span>

:::image type="content" source="../../images/scenario-task-reminders-spreadsheet-after.png" alt-text="Ein Arbeitsblatt mit einem Statusbericht mit einem jetzt ausgefüllten Statuseintrag.":::
