---
title: 'Office Skripts-Beispielszenario: Automatisierte Aufgabenerinnerungen'
description: Ein Beispiel, das Power Automate und Adaptive Cards verwendet, automatisiert Vorgangserinnerungen in einer Projektmanagementtabelle.
ms.date: 11/30/2020
localization_priority: Normal
ms.openlocfilehash: c254a627da8442c0974263908a41275182740b6e
ms.sourcegitcommit: 4687693f02fc90a57ba30c461f35046e02e6f5fb
ms.translationtype: MT
ms.contentlocale: de-DE
ms.lasthandoff: 05/19/2021
ms.locfileid: "52545605"
---
# <a name="office-scripts-sample-scenario-automated-task-reminders"></a><span data-ttu-id="eb630-103">Office Skripts-Beispielszenario: Automatisierte Aufgabenerinnerungen</span><span class="sxs-lookup"><span data-stu-id="eb630-103">Office Scripts sample scenario: Automated task reminders</span></span>

<span data-ttu-id="eb630-104">In diesem Szenario verwalten Sie ein Projekt.</span><span class="sxs-lookup"><span data-stu-id="eb630-104">In this scenario you're managing a project.</span></span> <span data-ttu-id="eb630-105">Sie verwenden einen Excel Arbeitsblatt, um den Status Ihrer Mitarbeiter jeden Monat nachzuverfolgen.</span><span class="sxs-lookup"><span data-stu-id="eb630-105">You use an Excel worksheet to track your employees' status every month.</span></span> <span data-ttu-id="eb630-106">Sie müssen die Benutzer oft daran erinnern, ihren Status auszufüllen, daher haben Sie sich entschieden, diesen Erinnerungsprozess zu automatisieren.</span><span class="sxs-lookup"><span data-stu-id="eb630-106">You often need to remind people to fill out their status, so you've decided to automate that reminder process.</span></span>

<span data-ttu-id="eb630-107">Sie erstellen einen Power Automate Fluss, um Personen mit fehlenden Statusfeldern zu nachrichten und ihre Antworten auf die Kalkulationstabelle anzuwenden.</span><span class="sxs-lookup"><span data-stu-id="eb630-107">You'll create a Power Automate flow to message people with missing status fields and apply their responses to the spreadsheet.</span></span> <span data-ttu-id="eb630-108">Dazu entwickeln Sie ein Skriptpaar, um die Arbeit mit der Arbeitsmappe zu verarbeiten.</span><span class="sxs-lookup"><span data-stu-id="eb630-108">To do this, you'll develop a pair of scripts to handle the working with the workbook.</span></span> <span data-ttu-id="eb630-109">Das erste Skript ruft eine Liste von Personen mit leeren Status ab, und das zweite Skript fügt der rechten Zeile eine Statuszeichenfolge hinzu.</span><span class="sxs-lookup"><span data-stu-id="eb630-109">The first script gets a list of people with blank statuses and the second script adds a status string to the right row.</span></span> <span data-ttu-id="eb630-110">Sie nutzen auch [Teams Adaptive Cards,](/microsoftteams/platform/task-modules-and-cards/what-are-cards) damit Mitarbeiter ihren Status direkt aus der Benachrichtigung eingeben.</span><span class="sxs-lookup"><span data-stu-id="eb630-110">You'll also make use of [Teams Adaptive Cards](/microsoftteams/platform/task-modules-and-cards/what-are-cards) to have employees enter their status directly from the notification.</span></span>

## <a name="scripting-skills-covered"></a><span data-ttu-id="eb630-111">Scripting-Fähigkeiten abgedeckt</span><span class="sxs-lookup"><span data-stu-id="eb630-111">Scripting skills covered</span></span>

- <span data-ttu-id="eb630-112">Erstellen von Flows in Power Automate</span><span class="sxs-lookup"><span data-stu-id="eb630-112">Create flows in Power Automate</span></span>
- <span data-ttu-id="eb630-113">Übergeben von Daten an Skripts</span><span class="sxs-lookup"><span data-stu-id="eb630-113">Pass data to scripts</span></span>
- <span data-ttu-id="eb630-114">Zurückgeben von Daten aus Skripts</span><span class="sxs-lookup"><span data-stu-id="eb630-114">Return data from scripts</span></span>
- <span data-ttu-id="eb630-115">Teams Adaptive Karten</span><span class="sxs-lookup"><span data-stu-id="eb630-115">Teams Adaptive Cards</span></span>
- <span data-ttu-id="eb630-116">Tabellen</span><span class="sxs-lookup"><span data-stu-id="eb630-116">Tables</span></span>

## <a name="prerequisites"></a><span data-ttu-id="eb630-117">Voraussetzungen</span><span class="sxs-lookup"><span data-stu-id="eb630-117">Prerequisites</span></span>

<span data-ttu-id="eb630-118">In diesem Szenario werden [Power Automate](https://flow.microsoft.com) und [Microsoft Teams verwendet.](https://www.microsoft.com/microsoft-365/microsoft-teams/group-chat-software)</span><span class="sxs-lookup"><span data-stu-id="eb630-118">This scenario uses [Power Automate](https://flow.microsoft.com) and [Microsoft Teams](https://www.microsoft.com/microsoft-365/microsoft-teams/group-chat-software).</span></span> <span data-ttu-id="eb630-119">Sie müssen beide dem Konto zugeordnet sein, das Sie für die Entwicklung Office Skripts verwenden.</span><span class="sxs-lookup"><span data-stu-id="eb630-119">You will need both associated with the account that you use for developing Office Scripts.</span></span> <span data-ttu-id="eb630-120">Für den kostenlosen Zugriff auf ein Microsoft Developer-Abonnement, um mehr über diese Anwendungen zu erfahren und mit ihnen zu arbeiten, sollten Sie dem [Microsoft 365 Developer Program](https://developer.microsoft.com/microsoft-365/dev-program)beitreten.</span><span class="sxs-lookup"><span data-stu-id="eb630-120">For free access to a Microsoft Developer subscription to learn about and work with these applications, consider joining the [Microsoft 365 Developer Program](https://developer.microsoft.com/microsoft-365/dev-program).</span></span>

## <a name="setup-instructions"></a><span data-ttu-id="eb630-121">Setup-Anweisungen</span><span class="sxs-lookup"><span data-stu-id="eb630-121">Setup instructions</span></span>

1. <span data-ttu-id="eb630-122">Laden Sie <a href="task-reminders.xlsx">task-reminders.xlsx</a> auf Ihre OneDrive herunter.</span><span class="sxs-lookup"><span data-stu-id="eb630-122">Download <a href="task-reminders.xlsx">task-reminders.xlsx</a> to your OneDrive.</span></span>

2. <span data-ttu-id="eb630-123">Öffnen Sie die Arbeitsmappe in Excel im Web.</span><span class="sxs-lookup"><span data-stu-id="eb630-123">Open the workbook in Excel on the web.</span></span>

3. <span data-ttu-id="eb630-124">Öffnen Sie unter der Registerkarte **Automatisieren** **alle Skripts**.</span><span class="sxs-lookup"><span data-stu-id="eb630-124">Under the **Automate** tab, open **All Scripts**.</span></span>

4. <span data-ttu-id="eb630-125">Zuerst benötigen wir ein Skript, um alle Mitarbeiter mit Statusberichten zu erhalten, die in der Kalkulationstabelle fehlen.</span><span class="sxs-lookup"><span data-stu-id="eb630-125">First, we need a script to get all the employees with status reports that are missing from the spreadsheet.</span></span> <span data-ttu-id="eb630-126">Drücken Sie im Aufgabenbereich **Code-Editor** auf **Neues Skript,** und fügen Sie das folgende Skript in den Editor ein.</span><span class="sxs-lookup"><span data-stu-id="eb630-126">In the **Code Editor** task pane, press **New Script** and paste the following script into the editor.</span></span>

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

5. <span data-ttu-id="eb630-127">Speichern Sie das Skript mit dem Namen **Get People**.</span><span class="sxs-lookup"><span data-stu-id="eb630-127">Save the script with the name **Get People**.</span></span>

6. <span data-ttu-id="eb630-128">Als Nächstes benötigen wir ein zweites Skript, um die Statusberichtskarten zu verarbeiten und die neuen Informationen in die Kalkulationstabelle zu setzen.</span><span class="sxs-lookup"><span data-stu-id="eb630-128">Next, we need a second script to process the status report cards and put the new information in the spreadsheet.</span></span> <span data-ttu-id="eb630-129">Drücken Sie im Aufgabenbereich **Code-Editor** auf **Neues Skript,** und fügen Sie das folgende Skript in den Editor ein.</span><span class="sxs-lookup"><span data-stu-id="eb630-129">In the **Code Editor** task pane, press **New Script** and paste the following script into the editor.</span></span>

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

7. <span data-ttu-id="eb630-130">Speichern Sie das Skript mit dem Namen **Status speichern**.</span><span class="sxs-lookup"><span data-stu-id="eb630-130">Save the script with the name **Save Status**.</span></span>

8. <span data-ttu-id="eb630-131">Jetzt müssen wir den Fluss schaffen.</span><span class="sxs-lookup"><span data-stu-id="eb630-131">Now, we need to create the flow.</span></span> <span data-ttu-id="eb630-132">Öffnen [Sie Power Automate](https://flow.microsoft.com/).</span><span class="sxs-lookup"><span data-stu-id="eb630-132">Open [Power Automate](https://flow.microsoft.com/).</span></span>

    > [!TIP]
    > <span data-ttu-id="eb630-133">Wenn Sie noch keinen Flow erstellt haben, schauen Sie sich bitte unser Tutorial [Beginnen Sie mit Skripten mit Power Automate,](../../tutorials/excel-power-automate-manual.md) um die Grundlagen zu lernen.</span><span class="sxs-lookup"><span data-stu-id="eb630-133">If you haven't created a flow before, please check out our tutorial [Start using scripts with Power Automate](../../tutorials/excel-power-automate-manual.md) to learn the basics.</span></span>

9. <span data-ttu-id="eb630-134">Erstellen Sie einen neuen **Instant-Flow**.</span><span class="sxs-lookup"><span data-stu-id="eb630-134">Create a new **Instant flow**.</span></span>

10. <span data-ttu-id="eb630-135">Wählen Sie **Manuell auslösen einen Fluss** aus den Optionen und drücken Sie **Erstellen**.</span><span class="sxs-lookup"><span data-stu-id="eb630-135">Choose **Manually trigger a flow** from the options and press **Create**.</span></span>

11. <span data-ttu-id="eb630-136">Der Flow muss das **Skript "Personen abrufen"** aufrufen, um alle Mitarbeiter mit leeren Statusfeldern zu erhalten.</span><span class="sxs-lookup"><span data-stu-id="eb630-136">The flow needs to call the **Get People** script to get all the employees with empty status fields.</span></span> <span data-ttu-id="eb630-137">Drücken Sie **Den neuen Schritt,** und wählen Sie **Excel Online (Geschäft)** aus.</span><span class="sxs-lookup"><span data-stu-id="eb630-137">Press **New step** and select **Excel Online (Business)**.</span></span> <span data-ttu-id="eb630-138">Wählen Sie unter **Aktionen** die Option **Skript ausführen** aus.</span><span class="sxs-lookup"><span data-stu-id="eb630-138">Under **Actions**, select **Run script**.</span></span> <span data-ttu-id="eb630-139">Geben Sie die folgenden Einträge für den Flow-Schritt an:</span><span class="sxs-lookup"><span data-stu-id="eb630-139">Provide the following entries for the flow step:</span></span>

    - <span data-ttu-id="eb630-140">**Location**: OneDrive for Business</span><span class="sxs-lookup"><span data-stu-id="eb630-140">**Location**: OneDrive for Business</span></span>
    - <span data-ttu-id="eb630-141">**Document Library**: OneDrive</span><span class="sxs-lookup"><span data-stu-id="eb630-141">**Document Library**: OneDrive</span></span>
    - <span data-ttu-id="eb630-142">**Datei**: task-reminders.xlsx *(Über den Dateibrowser ausgewählt)*</span><span class="sxs-lookup"><span data-stu-id="eb630-142">**File**: task-reminders.xlsx *(Chosen through the file browser)*</span></span>
    - <span data-ttu-id="eb630-143">**Skript**: Get People</span><span class="sxs-lookup"><span data-stu-id="eb630-143">**Script**: Get People</span></span>

    :::image type="content" source="../../images/scenario-task-reminders-first-flow-step.png" alt-text="Der Power Automate-Flow, der den ersten Skriptflussschritt ausführen zeigt":::

12. <span data-ttu-id="eb630-145">Als Nächstes muss der Flow jeden Mitarbeiter in dem vom Skript zurückgegebenen Array verarbeiten.</span><span class="sxs-lookup"><span data-stu-id="eb630-145">Next, the flow needs to process each Employee in the array returned by the script.</span></span> <span data-ttu-id="eb630-146">Drücken Sie Den **neuen Schritt,** und wählen Sie Eine adaptive Karte an einen Teams Benutzer sende **aus, und warten Sie auf eine Antwort**.</span><span class="sxs-lookup"><span data-stu-id="eb630-146">Press **New step** and select **Post an Adaptive Card to a Teams user and wait for a response**.</span></span>

13. <span data-ttu-id="eb630-147">Fügen Sie für das Feld **Empfänger** **E-Mails** aus dem dynamischen Inhalt hinzu (die Auswahl enthält das Excel Logo).</span><span class="sxs-lookup"><span data-stu-id="eb630-147">For the **Recipient** field, add **email** from the dynamic content (the selection will have the Excel logo by it).</span></span> <span data-ttu-id="eb630-148">Das Hinzufügen von **E-Mails** bewirkt, dass der Flussschritt von einem **Apply auf jeden** Block umgeben ist.</span><span class="sxs-lookup"><span data-stu-id="eb630-148">Adding **email** causes the flow step to be surrounded by an **Apply to each** block.</span></span> <span data-ttu-id="eb630-149">Das bedeutet, dass das Array von Power Automate iteriert wird.</span><span class="sxs-lookup"><span data-stu-id="eb630-149">That means the array will be iterated over by Power Automate.</span></span>

14. <span data-ttu-id="eb630-150">Für das Senden einer Adaptive Card muss das JSON der Karte als **Nachricht** bereitgestellt werden.</span><span class="sxs-lookup"><span data-stu-id="eb630-150">Sending an Adaptive Card requires the card's JSON to be provided as the **Message**.</span></span> <span data-ttu-id="eb630-151">Sie können den [Adaptive Card Designer](https://adaptivecards.io/designer/) verwenden, um benutzerdefinierte Karten zu erstellen.</span><span class="sxs-lookup"><span data-stu-id="eb630-151">You can use the [Adaptive Card Designer](https://adaptivecards.io/designer/) to create custom cards.</span></span> <span data-ttu-id="eb630-152">Verwenden Sie für dieses Beispiel die folgende JSON.For This-Beispiel.</span><span class="sxs-lookup"><span data-stu-id="eb630-152">For this sample, use the following JSON.</span></span>  

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

15. <span data-ttu-id="eb630-153">Füllen Sie die verbleibenden Felder wie folgt aus:</span><span class="sxs-lookup"><span data-stu-id="eb630-153">Fill out the remaining fields as follows:</span></span>

    - <span data-ttu-id="eb630-154">**Update-Meldung**: Vielen Dank für die Übermittlung Ihres Statusberichts.</span><span class="sxs-lookup"><span data-stu-id="eb630-154">**Update message**: Thank you for submitting your status report.</span></span> <span data-ttu-id="eb630-155">Ihre Antwort wurde erfolgreich zur Kalkulationstabelle hinzugefügt.</span><span class="sxs-lookup"><span data-stu-id="eb630-155">Your response has been successfully added to the spreadsheet.</span></span>
    - <span data-ttu-id="eb630-156">**Sollte Karte aktualisieren**: Ja</span><span class="sxs-lookup"><span data-stu-id="eb630-156">**Should update card**: Yes</span></span>

16. <span data-ttu-id="eb630-157">Drücken Sie unter **Auf jeden** Block anwenden, nach dem Posten **einer adaptiven Karte an einen Teams Benutzer und warten Sie auf eine Antwort,** die Taste **Eine Aktion hinzufügen**.</span><span class="sxs-lookup"><span data-stu-id="eb630-157">In the **Apply to each** block, following the **Post an Adaptive Card to a Teams user and wait for a response**, press **Add an action**.</span></span> <span data-ttu-id="eb630-158">Wählen Sie **Excel Online (Geschäft)** aus.</span><span class="sxs-lookup"><span data-stu-id="eb630-158">Select **Excel Online (Business)**.</span></span> <span data-ttu-id="eb630-159">Wählen Sie unter **Aktionen** die Option **Skript ausführen** aus.</span><span class="sxs-lookup"><span data-stu-id="eb630-159">Under **Actions**, select **Run script**.</span></span> <span data-ttu-id="eb630-160">Geben Sie die folgenden Einträge für den Flow-Schritt an:</span><span class="sxs-lookup"><span data-stu-id="eb630-160">Provide the following entries for the flow step:</span></span>

    - <span data-ttu-id="eb630-161">**Location**: OneDrive for Business</span><span class="sxs-lookup"><span data-stu-id="eb630-161">**Location**: OneDrive for Business</span></span>
    - <span data-ttu-id="eb630-162">**Document Library**: OneDrive</span><span class="sxs-lookup"><span data-stu-id="eb630-162">**Document Library**: OneDrive</span></span>
    - <span data-ttu-id="eb630-163">**Datei**: task-reminders.xlsx *(Über den Dateibrowser ausgewählt)*</span><span class="sxs-lookup"><span data-stu-id="eb630-163">**File**: task-reminders.xlsx *(Chosen through the file browser)*</span></span>
    - <span data-ttu-id="eb630-164">**Skript**: Status speichern</span><span class="sxs-lookup"><span data-stu-id="eb630-164">**Script**: Save Status</span></span>
    - <span data-ttu-id="eb630-165">**absenderEmail**: E-Mail *(dynamischer Inhalt von Excel)*</span><span class="sxs-lookup"><span data-stu-id="eb630-165">**senderEmail**: email *(dynamic content from Excel)*</span></span>
    - <span data-ttu-id="eb630-166">**statusReportResponse**: Antwort *(dynamischer Inhalt aus Teams)*</span><span class="sxs-lookup"><span data-stu-id="eb630-166">**statusReportResponse**: response *(dynamic content from Teams)*</span></span>

    :::image type="content" source="../../images/scenario-task-reminders-last-flow-step.png" alt-text="Der Power Automate-Flow, der den Auf-jeden-Schritt anzeigt":::

17. <span data-ttu-id="eb630-168">Speichern Sie den Fluss.</span><span class="sxs-lookup"><span data-stu-id="eb630-168">Save the flow.</span></span>

## <a name="running-the-flow"></a><span data-ttu-id="eb630-169">Ausführen des Flows</span><span class="sxs-lookup"><span data-stu-id="eb630-169">Running the flow</span></span>

<span data-ttu-id="eb630-170">Um den Flow zu testen, stellen Sie sicher, dass alle Tabellenzeilen mit leerem Status eine E-Mail-Adresse verwenden, die an ein Teams Konto gebunden ist (Sie sollten wahrscheinlich Ihre eigene E-Mail-Adresse beim Testen verwenden).</span><span class="sxs-lookup"><span data-stu-id="eb630-170">To test the flow, make sure any table rows with blank status use an email address tied to a Teams account (you should probably use your own email address while testing).</span></span>

<span data-ttu-id="eb630-171">Sie können entweder **Test** aus dem Flow-Designer auswählen oder den Flow auf der Seite **Meine Flows** ausführen.</span><span class="sxs-lookup"><span data-stu-id="eb630-171">You can either select **Test** from the flow designer, or run the flow from the **My flows** page.</span></span> <span data-ttu-id="eb630-172">Nachdem Sie den Flow gestartet und die Verwendung der erforderlichen Verbindungen akzeptiert haben, sollten Sie eine Adaptive Card von Power Automate bis Teams erhalten.</span><span class="sxs-lookup"><span data-stu-id="eb630-172">After starting the flow and accepting the use of the required connections, you should receive an Adaptive Card from Power Automate through Teams.</span></span> <span data-ttu-id="eb630-173">Sobald Sie das Statusfeld in der Karte ausgefüllt haben, wird der Flow fortgesetzt und die Kalkulationstabelle mit dem von Ihnen angezeigten Status aktualisiert.</span><span class="sxs-lookup"><span data-stu-id="eb630-173">Once you fill out the status field in the card, the flow will continue and update the spreadsheet with the status you provide.</span></span>

### <a name="before-running-the-flow"></a><span data-ttu-id="eb630-174">Vor dem Ausführen des Flows</span><span class="sxs-lookup"><span data-stu-id="eb630-174">Before running the flow</span></span>

:::image type="content" source="../../images/scenario-task-reminders-spreadsheet-before.png" alt-text="Ein Arbeitsblatt mit einem Statusbericht, der einen fehlenden Statuseintrag enthält":::

### <a name="receiving-the-adaptive-card"></a><span data-ttu-id="eb630-176">Empfangen der Adaptive Card</span><span class="sxs-lookup"><span data-stu-id="eb630-176">Receiving the Adaptive Card</span></span>

:::image type="content" source="../../images/scenario-task-reminders-adaptive-card.png" alt-text="Eine adaptive Karte in Teams, die den Mitarbeiter um eine Statusaktualisierung bittet":::

### <a name="after-running-the-flow"></a><span data-ttu-id="eb630-178">Nach dem Ausführen des Flows</span><span class="sxs-lookup"><span data-stu-id="eb630-178">After running the flow</span></span>

:::image type="content" source="../../images/scenario-task-reminders-spreadsheet-after.png" alt-text="Ein Arbeitsblatt mit einem Statusbericht mit einem jetzt ausgefüllten Statusposten":::
