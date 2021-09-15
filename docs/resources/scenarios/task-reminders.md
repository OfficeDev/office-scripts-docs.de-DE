---
title: 'Office Skript-Beispielszenario: Automatisierte Aufgabenerinnerungen'
description: Ein Beispiel, das Power Automate und adaptive Karten verwendet, automatisiert Aufgabenerinnerungen in einer Projektverwaltungstabelle.
ms.date: 06/29/2021
ms.localizationpriority: medium
ms.openlocfilehash: a2d4701fb7a42953de669c84dbb93104d199d5b8
ms.sourcegitcommit: d3ed4bdeeba805d97c930394e172e8306a0cf484
ms.translationtype: MT
ms.contentlocale: de-DE
ms.lasthandoff: 09/15/2021
ms.locfileid: "59330048"
---
# <a name="office-scripts-sample-scenario-automated-task-reminders"></a>Office Skript-Beispielszenario: Automatisierte Aufgabenerinnerungen

In diesem Szenario verwalten Sie ein Projekt. Sie verwenden ein Excel Arbeitsblatt, um den Status Ihrer Mitarbeiter jeden Monat nachzuverfolgen. Sie müssen die Personen häufig daran erinnern, ihren Status auszufüllen, daher haben Sie sich entschieden, diesen Erinnerungsprozess zu automatisieren.

Sie erstellen einen Power Automate Fluss an Nachrichtenmitarbeiter mit fehlenden Statusfeldern und wenden deren Antworten auf das Arbeitsblatt an. Zu diesem Zweck entwickeln Sie ein Skriptpaar, um die Arbeit mit der Arbeitsmappe zu verarbeiten. Das erste Skript ruft eine Liste von Personen mit leeren Status ab, und das zweite Skript fügt der rechten Zeile eine Statuszeichenfolge hinzu. Sie verwenden auch [Teams adaptive Karten,](/microsoftteams/platform/task-modules-and-cards/what-are-cards) damit Mitarbeiter ihren Status direkt über die Benachrichtigung eingeben können.

## <a name="scripting-skills-covered"></a>Abgedeckte Skriptfähigkeiten

- Erstellen von Flüssen in Power Automate
- Übergeben von Daten an Skripts
- Zurückgeben von Daten aus Skripts
- Teams Adaptive Karten
- Tabellen

## <a name="prerequisites"></a>Voraussetzungen

In diesem Szenario [werden Power Automate](https://flow.microsoft.com) und [Microsoft Teams](https://www.microsoft.com/microsoft-365/microsoft-teams/group-chat-software)verwendet. Sie müssen beide mit dem Konto verknüpft sein, das Sie für die Entwicklung von Office Skripts verwenden. Für den kostenlosen Zugriff auf ein Microsoft Developer-Abonnement, um mehr über diese Anwendungen zu erfahren und mit diesen zu arbeiten, sollten Sie dem [Microsoft 365-Entwicklerprogramm](https://developer.microsoft.com/microsoft-365/dev-program)beitreten.

## <a name="setup-instructions"></a>Setup-Anweisungen

1. Laden Sie <a href="task-reminders.xlsx">task-reminders.xlsx</a> auf Ihre OneDrive herunter.

1. Öffnen Sie die Arbeitsmappe in Excel im Web.

1. Zunächst benötigen wir ein Skript, um alle Mitarbeiter mit Statusberichten zu erhalten, die in der Tabelle fehlen. Wählen Sie auf der Registerkarte **"Automatisieren"** die Option **"Neues Skript"** aus, und fügen Sie das folgende Skript in den Editor ein.

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

1. Speichern Sie das Skript mit dem Namen **"Personen abrufen".**

1. Als Nächstes benötigen wir ein zweites Skript, um die Statusberichtskarten zu verarbeiten und die neuen Informationen in die Tabelle einzufügen. Wählen Sie im Code-Editor-Aufgabenbereich **"Neues Skript"** aus, und fügen Sie das folgende Skript in den Editor ein.

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

1. Speichern Sie das Skript mit dem Namen **"Save Status".**

1. Jetzt müssen wir den Fluss erstellen. Öffnen [Sie Power Automate](https://flow.microsoft.com/).

    > [!TIP]
    > Wenn Sie noch keinen Flow erstellt haben, sehen Sie sich unser Lernprogramm ["Start using scripts with Power Automate"](../../tutorials/excel-power-automate-manual.md) an, um die Grundlagen zu erlernen.

1. Erstellen Sie einen neuen **Sofortablauf.**

1. Wählen Sie manuell einen Fluss aus den Optionen **aus,** und wählen Sie **Erstellen** aus.

1. Der Fluss muss das Skript **"Personen** abrufen" aufrufen, um alle Mitarbeiter mit leeren Statusfeldern abzurufen. Wählen Sie **"Neu" aus,** und wählen Sie dann **Excel Online (Business)** aus. Wählen Sie unter **Aktionen** die Option **Skript ausführen** aus. Geben Sie die folgenden Einträge für den Flussschritt an:

    - **Location**: OneDrive for Business
    - **Document Library**: OneDrive
    - **Datei:** task-reminders.xlsx *(über den Dateibrowser ausgewählt)*
    - **Skript:** Personen abrufen

    :::image type="content" source="../../images/scenario-task-reminders-first-flow-step.png" alt-text="Der Power Automate-Ablauf mit dem ersten Ausführungsskriptflussschritt.":::

1. Als Nächstes muss der Fluss jeden Mitarbeiter in dem array verarbeiten, das vom Skript zurückgegeben wird. Wählen Sie **"Neuer Schritt"** aus, und wählen Sie dann **"Adaptive Karte an einen Teams Benutzer posten" aus, und warten Sie auf eine Antwort.**

1. Fügen Sie für das **Feld "Empfänger"** **E-Mails** aus dem dynamischen Inhalt hinzu (die Auswahl enthält das Excel Logo). Durch das Hinzufügen von **E-Mails** wird der Flussschritt von einem **Apply-Element für jeden** Block umgeben. Das bedeutet, dass das Array von Power Automate durchlaufen wird.

1. Das Senden einer adaptiven Karte erfordert, dass der JSON-Code der Karte als **Nachricht** bereitgestellt wird. Sie können den Designer für [adaptive Karten](https://adaptivecards.io/designer/) verwenden, um benutzerdefinierte Karten zu erstellen. Verwenden Sie für dieses Beispiel den folgenden JSON-Code.  

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

1. Füllen Sie die verbleibenden Felder wie folgt aus:

    - **Nachricht aktualisieren:** Vielen Dank für die Übermittlung Ihres Statusberichts. Ihre Antwort wurde erfolgreich zur Kalkulationstabelle hinzugefügt.
    - **Sollte Karte aktualisieren:** Ja

1. Wählen Sie im Abschnitt **"Auf jeden** Block anwenden" nach dem **Posten einer adaptiven Karte für einen Teams Benutzer und Warten auf eine Antwort** eine Aktion **hinzufügen** aus. Wählen Sie **Excel Online (Business)** aus. Wählen Sie unter **Aktionen** die Option **Skript ausführen** aus. Geben Sie die folgenden Einträge für den Flussschritt an:

    - **Location**: OneDrive for Business
    - **Document Library**: OneDrive
    - **Datei:** task-reminders.xlsx *(über den Dateibrowser ausgewählt)*
    - **Skript:** Save Status
    - **senderEmail**: E-Mail *(dynamischer Inhalt von Excel)*
    - **statusReportResponse:** response *(dynamischer Inhalt aus Teams)*

    :::image type="content" source="../../images/scenario-task-reminders-last-flow-step.png" alt-text="Der Power Automate-Ablauf, in dem die einzelnen Schritte angezeigt werden.":::

1. Speichern Sie den Fluss.

## <a name="running-the-flow"></a>Ausführen des Flusses

Um den Fluss zu testen, stellen Sie sicher, dass alle Tabellenzeilen mit leerem Status eine E-Mail-Adresse verwenden, die an ein Teams Konto gebunden ist (Sie sollten beim Testen wahrscheinlich Ihre eigene E-Mail-Adresse verwenden). Verwenden Sie die Schaltfläche **"Test"** auf der Flow-Editor-Seite, oder führen Sie den Fluss über Ihre Registerkarte **"Meine Flüsse"** aus. Achten Sie darauf, den Zugriff zuzulassen, wenn Sie dazu aufgefordert werden.

Sie sollten eine adaptive Karte von Power Automate über Teams erhalten. Nachdem Sie das Statusfeld in der Karte ausgefüllt haben, wird der Fluss fortgesetzt und die Tabelle mit dem von Ihnen angegebenen Status aktualisiert.

### <a name="before-running-the-flow"></a>Vor dem Ausführen des Flusses

:::image type="content" source="../../images/scenario-task-reminders-spreadsheet-before.png" alt-text="Ein Arbeitsblatt mit einem Statusbericht, der einen fehlenden Statuseintrag enthält.":::

### <a name="receiving-the-adaptive-card"></a>Empfangen der adaptiven Karte

:::image type="content" source="../../images/scenario-task-reminders-adaptive-card.png" alt-text="Eine adaptive Karte in Teams, die den Mitarbeiter um eine Statusaktualisierung bittet.":::

### <a name="after-running-the-flow"></a>Nach dem Ausführen des Flusses

:::image type="content" source="../../images/scenario-task-reminders-spreadsheet-after.png" alt-text="Ein Arbeitsblatt mit einem Statusbericht mit einem jetzt ausgefüllten Statuseintrag.":::
