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
# <a name="office-scripts-sample-scenario-automated-task-reminders"></a>Office Skripts-Beispielszenario: Automatisierte Aufgabenerinnerungen

In diesem Szenario verwalten Sie ein Projekt. Sie verwenden einen Excel Arbeitsblatt, um den Status Ihrer Mitarbeiter jeden Monat nachzuverfolgen. Sie müssen die Benutzer oft daran erinnern, ihren Status auszufüllen, daher haben Sie sich entschieden, diesen Erinnerungsprozess zu automatisieren.

Sie erstellen einen Power Automate Fluss, um Personen mit fehlenden Statusfeldern zu nachrichten und ihre Antworten auf die Kalkulationstabelle anzuwenden. Dazu entwickeln Sie ein Skriptpaar, um die Arbeit mit der Arbeitsmappe zu verarbeiten. Das erste Skript ruft eine Liste von Personen mit leeren Status ab, und das zweite Skript fügt der rechten Zeile eine Statuszeichenfolge hinzu. Sie nutzen auch [Teams Adaptive Cards,](/microsoftteams/platform/task-modules-and-cards/what-are-cards) damit Mitarbeiter ihren Status direkt aus der Benachrichtigung eingeben.

## <a name="scripting-skills-covered"></a>Scripting-Fähigkeiten abgedeckt

- Erstellen von Flows in Power Automate
- Übergeben von Daten an Skripts
- Zurückgeben von Daten aus Skripts
- Teams Adaptive Karten
- Tabellen

## <a name="prerequisites"></a>Voraussetzungen

In diesem Szenario werden [Power Automate](https://flow.microsoft.com) und [Microsoft Teams verwendet.](https://www.microsoft.com/microsoft-365/microsoft-teams/group-chat-software) Sie müssen beide dem Konto zugeordnet sein, das Sie für die Entwicklung Office Skripts verwenden. Für den kostenlosen Zugriff auf ein Microsoft Developer-Abonnement, um mehr über diese Anwendungen zu erfahren und mit ihnen zu arbeiten, sollten Sie dem [Microsoft 365 Developer Program](https://developer.microsoft.com/microsoft-365/dev-program)beitreten.

## <a name="setup-instructions"></a>Setup-Anweisungen

1. Laden Sie <a href="task-reminders.xlsx">task-reminders.xlsx</a> auf Ihre OneDrive herunter.

2. Öffnen Sie die Arbeitsmappe in Excel im Web.

3. Öffnen Sie unter der Registerkarte **Automatisieren** **alle Skripts**.

4. Zuerst benötigen wir ein Skript, um alle Mitarbeiter mit Statusberichten zu erhalten, die in der Kalkulationstabelle fehlen. Drücken Sie im Aufgabenbereich **Code-Editor** auf **Neues Skript,** und fügen Sie das folgende Skript in den Editor ein.

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

5. Speichern Sie das Skript mit dem Namen **Get People**.

6. Als Nächstes benötigen wir ein zweites Skript, um die Statusberichtskarten zu verarbeiten und die neuen Informationen in die Kalkulationstabelle zu setzen. Drücken Sie im Aufgabenbereich **Code-Editor** auf **Neues Skript,** und fügen Sie das folgende Skript in den Editor ein.

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

7. Speichern Sie das Skript mit dem Namen **Status speichern**.

8. Jetzt müssen wir den Fluss schaffen. Öffnen [Sie Power Automate](https://flow.microsoft.com/).

    > [!TIP]
    > Wenn Sie noch keinen Flow erstellt haben, schauen Sie sich bitte unser Tutorial [Beginnen Sie mit Skripten mit Power Automate,](../../tutorials/excel-power-automate-manual.md) um die Grundlagen zu lernen.

9. Erstellen Sie einen neuen **Instant-Flow**.

10. Wählen Sie **Manuell auslösen einen Fluss** aus den Optionen und drücken Sie **Erstellen**.

11. Der Flow muss das **Skript "Personen abrufen"** aufrufen, um alle Mitarbeiter mit leeren Statusfeldern zu erhalten. Drücken Sie **Den neuen Schritt,** und wählen Sie **Excel Online (Geschäft)** aus. Wählen Sie unter **Aktionen** die Option **Skript ausführen** aus. Geben Sie die folgenden Einträge für den Flow-Schritt an:

    - **Location**: OneDrive for Business
    - **Document Library**: OneDrive
    - **Datei**: task-reminders.xlsx *(Über den Dateibrowser ausgewählt)*
    - **Skript**: Get People

    :::image type="content" source="../../images/scenario-task-reminders-first-flow-step.png" alt-text="Der Power Automate-Flow, der den ersten Skriptflussschritt ausführen zeigt":::

12. Als Nächstes muss der Flow jeden Mitarbeiter in dem vom Skript zurückgegebenen Array verarbeiten. Drücken Sie Den **neuen Schritt,** und wählen Sie Eine adaptive Karte an einen Teams Benutzer sende **aus, und warten Sie auf eine Antwort**.

13. Fügen Sie für das Feld **Empfänger** **E-Mails** aus dem dynamischen Inhalt hinzu (die Auswahl enthält das Excel Logo). Das Hinzufügen von **E-Mails** bewirkt, dass der Flussschritt von einem **Apply auf jeden** Block umgeben ist. Das bedeutet, dass das Array von Power Automate iteriert wird.

14. Für das Senden einer Adaptive Card muss das JSON der Karte als **Nachricht** bereitgestellt werden. Sie können den [Adaptive Card Designer](https://adaptivecards.io/designer/) verwenden, um benutzerdefinierte Karten zu erstellen. Verwenden Sie für dieses Beispiel die folgende JSON.For This-Beispiel.  

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

15. Füllen Sie die verbleibenden Felder wie folgt aus:

    - **Update-Meldung**: Vielen Dank für die Übermittlung Ihres Statusberichts. Ihre Antwort wurde erfolgreich zur Kalkulationstabelle hinzugefügt.
    - **Sollte Karte aktualisieren**: Ja

16. Drücken Sie unter **Auf jeden** Block anwenden, nach dem Posten **einer adaptiven Karte an einen Teams Benutzer und warten Sie auf eine Antwort,** die Taste **Eine Aktion hinzufügen**. Wählen Sie **Excel Online (Geschäft)** aus. Wählen Sie unter **Aktionen** die Option **Skript ausführen** aus. Geben Sie die folgenden Einträge für den Flow-Schritt an:

    - **Location**: OneDrive for Business
    - **Document Library**: OneDrive
    - **Datei**: task-reminders.xlsx *(Über den Dateibrowser ausgewählt)*
    - **Skript**: Status speichern
    - **absenderEmail**: E-Mail *(dynamischer Inhalt von Excel)*
    - **statusReportResponse**: Antwort *(dynamischer Inhalt aus Teams)*

    :::image type="content" source="../../images/scenario-task-reminders-last-flow-step.png" alt-text="Der Power Automate-Flow, der den Auf-jeden-Schritt anzeigt":::

17. Speichern Sie den Fluss.

## <a name="running-the-flow"></a>Ausführen des Flows

Um den Flow zu testen, stellen Sie sicher, dass alle Tabellenzeilen mit leerem Status eine E-Mail-Adresse verwenden, die an ein Teams Konto gebunden ist (Sie sollten wahrscheinlich Ihre eigene E-Mail-Adresse beim Testen verwenden).

Sie können entweder **Test** aus dem Flow-Designer auswählen oder den Flow auf der Seite **Meine Flows** ausführen. Nachdem Sie den Flow gestartet und die Verwendung der erforderlichen Verbindungen akzeptiert haben, sollten Sie eine Adaptive Card von Power Automate bis Teams erhalten. Sobald Sie das Statusfeld in der Karte ausgefüllt haben, wird der Flow fortgesetzt und die Kalkulationstabelle mit dem von Ihnen angezeigten Status aktualisiert.

### <a name="before-running-the-flow"></a>Vor dem Ausführen des Flows

:::image type="content" source="../../images/scenario-task-reminders-spreadsheet-before.png" alt-text="Ein Arbeitsblatt mit einem Statusbericht, der einen fehlenden Statuseintrag enthält":::

### <a name="receiving-the-adaptive-card"></a>Empfangen der Adaptive Card

:::image type="content" source="../../images/scenario-task-reminders-adaptive-card.png" alt-text="Eine adaptive Karte in Teams, die den Mitarbeiter um eine Statusaktualisierung bittet":::

### <a name="after-running-the-flow"></a>Nach dem Ausführen des Flows

:::image type="content" source="../../images/scenario-task-reminders-spreadsheet-after.png" alt-text="Ein Arbeitsblatt mit einem Statusbericht mit einem jetzt ausgefüllten Statusposten":::
