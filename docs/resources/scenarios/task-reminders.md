---
title: 'Beispielszenario für Office-Skripts: automatische Aufgaben Erinnerungen'
description: Ein Beispiel, das Power automatisieren und Adaptive Karten zum Automatisieren von Aufgaben Erinnerungen in einer Project Management-Kalkulationstabelle verwendet.
ms.date: 06/09/2020
localization_priority: Normal
ms.openlocfilehash: f764c37dafdd964e9435d504770d10b1608428b8
ms.sourcegitcommit: aec3c971c6640429f89b6bb99d2c95ea06725599
ms.translationtype: Auto
ms.contentlocale: de-DE
ms.lasthandoff: 06/25/2020
ms.locfileid: "44878910"
---
# <a name="office-scripts-sample-scenario-automated-task-reminders"></a>Beispielszenario für Office-Skripts: automatische Aufgaben Erinnerungen

In diesem Szenario werden Sie ein Projekt verwalten. Sie verwenden ein Excel-Arbeitsblatt, um den Status Ihrer Mitarbeiter monatlich nachzuverfolgen. Sie müssen die Benutzer häufig darauf hinweisen, dass Sie Ihren Status ausfüllen, sodass Sie sich entschieden haben, diesen erinnerungsprozess zu automatisieren.

Sie erstellen einen Power-Automatisierungs Fluss für Nachrichten Personen mit fehlenden Statusfeldern und wenden ihre Antworten auf das Arbeitsblatt an. Dazu entwickeln Sie ein Skript paar zur Verarbeitung der Arbeitsmappe. Das erste Skript ruft eine Liste von Personen mit leeren Status ab, und das zweite Skript fügt der rechten Zeile eine Statuszeichenfolge hinzu. Sie verwenden auch [Adaptive Teams-Karten](/microsoftteams/platform/task-modules-and-cards/what-are-cards) , damit Mitarbeiter Ihren Status direkt aus der Benachrichtigung eingeben können.

## <a name="scripting-skills-covered"></a>Abgedeckte Skript Fertigkeiten

- Erstellen von Flows in Power Automation
- Übergeben von Daten an Skripts
- Zurückgeben von Daten aus Skripts
- Adaptive Teams-Karten
- Tabellen

## <a name="prerequisites"></a>Voraussetzungen

Dieses Szenario verwendet [Power Automation](https://flow.microsoft.com) und [Microsoft Teams](https://www.microsoft.com/microsoft-365/microsoft-teams/group-chat-software). Sie müssen beide mit dem Konto verbunden sein, das Sie für die Entwicklung von Office-Skripts verwenden. Für den kostenlosen Zugriff auf ein Microsoft Developer-Abonnement, um mehr über diese Anwendungen zu erfahren und mit diesen zu arbeiten, sollten Sie sich an das [Microsoft 365 Developer Program](https://developer.microsoft.com/microsoft-365/dev-program)anschließen.

## <a name="setup-instructions"></a>Setup Anweisungen

1. Laden Sie <a href="task-reminders.xlsx">task-reminders.xlsx</a> auf Ihre OneDrive herunter.

2. Öffnen Sie die Arbeitsmappe in Excel im Internet.

3. Öffnen Sie auf der Registerkarte **automatisieren** den **Code-Editor**.

4. Zunächst benötigen wir ein Skript, um alle Mitarbeiter mit Statusberichten abzurufen, die in der Kalkulationstabelle fehlen. Klicken Sie im Aufgabenbereich **Code-Editor** auf **Neues Skript** , und fügen Sie das folgende Skript in den Editor ein.

    ```typescript
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
          people.push({ name: row[NAME_INDEX], email: row[EMAIL_INDEX] });
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

6. Als nächstes benötigen wir ein zweites Skript, um die Statusberichts Karten zu verarbeiten und die neuen Informationen in das Arbeitsblatt einzufügen. Klicken Sie im Aufgabenbereich **Code-Editor** auf **Neues Skript** , und fügen Sie das folgende Skript in den Editor ein.

    ```typescript
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

7. Speichern Sie das Skript mit dem Namen " **Save Status**".

8. Jetzt müssen wir den Fluss erstellen. Öffnen Sie [Power automatisieren](https://flow.microsoft.com/).

    > [!TIP]
    > Wenn Sie noch keinen Flow erstellt haben, lesen Sie bitte unser Tutorial [Start using scripts with Power automatisieren](../../tutorials/excel-power-automate-manual.md) , um die Grundlagen zu erfahren.

9. Erstellen Sie einen neuen **sofort Ablauf**.

10. Wählen Sie einen Fluss aus den Optionen **manuell auslösen** aus, und klicken Sie dann auf **Erstellen**.

11. Der Fluss muss das **Get People** -Skript aufrufen, um alle Mitarbeiter mit leeren Statusfeldern abzurufen. Klicken Sie auf **New Step** , und wählen Sie **Excel Online (Business)** aus. Wählen Sie unter **Aktionen**die Option **Skript ausführen (Vorschau)** aus. Geben Sie für den Ablauf Schritt folgende Einträge an:

    - **Speicherort**: OneDrive für Unternehmen
    - **Dokumentbibliothek**: OneDrive
    - **Datei**: task-reminders.xlsx
    - **Skript**: Get People

    ![Der erste Lauf Skriptfluss Schritt.](../../images/scenario-task-reminders-first-flow-step.png)

12. Im nächsten Schritt muss der Fluss jeden Mitarbeiter in dem vom Skript zurückgegebenen Array verarbeiten. Klicken Sie auf **New Step** , und wählen Sie **in einem Teams-Benutzer eine Adaptive Karte bereit**stellen aus, und warten Sie auf eine Antwort.

13. Fügen Sie **Recipient** für das Feld Empfänger **e-Mail-** Nachweise aus dem dynamischen Inhalt hinzu (die Auswahl wird das Excel-Logo haben). Das Hinzufügen von **e-Mail** bewirkt, dass der Fluss Schritt von einem für **jeden Block angewendet** wird. Das bedeutet, dass das Array von Power Automation iteriert wird.

14. Für das Senden einer adaptiven Karte muss die JSON der Karte als **Nachricht**bereitgestellt werden. Sie können den [Adaptive Card-Designer](https://adaptivecards.io/designer/) verwenden, um benutzerdefinierte Karten zu erstellen. Verwenden Sie für dieses Beispiel die folgende JSON.  

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

15. Füllen Sie die restlichen Felder wie folgt aus:

    - **Update Nachricht**: Vielen Dank, dass Sie Ihren Statusbericht übermitteln. Ihre Antwort wurde dem Arbeitsblatt erfolgreich hinzugefügt.
    - **Sollte Update Card**: Ja

16. Klicken Sie im Feld **auf jeden Block anwenden** , indem Sie eine **Adaptive Karte an einen Microsoft Teams-Benutzer senden und auf eine Antwort warten**auf **Aktion hinzufügen**. Wählen Sie **Excel Online (Business)** aus. Wählen Sie unter **Aktionen**die Option **Skript ausführen (Vorschau)** aus. Geben Sie für den Ablauf Schritt folgende Einträge an:

    - **Speicherort**: OneDrive für Unternehmen
    - **Dokumentbibliothek**: OneDrive
    - **Datei**: task-reminders.xlsx
    - **Skript**: Save Status
    - **senderEmail**: e-Mail *(dynamischer Inhalt aus Excel)*
    - **statusReportResponse**: Antwort *(dynamischer Inhalt aus Teams)*

    ![Der Ablauf Schritt für jeden Fluss.](../../images/scenario-task-reminders-last-flow-step.png)

17. Speichern Sie den Fluss.

## <a name="running-the-flow"></a>Ablauf des Flusses

Um den Fluss zu testen, stellen Sie sicher, dass Tabellenzeilen mit leerem Status eine e-Mail-Adresse verwenden, die mit einem Microsoft Teams-Konto verknüpft ist (Sie sollten beim Testen wahrscheinlich ihre eigene e-Mail-Adresse verwenden).

Sie können entweder **Test** aus dem Flow Designer auswählen oder den Fluss auf der Seite **meine Flows** ausführen. Nachdem Sie den Fluss gestartet und die Verwendung der erforderlichen Verbindungen akzeptiert haben, sollten Sie eine Adaptive Karte von Power Automation über Teams erhalten. Nachdem Sie das Feld Status in der Karte ausgefüllt haben, wird der Fluss fortgesetzt und das Arbeitsblatt mit dem von Ihnen bereitgestellten Status aktualisiert.

### <a name="before-running-the-flow"></a>Vor dem Ausführen des Flusses

![Ein Arbeitsblatt mit einem Statusbericht, der einen fehlenden Statuseintrag enthält.](../../images/scenario-task-reminders-spreadsheet-before.png)

### <a name="receiving-the-adaptive-card"></a>Empfangen der adaptiven Karte

![Eine Adaptive Karte in Microsoft Teams, in der der Mitarbeiter um eine Statusaktualisierung ersucht wird.](../../images/scenario-task-reminders-adaptive-card.png)

### <a name="after-running-the-flow"></a>Nach Ausführung des Flusses

![Ein Arbeitsblatt mit einem Statusbericht mit einem jetzt ausgefüllten Statuseintrag.](../../images/scenario-task-reminders-spreadsheet-after.png)
