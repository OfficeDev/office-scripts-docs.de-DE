---
title: 'Beispielszenario für Office-Skripts: Automatisierte Aufgabenerinnerungen'
description: Ein Beispiel, das Power Automate und adaptive Karten verwendet, automatisiert Aufgabenerinnerungen in einer Projektverwaltungskalkulationstabelle.
ms.date: 11/30/2020
localization_priority: Normal
ms.openlocfilehash: a229a06e9f1f9118d57dadac8864bbc7eae7315b
ms.sourcegitcommit: 45ffe3dbd2c834b78592ad35928cf8096f5e80bc
ms.translationtype: MT
ms.contentlocale: de-DE
ms.lasthandoff: 04/14/2021
ms.locfileid: "51755154"
---
# <a name="office-scripts-sample-scenario-automated-task-reminders"></a>Beispielszenario für Office-Skripts: Automatisierte Aufgabenerinnerungen

In diesem Szenario verwalten Sie ein Projekt. Sie verwenden ein Excel-Arbeitsblatt, um den Status Ihrer Mitarbeiter jeden Monat nachverfolgt zu haben. Sie müssen die Personen häufig daran erinnern, ihren Status zu füllen, daher haben Sie sich entschieden, diesen Erinnerungsprozess zu automatisieren.

Sie erstellen einen Power Automate-Fluss, um Personen mit fehlenden Statusfeldern zu senden und deren Antworten auf das Arbeitsblatt anzuwenden. Dazu entwickeln Sie ein Skriptpaar, um die Arbeit mit der Arbeitsmappe zu verarbeiten. Das erste Skript ruft eine Liste von Personen mit leeren Status ab, und das zweite Skript fügt der rechten Zeile eine Statuszeichenfolge hinzu. Sie verwenden auch adaptive [Teams-Karten,](/microsoftteams/platform/task-modules-and-cards/what-are-cards) damit Mitarbeiter ihren Status direkt aus der Benachrichtigung eingeben können.

## <a name="scripting-skills-covered"></a>Abgedeckte Skriptkenntnisse

- Erstellen von Flüssen in Power Automate
- Übergeben von Daten an Skripts
- Zurückgeben von Daten aus Skripts
- Adaptive Teams-Karten
- Tabellen

## <a name="prerequisites"></a>Voraussetzungen

In diesem Szenario werden [Power Automate und](https://flow.microsoft.com) Microsoft Teams [verwendet.](https://www.microsoft.com/microsoft-365/microsoft-teams/group-chat-software) Sie müssen beide dem Konto zugeordnet sein, das Sie für die Entwicklung von Office-Skripts verwenden. Um kostenlosen Zugriff auf ein Microsoft Developer-Abonnement zu erhalten, um mehr über diese Anwendungen zu erfahren und mit diesen zu arbeiten, sollten Sie am [Microsoft 365 Developer Program teilnehmen.](https://developer.microsoft.com/microsoft-365/dev-program)

## <a name="setup-instructions"></a>Setupanweisungen

1. Laden <a href="task-reminders.xlsx">task-reminders.xlsx</a> auf Ihr OneDrive herunter.

2. Öffnen Sie die Arbeitsmappe in Excel im Web.

3. Öffnen Sie **auf** der Registerkarte Automatisieren **alle Skripts**.

4. Zunächst benötigen wir ein Skript, um alle Mitarbeiter mit Statusberichten zu erhalten, die in der Tabelle fehlen. Drücken Sie **im Aufgabenbereich Code-Editor** die **Taste Neues Skript,** und fügen Sie das folgende Skript in den Editor ein.

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

6. Als Nächstes benötigen wir ein zweites Skript, um die Statusberichtskarten zu verarbeiten und die neuen Informationen in der Tabelle zu speichern. Drücken Sie **im Aufgabenbereich Code-Editor** die **Taste Neues Skript,** und fügen Sie das folgende Skript in den Editor ein.

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

7. Speichern Sie das Skript mit dem Namen **Save Status**.

8. Jetzt müssen wir den Fluss erstellen. Öffnen [Sie Power Automate](https://flow.microsoft.com/).

    > [!TIP]
    > Wenn Sie noch keinen Fluss erstellt haben, lesen Sie bitte unser Lernprogramm [Start using scripts with Power Automate,](../../tutorials/excel-power-automate-manual.md) um die Grundlagen zu erfahren.

9. Erstellen Sie einen neuen **Instant-Fluss.**

10. Wählen **Sie Manuell einen Fluss aus den** Optionen auslösen aus, und drücken Sie **erstellen**.

11. Der Fluss muss das **Skript Get People aufrufen,** um alle Mitarbeiter mit leeren Statusfeldern zu erhalten. Drücken **Sie den Schritt Neu,** und wählen Sie Excel Online **(Business) aus.** Wählen Sie unter **Actions** die Option **Run script (preview)** aus. Geben Sie die folgenden Einträge für den Flussschritt an:

    - **Location**: OneDrive for Business
    - **Document Library**: OneDrive
    - **Datei**: task-reminders.xlsx *(über den Dateibrowser ausgewählt)*
    - **Skript**: Get People

    :::image type="content" source="../../images/scenario-task-reminders-first-flow-step.png" alt-text="Der Power Automate-Fluss, der den ersten Schritt zum Ausführen des Skriptflusses zeigt.":::

12. Als Nächstes muss der Fluss jeden Mitarbeiter im vom Skript zurückgegebenen Array verarbeiten. Drücken **Sie den Schritt Neu,** und wählen Sie Adaptive Karte an einen Teams-Benutzer posten **aus, und warten Sie auf eine Antwort.**

13. Fügen Sie **für das Feld** Empfänger **E-Mails** aus dem dynamischen Inhalt hinzu (die Auswahl hat das Excel-Logo davon). Das **Hinzufügen von** E-Mails bewirkt, dass der Flussschritt von einem Auf jeden Block anwenden **umgeben** ist. Das bedeutet, dass das Array von Power Automate durch iteriert wird.

14. Für das Senden einer adaptiven Karte muss das JSON der Karte als Nachricht **bereitgestellt werden.** Sie können den [Adaptive Card Designer verwenden,](https://adaptivecards.io/designer/) um benutzerdefinierte Karten zu erstellen. Verwenden Sie für dieses Beispiel den folgenden JSON.  

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

    - **Updatenachricht**: Vielen Dank für die Übermittlung Ihres Statusberichts. Ihre Antwort wurde der Tabelle erfolgreich hinzugefügt.
    - **Sollte Karte aktualisieren:** Ja

16. Drücken Sie **im Abschnitt Auf jeden** Block anwenden, nachdem Sie auf die Aktion Adaptive Karte an einen **Teams-Benutzer** posten und auf eine Antwort **warten, die Aktion hinzufügen.** Wählen **Sie Excel Online (Business) aus.** Wählen Sie unter **Actions** die Option **Run script (preview)** aus. Geben Sie die folgenden Einträge für den Flussschritt an:

    - **Location**: OneDrive for Business
    - **Document Library**: OneDrive
    - **Datei**: task-reminders.xlsx *(über den Dateibrowser ausgewählt)*
    - **Skript**: Save Status
    - **senderEmail**: E-Mail *(dynamischer Inhalt aus Excel)*
    - **statusReportResponse**: response *(dynamischer Inhalt aus Teams)*

    :::image type="content" source="../../images/scenario-task-reminders-last-flow-step.png" alt-text="Der Power Automate-Fluss, der den jeweiligen Schritt zeigt.":::

17. Speichern Sie den Fluss.

## <a name="running-the-flow"></a>Ausführen des Ablaufs

Stellen Sie zum Testen des Ablaufs sicher, dass tabellenzeilen mit leerem Status eine E-Mail-Adresse verwenden, die an ein Teams-Konto gebunden ist (Sie sollten beim Testen wahrscheinlich Ihre eigene E-Mail-Adresse verwenden).

Sie können entweder **test aus** dem Fluss-Designer auswählen oder den Fluss auf der Seite **Meine Flüsse** ausführen. Nachdem Sie den Fluss gestartet und die Verwendung der erforderlichen Verbindungen akzeptiert haben, sollten Sie eine adaptive Karte von Power Automate über Teams erhalten. Sobald Sie das Statusfeld auf der Karte ausfüllen, wird der Fluss fortgesetzt und die Tabelle mit dem von Ihnen angezeigten Status aktualisiert.

### <a name="before-running-the-flow"></a>Vor dem Ausführen des Datenflusses

:::image type="content" source="../../images/scenario-task-reminders-spreadsheet-before.png" alt-text="Ein Arbeitsblatt mit einem Statusbericht mit einem fehlenden Statuseintrag.":::

### <a name="receiving-the-adaptive-card"></a>Empfangen der adaptiven Karte

:::image type="content" source="../../images/scenario-task-reminders-adaptive-card.png" alt-text="Eine adaptive Karte in Teams, die den Mitarbeiter um eine Statusaktualisierung bittet.":::

### <a name="after-running-the-flow"></a>Nach dem Ausführen des Datenflusses

:::image type="content" source="../../images/scenario-task-reminders-spreadsheet-after.png" alt-text="Ein Arbeitsblatt mit einem Statusbericht mit einem jetzt ausgefüllten Statuseintrag.":::
