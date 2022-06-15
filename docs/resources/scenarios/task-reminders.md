---
title: 'beispielszenario für Office Skripts: Automatisierte Aufgabenerinnerungen'
description: Ein Beispiel, das Power Automate und adaptive Karten verwendet, automatisiert Aufgabenerinnerungen in einer Projektmanagementtabelle.
ms.date: 06/29/2021
ms.localizationpriority: medium
ms.openlocfilehash: 08f3713210e83162f86d38bc8eb33d76bf8a7288
ms.sourcegitcommit: dd01979d34b3499360d2f79a56f8a8f24f480eed
ms.translationtype: MT
ms.contentlocale: de-DE
ms.lasthandoff: 06/15/2022
ms.locfileid: "66088113"
---
# <a name="office-scripts-sample-scenario-automated-task-reminders"></a>beispielszenario für Office Skripts: Automatisierte Aufgabenerinnerungen

In diesem Szenario verwalten Sie ein Projekt. Sie verwenden ein Excel Arbeitsblatt, um den Status Ihrer Mitarbeiter jeden Monat nachzuverfolgen. Sie müssen Die Benutzer häufig daran erinnern, ihren Status auszufüllen, daher haben Sie sich entschieden, diesen Erinnerungsprozess zu automatisieren.

Sie erstellen einen Power Automate Fluss für Nachrichten mit fehlenden Statusfeldern und wenden deren Antworten auf das Arbeitsblatt an. Dazu entwickeln Sie ein Skriptpaar für die Arbeit mit der Arbeitsmappe. Das erste Skript ruft eine Liste von Personen mit leeren Status ab, und das zweite Skript fügt der rechten Zeile eine Statuszeichenfolge hinzu. Sie verwenden auch [Teams adaptive Karten](/microsoftteams/platform/task-modules-and-cards/what-are-cards), damit Mitarbeiter ihren Status direkt aus der Benachrichtigung eingeben können.

## <a name="scripting-skills-covered"></a>Abgedeckte Skriptfähigkeiten

- Erstellen von Flüssen in Power Automate
- Übergeben von Daten an Skripts
- Zurückgeben von Daten aus Skripts
- Teams adaptive Karten
- Tabellen

## <a name="prerequisites"></a>Voraussetzungen

In diesem Szenario werden [Power Automate](https://flow.microsoft.com) und [Microsoft Teams](https://www.microsoft.com/microsoft-365/microsoft-teams/group-chat-software) verwendet. Sie müssen beide mit dem Konto verknüpft sein, das Sie für die Entwicklung Office Skripts verwenden. Wenn Sie kostenlosen Zugriff auf ein Microsoft Developer-Abonnement erhalten möchten, um mehr über diese Anwendungen zu erfahren und mit diesen zu arbeiten, sollten Sie dem [Microsoft 365-Entwicklerprogramm](https://developer.microsoft.com/microsoft-365/dev-program) beitreten.

## <a name="setup-instructions"></a>Setupanweisungen

1. Laden Sie <a href="task-reminders.xlsx">task-reminders.xlsx</a> in Ihre OneDrive herunter.

1. Öffnen Sie die Arbeitsmappe in Excel im Web.

1. Zunächst benötigen wir ein Skript, um alle Mitarbeiter mit Statusberichten abzurufen, die in der Tabelle fehlen. Wählen Sie auf der Registerkarte **"Automatisieren** " die Option **"Neues Skript** " aus, und fügen Sie das folgende Skript in den Editor ein.

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

1. Speichern Sie das Skript unter dem Namen **"Personen abrufen"**.

1. Als Nächstes benötigen wir ein zweites Skript zum Verarbeiten der Statusberichtskarten und zum Einfügen der neuen Informationen in die Tabelle. Wählen Sie im Aufgabenbereich "Code-Editor" die Option **"Neues Skript** " aus, und fügen Sie das folgende Skript in den Editor ein.

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

1. Speichern Sie das Skript unter dem Namen **"Status speichern"**.

1. Jetzt müssen wir den Fluss erstellen. Öffnen [Sie Power Automate](https://flow.microsoft.com/).

    > [!TIP]
    > Wenn Sie zuvor noch keinen Fluss erstellt haben, lesen Sie bitte unser Lernprogramm ["Mit skripts mit Power Automate](../../tutorials/excel-power-automate-manual.md) beginnen", um die Grundlagen zu erlernen.

1. Erstellen Sie einen neuen **Sofortfluss**.

1. Wählen Sie **"Manuell auslösen"** aus den Optionen und dann " **Erstellen**" aus.

1. Der Fluss muss das Skript " **Personen** abrufen" aufrufen, um alle Mitarbeiter mit leeren Statusfeldern abzurufen. Wählen Sie **"Neuer Schritt**" und dann **Excel Online (Business)** aus. Wählen Sie unter **Aktionen** die Option **Skript ausführen** aus. Geben Sie die folgenden Einträge für den Ablaufschritt an:

    - **Location**: OneDrive for Business
    - **Document Library**: OneDrive
    - **Datei**: task-reminders.xlsx *(Über den Dateibrowser ausgewählt)*
    - **Skript**: Abrufen von Personen

    :::image type="content" source="../../images/scenario-task-reminders-first-flow-step.png" alt-text="Der Power Automate-Fluss mit dem ersten Schritt zum Ausführen des Skriptflusses.":::

1. Als Nächstes muss der Fluss jeden Mitarbeiter in dem vom Skript zurückgegebenen Array verarbeiten. Wählen Sie **"Neuer Schritt**" und dann "**Eine adaptive Karte für einen Teams Benutzer posten" aus, und warten Sie auf eine Antwort**.

1. Fügen Sie für das Feld **"Empfänger**" **E-Mails** aus dem dynamischen Inhalt hinzu (die Auswahl hat das Excel Logo). Durch das Hinzufügen von **E-Mails** wird der Flussschritt von einem **"Für jeden Block übernehmen"** umgeben. Das bedeutet, dass das Array von Power Automate iteriert wird.

1. Zum Senden einer adaptiven Karte muss der [JSON-Code](https://www.w3schools.com/whatis/whatis_json.asp) der Karte als **Nachricht** bereitgestellt werden. Sie können den [Designer für adaptive Karten](https://adaptivecards.io/designer/) verwenden, um benutzerdefinierte Karten zu erstellen. Verwenden Sie für dieses Beispiel den folgenden JSON-Code.  

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

    - **Updatenachricht**: Vielen Dank, dass Sie Ihren Statusbericht übermittelt haben. Ihre Antwort wurde erfolgreich zur Kalkulationstabelle hinzugefügt.
    - **Karte aktualisieren**: Ja

1. Wählen Sie in der Option **"Für jeden Block übernehmen**" nach dem **Posten einer adaptiven Karte an einen Teams Benutzer und warten Sie auf eine Antwort**, und wählen Sie "**Aktion hinzufügen"** aus. Wählen Sie **Excel Online (Business)** aus. Wählen Sie unter **Aktionen** die Option **Skript ausführen** aus. Geben Sie die folgenden Einträge für den Ablaufschritt an:

    - **Location**: OneDrive for Business
    - **Document Library**: OneDrive
    - **Datei**: task-reminders.xlsx *(Über den Dateibrowser ausgewählt)*
    - **Skript**: Speichern des Status
    - **senderEmail**: E-Mail *(dynamischer Inhalt aus Excel)*
    - **statusReportResponse**: Antwort *(dynamischer Inhalt aus Teams)*

    :::image type="content" source="../../images/scenario-task-reminders-last-flow-step.png" alt-text="Der Power Automate Fluss, der den Schritt &quot;Für jeden anwenden&quot; anzeigt.":::

1. Speichern Sie den Fluss.

## <a name="running-the-flow"></a>Ausführen des Flusses

Stellen Sie zum Testen des Flusses sicher, dass alle Tabellenzeilen mit leerem Status eine E-Mail-Adresse verwenden, die an ein Teams-Konto gebunden ist (Sie sollten beim Testen wahrscheinlich Ihre eigene E-Mail-Adresse verwenden). Verwenden Sie die Schaltfläche " **Testen** " auf der Fluss-Editor-Seite, oder führen Sie den Fluss über die Registerkarte **"Meine Flüsse** " aus. Achten Sie darauf, den Zugriff zuzulassen, wenn Sie dazu aufgefordert werden.

Sie sollten eine adaptive Karte von Power Automate bis Teams erhalten. Nachdem Sie das Statusfeld auf der Karte ausgefüllt haben, wird der Fluss fortgesetzt und die Tabelle mit dem von Ihnen bereitgestellten Status aktualisiert.

### <a name="before-running-the-flow"></a>Vor dem Ausführen des Flusses

:::image type="content" source="../../images/scenario-task-reminders-spreadsheet-before.png" alt-text="Ein Arbeitsblatt mit einem Statusbericht, der einen fehlenden Statuseintrag enthält.":::

### <a name="receiving-the-adaptive-card"></a>Empfangen der adaptiven Karte

:::image type="content" source="../../images/scenario-task-reminders-adaptive-card.png" alt-text="Eine adaptive Karte in Teams den Mitarbeiter um eine Statusaktualisierung bitten.":::

### <a name="after-running-the-flow"></a>Nach dem Ausführen des Flusses

:::image type="content" source="../../images/scenario-task-reminders-spreadsheet-after.png" alt-text="Ein Arbeitsblatt mit einem Statusbericht mit einem jetzt ausgefüllten Statuseintrag.":::
