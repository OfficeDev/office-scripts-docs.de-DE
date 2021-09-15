---
title: Übergeben von Daten zu Skripts in einem automatisch ausgeführten Power Automate-Datenfluss
description: Ein Lernprogramm zum Ausführen von Office-Skripts für Excel im Web mithilfe von Power Automate, wenn E-Mails empfangen und Flussdaten an das Skript übergeben werden.
ms.date: 06/29/2021
ms.localizationpriority: high
ms.openlocfilehash: 333ccfc753da067111ca4dc0c3e59ce9db360e0a
ms.sourcegitcommit: d3ed4bdeeba805d97c930394e172e8306a0cf484
ms.translationtype: HT
ms.contentlocale: de-DE
ms.lasthandoff: 09/15/2021
ms.locfileid: "59326870"
---
# <a name="pass-data-to-scripts-in-an-automatically-run-power-automate-flow"></a>Übergeben von Daten zu Skripts in einem automatisch ausgeführten Power Automate-Datenfluss

In diesem Lernprogramm erfahren Sie, wie Sie ein Office-Skript für Excel im Web mit einem automatisierten [Power Automate](https://flow.microsoft.com)-Workflow verwenden. Das Skript wird jedes Mal, wenn Sie eine E-Mail erhalten, automatisch ausgeführt, um Informationen aus der E-Mail in einer Excel-Arbeitsmappe aufzuzeichnen. Die Möglichkeit, Daten aus anderen Anwendungen in ein Office-Skript zu übertragen, bietet Ihnen ein hohes Maß an Flexibilität und Freiheit für Ihre automatisierten Prozesse.

> [!TIP]
> Wenn Sie mit Office-Skripten noch nicht vertraut sind, empfehlen wir, mit dem [Aufzeichnen, Bearbeiten und Erstellen von Office-Skripten in Excel im Web](excel-tutorial.md)-Lernprogramm zu beginnen. Wenn Sie noch nicht mit Power Automate vertraut sind, empfehlen wir, dass Sie mit dem Lernprogramm [Aufrufen von Skripts aus einem manuellen Power Automate-Datenfluss](excel-power-automate-manual.md) beginnen. [Office-Skripts verwenden TypeScript](../overview/code-editor-environment.md), und dieses Lernprogramm richten sich an Anfänger bis Fortgeschrittene mit JavaScript oder TypeScript. Wenn Sie noch nicht mit JavaScript vertraut sind, empfehlen wir Ihnen, mit dem [Mozilla-JavaScript-Lernprogramm](https://developer.mozilla.org/docs/Web/JavaScript/Guide/Introduction) zu beginnen.

## <a name="prerequisites"></a>Voraussetzungen

[!INCLUDE [Tutorial prerequisites](../includes/power-automate-tutorial-prerequisites.md)]

## <a name="prepare-the-workbook"></a>Vorbereiten der Arbeitsmappe

Power Automate sollte für den Zugriff auf Arbeitsmappenkomponenten keine [relativen Bezüge](../testing/power-automate-troubleshooting.md#avoid-relative-references) wie `Workbook.getActiveWorksheet` verwenden. Deshalb benötigen wir eine Arbeitsmappe und ein Arbeitsblatt mit konsistenten Namen, die Power Automate als Referenz verwenden kann.

1. Erstellen Sie eine neue Arbeitsmappe mit dem Namen **Mein Arbeitsblatt**.

2. Wechseln Sie zur Registerkarte **Automatisieren**, und wählen Sie **Alle Skripts** aus.

3. Wählen Sie **New Script** aus.

4. Ersetzen Sie den vorhandenen Code durch den folgenden Code, und wählen Sie **Ausführen** aus. Dadurch wird die Arbeitsmappe mit konsistenten Namen für Arbeitsblatt, Tabelle und PivotTable eingerichtet.

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

## <a name="create-an-office-script"></a>Erstellen eines Office-Skripts

Jetzt erstellen Sie ein Skript, das Informationen aus einer E-Mail protokolliert. Wir möchten wissen, an welchen Wochentagen wir die meisten E-Mails empfangen, und wie viele eindeutige Absender diese E-Mails senden. Die Arbeitsmappe enthält eine Tabelle mit den Spalten **Date**, **Day of the week**, **Email address** und **Subject**. Unser Arbeitsblatt enthält außerdem eine PivotTable, die die Spalten **Day of the week** und **Email address** verwendet (dabei handelt es sich um die Zeilenhierarchien). Die Anzahl eindeutiger **Themen** entspricht den aggregierten Informationen, die angezeigt werden (die Datenhierarchie). Nachdem die E-Mail-Tabelle aktualisiert wurde, wird auch das Skript aktualisiert.

1. Wählen Sie im Aufgabenbereich Code-Editor die Option **Neues Skript** aus.

2. Der Datenstrom, den wir später im Lernprogramm erstellen, sendet die Skriptinformationen zu jeder empfangenen E-Mail-Nachricht. Das Skript muss diese Eingabe über Parameter in der `main`-Funktion akzeptieren. Ersetzen Sie das Standardskript durch das folgende Skript:

    ```TypeScript
    function main(
      workbook: ExcelScript.Workbook,
      from: string,
      dateReceived: string,
      subject: string) {

    }
    ```

3. Das Skript benötigt Zugriff auf die Tabelle und die PivotTable der Arbeitsmappe. Fügen Sie den folgenden Code dem Skripttext nach dem öffnenden `{` hinzu:

    ```TypeScript
    // Get the email table.
    let emailWorksheet = workbook.getWorksheet("Emails");
    let table = emailWorksheet.getTable("EmailTable");
  
    // Get the PivotTable.
    let pivotTableWorksheet = workbook.getWorksheet("Subjects");
    let pivotTable = pivotTableWorksheet.getPivotTable("Pivot");
    ```

4. Der `dateReceived`-Parameter ist vom Typ `string`. Dies wird jetzt ein [`Date`-Objekt](../develop/javascript-objects.md#date) konvertiert, damit wir den Wochentag ganz einfach abrufen können. Danach muss der Zahlenwert des Tages einer besser lesbaren Version zugeordnet werden. Fügen Sie den folgenden Code am Ende des Skripts vor der schließenden `}` hinzu:

    ```TypeScript
      // Parse the received date string to determine the day of the week.
      let emailDate = new Date(dateReceived);
      let dayName = emailDate.toLocaleDateString("en-US", { weekday: 'long' });
    ```

5. Die `subject`-Zeichenfolge enthält möglicherweise das „RE:“-Antworttag. Wir entfernen den Tag aus der Zeichenfolge, damit E-Mails im selben Thread den gleichen Betreff für die Tabelle aufweisen. Fügen Sie den folgenden Code am Ende des Skripts vor der schließenden `}` hinzu:

    ```TypeScript
    // Remove the reply tag from the email subject to group emails on the same thread.
    let subjectText = subject.replace("Re: ", "");
    subjectText = subjectText.replace("RE: ", "");
    ```

6. Nachdem Sie die E-Mail-Daten nach Wunsch formatiert haben, fügen Sie eine Zeile zur E-Mail-Tabelle hinzu. Fügen Sie den folgenden Code am Ende des Skripts vor der schließenden `}` hinzu:

    ```TypeScript
    // Add the parsed text to the table.
    table.addRow(-1, [dateReceived, dayName, from, subjectText]);
    ```

7. Abschließend stellen Sie sicher, dass die PivotTable aktualisiert wird. Fügen Sie den folgenden Code am Ende des Skripts vor der schließenden `}` hinzu:

    ```TypeScript
    // Refresh the PivotTable to include the new row.
    pivotTable.refresh();
    ```

8. Benennen Sie das Skript in **E-Mail aufzeichnen** um, und wählen Sie **Skript speichern** aus.

Jetzt ist Ihr Skript bereit für einen Power Automate-Workflow. Es sollte wie das folgende Skript aussehen:

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

## <a name="create-an-automated-workflow-with-power-automate"></a>Erstellen eines automatisierten Workflows mit Power Automate

1. Melden Sie sich an der [Power Automate-Website](https://flow.microsoft.com) an.

2. Wählen Sie in dem Menü, das auf der linken Seite des Bildschirms angezeigt wird, **Erstellen** aus. Damit gelangen Sie zur Liste der Möglichkeiten zum Erstellen neuer Workflows.

    :::image type="content" source="../images/power-automate-tutorial-1.png" alt-text="Die Power Automate-Schaltfläche „Erstellen“.":::

3. Wählen Sie im Abschnitt **Start from blank** die Option **Automated flow** aus. Dadurch wird ein Workflow erstellt, der von einem Ereignis ausgelöst wird, z. B. das Empfangen einer E-Mail.

    :::image type="content" source="../images/power-automate-params-tutorial-1.png" alt-text="Die Option für den automatisierten Fluss in Power Automate.":::

4. Geben Sie im daraufhin angezeigten Dialogfenster einen Namen für den Fluss im Textfeld **Flow name** ein. Wählen Sie dann **When a new email arrives** aus der Liste der Optionen unter **Choose your flow's trigger** aus. Möglicherweise müssen Sie mithilfe des Suchfelds nach der Option suchen. Wählen Sie abschließend **Erstellen** aus.

    :::image type="content" source="../images/power-automate-params-tutorial-2.png" alt-text="Ein Teil des Power Automate-Flusses zeigt den ‚Flussnamen‘ und die Optionen zum Auswählen des Triggers für den Fluss. Der Flussname ist ‚E-Mail-Fluss-Datensatz‘, und der Trigger ist die Option ‚Wenn eine neue E-Mail in Outlook eintrifft‘.":::

    > [!NOTE]
    > In diesem Tutorial wird Outlook verwendet. Sie können stattdessen auch Ihren bevorzugten E-Mail-Dienst verwenden, obwohl einige Optionen anders sein können.

5. Wählen Sie **Neuer Schritt** aus.

6. Wählen Sie die Registerkarte **Standard** aus, und wählen Sie dann **Excel Online (Business)** aus.

    :::image type="content" source="../images/power-automate-tutorial-4.png" alt-text="Excel Online (Business)-Option in Power Automate.":::

7. Wählen Sie unter **Aktionen** die Option **Skript ausführen** aus.

    :::image type="content" source="../images/power-automate-tutorial-5.png" alt-text="Aktionsoption „Skript ausführen“ in Power Automate.":::

8. Als Nächstes wählen Sie die Arbeitsmappe, das Skript und die Eingabeargumente für das Skript aus, die im Datenfluss-Schritt verwendet werden sollen. In diesem Lernprogramm verwenden Sie die Arbeitsmappe, die Sie in Ihrem OneDrive erstellt haben. Sie könnten jedoch jede beliebige Arbeitsmappe auf einer OneDrive- oder SharePoint-Website verwenden. Geben Sie die folgenden Einstellungen für den Konnektor **Run script** an:

    - **Location**: OneDrive for Business
    - **Document Library**: OneDrive
    - **File**: MyWorkbook.xlsx *(Ausgewählt über den Dateibrowser)*
    - **Script**: E-Mail aufzeichnen
    - **from**: Von *(dynamischer Inhalt aus Outlook)*
    - **dateReceived**: Uhrzeit des Empfangs *(dynamischer Inhalt aus Outlook)*
    - **subject**: Betreff *(dynamischer Inhalt aus Outlook)*

    *Beachten Sie, dass die Parameter für das Skript nur angezeigt werden, wenn das Skript ausgewählt wurde.*

    :::image type="content" source="../images/power-automate-params-tutorial-3.png" alt-text="Die Aktion zum Ausführen eines Skripts in PowerAutomate zeigt die Optionen an, die erscheinen, nachdem das Skript ausgewählt wurde.":::

9. Wählen Sie **Speichern** aus.

Der Fluss ist nun aktiviert. Er wird das Skript automatisch jedes Mal ausführen, wenn Sie eine E-Mail über Outlook erhalten.

## <a name="manage-the-script-in-power-automate"></a>Verwalten des Skripts in Power Automate

1. Wählen Sie auf der Hauptseite der Power Automate-Seite **My Flows** aus.

    :::image type="content" source="../images/power-automate-tutorial-7.png" alt-text="Die Schaltfläche „Meine Flows“ in Power Automate.":::

2. Wählen Sie Ihren Flow aus. Hier sehen Sie den Ausführungsverlauf. Sie können die Seite aktualisieren, oder Sie können die Schaltfläche **All runs** auswählen, um den Verlauf zu aktualisieren. Der Flow wird kurz nach Empfang einer E-Mail ausgelöst. Testen Sie den Flow durch Senden von E-Mails.

Wenn der Flow ausgelöst und das Skript erfolgreich ausgeführt wird, sollten die Tabelle und die PivotTable der Arbeitsmappe aktualisiert werden.

:::image type="content" source="../images/power-automate-params-tutorial-4.png" alt-text="Ein Arbeitsblatt, auf dem die E-Mail-Tabelle angezeigt wird, nachdem der Fluss dreimal ausgeführt wurde.":::

:::image type="content" source="../images/power-automate-params-tutorial-5.png" alt-text="Ein Arbeitsblatt, auf dem die PivotTable angezeigt wird, nachdem der Fluss dreimal ausgeführt wurde":::

## <a name="next-steps"></a>Nächste Schritte

Schließen Sie das Tutorial[Zurückgeben von Daten aus einem Skript an einen automatisch ausgeführten Power Automate-Flow](excel-power-automate-returns.md) ab. Hier lernen Sie, wie Sie Daten aus einem Skript an den Flow zurückgeben.

Sie können sich auch das Beispielszenario [Automatisierte Aufgabenerinnerungen](../resources/scenarios/task-reminders.md) ansehen, um zu erfahren, wie Sie Office-Skripts und Power Automate mit Teams Adaptive Cards kombinieren können.
