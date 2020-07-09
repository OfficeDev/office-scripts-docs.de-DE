---
title: Automatisches Ausführen von Skripts mit automatisiertem Power-Automatisierungs Fluss
description: Ein Lernprogramm zum Ausführen von Office-Skripts für Excel im Internet durch Power Automation mithilfe eines automatischen externen Triggers (empfangen von e-Mails über Outlook).
ms.date: 07/01/2020
localization_priority: Priority
ms.openlocfilehash: fc98fb36fd5a8c5ef10bc3b767d6f5add0306246
ms.sourcegitcommit: edf58aed3cd38f57e5e7227465a1ef5515e15703
ms.translationtype: MT
ms.contentlocale: de-DE
ms.lasthandoff: 07/08/2020
ms.locfileid: "45081628"
---
# <a name="automatically-run-scripts-with-automated-power-automate-flows-preview"></a>Automatisches Ausführen von Skripts mit automatisierten Power-Automatisierungs Flüssen (Vorschau)

In diesem Lernprogramm erfahren Sie, wie Sie ein Office-Skript für Excel im Internet mit einem automatisierten [Power-Automatisierungs](https://flow.microsoft.com) Workflow verwenden. Ihr Skript wird automatisch jedes Mal ausgeführt, wenn Sie eine e-Mail erhalten, und es werden Informationen aus der e-Mail in einer Excel-Arbeitsmappe aufgezeichnet.

## <a name="prerequisites"></a>Voraussetzungen

[!INCLUDE [Tutorial prerequisites](../includes/tutorial-prerequisites.md)]

> [!IMPORTANT]
> In diesem Lernprogramm wird davon ausgegangen, dass Sie das Tutorial [Ausführen von Office-Skripts in Excel im Internet mit Power automatisieren](excel-power-automate-manual.md) abgeschlossen haben.

## <a name="prepare-the-workbook"></a>Vorbereiten der Arbeitsmappe

Power Automation kann keine [relativen Verweise](../develop/power-automate-integration.md#avoid-using-relative-references) wie `Workbook.getActiveWorksheet` den Zugriff auf Arbeitsmappen-Komponenten verwenden. Daher benötigen wir eine Arbeitsmappe und ein Arbeitsblatt mit konsistenten Namen für Power Automation to Reference.

1. Erstellen Sie eine neue Arbeitsmappe mit dem Namen **myworkbook**.

2. Wechseln Sie zur Registerkarte **automatisieren** , und wählen Sie **Code-Editor**aus.

3. Wählen Sie **Neues Skript**aus.

4. Ersetzen Sie den vorhandenen Code durch das folgende Skript, und drücken Sie **Run**. Dadurch wird die Arbeitsmappe mit konsistenten Tabellen-, Tabellen-und PivotTable-Namen eingerichtet.

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
      let pivotWorksheet = workbook.addWorksheet("SubjectPivot");
      let newPivotTable = workbook.addPivotTable("Pivot", "EmailTable", pivotWorksheet.getRange("A3:C20"));

      // Setup the pivot hierarchies
      newPivotTable.addRowHierarchy(newPivotTable.getHierarchy("Day of the week"));
      newPivotTable.addRowHierarchy(newPivotTable.getHierarchy("Email address"));
      newPivotTable.addDataHierarchy(newPivotTable.getHierarchy("Subject"));
    }
    ```

## <a name="create-an-office-script-for-your-automated-workflow"></a>Erstellen eines Office-Skripts für Ihren automatisierten Workflow

Lassen Sie uns ein Skript erstellen, das Informationen aus einer e-Mail protokolliert. Wir möchten wissen, wie an welchen Wochentagen die meisten e-Mails empfangen werden und wie viele eindeutige Absender diese e-Mail senden. Unsere Arbeitsmappe enthält eine Tabelle mit **Datum**, **Wochentag**, **e-Mail-Adresse**und **Betreff** -Spalten. Unser Arbeitsblatt verfügt auch über eine PivotTable, die am **Tag der Woche** und der **e-Mail-Adresse** (Dies sind die Zeilen Hierarchien) pivotiert. Die Anzahl der eindeutigen **Subjekte** ist die aggregierte Information, die angezeigt wird (die Datenhierarchie). Unser Skript aktualisiert diese PivotTable nach dem Aktualisieren der e-Mail-Tabelle.

1. Wählen Sie im **Code-Editor**die Option **Neues Skript**aus.

2. Der Ablauf, den wir später im Lernprogramm erstellen werden, sendet unsere Skript Informationen zu jeder empfangenen e-Mail. Das Skript muss diese Eingabe über Parameter in der `main` -Funktion akzeptieren. Ersetzen Sie das Standardskript durch das folgende Skript:

    ```TypeScript
    function main(
      workbook: ExcelScript.Workbook,
      from: string,
      dateReceived: string,
      subject: string) {

    }
    ```

3. Das Skript benötigt Zugriff auf die Tabelle und die PivotTable der Arbeitsmappe. Fügen Sie nach dem Öffnen den folgenden Code zum Text des Skripts hinzu `{` :

    ```TypeScript
    // Get the email table.
    let emailWorksheet = workbook.getWorksheet("Emails");
    let table = emailWorksheet.getTable("EmailTable");
  
    // Get the PivotTable.
    let pivotTableWorksheet = workbook.getWorksheet("SubjectPivot");
    let pivotTable = pivotTableWorksheet.getPivotTable("Pivot");
    ```

4. Der- `dateReceived` Parameter ist vom Typ `string` . Let es Convert, dass ein [ `Date` Objekt](../develop/javascript-objects.md#date) , damit wir ganz einfach den Tag der Woche erhalten können. Anschließend müssen wir den Zahlenwert des Tages einer besser lesbaren Version zuordnen. Fügen Sie den folgenden Code am Ende Ihres Skripts vor dem Schließen hinzu `}` :

    ```TypeScript
    // Parse the received date string.
    let date = new Date(dateReceived);

    // Convert number representing the day of the week into the name of the day.
    let dayText : string;
    switch (date.getDay()) {
      case 0:
        dayText = "Sunday";
        break;
      case 1:
        dayText = "Monday";
        break;
      case 2:
        dayText = "Tuesday";
        break;
      case 3:
        dayText = "Wednesday";
        break;
      case 4:
        dayText = "Thursday";
        break;
      case 5:
        dayText = "Friday";
        break;
      default:
        dayText = "Saturday";
        break;
    }
    ```

5. Die `subject` Zeichenfolge kann das Antwort-Tag "Re:" enthalten. Lassen Sie uns diese aus der Zeichenfolge entfernen, damit e-Mails im gleichen Thread denselben Betreff für die Tabelle haben. Fügen Sie den folgenden Code am Ende Ihres Skripts vor dem Schließen hinzu `}` :

    ```TypeScript
    // Remove the reply tag from the email subject to group emails on the same thread.
    let subjectText = subject.replace("Re: ", "");
    subjectText = subjectText.replace("RE: ", "");
    ```

6. Nachdem die e-Mail-Daten nach unserem Geschmack formatiert wurden, fügen wir der e-Mail-Tabelle eine Zeile hinzu. Fügen Sie den folgenden Code am Ende Ihres Skripts vor dem Schließen hinzu `}` :

    ```TypeScript
    // Add the parsed text to the table.
    table.addRow(-1, [dateReceived, dayText, from, subjectText]);
    ```

7. Lassen Sie uns schließlich sicherstellen, dass die PivotTable aktualisiert wird. Fügen Sie den folgenden Code am Ende Ihres Skripts vor dem Schließen hinzu `}` :

    ```TypeScript
    // Refresh the PivotTable to include the new row.
    pivotTable.refresh();
    ```

8. Benennen Sie die **e-Mail-Skriptaufzeichnung** um, und drücken Sie **Skript speichern**.

Ihr Skript ist jetzt für einen Power automatisieren-Workflow verfügbar. Es sollte wie das folgende Skript aussehen:

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
  let pivotTableWorksheet = workbook.getWorksheet("Pivot");
  let pivotTable = pivotTableWorksheet.getPivotTable("SubjectPivot");

  // Parse the received date string.
  let date = new Date(dateReceived);

  // Convert number representing the day of the week into the name of the day.
  let dayText: string;
  switch (date.getDay()) {
    case 0:
      dayText = "Sunday";
      break;
    case 1:
      dayText = "Monday";
      break;
    case 2:
      dayText = "Tuesday";
      break;
    case 3:
      dayText = "Wednesday";
      break;
    case 4:
      dayText = "Thursday";
      break;
    case 5:
      dayText = "Friday";
      break;
    default:
      dayText = "Saturday";
      break;
  }

  // Remove the reply tag from the email subject to group emails on the same thread.
  let subjectText = subject.replace("Re: ", "");
  subjectText = subjectText.replace("RE: ", "");

  // Add the parsed text to the table.
  table.addRow(-1, [dateReceived, dayText, from, subjectText]);

  // Refresh the PivotTable to include the new row.
  pivotTable.refresh();
}
```

## <a name="create-an-automated-workflow-with-power-automate"></a>Erstellen eines automatisierten Workflows mit Power Automation

1. Melden Sie sich bei der [Power Automation Preview-Website](https://flow.microsoft.com)an.

2. Klicken Sie im Menü, das auf der linken Seite des Bildschirms angezeigt wird, auf **Erstellen**. Damit gelangen Sie zur Liste der Möglichkeiten zum Erstellen neuer Workflows.

    ![Die Schaltfläche erstellen in Power automatisieren.](../images/power-automate-tutorial-1.png)

3. Wählen Sie im Abschnitt **Anfang von leer** den Eintrag **automatischer Fluss**aus. Dadurch wird ein Workflow erstellt, der durch ein Ereignis ausgelöst wird, beispielsweise das Empfangen einer e-Mail.

    ![Die automatische Fluss Option in Power Automation.](../images/power-automate-params-tutorial-1.png)

4. Geben Sie im angezeigten Dialogfeld einen Namen für den Fluss in das Textfeld **Fluss Name** ein. Wählen Sie dann, **Wenn eine neue e-Mail** aus der Liste der Optionen unter **Choose your Flow es Trigger**kommt. Möglicherweise müssen Sie mithilfe des Felds Suchen nach der Option suchen. Klicken Sie abschließend auf **Erstellen**.

    ![Teil des Fensters zum Erstellen eines automatisierten Flows in Power Automation, das die Option "neue e-Mail ankommt" anzeigt.](../images/power-automate-params-tutorial-2.png)

    > [!NOTE]
    > In diesem Lernprogramm wird Outlook verwendet. Fühlen Sie sich frei, stattdessen Ihren bevorzugten e-Mail-Dienst zu verwenden, obwohl einige Optionen unterschiedlich sein können.

5. Klicken Sie auf **New Step**.

6. Wählen Sie die Registerkarte **Standard** aus, und wählen Sie dann **Excel Online (Business)** aus.

    ![Die Power-Automatisierungsoption für Excel Online (Business).](../images/power-automate-tutorial-4.png)

7. Wählen Sie unter **Aktionen**die Option **Skript ausführen (Vorschau)** aus.

    ![Die Power automatisieren-Aktionsoption für Run Script (Preview).](../images/power-automate-tutorial-5.png)

8. Geben Sie die folgenden Einstellungen für den **Run Script** Connector an:

    - **Speicherort**: OneDrive für Unternehmen
    - **Dokumentbibliothek**: OneDrive
    - **Datei**: MyWorkbook.xlsx
    - **Skript**: Aufzeichnen von e-Mails
    - **von**: from *(dynamischer Inhalt aus Outlook)*
    - **dateReceived**: Empfangszeit *(dynamischer Inhalt aus Outlook)*
    - **Betreff**: Subject *(dynamischer Inhalt aus Outlook)*

    *Beachten Sie, dass die Parameter für das Skript nur dann angezeigt werden, wenn das Skript ausgewählt ist.*

    ![Die Power automatisieren-Aktionsoption für Run Script (Preview).](../images/power-automate-params-tutorial-3.png)

9. Klicken Sie auf **Speichern**.

Ihr Flow ist jetzt aktiviert. Jedes Mal, wenn Sie eine e-Mail über Outlook erhalten, wird das Skript automatisch ausgeführt.

## <a name="manage-the-script-in-power-automate"></a>Verwalten des Skripts in Power Automation

1. Wählen Sie auf der Seite Main Power automatisieren die Option **meine Flows**aus.

    ![Die Schaltfläche "meine Flüsse" in Power automatisieren.](../images/power-automate-tutorial-7.png)

2. Wählen Sie den Fluss aus. Hier können Sie den Ausführungsverlauf anzeigen. Sie können die Seite aktualisieren oder die Schaltfläche **alle Läufe** aktualisieren drücken, um den Verlauf zu aktualisieren. Der Fluss wird kurz nach dem Empfang einer e-Mail ausgelöst. Testen Sie den Ablauf durch Senden von e-Mails.

Wenn der Fluss ausgelöst wird und Ihr Skript erfolgreich ausgeführt wird, sollten Sie die Tabelle und das PivotTable-Update der Arbeitsmappe sehen.

![Die e-Mail-Tabelle, nachdem der Datenfluss ein paar Mal ausgeführt wurde.](../images/power-automate-params-tutorial-4.png)

![Die PivotTable, nachdem der Fluss ein paar Mal ausgeführt wurde.](../images/power-automate-params-tutorial-5.png)

## <a name="next-steps"></a>Nächste Schritte

Besuchen Sie [Office-Skripts mit Power Automation ausführen](../develop/power-automate-integration.md) , um mehr über das Verbinden von Office-Skripts mit Power Automation zu erfahren.

Sie können auch das [Beispielszenario für automatisierte Aufgaben Erinnerungen](../resources/scenarios/task-reminders.md) lesen, um zu erfahren, wie Sie Office-Skripts und Power Automation mit Adaptive Teams-Karten kombinieren.
