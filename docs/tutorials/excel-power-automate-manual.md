---
title: Aufrufen von Skripts aus einem manuellen Power Automate-Datenfluss
description: Ein Lernprogramm zur Verwendung von Office-Skripts in Power Automate durch einen manuellen Auslöser.
ms.date: 11/30/2020
localization_priority: Priority
ms.openlocfilehash: 831812f5ead549ee3ea3b8c643fc16d5467edbe8
ms.sourcegitcommit: af487756dffea0f8f0cd62710c586842cb08073c
ms.translationtype: HT
ms.contentlocale: de-DE
ms.lasthandoff: 12/04/2020
ms.locfileid: "49571472"
---
# <a name="call-scripts-from-a-manual-power-automate-flow-preview"></a>Aufrufen von Skripts aus einem manuellen Power Automate-Datenfluss (Vorschau)

In diesem Lernprogramm erfahren Sie, wie Sie ein Office-Skript für Excel im Web durch [Power Automate](https://flow.microsoft.com) ausführen. Sie erstellen ein Skript, das die Werte von zwei Zellen mit der aktuellen Zeit aktualisiert. Sie verbinden dieses Skript dann mit einem manuell ausgelösten Power Automate-Flow, so dass das Skript immer dann ausgeführt wird, wenn eine Taste in Power Automate gedrückt wird. Sobald Sie das Grundmuster verstanden haben, können Sie den Ablauf auf andere Anwendungen ausweiten und Ihren täglichen Arbeitsablauf stärker automatisieren.

> [!TIP]
> Wenn Sie mit Office-Skripten noch nicht vertraut sind, empfehlen wir, mit dem [Aufzeichnen, Bearbeiten und Erstellen von Office-Skripten in Excel im Web](excel-tutorial.md)-Lernprogramm zu beginnen. [Office-Skripts verwenden TypeScript](../overview/code-editor-environment.md), und dieses Lernprogramm richten sich an Anfänger bis Fortgeschrittene mit JavaScript oder TypeScript. Wenn Sie noch nicht mit JavaScript vertraut sind, empfehlen wir Ihnen, mit dem [Mozilla-JavaScript-Lernprogramm](https://developer.mozilla.org/docs/Web/JavaScript/Guide/Introduction) zu beginnen.

## <a name="prerequisites"></a>Voraussetzungen

[!INCLUDE [Tutorial prerequisites](../includes/power-automate-tutorial-prerequisites.md)]

## <a name="prepare-the-workbook"></a>Vorbereiten der Arbeitsmappe

Power Automation kann relative Bezüge wie `Workbook.getActiveWorksheet` nicht verwenden, um auf Arbeitsmappenfunktionen zuzugreifen. Deshalb benötigen wir eine Arbeitsmappe und ein Arbeitsblatt mit konsistenten Namen, die Power Automate als Referenz verwenden kann.

1. Erstellen Sie eine neue Arbeitsmappe mit dem Namen **MeineArbeitsmappe**.

2. Erstellen Sie in der Arbeitsmappe **MeineArbeitsmappe** ein Arbeitsblatt namens **Lernprogramm-Arbeitsblatt**.

## <a name="create-an-office-script"></a>Erstellen eines Office-Scripts

1. Wechseln Sie zur Registerkarte **Automate**, und wählen Sie **Code Editor** aus.

2. Wählen Sie **New Script** aus.

3. Ersetzen Sie das Standardskript durch das folgende Skript. Dieses Skript fügt das aktuelle Datum und die aktuelle Uhrzeit zu den ersten beiden Zellen des **Lernprogramm-Arbeitsblatt**-Arbeitsblatts hinzu.

    ```TypeScript
    function main(workbook: ExcelScript.Workbook) {
      // Get the "TutorialWorksheet" worksheet from the workbook.
      let worksheet = workbook.getWorksheet("TutorialWorksheet");

      // Get the cells at A1 and B1.
      let dateRange = worksheet.getRange("A1");
      let timeRange = worksheet.getRange("B1");

      // Get the current date and time using the JavaScript Date object.
      let date = new Date(Date.now());

      // Add the date string to A1.
      dateRange.setValue(date.toLocaleDateString());

      // Add the time string to B1.
      timeRange.setValue(date.toLocaleTimeString());
    }
    ```

4. Benennen Sie das Skript in **Datum und Uhrzeit festlegen** um. Klicken Sie auf den Skriptnamen, um ihn zu ändern.

5. Klicken Sie auf **Save Script**, um das Skript zu speichern.

## <a name="create-an-automated-workflow-with-power-automate"></a>Erstellen eines automatisierten Workflows mit Power Automate

1. Melden Sie sich an der [Power Automate-Website](https://flow.microsoft.com) an.

2. Klicken Sie in dem Menü, das auf der linken Seite des Bildschirms angezeigt wird, auf **Create**. Damit gelangen Sie zur Liste der Möglichkeiten zum Erstellen neuer Workflows.

    ![Die Schaltfläche „Erstellen“ in Power Automate.](../images/power-automate-tutorial-1.png)

3. Wählen Sie im Abschnitt **Start from blank** die Option **Instant flow** aus. Dadurch wird ein manuell aktivierter Workflow erstellt.

    ![Die Option „Instant flow“ zum Erstellen eines neuen Workflows.](../images/power-automate-tutorial-2.png)

4. Geben Sie im daraufhin angezeigten Dialogfenster einen Namen für den Flow in das Textfeld **Flow name** ein, wählen Sie in der Liste der Optionen unter **Choose how to trigger the flow** die Option **Manually trigger a flow** aus, und klicken Sie auf **Create**.

    ![Die manuelle Auslöseroption zum Erstellen eines neuen „Instant flow“.](../images/power-automate-tutorial-3.png)

    Beachten Sie, dass ein manuell ausgelöster Flow nur einer von vielen Arten von Flows ist. Im nächsten Lernprogramm erstellen Sie einen Flow, der automatisch ausgeführt wird, wenn Sie eine E-Mail erhalten.

5. Klicken Sie auf **New step**.

6. Wählen Sie die Registerkarte **Standard** aus, und wählen Sie dann **Excel Online (Business)** aus.

    ![Die Power Automate-Option für Excel Online (Business).](../images/power-automate-tutorial-4.png)

7. Wählen Sie unter **Actions** die Option **Run script (preview)** aus.

    ![Die Power Automate-Aktionsoption für „Run script (preview)“.](../images/power-automate-tutorial-5.png)

8. Als nächstes wählen Sie die Arbeitsmappe und das Skript aus, die im Ablaufschritt verwendet werden sollen. Für das Tutorial verwenden Sie die Arbeitsmappe, die Sie in Ihrem OneDrive erstellt haben. Sie können auch eine beliebige Arbeitsmappe in einer OneDrive- oder SharePoint-Website verwenden. Geben Sie die folgenden Einstellungen für den Konnektor **Run script** an:

    - **Location**: OneDrive for Business
    - **Document Library**: OneDrive
    - **File**: MyWorkbook.xlsx *(Ausgewählt über den Dateibrowser)*
    - **Script**: Datum und Uhrzeit festlegen

    ![Die Konnektoreinstellungen zum Ausführen eines Skripts in Power Automate.](../images/power-automate-tutorial-6.png)

9. Klicken Sie auf **Save**.

Jetzt kann ihr Flow über Power Automate ausgeführt werden. Sie können ihn mithilfe der Schaltfläche **Test** im Flow-Editor testen oder die weiteren Lernprogrammschritte zum Ausführen des Flows aus der Flow-Sammlung ausführen.

## <a name="run-the-script-through-power-automate"></a>Ausführen des Skripts über Power Automate

1. Wählen Sie auf der Hauptseite der Power Automate-Seite **My Flows** aus.

    ![Die Schaltfläche „My Flows“ in Power Automate.](../images/power-automate-tutorial-7.png)

2. Wählen Sie in der Liste der Flows, die auf der Registerkarte **My Flows** angezeigt werden, **Mein Lernprogramm-Flow** aus. Die Details des zuvor erstellten Flows werden angezeigt.

3. Klicken Sie auf **Run**.

    ![Die Schaltfläche „Run“ in Power Automate.](../images/power-automate-tutorial-8.png)

4. Für die Ausführung des Flows wird ein Aufgabenbereich angezeigt. Wenn Sie aufgefordert werden, sich bei Excel Online **anzumelden** klicken Sie auf **Continue**.

5. Klicken Sie auf **Run flow**. Damit wird der Flow ausgeführt, der das zugehörige Office-Skript ausführt.

6. Klicken Sie auf **Done**. Der Abschnitt **Runs** wird entsprechend aktualisiert.

7. Aktualisieren Sie die Seite, um die Ergebnisse von Power Automate anzuzeigen. Wenn der Vorgang erfolgreich war, wechseln Sie zur Arbeitsmappe, um die aktualisierten Zellen anzuzeigen. Falls ein Fehler aufgetreten ist, überprüfen Sie die Einstellungen des Flows, und führen Sie ihn ein zweites Mal aus.

    ![Power Automate-Ausgabe nach einer erfolgreichen Flow-Ausführung.](../images/power-automate-tutorial-9.png)

## <a name="next-steps"></a>Nächste Schritte

Schließen Sie das Lernprogramm [Übergeben von Daten zu Skripts in einem automatisch ausgeführten Power Automate-Datenfluss](excel-power-automate-trigger.md) ab. Hier erfahren Sie, wie Sie Daten aus einem Workflowdienst an Ihr Office-Skript übergeben und den Power Automate-Flow ausführen, wenn bestimmte Ereignisse eintreten.
