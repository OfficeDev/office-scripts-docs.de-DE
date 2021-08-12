---
title: Aufrufen von Skripts aus einem manuellen Power Automate-Datenfluss
description: Ein Lernprogramm zur Verwendung von Office-Skripts in Power Automate durch einen manuellen Auslöser.
ms.date: 06/29/2021
localization_priority: Priority
ms.openlocfilehash: 7244e3b637e21d66bdc83fda2c0cd23b8445a3a6f38a35b0bc3c35cfa8a58440
ms.sourcegitcommit: 75f7ed8c2d23a104acc293f8ce29ea580b4fcdc5
ms.translationtype: HT
ms.contentlocale: de-DE
ms.lasthandoff: 08/11/2021
ms.locfileid: "57847664"
---
# <a name="call-scripts-from-a-manual-power-automate-flow"></a>Aufrufen von Skripts aus einem manuellen Power Automate-Datenfluss

In diesem Lernprogramm erfahren Sie, wie Sie ein Office-Skript für Excel im Web durch [Power Automate](https://flow.microsoft.com) ausführen. Sie erstellen ein Skript, das die Werte von zwei Zellen mit der aktuellen Zeit aktualisiert. Sie verbinden dieses Skript dann mit einem manuell ausgelösten Power Automate-Flow, so dass das Skript immer dann ausgeführt wird, wenn eine Schaltfläche in Power Automate ausgewählt wird. Sobald Sie das Grundmuster verstanden haben, können Sie den Ablauf auf andere Anwendungen ausweiten und Ihren täglichen Arbeitsablauf stärker automatisieren.

> [!TIP]
> Wenn Sie mit Office-Skripten noch nicht vertraut sind, empfehlen wir, mit dem [Aufzeichnen, Bearbeiten und Erstellen von Office-Skripten in Excel im Web](excel-tutorial.md)-Lernprogramm zu beginnen. [Office-Skripts verwenden TypeScript](../overview/code-editor-environment.md), und dieses Lernprogramm richten sich an Anfänger bis Fortgeschrittene mit JavaScript oder TypeScript. Wenn Sie noch nicht mit JavaScript vertraut sind, empfehlen wir Ihnen, mit dem [Mozilla-JavaScript-Lernprogramm](https://developer.mozilla.org/docs/Web/JavaScript/Guide/Introduction) zu beginnen.

## <a name="prerequisites"></a>Voraussetzungen

[!INCLUDE [Tutorial prerequisites](../includes/power-automate-tutorial-prerequisites.md)]

## <a name="prepare-the-workbook"></a>Vorbereiten der Arbeitsmappe

Power Automate sollte für den Zugriff auf Arbeitsmappenkomponenten keine [relativen Bezüge](../testing/power-automate-troubleshooting.md#avoid-relative-references) wie `Workbook.getActiveWorksheet` verwenden. Deshalb benötigen wir eine Arbeitsmappe und ein Arbeitsblatt mit konsistenten Namen, die Power Automate als Referenz verwenden kann.

1. Erstellen Sie eine neue Arbeitsmappe mit dem Namen **MeineArbeitsmappe**.

2. Erstellen Sie in der Arbeitsmappe **MeineArbeitsmappe** ein Arbeitsblatt namens **Lernprogramm-Arbeitsblatt**.

## <a name="create-an-office-script"></a>Erstellen eines Office-Scripts

1. Wechseln Sie zur Registerkarte **Automatisieren**, und wählen Sie **Alle Skripts** aus.

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

4. Benennen Sie das Skript in **Datum und Uhrzeit festlegen** um. Wählen Sie den Skriptnamen aus, um ihn zu ändern.

5. Wählen Sie **Skript speichern** aus, um das Skript zu speichern.

## <a name="create-an-automated-workflow-with-power-automate"></a>Erstellen eines automatisierten Workflows mit Power Automate

1. Melden Sie sich an der [Power Automate-Website](https://flow.microsoft.com) an.

2. Wählen Sie in dem Menü, das auf der linken Seite des Bildschirms angezeigt wird, **Erstellen** aus. Damit gelangen Sie zur Liste der Möglichkeiten zum Erstellen neuer Workflows.

    :::image type="content" source="../images/power-automate-tutorial-1.png" alt-text="Die Schaltfläche „Erstellen“ in Power Automate.":::

3. Wählen Sie im Abschnitt **Start from blank** die Option **Instant flow** aus. Dadurch wird ein manuell aktivierter Workflow erstellt.

    :::image type="content" source="../images/power-automate-tutorial-2.png" alt-text="Die Option „Direktflow“ zum Erstellen eines neuen Workflows in Power Automate.":::

4. Geben Sie im daraufhin angezeigten Dialogfenster einen Namen für den Flow in das Textfeld **Flowname** ein, wählen Sie in der Liste der Optionen unter **Auswählen einer Triggermethode für den Flow** die Option **Einen Flow manuell auslösen** aus, und wählen Sie **Erstellen** aus.

    :::image type="content" source="../images/power-automate-tutorial-3.png" alt-text="Die Option „Einen Flow manuell auslösen“ in Power Automate.":::

    Beachten Sie, dass ein manuell ausgelöster Flow nur einer von vielen Arten von Flows ist. Im nächsten Lernprogramm erstellen Sie einen Flow, der automatisch ausgeführt wird, wenn Sie eine E-Mail erhalten.

5. Wählen Sie **Neuer Schritt** aus.

6. Wählen Sie die Registerkarte **Standard** aus, und wählen Sie dann **Excel Online (Business)** aus.

    :::image type="content" source="../images/power-automate-tutorial-4.png" alt-text="Excel Online (Business)-Option in Power Automate.":::

7. Wählen Sie unter **Aktionen** die Option **Skript ausführen** aus.

    :::image type="content" source="../images/power-automate-tutorial-5.png" alt-text="Aktionsoption „Skript ausführen“ in Power Automate.":::

8. Als nächstes wählen Sie die Arbeitsmappe und das Skript aus, die im Ablaufschritt verwendet werden sollen. Für das Tutorial verwenden Sie die Arbeitsmappe, die Sie in Ihrem OneDrive erstellt haben. Sie können auch eine beliebige Arbeitsmappe in einer OneDrive- oder SharePoint-Website verwenden. Geben Sie die folgenden Einstellungen für den Konnektor **Run script** an:

    - **Location**: OneDrive for Business
    - **Document Library**: OneDrive
    - **File**: MyWorkbook.xlsx *(Ausgewählt über den Dateibrowser)*
    - **Script**: Datum und Uhrzeit festlegen

    :::image type="content" source="../images/power-automate-tutorial-6.png" alt-text="Die Einstellungen des Power Automate-Connectors zum Ausführen eines Skripts.":::

9. Wählen Sie **Speichern** aus.

Jetzt kann ihr Flow über Power Automate ausgeführt werden. Sie können ihn mithilfe der Schaltfläche **Test** im Flow-Editor testen oder die weiteren Lernprogrammschritte zum Ausführen des Flows aus der Flow-Sammlung ausführen.

## <a name="run-the-script-through-power-automate"></a>Ausführen des Skripts über Power Automate

1. Wählen Sie auf der Hauptseite der Power Automate-Seite **My Flows** aus.

    :::image type="content" source="../images/power-automate-tutorial-7.png" alt-text="Die Schaltfläche „Meine Flows“ in Power Automate.":::

2. Wählen Sie in der Liste der Flows, die auf der Registerkarte **My Flows** angezeigt werden, **Mein Lernprogramm-Flow** aus. Die Details des zuvor erstellten Flows werden angezeigt.

3. Wählen Sie **Ausführen** aus.

    :::image type="content" source="../images/power-automate-tutorial-8.png" alt-text="Schaltfläche „Ausführen“ in Power Automate.":::

4. Für die Ausführung des Flows wird ein Aufgabenbereich angezeigt. Wenn Sie aufgefordert werden, sich bei Excel Online **anzumelden** klicken Sie auf **Weiter**.

5. Wählen Sie **Flow ausführen** aus. Damit wird der Flow ausgeführt, der das zugehörige Office-Skript ausführt.

6. Wählen Sie **Fertig** aus. Der Abschnitt **Runs** wird entsprechend aktualisiert.

7. Aktualisieren Sie die Seite, um die Ergebnisse von Power Automate anzuzeigen. Wenn der Vorgang erfolgreich war, wechseln Sie zur Arbeitsmappe, um die aktualisierten Zellen anzuzeigen. Falls ein Fehler aufgetreten ist, überprüfen Sie die Einstellungen des Flows, und führen Sie ihn ein zweites Mal aus.

    :::image type="content" source="../images/power-automate-tutorial-9.png" alt-text="Power Automate-Ausgabe mit einer erfolgreichen Ausführung des Flows.":::

## <a name="next-steps"></a>Nächste Schritte

Schließen Sie das Lernprogramm [Übergeben von Daten zu Skripts in einem automatisch ausgeführten Power Automate-Datenfluss](excel-power-automate-trigger.md) ab. Hier erfahren Sie, wie Sie Daten aus einem Workflowdienst an Ihr Office-Skript übergeben und den Power Automate-Flow ausführen, wenn bestimmte Ereignisse eintreten.
