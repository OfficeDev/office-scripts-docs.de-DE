---
title: Zurückgeben von Daten aus einem Skript an einen automatisch ausgeführten Power Automate-Flow
description: Ein Tutorial, das zeigt, wie Sie Erinnerungs-E-Mails senden, indem Sie Office-Skripts für Excel im Web über Power Automate ausführen.
ms.date: 06/29/2021
ms.localizationpriority: high
ms.openlocfilehash: e100fac263dee8f1f39529bd83610576e68eb2e6
ms.sourcegitcommit: 7023b9e23499806901a5ecf8ebc460b76887cca6
ms.translationtype: HT
ms.contentlocale: de-DE
ms.lasthandoff: 03/31/2022
ms.locfileid: "64586052"
---
# <a name="return-data-from-a-script-to-an-automatically-run-power-automate-flow"></a>Zurückgeben von Daten aus einem Skript an einen automatisch ausgeführten Power Automate-Flow

In diesem Tutorial erfahren Sie, wie Sie Informationen eines Office-Skripts für Excel im Web mit als Teil eines automatisierten [Power Automate](https://flow.microsoft.com)-Workflows zurückgeben. Sie erstellen ein Skript, das einen Zeitplan durchsucht und einen Flow verwendet, um Erinnerungs-E-Mails zu senden. Dieser Flow wird nach einem regelmäßigen Zeitplan ausgeführt und stellt diese Erinnerungen in Ihrem Namen bereit.

> [!TIP]
> Wenn Sie mit Office-Skripten noch nicht vertraut sind, empfehlen wir, mit dem [Aufzeichnen, Bearbeiten und Erstellen von Office-Skripten in Excel im Web](excel-tutorial.md)-Lernprogramm zu beginnen.
>
> Wenn Sie noch nicht mit Power Automate vertraut sind, empfehlen wir, mit den Tutorials [Aufrufen von Skripts aus einem manuellen Power Automate-Datenfluss](excel-power-automate-manual.md) und [Übergeben von Daten an Skripts in einem automatisch ausgeführten Power Automate-Flow](excel-power-automate-trigger.md) zu beginnen.
>
> [Office-Skripts verwenden TypeScript](../overview/code-editor-environment.md), und dieses Lernprogramm richten sich an Anfänger bis Fortgeschrittene mit JavaScript oder TypeScript. Wenn Sie noch nicht mit JavaScript vertraut sind, empfehlen wir Ihnen, mit dem [Mozilla-JavaScript-Lernprogramm](https://developer.mozilla.org/docs/Web/JavaScript/Guide/Introduction) zu beginnen.

## <a name="prerequisites"></a>Voraussetzungen

[!INCLUDE [Tutorial prerequisites](../includes/power-automate-tutorial-prerequisites.md)]

## <a name="prepare-the-workbook"></a>Vorbereiten der Arbeitsmappe

1. Laden Sie die Arbeitsmappe <a href="on-call-rotation.xlsx">on-call-rotation.xlsx</a> auf Ihr OneDrive herunter.

1. Öffnen Sie **on-call-rotation.xlsx** in Excel im Web.

1. Fügen Sie der Tabelle eine Zeile mit Ihrem Namen, Ihrer E-Mail-Adresse sowie dem Start- und Enddatum hinzu, die sich mit dem aktuellen Datum überschneiden.

    > [!IMPORTANT]
    > Das Skript, das Sie schreiben, verwendet den ersten übereinstimmenden Eintrag in der Tabelle, stellen Sie also sicher, dass Ihr Name über einer Zeile mit der aktuellen Woche steht.

    :::image type="content" source="../images/power-automate-return-tutorial-1.png" alt-text="Ein Arbeitsblatt mit den Daten der Rotationstabelle auf Abruf.":::

## <a name="create-an-office-script"></a>Erstellen eines Office-Skripts

1. Wechseln Sie zur Registerkarte **Automatisieren**, und wählen Sie **Alle Skripts** aus.

1. Wählen Sie **Neues Skript** aus.

1. Nennen Sie das Skript **Person mit Rufbereitschaft abrufen**.

1. Sie sollten nun ein leeres Skript haben. Sie möchten das Skript verwenden, um eine E-Mail-Adresse aus dem Arbeitsblatt abzurufen. Ändern Sie `main` folgendermaßen, damit eine Zeichenfolge zurückgegeben wird:

    ```TypeScript
    function main(workbook: ExcelScript.Workbook) : string {
    }
    ```

1. Als nächstes müssen alle Daten aus der Tabelle abgerufen werden. Damit können Sie jede Zeile mit dem Skript durchsuchen. Fügen Sie im `main`-Funktion den folgenden Code hinzu.

    ```TypeScript
    // Get the H1 worksheet.
    let worksheet = workbook.getWorksheet("H1");

    // Get the first (and only) table in the worksheet.
    let table = worksheet.getTables()[0];

    // Get the data from the table.
    let tableValues = table.getRangeBetweenHeaderAndTotal().getValues();
    ```

1. Die Datumsangaben in der Tabelle werden mit der [fortlaufenden Zahl für das Datum von Excel](https://support.microsoft.com/office/e7fe7167-48a9-4b96-bb53-5612a800b487) gespeichert. Diese Datumsangaben müssen in JavaScript-Datumsangaben konvertiert werden, um sie vergleichen zu können. Sie fügen eine Hilfsfunktion in Ihr Skript ein. Fügen Sie außerhalb der `main`-Funktion den folgenden Code hinzu:

    ```TypeScript
    // Convert the Excel date to a JavaScript Date object.
    function convertDate(excelDateValue: number) {
        let javaScriptDate = new Date(Math.round((excelDateValue - 25569) * 86400 * 1000));
        return javaScriptDate;
    }
    ```

1. Jetzt müssen Sie herausfinden, welche Person gerade Rufbereitschaft hat. Ihre Zeile enthält ein Start- und Enddatum, die das aktuelle Datum einschließen. Beim Schreiben des Skripts wird davon ausgegangen, dass immer nur eine Person in Rufbereitschaft ist. Skripts können Arrays zurückgeben, um mehrere Werte zu verarbeiten, aber für den Moment wird die erste übereinstimmende E-Mail-Adresse zurückgegeben. Fügen Sie am Ende der `main`-Funktion den folgenden Code hinzu.

    ```TypeScript
    // Look for the first row where today's date is between the row's start and end dates.
    let currentDate = new Date();
    for (let row = 0; row < tableValues.length; row++) {
        let startDate = convertDate(tableValues[row][2] as number);
        let endDate = convertDate(tableValues[row][3] as number);
        if (startDate <= currentDate && endDate >= currentDate) {
            // Return the first matching email address.
            return tableValues[row][1].toString();
        }
    }
    ```

1. Das fertige Skript sollte wie folgt aussehen:

    ```TypeScript
    function main(workbook: ExcelScript.Workbook) : string {
        // Get the H1 worksheet.
        let worksheet = workbook.getWorksheet("H1");

        // Get the first (and only) table in the worksheet.
        let table = worksheet.getTables()[0];
    
        // Get the data from the table.
        let tableValues = table.getRangeBetweenHeaderAndTotal().getValues();
    
        // Look for the first row where today's date is between the row's start and end dates.
        let currentDate = new Date();
        for (let row = 0; row < tableValues.length; row++) {
            let startDate = convertDate(tableValues[row][2] as number);
            let endDate = convertDate(tableValues[row][3] as number);
            if (startDate <= currentDate && endDate >= currentDate) {
                // Return the first matching email address.
                return tableValues[row][1].toString();
            }
        }
    }

    // Convert the Excel date to a JavaScript Date object.
    function convertDate(excelDateValue: number) {
        let javaScriptDate = new Date(Math.round((excelDateValue - 25569) * 86400 * 1000));
        return javaScriptDate;
    }
    ```

## <a name="create-an-automated-workflow-with-power-automate"></a>Erstellen eines automatisierten Workflows mit Power Automate

1. Melden Sie sich an der [Power Automate-Website](https://flow.microsoft.com) an.

1. Wählen Sie in dem Menü, das auf der linken Seite des Bildschirms angezeigt wird, **Erstellen** aus. Damit gelangen Sie zur Liste der Möglichkeiten zum Erstellen neuer Workflows.

    :::image type="content" source="../images/power-automate-tutorial-1.png" alt-text="Schaltfläche „Erstellen“ in Power Automate.":::

1. Wählen Sie im Abschnitt **Ohne Vorlage anfangen** die Option **Automatisierter Cloudfluss** aus.

    :::image type="content" source="../images/power-automate-return-tutorial-2.png" alt-text="Schaltfläche „Geplanter Cloudfluss“ in Power Automate.":::

1. Nun müssen Sie den Zeitplan für diesen Flow festlegen. Unser Arbeitsblatt weist eine neue Rufbereitschaftszuweisung auf, die in der ersten Hälfte des Jahres 2021 jeweils montags beginnt. Legen Sie den Flow so fest, dass er montagmorgens als erstes ausgeführt wird. Verwenden Sie die folgenden Optionen, um den Flow so zu konfigurieren, dass er jede Woche montags ausgeführt wird.

    - **Flowname**: Person mit Rufbereitschaft benachrichtigen
    - **Startdatum**: 4.1.21 um 1:00 Uhr
    - **Wiederholen**: 1-mal wöchentlich
    - **an diesen Tagen**: Mo

    :::image type="content" source="../images/power-automate-return-tutorial-3.png" alt-text="Das Power Automate-Dialogfeld ‚Erstellen eines geplanten Cloudflusses‘ mit Optionen. Zu den Optionen gehören der Flussname, die Anfangszeit, die Wiederholungszeit und ein Wochentag, an dem der Fluss ausgeführt werden soll.":::

1. Wählen Sie **Erstellen** aus.

1. Wählen Sie **Neuer Schritt** aus.

1. Wählen Sie die Registerkarte **Standard** aus, und wählen Sie dann **Excel Online (Business)** aus.

    :::image type="content" source="../images/power-automate-tutorial-4.png" alt-text="Excel Online (Business)-Option in Power Automate.":::

1. Wählen Sie unter **Aktionen** die Option **Skript ausführen** aus.

    :::image type="content" source="../images/power-automate-tutorial-5.png" alt-text="Aktionsoption „Skript ausführen“ in Power Automate.":::

1. Als nächstes wählen Sie die Arbeitsmappe und das Skript aus, die im Flowschritt verwendet werden sollen. Verwenden Sie die Arbeitsmappe **on-call-rotation.xlsx**, die Sie in Ihrem OneDrive erstellt haben. Geben Sie die folgenden Einstellungen für den Konnektor **Run script** an:

    - **Location**: OneDrive for Business
    - **Document Library**: OneDrive
    - **Datei**: MyWorkbook.xlsx *(ausgewählt über den Dateibrowser)*
    - **Skript**: Person mit Rufbereitschaft abrufen

    :::image type="content" source="../images/power-automate-return-tutorial-4.png" alt-text="Die Einstellungen des Power Automate-Connectors zum Ausführen eines Skripts.":::

1. Wählen Sie **Neuer Schritt** aus.

1. Der Flow wird mit dem Senden der Erinnerungs-E-Mail beendet. Wählen Sie **E-Mail senden (V2)** über die Suchleiste des Connectors aus. Verwenden Sie das Steuerelement **Dynamischen Inhalt hinzufügen**, um die vom Skript zurückgegebene E-Mail-Adresse hinzuzufügen. Dies wird mit dem Excel-Symbol daneben als **Ergebnis** gekennzeichnet. Sie können einen beliebigen Betreff und Text eingeben.

    :::image type="content" source="../images/power-automate-return-tutorial-5.png" alt-text="Die Power Automate Outlook Connector-Einstellungen zum Senden einer E-Mail. Zu den Optionen gehören die zu sendende Datei, der Betreff der E-Mail und der Textkörper der E-Mail sowie erweiterte Optionen.":::

    > [!NOTE]
    > In diesem Tutorial wird Outlook verwendet. Sie können stattdessen auch Ihren bevorzugten E-Mail-Dienst verwenden, obwohl einige Optionen anders sein können.

1. Wählen Sie **Speichern** aus.

## <a name="test-the-script-in-power-automate"></a>Testen des Skripts in Power Automate

Ihr Flow wird jeden Montagmorgen ausgeführt. Sie können das Skript jetzt testen, indem Sie die Schaltfläche **Test** in der oberen rechten Ecke des Bildschirms auswählen. Wählen Sie **Manuell** und anschließend **Test ausführen** aus, um den Flow jetzt auszuführen und das Verhalten zu testen. Möglicherweise müssen Sie Excel und Outlook Berechtigungen erteilen, um fortzufahren.

:::image type="content" source="../images/power-automate-return-tutorial-6.png" alt-text="Schaltfläche „Power Automate-Test“.":::

> [!TIP]
> Wenn Ihr Flow keine E-Mail senden kann, vergewissern Sie sich, dass in das Arbeitsblatt eine gültige E-Mail für den aktuellen Datumsbereich oben in der Tabelle aufgeführt ist.

## <a name="next-steps"></a>Nächste Schritte

Besuchen Sie [Ausführen von Office-Skripts mit Power Automate](../develop/power-automate-integration.md), um mehr über das Verbinden von Office-Skripts mit Power Automate zu erfahren.

Sie können sich auch das Beispielszenario [Automatisierte Aufgabenerinnerungen](../resources/scenarios/task-reminders.md) ansehen, um zu erfahren, wie Sie Office-Skripts und Power Automate mit Teams Adaptive Cards kombinieren können.
