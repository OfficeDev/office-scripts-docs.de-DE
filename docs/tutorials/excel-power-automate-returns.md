---
title: Zurückgeben von Daten aus einem Skript an einen automatisch ausgeführten Power Automate-Flow
description: Ein Tutorial, das zeigt, wie Sie Erinnerungs-E-Mails senden, indem Sie Office-Skripts für Excel im Web über Power Automate ausführen.
ms.date: 06/29/2021
localization_priority: Priority
ms.openlocfilehash: 6c94ba4382f9d481c0064e89b5f7afa147ab23f4
ms.sourcegitcommit: 211c157ca746e266eeb079f5fa1925a1e35ab702
ms.translationtype: HT
ms.contentlocale: de-DE
ms.lasthandoff: 07/07/2021
ms.locfileid: "53314002"
---
# <a name="return-data-from-a-script-to-an-automatically-run-power-automate-flow"></a><span data-ttu-id="36822-103">Zurückgeben von Daten aus einem Skript an einen automatisch ausgeführten Power Automate-Flow</span><span class="sxs-lookup"><span data-stu-id="36822-103">Return data from a script to an automatically-run Power Automate flow</span></span>

<span data-ttu-id="36822-104">In diesem Tutorial erfahren Sie, wie Sie Informationen eines Office-Skripts für Excel im Web mit als Teil eines automatisierten [Power Automate](https://flow.microsoft.com)-Workflows zurückgeben.</span><span class="sxs-lookup"><span data-stu-id="36822-104">This tutorial teaches you how to return information from an Office Script for Excel on the web as part of an automated [Power Automate](https://flow.microsoft.com) workflow.</span></span> <span data-ttu-id="36822-105">Sie erstellen ein Skript, das einen Zeitplan durchsucht und einen Flow verwendet, um Erinnerungs-E-Mails zu senden.</span><span class="sxs-lookup"><span data-stu-id="36822-105">You'll make a script that looks through a schedule and works with a flow to send reminder emails.</span></span> <span data-ttu-id="36822-106">Dieser Flow wird nach einem regelmäßigen Zeitplan ausgeführt und stellt diese Erinnerungen in Ihrem Namen bereit.</span><span class="sxs-lookup"><span data-stu-id="36822-106">This flow will run on a regular schedule, providing these reminders on your behalf.</span></span>

> [!TIP]
> <span data-ttu-id="36822-107">Wenn Sie mit Office-Skripten noch nicht vertraut sind, empfehlen wir, mit dem [Aufzeichnen, Bearbeiten und Erstellen von Office-Skripten in Excel im Web](excel-tutorial.md)-Lernprogramm zu beginnen.</span><span class="sxs-lookup"><span data-stu-id="36822-107">If you are new to Office Scripts, we recommend starting with the [Record, edit, and create Office Scripts in Excel on the web](excel-tutorial.md) tutorial.</span></span>
>
> <span data-ttu-id="36822-108">Wenn Sie noch nicht mit Power Automate vertraut sind, empfehlen wir, mit den Tutorials [Aufrufen von Skripts aus einem manuellen Power Automate-Datenfluss](excel-power-automate-manual.md) und [Übergeben von Daten an Skripts in einem automatisch ausgeführten Power Automate-Flow](excel-power-automate-trigger.md) zu beginnen.</span><span class="sxs-lookup"><span data-stu-id="36822-108">If you are new to Power Automate, we recommend starting with the [Call scripts from a manual Power Automate flow](excel-power-automate-manual.md) and [Pass data to scripts in an automatically-run Power Automate flow](excel-power-automate-trigger.md) tutorials.</span></span>
>
> <span data-ttu-id="36822-109">[Office-Skripts verwenden TypeScript](../overview/code-editor-environment.md), und dieses Lernprogramm richten sich an Anfänger bis Fortgeschrittene mit JavaScript oder TypeScript.</span><span class="sxs-lookup"><span data-stu-id="36822-109">[Office Scripts use TypeScript](../overview/code-editor-environment.md) and this tutorial is intended for people with beginner to intermediate-level knowledge of JavaScript or TypeScript.</span></span> <span data-ttu-id="36822-110">Wenn Sie noch nicht mit JavaScript vertraut sind, empfehlen wir Ihnen, mit dem [Mozilla-JavaScript-Lernprogramm](https://developer.mozilla.org/docs/Web/JavaScript/Guide/Introduction) zu beginnen.</span><span class="sxs-lookup"><span data-stu-id="36822-110">If you're new to JavaScript, we recommend starting with the [Mozilla JavaScript tutorial](https://developer.mozilla.org/docs/Web/JavaScript/Guide/Introduction).</span></span>

## <a name="prerequisites"></a><span data-ttu-id="36822-111">Voraussetzungen</span><span class="sxs-lookup"><span data-stu-id="36822-111">Prerequisites</span></span>

[!INCLUDE [Tutorial prerequisites](../includes/power-automate-tutorial-prerequisites.md)]

## <a name="prepare-the-workbook"></a><span data-ttu-id="36822-112">Vorbereiten der Arbeitsmappe</span><span class="sxs-lookup"><span data-stu-id="36822-112">Prepare the workbook</span></span>

1. <span data-ttu-id="36822-113">Laden Sie die Arbeitsmappe <a href="on-call-rotation.xlsx">on-call-rotation.xlsx</a> auf Ihr OneDrive herunter.</span><span class="sxs-lookup"><span data-stu-id="36822-113">Download the workbook <a href="on-call-rotation.xlsx">on-call-rotation.xlsx</a> to your OneDrive.</span></span>

1. <span data-ttu-id="36822-114">Öffnen Sie **on-call-rotation.xlsx** in Excel im Web.</span><span class="sxs-lookup"><span data-stu-id="36822-114">Open **on-call-rotation.xlsx** in Excel on the web.</span></span>

1. <span data-ttu-id="36822-115">Fügen Sie der Tabelle eine Zeile mit Ihrem Namen, Ihrer E-Mail-Adresse sowie dem Start- und Enddatum hinzu, die sich mit dem aktuellen Datum überschneiden.</span><span class="sxs-lookup"><span data-stu-id="36822-115">Add a row to the table with your name, email address, and start and end dates that overlap with the current date.</span></span>

    > [!IMPORTANT]
    > <span data-ttu-id="36822-116">Das Skript, das Sie schreiben, verwendet den ersten übereinstimmenden Eintrag in der Tabelle, stellen Sie also sicher, dass Ihr Name über einer Zeile mit der aktuellen Woche steht.</span><span class="sxs-lookup"><span data-stu-id="36822-116">The script you'll write uses the first matching entry in the table, so make sure your name is above any row with the current week.</span></span>

    :::image type="content" source="../images/power-automate-return-tutorial-1.png" alt-text="Ein Arbeitsblatt mit den Daten der Rotationstabelle auf Abruf.":::

## <a name="create-an-office-script"></a><span data-ttu-id="36822-118">Erstellen eines Office-Skripts</span><span class="sxs-lookup"><span data-stu-id="36822-118">Create an Office Script</span></span>

1. <span data-ttu-id="36822-119">Wechseln Sie zur Registerkarte **Automatisieren**, und wählen Sie **Alle Skripts** aus.</span><span class="sxs-lookup"><span data-stu-id="36822-119">Go to the **Automate** tab and select **All Scripts**.</span></span>

1. <span data-ttu-id="36822-120">Wählen Sie **Neues Skript** aus.</span><span class="sxs-lookup"><span data-stu-id="36822-120">Select **New Script**.</span></span>

1. <span data-ttu-id="36822-121">Nennen Sie das Skript **Person mit Rufbereitschaft abrufen**.</span><span class="sxs-lookup"><span data-stu-id="36822-121">Name the script **Get On-Call Person**.</span></span>

1. <span data-ttu-id="36822-122">Sie sollten nun ein leeres Skript haben.</span><span class="sxs-lookup"><span data-stu-id="36822-122">You should now have an empty script.</span></span> <span data-ttu-id="36822-123">Sie möchten das Skript verwenden, um eine E-Mail-Adresse aus dem Arbeitsblatt abzurufen.</span><span class="sxs-lookup"><span data-stu-id="36822-123">We want to use the script to get an email address from the spreadsheet.</span></span> <span data-ttu-id="36822-124">Ändern Sie `main` folgendermaßen, damit eine Zeichenfolge zurückgegeben wird:</span><span class="sxs-lookup"><span data-stu-id="36822-124">Change `main` to return a string, like this:</span></span>

    ```TypeScript
    function main(workbook: ExcelScript.Workbook) : string {
    }
    ```

1. <span data-ttu-id="36822-125">Als nächstes müssen alle Daten aus der Tabelle abgerufen werden.</span><span class="sxs-lookup"><span data-stu-id="36822-125">Next, we need to get all the data from the table.</span></span> <span data-ttu-id="36822-126">Damit können Sie jede Zeile mit dem Skript durchsuchen.</span><span class="sxs-lookup"><span data-stu-id="36822-126">That lets us look through each row with the script.</span></span> <span data-ttu-id="36822-127">Fügen Sie im `main`-Funktion den folgenden Code hinzu.</span><span class="sxs-lookup"><span data-stu-id="36822-127">Add the following code inside the `main` function.</span></span>

    ```TypeScript
    // Get the H1 worksheet.
    let worksheet = workbook.getWorksheet("H1");

    // Get the first (and only) table in the worksheet.
    let table = worksheet.getTables()[0];

    // Get the data from the table.
    let tableValues = table.getRangeBetweenHeaderAndTotal().getValues();
    ```

1. <span data-ttu-id="36822-128">Die Datumsangaben in der Tabelle werden mit der [fortlaufenden Zahl für das Datum von Excel](https://support.microsoft.com/office/date-systems-in-excel-e7fe7167-48a9-4b96-bb53-5612a800b487) gespeichert.</span><span class="sxs-lookup"><span data-stu-id="36822-128">The dates in the table are stored using [Excel's date serial number](https://support.microsoft.com/office/date-systems-in-excel-e7fe7167-48a9-4b96-bb53-5612a800b487).</span></span> <span data-ttu-id="36822-129">Diese Datumsangaben müssen in JavaScript-Datumsangaben konvertiert werden, um sie vergleichen zu können.</span><span class="sxs-lookup"><span data-stu-id="36822-129">We need to convert those dates to JavaScript dates in order to compare them.</span></span> <span data-ttu-id="36822-130">Sie fügen eine Hilfsfunktion in Ihr Skript ein.</span><span class="sxs-lookup"><span data-stu-id="36822-130">We'll add a helper function to our script.</span></span> <span data-ttu-id="36822-131">Fügen Sie außerhalb der `main`-Funktion den folgenden Code hinzu:</span><span class="sxs-lookup"><span data-stu-id="36822-131">Add the following code outside of the `main` function:</span></span>

    ```TypeScript
    // Convert the Excel date to a JavaScript Date object.
    function convertDate(excelDateValue: number) {
        let javaScriptDate = new Date(Math.round((excelDateValue - 25569) * 86400 * 1000));
        return javaScriptDate;
    }
    ```

1. <span data-ttu-id="36822-132">Jetzt müssen Sie herausfinden, welche Person gerade Rufbereitschaft hat.</span><span class="sxs-lookup"><span data-stu-id="36822-132">Now, we need to figure out which person is on call right now.</span></span> <span data-ttu-id="36822-133">Ihre Zeile enthält ein Start- und Enddatum, die das aktuelle Datum einschließen.</span><span class="sxs-lookup"><span data-stu-id="36822-133">Their row will have a start and end date surrounding the current date.</span></span> <span data-ttu-id="36822-134">Beim Schreiben des Skripts wird davon ausgegangen, dass immer nur eine Person in Rufbereitschaft ist.</span><span class="sxs-lookup"><span data-stu-id="36822-134">We'll write the script to assume only one person is on call at a time.</span></span> <span data-ttu-id="36822-135">Skripts können Arrays zurückgeben, um mehrere Werte zu verarbeiten, aber für den Moment wird die erste übereinstimmende E-Mail-Adresse zurückgegeben.</span><span class="sxs-lookup"><span data-stu-id="36822-135">Scripts can return arrays to handle multiple values, but for now we'll return the first matching email address.</span></span> <span data-ttu-id="36822-136">Fügen Sie am Ende der `main`-Funktion den folgenden Code hinzu.</span><span class="sxs-lookup"><span data-stu-id="36822-136">Add the following code to the end of the `main` function.</span></span>

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

1. <span data-ttu-id="36822-137">Das fertige Skript sollte wie folgt aussehen:</span><span class="sxs-lookup"><span data-stu-id="36822-137">The final script should look like this:</span></span>

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

## <a name="create-an-automated-workflow-with-power-automate"></a><span data-ttu-id="36822-138">Erstellen eines automatisierten Workflows mit Power Automate</span><span class="sxs-lookup"><span data-stu-id="36822-138">Create an automated workflow with Power Automate</span></span>

1. <span data-ttu-id="36822-139">Melden Sie sich an der [Power Automate-Website](https://flow.microsoft.com) an.</span><span class="sxs-lookup"><span data-stu-id="36822-139">Sign in to the [Power Automate site](https://flow.microsoft.com).</span></span>

1. <span data-ttu-id="36822-140">Wählen Sie in dem Menü, das auf der linken Seite des Bildschirms angezeigt wird, **Erstellen** aus.</span><span class="sxs-lookup"><span data-stu-id="36822-140">In the menu that's displayed on the left side of the screen, select **Create**.</span></span> <span data-ttu-id="36822-141">Damit gelangen Sie zur Liste der Möglichkeiten zum Erstellen neuer Workflows.</span><span class="sxs-lookup"><span data-stu-id="36822-141">This brings you to list of ways to create new workflows.</span></span>

    :::image type="content" source="../images/power-automate-tutorial-1.png" alt-text="Schaltfläche „Erstellen“ in Power Automate.":::

1. <span data-ttu-id="36822-143">Wählen Sie im Abschnitt **Ohne Vorlage anfangen** die Option **Automatisierter Cloudfluss** aus.</span><span class="sxs-lookup"><span data-stu-id="36822-143">Under the **Start from blank** section, select **Scheduled cloud flow**.</span></span>

    :::image type="content" source="../images/power-automate-return-tutorial-2.png" alt-text="Schaltfläche „Geplanter Cloudfluss“ in Power Automate.":::

1. <span data-ttu-id="36822-145">Nun müssen Sie den Zeitplan für diesen Flow festlegen.</span><span class="sxs-lookup"><span data-stu-id="36822-145">Now we need to set the schedule for this flow.</span></span> <span data-ttu-id="36822-146">Unser Arbeitsblatt weist eine neue Rufbereitschaftszuweisung auf, die in der ersten Hälfte des Jahres 2021 jeweils montags beginnt.</span><span class="sxs-lookup"><span data-stu-id="36822-146">Our spreadsheet has a new on-call assignment starting every Monday in the first half of 2021.</span></span> <span data-ttu-id="36822-147">Legen Sie den Flow so fest, dass er montagmorgens als erstes ausgeführt wird.</span><span class="sxs-lookup"><span data-stu-id="36822-147">Let's set the flow to run first thing Monday mornings.</span></span> <span data-ttu-id="36822-148">Verwenden Sie die folgenden Optionen, um den Flow so zu konfigurieren, dass er jede Woche montags ausgeführt wird.</span><span class="sxs-lookup"><span data-stu-id="36822-148">Use the following options to configure the flow to run on Monday each week.</span></span>

    - <span data-ttu-id="36822-149">**Flowname**: Person mit Rufbereitschaft benachrichtigen</span><span class="sxs-lookup"><span data-stu-id="36822-149">**Flow name**: Notify On-Call Person</span></span>
    - <span data-ttu-id="36822-150">**Startdatum**: 4.1.21 um 1:00 Uhr</span><span class="sxs-lookup"><span data-stu-id="36822-150">**Starting**: 1/4/21 at 1:00am</span></span>
    - <span data-ttu-id="36822-151">**Wiederholen**: 1-mal wöchentlich</span><span class="sxs-lookup"><span data-stu-id="36822-151">**Repeat every**: 1 Week</span></span>
    - <span data-ttu-id="36822-152">**an diesen Tagen**: Mo</span><span class="sxs-lookup"><span data-stu-id="36822-152">**On these days**: M</span></span>

    :::image type="content" source="../images/power-automate-return-tutorial-3.png" alt-text="Das Power Automate-Dialogfeld ‚Erstellen eines geplanten Cloudflusses‘ mit Optionen. Zu den Optionen gehören der Flussname, die Anfangszeit, die Wiederholungszeit und ein Wochentag, an dem der Fluss ausgeführt werden soll.":::

1. <span data-ttu-id="36822-154">Wählen Sie **Erstellen** aus.</span><span class="sxs-lookup"><span data-stu-id="36822-154">Select **Create**.</span></span>

1. <span data-ttu-id="36822-155">Wählen Sie **Neuer Schritt** aus.</span><span class="sxs-lookup"><span data-stu-id="36822-155">Select **New step**.</span></span>

1. <span data-ttu-id="36822-156">Wählen Sie die Registerkarte **Standard** aus, und wählen Sie dann **Excel Online (Business)** aus.</span><span class="sxs-lookup"><span data-stu-id="36822-156">Select the **Standard** tab, then select **Excel Online (Business)**.</span></span>

    :::image type="content" source="../images/power-automate-tutorial-4.png" alt-text="Excel Online (Business)-Option in Power Automate.":::

1. <span data-ttu-id="36822-158">Wählen Sie unter **Aktionen** die Option **Skript ausführen** aus.</span><span class="sxs-lookup"><span data-stu-id="36822-158">Under **Actions**, select **Run script**.</span></span>

    :::image type="content" source="../images/power-automate-tutorial-5.png" alt-text="Aktionsoption „Skript ausführen“ in Power Automate.":::

1. <span data-ttu-id="36822-160">Als nächstes wählen Sie die Arbeitsmappe und das Skript aus, die im Flowschritt verwendet werden sollen.</span><span class="sxs-lookup"><span data-stu-id="36822-160">Next, you'll select the workbook and script to use in the flow step.</span></span> <span data-ttu-id="36822-161">Verwenden Sie die Arbeitsmappe **on-call-rotation.xlsx**, die Sie in Ihrem OneDrive erstellt haben.</span><span class="sxs-lookup"><span data-stu-id="36822-161">Use the **on-call-rotation.xlsx** workbook you created in your OneDrive.</span></span> <span data-ttu-id="36822-162">Geben Sie die folgenden Einstellungen für den Konnektor **Run script** an:</span><span class="sxs-lookup"><span data-stu-id="36822-162">Specify the following settings for the **Run script** connector:</span></span>

    - <span data-ttu-id="36822-163">**Location**: OneDrive for Business</span><span class="sxs-lookup"><span data-stu-id="36822-163">**Location**: OneDrive for Business</span></span>
    - <span data-ttu-id="36822-164">**Document Library**: OneDrive</span><span class="sxs-lookup"><span data-stu-id="36822-164">**Document Library**: OneDrive</span></span>
    - <span data-ttu-id="36822-165">**Datei**: MyWorkbook.xlsx *(ausgewählt über den Dateibrowser)*</span><span class="sxs-lookup"><span data-stu-id="36822-165">**File**: on-call-rotation.xlsx *(Chosen through the file browser)*</span></span>
    - <span data-ttu-id="36822-166">**Skript**: Person mit Rufbereitschaft abrufen</span><span class="sxs-lookup"><span data-stu-id="36822-166">**Script**: Get On-Call Person</span></span>

    :::image type="content" source="../images/power-automate-return-tutorial-4.png" alt-text="Die Einstellungen des Power Automate-Connectors zum Ausführen eines Skripts.":::

1. <span data-ttu-id="36822-168">Wählen Sie **Neuer Schritt** aus.</span><span class="sxs-lookup"><span data-stu-id="36822-168">Select **New step**.</span></span>

1. <span data-ttu-id="36822-169">Der Flow wird mit dem Senden der Erinnerungs-E-Mail beendet.</span><span class="sxs-lookup"><span data-stu-id="36822-169">We'll end the flow by sending the reminder email.</span></span> <span data-ttu-id="36822-170">Wählen Sie **E-Mail senden (V2)** über die Suchleiste des Connectors aus.</span><span class="sxs-lookup"><span data-stu-id="36822-170">Select **Send an email (V2)** by using the connector's search bar.</span></span> <span data-ttu-id="36822-171">Verwenden Sie das Steuerelement **Dynamischen Inhalt hinzufügen**, um die vom Skript zurückgegebene E-Mail-Adresse hinzuzufügen.</span><span class="sxs-lookup"><span data-stu-id="36822-171">Use the **Add dynamic content** control to add the email address returned by the script.</span></span> <span data-ttu-id="36822-172">Dies wird mit dem Excel-Symbol daneben als **Ergebnis** gekennzeichnet.</span><span class="sxs-lookup"><span data-stu-id="36822-172">This will be labelled **result** with the Excel icon next to it.</span></span> <span data-ttu-id="36822-173">Sie können einen beliebigen Betreff und Text eingeben.</span><span class="sxs-lookup"><span data-stu-id="36822-173">You can provide whatever subject and body text you'd like.</span></span>

    :::image type="content" source="../images/power-automate-return-tutorial-5.png" alt-text="Die Power Automate Outlook Connector-Einstellungen zum Senden einer E-Mail. Zu den Optionen gehören die zu sendende Datei, der Betreff der E-Mail und der Textkörper der E-Mail sowie erweiterte Optionen.":::

    > [!NOTE]
    > <span data-ttu-id="36822-p111">In diesem Tutorial wird Outlook verwendet. Sie können stattdessen auch Ihren bevorzugten E-Mail-Dienst verwenden, obwohl einige Optionen anders sein können.</span><span class="sxs-lookup"><span data-stu-id="36822-p111">This tutorial uses Outlook. Feel free to use your preferred email service instead, though some options may be different.</span></span>

1. <span data-ttu-id="36822-177">Wählen Sie **Speichern** aus.</span><span class="sxs-lookup"><span data-stu-id="36822-177">Select **Save**.</span></span>

## <a name="test-the-script-in-power-automate"></a><span data-ttu-id="36822-178">Testen des Skripts in Power Automate</span><span class="sxs-lookup"><span data-stu-id="36822-178">Test the script in Power Automate</span></span>

<span data-ttu-id="36822-179">Ihr Flow wird jeden Montagmorgen ausgeführt.</span><span class="sxs-lookup"><span data-stu-id="36822-179">Your flow will run every Monday morning.</span></span> <span data-ttu-id="36822-180">Sie können das Skript jetzt testen, indem Sie die Schaltfläche **Test** in der oberen rechten Ecke des Bildschirms auswählen.</span><span class="sxs-lookup"><span data-stu-id="36822-180">You can test the script now by selecting the **Test** button in the upper-right corner of the screen.</span></span> <span data-ttu-id="36822-181">Wählen Sie **Manuell** und anschließend **Test ausführen** aus, um den Flow jetzt auszuführen und das Verhalten zu testen.</span><span class="sxs-lookup"><span data-stu-id="36822-181">Select **Manually**, then select **Run Test** to run the flow now and test the behavior.</span></span> <span data-ttu-id="36822-182">Möglicherweise müssen Sie Excel und Outlook Berechtigungen erteilen, um fortzufahren.</span><span class="sxs-lookup"><span data-stu-id="36822-182">You may need to grant permissions to Excel and Outlook to continue.</span></span>

:::image type="content" source="../images/power-automate-return-tutorial-6.png" alt-text="Schaltfläche „Power Automate-Test“.":::

> [!TIP]
> <span data-ttu-id="36822-184">Wenn Ihr Flow keine E-Mail senden kann, vergewissern Sie sich, dass in das Arbeitsblatt eine gültige E-Mail für den aktuellen Datumsbereich oben in der Tabelle aufgeführt ist.</span><span class="sxs-lookup"><span data-stu-id="36822-184">If your flow fails to send an email, double-check in the spreadsheet that a valid email is listed for the current date range at the top of the table.</span></span>

## <a name="next-steps"></a><span data-ttu-id="36822-185">Nächste Schritte</span><span class="sxs-lookup"><span data-stu-id="36822-185">Next steps</span></span>

<span data-ttu-id="36822-186">Besuchen Sie [Ausführen von Office-Skripts mit Power Automate](../develop/power-automate-integration.md), um mehr über das Verbinden von Office-Skripts mit Power Automate zu erfahren.</span><span class="sxs-lookup"><span data-stu-id="36822-186">Visit [Run Office Scripts with Power Automate](../develop/power-automate-integration.md) to learn more about connecting Office Scripts with Power Automate.</span></span>

<span data-ttu-id="36822-187">Sie können sich auch das Beispielszenario [Automatisierte Aufgabenerinnerungen](../resources/scenarios/task-reminders.md) ansehen, um zu erfahren, wie Sie Office-Skripts und Power Automate mit Teams Adaptive Cards kombinieren können.</span><span class="sxs-lookup"><span data-stu-id="36822-187">You can also check out the [Automated task reminders sample scenario](../resources/scenarios/task-reminders.md) to learn how to combine Office Scripts and Power Automate with Teams Adaptive Cards.</span></span>
