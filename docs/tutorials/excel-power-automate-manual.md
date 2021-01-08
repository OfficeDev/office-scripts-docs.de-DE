---
title: Aufrufen von Skripts aus einem manuellen Power Automate-Datenfluss
description: Ein Lernprogramm zur Verwendung von Office-Skripts in Power Automate durch einen manuellen Auslöser.
ms.date: 12/28/2020
localization_priority: Priority
ms.openlocfilehash: 763a61f65afbdbb55346bb3d99f216da33e8419c
ms.sourcegitcommit: 9df67e007ddbfec79a7360df9f4ea5ac6c86fb08
ms.translationtype: HT
ms.contentlocale: de-DE
ms.lasthandoff: 01/06/2021
ms.locfileid: "49772985"
---
# <a name="call-scripts-from-a-manual-power-automate-flow-preview"></a><span data-ttu-id="6d303-103">Aufrufen von Skripts aus einem manuellen Power Automate-Datenfluss (Vorschau)</span><span class="sxs-lookup"><span data-stu-id="6d303-103">Call scripts from a manual Power Automate flow (preview)</span></span>

<span data-ttu-id="6d303-104">In diesem Lernprogramm erfahren Sie, wie Sie ein Office-Skript für Excel im Web durch [Power Automate](https://flow.microsoft.com) ausführen.</span><span class="sxs-lookup"><span data-stu-id="6d303-104">This tutorial teaches you how to run an Office Script for Excel on the web through [Power Automate](https://flow.microsoft.com).</span></span> <span data-ttu-id="6d303-105">Sie erstellen ein Skript, das die Werte von zwei Zellen mit der aktuellen Zeit aktualisiert.</span><span class="sxs-lookup"><span data-stu-id="6d303-105">You'll make a script that updates the values of two cells with the current time.</span></span> <span data-ttu-id="6d303-106">Sie verbinden dieses Skript dann mit einem manuell ausgelösten Power Automate-Flow, so dass das Skript immer dann ausgeführt wird, wenn eine Taste in Power Automate gedrückt wird.</span><span class="sxs-lookup"><span data-stu-id="6d303-106">You'll then connect that script to a manually triggered Power Automate flow, so that the script is run whenever a button in Power Automate is pressed.</span></span> <span data-ttu-id="6d303-107">Sobald Sie das Grundmuster verstanden haben, können Sie den Ablauf auf andere Anwendungen ausweiten und Ihren täglichen Arbeitsablauf stärker automatisieren.</span><span class="sxs-lookup"><span data-stu-id="6d303-107">Once you understand the basic pattern, you can expand the flow to include other applications and automate more of your daily workflow.</span></span>

> [!TIP]
> <span data-ttu-id="6d303-108">Wenn Sie mit Office-Skripten noch nicht vertraut sind, empfehlen wir, mit dem [Aufzeichnen, Bearbeiten und Erstellen von Office-Skripten in Excel im Web](excel-tutorial.md)-Lernprogramm zu beginnen.</span><span class="sxs-lookup"><span data-stu-id="6d303-108">If you are new to Office Scripts, we recommend starting with the [Record, edit, and create Office Scripts in Excel on the web](excel-tutorial.md) tutorial.</span></span> <span data-ttu-id="6d303-109">[Office-Skripts verwenden TypeScript](../overview/code-editor-environment.md), und dieses Lernprogramm richten sich an Anfänger bis Fortgeschrittene mit JavaScript oder TypeScript.</span><span class="sxs-lookup"><span data-stu-id="6d303-109">[Office Scripts use TypeScript](../overview/code-editor-environment.md) and this tutorial is intended for people with beginner to intermediate-level knowledge of JavaScript or TypeScript.</span></span> <span data-ttu-id="6d303-110">Wenn Sie noch nicht mit JavaScript vertraut sind, empfehlen wir Ihnen, mit dem [Mozilla-JavaScript-Lernprogramm](https://developer.mozilla.org/docs/Web/JavaScript/Guide/Introduction) zu beginnen.</span><span class="sxs-lookup"><span data-stu-id="6d303-110">If you're new to JavaScript, we recommend starting with the [Mozilla JavaScript tutorial](https://developer.mozilla.org/docs/Web/JavaScript/Guide/Introduction).</span></span>

## <a name="prerequisites"></a><span data-ttu-id="6d303-111">Voraussetzungen</span><span class="sxs-lookup"><span data-stu-id="6d303-111">Prerequisites</span></span>

[!INCLUDE [Tutorial prerequisites](../includes/power-automate-tutorial-prerequisites.md)]

## <a name="prepare-the-workbook"></a><span data-ttu-id="6d303-112">Vorbereiten der Arbeitsmappe</span><span class="sxs-lookup"><span data-stu-id="6d303-112">Prepare the workbook</span></span>

<span data-ttu-id="6d303-113">Power Automate sollte für den Zugriff auf Arbeitsmappenkomponenten keine [relativen Bezüge](../testing/power-automate-troubleshooting.md#avoid-using-relative-references) wie `Workbook.getActiveWorksheet` verwenden.</span><span class="sxs-lookup"><span data-stu-id="6d303-113">Power Automate shouldn't use [relative references](../testing/power-automate-troubleshooting.md#avoid-using-relative-references) like `Workbook.getActiveWorksheet` to access workbook components.</span></span> <span data-ttu-id="6d303-114">Deshalb benötigen wir eine Arbeitsmappe und ein Arbeitsblatt mit konsistenten Namen, die Power Automate als Referenz verwenden kann.</span><span class="sxs-lookup"><span data-stu-id="6d303-114">So, we need a workbook and worksheet with consistent names that Power Automate can reference.</span></span>

1. <span data-ttu-id="6d303-115">Erstellen Sie eine neue Arbeitsmappe mit dem Namen **MeineArbeitsmappe**.</span><span class="sxs-lookup"><span data-stu-id="6d303-115">Create a new workbook named **MyWorkbook**.</span></span>

2. <span data-ttu-id="6d303-116">Erstellen Sie in der Arbeitsmappe **MeineArbeitsmappe** ein Arbeitsblatt namens **Lernprogramm-Arbeitsblatt**.</span><span class="sxs-lookup"><span data-stu-id="6d303-116">In the **MyWorkbook** workbook, create a worksheet called **TutorialWorksheet**.</span></span>

## <a name="create-an-office-script"></a><span data-ttu-id="6d303-117">Erstellen eines Office-Scripts</span><span class="sxs-lookup"><span data-stu-id="6d303-117">Create an Office Script</span></span>

1. <span data-ttu-id="6d303-118">Wechseln Sie zur Registerkarte **Automatisieren**, und wählen Sie **Alle Skripts** aus.</span><span class="sxs-lookup"><span data-stu-id="6d303-118">Go to the **Automate** tab and select **All Scripts**.</span></span>

2. <span data-ttu-id="6d303-119">Wählen Sie **New Script** aus.</span><span class="sxs-lookup"><span data-stu-id="6d303-119">Select **New Script**.</span></span>

3. <span data-ttu-id="6d303-120">Ersetzen Sie das Standardskript durch das folgende Skript.</span><span class="sxs-lookup"><span data-stu-id="6d303-120">Replace the default script with the following script.</span></span> <span data-ttu-id="6d303-121">Dieses Skript fügt das aktuelle Datum und die aktuelle Uhrzeit zu den ersten beiden Zellen des **Lernprogramm-Arbeitsblatt**-Arbeitsblatts hinzu.</span><span class="sxs-lookup"><span data-stu-id="6d303-121">This script adds the current date and time to the first two cells of the **TutorialWorksheet** worksheet.</span></span>

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

4. <span data-ttu-id="6d303-122">Benennen Sie das Skript in **Datum und Uhrzeit festlegen** um.</span><span class="sxs-lookup"><span data-stu-id="6d303-122">Rename the script to **Set date and time**.</span></span> <span data-ttu-id="6d303-123">Klicken Sie auf den Skriptnamen, um ihn zu ändern.</span><span class="sxs-lookup"><span data-stu-id="6d303-123">Press the script name to change it.</span></span>

5. <span data-ttu-id="6d303-124">Klicken Sie auf **Save Script**, um das Skript zu speichern.</span><span class="sxs-lookup"><span data-stu-id="6d303-124">Save the script by pressing **Save Script**.</span></span>

## <a name="create-an-automated-workflow-with-power-automate"></a><span data-ttu-id="6d303-125">Erstellen eines automatisierten Workflows mit Power Automate</span><span class="sxs-lookup"><span data-stu-id="6d303-125">Create an automated workflow with Power Automate</span></span>

1. <span data-ttu-id="6d303-126">Melden Sie sich an der [Power Automate-Website](https://flow.microsoft.com) an.</span><span class="sxs-lookup"><span data-stu-id="6d303-126">Sign in to the [Power Automate site](https://flow.microsoft.com).</span></span>

2. <span data-ttu-id="6d303-127">Klicken Sie in dem Menü, das auf der linken Seite des Bildschirms angezeigt wird, auf **Create**.</span><span class="sxs-lookup"><span data-stu-id="6d303-127">In the menu that's displayed on the left side of the screen, press **Create**.</span></span> <span data-ttu-id="6d303-128">Damit gelangen Sie zur Liste der Möglichkeiten zum Erstellen neuer Workflows.</span><span class="sxs-lookup"><span data-stu-id="6d303-128">This brings you to list of ways to create new workflows.</span></span>

    ![Die Schaltfläche „Erstellen“ in Power Automate](../images/power-automate-tutorial-1.png)

3. <span data-ttu-id="6d303-130">Wählen Sie im Abschnitt **Start from blank** die Option **Instant flow** aus.</span><span class="sxs-lookup"><span data-stu-id="6d303-130">In the **Start from blank** section, select **Instant flow**.</span></span> <span data-ttu-id="6d303-131">Dadurch wird ein manuell aktivierter Workflow erstellt.</span><span class="sxs-lookup"><span data-stu-id="6d303-131">This creates a manually activated workflow.</span></span>

    ![Die Option „Direktflow“ zum Erstellen eines neuen Workflows](../images/power-automate-tutorial-2.png)

4. <span data-ttu-id="6d303-133">Geben Sie im daraufhin angezeigten Dialogfenster einen Namen für den Flow in das Textfeld **Flow name** ein, wählen Sie in der Liste der Optionen unter **Choose how to trigger the flow** die Option **Manually trigger a flow** aus, und klicken Sie auf **Create**.</span><span class="sxs-lookup"><span data-stu-id="6d303-133">In the dialog window that appears, enter a name for your flow in the **Flow name** text box, select **Manually trigger a flow** from the list of options under **Choose how to trigger the flow**, and press **Create**.</span></span>

    ![Die Option „Manueller Trigger“ zum Erstellen eines neuen Direktflows](../images/power-automate-tutorial-3.png)

    <span data-ttu-id="6d303-135">Beachten Sie, dass ein manuell ausgelöster Flow nur einer von vielen Arten von Flows ist.</span><span class="sxs-lookup"><span data-stu-id="6d303-135">Note that a manually triggered flow is just one of many types of flows.</span></span> <span data-ttu-id="6d303-136">Im nächsten Lernprogramm erstellen Sie einen Flow, der automatisch ausgeführt wird, wenn Sie eine E-Mail erhalten.</span><span class="sxs-lookup"><span data-stu-id="6d303-136">In the next tutorial, you'll make a flow that automatically runs when you receive an email.</span></span>

5. <span data-ttu-id="6d303-137">Klicken Sie auf **New step**.</span><span class="sxs-lookup"><span data-stu-id="6d303-137">Press **New step**.</span></span>

6. <span data-ttu-id="6d303-138">Wählen Sie die Registerkarte **Standard** aus, und wählen Sie dann **Excel Online (Business)** aus.</span><span class="sxs-lookup"><span data-stu-id="6d303-138">Select the **Standard** tab, then select **Excel Online (Business)**.</span></span>

    ![Excel Online (Business)-Option in Power Automate](../images/power-automate-tutorial-4.png)

7. <span data-ttu-id="6d303-140">Wählen Sie unter **Actions** die Option **Run script (preview)** aus.</span><span class="sxs-lookup"><span data-stu-id="6d303-140">Under **Actions**, select **Run script (preview)**.</span></span>

    ![Aktionsoption „Skript ausführen (Vorschau)“ in Power Automate](../images/power-automate-tutorial-5.png)

8. <span data-ttu-id="6d303-142">Als nächstes wählen Sie die Arbeitsmappe und das Skript aus, die im Ablaufschritt verwendet werden sollen.</span><span class="sxs-lookup"><span data-stu-id="6d303-142">Next, you'll select the workbook and script to use in the flow step.</span></span> <span data-ttu-id="6d303-143">Für das Tutorial verwenden Sie die Arbeitsmappe, die Sie in Ihrem OneDrive erstellt haben. Sie können auch eine beliebige Arbeitsmappe in einer OneDrive- oder SharePoint-Website verwenden.</span><span class="sxs-lookup"><span data-stu-id="6d303-143">For the tutorial, you'll use the workbook you created in your OneDrive, but you could use any workbook in a OneDrive or SharePoint site.</span></span> <span data-ttu-id="6d303-144">Geben Sie die folgenden Einstellungen für den Konnektor **Run script** an:</span><span class="sxs-lookup"><span data-stu-id="6d303-144">Specify the following settings for the **Run script** connector:</span></span>

    - <span data-ttu-id="6d303-145">**Location**: OneDrive for Business</span><span class="sxs-lookup"><span data-stu-id="6d303-145">**Location**: OneDrive for Business</span></span>
    - <span data-ttu-id="6d303-146">**Document Library**: OneDrive</span><span class="sxs-lookup"><span data-stu-id="6d303-146">**Document Library**: OneDrive</span></span>
    - <span data-ttu-id="6d303-147">**File**: MyWorkbook.xlsx *(Ausgewählt über den Dateibrowser)*</span><span class="sxs-lookup"><span data-stu-id="6d303-147">**File**: MyWorkbook.xlsx *(Chosen through the file browser)*</span></span>
    - <span data-ttu-id="6d303-148">**Script**: Datum und Uhrzeit festlegen</span><span class="sxs-lookup"><span data-stu-id="6d303-148">**Script**: Set date and time</span></span>

    ![Die Connectoreinstellungen zum Ausführen eines Skripts in Power Automate](../images/power-automate-tutorial-6.png)

9. <span data-ttu-id="6d303-150">Klicken Sie auf **Save**.</span><span class="sxs-lookup"><span data-stu-id="6d303-150">Press **Save**.</span></span>

<span data-ttu-id="6d303-151">Jetzt kann ihr Flow über Power Automate ausgeführt werden.</span><span class="sxs-lookup"><span data-stu-id="6d303-151">Your flow is now ready to be run through Power Automate.</span></span> <span data-ttu-id="6d303-152">Sie können ihn mithilfe der Schaltfläche **Test** im Flow-Editor testen oder die weiteren Lernprogrammschritte zum Ausführen des Flows aus der Flow-Sammlung ausführen.</span><span class="sxs-lookup"><span data-stu-id="6d303-152">You can test it using the **Test** button in the flow editor or follow the remaining tutorial steps to run the flow from your flow collection.</span></span>

## <a name="run-the-script-through-power-automate"></a><span data-ttu-id="6d303-153">Ausführen des Skripts über Power Automate</span><span class="sxs-lookup"><span data-stu-id="6d303-153">Run the script through Power Automate</span></span>

1. <span data-ttu-id="6d303-154">Wählen Sie auf der Hauptseite der Power Automate-Seite **My Flows** aus.</span><span class="sxs-lookup"><span data-stu-id="6d303-154">From the main Power Automate page, select **My flows**.</span></span>

    ![Die Schaltfläche „Meine Flows“ in Power Automate](../images/power-automate-tutorial-7.png)

2. <span data-ttu-id="6d303-156">Wählen Sie in der Liste der Flows, die auf der Registerkarte **My Flows** angezeigt werden, **Mein Lernprogramm-Flow** aus. Die Details des zuvor erstellten Flows werden angezeigt.</span><span class="sxs-lookup"><span data-stu-id="6d303-156">Select **My tutorial flow** from the list of flows displayed in the **My flows** tab. This shows the details of the flow we previously created.</span></span>

3. <span data-ttu-id="6d303-157">Klicken Sie auf **Run**.</span><span class="sxs-lookup"><span data-stu-id="6d303-157">Press **Run**.</span></span>

    ![Die Schaltfläche „Ausführen“ in Power Automate](../images/power-automate-tutorial-8.png)

4. <span data-ttu-id="6d303-159">Für die Ausführung des Flows wird ein Aufgabenbereich angezeigt.</span><span class="sxs-lookup"><span data-stu-id="6d303-159">A task pane will appear for running the flow.</span></span> <span data-ttu-id="6d303-160">Wenn Sie aufgefordert werden, sich bei Excel Online **anzumelden** klicken Sie auf **Continue**.</span><span class="sxs-lookup"><span data-stu-id="6d303-160">If you are asked to **Sign in** to Excel Online, do so by pressing **Continue**.</span></span>

5. <span data-ttu-id="6d303-161">Klicken Sie auf **Run flow**.</span><span class="sxs-lookup"><span data-stu-id="6d303-161">Press **Run flow**.</span></span> <span data-ttu-id="6d303-162">Damit wird der Flow ausgeführt, der das zugehörige Office-Skript ausführt.</span><span class="sxs-lookup"><span data-stu-id="6d303-162">This runs the flow, which runs the related Office Script.</span></span>

6. <span data-ttu-id="6d303-163">Klicken Sie auf **Done**.</span><span class="sxs-lookup"><span data-stu-id="6d303-163">Press **Done**.</span></span> <span data-ttu-id="6d303-164">Der Abschnitt **Runs** wird entsprechend aktualisiert.</span><span class="sxs-lookup"><span data-stu-id="6d303-164">You should see the **Runs** section update accordingly.</span></span>

7. <span data-ttu-id="6d303-165">Aktualisieren Sie die Seite, um die Ergebnisse von Power Automate anzuzeigen.</span><span class="sxs-lookup"><span data-stu-id="6d303-165">Refresh the page to see the results of the Power Automate.</span></span> <span data-ttu-id="6d303-166">Wenn der Vorgang erfolgreich war, wechseln Sie zur Arbeitsmappe, um die aktualisierten Zellen anzuzeigen.</span><span class="sxs-lookup"><span data-stu-id="6d303-166">If it succeeded, go to the workbook to see the updated cells.</span></span> <span data-ttu-id="6d303-167">Falls ein Fehler aufgetreten ist, überprüfen Sie die Einstellungen des Flows, und führen Sie ihn ein zweites Mal aus.</span><span class="sxs-lookup"><span data-stu-id="6d303-167">If it failed, verify the flow's settings and run it a second time.</span></span>

    ![Power Automate-Ausgabe nach einer erfolgreichen Flowausführung](../images/power-automate-tutorial-9.png)

## <a name="next-steps"></a><span data-ttu-id="6d303-169">Nächste Schritte</span><span class="sxs-lookup"><span data-stu-id="6d303-169">Next steps</span></span>

<span data-ttu-id="6d303-170">Schließen Sie das Lernprogramm [Übergeben von Daten zu Skripts in einem automatisch ausgeführten Power Automate-Datenfluss](excel-power-automate-trigger.md) ab.</span><span class="sxs-lookup"><span data-stu-id="6d303-170">Complete the [Pass data to scripts in an automatically-run Power Automate flow](excel-power-automate-trigger.md) tutorial.</span></span> <span data-ttu-id="6d303-171">Hier erfahren Sie, wie Sie Daten aus einem Workflowdienst an Ihr Office-Skript übergeben und den Power Automate-Flow ausführen, wenn bestimmte Ereignisse eintreten.</span><span class="sxs-lookup"><span data-stu-id="6d303-171">It teaches you how to pass data from a workflow service to your Office Script and run the Power Automate flow when certain events occur.</span></span>
