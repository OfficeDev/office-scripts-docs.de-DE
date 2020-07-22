---
title: Aufrufen von Skripts aus einem manuellen Power Automate-Datenfluss
description: Ein Lernprogramm zur Verwendung von Office-Skripts in Power Automate durch einen manuellen Auslöser.
ms.date: 07/14/2020
localization_priority: Priority
ms.openlocfilehash: 70fca2620973ecefe9eda40f02e28f064b713677
ms.sourcegitcommit: ebd1079c7e2695ac0e7e4c616f2439975e196875
ms.translationtype: HT
ms.contentlocale: de-DE
ms.lasthandoff: 07/17/2020
ms.locfileid: "45160436"
---
# <a name="call-scripts-from-a-manual-power-automate-flow-preview"></a><span data-ttu-id="cf129-103">Aufrufen von Skripts aus einem manuellen Power Automate-Datenfluss (Vorschau)</span><span class="sxs-lookup"><span data-stu-id="cf129-103">Call scripts from a manual Power Automate flow (preview)</span></span>

<span data-ttu-id="cf129-104">In diesem Lernprogramm erfahren Sie, wie Sie ein Office-Skript für Excel im Web durch [Power Automate](https://flow.microsoft.com) ausführen.</span><span class="sxs-lookup"><span data-stu-id="cf129-104">This tutorial teaches you how to run an Office Script for Excel on the web through [Power Automate](https://flow.microsoft.com).</span></span>

## <a name="prerequisites"></a><span data-ttu-id="cf129-105">Voraussetzungen</span><span class="sxs-lookup"><span data-stu-id="cf129-105">Prerequisites</span></span>

[!INCLUDE [Tutorial prerequisites](../includes/power-automate-tutorial-prerequisites.md)]

> [!IMPORTANT]
> <span data-ttu-id="cf129-106">In diesem Lernprogramm wird davon ausgegangen, dass Sie das Lernprogramm [Aufzeichnen, Bearbeiten und Erstellen von Office-Skripts in Excel im Web](excel-tutorial.md) abgeschlossen haben.</span><span class="sxs-lookup"><span data-stu-id="cf129-106">This tutorial assumes you have completed the [Record, edit, and create Office Scripts in Excel on the web](excel-tutorial.md) tutorial.</span></span>

## <a name="prepare-the-workbook"></a><span data-ttu-id="cf129-107">Vorbereiten der Arbeitsmappe</span><span class="sxs-lookup"><span data-stu-id="cf129-107">Prepare the workbook</span></span>

<span data-ttu-id="cf129-108">Power Automation kann relative Bezüge wie `Workbook.getActiveWorksheet` nicht verwenden, um auf Arbeitsmappenfunktionen zuzugreifen.</span><span class="sxs-lookup"><span data-stu-id="cf129-108">Power Automate can't use relative references like `Workbook.getActiveWorksheet` to access workbook components.</span></span> <span data-ttu-id="cf129-109">Deshalb benötigen wir eine Arbeitsmappe und ein Arbeitsblatt mit konsistenten Namen, die Power Automate als Referenz verwenden kann.</span><span class="sxs-lookup"><span data-stu-id="cf129-109">So, we need a workbook and worksheet with consistent names that Power Automate can reference.</span></span>

1. <span data-ttu-id="cf129-110">Erstellen Sie eine neue Arbeitsmappe mit dem Namen **MeineArbeitsmappe**.</span><span class="sxs-lookup"><span data-stu-id="cf129-110">Create a new workbook named **MyWorkbook**.</span></span>

2. <span data-ttu-id="cf129-111">Erstellen Sie in der Arbeitsmappe **MeineArbeitsmappe** ein Arbeitsblatt namens **Lernprogramm-Arbeitsblatt**.</span><span class="sxs-lookup"><span data-stu-id="cf129-111">In the **MyWorkbook** workbook, create a worksheet called **TutorialWorksheet**.</span></span>

## <a name="create-an-office-script"></a><span data-ttu-id="cf129-112">Erstellen eines Office-Scripts</span><span class="sxs-lookup"><span data-stu-id="cf129-112">Create an Office Script</span></span>

1. <span data-ttu-id="cf129-113">Wechseln Sie zur Registerkarte **Automate**, und wählen Sie **Code Editor** aus.</span><span class="sxs-lookup"><span data-stu-id="cf129-113">Go to the **Automate** tab and select **Code Editor**.</span></span>

2. <span data-ttu-id="cf129-114">Wählen Sie **New Script** aus.</span><span class="sxs-lookup"><span data-stu-id="cf129-114">Select **New Script**.</span></span>

3. <span data-ttu-id="cf129-115">Ersetzen Sie das Standardskript durch das folgende Skript.</span><span class="sxs-lookup"><span data-stu-id="cf129-115">Replace the default script with the following script.</span></span> <span data-ttu-id="cf129-116">Dieses Skript fügt das aktuelle Datum und die aktuelle Uhrzeit zu den ersten beiden Zellen des **Lernprogramm-Arbeitsblatt**-Arbeitsblatts hinzu.</span><span class="sxs-lookup"><span data-stu-id="cf129-116">This script adds the current date and time to the first two cells of the **TutorialWorksheet** worksheet.</span></span>

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

4. <span data-ttu-id="cf129-117">Benennen Sie das Skript in **Datum und Uhrzeit festlegen** um.</span><span class="sxs-lookup"><span data-stu-id="cf129-117">Rename the script to **Set date and time**.</span></span> <span data-ttu-id="cf129-118">Klicken Sie auf den Skriptnamen, um ihn zu ändern.</span><span class="sxs-lookup"><span data-stu-id="cf129-118">Press the script name to change it.</span></span>

5. <span data-ttu-id="cf129-119">Klicken Sie auf **Save Script**, um das Skript zu speichern.</span><span class="sxs-lookup"><span data-stu-id="cf129-119">Save the script by pressing **Save Script**.</span></span>

## <a name="create-an-automated-workflow-with-power-automate"></a><span data-ttu-id="cf129-120">Erstellen eines automatisierten Workflows mit Power Automate</span><span class="sxs-lookup"><span data-stu-id="cf129-120">Create an automated workflow with Power Automate</span></span>

1. <span data-ttu-id="cf129-121">Melden Sie sich an der [Power Automate-Website](https://flow.microsoft.com) an.</span><span class="sxs-lookup"><span data-stu-id="cf129-121">Sign in to the [Power Automate site](https://flow.microsoft.com).</span></span>

2. <span data-ttu-id="cf129-122">Klicken Sie in dem Menü, das auf der linken Seite des Bildschirms angezeigt wird, auf **Create**.</span><span class="sxs-lookup"><span data-stu-id="cf129-122">In the menu that's displayed on the left side of the screen, press **Create**.</span></span> <span data-ttu-id="cf129-123">Damit gelangen Sie zur Liste der Möglichkeiten zum Erstellen neuer Workflows.</span><span class="sxs-lookup"><span data-stu-id="cf129-123">This brings you to list of ways to create new workflows.</span></span>

    ![Die Schaltfläche „Erstellen“ in Power Automate.](../images/power-automate-tutorial-1.png)

3. <span data-ttu-id="cf129-125">Wählen Sie im Abschnitt **Start from blank** die Option **Instant flow** aus.</span><span class="sxs-lookup"><span data-stu-id="cf129-125">In the **Start from blank** section, select **Instant flow**.</span></span> <span data-ttu-id="cf129-126">Dadurch wird ein manuell aktivierter Workflow erstellt.</span><span class="sxs-lookup"><span data-stu-id="cf129-126">This creates a manually activated workflow.</span></span>

    ![Die Option „Instant flow“ zum Erstellen eines neuen Workflows.](../images/power-automate-tutorial-2.png)

4. <span data-ttu-id="cf129-128">Geben Sie im daraufhin angezeigten Dialogfenster einen Namen für den Flow in das Textfeld **Flow name** ein, wählen Sie in der Liste der Optionen unter **Choose how to trigger the flow** die Option **Manually trigger a flow** aus, und klicken Sie auf **Create**.</span><span class="sxs-lookup"><span data-stu-id="cf129-128">In the dialog window that appears, enter a name for your flow in the **Flow name** text box, select **Manually trigger a flow** from the list of options under **Choose how to trigger the flow**, and press **Create**.</span></span>

    ![Die manuelle Auslöseroption zum Erstellen eines neuen „Instant flow“.](../images/power-automate-tutorial-3.png)

    <span data-ttu-id="cf129-130">Beachten Sie, dass ein manuell ausgelöster Flow nur einer von vielen Arten von Flows ist.</span><span class="sxs-lookup"><span data-stu-id="cf129-130">Note that a manually triggered flow is just one of many types of flows.</span></span> <span data-ttu-id="cf129-131">Im nächsten Lernprogramm erstellen Sie einen Flow, der automatisch ausgeführt wird, wenn Sie eine E-Mail erhalten.</span><span class="sxs-lookup"><span data-stu-id="cf129-131">In the next tutorial, you'll make a flow that automatically runs when you receive an email.</span></span>

5. <span data-ttu-id="cf129-132">Klicken Sie auf **New step**.</span><span class="sxs-lookup"><span data-stu-id="cf129-132">Press **New step**.</span></span>

6. <span data-ttu-id="cf129-133">Wählen Sie die Registerkarte **Standard** aus, und wählen Sie dann **Excel Online (Business)** aus.</span><span class="sxs-lookup"><span data-stu-id="cf129-133">Select the **Standard** tab, then select **Excel Online (Business)**.</span></span>

    ![Die Power Automate-Option für Excel Online (Business).](../images/power-automate-tutorial-4.png)

7. <span data-ttu-id="cf129-135">Wählen Sie unter **Actions** die Option **Run script (preview)** aus.</span><span class="sxs-lookup"><span data-stu-id="cf129-135">Under **Actions**, select **Run script (preview)**.</span></span>

    ![Die Power Automate-Aktionsoption für „Run script (preview)“.](../images/power-automate-tutorial-5.png)

8. <span data-ttu-id="cf129-137">Geben Sie die folgenden Einstellungen für den Konnektor **Run script** an:</span><span class="sxs-lookup"><span data-stu-id="cf129-137">Specify the following settings for the **Run script** connector:</span></span>

    - <span data-ttu-id="cf129-138">**Location**: OneDrive for Business</span><span class="sxs-lookup"><span data-stu-id="cf129-138">**Location**: OneDrive for Business</span></span>
    - <span data-ttu-id="cf129-139">**Document Library**: OneDrive</span><span class="sxs-lookup"><span data-stu-id="cf129-139">**Document Library**: OneDrive</span></span>
    - <span data-ttu-id="cf129-140">**File**: MeineArbeitsmappe. xlsx</span><span class="sxs-lookup"><span data-stu-id="cf129-140">**File**: MyWorkbook.xlsx</span></span>
    - <span data-ttu-id="cf129-141">**Script**: Datum und Uhrzeit festlegen</span><span class="sxs-lookup"><span data-stu-id="cf129-141">**Script**: Set date and time</span></span>

    ![Die Konnektoreinstellungen zum Ausführen eines Skripts in Power Automate.](../images/power-automate-tutorial-6.png)

9. <span data-ttu-id="cf129-143">Klicken Sie auf **Save**.</span><span class="sxs-lookup"><span data-stu-id="cf129-143">Press **Save**.</span></span>

<span data-ttu-id="cf129-144">Jetzt kann ihr Flow über Power Automate ausgeführt werden.</span><span class="sxs-lookup"><span data-stu-id="cf129-144">Your flow is now ready to be run through Power Automate.</span></span> <span data-ttu-id="cf129-145">Sie können ihn mithilfe der Schaltfläche **Test** im Flow-Editor testen oder die weiteren Lernprogrammschritte zum Ausführen des Flows aus der Flow-Sammlung ausführen.</span><span class="sxs-lookup"><span data-stu-id="cf129-145">You can test it using the **Test** button in the flow editor or follow the remaining tutorial steps to run the flow from your flow collection.</span></span>

## <a name="run-the-script-through-power-automate"></a><span data-ttu-id="cf129-146">Ausführen des Skripts über Power Automate</span><span class="sxs-lookup"><span data-stu-id="cf129-146">Run the script through Power Automate</span></span>

1. <span data-ttu-id="cf129-147">Wählen Sie auf der Hauptseite der Power Automate-Seite **My Flows** aus.</span><span class="sxs-lookup"><span data-stu-id="cf129-147">From the main Power Automate page, select **My flows**.</span></span>

    ![Die Schaltfläche „My Flows“ in Power Automate.](../images/power-automate-tutorial-7.png)

2. <span data-ttu-id="cf129-149">Wählen Sie in der Liste der Flows, die auf der Registerkarte **My Flows** angezeigt werden, **Mein Lernprogramm-Flow** aus. Die Details des zuvor erstellten Flows werden angezeigt.</span><span class="sxs-lookup"><span data-stu-id="cf129-149">Select **My tutorial flow** from the list of flows displayed in the **My flows** tab. This shows the details of the flow we previously created.</span></span>

3. <span data-ttu-id="cf129-150">Klicken Sie auf **Run**.</span><span class="sxs-lookup"><span data-stu-id="cf129-150">Press **Run**.</span></span>

    ![Die Schaltfläche „Run“ in Power Automate.](../images/power-automate-tutorial-8.png)

4. <span data-ttu-id="cf129-152">Für die Ausführung des Flows wird ein Aufgabenbereich angezeigt.</span><span class="sxs-lookup"><span data-stu-id="cf129-152">A task pane will appear for running the flow.</span></span> <span data-ttu-id="cf129-153">Wenn Sie aufgefordert werden, sich bei Excel Online **anzumelden** klicken Sie auf **Continue**.</span><span class="sxs-lookup"><span data-stu-id="cf129-153">If you are asked to **Sign in** to Excel Online, do so by pressing **Continue**.</span></span>

5. <span data-ttu-id="cf129-154">Klicken Sie auf **Run flow**.</span><span class="sxs-lookup"><span data-stu-id="cf129-154">Press **Run flow**.</span></span> <span data-ttu-id="cf129-155">Damit wird der Flow ausgeführt, der das zugehörige Office-Skript ausführt.</span><span class="sxs-lookup"><span data-stu-id="cf129-155">This runs the flow, which runs the related Office Script.</span></span>

6. <span data-ttu-id="cf129-156">Klicken Sie auf **Done**.</span><span class="sxs-lookup"><span data-stu-id="cf129-156">Press **Done**.</span></span> <span data-ttu-id="cf129-157">Der Abschnitt **Runs** wird entsprechend aktualisiert.</span><span class="sxs-lookup"><span data-stu-id="cf129-157">You should see the **Runs** section update accordingly.</span></span>

7. <span data-ttu-id="cf129-158">Aktualisieren Sie die Seite, um die Ergebnisse von Power Automate anzuzeigen.</span><span class="sxs-lookup"><span data-stu-id="cf129-158">Refresh the page to see the results of the Power Automate.</span></span> <span data-ttu-id="cf129-159">Wenn der Vorgang erfolgreich war, wechseln Sie zur Arbeitsmappe, um die aktualisierten Zellen anzuzeigen.</span><span class="sxs-lookup"><span data-stu-id="cf129-159">If it succeeded, go to the workbook to see the updated cells.</span></span> <span data-ttu-id="cf129-160">Falls ein Fehler aufgetreten ist, überprüfen Sie die Einstellungen des Flows, und führen Sie ihn ein zweites Mal aus.</span><span class="sxs-lookup"><span data-stu-id="cf129-160">If it failed, verify the flow's settings and run it a second time.</span></span>

    ![Power Automate-Ausgabe nach einer erfolgreichen Flow-Ausführung.](../images/power-automate-tutorial-9.png)

## <a name="next-steps"></a><span data-ttu-id="cf129-162">Nächste Schritte</span><span class="sxs-lookup"><span data-stu-id="cf129-162">Next steps</span></span>

<span data-ttu-id="cf129-163">Schließen Sie das Lernprogramm [Übergeben von Daten zu Skripts in einem automatisch ausgeführten Power Automate-Datenfluss](excel-power-automate-trigger.md) ab.</span><span class="sxs-lookup"><span data-stu-id="cf129-163">Complete the [Pass data to scripts in an automatically-run Power Automate flow](excel-power-automate-trigger.md) tutorial.</span></span> <span data-ttu-id="cf129-164">Hier erfahren Sie, wie Sie Daten aus einem Workflowdienst an Ihr Office-Skript übergeben und den Power Automate-Flow ausführen, wenn bestimmte Ereignisse eintreten.</span><span class="sxs-lookup"><span data-stu-id="cf129-164">It teaches you how to pass data from a workflow service to your Office Script and run the Power Automate flow when certain events occur.</span></span>
