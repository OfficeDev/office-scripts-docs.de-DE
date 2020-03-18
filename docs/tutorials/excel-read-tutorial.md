---
title: Lesen Sie Arbeitsmappendaten mit Office-Skripts in Excel im Web
description: Ein Office Skripts-Lernprogramm zum Lesen von Daten aus Arbeitsmappen und zum Auswerten dieser Daten im Skript.
ms.date: 01/27/2020
localization_priority: Priority
ms.openlocfilehash: 42ed0fe5843a78692f9660b873211e3668702164
ms.sourcegitcommit: b075eed5a6f275274fbbf6d62633219eac416f26
ms.translationtype: HT
ms.contentlocale: de-DE
ms.lasthandoff: 03/10/2020
ms.locfileid: "42700182"
---
# <a name="read-workbook-data-with-office-scripts-in-excel-on-the-web"></a><span data-ttu-id="00f70-103">Lesen Sie Arbeitsmappendaten mit Office-Skripts in Excel im Web</span><span class="sxs-lookup"><span data-stu-id="00f70-103">Read workbook data with Office Scripts in Excel on the web</span></span>

<span data-ttu-id="00f70-104">In diesem Lernprogramm erfahren Sie, wie Sie Daten aus einer Arbeitsmappe mit einem Office-Skript für Excel im Web lesen.</span><span class="sxs-lookup"><span data-stu-id="00f70-104">This tutorial will teach you how to read data from a workbook with an Office Script for Excel on the web.</span></span> <span data-ttu-id="00f70-105">Anschließend bearbeiten Sie die gelesenen Daten und fügen sie wieder in die Arbeitsmappe ein.</span><span class="sxs-lookup"><span data-stu-id="00f70-105">You'll then edit the data you read and put it back in the workbook.</span></span>

> [!TIP]
> <span data-ttu-id="00f70-106">Wenn Sie mit Office-Skripten noch nicht vertraut sind, empfehlen wir, mit dem [Aufzeichnen, Bearbeiten und Erstellen von Office-Skripten in Excel im Web](excel-tutorial.md)-Lernprogramm zu beginnen.</span><span class="sxs-lookup"><span data-stu-id="00f70-106">If you are new to Office Scripts, we recommend starting with the [Record, edit, and create Office Scripts in Excel on the web](excel-tutorial.md) tutorial.</span></span>

## <a name="prerequisites"></a><span data-ttu-id="00f70-107">Voraussetzungen</span><span class="sxs-lookup"><span data-stu-id="00f70-107">Prerequisites</span></span>

[!INCLUDE [Preview note](../includes/preview-note.md)]

<span data-ttu-id="00f70-108">Bevor Sie mit diesem Lernprogramm beginnen, benötigen Sie Zugriff auf Office-Skripts. Dies setzt Folgendes voraus:</span><span class="sxs-lookup"><span data-stu-id="00f70-108">Before starting this tutorial, you'll need access to Office Scripts, which requires the following:</span></span>

- <span data-ttu-id="00f70-109">[Excel im Web](https://www.office.com/launch/excel).</span><span class="sxs-lookup"><span data-stu-id="00f70-109">[Excel on the web](https://www.office.com/launch/excel).</span></span>
- <span data-ttu-id="00f70-110">Bitten Sie Ihren Administrator, [Office-Skripts für Ihre Organisation zu aktivieren](https://support.office.com/article/office-scripts-settings-in-m365-19d3c51a-6ca2-40ab-978d-60fa49554dcf). Dadurch wird die Registerkarte **Automatisieren** zum Menüband hinzugefügt.</span><span class="sxs-lookup"><span data-stu-id="00f70-110">Ask your administrator to [enable Office Scripts for your organization](https://support.office.com/article/office-scripts-settings-in-m365-19d3c51a-6ca2-40ab-978d-60fa49554dcf), which adds the **Automate** tab to the ribbon.</span></span>

> [!IMPORTANT]
> <span data-ttu-id="00f70-111">Dieses Lernprogramm richtet sich an Anfänger bis Fortgeschrittene mit JavaScript oder TypeScript.</span><span class="sxs-lookup"><span data-stu-id="00f70-111">This tutorial is intended for people with beginner to intermediate-level knowledge of JavaScript or TypeScript.</span></span> <span data-ttu-id="00f70-112">Wenn Sie noch nicht mit JavaScript vertraut sind, empfehlen wir Ihnen, das [Mozilla-JavaScript-Lernprogramm](https://developer.mozilla.org/docs/Web/JavaScript/Guide/Introduction) durchzusehen.</span><span class="sxs-lookup"><span data-stu-id="00f70-112">If you're new to JavaScript, we recommend reviewing the [Mozilla JavaScript tutorial](https://developer.mozilla.org/docs/Web/JavaScript/Guide/Introduction).</span></span> <span data-ttu-id="00f70-113">Weitere Informationen über die Skriptumgebung finden Sie unter [Office-Skripts in Excel im Web](../overview/excel.md).</span><span class="sxs-lookup"><span data-stu-id="00f70-113">Visit [Office Scripts in Excel on the web](../overview/excel.md) to learn more about the script environment.</span></span>

## <a name="read-a-cell"></a><span data-ttu-id="00f70-114">Lesen einer Zelle</span><span class="sxs-lookup"><span data-stu-id="00f70-114">Read a cell</span></span>

<span data-ttu-id="00f70-115">Mit dem Action Recorder erstellte Skripte können nur Informationen in die Arbeitsmappe schreiben.</span><span class="sxs-lookup"><span data-stu-id="00f70-115">Scripts made with the Action Recorder can only write information to the workbook.</span></span> <span data-ttu-id="00f70-116">Mit dem Code-Editor können Sie Skripte bearbeiten und erstellen, die auch Daten aus einer Arbeitsmappe lesen.</span><span class="sxs-lookup"><span data-stu-id="00f70-116">With the Code Editor, you can edit and make scripts that also read data from a workbook.</span></span>

<span data-ttu-id="00f70-117">Lassen Sie uns ein Skript erstellen, das Daten liest und basierend auf dem Gelesenen agiert.</span><span class="sxs-lookup"><span data-stu-id="00f70-117">Let's make a script that reads data and acts based on what was read.</span></span> <span data-ttu-id="00f70-118">Wir werden mit einem Muster-Kontoauszug arbeiten.</span><span class="sxs-lookup"><span data-stu-id="00f70-118">We're going to work with a sample banking statement.</span></span> <span data-ttu-id="00f70-119">Diese Erklärung ist ein kombinierter Prüfungs- und Gutschrift-Kontoauszug.</span><span class="sxs-lookup"><span data-stu-id="00f70-119">This statement is a combined checking and credit statement.</span></span> <span data-ttu-id="00f70-120">Leider melden sie Bilanzänderungen unterschiedlich.</span><span class="sxs-lookup"><span data-stu-id="00f70-120">Unfortunately, they report balance changes differently.</span></span> <span data-ttu-id="00f70-121">Der Prüfungskontoauszug gibt die Einnahmen als positive Gutschrift und die Kosten als negative Belastung an.</span><span class="sxs-lookup"><span data-stu-id="00f70-121">The checking statement gives income as positive credit and costs as negative debit.</span></span> <span data-ttu-id="00f70-122">Der Gutschrift-Kontoauszug macht das Gegenteil.</span><span class="sxs-lookup"><span data-stu-id="00f70-122">The credit statement does the opposite.</span></span>

<span data-ttu-id="00f70-123">Im weiteren Verlauf des Lernprogramms werden diese Daten mithilfe eines Skripts normalisiert.</span><span class="sxs-lookup"><span data-stu-id="00f70-123">Over the rest of the tutorial, we will normalize this data using a script.</span></span> <span data-ttu-id="00f70-124">Lassen Sie uns zunächst lernen, wie Sie Daten aus der Arbeitsmappe lesen.</span><span class="sxs-lookup"><span data-stu-id="00f70-124">First, let's learn how to read data from the workbook.</span></span>

1. <span data-ttu-id="00f70-125">Erstellen Sie ein neues Arbeitsblatt in der Arbeitsmappe, die Sie für den Rest des Lernprogramms verwendet haben.</span><span class="sxs-lookup"><span data-stu-id="00f70-125">Create a new worksheet in the workbook you've used for the rest of the tutorial.</span></span>
2. <span data-ttu-id="00f70-126">Kopieren Sie die folgenden Daten und fügen Sie sie ab Zelle **A1** in das neue Arbeitsblatt ein.</span><span class="sxs-lookup"><span data-stu-id="00f70-126">Copy the following data and paste it into the new worksheet, starting at cell **A1**.</span></span>

    |<span data-ttu-id="00f70-127">Datum</span><span class="sxs-lookup"><span data-stu-id="00f70-127">Date</span></span> |<span data-ttu-id="00f70-128">Konto</span><span class="sxs-lookup"><span data-stu-id="00f70-128">Account</span></span> |<span data-ttu-id="00f70-129">Beschreibung</span><span class="sxs-lookup"><span data-stu-id="00f70-129">Description</span></span> |<span data-ttu-id="00f70-130">Lastschrift</span><span class="sxs-lookup"><span data-stu-id="00f70-130">Debit</span></span> |<span data-ttu-id="00f70-131">Guthaben</span><span class="sxs-lookup"><span data-stu-id="00f70-131">Credit</span></span> |
    |:--|:--|:--|:--|:--|
    |<span data-ttu-id="00f70-132">10.10.2019</span><span class="sxs-lookup"><span data-stu-id="00f70-132">10/10/2019</span></span> |<span data-ttu-id="00f70-133">Wird geprüft</span><span class="sxs-lookup"><span data-stu-id="00f70-133">Checking</span></span> |<span data-ttu-id="00f70-134">Coho Vineyard</span><span class="sxs-lookup"><span data-stu-id="00f70-134">Coho Vineyard</span></span> |<span data-ttu-id="00f70-135">-20.05</span><span class="sxs-lookup"><span data-stu-id="00f70-135">-20.05</span></span> | |
    |<span data-ttu-id="00f70-136">11.10.2019</span><span class="sxs-lookup"><span data-stu-id="00f70-136">10/11/2019</span></span> |<span data-ttu-id="00f70-137">Guthaben</span><span class="sxs-lookup"><span data-stu-id="00f70-137">Credit</span></span> |<span data-ttu-id="00f70-138">The Phone Company</span><span class="sxs-lookup"><span data-stu-id="00f70-138">The Phone Company</span></span> |<span data-ttu-id="00f70-139">99.95</span><span class="sxs-lookup"><span data-stu-id="00f70-139">99.95</span></span> | |
    |<span data-ttu-id="00f70-140">13.10.2019</span><span class="sxs-lookup"><span data-stu-id="00f70-140">10/13/2019</span></span> |<span data-ttu-id="00f70-141">Guthaben</span><span class="sxs-lookup"><span data-stu-id="00f70-141">Credit</span></span> |<span data-ttu-id="00f70-142">Coho Vineyard</span><span class="sxs-lookup"><span data-stu-id="00f70-142">Coho Vineyard</span></span> |<span data-ttu-id="00f70-143">154.43</span><span class="sxs-lookup"><span data-stu-id="00f70-143">154.43</span></span> | |
    |<span data-ttu-id="00f70-144">15.10.2019</span><span class="sxs-lookup"><span data-stu-id="00f70-144">10/15/2019</span></span> |<span data-ttu-id="00f70-145">Wird geprüft</span><span class="sxs-lookup"><span data-stu-id="00f70-145">Checking</span></span> |<span data-ttu-id="00f70-146">Externe Einzahlung</span><span class="sxs-lookup"><span data-stu-id="00f70-146">External Deposit</span></span> | |<span data-ttu-id="00f70-147">1000</span><span class="sxs-lookup"><span data-stu-id="00f70-147">1000</span></span> |
    |<span data-ttu-id="00f70-148">20.10.2019</span><span class="sxs-lookup"><span data-stu-id="00f70-148">10/20/2019</span></span> |<span data-ttu-id="00f70-149">Guthaben</span><span class="sxs-lookup"><span data-stu-id="00f70-149">Credit</span></span> |<span data-ttu-id="00f70-150">Coho Vineyard – Rückerstattung</span><span class="sxs-lookup"><span data-stu-id="00f70-150">Coho Vineyard - Refund</span></span> | |<span data-ttu-id="00f70-151">-35.45</span><span class="sxs-lookup"><span data-stu-id="00f70-151">-35.45</span></span> |
    |<span data-ttu-id="00f70-152">25.10.2019</span><span class="sxs-lookup"><span data-stu-id="00f70-152">10/25/2019</span></span> |<span data-ttu-id="00f70-153">Wird geprüft</span><span class="sxs-lookup"><span data-stu-id="00f70-153">Checking</span></span> |<span data-ttu-id="00f70-154">Best For You Organics Company</span><span class="sxs-lookup"><span data-stu-id="00f70-154">Best For You Organics Company</span></span> | <span data-ttu-id="00f70-155">-85.64</span><span class="sxs-lookup"><span data-stu-id="00f70-155">-85.64</span></span> | |
    |<span data-ttu-id="00f70-156">01.11.2019</span><span class="sxs-lookup"><span data-stu-id="00f70-156">11/01/2019</span></span> |<span data-ttu-id="00f70-157">Wird geprüft</span><span class="sxs-lookup"><span data-stu-id="00f70-157">Checking</span></span> |<span data-ttu-id="00f70-158">Externe Einzahlung</span><span class="sxs-lookup"><span data-stu-id="00f70-158">External Deposit</span></span> | |<span data-ttu-id="00f70-159">1000</span><span class="sxs-lookup"><span data-stu-id="00f70-159">1000</span></span> |

3. <span data-ttu-id="00f70-160">Öffnen Sie den **Code-Editor**, und wählen Sie **Neuer Skript**aus.</span><span class="sxs-lookup"><span data-stu-id="00f70-160">Open the **Code Editor** and select **New Script**.</span></span>
4. <span data-ttu-id="00f70-161">Lassen Sie uns die Formatierung zurechtmachen.</span><span class="sxs-lookup"><span data-stu-id="00f70-161">Let's clean up the formatting.</span></span> <span data-ttu-id="00f70-162">Dies ist ein Finanzdokument. Ändern Sie daher die Zahlenformatierung in den Spalten **Lastschrift** und **Gutschrift**, um Werte als Dollarbeträge anzuzeigen.</span><span class="sxs-lookup"><span data-stu-id="00f70-162">This is a financial document, so let's change the number formatting in the **Debit** and **Credit** columns to show values as dollar amounts.</span></span> <span data-ttu-id="00f70-163">Passen wir auch die Spaltenbreite an die Daten an.</span><span class="sxs-lookup"><span data-stu-id="00f70-163">Let's also fit the column width to the data.</span></span>

    <span data-ttu-id="00f70-164">Ersetzen Sie den Skriptinhalt durch den folgenden Code:</span><span class="sxs-lookup"><span data-stu-id="00f70-164">Replace the script contents with the following code:</span></span>

    ```TypeScript
    async function main(context: Excel.RequestContext) {
      // Get the current worksheet.
      let workbook = context.workbook;
      let worksheets = workbook.worksheets;
      let selectedSheet = worksheets.getActiveWorksheet();

      // Format the range to display numerical dollar amounts.
      selectedSheet.getRange("D2:E8").numberFormat = [["$#,##0.00"]];

      // Fit the width of all the used columns to the data.
      selectedSheet.getUsedRange().format.autofitColumns();
    }
    ```

5. <span data-ttu-id="00f70-165">Lesen wir nun einen Wert aus einer der Zahlenspalten.</span><span class="sxs-lookup"><span data-stu-id="00f70-165">Now let's read a value from one of the number columns.</span></span> <span data-ttu-id="00f70-166">Fügen Sie am Ende des Skripts den folgenden Code hinzu:</span><span class="sxs-lookup"><span data-stu-id="00f70-166">Add the following code to the end of the script:</span></span>

    ```TypeScript
    // Get the value of cell D2.
    let range = selectedSheet.getRange("D2");
    range.load("values");
    await context.sync();
  
    // Print the value of D2.
    console.log(range.values);
    ```

    <span data-ttu-id="00f70-167">Beachten Sie die Anrufe zu `load` und `sync`.</span><span class="sxs-lookup"><span data-stu-id="00f70-167">Note the calls to `load` and `sync`.</span></span> <span data-ttu-id="00f70-168">Einzelheiten zu diesen Methoden finden Sie unter Grundlagen der [Skripterstellung für Office-Skripts in Excel im Web](../develop/scripting-fundamentals.md#sync-and-load).</span><span class="sxs-lookup"><span data-stu-id="00f70-168">You can learn the details of those methods in [Scripting Fundamentals for Office Scripts in Excel on the web](../develop/scripting-fundamentals.md#sync-and-load).</span></span> <span data-ttu-id="00f70-169">Beachten Sie zunächst, dass Sie das Lesen von Daten anfordern und dann Ihr Skript mit der Arbeitsmappe synchronisieren müssen, um diese Daten zu lesen.</span><span class="sxs-lookup"><span data-stu-id="00f70-169">For now, know that you must request data to be read and then sync your script with the workbook to read that data.</span></span>

6. <span data-ttu-id="00f70-170">Führen Sie das Skript aus.</span><span class="sxs-lookup"><span data-stu-id="00f70-170">Run the script.</span></span>
7. <span data-ttu-id="00f70-171">Öffnen Sie die Konsole.</span><span class="sxs-lookup"><span data-stu-id="00f70-171">Open the console.</span></span> <span data-ttu-id="00f70-172">Wechseln Sie zum Menü **Ellipsen** und drücken Sie **Protokolle...**.</span><span class="sxs-lookup"><span data-stu-id="00f70-172">Go to the **Ellipses** menu and press **Logs...**.</span></span>
8. <span data-ttu-id="00f70-173">Sie sollten `[Array[1]]` in der Konsole sehen.</span><span class="sxs-lookup"><span data-stu-id="00f70-173">You should see `[Array[1]]` in the console.</span></span> <span data-ttu-id="00f70-174">Dies ist keine Zahl, da Bereiche zweidimensionale Datenfelder sind.</span><span class="sxs-lookup"><span data-stu-id="00f70-174">This is not a number because ranges are two-dimensional arrays of data.</span></span> <span data-ttu-id="00f70-175">Dieser zweidimensionale Bereich wird direkt in der Konsole protokolliert.</span><span class="sxs-lookup"><span data-stu-id="00f70-175">That two-dimensional range is being logged to the console directly.</span></span> <span data-ttu-id="00f70-176">Glücklicherweise können Sie mit dem Code-Editor den Inhalt des Arrays anzeigen.</span><span class="sxs-lookup"><span data-stu-id="00f70-176">Luckily, the Code Editor does let you see the contents of the array.</span></span>
9. <span data-ttu-id="00f70-177">Wenn ein zweidimensionales Array in der Konsole protokolliert wird, werden Spaltenwerte unter jeder Zeile gruppiert.</span><span class="sxs-lookup"><span data-stu-id="00f70-177">When a two-dimensional array is logged to the console, it groups column values under each row.</span></span> <span data-ttu-id="00f70-178">Erweitern Sie das Array-Protokoll, indem Sie auf das blaue Dreieck klicken.</span><span class="sxs-lookup"><span data-stu-id="00f70-178">Expand the array log by pressing the blue triangle.</span></span>
10. <span data-ttu-id="00f70-179">Erweitern Sie die zweite Ebene des Arrays, indem Sie auf das neu aufgedeckte blaue Dreieck klicken.</span><span class="sxs-lookup"><span data-stu-id="00f70-179">Expand the second level of the array by pressing the newly revealed blue triangle.</span></span> <span data-ttu-id="00f70-180">Sie sollten jetzt Folgendes sehen:</span><span class="sxs-lookup"><span data-stu-id="00f70-180">You should now see this:</span></span>

    ![Das Konsolenprotokoll mit der Ausgabe "-20.05", verschachtelt unter zwei Arrays.](../images/tutorial-4.png)

## <a name="modify-the-value-of-a-cell"></a><span data-ttu-id="00f70-182">Ändern Sie den Wert einer Zelle</span><span class="sxs-lookup"><span data-stu-id="00f70-182">Modify the value of a cell</span></span>

<span data-ttu-id="00f70-183">Nachdem wir nun Daten lesen können, verwenden wir diese Daten, um die Arbeitsmappe zu ändern.</span><span class="sxs-lookup"><span data-stu-id="00f70-183">Now that we can read data, let's use that data to modify the workbook.</span></span> <span data-ttu-id="00f70-184">Wir werden den Wert der Zelle **D2** mit der `Math.abs` Funktion positiv machen.</span><span class="sxs-lookup"><span data-stu-id="00f70-184">We'll make the value of the cell **D2** positive with the `Math.abs` function.</span></span> <span data-ttu-id="00f70-185">Das [Math](https://developer.mozilla.org/docs/web/javascript/reference/global_objects/math)-Objekt enthält viele Funktionen, auf die Ihre Skripte Zugriff haben.</span><span class="sxs-lookup"><span data-stu-id="00f70-185">The [Math](https://developer.mozilla.org/docs/web/javascript/reference/global_objects/math) object contains many functions to which your scripts have access.</span></span> <span data-ttu-id="00f70-186">Weitere Informationen zu `Math` und andere integrierte Objekte finden Sie unter [Verwenden von integrierten JavaScript-Objekten in Office-Skripts](../develop/javascript-objects.md).</span><span class="sxs-lookup"><span data-stu-id="00f70-186">More information about `Math` and other built-in objects can be found at [Using built-in JavaScript objects in Office Scripts](../develop/javascript-objects.md).</span></span>

1. <span data-ttu-id="00f70-187">Fügen Sie am Ende des Skripts den folgenden Code hinzu:</span><span class="sxs-lookup"><span data-stu-id="00f70-187">Add the following code to the end of the script:</span></span>

    ```TypeScript
    // Run the `Math.abs` function with the value at D2 and apply that value back to D2.
    let positiveValue = Math.abs(range.values[0][0]);
    range.values = [[positiveValue]];
    ```

2. <span data-ttu-id="00f70-188">Der Wert von Zelle **D2** sollte jetzt positiv sein.</span><span class="sxs-lookup"><span data-stu-id="00f70-188">The value of cell **D2** should now be positive.</span></span>

## <a name="modify-the-values-of-a-column"></a><span data-ttu-id="00f70-189">Ändern Sie die Werte einer Spalte</span><span class="sxs-lookup"><span data-stu-id="00f70-189">Modify the values of a column</span></span>

<span data-ttu-id="00f70-190">Nachdem wir nun wissen, wie man in eine einzelne Zelle liest und schreibt, verallgemeinern wir das Skript so, dass es für die gesamten Spalten **Lastschrift** und **Gutschrift** funktioniert.</span><span class="sxs-lookup"><span data-stu-id="00f70-190">Now that we know how to read and write to a single cell, let's generalize the script to work on the entire **Debit** and **Credit** columns.</span></span>

1. <span data-ttu-id="00f70-191">Entfernen Sie den Code, der nur eine einzelne Zelle betrifft (den vorherigen Absolutwertcode), sodass Ihr Skript jetzt folgendermaßen aussieht:</span><span class="sxs-lookup"><span data-stu-id="00f70-191">Remove the code that affects only a single cell (the previous absolute value code), such that your script now looks like this:</span></span>

    ```TypeScript
    async function main(context: Excel.RequestContext) {
      // Get the current worksheet.
      let workbook = context.workbook;
      let worksheets = workbook.worksheets;
      let selectedSheet = worksheets.getActiveWorksheet();

      // Format the range to display numerical dollar amounts.
      selectedSheet.getRange("D2:E8").numberFormat = [["$#,##0.00"]];

      // Fit the width of all the used columns to the data.
      selectedSheet.getUsedRange().format.autofitColumns();
    }
    ```

2. <span data-ttu-id="00f70-192">Fügen Sie eine Schleife hinzu, die die Zeilen in den letzten beiden Spalten durchläuft.</span><span class="sxs-lookup"><span data-stu-id="00f70-192">Add a loop that iterates through the rows in the last two columns.</span></span> <span data-ttu-id="00f70-193">Für jede Zelle setzt das Skript den Wert auf den absoluten Wert des aktuellen Werts.</span><span class="sxs-lookup"><span data-stu-id="00f70-193">For each cell, the script sets the value to the current value's absolute value.</span></span>

    <span data-ttu-id="00f70-194">Beachten Sie, dass das Array, das die Zellenpositionen definiert, auf Null basiert.</span><span class="sxs-lookup"><span data-stu-id="00f70-194">Note that the array defining cell locations is zero-based.</span></span> <span data-ttu-id="00f70-195">Das heißt, Zelle **A1** ist `range[0][0]`.</span><span class="sxs-lookup"><span data-stu-id="00f70-195">That means cell **A1** is `range[0][0]`.</span></span>

    ```TypeScript
    // Get the values of the used range.
    let range = selectedSheet.getUsedRange();
    range.load("rowCount,values");
    await context.sync();

    // Iterate over the fourth and fifth columns and set their values to their absolute value.
    for (let i = 1; i < range.rowCount; i++) {
      // The column at index 3 is column "4" in the worksheet.
      if (range.values[i][3] != 0) {
        let positiveValue = Math.abs(range.values[i][3]);
        selectedSheet.getCell(i, 3).values = [[positiveValue]];
      }

      // The column at index 4 is column "5" in the worksheet.
      if (range.values[i][4] != 0) {
        let positiveValue = Math.abs(range.values[i][4]);
        selectedSheet.getCell(i, 4).values = [[positiveValue]];
      }
    }
    ```

    <span data-ttu-id="00f70-196">Dieser Teil des Skripts führt mehrere wichtige Aufgaben aus.</span><span class="sxs-lookup"><span data-stu-id="00f70-196">This portion of the script does several important tasks.</span></span> <span data-ttu-id="00f70-197">Zunächst werden die Werte und die Zeilenanzahl des verwendeten Bereichs geladen.</span><span class="sxs-lookup"><span data-stu-id="00f70-197">First, it loads the values and row count of the used range.</span></span> <span data-ttu-id="00f70-198">Auf diese Weise können wir uns die Werte ansehen und wissen, wann wir aufhören müssen.</span><span class="sxs-lookup"><span data-stu-id="00f70-198">This lets us look at values and know when to stop.</span></span> <span data-ttu-id="00f70-199">Zweitens durchläuft es den verwendeten Bereich und überprüft jede Zelle in den Spalten **Lastschrift** und **Gutschrift**.</span><span class="sxs-lookup"><span data-stu-id="00f70-199">Second, it iterates through the used range, checking each cell in the **Debit** or **Credit** columns.</span></span> <span data-ttu-id="00f70-200">Wenn der Wert in der Zelle nicht 0 ist, wird er durch seinen absoluten Wert ersetzt.</span><span class="sxs-lookup"><span data-stu-id="00f70-200">Finally, if the value in the cell is not 0, it is replaced by its absolute value.</span></span> <span data-ttu-id="00f70-201">Wir vermeiden Nullen, damit wir die leeren Zellen so lassen können, wie sie waren.</span><span class="sxs-lookup"><span data-stu-id="00f70-201">We're avoiding zeroes so we can leave the blank cells as they were.</span></span>

3. <span data-ttu-id="00f70-202">Führen Sie das Skript aus.</span><span class="sxs-lookup"><span data-stu-id="00f70-202">Run the script.</span></span>

    <span data-ttu-id="00f70-203">Ihr Kontoauszug sollte nun folgendermaßen aussehen:</span><span class="sxs-lookup"><span data-stu-id="00f70-203">Your banking statement should now look like this:</span></span>

    ![Der Kontoauszug als formatierte Tabelle mit nur positiven Werten.](../images/tutorial-5.png)

## <a name="next-steps"></a><span data-ttu-id="00f70-205">Nächste Schritte</span><span class="sxs-lookup"><span data-stu-id="00f70-205">Next steps</span></span>

<span data-ttu-id="00f70-206">Öffnen Sie den Code-Editor, und probieren Sie einige unserer [Beispielskripts für Office-Skripts in Excel im Web](../resources/excel-samples.md)aus.</span><span class="sxs-lookup"><span data-stu-id="00f70-206">Open the Code Editor and try out some of our [Sample scripts for Office Scripts in Excel on the web](../resources/excel-samples.md).</span></span> <span data-ttu-id="00f70-207">Sie können auch [Skripting-Grundlagen für Office-Skripts in Excel im Web](../develop/scripting-fundamentals.md) aufrufen, um weitere Informationen zum Erstellen von Office-Skripts zu erhalten.</span><span class="sxs-lookup"><span data-stu-id="00f70-207">You can also visit [Scripting Fundamentals for Office Scripts in Excel on the web](../develop/scripting-fundamentals.md) to learn more about creating Office Scripts.</span></span>
