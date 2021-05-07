---
title: Lesen Sie Arbeitsmappendaten mit Office-Skripts in Excel im Web
description: Ein Office Skripts-Lernprogramm zum Lesen von Daten aus Arbeitsmappen und zum Auswerten dieser Daten im Skript.
ms.date: 01/06/2021
localization_priority: Priority
ms.openlocfilehash: 979ad949f999f67ee9de750c53ba4690ce6294e6
ms.sourcegitcommit: f7a7aebfb687f2a35dbed07ed62ff352a114525a
ms.translationtype: HT
ms.contentlocale: de-DE
ms.lasthandoff: 05/06/2021
ms.locfileid: "52232851"
---
# <a name="read-workbook-data-with-office-scripts-in-excel-on-the-web"></a><span data-ttu-id="32f04-103">Lesen Sie Arbeitsmappendaten mit Office-Skripts in Excel im Web</span><span class="sxs-lookup"><span data-stu-id="32f04-103">Read workbook data with Office Scripts in Excel on the web</span></span>

<span data-ttu-id="32f04-104">In diesem Lernprogramm erfahren Sie, wie Sie Daten aus einer Arbeitsmappe mit einem Office-Skript für Excel im Web lesen.</span><span class="sxs-lookup"><span data-stu-id="32f04-104">This tutorial teaches you how to read data from a workbook with an Office Script for Excel on the web.</span></span> <span data-ttu-id="32f04-105">Sie werden ein neues Skript schreiben, das einen Kontoauszug formatiert und die Daten in diesem Auszug normalisiert.</span><span class="sxs-lookup"><span data-stu-id="32f04-105">You'll be writing a new script that formats a bank statement and normalizes the data in that statement.</span></span> <span data-ttu-id="32f04-106">Im Rahmen der Datenbereinigung liest Ihr Skript Werte aus den Buchungszellen, wendet auf jeden Wert eine einfache Formel an und schreibt die resultierende Antwort in die Arbeitsmappe.</span><span class="sxs-lookup"><span data-stu-id="32f04-106">As part of that data clean-up, your script will read values from the transaction cells, apply a simple formula to each value, and write the resulting answer to the workbook.</span></span> <span data-ttu-id="32f04-107">Durch das Lesen von Daten aus der Arbeitsmappe können Sie einige Ihrer Entscheidungsfindungsprozesse im Skript automatisieren.</span><span class="sxs-lookup"><span data-stu-id="32f04-107">Reading data from the workbook lets you automate some of your decision making processes in the script.</span></span>

> [!TIP]
> <span data-ttu-id="32f04-108">Wenn Sie mit Office-Skripten noch nicht vertraut sind, empfehlen wir, mit dem [Aufzeichnen, Bearbeiten und Erstellen von Office-Skripten in Excel im Web](excel-tutorial.md)-Lernprogramm zu beginnen.</span><span class="sxs-lookup"><span data-stu-id="32f04-108">If you are new to Office Scripts, we recommend starting with the [Record, edit, and create Office Scripts in Excel on the web](excel-tutorial.md) tutorial.</span></span> <span data-ttu-id="32f04-109">[Office-Skripts verwenden TypeScript](../overview/code-editor-environment.md), und dieses Lernprogramm richten sich an Anfänger bis Fortgeschrittene mit JavaScript oder TypeScript.</span><span class="sxs-lookup"><span data-stu-id="32f04-109">[Office Scripts use TypeScript](../overview/code-editor-environment.md) and this tutorial is intended for people with beginner to intermediate-level knowledge of JavaScript or TypeScript.</span></span> <span data-ttu-id="32f04-110">Wenn Sie noch nicht mit JavaScript vertraut sind, empfehlen wir Ihnen, mit dem [Mozilla-JavaScript-Lernprogramm](https://developer.mozilla.org/docs/Web/JavaScript/Guide/Introduction) zu beginnen.</span><span class="sxs-lookup"><span data-stu-id="32f04-110">If you're new to JavaScript, we recommend starting with the [Mozilla JavaScript tutorial](https://developer.mozilla.org/docs/Web/JavaScript/Guide/Introduction).</span></span>

## <a name="prerequisites"></a><span data-ttu-id="32f04-111">Voraussetzungen</span><span class="sxs-lookup"><span data-stu-id="32f04-111">Prerequisites</span></span>

[!INCLUDE [Tutorial prerequisites](../includes/tutorial-prerequisites.md)]

## <a name="read-a-cell"></a><span data-ttu-id="32f04-112">Lesen einer Zelle</span><span class="sxs-lookup"><span data-stu-id="32f04-112">Read a cell</span></span>

<span data-ttu-id="32f04-113">Mit dem Action Recorder erstellte Skripte können nur Informationen in die Arbeitsmappe schreiben.</span><span class="sxs-lookup"><span data-stu-id="32f04-113">Scripts made with the Action Recorder can only write information to the workbook.</span></span> <span data-ttu-id="32f04-114">Mit dem Code-Editor können Sie Skripte bearbeiten und erstellen, die auch Daten aus einer Arbeitsmappe lesen.</span><span class="sxs-lookup"><span data-stu-id="32f04-114">With the Code Editor, you can edit and make scripts that also read data from a workbook.</span></span>

<span data-ttu-id="32f04-115">Lassen Sie uns ein Skript erstellen, das Daten liest und basierend auf dem Gelesenen agiert.</span><span class="sxs-lookup"><span data-stu-id="32f04-115">Let's make a script that reads data and acts based on what was read.</span></span> <span data-ttu-id="32f04-116">Wir werden mit einem Muster-Kontoauszug arbeiten.</span><span class="sxs-lookup"><span data-stu-id="32f04-116">We're going to work with a sample banking statement.</span></span> <span data-ttu-id="32f04-117">Diese Erklärung ist ein kombinierter Prüfungs- und Gutschrift-Kontoauszug.</span><span class="sxs-lookup"><span data-stu-id="32f04-117">This statement is a combined checking and credit statement.</span></span> <span data-ttu-id="32f04-118">Leider melden sie Bilanzänderungen unterschiedlich.</span><span class="sxs-lookup"><span data-stu-id="32f04-118">Unfortunately, they report balance changes differently.</span></span> <span data-ttu-id="32f04-119">Der Prüfungskontoauszug gibt die Einnahmen als positive Gutschrift und die Kosten als negative Belastung an.</span><span class="sxs-lookup"><span data-stu-id="32f04-119">The checking statement gives income as positive credit and costs as negative debit.</span></span> <span data-ttu-id="32f04-120">Der Gutschrift-Kontoauszug macht das Gegenteil.</span><span class="sxs-lookup"><span data-stu-id="32f04-120">The credit statement does the opposite.</span></span>

<span data-ttu-id="32f04-121">Im weiteren Verlauf des Lernprogramms werden diese Daten mithilfe eines Skripts normalisiert.</span><span class="sxs-lookup"><span data-stu-id="32f04-121">Over the rest of the tutorial, we will normalize this data using a script.</span></span> <span data-ttu-id="32f04-122">Lassen Sie uns zunächst lernen, wie Sie Daten aus der Arbeitsmappe lesen.</span><span class="sxs-lookup"><span data-stu-id="32f04-122">First, let's learn how to read data from the workbook.</span></span>

1. <span data-ttu-id="32f04-123">Erstellen Sie ein neues Arbeitsblatt in der Arbeitsmappe, die Sie für den Rest des Lernprogramms verwendet haben.</span><span class="sxs-lookup"><span data-stu-id="32f04-123">Create a new worksheet in the workbook you've used for the rest of the tutorial.</span></span>
2. <span data-ttu-id="32f04-124">Kopieren Sie die folgenden Daten und fügen Sie sie ab Zelle **A1** in das neue Arbeitsblatt ein.</span><span class="sxs-lookup"><span data-stu-id="32f04-124">Copy the following data and paste it into the new worksheet, starting at cell **A1**.</span></span>

    |<span data-ttu-id="32f04-125">Datum</span><span class="sxs-lookup"><span data-stu-id="32f04-125">Date</span></span> |<span data-ttu-id="32f04-126">Konto</span><span class="sxs-lookup"><span data-stu-id="32f04-126">Account</span></span> |<span data-ttu-id="32f04-127">Beschreibung</span><span class="sxs-lookup"><span data-stu-id="32f04-127">Description</span></span> |<span data-ttu-id="32f04-128">Lastschrift</span><span class="sxs-lookup"><span data-stu-id="32f04-128">Debit</span></span> |<span data-ttu-id="32f04-129">Guthaben</span><span class="sxs-lookup"><span data-stu-id="32f04-129">Credit</span></span> |
    |:--|:--|:--|:--|:--|
    |<span data-ttu-id="32f04-130">10.10.2019</span><span class="sxs-lookup"><span data-stu-id="32f04-130">10/10/2019</span></span> |<span data-ttu-id="32f04-131">Wird geprüft</span><span class="sxs-lookup"><span data-stu-id="32f04-131">Checking</span></span> |<span data-ttu-id="32f04-132">Coho Vineyard</span><span class="sxs-lookup"><span data-stu-id="32f04-132">Coho Vineyard</span></span> |<span data-ttu-id="32f04-133">-20.05</span><span class="sxs-lookup"><span data-stu-id="32f04-133">-20.05</span></span> | |
    |<span data-ttu-id="32f04-134">11.10.2019</span><span class="sxs-lookup"><span data-stu-id="32f04-134">10/11/2019</span></span> |<span data-ttu-id="32f04-135">Guthaben</span><span class="sxs-lookup"><span data-stu-id="32f04-135">Credit</span></span> |<span data-ttu-id="32f04-136">The Phone Company</span><span class="sxs-lookup"><span data-stu-id="32f04-136">The Phone Company</span></span> |<span data-ttu-id="32f04-137">99.95</span><span class="sxs-lookup"><span data-stu-id="32f04-137">99.95</span></span> | |
    |<span data-ttu-id="32f04-138">13.10.2019</span><span class="sxs-lookup"><span data-stu-id="32f04-138">10/13/2019</span></span> |<span data-ttu-id="32f04-139">Guthaben</span><span class="sxs-lookup"><span data-stu-id="32f04-139">Credit</span></span> |<span data-ttu-id="32f04-140">Coho Vineyard</span><span class="sxs-lookup"><span data-stu-id="32f04-140">Coho Vineyard</span></span> |<span data-ttu-id="32f04-141">154.43</span><span class="sxs-lookup"><span data-stu-id="32f04-141">154.43</span></span> | |
    |<span data-ttu-id="32f04-142">15.10.2019</span><span class="sxs-lookup"><span data-stu-id="32f04-142">10/15/2019</span></span> |<span data-ttu-id="32f04-143">Wird geprüft</span><span class="sxs-lookup"><span data-stu-id="32f04-143">Checking</span></span> |<span data-ttu-id="32f04-144">Externe Einzahlung</span><span class="sxs-lookup"><span data-stu-id="32f04-144">External Deposit</span></span> | |<span data-ttu-id="32f04-145">1000</span><span class="sxs-lookup"><span data-stu-id="32f04-145">1000</span></span> |
    |<span data-ttu-id="32f04-146">20.10.2019</span><span class="sxs-lookup"><span data-stu-id="32f04-146">10/20/2019</span></span> |<span data-ttu-id="32f04-147">Guthaben</span><span class="sxs-lookup"><span data-stu-id="32f04-147">Credit</span></span> |<span data-ttu-id="32f04-148">Coho Vineyard – Rückerstattung</span><span class="sxs-lookup"><span data-stu-id="32f04-148">Coho Vineyard - Refund</span></span> | |<span data-ttu-id="32f04-149">-35.45</span><span class="sxs-lookup"><span data-stu-id="32f04-149">-35.45</span></span> |
    |<span data-ttu-id="32f04-150">25.10.2019</span><span class="sxs-lookup"><span data-stu-id="32f04-150">10/25/2019</span></span> |<span data-ttu-id="32f04-151">Wird geprüft</span><span class="sxs-lookup"><span data-stu-id="32f04-151">Checking</span></span> |<span data-ttu-id="32f04-152">Best For You Organics Company</span><span class="sxs-lookup"><span data-stu-id="32f04-152">Best For You Organics Company</span></span> | <span data-ttu-id="32f04-153">-85.64</span><span class="sxs-lookup"><span data-stu-id="32f04-153">-85.64</span></span> | |
    |<span data-ttu-id="32f04-154">01.11.2019</span><span class="sxs-lookup"><span data-stu-id="32f04-154">11/01/2019</span></span> |<span data-ttu-id="32f04-155">Wird geprüft</span><span class="sxs-lookup"><span data-stu-id="32f04-155">Checking</span></span> |<span data-ttu-id="32f04-156">Externe Einzahlung</span><span class="sxs-lookup"><span data-stu-id="32f04-156">External Deposit</span></span> | |<span data-ttu-id="32f04-157">1000</span><span class="sxs-lookup"><span data-stu-id="32f04-157">1000</span></span> |

3. <span data-ttu-id="32f04-158">Öffnen Sie **Alle Skripts** und wählen Sie **Neues Skript** aus.</span><span class="sxs-lookup"><span data-stu-id="32f04-158">Open **All Scripts** and select **New Script**.</span></span>
4. <span data-ttu-id="32f04-159">Lassen Sie uns die Formatierung zurechtmachen.</span><span class="sxs-lookup"><span data-stu-id="32f04-159">Let's clean up the formatting.</span></span> <span data-ttu-id="32f04-160">Dies ist ein Finanzdokument. Ändern Sie daher die Zahlenformatierung in den Spalten **Lastschrift** und **Gutschrift**, um Werte als Dollarbeträge anzuzeigen.</span><span class="sxs-lookup"><span data-stu-id="32f04-160">This is a financial document, so let's change the number formatting in the **Debit** and **Credit** columns to show values as dollar amounts.</span></span> <span data-ttu-id="32f04-161">Passen wir auch die Spaltenbreite an die Daten an.</span><span class="sxs-lookup"><span data-stu-id="32f04-161">Let's also fit the column width to the data.</span></span>

    <span data-ttu-id="32f04-162">Ersetzen Sie den Skriptinhalt durch den folgenden Code:</span><span class="sxs-lookup"><span data-stu-id="32f04-162">Replace the script contents with the following code:</span></span>

    ```TypeScript
    function main(workbook: ExcelScript.Workbook) {
        // Get the current worksheet.
        let selectedSheet = workbook.getActiveWorksheet();

        // Format the range to display numerical dollar amounts.
        selectedSheet.getRange("D2:E8").setNumberFormat("$#,##0.00");

        // Fit the width of all the used columns to the data.
        selectedSheet.getUsedRange().getFormat().autofitColumns();
    }
    ```

5. <span data-ttu-id="32f04-163">Lesen wir nun einen Wert aus einer der Zahlenspalten.</span><span class="sxs-lookup"><span data-stu-id="32f04-163">Now let's read a value from one of the number columns.</span></span> <span data-ttu-id="32f04-164">Fügen Sie den folgenden Code am Ende des Skripts hinzu (vor der schließenden `}`):</span><span class="sxs-lookup"><span data-stu-id="32f04-164">Add the following code to the end of the script (before the closing `}`):</span></span>

    ```TypeScript
    // Get the value of cell D2.
    let range = selectedSheet.getRange("D2");
    console.log(range.getValues());
    ```

6. <span data-ttu-id="32f04-165">Führen Sie das Skript aus.</span><span class="sxs-lookup"><span data-stu-id="32f04-165">Run the script.</span></span>
7. <span data-ttu-id="32f04-166">Sie sollten `[Array[1]]` in der Konsole sehen.</span><span class="sxs-lookup"><span data-stu-id="32f04-166">You should see `[Array[1]]` in the console.</span></span> <span data-ttu-id="32f04-167">Dies ist keine Zahl, da Bereiche zweidimensionale Datenfelder sind.</span><span class="sxs-lookup"><span data-stu-id="32f04-167">This is not a number because ranges are two-dimensional arrays of data.</span></span> <span data-ttu-id="32f04-168">Dieser zweidimensionale Bereich wird direkt in der Konsole protokolliert.</span><span class="sxs-lookup"><span data-stu-id="32f04-168">That two-dimensional range is being logged to the console directly.</span></span> <span data-ttu-id="32f04-169">Glücklicherweise können Sie mit dem Code-Editor den Inhalt des Arrays anzeigen.</span><span class="sxs-lookup"><span data-stu-id="32f04-169">Luckily, the Code Editor lets you see the contents of the array.</span></span>
8. <span data-ttu-id="32f04-170">Wenn ein zweidimensionales Array in der Konsole protokolliert wird, werden Spaltenwerte unter jeder Zeile gruppiert.</span><span class="sxs-lookup"><span data-stu-id="32f04-170">When a two-dimensional array is logged to the console, it groups column values under each row.</span></span> <span data-ttu-id="32f04-171">Erweitern Sie das Array-Protokoll, indem Sie auf das blaue Dreieck klicken.</span><span class="sxs-lookup"><span data-stu-id="32f04-171">Expand the array log by pressing the blue triangle.</span></span>
9. <span data-ttu-id="32f04-p110">Erweitern Sie die zweite Ebene des Arrays, indem Sie auf das neu aufgedeckte blaue Dreieck klicken. Jetzt sollten Sie folgendes sehen:</span><span class="sxs-lookup"><span data-stu-id="32f04-p110">Expand the second level of the array by pressing the newly revealed blue triangle. You should now see this:</span></span>

    :::image type="content" source="../images/tutorial-4.png" alt-text="Das Konsolenprotokoll mit der Ausgabe „-20,05“, verschachtelt unter zwei Arrays":::

## <a name="modify-the-value-of-a-cell"></a><span data-ttu-id="32f04-175">Ändern Sie den Wert einer Zelle</span><span class="sxs-lookup"><span data-stu-id="32f04-175">Modify the value of a cell</span></span>

<span data-ttu-id="32f04-176">Nachdem wir nun Daten lesen können, verwenden wir diese Daten, um die Arbeitsmappe zu ändern.</span><span class="sxs-lookup"><span data-stu-id="32f04-176">Now that we can read data, let's use that data to modify the workbook.</span></span> <span data-ttu-id="32f04-177">Wir werden den Wert der Zelle **D2** mit der `Math.abs` Funktion positiv machen.</span><span class="sxs-lookup"><span data-stu-id="32f04-177">We'll make the value of the cell **D2** positive with the `Math.abs` function.</span></span> <span data-ttu-id="32f04-178">Das [Math](https://developer.mozilla.org/docs/web/javascript/reference/global_objects/math)-Objekt enthält viele Funktionen, auf die Ihre Skripte Zugriff haben.</span><span class="sxs-lookup"><span data-stu-id="32f04-178">The [Math](https://developer.mozilla.org/docs/web/javascript/reference/global_objects/math) object contains many functions to which your scripts have access.</span></span> <span data-ttu-id="32f04-179">Weitere Informationen zu `Math` und andere integrierte Objekte finden Sie unter [Verwenden von integrierten JavaScript-Objekten in Office-Skripts](../develop/javascript-objects.md).</span><span class="sxs-lookup"><span data-stu-id="32f04-179">More information about `Math` and other built-in objects can be found at [Using built-in JavaScript objects in Office Scripts](../develop/javascript-objects.md).</span></span>

1. <span data-ttu-id="32f04-180">Sie ändern den Wert der Zelle mithilfe der Methoden `getValue` und `setValue`.</span><span class="sxs-lookup"><span data-stu-id="32f04-180">We'll use `getValue` and `setValue` methods to change the value of the cell.</span></span> <span data-ttu-id="32f04-181">Diese Methoden funktionieren bei einer einzelnen Zelle.</span><span class="sxs-lookup"><span data-stu-id="32f04-181">These methods work on a single cell.</span></span> <span data-ttu-id="32f04-182">Wenn Sie mehrere Zellbereiche bearbeiten möchten, verwenden Sie `getValues` und `setValues`.</span><span class="sxs-lookup"><span data-stu-id="32f04-182">When handling multi-cell ranges, you'll want to use `getValues` and `setValues`.</span></span> <span data-ttu-id="32f04-183">Fügen Sie am Ende des Skripts den folgenden Code hinzu:</span><span class="sxs-lookup"><span data-stu-id="32f04-183">Add the following code to the end of the script:</span></span>

    ```TypeScript
    // Run the `Math.abs` function with the value at D2 and apply that value back to D2.
    let positiveValue = Math.abs(range.getValue() as number);
    range.setValue(positiveValue);
    ```

    > [!NOTE]
    > <span data-ttu-id="32f04-184">Der zurückgegebene Wert wird unter Verwendung des Schlüsselworts `as` von `range.getValue()`in `number` [umgewandelt](https://www.typescripttutorial.net/typescript-tutorial/type-casting/).</span><span class="sxs-lookup"><span data-stu-id="32f04-184">We are [casting](https://www.typescripttutorial.net/typescript-tutorial/type-casting/) the returned value of `range.getValue()` to a `number` by using the `as` keyword.</span></span> <span data-ttu-id="32f04-185">Dies ist notwendig, da ein Bereich aus Zeichenfolgen, Zahlen oder booleschen Werten bestehen kann.</span><span class="sxs-lookup"><span data-stu-id="32f04-185">This is necessary because a range could be strings, numbers, or booleans.</span></span> <span data-ttu-id="32f04-186">In diesem Fall wird explizit eine Zahl benötigt.</span><span class="sxs-lookup"><span data-stu-id="32f04-186">In this instance, we explicitly need a number.</span></span>

2. <span data-ttu-id="32f04-187">Der Wert von Zelle **D2** sollte jetzt positiv sein.</span><span class="sxs-lookup"><span data-stu-id="32f04-187">The value of cell **D2** should now be positive.</span></span>

## <a name="modify-the-values-of-a-column"></a><span data-ttu-id="32f04-188">Ändern Sie die Werte einer Spalte</span><span class="sxs-lookup"><span data-stu-id="32f04-188">Modify the values of a column</span></span>

<span data-ttu-id="32f04-189">Nachdem wir nun wissen, wie man in eine einzelne Zelle liest und schreibt, verallgemeinern wir das Skript so, dass es für die gesamten Spalten **Lastschrift** und **Gutschrift** funktioniert.</span><span class="sxs-lookup"><span data-stu-id="32f04-189">Now that we know how to read and write to a single cell, let's generalize the script to work on the entire **Debit** and **Credit** columns.</span></span>

1. <span data-ttu-id="32f04-190">Entfernen Sie den Code, der nur eine einzelne Zelle betrifft (den vorherigen Absolutwertcode), sodass Ihr Skript jetzt folgendermaßen aussieht:</span><span class="sxs-lookup"><span data-stu-id="32f04-190">Remove the code that affects only a single cell (the previous absolute value code), such that your script now looks like this:</span></span>

    ```TypeScript
    function main(workbook: ExcelScript.Workbook) {
        // Get the current worksheet.
        let selectedSheet = workbook.getActiveWorksheet();

        // Format the range to display numerical dollar amounts.
        selectedSheet.getRange("D2:E8").setNumberFormat("$#,##0.00");

        // Fit the width of all the used columns to the data.
        selectedSheet.getUsedRange().getFormat().autofitColumns();
    }
    ```

2. <span data-ttu-id="32f04-191">Fügen Sie am Ende des Skripts eine Schleife hinzu, die die Zeilen in den letzten beiden Spalten durchläuft.</span><span class="sxs-lookup"><span data-stu-id="32f04-191">Add a loop to the end of the script that iterates through the rows in the last two columns.</span></span> <span data-ttu-id="32f04-192">Für jede Zelle setzt das Skript den Wert auf den absoluten Wert des aktuellen Werts.</span><span class="sxs-lookup"><span data-stu-id="32f04-192">For each cell, the script sets the value to the current value's absolute value.</span></span>

    <span data-ttu-id="32f04-193">Beachten Sie, dass das Array, das die Zellenpositionen definiert, auf Null basiert.</span><span class="sxs-lookup"><span data-stu-id="32f04-193">Note that the array defining cell locations is zero-based.</span></span> <span data-ttu-id="32f04-194">Das heißt, Zelle **A1** ist `range[0][0]`.</span><span class="sxs-lookup"><span data-stu-id="32f04-194">That means cell **A1** is `range[0][0]`.</span></span>

    ```TypeScript
    // Get the values of the used range.
    let range = selectedSheet.getUsedRange();
    let rangeValues = range.getValues();

    // Iterate over the fourth and fifth columns and set their values to their absolute value.
    let rowCount = range.getRowCount();
    for (let i = 1; i < rowCount; i++) {
        // The column at index 3 is column "4" in the worksheet.
        if (rangeValues[i][3] != 0) {
            let positiveValue = Math.abs(rangeValues[i][3] as number);
            selectedSheet.getCell(i, 3).setValue(positiveValue);
        }

        // The column at index 4 is column "5" in the worksheet.
        if (rangeValues[i][4] != 0) {
            let positiveValue = Math.abs(rangeValues[i][4] as number);
            selectedSheet.getCell(i, 4).setValue(positiveValue);
        }
    }
    ```

    <span data-ttu-id="32f04-195">Dieser Teil des Skripts führt mehrere wichtige Aufgaben aus.</span><span class="sxs-lookup"><span data-stu-id="32f04-195">This portion of the script does several important tasks.</span></span> <span data-ttu-id="32f04-196">Zunächst werden die Werte und die Zeilenanzahl des verwendeten Bereichs abgerufen.</span><span class="sxs-lookup"><span data-stu-id="32f04-196">First, it gets the values and row count of the used range.</span></span> <span data-ttu-id="32f04-197">Auf diese Weise können wir uns die Werte ansehen und wissen, wann wir aufhören müssen.</span><span class="sxs-lookup"><span data-stu-id="32f04-197">This lets us look at values and know when to stop.</span></span> <span data-ttu-id="32f04-198">Zweitens durchläuft es den verwendeten Bereich und überprüft jede Zelle in den Spalten **Lastschrift** und **Gutschrift**.</span><span class="sxs-lookup"><span data-stu-id="32f04-198">Second, it iterates through the used range, checking each cell in the **Debit** or **Credit** columns.</span></span> <span data-ttu-id="32f04-199">Wenn der Wert in der Zelle nicht 0 ist, wird er durch seinen absoluten Wert ersetzt.</span><span class="sxs-lookup"><span data-stu-id="32f04-199">Finally, if the value in the cell is not 0, it is replaced by its absolute value.</span></span> <span data-ttu-id="32f04-200">Wir vermeiden Nullen, damit wir die leeren Zellen so lassen können, wie sie waren.</span><span class="sxs-lookup"><span data-stu-id="32f04-200">We're avoiding zeroes so we can leave the blank cells as they were.</span></span>

3. <span data-ttu-id="32f04-201">Führen Sie das Skript aus.</span><span class="sxs-lookup"><span data-stu-id="32f04-201">Run the script.</span></span>

    <span data-ttu-id="32f04-202">Ihr Kontoauszug sollte nun folgendermaßen aussehen:</span><span class="sxs-lookup"><span data-stu-id="32f04-202">Your banking statement should now look like this:</span></span>

    :::image type="content" source="../images/tutorial-5.png" alt-text="Ein Arbeitsblatt, auf dem der Kontoauszug als formatierte Tabelle mit nur positiven Werten angezeigt wird":::

## <a name="next-steps"></a><span data-ttu-id="32f04-204">Nächste Schritte</span><span class="sxs-lookup"><span data-stu-id="32f04-204">Next steps</span></span>

<span data-ttu-id="32f04-205">Öffnen Sie den Code-Editor, und probieren Sie einige unserer [Beispielskripts für Office-Skripts in Excel im Web](../resources/samples/excel-samples.md)aus.</span><span class="sxs-lookup"><span data-stu-id="32f04-205">Open the Code Editor and try out some of our [Sample scripts for Office Scripts in Excel on the web](../resources/samples/excel-samples.md).</span></span> <span data-ttu-id="32f04-206">Sie können auch [Skripting-Grundlagen für Office-Skripts in Excel im Web](../develop/scripting-fundamentals.md) aufrufen, um weitere Informationen zum Erstellen von Office-Skripts zu erhalten.</span><span class="sxs-lookup"><span data-stu-id="32f04-206">You can also visit [Scripting Fundamentals for Office Scripts in Excel on the web](../develop/scripting-fundamentals.md) to learn more about creating Office Scripts.</span></span>

<span data-ttu-id="32f04-207">Die nächste Reihe von Office-Skripts-Lernprogrammen konzentriert sich auf die Verwendung von Office-Skripts mit Power Automate.</span><span class="sxs-lookup"><span data-stu-id="32f04-207">The next series of Office Scripts tutorials focus on using Office Scripts with Power Automate.</span></span> <span data-ttu-id="32f04-208">Weitere Informationen über die Vorteile, diese beiden Plattformen miteinander zu kombinieren, finden Sie in [Ausführen von Office-Skripts mit Power Automate](../develop/power-automate-integration.md), oder sehen Sie sich das Lernprogramm [Aufrufen von Skripts aus einem manuellen Power Automate-Datenfluss](excel-power-automate-manual.md) an, um einen Power Automate-Datenfluss zu erstellen, der ein Office-Skript verwendet.</span><span class="sxs-lookup"><span data-stu-id="32f04-208">Learn more about the advantages combining the two platforms in [Run Office Scripts with Power Automate](../develop/power-automate-integration.md) or try the [Call scripts from a manual Power Automate flow](excel-power-automate-manual.md) tutorial to create a Power Automate flow that uses an Office Script.</span></span>
