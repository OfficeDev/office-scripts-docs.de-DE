---
title: Aufzeichnen, Bearbeiten und Erstellen von Office-Skripts in Excel im Web
description: Dies ist ein Lernprogramm zu den Grundlagen von Office-Skripts, einschließlich dem Aufzeichnen von Skripts mithilfe der Aktionsaufzeichnung und dem Schreiben von Daten in eine Arbeitsmappe.
ms.date: 01/06/2021
localization_priority: Priority
ms.openlocfilehash: ae864cc08453a9c8a2538f15ceee1275e131725d
ms.sourcegitcommit: 45ffe3dbd2c834b78592ad35928cf8096f5e80bc
ms.translationtype: HT
ms.contentlocale: de-DE
ms.lasthandoff: 04/14/2021
ms.locfileid: "51754845"
---
# <a name="record-edit-and-create-office-scripts-in-excel-on-the-web"></a><span data-ttu-id="b1137-103">Aufzeichnen, Bearbeiten und Erstellen von Office-Skripts in Excel im Web</span><span class="sxs-lookup"><span data-stu-id="b1137-103">Record, edit, and create Office Scripts in Excel on the web</span></span>

<span data-ttu-id="b1137-104">In diesem Lernprogramm lernen Sie die Grundlagen zum Aufzeichnen, Bearbeiten und Schreiben eines Office-Skripts für Excel im Web kennen.</span><span class="sxs-lookup"><span data-stu-id="b1137-104">This tutorial teaches you the basics of recording, editing, and writing an Office Script for Excel on the web.</span></span> <span data-ttu-id="b1137-105">Zuerst zeichnen Sie ein Skript auf, das ein Arbeitsblatt für Umsatzdatensätze formatiert.</span><span class="sxs-lookup"><span data-stu-id="b1137-105">You'll record a script that applies some formatting to a sales record worksheet.</span></span> <span data-ttu-id="b1137-106">Dann bearbeiten Sie das aufgezeichnete Skript, um weitere Formatierungen anzuwenden, eine Tabelle zu erstellen und diese Tabelle zu sortieren.</span><span class="sxs-lookup"><span data-stu-id="b1137-106">You'll then edit the recorded script to apply more formatting, create a table, and sort that table.</span></span> <span data-ttu-id="b1137-107">Dieses Aufzeichnen-und-Bearbeiten-Muster ist ein wichtiges Tool, um zu sehen, wie Ihre Excel-Aktionen als Code aussehen.</span><span class="sxs-lookup"><span data-stu-id="b1137-107">This record-then-edit pattern is an important tool to see what your Excel actions look like as code.</span></span>

## <a name="prerequisites"></a><span data-ttu-id="b1137-108">Voraussetzungen</span><span class="sxs-lookup"><span data-stu-id="b1137-108">Prerequisites</span></span>

[!INCLUDE [Tutorial prerequisites](../includes/tutorial-prerequisites.md)]

> [!IMPORTANT]
> <span data-ttu-id="b1137-109">Dieses Lernprogramm richtet sich an Anfänger bis Fortgeschrittene mit JavaScript oder TypeScript.</span><span class="sxs-lookup"><span data-stu-id="b1137-109">This tutorial is intended for people with beginner to intermediate-level knowledge of JavaScript or TypeScript.</span></span> <span data-ttu-id="b1137-110">Wenn Sie noch nicht mit JavaScript vertraut sind, empfehlen wir Ihnen, mit dem [Mozilla-JavaScript-Lernprogramm](https://developer.mozilla.org/docs/Web/JavaScript/Guide/Introduction) zu beginnen.</span><span class="sxs-lookup"><span data-stu-id="b1137-110">If you're new to JavaScript, we recommend starting with the [Mozilla JavaScript tutorial](https://developer.mozilla.org/docs/Web/JavaScript/Guide/Introduction).</span></span> <span data-ttu-id="b1137-111">Weitere Informationen über die Skriptumgebung finden Sie unter [ Umgebung des Code Editor für Office-Skripts](../overview/code-editor-environment.md).</span><span class="sxs-lookup"><span data-stu-id="b1137-111">Visit [Office Scripts Code Editor environment](../overview/code-editor-environment.md) to learn more about the script environment.</span></span>

## <a name="add-data-and-record-a-basic-script"></a><span data-ttu-id="b1137-112">Hinzufügen von Daten und Aufzeichnen eines einfachen Skripts</span><span class="sxs-lookup"><span data-stu-id="b1137-112">Add data and record a basic script</span></span>

<span data-ttu-id="b1137-113">Zuerst benötigen wir einige Daten und ein kleines Startskript.</span><span class="sxs-lookup"><span data-stu-id="b1137-113">First, we'll need some data and a small starting script.</span></span>

1. <span data-ttu-id="b1137-114">Erstellen Sie eine neue Arbeitsmappe in Excel im Web.</span><span class="sxs-lookup"><span data-stu-id="b1137-114">Create a new workbook in Excel for the Web.</span></span>
2. <span data-ttu-id="b1137-115">Kopieren Sie die folgenden Obst-Umsatzdaten, und fügen Sie diese beginnend bei Zelle **A1** in das Arbeitsblatt ein.</span><span class="sxs-lookup"><span data-stu-id="b1137-115">Copy the following fruit sales data and paste it into the worksheet, starting at cell **A1**.</span></span>

    |<span data-ttu-id="b1137-116">Obstsorte</span><span class="sxs-lookup"><span data-stu-id="b1137-116">Fruit</span></span> |<span data-ttu-id="b1137-117">2018</span><span class="sxs-lookup"><span data-stu-id="b1137-117">2018</span></span> |<span data-ttu-id="b1137-118">2019</span><span class="sxs-lookup"><span data-stu-id="b1137-118">2019</span></span> |
    |:---|:---|:---|
    |<span data-ttu-id="b1137-119">Orangen</span><span class="sxs-lookup"><span data-stu-id="b1137-119">Oranges</span></span> |<span data-ttu-id="b1137-120">1000</span><span class="sxs-lookup"><span data-stu-id="b1137-120">1000</span></span> |<span data-ttu-id="b1137-121">1200</span><span class="sxs-lookup"><span data-stu-id="b1137-121">1200</span></span> |
    |<span data-ttu-id="b1137-122">Zitronen</span><span class="sxs-lookup"><span data-stu-id="b1137-122">Lemons</span></span> |<span data-ttu-id="b1137-123">800</span><span class="sxs-lookup"><span data-stu-id="b1137-123">800</span></span> |<span data-ttu-id="b1137-124">900</span><span class="sxs-lookup"><span data-stu-id="b1137-124">900</span></span> |
    |<span data-ttu-id="b1137-125">Limetten</span><span class="sxs-lookup"><span data-stu-id="b1137-125">Limes</span></span> |<span data-ttu-id="b1137-126">600</span><span class="sxs-lookup"><span data-stu-id="b1137-126">600</span></span> |<span data-ttu-id="b1137-127">500</span><span class="sxs-lookup"><span data-stu-id="b1137-127">500</span></span> |
    |<span data-ttu-id="b1137-128">Grapefruits</span><span class="sxs-lookup"><span data-stu-id="b1137-128">Grapefruits</span></span> |<span data-ttu-id="b1137-129">900</span><span class="sxs-lookup"><span data-stu-id="b1137-129">900</span></span> |<span data-ttu-id="b1137-130">700</span><span class="sxs-lookup"><span data-stu-id="b1137-130">700</span></span> |

3. <span data-ttu-id="b1137-131">Öffnen Sie die Registerkarte **Automatisieren**. Falls die Registerkarte **Automatisieren** nicht angezeigt wird, überprüfen Sie den Menüband-Überlauf, indem Sie auf den Dropdownpfeil klicken.</span><span class="sxs-lookup"><span data-stu-id="b1137-131">Open the **Automate** tab. If you do not see the **Automate** tab, check the ribbon overflow by pressing the drop-down arrow.</span></span>
4. <span data-ttu-id="b1137-132">Klicken Sie auf die Schaltfläche **Aktionen aufzeichnen**.</span><span class="sxs-lookup"><span data-stu-id="b1137-132">Press the **Record Actions** button.</span></span>
5. <span data-ttu-id="b1137-133">Wählen Sie Zellen **A2:C2** (die Zeile "Orangen") aus, und legen Sie die Füllfarbe auf Orange fest.</span><span class="sxs-lookup"><span data-stu-id="b1137-133">Select cells **A2:C2** (the "Oranges" row) and set the fill color to orange.</span></span>
6. <span data-ttu-id="b1137-134">Beenden Sie die Aufzeichnung, indem Sie die Schaltfläche **Stopp** drücken.</span><span class="sxs-lookup"><span data-stu-id="b1137-134">Stop the recording by pressing the **Stop** button.</span></span>
7. <span data-ttu-id="b1137-135">Geben Sie in das Feld **Skriptnamen** einen einprägsamen Namen ein.</span><span class="sxs-lookup"><span data-stu-id="b1137-135">Fill in the **Script Name** field with a memorable name.</span></span>
8. <span data-ttu-id="b1137-136">*Optional:* Füllen Sie das Feld **Beschreibung** mit einer aussagekräftigen Beschreibung aus.</span><span class="sxs-lookup"><span data-stu-id="b1137-136">*Optional:* Fill in the **Description** field with a meaningful description.</span></span> <span data-ttu-id="b1137-137">Diese wird verwendet, um den Kontext des Skripts bereitzustellen.</span><span class="sxs-lookup"><span data-stu-id="b1137-137">This is used to provide context as to what the script does.</span></span> <span data-ttu-id="b1137-138">Für dieses Lernprogramm können Sie "Farbcodierte Zeilen einer Tabelle" verwenden.</span><span class="sxs-lookup"><span data-stu-id="b1137-138">For the tutorial, you can use "Color-codes rows of a table".</span></span>

   > [!TIP]
   > <span data-ttu-id="b1137-139">Sie können die Beschreibung eines Skripts später im Bereich **Skriptdetails** bearbeiten, der sich im Menü **...** des Code-Editors befindet.</span><span class="sxs-lookup"><span data-stu-id="b1137-139">You can edit a script's description later from the **Script Details** pane, which is located under the Code Editor's **...** menu.</span></span>

9. <span data-ttu-id="b1137-140">Klicken Sie auf die Schaltfläche **Speichern**, um das Skript zu speichern.</span><span class="sxs-lookup"><span data-stu-id="b1137-140">Save the script by pressing the **Save** button.</span></span>

    <span data-ttu-id="b1137-141">Ihr Arbeitsblatt sollte wie folgt aussehen (machen Sie sich keine Sorgen, wenn die Farbe anders ist):</span><span class="sxs-lookup"><span data-stu-id="b1137-141">Your worksheet should look like this (don't worry if the color is different):</span></span>

    :::image type="content" source="../images/tutorial-1.png" alt-text="Ein Arbeitsblatt mit der Datenzeile mit den Verkaufszahlen für Obst, in der die Zeile „Orangen“ in der Farbe Orange hervorgehoben ist.":::

## <a name="edit-an-existing-script"></a><span data-ttu-id="b1137-143">Bearbeiten eines vorhandenen Skripts</span><span class="sxs-lookup"><span data-stu-id="b1137-143">Edit an existing script</span></span>

<span data-ttu-id="b1137-144">Das vorherige Skript hat die Zeile "Orangen" orangefarben eingefärbt.</span><span class="sxs-lookup"><span data-stu-id="b1137-144">The previous script colored the "Oranges" row to be orange.</span></span> <span data-ttu-id="b1137-145">Jetzt fügen wir eine gelbe Zeile für die "Zitronen" hinzu.</span><span class="sxs-lookup"><span data-stu-id="b1137-145">Let's add a yellow row for the "Lemons".</span></span>

1. <span data-ttu-id="b1137-146">Klicken Sie im nun geöffneten **Detailbereich** auf die Schaltfläche **Bearbeiten**.</span><span class="sxs-lookup"><span data-stu-id="b1137-146">From the now-open **Details** pane, press the **Edit** button.</span></span>
2. <span data-ttu-id="b1137-147">Folgendes sollte nun auf dem Bildschirm angezeigt werden:</span><span class="sxs-lookup"><span data-stu-id="b1137-147">You should see something similar to this code:</span></span>

    ```TypeScript
    function main(workbook: ExcelScript.Workbook) {
      // Set fill color to FFC000 for range Sheet1!A2:C2
      let selectedSheet = workbook.getActiveWorksheet();
      selectedSheet.getRange("A2:C2").getFormat().getFill().setColor("FFC000");
    }
    ```

    <span data-ttu-id="b1137-148">Dieser Code ruft das aktuelle Arbeitsblatt aus der Arbeitsmappe ab.</span><span class="sxs-lookup"><span data-stu-id="b1137-148">This code gets the current worksheet from the workbook.</span></span> <span data-ttu-id="b1137-149">Anschließend legt er die Füllfarbe des Bereichs **A2:C2** fest.</span><span class="sxs-lookup"><span data-stu-id="b1137-149">Then, it sets the fill color of the range **A2:C2**.</span></span>

    <span data-ttu-id="b1137-150">Bereiche sind ein wesentliches Element von Office-Skripts in Excel im Web.</span><span class="sxs-lookup"><span data-stu-id="b1137-150">Ranges are a fundamental part of Office Scripts in Excel on the web.</span></span> <span data-ttu-id="b1137-151">Ein Bereich ist ein zusammenhängender, rechteckiger Block von Zellen, die Werte, Formeln und Formatierungen enthalten.</span><span class="sxs-lookup"><span data-stu-id="b1137-151">A range is a contiguous, rectangular block of cells that contains values, formula, and formatting.</span></span> <span data-ttu-id="b1137-152">Hierbei handelt es sich um die grundlegende Struktur von Zellen, durch die Sie die meisten Skript-Aufgaben ausführen werden.</span><span class="sxs-lookup"><span data-stu-id="b1137-152">They are the basic structure of cells through which you'll perform most of your scripting tasks.</span></span>

3. <span data-ttu-id="b1137-153">Fügen Sie am Ende des Skripts die folgende Zeile ein (zwischen dem festgelegten `color` und der schließenden `}`):</span><span class="sxs-lookup"><span data-stu-id="b1137-153">Add the following line to the end of the script (between where the `color` is set and the closing `}`):</span></span>

    ```TypeScript
    selectedSheet.getRange("A3:C3").getFormat().getFill().setColor("yellow");
    ```

4. <span data-ttu-id="b1137-154">Testen Sie das Skript, indem Sie **Ausführen** drücken.</span><span class="sxs-lookup"><span data-stu-id="b1137-154">Test the script by pressing **Run**.</span></span> <span data-ttu-id="b1137-155">Ihre Arbeitsmappe sollte nun wie folgt aussehen:</span><span class="sxs-lookup"><span data-stu-id="b1137-155">Your workbook should now look like this:</span></span>

    :::image type="content" source="../images/tutorial-2.png" alt-text="Ein Arbeitsblatt mit der Datenzeile mit den Verkaufszahlen für Obst, in der die Zeile „Orangen“ in der Farbe Orange und die Zeile „Zitronen“ in der Farbe Gelb hervorgehoben ist.":::

## <a name="create-a-table"></a><span data-ttu-id="b1137-157">Erstellen einer Tabelle</span><span class="sxs-lookup"><span data-stu-id="b1137-157">Create a table</span></span>

<span data-ttu-id="b1137-158">Wandeln wir diese Obst-Umsatzdaten in eine Tabelle um.</span><span class="sxs-lookup"><span data-stu-id="b1137-158">Let's convert this fruit sales data into a table.</span></span> <span data-ttu-id="b1137-159">Wir verwenden unser Skript für den gesamten Prozess.</span><span class="sxs-lookup"><span data-stu-id="b1137-159">We'll use our script for the entire process.</span></span>

1. <span data-ttu-id="b1137-160">Fügen Sie die folgende Zeile am Ende des Skripts hinzu (vor der schließenden `}`):</span><span class="sxs-lookup"><span data-stu-id="b1137-160">Add the following line to the end of the script (before the closing `}`):</span></span>

    ```TypeScript
    let table = selectedSheet.addTable("A1:C5", true);
    ```

2. <span data-ttu-id="b1137-161">Dieser Aufruf gibt ein `Table`-Objekt zurück.</span><span class="sxs-lookup"><span data-stu-id="b1137-161">That call returns a `Table` object.</span></span> <span data-ttu-id="b1137-162">Verwenden wir diese Tabelle zum Sortieren der Daten.</span><span class="sxs-lookup"><span data-stu-id="b1137-162">Let's use that table to sort the data.</span></span> <span data-ttu-id="b1137-163">Wir werden die Daten basierend auf den Werten in der Spalte "Obstsorte" in aufsteigender Reihenfolge sortieren.</span><span class="sxs-lookup"><span data-stu-id="b1137-163">We'll sort the data in ascending order based on the values in the "Fruit" column.</span></span> <span data-ttu-id="b1137-164">Fügen Sie dann die folgende Zeile nach der Tabellenerstellung hinzu:</span><span class="sxs-lookup"><span data-stu-id="b1137-164">Add the following line after the table creation:</span></span>

    ```TypeScript
    table.getSort().apply([{ key: 0, ascending: true }]);
    ```

    <span data-ttu-id="b1137-165">Das Skript sollte wie folgt aussehen:</span><span class="sxs-lookup"><span data-stu-id="b1137-165">Your script should look like this:</span></span>

    ```TypeScript
    function main(workbook: ExcelScript.Workbook) {
        // Set fill color to FFC000 for range Sheet1!A2:C2
        let selectedSheet = workbook.getActiveWorksheet();
        selectedSheet.getRange("A2:C2").getFormat().getFill().setColor("FFC000");
        selectedSheet.getRange("A3:C3").getFormat().getFill().setColor("yellow");
        let table = selectedSheet.addTable("A1:C5", true);
        table.getSort().apply([{ key: 0, ascending: true }]);
    }
    ```

    <span data-ttu-id="b1137-166">Tabellen beinhalten ein `TableSort`-Objekt, auf das über die `Table.getSort`-Methode zugegriffen wird.</span><span class="sxs-lookup"><span data-stu-id="b1137-166">Tables have a `TableSort` object, accessed through the `Table.getSort` method.</span></span> <span data-ttu-id="b1137-167">Sie können auf dieses Objekt Sortierkriterien anwenden.</span><span class="sxs-lookup"><span data-stu-id="b1137-167">You can apply sorting criteria to that object.</span></span> <span data-ttu-id="b1137-168">Die `apply`-Methode bezieht eine Reihe von `SortField`-Objekten ein.</span><span class="sxs-lookup"><span data-stu-id="b1137-168">The `apply` method takes in an array of `SortField` objects.</span></span> <span data-ttu-id="b1137-169">In diesem Fall gibt es nur ein Sortierkriterium, also verwenden wir nur ein `SortField`.</span><span class="sxs-lookup"><span data-stu-id="b1137-169">In this case, we only have one sorting criteria, so we only use one `SortField`.</span></span> <span data-ttu-id="b1137-170">`key: 0` legt für die Spalte mit den die Sortierung bestimmenden Werten "0" fest (dies ist die erste Spalte in der Tabelle, in diesem Fall **A**).</span><span class="sxs-lookup"><span data-stu-id="b1137-170">`key: 0` sets the column with the sort-defining values to "0" (which is the first column on the table, **A** in this case).</span></span> <span data-ttu-id="b1137-171">`ascending: true` sortiert die Daten in aufsteigender Reihenfolge (statt in absteigender Reihenfolge).</span><span class="sxs-lookup"><span data-stu-id="b1137-171">`ascending: true` sorts the data in ascending order (instead of descending order).</span></span>

3. <span data-ttu-id="b1137-172">Führen Sie das Skript aus.</span><span class="sxs-lookup"><span data-stu-id="b1137-172">Run the script.</span></span> <span data-ttu-id="b1137-173">Es sollte eine Tabelle wie die folgende angezeigt werden:</span><span class="sxs-lookup"><span data-stu-id="b1137-173">You should see a table like this:</span></span>

    :::image type="content" source="../images/tutorial-3.png" alt-text="Ein Arbeitsblatt mit der sortierten Tabelle zum Verkauf von Obst.":::

    > [!NOTE]
    > <span data-ttu-id="b1137-175">Wenn Sie das Skript erneut ausführen, wird eine Fehlermeldung angezeigt.</span><span class="sxs-lookup"><span data-stu-id="b1137-175">If you re-run the script, you'll get an error.</span></span> <span data-ttu-id="b1137-176">Der Grund dafür ist, dass Sie keine Tabelle über eine andere Tabelle erstellen können.</span><span class="sxs-lookup"><span data-stu-id="b1137-176">This is because you cannot create a table on top of another table.</span></span> <span data-ttu-id="b1137-177">Sie können das Skript jedoch auf ein anderes Arbeitsblatt oder eine andere Arbeitsmappe anwenden.</span><span class="sxs-lookup"><span data-stu-id="b1137-177">However, you can run the script on a different worksheet or workbook.</span></span>

### <a name="re-run-the-script"></a><span data-ttu-id="b1137-178">Das Skript erneut ausführen</span><span class="sxs-lookup"><span data-stu-id="b1137-178">Re-run the script</span></span>

1. <span data-ttu-id="b1137-179">Erstellen Sie ein neues Arbeitsblatt in der aktuellen Arbeitsmappe.</span><span class="sxs-lookup"><span data-stu-id="b1137-179">Create a new worksheet in the current workbook.</span></span>
2. <span data-ttu-id="b1137-180">Kopieren Sie die Obstdaten am Anfang dieses Lernprogramms, und fügen Sie sie in das neue Arbeitsblatt ein, beginnend bei Zelle **A1**.</span><span class="sxs-lookup"><span data-stu-id="b1137-180">Copy the fruit data from the beginning of the tutorial and paste it into the new worksheet, starting at cell **A1**.</span></span>
3. <span data-ttu-id="b1137-181">Führen Sie das Skript aus.</span><span class="sxs-lookup"><span data-stu-id="b1137-181">Run the script.</span></span>

## <a name="next-steps"></a><span data-ttu-id="b1137-182">Nächste Schritte</span><span class="sxs-lookup"><span data-stu-id="b1137-182">Next steps</span></span>

<span data-ttu-id="b1137-183">Führen Sie das Lernprogramm [Lesen von Arbeitsmappendaten mit Office-Skripts in Excel im Web](excel-read-tutorial.md) aus.</span><span class="sxs-lookup"><span data-stu-id="b1137-183">Complete the [Read workbook data with Office Scripts in Excel on the web](excel-read-tutorial.md) tutorial.</span></span> <span data-ttu-id="b1137-184">Hier erfahren Sie, wie Sie mithilfe eines Office-Skripts Daten aus einer Arbeitsmappe lesen können.</span><span class="sxs-lookup"><span data-stu-id="b1137-184">It teaches you how to read data from a workbook with an Office Script.</span></span>
