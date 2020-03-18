---
title: Grundlegendes zur Skripterstellung für Office-Skripts in Excel im Internet
description: Informationen zum Objektmodell und andere Grundlagen, bevor Sie Office-Skripts schreiben.
ms.date: 01/27/2020
localization_priority: Priority
ms.openlocfilehash: 5a709c16e23c00ffc7ee7949a3cb11459dc2d530
ms.sourcegitcommit: b075eed5a6f275274fbbf6d62633219eac416f26
ms.translationtype: MT
ms.contentlocale: de-DE
ms.lasthandoff: 03/10/2020
ms.locfileid: "42700199"
---
# <a name="scripting-fundamentals-for-office-scripts-in-excel-on-the-web-preview"></a><span data-ttu-id="3fa39-103">Grundlegendes zur Skripterstellung für Office-Skripts in Excel im Internet (Vorschau)</span><span class="sxs-lookup"><span data-stu-id="3fa39-103">Scripting fundamentals for Office Scripts in Excel on the web (preview)</span></span>

<span data-ttu-id="3fa39-104">In diesem Artikel werden die technischen Aspekte von Office-Skripts vorgestellt.</span><span class="sxs-lookup"><span data-stu-id="3fa39-104">This article will introduce you to the technical aspects of Office Scripts.</span></span> <span data-ttu-id="3fa39-105">Sie erfahren, wie die Excel-Objekte zusammenarbeiten und wie der Code-Editor mit einer Arbeitsmappe synchronisiert wird.</span><span class="sxs-lookup"><span data-stu-id="3fa39-105">You'll learn how the Excel objects work together and how the Code Editor synchronizes with a workbook.</span></span>

[!INCLUDE [Preview note](../includes/preview-note.md)]

## <a name="object-model"></a><span data-ttu-id="3fa39-106">Objektmodell</span><span class="sxs-lookup"><span data-stu-id="3fa39-106">Object model</span></span>

<span data-ttu-id="3fa39-107">Um die Excel-APIs zu verstehen, müssen Sie verstehen, wie die Komponenten einer Arbeitsmappe miteinander verwandt sind.</span><span class="sxs-lookup"><span data-stu-id="3fa39-107">To understand the Excel APIs, you must understand how the components of a workbook are related to one another.</span></span>

- <span data-ttu-id="3fa39-108">Eine **Arbeitsmappe** enthält ein oder mehrere **Arbeitsblätter**.</span><span class="sxs-lookup"><span data-stu-id="3fa39-108">A **Workbook** contains one or more **Worksheets**.</span></span>
- <span data-ttu-id="3fa39-109">Ein **Arbeitsblatt** ermöglicht den Zugriff auf Zellen über **Bereichs** Objekte.</span><span class="sxs-lookup"><span data-stu-id="3fa39-109">A **Worksheet** gives access to cells through **Range** objects.</span></span>
- <span data-ttu-id="3fa39-110">Ein **Bereich** stellt eine Gruppe von zusammenhängenden Zellen dar.</span><span class="sxs-lookup"><span data-stu-id="3fa39-110">A **Range** represents a group of contiguous cells.</span></span>
- <span data-ttu-id="3fa39-111">**Bereiche** werden zum Erstellen und Platzieren von **Tabellen**, **Diagrammen**, **Formen**und anderen Daten Visualisierungs-oder Organisationsobjekten verwendet.</span><span class="sxs-lookup"><span data-stu-id="3fa39-111">**Ranges** are used to create and place **Tables**, **Charts**, **Shapes**, and other data visualization or organization objects.</span></span>
- <span data-ttu-id="3fa39-112">Ein **Arbeitsblatt** enthält Auflistungen dieser Datenobjekte, die im einzelnen Blatt vorhanden sind.</span><span class="sxs-lookup"><span data-stu-id="3fa39-112">A **Worksheet** contains collections of those data objects that are present in the individual sheet.</span></span>
- <span data-ttu-id="3fa39-113">**Arbeitsmappen** enthalten Auflistungen einiger dieser Datenobjekte (beispielsweise **Tabellen**) für die gesamte **Arbeitsmappe**.</span><span class="sxs-lookup"><span data-stu-id="3fa39-113">**Workbooks** contain collections of some of those data objects (such as **Tables**) for the entire **Workbook**.</span></span>

### <a name="ranges"></a><span data-ttu-id="3fa39-114">Bereiche</span><span class="sxs-lookup"><span data-stu-id="3fa39-114">Ranges</span></span>

<span data-ttu-id="3fa39-115">Ein Bereich ist eine Gruppe von zusammenhängenden Zellen in der Arbeitsmappe.</span><span class="sxs-lookup"><span data-stu-id="3fa39-115">A range is a group of contiguous cells in the workbook.</span></span> <span data-ttu-id="3fa39-116">Skripts verwenden normalerweise die a1-Notation (beispielsweise **B3** für die einzelne Zelle in Zeile **B** und Spalte **3** oder **C2: F4** für die Zellen von Zeilen **C** bis **F** und Spalten **2** bis **4**), um Bereiche zu definieren.</span><span class="sxs-lookup"><span data-stu-id="3fa39-116">Scripts typically use A1-style notation (e.g. **B3** for the single cell in row **B** and column **3** or **C2:F4** for the cells from rows **C** through **F** and columns **2** through **4**) to define ranges.</span></span>

<span data-ttu-id="3fa39-117">Bereiche weisen drei Haupteigenschaften auf `values`: `formulas`, und `format`.</span><span class="sxs-lookup"><span data-stu-id="3fa39-117">Ranges have three core properties: `values`, `formulas`, and `format`.</span></span> <span data-ttu-id="3fa39-118">Diese Eigenschaften abrufen oder Festlegen der Zellenwerte, Formeln ausgewertet werden, und die visuelle Formatierung der Zellen.</span><span class="sxs-lookup"><span data-stu-id="3fa39-118">These properties get or set the cell values, formulas to be evaluated, and the visual formatting of the cells.</span></span>

#### <a name="range-sample"></a><span data-ttu-id="3fa39-119">Range-Beispiel</span><span class="sxs-lookup"><span data-stu-id="3fa39-119">Range sample</span></span>

<span data-ttu-id="3fa39-120">Im folgenden Beispiel wird gezeigt, wie Sie Umsatzdatensätze erstellen.</span><span class="sxs-lookup"><span data-stu-id="3fa39-120">The following sample shows how to create sales records.</span></span> <span data-ttu-id="3fa39-121">Dieses Skript verwendet `Range` Objekte, um die Werte, Formeln und Formate festzulegen.</span><span class="sxs-lookup"><span data-stu-id="3fa39-121">This script uses `Range` objects to set the values, formulas, and formats.</span></span>

```TypeScript
async function main(context: Excel.RequestContext) {
  // Get the active worksheet.
  let sheet = context.workbook.worksheets.getActiveWorksheet();

  // Create the headers and format them to stand out.
  let headers = [
    ["Product", "Quantity", "Unit Price", "Totals"]
  ];
  let headerRange = sheet.getRange("B2:E2");
  headerRange.values = headers;
  headerRange.format.fill.color = "#4472C4";
  headerRange.format.font.color = "white";

  // Create the product data rows.
  let productData = [
    ["Almonds", 6, 7.5],
    ["Coffee", 20, 34.5],
    ["Chocolate", 10, 9.56],
  ];
  let dataRange = sheet.getRange("B3:D5");
  dataRange.values = productData;

  // Create the formulas to total the amounts sold.
  let totalFormulas = [
    ["=C3 * D3"],
    ["=C4 * D4"],
    ["=C5 * D5"],
    ["=SUM(E3:E5)"]
  ];
  let totalRange = sheet.getRange("E3:E6");
  totalRange.formulas = totalFormulas;
  totalRange.format.font.bold = true;

  // Display the totals as US dollar amounts.
  totalRange.numberFormat = [["$0.00"]];
}
```

<span data-ttu-id="3fa39-122">Beim Ausführen dieses Skripts werden die folgenden Daten im aktuellen Arbeitsblatt erstellt:</span><span class="sxs-lookup"><span data-stu-id="3fa39-122">Running this script creates the following data in the current worksheet:</span></span>

![Ein verkaufsdatensatz mit Wert Zeilen, einer Formelspalte und formatierten Kopfzeilen.](../images/range-sample.png)

### <a name="charts-tables-and-other-data-objects"></a><span data-ttu-id="3fa39-124">Diagramme, Tabellen und andere Datenobjekte</span><span class="sxs-lookup"><span data-stu-id="3fa39-124">Charts, tables, and other data objects</span></span>

<span data-ttu-id="3fa39-125">Skripts können die Datenstrukturen und Visualisierungen in Excel erstellen und bearbeiten.</span><span class="sxs-lookup"><span data-stu-id="3fa39-125">Scripts can create and manipulate the data structures and visualizations within Excel.</span></span> <span data-ttu-id="3fa39-126">Tabellen und Diagramme sind zwei der häufiger verwendeten Objekte, die APIs unterstützen jedoch PivotTables, Formen, Bilder und vieles mehr.</span><span class="sxs-lookup"><span data-stu-id="3fa39-126">Tables and charts are two of the more commonly used objects, but the APIs support PivotTables, shapes, images, and more.</span></span>

#### <a name="creating-a-table"></a><span data-ttu-id="3fa39-127">Erstellen einer Tabelle</span><span class="sxs-lookup"><span data-stu-id="3fa39-127">Creating a table</span></span>

<span data-ttu-id="3fa39-128">Erstellen von Tabellen mithilfe von Daten gefüllten Bereichen.</span><span class="sxs-lookup"><span data-stu-id="3fa39-128">Create tables by using data-filled ranges.</span></span> <span data-ttu-id="3fa39-129">Formatierungs-und Tabellensteuerelemente (wie Filter) werden automatisch auf den Bereich angewendet.</span><span class="sxs-lookup"><span data-stu-id="3fa39-129">Formatting and table controls (such as filters) are automatically applied to the range.</span></span>

<span data-ttu-id="3fa39-130">Das folgende Skript erstellt eine Tabelle mithilfe der Bereiche aus dem vorherigen Beispiel.</span><span class="sxs-lookup"><span data-stu-id="3fa39-130">The following script creates a table using the ranges from the previous sample.</span></span>

```TypeScript
async function main(context: Excel.RequestContext) {
   let sheet = context.workbook.worksheets.getActiveWorksheet();
   sheet.tables.add("B2:E5", true);
}
```

<span data-ttu-id="3fa39-131">Wenn Sie dieses Skript auf dem Arbeitsblatt mit den vorherigen Daten ausführen, wird die folgende Tabelle erstellt:</span><span class="sxs-lookup"><span data-stu-id="3fa39-131">Running this script on the worksheet with the previous data creates the following table:</span></span>

![Eine Tabelle, die aus dem vorherigen Umsatzdatensatz erstellt wurde.](../images/table-sample.png)

#### <a name="creating-a-chart"></a><span data-ttu-id="3fa39-133">Erstellen eines Diagramms</span><span class="sxs-lookup"><span data-stu-id="3fa39-133">Creating a chart</span></span>

<span data-ttu-id="3fa39-134">Erstellen Sie Diagramme, um die Daten in einem Bereich zu visualisieren.</span><span class="sxs-lookup"><span data-stu-id="3fa39-134">Create charts to visualize the data in a range.</span></span> <span data-ttu-id="3fa39-135">Skripts ermöglichen Dutzende von Diagramm Varianten, die jeweils an Ihre Anforderungen angepasst werden können.</span><span class="sxs-lookup"><span data-stu-id="3fa39-135">Scripts allow for dozens of chart varieties, each of which can be customized to suit your needs.</span></span>

<span data-ttu-id="3fa39-136">Das folgende Skript erstellt ein einfaches Säulendiagramm für drei Elemente und platziert es 100 Pixel unter dem oberen Rand des Arbeitsblatts.</span><span class="sxs-lookup"><span data-stu-id="3fa39-136">The following script creates a simple column chart for three items and places it 100 pixels below the top of the worksheet.</span></span>

```TypeScript
async function main(context: Excel.RequestContext) {
  let sheet = context.workbook.worksheets.getActiveWorksheet();
  let chart = sheet.charts.add(Excel.ChartType.columnStacked, sheet.getRange("B3:C5"));
  chart.top = 100;
}
```

<span data-ttu-id="3fa39-137">Wenn Sie dieses Skript auf dem Arbeitsblatt mit der vorherigen Tabelle ausführen, wird das folgende Diagramm erstellt:</span><span class="sxs-lookup"><span data-stu-id="3fa39-137">Running this script on the worksheet with the previous table creates the following chart:</span></span>

![Ein Säulendiagramm mit Mengen von drei Elementen aus dem vorherigen Umsatzdatensatz.](../images/chart-sample.png)

### <a name="further-reading-on-the-object-model"></a><span data-ttu-id="3fa39-139">Weiterführende Informationen zum Objektmodell</span><span class="sxs-lookup"><span data-stu-id="3fa39-139">Further reading on the object model</span></span>

<span data-ttu-id="3fa39-140">Die [Office Scripts-API-Referenzdokumentation](/javascript/api/office-scripts/overview) ist eine umfassende Auflistung der in Office-Skripts verwendeten Objekte.</span><span class="sxs-lookup"><span data-stu-id="3fa39-140">The [Office Scripts API reference documentation](/javascript/api/office-scripts/overview) is a comprehensive listing of the objects used in Office Scripts.</span></span> <span data-ttu-id="3fa39-141">Dort können Sie das Inhaltsverzeichnis verwenden, um zu einer beliebigen Klasse zu navigieren, zu der Sie mehr erfahren möchten.</span><span class="sxs-lookup"><span data-stu-id="3fa39-141">There, you can use the table of contents to navigate to any class you'd like to learn more about.</span></span> <span data-ttu-id="3fa39-142">Es folgen einige häufig angezeigte Seiten.</span><span class="sxs-lookup"><span data-stu-id="3fa39-142">The following are several commonly viewed pages.</span></span>

- [<span data-ttu-id="3fa39-143">Chart</span><span class="sxs-lookup"><span data-stu-id="3fa39-143">Chart</span></span>](/javascript/api/office-scripts/excel/excel.chart)
- [<span data-ttu-id="3fa39-144">Kommentar</span><span class="sxs-lookup"><span data-stu-id="3fa39-144">Comment</span></span>](/javascript/api/office-scripts/excel/excel.comment)
- [<span data-ttu-id="3fa39-145">PivotTable</span><span class="sxs-lookup"><span data-stu-id="3fa39-145">PivotTable</span></span>](/javascript/api/office-scripts/excel/excel.pivottable)
- [<span data-ttu-id="3fa39-146">Range</span><span class="sxs-lookup"><span data-stu-id="3fa39-146">Range</span></span>](/javascript/api/office-scripts/excel/excel.range)
- [<span data-ttu-id="3fa39-147">RangeFormat</span><span class="sxs-lookup"><span data-stu-id="3fa39-147">RangeFormat</span></span>](/javascript/api/office-scripts/excel/excel.rangeformat)
- [<span data-ttu-id="3fa39-148">Shape</span><span class="sxs-lookup"><span data-stu-id="3fa39-148">Shape</span></span>](/javascript/api/office-scripts/excel/excel.shape)
- [<span data-ttu-id="3fa39-149">Table</span><span class="sxs-lookup"><span data-stu-id="3fa39-149">Table</span></span>](/javascript/api/office-scripts/excel/excel.table)
- [<span data-ttu-id="3fa39-150">Workbook</span><span class="sxs-lookup"><span data-stu-id="3fa39-150">Workbook</span></span>](/javascript/api/office-scripts/excel/excel.workbook)
- [<span data-ttu-id="3fa39-151">Arbeitsblatt</span><span class="sxs-lookup"><span data-stu-id="3fa39-151">Worksheet</span></span>](/javascript/api/office-scripts/excel/excel.worksheet)

## <a name="main-function"></a><span data-ttu-id="3fa39-152">`main`Funktion</span><span class="sxs-lookup"><span data-stu-id="3fa39-152">`main` function</span></span>

<span data-ttu-id="3fa39-153">Jedes Office-Skript muss eine `main` Funktion mit der folgenden Signatur enthalten, einschließlich `Excel.RequestContext` der Typendefinition:</span><span class="sxs-lookup"><span data-stu-id="3fa39-153">Every Office Script must contain a `main` function with the following signature, including the `Excel.RequestContext` type definition:</span></span>

```TypeScript
async function main(context: Excel.RequestContext) {
    // Your Excel Script
}
```

<span data-ttu-id="3fa39-154">Der Code in der `main` -Funktion wird ausgeführt, wenn das Skript ausgeführt wird.</span><span class="sxs-lookup"><span data-stu-id="3fa39-154">The code inside the `main` function runs when the script is run.</span></span> <span data-ttu-id="3fa39-155">`main`kann andere Funktionen in Ihrem Skript aufrufen, aber Code, der nicht in einer Funktion enthalten ist, wird nicht ausgeführt.</span><span class="sxs-lookup"><span data-stu-id="3fa39-155">`main` can call other functions in your script, but code that's not contained in a function will not run.</span></span>

## <a name="context"></a><span data-ttu-id="3fa39-156">Kontext</span><span class="sxs-lookup"><span data-stu-id="3fa39-156">Context</span></span>

<span data-ttu-id="3fa39-157">Die `main` Funktion akzeptiert einen `Excel.RequestContext` Parameter mit dem `context`Namen.</span><span class="sxs-lookup"><span data-stu-id="3fa39-157">The `main` function accepts an `Excel.RequestContext` parameter, named `context`.</span></span> <span data-ttu-id="3fa39-158">Denken Sie `context` an die Brücke zwischen Ihrem Skript und der Arbeitsmappe.</span><span class="sxs-lookup"><span data-stu-id="3fa39-158">Think of `context` as the bridge between your script and the workbook.</span></span> <span data-ttu-id="3fa39-159">Ihr Skript greift mit dem `context` Objekt auf die Arbeitsmappe zu und `context` verwendet dieses zum Senden von Daten hin und her.</span><span class="sxs-lookup"><span data-stu-id="3fa39-159">Your script accesses the workbook with the `context` object and uses that `context` to send data back and forth.</span></span>

<span data-ttu-id="3fa39-160">Das `context` Objekt ist erforderlich, da das Skript und Excel in unterschiedlichen Prozessen und Speicherorten aktiv sind.</span><span class="sxs-lookup"><span data-stu-id="3fa39-160">The `context` object is necessary because the script and Excel are running in different processes and locations.</span></span> <span data-ttu-id="3fa39-161">Das Skript muss Änderungen an oder Abfragen von Daten aus der Arbeitsmappe in der Cloud vornehmen.</span><span class="sxs-lookup"><span data-stu-id="3fa39-161">The script will need to make changes to or query data from the workbook in the cloud.</span></span> <span data-ttu-id="3fa39-162">Das `context` Objekt verwaltet diese Transaktionen.</span><span class="sxs-lookup"><span data-stu-id="3fa39-162">The `context` object manages those transactions.</span></span>

## <a name="sync-and-load"></a><span data-ttu-id="3fa39-163">Synchronisieren und laden</span><span class="sxs-lookup"><span data-stu-id="3fa39-163">Sync and Load</span></span>

<span data-ttu-id="3fa39-164">Da das Skript und die Arbeitsmappe an unterschiedlichen Speicherorten ausgeführt werden, dauert die Datenübertragung zwischen den beiden Zeitpunkten.</span><span class="sxs-lookup"><span data-stu-id="3fa39-164">Because your script and workbook run in different locations, any data transfer between the two takes time.</span></span> <span data-ttu-id="3fa39-165">Um die Skript Leistung zu verbessern, werden Befehle in die Warteschlange eingereiht, bis das Skript den `sync` Vorgang explizit aufruft, um das Skript und die Arbeitsmappe zu synchronisieren.</span><span class="sxs-lookup"><span data-stu-id="3fa39-165">To improve script performance, commands are queued up until the script explicitly calls the `sync` operation to synchronize the script and workbook.</span></span> <span data-ttu-id="3fa39-166">Ihr Skript kann unabhängig arbeiten, bis es eine der folgenden Aufgaben erfüllt:</span><span class="sxs-lookup"><span data-stu-id="3fa39-166">Your script can work independently until it needs to do either of the following:</span></span>

- <span data-ttu-id="3fa39-167">Lesen von Daten aus der Arbeitsmappe ( `load` nach einem Vorgang).</span><span class="sxs-lookup"><span data-stu-id="3fa39-167">Read data from the workbook (following a `load` operation).</span></span>
- <span data-ttu-id="3fa39-168">Schreiben von Daten in die Arbeitsmappe (in der Regel, weil das Skript abgeschlossen ist).</span><span class="sxs-lookup"><span data-stu-id="3fa39-168">Write data to the workbook (usually because the script has finished).</span></span>

<span data-ttu-id="3fa39-169">In der folgenden Abbildung ist ein Beispiel für die Ablaufsteuerung zwischen Skript und Arbeitsmappe dargestellt:</span><span class="sxs-lookup"><span data-stu-id="3fa39-169">The following image shows an example control flow between the script and workbook:</span></span>

![Ein Diagramm mit Lese-und Schreibvorgängen, die aus dem Skript zur Arbeitsmappe wechseln.](../images/load-sync.png)

### <a name="sync"></a><span data-ttu-id="3fa39-171">Synchronisierung</span><span class="sxs-lookup"><span data-stu-id="3fa39-171">Sync</span></span>

<span data-ttu-id="3fa39-172">Wenn Ihr Skript Daten lesen oder Daten in die Arbeitsmappe schreiben muss, rufen Sie die `RequestContext.sync` -Methode wie hier gezeigt auf:</span><span class="sxs-lookup"><span data-stu-id="3fa39-172">Whenever your script needs to read data from or write data to the workbook, call the `RequestContext.sync` method as shown here:</span></span>

```TypeScript
await context.sync();
```

> [!NOTE]
> <span data-ttu-id="3fa39-173">`context.sync()`wird implizit aufgerufen, wenn ein Skript beendet wird.</span><span class="sxs-lookup"><span data-stu-id="3fa39-173">`context.sync()` is implicitly called when a script ends.</span></span>

<span data-ttu-id="3fa39-174">Nach Abschluss `sync` des Vorgangs wird die Arbeitsmappe so aktualisiert, dass alle Schreibvorgänge wiedergegeben werden, die von Skript angegeben wurden.</span><span class="sxs-lookup"><span data-stu-id="3fa39-174">After the `sync` operation completes, the workbook updates to reflect any write operations that script has specified.</span></span> <span data-ttu-id="3fa39-175">Ein Schreibvorgang legt eine Eigenschaft für ein Excel-Objekt fest (Beispiels `range.format.fill.color = "red"`Weise) oder ruft eine Methode auf, die eine Eigenschaft ändert ( `range.format.autoFitColumns()`beispielsweise).</span><span class="sxs-lookup"><span data-stu-id="3fa39-175">A write operation is setting any property on a Excel object (e.g. `range.format.fill.color = "red"`) or calling a method that changes a property (e.g., `range.format.autoFitColumns()`).</span></span> <span data-ttu-id="3fa39-176">Der `sync` Vorgang liest auch alle Werte aus der Arbeitsmappe, die das Skript mithilfe eines `load` Vorgangs angefordert hat (wie im nächsten Abschnitt beschrieben).</span><span class="sxs-lookup"><span data-stu-id="3fa39-176">The `sync` operation also reads any values from the workbook that the script requested by using a `load` operation (as discussed in the next section).</span></span>

<span data-ttu-id="3fa39-177">Die Synchronisierung Ihres Skripts mit der Arbeitsmappe kann je nach Netzwerkzeit in Anspruch nehmen.</span><span class="sxs-lookup"><span data-stu-id="3fa39-177">Synchronizing your script with the workbook can take time, depending on your network.</span></span> <span data-ttu-id="3fa39-178">Sie sollten die Anzahl der `sync` Aufrufe minimieren, die dazu beitragen, dass Ihr Skript schnell ausgeführt wird.</span><span class="sxs-lookup"><span data-stu-id="3fa39-178">You should minimize the number of `sync` calls to help your script run fast.</span></span>  

### <a name="load"></a><span data-ttu-id="3fa39-179">Load</span><span class="sxs-lookup"><span data-stu-id="3fa39-179">Load</span></span>

<span data-ttu-id="3fa39-180">Ein Skript muss Daten aus der Arbeitsmappe vor dem Lesen laden.</span><span class="sxs-lookup"><span data-stu-id="3fa39-180">A script must load data from the workbook before reading it.</span></span> <span data-ttu-id="3fa39-181">Das häufiges Laden von Daten aus der gesamten Arbeitsmappe würde jedoch die Geschwindigkeit des Skripts erheblich reduzieren.</span><span class="sxs-lookup"><span data-stu-id="3fa39-181">However, frequently loading data from the entire workbook would greatly reduce the script's speed.</span></span> <span data-ttu-id="3fa39-182">Stattdessen ermöglicht die `load` Methode Ihrem Skript, die Daten, die aus der Arbeitsmappe abgerufen werden sollen, explizit anzugeben.</span><span class="sxs-lookup"><span data-stu-id="3fa39-182">Instead, the `load` method lets your script state specifically which data should be retrieved from the workbook.</span></span>

<span data-ttu-id="3fa39-183">Die `load` -Methode ist für jedes Excel-Objekt verfügbar.</span><span class="sxs-lookup"><span data-stu-id="3fa39-183">The `load` method is available on every Excel object.</span></span> <span data-ttu-id="3fa39-184">Ihr Skript muss die Eigenschaften eines Objekts laden, bevor Sie diese lesen können.</span><span class="sxs-lookup"><span data-stu-id="3fa39-184">Your script must load an object's properties before it can read them.</span></span> <span data-ttu-id="3fa39-185">Andernfalls führt dies zu einem Fehler.</span><span class="sxs-lookup"><span data-stu-id="3fa39-185">Not doing so will result in an error.</span></span>

<span data-ttu-id="3fa39-186">In den folgenden Beispielen wird `Range` ein Objekt verwendet, um die drei `load` Methoden anzuzeigen, mit denen die Methode zum Laden von Daten verwendet werden kann.</span><span class="sxs-lookup"><span data-stu-id="3fa39-186">The following examples use a `Range` object to show the three ways the `load` method can be used to load data.</span></span>

|<span data-ttu-id="3fa39-187">Intent</span><span class="sxs-lookup"><span data-stu-id="3fa39-187">Intent</span></span> |<span data-ttu-id="3fa39-188">Beispielbefehl</span><span class="sxs-lookup"><span data-stu-id="3fa39-188">Example Command</span></span> | <span data-ttu-id="3fa39-189">Effekt</span><span class="sxs-lookup"><span data-stu-id="3fa39-189">Effect</span></span> |
|:--|:--|:--|
|<span data-ttu-id="3fa39-190">Laden einer Eigenschaft</span><span class="sxs-lookup"><span data-stu-id="3fa39-190">Load one property</span></span> |`myRange.load("values");` | <span data-ttu-id="3fa39-191">Lädt eine einzelne Eigenschaft, in diesem Fall das zweidimensionale Array von Werten in diesem Bereich.</span><span class="sxs-lookup"><span data-stu-id="3fa39-191">Loads a single property, in this case the two-dimensional array of values in this range.</span></span> |
|<span data-ttu-id="3fa39-192">Laden mehrerer Eigenschaften</span><span class="sxs-lookup"><span data-stu-id="3fa39-192">Load multiple properties</span></span> |`myRange.load("values, rowCount, columnCount");`| <span data-ttu-id="3fa39-193">Lädt alle Eigenschaften aus einer durch Trennzeichen getrennten Liste, in diesem Beispiel die Werte, die Zeilenanzahl und die Spaltenanzahl.</span><span class="sxs-lookup"><span data-stu-id="3fa39-193">Loads all the properties from a comma-delimited list, in this example the values, row count, and column count.</span></span> |
|<span data-ttu-id="3fa39-194">Alles laden</span><span class="sxs-lookup"><span data-stu-id="3fa39-194">Load everything</span></span> | `myRange.load();`|<span data-ttu-id="3fa39-195">Lädt alle Eigenschaften des Bereichs.</span><span class="sxs-lookup"><span data-stu-id="3fa39-195">Loads all the properties on the range.</span></span> <span data-ttu-id="3fa39-196">Dies ist keine empfohlene Lösung, da Ihr Skript dadurch verlangsamt wird, dass unnötige Daten abgerufen werden.</span><span class="sxs-lookup"><span data-stu-id="3fa39-196">This is not a recommended solution, since it will slow down your script by getting unnecessary data.</span></span> <span data-ttu-id="3fa39-197">Verwenden Sie dies nur, wenn Sie Ihr Skript testen oder wenn Sie jede Eigenschaft aus dem Objekt benötigen.</span><span class="sxs-lookup"><span data-stu-id="3fa39-197">You should only use this while testing your script or if you need every property from the object.</span></span> |

<span data-ttu-id="3fa39-198">Ihr Skript muss aufgerufen `context.sync()` werden, bevor geladene Werte gelesen werden.</span><span class="sxs-lookup"><span data-stu-id="3fa39-198">Your script must call `context.sync()` before reading any loaded values.</span></span>

```TypeScript
let range = selectedSheet.getRange("A1:B3");
range.load ("rowCount"); // Load the property.
await context.sync(); // Synchronize with the workbook to get the property.
console.log(range.rowCount); // Read and log the property value (3).
```

<span data-ttu-id="3fa39-199">Sie können auch Eigenschaften für eine gesamte Auflistung laden.</span><span class="sxs-lookup"><span data-stu-id="3fa39-199">You can also load properties across an entire collection.</span></span> <span data-ttu-id="3fa39-200">Jedes Collection-Objekt verfügt `items` über eine Eigenschaft, die ein Array mit den Objekten in dieser Auflistung darstellt.</span><span class="sxs-lookup"><span data-stu-id="3fa39-200">Every collection object has an `items` property that is an array containing the objects in that collection.</span></span> <span data-ttu-id="3fa39-201">Verwenden `items` als Anfang eines hierarchischen Aufrufs (`items\myProperty`), `load` um die angegebenen Eigenschaften für jedes dieser Elemente zu laden.</span><span class="sxs-lookup"><span data-stu-id="3fa39-201">Using `items` as the start of a hierarchical call (`items\myProperty`) to `load` loads the specified properties on each of those items.</span></span> <span data-ttu-id="3fa39-202">Im folgenden Beispiel wird die `resolved` Eigenschaft für jedes `Comment` Objekt im `CommentCollection` Objekt eines Arbeitsblatts geladen.</span><span class="sxs-lookup"><span data-stu-id="3fa39-202">The following example loads the `resolved` property on every `Comment` object in the `CommentCollection` object of a worksheet.</span></span>

```TypeScript
let comments = selectedSheet.comments;
comments.load("items/resolved"); // Load the `resolved` property from every comment in this collection.
await context.sync(); // Synchronize with the workbook to get the properties.
```

> [!TIP]
> <span data-ttu-id="3fa39-203">Weitere Informationen zum Arbeiten mit Auflistungen in Office-Skripts finden Sie im [Abschnitt "Array" im Artikel Verwenden integrierter JavaScript-Objekte in Office-Skripts](javascript-objects.md#array) .</span><span class="sxs-lookup"><span data-stu-id="3fa39-203">To learn more about working with collections in Office Scripts, see the [Array section of the Using built-in JavaScript objects in Office Scripts](javascript-objects.md#array) article.</span></span>

## <a name="see-also"></a><span data-ttu-id="3fa39-204">Siehe auch</span><span class="sxs-lookup"><span data-stu-id="3fa39-204">See also</span></span>

- [<span data-ttu-id="3fa39-205">Aufzeichnen, bearbeiten und Erstellen von Office-Skripts in Excel im Internet</span><span class="sxs-lookup"><span data-stu-id="3fa39-205">Record, edit, and create Office Scripts in Excel on the web</span></span>](../tutorials/excel-tutorial.md)
- [<span data-ttu-id="3fa39-206">Lesen von Arbeitsmappen-Daten mit Office-Skripts in Excel im Web</span><span class="sxs-lookup"><span data-stu-id="3fa39-206">Read workbook data with Office Scripts in Excel on the web</span></span>](../tutorials/excel-read-tutorial.md)
- [<span data-ttu-id="3fa39-207">Office Scripts-API-Referenz</span><span class="sxs-lookup"><span data-stu-id="3fa39-207">Office Scripts API reference</span></span>](/javascript/api/office-scripts/overview)
- [<span data-ttu-id="3fa39-208">Verwenden integrierter JavaScript-Objekte in Office-Skripts</span><span class="sxs-lookup"><span data-stu-id="3fa39-208">Using built-in JavaScript objects in Office Scripts</span></span>](javascript-objects.md)
