---
title: Grundlegendes zur Skripterstellung für Office-Skripts in Excel im Web
description: Informationen zu Objektmodellen und andere Grundlagen, die Sie vor dem Schreiben von Office-Skripts benötigen.
ms.date: 04/24/2020
localization_priority: Priority
ms.openlocfilehash: 8449654e359f665677f3d416a8e28fa4d6930f26
ms.sourcegitcommit: 350bd2447f616fa87bb23ac826c7731fb813986b
ms.translationtype: HT
ms.contentlocale: de-DE
ms.lasthandoff: 04/28/2020
ms.locfileid: "43919798"
---
# <a name="scripting-fundamentals-for-office-scripts-in-excel-on-the-web-preview"></a><span data-ttu-id="8397c-103">Grundlegendes zur Skripterstellung für Office-Skripts in Excel im Web (Vorschau)</span><span class="sxs-lookup"><span data-stu-id="8397c-103">Scripting fundamentals for Office Scripts in Excel on the web (preview)</span></span>

<span data-ttu-id="8397c-104">In diesem Artikel werden die technischen Aspekte von Office-Skripts vorgestellt.</span><span class="sxs-lookup"><span data-stu-id="8397c-104">This article will introduce you to the technical aspects of Office Scripts.</span></span> <span data-ttu-id="8397c-105">Sie erfahren, wie die einzelnen Excel-Objekte zusammenarbeiten und wie der Code-Editor mit einer Arbeitsmappe synchronisiert wird.</span><span class="sxs-lookup"><span data-stu-id="8397c-105">You'll learn how the Excel objects work together and how the Code Editor synchronizes with a workbook.</span></span>

[!INCLUDE [Preview note](../includes/preview-note.md)]

## <a name="object-model"></a><span data-ttu-id="8397c-106">Objektmodell</span><span class="sxs-lookup"><span data-stu-id="8397c-106">Object model</span></span>

<span data-ttu-id="8397c-107">Um die Excel-APIs zu verstehen, müssen Sie wissen, wie die Komponenten einer Arbeitsmappe miteinander verknüpft sind.</span><span class="sxs-lookup"><span data-stu-id="8397c-107">To understand the Excel APIs, you must understand how the components of a workbook are related to one another.</span></span>

- <span data-ttu-id="8397c-108">Eine **Arbeitsmappe** enthält mindestens ein **Arbeitsblatt**.</span><span class="sxs-lookup"><span data-stu-id="8397c-108">A **Workbook** contains one or more **Worksheets**.</span></span>
- <span data-ttu-id="8397c-109">Ein **Arbeitsblatt** ermöglicht den Zugriff auf Zellen über **Bereichsobjekte**.</span><span class="sxs-lookup"><span data-stu-id="8397c-109">A **Worksheet** gives access to cells through **Range** objects.</span></span>
- <span data-ttu-id="8397c-110">Ein **Bereich** besteht aus einer Gruppe zusammenhängender Zellen.</span><span class="sxs-lookup"><span data-stu-id="8397c-110">A **Range** represents a group of contiguous cells.</span></span>
- <span data-ttu-id="8397c-111">**Bereiche** werden verwendet, um **Tabellen**, **Diagramme**, **Formen** sowie andere Objekte für die Datenvisualisierung oder -organisation zu erstellen und zu platzieren.</span><span class="sxs-lookup"><span data-stu-id="8397c-111">**Ranges** are used to create and place **Tables**, **Charts**, **Shapes**, and other data visualization or organization objects.</span></span>
- <span data-ttu-id="8397c-112">Ein **Arbeitsblatt** enthält Sammlungen dieser Datenobjekte, die auf dem jeweiligen Blatt vorhanden sind.</span><span class="sxs-lookup"><span data-stu-id="8397c-112">A **Worksheet** contains collections of those data objects that are present in the individual sheet.</span></span>
- <span data-ttu-id="8397c-113">**Arbeitsmappen** enthalten Sammlungen einiger dieser Datenobjekte (z. B. **Tabellen**) für die gesamte **Arbeitsmappe**.</span><span class="sxs-lookup"><span data-stu-id="8397c-113">**Workbooks** contain collections of some of those data objects (such as **Tables**) for the entire **Workbook**.</span></span>

### <a name="ranges"></a><span data-ttu-id="8397c-114">Bereiche</span><span class="sxs-lookup"><span data-stu-id="8397c-114">Ranges</span></span>

<span data-ttu-id="8397c-115">Ein Bereich ist eine Gruppe zusammenhängender Zellen in der Arbeitsmappe.</span><span class="sxs-lookup"><span data-stu-id="8397c-115">A range is a group of contiguous cells in the workbook.</span></span> <span data-ttu-id="8397c-116">In Skripts wird in der Regel eine Notation im A1-Format verwendet (z. B. **B3** für die einzelne Zelle in Spalte **B** und Zeile **3** oder **C2:F4** für die Zellen in den Spalten **C** bis **F** und den Zeilen **2** bis **4**), um Bereiche zu definieren.</span><span class="sxs-lookup"><span data-stu-id="8397c-116">Scripts typically use A1-style notation (e.g. **B3** for the single cell in column **B** and row **3** or **C2:F4** for the cells from columns **C** through **F** and rows **2** through **4**) to define ranges.</span></span>

<span data-ttu-id="8397c-117">Bereiche besitzen drei Haupteigenschaften: `values`, `formulas` und `format`.</span><span class="sxs-lookup"><span data-stu-id="8397c-117">Ranges have three core properties: `values`, `formulas`, and `format`.</span></span> <span data-ttu-id="8397c-118">Durch diese Eigenschaften können die Zellwerte, die zu prüfenden Formeln sowie die visuelle Formatierung der Zellen abgerufen oder festgelegt werden.</span><span class="sxs-lookup"><span data-stu-id="8397c-118">These properties get or set the cell values, formulas to be evaluated, and the visual formatting of the cells.</span></span>

#### <a name="range-sample"></a><span data-ttu-id="8397c-119">Beispiel für einen Bereich</span><span class="sxs-lookup"><span data-stu-id="8397c-119">Range sample</span></span>

<span data-ttu-id="8397c-120">Das folgende Beispiel zeigt, wie Sie Verkaufsdatensätze erstellen können.</span><span class="sxs-lookup"><span data-stu-id="8397c-120">The following sample shows how to create sales records.</span></span> <span data-ttu-id="8397c-121">In diesem Skript werden `Range`-Objekte zum Festlegen der Werte, Formeln und Formate verwendet.</span><span class="sxs-lookup"><span data-stu-id="8397c-121">This script uses `Range` objects to set the values, formulas, and formats.</span></span>

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

<span data-ttu-id="8397c-122">Wenn Sie dieses Skript ausführen, werden die folgenden Daten im aktuellen Arbeitsblatt erstellt:</span><span class="sxs-lookup"><span data-stu-id="8397c-122">Running this script creates the following data in the current worksheet:</span></span>

![Ein Umsatzdatensatz mit Wert-Zeilen, einer Formelspalte sowie formatierten Überschriften.](../images/range-sample.png)

### <a name="charts-tables-and-other-data-objects"></a><span data-ttu-id="8397c-124">Diagramme, Tabellen und andere Datenobjekte</span><span class="sxs-lookup"><span data-stu-id="8397c-124">Charts, tables, and other data objects</span></span>

<span data-ttu-id="8397c-125">Skripts können die Datenstrukturen und -visualisierungen in Excel erstellen und ändern.</span><span class="sxs-lookup"><span data-stu-id="8397c-125">Scripts can create and manipulate the data structures and visualizations within Excel.</span></span> <span data-ttu-id="8397c-126">Tabellen und Diagramme sind zwei der am häufigsten verwendeten Objekte, die APIs unterstützen aber auch PivotTables, Formen, Bilder und vieles mehr.</span><span class="sxs-lookup"><span data-stu-id="8397c-126">Tables and charts are two of the more commonly used objects, but the APIs support PivotTables, shapes, images, and more.</span></span>

#### <a name="creating-a-table"></a><span data-ttu-id="8397c-127">Erstellen einer Tabelle</span><span class="sxs-lookup"><span data-stu-id="8397c-127">Creating a table</span></span>

<span data-ttu-id="8397c-128">Erstellen Sie Tabellen mithilfe von mit Daten ausgefüllten Bereichen.</span><span class="sxs-lookup"><span data-stu-id="8397c-128">Create tables by using data-filled ranges.</span></span> <span data-ttu-id="8397c-129">Auf den Bereich werden automatisch Formatierungs- und Tabellen-Steuerelemente (wie z. B. Filter) angewendet.</span><span class="sxs-lookup"><span data-stu-id="8397c-129">Formatting and table controls (such as filters) are automatically applied to the range.</span></span>

<span data-ttu-id="8397c-130">Durch das folgende Skript wird eine Tabelle auf Grundlage der Bereiche aus dem vorherigen Beispiel erstellt.</span><span class="sxs-lookup"><span data-stu-id="8397c-130">The following script creates a table using the ranges from the previous sample.</span></span>

```TypeScript
async function main(context: Excel.RequestContext) {
   let sheet = context.workbook.worksheets.getActiveWorksheet();
   sheet.tables.add("B2:E5", true);
}
```

<span data-ttu-id="8397c-131">Wenn Sie dieses Skript auf das Arbeitsblatt mit den vorherigen Daten anwenden, wird die folgende Tabelle erstellt:</span><span class="sxs-lookup"><span data-stu-id="8397c-131">Running this script on the worksheet with the previous data creates the following table:</span></span>

![Eine Tabelle aus dem vorherigen Umsatzeintrag.](../images/table-sample.png)

#### <a name="creating-a-chart"></a><span data-ttu-id="8397c-133">Erstellen eines Diagramms</span><span class="sxs-lookup"><span data-stu-id="8397c-133">Creating a chart</span></span>

<span data-ttu-id="8397c-134">Erstellen Sie Diagramme, um die Daten in einem Bereich darzustellen.</span><span class="sxs-lookup"><span data-stu-id="8397c-134">Create charts to visualize the data in a range.</span></span> <span data-ttu-id="8397c-135">In Skripts sind Dutzende von Diagrammvarianten zulässig, die jeweils an Ihre Anforderungen angepasst werden können.</span><span class="sxs-lookup"><span data-stu-id="8397c-135">Scripts allow for dozens of chart varieties, each of which can be customized to suit your needs.</span></span>

<span data-ttu-id="8397c-136">Mit dem folgenden Skript wird ein einfaches Säulendiagramm für drei Elemente erstellt und 100 Pixel unterhalb des oberen Rands des Arbeitsblatts platziert.</span><span class="sxs-lookup"><span data-stu-id="8397c-136">The following script creates a simple column chart for three items and places it 100 pixels below the top of the worksheet.</span></span>

```TypeScript
async function main(context: Excel.RequestContext) {
  let sheet = context.workbook.worksheets.getActiveWorksheet();
  let chart = sheet.charts.add(Excel.ChartType.columnStacked, sheet.getRange("B3:C5"));
  chart.top = 100;
}
```

<span data-ttu-id="8397c-137">Wenn Sie dieses Skript auf das Arbeitsblatt mit der vorherigen Tabelle anwenden, wird das folgende Diagramm erstellt:</span><span class="sxs-lookup"><span data-stu-id="8397c-137">Running this script on the worksheet with the previous table creates the following chart:</span></span>

![Ein Säulendiagramm, in dem die Mengenangaben zu drei Elementen aus dem vorherigen Umsatzeintrag angezeigt werden.](../images/chart-sample.png)

### <a name="further-reading-on-the-object-model"></a><span data-ttu-id="8397c-139">Weitere Informationen zum Objektmodell</span><span class="sxs-lookup"><span data-stu-id="8397c-139">Further reading on the object model</span></span>

<span data-ttu-id="8397c-140">Die [Referenzdokumentation zur Office Scripts-API](/javascript/api/office-scripts/overview) besteht aus einer umfassender Liste der Objekte, die in Office-Skripts verwendet werden.</span><span class="sxs-lookup"><span data-stu-id="8397c-140">The [Office Scripts API reference documentation](/javascript/api/office-scripts/overview) is a comprehensive listing of the objects used in Office Scripts.</span></span> <span data-ttu-id="8397c-141">Dort können Sie über das Inhaltsverzeichnis zu jedem Thema navigieren, über das Sie mehr erfahren möchten.</span><span class="sxs-lookup"><span data-stu-id="8397c-141">There, you can use the table of contents to navigate to any class you'd like to learn more about.</span></span> <span data-ttu-id="8397c-142">Nachstehend finden Sie einige häufig besuchte Seiten.</span><span class="sxs-lookup"><span data-stu-id="8397c-142">The following are several commonly viewed pages.</span></span>

- [<span data-ttu-id="8397c-143">Chart</span><span class="sxs-lookup"><span data-stu-id="8397c-143">Chart</span></span>](/javascript/api/office-scripts/excel/excel.chart)
- [<span data-ttu-id="8397c-144">Kommentar</span><span class="sxs-lookup"><span data-stu-id="8397c-144">Comment</span></span>](/javascript/api/office-scripts/excel/excel.comment)
- [<span data-ttu-id="8397c-145">PivotTable</span><span class="sxs-lookup"><span data-stu-id="8397c-145">PivotTable</span></span>](/javascript/api/office-scripts/excel/excel.pivottable)
- [<span data-ttu-id="8397c-146">Range</span><span class="sxs-lookup"><span data-stu-id="8397c-146">Range</span></span>](/javascript/api/office-scripts/excel/excel.range)
- [<span data-ttu-id="8397c-147">RangeFormat</span><span class="sxs-lookup"><span data-stu-id="8397c-147">RangeFormat</span></span>](/javascript/api/office-scripts/excel/excel.rangeformat)
- [<span data-ttu-id="8397c-148">Form</span><span class="sxs-lookup"><span data-stu-id="8397c-148">Shape</span></span>](/javascript/api/office-scripts/excel/excel.shape)
- [<span data-ttu-id="8397c-149">Table</span><span class="sxs-lookup"><span data-stu-id="8397c-149">Table</span></span>](/javascript/api/office-scripts/excel/excel.table)
- [<span data-ttu-id="8397c-150">Workbook</span><span class="sxs-lookup"><span data-stu-id="8397c-150">Workbook</span></span>](/javascript/api/office-scripts/excel/excel.workbook)
- [<span data-ttu-id="8397c-151">Worksheet</span><span class="sxs-lookup"><span data-stu-id="8397c-151">Worksheet</span></span>](/javascript/api/office-scripts/excel/excel.worksheet)

## <a name="main-function"></a><span data-ttu-id="8397c-152">Die `main`-Funktion</span><span class="sxs-lookup"><span data-stu-id="8397c-152">`main` function</span></span>

<span data-ttu-id="8397c-153">Jedes Office-Skript muss eine `main`-Funktion mit der folgenden Signatur enthalten, einschließlich der `Excel.RequestContext`-Typdefinition:</span><span class="sxs-lookup"><span data-stu-id="8397c-153">Every Office Script must contain a `main` function with the following signature, including the `Excel.RequestContext` type definition:</span></span>

```TypeScript
async function main(context: Excel.RequestContext) {
    // Your Excel Script
}
```

<span data-ttu-id="8397c-154">Der Code innerhalb der `main`-Funktion wird beim Ausführen des Skripts ausgeführt.</span><span class="sxs-lookup"><span data-stu-id="8397c-154">The code inside the `main` function runs when the script is run.</span></span> <span data-ttu-id="8397c-155">`main` kann andere Funktionen in Ihrem Skript aufrufen, Code, der nicht in einer Funktion enthalten ist, wird jedoch nicht ausgeführt.</span><span class="sxs-lookup"><span data-stu-id="8397c-155">`main` can call other functions in your script, but code that's not contained in a function will not run.</span></span>

## <a name="context"></a><span data-ttu-id="8397c-156">Context</span><span class="sxs-lookup"><span data-stu-id="8397c-156">Context</span></span>

<span data-ttu-id="8397c-157">Die `main`-Funktion lässt einen `Excel.RequestContext`-Parameter namens `context` zu.</span><span class="sxs-lookup"><span data-stu-id="8397c-157">The `main` function accepts an `Excel.RequestContext` parameter, named `context`.</span></span> <span data-ttu-id="8397c-158">Stellen Sie sich `context` als die Brücke zwischen Ihrem Skript und der Arbeitsmappe vor.</span><span class="sxs-lookup"><span data-stu-id="8397c-158">Think of `context` as the bridge between your script and the workbook.</span></span> <span data-ttu-id="8397c-159">Das Skript greift auf die Arbeitsmappe mit dem `context`-Objekt zu und verwendet diesen `context` zum Hin- und Hersenden von Daten.</span><span class="sxs-lookup"><span data-stu-id="8397c-159">Your script accesses the workbook with the `context` object and uses that `context` to send data back and forth.</span></span>

<span data-ttu-id="8397c-160">Das `context`-Objekt ist erforderlich, weil das Skript und Excel in unterschiedlichen Prozessen und Speicherorten ausgeführt werden.</span><span class="sxs-lookup"><span data-stu-id="8397c-160">The `context` object is necessary because the script and Excel are running in different processes and locations.</span></span> <span data-ttu-id="8397c-161">Das Skript muss Änderungen an den Daten in der Arbeitsmappe in der Cloud vornehmen oder diese abrufen können.</span><span class="sxs-lookup"><span data-stu-id="8397c-161">The script will need to make changes to or query data from the workbook in the cloud.</span></span> <span data-ttu-id="8397c-162">Das `context`-Objekt verwaltet diese Transaktionen.</span><span class="sxs-lookup"><span data-stu-id="8397c-162">The `context` object manages those transactions.</span></span>

## <a name="sync-and-load"></a><span data-ttu-id="8397c-163">Synchronisieren und Laden</span><span class="sxs-lookup"><span data-stu-id="8397c-163">Sync and Load</span></span>

<span data-ttu-id="8397c-164">Da Ihr Skript und die Arbeitsmappe an unterschiedlichen Orten ausgeführt werden, dauert die Datenübertragung zwischen diesen etwas.</span><span class="sxs-lookup"><span data-stu-id="8397c-164">Because your script and workbook run in different locations, any data transfer between the two takes time.</span></span> <span data-ttu-id="8397c-165">Um die Skriptleistung zu verbessern, werden Befehle in die Warteschlange gesetzt, bis das Skript explizit den `sync`-Vorgang aufruft, um das Skript und die Arbeitsmappe miteinander zu synchronisieren.</span><span class="sxs-lookup"><span data-stu-id="8397c-165">To improve script performance, commands are queued up until the script explicitly calls the `sync` operation to synchronize the script and workbook.</span></span> <span data-ttu-id="8397c-166">Ihr Skript kann unabhängig funktionieren, bis es eine der folgenden Aktionen durchführen muss:</span><span class="sxs-lookup"><span data-stu-id="8397c-166">Your script can work independently until it needs to do either of the following:</span></span>

- <span data-ttu-id="8397c-167">Daten aus der Arbeitsmappe lesen (nach einem `load`-Vorgang oder einer Methode, die ein[ClientResultat](/javascript/api/office-scripts/excel/excel.clientresult) zurückgibt).</span><span class="sxs-lookup"><span data-stu-id="8397c-167">Read data from the workbook (following a `load` operation or method that returns a [ClientResult](/javascript/api/office-scripts/excel/excel.clientresult)).</span></span>
- <span data-ttu-id="8397c-168">Daten in die Arbeitsmappe schreiben (in der Regel, weil das Skript abgeschlossen wurde).</span><span class="sxs-lookup"><span data-stu-id="8397c-168">Write data to the workbook (usually because the script has finished).</span></span>

<span data-ttu-id="8397c-169">In der folgenden Abbildung wird ein Beispiel für eine Ablaufsteuerung zwischen dem Skript und der Arbeitsmappe dargestellt:</span><span class="sxs-lookup"><span data-stu-id="8397c-169">The following image shows an example control flow between the script and workbook:</span></span>

![Ein Diagramm mit Lese- und Schreibvorgängen, die vom Skript in der Arbeitsmappe ausgeführt werden.](../images/load-sync.png)

### <a name="sync"></a><span data-ttu-id="8397c-171">Synchronisierung</span><span class="sxs-lookup"><span data-stu-id="8397c-171">Sync</span></span>

<span data-ttu-id="8397c-172">Wenn das Skript Daten aus der Arbeitsmappe auslesen oder in diese schreiben muss, rufen Sie die `RequestContext.sync`-Methode wie hier dargestellt auf:</span><span class="sxs-lookup"><span data-stu-id="8397c-172">Whenever your script needs to read data from or write data to the workbook, call the `RequestContext.sync` method as shown here:</span></span>

```TypeScript
await context.sync();
```

> [!NOTE]
> <span data-ttu-id="8397c-173">`context.sync()` wird implizit aufgerufen, wenn ein Skript endet.</span><span class="sxs-lookup"><span data-stu-id="8397c-173">`context.sync()` is implicitly called when a script ends.</span></span>

<span data-ttu-id="8397c-174">Nachdem der `sync`-Vorgang abgeschlossen ist, wird die Arbeitsmappe entsprechend den Schreibvorgängen aktualisiert, die vom Skript angegeben wurden.</span><span class="sxs-lookup"><span data-stu-id="8397c-174">After the `sync` operation completes, the workbook updates to reflect any write operations that script has specified.</span></span> <span data-ttu-id="8397c-175">Bei einem Schreibvorgang wird eine beliebige Eigenschaft eines Excel-Objekts festgelegt (z. B. `range.format.fill.color = "red"`) oder eine Methode aufgerufen, über die eine Eigenschaft geändert wird (z. B. `range.format.autoFitColumns()`).</span><span class="sxs-lookup"><span data-stu-id="8397c-175">A write operation is setting any property on a Excel object (e.g. `range.format.fill.color = "red"`) or calling a method that changes a property (e.g., `range.format.autoFitColumns()`).</span></span> <span data-ttu-id="8397c-176">Der `sync`-Vorgang liest auch alle Werte aus der Arbeitsmappe, die das Skript angefordert hat, indem es einen `load`-Vorgang oder eine Methode verwendet, die ein `ClientResult` zurückgibt (wie in den nächsten Abschnitten besprochen).</span><span class="sxs-lookup"><span data-stu-id="8397c-176">The `sync` operation also reads any values from the workbook that the script requested by using a `load` operation or a method that returns a `ClientResult` (as discussed in the next sections).</span></span>

<span data-ttu-id="8397c-177">Je nach Netzwerk kann es einige Zeit dauern, bis das Skript mit der Arbeitsmappe synchronisiert wurde.</span><span class="sxs-lookup"><span data-stu-id="8397c-177">Synchronizing your script with the workbook can take time, depending on your network.</span></span> <span data-ttu-id="8397c-178">Sie sollten die Anzahl von `sync`-Aufrufen minimieren, damit das Ausführen des Skripts möglichst schnell geht.</span><span class="sxs-lookup"><span data-stu-id="8397c-178">You should minimize the number of `sync` calls to help your script run fast.</span></span>  

### <a name="load"></a><span data-ttu-id="8397c-179">Laden</span><span class="sxs-lookup"><span data-stu-id="8397c-179">Load</span></span>

<span data-ttu-id="8397c-180">Ein Skript muss Daten aus der Arbeitsmappe laden, bevor es sie liest.</span><span class="sxs-lookup"><span data-stu-id="8397c-180">A script must load data from the workbook before reading it.</span></span> <span data-ttu-id="8397c-181">Das häufige Laden von Daten aus der gesamten Arbeitsmappe würde jedoch die Geschwindigkeit des Skripts erheblich verringern.</span><span class="sxs-lookup"><span data-stu-id="8397c-181">However, frequently loading data from the entire workbook would greatly reduce the script's speed.</span></span> <span data-ttu-id="8397c-182">Stattdessen können Sie mit der `load`-Methode präzise angeben, welche Daten aus der Arbeitsmappe abgerufen werden sollen.</span><span class="sxs-lookup"><span data-stu-id="8397c-182">Instead, the `load` method lets your script state specifically which data should be retrieved from the workbook.</span></span>

<span data-ttu-id="8397c-183">Die `load`-Methode ist für jedes Excel-Objekt verfügbar.</span><span class="sxs-lookup"><span data-stu-id="8397c-183">The `load` method is available on every Excel object.</span></span> <span data-ttu-id="8397c-184">Ihr Skript muss die Eigenschaften eines Objekts laden, bevor es sie lesen kann.</span><span class="sxs-lookup"><span data-stu-id="8397c-184">Your script must load an object's properties before it can read them.</span></span> <span data-ttu-id="8397c-185">Andernfalls wird ein Fehler zurückgegeben.</span><span class="sxs-lookup"><span data-stu-id="8397c-185">Not doing so will result in an error.</span></span>

<span data-ttu-id="8397c-186">In den folgenden Beispielen wird ein `Range`-Objekt verwendet, um die drei Arten darzustellen, wie die `load`-Methode zum Laden von Daten verwendet werden kann.</span><span class="sxs-lookup"><span data-stu-id="8397c-186">The following examples use a `Range` object to show the three ways the `load` method can be used to load data.</span></span>

|<span data-ttu-id="8397c-187">Absicht</span><span class="sxs-lookup"><span data-stu-id="8397c-187">Intent</span></span> |<span data-ttu-id="8397c-188">Beispielbefehl</span><span class="sxs-lookup"><span data-stu-id="8397c-188">Example Command</span></span> | <span data-ttu-id="8397c-189">Auswirkung</span><span class="sxs-lookup"><span data-stu-id="8397c-189">Effect</span></span> |
|:--|:--|:--|
|<span data-ttu-id="8397c-190">Laden einer Eigenschaft</span><span class="sxs-lookup"><span data-stu-id="8397c-190">Load one property</span></span> |`myRange.load("values");` | <span data-ttu-id="8397c-191">Lädt eine einzelne Eigenschaft, in diesem Fall den zweidimensionalen Wertearray in diesem Bereich.</span><span class="sxs-lookup"><span data-stu-id="8397c-191">Loads a single property, in this case the two-dimensional array of values in this range.</span></span> |
|<span data-ttu-id="8397c-192">Laden mehrerer Eigenschaften</span><span class="sxs-lookup"><span data-stu-id="8397c-192">Load multiple properties</span></span> |`myRange.load("values, rowCount, columnCount");`| <span data-ttu-id="8397c-193">Lädt alle Eigenschaften aus einer durch Kommas getrennten Liste, in diesem Beispiel die Werte, die Zeilenanzahl und die Spaltenanzahl.</span><span class="sxs-lookup"><span data-stu-id="8397c-193">Loads all the properties from a comma-delimited list, in this example the values, row count, and column count.</span></span> |
|<span data-ttu-id="8397c-194">Alles laden</span><span class="sxs-lookup"><span data-stu-id="8397c-194">Load everything</span></span> | `myRange.load();`|<span data-ttu-id="8397c-195">Lädt alle Eigenschaften des Zellbereichs.</span><span class="sxs-lookup"><span data-stu-id="8397c-195">Loads all the properties on the range.</span></span> <span data-ttu-id="8397c-196">Dies ist keine empfohlene Lösung, da das Skript durch das Abrufen unnötiger Daten verlangsamt wird.</span><span class="sxs-lookup"><span data-stu-id="8397c-196">This is not a recommended solution, since it will slow down your script by getting unnecessary data.</span></span> <span data-ttu-id="8397c-197">Sie sollten diesen Wert nur verwenden, wenn Sie das Skript testen, oder wenn Sie alle Eigenschaften des Objekts benötigen.</span><span class="sxs-lookup"><span data-stu-id="8397c-197">You should only use this while testing your script or if you need every property from the object.</span></span> |

<span data-ttu-id="8397c-198">Ihr Skript muss `context.sync()` aufrufen, bevor es geladene Werte ausliest.</span><span class="sxs-lookup"><span data-stu-id="8397c-198">Your script must call `context.sync()` before reading any loaded values.</span></span>

```TypeScript
let range = selectedSheet.getRange("A1:B3");
range.load ("rowCount"); // Load the property.
await context.sync(); // Synchronize with the workbook to get the property.
console.log(range.rowCount); // Read and log the property value (3).
```

<span data-ttu-id="8397c-199">Sie können auch Eigenschaften aus einer ganzen Sammlung laden.</span><span class="sxs-lookup"><span data-stu-id="8397c-199">You can also load properties across an entire collection.</span></span> <span data-ttu-id="8397c-200">Jedes Sammlungsobjekt verfügt über eine `items`-Eigenschaft, bei der es sich um ein Array handelt, das die Objekte in dieser Sammlung enthält.</span><span class="sxs-lookup"><span data-stu-id="8397c-200">Every collection object has an `items` property that is an array containing the objects in that collection.</span></span> <span data-ttu-id="8397c-201">Durch die Verwendung von `items` als Anfang eines hierarchischen Aufrufs (`items\myProperty`) für `load` werden die angegebenen Eigenschaften für jedes dieser Elemente geladen.</span><span class="sxs-lookup"><span data-stu-id="8397c-201">Using `items` as the start of a hierarchical call (`items\myProperty`) to `load` loads the specified properties on each of those items.</span></span> <span data-ttu-id="8397c-202">Im folgenden Beispiel wird die `resolved`-Eigenschaft für jedes `Comment`-Objekt im `CommentCollection`-Objekt eines Arbeitsblatts geladen.</span><span class="sxs-lookup"><span data-stu-id="8397c-202">The following example loads the `resolved` property on every `Comment` object in the `CommentCollection` object of a worksheet.</span></span>

```TypeScript
let comments = selectedSheet.comments;
comments.load("items/resolved"); // Load the `resolved` property from every comment in this collection.
await context.sync(); // Synchronize with the workbook to get the properties.
```

> [!TIP]
> <span data-ttu-id="8397c-203">Wenn Sie mehr über das Arbeiten mit Sammlungen in Office-Skripts wissen möchten, lesen Sie den [Array-Abschnitt des Artikels "Verwenden von integrierten JavaScript-Objekten in Office-Skripts"](javascript-objects.md#array).</span><span class="sxs-lookup"><span data-stu-id="8397c-203">To learn more about working with collections in Office Scripts, see the [Array section of the Using built-in JavaScript objects in Office Scripts](javascript-objects.md#array) article.</span></span>

### <a name="clientresult"></a><span data-ttu-id="8397c-204">ClientResult</span><span class="sxs-lookup"><span data-stu-id="8397c-204">ClientResult</span></span>

<span data-ttu-id="8397c-205">Methoden, die Informationen aus dem Arbeitsbuch zurückgeben, haben ein ähnliches Muster wie das `load`/`sync`-Paradigma.</span><span class="sxs-lookup"><span data-stu-id="8397c-205">Methods that return information from the workbook have a similar pattern to the `load`/`sync` paradigm.</span></span> <span data-ttu-id="8397c-206">`TableCollection.getCount` ruft zum Beispiel die Anzahl von Tabellen in der Auflistung ab.</span><span class="sxs-lookup"><span data-stu-id="8397c-206">As an example, `TableCollection.getCount` gets the number of tables in the collection.</span></span> <span data-ttu-id="8397c-207">`getCount` gibt eine `ClientResult<number>` zurück, was bedeutet, dass die `value`-Eigenschaft im zurückgegebenen `ClientResult` eine Zahl ist.</span><span class="sxs-lookup"><span data-stu-id="8397c-207">`getCount` returns a `ClientResult<number>`, meaning the `value` property in the return `ClientResult` is a number.</span></span> <span data-ttu-id="8397c-208">Ihr Skript kann erst auf diesen Wert zugreifen, wenn `context.sync()` aufgerufen wird.</span><span class="sxs-lookup"><span data-stu-id="8397c-208">Your script can't access that value until `context.sync()` is called.</span></span> <span data-ttu-id="8397c-209">Ähnlich wie beim Laden einer Eigenschaft ist der `value` bis zu diesem `sync`-Aufruf ein lokaler "leerer" Wert.</span><span class="sxs-lookup"><span data-stu-id="8397c-209">Much like loading a property, the `value` is a local "empty" value until that `sync` call.</span></span>

<span data-ttu-id="8397c-210">Das folgende Skript ruft die Gesamtanzahl der Tabellen in der Arbeitsmappe ab und protokolliert diese Anzahl in der Konsole.</span><span class="sxs-lookup"><span data-stu-id="8397c-210">The following script gets the total number of tables in the workbook and logs that number to the console.</span></span>

```TypeScript
async function main(context: Excel.RequestContext) {
  let tableCount = context.workbook.tables.getCount();

  // This sync call implicitly loads tableCount.value.
  // Any other ClientResult values are loaded too.
  await context.sync();

  // Trying to log the value before calling sync would throw an error.
  console.log(tableCount.value);
}
```

## <a name="see-also"></a><span data-ttu-id="8397c-211">Siehe auch</span><span class="sxs-lookup"><span data-stu-id="8397c-211">See also</span></span>

- [<span data-ttu-id="8397c-212">Aufzeichnen, Bearbeiten und Erstellen von Office-Skripts in Excel im Web</span><span class="sxs-lookup"><span data-stu-id="8397c-212">Record, edit, and create Office Scripts in Excel on the web</span></span>](../tutorials/excel-tutorial.md)
- [<span data-ttu-id="8397c-213">Auslesen von Arbeitsmappendaten mit Office-Skripts in Excel im Web</span><span class="sxs-lookup"><span data-stu-id="8397c-213">Read workbook data with Office Scripts in Excel on the web</span></span>](../tutorials/excel-read-tutorial.md)
- [<span data-ttu-id="8397c-214">Referenzdokumentation zur Office Scripts-API</span><span class="sxs-lookup"><span data-stu-id="8397c-214">Office Scripts API reference</span></span>](/javascript/api/office-scripts/overview)
- [<span data-ttu-id="8397c-215">Verwenden von integrierten JavaScript-Objekten in Office-Skripts</span><span class="sxs-lookup"><span data-stu-id="8397c-215">Using built-in JavaScript objects in Office Scripts</span></span>](javascript-objects.md)
