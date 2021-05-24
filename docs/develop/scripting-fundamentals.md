---
title: Grundlegendes zur Skripterstellung für Office-Skripts in Excel im Web
description: Informationen zu Objektmodellen und andere Grundlagen, die Sie vor dem Schreiben von Office-Skripts benötigen.
ms.date: 05/10/2021
localization_priority: Priority
ms.openlocfilehash: d930c9ee36933cb0458de8cce4f1d1adc7b6a001
ms.sourcegitcommit: 4687693f02fc90a57ba30c461f35046e02e6f5fb
ms.translationtype: HT
ms.contentlocale: de-DE
ms.lasthandoff: 05/19/2021
ms.locfileid: "52545100"
---
# <a name="scripting-fundamentals-for-office-scripts-in-excel-on-the-web-preview"></a><span data-ttu-id="8fccf-103">Grundlegendes zur Skripterstellung für Office-Skripts in Excel im Web (Vorschau)</span><span class="sxs-lookup"><span data-stu-id="8fccf-103">Scripting fundamentals for Office Scripts in Excel on the web (preview)</span></span>

<span data-ttu-id="8fccf-104">In diesem Artikel werden die technischen Aspekte von Office-Skripts vorgestellt.</span><span class="sxs-lookup"><span data-stu-id="8fccf-104">This article will introduce you to the technical aspects of Office Scripts.</span></span> <span data-ttu-id="8fccf-105">Sie erfahren, wie die einzelnen Excel-Objekte zusammenarbeiten und wie der Code-Editor mit einer Arbeitsmappe synchronisiert wird.</span><span class="sxs-lookup"><span data-stu-id="8fccf-105">You'll learn how the Excel objects work together and how the Code Editor synchronizes with a workbook.</span></span>

[!INCLUDE [Preview note](../includes/preview-note.md)]

## <a name="typescript-the-language-of-office-scripts"></a><span data-ttu-id="8fccf-106">TypeScript: Die Sprache von Office-Skripts</span><span class="sxs-lookup"><span data-stu-id="8fccf-106">TypeScript: The language of Office Scripts</span></span>

<span data-ttu-id="8fccf-107">Office-Skripts werden in [TypeScript](https://www.typescriptlang.org/docs/home.html) geschrieben, einer Obermenge von [JavaScript-](https://developer.mozilla.org/docs/Web/JavaScript).</span><span class="sxs-lookup"><span data-stu-id="8fccf-107">Office Scripts are written in [TypeScript](https://www.typescriptlang.org/docs/home.html), which is a superset of [JavaScript](https://developer.mozilla.org/docs/Web/JavaScript).</span></span> <span data-ttu-id="8fccf-108">Wenn Sie mit JavaScript vertraut sind, können Sie dieses Wissen nutzen, da ein Teil des Codes in beiden Sprachen identisch ist.</span><span class="sxs-lookup"><span data-stu-id="8fccf-108">If you're familiar with JavaScript, your knowledge will carry over because much of the code is the same in both languages.</span></span> <span data-ttu-id="8fccf-109">Es empfiehlt sich, über Programmierkenntnisse auf Anfängerniveau zu verfügen, bevor Sie mit dem Codieren von Office Scripts beginnen.</span><span class="sxs-lookup"><span data-stu-id="8fccf-109">We recommend you have some beginner-level programming knowledge before starting your Office Scripts coding journey.</span></span> <span data-ttu-id="8fccf-110">Die folgenden Ressourcen können Ihnen dabei helfen, das Coding von Office-Skripts zu verstehen.</span><span class="sxs-lookup"><span data-stu-id="8fccf-110">The following resources can help you understand the coding side of Office Scripts.</span></span>

[!INCLUDE [Preview note](../includes/coding-basics-references.md)]

## <a name="main-function-the-scripts-starting-point"></a><span data-ttu-id="8fccf-111">`main`-Funktion: Ausgangspunkt des Skripts</span><span class="sxs-lookup"><span data-stu-id="8fccf-111">`main` function: The script's starting point</span></span>

<span data-ttu-id="8fccf-112">Jedes Skript muss eine `main`-Funktion mit dem `ExcelScript.Workbook`-Typ als ersten Parameter enthalten.</span><span class="sxs-lookup"><span data-stu-id="8fccf-112">Each script must contain a `main` function with the `ExcelScript.Workbook` type as its first parameter.</span></span> <span data-ttu-id="8fccf-113">Wenn die Funktion ausgeführt wird, ruft die Excel-Anwendung diese `main`-Funktion auf, indem sie die Arbeitsmappe als ersten Parameter bereitstellt.</span><span class="sxs-lookup"><span data-stu-id="8fccf-113">When the function runs, the Excel application invokes the `main` function by providing the workbook as its first parameter.</span></span> <span data-ttu-id="8fccf-114">Ein `ExcelScript.Workbook` sollte immer der erste Parameter sein.</span><span class="sxs-lookup"><span data-stu-id="8fccf-114">An `ExcelScript.Workbook` should always be the first parameter.</span></span>

```TypeScript
function main(workbook: ExcelScript.Workbook) {
  // Your code goes here
}
```

<span data-ttu-id="8fccf-115">Der Code innerhalb der `main`-Funktion wird beim Ausführen des Skripts ausgeführt.</span><span class="sxs-lookup"><span data-stu-id="8fccf-115">The code inside the `main` function runs when the script is run.</span></span> <span data-ttu-id="8fccf-116">`main` kann andere Funktionen in Ihrem Skript aufrufen, Code, der nicht in einer Funktion enthalten ist, wird jedoch nicht ausgeführt.</span><span class="sxs-lookup"><span data-stu-id="8fccf-116">`main` can call other functions in your script, but code that's not contained in a function will not run.</span></span> <span data-ttu-id="8fccf-117">Skripts können keine anderen Office-Skripts aufrufen.</span><span class="sxs-lookup"><span data-stu-id="8fccf-117">Scripts cannot invoke or call other Office Scripts.</span></span>

<span data-ttu-id="8fccf-118">[Power Automate](https://flow.microsoft.com) ermöglicht es Ihnen, Skripts in Flüssen zu verbinden.</span><span class="sxs-lookup"><span data-stu-id="8fccf-118">[Power Automate](https://flow.microsoft.com) allows you to connect scripts in flows.</span></span> <span data-ttu-id="8fccf-119">Die Daten werden zwischen den Skripts und dem Fluss durch die Parameter und Rückgabewerte der `main`-Methode übergeben.</span><span class="sxs-lookup"><span data-stu-id="8fccf-119">Data is passed between the scripts and the flow through the parameters and returns of the`main` method.</span></span> <span data-ttu-id="8fccf-120">Die Integration von Office-Skripts mit Power Automate wird im Detail unter [Ausführen von Office-Skripts mit Power Automate](power-automate-integration.md) behandelt.</span><span class="sxs-lookup"><span data-stu-id="8fccf-120">How to integrate Office Scripts with Power Automate is covered in detail in [Run Office Scripts with Power Automate](power-automate-integration.md).</span></span>

## <a name="object-model-overview"></a><span data-ttu-id="8fccf-121">Übersicht über das Objektmodell</span><span class="sxs-lookup"><span data-stu-id="8fccf-121">Object model overview</span></span>

<span data-ttu-id="8fccf-122">Wenn Sie ein Skript schreiben möchten, müssen Sie verstehen, wie die Office-Scripts-APIs zusammenpassen.</span><span class="sxs-lookup"><span data-stu-id="8fccf-122">To write a script, you need to understand how the Office Scripts APIs fit together.</span></span> <span data-ttu-id="8fccf-123">Die Komponenten einer Arbeitsmappe haben bestimmte Beziehungen zueinander.</span><span class="sxs-lookup"><span data-stu-id="8fccf-123">The components of a workbook have specific relations to one another.</span></span> <span data-ttu-id="8fccf-124">Auf vielerlei Weise entsprechen diese Beziehungen denen der Excel-Benutzeroberfläche.</span><span class="sxs-lookup"><span data-stu-id="8fccf-124">In many ways, these relations match those of the Excel UI.</span></span>

- <span data-ttu-id="8fccf-125">Eine **Arbeitsmappe** enthält mindestens ein **Arbeitsblatt**.</span><span class="sxs-lookup"><span data-stu-id="8fccf-125">A **Workbook** contains one or more **Worksheets**.</span></span>
- <span data-ttu-id="8fccf-126">Ein **Arbeitsblatt** ermöglicht den Zugriff auf Zellen über **Bereichsobjekte**.</span><span class="sxs-lookup"><span data-stu-id="8fccf-126">A **Worksheet** gives access to cells through **Range** objects.</span></span>
- <span data-ttu-id="8fccf-127">Ein **Bereich** besteht aus einer Gruppe zusammenhängender Zellen.</span><span class="sxs-lookup"><span data-stu-id="8fccf-127">A **Range** represents a group of contiguous cells.</span></span>
- <span data-ttu-id="8fccf-128">**Bereiche** werden verwendet, um **Tabellen**, **Diagramme**, **Formen** sowie andere Objekte für die Datenvisualisierung oder -organisation zu erstellen und zu platzieren.</span><span class="sxs-lookup"><span data-stu-id="8fccf-128">**Ranges** are used to create and place **Tables**, **Charts**, **Shapes**, and other data visualization or organization objects.</span></span>
- <span data-ttu-id="8fccf-129">Ein **Arbeitsblatt** enthält Sammlungen dieser Datenobjekte, die auf dem jeweiligen Blatt vorhanden sind.</span><span class="sxs-lookup"><span data-stu-id="8fccf-129">A **Worksheet** contains collections of those data objects that are present in the individual sheet.</span></span>
- <span data-ttu-id="8fccf-130">**Arbeitsmappen** enthalten Sammlungen einiger dieser Datenobjekte (z. B. **Tabellen**) für die gesamte **Arbeitsmappe**.</span><span class="sxs-lookup"><span data-stu-id="8fccf-130">**Workbooks** contain collections of some of those data objects (such as **Tables**) for the entire **Workbook**.</span></span>

## <a name="workbook"></a><span data-ttu-id="8fccf-131">Arbeitsmappe</span><span class="sxs-lookup"><span data-stu-id="8fccf-131">Workbook</span></span>

<span data-ttu-id="8fccf-132">Jedes Skript wird von der `main`-Funktion als `workbook`-Objekt vom Typ `Workbook` bereitgestellt.</span><span class="sxs-lookup"><span data-stu-id="8fccf-132">Every script is provided a `workbook` object of type `Workbook` by the `main` function.</span></span> <span data-ttu-id="8fccf-133">Damit wird das Objekt der obersten Ebene dargestellt, durch das das Skript mit der Excel-Arbeitsmappe interagiert.</span><span class="sxs-lookup"><span data-stu-id="8fccf-133">This represents the top level object through which your script interacts with the Excel workbook.</span></span>

<span data-ttu-id="8fccf-134">Das folgende Skript ruft das aktive Arbeitsblatt aus der Arbeitsmappe ab und protokolliert den Namen.</span><span class="sxs-lookup"><span data-stu-id="8fccf-134">The following script gets the active worksheet from the workbook and logs its name.</span></span>

```TypeScript
function main(workbook: ExcelScript.Workbook) {
    // Get the active worksheet.
    let sheet = workbook.getActiveWorksheet();

    // Display the current worksheet's name.
    console.log(sheet.getName());
}
```

## <a name="ranges"></a><span data-ttu-id="8fccf-135">Bereiche</span><span class="sxs-lookup"><span data-stu-id="8fccf-135">Ranges</span></span>

<span data-ttu-id="8fccf-136">Ein Bereich ist eine Gruppe zusammenhängender Zellen in der Arbeitsmappe.</span><span class="sxs-lookup"><span data-stu-id="8fccf-136">A range is a group of contiguous cells in the workbook.</span></span> <span data-ttu-id="8fccf-137">In Skripts wird in der Regel eine Notation im A1-Format verwendet (z. B. **B3** für die einzelne Zelle in Spalte **B** und Zeile **3** oder **C2:F4** für die Zellen in den Spalten **C** bis **F** und den Zeilen **2** bis **4**), um Bereiche zu definieren.</span><span class="sxs-lookup"><span data-stu-id="8fccf-137">Scripts typically use A1-style notation (e.g., **B3** for the single cell in column **B** and row **3** or **C2:F4** for the cells from columns **C** through **F** and rows **2** through **4**) to define ranges.</span></span>

<span data-ttu-id="8fccf-138">Bereiche besitzen drei Haupteigenschaften: Werte, Formeln und Format.</span><span class="sxs-lookup"><span data-stu-id="8fccf-138">Ranges have three core properties: values, formulas, and format.</span></span> <span data-ttu-id="8fccf-139">Durch diese Eigenschaften können die Zellwerte, die zu prüfenden Formeln sowie die visuelle Formatierung der Zellen abgerufen oder festgelegt werden.</span><span class="sxs-lookup"><span data-stu-id="8fccf-139">These properties get or set the cell values, formulas to be evaluated, and the visual formatting of the cells.</span></span> <span data-ttu-id="8fccf-140">Sie können über `getValues`, `getFormulas` und `getFormat`auf sie zugreifen.</span><span class="sxs-lookup"><span data-stu-id="8fccf-140">They are accessed through `getValues`, `getFormulas`, and `getFormat`.</span></span> <span data-ttu-id="8fccf-141">Werte und Formeln können mit `setValues` und `setFormulas`geändert werden, wohingegen das Format ein `RangeFormat`-Objekt ist, das aus mehreren kleineren Objekten besteht, die einzeln festgelegt werden.</span><span class="sxs-lookup"><span data-stu-id="8fccf-141">Values and formulas can be changed with `setValues` and `setFormulas`, while the format is a `RangeFormat` object comprised of several smaller objects that are individually set.</span></span>

<span data-ttu-id="8fccf-142">Bereiche verwenden zweidimensionale Arrays zum Verwalten von Informationen.</span><span class="sxs-lookup"><span data-stu-id="8fccf-142">Ranges use two-dimensional arrays to manage information.</span></span> <span data-ttu-id="8fccf-143">Weitere Informationen zum Umgang mit Arrays im Office Scripts-Framework finden Sie unter [Arbeiten mit Bereichen](javascript-objects.md#work-with-ranges).</span><span class="sxs-lookup"><span data-stu-id="8fccf-143">For more information on handling arrays in the Office Scripts framework, see [Work with ranges](javascript-objects.md#work-with-ranges).</span></span>

### <a name="range-sample"></a><span data-ttu-id="8fccf-144">Beispiel für einen Bereich</span><span class="sxs-lookup"><span data-stu-id="8fccf-144">Range sample</span></span>

<span data-ttu-id="8fccf-145">Das folgende Beispiel zeigt, wie Sie Verkaufsdatensätze erstellen können.</span><span class="sxs-lookup"><span data-stu-id="8fccf-145">The following sample shows how to create sales records.</span></span> <span data-ttu-id="8fccf-146">In diesem Skript werden `Range`-Objekte zum Festlegen der Werte, Formeln und Teilen des Formats verwendet.</span><span class="sxs-lookup"><span data-stu-id="8fccf-146">This script uses `Range` objects to set the values, formulas, and parts of the format.</span></span>

```TypeScript
function main(workbook: ExcelScript.Workbook) {
    // Get the active worksheet.
    let sheet = workbook.getActiveWorksheet();

    // Create the headers and format them to stand out.
    let headers = [["Product", "Quantity", "Unit Price", "Totals"]];
    let headerRange = sheet.getRange("B2:E2");
    headerRange.setValues(headers);
    headerRange.getFormat().getFill().setColor("#4472C4");
    headerRange.getFormat().getFont().setColor("white");

    // Create the product data rows.
    let productData = [
        ["Almonds", 6, 7.5],
        ["Coffee", 20, 34.5],
        ["Chocolate", 10, 9.56],
    ];
    let dataRange = sheet.getRange("B3:D5");
    dataRange.setValues(productData);

    // Create the formulas to total the amounts sold.
    let totalFormulas = [
        ["=C3 * D3"],
        ["=C4 * D4"],
        ["=C5 * D5"],
        ["=SUM(E3:E5)"],
    ];
    let totalRange = sheet.getRange("E3:E6");
    totalRange.setFormulas(totalFormulas);
    totalRange.getFormat().getFont().setBold(true);

    // Display the totals as US dollar amounts.
    totalRange.setNumberFormat("$0.00");
}
```

<span data-ttu-id="8fccf-147">Wenn Sie dieses Skript ausführen, werden die folgenden Daten im aktuellen Arbeitsblatt erstellt:</span><span class="sxs-lookup"><span data-stu-id="8fccf-147">Running this script creates the following data in the current worksheet:</span></span>

:::image type="content" source="../images/range-sample.png" alt-text="Ein Arbeitsblatt mit einem Verkaufsdatensatz, der aus Zeilen mit Werten, einer Spalte mit Formeln und formatierten Überschriften besteht":::

## <a name="charts-tables-and-other-data-objects"></a><span data-ttu-id="8fccf-149">Diagramme, Tabellen und andere Datenobjekte</span><span class="sxs-lookup"><span data-stu-id="8fccf-149">Charts, tables, and other data objects</span></span>

<span data-ttu-id="8fccf-150">Skripts können die Datenstrukturen und -visualisierungen in Excel erstellen und ändern.</span><span class="sxs-lookup"><span data-stu-id="8fccf-150">Scripts can create and manipulate the data structures and visualizations within Excel.</span></span> <span data-ttu-id="8fccf-151">Tabellen und Diagramme sind zwei der am häufigsten verwendeten Objekte, die APIs unterstützen aber auch PivotTables, Formen, Bilder und vieles mehr.</span><span class="sxs-lookup"><span data-stu-id="8fccf-151">Tables and charts are two of the more commonly used objects, but the APIs support PivotTables, shapes, images, and more.</span></span> <span data-ttu-id="8fccf-152">Diese werden in Sammlungen gespeichert, die weiter unten in diesem Artikel erläutert werden.</span><span class="sxs-lookup"><span data-stu-id="8fccf-152">These are stored in collections, which will be discussed later in this article.</span></span>

### <a name="create-a-table"></a><span data-ttu-id="8fccf-153">Erstellen einer Tabelle</span><span class="sxs-lookup"><span data-stu-id="8fccf-153">Create a table</span></span>

<span data-ttu-id="8fccf-p113">Erstellen Sie Tabellen mithilfe von mit Daten gefüllten Bereichen. Formatierungen und Tabellensteuerelemente (z. B. Filter) werden automatisch auf den Bereich angewendet.</span><span class="sxs-lookup"><span data-stu-id="8fccf-p113">Create tables by using data-filled ranges. Formatting and table controls (such as filters) are automatically applied to the range.</span></span>

<span data-ttu-id="8fccf-156">Durch das folgende Skript wird eine Tabelle auf Grundlage der Bereiche aus dem vorherigen Beispiel erstellt.</span><span class="sxs-lookup"><span data-stu-id="8fccf-156">The following script creates a table using the ranges from the previous sample.</span></span>

```TypeScript
function main(workbook: ExcelScript.Workbook) {
    // Get the active worksheet.
    let sheet = workbook.getActiveWorksheet();

    // Add a table that has headers using the data from B2:E5.
    sheet.addTable("B2:E5", true);
}
```

<span data-ttu-id="8fccf-157">Wenn Sie dieses Skript auf das Arbeitsblatt mit den vorherigen Daten anwenden, wird die folgende Tabelle erstellt:</span><span class="sxs-lookup"><span data-stu-id="8fccf-157">Running this script on the worksheet with the previous data creates the following table:</span></span>

:::image type="content" source="../images/table-sample.png" alt-text="Ein Arbeitsblatt, das eine Tabelle enthält, die aus dem vorherigen Verkaufsdatensatz erstellt wurde.":::

### <a name="create-a-chart"></a><span data-ttu-id="8fccf-159">Erstellen eines Diagramms</span><span class="sxs-lookup"><span data-stu-id="8fccf-159">Create a chart</span></span>

<span data-ttu-id="8fccf-160">Erstellen Sie Diagramme, um die Daten in einem Bereich darzustellen.</span><span class="sxs-lookup"><span data-stu-id="8fccf-160">Create charts to visualize the data in a range.</span></span> <span data-ttu-id="8fccf-161">In Skripts sind Dutzende von Diagrammvarianten zulässig, die jeweils an Ihre Anforderungen angepasst werden können.</span><span class="sxs-lookup"><span data-stu-id="8fccf-161">Scripts allow for dozens of chart varieties, each of which can be customized to suit your needs.</span></span>

<span data-ttu-id="8fccf-162">Mit dem folgenden Skript wird ein einfaches Säulendiagramm für drei Elemente erstellt und 100 Pixel unterhalb des oberen Rands des Arbeitsblatts platziert.</span><span class="sxs-lookup"><span data-stu-id="8fccf-162">The following script creates a simple column chart for three items and places it 100 pixels below the top of the worksheet.</span></span>

```TypeScript
function main(workbook: ExcelScript.Workbook) {
    // Get the active worksheet.
    let sheet = workbook.getActiveWorksheet();

    // Create a column chart using the data from B3:C5.
    let chart = sheet.addChart(
        ExcelScript.ChartType.columnStacked,
        sheet.getRange("B3:C5")
    );

    // Set the margin of the chart to be 100 pixels from the top of the screen.
    chart.setTop(100);
}
```

<span data-ttu-id="8fccf-163">Wenn Sie dieses Skript auf das Arbeitsblatt mit der vorherigen Tabelle anwenden, wird das folgende Diagramm erstellt:</span><span class="sxs-lookup"><span data-stu-id="8fccf-163">Running this script on the worksheet with the previous table creates the following chart:</span></span>

:::image type="content" source="../images/chart-sample.png" alt-text="Ein Säulendiagramm mit den Mengen von drei Elementen aus dem vorherigen Verkaufsdatensatz":::

## <a name="collections"></a><span data-ttu-id="8fccf-165">Sammlungen</span><span class="sxs-lookup"><span data-stu-id="8fccf-165">Collections</span></span>

<span data-ttu-id="8fccf-166">Wenn ein Excel-Objekt eine Sammlung von mindestens einem Objekt desselben Typs enthält, werden diese in einem Array gespeichert.</span><span class="sxs-lookup"><span data-stu-id="8fccf-166">When an Excel object has a collection of one or more objects of the same type, it stores them in an array.</span></span> <span data-ttu-id="8fccf-167">Beispielsweise enthält ein `Workbook`-Objekt ein `Worksheet[]`.</span><span class="sxs-lookup"><span data-stu-id="8fccf-167">For example, a `Workbook` object contains a `Worksheet[]`.</span></span> <span data-ttu-id="8fccf-168">Auf dieses Array wird über die `Workbook.getWorksheets()`-Methode zugegriffen.</span><span class="sxs-lookup"><span data-stu-id="8fccf-168">This array is accessed by the `Workbook.getWorksheets()` method.</span></span> <span data-ttu-id="8fccf-169">`get`-Methoden im Plural (z. B. `Worksheet.getCharts()`) geben die gesamte Objektsammlung als Array zurück.</span><span class="sxs-lookup"><span data-stu-id="8fccf-169">`get` methods that are plural, such as `Worksheet.getCharts()`, return the entire object collection as an array.</span></span> <span data-ttu-id="8fccf-170">Dieses Muster gilt für alle Office Scripts-APIs: Das `Worksheet`-Objekt beinhaltet eine `getTables()`-Methode, die ein `Table[]` zurückgibt, das `Table`-Objekt beinhaltet eine `getColumns()`-Methode, die ein `TableColumn[]` zurückgibt, und so weiter.</span><span class="sxs-lookup"><span data-stu-id="8fccf-170">You'll see this pattern throughout the Office Scripts APIs: the `Worksheet` object has a `getTables()` method that returns a `Table[]`, the `Table` object has a `getColumns()` method that returns a `TableColumn[]`, as so on.</span></span>

<span data-ttu-id="8fccf-171">Das zurückgegebene Array ist ein normales Array, daher stehen alle normalen Arrayoperationen für Ihr Skript zur Verfügung.</span><span class="sxs-lookup"><span data-stu-id="8fccf-171">The returned array is a normal array, so all the regular array operations are available for your script.</span></span> <span data-ttu-id="8fccf-172">Sie können auch auf einzelne Objekte innerhalb der Sammlung zugreifen, indem Sie den Arrayindexwert verwenden.</span><span class="sxs-lookup"><span data-stu-id="8fccf-172">You can also access individual objects within the collection using the array index value.</span></span> <span data-ttu-id="8fccf-173">`workbook.getTables()[0]` gibt beispielsweise die erste Tabelle in der Sammlung zurück.</span><span class="sxs-lookup"><span data-stu-id="8fccf-173">For example, `workbook.getTables()[0]` returns the first table in the collection.</span></span> <span data-ttu-id="8fccf-174">Weitere Informationen zur Verwendung der integrierten Arrayfunktionen mit dem Office Scripts-Framework finden Sie unter [Arbeiten mit Sammlungen](javascript-objects.md#work-with-collections).</span><span class="sxs-lookup"><span data-stu-id="8fccf-174">For more information on using the built-in array functionality with the Office Scripts framework, see [Work with collections](javascript-objects.md#work-with-collections).</span></span> 

<span data-ttu-id="8fccf-175">Der Zugriff auf einzelne Objekte über die Sammlung erfolgt auch über eine `get`-Methode.</span><span class="sxs-lookup"><span data-stu-id="8fccf-175">Individual objects are also accessed from the collection through a `get` method.</span></span> <span data-ttu-id="8fccf-176">`get`-Methoden im Singular (z. B. `Worksheet.getTable(name)`) geben ein einzelnes Objekt zurück und benötigen eine ID oder einen Namen für das jeweilige Objekt.</span><span class="sxs-lookup"><span data-stu-id="8fccf-176">`get` methods that are singular, such as `Worksheet.getTable(name)`, return a single object and require an ID or name for the specific object.</span></span> <span data-ttu-id="8fccf-177">Diese ID oder dieser Name wird normalerweise durch das Skript oder über die Excel-Benutzeroberfläche festgelegt.</span><span class="sxs-lookup"><span data-stu-id="8fccf-177">This ID or name is usually set by the script or through the Excel UI.</span></span>

<span data-ttu-id="8fccf-p118">Das folgende Skript ruft alle Tabellen in der Arbeitsmappe ab. Dadurch wird sichergestellt, dass die Kopfzeilen angezeigt werden, die Filterschaltflächen sichtbar sind und das Tabellenformat auf „TableStyleLight1“ festgelegt ist.</span><span class="sxs-lookup"><span data-stu-id="8fccf-p118">The following script gets all tables in the workbook. It then ensures the headers are displays, the filter buttons are visible, and the table style is set to "TableStyleLight1".</span></span>

```TypeScript
function main(workbook: ExcelScript.Workbook) {
  // Get the table collection.
  let tables = workbook.getTables();

  // Set the table formatting properties for every table.
  tables.forEach(table => {
    table.setShowHeaders(true);
    table.setShowFilterButton(true);
    table.setPredefinedTableStyle("TableStyleLight1");
  })
}
```

## <a name="add-excel-objects-with-a-script"></a><span data-ttu-id="8fccf-180">Hinzufügen von Excel-Objekten mit einem Skript</span><span class="sxs-lookup"><span data-stu-id="8fccf-180">Add Excel objects with a script</span></span>

<span data-ttu-id="8fccf-181">Sie können Dokumentobjekte, z. B. Tabellen oder Diagramme, programmgesteuert hinzufügen, indem Sie die entsprechende `add`-Methode aufrufen, die für das übergeordnete Objekt verfügbar ist.</span><span class="sxs-lookup"><span data-stu-id="8fccf-181">You can programmatically add document objects, such as tables or charts, by calling the corresponding `add` method available on the parent object.</span></span>

> [!IMPORTANT]
> <span data-ttu-id="8fccf-182">Fügen Sie keine Objekte manuell zu Sammlungsarrays hinzu.</span><span class="sxs-lookup"><span data-stu-id="8fccf-182">Do not manually add objects to collection arrays.</span></span> <span data-ttu-id="8fccf-183">Verwenden Sie die `add`-Methoden in den übergeordneten Objekten. Fügen Sie z. B. `Table` mit der `Worksheet.addTable`-Methode zu `Worksheet` hinzu.</span><span class="sxs-lookup"><span data-stu-id="8fccf-183">Use the `add` methods on the parent objects For example, add a `Table` to a `Worksheet` with the `Worksheet.addTable` method.</span></span>

<span data-ttu-id="8fccf-184">Mit dem folgenden Skript wird eine Tabelle in Excel auf dem ersten Arbeitsblatt in der Arbeitsmappe erstellt.</span><span class="sxs-lookup"><span data-stu-id="8fccf-184">The following script creates a table in Excel on the first worksheet in the workbook.</span></span> <span data-ttu-id="8fccf-185">Beachten Sie, dass die erstellte Tabelle von der `addTable`-Methode zurückgegeben wird.</span><span class="sxs-lookup"><span data-stu-id="8fccf-185">Note that the created table is returned by the `addTable` method.</span></span>

```TypeScript
function main(workbook: ExcelScript.Workbook) {
    // Get the first worksheet.
    let sheet = workbook.getWorksheets()[0];

    // Add a table that uses the data in A1:G10.
    let table = sheet.addTable(
      "A1:G10",
       true /* True because the table has headers. */
    );
    
    // Give the table a name for easy reference in other scripts.
    table.setName("MyTable");
}
```

> [!TIP]
> <span data-ttu-id="8fccf-186">Die meisten Excel-Objekte besitzen eine `setName`-Methode.</span><span class="sxs-lookup"><span data-stu-id="8fccf-186">Most Excel objects have a `setName` method.</span></span> <span data-ttu-id="8fccf-187">Dies bietet Ihnen eine einfache Möglichkeit, später im Skript oder in anderen Skripts für dieselbe Arbeitsmappe auf Excel-Objekte zuzugreifen.</span><span class="sxs-lookup"><span data-stu-id="8fccf-187">This gives you an easy way to access Excel objects later in the script or in other scripts for the same workbook.</span></span>

### <a name="verify-an-object-exists-in-the-collection"></a><span data-ttu-id="8fccf-188">Überprüfen, ob ein Objekt in der Sammlung vorhanden ist</span><span class="sxs-lookup"><span data-stu-id="8fccf-188">Verify an object exists in the collection</span></span>

<span data-ttu-id="8fccf-189">Skripts müssen häufig überprüfen, ob eine Tabelle oder ein ähnliches Objekt vorhanden ist, bevor sie fortfahren.</span><span class="sxs-lookup"><span data-stu-id="8fccf-189">Scripts often need to check if a table or similar object exists before continuing.</span></span> <span data-ttu-id="8fccf-190">Verwenden Sie die Namen, die in Skripts oder über die Excel-Benutzeroberfläche angegeben sind, um erforderliche Objekte zu identifizieren und entsprechend zu handeln.</span><span class="sxs-lookup"><span data-stu-id="8fccf-190">Use the names given by scripts or through the Excel UI to identify necessary objects and act accordingly.</span></span> <span data-ttu-id="8fccf-191">`get`-Methoden geben `undefined` zurück, wenn sich das angeforderte Objekt nicht in der Sammlung befindet.</span><span class="sxs-lookup"><span data-stu-id="8fccf-191">`get` methods return `undefined` when the requested object is not in the collection.</span></span>

<span data-ttu-id="8fccf-192">Das folgende Skript fordert eine Tabelle namens "MyTable" an und verwendet eine `if...else`-Anweisung, um zu überprüfen, ob die Tabelle gefunden wurde.</span><span class="sxs-lookup"><span data-stu-id="8fccf-192">The following script requests a table named "MyTable" and uses an `if...else` statement to check if the table was found.</span></span>

```TypeScript
function main(workbook: ExcelScript.Workbook) {
  // Get the table named "MyTable".
  let myTable = workbook.getTable("MyTable");

  // If the table is in the workbook, myTable will have a value.
  // Otherwise, the variable will be undefined and go to the else clause.
  if (myTable) {
    let worksheetName = myTable.getWorksheet().getName();
    console.log(`MyTable is on the ${worksheetName} worksheet`);
  } else {
    console.log(`MyTable is not in the workbook.`);
  }
}
```

<span data-ttu-id="8fccf-193">Ein gängiges Muster in Office-Skripts besteht in der Neuerstellung einer Tabelle, eines Diagramms oder eines anderen Objekts bei jeder Ausführung des Skripts.</span><span class="sxs-lookup"><span data-stu-id="8fccf-193">A common pattern in Office Scripts is to recreate a table, chart, or other object every time the script is run.</span></span> <span data-ttu-id="8fccf-194">Wenn Sie die alten Daten nicht benötigen, ist es am besten, das alte Objekt zu löschen, bevor Sie das neue erstellen.</span><span class="sxs-lookup"><span data-stu-id="8fccf-194">If you don't need the old data, it's best to delete the old object before creating the new one.</span></span> <span data-ttu-id="8fccf-195">Dadurch werden Namenskonflikte oder andere Abweichungen vermieden, die evtl. durch andere Benutzern eingeführt wurden.</span><span class="sxs-lookup"><span data-stu-id="8fccf-195">This avoids name conflicts or other differences that may have been introduced by other users.</span></span>

<span data-ttu-id="8fccf-196">Das folgende Skript entfernt die Tabelle mit dem Namen "MyTable", wenn sie vorhanden ist, und fügt dann eine neue Tabelle mit demselben Namen hinzu.</span><span class="sxs-lookup"><span data-stu-id="8fccf-196">The following script removes the table named "MyTable", if it is present, then adds a new table with the same name.</span></span>

```TypeScript
function main(workbook: ExcelScript.Workbook) {
  // Get the table named "MyTable" from the first worksheet.
  let sheet = workbook.getWorksheets()[0];
  let tableName = "MyTable";
  let oldTable = sheet.getTable(tableName);

  // If the table exists, remove it.
  if (oldTable) {
    oldTable.delete();
  }

  // Add a new table with the same name.
  let newTable = sheet.addTable("A1:G10", true);
  newTable.setName(tableName);
}
```

## <a name="remove-excel-objects-with-a-script"></a><span data-ttu-id="8fccf-197">Entfernen von Excel-Objekten mit einem Skript</span><span class="sxs-lookup"><span data-stu-id="8fccf-197">Remove Excel objects with a script</span></span>

<span data-ttu-id="8fccf-198">Wenn Sie ein Objekt löschen möchten, rufen Sie die `delete`-Methode des Objekts auf.</span><span class="sxs-lookup"><span data-stu-id="8fccf-198">To delete an object, call the object's `delete` method.</span></span>

> [!NOTE]
> <span data-ttu-id="8fccf-199">Wie beim Hinzufügen von Objekten dürfen Sie keine Objekte manuell aus Sammlungsarrays entfernen.</span><span class="sxs-lookup"><span data-stu-id="8fccf-199">As with adding objects, do not manually remove objects from collection arrays.</span></span> <span data-ttu-id="8fccf-200">Verwenden Sie die `delete`-Methoden in den Sammlungstypobjekten.</span><span class="sxs-lookup"><span data-stu-id="8fccf-200">Use the `delete` methods on the collection-type objects.</span></span> <span data-ttu-id="8fccf-201">Entfernen Sie beispielsweise `Table` mit `Table.delete` aus `Worksheet`.</span><span class="sxs-lookup"><span data-stu-id="8fccf-201">For example, remove a `Table` from a `Worksheet` using `Table.delete`.</span></span>

<span data-ttu-id="8fccf-202">Mit dem folgenden Skript wird das erste Arbeitsblatt in der Arbeitsmappe entfernt.</span><span class="sxs-lookup"><span data-stu-id="8fccf-202">The following script removes the first worksheet in the workbook.</span></span>

```TypeScript
function main(workbook: ExcelScript.Workbook) {
    // Get first worksheet.
    let sheet = workbook.getWorksheets()[0];

    // Remove that worksheet from the workbook.
    sheet.delete();
}
```

## <a name="further-reading-on-the-object-model"></a><span data-ttu-id="8fccf-203">Weitere Informationen zum Objektmodell</span><span class="sxs-lookup"><span data-stu-id="8fccf-203">Further reading on the object model</span></span>

<span data-ttu-id="8fccf-204">Die [Referenzdokumentation zur Office Scripts-API](/javascript/api/office-scripts/overview) besteht aus einer umfassender Liste der Objekte, die in Office-Skripts verwendet werden.</span><span class="sxs-lookup"><span data-stu-id="8fccf-204">The [Office Scripts API reference documentation](/javascript/api/office-scripts/overview) is a comprehensive listing of the objects used in Office Scripts.</span></span> <span data-ttu-id="8fccf-205">Dort können Sie über das Inhaltsverzeichnis zu jedem Thema navigieren, über das Sie mehr erfahren möchten.</span><span class="sxs-lookup"><span data-stu-id="8fccf-205">There, you can use the table of contents to navigate to any class you'd like to learn more about.</span></span> <span data-ttu-id="8fccf-206">Nachstehend finden Sie einige häufig besuchte Seiten.</span><span class="sxs-lookup"><span data-stu-id="8fccf-206">The following are several commonly viewed pages.</span></span>

- [<span data-ttu-id="8fccf-207">Chart</span><span class="sxs-lookup"><span data-stu-id="8fccf-207">Chart</span></span>](/javascript/api/office-scripts/excelscript/excelscript.chart)
- [<span data-ttu-id="8fccf-208">Kommentar</span><span class="sxs-lookup"><span data-stu-id="8fccf-208">Comment</span></span>](/javascript/api/office-scripts/excelscript/excelscript.comment)
- [<span data-ttu-id="8fccf-209">PivotTable</span><span class="sxs-lookup"><span data-stu-id="8fccf-209">PivotTable</span></span>](/javascript/api/office-scripts/excelscript/excelscript.pivottable)
- [<span data-ttu-id="8fccf-210">Range</span><span class="sxs-lookup"><span data-stu-id="8fccf-210">Range</span></span>](/javascript/api/office-scripts/excelscript/excelscript.range)
- [<span data-ttu-id="8fccf-211">RangeFormat</span><span class="sxs-lookup"><span data-stu-id="8fccf-211">RangeFormat</span></span>](/javascript/api/office-scripts/excelscript/excelscript.rangeformat)
- [<span data-ttu-id="8fccf-212">Form</span><span class="sxs-lookup"><span data-stu-id="8fccf-212">Shape</span></span>](/javascript/api/office-scripts/excelscript/excelscript.shape)
- [<span data-ttu-id="8fccf-213">Table</span><span class="sxs-lookup"><span data-stu-id="8fccf-213">Table</span></span>](/javascript/api/office-scripts/excelscript/excelscript.table)
- [<span data-ttu-id="8fccf-214">Workbook</span><span class="sxs-lookup"><span data-stu-id="8fccf-214">Workbook</span></span>](/javascript/api/office-scripts/excelscript/excelscript.workbook)
- [<span data-ttu-id="8fccf-215">Arbeitsblatt</span><span class="sxs-lookup"><span data-stu-id="8fccf-215">Worksheet</span></span>](/javascript/api/office-scripts/excelscript/excelscript.worksheet)

## <a name="see-also"></a><span data-ttu-id="8fccf-216">Siehe auch</span><span class="sxs-lookup"><span data-stu-id="8fccf-216">See also</span></span>

- [<span data-ttu-id="8fccf-217">Aufzeichnen, Bearbeiten und Erstellen von Office-Skripts in Excel im Web</span><span class="sxs-lookup"><span data-stu-id="8fccf-217">Record, edit, and create Office Scripts in Excel on the web</span></span>](../tutorials/excel-tutorial.md)
- [<span data-ttu-id="8fccf-218">Auslesen von Arbeitsmappendaten mit Office-Skripts in Excel im Web</span><span class="sxs-lookup"><span data-stu-id="8fccf-218">Read workbook data with Office Scripts in Excel on the web</span></span>](../tutorials/excel-read-tutorial.md)
- [<span data-ttu-id="8fccf-219">Referenzdokumentation zur Office Scripts-API</span><span class="sxs-lookup"><span data-stu-id="8fccf-219">Office Scripts API reference</span></span>](/javascript/api/office-scripts/overview)
- [<span data-ttu-id="8fccf-220">Verwenden von integrierten JavaScript-Objekten in Office-Skripts</span><span class="sxs-lookup"><span data-stu-id="8fccf-220">Using built-in JavaScript objects in Office Scripts</span></span>](javascript-objects.md)
- [<span data-ttu-id="8fccf-221">Bewährte Methoden in Office-Skripts</span><span class="sxs-lookup"><span data-stu-id="8fccf-221">Best practices in Office Scripts</span></span>](best-practices.md)
