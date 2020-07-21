---
title: Grundlegendes zur Skripterstellung für Office-Skripts in Excel im Web
description: Informationen zu Objektmodellen und andere Grundlagen, die Sie vor dem Schreiben von Office-Skripts benötigen.
ms.date: 07/08/2020
localization_priority: Priority
ms.openlocfilehash: 6c02f4fb986e6a0ed1dd7afb099aaa1c9d1ea276
ms.sourcegitcommit: ebd1079c7e2695ac0e7e4c616f2439975e196875
ms.translationtype: HT
ms.contentlocale: de-DE
ms.lasthandoff: 07/17/2020
ms.locfileid: "45160474"
---
# <a name="scripting-fundamentals-for-office-scripts-in-excel-on-the-web-preview"></a><span data-ttu-id="e0cb3-103">Grundlegendes zur Skripterstellung für Office-Skripts in Excel im Web (Vorschau)</span><span class="sxs-lookup"><span data-stu-id="e0cb3-103">Scripting fundamentals for Office Scripts in Excel on the web (preview)</span></span>

<span data-ttu-id="e0cb3-104">In diesem Artikel werden die technischen Aspekte von Office-Skripts vorgestellt.</span><span class="sxs-lookup"><span data-stu-id="e0cb3-104">This article will introduce you to the technical aspects of Office Scripts.</span></span> <span data-ttu-id="e0cb3-105">Sie erfahren, wie die einzelnen Excel-Objekte zusammenarbeiten und wie der Code-Editor mit einer Arbeitsmappe synchronisiert wird.</span><span class="sxs-lookup"><span data-stu-id="e0cb3-105">You'll learn how the Excel objects work together and how the Code Editor synchronizes with a workbook.</span></span>

[!INCLUDE [Preview note](../includes/preview-note.md)]

## <a name="main-function"></a><span data-ttu-id="e0cb3-106">Die `main`-Funktion</span><span class="sxs-lookup"><span data-stu-id="e0cb3-106">`main` function</span></span>

<span data-ttu-id="e0cb3-107">Jedes Office-Skript muss eine `main`-Funktion mit dem `ExcelScript.Workbook`-Typ als ersten Parameter enthalten.</span><span class="sxs-lookup"><span data-stu-id="e0cb3-107">Each Office Script must contain a `main` function with the `ExcelScript.Workbook` type as its first parameter.</span></span> <span data-ttu-id="e0cb3-108">Wenn die Funktion ausgeführt wird, ruft die Excel-Anwendung diese `main`-Funktion auf, indem sie die Arbeitsmappe als ersten Parameter bereitstellt.</span><span class="sxs-lookup"><span data-stu-id="e0cb3-108">When the function is executed, Excel application invokes this `main` function by providing the workbook as its first parameter.</span></span> <span data-ttu-id="e0cb3-109">Deshalb ist es wichtig, dass Sie die Standardsignatur der `main`-Funktion nicht ändern, nachdem Sie das Skript aufgezeichnet oder im Code-Editor ein neues Skript erstellt haben.</span><span class="sxs-lookup"><span data-stu-id="e0cb3-109">Hence, it is important to not modify the basic signature of the `main` function once you have either recorded the script or created a new script from the code editor.</span></span>

```typescript
function main(workbook: ExcelScript.Workbook) {
  // Your code goes here
}
```

<span data-ttu-id="e0cb3-110">Der Code innerhalb der `main`-Funktion wird beim Ausführen des Skripts ausgeführt.</span><span class="sxs-lookup"><span data-stu-id="e0cb3-110">The code inside the `main` function runs when the script is run.</span></span> <span data-ttu-id="e0cb3-111">`main` kann andere Funktionen in Ihrem Skript aufrufen, Code, der nicht in einer Funktion enthalten ist, wird jedoch nicht ausgeführt.</span><span class="sxs-lookup"><span data-stu-id="e0cb3-111">`main` can call other functions in your script, but code that's not contained in a function will not run.</span></span>

> [!CAUTION]
> <span data-ttu-id="e0cb3-112">Wenn die `main`-Funktion wie `async function main(context: Excel.RequestContext)` aussieht, verwendet das Skript das ältere asynchrone API-Modell.</span><span class="sxs-lookup"><span data-stu-id="e0cb3-112">If your `main` function looks like `async function main(context: Excel.RequestContext)`, your script is using the older async API model.</span></span> <span data-ttu-id="e0cb3-113">Weitere Informationen (auch Informationen zum Konvertieren Ihres Skripts in das aktuelle API-Modell) finden Sie unter [Unterstützung für ältere Office-Skripts, die die Async-APIs verwenden](excel-async-model.md).</span><span class="sxs-lookup"><span data-stu-id="e0cb3-113">For more information (including how to convert your script to the current API model), refer to [Support older Office Scripts that use the Async APIs](excel-async-model.md).</span></span>

## <a name="object-model"></a><span data-ttu-id="e0cb3-114">Objektmodell</span><span class="sxs-lookup"><span data-stu-id="e0cb3-114">Object model</span></span>

<span data-ttu-id="e0cb3-115">Wenn Sie ein Skript schreiben möchten, müssen Sie verstehen, wie die Office-Skript-APIs zusammenpassen.</span><span class="sxs-lookup"><span data-stu-id="e0cb3-115">To write a script, you need to understand how the Office Script APIs fit together.</span></span> <span data-ttu-id="e0cb3-116">Die Komponenten einer Arbeitsmappe haben bestimmte Beziehungen zueinander.</span><span class="sxs-lookup"><span data-stu-id="e0cb3-116">The components of a workbook have specific relations to one another.</span></span> <span data-ttu-id="e0cb3-117">Auf vielerlei Weise entsprechen diese Beziehungen denen der Excel-Benutzeroberfläche.</span><span class="sxs-lookup"><span data-stu-id="e0cb3-117">In many ways, these relations match those of the Excel UI.</span></span>

- <span data-ttu-id="e0cb3-118">Eine **Arbeitsmappe** enthält mindestens ein **Arbeitsblatt**.</span><span class="sxs-lookup"><span data-stu-id="e0cb3-118">A **Workbook** contains one or more **Worksheets**.</span></span>
- <span data-ttu-id="e0cb3-119">Ein **Arbeitsblatt** ermöglicht den Zugriff auf Zellen über **Bereichsobjekte**.</span><span class="sxs-lookup"><span data-stu-id="e0cb3-119">A **Worksheet** gives access to cells through **Range** objects.</span></span>
- <span data-ttu-id="e0cb3-120">Ein **Bereich** besteht aus einer Gruppe zusammenhängender Zellen.</span><span class="sxs-lookup"><span data-stu-id="e0cb3-120">A **Range** represents a group of contiguous cells.</span></span>
- <span data-ttu-id="e0cb3-121">**Bereiche** werden verwendet, um **Tabellen**, **Diagramme**, **Formen** sowie andere Objekte für die Datenvisualisierung oder -organisation zu erstellen und zu platzieren.</span><span class="sxs-lookup"><span data-stu-id="e0cb3-121">**Ranges** are used to create and place **Tables**, **Charts**, **Shapes**, and other data visualization or organization objects.</span></span>
- <span data-ttu-id="e0cb3-122">Ein **Arbeitsblatt** enthält Sammlungen dieser Datenobjekte, die auf dem jeweiligen Blatt vorhanden sind.</span><span class="sxs-lookup"><span data-stu-id="e0cb3-122">A **Worksheet** contains collections of those data objects that are present in the individual sheet.</span></span>
- <span data-ttu-id="e0cb3-123">**Arbeitsmappen** enthalten Sammlungen einiger dieser Datenobjekte (z. B. **Tabellen**) für die gesamte **Arbeitsmappe**.</span><span class="sxs-lookup"><span data-stu-id="e0cb3-123">**Workbooks** contain collections of some of those data objects (such as **Tables**) for the entire **Workbook**.</span></span>

### <a name="workbook"></a><span data-ttu-id="e0cb3-124">Arbeitsmappe</span><span class="sxs-lookup"><span data-stu-id="e0cb3-124">Workbook</span></span>

<span data-ttu-id="e0cb3-125">Jedes Skript wird von der `main`-Funktion als `workbook`-Objekt vom Typ `Workbook` bereitgestellt.</span><span class="sxs-lookup"><span data-stu-id="e0cb3-125">Every script is provided a `workbook` object of type `Workbook` by the `main` function.</span></span> <span data-ttu-id="e0cb3-126">Damit wird das Objekt der obersten Ebene dargestellt, durch das das Skript mit der Excel-Arbeitsmappe interagiert.</span><span class="sxs-lookup"><span data-stu-id="e0cb3-126">This represents the top level object through which your script interacts with the Excel workbook.</span></span>

<span data-ttu-id="e0cb3-127">Das folgende Skript ruft das aktive Arbeitsblatt aus der Arbeitsmappe ab und protokolliert den Namen.</span><span class="sxs-lookup"><span data-stu-id="e0cb3-127">The following script gets the active worksheet from the workbook and logs its name.</span></span>

```typescript
function main(workbook: ExcelScript.Workbook) {
    // Get the active worksheet.
    let sheet = workbook.getActiveWorksheet();

    // Display the current worksheet's name.
    console.log(sheet.getName());
}
```

### <a name="ranges"></a><span data-ttu-id="e0cb3-128">Bereiche</span><span class="sxs-lookup"><span data-stu-id="e0cb3-128">Ranges</span></span>

<span data-ttu-id="e0cb3-129">Ein Bereich ist eine Gruppe zusammenhängender Zellen in der Arbeitsmappe.</span><span class="sxs-lookup"><span data-stu-id="e0cb3-129">A range is a group of contiguous cells in the workbook.</span></span> <span data-ttu-id="e0cb3-130">In Skripts wird in der Regel eine Notation im A1-Format verwendet (z. B. **B3** für die einzelne Zelle in Spalte **B** und Zeile **3** oder **C2:F4** für die Zellen in den Spalten **C** bis **F** und den Zeilen **2** bis **4**), um Bereiche zu definieren.</span><span class="sxs-lookup"><span data-stu-id="e0cb3-130">Scripts typically use A1-style notation (e.g., **B3** for the single cell in column **B** and row **3** or **C2:F4** for the cells from columns **C** through **F** and rows **2** through **4**) to define ranges.</span></span>

<span data-ttu-id="e0cb3-131">Bereiche besitzen drei Haupteigenschaften: Werte, Formeln und Format.</span><span class="sxs-lookup"><span data-stu-id="e0cb3-131">Ranges have three core properties: values, formulas, and format.</span></span> <span data-ttu-id="e0cb3-132">Durch diese Eigenschaften können die Zellwerte, die zu prüfenden Formeln sowie die visuelle Formatierung der Zellen abgerufen oder festgelegt werden.</span><span class="sxs-lookup"><span data-stu-id="e0cb3-132">These properties get or set the cell values, formulas to be evaluated, and the visual formatting of the cells.</span></span> <span data-ttu-id="e0cb3-133">Sie können über `getValues`, `getFormulas` und `getFormat`auf sie zugreifen.</span><span class="sxs-lookup"><span data-stu-id="e0cb3-133">They are accessed through `getValues`, `getFormulas`, and `getFormat`.</span></span> <span data-ttu-id="e0cb3-134">Werte und Formeln können mit `setValues` und `setFormulas`geändert werden, wohingegen das Format ein `RangeFormat`-Objekt ist, das aus mehreren kleineren Objekten besteht, die einzeln festgelegt werden.</span><span class="sxs-lookup"><span data-stu-id="e0cb3-134">Values and formulas can be changed with `setValues` and `setFormulas`, while the format is a `RangeFormat` object comprised of several smaller objects that are individually set.</span></span>

<span data-ttu-id="e0cb3-135">Bereiche verwenden zweidimensionale Arrays zum Verwalten von Informationen.</span><span class="sxs-lookup"><span data-stu-id="e0cb3-135">Ranges use two-dimensional arrays to manage information.</span></span> <span data-ttu-id="e0cb3-136">Lesen Sie den Abschnitt [„Arbeiten mit Bereichen“ des Artikels „Verwenden von integrierten JavaScript-Objekten in Office-Skripts“](javascript-objects.md#working-with-ranges), um weitere Informationen zum Umgang mit diesen Arrays im Office-Skripts-Framework zu erhalten.</span><span class="sxs-lookup"><span data-stu-id="e0cb3-136">Read the [Working with ranges section of Using built-in JavaScript objects in Office Scripts](javascript-objects.md#working-with-ranges) for more information on handling those arrays in the Office Scripts framework.</span></span>

#### <a name="range-sample"></a><span data-ttu-id="e0cb3-137">Beispiel für einen Bereich</span><span class="sxs-lookup"><span data-stu-id="e0cb3-137">Range sample</span></span>

<span data-ttu-id="e0cb3-138">Das folgende Beispiel zeigt, wie Sie Verkaufsdatensätze erstellen können.</span><span class="sxs-lookup"><span data-stu-id="e0cb3-138">The following sample shows how to create sales records.</span></span> <span data-ttu-id="e0cb3-139">In diesem Skript werden `Range`-Objekte zum Festlegen der Werte, Formeln und Teilen des Formats verwendet.</span><span class="sxs-lookup"><span data-stu-id="e0cb3-139">This script uses `Range` objects to set the values, formulas, and parts of the format.</span></span>

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

<span data-ttu-id="e0cb3-140">Wenn Sie dieses Skript ausführen, werden die folgenden Daten im aktuellen Arbeitsblatt erstellt:</span><span class="sxs-lookup"><span data-stu-id="e0cb3-140">Running this script creates the following data in the current worksheet:</span></span>

![Ein Umsatzdatensatz mit Wert-Zeilen, einer Formelspalte sowie formatierten Überschriften.](../images/range-sample.png)

### <a name="charts-tables-and-other-data-objects"></a><span data-ttu-id="e0cb3-142">Diagramme, Tabellen und andere Datenobjekte</span><span class="sxs-lookup"><span data-stu-id="e0cb3-142">Charts, tables, and other data objects</span></span>

<span data-ttu-id="e0cb3-143">Skripts können die Datenstrukturen und -visualisierungen in Excel erstellen und ändern.</span><span class="sxs-lookup"><span data-stu-id="e0cb3-143">Scripts can create and manipulate the data structures and visualizations within Excel.</span></span> <span data-ttu-id="e0cb3-144">Tabellen und Diagramme sind zwei der am häufigsten verwendeten Objekte, die APIs unterstützen aber auch PivotTables, Formen, Bilder und vieles mehr.</span><span class="sxs-lookup"><span data-stu-id="e0cb3-144">Tables and charts are two of the more commonly used objects, but the APIs support PivotTables, shapes, images, and more.</span></span> <span data-ttu-id="e0cb3-145">Diese werden in Sammlungen gespeichert, die weiter unten in diesem Artikel erläutert werden.</span><span class="sxs-lookup"><span data-stu-id="e0cb3-145">These are stored in collections, which will be discussed later in this article.</span></span>

#### <a name="creating-a-table"></a><span data-ttu-id="e0cb3-146">Erstellen einer Tabelle</span><span class="sxs-lookup"><span data-stu-id="e0cb3-146">Creating a table</span></span>

<span data-ttu-id="e0cb3-147">Erstellen Sie Tabellen mithilfe von mit Daten ausgefüllten Bereichen.</span><span class="sxs-lookup"><span data-stu-id="e0cb3-147">Create tables by using data-filled ranges.</span></span> <span data-ttu-id="e0cb3-148">Auf den Bereich werden automatisch Formatierungs- und Tabellen-Steuerelemente (wie z. B. Filter) angewendet.</span><span class="sxs-lookup"><span data-stu-id="e0cb3-148">Formatting and table controls (such as filters) are automatically applied to the range.</span></span>

<span data-ttu-id="e0cb3-149">Durch das folgende Skript wird eine Tabelle auf Grundlage der Bereiche aus dem vorherigen Beispiel erstellt.</span><span class="sxs-lookup"><span data-stu-id="e0cb3-149">The following script creates a table using the ranges from the previous sample.</span></span>

```TypeScript
function main(workbook: ExcelScript.Workbook) {
    // Get the active worksheet.
    let sheet = workbook.getActiveWorksheet();

    // Add a table that has headers using the data from B2:E5.
    sheet.addTable("B2:E5", true);
}
```

<span data-ttu-id="e0cb3-150">Wenn Sie dieses Skript auf das Arbeitsblatt mit den vorherigen Daten anwenden, wird die folgende Tabelle erstellt:</span><span class="sxs-lookup"><span data-stu-id="e0cb3-150">Running this script on the worksheet with the previous data creates the following table:</span></span>

![Eine Tabelle aus dem vorherigen Umsatzeintrag.](../images/table-sample.png)

#### <a name="creating-a-chart"></a><span data-ttu-id="e0cb3-152">Erstellen eines Diagramms</span><span class="sxs-lookup"><span data-stu-id="e0cb3-152">Creating a chart</span></span>

<span data-ttu-id="e0cb3-153">Erstellen Sie Diagramme, um die Daten in einem Bereich darzustellen.</span><span class="sxs-lookup"><span data-stu-id="e0cb3-153">Create charts to visualize the data in a range.</span></span> <span data-ttu-id="e0cb3-154">In Skripts sind Dutzende von Diagrammvarianten zulässig, die jeweils an Ihre Anforderungen angepasst werden können.</span><span class="sxs-lookup"><span data-stu-id="e0cb3-154">Scripts allow for dozens of chart varieties, each of which can be customized to suit your needs.</span></span>

<span data-ttu-id="e0cb3-155">Mit dem folgenden Skript wird ein einfaches Säulendiagramm für drei Elemente erstellt und 100 Pixel unterhalb des oberen Rands des Arbeitsblatts platziert.</span><span class="sxs-lookup"><span data-stu-id="e0cb3-155">The following script creates a simple column chart for three items and places it 100 pixels below the top of the worksheet.</span></span>

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

<span data-ttu-id="e0cb3-156">Wenn Sie dieses Skript auf das Arbeitsblatt mit der vorherigen Tabelle anwenden, wird das folgende Diagramm erstellt:</span><span class="sxs-lookup"><span data-stu-id="e0cb3-156">Running this script on the worksheet with the previous table creates the following chart:</span></span>

![Ein Säulendiagramm, in dem die Mengenangaben zu drei Elementen aus dem vorherigen Umsatzeintrag angezeigt werden.](../images/chart-sample.png)

### <a name="collections-and-other-object-relations"></a><span data-ttu-id="e0cb3-158">Sammlungen und andere Objektbeziehungen</span><span class="sxs-lookup"><span data-stu-id="e0cb3-158">Collections and other object relations</span></span>

<span data-ttu-id="e0cb3-159">Auf jedes untergeordnete Objekt kann über das übergeordnete Objekt zugegriffen werden.</span><span class="sxs-lookup"><span data-stu-id="e0cb3-159">Any child object can be accessed through its parent object.</span></span> <span data-ttu-id="e0cb3-160">Sie können z. B. `Worksheets` aus dem `Workbook`-Objekt lesen.</span><span class="sxs-lookup"><span data-stu-id="e0cb3-160">For example, you can read `Worksheets` from the `Workbook` object.</span></span> <span data-ttu-id="e0cb3-161">Für die übergeordnete Klasse gibt eine zugehörige `get`-Methode vorhanden sein (z. B. `Workbook.getWorksheets()` oder `Workbook.getWorksheet(name)`).</span><span class="sxs-lookup"><span data-stu-id="e0cb3-161">There will be a related `get` method on the parent class that (e.g., `Workbook.getWorksheets()` or `Workbook.getWorksheet(name)`).</span></span> <span data-ttu-id="e0cb3-162">`get`-Methoden im Singular geben ein einzelnes Objekt zurück und benötigen eine ID oder einen Namen für das jeweilige Objekt (z. B. den Namen eines Arbeitsblatts).</span><span class="sxs-lookup"><span data-stu-id="e0cb3-162">`get` methods that are singular return a single object and require an ID or name for the specific object (such as the name of a worksheet).</span></span> <span data-ttu-id="e0cb3-163">`get`-Methoden im Plural geben die gesamte Objektsammlung als Array zurück.</span><span class="sxs-lookup"><span data-stu-id="e0cb3-163">`get` methods that are plural return the entire object collection as an array.</span></span> <span data-ttu-id="e0cb3-164">Wenn die Sammlung leer ist, erhalten Sie ein leeres Array (`[]`).</span><span class="sxs-lookup"><span data-stu-id="e0cb3-164">If the collection is empty, you'll get an empty array (`[]`).</span></span>

<span data-ttu-id="e0cb3-165">Sobald die Sammlung abgerufen wurde, können Sie reguläre Arrayoperationen wie das Abrufen der `length` oder die Verwendung von `for`, `for..of`, `while` Schleifen für Iterationen oder die Verwendung von TypeScript-Arraymethoden wie `map` oder `forEach` verwenden.</span><span class="sxs-lookup"><span data-stu-id="e0cb3-165">Once the collection is retrieved, you can use regular array operations such as getting its `length` or use `for`, `for..of`, `while` loops for iteration or use TypeScript array methods such as `map`, `forEach` on them.</span></span> <span data-ttu-id="e0cb3-166">Sie können auch auf einzelne Objekte innerhalb der Sammlung zugreifen, indem Sie den Arrayindexwert verwenden.</span><span class="sxs-lookup"><span data-stu-id="e0cb3-166">You can also access individual objects within the collection using the array index value.</span></span> <span data-ttu-id="e0cb3-167">`workbook.getTables()[0]` gibt beispielsweise die erste Tabelle in der Sammlung zurück.</span><span class="sxs-lookup"><span data-stu-id="e0cb3-167">For example, `workbook.getTables()[0]` returns the first table in the collection.</span></span> <span data-ttu-id="e0cb3-168">Lesen Sie den Abschnitt [„Arbeiten mit Sammlungen“ des Artikels „Verwenden von integrierten JavaScript-Objekten in Office-Skripts“](javascript-objects.md#working-with-collections), um weitere Informationen zur Verwendung der integrierten Arrayfunktionen mit dem Office-Skripts-Framework zu erhalten.</span><span class="sxs-lookup"><span data-stu-id="e0cb3-168">Read the [Working with collections section of Using built-in JavaScript objects in Office Scripts](javascript-objects.md#working-with-collections) to learn more about using built-in array functionality with the Office Scripts framework.</span></span>

<span data-ttu-id="e0cb3-169">Das folgende Skript ruft alle Tabellen in der Arbeitsmappe ab.</span><span class="sxs-lookup"><span data-stu-id="e0cb3-169">The following script gets all tables in the workbook.</span></span> <span data-ttu-id="e0cb3-170">Dann wird sichergestellt, dass die Kopfzeilen angezeigt werden, die Filterschaltflächen sichtbar sind und das Tabellenformat auf „TableStyleLight1“ festgelegt ist.</span><span class="sxs-lookup"><span data-stu-id="e0cb3-170">It then ensures the headers are displays, the filter buttons are visible, and the table style is set to "TableStyleLight1".</span></span>

```typescript
function main(workbook: ExcelScript.Workbook) {
  /* Get table collection */
  const tables = workbook.getTables();
  /* Set table formatting properties */
  tables.forEach(table => {
    table.setShowHeaders(true);
    table.setShowFilterButton(true);
    table.setPredefinedTableStyle("TableStyleLight1");
  })
}
```

#### <a name="adding-excel-objects-with-a-script"></a><span data-ttu-id="e0cb3-171">Hinzufügen von Excel-Objekten mit einem Skript</span><span class="sxs-lookup"><span data-stu-id="e0cb3-171">Adding Excel objects with a script</span></span>

<span data-ttu-id="e0cb3-172">Sie können Dokumentobjekte, z. B. Tabellen oder Diagramme, programmgesteuert hinzufügen, indem Sie die entsprechende `add`-Methode aufrufen, die für das übergeordnete Objekt verfügbar ist.</span><span class="sxs-lookup"><span data-stu-id="e0cb3-172">You can programmatically add document objects, such as tables or charts, by calling the corresponding `add` method available on the parent object.</span></span>

> [!NOTE]
> <span data-ttu-id="e0cb3-173">Fügen Sie keine Objekte manuell zu Sammlungsarrays hinzu.</span><span class="sxs-lookup"><span data-stu-id="e0cb3-173">Do not manually add objects to collection arrays.</span></span> <span data-ttu-id="e0cb3-174">Verwenden Sie die `add`-Methoden in den übergeordneten Objekten. Fügen Sie z. B. `Table` mit der `Worksheet.addTable`-Methode zu `Worksheet` hinzu.</span><span class="sxs-lookup"><span data-stu-id="e0cb3-174">Use the `add` methods on the parent objects For example, add a `Table` to a `Worksheet` with the `Worksheet.addTable` method.</span></span>

<span data-ttu-id="e0cb3-175">Mit dem folgenden Skript wird eine Tabelle in Excel auf dem ersten Arbeitsblatt in der Arbeitsmappe erstellt.</span><span class="sxs-lookup"><span data-stu-id="e0cb3-175">The following script creates a table in Excel on the first worksheet in the workbook.</span></span> <span data-ttu-id="e0cb3-176">Beachten Sie, dass die erstellte Tabelle von der `addTable`-Methode zurückgegeben wird.</span><span class="sxs-lookup"><span data-stu-id="e0cb3-176">Note that the created table is returned by the `addTable` method.</span></span>

```typescript
function main(workbook: ExcelScript.Workbook) {
    // Get the first worksheet.
    let sheet = workbook.getWorksheets()[0];

    // Add a table that uses the data in C3:G10.
    let table = sheet.addTable(
      "C3:G10",
       true /* True because the table has headers. */
    );
}
```

## <a name="removing-excel-objects-with-a-script"></a><span data-ttu-id="e0cb3-177">Entfernen von Excel-Objekten mit einem Skript</span><span class="sxs-lookup"><span data-stu-id="e0cb3-177">Removing Excel objects with a script</span></span>

<span data-ttu-id="e0cb3-178">Wenn Sie ein Objekt löschen möchten, rufen Sie die `delete`-Methode des Objekts auf.</span><span class="sxs-lookup"><span data-stu-id="e0cb3-178">To delete an object, call the object's `delete` method.</span></span>

> [!NOTE]
> <span data-ttu-id="e0cb3-179">Wie beim Hinzufügen von Objekten dürfen Sie keine Objekte manuell aus Sammlungsarrays entfernen.</span><span class="sxs-lookup"><span data-stu-id="e0cb3-179">As with adding objects, do not manually remove objects from collection arrays.</span></span> <span data-ttu-id="e0cb3-180">Verwenden Sie die `delete`-Methoden in den Sammlungstypobjekten.</span><span class="sxs-lookup"><span data-stu-id="e0cb3-180">Use the `delete` methods on the collection-type objects.</span></span> <span data-ttu-id="e0cb3-181">Entfernen Sie beispielsweise `Table` mit `Table.delete` aus `Worksheet`.</span><span class="sxs-lookup"><span data-stu-id="e0cb3-181">For example, remove a `Table` from a `Worksheet` using `Table.delete`.</span></span>

<span data-ttu-id="e0cb3-182">Mit dem folgenden Skript wird das erste Arbeitsblatt in der Arbeitsmappe entfernt.</span><span class="sxs-lookup"><span data-stu-id="e0cb3-182">The following script removes the first worksheet in the workbook.</span></span>

```typescript
function main(workbook: ExcelScript.Workbook) {
    // Get first worksheet.
    let sheet = workbook.getWorksheets()[0];

    // Remove that worksheet from the workbook.
    sheet.delete();
}
```

### <a name="further-reading-on-the-object-model"></a><span data-ttu-id="e0cb3-183">Weitere Informationen zum Objektmodell</span><span class="sxs-lookup"><span data-stu-id="e0cb3-183">Further reading on the object model</span></span>

<span data-ttu-id="e0cb3-184">Die [Referenzdokumentation zur Office Scripts-API](/javascript/api/office-scripts/overview) besteht aus einer umfassender Liste der Objekte, die in Office-Skripts verwendet werden.</span><span class="sxs-lookup"><span data-stu-id="e0cb3-184">The [Office Scripts API reference documentation](/javascript/api/office-scripts/overview) is a comprehensive listing of the objects used in Office Scripts.</span></span> <span data-ttu-id="e0cb3-185">Dort können Sie über das Inhaltsverzeichnis zu jedem Thema navigieren, über das Sie mehr erfahren möchten.</span><span class="sxs-lookup"><span data-stu-id="e0cb3-185">There, you can use the table of contents to navigate to any class you'd like to learn more about.</span></span> <span data-ttu-id="e0cb3-186">Nachstehend finden Sie einige häufig besuchte Seiten.</span><span class="sxs-lookup"><span data-stu-id="e0cb3-186">The following are several commonly viewed pages.</span></span>

- [<span data-ttu-id="e0cb3-187">Chart</span><span class="sxs-lookup"><span data-stu-id="e0cb3-187">Chart</span></span>](/javascript/api/office-scripts/excelscript/excelscript.chart)
- [<span data-ttu-id="e0cb3-188">Kommentar</span><span class="sxs-lookup"><span data-stu-id="e0cb3-188">Comment</span></span>](/javascript/api/office-scripts/excelscript/excelscript.comment)
- [<span data-ttu-id="e0cb3-189">PivotTable</span><span class="sxs-lookup"><span data-stu-id="e0cb3-189">PivotTable</span></span>](/javascript/api/office-scripts/excelscript/excelscript.pivottable)
- [<span data-ttu-id="e0cb3-190">Range</span><span class="sxs-lookup"><span data-stu-id="e0cb3-190">Range</span></span>](/javascript/api/office-scripts/excelscript/excelscript.range)
- [<span data-ttu-id="e0cb3-191">RangeFormat</span><span class="sxs-lookup"><span data-stu-id="e0cb3-191">RangeFormat</span></span>](/javascript/api/office-scripts/excelscript/excelscript.rangeformat)
- [<span data-ttu-id="e0cb3-192">Form</span><span class="sxs-lookup"><span data-stu-id="e0cb3-192">Shape</span></span>](/javascript/api/office-scripts/excelscript/excelscript.shape)
- [<span data-ttu-id="e0cb3-193">Table</span><span class="sxs-lookup"><span data-stu-id="e0cb3-193">Table</span></span>](/javascript/api/office-scripts/excelscript/excelscript.table)
- [<span data-ttu-id="e0cb3-194">Workbook</span><span class="sxs-lookup"><span data-stu-id="e0cb3-194">Workbook</span></span>](/javascript/api/office-scripts/excelscript/excelscript.workbook)
- [<span data-ttu-id="e0cb3-195">Arbeitsblatt</span><span class="sxs-lookup"><span data-stu-id="e0cb3-195">Worksheet</span></span>](/javascript/api/office-scripts/excelscript/excelscript.worksheet)

## <a name="see-also"></a><span data-ttu-id="e0cb3-196">Siehe auch</span><span class="sxs-lookup"><span data-stu-id="e0cb3-196">See also</span></span>

- [<span data-ttu-id="e0cb3-197">Aufzeichnen, Bearbeiten und Erstellen von Office-Skripts in Excel im Web</span><span class="sxs-lookup"><span data-stu-id="e0cb3-197">Record, edit, and create Office Scripts in Excel on the web</span></span>](../tutorials/excel-tutorial.md)
- [<span data-ttu-id="e0cb3-198">Auslesen von Arbeitsmappendaten mit Office-Skripts in Excel im Web</span><span class="sxs-lookup"><span data-stu-id="e0cb3-198">Read workbook data with Office Scripts in Excel on the web</span></span>](../tutorials/excel-read-tutorial.md)
- [<span data-ttu-id="e0cb3-199">Referenzdokumentation zur Office Scripts-API</span><span class="sxs-lookup"><span data-stu-id="e0cb3-199">Office Scripts API reference</span></span>](/javascript/api/office-scripts/overview)
- [<span data-ttu-id="e0cb3-200">Verwenden von integrierten JavaScript-Objekten in Office-Skripts</span><span class="sxs-lookup"><span data-stu-id="e0cb3-200">Using built-in JavaScript objects in Office Scripts</span></span>](javascript-objects.md)
