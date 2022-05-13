---
title: Grundlegendes zur Skripterstellung für Office-Skripts in Excel im Web
description: Informationen zu Objektmodellen und andere Grundlagen, die Sie vor dem Schreiben von Office-Skripts benötigen.
ms.date: 05/24/2021
ms.localizationpriority: high
ms.openlocfilehash: bd51f814de60da8006413096f4d6aad125f78fab
ms.sourcegitcommit: 34c7740c9bff0e4c7426e01029f967724bfee566
ms.translationtype: HT
ms.contentlocale: de-DE
ms.lasthandoff: 05/13/2022
ms.locfileid: "65393600"
---
# <a name="scripting-fundamentals-for-office-scripts-in-excel-on-the-web"></a>Grundlegendes zur Skripterstellung für Office-Skripts in Excel im Web

In diesem Artikel werden die technischen Aspekte von Office-Skripts vorgestellt. Sie lernen die wichtigen Teile des TypeScript-basierten Skriptcodes kennen und erfahren, wie die Excel-Objekte und -APIs zusammenarbeiten.

## <a name="typescript-the-language-of-office-scripts"></a>TypeScript: Die Sprache von Office-Skripts

Office-Skripts werden in [TypeScript](https://www.typescriptlang.org/docs/home.html) geschrieben, einer Obermenge von [JavaScript-](https://developer.mozilla.org/docs/Web/JavaScript). Wenn Sie mit JavaScript vertraut sind, können Sie dieses Wissen nutzen, da ein Teil des Codes in beiden Sprachen identisch ist. Es empfiehlt sich, über Programmierkenntnisse auf Anfängerniveau zu verfügen, bevor Sie mit dem Codieren von Office Scripts beginnen. Die folgenden Ressourcen können Ihnen dabei helfen, das Coding von Office-Skripts zu verstehen.

[!INCLUDE [Recommended coding resources](../includes/coding-basics-references.md)]

## <a name="main-function-the-scripts-starting-point"></a>`main`-Funktion: Ausgangspunkt des Skripts

Jedes Skript muss eine `main`-Funktion mit dem `ExcelScript.Workbook`-Typ als ersten Parameter enthalten. Wenn die Funktion ausgeführt wird, ruft die Excel-Anwendung diese `main`-Funktion auf, indem sie die Arbeitsmappe als ersten Parameter bereitstellt. Ein `ExcelScript.Workbook` sollte immer der erste Parameter sein.

```TypeScript
function main(workbook: ExcelScript.Workbook) {
  // Your code goes here
}
```

Der Code innerhalb der `main`-Funktion wird beim Ausführen des Skripts ausgeführt. `main` kann andere Funktionen in Ihrem Skript aufrufen, Code, der nicht in einer Funktion enthalten ist, wird jedoch nicht ausgeführt. Skripts können keine anderen Office-Skripts aufrufen.

[Power Automate](https://flow.microsoft.com) ermöglicht es Ihnen, Skripts in Flüssen zu verbinden. Die Daten werden zwischen den Skripts und dem Fluss durch die Parameter und Rückgabewerte der `main`-Methode übergeben. Die Integration von Office-Skripts mit Power Automate wird im Detail unter [Ausführen von Office-Skripts mit Power Automate](power-automate-integration.md) behandelt.

## <a name="object-model-overview"></a>Übersicht über das Objektmodell

Wenn Sie ein Skript schreiben möchten, müssen Sie verstehen, wie die Office-Scripts-APIs zusammenpassen. Die Komponenten einer Arbeitsmappe haben bestimmte Beziehungen zueinander. Auf vielerlei Weise entsprechen diese Beziehungen denen der Excel-Benutzeroberfläche.

- Eine **Arbeitsmappe** enthält mindestens ein **Arbeitsblatt**.
- Ein **Arbeitsblatt** ermöglicht den Zugriff auf Zellen über **Bereichsobjekte**.
- Ein **Bereich** besteht aus einer Gruppe zusammenhängender Zellen.
- **Bereiche** werden verwendet, um **Tabellen**, **Diagramme**, **Formen** sowie andere Objekte für die Datenvisualisierung oder -organisation zu erstellen und zu platzieren.
- Ein **Arbeitsblatt** enthält Sammlungen dieser Datenobjekte, die auf dem jeweiligen Blatt vorhanden sind.
- **Arbeitsmappen** enthalten Sammlungen einiger dieser Datenobjekte (z. B. **Tabellen**) für die gesamte **Arbeitsmappe**.

## <a name="workbook"></a>Arbeitsmappe

Jedes Skript wird von der `main`-Funktion als `workbook`-Objekt vom Typ `Workbook` bereitgestellt. Damit wird das Objekt der obersten Ebene dargestellt, durch das das Skript mit der Excel-Arbeitsmappe interagiert.

Das folgende Skript ruft das aktive Arbeitsblatt aus der Arbeitsmappe ab und protokolliert den Namen.

```TypeScript
function main(workbook: ExcelScript.Workbook) {
    // Get the active worksheet.
    let sheet = workbook.getActiveWorksheet();

    // Display the current worksheet's name.
    console.log(sheet.getName());
}
```

## <a name="ranges"></a>Bereiche

Ein Bereich ist eine Gruppe zusammenhängender Zellen in der Arbeitsmappe. In Skripts wird in der Regel eine Notation im A1-Format verwendet, um Bereiche zu definieren (z. B. **B3** für die einzelne Zelle in Spalte **B** und Zeile **3**, oder **C2:F4** für die Zellen in den Spalten **C** bis **F** und den Zeilen **2** bis **4**).

Bereiche besitzen drei Haupteigenschaften: Werte, Formeln und Format. Durch diese Eigenschaften können die Zellwerte, die zu prüfenden Formeln sowie die visuelle Formatierung der Zellen abgerufen oder festgelegt werden. Sie können über `getValues`, `getFormulas` und `getFormat`auf sie zugreifen. Werte und Formeln können mit `setValues` und `setFormulas`geändert werden, wohingegen das Format ein `RangeFormat`-Objekt ist, das aus mehreren kleineren Objekten besteht, die einzeln festgelegt werden.

Bereiche verwenden zweidimensionale Arrays zum Verwalten von Informationen. Weitere Informationen zum Umgang mit Arrays im Office Scripts-Framework finden Sie unter [Arbeiten mit Bereichen](javascript-objects.md#work-with-ranges).

### <a name="range-sample"></a>Beispiel für einen Bereich

Das folgende Beispiel zeigt, wie Sie Verkaufsdatensätze erstellen können. In diesem Skript werden `Range`-Objekte zum Festlegen der Werte, Formeln und Teilen des Formats verwendet.

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
        ["Chocolate", 10, 9.54],
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

Wenn Sie dieses Skript ausführen, werden die folgenden Daten im aktuellen Arbeitsblatt erstellt:

:::image type="content" source="../images/range-sample.png" alt-text="Ein Arbeitsblatt mit einem Verkaufsdatensatz, der aus Zeilen mit Werten, einer Spalte mit Formeln und formatierten Überschriften besteht.":::

### <a name="the-types-of-range-values"></a>Typen von Bereichswerten

Jede Zelle verfügt über einen Wert. Dieser Wert ist der in die Zelle eingegebene zugrunde liegende Wert, der sich von dem in Excel angezeigten Text unterscheiden kann. Beispielsweise könnte "02.05.2021" in der Zelle als Datum angezeigt werden, aber der tatsächliche Wert ist "44318". Diese Darstellung kann über das Zahlenformat geändert werden, aber der tatsächliche Wert und Typ in der Zelle ändern sich nur, wenn ein neuer Wert festgelegt wird.

Wenn Sie den Zellwert verwenden, ist es wichtig, TypeScript mitzuteilen, welchen Wert Sie von einer Zelle oder einem Bereich erhalten möchten. Eine Zelle enthält einen der folgenden Typen: `string`, `number` oder `boolean`. Damit die zurückgegebenen Werte vom Skript als einer dieser Typen behandelt werden, müssen Sie den Typ deklarieren.

Das folgende Skript ruft den Durchschnittspreis aus der Tabelle aus dem vorherigen Beispiel ab. Beachten Sie den Code `priceRange.getValues() as number[][]`. Dieser [legt](https://www.typescriptlang.org/docs/handbook/2/everyday-types.html#type-assertions) für die Bereichswerte den Typ `number[][]` fest. Alle Werte in diesem Array können dann im Skript als Zahlen behandelt werden.

```TypeScript
function main(workbook: ExcelScript.Workbook) {
  // Get the active worksheet.
  let sheet = workbook.getActiveWorksheet();

  // Get the "Unit Price" column. 
  // The result of calling getValues is declared to be a number[][] so that we can perform arithmetic operations.
  let priceRange = sheet.getRange("D3:D5");
  let prices = priceRange.getValues() as number[][];

  // Get the average price.
  let totalPrices = 0;
  prices.forEach((price) => totalPrices += price[0]);
  let averagePrice = totalPrices / prices.length;
  console.log(averagePrice);
}
```

## <a name="charts-tables-and-other-data-objects"></a>Diagramme, Tabellen und andere Datenobjekte

Skripts können die Datenstrukturen und -visualisierungen in Excel erstellen und ändern. Tabellen und Diagramme sind zwei der am häufigsten verwendeten Objekte, die APIs unterstützen aber auch PivotTables, Formen, Bilder und vieles mehr. Diese werden in Sammlungen gespeichert, die weiter unten in diesem Artikel erläutert werden.

### <a name="create-a-table"></a>Erstellen einer Tabelle

Erstellen Sie Tabellen mithilfe von mit Daten gefüllten Bereichen. Formatierungen und Tabellensteuerelemente (z. B. Filter) werden automatisch auf den Bereich angewendet.

Durch das folgende Skript wird eine Tabelle auf Grundlage der Bereiche aus dem vorherigen Beispiel erstellt.

```TypeScript
function main(workbook: ExcelScript.Workbook) {
    // Get the active worksheet.
    let sheet = workbook.getActiveWorksheet();

    // Add a table that has headers using the data from B2:E5.
    sheet.addTable("B2:E5", true);
}
```

Wenn Sie dieses Skript auf das Arbeitsblatt mit den vorherigen Daten anwenden, wird die folgende Tabelle erstellt:

:::image type="content" source="../images/table-sample.png" alt-text="Ein Arbeitsblatt, das eine Tabelle enthält, die aus dem vorherigen Verkaufsdatensatz erstellt wurde.":::

### <a name="create-a-chart"></a>Erstellen eines Diagramms

Erstellen Sie Diagramme, um die Daten in einem Bereich darzustellen. In Skripts sind Dutzende von Diagrammvarianten zulässig, die jeweils an Ihre Anforderungen angepasst werden können.

Mit dem folgenden Skript wird ein einfaches Säulendiagramm für drei Elemente erstellt und 100 Pixel unterhalb des oberen Rands des Arbeitsblatts platziert.

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

Wenn Sie dieses Skript auf das Arbeitsblatt mit der vorherigen Tabelle anwenden, wird das folgende Diagramm erstellt:

:::image type="content" source="../images/chart-sample.png" alt-text="Ein Säulendiagramm mit den Mengen von drei Elementen aus dem vorherigen Verkaufsdatensatz.":::

## <a name="collections"></a>Sammlungen

Wenn ein Excel-Objekt eine Sammlung von mindestens einem Objekt desselben Typs enthält, werden diese in einem Array gespeichert. Beispielsweise enthält ein `Workbook`-Objekt ein `Worksheet[]`. Auf dieses Array wird über die `Workbook.getWorksheets()`-Methode zugegriffen. `get`-Methoden im Plural (z. B. `Worksheet.getCharts()`) geben die gesamte Objektsammlung als Array zurück. Dieses Muster gilt für alle Office Scripts-APIs: Das `Worksheet`-Objekt beinhaltet eine `getTables()`-Methode, die ein `Table[]` zurückgibt, das `Table`-Objekt beinhaltet eine `getColumns()`-Methode, die ein `TableColumn[]` zurückgibt, und so weiter.

Das zurückgegebene Array ist ein normales Array, daher stehen alle normalen Arrayoperationen für Ihr Skript zur Verfügung. Sie können auch auf einzelne Objekte innerhalb der Sammlung zugreifen, indem Sie den Arrayindexwert verwenden. `workbook.getTables()[0]` gibt beispielsweise die erste Tabelle in der Sammlung zurück. Weitere Informationen zur Verwendung der integrierten Arrayfunktionen mit dem Office Scripts-Framework finden Sie unter [Arbeiten mit Sammlungen](javascript-objects.md#work-with-collections).

Der Zugriff auf einzelne Objekte über die Sammlung erfolgt auch über eine `get`-Methode. `get`-Methoden im Singular (z. B. `Worksheet.getTable(name)`) geben ein einzelnes Objekt zurück und benötigen eine ID oder einen Namen für das jeweilige Objekt. Diese ID oder dieser Name wird normalerweise durch das Skript oder über die Excel-Benutzeroberfläche festgelegt.

Das folgende Skript ruft alle Tabellen in der Arbeitsmappe ab. Dadurch wird sichergestellt, dass die Kopfzeilen angezeigt werden, die Filterschaltflächen sichtbar sind und das Tabellenformat auf „TableStyleLight1“ festgelegt ist.

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

## <a name="add-excel-objects-with-a-script"></a>Hinzufügen von Excel-Objekten mit einem Skript

Sie können Dokumentobjekte, z. B. Tabellen oder Diagramme, programmgesteuert hinzufügen, indem Sie die entsprechende `add`-Methode aufrufen, die für das übergeordnete Objekt verfügbar ist.

> [!IMPORTANT]
> Fügen Sie keine Objekte manuell zu Sammlungsarrays hinzu. Verwenden Sie die `add`-Methoden in den übergeordneten Objekten. Fügen Sie z. B. `Table` mit der `Worksheet.addTable`-Methode zu `Worksheet` hinzu.

Mit dem folgenden Skript wird eine Tabelle in Excel auf dem ersten Arbeitsblatt in der Arbeitsmappe erstellt. Beachten Sie, dass die erstellte Tabelle von der `addTable`-Methode zurückgegeben wird.

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
> Die meisten Excel-Objekte besitzen eine `setName`-Methode. Dies bietet Ihnen eine einfache Möglichkeit, später im Skript oder in anderen Skripts für dieselbe Arbeitsmappe auf Excel-Objekte zuzugreifen.

### <a name="verify-an-object-exists-in-the-collection"></a>Überprüfen, ob ein Objekt in der Sammlung vorhanden ist

Skripts müssen häufig überprüfen, ob eine Tabelle oder ein ähnliches Objekt vorhanden ist, bevor sie fortfahren. Verwenden Sie die Namen, die in Skripts oder über die Excel-Benutzeroberfläche angegeben sind, um erforderliche Objekte zu identifizieren und entsprechend zu handeln. `get`-Methoden geben `undefined` zurück, wenn sich das angeforderte Objekt nicht in der Sammlung befindet.

Das folgende Skript fordert eine Tabelle namens "MyTable" an und verwendet eine `if...else`-Anweisung, um zu überprüfen, ob die Tabelle gefunden wurde.

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

Ein gängiges Muster in Office-Skripts besteht in der Neuerstellung einer Tabelle, eines Diagramms oder eines anderen Objekts bei jeder Ausführung des Skripts. Wenn Sie die alten Daten nicht benötigen, ist es am besten, das alte Objekt zu löschen, bevor Sie das neue erstellen. Dadurch werden Namenskonflikte oder andere Abweichungen vermieden, die evtl. durch andere Benutzern eingeführt wurden.

Das folgende Skript entfernt die Tabelle mit dem Namen "MyTable", wenn sie vorhanden ist, und fügt dann eine neue Tabelle mit demselben Namen hinzu.

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

## <a name="remove-excel-objects-with-a-script"></a>Entfernen von Excel-Objekten mit einem Skript

Wenn Sie ein Objekt löschen möchten, rufen Sie die `delete`-Methode des Objekts auf.

> [!NOTE]
> Wie beim Hinzufügen von Objekten dürfen Sie keine Objekte manuell aus Sammlungsarrays entfernen. Verwenden Sie die `delete`-Methoden in den Sammlungstypobjekten. Entfernen Sie beispielsweise `Table` mit `Table.delete` aus `Worksheet`.

Mit dem folgenden Skript wird das erste Arbeitsblatt in der Arbeitsmappe entfernt.

```TypeScript
function main(workbook: ExcelScript.Workbook) {
    // Get first worksheet.
    let sheet = workbook.getWorksheets()[0];

    // Remove that worksheet from the workbook.
    sheet.delete();
}
```

## <a name="further-reading-on-the-object-model"></a>Weitere Informationen zum Objektmodell

Die [Referenzdokumentation zur Office Scripts-API](/javascript/api/office-scripts/overview) besteht aus einer umfassender Liste der Objekte, die in Office-Skripts verwendet werden. Dort können Sie über das Inhaltsverzeichnis zu jedem Thema navigieren, über das Sie mehr erfahren möchten. Nachstehend finden Sie einige häufig besuchte Seiten.

- [Chart](/javascript/api/office-scripts/excelscript/excelscript.chart)
- [Kommentar](/javascript/api/office-scripts/excelscript/excelscript.comment)
- [PivotTable](/javascript/api/office-scripts/excelscript/excelscript.pivottable)
- [Range](/javascript/api/office-scripts/excelscript/excelscript.range)
- [RangeFormat](/javascript/api/office-scripts/excelscript/excelscript.rangeformat)
- [Form](/javascript/api/office-scripts/excelscript/excelscript.shape)
- [Table](/javascript/api/office-scripts/excelscript/excelscript.table)
- [Workbook](/javascript/api/office-scripts/excelscript/excelscript.workbook)
- [Arbeitsblatt](/javascript/api/office-scripts/excelscript/excelscript.worksheet)

Spezifische Informationen zum PivotTable-Objektmodell finden Sie unter [Arbeiten mit PivotTables in Office-Skripts](pivottables.md).

## <a name="see-also"></a>Siehe auch

- [Aufzeichnen, Bearbeiten und Erstellen von Office-Skripts in Excel im Web](../tutorials/excel-tutorial.md)
- [Auslesen von Arbeitsmappendaten mit Office-Skripts in Excel im Web](../tutorials/excel-read-tutorial.md)
- [Referenzdokumentation zur Office Scripts-API](/javascript/api/office-scripts/overview)
- [Arbeiten mit PivotTables in Office-Skripts](pivottables.md)
- [Verwenden von integrierten JavaScript-Objekten in Office-Skripts](javascript-objects.md)
- [Bewährte Methoden in Office-Skripts](best-practices.md)
- [Office-Skripts Dev Center](https://developer.microsoft.com/office-scripts)
