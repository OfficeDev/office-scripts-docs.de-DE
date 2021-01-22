---
title: Grundlegendes zur Skripterstellung für Office-Skripts in Excel im Web
description: Informationen zu Objektmodellen und andere Grundlagen, die Sie vor dem Schreiben von Office-Skripts benötigen.
ms.date: 07/08/2020
localization_priority: Priority
ms.openlocfilehash: acbeec69a5d9ae9e3ebfa95c9070033d1cca2265
ms.sourcegitcommit: e7e019ba36c2f49451ec08c71a1679eb6dba4268
ms.translationtype: HT
ms.contentlocale: de-DE
ms.lasthandoff: 01/22/2021
ms.locfileid: "49933273"
---
# <a name="scripting-fundamentals-for-office-scripts-in-excel-on-the-web-preview"></a>Grundlegendes zur Skripterstellung für Office-Skripts in Excel im Web (Vorschau)

In diesem Artikel werden die technischen Aspekte von Office-Skripts vorgestellt. Sie erfahren, wie die einzelnen Excel-Objekte zusammenarbeiten und wie der Code-Editor mit einer Arbeitsmappe synchronisiert wird.

[!INCLUDE [Preview note](../includes/preview-note.md)]

## <a name="main-function"></a>Die `main`-Funktion

Jedes Office-Skript muss eine `main`-Funktion mit dem `ExcelScript.Workbook`-Typ als ersten Parameter enthalten. Wenn die Funktion ausgeführt wird, ruft die Excel-Anwendung diese `main`-Funktion auf, indem sie die Arbeitsmappe als ersten Parameter bereitstellt. Deshalb ist es wichtig, dass Sie die Standardsignatur der `main`-Funktion nicht ändern, nachdem Sie das Skript aufgezeichnet oder im Code-Editor ein neues Skript erstellt haben.

```typescript
function main(workbook: ExcelScript.Workbook) {
  // Your code goes here
}
```

Der Code innerhalb der `main`-Funktion wird beim Ausführen des Skripts ausgeführt. `main` kann andere Funktionen in Ihrem Skript aufrufen, Code, der nicht in einer Funktion enthalten ist, wird jedoch nicht ausgeführt.

> [!CAUTION]
> Wenn die `main`-Funktion wie `async function main(context: Excel.RequestContext)` aussieht, verwendet das Skript das ältere asynchrone API-Modell. Weitere Informationen (auch Informationen zum Konvertieren Ihres Skripts in das aktuelle API-Modell) finden Sie unter [Unterstützung für ältere Office-Skripts, die die Async-APIs verwenden](excel-async-model.md).

## <a name="object-model"></a>Objektmodell

Wenn Sie ein Skript schreiben möchten, müssen Sie verstehen, wie die Office-Skript-APIs zusammenpassen. Die Komponenten einer Arbeitsmappe haben bestimmte Beziehungen zueinander. Auf vielerlei Weise entsprechen diese Beziehungen denen der Excel-Benutzeroberfläche.

- Eine **Arbeitsmappe** enthält mindestens ein **Arbeitsblatt**.
- Ein **Arbeitsblatt** ermöglicht den Zugriff auf Zellen über **Bereichsobjekte**.
- Ein **Bereich** besteht aus einer Gruppe zusammenhängender Zellen.
- **Bereiche** werden verwendet, um **Tabellen**, **Diagramme**, **Formen** sowie andere Objekte für die Datenvisualisierung oder -organisation zu erstellen und zu platzieren.
- Ein **Arbeitsblatt** enthält Sammlungen dieser Datenobjekte, die auf dem jeweiligen Blatt vorhanden sind.
- **Arbeitsmappen** enthalten Sammlungen einiger dieser Datenobjekte (z. B. **Tabellen**) für die gesamte **Arbeitsmappe**.

### <a name="workbook"></a>Arbeitsmappe

Jedes Skript wird von der `main`-Funktion als `workbook`-Objekt vom Typ `Workbook` bereitgestellt. Damit wird das Objekt der obersten Ebene dargestellt, durch das das Skript mit der Excel-Arbeitsmappe interagiert.

Das folgende Skript ruft das aktive Arbeitsblatt aus der Arbeitsmappe ab und protokolliert den Namen.

```typescript
function main(workbook: ExcelScript.Workbook) {
    // Get the active worksheet.
    let sheet = workbook.getActiveWorksheet();

    // Display the current worksheet's name.
    console.log(sheet.getName());
}
```

### <a name="ranges"></a>Bereiche

Ein Bereich ist eine Gruppe zusammenhängender Zellen in der Arbeitsmappe. In Skripts wird in der Regel eine Notation im A1-Format verwendet (z. B. **B3** für die einzelne Zelle in Spalte **B** und Zeile **3** oder **C2:F4** für die Zellen in den Spalten **C** bis **F** und den Zeilen **2** bis **4**), um Bereiche zu definieren.

Bereiche besitzen drei Haupteigenschaften: Werte, Formeln und Format. Durch diese Eigenschaften können die Zellwerte, die zu prüfenden Formeln sowie die visuelle Formatierung der Zellen abgerufen oder festgelegt werden. Sie können über `getValues`, `getFormulas` und `getFormat`auf sie zugreifen. Werte und Formeln können mit `setValues` und `setFormulas`geändert werden, wohingegen das Format ein `RangeFormat`-Objekt ist, das aus mehreren kleineren Objekten besteht, die einzeln festgelegt werden.

Bereiche verwenden zweidimensionale Arrays zum Verwalten von Informationen. Lesen Sie den Abschnitt [„Arbeiten mit Bereichen“ des Artikels „Verwenden von integrierten JavaScript-Objekten in Office-Skripts“](javascript-objects.md#working-with-ranges), um weitere Informationen zum Umgang mit diesen Arrays im Office-Skripts-Framework zu erhalten.

#### <a name="range-sample"></a>Beispiel für einen Bereich

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

Wenn Sie dieses Skript ausführen, werden die folgenden Daten im aktuellen Arbeitsblatt erstellt:

![Ein Umsatzdatensatz mit Wert-Zeilen, einer Formelspalte sowie formatierten Überschriften.](../images/range-sample.png)

### <a name="charts-tables-and-other-data-objects"></a>Diagramme, Tabellen und andere Datenobjekte

Skripts können die Datenstrukturen und -visualisierungen in Excel erstellen und ändern. Tabellen und Diagramme sind zwei der am häufigsten verwendeten Objekte, die APIs unterstützen aber auch PivotTables, Formen, Bilder und vieles mehr. Diese werden in Sammlungen gespeichert, die weiter unten in diesem Artikel erläutert werden.

#### <a name="creating-a-table"></a>Erstellen einer Tabelle

Erstellen Sie Tabellen mithilfe von mit Daten ausgefüllten Bereichen. Auf den Bereich werden automatisch Formatierungs- und Tabellen-Steuerelemente (wie z. B. Filter) angewendet.

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

![Eine Tabelle aus dem vorherigen Umsatzeintrag.](../images/table-sample.png)

#### <a name="creating-a-chart"></a>Erstellen eines Diagramms

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

![Ein Säulendiagramm, in dem die Mengenangaben zu drei Elementen aus dem vorherigen Umsatzeintrag angezeigt werden.](../images/chart-sample.png)

### <a name="collections-and-other-object-relations"></a>Sammlungen und andere Objektbeziehungen

Auf jedes untergeordnete Objekt kann über das übergeordnete Objekt zugegriffen werden. Sie können z. B. `Worksheets` aus dem `Workbook`-Objekt lesen. Für die übergeordnete Klasse gibt eine zugehörige `get`-Methode vorhanden sein (z. B. `Workbook.getWorksheets()` oder `Workbook.getWorksheet(name)`). `get`-Methoden im Singular geben ein einzelnes Objekt zurück und benötigen eine ID oder einen Namen für das jeweilige Objekt (z. B. den Namen eines Arbeitsblatts). `get`-Methoden im Plural geben die gesamte Objektsammlung als Array zurück. Wenn die Sammlung leer ist, erhalten Sie ein leeres Array (`[]`).

Sobald die Sammlung abgerufen wurde, können Sie reguläre Arrayoperationen wie das Abrufen der `length` oder die Verwendung von `for`, `for..of`, `while` Schleifen für Iterationen oder die Verwendung von TypeScript-Arraymethoden wie `map` oder `forEach` verwenden. Sie können auch auf einzelne Objekte innerhalb der Sammlung zugreifen, indem Sie den Arrayindexwert verwenden. `workbook.getTables()[0]` gibt beispielsweise die erste Tabelle in der Sammlung zurück. Lesen Sie den Abschnitt [„Arbeiten mit Sammlungen“ des Artikels „Verwenden von integrierten JavaScript-Objekten in Office-Skripts“](javascript-objects.md#working-with-collections), um weitere Informationen zur Verwendung der integrierten Arrayfunktionen mit dem Office-Skripts-Framework zu erhalten.

Das folgende Skript ruft alle Tabellen in der Arbeitsmappe ab. Dann wird sichergestellt, dass die Kopfzeilen angezeigt werden, die Filterschaltflächen sichtbar sind und das Tabellenformat auf „TableStyleLight1“ festgelegt ist.

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

#### <a name="adding-excel-objects-with-a-script"></a>Hinzufügen von Excel-Objekten mit einem Skript

Sie können Dokumentobjekte, z. B. Tabellen oder Diagramme, programmgesteuert hinzufügen, indem Sie die entsprechende `add`-Methode aufrufen, die für das übergeordnete Objekt verfügbar ist.

> [!NOTE]
> Fügen Sie keine Objekte manuell zu Sammlungsarrays hinzu. Verwenden Sie die `add`-Methoden in den übergeordneten Objekten. Fügen Sie z. B. `Table` mit der `Worksheet.addTable`-Methode zu `Worksheet` hinzu.

Mit dem folgenden Skript wird eine Tabelle in Excel auf dem ersten Arbeitsblatt in der Arbeitsmappe erstellt. Beachten Sie, dass die erstellte Tabelle von der `addTable`-Methode zurückgegeben wird.

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

## <a name="removing-excel-objects-with-a-script"></a>Entfernen von Excel-Objekten mit einem Skript

Wenn Sie ein Objekt löschen möchten, rufen Sie die `delete`-Methode des Objekts auf.

> [!NOTE]
> Wie beim Hinzufügen von Objekten dürfen Sie keine Objekte manuell aus Sammlungsarrays entfernen. Verwenden Sie die `delete`-Methoden in den Sammlungstypobjekten. Entfernen Sie beispielsweise `Table` mit `Table.delete` aus `Worksheet`.

Mit dem folgenden Skript wird das erste Arbeitsblatt in der Arbeitsmappe entfernt.

```typescript
function main(workbook: ExcelScript.Workbook) {
    // Get first worksheet.
    let sheet = workbook.getWorksheets()[0];

    // Remove that worksheet from the workbook.
    sheet.delete();
}
```

### <a name="further-reading-on-the-object-model"></a>Weitere Informationen zum Objektmodell

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

## <a name="see-also"></a>Siehe auch

- [Aufzeichnen, Bearbeiten und Erstellen von Office-Skripts in Excel im Web](../tutorials/excel-tutorial.md)
- [Auslesen von Arbeitsmappendaten mit Office-Skripts in Excel im Web](../tutorials/excel-read-tutorial.md)
- [Referenzdokumentation zur Office Scripts-API](/javascript/api/office-scripts/overview)
- [Verwenden von integrierten JavaScript-Objekten in Office-Skripts](javascript-objects.md)
