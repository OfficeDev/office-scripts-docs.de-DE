---
title: Grundlegendes zur Skripterstellung für Office-Skripts in Excel im Web
description: Informationen zu Objektmodellen und andere Grundlagen, die Sie vor dem Schreiben von Office-Skripts benötigen.
ms.date: 01/27/2020
localization_priority: Priority
ms.openlocfilehash: 5a709c16e23c00ffc7ee7949a3cb11459dc2d530
ms.sourcegitcommit: d556aaefac80e55f53ac56b7f6ecbc657ebd426f
ms.translationtype: HT
ms.contentlocale: de-DE
ms.lasthandoff: 03/26/2020
ms.locfileid: "42978719"
---
# <a name="scripting-fundamentals-for-office-scripts-in-excel-on-the-web-preview"></a>Grundlegendes zur Skripterstellung für Office-Skripts in Excel im Web (Vorschau)

In diesem Artikel werden die technischen Aspekte von Office-Skripts vorgestellt. Sie erfahren, wie die einzelnen Excel-Objekte zusammenarbeiten und wie der Code-Editor mit einer Arbeitsmappe synchronisiert wird.

[!INCLUDE [Preview note](../includes/preview-note.md)]

## <a name="object-model"></a>Objektmodell

Um die Excel-APIs zu verstehen, müssen Sie wissen, wie die Komponenten einer Arbeitsmappe miteinander verknüpft sind.

- Eine **Arbeitsmappe** enthält mindestens ein **Arbeitsblatt**.
- Ein **Arbeitsblatt** ermöglicht den Zugriff auf Zellen über **Bereichsobjekte**.
- Ein **Bereich** besteht aus einer Gruppe zusammenhängender Zellen.
- **Bereiche** werden verwendet, um **Tabellen**, **Diagramme**, **Formen** sowie andere Objekte für die Datenvisualisierung oder -organisation zu erstellen und zu platzieren.
- Ein **Arbeitsblatt** enthält Sammlungen dieser Datenobjekte, die auf dem jeweiligen Blatt vorhanden sind.
- **Arbeitsmappen** enthalten Sammlungen einiger dieser Datenobjekte (z. B. **Tabellen**) für die gesamte **Arbeitsmappe**.

### <a name="ranges"></a>Bereiche

Ein Bereich ist eine Gruppe zusammenhängender Zellen in der Arbeitsmappe. In Skripts wird in der Regel eine Notation im A1-Format verwendet (z. B. **B3** für die einzelne Zelle in Zeile **B**und Spalte **3**, oder **C2:F4** für die Zellen in den Zeilen **C** bis **F** und den Spalten **2** bis **4**), um Bereiche zu definieren.

Bereiche besitzen drei Haupteigenschaften: `values`, `formulas` und `format`. Durch diese Eigenschaften können die Zellwerte, die zu prüfenden Formeln sowie die visuelle Formatierung der Zellen abgerufen oder festgelegt werden.

#### <a name="range-sample"></a>Beispiel für einen Bereich

Das folgende Beispiel zeigt, wie Sie Verkaufsdatensätze erstellen können. In diesem Skript werden `Range`-Objekte zum Festlegen der Werte, Formeln und Formate verwendet.

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

Wenn Sie dieses Skript ausführen, werden die folgenden Daten im aktuellen Arbeitsblatt erstellt:

![Ein Umsatzdatensatz mit Wert-Zeilen, einer Formelspalte sowie formatierten Überschriften.](../images/range-sample.png)

### <a name="charts-tables-and-other-data-objects"></a>Diagramme, Tabellen und andere Datenobjekte

Skripts können die Datenstrukturen und -visualisierungen in Excel erstellen und ändern. Tabellen und Diagramme sind zwei der am häufigsten verwendeten Objekte, die APIs unterstützen aber auch PivotTables, Formen, Bilder und vieles mehr.

#### <a name="creating-a-table"></a>Erstellen einer Tabelle

Erstellen Sie Tabellen mithilfe von mit Daten ausgefüllten Bereichen. Auf den Bereich werden automatisch Formatierungs- und Tabellen-Steuerelemente (wie z. B. Filter) angewendet.

Durch das folgende Skript wird eine Tabelle auf Grundlage der Bereiche aus dem vorherigen Beispiel erstellt.

```TypeScript
async function main(context: Excel.RequestContext) {
   let sheet = context.workbook.worksheets.getActiveWorksheet();
   sheet.tables.add("B2:E5", true);
}
```

Wenn Sie dieses Skript auf das Arbeitsblatt mit den vorherigen Daten anwenden, wird die folgende Tabelle erstellt:

![Eine Tabelle aus dem vorherigen Umsatzeintrag.](../images/table-sample.png)

#### <a name="creating-a-chart"></a>Erstellen eines Diagramms

Erstellen Sie Diagramme, um die Daten in einem Bereich darzustellen. In Skripts sind Dutzende von Diagrammvarianten zulässig, die jeweils an Ihre Anforderungen angepasst werden können.

Mit dem folgenden Skript wird ein einfaches Säulendiagramm für drei Elemente erstellt und 100 Pixel unterhalb des oberen Rands des Arbeitsblatts platziert.

```TypeScript
async function main(context: Excel.RequestContext) {
  let sheet = context.workbook.worksheets.getActiveWorksheet();
  let chart = sheet.charts.add(Excel.ChartType.columnStacked, sheet.getRange("B3:C5"));
  chart.top = 100;
}
```

Wenn Sie dieses Skript auf das Arbeitsblatt mit der vorherigen Tabelle anwenden, wird das folgende Diagramm erstellt:

![Ein Säulendiagramm, in dem die Mengenangaben zu drei Elementen aus dem vorherigen Umsatzeintrag angezeigt werden.](../images/chart-sample.png)

### <a name="further-reading-on-the-object-model"></a>Weitere Informationen zum Objektmodell

Die [Referenzdokumentation zur Office Scripts-API](/javascript/api/office-scripts/overview) besteht aus einer umfassender Liste der Objekte, die in Office-Skripts verwendet werden. Dort können Sie über das Inhaltsverzeichnis zu jedem Thema navigieren, über das Sie mehr erfahren möchten. Nachstehend finden Sie einige häufig besuchte Seiten.

- [Chart](/javascript/api/office-scripts/excel/excel.chart)
- [Kommentar](/javascript/api/office-scripts/excel/excel.comment)
- [PivotTable](/javascript/api/office-scripts/excel/excel.pivottable)
- [Range](/javascript/api/office-scripts/excel/excel.range)
- [RangeFormat](/javascript/api/office-scripts/excel/excel.rangeformat)
- [Form](/javascript/api/office-scripts/excel/excel.shape)
- [Table](/javascript/api/office-scripts/excel/excel.table)
- [Workbook](/javascript/api/office-scripts/excel/excel.workbook)
- [Worksheet](/javascript/api/office-scripts/excel/excel.worksheet)

## <a name="main-function"></a>Die `main`-Funktion

Jedes Office-Skript muss eine `main`-Funktion mit der folgenden Signatur enthalten, einschließlich der `Excel.RequestContext`-Typdefinition:

```TypeScript
async function main(context: Excel.RequestContext) {
    // Your Excel Script
}
```

Der Code innerhalb der `main`-Funktion wird beim Ausführen des Skripts ausgeführt. `main` kann andere Funktionen in Ihrem Skript aufrufen, Code, der nicht in einer Funktion enthalten ist, wird jedoch nicht ausgeführt.

## <a name="context"></a>Context

Die `main`-Funktion lässt einen `Excel.RequestContext`-Parameter namens `context` zu. Stellen Sie sich `context` als die Brücke zwischen Ihrem Skript und der Arbeitsmappe vor. Das Skript greift auf die Arbeitsmappe mit dem `context`-Objekt zu und verwendet diesen `context` zum Hin- und Hersenden von Daten.

Das `context`-Objekt ist erforderlich, weil das Skript und Excel in unterschiedlichen Prozessen und Speicherorten ausgeführt werden. Das Skript muss Änderungen an den Daten in der Arbeitsmappe in der Cloud vornehmen oder diese abrufen können. Das `context`-Objekt verwaltet diese Transaktionen.

## <a name="sync-and-load"></a>Synchronisieren und Laden

Da Ihr Skript und die Arbeitsmappe an unterschiedlichen Orten ausgeführt werden, dauert die Datenübertragung zwischen diesen etwas. Um die Skriptleistung zu verbessern, werden Befehle in die Warteschlange gesetzt, bis das Skript explizit den `sync`-Vorgang aufruft, um das Skript und die Arbeitsmappe miteinander zu synchronisieren. Ihr Skript kann unabhängig funktionieren, bis es eine der folgenden Aktionen durchführen muss:

- Daten aus der Arbeitsmappe auslesen (nach einem `load`-Vorgang)
- Daten in die Arbeitsmappe schreiben (in der Regel, weil das Skript abgeschlossen wurde).

In der folgenden Abbildung wird ein Beispiel für eine Ablaufsteuerung zwischen dem Skript und der Arbeitsmappe dargestellt:

![Ein Diagramm mit Lese- und Schreibvorgängen, die vom Skript in der Arbeitsmappe ausgeführt werden.](../images/load-sync.png)

### <a name="sync"></a>Synchronisierung

Wenn das Skript Daten aus der Arbeitsmappe auslesen oder in diese schreiben muss, rufen Sie die `RequestContext.sync`-Methode wie hier dargestellt auf:

```TypeScript
await context.sync();
```

> [!NOTE]
> `context.sync()` wird implizit aufgerufen, wenn ein Skript endet.

Nachdem der `sync`-Vorgang abgeschlossen ist, wird die Arbeitsmappe entsprechend den Schreibvorgängen aktualisiert, die vom Skript angegeben wurden. Bei einem Schreibvorgang wird eine beliebige Eigenschaft eines Excel-Objekts festgelegt (z. B. `range.format.fill.color = "red"`) oder eine Methode aufgerufen, über die eine Eigenschaft geändert wird (z. B. `range.format.autoFitColumns()`). Der `sync`-Vorgang liest außerdem über einen `load`-Vorgang (wie im nächsten Abschnitt beschrieben) alle Werte aus der Arbeitsmappe aus, die das Skript angefordert hat.

Je nach Netzwerk kann es einige Zeit dauern, bis das Skript mit der Arbeitsmappe synchronisiert wurde. Sie sollten die Anzahl von `sync`-Aufrufen minimieren, damit das Ausführen des Skripts möglichst schnell geht.  

### <a name="load"></a>Laden

Ein Skript muss Daten aus der Arbeitsmappe laden, bevor es sie liest. Das häufige Laden von Daten aus der gesamten Arbeitsmappe würde jedoch die Geschwindigkeit des Skripts erheblich verringern. Stattdessen können Sie mit der `load`-Methode präzise angeben, welche Daten aus der Arbeitsmappe abgerufen werden sollen.

Die `load`-Methode ist für jedes Excel-Objekt verfügbar. Ihr Skript muss die Eigenschaften eines Objekts laden, bevor es sie lesen kann. Andernfalls wird ein Fehler zurückgegeben.

In den folgenden Beispielen wird ein `Range`-Objekt verwendet, um die drei Arten darzustellen, wie die `load`-Methode zum Laden von Daten verwendet werden kann.

|Absicht |Beispielbefehl | Auswirkung |
|:--|:--|:--|
|Laden einer Eigenschaft |`myRange.load("values");` | Lädt eine einzelne Eigenschaft, in diesem Fall den zweidimensionalen Wertearray in diesem Bereich. |
|Laden mehrerer Eigenschaften |`myRange.load("values, rowCount, columnCount");`| Lädt alle Eigenschaften aus einer durch Kommas getrennten Liste, in diesem Beispiel die Werte, die Zeilenanzahl und die Spaltenanzahl. |
|Alles laden | `myRange.load();`|Lädt alle Eigenschaften des Zellbereichs. Dies ist keine empfohlene Lösung, da das Skript durch das Abrufen unnötiger Daten verlangsamt wird. Sie sollten diesen Wert nur verwenden, wenn Sie das Skript testen, oder wenn Sie alle Eigenschaften des Objekts benötigen. |

Ihr Skript muss `context.sync()` aufrufen, bevor es geladene Werte ausliest.

```TypeScript
let range = selectedSheet.getRange("A1:B3");
range.load ("rowCount"); // Load the property.
await context.sync(); // Synchronize with the workbook to get the property.
console.log(range.rowCount); // Read and log the property value (3).
```

Sie können auch Eigenschaften aus einer ganzen Sammlung laden. Jedes Sammlungsobjekt verfügt über eine `items`-Eigenschaft, bei der es sich um ein Array handelt, das die Objekte in dieser Sammlung enthält. Durch die Verwendung von `items` als Anfang eines hierarchischen Aufrufs (`items\myProperty`) für `load` werden die angegebenen Eigenschaften für jedes dieser Elemente geladen. Im folgenden Beispiel wird die `resolved`-Eigenschaft für jedes `Comment`-Objekt im `CommentCollection`-Objekt eines Arbeitsblatts geladen.

```TypeScript
let comments = selectedSheet.comments;
comments.load("items/resolved"); // Load the `resolved` property from every comment in this collection.
await context.sync(); // Synchronize with the workbook to get the properties.
```

> [!TIP]
> Wenn Sie mehr über das Arbeiten mit Sammlungen in Office-Skripts wissen möchten, lesen Sie den [Array-Abschnitt des Artikels "Verwenden von integrierten JavaScript-Objekten in Office-Skripts"](javascript-objects.md#array).

## <a name="see-also"></a>Siehe auch

- [Aufzeichnen, Bearbeiten und Erstellen von Office-Skripts in Excel im Web](../tutorials/excel-tutorial.md)
- [Auslesen von Arbeitsmappendaten mit Office-Skripts in Excel im Web](../tutorials/excel-read-tutorial.md)
- [Referenzdokumentation zur Office Scripts-API](/javascript/api/office-scripts/overview)
- [Verwenden von integrierten JavaScript-Objekten in Office-Skripts](javascript-objects.md)
