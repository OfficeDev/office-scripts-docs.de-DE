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
# <a name="scripting-fundamentals-for-office-scripts-in-excel-on-the-web-preview"></a>Grundlegendes zur Skripterstellung für Office-Skripts in Excel im Internet (Vorschau)

In diesem Artikel werden die technischen Aspekte von Office-Skripts vorgestellt. Sie erfahren, wie die Excel-Objekte zusammenarbeiten und wie der Code-Editor mit einer Arbeitsmappe synchronisiert wird.

[!INCLUDE [Preview note](../includes/preview-note.md)]

## <a name="object-model"></a>Objektmodell

Um die Excel-APIs zu verstehen, müssen Sie verstehen, wie die Komponenten einer Arbeitsmappe miteinander verwandt sind.

- Eine **Arbeitsmappe** enthält ein oder mehrere **Arbeitsblätter**.
- Ein **Arbeitsblatt** ermöglicht den Zugriff auf Zellen über **Bereichs** Objekte.
- Ein **Bereich** stellt eine Gruppe von zusammenhängenden Zellen dar.
- **Bereiche** werden zum Erstellen und Platzieren von **Tabellen**, **Diagrammen**, **Formen**und anderen Daten Visualisierungs-oder Organisationsobjekten verwendet.
- Ein **Arbeitsblatt** enthält Auflistungen dieser Datenobjekte, die im einzelnen Blatt vorhanden sind.
- **Arbeitsmappen** enthalten Auflistungen einiger dieser Datenobjekte (beispielsweise **Tabellen**) für die gesamte **Arbeitsmappe**.

### <a name="ranges"></a>Bereiche

Ein Bereich ist eine Gruppe von zusammenhängenden Zellen in der Arbeitsmappe. Skripts verwenden normalerweise die a1-Notation (beispielsweise **B3** für die einzelne Zelle in Zeile **B** und Spalte **3** oder **C2: F4** für die Zellen von Zeilen **C** bis **F** und Spalten **2** bis **4**), um Bereiche zu definieren.

Bereiche weisen drei Haupteigenschaften auf `values`: `formulas`, und `format`. Diese Eigenschaften abrufen oder Festlegen der Zellenwerte, Formeln ausgewertet werden, und die visuelle Formatierung der Zellen.

#### <a name="range-sample"></a>Range-Beispiel

Im folgenden Beispiel wird gezeigt, wie Sie Umsatzdatensätze erstellen. Dieses Skript verwendet `Range` Objekte, um die Werte, Formeln und Formate festzulegen.

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

Beim Ausführen dieses Skripts werden die folgenden Daten im aktuellen Arbeitsblatt erstellt:

![Ein verkaufsdatensatz mit Wert Zeilen, einer Formelspalte und formatierten Kopfzeilen.](../images/range-sample.png)

### <a name="charts-tables-and-other-data-objects"></a>Diagramme, Tabellen und andere Datenobjekte

Skripts können die Datenstrukturen und Visualisierungen in Excel erstellen und bearbeiten. Tabellen und Diagramme sind zwei der häufiger verwendeten Objekte, die APIs unterstützen jedoch PivotTables, Formen, Bilder und vieles mehr.

#### <a name="creating-a-table"></a>Erstellen einer Tabelle

Erstellen von Tabellen mithilfe von Daten gefüllten Bereichen. Formatierungs-und Tabellensteuerelemente (wie Filter) werden automatisch auf den Bereich angewendet.

Das folgende Skript erstellt eine Tabelle mithilfe der Bereiche aus dem vorherigen Beispiel.

```TypeScript
async function main(context: Excel.RequestContext) {
   let sheet = context.workbook.worksheets.getActiveWorksheet();
   sheet.tables.add("B2:E5", true);
}
```

Wenn Sie dieses Skript auf dem Arbeitsblatt mit den vorherigen Daten ausführen, wird die folgende Tabelle erstellt:

![Eine Tabelle, die aus dem vorherigen Umsatzdatensatz erstellt wurde.](../images/table-sample.png)

#### <a name="creating-a-chart"></a>Erstellen eines Diagramms

Erstellen Sie Diagramme, um die Daten in einem Bereich zu visualisieren. Skripts ermöglichen Dutzende von Diagramm Varianten, die jeweils an Ihre Anforderungen angepasst werden können.

Das folgende Skript erstellt ein einfaches Säulendiagramm für drei Elemente und platziert es 100 Pixel unter dem oberen Rand des Arbeitsblatts.

```TypeScript
async function main(context: Excel.RequestContext) {
  let sheet = context.workbook.worksheets.getActiveWorksheet();
  let chart = sheet.charts.add(Excel.ChartType.columnStacked, sheet.getRange("B3:C5"));
  chart.top = 100;
}
```

Wenn Sie dieses Skript auf dem Arbeitsblatt mit der vorherigen Tabelle ausführen, wird das folgende Diagramm erstellt:

![Ein Säulendiagramm mit Mengen von drei Elementen aus dem vorherigen Umsatzdatensatz.](../images/chart-sample.png)

### <a name="further-reading-on-the-object-model"></a>Weiterführende Informationen zum Objektmodell

Die [Office Scripts-API-Referenzdokumentation](/javascript/api/office-scripts/overview) ist eine umfassende Auflistung der in Office-Skripts verwendeten Objekte. Dort können Sie das Inhaltsverzeichnis verwenden, um zu einer beliebigen Klasse zu navigieren, zu der Sie mehr erfahren möchten. Es folgen einige häufig angezeigte Seiten.

- [Chart](/javascript/api/office-scripts/excel/excel.chart)
- [Kommentar](/javascript/api/office-scripts/excel/excel.comment)
- [PivotTable](/javascript/api/office-scripts/excel/excel.pivottable)
- [Range](/javascript/api/office-scripts/excel/excel.range)
- [RangeFormat](/javascript/api/office-scripts/excel/excel.rangeformat)
- [Shape](/javascript/api/office-scripts/excel/excel.shape)
- [Table](/javascript/api/office-scripts/excel/excel.table)
- [Workbook](/javascript/api/office-scripts/excel/excel.workbook)
- [Arbeitsblatt](/javascript/api/office-scripts/excel/excel.worksheet)

## <a name="main-function"></a>`main`Funktion

Jedes Office-Skript muss eine `main` Funktion mit der folgenden Signatur enthalten, einschließlich `Excel.RequestContext` der Typendefinition:

```TypeScript
async function main(context: Excel.RequestContext) {
    // Your Excel Script
}
```

Der Code in der `main` -Funktion wird ausgeführt, wenn das Skript ausgeführt wird. `main`kann andere Funktionen in Ihrem Skript aufrufen, aber Code, der nicht in einer Funktion enthalten ist, wird nicht ausgeführt.

## <a name="context"></a>Kontext

Die `main` Funktion akzeptiert einen `Excel.RequestContext` Parameter mit dem `context`Namen. Denken Sie `context` an die Brücke zwischen Ihrem Skript und der Arbeitsmappe. Ihr Skript greift mit dem `context` Objekt auf die Arbeitsmappe zu und `context` verwendet dieses zum Senden von Daten hin und her.

Das `context` Objekt ist erforderlich, da das Skript und Excel in unterschiedlichen Prozessen und Speicherorten aktiv sind. Das Skript muss Änderungen an oder Abfragen von Daten aus der Arbeitsmappe in der Cloud vornehmen. Das `context` Objekt verwaltet diese Transaktionen.

## <a name="sync-and-load"></a>Synchronisieren und laden

Da das Skript und die Arbeitsmappe an unterschiedlichen Speicherorten ausgeführt werden, dauert die Datenübertragung zwischen den beiden Zeitpunkten. Um die Skript Leistung zu verbessern, werden Befehle in die Warteschlange eingereiht, bis das Skript den `sync` Vorgang explizit aufruft, um das Skript und die Arbeitsmappe zu synchronisieren. Ihr Skript kann unabhängig arbeiten, bis es eine der folgenden Aufgaben erfüllt:

- Lesen von Daten aus der Arbeitsmappe ( `load` nach einem Vorgang).
- Schreiben von Daten in die Arbeitsmappe (in der Regel, weil das Skript abgeschlossen ist).

In der folgenden Abbildung ist ein Beispiel für die Ablaufsteuerung zwischen Skript und Arbeitsmappe dargestellt:

![Ein Diagramm mit Lese-und Schreibvorgängen, die aus dem Skript zur Arbeitsmappe wechseln.](../images/load-sync.png)

### <a name="sync"></a>Synchronisierung

Wenn Ihr Skript Daten lesen oder Daten in die Arbeitsmappe schreiben muss, rufen Sie die `RequestContext.sync` -Methode wie hier gezeigt auf:

```TypeScript
await context.sync();
```

> [!NOTE]
> `context.sync()`wird implizit aufgerufen, wenn ein Skript beendet wird.

Nach Abschluss `sync` des Vorgangs wird die Arbeitsmappe so aktualisiert, dass alle Schreibvorgänge wiedergegeben werden, die von Skript angegeben wurden. Ein Schreibvorgang legt eine Eigenschaft für ein Excel-Objekt fest (Beispiels `range.format.fill.color = "red"`Weise) oder ruft eine Methode auf, die eine Eigenschaft ändert ( `range.format.autoFitColumns()`beispielsweise). Der `sync` Vorgang liest auch alle Werte aus der Arbeitsmappe, die das Skript mithilfe eines `load` Vorgangs angefordert hat (wie im nächsten Abschnitt beschrieben).

Die Synchronisierung Ihres Skripts mit der Arbeitsmappe kann je nach Netzwerkzeit in Anspruch nehmen. Sie sollten die Anzahl der `sync` Aufrufe minimieren, die dazu beitragen, dass Ihr Skript schnell ausgeführt wird.  

### <a name="load"></a>Load

Ein Skript muss Daten aus der Arbeitsmappe vor dem Lesen laden. Das häufiges Laden von Daten aus der gesamten Arbeitsmappe würde jedoch die Geschwindigkeit des Skripts erheblich reduzieren. Stattdessen ermöglicht die `load` Methode Ihrem Skript, die Daten, die aus der Arbeitsmappe abgerufen werden sollen, explizit anzugeben.

Die `load` -Methode ist für jedes Excel-Objekt verfügbar. Ihr Skript muss die Eigenschaften eines Objekts laden, bevor Sie diese lesen können. Andernfalls führt dies zu einem Fehler.

In den folgenden Beispielen wird `Range` ein Objekt verwendet, um die drei `load` Methoden anzuzeigen, mit denen die Methode zum Laden von Daten verwendet werden kann.

|Intent |Beispielbefehl | Effekt |
|:--|:--|:--|
|Laden einer Eigenschaft |`myRange.load("values");` | Lädt eine einzelne Eigenschaft, in diesem Fall das zweidimensionale Array von Werten in diesem Bereich. |
|Laden mehrerer Eigenschaften |`myRange.load("values, rowCount, columnCount");`| Lädt alle Eigenschaften aus einer durch Trennzeichen getrennten Liste, in diesem Beispiel die Werte, die Zeilenanzahl und die Spaltenanzahl. |
|Alles laden | `myRange.load();`|Lädt alle Eigenschaften des Bereichs. Dies ist keine empfohlene Lösung, da Ihr Skript dadurch verlangsamt wird, dass unnötige Daten abgerufen werden. Verwenden Sie dies nur, wenn Sie Ihr Skript testen oder wenn Sie jede Eigenschaft aus dem Objekt benötigen. |

Ihr Skript muss aufgerufen `context.sync()` werden, bevor geladene Werte gelesen werden.

```TypeScript
let range = selectedSheet.getRange("A1:B3");
range.load ("rowCount"); // Load the property.
await context.sync(); // Synchronize with the workbook to get the property.
console.log(range.rowCount); // Read and log the property value (3).
```

Sie können auch Eigenschaften für eine gesamte Auflistung laden. Jedes Collection-Objekt verfügt `items` über eine Eigenschaft, die ein Array mit den Objekten in dieser Auflistung darstellt. Verwenden `items` als Anfang eines hierarchischen Aufrufs (`items\myProperty`), `load` um die angegebenen Eigenschaften für jedes dieser Elemente zu laden. Im folgenden Beispiel wird die `resolved` Eigenschaft für jedes `Comment` Objekt im `CommentCollection` Objekt eines Arbeitsblatts geladen.

```TypeScript
let comments = selectedSheet.comments;
comments.load("items/resolved"); // Load the `resolved` property from every comment in this collection.
await context.sync(); // Synchronize with the workbook to get the properties.
```

> [!TIP]
> Weitere Informationen zum Arbeiten mit Auflistungen in Office-Skripts finden Sie im [Abschnitt "Array" im Artikel Verwenden integrierter JavaScript-Objekte in Office-Skripts](javascript-objects.md#array) .

## <a name="see-also"></a>Siehe auch

- [Aufzeichnen, bearbeiten und Erstellen von Office-Skripts in Excel im Internet](../tutorials/excel-tutorial.md)
- [Lesen von Arbeitsmappen-Daten mit Office-Skripts in Excel im Web](../tutorials/excel-read-tutorial.md)
- [Office Scripts-API-Referenz](/javascript/api/office-scripts/overview)
- [Verwenden integrierter JavaScript-Objekte in Office-Skripts](javascript-objects.md)
