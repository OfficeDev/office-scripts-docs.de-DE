---
title: Arbeiten mit PivotTables in Office-Skripts
description: Erfahren Sie mehr über das Objektmodell für PivotTables in der JavaScript-API für Office-Skripts.
ms.date: 04/20/2022
ms.localizationpriority: medium
ms.openlocfilehash: a457c41bd1205f4e17636c43d7ba78addc80d0e4
ms.sourcegitcommit: a6504f8b0d6b717457c6e0b5306c35ad3900914e
ms.translationtype: MT
ms.contentlocale: de-DE
ms.lasthandoff: 09/02/2022
ms.locfileid: "67572584"
---
# <a name="work-with-pivottables-in-office-scripts"></a>Arbeiten mit PivotTables in Office-Skripts

Mit PivotTables können Sie große Datensammlungen schnell analysieren. Mit ihrer Macht kommt Komplexität. Mit den Office-Skripts-APIs können Sie eine PivotTable an Ihre Anforderungen anpassen, aber der Umfang des API-Satzes macht die ersten Schritte zu einer Herausforderung. In diesem Artikel wird veranschaulicht, wie allgemeine PivotTable-Aufgaben ausgeführt werden, und wichtige Klassen und Methoden werden erläutert.

> [!NOTE]
> Um den Kontext für die von den APIs verwendeten Begriffe besser zu verstehen, lesen Sie zuerst die PivotTable-Dokumentation von Excel. Beginnen [Sie mit dem Erstellen einer PivotTable, um Arbeitsblattdaten zu analysieren](https://support.microsoft.com/office/a9a84538-bfe9-40a9-a8e9-f99134456576).

## <a name="object-model"></a>Objektmodell

:::image type="content" source="../images/pivottable-object-model.png" alt-text="Ein vereinfachtes Bild der Klassen, Methoden und Eigenschaften, die beim Arbeiten mit PivotTables verwendet werden.":::

Die [PivotTable](/javascript/api/office-scripts/excelscript/excelscript.pivottable) ist das zentrale Objekt für PivotTables in der Office-Skripts-API.

- Das [Workbook-Objekt](/javascript/api/office-scripts/excelscript/excelscript.workbook) verfügt über eine Auflistung aller [PivotTables](/javascript/api/office-scripts/excelscript/excelscript.pivottable). Jedes [Arbeitsblatt](/javascript/api/office-scripts/excelscript/excelscript.worksheet) enthält auch eine PivotTable-Auflistung, die für dieses Blatt lokal ist.
- Eine [PivotTable](/javascript/api/office-scripts/excelscript/excelscript.pivottable) enthält [PivotHierarchies](/javascript/api/office-scripts/excelscript/excelscript.pivothierarchy). Eine Hierarchie kann als Spalte in einer Tabelle betrachtet werden.
- [PivotHierarchies](/javascript/api/office-scripts/excelscript/excelscript.pivothierarchy) können als Zeilen oder Spalten ([RowColumnPivotHierarchy](/javascript/api/office-scripts/excelscript/excelscript.rowcolumnpivothierarchy)), Daten ([DataPivotHierarchy](/javascript/api/office-scripts/excelscript/excelscript.datapivothierarchy)) oder Filter ([FilterPivotHierarchy)](/javascript/api/office-scripts/excelscript/excelscript.filterpivothierarchy) hinzugefügt werden.
- Jedes [PivotHierarchy-Objekt](/javascript/api/office-scripts/excelscript/excelscript.pivothierarchy) enthält genau ein [PivotField](/javascript/api/office-scripts/excelscript/excelscript.pivotfield). PivotTable-Strukturen außerhalb von Excel können mehrere Felder pro Hierarchie enthalten, sodass dieser Entwurf zur Unterstützung zukünftiger Optionen vorhanden ist. Bei Office-Skripts werden Felder und Hierarchien denselben Informationen zugeordnet.
- Ein [PivotField](/javascript/api/office-scripts/excelscript/excelscript.pivotfield) enthält mehrere [PivotItems](/javascript/api/office-scripts/excelscript/excelscript.pivotitem). Jedes PivotItem ist ein eindeutiger Wert im Feld. Stellen Sie sich jedes Element als Wert in der Tabellenspalte vor. Elemente können auch aggregierte Werte sein, z. B. Summen, wenn das Feld für Daten verwendet wird.
- Das [PivotLayout](/javascript/api/office-scripts/excelscript/excelscript.pivotlayout) definiert, wie die [PivotFields](/javascript/api/office-scripts/excelscript/excelscript.pivotfield) und [PivotItems](/javascript/api/office-scripts/excelscript/excelscript.pivotitem) angezeigt werden.
- [PivotFilter](/javascript/api/office-scripts/excelscript/excelscript.pivotfilters) filtern Daten aus der [PivotTable](/javascript/api/office-scripts/excelscript/excelscript.pivottable) anhand unterschiedlicher Kriterien.

Schauen Sie sich an, wie diese Beziehungen in der Praxis funktionieren. Die folgenden Daten beschreiben den Umsatz von Obst aus verschiedenen Farmen. Es ist die Basis für alle Beispiele in diesem Artikel. Verwenden Sie [pivottable-sample.xlsx](pivottable-sample.xlsx) , um dies zu verfolgen.

:::image type="content" source="../images/pivottable-raw-data.png" alt-text="Eine Sammlung von Obstverkäufen verschiedener Arten von verschiedenen Farmen.":::

## <a name="create-a-pivottable-with-fields"></a>Erstellen einer PivotTable mit Feldern

PivotTables werden mit Verweisen auf vorhandene Daten erstellt. Sowohl Bereiche als auch Tabellen können die Quelle für eine PivotTable sein. Sie benötigen auch einen Ort in der Arbeitsmappe. Da die Größe einer PivotTable dynamisch ist, wird nur die obere linke Ecke des Zielbereichs angegeben.

Mit dem folgenden Codeausschnitt wird eine PivotTable basierend auf einem Datenbereich erstellt. Die PivotTable verfügt über keine Hierarchien, daher sind die Daten noch nicht in irgendeiner Weise gruppiert.

```typescript
  const dataSheet = workbook.getWorksheet("Data");
  const pivotSheet = workbook.getWorksheet("Pivot");

  const farmPivot = pivotSheet.addPivotTable(
    "Farm Pivot", /* The name of the PivotTable. */
    dataSheet.getUsedRange(), /* The source data range. */
    pivotSheet.getRange("A1") /* The location to put the new PivotTable. */);
```

:::image type="content" source="../images/pivottable-empty.png" alt-text="Eine PivotTable mit dem Namen &quot;Farmpivot&quot; ohne Hierarchien.":::

### <a name="hierarchies-and-fields"></a>Hierarchien und Felder

PivotTables werden über Hierarchien organisiert. Diese Hierarchien werden verwendet, um Daten zu pivotieren, wenn sie als einen bestimmten Hierarchietyp hinzugefügt werden. Es gibt vier Arten von Hierarchien.

- **Zeile**: Zeigt Elemente in horizontalen Zeilen an.
- **Spalte**: Zeigt Elemente in vertikalen Spalten an.
- **Daten**: Zeigt Aggregate von Werten basierend auf den Zeilen und Spalten an.
- **Filter**: Hinzufügen oder Entfernen von Elementen aus der PivotTable.

Einer PivotTable können so viele oder nur wenige Felder zugewiesen sein, die diesen spezifischen Hierarchien zugewiesen sind. Eine PivotTable benötigt mindestens eine Datenhierarchie, um zusammengefasste numerische Daten anzuzeigen, und mindestens eine Zeile oder Spalte, um diese Zusammenfassung zu pivotieren. Im folgenden Codeausschnitt werden zwei Zeilenhierarchien und zwei Datenhierarchien hinzugefügt.

```typescript
  farmPivot.addRowHierarchy(farmPivot.getHierarchy("Farm"));
  farmPivot.addRowHierarchy(farmPivot.getHierarchy("Type"));
  farmPivot.addDataHierarchy(farmPivot.getHierarchy("Crates Sold at Farm"));
  farmPivot.addDataHierarchy(farmPivot.getHierarchy("Crates Sold Wholesale"));
```

:::image type="content" source="../images/pivottable-data-hierarchy.png" alt-text="Eine PivotTable, in der der Gesamtumsatz verschiedener Früchte basierend auf der Farm angezeigt wird, aus der sie stammen.":::

## <a name="layout-ranges"></a>Layoutbereiche

Jeder Teil der PivotTable wird einem Bereich zugeordnet. Auf diese Weise kann Ihr Skript Daten aus der PivotTable abrufen, die später im Skript verwendet werden oder in einem [Power Automate-Fluss](power-automate-integration.md) zurückgegeben werden können. Auf diese Bereiche wird über das [PivotLayout](/javascript/api/office-scripts/excelscript/excelscript.pivotlayout) -Objekt zugegriffen, das von abgerufen wird `PivotTable.getLayout()`. Das folgende Diagramm zeigt die Bereiche, die von den Methoden in `PivotLayout`zurückgegeben werden.

:::image type="content" source="../images/pivottable-layout-breakdown.png" alt-text="Ein Diagramm, das zeigt, welche Abschnitte einer PivotTable von den Get-Bereichsfunktionen des Layouts zurückgegeben werden.":::

## <a name="filters-and-slicers"></a>Filter und Datenschnitte

Es gibt drei Möglichkeiten zum Filtern einer PivotTable.

- [FilterPivotHierarchies](/javascript/api/office-scripts/excelscript/excelscript.filterpivothierarchy)
- [PivotFilters](/javascript/api/office-scripts/excelscript/excelscript.pivotfilters)
- [Slicers](/javascript/api/office-scripts/excelscript/excelscript.slicer)

### <a name="filterpivothierarchies"></a>FilterPivotHierarchies

`FilterPivotHierarchies` fügen Sie eine zusätzliche Hierarchie hinzu, um jede Datenzeile zu filtern. Jede Zeile mit einem Element, das herausgefiltert ist, wird aus der PivotTable und deren Zusammenfassungen ausgeschlossen. Da diese Filter auf Elementen basieren, funktionieren sie nur mit einzelnen Werten. Wenn "Klassifizierung" eine Filterhierarchie in unserem Beispiel ist, können Benutzer die Werte "Organisch" und "Konventionell" für den Filter auswählen. Wenn "Kisten verkaufter Großhandel" ausgewählt wird, wären die Filteroptionen die einzelnen Zahlen, z. B. 120 und 150, anstelle numerischer Bereiche.

`FilterPivotHierarchies` werden mit allen ausgewählten Werten erstellt. Dies bedeutet, dass nichts gefiltert wird, bis der Benutzer manuell mit dem Filtersteuerelement interagiert oder ein `PivotManualFilter` Element für das Feld festgelegt ist, das `FilterPivotHierarchy`zum .

Der folgende Codeausschnitt fügt "Classification" als Filterhierarchie hinzu.

```typescript
  farmPivot.addFilterHierarchy(farmPivot.getHierarchy("Classification"));
```

:::image type="content" source="../images/pivottable-filter-hierarchy.png" alt-text="Ein Filtersteuerelement, das &quot;Klassifizierung&quot; für eine PivotTable verwendet.":::

### <a name="pivotfilters"></a>PivotFilters

Das `PivotFilters` Objekt ist eine Auflistung von Filtern, die auf ein einzelnes Feld angewendet werden. Da jede Hierarchie genau ein Feld aufweist, sollten Sie beim Anwenden von Filtern immer das erste Feld `PivotHierarchy.getFields()` verwenden. Es gibt vier Filtertypen.

- **Datumsfilter**: Kalenderdatumsbasierte Filterung.
- **Bezeichnungsfilter**: Textvergleichsfilterung.
- **Manueller Filter**: Benutzerdefinierte Eingabefilterung.
- **Wertfilter**: Zahlenvergleichsfilterung. Dadurch werden Elemente in der zugeordneten Hierarchie mit Werten in einer angegebenen Datenhierarchie verglichen.

In der Regel wird nur einer der vier Filtertypen erstellt und auf das Feld angewendet. Wenn das Skript versucht, inkompatible Filter zu verwenden, wird ein Fehler mit dem Text "Das Argument ist ungültig oder fehlt oder hat ein falsches Format" ausgelöst.

Im folgenden Codeausschnitt werden zwei Filter hinzugefügt. Der erste ist ein manueller Filter, der Elemente in einer vorhandenen Filterhierarchie "Klassifizierung" auswählt. Der zweite Filter entfernt alle Farmen, die weniger als 300 "Krates Sold Wholesale" haben. Beachten Sie, dass dadurch die "Summe" dieser Farmen herausfiltert, nicht die einzelnen Zeilen aus den ursprünglichen Daten.

```typescript
  const classificationField = farmPivot.getFilterHierarchy("Classification").getFields()[0];
  classificationField.applyFilter({
    manualFilter: { 
      selectedItems: ["Organic"] /* The included items. */
    }
  });

  const farmField = farmPivot.getHierarchy("Farm").getFields()[0];
  farmField.applyFilter({
    valueFilter: {
      condition: ExcelScript.ValueFilterCondition.greaterThan, /* The relationship of the value to the comparator. */
      comparator: 300, /* The value to which items are compared. */
      value: "Sum of Crates Sold Wholesale" /* The name of the data hierarchy. Note the "Sum of" prefix. */
      }
  });
```

:::image type="content" source="../images/pivottable-filters.png" alt-text="Eine PivotTable nach der Anwendung des Wertfilters und des manuellen Filters.":::

### <a name="slicers"></a>Datenschnitte

[Datenschnitte](https://support.microsoft.com/office/249f966b-a9d5-4b0f-b31a-12651785d29d) filtern Daten in einer PivotTable (oder Standardtabelle). Sie sind verschiebbare Objekte im Arbeitsblatt, die eine schnelle Filterauswahl ermöglichen. Ein Datenschnitt arbeitet ähnlich wie der manuelle Filter und `PivotFilterHierarchy`. Elemente aus der `PivotField` PivotTable werden umgeschaltet, um sie ein- oder aus der PivotTable auszuschließen.

Der folgende Codeausschnitt fügt einen Datenschnitt für das Feld "Typ" hinzu. Die ausgewählten Elemente werden auf "Lemon" und "Lime" festgelegt, und der Datenschnitt wird dann um 400 Pixel nach links verschoben.

```typescript
  const fruitSlicer = pivotSheet.addSlicer(
    farmPivot, /* The table or PivotTale to be sliced. */
    farmPivot.getHierarchy("Type").getFields()[0] /* What source to use as the slicer options. */
  );
  fruitSlicer.selectItems(["Lemon", "Lime"]);
  fruitSlicer.setLeft(400);
```

:::image type="content" source="../images/slicer.png" alt-text="Ein Datenschnitt, der Daten in einer PivotTable filtert.":::

## <a name="see-also"></a>Siehe auch

- [Grundlagen der Skripterstellung für Office-Skripts in Excel im Web](scripting-fundamentals.md)
- [Referenzdokumentation zur Office Scripts-API](/javascript/api/office-scripts/overview)
