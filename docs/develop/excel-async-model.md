---
title: Verwenden der Office Scripts Async-APIs zur Unterstützung von älteren Skripts
description: Eine Einführung in die Async-APIs für Office-Skripts und die Verwendung des Musters zum Laden/synchronisieren für ältere Skripts.
ms.date: 06/29/2020
localization_priority: Normal
ms.openlocfilehash: 6c31a39c8e1fe53f2f5587183a6b32e100d2b457
ms.sourcegitcommit: bf9f33c37c6f7805d6b408aa648bb9785a7cd133
ms.translationtype: MT
ms.contentlocale: de-DE
ms.lasthandoff: 07/06/2020
ms.locfileid: "45043398"
---
# <a name="using-the-office-scripts-async-apis-to-support-legacy-scripts"></a>Verwenden der Office Scripts Async-APIs zur Unterstützung von älteren Skripts

In diesem Artikel erfahren Sie, wie Sie Skripts mit den Legacy-, Async-APIs schreiben. Diese APIs haben die gleiche Kernfunktionalität wie die synchronen Office Scripts-APIs, aber Sie erfordern, dass Ihr Skript die Datensynchronisierung zwischen dem Skript und der Arbeitsmappe steuert.

> [!IMPORTANT]
> Das Async-Modell kann nur mit Skripts verwendet werden, die vor der Implementierung des aktuellen [API-Modells](scripting-fundamentals.md?view=office-scripts)erstellt wurden. Skripts sind dauerhaft mit dem API-Modell gesperrt, das Sie bei der Erstellung haben. Dies bedeutet auch, dass Sie ein nagelneues Skript verwenden müssen, wenn Sie ein Legacy Skript in das neue Modell konvertieren möchten. Es wird empfohlen, die alten Skripts beim Vornehmen von Änderungen auf das neue Modell zu aktualisieren, da das aktuelle Modell einfacher zu verwenden ist. Die [Konvertierung von asynchronen Legacy Skripts in den Abschnitt Aktuelles Modell](#converting-legacy-async-scripts-to-the-current-model) enthält Ratschläge dazu, wie Sie diesen Übergang vornehmen.

## <a name="main-function"></a>Die `main`-Funktion

Skripts, die die Async-APIs verwenden, haben eine andere `main` Funktion. Es handelt sich um eine `async` Funktion, die einen `Excel.RequestContext` als ersten Parameter aufweist.

```TypeScript
async function main(context: Excel.RequestContext) {
    // Your async Office Script
}
```

## <a name="context"></a>Context

Die `main`-Funktion lässt einen `Excel.RequestContext`-Parameter namens `context` zu. Stellen Sie sich `context` als die Brücke zwischen Ihrem Skript und der Arbeitsmappe vor. Das Skript greift auf die Arbeitsmappe mit dem `context`-Objekt zu und verwendet diesen `context` zum Hin- und Hersenden von Daten.

Das `context`-Objekt ist erforderlich, weil das Skript und Excel in unterschiedlichen Prozessen und Speicherorten ausgeführt werden. Das Skript muss Änderungen an den Daten in der Arbeitsmappe in der Cloud vornehmen oder diese abrufen können. Das `context`-Objekt verwaltet diese Transaktionen.

## <a name="sync-and-load"></a>Synchronisieren und Laden

Da Ihr Skript und die Arbeitsmappe an unterschiedlichen Orten ausgeführt werden, dauert die Datenübertragung zwischen diesen etwas. In der Async-API werden Befehle in die Warteschlange eingereiht, bis das Skript den Vorgang explizit aufruft `sync` , um das Skript und die Arbeitsmappe zu synchronisieren. Ihr Skript kann unabhängig funktionieren, bis es eine der folgenden Aktionen durchführen muss:

- Daten aus der Arbeitsmappe lesen (nach einem `load`-Vorgang oder einer Methode, die ein[ClientResultat](/javascript/api/office-scripts/excelscript/excelscript.clientresult?view=office-scripts-async) zurückgibt).
- Daten in die Arbeitsmappe schreiben (in der Regel, weil das Skript abgeschlossen wurde).

In der folgenden Abbildung wird ein Beispiel für eine Ablaufsteuerung zwischen dem Skript und der Arbeitsmappe dargestellt:

![Ein Diagramm mit Lese- und Schreibvorgängen, die vom Skript in der Arbeitsmappe ausgeführt werden.](../images/load-sync.png)

### <a name="sync"></a>Synchronisierung

Wenn Ihr Async-Skript Daten lesen oder Daten in die Arbeitsmappe schreiben muss, rufen Sie die `RequestContext.sync` -Methode wie hier gezeigt auf:

```TypeScript
await context.sync();
```

> [!NOTE]
> `context.sync()` wird implizit aufgerufen, wenn ein Skript endet.

Nachdem der `sync`-Vorgang abgeschlossen ist, wird die Arbeitsmappe entsprechend den Schreibvorgängen aktualisiert, die vom Skript angegeben wurden. Bei einem Schreibvorgang wird eine beliebige Eigenschaft eines Excel-Objekts festgelegt (z. B. `range.format.fill.color = "red"`) oder eine Methode aufgerufen, über die eine Eigenschaft geändert wird (z. B. `range.format.autoFitColumns()`). Der `sync`-Vorgang liest auch alle Werte aus der Arbeitsmappe, die das Skript angefordert hat, indem es einen `load`-Vorgang oder eine Methode verwendet, die ein `ClientResult` zurückgibt (wie in den nächsten Abschnitten besprochen).

Je nach Netzwerk kann es einige Zeit dauern, bis das Skript mit der Arbeitsmappe synchronisiert wurde. Minimieren Sie die Anzahl der `sync` Aufrufe, mit denen Ihr Skript schnell ausgeführt werden kann. Andernfalls sind die asynchronen APIs nicht schneller die Standard synchronen APIs.

### <a name="load"></a>Laden

Ein Async-Skript muss Daten aus der Arbeitsmappe vor dem Lesen laden. Das Laden von Daten aus der gesamten Arbeitsmappe würde jedoch die Geschwindigkeit des Skripts erheblich reduzieren. `load`Mit der-Methode kann Ihr Skript explizit angeben, welche Daten aus der Arbeitsmappe abgerufen werden sollen.

Die `load`-Methode ist für jedes Excel-Objekt verfügbar. Ihr Skript muss die Eigenschaften eines Objekts laden, bevor es sie lesen kann. Wenn Sie dies nicht tun, tritt ein Fehler auf.

In den folgenden Beispielen wird ein `Range`-Objekt verwendet, um die drei Arten darzustellen, wie die `load`-Methode zum Laden von Daten verwendet werden kann.

|Absicht |Beispielbefehl | Auswirkung |
|:--|:--|:--|
|Laden einer Eigenschaft |`myRange.load("values");` | Lädt eine einzelne Eigenschaft, in diesem Fall den zweidimensionalen Wertearray in diesem Bereich. |
|Laden mehrerer Eigenschaften |`myRange.load("values, rowCount, columnCount");`| Lädt alle Eigenschaften aus einer durch Kommas getrennten Liste, in diesem Beispiel die Werte, die Zeilenanzahl und die Spaltenanzahl. |
|Alles laden | `myRange.load();`|Lädt alle Eigenschaften des Zellbereichs. Dies ist keine empfohlene Lösung, da Ihr Skript dadurch verlangsamt wird, dass unnötige Daten abgerufen werden. Verwenden Sie dies nur, wenn Sie Ihr Skript testen oder wenn Sie jede Eigenschaft aus dem Objekt benötigen. |

Ihr Skript muss `context.sync()` aufrufen, bevor es geladene Werte ausliest.

```TypeScript
/**
 * This script uses the async API to get the row count for a range.
 * It shows how to load a property in the async model.
 */
async function main(context: Excel.RequestContext) {
    let selectedSheet = context.workbook.worksheets.getActiveWorksheet();
    let range = selectedSheet.getRange("A1:B3");

    // Load the property.
    range.load("rowCount");

    // Synchronize with the workbook to get the property.
    await context.sync();

    // Read and log the property value (3).
    console.log(range.rowCount);
}
```

Sie können auch Eigenschaften aus einer ganzen Sammlung laden. Jedes Collection-Objekt in der Async-API verfügt über eine `items` Eigenschaft, die ein Array mit den Objekten in dieser Auflistung darstellt. Durch die Verwendung von `items` als Anfang eines hierarchischen Aufrufs (`items\myProperty`) für `load` werden die angegebenen Eigenschaften für jedes dieser Elemente geladen. Im folgenden Beispiel wird die `resolved`-Eigenschaft für jedes `Comment`-Objekt im `CommentCollection`-Objekt eines Arbeitsblatts geladen.

```TypeScript
/**
 * This script uses the async API to get resolved property on every comment in the worksheet.
 * It shows how to load a property from every object in a collection.
 */
async function main(context: Excel.RequestContext){
    let selectedSheet = context.workbook.worksheets.getActiveWorksheet();
    let comments = selectedSheet.comments;

    // Load the `resolved` property from every comment in this collection.
    comments.load("items/resolved");

    // Synchronize with the workbook to get the properties.
    await context.sync();
}
```

### <a name="clientresult"></a>ClientResult

Methoden in der Async-API, die Informationen aus der Arbeitsmappe zurückgeben, weisen ein ähnliches Muster wie das `load` / `sync` Paradigma auf. `TableCollection.getCount` ruft zum Beispiel die Anzahl von Tabellen in der Auflistung ab. `getCount`gibt a zurück `ClientResult<number>` , was bedeutet, dass die `value` Eigenschaft im zurückgegebenen [`ClientResult`](/javascript/api/office-scripts/excelscript/excelscript.clientresult?view=office-scripts-async) eine Zahl ist. Ihr Skript kann erst auf diesen Wert zugreifen, wenn `context.sync()` aufgerufen wird. Ähnlich wie beim Laden einer Eigenschaft ist der `value` bis zu diesem `sync`-Aufruf ein lokaler "leerer" Wert.

Das folgende Skript ruft die Gesamtanzahl der Tabellen in der Arbeitsmappe ab und protokolliert diese Anzahl in der Konsole.

```TypeScript
/**
 * This script uses the async API to get the table count of the workbook.
 * It shows how ClientResult objects return workbook information.
 */
async function main(context: Excel.RequestContext) {
    let tableCount = context.workbook.tables.getCount();

    // This sync call implicitly loads tableCount.value.
    // Any other ClientResult values are loaded too.
    await context.sync();

    // Trying to log the value before calling sync would throw an error.
    console.log(tableCount.value);
}
```

## <a name="converting-legacy-async-scripts-to-the-current-model"></a>Konvertieren von asynchronen Legacy Skripts in das aktuelle Modell

Das aktuelle API-Modell verwendet weder, `load` `sync` noch eine `RequestContext` . Dadurch wird das Schreiben und Verwalten der Skripts wesentlich vereinfacht. Die beste Ressource für die Konvertierung Alter Skripts ist der [Stapelüberlauf](https://stackoverflow.com/questions/tagged/office-scripts). Dort können Sie die Community um Hilfe bei bestimmten Szenarien bitten. Die folgenden Anleitungen sollten Ihnen helfen, die allgemeinen Schritte zu erläutern, die Sie ausführen müssen.

1. Erstellen Sie ein neues Skript, und kopieren Sie den alten Async-Code hinein. Achten Sie darauf, die alte `main` Methodensignatur nicht mit dem aktuellen, `function main(workbook: ExcelScript.Workbook)` sondern mit einzubeziehen.

2. Entfernen Sie alle `load` und `sync` ruft. Sie sind nicht mehr erforderlich.

3. Alle Eigenschaften wurden entfernt. Sie können nun über und Methoden auf diese Objekte zugreifen `get` `set` , sodass Sie diese Eigenschaftenverweise auf Methodenaufrufe umstellen müssen. Anstatt beispielsweise die Füllfarbe einer Zelle durch den Eigenschaftenzugriff wie folgt festzulegen: `mySheet.getRange("A2:C2").format.fill.color = "blue";` , verwenden Sie jetzt Methoden wie die folgende:`mySheet.getRange("A2:C2").getFormat().getFill().setColor("blue");`

4. Auflistungsklassen wurden durch Arrays ersetzt. Die `add` und- `get` Methoden dieser Auflistungsklassen wurden auf das Objekt verschoben, das die Auflistung besaß, sodass Ihre Verweise entsprechend aktualisiert werden müssen. Um beispielsweise ein Diagramm mit dem Namen "myChart" aus dem ersten Arbeitsblatt in der Arbeitsmappe abzurufen, verwenden Sie den folgenden Code: `workbook.getWorksheets()[0].getChart("MyChart");` . Beachten Sie `[0]` , dass für den Zugriff auf den ersten Wert der `Worksheet[]` zurückgegebenen `getWorksheets()` .

5. Einige Methoden wurden aus Gründen der Übersichtlichkeit umbenannt und zur Vereinfachung hinzugefügt. Weitere Informationen finden Sie in der [Office Scripts-API-Referenz](/javascript/api/office-scripts/overview?view=office-scripts) .

## <a name="office-scripts-async-api-reference-documentation"></a>Office Scripts Async-API-Referenzdokumentation

[!INCLUDE [Async reference documentation](../includes/async-reference-documentation-link.md)]
