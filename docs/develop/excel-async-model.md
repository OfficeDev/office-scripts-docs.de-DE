---
title: Unterstützen älterer Office-Skripts, die die asynchronen APIs verwenden
description: Eine Einführung in die Office Scripts Async-APIs und die Verwendung des Lade-/Synchronisierungsmusters für ältere Skripts.
ms.date: 02/08/2021
localization_priority: Normal
ms.openlocfilehash: d61a5d8affae2077b23e140645c19dac977ff0d2
ms.sourcegitcommit: 5d24e77df70aa2c1c982275d53213c2a9323ff86
ms.translationtype: MT
ms.contentlocale: de-DE
ms.lasthandoff: 04/02/2021
ms.locfileid: "51570283"
---
# <a name="support-older-office-scripts-that-use-the-async-apis"></a>Unterstützen älterer Office-Skripts, die die asynchronen APIs verwenden

In diesem Artikel erfahren Sie, wie Sie Skripts verwalten und aktualisieren, die die asynchronen APIs des älteren Modells verwenden. Diese APIs verfügen über die gleiche Kernfunktionalität wie die jetzt standardmäßigen synchronen Office Scripts-APIs, sie erfordern jedoch, dass Ihr Skript die Datensynchronisierung zwischen dem Skript und der Arbeitsmappe steuern kann.

> [!IMPORTANT]
> Das asynchrone Modell kann nur mit Skripts verwendet werden, die vor der Implementierung des aktuellen [API-Modells erstellt wurden.](scripting-fundamentals.md) Skripts sind dauerhaft für das API-Modell gesperrt, das sie bei der Erstellung haben. Dies bedeutet auch, dass Sie ein ganz neues Skript erstellen müssen, wenn Sie ein altes Skript in das neue Modell konvertieren möchten. Es wird empfohlen, die alten Skripts beim Vornehmen von Änderungen auf das neue Modell zu aktualisieren, da das aktuelle Modell einfacher zu verwenden ist. Der [Abschnitt Konvertieren von asynchronen Skripts](#converting-async-scripts-to-the-current-model) in das aktuelle Modell enthält Hinweise zum Übergang.

## <a name="main-function"></a>Die `main`-Funktion

Skripts, die die asynchronen APIs verwenden, haben eine andere `main` Funktion. Es ist eine `async` Funktion, die einen `Excel.RequestContext` als ersten Parameter hat.

```TypeScript
async function main(context: Excel.RequestContext) {
    // Your async Office Script
}
```

## <a name="context"></a>Context

Die `main`-Funktion lässt einen `Excel.RequestContext`-Parameter namens `context` zu. Stellen Sie sich `context` als die Brücke zwischen Ihrem Skript und der Arbeitsmappe vor. Das Skript greift auf die Arbeitsmappe mit dem `context`-Objekt zu und verwendet diesen `context` zum Hin- und Hersenden von Daten.

Das `context`-Objekt ist erforderlich, weil das Skript und Excel in unterschiedlichen Prozessen und Speicherorten ausgeführt werden. Das Skript muss Änderungen an den Daten in der Arbeitsmappe in der Cloud vornehmen oder diese abrufen können. Das `context`-Objekt verwaltet diese Transaktionen.

## <a name="sync-and-load"></a>Synchronisieren und Laden

Da Ihr Skript und die Arbeitsmappe an unterschiedlichen Orten ausgeführt werden, dauert die Datenübertragung zwischen diesen etwas. In der asynchronen API werden Befehle in die Warteschlange eingereiht, bis das Skript den Vorgang explizit aufruft, um das Skript und die `sync` Arbeitsmappe zu synchronisieren. Ihr Skript kann unabhängig funktionieren, bis es eine der folgenden Aktionen durchführen muss:

- Daten aus der Arbeitsmappe lesen (nach einem `load`-Vorgang oder einer Methode, die ein[ClientResultat](/javascript/api/office/officeextension.clientresult?view=excel-js-online&preserve-view=true) zurückgibt).
- Daten in die Arbeitsmappe schreiben (in der Regel, weil das Skript abgeschlossen wurde).

In der folgenden Abbildung wird ein Beispiel für eine Ablaufsteuerung zwischen dem Skript und der Arbeitsmappe dargestellt:

![Ein Diagramm mit Lese- und Schreibvorgängen, die vom Skript in der Arbeitsmappe ausgeführt werden.](../images/load-sync.png)

### <a name="sync"></a>Synchronisierung

Wenn Ihr asynchrones Skript Daten aus der Arbeitsmappe lesen oder in die Arbeitsmappe schreiben muss, rufen Sie die `RequestContext.sync` -Methode wie hier gezeigt auf:

```TypeScript
await context.sync();
```

> [!NOTE]
> `context.sync()` wird implizit aufgerufen, wenn ein Skript endet.

Nachdem der `sync`-Vorgang abgeschlossen ist, wird die Arbeitsmappe entsprechend den Schreibvorgängen aktualisiert, die vom Skript angegeben wurden. Ein Schreibvorgang setzt eine beliebige Eigenschaft für ein Excel-Objekt (z. B. ) oder ruft eine Methode auf, die eine Eigenschaft `range.format.fill.color = "red"` ändert (z. B. `range.format.autoFitColumns()` ). Der `sync`-Vorgang liest auch alle Werte aus der Arbeitsmappe, die das Skript angefordert hat, indem es einen `load`-Vorgang oder eine Methode verwendet, die ein `ClientResult` zurückgibt (wie in den nächsten Abschnitten besprochen).

Je nach Netzwerk kann es einige Zeit dauern, bis das Skript mit der Arbeitsmappe synchronisiert wurde. Minimieren Sie die Anzahl `sync` der Aufrufe, damit Ihr Skript schnell ausgeführt werden kann. Andernfalls sind die asynchronen APIs nicht schneller als die standardmäßigen synchronen APIs.

### <a name="load"></a>Laden

Ein asynchrones Skript muss Daten aus der Arbeitsmappe laden, bevor es gelesen wird. Das Laden von Daten aus der gesamten Arbeitsmappe würde jedoch die Geschwindigkeit des Skripts erheblich reduzieren. Mit `load` der -Methode kann Ihr Skript explizit die Daten aus der Arbeitsmappe abrufen.

Die `load`-Methode ist für jedes Excel-Objekt verfügbar. Ihr Skript muss die Eigenschaften eines Objekts laden, bevor es sie lesen kann. Dies führt zu einem Fehler.

In den folgenden Beispielen wird ein `Range`-Objekt verwendet, um die drei Arten darzustellen, wie die `load`-Methode zum Laden von Daten verwendet werden kann.

|Absicht |Beispielbefehl | Auswirkung |
|:--|:--|:--|
|Laden einer Eigenschaft |`myRange.load("values");` | Lädt eine einzelne Eigenschaft, in diesem Fall den zweidimensionalen Wertearray in diesem Bereich. |
|Laden mehrerer Eigenschaften |`myRange.load("values, rowCount, columnCount");`| Lädt alle Eigenschaften aus einer durch Kommas getrennten Liste, in diesem Beispiel die Werte, die Zeilenanzahl und die Spaltenanzahl. |
|Alles laden | `myRange.load();`|Lädt alle Eigenschaften des Zellbereichs. Dies ist keine empfohlene Lösung, da ihr Skript dadurch verlangsamt wird, dass unnötige Daten erhalten werden. Verwenden Sie dies nur, wenn Sie Ihr Skript testen oder wenn Sie jede Eigenschaft aus dem Objekt benötigen. |

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

Sie können auch Eigenschaften aus einer ganzen Sammlung laden. Jedes Auflistungsobjekt in der asynchronen API verfügt über eine Eigenschaft, bei der es sich um ein Array handelt, das `items` die Objekte in dieser Auflistung enthält. Durch die Verwendung von `items` als Anfang eines hierarchischen Aufrufs (`items\myProperty`) für `load` werden die angegebenen Eigenschaften für jedes dieser Elemente geladen. Im folgenden Beispiel wird die `resolved`-Eigenschaft für jedes `Comment`-Objekt im `CommentCollection`-Objekt eines Arbeitsblatts geladen.

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

Methoden in der asynchronen API, die Informationen aus der Arbeitsmappe zurückgeben, haben ein ähnliches Muster wie das `load` / `sync` Paradigma. `TableCollection.getCount` ruft zum Beispiel die Anzahl von Tabellen in der Auflistung ab. `getCount` gibt ein `ClientResult<number>` zurück, was bedeutet, dass die `value`-Eigenschaft in der zurückgegebenen [`ClientResult`](/javascript/api/office/officeextension.clientresult?view=excel-js-online&preserve-view=true) eine Zahl ist. Ihr Skript kann erst auf diesen Wert zugreifen, wenn `context.sync()` aufgerufen wird. Ähnlich wie beim Laden einer Eigenschaft ist der `value` bis zu diesem `sync`-Aufruf ein lokaler "leerer" Wert.

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

## <a name="converting-async-scripts-to-the-current-model"></a>Konvertieren von asynchronen Skripts in das aktuelle Modell

Das aktuelle API-Modell verwendet keine `load` , `sync` oder einen `RequestContext` . Dies erleichtert das Schreiben und Verwalten der Skripts. Die beste Ressource zum Konvertieren alter Skripts ist [Stack Overflow](https://stackoverflow.com/questions/tagged/office-scripts). Dort können Sie die Community um Hilfe bei bestimmten Szenarien bitten. Die folgenden Anleitungen sollten Ihnen dabei helfen, die allgemeinen Schritte zu beschreiben, die Sie ausführen müssen.

1. Erstellen Sie ein neues Skript, und kopieren Sie den alten asynchronen Code in das Skript. Achten Sie darauf, die alte Methodensignatur `main` nicht mit der aktuellen zu `function main(workbook: ExcelScript.Workbook)` verwenden.

2. Entfernen Sie alle `load` `sync` Und-Aufrufe. Sie sind nicht mehr erforderlich.

3. Alle Eigenschaften wurden entfernt. Sie greifen nun über und Methoden auf diese Objekte zu, sodass Sie diese Eigenschaftsverweise auf `get` `set` Methodenaufrufe umstellen müssen. Anstatt beispielsweise die Füllfarbe einer Zelle über den Eigenschaftszugriff wie die folgende zu setzen: , verwenden Sie jetzt `mySheet.getRange("A2:C2").format.fill.color = "blue";` Methoden wie die folgende: `mySheet.getRange("A2:C2").getFormat().getFill().setColor("blue");`

4. Auflistungsklassen wurden durch Arrays ersetzt. Die `add` Und-Methoden dieser Auflistungsklassen wurden in das Objekt verschoben, das im Besitz der Auflistung war, sodass Ihre `get` Verweise entsprechend aktualisiert werden müssen. Verwenden Sie beispielsweise den folgenden Code, um ein Diagramm mit dem Namen "MyChart" aus dem ersten Arbeitsblatt in der Arbeitsmappe zu erhalten: `workbook.getWorksheets()[0].getChart("MyChart");` . Notieren `[0]` Sie sich den, um auf den ersten Wert des `Worksheet[]` zurückgegebenen von zu `getWorksheets()` zugreifen.

5. Einige Methoden wurden aus Gründen der Übersichtlichkeit umbenannt und zur Vereinfachung hinzugefügt. Weitere Informationen finden Sie in [der Office Scripts-API-Referenz.](/javascript/api/office-scripts/overview)

## <a name="office-scripts-async-api-reference-documentation"></a>Dokumentation zur asynchronen API-Referenz für Office Scripts

Die asynchronen APIs entsprechen denen in Office-Add-Ins. Die Referenzdokumentation finden Sie im [Abschnitt Excel der JavaScript-API-Referenz für Office-Add-Ins.](/javascript/api/excel?view=excel-js-online&preserve-view=true)
