---
title: Unterstützung älterer Office Skripts, die die asynchronen APIs verwenden
description: Ein Primer auf den Office Scripts Async-APIs und wie das Lade-/Synchronisierungsmuster für ältere Skripts verwendet wird.
ms.date: 05/17/2021
localization_priority: Normal
ms.openlocfilehash: 80a1c0dec5393d8882ddb37eea5f81ef23b1ebb1
ms.sourcegitcommit: 4687693f02fc90a57ba30c461f35046e02e6f5fb
ms.translationtype: MT
ms.contentlocale: de-DE
ms.lasthandoff: 05/19/2021
ms.locfileid: "52545075"
---
# <a name="support-older-office-scripts-that-use-the-async-apis"></a>Unterstützung älterer Office Skripts, die die asynchronen APIs verwenden

In diesem Artikel erfahren Sie, wie Sie Skripts verwalten und aktualisieren, die die asynchronen APIs des älteren Modells verwenden. Diese APIs verfügen über die gleiche Kernfunktionalität wie die jetzt standardmäßigen, synchronen Office Scripts-APIs, erfordern jedoch, dass Ihr Skript die Datensynchronisierung zwischen dem Skript und der Arbeitsmappe steuert.

> [!IMPORTANT]
> Das async-Modell kann nur mit Skripts verwendet werden, die vor der Implementierung des aktuellen [API-Modells](scripting-fundamentals.md)erstellt wurden. Skripts sind bei der Erstellung dauerhaft an das API-Modell gesperrt, das sie haben. Dies bedeutet auch, dass Sie ein brandneues Skript erstellen müssen, wenn Sie ein altes Skript in das neue Modell konvertieren möchten. Es wird empfohlen, dass Sie Ihre alten Skripts auf das neue Modell aktualisieren, wenn Sie Änderungen vornehmen, da das aktuelle Modell einfacher zu verwenden ist. Der Abschnitt [Konvertieren von Asynchronskripten in den aktuellen Modellabschnitt](#convert-async-scripts-to-the-current-model) enthält Hinweise zum Umwandeln dieses Übergangs.

## <a name="older-main-function-signature"></a>Ältere `main` Funktionssignatur

Skripts, die die async-APIs verwenden, haben eine andere `main` Funktion. Es ist eine `async` Funktion, die einen `Excel.RequestContext` als ersten Parameter hat.

```TypeScript
async function main(context: Excel.RequestContext) {
    // Your async Office Script
}
```

## <a name="context"></a>Context

Die `main`-Funktion lässt einen `Excel.RequestContext`-Parameter namens `context` zu. Stellen Sie sich `context` als die Brücke zwischen Ihrem Skript und der Arbeitsmappe vor. Das Skript greift auf die Arbeitsmappe mit dem `context`-Objekt zu und verwendet diesen `context` zum Hin- und Hersenden von Daten.

Das `context`-Objekt ist erforderlich, weil das Skript und Excel in unterschiedlichen Prozessen und Speicherorten ausgeführt werden. Das Skript muss Änderungen an den Daten in der Arbeitsmappe in der Cloud vornehmen oder diese abrufen können. Das `context`-Objekt verwaltet diese Transaktionen.

## <a name="sync-and-load"></a>Synchronisieren und Laden

Da Ihr Skript und die Arbeitsmappe an unterschiedlichen Orten ausgeführt werden, dauert die Datenübertragung zwischen diesen etwas. In der async-API werden Befehle in die Warteschlange gestellt, bis das Skript den `sync` Vorgang zum Synchronisieren des Skripts und der Arbeitsmappe explizit aufruft. Ihr Skript kann unabhängig funktionieren, bis es eine der folgenden Aktionen durchführen muss:

- Daten aus der Arbeitsmappe lesen (nach einem `load`-Vorgang oder einer Methode, die ein[ClientResultat](/javascript/api/office/officeextension.clientresult?view=excel-js-online&preserve-view=true) zurückgibt).
- Daten in die Arbeitsmappe schreiben (in der Regel, weil das Skript abgeschlossen wurde).

In der folgenden Abbildung wird ein Beispiel für eine Ablaufsteuerung zwischen dem Skript und der Arbeitsmappe dargestellt:

:::image type="content" source="../images/load-sync.png" alt-text="Ein Diagramm, in dem Lese- und Schreibvorgänge angezeigt werden, die aus dem Skript an die Arbeitsmappe":::

### <a name="sync"></a>Synchronisierung

Wenn Ihr async-Skript Daten aus der Arbeitsmappe lesen oder in die Arbeitsmappe schreiben muss, rufen Sie die Methode auf, `RequestContext.sync` wie im folgenden Codeausschnitt gezeigt:

```TypeScript
await context.sync();
```

> [!NOTE]
> `context.sync()` wird implizit aufgerufen, wenn ein Skript endet.

Nachdem der `sync`-Vorgang abgeschlossen ist, wird die Arbeitsmappe entsprechend den Schreibvorgängen aktualisiert, die vom Skript angegeben wurden. Ein Schreibvorgang setzt eine beliebige Eigenschaft für ein Excel Objekt (z. B. ) oder eine Methode auf, die `range.format.fill.color = "red"` eine Eigenschaft ändert (z. B. `range.format.autoFitColumns()` ). Der `sync`-Vorgang liest auch alle Werte aus der Arbeitsmappe, die das Skript angefordert hat, indem es einen `load`-Vorgang oder eine Methode verwendet, die ein `ClientResult` zurückgibt (wie in den nächsten Abschnitten besprochen).

Je nach Netzwerk kann es einige Zeit dauern, bis das Skript mit der Arbeitsmappe synchronisiert wurde. Minimieren Sie die Anzahl der `sync` Aufrufe, damit Ihr Skript schnell ausgeführt wird. Andernfalls sind die asynchronen APIs nicht schneller als die standardmäßigen, synchronen APIs.

### <a name="load"></a>Laden

Ein async-Skript muss Daten aus der Arbeitsmappe laden, bevor es gelesen wird. Das Laden von Daten aus der gesamten Arbeitsmappe würde jedoch die Geschwindigkeit des Skripts erheblich reduzieren. Mit der `load` Methode kann das Skript genau angeben, welche Daten aus der Arbeitsmappe abgerufen werden sollen.

Die `load`-Methode ist für jedes Excel-Objekt verfügbar. Ihr Skript muss die Eigenschaften eines Objekts laden, bevor es sie lesen kann. Wenn Sie dies nicht tun, führt dies zu einem Fehler.

In den folgenden Beispielen wird ein `Range`-Objekt verwendet, um die drei Arten darzustellen, wie die `load`-Methode zum Laden von Daten verwendet werden kann.

|Absicht |Beispielbefehl | Auswirkung |
|:--|:--|:--|
|Laden einer Eigenschaft |`myRange.load("values");` | Lädt eine einzelne Eigenschaft, in diesem Fall den zweidimensionalen Wertearray in diesem Bereich. |
|Laden mehrerer Eigenschaften |`myRange.load("values, rowCount, columnCount");`| Lädt alle Eigenschaften aus einer durch Kommas getrennten Liste, in diesem Beispiel die Werte, die Zeilenanzahl und die Spaltenanzahl. |
|Alles laden | `myRange.load();`|Lädt alle Eigenschaften des Zellbereichs. Dies ist keine empfohlene Lösung, da es Ihr Skript verlangsamen wird, indem unnötige Daten erhalten. Verwenden Sie dies nur, während Sie Ihr Skript testen oder wenn Sie jede Eigenschaft aus dem Objekt benötigen. |

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

Sie können auch Eigenschaften aus einer ganzen Sammlung laden. Jedes Auflistungsobjekt in der async-API verfügt über eine `items` Eigenschaft, die ein Array ist, das die Objekte in dieser Auflistung enthält. Durch die Verwendung von `items` als Anfang eines hierarchischen Aufrufs (`items\myProperty`) für `load` werden die angegebenen Eigenschaften für jedes dieser Elemente geladen. Im folgenden Beispiel wird die `resolved`-Eigenschaft für jedes `Comment`-Objekt im `CommentCollection`-Objekt eines Arbeitsblatts geladen.

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

## <a name="convert-async-scripts-to-the-current-model"></a>Konvertieren von Async-Skripts in das aktuelle Modell

Das aktuelle API-Modell verwendet keine `load` , `sync` oder eine `RequestContext` . Dies macht die Skripte viel einfacher zu schreiben und zu verwalten. Ihre beste Ressource zum Konvertieren alter Skripts ist [Microsoft Q&A](/answers/topics/office-scripts-dev.html). Dort können Sie die Community um Hilfe bei bestimmten Szenarien bitten. In den folgenden Anleitungen sollten Sie die allgemeinen Schritte erläutern, die Sie ausführen müssen.

1. Erstellen Sie ein neues Skript, und kopieren Sie den alten asynchronen Code in dieses. Achten Sie darauf, die alte Methodensignatur nicht einzuschließen, `main` sondern stattdessen die aktuelle. `function main(workbook: ExcelScript.Workbook)`

2. Entfernen Sie alle `load` und `sync` Anrufe. Sie sind nicht mehr notwendig.

3. Alle Eigenschaften wurden entfernt. Sie greifen nun über diese Objekte und Methoden auf diese Objekte `get` `set` zu, daher müssen Sie diese Eigenschaftsverweise auf Methodenaufrufe umstellen. Anstatt beispielsweise die Füllfarbe einer Zelle durch den Eigenschaftenzugriff wie folgt `mySheet.getRange("A2:C2").format.fill.color = "blue";` festzulegen: verwenden Sie jetzt Methoden wie die folgende: `mySheet.getRange("A2:C2").getFormat().getFill().setColor("blue");`

4. Auflistungsklassen wurden durch Arrays ersetzt. Die `add` und Methoden dieser `get` Auflistungsklassen wurden in das Objekt verschoben, das die Auflistung besaß, daher müssen Ihre Verweise entsprechend aktualisiert werden. Verwenden Sie beispielsweise den folgenden Code, um ein Diagramm mit dem Namen "MyChart" aus dem ersten Arbeitsblatt in der Arbeitsmappe abzubekommen: `workbook.getWorksheets()[0].getChart("MyChart");` . Beachten Sie `[0]` die, um auf den ersten Wert des `Worksheet[]` zurückgegebenen von `getWorksheets()` zuzugreifen.

5. Einige Methoden wurden aus Gründen der Übersichtlichkeit umbenannt und aus Gründen der Benutzerfreundlichkeit hinzugefügt. Weitere Informationen finden Sie in der [API-Referenz Office Scripts.](/javascript/api/office-scripts/overview)

## <a name="office-scripts-async-api-reference-documentation"></a>Office Skripts async API Referenzdokumentation

Die asynchronen APIs entsprechen denen, die in Office Add-Ins verwendet werden. Die Referenzdokumentation finden Sie im [Excel Abschnitt der Office JavaScript-API-Referenz .](/javascript/api/excel?view=excel-js-online&preserve-view=true)
