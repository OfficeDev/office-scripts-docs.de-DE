---
title: Unterstützung älterer Office-Skripts, die die asynchronen APIs verwenden
description: Einführung in die Office Scripts Async-APIs und die Verwendung des Lade-/Synchronisierungsmusters für ältere Skripts.
ms.date: 02/08/2021
localization_priority: Normal
ms.openlocfilehash: be7847efe59dc6026875b8a8e3b3c93e0eb82e4d
ms.sourcegitcommit: 345f1dd96d80471b246044b199fe11126a192a88
ms.translationtype: MT
ms.contentlocale: de-DE
ms.lasthandoff: 02/14/2021
ms.locfileid: "50242025"
---
# <a name="support-older-office-scripts-that-use-the-async-apis"></a>Unterstützung älterer Office-Skripts, die die asynchronen APIs verwenden

In diesem Artikel erfahren Sie, wie Sie Skripts verwalten und aktualisieren, die die asynchronen APIs des älteren Modells verwenden. Diese APIs verfügen über die gleiche Kernfunktionalität wie die jetzt standardmäßigen, synchronen Office-Skript-APIs, aber sie erfordern ihr Skript, um die Datensynchronisierung zwischen dem Skript und der Arbeitsmappe zu steuern.

> [!IMPORTANT]
> Das asynchrone Modell kann nur mit Skripts verwendet werden, die vor der Implementierung des aktuellen [API-Modells erstellt wurden.](scripting-fundamentals.md?view=office-scripts&preserve-view=true) Skripts sind dauerhaft für das API-Modell gesperrt, das sie bei der Erstellung haben. Dies bedeutet auch, dass Sie ein völlig neues Skript erstellen müssen, wenn Sie ein altes Skript in das neue Modell konvertieren möchten. Es wird empfohlen, die alten Skripts beim Vornehmen von Änderungen auf das neue Modell zu aktualisieren, da das aktuelle Modell einfacher zu verwenden ist. Der [Abschnitt "Konvertieren asynchroner Skripts](#converting-async-scripts-to-the-current-model) in den aktuellen Modellabschnitt" enthält Hinweise zum Übergang.

## <a name="main-function"></a>Die `main`-Funktion

Skripts, die die asynchronen APIs verwenden, haben eine andere `main` Funktion. Es ist eine `async` Funktion, die als `Excel.RequestContext` erster Parameter verwendet wird.

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

Wenn Ihr asynchrones Skript Daten aus der Arbeitsmappe lesen oder in die Arbeitsmappe schreiben muss, rufen Sie die Methode wie `RequestContext.sync` hier gezeigt auf:

```TypeScript
await context.sync();
```

> [!NOTE]
> `context.sync()` wird implizit aufgerufen, wenn ein Skript endet.

Nachdem der `sync`-Vorgang abgeschlossen ist, wird die Arbeitsmappe entsprechend den Schreibvorgängen aktualisiert, die vom Skript angegeben wurden. Ein Schreibvorgang ist das Festlegen einer beliebigen Eigenschaft für ein Excel-Objekt (z. B. ) oder das Aufrufen einer Methode, die eine Eigenschaft ändert `range.format.fill.color = "red"` (z. B. `range.format.autoFitColumns()` ). Der `sync`-Vorgang liest auch alle Werte aus der Arbeitsmappe, die das Skript angefordert hat, indem es einen `load`-Vorgang oder eine Methode verwendet, die ein `ClientResult` zurückgibt (wie in den nächsten Abschnitten besprochen).

Je nach Netzwerk kann es einige Zeit dauern, bis das Skript mit der Arbeitsmappe synchronisiert wurde. Minimieren Sie die Anzahl der `sync` Aufrufe, damit Ihr Skript schnell ausgeführt werden kann. Andernfalls sind die asynchronen APIs nicht schneller als die standardmäßigen synchronen APIs.

### <a name="load"></a>Laden

Ein asynchrones Skript muss Daten aus der Arbeitsmappe laden, bevor es gelesen wird. Das Laden von Daten aus der gesamten Arbeitsmappe würde jedoch die Geschwindigkeit des Skripts erheblich reduzieren. Mit der Methode kann Ihr Skript explizit die Daten aus der Arbeitsmappe `load` abrufen.

Die `load`-Methode ist für jedes Excel-Objekt verfügbar. Ihr Skript muss die Eigenschaften eines Objekts laden, bevor es sie lesen kann. Wenn Sie dies nicht tun, tritt ein Fehler auf.

In den folgenden Beispielen wird ein `Range`-Objekt verwendet, um die drei Arten darzustellen, wie die `load`-Methode zum Laden von Daten verwendet werden kann.

|Absicht |Beispielbefehl | Auswirkung |
|:--|:--|:--|
|Laden einer Eigenschaft |`myRange.load("values");` | Lädt eine einzelne Eigenschaft, in diesem Fall den zweidimensionalen Wertearray in diesem Bereich. |
|Laden mehrerer Eigenschaften |`myRange.load("values, rowCount, columnCount");`| Lädt alle Eigenschaften aus einer durch Kommas getrennten Liste, in diesem Beispiel die Werte, die Zeilenanzahl und die Spaltenanzahl. |
|Alles laden | `myRange.load();`|Lädt alle Eigenschaften des Zellbereichs. Dies ist keine empfohlene Lösung, da das Skript dadurch verlangsamt wird, dass unnötige Daten erhalten werden. Verwenden Sie dies nur beim Testen Des Skripts oder wenn Sie jede Eigenschaft aus dem Objekt benötigen. |

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

Methoden in der asynchronen API, die Informationen aus der Arbeitsmappe zurückgeben, haben ein ähnliches Muster wie das `load` / `sync` Paradigma. `TableCollection.getCount` ruft zum Beispiel die Anzahl von Tabellen in der Auflistung ab. `getCount` gibt eine `ClientResult<number>` zurück, `value` d. h. die Eigenschaft in der zurückgegebenen [`ClientResult`](/javascript/api/office/officeextension.clientresult?view=excel-js-online&preserve-view=true) ist eine Zahl. Ihr Skript kann erst auf diesen Wert zugreifen, wenn `context.sync()` aufgerufen wird. Ähnlich wie beim Laden einer Eigenschaft ist der `value` bis zu diesem `sync`-Aufruf ein lokaler "leerer" Wert.

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

## <a name="converting-async-scripts-to-the-current-model"></a>Konvertieren asynchroner Skripts in das aktuelle Modell

Das aktuelle API-Modell verwendet keine `load` , `sync` oder eine `RequestContext` . Dies erleichtert das Schreiben und Warten der Skripts erheblich. Die beste Ressource für die Konvertierung alter Skripts ist [Stack Overflow](https://stackoverflow.com/questions/tagged/office-scripts). Dort können Sie die Community um Hilfe bei bestimmten Szenarien bitten. Die folgenden Anleitungen sollten Ihnen dabei helfen, die allgemeinen Schritte zu beschreiben, die Sie ausführen müssen.

1. Erstellen Sie ein neues Skript, und kopieren Sie den alten asynchronen Code in das Skript. Achten Sie darauf, die alte Methodensignatur `main` nicht mit der aktuellen zu `function main(workbook: ExcelScript.Workbook)` verwenden.

2. Entfernen Sie alle `load` und `sync` alle Aufrufe. Sie sind nicht mehr erforderlich.

3. Alle Eigenschaften wurden entfernt. Sie greifen nun über und methoden auf diese Objekte zu, sodass Sie diese Eigenschaftsverweise auf `get` `set` Methodenaufrufe umstellen müssen. Statt beispielsweise die Füllfarbe einer Zelle über den Eigenschaftszugriff wie die folgende festlegen zu müssen, verwenden Sie jetzt `mySheet.getRange("A2:C2").format.fill.color = "blue";` Methoden wie die folgende: `mySheet.getRange("A2:C2").getFormat().getFill().setColor("blue");`

4. Sammlungsklassen wurden durch Arrays ersetzt. Die `add` und die Methoden dieser Auflistungsklassen wurden in das Objekt verschoben, das die Auflistung besaß, sodass Ihre Verweise `get` entsprechend aktualisiert werden müssen. Verwenden Sie beispielsweise den folgenden Code, um ein Diagramm mit dem Namen "MyChart" aus dem ersten Arbeitsblatt in der Arbeitsmappe zu `workbook.getWorksheets()[0].getChart("MyChart");` erhalten: Note the `[0]` to access the first value of the `Worksheet[]` returned by `getWorksheets()` .

5. Einige Methoden wurden aus Gründen der Übersichtlichkeit umbenannt und zur Vereinfachung hinzugefügt. Weitere Informationen finden Sie in der [Office Scripts-API-Referenz.](/javascript/api/office-scripts/overview?view=office-scripts&preserve-view=true)

## <a name="office-scripts-async-api-reference-documentation"></a>Referenzdokumentation zur asynchronen API für Office-Skripts

Die asynchronen APIs entsprechen denen in Office-Add-Ins. Die Referenzdokumentation finden Sie im [Abschnitt "Excel" der JavaScript-API-Referenz für Office-Add-Ins.](/javascript/api/excel?view=excel-js-online&preserve-view=true)
