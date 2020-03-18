---
title: Lesen Sie Arbeitsmappendaten mit Office-Skripts in Excel im Web
description: Ein Office Skripts-Lernprogramm zum Lesen von Daten aus Arbeitsmappen und zum Auswerten dieser Daten im Skript.
ms.date: 01/27/2020
localization_priority: Priority
ms.openlocfilehash: 42ed0fe5843a78692f9660b873211e3668702164
ms.sourcegitcommit: b075eed5a6f275274fbbf6d62633219eac416f26
ms.translationtype: HT
ms.contentlocale: de-DE
ms.lasthandoff: 03/10/2020
ms.locfileid: "42700182"
---
# <a name="read-workbook-data-with-office-scripts-in-excel-on-the-web"></a>Lesen Sie Arbeitsmappendaten mit Office-Skripts in Excel im Web

In diesem Lernprogramm erfahren Sie, wie Sie Daten aus einer Arbeitsmappe mit einem Office-Skript für Excel im Web lesen. Anschließend bearbeiten Sie die gelesenen Daten und fügen sie wieder in die Arbeitsmappe ein.

> [!TIP]
> Wenn Sie mit Office-Skripten noch nicht vertraut sind, empfehlen wir, mit dem [Aufzeichnen, Bearbeiten und Erstellen von Office-Skripten in Excel im Web](excel-tutorial.md)-Lernprogramm zu beginnen.

## <a name="prerequisites"></a>Voraussetzungen

[!INCLUDE [Preview note](../includes/preview-note.md)]

Bevor Sie mit diesem Lernprogramm beginnen, benötigen Sie Zugriff auf Office-Skripts. Dies setzt Folgendes voraus:

- [Excel im Web](https://www.office.com/launch/excel).
- Bitten Sie Ihren Administrator, [Office-Skripts für Ihre Organisation zu aktivieren](https://support.office.com/article/office-scripts-settings-in-m365-19d3c51a-6ca2-40ab-978d-60fa49554dcf). Dadurch wird die Registerkarte **Automatisieren** zum Menüband hinzugefügt.

> [!IMPORTANT]
> Dieses Lernprogramm richtet sich an Anfänger bis Fortgeschrittene mit JavaScript oder TypeScript. Wenn Sie noch nicht mit JavaScript vertraut sind, empfehlen wir Ihnen, das [Mozilla-JavaScript-Lernprogramm](https://developer.mozilla.org/docs/Web/JavaScript/Guide/Introduction) durchzusehen. Weitere Informationen über die Skriptumgebung finden Sie unter [Office-Skripts in Excel im Web](../overview/excel.md).

## <a name="read-a-cell"></a>Lesen einer Zelle

Mit dem Action Recorder erstellte Skripte können nur Informationen in die Arbeitsmappe schreiben. Mit dem Code-Editor können Sie Skripte bearbeiten und erstellen, die auch Daten aus einer Arbeitsmappe lesen.

Lassen Sie uns ein Skript erstellen, das Daten liest und basierend auf dem Gelesenen agiert. Wir werden mit einem Muster-Kontoauszug arbeiten. Diese Erklärung ist ein kombinierter Prüfungs- und Gutschrift-Kontoauszug. Leider melden sie Bilanzänderungen unterschiedlich. Der Prüfungskontoauszug gibt die Einnahmen als positive Gutschrift und die Kosten als negative Belastung an. Der Gutschrift-Kontoauszug macht das Gegenteil.

Im weiteren Verlauf des Lernprogramms werden diese Daten mithilfe eines Skripts normalisiert. Lassen Sie uns zunächst lernen, wie Sie Daten aus der Arbeitsmappe lesen.

1. Erstellen Sie ein neues Arbeitsblatt in der Arbeitsmappe, die Sie für den Rest des Lernprogramms verwendet haben.
2. Kopieren Sie die folgenden Daten und fügen Sie sie ab Zelle **A1** in das neue Arbeitsblatt ein.

    |Datum |Konto |Beschreibung |Lastschrift |Guthaben |
    |:--|:--|:--|:--|:--|
    |10.10.2019 |Wird geprüft |Coho Vineyard |-20.05 | |
    |11.10.2019 |Guthaben |The Phone Company |99.95 | |
    |13.10.2019 |Guthaben |Coho Vineyard |154.43 | |
    |15.10.2019 |Wird geprüft |Externe Einzahlung | |1000 |
    |20.10.2019 |Guthaben |Coho Vineyard – Rückerstattung | |-35.45 |
    |25.10.2019 |Wird geprüft |Best For You Organics Company | -85.64 | |
    |01.11.2019 |Wird geprüft |Externe Einzahlung | |1000 |

3. Öffnen Sie den **Code-Editor**, und wählen Sie **Neuer Skript**aus.
4. Lassen Sie uns die Formatierung zurechtmachen. Dies ist ein Finanzdokument. Ändern Sie daher die Zahlenformatierung in den Spalten **Lastschrift** und **Gutschrift**, um Werte als Dollarbeträge anzuzeigen. Passen wir auch die Spaltenbreite an die Daten an.

    Ersetzen Sie den Skriptinhalt durch den folgenden Code:

    ```TypeScript
    async function main(context: Excel.RequestContext) {
      // Get the current worksheet.
      let workbook = context.workbook;
      let worksheets = workbook.worksheets;
      let selectedSheet = worksheets.getActiveWorksheet();

      // Format the range to display numerical dollar amounts.
      selectedSheet.getRange("D2:E8").numberFormat = [["$#,##0.00"]];

      // Fit the width of all the used columns to the data.
      selectedSheet.getUsedRange().format.autofitColumns();
    }
    ```

5. Lesen wir nun einen Wert aus einer der Zahlenspalten. Fügen Sie am Ende des Skripts den folgenden Code hinzu:

    ```TypeScript
    // Get the value of cell D2.
    let range = selectedSheet.getRange("D2");
    range.load("values");
    await context.sync();
  
    // Print the value of D2.
    console.log(range.values);
    ```

    Beachten Sie die Anrufe zu `load` und `sync`. Einzelheiten zu diesen Methoden finden Sie unter Grundlagen der [Skripterstellung für Office-Skripts in Excel im Web](../develop/scripting-fundamentals.md#sync-and-load). Beachten Sie zunächst, dass Sie das Lesen von Daten anfordern und dann Ihr Skript mit der Arbeitsmappe synchronisieren müssen, um diese Daten zu lesen.

6. Führen Sie das Skript aus.
7. Öffnen Sie die Konsole. Wechseln Sie zum Menü **Ellipsen** und drücken Sie **Protokolle...**.
8. Sie sollten `[Array[1]]` in der Konsole sehen. Dies ist keine Zahl, da Bereiche zweidimensionale Datenfelder sind. Dieser zweidimensionale Bereich wird direkt in der Konsole protokolliert. Glücklicherweise können Sie mit dem Code-Editor den Inhalt des Arrays anzeigen.
9. Wenn ein zweidimensionales Array in der Konsole protokolliert wird, werden Spaltenwerte unter jeder Zeile gruppiert. Erweitern Sie das Array-Protokoll, indem Sie auf das blaue Dreieck klicken.
10. Erweitern Sie die zweite Ebene des Arrays, indem Sie auf das neu aufgedeckte blaue Dreieck klicken. Sie sollten jetzt Folgendes sehen:

    ![Das Konsolenprotokoll mit der Ausgabe "-20.05", verschachtelt unter zwei Arrays.](../images/tutorial-4.png)

## <a name="modify-the-value-of-a-cell"></a>Ändern Sie den Wert einer Zelle

Nachdem wir nun Daten lesen können, verwenden wir diese Daten, um die Arbeitsmappe zu ändern. Wir werden den Wert der Zelle **D2** mit der `Math.abs` Funktion positiv machen. Das [Math](https://developer.mozilla.org/docs/web/javascript/reference/global_objects/math)-Objekt enthält viele Funktionen, auf die Ihre Skripte Zugriff haben. Weitere Informationen zu `Math` und andere integrierte Objekte finden Sie unter [Verwenden von integrierten JavaScript-Objekten in Office-Skripts](../develop/javascript-objects.md).

1. Fügen Sie am Ende des Skripts den folgenden Code hinzu:

    ```TypeScript
    // Run the `Math.abs` function with the value at D2 and apply that value back to D2.
    let positiveValue = Math.abs(range.values[0][0]);
    range.values = [[positiveValue]];
    ```

2. Der Wert von Zelle **D2** sollte jetzt positiv sein.

## <a name="modify-the-values-of-a-column"></a>Ändern Sie die Werte einer Spalte

Nachdem wir nun wissen, wie man in eine einzelne Zelle liest und schreibt, verallgemeinern wir das Skript so, dass es für die gesamten Spalten **Lastschrift** und **Gutschrift** funktioniert.

1. Entfernen Sie den Code, der nur eine einzelne Zelle betrifft (den vorherigen Absolutwertcode), sodass Ihr Skript jetzt folgendermaßen aussieht:

    ```TypeScript
    async function main(context: Excel.RequestContext) {
      // Get the current worksheet.
      let workbook = context.workbook;
      let worksheets = workbook.worksheets;
      let selectedSheet = worksheets.getActiveWorksheet();

      // Format the range to display numerical dollar amounts.
      selectedSheet.getRange("D2:E8").numberFormat = [["$#,##0.00"]];

      // Fit the width of all the used columns to the data.
      selectedSheet.getUsedRange().format.autofitColumns();
    }
    ```

2. Fügen Sie eine Schleife hinzu, die die Zeilen in den letzten beiden Spalten durchläuft. Für jede Zelle setzt das Skript den Wert auf den absoluten Wert des aktuellen Werts.

    Beachten Sie, dass das Array, das die Zellenpositionen definiert, auf Null basiert. Das heißt, Zelle **A1** ist `range[0][0]`.

    ```TypeScript
    // Get the values of the used range.
    let range = selectedSheet.getUsedRange();
    range.load("rowCount,values");
    await context.sync();

    // Iterate over the fourth and fifth columns and set their values to their absolute value.
    for (let i = 1; i < range.rowCount; i++) {
      // The column at index 3 is column "4" in the worksheet.
      if (range.values[i][3] != 0) {
        let positiveValue = Math.abs(range.values[i][3]);
        selectedSheet.getCell(i, 3).values = [[positiveValue]];
      }

      // The column at index 4 is column "5" in the worksheet.
      if (range.values[i][4] != 0) {
        let positiveValue = Math.abs(range.values[i][4]);
        selectedSheet.getCell(i, 4).values = [[positiveValue]];
      }
    }
    ```

    Dieser Teil des Skripts führt mehrere wichtige Aufgaben aus. Zunächst werden die Werte und die Zeilenanzahl des verwendeten Bereichs geladen. Auf diese Weise können wir uns die Werte ansehen und wissen, wann wir aufhören müssen. Zweitens durchläuft es den verwendeten Bereich und überprüft jede Zelle in den Spalten **Lastschrift** und **Gutschrift**. Wenn der Wert in der Zelle nicht 0 ist, wird er durch seinen absoluten Wert ersetzt. Wir vermeiden Nullen, damit wir die leeren Zellen so lassen können, wie sie waren.

3. Führen Sie das Skript aus.

    Ihr Kontoauszug sollte nun folgendermaßen aussehen:

    ![Der Kontoauszug als formatierte Tabelle mit nur positiven Werten.](../images/tutorial-5.png)

## <a name="next-steps"></a>Nächste Schritte

Öffnen Sie den Code-Editor, und probieren Sie einige unserer [Beispielskripts für Office-Skripts in Excel im Web](../resources/excel-samples.md)aus. Sie können auch [Skripting-Grundlagen für Office-Skripts in Excel im Web](../develop/scripting-fundamentals.md) aufrufen, um weitere Informationen zum Erstellen von Office-Skripts zu erhalten.
