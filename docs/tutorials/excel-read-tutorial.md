---
title: Lesen Sie Arbeitsmappendaten mit Office-Skripts in Excel im Web
description: Ein Office Skripts-Lernprogramm zum Lesen von Daten aus Arbeitsmappen und zum Auswerten dieser Daten im Skript.
ms.date: 01/06/2021
localization_priority: Priority
ms.openlocfilehash: aa05533f0d7cd3b53e4eb616ae3216d672d026f9
ms.sourcegitcommit: 4693c8f79428ec74695328275703af0ba1bfea8f
ms.translationtype: HT
ms.contentlocale: de-DE
ms.lasthandoff: 06/23/2021
ms.locfileid: "53074690"
---
# <a name="read-workbook-data-with-office-scripts-in-excel-on-the-web"></a>Lesen Sie Arbeitsmappendaten mit Office-Skripts in Excel im Web

In diesem Lernprogramm erfahren Sie, wie Sie Daten aus einer Arbeitsmappe mit einem Office-Skript für Excel im Web lesen. Sie werden ein neues Skript schreiben, das einen Kontoauszug formatiert und die Daten in diesem Auszug normalisiert. Im Rahmen der Datenbereinigung liest Ihr Skript Werte aus den Buchungszellen, wendet auf jeden Wert eine einfache Formel an und schreibt die resultierende Antwort in die Arbeitsmappe. Durch das Lesen von Daten aus der Arbeitsmappe können Sie einige Ihrer Entscheidungsfindungsprozesse im Skript automatisieren.

> [!TIP]
> Wenn Sie mit Office-Skripten noch nicht vertraut sind, empfehlen wir, mit dem [Aufzeichnen, Bearbeiten und Erstellen von Office-Skripten in Excel im Web](excel-tutorial.md)-Lernprogramm zu beginnen. [Office-Skripts verwenden TypeScript](../overview/code-editor-environment.md), und dieses Lernprogramm richten sich an Anfänger bis Fortgeschrittene mit JavaScript oder TypeScript. Wenn Sie noch nicht mit JavaScript vertraut sind, empfehlen wir Ihnen, mit dem [Mozilla-JavaScript-Lernprogramm](https://developer.mozilla.org/docs/Web/JavaScript/Guide/Introduction) zu beginnen.

## <a name="prerequisites"></a>Voraussetzungen

[!INCLUDE [Tutorial prerequisites](../includes/tutorial-prerequisites.md)]

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

3. Öffnen Sie **Alle Skripts** und wählen Sie **Neues Skript** aus.
4. Lassen Sie uns die Formatierung zurechtmachen. Dies ist ein Finanzdokument. Ändern Sie daher die Zahlenformatierung in den Spalten **Lastschrift** und **Gutschrift**, um Werte als Dollarbeträge anzuzeigen. Passen wir auch die Spaltenbreite an die Daten an.

    Ersetzen Sie den Skriptinhalt durch den folgenden Code:

    ```TypeScript
    function main(workbook: ExcelScript.Workbook) {
        // Get the current worksheet.
        let selectedSheet = workbook.getActiveWorksheet();

        // Format the range to display numerical dollar amounts.
        selectedSheet.getRange("D2:E8").setNumberFormat("$#,##0.00");

        // Fit the width of all the used columns to the data.
        selectedSheet.getUsedRange().getFormat().autofitColumns();
    }
    ```

5. Lesen wir nun einen Wert aus einer der Zahlenspalten. Fügen Sie den folgenden Code am Ende des Skripts hinzu (vor der schließenden `}`):

    ```TypeScript
    // Get the value of cell D2.
    let range = selectedSheet.getRange("D2");
    console.log(range.getValues());
    ```

6. Führen Sie das Skript aus.
7. Sie sollten `[Array[1]]` in der Konsole sehen. Dies ist keine Zahl, da Bereiche zweidimensionale Datenfelder sind. Dieser zweidimensionale Bereich wird direkt in der Konsole protokolliert. Glücklicherweise können Sie mit dem Code-Editor den Inhalt des Arrays anzeigen.
8. Wenn ein zweidimensionales Array in der Konsole protokolliert wird, werden Spaltenwerte unter jeder Zeile gruppiert. Erweitern Sie das Array-Protokoll, indem Sie auf das blaue Dreieck klicken.
9. Erweitern Sie die zweite Ebene des Arrays, indem Sie auf das neu aufgedeckte blaue Dreieck klicken. Jetzt sollten Sie folgendes sehen:

    :::image type="content" source="../images/tutorial-4.png" alt-text="Das Konsolenprotokoll mit der Ausgabe „-20,05“, verschachtelt unter zwei Arrays.":::

## <a name="modify-the-value-of-a-cell"></a>Ändern Sie den Wert einer Zelle

Nachdem wir nun Daten lesen können, verwenden wir diese Daten, um die Arbeitsmappe zu ändern. Wir werden den Wert der Zelle **D2** mit der `Math.abs` Funktion positiv machen. Das [Math](https://developer.mozilla.org/docs/web/javascript/reference/global_objects/math)-Objekt enthält viele Funktionen, auf die Ihre Skripte Zugriff haben. Weitere Informationen zu `Math` und andere integrierte Objekte finden Sie unter [Verwenden von integrierten JavaScript-Objekten in Office-Skripts](../develop/javascript-objects.md).

1. Sie ändern den Wert der Zelle mithilfe der Methoden `getValue` und `setValue`. Diese Methoden funktionieren bei einer einzelnen Zelle. Wenn Sie mehrere Zellbereiche bearbeiten möchten, verwenden Sie `getValues` und `setValues`. Fügen Sie am Ende des Skripts den folgenden Code hinzu:

    ```TypeScript
    // Run the `Math.abs` function with the value at D2 and apply that value back to D2.
    let positiveValue = Math.abs(range.getValue() as number);
    range.setValue(positiveValue);
    ```

    > [!NOTE]
    > Der zurückgegebene Wert wird unter Verwendung des Schlüsselworts `as` von `range.getValue()`in `number` [umgewandelt](https://www.typescripttutorial.net/typescript-tutorial/type-casting/). Dies ist notwendig, da ein Bereich aus Zeichenfolgen, Zahlen oder booleschen Werten bestehen kann. In diesem Fall wird explizit eine Zahl benötigt.

2. Der Wert von Zelle **D2** sollte jetzt positiv sein.

## <a name="modify-the-values-of-a-column"></a>Ändern Sie die Werte einer Spalte

Nachdem wir nun wissen, wie man in eine einzelne Zelle liest und schreibt, verallgemeinern wir das Skript so, dass es für die gesamten Spalten **Lastschrift** und **Gutschrift** funktioniert.

1. Entfernen Sie den Code, der nur eine einzelne Zelle betrifft (den vorherigen Absolutwertcode), sodass Ihr Skript jetzt folgendermaßen aussieht:

    ```TypeScript
    function main(workbook: ExcelScript.Workbook) {
        // Get the current worksheet.
        let selectedSheet = workbook.getActiveWorksheet();

        // Format the range to display numerical dollar amounts.
        selectedSheet.getRange("D2:E8").setNumberFormat("$#,##0.00");

        // Fit the width of all the used columns to the data.
        selectedSheet.getUsedRange().getFormat().autofitColumns();
    }
    ```

2. Fügen Sie am Ende des Skripts eine Schleife hinzu, die die Zeilen in den letzten beiden Spalten durchläuft. Für jede Zelle setzt das Skript den Wert auf den absoluten Wert des aktuellen Werts.

    Beachten Sie, dass das Array, das die Zellenpositionen definiert, auf Null basiert. Das heißt, Zelle **A1** ist `range[0][0]`.

    ```TypeScript
    // Get the values of the used range.
    let range = selectedSheet.getUsedRange();
    let rangeValues = range.getValues();

    // Iterate over the fourth and fifth columns and set their values to their absolute value.
    let rowCount = range.getRowCount();
    for (let i = 1; i < rowCount; i++) {
        // The column at index 3 is column "4" in the worksheet.
        if (rangeValues[i][3] != 0) {
            let positiveValue = Math.abs(rangeValues[i][3] as number);
            selectedSheet.getCell(i, 3).setValue(positiveValue);
        }

        // The column at index 4 is column "5" in the worksheet.
        if (rangeValues[i][4] != 0) {
            let positiveValue = Math.abs(rangeValues[i][4] as number);
            selectedSheet.getCell(i, 4).setValue(positiveValue);
        }
    }
    ```

    Dieser Teil des Skripts führt mehrere wichtige Aufgaben aus. Zunächst werden die Werte und die Zeilenanzahl des verwendeten Bereichs abgerufen. Auf diese Weise können wir uns die Werte ansehen und wissen, wann wir aufhören müssen. Zweitens durchläuft es den verwendeten Bereich und überprüft jede Zelle in den Spalten **Lastschrift** und **Gutschrift**. Wenn der Wert in der Zelle nicht 0 ist, wird er durch seinen absoluten Wert ersetzt. Wir vermeiden Nullen, damit wir die leeren Zellen so lassen können, wie sie waren.

3. Führen Sie das Skript aus.

    Ihr Kontoauszug sollte nun folgendermaßen aussehen:

    :::image type="content" source="../images/tutorial-5.png" alt-text="Ein Arbeitsblatt, auf dem der Kontoauszug als formatierte Tabelle mit nur positiven Werten angezeigt wird.":::

## <a name="next-steps"></a>Nächste Schritte

Öffnen Sie den Code-Editor, und probieren Sie einige unserer [Beispielskripts für Office-Skripts in Excel im Web](../resources/samples/excel-samples.md)aus. Sie können auch [Skripting-Grundlagen für Office-Skripts in Excel im Web](../develop/scripting-fundamentals.md) aufrufen, um weitere Informationen zum Erstellen von Office-Skripts zu erhalten.

Die nächste Reihe von Office-Skripts-Lernprogrammen konzentriert sich auf die Verwendung von Office-Skripts mit Power Automate. Weitere Informationen über die Vorteile, diese beiden Plattformen miteinander zu kombinieren, finden Sie in [Ausführen von Office-Skripts mit Power Automate](../develop/power-automate-integration.md), oder sehen Sie sich das Lernprogramm [Aufrufen von Skripts aus einem manuellen Power Automate-Datenfluss](excel-power-automate-manual.md) an, um einen Power Automate-Datenfluss zu erstellen, der ein Office-Skript verwendet.
