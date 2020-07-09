---
title: Ausführen von Office-Skripts mit Power Automation
description: Vorgehensweise Abrufen von Office-Skripts für Excel im Internet arbeiten mit einem Power automatisieren Workflow.
ms.date: 07/01/2020
localization_priority: Normal
ms.openlocfilehash: 40a67f3d0e8f049a8ec5516c0af54c5fc6fb9319
ms.sourcegitcommit: edf58aed3cd38f57e5e7227465a1ef5515e15703
ms.translationtype: MT
ms.contentlocale: de-DE
ms.lasthandoff: 07/08/2020
ms.locfileid: "45081593"
---
# <a name="run-office-scripts-with-power-automate"></a>Ausführen von Office-Skripts mit Power Automation

[Power Automation](https://flow.microsoft.com) ermöglicht Ihnen das Hinzufügen von Office-Skripts zu einem größeren, automatisierten Workflow. Sie können Power automatisieren verwenden Sie Dinge wie das Hinzufügen des Inhalts einer e-Mail zur Tabelle eines Arbeitsblatts oder das Erstellen von Aktionen in ihren Projektverwaltungstools basierend auf Arbeitsmappen-Kommentaren. Wenn Sie neu bei Power Automation sind, empfehlen wir Ihnen, die [Erste Schritte mit Power automatisieren](/power-automate/getting-started)zu besuchen. Hier erfahren Sie mehr über die Automatisierung Ihrer Workflows in mehreren Diensten.

> [!IMPORTANT]
> Derzeit können Sie Office-Skripts nicht über einen [freigegebenen Fluss](/power-automate/share-buttons)ausführen. Nur der Benutzer, der ein Skript erstellt hat, kann es auch mithilfe von Power Automation ausführen.

## <a name="getting-started"></a>Erste Schritte

Um mit der Kombination von Power Automation und Office-Skripts zu beginnen, führen Sie das Lernprogramm [Start using scripts with Power Automation aus](../tutorials/excel-power-automate-manual.md). In diesem Artikel erfahren Sie, wie Sie einen Fluss erstellen, der ein einfaches Skript aufruft. Wenn Sie das Lernprogramm und das Lernprogramm [automatisch ausgeführte Skripts mit automatischer Power-Automatisierung ausführen](../tutorials/excel-power-automate-trigger.md) abgeschlossen haben, geben Sie hier ausführliche Informationen zum Verbinden von Office-Skripts mit Power Automation Flows ein.

## <a name="excel-online-business-connector"></a>Excel Online-Connector (Business)

[Connectors](/connectors/connectors) sind die Brücken zwischen Power Automation und Anwendungen. Der [Excel Online (Business)-Connector](/connectors/excelonlinebusiness) gibt Ihrem Fluss Zugriff auf Excel-Arbeitsmappen. Mit der Aktion "Skript ausführen" können Sie ein beliebiges Office-Skript aufrufen, das über die ausgewählte Arbeitsmappe zugänglich ist. Sie können nicht nur Skripts über einen Fluss ausführen, sondern auch Daten an und von der Arbeitsmappe mit dem Fluss durch die Skripts übergeben.

> [!IMPORTANT]
> Die Aktion "Skript ausführen" gibt Benutzern, die den Excel Connector verwenden, wichtigen Zugriff auf Ihre Arbeitsmappe und deren Daten. Darüber hinaus gibt es Sicherheitsrisiken mit Skripts, die externe API-Aufrufe durchführen, wie in [externe Aufrufe von Power Automation](external-calls.md)erläutert. Wenn Ihr Administrator mit der Exposition hoch vertraulicher Daten befasst ist, können Sie entweder den Excel Online Connector deaktivieren oder den Zugriff auf Office-Skripts über die [Office Scripts-Administrator Steuerelemente](https://support.microsoft.com/office/19d3c51a-6ca2-40ab-978d-60fa49554dcf)einschränken.

## <a name="data-transfer-in-flows-for-scripts"></a>Datenübertragung in Flows für Skripts

Mit Power Automation können Sie Datenteile zwischen den einzelnen Schritten Ihres Flows übergeben. Skripts können so konfiguriert werden, dass alle Arten von Informationen akzeptiert werden, die Sie benötigen, und Sie geben alles aus Ihrer Arbeitsmappe zurück, die Sie in Ihrem Flow wünschen. Eingabe für Ihr Skript wird durch Hinzufügen von Parametern zur `main` Funktion (zusätzlich zu `workbook: ExcelScript.Workbook` ) angegeben. Die Ausgabe des Skripts wird durch Hinzufügen eines Rückgabetyps zu deklariert `main` .

> [!NOTE]
> Wenn Sie einen Block "Skript ausführen" im Flow erstellen, werden die akzeptierten Parameter und die zurückgegebenen Typen aufgefüllt. Wenn Sie die Parameter oder Rückgabetypen Ihres Skripts ändern, müssen Sie den Block "Skript ausführen" des Flusses wiederholen. Dadurch wird sichergestellt, dass die Daten ordnungsgemäß analysiert werden.

In den folgenden Abschnitten werden die Details der Eingabe und Ausgabe für Skripts behandelt, die in Power Automation verwendet werden. Wenn Sie eine praktische Herangehensweise zum Erlernen dieses Themas wünschen, probieren Sie das Lernprogramm [automatisch ausgeführte Skripts mit automatischer Power-Automatisierung](../tutorials/excel-power-automate-trigger.md) aus, oder erkunden Sie das Beispielszenario für [automatisierte Aufgaben Erinnerungen](../resources/scenarios/task-reminders.md) .

### <a name="main-parameters-passing-data-to-a-script"></a>`main`Parameter: übergeben von Daten an ein Skript

Alle Skript Eingaben werden als zusätzliche Parameter für die `main` Funktion angegeben. Wenn Sie beispielsweise möchten, dass ein Skript einen akzeptiert, `string` das einen Namen als Eingabe darstellt, ändern Sie die `main` Signatur in `function main(workbook: ExcelScript.Workbook, name: string)` .

Wenn Sie einen Fluss in Power Automation konfigurieren, können Sie Skript Eingaben als statische Werte, [Ausdrücke](/power-automate/use-expressions-in-conditions)oder dynamische Inhalte angeben. Details zu den Connectors eines einzelnen Diensts finden Sie in der [Power Automation Connector-Dokumentation](/connectors/).

Berücksichtigen Sie beim Hinzufügen von Eingabeparametern zur Funktion eines Skripts `main` die folgenden Zulagen und Einschränkungen.

1. Der erste Parameter muss vom Typ sein `ExcelScript.Workbook` . Der Name des Parameters spielt keine Rolle.

2. Jeder Parameter muss einen Typ aufweisen.

3. Die Grundtypen `string` , `number` ,,, `boolean` `any` `unknown` , `object` und `undefined` werden unterstützt.

4. Arrays der zuvor aufgelisteten Grundtypen werden unterstützt.

5. Geschachtelte Arrays werden als Parameter unterstützt (jedoch nicht als Rückgabetypen).

6. Union-Typen sind zulässig, wenn es sich um eine Vereinigung von literalen handelt, die zu einem einzelnen Typ ( `string` , `number` oder `boolean` ) gehören. Gewerkschaften eines unterstützten Typs mit undefined werden ebenfalls unterstützt.

7. Objekttypen sind zulässig, wenn Sie Eigenschaften vom Typ `string` , `number` ,, `boolean` unterstützte Arrays oder andere unterstützte Objekte enthalten. Im folgenden Beispiel werden geschachtelte Objekte gezeigt, die als Parametertypen unterstützt werden:

    ```TypeScript
    // Office Scripts can return an Employee object because Position only contains strings and numbers.
    interface Employee {
        name: string;
        job: Position;
    }

    interface Position {
        id: number;
        title: string;
    }
    ```

8. Objekte müssen Ihre Schnittstelle oder Klassendefinition im Skript definiert haben. Ein Objekt kann auch anonym Inline definiert werden, wie im folgenden Beispiel dargestellt:

    ```TypeScript
    function main(workbook: ExcelScript.Workbook): {name: string, email: string}
    ```

9. Optionale Parameter sind zulässig und können mit dem optionalen Modifizierer als solche gekennzeichnet werden `?` (beispielsweise `function main(workbook: ExcelScript.Workbook, Name?: string)` ).

10. Standardparameterwerte sind zulässig (beispielsweise `async function main(workbook: ExcelScript.Workbook, Name: string = 'Jane Doe')` .

## <a name="returning-data-from-a-script"></a>Zurückgeben von Daten aus einem Skript

Skripts können Daten aus der Arbeitsmappe zurückgeben, die als dynamischer Inhalt in einem Power-Automatisierungs Fluss verwendet werden. Wie bei Eingabeparametern stellt Power Automation einige Einschränkungen für den Rückgabetyp dar.

1. Die Grundtypen `string` , `number` , `boolean` , `void` und `undefined` werden unterstützt.

2. Union-Typen, die als Rückgabetypen verwendet werden, verfolgen dieselben Einschränkungen wie bei der Verwendung als Skriptparameter.

3. Array Typen sind zulässig, wenn Sie vom Typ `string` , `number` oder sind `boolean` . Sie sind auch zulässig, wenn der Typ eine unterstützte Union oder ein unterstützter Literaltyp ist.

4. Als Rückgabetypen verwendete Objekttypen entsprechen denselben Einschränkungen wie bei der Verwendung als Skriptparameter.

5. Die implizite Typisierung wird unterstützt, muss aber denselben Regeln wie ein definierter Typ entsprechen.

## <a name="avoid-using-relative-references"></a>Vermeiden der Verwendung relativer Verweise

Power Automation führt Ihr Skript in der ausgewählten Excel-Arbeitsmappe in Ihrem Auftrag aus. Die Arbeitsmappe wird möglicherweise geschlossen, wenn dies geschieht. Jede API, die den aktuellen Status des Benutzers verwendet, beispielsweise `Workbook.getActiveWorksheet` , schlägt fehl, wenn die Leistung automatisiert ausgeführt wird. Achten Sie beim Entwerfen Ihrer Skripts darauf, absolute Bezüge für Arbeitsblätter und Bereiche zu verwenden.

Die folgenden Funktionen lösen einen Fehler aus und schlagen fehl, wenn Sie aus einem Skript in einem Power automatisieren-Flow aufgerufen werden.

- `Chart.activate`
- `Range.select`
- `Workbook.getActiveCell`
- `Workbook.getActiveChart`
- `Workbook.getActiveChartOrNullObject`
- `Workbook.getActiveSlicer`
- `Workbook.getActiveSlicerOrNullObject`
- `Workbook.getActiveWorksheet`
- `Workbook.getSelectedRange`
- `Workbook.getSelectedRanges`
- `Worksheet.activate`

## <a name="example"></a>Beispiel

Der folgende Screenshot zeigt einen Power-Automatisierungs Fluss, der ausgelöst wird, wenn Ihnen ein [GitHub](https://github.com/) -Problem zugewiesen ist. Der Fluss führt ein Skript aus, mit dem das Problem einer Tabelle in einer Excel-Arbeitsmappe hinzugefügt wird. Wenn es fünf oder mehr Probleme in dieser Tabelle gibt, sendet der Fluss eine e-Mail-Erinnerung.

![Das Beispiel fließt wie im Power automatisieren-Fluss-Editor dargestellt.](../images/power-automate-parameter-return-sample.png)

Die `main` Funktion des Skripts gibt die Problem-ID und den Titel des Problems als Eingabeparameter an, und das Skript gibt die Anzahl der Zeilen in der Ausgabetabelle zurück.

```TypeScript
function main(
  workbook: ExcelScript.Workbook,
  issueId: string,
  issueTitle: string): number {
  // Get the "GitHub" worksheet.
  let worksheet = workbook.getWorksheet("GitHub");

  // Get the first table in this worksheet, which contains the table of GitHub issues.
  let issueTable = worksheet.getTables()[0];

  // Add the issue ID and issue title as a row.
  issueTable.addRow(-1, [issueId, issueTitle]);

  // Return the number of rows in the table, which represents how many issues are assigned to this user.
  return issueTable.getRangeBetweenHeaderAndTotal().getRowCount();
}
```

## <a name="see-also"></a>Siehe auch

- [Ausführen von Office-Skripts in Excel im Internet mit Power Automation](../tutorials/excel-power-automate-manual.md)
- [Automatisches Ausführen von Skripts mit automatisiertem Power-Automatisierungs Fluss](../tutorials/excel-power-automate-trigger.md)
- [Grundlegendes zur Skripterstellung für Office-Skripts in Excel im Web](scripting-fundamentals.md)
- [Erste Schritte mit Power Automate](/power-automate/getting-started)
- [Referenzdokumentation zu Excel Online (Business) Connector](/connectors/excelonlinebusiness/)
