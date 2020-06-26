---
title: Integrieren von Office-Skripts mit Power Automation
description: Vorgehensweise Abrufen von Office-Skripts für Excel im Internet arbeiten mit einem Power automatisieren Workflow.
ms.date: 06/24/2020
localization_priority: Normal
ms.openlocfilehash: 977d9c88d75c8070eb729a443b4e8bc9a32e456d
ms.sourcegitcommit: aec3c971c6640429f89b6bb99d2c95ea06725599
ms.translationtype: MT
ms.contentlocale: de-DE
ms.lasthandoff: 06/25/2020
ms.locfileid: "44878776"
---
# <a name="integrate-office-scripts-with-power-automate"></a>Integrieren von Office-Skripts mit Power Automation

[Power Automation](https://flow.microsoft.com) integriert Ihr Skript in einen größeren Workflow. Sie können Power automatisieren verwenden Sie Dinge wie das Hinzufügen des Inhalts einer e-Mail zur Tabelle eines Arbeitsblatts oder das Erstellen von Aktionen in ihren Projektverwaltungstools basierend auf Arbeitsmappen-Kommentaren. Wenn Sie neu bei Power Automation sind, empfehlen wir Ihnen, die [Erste Schritte mit Power automatisieren](/power-automate/getting-started)zu besuchen. Hier erfahren Sie mehr über die Automatisierung Ihrer Workflows in mehreren Diensten.

> [!IMPORTANT]
> Derzeit können Sie Office-Skripts nicht über einen [freigegebenen Fluss](/power-automate/share-buttons)ausführen. Nur der Benutzer, der ein Skript erstellt hat, kann es auch mithilfe von Power Automation ausführen.

## <a name="getting-started"></a>Erste Schritte

Um mit der Kombination von Power Automation und Office-Skripts zu beginnen, führen Sie das Lernprogramm [Start using scripts with Power Automation aus](../tutorials/excel-power-automate-manual.md). In diesem Artikel erfahren Sie, wie Sie einen Fluss erstellen, der ein einfaches Skript aufruft. Nachdem Sie das Lernprogramm und die [automatisch ausgeführten Skripts mit Power Automation](../tutorials/excel-power-automate-trigger.md) Tutorial abgeschlossen haben, erfahren Sie mehr über die Platt Form Integrationen.

## <a name="excel-online-business-connector"></a>Excel Online-Connector (Business)

[Connectors](/connectors/connectors) sind die Brücken zwischen Power Automation und Anwendungen. Der [Excel Online (Business)-Connector](/connectors/excelonlinebusiness) gibt Ihrem Fluss Zugriff auf Excel-Arbeitsmappen. Mit der Aktion "Skript ausführen" können Sie ein beliebiges Office-Skript aufrufen, das über die ausgewählte Arbeitsmappe zugänglich ist. Sie können nicht nur Skripts über einen Fluss ausführen, sondern auch Daten an und von der Arbeitsmappe mit dem Fluss durch die Skripts übergeben.

> [!IMPORTANT]
> Die Aktion "Skript ausführen" gibt Benutzern, die den Excel Connector verwenden, wichtigen Zugriff auf Ihre Arbeitsmappe und deren Daten. Darüber hinaus gibt es Sicherheitsrisiken mit Skripts, die externe API-Aufrufe durchführen, wie in [externe Aufrufe von Power Automation](external-calls.md)erläutert. Wenn Ihr Administrator mit der Exposition hoch vertraulicher Daten befasst ist, können Sie entweder den Excel Online Connector deaktivieren oder den Zugriff auf Office-Skripts über die [Office Scripts-Administrator Steuerelemente](https://support.microsoft.com/office/19d3c51a-6ca2-40ab-978d-60fa49554dcf)einschränken.

## <a name="passing-data-from-power-automate-into-a-script"></a>Übergeben von Daten aus Power Automation in ein Skript

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

## <a name="returning-data-from-a-script-back-to-power-automate"></a>Zurückgeben von Daten aus einem Skript an Power Automation

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
- [Automatisches Ausführen von Skripts mit Power Automation](../tutorials/excel-power-automate-trigger.md)
- [Grundlegendes zur Skripterstellung für Office-Skripts in Excel im Web](scripting-fundamentals.md)
- [Erste Schritte mit Power Automate](/power-automate/getting-started)
- [Referenzdokumentation zu Excel Online (Business) Connector](/connectors/excelonlinebusiness/)
