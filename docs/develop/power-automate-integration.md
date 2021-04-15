---
title: Ausführen von Office-Skripts mit Power Automate
description: So erhalten Sie Office Scripts for Excel im Web, die mit einem Power Automate-Workflow arbeiten.
ms.date: 12/16/2020
localization_priority: Normal
ms.openlocfilehash: 1ca9aa14efe7cf2c91100a32fbc9a69054012f06
ms.sourcegitcommit: 45ffe3dbd2c834b78592ad35928cf8096f5e80bc
ms.translationtype: MT
ms.contentlocale: de-DE
ms.lasthandoff: 04/14/2021
ms.locfileid: "51755070"
---
# <a name="run-office-scripts-with-power-automate"></a>Ausführen von Office-Skripts mit Power Automate

[Mit Power Automate](https://flow.microsoft.com) können Sie Office-Skripts zu einem größeren, automatisierten Workflow hinzufügen. Sie können Power Automate do-Aktionen verwenden, z. B. den Inhalt einer E-Mail zur Tabelle eines Arbeitsblatts hinzufügen oder Aktionen in Ihren Projektverwaltungstools basierend auf Arbeitsmappenkommentaren erstellen.

## <a name="getting-started"></a>Erste Schritte

Wenn Sie noch keine Power Automate-Version haben, empfehlen wir, [erste Schritte mit Power Automate zu besuchen.](/power-automate/getting-started) Dort erfahren Sie mehr über alle Automatisierungsmöglichkeiten, die Ihnen zur Verfügung stehen. Die Dokumente hier konzentrieren sich auf die Funktionsweise von Office-Skripts mit Power Automate und die Verbesserung Ihrer Excel-Erfahrung.

Um mit der Kombination von Power Automate- und Office-Skripts zu beginnen, folgen Sie dem Lernprogramm [Start using scripts with Power Automate](../tutorials/excel-power-automate-manual.md). Dadurch erfahren Sie, wie Sie einen Fluss erstellen, der ein einfaches Skript aufruft. Nachdem Sie dieses Lernprogramm und das Übergeben von Daten an Skripts in einem automatisch ausgeführten Power Automate-Fluss-Lernprogramm abgeschlossen haben, geben Sie hier ausführliche Informationen zum Verbinden von Office-Skripts mit Power [Automate-Flüssen](../tutorials/excel-power-automate-trigger.md) zurück.

## <a name="excel-online-business-connector"></a>Excel Online (Business)-Connector

[Connectors](/connectors/connectors) sind die Brücken zwischen Power Automate und Anwendungen. Der [Excel Online (Business)-Connector](/connectors/excelonlinebusiness) gewährt Ihren Flüssen Zugriff auf Excel-Arbeitsmappen. Mit der Aktion "Skript ausführen" können Sie ein beliebiges Office-Skript aufrufen, auf das über die ausgewählte Arbeitsmappe zugegriffen werden kann. Sie können Ihren Skripts auch Eingabeparameter geben, damit Daten vom Fluss bereitgestellt werden können, oder Ihr Skript kann Informationen zu späteren Schritten im Fluss zurückgeben.

> [!IMPORTANT]
> Die Aktion "Skript ausführen" ermöglicht Personen, die den Excel-Connector verwenden, erheblichen Zugriff auf Ihre Arbeitsmappe und ihre Daten. Darüber hinaus gibt es Sicherheitsrisiken bei Skripts, die externe API-Aufrufe ausführen, wie unter [Externe Aufrufe von Power Automate erläutert.](external-calls.md) Wenn Ihr Administrator mit der Belichtung besonders vertraulicher Daten zurecht kommt, kann er entweder den Excel Online-Connector deaktivieren oder den Zugriff auf Office-Skripts über die [Office Scripts-Administratorsteuerelemente einschränken.](/microsoft-365/admin/manage/manage-office-scripts-settings)

## <a name="data-transfer-in-flows-for-scripts"></a>Datenübertragung in Flüssen für Skripts

Mit Power Automate können Sie Datenteile zwischen den Schritten Ihres Datenflusses übergeben. Skripts können so konfiguriert werden, dass sie alle arten von Informationen akzeptieren, die Sie benötigen, und alle Informationen aus Ihrer Arbeitsmappe zurückgeben, die Sie in Ihrem Fluss wünschen. Die Eingabe für Ihr Skript wird durch Hinzufügen von Parametern zur `main` Funktion (zusätzlich zu `workbook: ExcelScript.Workbook` ) angegeben. Die Ausgabe des Skripts wird durch Hinzufügen eines Rückgabetyps zu `main` deklariert.

> [!NOTE]
> Wenn Sie einen "Skript ausführen"-Block in Ihrem Fluss erstellen, werden die akzeptierten Parameter und zurückgegebenen Typen aufgefüllt. Wenn Sie die Parameter oder Rückgabetypen Ihres Skripts ändern, müssen Sie den Block "Skript ausführen" des Ablaufs wiederholen. Dadurch wird sichergestellt, dass die Daten ordnungsgemäß analysiert werden.

In den folgenden Abschnitten werden die Details der Eingabe und Ausgabe für Skripts beschrieben, die in Power Automate verwendet werden. Wenn Sie einen praktischen Ansatz zum Erlernen dieses [Themas](../resources/scenarios/task-reminders.md) verwenden möchten, testen Sie das Beispielszenario Daten an Skripts übergeben in einem automatisch ausgeführten [Power Automate-Fluss-Lernprogramm](../tutorials/excel-power-automate-trigger.md) oder erkunden Sie das Beispielszenario Automatisierte Aufgabenerinnerungen.

### <a name="main-parameters-passing-data-to-a-script"></a>`main` Parameter: Übergeben von Daten an ein Skript

Alle Skripteingaben werden als zusätzliche Parameter für die Funktion `main` angegeben. Wenn Sie z. B. möchten, dass ein Skript einen Namen als Eingabe akzeptiert, ändern Sie `string` die `main` Signatur in `function main(workbook: ExcelScript.Workbook, name: string)` .

Wenn Sie einen Fluss in Power Automate konfigurieren, können Sie Skripteingaben als statische [Werte,](/power-automate/use-expressions-in-conditions)Ausdrücke oder dynamische Inhalte angeben. Details zum Connector eines einzelnen Diensts finden Sie in der [Power Automate Connector-Dokumentation](/connectors/).

Berücksichtigen Sie beim Hinzufügen von Eingabeparametern zur `main` Skriptfunktion die folgenden Freibeträge und Einschränkungen.

1. Der erste Parameter muss vom Typ `ExcelScript.Workbook` sein. Der Parametername spielt keine Rolle.

2. Jeder Parameter muss einen Typ (z. B. `string` oder `number` ) haben.

3. Die grundlegenden `string` Typen , , , , , und werden `number` `boolean` `any` `unknown` `object` `undefined` unterstützt.

4. Arrays der zuvor aufgeführten Basistypen werden unterstützt.

5. Geschachtelte Arrays werden als Parameter (jedoch nicht als Rückgabetypen) unterstützt.

6. Union-Typen sind zulässig, wenn sie eine Vereinigung von Literalen sind, die zu einem einzelnen Typ gehören (z. B. `"Left" | "Right"` ). Auch Unionen eines unterstützten Typs mit nicht definierten Werden unterstützt (z. B. `string | undefined` ).

7. Objekttypen sind zulässig, wenn sie Eigenschaften vom Typ `string` , , , unterstützte Arrays oder andere unterstützte Objekte `number` `boolean` enthalten. Das folgende Beispiel zeigt geschachtelte Objekte, die als Parametertypen unterstützt werden:

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

8. Objekte müssen ihre Schnittstelle oder Klassendefinition im Skript definiert haben. Ein Objekt kann auch anonym inline definiert werden, wie im folgenden Beispiel:

    ```TypeScript
    function main(workbook: ExcelScript.Workbook): {name: string, email: string}
    ```

9. Optionale Parameter sind zulässig und können mit dem optionalen Modifizierer (z. B. ) als solche bezeichnet `?` `function main(workbook: ExcelScript.Workbook, Name?: string)` werden.

10. Standardparameterwerte sind zulässig (z. `async function main(workbook: ExcelScript.Workbook, Name: string = 'Jane Doe')` B. .

### <a name="returning-data-from-a-script"></a>Zurückgeben von Daten aus einem Skript

Skripts können Daten aus der Arbeitsmappe zurückgeben, die als dynamischer Inhalt in einem Power Automate-Fluss verwendet werden sollen. Wie bei Eingabeparametern legt Power Automate einige Einschränkungen für den Rückgabetyp fest.

1. Die grundlegenden `string` Typen , , , und werden `number` `boolean` `void` `undefined` unterstützt.

2. Union-Typen, die als Rückgabetypen verwendet werden, folgen den gleichen Einschränkungen wie bei Verwendung als Skriptparameter.

3. Arraytypen sind zulässig, wenn sie vom Typ `string` , `number` oder `boolean` sind. Sie sind auch zulässig, wenn der Typ eine unterstützte Union oder ein unterstützter Literaltyp ist.

4. Objekttypen, die als Rückgabetypen verwendet werden, folgen den gleichen Einschränkungen wie bei Verwendung als Skriptparameter.

5. Die implizite Eingabe wird unterstützt, muss jedoch dieselben Regeln wie ein definierter Typ befolgen.

## <a name="example"></a>Beispiel

Der folgende Screenshot zeigt einen Power Automate-Fluss, der ausgelöst wird, wenn Ihnen ein [GitHub-Problem](https://github.com/) zugewiesen wird. Der Fluss führt ein Skript aus, das das Problem einer Tabelle in einer Excel-Arbeitsmappe hinzufügt. Wenn in dieser Tabelle fünf oder mehr Probleme auftreten, sendet der Fluss eine E-Mail-Erinnerung.

:::image type="content" source="../images/power-automate-parameter-return-sample.png" alt-text="Der Power Automate-Fluss-Editor, der den Beispielfluss zeigt.":::

Die Funktion des Skripts gibt die Problem-ID und den Ausgabetitel als Eingabeparameter an, und das Skript gibt die Anzahl der `main` Zeilen in der Problemtabelle zurück.

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

## <a name="see-also"></a>Weitere Informationen

- [Ausführen von Office-Skripts in Excel im Web mit Power Automate](../tutorials/excel-power-automate-manual.md)
- [Übergeben von Daten zu Skripts in einem automatisch ausgeführten Power Automate-Datenfluss](../tutorials/excel-power-automate-trigger.md)
- [Zurückgeben von Daten aus einem Skript an einen automatisch ausgeführten Power Automate-Flow](../tutorials/excel-power-automate-returns.md)
- [Problembehandlungsinformationen für Power Automate mit Office-Skripts](../testing/power-automate-troubleshooting.md)
- [Erste Schritte mit Power Automate](/power-automate/getting-started)
- [Referenzdokumentation zu Excel Online (Business)-Connectors](/connectors/excelonlinebusiness/)
