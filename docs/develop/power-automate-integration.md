---
title: Ausführen Office Skripts mit Power Automate
description: So erhalten Sie Office Skripts für Excel im Web, die mit einem Power Automate-Workflow arbeiten.
ms.date: 03/08/2022
ms.localizationpriority: medium
ms.openlocfilehash: dbf65086e564b20ca0fc3a4dc1c527188540be6b
ms.sourcegitcommit: 7023b9e23499806901a5ecf8ebc460b76887cca6
ms.translationtype: MT
ms.contentlocale: de-DE
ms.lasthandoff: 03/31/2022
ms.locfileid: "64585751"
---
# <a name="run-office-scripts-with-power-automate"></a>Ausführen Office Skripts mit Power Automate

[Power Automate](https://flow.microsoft.com) können Sie einem größeren, automatisierten Workflow Office Skripts hinzufügen. Mit Power Automate können Sie z. B. den Inhalt einer E-Mail zur Tabelle eines Arbeitsblatts hinzufügen oder Aktionen in Ihren Projektverwaltungstools basierend auf Arbeitsmappenkommentaren erstellen.

## <a name="get-started"></a>Erste Schritte

Wenn Sie noch nicht mit Power Automate vertraut sind, empfehlen wir, [Erste Schritte mit Power Automate](/power-automate/getting-started) zu besuchen. Dort können Sie mehr über alle Automatisierungsmöglichkeiten erfahren, die Ihnen zur Verfügung stehen. Die dokumente hier konzentrieren sich darauf, wie Office Skripts mit Power Automate arbeiten und wie dies dazu beitragen kann, Ihre Excel Erfahrung zu verbessern.

Um mit der Kombination von Power Automate und Office Skripts zu beginnen, folgen Sie dem Lernprogramm [" Verwenden von Skripts mit Power Automate](../tutorials/excel-power-automate-manual.md)". Dadurch erfahren Sie, wie Sie einen Fluss erstellen, der ein einfaches Skript aufruft. Nachdem Sie dieses Lernprogramm abgeschlossen haben und die [Daten an Skripts in einem automatisch ausgeführten Power Automate-Fluss-Lernprogramm übergeben](../tutorials/excel-power-automate-trigger.md) haben, kehren Sie hier zurück, um ausführliche Informationen zum Verbinden von Office Skripts mit Power Automate Flüssen zu finden.

## <a name="excel-online-business-connector"></a>Excel Online (Business)-Connector

[Connectors](/connectors/connectors) sind die Brücken zwischen Power Automate und Anwendungen. Der [connector Excel Online (Business)](/connectors/excelonlinebusiness) bietet Ihren Flüssen Zugriff auf Excel Arbeitsmappen. Mit der Aktion "Skript ausführen" können Sie jedes Office Skript aufrufen, auf das über die ausgewählte Arbeitsmappe zugegriffen werden kann. Sie können Ihren Skripts auch Eingabeparameter geben, damit Daten vom Fluss bereitgestellt werden können, oder Sie können festlegen, dass Ihr Skript Informationen für spätere Schritte im Fluss zurückgibt.

> [!IMPORTANT]
> Die Aktion "Skript ausführen" bietet Personen, die den Excel Connector verwenden, erheblichen Zugriff auf Ihre Arbeitsmappe und ihre Daten. Darüber hinaus gibt es Sicherheitsrisiken bei Skripts, die externe API-Aufrufe ausführen, wie in [externen Aufrufen von Power Automate](external-calls.md) erläutert. Wenn Sich Ihr Administrator mit der Offenlegung streng vertraulicher Daten befasst, kann er entweder den Excel Online-Connector deaktivieren oder den Zugriff auf Office Skripts über die [Administratorsteuerelemente Office Skripts](/microsoft-365/admin/manage/manage-office-scripts-settings) einschränken.

## <a name="data-transfer-in-flows-for-scripts"></a>Datenübertragung in Flüssen für Skripts

mit Power Automate können Sie Datenteile zwischen den Schritten des Flusses übergeben. Skripts können so konfiguriert werden, dass sie alle benötigten Informationstypen akzeptieren und alles aus Ihrer Arbeitsmappe zurückgeben, die Sie in Ihrem Flow wünschen. Die Eingabe für Ihr Skript wird durch Hinzufügen von Parametern zur `main` Funktion (zusätzlich zu `workbook: ExcelScript.Workbook`) angegeben. Die Ausgabe des Skripts wird deklariert, indem ein Rückgabetyp zu `main`hinzugefügt wird.

> [!NOTE]
> Wenn Sie einen "Skript ausführen"-Block in Ihrem Flow erstellen, werden die akzeptierten Parameter und zurückgegebenen Typen aufgefüllt. Wenn Sie die Parameter ändern oder Typen ihres Skripts zurückgeben, müssen Sie den "Skript ausführen"-Block des Flusses wiederholen. Dadurch wird sichergestellt, dass die Daten richtig analysiert werden.

In den folgenden Abschnitten werden die Details der Eingabe und Ausgabe für Skripts behandelt, die in Power Automate verwendet werden. Wenn Sie einen praktischen Ansatz zum Erlernen dieses Themas wünschen, testen Sie die [Pass-Daten an Skripts in einem automatisch ausgeführten Lernprogramm Power Automate Fluss](../tutorials/excel-power-automate-trigger.md) oder erkunden Sie das Beispielszenario für [automatisierte Aufgabenerinnerungen](../resources/scenarios/task-reminders.md).

### <a name="main-parameters-pass-data-to-a-script"></a>`main` Parameter: Übergeben von Daten an ein Skript

Alle Skripteingaben werden als zusätzliche Parameter für die `main` Funktion angegeben. Wenn sie beispielsweise möchten, dass ein Skript einen `string` Namen akzeptiert, der einen Namen als Eingabe darstellt, ändern Sie die `main` Signatur in `function main(workbook: ExcelScript.Workbook, name: string)`.

Wenn Sie einen Fluss in Power Automate konfigurieren, können Sie Skripteingaben als statische Werte, [Ausdrücke](/power-automate/use-expressions-in-conditions) oder dynamischen Inhalt angeben. Details zum Connector eines einzelnen Diensts finden Sie in der [Dokumentation zu Power Automate Connector](/connectors/).

#### <a name="type-restrictions"></a>Typeinschränkungen

Berücksichtigen Sie beim Hinzufügen von Eingabeparametern zur Funktion eines Skripts `main` die folgenden Einschränkungen und Einschränkungen. Diese gelten auch für den Rückgabetyp des Skripts.

1. Der erste Parameter muss vom Typ `ExcelScript.Workbook`sein. Der Parametername spielt keine Rolle.

1. Die Typen `string`, `number`, `boolean`, , `unknown`, und `undefined` `object`werden unterstützt.

1. Arrays (sowohl als `Array<T>` auch `[]` Formatvorlagen) der zuvor aufgeführten Typen werden unterstützt. Geschachtelte Arrays werden ebenfalls unterstützt.

1. Union-Typen sind zulässig, wenn sie eine Vereinigung von Literalen sind, die zu einem einzelnen Typ gehören (z `"Left" | "Right"`. B. , nicht `"Left", 5`). Außerdem werden Verbindungen eines unterstützten Typs mit nicht definierten Typen unterstützt (z `string | undefined`. B. ).

1. Objekttypen sind zulässig, wenn sie Eigenschaften vom Typ `string`, `number`, `boolean`unterstützten Arrays oder anderen unterstützten Objekten enthalten. Das folgende Beispiel zeigt geschachtelte Objekte, die als Parametertypen unterstützt werden.

    ```TypeScript
    // The Employee object is supported because Position is also composed of supported types.
    interface Employee {
        name: string;
        job: Position;
    }

    interface Position {
        id: number;
        title: string;
    }
    ```

1. Für Objekte muss ihre Schnittstelle oder Klassendefinition im Skript definiert sein. Ein Objekt kann auch anonym inline definiert werden, wie im folgenden Beispiel gezeigt.

    ```TypeScript
    function main(workbook: ExcelScript.Workbook): {name: string, email: string}
    ```

#### <a name="optional-and-default-parameters"></a>Optionale und Standardparameter

1. Optionale Parameter sind zulässig und werden mit dem optionalen Modifizierer `?` (z. B `function main(workbook: ExcelScript.Workbook, Name?: string)`. ) angegeben.

1. Standardwerte sind zulässig (z. B `function main(workbook: ExcelScript.Workbook, Name: string = 'Jane Doe')`. .

### <a name="return-data-from-a-script"></a>Zurückgeben von Daten aus einem Skript

Skripts können Daten aus der Arbeitsmappe zurückgeben, die als dynamischer Inhalt in einem Power Automate Fluss verwendet werden sollen. Die [gleichen Zuvor aufgeführten Typeinschränkungen](#type-restrictions) gelten für den Rückgabetyp. Um ein Objekt zurückzugeben, fügen Sie der Funktion die Rückgabetypsyntax `main` hinzu. Wenn Sie beispielsweise einen `string` Wert aus dem Skript zurückgeben möchten, lautet `function main(workbook: ExcelScript.Workbook): string`die `main` Signatur .

## <a name="example"></a>Beispiel

Der folgende Screenshot zeigt einen Power Automate Fluss, der ausgelöst wird, wenn Ihnen ein [GitHub](https://github.com/) Problem zugewiesen wird. Der Fluss führt ein Skript aus, das das Problem einer Tabelle in einer Excel Arbeitsmappe hinzufügt. Wenn in dieser Tabelle fünf oder mehr Probleme auftreten, sendet der Fluss eine E-Mail-Erinnerung.

:::image type="content" source="../images/power-automate-parameter-return-sample.png" alt-text="Der Power Automate Fluss-Editor mit dem Beispielfluss.":::

Die `main` Funktion des Skripts gibt die Problem-ID und den Ausgabetitel als Eingabeparameter an, und das Skript gibt die Anzahl der Zeilen in der Problemtabelle zurück.

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

- [Ausführen Office Skripts in Excel im Web mit Power Automate](../tutorials/excel-power-automate-manual.md)
- [Übergeben von Daten zu Skripts in einem automatisch ausgeführten Power Automate-Datenfluss](../tutorials/excel-power-automate-trigger.md)
- [Zurückgeben von Daten aus einem Skript an einen automatisch ausgeführten Power Automate-Flow](../tutorials/excel-power-automate-returns.md)
- [Problembehandlungsinformationen für Power Automate mit Office Skripts](../testing/power-automate-troubleshooting.md)
- [Erste Schritte mit Power Automate](/power-automate/getting-started)
- [Referenzdokumentation für Excel Online connector (Business)](/connectors/excelonlinebusiness/)
