---
title: Ausführen von Office-Skripts mit Power Automate
description: So erhalten Sie Office-Skripts für Excel im Web arbeiten mit einem Power Automate-Workflow.
ms.date: 06/22/2022
ms.localizationpriority: medium
ms.openlocfilehash: 61e51861bd2c987c25d40e9ac6d2247122256918
ms.sourcegitcommit: c5ffe0a95b962936ee92e7ffe17388bef6d4fad8
ms.translationtype: MT
ms.contentlocale: de-DE
ms.lasthandoff: 06/29/2022
ms.locfileid: "66241856"
---
# <a name="run-office-scripts-with-power-automate"></a>Ausführen von Office-Skripts mit Power Automate

[Mit Power Automate](https://flow.microsoft.com) können Sie Office-Skripts zu einem größeren, automatisierten Workflow hinzufügen. Mit Power Automate können Sie beispielsweise den Inhalt einer E-Mail zur Tabelle eines Arbeitsblatts hinzufügen oder Aktionen in Ihren Projektverwaltungstools basierend auf Arbeitsmappenkommentaren erstellen.

## <a name="get-started"></a>Erste Schritte

Wenn Sie noch nicht mit Power Automate begonnen haben, empfehlen wir Ihnen, ["Erste Schritte mit Power Automate](/power-automate/getting-started)" zu besuchen. Dort erfahren Sie mehr über alle Automatisierungsmöglichkeiten, die Ihnen zur Verfügung stehen. In den hier aufgeführten Dokumenten wird erläutert, wie Office-Skripts mit Power Automate funktionieren und wie dies dazu beitragen kann, Ihre Excel-Erfahrung zu verbessern.

### <a name="step-by-step-tutorials"></a>Schritt-für-Schritt-Lernprogramme

Es gibt drei schrittweise Lernprogramme für Power Automate und Office-Skripts. Diese zeigen, wie Sie die Automatisierungsdienste kombinieren und Daten zwischen einer Arbeitsmappe und einem Fluss übergeben.

- [Aufrufen von Skripts aus einem manuellen Power Automate-Datenfluss](../tutorials/excel-power-automate-manual.md)
- [Übergeben von Daten zu Skripts in einem automatisch ausgeführten Power Automate-Datenfluss](../tutorials/excel-power-automate-trigger.md)
- [Zurückgeben von Daten aus einem Skript an einen automatisch ausgeführten Power Automate-Flow](../tutorials//excel-power-automate-returns.md)

### <a name="create-a-flow-from-excel"></a>Erstellen eines Flusses aus Excel

Sie können mit einer Vielzahl von Flussvorlagen mit Power Automate in Excel beginnen. Wählen Sie auf der Registerkarte **"Automatisieren** " die Option " **Aufgabe automatisieren**" aus.

:::image type="content" source="../images/automate-a-task-button.png" alt-text="Die Schaltfläche &quot;Aufgabe automatisieren&quot; im Menüband.":::

Dadurch wird ein Aufgabenbereich mit mehreren Optionen geöffnet, um ihre Office-Skripts mit größeren automatisierten Lösungen zu verbinden. Wählen Sie eine beliebige Option aus, um zu beginnen. Ihr Fluss wird mit der aktuellen Arbeitsmappe bereitgestellt.

:::image type="content" source="../images/automate-a-task-choices.png" alt-text="Ein Aufgabenbereich mit Flussvorlagenoptionen wie &quot;Planen der Ausführung eines Office-Skripts in Excel und senden Sie dann eine E-Mail&quot; und &quot;Ausführen eines Office-Skripts in Excel, wenn eine Microsoft Forms Antwort empfangen wird&quot;.":::

> [!TIP]
> Sie können auch mit dem Erstellen eines Flusses über das Menü **"Weitere Optionen (...)"** in einem einzelnen Skript beginnen.

## <a name="excel-online-business-connector"></a>Excel Online (Business)-Connector

[Connectors](/connectors/connectors) sind die Brücken zwischen Power Automate und Anwendungen. Der [Excel Online (Business)-Connector](/connectors/excelonlinebusiness) bietet Ihren Flüssen Zugriff auf Excel-Arbeitsmappen. Mit der Aktion "Skript ausführen" können Sie ein beliebiges Office-Skript aufrufen, auf das über die ausgewählte Arbeitsmappe zugegriffen werden kann. Sie können Ihren Skripts auch Eingabeparameter angeben, damit Daten vom Fluss bereitgestellt werden können, oder Sie können festlegen, dass das Skript Informationen für spätere Schritte im Fluss zurückgibt.

> [!IMPORTANT]
> Die Aktion "Skript ausführen" ermöglicht Personen, die den Excel-Connector verwenden, erheblichen Zugriff auf Ihre Arbeitsmappe und deren Daten. Darüber hinaus bestehen Sicherheitsrisiken mit Skripts, die externe API-Aufrufe ausführen, wie in [externen Aufrufen von Power Automate](external-calls.md) erläutert. Wenn Sich Ihr Administrator mit der Belichtung von streng vertraulichen Daten beschäftigt, kann er entweder den Excel Online-Connector deaktivieren oder den Zugriff auf Office-Skripts über die [Administratorsteuerelemente von Office-Skripts](/microsoft-365/admin/manage/manage-office-scripts-settings) einschränken.

> [!IMPORTANT]
> Power Automate unterstützt derzeit **keine** Skripts, die in SharePoint gespeichert sind.

## <a name="data-transfer-in-flows-for-scripts"></a>Datenübertragung in Flows für Skripts

Mit Power Automate können Sie Datenteile zwischen den Schritten Ihres Flusses übergeben. Skripts können so konfiguriert werden, dass sie alle arten von Informationen akzeptieren, die Sie benötigen, und alles aus Ihrer Arbeitsmappe zurückgeben, die Sie in Ihrem Fluss benötigen. Die Eingabe für Das Skript wird durch Hinzufügen von Parametern zur `main` Funktion (zusätzlich zu `workbook: ExcelScript.Workbook`) angegeben. Die Ausgabe des Skripts wird durch Hinzufügen eines Rückgabetyps zu `main`deklariert.

> [!NOTE]
> Wenn Sie einen "Skript ausführen"-Block in Ihrem Fluss erstellen, werden die akzeptierten Parameter und zurückgegebenen Typen aufgefüllt. Wenn Sie die Parameter ändern oder Typen Ihres Skripts zurückgeben, müssen Sie den Block "Skript ausführen" ihres Flows wiederholen. Dadurch wird sichergestellt, dass die Daten ordnungsgemäß analysiert werden.

In den folgenden Abschnitten werden die Details der Eingabe und Ausgabe für Skripts behandelt, die in Power Automate verwendet werden. Wenn Sie einen praktischen Ansatz zum Erlernen dieses Themas wünschen, probieren Sie die [Übergabe von Daten an Skripts in einem automatisch ausgeführten Power Automate-Ablauflernprogramm](../tutorials/excel-power-automate-trigger.md) aus, oder erkunden Sie das Beispielszenario [für automatisierte Aufgabenerinnerungen](../resources/scenarios/task-reminders.md) .

### <a name="main-parameters-pass-data-to-a-script"></a>`main` Parameter: Übergeben von Daten an ein Skript

Alle Skripteingaben werden als zusätzliche Parameter für die `main` Funktion angegeben. Wenn Sie beispielsweise möchten, dass ein Skript ein `string` Skript akzeptiert, das einen Namen als Eingabe darstellt, ändern Sie die `main` Signatur in `function main(workbook: ExcelScript.Workbook, name: string)`.

Wenn Sie einen Fluss in Power Automate konfigurieren, können Sie Skripteingaben als statische Werte, [Ausdrücke](/power-automate/use-expressions-in-conditions) oder dynamische Inhalte angeben. Details zum Connector eines einzelnen Diensts finden Sie in der Dokumentation zu [Power Automate Connector](/connectors/).

#### <a name="type-restrictions"></a>Typeinschränkungen

Berücksichtigen Sie beim Hinzufügen von Eingabeparametern zur Funktion eines Skripts `main` die folgenden Zertifikate und Einschränkungen. Diese gelten auch für den Rückgabetyp des Skripts.

1. Der erste Parameter muss vom Typ `ExcelScript.Workbook`sein. Der Name des Parameters spielt keine Rolle.

1. Die Typen `string`, `number`, `boolean`, `unknown`, und `object``undefined` werden unterstützt.

1. Arrays (sowohl als `Array<T>` auch `[]` Formatvorlagen) der zuvor aufgeführten Typen werden unterstützt. Geschachtelte Arrays werden ebenfalls unterstützt.

1. Union-Typen sind zulässig, wenn es sich um eine Vereinigung von Literalen handelt, die zu einem einzelnen Typ gehören (z `"Left" | "Right"`. B. nicht `"Left", 5`). Auch Vereinigungen eines unterstützten Typs mit nicht definierter Größe werden unterstützt (z `string | undefined`. B. ).

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

1. Optionale Parameter sind zulässig und werden mit dem optionalen Modifizierer `?` angegeben (z. B `function main(workbook: ExcelScript.Workbook, Name?: string)`. ).

1. Standardwerte sind zulässig (z. B `function main(workbook: ExcelScript.Workbook, Name: string = 'Jane Doe')`. .

### <a name="return-data-from-a-script"></a>Zurückgeben von Daten aus einem Skript

Skripts können Daten aus der Arbeitsmappe zurückgeben, die als dynamischer Inhalt in einem Power Automate-Fluss verwendet werden sollen. Die [zuvor aufgeführten Typeinschränkungen](#type-restrictions) gelten auch für den Rückgabetyp. Um ein Objekt zurückzugeben, fügen Sie der Funktion die Rückgabetypsyntax `main` hinzu. Wenn Sie beispielsweise einen `string` Wert aus dem Skript zurückgeben möchten, lautet `function main(workbook: ExcelScript.Workbook): string`Ihre `main` Signatur .

## <a name="example"></a>Beispiel

Der folgende Screenshot zeigt einen Power Automate-Fluss, der ausgelöst wird, wenn Ihnen ein [GitHub-Problem](https://github.com/) zugewiesen wird. Der Fluss führt ein Skript aus, das das Problem einer Tabelle in einer Excel-Arbeitsmappe hinzufügt. Wenn diese Tabelle fünf oder mehr Probleme enthält, sendet der Fluss eine E-Mail-Erinnerung.

:::image type="content" source="../images/power-automate-parameter-return-sample.png" alt-text="Der Power Automate-Fluss-Editor mit dem Beispielfluss.":::

Die `main` Funktion des Skripts gibt die Problem-ID und den Problemtitel als Eingabeparameter an, und das Skript gibt die Anzahl der Zeilen in der Problemtabelle zurück.

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

- [Aufrufen von Skripts aus einem manuellen Power Automate-Datenfluss](../tutorials/excel-power-automate-manual.md)
- [Übergeben von Daten zu Skripts in einem automatisch ausgeführten Power Automate-Datenfluss](../tutorials/excel-power-automate-trigger.md)
- [Zurückgeben von Daten aus einem Skript an einen automatisch ausgeführten Power Automate-Flow](../tutorials/excel-power-automate-returns.md)
- [Problembehandlungsinformationen für Power Automate mit Office-Skripts](../testing/power-automate-troubleshooting.md)
- [Erste Schritte mit Power Automate](/power-automate/getting-started)
- [Referenzdokumentation zum Excel Online (Business)-Connector](/connectors/excelonlinebusiness/)
