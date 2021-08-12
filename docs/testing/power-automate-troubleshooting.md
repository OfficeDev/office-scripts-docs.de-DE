---
title: Problembehandlung bei Office Skripts, die in Power Automate ausgeführt werden
description: Tipps, Plattforminformationen und bekannte Probleme bei der Integration von Office Skripts und Power Automate.
ms.date: 05/18/2021
localization_priority: Normal
ms.openlocfilehash: 1746a03022b6d1aa9fc35e1a8875add301dd6a0f2d6d45cedd64308f0738d2f8
ms.sourcegitcommit: 75f7ed8c2d23a104acc293f8ce29ea580b4fcdc5
ms.translationtype: MT
ms.contentlocale: de-DE
ms.lasthandoff: 08/11/2021
ms.locfileid: "57847207"
---
# <a name="troubleshoot-office-scripts-running-in-power-automate"></a>Problembehandlung bei Office Skripts, die in Power Automate ausgeführt werden

Power Automate können Sie ihre Office Skriptautomatisierung auf die nächste Ebene bringen. Da Power Automate jedoch Skripts in Ihrem Namen in unabhängigen Excel Sitzungen ausführt, sind einige wichtige Dinge zu beachten.

> [!TIP]
> Wenn Sie gerade erst mit der Verwendung Office Skripts mit Power Automate beginnen, beginnen Sie mit ["Ausführen Office Skripts mit Power Automate",](../develop/power-automate-integration.md) um mehr über die Plattformen zu erfahren.

## <a name="avoid-relative-references"></a>Vermeiden von relativen Verweisen

Power Automate führt Ihr Skript in der ausgewählten Excel Arbeitsmappe in Ihrem Namen aus. In diesem Fall wird die Arbeitsmappe möglicherweise geschlossen. Jede API, die auf dem aktuellen Status des Benutzers basiert, z. B. `Workbook.getActiveWorksheet` kann sich in Power Automate anders verhalten. Dies liegt daran, dass die APIs auf einer relativen Position der Ansicht oder des Cursors des Benutzers basieren und dieser Verweis in einem Power Automate Fluss nicht vorhanden ist.

Einige relative Verweis-APIs lösen Fehler in Power Automate aus. Andere weisen ein Standardverhalten auf, das den Status eines Benutzers impliziert. Achten Sie beim Entwerfen ihrer Skripts darauf, absolute Verweise für Arbeitsblätter und Bereiche zu verwenden. Dadurch wird ihr Power Automate Fluss konsistent, auch wenn Arbeitsblätter neu angeordnet werden.

### <a name="script-methods-that-fail-when-run-in-power-automate-flows"></a>Skriptmethoden, die fehlschlagen, wenn sie in Power Automate Flüssen ausgeführt werden

Die folgenden Methoden lösen einen Fehler aus und schlagen fehl, wenn sie von einem Skript in einem Power Automate-Fluss aufgerufen werden.

| Klasse | Methode |
|--|--|
| [Chart](/javascript/api/office-scripts/excelscript/excelscript.chart) | `activate` |
| [Range](/javascript/api/office-scripts/excelscript/excelscript.range) | `select` |
| [Workbook](/javascript/api/office-scripts/excelscript/excelscript.workbook) | `getActiveCell` |
| [Workbook](/javascript/api/office-scripts/excelscript/excelscript.workbook) | `getActiveChart` |
| [Workbook](/javascript/api/office-scripts/excelscript/excelscript.workbook) | `getActiveSlicer` |
| [Workbook](/javascript/api/office-scripts/excelscript/excelscript.workbook) | `getSelectedRange` |
| [Workbook](/javascript/api/office-scripts/excelscript/excelscript.workbook) | `getSelectedRanges` |

### <a name="script-methods-with-a-default-behavior-in-power-automate-flows"></a>Skriptmethoden mit einem Standardverhalten in Power Automate Flüssen

Die folgenden Methoden verwenden ein Standardverhalten anstelle des aktuellen Status eines beliebigen Benutzers.

| Klasse | Methode | Power Automate Verhalten |
|--|--|--|
| [Workbook](/javascript/api/office-scripts/excelscript/excelscript.workbook) | `getActiveWorksheet` | Gibt entweder das erste Arbeitsblatt in der Arbeitsmappe oder das Arbeitsblatt zurück, das derzeit von der Methode aktiviert `Worksheet.activate` wird. |
| [Arbeitsblatt](/javascript/api/office-scripts/excelscript/excelscript.worksheet) | `activate` | Markiert das Arbeitsblatt als aktives Arbeitsblatt für die Zwecke von `Workbook.getActiveWorksheet` . |

## <a name="data-refresh-not-supported-in-power-automate"></a>Datenaktualisierung in Power Automate nicht unterstützt

Office Skripts können Daten nicht aktualisieren, wenn sie in Power Automate ausgeführt werden. Methoden, `PivotTable.refresh` z. B. "Nichts tun", wenn sie in einem Fluss aufgerufen werden. Darüber hinaus löst Power Automate keine Datenaktualisierung für Formeln aus, die Arbeitsmappenlinks verwenden.

### <a name="script-methods-that-do-nothing-when-run-in-power-automate-flows"></a>Skriptmethoden, die keine Aktionen ausführen, wenn sie in Power Automate Flüssen ausgeführt werden

Die folgenden Methoden führen keine Aktionen in einem Skript durch, wenn sie über Power Automate aufgerufen werden. Sie werden weiterhin erfolgreich zurückgegeben und lösen keine Fehler aus.

| Klasse | Methode |
|--|--|
| [PivotTable](/javascript/api/office-scripts/excelscript/excelscript.pivottable) | `refresh` |
| [Workbook](/javascript/api/office-scripts/excelscript/excelscript.workbook) | `refreshAllDataConnections` |
| [Workbook](/javascript/api/office-scripts/excelscript/excelscript.workbook) | `refreshAllPivotTables` |
| [Arbeitsblatt](/javascript/api/office-scripts/excelscript/excelscript.worksheet) | `refreshAllPivotTables` |

## <a name="select-workbooks-with-the-file-browser-control"></a>Auswählen von Arbeitsmappen mit dem Dateibrowsersteuerelement

Beim Erstellen des **Ausführungsskriptschritts** eines Power Automate Flusses müssen Sie auswählen, welche Arbeitsmappe Teil des Flusses ist. Verwenden Sie den Dateibrowser, um Ihre Arbeitsmappe auszuwählen, anstatt den Namen der Arbeitsmappe manuell einzugeben.

:::image type="content" source="../images/power-automate-file-browser.png" alt-text="Die Skriptaktion Power Automate Ausführen mit der Browseroption &quot;Dateiauswahl anzeigen&quot;.":::

For more context on the Power Automate limitation and a discussion of potential workarounds for the dynamic selection of workbooks, see [this thread in the Microsoft Power Automate Community](https://powerusers.microsoft.com/t5/Power-Automate-Ideas/Allow-for-dynamic-quot-file-quot-value-for-excel-quot-get-a-row/idi-p/103091#).

## <a name="time-zone-differences"></a>Zeitzonenunterschiede

Excel Dateien haben keinen inhärenten Speicherort oder keine Zeitzone. Jedes Mal, wenn ein Benutzer die Arbeitsmappe öffnet, verwendet seine Sitzung die lokale Zeitzone dieses Benutzers für Datumsberechnungen. Power Automate verwendet immer UTC.

Wenn Ihr Skript Datumsangaben oder Uhrzeiten verwendet, kann es Verhaltensunterschiede geben, wenn das Skript lokal getestet wird und nicht, wenn es Power Automate ausgeführt wird. Power Automate können Sie Zeiten konvertieren, formatieren und anpassen. Anweisungen zur Verwendung dieser Funktionen in Power Automate und Parametern finden Sie unter [Working with Dates and Times inside of your](https://flow.microsoft.com/blog/working-with-dates-and-times/) [ `main` flows: Pass data to a script](../develop/power-automate-integration.md#main-parameters-pass-data-to-a-script) to learn how to provide that time information for the script.

## <a name="see-also"></a>Weitere Artikel

- [Problembehandlung bei Office Skripts](troubleshooting.md)
- [Ausführen von Office Skripts mit Power Automate](../develop/power-automate-integration.md)
- [Excel Referenzdokumentation zu Online-Connectors (Business)](/connectors/excelonlinebusiness/)
