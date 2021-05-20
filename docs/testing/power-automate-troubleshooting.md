---
title: Beheben Office Skripts, die in Power Automate ausgeführt werden
description: Tipps, Plattforminformationen und bekannte Probleme bei der Integration zwischen Office Skripts und Power Automate.
ms.date: 05/17/2021
localization_priority: Normal
ms.openlocfilehash: e26378051c764d97b4e8d748abc85fbe095c7b03
ms.sourcegitcommit: 4687693f02fc90a57ba30c461f35046e02e6f5fb
ms.translationtype: MT
ms.contentlocale: de-DE
ms.lasthandoff: 05/19/2021
ms.locfileid: "52545570"
---
# <a name="troubleshoot-office-scripts-running-in-power-automate"></a>Beheben Office Skripts, die in Power Automate ausgeführt werden

Power Automate können Sie Ihre Office Script-Automatisierung auf die nächste Stufe bringen. Da Power Automate Skripts in Ihrem Namen in unabhängigen Excel Sitzungen ausführt, gibt es jedoch einige wichtige Dinge zu beachten.

> [!TIP]
> Wenn Sie gerade erst anfangen, Office Scripts mit Power Automate zu verwenden, beginnen Sie bitte mit [Ausführen Office Skripts mit Power Automate,](../develop/power-automate-integration.md) um mehr über die Plattformen zu erfahren.

## <a name="avoid-relative-references"></a>Vermeiden Sie relative Referenzen

Power Automate führt Ihr Skript in der ausgewählten arbeitsmappe Excel in Ihrem Namen aus. Die Arbeitsmappe wird möglicherweise geschlossen, wenn dies geschieht. Jede API, die auf dem aktuellen Status des Benutzers basiert, z. `Workbook.getActiveWorksheet` B. , verhält sich in Power Automate möglicherweise anders. Dies liegt daran, dass die APIs auf einer relativen Position der Ansicht oder des Cursors des Benutzers basieren und dieser Verweis in einem Power Automate-Flow nicht vorhanden ist.

Einige relative Referenz-APIs werfen Fehler in Power Automate aus. Andere haben ein Standardverhalten, das den Status eines Benutzers impliziert. Achten Sie beim Entwerfen Ihrer Skripts darauf, absolute Referenzen für Arbeitsblätter und Bereiche zu verwenden. Dadurch wird ihr Power Automate Fluss konsistent, auch wenn Arbeitsblätter neu angeordnet werden.

### <a name="script-methods-that-fail-when-run-power-automate-flows"></a>Skriptmethoden, die beim Ausführen Power Automate-Flows fehlschlagen

Die folgenden Methoden werden einen Fehler auslösen und schlagen fehl, wenn sie von einem Skript in einem Power Automate-Flow aufgerufen werden.

| Klasse | Methode |
|--|--|
| [Chart](/javascript/api/office-scripts/excelscript/excelscript.chart) | `activate` |
| [Range](/javascript/api/office-scripts/excelscript/excelscript.range) | `select` |
| [Workbook](/javascript/api/office-scripts/excelscript/excelscript.workbook) | `getActiveCell` |
| [Workbook](/javascript/api/office-scripts/excelscript/excelscript.workbook) | `getActiveChart` |
| [Workbook](/javascript/api/office-scripts/excelscript/excelscript.workbook) | `getActiveSlicer` |
| [Workbook](/javascript/api/office-scripts/excelscript/excelscript.workbook) | `getSelectedRange` |
| [Workbook](/javascript/api/office-scripts/excelscript/excelscript.workbook) | `getSelectedRanges` |

### <a name="script-methods-with-a-default-behavior-in-power-automate-flows"></a>Skriptmethoden mit einem Standardverhalten in Power Automate-Flows

Die folgenden Methoden verwenden ein Standardverhalten anstelle des aktuellen Status eines Benutzers.

| Klasse | Methode | Power Automate Verhalten |
|--|--|--|
| [Workbook](/javascript/api/office-scripts/excelscript/excelscript.workbook) | `getActiveWorksheet` | Gibt entweder das erste Arbeitsblatt in der Arbeitsmappe oder das Arbeitsblatt zurück, das derzeit von der Methode aktiviert `Worksheet.activate` ist. |
| [Arbeitsblatt](/javascript/api/office-scripts/excelscript/excelscript.worksheet) | `activate` | Markiert das Arbeitsblatt als aktives Arbeitsblatt für die Zwecke `Workbook.getActiveWorksheet` von . |

## <a name="select-workbooks-with-the-file-browser-control"></a>Auswählen von Arbeitsmappen mit dem Dateibrowsersteuerelement

Beim Erstellen des **Run-Skriptschritts** eines Power Automate-Flows müssen Sie auswählen, welche Arbeitsmappe Teil des Flows ist. Verwenden Sie den Dateibrowser, um Ihre Arbeitsmappe auszuwählen, anstatt den Namen der Arbeitsmappe manuell einzugeben.

:::image type="content" source="../images/power-automate-file-browser.png" alt-text="Die Power Automate Skriptaktion ausführen, die die Browseroption &quot;Auswahldatei anzeigen&quot; anzeigt":::

Weitere Kontextinformationen zur Power Automate Einschränkung und eine Erläuterung möglicher Problemumgehungen für die dynamische Auswahl von Arbeitsmappen finden Sie [in diesem Thread im Microsoft Power Automate Community](https://powerusers.microsoft.com/t5/Power-Automate-Ideas/Allow-for-dynamic-quot-file-quot-value-for-excel-quot-get-a-row/idi-p/103091#).

## <a name="time-zone-differences"></a>Zeitzonenunterschiede

Excel Dateien haben keinen inhärenten Speicherort oder keine Zeitzone. Jedes Mal, wenn ein Benutzer die Arbeitsmappe öffnet, verwendet seine Sitzung die lokale Zeitzone dieses Benutzers für Datumsberechnungen. Power Automate verwendet immer UTC.

Wenn Ihr Skript Datums- oder Uhrzeitangaben verwendet, kann es Verhaltensunterschiede geben, wenn das Skript lokal getestet wird, und wenn es durch Power Automate ausgeführt wird. Power Automate können Sie Die Zeiten konvertieren, formatieren und anpassen. Unter [Arbeiten mit Datumsangaben und Zeiten innerhalb Ihrer Flows](https://flow.microsoft.com/blog/working-with-dates-and-times/) finden Sie Anweisungen zur Verwendung dieser Funktionen in Power Automate und [ `main` Parametern: Übergeben Sie Daten an ein Skript,](../develop/power-automate-integration.md#main-parameters-pass-data-to-a-script) um zu erfahren, wie Sie diese Zeitinformationen für das Skript bereitstellen.

## <a name="see-also"></a>Siehe auch

- [Fehlerbehebung Office Skripts](troubleshooting.md)
- [Ausführen Office Skripts mit Power Automate](../develop/power-automate-integration.md)
- [Excel Online -Connector-Referenzdokumentation (Business)](/connectors/excelonlinebusiness/)
