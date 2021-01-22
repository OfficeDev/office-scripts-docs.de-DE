---
title: Problembehandlungsinformationen für Power Automate mit Office-Skripts
description: Tipps, Plattforminformationen und bekannte Probleme bei der Integration zwischen Office Scripts und Power Automate.
ms.date: 01/14/2021
localization_priority: Normal
ms.openlocfilehash: b0f5b2f542216789f0d96f309cb7d799d201ba0f
ms.sourcegitcommit: e7e019ba36c2f49451ec08c71a1679eb6dba4268
ms.translationtype: MT
ms.contentlocale: de-DE
ms.lasthandoff: 01/22/2021
ms.locfileid: "49933266"
---
# <a name="troubleshooting-information-for-power-automate-with-office-scripts"></a>Problembehandlungsinformationen für Power Automate mit Office-Skripts

Mit Power Automate können Sie Ihre Office -Skriptautomatisierung auf die nächste Ebene bringen. Da Power Automate Skripts in Ihrem Auftrag in unabhängigen Excel-Sitzungen ausgeführt, sind jedoch einige wichtige Dinge zu beachten.

> [!TIP]
> Wenn Sie gerade mit der Verwendung von Office Scripts mit Power Automate beginnen, beginnen Sie mit der Ausführung von Office-Skripts [mit Power Automate,](../develop/power-automate-integration.md) um mehr über die Plattformen zu erfahren.

## <a name="avoid-using-relative-references"></a>Vermeiden von relativen Verweisen

Power Automate führt Ihr Skript in der ausgewählten Excel-Arbeitsmappe in Ihrem Auftrag aus. Die Arbeitsmappe wird in diesem Fall möglicherweise geschlossen. Jede API, die auf dem aktuellen Status des Benutzers basiert, z. B. , verhält sich `Workbook.getActiveWorksheet` in Power Automate möglicherweise anders. Dies liegt daran, dass die APIs auf einer relativen Position der Ansicht oder des Cursors des Benutzers basieren und dieser Verweis in einem Power Automate-Fluss nicht vorhanden ist.

Einige relative Referenz-APIs verursachen Fehler in Power Automate. Andere haben ein Standardverhalten, das den Status eines Benutzers impliziert. Achten Sie beim Entwerfen von Skripts darauf, absolute Verweise für Arbeitsblätter und Bereiche zu verwenden. Dadurch ist Ihr Power Automate-Fluss konsistent, auch wenn Arbeitsblätter neu angeordnet werden.

### <a name="script-methods-that-fail-when-run-power-automate-flows"></a>Skriptmethoden, die beim Ausführen von Power Automate-Flüssen fehlschlagen

Mit den folgenden Methoden wird ein Fehler ausgelöst, und beim Ausführen eines Skripts in einem Power Automate-Fluss tritt ein Fehler auf.

| Klasse | Methode |
|--|--|
| [Chart](/javascript/api/office-scripts/excelscript/excelscript.chart) | `activate` |
| [Range](/javascript/api/office-scripts/excelscript/excelscript.range) | `select` |
| [Workbook](/javascript/api/office-scripts/excelscript/excelscript.workbook) | `getActiveCell` |
| [Workbook](/javascript/api/office-scripts/excelscript/excelscript.workbook) | `getActiveChart` |
| [Workbook](/javascript/api/office-scripts/excelscript/excelscript.workbook) | `getActiveSlicer` |
| [Workbook](/javascript/api/office-scripts/excelscript/excelscript.workbook) | `getSelectedRange` |
| [Workbook](/javascript/api/office-scripts/excelscript/excelscript.workbook) | `getSelectedRanges` |

### <a name="script-methods-with-a-default-behavior-in-power-automate-flows"></a>Skriptmethoden mit Standardverhalten in Power Automate-Flüssen

Die folgenden Methoden verwenden ein Standardverhalten, und nicht den aktuellen Status eines Benutzers.

| Klasse | Methode | Verhalten von Power Automate |
|--|--|--|
| [Workbook](/javascript/api/office-scripts/excelscript/excelscript.workbook) | `getActiveWorksheet` | Gibt entweder das erste Arbeitsblatt in der Arbeitsmappe oder das Arbeitsblatt zurück, das derzeit von der Methode aktiviert `Worksheet.activate` wird. |
| [Arbeitsblatt](/javascript/api/office-scripts/excelscript/excelscript.worksheet) | `activate` | Markiert das Arbeitsblatt als aktives Arbeitsblatt für Zwecke von `Workbook.getActiveWorksheet` . |

## <a name="select-workbooks-with-the-file-browser-control"></a>Auswählen von Arbeitsmappen mit dem Dateibrowsersteuerelement

Beim Erstellen des **Schritts "Skript** ausführen" eines Power Automate-Ablaufs müssen Sie auswählen, welche Arbeitsmappe Teil des Ablaufs ist. Verwenden Sie den Dateibrowser, um Ihre Arbeitsmappe auszuwählen, anstatt den Namen der Arbeitsmappe manuell eingibt.

![Die Dateibrowseroption beim Erstellen einer Aktion "Skript ausführen" in Power Automate](../images/power-automate-file-browser.png)

Weitere Informationen zur Power Automate-Einschränkung und eine Diskussion über mögliche Problemumgehungen für die dynamische Auswahl von Arbeitsmappen finden Sie in diesem Thread in der [Microsoft Power Automate Community.](https://powerusers.microsoft.com/t5/Power-Automate-Ideas/Allow-for-dynamic-quot-file-quot-value-for-excel-quot-get-a-row/idi-p/103091#)

## <a name="time-zone-differences"></a>Zeitzonenunterschiede

Excel-Dateien verfügen nicht über einen inhärenten Speicherort oder eine Zeitzone. Jedes Mal, wenn ein Benutzer die Arbeitsmappe öffnet, verwendet seine Sitzung die lokale Zeitzone dieses Benutzers für Datumsberechnungen. Power Automate verwendet immer UTC.

Wenn Ihr Skript Datums- oder Zeitangaben verwendet, kann es Verhaltensunterschiede beim lokalen Test des Skripts oder beim Ausführen von Power Automate gibt. Mit Power Automate können Sie Zeiten konvertieren, formatieren und anpassen. Unter ["Arbeiten](https://flow.microsoft.com/blog/working-with-dates-and-times/) mit Datumsangaben und Uhrzeiten" in Ihren Flüssen finden Sie Anweisungen zur Verwendung dieser Funktionen in Power Automate und [ `main` Parametern:](../develop/power-automate-integration.md#main-parameters-passing-data-to-a-script) Übergeben von Daten an ein Skript, um zu erfahren, wie Sie diese Zeitinformationen für das Skript bereitstellen.

## <a name="see-also"></a>Siehe auch

- [Behandeln von Problemen mit Office-Skripts](troubleshooting.md)
- [Ausführen von Office-Skripts mit Power Automate](../develop/power-automate-integration.md)
- [Referenzdokumentation zu Excel Online (Business)-Connectors](/connectors/excelonlinebusiness/)
