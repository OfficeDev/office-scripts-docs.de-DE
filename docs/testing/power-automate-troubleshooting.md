---
title: Behandeln von Office skripts, die in Power Automate
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
# <a name="troubleshoot-office-scripts-running-in-power-automate"></a>Behandeln von Office skripts, die in Power Automate

Power Automate können Sie Ihre Office Skriptautomatisierung auf die nächste Ebene bringen. Da Power Automate Skripts in Ihrem Namen in unabhängigen Sitzungen Excel, müssen Sie jedoch einige wichtige Dinge beachten.

> [!TIP]
> Wenn Sie gerade beginnen, Office Skripts mit Power Automate zu verwenden, beginnen Sie mit Ausführen von [Office Skripts](../develop/power-automate-integration.md) mit Power Automate, um mehr über die Plattformen zu erfahren.

## <a name="avoid-relative-references"></a>Vermeiden relativer Verweise

Power Automate führt Ihr Skript in der ausgewählten Excel in Ihrem Namen aus. In diesem Fall wird die Arbeitsmappe möglicherweise geschlossen. Jede API, die auf dem aktuellen Status des Benutzers basiert, z. B. , verhält sich in der `Workbook.getActiveWorksheet` Power Automate. Dies liegt daran, dass die APIs auf einer relativen Position der Ansicht oder des Cursors des Benutzers basieren und dieser Verweis nicht in einem Power Automate ist.

Einige relative Referenz-APIs verursachen Fehler in Power Automate. Andere haben ein Standardverhalten, das den Status eines Benutzers impliziert. Achten Sie beim Entwerfen ihrer Skripts darauf, absolute Verweise für Arbeitsblätter und Bereiche zu verwenden. Dadurch wird der Power Automate, auch wenn Arbeitsblätter neu angeordnet werden, konsistent.

### <a name="script-methods-that-fail-when-run-power-automate-flows"></a>Skriptmethoden, die beim Ausführen von Power Automate fehlschlagen

Die folgenden Methoden führen zu einem Fehler und einem Fehler, wenn sie von einem Skript in einem Power Automate werden.

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

Die folgenden Methoden verwenden ein Standardverhalten statt des aktuellen Status eines Benutzers.

| Klasse | Methode | Power Automate Verhalten |
|--|--|--|
| [Workbook](/javascript/api/office-scripts/excelscript/excelscript.workbook) | `getActiveWorksheet` | Gibt entweder das erste Arbeitsblatt in der Arbeitsmappe oder das Arbeitsblatt zurück, das derzeit von der Methode aktiviert `Worksheet.activate` wird. |
| [Arbeitsblatt](/javascript/api/office-scripts/excelscript/excelscript.worksheet) | `activate` | Markiert das Arbeitsblatt zu Zwecken von als aktives `Workbook.getActiveWorksheet` Arbeitsblatt. |

## <a name="select-workbooks-with-the-file-browser-control"></a>Auswählen von Arbeitsmappen mit dem Dateibrowsersteuerelement

Wenn Sie den **Schritt Skript ausführen** eines Power Automate erstellen, müssen Sie auswählen, welche Arbeitsmappe Teil des Ablaufs ist. Verwenden Sie den Dateibrowser, um Ihre Arbeitsmappe auszuwählen, anstatt den Namen der Arbeitsmappe manuell eintippen zu müssen.

:::image type="content" source="../images/power-automate-file-browser.png" alt-text="Die Power Automate Ausführen der Skriptaktion mit der Option Dateiauswahl anzeigen":::

Weitere Kontexte zur Power Automate und eine Diskussion über mögliche Problemumgehungen für die dynamische Auswahl von Arbeitsmappen finden Sie in diesem Thread [im Microsoft Power Automate Community](https://powerusers.microsoft.com/t5/Power-Automate-Ideas/Allow-for-dynamic-quot-file-quot-value-for-excel-quot-get-a-row/idi-p/103091#).

## <a name="time-zone-differences"></a>Zeitzonenunterschiede

Excel Dateien verfügen nicht über einen inhärenten Speicherort oder eine Zeitzone. Jedes Mal, wenn ein Benutzer die Arbeitsmappe öffnet, verwendet seine Sitzung die lokale Zeitzone dieses Benutzers für Datumsberechnungen. Power Automate verwendet immer UTC.

Wenn Ihr Skript Datums- oder Zeitangaben verwendet, können Verhaltensunterschiede auftreten, wenn das Skript lokal getestet wird, im Vergleich zum Ausführen von Power Automate. Power Automate können Sie Zeiten konvertieren, formatieren und anpassen. Anweisungen zur Verwendung dieser Funktionen in Power Automate und [ `main` Parameters: Übergeben](../develop/power-automate-integration.md#main-parameters-pass-data-to-a-script) von Daten an ein Skript finden Sie unter Working with Dates and Times inside of [your flows,](https://flow.microsoft.com/blog/working-with-dates-and-times/) um zu erfahren, wie Sie diese Zeitinformationen für das Skript bereitstellen.

## <a name="see-also"></a>Siehe auch

- [Problembehandlung Office Skripts](troubleshooting.md)
- [Ausführen Office Skripts mit Power Automate](../develop/power-automate-integration.md)
- [Excel Referenzdokumentation zu Onlineconnector (Business)](/connectors/excelonlinebusiness/)
