---
title: Problembehandlungsinformationen für Power Automate mit Office-Skripts
description: Tipps, Plattforminformationen und bekannte Probleme bei der Integration zwischen Office-Skripts und Power Automate.
ms.date: 01/14/2021
localization_priority: Normal
ms.openlocfilehash: 59f4cd8b3476c2ee2a1a862f136173a543ba8a15
ms.sourcegitcommit: 45ffe3dbd2c834b78592ad35928cf8096f5e80bc
ms.translationtype: MT
ms.contentlocale: de-DE
ms.lasthandoff: 04/14/2021
ms.locfileid: "51755007"
---
# <a name="troubleshooting-information-for-power-automate-with-office-scripts"></a>Problembehandlungsinformationen für Power Automate mit Office-Skripts

Mit Power Automate können Sie Ihre Office Script-Automatisierung auf die nächste Ebene bringen. Da Power Automate skripts jedoch in Ihrem Auftrag in unabhängigen Excel-Sitzungen ausgeführt, sind einige wichtige Dinge zu beachten.

> [!TIP]
> Wenn Sie gerade mit der Verwendung von Office-Skripts mit Power Automate beginnen, beginnen Sie bitte mit Ausführen von [Office-Skripts mit Power Automate,](../develop/power-automate-integration.md) um mehr über die Plattformen zu erfahren.

## <a name="avoid-using-relative-references"></a>Vermeiden der Verwendung relativer Verweise

Power Automate führt Ihr Skript in der ausgewählten Excel-Arbeitsmappe in Ihrem Auftrag aus. In diesem Fall wird die Arbeitsmappe möglicherweise geschlossen. Jede API, die auf dem aktuellen Status des Benutzers basiert, z. B. , verhält sich `Workbook.getActiveWorksheet` in Power Automate möglicherweise anders. Dies liegt daran, dass die APIs auf einer relativen Position der Ansicht oder des Cursors des Benutzers basieren und dieser Verweis in einem Power Automate-Fluss nicht vorhanden ist.

Einige relative Referenz-APIs verursachen Fehler in Power Automate. Andere haben ein Standardverhalten, das den Status eines Benutzers impliziert. Achten Sie beim Entwerfen ihrer Skripts darauf, absolute Verweise für Arbeitsblätter und Bereiche zu verwenden. Dadurch ist der Power Automate-Fluss konsistent, auch wenn Arbeitsblätter neu angeordnet werden.

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

### <a name="script-methods-with-a-default-behavior-in-power-automate-flows"></a>Skriptmethoden mit einem Standardverhalten in Power Automate-Flüssen

Die folgenden Methoden verwenden ein Standardverhalten statt des aktuellen Status eines Benutzers.

| Klasse | Methode | Power Automate-Verhalten |
|--|--|--|
| [Workbook](/javascript/api/office-scripts/excelscript/excelscript.workbook) | `getActiveWorksheet` | Gibt entweder das erste Arbeitsblatt in der Arbeitsmappe oder das Arbeitsblatt zurück, das derzeit von der Methode aktiviert `Worksheet.activate` wird. |
| [Arbeitsblatt](/javascript/api/office-scripts/excelscript/excelscript.worksheet) | `activate` | Markiert das Arbeitsblatt zu Zwecken von als aktives `Workbook.getActiveWorksheet` Arbeitsblatt. |

## <a name="select-workbooks-with-the-file-browser-control"></a>Auswählen von Arbeitsmappen mit dem Dateibrowsersteuerelement

Beim Erstellen des **Run-Skriptschritts** eines Power Automate-Fluss müssen Sie auswählen, welche Arbeitsmappe Teil des Datenflusses ist. Verwenden Sie den Dateibrowser, um Ihre Arbeitsmappe auszuwählen, anstatt den Namen der Arbeitsmappe manuell eintippen zu müssen.

:::image type="content" source="../images/power-automate-file-browser.png" alt-text="Die Power Automate Run-Skriptaktion mit der Option Dateiauswahl anzeigen.":::

Weitere Informationen zur Einschränkung von Power Automate und eine Diskussion über mögliche Problemumgehungen für die dynamische Auswahl von Arbeitsmappen finden Sie in diesem Thread in der [Microsoft Power Automate Community](https://powerusers.microsoft.com/t5/Power-Automate-Ideas/Allow-for-dynamic-quot-file-quot-value-for-excel-quot-get-a-row/idi-p/103091#).

## <a name="time-zone-differences"></a>Zeitzonenunterschiede

Excel-Dateien verfügen nicht über einen inhärenten Speicherort oder eine Zeitzone. Jedes Mal, wenn ein Benutzer die Arbeitsmappe öffnet, verwendet seine Sitzung die lokale Zeitzone dieses Benutzers für Datumsberechnungen. Power Automate verwendet immer UTC.

Wenn Ihr Skript Datums- oder Zeitangaben verwendet, kann es zu Verhaltensunterschieden beim lokalen Test des Skripts im Vergleich zum Ausführen von Power Automate führen. Mit Power Automate können Sie Zeiten konvertieren, formatieren und anpassen. Anweisungen zur Verwendung dieser Funktionen in Power Automate und [ `main` Parameters: Passing data to a script](../develop/power-automate-integration.md#main-parameters-passing-data-to-a-script) finden Sie unter Working with Dates and Times inside of your [flows,](https://flow.microsoft.com/blog/working-with-dates-and-times/) um zu erfahren, wie Sie diese Zeitinformationen für das Skript bereitstellen.

## <a name="see-also"></a>Weitere Informationen

- [Behandeln von Problemen mit Office-Skripts](troubleshooting.md)
- [Ausführen von Office-Skripts mit Power Automate](../develop/power-automate-integration.md)
- [Referenzdokumentation zu Excel Online (Business)-Connectors](/connectors/excelonlinebusiness/)
