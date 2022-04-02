---
title: Problembehandlung bei Office Skripts, die in Power Automate ausgeführt werden
description: Tipps, Plattforminformationen und bekannte Probleme bei der Integration von Office Skripts und Power Automate.
ms.date: 11/01/2021
ms.localizationpriority: medium
ms.openlocfilehash: 2c256c2ddc64fcfc510f24e27662234f44b65ac0
ms.sourcegitcommit: 7023b9e23499806901a5ecf8ebc460b76887cca6
ms.translationtype: MT
ms.contentlocale: de-DE
ms.lasthandoff: 03/31/2022
ms.locfileid: "64586031"
---
# <a name="troubleshoot-office-scripts-running-in-power-automate"></a>Problembehandlung bei Office Skripts, die in Power Automate ausgeführt werden

Power Automate können Sie ihre Office Skriptautomatisierung auf die nächste Ebene bringen. Da Power Automate jedoch Skripts in Ihrem Auftrag in unabhängigen Excel Sitzungen ausführt, sind einige wichtige Dinge zu beachten.

> [!TIP]
> Wenn Sie gerade erst damit beginnen, Office Skripts mit Power Automate zu verwenden, beginnen Sie bitte mit ["Run Office Scripts with Power Automate](../develop/power-automate-integration.md)", um mehr über die Plattformen zu erfahren.

## <a name="avoid-relative-references"></a>Vermeiden von relativen Verweisen

Power Automate führt Ihr Skript in der ausgewählten Excel Arbeitsmappe in Ihrem Auftrag aus. In diesem Fall wird die Arbeitsmappe möglicherweise geschlossen. Jede API, die auf dem aktuellen Status des Benutzers basiert, z`Workbook.getActiveWorksheet`. B. , verhält sich in Power Automate möglicherweise anders. Dies liegt daran, dass die APIs auf einer relativen Position der Ansicht oder des Cursors des Benutzers basieren und dieser Verweis in einem Power Automate Fluss nicht vorhanden ist.

Einige relative Verweis-APIs lösen Fehler in Power Automate aus. Andere weisen ein Standardverhalten auf, das den Status eines Benutzers impliziert. Achten Sie beim Entwerfen ihrer Skripts darauf, absolute Verweise für Arbeitsblätter und Bereiche zu verwenden. Dadurch wird ihr Power Automate-Ablauf konsistent, auch wenn Arbeitsblätter neu angeordnet werden.

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
| [Workbook](/javascript/api/office-scripts/excelscript/excelscript.workbook) | `getActiveWorksheet` | Gibt entweder das erste Arbeitsblatt in der Arbeitsmappe oder das Arbeitsblatt zurück, das derzeit von der `Worksheet.activate` Methode aktiviert wird. |
| [Arbeitsblatt](/javascript/api/office-scripts/excelscript/excelscript.worksheet) | `activate` | Markiert das Arbeitsblatt als aktives Arbeitsblatt für die Zwecke von `Workbook.getActiveWorksheet`. |

## <a name="data-refresh-not-supported-in-power-automate"></a>Datenaktualisierung in Power Automate nicht unterstützt

Office Skripts können Daten nicht aktualisieren, wenn sie in Power Automate ausgeführt werden. Methoden, z `PivotTable.refresh` . B. "Nichts tun", wenn sie in einem Fluss aufgerufen werden. Darüber hinaus löst Power Automate keine Datenaktualisierung für Formeln aus, die Arbeitsmappenlinks verwenden.

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

Weitere Informationen zur Power Automate Einschränkung und eine Erläuterung möglicher Problemumgehungen für die dynamische Auswahl von Arbeitsmappen finden Sie [in diesem Thread im Microsoft Power Automate Community](https://powerusers.microsoft.com/t5/Power-Automate-Ideas/Allow-for-dynamic-quot-file-quot-value-for-excel-quot-get-a-row/idi-p/103091#).

## <a name="pass-entire-arrays-as-script-parameters"></a>Übergeben ganzer Arrays als Skriptparameter

Power Automate ermöglicht Es Benutzern, Arrays als Variable oder als einzelne Elemente im Array an Connectors zu übergeben. Standardmäßig werden einzelne Elemente übergeben, die das Array im Fluss erstellen. Für Skripts oder andere Connectors, die ganze Arrays als Argumente verwenden, müssen Sie den Switch auswählen, **um die gesamte Arrayschaltfläche einzugeben** , um das Array als ein vollständiges Objekt zu übergeben. Diese Schaltfläche befindet sich in der oberen rechten Ecke jedes Eingabefelds für Arrayparameter.

:::image type="content" source="../images/combine-worksheets-flow-3.png" alt-text="Die Schaltfläche, um zu wechseln, um ein gesamtes Array in ein Eingabefeld eines Steuerelementfelds einzugeben.":::

## <a name="time-zone-differences"></a>Zeitzonenunterschiede

Excel Dateien haben keinen inhärenten Speicherort oder keine Zeitzone. Jedes Mal, wenn ein Benutzer die Arbeitsmappe öffnet, verwendet seine Sitzung die lokale Zeitzone dieses Benutzers für Datumsberechnungen. Power Automate verwendet immer UTC.

Wenn Ihr Skript Datumsangaben oder Uhrzeiten verwendet, kann es Verhaltensunterschiede geben, wenn das Skript lokal getestet wird und nicht, wenn es Power Automate ausgeführt wird. mit Power Automate können Sie Zeiten konvertieren, formatieren und anpassen. Anweisungen zur Verwendung dieser Funktionen in Power Automate und [`main` Parametern finden](../develop/power-automate-integration.md#main-parameters-pass-data-to-a-script) Sie unter [Working with Dates and Times inside of your flows](https://flow.microsoft.com/blog/working-with-dates-and-times/): Pass data to a script to learn how to provide that time information for the script.

## <a name="script-parameter-fields-or-returned-output-not-appearing-in-power-automate"></a>Skriptparameterfelder oder zurückgegebene Ausgabe, die nicht in Power Automate

Es gibt zwei Gründe, warum die Parameter oder zurückgegebenen Daten eines Skripts im Power Automate Fluss-Generator nicht korrekt wiedergegeben werden.

- Die Skriptsignatur (die Parameter oder der Rückgabewert) wurde seit dem Hinzufügen des **Excel Business (Online)**-Connectors geändert.
- Die Skriptsignatur verwendet nicht unterstützte Typen. Überprüfen Sie Ihre Typen anhand der Listen unter den [Parametern](../develop/power-automate-integration.md#main-parameters-pass-data-to-a-script) und [gibt](../develop/power-automate-integration.md#return-data-from-a-script) Abschnitte von [Run Office Scripts mit Power Automate](../develop/power-automate-integration.md) Artikel zurück.

Die Signatur eines Skripts wird beim Erstellen mit dem **Excel Business (Online)**-Connector gespeichert. Entfernen Sie den alten Connector, und erstellen Sie einen neuen, um die neuesten Parameter abzurufen und Werte für die **Skriptaktion ausführen** zurückzugeben.

## <a name="see-also"></a>Weitere Informationen

- [Problembehandlung bei Office Skripts](troubleshooting.md)
- [Ausführen Office Skripts mit Power Automate](../develop/power-automate-integration.md)
- [Referenzdokumentation für Excel Online connector (Business)](/connectors/excelonlinebusiness/)
