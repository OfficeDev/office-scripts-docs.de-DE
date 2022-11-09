---
title: Plattformlimits und -anforderungen mit Office-Skripts
description: Ressourcenlimits und Browserunterstützung für Office-Skripts bei Verwendung mit Excel im Web.
ms.date: 11/07/2022
ms.localizationpriority: medium
ms.openlocfilehash: 764d1eddaf303a941a098ec1d3f3056d63e8693f
ms.sourcegitcommit: 7cadf2b637bf62874e43b6e595286101816662aa
ms.translationtype: MT
ms.contentlocale: de-DE
ms.lasthandoff: 11/09/2022
ms.locfileid: "68891246"
---
# <a name="platform-limits-and-requirements-with-office-scripts"></a>Plattformlimits und -anforderungen mit Office-Skripts

Es gibt einige Plattformbeschränkungen, die Sie beim Entwickeln von Office-Skripts beachten sollten. In diesem Artikel werden die Browserunterstützung und Datengrenzwerte für Office-Skripts für Excel im Web beschrieben.

## <a name="browser-support"></a>Browserunterstützung

Office-Skripts funktionieren in jedem Browser, der [Office für das Web unterstützt](https://support.microsoft.com/office/ad1303e0-a318-47aa-b409-d3a5eb44e452). Einige JavaScript-Features werden in Internet Explorer 11 (IE 11) jedoch nicht unterstützt. Alle in [ES6 oder höher eingeführten](https://www.w3schools.com/Js/js_es6.asp) Features funktionieren nicht mit IE 11. Wenn Personen in Ihrer Organisation diesen Browser weiterhin verwenden, sollten Sie Ihre Skripts in dieser Umgebung testen, wenn Sie sie freigeben.

[!INCLUDE [Teams support note](../includes/teams-support-note.md)]

### <a name="third-party-cookies"></a>Cookies von Drittanbietern

Ihr Browser muss Cookies von Drittanbietern aktivieren, um die Registerkarte **Automatisieren** in Excel im Web anzuzeigen. Überprüfen Sie die Browsereinstellungen, wenn die Registerkarte nicht angezeigt wird. Wenn Sie eine private Browsersitzung verwenden, müssen Sie diese Einstellung möglicherweise jedes Mal erneut aktivieren.

> [!NOTE]
> Einige Browser bezeichnen diese Einstellung als "alle Cookies" anstelle von "Cookies von Drittanbietern".

#### <a name="instructions-for-adjusting-cookie-settings-in-popular-browsers"></a>Anweisungen zum Anpassen von Cookieeinstellungen in beliebten Browsern

- [Chrome](https://support.google.com/chrome/answer/95647)
- [Microsoft Edge](https://support.microsoft.com/microsoft-edge/597f04f2-c0ce-f08c-7c2b-541086362bd2)
- [Firefox](https://support.mozilla.org/kb/disable-third-party-cookies)
- [Safari](https://support.apple.com/guide/safari/manage-cookies-and-website-data-sfri11471/mac)

## <a name="data-limits"></a>Beschränkungen für Daten

Es gibt Beschränkungen, wie viele Excel-Daten gleichzeitig übertragen werden können und wie viele einzelne Power Automate-Transaktionen durchgeführt werden können.

### <a name="excel"></a>Excel

Excel für das Web weist die folgenden Einschränkungen auf, wenn Aufrufe an die Arbeitsmappe über ein Skript ausgeführt werden:

- Anforderungen und Antworten sind auf **5 MB** beschränkt.
- Ein Bereich ist auf **fünf Millionen Zellen** begrenzt.

Wenn beim Umgang mit großen Datasets Fehler auftreten, versuchen Sie, mehrere kleinere Bereiche anstelle größerer Bereiche zu verwenden. Ein Beispiel finden Sie im Beispiel [Schreiben eines großen Datasets](../resources/samples/write-large-dataset.md) . Sie können auch APIs wie [Range.getSpecialCells](/javascript/api/office-scripts/excelscript/excelscript.range#excelscript-excelscript-range-getspecialcells-member(1)) verwenden, um bestimmte Zellen anstelle großer Bereiche als Ziel zu verwenden.

Excel-Grenzwerte, die nicht spezifisch für Office-Skripts sind, finden Sie im Artikel [Excel-Spezifikationen und -Grenzwerte](https://support.microsoft.com/office/excel-specifications-and-limits-1672b34d-7043-467e-8e27-269d656771c3).

### <a name="power-automate"></a>Power Automate

Bei Verwendung von Office-Skripts mit Power Automate ist jeder Benutzer auf **1.600 Aufrufe der Aktion Skript ausführen pro Tag** beschränkt. Dieser Grenzwert wird um 12:00 Uhr UTC zurückgesetzt.

Die Power Automate-Plattform weist auch Nutzungseinschränkungen auf, die in den folgenden Artikeln zu finden sind.

- [Grenzwerte und Konfiguration in Power Automate](/power-automate/limits-and-config)
- [Bekannte Probleme und Einschränkungen für den Excel Online (Business)-Connector](/connectors/excelonlinebusiness/#known-issues-and-limitations)

> [!NOTE]
> Wenn Sie über ein Skript mit langer Ausführungsdauer verfügen, beachten Sie das [Timeout von 120 Sekunden für synchrone Power Automate-Vorgänge](/power-automate/limits-and-config#timeout). Sie müssen entweder [Ihr Skript optimieren](../develop/web-client-performance.md) oder Ihre Excel-Automatisierung in mehrere Skripts aufteilen.

## <a name="see-also"></a>Siehe auch

- [Spezifikationen und Beschränkungen in Excel](https://support.microsoft.com/office/excel-specifications-and-limits-1672b34d-7043-467e-8e27-269d656771c3)
- [Problembehandlung bei Office-Skripts](troubleshooting.md)
- [Auswirkungen von Office-Skripts rückgängig machen](undo.md)
- [Verbessern der Leistung Ihrer Office-Skripts](../develop/web-client-performance.md)
