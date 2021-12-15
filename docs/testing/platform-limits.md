---
title: Plattformbeschränkungen und -anforderungen mit Office-Skripts
description: Ressourcenbeschränkungen und Browserunterstützung für Office Skripts bei Verwendung mit Excel im Web
ms.date: 12/14/2021
ms.localizationpriority: medium
ms.openlocfilehash: e0c2c8108bb36dc979b33e195cdc0e0e17a586ba
ms.sourcegitcommit: 1e561d245d9ab7f2529a1b071b1fa0326e37c4ca
ms.translationtype: MT
ms.contentlocale: de-DE
ms.lasthandoff: 12/15/2021
ms.locfileid: "61513629"
---
# <a name="platform-limits-and-requirements-with-office-scripts"></a>Plattformbeschränkungen und -anforderungen mit Office-Skripts

Es gibt einige Plattformeinschränkungen, die Sie beim Entwickeln von Office Skripts beachten sollten. In diesem Artikel werden die Browserunterstützung und Datenbeschränkungen für Office Skripts für Excel im Web beschrieben.

## <a name="browser-support"></a>Browserunterstützung

Office Skripts funktionieren in jedem Browser, [der Office für das Web unterstützt.](https://support.microsoft.com/office/ad1303e0-a318-47aa-b409-d3a5eb44e452) Einige JavaScript-Features werden in Internet Explorer 11 (IE 11) jedoch nicht unterstützt. Alle in [ES6 oder höher](https://www.w3schools.com/Js/js_es6.asp) eingeführten Features funktionieren nicht mit IE 11. Wenn Personen in Ihrer Organisation diesen Browser weiterhin verwenden, sollten Sie Ihre Skripts in dieser Umgebung testen, wenn Sie sie freigeben.

[!INCLUDE [Teams support note](../includes/teams-support-note.md)]

### <a name="third-party-cookies"></a>Cookies von Drittanbietern

Ihr Browser benötigt Cookies von Drittanbietern, die aktiviert sind, um die Registerkarte **"Automatisieren"** in Excel im Web anzuzeigen. Überprüfen Sie die Browsereinstellungen, wenn die Registerkarte nicht angezeigt wird. Wenn Sie eine private Browsersitzung verwenden, müssen Sie diese Einstellung möglicherweise jedes Mal erneut aktivieren.

> [!NOTE]
> Einige Browser bezeichnen diese Einstellung als "alle Cookies" anstelle von "Cookies von Drittanbietern".

#### <a name="instructions-for-adjusting-cookie-settings-in-popular-browsers"></a>Anweisungen zum Anpassen von Cookieeinstellungen in beliebten Browsern

- [Chrome](https://support.google.com/chrome/answer/95647)
- [Microsoft Edge](https://support.microsoft.com/microsoft-edge/597f04f2-c0ce-f08c-7c2b-541086362bd2)
- [Firefox](https://support.mozilla.org/kb/disable-third-party-cookies)
- [Safari](https://support.apple.com/guide/safari/manage-cookies-and-website-data-sfri11471/mac)

## <a name="data-limits"></a>Beschränkungen für Daten

Es gibt Beschränkungen, wie viele Excel Daten gleichzeitig übertragen werden können und wie viele einzelne Power Automate Transaktionen durchgeführt werden können.

### <a name="excel"></a>Excel

bei Aufrufen der Arbeitsmappe über ein Skript gelten für Excel für das Web die folgenden Einschränkungen:

- Anforderungen und Antworten sind auf **5 MB** beschränkt.
- Ein Bereich ist auf **fünf Millionen Zellen** begrenzt.

Wenn beim Umgang mit großen Datasets Fehler auftreten, versuchen Sie, mehrere kleinere Bereiche anstelle von größeren Bereichen zu verwenden. Ein Beispiel finden Sie im Beispiel zum [Schreiben eines großen Datasets.](../resources/samples/write-large-dataset.md) Sie können auch APIs wie [Range.getSpecialCells](/javascript/api/office-scripts/excelscript/excelscript.range#getSpecialCells_cellType__cellValueType_) verwenden, um bestimmte Zellen anstelle großer Bereiche als Ziel zu verwenden.

### <a name="power-automate"></a>Power Automate

Wenn sie Office Skripts mit Power Automate verwenden, ist jeder Benutzer auf **800 Aufrufe der Aktion "Skript ausführen" pro Tag** beschränkt. Dieser Grenzwert wird um 12:00 Uhr UTC zurückgesetzt.

Die Power Automate-Plattform weist ebenfalls Nutzungseinschränkungen auf, die in den folgenden Artikeln zu finden sind.

- [Grenzwerte und Konfiguration in Power Automate](/power-automate/limits-and-config)
- [Bekannte Probleme und Einschränkungen für den connector Excel Online (Business)](/connectors/excelonlinebusiness/#known-issues-and-limitations)

> [!NOTE]
> Wenn Sie über ein langes Skript verfügen, beachten Sie das Timeout von [120 Sekunden für synchrone Power Automate Vorgänge.](/power-automate/limits-and-config#timeout) Sie müssen entweder [Ihr Skript optimieren](../develop/web-client-performance.md) oder Ihre Excel Automatisierung in mehrere Skripts aufteilen.

## <a name="see-also"></a>Weitere Artikel

- [Problembehandlung bei Office Skripts](troubleshooting.md)
- [Auswirkungen von Office-Skripts rückgängig machen](undo.md)
- [Verbessern der Leistung Ihrer Office-Skripts](../develop/web-client-performance.md)
- [Grundlagen der Skripterstellung für Office Skripts in Excel im Web](../develop/scripting-fundamentals.md)
