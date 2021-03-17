---
title: Plattformbeschränkungen und -anforderungen mit Office-Skripts
description: Ressourcenbeschränkungen und Browserunterstützung für Office-Skripts bei Verwendung mit Excel im Web
ms.date: 03/12/2021
localization_priority: Normal
ms.openlocfilehash: 93307b6204f409f26c77b5ead33188205d5c4b4d
ms.sourcegitcommit: 5bde455b06ee2ed007f3e462d8ad485b257774ef
ms.translationtype: MT
ms.contentlocale: de-DE
ms.lasthandoff: 03/17/2021
ms.locfileid: "50837265"
---
# <a name="platform-limits-and-requirements-with-office-scripts"></a>Plattformbeschränkungen und -anforderungen mit Office-Skripts

Es gibt einige Plattformeinschränkungen, die Sie bei der Entwicklung von Office-Skripts beachten sollten. In diesem Artikel werden die Browserunterstützung und Datenbeschränkungen für Office-Skripts für Excel im Web erläutert.

## <a name="browser-support"></a>Browserunterstützung

Office-Skripts funktionieren in jedem Browser, [der Office für das Web unterstützt.](https://support.microsoft.com/office/ad1303e0-a318-47aa-b409-d3a5eb44e452) Einige JavaScript-Features werden jedoch in Internet Explorer 11 (IE 11) nicht unterstützt. In ES6 oder [höher eingeführte](https://www.w3schools.com/Js/js_es6.asp) Features funktionieren nicht mit IE 11. Wenn Personen in Ihrer Organisation diesen Browser weiterhin verwenden, sollten Sie ihre Skripts in dieser Umgebung testen, wenn Sie sie freigeben.

[!INCLUDE [Teams support note](../includes/teams-support-note.md)]

### <a name="third-party-cookies"></a>Cookies von Drittanbietern

Ihr Browser benötigt Cookies von Drittanbietern, die aktiviert sind, um die Registerkarte **Automatisieren** in Excel im Web anzuzeigen. Überprüfen Sie die Browsereinstellungen, wenn die Registerkarte nicht angezeigt wird. Wenn Sie eine private Browsersitzung verwenden, müssen Sie diese Einstellung möglicherweise jedes Mal erneut aktivieren.

> [!NOTE]
> In einigen Browsern wird diese Einstellung als "alle Cookies" anstelle von "Cookies von Drittanbietern" bezeichnet.

#### <a name="instructions-for-adjusting-cookie-settings-in-popular-browsers"></a>Anweisungen zum Anpassen von Cookieeinstellungen in beliebten Browsern

- [Chrome](https://support.google.com/chrome/answer/95647)
- [Microsoft Edge](https://support.microsoft.com/microsoft-edge/temporarily-allow-cookies-and-site-data-in-microsoft-edge-597f04f2-c0ce-f08c-7c2b-541086362bd2)
- [Firefox](https://support.mozilla.org/kb/disable-third-party-cookies)
- [Safari](https://support.apple.com/guide/safari/manage-cookies-and-website-data-sfri11471/mac)

## <a name="data-limits"></a>Beschränkungen für Daten

Es gibt Beschränkungen, wie viele Excel-Daten gleichzeitig übertragen werden können und wie viele einzelne Power Automate-Transaktionen durchgeführt werden können.

### <a name="excel"></a>Excel

Excel für das Web hat die folgenden Einschränkungen, wenn Aufrufe der Arbeitsmappe über ein Skript ausgeführt werden:

- Anforderungen und Antworten sind auf **5 MB beschränkt.**
- Ein Bereich ist auf **fünf Millionen Zellen beschränkt.**

Wenn beim Umgang mit großen Datasets Fehler auftreten, versuchen Sie, mehrere kleinere Bereiche anstelle größerer Bereiche zu verwenden. Sie können apIs wie [Range.getSpecialCells](/javascript/api/office-scripts/excelscript/excelscript.range#getspecialcells-celltype--cellvaluetype-) auch auf bestimmte Zellen statt auf große Bereiche zielen.

### <a name="power-automate"></a>Power Automate

Bei verwendung von Office Scripts mit Power Automate ist jeder Benutzer auf **200** Anrufe pro Tag beschränkt. Dieser Grenzwert wird um 12:00 Uhr UTC zurückgesetzt.

Die Power Automate-Plattform hat auch Nutzungseinschränkungen, die in den folgenden Artikeln zu finden sind:

- [Grenzwerte und Konfiguration in Power Automate](/power-automate/limits-and-config)
- [Bekannte Probleme und Einschränkungen für den Excel Online (Business)-Connector](/connectors/excelonlinebusiness/#known-issues-and-limitations)

## <a name="see-also"></a>Siehe auch

- [Behandeln von Problemen mit Office-Skripts](troubleshooting.md)
- [Rückgängigmachen der Auswirkungen eines Office-Skripts](undo.md)
- [Verbessern der Leistung Ihrer Office-Skripts](../develop/web-client-performance.md)
- [Skripting-Grundlagen für Office-Skripts in Excel im Web](../develop/scripting-fundamentals.md)
