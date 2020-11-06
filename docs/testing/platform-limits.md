---
title: Plattformbeschränkungen und-Anforderungen mit Office-Skripts
description: Ressourcengrenzwerte und Browserunterstützung für Office-Skripts bei Verwendung mit Excel im Internet
ms.date: 10/23/2020
localization_priority: Normal
ms.openlocfilehash: 61f5c55be278ae056014d3b01e4176354d913f87
ms.sourcegitcommit: d3e7681e262bdccc281fcb7b3c719494202e846b
ms.translationtype: MT
ms.contentlocale: de-DE
ms.lasthandoff: 11/06/2020
ms.locfileid: "48930078"
---
# <a name="platform-limits-and-requirements-with-office-scripts"></a>Plattformbeschränkungen und-Anforderungen mit Office-Skripts

Es gibt einige Plattformeinschränkungen, die Sie beim Entwickeln von Office-Skripts beachten sollten. In diesem Artikel werden die Browserunterstützung und Daten Grenzwerte für Office-Skripts für Excel im Internet erläutert.

## <a name="browser-support"></a>Browserunterstützung

Office-Skripts funktionieren in jedem Browser, [der Office für das Internet unterstützt](https://support.microsoft.com/office/ad1303e0-a318-47aa-b409-d3a5eb44e452). Einige JavaScript-Funktionen werden in Internet Explorer 11 jedoch nicht unterstützt (IE 11). Alle in [ES6 oder höher](https://www.w3schools.com/Js/js_es6.asp) eingeführten Features funktionieren nicht mit IE 11. Wenn Personen in Ihrer Organisation diesen Browser weiterhin verwenden, müssen Sie Ihre Skripts in dieser Umgebung testen, wenn Sie Sie freigeben.

[!INCLUDE [Teams support note](../includes/teams-support-note.md)]

### <a name="third-party-cookies"></a>Drittanbieter-Cookies

Ihr Browser benötigt Cookies von Drittanbietern, die zum Anzeigen der Registerkarte " **automatisieren** " in Excel im Internet aktiviert sind. Überprüfen Sie Ihre Browsereinstellungen, wenn die Registerkarte nicht angezeigt wird. Wenn Sie eine private Browsersitzung verwenden, müssen Sie diese Einstellung möglicherweise jedes Mal erneut aktivieren.

> [!NOTE]
> Einige Browser bezeichnen diese Einstellung als "alle Cookies" anstelle von "Drittanbieter-Cookies".

#### <a name="instructions-for-adjusting-cookie-settings-in-popular-browsers"></a>Anweisungen zum Anpassen von Cookieeinstellungen in gängigen Browsern

- [Chrome](https://support.google.com/chrome/answer/95647)
- [Microsoft Edge](https://support.microsoft.com/microsoft-edge/temporarily-allow-cookies-and-site-data-in-microsoft-edge-597f04f2-c0ce-f08c-7c2b-541086362bd2)
- [Firefox](https://support.mozilla.org/kb/disable-third-party-cookies)
- [Safari](https://support.apple.com/guide/safari/manage-cookies-and-website-data-sfri11471/mac)

## <a name="data-limits"></a>Beschränkungen für Daten

Es gibt Grenzen dafür, wie viel Excel-Daten gleichzeitig übertragen werden können und wie viele einzelne Power-Automatisierungs Transaktionen durchgeführt werden können.

### <a name="excel"></a>Excel

Excel für das Internet hat die folgenden Einschränkungen beim Ausführen von Aufrufen der Arbeitsmappe über ein Skript:

- Anforderungen und Antworten sind auf **5MB** limitiert.
- Ein Bereich ist auf **5 Millionen Zellen** limitiert.

Wenn beim Umgang mit großen Datasets Fehler auftreten, verwenden Sie mehrere kleinere Bereiche anstelle größerer Bereiche. Sie können auch APIs wie [Range. getSpecialCells](/javascript/api/office-scripts/excelscript/excelscript.range#getspecialcells-celltype--cellvaluetype-) , um bestimmte Zellen anstelle von großen Bereichen Ziel.

### <a name="power-automate"></a>Power Automate

Bei der Verwendung von Office-Skripts mit Power Automation sind Sie auf **200 Anrufe pro Tag** limitiert. Dieses Limit wird auf 12:00 Uhr UTC zurückgesetzt.

Die Power Automation-Plattform hat auch Nutzungseinschränkungen, die im Artikel [Limits and Configuration in Power Automation](/power-automate/limits-and-config)zu finden sind.

## <a name="see-also"></a>Siehe auch

- [Behandeln von Problemen mit Office-Skripts](troubleshooting.md)
- [Rückgängigmachen der Auswirkungen eines Office-Skripts](undo.md)
- [Verbessern der Leistung Ihrer Office-Skripts](../develop/web-client-performance.md)
- [Grundlegendes zur Skripterstellung für Office-Skripts in Excel im Internet](../develop/scripting-fundamentals.md)
