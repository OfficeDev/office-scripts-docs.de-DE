---
title: Plattformbeschränkungen und-Anforderungen mit Office-Skripts
description: Ressourcengrenzwerte und Browserunterstützung für Office-Skripts bei Verwendung mit Excel im Internet
ms.date: 07/23/2020
localization_priority: Normal
ms.openlocfilehash: 6e297cba0b9f984f2d541cc3c441a666f9ebfcef
ms.sourcegitcommit: ff7fde04ce5a66d8df06ed505951c8111e2e9833
ms.translationtype: MT
ms.contentlocale: de-DE
ms.lasthandoff: 08/11/2020
ms.locfileid: "46618159"
---
# <a name="platform-limits-and-requirements-with-office-scripts"></a>Plattformbeschränkungen und-Anforderungen mit Office-Skripts

Es gibt einige Plattformeinschränkungen, die Sie beim Entwickeln von Office-Skripts beachten sollten. In diesem Artikel werden die Browserunterstützung und Daten Grenzwerte für Office-Skripts für Excel im Internet erläutert.

## <a name="browser-support"></a>Browserunterstützung

Office-Skripts funktionieren in jedem Browser, [der Office für das Internet unterstützt](https://support.microsoft.com/office/ad1303e0-a318-47aa-b409-d3a5eb44e452). Einige JavaScript-Funktionen werden in Internet Explorer 11 jedoch nicht unterstützt (IE 11). Alle in [ES6 oder höher](https://www.w3schools.com/Js/js_es6.asp) eingeführten Features funktionieren nicht mit IE 11. Wenn Personen in Ihrer Organisation diesen Browser weiterhin verwenden, müssen Sie Ihre Skripts in dieser Umgebung testen, wenn Sie Sie freigeben.

### <a name="third-party-cookies"></a>Drittanbieter-Cookies

Ihr Browser benötigt Cookies von Drittanbietern, die zum Anzeigen der Registerkarte " **automatisieren** " in Excel im Internet aktiviert sind. Überprüfen Sie Ihre Browsereinstellungen, wenn die Registerkarte nicht angezeigt wird. Wenn Sie eine private Browsersitzung verwenden, müssen Sie diese Einstellung möglicherweise jedes Mal erneut aktivieren.

> [!NOTE]
> Einige Browser bezeichnen diese Einstellung als "alle Cookies" anstelle von "Drittanbieter-Cookies".

## <a name="data-limits"></a>Beschränkungen für Daten

Es gibt Grenzen dafür, wie viel Excel-Daten gleichzeitig übertragen werden können und wie viele einzelne Power-Automatisierungs Transaktionen durchgeführt werden können.

### <a name="excel"></a>Excel

Excel für das Internet hat die folgenden Einschränkungen beim Ausführen von Aufrufen der Arbeitsmappe über ein Skript:

- Anforderungen und Antworten sind auf **5MB**limitiert.
- Ein Bereich ist auf **5 Millionen Zellen**limitiert.

Wenn beim Umgang mit großen Datasets Fehler auftreten, verwenden Sie mehrere kleinere Bereiche anstelle größerer Bereiche. Sie können auch APIs wie [Range. getSpecialCells](/javascript/api/office-scripts/excelscript/excelscript.range#getspecialcells-celltype--cellvaluetype-) , um bestimmte Zellen anstelle von großen Bereichen Ziel.

### <a name="power-automate"></a>Power Automate

Bei der Verwendung von Office-Skripts mit Power Automation sind Sie auf **200 Anrufe pro Tag**limitiert. Dieses Limit wird auf 12:00 Uhr UTC zurückgesetzt.

Die Power Automation-Plattform hat auch Nutzungseinschränkungen, die im Artikel [Limits and Configuration in Power Automation](/power-automate/limits-and-config)zu finden sind.

## <a name="see-also"></a>Siehe auch

- [Behandeln von Problemen mit Office-Skripts](troubleshooting.md)
- [Rückgängigmachen der Auswirkungen eines Office-Skripts](undo.md)
- [Verbessern der Leistung Ihrer Office-Skripts](../develop/web-client-performance.md)
- [Grundlegendes zur Skripterstellung für Office-Skripts in Excel im Internet](../develop/scripting-fundamentals.md)
