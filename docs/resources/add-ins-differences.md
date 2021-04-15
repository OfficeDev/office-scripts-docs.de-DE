---
title: Unterschiede zwischen Office-Skripts und Office-Add-Ins
description: Das Verhalten und die API-Unterschiede zwischen Office-Skripts und Office-Add-Ins.
ms.date: 06/01/2020
localization_priority: Normal
ms.openlocfilehash: 96af98ca9f247406c5cc916f38892c318d33c560
ms.sourcegitcommit: 45ffe3dbd2c834b78592ad35928cf8096f5e80bc
ms.translationtype: MT
ms.contentlocale: de-DE
ms.lasthandoff: 04/14/2021
ms.locfileid: "51755098"
---
# <a name="differences-between-office-scripts-and-office-add-ins"></a>Unterschiede zwischen Office-Skripts und Office-Add-Ins

Office-Add-Ins und Office-Skripts haben viele gemeinsam. Beide bieten eine automatisierte Steuerung einer Excel-Arbeitsmappe mit einer JavaScript-API. Die Office Scripts-APIs sind jedoch eine spezielle, synchrone Version der Office-JavaScript-API.

:::image type="content" source="../images/office-programmability-diagram.png" alt-text="Ein Diagramm mit vier Quadranten, das die Fokusbereiche für verschiedene Office-Erweiterbarkeitslösungen zeigt. Sowohl Office-Skripts als auch Office Web-Add-Ins konzentrieren sich auf das Web und die Zusammenarbeit, aber Office-Skripts sind für Endbenutzer (während Office Web-Add-Ins auf professionelle Entwickler ausgerichtet sind).":::

Office-Skripts werden mit einem manuellen Tastendruck oder als Schritt in [Power Automate](https://flow.microsoft.com/)ausgeführt, während Office-Add-Ins beibehalten werden, während ihre Aufgabenbereiche geöffnet sind. Dies bedeutet, dass die Add-Ins während einer Sitzung den Status beibehalten können, während Office-Skripts keinen internen Status zwischen den Einzelnen Ausführen beibehalten. Wenn Sie feststellen, dass Ihre Excel-Erweiterung die Funktionen der Skriptplattform überschreiten muss, lesen Sie die [Office-Add-Ins-Dokumentation,](/office/dev/add-ins) um mehr über Office-Add-Ins zu erfahren.

Der Rest dieses Artikels beschreibt die wichtigsten Unterschiede zwischen Office-Add-Ins und Office-Skripts.

## <a name="platform-support"></a>Plattformunterstützung

Office-Add-Ins sind plattformübergreifender. Sie funktionieren auf allen Windows-Desktop-, Mac-, iOS- und Webplattformen und bieten die gleiche Benutzererfahrung. Eine Ausnahme davon ist in der Dokumentation der einzelnen API vermerkt.

Office-Skripts werden derzeit nur für Excel im Web unterstützt. Die Aufzeichnung, Bearbeitung und Ausführung erfolgt auf der Webplattform.

## <a name="apis"></a>APIs

Es gibt keine synchrone Version der Office-JavaScript-APIs für Office-Add-Ins. Die standardmäßigen Office Scripts-APIs sind für die Plattform einzigartig und verfügen über zahlreiche Optimierungen und Änderungen, um die Verwendung des Paradigmas zu `load` / `sync` vermeiden.

Einige der [Excel-JavaScript-APIs](/javascript/api/excel?view=excel-js-preview&preserve-view=true) sind mit den [Office Scripts Async-APIs kompatibel.](../develop/excel-async-model.md) Einige Beispiele und Add-In-Codeblöcke können in Blöcke mit `Excel.run` minimaler Übersetzung portiert werden. Während die beiden Plattformen die Funktionalität gemeinsam nutzen, gibt es Lücken. Die beiden wichtigsten API-Sätze, über die Office-Add-Ins verfügen, office-Skripts sind jedoch keine Ereignisse und die allgemeinen APIs.

### <a name="events"></a>Ereignisse

Office-Skripts unterstützen keine [Ereignisse.](/office/dev/add-ins/excel/excel-add-ins-events) Jedes Skript führt den Code in einer einzelnen `main` Methode aus und endet dann. Es wird nicht reaktiviert, wenn Ereignisse ausgelöst werden, und kann daher keine Ereignisse registrieren.

### <a name="common-apis"></a>Allgemeine APIs

Office-Skripts können keine [allgemeinen APIs verwenden.](/javascript/api/office) Wenn Sie Authentifizierung, Dialogfelder oder andere Features benötigen, die nur von allgemeinen APIs unterstützt werden, müssen Sie wahrscheinlich ein Office-Add-In anstelle eines Office-Skripts erstellen.

## <a name="see-also"></a>Weitere Informationen

- [Office-Skripts in Excel im Web](../overview/excel.md)
- [Unterschiede zwischen Office-Skripts und VBA-Makros](vba-differences.md)
- [Behandeln von Problemen mit Office-Skripts](../testing/troubleshooting.md)
- [Erstellen eines Excel-Aufgabenbereich-Add-Ins](/office/dev/add-ins/quickstarts/excel-quickstart-jquery)
