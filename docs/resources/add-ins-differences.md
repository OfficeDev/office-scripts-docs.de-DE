---
title: Unterschiede zwischen Office-Skripts und Office-Add-Ins
description: Das Verhalten und die API-Unterschiede zwischen Office Skripts und Office Add-Ins.
ms.date: 06/01/2020
localization_priority: Normal
ms.openlocfilehash: 45993d08d85cfceb299216dddbe2e7da9fd2e404
ms.sourcegitcommit: f7a7aebfb687f2a35dbed07ed62ff352a114525a
ms.translationtype: MT
ms.contentlocale: de-DE
ms.lasthandoff: 05/06/2021
ms.locfileid: "52232634"
---
# <a name="differences-between-office-scripts-and-office-add-ins"></a>Unterschiede zwischen Office-Skripts und Office-Add-Ins

Office Add-Ins und Office haben viele gemeinsam. Beide bieten eine automatisierte Steuerung einer Excel eine JavaScript-API. Bei den Office-Skript-APIs handelt es sich jedoch um eine spezielle, synchrone Version der Office JavaScript-API.

:::image type="content" source="../images/office-programmability-diagram.png" alt-text="Ein Diagramm mit vier Quadranten, in dem die Fokusbereiche für unterschiedliche Office dargestellt werden. Sowohl Office-Skripts als auch Office-Web-Add-Ins konzentrieren sich auf das Web und die Zusammenarbeit, aber Office-Skripts sind für Endbenutzer (während Office Web-Add-Ins auf professionelle Entwickler ausgerichtet sind)":::

Office Skripts werden mit einem manuellen Tastendruck oder als Schritt in [Power Automate](https://flow.microsoft.com/)ausgeführt, während Office Add-Ins beibehalten werden, während ihre Aufgabenbereiche geöffnet sind. Dies bedeutet, dass die Add-Ins während einer Sitzung den Status beibehalten können, während Office Skripts keinen internen Status zwischen den Einzelnen Ausführen beibehalten. Wenn Sie feststellen, Excel Erweiterung die Funktionen der Skriptplattform überschreiten muss, lesen Sie die [Office-Add-Ins-Dokumentation,](/office/dev/add-ins) um mehr über Office-Add-Ins zu erfahren.

Der Rest dieses Artikels beschreibt die wichtigsten Unterschiede zwischen Office-Add-Ins und Office Skripts.

## <a name="platform-support"></a>Plattformunterstützung

Office Add-Ins sind plattformübergreifender. Sie arbeiten über Windows Desktop-, Mac-, iOS- und Webplattformen hinweg und bieten die gleiche Benutzererfahrung auf allen Plattformen. Eine Ausnahme davon ist in der Dokumentation der einzelnen API vermerkt.

Office Skripts werden derzeit nur für Excel im Web. Die Aufzeichnung, Bearbeitung und Ausführung erfolgt auf der Webplattform.

## <a name="apis"></a>APIs

Es gibt keine synchrone Version der Office JavaScript-APIs für Office-Add-Ins. Die standardmäßigen Office Skript-APIs sind für die Plattform einzigartig und verfügen über zahlreiche Optimierungen und Änderungen, um die Verwendung des Paradigmas zu `load` / `sync` vermeiden.

Einige der Excel [JavaScript-APIs](/javascript/api/excel?view=excel-js-preview&preserve-view=true) sind mit den Office [Skripts Async-APIs kompatibel.](../develop/excel-async-model.md) Einige Beispiele und Add-In-Codeblöcke können in Blöcke mit `Excel.run` minimaler Übersetzung portiert werden. Während die beiden Plattformen die Funktionalität gemeinsam nutzen, gibt es Lücken. Die beiden wichtigsten API-Sätze, die Office Add-Ins haben, aber Office Skripts sind keine Ereignisse und die allgemeinen APIs.

### <a name="events"></a>Ereignisse

Office Skripts unterstützen keine [Ereignisse](/office/dev/add-ins/excel/excel-add-ins-events). Jedes Skript führt den Code in einer einzelnen `main` Methode aus und endet dann. Es wird nicht reaktiviert, wenn Ereignisse ausgelöst werden, und kann daher keine Ereignisse registrieren.

### <a name="common-apis"></a>Allgemeine APIs

Office Skripts können keine [allgemeinen APIs verwenden.](/javascript/api/office) Wenn Sie Authentifizierung, Dialogfelder oder andere Features benötigen, die nur von allgemeinen APIs unterstützt werden, müssen Sie wahrscheinlich ein Office-Add-In anstelle eines Office erstellen.

## <a name="see-also"></a>Siehe auch

- [Office-Skripts in Excel im Web](../overview/excel.md)
- [Unterschiede zwischen Office Skripts und VBA-Makros](vba-differences.md)
- [Behandeln von Problemen mit Office-Skripts](../testing/troubleshooting.md)
- [Erstellen eines Excel-Aufgabenbereich-Add-Ins](/office/dev/add-ins/quickstarts/excel-quickstart-jquery)
