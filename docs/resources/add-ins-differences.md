---
title: Unterschiede zwischen Office-Skripts und Office-Add-Ins
description: Das Verhalten und die API-Unterschiede zwischen Office-Skripts und Office-Add-Ins.
ms.date: 06/01/2020
localization_priority: Normal
ms.openlocfilehash: ddac6cc68874da34ae76c66a5c5b84ffa7a60eec
ms.sourcegitcommit: ce72354381561dc167ea0092efd915642a9161b3
ms.translationtype: MT
ms.contentlocale: de-DE
ms.lasthandoff: 09/30/2020
ms.locfileid: "48319651"
---
# <a name="differences-between-office-scripts-and-office-add-ins"></a>Unterschiede zwischen Office-Skripts und Office-Add-Ins

Office-Add-Ins-und Office-Skripts haben viele Gemeinsamkeiten. Sie bieten sowohl eine automatisierte Steuerung einer Excel-Arbeitsmappe als auch eine JavaScript-API. Die Office-Skript-APIs sind jedoch eine spezielle synchrone Version der Office-JavaScript-API.

![Ein Diagramm mit vier Quadranten, in dem die Fokusbereiche für unterschiedliche Office-Erweiterbarkeits Lösungen angezeigt werden. Sowohl Office-Skripts als auch Office-WebAdd-ins konzentrieren sich auf das Internet und die Zusammenarbeit, aber Office-Skripts bieten Endbenutzern (während Office-WebAdd-ins als Ziel für professionelle Entwickler gedacht sind).)](../images/office-programmability-diagram.png)

Office-Skripts werden bis zum Abschluss mit einer manuellen Tastendruck oder als Schritt in [Power automatisieren](https://flow.microsoft.com/)ausgeführt, während Office-Add-Ins bleiben, während Ihre Aufgabenbereiche geöffnet sind. Dies bedeutet, dass die Add-Ins den Status während einer Sitzung beibehalten können, während Office-Skripts keinen internen Status zwischen Läufen aufrecht erhalten. Wenn Sie feststellen, dass Ihre Excel-Erweiterung die Funktionen der Skriptplattform überschreiten muss, lesen Sie die [Office-Add-ins Dokumentation](/office/dev/add-ins) , um mehr über Office-Add-Ins zu erfahren.

Der Rest dieses Artikels beschreibt die Hauptunterschiede zwischen Office-Add-Ins-und Office-Skripts.

## <a name="platform-support"></a>Plattformunterstützung

Office-Add-Ins sind plattformübergreifend. Sie funktionieren auf Windows-Desktop-, Mac-, IOS-und Webplattformen und bieten die gleiche Benutzeroberfläche. Jede Ausnahme hiervon wird in der Dokumentation der einzelnen API notiert.

Office-Skripts werden derzeit nur von für Excel im Internet unterstützt. Alle Aufzeichnung, Bearbeitung und Ausführung erfolgt auf der Webplattform.

## <a name="apis"></a>APIs

Es gibt keine synchrone Version der Office-JavaScript-APIs für Office-Add-Ins. Die standardmäßigen Office-Skript-APIs sind für die Plattform eindeutig und weisen zahlreiche Optimierungen und Änderungen auf, um die Verwendung des `load` / `sync` Paradigmas zu vermeiden.

Einige der [Excel-JavaScript-APIs](/javascript/api/excel?view=excel-js-preview&preserve-view=true) sind mit den [Async-APIs von Office Scripts](../develop/excel-async-model.md)kompatibel. Einige Beispiele und Add-in-Codeblöcke konnten in `Excel.run` Blöcke mit minimaler Übersetzung portiert werden. Während die beiden Plattformen die Funktionalität teilen, gibt es Lücken. Die beiden Haupt-API-Sätze, die Office-Add-ins haben, aber Office-Skripts sind keine Ereignisse und die allgemeinen APIs.

### <a name="events"></a>Ereignisse

Office-Skripts unterstützen keine [Ereignisse](/office/dev/add-ins/excel/excel-add-ins-events). Jedes Skript führt den Code in einer einzigen `main` Methode aus und endet dann. Es wird nicht erneut aktiviert, wenn Ereignisse ausgelöst werden und daher keine Ereignisse registrieren können.

### <a name="common-apis"></a>Allgemeine APIs

Office-Skripts können keine [allgemeinen APIs](/javascript/api/office)verwenden. Wenn Sie Authentifizierung, Dialogfenster oder andere Features benötigen, die nur von gängigen APIs unterstützt werden, müssen Sie wahrscheinlich ein Office-Add-in anstelle eines Office-Skripts erstellen.

## <a name="see-also"></a>Mehr dazu

- [Office-Skripts in Excel im Web](../overview/excel.md)
- [Unterschiede zwischen Office-Skripts und VBA-Makros](vba-differences.md)
- [Behandeln von Problemen mit Office-Skripts](../testing/troubleshooting.md)
- [Erstellen eines Excel-Aufgabenbereich-Add-Ins](/office/dev/add-ins/quickstarts/excel-quickstart-jquery)
