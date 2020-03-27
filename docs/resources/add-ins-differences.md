---
title: Unterschiede zwischen Office-Skripts und Office-Add-ins
description: Das Verhalten und die API-Unterschiede zwischen Office-Skripts und Office-Add-Ins.
ms.date: 03/23/2020
localization_priority: Normal
ms.openlocfilehash: 2290d4e34b7a7286d67443de9e9c64bad4fcd4b7
ms.sourcegitcommit: d556aaefac80e55f53ac56b7f6ecbc657ebd426f
ms.translationtype: MT
ms.contentlocale: de-DE
ms.lasthandoff: 03/26/2020
ms.locfileid: "42978715"
---
# <a name="differences-between-office-scripts-and-office-add-ins"></a>Unterschiede zwischen Office-Skripts und Office-Add-ins

Office-Add-Ins-und Office-Skripts haben viele Gemeinsamkeiten. Beide bieten eine automatisierte Steuerung einer Excel-Arbeitsmappe über `Excel` den Namespace der Office-JavaScript-API. Office-Skripts sind jedoch in ihrem Umfang beschränkter.

![Ein Diagramm mit vier Quadranten, in dem die Fokusbereiche für unterschiedliche Office-Erweiterbarkeits Lösungen angezeigt werden. Sowohl Office-Skripts als auch Office-WebAdd-ins konzentrieren sich auf das Internet und die Zusammenarbeit, aber Office-Skripts bieten Endbenutzern (während Office-WebAdd-ins als Ziel für professionelle Entwickler gedacht sind).)](../images/office-programmability-diagram.png)

Office-Skripts werden bis zum Abschluss mit einer manuellen Tastendruck oder als Schritt in [Power automatisieren](https://flow.microsoft.com/)ausgeführt, während Office-Add-Ins bleiben, während Ihre Aufgabenbereiche geöffnet sind. Dies bedeutet, dass die Add-Ins den Status während einer Sitzung beibehalten können, während Office-Skripts keinen internen Status zwischen Läufen aufrecht erhalten. Wenn Sie feststellen, dass Ihre Excel-Erweiterung die Funktionen der Skriptplattform überschreiten muss, lesen Sie die [Office-Add-ins Dokumentation](/office/dev/add-ins) , um mehr über Office-Add-Ins zu erfahren.

Der Rest dieses Artikels beschreibt die Hauptunterschiede zwischen Office-Add-Ins-und Office-Skripts.

## <a name="platform-support"></a>Plattformunterstützung

Office-Add-Ins sind plattformübergreifend. Sie funktionieren auf Windows-Desktop-, Mac-, IOS-und Webplattformen und bieten die gleiche Benutzeroberfläche. Jede Ausnahme hiervon wird in der Dokumentation der einzelnen API notiert.

Office-Skripts werden derzeit nur von für Excel im Internet unterstützt. Alle Aufzeichnung, Bearbeitung und Ausführung erfolgt auf der Webplattform.

## <a name="apis"></a>APIs

Office-Skripts unterstützen die meisten der Excel-JavaScript-APIs, was bedeutet, dass es eine Vielzahl von Funktionsüberschneidungen zwischen den beiden Plattformen gibt. Es gibt zwei Ausnahmen: Ereignisse und allgemeine APIs.

### <a name="events"></a>Ereignisse

Office-Skripts unterstützen keine [Ereignisse](/office/dev/add-ins/excel/excel-add-ins-events). Jedes Skript führt den Code in einer einzigen `main` Methode aus und endet dann. Es wird nicht erneut aktiviert, wenn Ereignisse ausgelöst werden und daher keine Ereignisse registrieren können.

### <a name="common-apis"></a>Allgemeine APIs

Office-Skripts können keine [allgemeinen APIs](/javascript/api/office)verwenden. Wenn Sie Authentifizierung, Dialogfenster oder andere Features benötigen, die nur von gängigen APIs unterstützt werden, müssen Sie wahrscheinlich ein Office-Add-in anstelle eines Office-Skripts erstellen.

## <a name="see-also"></a>Siehe auch

- [Office-Skripts in Excel im Web](../overview/excel.md)
- [Unterschiede zwischen Office-Skripts und VBA-Makros](vba-differences.md)
- [Behandeln von Problemen mit Office-Skripts](../testing/troubleshooting.md)
- [Erstellen eines Excel-Aufgabenbereich-Add-Ins](/office/dev/add-ins/quickstarts/excel-quickstart-jquery)
