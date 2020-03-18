---
title: Unterschiede zwischen Office-Skripts und Office-Add-ins
description: Das Verhalten und die API-Unterschiede zwischen Office-Skripts und Office-Add-Ins.
ms.date: 12/12/2019
localization_priority: Normal
ms.openlocfilehash: 4626afb66b54c94a72f29b039c601435c089d64d
ms.sourcegitcommit: b075eed5a6f275274fbbf6d62633219eac416f26
ms.translationtype: MT
ms.contentlocale: de-DE
ms.lasthandoff: 03/10/2020
ms.locfileid: "42700230"
---
# <a name="differences-between-office-scripts-and-office-add-ins"></a>Unterschiede zwischen Office-Skripts und Office-Add-ins

Office-Add-Ins-und Office-Skripts haben viele Gemeinsamkeiten. Beide bieten eine automatisierte Steuerung einer Excel-Arbeitsmappe über `Excel` den Namespace der Office-JavaScript-API. Office-Skripts sind jedoch in ihrem Umfang beschränkter.

Office-Skripts werden bis zum Abschluss mit einer manuellen Schaltfläche drücken, während Office-Add-Ins auf Benutzerinteraktion basieren und dauerhaft bleiben, während die Arbeitsmappe verwendet wird. Wenn Sie feststellen, dass Ihre Excel-Erweiterung die Funktionen der Skriptplattform überschreiten muss, lesen Sie die [Office-Add-ins Dokumentation](/office/dev/add-ins) , um mehr über Office-Add-Ins zu erfahren.

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

- [Office-Skripts in Excel im Internet](../overview/excel.md)
- [Problembehandlung bei Office-Skripts](../testing/troubleshooting.md)
- [Erstellen eines Excel-Aufgabenbereich-Add-Ins](/office/dev/add-ins/quickstarts/excel-quickstart-jquery)