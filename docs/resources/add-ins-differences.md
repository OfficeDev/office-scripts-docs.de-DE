---
title: Unterschiede zwischen Office-Skripts und Office-Add-Ins
description: Das Verhalten und die API-Unterschiede zwischen Office Skripts und Office Add-Ins.
ms.date: 06/01/2020
localization_priority: Normal
ms.openlocfilehash: 5c30406867da05952dedda684f765df5e7a7e53f
ms.sourcegitcommit: 09d8859d5269ada8f1d0e141f6b5a4f96d95a739
ms.translationtype: MT
ms.contentlocale: de-DE
ms.lasthandoff: 05/24/2021
ms.locfileid: "52631678"
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

Während die Office JavaScript-APIs für Office-Add-Ins und die Office-Skript-APIs einige Funktionen gemeinsam nutzen, sind sie unterschiedliche Plattformen. Die Office Skript-APIs sind eine optimierte, synchrone Version des Excel JavaScript-API-Modells. Der Hauptunterschied besteht in der Verwendung des `load` / `sync` Paradigmas mit Add-Ins. Darüber hinaus bieten Add-Ins APIs für Ereignisse und einen umfassenderen Funktionsumfang außerhalb von Excel, die als allgemeine APIs bezeichnet werden.

### <a name="events"></a>Ereignisse

Office Skripts unterstützen keine [Ereignisse](/office/dev/add-ins/excel/excel-add-ins-events). Jedes Skript führt den Code in einer einzelnen `main` Methode aus und endet dann. Es wird nicht reaktiviert, wenn Ereignisse ausgelöst werden, und kann daher keine Ereignisse registrieren.

### <a name="common-apis"></a>Allgemeine APIs

Office Skripts können keine [allgemeinen APIs verwenden.](/javascript/api/office) Wenn Sie Authentifizierung, Dialogfelder oder andere Features benötigen, die nur von allgemeinen APIs unterstützt werden, müssen Sie wahrscheinlich ein Office-Add-In anstelle eines Office erstellen.

## <a name="see-also"></a>Sehen Sie ebenfalls

- [Office-Skripts in Excel im Web](../overview/excel.md)
- [Unterschiede zwischen Office Skripts und VBA-Makros](vba-differences.md)
- [Behandeln von Problemen mit Office-Skripts](../testing/troubleshooting.md)
- [Erstellen eines Excel-Aufgabenbereich-Add-Ins](/office/dev/add-ins/quickstarts/excel-quickstart-jquery)
