---
title: Unterschiede zwischen Office-Skripts und Office-Add-Ins
description: Das Verhalten und die API-Unterschiede zwischen Office-Skripts und Office-Add-Ins.
ms.date: 02/04/2022
ms.localizationpriority: medium
ms.openlocfilehash: a3df4daf04f963598d2cb31f82dd2c1c9923fdc8
ms.sourcegitcommit: 33fe0f6807daefb16b148fd73c863de101f47cea
ms.translationtype: MT
ms.contentlocale: de-DE
ms.lasthandoff: 08/08/2022
ms.locfileid: "67281910"
---
# <a name="differences-between-office-scripts-and-office-add-ins"></a>Unterschiede zwischen Office-Skripts und Office-Add-Ins

Verstehen Sie die Unterschiede zwischen Office-Skripts und Office-Add-Ins, um zu wissen, wann sie verwendet werden sollten. Office-Skripts sind so konzipiert, dass sie schnell von allen Benutzern erstellt werden können, die ihren Workflow verbessern möchten. Office-Add-Ins können in die Office-Benutzeroberfläche integriert werden, um eine interaktivere Benutzeroberfläche über Menübandschaltflächen und Aufgabenbereiche zu ermöglichen. Office-Add-Ins können auch integrierte Excel-Funktionen erweitern, indem sie benutzerdefinierte Funktionen bereitstellen.

:::image type="content" source="../images/office-programmability-diagram.png" alt-text="Ein Vier-Quadranten-Diagramm mit den Fokusbereichen für verschiedene Office-Erweiterbarkeitslösungen. Sowohl Office-Skripts als auch Office-Web-Add-Ins konzentrieren sich auf das Web und die Zusammenarbeit, aber Office-Skripts richten sich an Endbenutzer (während Office-Web-Add-Ins professionelle Entwickler ansprechen).":::

Office-Skripts werden bis zum Abschluss mit einem manuellen Tastendruck oder als Schritt in [Power Automate](https://flow.microsoft.com/) ausgeführt, während Office-Add-Ins weiterhin ausgeführt werden, je nachdem, wie sie konfiguriert sind. Sie können z. B. ein Office-Add-In so konfigurieren, dass es auch dann weiter ausgeführt wird, wenn sein Aufgabenbereich geschlossen ist. Dies bedeutet, dass Office-Add-Ins während einer Sitzung den Status beibehalten, während Office-Skripts keinen internen Zustand zwischen den Ausführungen beibehalten. Wenn die von Ihnen entwickelte Lösung einen verwalteten Status erfordert, sollten Sie die [Dokumentation zu Office-Add-Ins](/office/dev/add-ins) besuchen, um mehr über Office-Add-Ins zu erfahren.

Der Rest dieses Artikels beschreibt die wichtigsten Unterschiede zwischen Office-Add-Ins und Office-Skripts.

## <a name="platform-support"></a>Plattformunterstützung

Office-Add-Ins sind plattformübergreifend. Sie funktionieren auf Windows-Desktop-, Mac-, iOS- und Webplattformen und bieten die gleiche Erfahrung auf den einzelnen Plattformen. Jede Ausnahme ist in der Dokumentation der einzelnen API aufgeführt.

Office-Skripts werden derzeit nur für Excel im Web unterstützt. Die gesamte Aufzeichnungs-, Bearbeitungs- und Skriptverwaltung erfolgt auf der Webplattform.

### <a name="script-support-for-excel-on-windows"></a>Skriptunterstützung für Excel unter Windows

[!INCLUDE [Run-from-button support](../includes/run-from-button-desktop-support.md)]

## <a name="apis"></a>APIs

Während die Office-JavaScript-APIs für Office-Add-Ins und die Office-Skript-APIs einige Funktionen gemeinsam nutzen, sind sie unterschiedliche Plattformen. Die Office-Skripts-APIs sind eine optimierte, synchrone Teilmenge des Excel-JavaScript-API-Modells. Der hauptunterschied ist die Verwendung des `load`/`sync` Paradigmas mit Add-Ins. Darüber hinaus bieten Add-Ins APIs für Ereignisse und einen breiteren Satz von Funktionen außerhalb von Excel, die als allgemeine APIs bezeichnet werden.

### <a name="events"></a>Events

Office-Skripts unterstützen keine [Ereignisse](/office/dev/add-ins/excel/excel-add-ins-events) auf Arbeitsmappenebene. Skripts werden entweder ausgelöst, indem Benutzer die Schaltfläche " **Ausführen** " für ein Skript oder über Power Automate auswählen. Jedes Skript führt den Code in einer einzigen `main` Funktion aus und endet dann.

### <a name="common-apis"></a>Allgemeine APIs

Office-Skripts können keine [allgemeinen APIs](/javascript/api/office) verwenden. Wenn Sie Authentifizierung, Dialogfelder oder andere Features benötigen, die nur von allgemeinen APIs unterstützt werden, müssen Sie wahrscheinlich ein Office-Add-In anstelle eines Office-Skripts erstellen.

## <a name="see-also"></a>Siehe auch

- [Office-Skripts in Excel](../overview/excel.md)
- [Unterschiede zwischen Office-Skripts und VBA-Makros](vba-differences.md)
- [Behandeln von Problemen mit Office-Skripts](../testing/troubleshooting.md)
- [Erstellen eines Excel-Aufgabenbereich-Add-Ins](/office/dev/add-ins/quickstarts/excel-quickstart-jquery)
