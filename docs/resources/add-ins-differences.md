---
title: Unterschiede zwischen Office-Skripts und Office-Add-Ins
description: Das Verhalten und die API-Unterschiede zwischen Office Skripts und Office-Add-Ins.
ms.date: 02/04/2022
ms.localizationpriority: medium
ms.openlocfilehash: 018d210208bc78da894678d21e368864522cb83e
ms.sourcegitcommit: 7023b9e23499806901a5ecf8ebc460b76887cca6
ms.translationtype: MT
ms.contentlocale: de-DE
ms.lasthandoff: 03/31/2022
ms.locfileid: "64585608"
---
# <a name="differences-between-office-scripts-and-office-add-ins"></a>Unterschiede zwischen Office-Skripts und Office-Add-Ins

Verstehen Sie die Unterschiede zwischen Office Skripts und Office-Add-Ins, um zu wissen, wann die einzelnen Skripts verwendet werden sollen. Office Skripts sind so konzipiert, dass sie schnell von jedem erstellt werden können, der den Workflow verbessern möchte. Office-Add-Ins können über Menübandschaltflächen und Aufgabenbereiche in die Office Ui integriert werden. Office-Add-Ins können auch integrierte Excel Funktionen erweitern, indem benutzerdefinierte Funktionen bereitgestellt werden.

:::image type="content" source="../images/office-programmability-diagram.png" alt-text="Ein Vier-Quadranten-Diagramm mit den Fokusbereichen für verschiedene Office Erweiterbarkeitslösungen. Sowohl Office Skripts als auch Office Web-Add-Ins konzentrieren sich auf das Web und die Zusammenarbeit, aber Office Skripts richten sich an Endbenutzer (während Office Web-Add-Ins auf professionelle Entwickler abzielen).":::

Office Skripts werden bis zum Abschluss mit einem manuellen Tastendruck oder als Schritt in [Power Automate](https://flow.microsoft.com/) ausgeführt, während Office Add-Ins je nach Konfiguration weiter ausgeführt werden. Beispielsweise können Sie ein Office-Add-In so konfigurieren, dass es auch dann weiterhin ausgeführt wird, wenn der Aufgabenbereich geschlossen ist. Dies bedeutet, dass Office-Add-Ins den Status während einer Sitzung beibehalten, während Office Skripts keinen internen Status zwischen Ausführungen beibehalten. Wenn die Lösung, die Sie erstellen, einen verwalteten Zustand erfordert, sollten Sie die [Dokumentation zu Office Add-Ins](/office/dev/add-ins) besuchen, um mehr über Office Add-Ins zu erfahren.

Der Rest dieses Artikels beschreibt die Hauptunterschiede zwischen Office-Add-Ins und Office-Skripts.

## <a name="platform-support"></a>Plattformunterstützung

Office-Add-Ins sind plattformübergreifend. Sie funktionieren auf Windows Desktop-, Mac-, iOS- und Webplattformen und bieten jeweils dieselbe Oberfläche. Ausnahmen hierzu sind in der Dokumentation der einzelnen API aufgeführt.

Office Skripts werden derzeit nur für Excel im Web unterstützt. Die gesamte Aufzeichnung, Bearbeitung und Skriptverwaltung erfolgt auf der Webplattform.

### <a name="script-support-for-excel-on-windows"></a>Skriptunterstützung für Excel auf Windows

[!INCLUDE [Run-from-button support](../includes/run-from-button-desktop-support.md)]

## <a name="apis"></a>APIs

Obwohl die Office JavaScript-APIs für Office-Add-Ins und die APIs für Office Skripts einige Funktionen nutzen, sind sie unterschiedliche Plattformen. Die APIs Office Skripts sind eine optimierte, synchrone Teilmenge des Excel JavaScript-API-Modells. Der Hauptunterschied besteht in der Verwendung des `load`/`sync` Paradigmas mit Add-Ins. Darüber hinaus bieten Add-Ins APIs für Ereignisse und einen breiteren Satz von Funktionen außerhalb von Excel, die als allgemeine APIs bezeichnet werden.

### <a name="events"></a>Ereignisse

Office Skripts unterstützen keine [Ereignisse](/office/dev/add-ins/excel/excel-add-ins-events) auf Arbeitsmappenebene. Skripts werden entweder durch Benutzer ausgelöst, die die Schaltfläche "**Ausführen**" für ein Skript auswählen, oder über Power Automate. Jedes Skript führt den Code in einer einzigen `main` Methode aus und endet dann.

### <a name="common-apis"></a>Allgemeine APIs

Office Skripts können keine [allgemeinen APIs](/javascript/api/office) verwenden. Wenn Sie Authentifizierung, Dialogfelder oder andere Features benötigen, die nur von allgemeinen APIs unterstützt werden, müssen Sie wahrscheinlich ein Office-Add-In anstelle eines Office Skripts erstellen.

## <a name="see-also"></a>Weitere Informationen

- [Office-Skripts in Excel im Web](../overview/excel.md)
- [Unterschiede zwischen Office Skripts und VBA-Makros](vba-differences.md)
- [Behandeln von Problemen mit Office-Skripts](../testing/troubleshooting.md)
- [Erstellen eines Excel-Aufgabenbereich-Add-Ins](/office/dev/add-ins/quickstarts/excel-quickstart-jquery)
