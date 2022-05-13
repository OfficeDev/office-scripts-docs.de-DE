---
title: Unterschiede zwischen Office-Skripts und VBA-Makros
description: Die Verhaltens- und API-Unterschiede zwischen Office-Skripts und Excel VBA-Makros.
ms.date: 02/04/2022
ms.localizationpriority: medium
ms.openlocfilehash: 60e4fba6e63967302066f544b76fb20a8c8630a6
ms.sourcegitcommit: 34c7740c9bff0e4c7426e01029f967724bfee566
ms.translationtype: MT
ms.contentlocale: de-DE
ms.lasthandoff: 05/13/2022
ms.locfileid: "65393614"
---
# <a name="differences-between-office-scripts-and-vba-macros"></a>Unterschiede zwischen Office-Skripts und VBA-Makros

Office Skripts und VBA-Makros haben viel gemeinsam. Beide ermöglichen es Benutzern, Lösungen über eine einfach zu bedienende Aktionsaufzeichnung zu automatisieren und die Bearbeitung dieser Aufzeichnungen zu ermöglichen. Beide Frameworks sind darauf ausgelegt, Personen, die sich nicht als Programmierer betrachten, zu befähigen, kleine Programme in Excel zu erstellen.

Der grundlegende Unterschied besteht darin, dass VBA-Makros für Desktoplösungen entwickelt werden und Office Skripts für sichere, cloudbasierte Lösungen entwickelt wurden. Derzeit werden Office Skripts nur in Excel im Web unterstützt.

:::image type="content" source="../images/office-programmability-diagram.png" alt-text="Ein Diagramm mit vier Quadranten, das die Fokusbereiche für verschiedene Office Erweiterbarkeitslösungen zeigt. Sowohl Office Skripts als auch VBA-Makros sollen Endbenutzern beim Erstellen von Lösungen helfen, aber Office Skripts werden für das Web und die Zusammenarbeit erstellt (während VBA für den Desktop vorgesehen ist).":::

In diesem Artikel werden die wichtigsten Unterschiede zwischen VBA-Makros (sowie VBA im Allgemeinen) und Office Skripts beschrieben. Da Office Skripts nur für Excel verfügbar sind, ist dies der einzige Host, der hier besprochen wird.

## <a name="platform-and-ecosystem"></a>Plattform und Ökosystem

VBA wird von Excel auf Windows und Mac unterstützt. Office Skripts wird von Excel im Web unterstützt.

Die beiden Lösungen wurden für ihre jeweiligen Plattformen konzipiert. VBA kann mit dem Desktop eines Benutzers interagieren, um eine Verbindung mit ähnlichen Technologien wie COM und OLE herzustellen. VBA hat jedoch keine bequeme Möglichkeit, sich an das Internet zu wenden. Office Skripts verwenden eine universelle Laufzeit für JavaScript. Dies bietet ein konsistentes Verhalten und Barrierefreiheit, unabhängig davon, auf welchem Computer das Skript ausgeführt wird. Sie können auch Anrufe an andere Webdienste tätigen.

### <a name="script-support-for-excel-on-windows"></a>Skriptunterstützung für Excel auf Windows

[!INCLUDE [Run-from-button support](../includes/run-from-button-desktop-support.md)]

## <a name="security"></a>Sicherheit

VBA-Makros verfügen über die gleiche Sicherheitsfreigabe wie Excel. Dadurch erhalten sie vollen Zugriff auf Ihren Desktop. Office Skripts haben nur Zugriff auf die Arbeitsmappe, nicht auf den Computer, auf dem die Arbeitsmappe gehostet wird. Darüber hinaus können keine JavaScript-Authentifizierungstoken für Skripts freigegeben werden. Dies bedeutet, dass das Skript weder über die Token des angemeldeten Benutzers verfügt noch über API-Funktionen zum Anmelden bei einem externen Dienst verfügt, sodass es vorhandene Token nicht verwenden kann, um externe Aufrufe im Namen des Benutzers auszuführen.

Administratoren haben drei Optionen für VBA-Makros: Alle Makros im Mandanten zulassen, keine Makros auf dem Mandanten zulassen oder nur Makros mit signierten Zertifikaten zulassen. Diese mangelnde Granularität macht es schwierig, einen einzelnen schlechten Akteur zu isolieren. Derzeit kann Office Skripts für einen gesamten Mandanten, für einen gesamten Mandanten oder für eine Gruppe von Benutzern in einem Mandanten deaktiviert sein. Administratoren haben auch die Kontrolle darüber, wer Skripts für andere Personen freigeben und wer Skripts in Power Automate verwenden kann.

## <a name="coverage"></a>Abdeckung

Derzeit bietet VBA eine umfassendere Abdeckung Excel Features, insbesondere derjenigen, die auf dem Desktopclient verfügbar sind. Office Skripts decken fast alle Szenarien für Excel im Web ab. Darüber hinaus werden Office Skripts, sobald neue Features im Web veröffentlicht werden, sowohl für die Action Recorder- als auch die JavaScript-APIs unterstützt.

Office Skripts unterstützen keine [Ereignisse](/office/vba/excel/concepts/events-worksheetfunctions-shapes/using-events-with-excel-objects) auf Excel-Ebene. Skripts werden nur ausgeführt, wenn ein Benutzer sie manuell startet oder wenn ein Power Automate-Fluss das Skript aufruft.

## <a name="power-automate"></a>Power Automate

Office Skripts können über Power Automate ausgeführt werden. Ihre Arbeitsmappe kann durch geplante oder ereignisgesteuerte Abläufe aktualisiert werden, sodass Sie Workflows automatisieren können, ohne Excel zu öffnen. Dies bedeutet: Solange Ihre Arbeitsmappe in OneDrive gespeichert ist (und für Power Automate zugänglich ist), kann ein Fluss Ihre Skripts unabhängig davon ausführen, ob Sie und Ihre Organisation den Desktop-, Mac- oder Webclient von Excel verwenden.

VBA verfügt nicht über einen Power Automate Connector. In allen unterstützten VBA-Szenarien ist ein Benutzer an der Ausführung des Makros beteiligt.

Probieren Sie die [Anrufskripts aus einem manuellen Power Automate Flusslernprogramm](../tutorials/excel-power-automate-manual.md) aus, um mehr über Power Automate zu erfahren. Sie können sich auch das Beispiel für [automatisierte Aufgabenerinnerungen](scenarios/task-reminders.md) ansehen, um Office Skripts anzuzeigen, die über Power Automate in einem realen Szenario mit Teams verbunden sind.

## <a name="see-also"></a>Siehe auch

- [Office Skripts in Excel](../overview/excel.md)
- [Ausführen Office Skripts mit Power Automate](../develop/power-automate-integration.md)
- [Unterschiede zwischen Office-Skripts und Office-Add-Ins](add-ins-differences.md)
- [Behandeln von Problemen mit Office-Skripts](../testing/troubleshooting.md)
- [VBA-Referenz für Excel](/office/vba/api/overview/excel)
