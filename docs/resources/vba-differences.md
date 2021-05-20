---
title: Unterschiede zwischen Office Skripts und VBA-Makros
description: Das Verhalten und die API-Unterschiede zwischen Office Skripts und Excel VBA-Makros.
ms.date: 05/18/2021
localization_priority: Normal
ms.openlocfilehash: 612a5f21d935fd262a6e9fd12a3431956105636a
ms.sourcegitcommit: 4687693f02fc90a57ba30c461f35046e02e6f5fb
ms.translationtype: MT
ms.contentlocale: de-DE
ms.lasthandoff: 05/19/2021
ms.locfileid: "52545588"
---
# <a name="differences-between-office-scripts-and-vba-macros"></a>Unterschiede zwischen Office Skripts und VBA-Makros

Office Skripts und VBA-Makros haben viel gemeinsam. Beide ermöglichen es Benutzern, Lösungen über einen benutzerfreundlichen Actionrecorder zu automatisieren und diese Aufnahmen zu bearbeiten. Beide Frameworks sind so konzipiert, dass Menschen, die sich nicht als Programmierer betrachten, kleine Programme in Excel erstellen können.
Der grundlegende Unterschied besteht darin, dass VBA-Makros für Desktop-Lösungen entwickelt wurden und Office Scripts für sichere, Cloud-basierte Lösungen entwickelt wurden. Derzeit werden Office Skripts nur in Excel im Web unterstützt.

:::image type="content" source="../images/office-programmability-diagram.png" alt-text="Ein Vier-Quadranten-Diagramm, das die Schwerpunkte für verschiedene Office Erweiterbarkeitslösungen zeigt. Sowohl Office-Skripts als auch VBA-Makros sollen Endbenutzern beim Erstellen von Lösungen helfen, aber Office Skripts für das Web und die Zusammenarbeit erstellt werden (wobei VBA für den Desktop bestimmt ist).":::

In diesem Artikel werden die Hauptunterschiede zwischen VBA-Makros (sowie VBA im Allgemeinen) und Office Skripts beschrieben. Da Office Skripts nur für Excel verfügbar sind, ist dies der einzige Host, der hier besprochen wird.

## <a name="platform-and-ecosystem"></a>Plattform und Ökosystem

VBA ist für den Desktop konzipiert und Office Scripts für das Web entwickelt wurden. VBA kann mit dem Desktop eines Benutzers interagieren, um eine Verbindung mit ähnlichen Technologien wie COM und OLE herzustellen. VBA hat jedoch keine bequeme Möglichkeit, ins Internet zu rufen.

Office Skripts verwenden eine universelle Laufzeit für JavaScript. Dies ermöglicht konsistentes Verhalten und Zugänglichkeit, unabhängig davon, auf dem Computer das Skript ausgeführt wird. Sie können auch Anrufe an andere Webdienste tätigen.

## <a name="security"></a>Sicherheit

VBA-Makros haben die gleiche Sicherheitsfreigabe wie Excel. Dadurch erhalten sie vollen Zugriff auf Ihren Desktop. Office Skripts haben nur Zugriff auf die Arbeitsmappe, nicht auf den Computer, auf dem die Arbeitsmappe gehostet wird. Darüber hinaus können keine JavaScript-Authentifizierungstoken für Skripts freigegeben werden. Dies bedeutet, dass das Skript weder über die Token des angemeldeten Benutzers noch über API-Funktionen zum Anmelden bei einem externen Dienst verfügt, sodass sie keine vorhandenen Token verwenden können, um externe Aufrufe im Namen des Benutzers zu tätigen.

Administratoren haben drei Optionen für VBA-Makros: Alle Makros auf dem Mandanten zulassen, keine Makros für den Mandanten zulassen oder nur Makros mit signierten Zertifikaten zulassen. Dieser Mangel an Granularität macht es schwierig, einen einzelnen schlechten Akteur zu isolieren. Derzeit sind Office Skripts für einen Mandanten entweder ein- oder ausgeschaltet. Wir arbeiten jedoch daran, Administratoren mehr Kontrolle über einzelne Skripts und Skripters zu geben.

## <a name="coverage"></a>Abdeckung

Derzeit bietet VBA eine umfassendere Abdeckung Excel Funktionen, insbesondere die auf dem Desktop-Client verfügbaren Funktionen. Office Skripts decken fast alle Szenarien für Excel im Web ab. Darüber hinaus werden Office Scripts sie als neues Features-Debüt im Web sowohl für die Action Recorder- als auch für JavaScript-APIs unterstützen.

Office Skripts unterstützen keine [Ereignisse](/office/vba/excel/concepts/events-worksheetfunctions-shapes/using-events-with-excel-objects)auf Excel Ebene . Skripts werden nur ausgeführt, wenn sie manuell von einem Benutzer gestartet werden oder wenn ein Power Automate das Skript aufruft.

## <a name="power-automate"></a>Power Automate

Office Skripts können über Power Automate ausgeführt werden. Ihre Arbeitsmappe kann durch geplante oder ereignisgesteuerte Flows aktualisiert werden, sodass Sie Workflows automatisieren können, ohne auch nur Excel zu öffnen. Dies bedeutet, dass, solange Ihre Arbeitsmappe in OneDrive gespeichert ist (und für Power Automate zugänglich ist), ein Flow Ihre Skripts ausführen kann, unabhängig davon, ob Sie und Ihre Organisation den Desktop, Mac oder Webclient von Excel verwenden.

VBA verfügt nicht über einen Power Automate-Anschluss. Alle unterstützten VBA-Szenarien umfassen einen Benutzer, der sich um die Ausführung des Makros kümmert.

Probieren Sie die [Anrufskripts aus einem manuellen Power Automate-Flow-Tutorial](../tutorials/excel-power-automate-manual.md) aus, um mehr über Power Automate zu erfahren. Sie können auch das Beispiel für [automatische Aufgabenerinnerungen auschecken,](scenarios/task-reminders.md) um zu sehen, Office Skripts, die mit Teams über Power Automate verbunden sind, in einem realen Szenario.

## <a name="see-also"></a>Siehe auch

- [Office-Skripts in Excel im Web](../overview/excel.md)
- [Ausführen Office Skripts mit Power Automate](../develop/power-automate-integration.md)
- [Unterschiede zwischen Office-Skripts und Office-Add-Ins](add-ins-differences.md)
- [Behandeln von Problemen mit Office-Skripts](../testing/troubleshooting.md)
- [VBA-Referenz für Excel](/office/vba/api/overview/excel)
