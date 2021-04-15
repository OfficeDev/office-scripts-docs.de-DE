---
title: Unterschiede zwischen Office-Skripts und VBA-Makros
description: Das Verhalten und die API-Unterschiede zwischen Office-Skripts und Excel-VBA-Makros.
ms.date: 12/14/2020
localization_priority: Normal
ms.openlocfilehash: a56409a5de3eb07876faa88bfbfe78eeca59f70f
ms.sourcegitcommit: 45ffe3dbd2c834b78592ad35928cf8096f5e80bc
ms.translationtype: MT
ms.contentlocale: de-DE
ms.lasthandoff: 04/14/2021
ms.locfileid: "51755021"
---
# <a name="differences-between-office-scripts-and-vba-macros"></a>Unterschiede zwischen Office-Skripts und VBA-Makros

Office-Skripts und VBA-Makros haben viele gemeinsam. Sie ermöglichen es Benutzern, Lösungen mithilfe einer einfach zu verwendenden Aktionsaufzeichnung zu automatisieren und Bearbeitungen dieser Aufzeichnungen zu ermöglichen. Beide Frameworks sollen Personen, die sich möglicherweise nicht als Programmierer betrachten, die Erstellung kleiner Programme in Excel ermöglichen.
Der wesentliche Unterschied besteht in der Entwicklung von VBA-Makros für Desktoplösungen, und Office-Skripts sind mit plattformübergreifender Unterstützung und Sicherheit als Leitprinzipien konzipiert. Derzeit werden Office-Skripts nur in Excel im Web unterstützt.

:::image type="content" source="../images/office-programmability-diagram.png" alt-text="Ein Diagramm mit vier Quadranten, das die Fokusbereiche für verschiedene Office-Erweiterbarkeitslösungen zeigt. Sowohl Office-Skripts als auch VBA-Makros sollen Endbenutzern beim Erstellen von Lösungen helfen, aber Office-Skripts sind für das Web und die Zusammenarbeit entwickelt (während VBA für den Desktop verwendet wird).":::

In diesem Artikel werden die wichtigsten Unterschiede zwischen VBA-Makros (sowie VBA im Allgemeinen) und Office-Skripts beschrieben. Da Office-Skripts nur für Excel verfügbar sind, ist dies der einzige Host, der hier behandelt wird.

## <a name="platform-and-ecosystem"></a>Plattform und Ökosystem

VBA ist für den Desktop konzipiert, und Office-Skripts sind für das Web konzipiert. VBA kann mit dem Desktop eines Benutzers interagieren, um eine Verbindung mit ähnlichen Technologien wie COM und OLE herzustellen. VBA hat jedoch keine bequeme Möglichkeit, das Internet an sich zu rufen.

Office-Skripts verwenden eine universelle Laufzeit für JavaScript. Dies bietet konsistentes Verhalten und Barrierefreiheit, unabhängig vom Computer, der zum Ausführen des Skripts verwendet wird. Sie können auch Anrufe an andere Webdienste nehmen.

## <a name="security"></a>Sicherheit

VBA-Makros haben die gleiche Sicherheitsfreigabe wie Excel. Dadurch erhalten sie vollzugriff auf Ihren Desktop. Office-Skripts haben nur Zugriff auf die Arbeitsmappe, nicht auf den Computer, auf dem die Arbeitsmappe hostend ist. Darüber hinaus können keine JavaScript-Authentifizierungstoken für Skripts freigegeben werden. Dies bedeutet, dass das Skript weder über die Token des angemeldeten Benutzers noch über API-Funktionen für die Anmeldung bei einem externen Dienst verfügt, sodass es keine vorhandenen Token zum Ausführen externer Aufrufe im Namen des Benutzers verwenden kann.

Administratoren haben drei Optionen für VBA-Makros: Alle Makros für den Mandanten zulassen, keine Makros für den Mandanten zulassen oder nur Makros mit signierten Zertifikaten zulassen. Dieser Mangel an Granularität macht es schwierig, einen einzelnen schlechten Akteur zu isolieren. Derzeit sind Office-Skripts für einen Mandanten entweder ein- oder ausgeschaltet. Wir arbeiten jedoch daran, Administratoren mehr Kontrolle über einzelne Skripts und Skriptersteller zu geben.

## <a name="coverage"></a>Abdeckung

Derzeit bietet VBA eine vollständigere Abdeckung der Excel-Features, insbesondere derjenigen, die auf dem Desktopclient verfügbar sind. Office-Skripts umfassen fast alle Szenarien für Excel im Web. Darüber hinaus unterstützen Office-Skripts als debütierte neue Features im Web diese sowohl für die Action Recorder- als auch für die JavaScript-APIs.

Office-Skripts unterstützen keine Ereignisse auf [Excel-Ebene.](/office/vba/excel/concepts/events-worksheetfunctions-shapes/using-events-with-excel-objects) Skripts werden nur ausgeführt, wenn ein Benutzer sie manuell startet oder wenn ein Power Automate-Fluss das Skript aufruft.

## <a name="power-automate"></a>Power Automate

Office-Skripts können über Power Automate ausgeführt werden. Ihre Arbeitsmappe kann über geplante oder ereignisgesteuerte Flüsse aktualisiert werden, damit Sie Workflows automatisieren können, ohne excel zu öffnen. Dies bedeutet, dass ein Fluss Ihre Skripts unabhängig davon ausführen kann, ob Sie und Ihre Organisation den Desktop, Mac oder Webclient von Excel verwenden, solange Ihre Arbeitsmappe in OneDrive gespeichert ist (und auf den Power Automate zugegriffen werden kann).

VBA hat keinen Power Automate-Connector. Bei allen unterstützten VBA-Szenarien musste ein Benutzer an der Ausführung des Makros teilnehmen.

## <a name="see-also"></a>Weitere Informationen

- [Office-Skripts in Excel im Web](../overview/excel.md)
- [Unterschiede zwischen Office-Skripts und Office-Add-Ins](add-ins-differences.md)
- [Behandeln von Problemen mit Office-Skripts](../testing/troubleshooting.md)
- [VBA-Referenz für Excel](/office/vba/api/overview/excel)
