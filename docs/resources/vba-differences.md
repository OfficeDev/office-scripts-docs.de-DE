---
title: Unterschiede zwischen Office-Skripts und VBA-Makros
description: Das Verhalten und die API-Unterschiede zwischen Office-Skripts und Excel-VBA-Makros.
ms.date: 06/30/2020
localization_priority: Normal
ms.openlocfilehash: 8a8929f0c6a73a8e9041bb4b55cce1edd539e166
ms.sourcegitcommit: bf9f33c37c6f7805d6b408aa648bb9785a7cd133
ms.translationtype: MT
ms.contentlocale: de-DE
ms.lasthandoff: 07/06/2020
ms.locfileid: "45043391"
---
# <a name="differences-between-office-scripts-and-vba-macros"></a>Unterschiede zwischen Office-Skripts und VBA-Makros

Office-Skripts und VBA-Makros haben viele Gemeinsamkeiten. Beide ermöglichen Benutzern das Automatisieren von Lösungen über einen einfach zu bedienende Action Recorder und erlauben das Bearbeiten dieser Aufzeichnungen. Beide Frameworks sind darauf ausgelegt, Personen, die sich möglicherweise nicht als Programmierer bezeichnen, das Erstellen von kleinen Programmen in Excel zu ermöglichen.
Der wesentliche Unterschied besteht darin, dass VBA-Makros für Desktop Lösungen entwickelt werden und Office-Skripts mit plattformübergreifender Unterstützung und Sicherheit als Leitprinzipien konzipiert wurden. Zurzeit werden Office-Skripts nur in Excel im Internet unterstützt.

![Ein Diagramm mit vier Quadranten, in dem die Fokusbereiche für unterschiedliche Office-Erweiterbarkeits Lösungen dargestellt werden. Sowohl Office-Skripts als auch VBA-Makros sollen Endbenutzern beim Erstellen von Lösungen helfen, aber Office-Skripts werden für das Internet und die Zusammenarbeit erstellt (während VBA für den Desktop gilt).)](../images/office-programmability-diagram.png)

In diesem Artikel werden die wichtigsten Unterschiede zwischen VBA-Makros (sowie VBA im allgemeinen) und Office-Skripts beschrieben. Da Office-Skripts nur für Excel zur Verfügung stehen, ist dies der einzige Host, der hier erörtert wird.

## <a name="platform-and-ecosystem"></a>Plattform und Ökosystem

VBA wurde für den Desktop entwickelt, und Office-Skripts sind für das Internet konzipiert. VBA kann mit dem Desktop eines Benutzers interagieren, um eine Verbindung mit ähnlichen Technologien wie com und OLE herzustellen. VBA verfügt jedoch nicht über eine bequeme Möglichkeit zum Aufrufen des Internets.

Office-Skripts verwenden eine universelle Laufzeit oder JavaScript. Dadurch wird konsistentes Verhalten und Barrierefreiheit gewährleistet, unabhängig davon, welcher Computer zum Ausführen des Skripts verwendet wird. Sie können auch Anrufe an andere Webdienste vornehmen.

## <a name="security"></a>Sicherheit

VBA-Makros haben denselben Sicherheitsabstand wie Excel. Dadurch erhalten Sie vollen Zugriff auf Ihren Desktop. Office-Skripts haben nur Zugriff auf die Arbeitsmappe, nicht auf den Computer, der die Arbeitsmappe hostet. Darüber hinaus können keine JavaScript-Authentifizierungstoken für Skripts freigegeben werden, sodass Skripts sich nie bei einem externen Dienst authentifizieren können.

Administratoren haben drei Optionen für VBA-Makros: alle Makros im Mandanten zulassen, keine Makros auf dem Mandanten zulassen oder nur Makros mit signierten Zertifikaten zulassen. Durch diesen Mangel an Granularität ist es schwierig, einen einzelnen fehlerhaften Akteur zu isolieren. Office-Skripts sind derzeit entweder für einen Mandanten ein-oder ausgeschaltet. Wir arbeiten jedoch daran, Administratoren mehr Kontrolle über einzelne Skripts und Skriptersteller zu geben.

## <a name="coverage"></a>Abdeckung

Derzeit bietet VBA eine umfassendere Abdeckung der Excel-Features, insbesondere die auf dem Desktop Client verfügbaren. Office-Skripts umfassen nahezu alle Szenarien für Excel im Internet. Da neue Features im Internet debütieren, werden diese von Office-Skripts sowohl für die aktionsaufzeichnung als auch für JavaScript-APIs unterstützt.

## <a name="power-automate"></a>Power Automate

Office-Skripts können über Power Automation ausgeführt werden. Ihre Arbeitsmappe kann über geplante oder ereignisgesteuerte Abläufe aktualisiert werden, sodass Sie Workflows automatisieren können, ohne Excel zu öffnen. Dies bedeutet, dass, solange Ihre Arbeitsmappe in OneDrive gespeichert ist (und für Power Automation zugänglich), ein Flow Ihre Skripts ausführen kann, unabhängig davon, ob Sie und Ihre Organisation den Desktop, Mac oder Webclient von Excel verwenden.

VBA verfügt über keinen Power Automation-Connector. Alle unterstützten VBA-Szenarien betrafen einen Benutzer, der die Ausführung des Makros teilgenommen hat.

## <a name="see-also"></a>Siehe auch

- [Office-Skripts in Excel im Web](../overview/excel.md)
- [Unterschiede zwischen Office-Skripts und Office-Add-Ins](add-ins-differences.md)
- [Behandeln von Problemen mit Office-Skripts](../testing/troubleshooting.md)
- [VBA-Referenz für Excel](/office/vba/api/overview/excel)
