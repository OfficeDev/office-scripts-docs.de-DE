---
title: Unterschiede zwischen Office Skripts und VBA-Makros
description: Das Verhalten und die API-Unterschiede zwischen Office Skripts und Excel VBA-Makros.
ms.date: 05/21/2021
localization_priority: Normal
ms.openlocfilehash: 0d94607902fa62e07ce378b94ec3b9c328937e16535b1882b6cad5bd76212b33
ms.sourcegitcommit: 75f7ed8c2d23a104acc293f8ce29ea580b4fcdc5
ms.translationtype: MT
ms.contentlocale: de-DE
ms.lasthandoff: 08/11/2021
ms.locfileid: "57847268"
---
# <a name="differences-between-office-scripts-and-vba-macros"></a>Unterschiede zwischen Office Skripts und VBA-Makros

Office Skripts und VBA-Makros haben viele Gemeinsamkeiten. Beide ermöglichen Benutzern die Automatisierung von Lösungen über einen einfach zu verwendenden Action Recorder und ermöglichen die Bearbeitung dieser Aufzeichnungen. Beide Frameworks sind so konzipiert, dass Benutzer, die sich nicht selbst als Programmierer betrachten, kleine Programme in Excel erstellen können.
Der grundlegende Unterschied besteht darin, dass VBA-Makros für Desktoplösungen und Office Skripts für sichere, cloudbasierte Lösungen entwickelt werden. Derzeit werden Office Skripts nur in Excel im Web unterstützt.

:::image type="content" source="../images/office-programmability-diagram.png" alt-text="Ein Vier-Quadranten-Diagramm mit den Schwerpunktbereichen für verschiedene Office Erweiterbarkeitslösungen. Sowohl Office Skripts als auch VBA-Makros sollen Endbenutzern helfen, Lösungen zu erstellen, aber Office Skripts werden für das Web und die Zusammenarbeit erstellt (während VBA für den Desktop vorgesehen ist).":::

In diesem Artikel werden die Hauptunterschiede zwischen VBA-Makros (sowie VBA im Allgemeinen) und Office Skripts beschrieben. Da Office Skripts nur für Excel verfügbar sind, ist dies der einzige Host, der hier erläutert wird.

## <a name="platform-and-ecosystem"></a>Plattform und Ökosystem

VBA wurde für den Desktop entwickelt, und Office Skripts sind für das Web konzipiert. VBA kann mit dem Desktop eines Benutzers interagieren, um eine Verbindung mit ähnlichen Technologien wie COM und OLE herzustellen. VBA bietet jedoch keine bequeme Möglichkeit zum Aufrufen des Internets.

Office Skripts verwenden eine universelle Laufzeit für JavaScript. Dies sorgt für konsistentes Verhalten und Barrierefreiheit, unabhängig davon, auf welchem Computer das Skript ausgeführt wird. Sie können auch Aufrufe an andere Webdienste senden.

## <a name="security"></a>Sicherheit

VBA-Makros verfügen über dieselbe Sicherheitsfreigabe wie Excel. Dadurch erhalten sie Vollzugriff auf Ihren Desktop. Office Skripts haben nur Zugriff auf die Arbeitsmappe, nicht auf den Computer, auf dem die Arbeitsmappe gehostet wird. Darüber hinaus können keine JavaScript-Authentifizierungstoken für Skripts freigegeben werden. Dies bedeutet, dass das Skript weder über die Token des angemeldeten Benutzers verfügt noch über API-Funktionen zum Anmelden bei einem externen Dienst verfügt, sodass es nicht in der Lage ist, vorhandene Token für externe Aufrufe im Auftrag des Benutzers zu verwenden.

Administratoren haben drei Optionen für VBA-Makros: alle Makros auf dem Mandanten zulassen, keine Makros auf dem Mandanten zulassen oder nur Makros mit signierten Zertifikaten zulassen. Aufgrund dieser fehlenden Granularität ist es schwierig, einen einzelnen ungültigen Akteur zu isolieren. Derzeit können Office Skripts für einen gesamten Mandanten, für einen gesamten Mandanten oder für eine Gruppe von Benutzern in einem Mandanten deaktiviert sein. Administratoren haben auch die Kontrolle darüber, wer Skripts für andere personen freigeben kann und wer Skripts in Power Automate verwenden kann.

## <a name="coverage"></a>Abdeckung

Derzeit bietet VBA eine umfassendere Abdeckung von Excel Features, insbesondere diejenigen, die auf dem Desktopclient verfügbar sind. Office Skripts umfassen fast alle Szenarien für Excel im Web. Darüber hinaus unterstützen Office Skripts diese bei der Einführung neuer Features im Web sowohl für die Action Recorder- als auch die JavaScript-APIs.

Office Skripts unterstützen keine [Ereignisse](/office/vba/excel/concepts/events-worksheetfunctions-shapes/using-events-with-excel-objects)auf Excel Ebene. Skripts werden nur ausgeführt, wenn ein Benutzer sie manuell startet oder wenn ein Power Automate Fluss das Skript aufruft.

## <a name="power-automate"></a>Power Automate

Office Skripts können über Power Automate ausgeführt werden. Ihre Arbeitsmappe kann über geplante oder ereignisgesteuerte Flüsse aktualisiert werden, sodass Sie Workflows automatisieren können, ohne Excel zu öffnen. Dies bedeutet, dass, solange Ihre Arbeitsmappe in OneDrive gespeichert ist (und für Power Automate zugänglich ist), ein Fluss Ihre Skripts unabhängig davon ausführen kann, ob Sie und Ihre Organisation den Desktop-, Mac- oder Webclient Excel verwenden.

VBA verfügt nicht über einen Power Automate Connector. Alle unterstützten VBA-Szenarien umfassen einen Benutzer, der an der Ausführung des Makros teilnimmt.

Probieren Sie die [Call-Skripts aus einem manuellen Power Automate-Flow-Lernprogramm](../tutorials/excel-power-automate-manual.md) aus, um mehr über Power Automate zu erfahren. Sie können sich auch das Beispiel für [automatisierte Aufgabenerinnerungen](scenarios/task-reminders.md) ansehen, um Office Skripts anzuzeigen, die mit Teams über Power Automate in einem realen Szenario verbunden sind.

## <a name="see-also"></a>Weitere Artikel

- [Office-Skripts in Excel im Web](../overview/excel.md)
- [Ausführen von Office Skripts mit Power Automate](../develop/power-automate-integration.md)
- [Unterschiede zwischen Office-Skripts und Office-Add-Ins](add-ins-differences.md)
- [Behandeln von Problemen mit Office-Skripts](../testing/troubleshooting.md)
- [VBA-Referenz für Excel](/office/vba/api/overview/excel)
