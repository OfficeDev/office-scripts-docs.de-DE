---
title: Office Dateispeicherung und -besitz von Skripts
description: Informationen dazu, wie Office Skripts in Microsoft OneDrive gespeichert und zwischen Besitzern übertragen werden.
ms.date: 06/04/2021
localization_priority: Normal
ms.openlocfilehash: 788850db9c8e07ad59ea6d42eb9958779efcb06f
ms.sourcegitcommit: 6654aeae8a3ee2af84b4d4c4d8ff45b360a303eb
ms.translationtype: MT
ms.contentlocale: de-DE
ms.lasthandoff: 09/02/2021
ms.locfileid: "58862159"
---
# <a name="office-scripts-file-storage-and-ownership"></a>Office Dateispeicherung und -besitz von Skripts

Office Skripts werden als **OSTS-Dateien** in Ihrem Microsoft OneDrive gespeichert. Sie werden getrennt von einer Arbeitsmappe gespeichert. Um anderen Zugriff zu gewähren, [geben Sie das Skript mit einer Excel Arbeitsmappe frei.](excel.md#sharing-scripts) Dies bedeutet, dass Sie das Skript mit der Datei verknüpfen und nicht anfügen. Wer Zugriff auf die Excel Datei hat, kann auch eine Kopie des Skripts anzeigen, ausführen oder erstellen.

Wenn Sie Ihre Skripts nicht freigeben, kann kein anderer Benutzer darauf zugreifen. Ihre OneDrive-Einstellungen steuern den freigegebenen Zugriff  und die Berechtigungen für alle OSTS-Skriptdateien, unabhängig von Excel Einstellungen. Skripts können nicht von einem lokalen Datenträger oder benutzerdefinierten Cloudspeicherorten verknüpft werden. Office Skripts erkennen und führen ein Skript nur aus, wenn es sich in Ihrem OneDrive Ordner befindet oder für die Arbeitsmappe freigegeben ist.

## <a name="file-storage"></a>Dateispeicher

Sie Office Skripts werden in Ihrem OneDrive gespeichert. Die **OSTS-Dateien** befinden sich im Ordner **"/Documents/Office Scripts/".** Alle an diesen **OSTS-Dateien** vorgenommenen Änderungen, z. B. das Umbenennen oder Löschen von Dateien, werden im Code-Editor und im Skriptkatalog wiedergegeben.

Skripts, die für eine Ihrer Arbeitsmappen freigegeben sind, verbleiben im OneDrive des Skripterstellers. Sie werden nicht in Ihre lokalen oder OneDrive Ordner kopiert, wenn Sie das freigegebene Skript in Excel ausführen. Die Schaltfläche **"Kopieren"** des Code-Editors speichert eine separate Kopie des Skripts in Ihrer OneDrive. Änderungen an der Kopie wirken sich nicht auf das Originalskript aus.

### <a name="restore-deleted-scripts"></a>Wiederherstellen gelöschter Skripts

Wenn Sie ein Skript in Excel löschen, wird es in Den OneDrive Papierkorb verschoben. Führen Sie zum Wiederherstellen eines gelöschten Skripts die unter [Wiederherstellen gelöschter Dateien oder Ordner in OneDrive](https://support.microsoft.com/office/949ada80-0026-4db3-a953-c99083e6a84f)aufgeführten Schritte aus. Durch Wiederherstellen einer **OSTS-Datei** wird sie an die Liste **"Alle Skripts"** zurückgegeben.

Ein gelöschtes Skript ist nicht mit der Arbeitsmappe verbunden. Wenn Sie ein Skript wiederherstellen, behält es den Skriptzugriff **nicht** bei. Sie müssen das Skript erneut freigeben.

Wiederhergestellte Skripts funktionieren weiterhin wie erwartet mit Power Automate Flüssen. Sie müssen den Flusskonnektor nicht neu erstellen.

## <a name="file-ownership-and-retention"></a>Dateibesitz und -aufbewahrung

Office Skripts werden im OneDrive eines Benutzers gespeichert. Sie folgen den Von Microsoft OneDrive angegebenen Aufbewahrungs- und Löschrichtlinien. Wenn Sie wissen möchten, wie Skripts behandelt werden, die von einem Benutzer erstellt und freigegeben wurden, der aus Ihrer Organisation entfernt wird, lesen Sie [OneDrive Aufbewahrungs-und Löschvorgänge](/onedrive/retention-and-deletion).

Während der Bearbeitung werden Dateien vorübergehend im Browser gespeichert. Sie müssen das Skript speichern, bevor Sie das Excel Fenster schließen, um es am OneDrive Speicherort zu speichern. Vergessen Sie nicht, die Datei nach den Bearbeitungen zu speichern, andernfalls befinden sich diese Bearbeitungen nur in der Browserversion der Datei.

## <a name="audit-office-scripts-usage-at-the-admin-level"></a>Überwachen der Verwendung von Office Skripts auf Administratorebene

Ermitteln Sie, welche Mandanten Office Skripts mit dem Überwachungsprotokoll im Compliance Center verwenden. Informationen zur Verwendung dieses Tools finden Sie unter [Durchsuchen des Überwachungsprotokolls im Security & Compliance Center.](/microsoft-365/compliance/search-the-audit-log-in-security-and-compliance?view=o365-worldwide&preserve-view=true#search-the-audit-log)

Um zu ermitteln, wer Office Skripts mit dem Suchtool verwendet, fügen Sie `.osts` das **Feld "Datei", "Ordner" oder "Website"** hinzu. Dadurch werden alle Dateien mit der Dateierweiterung Office Scripts gesucht. Wenn jemand in Ihrer Organisation das Feature Office Skripts verwendet hat, wird die Benutzeraktivität in den Suchergebnissen des Überwachungsprotokolls angezeigt.

> [!NOTE]
> Das Ausführen eines Skripts wird derzeit nicht protokolliert. Es werden nur die Erstellungs-, Ansichts- und Änderungsaktionen protokolliert.

## <a name="see-also"></a>Siehe auch

- [Office-Skripts in Excel für das Web](https://support.microsoft.com/office/226eddbc-3a44-4540-acfe-fccda3d1122b)
- [Behandeln von Problemen mit Office-Skripts](../testing/troubleshooting.md)
- [Office-Skripts-Einstellungen in M365](/microsoft-365/admin/manage/manage-office-scripts-settings)
- [Auswirkungen von Office-Skripts rückgängig machen](../testing/undo.md)
