---
title: Office Skripts– Dateispeicher und -besitz
description: Informationen dazu, wie Office Skripts in Microsoft OneDrive gespeichert und zwischen Besitzern übertragen werden.
ms.date: 05/10/2022
ms.localizationpriority: medium
ms.openlocfilehash: 5e2bc89db54ee5520c3b911ebd0f182777a78e2b
ms.sourcegitcommit: 8ae932e8b4e521fec8576ab16126eb9fe22a8dd7
ms.translationtype: MT
ms.contentlocale: de-DE
ms.lasthandoff: 05/11/2022
ms.locfileid: "65310757"
---
# <a name="office-scripts-file-storage-and-ownership"></a>Office Skripts– Dateispeicher und -besitz

Office Skripts werden als **OSTS-Dateien** in Ihrem Microsoft OneDrive gespeichert. Sie werden getrennt von einer Arbeitsmappe gespeichert. Wenn Sie anderen Zugriff gewähren möchten, [geben Sie das Skript für eine Excel Arbeitsmappe frei](excel.md#share-office-scripts). Dies bedeutet, dass Sie das Skript mit der Datei verknüpfen und nicht anfügen. Wer Zugriff auf die Excel Datei hat, kann das Skript auch anzeigen, ausführen oder eine Kopie davon erstellen.

Wenn Sie Ihre Skripts nicht freigeben, kann niemand anderes darauf zugreifen. Ihre OneDrive-Einstellungen steuern den freigegebenen Zugriff und die Berechtigungen für alle **OSTS-Skriptdateien**, unabhängig von Excel Einstellungen. Skripts können nicht von einem lokalen Datenträger oder benutzerdefinierten Cloudspeicherorten verknüpft werden. Office Skripts erkennt und führt ein Skript nur aus, wenn es sich in Ihrem OneDrive Ordner befindet oder für die Arbeitsmappe freigegeben ist.

## <a name="file-storage"></a>Dateispeicher

Sie Office Skripts in Ihrem OneDrive gespeichert werden. Die **OSTS-Dateien** befinden sich im Ordner **"/Documents/Office Scripts/**". Alle Änderungen an diesen **OSTS-Dateien** , z. B. Umbenennen oder Löschen von Dateien, werden im Code-Editor und im Skriptkatalog angezeigt.

Skripts, die für eine Ihrer Arbeitsmappen freigegeben sind, verbleiben im OneDrive des Skripterstellers. Sie werden nicht in ihre lokalen oder OneDrive Ordner kopiert, wenn Sie das freigegebene Skript in Excel ausführen. Die Schaltfläche "**Kopie erstellen**" des Code-Editors speichert eine separate Kopie des Skripts in Ihrem OneDrive. Änderungen an der Kopie wirken sich nicht auf das ursprüngliche Skript aus.

### <a name="restore-deleted-scripts"></a>Wiederherstellen gelöschter Skripts

Wenn Sie ein Skript in Excel löschen, wird es in Ihren OneDrive Papierkorb verschoben. Führen Sie zum Wiederherstellen eines gelöschten Skripts die unter ["Wiederherstellen gelöschter Dateien oder Ordner in OneDrive" aufgeführten](https://support.microsoft.com/office/949ada80-0026-4db3-a953-c99083e6a84f) Schritte aus. Durch das Wiederherstellen einer **OSTS-Datei** wird sie in die Liste **"Alle Skripts** " zurückgegeben.

Ein gelöschtes Skript wird mit der Arbeitsmappe nicht freigegeben. Wenn Sie ein Skript wiederherstellen, behält es **seinen** Skriptzugriff nicht bei. Sie müssen das Skript erneut freigeben.

Wiederhergestellte Skripts funktionieren mit Power Automate Flüssen weiterhin erwartungsgemäß. Sie müssen den Flussverbinder nicht neu erstellen.

## <a name="file-ownership-and-retention"></a>Dateibesitz und -aufbewahrung

Office Skripts werden im OneDrive eines Benutzers gespeichert. Sie folgen den durch Microsoft OneDrive angegebenen Aufbewahrungs- und Löschrichtlinien. Wenn Sie wissen möchten, wie Skripts behandelt werden, die von einem Benutzer erstellt und freigegeben wurden, der aus Ihrer Organisation entfernt wird, lesen Sie [OneDrive Aufbewahrungs-und Löschvorgänge](/onedrive/retention-and-deletion).

Während der Bearbeitung werden Dateien vorübergehend im Browser gespeichert. Sie müssen das Skript speichern, bevor Sie das Excel Fenster schließen, um es am OneDrive Speicherort zu speichern. Vergessen Sie nicht, die Datei nach Bearbeitungen zu speichern. Andernfalls befinden sich diese Bearbeitungen nur in der Browserversion der Datei.

## <a name="audit-office-scripts-usage-at-the-admin-level"></a>Überwachen der Verwendung von Office Skripts auf Administratorebene

Ermitteln Sie mithilfe des Überwachungsprotokolls im Compliance Center, welche Mandanten Office Skripts verwenden. Informationen zur Verwendung dieses Tools finden Sie [im Security & Compliance Center unter Durchsuchen des Überwachungsprotokolls](/microsoft-365/compliance/search-the-audit-log-in-security-and-compliance?view=o365-worldwide&preserve-view=true#search-the-audit-log).

Um zu finden, wer Office Skripts mit dem Suchtool verwendet, fügen Sie das **Feld "Datei", "Ordner" oder "Website**" hinzu`.osts`. Dadurch werden alle Dateien mit der Dateierweiterung Office Scripts gesucht. Wenn jemand in Ihrer Organisation das Feature Office Skripts verwendet hat, wird die Benutzeraktivität in den Suchergebnissen des Überwachungsprotokolls angezeigt.

## <a name="see-also"></a>Siehe auch

- [Office-Skripts in Excel für das Web](https://support.microsoft.com/office/226eddbc-3a44-4540-acfe-fccda3d1122b)
- [Behandeln von Problemen mit Office-Skripts](../testing/troubleshooting.md)
- [Office-Skripts-Einstellungen in M365](/microsoft-365/admin/manage/manage-office-scripts-settings)
- [Auswirkungen von Office-Skripts rückgängig machen](../testing/undo.md)
