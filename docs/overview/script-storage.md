---
title: Office Speicher und Besitz von Skripts
description: Informationen dazu, Office Skripts in Microsoft OneDrive zwischen Besitzern gespeichert und übertragen werden.
ms.date: 05/21/2021
localization_priority: Normal
ms.openlocfilehash: 25683d2b6ac2e8ac47b465b24fa087af83175806
ms.sourcegitcommit: 09d8859d5269ada8f1d0e141f6b5a4f96d95a739
ms.translationtype: MT
ms.contentlocale: de-DE
ms.lasthandoff: 05/24/2021
ms.locfileid: "52631657"
---
# <a name="office-scripts-file-storage-and-ownership"></a>Office Speicher und Besitz von Skripts

Office Skripts werden als **OSTS-Dateien** in Ihrem Microsoft OneDrive. Sie werden getrennt von einer Arbeitsmappe gespeichert. Um anderen Zugriff zu geben, [teilen Sie das Skript mit einer Excel Arbeitsmappe.](excel.md#sharing-scripts) Dies bedeutet, dass Sie das Skript mit der Datei verknüpfen und nicht anfügen. Wer Zugriff auf die Excel hat, kann auch eine Kopie des Skripts anzeigen, ausführen oder erstellen.

Wenn Sie Ihre Skripts nicht freigeben, kann niemand darauf zugreifen. Ihre OneDrive steuern den freigegebenen Zugriff und die Berechtigungen für alle **Skript-OSTS-Dateien,** unabhängig von Excel Einstellungen. Skripts können nicht von einem lokalen Datenträger oder benutzerdefinierten Cloudspeicherorten verknüpft werden. Office Skripts erkennen und ausführen ein Skript nur, wenn es sich in Ihrem ordner OneDrive oder für die Arbeitsmappe freigegeben ist.

## <a name="file-storage"></a>Dateispeicher

Sie Office Skripts in Ihrem OneDrive. Die **OSTS-Dateien** befinden sich im Ordner **/Documents/Office Scripts/.** Alle an diesen **OSTS-Dateien** vorgenommenen Änderungen, z. B. das Umbenennen oder Löschen von Dateien, werden im Code-Editor und im Skriptkatalog angezeigt.

Skripts, die für eine Ihrer Arbeitsmappen freigegeben werden, verbleiben in der OneDrive. Sie werden nicht in einen Ihrer lokalen oder OneDrive kopiert, wenn Sie das freigegebene Skript in Excel. Die **Schaltfläche Kopieren erstellen** des Code-Editors speichert eine separate Kopie des Skripts in OneDrive. Änderungen an der Kopie wirken sich nicht auf das ursprüngliche Skript aus.

## <a name="file-ownership-and-retention"></a>Dateibesitz und Aufbewahrung

Office Skripts werden in der Datei eines Benutzers OneDrive. Sie folgen den aufbewahrungs- und löschungsrichtlinien, die von Microsoft OneDrive. Wenn Sie wissen möchten, wie Skripts behandelt werden, die von einem Benutzer erstellt und freigegeben wurden, der aus Ihrer Organisation entfernt wird, lesen Sie [OneDrive Aufbewahrungs-und Löschvorgänge](/onedrive/retention-and-deletion).

Während der Bearbeitung werden Dateien vorübergehend im Browser gespeichert. Sie müssen das Skript speichern, bevor Sie das Excel schließen, um es am speicherort OneDrive speichern. Vergessen Sie nicht, die Datei nach den Bearbeitungen zu speichern, da sich diese Bearbeitungen sonst nur in der Version der Datei des Browsers finden.

## <a name="audit-office-scripts-usage-at-the-admin-level"></a>Überwachen Office Skriptverwendung auf Administratorebene

Ermitteln Sie, welche Mandanten Office Skripts mit dem Überwachungsprotokoll im Compliance Center verwenden. Informationen zur Verwendung dieses Tools finden Sie unter Durchsuchen des Überwachungsprotokolls im [Security & Compliance Center](/microsoft-365/compliance/search-the-audit-log-in-security-and-compliance?view=o365-worldwide#search-the-audit-log).

Um zu finden, wer Office Skripts mit dem Suchtool verwendet, fügen Sie das Feld `.osts` **Datei, Ordner oder Website** hinzu. Dadurch werden alle Dateien mit der Dateierweiterung Office Skripts gesucht. Wenn jemand in Ihrer Organisation das Feature Office Skripts verwendet hat, wird die Benutzeraktivität in den Suchergebnissen des Überwachungsprotokolls angezeigt.

> [!NOTE]
> Das Ausführen eines Skripts wird derzeit nicht protokolliert. Nur die Aktionen zum Erstellen, Anzeigen und Ändern werden protokolliert.

## <a name="see-also"></a>Sehen Sie ebenfalls

- [Office-Skripts in Excel für das Web](https://support.microsoft.com/office/sharing-office-scripts-in-excel-for-the-web-226eddbc-3a44-4540-acfe-fccda3d1122b)
- [Behandeln von Problemen mit Office-Skripts](../testing/troubleshooting.md)
- [Office-Skripts-Einstellungen in M365](https://support.office.com/article/office-scripts-settings-in-m365-19d3c51a-6ca2-40ab-978d-60fa49554dcf)
- [Auswirkungen von Office-Skripts rückgängig machen](../testing/undo.md)
