---
title: Office Speicher und Besitz von Skripts
description: Informationen dazu, Office Skripts in Microsoft OneDrive zwischen Besitzern gespeichert und übertragen werden.
ms.date: 11/13/2020
localization_priority: Normal
ms.openlocfilehash: 47b732399c3068bea78b027e01324bbd73a83bc7
ms.sourcegitcommit: f7a7aebfb687f2a35dbed07ed62ff352a114525a
ms.translationtype: MT
ms.contentlocale: de-DE
ms.lasthandoff: 05/06/2021
ms.locfileid: "52232529"
---
# <a name="office-scripts-file-storage-and-ownership"></a>Office Speicher und Besitz von Skripts

Office Skripts werden als **OSTS-Dateien** in Ihrem Microsoft OneDrive. Dadurch können Ihre Skripts außerhalb einer bestimmten Arbeitsmappe vorhanden sein. Ihre OneDrive steuern den freigegebenen Zugriff und die Berechtigungen für alle **#A0 des Skripts.** unabhängig von den Excel werden.

## <a name="file-storage"></a>Dateispeicher

Sie Office Skripts in Ihrem OneDrive. Die **OSTS-Dateien** befinden sich im Ordner **/Documents/Office Scripts/.** Alle an diesen **OSTS-Dateien** vorgenommenen Änderungen, z. B. das Umbenennen oder Löschen von Dateien, werden im Code-Editor und im Skriptkatalog angezeigt.

Skripts, die für eine Ihrer Arbeitsmappen freigegeben werden, verbleiben in der OneDrive. Sie werden nicht in einen Ihrer lokalen oder OneDrive kopiert, wenn Sie das freigegebene Skript in Excel. Die **Schaltfläche Kopieren erstellen** des Code-Editors speichert eine separate Kopie des Skripts in OneDrive. Änderungen an der Kopie wirken sich nicht auf das ursprüngliche Skript aus.

### <a name="script-folders"></a>Skriptordner

Das Hinzufügen von Ordnern zu OneDrive trägt dazu bei, ihre Skripts organisiert zu halten. Alle Ordner unter **/Documents/Office Scripts/** werden im Abschnitt **Meine Skripts** des Code-Editors angezeigt. Beachten Sie, dass diese Ordner nicht mithilfe des Code-Editors erstellt oder gelöscht werden können. Ebenso können Skripts nicht in Ordnern platziert oder mithilfe des Code-Editors über Ordner verschoben werden.

:::image type="content" source="../images/script-folders.png" alt-text="Das Dialogfeld Neues Skript im Code-Editor, in dem Skripts in Ordnern angezeigt werden, wie im Aufgabenbereich angezeigt":::

## <a name="file-ownership-and-retention"></a>Dateibesitz und Aufbewahrung

Office Skripts werden in der Datei eines Benutzers OneDrive. Sie folgen den aufbewahrungs- und löschungsrichtlinien, die von Microsoft OneDrive. Wenn Sie wissen möchten, wie Skripts behandelt werden, die von einem Benutzer erstellt und freigegeben wurden, der aus Ihrer Organisation entfernt wird, lesen Sie [OneDrive Aufbewahrungs-und Löschvorgänge](/onedrive/retention-and-deletion).

## <a name="see-also"></a>Siehe auch

- [Office-Skripts in Excel für das Web](https://support.microsoft.com/office/sharing-office-scripts-in-excel-for-the-web-226eddbc-3a44-4540-acfe-fccda3d1122b)
- [Behandeln von Problemen mit Office-Skripts](../testing/troubleshooting.md)
- [Office-Skripts-Einstellungen in M365](https://support.office.com/article/office-scripts-settings-in-m365-19d3c51a-6ca2-40ab-978d-60fa49554dcf)
- [Rückgängigmachen der Auswirkungen eines Office-Skripts](../testing/undo.md)
