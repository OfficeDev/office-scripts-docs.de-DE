---
title: Speicher und Besitz von Office Scripts-Dateien
description: 'Informationen dazu, wie #A0 in Microsoft OneDrive gespeichert und zwischen Besitzern übertragen werden.'
ms.date: 11/13/2020
localization_priority: Normal
ms.openlocfilehash: bd868c1dbfd0b33d3cd9fc4ee774c654d86f9b07
ms.sourcegitcommit: 45ffe3dbd2c834b78592ad35928cf8096f5e80bc
ms.translationtype: MT
ms.contentlocale: de-DE
ms.lasthandoff: 04/14/2021
ms.locfileid: "51755105"
---
# <a name="office-scripts-file-storage-and-ownership"></a>Speicher und Besitz von Office Scripts-Dateien

#A0 werden als **#A1** in Ihrem Microsoft OneDrive gespeichert. Dadurch können Ihre Skripts außerhalb einer bestimmten Arbeitsmappe vorhanden sein. Ihre #A0 steuern den freigegebenen Zugriff und die Berechtigungen für alle **#A1 des Skripts.** unabhängig von den Excel-Einstellungen.

## <a name="file-storage"></a>Dateispeicher

Sie #A0 werden in Ihrem OneDrive gespeichert. Die **OSTS-Dateien** befinden sich im Ordner **/Documents/Office Scripts/.** Alle an diesen **OSTS-Dateien** vorgenommenen Änderungen, z. B. das Umbenennen oder Löschen von Dateien, werden im Code-Editor und im Skriptkatalog angezeigt.

Skripts, die für eine Ihrer Arbeitsmappen freigegeben werden, verbleiben im OneDrive des Skripterstellers. Sie werden nicht in Ihren lokalen oder #A0 kopiert, wenn Sie das freigegebene Skript in Excel ausführen. Die **Schaltfläche Kopieren erstellen** des #A0 speichert eine separate Kopie des Skripts in Ihrem OneDrive. Änderungen an der Kopie wirken sich nicht auf das ursprüngliche Skript aus.

### <a name="script-folders"></a>Skriptordner

Das Hinzufügen von Ordnern zu Ihrem OneDrive trägt dazu bei, dass Ihre Skripts organisiert bleiben. Alle Ordner unter **/Documents/Office Scripts/** werden im Abschnitt **Meine Skripts** des Code-Editors angezeigt. Beachten Sie, dass diese Ordner nicht mithilfe des Code-Editors erstellt oder gelöscht werden können. Ebenso können Skripts nicht in Ordnern platziert oder mithilfe des Code-Editors über Ordner verschoben werden.

:::image type="content" source="../images/script-folders.png" alt-text="Das Dialogfeld Neues Skript im Code-Editor zeigt Skripts in Ordnern an, wie im Aufgabenbereich angezeigt.":::

## <a name="file-ownership-and-retention"></a>Dateibesitz und Aufbewahrung

#A0 werden im OneDrive eines Benutzers gespeichert. Sie folgen den von Microsoft OneDrive angegebenen Aufbewahrungs- und Löschrichtlinien. Wenn Sie wissen möchten, wie Skripts behandelt werden, die von einem Benutzer erstellt und freigegeben wurden, der aus Ihrer Organisation entfernt wird, lesen Sie [OneDrive Aufbewahrungs-und Löschvorgänge](/onedrive/retention-and-deletion).

## <a name="see-also"></a>Weitere Informationen

- [Office-Skripts in Excel für das Web](https://support.microsoft.com/office/sharing-office-scripts-in-excel-for-the-web-226eddbc-3a44-4540-acfe-fccda3d1122b)
- [Behandeln von Problemen mit Office-Skripts](../testing/troubleshooting.md)
- [Office-Skripts-Einstellungen in M365](https://support.office.com/article/office-scripts-settings-in-m365-19d3c51a-6ca2-40ab-978d-60fa49554dcf)
- [Rückgängigmachen der Auswirkungen eines Office-Skripts](../testing/undo.md)
