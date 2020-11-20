---
title: Dateispeicherung und Besitz von Office-Skripts
description: Informationen dazu, wie Office-Skripts in Microsoft OneDrive gespeichert und zwischen Besitzern übertragen werden.
ms.date: 11/13/2020
localization_priority: Normal
ms.openlocfilehash: 648f3b2cf7e7d8d3bab2cf07a090e116e267a99a
ms.sourcegitcommit: 82d3c0ef1e187bcdeceb2b5fc3411186674fe150
ms.translationtype: MT
ms.contentlocale: de-DE
ms.lasthandoff: 11/18/2020
ms.locfileid: "49346865"
---
# <a name="office-scripts-file-storage-and-ownership"></a>Dateispeicherung und Besitz von Office-Skripts

Office-Skripts werden als **Kosten** -Dateien in Ihrem Microsoft OneDrive gespeichert. Auf diese Weise können Ihre Skripts außerhalb einer bestimmten Arbeitsmappe vorhanden sein. Ihre OneDrive-Einstellungen steuern den freigegebenen Zugriff und die Berechtigungen für alle Script **. Kosten** -Dateien; unabhängig von Excel-Einstellungen.

## <a name="file-storage"></a>Dateispeicher

Ihre Office-Skripts werden in Ihrem OneDrive gespeichert. Die **Kosten** -Dateien werden in den **/Documents/Office-Skripts/-** Ordnern gefunden. Alle an diesen **Kosten** -Dateien vorgenommenen Änderungen, beispielsweise das Umbenennen oder Löschen von Dateien, werden im Code-Editor und im Skript Katalog wiedergegeben.

Skripts, die für eine ihrer Arbeitsmappen freigegeben sind, verbleiben im OneDrive des Skript Erstellers. Sie werden nicht in einen der lokalen oder OneDrive Ordner kopiert, wenn Sie das freigegebene Skript in Excel ausführen. Auf der Schaltfläche **Kopie erstellen** des Code-Editors wird eine separate Kopie des Skripts in Ihrem OneDrive gespeichert. Änderungen an der Kopie wirken sich nicht auf das ursprüngliche Skript aus.

### <a name="script-folders"></a>Skriptordner

Durch das Hinzufügen von Ordnern zu Ihrem OneDrive können Sie Ihre Skripts organisieren. Alle Ordner unter **/Documents/Office-Skripts/** werden im Abschnitt **meine Skripts** des Code-Editors angezeigt. Beachten Sie, dass diese Ordner nicht mithilfe des Code-Editors erstellt oder gelöscht werden können. Ebenso können Skripts nicht mithilfe des Code-Editors in Ordner verschoben oder in Ordner verschoben werden.

![Einige in Ordnern enthaltene Skripts, die im Aufgabenbereich Code-Editor angezeigt werden](../images/script-folders.png)

## <a name="file-ownership-and-retention"></a>Dateibesitz und-Aufbewahrung

Office-Skripts werden im OneDrive eines Benutzers gespeichert. Sie entsprechen den von Microsoft OneDrive angegebenen Aufbewahrungs-und Löschungsrichtlinien. Wenn Sie wissen möchten, wie Skripts behandelt werden, die von einem Benutzer erstellt und freigegeben wurden, der aus Ihrer Organisation entfernt wird, lesen Sie [OneDrive Aufbewahrungs-und Löschvorgänge](/onedrive/retention-and-deletion).

## <a name="see-also"></a>Siehe auch

- [Office-Skripts in Excel für das Web](https://support.microsoft.com/office/sharing-office-scripts-in-excel-for-the-web-226eddbc-3a44-4540-acfe-fccda3d1122b)
- [Behandeln von Problemen mit Office-Skripts](../testing/troubleshooting.md)
- [Office-Skripts-Einstellungen in M365](https://support.office.com/article/office-scripts-settings-in-m365-19d3c51a-6ca2-40ab-978d-60fa49554dcf)
- [Rückgängigmachen der Auswirkungen eines Office-Skripts](../testing/undo.md)
