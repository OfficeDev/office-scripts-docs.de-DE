---
title: Office Skripts-Dateispeicherung und -besitz
description: Informationen darüber, wie Office Skripts in Microsoft OneDrive gespeichert und zwischen Besitzern übertragen werden.
ms.date: 05/17/2021
localization_priority: Normal
ms.openlocfilehash: 556d784dc1fe64873866c49ab2726a4c68abc1a7
ms.sourcegitcommit: 4687693f02fc90a57ba30c461f35046e02e6f5fb
ms.translationtype: MT
ms.contentlocale: de-DE
ms.lasthandoff: 05/19/2021
ms.locfileid: "52545801"
---
# <a name="office-scripts-file-storage-and-ownership"></a>Office Skripts-Dateispeicherung und -besitz

Office Skripts werden als **.osts-Dateien** in Ihrem Microsoft OneDrive gespeichert. Sie werden getrennt von einer Arbeitsmappe gespeichert. Um anderen Zugriff zu gewähren, [geben Sie das Skript für eine Excel Arbeitsmappe](excel.md#sharing-scripts). Dies bedeutet, dass Sie das Skript mit der Datei verknüpfen und nicht anfügen. Wer Zugriff auf die Excel Datei hat, kann auch das Skript anzeigen, ausführen oder eine Kopie erstellen.

Wenn Sie Ihre Skripts nicht freigeben, kann niemand sonst darauf zugreifen. Ihre OneDrive Einstellungen steuern den freigegebenen Zugriff und die Berechtigungen für alle **Skript.osts-Dateien,** unabhängig von Excel Einstellungen. Skripts können nicht von einem lokalen Datenträger oder benutzerdefinierten Cloudspeicherorten verknüpft werden. Office Skripts erkennt und führt ein Skript nur aus, wenn es sich in Ihrem OneDrive Ordner befindet oder für die Arbeitsmappe freigegeben ist.

## <a name="file-storage"></a>Dateispeicher

Sie Office Skripts in Ihrem OneDrive gespeichert werden. Die **.osts-Dateien** befinden sich im Ordner **/Documents/Office Scripts/.** Alle an diesen **.osts-Dateien vorgenommenen** Bearbeitungen, wie z. B. das Umbenennen oder Löschen von Dateien, werden im Code-Editor und im Skriptkatalog widergespiegelt.

Skripts, die für eine Ihrer Arbeitsmappen freigegeben sind, verbleiben im OneDrive des Skripters. Sie werden nicht in einen Ihrer lokalen oder OneDrive Ordner kopiert, wenn Sie das freigegebene Skript in Excel ausführen. Mit der Schaltfläche **Kopieren** des Code-Editors wird eine separate Kopie des Skripts in Ihrem OneDrive. Änderungen an der Kopie wirken sich nicht auf das ursprüngliche Skript aus.

## <a name="file-ownership-and-retention"></a>Dateibesitz und -aufbewahrung

Office Skripts werden im OneDrive eines Benutzers gespeichert. Sie folgen den von Microsoft OneDrive angegebenen Aufbewahrungs- und Löschrichtlinien. Wenn Sie wissen möchten, wie Skripts behandelt werden, die von einem Benutzer erstellt und freigegeben wurden, der aus Ihrer Organisation entfernt wird, lesen Sie [OneDrive Aufbewahrungs-und Löschvorgänge](/onedrive/retention-and-deletion).

Während der Bearbeitung werden Dateien vorübergehend im Browser gespeichert. Sie müssen das Skript speichern, bevor Sie das Excel Fenster schließen, um es am OneDrive Speicherort zu speichern. Vergessen Sie nicht, die Datei nach Denkarbeiten zu speichern, sonst befinden sich diese Änderungen nur in der Browserversion der Datei.

## <a name="see-also"></a>Siehe auch

- [Office-Skripts in Excel für das Web](https://support.microsoft.com/office/sharing-office-scripts-in-excel-for-the-web-226eddbc-3a44-4540-acfe-fccda3d1122b)
- [Behandeln von Problemen mit Office-Skripts](../testing/troubleshooting.md)
- [Office-Skripts-Einstellungen in M365](https://support.office.com/article/office-scripts-settings-in-m365-19d3c51a-6ca2-40ab-978d-60fa49554dcf)
- [Auswirkungen von Office-Skripts rückgängig machen](../testing/undo.md)
