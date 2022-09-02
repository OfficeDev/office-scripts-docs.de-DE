---
title: Dateispeicher und Besitz von Office-Skripts
description: Informationen dazu, wie Office-Skripts in Microsoft OneDrive gespeichert und zwischen Besitzern übertragen werden.
ms.date: 08/31/2022
ms.localizationpriority: medium
ms.openlocfilehash: 573f65f299c29b4f481c9a2e23ebe7e36181706b
ms.sourcegitcommit: a6504f8b0d6b717457c6e0b5306c35ad3900914e
ms.translationtype: MT
ms.contentlocale: de-DE
ms.lasthandoff: 09/02/2022
ms.locfileid: "67572507"
---
# <a name="office-scripts-file-storage-and-ownership"></a>Dateispeicher und Besitz von Office-Skripts

Office-Skripts werden als **OSTS-Dateien** in Ihrem Microsoft OneDrive oder einem SharePoint-Ordner gespeichert. Sie werden getrennt von einer Arbeitsmappe gespeichert. Um Benutzern außerhalb der SharePoint-Website Zugriff auf das Skript zu gewähren, [geben Sie das Skript für eine Excel-Arbeitsmappe frei](excel.md#share-office-scripts). Dies bedeutet, dass Sie das Skript mit der Datei verknüpfen und nicht anfügen. Wer Zugriff auf die Excel-Datei hat, kann das Skript auch anzeigen, ausführen oder eine Kopie davon erstellen.

Excel erkennt und führt ein Skript nur aus, wenn es sich in Ihrem OneDrive-Ordner, einem SharePoint-Ordner oder in der Arbeitsmappe befindet.

## <a name="onedrive"></a>OneDrive

Das Standardverhalten ist, dass Office-Skripts auf Ihrem OneDrive gespeichert werden. Die **OSTS-Dateien** befinden sich im Ordner **"/Documents/Office Scripts/** ". Alle Änderungen an diesen **OSTS-Dateien** , z. B. Umbenennen oder Löschen von Dateien, werden im Code-Editor und im Skriptkatalog angezeigt.

Skripts, die für eine Ihrer Arbeitsmappen freigegeben sind, verbleiben im OneDrive des Skripterstellers. Sie werden nicht in Einen Ihrer lokalen oder OneDrive-Ordner kopiert, wenn Sie das freigegebene Skript in Excel ausführen. Die Schaltfläche " **Kopie erstellen** " des Code-Editors speichert eine separate Kopie des Skripts auf Ihrem OneDrive. Änderungen an der Kopie wirken sich nicht auf das ursprüngliche Skript aus.

Sofern Sie Ihre persönlichen Skripts nicht freigeben, kann niemand anderes darauf zugreifen. Ihre OneDrive-Einstellungen steuern den freigegebenen Zugriff und die Berechtigungen für alle **OSTS-Skriptdateien** , unabhängig von allen Excel-Einstellungen. Skripts können nicht von einem lokalen Datenträger oder benutzerdefinierten Cloudspeicherorten verknüpft werden.

## <a name="sharepoint"></a>SharePoint

Office-Skripts, die auf einer SharePoint-Website gespeichert werden, gehören Ihrem Team. Sie und Mitglieder Ihrer Organisation mit dem entsprechenden Zugriff können Skripts aus SharePoint ausführen und bearbeiten. Diese Skripts werden auch im Skriptkatalog der Registerkarte **"Automatisieren** " angezeigt.

Um ein Skript aus SharePoint zu laden, wechseln **Sie zu "Alle Skripts** ", und wählen Sie unten in der Liste " **Weitere Skripts anzeigen** " aus. Dadurch wird eine Dateiauswahl angezeigt, in der Sie **OSTS-Dateien** von jeder SharePoint-Website auswählen können, auf die Sie Zugriff haben. Beachten Sie, dass Skripts aus SharePoint, die Sie bereits geöffnet haben, in der Liste der zuletzt verwendeten Skripts angezeigt werden.

Um ein Skript in SharePoint zu speichern, wechseln Sie zum Menü **"Weitere Optionen (...)** ", und wählen Sie **"Speichern unter" aus**. Dadurch wird eine Dateiauswahl geöffnet, in der Sie Ordner auf Ihrer SharePoint-Website auswählen können. Beim Speichern an einem neuen Speicherort wird eine Kopie des Skripts an diesem Speicherort erstellt. Die ursprüngliche Version befindet sich weiterhin auf Ihrem OneDrive oder einem anderen SharePoint-Speicherort.

> [!IMPORTANT]
> Skripts mit [externen Aufrufen](../develop/external-calls.md) können von SharePoint aus nicht ausgeführt werden. Sie erhalten die Fehlermeldung "Netzwerkzugriffsaufrufe werden derzeit für Skripts, die auf einer SharePoint-Website gespeichert sind, nicht unterstützt".

> [!IMPORTANT]
> Power Automate unterstützt derzeit **keine** Skripts, die in SharePoint gespeichert sind.

## <a name="restore-deleted-scripts"></a>Wiederherstellen gelöschter Skripts

Wenn Sie ein Skript in Excel löschen, wird es in Ihren OneDrive- oder SharePoint-Papierkorb verschoben. Führen Sie zum Wiederherstellen eines gelöschten Skripts die unter ["Wiederherstellen fehlender, gelöschter oder beschädigter Elemente in SharePoint und OneDrive für Arbeit oder Schule](https://support.microsoft.com/office/how-to-recover-missing-deleted-or-corrupted-items-in-sharepoint-and-onedrive-for-work-or-school-3d748edf-c072-46c9-81a4-4989056ebc87)" aufgeführten Schritte aus. Durch das Wiederherstellen einer **OSTS-Datei** wird sie in die Liste **"Alle Skripts** " zurückgegeben.

Ein gelöschtes Skript wird mit der Arbeitsmappe nicht freigegeben. Wenn Sie ein Skript wiederherstellen, behält es **seinen** Skriptzugriff nicht bei. Sie müssen das Skript erneut freigeben.

Wiederhergestellte Skripts funktionieren mit Power Automate-Flüssen weiterhin erwartungsgemäß. Sie müssen den Flussverbinder nicht neu erstellen.

## <a name="file-ownership-and-retention"></a>Dateibesitz und -aufbewahrung

Office-Skripts folgen den von Microsoft OneDrive und Microsoft SharePoint angegebenen Aufbewahrungs- und Löschrichtlinien. Informationen zum Behandeln von Skripts, die von einem Benutzer erstellt und freigegeben wurden, der aus Ihrer Organisation entfernt wurde, finden [Sie unter "Informationen zur Aufbewahrung für SharePoint und OneDrive](/microsoft-365/compliance/retention-policies-sharepoint?view=o365-worldwide&preserve-view=true)".

Während der Bearbeitung werden Dateien vorübergehend im Browser gespeichert. Sie müssen das Skript speichern, bevor Sie das Excel-Fenster schließen, um es am OneDrive-Speicherort zu speichern. Vergessen Sie nicht, die Datei nach Bearbeitungen zu speichern. Andernfalls befinden sich diese Bearbeitungen nur in der Browserversion der Datei.

## <a name="audit-office-scripts-usage-at-the-admin-level"></a>Überwachen der Verwendung von Office-Skripts auf Administratorebene

Ermitteln Sie mithilfe des Compliance Center-Überwachungsprotokolls, wer Office-Skripts in Ihrer Organisation verwendet. Details zum Überwachungsprotokoll finden Sie unter [Durchsuchen des Überwachungsprotokolls im Security & Compliance Center](/microsoft-365/compliance/search-the-audit-log-in-security-and-compliance?view=o365-worldwide&preserve-view=true#search-the-audit-log).

Führen Sie die folgenden Schritte aus, um die Aktivitäten im Zusammenhang mit Office-Skripts als Administrator speziell zu überwachen.

1. Öffnen Sie in einem InPrivate-Browserfenster (oder Inkognito oder einem anderen browserspezifischen Modus mit eingeschränkter Nachverfolgung) das [Compliance Center](https://compliance.microsoft.com/), und melden Sie sich an.
1. Wechseln Sie zur Seite **"Überwachung** ".
1. *(Nur einmal)* Wählen Sie auf der Registerkarte **"Suchen** " die Option **"Aufzeichnung von Benutzer- und Administratoraktivitäten starten**" aus.

    > [!IMPORTANT]
    > Es kann ein bis zwei Stunden nach dem Aktivieren der Aufzeichnung dauern, bis alle Aktivitäten im gesamten Mandanten aufgezeichnet werden.

1. Legen Sie die gewünschten Suchoptionen fest, und drücken **Sie die Suche**. Filtern Sie **Aktivitäten** nach **dem Skript "Ausführen" in der Arbeitsmappe** , um jedes Mal zu sehen, wenn ein Skript ausgeführt wurde. Sie können auch das **Datei-, Ordner- oder Websitefeld** nach `.osts`filtern. Dies zeigt, wer in Ihrer Organisation Skripts erstellt oder ändert.

    :::image type="content" source="../images/audit-log-example.png" alt-text="Einige Zeilen mit Überwachungsprotokollsuchergebnissen, einschließlich der Aktion &quot;Skript für Arbeitsmappe ausgeführt&quot; und dem Hochladen und Ändern einer OSTS-Datei.":::

## <a name="see-also"></a>Siehe auch

- [Office-Skripts in Excel für das Web](https://support.microsoft.com/office/226eddbc-3a44-4540-acfe-fccda3d1122b)
- [Behandeln von Problemen mit Office-Skripts](../testing/troubleshooting.md)
- [Office-Skripts-Einstellungen in M365](/microsoft-365/admin/manage/manage-office-scripts-settings)
- [Auswirkungen von Office-Skripts rückgängig machen](../testing/undo.md)
