---
title: Verwenden von Makrodateien in Power Automate Flüssen
description: Erfahren Sie, wie Sie Makrodateien oder XLSM-Dateien in Power Automate Flüssen verwenden.
ms.date: 09/01/2021
localization_priority: Normal
ms.openlocfilehash: 204eb8315f90c0ab0e20a34ec64517fbfbf304b1
ms.sourcegitcommit: 9472e78eb186ceffdaaceb2718d5074ce55a0e74
ms.translationtype: MT
ms.contentlocale: de-DE
ms.lasthandoff: 09/02/2021
ms.locfileid: "58866539"
---
# <a name="how-to-use-macro-files-in-power-automate-flows"></a>Verwenden von Makrodateien in Power Automate Flüssen

Der [Excel Online-Connector (Business)](https://flow.microsoft.com/connectors/shared_excelonlinebusiness/excel-online-business/) in [Power Automate](https://flow.microsoft.com/) funktioniert in der Regel nur mit Dateien im Microsoft Excel Open XML-Tabellenkalkulationsformat (.xlsx). Der Dateibrowser beschränkt Ihre Auswahl auf .xlsx Dateien innerhalb des Connectors. Makrodateien sind jedoch mit der **Skriptaktion "Ausführen"** des Connectors kompatibel, wenn die Dateimetadaten verwendet werden.

Verwenden Sie in Ihrem Flow die Aktion **"Dateimetadaten** abrufen" aus dem [OneDrive for Business](https://flow.microsoft.com/connectors/shared_onedriveforbusiness/onedrive-for-business/) oder [SharePoint](https://flow.microsoft.com/connectors/shared_sharepointonline/sharepoint/) Connectors. Die **Skriptaktion ausführen** akzeptiert diese Metadaten als gültige Datei. Verwenden Sie beim Ausführen des Skripts den dynamischen *ID-Inhalt,* der von der Aktion **"Dateimetadaten abrufen"** zurückgegeben wird, als Argument "Datei". Der folgende Screenshot zeigt einen Fluss, der die Metadaten für eine Datei namens "Makro File.xlsm testen" für eine **Ausführungsskriptaktion** bereitstellt.

:::image type="content" source="../images/xlsm-in-power-automate.png" alt-text="Ein Fluss mit einer Aktion zum Abrufen von Dateimetadaten, der die Metadaten einer Makrodatei an eine Ausführungsskriptaktion übergibt.":::

> [!WARNING]
> Einige XLSM-Dateien, insbesondere solche mit ActiveX oder Formularsteuerelementen, funktionieren möglicherweise nicht im Excel Onlineconnector. Testen Sie die Lösung unbedingt, bevor Sie sie bereitstellen.

## <a name="other-resources"></a>Sonstige Ressourcen

[Sehen Sie sich das YouTube-Video von Adrianhi Ramamurthy über die Verwendung einer XLSM-Datei in einer Ausführungsskriptaktion](https://youtu.be/o-H9BbywJQQ)an.
