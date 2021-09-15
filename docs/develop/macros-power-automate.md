---
title: Verwenden von Makrodateien in Power Automate-Flows
description: Erfahren Sie, wie Sie Makrodateien oder XLSM-Dateien in Power Automate Flüssen verwenden.
ms.date: 09/01/2021
ms.localizationpriority: medium
ms.openlocfilehash: ab83c62d219ec215497e02d6cfe5718c628ec1bf
ms.sourcegitcommit: d3ed4bdeeba805d97c930394e172e8306a0cf484
ms.translationtype: MT
ms.contentlocale: de-DE
ms.lasthandoff: 09/15/2021
ms.locfileid: "59326905"
---
# <a name="how-to-use-macro-files-in-power-automate-flows"></a>Verwenden von Makrodateien in Power Automate Flüssen

Der [Excel Online-Connector (Business)](https://flow.microsoft.com/connectors/shared_excelonlinebusiness/excel-online-business/) in [Power Automate](https://flow.microsoft.com/) funktioniert in der Regel nur mit Dateien im Microsoft Excel Open XML-Tabellenkalkulationsformat (.xlsx). Der Dateibrowser beschränkt Ihre Auswahl auf .xlsx Dateien innerhalb des Connectors. Makrodateien sind jedoch mit der **Skriptaktion "Ausführen"** des Connectors kompatibel, wenn die Dateimetadaten verwendet werden.

Verwenden Sie in Ihrem Flow die Aktion **"Dateimetadaten** abrufen" aus dem [OneDrive for Business](https://flow.microsoft.com/connectors/shared_onedriveforbusiness/onedrive-for-business/) oder [SharePoint](https://flow.microsoft.com/connectors/shared_sharepointonline/sharepoint/) Connectors. Die **Skriptaktion ausführen** akzeptiert diese Metadaten als gültige Datei. Verwenden Sie beim Ausführen des Skripts den dynamischen *ID-Inhalt,* der von der Aktion **"Dateimetadaten abrufen"** zurückgegeben wird, als Argument "Datei". Der folgende Screenshot zeigt einen Fluss, der die Metadaten für eine Datei namens "Test Macro File.xlsm" für eine **Ausführungsskriptaktion** bereitstellt.

:::image type="content" source="../images/xlsm-in-power-automate.png" alt-text="Ein Fluss mit einer Aktion zum Abrufen von Dateimetadaten, der die Metadaten einer Makrodatei an eine Ausführungsskriptaktion übergibt.":::

> [!WARNING]
> Einige XLSM-Dateien, insbesondere solche mit ActiveX oder Formularsteuerelementen, funktionieren möglicherweise nicht im Excel Onlineconnector. Testen Sie die Lösung unbedingt, bevor Sie sie bereitstellen.

## <a name="other-resources"></a>Weitere Ressourcen

[Sehen Sie sich das YouTube-Video von Adrianhi Ramamurthy über die Verwendung einer XLSM-Datei in einer Ausführungsskriptaktion](https://youtu.be/o-H9BbywJQQ)an.
