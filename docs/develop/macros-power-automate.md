---
title: Verwenden von Makros aktivierten Dateien in Power Automate Flüssen
description: Erfahren Sie, wie Sie makrofähige Dateien oder XLSM-Dateien in Power Automate Flüssen verwenden.
ms.date: 03/24/2022
ms.localizationpriority: medium
ms.openlocfilehash: 9f2ecefe9fb97d1c5514ddb52c3cbcd0596df426
ms.sourcegitcommit: 7023b9e23499806901a5ecf8ebc460b76887cca6
ms.translationtype: MT
ms.contentlocale: de-DE
ms.lasthandoff: 03/31/2022
ms.locfileid: "64585744"
---
# <a name="how-to-use-macro-enabled-files-in-power-automate-flows"></a>Verwenden von Makros aktivierten Dateien in Power Automate Flüssen

Sie können Ihre XLSM-Dateien in einen Power Automate Fluss integrieren. Auf diese Weise können Sie ihre vorhandenen Automatisierungslösungen in webbasierte Formate konvertieren. Beachten Sie, dass die in den XSLM-Dateien enthaltenen Makros nicht über Power Automate ausgeführt werden können. Nur Office Skripts sind dort aktiviert.

Der [Excel Online -Connector (Business)](https://flow.microsoft.com/connectors/shared_excelonlinebusiness/excel-online-business/) in [Power Automate](https://flow.microsoft.com/) ist in der Regel auf Dateien im Microsoft Excel Open XML-Tabellenkalkulationsformat (.xlsx) beschränkt. Im Dateibrowser können Sie nur .xlsx Dateien auswählen. Makrofähige Dateien sind jedoch mit der **Skriptaktion "Ausführen** " des Connectors kompatibel, wenn die Dateimetadaten verwendet werden.

Verwenden Sie in Ihrem Flow die Aktion "**Dateimetadaten** abrufen" aus dem [OneDrive for Business](https://flow.microsoft.com/connectors/shared_onedriveforbusiness/onedrive-for-business/) oder [SharePoint](https://flow.microsoft.com/connectors/shared_sharepointonline/sharepoint/) Connectors. Die **Skriptaktion ausführen** akzeptiert diese Metadaten als gültige Datei. Verwenden Sie beim Ausführen des Skripts den dynamischen *ID-Inhalt* , der von der Aktion " **Dateimetadaten abrufen** " zurückgegeben wird, als Argument "Datei". Der folgende Screenshot zeigt einen Fluss, der die Metadaten für eine Datei namens "Test Macro File.xlsm" für eine **Ausführungsskriptaktion** bereitstellt.

:::image type="content" source="../images/xlsm-in-power-automate.png" alt-text="Ein Fluss mit einer Aktion zum Abrufen von Dateimetadaten, der die Metadaten einer Makrodatei an eine Ausführungsskriptaktion übergibt.":::

> [!WARNING]
> Einige XLSM-Dateien, insbesondere solche mit ActiveX oder Formularsteuerelementen, funktionieren möglicherweise nicht im Excel Onlineconnector. Testen Sie die Lösung unbedingt, bevor Sie sie bereitstellen.

## <a name="other-resources"></a>Weitere Ressourcen

[Sehen Sie sich das YouTube-Video von Adrianhi Ramamurthy zur Verwendung einer XLSM-Datei in einer Ausführungsskriptaktion an](https://youtu.be/o-H9BbywJQQ).
