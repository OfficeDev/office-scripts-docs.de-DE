---
title: Verwenden von Makrodateien in Power Automate Flüssen
description: Erfahren Sie, wie Sie Makrodateien oder xlsm-Dateien in Power Automate verwenden.
ms.date: 04/28/2021
localization_priority: Normal
ms.openlocfilehash: b232a1d31a7ff6e28016c5e28fd8a83c8d3f1859
ms.sourcegitcommit: f7a7aebfb687f2a35dbed07ed62ff352a114525a
ms.translationtype: MT
ms.contentlocale: de-DE
ms.lasthandoff: 05/06/2021
ms.locfileid: "52232655"
---
# <a name="how-to-use-macro-files-in-power-automate-flows"></a>Verwenden von Makrodateien in Power Automate Flüssen

[Power Automate-Flüsse](https://flow.microsoft.com/) stellen Excel [Connectors](https://flow.microsoft.com/connectors/shared_excelonlinebusiness/excel-online-business/) zur Verfügung, mit deren Hilfe Excel-Dateien mit den restlichen Organisationsdaten und Apps wie Teams, Outlook und SharePoint.

Makrodateien können jedoch nicht im Dateidropdown ausgewählt werden (siehe ein Beispiel im folgenden Screenshot).

:::image type="content" source="../images/no-xlsm.png" alt-text="Die Power Automate Skriptaktion ausführen, bei der keine Makrodatei ausgewählt ist. Der angezeigte Fehler ist &quot;Datei&quot; erforderlich.":::

Eine Möglichkeit, dieses Problem zu beheben, besteht in der Einbeziehung der Aktion "Dateimetadaten" (OneDrive oder SharePoint) und verwenden Sie die ID-Eigenschaft in der Aktion "Skript ausführen", wie im folgenden Screenshot gezeigt.

:::image type="content" source="../images/xlsm-in-pa.png" alt-text="Die Power Automate Ausführen der Skriptaktion mit der ausgewählten Makrodatei und keine Ausführungsskriptfehler":::

> [!NOTE]
> Einige XLSM (insbesondere diejenigen mit ActiveX/Formularsteuerelementen) funktionieren möglicherweise nicht im Excel Onlineconnector. Testen Sie unbedingt, bevor Sie Ihre Lösung bereitstellen.

## <a name="other-resources"></a>Sonstige Ressourcen

[Sehen Sie sich das YouTube-Video von Sudhi Ramamurthy](https://youtu.be/o-H9BbywJQQ)über die Verwendung einer XLSM-Datei in einer Run Script-Aktion an.
