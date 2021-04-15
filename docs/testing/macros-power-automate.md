---
title: Verwenden von Makrodateien in Power Automate-Flüssen
description: Erfahren Sie, wie Sie Makrodateien oder xlsm-Dateien in Power Automate-Flüssen verwenden.
ms.date: 03/18/2021
localization_priority: Normal
ms.openlocfilehash: a7929fc485ae2118d30a4f2783538d0e04deca2a
ms.sourcegitcommit: 45ffe3dbd2c834b78592ad35928cf8096f5e80bc
ms.translationtype: MT
ms.contentlocale: de-DE
ms.lasthandoff: 04/14/2021
ms.locfileid: "51755014"
---
# <a name="how-to-use-macro-files-in-power-automate-flows"></a>Verwenden von Makrodateien in Power Automate-Flüssen

[Power Automate-Flüsse](https://flow.microsoft.com/) stellen [Excel-Connectors](https://flow.microsoft.com/connectors/shared_excelonlinebusiness/excel-online-business/) zur Verfügung, mit deren Hilfe Sie Excel-Dateien mit den restlichen Organisationsdaten und Apps wie Teams, Outlook und SharePoint verbinden können.

Makrodateien können jedoch nicht im Dateidropdown ausgewählt werden (siehe ein Beispiel im folgenden Screenshot).

:::image type="content" source="../images/no-xlsm.png" alt-text="Die Power Automate Run-Skriptaktion, in der keine Makrodatei ausgewählt ist. Der angezeigte Fehler ist &quot;Datei&quot; erforderlich.":::

Eine Möglichkeit, dieses Problem zu beheben, besteht in der Einbeziehung der Aktion "Dateimetadaten erhalten" (OneDrive oder SharePoint) und verwenden Sie die ID-Eigenschaft in der Aktion "Skript ausführen", wie im folgenden Screenshot gezeigt.

:::image type="content" source="../images/xlsm-in-pa.png" alt-text="Die Power Automate Run-Skriptaktion mit der ausgewählten Makrodatei und keine Ausführungsskriptfehler.":::

> [!NOTE]
> Einige XLSM (insbesondere diejenigen mit ActiveX/Formularsteuerelementen) funktionieren möglicherweise nicht im Excel-Onlineconnector. Testen Sie unbedingt, bevor Sie Ihre Lösung bereitstellen.

[![Video zur Verwendung von XLSM in Ausführen der Skriptaktion ansehen](../images/xlsm-vid.png)](https://youtu.be/o-H9BbywJQQ "Video zur Verwendung von XLSM in der Aktion Skript ausführen")
