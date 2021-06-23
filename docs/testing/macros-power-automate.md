---
title: Verwenden von Makrodateien in Power Automate Flüssen
description: Erfahren Sie, wie Sie Makrodateien oder XLSM-Dateien in Power Automate Flüssen verwenden.
ms.date: 04/28/2021
localization_priority: Normal
ms.openlocfilehash: 91e11424e4220a3e1f80cdd2711d05f219016147
ms.sourcegitcommit: 4693c8f79428ec74695328275703af0ba1bfea8f
ms.translationtype: MT
ms.contentlocale: de-DE
ms.lasthandoff: 06/23/2021
ms.locfileid: "53074641"
---
# <a name="how-to-use-macro-files-in-power-automate-flows"></a>Verwenden von Makrodateien in Power Automate Flüssen

[Power Automate Flüsse](https://flow.microsoft.com/) bieten [Excel Connectors,](https://flow.microsoft.com/connectors/shared_excelonlinebusiness/excel-online-business/) um Excel Dateien mit den restlichen Organisationsdaten und Apps wie Teams, Outlook und SharePoint zu verbinden.

Makrodateien können jedoch nicht im Dateidropdown ausgewählt werden (siehe Beispiel im folgenden Screenshot).

:::image type="content" source="../images/no-xlsm.png" alt-text="Die Skriptaktion Power Automate Ausführen, bei der keine Makrodatei ausgewählt ist. Der angezeigte Fehler lautet &quot;Datei&quot; ist erforderlich.":::

Eine Möglichkeit, dieses Problem zu umgehen, besteht darin, die Aktion "Dateimetadaten abrufen" (OneDrive oder SharePoint) einzublenden und die ID-Eigenschaft in der Aktion "Skript ausführen" zu verwenden, wie im folgenden Screenshot dargestellt.

:::image type="content" source="../images/xlsm-in-pa.png" alt-text="Die Aktion Power Automate Skript ausführen zeigt die ausgewählte Makrodatei und keinen Fehler beim Ausführen des Skripts an.":::

> [!NOTE]
> Einige XLSM (insbesondere die mit ActiveX/Formularsteuerelementen) funktionieren möglicherweise nicht im Excel Onlineconnector. Testen Sie die Lösung unbedingt, bevor Sie sie bereitstellen.

## <a name="other-resources"></a>Sonstige Ressourcen

[Sehen Sie sich das YouTube-Video von Adrianhi Ramamurthy über die Verwendung einer XLSM-Datei in einer Ausführungsskriptaktion](https://youtu.be/o-H9BbywJQQ)an.
