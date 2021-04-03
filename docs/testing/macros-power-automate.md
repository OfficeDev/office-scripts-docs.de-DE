---
title: Verwenden von Makrodateien in Power Automate-Flüssen
description: Erfahren Sie, wie Sie Makrodateien oder xlsm-Dateien in Power Automate-Flüssen verwenden.
ms.date: 03/18/2021
localization_priority: Normal
ms.openlocfilehash: ec1fe00eb9ddc382ae4bc02187de7a36c97288b1
ms.sourcegitcommit: 5d24e77df70aa2c1c982275d53213c2a9323ff86
ms.translationtype: MT
ms.contentlocale: de-DE
ms.lasthandoff: 04/02/2021
ms.locfileid: "51571478"
---
# <a name="how-to-use-macro-files-in-power-automate-flows"></a>Verwenden von Makrodateien in Power Automate-Flüssen

[Power Automate-Flüsse](https://flow.microsoft.com/) stellen [Excel-Connectors](https://flow.microsoft.com/connectors/shared_excelonlinebusiness/excel-online-business/) zur Verfügung, mit deren Hilfe Sie Excel-Dateien mit den restlichen Organisationsdaten und Apps wie Teams, Outlook und SharePoint verbinden können.

Makrodateien können jedoch nicht im Dateidropdown ausgewählt werden (siehe ein Beispiel im folgenden Screenshot).

![Keine xlsm in Ausführen der Skriptaktion](../images/no-xlsm.png)

Eine Möglichkeit, dieses Problem zu beheben, besteht in der Einbeziehung der Aktion "Dateimetadaten erhalten" (OneDrive oder SharePoint) und verwenden Sie die ID-Eigenschaft in der Aktion "Skript ausführen", wie im folgenden Screenshot gezeigt.

![xlsm in Ausführen der Skriptaktion](../images/xlsm-in-pa.png)

> [!NOTE]
> Einige XLSM (insbesondere diejenigen mit ActiveX/Formularsteuerelementen) funktionieren möglicherweise nicht im Excel-Onlineconnector. Testen Sie unbedingt, bevor Sie Ihre Lösung bereitstellen.

[![Video zur Verwendung von XLSM in Ausführen der Skriptaktion ansehen](../images/xlsm-vid.png)](https://youtu.be/o-H9BbywJQQ "Video zur Verwendung von XLSM in der Aktion Skript ausführen")
