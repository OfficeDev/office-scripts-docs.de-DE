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
# <a name="how-to-use-macro-files-in-power-automate-flows"></a><span data-ttu-id="dc435-103">Verwenden von Makrodateien in Power Automate Flüssen</span><span class="sxs-lookup"><span data-stu-id="dc435-103">How to use macro files in Power Automate flows</span></span>

<span data-ttu-id="dc435-104">[Power Automate Flüsse](https://flow.microsoft.com/) bieten [Excel Connectors,](https://flow.microsoft.com/connectors/shared_excelonlinebusiness/excel-online-business/) um Excel Dateien mit den restlichen Organisationsdaten und Apps wie Teams, Outlook und SharePoint zu verbinden.</span><span class="sxs-lookup"><span data-stu-id="dc435-104">[Power Automate flows](https://flow.microsoft.com/) provide [Excel connectors](https://flow.microsoft.com/connectors/shared_excelonlinebusiness/excel-online-business/) to help connect Excel files with the rest of your organizational data and apps such as Teams, Outlook, and SharePoint.</span></span>

<span data-ttu-id="dc435-105">Makrodateien können jedoch nicht im Dateidropdown ausgewählt werden (siehe Beispiel im folgenden Screenshot).</span><span class="sxs-lookup"><span data-stu-id="dc435-105">However, macro files can't be selected in the file dropdown (see an example in the following screenshot).</span></span>

:::image type="content" source="../images/no-xlsm.png" alt-text="Die Skriptaktion Power Automate Ausführen, bei der keine Makrodatei ausgewählt ist. Der angezeigte Fehler lautet &quot;Datei&quot; ist erforderlich.":::

<span data-ttu-id="dc435-107">Eine Möglichkeit, dieses Problem zu umgehen, besteht darin, die Aktion "Dateimetadaten abrufen" (OneDrive oder SharePoint) einzublenden und die ID-Eigenschaft in der Aktion "Skript ausführen" zu verwenden, wie im folgenden Screenshot dargestellt.</span><span class="sxs-lookup"><span data-stu-id="dc435-107">One way to get around this issue is by including the "Get File Metadata" action (OneDrive or SharePoint) and use the ID property in the "Run Script" action as shown in the following screenshot.</span></span>

:::image type="content" source="../images/xlsm-in-pa.png" alt-text="Die Aktion Power Automate Skript ausführen zeigt die ausgewählte Makrodatei und keinen Fehler beim Ausführen des Skripts an.":::

> [!NOTE]
> <span data-ttu-id="dc435-109">Einige XLSM (insbesondere die mit ActiveX/Formularsteuerelementen) funktionieren möglicherweise nicht im Excel Onlineconnector.</span><span class="sxs-lookup"><span data-stu-id="dc435-109">Some XLSM (especially the ones with ActiveX/Form controls) may not work in the Excel online connector.</span></span> <span data-ttu-id="dc435-110">Testen Sie die Lösung unbedingt, bevor Sie sie bereitstellen.</span><span class="sxs-lookup"><span data-stu-id="dc435-110">Be sure to test before deploying your solution.</span></span>

## <a name="other-resources"></a><span data-ttu-id="dc435-111">Sonstige Ressourcen</span><span class="sxs-lookup"><span data-stu-id="dc435-111">Other resources</span></span>

<span data-ttu-id="dc435-112">[Sehen Sie sich das YouTube-Video von Adrianhi Ramamurthy über die Verwendung einer XLSM-Datei in einer Ausführungsskriptaktion](https://youtu.be/o-H9BbywJQQ)an.</span><span class="sxs-lookup"><span data-stu-id="dc435-112">[Watch Sudhi Ramamurthy's YouTube video on how use an .xlsm file in a Run Script action](https://youtu.be/o-H9BbywJQQ).</span></span>
