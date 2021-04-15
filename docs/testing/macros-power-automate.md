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
# <a name="how-to-use-macro-files-in-power-automate-flows"></a><span data-ttu-id="db033-103">Verwenden von Makrodateien in Power Automate-Flüssen</span><span class="sxs-lookup"><span data-stu-id="db033-103">How to use macro files in Power Automate flows</span></span>

<span data-ttu-id="db033-104">[Power Automate-Flüsse](https://flow.microsoft.com/) stellen [Excel-Connectors](https://flow.microsoft.com/connectors/shared_excelonlinebusiness/excel-online-business/) zur Verfügung, mit deren Hilfe Sie Excel-Dateien mit den restlichen Organisationsdaten und Apps wie Teams, Outlook und SharePoint verbinden können.</span><span class="sxs-lookup"><span data-stu-id="db033-104">[Power Automate flows](https://flow.microsoft.com/) provide [Excel connectors](https://flow.microsoft.com/connectors/shared_excelonlinebusiness/excel-online-business/) to help connect Excel files with the rest of your organizational data and apps such as Teams, Outlook, and SharePoint.</span></span>

<span data-ttu-id="db033-105">Makrodateien können jedoch nicht im Dateidropdown ausgewählt werden (siehe ein Beispiel im folgenden Screenshot).</span><span class="sxs-lookup"><span data-stu-id="db033-105">However, macro files can't be selected in the file dropdown (see an example in the following screenshot).</span></span>

:::image type="content" source="../images/no-xlsm.png" alt-text="Die Power Automate Run-Skriptaktion, in der keine Makrodatei ausgewählt ist. Der angezeigte Fehler ist &quot;Datei&quot; erforderlich.":::

<span data-ttu-id="db033-107">Eine Möglichkeit, dieses Problem zu beheben, besteht in der Einbeziehung der Aktion "Dateimetadaten erhalten" (OneDrive oder SharePoint) und verwenden Sie die ID-Eigenschaft in der Aktion "Skript ausführen", wie im folgenden Screenshot gezeigt.</span><span class="sxs-lookup"><span data-stu-id="db033-107">One way to get around this issue is by including the "Get File Metadata" action (OneDrive or SharePoint) and use the ID property in the "Run Script" action as shown in the following screenshot.</span></span>

:::image type="content" source="../images/xlsm-in-pa.png" alt-text="Die Power Automate Run-Skriptaktion mit der ausgewählten Makrodatei und keine Ausführungsskriptfehler.":::

> [!NOTE]
> <span data-ttu-id="db033-109">Einige XLSM (insbesondere diejenigen mit ActiveX/Formularsteuerelementen) funktionieren möglicherweise nicht im Excel-Onlineconnector.</span><span class="sxs-lookup"><span data-stu-id="db033-109">Some XLSM (especially the ones with ActiveX/Form controls) may not work in the Excel online connector.</span></span> <span data-ttu-id="db033-110">Testen Sie unbedingt, bevor Sie Ihre Lösung bereitstellen.</span><span class="sxs-lookup"><span data-stu-id="db033-110">Be sure to test before deploying your solution.</span></span>

<span data-ttu-id="db033-111">[![Video zur Verwendung von XLSM in Ausführen der Skriptaktion ansehen](../images/xlsm-vid.png)](https://youtu.be/o-H9BbywJQQ "Video zur Verwendung von XLSM in der Aktion Skript ausführen")</span><span class="sxs-lookup"><span data-stu-id="db033-111">[![Watch video about using XLSM in Run Script action](../images/xlsm-vid.png)](https://youtu.be/o-H9BbywJQQ "Video about using XLSM in Run Script action")</span></span>
