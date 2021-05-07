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
# <a name="how-to-use-macro-files-in-power-automate-flows"></a><span data-ttu-id="051e1-103">Verwenden von Makrodateien in Power Automate Flüssen</span><span class="sxs-lookup"><span data-stu-id="051e1-103">How to use macro files in Power Automate flows</span></span>

<span data-ttu-id="051e1-104">[Power Automate-Flüsse](https://flow.microsoft.com/) stellen Excel [Connectors](https://flow.microsoft.com/connectors/shared_excelonlinebusiness/excel-online-business/) zur Verfügung, mit deren Hilfe Excel-Dateien mit den restlichen Organisationsdaten und Apps wie Teams, Outlook und SharePoint.</span><span class="sxs-lookup"><span data-stu-id="051e1-104">[Power Automate flows](https://flow.microsoft.com/) provide [Excel connectors](https://flow.microsoft.com/connectors/shared_excelonlinebusiness/excel-online-business/) to help connect Excel files with the rest of your organizational data and apps such as Teams, Outlook, and SharePoint.</span></span>

<span data-ttu-id="051e1-105">Makrodateien können jedoch nicht im Dateidropdown ausgewählt werden (siehe ein Beispiel im folgenden Screenshot).</span><span class="sxs-lookup"><span data-stu-id="051e1-105">However, macro files can't be selected in the file dropdown (see an example in the following screenshot).</span></span>

:::image type="content" source="../images/no-xlsm.png" alt-text="Die Power Automate Skriptaktion ausführen, bei der keine Makrodatei ausgewählt ist. Der angezeigte Fehler ist &quot;Datei&quot; erforderlich.":::

<span data-ttu-id="051e1-107">Eine Möglichkeit, dieses Problem zu beheben, besteht in der Einbeziehung der Aktion "Dateimetadaten" (OneDrive oder SharePoint) und verwenden Sie die ID-Eigenschaft in der Aktion "Skript ausführen", wie im folgenden Screenshot gezeigt.</span><span class="sxs-lookup"><span data-stu-id="051e1-107">One way to get around this issue is by including the "Get File Metadata" action (OneDrive or SharePoint) and use the ID property in the "Run Script" action as shown in the following screenshot.</span></span>

:::image type="content" source="../images/xlsm-in-pa.png" alt-text="Die Power Automate Ausführen der Skriptaktion mit der ausgewählten Makrodatei und keine Ausführungsskriptfehler":::

> [!NOTE]
> <span data-ttu-id="051e1-109">Einige XLSM (insbesondere diejenigen mit ActiveX/Formularsteuerelementen) funktionieren möglicherweise nicht im Excel Onlineconnector.</span><span class="sxs-lookup"><span data-stu-id="051e1-109">Some XLSM (especially the ones with ActiveX/Form controls) may not work in the Excel online connector.</span></span> <span data-ttu-id="051e1-110">Testen Sie unbedingt, bevor Sie Ihre Lösung bereitstellen.</span><span class="sxs-lookup"><span data-stu-id="051e1-110">Be sure to test before deploying your solution.</span></span>

## <a name="other-resources"></a><span data-ttu-id="051e1-111">Sonstige Ressourcen</span><span class="sxs-lookup"><span data-stu-id="051e1-111">Other resources</span></span>

<span data-ttu-id="051e1-112">[Sehen Sie sich das YouTube-Video von Sudhi Ramamurthy](https://youtu.be/o-H9BbywJQQ)über die Verwendung einer XLSM-Datei in einer Run Script-Aktion an.</span><span class="sxs-lookup"><span data-stu-id="051e1-112">[Watch Sudhi Ramamurthy's YouTube video on how use an .xlsm file in a Run Script action](https://youtu.be/o-H9BbywJQQ).</span></span>
