---
title: Hinzufügen von Bildern zu einer Arbeitsmappe
description: Erfahren Sie, wie Sie Office Skripts verwenden, um einer Arbeitsmappe ein Bild hinzuzufügen und es blätterübergreifend zu kopieren.
ms.date: 07/12/2021
localization_priority: Normal
ms.openlocfilehash: 993444aa328356f872db90d1b9d2403bf28be4de
ms.sourcegitcommit: a86b91c7e104bb7c26efd56de53b9e3976a34828
ms.translationtype: MT
ms.contentlocale: de-DE
ms.lasthandoff: 07/12/2021
ms.locfileid: "53394416"
---
# <a name="add-images-to-a-workbook"></a><span data-ttu-id="e0c28-103">Hinzufügen von Bildern zu einer Arbeitsmappe</span><span class="sxs-lookup"><span data-stu-id="e0c28-103">Add images to a workbook</span></span>

<span data-ttu-id="e0c28-104">In diesem Beispiel wird gezeigt, wie Sie mit Bildern mithilfe eines Office Skripts in Excel arbeiten.</span><span class="sxs-lookup"><span data-stu-id="e0c28-104">This sample shows how to work with images using an Office Script in Excel.</span></span>

## <a name="scenario"></a><span data-ttu-id="e0c28-105">Szenario</span><span class="sxs-lookup"><span data-stu-id="e0c28-105">Scenario</span></span>

<span data-ttu-id="e0c28-106">Bilder helfen bei Branding, visueller Identität und Vorlagen.</span><span class="sxs-lookup"><span data-stu-id="e0c28-106">Images help with branding, visual identity, and templates.</span></span> <span data-ttu-id="e0c28-107">Sie tragen dazu bei, eine Arbeitsmappe nicht nur zu einer Tabellentabelle zu machen.</span><span class="sxs-lookup"><span data-stu-id="e0c28-107">They help make a workbook more than just a giant table.</span></span>

<span data-ttu-id="e0c28-108">Im ersten Beispiel wird ein Bild von einem Arbeitsblatt in ein anderes kopiert.</span><span class="sxs-lookup"><span data-stu-id="e0c28-108">The first sample copies an image from one worksheet to another.</span></span> <span data-ttu-id="e0c28-109">Dies kann verwendet werden, um das Logo Ihres Unternehmens an der gleichen Position auf jedem Blatt zu platzieren.</span><span class="sxs-lookup"><span data-stu-id="e0c28-109">This could be used to put your company's logo in the same position on every sheet.</span></span>

<span data-ttu-id="e0c28-110">Im zweiten Beispiel wird ein Bild von einer URL kopiert.</span><span class="sxs-lookup"><span data-stu-id="e0c28-110">The second sample copies an image from a URL.</span></span> <span data-ttu-id="e0c28-111">Dies kann verwendet werden, um Fotos, die ein Kollege in einem freigegebenen Ordner gespeichert hat, in eine verwandte Arbeitsmappe zu kopieren.</span><span class="sxs-lookup"><span data-stu-id="e0c28-111">This could be used to copy photos that a colleague stored in a shared folder to a related workbook.</span></span>

## <a name="sample-excel-file"></a><span data-ttu-id="e0c28-112">Beispieldatei für Excel</span><span class="sxs-lookup"><span data-stu-id="e0c28-112">Sample Excel file</span></span>

<span data-ttu-id="e0c28-113">Laden Sie <a href="add-images.xlsx">add-images.xlsx</a> für eine einsatzbereite Arbeitsmappe herunter.</span><span class="sxs-lookup"><span data-stu-id="e0c28-113">Download <a href="add-images.xlsx">add-images.xlsx</a> for a ready-to-use workbook.</span></span> <span data-ttu-id="e0c28-114">Fügen Sie die folgenden Skripts hinzu, und testen Sie das Beispiel selbst!</span><span class="sxs-lookup"><span data-stu-id="e0c28-114">Add the following scripts and try the sample yourself!</span></span>

## <a name="sample-code-copy-an-image-across-worksheets"></a><span data-ttu-id="e0c28-115">Beispielcode: Kopieren eines Bilds auf mehrere Arbeitsblätter</span><span class="sxs-lookup"><span data-stu-id="e0c28-115">Sample code: Copy an image across worksheets</span></span>

```TypeScript
/**
 * This script transfers an image from one worksheet to another.
 */
function main(workbook: ExcelScript.Workbook)
{
  // Get the worksheet with the image on it.
  let firstWorksheet = workbook.getWorksheet("FirstSheet");

  // Get the first image from the worksheet.
  // If a script added the image, you could add a name to make it easier to find.
  let image: ExcelScript.Image;
  firstWorksheet.getShapes().forEach((shape, index) => {
    if (shape.getType() === ExcelScript.ShapeType.image) {
      image = shape.getImage();
      return;
    }
  });

  // Copy the image to another worksheet.
  image.getShape().copyTo("SecondSheet");
}
```

## <a name="sample-code-add-an-image-from-a-url-to-a-workbook"></a><span data-ttu-id="e0c28-116">Beispielcode: Hinzufügen eines Bilds von einer URL zu einer Arbeitsmappe</span><span class="sxs-lookup"><span data-stu-id="e0c28-116">Sample code: Add an image from a URL to a workbook</span></span>

```TypeScript
async function main(workbook: ExcelScript.Workbook) {
  // Fetch the image from a URL.
  const link = "https://raw.githubusercontent.com/OfficeDev/office-scripts-docs/master/docs/images/git-octocat.png";
  const response = await fetch(link);

  // Store the response as an ArrayBuffer, since it is a raw image file.
  const data = await response.arrayBuffer();

  // Convert the image data into a base64-encoded string.
  const image = convertToBase64(data);

  // Add the image to a worksheet.
  workbook.getWorksheet("WebSheet").addImage(image);
}

/**
 * Converts an ArrayBuffer containing a .png image into a base64-encoded string.
 */
function convertToBase64(input: ArrayBuffer) {
  const uInt8Array = new Uint8Array(input);
  const count = uInt8Array.length;

  // Allocate the necessary space up front.
  const charCodeArray = new Array(count) as string[];
  
  // Convert every entry in the array to a character.
  for (let i = count; i >= 0; i--) { 
    charCodeArray[i] = String.fromCharCode(uInt8Array[i]);
  }

  // Convert the characters to base64.
  const base64 = btoa(charCodeArray.join(''));
  return base64;
}
```
