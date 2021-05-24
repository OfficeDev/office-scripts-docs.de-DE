---
title: Hinzufügen von Bildern zu einer Arbeitsmappe
description: Erfahren Sie, wie Sie Office Skripts verwenden, um einer Arbeitsmappe ein Bild hinzuzufügen und es über Blätter hinweg zu kopieren.
ms.date: 05/17/2021
localization_priority: Normal
ms.openlocfilehash: 64c356b2d76a561276b2955263555b16de27b3ba
ms.sourcegitcommit: a2b85168d2b5e2c4e6951f808368f7d726400df0
ms.translationtype: MT
ms.contentlocale: de-DE
ms.lasthandoff: 05/21/2021
ms.locfileid: "52592754"
---
# <a name="add-images-to-a-workbook"></a><span data-ttu-id="83d55-103">Hinzufügen von Bildern zu einer Arbeitsmappe</span><span class="sxs-lookup"><span data-stu-id="83d55-103">Add images to a workbook</span></span>

<span data-ttu-id="83d55-104">In diesem Beispiel wird gezeigt, wie Sie mit Bildern mithilfe eines Office Skripts in Excel.</span><span class="sxs-lookup"><span data-stu-id="83d55-104">This sample shows how to work with images using an Office Script in Excel.</span></span>

## <a name="scenario"></a><span data-ttu-id="83d55-105">Szenario</span><span class="sxs-lookup"><span data-stu-id="83d55-105">Scenario</span></span>

<span data-ttu-id="83d55-106">Bilder helfen bei Branding, visueller Identität und Vorlagen.</span><span class="sxs-lookup"><span data-stu-id="83d55-106">Images help with branding, visual identity, and templates.</span></span> <span data-ttu-id="83d55-107">Sie helfen, eine Arbeitsmappe mehr als nur einen riesigen Tisch zu erstellen.</span><span class="sxs-lookup"><span data-stu-id="83d55-107">They help make a workbook more than just a giant table.</span></span>

<span data-ttu-id="83d55-108">Im ersten Beispiel wird ein Bild von einem Arbeitsblatt in ein anderes kopiert.</span><span class="sxs-lookup"><span data-stu-id="83d55-108">The first sample copies an image from one worksheet to another.</span></span> <span data-ttu-id="83d55-109">Dies kann verwendet werden, um das Logo Ihres Unternehmens auf jedem Blatt an die gleiche Position zu setzen.</span><span class="sxs-lookup"><span data-stu-id="83d55-109">This could be used to put your company's logo in the same position on every sheet.</span></span>

<span data-ttu-id="83d55-110">Im zweiten Beispiel wird ein Bild aus einer URL kopiert.</span><span class="sxs-lookup"><span data-stu-id="83d55-110">The second sample copies an image from a URL.</span></span> <span data-ttu-id="83d55-111">Dies kann verwendet werden, um Fotos, die ein Kollege in einem freigegebenen Ordner gespeichert hat, in eine zugehörige Arbeitsmappe zu kopieren.</span><span class="sxs-lookup"><span data-stu-id="83d55-111">This could be used to copy photos that a colleague stored in a shared folder to a related workbook.</span></span>

## <a name="sample-excel-file"></a><span data-ttu-id="83d55-112">Beispieldatei Excel Datei</span><span class="sxs-lookup"><span data-stu-id="83d55-112">Sample Excel file</span></span>

<span data-ttu-id="83d55-113">Laden Sie die Datei <a href="add-images.xlsx">add-images.xlsx, </a> die in diesen Beispielen verwendet wird, herunter, und testen Sie sie selbst!</span><span class="sxs-lookup"><span data-stu-id="83d55-113">Download the file <a href="add-images.xlsx">add-images.xlsx</a> used in these samples and try it out yourself!</span></span>

## <a name="sample-code-copy-an-image-across-worksheets"></a><span data-ttu-id="83d55-114">Beispielcode: Kopieren eines Bilds über Arbeitsblätter hinweg</span><span class="sxs-lookup"><span data-stu-id="83d55-114">Sample code: Copy an image across worksheets</span></span>

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

## <a name="sample-code-add-an-image-from-a-url-to-a-workbook"></a><span data-ttu-id="83d55-115">Beispielcode: Hinzufügen eines Bilds aus einer URL zu einer Arbeitsmappe</span><span class="sxs-lookup"><span data-stu-id="83d55-115">Sample code: Add an image from a URL to a workbook</span></span>

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
  workbook.getWorksheet("WebSheet").addImage(image)
}

/**
 * Converts an ArrayBuffer containing a .png image into a base64-encoded string.
 */
function convertToBase64(input: ArrayBuffer) {
  const uInt8Array = new Uint8Array(input);
  const count = uInt8Array.length;

  // Allocate the necessary space up front.
  const charCodeArray = new Array(count) 
  
  // Convert every entry in the array to a character.
  for (let i = count; i >= 0; i--) { 
    charCodeArray[i] = String.fromCharCode(uInt8Array[i]);
  }

  // Convert the characters to base64.
  const base64 = btoa(charCodeArray.join(''));
  return base64;
}
```
