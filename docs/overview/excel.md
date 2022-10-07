---
title: Office-Skripts in Excel
description: Eine kurze Einführung in den Action Recorder und den Code Editor für Office-Skripts.
ms.topic: overview
ms.date: 10/05/2022
ms.localizationpriority: high
ms.openlocfilehash: 02a45e5aae468cff2c61e18b35c54ba656d0484b
ms.sourcegitcommit: 64d506257bee282fb01aedbf4d090781b06e4900
ms.translationtype: MT
ms.contentlocale: de-DE
ms.lasthandoff: 10/07/2022
ms.locfileid: "68495474"
---
# <a name="office-scripts-in-excel"></a>Office-Skripts in Excel

Mit Office-Skripts in Excel können Sie Ihre täglichen Aufgaben automatisieren. In Excel im Web können Sie Ihre Aktionen mit dem Aktionsrekorder aufzeichnen. Dadurch wird ein TypeScript-Sprachskript erstellt, das jederzeit erneut ausgeführt werden kann. Sie können Skripts auch mit dem Code Editor erstellen und bearbeiten. Ihre Skripts können dann für ihre gesamte Organisation freigegeben werden, sodass Ihre Kollegen den Workflow entsprechend automatisieren können.

In dieser Reihe von Dokumenten lernen Sie, wie Sie diese Tools verwenden. Sie werden in den Action Recorder eingeführt und erfahren, wie Sie Ihre häufigen Excel-Aktionen aufzeichnen können. Außerdem erfahren Sie, wie Sie eigene Skripts mit dem Code Editor erstellen oder aktualisieren können.

<br>

> [!VIDEO https://www.microsoft.com/videoplayer/embed/RE4qdFF]

## <a name="requirements"></a>Anforderungen

Wenn Sie Office-Skripts verwenden möchten, benötigen Sie Folgendes.

1. [Excel im Web](https://www.office.com/launch/excel) (Excel unter Windows kann Office-Skripts nur mit [Skriptschaltflächen](../develop/script-buttons.md) verwenden).

    > [!TIP]
    > Office-Skripts sind jetzt in Office unter Windows und auf dem Mac für [Office-Insider](https://insider.office.com/) verfügbar.

1. OneDrive for Business.
1. Jede kommerzielle oder pädagogische Microsoft 365-Lizenz mit Zugriff auf Microsoft 365 Office-Desktop-Apps wie:

    - Office 365 Business
    - Office 365 Business Premium
    - Office 365 ProPlus
    - Office 365 ProPlus für Geräte
    - Office 365 Enterprise E3
    - Office 365 Enterprise E5
    - Office 365 A3
    - Office 365 A5

> [!NOTE]
> Wenn Sie diese Anforderungen erfüllen und die Registerkarte **Automatisieren** immer noch nicht angezeigt wird, ist es möglich, dass Ihr Administrator das Feature deaktiviert hat oder ein anderes Problem mit Ihrer Umgebung vorliegt. Folgen Sie den Schritten unter [Automatisierungs-Registerkarte wird nicht angezeigt oder Office-Skripts sind nicht verfügbar](../testing/troubleshooting.md#automate-tab-not-appearing-or-office-scripts-unavailable), um Office-Skripts verwenden zu können.

## <a name="when-to-use-office-scripts"></a>Wann empfiehlt sich die Verwendung von Office-Skripts

Mit Skripts können Sie Ihre Excel-Aktionen in verschiedenen Arbeitsmappen und Arbeitsblättern aufzeichnen und wiedergeben. Wenn Sie immer wieder dieselben Aufgaben ausführen, können Sie alle diese Aufgaben in ein einfaches Office-Skript umwandeln. Führen Sie das Skript durch den Klick auf eine Schaltfläche in Excel aus, oder kombinieren Sie es mit Power Automation, um den gesamten Workflow zu optimieren.

Angenommen, Sie beginnen Ihren Arbeitstag, indem Sie eine .csv-Datei auf einer Kontoführungswebsite in Excel öffnen. Sie müssen dann mehrere Minuten lang nicht benötigte Spalten löschen, eine Tabelle formatieren, Formeln hinzufügen und eine PivotTable in einem neuen Arbeitsblatt erstellen. Die Aktivitäten, die Sie täglich wiederholen, können einmalig mit dem Action Recorder aufzeichnet werden. Ab dem Zeitpunkt wird durch Ausführen des Skripts die gesamte .csv-Konvertierung für Sie vorgenommen. Sie können nicht nur das Risiko minimieren, Schritte zu vergessen, sondern sind in der Lage, Ihren Prozess mit anderen zu teilen, ohne ihnen etwas beibringen zu müssen. Office-Skripts erlauben Ihnen, Ihre allgemeinen Aufgaben zu automatisieren, damit Sie und Ihre Arbeitsumgebung effizienter und produktiver arbeiten können.

## <a name="action-recorder"></a>Action Recorder

:::image type="content" source="../images/action-recorder-intro.png" alt-text="Eine Liste der vom Aktionsrekorder aufgezeichneten Aktionen.":::

Die Action Recorder zeichnet Aktionen auf, die Sie in Excel ausführen, und speichert diese als Skript. Während der Action Recorder ausgeführt wird, können Sie die Excel-Aktionen erfassen, während Sie Zellen bearbeiten, Formatierungen ändern und Tabellen erstellen. Das resultierende Skript kann in anderen Arbeitsblättern und Arbeitsmappen ausgeführt werden, um die ursprünglichen Aktionen neuerlich zu erstellen.

## <a name="code-editor"></a>Code Editor

:::image type="content" source="../images/code-editor-intro.png" alt-text="Der Code-Editor mit dem in diesem Lernprogramm verwendeten Skriptcode.":::

Alle mit dem Action Recorder aufgezeichneten Skripts können über den Code Editor bearbeitet werden. Auf diese Weise können Sie das Skript so optimieren und anpassen, so dass es Ihren genauen Anforderungen besser entspricht. Sie können auch Logiken und Funktionen hinzufügen, die nicht direkt über die Excel-Benutzeroberfläche zugänglich sind, z. B. bedingte Anweisungen (sofern/andernfalls) und Schleifen.

> [!TIP]
> Der Aktionsrekorder verfügt über eine Schaltfläche **Als Code kopieren**, um Aktionen in Skriptcode aufzuzeichnen, ohne das gesamte Skript zu speichern.
>
> :::image type="content" source="../images/action-recorder-copy-code.png" alt-text="Der Aufgabenbereich „Aktionsrekorder“ mit hervorgehobener Schaltfläche „Als Code kopieren“.":::

Unsere [Lernprogramme](../tutorials/excel-tutorial.md) bieten eine geführte und strukturierte Möglichkeit zum Erlernen der Funktionen von Office-Skripts. Lesen Sie nach Abschluss der Lernprogramme die [Grundlagen der Skripterstellung für Office-Skripts in Excel für das Web](../develop/scripting-fundamentals.md), um mehr über den Code-Editor zu erfahren und zu lernen, wie Sie Ihre eigenen Skripte schreiben und bearbeiten können. Zusätzliche Informationen über den Code-Editor und wie Ihr Skript-Code interpretiert wird, finden Sie unter [Code-Editor-Umgebung für Office-Skripts](code-editor-environment.md).

## <a name="share-office-scripts"></a>Freigeben von Office-Skripts

Office-Skripts können für andere Benutzer einer Excel-Arbeitsmappe freigegeben werden. Wenn Sie ein Skript in einer geteilten Arbeitsmappe freigeben, kann auch jeder Benutzer mit Zugriff auf die Arbeitsmappe Ihr Skript anzeigen und ausführen. Weitere Informationen zu Freigabe- und Freigabeskripts finden Sie unter [„Freigabe von Office-Skripts“ in Excel für das Web.](https://support.microsoft.com/office/226eddbc-3a44-4540-acfe-fccda3d1122b)

:::image type="content" source="../images/script-sharing.png" alt-text="Die Skriptdetailseite mit der Option &quot;Für andere Personen in dieser Arbeitsmappe freigeben&quot;.":::

Fügen Sie Schaltflächen hinzu, mit denen Skripts ausgeführt werden, damit Ihre Kollegen Ihre wertvollen Lösungen entdecken und Skripts in Excel auf dem Desktop ausführen können. Erfahren Sie mehr über Skriptschaltflächen in [Ausführung von Office-Skripts mit Schaltflächen](../develop/script-buttons.md).

:::image type="content" source="../images/add-button.png" alt-text="Eine Schaltfläche im Arbeitsblatt, die ein Skript ausführt, wenn darauf geklickt wird.":::

> [!NOTE]
> Erfahren Sie mehr darüber, wie Skripts in ihrem OneDrive gespeichert werden: [Office-Skripts-Dateispeicher und -Besitz](script-storage.md).

## <a name="connect-office-scripts-to-power-automate"></a>Verbinden von Office-Skripts mit Power Automate

[Power Automate](https://flow.microsoft.com/) ist ein Dienst, der Ihnen hilft, automatisierte Workflows zwischen mehreren Apps und Diensten zu erstellen. Office-Skripts können in diesen Workflows verwendet werden. Sie erhalten somit die Kontrolle über Ihre Skripts außerhalb der Arbeitsmappe. Sie können Ihre Skripts nach einem Zeitplan ausführen, sie als Antwort auf E-Mails auslösen und vieles mehr. Im Tutorial [Ausführen von Office-Skripts mit Power Automate](../tutorials/excel-power-automate-manual.md) lernen Sie die Grundlagen des Verbindungsaufbaus zu diesen Automatisierungsdiensten.

## <a name="next-steps"></a>Nächste Schritte

Führen Sie die [Office-Skripts in Excel im Web-Lernprogramm](../tutorials/excel-tutorial.md) aus, um zu erfahren, wie Sie Ihr erstes Office-Skript erstellen können.

## <a name="see-also"></a>Siehe auch

- [Grundlagen der Skripterstellung für Office-Skripts in Excel im Web](../develop/scripting-fundamentals.md)
- [Referenzdokumentation zur Office Scripts-API](/javascript/api/office-scripts/overview)
- [Behandeln von Problemen mit Office-Skripts](../testing/troubleshooting.md)
- [Office-Skripts-Einstellungen in M365](https://support.office.com/article/office-scripts-settings-in-m365-19d3c51a-6ca2-40ab-978d-60fa49554dcf)
- [Einführung in Office-Skripts in Excel](https://support.microsoft.com/office/9fbe283d-adb8-4f13-a75b-a81c6baf163a)
- [Office-Skripts in Excel für das Web](https://support.microsoft.com/office/226eddbc-3a44-4540-acfe-fccda3d1122b)
- [Office-Skripts Dev Center](https://developer.microsoft.com/office-scripts)
