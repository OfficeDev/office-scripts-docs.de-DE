---
title: Office-Skripts in Excel im Web
description: Eine kurze Einführung in den Action Recorder und den Code Editor für Office-Skripts.
ms.date: 05/24/2021
localization_priority: Priority
ms.openlocfilehash: 93d6457338442472fc691d6d020dff99e7963b1b
ms.sourcegitcommit: 4693c8f79428ec74695328275703af0ba1bfea8f
ms.translationtype: HT
ms.contentlocale: de-DE
ms.lasthandoff: 06/23/2021
ms.locfileid: "53074599"
---
# <a name="office-scripts-in-excel-on-the-web"></a>Office-Skripts in Excel im Web

Mit Office-Skripts in Excel im Web können Sie Ihre täglichen Aufgaben automatisieren. Sie können Ihre Excel-Aktionen mit dem Action Recorder aufzeichnen, wodurch ein TypeScript-Sprachskript erstellt wird. Sie können Skripts auch mit dem Code Editor erstellen und bearbeiten. Ihre Skripts können dann für ihre gesamte Organisation freigegeben werden, sodass Ihre Kollegen den Workflow entsprechend automatisieren können.

In dieser Reihe von Dokumenten lernen Sie, wie Sie diese Tools verwenden. Sie werden in den Action Recorder eingeführt und erfahren, wie Sie Ihre häufigen Excel-Aktionen aufzeichnen können. Außerdem erfahren Sie, wie Sie eigene Skripts mit dem Code Editor erstellen oder aktualisieren können.

<br>

> [!VIDEO https://www.microsoft.com/videoplayer/embed/RE4qdFF]

## <a name="requirements"></a>Anforderungen

Wenn Sie Office-Skripts verwenden möchten, benötigen Sie Folgendes.

1. [Excel im Web](https://www.office.com/launch/excel) (andere Plattformen, z. B. Desktop, werden nicht unterstützt).
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

Eine einfache Möglichkeit, die Fähigkeiten von Office-Skripts zu erlernen besteht darin, Skripts in Excel im Web aufzuzeichnen und sich den resultierenden Code anzeigen zu lassen. Eine weitere Möglichkeit besteht darin, unseren [Lernprogrammen](../tutorials/excel-tutorial.md) zu folgen, um geleitet und strukturierter zu lernen.

Lesen Sie nach Abschluss des Lernprogramms die [Grundlagen der Skripterstellung für Office-Skripte in Excel für das Web](../develop/scripting-fundamentals.md), um mehr über den Code-Editor zu erfahren und zu lernen, wie Sie Ihre eigenen Skripte schreiben und bearbeiten können. Zusätzliche Informationen über den Code-Editor und wie Ihr Skript-Code interpretiert wird, finden Sie unter [Code-Editor-Umgebung für Office-Skripts](code-editor-environment.md).

## <a name="sharing-scripts"></a>Freigeben von Skripts

:::image type="content" source="../images/script-sharing.png" alt-text="Die Seite „Skriptdetails“ mit der Option ‚Für andere Personen in dieser Arbeitsmappe freigeben‘":::

Office-Skripts können für andere Benutzer einer Excel-Arbeitsmappe freigegeben werden. Wenn Sie ein Skript in einer geteilten Arbeitsmappe freigeben, kann auch jeder Benutzer mit Zugriff auf die Arbeitsmappe Ihr Skript anzeigen und ausführen.

Weitere Informationen zur Freigabe und zum Aufheben der Freigabe von Skripts entnehmen Sie dem Artikel [Freigeben von Office-Skripts in Excel für das Web](https://support.microsoft.com/office/sharing-office-scripts-in-excel-for-the-web-226eddbc-3a44-4540-acfe-fccda3d1122b).

> [!NOTE]
> Erfahren Sie mehr darüber, wie Skripts in ihrem OneDrive gespeichert werden: [Office-Skripts-Dateispeicher und -Besitz](script-storage.md).

## <a name="connecting-office-scripts-to-power-automate"></a>Verbinden von Office-Skripts mit Power Automate

[Power Automate](https://flow.microsoft.com/) ist ein Dienst, der Ihnen hilft, automatisierte Workflows zwischen mehreren Apps und Diensten zu erstellen. Office-Skripts können in diesen Workflows verwendet werden. Sie erhalten somit die Kontrolle über Ihre Skripts außerhalb der Arbeitsmappe. Sie können Ihre Skripts nach einem Zeitplan ausführen, sie als Antwort auf E-Mails auslösen und vieles mehr. Im Lernprogramm [Ausführen von Office-Skripts in Excel im Web mit Power Automate](../tutorials/excel-power-automate-manual.md) erfahren Sie die Grundlagen des Verbindungsaufbaus zu diesen Automatisierungsdiensten.

## <a name="next-steps"></a>Nächste Schritte

Führen Sie die [Office-Skripts in Excel im Web-Lernprogramm](../tutorials/excel-tutorial.md) aus, um zu erfahren, wie Sie Ihr erstes Office-Skript erstellen können.

## <a name="see-also"></a>Siehe auch

- [Grundlagen der Skripterstellung für Office-Skripts in Excel im Web](../develop/scripting-fundamentals.md)
- [Referenzdokumentation zur Office Scripts-API](/javascript/api/office-scripts/overview)
- [Behandeln von Problemen mit Office-Skripts](../testing/troubleshooting.md)
- [Office-Skripts-Einstellungen in M365](https://support.office.com/article/office-scripts-settings-in-m365-19d3c51a-6ca2-40ab-978d-60fa49554dcf)
- [Einführung in Office-Skripts in Excel (unter support.office.com)](https://support.office.com/article/introduction-to-office-scripts-in-excel-9fbe283d-adb8-4f13-a75b-a81c6baf163a)
- [Office-Skripts in Excel für das Web](https://support.microsoft.com/office/sharing-office-scripts-in-excel-for-the-web-226eddbc-3a44-4540-acfe-fccda3d1122b)
