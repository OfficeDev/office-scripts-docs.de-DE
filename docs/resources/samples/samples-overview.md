---
title: Office Skriptbeispiele
description: Verfügbare Office Skripts Beispiele und Szenarien.
ms.date: 04/05/2021
localization_priority: Normal
ms.openlocfilehash: dc09db00cb63e6873b255360aff17ad2a56fa89e
ms.sourcegitcommit: 1f003c9924e651600c913d84094506125f1055ab
ms.translationtype: MT
ms.contentlocale: de-DE
ms.lasthandoff: 04/26/2021
ms.locfileid: "52026828"
---
# <a name="office-scripts-samples-and-scenarios"></a>Office Skriptbeispiele und -szenarien

Dieser Abschnitt enthält [Office Skripts](../../overview/excel.md) basierende Automatisierungslösungen, mit deren Hilfe Endbenutzer die Automatisierung täglicher Aufgaben erreichen können. Es enthält realistische Szenarien, denen Geschäftsbenutzer gegenüber stehen, und bietet detaillierte Lösungen sowie schrittweise Anleitungen für Videolinks.

Für jedes der Projekte in [Basics](#basics) und Beyond the [basics](#beyond-the-basics)lesen Sie den Quellcode, schrittweise [**YouTube-Videos**](https://www.youtube.com/playlist?list=PLr3zVPZrMOUMl88fs8uc2GGAePRnNe6m0)und vieles mehr.

In [Szenarien](#scenarios)haben wir einige größere Szenariobeispiele enthalten, die reale Anwendungsfälle veranschaulichen.

Wir freuen uns auch [über Beiträge von der Community](#community-contributions).

[!INCLUDE [Preview note](../../includes/preview-note.md)]

## <a name="basics"></a>Grundlagen

| Project | Details |
|---------|---------|
| [Skripting-Grundlagen](../excel-samples.md) | Diese Beispiele zeigen grundlegende Bausteine für Office Skripts. |
| [Grundlegendes zur Verwendung des Range-Objekts in Office Skripts](range-basics.md) | In diesem Artikel werden die Grundlagen der Verwendung des Range-Objekts und seiner APIs erläutert. Dies ist ein Grundthema, das in allen anderen Projekten verwendet wird. |

## <a name="beyond-the-basics"></a>Weitere Tipps und Tricks

Sehen Sie sich das folgende End-to-End-Projekt an, das Beispielszenarien zusammen mit vollständigen Skripts, Beispieldateien Excel verwendeten Dateien und Videos [automatisiert.](https://www.youtube.com/playlist?list=PLr3zVPZrMOUMl88fs8uc2GGAePRnNe6m0)

| Project | Details |
|---------|---------|
| [Hinzufügen von Kommentaren in Excel](add-excel-comments.md) | In diesem Beispiel wird gezeigt, wie Sie einer Zelle Kommentare hinzufügen, einschließlich @mentioning Kollegen. |
| [Zählen leerer Zeilen in einem bestimmten Blatt oder in allen Blättern](count-blank-rows.md) | In diesem Beispiel wird ermittelt, ob leere Zeilen in Blättern vorhanden sind, in denen Daten vorhanden sein sollen, und anschließend die Leere Zeilenanzahl für die Verwendung in einem Power Automate melden. |
| [Querverweis und Formatieren einer Excel Datei](excel-cross-reference.md) | Diese Lösung zeigt, wie zwei Excel mithilfe von Skripts und Office formatiert und formatiert werden Power Automate. |
| [E-Mail-Diagramm- und Tabellenbilder](email-images-chart-table.md) | In diesem Beispiel werden Office Skripts und Power Automate verwendet, um ein Diagramm zu erstellen und dieses Diagramm als Bild per E-Mail zu senden. |
| [Externe Abrufaufrufe](external-fetch-calls.md) | In diesem Beispiel `fetch` werden Informationen aus GitHub Skripts erhalten. |
| [Filtern Excel Tabelle und Erhalten des sichtbaren Bereichs](filter-table-get-visible-range.md) | In diesem Beispiel wird eine Excel filtert und der sichtbare Bereich als JSON-Objekt zurückgegeben. Diese JSON könnte einem Power Automate als Teil einer größeren Lösung bereitgestellt werden. |
| [Generieren eines eindeutigen Bezeichners in einer Arbeitsmappe](document-number-generator.md) | Dieses Szenario hilft einem Benutzer, eine eindeutige Dokumentnummer mit einem bestimmten Format zu generieren und einem Bereich oder einer Tabelle einen Eintrag hinzuzufügen. |
| [Verwalten des Berechnungsmodus in Excel](excel-calculation.md) | In diesem Beispiel wird gezeigt, wie Sie den Berechnungsmodus verwenden und Methoden in Excel im Web mithilfe Office berechnen. |
| [Zusammenführen mehrerer Excel Tabellen in eine einzelne Tabelle](copy-tables-combine.md) | In diesem Beispiel werden Daten aus Excel Tabellen in einer einzigen Tabelle kombiniert, die alle Zeilen enthält. |
| [Verschieben von Zeilen über Tabellen](move-rows-across-tables.md) | In diesem Beispiel wird gezeigt, wie Sie Zeilen über Tabellen verschieben, indem Sie Filter speichern und anschließend die Filter verarbeiten und erneut anwenden. |
| [Ausgabe Excel Daten als JSON](get-table-data.md) | Diese Lösung zeigt, wie Excel als JSON ausgegeben werden, die in der Power Automate. |
| [Entfernen von Hyperlinks aus jeder Zelle in Excel Arbeitsblatt](remove-hyperlinks-from-cells.md) | In diesem Beispiel werden alle Hyperlinks aus dem aktuellen Arbeitsblatt entfernt. |
| [Ausführen eines Scripts für alle Excel-Dateien in einem Ordner](automate-tasks-on-all-excel-files-in-folder.md) | Dieses Projekt führt eine Reihe von Automatisierungsaufgaben für alle Dateien aus, die sich in einem Ordner auf OneDrive for Business befinden (kann auch für einen Ordner SharePoint werden). Es führt Berechnungen für die Excel aus, fügt Formatierungen hinzu und fügt einen Kommentar ein, der @mentions kollegen. |
| [Senden einer Teams aus Excel Daten](send-teams-invite-from-excel-data.md) | Diese Lösung zeigt, wie Sie Office Skripts und Power Automate-Aktionen verwenden, um Zeilen aus der Excel-Datei auszuwählen und sie zu verwenden, um eine Teams-Besprechungs-Einladung zu senden und dann Excel. |

## <a name="scenarios"></a>Szenarien

Office Skripts können Teile Ihrer täglichen Routine automatisieren. Diese täglichen Aufgaben sind häufig in einzigartigen Ökosystemen vorhanden, mit Excel arbeitsmappen, die auf besondere Weise eingerichtet sind. Diese größeren Szenariobeispiele veranschaulichen solche realen Verwendungsfälle. Sie enthalten sowohl die Office Skripts als auch die Arbeitsmappen, sodass Sie das Szenario von Ende zu Ende sehen können.

| Szenario | Details |
|---------|---------|
| [Analysieren von Webdownloads](../scenarios/analyze-web-downloads.md) | Dieses Szenario verfügt über ein Skript, das Webdatenverkehrdatensätze analysiert, um das Ursprungsland eines Benutzers zu bestimmen. Es zeigt die Fähigkeiten der Textparsierung, unter Verwendung von Unterfunktion in Skripts, anwenden bedingter Formatierung und Arbeiten mit Tabellen. |
| [Wasserstandsdaten von NOAA abrufen und grafisch darstellen](../scenarios/noaa-data-fetch.md) | In diesem Szenario wird ein Office-Skript verwendet, um Daten aus einer externen Quelle (der [NOAA-Flut-](https://tidesandcurrents.noaa.gov/)und Currents-Datenbank) zu ziehen und die resultierenden Informationen zu graphieren. Es hebt die Fähigkeiten der Verwendung von `fetch` Daten und der Verwendung von Diagrammen hervor. |
| [Bewertungsrechner](../scenarios/grade-calculator.md) | Dieses Szenario verfügt über ein Skript, das den Datensatz eines Kursleiters für die Noten ihrer Klasse überprüft. Es zeigt die Fähigkeiten der Fehlerüberprüfung, der Zellformatierung und regulärer Ausdrücke. |
| [Aufgabenerinnerungen](../scenarios/task-reminders.md) | In diesem Szenario wird Office Skript in einem Power Automate verwendet, um Erinnerungen an Kollegen zu senden, um den Status eines Projekts zu aktualisieren. Es unterstreicht die Fähigkeiten Power Automate Integration und Datenübertragung zu und von Skripts. |

## <a name="community-contributions"></a>Beiträge der Community

Wir freuen [uns über](https://github.com/OfficeDev/office-scripts-docs/blob/master/Contributing.md) Beiträge unserer Office Scripts-Community! Sie können eine Pullanforderung zur Überprüfung erstellen.

| Project | Details |
|---------|---------|
| [Begrüßungsanimation für die Jahreszeiten](community-seasons-greetings.md) | Dieses Skript wurde von [Leslie Black](https://www.linkedin.com/in/lesblackconsultant/) im Geiste der Feiertagszeit beigetragen! Es ist ein lustiges Skript, das einen besingten Weihnachtsbaum in einem Excel im Web mit Office zeigt. |

## <a name="try-it-out"></a>Probieren Sie es aus

Diese Beispiele sind Open Source. Testen Sie sie selbst. Sie benötigen ein Microsoft-Arbeits- oder Schulkonto von der Arbeit oder Schule mit einer Lizenz zum Microsoft 365 (E3 oder höher). Melden Sie sich https://office.com einfach bei Ihrem Konto an, und beginnen Sie.

## <a name="leave-a-comment"></a>Hinterlassen eines Kommentars

Sie können einen Kommentar hinterlassen, einen Vorschlag machen oder ein Problem mithilfe des Abschnitts **Feedback** am Ende der Dokumentationsseite des jeweiligen Beispiels protokollieren.
