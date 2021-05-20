---
title: Office Skriptbeispiele
description: Verfügbare Office Skripts-Beispielen und -Szenarien.
ms.date: 05/17/2021
localization_priority: Normal
ms.openlocfilehash: 0ea9a8a8986681fca0e45784e2923c1d3b34576d
ms.sourcegitcommit: 4687693f02fc90a57ba30c461f35046e02e6f5fb
ms.translationtype: MT
ms.contentlocale: de-DE
ms.lasthandoff: 05/19/2021
ms.locfileid: "52545709"
---
# <a name="office-scripts-samples-and-scenarios"></a>Office Skriptbeispiele und Szenarien

Dieser Abschnitt enthält [Office Skripts-basierten Automatisierungslösungen,](../../overview/excel.md) die Endbenutzern dabei helfen, die Automatisierung täglicher Aufgaben zu erreichen. Es enthält realistische Szenarien, denen Geschäftsanwender gegenüberstehen, und bietet detaillierte Lösungen zusammen mit Schritt-für-Schritt-Anleitungsvideolinks.

Für jedes der Projekte in [Basics](#basics) und [Beyond the Basics,](#beyond-the-basics)schauen Sie sich den Quellcode, Schritt-für-Schritt-YouTube-Videos und vieles mehr an. [](https://www.youtube.com/playlist?list=PLr3zVPZrMOUMl88fs8uc2GGAePRnNe6m0)

In [Szenarien](#scenarios)haben wir einige größere Szenariobeispiele enthalten, die reale Anwendungsfälle veranschaulichen.

Wir freuen uns auch über [Beiträge aus der Community](#community-contributions-and-fun-samples).

[!INCLUDE [Preview note](../../includes/preview-note.md)]

## <a name="basics"></a>Grundlagen

| Project | Details |
|---------|---------|
| [Skripting-Grundlagen](../excel-samples.md) | Diese Beispiele zeigen grundlegende Bausteine für Office Skripts. |
| [Hinzufügen von Kommentaren in Excel](add-excel-comments.md) | In diesem Beispiel werden einer Zelle Kommentare hinzugefügt, einschließlich @mentioning einem Kollegen. |
| [Hinzufügen von Bildern zu einer Arbeitsmappe](add-image-to-workbook.md) | In diesem Beispiel wird einer Arbeitsmappe ein Bild hinzugefügt und ein Bild über Blätter hinweg kopiert.|
| [Kopieren mehrerer Excel Tabellen in eine einzelne Tabelle](copy-tables-combine.md) | In diesem Beispiel werden Daten aus mehreren Excel Tabellen in einer einzigen Tabelle zusammengefasst, die alle Zeilen enthält. |

## <a name="beyond-the-basics"></a>Weitere Tipps und Tricks

Sehen Sie sich das folgende End-to-End-Projekt an, das Beispielszenarien zusammen mit vollständigen Skripts, Beispiel Excel verwendeten Dateien und [Videos (gehostet auf YouTube)](https://www.youtube.com/playlist?list=PLr3zVPZrMOUMl88fs8uc2GGAePRnNe6m0)automatisiert.

| Project | Details |
|---------|---------|
| [Zählen leerer Zeilen in einem bestimmten Blatt oder in allen Blättern](count-blank-rows.md) | In diesem Beispiel wird erkannt, ob leere Zeilen in Blättern vorhanden sind, in denen Sie die vorhandenen Daten erwarten und dann die anzahl leerer Zeilen für die Verwendung in einem Power Automate Fluss melden. |
| [E-Mail-Diagramm und Tabellenbilder](email-images-chart-table.md) | In diesem Beispiel werden Office Skripts und Power Automate Aktionen verwendet, um ein Diagramm zu erstellen und dieses Diagramm als Bild per E-Mail zu senden. |
| [Externe Abrufanrufe](external-fetch-calls.md) | In diesem Beispiel werden `fetch` Informationen aus GitHub für das Skript abgesucht. |
| [Filtern Excel Tabelle und abrufen des sichtbaren Bereichs](filter-table-get-visible-range.md) | In diesem Beispiel wird eine Excel Tabelle gefiltert und der sichtbare Bereich als JSON-Objekt zurückgegeben. Dieser JSON könnte einem Power Automate Flow als Teil einer größeren Lösung zur Verfügung gestellt werden. |
| [Verwalten des Berechnungsmodus in Excel](excel-calculation.md) | Dieses Beispiel zeigt, wie Sie den Berechnungsmodus verwenden und Methoden in Excel im Web mit Office Skripts berechnen. |
| [Verschieben von Zeilen über Tabellen](move-rows-across-tables.md) | In diesem Beispiel wird gezeigt, wie Zeilen über Tabellen verschoben werden, indem Filter gespeichert und anschließend verarbeitet und erneut angewendet werden. |
| [Ausgabe Excel Daten als JSON](get-table-data.md) | Diese Lösung zeigt, wie Excel Tabellendaten als JSON ausgegeben werden, um sie in Power Automate zu verwenden. |
| [Entfernen von Hyperlinks aus jeder Zelle in einem Excel Arbeitsblatt](remove-hyperlinks-from-cells.md) | In diesem Beispiel werden alle Hyperlinks aus dem aktuellen Arbeitsblatt entfernt. |
| [Ausführen eines Scripts für alle Excel-Dateien in einem Ordner](automate-tasks-on-all-excel-files-in-folder.md) | Dieses Projekt führt eine Reihe von Automatisierungsaufgaben für alle Dateien aus, die sich in einem Ordner auf OneDrive for Business befinden (kann auch für einen SharePoint Ordner verwendet werden). Es führt Berechnungen für die Excel Dateien durch, fügt Formatierungen hinzu und fügt einen Kommentar ein, der einen Kollegen @mentions. |
| [Erstellen eines großen Datasets](write-large-dataset.md) | In diesem Beispiel wird gezeigt, wie ein großer Bereich als kleinerer Teilbereich gesendet wird. |

## <a name="scenarios"></a>Szenarien

Office Skripts können Teile Ihrer täglichen Routine automatisieren. Diese täglichen Aufgaben existieren häufig in einzigartigen Ökosystemen, mit Excel Arbeitsmappen, die auf besondere Weise eingerichtet sind. Diese größeren Szenariobeispiele zeigen solche realen Anwendungsfälle. Sie umfassen sowohl die Office Skripts als auch die Arbeitsmappen, sodass Sie das Szenario von Ende zu Ende sehen können.

| Szenario | Details |
|---------|---------|
| [Analysieren von Webdownloads](../scenarios/analyze-web-downloads.md) | Dieses Szenario enthält ein Skript, das Webdatenverkehrsdatensätze analysiert, um das Herkunftsland eines Benutzers zu bestimmen. Es zeigt die Fähigkeiten der Textanalyse, der Verwendung von Unterfunktionen in Skripts, der Anwendung bedingter Formatierung und der Arbeit mit Tabellen. |
| [Wasserstandsdaten von NOAA abrufen und grafisch darstellen](../scenarios/noaa-data-fetch.md) | In diesem Szenario wird ein Office Skript verwendet, um Daten aus einer externen Quelle (der [NOAA-Gezeiten- und Stromdatenbank)](https://tidesandcurrents.noaa.gov/)abzurufen und die resultierenden Informationen grafisch zu zeichnen. Es zeigt die Fähigkeiten der Verwendung von Daten und die Verwendung von `fetch` Diagrammen. |
| [Bewertungsrechner](../scenarios/grade-calculator.md) | Dieses Szenario enthält ein Skript, das den Datensatz eines Kursleiters für die Klassennoten überprüft. Es zeigt die Fähigkeiten der Fehlerüberprüfung, Zellenformatierung und reguläre ausdrücke. |
| [Aufgabenerinnerungen](../scenarios/task-reminders.md) | In diesem Szenario wird ein Office Skript in einem Power Automate Ablauf verwendet, um Erinnerungen an Kollegen zu senden, um den Status eines Projekts zu aktualisieren. Es werden die Fähigkeiten Power Automate Integration und Datenübertragung zu und von Skripten hervorgehoben. |

## <a name="community-contributions-and-fun-samples"></a>Community Beiträge und lustige Proben

Wir freuen uns über [Beiträge](https://github.com/OfficeDev/office-scripts-docs/blob/master/Contributing.md) aus unserer Office Scripts Community! Fühlen Sie sich frei, eine Pull-Anfrage für die Überprüfung zu erstellen.

| Project | Details |
|---------|---------|
| [Spiel des Lebens](https://techcommunity.microsoft.com/t5/excel-blog/ready-player-zero/ba-p/2246208) | Der Blog "Ready Player Zero" von Yutao Huang auf der Excel Tech Community enthält ein Skript zum Modellieren von John Conways [*The Game of Life*](https://en.wikipedia.org/wiki/Conway%27s_Game_of_Life). |
| [Jahreszeiten Grüße Animation](community-seasons-greetings.md) | Dieses Drehbuch wurde von [Leslie Black](https://www.linkedin.com/in/lesblackconsultant/) im Geiste der Weihnachtszeit beigesteuert! Es ist ein lustiges Skript, das einen singenden Weihnachtsbaum in Excel im Web mit Office Scripts zeigt. |

## <a name="try-it-out"></a>Probieren Sie es aus

Diese Beispiele sind Open Source. Probieren Sie sie selbst aus. Sie benötigen ein Microsoft-Arbeits- oder Schulkonto von der Arbeit oder Schule mit einer Lizenz für Microsoft 365 Abonnement (E3 oder höher). Gehen Sie einfach https://office.com hinüber, um sich bei Ihrem Konto anzumelden und loszulegen.

## <a name="leave-a-comment"></a>Hinterlassen Sie einen Kommentar

Lassen Sie einen Kommentar hinterlassen, einen Vorschlag machen oder ein Problem protokollieren, indem Sie den Abschnitt **Feedback** am unteren Rand der Dokumentationsseite des jeweiligen Beispiels verwenden.
