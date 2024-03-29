---
title: Beispiele für Office-Skripts
description: Verfügbare Beispiele und Szenarien für Office-Skripts.
ms.date: 08/25/2022
ms.localizationpriority: medium
ms.openlocfilehash: 5798da37bd4166d18b41c005c4d8cc8a4b6c401d
ms.sourcegitcommit: a6504f8b0d6b717457c6e0b5306c35ad3900914e
ms.translationtype: MT
ms.contentlocale: de-DE
ms.lasthandoff: 09/02/2022
ms.locfileid: "67572486"
---
# <a name="office-scripts-samples-and-scenarios"></a>Beispiele und Szenarien für Office-Skripts

Dieser Abschnitt enthält [auf Office-Skripts basierende](../../overview/excel.md) Automatisierungslösungen, mit denen Endbenutzer die Automatisierung täglicher Aufgaben erreichen können. Es enthält realistische Szenarien, mit denen Geschäftsbenutzer konfrontiert sind, und bietet detaillierte Lösungen zusammen mit schrittweisen Videolinks mit Anleitungen.

Sehen Sie sich für jedes der Projekte in " [Grundlagen](#basics) " und " [Über die Grundlagen hinaus](#beyond-the-basics)" den Quellcode, schrittweise [**YouTube-Videos**](https://www.youtube.com/playlist?list=PLr3zVPZrMOUMl88fs8uc2GGAePRnNe6m0) und vieles mehr an.

In [Szenarien](#scenarios) haben wir einige größere Szenariobeispiele einbezogen, die reale Anwendungsfälle veranschaulichen.

Wir freuen uns auch [über Beiträge aus der Community](#community-contributions-and-fun-samples). Diese Beispiele sind Open Source.

> [!IMPORTANT]
> Stellen Sie sicher, dass Sie die Voraussetzungen für Office-Skripts erfüllen, bevor Sie die Beispiele ausprobieren. Die Anforderungen für Ihr Microsoft 365-Abonnement und -Konto finden Sie im [Abschnitt "Anforderungen" im Abschnitt "Office-Skripts für Excel– Übersicht"](../../overview/excel.md#requirements).

## <a name="basics"></a>Grundlagen

| Project | Details |
|---------|---------|
| [Skripting-Grundlagen](excel-samples.md) | Diese Beispiele veranschaulichen grundlegende Bausteine für Office-Skripts. |
| [Hinzufügen von Kommentaren in Excel](add-excel-comments.md) | In diesem Beispiel werden einer Zelle Kommentare hinzugefügt, einschließlich @mentioning eines Kollegen. |
| [Hinzufügen von Bildern zu einer Arbeitsmappe](add-image-to-workbook.md) | In diesem Beispiel wird einer Arbeitsmappe ein Bild hinzugefügt und ein Bild über Blätter hinweg kopiert.|
| [Kopieren mehrerer Excel-Tabellen in eine einzelne Tabelle](copy-tables-combine.md) | In diesem Beispiel werden Daten aus mehreren Excel-Tabellen in einer einzigen Tabelle kombiniert, die alle Zeilen enthält. |
| [Inhaltsverzeichnis für eine Arbeitsmappe erstellen](table-of-contents.md) | In diesem Beispiel wird ein Inhaltsverzeichnis mit Verknüpfungen zu jedem Arbeitsblatt erstellt. |
| [Tabellenspaltenfilter entfernen](clear-table-filter-for-active-cell.md) | In diesem Beispiel werden alle Filter aus einer Tabellenspalte gelöscht. |
| [Aufzeichnen von täglichen Änderungen in Excel und Berichterstellung mit einem Power Automate-Fluss](report-day-to-day-changes.md) | In diesem Beispiel wird ein geplanter Power Automate-Fluss verwendet, um tägliche Messwerte aufzuzeichnen und die Änderungen zu melden. |

## <a name="beyond-the-basics"></a>Weitere Tipps und Tricks

Sehen Sie sich das folgende End-to-End-Projekt an, das Beispielszenarien zusammen mit vollständigen Skripts, verwendeten Excel-Beispieldateien und [Videos (auf YouTube gehostet)](https://www.youtube.com/playlist?list=PLr3zVPZrMOUMl88fs8uc2GGAePRnNe6m0) automatisiert.

| Project | Details |
|---------|---------|
| [Kombinieren von Arbeitsblättern in einer einzelnen Arbeitsmappe](combine-worksheets-into-single-workbook.md) | In diesem Beispiel werden Office-Skripts und Power Automate verwendet, um Daten aus anderen Arbeitsmappen in eine einzelne Arbeitsmappe zu übertragen. |
| [Konvertieren von CSV-Dateien in Excel-Arbeitsmappen](convert-csv.md) | In diesem Beispiel werden Office-Skripts und Power Automate verwendet, um .xlsx Dateien aus .csv Dateien zu erstellen. |
| [Querverweis-Arbeitsmappen](excel-cross-reference.md) | In diesem Beispiel werden Office-Skripts und Power Automate verwendet, um Informationen in verschiedenen Arbeitsmappen querzuverweisen und zu überprüfen. |
| [Zählen leerer Zeilen in einem bestimmten Blatt oder in allen Blättern](count-blank-rows.md) | In diesem Beispiel wird erkannt, ob leere Zeilen in Blättern vorhanden sind, in denen Sie davon ausgehen, dass Daten vorhanden sind, und dann die Anzahl leerer Zeilen für die Verwendung in einem Power Automate-Fluss melden. |
| [Email Diagramm- und Tabellenbilder](email-images-chart-table.md) | In diesem Beispiel werden Office-Skripts und Power Automate-Aktionen verwendet, um ein Diagramm zu erstellen und es als Bild per E-Mail zu senden. |
| [Externe Abrufaufrufe](external-fetch-calls.md) | In diesem Beispiel werden `fetch` Informationen von GitHub für das Skript abgerufen. |
| [Verwalten des Berechnungsmodus in Excel](excel-calculation.md) | In diesem Beispiel wird gezeigt, wie Sie den Berechnungsmodus verwenden und Methoden in Excel im Web mithilfe von Office-Skripts berechnen. |
| [Verschieben von Zeilen über Tabellen hinweg](move-rows-across-tables.md) | In diesem Beispiel wird gezeigt, wie Sie Zeilen über Tabellen hinweg verschieben, indem Sie Filter speichern, dann die Filter verarbeiten und erneut anwenden. |
| [Excel-Daten als JSON ausgeben](get-table-data.md) | Diese Lösung zeigt, wie Excel-Tabellendaten als JSON ausgegeben werden, um sie in Power Automate zu verwenden. |
| [Entfernen von Links aus jeder Zelle in einem Excel-Arbeitsblatt](remove-hyperlinks-from-cells.md) | In diesem Beispiel werden alle Hyperlinks aus dem aktuellen Arbeitsblatt gelöscht. |
| [Ausführen eines Scripts für alle Excel-Dateien in einem Ordner](automate-tasks-on-all-excel-files-in-folder.md) | Dieses Projekt führt eine Reihe von Automatisierungsaufgaben für alle Dateien aus, die sich in einem Ordner auf OneDrive for Business befinden (kann auch für einen SharePoint-Ordner verwendet werden). Es führt Berechnungen für die Excel-Dateien aus, fügt Formatierung hinzu und fügt einen Kommentar ein, der einen Kollegen @mentions. |
| [Erstellen eines großen Datasets](write-large-dataset.md) | In diesem Beispiel wird gezeigt, wie ein großer Bereich als kleinere Unterbereiche gesendet wird. |

## <a name="scenarios"></a>Szenarien

Office-Skripts können Teile Ihrer täglichen Routine automatisieren. Diese täglichen Aufgaben befinden sich häufig in einzigartigen Ökosystemen, mit Excel-Arbeitsmappen, die auf besondere Weise eingerichtet sind. Diese größeren Szenariobeispiele veranschaulichen solche realen Anwendungsfälle. Sie enthalten sowohl die Office-Skripts als auch die Arbeitsmappen, sodass Sie das Szenario von Ende zu Ende sehen können.

| Szenario | Details |
|---------|---------|
| [Analysieren von Webdownloads](../scenarios/analyze-web-downloads.md) | Dieses Szenario enthält ein Skript, das Webdatenverkehrsdatensätze analysiert, um das Herkunftsland eines Benutzers zu ermitteln. Es zeigt die Fähigkeiten der Textparsierung, die Verwendung von Unterfunktionen in Skripts, das Anwenden bedingter Formatierung und das Arbeiten mit Tabellen. |
| [Wasserstandsdaten von NOAA abrufen und grafisch darstellen](../scenarios/noaa-data-fetch.md) | In diesem Szenario wird ein Office-Skript verwendet, um Daten aus einer externen Quelle (der [NOAA Tides and Currents-Datenbank](https://tidesandcurrents.noaa.gov/)) abzurufen und die resultierenden Informationen darzustellen. Es hebt die Fähigkeiten der Verwendung `fetch` von Daten und der Verwendung von Diagrammen hervor. |
| [Bewertungsrechner](../scenarios/grade-calculator.md) | Dieses Szenario enthält ein Skript, das den Datensatz eines Kursleiters für die Noten seines Kurses überprüft. Es zeigt die Fähigkeiten der Fehlerüberprüfung, Zellformatierung und reguläre Ausdrücke. |
| [Vorstellungsgespräche in Teams planen](../scenarios/schedule-interviews-in-teams.md) | In diesem Szenario wird gezeigt, wie Sie mithilfe eines Excel-Arbeitsblatts Besprechungszeiten verwalten und einen Fluss zum Planen von Besprechungen in Teams erstellen. |
| [Aufgabenerinnerungen](../scenarios/task-reminders.md) | In diesem Szenario wird ein Office-Skript in einem Power Automate-Fluss verwendet, um Erinnerungen an Kollegen zu senden, um den Status eines Projekts zu aktualisieren. Es hebt die Fähigkeiten der Power Automate-Integration und Datenübertragung zu und von Skripts hervor. |

## <a name="community-contributions-and-fun-samples"></a>Community-Beiträge und lustige Beispiele

Wir freuen uns über [Beiträge](https://github.com/OfficeDev/office-scripts-docs/blob/master/Contributing.md) aus unserer Office Scripts-Community! Sie können eine Pull-Anforderung zur Überprüfung erstellen.

| Project | Details |
|---------|---------|
| [Spiel des Lebens](https://techcommunity.microsoft.com/t5/excel-blog/ready-player-zero/ba-p/2246208) | Der Blog "Ready Player Zero" von Yutao Huang in der Excel Tech Community enthält ein Skript zum Modellieren von John Conways [*The Game of Life*](https://en.wikipedia.org/wiki/Conway%27s_Game_of_Life). |
| [Stempeluhr-Taste](../scenarios/punch-clock.md) | Dieses Skript wurde von [Brian Gonzalez](https://github.com/b-gonzalez) beigesteuert. Das Szenario enthält ein Skript und eine Skriptschaltfläche, die die aktuelle Uhrzeit aufzeichnet. |
| [Animation "Seasons Greetings"](community-seasons-greetings.md) | Dieses Skript wurde von [Leslie Black](https://www.linkedin.com/in/lesblackconsultant/) im Geiste der Ferienzeit beigetragen! Es ist ein lustiges Skript, das einen singenden Weihnachtsbaum in Excel im Web mitHilfe von Office-Skripts zeigt. |

## <a name="leave-a-comment"></a>Hinterlasse einen Kommentar

Sie können einen Kommentar hinterlassen, einen Vorschlag machen oder ein Problem protokollieren, indem Sie den Abschnitt **"Feedback** " unten auf der Dokumentationsseite des jeweiligen Beispiels verwenden.
