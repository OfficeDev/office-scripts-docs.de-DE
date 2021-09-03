---
title: Office Skriptbeispiele
description: Verfügbare Beispiele und Szenarien für Office Skripts.
ms.date: 07/19/2021
localization_priority: Normal
ms.openlocfilehash: 042e49145285fe100cace1df2571f5aa06dab7d5
ms.sourcegitcommit: 6654aeae8a3ee2af84b4d4c4d8ff45b360a303eb
ms.translationtype: MT
ms.contentlocale: de-DE
ms.lasthandoff: 09/02/2021
ms.locfileid: "58862166"
---
# <a name="office-scripts-samples-and-scenarios"></a>Office Skriptbeispiele und -szenarien

Dieser Abschnitt enthält [Office Skripts-basierten](../../overview/excel.md) Automatisierungslösungen, die Endbenutzern helfen, tägliche Aufgaben zu automatisieren. Es enthält realistische Szenarien, mit denen Geschäftsbenutzer konfrontiert sind, und bietet detaillierte Lösungen zusammen mit schrittweisen Anleitungsvideolinks.

Sehen Sie sich für jedes der Projekte in ["Grundlagen"](#basics) und ["Über die Grundlagen hinaus"](#beyond-the-basics)den Quellcode, Schritt-für-Schritt-YouTube-Videos und vieles mehr an. [](https://www.youtube.com/playlist?list=PLr3zVPZrMOUMl88fs8uc2GGAePRnNe6m0)

In [Szenarien](#scenarios)haben wir ein paar größere Szenariobeispiele hinzugefügt, die anwendungsfälle in der Praxis veranschaulichen.

Wir freuen uns auch über [Beiträge der Community.](#community-contributions-and-fun-samples)

## <a name="basics"></a>Grundlagen

| Project | Details |
|---------|---------|
| [Skripting-Grundlagen](../excel-samples.md) | Diese Beispiele veranschaulichen grundlegende Bausteine für Office Skripts. |
| [Hinzufügen von Kommentaren in Excel](add-excel-comments.md) | In diesem Beispiel werden einer Zelle, die @mentioning einen Kollegen enthält, Kommentare hinzugefügt. |
| [Hinzufügen von Bildern zu einer Arbeitsmappe](add-image-to-workbook.md) | In diesem Beispiel wird einer Arbeitsmappe ein Bild hinzugefügt und ein Bild über Blätter hinweg kopiert.|
| [Kopieren mehrerer Excel Tabellen in eine einzelne Tabelle](copy-tables-combine.md) | In diesem Beispiel werden Daten aus mehreren Excel Tabellen in einer einzigen Tabelle kombiniert, die alle Zeilen enthält. |

## <a name="beyond-the-basics"></a>Weitere Tipps und Tricks

Sehen Sie sich das folgende End-to-End-Projekt an, das Beispielszenarien zusammen mit vollständigen Skripts, Beispielen Excel verwendeten Dateien und [Videos (gehostet auf YouTube)](https://www.youtube.com/playlist?list=PLr3zVPZrMOUMl88fs8uc2GGAePRnNe6m0)automatisiert.

| Project | Details |
|---------|---------|
| [Konvertieren von CSV-Dateien in Excel Arbeitsmappen](convert-csv.md) | In diesem Beispiel werden Office Skripts und Power Automate verwendet, um .xlsx Dateien aus .csv Dateien zu erstellen. |
| [Querverweis auf Arbeitsmappen](excel-cross-reference.md) | In diesem Beispiel werden Office Skripts und Power Automate zum Querverweisen und Überprüfen von Informationen in verschiedenen Arbeitsmappen verwendet. |
| [Zählen leerer Zeilen in einem bestimmten Blatt oder in allen Blättern](count-blank-rows.md) | In diesem Beispiel wird erkannt, ob leere Zeilen in Blättern vorhanden sind, in denen Daten voraussichtlich vorhanden sind, und dann die Anzahl leerer Zeilen für die Verwendung in einem Power Automate Fluss gemeldet. |
| [E-Mail-Diagramm- und Tabellenbilder](email-images-chart-table.md) | In diesem Beispiel werden Office Skripts und Power Automate Aktionen verwendet, um ein Diagramm zu erstellen und dieses Diagramm per E-Mail als Bild zu senden. |
| [Externe Abrufanrufe](external-fetch-calls.md) | In diesem Beispiel werden `fetch` Informationen aus GitHub für das Skript abgerufen. |
| [Filtern Excel Tabelle und Abrufen des sichtbaren Bereichs](filter-table-get-visible-range.md) | In diesem Beispiel wird eine Excel Tabelle gefiltert und der sichtbare Bereich als JSON-Objekt zurückgegeben. Dieser JSON-Code könnte einem Power Automate Fluss als Teil einer größeren Lösung bereitgestellt werden. |
| [Verwalten des Berechnungsmodus in Excel](excel-calculation.md) | Dieses Beispiel zeigt, wie Sie den Berechnungsmodus verwenden und Methoden in Excel im Web mit Office Skripts berechnen. |
| [Verschieben von Zeilen über Tabellen](move-rows-across-tables.md) | In diesem Beispiel wird gezeigt, wie Sie Zeilen über Tabellen verschieben, indem Sie Filter speichern und dann die Filter verarbeiten und erneut anwenden. |
| [Ausgabe Excel Daten als JSON](get-table-data.md) | Diese Lösung zeigt, wie Excel Tabellendaten als JSON ausgegeben werden, die in Power Automate verwendet werden. |
| [Entfernen von Hyperlinks aus jeder Zelle in einem Excel Arbeitsblatt](remove-hyperlinks-from-cells.md) | In diesem Beispiel werden alle Hyperlinks aus dem aktuellen Arbeitsblatt gelöscht. |
| [Ausführen eines Scripts für alle Excel-Dateien in einem Ordner](automate-tasks-on-all-excel-files-in-folder.md) | Dieses Projekt führt eine Reihe von Automatisierungsaufgaben für alle Dateien aus, die sich in einem Ordner auf OneDrive for Business befinden (kann auch für einen SharePoint Ordner verwendet werden). Es führt Berechnungen für die Excel Dateien durch, fügt Formatierungen hinzu und fügt einen Kommentar ein, der einen Kollegen @mentions. |
| [Erstellen eines großen Datasets](write-large-dataset.md) | Dieses Beispiel zeigt, wie ein großer Bereich als kleinere Unterbereiche gesendet wird. |

## <a name="scenarios"></a>Szenarien

Office Skripts können Teile Ihrer täglichen Routine automatisieren. Diese täglichen Aufgaben sind häufig in einzigartigen Ökosystemen vorhanden, mit Excel Arbeitsmappen, die auf bestimmte Weise eingerichtet sind. Diese größeren Szenariobeispiele veranschaulichen solche realen Anwendungsfälle. Sie enthalten sowohl die Office Skripts als auch die Arbeitsmappen, sodass Sie das Szenario von Ende bis Ende sehen können.

| Szenario | Details |
|---------|---------|
| [Analysieren von Webdownloads](../scenarios/analyze-web-downloads.md) | Dieses Szenario enthält ein Skript, das Webdatenverkehrsdatensätze analysiert, um das Ursprungsland eines Benutzers zu ermitteln. Es zeigt die Fähigkeiten der Textparsing, die Verwendung von Unterfunktionen in Skripts, das Anwenden bedingter Formatierung und das Arbeiten mit Tabellen. |
| [Wasserstandsdaten von NOAA abrufen und grafisch darstellen](../scenarios/noaa-data-fetch.md) | In diesem Szenario wird ein Office-Skript verwendet, um Daten aus einer externen Quelle (die [NOAA-Datenbank für Käufe und Currents)](https://tidesandcurrents.noaa.gov/)abzurufen und die resultierenden Informationen zu graphen. Es werden die Fähigkeiten der Verwendung zum Abrufen von `fetch` Daten und der Verwendung von Diagrammen hervorgehoben. |
| [Bewertungsrechner](../scenarios/grade-calculator.md) | Dieses Szenario enthält ein Skript, das den Datensatz eines Kursleiters auf die Noten seiner Klasse überprüft. Es zeigt die Fähigkeiten der Fehlerprüfung, Zellenformatierung und regulärer Ausdrücke. |
| [Vorstellungsgespräche in Teams planen](../scenarios/schedule-interviews-in-teams.md) | In diesem Szenario wird gezeigt, wie Sie eine Excel Kalkulationstabelle verwenden, um Besprechungszeiten zu verwalten und einen Fluss zum Planen von Besprechungen in Teams zu erstellen. |
| [Aufgabenerinnerungen](../scenarios/task-reminders.md) | In diesem Szenario wird ein Office Skript in einem Power Automate-Fluss verwendet, um Erinnerungen an Kollegen zu senden, um den Status eines Projekts zu aktualisieren. Es hebt die Fähigkeiten der Power Automate Integration und Datenübertragung zu und von Skripts hervor. |

## <a name="community-contributions-and-fun-samples"></a>Community Beiträge und Unterhaltungsbeispiele

Wir freuen uns über [Beiträge](https://github.com/OfficeDev/office-scripts-docs/blob/master/Contributing.md) unserer Office Scripts-Community! Sie können eine Pull-Anforderung zur Überprüfung erstellen.

| Project | Details |
|---------|---------|
| [Spiel des Lebens](https://techcommunity.microsoft.com/t5/excel-blog/ready-player-zero/ba-p/2246208) | Der Blog "Ready Player Zero" von Yutao The Excel Tech Community enthält ein Skript zum Modellieren von John Conways [*The Game of Life.*](https://en.wikipedia.org/wiki/Conway%27s_Game_of_Life) |
| [Begrüßungsanimation für "Spielzeiten"](community-seasons-greetings.md) | Dieses Skript wurde von [Leslie Black](https://www.linkedin.com/in/lesblackconsultant/) in der Feiertagszeit beigesteuert! Es ist ein interessantes Skript, das mithilfe von Office Skripts in Excel im Web eine tolle Tannenstruktur zeigt. |

## <a name="try-it-out"></a>Probieren Sie es aus

Diese Beispiele sind Open Source. Probieren Sie sie selbst aus. Sie benötigen ein Microsoft-Geschäfts-, Schul- oder Unikonto mit einer Lizenz für Microsoft 365 Abonnement (E3 oder höher). Melden Sie sich einfach bei https://office.com Ihrem Konto an und beginnen Sie.

## <a name="leave-a-comment"></a>Kommentar hinterlassen

Sie können einen Kommentar hinterlassen, einen Vorschlag machen oder  ein Problem mithilfe des Feedback-Abschnitts unten auf der Dokumentationsseite des jeweiligen Beispiels protokollieren.
