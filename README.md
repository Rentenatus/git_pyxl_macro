# git_pyxl_macro
Tool zur Migration von VBA-Referenzrechnern in Excel in Python-Services, die ebenfalls Excel als Einngabe- und Ausgabemedium benutzen.



# Motivation
DORA definiert, was an digitaler Resilienz erreicht werden muss – aber nicht, wie die Umsetzung konkret auszusehen hat. Dieses Projekt unterstützt dabei, bestehende VBA-Referenzrechner aus der „Individualsoftware-Ecke“ herauszuholen und in eine strukturierte, versionierbare und auditierbare Python-Codebasis zu überführen.

In vielen Häusern laufen fachkritische Referenzrechner noch als Excel/VBA-Lösungen. Sie werden nicht nur für Tests, sondern auch gerne Mal in der Produktion eingesetzt. Sie sind sehr schwer zu versionieren, die Formeln sind kaum nachvollziehbar und die Funktion schwer zu dokumentieren. 
Damit ist die regulatorisch konforme Pflege der Referenzrechner (DORA, IDV, Revisionssicherheit) nur mit großem und aufgesetztem Aufwand abzusichern.
Außerdem benutzt dieses Repo eine lokal installierte Ollama Instanz, was ein Vorteil gegenüber möglicherweise strengen Datenschutzvorschriften gegenüber Cloudbasierten Modellen sein kann.

Ziel dieses Repos ist es, einen reproduzierbaren, halbautomatisierten Weg anzubieten, um:
- VBA-Makros technisch aufzubereiten,
- ihre Logik sauber zu dokumentieren,
- robusten Python-Code zu erzeugen.
- die Migration trotz LLMs lokal und damit gesichert auszuführen

Leider erzeugt auch dieses Repo wie jeder 1 zu 1 Übersetzer einen „Workslop“ in modernen Mantel.

Dieses Tool kann also nur ein erster Schritt zu einer echten Softwaremigration sein.

## Abgrenzung

Dieses Repo führt nicht die Softwarearchitektur ein, die bei professionellen Modernisierung zu erwartet wäre.


# Funktionsumfang
Der aktuelle Workflow besteht aus mehreren Schritten, die nacheinander ausgeführt werden:

- Auslesen der VBA-Makros

- Extraktion des VBA-Codes aus Excel-Dateien

- Technische Zerlegung in Abschnitte, Deklarationen und Prozeduren

- Iterative Dokumentation pro Chunk

- Generierung von Deklarationen in Python

- Generierung von Methodensignaturen je Prozedur

- Erzeugung von Python-Code je Prozedur



## Ziel
Dokumentation, 
Bessere Kontrolle über Versionierung, Zugriff, Logging und Ausfallsicherheit.

## Zielgruppe
Aktuariate und Fachbereiche, die ihre Referenzrechner und IDV-Lösungen modernisieren wollen.
IT- und Architektur-Teams, die regulatorisch saubere Wege aus der Excel/VBA-Welt suchen.
Compliance- und Revisionsverantwortliche, die Nachvollziehbarkeit, Audit-Trails und Dokumentation stärken möchten.
