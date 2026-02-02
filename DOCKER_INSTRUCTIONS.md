# Docker Anleitung für Antigravity

Dieses Dokument beschreibt, wie man die Applikation mithilfe von Docker ausführen kann.
Docker ermöglicht es, das gesamte System (Frontend + Backend + Datenbanken) in einer isolierten Umgebung zu starten, ohne Python oder Node.js lokal installieren zu müssen.

## Voraussetzungen
*   Docker Desktop muss installiert sein.

## Schnellstart

1.  Öffne ein Terminal im Projektordner.
2.  Führe folgenden Befehl aus:

```bash
docker-compose up --build
```

Das war's! 
*   Der Build-Prozess startet (dies kann beim ersten Mal ein paar Minuten dauern, da das Frontend kompiliert wird).
*   Die App ist danach unter `http://localhost:5000` erreichbar.

## Was passiert im Hintergrund?

Das `Dockerfile` führt einen **Multi-Stage Build** durch:
1.  **Stage 1 (Node.js)**: Installiert alle Frontend-Abhängigkeiten und baut die React-App (`npm run build`). Die fertigen HTML/CSS/JS Dateien landen im `dist` Ordner.
2.  **Stage 2 (Python)**:
    *   Installiert Python und die Pakete aus `requirements.txt` (`flask`, `python-pptx`, etc.).
    *   Kopiert die fertigen Frontend-Dateien an die richtige Stelle für Flask (`backend/templates` und `backend/static`).
    *   Passt die Pfade in der `index.html` automatisch an (damit Flask die Skripte findet).

## Daten-Persistenz

Die Datei `docker-compose.yml` sorgt dafür, dass deine Daten nicht verloren gehen, wenn du den Container stoppst.
Folgende Ordner auf deinem Computer werden in den Container "gemountet":
*   `./uploads` -> Hier landen die hochgeladenen CSVs.
*   `./backend/outputs` -> Hier landen die generierten PowerPoint-Dateien.

Das bedeutet: Du kannst die Ergebnisse ganz normal auf deinem Desktop im Ordner sehen, auch wenn sie im Container erzeugt wurden.

## Container beenden

Drücke im Terminal `STRG+C` oder führe in einem neuen Fenster aus:
```bash
docker-compose down
```
