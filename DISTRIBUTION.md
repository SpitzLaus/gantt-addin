# Verteilung ohne Serverstart & ohne Installation

Ziel: **Kollegen sollen das Add-in nutzen, ohne Node/npm zu installieren und ohne
einen lokalen Server zu starten.**

Der Trick: Die Add-in-Oberfläche wird **einmal** auf einem öffentlichen HTTPS-Host
(z. B. GitHub Pages) veröffentlicht. Danach muss jeder Kollege nur **einmal ein
kleines Skript ausführen** (oder das Manifest sideloaden) – fertig.

> Warum kein reines „Datei öffnen"? Office-Web-Add-ins **müssen** ihre Oberfläche
> über HTTPS laden. Eine lokale HTML-Datei ohne Server ist technisch nicht möglich.
> Deshalb: einmal hosten, beliebig oft nutzen.

---

## Schritt 1 (einmalig, nur du): Hosten

1. Self-contained Build erzeugen und Produktions-Manifest für deine URL bauen:

   ```powershell
   .\make-prod.ps1 -HostUrl "https://<dein-user>.github.io/<repo>"
   ```

   Das erzeugt:
   - `docs\taskpane.html` – **eine** Datei mit eingebettetem CSS/JS (keine Abhängigkeiten)
   - `docs\assets\*` – Icons
   - `manifest.prod.xml` – Manifest, das auf deine URL zeigt
   - `docs\manifest.xml` – Kopie zum Weitergeben

2. Ordner `docs` auf **GitHub Pages** veröffentlichen:
   - Neues GitHub-Repo anlegen, Dateien pushen.
   - **Settings → Pages → Source: „Deploy from a branch" → Branch `main`, Ordner `/docs`.**
   - Nach ~1 Minute ist die App unter
     `https://<dein-user>.github.io/<repo>/taskpane.html` erreichbar.

   > Alternativen zu GitHub Pages: jeder statische HTTPS-Host
   > (Azure Static Web Apps, Netlify, ein interner Webserver, SharePoint mit
   > direktem HTTPS-Zugriff auf die Datei).

3. Prüfen: Öffne `https://<dein-user>.github.io/<repo>/taskpane.html` im Browser –
   die Oberfläche muss laden.

---

## Schritt 2 (jeder Kollege, einmalig): Aktivieren

Gib den Kollegen diese **zwei Dateien** (z. B. per E-Mail/Teams/Netzlaufwerk):

- `manifest.prod.xml`
- `install-addin.cmd`

Der Kollege legt beide in **denselben Ordner** und macht einen **Doppelklick auf
`install-addin.cmd`**. Das Skript:

- registriert das Add-in für den Benutzer (Registry, **keine** Admin-Rechte nötig),
- leert den Office-Cache.

Danach **PowerPoint neu starten** → im **Start-Tab** erscheint die Gruppe **Gantt**
mit dem Button **Gantt Chart**.

**Kein Node, kein npm, kein Server** auf dem Kollegen-Rechner. Die Oberfläche kommt
live von deiner gehosteten URL.

### Deinstallieren
Doppelklick auf `uninstall-addin.cmd`, dann PowerPoint neu starten.

---

## Noch einfacher: PowerPoint im Web

Wenn Kollegen **PowerPoint im Web** (office.com) nutzen und eure Organisation den
Upload erlaubt:
**Einfügen → Add-Ins → Meine Add-Ins → Add-In hochladen → `manifest.prod.xml`.**
Kein Skript nötig.

---

## Der „richtige" Firmenweg (optional, für viele Nutzer)

Für eine große Verteilung ist der saubere Weg das **zentrale Deployment durch die
IT** (Microsoft 365 Admin Center → *Integrierte Apps* / *Add-Ins*). Dann taucht das
Add-in bei allen Nutzern **automatisch** auf – ganz ohne Skript. Dafür brauchst du
die gehostete URL (Schritt 1) und die Freigabe deiner IT.
