# Deployment-Hinweise

Damit alle Zielserver die Anwendung ohne Browser-Caching ausliefern, müssen die Cache-Einstellungen für Apache und IIS parallel gepflegt werden:

- **Apache:** Die Datei `.htaccess` enthält die erforderlichen Header-Anweisungen für das Caching und muss bei Änderungen aktualisiert werden.
- **IIS:** Die Datei `web.config` liefert dieselben `Cache-Control`, `Pragma` und `Expires`-Header sowie deaktiviert den Client-Cache über `<staticContent>`. Änderungen an den Cache-Regeln sind hier ebenfalls einzutragen.

Nur wenn beide Konfigurationsdateien synchron gehalten werden, ist sichergestellt, dass alle unterstützten Webserver die Anwendung cache-frei bereitstellen.
