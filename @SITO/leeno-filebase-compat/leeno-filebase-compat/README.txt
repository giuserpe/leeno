=== LeenO WP Filebase Compatibility ===
Contributors: LeenO Team
License: GPL2
License URI: https://www.gnu.org/licenses/gpl-2.0.html
Text Domain: leeno-filebase-compat
Domain Path: /languages

== Description ==

Plugin di compatibilità per interpretare gli shortcode [wpfilebase] legacy SENZA dipendere dal plugin WP Filebase (che è incompatibile con PHP 8+).

Fornisce una soluzione completa e moderna per la gestione di download, con styling responsivo, caching, e supporto per cartelle fisiche.

== Features ==

✓ Interpreta shortcode [wpfilebase tag=file] e [wpfilebase tag=list]
✓ Compatibile con PHP 8.1 e superiori
✓ Lettura dal DB di WP Filebase (se esiste) con fallback a cartelle fisiche
✓ Caching transient per performance
✓ Styling CSS responsive e tema-aware
✓ Supporto mobile, tablet, desktop
✓ Tema chiaro/scuro automatico
✓ Print-friendly
✓ Accessibilità ARIA

== Installation ==

1. Scarica il ZIP del plugin
2. Estrai in /wp-content/plugins/
3. Attiva da WP Admin › Plugin
4. Usa gli shortcode normalmente

== Shortcode Usage ==

=== File singolo ===
[wpfilebase tag=file id=1937]

Mostra un link di download per il file ID 1937.

=== Lista file ===
[wpfilebase tag=list id=38 sort=date order=DESC limit=20]

Opzioni:
- id: categoria di file (obbligatorio)
- sort: "name", "size", "date" (default: date)
- order: "ASC", "DESC" (default: DESC)
- limit: numero massimo di file (default: tutti)

== Compatibility ==

- PHP: 8.1+
- WordPress: 6.0+
- WP Filebase: opzionale (il plugin funziona anche senza)

== Fallback Logic ==

Se la categoria non esiste nel DB di WP Filebase, il plugin cerca automaticamente i file nella cartella:
/wp-content/uploads/filebase/LeenO/LeenO/

== Performance ==

- Caching transient: 1 ora per categoria
- Pulizia automatica daily
- Zero carico su ogni pagina (cache hit)

== Support ==

Per supporto, visita: https://leeno.org

== Changelog ==

= 1.1.0 =
- Aggiunto CSS styling responsive
- Aggiunto caching transient
- Aggiunto fallback a cartelle fisiche
- Migliorata accessibilità ARIA
- Tema chiaro/scuro automatico

= 1.0.0 =
- Versione iniziale
