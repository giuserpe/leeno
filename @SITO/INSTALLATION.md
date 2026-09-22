# LeenO WP Filebase Compatibility - Guida di Installazione

## Versione
**1.1.0** | Compatibile con PHP 8.1+ | WordPress 6.0+

## Cosa fa questo plugin?

Interpreta gli shortcode `[wpfilebase]` legacy senza dipendere dal plugin WP Filebase (che è broken con PHP 8+).

Se hai pagine in cui usi:
```
[wpfilebase tag=file id=1937]
[wpfilebase tag=list id=38]
```

Questo plugin le farà funzionare di nuovo, senza errori.

## Installazione rapida

### Metodo 1: Via WP Admin (consigliato)
1. **Plugin › Aggiungi nuovo › Carica plugin**
2. Seleziona il file `leeno-filebase-compat.zip`
3. Clicca "Installa ora"
4. Clicca "Attiva plugin"

### Metodo 2: Via FTP
1. Estrai `leeno-filebase-compat.zip`
2. Carica la cartella `leeno-filebase-compat/` in `/wp-content/plugins/`
3. Attiva da WP Admin › Plugin

## Struttura del plugin

```
leeno-filebase-compat/
├── leeno-filebase-compat.php          # Entry point del plugin
├── includes/
│   └── class-leeno-filebase-compat.php # Logica principale
├── assets/
│   └── css/
│       └── leeno-filebase-compat.css  # Styling responsive
├── languages/
│   └── leeno-filebase-compat.pot      # Template traduzioni
├── README.md                          # Documentazione estesa
├── README.txt                         # Documentazione WordPress
├── LICENSE                            # Licenza GPL2
├── composer.json                      # Metadata Composer
├── .gitignore                         # Git ignore rules
└── build_zip.ps1                      # Script build PowerShell
```

## Utilizzo

### Link singolo file
```
[wpfilebase tag=file id=1937]
```

Renderizza un link di download per il file ID 1937.

### Tabella di file
```
[wpfilebase tag=list id=38]
[wpfilebase tag=list id=38 sort=name order=ASC limit=10]
```

#### Parametri:
- `id` (obbligatorio): ID categoria
- `sort`: "name" | "size" | "date" (default: date)
- `order`: "ASC" | "DESC" (default: DESC)
- `limit`: numero massimo file (default: tutti)

## Come funziona

### Lettura dati

Il plugin legge i file in questo ordine:

1. **DB di WP Filebase** (se tabelle `wp_wpfb_files` esistono)
2. **Fallback**: Cartella `/wp-content/uploads/filebase/LeenO/LeenO/`

Se la categoria non esiste nel DB, il plugin automaticamente legge dalla cartella fisica.

### Caching

- **Transient cache:** 1 ora per categoria
- **Pulizia automatica:** giornaliera (daily schedule)
- **Performance:** zero query su cache hit

## Features

✓ PHP 8.1+ compatibile  
✓ Styling responsive (mobile, tablet, desktop)  
✓ Tema chiaro/scuro automatico  
✓ Print-friendly  
✓ Accessibilità ARIA + semantica HTML5  
✓ Caching intelligente  
✓ Fallback cartelle fisiche  

## Compatibilità

| Ambiente | Richiesto | Supportato |
|----------|-----------|-----------|
| **PHP** | 8.1+ | 8.1, 8.2, 8.3, 8.4 |
| **WordPress** | 6.0+ | 6.0, 6.1, 6.2, 6.3, 6.4, 6.5 |
| **WP Filebase** | No (opzionale) | Se esiste, legge il DB |

## Troubleshooting

### Non vedo la lista file
**Possibile causa:** Categoria non esiste nel DB

**Soluzione:** Il plugin cercherà nella cartella `/wp-content/uploads/filebase/LeenO/LeenO/`

### Gli link non funzionano
**Possibile causa:** Percorso file errato

**Soluzione:** Verifica che i file siano in:
```
/wp-content/uploads/filebase/LeenO/LeenO/
```

### Errori PHP
**Possibile causa:** PHP < 8.1

**Soluzione:** Aggiorna a PHP 8.1 minimo

## Sviluppo

### Build del plugin

Se usi PowerShell su Windows:
```powershell
cd leeno-filebase-compat
./build_zip.ps1
```

Crea `leeno-filebase-compat.zip` in `build/`

### Struttura codice

```php
LeenO_Filebase_Compat::handle_shortcode($atts)
  ├── render_single_file($file_id)
  └── render_file_list($category_id)
      ├── get_files_by_category() [con caching]
      └── format_bytes()
```

## Licenza

GPL2 - Vedi file `LICENSE`

## Supporto

Per supporto e segnalazioni bug: https://leeno.org

---

**Versione:** 1.1.0  
**Data:** Settembre 2026  
**Autore:** LeenO Team
