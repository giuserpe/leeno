<?php
/**
 * Template Name: Brogliaccio (presentazione)
 *
 * Pagina di presentazione di Brogliaccio, l'agenda di cantiere di LeenO.
 * Si applica in automatico alla pagina con slug "brogliaccio"; in alternativa
 * si sceglie dal selettore "Modello" nell'editor della pagina.
 * Il testo inserito nell'editor della pagina, se presente, compare in coda.
 */

get_header();

$url_app      = 'https://brogliaccio.leeno.org';
$url_download = home_url( '/scarica-leeno/' );
$url_docs     = home_url( '/category/documentazione/' );
$url_forum    = home_url( '/forums/' );
?>

<style>
    .brogliaccio-lead { font-size: 1.25rem; line-height: 1.7; max-width: 62ch; color: var(--text-dark); margin-bottom: 28px; }
    .brogliaccio-ctas { display: flex; flex-wrap: wrap; gap: 16px; align-items: center; margin-bottom: 64px; }
    .brogliaccio-ctas .btn-big, .brogliaccio-ctas .btn-outline { text-decoration: none !important; }
    .brogliaccio-section { margin-bottom: 72px; }
    .brogliaccio-section h2 { text-transform: uppercase; font-size: clamp(1.4rem, 3vw, 2rem); margin-bottom: 28px; color: var(--text-dark); }
    .brogliaccio-grid { display: grid; grid-template-columns: repeat(auto-fit, minmax(260px, 1fr)); gap: 24px; }
    .brogliaccio-steps { list-style: none; counter-reset: passo; padding: 0; margin: 0; display: grid; gap: 20px; }
    .brogliaccio-steps li { counter-increment: passo; position: relative; background: #fff; padding: 24px 24px 24px 88px; border-left: 4px solid var(--accent-cyan); }
    .brogliaccio-steps li::before { content: counter(passo); position: absolute; left: 24px; top: 20px; width: 44px; height: 44px; background: var(--bg-dark); color: var(--accent-cyan); font-family: var(--font-display); font-weight: 700; font-size: 1.3rem; display: flex; align-items: center; justify-content: center; }
    .brogliaccio-steps strong { display: block; font-family: var(--font-display); text-transform: uppercase; margin-bottom: 6px; color: var(--text-dark); }
    .brogliaccio-steps p { margin: 0; color: var(--text-dark); line-height: 1.7; }
    .brogliaccio-menu { font-family: var(--font-mono); font-size: 0.9em; background: var(--bg-primary); padding: 2px 6px; }
    .brogliaccio-note { background: var(--bg-dark); color: #fff; padding: 32px; border-left: 4px solid var(--accent-cyan); }
    .brogliaccio-note h2 { color: #fff; margin-bottom: 16px; }
    .brogliaccio-note p { color: #d5dbe5; line-height: 1.7; margin-bottom: 12px; }
    .brogliaccio-note p:last-child { margin-bottom: 0; }
    .brogliaccio-campi { columns: 2 280px; column-gap: 32px; padding-left: 20px; line-height: 1.9; color: var(--text-dark); }
    @media (max-width: 600px) { .brogliaccio-steps li { padding-left: 24px; padding-top: 80px; } }
</style>

<main id="main-content" class="main-content page-content page-brogliaccio">

    <div class="page-header">
        <div class="container">
            <nav class="breadcrumbs" aria-label="Percorso di navigazione">
                <a href="<?php echo esc_url( home_url( '/' ) ); ?>">Home</a>
                <span class="sep" aria-hidden="true">&rsaquo;</span>
                <span class="current">Brogliaccio</span>
            </nav>
            <h1 class="page-title">Brogliaccio</h1>
        </div>
    </div>

    <div class="container">

        <p class="brogliaccio-lead">
            <strong>Brogliaccio</strong> è l'agenda di cantiere di LeenO: un'applicazione gratuita per annotare sul telefono,
            giorno per giorno, i fatti da riportare nel Giornale dei Lavori. Poi importi tutto in LeenO.
        </p>

        <div class="brogliaccio-ctas">
            <a class="btn-big" href="<?php echo esc_url( $url_app ); ?>" rel="noopener">Apri Brogliaccio</a>
            <!-- <a class="btn-outline" href="<?php echo esc_url( $url_docs ); ?>">Leggi la documentazione</a> -->
        </div>

        <section class="brogliaccio-section" aria-labelledby="brogliaccio-perche">
            <h2 id="brogliaccio-perche">Cosa ti dà</h2>
            <div class="brogliaccio-grid">

                <div class="feature-card">
                    <div class="feature-icon">
                        <svg width="24" height="24" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" aria-hidden="true">
                            <rect x="5" y="2" width="14" height="20" rx="2" /><line x1="12" y1="18" x2="12.01" y2="18" />
                        </svg>
                    </div>
                    <h3 class="feature-title">Nessuna installazione</h3>
                    <p class="feature-desc">Si apre dal browser del telefono, senza store e senza account. Se vuoi, la aggiungi alla schermata Home come un'app.</p>
                </div>

                <div class="feature-card">
                    <div class="feature-icon">
                        <svg width="24" height="24" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" aria-hidden="true">
                            <path d="M1 1l22 22" /><path d="M16.72 11.06A10.94 10.94 0 0 1 19 12.55" /><path d="M5 12.55a10.94 10.94 0 0 1 5.17-2.39" /><path d="M10.71 5.05A16 16 0 0 1 22.58 9" /><path d="M1.42 9a15.91 15.91 0 0 1 4.7-2.88" /><path d="M8.53 16.11a6 6 0 0 1 6.95 0" /><line x1="12" y1="20" x2="12.01" y2="20" />
                        </svg>
                    </div>
                    <h3 class="feature-title">Anche senza campo</h3>
                    <p class="feature-desc">Dopo il primo caricamento funziona senza connessione: in cantiere non serve la rete per scrivere.</p>
                </div>

                <div class="feature-card">
                    <div class="feature-icon">
                        <svg width="24" height="24" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" aria-hidden="true">
                            <path d="M23 19a2 2 0 0 1-2 2H3a2 2 0 0 1-2-2V8a2 2 0 0 1 2-2h4l2-3h6l2 3h4a2 2 0 0 1 2 2z" /><circle cx="12" cy="13" r="4" />
                        </svg>
                    </div>
                    <h3 class="feature-title">Foto del giorno</h3>
                    <p class="feature-desc">Scatti o scegli dalla galleria. Le foto vengono ridimensionate e ripulite da tutti i metadati, posizione GPS compresa.</p>
                </div>

                <div class="feature-card">
                    <div class="feature-icon">
                        <svg width="24" height="24" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" aria-hidden="true">
                            <path d="M14 2H6a2 2 0 0 0-2 2v16a2 2 0 0 0 2 2h12a2 2 0 0 0 2-2V8z" /><polyline points="14 2 14 8 20 8" /><line x1="16" y1="13" x2="8" y2="13" /><line x1="16" y1="17" x2="8" y2="17" />
                        </svg>
                    </div>
                    <h3 class="feature-title">Gli stessi campi del giornale</h3>
                    <p class="feature-desc">La scheda del giorno segue l'ordine del Giornale dei Lavori di LeenO. Compili solo ciò che serve, il salvataggio è automatico.</p>
                </div>

                <div class="feature-card">
                    <div class="feature-icon">
                        <svg width="24" height="24" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" aria-hidden="true">
                            <polygon points="12 2 2 7 12 12 22 7 12 2" /><polyline points="2 17 12 22 22 17" /><polyline points="2 12 12 17 22 12" />
                        </svg>
                    </div>
                    <h3 class="feature-title">Più cantieri</h3>
                    <p class="feature-desc">Ogni cantiere ha il suo elenco di giornate, da scegliere con un menù. Per ognuno l'app ti dice quante giornate non hai ancora esportato e quando hai esportato l'ultima volta.</p>
                </div>

                <div class="feature-card">
                    <div class="feature-icon">
                        <svg width="24" height="24" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" aria-hidden="true">
                            <path d="M21 15v4a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2v-4" /><polyline points="7 10 12 15 17 10" /><line x1="12" y1="15" x2="12" y2="3" />
                        </svg>
                    </div>
                    <h3 class="feature-title">Un file per LeenO</h3>
                    <p class="feature-desc">Esporti un cantiere in un solo file: un .json, oppure uno .zip con le foto se ne hai allegate. Con più cantieri puoi esportarli tutti insieme in un unico .zip.</p>
                </div>

                <div class="feature-card">
                    <div class="feature-icon">
                        <svg width="24" height="24" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" aria-hidden="true">
                            <polyline points="1 4 1 10 7 10" /><path d="M3.51 15a9 9 0 1 0 2.13-9.36L1 10" />
                        </svg>
                    </div>
                    <h3 class="feature-title">Ripristina da file</h3>
                    <p class="feature-desc">Hai cambiato telefono o perso i dati? Da un file esportato riporti le giornate e, con lo .zip, anche cantieri e foto. Ripetere il ripristino non duplica le foto.</p>
                </div>

                <div class="feature-card">
                    <div class="feature-icon">
                        <svg width="24" height="24" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" aria-hidden="true">
                            <polyline points="6 9 6 2 18 2 18 9" /><path d="M6 18H4a2 2 0 0 1-2-2v-5a2 2 0 0 1 2-2h16a2 2 0 0 1 2 2v5a2 2 0 0 1-2 2h-2" /><rect x="6" y="14" width="12" height="8" />
                        </svg>
                    </div>
                    <h3 class="feature-title">Stampa o PDF</h3>
                    <p class="feature-desc">Apre la stampa del telefono o del browser: stampi oppure scegli «Salva come PDF». Compaiono solo i campi compilati e le foto, con la dicitura che non è un registro ufficiale.</p>
                </div>

            </div>
        </section>

        <section class="brogliaccio-section" aria-labelledby="brogliaccio-come">
            <h2 id="brogliaccio-come">Come funziona</h2>
            <ol class="brogliaccio-steps">
                <li>
                    <strong>Annota in cantiere</strong>
                    <p>Crea un cantiere, scegli la data e compila la giornata: meteo, presenti, attività svolte, operai, attrezzature, ordini di servizio, osservazioni e gli altri campi del giornale. Aggiungi le foto.</p>
                </li>
                <li>
                    <strong>Esporta</strong>
                    <p>Con <em>Esporta per LeenO</em> ottieni il file dell'agenda: un .json oppure, se hai allegato foto, uno .zip. Il .json puoi condividerlo (posta, messaggi, cloud) per portarlo al computer; lo .zip viene salvato sul telefono, di solito nella cartella Download. Con più cantieri l'app ti chiede se esportare solo quello aperto o tutti.</p>
                </li>
                <li>
                    <strong>Importa in LeenO</strong>
                    <p>In un Giornale Lavori aperto scegli <span class="brogliaccio-menu">LeenO &gt; Importa/Esporta... &gt; Importa agenda di cantiere nel Giornale Lavori...</span> e seleziona il file (.json o .zip). Le giornate nuove vengono aggiunte; per quelle già presenti decidi tu se sovrascrivere. Le foto finiscono nella cartella FOTO accanto al documento, con un collegamento nel giornale: con uno .zip il documento deve essere già stato salvato. L'importazione legge un cantiere per volta, quindi uno .zip con tutti i cantieri non viene accettato.</p>
                </li>
            </ol>
        </section>

        <section class="brogliaccio-section" aria-labelledby="brogliaccio-campi">
            <h2 id="brogliaccio-campi">I campi di ogni giornata</h2>
            <ul class="brogliaccio-campi">
                <li>Meteo</li>
                <li>Presenti/intervenuti</li>
                <li>Annotazioni, attività svolte</li>
                <li>Qualifica e n. operai</li>
                <li>Attrezzature impiegate</li>
                <li>Provviste</li>
                <li>Rifiuto di materiali e/o manufatti</li>
                <li>Disposizioni e ordini di servizio del R.U.P. e del D.L.</li>
                <li>Relazione indirizzata al R.U.P.</li>
                <li>Verbali di accertamento e prove</li>
                <li>Contestazioni, sospensioni e riprese lavori</li>
                <li>Varianti disposte, modifiche e/o aggiunte prezzi</li>
                <li>Evento infortunistico</li>
                <li>Osservazioni, prescrizioni, avvertenze della D.L.</li>
            </ul>
        </section>

        <section class="brogliaccio-section brogliaccio-note" aria-labelledby="brogliaccio-dati">
            <h2 id="brogliaccio-dati">I tuoi dati restano tuoi</h2>
            <p>I dati dell'agenda restano nella memoria del browser del telefono: non esiste alcun server e non serve alcun account. Per questo conviene esportare spesso: se cancelli i dati del browser o cambi telefono, ciò che non hai esportato va perso. Ciò che hai esportato lo recuperi con «Ripristina da file».</p>
            <p>Brogliaccio non è un registro ufficiale. Raccoglie annotazioni da consolidare nel Giornale dei Lavori di LeenO, che resta il documento di riferimento.</p>
            <p>Il campo Evento infortunistico può contenere dati sulla salute di una persona identificabile: prima di condividere un file che lo contiene, l'app ti chiede conferma.</p>
        </section>

        <?php
        // Testo opzionale scritto nell'editor della pagina.
        while ( have_posts() ) :
            the_post();
            $extra = get_the_content();
            if ( trim( wp_strip_all_tags( $extra ) ) ) : ?>
                <div class="entry-content brogliaccio-section">
                    <?php echo apply_filters( 'the_content', $extra ); ?>
                </div>
            <?php endif;
        endwhile;
        ?>

        <p class="brogliaccio-section">
            Hai bisogno di una mano o vuoi segnalare qualcosa? Scrivi nel <a href="<?php echo esc_url( $url_forum ); ?>">forum</a>.
            Ti serve LeenO? Lo trovi nella pagina di <a href="<?php echo esc_url( $url_download ); ?>">download</a>.
        </p>

    </div><!-- .container -->

</main>

<?php get_footer(); ?>
