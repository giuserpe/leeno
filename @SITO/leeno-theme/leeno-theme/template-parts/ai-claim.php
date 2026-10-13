<?php
/**
 * Dichiarazione sull'uso dell'intelligenza artificiale.
 *
 * Unica fonte del testo: si modifica solo qui, mai nei template che la includono.
 * Perimetro: riguarda il comportamento del software, non il modo in cui viene sviluppato.
 *
 * Varianti:
 *   'full'    sfondo chiaro, pagina download (testo completo, ancora #dichiarazione-ia)
 *   'compact' sfondo scuro, fascia sotto l'hero della home (frase breve + link alla versione completa)
 *
 * Uso:
 *   get_template_part( 'template-parts/ai-claim', null, array( 'variant' => 'compact' ) );
 *
 * Data di revisione: costante LEENO_AI_CLAIM_REVISIONE in functions.php, da rivedere a ogni rilascio.
 */

if ( ! defined( 'ABSPATH' ) ) exit;

$variant = ( isset( $args['variant'] ) && 'compact' === $args['variant'] ) ? 'compact' : 'full';
?>
<?php if ( 'compact' === $variant ) : ?>
<div class="ai-free-strip"><div class="container">
<?php endif; ?>
<div class="ai-free ai-free--<?php echo esc_attr( $variant ); ?>"<?php echo ( 'full' === $variant ) ? ' id="dichiarazione-ia"' : ''; ?> role="note" aria-label="Dichiarazione sull'intelligenza artificiale">
    <div class="ai-free__icon" aria-hidden="true">
        <svg width="24" height="24" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><rect x="6" y="6" width="12" height="12"/><path d="M9 2v4M15 2v4M9 18v4M15 18v4M2 9h4M2 15h4M18 9h4M18 15h4"/><path d="M4 20L20 4"/></svg>
    </div>
    <div class="ai-free__body">
        <p class="ai-free__title">Nessuna IA nel software</p>
<?php if ( 'compact' === $variant ) : ?>
        <p class="ai-free__text">Le macro non usano servizi di intelligenza artificiale e non inviano i contenuti dei tuoi computi all'esterno.</p>
        <a class="ai-free__link" href="<?php echo esc_url( home_url( '/scarica-leeno/#dichiarazione-ia' ) ); ?>">Perimetro e connessioni</a>
<?php else : ?>
        <p class="ai-free__text">Le macro del software non usano servizi di intelligenza artificiale e non inviano i contenuti dei tuoi computi a servizi esterni. Il software si collega a internet solo per due funzioni: il controllo degli aggiornamenti (leeno.org) e il download opzionale degli indici ISTAT TOL.</p>
        <p class="ai-free__note">Riferito al comportamento del software &middot; Aggiornato a <?php echo esc_html( LEENO_AI_CLAIM_REVISIONE ); ?></p>
<?php endif; ?>
    </div>
</div>
<?php if ( 'compact' === $variant ) : ?>
</div></div>
<?php endif; ?>
