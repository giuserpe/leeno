<?php
/**
 * Template Name: Download Software
 *
 * Pagina download software — lista piatta dei file della categoria.
 * Sostituisce lo shortcode: [wpfilebase tag=list id='38' sort=file_name tpl=data-table]
 */

get_header();
?>

<main id="main-content" class="main-content page-content page-software">

    <div class="page-header">
        <div class="container">
            <nav class="breadcrumbs" aria-label="Percorso di navigazione">
                <a href="<?php echo esc_url( home_url('/') ); ?>">Home</a>
                <span class="sep" aria-hidden="true">›</span>
                <span class="current"><?php the_title(); ?></span>
            </nav>
            <h1 class="page-title"><?php the_title(); ?></h1>
            <?php
            $desc = get_the_content();
            $intro = '';
            $iframe_content = '';
            
            $pos = strpos( $desc, '[advanced_iframe' );
            if ( $pos !== false ) {
                $intro = substr( $desc, 0, $pos );
                $iframe_content = substr( $desc, $pos );
            } else {
                $intro = $desc;
            }

            if ( trim( strip_tags( $intro ) ) || trim( $intro ) ) : ?>
            <div class="page-desc software-intro">
                <?php echo apply_filters( 'the_content', $intro ); ?>
            </div>
            <?php endif; ?>
        </div>
    </div>

    <div class="container prezzari-container">
        <?php
        // ── Carica WP Filebase se non già caricato ──────────────────────────
        if ( ! class_exists('WPFB_Core') ) {
            $wpfb_path = WP_PLUGIN_DIR . '/wp-filebase/wp-filebase.php';
            if ( file_exists( $wpfb_path ) ) {
                include_once $wpfb_path;
            }
        }

        if ( ! class_exists('WPFB_Core') ) : ?>
            <p class="prezzari-error">Plugin WP Filebase non attivo.</p>
        <?php else :

            global $wpdb;

            // ── Recupera la categoria desiderata (Software = 38) ──────────────────────
            $cat_id = 38;

            $software_cat = $wpdb->get_row( $wpdb->prepare(
                "SELECT * FROM {$wpdb->prefix}wpfb_cats WHERE cat_id = %d",
                $cat_id
            ) );

            if ( empty( $software_cat ) ) : ?>
                <p class="prezzari-error">Categoria software non trovata.</p>
            <?php else :

                // Recupera solo i file diretti in questa categoria
                $files = $wpdb->get_results( $wpdb->prepare(
                    "SELECT * FROM {$wpdb->prefix}wpfb_files
                     WHERE file_category = %d
                     ORDER BY file_name ASC",
                    $cat_id
                ) );

                if ( empty( $files ) ) : ?>
                    <p class="prezzari-error">Nessun file trovato in questa categoria.</p>
                <?php else : ?>
                    <div class="prezzari-block prezzari-region">
                        <table class="prezzari-table" role="table">
                            <thead>
                                <tr>
                                    <th scope="col">Nome</th>
                                    <th scope="col" class="col-size">Dim.</th>
                                    <th scope="col" class="col-hits">↓</th>
                                    <th scope="col" class="col-dl"></th>
                                </tr>
                            </thead>
                            <tbody>
                            <?php foreach ( $files as $file ) :
                                $dl_url   = home_url( '?wpfb_dl=' . $file->file_id );
                                $name     = $file->file_display_name ?: $file->file_name;
                                $size     = size_format( $file->file_size, 1 );
                                $hits     = intval( $file->file_hits );
                            ?>
                                <tr class="prezzari-row">
                                    <td class="col-name">
                                        <a href="<?php echo esc_url($dl_url); ?>" class="prezzari-file-link">
                                            <?php echo esc_html($name); ?>
                                        </a>
                                    </td>
                                    <td class="col-size">
                                        <span class="prezzari-size"><?php echo esc_html($size); ?></span>
                                    </td>
                                    <td class="col-hits">
                                        <span class="prezzari-hits" title="Download effettuati"><?php echo number_format_i18n($hits); ?></span>
                                    </td>
                                    <td class="col-dl">
                                        <a href="<?php echo esc_url($dl_url); ?>" class="prezzari-dl-btn" aria-label="Scarica <?php echo esc_attr($name); ?>">
                                            <svg width="12" height="12" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.5" aria-hidden="true">
                                                <path d="M21 15v4a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2v-4"/>
                                                <polyline points="7 10 12 15 17 10"/>
                                                <line x1="12" y1="15" x2="12" y2="3"/>
                                            </svg>
                                            Scarica
                                        </a>
                                    </td>
                                </tr>
                            <?php endforeach; ?>
                            </tbody>
                        </table>
                    </div>
                <?php endif;

            endif; // software_cat
        endif; // WPFB_Core
        ?>
    </div><!-- .prezzari-container -->

    <?php if ( ! empty( $iframe_content ) && trim( $iframe_content ) ) : ?>
    <div class="container software-after-table" style="margin-top: 2rem; margin-bottom: 3rem;">
        <?php echo apply_filters( 'the_content', $iframe_content ); ?>
    </div>
    <?php endif; ?>

</main>

<?php get_footer(); ?>
