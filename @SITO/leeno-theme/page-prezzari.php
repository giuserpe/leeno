<?php
/**
 * Template Name: Prezzari
 *
 * Pagina download prezzari — lista piatta raggruppata per regione/categoria.
 * Usa direttamente l'API di WP Filebase (WPFB) senza shortcode browser.
 * Assegna questo template alla pagina Prezzari da: Pagina → Attributi → Template.
 */

get_header();
?>

<main id="main-content" class="main-content page-content page-prezzari">

    <div class="page-header">
        <div class="container">
            <nav class="breadcrumbs" aria-label="Percorso di navigazione">
                <a href="<?php echo esc_url( home_url('/') ); ?>">Home</a>
                <span class="sep" aria-hidden="true">›</span>
                <span class="current"><?php the_title(); ?></span>
            </nav>
            <h1 class="page-title"><?php the_title(); ?></h1>
            <div class="prezzari-search-bar">
                <input
                    type="text"
                    id="prezzari-filter"
                    placeholder="Filtra per nome, regione o anno…"
                    aria-label="Filtra prezzari"
                    autocomplete="off"
                >
            </div>
        </div>
    </div>

    <div class="container prezzari-container">
        <div class="content-layout">
            <div class="content-main">

            <?php
            $desc = get_the_content();
            if ( $desc ) : ?>
            <div class="prezzari-intro entry-content" style="margin-bottom:2rem;">
                <?php echo wp_kses_post( $desc ); ?>
            </div>
            <?php endif; ?>

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

            // Cerchiamo l'ID della categoria principale dei listini
            $listini_cat_id = $wpdb->get_var(
                "SELECT cat_id FROM {$wpdb->prefix}wpfb_cats 
                 WHERE cat_name = 'LISTINI' OR cat_name = 'Prezzari' 
                 LIMIT 1"
            );

            // Se non troviamo la categoria specifica, mostriamo tutto come prima (fallback)
            $parent_id = $listini_cat_id ? (int)$listini_cat_id : 0;

            $top_cats = $wpdb->get_results( $wpdb->prepare(
                "SELECT * FROM {$wpdb->prefix}wpfb_cats
                 WHERE cat_parent = %d
                 ORDER BY cat_name ASC",
                $parent_id
            ) );

            if ( empty( $top_cats ) && $parent_id === 0 ) {
                // Fallback: se siamo alla radice e non c'è nulla, prendi tutto
                $top_cats = $wpdb->get_results(
                    "SELECT * FROM {$wpdb->prefix}wpfb_cats
                     ORDER BY cat_name ASC"
                );
            }

            if ( empty( $top_cats ) ) : ?>
                <p class="prezzari-error">Nessuna categoria trovata nella repository.</p>
            <?php else :

                // Funzione ricorsiva per renderizzare categorie e file
                function leeno_render_cat( $cat, $depth = 0 ) {
                    global $wpdb;

                    $prefix = $depth === 0 ? 'prezzari-region' : 'prezzari-subcat';

                    // File diretti in questa categoria
                    $files = $wpdb->get_results( $wpdb->prepare(
                        "SELECT * FROM {$wpdb->prefix}wpfb_files
                         WHERE file_category = %d
                         ORDER BY file_display_name ASC",
                        $cat->cat_id
                    ) );

                    // Sottocategorie
                    $subcats = $wpdb->get_results( $wpdb->prepare(
                        "SELECT * FROM {$wpdb->prefix}wpfb_cats
                         WHERE cat_parent = %d
                         ORDER BY cat_name ASC",
                        $cat->cat_id
                    ) );

                    // Non mostrare categorie vuote (né file né subcats)
                    $has_content = ! empty( $files ) || ! empty( $subcats );
                    if ( ! $has_content ) return;

                    $cat_slug = sanitize_title( $cat->cat_name );
                    ?>
                    <div class="prezzari-block <?php echo esc_attr($prefix); ?>" id="<?php echo esc_attr($cat_slug); ?>">

                        <div class="prezzari-block-header">
                            <h2 class="prezzari-block-title"><?php echo esc_html( $cat->cat_name ); ?></h2>
                            <?php if ( ! empty( $files ) ) : ?>
                            <span class="prezzari-block-count"><?php echo count($files); ?> file</span>
                            <?php endif; ?>
                        </div>

                        <?php if ( ! empty( $files ) ) : ?>
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
                        $dl_url = '';
                                if ( method_exists('WPFB_Core', 'GetUrl') ) {
                                    $dl_url = WPFB_Core::GetUrl( $file );
                                } elseif ( isset($file->file_url) ) {
                                    $dl_url = $file->file_url;
                                } else {
                                    $dl_url = home_url( '?wpfb_dl=' . $file->file_id );
                                }
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
                        <?php endif; ?>

                        <?php
                        // Sottocategorie ricorsive (es. anno dentro regione)
                        if ( ! empty( $subcats ) ) :
                            foreach ( $subcats as $subcat ) :
                                leeno_render_cat( $subcat, $depth + 1 );
                            endforeach;
                        endif;
                        ?>

                    </div><!-- .prezzari-block -->
                    <?php
                }

                // Render di tutte le top-level categories
                foreach ( $top_cats as $cat ) :
                    leeno_render_cat( $cat, 0 );
                endforeach;

            endif; // top_cats
        endif; // WPFB_Core
        ?>
            </div><!-- .content-main -->

            <?php if ( is_active_sidebar('sidebar-blog') ) : ?>
            <aside class="content-sidebar">
                <?php dynamic_sidebar('sidebar-blog'); ?>
            </aside>
            <?php endif; ?>

        </div><!-- .content-layout -->
    </div><!-- .prezzari-container -->

</main>

<?php get_footer(); ?>
