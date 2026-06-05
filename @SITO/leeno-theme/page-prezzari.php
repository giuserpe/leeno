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
                 WHERE cat_name = 'Listini' OR cat_name = 'Prezzari' 
                 LIMIT 1"
            );

            $parent_id = $listini_cat_id ? (int)$listini_cat_id : 0;

            // 1. Estraiamo tutte le categorie
            $all_cats = $wpdb->get_results("SELECT * FROM {$wpdb->prefix}wpfb_cats");
            $cats_by_id = array();
            if ( $all_cats ) {
                foreach ( $all_cats as $c ) {
                    $cats_by_id[$c->cat_id] = $c;
                }
            }

            // 2. Estraiamo tutti i file
            $all_files = $wpdb->get_results("SELECT * FROM {$wpdb->prefix}wpfb_files ORDER BY file_display_name ASC");

            if ( empty( $all_files ) ) : ?>
                <p class="prezzari-error">Nessun file trovato nella repository.</p>
            <?php else :

                // 3. Raggruppamento file: Regione -> Anno -> File[]
                $grouped = array();

                foreach ( $all_files as $file ) {
                    $cat_id = $file->file_category;
                    if ( ! isset( $cats_by_id[$cat_id] ) ) continue;

                    // Risale l'albero delle categorie per costruire il percorso
                    $path = array();
                    $curr = $cat_id;
                    $is_under_parent = false;

                    while ( $curr != 0 && isset( $cats_by_id[$curr] ) ) {
                        $path[] = $cats_by_id[$curr];
                        if ( $curr == $parent_id ) {
                            $is_under_parent = true;
                        }
                        $curr = $cats_by_id[$curr]->cat_parent;
                    }

                    // Se $parent_id è 0, tutto è valido (fallback)
                    if ( $parent_id == 0 ) {
                        $is_under_parent = true;
                    }

                    if ( ! $is_under_parent ) continue;

                    // Il percorso è costruito dal basso verso l'alto (es. Regione, Anno, Listini)
                    // Lo invertiamo per avere (Listini, Anno, Regione)
                    $path = array_reverse( $path );

                    // Trova l'indice di $parent_id nel percorso (se $parent_id non è 0)
                    $parent_idx = -1;
                    if ( $parent_id != 0 ) {
                        foreach ( $path as $i => $p ) {
                            if ( $p->cat_id == $parent_id ) {
                                $parent_idx = $i;
                                break;
                            }
                        }
                    }

                    // Estraiamo i primi due livelli sotto la cartella padre
                    $livello1 = isset( $path[$parent_idx + 1] ) ? $path[$parent_idx + 1]->cat_name : '';
                    $livello2 = isset( $path[$parent_idx + 2] ) ? $path[$parent_idx + 2]->cat_name : '';

                    // Rilevamento automatico di Regione e Anno basato sulla presenza di '20xx'
                    if ( $livello1 && $livello2 ) {
                        // Controlliamo se uno dei due livelli contiene l'anno
                        if ( preg_match('/\b20\d{2}\b/', $livello2) ) {
                            $year_name = $livello2;
                            $region_name = $livello1;
                        } elseif ( preg_match('/\b20\d{2}\b/', $livello1) ) {
                            $year_name = $livello1;
                            $region_name = $livello2;
                        } else {
                            // Fallback: assumiamo che il primo sia la regione e il secondo l'anno
                            $region_name = $livello1;
                            $year_name = $livello2;
                        }
                    } elseif ( $livello1 ) {
                        if ( preg_match('/\b20\d{2}\b/', $livello1) ) {
                            $year_name = $livello1;
                            $region_name = 'Nazionale/Generale';
                        } else {
                            $region_name = $livello1;
                            $year_name = 'Generale';
                        }
                    } else {
                        $region_name = 'Nazionale/Generale';
                        $year_name = 'Generale';
                    }

                    if ( ! isset( $grouped[$region_name] ) ) {
                        $grouped[$region_name] = array();
                    }
                    if ( ! isset( $grouped[$region_name][$year_name] ) ) {
                        $grouped[$region_name][$year_name] = array();
                    }
                    
                    $grouped[$region_name][$year_name][] = $file;
                }

                if ( empty( $grouped ) ) : ?>
                    <p class="prezzari-error">Nessun file presente nella categoria Prezzari/Listini.</p>
                <?php else :

                    // 4. Ordinamento: Regioni alfabetiche, Anni decrescenti
                    ksort( $grouped );
                    foreach ( $grouped as $region => $years ) {
                        krsort( $years ); // Anni dal più recente al più vecchio
                        $grouped[$region] = $years;
                    }

                    // 5. Render HTML
                    foreach ( $grouped as $region_name => $years ) :
                        $cat_slug = sanitize_title( $region_name );
                        ?>
                        <div class="prezzari-block prezzari-region" id="<?php echo esc_attr($cat_slug); ?>">
                            
                            <div class="prezzari-block-header">
                                <h2 class="prezzari-block-title" style="color: #ffffff !important;"><?php echo esc_html( $region_name ); ?></h2>
                                <?php 
                                $total_region_files = 0;
                                foreach ( $years as $fs ) {
                                    $total_region_files += count($fs);
                                }
                                ?>
                                <span class="prezzari-block-count" style="color: rgba(255,255,255,0.7);"><?php echo $total_region_files; ?> file</span>
                            </div>

                            <?php foreach ( $years as $year_name => $files ) : ?>
                                <div class="prezzari-block prezzari-subcat" style="margin-top: 1.5rem; padding-left: 1rem; border-left: 2px solid var(--border-color, #e2e8f0);">
                                    <div class="prezzari-block-header" style="margin-bottom: 1rem;">
                                        <h3 class="prezzari-block-title" style="font-size: 1.25rem; color: #ffffff !important;"><?php echo esc_html( $year_name ); ?></h3>
                                        <span class="prezzari-block-count"><?php echo count($files); ?> file</span>
                                    </div>

                                    <div class="leeno-table-wrap">
                                        <table class="leeno-table" role="table">
                                            <thead>
                                                <tr>
                                                    <th scope="col" class="col-name">Nome</th>
                                                    <th scope="col" class="col-dim" style="text-align: right; width: 100px;">Dim.</th>
                                                    <th scope="col" class="col-dl" style="width: 150px;"></th>
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
                                            ?>
                                                <tr class="leeno-row prezzari-row">
                                                    <td class="col-name">
                                                        <a href="<?php echo esc_url($dl_url); ?>">
                                                            <?php echo esc_html($name); ?>
                                                        </a>
                                                    </td>
                                                    <td class="col-dim" style="text-align: right;">
                                                        <?php echo esc_html($size); ?>
                                                    </td>
                                                    <td class="col-dl" style="text-align: right;">
                                                        <a href="<?php echo esc_url($dl_url); ?>" class="btn-leeno-download" aria-label="Scarica <?php echo esc_attr($name); ?>">
                                                            <svg width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="3" stroke-linecap="round" stroke-linejoin="round" aria-hidden="true">
                                                                <path d="M21 15v4a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2v-4"/>
                                                                <polyline points="7 10 12 15 17 10"/>
                                                                <line x1="12" y1="15" x2="12" y2="3"/>
                                                            </svg>
                                                            <span>Scarica</span>
                                                        </a>
                                                    </td>
                                                </tr>
                                            <?php endforeach; ?>
                                            </tbody>
                                        </table>
                                    </div>
                                </div>
                            <?php endforeach; ?>

                        </div><!-- .prezzari-block -->
                    <?php endforeach;
                endif; // end if grouped empty
            endif; // end if all_files empty
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
