<?php
/**
 * Template Name: Archivio Download
 *
 * Lista file WP Filebase ordinata per data decrescente (più recente prima).
 * Repository id=199 — Archivio versioni LeenO.
 * Assegna questo template alla pagina: Archivio › Download › LeenO › About LeenO
 */

get_header();
?>

<main id="main-content" class="main-content page-content page-archivio">

    <div class="page-header">
        <div class="container">
            <nav class="breadcrumbs" aria-label="Percorso di navigazione">
                <a href="<?php echo esc_url( home_url('/') ); ?>">Home</a>
                <span class="sep" aria-hidden="true">›</span>
                <a href="<?php echo esc_url( home_url('/about-leeno/') ); ?>">About LeenO</a>
                <span class="sep" aria-hidden="true">›</span>
                <a href="<?php echo esc_url( home_url('/about-leeno/leeno/download/') ); ?>">Download</a>
                <span class="sep" aria-hidden="true">›</span>
                <span class="current"><?php the_title(); ?></span>
            </nav>
            <h1 class="page-title"><?php the_title(); ?></h1>
        </div>
    </div>

    <div class="container archivio-container">
        <div class="content-layout">
            <div class="content-main">

            <?php
            $content = get_the_content();
            if ( $content ) : ?>
            <div class="archivio-intro entry-content" style="margin-bottom:2rem;">
                <?php echo wp_kses_post( $content ); ?>
            </div>
            <?php endif; ?>

            <?php
        if ( ! class_exists('WPFB_Core') ) : ?>
            <p class="prezzari-error">Plugin WP Filebase non attivo.</p>
        <?php else :
            global $wpdb;

            // Recupera tutti i file della repository 199, ordinati per data decrescente
            $files = $wpdb->get_results(
                "SELECT f.*, c.cat_name
                 FROM {$wpdb->prefix}wpfb_files f
                 LEFT JOIN {$wpdb->prefix}wpfb_cats c ON f.file_category = c.cat_id
                 WHERE f.file_repository = 199
                    OR f.file_category IN (
                        SELECT cat_id FROM {$wpdb->prefix}wpfb_cats
                        WHERE cat_id = 199 OR cat_parent = 199
                    )
                 ORDER BY f.file_date DESC, f.file_display_name ASC"
            );

            // Fallback: cerca per cat_id = 199 direttamente
            if ( empty( $files ) ) {
                $files = $wpdb->get_results(
                    "SELECT f.*, c.cat_name
                     FROM {$wpdb->prefix}wpfb_files f
                     LEFT JOIN {$wpdb->prefix}wpfb_cats c ON f.file_category = c.cat_id
                     WHERE f.file_category = 199
                     ORDER BY f.file_date DESC, f.file_display_name ASC"
                );
            }

            if ( empty( $files ) ) : ?>
                <p class="prezzari-error">Nessun file trovato nella repository.</p>

                <?php if ( current_user_can('administrator') ) :
                    // Diagnostica per admin
                    $sample = $wpdb->get_results(
                        "SELECT file_id, file_display_name, file_category, file_repository, file_date
                         FROM {$wpdb->prefix}wpfb_files
                         ORDER BY file_date DESC LIMIT 10"
                    );
                    echo '<div style="background:#1a2010;color:#aad400;font-family:monospace;font-size:11px;padding:12px 20px;margin:16px 0;border-left:4px solid #aad400">';
                    echo '<strong>DEBUG — ultimi 10 file nel DB:</strong><br>';
                    foreach ( $sample as $f ) {
                        echo "ID={$f->file_id} | cat={$f->file_category} | repo={$f->file_repository} | data={$f->file_date} | " . esc_html($f->file_display_name) . "<br>";
                    }
                    echo '</div>';
                endif;

            else :

                $total = count( $files );
            ?>

            <div class="archivio-header">
                <span class="archivio-count">
                    <?php printf( _n('%s versione disponibile', '%s versioni disponibili', $total, 'leeno-dm'), number_format_i18n($total) ); ?>
                </span>
                <span class="archivio-sort-label">Ordinate per data — più recente prima</span>
            </div>

            <div class="leeno-table-wrap">
                <table class="leeno-table" role="table">
                    <thead>
                        <tr>
                            <th scope="col">Versione / File</th>
                            <th scope="col" style="text-align: right; width: 100px;">Data</th>
                            <th scope="col" style="text-align: right; width: 100px;">Dim.</th>
                            <th scope="col" style="text-align: right; width: 60px;">&darr;</th>
                            <th scope="col" style="width: 150px;"></th>
                        </tr>
                    </thead>
                    <tbody>
                    <?php foreach ( $files as $file ) :
                        $dl_url  = '';
                        if ( method_exists('WPFB_Core', 'GetUrl') ) {
                            $dl_url = WPFB_Core::GetUrl( $file );
                        } elseif ( isset($file->file_url) ) {
                            $dl_url = $file->file_url;
                        } elseif ( isset($file->file_id) ) {
                            $dl_url = home_url( '/wp-content/plugins/wp-filebase/download.php?id=' . intval($file->file_id) );
                        }
                        $name    = $file->file_display_name ?: $file->file_name;
                        $size    = size_format( $file->file_size, 1 );
                        $hits    = intval( $file->file_hits );
                        $date    = $file->file_date ? date_i18n( 'd M Y', strtotime($file->file_date) ) : '—';
                        // Usa file_version se disponibile, altrimenti estrai dal nome
                        $version = ! empty( $file->file_version )
                            ? $file->file_version
                            : ( preg_match('/(\d+\.\d+[\.\d]*)/u', $name, $m) ? $m[1] : null );
                    ?>
                        <tr class="leeno-row">
                            <td class="col-name">
                                <a href="<?php echo esc_url($dl_url); ?>">
                                    <?php if ( $version ) : ?>
                                    <span class="archivio-ver-badge" style="background: var(--bg-dark); color: var(--accent-cyan); padding: 2px 6px; font-size: 0.75rem; margin-right: 8px; font-family: var(--font-mono);">v<?php echo esc_html($version); ?></span>
                                    <?php endif; ?>
                                    <?php echo esc_html($name); ?>
                                </a>
                            </td>
                            <td class="col-extra" style="color: #666;">
                                <?php echo esc_html($date); ?>
                            </td>
                            <td class="col-dim">
                                <?php echo esc_html($size); ?>
                            </td>
                            <td class="col-extra">
                                <span title="Download effettuati"><?php echo number_format_i18n($hits); ?></span>
                            </td>
                            <td style="text-align: right;">
                                <a href="<?php echo esc_url($dl_url); ?>" class="btn-leeno-download" aria-label="Scarica <?php echo esc_attr($name); ?>">
                                    <svg width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="3" stroke-linecap="round" stroke-linejoin="round" aria-hidden="true">
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

            <?php endif; // files
        endif; // WPFB_Core
        ?>
            </div><!-- .content-main -->

            <?php if ( is_active_sidebar('sidebar-blog') ) : ?>
            <aside class="content-sidebar">
                <?php dynamic_sidebar('sidebar-blog'); ?>
            </aside>
            <?php endif; ?>

        </div><!-- .content-layout -->
    </div><!-- .container -->

</main>

<?php get_footer(); ?>
