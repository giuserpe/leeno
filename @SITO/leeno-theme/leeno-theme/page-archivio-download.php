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
                <a href="<?php echo esc_url( home_url('/scarica-leeno/') ); ?>">Download</a>
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
            global $wpdb;

            $files = array();
            $files_table = isset($wpdb->prefix) ? $wpdb->prefix . 'wpfb_files' : 'wp_wpfb_files';
            $cats_table  = isset($wpdb->prefix) ? $wpdb->prefix . 'wpfb_cats' : 'wp_wpfb_cats';
            $has_table   = (isset($wpdb) && is_object($wpdb)) ? (bool) $wpdb->get_var($wpdb->prepare('SHOW TABLES LIKE %s', $files_table)) : false;

            if ( $has_table ) {
                $has_repo = (bool) $wpdb->get_var("SHOW COLUMNS FROM {$files_table} LIKE 'file_repository'");

                $cat_filter = "f.file_category IN (
                            SELECT cat_id FROM {$cats_table}
                            WHERE cat_id = 199 OR cat_parent = 199
                        )
                     OR f.file_category = 199";

                $where = $has_repo
                    ? "(f.file_repository = 199 OR {$cat_filter})"
                    : "({$cat_filter})";

                $files = $wpdb->get_results(
                    "SELECT f.*, c.cat_name
                     FROM {$files_table} f
                     LEFT JOIN {$cats_table} c ON f.file_category = c.cat_id
                     WHERE {$where}
                     ORDER BY f.file_date DESC, f.file_display_name ASC"
                );

                if ( empty( $files ) ) {
                    $files = $wpdb->get_results(
                        "SELECT f.*, c.cat_name
                         FROM {$files_table} f
                         LEFT JOIN {$cats_table} c ON f.file_category = c.cat_id
                         WHERE f.file_path LIKE 'LeenO/Archivio/%'
                         ORDER BY f.file_date DESC, f.file_display_name ASC"
                    );
                }
            }

            if ( empty( $files ) && function_exists('leeno_fc_folder_list') ) :
                echo leeno_fc_folder_list(199);
            elseif ( empty( $files ) ) : ?>
                <p class="prezzari-error">Nessun file trovato nella repository.</p>

                <?php if ( current_user_can('administrator') ) :
                    $repo_col = $has_repo ? ', file_repository' : '';
                    $sample = $wpdb->get_results(
                        "SELECT file_id, file_display_name, file_category, file_date{$repo_col}
                         FROM {$files_table}
                         ORDER BY file_date DESC LIMIT 10"
                    );
                    echo '<div style="background:#1a2010;color:#aad400;font-family:monospace;font-size:11px;padding:12px 20px;margin:16px 0;border-left:4px solid #aad400">';
                    echo '<strong>DEBUG — ultimi 10 file nel DB:</strong><br>';
                    foreach ( $sample as $f ) {
                        $repo = isset($f->file_repository) ? $f->file_repository : '—';
                        echo "ID={$f->file_id} | cat={$f->file_category} | repo={$repo} | data={$f->file_date} | " . esc_html($f->file_display_name) . "<br>";
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
                            <th scope="col" class="col-name">Versione / File</th>
                            <th scope="col" class="col-date" style="text-align: right; width: 100px;">Data</th>
                            <th scope="col" class="col-dim" style="text-align: right; width: 100px;">Dim.</th>
                            <th scope="col" class="col-dl" style="width: 150px;"></th>
                        </tr>
                    </thead>
                    <tbody>
                    <?php foreach ( $files as $file ) :
                        $dl_url = leeno_fc_file_url( $file );
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
                            <td class="col-date" style="color: #666; text-align: right;">
                                <?php echo esc_html($date); ?>
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

            <?php endif; // files
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
