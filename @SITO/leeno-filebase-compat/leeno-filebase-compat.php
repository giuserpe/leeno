<?php
/**
 * Plugin Name: LeenO WP Filebase Compatibility
 * Plugin URI: https://leeno.org
 * Description: Interpreta shortcode [wpfilebase] legacy senza dipendere dal plugin WP Filebase
 * Version: 1.2.0
 * Author: LeenO Team
 * License: GPL2
 * Text Domain: leeno-filebase-compat
 * Domain Path: /languages
 */

if (!defined('ABSPATH')) exit;

define('LEENO_FC_URL', plugin_dir_url(__FILE__));
define('LEENO_FC_VERSION', '1.2.0');

add_shortcode('wpfilebase', 'leeno_fc_shortcode');
add_action('wp_enqueue_scripts', 'leeno_fc_enqueue');

function leeno_fc_enqueue() {
    // Il CSS viene caricato sempre: su leeno-theme le regole del tema
    // già coprono .leeno-table-wrap, .leeno-table, .btn-leeno-download —
    // questo foglio aggiunge solo le variabili di fallback e il badge,
    // senza sovrapporsi a nulla.
    wp_enqueue_style(
        'leeno-fc',
        LEENO_FC_URL . 'assets/css/leeno-filebase-compat.css',
        [],
        LEENO_FC_VERSION
    );
}

function leeno_fc_shortcode($atts) {
    $atts = shortcode_atts(array('tag' => 'file', 'id' => ''), $atts);
    $id = intval($atts['id']);

    if ($atts['tag'] === 'file') {
        return leeno_fc_single($id);
    } elseif ($atts['tag'] === 'list') {
        return leeno_fc_list($id);
    }
    return '';
}

function leeno_fc_single($id) {
    global $wpdb;
    $table = $wpdb->prefix . 'wpfb_files';

    if (!$wpdb->get_var("SHOW TABLES LIKE '{$table}'")) return '';

    $file = $wpdb->get_row($wpdb->prepare(
        "SELECT file_name, file_display_name, file_size FROM {$table} WHERE file_id = %d",
        $id
    ), ARRAY_A);

    if (!$file) return '';

    $url  = content_url('/uploads/filebase/LeenO/LeenO/' . urlencode($file['file_name']));
    $name = $file['file_display_name'] ?: $file['file_name'];
    $size = $file['file_size'] ? leeno_fc_bytes($file['file_size']) : '';

    // Link inline con aspetto di pulsante download (stile tema)
    $out  = '<a href="' . esc_url($url) . '" class="btn-leeno-download" download>';
    $out .= '<svg width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="3" stroke-linecap="round" stroke-linejoin="round" aria-hidden="true">';
    $out .= '<path d="M21 15v4a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2v-4"/>';
    $out .= '<polyline points="7 10 12 15 17 10"/>';
    $out .= '<line x1="12" y1="15" x2="12" y2="3"/>';
    $out .= '</svg>';
    $out .= '<span>' . esc_html($name);
    if ($size) {
        $out .= ' (' . esc_html($size) . ')';
    }
    $out .= '</span></a>';
    return $out;
}

function leeno_fc_list($cat) {
    global $wpdb;
    $table = $wpdb->prefix . 'wpfb_files';

    if (!$wpdb->get_var("SHOW TABLES LIKE '{$table}'")) {
        return leeno_fc_folder_list();
    }

    $files = $wpdb->get_results($wpdb->prepare(
        "SELECT file_name, file_display_name, file_size, file_date FROM {$table} WHERE file_category = %d ORDER BY file_date DESC",
        $cat
    ), ARRAY_A);

    return empty($files) ? leeno_fc_folder_list() : leeno_fc_table($files);
}

function leeno_fc_folder_list() {
    $path = WP_CONTENT_DIR . '/uploads/filebase/LeenO/LeenO';
    if (!is_dir($path)) return '';

    $list = array_diff(scandir($path), ['.', '..']);
    if (empty($list)) return '';

    $files = [];
    foreach ($list as $f) {
        $fp = $path . '/' . $f;
        if (!is_dir($fp)) {
            $files[] = [
                'file_name'         => $f,
                'file_display_name' => $f,
                'file_size'         => filesize($fp),
                'file_date'         => date('Y-m-d H:i:s', filemtime($fp)),
            ];
        }
    }

    usort($files, fn($a, $b) => strtotime($b['file_date']) <=> strtotime($a['file_date']));
    return leeno_fc_table($files);
}

/**
 * Genera la tabella file con classi identiche a leeno-theme
 * (page-archivio-download.php): .leeno-table-wrap, .leeno-table,
 * .col-name, .col-date, .col-dim, .col-dl, .btn-leeno-download.
 */
function leeno_fc_table($files) {
    if (empty($files)) return '';

    $total = count($files);
    $label = $total === 1 ? '1 versione disponibile' : $total . ' versioni disponibili';

    $svg_dl = '<svg width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor"'
            . ' stroke-width="3" stroke-linecap="round" stroke-linejoin="round" aria-hidden="true">'
            . '<path d="M21 15v4a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2v-4"/>'
            . '<polyline points="7 10 12 15 17 10"/>'
            . '<line x1="12" y1="15" x2="12" y2="3"/>'
            . '</svg>';

    $html  = '<div class="wpfb-file-list">';
    $html .= '<div class="archivio-header">';
    $html .= '<span class="archivio-count">' . esc_html($label) . '</span>';
    $html .= '</div>';

    $html .= '<div class="leeno-table-wrap">';
    $html .= '<table class="leeno-table" role="table">';
    $html .= '<thead><tr>';
    $html .= '<th scope="col" class="col-name">Versione / File</th>';
    $html .= '<th scope="col" class="col-date" style="text-align:right;width:100px;">Data</th>';
    $html .= '<th scope="col" class="col-dim" style="text-align:right;width:100px;">Dim.</th>';
    $html .= '<th scope="col" class="col-dl" style="width:150px;"></th>';
    $html .= '</tr></thead>';
    $html .= '<tbody>';

    $i = 0;
    foreach ($files as $f) {
        $url     = content_url('/uploads/filebase/LeenO/LeenO/' . urlencode($f['file_name']));
        $name    = $f['file_display_name'] ?: $f['file_name'];
        $size    = $f['file_size'] ? leeno_fc_bytes($f['file_size']) : '—';
        $date    = $f['file_date'] ? date_i18n('d M Y', strtotime($f['file_date'])) : '—';
        $badge   = ($i++ === 0) ? ' <span class="wpfb-badge">ULTIMA</span>' : '';

        // Estrai numero versione dal nome file (es. "LeenO_5.3.2.oxt" → "5.3.2")
        $version = preg_match('/(\d+\.\d+[\.\d]*)/', $name, $m) ? $m[1] : null;
        $ver_badge = $version
            ? '<span class="archivio-ver-badge">v' . esc_html($version) . '</span> '
            : '';

        $html .= '<tr class="leeno-row">';

        // col-name
        $html .= '<td class="col-name">';
        $html .= '<a href="' . esc_url($url) . '">';
        $html .= $ver_badge . esc_html($name);
        $html .= '</a>' . $badge;
        $html .= '</td>';

        // col-date
        $html .= '<td class="col-date" style="color:#666;text-align:right;">';
        $html .= esc_html($date);
        $html .= '</td>';

        // col-dim
        $html .= '<td class="col-dim" style="text-align:right;">';
        $html .= esc_html($size);
        $html .= '</td>';

        // col-dl — pulsante identico a page-archivio-download.php
        $html .= '<td class="col-dl" style="text-align:right;">';
        $html .= '<a href="' . esc_url($url) . '" class="btn-leeno-download"'
              .  ' aria-label="Scarica ' . esc_attr($name) . '" download>';
        $html .= $svg_dl . '<span>Scarica</span>';
        $html .= '</a>';
        $html .= '</td>';

        $html .= '</tr>';
    }

    $html .= '</tbody></table></div>';
    $html .= '</div>';

    return $html;
}

function leeno_fc_bytes($b) {
    $u = ['B', 'KB', 'MB', 'GB'];
    $b = max($b, 0);
    $p = floor(($b ? log($b) : 0) / log(1024));
    $b /= (1 << (10 * $p));
    return round($b, 2) . ' ' . $u[$p];
}
