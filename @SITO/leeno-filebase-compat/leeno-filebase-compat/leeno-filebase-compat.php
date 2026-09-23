<?php
/**
 * Plugin Name: LeenO WP Filebase Compatibility
 * Plugin URI: https://leeno.org
 * Description: Interpreta shortcode [wpfilebase] legacy senza dipendere dal plugin WP Filebase
 * Version: 1.3.0
 * Author: LeenO Team
 * License: GPL2
 * Text Domain: leeno-filebase-compat
 * Domain Path: /languages
 */

if (!defined('ABSPATH')) exit;

define('LEENO_FC_URL', plugin_dir_url(__FILE__));
define('LEENO_FC_VERSION', '1.3.0');
define('LEENO_FC_ROOT_REL', 'uploads/filebase');

add_shortcode('wpfilebase', 'leeno_fc_shortcode');
add_action('wp_enqueue_scripts', 'leeno_fc_enqueue');

function leeno_fc_enqueue() {
    wp_enqueue_style(
        'leeno-fc',
        LEENO_FC_URL . 'assets/css/leeno-filebase-compat.css',
        [],
        LEENO_FC_VERSION
    );
}

function leeno_fc_files_table() {
    global $wpdb;
    return $wpdb->prefix . 'wpfb_files';
}

function leeno_fc_cats_table() {
    global $wpdb;
    return $wpdb->prefix . 'wpfb_cats';
}

function leeno_fc_ready() {
    global $wpdb;
    static $ok = null;
    if ($ok === null) {
        $table = leeno_fc_files_table();
        $has_table = (bool) $wpdb->get_var($wpdb->prepare('SHOW TABLES LIKE %s', $table));
        $has_dir   = is_dir(leeno_fc_root_dir());
        $ok = $has_table || $has_dir;
    }
    return $ok;
}

function leeno_fc_cats_ready() {
    global $wpdb;
    static $ok = null;
    if ($ok === null) {
        $table = leeno_fc_cats_table();
        $ok = (bool) $wpdb->get_var($wpdb->prepare('SHOW TABLES LIKE %s', $table));
    }
    return $ok;
}

function leeno_fc_field($file, $key) {
    if (is_array($file)) {
        if (isset($file[$key])) {
            return $file[$key];
        }
        $alt = (strpos($key, 'file_') === 0) ? substr($key, 5) : 'file_' . $key;
        return $file[$alt] ?? '';
    }
    if (is_object($file)) {
        return $file->$key ?? '';
    }
    return '';
}

function leeno_fc_root_dir() {
    return WP_CONTENT_DIR . '/' . LEENO_FC_ROOT_REL;
}

function leeno_fc_norm_rel($rel) {
    $rel = str_replace('\\', '/', (string) $rel);
    $rel = preg_replace('#/+#', '/', $rel);
    return trim($rel, '/');
}

function leeno_fc_encode_rel($rel) {
    $rel = leeno_fc_norm_rel($rel);
    if ($rel === '') {
        return '';
    }
    $parts = array_map('rawurlencode', explode('/', $rel));
    return implode('/', $parts);
}

function leeno_fc_rel_url($rel) {
    $enc = leeno_fc_encode_rel($rel);
    if ($enc === '') {
        return '';
    }
    return content_url('/' . LEENO_FC_ROOT_REL . '/' . $enc);
}

function leeno_fc_rel_exists($rel) {
    $rel = leeno_fc_norm_rel($rel);
    if ($rel === '') {
        return false;
    }
    $abs = leeno_fc_root_dir() . '/' . $rel;
    return is_file($abs);
}

function leeno_fc_cat_relpath($cat_id) {
    static $cache = [];
    $cat_id = (int) $cat_id;
    if (!$cat_id || !leeno_fc_cats_ready()) {
        return '';
    }
    if (array_key_exists($cat_id, $cache)) {
        return $cache[$cat_id];
    }
    global $wpdb;
    $table = leeno_fc_cats_table();
    $row = $wpdb->get_row($wpdb->prepare(
        "SELECT cat_path, cat_folder FROM {$table} WHERE cat_id = %d",
        $cat_id
    ), ARRAY_A);
    if (!$row) {
        $cache[$cat_id] = '';
        return '';
    }
    $rel = !empty($row['cat_path'])
        ? leeno_fc_norm_rel($row['cat_path'])
        : leeno_fc_norm_rel($row['cat_folder'] ?? '');
    $cache[$cat_id] = $rel;
    return $rel;
}

function leeno_fc_index_by_name() {
    static $index = null;
    if ($index !== null) {
        return $index;
    }
    $index = [];
    $root = leeno_fc_root_dir();
    if (!is_dir($root)) {
        return $index;
    }
    $skip = ['.tmp', '.git'];
    try {
        $it = new RecursiveIteratorIterator(
            new RecursiveDirectoryIterator($root, FilesystemIterator::SKIP_DOTS)
        );
        foreach ($it as $f) {
            if (!$f->isFile()) {
                continue;
            }
            $name = $f->getFilename();
            if (strpos($name, 'thumb_') === 0) {
                continue;
            }
            $rel = leeno_fc_norm_rel(str_replace($root, '', $f->getPathname()));
            $top = explode('/', $rel)[0] ?? '';
            if (in_array($top, $skip, true)) {
                continue;
            }
            if (!isset($index[$name])) {
                $index[$name] = $rel;
            }
        }
    } catch (Exception $e) {
        return $index;
    }
    return $index;
}

function leeno_fc_find_rel_by_name($filename) {
    $filename = basename((string) $filename);
    if ($filename === '' || $filename === '.' || $filename === '..') {
        return '';
    }
    $index = leeno_fc_index_by_name();
    return $index[$filename] ?? '';
}

/**
 * URL pubblica del file. Accetta riga DB (oggetto o array) con file_path/file_name.
 */
function leeno_fc_file_url($file) {
    $path = leeno_fc_norm_rel(leeno_fc_field($file, 'file_path'));
    $name = basename((string) leeno_fc_field($file, 'file_name'));
    $cat  = (int) leeno_fc_field($file, 'file_category');

    $candidates = [];
    if ($path !== '') {
        $candidates[] = $path;
        if ($name !== '' && substr($path, -strlen($name)) !== $name) {
            $candidates[] = $path . '/' . $name;
        }
    }
    if ($name !== '') {
        $cat_rel = leeno_fc_cat_relpath($cat);
        if ($cat_rel !== '') {
            $candidates[] = $cat_rel . '/' . $name;
        }
        $candidates[] = 'LeenO/LeenO/' . $name;
    }

    foreach ($candidates as $rel) {
        if (leeno_fc_rel_exists($rel)) {
            return leeno_fc_rel_url($rel);
        }
    }

    if ($name !== '') {
        $found = leeno_fc_find_rel_by_name($name);
        if ($found !== '') {
            return leeno_fc_rel_url($found);
        }
    }

    if (!empty($candidates)) {
        return leeno_fc_rel_url($candidates[0]);
    }
    return '';
}

function leeno_fc_shortcode($atts) {
    $atts = shortcode_atts(array('tag' => 'file', 'id' => ''), $atts);
    $id = intval($atts['id']);

    if ($atts['tag'] === 'file') {
        return leeno_fc_single($id);
    } elseif ($atts['tag'] === 'list') {
        return leeno_fc_list($id, false);
    } elseif ($atts['tag'] === 'browser') {
        return leeno_fc_list($id, true);
    }
    return '';
}

function leeno_fc_single($id) {
    global $wpdb;
    if (!leeno_fc_ready()) {
        return '';
    }

    $table = leeno_fc_files_table();
    $file = $wpdb->get_row($wpdb->prepare(
        "SELECT file_name, file_display_name, file_size, file_path, file_category
         FROM {$table} WHERE file_id = %d",
        $id
    ), ARRAY_A);

    if (!$file) {
        return '';
    }

    $url  = leeno_fc_file_url($file);
    $name = $file['file_display_name'] ?: $file['file_name'];
    $size = $file['file_size'] ? leeno_fc_bytes($file['file_size']) : '';

    $out  = '<a href="' . esc_url($url) . '" class="btn-leeno-download" download>';
    $out .= leeno_fc_svg_dl();
    $out .= '<span>' . esc_html($name);
    if ($size) {
        $out .= ' (' . esc_html($size) . ')';
    }
    $out .= '</span></a>';
    return $out;
}

function leeno_fc_descendant_cat_ids($cat) {
    $cat = (int) $cat;
    $ids = [$cat];
    if (!$cat || !leeno_fc_cats_ready()) {
        return $ids;
    }
    global $wpdb;
    $table = leeno_fc_cats_table();
    $path = leeno_fc_cat_relpath($cat);
    if ($path === '') {
        $children = $wpdb->get_col($wpdb->prepare(
            "SELECT cat_id FROM {$table} WHERE cat_parent = %d",
            $cat
        ));
        return array_map('intval', array_merge($ids, $children ?: []));
    }
    $like = $wpdb->esc_like($path) . '/%';
    $children = $wpdb->get_col($wpdb->prepare(
        "SELECT cat_id FROM {$table} WHERE cat_id = %d OR cat_path = %s OR cat_path LIKE %s",
        $cat,
        $path,
        $like
    ));
    return array_map('intval', $children ?: $ids);
}

function leeno_fc_list($cat, $descendants = false) {
    global $wpdb;
    if (!leeno_fc_ready()) {
        return leeno_fc_folder_list($cat);
    }

    $table = leeno_fc_files_table();
    $ids = $descendants ? leeno_fc_descendant_cat_ids($cat) : [(int) $cat];
    $in  = implode(',', array_map('intval', $ids));
    if ($in === '') {
        return leeno_fc_folder_list($cat);
    }

    $files = $wpdb->get_results(
        "SELECT file_name, file_display_name, file_size, file_date, file_path, file_category
         FROM {$table}
         WHERE file_category IN ({$in})
         ORDER BY file_date DESC",
        ARRAY_A
    );

    return empty($files) ? leeno_fc_folder_list($cat) : leeno_fc_table($files);
}

function leeno_fc_folder_list($cat = 0) {
    $rel = leeno_fc_cat_relpath((int) $cat);
    if ($rel === '') {
        $rel = 'LeenO/LeenO';
    }
    $path = leeno_fc_root_dir() . '/' . $rel;
    if (!is_dir($path)) {
        return '';
    }

    $list = array_diff(scandir($path), ['.', '..']);
    if (empty($list)) {
        return '';
    }

    $files = [];
    foreach ($list as $f) {
        $fp = $path . '/' . $f;
        if (is_dir($fp) || strpos($f, 'thumb_') === 0) {
            continue;
        }
        $files[] = [
            'file_name'         => $f,
            'file_display_name' => $f,
            'file_size'         => filesize($fp),
            'file_date'         => date('Y-m-d H:i:s', filemtime($fp)),
            'file_path'         => $rel . '/' . $f,
            'file_category'     => (int) $cat,
        ];
    }

    usort($files, fn($a, $b) => strtotime($b['file_date']) <=> strtotime($a['file_date']));
    return leeno_fc_table($files);
}

function leeno_fc_svg_dl() {
    return '<svg width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor"'
         . ' stroke-width="3" stroke-linecap="round" stroke-linejoin="round" aria-hidden="true">'
         . '<path d="M21 15v4a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2v-4"/>'
         . '<polyline points="7 10 12 15 17 10"/>'
         . '<line x1="12" y1="15" x2="12" y2="3"/>'
         . '</svg>';
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
    $svg_dl = leeno_fc_svg_dl();

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
        $url     = leeno_fc_file_url($f);
        $name    = leeno_fc_field($f, 'file_display_name') ?: leeno_fc_field($f, 'file_name');
        $size    = leeno_fc_field($f, 'file_size') ? leeno_fc_bytes(leeno_fc_field($f, 'file_size')) : '—';
        $date    = leeno_fc_field($f, 'file_date') ? date_i18n('d M Y', strtotime(leeno_fc_field($f, 'file_date'))) : '—';
        $badge   = ($i++ === 0) ? ' <span class="wpfb-badge">ULTIMA</span>' : '';

        $version = preg_match('/(\d+\.\d+[\.\d]*)/', $name, $m) ? $m[1] : null;
        $ver_badge = $version
            ? '<span class="archivio-ver-badge">v' . esc_html($version) . '</span> '
            : '';

        $html .= '<tr class="leeno-row">';

        $html .= '<td class="col-name">';
        $html .= '<a href="' . esc_url($url) . '">';
        $html .= $ver_badge . esc_html($name);
        $html .= '</a>' . $badge;
        $html .= '</td>';

        $html .= '<td class="col-date" style="color:#666;text-align:right;">';
        $html .= esc_html($date);
        $html .= '</td>';

        $html .= '<td class="col-dim" style="text-align:right;">';
        $html .= esc_html($size);
        $html .= '</td>';

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
