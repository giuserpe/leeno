<?php
/**
 * Plugin Name: LeenO WP Filebase Compatibility
 * Plugin URI: https://leeno.org
 * Description: Interpreta shortcode [wpfilebase] legacy senza dipendere dal plugin WP Filebase (broken con PHP 8+). Include styling, caching, e ottimizzazioni.
 * Version: 1.1.0
 * Author: LeenO Team
 * Author URI: https://leeno.org
 * License: GPL2
 * Text Domain: leeno-filebase-compat
 * Domain Path: /languages
 */

if (!defined('ABSPATH')) {
    exit;
}

class LeenO_Filebase_Compat {
    
    private static $file_cache = [];
    private static $cache_ttl = 3600; // 1 ora
    private static $debug = false;
    
    public static function init() {
        // Shortcode
        add_shortcode('wpfilebase', [self::class, 'handle_shortcode']);
        
        // CSS
        add_action('wp_enqueue_scripts', [self::class, 'enqueue_styles']);
        add_action('admin_enqueue_scripts', [self::class, 'enqueue_styles']);
        
        // Pulizia cache
        add_action('wp_scheduled_event_leeno_clear_cache', [self::class, 'clear_cache']);
        
        // Schedule pulizia daily
        if (!wp_next_scheduled('wp_scheduled_event_leeno_clear_cache')) {
            wp_schedule_event(time(), 'daily', 'wp_scheduled_event_leeno_clear_cache');
        }
    }
    
    /**
     * Enqueue CSS
     */
    public static function enqueue_styles() {
        wp_enqueue_style(
            'leeno-filebase-compat',
            plugin_dir_url(__FILE__) . 'leeno-filebase-compat.css',
            [],
            '1.1.0'
        );
    }
    
    /**
     * Interpreta shortcode [wpfilebase tag=... id=...]
     */
    public static function handle_shortcode($atts) {
        $atts = shortcode_atts([
            'tag' => 'file',
            'id' => '',
            'cat' => '',
            'sort' => 'date',
            'order' => 'DESC',
            'limit' => '',
            'class' => '',
        ], $atts, 'wpfilebase');
        
        if (empty($atts['id'])) {
            return self::debug_note('richiede id');
        }
        
        $tag = strtolower(trim($atts['tag']));
        $id = intval($atts['id']);
        
        if ($tag === 'file') {
            return self::render_single_file($id, $atts);
        } elseif ($tag === 'list') {
            return self::render_file_list($id, $atts);
        }
        
        return self::debug_note('tag="' . esc_attr($tag) . '" non supportato');
    }
    
    /**
     * Renderizza un singolo file per il download
     */
    private static function render_single_file($file_id, $atts) {
        $file = self::get_file_by_id($file_id);
        
        if (!$file) {
            return self::debug_note('File ID ' . $file_id . ' non trovato');
        }
        
        $download_url = self::get_file_url($file);
        $filename = !empty($file['display_name']) ? $file['display_name'] : $file['name'];
        $size = !empty($file['size']) ? self::format_bytes($file['size']) : '';
        $class = !empty($atts['class']) ? ' ' . sanitize_html_class($atts['class']) : '';
        
        $html = '<a href="' . esc_url($download_url) . '" ';
        $html .= 'class="wpfb-file-link' . $class . '" ';
        $html .= 'title="' . esc_attr($filename) . '" ';
        $html .= 'download>';
        $html .= '<svg class="wpfb-icon" width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><path d="M21 15v4a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2v-4"/><polyline points="7 10 12 15 17 10"/><line x1="12" y1="15" x2="12" y2="3"/></svg>';
        $html .= esc_html($filename);
        if ($size) {
            $html .= ' <span class="wpfb-filesize">(' . esc_html($size) . ')</span>';
        }
        $html .= '</a>';
        
        return $html;
    }
    
    /**
     * Renderizza lista di file di una categoria
     */
    private static function render_file_list($category_id, $atts) {
        $files = self::get_files_by_category($category_id);
        
        if (empty($files)) {
            return self::debug_note('Nessun file in categoria ' . $category_id);
        }
        
        $sort = isset($atts['sort']) ? strtolower($atts['sort']) : 'date';
        $order = isset($atts['order']) ? strtoupper($atts['order']) : 'DESC';
        $limit = !empty($atts['limit']) ? intval($atts['limit']) : 0;
        
        // Ordina file
        usort($files, function($a, $b) use ($sort, $order) {
            if ($sort === 'name') {
                $cmp = strcmp($a['name'], $b['name']);
            } elseif ($sort === 'size') {
                $cmp = ($a['size'] <=> $b['size']);
            } else { // date
                $cmp = strtotime($a['date'] ?? '1970-01-01') <=> strtotime($b['date'] ?? '1970-01-01');
            }
            return $order === 'ASC' ? $cmp : -$cmp;
        });
        
        // Applica limit
        if ($limit > 0) {
            $files = array_slice($files, 0, $limit);
        }
        
        $class = !empty($atts['class']) ? ' ' . sanitize_html_class($atts['class']) : '';
        
        $html = '<div class="wpfb-file-list' . $class . '">';
        $html .= '<table class="wpfb-table">';
        $html .= '<thead>';
        $html .= '<tr>';
        $html .= '<th class="col-name">File</th>';
        $html .= '<th class="col-size">Dimensione</th>';
        $html .= '<th class="col-date">Data</th>';
        $html .= '<th class="col-action">Download</th>';
        $html .= '</tr>';
        $html .= '</thead>';
        $html .= '<tbody>';
        
        $row_count = 0;
        foreach ($files as $file) {
            $download_url = self::get_file_url($file);
            $filename = !empty($file['display_name']) ? $file['display_name'] : $file['name'];
            $size = !empty($file['size']) ? self::format_bytes($file['size']) : '—';
            $date = !empty($file['date']) ? date_i18n('d M Y', strtotime($file['date'])) : '—';
            
            $row_class = ($row_count % 2 === 0) ? 'even' : 'odd';
            
            $html .= '<tr class="' . $row_class . '">';
            $html .= '<td class="col-name">' . esc_html($filename) . '</td>';
            $html .= '<td class="col-size">' . esc_html($size) . '</td>';
            $html .= '<td class="col-date">' . esc_html($date) . '</td>';
            $html .= '<td class="col-action">';
            $html .= '<a href="' . esc_url($download_url) . '" class="wpfb-btn wpfb-btn-download" download aria-label="Scarica ' . esc_attr($filename) . '">';
            $html .= '<svg class="wpfb-icon" width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><path d="M21 15v4a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2v-4"/><polyline points="7 10 12 15 17 10"/><line x1="12" y1="15" x2="12" y2="3"/></svg>';
            $html .= 'Scarica</a>';
            $html .= '</td>';
            $html .= '</tr>';
            
            $row_count++;
        }
        
        $html .= '</tbody>';
        $html .= '</table>';
        $html .= '</div>';
        
        return $html;
    }
    
    /**
     * Recupera un file per ID (con caching)
     */
    private static function get_file_by_id($file_id) {
        if (isset(self::$file_cache[$file_id])) {
            return self::$file_cache[$file_id];
        }
        
        // Prova cache transient
        $cache_key = 'leeno_wpfb_file_' . $file_id;
        $cached = get_transient($cache_key);
        if ($cached !== false) {
            self::$file_cache[$file_id] = $cached;
            return $cached;
        }
        
        global $wpdb;
        
        $table_files = $wpdb->prefix . 'wpfb_files';
        if (!$wpdb->get_var("SHOW TABLES LIKE '{$table_files}'")) {
            return null;
        }
        
        $file = $wpdb->get_row(
            $wpdb->prepare(
                "SELECT file_id, file_name, file_display_name, file_size, file_date, file_url 
                 FROM {$table_files} WHERE file_id = %d LIMIT 1",
                $file_id
            ),
            ARRAY_A
        );
        
        if ($file) {
            $result = [
                'id' => $file['file_id'],
                'name' => $file['file_name'],
                'display_name' => $file['file_display_name'],
                'size' => $file['file_size'],
                'date' => $file['file_date'],
                'url' => $file['file_url'],
            ];
            
            // Cache per 1 ora
            set_transient($cache_key, $result, self::$cache_ttl);
            self::$file_cache[$file_id] = $result;
            
            return $result;
        }
        
        return null;
    }
    
    /**
     * Recupera file di una categoria (con caching e fallback a cartella fisica)
     */
    private static function get_files_by_category($category_id) {
        // Prova cache transient
        $cache_key = 'leeno_wpfb_cat_' . $category_id;
        $cached = get_transient($cache_key);
        if ($cached !== false) {
            return $cached;
        }
        
        global $wpdb;
        
        $table_files = $wpdb->prefix . 'wpfb_files';
        $result = [];
        
        // Prova a leggere dal DB
        if ($wpdb->get_var("SHOW TABLES LIKE '{$table_files}'")) {
            $files = $wpdb->get_results(
                $wpdb->prepare(
                    "SELECT file_id, file_name, file_display_name, file_size, file_date, file_url 
                     FROM {$table_files} 
                     WHERE file_category = %d 
                     ORDER BY file_date DESC, file_name ASC",
                    $category_id
                ),
                ARRAY_A
            );
            
            if (!empty($files)) {
                foreach ($files as $file) {
                    $result[] = [
                        'id' => $file['file_id'],
                        'name' => $file['file_name'],
                        'display_name' => $file['file_display_name'],
                        'size' => $file['file_size'],
                        'date' => $file['file_date'],
                        'url' => $file['file_url'],
                    ];
                }
                
                // Cache per 1 ora
                set_transient($cache_key, $result, self::$cache_ttl);
                return $result;
            }
        }
        
        // ── FALLBACK: Leggi da cartella fisica se DB è vuoto ──
        $base_path = WP_CONTENT_DIR . '/uploads/filebase/LeenO/LeenO';
        if (is_dir($base_path)) {
            $files = array_diff(scandir($base_path), array('.', '..'));
            foreach ($files as $filename) {
                $filepath = $base_path . '/' . $filename;
                if (is_dir($filepath)) continue;
                
                $result[] = [
                    'id' => crc32($filename),
                    'name' => $filename,
                    'display_name' => $filename,
                    'size' => filesize($filepath),
                    'date' => date('Y-m-d H:i:s', filemtime($filepath)),
                    'url' => content_url('/uploads/filebase/LeenO/LeenO/' . urlencode($filename)),
                ];
            }
            
            // Ordina per data decrescente
            usort($result, function($a, $b) {
                return strtotime($b['date']) <=> strtotime($a['date']);
            });
            
            // Cache per 1 ora
            set_transient($cache_key, $result, self::$cache_ttl);
            return $result;
        }
        
        return [];
    }
    
    /**
     * Genera URL di download per il file
     */
    private static function get_file_url($file) {
        if (!empty($file['url'])) {
            return $file['url'];
        }
        
        $base_path = WP_CONTENT_DIR . '/uploads/filebase/LeenO/LeenO';
        $file_path = $base_path . '/' . basename($file['name']);
        
        if (file_exists($file_path)) {
            return content_url('/uploads/filebase/LeenO/LeenO/' . urlencode(basename($file['name'])));
        }
        
        return home_url('/wp-content/plugins/wp-filebase/download.php?id=' . intval($file['id']));
    }
    
    /**
     * Formatta byte in KB/MB/GB
     */
    private static function format_bytes($bytes, $precision = 2) {
        $units = ['B', 'KB', 'MB', 'GB', 'TB'];
        $bytes = max($bytes, 0);
        $pow = floor(($bytes ? log($bytes) : 0) / log(1024));
        $pow = min($pow, count($units) - 1);
        $bytes /= (1 << (10 * $pow));
        
        return round($bytes, $precision) . ' ' . $units[$pow];
    }
    
    /**
     * Nota di debug (solo per admin)
     */
    private static function debug_note($msg) {
        if (current_user_can('manage_options')) {
            return '<!-- [wpfilebase] ' . esc_html($msg) . ' -->';
        }
        return '';
    }
    
    /**
     * Svuota la cache
     */
    public static function clear_cache() {
        global $wpdb;
        $wpdb->query("DELETE FROM $wpdb->options WHERE option_name LIKE '_transient_leeno_wpfb_%'");
    }
}

// Deactivation hook
register_deactivation_hook(__FILE__, function() {
    wp_clear_scheduled_hook('wp_scheduled_event_leeno_clear_cache');
    LeenO_Filebase_Compat::clear_cache();
});

// Init
LeenO_Filebase_Compat::init();
