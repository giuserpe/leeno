<?php
/**
 * Plugin Name: LeenO Versioni Sviluppo
 * Description: Mostra il file versions.html generato automaticamente da GitHub.
 * Version: 1.0
 * Author: giuserpe
 */

add_shortcode('leeno_versions', function() {
    $file = plugin_dir_path(__FILE__) . 'versions.html';
    if (file_exists($file)) {
        return file_get_contents($file);
    } else {
        return '<p><em>File versions.html non trovato.</em></p>';
    }
});

// Aggiunge stile base (opzionale)
add_action('wp_enqueue_scripts', function() {
    wp_register_style('leeno-versions-style', false);
    wp_enqueue_style('leeno-versions-style');
    wp_add_inline_style('leeno-versions-style', '.version-card{border:1px solid #ccc;padding:1rem;margin-bottom:1rem;border-radius:8px;background:#f9f9f9;}.version-card .badge{background:#28a745;color:white;padding:0.2rem .5rem;border-radius:5px;font-size:.8rem;}');
});
