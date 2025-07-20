<?php
/*
Plugin Name: LeenO - Versioni di Sviluppo
Description: Mostra le versioni di sviluppo di LeenO tramite shortcode.
Version: 1.0
Author: LeenO Project
*/

function leeno_versioni_shortcode() {
    $file_path = plugin_dir_path(__FILE__) . 'versions.html';
    if (!file_exists($file_path)) {
        return '<p>⚠️ File delle versioni non trovato.</p>';
    }
    $html = file_get_contents($file_path);
    return $html;
}
add_shortcode('leeno_versioni', 'leeno_versioni_shortcode');
