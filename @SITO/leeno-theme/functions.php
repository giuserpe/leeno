<?php
/**
 * LeenO Digital Masonry - Functions
 */

if (!defined('ABSPATH')) exit;

// Theme version for cache busting
define('LEENO_DM_VERSION', '1.0.0');

/**
 * Theme setup
 */
add_action('after_setup_theme', function () {
    // Title tag support
    add_theme_support('title-tag');

    // Post thumbnails
    add_theme_support('post-thumbnails');
    set_post_thumbnail_size(150, 150, true);

    // HTML5 markup
    add_theme_support('html5', ['search-form', 'comment-form', 'comment-list', 'gallery', 'caption']);

    // Responsive embeds
    add_theme_support('responsive-embeds');

    // Custom logo
    add_theme_support('custom-logo', [
        'height'      => 60,
        'width'       => 200,
        'flex-height' => true,
        'flex-width'  => true,
    ]);

    // Menus
    register_nav_menus([
        'primary' => __('Menu Principale', 'leeno-dm'),
        'footer'  => __('Menu Footer', 'leeno-dm'),
    ]);
});

/**
 * Enqueue scripts and styles
 */
add_action('wp_enqueue_scripts', function () {
    // Google Fonts
    wp_enqueue_style(
        'leeno-fonts',
        'https://fonts.googleapis.com/css2?family=Inter:wght@300;400;500;600;700&family=JetBrains+Mono:wght@400;500;600&family=Space+Grotesk:wght@400;500;600;700&display=swap',
        [],
        null
    );

    // Main theme styles
    wp_enqueue_style(
        'leeno-main',
        get_template_directory_uri() . '/assets/css/main.css',
        ['leeno-fonts'],
        LEENO_DM_VERSION
    );

    // GSAP
    wp_enqueue_script(
        'gsap',
        'https://cdnjs.cloudflare.com/ajax/libs/gsap/3.12.5/gsap.min.js',
        [],
        '3.12.5',
        true
    );
    wp_enqueue_script(
        'gsap-scrolltrigger',
        'https://cdnjs.cloudflare.com/ajax/libs/gsap/3.12.5/ScrollTrigger.min.js',
        ['gsap'],
        '3.12.5',
        true
    );

    // Main JS - WebGL shader + CAD grid + animations
    wp_enqueue_script(
        'leeno-main',
        get_template_directory_uri() . '/assets/js/main.js',
        ['gsap', 'gsap-scrolltrigger'],
        LEENO_DM_VERSION,
        true
    );

    // Pass theme URL to JS
    wp_localize_script('leeno-main', 'leenoData', [
        'themeUrl' => get_template_directory_uri(),
        'ajaxUrl'  => admin_url('admin-ajax.php'),
    ]);
});

/**
 * Custom excerpt length
 */
add_filter('excerpt_length', function () {
    return 25;
});

add_filter('excerpt_more', function () {
    return '...';
});

/**
 * Add custom image sizes
 */
add_action('after_setup_theme', function () {
    add_image_size('leeno-card', 400, 250, true);
    add_image_size('leeno-hero', 1200, 600, true);
});

/**
 * Customizer: Accent color
 */
add_action('customize_register', function ($wp_customize) {
    $wp_customize->add_setting('leeno_accent_color', [
        'default'           => '#00e5ff',
        'sanitize_callback' => 'sanitize_hex_color',
        'transport'         => 'refresh',
    ]);

    $wp_customize->add_control(new WP_Customize_Color_Control($wp_customize, 'leeno_accent_color', [
        'label'   => __('Colore Accento', 'leeno-dm'),
        'section' => 'colors',
    ]));
});

/**
 * Widget areas
 */
add_action('widgets_init', function () {
    register_sidebar([
        'name'          => __('Sidebar Blog', 'leeno-dm'),
        'id'            => 'sidebar-blog',
        'before_widget' => '<div class="widget">',
        'after_widget'  => '</div>',
        'before_title'  => '<h4 class="widget-title">',
        'after_title'   => '</h4>',
    ]);
});
