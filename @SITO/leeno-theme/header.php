<!DOCTYPE html>
<html <?php language_attributes(); ?>>
<head>
    <meta charset="<?php bloginfo('charset'); ?>">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <?php wp_head(); ?>
</head>
<body <?php body_class(); ?>>
<?php wp_body_open(); ?>

<a class="skip-link screen-reader-text" href="#main-content"><?php esc_html_e('Vai al contenuto', 'leeno-dm'); ?></a>

<header class="site-header" id="siteHeader" role="banner">
    <div class="header-inner">
        <a href="<?php echo esc_url(home_url('/')); ?>" class="site-logo" rel="home">
            <img src="<?php echo esc_url(get_template_directory_uri() . '/assets/images/logo-leeno.png'); ?>" alt="<?php bloginfo('name'); ?>">
            <span class="site-name">LEENO</span>
        </a>

        <nav class="main-nav" id="mainNav">
            <?php
            wp_nav_menu([
                'theme_location'  => 'primary',
                'menu_class'      => 'nav-menu',
                'container'       => false,
                'fallback_cb'     => false,
                'depth'           => 2,
            ]);
            ?>
        </nav>

        <div class="header-search" id="headerSearch" role="search">
            <button class="header-search-toggle" id="headerSearchToggle" aria-label="Apri ricerca" aria-expanded="false">
                <svg width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.5" aria-hidden="true"><circle cx="11" cy="11" r="8"/><line x1="21" y1="21" x2="16.65" y2="16.65"/></svg>
            </button>
            <div class="header-search-box" id="headerSearchBox" hidden>
                <form role="search" method="get" action="https://leeno.org/">
                    <input
                        type="search"
                        class="header-search-input"
                        id="headerSearchInput"
                        placeholder="Cerca nel sito…"
                        value="<?php echo esc_attr( get_search_query() ); ?>"
                        name="s"
                        autocomplete="off"
                        aria-label="Cerca"
                    >
                    <button type="submit" class="header-search-submit" aria-label="Invia ricerca">
                        <svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.5" aria-hidden="true"><line x1="5" y1="12" x2="19" y2="12"/><polyline points="12 5 19 12 12 19"/></svg>
                    </button>
                </form>
            </div>
        </div>

        <a href="<?php echo esc_url(home_url('/about-leeno/leeno/download/')); ?>" class="btn-download" aria-label="Scarica LeenO">
            <svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.5" aria-hidden="true"><path d="M21 15v4a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2v-4"/><polyline points="7 10 12 15 17 10"/><line x1="12" y1="15" x2="12" y2="3"/></svg>
            Scarica
        </a>

        <button class="menu-toggle" id="menuToggle" aria-label="Apri menu" aria-expanded="false" aria-controls="mainNav">
            <span></span>
            <span></span>
            <span></span>
        </button>
    </div>
</header>
