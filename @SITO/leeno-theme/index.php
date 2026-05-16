<?php get_header(); ?>

<main id="main-content">

<!-- HERO SECTION -->
<section class="hero-section" id="heroSection" aria-label="Introduzione LeenO">
    <canvas class="hero-canvas" id="heroCanvas" aria-hidden="true"></canvas>
    <div class="hero-content">
        <h1 class="hero-title" id="heroTitle">
            <span class="word-wrap"><span class="word-inner">IL</span></span>
            <span class="word-wrap"><span class="word-inner">COMPUTO</span></span><br>
            <span class="word-wrap"><span class="word-inner text-accent">METRICO</span></span><br>
            <span class="word-wrap"><span class="word-inner">EVOLUTO</span></span>
        </h1>
        <p class="hero-subtitle" id="heroSubtitle">
            Software open-source per computi metrici estimativi su LibreOffice Calc.
            <br>
            <span class="version-badge">v3.25.0 &mdash; Marzo 2026</span>
        </p>
        <div class="hero-ctas">
            <a href="<?php echo esc_url(home_url('/about-leeno/leeno/download/')); ?>" class="btn-hero">
                <svg width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.5" aria-hidden="true"><path d="M21 15v4a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2v-4"/><polyline points="7 10 12 15 17 10"/><line x1="12" y1="15" x2="12" y2="3"/></svg>
                Scarica Gratuitamente
            </a>
            <a href="#featuresSection" class="btn-hero-ghost" id="heroScrollCta">
                Scopri le funzioni
                <svg width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" aria-hidden="true"><polyline points="6 9 12 15 18 9"/></svg>
            </a>
        </div>
    </div>
    <div class="hero-scroll-hint" aria-hidden="true">
        <span></span>
    </div>
</section>

<!-- TRANSITION / CAD GRID SECTION -->
<section class="transition-section" id="transitionSection">
    <canvas class="cad-canvas" id="cadCanvas"></canvas>
    <div class="transition-text" id="transitionText">
        <p>
            <span class="text-accent">PRECISIONE</span> ARCHITETTONICA.
            <br>
            ALGORITMI DI <span class="text-accent">ULTIMA GENERAZIONE</span>.
        </p>
    </div>
</section>

<!-- FEATURES SECTION -->
<section class="features-section" id="featuresSection">
    <div class="container">
        <h2 class="section-title" id="featuresTitle">
            <?php
            $title_words = ['STRUMENTI', 'PROFESSIONALI'];
            foreach ($title_words as $i => $word) {
                echo '<span class="title-word-outer"><span class="title-word-inner">' . esc_html($word) . '</span></span>';
                if ($i < count($title_words) - 1) echo ' ';
            }
            ?>
        </h2>

        <div class="features-grid">
            <div class="feature-card">
                <div class="feature-icon">
                    <svg width="24" height="24" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2">
                        <rect x="3" y="3" width="18" height="18" rx="2"/>
                        <path d="M3 9h18M9 21V9"/>
                    </svg>
                </div>
                <h3 class="feature-title">
                    <a href="https://leeno.org/about-leeno/" class="feature-title-link">
                        Automazione Completa
                        <svg width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.5" aria-hidden="true" style="vertical-align:middle;margin-left:6px;opacity:0.6"><line x1="5" y1="12" x2="19" y2="12"/><polyline points="12 5 19 12 12 19"/></svg>
                    </a>
                </h3>
                <p class="feature-desc">Genera computi metrici, contabilità di cantiere, varianti e libretti delle misure automaticamente. Ogni calcolo è gestito dalle macro integrate.</p>
            </div>

            <div class="feature-card">
                <div class="feature-icon">
                    <svg width="24" height="24" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2">
                        <ellipse cx="12" cy="5" rx="9" ry="3"/>
                        <path d="M21 12c0 1.66-4 3-9 3s-9-1.34-9-3"/>
                        <path d="M3 5v14c0 1.66 4 3 9 3s9-1.34 9-3V5"/>
                    </svg>
                </div>
                <h3 class="feature-title">
                    <a href="https://leeno.org/prezziari/download-prezziari/" class="feature-title-link">
                        Prezzari Regionali
                        <svg width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.5" aria-hidden="true" style="vertical-align:middle;margin-left:6px;opacity:0.6"><line x1="5" y1="12" x2="19" y2="12"/><polyline points="12 5 19 12 12 19"/></svg>
                    </a>
                </h3>
                <p class="feature-desc">Accesso ai prezzari regionali aggiornati annualmente. Importa gli elenchi prezzi di tutte le regioni e enti.</p>
            </div>

            <div class="feature-card">
                <div class="feature-icon">
                    <svg width="24" height="24" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2">
                        <polyline points="9 11 12 14 22 4"/>
                        <path d="M21 12v7a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2V5a2 2 0 0 1 2-2h11"/>
                    </svg>
                </div>
                <h3 class="feature-title">
                    <a href="https://leeno.org/leeno-e-il-computo-su-formato-opendocument-odf/" class="feature-title-link">
                        Conforme al DM 49/2018
                        <svg width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.5" aria-hidden="true" style="vertical-align:middle;margin-left:6px;opacity:0.6"><line x1="5" y1="12" x2="19" y2="12"/><polyline points="12 5 19 12 12 19"/></svg>
                    </a>
                </h3>
                <p class="feature-desc">Formato ODF aperto e interoperabile. Conforme al decreto per la contabilità dei lavori pubblici. Nessuna dipendenza da software proprietario.</p>
            </div>
        </div>

        <!-- Screenshot -->
        <div class="screenshot-wrap">
            <div class="screenshot-frame">
                <img src="<?php echo esc_url(get_template_directory_uri() . '/assets/images/leeno-screenshot.jpg'); ?>" alt="Interfaccia LeenO su LibreOffice Calc" class="screenshot-img">
            </div>
            <div class="screenshot-caption">
                <span class="caption-text">Interfaccia LeenO su LibreOffice Calc</span>
            </div>
        </div>
    </div>
</section>

<!-- METRICS SECTION -->
<section class="metrics-section" id="metricsSection">
    <canvas class="cad-canvas cad-curved" id="cadCanvasCurved"></canvas>
    <div class="container relative">
        <h2 class="metrics-heading">
            Numeri che <span class="text-accent">contano</span>
        </h2>

        <div class="metrics-grid">
            <!-- 1. Ultima versione -->
            <div class="metric-item metric-item--text">
                <div class="metric-number metric-number--text">
                    <span class="metric-text-value">3.25</span>
                </div>
                <span class="metric-label">Ultima Versione</span>
            </div>

            <!-- 2. Iscritti newsletter (MailPoet) -->
            <div class="metric-item" data-target="<?php
                global $wpdb;
                $nc = intval( $wpdb->get_var( "SELECT COUNT(*) FROM {$wpdb->prefix}mailpoet_subscribers WHERE status = 'subscribed'" ) );
                echo $nc > 0 ? $nc : 382;
            ?>">
                <div class="metric-number">
                    <span class="flip-stat">0</span><span class="metric-suffix">+</span>
                </div>
                <span class="metric-label">Iscritti alla Newsletter</span>
            </div>

            <!-- 3. Anni di sviluppo -->
            <div class="metric-item" data-target="15">
                <div class="metric-number">
                    <span class="flip-stat">0</span><span class="metric-suffix">+</span>
                </div>
                <span class="metric-label">Anni di Sviluppo</span>
            </div>

            <!-- 4. FOSS -->
            <div class="metric-item metric-item--text">
                <div class="metric-number metric-number--text">
                    <span class="metric-text-value">100%</span>
                </div>
                <span class="metric-label">
                    <a href="https://it.wikipedia.org/wiki/Free_and_Open_Source_Software" target="_blank" rel="noopener" class="metric-foss-link">Free &amp; Open Source</a>
                    &mdash; Licenza LGPL
                </span>
            </div>
        </div>

        <!-- Prezzari preview -->
        <div class="prezzari-section">
            <h3 class="prezzari-title">
                <a href="https://leeno.org/prezziari/download-prezziari/" class="prezzari-title-link">
                    Prezzari <span class="text-accent">Disponibili</span>
                    <svg width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.5" aria-hidden="true" style="vertical-align:middle;margin-left:6px;opacity:0.6"><line x1="5" y1="12" x2="19" y2="12"/><polyline points="12 5 19 12 12 19"/></svg>
                </a>
            </h3>
            <div class="prezzari-grid">
                <?php
                $prezzari_query = new WP_Query([
                    'posts_per_page' => 8,
                    'post_status'    => 'publish',
                    'orderby'        => 'date',
                    'order'          => 'DESC',
                    'category_name'  => 'elenchi-prezzi',  // slug categoria Prezzari
                ]);
                if ( $prezzari_query->have_posts() ) :
                    while ( $prezzari_query->have_posts() ) : $prezzari_query->the_post();
                ?>
                <a href="<?php the_permalink(); ?>" class="prezzari-tag">
                    <?php the_title(); ?>
                </a>
                <?php
                    endwhile;
                    wp_reset_postdata();
                else : ?>
                <p style="color:rgba(255,255,255,0.4);font-size:0.85rem;">Nessun prezzario trovato.</p>
                <?php endif; ?>
            </div>
        </div>
    </div>
</section>

<!-- BLOG POSTS SECTION -->
<section class="blog-section" id="blogSection">
    <div class="container">
        <h2 class="section-title section-title-dark">
            <span class="title-word-outer"><span class="title-word-inner">ULTIME</span></span>
            <span class="title-word-outer"><span class="title-word-inner">NOVIT&Agrave;</span></span>
        </h2>

        <div class="posts-grid">
            <?php
            $latest = new WP_Query([
                'posts_per_page' => 6,
                'post_status'    => 'publish',
            ]);
            if ($latest->have_posts()) :
                while ($latest->have_posts()) : $latest->the_post();
            ?>
            <article class="post-card<?php echo !has_post_thumbnail() ? ' post-card--no-thumb' : ''; ?>">
                <?php if (has_post_thumbnail()) : ?>
                <a href="<?php the_permalink(); ?>" class="post-thumb">
                    <?php the_post_thumbnail('leeno-card', ['class' => 'post-img']); ?>
                </a>
                <?php else : ?>
                <a href="<?php the_permalink(); ?>" class="post-thumb post-thumb--placeholder" aria-hidden="true">
                    <span class="post-thumb-cat">
                        <?php
                        $cats = get_the_category();
                        echo $cats ? esc_html($cats[0]->name) : 'LeenO';
                        ?>
                    </span>
                </a>
                <?php endif; ?>
                <div class="post-body">
                    <div class="post-meta">
                        <span class="post-cat">
                            <?php
                            $cats = get_the_category();
                            echo $cats ? esc_html($cats[0]->name) : 'Blog';
                            ?>
                        </span>
                        <span class="post-date"><?php echo get_the_date('j F Y'); ?></span>
                    </div>
                    <h3 class="post-title">
                        <a href="<?php the_permalink(); ?>"><?php the_title(); ?></a>
                    </h3>
                    <p class="post-excerpt"><?php echo wp_trim_words(get_the_excerpt(), 18); ?></p>
                    <a href="<?php the_permalink(); ?>" class="post-read">Leggi tutto &rarr;</a>
                </div>
            </article>
            <?php
                endwhile;
                wp_reset_postdata();
            else :
            ?>
            <p class="no-posts">Nessun articolo disponibile.</p>
            <?php endif; ?>
        </div>

        <div class="view-all">
            <a href="<?php echo esc_url(home_url('/category/blog/')); ?>" class="btn-outline">
                Vedi tutti gli articoli
            </a>
        </div>
    </div>
</section>

</main><!-- #main-content -->

<?php get_footer(); ?>
