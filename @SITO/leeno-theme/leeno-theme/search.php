<?php get_header(); ?>

<main id="main-content" class="main-content page-content">

    <div class="page-header">
        <div class="container">
            <nav class="breadcrumbs" aria-label="Percorso di navigazione">
                <a href="<?php echo esc_url(home_url('/')); ?>">Home</a>
                <span class="sep" aria-hidden="true">›</span>
                <span class="current">Ricerca</span>
            </nav>
            <h1 class="page-title">
                <?php if ( get_search_query() ) : ?>
                    Risultati per: <span style="color:var(--accent-cyan)"><?php echo esc_html(get_search_query()); ?></span>
                <?php else : ?>
                    Ricerca
                <?php endif; ?>
            </h1>
            <?php if ( have_posts() ) : ?>
            <p class="page-desc" style="color:rgba(255,255,255,0.5);font-size:0.85rem;margin-top:8px;">
                <?php
                global $wp_query;
                $count = $wp_query->found_posts;
                printf(
                    _n('%s risultato trovato', '%s risultati trovati', $count, 'leeno-dm'),
                    number_format_i18n($count)
                );
                ?>
            </p>
            <?php endif; ?>
        </div>
    </div>

    <div class="container" style="padding-top:40px;padding-bottom:80px;">

        <?php if ( have_posts() ) : ?>

            <div class="search-results">
                <?php while ( have_posts() ) : the_post(); ?>
                <article class="search-result-item">
                    <div class="search-result-meta">
                        <?php
                        $cats = get_the_category();
                        if ( $cats ) echo '<span class="search-result-cat">' . esc_html($cats[0]->name) . '</span>';
                        ?>
                        <span class="search-result-date"><?php echo get_the_date('d M Y'); ?></span>
                    </div>
                    <h2 class="search-result-title">
                        <a href="<?php the_permalink(); ?>"><?php the_title(); ?></a>
                    </h2>
                    <?php if ( get_the_excerpt() ) : ?>
                    <p class="search-result-excerpt"><?php the_excerpt(); ?></p>
                    <?php endif; ?>
                    <a href="<?php the_permalink(); ?>" class="search-result-link">
                        Leggi
                        <svg width="12" height="12" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.5" aria-hidden="true"><line x1="5" y1="12" x2="19" y2="12"/><polyline points="12 5 19 12 12 19"/></svg>
                    </a>
                </article>
                <?php endwhile; ?>
            </div>

            <div class="search-pagination">
                <?php
                the_posts_pagination([
                    'prev_text' => '← Precedente',
                    'next_text' => 'Successivo →',
                ]);
                ?>
            </div>

        <?php else : ?>

            <div class="search-no-results">
                <p>Nessun risultato per <strong><?php echo esc_html(get_search_query()); ?></strong>.</p>
                <p style="margin-top:8px;color:var(--text-secondary);font-size:0.9rem;">Prova con termini diversi o controlla l'ortografia.</p>
                <form role="search" method="get" action="<?php echo esc_url(home_url('/')); ?>" style="margin-top:24px;display:flex;gap:8px;max-width:400px;">
                    <input type="search" name="s" placeholder="Nuova ricerca…" value="<?php echo esc_attr(get_search_query()); ?>" style="flex:1;background:rgba(255,255,255,0.08);border:1px solid rgba(255,255,255,0.2);color:#fff;padding:10px 14px;font-size:0.9rem;outline:none;">
                    <button type="submit" style="background:var(--accent-rust);color:var(--bg-dark);font-family:var(--font-display);font-weight:700;font-size:0.8rem;text-transform:uppercase;padding:10px 20px;border:none;cursor:pointer;">Cerca</button>
                </form>
            </div>

        <?php endif; ?>

    </div>
</main>

<?php get_footer(); ?>
