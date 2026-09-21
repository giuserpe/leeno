<?php get_header(); ?>

<main class="main-content archive-content">
    <div class="page-header">
        <div class="container">
            <h1 class="page-title">
                <?php
                if (is_category()) {
                    single_cat_title('Categoria: ');
                } elseif (is_tag()) {
                    single_tag_title('Tag: ');
                } elseif (is_author()) {
                    echo 'Autore: ' . get_the_author();
                } elseif (is_date()) {
                    echo 'Archivio: ' . get_the_date('F Y');
                } else {
                    post_type_archive_title();
                }
                ?>
            </h1>
            <?php if (is_category() || is_tag()) {
                $desc = term_description();
                if ($desc) echo '<div class="page-desc">' . $desc . '</div>';
            } ?>
        </div>
    </div>

    <div class="container">
        <div class="content-layout">
            <div class="content-main">
                <div class="posts-grid">
                    <?php
                    if (have_posts()) :
                        while (have_posts()) : the_post();
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
                    else :
                    ?>
                    <p class="no-posts">Nessun articolo trovato.</p>
                    <?php endif; ?>
                </div>

                <div class="pagination">
                    <?php
                    echo paginate_links([
                        'prev_text' => '&larr; Precedente',
                        'next_text' => 'Successiva &rarr;',
                    ]);
                    ?>
                </div>
            </div>

            <?php if (is_active_sidebar('sidebar-blog')) : ?>
            <aside class="content-sidebar">
                <?php dynamic_sidebar('sidebar-blog'); ?>
            </aside>
            <?php endif; ?>
        </div>
    </div>
</main>

<?php get_footer(); ?>
