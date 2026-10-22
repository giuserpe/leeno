<?php get_header(); ?>

<main id="main-content" class="main-content single-content">
    <?php while (have_posts()) : the_post(); ?>
    <article class="single-article">
        <div class="single-header">
            <div class="container">
                <div class="single-meta">
                    <span class="single-cat">
                        <?php
                        $cats = get_the_category();
                        echo $cats ? esc_html($cats[0]->name) : 'Blog';
                        ?>
                    </span>
                    <span class="single-date"><?php echo get_the_date('j F Y'); ?></span>
                    <span class="single-author"><?php the_author(); ?></span>
                </div>
                <h1 class="single-title"><?php the_title(); ?></h1>
            </div>
        </div>

        <?php if (has_post_thumbnail()) : ?>
        <div class="single-featured">
            <div class="container">
                <?php the_post_thumbnail('leeno-hero', ['class' => 'single-img']); ?>
            </div>
        </div>
        <?php endif; ?>

        <div class="single-body">
            <div class="container">
                <div class="content-layout">
                    <div class="content-main">
                        <div class="entry-content">
                            <?php the_content(); ?>
                        </div>

                        <nav class="post-navigation" aria-label="<?php esc_attr_e('Navigazione articolo', 'leeno-dm'); ?>">
                            <?php
                            previous_post_link('<span class="nav-prev">&larr; %link</span>');
                            next_post_link('<span class="nav-next">%link &rarr;</span>');
                            ?>
                        </nav>
                    </div>

                    <?php if (is_active_sidebar('sidebar-blog')) : ?>
                    <aside class="content-sidebar" aria-label="<?php esc_attr_e('Barra laterale blog', 'leeno-dm'); ?>">
                        <?php dynamic_sidebar('sidebar-blog'); ?>
                    </aside>
                    <?php endif; ?>
                </div>
            </div>
        </div>
    </article>
    <?php endwhile; ?>
</main>

<?php get_footer(); ?>
