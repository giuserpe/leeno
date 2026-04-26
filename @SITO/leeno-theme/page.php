<?php get_header(); ?>

<?php
// Determina se siamo nella pagina Prezzari
$is_prezzari = has_term('prezzari', 'category') || stripos(get_the_title(), 'prezzari') !== false || is_page('prezzari');
$page_class   = $is_prezzari ? 'page-prezzari' : '';
?>

<main id="main-content" class="main-content page-content <?php echo esc_attr($page_class); ?>">
    <div class="page-header">
        <div class="container">

            <?php /* Breadcrumb — compatibile con Yoast, RankMath o fallback */ ?>
            <nav class="breadcrumbs" aria-label="Percorso di navigazione">
                <?php if (function_exists('yoast_breadcrumb')) :
                    yoast_breadcrumb('', '');
                elseif (function_exists('rank_math_the_breadcrumbs')) :
                    rank_math_the_breadcrumbs();
                else : ?>
                    <a href="<?php echo esc_url(home_url('/')); ?>">Home</a>
                    <span class="sep" aria-hidden="true">›</span>
                    <span class="current"><?php the_title(); ?></span>
                <?php endif; ?>
            </nav>

            <h1 class="page-title"><?php the_title(); ?></h1>

            <?php if ($is_prezzari) : ?>
            <?php $desc = get_the_excerpt(); ?>
            <?php if ($desc) : ?>
            <p class="page-desc"><?php echo esc_html($desc); ?></p>
            <?php endif; ?>

            <div class="prezzari-search-bar">
                <input
                    type="text"
                    id="prezzari-filter"
                    placeholder="Filtra per regione o anno…"
                    aria-label="Filtra prezzari"
                    autocomplete="off"
                >
            </div>
            <?php endif; ?>

        </div>
    </div>

    <div class="container" style="padding-top: 40px; padding-bottom: 80px;">
        <div class="content-layout">
            <div class="content-main">
                <?php
                while (have_posts()) : the_post();
                    the_content();
                endwhile;
                ?>
            </div>

            <?php if (is_active_sidebar('sidebar-blog') && !$is_prezzari) : ?>
            <aside class="content-sidebar">
                <?php dynamic_sidebar('sidebar-blog'); ?>
            </aside>
            <?php endif; ?>
        </div>
    </div>
</main>

<?php get_footer(); ?>
