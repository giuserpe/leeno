<?php
/**
 * Template Name: Download Software
 *
 * Pagina download software — lista file della categoria.
 * Usa shortcode [wpfilebase tag=list id=38] via plugin LeenO WP Filebase Compatibility
 */

get_header();
?>

<main id="main-content" class="main-content page-content page-software">

    <div class="page-header">
        <div class="container">
            <nav class="breadcrumbs" aria-label="Percorso di navigazione">
                <a href="<?php echo esc_url( home_url('/') ); ?>">Home</a>
                <span class="sep" aria-hidden="true">›</span>
                <span class="current"><?php the_title(); ?></span>
            </nav>
            <h1 class="page-title"><?php the_title(); ?></h1>
        </div>
    </div>

    <div class="container prezzari-container">
        <div class="content-layout">
            <div class="content-main">

            <?php
            $desc = get_the_content();
            $intro = '';
            $iframe_content = '';

            $pos = strpos( $desc, '[advanced_iframe' );
            if ( $pos !== false ) {
                $intro = substr( $desc, 0, $pos );
                $iframe_content = substr( $desc, $pos );
            } else {
                $intro = $desc;
            }

            if ( trim( strip_tags( $intro ) ) || trim( $intro ) ) : ?>
            <div class="software-intro entry-content" style="margin-bottom: 2rem;">
                <?php echo apply_filters( 'the_content', $intro ); ?>
            </div>
            <?php endif; ?>

            <?php
            // Carica file using shortcode (no WP Filebase dependency)
            echo do_shortcode('[wpfilebase tag=list id=38 sort=name]');
            ?>
            </div><!-- .content-main -->

            <?php if ( is_active_sidebar('sidebar-blog') ) : ?>
            <aside class="content-sidebar">
                <?php dynamic_sidebar('sidebar-blog'); ?>
            </aside>
            <?php endif; ?>

        </div><!-- .content-layout -->
    </div><!-- .prezzari-container -->

    <?php if ( ! empty( $iframe_content ) && trim( $iframe_content ) ) : ?>
    <div class="container software-after-table" style="margin-top: 2rem; margin-bottom: 3rem;">
        <?php echo apply_filters( 'the_content', $iframe_content ); ?>
    </div>
    <?php endif; ?>

</main>

<?php get_footer(); ?>
