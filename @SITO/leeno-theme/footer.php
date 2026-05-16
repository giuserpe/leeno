<footer class="site-footer" id="siteFooter">
    <div class="footer-cta">
        <div class="container">
            <div class="footer-links-grid">
                <a href="<?php echo esc_url(home_url('/about-leeno/leeno/download/')); ?>" class="footer-cta-card">
                    <div class="cta-card-header">
                        <svg width="24" height="24" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><path d="M21 15v4a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2v-4"/><polyline points="7 10 12 15 17 10"/><line x1="12" y1="15" x2="12" y2="3"/></svg>
                        <span class="cta-title">SCARICA LEENO</span>
                        <svg width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><path d="M18 13v6a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2V8a2 2 0 0 1 2-2h6"/><polyline points="15 3 21 3 21 9"/><line x1="10" y1="14" x2="21" y2="3"/></svg>
                    </div>
                    <span class="cta-desc">v3.25.0 &mdash; LibreOffice Extension</span>
                </a>

                <a href="<?php echo esc_url(home_url('/category/documentazione/')); ?>" class="footer-cta-card">
                    <div class="cta-card-header">
                        <svg width="24" height="24" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><path d="M2 3h6a4 4 0 0 1 4 4v14a3 3 0 0 0-3-3H2z"/><path d="M22 3h-6a4 4 0 0 0-4 4v14a3 3 0 0 1 3-3h7z"/></svg>
                        <span class="cta-title">DOCUMENTAZIONE</span>
                        <svg width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><path d="M18 13v6a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2V8a2 2 0 0 1 2-2h6"/><polyline points="15 3 21 3 21 9"/><line x1="10" y1="14" x2="21" y2="3"/></svg>
                    </div>
                    <span class="cta-desc">Manuale, guide e tutorial</span>
                </a>

                <a href="<?php echo esc_url(home_url('/forums/')); ?>" class="footer-cta-card">
                    <div class="cta-card-header">
                        <svg width="24" height="24" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><path d="M21 15a2 2 0 0 1-2 2H7l-4 4V5a2 2 0 0 1 2-2h14a2 2 0 0 1 2 2z"/></svg>
                        <span class="cta-title">FORUM</span>
                        <svg width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><path d="M18 13v6a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2V8a2 2 0 0 1 2-2h6"/><polyline points="15 3 21 3 21 9"/><line x1="10" y1="14" x2="21" y2="3"/></svg>
                    </div>
                    <span class="cta-desc">Community e supporto</span>
                </a>
            </div>

            <div class="footer-big-cta">
                <a href="<?php echo esc_url(home_url('/about-leeno/leeno/download/')); ?>" class="btn-big">
                    Download LeenO 3.25.0
                </a>
                <span class="btn-sub">LibreOffice Extension (.oxt) &mdash; Windows, macOS, Linux</span>
            </div>
        </div>
    </div>

    <div class="footer-bottom">
        <div class="container">
            <div class="footer-row">
                <div class="footer-brand">
                    <img src="<?php echo esc_url(get_template_directory_uri() . '/assets/images/logo-leeno.png'); ?>" alt="LeenO" class="footer-logo">
                    <span class="footer-license">Open Source &mdash; GPL v3</span>
                </div>

                <div class="footer-external">
                    <a href="https://www.libreoffice.org/" target="_blank" rel="noopener">LibreOffice</a>
                    <a href="https://extensions.libreoffice.org/extensions/leeno-2" target="_blank" rel="noopener">Extensions</a>
                    <a href="https://gitlab.com/giuserpe/leeno" target="_blank" rel="noopener">GitLab</a>
                    <a href="https://github.com/giuserpe/leeno" target="_blank" rel="noopener">GitHub</a>
                    <a href="https://t.me/leeno_computometrico" target="_blank" rel="noopener">Telegram</a>
                    <a href="https://m.facebook.com/groups/433206393972197" target="_blank" rel="noopener">Facebook</a>
                    <a href="https://leeno.org/donazioni/" target="_blank" rel="noopener">Dona!</a>
                </div>
            </div>

            <div class="footer-copyright">
                <p>&copy; <?php echo date('Y'); ?> LeenO.org &mdash; Giuseppe Vizziello &amp; Contributors. Software libero per il computo metrico estimativo.</p>
            </div>
        </div>
    </div>
</footer>

<?php wp_footer(); ?>
</body>
</html>
