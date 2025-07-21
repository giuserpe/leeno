def _generate_versions_html(self, version_info: Dict[str, str], build_info: Dict[str, str]) -> None:
    """Genera la pagina HTML completa con i file dal WebDAV"""
    try:
        # 1. Recupera i file .oxt dal WebDAV
        oxt_files = list_webdav_oxt_files()
        
        # 2. Recupera i tag GitHub
        github_tags = get_github_tags()
        
        # 3. Genera HTML dinamico
        html_content = f"""<!DOCTYPE html>
        <html lang="it">
        <head>
            <meta charset="UTF-8">
            <title>Versioni LeenO - {version_info['full']}</title>
            <style>
                /* Stili migliorati... */
                .oxt-file {{ background-color: #f8f9fa; margin: 10px 0; padding: 15px; }}
            </style>
        </head>
        <body>
            <h1>Versione Corrente: {version_info['full']}</h1>
            
            <!-- Sezione versione attuale -->
            <div class="version-info">
                <table>
                    <tr><th>Build</th><td>{build_info['build_number']}</td></tr>
                    <tr><th>Data</th><td>{build_info['build_date']}</td></tr>
                    <tr><th>Commit</th><td>{build_info['git_sha']}</td></tr>
                </table>
            </div>
            
            <!-- Sezione file disponibili -->
            <h2>File disponibili sul WebDAV</h2>
            <div class="oxt-files">
                {''.join(
                    f'<div class="oxt-file">'
                    f'<h3>{file["filename"]}</h3>'
                    f'<p>Size: {format_size(file["size"])} | '
                    f'Modified: {format_date(file["lastmod"])}</p>'
                    f'<a href="{file["public_url"]}" download>Download</a>'
                    f'</div>'
                    for file in oxt_files[:10]  # Mostra solo gli ultimi 10
                )}
            </div>
        </body>
        </html>
        """
        
        with open(self.tools_version_dir / 'versions.html', 'w') as f:
            f.write(html_content)
            
    except Exception as e:
        logger.error(f"Errore generazione HTML: {str(e)}")
        raise