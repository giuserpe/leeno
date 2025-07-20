function leeno_versions_shortcode() {
  return '<iframe src="https://leeno.org/versions/versions.html" width="100%" height="800" frameborder="0" style="border:0; overflow:auto;"></iframe>';
}
add_shortcode('leenoversions', 'leeno_versions_shortcode');
function leeno_versions_enqueue_scripts() {
  wp_enqueue_style('leeno-versions-style', get_template_directory_uri() . '/css/leeno-versions.css');
}