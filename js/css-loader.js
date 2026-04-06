// css-loader.js — Switch preloaded stylesheets from print to all media
// Replaces inline onload="this.media='all'" for CSP compliance
(function(){
  var links = document.querySelectorAll('link[rel="stylesheet"][media="print"]');
  for (var i = 0; i < links.length; i++) {
    links[i].media = 'all';
  }
})();
