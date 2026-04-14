// Google Analytics: 3秒遅延読み込み
window.addEventListener('load', function() {
  setTimeout(function() {
    var s = document.createElement('script');
    s.src = 'https://www.googletagmanager.com/gtag/js?id=G-C5D57P05PM';
    document.head.appendChild(s);
    var g = document.createElement('script');
    g.src = '/js/gtag.js';
    document.head.appendChild(g);
  }, 3000);
});
