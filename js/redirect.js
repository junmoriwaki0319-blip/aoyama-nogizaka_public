// GitHub Pages clean URL redirect: /team → /team.html
(function(){
  var defined = ['team','privacy','risk-assessment','activist-dashboard','activist-screener','food-service','saas'];
  var path = location.pathname.replace(/\/$/,'');
  var slug = path.substring(1);
  if (defined.indexOf(slug) !== -1) {
    location.replace('/' + slug + '.html');
  }
})();
