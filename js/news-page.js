const GAS_URL = 'https://script.google.com/macros/s/AKfycbwQby968Ts6rSFyzTZZFfblfFDYOogRPWYa7kYl5eHdXaemLIilu6FnOemvX5ygIBwt/exec';

/* ── Filter (event delegation) ── */
function filterArticles(category, btn) {
  document.querySelectorAll('.filter-btn').forEach(b => b.classList.remove('active'));
  btn.classList.add('active');

  const cards = document.querySelectorAll('#articles-grid .article-card:not(.coming-soon)');
  const featured = document.querySelector('.featured-wrap');
  let visible = 0;

  cards.forEach(card => {
    const match = category === 'all' || card.dataset.category === category;
    card.style.display = match ? '' : 'none';
    if (match) visible++;
  });

  if (featured) {
    const fMatch = category === 'all' || featured.dataset.category === category;
    featured.style.display = fMatch ? '' : 'none';
  }

  document.getElementById('article-count').textContent = visible + '件';
}

async function submitNewsletter(e) {
  e.preventDefault();
  const form = document.getElementById('newsletterForm');
  const btn = document.getElementById('nlSubmitBtn');
  const status = document.getElementById('nlStatus');
  const email = form.querySelector('input[name="email"]').value;
  btn.disabled = true; btn.textContent = '登録中...';
  status.textContent = ''; status.style.color = '';
  try {
    await fetch(GAS_URL, { method: 'POST', body: JSON.stringify({ _formType: 'newsletter', email }), mode: 'no-cors' });
    status.textContent = '登録が完了しました。確認メールをお送りしました。';
    status.style.color = '#a3e4a3';
    form.reset();
  } catch (err) {
    status.textContent = '登録に失敗しました。再度お試しください。';
    status.style.color = '#ff9999';
  } finally { btn.disabled = false; btn.textContent = '登録'; }
  return false;
}

/* ── Event Delegation ── */
document.addEventListener('click', function(e) {
  const filterBtn = e.target.closest('[data-filter]');
  if (filterBtn) {
    filterArticles(filterBtn.dataset.filter, filterBtn);
  }
});

document.addEventListener('submit', function(e) {
  const form = e.target.closest('[data-form-action]');
  if (form && form.dataset.formAction === 'submitNewsletter') {
    submitNewsletter(e);
  }
});
