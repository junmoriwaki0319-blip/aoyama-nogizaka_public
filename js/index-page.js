const GAS_URL = 'https://script.google.com/macros/s/AKfycbwQby968Ts6rSFyzTZZFfblfFDYOogRPWYa7kYl5eHdXaemLIilu6FnOemvX5ygIBwt/exec';

async function submitContactForm(e) {
  e.preventDefault();
  const form = document.getElementById('contactForm');
  const btn = document.getElementById('contactSubmitBtn');
  const status = document.getElementById('contactStatus');
  const fd = new FormData(form);
  const data = { _formType: 'contact' };
  fd.forEach((v, k) => { data[k] = v; });
  btn.disabled = true; btn.textContent = '送信中...';
  status.className = 'form-status'; status.style.display = 'none';
  try {
    await fetch(GAS_URL, { method: 'POST', body: JSON.stringify(data), mode: 'no-cors' });
    status.className = 'form-status success'; status.textContent = 'お問い合わせを受け付けました。確認メールをお送りしましたのでご確認ください。';
    form.reset();
  } catch (err) {
    status.className = 'form-status error'; status.textContent = '送信に失敗しました。時間をおいて再度お試しください。';
  } finally { btn.disabled = false; btn.textContent = '送信する'; }
  return false;
}

/* ── Event Delegation: form submit ── */
document.addEventListener('submit', function(e) {
  var form = e.target.closest('[data-form-action]');
  if (form && form.dataset.formAction === 'submitContactForm') {
    submitContactForm(e);
  }
});
