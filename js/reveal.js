// Scroll reveal - IntersectionObserver pattern
const reveals = document.querySelectorAll('.reveal');
const observer = new IntersectionObserver((entries) => {
  entries.forEach(e => {
    if (e.isIntersecting) {
      e.target.classList.add('visible');
      observer.unobserve(e.target);
    }
  });
}, { threshold: 0.08 });
reveals.forEach(el => observer.observe(el));

// Trigger hero reveals immediately on load
window.addEventListener('load', () => {
  // Check for either .page-hero or .hero parent
  const heroReveals = document.querySelectorAll('.page-hero .reveal, .hero .reveal');
  heroReveals.forEach((el, i) => {
    setTimeout(() => el.classList.add('visible'), 100 + i * 130);
  });
});
