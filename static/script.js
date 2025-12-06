// main.js

document.addEventListener('DOMContentLoaded', () => {
  // Highlight active nav link
  const path = window.location.pathname;
  document.querySelectorAll('nav a.nav-link').forEach(link => {
    if (link.getAttribute('href') === path) {
      link.classList.add('active');
    }
  });

  // Auto-dismiss flash messages after 5 seconds
  setTimeout(() => {
    document.querySelectorAll('.alert').forEach(alert => {
      // Use Bootstrap’s JS API to close
      bootstrap.Alert.getOrCreateInstance(alert).close();
    });
  }, 5000);
});
