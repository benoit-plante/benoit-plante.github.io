/* ==========================================================================
   Benoit Plante — Script partagé
   ========================================================================== */

(function () {
  'use strict';

  /* ----------------------------------------------------------------------
     Menu mobile
     ---------------------------------------------------------------------- */
  const navToggle = document.querySelector('.nav-toggle');
  const navLinks = document.querySelector('.nav-links');

  if (navToggle && navLinks) {
    navToggle.addEventListener('click', () => {
      const isOpen = navLinks.classList.toggle('open');
      navToggle.setAttribute('aria-expanded', isOpen);
    });

    // Fermer au clic sur un lien (mobile)
    navLinks.querySelectorAll('a').forEach(link => {
      link.addEventListener('click', () => {
        navLinks.classList.remove('open');
        navToggle.setAttribute('aria-expanded', 'false');
      });
    });
  }

  /* ----------------------------------------------------------------------
     Reveal au scroll
     ---------------------------------------------------------------------- */
  const reveals = document.querySelectorAll('.reveal');
  if (reveals.length && 'IntersectionObserver' in window) {
    const io = new IntersectionObserver(
      entries => {
        entries.forEach(e => {
          if (e.isIntersecting) {
            e.target.classList.add('visible');
            io.unobserve(e.target);
          }
        });
      },
      { threshold: 0.1 }
    );
    reveals.forEach(el => io.observe(el));
  } else {
    reveals.forEach(el => el.classList.add('visible'));
  }

  /* ----------------------------------------------------------------------
     Filtres de la bibliothèque de ressources
     ---------------------------------------------------------------------- */
  const filterContainer = document.querySelector('[data-filter-container]');
  if (filterContainer) {
    const cards = filterContainer.querySelectorAll('.resource-card');
    const chips = document.querySelectorAll('.filter-chip');
    const countEl = document.querySelector('[data-results-count]');
    const emptyEl = document.querySelector('[data-empty-state]');

    const state = { categorie: 'tous', prix: 'tous' };

    function applyFilters() {
      let visible = 0;
      cards.forEach(card => {
        const cat = card.getAttribute('data-categorie') || '';
        const prix = card.getAttribute('data-prix') || '';
        const showCat = state.categorie === 'tous' || cat === state.categorie;
        const showPrix = state.prix === 'tous' || prix === state.prix;
        const show = showCat && showPrix;
        card.style.display = show ? '' : 'none';
        if (show) visible++;
      });

      if (countEl) {
        countEl.textContent =
          visible === 0
            ? 'Aucune ressource trouvée'
            : visible === 1
            ? '1 ressource'
            : `${visible} ressources`;
      }

      if (emptyEl) {
        emptyEl.style.display = visible === 0 ? 'block' : 'none';
      }
    }

    chips.forEach(chip => {
      chip.addEventListener('click', () => {
        const group = chip.getAttribute('data-group');
        const value = chip.getAttribute('data-value');
        if (!group || !value) return;

        // Désactiver les autres du même groupe
        document
          .querySelectorAll(`.filter-chip[data-group="${group}"]`)
          .forEach(c => c.classList.remove('active'));
        chip.classList.add('active');

        state[group] = value;
        applyFilters();
      });
    });

    applyFilters();
  }

  /* ----------------------------------------------------------------------
     Formulaire de contact — Web3Forms (soumission AJAX)
     ---------------------------------------------------------------------- */
  document.querySelectorAll('form[data-form-type="contact"]').forEach(form => {
    form.addEventListener('submit', async e => {
      e.preventDefault();

      const button = form.querySelector('button[type="submit"]');
      const confirm = form.querySelector('.form-confirm');
      const originalText = button ? button.textContent : '';

      // Honeypot : si le bot a coché la case cachée, on simule un succès et on stoppe
      const honeypot = form.querySelector('input[name="botcheck"]');
      if (honeypot && honeypot.checked) {
        if (confirm) {
          confirm.textContent = '✓ Merci !';
          confirm.classList.add('show');
        }
        form.reset();
        return;
      }

      if (button) {
        button.disabled = true;
        button.textContent = 'Envoi en cours…';
      }
      if (confirm) {
        confirm.classList.remove('show');
        confirm.classList.remove('error');
      }

      try {
        const formData = new FormData(form);
        const response = await fetch('https://api.web3forms.com/submit', {
          method: 'POST',
          headers: { 'Accept': 'application/json' },
          body: formData
        });
        const data = await response.json();

        if (data.success) {
          if (confirm) {
            confirm.textContent = '✓ Merci ! Votre message a bien été envoyé. Je vous répondrai dans les meilleurs délais.';
            confirm.classList.add('show');
          }
          form.reset();
        } else {
          throw new Error(data.message || 'Erreur lors de l\'envoi');
        }
      } catch (err) {
        if (confirm) {
          confirm.textContent = '⚠ Une erreur est survenue. Réessayez ou contactez-moi via LinkedIn.';
          confirm.classList.add('show', 'error');
        }
        console.error('Erreur Web3Forms:', err);
      } finally {
        if (button) {
          button.disabled = false;
          button.textContent = originalText;
        }
      }
    });
  });

  /* ----------------------------------------------------------------------
     Callback de succès MailerLite (déclenché par webforms.min.js)
     Form ID: 41268111 — partagé entre toutes les pages avec inscription
     ---------------------------------------------------------------------- */
  window.ml_webform_success_41268111 = function () {
    var $ = window.ml_jQuery || window.jQuery;
    if ($) {
      $('.ml-subscribe-form-41268111 .row-success').show();
      $('.ml-subscribe-form-41268111 .row-form').hide();
    } else {
      document.querySelectorAll('.ml-subscribe-form-41268111 .row-success').forEach(el => el.style.display = 'block');
      document.querySelectorAll('.ml-subscribe-form-41268111 .row-form').forEach(el => el.style.display = 'none');
    }
  };

  // Tracking pixel MailerLite — déclenché si un formulaire est présent
  if (document.querySelector('.ml-subscribe-form-41268111')) {
    fetch('https://assets.mailerlite.com/jsonp/2346008/forms/187388314435716606/takel').catch(() => {});
  }

  /* ----------------------------------------------------------------------
     Mise en évidence du lien actif dans la nav
     ---------------------------------------------------------------------- */
  const currentPage = (window.location.pathname.split('/').pop() || 'index.html').toLowerCase();
  document.querySelectorAll('.nav-links a').forEach(link => {
    const href = (link.getAttribute('href') || '').toLowerCase();
    if (href === currentPage || (currentPage === '' && href === 'index.html')) {
      link.classList.add('active');
    }
  });
})();
