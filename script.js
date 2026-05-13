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
     Formulaire de contact (placeholder — à remplacer par Formspree/Web3Forms)
     ---------------------------------------------------------------------- */
  document.querySelectorAll('form[data-form-type="contact"]').forEach(form => {
    form.addEventListener('submit', e => {
      e.preventDefault();
      const confirm = form.querySelector('.form-confirm');
      if (confirm) confirm.classList.add('show');
      form.reset();
      // TODO: Remplacer par appel Formspree / Web3Forms
      // ex: fetch('https://formspree.io/f/XXXXX', { method: 'POST', body: new FormData(form) })
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
