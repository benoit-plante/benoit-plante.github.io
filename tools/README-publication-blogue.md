# Pipeline de publication du blogue

Ce dossier contient l'outillage qui publie un article de fond chaque lundi sur `benoitplante.com/blogue/`. Tu prépares un `.docx`, le système gère la conversion, la mise à jour de l'index, le sitemap et le commit Git.

## Cycle en bref

1. Tu rédiges un article dans Word (voir convention ci-bas).
2. Tu déposes le `.docx` dans `articles-de-fond/_a-publier/` avec un nom au format `YYYY-MM-DD-slug.docx` (la date = date de publication souhaitée, un lundi typiquement).
3. La tâche planifiée `publication-blogue-lundi` se déclenche chaque lundi à 7h, prend le plus ancien article dont la date est passée, le convertit en page HTML stylisée, met à jour `blogue/index.html` + `sitemap.xml`, archive le source et fait un commit + push Git.
4. GitHub Pages redéploie le site dans la minute qui suit.
5. Tu reçois une notification avec l'URL à inclure dans l'infolettre.

## Convention de nommage

```
articles-de-fond/_a-publier/2026-05-18-loi-25-ia-recherche.docx
                            └────────┘ └───────────────────┘
                            date publi          slug URL
```

→ URL finale : `https://benoitplante.com/blogue/loi-25-ia-recherche.html`

## Structure attendue du .docx

| Para # | Contenu | Style Word |
|--------|---------|------------|
| 1 | "Article de fond" | Normal |
| 2 | **Titre principal** | Normal |
| 3 | Sous-titre / dek | Normal |
| 4 | Signature et date | Normal |
| 5 | (vide) | Normal |
| 6 | Chapeau (devient la méta-description et l'intro mise en italique) | Normal |
| 7+ | Corps de l'article | voir ci-bas |

**Styles à utiliser dans le corps :**

- `Heading 1` → grand titre de section (devient `<h2>` + entrée dans le sommaire)
- `Heading 2` → sous-titre (devient `<h3>`)
- `Normal` → paragraphe
- `List Paragraph` → item de liste. Le script regroupe les items consécutifs et choisit automatiquement `<ol>` (si le premier item commence par "Étape N", "Phase N" ou un nombre) ou `<ul>` (sinon)
- **Tableau Word à 1 cellule** → encadré stylisé (idéal pour cas d'usage, exemples, mises en garde)
- **Tableau Word multi-cellules** → table HTML avec en-tête (première ligne)

**Section finale obligatoire :** `Heading 2` intitulé "Sources et références" (ou "Sources" ou "Bibliographie"). Tout `List Paragraph` qui suit est traité comme bibliographie. Les URLs sont automatiquement liées.

## Commandes

```bash
# Mode auto — prend le plus ancien article dû
python tools/publish_article.py

# Simulation
python tools/publish_article.py --dry-run

# Forcer un .docx précis
python tools/publish_article.py --file articles-de-fond/_a-publier/2026-05-18-loi-25-ia-recherche.docx

# Sans archiver la source (test)
python tools/publish_article.py --no-archive --file ...

# Régénérer index.html + sitemap depuis le manifest (utile après édition manuelle)
python tools/publish_article.py --rebuild-index
```

## Structure des dossiers

```
benoit-plante.github.io/         (repo Git, déployé sur GitHub Pages)
├── blogue/
│   ├── index.html               ← page liste (régénérée à chaque publication)
│   ├── article.css              ← styles dédiés au blogue
│   ├── manifest.json            ← source de vérité des billets
│   └── {slug}.html              ← une page par article publié
├── tools/
│   ├── publish_article.py       ← le pipeline
│   └── README-publication-blogue.md
├── articles-de-fond/            ← EXCLU du Git (.gitignore)
│   ├── _a-publier/              ← file d'attente
│   ├── _publies/                ← archive
│   └── _brouillons/             ← WIP
└── sitemap.xml                  ← section <!-- BLOGUE_START --> ... <!-- BLOGUE_END --> auto-gérée
```

## Publication programmée (« preloading »)

L'index `blogue/index.html` et le sitemap **filtrent automatiquement par date** : un article dont la date de publication est dans le futur reste invisible publiquement (pas de carte sur l'index, pas d'URL dans le sitemap).

Tu peux donc « précharger » plusieurs articles à l'avance :
1. Génère les pages `.html` à partir des `.docx` (via `python tools/publish_article.py --file ...`)
2. Les pages existent dans `blogue/` mais ne sont pas listées tant que la date est future
3. Chaque lundi, `--rebuild-index` (exécuté par la tâche planifiée) régénère l'index avec les articles devenus dus

C'est ce qui est en place pour les 4 premiers articles : ils sortent les 18 mai, 25 mai, 1er juin et 8 juin.

## Articulation avec l'infolettre

- Lundi 7h → publication automatique
- Lundi matin → notification avec URL
- Mardi (ou jour fixe) → l'infolettre hebdo intègre le lien vers l'article

Le pied de chaque article contient déjà un CTA vers `/infolettre/` pour la conversion lecteur → abonné.

## Dépannage

| Symptôme | Cause | Solution |
|----------|-------|----------|
| "Aucun article du aujourd'hui" | File vide ou dates futures | Ajouter un .docx daté |
| "structure trop courte" | .docx ne suit pas la convention | Vérifier les 4 paragraphes d'en-tête |
| "titre vide" | Le 2ᵉ paragraphe est blanc | Mettre le titre au bon endroit |
| Encadré attendu mais non rendu | Texte normal au lieu d'une table 1×1 | Insérer un tableau Word à 1 cellule |
| Liste non détectée comme numérotée | Premier item ne commence pas par "Étape N" ou un nombre | Modifier le 1er item ("Étape 1 —" ou "1.") |
| Publication locale OK mais site pas à jour | Push Git a échoué | Vérifier `git status` et `git push origin main` manuellement |

## Personnalisation

- **Heure de publication** : modifier le cron `0 7 * * 1` de la tâche `publication-blogue-lundi`
- **Style des pages** : `blogue/article.css`
- **Gabarit HTML** : fonctions `page_html()` et `regenerer_index()` dans `publish_article.py`
- **Vitesse de lecture** : 220 mots/min par défaut (propriété `Article.temps_lecture`)
