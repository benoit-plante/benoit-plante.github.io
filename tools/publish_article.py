#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
publish_article.py — Pipeline de publication du blogue benoitplante.com

File d'attente : articles-de-fond/_a-publier/YYYY-MM-DD-slug.docx
Sortie         : blogue/{slug}.html + blogue/index.html régénéré
Archive        : articles-de-fond/_publies/

Le dossier articles-de-fond/ est exclu via .gitignore — il ne sera donc
pas déployé sur GitHub Pages. Le script + le manifest restent versionnés.

Convention .docx
----------------
Para 1 : "Article de fond"
Para 2 : Titre
Para 3 : Sous-titre
Para 4 : Signature/date
Para 5+ : Corps. Premier paragraphe non-vide = chapeau.
Suite : Heading 1 (→h2), Heading 2 (→h3), Normal, List Paragraph,
        tables 1×1 (→encadré), tables NxM (→table HTML).
Section finale "Sources et références" (Heading 2 + List Paragraph).

CLI
---
    python tools/publish_article.py                # auto (le plus ancien dû)
    python tools/publish_article.py --file FILE    # forcer un .docx
    python tools/publish_article.py --dry-run      # sans écrire
    python tools/publish_article.py --no-archive   # ne pas déplacer la source
    python tools/publish_article.py --rebuild-index  # régénère index + sitemap
"""
from __future__ import annotations

import argparse, datetime as dt, html, json, re, shutil, sys
from dataclasses import dataclass, field
from pathlib import Path

try:
    from docx import Document
    from docx.oxml.ns import qn
    from docx.text.paragraph import Paragraph
    from docx.table import Table
except ImportError:
    sys.stderr.write("ERREUR : python-docx requis. pip install python-docx\n")
    sys.exit(2)

# ---------------------------------------------------------------------------
SITE_ROOT     = Path(__file__).resolve().parent.parent
QUEUE_DIR     = SITE_ROOT / "articles-de-fond" / "_a-publier"
ARCHIVE_DIR   = SITE_ROOT / "articles-de-fond" / "_publies"
BLOG_DIR      = SITE_ROOT / "blogue"
MANIFEST_PATH = BLOG_DIR / "manifest.json"
SITEMAP_PATH  = SITE_ROOT / "sitemap.xml"

SITE_BASE_URL = "https://benoitplante.com"

MOIS_FR = {1:"janvier",2:"février",3:"mars",4:"avril",5:"mai",6:"juin",
           7:"juillet",8:"août",9:"septembre",10:"octobre",11:"novembre",12:"décembre"}

QUEUE_FILE_RE = re.compile(r"^(\d{4})-(\d{2})-(\d{2})-(.+)\.docx$")

# Favicon SVG inline (cohérent avec le reste du site)
FAVICON = "data:image/svg+xml,%3Csvg xmlns='http://www.w3.org/2000/svg' viewBox='0 0 100 100'%3E%3Crect width='100' height='100' rx='20' fill='%231e3a5f'/%3E%3Ctext x='50' y='66' font-family='Georgia,serif' font-size='52' font-weight='700' text-anchor='middle' fill='%23b8975a'%3EB%3C/text%3E%3C/svg%3E"

# ---------------------------------------------------------------------------
@dataclass
class ArticleBloc:
    type: str
    contenu: object

@dataclass
class Article:
    slug: str
    date_publi: dt.date
    eyebrow: str
    titre: str
    sous_titre: str
    signature: str
    chapeau: str
    blocs: list = field(default_factory=list)
    sources: list = field(default_factory=list)
    nb_mots: int = 0

    @property
    def temps_lecture(self) -> int:
        return max(1, round(self.nb_mots / 220))

    @property
    def date_fr(self) -> str:
        return f"{self.date_publi.day} {MOIS_FR[self.date_publi.month]} {self.date_publi.year}"

    @property
    def date_iso(self) -> str:
        return self.date_publi.isoformat()

# ---------------------------------------------------------------------------
def _iter_corps(doc):
    for child in doc.element.body.iterchildren():
        if child.tag == qn("w:p"):
            yield ("p", Paragraph(child, doc))
        elif child.tag == qn("w:tbl"):
            yield ("table", Table(child, doc))

def lire_docx(path: Path, slug: str, date_publi: dt.date) -> Article:
    doc = Document(str(path))
    items = list(_iter_corps(doc))

    paragraphes_only = [
        (p.style.name if p.style else "Normal", p.text.strip())
        for kind, p in items if kind == "p"
    ]
    if len(paragraphes_only) < 6:
        raise ValueError(f"{path.name} : structure trop courte (< 6 paragraphes)")

    eyebrow    = paragraphes_only[0][1] or "Article de fond"
    titre      = paragraphes_only[1][1]
    sous_titre = paragraphes_only[2][1]
    signature  = paragraphes_only[3][1]

    if not titre:
        raise ValueError(f"{path.name} : titre vide")

    chapeau = ""
    items_corps = []
    p_vus = 0
    for kind, obj in items:
        if kind == "p":
            p_vus += 1
            if p_vus <= 4:
                continue
            texte = obj.text.strip()
            if not chapeau:
                if texte:
                    chapeau = texte
                continue
        items_corps.append((kind, obj))

    if not chapeau:
        raise ValueError(f"{path.name} : chapeau introuvable")

    blocs, sources = [], []
    en_sources = False
    liste_courante = None
    type_liste_courant = "ul"

    def flush():
        nonlocal liste_courante
        if liste_courante:
            blocs.append(ArticleBloc(type=type_liste_courant, contenu=liste_courante))
            liste_courante = None

    nb_mots = len(chapeau.split())

    for kind, obj in items_corps:
        if kind == "table":
            flush()
            if en_sources:
                continue
            rows = [[c.text.strip() for c in r.cells] for r in obj.rows]
            if len(rows) == 1 and len(rows[0]) == 1:
                t = rows[0][0]
                if t:
                    blocs.append(ArticleBloc(type="callout", contenu=t))
                    nb_mots += len(t.split())
            elif rows:
                blocs.append(ArticleBloc(type="table", contenu=rows))
                for r in rows:
                    for c in r:
                        nb_mots += len(c.split())
            continue

        style = obj.style.name if obj.style else "Normal"
        texte = obj.text.strip()
        if not texte:
            flush()
            continue

        if style == "Heading 2" and texte.lower().startswith(("source","référence","bibliographie")):
            flush()
            en_sources = True
            continue

        if en_sources:
            if style == "List Paragraph":
                sources.append(texte)
            continue

        if style == "Heading 1":
            flush()
            blocs.append(ArticleBloc(type="h2", contenu=texte))
            nb_mots += len(texte.split())
        elif style == "Heading 2":
            flush()
            blocs.append(ArticleBloc(type="h3", contenu=texte))
            nb_mots += len(texte.split())
        elif style == "List Paragraph":
            if liste_courante is None:
                liste_courante = []
                if re.match(r"^(étape\s*\d|phase\s*\d|\d+[\.\)]|0\d[\.\)])", texte.lower()):
                    type_liste_courant = "ol"
                else:
                    type_liste_courant = "ul"
            liste_courante.append(texte)
            nb_mots += len(texte.split())
        else:
            flush()
            blocs.append(ArticleBloc(type="p", contenu=texte))
            nb_mots += len(texte.split())

    flush()

    return Article(slug=slug, date_publi=date_publi,
                   eyebrow=eyebrow, titre=titre, sous_titre=sous_titre,
                   signature=signature, chapeau=chapeau,
                   blocs=blocs, sources=sources, nb_mots=nb_mots)

# ---------------------------------------------------------------------------
def esc(s: str) -> str:
    return html.escape(s, quote=True)

def linkify(s: str) -> str:
    return re.sub(r"(https?://[^\s)]+)",
                  r'<a href="\1" target="_blank" rel="noopener">\1</a>', s)

def slug_to_id(s: str) -> str:
    s = s.lower()
    s = re.sub(r"[^a-z0-9]+", "-", s)
    return s.strip("-")

def description_meta(article: Article, max_len: int = 170) -> str:
    desc = article.chapeau.replace("\n", " ").strip()
    if len(desc) <= max_len:
        return desc
    return desc[: max_len - 1].rsplit(" ", 1)[0] + "…"

def rendu_callout(texte: str) -> str:
    premiere = texte.split("\n", 1)[0].strip()
    m = re.match(r"^([A-ZÉÊÀÂÔÎÛÇ][^.\n]{2,80}?)\s*[—\-]\s*(.+)$", premiere)
    if m:
        prefixe = esc(m.group(1).strip())
        suite = esc(m.group(2).strip())
        rest = texte.split("\n", 1)[1].strip() if "\n" in texte else ""
        lines = ["<p><strong>" + prefixe + "</strong> — " + suite + "</p>"]
        if rest:
            for para in rest.split("\n"):
                if para.strip():
                    lines.append("<p>" + esc(para.strip()) + "</p>")
        body = "\n          ".join(lines)
    else:
        body = "\n          ".join(
            "<p>" + esc(p.strip()) + "</p>" for p in texte.split("\n") if p.strip()
        )
    return '        <aside class="article-callout">\n          ' + body + '\n        </aside>'

def rendu_table(rows) -> str:
    out = ['        <div class="article-table-wrap"><table class="article-table">']
    if rows:
        out.append("          <thead><tr>" + "".join("<th>" + esc(c) + "</th>" for c in rows[0]) + "</tr></thead>")
        if len(rows) > 1:
            out.append("          <tbody>")
            for r in rows[1:]:
                out.append("            <tr>" + "".join("<td>" + esc(c) + "</td>" for c in r) + "</tr>")
            out.append("          </tbody>")
    out.append("        </table></div>")
    return "\n".join(out)

def rendu_blocs(blocs) -> str:
    out = []
    for b in blocs:
        if b.type == "h2":
            out.append('        <h2 id="' + slug_to_id(b.contenu) + '">' + esc(b.contenu) + '</h2>')
        elif b.type == "h3":
            out.append("        <h3>" + esc(b.contenu) + "</h3>")
        elif b.type == "p":
            out.append("        <p>" + esc(b.contenu) + "</p>")
        elif b.type in ("ol", "ul"):
            out.append("        <" + b.type + ">")
            for item in b.contenu:
                out.append("          <li>" + esc(item) + "</li>")
            out.append("        </" + b.type + ">")
        elif b.type == "callout":
            out.append(rendu_callout(b.contenu))
        elif b.type == "table":
            out.append(rendu_table(b.contenu))
    return "\n".join(out)

def rendu_sommaire(blocs) -> str:
    titres = [(slug_to_id(b.contenu), b.contenu) for b in blocs if b.type == "h2"]
    if not titres:
        return ""
    items = "\n".join('            <li><a href="#' + a + '">' + esc(t) + '</a></li>' for a, t in titres)
    return ('        <aside class="article-toc" aria-label="Sommaire">\n'
            '          <div class="article-toc-inner">\n'
            '            <p class="article-toc-title">Dans cet article</p>\n'
            '            <ol>\n' + items + '\n'
            '            </ol>\n'
            '          </div>\n'
            '        </aside>')

def rendu_sources(sources) -> str:
    if not sources:
        return ""
    items = "\n".join("          <li>" + linkify(esc(s)) + "</li>" for s in sources)
    return ('\n        <section id="sources" class="article-sources">\n'
            '          <h2>Sources et références</h2>\n'
            '          <ol>\n' + items + '\n'
            '          </ol>\n'
            '        </section>')

# ---------------------------------------------------------------------------
# Templates de navigation et footer alignés sur la convention du repo Git
# ---------------------------------------------------------------------------
NAV_INDEX = '''      <li><a href="../services/">Services</a></li>
      <li><a href="index.html" class="active">Blogue</a></li>
      <li><a href="../ressources/">Ressources</a></li>
      <li><a href="../formations/">Formations</a></li>
      <li><a href="../a-propos/">À propos</a></li>
      <li><a href="../contact/" class="nav-cta">Me contacter</a></li>'''

NAV_ARTICLE = '''      <li><a href="../services/">Services</a></li>
      <li><a href="index.html" class="active">Blogue</a></li>
      <li><a href="../ressources/">Ressources</a></li>
      <li><a href="../formations/">Formations</a></li>
      <li><a href="../a-propos/">À propos</a></li>
      <li><a href="../contact/" class="nav-cta">Me contacter</a></li>'''

FOOTER = '''  <footer>
    <div class="footer-grid">
      <div class="footer-brand">
        <div class="footer-logo">Benoit Plante</div>
        <p>Consultant en recherche à l'ère de l'IA. Veille, formations et accompagnement pour intégrer l'intelligence artificielle dans les pratiques de recherche, sans compromettre la rigueur.</p>
      </div>
      <div class="footer-col">
        <h4>Offre</h4>
        <ul>
          <li><a href="../services/">Services</a></li>
          <li><a href="../blogue/">Blogue</a></li>
          <li><a href="../ressources/">Ressources</a></li>
          <li><a href="../formations/">Formations</a></li>
          <li><a href="../outils/">Outils</a></li>
        </ul>
      </div>
      <div class="footer-col">
        <h4>À propos</h4>
        <ul>
          <li><a href="../a-propos/">Parcours</a></li>
          <li><a href="../contact/">Contact</a></li>
          <li><a href="../infolettre/">Infolettre</a></li>
        </ul>
      </div>
      <div class="footer-col">
        <h4>Légal &amp; liens</h4>
        <ul>
          <li><a href="../confidentialite/">Confidentialité</a></li>
          <li><a href="https://www.linkedin.com/in/benoitplante/" target="_blank" rel="noopener">LinkedIn ↗</a></li>
          <li><a href="https://orcid.org/0009-0003-5449-427X" target="_blank" rel="noopener">ORCID ↗</a></li>
        </ul>
      </div>
    </div>
    <div class="footer-bottom">
      <p>© 2026 Benoit Plante. Tous droits réservés.</p>
      <span class="footer-mark">IA &amp; recherche</span>
    </div>
  </footer>'''

NAV_OPEN = '''  <nav class="nav" aria-label="Navigation principale">
    <a class="nav-logo" href="../">Benoit Plante</a>
    <button class="nav-toggle" aria-label="Ouvrir le menu" aria-expanded="false">
      <svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round"><line x1="3" y1="6" x2="21" y2="6"/><line x1="3" y1="12" x2="21" y2="12"/><line x1="3" y1="18" x2="21" y2="18"/></svg>
    </button>
    <ul class="nav-links">
'''

NAV_CLOSE = '''
    </ul>
  </nav>'''

# ---------------------------------------------------------------------------
def page_html(article: Article) -> str:
    desc = description_meta(article)
    sommaire = rendu_sommaire(article.blocs)
    corps = rendu_blocs(article.blocs)
    sources = rendu_sources(article.sources)
    url = SITE_BASE_URL + "/blogue/" + article.slug + ".html"

    json_ld = {
        "@context": "https://schema.org",
        "@type": "BlogPosting",
        "headline": article.titre,
        "description": desc,
        "datePublished": article.date_iso,
        "author": {"@type": "Person", "name": "Benoit Plante", "url": SITE_BASE_URL + "/a-propos/"},
        "publisher": {"@type": "Person", "name": "Benoit Plante"},
        "mainEntityOfPage": url,
    }
    json_ld_str = json.dumps(json_ld, ensure_ascii=False, indent=2)

    return (
        '<!DOCTYPE html>\n<html lang="fr">\n<head>\n'
        '  <meta charset="UTF-8" />\n'
        '  <meta name="viewport" content="width=device-width, initial-scale=1.0" />\n'
        '  <title>' + esc(article.titre) + ' — Benoit Plante</title>\n'
        '  <meta name="description" content="' + esc(desc) + '" />\n\n'
        '  <meta property="og:type" content="article" />\n'
        '  <meta property="og:title" content="' + esc(article.titre) + '" />\n'
        '  <meta property="og:description" content="' + esc(desc) + '" />\n'
        '  <meta property="og:url" content="' + url + '" />\n'
        '  <meta property="og:locale" content="fr_CA" />\n'
        '  <meta property="article:published_time" content="' + article.date_iso + '" />\n'
        '  <meta property="article:author" content="Benoit Plante" />\n'
        '  <meta name="twitter:card" content="summary_large_image" />\n\n'
        '  <link rel="canonical" href="' + url + '" />\n'
        '  <link rel="icon" href="' + FAVICON + '" />\n\n'
        '  <link rel="preconnect" href="https://fonts.googleapis.com" />\n'
        '  <link rel="preconnect" href="https://fonts.gstatic.com" crossorigin />\n'
        '  <link href="https://fonts.googleapis.com/css2?family=Playfair+Display:wght@600;700&family=DM+Sans:ital,wght@0,300;0,400;0,500;0,600;1,300&display=swap" rel="stylesheet" />\n'
        '  <link rel="stylesheet" href="../styles.css" />\n'
        '  <link rel="stylesheet" href="article.css" />\n\n'
        '  <script type="application/ld+json">\n' + json_ld_str + '\n  </script>\n'
        '</head>\n<body>\n\n'
        + NAV_OPEN + NAV_ARTICLE + NAV_CLOSE + '\n\n'
        '  <section class="hero hero-compact article-hero">\n'
        '    <div class="hero-content" style="max-width: 820px;">\n'
        '      <div class="hero-eyebrow"><a href="index.html" style="color: inherit; text-decoration: none;">Blogue</a> · ' + esc(article.eyebrow) + '</div>\n'
        '      <h1 class="hero-h1">' + esc(article.titre) + '</h1>\n'
        '      <p class="hero-sub">' + esc(article.sous_titre) + '</p>\n'
        '      <p class="article-meta">\n'
        '        <span>Par Benoit Plante</span>\n'
        '        <span>·</span>\n'
        '        <time datetime="' + article.date_iso + '">' + article.date_fr + '</time>\n'
        '        <span>·</span>\n'
        '        <span>Lecture : ~' + str(article.temps_lecture) + ' min</span>\n'
        '      </p>\n'
        '    </div>\n'
        '  </section>\n\n'
        '  <section class="article-body">\n'
        '    <div class="inner article-layout">\n'
        + sommaire + '\n\n'
        '      <article class="article-content">\n'
        '        <p class="article-lead">' + esc(article.chapeau) + '</p>\n\n'
        + corps + '\n'
        + sources + '\n\n'
        '        <aside class="article-cta">\n'
        '          <p class="article-cta-eyebrow">Pour ne rien manquer</p>\n'
        '          <h3>Recevez la veille IA chaque semaine</h3>\n'
        "          <p>L'infolettre récapitule les nouveaux outils, méthodes et enjeux pour la recherche en sciences sociales. Une fois par semaine, sans spam.</p>\n"
        '          <a href="../infolettre/" class="btn-primary">M\'abonner à l\'infolettre →</a>\n'
        '        </aside>\n'
        '      </article>\n'
        '    </div>\n'
        '  </section>\n\n'
        + FOOTER + '\n\n'
        '  <script src="../script.js"></script>\n'
        '</body>\n</html>\n'
    )

# ---------------------------------------------------------------------------
def charger_manifest():
    if MANIFEST_PATH.exists():
        return json.loads(MANIFEST_PATH.read_text(encoding="utf-8"))
    return []

def sauver_manifest(entries):
    MANIFEST_PATH.write_text(json.dumps(entries, ensure_ascii=False, indent=2), encoding="utf-8")

def regenerer_index(entries) -> str:
    # Ne lister publiquement que les articles dont la date est passée
    aujourd_hui = dt.date.today().isoformat()
    visibles = [e for e in entries if e["date"] <= aujourd_hui]
    ordered = sorted(visibles, key=lambda e: e["date"], reverse=True)

    if ordered:
        cards = []
        for e in ordered:
            d = dt.date.fromisoformat(e["date"])
            date_fr = str(d.day) + " " + MOIS_FR[d.month] + " " + str(d.year)
            cards.append(
                '        <article class="blog-card">\n'
                '          <a href="' + esc(e["slug"]) + '.html" class="blog-card-link" aria-label="Lire ' + esc(e["titre"]) + '"></a>\n'
                '          <div class="blog-card-meta">\n'
                '            <span class="blog-card-eyebrow">' + esc(e.get("eyebrow","Article de fond")) + '</span>\n'
                '            <span class="blog-card-sep">·</span>\n'
                '            <time datetime="' + e["date"] + '">' + date_fr + '</time>\n'
                '            <span class="blog-card-sep">·</span>\n'
                '            <span>~' + str(e["temps_lecture"]) + ' min</span>\n'
                '          </div>\n'
                '          <h2 class="blog-card-title">' + esc(e["titre"]) + '</h2>\n'
                '          <p class="blog-card-dek">' + esc(e["sous_titre"]) + '</p>\n'
                "          <span class=\"blog-card-cta\">Lire l'article →</span>\n"
                '        </article>'
            )
        liste_html = "\n".join(cards)
    else:
        liste_html = (
            '        <div class="blog-empty">\n'
            "          <p>Le premier billet sera publié sous peu. Inscrivez-vous à l'infolettre pour être notifié.</p>\n"
            '          <a href="../infolettre/" class="btn-primary">M\'abonner →</a>\n'
            '        </div>'
        )

    return (
        '<!DOCTYPE html>\n<html lang="fr">\n<head>\n'
        '  <meta charset="UTF-8" />\n'
        '  <meta name="viewport" content="width=device-width, initial-scale=1.0" />\n'
        '  <title>Blogue — Benoit Plante</title>\n'
        "  <meta name=\"description\" content=\"Articles de fond sur la recherche à l'ère de l'IA : méthodes, outils, enjeux critiques. Un nouveau billet chaque lundi.\" />\n\n"
        '  <meta property="og:type" content="website" />\n'
        '  <meta property="og:title" content="Blogue — Benoit Plante" />\n'
        "  <meta property=\"og:description\" content=\"Articles de fond sur la recherche à l'ère de l'IA. Un nouveau billet chaque lundi.\" />\n"
        '  <meta property="og:locale" content="fr_CA" />\n\n'
        '  <link rel="canonical" href="' + SITE_BASE_URL + '/blogue/" />\n'
        '  <link rel="icon" href="' + FAVICON + '" />\n\n'
        '  <link rel="preconnect" href="https://fonts.googleapis.com" />\n'
        '  <link rel="preconnect" href="https://fonts.gstatic.com" crossorigin />\n'
        '  <link href="https://fonts.googleapis.com/css2?family=Playfair+Display:wght@600;700&family=DM+Sans:ital,wght@0,300;0,400;0,500;0,600;1,300&display=swap" rel="stylesheet" />\n'
        '  <link rel="stylesheet" href="../styles.css" />\n'
        '  <link rel="stylesheet" href="article.css" />\n'
        '</head>\n<body>\n\n'
        + NAV_OPEN + NAV_INDEX + NAV_CLOSE + '\n\n'
        '  <section class="hero hero-compact">\n'
        '    <div class="hero-content" style="max-width: 780px; text-align: center;">\n'
        '      <div class="hero-eyebrow" style="margin-left:auto; margin-right:auto;">Le blogue</div>\n'
        "      <h1 class=\"hero-h1\">Veille critique sur la <em>recherche</em> à l'ère de l'IA</h1>\n"
        '      <p class="hero-sub">Articles de fond sur les outils, méthodes et enjeux qui transforment la recherche en sciences sociales. Un nouveau billet chaque lundi.</p>\n'
        '    </div>\n'
        '  </section>\n\n'
        '  <section class="blog-list-section">\n'
        '    <div class="inner">\n'
        '      <div class="blog-list">\n'
        + liste_html + '\n'
        '      </div>\n'
        '    </div>\n'
        '  </section>\n\n'
        '  <section style="background: var(--gris-50); border-top: 1px solid var(--gris-200);">\n'
        '    <div class="inner" style="text-align: center; max-width: 720px;">\n'
        '      <span class="section-tag" style="margin-left:auto; margin-right:auto;">Pour ne rien manquer</span>\n'
        '      <h2 class="section-h2">Recevez chaque lundi le nouvel article + le digest hebdo</h2>\n'
        "      <p class=\"section-lead\" style=\"margin: 0 auto 2rem;\">L'infolettre relaie chaque semaine le nouvel article du blogue et résume les annonces marquantes en IA pour la recherche en sciences sociales.</p>\n"
        '      <a href="../infolettre/" class="btn-primary">M\'abonner gratuitement →</a>\n'
        '    </div>\n'
        '  </section>\n\n'
        + FOOTER + '\n\n'
        '  <script src="../script.js"></script>\n'
        '</body>\n</html>\n'
    )

# ---------------------------------------------------------------------------
def maj_sitemap(entries):
    if not SITEMAP_PATH.exists():
        return
    contenu = SITEMAP_PATH.read_text(encoding="utf-8")
    debut = "<!-- BLOGUE_START -->"
    fin   = "<!-- BLOGUE_END -->"

    bloc = [debut, "  <url>",
            "    <loc>" + SITE_BASE_URL + "/blogue/</loc>",
            "    <changefreq>weekly</changefreq>",
            "    <priority>0.8</priority>",
            "  </url>"]
    aujourd_hui = dt.date.today().isoformat()
    visibles_sitemap = [e for e in entries if e["date"] <= aujourd_hui]
    for e in sorted(visibles_sitemap, key=lambda x: x["date"], reverse=True):
        bloc.extend(["  <url>",
                     "    <loc>" + SITE_BASE_URL + "/blogue/" + e["slug"] + ".html</loc>",
                     "    <lastmod>" + e["date"] + "</lastmod>",
                     "    <changefreq>monthly</changefreq>",
                     "    <priority>0.7</priority>",
                     "  </url>"])
    bloc.append(fin)
    nouveau = "\n".join(bloc)

    if debut in contenu and fin in contenu:
        contenu = re.sub(re.escape(debut) + r".*?" + re.escape(fin),
                         nouveau, contenu, flags=re.DOTALL)
    else:
        contenu = contenu.replace("</urlset>", nouveau + "\n</urlset>")
    SITEMAP_PATH.write_text(contenu, encoding="utf-8")

# ---------------------------------------------------------------------------
def lister_file_attente():
    out = []
    if not QUEUE_DIR.exists():
        return out
    for f in sorted(QUEUE_DIR.iterdir()):
        if not f.is_file() or not f.name.lower().endswith(".docx"):
            continue
        if f.name.startswith("~$"):
            continue
        m = QUEUE_FILE_RE.match(f.name)
        if not m:
            sys.stderr.write("WARN : " + f.name + " ignoré (format invalide)\n")
            continue
        annee, mois, jour, slug = m.groups()
        try:
            d = dt.date(int(annee), int(mois), int(jour))
        except ValueError:
            sys.stderr.write("WARN : date invalide dans " + f.name + "\n")
            continue
        out.append((d, slug, f))
    return out

def choisir_prochain(aujourd_hui):
    candidats = [(d, s, p) for d, s, p in lister_file_attente() if d <= aujourd_hui]
    if not candidats:
        return None
    candidats.sort(key=lambda x: x[0])
    return candidats[0]

# ---------------------------------------------------------------------------
def publier(docx_path, slug, date_publi, dry_run=False, archive=True):
    article = lire_docx(docx_path, slug, date_publi)
    page = page_html(article)
    cible = BLOG_DIR / (slug + ".html")

    print("  -> Article  : " + article.titre)
    print("  -> Slug     : " + slug)
    print("  -> Date     : " + article.date_fr)
    print("  -> Mots     : " + str(article.nb_mots) + " (~" + str(article.temps_lecture) + " min)")
    print("  -> Sections : " + str(sum(1 for b in article.blocs if b.type == "h2")))
    print("  -> Sources  : " + str(len(article.sources)))

    entree = {
        "slug": slug, "titre": article.titre, "sous_titre": article.sous_titre,
        "eyebrow": article.eyebrow, "date": article.date_iso,
        "temps_lecture": article.temps_lecture,
    }
    if dry_run:
        print("  -> [DRY-RUN] Cible : " + str(cible))
        return entree

    BLOG_DIR.mkdir(parents=True, exist_ok=True)
    cible.write_text(page, encoding="utf-8")
    print("  OK Page : " + str(cible.relative_to(SITE_ROOT)))

    entries = charger_manifest()
    entries = [e for e in entries if e["slug"] != slug]
    entries.append(entree)
    sauver_manifest(entries)
    print("  OK Manifest (" + str(len(entries)) + " billets)")

    (BLOG_DIR / "index.html").write_text(regenerer_index(entries), encoding="utf-8")
    print("  OK Index regenere")

    maj_sitemap(entries)
    print("  OK sitemap.xml")

    if archive:
        ARCHIVE_DIR.mkdir(parents=True, exist_ok=True)
        cible_arch = ARCHIVE_DIR / docx_path.name
        shutil.move(str(docx_path), str(cible_arch))
        print("  OK Source archivee -> " + str(cible_arch.relative_to(SITE_ROOT)))
    return entree

# ---------------------------------------------------------------------------
def main():
    parser = argparse.ArgumentParser(description="Pipeline de publication blogue benoitplante.com")
    parser.add_argument("--file", type=Path)
    parser.add_argument("--dry-run", action="store_true")
    parser.add_argument("--no-archive", action="store_true")
    parser.add_argument("--date", type=str)
    parser.add_argument("--rebuild-index", action="store_true")
    args = parser.parse_args()

    print("== Pipeline de publication blogue — " + dt.datetime.now().strftime("%Y-%m-%d %H:%M") + " ==")
    print("Racine : " + str(SITE_ROOT))

    if args.rebuild_index:
        entries = charger_manifest()
        BLOG_DIR.mkdir(parents=True, exist_ok=True)
        (BLOG_DIR / "index.html").write_text(regenerer_index(entries), encoding="utf-8")
        maj_sitemap(entries)
        print("OK Index et sitemap regeneres (" + str(len(entries)) + " billets)")
        return 0

    aujourd_hui = dt.date.today()

    if args.file:
        path = args.file
        if not path.exists():
            sys.stderr.write("ERREUR : " + str(path) + " introuvable\n")
            return 1
        m = QUEUE_FILE_RE.match(path.name)
        if m:
            annee, mois, jour, slug = m.groups()
            date_publi = dt.date(int(annee), int(mois), int(jour))
        else:
            slug = path.stem
            date_publi = aujourd_hui
        if args.date:
            date_publi = dt.date.fromisoformat(args.date)
        print("\nPublication forcee : " + path.name)
        publier(path, slug, date_publi, dry_run=args.dry_run, archive=not args.no_archive)
        return 0

    print("Date : " + str(aujourd_hui))
    file_attente = lister_file_attente()
    print("\nFile d'attente : " + str(len(file_attente)) + " article(s)")
    for d, s, _ in file_attente:
        marqueur = "  <- prochain" if d <= aujourd_hui else ""
        print("  - " + str(d) + " - " + s + marqueur)

    prochain = choisir_prochain(aujourd_hui)
    if not prochain:
        print("\nAucun article du aujourd'hui. Rien a publier.")
        return 0

    d, slug, path = prochain
    print("\nPublication : " + path.name)
    publier(path, slug, d, dry_run=args.dry_run, archive=not args.no_archive)
    print("\nPublication terminee.")
    return 0

if __name__ == "__main__":
    sys.exit(main())
