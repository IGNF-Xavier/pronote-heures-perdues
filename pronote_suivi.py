#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""
pronote_suivi - parcourt l'emploi du temps hebdomadaire Pronote (compte parent),
fait une capture d'écran par semaine et produit un Excel d'analyse
(annulations, grèves, absences de professeurs, déplacements, modifications...).

Outil autonome : Python + Playwright + openpyxl.
Aucun identifiant n'est lu ni stocké : vous vous connectez vous-même dans la
fenêtre du navigateur.

Deux étapes, séparables :
  capture  : pilote le navigateur, enregistre semaines/AAAA-MM-JJ.png + .json
  analyse  : relit les .json et (re)génère le classeur Excel (sans navigateur)
Par défaut : capture puis analyse. Sans option, l'outil pose les questions
(dates de début et de fin) en proposant des valeurs par défaut.
"""
import argparse
import datetime as dt
import json
import re
import statistics
import sys
import unicodedata
from pathlib import Path

__version__ = "0.9"

JOURS = ["lundi", "mardi", "mercredi", "jeudi", "vendredi", "samedi", "dimanche"]

# ----------------------------------------------------------------------------
# Classification : (catégorie, motifs sur texte normalisé sans accents).
# L'ordre = priorité. Modifiable ici, puis relancer avec --etape analyse.
# ----------------------------------------------------------------------------
REGLES = [
    ("Grève",            [r"greve"]),
    ("Prof absent",      [r"(prof\w*|enseignant|intervenant)\.?\s*(/\w+\.?)?\s*absent",
                          r"absence (du |de l.)?(prof|enseignant)", r"absent\b.*prof"]),
    ("Annulé",           [r"annul"]),
    ("Remplacé",         [r"remplac"]),
    # simple changement de salle : le cours a bien lieu -> Maintenu
    ("Maintenu",         [r"changement de salle", r"salle changee", r"salle modifiee",
                          r"cours maintenu"]),
    # nouveau créneau d'un cours perdu : Pronote écrit "Cours déplacé" ou "Cours modifié"
    ("Déplacé/Modifié",  [r"deplac", r"modifi", r"changement d.horaire", r"exceptionnel"]),
    ("Sortie/Événement", [r"sortie", r"voyage", r"evenement", r"\bstage\b"]),
    ("Dispensé",         [r"dispens"]),
    ("Absence élève",    [r"\beleve absent", r"\babsence\b"]),
]
CATEGORIE_DEFAUT = "Maintenu"
RATTRAPAGE = "Déplacé/Modifié"
# heures perdues (le cours n'a pas eu lieu à ce créneau)
PERDUES = ["Prof absent", "Annulé", "Grève", "Dispensé", "Absence élève"]
EXCLUES = ["vie de classe"]   # matières affichées mais hors calculs globaux (noms normalisés)
# onglets FCPE : valeurs reprises de parametres.ini (motifs = noms de catégories normalisés)
FCPE = {"ville": "", "etablissement": "", "classe": "", "motifs": ["prof absent"],
        "regroupement": "matiere"}     # une déclaration par "matiere" ou par "periode" d'absence
FENETRE_JOURS = 14  # un rattrapage est apparié à une heure perdue de la même matière à +/- N jours

JS_EXTRACTION = r"""
() => {
  const vis = e => { const r = e.getBoundingClientRect(); return r.width > 2 && r.height > 2; };
  const rect = e => { const r = e.getBoundingClientRect();
    return {x: r.left + scrollX, y: r.top + scrollY, w: r.width, h: r.height}; };
  const txt = e => (e.innerText || '').trim();
  const all = Array.from(document.querySelectorAll('body *')).filter(vis);

  // en-têtes de jours
  const reJour = /^(lun|mar|mer|jeu|ven|sam|dim)[a-zé]*\.?\s*\d{0,2}/i;
  const headers = all.filter(e => { const t = txt(e);
      return t.length < 40 && reJour.test(t) && !t.includes('\n') &&
             Array.from(e.children).every(c => !reJour.test(txt(c))); })
    .map(e => ({text: txt(e), ...rect(e)}));

  // graduations horaires
  const reHeure = /^\d{1,2}\s*[h:]\s*(\d{2})?$/i;
  const hours = all.filter(e => { const t = txt(e);
      return reHeure.test(t) && e.children.length === 0; })
    .map(e => ({text: txt(e), ...rect(e)}));

  // cours : sélecteur personnalisé sinon heuristique sur la classe
  const sel = window.__SEL_COURS__;
  let cand = sel ? Array.from(document.querySelectorAll(sel)).filter(vis)
                 : all.filter(e => /cours/i.test(e.className && e.className.baseVal === undefined
                        ? e.className : '') );
  const W = innerWidth, H = innerHeight;
  cand = cand.filter(e => { const r = e.getBoundingClientRect();
      return txt(e).length > 2 && r.width < W * 0.45 && r.height < H * 0.8; });
  const set = new Set(cand);
  cand = cand.filter(e => { let p = e.parentElement;
      while (p) { if (set.has(p)) return false; p = p.parentElement; } return true; });

  const courses = cand.map(e => ({
    text: txt(e), title: e.getAttribute('title') || e.getAttribute('aria-label') || '',
    cls: String(e.className && e.className.baseVal === undefined ? e.className : ''),
    ...rect(e)}));
  return {headers, hours, courses, url: location.href, title: document.title};
}
"""


# ----------------------------------------------------------------------------
# Utilitaires
# ----------------------------------------------------------------------------
def norm(s):
    s = unicodedata.normalize("NFKD", s or "")
    return "".join(c for c in s if not unicodedata.combining(c)).lower()


def lundi_de(d):
    return d - dt.timedelta(days=d.weekday())


def parse_date(s):
    """Accepte JJ/MM/AAAA (usage courant) ou AAAA-MM-JJ."""
    s = s.strip()
    for fmt in ("%d/%m/%Y", "%Y-%m-%d"):
        try:
            d = dt.datetime.strptime(s, fmt).date()
        except ValueError:
            continue
        if d.year >= 2000:
            return d
    raise ValueError(f"date illisible : {s!r}")


def classer(texte, titre="", classes=""):
    """Retourne (catégorie, indice de détection)."""
    base = norm(texte + " " + titre)
    cl = norm(classes)
    for cat, motifs in REGLES:
        for m in motifs:
            hit = re.search(m, base)
            if hit:
                return cat, f"texte: {hit.group(0)}"
    for cat, motifs in REGLES:      # indices CSS, seulement les plus sûrs
        if cat not in ("Grève", "Prof absent", "Annulé", RATTRAPAGE):
            continue
        for m in motifs[:1]:
            if re.search(m, cl):
                return cat, f"css: {m}"
    return CATEGORIE_DEFAUT, ""


def moindres_carres(pts):
    xs, ys = zip(*pts)
    mx, my = statistics.fmean(xs), statistics.fmean(ys)
    den = sum((x - mx) ** 2 for x in xs)
    if den == 0:
        return None
    a = sum((x - mx) * (y - my) for x, y in pts) / den
    return a, my - a * mx


def vers_minutes(txt):
    m = re.match(r"(\d{1,2})\s*[h:]\s*(\d{2})?", txt.strip(), re.I)
    return int(m.group(1)) * 60 + int(m.group(2) or 0) if m else None


MOIS = ["janvier", "fevrier", "mars", "avril", "mai", "juin", "juillet", "aout",
        "septembre", "octobre", "novembre", "decembre"]
RE_TITRE = re.compile(r"du (\d{1,2})\w* ([a-z]+) de (\d{1,2}) heures?\s*(\d{2})?\s*a\s*"
                      r"(\d{1,2}) heures?\s*(\d{2})?")


def lire_titre(titre, lundi):
    """'Cours du 1 septembre de 10 heures 30 à 11 heures 30' -> (date, debut_min, fin_min)."""
    m = RE_TITRE.search(norm(titre))
    if not m or m.group(2) not in MOIS:
        return None
    mois = MOIS.index(m.group(2)) + 1
    for annee in (lundi.year - 1, lundi.year, lundi.year + 1):
        try:
            d = dt.date(annee, mois, int(m.group(1)))
        except ValueError:
            continue
        if -3 <= (d - lundi).days <= 9:
            return (d, int(m.group(3)) * 60 + int(m.group(4) or 0),
                    int(m.group(5)) * 60 + int(m.group(6) or 0))
    return None


def traiter_semaine(brut, lundi):
    """Transforme l'extraction brute d'une page en lignes de cours."""
    # axe vertical : y (px) -> minutes
    pts = [(h["y"] + h["h"] / 2, vers_minutes(h["text"])) for h in brut["hours"]]
    pts = [p for p in pts if p[1] is not None]
    axe = moindres_carres(pts) if len(pts) >= 2 else None
    # axe horizontal : jours
    entetes = []
    for h in brut["headers"]:
        mj = re.match(r"(lun|mar|mer|jeu|ven|sam|dim)", norm(h["text"]))
        if mj:
            idx = ["lun", "mar", "mer", "jeu", "ven", "sam", "dim"].index(mj.group(1))
            entetes.append((h["x"] + h["w"] / 2, idx))
    lignes, vus = [], set()
    for c in brut["courses"]:
        cx = c["x"] + c["w"] / 2
        jour = min(entetes, key=lambda e: abs(e[0] - cx))[1] if entetes else None
        debut = fin = None
        lt = lire_titre(c["title"], lundi)
        m = None if lt else re.search(r"(\d{1,2})\s*[h:]\s*(\d{2})\D{1,15}?(\d{1,2})\s*[h:]\s*(\d{2})", c["text"])
        if lt:
            jour, debut, fin = (lt[0] - lundi).days, lt[1], lt[2]
        elif m:
            debut = int(m[1]) * 60 + int(m[2])
            fin = int(m[3]) * 60 + int(m[4])
        elif axe:
            a, b = axe
            snap = lambda v: int(round(v / 5.0)) * 5
            debut, fin = snap(a * c["y"] + b), snap(a * (c["y"] + c["h"]) + b)
        cle = (jour, debut, fin, c["text"])
        if cle in vus:
            continue
        vus.add(cle)
        parts = [p.strip() for p in c["text"].split("\n") if p.strip()
                 and not norm(p).startswith("ouverture des details")]
        cat, indice = classer("\n".join(parts), c["title"], c["cls"])
        est_statut = lambda p: any(re.search(m, norm(p)) for _, ms in REGLES for m in ms)
        matiere = next((p for p in parts if not est_statut(p)), parts[0] if parts else "")
        hhmm = lambda v: "" if v is None else f"{v // 60:02d}:{v % 60:02d}"
        lignes.append({
            "semaine": lundi.isoformat(),
            "date": (lundi + dt.timedelta(days=jour)).isoformat() if jour is not None else "",
            "jour": JOURS[jour].capitalize() if jour is not None and 0 <= jour < 7 else "?",
            "debut": hhmm(debut), "fin": hhmm(fin),
            "duree_h": round((fin - debut) / 60, 2) if debut is not None and fin is not None
                       and fin > debut else None,
            "matiere": matiere,
            "categorie": cat, "indice": indice,
            "texte": " | ".join(parts),
            "boite": (c["x"], c["y"], c["w"], c["h"]),
        })
    lignes.sort(key=lambda l: (l["date"], l["debut"]))
    return lignes


# ----------------------------------------------------------------------------
# Étape 1 : capture
# ----------------------------------------------------------------------------
RE_PLAGE = re.compile(r"du (\d{2})/(\d{2})/(\d{4}) au (\d{2})/(\d{2})/(\d{4})")

JS_BANDE = r"""
() => Array.from(document.querySelectorAll('body *')).filter(e => {
    const r = e.getBoundingClientRect(), t = (e.innerText || '').trim();
    return /^\d{1,2}$/.test(t) && e.children.length === 0 && r.width > 4 && r.top < 260 && r.top > 100;
  }).map(e => { const r = e.getBoundingClientRect();
    return {n: parseInt(e.innerText.trim()), x: r.left + r.width / 2, y: r.top + r.height / 2}; })
"""


def sur(page, f, essais=8):
    """Exécute f(page) en réessayant si Pronote navigue/recharge pendant la lecture."""
    for k in range(essais):
        try:
            return f(page)
        except Exception as e:
            if k == essais - 1 or not any(t in str(e) for t in
                                          ("context was destroyed", "navigation", "Target closed")):
                raise
            try:
                page.wait_for_load_state("domcontentloaded", timeout=5000)
            except Exception:
                pass
            page.wait_for_timeout(1000)


def plage_affichee(page):
    """Premier jour de la semaine affichée, lu dans l'en-tête 'du JJ/MM/AAAA au JJ/MM/AAAA'."""
    m = RE_PLAGE.search(sur(page, lambda p: p.inner_text("body")))
    if not m:
        return None
    return dt.date(int(m[3]), int(m[2]), int(m[1]))


def cliquer_suivant(page, selecteur):
    cands = []
    if selecteur:
        cands.append(page.locator(selecteur))
    cands += [
        page.get_by_role("button", name=re.compile(r"semaine suivante", re.I)),
        page.locator("[title*='emaine suivante' i], [aria-label*='emaine suivante' i]"),
    ]
    for loc in cands:
        try:
            if loc.count() > 0 and loc.first.is_visible():
                loc.first.click(timeout=3000)
                return True
        except Exception:
            continue
    return False


def aller_a_semaine(page, x, y, avant, attente):
    """Clique sur un numéro de semaine et attend que la semaine affichée change."""
    page.mouse.click(x, y)
    page.mouse.move(700, 700)
    for _ in range(12):
        page.wait_for_timeout(max(attente // 3, 300))
        lundi = plage_affichee(page)
        if lundi and lundi != avant:
            page.wait_for_timeout(attente)
            return lundi
    return plage_affichee(page)


def enregistrer(page, sem_dir, lundi, selecteur_cours):
    png = sem_dir / f"{lundi.isoformat()}.png"
    page.screenshot(path=str(png), full_page=True)
    sur(page, lambda p: p.evaluate(f"window.__SEL_COURS__ = {json.dumps(selecteur_cours)}"))
    brut = sur(page, lambda p: p.evaluate(JS_EXTRACTION))
    brut["lundi"] = lundi.isoformat()
    (sem_dir / f"{lundi.isoformat()}.json").write_text(
        json.dumps(brut, ensure_ascii=False, indent=1), encoding="utf-8")
    return len(brut["courses"])


def capture(args, dossier):
    from playwright.sync_api import sync_playwright
    rentree = lundi_de(parse_date(args.rentree))
    cible = lundi_de(parse_date(args.semaine))
    if cible < rentree:
        sys.exit("La semaine choisie est avant la rentrée.")
    sem_dir = dossier / "semaines"
    sem_dir.mkdir(parents=True, exist_ok=True)

    with sync_playwright() as pw:
        ctx = None
        for canal in ([args.navigateur] if args.navigateur else ["msedge", "chrome", None]):
            try:
                kw = {"channel": canal} if canal else {}
                ctx = pw.chromium.launch_persistent_context(
                    str(dossier / "profil_navigateur"), headless=False,
                    viewport={"width": 1500, "height": 1000}, locale="fr-FR", **kw)
                break
            except Exception as e:
                err = e
        if ctx is None:
            sys.exit(f"Impossible de lancer un navigateur : {err}\n"
                     "Installez Edge/Chrome ou lancez : python -m playwright install chromium")
        page = ctx.pages[0] if ctx.pages else ctx.new_page()
        if args.url:
            page.goto(args.url)
        print("\n=== Dans la fenêtre du navigateur ===")
        print(" 1. Connectez-vous à Pronote (compte parent) et choisissez l'enfant.")
        print(" 2. Ouvrez 'Emploi du temps' (vue grille hebdomadaire).")
        while True:
            input("Quand la grille est affichée, appuyez sur Entrée ici... ")
            page = None
            for onglet in reversed(ctx.pages):      # l'ENT ouvre parfois Pronote dans un nouvel onglet
                try:
                    if plage_affichee(onglet):
                        page = onglet
                        break
                except Exception:
                    continue
            if page:
                break
            print("\n/!\\ Je ne vois pas l'emploi du temps (je cherche « du JJ/MM/AAAA au JJ/MM/AAAA » en haut "
                  "de la page).\n    Ouvrez Vie scolaire > Emploi du temps, attendez l'affichage, puis réessayez "
                  "(Ctrl+C pour abandonner).")
        page.bring_to_front()

        page.wait_for_timeout(1500)
        sur(page, lambda p: p.evaluate("window.scrollTo(0, 0)"))
        bande = sorted({b["n"]: b for b in sur(page, lambda p: p.evaluate(JS_BANDE))}.values(),
                       key=lambda b: b["n"])
        faits, vus = 0, set()
        if len(bande) >= 5 and not args.manuel:
            print(f"Bande des semaines détectée ({len(bande)} semaines d'école).")
            courante = plage_affichee(page)
            for b in bande:
                sur(page, lambda p: p.evaluate("window.scrollTo(0, 0)"))
                lundi = aller_a_semaine(page, b["x"], b["y"], courante, args.attente)
                if lundi is None:
                    print(f"  semaine {b['n']} : dates illisibles, ignorée")
                    continue
                courante = lundi
                lundi = lundi_de(lundi)
                if lundi in vus:
                    continue
                vus.add(lundi)
                if lundi < rentree:
                    continue
                if lundi > cible:
                    break
                n = enregistrer(page, sem_dir, lundi, args.selecteur_cours)
                faits += 1
                print(f"[sem. {b['n']:>2}] {lundi:%d/%m/%Y} -> {n} cours détectés")
        else:
            print("Navigation semaine par semaine (bouton suivant ou manuelle).")
            print(f"Placez-vous sur la semaine de rentrée (du {rentree:%d/%m/%Y}).")
            if args.manuel:
                input("Puis Entrée... ")
            manuel = args.manuel
            while True:
                page.wait_for_timeout(args.attente)
                lundi = plage_affichee(page)
                lundi = lundi_de(lundi) if lundi else None
                if lundi is None or lundi in vus or lundi > cible:
                    break
                vus.add(lundi)
                if lundi >= rentree:
                    n = enregistrer(page, sem_dir, lundi, args.selecteur_cours)
                    faits += 1
                    print(f"{lundi:%d/%m/%Y} -> {n} cours détectés")
                if lundi >= cible:
                    break
                if not manuel and not cliquer_suivant(page, args.selecteur_suivant):
                    print("   Bouton 'semaine suivante' introuvable -> mode manuel.")
                    manuel = True
                if manuel:
                    input("   Passez à la semaine suivante dans le navigateur, puis Entrée... ")
        ctx.close()
    print(f"{faits} semaine(s) capturée(s).")
    return faits


# ----------------------------------------------------------------------------
# Étape 2 : analyse -> Excel
# ----------------------------------------------------------------------------
# ----------------------------------------------------------------------------
# Onglets FCPE (outil en ligne Ouyapacours)
# ----------------------------------------------------------------------------
URL_FCPE = "https://ouyapacours.fcpe.asso.fr/"


def declarations_fcpe(lignes, jusqu_au):
    """Cours non assurés regroupés en déclarations : une par matière (défaut), ou une par période
    d'absence (close dès que la matière a de nouveau cours)."""
    par_mat = {}
    for l in sorted(lignes, key=lambda l: (l["date"], l["debut"])):
        if (not l["date"] or l["date"] > jusqu_au.isoformat()
                or norm(l["matiere"]).strip() in EXCLUES):
            continue
        par_mat.setdefault(l["matiere"], []).append(l)
    par_periode = FCPE["regroupement"] == "periode"
    decl = []
    for matiere, cours in sorted(par_mat.items()):
        courante = None
        for l in cours:
            if norm(l["categorie"]) in FCPE["motifs"] and l["compense"] != "Oui":
                if courante is None:
                    courante = {"matiere": matiere, "cours": [], "reprise": ""}
                    decl.append(courante)
                courante["cours"].append(l)
                courante["reprise"] = ""
            elif l["categorie"] not in PERDUES and courante is not None:
                courante["reprise"] = courante["reprise"] or l["date"]   # un cours a eu lieu
                if par_periode:
                    courante = None
    if par_periode:
        decl.sort(key=lambda d: (d["cours"][0]["date"], d["cours"][0]["debut"], d["matiere"]))
    return decl


def heures_formulaire(cours):
    """Durée à choisir dans la liste du formulaire (heures entières : 55 min = 1 h)."""
    return max(1, int(sum(l["duree_h"] or 0 for l in cours) + 0.5))


def captures_fcpe(decl, sem_dir, cible):
    """Copies des captures pour justificatifs : bandeau du haut (noms, photo) retiré et
    cours déclarés entourés en rouge. Les captures d'origine ne sont pas modifiées."""
    from PIL import Image, ImageDraw
    par_semaine = {}
    for d in decl:
        for l in d["cours"]:
            par_semaine.setdefault(l["semaine"], []).append(l)
    faites = {}
    for semaine, cours in sorted(par_semaine.items()):
        source = sem_dir / f"{semaine}.png"
        donnees = sem_dir / f"{semaine}.json"
        if not source.exists() or not donnees.exists():
            continue
        brut = json.loads(donnees.read_text(encoding="utf-8"))
        jours = [h["y"] for h in brut["headers"]
                 if re.match(r"(lun|mar|mer|jeu|ven|sam|dim)[a-z]*\.? +\d", norm(h["text"]))]
        im = Image.open(source).convert("RGB")
        trait = ImageDraw.Draw(im)
        for l in cours:
            x, y, w, h = l["boite"]
            trait.rectangle([x - 2, y - 2, x + w + 2, y + h + 2], outline=(220, 0, 0), width=5)
        # on garde la ligne des dates (environ 85 px au-dessus des jours) jusqu'au bas de la grille
        haut = int(min(jours)) - 85 if jours else 125
        bas = int(max(c["y"] + c["h"] for c in brut["courses"])) + 12
        cible.mkdir(parents=True, exist_ok=True)
        dest = cible / f"semaine_{semaine}.png"
        im.crop((0, max(haut, 0), im.width, min(bas, im.height))).save(dest)
        faites[semaine] = dest
    return faites


def onglets_fcpe(wb, lignes, sem_dir, sortie, jusqu_au):
    import os
    from openpyxl.styles import Alignment, Font, PatternFill
    from openpyxl.utils import get_column_letter as L

    decl = declarations_fcpe(lignes, jusqu_au)
    cible = sortie.parent / "captures_fcpe" / sortie.stem
    images = captures_fcpe(decl, sem_dir, cible)

    blanc, bleu = Font(bold=True, color="FFFFFF"), PatternFill("solid", fgColor="1F4E78")
    jaune, total = PatternFill("solid", fgColor="FFF2CC"), PatternFill("solid", fgColor="DDEBF7")
    lien_police = Font(color="0563C1", underline="single")
    jour = lambda s: parse_date(s) if s else None

    def entete(ws, r, titres):
        for j, t in enumerate(titres, 1):
            c = ws.cell(r, j, t)
            c.font, c.fill = blanc, bleu
            c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        ws.row_dimensions[r].height = 32

    def lien(cellule, chemin, texte):
        cellule.value = texte
        cellule.hyperlink = os.path.relpath(chemin, sortie.parent).replace("\\", "/")
        cellule.font = lien_police

    ws = wb.create_sheet("Saisie FCPE", 1)
    ws["A1"] = "À saisir dans Ouyapacours (FCPE) : heures de cours perdues"
    ws["A1"].font = Font(bold=True, size=14)
    ws["A2"] = URL_FCPE
    ws["A2"].hyperlink, ws["A2"].font = URL_FCPE, lien_police
    noms_motifs = [c for c, _ in REGLES if norm(c) in FCPE["motifs"]]
    ws["A3"] = ("Pour chaque déclaration : choisir l'enfant, vérifier ville, établissement et classe, saisir la "
                "matière, puis la 1re ligne Date + Durée ; « Ajouter des heures » pour chaque ligne suivante ; "
                "« Valider ». La durée est en heures entières (un cours de 55 min = 1 h) : le total diffère donc "
                "de la Synthèse. Motifs retenus : " + ", ".join(dict.fromkeys(noms_motifs)) +
                ". Non comptés : cours rattrapés, matières exclues" +
                (f" ({', '.join(EXCLUES)})" if EXCLUES else "") +
                f", cours après le {jusqu_au:%d/%m/%Y}. Onglet figé à la génération.")
    ws.merge_cells("A3:J3")
    ws["A3"].alignment = Alignment(wrap_text=True, vertical="top")
    ws.row_dimensions[3].height = 48
    champs = [("Ville / code postal", FCPE["ville"]), ("Etablissement", FCPE["etablissement"]),
              ("Classe", FCPE["classe"])]
    for r, (libelle, valeur) in enumerate(champs, 5):
        ws.cell(r, 1, libelle).font = Font(bold=True)
        ws.cell(r, 2, valeur).fill = jaune
        ws.merge_cells(start_row=r, start_column=2, end_row=r, end_column=4)

    ws["A9"] = "1. Les lignes à saisir"
    ws["A9"].font = Font(bold=True, size=12)
    entete(ws, 10, ["Déclaration n°", "Matiere", "Date", "Duree", "Dans le formulaire", "Cours concernés",
                    "Motif", "Capture (justificatif)", "Après validation", "Saisi le"])
    from openpyxl.styles import Border, Side
    separation = Border(top=Side(style="medium", color="1F4E78"))
    gris = PatternFill("solid", fgColor="F2F2F2")
    recap, r = [], 11
    for n, d in enumerate(decl, 1):
        par_date = {}
        for l in d["cours"]:
            par_date.setdefault(l["date"], []).append(l)
        etat = (f"Cours repris le {jour(d['reprise']):%d/%m} : signaler la fin" if d["reprise"]
                else "Pas de reprise : revenir signaler la prolongation ou la fin")
        total_decl = 0
        for k, (date, cours) in enumerate(sorted(par_date.items())):
            h = heures_formulaire(cours)
            total_decl += h
            vals = [n, d["matiere"], jour(date), h,
                    "Nouvelle déclaration" if k == 0 else "+ Ajouter des heures",
                    ", ".join(f"{l['debut']}-{l['fin']}" for l in cours),
                    ", ".join(dict.fromkeys(l["categorie"] for l in cours)), None,
                    etat if k == 0 else None, None]
            for j, v in enumerate(vals, 1):
                c = ws.cell(r, j, v)
                if n % 2 == 0:
                    c.fill = gris
                if k == 0:
                    c.border = separation
            ws.cell(r, 3).number_format = ws.cell(r, 10).number_format = "dd/mm/yyyy"
            ws.cell(r, 4).number_format = '0"h"'
            ws.cell(r, 10).fill = jaune
            semaine = cours[0]["semaine"]
            if semaine in images:
                lien(ws.cell(r, 8), images[semaine], images[semaine].name)
            r += 1
        recap.append((n, d["matiere"], len(par_date), total_decl, min(par_date), max(par_date), etat))
    t = r
    ws.cell(t, 1, "TOTAL")
    ws.cell(t, 4, f"=SUM(D11:D{t - 1})" if decl else 0).number_format = '0"h"'
    for j in range(1, 11):
        ws.cell(t, j).font, ws.cell(t, j).fill = Font(bold=True), total

    r0 = t + 3
    ws.cell(r0 - 1, 1, "2. Récapitulatif des déclarations").font = Font(bold=True, size=12)
    entete(ws, r0, ["Déclaration n°", "Matiere", "Nb de lignes", "Heures", "Première date", "Dernière date",
                    "État"])
    for r, (n, m, nb_l, h, d1, d2, etat) in enumerate(recap, r0 + 1):
        for j, v in enumerate([n, m, nb_l, h, jour(d1), jour(d2), etat], 1):
            ws.cell(r, j, v)
        ws.cell(r, 4).number_format = '0"h"'
        ws.cell(r, 5).number_format = ws.cell(r, 6).number_format = "dd/mm/yyyy"
    for j, w in enumerate([19, 28, 14, 14, 23, 28, 14, 30, 58, 13], 1):
        ws.column_dimensions[L(j)].width = w
    ws.freeze_panes = "C11"

    wj = wb.create_sheet("Justificatifs FCPE", 2)
    entete(wj, 1, ["N° déclaration", "Date", "Jour", "Début", "Fin", "Discipline", "Durée (h)", "Motif",
                   "Capture (cours entouré en rouge)"])
    r = 2
    for n, d in enumerate(decl, 1):
        for l in d["cours"]:
            vals = [n, jour(l["date"]), l["jour"], l["debut"], l["fin"], l["matiere"], l["duree_h"],
                    l["categorie"]]
            for j, v in enumerate(vals, 1):
                wj.cell(r, j, v)
            wj.cell(r, 2).number_format = "dd/mm/yyyy"
            if l["semaine"] in images:
                lien(wj.cell(r, 9), images[l["semaine"]], images[l["semaine"]].name)
            r += 1
    for j, w in enumerate([14, 12, 11, 8, 8, 30, 10, 16, 34], 1):
        wj.column_dimensions[L(j)].width = w
    wj.freeze_panes = "A2"
    wj.auto_filter.ref = f"A1:I{max(r - 1, 2)}"

    for f in (ws, wj):
        f.page_setup.orientation = "landscape"
        f.sheet_properties.pageSetUpPr.fitToPage = True
        f.page_setup.fitToHeight = 0
    nb = sum(len(d["cours"]) for d in decl)
    heures = sum(x[3] for x in recap)
    return (f"FCPE : {len(decl)} déclaration(s), {nb} cours non assurés, {heures} h à saisir (onglet « Saisie FCPE »)"
            + (f"\n       captures prêtes : {cible}" if images else ""))


def apparier(lignes):
    """Associe chaque heure perdue à son rattrapage (même matière, dates proches)."""
    for l in lignes:
        l["compense"], l["lie"] = "", ""
    perdues = [l for l in lignes if l["categorie"] in PERDUES]
    for l in perdues:
        l["compense"] = "Non"
    gains = sorted((l for l in lignes if l["categorie"] == RATTRAPAGE), key=lambda l: l["date"])
    lib = lambda l: f'{l["date"]} {l["debut"]}'
    for g in gains:
        if not g["date"]:
            continue
        dg = dt.date.fromisoformat(g["date"])
        cands = [p for p in perdues if p["compense"] == "Non" and p["matiere"] == g["matiere"]
                 and p["date"] and abs((dt.date.fromisoformat(p["date"]) - dg).days) <= FENETRE_JOURS]
        if not cands:
            g["lie"] = "(heure d'origine non trouvée)"
            continue
        p = min(cands, key=lambda p: (abs((dt.date.fromisoformat(p["date"]) - dg).days), p["debut"]))
        p["compense"], p["lie"], g["lie"] = "Oui", "rattrapé le " + lib(g), "compense " + lib(p)


def analyse(args, dossier):
    import openpyxl
    from openpyxl.chart import BarChart, Reference
    from openpyxl.drawing.image import Image as XLImage
    from openpyxl.styles import Alignment, Font, PatternFill
    from openpyxl.utils import get_column_letter as L
    from PIL import Image

    sem_dir = dossier / "semaines"
    fichiers = sorted(sem_dir.glob("*.json"))
    if not fichiers:
        sys.exit("Aucune semaine capturée dans " + str(sem_dir))
    lignes, semaines = [], []
    for f in fichiers:
        lundi = parse_date(f.stem)
        if args.rentree and lundi < lundi_de(parse_date(args.rentree)):
            continue
        if args.semaine and lundi > lundi_de(parse_date(args.semaine)):
            continue
        lignes += traiter_semaine(json.loads(f.read_text(encoding="utf-8")), lundi)
        semaines.append(lundi)
    if not semaines:
        sys.exit("Aucune semaine capturée entre les deux dates demandées : vérifiez les dates.")
    apparier(lignes)

    cats = []
    for c, _ in REGLES:
        if c not in cats:
            cats.append(c)
    wb = openpyxl.Workbook()
    gras = Font(bold=True, color="FFFFFF")
    fond = PatternFill("solid", fgColor="1F4E78")
    fond_tot = PatternFill("solid", fgColor="DDEBF7")

    def entete(ws, row, valeurs):
        for j, v in enumerate(valeurs, 1):
            c = ws.cell(row, j, v)
            c.font, c.fill = gras, fond
            c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        ws.row_dimensions[row].height = 45

    # ------------------------------------------------------------------ Détail
    wd = wb.active
    wd.title = "Détail"
    cols = ["Semaine", "Date", "Jour", "Début", "Fin", "Durée (h)", "Matière", "Catégorie",
            "Heure perdue compensée ?", "Lien perdue / rattrapage", "Indice détecté", "Texte brut",
            "Inclus dans les calculs ?", "Durée comptée (h)"]
    entete(wd, 1, cols)
    for i, l in enumerate(lignes, 2):
        vals = [parse_date(l["semaine"]), parse_date(l["date"]) if l["date"] else None, l["jour"],
                l["debut"], l["fin"], l["duree_h"], l["matiere"], l["categorie"],
                l["compense"], l["lie"], l["indice"], l["texte"],
                "Non" if norm(l["matiere"]).strip() in EXCLUES else "Oui", f'=IF(M{i}="Oui",F{i},0)']
        for j, v in enumerate(vals, 1):
            wd.cell(i, j, v)
        wd.cell(i, 1).number_format = wd.cell(i, 2).number_format = "dd/mm/yyyy"
    nlast = max(len(lignes) + 1, 2)
    wd.auto_filter.ref = f"A1:{L(len(cols))}{nlast}"
    wd.freeze_panes = "A2"
    for j, w in enumerate([12, 12, 11, 8, 8, 10, 28, 18, 16, 34, 24, 80, 14, 14], 1):
        wd.column_dimensions[L(j)].width = w
    rng = lambda col: f"Détail!${col}$2:${col}${nlast}"
    SEM, DUR, MAT, CAT, COMP, RAW, INC = "A", "N", "G", "H", "I", "F", "M"

    # -------------------------------------------- tableaux par semaine / matière
    ent = ["Heures prévues", "Prof absent (h)", "Annulé (h)", "Grève (h)", "Autres perdues (h)",
           "TOTAL perdues (h)", "Rattrapées : déplacé/modifié (h)", "Dont remplacées (info, h)",
           "Solde net (rattrapées − perdues, h)", "% perdu brut", "% perdu net"]

    def tableau(ws, titre_cle, cle_col, valeurs, fmt=None, brut=False):
        """brut=True : durées réelles de chaque ligne (onglet matières) ; le TOTAL ne
        retient que les lignes 'Inclus = Oui'."""
        D = RAW if brut else DUR
        entete(ws, 1, [titre_cle] + ent + (["Inclus dans le total ?", "Heures assurées (prévues − perdues, h)"]
                                           if brut else []))
        n = len(valeurs)
        for i, v in enumerate(valeurs, 2):
            c = ws.cell(i, 1, v)
            if fmt:
                c.number_format = fmt
            k = lambda crit: f'SUMIFS({rng(D)},{rng(cle_col)},$A{i},{rng(CAT)},"{crit}")'
            if brut:
                ws.cell(i, 13, "Non" if norm(v).strip() in EXCLUES else "Oui")
                if ws.cell(i, 13).value == "Non":
                    for jj in range(1, 15):
                        ws.cell(i, jj).font = Font(italic=True, color="808080")
            ws.cell(i, 2, f'=SUMIF({rng(cle_col)},$A{i},{rng(D)})-{k(RATTRAPAGE)}')
            ws.cell(i, 3, "=" + k("Prof absent"))
            ws.cell(i, 4, "=" + k("Annulé"))
            ws.cell(i, 5, "=" + k("Grève"))
            ws.cell(i, 6, "=" + "+".join(k(c) for c in PERDUES if c not in ("Prof absent", "Annulé", "Grève")))
            ws.cell(i, 7, f"=SUM(C{i}:F{i})")
            ws.cell(i, 8, "=" + k(RATTRAPAGE))
            ws.cell(i, 9, "=" + k("Remplacé"))
            ws.cell(i, 10, f"=H{i}-G{i}")
            ws.cell(i, 11, f"=IF(B{i}=0,0,G{i}/B{i})")
            ws.cell(i, 12, f"=IF(B{i}=0,0,MAX(G{i}-H{i},0)/B{i})")
            if brut:
                ws.cell(i, 14, f"=B{i}-G{i}")
        t = n + 2
        ws.cell(t, 1, "TOTAL")
        for j in range(2, 11):
            if brut:
                ws.cell(t, j, f'=SUMIF($M$2:$M${t - 1},"Oui",{L(j)}2:{L(j)}{t - 1})')
            else:
                ws.cell(t, j, f"=SUM({L(j)}2:{L(j)}{t - 1})")
        ws.cell(t, 11, f"=IF(B{t}=0,0,G{t}/B{t})")
        ws.cell(t, 12, f"=IF(B{t}=0,0,MAX(G{t}-H{t},0)/B{t})")
        if brut:
            ws.cell(t, 14, f'=SUMIF($M$2:$M${t - 1},"Oui",N2:N{t - 1})')
        for j in range(1, 13 + (2 if brut else 0)):
            ws.cell(t, j).font = Font(bold=True)
            ws.cell(t, j).fill = fond_tot
        for r in range(2, t + 1):
            for j in (2, 3, 4, 5, 6, 7, 8, 9, 10, 14):
                ws.cell(r, j).number_format = "0.0"
            if brut:      # unité dans les colonnes tracées : les étiquettes du graphique la reprennent
                ws.cell(r, 7).number_format = ws.cell(r, 14).number_format = '0.0" h"'
            ws.cell(r, 11).number_format = ws.cell(r, 12).number_format = "0.0%"
        ws.freeze_panes = "B2"
        ws.column_dimensions["A"].width = 34
        for j in range(2, 15):
            ws.column_dimensions[L(j)].width = 15
        if brut and EXCLUES:
            ws.cell(t + 1, 1, "Matières en gris : affichées mais exclues du TOTAL, de la Synthèse et de "
                              "Par semaine (colonne « Inclus dans le total ? »).").font = Font(italic=True)
        return t

    matieres = sorted({l["matiere"] for l in lignes if l["matiere"]})
    wm = wb.create_sheet("Par matière", 0)
    tm = tableau(wm, "Matière", MAT, matieres, brut=True)
    from openpyxl.chart.label import DataLabelList
    from openpyxl.chart.shapes import GraphicalProperties
    from openpyxl.chart.text import RichText
    from openpyxl.drawing.line import LineProperties
    from openpyxl.drawing.text import CharacterProperties, Paragraph, ParagraphProperties

    # ---- tableau de données dédié aux graphiques (colonnes P:T) ------------------------------
    # matières incluses seulement (la vie de classe etc. sont hors graphiques), ordre inversé
    # pour que la 1re matière soit en haut sans retourner l'axe, #N/A à la place de 0 pour ne
    # pas afficher d'étiquettes « 0,0 ».
    lignes_incl = [(i, v) for i, v in enumerate(matieres, 2) if norm(v).strip() not in EXCLUES]
    lignes_incl.reverse()
    CG = 16   # colonne P
    titres_g = ["Données des graphiques (matières incluses)", "Heures assurées", "Heures perdues",
                "% perdu brut", "% perdu net (après rattrapage)", "Total prévu (valeur à droite des barres)"]
    for j, t in enumerate(titres_g):
        c = wm.cell(1, CG + j, t)
        c.font = Font(bold=True, color="808080", size=9)
        c.alignment = Alignment(wrap_text=True, vertical="center")
    for r, (i, _) in enumerate(lignes_incl, 2):
        wm.cell(r, CG, f"=A{i}")
        wm.cell(r, CG + 1, f"=IF(N{i}<=0,NA(),N{i})").number_format = '0.0" h"'
        wm.cell(r, CG + 2, f"=IF(G{i}<=0,NA(),G{i})").number_format = '0.0" h"'
        wm.cell(r, CG + 3, f"=IF(K{i}<=0,NA(),K{i})").number_format = "0%"
        wm.cell(r, CG + 4, f"=IF(L{i}<=0,NA(),L{i})").number_format = "0%"
        wm.cell(r, CG + 5, f"=B{i}").number_format = '0.0" h"'
        for j in range(6):
            wm.cell(r, CG + j).font = Font(color="808080", size=9)
    wm.column_dimensions[L(CG)].width = 26
    for j in range(1, 6):
        wm.column_dimensions[L(CG + j)].width = 14
    fin = len(lignes_incl) + 1
    # échelle commune de l'axe des heures (nécessaire pour superposer la série « total »)
    prevu_mat = {}
    for l in lignes:
        if l["categorie"] != RATTRAPAGE:
            prevu_mat[l["matiere"]] = prevu_mat.get(l["matiere"], 0) + (l["duree_h"] or 0)
    haut = max([v for m, v in prevu_mat.items() if norm(m).strip() not in EXCLUES] or [10])
    axe_max = (int(haut * 1.18 / 5) + 1) * 5

    def graphique(titre, colonnes, axe_titre, fmt_axe, empile=False):
        ch = BarChart()
        ch.type, ch.title = "bar", titre
        for col in colonnes:
            ch.add_data(Reference(wm, min_col=col, min_row=1, max_row=fin), titles_from_data=True)
        ch.set_categories(Reference(wm, min_col=CG, min_row=2, max_row=fin))
        ch.height, ch.width = max(10, 0.85 * len(lignes_incl) + 6), 26
        ch.x_axis.delete = False          # axes visibles dans Excel récent
        ch.y_axis.delete = False
        ch.y_axis.title = axe_titre
        ch.y_axis.number_format = fmt_axe
        ch.y_axis.majorGridlines.spPr = GraphicalProperties(ln=LineProperties(solidFill="D9D9D9", w=6350))
        ch.legend.position = "t"          # en haut : ne gêne pas les graduations
        ch.gapWidth = 60
        if empile:
            ch.grouping, ch.overlap = "stacked", 100
        for serie in ch.series:
            serie.dLbls = DataLabelList()
            serie.dLbls.showVal = True
            serie.dLbls.showSerName = serie.dLbls.showCatName = serie.dLbls.showLegendKey = False
            serie.dLbls.showPercent = False
        return ch

    def couleur_etiquettes(serie, couleur):
        cp = CharacterProperties(solidFill=couleur, b=True)
        serie.dLbls.txPr = RichText(p=[Paragraph(pPr=ParagraphProperties(defRPr=cp), endParaRPr=cp)])

    # 1) heures prévues = assurées + perdues : la part rouge est le perdu sur le prévu
    ch1 = graphique("Heures prévues par matière : assurées et perdues", (CG + 1, CG + 2),
                    "Heures", '0" h"', empile=True)
    ch1.series[0].graphicalProperties.solidFill = "9DC3E6"
    ch1.series[1].graphicalProperties.solidFill = "C00000"
    couleur_etiquettes(ch1.series[1], "FFFFFF")
    ch1.y_axis.scaling.min, ch1.y_axis.scaling.max = 0, axe_max
    tot = BarChart()
    tot.type, tot.gapWidth = "bar", 60
    tot.add_data(Reference(wm, min_col=CG + 5, min_row=1, max_row=fin), titles_from_data=True)
    tot.set_categories(Reference(wm, min_col=CG, min_row=2, max_row=fin))
    st = tot.series[0]
    st.idx, st.order = 2, 2              # numéro distinct des deux séries empilées
    st.graphicalProperties.noFill = True
    st.graphicalProperties.line.noFill = True
    st.dLbls = DataLabelList()
    st.dLbls.showVal, st.dLbls.position = True, "outEnd"
    st.dLbls.showSerName = st.dLbls.showCatName = st.dLbls.showLegendKey = st.dLbls.showPercent = False
    couleur_etiquettes(st, "1F4E78")
    tot.y_axis.axId = 200                      # axe secondaire, masqué, même échelle que le principal
    tot.y_axis.scaling.min, tot.y_axis.scaling.max = 0, axe_max
    tot.y_axis.delete = True
    tot.y_axis.crosses = "max"
    ch1 += tot
    wm.add_chart(ch1, f"A{tm + 4}")
    # 2) part perdue en % du prévu
    ch2 = graphique("Part des heures prévues perdue, par matière", (CG + 3, CG + 4),
                    "% des heures prévues", "0%")
    ch2.series[0].graphicalProperties.solidFill = "C00000"
    ch2.series[1].graphicalProperties.solidFill = "F4B183"
    wm.add_chart(ch2, f"A{tm + 4 + int(ch1.height * 2) + 3}")
    wk = wb.create_sheet("Par semaine", 1)
    tableau(wk, "Semaine du", SEM, semaines, "dd/mm/yyyy")

    # --------------------------------------------------------------- Synthèse
    ws = wb.create_sheet("Synthèse", 0)
    ws["A1"] = "Analyse de l'emploi du temps Pronote"
    ws["A1"].font = Font(bold=True, size=14)
    ws["A2"] = (f"Du {min(semaines):%d/%m/%Y} au {max(semaines) + dt.timedelta(days=6):%d/%m/%Y}"
                f"  -  {len(semaines)} semaines avec cours  -  généré le {dt.date.today():%d/%m/%Y}")
    cat_h = lambda c: f'SUMIF({rng(CAT)},"{c}",{rng(DUR)})'
    bilan = [
        ("Heures affichées dans l'emploi du temps (tous créneaux)", f"=SUM({rng(DUR)})"),
        ("  dont rattrapages (cours déplacés / modifiés, créneaux ajoutés)", "=" + cat_h(RATTRAPAGE)),
        ("HEURES PRÉVUES à l'origine", "=B4-B5"),
        ("", None),
        ("Heures perdues : prof absent", "=" + cat_h("Prof absent")),
        ("Heures perdues : annulé", "=" + cat_h("Annulé")),
        ("Heures perdues : grève", "=" + cat_h("Grève")),
        ("Heures perdues : autres (dispensé, absence élève)",
         "=" + "+".join(cat_h(c) for c in PERDUES if c not in ("Prof absent", "Annulé", "Grève"))),
        ("TOTAL heures perdues", "=SUM(B8:B11)"),
        ("Heures rattrapées (déplacé / modifié)", "=B5"),
        ("SOLDE NET (rattrapées − perdues)", "=B13-B12"),
        ("", None),
        ("Heures perdues appariées à un rattrapage identifié",
         f'=SUMIF({rng(COMP)},"Oui",{rng(DUR)})'),
        ("Heures remplacées par un autre enseignant (cours assuré, info)", "=" + cat_h("Remplacé")),
        ("", None),
        ("% d'heures perdues (brut)", "=IF(B6=0,0,B12/B6)"),
        ("% d'heures perdues après rattrapage (net)", "=IF(B6=0,0,MAX(B12-B13,0)/B6)"),
    ]
    entete(ws, 3, ["Indicateur", "Heures / taux"])
    for k, (lib, f) in enumerate(bilan, 4):
        ws.cell(k, 1, lib)
        if f:
            ws.cell(k, 2, f)
            ws.cell(k, 2).number_format = "0.0%" if lib.startswith("%") else "0.0"
        if lib.isupper() or lib.startswith(("HEURES", "TOTAL", "SOLDE")):
            ws.cell(k, 1).font = ws.cell(k, 2).font = Font(bold=True)
            ws.cell(k, 1).fill = ws.cell(k, 2).fill = fond_tot
    r0 = 4 + len(bilan) + 2
    entete(ws, r0, ["Catégorie (détail)", "Nb de cours", "Heures"])
    for k, cat in enumerate(cats, r0 + 1):
        ws.cell(k, 1, cat)
        ws.cell(k, 2, f'=COUNTIFS({rng(CAT)},A{k},{rng(INC)},"Oui")')
        ws.cell(k, 3, f'=SUMIF({rng(CAT)},A{k},{rng(DUR)})').number_format = "0.0"
    kt = r0 + 1 + len(cats)
    ws.cell(kt, 1, "TOTAL").font = Font(bold=True)
    ws.cell(kt, 2, f"=SUM(B{r0 + 1}:B{kt - 1})")
    ws.cell(kt, 3, f"=SUM(C{r0 + 1}:C{kt - 1})").number_format = "0.0"
    ws.cell(kt + 2, 1, "Principe : un cours « déplacé / modifié » est le nouveau créneau d'un cours perdu "
                       "(annulé, prof absent, grève). L'outil apparie chaque rattrapage à une heure perdue "
                       "de la même matière à ±" + str(FENETRE_JOURS) + " jours (colonne « Lien » de Détail). "
                       "Les heures viennent des infobulles Pronote ; un cours dure ici 55 min = 0,92 h. "
                       "Corrigez une catégorie dans Détail : tout se met à jour.").alignment = \
        Alignment(wrap_text=True, vertical="top")
    ws.merge_cells(start_row=kt + 2, start_column=1, end_row=kt + 2, end_column=3)
    ws.row_dimensions[kt + 2].height = 90
    exclus = [l for l in lignes if norm(l["matiere"]).strip() in EXCLUES]
    if exclus:
        noms = ", ".join(sorted({l["matiere"] for l in exclus}))
        ws.cell(kt + 4, 1, f"Exclus des calculs ci-dessus (visibles dans Détail et Par matière) : {noms} - "
                           f"{len(exclus)} créneaux, {sum(l['duree_h'] or 0 for l in exclus):.1f} h "
                           f"dont {sum(l['duree_h'] or 0 for l in exclus if l['categorie'] in PERDUES):.1f} h perdues."
                ).font = Font(italic=True, color="7F6000")
    ws.column_dimensions["A"].width = 68
    ws.column_dimensions["B"].width = 16
    ws.column_dimensions["C"].width = 14

    # ---------------------------------------------------------------- Captures
    wc = wb.create_sheet("Captures")
    entete(wc, 1, ["Semaine du", "Fichier", "Aperçu"])
    wc.column_dimensions["A"].width = 14
    wc.column_dimensions["B"].width = 30
    wc.column_dimensions["C"].width = 90
    miniatures = dossier / "_miniatures"
    miniatures.mkdir(exist_ok=True)
    r = 2
    for lundi in semaines:
        png = sem_dir / f"{lundi.isoformat()}.png"
        wc.cell(r, 1, lundi).number_format = "dd/mm/yyyy"
        if png.exists():
            lien = wc.cell(r, 2, png.name)
            lien.hyperlink = str(png.resolve())
            lien.font = Font(color="0563C1", underline="single")
            im = Image.open(png)
            im.thumbnail((640, 420))
            mini = miniatures / png.name
            im.save(mini)
            xi = XLImage(str(mini))
            wc.add_image(xi, f"C{r}")
            wc.row_dimensions[r].height = xi.height * 0.75 + 4
        r += 1

    sortie = Path(args.excel) if args.excel else dossier / f"pronote_analyse_{dt.date.today():%Y%m%d}.xlsx"
    if sortie.exists():       # ne jamais écraser un classeur existant (peut être annoté à la main)
        k = 2
        while sortie.with_name(f"{sortie.stem}_v{k}{sortie.suffix}").exists():
            k += 1
        sortie = sortie.with_name(f"{sortie.stem}_v{k}{sortie.suffix}")
    capture_le = max(dt.date.fromtimestamp(f.stat().st_mtime) for f in fichiers)
    resume_fcpe = onglets_fcpe(wb, lignes, sem_dir, sortie.resolve(), min(capture_le, dt.date.today()))
    wb.save(sortie)

    from collections import Counter
    h = Counter()
    for l in lignes:
        if norm(l["matiere"]).strip() not in EXCLUES:
            h[l["categorie"]] += l["duree_h"] or 0
    perdu = sum(h[c] for c in PERDUES)
    prevu = sum(h.values()) - h[RATTRAPAGE]
    print(f"\nExcel : {sortie}")
    print(f"{len(lignes)} cours sur {len(semaines)} semaines")
    for c in cats:
        if h[c]:
            print(f"  {c:<18} {h[c]:6.1f} h")
    print(f"Heures prévues à l'origine : {prevu:.1f} h")
    print(f"Heures perdues             : {perdu:.1f} h ({perdu / prevu:.1%})" if prevu else "")
    print(f"Heures rattrapées          : {h[RATTRAPAGE]:.1f} h  -> solde net {h[RATTRAPAGE] - perdu:+.1f} h")
    print(resume_fcpe)
    sans = sum(1 for l in lignes if l["duree_h"] is None)
    if sans:
        print(f"  /!\\ {sans} cours sans horaire exploitable (durée vide)")


RACINE = Path(__file__).resolve().parent
PARAM_DEFAUT = RACINE / "parametres.ini"


def lire_parametres(chemin):
    """Lit parametres.ini (section [pronote]) ; retourne un dict (valeurs vides ignorées)."""
    import configparser
    if not chemin.exists():
        return {}
    cp = configparser.ConfigParser(inline_comment_prefixes=("#", ";"))
    cp.read(chemin, encoding="utf-8-sig")
    if not cp.has_section("pronote"):
        sys.exit(f"{chemin.name} : section [pronote] absente.")
    return {k: v.strip() for k, v in cp["pronote"].items() if v.strip()}


RENTREE_DEFAUT = "2026-08-31"


def normaliser_url(u):
    """Nettoie une adresse copiée du navigateur et vise l'espace Parents."""
    u = u.strip().strip('"').split("?")[0].split("#")[0]
    if not u:
        return ""
    if "://" not in u:
        u = "https://" + u
    if "/pronote" in u and not u.lower().endswith(".html"):
        u = u.rstrip("/") + "/parent.html"
    return u


def memoriser_url(chemin, url):
    """Écrit l'adresse dans parametres.ini en conservant les commentaires."""
    if chemin.exists():
        texte = chemin.read_text(encoding="utf-8-sig")
    else:
        texte = "[pronote]\nurl = \n"
    nouveau, n = re.subn(r"(?m)^[ \t]*url[ \t]*=.*$", lambda m: "url = " + url, texte, count=1)
    if not n:
        nouveau = texte.replace("[pronote]", "[pronote]\nurl = " + url, 1)
    chemin.write_text(nouveau, encoding="utf-8")


def demander_date(question, defaut):
    """Entrée = valeur par défaut ; sinon redemande tant que la date est illisible."""
    while True:
        rep = input(f"{question}\n   [Entrée = {defaut:%d/%m/%Y}]  ou tapez une date JJ/MM/AAAA : ").strip()
        if not rep:
            return defaut
        try:
            return parse_date(rep)
        except ValueError:
            print("   Date non comprise. Exemple attendu : 31/08/2026 (jour/mois/année).")


def questions(args, chemin_param, date_debut_fixee, date_fin_fixee):
    print(f"=== pronote_suivi {__version__} : réglage de la période ===\n")
    if not args.url:
        saisie = input("Adresse de votre Pronote (copiez-la depuis votre navigateur, puis collez-la ici)\n"
                       "   [Entrée = ouvrir un navigateur vide, par exemple pour passer par votre ENT] : ")
        args.url = normaliser_url(saisie)
        if args.url:
            memoriser_url(chemin_param, args.url)
            print(f"   Adresse enregistrée dans {chemin_param.name} : {args.url}")
        print()
    debut = parse_date(args.rentree)
    if not date_debut_fixee:
        debut = demander_date("Date de la rentrée (début de l'analyse) ?", debut)
    fin = parse_date(args.semaine)
    if not date_fin_fixee:
        auj = lundi_de(fin)
        libelle = ("Jusqu'à quelle semaine ? Par défaut la semaine d'aujourd'hui"
                   if fin == dt.date.today() else "Jusqu'à quelle semaine ?")
        while True:
            fin = demander_date(f"{libelle} (semaine du {auj:%d/%m/%Y}).", fin)
            if lundi_de(fin) >= lundi_de(debut):
                break
            print("   La fin doit être après la rentrée.")
            fin = parse_date(args.semaine)
    args.rentree, args.semaine = debut.isoformat(), fin.isoformat()
    print(f"\nAnalyse du {lundi_de(debut):%d/%m/%Y} à la semaine du {lundi_de(fin):%d/%m/%Y}.\n")


def main():
    global FENETRE_JOURS
    try:
        sys.stdout.reconfigure(encoding="utf-8", errors="replace")
    except Exception:
        pass
    p = argparse.ArgumentParser(
        description=__doc__ + "\nLes réglages se font dans parametres.ini ; les options ci-dessous "
                              "les remplacent ponctuellement.",
        formatter_class=argparse.RawTextHelpFormatter)
    p.add_argument("--param", default=str(PARAM_DEFAUT), help="fichier de paramètres (défaut : parametres.ini)")
    p.add_argument("--url")
    p.add_argument("--rentree", help="date de rentrée JJ/MM/AAAA (défaut : 31/08/2026)")
    p.add_argument("--semaine", help="semaine finale JJ/MM/AAAA (défaut : semaine en cours)")
    p.add_argument("--sans-questions", action="store_true",
                   help="ne rien demander : prendre parametres.ini et les valeurs par défaut")
    p.add_argument("--dossier")
    p.add_argument("--etape", choices=["tout", "capture", "analyse"], default="tout")
    p.add_argument("--excel", help="chemin du .xlsx de sortie")
    p.add_argument("--navigateur")
    p.add_argument("--manuel", action="store_true", default=None)
    p.add_argument("--attente", type=int)
    p.add_argument("--selecteur-suivant")
    p.add_argument("--selecteur-cours")
    p.add_argument("--fenetre-jours", type=int)
    p.add_argument("--matieres-exclues", help="matières hors calculs, séparées par ; (vide = aucune)")
    p.add_argument("--version", action="version", version=f"pronote_suivi {__version__}")
    args = p.parse_args()

    debut_cli, fin_cli = bool(args.rentree), bool(args.semaine)
    cfg = lire_parametres(Path(args.param))
    pick = lambda cli, cle, defaut=None: cli if cli not in (None, "") else cfg.get(cle, defaut)
    args.url = pick(args.url, "url")
    args.rentree = pick(args.rentree, "rentree", RENTREE_DEFAUT)
    args.semaine = pick(args.semaine, "semaine") or dt.date.today().isoformat()
    args.dossier = pick(args.dossier, "dossier", "sortie")
    args.navigateur = pick(args.navigateur, "navigateur")
    args.attente = int(pick(args.attente, "attente_ms", 1500))
    args.selecteur_suivant = pick(args.selecteur_suivant, "selecteur_suivant")
    args.selecteur_cours = pick(args.selecteur_cours, "selecteur_cours")
    args.manuel = bool(args.manuel) or cfg.get("mode_manuel", "non").lower() in ("oui", "1", "true", "vrai")
    FENETRE_JOURS = int(pick(args.fenetre_jours, "fenetre_jours", FENETRE_JOURS))
    if "matieres_exclues" in cfg or args.matieres_exclues is not None:
        brut = pick(args.matieres_exclues, "matieres_exclues", "")
        EXCLUES[:] = [norm(m).strip() for m in re.split(r"[;,]", brut) if m.strip()]

    args.url = normaliser_url(args.url or "")
    FCPE["etablissement"] = cfg.get("etablissement", "")
    FCPE["ville"] = cfg.get("ville", "")
    FCPE["classe"] = cfg.get("classe") or cfg.get("niveau", "")
    FCPE["regroupement"] = "periode" if norm(cfg.get("regroupement_fcpe", "")).startswith("period") else "matiere"
    if cfg.get("motifs_fcpe"):
        FCPE["motifs"] = [norm(m).strip() for m in re.split(r"[;,]", cfg["motifs_fcpe"]) if m.strip()]
    dossier = Path(args.dossier)
    if not dossier.is_absolute():
        dossier = RACINE / dossier
    dossier.mkdir(parents=True, exist_ok=True)
    try:
        import openpyxl, PIL                                  # noqa: F401
        if args.etape != "analyse":
            import playwright                                 # noqa: F401
    except ImportError as e:
        sys.exit(f"Un composant n'est pas installé ({e.name}). Dans une invite de commandes ouverte dans "
                 "le dossier de l'outil, tapez :\n    python -m pip install -r requirements.txt")
    try:
        parse_date(args.rentree), parse_date(args.semaine)
    except ValueError:
        sys.exit("Date illisible dans parametres.ini ou en option : écrivez JJ/MM/AAAA (ex. 31/08/2026).")
    if args.etape in ("tout", "capture"):
        if not args.sans_questions:
            questions(args, Path(args.param), debut_cli, fin_cli)
        capture(args, dossier)
    if args.etape in ("tout", "analyse"):
        analyse(args, dossier)


if __name__ == "__main__":
    try:
        main()
    except (KeyboardInterrupt, EOFError):
        print("\nArrêt demandé.")
