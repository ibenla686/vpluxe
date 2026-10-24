/* =====================================================================
   VPLuxe — application du site (sans dépendance serveur)
   - Charge la configuration (data/site.json), les maisons (data/maisons.json),
     les ventes (data/ventes.json) et le catalogue (data/catalogue.xlsx, puis
     catalogue.csv, puis catalogue.json en secours).
   - Rend l'en-tête, le pied de page, le menu mobile, la recherche, le panier.
   - Pilote chaque page selon l'attribut data-page du <body>.
   ===================================================================== */
(function () {
  "use strict";

  /* ------------------------------------------------------------ utilitaires */
  const $ = (sel, root) => (root || document).querySelector(sel);
  const $$ = (sel, root) => Array.from((root || document).querySelectorAll(sel));
  const esc = (s) => String(s == null ? "" : s).replace(/[&<>"']/g, (c) => ({ "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;", "'": "&#39;" }[c]));
  const slug = (s) => String(s || "").normalize("NFKD").replace(/[̀-ͯ]/g, "").replace(/[^a-zA-Z0-9]+/g, "-").replace(/^-|-$/g, "").toLowerCase();
  const norm = (s) => String(s || "").normalize("NFKD").replace(/[̀-ͯ]/g, "").toLowerCase().trim();
  const oui = (v) => /^(oui|yes|true|1|x)$/i.test(String(v || "").trim());
  const num = (v) => { if (v === "" || v == null) return null; const n = parseFloat(String(v).replace(/\s/g, "").replace(",", ".")); return isNaN(n) ? null : n; };
  const params = new URLSearchParams(location.search);
  const fmtPrix = (n) => n == null ? "" : new Intl.NumberFormat("fr-FR", { style: "currency", currency: "EUR", maximumFractionDigits: 0 }).format(n);
  const fmtDate = (d) => d ? new Intl.DateTimeFormat("fr-FR", { day: "numeric", month: "long" }).format(d) : "";
  const parseDate = (v) => {
    if (!v && v !== 0) return null;
    if (v instanceof Date) return isNaN(v) ? null : v;
    if (typeof v === "number") { const d = new Date(Math.round((v - 25569) * 86400 * 1000)); return isNaN(d) ? null : new Date(d.getUTCFullYear(), d.getUTCMonth(), d.getUTCDate()); }
    const s = String(v).trim(); let m;
    if ((m = s.match(/^(\d{4})-(\d{1,2})-(\d{1,2})/))) return new Date(+m[1], +m[2] - 1, +m[3]);
    if ((m = s.match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})/))) return new Date(+m[3], +m[2] - 1, +m[1]);
    const d = new Date(s); return isNaN(d) ? null : d;
  };
  const store = {
    get(k, def) { try { const v = localStorage.getItem(k); return v ? JSON.parse(v) : def; } catch (e) { return def; } },
    set(k, v) { try { localStorage.setItem(k, JSON.stringify(v)); } catch (e) { /* stockage indisponible */ } },
  };
  const hash = (s) => { let h = 0; for (let i = 0; i < s.length; i++) h = (h * 31 + s.charCodeAt(i)) | 0; return Math.abs(h); };

  /* ------------------------------------------------------------ données */
  const D = { site: null, maisons: [], ventes: [], produits: [], parRef: new Map(), maisonParSlug: new Map() };

  async function json(url) { const r = await fetch(url, { cache: "no-cache" }); if (!r.ok) throw new Error(url + " : " + r.status); return r.json(); }

  async function chargerCatalogue() {
    // 1) Excel (SheetJS)
    if (window.XLSX) {
      try {
        const r = await fetch("data/catalogue.xlsx", { cache: "no-cache" });
        if (r.ok) {
          const wb = XLSX.read(await r.arrayBuffer(), { type: "array", cellDates: true });
          const feuille = wb.Sheets["Catalogue"] || wb.Sheets[wb.SheetNames[0]];
          const lignes = XLSX.utils.sheet_to_json(feuille, { defval: "" });
          let ventes = [];
          if (wb.Sheets["Ventes"]) ventes = XLSX.utils.sheet_to_json(wb.Sheets["Ventes"], { defval: "" });
          if (lignes.length) return { produits: lignes, ventes };
        }
      } catch (e) { console.warn("Catalogue Excel illisible, passage au CSV.", e); }
    }
    // 2) CSV
    try {
      const r = await fetch("data/catalogue.csv", { cache: "no-cache" });
      if (r.ok) { const lignes = parseCSV(await r.text()); if (lignes.length) return { produits: lignes, ventes: null }; }
    } catch (e) { console.warn("CSV illisible, passage au JSON.", e); }
    // 3) JSON
    return { produits: await json("data/catalogue.json"), ventes: null };
  }

  function parseCSV(texte) {
    texte = texte.replace(/^﻿/, "");
    const premiere = texte.split(/\r?\n/)[0] || "";
    const sep = (premiere.match(/;/g) || []).length >= (premiere.match(/,/g) || []).length ? ";" : ",";
    const lignes = []; let champ = "", ligne = [], q = false;
    for (let i = 0; i < texte.length; i++) {
      const c = texte[i];
      if (q) { if (c === '"') { if (texte[i + 1] === '"') { champ += '"'; i++; } else q = false; } else champ += c; }
      else if (c === '"') q = true;
      else if (c === sep) { ligne.push(champ); champ = ""; }
      else if (c === "\n" || c === "\r") { if (c === "\r" && texte[i + 1] === "\n") i++; ligne.push(champ); lignes.push(ligne); ligne = []; champ = ""; }
      else champ += c;
    }
    if (champ !== "" || ligne.length) { ligne.push(champ); lignes.push(ligne); }
    const entetes = (lignes.shift() || []).map((h) => h.trim());
    return lignes.filter((l) => l.some((v) => v !== "")).map((l) => { const o = {}; entetes.forEach((h, i) => (o[h] = l[i] == null ? "" : l[i])); return o; });
  }

  const PLACEHOLDERS = [
    [/sac porte main|sac a main|sac a rabat|rabat/, "sac-main"], [/bandoul|porte croise|mini sac/, "sac-bandouliere"], [/cabas|tote/, "cabas"],
    [/portefeuille|porte-cartes|porte cartes|pochette/, "portefeuille"], [/manteau|trench|pardessus/, "manteau"], [/doudoune|parka/, "doudoune"],
    [/veste|blazer|blouson/, "veste"], [/pull|cardigan|maille|sweat/, "pull"], [/robe|jupe/, "robe"], [/sneaker|basket|tennis/, "sneakers"],
    [/mocassin|derby|richelieu|loafer/, "mocassins"], [/escarpin|sandale|mule/, "escarpins"], [/botte|bottine|boots/, "bottes"],
    [/ceinture/, "ceinture"], [/lunettes/, "lunettes"], [/carre|foulard|etole|bandana/, "foulard"], [/echarpe|snood|bonnet/, "echarpe"],
    [/valise|trolley/, "valise"], [/voyage|week-end|weekend|duffle/, "sac-voyage"],
  ];
  const PAR_CATEGORIE = { maroquinerie: "sac-main", "pret-a-porter": "veste", souliers: "mocassins", accessoires: "ceinture", bagagerie: "valise" };

  function placeholderPour(p) {
    // le type (sous-catégorie) prime sur le nom, pour éviter les faux amis (« lunettes carrées » ≠ carré de soie)
    for (const texte of [norm(p.type), norm(p.nom)]) { if (!texte) continue; for (const [re, k] of PLACEHOLDERS) if (re.test(texte)) return k; }
    return PAR_CATEGORIE[slug(p.categorie)] || "generique";
  }

  function normaliserProduit(l) {
    const g = (k) => { const key = Object.keys(l).find((x) => norm(x) === norm(k)); return key == null ? "" : l[key]; };
    const p = {
      ref: String(g("Référence") || "").trim(), maison: String(g("Maison") || "").trim(), vente: String(g("Vente") || "").trim(),
      nom: String(g("Nom") || "").trim(), categorie: String(g("Catégorie") || "").trim(), type: String(g("Type") || "").trim(),
      genre: String(g("Genre") || "Mixte").trim() || "Mixte", prix: num(g("Prix")), ancienPrix: num(g("Ancien prix")),
      couleur: String(g("Couleur") || "").trim(), matiere: String(g("Matière") || "").trim(), dimensions: String(g("Dimensions") || "").trim(),
      tailles: String(g("Tailles") || "").split(/[;,]/).map((s) => s.trim()).filter(Boolean),
      stock: num(g("Stock")), nouveaute: oui(g("Nouveauté")), miseEnAvant: oui(g("Mise en avant")),
      description: String(g("Description") || "").trim(), lienPaiement: String(g("Lien paiement") || "").trim(),
      statut: String(g("Statut") || "Publié").trim(),
    };
    if (!p.ref || !p.nom) return null;
    if (/brouillon/i.test(p.statut)) return null;
    p.maisonSlug = slug(p.maison);
    const imgs = String(g("Images") || "").split(/[;\n]/).map((s) => s.trim()).filter(Boolean);
    if (imgs.length) { p.images = imgs.map((i) => /^https?:\/\//i.test(i) ? i : "assets/img/produits/" + i); p.placeholder = false; }
    else {
      const k = placeholderPour(p); const d = hash(p.ref) % 2 === 0 ? "detail-1" : "detail-2";
      p.images = ["assets/img/placeholders/" + k + ".svg", "assets/img/placeholders/" + d + ".svg", "assets/img/placeholders/ecrin.svg"]; p.placeholder = true;
    }
    p.remise = p.ancienPrix && p.prix && p.ancienPrix > p.prix ? Math.round((1 - p.prix / p.ancienPrix) * 100) : 0;
    p.epuise = p.stock === 0;
    p.recherche = norm([p.ref, p.maison, p.nom, p.categorie, p.type, p.couleur, p.matiere, p.genre].join(" "));
    return p;
  }

  function normaliserVente(v) {
    const g = (k) => { const key = Object.keys(v).find((x) => norm(x) === norm(k)); return key == null ? "" : v[key]; };
    const o = { code: String(g("Code vente") || "").trim(), maison: String(g("Maison") || "").trim(), titre: String(g("Titre") || "").trim(),
      debut: parseDate(g("Début")), fin: parseDate(g("Fin")), remise: String(g("Remise") || "").trim(), statut: String(g("Statut") || "Publié").trim() };
    if (!o.maison || /brouillon/i.test(o.statut)) return null;
    o.maisonSlug = slug(o.maison);
    const auj = new Date(); auj.setHours(0, 0, 0, 0);
    const finJ = o.fin ? new Date(o.fin.getFullYear(), o.fin.getMonth(), o.fin.getDate(), 23, 59, 59) : null;
    if (o.debut && o.debut > auj) o.etat = "a-venir";
    else if (finJ && finJ < new Date()) o.etat = "terminee";
    else o.etat = "en-cours";
    o.joursRestants = finJ ? Math.max(0, Math.ceil((finJ - new Date()) / 86400000)) : null;
    o.joursAvant = o.debut ? Math.max(0, Math.ceil((o.debut - new Date()) / 86400000)) : null;
    return o;
  }

  let chargement = null;
  function charger() {
    if (chargement) return chargement;
    chargement = (async () => {
      const [site, maisons, ventesJson, cat] = await Promise.all([json("data/site.json"), json("data/maisons.json"), json("data/ventes.json").catch(() => []), chargerCatalogue()]);
      D.site = site;
      D.maisons = maisons.map((m) => ({ ...m, slug: m.slug || slug(m.nom) }));
      D.maisons.forEach((m) => D.maisonParSlug.set(m.slug, m));
      D.produits = cat.produits.map(normaliserProduit).filter(Boolean);
      D.produits.forEach((p) => D.parRef.set(p.ref, p));
      D.ventes = (cat.ventes || ventesJson || []).map(normaliserVente).filter(Boolean);
      // nombre de pièces par maison
      D.maisons.forEach((m) => { m.nb = D.produits.filter((p) => p.maisonSlug === m.slug).length; });
      // maisons présentes au catalogue mais absentes du référentiel
      D.produits.forEach((p) => { if (!D.maisonParSlug.has(p.maisonSlug)) { const m = { slug: p.maisonSlug, nom: p.maison, pays: "", univers: [], presentation: "", nb: 0 }; D.maisons.push(m); D.maisonParSlug.set(m.slug, m); } });
      D.maisons.forEach((m) => { m.nb = D.produits.filter((p) => p.maisonSlug === m.slug).length; m.vente = D.ventes.find((v) => v.maisonSlug === m.slug && v.etat !== "terminee") || null; });
      D.produits.forEach((p) => { p.venteObj = D.ventes.find((v) => v.code && v.code === p.vente) || D.ventes.find((v) => v.maisonSlug === p.maisonSlug && v.etat === "en-cours") || null; });
      return D;
    })();
    return chargement;
  }

  /* ------------------------------------------------------------ panier */
  const Panier = {
    cle: "vpluxe_panier",
    lire() { const l = store.get(this.cle, []); return Array.isArray(l) ? l : []; },
    ecrire(l) { store.set(this.cle, l); majBadge(); },
    ajouter(ref, taille, qte) {
      const l = this.lire(); const i = l.findIndex((x) => x.ref === ref && (x.taille || "") === (taille || ""));
      if (i >= 0) l[i].qte = Math.min(9, (l[i].qte || 1) + (qte || 1)); else l.push({ ref, taille: taille || "", qte: qte || 1 });
      this.ecrire(l);
    },
    retirer(ref, taille) { this.ecrire(this.lire().filter((x) => !(x.ref === ref && (x.taille || "") === (taille || "")))); },
    quantite(ref, taille, qte) { const l = this.lire(); const x = l.find((y) => y.ref === ref && (y.taille || "") === (taille || "")); if (x) { x.qte = Math.max(1, Math.min(9, qte)); this.ecrire(l); } },
    vider() { this.ecrire([]); },
    nombre() { return this.lire().reduce((s, x) => s + (x.qte || 1), 0); },
  };
  function majBadge() { $$(".badge-panier").forEach((b) => { const n = Panier.nombre(); b.textContent = n || ""; if (n) b.removeAttribute("data-zero"); else b.setAttribute("data-zero", ""); }); }

  /* ------------------------------------------------------------ gabarit commun */
  const ICONES = {
    loupe: '<svg viewBox="0 0 24 24" width="18" height="18" fill="none" stroke="currentColor" stroke-width="1.4"><circle cx="11" cy="11" r="7"/><path d="M16.5 16.5 21 21"/></svg>',
    sac: '<svg viewBox="0 0 24 24" width="18" height="18" fill="none" stroke="currentColor" stroke-width="1.4"><path d="M5 8h14l-1 13H6z"/><path d="M9 8V6a3 3 0 0 1 6 0v2"/></svg>',
    fleche: '<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.4"><path d="M4 12h15M13 6l6 6-6 6"/></svg>',
    auth: '<svg viewBox="0 0 24 24"><path d="M12 3l7 3v6c0 4.5-3 8-7 9-4-1-7-4.5-7-9V6z"/><path d="M9 12l2 2 4-4"/></svg>',
    livraison: '<svg viewBox="0 0 24 24"><path d="M3 7h11v9H3zM14 10h4l3 3v3h-7z"/><circle cx="7" cy="18" r="1.6"/><circle cx="17" cy="18" r="1.6"/></svg>',
    retour: '<svg viewBox="0 0 24 24"><path d="M9 14 4 9l5-5"/><path d="M4 9h10a6 6 0 0 1 0 12H8"/></svg>',
    conseil: '<svg viewBox="0 0 24 24"><path d="M4 5h16v11H9l-5 4z"/><path d="M8 9h8M8 12h5"/></svg>',
    neuf: '<svg viewBox="0 0 24 24"><path d="M12 3l2.5 5.5L20 9l-4 4 1 6-5-3-5 3 1-6-4-4 5.5-.5z"/></svg>',
  };

  function rendreGabarit() {
    const s = D.site; const page = document.body.dataset.page || "";
    const nav = [["ventes.html", "Ventes", "ventes"], ["maisons.html", "Maisons", "maisons"], ["catalogue.html?genre=Femme", "Femme", "femme"], ["catalogue.html?genre=Homme", "Homme", "homme"], ["catalogue.html", "Catalogue", "catalogue"]];
    const lien = (h, t, k) => `<a href="${h}" ${page === k ? 'aria-current="page"' : ""}>${t}</a>`;
    const tete = $("#entete");
    if (tete) tete.innerHTML = `
      <div class="annonce"><span>Ventes privées réservées aux membres</span><span>Pièces neuves et authentiques</span><span>Livraison assurée en France et en Europe</span></div>
      <header class="entete" id="barre"><div class="container entete__inner">
        <div style="display:flex;align-items:center;gap:8px"><button class="burger" aria-label="Ouvrir le menu" id="ouvrir-menu"><span></span></button>
        <nav class="entete__nav" aria-label="Navigation principale">${nav.map((n) => lien(...n)).join("")}</nav></div>
        <a class="logo" href="index.html" aria-label="${esc(s.nom)} — accueil">${esc(s.nom)}<small>${esc(s.signature || "")}</small></a>
        <div class="entete__actions">
          <a class="only-desktop" href="invitation.html">Invitation</a>
          <button id="ouvrir-recherche" aria-label="Rechercher" style="display:inline-flex;align-items:center;gap:8px">${ICONES.loupe}<span class="only-desktop">Rechercher</span></button>
          <a href="panier.html" aria-label="Panier" style="display:inline-flex;align-items:center;gap:8px">${ICONES.sac}<span class="only-desktop">Panier</span><span class="badge badge-panier" data-zero></span></a>
        </div></div></header>
      <div class="menu-mobile" id="menu-mobile" aria-hidden="true">
        <div class="menu-mobile__top"><span class="logo">${esc(s.nom)}</span><button class="fermer" id="fermer-menu" aria-label="Fermer">×</button></div>
        <nav>${nav.map((n) => lien(...n)).join("")}</nav>
        <div class="menu-mobile__sec"><a href="invitation.html">Demander une invitation</a><a href="authenticite.html">Authenticité & garanties</a><a href="contact.html">Contact</a></div>
      </div>
      <div class="recherche" id="recherche" role="dialog" aria-label="Recherche"><div class="recherche__box">
        <form action="catalogue.html" method="get"><input type="search" name="q" placeholder="Rechercher une maison, une pièce…" autocomplete="off" aria-label="Rechercher"><button class="btn" type="submit">Rechercher</button></form>
        <div class="recherche__sugg"><a href="catalogue.html?nouveautes=1">Nouveautés</a><a href="catalogue.html?categorie=Maroquinerie">Maroquinerie</a><a href="catalogue.html?categorie=Souliers">Souliers</a><a href="catalogue.html?categorie=Pr%C3%AAt-%C3%A0-porter">Prêt-à-porter</a><a href="ventes.html">Ventes en cours</a></div>
      </div></div>`;

    const pied = $("#pied");
    if (pied) pied.innerHTML = `<footer class="pied"><div class="container">
      <div class="pied__haut">
        <div><span class="logo">${esc(s.nom)}<small>${esc(s.signature || "")}</small></span><p class="pied__pitch">${esc(s.sous_accroche || "")}</p></div>
        <div><h4>Découvrir</h4><ul><li><a href="ventes.html">Ventes en cours</a></li><li><a href="maisons.html">Les maisons</a></li><li><a href="catalogue.html?nouveautes=1">Nouveautés</a></li><li><a href="catalogue.html">Tout le catalogue</a></li></ul></div>
        <div><h4>Service</h4><ul><li><a href="invitation.html">Devenir membre</a></li><li><a href="authenticite.html">Authenticité & garanties</a></li><li><a href="cgv.html#livraison">Livraison & retours</a></li><li><a href="contact.html">Contact</a></li></ul></div>
        <div><h4>Informations</h4><ul><li><a href="a-propos.html">À propos</a></li><li><a href="mentions-legales.html">Mentions légales</a></li><li><a href="cgv.html">Conditions générales de vente</a></li><li><a href="confidentialite.html">Confidentialité</a></li></ul></div>
      </div>
      <div class="pied__bas"><p>© ${new Date().getFullYear()} ${esc(s.nom)}. ${esc(s.nom)} est un distributeur indépendant. Les noms des maisons sont cités uniquement pour identifier les pièces proposées et demeurent la propriété de leurs titulaires respectifs.</p><p>${esc(s.adresse || "")}</p></div>
    </div></footer><div class="toast" id="toast" role="status" aria-live="polite"></div>`;

    // titres de page
    const t = document.querySelector("title"); if (t && t.dataset.titre) document.title = t.dataset.titre + " · " + s.nom;
    $$("[data-site]").forEach((el) => { const v = el.dataset.site.split(".").reduce((o, k) => (o ? o[k] : undefined), s); if (v != null && v !== "") el.textContent = v; });
    $$("[data-site-href]").forEach((el) => { const k = el.dataset.siteHref; if (k === "email" && s.email) el.href = "mailto:" + s.email; if (k === "telephone" && s.telephone) el.href = "tel:" + s.telephone.replace(/\s/g, ""); if (k === "whatsapp" && s.whatsapp) el.href = "https://wa.me/" + s.whatsapp.replace(/\D/g, ""); if (k === "instagram") { if (s.instagram) el.href = s.instagram; else el.style.display = "none"; } });

    // interactions
    const barre = $("#barre"); const surScroll = () => barre && barre.classList.toggle("is-scrolled", scrollY > 8); addEventListener("scroll", surScroll, { passive: true }); surScroll();
    const menu = $("#menu-mobile"); $("#ouvrir-menu") && $("#ouvrir-menu").addEventListener("click", () => { menu.classList.add("is-open"); menu.setAttribute("aria-hidden", "false"); });
    $("#fermer-menu") && $("#fermer-menu").addEventListener("click", () => { menu.classList.remove("is-open"); menu.setAttribute("aria-hidden", "true"); });
    const rech = $("#recherche"); $("#ouvrir-recherche") && $("#ouvrir-recherche").addEventListener("click", () => { rech.classList.add("is-open"); setTimeout(() => $("input", rech).focus(), 50); });
    rech && rech.addEventListener("click", (e) => { if (e.target === rech) rech.classList.remove("is-open"); });
    addEventListener("keydown", (e) => { if (e.key === "Escape") { rech && rech.classList.remove("is-open"); menu && menu.classList.remove("is-open"); } });
    majBadge();
  }

  function toast(msg) { const t = $("#toast"); if (!t) return; t.textContent = msg; t.classList.add("is-visible"); clearTimeout(t._h); t._h = setTimeout(() => t.classList.remove("is-visible"), 2600); }

  /* ------------------------------------------------------------ composants */
  function carte(p) {
    const etiquettes = [];
    if (p.epuise) etiquettes.push('<span class="etiquette etiquette--epuise">Épuisé</span>');
    else if (p.remise) etiquettes.push(`<span class="etiquette etiquette--or">−${p.remise} %</span>`);
    if (p.nouveaute && !p.epuise) etiquettes.push('<span class="etiquette">Nouveauté</span>');
    if (p.venteObj && p.venteObj.etat === "en-cours" && p.venteObj.joursRestants != null) etiquettes.push(`<span class="etiquette">Vente · J-${p.venteObj.joursRestants}</span>`);
    return `<a class="carte apparait" href="produit.html?ref=${encodeURIComponent(p.ref)}">
      <div class="carte__visuel"><img src="${esc(p.images[0])}" alt="${esc(p.nom)} — ${esc(p.maison)}" loading="lazy" class="${p.placeholder ? "is-placeholder" : ""}" onerror="this.onerror=null;this.src='assets/img/placeholders/generique.svg';this.classList.add('is-placeholder')">
        <div class="carte__etiquettes">${etiquettes.join("")}</div><div class="carte__cta">Voir la pièce</div></div>
      <div class="carte__corps"><div class="carte__maison">${esc(p.maison)}</div><div class="carte__nom">${esc(p.nom)}</div>
        <div class="carte__prix"><span>${fmtPrix(p.prix)}</span>${p.ancienPrix && p.ancienPrix > (p.prix || 0) ? `<s>${fmtPrix(p.ancienPrix)}</s>` : ""}</div></div></a>`;
  }

  function carteVente(v, compact) {
    const m = D.maisonParSlug.get(v.maisonSlug) || { nom: v.maison, nb: 0, univers: [] };
    const etat = compact
      ? (v.etat === "en-cours" ? `Jusqu'au ${fmtDate(v.fin)}${v.joursRestants != null ? ` · J-${v.joursRestants}` : ""}` : v.etat === "a-venir" ? `Ouvre le ${fmtDate(v.debut)}` : `Terminée le ${fmtDate(v.fin)}`)
      : (v.etat === "en-cours" ? `En cours · se termine le ${fmtDate(v.fin)}${v.joursRestants != null ? ` · J-${v.joursRestants}` : ""}` : v.etat === "a-venir" ? `Ouvre le ${fmtDate(v.debut)}${v.joursAvant ? ` · dans ${v.joursAvant} jour${v.joursAvant > 1 ? "s" : ""}` : ""}` : `Terminée le ${fmtDate(v.fin)}`);
    return `<a class="vente ${compact ? "vente--compacte" : ""} vente--${v.etat} apparait" href="maison.html?m=${encodeURIComponent(v.maisonSlug)}">
      <div class="vente__etat">${etat}</div>
      <h3 class="vente__maison">${esc(m.nom)}</h3>
      <div class="vente__titre">${esc(v.titre || "Vente privée")}</div>
      <div class="vente__meta"><span>${m.nb} pièce${m.nb > 1 ? "s" : ""}</span>${v.remise ? `<span>${esc(v.remise)}</span>` : ""}</div>
      <span class="vente__lien">${v.etat === "a-venir" ? "Découvrir la maison" : "Accéder à la vente"} ${ICONES.fleche}</span></a>`;
  }

  /* ------------------------------------------------------------ pages */
  const Pages = {};

  Pages.accueil = function () {
    const enCours = D.ventes.filter((v) => v.etat === "en-cours").sort((a, b) => (a.fin || 0) - (b.fin || 0));
    const aVenir = D.ventes.filter((v) => v.etat === "a-venir").sort((a, b) => (a.debut || 0) - (b.debut || 0));
    const z = $("#ventes-en-cours"); if (z) z.innerHTML = enCours.length ? enCours.slice(0, 6).map((v) => carteVente(v)).join("") : `<div class="vide"><h3>Prochaines ventes en préparation</h3><p>Demandez une invitation pour être prévenu de l'ouverture.</p></div>`;
    const y = $("#ventes-a-venir"); if (y) y.innerHTML = aVenir.slice(0, 6).map((v) => carteVente(v, true)).join("");
    const vedettes = D.produits.filter((p) => p.miseEnAvant && !p.epuise); const sel = (vedettes.length >= 8 ? vedettes : vedettes.concat(D.produits.filter((p) => !p.miseEnAvant && !p.epuise))).slice(0, 8);
    const g = $("#selection"); if (g) g.innerHTML = sel.map(carte).join("");
    const nv = D.produits.filter((p) => p.nouveaute && !p.epuise).slice(0, 8); const n = $("#nouveautes"); if (n) n.innerHTML = nv.map(carte).join("");
    const cats = [["Maroquinerie", "sac-main"], ["Prêt-à-porter", "manteau"], ["Souliers", "escarpins"], ["Accessoires", "foulard"], ["Bagagerie", "valise"]];
    const t = $("#tuiles"); if (t) t.innerHTML = cats.map(([c, k]) => { const nb = D.produits.filter((p) => norm(p.categorie) === norm(c)).length; return `<a class="tuile" href="catalogue.html?categorie=${encodeURIComponent(c)}"><img src="assets/img/placeholders/${k}.svg" alt=""><div class="tuile__titre">${c}</div><div class="tuile__sous">${nb} pièce${nb > 1 ? "s" : ""}</div></a>`; }).join("");
    const bandeau = $("#bandeau-maisons"); if (bandeau) { const l = D.maisons.map((m) => `<a href="maison.html?m=${m.slug}">${esc(m.nom)}</a>`).join(""); bandeau.innerHTML = l + l; }
    const idx = $("#index-maisons"); if (idx) idx.innerHTML = D.maisons.slice().sort((a, b) => a.nom.localeCompare(b.nom, "fr")).map((m) => `<a href="maison.html?m=${m.slug}">${esc(m.nom)}</a>`).join("");
    const ch = $("#chiffre-maisons"); if (ch) ch.innerHTML = `${D.maisons.length}<sup>maisons</sup>`;
    const cp = $("#chiffre-pieces"); if (cp) cp.textContent = D.produits.length;
    const cv = $("#chiffre-ventes"); if (cv) cv.textContent = enCours.length;
  };

  Pages.ventes = function () {
    const enCours = D.ventes.filter((v) => v.etat === "en-cours").sort((a, b) => (a.fin || 0) - (b.fin || 0));
    const aVenir = D.ventes.filter((v) => v.etat === "a-venir").sort((a, b) => (a.debut || 0) - (b.debut || 0));
    const term = D.ventes.filter((v) => v.etat === "terminee").sort((a, b) => (b.fin || 0) - (a.fin || 0));
    $("#ventes-en-cours").innerHTML = enCours.length ? enCours.map((v) => carteVente(v)).join("") : `<div class="vide"><h3>Aucune vente en cours</h3><p>Les prochaines ventes sont annoncées ci-dessous.</p></div>`;
    $("#ventes-a-venir").innerHTML = aVenir.length ? aVenir.map((v) => carteVente(v, true)).join("") : `<p class="muted">Aucune vente annoncée pour le moment.</p>`;
    const zt = $("#ventes-terminees"); if (zt) zt.innerHTML = term.length ? term.slice(0, 6).map((v) => carteVente(v, true)).join("") : "";
    const st = $("#section-terminees"); if (st && !term.length) st.style.display = "none";
    $("#nb-en-cours") && ($("#nb-en-cours").textContent = enCours.length);
  };

  Pages.maisons = function () {
    const lettres = "ABCDEFGHIJKLMNOPQRSTUVWXYZ".split("");
    const groupes = {}; D.maisons.slice().sort((a, b) => a.nom.localeCompare(b.nom, "fr")).forEach((m) => { const L = norm(m.nom)[0].toUpperCase(); (groupes[L] = groupes[L] || []).push(m); });
    $("#lettres").innerHTML = lettres.map((L) => `<a href="#lettre-${L}" class="${groupes[L] ? "" : "is-vide"}">${L}</a>`).join("");
    $("#groupes").innerHTML = Object.keys(groupes).sort().map((L) => `<section class="groupe-lettre" id="lettre-${L}"><div class="groupe-lettre__lettre">${L}</div><div class="maisons-liste">${groupes[L].map((m) => `
      <a class="maison-carte" href="maison.html?m=${m.slug}"><h3>${esc(m.nom)}</h3><div class="pays">${esc(m.pays || "")}${m.vente ? ` · <strong>${m.vente.etat === "en-cours" ? "Vente en cours" : "Vente à venir"}</strong>` : ""}</div><p>${esc(m.presentation || "")}</p>
      <div class="univers">${(m.univers || []).map((u) => `<span>${esc(u)}</span>`).join("")}<span>${m.nb} pièce${m.nb > 1 ? "s" : ""}</span></div></a>`).join("")}</div></section>`).join("");
    $("#nb-maisons") && ($("#nb-maisons").textContent = D.maisons.length);
  };

  Pages.maison = function () {
    const m = D.maisonParSlug.get(params.get("m") || "");
    const racine = $("#espace-maison");
    if (!m) { racine.innerHTML = `<div class="container vide"><h3>Maison introuvable</h3><p><a class="link" href="maisons.html">Voir toutes les maisons</a></p></div>`; return; }
    document.title = m.nom + " · " + D.site.nom;
    const produits = D.produits.filter((p) => p.maisonSlug === m.slug);
    const v = m.vente; const cats = Array.from(new Set(produits.map((p) => p.categorie))).filter(Boolean);
    const remiseMax = Math.max(0, ...produits.map((p) => p.remise || 0));
    const etat = v ? (v.etat === "en-cours" ? `<span class="pastille pastille--vert"></span>Vente en cours · jusqu'au ${fmtDate(v.fin)} · J-${v.joursRestants}` : `<span class="pastille pastille--or"></span>Prochaine vente · du ${fmtDate(v.debut)} au ${fmtDate(v.fin)}`) : `<span class="pastille"></span>Sélection permanente`;
    racine.innerHTML = `
      <section class="maison-hero"><div class="container">
        <nav class="fil" aria-label="Fil d'Ariane"><a href="index.html">Accueil</a><span><a href="maisons.html">Maisons</a></span><span>${esc(m.nom)}</span></nav>
        <div class="maison-hero__grid">
          <div><div class="eyebrow">${esc(m.pays || "Maison")}${(m.univers || []).length ? " · " + esc(m.univers.join(" · ")) : ""}</div>
            <h1 class="display">${esc(m.nom)}</h1>
            <p class="lead">${esc(m.presentation || "")}</p>
            <div class="maison-hero__etat">${etat}</div>
            <div class="hero__cta"><a class="btn" href="#pieces">Voir les ${produits.length} pièce${produits.length > 1 ? "s" : ""}</a><a class="btn btn--ghost" href="invitation.html">Demander une invitation</a></div>
          </div>
          <div class="maison-hero__visuel" aria-hidden="true"><div class="maison-hero__cadre"><img src="assets/img/identite/monogramme.svg" alt=""><div class="maison-hero__nom">${esc(m.nom)}</div><div class="maison-hero__sous">${v && v.titre ? esc(v.titre) : "Vente privée"}</div>${remiseMax ? `<div class="maison-hero__remise">jusqu'à −${remiseMax} %</div>` : ""}</div></div>
        </div></div></section>
      <section class="section--tight"><div class="container"><div class="stats">
        <div><strong>${produits.length}</strong><span>pièces proposées</span></div><div><strong>${cats.length}</strong><span>univers</span></div>
        <div><strong>${remiseMax ? "−" + remiseMax + " %" : "—"}</strong><span>remise maximale</span></div><div><strong>${v ? (v.etat === "en-cours" ? "J-" + v.joursRestants : "J+" + v.joursAvant) : "—"}</strong><span>${v ? (v.etat === "en-cours" ? "avant clôture" : "avant ouverture") : "vente"}</span></div>
      </div></div></section>
      <section class="section" id="pieces"><div class="container">
        <div class="section__tete"><div><div class="eyebrow">La sélection</div><h2 class="h2">Les pièces ${esc(m.nom)}</h2></div>
          <div class="onglets" id="onglets"><button class="is-active" data-cat="">Tout</button>${cats.map((c) => `<button data-cat="${esc(c)}">${esc(c)}</button>`).join("")}</div></div>
        <div class="grille" id="grille-maison"></div>
      </div></section>
      <section class="section section--blanc"><div class="container"><div class="colonnes-2">
        <div><div class="eyebrow">Conditions de la vente</div><h2 class="h2" style="margin:12px 0 18px">Comment cela fonctionne</h2>
          <ul class="liste-fine"><li>Vente ouverte aux membres ${esc(D.site.nom)}, pour une durée limitée.</li><li>Prix privés affichés TTC, prix boutique conseillé indiqué à titre de comparaison.</li><li>Pièces neuves, livrées dans leur emballage d'origine avec leur facture.</li><li>${esc(D.site.livraison || "")}</li><li>${esc(D.site.retours || "")}</li></ul></div>
        <div><div class="eyebrow">Authenticité</div><h2 class="h2" style="margin:12px 0 18px">Notre engagement</h2>
          <p class="lead" style="font-size:15px">Chaque pièce est contrôlée avant expédition. ${esc(D.site.nom)} est un distributeur indépendant ; le nom de la maison est cité pour identifier les pièces proposées. <a class="link" href="authenticite.html" style="margin-top:12px">Authenticité & garanties ${ICONES.fleche}</a></p></div>
      </div></div></section>`;
    const grille = $("#grille-maison");
    const rendre = (cat) => { const l = produits.filter((p) => !cat || p.categorie === cat); grille.innerHTML = l.length ? l.map(carte).join("") : `<div class="vide" style="grid-column:1/-1"><h3>Aucune pièce dans cet univers</h3></div>`; };
    rendre("");
    $$("#onglets button").forEach((b) => b.addEventListener("click", () => { $$("#onglets button").forEach((x) => x.classList.remove("is-active")); b.classList.add("is-active"); rendre(b.dataset.cat); }));
  };

  Pages.catalogue = function () {
    const etat = { maisons: new Set((params.get("maison") || "").split(",").filter(Boolean)), categories: new Set((params.get("categorie") || "").split(",").filter(Boolean)),
      genres: new Set((params.get("genre") || "").split(",").filter(Boolean)), q: params.get("q") || "", tri: params.get("tri") || "pertinence",
      nouveautes: params.get("nouveautes") === "1", dispo: params.get("dispo") === "1", min: num(params.get("min")), max: num(params.get("max")), vente: params.get("vente") || "", page: 1 };
    const PAR_PAGE = 24;
    const CATS = ["Maroquinerie", "Prêt-à-porter", "Souliers", "Accessoires", "Bagagerie"];
    const titre = $("#titre-catalogue"), sousTitre = $("#sous-titre-catalogue");

    function filtrer() {
      const q = norm(etat.q); const mots = q.split(/\s+/).filter(Boolean);
      let l = D.produits.filter((p) => {
        if (etat.maisons.size && !etat.maisons.has(p.maisonSlug)) return false;
        if (etat.categories.size && !Array.from(etat.categories).some((c) => norm(c) === norm(p.categorie))) return false;
        if (etat.genres.size && !Array.from(etat.genres).some((g) => norm(g) === norm(p.genre) || (norm(p.genre) === "mixte"))) return false;
        if (etat.nouveautes && !p.nouveaute) return false;
        if (etat.dispo && p.epuise) return false;
        if (etat.vente && p.vente !== etat.vente) return false;
        if (etat.min != null && (p.prix == null || p.prix < etat.min)) return false;
        if (etat.max != null && (p.prix == null || p.prix > etat.max)) return false;
        if (mots.length && !mots.every((m) => p.recherche.includes(m))) return false;
        return true;
      });
      const tris = {
        pertinence: (a, b) => (a.epuise - b.epuise) || (b.miseEnAvant - a.miseEnAvant) || (b.nouveaute - a.nouveaute) || a.maison.localeCompare(b.maison, "fr") || a.nom.localeCompare(b.nom, "fr"),
        "prix-asc": (a, b) => (a.prix || 0) - (b.prix || 0), "prix-desc": (a, b) => (b.prix || 0) - (a.prix || 0),
        nouveautes: (a, b) => (b.nouveaute - a.nouveaute) || a.maison.localeCompare(b.maison, "fr"), maison: (a, b) => a.maison.localeCompare(b.maison, "fr") || a.nom.localeCompare(b.nom, "fr"),
        remise: (a, b) => (b.remise || 0) - (a.remise || 0),
      };
      return l.sort(tris[etat.tri] || tris.pertinence);
    }

    function majURL() {
      const u = new URLSearchParams();
      if (etat.maisons.size) u.set("maison", Array.from(etat.maisons).join(",")); if (etat.categories.size) u.set("categorie", Array.from(etat.categories).join(","));
      if (etat.genres.size) u.set("genre", Array.from(etat.genres).join(",")); if (etat.q) u.set("q", etat.q); if (etat.tri !== "pertinence") u.set("tri", etat.tri);
      if (etat.nouveautes) u.set("nouveautes", "1"); if (etat.dispo) u.set("dispo", "1"); if (etat.min != null) u.set("min", etat.min); if (etat.max != null) u.set("max", etat.max); if (etat.vente) u.set("vente", etat.vente);
      history.replaceState(null, "", location.pathname + (u.toString() ? "?" + u.toString() : ""));
    }

    function rendreFiltres() {
      const cherche = norm($("#cherche-maison") ? $("#cherche-maison").value : "");
      const comptes = {}; D.produits.forEach((p) => (comptes[p.maisonSlug] = (comptes[p.maisonSlug] || 0) + 1));
      $("#filtre-maisons").innerHTML = D.maisons.slice().sort((a, b) => a.nom.localeCompare(b.nom, "fr")).filter((m) => !cherche || norm(m.nom).includes(cherche)).map((m) => `<label class="case"><input type="checkbox" data-f="maisons" value="${m.slug}" ${etat.maisons.has(m.slug) ? "checked" : ""}>${esc(m.nom)}<span class="n">${comptes[m.slug] || 0}</span></label>`).join("");
      $("#filtre-categories").innerHTML = CATS.map((c) => `<label class="case"><input type="checkbox" data-f="categories" value="${esc(c)}" ${etat.categories.has(c) ? "checked" : ""}>${c}<span class="n">${D.produits.filter((p) => norm(p.categorie) === norm(c)).length}</span></label>`).join("");
      $("#filtre-genres").innerHTML = ["Femme", "Homme"].map((g) => `<label class="case"><input type="checkbox" data-f="genres" value="${g}" ${etat.genres.has(g) ? "checked" : ""}>${g}</label>`).join("");
      $("#filtre-nouveautes").checked = etat.nouveautes; $("#filtre-dispo").checked = etat.dispo;
      $("#prix-min").value = etat.min != null ? etat.min : ""; $("#prix-max").value = etat.max != null ? etat.max : "";
      $("#tri").value = etat.tri;
    }

    function rendreActifs() {
      const puces = [];
      etat.maisons.forEach((s) => { const m = D.maisonParSlug.get(s); puces.push([`${m ? m.nom : s}`, () => etat.maisons.delete(s)]); });
      etat.categories.forEach((c) => puces.push([c, () => etat.categories.delete(c)])); etat.genres.forEach((g) => puces.push([g, () => etat.genres.delete(g)]));
      if (etat.nouveautes) puces.push(["Nouveautés", () => (etat.nouveautes = false)]); if (etat.dispo) puces.push(["Disponible", () => (etat.dispo = false)]);
      if (etat.min != null || etat.max != null) puces.push([`${etat.min != null ? fmtPrix(etat.min) : "0 €"} – ${etat.max != null ? fmtPrix(etat.max) : "∞"}`, () => { etat.min = null; etat.max = null; }]);
      if (etat.q) puces.push([`« ${etat.q} »`, () => (etat.q = "")]); if (etat.vente) puces.push([`Vente ${etat.vente}`, () => (etat.vente = "")]);
      $("#actifs").innerHTML = puces.map(([t], i) => `<span class="puce">${esc(t)}<button data-i="${i}" aria-label="Retirer le filtre ${esc(t)}">×</button></span>`).join("");
      $$("#actifs button").forEach((b) => b.addEventListener("click", () => { puces[+b.dataset.i][1](); etat.page = 1; tout(); }));
    }

    function rendreGrille() {
      const l = filtrer(); const vis = l.slice(0, etat.page * PAR_PAGE);
      $("#compte").textContent = `${l.length} pièce${l.length > 1 ? "s" : ""}`;
      $("#grille").innerHTML = l.length ? vis.map(carte).join("") : `<div class="vide" style="grid-column:1/-1"><h3>Aucune pièce ne correspond</h3><p>Élargissez vos critères ou <button class="link" id="reset-vide">réinitialisez les filtres</button>.</p></div>`;
      $("#plus").style.display = vis.length < l.length ? "" : "none";
      $("#plus-btn") && ($("#plus-btn").textContent = `Afficher plus (${l.length - vis.length})`);
      $("#reset-vide") && $("#reset-vide").addEventListener("click", reinitialiser);
      // titre contextuel
      let t = "Le catalogue", st = "Toutes les maisons, toutes les pièces.";
      if (etat.maisons.size === 1) { const m = D.maisonParSlug.get(Array.from(etat.maisons)[0]); if (m) { t = m.nom; st = m.presentation || ""; } }
      else if (etat.categories.size === 1) { t = Array.from(etat.categories)[0]; st = "La sélection " + t.toLowerCase() + " de toutes les maisons."; }
      else if (etat.genres.size === 1) { t = Array.from(etat.genres)[0]; st = "Le vestiaire " + (Array.from(etat.genres)[0] === "Femme" ? "féminin" : "masculin") + ", toutes maisons confondues."; }
      else if (etat.nouveautes) { t = "Nouveautés"; st = "Les dernières pièces entrées au catalogue."; }
      else if (etat.q) { t = `Résultats pour « ${etat.q} »`; st = ""; }
      titre.textContent = t; sousTitre.textContent = st; document.title = t + " · " + D.site.nom;
    }
    function tout() { majURL(); rendreFiltres(); rendreActifs(); rendreGrille(); }
    function reinitialiser() { etat.maisons.clear(); etat.categories.clear(); etat.genres.clear(); etat.q = ""; etat.nouveautes = false; etat.dispo = false; etat.min = null; etat.max = null; etat.vente = ""; etat.page = 1; tout(); }

    $("#filtres").addEventListener("change", (e) => {
      const i = e.target; if (i.dataset.f) { const set = etat[i.dataset.f]; i.checked ? set.add(i.value) : set.delete(i.value); }
      if (i.id === "filtre-nouveautes") etat.nouveautes = i.checked; if (i.id === "filtre-dispo") etat.dispo = i.checked;
      if (i.id === "prix-min") etat.min = num(i.value); if (i.id === "prix-max") etat.max = num(i.value);
      etat.page = 1; if (i.id !== "cherche-maison") tout();
    });
    $("#cherche-maison").addEventListener("input", rendreFiltres);
    $("#tri").addEventListener("change", (e) => { etat.tri = e.target.value; tout(); });
    $("#plus-btn").addEventListener("click", () => { etat.page++; rendreGrille(); });
    $("#reset").addEventListener("click", reinitialiser);
    const ouvrir = () => $("#filtres").classList.add("is-open"), fermer = () => $("#filtres").classList.remove("is-open");
    $("#ouvrir-filtres").addEventListener("click", ouvrir); $("#fermer-filtres").addEventListener("click", fermer); $("#appliquer-filtres").addEventListener("click", fermer);
    tout();
  };

  Pages.produit = function () {
    const p = D.parRef.get(params.get("ref") || "");
    const racine = $("#fiche");
    if (!p) { racine.innerHTML = `<div class="container vide"><h3>Pièce introuvable</h3><p>Cette référence n'est plus disponible.</p><p><a class="link" href="catalogue.html">Retour au catalogue</a></p></div>`; return; }
    document.title = `${p.nom} — ${p.maison} · ${D.site.nom}`;
    const m = D.maisonParSlug.get(p.maisonSlug); const s = D.site;
    const dispo = p.epuise ? '<span class="dispo dispo--epuise">Épuisé</span>' : p.stock == null ? '<span class="dispo dispo--demande">Disponible sur demande</span>' : p.stock <= 2 ? `<span class="dispo">Dernières pièces (${p.stock})</span>` : '<span class="dispo">Disponible</span>';
    const sujet = encodeURIComponent(`Demande d'information — ${p.ref} ${p.maison} ${p.nom}`);
    const corps = encodeURIComponent(`Bonjour,\n\nJe souhaite obtenir des informations sur la pièce suivante :\n${p.maison} — ${p.nom} (réf. ${p.ref}) — ${fmtPrix(p.prix)}\n\nMerci.`);
    const wa = s.whatsapp ? `https://wa.me/${s.whatsapp.replace(/\D/g, "")}?text=${encodeURIComponent(`Bonjour, je souhaite des informations sur ${p.maison} — ${p.nom} (réf. ${p.ref}).`)}` : "";
    racine.innerHTML = `<div class="container">
      <nav class="fil" aria-label="Fil d'Ariane" style="padding-top:24px"><a href="index.html">Accueil</a><span><a href="catalogue.html">Catalogue</a></span><span><a href="maison.html?m=${p.maisonSlug}">${esc(p.maison)}</a></span><span>${esc(p.nom)}</span></nav>
      <div class="produit">
        <div class="galerie ${p.images.length < 2 ? "galerie--seule" : ""}">
          <div class="galerie__vignettes">${p.images.map((src, i) => `<button class="${i === 0 ? "is-active" : ""}" data-i="${i}" aria-label="Visuel ${i + 1}"><img src="${esc(src)}" alt="" loading="lazy"></button>`).join("")}</div>
          <div class="galerie__principale"><img id="visuel-principal" src="${esc(p.images[0])}" alt="${esc(p.nom)} — ${esc(p.maison)}" class="${p.placeholder ? "is-placeholder" : ""}" onerror="this.onerror=null;this.src='assets/img/placeholders/generique.svg';this.classList.add('is-placeholder')"></div>
        </div>
        <div class="infos">
          <a class="maison-lien" href="maison.html?m=${p.maisonSlug}">${esc(p.maison)}</a>
          <h1>${esc(p.nom)}</h1>
          <div class="prix"><span>${fmtPrix(p.prix)}</span>${p.ancienPrix && p.ancienPrix > (p.prix || 0) ? `<s>${fmtPrix(p.ancienPrix)}</s><span class="etiquette etiquette--or">−${p.remise} %</span>` : ""}</div>
          <div class="taxes">Prix TTC${p.ancienPrix && p.ancienPrix > (p.prix || 0) ? " · prix boutique conseillé barré" : ""}${p.venteObj && p.venteObj.etat === "en-cours" ? ` · vente jusqu'au ${fmtDate(p.venteObj.fin)}` : ""}</div>
          <div class="ref">Réf. ${esc(p.ref)}</div>
          <dl class="caracs">${[["Couleur", p.couleur], ["Matière", p.matiere], ["Dimensions", p.dimensions], ["Univers", [p.categorie, p.type].filter(Boolean).join(" · ")], ["Genre", p.genre]].filter(([, v]) => v).map(([k, v]) => `<div class="carac"><dt>${k}</dt><dd>${esc(v)}</dd></div>`).join("")}</dl>
          ${p.tailles.length ? `<div class="tailles"><div class="tailles__titre"><span>Taille</span><span class="muted" id="taille-choisie">Sélectionnez une taille</span></div><div class="tailles__liste" id="tailles">${p.tailles.map((t) => `<button type="button" data-t="${esc(t)}">${esc(t)}</button>`).join("")}</div></div>` : ""}
          ${dispo}
          <div class="achat">
            <div class="ligne"><div class="qte"><button type="button" id="qte-moins" aria-label="Moins">−</button><input id="qte" type="number" min="1" max="9" value="1" aria-label="Quantité"><button type="button" id="qte-plus" aria-label="Plus">+</button></div>
              <button class="btn btn--block" id="ajouter" ${p.epuise ? "disabled" : ""}>${p.epuise ? "Épuisé" : "Ajouter au panier"}</button></div>
            ${p.lienPaiement && !p.epuise ? `<a class="btn btn--ghost btn--block" href="${esc(p.lienPaiement)}" target="_blank" rel="noopener">Acheter maintenant — paiement sécurisé</a>` : ""}
          </div>
          <div class="contact-rapide"><a class="link" href="mailto:${esc(s.email)}?subject=${sujet}&body=${corps}">Poser une question</a>${wa ? `<a class="link" href="${wa}" target="_blank" rel="noopener">WhatsApp</a>` : ""}</div>
          <div class="accordeon">
            <details open><summary>Description</summary><div class="corps">${esc(p.description || "Description à venir.")}</div></details>
            <details><summary>Authenticité</summary><div class="corps">Pièce neuve, contrôlée avant expédition et livrée dans son emballage d'origine avec sa facture. ${esc(s.nom)} est un distributeur indépendant ; le nom de la maison est cité pour identifier la pièce. <a href="authenticite.html">En savoir plus</a>.</div></details>
            <details><summary>Livraison & retours</summary><div class="corps"><p>${esc(s.livraison || "")}</p><p>${esc(s.retours || "")}</p></div></details>
            ${m && m.presentation ? `<details><summary>La maison ${esc(m.nom)}</summary><div class="corps">${esc(m.presentation)} <a href="maison.html?m=${m.slug}">Voir l'espace ${esc(m.nom)}</a>.</div></details>` : ""}
          </div>
          <div class="reassurance"><div>${ICONES.auth}Authenticité garantie</div><div>${ICONES.livraison}Livraison assurée</div><div>${ICONES.retour}Retour sous 14 jours</div></div>
        </div>
      </div>
      <section class="section section--tight"><div class="section__tete"><div><div class="eyebrow">Dans le même esprit</div><h2 class="h2">Vous aimerez aussi</h2></div><a class="link" href="maison.html?m=${p.maisonSlug}">Tout ${esc(p.maison)} ${ICONES.fleche}</a></div><div class="grille" id="similaires"></div></section>
    </div>`;
    let taille = ""; let qte = 1;
    $$("#tailles button").forEach((b) => b.addEventListener("click", () => { $$("#tailles button").forEach((x) => x.classList.remove("is-active")); b.classList.add("is-active"); taille = b.dataset.t; $("#taille-choisie").textContent = "Taille " + taille; }));
    $$(".galerie__vignettes button").forEach((b) => b.addEventListener("click", () => { $$(".galerie__vignettes button").forEach((x) => x.classList.remove("is-active")); b.classList.add("is-active"); $("#visuel-principal").src = p.images[+b.dataset.i]; }));
    const inQ = $("#qte"); $("#qte-moins").addEventListener("click", () => (inQ.value = Math.max(1, +inQ.value - 1))); $("#qte-plus").addEventListener("click", () => (inQ.value = Math.min(9, +inQ.value + 1)));
    $("#ajouter").addEventListener("click", () => {
      if (p.tailles.length && !taille) { toast("Veuillez choisir une taille"); $("#tailles").scrollIntoView({ behavior: "smooth", block: "center" }); return; }
      qte = Math.max(1, Math.min(9, +inQ.value || 1)); Panier.ajouter(p.ref, taille, qte); toast("Ajouté au panier");
    });
    const sim = D.produits.filter((x) => x.ref !== p.ref && !x.epuise && (x.maisonSlug === p.maisonSlug || norm(x.type) === norm(p.type))).sort((a, b) => (b.maisonSlug === p.maisonSlug) - (a.maisonSlug === p.maisonSlug)).slice(0, 4);
    $("#similaires").innerHTML = sim.map(carte).join("");
  };

  Pages.panier = function () {
    const racine = $("#contenu-panier"); const s = D.site;
    function rendre() {
      const lignes = Panier.lire().map((l) => ({ ...l, p: D.parRef.get(l.ref) })).filter((l) => l.p);
      if (!lignes.length) { racine.innerHTML = `<div class="panier-vide"><h2>Votre panier est vide</h2><p class="lead" style="margin:0 auto 24px">Découvrez les ventes en cours et composez votre sélection.</p><a class="btn" href="ventes.html">Voir les ventes</a></div>`; return; }
      const total = lignes.reduce((t, l) => t + (l.p.prix || 0) * l.qte, 0); const totalBoutique = lignes.reduce((t, l) => t + (l.p.ancienPrix || l.p.prix || 0) * l.qte, 0);
      racine.innerHTML = `<div class="panier">
        <div>${lignes.map((l, i) => `<div class="ligne-panier"><a href="produit.html?ref=${encodeURIComponent(l.ref)}"><img src="${esc(l.p.images[0])}" alt="" class="${l.p.placeholder ? "is-placeholder" : ""}"></a>
          <div><div class="maison">${esc(l.p.maison)}</div><a class="nom" href="produit.html?ref=${encodeURIComponent(l.ref)}">${esc(l.p.nom)}</a><div class="detail">Réf. ${esc(l.ref)}${l.taille ? " · Taille " + esc(l.taille) : ""}${l.p.couleur ? " · " + esc(l.p.couleur) : ""}</div>
            <div class="bas"><div class="qte"><button data-i="${i}" data-d="-1" aria-label="Moins">−</button><input type="number" value="${l.qte}" min="1" max="9" data-i="${i}" aria-label="Quantité"><button data-i="${i}" data-d="1" aria-label="Plus">+</button></div><button class="retirer" data-r="${i}">Retirer</button></div></div>
          <div class="total-ligne">${fmtPrix((l.p.prix || 0) * l.qte)}${l.p.ancienPrix && l.p.ancienPrix > (l.p.prix || 0) ? `<div class="small muted"><s>${fmtPrix(l.p.ancienPrix * l.qte)}</s></div>` : ""}</div></div>`).join("")}
          <p class="small muted" style="margin-top:18px"><button class="link" id="vider" style="border-bottom-color:transparent">Vider le panier</button></p></div>
        <aside class="recap"><h2>Récapitulatif</h2>
          <div class="ligne"><span>${lignes.reduce((n, l) => n + l.qte, 0)} pièce(s)</span><span>${fmtPrix(total)}</span></div>
          ${totalBoutique > total ? `<div class="ligne"><span>Avantage membre</span><span style="color:var(--or)">− ${fmtPrix(totalBoutique - total)}</span></div>` : ""}
          <div class="ligne"><span>Livraison</span><span>Offerte</span></div>
          <div class="ligne total"><span>Total TTC</span><span>${fmtPrix(total)}</span></div>
          <p class="note">Votre demande est transmise à notre équipe, qui confirme la disponibilité de chaque pièce et vous adresse un lien de paiement sécurisé sous 24 h ouvrées.</p>
          <form id="form-commande">
            <div class="champ"><label for="c-nom">Nom et prénom</label><input id="c-nom" required autocomplete="name"></div>
            <div class="champ"><label for="c-email">E-mail</label><input id="c-email" type="email" required autocomplete="email"></div>
            <div class="champ"><label for="c-tel">Téléphone</label><input id="c-tel" type="tel" autocomplete="tel"></div>
            <div class="champ"><label for="c-msg">Message (facultatif)</label><textarea id="c-msg" style="min-height:80px"></textarea></div>
            <button class="btn btn--block" type="submit">Envoyer ma demande de commande</button>
          </form>
          ${s.whatsapp ? `<div class="ou">ou</div><a class="btn btn--ghost btn--block" id="wa" href="#" target="_blank" rel="noopener">Commander via WhatsApp</a>` : ""}
        </aside></div>`;
      const recap = () => lignes.map((l) => `• ${l.p.maison} — ${l.p.nom} (réf. ${l.ref}${l.taille ? ", taille " + l.taille : ""}) × ${l.qte} — ${fmtPrix((l.p.prix || 0) * l.qte)}`).join("\n") + `\n\nTotal TTC : ${fmtPrix(total)}`;
      $$(".ligne-panier .qte button").forEach((b) => b.addEventListener("click", () => { const l = lignes[+b.dataset.i]; Panier.quantite(l.ref, l.taille, l.qte + +b.dataset.d); rendre(); }));
      $$(".ligne-panier .qte input").forEach((i) => i.addEventListener("change", () => { const l = lignes[+i.dataset.i]; Panier.quantite(l.ref, l.taille, +i.value || 1); rendre(); }));
      $$(".retirer").forEach((b) => b.addEventListener("click", () => { const l = lignes[+b.dataset.r]; Panier.retirer(l.ref, l.taille); rendre(); }));
      $("#vider").addEventListener("click", () => { Panier.vider(); rendre(); });
      $("#form-commande").addEventListener("submit", (e) => {
        e.preventDefault();
        const nom = $("#c-nom").value.trim(), email = $("#c-email").value.trim(), tel = $("#c-tel").value.trim(), msg = $("#c-msg").value.trim();
        const corps = `Bonjour,\n\nJe souhaite commander les pièces suivantes :\n\n${recap()}\n\nNom : ${nom}\nE-mail : ${email}\nTéléphone : ${tel}\n${msg ? "\nMessage : " + msg + "\n" : ""}\nMerci de me confirmer la disponibilité et de m'adresser le lien de paiement.`;
        location.href = `mailto:${s.email}?subject=${encodeURIComponent("Demande de commande — " + nom)}&body=${encodeURIComponent(corps)}`;
        toast("Votre messagerie s'ouvre avec la demande pré-remplie");
      });
      const wa = $("#wa"); if (wa) wa.href = `https://wa.me/${s.whatsapp.replace(/\D/g, "")}?text=${encodeURIComponent("Bonjour, je souhaite commander :\n\n" + recap())}`;
    }
    rendre();
  };

  Pages.contact = function () {
    const s = D.site; const f = $("#form-contact"); if (!f) return;
    if (s.formulaire_action) { f.action = s.formulaire_action; f.method = "POST"; return; }
    f.addEventListener("submit", (e) => { e.preventDefault(); const d = new FormData(f); const corps = `Bonjour,\n\n${d.get("message")}\n\n${d.get("nom")}\n${d.get("email")}${d.get("telephone") ? "\n" + d.get("telephone") : ""}`;
      location.href = `mailto:${s.email}?subject=${encodeURIComponent(d.get("sujet") || "Contact")}&body=${encodeURIComponent(corps)}`; toast("Votre messagerie s'ouvre avec le message pré-rempli"); });
  };

  Pages.invitation = function () {
    const s = D.site; const f = $("#form-invitation"); if (!f) return;
    if (s.formulaire_action) { f.action = s.formulaire_action; f.method = "POST"; return; }
    f.addEventListener("submit", (e) => { e.preventDefault(); const d = new FormData(f); const corps = `Bonjour,\n\nJe souhaite rejoindre ${s.nom}.\n\nNom : ${d.get("nom")}\nE-mail : ${d.get("email")}\nTéléphone : ${d.get("telephone") || "—"}\nMaisons qui m'intéressent : ${d.get("maisons") || "—"}\nParrain / code : ${d.get("parrain") || "—"}`;
      location.href = `mailto:${s.email}?subject=${encodeURIComponent("Demande d'invitation — " + d.get("nom"))}&body=${encodeURIComponent(corps)}`; toast("Votre messagerie s'ouvre avec la demande pré-remplie"); });
  };

  /* ------------------------------------------------------------ démarrage */
  document.addEventListener("DOMContentLoaded", async () => {
    try {
      await charger();
    } catch (e) {
      console.error(e);
      const z = $("#entete"); if (z) z.insertAdjacentHTML("afterend", `<div class="container erreur-donnees">Les données du site n'ont pas pu être chargées (${esc(e.message)}). Vérifiez que le dossier <code>data/</code> est bien en ligne.</div>`);
      return;
    }
    rendreGabarit();
    const page = document.body.dataset.page; if (page && Pages[page]) { try { Pages[page](); } catch (e) { console.error(e); } }
    $$(".chargement").forEach((c) => c.remove());
  });

  window.VPLuxe = { data: D, panier: Panier, charger };
})();
