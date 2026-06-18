const { chromium } = require("playwright");
const ExcelJS = require("exceljs");
const fs = require("fs");

// Format urls.txt: "Nom Client|https://url.com" ou simplement "https://url.com"
const urlEntries = fs.readFileSync("urls.txt", "utf8")
  .split(/\r?\n/)
  .map(u => u.trim())
  .filter(Boolean)
  .map(line => {
    const parts = line.split("|");
    if (parts.length >= 2) {
      return { clientName: parts[0].trim(), url: parts[1].trim() };
    }
    return { clientName: null, url: line };
  });

// --- TEST CLAVIER : simulation réelle via Playwright ---
async function testKeyboardNavigation(page) {
  try {
    await page.click("body", { timeout: 2000 });
  } catch (e) { /* ignore */ }

  const seenIds = new Set();
  const invisibleFocus = [];
  let reachableCount = 0;

  for (let i = 0; i < 40; i++) {
    await page.keyboard.press("Tab");

    const focused = await page.evaluate(() => {
      const el = document.activeElement;
      if (!el || el.tagName === "BODY" || el === document.documentElement) return null;

      const style = getComputedStyle(el);
      const outlineWidth = parseFloat(style.outlineWidth) || 0;
      const outlineStyle = style.outlineStyle;
      const boxShadow = style.boxShadow;

      const hasFocusIndicator =
        (outlineWidth > 0 && outlineStyle !== "none") ||
        (boxShadow !== "none" && boxShadow !== "");

      const uid = `${el.tagName}|${el.id}|${(el.textContent || "").trim().slice(0, 30)}`;
      return { tag: el.tagName, hasFocusIndicator, uid };
    });

    if (!focused) break;
    if (seenIds.has(focused.uid)) break;
    seenIds.add(focused.uid);
    reachableCount++;
    if (!focused.hasFocusIndicator) invisibleFocus.push(focused.tag);
  }

  return {
    reachableCount,
    invisibleFocusCount: invisibleFocus.length,
    reachableStatus: reachableCount === 0 ? "KO" : "À vérifier",
    reachableComment: reachableCount === 0
      ? "Aucun élément n'a reçu le focus lors de la navigation Tab. Navigation clavier compromise."
      : `${reachableCount} élément(s) atteints par Tab. Vérification manuelle complète requise.`,
    usableStatus: invisibleFocus.length > 0 ? "KO" : reachableCount === 0 ? "KO" : "À vérifier",
    usableComment: invisibleFocus.length > 0
      ? `${invisibleFocus.length}/${reachableCount} élément(s) sans indicateur de focus visible (${[...new Set(invisibleFocus)].slice(0, 5).join(", ")}). Vérifier Entrée/Espace/flèches manuellement.`
      : reachableCount === 0
        ? "Aucun élément focusable atteint au clavier."
        : `Focus visible détecté sur les ${reachableCount} éléments testés. Vérifier l'utilisation Entrée/Espace/flèches manuellement.`
  };
}

// --- TEST FORMULAIRE : soumission réelle via Playwright ---
async function testFormSubmission(page) {
  const formInfo = await page.evaluate(() => {
    const forms = [...document.querySelectorAll("form")].filter(f => {
      const rect = f.getBoundingClientRect();
      return rect.width > 0 && rect.height > 0;
    });
    if (forms.length === 0) return { hasForms: false };
    const form = forms[0];
    const hasSubmit = !!form.querySelector('[type="submit"], button:not([type="button"]):not([type="reset"]), button[type="submit"]');
    const requiredCount = form.querySelectorAll("[required], [aria-required='true']").length;
    return { hasForms: true, hasSubmit, requiredCount, formCount: forms.length };
  });

  if (!formInfo.hasForms) return { status: "NA", comment: "Aucun formulaire détecté." };
  if (!formInfo.hasSubmit) return { status: "À vérifier", comment: `${formInfo.formCount} formulaire(s) détecté(s) — aucun bouton de soumission identifié.` };

  try {
    await page.evaluate(() => {
      const form = document.querySelector("form");
      const inputs = form.querySelectorAll('input:not([type="hidden"]):not([type="submit"]):not([type="button"]):not([type="reset"]), textarea');
      inputs.forEach(input => { input.value = ""; });
      const selects = form.querySelectorAll("select");
      selects.forEach(select => { select.selectedIndex = 0; });
    });

    const submitBtn = await page.$('form [type="submit"], form button[type="submit"], form button:not([type="button"]):not([type="reset"])');
    if (submitBtn) {
      await Promise.race([
        submitBtn.click(),
        new Promise(resolve => setTimeout(resolve, 3000))
      ]);
      await page.waitForTimeout(1500);
    }

    const errorCount = await page.evaluate(() => {
      const selectors = [
        '[role="alert"]', '[aria-live="polite"]', '[aria-live="assertive"]',
        '[aria-invalid="true"]', '[class*="error" i]', '[class*="erreur" i]',
        '[class*="invalid" i]', '[class*="danger" i]',
        '.field-error', '.form-error', '.validation-error', '[id*="error" i]'
      ];
      return document.querySelectorAll(selectors.join(",")).length;
    });

    return {
      status: errorCount > 0 ? "OK" : "KO",
      comment: errorCount > 0
        ? `${errorCount} message(s) d'erreur détecté(s) après soumission du formulaire vide.`
        : `Aucun message d'erreur détecté après soumission d'un formulaire vide (${formInfo.requiredCount} champ(s) obligatoire(s)).`
    };
  } catch (e) {
    return {
      status: "À vérifier",
      comment: `${formInfo.formCount} formulaire(s) détecté(s). Test de soumission échoué : ${e.message.slice(0, 80)}.`
    };
  }
}

async function auditSite(page, url) {
  await page.goto(url, { waitUntil: "domcontentloaded", timeout: 45000 });
  await page.waitForTimeout(2500);

  const result = await page.evaluate(() => {
    function isVisible(el) {
      const style = getComputedStyle(el);
      const rect = el.getBoundingClientRect();
      return style.display !== "none" &&
        style.visibility !== "hidden" &&
        rect.width > 0 &&
        rect.height > 0;
    }

    function addRow(theme, ref, test, expected, status, comment) {
      result.rows.push({ theme, ref, test, expected, status, comment, url: location.href });
    }

    const result = { rows: [] };
    const pageLang = document.documentElement.getAttribute("lang") || "";

    // --- CONTENU TEXTUEL ---

    addRow(
      "CONTENU TEXTUEL",
      "Donner un titre aux pages",
      "Lancer l'inspecteur de code du navigateur et examiner le titre de page (<title>[titre de la page]</title>).",
      "Chaque page possède un titre unique et descriptif du contenu, globalement du plus précis vers le plus général (exemple : [résumé du contenu de la page - nom du site]).",
      document.title && document.title.trim().length > 5 ? "OK" : "KO",
      document.title ? `Titre détecté : ${document.title}` : "Aucun titre de page détecté."
    );

    const headings = [...document.querySelectorAll("h1,h2,h3,h4,h5,h6")].filter(isVisible);

    addRow(
      "CONTENU TEXTUEL",
      "Donner un titre aux rubriques",
      "Installer le bookmarklet Headings en le glissant dans la barre des favoris de votre navigateur puis l'exécuter.",
      "Tous les contenus traités visuellement comme des titres possèdent une sémantique de titre (balises <h1> à <h6>).",
      headings.length > 0 ? "OK" : "KO",
      headings.length > 0
        ? `${headings.length} titre(s) détecté(s) : ${headings.slice(0, 8).map(h => `${h.tagName} ${(h.innerText || "").trim()}`).join(" | ")}`
        : "Aucun titre Hn détecté."
    );

    const levels = headings.map(h => Number(h.tagName.substring(1)));
    let hierarchyIssue = false;
    for (let i = 1; i < levels.length; i++) {
      if (levels[i] - levels[i - 1] > 1) hierarchyIssue = true;
    }

    addRow(
      "CONTENU TEXTUEL",
      "Donner un titre aux rubriques",
      "Installer le bookmarklet Headings en le glissant dans la barre des favoris de votre navigateur puis l'exécuter.",
      "Les titres de niveaux sont hiérarchisés de manière à refléter leur poids sémantique.",
      headings.length === 0 || hierarchyIssue ? "KO" : "OK",
      headings.length === 0
        ? "Impossible de vérifier : aucun Hn détecté."
        : hierarchyIssue
          ? `Saut de niveau détecté : ${levels.join(" > ")}`
          : `Hiérarchie détectée : ${levels.join(" > ")}`
    );

    addRow(
      "CONTENU TEXTUEL",
      "Indiquer la langue principale",
      "Lancer l'inspecteur de code du navigateur. Examiner l'élément <html>.",
      "Un attribut lang est présent dans l'élément <html> de la page.",
      pageLang ? "OK" : "KO",
      pageLang ? `Lang détecté : ${pageLang}` : "Attribut lang absent."
    );

    addRow(
      "CONTENU TEXTUEL",
      "Indiquer la langue principale",
      "Lancer l'inspecteur de code du navigateur. Examiner l'élément <html>.",
      "La valeur de l'attribut lang correspond à la langue principale du document, exemple : <html lang='fr'>, <html lang='en-US'>.",
      pageLang ? "À vérifier" : "KO",
      pageLang ? `Valeur détectée : ${pageLang}. À confirmer manuellement.` : "Langue principale non déclarée."
    );

    // --- CONTENU NON TEXTUEL ---

    const images = [...document.images].filter(isVisible);
    const imageLinks = images.filter(img => img.closest("a"));
    const missingAlt = images.filter(img => !img.hasAttribute("alt") && img.getAttribute("role") !== "presentation");
    const emptyAlt = images.filter(img => img.hasAttribute("alt") && img.getAttribute("alt") === "");

    addRow(
      "CONTENU NON TEXTUEL",
      "S'assurer que les images ont une alternative textuelle",
      "Installer puis lancer le bookmarklet List Images ou l'inspecteur de code.",
      "Image lien : le contenu de l'attribut alt de chaque image-lien est pertinent par rapport à la cible du lien.",
      imageLinks.length === 0 ? "NA" : imageLinks.every(img => img.hasAttribute("alt")) ? "OK" : "KO",
      imageLinks.length === 0
        ? "Aucune image-lien détectée."
        : `${imageLinks.length} image(s)-lien détectée(s), ${imageLinks.filter(img => !img.hasAttribute("alt")).length} sans alt.`
    );

    addRow(
      "CONTENU NON TEXTUEL",
      "S'assurer que les images ont une alternative textuelle",
      "Installer puis lancer le bookmarklet List Images ou l'inspecteur de code.",
      "Image porteuse d'information : l'attribut alt de chaque image est pertinent par rapport au rôle de l'image dans la page.",
      missingAlt.length === 0 ? "OK" : "KO",
      missingAlt.length === 0
        ? `${images.length} image(s) analysée(s), aucune image sans alt détectée.`
        : `${missingAlt.length} image(s) sans attribut alt.`
    );

    const imagesWithNonEmptyAlt = images.filter(img => {
      const alt = img.getAttribute("alt");
      return img.hasAttribute("alt") && alt !== "" && alt !== null && !img.closest("a");
    });

    addRow(
      "CONTENU NON TEXTUEL",
      "S'assurer que les images ont une alternative textuelle",
      "Installer puis lancer le bookmarklet List Images ou l'inspecteur de code.",
      "Image contenant du texte : l'attribut alt reprend au moins le texte de l'image.",
      imagesWithNonEmptyAlt.length === 0 ? "NA" : "À vérifier",
      imagesWithNonEmptyAlt.length === 0
        ? "Aucune image informative (hors lien) avec alt non vide détectée."
        : `${imagesWithNonEmptyAlt.length} image(s) avec alt non vide. Vérifier que l'alt reprend bien le texte visible dans l'image.`
    );

    const captchaImages = images.filter(img => {
      const alt = (img.getAttribute("alt") || "").toLowerCase();
      const src = (img.getAttribute("src") || "").toLowerCase();
      const cls = (img.getAttribute("class") || "").toLowerCase();
      const id = (img.getAttribute("id") || "").toLowerCase();
      return alt.includes("captcha") || src.includes("captcha") || cls.includes("captcha") || id.includes("captcha");
    });

    addRow(
      "CONTENU NON TEXTUEL",
      "S'assurer que les images ont une alternative textuelle",
      "",
      "Pour chaque image utilisée comme CAPTCHA ou comme image-test, ayant une alternative textuelle, cette alternative permet-elle d'identifier la nature et la fonction de l'image ?",
      captchaImages.length === 0 ? "NA" : "À vérifier",
      captchaImages.length === 0
        ? "Aucune image CAPTCHA détectée."
        : `${captchaImages.length} image(s) CAPTCHA détectée(s). Vérifier que l'alternative textuelle identifie la nature et la fonction.`
    );

    addRow(
      "CONTENU NON TEXTUEL",
      "S'assurer que les images ont une alternative textuelle",
      "",
      "Pour chaque image utilisée comme CAPTCHA, une solution d'accès alternatif au contenu ou à la fonction du CAPTCHA est-elle présente ?",
      captchaImages.length === 0 ? "NA" : "À vérifier",
      captchaImages.length === 0
        ? "Aucune image CAPTCHA détectée."
        : `${captchaImages.length} image(s) CAPTCHA détectée(s). Vérifier si une alternative d'accès est proposée (audio CAPTCHA, autre mécanisme).`
    );

    addRow(
      "CONTENU NON TEXTUEL",
      "S'assurer que les images ont une alternative textuelle",
      "Installer puis lancer le bookmarklet List Images ou l'inspecteur de code.",
      "Image décorative : l'attribut alt est présent mais vide.",
      emptyAlt.length > 0 ? "OK" : "NA",
      emptyAlt.length > 0
        ? `${emptyAlt.length} image(s) avec alt vide détectée(s).`
        : "Aucune image décorative avec alt vide détectée."
    );

    const imagesWithLongDesc = images.filter(img => img.hasAttribute("longdesc") || img.hasAttribute("aria-describedby"));
    const largeImages = images.filter(img => {
      const rect = img.getBoundingClientRect();
      return rect.width > 400 && !img.closest("a") && img.getAttribute("alt") !== "";
    });
    const figuresWithCaption = [...document.querySelectorAll("figure")].filter(fig => fig.querySelector("figcaption") && fig.querySelector("img"));

    addRow(
      "CONTENU NON TEXTUEL",
      "S'assurer que les images ont une alternative textuelle",
      "Installer puis lancer le bookmarklet List Images ou l'inspecteur de code.",
      "Image complexe dont le contenu du alt serait trop long (schémas, graphes...) : pour toute description d'image trop longue pour être mise dans un attribut alt, la description longue sous forme de texte est présente dans la page, soit consultable par lien à proximité de l'image à décrire et pointant vers une page html contenant la description.",
      images.length === 0 ? "NA" : imagesWithLongDesc.length > 0 ? "À vérifier" : largeImages.length > 0 ? "À vérifier" : "NA",
      images.length === 0
        ? "Aucune image détectée."
        : imagesWithLongDesc.length > 0
          ? `${imagesWithLongDesc.length} image(s) avec longdesc ou aria-describedby. Vérifier la pertinence de la description longue.`
          : largeImages.length > 0
            ? `${largeImages.length} grande(s) image(s) (>400px) potentiellement complexe(s). Vérifier si une description longue est nécessaire.${figuresWithCaption.length > 0 ? ` ${figuresWithCaption.length} figure(s) avec figcaption détectée(s).` : ""}`
            : "Aucune image complexe détectée automatiquement (vérification manuelle conseillée)."
    );

    const textImages = images.filter(img => {
      const alt = img.getAttribute("alt") || "";
      const src = (img.getAttribute("src") || "").toLowerCase();
      return alt !== "" && !img.closest("a") && !img.closest("figure") &&
        !src.includes("logo") && !src.includes("icon") && !src.includes("avatar") && !src.includes("sprite");
    });

    addRow(
      "CONTENU NON TEXTUEL",
      "S'assurer que les images ont une alternative textuelle",
      "",
      "Chaque image texte porteuse d'information, en l'absence d'un mécanisme de remplacement, doit si possible être remplacée par du texte stylé. Cette règle est-elle respectée (hors cas particuliers) ?",
      textImages.length === 0 ? "NA" : "À vérifier",
      textImages.length === 0
        ? "Aucune image texte porteuse d'information détectée."
        : `${textImages.length} image(s) potentiellement porteuse(s) de texte détectée(s). Vérifier si elles peuvent être remplacées par du texte stylé CSS.`
    );

    const figuresWithImg = [...document.querySelectorAll("figure")].filter(fig => fig.querySelector("img"));
    const figuresWithoutCaption = figuresWithImg.filter(fig => !fig.querySelector("figcaption"));
    const imagesWithTitle = images.filter(img => img.hasAttribute("title") && !img.closest("figure"));

    addRow(
      "CONTENU NON TEXTUEL",
      "S'assurer que les images ont une alternative textuelle",
      "",
      "Chaque légende d'image est-elle, si nécessaire, correctement reliée à l'image correspondante ?",
      figuresWithImg.length === 0 && imagesWithTitle.length === 0 ? "NA"
        : figuresWithoutCaption.length === 0 ? "OK" : "À vérifier",
      figuresWithImg.length === 0 && imagesWithTitle.length === 0
        ? "Aucune légende d'image (figure/figcaption ou title) détectée."
        : figuresWithoutCaption.length > 0
          ? `${figuresWithoutCaption.length} élément(s) <figure> sans <figcaption> détecté(s).`
          : `${figuresWithImg.length} figure(s) avec figcaption correctement liées.${imagesWithTitle.length > 0 ? ` ${imagesWithTitle.length} image(s) avec attribut title.` : ""}`
    );

    // --- COULEURS ET CONTRASTE ---

    function srgbToLin(c) {
      c = c / 255;
      return c <= 0.03928 ? c / 12.92 : Math.pow((c + 0.055) / 1.055, 2.4);
    }
    function luminance(rgb) {
      return 0.2126 * srgbToLin(rgb[0]) + 0.7152 * srgbToLin(rgb[1]) + 0.0722 * srgbToLin(rgb[2]);
    }
    function contrastRatio(fg, bg) {
      const L1 = luminance(fg), L2 = luminance(bg);
      return (Math.max(L1, L2) + 0.05) / (Math.min(L1, L2) + 0.05);
    }
    function parseRgb(value) {
      const match = String(value || "").match(/rgba?\((\d+),\s*(\d+),\s*(\d+)/);
      return match ? [Number(match[1]), Number(match[2]), Number(match[3])] : null;
    }
    function getEffectiveBackground(el) {
      let current = el;
      while (current && current !== document.documentElement) {
        const bg = getComputedStyle(current).backgroundColor;
        if (bg && bg !== "transparent" && !bg.includes("rgba(0, 0, 0, 0)")) return bg;
        current = current.parentElement;
      }
      return "rgb(255, 255, 255)";
    }

    const textIssues = [];
    const elements = [...document.querySelectorAll("body *")].filter(isVisible).slice(0, 900);
    for (const el of elements) {
      const text = (el.innerText || "").trim().replace(/\s+/g, " ");
      if (text.length < 3 || el.children.length > 4) continue;
      const style = getComputedStyle(el);
      const fg = parseRgb(style.color);
      const bg = parseRgb(getEffectiveBackground(el));
      if (!fg || !bg) continue;
      const ratio = contrastRatio(fg, bg);
      const fontSize = parseFloat(style.fontSize);
      const fontWeight = parseInt(style.fontWeight) || 400;
      const isLargeText = fontSize >= 24 || (fontSize >= 18.5 && fontWeight >= 700);
      const minimum = isLargeText ? 3 : 4.5;
      if (ratio < minimum) textIssues.push(`${text.slice(0, 70)} (${ratio.toFixed(2)}:1 attendu ${minimum}:1)`);
    }

    addRow(
      "COULEURS ET CONTRASTE",
      "Assurer un contraste suffisamment élevé entre texte et arrière-plan",
      "Installer et lancer Color Contrast Analyser.",
      "Color Contrast Analyser affiche 'Conforme' pour les critères AA : Texte normal : taille inférieure à 24px ou à 18,5px gras. Grand texte : Taille supérieure ou égale à 24px ou à 18,5px gras. Contenu non textuel : indicateurs de focus, graphiques, icônes, liens non soulignés.",
      textIssues.length === 0 ? "OK" : "KO",
      textIssues.length === 0
        ? "Aucun écart de contraste détecté automatiquement."
        : `Contrastes insuffisants : ${textIssues.slice(0, 5).join(" | ")}`
    );

    const colorStatusElements = [...document.querySelectorAll(
      "[class*='error'],[class*='success'],[class*='warning'],[class*='danger'],[class*='alert'],[class*='status'],[class*='badge'],[class*='tag']"
    )].filter(isVisible);

    addRow(
      "COULEURS ET CONTRASTE",
      "S'assurer que l'information n'est pas transmise uniquement par la couleur",
      "Installer et lancer Color Contrast Analyser.",
      "L'information transmise par la couleur peut également être obtenue par un texte explicite.",
      colorStatusElements.length === 0 ? "NA" : "À vérifier",
      colorStatusElements.length === 0
        ? "Aucun élément de statut/couleur détecté."
        : `${colorStatusElements.length} élément(s) de statut/couleur détecté(s). Vérifier qu'un texte explicite accompagne chaque information colorée.`
    );

    const statusWithIcons = colorStatusElements.filter(el => {
      const hasIconChild = el.querySelector("svg, img, [class*='icon'], [class*='fa-'], [class*='bi-'], [class*='material-icons']") !== null;
      const beforeContent = getComputedStyle(el, "::before").content;
      const afterContent = getComputedStyle(el, "::after").content;
      const hasPseudoContent =
        (beforeContent && beforeContent !== "none" && beforeContent !== '""' && beforeContent !== "''") ||
        (afterContent && afterContent !== "none" && afterContent !== '""' && afterContent !== "''");
      return hasIconChild || hasPseudoContent;
    });

    addRow(
      "COULEURS ET CONTRASTE",
      "S'assurer que l'information n'est pas transmise uniquement par la couleur",
      "S'assurer que l'information n'est pas transmise uniquement par la couleur.",
      "L'information transmise par la couleur est complétée par une autre information visuelle (exemple : icônes utilisant des couleurs et formes différentes).",
      colorStatusElements.length === 0 ? "NA" : "À vérifier",
      colorStatusElements.length === 0
        ? "Aucun élément de statut/couleur détecté."
        : statusWithIcons.length > 0
          ? `${statusWithIcons.length}/${colorStatusElements.length} élément(s) semblent accompagnés d'une icône ou forme. Vérification manuelle nécessaire.`
          : `${colorStatusElements.length} élément(s) de statut détectés sans icône détectée. Vérifier si une forme ou icône complète l'information colorée.`
    );

    const textLinks = [...document.querySelectorAll("p a, li a, td a, th a, dd a, dt a, span a")].filter(isVisible);
    const linksWithoutUnderline = textLinks.filter(el => {
      const style = getComputedStyle(el);
      const textDecoration = style.textDecorationLine || style.textDecoration || "";
      const borderBottomWidth = parseFloat(style.borderBottomWidth) || 0;
      return !textDecoration.includes("underline") && borderBottomWidth === 0;
    });

    addRow(
      "COULEURS ET CONTRASTE",
      "S'assurer que l'information n'est pas transmise uniquement par la couleur",
      "Installer et lancer Color Contrast Analyser.",
      "Cas particulier des liens dans le texte : s'ils ne sont pas soulignés, au focus clavier et au survol souris, fournir un autre moyen que la couleur pour les distinguer.",
      textLinks.length === 0 ? "NA" : linksWithoutUnderline.length === 0 ? "OK" : "À vérifier",
      textLinks.length === 0
        ? "Aucun lien dans du texte détecté."
        : linksWithoutUnderline.length === 0
          ? `${textLinks.length} lien(s) dans le texte détectés, tous soulignés.`
          : `${linksWithoutUnderline.length} lien(s) non soulignés. Vérifier si un autre indicateur visuel est présent au focus et au survol.`
    );

    // --- NAVIGATION GÉNÉRALE : détection améliorée carrousels + bouton pause ---

    const carouselSelectors = [
      "[class*='carousel']", "[class*='slider']", "[class*='swiper']",
      "[class*='splide']", "[class*='owl']", "[class*='glide']",
      "[class*='slideshow']", "[data-ride]", "[class*='slick']"
    ];
    const carousels = [...document.querySelectorAll(carouselSelectors.join(","))].filter(isVisible);
    const autoPlayVideos = [...document.querySelectorAll("video[autoplay]")].filter(isVisible);
    const allAnimated = [...carousels, ...autoPlayVideos];

    const pauseControls = allAnimated.filter(el =>
      el.querySelector(
        'button[aria-label*="pause" i], button[aria-label*="stop" i], button[aria-label*="arrêt" i], ' +
        '[class*="pause"], [class*="stop"], video[controls]'
      )
    );

    const hasReducedMotion = (() => {
      try {
        return [...document.styleSheets].some(sheet => {
          try {
            return [...sheet.cssRules].some(rule =>
              rule.media && rule.media.mediaText && rule.media.mediaText.includes("prefers-reduced-motion")
            );
          } catch { return false; }
        });
      } catch { return false; }
    })();

    const animationStatus =
      allAnimated.length === 0 ? "NA" :
      (pauseControls.length >= allAnimated.length || hasReducedMotion) ? "À vérifier" : "KO";

    addRow(
      "NAVIGATION GÉNÉRALE",
      "Permettre le contrôle des animations",
      "Identifier tout contenu en mouvement, mis à jour automatiquement, clignotant ou en défilement, durant plus de 5 secondes et lancé automatiquement (exemple : un carrousel).",
      "L'utilisateur peut mettre pause ou masquer les animations, les mouvements, les mises à jour ou les clignotements.",
      animationStatus,
      allAnimated.length === 0
        ? "Aucun carrousel, slider ou vidéo autoplay détecté."
        : hasReducedMotion
          ? `${allAnimated.length} élément(s) animé(s) détecté(s). Support prefers-reduced-motion détecté dans les CSS.`
          : pauseControls.length > 0
            ? `${pauseControls.length}/${allAnimated.length} élément(s) animé(s) avec contrôle pause/stop détecté. Vérification manuelle nécessaire.`
            : `${allAnimated.length} élément(s) animé(s) détecté(s) SANS bouton pause/stop ni prefers-reduced-motion.`
    );

    // --- NAVIGATION CLAVIER : placeholders mis à jour par Playwright après page.evaluate ---

    const interactive = [...document.querySelectorAll("a, button, input, select, textarea, [tabindex], [role='button'], [role='link']")].filter(isVisible);
    const tabIndexMinus1 = interactive.filter(el => {
      const tag = el.tagName.toLowerCase();
      const tabindex = el.getAttribute("tabindex");
      return tabindex === "-1" && ["a", "button", "input", "select", "textarea"].includes(tag);
    });

    addRow(
      "NAVIGATION CLAVIER",
      "Permettre l'utilisation de l'application au clavier",
      "Parcourir la page au clavier à l'aide des touches Tab ou Shift + Tab. Utiliser tous les éléments interactifs (en tapant sur les touches Entrée, Espace pour les boutons/liens, et les flèches directionnelles pour certains composants : une série de boutons radio, un système d'onglets…).",
      "Tous les éléments interactifs sont atteignables en naviguant au clavier.",
      tabIndexMinus1.length > 0 ? "KO" : "À vérifier",
      tabIndexMinus1.length > 0
        ? `${tabIndexMinus1.length} élément(s) avec tabindex="-1" exclu(s) de la navigation clavier.`
        : `${interactive.length} élément(s) interactif(s) détecté(s). Test clavier en cours...`
    );

    addRow(
      "NAVIGATION CLAVIER",
      "Permettre l'utilisation de l'application au clavier",
      "Parcourir la page au clavier à l'aide des touches Tab ou Shift + Tab. Utiliser tous les éléments interactifs (en tapant sur les touches Entrée, Espace pour les boutons/liens, et les flèches directionnelles pour certains composants : une série de boutons radio, un système d'onglets…).",
      "Tous les éléments interactifs sont utilisables depuis des interactions clavier.",
      "À vérifier",
      "Test clavier en cours..."
    );

    // --- MISE EN PAGE : détection améliorée des conteneurs à risque ---

    const fixedHeightIssues = [];
    const containers = [...document.querySelectorAll("p, div, li, td, th, section, article, header, nav, aside, footer, span")].filter(isVisible).slice(0, 300);
    for (const el of containers) {
      const style = getComputedStyle(el);
      const height = style.height;
      const overflow = style.overflow;
      const overflowY = style.overflowY;
      if (
        (overflow === "hidden" || overflowY === "hidden") &&
        height !== "auto" && height !== "0px" &&
        !height.includes("%") && !height.includes("vh") &&
        parseFloat(height) > 20
      ) {
        fixedHeightIssues.push(`${el.tagName} h=${height}`);
      }
    }

    const pxFontCount = [...document.querySelectorAll("*")].filter(el => {
      const inline = el.getAttribute("style") || "";
      return /font-size\s*:\s*\d+px/i.test(inline);
    }).length;

    addRow(
      "MISE EN PAGE",
      "Utiliser des tailles relatives et faire du web adaptatif (responsive)",
      "Avec Firefox, à partir du menu 'Affichage', sélectionner 'Zoom' puis 'Agrandir uniquement le texte' et activer un niveau de zoom à 200%.",
      "Absence de contenus tronqués ou masqués et absence de fonctionnalités inutilisables.",
      fixedHeightIssues.length > 0 || pxFontCount > 0 ? "À vérifier" : "À vérifier",
      fixedHeightIssues.length > 0
        ? `${fixedHeightIssues.length} conteneur(s) à hauteur fixe avec overflow:hidden détecté(s) — risque de troncature au zoom 200%. Vérifier manuellement.`
        : pxFontCount > 0
          ? `${pxFontCount} élément(s) avec font-size en px inline détecté(s) — peut ne pas s'agrandir avec le zoom texte. Vérifier manuellement.`
          : "Aucun risque évident détecté automatiquement. Test manuel à 200% requis sous Firefox."
    );

    // --- FORMULAIRES ---

    const fields = [...document.querySelectorAll("input, select, textarea")]
      .filter(isVisible)
      .filter(el => {
        const type = (el.getAttribute("type") || "").toLowerCase();
        return !["hidden", "submit", "button", "reset", "image"].includes(type);
      });

    function getAccessibleName(el) {
      const id = el.getAttribute("id");
      const ariaLabel = el.getAttribute("aria-label");
      const ariaLabelledby = el.getAttribute("aria-labelledby");
      const title = el.getAttribute("title");
      const placeholder = el.getAttribute("placeholder");
      let labelText = "";
      if (id) {
        const label = document.querySelector(`label[for="${CSS.escape(id)}"]`);
        if (label) labelText = label.innerText.trim();
      }
      const parentLabel = el.closest("label");
      if (!labelText && parentLabel) labelText = parentLabel.innerText.trim();
      if (ariaLabelledby) {
        labelText = ariaLabelledby
          .split(/\s+/)
          .map(id => document.getElementById(id)?.innerText?.trim())
          .filter(Boolean)
          .join(" ");
      }
      return {
        name: labelText || ariaLabel || title || "",
        onlyPlaceholder: !labelText && !ariaLabel && !title && !!placeholder,
        placeholder
      };
    }

    const fieldsWithoutName = [], fieldsOnlyPlaceholder = [], labelsToCheck = [];
    for (const field of fields) {
      const info = getAccessibleName(field);
      if (!info.name) fieldsWithoutName.push(field.outerHTML.slice(0, 120));
      if (info.onlyPlaceholder) fieldsOnlyPlaceholder.push(info.placeholder);
      if (info.name) labelsToCheck.push(info.name);
    }

    addRow(
      "FORMULAIRES",
      "S'assurer qu'un nom accessible est associé à chaque champ de formulaire",
      "Utiliser l'inspecteur de code du navigateur sur l'onglet 'Accessibilité'.",
      "Chaque champ a au moins un nom accessible pertinent et contient au moins le texte de l'étiquette de champ visible à l'écran (un placeholder n'est pas conforme).",
      fields.length === 0 ? "NA" : fieldsWithoutName.length === 0 && fieldsOnlyPlaceholder.length === 0 ? "OK" : "KO",
      fields.length === 0
        ? "Aucun champ de formulaire détecté."
        : fieldsWithoutName.length > 0
          ? `${fieldsWithoutName.length} champ(s) sans nom accessible.`
          : fieldsOnlyPlaceholder.length > 0
            ? `${fieldsOnlyPlaceholder.length} champ(s) utilisent uniquement un placeholder (non conforme).`
            : `${fields.length} champ(s) analysé(s), nom accessible détecté.`
    );

    addRow(
      "FORMULAIRES",
      "S'assurer qu'un nom accessible est associé à chaque champ de formulaire",
      "Utiliser l'inspecteur de code du navigateur sur l'onglet 'Accessibilité'.",
      "Chaque étiquette associée à un champ de formulaire est-elle pertinente (hors cas particuliers) ?",
      fields.length === 0 ? "NA" : "À vérifier",
      fields.length === 0
        ? "Aucun champ de formulaire détecté."
        : `Labels détectés : ${labelsToCheck.slice(0, 6).join(" | ")}. Pertinence à vérifier manuellement.`
    );

    const forms = [...document.querySelectorAll("form")].filter(isVisible);
    const requiredFields = fields.filter(el =>
      el.required || el.getAttribute("aria-required") === "true" || el.hasAttribute("required")
    );

    addRow(
      "FORMULAIRES",
      "S'assurer que les messages d'erreurs sont pertinents",
      "Renseigner les formulaires avec des données erronées et des champs obligatoires laissés vides. Soumettre le formulaire.",
      "Les messages d'erreurs sont présents, pertinents, et identifient les champs en erreur.",
      forms.length === 0 ? "NA" : "À vérifier",
      forms.length === 0
        ? "Aucun formulaire détecté."
        : `${forms.length} formulaire(s), ${requiredFields.length} champ(s) obligatoire(s). Test de soumission en cours...`
    );

    return result;
  });

  // --- Amélioration test 16 & 17 : simulation clavier réelle ---
  try {
    const kbResult = await testKeyboardNavigation(page);
    const row16 = result.rows.find(r => r.expected && r.expected.includes("atteignables"));
    const row17 = result.rows.find(r => r.expected && r.expected.includes("utilisables"));
    if (row16) { row16.status = kbResult.reachableStatus; row16.comment = kbResult.reachableComment; }
    if (row17) { row17.status = kbResult.usableStatus; row17.comment = kbResult.usableComment; }
  } catch (e) { /* garder résultats statiques */ }

  // --- Amélioration test 25 : soumission formulaire réelle ---
  try {
    const formResult = await testFormSubmission(page);
    const row25 = result.rows.find(r => r.expected && r.expected.includes("messages d'erreurs"));
    if (row25 && formResult.status !== "À vérifier") {
      row25.status = formResult.status;
      row25.comment = formResult.comment;
    } else if (row25 && formResult.status === "À vérifier") {
      row25.comment = formResult.comment;
    }
  } catch (e) { /* garder résultats statiques */ }

  return result;
}

(async () => {
  if (!fs.existsSync("screenshots")) fs.mkdirSync("screenshots");

  const browser = await chromium.launch({ headless: true });
  const workbook = new ExcelJS.Workbook();

  workbook.creator = "Bot audit accessibilité";
  workbook.created = new Date();

  for (const { clientName, url } of urlEntries) {
    const defaultName = new URL(url).hostname.replace("www.", "");
    const rawSheetName = (clientName || defaultName).substring(0, 31);

    let ws = workbook.getWorksheet(rawSheetName);
    const sheetName = ws
      ? rawSheetName.substring(0, 25) + "_" + Math.floor(Math.random() * 999)
      : rawSheetName;

    ws = workbook.addWorksheet(sheetName);

    ws.columns = [
      { header: "Thématique", key: "theme", width: 26 },
      { header: "Référence \"incontournable\"", key: "ref", width: 42 },
      { header: "Tests à réaliser", key: "test", width: 55 },
      { header: "Résultat attendu", key: "expected", width: 80 },
      { header: "Conformité", key: "status", width: 16 },
      { header: "Commentaire", key: "comment", width: 55 },
      { header: "URL de la page concernée", key: "url", width: 45 },
      { header: "Copie d'écran (si non conforme)", key: "screenshot", width: 45 },
      { header: "Nom et Date", key: "nameDate", width: 28 }
    ];

    const page = await browser.newPage({ viewport: { width: 1366, height: 768 } });

    try {
      console.log(`Audit : ${clientName ? clientName + " | " : ""}${url}`);

      const audit = await auditSite(page, url);
      const hasKO = audit.rows.some(r => r.status === "KO");

      let screenshotPath = "";
      if (hasKO) {
        const safeName = url.replace(/https?:\/\//, "").replace(/[^\w]/g, "_");
        screenshotPath = `screenshots/${safeName}.jpg`;
        await page.screenshot({ path: screenshotPath, type: "jpeg", quality: 45, fullPage: false });
      }

      let imageInserted = false;
      for (const row of audit.rows) {
        const addedRow = ws.addRow({
          theme: row.theme,
          ref: row.ref,
          test: row.test,
          expected: row.expected,
          status: row.status,
          comment: row.comment,
          url: row.url,
          screenshot: row.status === "KO" ? screenshotPath : "",
          nameDate: `Audit auto - ${new Date().toLocaleDateString("fr-FR")}`
        });

        if (row.status === "KO" && screenshotPath && !imageInserted) {
          const imageId = workbook.addImage({ filename: screenshotPath, extension: "jpeg" });
          ws.addImage(imageId, { tl: { col: 7, row: addedRow.number - 1 }, ext: { width: 220, height: 120 } });
          ws.getRow(addedRow.number).height = 95;
          imageInserted = true;
        }
      }
    } catch (error) {
      ws.addRow({
        theme: "ERREUR", ref: "Page inaccessible", test: "Chargement de la page",
        expected: "La page doit être accessible au bot.", status: "KO",
        comment: error.message, url, screenshot: "",
        nameDate: `Audit auto - ${new Date().toLocaleDateString("fr-FR")}`
      });
    }

    await page.close();

    ws.getRow(1).font = { bold: true, color: { argb: "FF000000" } };
    ws.getRow(1).fill = { type: "pattern", pattern: "solid", fgColor: { argb: "FFFF7F00" } };
    ws.getRow(1).alignment = { vertical: "middle", horizontal: "center", wrapText: true };
    ws.getRow(1).height = 28;
    ws.autoFilter = { from: "A1", to: "I1" };
    ws.views = [{ state: "frozen", ySplit: 1 }];

    ws.eachRow((row, rowNumber) => {
      row.alignment = { vertical: "top", wrapText: true };
      if (rowNumber > 1 && row.height < 55) row.height = 55;

      const themeCell = row.getCell(1);
      const refCell = row.getCell(2);
      const statusCell = row.getCell(5);

      if (rowNumber > 1) {
        const theme = String(themeCell.value || "");
        if (theme.includes("COULEURS ET CONTRASTE") || theme.includes("NAVIGATION") || theme.includes("FORMULAIRES")) {
          themeCell.fill = { type: "pattern", pattern: "solid", fgColor: { argb: "FFFFFF00" } };
          refCell.fill = { type: "pattern", pattern: "solid", fgColor: { argb: "FFFFFF00" } };
        }

        const statusColors = { "OK": "FF92D050", "KO": "FFFF0000", "NA": "FFD9D9D9", "À vérifier": "FFFFD966" };
        if (statusColors[statusCell.value]) {
          statusCell.fill = { type: "pattern", pattern: "solid", fgColor: { argb: statusColors[statusCell.value] } };
        }

        row.eachCell(cell => {
          cell.border = {
            top: { style: "thin", color: { argb: "FF999999" } },
            bottom: { style: "thin", color: { argb: "FF999999" } },
            left: { style: "thin", color: { argb: "FF999999" } },
            right: { style: "thin", color: { argb: "FF999999" } }
          };
        });
      }
    });
  }

  await browser.close();
  await workbook.xlsx.writeFile("audit-accessibilite-rgaa.xlsx");
  console.log("✅ Audit terminé : audit-accessibilite-rgaa.xlsx");
})();
