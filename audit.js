const { chromium } = require("playwright");
const ExcelJS = require("exceljs");
const fs = require("fs");

const urls = fs.readFileSync("urls.txt", "utf8")
  .split(/\r?\n/)
  .map(u => u.trim())
  .filter(Boolean);

function srgbToLin(c) {
  c = c / 255;
  return c <= 0.03928 ? c / 12.92 : Math.pow((c + 0.055) / 1.055, 2.4);
}

function luminance(rgb) {
  return 0.2126 * srgbToLin(rgb[0]) + 0.7152 * srgbToLin(rgb[1]) + 0.0722 * srgbToLin(rgb[2]);
}

function contrastRatio(fg, bg) {
  const L1 = luminance(fg);
  const L2 = luminance(bg);
  return (Math.max(L1, L2) + 0.05) / (Math.min(L1, L2) + 0.05);
}

async function auditSite(page, url) {
  await page.goto(url, { waitUntil: "domcontentloaded", timeout: 45000 });
  await page.waitForTimeout(2500);

  return await page.evaluate(({ contrastRatioString }) => {
    const contrastRatio = eval("(" + contrastRatioString + ")");

    function parseRgb(value) {
      const match = String(value || "").match(/rgba?\((\d+),\s*(\d+),\s*(\d+)/);
      return match ? [Number(match[1]), Number(match[2]), Number(match[3])] : null;
    }

    function isVisible(el) {
      const style = getComputedStyle(el);
      const rect = el.getBoundingClientRect();
      return style.display !== "none" &&
             style.visibility !== "hidden" &&
             rect.width > 0 &&
             rect.height > 0;
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

    const result = {
      title: document.title || "",
      pageLang: document.documentElement.getAttribute("lang") || "",
      rows: []
    };

    function addRow(theme, ref, test, expected, status, comment) {
      result.rows.push({
        theme,
        ref,
        test,
        expected,
        status,
        comment,
        url: location.href
      });
    }

    // 1. Title
    addRow(
      "CONTENU TEXTUEL",
      "Donner un titre aux pages",
      "Lancer l'inspecteur de code du navigateur et examiner le titre de page (<title>[titre de la page]</title>).",
      "Chaque page possède un titre unique et descriptif du contenu.",
      document.title && document.title.trim().length > 5 ? "OK" : "KO",
      document.title ? `Titre détecté : ${document.title}` : "Aucun titre de page détecté."
    );

    // 2. Headings semantics
    const headings = [...document.querySelectorAll("h1,h2,h3,h4,h5,h6")].filter(isVisible);
    const headingComment = headings.slice(0, 8).map(h => `${h.tagName}: ${(h.innerText || "").trim().slice(0, 50)}`).join(" | ");
    addRow(
      "CONTENU TEXTUEL",
      "Donner un titre aux rubriques",
      "Installer puis lancer le bookmarklet Headings ou inspecter les balises Hn.",
      "Tous les contenus traités visuellement comme des titres possèdent une sémantique de titre (balises <h1> à <h6>).",
      headings.length > 0 ? "OK" : "KO",
      headings.length > 0 ? `${headings.length} titre(s) Hn détecté(s). ${headingComment}` : "Aucun titre Hn détecté."
    );

    // 3. Headings hierarchy
    const levels = headings.map(h => Number(h.tagName.substring(1)));
    let hierarchyIssue = false;
    for (let i = 1; i < levels.length; i++) {
      if (levels[i] - levels[i - 1] > 1) hierarchyIssue = true;
    }
    addRow(
      "CONTENU TEXTUEL",
      "Donner un titre aux rubriques",
      "Installer puis lancer le bookmarklet Headings ou inspecter la hiérarchie Hn.",
      "Les titres de niveaux sont hiérarchisés de manière à refléter leur poids sémantique.",
      headings.length === 0 || hierarchyIssue ? "KO" : "OK",
      headings.length === 0 ? "Impossible de vérifier : aucun Hn détecté." : hierarchyIssue ? `Saut de niveau détecté dans la structure : ${levels.join(" > ")}` : `Hiérarchie détectée : ${levels.join(" > ")}`
    );

    // 4. Lang presence
    addRow(
      "CONTENU TEXTUEL",
      "Indiquer la langue principale",
      "Lancer l'inspecteur de code du navigateur. Examiner l'élément <html>.",
      "Un attribut lang est présent dans l'élément <html> de la page.",
      result.pageLang ? "OK" : "KO",
      result.pageLang ? `Lang détecté : ${result.pageLang}` : "Attribut lang absent sur <html>."
    );

    // 5. Lang value
    addRow(
      "CONTENU TEXTUEL",
      "Indiquer la langue principale",
      "Lancer l'inspecteur de code du navigateur. Examiner l'élément <html>.",
      "La valeur de l'attribut lang correspond à la langue principale du document.",
      result.pageLang ? "OK" : "KO",
      result.pageLang ? `Valeur lang : ${result.pageLang}. À confirmer selon la langue réelle du contenu.` : "Langue principale non déclarée."
    );

    // Images
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
      imageLinks.length === 0 ? "Aucune image-lien détectée." : `${imageLinks.length} image(s)-lien détectée(s). ${imageLinks.filter(img => !img.hasAttribute("alt")).length} sans alt.`
    );

    addRow(
      "CONTENU NON TEXTUEL",
      "S'assurer que les images ont une alternative textuelle",
      "Installer puis lancer le bookmarklet List Images ou l'inspecteur de code.",
      "Image porteuse d'information : l'attribut alt de chaque image est pertinent par rapport au rôle de l'image dans la page.",
      missingAlt.length === 0 ? "OK" : "KO",
      missingAlt.length === 0 ? `${images.length} image(s) analysée(s), pas d'image sans alt détectée.` : `${missingAlt.length} image(s) sans attribut alt.`
    );

    addRow(
      "CONTENU NON TEXTUEL",
      "S'assurer que les images ont une alternative textuelle",
      "Installer puis lancer le bookmarklet List Images ou l'inspecteur de code.",
      "Image décorative : l'attribut alt est présent mais vide.",
      emptyAlt.length > 0 ? "OK" : "NA",
      emptyAlt.length > 0 ? `${emptyAlt.length} image(s) avec alt vide détectée(s), à confirmer comme décoratives.` : "Aucune image décorative avec alt vide détectée."
    );

    // Contraste texte/fond + non textuel
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

      if (ratio < minimum) {
        textIssues.push(`${text.slice(0, 70)} (${ratio.toFixed(2)}:1 attendu ${minimum}:1)`);
      }
    }

    addRow(
      "COULEURS ET CONTRASTE",
      "Assurer un contraste suffisamment élevé entre texte et arrière-plan",
      "Installer et lancer Color Contrast Analyser.",
      "Color Contrast Analyser affiche Conforme pour les critères AA : texte normal 4.5:1, grand texte 3:1, contenu non textuel 3:1.",
      textIssues.length === 0 ? "OK" : "KO",
      textIssues.length === 0 ? "Aucun écart de contraste détecté automatiquement." : `Contrastes insuffisants détectés : ${textIssues.slice(0, 5).join(" | ")}`
    );

    // Info couleur seule
    const colorStatusElements = [...document.querySelectorAll("[class*='error'],[class*='success'],[class*='warning'],[class*='danger'],[class*='alert'],[class*='status'],[class*='badge'],[class*='tag']")].filter(isVisible);

    addRow(
      "COULEURS ET CONTRASTE",
      "S'assurer que l'information n'est pas transmise uniquement par la couleur",
      "Inspecter les messages d'erreur, statuts, badges, graphiques et indicateurs colorés.",
      "L'information transmise par la couleur peut également être obtenue par un texte explicite.",
      colorStatusElements.length === 0 ? "NA" : "À vérifier",
      colorStatusElements.length === 0 ? "Aucun élément de statut/couleur évident détecté." : `${colorStatusElements.length} élément(s) de statut/couleur détecté(s). Vérifier la présence d'un texte explicite.`
    );

    // Double coding
    const noDoubleCoding = colorStatusElements.filter(el => {
      const hasText = (el.innerText || "").trim().length > 1;
      const hasIcon = !!el.querySelector("svg,img,i,[class*='icon']");
      const hasAria = !!el.getAttribute("aria-label");
      return !hasText && !hasIcon && !hasAria;
    });

    addRow(
      "COULEURS ET CONTRASTE",
      "S'assurer que l'information n'est pas transmise uniquement par la couleur",
      "Inspecter les éléments colorés : icônes, formes, badges, messages, statuts.",
      "L'information transmise par la couleur est complétée par une autre information visuelle.",
      noDoubleCoding.length === 0 ? "OK" : "KO",
      noDoubleCoding.length === 0 ? "Aucun élément coloré sans double codage détecté automatiquement." : `${noDoubleCoding.length} élément(s) coloré(s) sans texte, icône ou aria-label détecté(s).`
    );

    // Links
    const paragraphLinks = [...document.querySelectorAll("p a, li a, article a, main a")].filter(isVisible);
    const linkIssues = paragraphLinks.filter(a => {
      const style = getComputedStyle(a);
      const underlined = String(style.textDecorationLine || style.textDecoration).includes("underline");
      const bold = (parseInt(style.fontWeight) || 400) >= 700;
      const hasIcon = !!a.querySelector("svg,img");
      return !underlined && !bold && !hasIcon;
    });

    addRow(
      "COULEURS ET CONTRASTE",
      "S'assurer que l'information n'est pas transmise uniquement par la couleur",
      "Inspecter les liens dans le texte, puis vérifier hover et focus clavier.",
      "Cas particulier des liens dans du texte : s'ils ne sont pas soulignés, au focus clavier et au survol souris, fournir un autre moyen que la couleur pour les distinguer.",
      linkIssues.length === 0 ? "OK" : "KO",
      linkIssues.length === 0 ? "Les liens analysés semblent distinguables." : `${linkIssues.length} lien(s) dans le texte semblent non soulignés ou non distinguables autrement que par la couleur.`
    );

    return result;
  }, { contrastRatioString: contrastRatio.toString() });
}

(async () => {
  const browser = await chromium.launch({ headless: true });
  const workbook = new ExcelJS.Workbook();
  workbook.creator = "Bot audit accessibilité";
  workbook.created = new Date();

  for (const url of urls) {
    const sheetName = new URL(url).hostname.replace("www.", "").substring(0, 31);
    let ws = workbook.getWorksheet(sheetName);
    if (ws) ws = workbook.addWorksheet(sheetName.substring(0, 25) + "_" + Math.floor(Math.random() * 999));
    else ws = workbook.addWorksheet(sheetName);

    ws.columns = [
      { header: "Thématique", key: "theme", width: 26 },
      { header: "Référence \"incontournable\"", key: "ref", width: 34 },
      { header: "Tests à réaliser", key: "test", width: 55 },
      { header: "Résultat attendu", key: "expected", width: 70 },
      { header: "Conformité", key: "status", width: 16 },
      { header: "Commentaire", key: "comment", width: 75 },
      { header: "URL de la page concernée", key: "url", width: 45 },
      { header: "Copie d'écran (si non conforme)", key: "screenshot", width: 35 },
      { header: "Nom et Date", key: "nameDate", width: 25 }
    ];

    const page = await browser.newPage({ viewport: { width: 1366, height: 900 } });

    try {
      console.log("Audit :", url);
      const audit = await auditSite(page, url);

      for (const row of audit.rows) {
        ws.addRow({
          theme: row.theme,
          ref: row.ref,
          test: row.test,
          expected: row.expected,
          status: row.status,
          comment: row.comment,
          url: row.url,
          screenshot: "",
          nameDate: `Audit auto - ${new Date().toLocaleDateString("fr-FR")}`
        });
      }
    } catch (error) {
      ws.addRow({
        theme: "ERREUR",
        ref: "Page inaccessible",
        test: "Chargement de la page",
        expected: "La page doit être accessible au bot.",
        status: "KO",
        comment: error.message,
        url,
        screenshot: "",
        nameDate: `Audit auto - ${new Date().toLocaleDateString("fr-FR")}`
      });
    }

    await page.close();

    ws.getRow(1).font = { bold: true, color: { argb: "FFFFFFFF" } };
    ws.getRow(1).fill = { type: "pattern", pattern: "solid", fgColor: { argb: "FFFF7F00" } };
    ws.getRow(1).alignment = { vertical: "middle", horizontal: "center", wrapText: true };
    ws.autoFilter = { from: "A1", to: "I1" };
    ws.views = [{ state: "frozen", ySplit: 1 }];

    ws.eachRow((row, rowNumber) => {
      row.alignment = { vertical: "top", wrapText: true };
      if (rowNumber > 1) row.height = 55;

      const statusCell = row.getCell(5);
      if (statusCell.value === "OK") {
        statusCell.fill = { type: "pattern", pattern: "solid", fgColor: { argb: "FF92D050" } };
      } else if (statusCell.value === "KO") {
        statusCell.fill = { type: "pattern", pattern: "solid", fgColor: { argb: "FFFF6961" } };
      } else if (statusCell.value === "NA") {
        statusCell.fill = { type: "pattern", pattern: "solid", fgColor: { argb: "FFD9D9D9" } };
      } else if (statusCell.value === "À vérifier") {
        statusCell.fill = { type: "pattern", pattern: "solid", fgColor: { argb: "FFFFD966" } };
      }

      row.eachCell(cell => {
        cell.border = {
          top: { style: "thin", color: { argb: "FF999999" } },
          bottom: { style: "thin", color: { argb: "FF999999" } },
          left: { style: "thin", color: { argb: "FF999999" } },
          right: { style: "thin", color: { argb: "FF999999" } }
        };
      });
    });
  }

  await browser.close();
  await workbook.xlsx.writeFile("audit-accessibilite-rgaa.xlsx");
  console.log("✅ Audit terminé : audit-accessibilite-rgaa.xlsx");
})();
