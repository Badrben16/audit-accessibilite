const { chromium } = require("playwright");
const ExcelJS = require("exceljs");
const fs = require("fs");

const urls = fs.readFileSync("urls.txt", "utf8")
  .split(/\r?\n/)
  .map(u => u.trim())
  .filter(Boolean);

function getRecommendation(row) {
  if (row.status !== "KO" && row.status !== "À vérifier") return "";

  const ref = `${row.theme} ${row.ref} ${row.expected}`.toLowerCase();

  if (ref.includes("contraste")) {
    return "Corriger les contrastes selon le RGAA/WCAG AA : 4.5:1 pour le texte normal, 3:1 pour le grand texte et 3:1 pour les composants graphiques essentiels.";
  }
  if (ref.includes("couleur")) {
    return "Ne pas transmettre une information uniquement par la couleur. Ajouter un texte explicite, une icône, une forme ou un libellé visible.";
  }
  if (ref.includes("lien")) {
    return "Rendre les liens identifiables sans dépendre uniquement de la couleur : soulignement, graisse, icône, focus visible ou autre indicateur visuel.";
  }
  if (ref.includes("titre") || ref.includes("rubrique")) {
    return "Structurer la page avec une hiérarchie de titres logique : H1 principal, puis H2, H3, sans saut incohérent.";
  }
  if (ref.includes("image") || ref.includes("alternative")) {
    return "Ajouter un attribut alt pertinent aux images porteuses d’information. Pour les images décoratives, utiliser alt=\"\".";
  }
  if (ref.includes("langue") || ref.includes("lang")) {
    return "Déclarer correctement la langue principale avec l’attribut lang sur la balise html, par exemple <html lang=\"fr\">.";
  }
  if (ref.includes("formulaire") || ref.includes("champ")) {
    return "Associer chaque champ à un nom accessible via label, aria-label ou aria-labelledby. Le placeholder seul n’est pas suffisant.";
  }
  if (ref.includes("erreur")) {
    return "Afficher des messages d’erreur explicites, reliés au champ concerné, compréhensibles et annoncés aux technologies d’assistance.";
  }

  return "Corriger l’écart selon les critères RGAA applicables.";
}

async function auditSite(page, url) {
  await page.goto(url, { waitUntil: "domcontentloaded", timeout: 45000 });
  await page.waitForTimeout(2500);

  return await page.evaluate(() => {
    function srgbToLin(c) {
      c = c / 255;
      return c <= 0.03928 ? c / 12.92 : Math.pow((c + 0.055) / 1.055, 2.4);
    }

    function luminance(rgb) {
      return 0.2126 * srgbToLin(rgb[0]) +
             0.7152 * srgbToLin(rgb[1]) +
             0.0722 * srgbToLin(rgb[2]);
    }

    function contrastRatio(fg, bg) {
      const L1 = luminance(fg);
      const L2 = luminance(bg);
      return (Math.max(L1, L2) + 0.05) / (Math.min(L1, L2) + 0.05);
    }

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

    addRow(
      "CONTENU TEXTUEL",
      "Donner un titre aux pages",
      "Examiner le titre de page (<title>).",
      "Chaque page possède un titre unique et descriptif.",
      document.title && document.title.trim().length > 5 ? "OK" : "KO",
      document.title ? `Titre détecté : ${document.title}` : "Aucun titre de page détecté."
    );

    const headings = [...document.querySelectorAll("h1,h2,h3,h4,h5,h6")].filter(isVisible);
    const headingComment = headings
      .slice(0, 8)
      .map(h => `${h.tagName}: ${(h.innerText || "").trim().slice(0, 50)}`)
      .join(" | ");

    addRow(
      "CONTENU TEXTUEL",
      "Donner un titre aux rubriques",
      "Inspecter les balises Hn.",
      "Tous les titres visuels possèdent une sémantique de titre H1 à H6.",
      headings.length > 0 ? "OK" : "KO",
      headings.length > 0 ? `${headings.length} titre(s) Hn détecté(s). ${headingComment}` : "Aucun titre Hn détecté."
    );

    const levels = headings.map(h => Number(h.tagName.substring(1)));
    let hierarchyIssue = false;

    for (let i = 1; i < levels.length; i++) {
      if (levels[i] - levels[i - 1] > 1) hierarchyIssue = true;
    }

    addRow(
      "CONTENU TEXTUEL",
      "Donner un titre aux rubriques",
      "Inspecter la hiérarchie des titres.",
      "Les titres sont hiérarchisés de manière logique.",
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
      "Examiner l'élément <html>.",
      "Un attribut lang est présent dans <html>.",
      result.pageLang ? "OK" : "KO",
      result.pageLang ? `Lang détecté : ${result.pageLang}` : "Attribut lang absent sur <html>."
    );

    addRow(
      "CONTENU TEXTUEL",
      "Indiquer la langue principale",
      "Examiner la valeur de l'attribut lang.",
      "La valeur de lang correspond à la langue principale du document.",
      result.pageLang ? "OK" : "KO",
      result.pageLang ? `Valeur lang : ${result.pageLang}. À confirmer selon le contenu réel.` : "Langue principale non déclarée."
    );

    const images = [...document.images].filter(isVisible);
    const imageLinks = images.filter(img => img.closest("a"));
    const missingAlt = images.filter(img => !img.hasAttribute("alt") && img.getAttribute("role") !== "presentation");
    const emptyAlt = images.filter(img => img.hasAttribute("alt") && img.getAttribute("alt") === "");

    addRow(
      "CONTENU NON TEXTUEL",
      "S'assurer que les images ont une alternative textuelle",
      "Inspecter les images-liens.",
      "Chaque image-lien possède une alternative pertinente.",
      imageLinks.length === 0 ? "NA" : imageLinks.every(img => img.hasAttribute("alt")) ? "OK" : "KO",
      imageLinks.length === 0
        ? "Aucune image-lien détectée."
        : `${imageLinks.length} image(s)-lien détectée(s). ${imageLinks.filter(img => !img.hasAttribute("alt")).length} sans alt.`
    );

    addRow(
      "CONTENU NON TEXTUEL",
      "S'assurer que les images ont une alternative textuelle",
      "Inspecter les images porteuses d'information.",
      "Chaque image porteuse d'information possède un alt pertinent.",
      missingAlt.length === 0 ? "OK" : "KO",
      missingAlt.length === 0
        ? `${images.length} image(s) analysée(s), pas d'image sans alt détectée.`
        : `${missingAlt.length} image(s) sans attribut alt.`
    );

    addRow(
      "CONTENU NON TEXTUEL",
      "S'assurer que les images ont une alternative textuelle",
      "Inspecter les images décoratives.",
      "Les images décoratives ont un alt vide.",
      emptyAlt.length > 0 ? "OK" : "NA",
      emptyAlt.length > 0
        ? `${emptyAlt.length} image(s) avec alt vide détectée(s), à confirmer comme décoratives.`
        : "Aucune image décorative avec alt vide détectée."
    );

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
      "Vérifier les ratios de contraste.",
      "Texte normal 4.5:1, grand texte 3:1, contenu non textuel 3:1.",
      textIssues.length === 0 ? "OK" : "KO",
      textIssues.length === 0
        ? "Aucun écart de contraste détecté automatiquement."
        : `Contrastes insuffisants détectés : ${textIssues.slice(0, 5).join(" | ")}`
    );

    const colorStatusElements = [...document.querySelectorAll("[class*='error'],[class*='success'],[class*='warning'],[class*='danger'],[class*='alert'],[class*='status'],[class*='badge'],[class*='tag']")].filter(isVisible);

    addRow(
      "COULEURS ET CONTRASTE",
      "S'assurer que l'information n'est pas transmise uniquement par la couleur",
      "Inspecter les messages, statuts, badges et indicateurs colorés.",
      "L'information transmise par la couleur est aussi disponible par un texte explicite.",
      colorStatusElements.length === 0 ? "NA" : "À vérifier",
      colorStatusElements.length === 0
        ? "Aucun élément de statut/couleur évident détecté."
        : `${colorStatusElements.length} élément(s) de statut/couleur détecté(s). Vérification manuelle nécessaire.`
    );

    const noDoubleCoding = colorStatusElements.filter(el => {
      const hasText = (el.innerText || "").trim().length > 1;
      const hasIcon = !!el.querySelector("svg,img,i,[class*='icon']");
      const hasAria = !!el.getAttribute("aria-label");
      return !hasText && !hasIcon && !hasAria;
    });

    addRow(
      "COULEURS ET CONTRASTE",
      "S'assurer que l'information n'est pas transmise uniquement par la couleur",
      "Inspecter le double codage visuel.",
      "La couleur est complétée par une icône, une forme, un texte ou un libellé.",
      noDoubleCoding.length === 0 ? "OK" : "KO",
      noDoubleCoding.length === 0
        ? "Aucun élément coloré sans double codage détecté automatiquement."
        : `${noDoubleCoding.length} élément(s) coloré(s) sans texte, icône ou aria-label.`
    );

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
      "Inspecter les liens dans le texte.",
      "Les liens non soulignés disposent d'un autre moyen visuel que la couleur seule.",
      linkIssues.length === 0 ? "OK" : "KO",
      linkIssues.length === 0
        ? "Les liens analysés semblent distinguables."
        : `${linkIssues.length} lien(s) semblent non distinguables autrement que par la couleur.`
    );

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
      if (!labelText && parentLabel) {
        labelText = parentLabel.innerText.trim();
      }

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

    const fieldsWithoutName = [];
    const fieldsOnlyPlaceholder = [];
    const labelsToCheck = [];

    for (const field of fields) {
      const info = getAccessibleName(field);
      if (!info.name) fieldsWithoutName.push(field.outerHTML.slice(0, 120));
      if (info.onlyPlaceholder) fieldsOnlyPlaceholder.push(info.placeholder);
      if (info.name) labelsToCheck.push(info.name);
    }

    addRow(
      "FORMULAIRES",
      "S'assurer qu'un nom accessible est associé à chaque champ de formulaire",
      "Inspecter le nom accessible des champs.",
      "Chaque champ possède un nom accessible pertinent. Le placeholder seul n'est pas conforme.",
      fields.length === 0 ? "NA" : (fieldsWithoutName.length === 0 && fieldsOnlyPlaceholder.length === 0 ? "OK" : "KO"),
      fields.length === 0
        ? "Aucun champ de formulaire détecté."
        : fieldsWithoutName.length > 0
          ? `${fieldsWithoutName.length} champ(s) sans nom accessible.`
          : fieldsOnlyPlaceholder.length > 0
            ? `${fieldsOnlyPlaceholder.length} champ(s) utilisent uniquement un placeholder.`
            : `${fields.length} champ(s) analysé(s), nom accessible détecté.`
    );

    addRow(
      "FORMULAIRES",
      "S'assurer qu'un nom accessible est associé à chaque champ de formulaire",
      "Inspecter la pertinence des étiquettes.",
      "Chaque étiquette associée à un champ est pertinente.",
      fields.length === 0 ? "NA" : "À vérifier",
      fields.length === 0
        ? "Aucun champ de formulaire détecté."
        : `Labels/noms détectés : ${labelsToCheck.slice(0, 6).join(" | ")}. Pertinence à vérifier manuellement.`
    );

    const forms = [...document.querySelectorAll("form")].filter(isVisible);
    const requiredFields = fields.filter(el =>
      el.required ||
      el.getAttribute("aria-required") === "true" ||
      el.hasAttribute("required")
    );

    addRow(
      "FORMULAIRES",
      "S'assurer que les messages d'erreurs sont pertinents",
      "Soumettre les formulaires avec erreurs.",
      "Les messages d'erreurs sont présents, pertinents et identifient les champs en erreur.",
      forms.length === 0 ? "NA" : "À vérifier",
      forms.length === 0
        ? "Aucun formulaire détecté."
        : `${forms.length} formulaire(s), ${requiredFields.length} champ(s) obligatoire(s). Test manuel nécessaire.`
    );

    return result;
  });
}

(async () => {
  if (!fs.existsSync("screenshots")) fs.mkdirSync("screenshots");

  const browser = await chromium.launch({ headless: true });
  const workbook = new ExcelJS.Workbook();
  workbook.creator = "Bot audit accessibilité";
  workbook.created = new Date();

  const summaryData = [];

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
      { header: "Recommandation RGAA", key: "recommendation", width: 75 },
      { header: "URL de la page concernée", key: "url", width: 45 },
      { header: "Copie d'écran (si non conforme)", key: "screenshot", width: 35 },
      { header: "Nom et Date", key: "nameDate", width: 25 }
    ];

    const page = await browser.newPage({ viewport: { width: 1024, height: 768 } });

    try {
      console.log("Audit :", url);
      const audit = await auditSite(page, url);

      let screenshotPath = "";
      const hasKO = audit.rows.some(r => r.status === "KO");

      if (hasKO) {
        const safeName = url.replace(/https?:\/\//, "").replace(/[^\w]/g, "_");
        screenshotPath = `screenshots/${safeName}.jpg`;

        await page.screenshot({
          path: screenshotPath,
          type: "jpeg",
          quality: 45,
          fullPage: false
        });
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
          recommendation: getRecommendation(row),
          url: row.url,
          screenshot: row.status === "KO" ? screenshotPath : "",
          nameDate: `Audit auto - ${new Date().toLocaleDateString("fr-FR")}`
        });

        if (row.status === "KO" && screenshotPath && !imageInserted) {
          const imageId = workbook.addImage({
            filename: screenshotPath,
            extension: "jpeg"
          });

          ws.addImage(imageId, {
            tl: { col: 8, row: addedRow.number - 1 },
            ext: { width: 180, height: 100 }
          });

          ws.getRow(addedRow.number).height = 80;
          imageInserted = true;
        }
      }

      const totalTests = audit.rows.length;
      const koCount = audit.rows.filter(r => r.status === "KO").length;
      const okCount = audit.rows.filter(r => r.status === "OK").length;
      const checkCount = audit.rows.filter(r => r.status === "À vérifier").length;
      const naCount = audit.rows.filter(r => r.status === "NA").length;

      const score = totalTests > 0
        ? Math.round(((okCount + naCount) / totalTests) * 100)
        : 0;

      summaryData.push({
        site: url,
        score,
        totalTests,
        okCount,
        koCount,
        checkCount,
        naCount,
        topProblems: audit.rows.filter(r => r.status === "KO").map(r => `${r.theme} - ${r.ref}`).slice(0, 5).join(" | "),
        priorityRecommendations: audit.rows.filter(r => r.status === "KO").map(r => getRecommendation(r)).filter(Boolean).slice(0, 3).join(" | ")
      });

    } catch (error) {
      ws.addRow({
        theme: "ERREUR",
        ref: "Page inaccessible",
        test: "Chargement de la page",
        expected: "La page doit être accessible au bot.",
        status: "KO",
        comment: error.message,
        recommendation: "Vérifier que la page est accessible, que le domaine ne bloque pas le bot ou que le chargement ne nécessite pas une interaction manuelle.",
        url,
        screenshot: "",
        nameDate: `Audit auto - ${new Date().toLocaleDateString("fr-FR")}`
      });
    }

    await page.close();

    ws.getRow(1).font = { bold: true, color: { argb: "FFFFFFFF" } };
    ws.getRow(1).fill = { type: "pattern", pattern: "solid", fgColor: { argb: "FFFF7F00" } };
    ws.getRow(1).alignment = { vertical: "middle", horizontal: "center", wrapText: true };
    ws.autoFilter = { from: "A1", to: "J1" };
    ws.views = [{ state: "frozen", ySplit: 1 }];

    ws.eachRow((row, rowNumber) => {
      row.alignment = { vertical: "top", wrapText: true };
      if (rowNumber > 1 && row.height < 80) row.height = 55;

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

  const summarySheet = workbook.addWorksheet("Résumé global");

  summarySheet.columns = [
    { header: "Site", key: "site", width: 45 },
    { header: "Score global", key: "score", width: 18 },
    { header: "Tests total", key: "totalTests", width: 15 },
    { header: "OK", key: "okCount", width: 10 },
    { header: "KO", key: "koCount", width: 10 },
    { header: "À vérifier", key: "checkCount", width: 15 },
    { header: "NA", key: "naCount", width: 10 },
    { header: "Top problèmes", key: "topProblems", width: 70 },
    { header: "Recommandations prioritaires", key: "priorityRecommendations", width: 90 }
  ];

  summaryData
    .sort((a, b) => b.koCount - a.koCount || a.score - b.score)
    .forEach(item => {
      summarySheet.addRow({
        site: item.site,
        score: `${item.score}%`,
        totalTests: item.totalTests,
        okCount: item.okCount,
        koCount: item.koCount,
        checkCount: item.checkCount,
        naCount: item.naCount,
        topProblems: item.topProblems || "Aucun KO détecté",
        priorityRecommendations: item.priorityRecommendations || "Aucune recommandation prioritaire"
      });
    });

  summarySheet.getRow(1).font = { bold: true, color: { argb: "FFFFFFFF" } };
  summarySheet.getRow(1).fill = { type: "pattern", pattern: "solid", fgColor: { argb: "FFFF7F00" } };
  summarySheet.autoFilter = { from: "A1", to: "I1" };
  summarySheet.views = [{ state: "frozen", ySplit: 1 }];

  await browser.close();

  await workbook.xlsx.writeFile("audit-accessibilite-rgaa.xlsx");
  console.log("✅ Audit terminé : audit-accessibilite-rgaa.xlsx");
})();