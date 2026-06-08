const { chromium } = require("playwright");
const ExcelJS = require("exceljs");
const fs = require("fs");

const urls = fs.readFileSync("urls.txt", "utf8")
  .split(/\r?\n/)
  .map(u => u.trim())
  .filter(Boolean);

async function auditSite(page, url) {
  await page.goto(url, { waitUntil: "domcontentloaded", timeout: 45000 });
  await page.waitForTimeout(2500);

  return await page.evaluate(() => {
    function isVisible(el) {
      const style = getComputedStyle(el);
      const rect = el.getBoundingClientRect();
      return style.display !== "none" &&
        style.visibility !== "hidden" &&
        rect.width > 0 &&
        rect.height > 0;
    }

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

    const result = {
      rows: []
    };

    const pageLang = document.documentElement.getAttribute("lang") || "";

    addRow(
      "CONTENU TEXTUEL",
      "Donner un titre aux pages",
      "Lancer l’inspecteur de code du navigateur et examiner le titre de page (<title>).",
      "Chaque page possède un titre unique et descriptif du contenu.",
      document.title && document.title.trim().length > 5 ? "OK" : "KO",
      document.title ? `Titre détecté : ${document.title}` : "Aucun titre de page détecté."
    );

    const headings = [...document.querySelectorAll("h1,h2,h3,h4,h5,h6")].filter(isVisible);

    addRow(
      "CONTENU TEXTUEL",
      "Donner un titre aux rubriques",
      "Installer le bookmarklet Headings et vérifier les titres.",
      "Tous les contenus traités visuellement comme des titres possèdent une sémantique de titre.",
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
      "Installer le bookmarklet Headings et vérifier la hiérarchie des titres.",
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
      "Lancer l’inspecteur de code du navigateur. Examiner l’élément <html>.",
      "Un attribut lang est présent dans l’élément <html> de la page.",
      pageLang ? "OK" : "KO",
      pageLang ? `Lang détecté : ${pageLang}` : "Attribut lang absent."
    );

    addRow(
      "CONTENU TEXTUEL",
      "Indiquer la langue principale",
      "Lancer l’inspecteur de code du navigateur. Examiner l’élément <html>.",
      "La valeur de l’attribut lang correspond à la langue principale du document.",
      pageLang ? "À vérifier" : "KO",
      pageLang ? `Valeur détectée : ${pageLang}. À confirmer manuellement.` : "Langue principale non déclarée."
    );

    const images = [...document.images].filter(isVisible);
    const imageLinks = images.filter(img => img.closest("a"));
    const missingAlt = images.filter(img => !img.hasAttribute("alt") && img.getAttribute("role") !== "presentation");
    const emptyAlt = images.filter(img => img.hasAttribute("alt") && img.getAttribute("alt") === "");

    addRow(
      "CONTENU NON TEXTUEL",
      "S’assurer que les images ont une alternative textuelle",
      "Installer puis lancer le bookmarklet List Images ou l’inspecteur de code.",
      "Image lien : le contenu de l’attribut alt de chaque image-lien est pertinent par rapport à la cible du lien.",
      imageLinks.length === 0 ? "NA" : imageLinks.every(img => img.hasAttribute("alt")) ? "OK" : "KO",
      imageLinks.length === 0
        ? "Aucune image-lien détectée."
        : `${imageLinks.length} image(s)-lien détectée(s), ${imageLinks.filter(img => !img.hasAttribute("alt")).length} sans alt.`
    );

    addRow(
      "CONTENU NON TEXTUEL",
      "S’assurer que les images ont une alternative textuelle",
      "Installer puis lancer le bookmarklet List Images ou l’inspecteur de code.",
      "Image porteuse d’information : l’attribut alt de chaque image est pertinent.",
      missingAlt.length === 0 ? "OK" : "KO",
      missingAlt.length === 0
        ? `${images.length} image(s) analysée(s), aucune image sans alt détectée.`
        : `${missingAlt.length} image(s) sans attribut alt.`
    );

    addRow(
      "CONTENU NON TEXTUEL",
      "S’assurer que les images ont une alternative textuelle",
      "Installer puis lancer le bookmarklet List Images ou l’inspecteur de code.",
      "Image décorative : l’attribut alt est présent mais vide.",
      emptyAlt.length > 0 ? "OK" : "NA",
      emptyAlt.length > 0
        ? `${emptyAlt.length} image(s) avec alt vide détectée(s).`
        : "Aucune image décorative avec alt vide détectée."
    );

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
      textIssues.length === 0
        ? "Aucun écart de contraste détecté automatiquement."
        : `Contrastes insuffisants : ${textIssues.slice(0, 5).join(" | ")}`
    );

    const colorStatusElements = [...document.querySelectorAll("[class*='error'],[class*='success'],[class*='warning'],[class*='danger'],[class*='alert'],[class*='status'],[class*='badge'],[class*='tag']")].filter(isVisible);

    addRow(
      "COULEURS ET CONTRASTE",
      "S’assurer que l’information n’est pas transmise uniquement par la couleur",
      "Installer et lancer Color Contrast Analyser.",
      "L’information transmise par la couleur peut également être obtenue par un texte explicite.",
      colorStatusElements.length === 0 ? "NA" : "À vérifier",
      colorStatusElements.length === 0
        ? "Aucun élément de statut/couleur détecté."
        : `${colorStatusElements.length} élément(s) de statut/couleur détecté(s). Vérification manuelle nécessaire.`
    );

    addRow(
      "NAVIGATION GÉNÉRALE",
      "Permettre le contrôle des animations",
      "Identifier tout contenu en mouvement, mis à jour automatiquement, clignotant ou en défilement.",
      "L’utilisateur peut mettre en pause ou masquer les animations, mouvements, mises à jour ou clignotements.",
      "À vérifier",
      "Test manuel nécessaire : carrousel, slider, vidéo, animation ou contenu dynamique."
    );

    const interactive = [...document.querySelectorAll("a, button, input, select, textarea, [tabindex], [role='button'], [role='link']")].filter(isVisible);
    const keyboardIssues = interactive.filter(el => {
      const tag = el.tagName.toLowerCase();
      const tabindex = el.getAttribute("tabindex");
      return tabindex === "-1" && ["a", "button", "input", "select", "textarea"].includes(tag);
    });

    addRow(
      "NAVIGATION CLAVIER",
      "Permettre l’utilisation de l’application au clavier",
      "Parcourir la page au clavier à l’aide des touches Tab ou Shift + Tab.",
      "Tous les éléments interactifs sont atteignables en naviguant au clavier.",
      keyboardIssues.length === 0 ? "À vérifier" : "KO",
      keyboardIssues.length === 0
        ? `${interactive.length} élément(s) interactif(s) détecté(s). Test clavier manuel nécessaire.`
        : `${keyboardIssues.length} élément(s) semblent exclus de la navigation clavier.`
    );

    addRow(
      "NAVIGATION CLAVIER",
      "Permettre l’utilisation de l’application au clavier",
      "Utiliser les éléments interactifs au clavier avec Entrée, Espace et les flèches directionnelles.",
      "Tous les éléments interactifs sont utilisables depuis des interactions clavier.",
      "À vérifier",
      "Doublon du test 22 / test manuel nécessaire."
    );

    addRow(
      "MISE EN PAGE",
      "Utiliser des tailles relatives et faire du web adaptatif",
      "Avec Firefox, sélectionner Zoom puis Agrandir uniquement le texte et activer un zoom à 200%.",
      "Absence de contenus tronqués ou masqués et absence de fonctionnalités inutilisables.",
      "À vérifier",
      "Test manuel nécessaire à 200%."
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
      "S’assurer qu’un nom accessible est associé à chaque champ de formulaire",
      "Utiliser l’inspecteur de code du navigateur sur l’onglet Accessibilité.",
      "Chaque champ a au moins un nom accessible pertinent et contient au moins le texte de l’étiquette de champ visible à l’écran.",
      fields.length === 0 ? "NA" : fieldsWithoutName.length === 0 && fieldsOnlyPlaceholder.length === 0 ? "OK" : "KO",
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
      "S’assurer qu’un nom accessible est associé à chaque champ de formulaire",
      "Vérifier la pertinence des étiquettes associées aux champs.",
      "Chaque étiquette associée à un champ de formulaire est pertinente.",
      fields.length === 0 ? "NA" : "À vérifier",
      fields.length === 0
        ? "Aucun champ de formulaire détecté."
        : `Labels détectés : ${labelsToCheck.slice(0, 6).join(" | ")}. Pertinence à vérifier manuellement.`
    );

    const forms = [...document.querySelectorAll("form")].filter(isVisible);
    const requiredFields = fields.filter(el =>
      el.required ||
      el.getAttribute("aria-required") === "true" ||
      el.hasAttribute("required")
    );

    addRow(
      "FORMULAIRES",
      "S’assurer que les messages d’erreurs sont pertinents",
      "Renseigner les formulaires avec des données erronées et des champs obligatoires laissés vides.",
      "Les messages d’erreurs sont présents, pertinents, et identifient les champs en erreur.",
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

  for (const url of urls) {
    const sheetName = new URL(url).hostname.replace("www.", "").substring(0, 31);
    let ws = workbook.getWorksheet(sheetName);

    if (ws) {
      ws = workbook.addWorksheet(sheetName.substring(0, 25) + "_" + Math.floor(Math.random() * 999));
    } else {
      ws = workbook.addWorksheet(sheetName);
    }

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
      console.log("Audit :", url);

      const audit = await auditSite(page, url);
      const hasKO = audit.rows.some(r => r.status === "KO");

      let screenshotPath = "";

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
            tl: { col: 7, row: addedRow.number - 1 },
            ext: { width: 220, height: 120 }
          });

          ws.getRow(addedRow.number).height = 95;
          imageInserted = true;
        }
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

    ws.getRow(1).font = {
      bold: true,
      color: { argb: "FF000000" }
    };

    ws.getRow(1).fill = {
      type: "pattern",
      pattern: "solid",
      fgColor: { argb: "FFFF7F00" }
    };

    ws.getRow(1).alignment = {
      vertical: "middle",
      horizontal: "center",
      wrapText: true
    };

    ws.getRow(1).height = 28;
    ws.autoFilter = { from: "A1", to: "I1" };
    ws.views = [{ state: "frozen", ySplit: 1 }];

    ws.eachRow((row, rowNumber) => {
      row.alignment = {
        vertical: "top",
        wrapText: true
      };

      if (rowNumber > 1 && row.height < 55) {
        row.height = 55;
      }

      const themeCell = row.getCell(1);
      const refCell = row.getCell(2);
      const statusCell = row.getCell(5);
      const commentCell = row.getCell(6);

      if (rowNumber > 1) {
        const theme = String(themeCell.value || "");

        if (
          theme.includes("COULEURS ET CONTRASTE") ||
          theme.includes("NAVIGATION") ||
          theme.includes("FORMULAIRES")
        ) {
          themeCell.fill = {
            type: "pattern",
            pattern: "solid",
            fgColor: { argb: "FFFFFF00" }
          };

          refCell.fill = {
            type: "pattern",
            pattern: "solid",
            fgColor: { argb: "FFFFFF00" }
          };
        }

        if (statusCell.value === "OK") {
          statusCell.fill = {
            type: "pattern",
            pattern: "solid",
            fgColor: { argb: "FF92D050" }
          };
        } else if (statusCell.value === "KO") {
          statusCell.fill = {
            type: "pattern",
            pattern: "solid",
            fgColor: { argb: "FFFF0000" }
          };
        } else if (statusCell.value === "NA") {
          statusCell.fill = {
            type: "pattern",
            pattern: "solid",
            fgColor: { argb: "FFD9D9D9" }
          };
        } else if (statusCell.value === "À vérifier") {
          statusCell.fill = {
            type: "pattern",
            pattern: "solid",
            fgColor: { argb: "FFFFD966" }
          };
        }

        if (String(commentCell.value || "").includes("Doublon")) {
          commentCell.fill = {
            type: "pattern",
            pattern: "solid",
            fgColor: { argb: "FFFF0000" }
          };

          commentCell.font = {
            bold: true,
            color: { argb: "FF000000" }
          };
        }
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
