const audit = await page.evaluate(() => {

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

  function parseRgb(str) {
    const m = String(str || "").match(/rgba?\((\d+),\s*(\d+),\s*(\d+)/);
    return m ? [Number(m[1]), Number(m[2]), Number(m[3])] : null;
  }

  function isVisible(el) {
    const s = getComputedStyle(el);
    const r = el.getBoundingClientRect();
    return s.display !== "none" && s.visibility !== "hidden" && r.width > 0 && r.height > 0;
  }

  function getBg(el) {
    let current = el;
    while (current && current !== document.documentElement) {
      const bg = getComputedStyle(current).backgroundColor;
      if (bg && bg !== "transparent" && !bg.includes("rgba(0,0,0,0)")) return bg;
      current = current.parentElement;
    }
    return "rgb(255,255,255)";
  }

  // 🔴 CONTRASTE
  const issues = [];

  document.querySelectorAll("body *").forEach(el => {
    if (!isVisible(el)) return;

    const text = (el.innerText || "").trim();
    if (text.length < 3) return;

    const style = getComputedStyle(el);
    const fg = parseRgb(style.color);
    const bg = parseRgb(getBg(el));

    if (!fg || !bg) return;

    const ratio = contrastRatio(fg, bg);

    const fontSize = parseFloat(style.fontSize);
    const fontWeight = parseInt(style.fontWeight) || 400;
    const large = fontSize >= 24 || (fontSize >= 18.5 && fontWeight >= 700);
    const min = large ? 3 : 4.5;

    if (ratio < min) {
      issues.push(`${text.slice(0, 50)} (${ratio.toFixed(2)} < ${min})`);
    }
  });

  return {
    contrast: issues.length === 0 ? "OK" : "KO",
    comment: issues.slice(0, 5).join(" | ")
  };

});
