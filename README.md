# BotBB99 Audit Accessibilité RGAA / WCAG

Ce bot audite automatiquement une liste de sites et génère un fichier Excel similaire au modèle RGAA.

## Tests inclus

- Titre de page `<title>`
- Titres H1-H6 / hiérarchie headings
- Langue principale `<html lang="">`
- Images avec attribut `alt`
- Contraste texte/fond WCAG AA
- Information transmise uniquement par la couleur
- Double codage visuel
- Liens distinguables

## Installation dans GitHub Codespaces

```bash
npm install
npx playwright install chromium
npx playwright install --with-deps chromium
npm start
```

Le fichier généré sera :

```text
audit-accessibilite-rgaa.xlsx
```

## Modifier la liste des sites

Modifie le fichier `urls.txt`.

