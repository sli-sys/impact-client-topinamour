# Impact clients Topinamour

Outil web local pour importer un export Pennylane des factures, calculer l’impact anti-gaspi de chaque client (mêmes ratios que le calculateur Topinamour) et générer un certificat PDF ou PNG.

Les calculs restent dans le navigateur : aucun fichier n’est envoyé sur un serveur.

## Lancer le projet en local

C’est un site statique (`index.html`, `app.js`, `styles.css`). Un petit serveur HTTP est recommandé, plutôt que d’ouvrir le fichier directement dans le navigateur.

### 1. Récupérer le projet

```bash
git clone https://github.com/sli-sys/impact-client-topinamour.git
cd impact-client-topinamour
```

### 2. Démarrer un serveur

**Python 3** (Windows : `py`, macOS / Linux : `python3`) :

```bash
py -3 -m http.server 8765
```

ou

```bash
python3 -m http.server 8765
```

**Node.js** (alternative) :

```bash
npx --yes serve -p 8765
```

### 3. Ouvrir l’application

Dans le navigateur : [http://127.0.0.1:8765/](http://127.0.0.1:8765/)

Pour arrêter le serveur : `Ctrl + C` dans le terminal.

## Utilisation

1. Déposez (ou choisissez) l’export Pennylane au format `.xlsx`, `.xls` ou `.csv`.
2. L’outil lit la feuille **Lignes de facture** (client, produit, quantité, SIREN).
3. Consultez le tableau d’impact, exportez-le en Excel si besoin.
4. Cliquez sur **Voir le certificat** pour un client, puis téléchargez le PDF ou l’image.

Le certificat dépend des kilogrammes sauvés :

- **Sauveur débutant** : moins de 500 kg
- **Sauveur confirmé** : de 500 kg à moins de 1 500 kg
- **Héros anti-gaspi** : 1 500 kg ou plus
