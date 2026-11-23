# Caisse du stand

Caisse des marchés et foires de la Ferme du Lacas : ventes au kilo ou à la pièce, remises, retours par lot, export Excel.

- Fonctionne **sans réseau** : tout est gardé sur le téléphone.
- Dès qu'il y a internet, chaque téléphone envoie ses ventes et retours dans un **tableau Google Sheets** et récupère ceux des autres.
- Scan des étiquettes avec l'**appareil photo** (bouton « Scanner ») ou avec une douchette Bluetooth.

Adresse : https://theoviatge.github.io/caisse-du-stand/

## Mise en place du tableau Google Sheets (une seule fois)

1. Crée un tableau Google Sheets vide, par exemple « Caisse du stand ».
2. Menu **Extensions > Apps Script**. Efface le contenu et colle celui de [`apps-script/Code.gs`](apps-script/Code.gs).
3. En haut du script, remplace `a-changer` par ton code secret, puis enregistre.
4. **Déployer > Nouveau déploiement** > type **Application Web**, « Exécuter en tant que : **moi** », « Qui a accès : **Tout le monde** ». Autorise l'accès demandé.
5. Copie l'adresse qui finit par `/exec`.

## Relier un téléphone

Ouvre sur le téléphone le lien de réglage :

```
https://theoviatge.github.io/caisse-du-stand/#synchro=ADRESSE_EXEC&code=TON_CODE
```

ou touche « Synchro à régler » en haut de la caisse et colle l'adresse et le code.
Ensuite, « Ajouter à l'écran d'accueil » pour l'avoir comme une appli.

Le tableau se remplit tout seul : onglets **Ventes** (une ligne par article), **Tickets**, **Retours**, **Produits**.
L'onglet caché `_sync` sert à la synchronisation : ne pas le modifier.

## Sauvegarde

Dans « Synchronisation » : **Télécharger une sauvegarde** (fichier .json) et **Importer une sauvegarde** (ajoute sans rien effacer).
