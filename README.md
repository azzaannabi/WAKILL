# Wakil : serveur complet

Agent de vente IA pour Instagram, Facebook et WhatsApp, en derja. Tableau de bord mobile inclus, sans dépendance à installer (Node 20 ou plus).

## Ce que fait le serveur
- Reçoit les messages (webhook Meta), répond avec Claude, prend les commandes (outil `creer_commande`, montant calculé par le code).
- **Mode brouillon** (par défaut) : l'agent propose, tu relis et tu envoies. Passe en **automatique** dans Agent, Boutique.
- Commande + 20 h : demande de confirmation. « Oui » : confirmée. Pas de réponse à 24 h : « Sans réponse » et alerte, sans relance.
- Réclamation ou question hors périmètre : l'agent se met en pause, tu reçois une alerte WhatsApp avec le lien direct vers la conversation.
- Changement de statut dans le tableau de bord (Emballée, En livraison, Livrée) : message automatique au client si l'option est cochée.
- Chaque soir : fichier Excel Intigo (11 colonnes) envoyé sur ton WhatsApp. Téléchargeable aussi à la main.

## Tester en local, sans aucun compte
```
node test/e2e.js
```
Lance un faux Claude, un faux Meta et un faux WhatsApp, et vérifie 31 points (webhook signé, brouillons, commande, réclamation, minuteurs, statuts, export Excel).

## Déployer
1. Héberge sur Railway, Render ou un VPS. Commande de démarrage : `node server.js`.
2. Ajoute un **volume persistant** et mets `DATA_DIR` dessus, sinon tes données disparaissent à chaque redéploiement.
3. La seule variable obligatoire côté hébergeur est `DASH_PASSWORD` (protège le tableau de bord). Tout le reste — clé Claude, Instagram, Facebook, WhatsApp — se colle **depuis l'appli**, dans Agent → Intégrations, pas besoin de retoucher l'hébergeur.
4. Ouvre l'URL de ton service sur ton téléphone, saisis le mot de passe. L'URL publique et l'adresse du webhook s'affichent automatiquement dans Intégrations.

## Brancher Instagram et Facebook
1. Compte Instagram professionnel relié à une page Facebook.
2. Meta for Developers : crée une app, ajoute la messagerie, génère le jeton de page.
3. Dans l'appli, Agent → Intégrations → Instagram et Facebook : copie l'**URL du webhook** affichée là, colle-la dans Meta avec un jeton de vérification de ton choix (le même des deux côtés), abonne « messages ». Colle le jeton de page et le secret de l'app dans les champs correspondants, puis Enregistrer.
4. Passe l'app en mode live après validation Meta. Les noms exacts des permissions changent : suis la documentation Meta du moment.

## Brancher WhatsApp
1. WhatsApp Cloud API : numéro dédié. Dans Agent → Intégrations → WhatsApp, colle le jeton, l'ID du numéro et ton propre numéro WhatsApp (pour recevoir les alertes), puis Enregistrer.
2. Hors des 24 h, WhatsApp n'accepte que des **modèles approuvés** (catégorie Utilitaire). Ceux-ci restent à configurer en variables d'environnement (`WA_TEMPLATES` etc., voir `.env.example`), variables dans cet ordre :
   - `confirmation_demande` : {{1}} prénom, {{2}} boutique, {{3}} produit, {{4}} montant
   - `commande_confirmee` : {{1}} prénom, {{2}} numéro
   - `commande_emballee` : {{1}} numéro
   - `commande_en_livraison` : {{1}} numéro, {{2}} date estimée, {{3}} montant
   - `commande_livree` : {{1}} prénom, {{2}} boutique
   - `alerte_wakil` : {{1}} texte de l'alerte
   - `export_quotidien` : en-tête document, {{1}} légende
   Reprends les textes de l'onglet Agent, Messages, puis renseigne `WA_TEMPLATES`, `WA_ALERT_TEMPLATE` et `WA_EXPORT_TEMPLATE`.

## Premier jour conseillé
1. Mode brouillon, statuts automatiques décochés.
2. Envoie-toi des messages de test depuis un autre compte Instagram.
3. Quand tu valides presque tout sans corriger, passe en automatique.

## Limites à connaître
- **TikTok** : non branché. TikTok a une vraie Business Messaging API, mais l'accès demande une candidature approuvée par TikTok (compte Business + Business Center + app validée sur business-api.tiktok.com), qui prend plusieurs jours. Une fois approuvé, envoie les identifiants pour que le webhook soit branché.
- **Notes vocales et images** : l'agent ne les lit pas, il demande poliment d'écrire.
- **Intigo et Converty** : pas d'envoi automatique des colis. Tu utilises le fichier Excel ; l'API se branche quand leur documentation est disponible.
- L'envoi Instagram et Facebook par DM n'est possible que dans les 24 h après le dernier message du client ; au-delà, le serveur passe par WhatsApp avec le numéro de la commande.
- Non testé face aux vraies API Meta et WhatsApp (pas d'accès ici) : fais le premier test en mode brouillon.
- Les données sont dans `data/state.json` : sauvegarde-le régulièrement.
