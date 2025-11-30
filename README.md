# Résumé d'Agenda Hebdomadaire Automatisé

## Description
Ce projet Google Apps Script récupère automatiquement les événements de votre calendrier Google principal pour les 7 prochains jours et vous envoie un résumé formaté et lisible par e-mail. Idéal pour préparer sa semaine le lundi matin.

## Auteur
**Fabrice Faucheux** Expert Google Apps Script

## Fonctionnalités Clés
* **Extraction dynamique** : Récupère les événements via l'API CalendarApp (fenêtre glissante de 7 jours).
* **Mise en forme HTML** : Génère un tableau propre et responsive pour une lecture facile sur mobile ou bureau.
* **Gestion des horaires** : Distingue automatiquement les événements ponctuels des événements "Toute la journée".
* **Automatisation** : Inclut un script de configuration pour un envoi automatique chaque lundi à 8h00.

## Installation Manuelle

1.  Ouvrez **Google Drive** et créez un nouveau projet Apps Script.
2.  Copiez le contenu du fichier `Code.js` fourni dans l'éditeur.
3.  Enregistrez le projet (Ctrl + S).
4.  Exécutez manuellement la fonction `creerDeclencheurHebdomadaire` **une seule fois**.
5.  Acceptez les demandes d'autorisation (Calendar, Gmail, ScriptApp) lors de la première exécution.

## Prérequis
* Un compte Google Workspace ou Gmail actif.
* L'accès aux services : Gmail et Calendar.
