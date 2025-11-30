/**
 * @fileoverview Script de gestion d'envoi de résumé d'agenda hebdomadaire.
 * @author Fabrice Faucheux
 */

// Constante globale pour l'email de l'utilisateur courant
const EMAIL_DESTINATAIRE = Session.getActiveUser().getEmail();

/**
 * Récupère les événements du calendrier pour les 7 prochains jours et envoie
 * un récapitulatif formaté par courriel.
 * * @return {void}
 */
function envoyerResumeHebdomadaire() {
  const aujourdhui = new Date();
  // Calcul de la date de fin (aujourd'hui + 7 jours)
  const dansUneSemaine = new Date(aujourdhui.getTime() + 7 * 24 * 60 * 60 * 1000);
  const optionsDate = { weekday: 'long', day: 'numeric', month: 'long', hour: '2-digit', minute: '2-digit' };

  try {
    const calendrier = CalendarApp.getDefaultCalendar();
    const evenements = calendrier.getEvents(aujourdhui, dansUneSemaine);

    const dateDebutStr = aujourdhui.toLocaleDateString('fr-FR');
    const dateFinStr = dansUneSemaine.toLocaleDateString('fr-FR');
    const sujet = `📅 Votre agenda du ${dateDebutStr} au ${dateFinStr}`;

    // Début de la construction HTML
    let contenuPrincipalHtml = '';

    if (evenements.length === 0) {
      contenuPrincipalHtml = `
        <div style="padding: 20px; background-color: #e8f0fe; border-radius: 5px; color: #1967d2;">
          <p>Aucun événement prévu cette semaine. Profitez de votre temps libre ! 🏖️</p>
        </div>`;
    } else {
      // Utilisation de .map() pour transformer les événements en lignes HTML (Approche ES6 moderne)
      const lignesTableau = evenements.map(evenement => {
        const titre = evenement.getTitle() || "Sans titre";
        
        let horaire;
        if (evenement.isAllDayEvent()) {
          horaire = "Toute la journée";
        } else {
          const debut = evenement.getStartTime().toLocaleTimeString('fr-FR', { hour: '2-digit', minute: '2-digit' });
          const fin = evenement.getEndTime().toLocaleTimeString('fr-FR', { hour: '2-digit', minute: '2-digit' });
          const jour = evenement.getStartTime().toLocaleDateString('fr-FR', { weekday: 'short', day: 'numeric' });
          horaire = `<strong>${jour}</strong> : ${debut} - ${fin}`;
        }

        return `
          <tr>
            <td style="padding: 12px; border-bottom: 1px solid #eee; color: #333;">${titre}</td>
            <td style="padding: 12px; border-bottom: 1px solid #eee; color: #555;">${horaire}</td>
          </tr>
        `;
      }).join(''); // Rejoint tous les éléments du tableau en une seule chaîne

      contenuPrincipalHtml = `
        <table style="width: 100%; border-collapse: collapse; margin-top: 15px;">
          <thead>
            <tr style="background-color: #f8f9fa; text-align: left;">
              <th style="padding: 12px; border-bottom: 2px solid #ddd; color: #5f6368;">Événement</th>
              <th style="padding: 12px; border-bottom: 2px solid #ddd; color: #5f6368;">Horaire</th>
            </tr>
          </thead>
          <tbody>
            ${lignesTableau}
          </tbody>
        </table>
      `;
    }

    // Assemblage final du corps HTML
    const corpsHtml = `
      <body style="font-family: 'Segoe UI', Arial, sans-serif; line-height: 1.6; color: #202124; max-width: 600px; margin: 0 auto;">
        <h2 style="color: #1a73e8; border-bottom: 1px solid #e0e0e0; padding-bottom: 10px;">Votre résumé hebdomadaire</h2>
        <p>Bonjour,</p>
        <p>Voici un aperçu de vos engagements à venir pour la semaine :</p>
        ${contenuPrincipalHtml}
        <p style="font-size: 11px; color: #9aa0a6; margin-top: 30px; border-top: 1px solid #e0e0e0; padding-top: 10px;">
          Généré automatiquement par Google Apps Script - Auteur : Fabrice Faucheux.
        </p>
      </body>
    `;

    // Envoi de l'email
    GmailApp.sendEmail(EMAIL_DESTINATAIRE, sujet, "Votre client mail ne supporte pas le HTML.", { htmlBody: corpsHtml });
    console.log(`E-mail de résumé envoyé avec succès à ${EMAIL_DESTINATAIRE}`);

  } catch (erreur) {
    console.error(`Erreur critique lors de l'envoi du résumé : ${erreur.stack}`);
  }
}

/**
 * Configure un déclencheur temporel pour exécuter le rapport automatiquement.
 * Nettoie les anciens déclencheurs similaires pour éviter les doublons.
 * * @return {void}
 */
function creerDeclencheurHebdomadaire() {
  try {
    const nomFonction = 'envoyerResumeHebdomadaire';
    const declencheursActuels = ScriptApp.getProjectTriggers();
    
    // Suppression préventive des déclencheurs existants pour cette fonction
    declencheursActuels
      .filter(d => d.getHandlerFunction() === nomFonction)
      .forEach(d => ScriptApp.deleteTrigger(d));

    // Création du nouveau déclencheur
    ScriptApp.newTrigger(nomFonction)
      .timeBased()
      .onWeekDay(ScriptApp.WeekDay.MONDAY)
      .atHour(8)
      .inTimezone(Session.getScriptTimeZone())
      .create();

    console.log(`Déclencheur planifié : ${nomFonction} s'exécutera chaque lundi à 08h00.`);
  } catch (erreur) {
    console.error(`Impossible de créer le déclencheur : ${erreur.message}`);
  }
}
