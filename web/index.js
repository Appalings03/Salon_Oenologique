function envoyerFormulaire(event) {
    event.preventDefault();
  
    // Récupérer les données saisies par l'utilisateur
    const nom = document.getElementById('nom').value;
    const prenom = document.getElementById('prenom').value;
    const email = document.getElementById('email').value;
    const date = new Date().toISOString();
  
    // Créer un objet avec les données saisies
    const donnees = {
      nom: nom,
      prenom: prenom,
      email: email,
      date: date
    };
  
    // Envoyer les données au serveur en utilisant fetch
    fetch('enregistrer_donnees.php', {
      method: 'POST',
      body: JSON.stringify(donnees)
    }).then(function(response) {
      // Rediriger l'utilisateur vers une page de confirmation
      window.location.href = 'confirmation.html';
    });
}
  
// Attacher un événement à la soumission du formulaire
document.getElementById('monFormulaire').addEventListener('submit', envoyerFormulaire);
  