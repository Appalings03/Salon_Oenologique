<?php
  // Récupérer les données envoyées par JavaScript
  $donnees = json_decode(file_get_contents('php://input'), true);

  // Récupérer les données individuelles
  $nom = $donnees['nom'];
  $prenom = $donnees['prenom'];
  $email = $donnees['email'];
  $date = $donnees['date'];

  // Ouvrir le fichier Excel en mode écriture
  $fichier = fopen('donnees.xlsx', 'a');

  // Écrire les données dans le fichier Excel
  fputcsv($fichier, array($nom, $prenom, $email, $date));

  // Fermer le fichier Excel
  fclose($fichier);
?>






