# client-file-manager

Mini-projet Python pour assurer le suivi des dossiers clients d'une entreprise.



But : Repérer les dossiers qui n'ont pas été traités depuis un certain temps afin :

			        - d'éviter d'oublier un dossier
				  
			        - effectuer des rappels au bon moment

Moyens :        		
				
				- Se connecter à l'adresse email de l'entreprise (via la bibliothèque Selenium)

			        - Parcourir chaque email envoyé sur une période donnée, et récupérer le numéro de dossier associé
				  
			        - Mettre à jour dans un tableau Excel la date de traitement du dossier si elle est plus récente que la précédente

Contraintes :   		

				- Effectuer ses correspondances par email

				- Indiquer, dans l'objet de chaque email, le numéro du dossier en question

Résultat : Production d'un fichier Excel contenant tous les dossiers triés par dernière date de traitement, et contenant pour chaque dossier : 

			          - Son numéro
				  
			          - Sa dernière date de traitement
				  
			          - L'objet et le destinataire de l'email envoyé à cette date

L'application a une interface minimaliste codée avec Tkinter, l'objectif étant surtout le fichier Excel produit.
