# RP_enquete_analyse
Mars 2023

Auteur : Kenneth Maussang

Code Python utilisés pour analyser les résultats d'une enquête en ligne réalisée avec la solution logiciel LimeSurvey

Analyse_Enquete.py : fichier principal. Le fichier Excel correspondant à l'export de LimeSurvey doit être dans le même dossier que ce fichier Python. Dans cette version, le fichier doit être nommé "results.xlsx". Ce nom de fichier peut être modifé à la ligne 35 du code. Nécessite les modules ExtractData.py et Name_cleaning.py

ExtractData.py : module de fonctions pour extraire les données du fichier Excel.

Name_cleaning.py : module permettant d'harmoniser les réponses (noms d'établissements, noms de villes).

ExtractAnswer.py : code permettant d'extraire l'ensemble des réponses individuelles au format .docx. Nécessite le modules ExtractData.py.

ExtractQuestion.py : code permettant d'extraire l'ensemble des réponsesà une question au format .docx. Nécessite le modules ExtractData.py.
