"""
Définition de dictionnaires pour le nettoyages de certains champs libres notamment le nom de l'établissement et la ville
version : 1.0
Python 3.10
Auteur : K. Maussang
"""
def get_dict_etab():
    dict_etab = {}
    ##########################
    ####### NON PUBLIC #######
    ##########################
    dict_etab['XXXXX']='XXxxxXXX'
    return dict_etab

def get_dict_ville():
    dict_ville = {}
    ##########################
    ####### NON PUBLIC #######
    ##########################
    dict_ville['XXXXXX'] = 'XXXxxxxxXXXX'
    return dict_ville

def get_dict_CNU():
    dict_CNU = {}
    dict_CNU["Droit privé, droit public, sciences politiques, histoire du droit (groupe 1 du CNU)"]="CNU-1"
    dict_CNU["Sciences économiques, sciences de gestion et du management (groupe 2 du CNU)"]="CNU-2"
    dict_CNU["Littérature, langues, sciences du langage, linguistique (groupe 3 du CNU)"]="CNU-3"
    dict_CNU["Psychologie, sociologie, histoire, géographie, urbanisme, ethnologie, philosophie, architecture (groupe 4 du CNU)"]="CNU-4"
    dict_CNU["Mathématiques, informatique (groupe 5 du CNU)"]="CNU-5"
    dict_CNU["Physique hors astrophysique/astronomie (groupe 6 du CNU)"]="CNU-6"
    dict_CNU["Chimie (groupe 7 du CNU)"]="CNU-7"
    dict_CNU["Astronomie/astrophysique, géologie, météorologie (groupe 8 du CNU)"]="CNU-8"
    dict_CNU["Mécanique, génie informatique, automatique, traitement du signal, génie électrique, électronique, photoniques (groupe 9 du CNU)"]="CNU-9"
    dict_CNU["Biologie, biochimie, physiologie, neurosciences, écologie (groupe 10 du CNU)"]="CNU-10"
    dict_CNU["Pluridisciplinaire (groupe 12 du CNU, sections 70, 71, 72, 73 et 74)"]="CNU-12"
    # cas des entrées "autres"
    dict_CNU["Sciences de l'information et de la communication"]="CNU-12"
    dict_CNU["Microélectronique"]="CNU-9"
    dict_CNU["Anthropologie"]="CNU-4"
    dict_CNU["archéologie"] = "CNU-4"
    dict_CNU["Archéologie"] = "CNU-4"
    dict_CNU["Microbiologie"] = "CNU-10"
    dict_CNU["Aménagement"] = "CNU-4"
    dict_CNU["24"] = "CNU-4"
    dict_CNU["24 "] = "CNU-4"
    dict_CNU[" 24"] = "CNU-4"
    dict_CNU["Santé publique, Epidémiologie"] = "Santé"
    dict_CNU["groupe 10 du CNU"]="CNU-10"
    dict_CNU["Hydrologie"] = "CNU-8"
    dict_CNU["Sociologie des Sciences et des techniques"] = "CNU-12"
    return dict_CNU

def get_dict_CNU_autres():
    dict_CNU_autres = {}
    dict_CNU_autres["Sciences de l'information et de la communication"]="CNU-12"
    dict_CNU_autres["Microélectronique"]="CNU-9"
    dict_CNU_autres["Anthropologie"]="CNU-4"
    dict_CNU_autres["archéologie"] = "CNU-4"
    dict_CNU_autres["Archéologie"] = "CNU-4"
    dict_CNU_autres["Microbiologie"] = "CNU-10"
    dict_CNU_autres["Aménagement"] = "CNU-4"
    dict_CNU_autres["24"] = "CNU-4"
    dict_CNU_autres["groupe 10 du CNU"]="CNU-10"
    dict_CNU_autres["Santé publique, Epidémiologie"] = "Santé"
    dict_CNU_autres["Hydrologie"]="CNU-8"
    dict_CNU_autres["Sociologie des Sciences et des techniques"]="CNU-12"
    return dict_CNU_autres


def get_dict_CNU_inv():
    dict_CNU_inv = {}
    dict_CNU_inv["CNU-1"]="Droit privé, droit public, sciences politiques, histoire du droit (sections 1 à 4)"
    dict_CNU_inv["CNU-2"]="Sciences économiques, sciences de gestion et du management (sections 5 et 6)"
    dict_CNU_inv["CNU-3"]="Littérature, langues, sciences du langage, linguistique (sections 7 à 15)"
    dict_CNU_inv["CNU-4"]="Psychologie, sociologie, histoire, géographie, urbanisme, ethnologie, philosophie, architecture (sections 16 à 24)"
    dict_CNU_inv["CNU-5"]="Mathématiques, informatique (sections 25 à 27)"
    dict_CNU_inv["CNU-6"]="Physique hors astrophysique/astronomie (section 28 à 30)"
    dict_CNU_inv["CNU-7"]="Chimie (sections 31 à 33)"
    dict_CNU_inv["CNU-8"]="Astronomie/astrophysique, géologie, météorologie (sections 34 à 37)"
    dict_CNU_inv["CNU-9"]="Mécanique, génie informatique, automatique, traitement du signal, génie électrique, électronique, photoniques (sections 60 à 63)"
    dict_CNU_inv["CNU-10"]="Biologie, biochimie, physiologie, neurosciences, écologie (sections 64 à 69)"
    dict_CNU_inv["CNU-12"]="Pluridisciplinaire (sections 70 à 74)"
    return dict_CNU_inv